using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using WordAgent.Core;

namespace WordAgent.WordInterop;


public sealed record TocOpenXmlParagraphAuthority(
    string Role,
    string StyleName,
    string StableTitle,
    string ReadingOrder,
    string Alignment,
    bool TextRunsRtl,
    bool ReadingOrderMatch,
    bool AlignmentMatch,
    bool TextRunsRtlMatch);

public sealed record TocOpenXmlAuthoritySnapshot(
    int StandardTocControls,
    int Paragraphs,
    int MappedParagraphs,
    int ReadingOrderMatches,
    int AlignmentMatches,
    int TextRunsRtlMatches,
    IReadOnlyList<TocOpenXmlParagraphAuthority> Items,
    IReadOnlyList<string> Issues)
{
    // "Repairable" is a structural safety statement, not a claim that the
    // currently broken visual formatting is already compliant. A standard TOC
    // whose entries can all be mapped deterministically is safe to normalize.
    public bool RepairableVisualLayout
        => StandardTocControls > 0
           && Paragraphs > 0
           && Paragraphs == MappedParagraphs
           && Issues.Count == 0;

    public bool FullyCompliant
        => RepairableVisualLayout
           && Paragraphs == ReadingOrderMatches
           && Paragraphs == AlignmentMatches
           && Paragraphs == TextRunsRtlMatches;
}

public static class TocOpenXmlComplianceAuthorityService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static TocOpenXmlAuthoritySnapshot Inspect(string inputPath, AgentProfile profile)
    {
        var path = PathPolicy.RequireSupportedWordFile(inputPath);
        var aliasToRole = BuildTocAliasMap(profile);
        var issues = new List<string>();
        var items = new List<TocOpenXmlParagraphAuthority>();

        using var zip = ZipFile.OpenRead(path);
        var documentEntry = zip.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("word/document.xml not found.");
        var stylesEntry = zip.GetEntry("word/styles.xml")
            ?? throw new InvalidDataException("word/styles.xml not found.");

        XDocument document;
        XDocument styles;
        using (var stream = documentEntry.Open())
            document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        using (var stream = stylesEntry.Open())
            styles = XDocument.Load(stream, LoadOptions.PreserveWhitespace);

        var styleCatalog = TocStyleCatalog.Create(styles);

        var tocControls = document
            .Descendants(W + "sdt")
            .Where(IsStandardTableOfContentsSdt)
            .ToList();

        List<XElement> paragraphs;
        var authorityScopes = tocControls.Count;

        if (tocControls.Count > 0)
        {
            paragraphs = tocControls
                .SelectMany(x => x.Element(W + "sdtContent")?.Descendants(W + "p") ?? Enumerable.Empty<XElement>())
                .Where(x => Normalize(ParagraphText(x)).Length > 0)
                .ToList();
        }
        else
        {
            paragraphs = FindRecoveredDocumentLevelTocParagraphs(document)
                .Where(x => Normalize(ParagraphText(x)).Length > 0)
                .ToList();
            authorityScopes = paragraphs.Count > 0 ? 1 : 0;
        }

        foreach (var paragraph in paragraphs)
        {
            var styleId = (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
            var styleName = styleCatalog.GetStyleName(styleId) ?? styleId ?? "(unknown)";
            var stableTitle = StableTitle(paragraph);

            var mappedByStyle = aliasToRole.TryGetValue(styleName, out var role)
                && (role.Equals("Toc1", StringComparison.OrdinalIgnoreCase)
                    || role.Equals("Toc2", StringComparison.OrdinalIgnoreCase));

            if (!mappedByStyle)
            {
                role = InferTocRoleFromStableTitle(stableTitle, profile);
                if (role is null)
                {
                    issues.Add(
                        $"Automatic TOC paragraph role cannot be inferred safely: style='{styleName}'; title='{stableTitle}'.");
                    continue;
                }
            }

            if (role is null || !profile.RoleDefaults.TryGetValue(role, out var expected))
            {
                issues.Add($"Automatic TOC role has no profile defaults: {role}.");
                continue;
            }

            var resolvedStyle = styleCatalog.Resolve(styleId);
            var pPr = paragraph.Element(W + "pPr");

            var directBidi = ParseOnOff(pPr?.Element(W + "bidi"));
            var reading = directBidi is not null
                ? (directBidi.Value ? "RTL" : "LTR")
                : resolvedStyle.ReadingOrder;

            var directAlignment = NormalizeAlignment((string?)pPr?.Element(W + "jc")?.Attribute(W + "val"));
            // QH01-E235 / A19R1:
            // Automatic TOC OpenXML authority uses the literal Word paragraph alignment token.
            // w:jc="right" is RIGHT; w:bidi is validated independently as reading-order authority.
            // Do not mirror LEFT/RIGHT a second time in the compliance authority.
            var alignment = !string.IsNullOrWhiteSpace(directAlignment)
                ? directAlignment!
                : resolvedStyle.Alignment;

            var textRuns = paragraph
                .Descendants(W + "r")
                .Where(r => r.Descendants(W + "t").Any(t => !string.IsNullOrWhiteSpace((string?)t)))
                .ToList();

            var allTextRunsRtl = textRuns.Count > 0 && textRuns.All(run =>
            {
                var direct = ParseOnOff(run.Element(W + "rPr")?.Element(W + "rtl"));
                if (direct is not null) return direct.Value;
                return resolvedStyle.RunRtl is true;
            });

            var expectedReadingRtl = string.Equals(expected.ReadingOrder?.Trim(), "RTL", StringComparison.OrdinalIgnoreCase);
            var readingMatch = MatchToken(expected.ReadingOrder, reading);
            var alignmentMatch = MatchToken(expected.Alignment, alignment);
            var runRtlMatch = !expectedReadingRtl || allTextRunsRtl;

            items.Add(new TocOpenXmlParagraphAuthority(
                role,
                styleName,
                stableTitle,
                reading,
                alignment,
                allTextRunsRtl,
                readingMatch,
                alignmentMatch,
                runRtlMatch));
        }

        return new TocOpenXmlAuthoritySnapshot(
            authorityScopes,
            paragraphs.Count,
            items.Count,
            items.Count(x => x.ReadingOrderMatch),
            items.Count(x => x.AlignmentMatch),
            items.Count(x => x.TextRunsRtlMatch),
            items,
            issues);
    }

    internal static string? InferTocRoleFromStableTitle(
        string stableTitle,
        AgentProfile profile)
    {
        if (string.IsNullOrWhiteSpace(stableTitle))
            return null;

        // Only infer roles that the selected target profile actually supports.
        var hasToc1 = profile.RoleDefaults.ContainsKey("Toc1");
        var hasToc2 = profile.RoleDefaults.ContainsKey("Toc2");
        if (!hasToc1 && !hasToc2)
            return null;

        // Deterministic hierarchy cues used by the supported Persian document families:
        //   1- Heading          => level 1
        //   2-1- Subheading    => level 2
        //   A- Domain heading  => level 2
        // Persian/Arabic/ASCII digits are accepted. This is applied only inside a
        // standard Word "Table of Contents" content control.
        var numeric = System.Text.RegularExpressions.Regex.Match(
            stableTitle,
            @"^\s*[0-9۰-۹٠-٩]+(?<tail>(?:\s*[-–—.]\s*[0-9۰-۹٠-٩]+)*)\s*[-–—.)]?\s+");

        if (numeric.Success)
        {
            var segmentCount = System.Text.RegularExpressions.Regex.Matches(
                numeric.Value,
                @"[0-9۰-۹٠-٩]+").Count;

            if (segmentCount <= 1 && hasToc1)
                return "Toc1";
            if (segmentCount >= 2 && hasToc2)
                return "Toc2";
        }

        var alpha = System.Text.RegularExpressions.Regex.Match(
            stableTitle,
            @"^\s*[A-Za-z]\s*[-–—.)]\s+");

        if (alpha.Success && hasToc2)
            return "Toc2";

        return null;
    }

    private static Dictionary<string, string> BuildTocAliasMap(AgentProfile profile)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in profile.RoleAliases)
        {
            var canonicalRole = CanonicalizeTocRole(pair.Key);
            if (canonicalRole is null)
                continue;

            foreach (var alias in pair.Value)
            {
                if (!string.IsNullOrWhiteSpace(alias))
                    map[alias.Trim()] = canonicalRole;
            }
        }
        return map;
    }

    private static string? CanonicalizeTocRole(string role)
    {
        if (role.Equals("Toc1", StringComparison.OrdinalIgnoreCase)
            || role.Equals("TocLegacy1", StringComparison.OrdinalIgnoreCase))
            return "Toc1";

        if (role.Equals("Toc2", StringComparison.OrdinalIgnoreCase)
            || role.Equals("TocLegacy2", StringComparison.OrdinalIgnoreCase))
            return "Toc2";

        return null;
    }

    private static bool IsStandardTableOfContentsSdt(XElement sdt)
    {
        var gallery = (string?)sdt
            .Element(W + "sdtPr")?
            .Element(W + "docPartObj")?
            .Element(W + "docPartGallery")?
            .Attribute(W + "val");

        if (string.Equals(
                gallery?.Trim(),
                "Table of Contents",
                StringComparison.OrdinalIgnoreCase))
            return true;

        var joinedInstruction = string.Join(
            " ",
            sdt.Descendants(W + "instrText")
                .Select(x => ((string?)x ?? string.Empty).Trim()));

        if (IsAutomaticTocMainInstruction(joinedInstruction))
            return true;

        return sdt
            .Descendants(W + "p")
            .Any(HasAutomaticTocEntrySignature);
    }

    private static bool IsAutomaticTocMainInstruction(string? value)
    {
        var normalized = Normalize(value);
        return normalized.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)
               && (normalized.Length == 3 || char.IsWhiteSpace(normalized[3]));
    }

    private static bool HasAutomaticTocEntrySignature(XElement paragraph)
    {
        var instruction = string.Join(
            " ",
            paragraph.Descendants(W + "instrText")
                .Select(x => ((string?)x ?? string.Empty).Trim()));

        if (instruction.Contains("PAGEREF", StringComparison.OrdinalIgnoreCase)
            && instruction.Contains("_Toc", StringComparison.OrdinalIgnoreCase))
            return true;

        return paragraph
            .Descendants(W + "hyperlink")
            .Select(link => (string?)link.Attribute(W + "anchor"))
            .Any(anchor =>
                !string.IsNullOrWhiteSpace(anchor)
                && anchor.StartsWith("_Toc", StringComparison.OrdinalIgnoreCase));
    }

    private static IReadOnlyList<XElement> FindRecoveredDocumentLevelTocParagraphs(XDocument document)
    {
        var paragraphs = document.Descendants(W + "p").ToArray();
        var entryIndexes = paragraphs
            .Select((paragraph, index) => new { paragraph, index })
            .Where(x => HasAutomaticTocEntrySignature(x.paragraph))
            .Select(x => x.index)
            .ToArray();

        if (entryIndexes.Length < 2)
            return Array.Empty<XElement>();

        var start = entryIndexes.Min();
        var end = entryIndexes.Max();

        var mainIndexes = paragraphs
            .Select((paragraph, index) => new
            {
                index,
                instruction = string.Join(
                    " ",
                    paragraph.Descendants(W + "instrText")
                        .Select(x => ((string?)x ?? string.Empty).Trim()))
            })
            .Where(x => IsAutomaticTocMainInstruction(x.instruction))
            .Select(x => x.index)
            .Where(index => index <= start)
            .ToArray();

        if (mainIndexes.Length > 0)
        {
            var nearestMain = mainIndexes.Max();
            if (start - nearestMain <= 3)
                start = nearestMain;
        }

        if (end - start > 250)
            return Array.Empty<XElement>();

        return paragraphs
            .Skip(start)
            .Take(end - start + 1)
            .ToArray();
    }

    private static string StableTitle(XElement paragraph)
    {
        var pieces = new List<string>();
        foreach (var run in paragraph.Descendants(W + "r"))
        {
            foreach (var child in run.Elements())
            {
                if (child.Name == W + "tab")
                    return Normalize(string.Concat(pieces));

                if (child.Name == W + "t")
                {
                    pieces.Add((string?)child ?? string.Empty);
                    continue;
                }

                foreach (var text in child.Descendants(W + "t"))
                    pieces.Add((string?)text ?? string.Empty);
            }
        }

        return Normalize(string.Concat(pieces));
    }

    private static string ParagraphText(XElement paragraph)
        => string.Concat(paragraph.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));

    private static string Normalize(string? value)
        => string.Join(
            ' ',
            (value ?? string.Empty)
                .Replace('\u200c', ' ')
                .Replace('\u200f', ' ')
                .Replace('\u200e', ' ')
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static bool MatchToken(string? expected, string? actual)
        => string.IsNullOrWhiteSpace(expected)
           || string.Equals(expected.Trim(), actual?.Trim(), StringComparison.OrdinalIgnoreCase);

    private static string? NormalizeAlignment(string? value)
    {
        var token = value?.Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(token)) return null;
        return token switch
        {
            "right" or "end" => "RIGHT",
            "left" or "start" => "LEFT",
            "center" => "CENTER",
            "both" => "JUSTIFY",
            "distribute" => "DISTRIBUTE",
            _ => token.ToUpperInvariant()
        };
    }

    private static bool? ParseOnOff(XElement? element)
    {
        if (element is null) return null;
        var value = ((string?)element.Attribute(W + "val"))?.Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(value)) return true;
        return value switch
        {
            "0" or "false" or "off" or "no" => false,
            _ => true
        };
    }

    private sealed class TocStyleCatalog
    {
        private readonly Dictionary<string, TocStyleNode> _styles;
        private readonly TocResolvedStyle _defaults;
        private readonly Dictionary<string, TocResolvedStyle> _cache = new(StringComparer.OrdinalIgnoreCase);

        private TocStyleCatalog(
            Dictionary<string, TocStyleNode> styles,
            TocResolvedStyle defaults)
        {
            _styles = styles;
            _defaults = defaults;
        }

        public static TocStyleCatalog Create(XDocument styles)
        {
            var root = styles.Root ?? throw new InvalidDataException("styles.xml root not found.");
            var defaults = ReadDefaults(root);
            var map = new Dictionary<string, TocStyleNode>(StringComparer.OrdinalIgnoreCase);

            foreach (var style in root.Elements(W + "style"))
            {
                var id = (string?)style.Attribute(W + "styleId");
                if (string.IsNullOrWhiteSpace(id)) continue;
                var name = (string?)style.Element(W + "name")?.Attribute(W + "val") ?? id;
                var basedOn = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
                map[id] = new TocStyleNode(id, name, basedOn, style);
            }

            return new TocStyleCatalog(map, defaults);
        }

        public string? GetStyleName(string? styleId)
            => !string.IsNullOrWhiteSpace(styleId) && _styles.TryGetValue(styleId, out var node)
                ? node.Name
                : null;

        public TocResolvedStyle Resolve(string? styleId)
        {
            if (string.IsNullOrWhiteSpace(styleId)) return _defaults;
            if (_cache.TryGetValue(styleId, out var cached)) return cached;
            var resolved = ResolveCore(styleId, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            _cache[styleId] = resolved;
            return resolved;
        }

        private TocResolvedStyle ResolveCore(string styleId, HashSet<string> visiting)
        {
            if (_cache.TryGetValue(styleId, out var cached)) return cached;
            if (!_styles.TryGetValue(styleId, out var node)) return _defaults;
            if (!visiting.Add(styleId)) return _defaults;

            var inherited = !string.IsNullOrWhiteSpace(node.BasedOn)
                ? ResolveCore(node.BasedOn!, visiting)
                : _defaults;

            var resolved = Apply(inherited, node.Element);
            visiting.Remove(styleId);
            _cache[styleId] = resolved;
            return resolved;
        }

        private static TocResolvedStyle ReadDefaults(XElement root)
        {
            var result = new TocResolvedStyle("UNKNOWN", "UNKNOWN", null);
            var docDefaults = root.Element(W + "docDefaults");
            if (docDefaults is null) return result;

            var pPr = docDefaults.Element(W + "pPrDefault")?.Element(W + "pPr");
            var rPr = docDefaults.Element(W + "rPrDefault")?.Element(W + "rPr");
            return ApplyProperties(result, pPr, rPr);
        }

        private static TocResolvedStyle Apply(TocResolvedStyle inherited, XElement style)
            => ApplyProperties(inherited, style.Element(W + "pPr"), style.Element(W + "rPr"));

        private static TocResolvedStyle ApplyProperties(
            TocResolvedStyle inherited,
            XElement? pPr,
            XElement? rPr)
        {
            var reading = inherited.ReadingOrder;
            var alignment = inherited.Alignment;
            var runRtl = inherited.RunRtl;

            var bidi = ParseOnOff(pPr?.Element(W + "bidi"));
            if (bidi is not null) reading = bidi.Value ? "RTL" : "LTR";

            var jc = NormalizeAlignment((string?)pPr?.Element(W + "jc")?.Attribute(W + "val"));
            if (!string.IsNullOrWhiteSpace(jc)) alignment = jc!;

            var rtl = ParseOnOff(rPr?.Element(W + "rtl"));
            if (rtl is not null) runRtl = rtl;

            return new TocResolvedStyle(reading, alignment, runRtl);
        }

        private sealed record TocStyleNode(
            string Id,
            string Name,
            string? BasedOn,
            XElement Element);

        public sealed record TocResolvedStyle(
            string ReadingOrder,
            string Alignment,
            bool? RunRtl);
    }
}

public sealed class ProfileComplianceService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const int MaxUnclassifiedSamples = 12;

    public ProfileComplianceSnapshot Inspect(string inputPath, AgentProfile profile, bool visible = false)
    {
        var path = PathPolicy.RequireSupportedWordFile(inputPath);
        var warnings = new List<string>();
        var aliasToRole = BuildAliasToRole(profile);
        var tocOpenXmlAuthority = TocOpenXmlComplianceAuthorityService.Inspect(path, profile);
        var styleCompliance = InspectStylesFromOpenXml(path, profile, aliasToRole, warnings);
        var styleStateByName = styleCompliance
            .Where(x => x.Found)
            .GroupBy(x => x.StyleName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        // R10R41: the three remaining S3 audit/plan timeout families use a persisted-OOXML
        // compliance read model. This preserves deterministic paragraph/action indexing while
        // removing unbounded full-document cross-process Word COM enumeration. Native Word
        // compatibility remains independently verified by the isolated Word-native smoke gate.
        if (UsePersistedOpenXmlFastPath(profile))
        {
            warnings.Add(
                "PROFILE_COMPLIANCE_AUTHORITY=PERSISTED_OOXML_FAST_PATH; WORD_NATIVE_COMPATIBILITY_IS_VERIFIED_SEPARATELY.");
            return InspectPersistedOpenXml(
                path,
                profile,
                aliasToRole,
                styleCompliance,
                styleStateByName,
                tocOpenXmlAuthority,
                warnings);
        }

        Word.Application? app = null;
        Word.Documents? documents = null;
        Word.Document? doc = null;

        try
        {
            app = new Word.Application();
            app.Visible = visible;
            app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            app.ScreenUpdating = visible;

            documents = app.Documents;
            doc = documents.Open(
                FileName: path,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: visible,
                OpenAndRepair: false,
                NoEncodingDialog: true);

            return InspectOpenDocument(
                doc,
                profile,
                aliasToRole,
                styleCompliance,
                styleStateByName,
                tocOpenXmlAuthority,
                warnings);
        }
        finally
        {
            if (doc is not null)
            {
                try { doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                ComUtil.Release(doc);
            }

            ComUtil.Release(documents);

            if (app is not null)
            {
                try { app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                ComUtil.Release(app);
            }
        }
    }


    private static bool UsePersistedOpenXmlFastPath(AgentProfile profile)
        => profile.Id.Equals("proposal.academic.fa", StringComparison.OrdinalIgnoreCase)
           || profile.Id.Equals("thesis-dissertation-academic-fa", StringComparison.OrdinalIgnoreCase)
           || profile.Id.Equals("thesis.dissertation.phd.fa", StringComparison.OrdinalIgnoreCase);

    private static ProfileComplianceSnapshot InspectPersistedOpenXml(
        string inputPath,
        AgentProfile profile,
        IReadOnlyDictionary<string, string> aliasToRole,
        IReadOnlyList<StyleComplianceItem> styleCompliance,
        IReadOnlyDictionary<string, StyleComplianceItem> styleStateByName,
        TocOpenXmlAuthoritySnapshot tocOpenXmlAuthority,
        List<string> warnings)
    {
        using var zip = ZipFile.OpenRead(inputPath);
        var documentEntry = zip.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("word/document.xml not found.");
        var stylesEntry = zip.GetEntry("word/styles.xml")
            ?? throw new InvalidDataException("word/styles.xml not found.");

        XDocument document;
        XDocument styles;
        using (var stream = documentEntry.Open())
            document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        using (var stream = stylesEntry.Open())
            styles = XDocument.Load(stream, LoadOptions.PreserveWhitespace);

        var catalog = OpenXmlStyleCatalog.Load(inputPath);
        var styleNodes = BuildPersistedStyleNodes(styles);
        var styleNames = styleNodes.Values
            .Where(x => !string.IsNullOrWhiteSpace(x.Name))
            .ToDictionary(x => x.StyleId, x => x.Name!, StringComparer.OrdinalIgnoreCase);
        var outlineCache = new Dictionary<string, int?>(StringComparer.OrdinalIgnoreCase);
        var persistedParagraphs = BuildPersistedWordParagraphSequence(document);

        var roleBuilders = new Dictionary<string, RoleBuilder>(StringComparer.OrdinalIgnoreCase);
        var unclassifiedSamples = new List<UnclassifiedParagraphSample>();
        int nonEmpty = 0, classified = 0, unclassified = 0, inTables = 0, overrideSuspected = 0;
        int recoveredOutlineChapter = 0, recoveredOutlineHeading1 = 0;
        int frontMatterOrdinal = 0;
        bool frontMatterClosed = false;
        bool lastNonEmptyParagraphWasInTable = false;

        foreach (var entry in persistedParagraphs)
        {
            var paragraph = entry.Paragraph;
            var rawText = PersistedParagraphRawText(paragraph);
            var cleanText = WordAuditService.CleanParagraphText(rawText);
            if (string.IsNullOrWhiteSpace(cleanText))
                continue;

            nonEmpty++;
            var inTable = entry.InTable;
            var followsTableBoundary = !inTable && lastNonEmptyParagraphWasInTable;
            lastNonEmptyParagraphWasInTable = inTable;
            if (inTable)
            {
                inTables++;
                frontMatterClosed = true;
            }

            var pPr = paragraph.Element(W + "pPr");
            var styleId = (string?)pPr?.Element(W + "pStyle")?.Attribute(W + "val");
            var styleName = !string.IsNullOrWhiteSpace(styleId) && styleNames.TryGetValue(styleId!, out var mappedName)
                ? mappedName
                : styleId ?? "(unknown)";

            var outlineLevel = ResolvePersistedOutlineLevel(
                pPr,
                styleId,
                styleNodes,
                outlineCache,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            var outlineHeadingRole = ResolvePersistedSourceHeadingRoleFromOutlineLevel(
                outlineLevel,
                profile,
                rawText,
                cleanText);

            var exactTocTitle = IsExactPersianTocTitle(cleanText);
            if (exactTocTitle)
                frontMatterClosed = true;

            var invocationMatch =
                !string.IsNullOrWhiteSpace(profile.SemanticPolicy.Invocation.Role)
                && SemanticPolicyMatcher.MatchesInvocation(
                    cleanText,
                    profile.SemanticPolicy.Invocation);

            string? frontMatterRole = null;
            if (!frontMatterClosed && !inTable && !invocationMatch)
            {
                frontMatterOrdinal++;
                frontMatterRole = ResolveFrontMatterRole(profile, frontMatterOrdinal);
            }

            var tocAuthority = ResolveAutomaticTocAuthorityByWordResult(
                tocOpenXmlAuthority,
                styleName,
                rawText);

            var structuralHeadingRole = tocAuthority is null
                ? ResolveSourceHeadingRoleFromToc(
                    tocOpenXmlAuthority,
                    rawText,
                    cleanText,
                    profile)
                : null;

            var recoveredHeadingRole = tocAuthority is null && structuralHeadingRole is null
                ? outlineHeadingRole
                : null;

            if (recoveredHeadingRole is not null)
            {
                if (recoveredHeadingRole.Equals("Chapter", StringComparison.OrdinalIgnoreCase))
                    recoveredOutlineChapter++;
                else if (recoveredHeadingRole.Equals("Heading1", StringComparison.OrdinalIgnoreCase))
                    recoveredOutlineHeading1++;
            }

            var mainTocFieldRole =
                IsPersistedMainTocFieldParagraph(paragraph)
                && profile.RoleDefaults.ContainsKey("TocField")
                    ? "TocField"
                    : null;

            var role = tocAuthority?.Role
                ?? structuralHeadingRole
                ?? recoveredHeadingRole
                ?? frontMatterRole
                ?? mainTocFieldRole
                ?? ResolveRole(
                    styleName,
                    inTable,
                    followsTableBoundary,
                    entry.TableRowIndex,
                    entry.ParagraphIndex,
                    cleanText,
                    profile,
                    aliasToRole);

            if (role is null || !profile.RoleDefaults.TryGetValue(role, out var expected))
            {
                unclassified++;
                if (unclassifiedSamples.Count < MaxUnclassifiedSamples)
                    unclassifiedSamples.Add(
                        new UnclassifiedParagraphSample(
                            entry.ParagraphIndex,
                            styleName,
                            Preview(cleanText)));
                continue;
            }

            classified++;
            if (!roleBuilders.TryGetValue(role, out var builder))
            {
                builder = new RoleBuilder(role, expected);
                roleBuilders[role] = builder;
            }

            var actual = ResolvePersistedParagraphFormatting(
                paragraph,
                styleId,
                catalog);

            var actualFont = actual.FontBi;
            var actualSize = actual.SizeBi;
            var actualReading = actual.ReadingOrder;
            var actualAlignment = actual.Alignment;

            styleStateByName.TryGetValue(styleName, out var styleState);
            var sameRoleStyle = styleState is not null
                && string.Equals(styleState.Role, role, StringComparison.OrdinalIgnoreCase);

            tocAuthority ??= ResolveAutomaticTocAuthority(
                tocOpenXmlAuthority,
                role,
                styleName,
                rawText);

            if (tocAuthority is not null)
            {
                actualReading = tocAuthority.ReadingOrder;
                actualAlignment = tocAuthority.Alignment;
            }

            var fontMatch = string.IsNullOrWhiteSpace(expected.FontBi)
                || (actualFont is not null
                    ? MatchesFont(expected.FontBi, actualFont)
                    : sameRoleStyle && styleState!.FontBiMatch);

            var sizeMatch = expected.SizeBi is null
                || (actualSize is not null
                    ? MatchesSize(expected.SizeBi, actualSize)
                    : sameRoleStyle && styleState!.SizeBiMatch);

            var readingMatch = string.IsNullOrWhiteSpace(expected.ReadingOrder)
                || (tocAuthority is not null
                    ? tocAuthority.ReadingOrderMatch && tocAuthority.TextRunsRtlMatch
                    : !IsUnknownToken(actualReading)
                        ? MatchesToken(expected.ReadingOrder, actualReading)
                        : sameRoleStyle && styleState!.ReadingOrderMatch);

            var alignmentMatch = string.IsNullOrWhiteSpace(expected.Alignment)
                || (tocAuthority is not null
                    ? true
                    : !IsUnknownToken(actualAlignment)
                        ? MatchesToken(expected.Alignment, actualAlignment)
                        : sameRoleStyle && styleState!.AlignmentMatch);

            var semanticTableNoteStyleRequired =
                role.Equals("TableNote", StringComparison.OrdinalIgnoreCase)
                && (inTable || followsTableBoundary)
                && IsSemanticTableNote(cleanText);

            var semanticHeadingStyleRequired =
                role.Equals("Chapter", StringComparison.OrdinalIgnoreCase)
                || role.Equals("Heading1", StringComparison.OrdinalIgnoreCase);

            var semanticStyleIdentityRequired =
                semanticTableNoteStyleRequired || semanticHeadingStyleRequired;

            var semanticStyleIdentityMatch =
                !semanticStyleIdentityRequired
                || (aliasToRole.TryGetValue(styleName, out var semanticStyleRole)
                    && semanticStyleRole.Equals(role, StringComparison.OrdinalIgnoreCase));

            var fullMatch =
                fontMatch && sizeMatch && readingMatch && alignmentMatch
                && semanticStyleIdentityMatch;

            var directOverride = false;
            if (!inTable && sameRoleStyle && tocAuthority is null)
            {
                directOverride = KnownFormatDifference(
                    styleState!.ActualFontBi,
                    styleState.ActualSizeBi,
                    styleState.ActualReadingOrder,
                    styleState.ActualAlignment,
                    actualFont,
                    actualSize,
                    actualReading,
                    actualAlignment);
                if (directOverride) overrideSuspected++;
            }

            if (semanticTableNoteStyleRequired
                && !semanticStyleIdentityMatch
                && !directOverride)
            {
                directOverride = true;
                overrideSuspected++;
            }

            builder.Add(
                fontMatch,
                sizeMatch,
                readingMatch,
                alignmentMatch,
                fullMatch,
                directOverride);

            if (!fullMatch || directOverride)
            {
                builder.Samples.Add(new ComplianceSample(
                    entry.ParagraphIndex,
                    styleName,
                    Preview(cleanText),
                    actualFont,
                    actualSize,
                    actualReading,
                    actualAlignment,
                    directOverride,
                    ExpectedSummary(expected)));
            }
        }

        if (warnings.Count < 80)
        {
            warnings.Add(
                $"PERSISTED_OOXML_COMPLIANCE_SUMMARY:WORD_INDEXED_PARAGRAPHS={persistedParagraphs.Count}:"
                + $"NONEMPTY={nonEmpty}:OUTLINE_CHAPTER={recoveredOutlineChapter}:"
                + $"OUTLINE_HEADING1={recoveredOutlineHeading1}");
        }

        var roleSummaries = roleBuilders.Values
            .OrderBy(x => x.Role, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.ToSummary())
            .ToArray();

        return new ProfileComplianceSnapshot(
            profile.Id,
            profile.DisplayName,
            nonEmpty,
            classified,
            unclassified,
            inTables,
            overrideSuspected,
            roleSummaries,
            styleCompliance,
            unclassifiedSamples,
            warnings);
    }

    private static IReadOnlyList<PersistedParagraphEntry> BuildPersistedWordParagraphSequence(
        XDocument document)
    {
        var rowIndexByElement = new Dictionary<XElement, int>();
        foreach (var table in document.Descendants(W + "tbl"))
        {
            var row = 0;
            foreach (var tr in table.Elements(W + "tr"))
            {
                row++;
                rowIndexByElement[tr] = row;
            }
        }

        var result = new List<PersistedParagraphEntry>();
        var wordIndex = 0;

        void Visit(XElement element)
        {
            if (element.Name == W + "p")
            {
                wordIndex++;
                var row = element.Ancestors(W + "tr").FirstOrDefault();
                var rowIndex = row is not null && rowIndexByElement.TryGetValue(row, out var value)
                    ? value
                    : (int?)null;

                result.Add(new PersistedParagraphEntry(
                    wordIndex,
                    element,
                    element.Ancestors(W + "tc").Any(),
                    rowIndex));
                return;
            }

            foreach (var child in element.Elements())
                Visit(child);

            // Match the structural Word paragraph-index authority already used by
            // ControlledCopyFixService: each table row contributes one Word paragraph marker
            // that has no standalone w:p node.
            if (element.Name == W + "tr")
                wordIndex++;
        }

        if (document.Root is null)
            throw new InvalidDataException("word/document.xml has no root element.");

        Visit(document.Root);
        return result;
    }

    private static Dictionary<string, PersistedStyleNode> BuildPersistedStyleNodes(XDocument styles)
    {
        var result = new Dictionary<string, PersistedStyleNode>(StringComparer.OrdinalIgnoreCase);
        var root = styles.Root;
        if (root is null) return result;

        foreach (var style in root.Elements(W + "style"))
        {
            var id = (string?)style.Attribute(W + "styleId");
            if (string.IsNullOrWhiteSpace(id)) continue;
            result[id] = new PersistedStyleNode(
                id,
                (string?)style.Element(W + "name")?.Attribute(W + "val"),
                (string?)style.Element(W + "basedOn")?.Attribute(W + "val"),
                style);
        }
        return result;
    }

    private static int? ResolvePersistedOutlineLevel(
        XElement? paragraphProperties,
        string? styleId,
        IReadOnlyDictionary<string, PersistedStyleNode> styles,
        IDictionary<string, int?> cache,
        HashSet<string> visiting)
    {
        var direct = ParseOutlineLevel(paragraphProperties?.Element(W + "outlineLvl"));
        if (direct is not null) return direct;

        if (string.IsNullOrWhiteSpace(styleId))
            return null;

        return ResolvePersistedStyleOutlineLevel(styleId!, styles, cache, visiting);
    }

    private static int? ResolvePersistedStyleOutlineLevel(
        string styleId,
        IReadOnlyDictionary<string, PersistedStyleNode> styles,
        IDictionary<string, int?> cache,
        HashSet<string> visiting)
    {
        if (cache.TryGetValue(styleId, out var cached))
            return cached;
        if (!styles.TryGetValue(styleId, out var node))
            return null;
        if (!visiting.Add(styleId))
            return null;

        int? value = null;
        if (!string.IsNullOrWhiteSpace(node.BasedOn))
            value = ResolvePersistedStyleOutlineLevel(node.BasedOn!, styles, cache, visiting);

        var direct = ParseOutlineLevel(node.Element.Element(W + "pPr")?.Element(W + "outlineLvl"));
        if (direct is not null)
            value = direct;

        visiting.Remove(styleId);
        cache[styleId] = value;
        return value;
    }

    private static int? ParseOutlineLevel(XElement? element)
    {
        var raw = ((string?)element?.Attribute(W + "val"))?.Trim();
        if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
            return null;
        return value is >= 0 and <= 8 ? value : null;
    }

    private static string? ResolvePersistedSourceHeadingRoleFromOutlineLevel(
        int? zeroBasedOutlineLevel,
        AgentProfile profile,
        string rawText,
        string cleanText)
    {
        if (zeroBasedOutlineLevel is null
            || string.IsNullOrWhiteSpace(cleanText)
            || (!string.IsNullOrEmpty(rawText) && rawText.Contains('\t'))
            || IsExactPersianTocTitle(cleanText))
            return null;

        if (zeroBasedOutlineLevel == 0 && profile.RoleDefaults.ContainsKey("Chapter"))
            return "Chapter";

        if (zeroBasedOutlineLevel == 1 && profile.RoleDefaults.ContainsKey("Heading1"))
            return "Heading1";

        return null;
    }

    private static PersistedParagraphFormatting ResolvePersistedParagraphFormatting(
        XElement paragraph,
        string? paragraphStyleId,
        OpenXmlStyleCatalog catalog)
    {
        var style = !string.IsNullOrWhiteSpace(paragraphStyleId)
            ? catalog.Resolve(paragraphStyleId!)
            : new OpenXmlStyleCatalog.ResolvedStyle(null, null, "UNKNOWN", "UNKNOWN");

        var pPr = paragraph.Element(W + "pPr");
        var reading = style.ReadingOrder;
        var alignment = style.Alignment;

        var directBidi = ParsePersistedOnOff(pPr?.Element(W + "bidi"));
        if (directBidi is not null)
            reading = directBidi.Value ? "RTL" : "LTR";

        var directAlignment = NormalizePersistedAlignment(
            (string?)pPr?.Element(W + "jc")?.Attribute(W + "val"));
        if (!string.IsNullOrWhiteSpace(directAlignment))
            alignment = directAlignment!;

        var runFonts = new List<string?>();
        var runSizes = new List<float?>();

        foreach (var run in paragraph.Descendants(W + "r")
                     .Where(r => r.Descendants(W + "t")
                         .Any(t => !string.IsNullOrWhiteSpace((string?)t))))
        {
            var font = style.FontBi;
            var size = style.SizeBi;
            var rPr = run.Element(W + "rPr");

            var rStyleId = (string?)rPr?.Element(W + "rStyle")?.Attribute(W + "val");
            if (!string.IsNullOrWhiteSpace(rStyleId))
            {
                var charStyle = catalog.Resolve(rStyleId!);
                if (!string.IsNullOrWhiteSpace(charStyle.FontBi))
                    font = charStyle.FontBi;
                if (charStyle.SizeBi is not null)
                    size = charStyle.SizeBi;
            }

            var cs = (string?)rPr?.Element(W + "rFonts")?.Attribute(W + "cs");
            if (!string.IsNullOrWhiteSpace(cs))
                font = cs.Trim();

            var szCs = (string?)rPr?.Element(W + "szCs")?.Attribute(W + "val");
            if (int.TryParse(szCs, NumberStyles.Integer, CultureInfo.InvariantCulture, out var halfPoints)
                && halfPoints > 0)
                size = halfPoints / 2f;

            runFonts.Add(font);
            runSizes.Add(size);
        }

        var actualFont = CollapseFont(runFonts, style.FontBi);
        var actualSize = CollapseSize(runSizes, style.SizeBi);

        return new PersistedParagraphFormatting(
            actualFont,
            actualSize,
            reading,
            alignment);
    }

    private static string? CollapseFont(IReadOnlyList<string?> values, string? fallback)
    {
        if (values.Count == 0) return fallback;
        var distinct = values
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x!.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (distinct.Length == 1) return distinct[0];
        if (distinct.Length == 0 && values.All(x => string.IsNullOrWhiteSpace(x)))
            return null;
        return null;
    }

    private static float? CollapseSize(IReadOnlyList<float?> values, float? fallback)
    {
        if (values.Count == 0) return fallback;
        var distinct = values
            .Where(x => x is not null)
            .Select(x => x!.Value)
            .Distinct()
            .ToArray();

        if (distinct.Length == 1) return distinct[0];
        if (distinct.Length == 0 && values.All(x => x is null))
            return null;
        return null;
    }

    private static string PersistedParagraphRawText(XElement paragraph)
    {
        var pieces = new List<string>();
        foreach (var node in paragraph.Descendants())
        {
            if (node.Name == W + "t")
                pieces.Add((string?)node ?? string.Empty);
            else if (node.Name == W + "tab")
                pieces.Add("\t");
            else if (node.Name == W + "br")
                pieces.Add("\n");
        }

        return string.Concat(pieces);
    }

    private static string? NormalizePersistedAlignment(string? value)
    {
        var token = value?.Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(token)) return null;
        return token switch
        {
            "right" or "end" => "RIGHT",
            "left" or "start" => "LEFT",
            "center" => "CENTER",
            "both" => "JUSTIFY",
            "distribute" => "DISTRIBUTE",
            _ => token.ToUpperInvariant()
        };
    }

    private static bool? ParsePersistedOnOff(XElement? element)
    {
        if (element is null) return null;
        var value = ((string?)element.Attribute(W + "val"))?.Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(value)) return true;
        return value switch
        {
            "0" or "false" or "off" or "no" => false,
            _ => true
        };
    }

    private sealed record PersistedParagraphEntry(
        int ParagraphIndex,
        XElement Paragraph,
        bool InTable,
        int? TableRowIndex);

    private sealed record PersistedStyleNode(
        string StyleId,
        string? Name,
        string? BasedOn,
        XElement Element);

    private sealed record PersistedParagraphFormatting(
        string? FontBi,
        float? SizeBi,
        string ReadingOrder,
        string Alignment);

    private static ProfileComplianceSnapshot InspectOpenDocument(
        Word.Document doc,
        AgentProfile profile,
        IReadOnlyDictionary<string, string> aliasToRole,
        IReadOnlyList<StyleComplianceItem> styleCompliance,
        IReadOnlyDictionary<string, StyleComplianceItem> styleStateByName,
        TocOpenXmlAuthoritySnapshot tocOpenXmlAuthority,
        List<string> warnings)
    {
        var roleBuilders = new Dictionary<string, RoleBuilder>(StringComparer.OrdinalIgnoreCase);
        var unclassifiedSamples = new List<UnclassifiedParagraphSample>();
        int nonEmpty = 0, classified = 0, unclassified = 0, inTables = 0, overrideSuspected = 0;
        int recoveredOutlineChapter = 0, recoveredOutlineHeading1 = 0, recoveredDegraded = 0, recoveryConflicts = 0;
        int frontMatterOrdinal = 0;
        bool frontMatterClosed = false;
        bool lastNonEmptyParagraphWasInTable = false;

        Word.Paragraphs? paragraphs = null;
        try
        {
            paragraphs = doc.Paragraphs;
            var degradedTocHeadingAuthority = BuildDegradedTocSourceHeadingAuthority(
                paragraphs,
                profile,
                aliasToRole,
                tocOpenXmlAuthority,
                warnings);

            for (int i = 1; i <= paragraphs.Count; i++)
            {
                Word.Paragraph? paragraph = null;
                Word.Range? range = null;
                Word.Range? formatRange = null;
                Word.Style? style = null;
                Word.Font? font = null;
                Word.ParagraphFormat? paragraphFormat = null;

                try
                {
                    paragraph = paragraphs[i];
                    range = paragraph.Range;
                    var rawText = range.Text ?? string.Empty;
                    var cleanText = WordAuditService.CleanParagraphText(rawText);
                    if (string.IsNullOrWhiteSpace(cleanText)) continue;
                    nonEmpty++;

                    var inTable = IsInsideTable(range);
                    var followsTableBoundary = !inTable && lastNonEmptyParagraphWasInTable;
                    lastNonEmptyParagraphWasInTable = inTable;
                    if (inTable)
                    {
                        inTables++;
                        frontMatterClosed = true;
                    }
                    var tableRow = inTable ? GetTableRowIndex(range) : null;

                    var styleName = GetStyleName(paragraph, range, out style);
                    paragraphFormat ??= range.ParagraphFormat;
                    var outlineHeadingRole = ResolveSourceHeadingRoleFromOutlineLevel(
                        paragraphFormat.OutlineLevel,
                        profile,
                        rawText,
                        cleanText);

                    var exactTocTitle = IsExactPersianTocTitle(cleanText);
                    if (exactTocTitle)
                        frontMatterClosed = true;

                    var invocationMatch =
                        !string.IsNullOrWhiteSpace(profile.SemanticPolicy.Invocation.Role)
                        && SemanticPolicyMatcher.MatchesInvocation(
                            cleanText,
                            profile.SemanticPolicy.Invocation);

                    string? frontMatterRole = null;
                    if (!frontMatterClosed && !inTable && !invocationMatch)
                    {
                        frontMatterOrdinal++;
                        frontMatterRole = ResolveFrontMatterRole(
                            profile,
                            frontMatterOrdinal);
                    }

                    // Word-managed TOC results can lose their paragraph style during severe
                    // manual corruption. The persisted standard-TOC authority is stronger than
                    // the damaged current style. A TOC result contains a tab before its page number.
                    var tocAuthority = ResolveAutomaticTocAuthorityByWordResult(
                        tocOpenXmlAuthority,
                        styleName,
                        rawText);

                    // Outside the TOC, an exact stable-title match back to one unique TOC entry
                    // is strong structural evidence that the source paragraph is a heading.
                    var structuralHeadingRole = tocAuthority is null
                        ? ResolveSourceHeadingRoleFromToc(
                            tocOpenXmlAuthority,
                            rawText,
                            cleanText,
                            profile)
                        : null;

                    // R8R23R11: when severe style collapse has removed the automatic TOC
                    // control/field authority entirely, the still-bounded degraded plain-text
                    // TOC block remains a stronger semantic source-title authority than current
                    // paragraph formatting. Reuse only exact title matches whose stale TOC
                    // paragraph Style still maps unambiguously to Toc1/Toc2. This recovers
                    // Chapter/Heading1 role identity without inventing heading heuristics.
                    var degradedTocHeadingRole = tocAuthority is null && structuralHeadingRole is null
                        ? ResolveSourceHeadingRoleFromDegradedToc(
                            degradedTocHeadingAuthority,
                            rawText,
                            cleanText)
                        : null;

                    string? recoveredHeadingRole = null;
                    if (tocAuthority is null && structuralHeadingRole is null)
                    {
                        if (outlineHeadingRole is not null
                            && degradedTocHeadingRole is not null
                            && !outlineHeadingRole.Equals(degradedTocHeadingRole, StringComparison.OrdinalIgnoreCase))
                        {
                            recoveryConflicts++;
                            if (warnings.Count < 80)
                                warnings.Add(
                                    $"HEADING_RECOVERY_CONFLICT:P={i}:OUTLINE={outlineHeadingRole}:DEGRADED={degradedTocHeadingRole}:TITLE={Preview(cleanText)}");
                        }
                        else
                        {
                            recoveredHeadingRole = outlineHeadingRole ?? degradedTocHeadingRole;
                            if (recoveredHeadingRole is not null)
                            {
                                if (outlineHeadingRole is not null)
                                {
                                    if (recoveredHeadingRole.Equals("Chapter", StringComparison.OrdinalIgnoreCase))
                                        recoveredOutlineChapter++;
                                    else if (recoveredHeadingRole.Equals("Heading1", StringComparison.OrdinalIgnoreCase))
                                        recoveredOutlineHeading1++;
                                }
                                if (degradedTocHeadingRole is not null)
                                    recoveredDegraded++;
                            }
                        }
                    }

                    var role = tocAuthority?.Role
                        ?? structuralHeadingRole
                        ?? recoveredHeadingRole
                        ?? frontMatterRole
                        ?? ResolveRole(styleName, inTable, followsTableBoundary, tableRow, i, cleanText, profile, aliasToRole);

                    if (role is null || !profile.RoleDefaults.TryGetValue(role, out var expected))
                    {
                        unclassified++;
                        if (unclassifiedSamples.Count < MaxUnclassifiedSamples)
                            unclassifiedSamples.Add(new UnclassifiedParagraphSample(i, styleName ?? "(unknown)", Preview(cleanText)));
                        continue;
                    }

                    classified++;
                    if (!roleBuilders.TryGetValue(role, out var builder))
                    {
                        builder = new RoleBuilder(role, expected);
                        roleBuilders[role] = builder;
                    }

                    formatRange = CreateTextOnlyRange(range, rawText);
                    font = formatRange.Font;

                    var actualFont = NormalizeFontName(font.NameBi);
                    var actualSize = NormalizeSize(font.SizeBi);
                    var actualReading = ToReadingOrder(paragraphFormat.ReadingOrder);
                    var actualAlignment = ToAlignment(paragraphFormat.Alignment);

                    styleStateByName.TryGetValue(styleName ?? string.Empty, out var styleState);
                    var sameRoleStyle = styleState is not null && string.Equals(styleState.Role, role, StringComparison.OrdinalIgnoreCase);

                    // QH01-E219 / A15 architecture, corrected by QH01-B009:
                    // Automatic TOC field results are Word-managed content. Persisted OOXML remains
                    // the reading-order authority because Word COM can report stale LTR values.
                    // TOC alignment authority follows the literal persisted Word w:jc token.
                    // w:bidi remains the independent RTL reading-order authority.
                    // Dedicated TOC visual-layout repair owns alignment; generic Safe-Fix must not
                    // re-enumerate Word-managed TOC paragraphs as ordinary actionable candidates.
                    tocAuthority ??= ResolveAutomaticTocAuthority(
                        tocOpenXmlAuthority,
                        role,
                        styleName,
                        rawText);

                    if (tocAuthority is not null)
                    {
                        actualReading = tocAuthority.ReadingOrder;
                        actualAlignment = tocAuthority.Alignment;
                    }

                    // Null profile values mean "(any)" and are unconditional matches.
                    // They must not become false merely because Word reports null/mixed or
                    // the paragraph style maps to another semantic role (e.g. Normal in TableBody).
                    var fontMatch = string.IsNullOrWhiteSpace(expected.FontBi)
                        || (actualFont is not null
                            ? MatchesFont(expected.FontBi, actualFont)
                            : sameRoleStyle && styleState!.FontBiMatch);

                    var sizeMatch = expected.SizeBi is null
                        || (actualSize is not null
                            ? MatchesSize(expected.SizeBi, actualSize)
                            : sameRoleStyle && styleState!.SizeBiMatch);

                    var readingMatch = string.IsNullOrWhiteSpace(expected.ReadingOrder)
                        || (tocAuthority is not null
                            ? tocAuthority.ReadingOrderMatch && tocAuthority.TextRunsRtlMatch
                            : !IsUnknownToken(actualReading)
                                ? MatchesToken(expected.ReadingOrder, actualReading)
                                : sameRoleStyle && styleState!.ReadingOrderMatch);

                    var alignmentMatch = string.IsNullOrWhiteSpace(expected.Alignment)
                        || (tocAuthority is not null
                            ? true
                            : !IsUnknownToken(actualAlignment)
                                ? MatchesToken(expected.Alignment, actualAlignment)
                                : sameRoleStyle && styleState!.AlignmentMatch);

                    var semanticTableNoteStyleRequired =
                        role.Equals("TableNote", StringComparison.OrdinalIgnoreCase)
                        && (inTable || followsTableBoundary)
                        && IsSemanticTableNote(cleanText);

                    // R8R23R9: Chapter/Heading1 style identity is semantic authority, not
                    // merely a visual preference. Automatic TOC compilation depends on these
                    // source paragraphs being materialized onto one of the profile RoleAlias
                    // styles. A structurally classified heading with matching font/size/RTL
                    // must therefore remain actionable until its paragraph style belongs to
                    // the same semantic role.
                    var semanticHeadingStyleRequired =
                        role.Equals("Chapter", StringComparison.OrdinalIgnoreCase)
                        || role.Equals("Heading1", StringComparison.OrdinalIgnoreCase);

                    var semanticStyleIdentityRequired =
                        semanticTableNoteStyleRequired || semanticHeadingStyleRequired;

                    var semanticStyleIdentityMatch =
                        !semanticStyleIdentityRequired
                        || (aliasToRole.TryGetValue(styleName ?? string.Empty, out var semanticStyleRole)
                            && semanticStyleRole.Equals(role, StringComparison.OrdinalIgnoreCase));

                    var fullMatch =
                        fontMatch && sizeMatch && readingMatch && alignmentMatch
                        && semanticStyleIdentityMatch;

                    var directOverride = false;
                    if (!inTable && sameRoleStyle && tocAuthority is null)
                    {
                        directOverride = KnownFormatDifference(
                            styleState!.ActualFontBi,
                            styleState.ActualSizeBi,
                            styleState.ActualReadingOrder,
                            styleState.ActualAlignment,
                            actualFont,
                            actualSize,
                            actualReading,
                            actualAlignment);
                        if (directOverride) overrideSuspected++;
                    }

                    if (semanticTableNoteStyleRequired
                        && !semanticStyleIdentityMatch
                        && !directOverride)
                    {
                        directOverride = true;
                        overrideSuspected++;
                    }

                    builder.Add(fontMatch, sizeMatch, readingMatch, alignmentMatch, fullMatch, directOverride);
                    // QH01-E209: Samples is machine-action authority for Controlled Fix, not only a UI preview.
                    // Keep every actionable non-compliant paragraph so fail-closed mutation can enumerate the complete set.
                    if (!fullMatch || directOverride)
                    {
                        builder.Samples.Add(new ComplianceSample(
                            i,
                            styleName ?? "(unknown)",
                            Preview(cleanText),
                            actualFont,
                            actualSize,
                            actualReading,
                            actualAlignment,
                            directOverride,
                            ExpectedSummary(expected)));
                    }
                }
                catch (Exception ex)
                {
                    if (warnings.Count < 80) warnings.Add($"Compliance paragraph {i}: {ex.Message}");
                }
                finally
                {
                    ComUtil.Release(paragraphFormat);
                    ComUtil.Release(font);
                    ComUtil.Release(style);
                    ComUtil.Release(formatRange);
                    ComUtil.Release(range);
                    ComUtil.Release(paragraph);
                }
            }
        }
        finally { ComUtil.Release(paragraphs); }

        if (warnings.Count < 80)
        {
            warnings.Add(
                $"HEADING_RECOVERY_STRUCTURAL_SUMMARY:OUTLINE_CHAPTER={recoveredOutlineChapter}:"
                + $"OUTLINE_HEADING1={recoveredOutlineHeading1}:DEGRADED_MATCHES={recoveredDegraded}:"
                + $"CONFLICTS={recoveryConflicts}");
        }

        var roleSummaries = roleBuilders.Values
            .OrderBy(x => x.Role, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.ToSummary())
            .ToArray();

        return new ProfileComplianceSnapshot(
            profile.Id,
            profile.DisplayName,
            nonEmpty,
            classified,
            unclassified,
            inTables,
            overrideSuspected,
            roleSummaries,
            styleCompliance,
            unclassifiedSamples,
            warnings);
    }

    private static IReadOnlyList<StyleComplianceItem> InspectStylesFromOpenXml(
        string inputPath,
        AgentProfile profile,
        IReadOnlyDictionary<string, string> aliasToRole,
        List<string> warnings)
    {
        try
        {
            var catalog = OpenXmlStyleCatalog.Load(inputPath);
            var result = new List<StyleComplianceItem>();

            foreach (var style in catalog.Styles.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
            {
                if (!aliasToRole.TryGetValue(style.Name, out var role)) continue;
                if (!profile.RoleDefaults.TryGetValue(role, out var expected)) continue;

                var resolved = catalog.Resolve(style.StyleId);
                result.Add(new StyleComplianceItem(
                    role,
                    style.Name,
                    true,
                    resolved.FontBi,
                    resolved.SizeBi,
                    resolved.ReadingOrder,
                    resolved.Alignment,
                    MatchesFont(expected.FontBi, resolved.FontBi),
                    MatchesSize(expected.SizeBi, resolved.SizeBi),
                    MatchesToken(expected.ReadingOrder, resolved.ReadingOrder),
                    MatchesToken(expected.Alignment, resolved.Alignment)));
            }

            return result
                .OrderBy(x => x.Role, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.StyleName, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
        catch (Exception ex)
        {
            if (warnings.Count < 80) warnings.Add("OOXML style compliance: " + ex.Message);
            return Array.Empty<StyleComplianceItem>();
        }
    }

    // R5R1 compile integration:
    // the earlier helper belongs to TocOpenXmlComplianceAuthorityService,
    // so ProfileComplianceService needs its own class-local canonicalizer.
    private static string? CanonicalizeTocRole(string role)
    {
        if (role.Equals("Toc1", StringComparison.OrdinalIgnoreCase)
            || role.Equals("TocLegacy1", StringComparison.OrdinalIgnoreCase))
            return "Toc1";

        if (role.Equals("Toc2", StringComparison.OrdinalIgnoreCase)
            || role.Equals("TocLegacy2", StringComparison.OrdinalIgnoreCase))
            return "Toc2";

        return null;
    }

    private static Dictionary<string, string> BuildAliasToRole(AgentProfile profile)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (role, aliases) in profile.RoleAliases)
        {
            var effectiveRole = CanonicalizeTocRole(role) ?? role;
            foreach (var alias in aliases ?? Array.Empty<string>())
            {
                if (!string.IsNullOrWhiteSpace(alias)) map[alias.Trim()] = effectiveRole;
            }
        }
        return map;
    }

    private static TocOpenXmlParagraphAuthority? ResolveAutomaticTocAuthorityByWordResult(
        TocOpenXmlAuthoritySnapshot snapshot,
        string? styleName,
        string rawWordText)
    {
        // A generated TOC entry uses a tab before the page number. Requiring the tab
        // prevents a source heading with the same title from being mistaken for the
        // field-managed TOC result.
        if (string.IsNullOrEmpty(rawWordText) || !rawWordText.Contains('\t'))
            return null;

        var stableTitle = StableTocTitleFromWord(rawWordText);
        if (stableTitle.Length == 0)
            return null;

        var matches = snapshot.Items
            .Where(x =>
                string.Equals(x.StableTitle, stableTitle, StringComparison.Ordinal)
                && (string.IsNullOrWhiteSpace(styleName)
                    || string.Equals(x.StyleName, styleName, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        return matches.Count == 1 ? matches[0] : null;
    }

    private static IReadOnlyDictionary<string, string> BuildDegradedTocSourceHeadingAuthority(
        Word.Paragraphs paragraphs,
        AgentProfile profile,
        IReadOnlyDictionary<string, string> aliasToRole,
        TocOpenXmlAuthoritySnapshot tocOpenXmlAuthority,
        List<string> warnings)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);

        if (!string.Equals(profile.Id, "proposal.governance.fa", StringComparison.OrdinalIgnoreCase)
            || tocOpenXmlAuthority.StandardTocControls != 0
            || tocOpenXmlAuthority.Paragraphs != 0)
            return result;

        var rawAuthorities = new List<(string Title, string Role, int ParagraphIndex)>();
        var titleIndex = 0;
        var staleStarted = false;
        var staleCount = 0;
        var mappedCount = 0;
        var styleMappedCount = 0;
        var numberingMappedCount = 0;
        var authorityConflictCount = 0;
        var unresolvedStyleCount = 0;

        for (var i = 1; i <= paragraphs.Count; i++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.Style? style = null;
            try
            {
                paragraph = paragraphs[i];
                range = paragraph.Range;
                var rawText = range.Text ?? string.Empty;
                var cleanText = WordAuditService.CleanParagraphText(rawText);

                if (titleIndex == 0)
                {
                    if (IsExactPersianTocTitle(cleanText))
                        titleIndex = i;
                    continue;
                }

                if (i <= titleIndex)
                    continue;

                if (LooksLikeDegradedTocEntryForSemanticRecovery(rawText))
                {
                    staleStarted = true;
                    staleCount++;
                    var styleName = GetStyleName(paragraph, range, out style) ?? string.Empty;
                    var stableTitle = StableTocTitleFromWord(rawText);
                    var mappedRole = aliasToRole.TryGetValue(styleName, out var rawRole)
                        ? CanonicalizeTocRole(rawRole)
                        : null;
                    var styleSourceRole = mappedRole switch
                    {
                        "Toc1" when profile.RoleDefaults.ContainsKey("Chapter") => "Chapter",
                        "Toc2" when profile.RoleDefaults.ContainsKey("Heading1") => "Heading1",
                        _ => null
                    };

                    var inferredTocRole = TocOpenXmlComplianceAuthorityService.InferTocRoleFromStableTitle(
                        stableTitle,
                        profile);
                    var numberingSourceRole = inferredTocRole switch
                    {
                        "Toc1" when profile.RoleDefaults.ContainsKey("Chapter") => "Chapter",
                        "Toc2" when profile.RoleDefaults.ContainsKey("Heading1") => "Heading1",
                        _ => null
                    };

                    string? sourceRole = null;
                    if (styleSourceRole is not null
                        && numberingSourceRole is not null
                        && !styleSourceRole.Equals(numberingSourceRole, StringComparison.OrdinalIgnoreCase))
                    {
                        authorityConflictCount++;
                    }
                    else
                    {
                        sourceRole = styleSourceRole ?? numberingSourceRole;
                    }

                    if (!string.IsNullOrWhiteSpace(stableTitle) && sourceRole is not null)
                    {
                        mappedCount++;
                        if (styleSourceRole is not null) styleMappedCount++;
                        if (styleSourceRole is null && numberingSourceRole is not null) numberingMappedCount++;
                        rawAuthorities.Add((stableTitle, sourceRole, i));
                    }
                    else
                    {
                        unresolvedStyleCount++;
                    }
                    continue;
                }

                var clean = WordAuditService.CleanParagraphText(rawText);
                if (string.IsNullOrWhiteSpace(clean) && !staleStarted)
                    continue;

                if (staleStarted)
                    break;

                break;
            }
            catch (Exception ex)
            {
                if (warnings.Count < 80)
                    warnings.Add($"Degraded TOC authority P{i}: {ex.Message}");
            }
            finally
            {
                ComUtil.Release(style);
                ComUtil.Release(range);
                ComUtil.Release(paragraph);
            }
        }

        var ambiguousTitles = 0;
        foreach (var group in rawAuthorities.GroupBy(x => x.Title, StringComparer.Ordinal))
        {
            var roles = group.Select(x => x.Role)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (roles.Length == 1)
                result[group.Key] = roles[0];
            else
                ambiguousTitles++;
        }

        if (warnings.Count < 80)
        {
            warnings.Add(
                $"DEGRADED_TOC_SOURCE_HEADING_AUTHORITY:TITLE_P={titleIndex}:STALE={staleCount}:" +
                $"MAPPED={mappedCount}:STYLE_MAPPED={styleMappedCount}:NUMBERING_MAPPED={numberingMappedCount}:" +
                $"AUTHORITY_CONFLICTS={authorityConflictCount}:UNRESOLVED={unresolvedStyleCount}:" +
                $"AMBIGUOUS_TITLES={ambiguousTitles}:UNIQUE_TITLES={result.Count}");
        }

        return result;
    }

    private static bool LooksLikeDegradedTocEntryForSemanticRecovery(string? raw)
    {
        var text = (raw ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\a", string.Empty)
            .Replace("\v", string.Empty)
            .Trim();
        if (text.Length == 0 || !text.Contains('\t'))
            return false;

        var tail = text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .LastOrDefault() ?? string.Empty;
        return tail.Length > 0 && tail.All(ch =>
            char.IsDigit(ch)
            || (ch >= '\u06F0' && ch <= '\u06F9')
            || (ch >= '\u0660' && ch <= '\u0669'));
    }

    private static string? ResolveSourceHeadingRoleFromOutlineLevel(
        Word.WdOutlineLevel outlineLevel,
        AgentProfile profile,
        string rawWordText,
        string cleanText)
    {
        if (string.IsNullOrWhiteSpace(cleanText)
            || (!string.IsNullOrEmpty(rawWordText) && rawWordText.Contains('\t'))
            || IsExactPersianTocTitle(cleanText))
            return null;

        if (outlineLevel == Word.WdOutlineLevel.wdOutlineLevel1
            && profile.RoleDefaults.ContainsKey("Chapter"))
            return "Chapter";

        if (outlineLevel == Word.WdOutlineLevel.wdOutlineLevel2
            && profile.RoleDefaults.ContainsKey("Heading1"))
            return "Heading1";

        return null;
    }

    private static string? ResolveSourceHeadingRoleFromDegradedToc(
        IReadOnlyDictionary<string, string> authority,
        string rawWordText,
        string cleanText)
    {
        if (authority.Count == 0
            || string.IsNullOrWhiteSpace(cleanText)
            || (!string.IsNullOrEmpty(rawWordText) && rawWordText.Contains('\t')))
            return null;

        var stableTitle = StableTocTitleFromWord(cleanText);
        return stableTitle.Length > 0 && authority.TryGetValue(stableTitle, out var role)
            ? role
            : null;
    }

    private static string? ResolveSourceHeadingRoleFromToc(
        TocOpenXmlAuthoritySnapshot snapshot,
        string rawWordText,
        string cleanText,
        AgentProfile profile)
    {
        if (string.IsNullOrWhiteSpace(cleanText)
            || (!string.IsNullOrEmpty(rawWordText) && rawWordText.Contains('\t')))
            return null;

        var stableTitle = StableTocTitleFromWord(cleanText);
        if (stableTitle.Length == 0)
            return null;

        var roles = snapshot.Items
            .Where(x => string.Equals(x.StableTitle, stableTitle, StringComparison.Ordinal))
            .Select(x => x.Role)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (roles.Length != 1)
            return null;

        if (roles[0].Equals("Toc1", StringComparison.OrdinalIgnoreCase)
            && profile.RoleDefaults.ContainsKey("Chapter"))
            return "Chapter";

        if (roles[0].Equals("Toc2", StringComparison.OrdinalIgnoreCase)
            && profile.RoleDefaults.ContainsKey("Heading1"))
            return "Heading1";

        return null;
    }

    private static TocOpenXmlParagraphAuthority? ResolveAutomaticTocAuthority(
        TocOpenXmlAuthoritySnapshot snapshot,
        string role,
        string? styleName,
        string rawWordText)
    {
        if (!(role.Equals("Toc1", StringComparison.OrdinalIgnoreCase)
              || role.Equals("Toc2", StringComparison.OrdinalIgnoreCase)))
            return null;

        var stableTitle = StableTocTitleFromWord(rawWordText);
        if (stableTitle.Length == 0)
            return null;

        var matches = snapshot.Items
            .Where(x => x.Role.Equals(role, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(x.StyleName, styleName ?? string.Empty, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(x.StableTitle, stableTitle, StringComparison.Ordinal))
            .ToList();

        return matches.Count == 1 ? matches[0] : null;
    }

    private static string StableTocTitleFromWord(string rawText)
    {
        var text = rawText ?? string.Empty;
        var tab = text.IndexOf('\t');
        if (tab >= 0)
            text = text[..tab];

        text = text
            .Replace('\r', ' ')
            .Replace('\a', ' ')
            .Replace('\u200c', ' ')
            .Replace('\u200f', ' ')
            .Replace('\u200e', ' ');

        return string.Join(
            ' ',
            text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static bool IsPersistedMainTocFieldParagraph(XElement paragraph)
    {
        var instruction = string.Join(
            " ",
            paragraph.Descendants(W + "instrText")
                .Select(x => ((string?)x ?? string.Empty).Trim()));
        var normalized = string.Join(
            " ",
            instruction.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        return normalized.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)
               && (normalized.Length == 3 || char.IsWhiteSpace(normalized[3]));
    }

    private static bool IsExactPersianTocTitle(string cleanText)
    {
        var compact = string.Concat((cleanText ?? string.Empty)
            .Where(ch => !char.IsWhiteSpace(ch)));
        return string.Equals(compact, "فهرستمطالب", StringComparison.Ordinal);
    }

    private static string? ResolveFrontMatterRole(
        AgentProfile profile,
        int nonInvocationOrdinal)
    {
        if (nonInvocationOrdinal <= 0)
            return null;

        // Supported Persian profile convention:
        // the first two non-invocation, non-table front-matter lines are CoverText.
        if (nonInvocationOrdinal <= 2
            && profile.RoleDefaults.ContainsKey("CoverText"))
            return "CoverText";

        var coverSubtitleRole = profile.SemanticPolicy.CoverSubtitle.Role;
        if (nonInvocationOrdinal == 3
            && !string.IsNullOrWhiteSpace(coverSubtitleRole)
            && profile.RoleDefaults.ContainsKey(coverSubtitleRole))
            return coverSubtitleRole;

        return null;
    }

    private static bool IsSemanticTableNote(string cleanText)
    {
        var normalized = string.Join(
            ' ',
            (cleanText ?? string.Empty)
                .Replace('\u200c', ' ')
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

        return normalized.StartsWith("منبع اتکا:", StringComparison.Ordinal)
               || normalized.StartsWith("یادداشت:", StringComparison.Ordinal)
               || normalized.StartsWith("توضیح:", StringComparison.Ordinal);
    }

    private static string? ResolveRole(
        string? styleName,
        bool inTable,
        bool followsTableBoundary,
        int? tableRow,
        int paragraphIndex,
        string cleanText,
        AgentProfile profile,
        IReadOnlyDictionary<string, string> aliasToRole)
    {
        // Strong TableNote cues are authoritative inside a table or in the
        // first non-empty paragraph immediately following a table.
        if ((inTable || followsTableBoundary)
            && profile.RoleDefaults.ContainsKey("TableNote")
            && IsSemanticTableNote(cleanText))
            return "TableNote";

        if (inTable)
        {
            var tableRole = SemanticPolicyMatcher.ResolveTableRole(
                tableRow,
                profile.SemanticPolicy.TableRoles);

            if (!string.IsNullOrWhiteSpace(tableRole) &&
                profile.RoleDefaults.ContainsKey(tableRole))
                return tableRole;
        }

        var invocationRole = profile.SemanticPolicy.Invocation.Role;
        if (!string.IsNullOrWhiteSpace(invocationRole) &&
            profile.RoleDefaults.ContainsKey(invocationRole) &&
            SemanticPolicyMatcher.MatchesInvocation(cleanText, profile.SemanticPolicy.Invocation))
            return invocationRole;

        var coverRole = profile.SemanticPolicy.CoverSubtitle.Role;
        if (!string.IsNullOrWhiteSpace(coverRole) &&
            profile.RoleDefaults.ContainsKey(coverRole) &&
            SemanticPolicyMatcher.MatchesCoverSubtitle(
                styleName,
                paragraphIndex,
                profile.SemanticPolicy.CoverSubtitle))
            return coverRole;

        // A standard Persian TOC title can lose its style while the actual TOC control
        // remains intact. Exact text is deterministic and independent of formatting.
        if (profile.RoleDefaults.ContainsKey("TocTitle")
            && IsExactPersianTocTitle(cleanText))
            return "TocTitle";

        // Semantic guard: a caption can have lost its paragraph style through manual editing.
        // Document-family caption cues are profile-driven; matching remains deterministic.
        var captionRole = profile.SemanticPolicy.FigureTableCaption.Role;
        if (!string.IsNullOrWhiteSpace(captionRole) &&
            profile.RoleDefaults.ContainsKey(captionRole) &&
            SemanticPolicyMatcher.MatchesFigureTableCaption(
                cleanText,
                profile.SemanticPolicy.FigureTableCaption))
            return captionRole;

        if (!string.IsNullOrWhiteSpace(styleName) && aliasToRole.TryGetValue(styleName, out var role)) return role;
        return null;
    }



    private static bool IsInsideTable(Word.Range range)
    {
        try { return Convert.ToBoolean(range.Information[Word.WdInformation.wdWithInTable], CultureInfo.InvariantCulture); }
        catch { return false; }
    }

    private static int? GetTableRowIndex(Word.Range range)
    {
        try
        {
            var value = range.Information[Word.WdInformation.wdStartOfRangeRowNumber];
            var row = Convert.ToInt32(value, CultureInfo.InvariantCulture);
            return row > 0 ? row : null;
        }
        catch { return null; }
    }

    private static string? GetStyleName(Word.Paragraph paragraph, Word.Range range, out Word.Style? style)
    {
        style = null;
        object? styleObject = null;

        // ParagraphStyle is the strongest signal for field-managed paragraphs such as TOC entries.
        // In Word, Range.Style may be null/mixed while Range.ParagraphStyle still reports TOC 1/2.
        try { styleObject = range.ParagraphStyle; } catch { }

        if (styleObject is null)
        {
            try { styleObject = paragraph.get_Style(); } catch { }
        }

        if (styleObject is null)
        {
            try { styleObject = range.get_Style(); } catch { }
        }

        style = styleObject as Word.Style;
        if (style is not null)
        {
            try { return style.NameLocal; } catch { }
        }

        return styleObject?.ToString();
    }

    private static bool IsDigitToken(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        foreach (var ch in value)
        {
            if ((ch >= '0' && ch <= '9') || (ch >= '۰' && ch <= '۹') || (ch >= '٠' && ch <= '٩'))
                continue;
            return false;
        }
        return true;
    }

    private static Word.Range CreateTextOnlyRange(Word.Range paragraphRange, string rawText)
    {
        var duplicate = paragraphRange.Duplicate;
        var trim = rawText.EndsWith("\r\a", StringComparison.Ordinal) ? 2 : rawText.EndsWith("\r", StringComparison.Ordinal) ? 1 : 0;
        if (trim > 0 && duplicate.End - duplicate.Start >= trim)
            duplicate.End -= trim;
        return duplicate;
    }

    private static string? NormalizeFontName(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static float? NormalizeSize(float value)
        => value > 0 && value < 1000 ? value : null;

    private static bool MatchesFont(string? expected, string? actual)
        => string.IsNullOrWhiteSpace(expected) || string.Equals(expected.Trim(), actual?.Trim(), StringComparison.OrdinalIgnoreCase);

    private static bool MatchesSize(float? expected, float? actual)
        => expected is null || (actual is not null && Math.Abs(expected.Value - actual.Value) <= 0.11f);

    private static bool MatchesToken(string? expected, string actual)
        => string.IsNullOrWhiteSpace(expected) || string.Equals(expected.Trim(), actual, StringComparison.OrdinalIgnoreCase);

    private static bool KnownFormatDifference(
        string? styleFont,
        float? styleSize,
        string styleReading,
        string styleAlignment,
        string? actualFont,
        float? actualSize,
        string actualReading,
        string actualAlignment)
    {
        if (styleFont is not null && actualFont is not null
            && !string.Equals(styleFont, actualFont, StringComparison.OrdinalIgnoreCase)) return true;

        if (styleSize is not null && actualSize is not null
            && Math.Abs(styleSize.Value - actualSize.Value) > 0.11f) return true;

        if (!IsUnknownToken(styleReading) && !IsUnknownToken(actualReading)
            && !string.Equals(styleReading, actualReading, StringComparison.OrdinalIgnoreCase)) return true;

        if (!IsUnknownToken(styleAlignment) && !IsUnknownToken(actualAlignment)
            && !string.Equals(styleAlignment, actualAlignment, StringComparison.OrdinalIgnoreCase)) return true;

        return false;
    }

    private static bool IsUnknownToken(string? value)
        => string.IsNullOrWhiteSpace(value)
           || string.Equals(value, "UNKNOWN", StringComparison.OrdinalIgnoreCase)
           || string.Equals(value, "MIXED_OR_UNDEFINED", StringComparison.OrdinalIgnoreCase);

    private static string ToReadingOrder(Word.WdReadingOrder value)
        => value == Word.WdReadingOrder.wdReadingOrderRtl ? "RTL"
            : value == Word.WdReadingOrder.wdReadingOrderLtr ? "LTR"
            : "MIXED_OR_UNDEFINED";

    private static string ToAlignment(Word.WdParagraphAlignment value)
        => value switch
        {
            Word.WdParagraphAlignment.wdAlignParagraphRight => "RIGHT",
            Word.WdParagraphAlignment.wdAlignParagraphLeft => "LEFT",
            Word.WdParagraphAlignment.wdAlignParagraphCenter => "CENTER",
            Word.WdParagraphAlignment.wdAlignParagraphJustify => "JUSTIFY",
            Word.WdParagraphAlignment.wdAlignParagraphDistribute => "DISTRIBUTE",
            _ => "MIXED_OR_UNDEFINED"
        };

    private static string ExpectedSummary(RoleDefaults d)
        => $"FontBi={d.FontBi ?? "(any)"}; SizeBi={(d.SizeBi?.ToString("0.##", CultureInfo.InvariantCulture) ?? "(any)")}; ReadingOrder={d.ReadingOrder ?? "(any)"}; Alignment={d.Alignment ?? "(any)"}";

    private static string Preview(string text)
    {
        var compact = string.Join(' ', text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        return compact.Length <= 120 ? compact : compact[..117] + "...";
    }

    private sealed class RoleBuilder
    {
        public string Role { get; }
        public RoleDefaults Expected { get; }
        public int Total { get; private set; }
        public int FontMatches { get; private set; }
        public int SizeMatches { get; private set; }
        public int ReadingMatches { get; private set; }
        public int AlignmentMatches { get; private set; }
        public int FullyCompliant { get; private set; }
        public int DirectOverrideSuspected { get; private set; }
        public List<ComplianceSample> Samples { get; } = new();

        public RoleBuilder(string role, RoleDefaults expected)
        {
            Role = role;
            Expected = expected;
        }

        public void Add(bool font, bool size, bool reading, bool alignment, bool full, bool directOverride)
        {
            Total++;
            if (font) FontMatches++;
            if (size) SizeMatches++;
            if (reading) ReadingMatches++;
            if (alignment) AlignmentMatches++;
            if (full) FullyCompliant++;
            if (directOverride) DirectOverrideSuspected++;
        }

        public RoleComplianceSummary ToSummary()
            => new(
                Role,
                Expected.FontBi,
                Expected.SizeBi,
                Expected.ReadingOrder,
                Expected.Alignment,
                Total,
                FontMatches,
                SizeMatches,
                ReadingMatches,
                AlignmentMatches,
                FullyCompliant,
                DirectOverrideSuspected,
                Samples.ToArray());
    }

    private sealed class OpenXmlStyleCatalog
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private readonly Dictionary<string, StyleNode> _byId;
        private readonly ResolvedStyle _defaults;
        private readonly Dictionary<string, ResolvedStyle> _cache = new(StringComparer.OrdinalIgnoreCase);

        public IReadOnlyCollection<StyleNode> Styles => _byId.Values;

        private OpenXmlStyleCatalog(Dictionary<string, StyleNode> byId, ResolvedStyle defaults)
        {
            _byId = byId;
            _defaults = defaults;
        }

        public static OpenXmlStyleCatalog Load(string docxPath)
        {
            using var zip = ZipFile.OpenRead(docxPath);
            var entry = zip.GetEntry("word/styles.xml") ?? throw new InvalidDataException("word/styles.xml not found.");
            using var stream = entry.Open();
            var xdoc = XDocument.Load(stream, LoadOptions.None);
            var root = xdoc.Root ?? throw new InvalidDataException("styles.xml root not found.");

            var defaults = ReadDefaults(root);
            var byId = new Dictionary<string, StyleNode>(StringComparer.OrdinalIgnoreCase);
            foreach (var style in root.Elements(W + "style"))
            {
                var id = (string?)style.Attribute(W + "styleId");
                var name = (string?)style.Element(W + "name")?.Attribute(W + "val");
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name)) continue;
                var basedOn = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
                byId[id] = new StyleNode(id, name, basedOn, style);
            }
            return new OpenXmlStyleCatalog(byId, defaults);
        }

        public ResolvedStyle Resolve(string styleId)
        {
            if (_cache.TryGetValue(styleId, out var cached)) return cached;
            var visiting = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var resolved = ResolveCore(styleId, visiting);
            _cache[styleId] = resolved;
            return resolved;
        }

        private ResolvedStyle ResolveCore(string styleId, HashSet<string> visiting)
        {
            if (_cache.TryGetValue(styleId, out var cached)) return cached;
            if (!_byId.TryGetValue(styleId, out var node)) return _defaults;
            if (!visiting.Add(styleId)) return _defaults;

            var current = !string.IsNullOrWhiteSpace(node.BasedOn)
                ? ResolveCore(node.BasedOn!, visiting)
                : _defaults;

            current = Apply(current, node.Element);
            visiting.Remove(styleId);
            _cache[styleId] = current;
            return current;
        }

        private static ResolvedStyle ReadDefaults(XElement root)
        {
            var defaults = new ResolvedStyle(null, null, "UNKNOWN", "UNKNOWN");
            var docDefaults = root.Element(W + "docDefaults");
            if (docDefaults is null) return defaults;

            var rPr = docDefaults.Element(W + "rPrDefault")?.Element(W + "rPr");
            var pPr = docDefaults.Element(W + "pPrDefault")?.Element(W + "pPr");
            return ApplyProperties(defaults, rPr, pPr);
        }

        private static ResolvedStyle Apply(ResolvedStyle inherited, XElement style)
            => ApplyProperties(inherited, style.Element(W + "rPr"), style.Element(W + "pPr"));

        private static ResolvedStyle ApplyProperties(ResolvedStyle inherited, XElement? rPr, XElement? pPr)
        {
            var font = inherited.FontBi;
            var size = inherited.SizeBi;
            var reading = inherited.ReadingOrder;
            var alignment = inherited.Alignment;

            if (rPr is not null)
            {
                var rFonts = rPr.Element(W + "rFonts");
                var cs = (string?)rFonts?.Attribute(W + "cs");
                if (!string.IsNullOrWhiteSpace(cs)) font = cs.Trim();

                // Font.SizeBi is the complex-script/BiDi size authority.
                // w:sz must not be substituted for a missing w:szCs. If szCs is absent,
                // retain inherited complex-script size so missing BiDi style materialization
                // is detected and repaired explicitly.
                var szCs = (string?)rPr.Element(W + "szCs")?.Attribute(W + "val");
                if (int.TryParse(szCs, NumberStyles.Integer, CultureInfo.InvariantCulture, out var halfPoints) && halfPoints > 0)
                    size = halfPoints / 2f;
            }

            if (pPr is not null)
            {
                var bidi = ParseOnOff(pPr.Element(W + "bidi"));
                if (bidi is not null) reading = bidi.Value ? "RTL" : "LTR";

                var jc = ((string?)pPr.Element(W + "jc")?.Attribute(W + "val"))?.Trim().ToLowerInvariant();
                if (!string.IsNullOrWhiteSpace(jc))
                {
                    alignment = jc switch
                    {
                        "right" or "end" => "RIGHT",
                        "left" or "start" => "LEFT",
                        "center" => "CENTER",
                        "both" => "JUSTIFY",
                        "distribute" => "DISTRIBUTE",
                        _ => jc.ToUpperInvariant()
                    };
                }
            }

            return new ResolvedStyle(font, size, reading, alignment);
        }

        private static bool? ParseOnOff(XElement? element)
        {
            if (element is null) return null;
            var value = ((string?)element.Attribute(W + "val"))?.Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(value)) return true;
            return value switch
            {
                "0" or "false" or "off" or "no" => false,
                _ => true
            };
        }

        public sealed record StyleNode(string StyleId, string Name, string? BasedOn, XElement Element);
        public sealed record ResolvedStyle(string? FontBi, float? SizeBi, string ReadingOrder, string Alignment);
    }
}

