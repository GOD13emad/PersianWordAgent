using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using WordAgent.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordAgent.WordInterop;

public sealed record ControlledFixAction(
    int ParagraphIndex,
    string Role,
    string Action,
    string SourceStyle,
    string? TargetStyle,
    string TextPreview);

public sealed record ControlledCopyFixResult(
    string InputFile,
    string OutputFile,
    string ProfileId,
    string InputSha256,
    string OutputSha256,
    int CandidateParagraphs,
    int MutatedParagraphs,
    IReadOnlyList<ControlledFixAction> Actions,
    WordAuditSnapshot BeforeWord,
    WordAuditSnapshot AfterWord,
    ProfileComplianceSnapshot BeforeCompliance,
    ProfileComplianceSnapshot AfterCompliance,
    bool WordIntegrityPassed,
    IReadOnlyList<string> IntegrityIssues,
    IReadOnlyList<string> Warnings);

public sealed record ControlledMutationTransaction(
    string InputFile,
    string OutputFile,
    string ProfileId,
    string InputSha256,
    string OutputSha256,
    int CandidateParagraphs,
    int MutatedParagraphs,
    IReadOnlyList<ControlledFixAction> Actions,
    WordAuditSnapshot BeforeWord,
    ProfileComplianceSnapshot BeforeCompliance,
    bool MissingRequiredTocNotApplicableNoSourceHierarchy,
    IReadOnlyList<string> Warnings);

public sealed class ControlledCopyFixService
{
    private const int RpcSServerUnavailable = unchecked((int)0x800706BA);
    private const int RpcSCallFailed = unchecked((int)0x800706BE);
    private const int RpcEDisconnected = unchecked((int)0x80010108);
    private const int RpcECallRejected = unchecked((int)0x80010001);

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public ControlledCopyFixResult Apply(
        string inputPath,
        AgentProfile profile,
        string? outputDirectory = null,
        bool visible = false,
        Func<string, WordAuditSnapshot>? wordInspector = null,
        Func<string, AgentProfile, ProfileComplianceSnapshot>? complianceInspector = null,
        Action<string>? trace = null)
    {
        var transaction = PrepareAndMutate(
            inputPath,
            profile,
            outputDirectory,
            visible,
            wordInspector,
            complianceInspector,
            trace);

        return VerifyTransaction(
            transaction,
            profile,
            visible,
            wordInspector,
            complianceInspector,
            trace);
    }

    private const string TocReconstructionBaselineSuffix =
        ".WordAgent_TocReconstructionBaseline.docx";

    private static string TocReconstructionBaselinePath(string outputPath)
        => outputPath + TocReconstructionBaselineSuffix;

    private static bool ShouldReconstructMissingRequiredToc(
        string inputPath,
        AgentProfile profile,
        WordAuditSnapshot beforeWord,
        ProfileComplianceSnapshot beforeCompliance,
        Action<string>? trace)
    {
        // QH01-E227: AR9 admits UniversityBook to the already-proven controlled
        // missing-TOC transaction. Thesis profiles remain excluded here.
        if (!string.Equals(profile.Id, "proposal.governance.fa", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(profile.Id, "book.university.fa", StringComparison.OrdinalIgnoreCase))
            return false;

        // R8R23R3: one reconstruction authority only.
        // The CLI already treats persisted OOXML as stronger than Word COM/current damaged style.
        // Do not independently veto that decision with legacy Word TocFields or unclassified counts:
        // severe style collapse can leave stale visible TOC text unclassified while no automatic
        // TOC authority remains. Those weaker signals are retained below only as diagnostics.
        var tocAuthority = TocOpenXmlComplianceAuthorityService.Inspect(inputPath, profile);
        var reconstructRequired =
            tocAuthority.StandardTocControls == 0
            && tocAuthority.Paragraphs == 0;

        trace?.Invoke(
            $"TOC_RECONSTRUCTION_DECISION_AUTHORITY:SOURCE=OPENXML:REQUIRED={(reconstructRequired ? 1 : 0)}:" +
            $"CONTROLS={tocAuthority.StandardTocControls}:PARAGRAPHS={tocAuthority.Paragraphs}:" +
            $"WORD_TOC_FIELDS={beforeWord.Counts.TocFields}:UNCLASSIFIED={beforeCompliance.UnclassifiedParagraphs}");

        return reconstructRequired;
    }

    private static string CleanWordParagraphTextForTocRecovery(string? raw)
        => (raw ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\a", string.Empty)
            .Replace("\v", string.Empty)
            .Trim();

    private static bool IsExactTocTitleForReconstruction(string? raw)
    {
        var compact = new string(
            CleanWordParagraphTextForTocRecovery(raw)
                .Where(ch => !char.IsWhiteSpace(ch) && ch != '\u200c' && ch != '\u200f' && ch != '\u200e')
                .ToArray());

        return compact.Equals("فهرستمطالب", StringComparison.Ordinal);
    }

    private static bool LooksLikeDegradedTocEntryForReconstruction(string? raw)
    {
        var text = CleanWordParagraphTextForTocRecovery(raw);
        if (string.IsNullOrWhiteSpace(text) || !text.Contains('\t'))
            return false;

        var tail = text
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .LastOrDefault() ?? string.Empty;

        return tail.Length > 0
               && tail.All(ch =>
                   char.IsDigit(ch)
                   || (ch >= '\u06F0' && ch <= '\u06F9')
                   || (ch >= '\u0660' && ch <= '\u0669'));
    }

    private static object ResolveWordStyleObjectForToc(
        dynamic styles,
        string role,
        string preferredName,
        IReadOnlyList<string> equivalentAliases,
        Action<string>? trace)
    {
        var targets = new List<string> { preferredName };
        targets.AddRange(equivalentAliases.Where(x => !string.IsNullOrWhiteSpace(x)));
        var normalizedTargets = targets
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var observed = new List<string>();
        var styleCount = (int)styles.Count;
        for (int i = 1; i <= styleCount; i++)
        {
            object? styleObject = null;
            var keep = false;
            try
            {
                styleObject = styles.Item(i);
                dynamic style = styleObject;
                string nameLocal;
                try { nameLocal = ((string?)style.NameLocal ?? string.Empty).Trim(); }
                catch { nameLocal = string.Empty; }

                if (!string.IsNullOrWhiteSpace(nameLocal) && observed.Count < 80)
                    observed.Add(nameLocal);

                if (normalizedTargets.Any(target => TokensEqual(target, nameLocal)))
                {
                    keep = true;
                    trace?.Invoke(
                        $"TOC_RECONSTRUCTION_WORD_STYLE_RESOLUTION:PASS:ROLE={role}:TARGET={preferredName}:"
                        + $"RESOLVED={nameLocal}:INDEX={i}:STYLE_COUNT={styleCount}");
                    return styleObject!;
                }
            }
            catch (Exception ex)
            {
                if (observed.Count < 80)
                    observed.Add($"<STYLE_INDEX_{i}_ERROR:{ex.Message}>");
            }
            finally
            {
                if (!keep)
                    ReleaseCom(ref styleObject);
            }
        }

        var nearby = observed
            .Where(name => normalizedTargets.Any(target =>
                NormalizeText(name).Contains(NormalizeText(target), StringComparison.OrdinalIgnoreCase)
                || NormalizeText(target).Contains(NormalizeText(name), StringComparison.OrdinalIgnoreCase)))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(12)
            .ToArray();

        trace?.Invoke(
            $"TOC_RECONSTRUCTION_WORD_STYLE_RESOLUTION:FAIL:ROLE={role}:TARGET={preferredName}:"
            + $"ALIASES={string.Join('|', normalizedTargets)}:STYLE_COUNT={styleCount}:"
            + $"NEARBY={string.Join('|', nearby)}");

        throw new InvalidOperationException(
            $"Word COM style identity could not be resolved safely for role '{role}'. "
            + $"Target='{preferredName}';Aliases='{string.Join(",", normalizedTargets)}';StyleCount={styleCount}.");
    }

    private sealed record TocSourceStyleMembership(
        int ChapterCount,
        int Heading1Count,
        IReadOnlyList<int> ChapterSamples,
        IReadOnlyList<int> Heading1Samples,
        IReadOnlyList<string> ChapterStyles,
        IReadOnlyList<string> Heading1Styles,
        int AmbiguousParagraphs,
        IReadOnlyList<int> AmbiguousSamples,
        int StyleReadErrors);

    private static string ReadWordStyleNameLocal(object styleObject)
    {
        try
        {
            dynamic style = styleObject;
            return ((string?)style.NameLocal ?? string.Empty).Trim();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string[] BuildTocRoleStyleTargets(
        string preferredStyleName,
        IReadOnlyList<string> aliases)
    {
        var all = new List<string>();
        if (!string.IsNullOrWhiteSpace(preferredStyleName))
            all.Add(preferredStyleName.Trim());
        all.AddRange((aliases ?? Array.Empty<string>())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim()));

        return all
            .GroupBy(NormalizeText, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToArray();
    }

    private static TocSourceStyleMembership InspectWordTocSourceStyleMembership(
        dynamic paragraphs,
        string chapterPreferredStyleName,
        IReadOnlyList<string> chapterAliases,
        string heading1PreferredStyleName,
        IReadOnlyList<string> heading1Aliases,
        int tocTitleIndex,
        Action<string>? trace)
    {
        var chapterTargets = BuildTocRoleStyleTargets(chapterPreferredStyleName, chapterAliases);
        var heading1Targets = BuildTocRoleStyleTargets(heading1PreferredStyleName, heading1Aliases);

        var overlap = chapterTargets
            .Where(c => heading1Targets.Any(h => TokensEqual(c, h)))
            .ToArray();
        if (overlap.Length > 0)
            throw new InvalidOperationException(
                "Chapter/Heading1 profile alias families overlap and cannot be mapped safely into one TOC. "
                + $"Overlap={string.Join("|", overlap)}.");

        var chapterCount = 0;
        var heading1Count = 0;
        var ambiguousParagraphs = 0;
        var styleReadErrors = 0;
        var chapterSamples = new List<int>();
        var heading1Samples = new List<int>();
        var ambiguousSamples = new List<int>();
        var chapterStyles = new List<string>();
        var heading1Styles = new List<string>();
        var paragraphCount = (int)paragraphs.Count;

        for (var i = 1; i <= paragraphCount; i++)
        {
            if (i == tocTitleIndex)
                continue;

            object? paragraphObject = null;
            object? rangeObject = null;
            object? styleObject = null;
            try
            {
                paragraphObject = paragraphs.Item(i);
                dynamic paragraph = paragraphObject;
                rangeObject = paragraph.Range;
                dynamic range = rangeObject;

                // Match the same paragraph-style authority used by ProfileComplianceService.
                try { styleObject = range.ParagraphStyle; } catch { }
                if (styleObject is null)
                {
                    try { styleObject = paragraph.get_Style(); } catch { }
                }
                if (styleObject is null)
                {
                    try { styleObject = range.get_Style(); } catch { }
                }

                string styleName;
                if (styleObject is string styleString)
                {
                    styleName = styleString.Trim();
                }
                else
                {
                    styleName = styleObject is null ? string.Empty : ReadWordStyleNameLocal(styleObject);
                }

                var isChapter = chapterTargets.Any(target => TokensEqual(styleName, target));
                var isHeading1 = heading1Targets.Any(target => TokensEqual(styleName, target));

                if (isChapter && isHeading1)
                {
                    ambiguousParagraphs++;
                    if (ambiguousSamples.Count < 12) ambiguousSamples.Add(i);
                    continue;
                }

                if (isChapter)
                {
                    chapterCount++;
                    if (chapterSamples.Count < 12) chapterSamples.Add(i);
                    if (!string.IsNullOrWhiteSpace(styleName))
                        chapterStyles.Add(styleName);
                }

                if (isHeading1)
                {
                    heading1Count++;
                    if (heading1Samples.Count < 12) heading1Samples.Add(i);
                    if (!string.IsNullOrWhiteSpace(styleName))
                        heading1Styles.Add(styleName);
                }
            }
            catch
            {
                styleReadErrors++;
            }
            finally
            {
                ReleaseCom(ref styleObject);
                ReleaseCom(ref rangeObject);
                ReleaseCom(ref paragraphObject);
            }
        }

        var chapterStylesDistinct = chapterStyles
            .GroupBy(NormalizeText, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .OrderBy(x => NormalizeText(x), StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var heading1StylesDistinct = heading1Styles
            .GroupBy(NormalizeText, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .OrderBy(x => NormalizeText(x), StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var result = new TocSourceStyleMembership(
            chapterCount,
            heading1Count,
            chapterSamples,
            heading1Samples,
            chapterStylesDistinct,
            heading1StylesDistinct,
            ambiguousParagraphs,
            ambiguousSamples,
            styleReadErrors);

        trace?.Invoke(
            $"TOC_RECONSTRUCTION_SOURCE_STYLE_MEMBERSHIP:"
            + $"MODE=ROLE_ALIAS_FAMILY:"
            + $"CHAPTER={result.ChapterCount}:HEADING1={result.Heading1Count}:"
            + $"CHAPTER_STYLES={string.Join("|", result.ChapterStyles)}:"
            + $"HEADING1_STYLES={string.Join("|", result.Heading1Styles)}:"
            + $"CHAPTER_SAMPLES={string.Join(",", result.ChapterSamples)}:"
            + $"HEADING1_SAMPLES={string.Join(",", result.Heading1Samples)}:"
            + $"AMBIGUOUS={result.AmbiguousParagraphs}:"
            + $"AMBIGUOUS_SAMPLES={string.Join(",", result.AmbiguousSamples)}:"
            + $"STYLE_READ_ERRORS={result.StyleReadErrors}");

        return result;
    }


    private static string ResolveWordTocListSeparator(object appObject, Action<string>? trace)
    {
        string separator;
        try
        {
            var wordApp = (Word.Application)appObject;
            separator = Convert.ToString(
                wordApp.International[Word.WdInternationalIndex.wdListSeparator],
                CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "Word list-separator authority could not be read from Application.International(wdListSeparator).",
                ex);
        }

        if (separator is not ("," or ";"))
            throw new InvalidOperationException(
                $"Unsupported Word TOC list separator '{separator}'. Expected comma or semicolon.");

        trace?.Invoke($"TOC_RECONSTRUCTION_LIST_SEPARATOR:PASS:VALUE={separator}:SOURCE=WORD_APPLICATION_INTERNATIONAL_17");
        return separator;
    }

    private static string BuildWordCustomStyleTocFieldCode(
        IReadOnlyList<string> chapterStyleNames,
        IReadOnlyList<string> heading1StyleNames,
        string listSeparator)
    {
        var chapter = (chapterStyleNames ?? Array.Empty<string>())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .GroupBy(NormalizeText, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First().Trim())
            .ToArray();
        var heading1 = (heading1StyleNames ?? Array.Empty<string>())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .GroupBy(NormalizeText, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First().Trim())
            .ToArray();

        if (chapter.Length == 0 || heading1.Length == 0)
            throw new InvalidOperationException(
                $"Custom TOC role-style mapping is empty. ChapterStyles={chapter.Length};Heading1Styles={heading1.Length}.");

        var overlap = chapter
            .Where(c => heading1.Any(h => TokensEqual(c, h)))
            .ToArray();
        if (overlap.Length > 0)
            throw new InvalidOperationException(
                "The same in-use paragraph style cannot safely map to both Chapter and Heading1 TOC levels. "
                + $"Overlap={string.Join("|", overlap)}.");

        foreach (var styleName in chapter.Concat(heading1))
        {
            if (string.IsNullOrWhiteSpace(styleName)
                || styleName.Contains('"')
                || styleName.Contains('\r')
                || styleName.Contains('\n')
                || styleName.Contains(listSeparator, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    $"Custom TOC style name cannot be encoded safely in the Word TOC \\t switch: '{styleName}'.");
        }

        var tokens = new List<string>();
        foreach (var styleName in chapter)
        {
            tokens.Add(styleName);
            tokens.Add("1");
        }
        foreach (var styleName in heading1)
        {
            tokens.Add(styleName);
            tokens.Add("2");
        }

        return $" TOC \\h \\z \\t \"{string.Join(listSeparator, tokens)}\" ";
    }


    private static void RebuildMissingRequiredTocWithWord(
        string outputPath,
        AgentProfile profile,
        ProfileComplianceSnapshot beforeCompliance,
        WordAuditSnapshot beforeWord,
        List<string> warnings,
        Action<string>? trace)
    {
        var governanceMode =
            string.Equals(profile.Id, "proposal.governance.fa", StringComparison.OrdinalIgnoreCase);
        var universityBookMode =
            string.Equals(profile.Id, "book.university.fa", StringComparison.OrdinalIgnoreCase);
        if (!governanceMode && !universityBookMode)
            throw new InvalidOperationException("TOC reconstruction is not enabled for this profile.");

        var baselinePids = SnapshotWinWordPids();
        if (baselinePids.Length != 0)
            throw new InvalidOperationException(
                "TOC reconstruction blocked because WINWORD already exists. PIDS="
                + string.Join(",", baselinePids)
                + ". Close Word yourself; WordAgent will not kill processes.");

        var tocHierarchy = ResolveControlledTocHierarchyRoles(profile);
        var chapterAliases = ResolveControlledRoleAliases(profile, tocHierarchy.Level1Role);
        var headingAliases = ResolveControlledRoleAliases(profile, tocHierarchy.Level2Role);
        var chapterStyle = ResolveControlledPreferredExistingStyle(
            profile, tocHierarchy.Level1Role, chapterAliases, beforeCompliance);
        var headingStyle = ResolveControlledPreferredExistingStyle(
            profile, tocHierarchy.Level2Role, headingAliases, beforeCompliance);

        if (string.IsNullOrWhiteSpace(chapterStyle) || string.IsNullOrWhiteSpace(headingStyle))
            throw new InvalidOperationException(
                $"TOC reconstruction cannot resolve target styles safely. Level1Role={tocHierarchy.Level1Role};Level2Role={tocHierarchy.Level2Role}");
        trace?.Invoke(
            $"TOC_HIERARCHY_ROLE_AUTHORITY:PASS:LEVEL1={tocHierarchy.Level1Role}:STYLE1={chapterStyle}:LEVEL2={tocHierarchy.Level2Role}:STYLE2={headingStyle}");

        object? appObject = null;
        object? documentsObject = null;
        object? documentObject = null;
        object? paragraphsObject = null;
        object? titleParagraphObject = null;
        object? titleRangeObject = null;
        object? insertionRangeObject = null;
        object? tocCollectionObject = null;
        object? tocObject = null;
        object? documentStylesObject = null;
        object? chapterWordStyleObject = null;
        object? heading1WordStyleObject = null;
        object? tocFieldObject = null;
        object? tocFieldCodeRangeObject = null;
        object? tocRangeObject = null;
        object? tocRangeParagraphsObject = null;
        object? tocRangeFieldsObject = null;
        object? tocRangeHyperlinksObject = null;
        object? tablesObject = null;
        object? sectionsObject = null;

        try
        {
            trace?.Invoke("TOC_RECONSTRUCTION_WORD:START");

            var wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new InvalidOperationException("Microsoft Word COM automation is unavailable.");

            appObject = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Microsoft Word COM instance could not be created.");

            dynamic app = appObject;
            app.Visible = false;
            app.DisplayAlerts = 0;

            documentsObject = app.Documents;
            dynamic documents = documentsObject;
            documentObject = documents.Open(outputPath, false, false, false);
            dynamic document = documentObject;

            if ((bool)document.ReadOnly)
                throw new InvalidOperationException("Controlled output opened read-only during TOC reconstruction.");

            tablesObject = document.Tables;
            sectionsObject = document.Sections;
            dynamic tables = tablesObject;
            dynamic sections = sectionsObject;

            if ((int)tables.Count != beforeWord.Counts.Tables
                || (int)sections.Count != beforeWord.Counts.Sections)
                throw new InvalidOperationException(
                    "Protected table/section counts changed before TOC reconstruction.");

            tocCollectionObject = document.TablesOfContents;
            dynamic tocCollection = tocCollectionObject;
            if ((int)tocCollection.Count != 0)
                throw new InvalidOperationException(
                    $"Missing-TOC reconstruction expected zero TOCs, found={tocCollection.Count}.");

            paragraphsObject = document.Paragraphs;
            dynamic paragraphs = paragraphsObject;
            var titleIndex = 0;
            var staleEntryIndexes = new List<int>();

            for (var i = 1; i <= (int)paragraphs.Count; i++)
            {
                object? paragraphObject = null;
                object? rangeObject = null;
                try
                {
                    paragraphObject = paragraphs.Item(i);
                    dynamic paragraph = paragraphObject;
                    rangeObject = paragraph.Range;
                    dynamic range = rangeObject;
                    var text = (string?)range.Text ?? string.Empty;

                    if (titleIndex == 0)
                    {
                        if (IsExactTocTitleForReconstruction(text))
                            titleIndex = i;
                        continue;
                    }

                    if (i <= titleIndex)
                        continue;

                    if (LooksLikeDegradedTocEntryForReconstruction(text))
                    {
                        staleEntryIndexes.Add(i);
                        continue;
                    }

                    var clean = CleanWordParagraphTextForTocRecovery(text);
                    if (string.IsNullOrWhiteSpace(clean) && staleEntryIndexes.Count == 0)
                        continue;

                    if (staleEntryIndexes.Count > 0)
                        break;

                    // The first non-empty non-TOC line after the title means the expected
                    // degraded TOC block was not found safely.
                    break;
                }
                finally
                {
                    ReleaseCom(ref rangeObject);
                    ReleaseCom(ref paragraphObject);
                }
            }

            if (titleIndex <= 0)
                throw new InvalidOperationException("Exact Persian TOC title was not found for reconstruction.");

            if (governanceMode)
            {
                if (staleEntryIndexes.Count < 2 || staleEntryIndexes.Count > 120)
                    throw new InvalidOperationException(
                        $"Degraded plain-text TOC block is not safely bounded. Entries={staleEntryIndexes.Count}.");
            }
            else
            {
                // QH01-E228: UniversityBook creation is a true missing-TOC case.
                // A pre-existing degraded text block is not silently deleted under this path.
                if (staleEntryIndexes.Count != 0)
                    throw new InvalidOperationException(
                        $"UniversityBook missing-TOC creation requires zero stale plain-text TOC entries. Entries={staleEntryIndexes.Count}.");
            }

            // QH01-E236: evidence labels follow the profile-declared hierarchy instead
            // of legacy Chapter/Heading1 names.
            trace?.Invoke(
                $"TOC_RECONSTRUCTION_PREFLIGHT:PASS:PROFILE={profile.Id}:TITLE_P={titleIndex}:STALE_ENTRIES={staleEntryIndexes.Count}:"
                + $"LEVEL1_ROLE={tocHierarchy.Level1Role}:LEVEL1_STYLE={chapterStyle}:"
                + $"LEVEL2_ROLE={tocHierarchy.Level2Role}:LEVEL2_STYLE={headingStyle}");

            foreach (var index in staleEntryIndexes.OrderByDescending(x => x))
            {
                object? paragraphObject = null;
                object? rangeObject = null;
                try
                {
                    paragraphObject = paragraphs.Item(index);
                    dynamic paragraph = paragraphObject;
                    rangeObject = paragraph.Range;
                    dynamic range = rangeObject;
                    range.Delete();
                }
                finally
                {
                    ReleaseCom(ref rangeObject);
                    ReleaseCom(ref paragraphObject);
                }
            }

            ReleaseCom(ref paragraphsObject);
            paragraphsObject = document.Paragraphs;
            paragraphs = paragraphsObject;

            titleParagraphObject = paragraphs.Item(titleIndex);
            dynamic titleParagraph = titleParagraphObject;
            titleRangeObject = titleParagraph.Range;
            dynamic titleRange = titleRangeObject;
            var insertionPosition = (int)titleRange.End;

            insertionRangeObject = document.Range(insertionPosition, insertionPosition);
            dynamic insertionRange = insertionRangeObject;

            // R8R23R8: preserve the proven Word.Style identity resolution from R6, but
            // fail closed before creating a TOC unless the controlled copy actually contains
            // source paragraphs using those resolved styles. Then compile the TOC from one
            // explicit persisted Word field contract (TOC \\t Style,Level,...) using Word's own
            // regional list separator. This removes the implicit collection-timing ambiguity
            // demonstrated by R6 while keeping the same semantic role authority.
            documentStylesObject = document.Styles;
            dynamic documentStyles = documentStylesObject;
            chapterWordStyleObject = ResolveWordStyleObjectForToc(
                documentStyles, "LEVEL1_" + tocHierarchy.Level1Role.ToUpperInvariant(), chapterStyle, chapterAliases, trace);
            heading1WordStyleObject = ResolveWordStyleObjectForToc(
                documentStyles, "LEVEL2_" + tocHierarchy.Level2Role.ToUpperInvariant(), headingStyle, headingAliases, trace);

            var chapterWordStyleName = ReadWordStyleNameLocal(chapterWordStyleObject);
            var heading1WordStyleName = ReadWordStyleNameLocal(heading1WordStyleObject);
            if (string.IsNullOrWhiteSpace(chapterWordStyleName)
                || string.IsNullOrWhiteSpace(heading1WordStyleName))
                throw new InvalidOperationException(
                    "Resolved Word Style objects did not expose decision-ready NameLocal identities.");

            // R8R23R8R1 compile recovery: force static dispatch so the typed
            // TocSourceStyleMembership return does not become dynamic (CS1977 at downstream LINQ).
            var membership = InspectWordTocSourceStyleMembership(
                (object)paragraphs,
                chapterWordStyleName,
                chapterAliases,
                heading1WordStyleName,
                headingAliases,
                titleIndex,
                trace);
            if (membership.StyleReadErrors != 0)
                throw new InvalidOperationException(
                    $"TOC source-style membership scan had unreadable paragraph Style identities. Errors={membership.StyleReadErrors}.");
            if (membership.AmbiguousParagraphs != 0)
                throw new InvalidOperationException(
                    "Semantic TOC role-style membership is ambiguous between Chapter and Heading1. "
                    + $"Ambiguous={membership.AmbiguousParagraphs};Samples={string.Join(",", membership.AmbiguousSamples)}.");
            if (membership.ChapterCount <= 0 || membership.Heading1Count <= 0)
                throw new InvalidOperationException(
                    "Semantic TOC source-style membership is missing across the controlled profile alias families. "
                    + $"Level1Role={tocHierarchy.Level1Role};Level1={membership.ChapterCount};"
                    + $"Level2Role={tocHierarchy.Level2Role};Level2={membership.Heading1Count};"
                    + $"Level1Styles={string.Join("|", membership.ChapterStyles)};"
                    + $"Level2Styles={string.Join("|", membership.Heading1Styles)}. "
                    + "TOC generation was not attempted.");
            trace?.Invoke(
                $"TOC_RECONSTRUCTION_SOURCE_STYLE_MEMBERSHIP:PASS:MODE=ROLE_ALIAS_FAMILY:"
                + $"LEVEL1_ROLE={tocHierarchy.Level1Role}:LEVEL1={membership.ChapterCount}:"
                + $"LEVEL2_ROLE={tocHierarchy.Level2Role}:LEVEL2={membership.Heading1Count}:"
                + $"LEVEL1_STYLES={string.Join("|", membership.ChapterStyles)}:"
                + $"LEVEL2_STYLES={string.Join("|", membership.Heading1Styles)}");

            var listSeparator = ResolveWordTocListSeparator(appObject, trace);
            var customStyleFieldCode = BuildWordCustomStyleTocFieldCode(
                membership.ChapterStyles,
                membership.Heading1Styles,
                listSeparator);

            tocObject = tocCollection.Add(
                insertionRange,
                false,
                1,
                2,
                false,
                "",
                true,
                true,
                "",
                true,
                false,
                false);

            dynamic toc = tocObject;
            tocRangeObject = toc.Range;
            dynamic tocRangeForField = tocRangeObject;
            tocRangeFieldsObject = tocRangeForField.Fields;
            dynamic tocFieldsForCode = tocRangeFieldsObject;

            // R8R23R12: Range.Fields contains every field whose result/code lies in the
            // TOC range (including entry/page-reference fields). Do not require the
            // collection cardinality to be one. Identify the single outer TOC field
            // by Word's WdFieldType.wdFieldTOC value (13) and fail closed on any
            // ambiguity or unreadable field identity.
            var tocRangeFieldTotal = (int)tocFieldsForCode.Count;
            var outerTocFieldCount = 0;
            var outerTocFieldIndex = 0;
            var fieldTypeReadErrors = 0;
            for (var fieldIndex = 1; fieldIndex <= tocRangeFieldTotal; fieldIndex++)
            {
                object? candidateFieldObject = null;
                try
                {
                    candidateFieldObject = tocFieldsForCode.Item(fieldIndex);
                    dynamic candidateField = candidateFieldObject;
                    var candidateFieldType = (int)candidateField.Type;
                    if (candidateFieldType == 13)
                    {
                        outerTocFieldCount++;
                        outerTocFieldIndex = fieldIndex;
                    }
                }
                catch
                {
                    fieldTypeReadErrors++;
                }
                finally
                {
                    ReleaseCom(ref candidateFieldObject);
                }
            }

            trace?.Invoke(
                $"TOC_RECONSTRUCTION_OUTER_FIELD_DISCOVERY:"
                + $"TOTAL_FIELDS={tocRangeFieldTotal}:TOC_FIELDS={outerTocFieldCount}:"
                + $"TOC_INDEX={outerTocFieldIndex}:TYPE_READ_ERRORS={fieldTypeReadErrors}");

            if (fieldTypeReadErrors != 0)
                throw new InvalidOperationException(
                    $"TOC field identity scan had unreadable Field.Type values. Errors={fieldTypeReadErrors};Total={tocRangeFieldTotal}.");
            if (outerTocFieldCount != 1 || outerTocFieldIndex <= 0)
                throw new InvalidOperationException(
                    $"New TOC container must expose exactly one wdFieldTOC outer field. "
                    + $"TotalFields={tocRangeFieldTotal};TocFields={outerTocFieldCount};Index={outerTocFieldIndex}.");

            trace?.Invoke(
                $"TOC_RECONSTRUCTION_OUTER_FIELD_DISCOVERY:PASS:"
                + $"TOTAL_FIELDS={tocRangeFieldTotal}:TOC_FIELDS=1:TOC_INDEX={outerTocFieldIndex}:TYPE=13");

            tocFieldObject = tocFieldsForCode.Item(outerTocFieldIndex);
            dynamic tocField = tocFieldObject;
            tocFieldCodeRangeObject = tocField.Code;
            dynamic tocFieldCodeRange = tocFieldCodeRangeObject;
            tocFieldCodeRange.Text = customStyleFieldCode;
            var readBackFieldCode = ((string?)tocFieldCodeRange.Text ?? string.Empty).Trim();
            var expectedFieldStyles = membership.ChapterStyles
                .Concat(membership.Heading1Styles)
                .ToArray();
            var missingFieldStyles = expectedFieldStyles
                .Where(styleName =>
                    !NormalizeText(readBackFieldCode).Contains(
                        NormalizeText(styleName),
                        StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (!readBackFieldCode.Contains("\\t", StringComparison.OrdinalIgnoreCase)
                || missingFieldStyles.Length != 0)
                throw new InvalidOperationException(
                    $"Custom-style TOC field-code readback failed. MissingStyles={string.Join("|", missingFieldStyles)};Code='{readBackFieldCode}'.");

            trace?.Invoke(
                $"TOC_RECONSTRUCTION_FIELD_CODE:PASS:SOURCE=CUSTOM_STYLE_T_SWITCH_ROLE_ALIAS_FAMILY:"
                + $"CHAPTER_STYLES={string.Join("|", membership.ChapterStyles)}:LEVEL=1:"
                + $"HEADING1_STYLES={string.Join("|", membership.Heading1Styles)}:LEVEL=2:"
                + $"LIST_SEPARATOR={listSeparator}");

            ReleaseCom(ref tocFieldCodeRangeObject);
            ReleaseCom(ref tocFieldObject);
            ReleaseCom(ref tocRangeFieldsObject);
            ReleaseCom(ref tocRangeObject);

            toc.Update();
            document.Repaginate();
            try { toc.UpdatePageNumbers(); } catch { toc.Update(); }

            if ((int)tocCollection.Count != 1)
                throw new InvalidOperationException(
                    $"TOC reconstruction did not create exactly one automatic TOC. Found={tocCollection.Count}.");

            tocRangeObject = toc.Range;
            dynamic tocRange = tocRangeObject;
            tocRangeParagraphsObject = tocRange.Paragraphs;
            tocRangeFieldsObject = tocRange.Fields;
            tocRangeHyperlinksObject = tocRange.Hyperlinks;
            dynamic tocRangeParagraphs = tocRangeParagraphsObject;
            dynamic tocRangeFields = tocRangeFieldsObject;
            dynamic tocRangeHyperlinks = tocRangeHyperlinksObject;
            var tocRangeParagraphCount = (int)tocRangeParagraphs.Count;
            var tocRangeFieldCount = (int)tocRangeFields.Count;
            var tocRangeHyperlinkCount = (int)tocRangeHyperlinks.Count;
            var tocRangeText = ((string?)tocRange.Text ?? string.Empty).Trim();

            trace?.Invoke(
                $"TOC_RECONSTRUCTION_COM_ENTRY_EVIDENCE:PARAGRAPHS={tocRangeParagraphCount}:"
                + $"FIELDS={tocRangeFieldCount}:HYPERLINKS={tocRangeHyperlinkCount}:TEXT_CHARS={tocRangeText.Length}");

            if (tocRangeParagraphCount < 2
                || tocRangeText.Length == 0
                || (tocRangeFieldCount <= 1 && tocRangeHyperlinkCount <= 0))
                throw new InvalidOperationException(
                    "Word created a TableOfContents object but produced no decision-ready TOC entry results. "
                    + $"Paragraphs={tocRangeParagraphCount};Fields={tocRangeFieldCount};"
                    + $"Hyperlinks={tocRangeHyperlinkCount};TextChars={tocRangeText.Length}.");

            trace?.Invoke("TOC_RECONSTRUCTION_COM_ENTRY_GATE:PASS");

            if ((int)tables.Count != beforeWord.Counts.Tables
                || (int)sections.Count != beforeWord.Counts.Sections)
                throw new InvalidOperationException(
                    "TOC reconstruction changed protected table/section counts.");

            document.Save();

            warnings.Add(
                $"فهرست مطالب حذف‌شده از روی تیترهای معنایی سند روی کپی کنترل‌شده بازسازی شد؛ "
                + $"ورودی‌های متنی قدیمی={staleEntryIndexes.Count}.");
            trace?.Invoke(
                $"TOC_RECONSTRUCTION_WORD:PASS:TOCS=1:REMOVED_STALE={staleEntryIndexes.Count}:"
                + $"CUSTOM_STYLE_FIELD_ROLE_ALIAS_FAMILY="
                + $"CHAPTER[{string.Join("|", membership.ChapterStyles)}]:1;"
                + $"HEADING1[{string.Join("|", membership.Heading1Styles)}]:2");
        }
        finally
        {
            ReleaseCom(ref tocRangeHyperlinksObject);
            ReleaseCom(ref tocRangeFieldsObject);
            ReleaseCom(ref tocRangeParagraphsObject);
            ReleaseCom(ref tocRangeObject);
            ReleaseCom(ref tocFieldCodeRangeObject);
            ReleaseCom(ref tocFieldObject);
            ReleaseCom(ref heading1WordStyleObject);
            ReleaseCom(ref chapterWordStyleObject);
            ReleaseCom(ref documentStylesObject);
            ReleaseCom(ref tocObject);
            ReleaseCom(ref tocCollectionObject);
            ReleaseCom(ref insertionRangeObject);
            ReleaseCom(ref titleRangeObject);
            ReleaseCom(ref titleParagraphObject);
            ReleaseCom(ref paragraphsObject);
            ReleaseCom(ref sectionsObject);
            ReleaseCom(ref tablesObject);

            if (documentObject is not null)
            {
                try
                {
                    dynamic document = documentObject;
                    document.Close(0);
                }
                catch { }
            }
            ReleaseCom(ref documentObject);
            ReleaseCom(ref documentsObject);

            if (appObject is not null)
            {
                try
                {
                    dynamic app = appObject;
                    app.Quit();
                }
                catch { }
            }
            ReleaseCom(ref appObject);
            StabilizeReleasedCom();
        }

        if (!WaitForNoNewWinWord(baselinePids, 15000, out var remaining))
            throw new InvalidOperationException(
                "TOC reconstruction left a WINWORD orphan. REMAINING="
                + string.Join(",", remaining)
                + ". WordAgent will not kill processes.");

        trace?.Invoke("TOC_RECONSTRUCTION_WORD:ORPHAN_GUARD:PASS:REMAINING=0");
    }

    public ControlledMutationTransaction PrepareAndMutate(
        string inputPath,
        AgentProfile profile,
        string? outputDirectory = null,
        bool visible = false,
        Func<string, WordAuditSnapshot>? wordInspector = null,
        Func<string, AgentProfile, ProfileComplianceSnapshot>? complianceInspector = null,
        Action<string>? trace = null)
    {
        var input = PathPolicy.RequireSupportedWordFile(inputPath);
        var warnings = new List<string>();
        var inputHash = Sha256(input);

        WordAuditSnapshot InspectWordStage(string path, string stage)
        {
            trace?.Invoke(stage + ":START");
            try
            {
                var snapshot = wordInspector is null
                    ? InspectWordWithRetry(path, visible: false, warnings, stage)
                    : wordInspector(path);
                trace?.Invoke(stage + ":PASS");
                return snapshot;
            }
            catch (Exception ex)
            {
                trace?.Invoke(stage + ":FAIL:" + ex.Message);
                throw;
            }
        }

        ProfileComplianceSnapshot InspectComplianceStage(string path, string stage)
        {
            trace?.Invoke(stage + ":START");
            try
            {
                var snapshot = complianceInspector is null
                    ? InspectComplianceWithRetry(path, profile, visible: false, warnings, stage)
                    : complianceInspector(path, profile);
                trace?.Invoke(stage + ":PASS");
                return snapshot;
            }
            catch (Exception ex)
            {
                trace?.Invoke(stage + ":FAIL:" + ex.Message);
                throw;
            }
        }

        trace?.Invoke("PRECHECK:START");
        var beforeWord = InspectWordStage(input, "PRE_WORD_AUDIT");
        var beforeCompliance = InspectComplianceStage(input, "PRE_PROFILE_AUDIT");
        foreach (var diagnostic in beforeCompliance.Warnings.Where(x =>
                     x.StartsWith("DEGRADED_TOC_SOURCE_HEADING_AUTHORITY:", StringComparison.Ordinal)
                     || x.StartsWith("HEADING_RECOVERY_", StringComparison.Ordinal)))
        {
            trace?.Invoke("TOC_RECONSTRUCTION_HEADING_AUTHORITY:" + diagnostic);
        }
        var styleRepairs = beforeCompliance.Styles
            .Where(x => x.Found && !x.FullyCompliant)
            .OrderBy(x => x.Role, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.StyleName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var actions = BuildActions(profile, beforeWord, beforeCompliance);
        var reconstructMissingToc = ShouldReconstructMissingRequiredToc(
            input,
            profile,
            beforeWord,
            beforeCompliance,
            trace);
        var missingRequiredTocNotApplicableNoSourceHierarchy = false;

        if (reconstructMissingToc)
        {
            // QH01-E232:
            // The semantic preflight must use the exact same profile-declared hierarchy
            // as the Word reconstruction path. UniversityBook authority is Heading1/Heading2.
            var tocHierarchy = ResolveControlledTocHierarchyRoles(profile);
            var level1Role = beforeCompliance.Roles.FirstOrDefault(x =>
                x.Role.Equals(tocHierarchy.Level1Role, StringComparison.OrdinalIgnoreCase));
            var level2Role = beforeCompliance.Roles.FirstOrDefault(x =>
                x.Role.Equals(tocHierarchy.Level2Role, StringComparison.OrdinalIgnoreCase));
            var level1Total = level1Role?.Total ?? 0;
            var level2Total = level2Role?.Total ?? 0;
            var level1Assign = actions.Count(x =>
                x.Role.Equals(tocHierarchy.Level1Role, StringComparison.OrdinalIgnoreCase)
                && x.Action.Equals("ASSIGN_STYLE_AND_RESET_DIRECT_FORMAT", StringComparison.OrdinalIgnoreCase));
            var level2Assign = actions.Count(x =>
                x.Role.Equals(tocHierarchy.Level2Role, StringComparison.OrdinalIgnoreCase)
                && x.Action.Equals("ASSIGN_STYLE_AND_RESET_DIRECT_FORMAT", StringComparison.OrdinalIgnoreCase));

            trace?.Invoke(
                $"TOC_RECONSTRUCTION_SOURCE_ROLE_CENSUS:PASS:"
                + $"LEVEL1_ROLE={tocHierarchy.Level1Role}:LEVEL1={level1Total}:ASSIGN_LEVEL1={level1Assign}:"
                + $"LEVEL2_ROLE={tocHierarchy.Level2Role}:LEVEL2={level2Total}:ASSIGN_LEVEL2={level2Assign}");

            if (level1Total <= 0 || level2Total <= 0)
            {
                if (level1Total == 0 && level2Total == 0)
                {
                    // R10R45: a styleless/synthetic or otherwise hierarchy-free document has no
                    // decision-ready source authority from which an automatic TOC can be rebuilt.
                    // Treat the missing TOC as document-level N/A, not as permission to invent headings.
                    reconstructMissingToc = false;
                    missingRequiredTocNotApplicableNoSourceHierarchy = true;
                    trace?.Invoke(
                        $"TOC_RECONSTRUCTION_SOURCE_ROLE_CENSUS:NOT_APPLICABLE:NO_SOURCE_HIERARCHY:"
                        + $"LEVEL1_ROLE={tocHierarchy.Level1Role}:LEVEL1=0:"
                        + $"LEVEL2_ROLE={tocHierarchy.Level2Role}:LEVEL2=0");
                }
                else
                {
                    throw new InvalidOperationException(
                        "Semantic TOC reconstruction requires decision-ready profile-declared level-1 and level-2 source roles before mutation. "
                        + $"Level1Role={tocHierarchy.Level1Role};Level1={level1Total};"
                        + $"Level2Role={tocHierarchy.Level2Role};Level2={level2Total}. "
                        + "TOC generation was not attempted.");
                }
            }
        }

        trace?.Invoke($"PRECHECK:PASS:CANDIDATES={actions.Count}:PARAGRAPH_ACTIONS={actions.Count}:STYLE_REPAIRS={styleRepairs.Length}");
        if (reconstructMissingToc)
            trace?.Invoke("TOC_RECONSTRUCTION_SEMANTIC_PREFLIGHT:PASS");

        var output = PathPolicy.CreateSafeOutputPath(input, outputDirectory, "_WordAgent_Controlled");
        File.Copy(input, output, overwrite: false);
        trace?.Invoke("OOXML_MUTATION:START");

        int mutated;
        try
        {
            mutated = ApplyOpenXmlMutation(
                output,
                profile,
                actions,
                styleRepairs,
                reconstructMissingToc,
                missingRequiredTocNotApplicableNoSourceHierarchy,
                warnings,
                trace);
        }
        catch (Exception ex)
        {
            TryDeleteFailedOutput(output, warnings);
            trace?.Invoke("OOXML_MUTATION:FAIL:" + ex.Message);
            throw new InvalidOperationException(
                $"Controlled Copy Fix OpenXML transaction failed. Original file is untouched. {ex.Message}", ex);
        }

        try
        {
            NormalizeGovernanceGanttLeadingManualPageBreak(output, profile, warnings, trace);

            if (reconstructMissingToc)
            {
                var reconstructionBaseline = TocReconstructionBaselinePath(output);
                if (File.Exists(reconstructionBaseline))
                    throw new InvalidOperationException(
                        "TOC reconstruction baseline already exists unexpectedly.");

                File.Copy(output, reconstructionBaseline, overwrite: false);
                trace?.Invoke("TOC_RECONSTRUCTION_BASELINE:CREATED");

                RebuildMissingRequiredTocWithWord(
                    output,
                    profile,
                    beforeCompliance,
                    beforeWord,
                    warnings,
                    trace);
            }
            else
            {
                RefreshTocResultsWithWord(output, profile, beforeWord, warnings, trace);
            }

            if (missingRequiredTocNotApplicableNoSourceHierarchy)
            {
                // R10R45R1: R10R45 intentionally classified a missing required TOC as
                // document-level N/A when BOTH profile-declared source hierarchy levels
                // are absent. In that exact state there is no TOC object to re-normalize.
                // Skipping only this TOC-specific post-Word pass preserves the fail-closed
                // TOC gates for every document where a TOC exists or reconstruction is required.
                trace?.Invoke(
                    "POST_WORD_TOC_OOXML_RENORMALIZATION:SKIP:NOT_APPLICABLE_NO_SOURCE_HIERARCHY");
            }
            else
            {
                RenormalizeTocOpenXmlAfterWordRefresh(output, profile, warnings, trace);
            }
            ReapplyProfileStyleRepairsAfterWordRefresh(output, profile, styleRepairs, warnings, trace);
            if (!reconstructMissingToc && beforeWord.Counts.TocFields > 0)
            {
                ReconcileExistingTocConcreteProfileStyleContractsAfterWordRefresh(
                    output,
                    profile,
                    warnings,
                    trace);
            }
            if (reconstructMissingToc)
            {
                ApplyBookMaterializedTocStyleContractsAfterWordRefresh(output, profile, warnings, trace);
            }
        }
        catch (Exception ex)
        {
            TryDeleteFailedOutput(output, warnings);
            trace?.Invoke("WORD_NATIVE_TOC_REFRESH:FAIL:" + ex.Message);
            throw new InvalidOperationException(
                $"Controlled Copy Fix Word-native TOC refresh failed. Original file is untouched. {ex.Message}", ex);
        }

        PersianDocumentFoundationResult foundation;
        try
        {
            foundation = new PersianDocumentFoundationService().ApplyInPlace(
                output,
                new PersianDocumentFoundationOptions(ProfileId: profile.Id));
            warnings.Add(
                $"Persian Foundation: paragraphs={foundation.PersianParagraphs}; " +
                $"bodyJustified={foundation.BodyParagraphsJustified}; headingsRight={foundation.HeadingParagraphsRightAligned}; " +
                $"toc={foundation.TocParagraphsNormalized}; runs={foundation.PersianRunsNormalized}; styles={foundation.StyleDefinitionsNormalized}.");
            trace?.Invoke(
                $"PERSIAN_FOUNDATION:PASS:PARAGRAPHS={foundation.PersianParagraphs}:BODY_JUSTIFIED={foundation.BodyParagraphsJustified}:" +
                $"HEADINGS_RIGHT={foundation.HeadingParagraphsRightAligned}:TOC={foundation.TocParagraphsNormalized}:" +
                $"RUNS={foundation.PersianRunsNormalized}:STYLES={foundation.StyleDefinitionsNormalized}");
        }
        catch (Exception ex)
        {
            TryDeleteFailedOutput(output, warnings);
            trace?.Invoke("PERSIAN_FOUNDATION:FAIL:" + ex.Message);
            throw new InvalidOperationException(
                $"Controlled Copy Fix Persian Foundation pass failed. Original file is untouched. {ex.Message}", ex);
        }

        var outputHash = Sha256(output);
        if (actions.Count > 0 && string.Equals(inputHash, outputHash, StringComparison.OrdinalIgnoreCase))
        {
            TryDeleteFailedOutput(output, warnings);
            trace?.Invoke("OOXML_MUTATION:FAIL:BYTE_IDENTICAL");
            throw new InvalidOperationException(
                "Controlled Copy Fix persistence gate failed: output SHA-256 is byte-identical to the input even though mutation candidates existed. Output is NOT approved.");
        }

        trace?.Invoke($"OOXML_MUTATION:PASS:MUTATED={mutated}");
        return new ControlledMutationTransaction(
            input,
            output,
            profile.Id,
            inputHash,
            outputHash,
            actions.Count,
            mutated,
            actions,
            beforeWord,
            beforeCompliance,
            missingRequiredTocNotApplicableNoSourceHierarchy,
            warnings);
    }

    public ControlledCopyFixResult VerifyTransaction(
        ControlledMutationTransaction transaction,
        AgentProfile profile,
        bool visible = false,
        Func<string, WordAuditSnapshot>? wordInspector = null,
        Func<string, AgentProfile, ProfileComplianceSnapshot>? complianceInspector = null,
        Action<string>? trace = null)
    {
        var warnings = transaction.Warnings.ToList();

        WordAuditSnapshot InspectWordStage(string path, string stage)
        {
            trace?.Invoke(stage + ":START");
            try
            {
                var snapshot = wordInspector is null
                    ? InspectWordWithRetry(path, visible: false, warnings, stage)
                    : wordInspector(path);
                trace?.Invoke(stage + ":PASS");
                return snapshot;
            }
            catch (Exception ex)
            {
                trace?.Invoke(stage + ":FAIL:" + ex.Message);
                throw;
            }
        }

        ProfileComplianceSnapshot InspectComplianceStage(string path, string stage)
        {
            trace?.Invoke(stage + ":START");
            try
            {
                var snapshot = complianceInspector is null
                    ? InspectComplianceWithRetry(path, profile, visible: false, warnings, stage)
                    : complianceInspector(path, profile);
                trace?.Invoke(stage + ":PASS");
                return snapshot;
            }
            catch (Exception ex)
            {
                trace?.Invoke(stage + ":FAIL:" + ex.Message);
                throw;
            }
        }

        try
        {
            trace?.Invoke("POST_WORD_VERIFICATION:START");
            var afterWord = InspectWordStage(transaction.OutputFile, "POST_WORD_AUDIT");
            var afterCompliance = InspectComplianceStage(transaction.OutputFile, "POST_PROFILE_AUDIT");
            trace?.Invoke("POST_WORD_VERIFICATION:PASS");

            var reconstructionBaseline = TocReconstructionBaselinePath(transaction.OutputFile);
            var inventoryBeforePath = File.Exists(reconstructionBaseline)
                ? reconstructionBaseline
                : transaction.InputFile;

            var inventoryIssues = CompareInventory(
                transaction.BeforeWord,
                afterWord,
                inventoryBeforePath,
                transaction.OutputFile,
                transaction.ProfileId,
                warnings,
                trace);
            var wordInventoryPassed = inventoryIssues.Count == 0;
            var integrityIssues = new List<string>(inventoryIssues);

            var remaining = afterCompliance.Roles.Sum(x => x.Total - x.FullyCompliant);
            if (afterCompliance.UnclassifiedParagraphs != 0)
                integrityIssues.Add($"Post-fix semantic classification regressed: Unclassified={afterCompliance.UnclassifiedParagraphs}");
            if (remaining != 0)
                integrityIssues.Add($"Post-fix profile compliance still has {remaining} non-compliant paragraph(s).");
            if (afterCompliance.Styles.Any(x => !x.FullyCompliant))
                integrityIssues.Add("Post-fix profile style-definition gate is not fully compliant.");

            return new ControlledCopyFixResult(
                transaction.InputFile,
                transaction.OutputFile,
                transaction.ProfileId,
                transaction.InputSha256,
                transaction.OutputSha256,
                transaction.CandidateParagraphs,
                transaction.MutatedParagraphs,
                transaction.Actions,
                transaction.BeforeWord,
                afterWord,
                transaction.BeforeCompliance,
                afterCompliance,
                wordInventoryPassed,
                integrityIssues,
                warnings);
        }
        catch
        {
            TryDeleteFailedOutput(transaction.OutputFile, warnings);
            trace?.Invoke("POST_WORD_VERIFICATION:FAILED_OUTPUT_DISCARD_ATTEMPTED");
            throw;
        }
    }

    private static string CanonicalizeControlledRoleAliasFamily(string role)
    {
        if (role.Equals("TocLegacy1", StringComparison.OrdinalIgnoreCase))
            return "Toc1";
        if (role.Equals("TocLegacy2", StringComparison.OrdinalIgnoreCase))
            return "Toc2";
        return role;
    }

    // QH01-E229:
    // Resolve the TOC hierarchy from the profile's declared role aliases instead of
    // assuming every profile exposes a literal Chapter role. UniversityBook may
    // declare Heading1/Heading2 as its level-1/level-2 hierarchy.
    private static (string Level1Role, string Level2Role) ResolveControlledTocHierarchyRoles(
        AgentProfile profile)
    {
        // QH01-E231:
        // Match the runner's case-insensitive role-key authority. JSON property casing
        // must not make Runner and ControlledCopyFixService resolve different hierarchies.
        static bool HasAliases(AgentProfile p, string role)
            => p.RoleAliases.Any(pair =>
                string.Equals(pair.Key, role, StringComparison.OrdinalIgnoreCase)
                && (pair.Value ?? Array.Empty<string>()).Any(x => !string.IsNullOrWhiteSpace(x)));

        if (HasAliases(profile, "Chapter") && HasAliases(profile, "Heading1"))
            return ("Chapter", "Heading1");

        if (HasAliases(profile, "Heading1") && HasAliases(profile, "Heading2"))
            return ("Heading1", "Heading2");

        throw new InvalidOperationException(
            "TOC hierarchy cannot be resolved from profile roleAliases. Available roles="
            + string.Join(",", profile.RoleAliases.Keys.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)));
    }

    // QH01-E289:
    // Style-definition ownership must come from the profile's explicit roleAliases only.
    // Semantic selectors (for example CoverSubtitle.StyleNames) classify individual paragraphs
    // and may intentionally reuse a generic/shared style. Treating those semantic selectors as
    // document-global style ownership can create false cross-role conflicts and can mutate a
    // shared style against the wrong contract. Existing-TOC full style closure therefore uses
    // explicit aliases only; semantic-role paragraph materialization remains handled by the
    // already-guarded semantic actions (QH01-E219/QH01-E220). Canonical TocLegacy families are
    // still merged with Toc1/Toc2, and conflicting explicit ownership remains fail-closed.
    private static string[] ResolveControlledExplicitRoleAliasesForStyleDefinition(
        AgentProfile profile,
        string role)
    {
        var canonical = CanonicalizeControlledRoleAliasFamily(role);
        return profile.RoleAliases
            .Where(pair => CanonicalizeControlledRoleAliasFamily(pair.Key)
                .Equals(canonical, StringComparison.OrdinalIgnoreCase))
            .SelectMany(pair => pair.Value ?? Array.Empty<string>())
            .Where(alias => !string.IsNullOrWhiteSpace(alias))
            .Select(alias => alias.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string[] ResolveControlledRoleAliases(AgentProfile profile, string role)
    {
        var canonical = CanonicalizeControlledRoleAliasFamily(role);
        var aliases = profile.RoleAliases
            .Where(pair => CanonicalizeControlledRoleAliasFamily(pair.Key)
                .Equals(canonical, StringComparison.OrdinalIgnoreCase))
            .SelectMany(pair => pair.Value ?? Array.Empty<string>())
            .Where(alias => !string.IsNullOrWhiteSpace(alias))
            .Select(alias => alias.Trim())
            .ToList();

        var coverSubtitleRole = profile.SemanticPolicy.CoverSubtitle.Role;
        if (!string.IsNullOrWhiteSpace(coverSubtitleRole)
            && CanonicalizeControlledRoleAliasFamily(coverSubtitleRole)
                .Equals(canonical, StringComparison.OrdinalIgnoreCase))
        {
            aliases.AddRange(
                (profile.SemanticPolicy.CoverSubtitle.StyleNames ?? Array.Empty<string>())
                .Where(alias => !string.IsNullOrWhiteSpace(alias))
                .Select(alias => alias.Trim()));
        }

        return aliases
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string? ResolveControlledPreferredExistingStyle(
        AgentProfile profile,
        string role,
        IReadOnlyList<string> equivalentAliases,
        ProfileComplianceSnapshot compliance)
    {
        var canonical = CanonicalizeControlledRoleAliasFamily(role);
        var ordered = new List<string>();

        if (profile.RoleAliases.TryGetValue(canonical, out var canonicalAliases))
        {
            ordered.AddRange(
                (canonicalAliases ?? Array.Empty<string>())
                .Where(alias => !string.IsNullOrWhiteSpace(alias))
                .Select(alias => alias.Trim()));
        }

        ordered.AddRange(equivalentAliases.Where(alias => !string.IsNullOrWhiteSpace(alias)));

        foreach (var candidate in ordered.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (compliance.Styles.Any(item =>
                    item.Found
                    && TokensEqual(item.StyleName, candidate)))
                return candidate.Trim();
        }

        return null;
    }

    private static bool HasControlledStyleRepairForSample(
        ProfileComplianceSnapshot compliance,
        string role,
        string? styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
            return false;

        var canonical = CanonicalizeControlledRoleAliasFamily(role);
        return compliance.Styles.Any(item =>
            item.Found
            && !item.FullyCompliant
            && CanonicalizeControlledRoleAliasFamily(item.Role)
                .Equals(canonical, StringComparison.OrdinalIgnoreCase)
            && TokensEqual(item.StyleName, styleName));
    }

    private static bool IsControlledManagedTocRole(string role)
    {
        var canonical = CanonicalizeControlledRoleAliasFamily(role);
        return canonical.Equals("Toc1", StringComparison.OrdinalIgnoreCase)
               || canonical.Equals("Toc2", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasControlledFullyCompliantExactStyle(
        ProfileComplianceSnapshot compliance,
        string role,
        string? styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
            return false;

        var canonical = CanonicalizeControlledRoleAliasFamily(role);
        return compliance.Styles.Any(item =>
            item.Found
            && item.FullyCompliant
            && CanonicalizeControlledRoleAliasFamily(item.Role)
                .Equals(canonical, StringComparison.OrdinalIgnoreCase)
            && TokensEqual(item.StyleName, styleName));
    }

    private static bool HasControlledFullyCompliantStyleByName(
        ProfileComplianceSnapshot compliance,
        string? styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
            return false;

        return compliance.Styles.Any(item =>
            item.Found
            && item.FullyCompliant
            && TokensEqual(item.StyleName, styleName));
    }

    private static List<ControlledFixAction> BuildActions(
        AgentProfile profile,
        WordAuditSnapshot beforeWord,
        ProfileComplianceSnapshot compliance)
    {
        if (compliance.UnclassifiedParagraphs != 0)
            throw new InvalidOperationException($"Controlled fix blocked: Unclassified paragraphs={compliance.UnclassifiedParagraphs}.");

        var styleFailures = compliance.Styles.Count(x => !x.Found && !x.FullyCompliant);
        if (styleFailures != 0)
            throw new InvalidOperationException($"Controlled fix blocked: profile style-definition failures/missing={styleFailures}.");

        if (beforeWord.Counts.Revisions != 0)
            throw new InvalidOperationException($"Controlled fix blocked: tracked revisions={beforeWord.Counts.Revisions}. Resolve or explicitly handle revisions first.");

        var expectedCandidates = compliance.Roles.Sum(x => x.Total - x.FullyCompliant);
        var sampledCandidates = compliance.Roles
            .Where(x => x.Total - x.FullyCompliant > 0)
            .Sum(x => x.Samples.Count);
        if (sampledCandidates != expectedCandidates)
            throw new InvalidOperationException($"Controlled fix blocked: actionable candidate enumeration is incomplete ({sampledCandidates}/{expectedCandidates}). Increase audit sample capacity before mutation.");

        var actions = new List<ControlledFixAction>();
        foreach (var role in compliance.Roles.OrderBy(x => x.Role, StringComparer.OrdinalIgnoreCase))
        {
            var delta = role.Total - role.FullyCompliant;
            if (delta <= 0) continue;

            var aliases = ResolveControlledRoleAliases(profile, role.Role);
            var canonicalStyle = ResolveControlledPreferredExistingStyle(profile, role.Role, aliases, compliance);

            foreach (var sample in role.Samples)
            {
                var styleBelongsToRole = aliases.Any(
                    x => TokensEqual(x, sample.StyleName));
                var styleRepairScheduled = HasControlledStyleRepairForSample(
                    compliance,
                    role.Role,
                    sample.StyleName);
                var exactExistingRoleStyleCompliant =
                    styleBelongsToRole
                    && HasControlledFullyCompliantExactStyle(
                        compliance,
                        role.Role,
                        sample.StyleName);

                // QH01-E218:
                // Role aliases are authoritative for paragraph-role membership. The compliance
                // surface may classify a shared exact style under a sibling canonical role, so
                // role-label equality must not hide an independently proven compliant style definition.
                // This remains safe because styleBelongsToRole is already required by the current
                // profile's alias contract before this style-name-only compliance proof is used.
                var exactStyleDefinitionCompliant =
                    styleBelongsToRole
                    && HasControlledFullyCompliantStyleByName(
                        compliance,
                        sample.StyleName);

                var managedTocResultOverride =
                    IsControlledManagedTocRole(role.Role)
                    && exactExistingRoleStyleCompliant;

                if (!styleBelongsToRole && !string.IsNullOrWhiteSpace(canonicalStyle))
                {
                    actions.Add(new ControlledFixAction(
                        sample.ParagraphIndex,
                        role.Role,
                        "ASSIGN_STYLE_AND_RESET_DIRECT_FORMAT",
                        sample.StyleName,
                        canonicalStyle,
                        sample.TextPreview));
                    continue;
                }

                if (!profile.RoleDefaults.TryGetValue(role.Role, out var defaultsForGuard))
                    throw new InvalidOperationException($"Controlled fix blocked at P{sample.ParagraphIndex}: role defaults missing for {role.Role}.");

                var fontContractAlreadySatisfied =
                    string.IsNullOrWhiteSpace(defaultsForGuard.FontBi)
                    || TokensEqual(sample.ActualFontBi, defaultsForGuard.FontBi);

                var sizeContractAlreadySatisfied =
                    defaultsForGuard.SizeBi is null
                    || (sample.ActualSizeBi is not null
                        && Math.Abs(sample.ActualSizeBi.Value - defaultsForGuard.SizeBi.Value) < 0.01f);

                // QH01-E217:
                // If the paragraph already belongs to the exact role style and that style is independently
                // proven fully compliant, an effective font/size mismatch cannot be repaired safely by
                // paragraph-only materialization. In that narrow case, reset run-level font/size overrides
                // back to the already-validated style identity. OpenXML dry-run, action verification and the
                // post-Word profile audit remain authoritative and fail closed if another inheritance source
                // still violates the role contract.
                var exactStyleRunOverrideRecovery =
                    exactStyleDefinitionCompliant
                    && (!fontContractAlreadySatisfied || !sizeContractAlreadySatisfied);

                // QH01-E219:
                // A semantic CoverSubtitle can be authoritative even when its source paragraph uses a
                // generic style (for example Normal) and none of the profile's CoverSubtitle style names
                // exists in the document. Mutating the shared Normal style would be document-global and
                // unsafe. In this one narrow case, materialize only the bidirectional font/size contract
                // on the already-enumerated semantic paragraph while preserving text and all unrelated
                // character properties. Paragraph bidi/alignment remains explicitly materialized below.
                var semanticCoverSubtitleRunMaterialization =
                    role.Role.Equals("CoverSubtitle", StringComparison.OrdinalIgnoreCase)
                    && !styleBelongsToRole
                    && string.IsNullOrWhiteSpace(canonicalStyle)
                    && !styleRepairScheduled
                    && (!fontContractAlreadySatisfied || !sizeContractAlreadySatisfied);

                // QH01-E220:
                // AR4 advanced past CoverSubtitle and exposed the same styleless semantic-role gap at
                // CoverText. Reuse the already-verified local run-contract action only for CoverText,
                // under the identical no-canonical-style / no-style-repair guard. Do not mutate Normal
                // or weaken the CoverText profile contract (B Titr 20 in the current real report).
                var semanticCoverTextRunMaterialization =
                    role.Role.Equals("CoverText", StringComparison.OrdinalIgnoreCase)
                    && !styleBelongsToRole
                    && string.IsNullOrWhiteSpace(canonicalStyle)
                    && !styleRepairScheduled
                    && (!fontContractAlreadySatisfied || !sizeContractAlreadySatisfied);

                // R10R47R2:
                // The proposal.academic styleless fixture deliberately has no paragraph styles.
                // R47/R47R1 proved that profile alias recovery can classify those paragraphs as Body,
                // but the generic paragraph-only path must still fail because B Nazanin 14 is not
                // inherited from an existing compliant style. Mutating shared Normal would be
                // document-global and unsafe. Reuse the already-verified local run-contract action
                // only for the exact academic styleless Body state:
                //   profile=proposal.academic.fa
                //   role=Body
                //   persisted source style=(unknown)
                //   preferred existing canonical style=Normal
                //   no style repair / no direct override authority
                //   no existing exact-style compliance
                //   missing effective FontBi/SizeBi
                // This preserves paragraph text, fields and style identity while materializing only
                // the profile-governed FontBi/SizeBi plus the existing paragraph RTL/alignment contract.
                var academicStylelessBodyRunMaterialization =
                    profile.Id.Equals("proposal.academic.fa", StringComparison.OrdinalIgnoreCase)
                    && role.Role.Equals("Body", StringComparison.OrdinalIgnoreCase)
                    && styleBelongsToRole
                    && TokensEqual(sample.StyleName, "(unknown)")
                    && TokensEqual(canonicalStyle, "Normal")
                    && !exactExistingRoleStyleCompliant
                    && !exactStyleDefinitionCompliant
                    && !styleRepairScheduled
                    && !sample.DirectOverrideSuspected
                    && string.IsNullOrWhiteSpace(sample.ActualFontBi)
                    && sample.ActualSizeBi is null
                    && (!fontContractAlreadySatisfied || !sizeContractAlreadySatisfied);

                if (semanticCoverSubtitleRunMaterialization
                    || semanticCoverTextRunMaterialization
                    || academicStylelessBodyRunMaterialization)
                {
                    actions.Add(new ControlledFixAction(
                        sample.ParagraphIndex,
                        role.Role,
                        "MATERIALIZE_SEMANTIC_ROLE_RUN_DEFAULTS",
                        sample.StyleName,
                        canonicalStyle,
                        sample.TextPreview));
                    continue;
                }

                // Direct override detection can miss effective run overrides in ordinary roles.
                // General exact-style recovery closes that blind spot without weakening style identity.
                if (managedTocResultOverride
                    || styleRepairScheduled
                    || sample.DirectOverrideSuspected
                    || exactStyleRunOverrideRecovery)
                {
                    actions.Add(new ControlledFixAction(
                        sample.ParagraphIndex,
                        role.Role,
                        "RESET_DIRECT_FORMAT_TO_STYLE",
                        sample.StyleName,
                        canonicalStyle,
                        sample.TextPreview));
                    continue;
                }

                // APPLY_ROLE_DEFAULTS_WITH_GUARD changes only paragraph-level bidi/alignment.
                // It remains admissible only when the run-level font/size contract is already satisfied.
                if (!fontContractAlreadySatisfied || !sizeContractAlreadySatisfied)
                {
                    throw new InvalidOperationException(
                        $"Controlled fix blocked at P{sample.ParagraphIndex} ({role.Role}): paragraph-only materialization is unsafe because an existing font/size contract is not already satisfied. "
                        + $"style='{sample.StyleName}'; canonical='{canonicalStyle}'; styleBelongs={styleBelongsToRole}; "
                        + $"exactRoleStyleCompliant={exactExistingRoleStyleCompliant}; exactStyleDefinitionCompliant={exactStyleDefinitionCompliant}; "
                        + $"styleRepairScheduled={styleRepairScheduled}; directOverride={sample.DirectOverrideSuspected}; "
                        + $"actualFont='{sample.ActualFontBi}'; expectedFont='{defaultsForGuard.FontBi}'; "
                        + $"actualSize={sample.ActualSizeBi?.ToString(CultureInfo.InvariantCulture) ?? "<null>"}; "
                        + $"expectedSize={defaultsForGuard.SizeBi?.ToString(CultureInfo.InvariantCulture) ?? "<null>"}.");
                }

                if (string.IsNullOrWhiteSpace(defaultsForGuard.ReadingOrder) &&
                    string.IsNullOrWhiteSpace(defaultsForGuard.Alignment))
                {
                    throw new InvalidOperationException(
                        $"Controlled fix blocked at P{sample.ParagraphIndex} ({role.Role}): no safe paragraph-format contract is available.");
                }

                actions.Add(new ControlledFixAction(
                    sample.ParagraphIndex,
                    role.Role,
                    "APPLY_ROLE_DEFAULTS_WITH_GUARD",
                    sample.StyleName,
                    canonicalStyle,
                    sample.TextPreview));
            }
        }

        if (actions.Count != expectedCandidates)
            throw new InvalidOperationException($"Controlled fix blocked: action count mismatch ({actions.Count}/{expectedCandidates}).");

        return actions;
    }

    private static int ApplyOpenXmlMutation(
        string outputPath,
        AgentProfile profile,
        IReadOnlyList<ControlledFixAction> actions,
        IReadOnlyList<StyleComplianceItem> styleRepairs,
        bool allowMissingAutomaticTocBeforeReconstruction,
        bool allowMissingAutomaticTocNoSourceHierarchy,
        List<string> warnings,
        Action<string>? trace = null)
    {
        using var file = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(file, ZipArchiveMode.Update, leaveOpen: false);

        if (archive.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException("Digitally signed Office package detected; controlled mutation is blocked because it would invalidate the package signature.");

        var styleRepairResult = ApplyProfileStyleRepairs(
            archive,
            profile,
            styleRepairs,
            mode: "MUTATION_REQUIRED",
            trace: trace);
        if (styleRepairResult.Requested != styleRepairs.Count
            || styleRepairResult.Changed != styleRepairs.Count)
            throw new InvalidOperationException(
                $"Profile style repair count mismatch: requested={styleRepairs.Count}, preflight={styleRepairResult.Requested}, changed={styleRepairResult.Changed}.");

        int mutated = 0;

        if (actions.Count > 0)
        {
            var document = LoadXmlEntry(archive, "word/document.xml");
            var styles = LoadXmlEntry(archive, "word/styles.xml");
            var catalog = OpenXmlStyleCatalogLite.Build(styles);
            var paragraphs = document.Descendants(W + "p").ToList();
            var wordParagraphIndexMap = BuildWordParagraphIndexMap(document);
            var used = new HashSet<XElement>();

            var missingWordIndexes = actions
                .Where(action => !wordParagraphIndexMap.ContainsKey(action.ParagraphIndex))
                .Select(action => action.ParagraphIndex)
                .Distinct()
                .OrderBy(index => index)
                .ToArray();

            if (missingWordIndexes.Length != 0)
            {
                throw new InvalidOperationException(
                    "Word/OpenXML paragraph-index locator could not map action index(es): "
                    + string.Join(",", missingWordIndexes.Take(40))
                    + (missingWordIndexes.Length > 40 ? ",..." : string.Empty));
            }

            trace?.Invoke(
                $"OOXML_ACTION_LOCATOR:PASS:MODE=WORD_PARAGRAPH_INDEX_STRUCTURAL:ACTIONS={actions.Count}:MAPPED={actions.Count}");

            // QH01-E213 safety invariant: controlled formatting must never change visible text
            // or field instructions, including the TOC field code/result text.
            var originalTextNodes = document.Descendants(W + "t")
                .Select(x => (string?)x ?? string.Empty)
                .ToArray();
            var originalFieldInstructions = document.Descendants(W + "instrText")
                .Select(x => (string?)x ?? string.Empty)
                .ToArray();
            var originalFieldCharTypes = document.Descendants(W + "fldChar")
                .Select(x => ((string?)x.Attribute(W + "fldCharType") ?? string.Empty).Trim())
                .ToArray();
            var originalSimpleFieldInstructions = document.Descendants(W + "fldSimple")
                .Select(x => (string?)x.Attribute(W + "instr") ?? string.Empty)
                .ToArray();

            // QH01-E214 / loop-breaker:
            // Resolve and guard every action before the first mutation. This prevents serial
            // "fix the first guard, rerun, discover the next guard" behavior. All structural
            // resolver/guard failures are aggregated in one preflight result.
            var resolvedActions = new List<(ControlledFixAction Action, XElement Paragraph, RoleDefaults Expected)>();
            var preflightErrors = new List<string>();

            foreach (var action in actions.OrderBy(x => x.ParagraphIndex))
            {
                try
                {
                    if (!profile.RoleDefaults.TryGetValue(action.Role, out var expected))
                        throw new InvalidOperationException($"no role contract for {action.Role}");

                    var paragraph = ResolveParagraph(paragraphs, catalog, action, used, wordParagraphIndexMap);
                    GuardParagraph(paragraph, action);

                    if (action.Role.Equals("FigureTableCaption", StringComparison.OrdinalIgnoreCase)
                        && ParagraphHasFieldMarkup(paragraph))
                    {
                        trace?.Invoke(
                            $"OOXML_ACTION_PREFLIGHT:SAFE_CAPTION_FIELD:P{action.ParagraphIndex}:" +
                            DescribeFieldMarkup(paragraph));
                    }

                    resolvedActions.Add((action, paragraph, expected));
                    used.Add(paragraph);
                }
                catch (Exception ex)
                {
                    preflightErrors.Add($"P{action.ParagraphIndex}/{action.Role}: {ex.Message}");
                }
            }

            if (preflightErrors.Count != 0)
            {
                trace?.Invoke($"OOXML_ACTION_PREFLIGHT:FAIL:ERRORS={preflightErrors.Count}");
                throw new InvalidOperationException(
                    $"OpenXML action preflight failed for {preflightErrors.Count}/{actions.Count} action(s): "
                    + string.Join(" || ", preflightErrors));
            }

            if (resolvedActions.Count != actions.Count)
                throw new InvalidOperationException($"OpenXML action preflight count mismatch: {resolvedActions.Count}/{actions.Count}.");

            trace?.Invoke($"OOXML_ACTION_PREFLIGHT:PASS:ACTIONS={resolvedActions.Count}");

            // QH01-E216 / loop-breaker phase 2:
            // Simulate Apply+Verify for every resolved action on an isolated paragraph clone.
            // Aggregate all action-level failures before mutating the real document tree.
            var dryRunErrors = new List<string>();
            foreach (var item in resolvedActions)
            {
                try
                {
                    var clone = new XElement(item.Paragraph);
                    ApplyOpenXmlAction(clone, catalog, item.Action, item.Expected);
                    VerifyOpenXmlAction(clone, catalog, item.Action, item.Expected);
                }
                catch (Exception ex)
                {
                    dryRunErrors.Add($"P{item.Action.ParagraphIndex}/{item.Action.Role}: {ex.Message}");
                }
            }

            if (dryRunErrors.Count != 0)
            {
                trace?.Invoke($"OOXML_ACTION_DRY_RUN:FAIL:ERRORS={dryRunErrors.Count}");
                throw new InvalidOperationException(
                    $"OpenXML action dry-run failed for {dryRunErrors.Count}/{actions.Count} action(s): "
                    + string.Join(" || ", dryRunErrors));
            }

            trace?.Invoke($"OOXML_ACTION_DRY_RUN:PASS:ACTIONS={resolvedActions.Count}");

            foreach (var item in resolvedActions)
            {
                ApplyOpenXmlAction(item.Paragraph, catalog, item.Action, item.Expected);
                VerifyOpenXmlAction(item.Paragraph, catalog, item.Action, item.Expected);
                mutated++;
            }

            if (mutated != actions.Count)
                throw new InvalidOperationException($"OpenXML mutation count mismatch: {mutated}/{actions.Count}.");

            var finalTextNodes = document.Descendants(W + "t")
                .Select(x => (string?)x ?? string.Empty)
                .ToArray();
            var finalFieldInstructions = document.Descendants(W + "instrText")
                .Select(x => (string?)x ?? string.Empty)
                .ToArray();
            var finalFieldCharTypes = document.Descendants(W + "fldChar")
                .Select(x => ((string?)x.Attribute(W + "fldCharType") ?? string.Empty).Trim())
                .ToArray();
            var finalSimpleFieldInstructions = document.Descendants(W + "fldSimple")
                .Select(x => (string?)x.Attribute(W + "instr") ?? string.Empty)
                .ToArray();

            if (!originalTextNodes.SequenceEqual(finalTextNodes, StringComparer.Ordinal))
                throw new InvalidOperationException("Controlled formatting invariant failed: visible document text changed during OpenXML mutation.");
            if (!originalFieldInstructions.SequenceEqual(finalFieldInstructions, StringComparer.Ordinal))
                throw new InvalidOperationException("Controlled formatting invariant failed: field instructions changed during OpenXML mutation.");
            if (!originalFieldCharTypes.SequenceEqual(finalFieldCharTypes, StringComparer.Ordinal))
                throw new InvalidOperationException("Controlled formatting invariant failed: field-char structure changed during OpenXML mutation.");
            if (!originalSimpleFieldInstructions.SequenceEqual(finalSimpleFieldInstructions, StringComparer.Ordinal))
                throw new InvalidOperationException("Controlled formatting invariant failed: simple-field instructions changed during OpenXML mutation.");

            ReplaceXmlEntry(archive, "word/document.xml", document);
            warnings.Add($"موتور تغییر: ویرایش مستقیم OpenXML؛ تعداد نوشتن با Word COM برابر صفر؛ پاراگراف‌های اصلاح‌شده={mutated}.");
            warnings.Add("word/document.xml برای نامزدهای تأییدشده پاراگراف بازنویسی شد؛ Styleها، دستورهای Field، محتوای فهرست مطالب، Relationshipها و رسانه بازتولید نشدند.");
        }
        else
        {
            warnings.Add("موتور OpenXML: هیچ نامزد قالب‌بندی پاراگرافی نیازمند تغییر نبود.");
        }

        var visualFormatting = ApplyVisualFormattingContract(archive, profile, trace);
        warnings.Add(
            $"قالب‌بندی بصری هدفمند: جدول‌ها={visualFormatting.Tables}; پاراگراف‌های جدول={visualFormatting.TableParagraphs}; " +
            $"سربرگ جدول وسط‌چین={visualFormatting.HeaderCentered}; بدنه وسط‌چین={visualFormatting.BodyCentered}; بدنه راست‌چین={visualFormatting.BodyRightAligned}; " +
            $"بخش‌های Header/Footer={visualFormatting.HeaderFooterParts}; پاراگراف‌های Header/Footer={visualFormatting.HeaderFooterParagraphs}; فونت Header/Footer={visualFormatting.HeaderFooterFont}.");

        var layoutStability = ApplyLayoutStabilityAndTocContract(
            archive,
            profile,
            allowMissingAutomaticTocBeforeReconstruction,
            allowMissingAutomaticTocNoSourceHierarchy,
            trace);
        warnings.Add(
            $"پایداری صفحه و فهرست: جدول‌های Fit={layoutStability.WidthFittedTables}/{layoutStability.Tables}; " +
            $"ردیف بدون شکست={layoutStability.RowsNoSplit}/{layoutStability.Rows}; تکرار سربرگ={layoutStability.HeaderRowsRepeat}; " +
            $"زنجیره نگه‌داری جدول={layoutStability.TableKeepChains}; سرستون کوتاه بدون شکست واژه={layoutStability.ShortHeaderNoWrap}; " +
            $"فهرست RTL راست‌چین فیزیکی={layoutStability.TocPhysicalRight}/{layoutStability.TocParagraphs}.");

        if (archive.GetEntry("word/settings.xml") is not null)
        {
            var settings = LoadXmlEntry(archive, "word/settings.xml");
            if (DisableAutomaticFieldUpdateOnOpen(settings))
            {
                ReplaceXmlEntry(archive, "word/settings.xml", settings);
                warnings.Add("به‌روزرسانی خودکار Field هنگام بازشدن سند غیرفعال شد؛ Fieldها سالم می‌مانند و فقط با درخواست صریح کاربر به‌روز می‌شوند.");
            }
        }

        return mutated;
    }


    private sealed record ProfileStyleRepairResult(int Requested, int Changed);

    private static ProfileStyleRepairResult ApplyProfileStyleRepairs(
        ZipArchive archive,
        AgentProfile profile,
        IReadOnlyList<StyleComplianceItem> requestedRepairs,
        string mode,
        Action<string>? trace)
    {
        if (requestedRepairs.Count == 0)
        {
            trace?.Invoke($"PROFILE_STYLE_REPAIR_PREFLIGHT:PASS:COUNT=0");
            trace?.Invoke($"PROFILE_STYLE_REPAIR:PASS:REQUESTED=0:CHANGED=0:MODE={mode}");
            return new ProfileStyleRepairResult(0, 0);
        }

        var styles = LoadXmlEntry(archive, "word/styles.xml");
        var errors = new List<string>();
        var resolved = new List<(StyleComplianceItem Item, XElement Style, RoleDefaults Expected)>();

        foreach (var item in requestedRepairs)
        {
            if (!profile.RoleDefaults.TryGetValue(item.Role, out var expected)
                || !HasConcreteStyleContract(expected))
            {
                errors.Add($"{item.Role}/{item.StyleName}: concrete role defaults missing.");
                continue;
            }

            var matches = styles
                .Descendants(W + "style")
                .Where(style => string.Equals(
                    ((string?)style.Element(W + "name")?.Attribute(W + "val"))?.Trim(),
                    item.StyleName.Trim(),
                    StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (matches.Length != 1)
            {
                errors.Add($"{item.Role}/{item.StyleName}: expected exactly one style match, found {matches.Length}.");
                continue;
            }

            resolved.Add((item, matches[0], expected));
        }

        if (errors.Count != 0)
        {
            trace?.Invoke($"PROFILE_STYLE_REPAIR_PREFLIGHT:FAIL:ERRORS={errors.Count}");
            throw new InvalidOperationException(
                "Profile style repair preflight failed: " + string.Join(" | ", errors));
        }

        trace?.Invoke($"PROFILE_STYLE_REPAIR_PREFLIGHT:PASS:COUNT={resolved.Count}");

        var changed = 0;
        foreach (var target in resolved)
        {
            if (ApplyRoleDefaultsToStyle(target.Style, target.Expected))
                changed++;
        }

        if (changed > 0)
            ReplaceXmlEntry(archive, "word/styles.xml", styles);

        trace?.Invoke(
            $"PROFILE_STYLE_REPAIR:PASS:REQUESTED={resolved.Count}:CHANGED={changed}:MODE={mode}");

        return new ProfileStyleRepairResult(resolved.Count, changed);
    }

    private static bool ApplyRoleDefaultsToStyle(XElement style, RoleDefaults expected)
    {
        var changed = false;

        if (!string.IsNullOrWhiteSpace(expected.FontBi)
            || expected.SizeBi is not null)
        {
            var rPr = style.Element(W + "rPr");
            if (rPr is null)
            {
                rPr = new XElement(W + "rPr");
                style.Add(rPr);
                changed = true;
            }

            if (!string.IsNullOrWhiteSpace(expected.FontBi))
            {
                var rFonts = rPr.Element(W + "rFonts");
                if (rFonts is null)
                {
                    rFonts = new XElement(W + "rFonts");
                    InsertRunPropertyInSchemaOrder(rPr, rFonts);
                    changed = true;
                }

                // QH01-E222:
                // QH01-E224: schema-correct complex-script theme attribute recovery.
                // The OOXML attribute is exactly w:cstheme (lower-case t), not w:csTheme.
                // A concrete profile FontBi is authoritative for complex-script text. Word uses
                // cstheme when it is defined, so retaining a theme token beside w:cs can cause
                // the post-Word effective font to resolve to +Headings CS instead of the
                // requested concrete font. Remove only the competing complex-script theme
                // attribute; preserve ASCII/HighAnsi/EastAsia theme/font dimensions.
                var current = (string?)rFonts.Attribute(W + "cs");
                if (!string.Equals(current?.Trim(), expected.FontBi.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    rFonts.SetAttributeValue(W + "cs", expected.FontBi.Trim());
                    changed = true;
                }

                var csTheme = rFonts.Attribute(W + "cstheme");
                if (csTheme is not null)
                {
                    csTheme.Remove();
                    changed = true;
                }
            }

            if (expected.SizeBi is not null)
            {
                var halfPoints = (int)Math.Round(
                    expected.SizeBi.Value * 2f,
                    MidpointRounding.AwayFromZero);
                var szCs = rPr.Element(W + "szCs");
                if (szCs is null)
                {
                    szCs = new XElement(W + "szCs");
                    InsertRunPropertyInSchemaOrder(rPr, szCs);
                    changed = true;
                }

                var expectedToken = halfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture);
                var current = (string?)szCs.Attribute(W + "val");
                if (!string.Equals(current?.Trim(), expectedToken, StringComparison.Ordinal))
                {
                    szCs.SetAttributeValue(W + "val", expectedToken);
                    changed = true;
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(expected.ReadingOrder)
            || !string.IsNullOrWhiteSpace(expected.Alignment))
        {
            var pPr = style.Element(W + "pPr");
            if (pPr is null)
            {
                pPr = new XElement(W + "pPr");
                var existingRPr = style.Element(W + "rPr");
                if (existingRPr is null)
                    style.Add(pPr);
                else
                    existingRPr.AddBeforeSelf(pPr);
                changed = true;
            }

            if (!string.IsNullOrWhiteSpace(expected.ReadingOrder))
            {
                var bidi = pPr.Element(W + "bidi");
                var shouldRtl = expected.ReadingOrder.Equals("RTL", StringComparison.OrdinalIgnoreCase);
                if (bidi is null)
                {
                    bidi = new XElement(W + "bidi");
                    InsertParagraphPropertyInSchemaOrder(pPr, bidi);
                    changed = true;
                }

                var current = ((string?)bidi.Attribute(W + "val"))?.Trim();
                var currentOn = string.IsNullOrWhiteSpace(current)
                    || !(current.Equals("0", StringComparison.OrdinalIgnoreCase)
                         || current.Equals("false", StringComparison.OrdinalIgnoreCase)
                         || current.Equals("off", StringComparison.OrdinalIgnoreCase)
                         || current.Equals("no", StringComparison.OrdinalIgnoreCase));

                if (currentOn != shouldRtl)
                {
                    bidi.SetAttributeValue(W + "val", shouldRtl ? "1" : "0");
                    changed = true;
                }
            }

            if (!string.IsNullOrWhiteSpace(expected.Alignment))
            {
                var expectedJc = expected.Alignment.Trim().ToUpperInvariant() switch
                {
                    "RIGHT" => "right",
                    "LEFT" => "left",
                    "CENTER" => "center",
                    "JUSTIFY" => "both",
                    "DISTRIBUTE" => "distribute",
                    _ => throw new InvalidOperationException(
                        $"Unsupported profile style alignment token: {expected.Alignment}")
                };

                var jc = pPr.Element(W + "jc");
                if (jc is null)
                {
                    jc = new XElement(W + "jc");
                    InsertParagraphPropertyInSchemaOrder(pPr, jc);
                    changed = true;
                }

                var current = (string?)jc.Attribute(W + "val");
                if (!string.Equals(current?.Trim(), expectedJc, StringComparison.OrdinalIgnoreCase))
                {
                    jc.SetAttributeValue(W + "val", expectedJc);
                    changed = true;
                }
            }
        }

        return changed;
    }

    private static bool HasConcreteStyleContract(RoleDefaults expected)
        => !string.IsNullOrWhiteSpace(expected.FontBi)
           || expected.SizeBi is not null
           || !string.IsNullOrWhiteSpace(expected.ReadingOrder)
           || !string.IsNullOrWhiteSpace(expected.Alignment);

    private static void ReapplyProfileStyleRepairsAfterWordRefresh(
        string outputPath,
        AgentProfile profile,
        IReadOnlyList<StyleComplianceItem> styleRepairs,
        List<string> warnings,
        Action<string>? trace)
    {
        if (styleRepairs.Count == 0)
        {
            trace?.Invoke("POST_WORD_PROFILE_STYLE_RENORMALIZATION:SKIP:REPAIRS=0");
            return;
        }

        trace?.Invoke($"POST_WORD_PROFILE_STYLE_RENORMALIZATION:START:REPAIRS={styleRepairs.Count}");
        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var result = ApplyProfileStyleRepairs(
            archive,
            profile,
            styleRepairs,
            mode: "IDEMPOTENT_VERIFY",
            trace: trace);

        if (result.Requested != styleRepairs.Count)
            throw new InvalidOperationException(
                $"Post-Word profile style repair count mismatch: requested={styleRepairs.Count}, resolved={result.Requested}.");

        warnings.Add(
            $"پس از مرحله Word، قرارداد Styleهای قابل‌اصلاح دوباره روی کپی کنترل‌شده اعمال شد؛ تعداد={styleRepairs.Count}.");

        trace?.Invoke(
            $"POST_WORD_PROFILE_STYLE_RENORMALIZATION:PASS:REPAIRS={result.Requested}:CHANGED={result.Changed}:STATE=REAPPLIED");
    }

    // QH01-E288 + QH01-E289:
    // On an existing-TOC controlled-fix path, Word's native TOC refresh can rewrite
    // style definitions that were fully compliant before refresh. Reapplying only the
    // pre-refresh non-compliant subset is therefore insufficient: a previously compliant
    // style can become non-compliant after Word. For documents that already had a TOC,
    // reconcile every PRESENT style that is explicitly governed by a concrete profile
    // role contract. Missing/unused aliases are not materialized, unrelated styles are
    // untouched, and ambiguous cross-role ownership fails closed.
    private static void ReconcileExistingTocConcreteProfileStyleContractsAfterWordRefresh(
        string outputPath,
        AgentProfile profile,
        List<string> warnings,
        Action<string>? trace)
    {
        trace?.Invoke("POST_WORD_PROFILE_STYLE_FULL_CLOSURE:START");
        trace?.Invoke("POST_WORD_PROFILE_STYLE_FULL_CLOSURE:OWNERSHIP_POLICY:EXPLICIT_ROLE_ALIASES_ONLY");

        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var styles = LoadXmlEntry(archive, "word/styles.xml");
        var styleElements = styles
            .Descendants(W + "style")
            .Select(style => new
            {
                Style = style,
                Name = ((string?)style.Element(W + "name")?.Attribute(W + "val"))?.Trim()
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Name))
            .ToArray();

        var claimed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var concreteRoles = 0;
        var resolvedStyles = 0;
        var changedStyles = 0;

        foreach (var pair in profile.RoleDefaults.OrderBy(
                     x => x.Key,
                     StringComparer.OrdinalIgnoreCase))
        {
            var role = pair.Key;
            var expected = pair.Value;
            if (!HasConcreteStyleContract(expected))
                continue;

            var aliases = ResolveControlledExplicitRoleAliasesForStyleDefinition(profile, role);
            if (aliases.Length == 0)
                continue;

            concreteRoles++;
            foreach (var target in styleElements.Where(x =>
                         aliases.Any(alias => TokensEqual(alias, x.Name!))))
            {
                var styleName = target.Name!;
                var canonicalRole = CanonicalizeControlledRoleAliasFamily(role);

                if (claimed.TryGetValue(styleName, out var priorRole))
                {
                    var priorCanonical = CanonicalizeControlledRoleAliasFamily(priorRole);
                    if (!priorCanonical.Equals(canonicalRole, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException(
                            "Existing-TOC profile style closure found ambiguous cross-role style ownership: "
                            + $"Style='{styleName}';Role1='{priorRole}';Role2='{role}'.");
                    }

                    continue;
                }

                claimed.Add(styleName, role);
                resolvedStyles++;
                if (ApplyRoleDefaultsToStyle(target.Style, expected))
                    changedStyles++;
            }
        }

        if (changedStyles > 0)
            ReplaceXmlEntry(archive, "word/styles.xml", styles);

        warnings.Add(
            "پس از نوسازی TOC موجود توسط Word، قرارداد تمام Styleهای موجود با defaults صریح پروفایل "
            + $"دوباره اعمال شد؛ roles={concreteRoles}; styles={resolvedStyles}; changed={changedStyles}.");
        trace?.Invoke(
            $"POST_WORD_PROFILE_STYLE_FULL_CLOSURE:PASS:"
            + $"CONCRETE_ROLES={concreteRoles}:RESOLVED_STYLES={resolvedStyles}:CHANGED_STYLES={changedStyles}");
    }

    // QH01-E278:
    // Word creates TOC1/TOC2 styles only when the missing automatic TOC is materialized.
    // The pre-Word compliance snapshot therefore cannot request repairs for those styles.
    // For the Book missing-TOC branch only, reconcile newly materialized TOC1/TOC2 style
    // definitions against the already-declared profile defaults after Word refresh.
    // This preserves the profile contract (including Toc2 11.5 pt) instead of weakening it
    // to Word's transient native 12 pt materialization. No other profile or non-reconstruction
    // branch is eligible.
    private static void ApplyBookMaterializedTocStyleContractsAfterWordRefresh(
        string outputPath,
        AgentProfile profile,
        List<string> warnings,
        Action<string>? trace)
    {
        if (!string.Equals(profile.Id, "book.university.fa", StringComparison.OrdinalIgnoreCase))
        {
            trace?.Invoke("BOOK_MATERIALIZED_TOC_STYLE_CONTRACT:NOT_APPLICABLE:PROFILE=" + profile.Id);
            return;
        }

        var roles = new[] { "Toc1", "Toc2" };
        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var styles = LoadXmlEntry(archive, "word/styles.xml");
        var changed = 0;
        var resolved = 0;

        foreach (var role in roles)
        {
            if (!profile.RoleDefaults.TryGetValue(role, out var expected)
                || !HasConcreteStyleContract(expected))
                throw new InvalidOperationException($"Book materialized TOC style contract missing concrete defaults for role {role}.");

            var aliases = ResolveControlledRoleAliases(profile, role);
            if (aliases.Length == 0)
                throw new InvalidOperationException($"Book materialized TOC style contract missing aliases for role {role}.");

            var matches = styles
                .Descendants(W + "style")
                .Where(style =>
                {
                    var name = ((string?)style.Element(W + "name")?.Attribute(W + "val"))?.Trim();
                    return !string.IsNullOrWhiteSpace(name)
                           && aliases.Any(alias => TokensEqual(alias, name));
                })
                .ToArray();

            if (matches.Length != 1)
                throw new InvalidOperationException(
                    $"Book materialized TOC style contract expected exactly one style for role {role}; found {matches.Length}; aliases=[{string.Join(",", aliases)}].");

            resolved++;
            if (ApplyRoleDefaultsToStyle(matches[0], expected))
                changed++;
        }

        if (resolved != roles.Length)
            throw new InvalidOperationException(
                $"Book materialized TOC style contract resolution mismatch: resolved={resolved}; expected={roles.Length}.");

        if (changed > 0)
            ReplaceXmlEntry(archive, "word/styles.xml", styles);

        warnings.Add(
            $"پس از ساخت TOC توسط Word، قرارداد Styleهای تازه‌ساخته‌شده Toc1/Toc2 برای پروفایل کتاب دوباره اعمال شد؛ resolved={resolved}; changed={changed}.");
        trace?.Invoke(
            $"BOOK_MATERIALIZED_TOC_STYLE_CONTRACT:PASS:RESOLVED={resolved}:CHANGED={changed}:ROLES=Toc1,Toc2");
    }

    private sealed record VisualFormattingStats(
        int Tables,
        int TableParagraphs,
        int HeaderCentered,
        int BodyCentered,
        int BodyRightAligned,
        int HeaderFooterParts,
        int HeaderFooterParagraphs,
        string HeaderFooterFont);

    private static VisualFormattingStats ApplyVisualFormattingContract(
        ZipArchive archive,
        AgentProfile profile,
        Action<string>? trace)
    {
        const string headerFooterFont = "B Titr";
        var tablePolicy = profile.LayoutPolicy.TableStyle;

        var document = LoadXmlEntry(archive, "word/document.xml");
        var originalDocumentText = document.Descendants(W + "t")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();
        var originalDocumentFields = document.Descendants(W + "instrText")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();

        // R8R26R1R7: schema-only normalization must not depend on profile TableStyle enablement.
        // Existing enabled bidiVisual values (1/true/on) are semantically equivalent to the
        // presence-only CT_OnOffOnly form, so remove only those explicit enabled values.
        // False/unknown explicit values are preserved and left for the fail-closed validator.
        var bidiCanonicalization = CanonicalizeEnabledBidiVisualValues(document);
        trace?.Invoke(
            $"BIDIVISUAL_CANONICALIZATION:PASS:TOTAL={bidiCanonicalization.Total}:" +
            $"EXPLICIT={bidiCanonicalization.Explicit}:NORMALIZED={bidiCanonicalization.Normalized}:" +
            $"UNHANDLED_EXPLICIT={bidiCanonicalization.UnhandledExplicit}");

        var tableInventory = InspectTableVisualFormatting(document, tablePolicy);
        var storyNames = archive.Entries
            .Select(x => x.FullName)
            .Where(IsHeaderFooterStoryPart)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var storyParagraphs = 0;
        foreach (var storyName in storyNames)
        {
            var story = LoadXmlEntry(archive, storyName);
            storyParagraphs += story.Descendants(W + "p")
                .Count(p => NormalizeText(ParagraphText(p)).Length > 0);
        }

        trace?.Invoke(
            $"VISUAL_FORMATTING_PREFLIGHT:PASS:TABLES={tableInventory.Tables}:TABLE_PARAGRAPHS={tableInventory.TableParagraphs}:HF_PARTS={storyNames.Length}:HF_PARAGRAPHS={storyParagraphs}");

        var tableMutation = ApplyTableVisualFormatting(document, tablePolicy);
        var tableVerification = InspectTableVisualFormatting(document, tablePolicy);
        if (!tableVerification.FullyCompliant)
        {
            throw new InvalidOperationException(
                "Table visual formatting verification failed: "
                + string.Join(" | ", tableVerification.Issues.Take(20)));
        }

        var finalDocumentText = document.Descendants(W + "t")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();
        var finalDocumentFields = document.Descendants(W + "instrText")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();

        if (!originalDocumentText.SequenceEqual(finalDocumentText, StringComparer.Ordinal))
            throw new InvalidOperationException("Table visual formatting invariant failed: visible document text changed.");
        if (!originalDocumentFields.SequenceEqual(finalDocumentFields, StringComparer.Ordinal))
            throw new InvalidOperationException("Table visual formatting invariant failed: field instructions changed.");

        if (bidiCanonicalization.Changed || tableMutation.Changed)
            ReplaceXmlEntry(archive, "word/document.xml", document);

        var headerFooterParagraphs = 0;
        foreach (var storyName in storyNames)
        {
            var story = LoadXmlEntry(archive, storyName);
            var originalText = story.Descendants(W + "t")
                .Select(x => (string?)x ?? string.Empty)
                .ToArray();
            var originalFields = story.Descendants(W + "instrText")
                .Select(x => (string?)x ?? string.Empty)
                .ToArray();

            var storyChanged = ApplyHeaderFooterVisualFormatting(story, headerFooterFont);
            var storyVerification = InspectHeaderFooterVisualFormatting(story, headerFooterFont);
            if (!storyVerification.FullyCompliant)
            {
                throw new InvalidOperationException(
                    $"Header/Footer formatting verification failed for {storyName}: "
                    + string.Join(" | ", storyVerification.Issues.Take(20)));
            }

            headerFooterParagraphs += storyVerification.Paragraphs;

            var finalText = story.Descendants(W + "t")
                .Select(x => (string?)x ?? string.Empty)
                .ToArray();
            var finalFields = story.Descendants(W + "instrText")
                .Select(x => (string?)x ?? string.Empty)
                .ToArray();

            if (!originalText.SequenceEqual(finalText, StringComparer.Ordinal))
                throw new InvalidOperationException($"Header/Footer invariant failed: visible text changed in {storyName}.");
            if (!originalFields.SequenceEqual(finalFields, StringComparer.Ordinal))
                throw new InvalidOperationException($"Header/Footer invariant failed: field instructions changed in {storyName}.");

            if (storyChanged)
                ReplaceXmlEntry(archive, storyName, story);
        }

        var stats = new VisualFormattingStats(
            tableVerification.Tables,
            tableVerification.TableParagraphs,
            tableVerification.HeaderCentered,
            tableVerification.BodyCentered,
            tableVerification.BodyRightAligned,
            storyNames.Length,
            headerFooterParagraphs,
            headerFooterFont);

        trace?.Invoke(
            $"VISUAL_FORMATTING_POST:PASS:TABLES={stats.Tables}:TABLE_PARAGRAPHS={stats.TableParagraphs}:HEADER_CENTER={stats.HeaderCentered}:BODY_CENTER={stats.BodyCentered}:BODY_RIGHT={stats.BodyRightAligned}:HF_PARTS={stats.HeaderFooterParts}:HF_PARAGRAPHS={stats.HeaderFooterParagraphs}:HF_FONT=B_TITR");

        return stats;
    }

    private sealed record BidiVisualCanonicalization(
        int Total,
        int Explicit,
        int Normalized,
        int UnhandledExplicit,
        bool Changed);

    private static BidiVisualCanonicalization CanonicalizeEnabledBidiVisualValues(XDocument document)
    {
        var total = 0;
        var explicitCount = 0;
        var normalized = 0;
        var unhandledExplicit = 0;
        var changed = false;

        foreach (var bidiVisual in document.Descendants(W + "bidiVisual"))
        {
            total++;
            var value = bidiVisual.Attribute(W + "val");
            if (value is null)
                continue;

            explicitCount++;
            var lexical = value.Value.Trim();
            var enabled = lexical.Equals("1", StringComparison.OrdinalIgnoreCase)
                || lexical.Equals("true", StringComparison.OrdinalIgnoreCase)
                || lexical.Equals("on", StringComparison.OrdinalIgnoreCase);

            if (enabled)
            {
                value.Remove();
                normalized++;
                changed = true;
            }
            else
            {
                // Do not silently turn an explicit false/unknown value into enabled state.
                // The exhaustive OpenXmlValidator gate remains authoritative for such input.
                unhandledExplicit++;
            }
        }

        return new BidiVisualCanonicalization(
            total, explicitCount, normalized, unhandledExplicit, changed);
    }

    private sealed record TableVisualInspection(
        int Tables,
        int TableParagraphs,
        int HeaderCentered,
        int BodyCentered,
        int BodyRightAligned,
        bool FullyCompliant,
        IReadOnlyList<string> Issues);

    private sealed record TableVisualMutation(bool Changed);

    private static TableVisualMutation ApplyTableVisualFormatting(
        XDocument document,
        ProfileTableStyleLayoutPolicy policy)
    {
        if (!policy.Enabled)
            return new TableVisualMutation(false);

        var changed = false;
        foreach (var table in document.Descendants(W + "tbl"))
        {
            var tblPr = table.Element(W + "tblPr");
            if (tblPr is null)
            {
                tblPr = new XElement(W + "tblPr");
                table.AddFirst(tblPr);
                changed = true;
            }

            if (policy.RtlVisual)
            {
                var bidiVisual = tblPr.Element(W + "bidiVisual");
                if (bidiVisual is null)
                {
                    // w:bidiVisual is CT_OnOffOnly.  The canonical enabled form is the
                    // empty element (<w:bidiVisual/>), not the ST_OnOff lexical value "1".
                    bidiVisual = new XElement(W + "bidiVisual");
                    InsertTablePropertyInSchemaOrder(tblPr, bidiVisual);
                    changed = true;
                }
                else
                {
                    // Policy requires RTL visual order, so normalize any explicit value
                    // (including legacy "1" or "off") to the schema-valid enabled form.
                    var bidiVisualValue = bidiVisual.Attribute(W + "val");
                    if (bidiVisualValue is not null)
                    {
                        bidiVisualValue.Remove();
                        changed = true;
                    }
                }
            }

            if (policy.CenterTable)
            {
                var jc = tblPr.Element(W + "jc");
                if (jc is null)
                {
                    jc = new XElement(W + "jc");
                    InsertTablePropertyInSchemaOrder(tblPr, jc);
                    changed = true;
                }
                var current = ((string?)jc.Attribute(W + "val") ?? string.Empty).Trim();
                if (!current.Equals("center", StringComparison.OrdinalIgnoreCase))
                {
                    jc.SetAttributeValue(W + "val", "center");
                    changed = true;
                }
            }

            var rows = table.Elements(W + "tr").ToList();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                foreach (var cell in rows[rowIndex].Elements(W + "tc"))
                {
                    foreach (var paragraph in DirectCellParagraphs(cell))
                    {
                        var text = ParagraphText(paragraph);
                        var alignment = rowIndex == 0
                            ? policy.HeaderAlignment
                            : DecideTableCellAlignment(text, policy);

                        if (MaterializeTableCellParagraphAlignment(
                                paragraph,
                                alignment,
                                ContainsPersian(text)))
                            changed = true;
                    }
                }
            }
        }

        return new TableVisualMutation(changed);
    }

    private static TableVisualInspection InspectTableVisualFormatting(
        XDocument document,
        ProfileTableStyleLayoutPolicy policy)
    {
        if (!policy.Enabled)
            return new TableVisualInspection(0, 0, 0, 0, 0, true, Array.Empty<string>());

        var issues = new List<string>();
        var tableCount = 0;
        var paragraphCount = 0;
        var headerCentered = 0;
        var bodyCentered = 0;
        var bodyRight = 0;

        foreach (var table in document.Descendants(W + "tbl"))
        {
            tableCount++;
            var tblPr = table.Element(W + "tblPr");

            if (policy.RtlVisual && !IsOnOffEnabled(tblPr?.Element(W + "bidiVisual")))
                issues.Add($"T{tableCount}: table visual direction is not RTL.");

            if (policy.CenterTable)
            {
                var tableJc = ((string?)tblPr?.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();
                if (!tableJc.Equals("center", StringComparison.OrdinalIgnoreCase))
                    issues.Add($"T{tableCount}: table alignment is not CENTER.");
            }

            var rows = table.Elements(W + "tr").ToList();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                foreach (var cell in rows[rowIndex].Elements(W + "tc"))
                {
                    foreach (var paragraph in DirectCellParagraphs(cell))
                    {
                        var text = ParagraphText(paragraph);
                        var normalized = NormalizeText(text);
                        if (normalized.Length == 0)
                            continue;

                        paragraphCount++;
                        var alignment = rowIndex == 0
                            ? policy.HeaderAlignment
                            : DecideTableCellAlignment(text, policy);

                        if (rowIndex == 0)
                            headerCentered++;
                        else if (alignment.Trim().Equals("RIGHT", StringComparison.OrdinalIgnoreCase))
                            bodyRight++;
                        else
                            bodyCentered++;

                        var pPr = paragraph.Element(W + "pPr");
                        var expectedJc = EncodeAlignmentForWordOoxml(
                            alignment,
                            ContainsPersian(text) ? "RTL" : null);
                        var actualJc = ((string?)pPr?.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();

                        if (!actualJc.Equals(expectedJc, StringComparison.OrdinalIgnoreCase))
                            issues.Add($"T{tableCount}/R{rowIndex + 1}: cell paragraph alignment mismatch actual='{actualJc}' expected='{expectedJc}'.");

                        if (ContainsPersian(text) && !IsOnOffEnabled(pPr?.Element(W + "bidi")))
                            issues.Add($"T{tableCount}/R{rowIndex + 1}: Persian cell paragraph is not RTL.");
                    }
                }
            }
        }

        return new TableVisualInspection(
            tableCount,
            paragraphCount,
            headerCentered,
            bodyCentered,
            bodyRight,
            issues.Count == 0,
            issues);
    }

    private sealed record LayoutStabilityStats(
        int Tables,
        int Rows,
        int WidthFittedTables,
        int RowsNoSplit,
        int HeaderRowsRepeat,
        int TableKeepChains,
        int KeepNextParagraphs,
        int ShortHeaderNoWrap,
        int CaptionsKeepNext,
        int TocParagraphs,
        int TocPhysicalRight,
        int TocTextRunsRtl,
        int PageContentWidthTwips);

    private sealed record LayoutStabilityInspection(
        LayoutStabilityStats Stats,
        bool FullyCompliant,
        IReadOnlyList<string> Issues);

    // QH01-E221:
    // Keep the ControlledCopyFixService absence policy identical to the CLI authority.
    // Missing automatic TOC is NOT_APPLICABLE for ordinary profiles, but remains mandatory
    // for governance reconstruction and the academic/book profiles that require structural TOC.
    private static bool IsAutomaticTocNotApplicableWhenAbsent(AgentProfile profile)
    {
        if (profile.Id.Equals("proposal.governance.fa", StringComparison.OrdinalIgnoreCase))
            return false;

        return !profile.Id.Equals("book.university.fa", StringComparison.OrdinalIgnoreCase)
            && !profile.Id.Equals("thesis.dissertation.phd.fa", StringComparison.OrdinalIgnoreCase)
            && !profile.Id.Equals("thesis-dissertation-academic-fa", StringComparison.OrdinalIgnoreCase);
    }

    private static LayoutStabilityStats ApplyLayoutStabilityAndTocContract(
        ZipArchive archive,
        AgentProfile profile,
        bool allowMissingAutomaticTocBeforeReconstruction,
        bool allowMissingAutomaticTocNoSourceHierarchy,
        Action<string>? trace)
    {
        var document = LoadXmlEntry(archive, "word/document.xml");
        var styles = LoadXmlEntry(archive, "word/styles.xml");
        var tocStyleTokens = ResolveAutomaticTocStyleTokens(styles, profile);
        trace?.Invoke(
            $"TOC_STYLE_AUTHORITY:PASS:TOKENS={string.Join(",", tocStyleTokens.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))}");
        var originalText = document.Descendants(W + "t")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();
        var originalFields = document.Descendants(W + "instrText")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();

        var pageContentWidth = ResolveMinimumPageContentWidthTwips(document);
        var pre = InspectLayoutStabilityAndToc(document, pageContentWidth, tocStyleTokens);

        trace?.Invoke(
            $"LAYOUT_STABILITY_PREFLIGHT:PASS:TABLES={pre.Stats.Tables}:ROWS={pre.Stats.Rows}:TOC={pre.Stats.TocParagraphs}:PAGE_CONTENT={pageContentWidth}");

        var tocTabsPre = InspectAutomaticTocTabLayout(document, styles, pageContentWidth, tocStyleTokens);
        var automaticTocStructurallyAbsent =
            pre.Stats.TocParagraphs == 0
            && tocTabsPre.Paragraphs == 0
            && tocTabsPre.Styles == 0
            && tocTabsPre.Issues.Count == 0;
        var deferMissingAutomaticToc =
            allowMissingAutomaticTocBeforeReconstruction
            && automaticTocStructurallyAbsent;
        var missingAutomaticTocNotApplicable =
            !allowMissingAutomaticTocBeforeReconstruction
            && (IsAutomaticTocNotApplicableWhenAbsent(profile) || allowMissingAutomaticTocNoSourceHierarchy)
            && automaticTocStructurallyAbsent;

        if (deferMissingAutomaticToc)
        {
            trace?.Invoke(
                $"TOC_TAB_LAYOUT_PREFLIGHT:DEFERRED_RECONSTRUCTION_REQUIRED:PARAGRAPHS=0:STYLES=0:TAB_POS={tocTabsPre.TabPosition}");
            trace?.Invoke(
                "TOC_PRE_RECONSTRUCTION_GATE:DEFERRED:REASON=SEMANTIC_RECONSTRUCTION_ALREADY_AUTHORIZED:STRICT_POST_WORD_GATE=REQUIRED");
        }
        else if (missingAutomaticTocNotApplicable)
        {
            trace?.Invoke(
                $"TOC_TAB_LAYOUT_PREFLIGHT:NOT_APPLICABLE:PASS:PROFILE={profile.Id}:REASON="
                + $"{(allowMissingAutomaticTocNoSourceHierarchy ? "NO_SOURCE_HIERARCHY" : "PROFILE_POLICY")}:"
                + $"PARAGRAPHS=0:STYLES=0:TAB_POS={tocTabsPre.TabPosition}");
        }
        else
        {
            trace?.Invoke(
                $"TOC_TAB_LAYOUT_PREFLIGHT:REPAIRABLE:PARAGRAPHS={tocTabsPre.Paragraphs}:PARAGRAPH_TABS_OK={tocTabsPre.ParagraphTabsOk}:STYLES={tocTabsPre.Styles}:STYLE_TABS_OK={tocTabsPre.StyleTabsOk}:TAB_POS={tocTabsPre.TabPosition}");
        }

        var changed = false;
        var stylesChanged = false;
        if (ApplyTableFitAndPagination(document, pageContentWidth))
            changed = true;
        if (ApplyAutomaticTocPhysicalRtl(document, tocStyleTokens))
            changed = true;
        if (ApplyAutomaticTocTabLayout(document, styles, pageContentWidth, tocStyleTokens))
        {
            changed = true;
            stylesChanged = true;
        }

        var post = InspectLayoutStabilityAndToc(document, pageContentWidth, tocStyleTokens);
        var tocTabsPost = InspectAutomaticTocTabLayout(document, styles, pageContentWidth, tocStyleTokens);

        var deferredTocTabsAccepted = deferMissingAutomaticToc
            && tocTabsPost.Paragraphs == 0
            && tocTabsPost.Styles == 0
            && tocTabsPost.Issues.Count == 0;
        var notApplicableTocTabsAccepted = missingAutomaticTocNotApplicable
            && tocTabsPost.Paragraphs == 0
            && tocTabsPost.Styles == 0
            && tocTabsPost.Issues.Count == 0;
        if (!tocTabsPost.FullyCompliant && !deferredTocTabsAccepted && !notApplicableTocTabsAccepted)
            throw new InvalidOperationException("Automatic TOC tab-layout verification failed: " + string.Join(" | ", tocTabsPost.Issues.Take(20)));

        var sPost = post.Stats;
        // QH01-E271: table layout obligations are vacuously satisfied when the document
        // contains zero tables. Missing-TOC pre-reconstruction acceptance must not require
        // an unrelated table to exist; every table-specific equality still remains enforced
        // whenever one or more tables are present.
        var missingTocTableLayoutAccepted =
            sPost.WidthFittedTables == sPost.Tables
            && sPost.RowsNoSplit == sPost.Rows
            && sPost.HeaderRowsRepeat == sPost.Tables
            && sPost.TableKeepChains == 0
            && sPost.KeepNextParagraphs == 0;
        var missingTocLayoutBaseAccepted =
            missingTocTableLayoutAccepted
            && sPost.TocParagraphs == 0
            && sPost.TocPhysicalRight == 0
            && sPost.TocTextRunsRtl == 0
            && post.Issues.Count == 0;
        var deferredLayoutAccepted =
            deferMissingAutomaticToc
            && missingTocLayoutBaseAccepted;
        var notApplicableLayoutAccepted =
            missingAutomaticTocNotApplicable
            && missingTocLayoutBaseAccepted;
        if (!post.FullyCompliant && !deferredLayoutAccepted && !notApplicableLayoutAccepted)
        {
            throw new InvalidOperationException(
                "Layout stability / TOC visual verification failed: "
                + string.Join(" | ", post.Issues.Take(30)));
        }

        var finalText = document.Descendants(W + "t")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();
        var finalFields = document.Descendants(W + "instrText")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();

        if (!originalText.SequenceEqual(finalText, StringComparer.Ordinal))
            throw new InvalidOperationException("Layout stability invariant failed: visible document text changed.");
        if (!originalFields.SequenceEqual(finalFields, StringComparer.Ordinal))
            throw new InvalidOperationException("Layout stability invariant failed: field instructions changed.");

        if (changed)
            ReplaceXmlEntry(archive, "word/document.xml", document);
        if (stylesChanged)
            ReplaceXmlEntry(archive, "word/styles.xml", styles);

        if (deferMissingAutomaticToc)
        {
            trace?.Invoke(
                $"TOC_TAB_LAYOUT_POST:DEFERRED_UNTIL_RECONSTRUCTION:PARAGRAPHS=0:STYLES=0:TAB_POS={tocTabsPost.TabPosition}:STRICT_POST_WORD_GATE=REQUIRED");
            trace?.Invoke(
                "TOC_COMPACT_PHYSICAL_RIGHT_POST:DEFERRED_UNTIL_RECONSTRUCTION:STRICT_POST_WORD_GATE=REQUIRED");
        }
        else if (missingAutomaticTocNotApplicable)
        {
            trace?.Invoke(
                $"TOC_TAB_LAYOUT_POST:NOT_APPLICABLE:PASS:PROFILE={profile.Id}:PARAGRAPHS=0:STYLES=0:TAB_POS={tocTabsPost.TabPosition}");
            trace?.Invoke(
                $"TOC_COMPACT_PHYSICAL_RIGHT_POST:NOT_APPLICABLE:PASS:PROFILE={profile.Id}:PARAGRAPHS=0");
        }
        else
        {
            trace?.Invoke(
                $"TOC_TAB_LAYOUT_POST:PASS:PARAGRAPHS={tocTabsPost.Paragraphs}:PARAGRAPH_TABS_OK={tocTabsPost.ParagraphTabsOk}:STYLES={tocTabsPost.Styles}:STYLE_TABS_OK={tocTabsPost.StyleTabsOk}:TAB_POS={tocTabsPost.TabPosition}");

            trace?.Invoke(
                $"TOC_COMPACT_PHYSICAL_RIGHT_POST:PASS:PARAGRAPHS={tocTabsPost.Paragraphs}:RIGHT={tocTabsPost.ParagraphTabsOk}:STYLES={tocTabsPost.Styles}:STYLE_RIGHT={tocTabsPost.StyleTabsOk}:TAB=RIGHT:SPACING=0:LINE=240");
        }

        var s = post.Stats;
        trace?.Invoke(
            $"LAYOUT_STABILITY_POST:PASS:TABLES={s.Tables}:WIDTH_FIT={s.WidthFittedTables}:ROWS={s.Rows}:ROWS_NOSPLIT={s.RowsNoSplit}:HEADER_REPEAT={s.HeaderRowsRepeat}:TABLE_KEEP_CHAIN={s.TableKeepChains}:KEEP_NEXT={s.KeepNextParagraphs}:SHORT_HEADER_NOWRAP={s.ShortHeaderNoWrap}:CAPTIONS_KEEP={s.CaptionsKeepNext}:TOC={s.TocParagraphs}:TOC_PHYSICAL_RIGHT={s.TocPhysicalRight}:TOC_RUNRTL={s.TocTextRunsRtl}:PAGE_CONTENT={s.PageContentWidthTwips}");

        return s;
    }

    private static bool ApplyTableFitAndPagination(
        XDocument document,
        int pageContentWidthTwips)
    {
        var changed = false;

        foreach (var table in document.Descendants(W + "tbl"))
        {
            var rows = table.Elements(W + "tr").ToList();
            if (rows.Count == 0)
                continue;

            var columnCount = ResolveTableColumnCount(table, rows[0]);
            if (columnCount <= 0)
                continue;

            var headers = ResolveFirstRowHeaderTexts(rows[0], columnCount);
            var preferred = ResolvePreferredColumnWidths(table, rows[0], columnCount);
            var minimums = headers
                .Select(x => MinimumColumnWidthTwips(x, columnCount))
                .ToArray();
            var fitted = FitColumnWidths(preferred, minimums, pageContentWidthTwips);

            var tblPr = table.Element(W + "tblPr");
            if (tblPr is null)
            {
                tblPr = new XElement(W + "tblPr");
                table.AddFirst(tblPr);
                changed = true;
            }

            var tblW = tblPr.Element(W + "tblW");
            if (tblW is null)
            {
                tblW = new XElement(W + "tblW");
                InsertTablePropertyInSchemaOrder(tblPr, tblW);
                changed = true;
            }
            if (!string.Equals((string?)tblW.Attribute(W + "type"), "dxa", StringComparison.OrdinalIgnoreCase)
                || !string.Equals((string?)tblW.Attribute(W + "w"), pageContentWidthTwips.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal))
            {
                tblW.SetAttributeValue(W + "type", "dxa");
                tblW.SetAttributeValue(W + "w", pageContentWidthTwips.ToString(CultureInfo.InvariantCulture));
                changed = true;
            }

            var layout = tblPr.Element(W + "tblLayout");
            if (layout is null)
            {
                layout = new XElement(W + "tblLayout");
                InsertTablePropertyInSchemaOrder(tblPr, layout);
                changed = true;
            }
            if (!string.Equals((string?)layout.Attribute(W + "type"), "fixed", StringComparison.OrdinalIgnoreCase))
            {
                layout.SetAttributeValue(W + "type", "fixed");
                changed = true;
            }

            var grid = table.Element(W + "tblGrid");
            if (grid is null)
            {
                grid = new XElement(W + "tblGrid");
                var firstRow = table.Elements(W + "tr").FirstOrDefault();
                if (firstRow is null)
                    table.Add(grid);
                else
                    firstRow.AddBeforeSelf(grid);
                changed = true;
            }

            var gridColumns = grid.Elements(W + "gridCol").ToList();
            if (gridColumns.Count != columnCount)
            {
                grid.RemoveNodes();
                foreach (var width in fitted)
                    grid.Add(new XElement(W + "gridCol", new XAttribute(W + "w", width.ToString(CultureInfo.InvariantCulture))));
                changed = true;
            }
            else
            {
                for (var i = 0; i < columnCount; i++)
                {
                    var expected = fitted[i].ToString(CultureInfo.InvariantCulture);
                    if (!string.Equals((string?)gridColumns[i].Attribute(W + "w"), expected, StringComparison.Ordinal))
                    {
                        gridColumns[i].SetAttributeValue(W + "w", expected);
                        changed = true;
                    }
                }
            }

            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                var trPr = EnsureTrPr(row);

                var cantSplit = trPr.Element(W + "cantSplit");
                if (cantSplit is null)
                {
                    cantSplit = new XElement(W + "cantSplit");
                    InsertRowPropertyInSchemaOrder(trPr, cantSplit);
                    changed = true;
                }
                if (!IsOnOffEnabled(cantSplit))
                {
                    cantSplit.SetAttributeValue(W + "val", "1");
                    changed = true;
                }

                if (rowIndex == 0)
                {
                    var tblHeader = trPr.Element(W + "tblHeader");
                    if (tblHeader is null)
                    {
                        tblHeader = new XElement(W + "tblHeader");
                        InsertRowPropertyInSchemaOrder(trPr, tblHeader);
                        changed = true;
                    }
                    if (!IsOnOffEnabled(tblHeader))
                    {
                        tblHeader.SetAttributeValue(W + "val", "1");
                        changed = true;
                    }
                }

                var logicalColumn = 0;
                foreach (var cell in row.Elements(W + "tc"))
                {
                    var span = ResolveGridSpan(cell);
                    if (logicalColumn >= fitted.Count)
                        break;

                    var cellWidth = 0;
                    for (var j = 0; j < span && logicalColumn + j < fitted.Count; j++)
                        cellWidth += fitted[logicalColumn + j];

                    var tcPr = EnsureTcPr(cell);
                    var tcW = tcPr.Element(W + "tcW");
                    if (tcW is null)
                    {
                        tcW = new XElement(W + "tcW");
                        InsertCellPropertyInSchemaOrder(tcPr, tcW);
                        changed = true;
                    }

                    var expectedWidth = cellWidth.ToString(CultureInfo.InvariantCulture);
                    if (!string.Equals((string?)tcW.Attribute(W + "type"), "dxa", StringComparison.OrdinalIgnoreCase)
                        || !string.Equals((string?)tcW.Attribute(W + "w"), expectedWidth, StringComparison.Ordinal))
                    {
                        tcW.SetAttributeValue(W + "type", "dxa");
                        tcW.SetAttributeValue(W + "w", expectedWidth);
                        changed = true;
                    }

                    if (rowIndex == 0)
                    {
                        var headerText = NormalizeText(ParagraphText(cell));
                        if (headerText.Length > 0 && headerText.Length <= 8)
                        {
                            var noWrap = tcPr.Element(W + "noWrap");
                            if (noWrap is null)
                            {
                                noWrap = new XElement(W + "noWrap");
                                InsertCellPropertyInSchemaOrder(tcPr, noWrap);
                                changed = true;
                            }
                            if (!IsOnOffEnabled(noWrap))
                            {
                                noWrap.SetAttributeValue(W + "val", "1");
                                changed = true;
                            }
                        }
                    }

                    // A18: do not chain table body paragraphs with keepNext.
                    foreach (var paragraph in DirectCellParagraphs(cell))
                    {
                        var keepNext = paragraph.Element(W + "pPr")?.Element(W + "keepNext");
                        if (keepNext is not null)
                        {
                            keepNext.Remove();
                            changed = true;
                        }
                    }

                    logicalColumn += span;
                }
            }

            var previous = table.NodesBeforeSelf().OfType<XElement>().LastOrDefault();
            if (previous is not null
                && previous.Name == W + "p"
                && NormalizeText(ParagraphText(previous)).Contains("جدول", StringComparison.Ordinal))
            {
                var pPr = EnsurePPr(previous);
                var keepNext = pPr.Element(W + "keepNext");
                if (keepNext is null)
                {
                    keepNext = new XElement(W + "keepNext");
                    InsertParagraphPropertyInSchemaOrder(pPr, keepNext);
                    changed = true;
                }
                if (!IsOnOffEnabled(keepNext))
                {
                    keepNext.SetAttributeValue(W + "val", "1");
                    changed = true;
                }
            }
        }

        return changed;
    }

    private static bool ApplyAutomaticTocPhysicalRtl(
        XDocument document,
        IReadOnlySet<string> tocStyleTokens)
    {
        var changed = false;

        foreach (var paragraph in document.Descendants(W + "p")
                     .Where(p => IsAutomaticTocParagraph(p, tocStyleTokens)))
        {
            var pPr = EnsurePPr(paragraph);

            var bidi = pPr.Element(W + "bidi");
            if (bidi is null)
            {
                bidi = new XElement(W + "bidi");
                InsertParagraphPropertyInSchemaOrder(pPr, bidi);
                changed = true;
            }
            if (!IsOnOffEnabled(bidi))
            {
                bidi.SetAttributeValue(W + "val", "1");
                changed = true;
            }

            var jc = pPr.Element(W + "jc");
            if (jc is null)
            {
                jc = new XElement(W + "jc");
                InsertParagraphPropertyInSchemaOrder(pPr, jc);
                changed = true;
            }

            const string physicalRightToken = "right"; // A19: literal Word Align Right
            var currentJc = ((string?)jc.Attribute(W + "val") ?? string.Empty).Trim();
            if (!currentJc.Equals(physicalRightToken, StringComparison.OrdinalIgnoreCase))
            {
                jc.SetAttributeValue(W + "val", physicalRightToken);
                changed = true;
            }

            foreach (var run in paragraph.Descendants(W + "r"))
            {
                if (!RunHasVisibleOrFieldContent(run))
                    continue;

                var rPr = run.Element(W + "rPr");
                if (rPr is null)
                {
                    rPr = new XElement(W + "rPr");
                    run.AddFirst(rPr);
                    changed = true;
                }

                var rtl = rPr.Element(W + "rtl");
                if (rtl is null)
                {
                    rtl = new XElement(W + "rtl");
                    InsertRunPropertyInSchemaOrder(rPr, rtl);
                    changed = true;
                }
                if (!IsOnOffEnabled(rtl))
                {
                    rtl.SetAttributeValue(W + "val", "1");
                    changed = true;
                }
            }
        }

        return changed;
    }

    private sealed record TocTabLayoutInspection(
        int Paragraphs, int ParagraphTabsOk, int Styles, int StyleTabsOk,
        int TabPosition, bool FullyCompliant, IReadOnlyList<string> Issues);

    private static bool ApplyAutomaticTocTabLayout(
        XDocument document,
        XDocument styles,
        int pageContentWidthTwips,
        IReadOnlySet<string> tocStyleTokens)
    {
        var changed = false;
        var tabPosition = Math.Max(1, pageContentWidthTwips - 1);
        const string physicalRightToken = "right"; // A19: literal Word Align Right
        var tocParagraphs = document.Descendants(W + "p")
            .Where(p => IsAutomaticTocParagraph(p, tocStyleTokens))
            .ToList();
        var styleIds = tocParagraphs
            .Select(p => ((string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") ?? string.Empty).Trim())
            .Where(x => x.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();

        foreach (var paragraph in tocParagraphs)
        {
            var pPr = EnsurePPr(paragraph);
            if (NormalizeSingleTocTab(pPr, tabPosition)) changed = true;
            if (NormalizeCompactTocSpacing(pPr)) changed = true;
            var bidi = pPr.Element(W + "bidi");
            if (bidi is null) { bidi = new XElement(W + "bidi"); InsertParagraphPropertyInSchemaOrder(pPr, bidi); changed = true; }
            if (!IsOnOffEnabled(bidi)) { bidi.SetAttributeValue(W + "val", "1"); changed = true; }
            var jc = pPr.Element(W + "jc");
            if (jc is null) { jc = new XElement(W + "jc"); InsertParagraphPropertyInSchemaOrder(pPr, jc); changed = true; }
            if (!string.Equals((string?)jc.Attribute(W + "val"), physicalRightToken, StringComparison.OrdinalIgnoreCase))
            { jc.SetAttributeValue(W + "val", physicalRightToken); changed = true; }
        }

        foreach (var styleId in styleIds)
        {
            var style = styles.Descendants(W + "style").FirstOrDefault(x =>
                string.Equals(((string?)x.Attribute(W + "styleId") ?? string.Empty).Trim(), styleId, StringComparison.OrdinalIgnoreCase));
            if (style is null) continue;
            var pPr = style.Element(W + "pPr");
            if (pPr is null)
            {
                pPr = new XElement(W + "pPr");
                var rPr = style.Element(W + "rPr");
                if (rPr is null) style.Add(pPr); else rPr.AddBeforeSelf(pPr);
                changed = true;
            }
            if (NormalizeSingleTocTab(pPr, tabPosition)) changed = true;
            if (NormalizeCompactTocSpacing(pPr)) changed = true;
            var bidi = pPr.Element(W + "bidi");
            if (bidi is null) { bidi = new XElement(W + "bidi"); InsertParagraphPropertyInSchemaOrder(pPr, bidi); changed = true; }
            if (!IsOnOffEnabled(bidi)) { bidi.SetAttributeValue(W + "val", "1"); changed = true; }
            var jc = pPr.Element(W + "jc");
            if (jc is null) { jc = new XElement(W + "jc"); InsertParagraphPropertyInSchemaOrder(pPr, jc); changed = true; }
            if (!string.Equals((string?)jc.Attribute(W + "val"), physicalRightToken, StringComparison.OrdinalIgnoreCase))
            { jc.SetAttributeValue(W + "val", physicalRightToken); changed = true; }
        }
        return changed;
    }

    private static bool NormalizeCompactTocSpacing(XElement pPr)
    {
        var changed = false;
        var spacing = pPr.Element(W + "spacing");
        if (spacing is null)
        {
            spacing = new XElement(W + "spacing");
            InsertParagraphPropertyInSchemaOrder(pPr, spacing);
            changed = true;
        }

        if (((string?)spacing.Attribute(W + "before") ?? string.Empty).Trim() != "0")
        {
            spacing.SetAttributeValue(W + "before", "0");
            changed = true;
        }
        if (((string?)spacing.Attribute(W + "after") ?? string.Empty).Trim() != "0")
        {
            spacing.SetAttributeValue(W + "after", "0");
            changed = true;
        }
        if (((string?)spacing.Attribute(W + "line") ?? string.Empty).Trim() != "240")
        {
            spacing.SetAttributeValue(W + "line", "240");
            changed = true;
        }
        if (!((string?)spacing.Attribute(W + "lineRule") ?? string.Empty)
            .Trim().Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            spacing.SetAttributeValue(W + "lineRule", "auto");
            changed = true;
        }

        return changed;
    }

    private static bool IsCompactTocSpacing(XElement? pPr)
    {
        var spacing = pPr?.Element(W + "spacing");
        if (spacing is null)
            return false;

        return ((string?)spacing.Attribute(W + "before") ?? string.Empty).Trim() == "0"
            && ((string?)spacing.Attribute(W + "after") ?? string.Empty).Trim() == "0"
            && ((string?)spacing.Attribute(W + "line") ?? string.Empty).Trim() == "240"
            && ((string?)spacing.Attribute(W + "lineRule") ?? string.Empty)
                .Trim().Equals("auto", StringComparison.OrdinalIgnoreCase);
    }

    private static bool NormalizeSingleTocTab(XElement pPr, int tabPosition)
    {
        var changed = false;
        var tabs = pPr.Element(W + "tabs");
        if (tabs is null) { tabs = new XElement(W + "tabs"); InsertParagraphPropertyInSchemaOrder(pPr, tabs); changed = true; }
        var existing = tabs.Elements(W + "tab").ToList();
        var exact = existing.Count == 1
            && string.Equals((string?)existing[0].Attribute(W + "val"), "right", StringComparison.OrdinalIgnoreCase)
            && ParsePositiveInt((string?)existing[0].Attribute(W + "pos")) == tabPosition
            && string.Equals((string?)existing[0].Attribute(W + "leader"), "dot", StringComparison.OrdinalIgnoreCase);
        if (!exact)
        {
            tabs.RemoveNodes();
            tabs.Add(new XElement(W + "tab",
                new XAttribute(W + "val", "right"),
                new XAttribute(W + "pos", tabPosition.ToString(CultureInfo.InvariantCulture)),
                new XAttribute(W + "leader", "dot")));
            changed = true;
        }
        return changed;
    }

    private static TocTabLayoutInspection InspectAutomaticTocTabLayout(
        XDocument document,
        XDocument styles,
        int pageContentWidthTwips,
        IReadOnlySet<string> tocStyleTokens)
    {
        var issues = new List<string>();
        var tabPosition = Math.Max(1, pageContentWidthTwips - 1);
        const string physicalRightToken = "right"; // A19: literal Word Align Right
        var tocParagraphs = document.Descendants(W + "p")
            .Where(p => IsAutomaticTocParagraph(p, tocStyleTokens))
            .ToList();
        var paragraphTabsOk = 0;

        foreach (var paragraph in tocParagraphs)
        {
            var pPr = paragraph.Element(W + "pPr");
            var tabs = pPr?.Element(W + "tabs")?.Elements(W + "tab").ToList() ?? new List<XElement>();
            var tabOk = tabs.Count == 1
                && string.Equals((string?)tabs[0].Attribute(W + "val"), "right", StringComparison.OrdinalIgnoreCase)
                && ParsePositiveInt((string?)tabs[0].Attribute(W + "pos")) == tabPosition
                && string.Equals((string?)tabs[0].Attribute(W + "leader"), "dot", StringComparison.OrdinalIgnoreCase);
            var bidiOk = IsOnOffEnabled(pPr?.Element(W + "bidi"));
            var spacingOk = IsCompactTocSpacing(pPr);
            var jc = ((string?)pPr?.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();
            if (tabOk && bidiOk && spacingOk && jc.Equals(physicalRightToken, StringComparison.OrdinalIgnoreCase)) paragraphTabsOk++;
            else issues.Add($"TOC paragraph compact/right mismatch: tabOk={tabOk}, bidiOk={bidiOk}, spacingOk={spacingOk}, jc='{jc}'.");
        }

        var styleIds = tocParagraphs
            .Select(p => ((string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") ?? string.Empty).Trim())
            .Where(x => x.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        var styleTabsOk = 0;

        foreach (var styleId in styleIds)
        {
            var style = styles.Descendants(W + "style").FirstOrDefault(x =>
                string.Equals(((string?)x.Attribute(W + "styleId") ?? string.Empty).Trim(), styleId, StringComparison.OrdinalIgnoreCase));
            if (style is null) { issues.Add($"TOC style '{styleId}' is missing."); continue; }
            var pPr = style.Element(W + "pPr");
            var tabs = pPr?.Element(W + "tabs")?.Elements(W + "tab").ToList() ?? new List<XElement>();
            var tabOk = tabs.Count == 1
                && string.Equals((string?)tabs[0].Attribute(W + "val"), "right", StringComparison.OrdinalIgnoreCase)
                && ParsePositiveInt((string?)tabs[0].Attribute(W + "pos")) == tabPosition
                && string.Equals((string?)tabs[0].Attribute(W + "leader"), "dot", StringComparison.OrdinalIgnoreCase);
            var bidiOk = IsOnOffEnabled(pPr?.Element(W + "bidi"));
            var spacingOk = IsCompactTocSpacing(pPr);
            var jc = ((string?)pPr?.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();
            if (tabOk && bidiOk && spacingOk && jc.Equals(physicalRightToken, StringComparison.OrdinalIgnoreCase)) styleTabsOk++;
            else issues.Add($"TOC style '{styleId}' compact/right mismatch: tabOk={tabOk}, bidiOk={bidiOk}, spacingOk={spacingOk}, jc='{jc}'.");
        }

        var ok = tocParagraphs.Count > 0 && paragraphTabsOk == tocParagraphs.Count
            && styleIds.Length > 0 && styleTabsOk == styleIds.Length && issues.Count == 0;
        return new TocTabLayoutInspection(tocParagraphs.Count, paragraphTabsOk, styleIds.Length, styleTabsOk, tabPosition, ok, issues);
    }

    private static LayoutStabilityInspection InspectLayoutStabilityAndToc(
        XDocument document,
        int pageContentWidthTwips,
        IReadOnlySet<string> tocStyleTokens)
    {
        var issues = new List<string>();
        var tables = 0;
        var rowsCount = 0;
        var widthFitted = 0;
        var rowsNoSplit = 0;
        var headerRowsRepeat = 0;
        var keepNextParagraphs = 0;
        var tableKeepChains = 0;
        var shortHeaderNoWrap = 0;
        var captionsKeepNext = 0;

        foreach (var table in document.Descendants(W + "tbl"))
        {
            tables++;
            var rows = table.Elements(W + "tr").ToList();
            rowsCount += rows.Count;

            var columnCount = ResolveTableColumnCount(table, rows.FirstOrDefault());
            var grid = table.Element(W + "tblGrid");
            var gridWidths = grid?.Elements(W + "gridCol")
                .Select(x => ParsePositiveInt((string?)x.Attribute(W + "w")))
                .ToArray() ?? Array.Empty<int>();

            var tblW = table.Element(W + "tblPr")?.Element(W + "tblW");
            var tableWidthOk = columnCount > 0
                && gridWidths.Length == columnCount
                && gridWidths.All(x => x > 0)
                && gridWidths.Sum() == pageContentWidthTwips
                && string.Equals((string?)tblW?.Attribute(W + "type"), "dxa", StringComparison.OrdinalIgnoreCase)
                && ParsePositiveInt((string?)tblW?.Attribute(W + "w")) == pageContentWidthTwips;

            if (tableWidthOk)
                widthFitted++;
            else
                issues.Add($"T{tables}: table/grid width is not fitted to {pageContentWidthTwips} twips.");

            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                var trPr = row.Element(W + "trPr");

                if (IsOnOffEnabled(trPr?.Element(W + "cantSplit")))
                    rowsNoSplit++;
                else
                    issues.Add($"T{tables}/R{rowIndex + 1}: row can split across pages.");

                if (rowIndex == 0)
                {
                    if (IsOnOffEnabled(trPr?.Element(W + "tblHeader")))
                        headerRowsRepeat++;
                    else
                        issues.Add($"T{tables}: first row is not marked as repeating table header.");
                }

                var logicalColumn = 0;
                foreach (var cell in row.Elements(W + "tc"))
                {
                    var span = ResolveGridSpan(cell);
                    if (tableWidthOk && logicalColumn < gridWidths.Length)
                    {
                        var expectedCellWidth = 0;
                        for (var j = 0; j < span && logicalColumn + j < gridWidths.Length; j++)
                            expectedCellWidth += gridWidths[logicalColumn + j];

                        var tcW = cell.Element(W + "tcPr")?.Element(W + "tcW");
                        var actual = ParsePositiveInt((string?)tcW?.Attribute(W + "w"));
                        var type = (string?)tcW?.Attribute(W + "type");
                        if (actual != expectedCellWidth
                            || !string.Equals(type, "dxa", StringComparison.OrdinalIgnoreCase))
                        {
                            issues.Add($"T{tables}/R{rowIndex + 1}/C{logicalColumn + 1}: cell width mismatch actual={actual} expected={expectedCellWidth}.");
                        }
                    }

                    if (rowIndex == 0)
                    {
                        var headerText = NormalizeText(ParagraphText(cell));
                        if (headerText.Length > 0 && headerText.Length <= 8)
                        {
                            if (IsOnOffEnabled(cell.Element(W + "tcPr")?.Element(W + "noWrap")))
                                shortHeaderNoWrap++;
                            else
                                issues.Add($"T{tables}/C{logicalColumn + 1}: short header '{headerText}' is allowed to wrap.");
                        }
                    }

                    foreach (var paragraph in DirectCellParagraphs(cell))
                    {
                        if (IsOnOffEnabled(paragraph.Element(W + "pPr")?.Element(W + "keepNext")))
                        {
                            keepNextParagraphs++;
                            issues.Add($"T{tables}/R{rowIndex + 1}: table-cell keepNext over-constrains pagination.");
                        }
                    }

                    logicalColumn += span;
                }
            }

            tableKeepChains += 0;

            var previous = table.NodesBeforeSelf().OfType<XElement>().LastOrDefault();
            if (previous is not null
                && previous.Name == W + "p"
                && NormalizeText(ParagraphText(previous)).Contains("جدول", StringComparison.Ordinal))
            {
                if (IsOnOffEnabled(previous.Element(W + "pPr")?.Element(W + "keepNext")))
                    captionsKeepNext++;
                else
                    issues.Add($"T{tables}: table caption is not kept with the table.");
            }
        }

        var tocParagraphs = 0;
        var tocPhysicalRight = 0;
        var tocTextRunsRtl = 0;
        const string physicalRightToken = "right"; // A19: literal Word Align Right

        foreach (var paragraph in document.Descendants(W + "p").Where(p => IsAutomaticTocParagraph(p, tocStyleTokens)))
        {
            tocParagraphs++;
            var pPr = paragraph.Element(W + "pPr");
            var bidi = IsOnOffEnabled(pPr?.Element(W + "bidi"));
            var jc = ((string?)pPr?.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();

            if (bidi && jc.Equals(physicalRightToken, StringComparison.OrdinalIgnoreCase))
                tocPhysicalRight++;
            else
                issues.Add($"TOC P{tocParagraphs}: paragraph is not physical-right RTL.");

            var visibleRuns = paragraph.Descendants(W + "r")
                .Where(RunHasVisibleOrFieldContent)
                .ToList();
            if (visibleRuns.Count > 0
                && visibleRuns.All(r => IsOnOffEnabled(r.Element(W + "rPr")?.Element(W + "rtl"))))
            {
                tocTextRunsRtl++;
            }
            else
            {
                issues.Add($"TOC P{tocParagraphs}: one or more visible/field runs are not RTL.");
            }
        }

        var stats = new LayoutStabilityStats(
            tables,
            rowsCount,
            widthFitted,
            rowsNoSplit,
            headerRowsRepeat,
            tableKeepChains,
            keepNextParagraphs,
            shortHeaderNoWrap,
            captionsKeepNext,
            tocParagraphs,
            tocPhysicalRight,
            tocTextRunsRtl,
            pageContentWidthTwips);

        // QH01-E271: zero-table documents are valid inputs. Table obligations are
        // universal over existing tables, not an existence requirement. This keeps the
        // strict TOC postcondition while allowing a table-free book fixture to qualify.
        var tableLayoutCompliant =
            widthFitted == tables
            && rowsNoSplit == rowsCount
            && headerRowsRepeat == tables
            && tableKeepChains == 0
            && keepNextParagraphs == 0;
        var fullyCompliant = tableLayoutCompliant
            && tocParagraphs > 0
            && tocPhysicalRight == tocParagraphs
            && tocTextRunsRtl == tocParagraphs
            && issues.Count == 0;

        return new LayoutStabilityInspection(stats, fullyCompliant, issues);
    }

    private static string NormalizeAutomaticTocStyleToken(string? value)
    {
        var normalized = NormalizeText(value);
        return new string(normalized
            .Where(ch => !char.IsWhiteSpace(ch) && ch != '-' && ch != '_')
            .ToArray())
            .ToUpperInvariant();
    }

    private static HashSet<string> ResolveAutomaticTocStyleTokens(
        XDocument styles,
        AgentProfile profile)
    {
        var aliasTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var pair in profile.RoleAliases)
        {
            var canonical = CanonicalizeControlledRoleAliasFamily(pair.Key);
            if (!canonical.Equals("Toc1", StringComparison.OrdinalIgnoreCase)
                && !canonical.Equals("Toc2", StringComparison.OrdinalIgnoreCase))
                continue;

            foreach (var alias in pair.Value ?? Array.Empty<string>())
            {
                var token = NormalizeAutomaticTocStyleToken(alias);
                if (token.Length > 0)
                    aliasTokens.Add(token);
            }
        }

        foreach (var builtIn in new[] { "TOC1", "TOC2", "TOC 1", "TOC 2" })
            aliasTokens.Add(NormalizeAutomaticTocStyleToken(builtIn));

        var resolved = new HashSet<string>(aliasTokens, StringComparer.OrdinalIgnoreCase);
        foreach (var style in styles.Descendants(W + "style"))
        {
            var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Trim();
            var styleName = ((string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty).Trim();
            var idToken = NormalizeAutomaticTocStyleToken(styleId);
            var nameToken = NormalizeAutomaticTocStyleToken(styleName);

            if (aliasTokens.Contains(idToken) || aliasTokens.Contains(nameToken))
            {
                if (idToken.Length > 0) resolved.Add(idToken);
                if (nameToken.Length > 0) resolved.Add(nameToken);
            }
        }

        return resolved;
    }

    private static bool IsAutomaticTocParagraph(
        XElement paragraph,
        IReadOnlySet<string> tocStyleTokens)
    {
        var styleId = ((string?)paragraph
                .Element(W + "pPr")?
                .Element(W + "pStyle")?
                .Attribute(W + "val") ?? string.Empty)
            .Trim();

        var token = NormalizeAutomaticTocStyleToken(styleId);
        if (token.Length > 0 && tocStyleTokens.Contains(token))
            return true;

        // R8R23R4 deep-research closure:
        // Persisted automatic-TOC structure is semantic authority when Word has regenerated
        // a TOC whose locale/document-specific paragraph style id/name is not present in the
        // damaged source profile aliases. PAGEREF _Toc... and _Toc... hyperlink anchors are
        // already used elsewhere in this service as strong automatic-TOC signatures.
        // This broadens discovery only; all strict post-Word layout/tab/text/field gates remain.
        return HasAutomaticTocEntrySignature(paragraph);
    }

    private static int ResolveMinimumPageContentWidthTwips(XDocument document)
    {
        var widths = new List<int>();

        foreach (var sectPr in document.Descendants(W + "sectPr"))
        {
            var pgSz = sectPr.Element(W + "pgSz");
            var pgMar = sectPr.Element(W + "pgMar");

            var page = ParsePositiveInt((string?)pgSz?.Attribute(W + "w"));
            var left = ParsePositiveInt((string?)pgMar?.Attribute(W + "left"));
            var right = ParsePositiveInt((string?)pgMar?.Attribute(W + "right"));
            var gutter = ParsePositiveInt((string?)pgMar?.Attribute(W + "gutter"));

            var content = page - left - right - gutter;
            if (content > 0)
                widths.Add(content);
        }

        return widths.Count > 0 ? widths.Min() : 9072;
    }

    private static int ResolveTableColumnCount(XElement table, XElement? firstRow)
    {
        var gridCount = table.Element(W + "tblGrid")?.Elements(W + "gridCol").Count() ?? 0;
        if (gridCount > 0)
            return gridCount;

        if (firstRow is null)
            return 0;

        return firstRow.Elements(W + "tc").Sum(ResolveGridSpan);
    }

    private static string[] ResolveFirstRowHeaderTexts(XElement firstRow, int columnCount)
    {
        var result = Enumerable.Repeat(string.Empty, columnCount).ToArray();
        var logicalColumn = 0;

        foreach (var cell in firstRow.Elements(W + "tc"))
        {
            var span = ResolveGridSpan(cell);
            var text = NormalizeText(ParagraphText(cell));
            for (var i = 0; i < span && logicalColumn + i < result.Length; i++)
                result[logicalColumn + i] = text;
            logicalColumn += span;
        }

        return result;
    }

    private static int[] ResolvePreferredColumnWidths(
        XElement table,
        XElement firstRow,
        int columnCount)
    {
        var gridWidths = table.Element(W + "tblGrid")?.Elements(W + "gridCol")
            .Select(x => ParsePositiveInt((string?)x.Attribute(W + "w")))
            .ToArray() ?? Array.Empty<int>();

        var firstCells = firstRow.Elements(W + "tc").ToList();
        var simpleFirstRow = firstCells.Count == columnCount
            && firstCells.All(x => ResolveGridSpan(x) == 1);

        if (simpleFirstRow)
        {
            var cellWidths = firstCells
                .Select(x => ParsePositiveInt((string?)x.Element(W + "tcPr")?.Element(W + "tcW")?.Attribute(W + "w")))
                .ToArray();

            if (cellWidths.All(x => x > 0))
                return cellWidths;
        }

        if (gridWidths.Length == columnCount && gridWidths.All(x => x > 0))
            return gridWidths;

        return Enumerable.Repeat(Math.Max(1, 9072 / Math.Max(1, columnCount)), columnCount).ToArray();
    }

    private static int MinimumColumnWidthTwips(string headerText, int columnCount)
    {
        var length = NormalizeText(headerText).Length;
        var baseWidth = columnCount >= 8 ? 650 : 800;

        if (length <= 4) return baseWidth;
        if (length <= 6) return baseWidth + 80;
        if (length <= 10) return baseWidth + 250;
        if (length <= 16) return baseWidth + 450;
        return baseWidth + 650;
    }

    private static IReadOnlyList<int> FitColumnWidths(
        IReadOnlyList<int> preferred,
        IReadOnlyList<int> minimums,
        int totalWidth)
    {
        if (preferred.Count == 0 || preferred.Count != minimums.Count)
            throw new InvalidOperationException("Table width allocation input is invalid.");

        var count = preferred.Count;
        var mins = minimums.Select(x => Math.Max(450, x)).ToArray();
        var prefs = preferred
            .Select((x, i) => Math.Max(Math.Max(1, x), mins[i]))
            .ToArray();

        var minSum = mins.Sum();
        var result = new int[count];

        if (minSum >= totalWidth)
        {
            var scale = (double)totalWidth / minSum;
            for (var i = 0; i < count; i++)
                result[i] = Math.Max(450, (int)Math.Floor(mins[i] * scale));
        }
        else
        {
            var excess = prefs.Select((x, i) => Math.Max(0, x - mins[i])).ToArray();
            var excessSum = excess.Sum();
            var remaining = totalWidth - minSum;

            if (excessSum > 0 && remaining <= excessSum)
            {
                for (var i = 0; i < count; i++)
                    result[i] = mins[i] + (int)Math.Floor((double)remaining * excess[i] / excessSum);
            }
            else
            {
                Array.Copy(prefs, result, count);
                var current = result.Sum();
                var extra = totalWidth - current;

                if (extra > 0)
                {
                    var weightSum = prefs.Sum();
                    for (var i = 0; i < count; i++)
                        result[i] += (int)Math.Floor((double)extra * prefs[i] / weightSum);
                }
            }
        }

        var diff = totalWidth - result.Sum();
        var widest = 0;
        for (var i = 1; i < count; i++)
        {
            if (result[i] > result[widest])
                widest = i;
        }

        result[widest] += diff;
        if (result.Any(x => x <= 0) || result.Sum() != totalWidth)
            throw new InvalidOperationException("Table width allocation could not produce a positive exact-fit grid.");

        return result;
    }

    private static int ResolveGridSpan(XElement cell)
    {
        var value = ParsePositiveInt((string?)cell
            .Element(W + "tcPr")?
            .Element(W + "gridSpan")?
            .Attribute(W + "val"));

        return value > 0 ? value : 1;
    }

    private static int ParsePositiveInt(string? value)
        => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0
            ? parsed
            : 0;

    private static XElement EnsureTrPr(XElement row)
    {
        var trPr = row.Element(W + "trPr");
        if (trPr is not null)
            return trPr;

        trPr = new XElement(W + "trPr");
        row.AddFirst(trPr);
        return trPr;
    }

    private static XElement EnsureTcPr(XElement cell)
    {
        var tcPr = cell.Element(W + "tcPr");
        if (tcPr is not null)
            return tcPr;

        tcPr = new XElement(W + "tcPr");
        cell.AddFirst(tcPr);
        return tcPr;
    }

    private static void InsertRowPropertyInSchemaOrder(
        XElement trPr,
        XElement property)
    {
        static int Rank(string name) => name switch
        {
            "cnfStyle" => 10,
            "divId" => 20,
            "gridBefore" => 30,
            "gridAfter" => 40,
            "wBefore" => 50,
            "wAfter" => 60,
            "cantSplit" => 70,
            "trHeight" => 80,
            "tblHeader" => 90,
            "tblCellSpacing" => 100,
            "jc" => 110,
            "hidden" => 120,
            "ins" => 130,
            "del" => 140,
            "trPrChange" => 150,
            _ => 1000
        };

        var targetRank = Rank(property.Name.LocalName);
        var next = trPr.Elements().FirstOrDefault(e => Rank(e.Name.LocalName) > targetRank);
        if (next is null)
            trPr.Add(property);
        else
            next.AddBeforeSelf(property);
    }

    private static void InsertCellPropertyInSchemaOrder(
        XElement tcPr,
        XElement property)
    {
        static int Rank(string name) => name switch
        {
            "cnfStyle" => 10,
            "tcW" => 20,
            "gridSpan" => 30,
            "hMerge" => 40,
            "vMerge" => 50,
            "tcBorders" => 60,
            "shd" => 70,
            "noWrap" => 80,
            "tcMar" => 90,
            "textDirection" => 100,
            "tcFitText" => 110,
            "vAlign" => 120,
            "hideMark" => 130,
            "headers" => 140,
            "cellIns" => 150,
            "cellDel" => 160,
            "cellMerge" => 170,
            "tcPrChange" => 180,
            _ => 1000
        };

        var targetRank = Rank(property.Name.LocalName);
        var next = tcPr.Elements().FirstOrDefault(e => Rank(e.Name.LocalName) > targetRank);
        if (next is null)
            tcPr.Add(property);
        else
            next.AddBeforeSelf(property);
    }

    private sealed record HeaderFooterVisualInspection(
        int Paragraphs,
        bool FullyCompliant,
        IReadOnlyList<string> Issues);

    private static bool ApplyHeaderFooterVisualFormatting(
        XDocument story,
        string fontName)
    {
        var changed = false;

        foreach (var paragraph in story.Descendants(W + "p"))
        {
            if (NormalizeText(ParagraphText(paragraph)).Length == 0)
                continue;

            var pPr = EnsurePPr(paragraph);

            var bidi = pPr.Element(W + "bidi");
            if (bidi is null)
            {
                bidi = new XElement(W + "bidi");
                InsertParagraphPropertyInSchemaOrder(pPr, bidi);
                changed = true;
            }
            if (!IsOnOffEnabled(bidi))
            {
                bidi.SetAttributeValue(W + "val", "1");
                changed = true;
            }

            var jc = pPr.Element(W + "jc");
            if (jc is null)
            {
                jc = new XElement(W + "jc");
                InsertParagraphPropertyInSchemaOrder(pPr, jc);
                changed = true;
            }

            var expectedJc = EncodeAlignmentForWordOoxml("RIGHT", "RTL");
            var currentJc = ((string?)jc.Attribute(W + "val") ?? string.Empty).Trim();
            if (!currentJc.Equals(expectedJc, StringComparison.OrdinalIgnoreCase))
            {
                jc.SetAttributeValue(W + "val", expectedJc);
                changed = true;
            }

            foreach (var run in paragraph.Descendants(W + "r"))
            {
                if (!RunHasVisibleOrFieldContent(run))
                    continue;

                var rPr = run.Element(W + "rPr");
                if (rPr is null)
                {
                    rPr = new XElement(W + "rPr");
                    run.AddFirst(rPr);
                    changed = true;
                }

                var rFonts = rPr.Element(W + "rFonts");
                if (rFonts is null)
                {
                    rFonts = new XElement(W + "rFonts");
                    InsertRunPropertyInSchemaOrder(rPr, rFonts);
                    changed = true;
                }

                foreach (var attributeName in new[] { "ascii", "hAnsi", "eastAsia", "cs" })
                {
                    var current = ((string?)rFonts.Attribute(W + attributeName) ?? string.Empty).Trim();
                    if (!current.Equals(fontName, StringComparison.OrdinalIgnoreCase))
                    {
                        rFonts.SetAttributeValue(W + attributeName, fontName);
                        changed = true;
                    }
                }

                foreach (var themeAttribute in new[] { "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme" })
                {
                    var attr = rFonts.Attribute(W + themeAttribute);
                    if (attr is not null)
                    {
                        attr.Remove();
                        changed = true;
                    }
                }

                var rtl = rPr.Element(W + "rtl");
                if (rtl is null)
                {
                    rtl = new XElement(W + "rtl");
                    InsertRunPropertyInSchemaOrder(rPr, rtl);
                    changed = true;
                }
                if (!IsOnOffEnabled(rtl))
                {
                    rtl.SetAttributeValue(W + "val", "1");
                    changed = true;
                }
            }
        }

        return changed;
    }

    private static HeaderFooterVisualInspection InspectHeaderFooterVisualFormatting(
        XDocument story,
        string fontName)
    {
        var issues = new List<string>();
        var paragraphCount = 0;

        foreach (var paragraph in story.Descendants(W + "p"))
        {
            if (NormalizeText(ParagraphText(paragraph)).Length == 0)
                continue;

            paragraphCount++;
            var pPr = paragraph.Element(W + "pPr");

            if (!IsOnOffEnabled(pPr?.Element(W + "bidi")))
                issues.Add($"P{paragraphCount}: Header/Footer paragraph is not RTL.");

            var expectedJc = EncodeAlignmentForWordOoxml("RIGHT", "RTL");
            var actualJc = ((string?)pPr?.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();
            if (!actualJc.Equals(expectedJc, StringComparison.OrdinalIgnoreCase))
                issues.Add($"P{paragraphCount}: Header/Footer paragraph is not physically right-aligned.");

            foreach (var run in paragraph.Descendants(W + "r"))
            {
                if (!RunHasVisibleOrFieldContent(run))
                    continue;

                var rPr = run.Element(W + "rPr");
                var rFonts = rPr?.Element(W + "rFonts");
                foreach (var attributeName in new[] { "ascii", "hAnsi", "eastAsia", "cs" })
                {
                    var actual = ((string?)rFonts?.Attribute(W + attributeName) ?? string.Empty).Trim();
                    if (!actual.Equals(fontName, StringComparison.OrdinalIgnoreCase))
                        issues.Add($"P{paragraphCount}: Header/Footer run font '{attributeName}' is '{actual}', expected '{fontName}'.");
                }

                if (!IsOnOffEnabled(rPr?.Element(W + "rtl")))
                    issues.Add($"P{paragraphCount}: Header/Footer run is not RTL.");
            }
        }

        return new HeaderFooterVisualInspection(
            paragraphCount,
            issues.Count == 0,
            issues);
    }

    private static bool IsHeaderFooterStoryPart(string name)
        => (name.StartsWith("word/header", StringComparison.OrdinalIgnoreCase)
            || name.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase))
           && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

    private static IEnumerable<XElement> DirectCellParagraphs(XElement cell)
        => cell.Descendants(W + "p")
            .Where(p => ReferenceEquals(p.Ancestors(W + "tc").FirstOrDefault(), cell));

    private static string DecideTableCellAlignment(
        string text,
        ProfileTableStyleLayoutPolicy policy)
    {
        var compact = NormalizeText(text);
        if (compact.Length == 0)
            return policy.ShortCellAlignment;
        if (LooksNumericOrCode(compact))
            return policy.ShortCellAlignment;
        if (ContainsPersian(compact) && compact.Length > policy.LongTextThreshold)
            return policy.LongPersianAlignment;
        return policy.ShortCellAlignment;
    }

    private static bool LooksNumericOrCode(string text)
    {
        var letters = 0;
        var digits = 0;

        foreach (var ch in text)
        {
            if (char.IsLetter(ch))
                letters++;
            if (char.IsDigit(ch)
                || (ch >= '\u06F0' && ch <= '\u06F9')
                || (ch >= '\u0660' && ch <= '\u0669'))
                digits++;
        }

        return digits > 0 && letters <= Math.Max(2, digits / 2);
    }

    private static bool ContainsPersian(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        foreach (var ch in text)
        {
            if ((ch >= '\u0600' && ch <= '\u06FF')
                || (ch >= '\u0750' && ch <= '\u077F')
                || (ch >= '\u08A0' && ch <= '\u08FF'))
                return true;
        }

        return false;
    }

    private static bool MaterializeTableCellParagraphAlignment(
        XElement paragraph,
        string alignment,
        bool rtl)
    {
        var changed = false;
        var pPr = EnsurePPr(paragraph);

        var jc = pPr.Element(W + "jc");
        if (jc is null)
        {
            jc = new XElement(W + "jc");
            InsertParagraphPropertyInSchemaOrder(pPr, jc);
            changed = true;
        }

        var expectedJc = EncodeAlignmentForWordOoxml(
            alignment,
            rtl ? "RTL" : null);
        var currentJc = ((string?)jc.Attribute(W + "val") ?? string.Empty).Trim();
        if (!currentJc.Equals(expectedJc, StringComparison.OrdinalIgnoreCase))
        {
            jc.SetAttributeValue(W + "val", expectedJc);
            changed = true;
        }

        if (rtl)
        {
            var bidi = pPr.Element(W + "bidi");
            if (bidi is null)
            {
                bidi = new XElement(W + "bidi");
                InsertParagraphPropertyInSchemaOrder(pPr, bidi);
                changed = true;
            }
            if (!IsOnOffEnabled(bidi))
            {
                bidi.SetAttributeValue(W + "val", "1");
                changed = true;
            }
        }

        return changed;
    }

    private static bool RunHasVisibleOrFieldContent(XElement run)
        => run.Descendants(W + "t").Any()
           || run.Descendants(W + "instrText").Any()
           || run.Descendants(W + "fldChar").Any();

    private static bool IsOnOffEnabled(XElement? element)
    {
        if (element is null)
            return false;

        var value = ((string?)element.Attribute(W + "val") ?? string.Empty)
            .Trim()
            .ToLowerInvariant();

        return value is not ("0" or "false" or "off" or "no");
    }

    private static void InsertTablePropertyInSchemaOrder(
        XElement tblPr,
        XElement property)
    {
        static int Rank(string name) => name switch
        {
            "tblStyle" => 10,
            "tblpPr" => 20,
            "tblOverlap" => 30,
            "bidiVisual" => 40,
            "tblStyleRowBandSize" => 50,
            "tblStyleColBandSize" => 60,
            "tblW" => 70,
            "jc" => 80,
            "tblCellSpacing" => 90,
            "tblInd" => 100,
            "tblBorders" => 110,
            "shd" => 120,
            "tblLayout" => 130,
            "tblCellMar" => 140,
            "tblLook" => 150,
            "tblCaption" => 160,
            "tblDescription" => 170,
            _ => 1000
        };

        var targetRank = Rank(property.Name.LocalName);
        var next = tblPr.Elements().FirstOrDefault(e => Rank(e.Name.LocalName) > targetRank);
        if (next is null)
            tblPr.Add(property);
        else
            next.AddBeforeSelf(property);
    }

    private static void InsertRunPropertyInSchemaOrder(
        XElement rPr,
        XElement property)
    {
        static int Rank(string name) => name switch
        {
            "rStyle" => 10,
            "rFonts" => 20,
            "b" => 30,
            "bCs" => 40,
            "i" => 50,
            "iCs" => 60,
            "caps" => 70,
            "smallCaps" => 80,
            "strike" => 90,
            "dstrike" => 100,
            "outline" => 110,
            "shadow" => 120,
            "emboss" => 130,
            "imprint" => 140,
            "noProof" => 150,
            "snapToGrid" => 160,
            "vanish" => 170,
            "webHidden" => 180,
            "color" => 190,
            "spacing" => 200,
            "w" => 210,
            "kern" => 220,
            "position" => 230,
            "sz" => 240,
            "szCs" => 250,
            "highlight" => 260,
            "u" => 270,
            "effect" => 280,
            "bdr" => 290,
            "shd" => 300,
            "fitText" => 310,
            "vertAlign" => 320,
            "rtl" => 330,
            "cs" => 340,
            "em" => 350,
            "lang" => 360,
            "eastAsianLayout" => 370,
            "specVanish" => 380,
            "oMath" => 390,
            "rPrChange" => 400,
            _ => 1000
        };

        var targetRank = Rank(property.Name.LocalName);
        var next = rPr.Elements().FirstOrDefault(e => Rank(e.Name.LocalName) > targetRank);
        if (next is null)
            rPr.Add(property);
        else
            next.AddBeforeSelf(property);
    }

    private static bool DisableAutomaticFieldUpdateOnOpen(XDocument settings)
    {
        var root = settings.Root;
        if (root is null) return false;

        var updateFields = root.Element(W + "updateFields");
        if (updateFields is null) return false;

        var current = ((string?)updateFields.Attribute(W + "val") ?? string.Empty).Trim();
        if (current.Equals("0", StringComparison.OrdinalIgnoreCase) ||
            current.Equals("false", StringComparison.OrdinalIgnoreCase) ||
            current.Equals("off", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        updateFields.SetAttributeValue(W + "val", "false");
        return true;
    }



    private static IReadOnlyDictionary<int, XElement> BuildWordParagraphIndexMap(
        XDocument document)
    {
        var result = new Dictionary<int, XElement>();
        var index = 0;

        void Visit(XElement element)
        {
            if (element.Name == W + "p")
            {
                index++;
                result[index] = element;
                return;
            }

            foreach (var child in element.Elements())
                Visit(child);

            // Microsoft Word's main-story Paragraphs collection exposes one additional
            // paragraph marker for each table row. It has no standalone w:p node.
            if (element.Name == W + "tr")
                index++;
        }

        if (document.Root is null)
            throw new InvalidOperationException("word/document.xml has no root element.");

        Visit(document.Root);
        return result;
    }

    private static void GuardStructuralParagraphIndexContext(
        XElement paragraph,
        ControlledFixAction action)
    {
        var insideToc = IsInsideStandardTocContentControl(paragraph);
        var insideTable = paragraph.Ancestors(W + "tbl").Any();

        if (IsTocRole(action.Role))
        {
            if (!insideToc)
                throw new InvalidOperationException(
                    $"Candidate P{action.ParagraphIndex} ({action.Role}) structural index is outside the standard TOC content control.");
            return;
        }

        if (action.Role.Equals("TableBody", StringComparison.OrdinalIgnoreCase)
            || action.Role.Equals("TableHeader", StringComparison.OrdinalIgnoreCase))
        {
            if (!insideTable || insideToc)
                throw new InvalidOperationException(
                    $"Candidate P{action.ParagraphIndex} ({action.Role}) structural index is not in the expected table context.");
            return;
        }

        if (action.Role.Equals("TableNote", StringComparison.OrdinalIgnoreCase))
        {
            var followsTableBoundary = IsFirstNonEmptyParagraphAfterTable(paragraph);
            if ((!insideTable && !followsTableBoundary) || insideToc)
                throw new InvalidOperationException(
                    $"Candidate P{action.ParagraphIndex} ({action.Role}) structural index is neither inside a table nor the first non-empty paragraph after a table.");
            return;
        }

        if (insideToc)
        {
            throw new InvalidOperationException(
                $"Candidate P{action.ParagraphIndex} ({action.Role}) structural index points inside the TOC but the semantic role is not a TOC role.");
        }
    }

    private static bool IsFirstNonEmptyParagraphAfterTable(XElement paragraph)
    {
        foreach (var sibling in paragraph.ElementsBeforeSelf().Reverse())
        {
            if (sibling.Name == W + "tbl")
                return true;

            if (sibling.Name == W + "p")
            {
                if (HasMeaningfulParagraphContent(sibling))
                    return false;
                continue;
            }

            if (sibling.Name == W + "sdt")
                return false;
        }

        return false;
    }

    private static bool HasMeaningfulParagraphContent(XElement paragraph)
    {
        if (paragraph.Descendants(W + "t")
            .Any(node => !string.IsNullOrWhiteSpace(node.Value)))
            return true;

        if (paragraph.Descendants(W + "instrText")
            .Any(node => !string.IsNullOrWhiteSpace(node.Value)))
            return true;

        return paragraph.Descendants(W + "fldChar").Any()
               || paragraph.Descendants(W + "drawing").Any()
               || paragraph.Descendants(W + "object").Any()
               || paragraph.Descendants(W + "pict").Any();
    }

    private static bool IsInsideStandardTocContentControl(XElement paragraph)
    {
        var ancestors = paragraph.Ancestors(W + "sdt").ToList();
        if (ancestors.Any(IsRecognizedAutomaticTocSdt))
            return true;

        // When aggressive style clearing removed the standard TOC SDT wrapper/metadata,
        // an automatic TOC entry still has strong structural field signatures.
        return HasAutomaticTocEntrySignature(paragraph);
    }

    private static bool IsRecognizedAutomaticTocSdt(XElement tocSdt)
    {
        var gallery = ((string?)tocSdt
                .Element(W + "sdtPr")?
                .Element(W + "docPartObj")?
                .Element(W + "docPartGallery")?
                .Attribute(W + "val") ?? string.Empty)
            .Trim();

        if (gallery.Equals("Table of Contents", StringComparison.OrdinalIgnoreCase))
            return true;

        var instruction = string.Join(
            " ",
            tocSdt.Descendants(W + "instrText")
                .Select(x => ((string?)x ?? string.Empty).Trim()));

        if (instruction.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)
            && (instruction.Length == 3 || char.IsWhiteSpace(instruction[3])))
            return true;

        return tocSdt.Descendants(W + "p").Any(HasAutomaticTocEntrySignature);
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

    private static XElement ResolveParagraph(
        IReadOnlyList<XElement> paragraphs,
        OpenXmlStyleCatalogLite catalog,
        ControlledFixAction action,
        HashSet<XElement> used,
        IReadOnlyDictionary<int, XElement> wordParagraphIndexMap)
    {
        var preview = NormalizePreview(action.TextPreview);
        if (preview.Length == 0)
            throw new InvalidOperationException($"Candidate P{action.ParagraphIndex} has an empty semantic text key.");

        // R8R12R2 semantic-reconstruction identity authority:
        // Word's Document.Paragraphs ordinal differs from raw w:p ordinal because Word
        // exposes one additional paragraph marker at the end of each table row. Rebuild
        // that deterministic ordinal from document.xml and use it as the primary identity.
        //
        // This is intentionally stronger than damaged style/text matching: under
        // HARD_COLLAPSE hundreds of unrelated paragraphs can all have Style=Normal and
        // repeated visible text. Current formatting is not authority for target selection.
        if (wordParagraphIndexMap.TryGetValue(action.ParagraphIndex, out var indexedParagraph))
        {
            if (used.Contains(indexedParagraph))
                throw new InvalidOperationException(
                    $"Candidate P{action.ParagraphIndex} ({action.Role}) structural paragraph index was already consumed by another action.");

            GuardStructuralParagraphIndexContext(indexedParagraph, action);
            return indexedParagraph;
        }

        // QH01-E213:
        // TOC page numbers are volatile field results. Word COM can expose an in-memory page number
        // that differs from the persisted document.xml result even though the entry title is unchanged.
        // For Toc1/Toc2 only, resolve by source style + stable title before the first tab and require
        // the remaining preview suffix to be page-number digits only.
        if (IsTocRole(action.Role))
        {
            var tocMatches = paragraphs
                .Where(p => !used.Contains(p))
                .Select(p => new
                {
                    Paragraph = p,
                    Title = TocStableTitle(p),
                    Style = catalog.GetParagraphStyleName(p)
                })
                .Where(x => TokensEqual(x.Style, action.SourceStyle)
                            && x.Title.Length > 0
                            && TocPreviewMatchesStableTitle(preview, x.Title))
                .ToList();

            if (tocMatches.Count == 1)
                return tocMatches[0].Paragraph;

            if (tocMatches.Count > 1)
            {
                throw new InvalidOperationException(
                    $"Candidate P{action.ParagraphIndex} ({action.Role}) is ambiguous in document.xml: {tocMatches.Count} TOC paragraphs match stable title and source style '{action.SourceStyle}'.");
            }

            // QH01-E244:
            // A Word COM audit can expose a freshly recalculated automatic-TOC result while the persisted
            // document.xml still contains the previous TOC title/page result. The old resolver assumed only
            // page-number volatility; real evidence proved that heading titles can also be stale on disk.
            //
            // Fail-closed fallback: when the stable title does not match, bind only by a unique hierarchical
            // TOC outline key (for example 17- or 7-2-) + exact source style. This key is derived from the
            // visible numbered heading prefix and is independent of the volatile/stale title and page number.
            var previewOutlineKey = TocOutlineKey(preview);
            if (previewOutlineKey.Length > 0)
            {
                var outlineMatches = paragraphs
                    .Where(p => !used.Contains(p))
                    .Select(p => new
                    {
                        Paragraph = p,
                        Title = TocStableTitle(p),
                        Style = catalog.GetParagraphStyleName(p)
                    })
                    .Where(x => TokensEqual(x.Style, action.SourceStyle)
                                && x.Title.Length > 0
                                && TocOutlineKey(x.Title).Equals(previewOutlineKey, StringComparison.Ordinal))
                    .ToList();

                if (outlineMatches.Count == 1)
                    return outlineMatches[0].Paragraph;

                if (outlineMatches.Count > 1)
                {
                    throw new InvalidOperationException(
                        $"Candidate P{action.ParagraphIndex} ({action.Role}) is ambiguous in document.xml: {outlineMatches.Count} TOC paragraphs match outline key '{previewOutlineKey}' and source style '{action.SourceStyle}'.");
                }
            }

            throw new InvalidOperationException(
                $"Candidate P{action.ParagraphIndex} ({action.Role}) could not be resolved by stable TOC title or unique outline key in document.xml; source style='{action.SourceStyle}'; outline key='{previewOutlineKey}'.");
        }

        var textMatches = paragraphs
            .Where(p => !used.Contains(p))
            .Select(p => new { Paragraph = p, Text = NormalizeText(ParagraphText(p)), Style = catalog.GetParagraphStyleName(p) })
            .Where(x => x.Text.Length > 0 && TextMatches(x.Text, preview))
            .ToList();

        var exactStyleMatches = textMatches
            .Where(x => TokensEqual(x.Style, action.SourceStyle))
            .ToList();

        if (exactStyleMatches.Count == 1)
            return exactStyleMatches[0].Paragraph;

        if (exactStyleMatches.Count > 1)
        {
            throw new InvalidOperationException(
                $"Candidate P{action.ParagraphIndex} ({action.Role}) is ambiguous in document.xml: {exactStyleMatches.Count} paragraphs match both text and source style '{action.SourceStyle}'.");
        }

        // A semantic style-loss action can safely fall back to a unique text match only when the
        // action itself proves the style is wrong.
        if (action.Action == "ASSIGN_STYLE_AND_RESET_DIRECT_FORMAT" && textMatches.Count == 1)
            return textMatches[0].Paragraph;

        var foundStyles = string.Join(", ", textMatches.Select(x => x.Style).Distinct(StringComparer.OrdinalIgnoreCase).Take(8));
        throw new InvalidOperationException(
            $"Candidate P{action.ParagraphIndex} ({action.Role}) could not be resolved uniquely in document.xml. " +
            $"TextMatches={textMatches.Count}; source style='{action.SourceStyle}'; matched styles=[{foundStyles}].");
    }

    private static void GuardParagraph(XElement paragraph, ControlledFixAction action)
    {
        var hasFieldMarkup = ParagraphHasFieldMarkup(paragraph);

        var safeTocField =
            hasFieldMarkup
            && IsTocRole(action.Role)
            && HasOnlySafeOuterTocBoundaryMarkup(paragraph);

        var safeCaptionSequenceField =
            hasFieldMarkup
            && action.Role.Equals("FigureTableCaption", StringComparison.OrdinalIgnoreCase)
            && HasOnlySafeCaptionSequenceFieldMarkup(paragraph);

        var safePageField =
            hasFieldMarkup
            && CanonicalizeControlledRoleAliasFamily(action.Role).Equals("Body", StringComparison.OrdinalIgnoreCase)
            && HasOnlySafeCanonicalPageFieldMarkup(paragraph);

        if (hasFieldMarkup && !safeTocField && !safeCaptionSequenceField && !safePageField)
            throw new InvalidOperationException(
                $"Candidate P{action.ParagraphIndex} contains unsupported Word field markup; " +
                $"only built-in TOC boundary fields, one structurally valid SEQ caption field, " +
                $"or one canonical complex PAGE field in a Body paragraph are permitted. " +
                $"FIELD_SIGNATURE={DescribeFieldMarkup(paragraph)}");

        var intersectsContentControl =
            paragraph.Ancestors(W + "sdt").Any()
            || paragraph.Descendants(W + "sdt").Any();

        if (intersectsContentControl && !IsSafeTocContentControlContext(paragraph, action))
            throw new InvalidOperationException($"Candidate P{action.ParagraphIndex} intersects a content control outside the narrowly permitted built-in TOC context; OpenXML mutation refused.");

        if (paragraph.Descendants(W + "ins").Any() || paragraph.Descendants(W + "del").Any())
            throw new InvalidOperationException($"Candidate P{action.ParagraphIndex} contains tracked-change markup; OpenXML mutation refused.");
    }

    private static bool IsTocRole(string role)
    {
        var canonical = CanonicalizeControlledRoleAliasFamily(role);
        return canonical.Equals("Toc1", StringComparison.OrdinalIgnoreCase)
               || canonical.Equals("Toc2", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ParagraphHasFieldMarkup(XElement paragraph)
        => paragraph.Descendants(W + "fldChar").Any()
           || paragraph.Descendants(W + "instrText").Any()
           || paragraph.Descendants(W + "fldSimple").Any();

    private static bool HasOnlySafeCanonicalPageFieldMarkup(XElement paragraph)
    {
        // R10R45: formatting-only controlled mutation may cross a standard PAGE field
        // only when its field structure is the single canonical complex form. The
        // transaction-level invariants below still require field instructions and
        // fldChar structure to remain byte-for-byte semantically unchanged.
        if (paragraph.Descendants(W + "fldSimple").Any())
            return false;

        var fieldChars = paragraph.Descendants(W + "fldChar").ToList();
        var instructionNodes = paragraph.Descendants(W + "instrText").ToList();
        if (fieldChars.Count != 3 || instructionNodes.Count == 0)
            return false;

        var types = fieldChars
            .Select(x => ((string?)x.Attribute(W + "fldCharType") ?? string.Empty).Trim())
            .ToArray();
        if (!types.SequenceEqual(new[] { "begin", "separate", "end" }, StringComparer.Ordinal))
            return false;

        var instruction = NormalizeFieldInstruction(
            string.Join(" ", instructionNodes.Select(x => (string?)x ?? string.Empty)));
        return instruction.Equals("PAGE", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasOnlySafeCaptionSequenceFieldMarkup(XElement paragraph)
    {
        var simpleFields = paragraph.Descendants(W + "fldSimple").ToList();
        var fieldChars = paragraph.Descendants(W + "fldChar").ToList();
        var instructionNodes = paragraph.Descendants(W + "instrText").ToList();

        // Never accept a mixture of simple-field and complex-field representations.
        if (simpleFields.Count > 0)
        {
            if (simpleFields.Count != 1 || fieldChars.Count != 0 || instructionNodes.Count != 0)
                return false;

            var simpleInstruction = ((string?)simpleFields[0].Attribute(W + "instr") ?? string.Empty).Trim();
            return IsSafeCaptionSequenceInstruction(simpleInstruction);
        }

        // Complex Word field: exactly one begin/end pair, at most one separator,
        // no unknown field-char token, and one SEQ instruction after concatenation.
        if (fieldChars.Count == 0 || instructionNodes.Count == 0)
            return false;

        var begin = 0;
        var separate = 0;
        var end = 0;

        foreach (var fieldChar in fieldChars)
        {
            var type = ((string?)fieldChar.Attribute(W + "fldCharType") ?? string.Empty).Trim();
            switch (type)
            {
                case "begin":
                    begin++;
                    break;
                case "separate":
                    separate++;
                    break;
                case "end":
                    end++;
                    break;
                default:
                    return false;
            }
        }

        if (begin != 1 || end != 1 || separate > 1)
            return false;

        var complexInstruction = NormalizeFieldInstruction(
            string.Join(" ", instructionNodes.Select(x => (string?)x ?? string.Empty)));

        return IsSafeCaptionSequenceInstruction(complexInstruction);
    }

    private static bool IsSafeCaptionSequenceInstruction(string? instruction)
    {
        var normalized = NormalizeFieldInstruction(instruction ?? string.Empty);
        if (normalized.Length == 0)
            return false;

        // Only the Word SEQ field family is permitted here. This covers ordinary
        // Figure/Table caption numbering, including chapter-number switches.
        if (!System.Text.RegularExpressions.Regex.IsMatch(
                normalized,
                @"^SEQ(?:\s|$)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant))
            return false;

        // Fail closed if a second active/dynamic field family appears in the same
        // instruction string. A normal SEQ instruction may contain switches only.
        if (System.Text.RegularExpressions.Regex.IsMatch(
                normalized,
                @"\b(?:DDE|DDEAUTO|INCLUDETEXT|INCLUDEPICTURE|LINK|MACROBUTTON|AUTOTEXT|AUTOTEXTLIST|HYPERLINK)\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant))
            return false;

        return true;
    }

    private static string NormalizeFieldInstruction(string value)
        => System.Text.RegularExpressions.Regex.Replace(value ?? string.Empty, @"\s+", " ").Trim();

    private static string DescribeFieldMarkup(XElement paragraph)
    {
        var simple = paragraph.Descendants(W + "fldSimple")
            .Select(x => NormalizeFieldInstruction((string?)x.Attribute(W + "instr") ?? string.Empty))
            .Where(x => x.Length > 0)
            .ToArray();

        var complex = paragraph.Descendants(W + "instrText")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();

        var joined = NormalizeFieldInstruction(string.Join(" ", complex));

        var types = paragraph.Descendants(W + "fldChar")
            .Select(x => ((string?)x.Attribute(W + "fldCharType") ?? string.Empty).Trim())
            .ToArray();

        return $"simple=[{string.Join("|", simple)}];complex=[{joined}];fldChar=[{string.Join(",", types)}]";
    }

    private static bool IsSafeTocContentControlContext(XElement paragraph, ControlledFixAction action)
    {
        if (!IsTocRole(action.Role))
            return false;

        // The paragraph itself may not contain a nested content control.
        if (paragraph.Descendants(W + "sdt").Any())
            return false;

        var ancestors = paragraph.Ancestors(W + "sdt").ToList();
        if (ancestors.Count != 1)
            return false;

        var tocSdt = ancestors[0];
        var gallery = ((string?)tocSdt
                .Element(W + "sdtPr")?
                .Element(W + "docPartObj")?
                .Element(W + "docPartGallery")?
                .Attribute(W + "val") ?? string.Empty)
            .Trim();

        if (!gallery.Equals("Table of Contents", StringComparison.OrdinalIgnoreCase))
            return false;

        // Fail closed for nested SDTs or tracked changes inside the built-in TOC container.
        if (tocSdt.Descendants(W + "sdt").Any())
            return false;
        if (tocSdt.Descendants(W + "ins").Any() || tocSdt.Descendants(W + "del").Any())
            return false;

        return true;
    }

    private static string TocStableTitle(XElement paragraph)
    {
        // QH01-E213R1:
        // Search only run content. Paragraph properties can contain <w:tabs><w:tab .../></w:tabs>
        // before the first text run; that tab-stop definition is not the TOC title/page separator.
        var pieces = new List<string>();
        foreach (var run in paragraph.Descendants(W + "r"))
        {
            foreach (var child in run.Elements())
            {
                if (child.Name == W + "tab")
                    return NormalizeText(string.Concat(pieces));

                if (child.Name == W + "t")
                {
                    pieces.Add((string?)child ?? string.Empty);
                    continue;
                }

                foreach (var text in child.Descendants(W + "t"))
                    pieces.Add((string?)text ?? string.Empty);
            }
        }

        return NormalizeText(string.Concat(pieces));
    }

    private static string TocOutlineKey(string value)
    {
        var normalized = NormalizeText(value);
        if (normalized.Length == 0)
            return string.Empty;

        var sb = new StringBuilder();
        var sawDigit = false;
        var expectDigitAfterSeparator = false;

        foreach (var c in normalized)
        {
            var digit = NormalizeTocOutlineDigit(c);
            if (digit != '\0')
            {
                sb.Append(digit);
                sawDigit = true;
                expectDigitAfterSeparator = false;
                continue;
            }

            if (c is '-' or '\u2010' or '\u2011' or '\u2012' or '\u2013' or '\u2014' or '\u2212')
            {
                if (!sawDigit)
                    return string.Empty;

                sb.Append('-');
                expectDigitAfterSeparator = true;
                continue;
            }

            if (char.IsWhiteSpace(c))
                continue;

            break;
        }

        if (!sawDigit)
            return string.Empty;

        // A numbered TOC key must end in a separator (e.g. 17- or 7-2-). Requiring the separator
        // prevents accidental binding to ordinary prose that merely begins with a number.
        if (sb.Length == 0 || sb[^1] != '-' || !expectDigitAfterSeparator)
            return string.Empty;

        return sb.ToString();
    }

    private static char NormalizeTocOutlineDigit(char c)
    {
        if (c is >= '0' and <= '9')
            return c;
        if (c is >= '\u06F0' and <= '\u06F9')
            return (char)('0' + (c - '\u06F0'));
        if (c is >= '\u0660' and <= '\u0669')
            return (char)('0' + (c - '\u0660'));
        return '\0';
    }

    private static bool TocPreviewMatchesStableTitle(string normalizedPreview, string stableTitle)
    {
        var title = NormalizeText(stableTitle);
        if (title.Length == 0) return false;
        if (normalizedPreview.Equals(title, StringComparison.Ordinal)) return true;
        if (!normalizedPreview.StartsWith(title, StringComparison.Ordinal)) return false;

        var suffix = normalizedPreview[title.Length..].Trim();
        if (suffix.Length == 0) return false;

        foreach (var c in suffix)
        {
            var digit =
                (c >= '0' && c <= '9')
                || (c >= '\u06F0' && c <= '\u06F9')
                || (c >= '\u0660' && c <= '\u0669');
            if (!digit) return false;
        }
        return true;
    }

    private static bool HasOnlySafeOuterTocBoundaryMarkup(XElement paragraph)
    {
        if (paragraph.Descendants(W + "fldSimple").Any())
            return false;

        var instructions = paragraph.Descendants(W + "instrText")
            .Select(x => NormalizeText((string?)x))
            .Where(x => x.Length > 0)
            .ToList();

        // QH01-E287: Word-native TOC result paragraphs commonly contain nested PAGEREF
        // fields for the displayed page number. These are part of the built-in TOC field
        // result and are safe for our paragraph-format-only mutation. Keep this narrow:
        // only TOC instructions and the canonical _Toc... PAGEREF form are admitted.
        foreach (var instruction in instructions)
        {
            if (instruction.StartsWith("TOC ", StringComparison.OrdinalIgnoreCase))
                continue;

            if (System.Text.RegularExpressions.Regex.IsMatch(
                    instruction,
                    @"^PAGEREF\s+_Toc[0-9A-Za-z_]+(?:\s+\\[A-Za-z]+)*$",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant))
                continue;

            return false;
        }

        foreach (var fieldChar in paragraph.Descendants(W + "fldChar"))
        {
            var type = ((string?)fieldChar.Attribute(W + "fldCharType") ?? string.Empty).Trim();
            if (type is not ("begin" or "separate" or "end"))
                return false;
        }

        return true;
    }

    private static void ApplyOpenXmlAction(
        XElement paragraph,
        OpenXmlStyleCatalogLite catalog,
        ControlledFixAction action,
        RoleDefaults expected)
    {
        var pPr = EnsurePPr(paragraph);

        if (action.Action == "ASSIGN_STYLE_AND_RESET_DIRECT_FORMAT")
        {
            if (string.IsNullOrWhiteSpace(action.TargetStyle))
                throw new InvalidOperationException($"Candidate P{action.ParagraphIndex} has no target style.");

            var styleId = catalog.ResolveStyleId(action.TargetStyle!);
            if (styleId is null)
                throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: target style '{action.TargetStyle}' was not found in styles.xml.");

            var pStyle = pPr.Element(W + "pStyle");
            if (pStyle is null)
            {
                pStyle = new XElement(W + "pStyle");
                pPr.AddFirst(pStyle);
            }
            pStyle.SetAttributeValue(W + "val", styleId);
        }
        else if (action.Action != "RESET_DIRECT_FORMAT_TO_STYLE" &&
                 action.Action != "APPLY_ROLE_DEFAULTS_WITH_GUARD" &&
                 action.Action != "MATERIALIZE_SEMANTIC_ROLE_RUN_DEFAULTS")
        {
            throw new InvalidOperationException($"Unsupported controlled-fix action: {action.Action}");
        }

        // Conservative direct-format reset: clear only the dimensions governed by the profile.
        // APPLY_ROLE_DEFAULTS_WITH_GUARD intentionally does not touch run font/size formatting.
        // MATERIALIZE_SEMANTIC_ROLE_RUN_DEFAULTS is intentionally different: it writes only the
        // bidirectional font/size dimensions for one already-proven semantic paragraph and does not
        // alter its shared source style or unrelated character properties.
        if (action.Action == "MATERIALIZE_SEMANTIC_ROLE_RUN_DEFAULTS")
            MaterializeSemanticRoleRunContract(paragraph, expected);
        else if (action.Action != "APPLY_ROLE_DEFAULTS_WITH_GUARD")
            ClearFontAndSizeOverrides(paragraph);

        // Materialize only paragraph-level alignment / bidi contract.
        MaterializeExistingParagraphContract(pPr, expected);
    }

    private static void MaterializeSemanticRoleRunContract(XElement paragraph, RoleDefaults expected)
    {
        if (string.IsNullOrWhiteSpace(expected.FontBi) && expected.SizeBi is null)
            return;

        var pPr = EnsurePPr(paragraph);
        var paragraphMarkRPr = pPr.Element(W + "rPr");
        if (paragraphMarkRPr is null)
        {
            paragraphMarkRPr = new XElement(W + "rPr");
            InsertParagraphPropertyInSchemaOrder(pPr, paragraphMarkRPr);
        }
        MaterializeBiFontSizeOnRunProperties(paragraphMarkRPr, expected);

        // QH01-E223:
        // Word's paragraph Range.Font may become mixed when any run in the semantic paragraph
        // remains on a different font/size authority, even if that run only carries tabs,
        // breaks, drawings, proofing/field scaffolding, or other non-text range content.
        // The semantic role contract governs the paragraph as a whole, so materialize only
        // FontBi/SizeBi on every run in this one semantic paragraph. Text, fields, drawings,
        // bold/italic and all unrelated run properties remain unchanged.
        var allRuns = paragraph.Descendants(W + "r").ToList();
        if (allRuns.Count == 0)
            throw new InvalidOperationException("Semantic role run materialization requires at least one run.");

        foreach (var run in allRuns)
        {
            var rPr = run.Element(W + "rPr");
            if (rPr is null)
            {
                rPr = new XElement(W + "rPr");
                run.AddFirst(rPr);
            }
            MaterializeBiFontSizeOnRunProperties(rPr, expected);
        }
    }

    private static void MaterializeBiFontSizeOnRunProperties(XElement rPr, RoleDefaults expected)
    {
        if (!string.IsNullOrWhiteSpace(expected.FontBi))
        {
            var rFonts = rPr.Element(W + "rFonts");
            if (rFonts is null)
            {
                rFonts = new XElement(W + "rFonts");
                InsertRunPropertyInSchemaOrder(rPr, rFonts);
            }
            rFonts.SetAttributeValue(W + "cs", expected.FontBi.Trim());
            rFonts.Attribute(W + "cstheme")?.Remove();
        }

        if (expected.SizeBi is not null)
        {
            var halfPoints = (int)Math.Round(expected.SizeBi.Value * 2f, MidpointRounding.AwayFromZero);
            var expectedToken = halfPoints.ToString(CultureInfo.InvariantCulture);
            var szCs = rPr.Element(W + "szCs");
            if (szCs is null)
            {
                szCs = new XElement(W + "szCs");
                InsertRunPropertyInSchemaOrder(rPr, szCs);
            }
            szCs.SetAttributeValue(W + "val", expectedToken);
        }
    }

    private static void ClearFontAndSizeOverrides(XElement paragraph)
    {
        var pPrRPr = paragraph.Element(W + "pPr")?.Element(W + "rPr");
        if (pPrRPr is not null)
            RemoveFormattingChildren(pPrRPr);

        foreach (var rPr in paragraph.Descendants(W + "rPr").ToList())
            RemoveFormattingChildren(rPr);
    }

    private static void RemoveFormattingChildren(XElement rPr)
    {
        rPr.Elements(W + "rFonts").Remove();
        rPr.Elements(W + "sz").Remove();
        rPr.Elements(W + "szCs").Remove();
    }

    private static void MaterializeExistingParagraphContract(XElement pPr, RoleDefaults expected)
    {
        // Patch 0015: do not rely on Word to materialize paragraph-format inheritance.
        // For controlled candidates, write the role contract explicitly into w:pPr so that
        // effective Alignment/ReadingOrder is deterministic when Word opens the OOXML package.
        if (!string.IsNullOrWhiteSpace(expected.Alignment))
        {
            // Microsoft Word compatibility behavior for bidi paragraphs is not the same as
            // treating w:jc left/right as physical page edges. In Word, when w:bidi is on,
            // left/right are evaluated in the RTL paragraph context. Therefore a physical
            // RIGHT contract must be serialized as w:jc=left for an RTL paragraph, and
            // physical LEFT as w:jc=right. CENTER/JUSTIFY/DISTRIBUTE are unchanged.
            var value = EncodeAlignmentForWordOoxml(expected.Alignment, expected.ReadingOrder);

            var jc = pPr.Element(W + "jc");
            if (jc is null)
            {
                jc = new XElement(W + "jc");
                InsertParagraphPropertyInSchemaOrder(pPr, jc);
            }
            jc.SetAttributeValue(W + "val", value);
        }

        if (!string.IsNullOrWhiteSpace(expected.ReadingOrder))
        {
            var rtl = expected.ReadingOrder.Trim().Equals("RTL", StringComparison.OrdinalIgnoreCase);
            var ltr = expected.ReadingOrder.Trim().Equals("LTR", StringComparison.OrdinalIgnoreCase);
            if (!rtl && !ltr)
                throw new InvalidOperationException($"Unsupported ReadingOrder contract: {expected.ReadingOrder}");

            var bidi = pPr.Element(W + "bidi");
            if (bidi is null)
            {
                bidi = new XElement(W + "bidi");
                InsertParagraphPropertyInSchemaOrder(pPr, bidi);
            }
            bidi.SetAttributeValue(W + "val", rtl ? "1" : "0");
        }
    }

    private static string EncodeAlignmentForWordOoxml(string alignment, string? readingOrder)
    {
        var token = alignment.Trim().ToUpperInvariant();
        var rtl = string.Equals(readingOrder?.Trim(), "RTL", StringComparison.OrdinalIgnoreCase);

        return token switch
        {
            // Word evaluates left/right through bidi paragraph layout. The inversion here is
            // deliberate and is required for Word COM to report the intended physical edge.
            "RIGHT" when rtl => "left",
            "LEFT" when rtl => "right",
            "RIGHT" => "right",
            "LEFT" => "left",
            "CENTER" => "center",
            "JUSTIFY" => "both",
            "DISTRIBUTE" => "distribute",
            _ => throw new InvalidOperationException($"Unsupported Alignment contract: {alignment}")
        };
    }

    private static void InsertParagraphPropertyInSchemaOrder(XElement pPr, XElement property)
    {
        // WordprocessingML pPr is order-sensitive. Keep the insertion order aligned with
        // CT_PPr so the existing OpenXML validation gate does not gain new schema errors.
        static int Rank(string name) => name switch
        {
            "pStyle" => 10,
            "keepNext" => 20,
            "keepLines" => 30,
            "pageBreakBefore" => 40,
            "framePr" => 50,
            "widowControl" => 60,
            "numPr" => 70,
            "suppressLineNumbers" => 80,
            "pBdr" => 90,
            "shd" => 100,
            "tabs" => 110,
            "suppressAutoHyphens" => 120,
            "kinsoku" => 130,
            "wordWrap" => 140,
            "overflowPunct" => 150,
            "topLinePunct" => 160,
            "autoSpaceDE" => 170,
            "autoSpaceDN" => 180,
            "bidi" => 190,
            "adjustRightInd" => 200,
            "snapToGrid" => 210,
            "spacing" => 220,
            "ind" => 230,
            "contextualSpacing" => 240,
            "mirrorIndents" => 250,
            "suppressOverlap" => 260,
            "jc" => 270,
            "textDirection" => 280,
            "textAlignment" => 290,
            "textboxTightWrap" => 300,
            "outlineLvl" => 310,
            "divId" => 320,
            "cnfStyle" => 330,
            "rPr" => 340,
            "sectPr" => 350,
            "pPrChange" => 360,
            _ => 1000
        };

        var targetRank = Rank(property.Name.LocalName);
        var next = pPr.Elements().FirstOrDefault(e => Rank(e.Name.LocalName) > targetRank);
        if (next is null)
            pPr.Add(property);
        else
            next.AddBeforeSelf(property);
    }

    private static void VerifyOpenXmlAction(
        XElement paragraph,
        OpenXmlStyleCatalogLite catalog,
        ControlledFixAction action,
        RoleDefaults expected)
    {
        if (action.Action == "ASSIGN_STYLE_AND_RESET_DIRECT_FORMAT")
        {
            var actualStyle = catalog.GetParagraphStyleName(paragraph);
            if (string.IsNullOrWhiteSpace(action.TargetStyle) || !TokensEqual(actualStyle, action.TargetStyle))
                throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: style assignment verification failed. Actual='{actualStyle}', expected='{action.TargetStyle}'.");
        }

        if (action.Action == "RESET_DIRECT_FORMAT_TO_STYLE")
        {
            var actualStyle = catalog.GetParagraphStyleName(paragraph);
            if (string.IsNullOrWhiteSpace(action.SourceStyle)
                || !TokensEqual(actualStyle, action.SourceStyle))
            {
                throw new InvalidOperationException(
                    $"Candidate P{action.ParagraphIndex}: style identity changed during direct-format reset. Actual='{actualStyle}', expected existing='{action.SourceStyle}'.");
            }
        }

        if (action.Action == "MATERIALIZE_SEMANTIC_ROLE_RUN_DEFAULTS")
        {
            VerifySemanticRoleRunContract(paragraph, action, expected);
        }
        else if (action.Action != "APPLY_ROLE_DEFAULTS_WITH_GUARD" &&
            (!string.IsNullOrWhiteSpace(expected.FontBi) || expected.SizeBi is not null))
        {
            var conflicting = paragraph.Descendants(W + "rPr")
                .Any(r => r.Element(W + "rFonts") is not null || r.Element(W + "sz") is not null || r.Element(W + "szCs") is not null);
            var pPrRPr = paragraph.Element(W + "pPr")?.Element(W + "rPr");
            conflicting |= pPrRPr is not null
                && (pPrRPr.Element(W + "rFonts") is not null || pPrRPr.Element(W + "sz") is not null || pPrRPr.Element(W + "szCs") is not null);
            if (conflicting)
                throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: font/size direct-format reset verification failed.");
        }

        var pPr = paragraph.Element(W + "pPr");
        if (pPr is null)
            throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: paragraph properties are missing after controlled mutation.");

        if (!string.IsNullOrWhiteSpace(expected.Alignment))
        {
            var jc = pPr.Element(W + "jc")
                ?? throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: explicit alignment contract was not materialized.");
            var actual = ((string?)jc.Attribute(W + "val") ?? string.Empty).Trim().ToUpperInvariant();
            var expectedToken = EncodeAlignmentForWordOoxml(expected.Alignment, expected.ReadingOrder).ToUpperInvariant();
            if (actual != expectedToken)
                throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: Word-compatible alignment materialization verification failed ({actual}/{expectedToken}).");
        }

        if (!string.IsNullOrWhiteSpace(expected.ReadingOrder))
        {
            var bidi = pPr.Element(W + "bidi")
                ?? throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: explicit reading-order contract was not materialized.");
            var actual = ((string?)bidi.Attribute(W + "val") ?? string.Empty).Trim();
            var expectedValue = expected.ReadingOrder.Trim().Equals("RTL", StringComparison.OrdinalIgnoreCase) ? "1" : "0";
            if (actual != expectedValue)
                throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: reading-order materialization verification failed ({actual}/{expectedValue}).");
        }
    }

    private static void VerifySemanticRoleRunContract(
        XElement paragraph,
        ControlledFixAction action,
        RoleDefaults expected)
    {
        var targets = new List<XElement>();
        var pPrRPr = paragraph.Element(W + "pPr")?.Element(W + "rPr");
        if (pPrRPr is not null)
            targets.Add(pPrRPr);
        targets.AddRange(paragraph.Descendants(W + "r")
            .Select(run => run.Element(W + "rPr"))
            .OfType<XElement>());

        if (targets.Count < 2)
            throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: semantic run-contract verification surface is incomplete.");

        var expectedSizeToken = expected.SizeBi is null
            ? null
            : ((int)Math.Round(expected.SizeBi.Value * 2f, MidpointRounding.AwayFromZero))
                .ToString(CultureInfo.InvariantCulture);

        foreach (var rPr in targets)
        {
            if (!string.IsNullOrWhiteSpace(expected.FontBi))
            {
                var rFonts = rPr.Element(W + "rFonts");
                var actualFont = ((string?)rFonts?.Attribute(W + "cs") ?? string.Empty).Trim();
                if (!TokensEqual(actualFont, expected.FontBi))
                    throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: semantic FontBi materialization verification failed. Actual='{actualFont}', expected='{expected.FontBi}'.");
                if (rFonts?.Attribute(W + "cstheme") is not null)
                    throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: semantic FontBi still has csTheme inheritance after materialization.");
            }

            if (expectedSizeToken is not null)
            {
                var actualSize = ((string?)rPr.Element(W + "szCs")?.Attribute(W + "val") ?? string.Empty).Trim();
                if (!string.Equals(actualSize, expectedSizeToken, StringComparison.Ordinal))
                    throw new InvalidOperationException($"Candidate P{action.ParagraphIndex}: semantic SizeBi materialization verification failed ({actualSize}/{expectedSizeToken}).");
            }
        }
    }

    private static XElement EnsurePPr(XElement paragraph)
    {
        var pPr = paragraph.Element(W + "pPr");
        if (pPr is not null) return pPr;

        pPr = new XElement(W + "pPr");
        paragraph.AddFirst(pPr);
        return pPr;
    }

    private static string ParagraphText(XElement paragraph)
        => string.Concat(paragraph.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));

    private static string NormalizePreview(string preview)
    {
        var value = preview ?? string.Empty;
        if (value.EndsWith("...", StringComparison.Ordinal) && value.Length > 3)
            value = value[..^3];
        return NormalizeText(value);
    }

    private static bool TextMatches(string paragraphText, string preview)
        => paragraphText.Equals(preview, StringComparison.Ordinal)
           || paragraphText.StartsWith(preview, StringComparison.Ordinal);

    private static string NormalizeText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var normalized = value.Normalize(NormalizationForm.FormC)
            .Replace('\u064A', '\u06CC')
            .Replace('\u0643', '\u06A9')
            .Replace('\u200C', ' ')
            .Replace('\u200F', ' ')
            .Replace('\u200E', ' ');
        return string.Join(' ', normalized.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).Trim();
    }

    private static bool TokensEqual(string? a, string? b)
        => string.Equals(NormalizeText(a), NormalizeText(b), StringComparison.OrdinalIgnoreCase);

    private static XDocument LoadXmlEntry(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path) ?? throw new InvalidOperationException($"Required package part missing: {path}");
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
    }

    private static void ReplaceXmlEntry(ZipArchive archive, string path, XDocument document)
    {
        var old = archive.GetEntry(path) ?? throw new InvalidOperationException($"Required package part missing: {path}");
        var time = old.LastWriteTime;
        old.Delete();

        var replacement = archive.CreateEntry(path, CompressionLevel.Optimal);
        replacement.LastWriteTime = time;
        using var stream = replacement.Open();
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = false,
            CloseOutput = false
        };
        using var writer = XmlWriter.Create(stream, settings);
        document.Save(writer);
    }

    private sealed class OpenXmlStyleCatalogLite
    {
        private readonly Dictionary<string, string> _styleIdToName;
        private readonly Dictionary<string, string> _normalizedNameToId;
        private readonly string _defaultParagraphStyleId;
        private readonly string _defaultParagraphStyleName;

        private OpenXmlStyleCatalogLite(
            Dictionary<string, string> styleIdToName,
            Dictionary<string, string> normalizedNameToId,
            string defaultParagraphStyleId,
            string defaultParagraphStyleName)
        {
            _styleIdToName = styleIdToName;
            _normalizedNameToId = normalizedNameToId;
            _defaultParagraphStyleId = defaultParagraphStyleId;
            _defaultParagraphStyleName = defaultParagraphStyleName;
        }

        public static OpenXmlStyleCatalogLite Build(XDocument styles)
        {
            var idToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var normalizedNameToId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string? defaultId = null;
            string? defaultName = null;

            foreach (var style in styles.Descendants(W + "style"))
            {
                var id = (string?)style.Attribute(W + "styleId");
                var name = (string?)style.Element(W + "name")?.Attribute(W + "val");
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name)) continue;

                idToName[id] = name;
                var normalized = NormalizeText(name);
                if (!normalizedNameToId.ContainsKey(normalized))
                    normalizedNameToId[normalized] = id;

                var type = (string?)style.Attribute(W + "type");
                var isDefault = IsOn(style.Attribute(W + "default"));
                if (isDefault && string.Equals(type, "paragraph", StringComparison.OrdinalIgnoreCase))
                {
                    defaultId = id;
                    defaultName = name;
                }
            }

            defaultId ??= idToName.ContainsKey("Normal") ? "Normal" : idToName.Keys.FirstOrDefault() ?? "Normal";
            defaultName ??= idToName.TryGetValue(defaultId, out var n) ? n : "Normal";
            if (!normalizedNameToId.ContainsKey(NormalizeText(defaultName)))
                normalizedNameToId[NormalizeText(defaultName)] = defaultId;

            return new OpenXmlStyleCatalogLite(idToName, normalizedNameToId, defaultId, defaultName);
        }

        public string GetParagraphStyleName(XElement paragraph)
        {
            var styleId = (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
            if (string.IsNullOrWhiteSpace(styleId)) return _defaultParagraphStyleName;
            return _styleIdToName.TryGetValue(styleId, out var name) ? name : styleId;
        }

        public string? ResolveStyleId(string styleName)
        {
            var key = NormalizeText(styleName);
            if (_normalizedNameToId.TryGetValue(key, out var id)) return id;
            if (TokensEqual(styleName, _defaultParagraphStyleName)) return _defaultParagraphStyleId;
            return null;
        }

        private static bool IsOn(XAttribute? attribute)
        {
            var value = ((string?)attribute ?? string.Empty).Trim();
            return value is "1" or "true" or "on";
        }
    }


    private static void NormalizeGovernanceGanttLeadingManualPageBreak(
        string outputPath,
        AgentProfile profile,
        List<string> warnings,
        Action<string>? trace)
    {
        const string caption =
            "جدول (۱۰-۲) گانت‌چارت نه‌ماهه پروژه سامانه تصمیم‌یار هوشمند";

        if (!string.Equals(profile.Id, "proposal.governance.fa", StringComparison.OrdinalIgnoreCase))
        {
            trace?.Invoke("GANTT_MANUAL_PAGEBREAK_PRODUCTIZATION:SKIP:PROFILE");
            return;
        }

        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var document = LoadXmlEntry(archive, "word/document.xml");

        var originalText = document.Descendants(W + "t")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();
        var originalFields = document.Descendants(W + "instrText")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();

        var globalPageBreaksBefore = document
            .Descendants(W + "br")
            .Count(x => string.Equals(
                (string?)x.Attribute(W + "type"),
                "page",
                StringComparison.OrdinalIgnoreCase));

        var matches = document
            .Descendants(W + "p")
            .Where(paragraph => TokensEqual(
                string.Concat(paragraph.Descendants(W + "t")
                    .Select(x => (string?)x ?? string.Empty)),
                caption))
            .ToArray();

        if (matches.Length == 0)
        {
            trace?.Invoke("GANTT_MANUAL_PAGEBREAK_PRODUCTIZATION:SKIP:CAPTION_NOT_FOUND");
            return;
        }

        if (matches.Length > 1)
            throw new InvalidOperationException(
                $"Gantt page-break productization found multiple caption matches. Found={matches.Length}.");

        var paragraph = matches[0];
        var firstVisibleText = paragraph
            .Descendants(W + "t")
            .FirstOrDefault(x => !string.IsNullOrWhiteSpace((string?)x));

        if (firstVisibleText is null)
            throw new InvalidOperationException(
                "Gantt caption has no visible text node.");

        var leadingPageBreaks = paragraph
            .Descendants(W + "br")
            .Where(x =>
                string.Equals(
                    (string?)x.Attribute(W + "type"),
                    "page",
                    StringComparison.OrdinalIgnoreCase)
                && XNode.DocumentOrderComparer.Compare(x, firstVisibleText) < 0)
            .ToArray();

        var leadingRenderedBreaks = paragraph
            .Descendants(W + "lastRenderedPageBreak")
            .Where(x => XNode.DocumentOrderComparer.Compare(x, firstVisibleText) < 0)
            .ToArray();

        if (leadingPageBreaks.Length == 0)
        {
            trace?.Invoke(
                $"GANTT_MANUAL_PAGEBREAK_PRODUCTIZATION:ALREADY_COMPLIANT:"
                + $"CAPTION_MATCHES=1:LEADING_MANUAL=0:"
                + $"LEADING_CACHED={leadingRenderedBreaks.Length}:"
                + $"GLOBAL_MANUAL={globalPageBreaksBefore}");
            return;
        }

        if (leadingPageBreaks.Length > 1)
            throw new InvalidOperationException(
                $"Gantt page-break productization found multiple leading manual page breaks. Found={leadingPageBreaks.Length}.");

        var breakParent = leadingPageBreaks[0].Parent;
        leadingPageBreaks[0].Remove();
        if (breakParent is not null && !breakParent.Nodes().Any())
            breakParent.Remove();

        foreach (var rendered in leadingRenderedBreaks)
            rendered.Remove();

        var globalPageBreaksAfter = document
            .Descendants(W + "br")
            .Count(x => string.Equals(
                (string?)x.Attribute(W + "type"),
                "page",
                StringComparison.OrdinalIgnoreCase));

        if (globalPageBreaksAfter != globalPageBreaksBefore - 1)
            throw new InvalidOperationException(
                $"Manual page-break delta is not exactly one. Global={globalPageBreaksBefore}->{globalPageBreaksAfter}.");

        var remainingLeadingPageBreaks = paragraph
            .Descendants(W + "br")
            .Count(x =>
                string.Equals(
                    (string?)x.Attribute(W + "type"),
                    "page",
                    StringComparison.OrdinalIgnoreCase)
                && XNode.DocumentOrderComparer.Compare(x, firstVisibleText) < 0);

        var remainingLeadingRenderedBreaks = paragraph
            .Descendants(W + "lastRenderedPageBreak")
            .Count(x => XNode.DocumentOrderComparer.Compare(x, firstVisibleText) < 0);

        if (remainingLeadingPageBreaks != 0 || remainingLeadingRenderedBreaks != 0)
            throw new InvalidOperationException(
                $"Gantt caption leading break cleanup incomplete: manual={remainingLeadingPageBreaks}, cached={remainingLeadingRenderedBreaks}.");

        var finalText = document.Descendants(W + "t")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();
        var finalFields = document.Descendants(W + "instrText")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();

        if (!originalText.SequenceEqual(finalText, StringComparer.Ordinal))
            throw new InvalidOperationException(
                "Gantt page-break productization changed visible text.");

        if (!originalFields.SequenceEqual(finalFields, StringComparer.Ordinal))
            throw new InvalidOperationException(
                "Gantt page-break productization changed field instructions.");

        ReplaceXmlEntry(archive, "word/document.xml", document);

        warnings.Add(
            "یک شکست صفحه دستیِ منحصربه‌فرد قبل از کپشن گانت در کپی کنترل‌شده حذف شد تا صفحه‌بندی حرفه‌ای B010 داخل محصول بازتولید شود؛ متن و Fieldها تغییر نکردند.");

        trace?.Invoke(
            $"GANTT_MANUAL_PAGEBREAK_PRODUCTIZATION:PASS:"
            + $"CAPTION_MATCHES=1:LEADING_MANUAL=1->0:"
            + $"LEADING_CACHED={leadingRenderedBreaks.Length}->0:"
            + $"GLOBAL_MANUAL={globalPageBreaksBefore}->{globalPageBreaksAfter}:"
            + $"VISIBLE_TEXT=UNCHANGED:FIELDS=UNCHANGED");
    }

    private static void RenormalizeTocOpenXmlAfterWordRefresh(
        string outputPath,
        AgentProfile profile,
        List<string> warnings,
        Action<string>? trace)
    {
        if (!string.Equals(profile.Id, "proposal.governance.fa", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(profile.Id, "book.university.fa", StringComparison.OrdinalIgnoreCase))
        {
            trace?.Invoke("POST_WORD_TOC_OOXML_RENORMALIZATION:SKIP:PROFILE");
            return;
        }

        trace?.Invoke("POST_WORD_TOC_OOXML_RENORMALIZATION:START");

        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var document = LoadXmlEntry(archive, "word/document.xml");
        var styles = LoadXmlEntry(archive, "word/styles.xml");
        var tocStyleTokens = ResolveAutomaticTocStyleTokens(styles, profile);

        var originalText = document.Descendants(W + "t")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();
        var originalFields = document.Descendants(W + "instrText")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();

        var pageContentWidth = ResolveMinimumPageContentWidthTwips(document);
        var changed = false;
        if (ApplyTableFitAndPagination(document, pageContentWidth))
            changed = true;

        if (ApplyAutomaticTocPhysicalRtl(document, tocStyleTokens))
            changed = true;
        var stylesChanged = false;
        if (ApplyAutomaticTocTabLayout(document, styles, pageContentWidth, tocStyleTokens))
        {
            changed = true;
            stylesChanged = true;
        }

        var post = InspectLayoutStabilityAndToc(document, pageContentWidth, tocStyleTokens);
        var tabs = InspectAutomaticTocTabLayout(document, styles, pageContentWidth, tocStyleTokens);

        var allParagraphs = document.Descendants(W + "p").ToList();
        var styleOnlyTocParagraphs = allParagraphs.Count(p =>
        {
            var styleId = ((string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") ?? string.Empty).Trim();
            var token = NormalizeAutomaticTocStyleToken(styleId);
            return token.Length > 0 && tocStyleTokens.Contains(token);
        });
        var structuralTocParagraphs = allParagraphs.Count(HasAutomaticTocEntrySignature);
        var unionTocParagraphs = allParagraphs.Count(p => IsAutomaticTocParagraph(p, tocStyleTokens));
        var discoveredStyleIds = allParagraphs
            .Where(p => IsAutomaticTocParagraph(p, tocStyleTokens))
            .Select(p => ((string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") ?? string.Empty).Trim())
            .Where(x => x.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .Take(12)
            .ToArray();

        trace?.Invoke(
            $"POST_WORD_TOC_DISCOVERY:{(unionTocParagraphs > 0 ? "PASS" : "FAIL")}:"
            + $"STYLE_ONLY={styleOnlyTocParagraphs}:STRUCTURAL={structuralTocParagraphs}:UNION={unionTocParagraphs}:"
            + $"STYLE_IDS={string.Join(",", discoveredStyleIds)}");

        if (!post.FullyCompliant)
            throw new InvalidOperationException(
                "Post-Word TOC OpenXML re-normalization layout verification failed: "
                + $"TABLES={post.Stats.Tables};ROWS={post.Stats.Rows};TOC={post.Stats.TocParagraphs};"
                + $"RIGHT={post.Stats.TocPhysicalRight};RUNRTL={post.Stats.TocTextRunsRtl};ISSUES={post.Issues.Count}; "
                + string.Join(" | ", post.Issues.Take(30)));

        if (!tabs.FullyCompliant)
            throw new InvalidOperationException(
                "Post-Word TOC OpenXML re-normalization tab verification failed: "
                + $"PARAGRAPHS={tabs.Paragraphs};PARAGRAPH_TABS_OK={tabs.ParagraphTabsOk};"
                + $"STYLES={tabs.Styles};STYLE_TABS_OK={tabs.StyleTabsOk};TAB_POS={tabs.TabPosition};ISSUES={tabs.Issues.Count}; "
                + string.Join(" | ", tabs.Issues.Take(20)));

        trace?.Invoke(
            "POST_WORD_ORDERING_GATE:PASS:LAYOUT_AND_TOC_REPAIRED_BEFORE_COMBINED_VERIFY");

        trace?.Invoke(
            $"POST_WORD_LAYOUT_RENORMALIZATION:PASS:"
            + $"TABLES={post.Stats.Tables}:ROWS={post.Stats.Rows}:"
            + $"ROWS_NOSPLIT={post.Stats.RowsNoSplit}:HEADER_REPEAT={post.Stats.HeaderRowsRepeat}:"
            + $"TABLE_KEEP_CHAIN={post.Stats.TableKeepChains}:KEEP_NEXT={post.Stats.KeepNextParagraphs}:"
            + $"CAPTIONS_KEEP={post.Stats.CaptionsKeepNext}:"
            + $"VISIBLE_TEXT=PENDING_INVARIANT:FIELDS=PENDING_INVARIANT");

        var finalText = document.Descendants(W + "t")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();
        var finalFields = document.Descendants(W + "instrText")
            .Select(x => (string?)x ?? string.Empty)
            .ToArray();

        if (!originalText.SequenceEqual(finalText, StringComparer.Ordinal))
            throw new InvalidOperationException(
                "Post-Word TOC OpenXML re-normalization changed visible document text.");

        if (!originalFields.SequenceEqual(finalFields, StringComparer.Ordinal))
            throw new InvalidOperationException(
                "Post-Word TOC OpenXML re-normalization changed field instructions.");

        if (changed)
            ReplaceXmlEntry(archive, "word/document.xml", document);
        if (stylesChanged)
            ReplaceXmlEntry(archive, "word/styles.xml", styles);

        var stats = post.Stats;
        warnings.Add(
            "پس از به‌روزرسانی فهرست توسط Microsoft Word، قرارداد محافظت‌شده صفحه‌بندی جدول‌ها و سپس RTL/Right/Tab فهرست روی همان کپی کنترل‌شده دوباره اعمال و بدون تغییر متن/Field تأیید شد.");

        trace?.Invoke(
            $"POST_WORD_TOC_OOXML_RENORMALIZATION:PASS:"
            + $"TOC={stats.TocParagraphs}:RIGHT={stats.TocPhysicalRight}:RUNRTL={stats.TocTextRunsRtl}:"
            + $"PARAGRAPH_TABS_OK={tabs.ParagraphTabsOk}:STYLE_TABS_OK={tabs.StyleTabsOk}:"
            + $"CAPTIONS_KEEP={stats.CaptionsKeepNext}:ROWS_NOSPLIT={stats.RowsNoSplit}:HEADER_REPEAT={stats.HeaderRowsRepeat}:"
            + $"VISIBLE_TEXT=UNCHANGED:FIELDS=UNCHANGED");
    }

    private static int[] SnapshotWinWordPids()
    {
        var ids = new List<int>();
        foreach (var process in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
        {
            try { ids.Add(process.Id); }
            catch { }
            finally { process.Dispose(); }
        }
        return ids.Distinct().OrderBy(x => x).ToArray();
    }

    private static bool WaitForNoNewWinWord(
        IReadOnlyCollection<int> baselinePids,
        int timeoutMs,
        out int[] remainingPids)
    {
        var baseline = new HashSet<int>(baselinePids);
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        do
        {
            var current = SnapshotWinWordPids();
            remainingPids = current.Where(id => !baseline.Contains(id)).ToArray();
            if (remainingPids.Length == 0)
                return true;
            Thread.Sleep(250);
        }
        while (DateTime.UtcNow < deadline);

        var finalCurrent = SnapshotWinWordPids();
        remainingPids = finalCurrent.Where(id => !baseline.Contains(id)).ToArray();
        return remainingPids.Length == 0;
    }

    private static void RefreshTocResultsWithWord(
        string outputPath,
        AgentProfile profile,
        WordAuditSnapshot beforeWord,
        List<string> warnings,
        Action<string>? trace)
    {
        if (!string.Equals(profile.Id, "proposal.governance.fa", StringComparison.OrdinalIgnoreCase))
        {
            trace?.Invoke("WORD_NATIVE_TOC_REFRESH:SKIP:PROFILE");
            return;
        }

        if (beforeWord.Counts.TocFields == 0)
        {
            trace?.Invoke("WORD_NATIVE_TOC_REFRESH:SKIP:NO_TOC");
            return;
        }

        if (beforeWord.Counts.TocFields != 1)
            throw new InvalidOperationException(
                $"Word-native TOC refresh is fail-closed for this profile unless exactly one TOC exists. Found={beforeWord.Counts.TocFields}.");

        var baselinePids = SnapshotWinWordPids();
        if (baselinePids.Length != 0)
            throw new InvalidOperationException(
                "Word-native TOC refresh blocked because WINWORD already exists. PIDS="
                + string.Join(",", baselinePids)
                + ". Close Word yourself; WordAgent will not kill processes.");

        Exception? last = null;
        for (int attempt = 1; attempt <= 3; attempt++)
        {
            trace?.Invoke($"WORD_NATIVE_TOC_REFRESH:START:ATTEMPT={attempt}");
            try
            {
                RefreshSingleTocWithWord(outputPath, beforeWord, trace);

                if (!WaitForNoNewWinWord(baselinePids, 15000, out var remaining))
                    throw new InvalidOperationException(
                        "Word-native TOC refresh orphan guard failed. REMAINING="
                        + string.Join(",", remaining)
                        + ". WordAgent will not kill processes.");

                warnings.Add("فهرست مطالب با Microsoft Word روی کپی کنترل‌شده به‌روزرسانی شد؛ متن خارج از محدوده فهرست بدون تغییر ماند.");
                trace?.Invoke("WORD_NATIVE_TOC_REFRESH:ORPHAN_GUARD:PASS:REMAINING=0");
                trace?.Invoke("WORD_NATIVE_TOC_REFRESH:PASS:TOCS=1:OUTSIDE_TEXT=UNCHANGED");
                return;
            }
            catch (COMException ex) when (IsTransientRpc(ex) && attempt < 3)
            {
                last = ex;
                if (!WaitForNoNewWinWord(baselinePids, 15000, out var remaining))
                    throw new InvalidOperationException(
                        "Transient Word RPC failure left a WINWORD orphan. REMAINING="
                        + string.Join(",", remaining)
                        + ". WordAgent will not kill processes.", ex);

                warnings.Add($"تلاش مجدد RPC موقت Word برای به‌روزرسانی فهرست: تلاش={attempt}، HRESULT=0x{ex.HResult:X8}.");
                trace?.Invoke("WORD_NATIVE_TOC_REFRESH:ORPHAN_GUARD:PASS:REMAINING=0:AFTER_TRANSIENT_RPC");
                StabilizeReleasedCom();
                Thread.Sleep(300 * attempt);
            }
            catch
            {
                if (!WaitForNoNewWinWord(baselinePids, 15000, out var remaining))
                    trace?.Invoke(
                        "WORD_NATIVE_TOC_REFRESH:ORPHAN_GUARD:FAIL:REMAINING="
                        + string.Join(",", remaining)
                        + ":NO_KILL");
                throw;
            }
        }

        throw new InvalidOperationException(
            "Word-native TOC refresh failed after transient-RPC retries.", last);
    }

    private static void RefreshSingleTocWithWord(
        string outputPath,
        WordAuditSnapshot beforeWord,
        Action<string>? trace)
    {
        object? appObject = null;
        object? optionsObject = null;
        object? documentsObject = null;
        object? documentObject = null;
        object? tablesObject = null;
        object? sectionsObject = null;
        object? tocCollectionObject = null;
        object? tocObject = null;
        object? tocRangeObject = null;
        object? prefixRangeObject = null;
        object? suffixRangeObject = null;
        object? contentRangeObject = null;

        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new InvalidOperationException("Microsoft Word COM automation is unavailable.");

            appObject = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Microsoft Word COM instance could not be created.");

            dynamic app = appObject;
            app.Visible = false;
            app.DisplayAlerts = 0;

            try
            {
                optionsObject = app.Options;
                dynamic options = optionsObject;
                options.UpdateFieldsAtPrint = false;
                options.UpdateLinksAtOpen = false;
                options.UpdateLinksAtPrint = false;
                options.SaveNormalPrompt = false;
            }
            catch
            {
                // Best-effort suppression only; document-level invariants remain authoritative.
            }

            documentsObject = app.Documents;
            dynamic documents = documentsObject;
            documentObject = documents.Open(outputPath, false, false, false);
            dynamic document = documentObject;

            if ((bool)document.ReadOnly)
                throw new InvalidOperationException("Controlled output opened read-only during Word-native TOC refresh.");

            tablesObject = document.Tables;
            dynamic tables = tablesObject;
            if ((int)tables.Count != beforeWord.Counts.Tables)
                throw new InvalidOperationException(
                    $"Table count changed before TOC refresh. Before={beforeWord.Counts.Tables}, Word={tables.Count}.");

            sectionsObject = document.Sections;
            dynamic sections = sectionsObject;
            if ((int)sections.Count != beforeWord.Counts.Sections)
                throw new InvalidOperationException(
                    $"Section count changed before TOC refresh. Before={beforeWord.Counts.Sections}, Word={sections.Count}.");

            tocCollectionObject = document.TablesOfContents;
            dynamic tocCollection = tocCollectionObject;
            if ((int)tocCollection.Count != 1)
                throw new InvalidOperationException(
                    $"Expected exactly one TOC during Word-native refresh. Found={tocCollection.Count}.");

            document.Repaginate();

            tocObject = tocCollection.Item(1);
            dynamic toc = tocObject;
            tocRangeObject = toc.Range;
            dynamic tocRange = tocRangeObject;
            int tocStartBefore = (int)tocRange.Start;
            int tocEndBefore = (int)tocRange.End;

            prefixRangeObject = document.Range(0, tocStartBefore);
            dynamic prefixRange = prefixRangeObject;
            string prefixBefore = (string)prefixRange.Text;

            contentRangeObject = document.Content;
            dynamic contentRange = contentRangeObject;
            int contentEndBefore = (int)contentRange.End;
            ReleaseCom(ref contentRangeObject);

            suffixRangeObject = document.Range(tocEndBefore, contentEndBefore);
            dynamic suffixRange = suffixRangeObject;
            string suffixBefore = (string)suffixRange.Text;

            ReleaseCom(ref prefixRangeObject);
            ReleaseCom(ref suffixRangeObject);
            ReleaseCom(ref tocRangeObject);

            toc.Update();
            document.Repaginate();

            ReleaseCom(ref tocObject);
            tocObject = tocCollection.Item(1);
            toc = tocObject;
            tocRangeObject = toc.Range;
            tocRange = tocRangeObject;

            int tocStartAfter = (int)tocRange.Start;
            int tocEndAfter = (int)tocRange.End;

            prefixRangeObject = document.Range(0, tocStartAfter);
            prefixRange = prefixRangeObject;
            string prefixAfter = (string)prefixRange.Text;

            contentRangeObject = document.Content;
            contentRange = contentRangeObject;
            int contentEndAfter = (int)contentRange.End;
            ReleaseCom(ref contentRangeObject);

            suffixRangeObject = document.Range(tocEndAfter, contentEndAfter);
            suffixRange = suffixRangeObject;
            string suffixAfter = (string)suffixRange.Text;

            if (!string.Equals(prefixBefore, prefixAfter, StringComparison.Ordinal) ||
                !string.Equals(suffixBefore, suffixAfter, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Word-native TOC refresh changed visible text outside the TOC range.");
            }

            if ((int)tables.Count != beforeWord.Counts.Tables ||
                (int)sections.Count != beforeWord.Counts.Sections ||
                (int)tocCollection.Count != 1)
            {
                throw new InvalidOperationException(
                    "Word-native TOC refresh changed protected structural counts.");
            }

            document.Save();
            trace?.Invoke(
                $"WORD_NATIVE_TOC_REFRESH:PROOF:START={tocStartBefore}->{tocStartAfter}:END={tocEndBefore}->{tocEndAfter}");
        }
        finally
        {
            ReleaseCom(ref prefixRangeObject);
            ReleaseCom(ref suffixRangeObject);
            ReleaseCom(ref contentRangeObject);
            ReleaseCom(ref tocRangeObject);
            ReleaseCom(ref tocObject);
            ReleaseCom(ref tocCollectionObject);
            ReleaseCom(ref sectionsObject);
            ReleaseCom(ref tablesObject);

            if (documentObject is not null)
            {
                try
                {
                    dynamic document = documentObject;
                    document.Close(0);
                }
                catch { }
            }
            ReleaseCom(ref documentObject);
            ReleaseCom(ref documentsObject);
            ReleaseCom(ref optionsObject);

            if (appObject is not null)
            {
                try
                {
                    dynamic app = appObject;
                    app.Quit();
                }
                catch { }
            }
            ReleaseCom(ref appObject);
            StabilizeReleasedCom();
        }
    }

    private static void ReleaseCom(ref object? value)
    {
        if (value is null) return;
        try
        {
            if (Marshal.IsComObject(value))
                Marshal.FinalReleaseComObject(value);
        }
        catch { }
        finally
        {
            value = null;
        }
    }

    private static WordAuditSnapshot InspectWordWithRetry(
        string path,
        bool visible,
        List<string> warnings,
        string stage)
    {
        Exception? last = null;
        for (int attempt = 1; attempt <= 3; attempt++)
        {
            try
            {
                return new WordAuditService().Inspect(path, visible);
            }
            catch (COMException ex) when (IsTransientRpc(ex) && attempt < 3)
            {
                last = ex;
                warnings.Add($"تلاش مجدد RPC موقت Word: مرحله={stage}، تلاش={attempt}، HRESULT=0x{ex.HResult:X8}.");
                StabilizeReleasedCom();
                Thread.Sleep(300 * attempt);
            }
        }

        throw new InvalidOperationException($"{stage} failed after transient-RPC retries.", last);
    }

    private static ProfileComplianceSnapshot InspectComplianceWithRetry(
        string path,
        AgentProfile profile,
        bool visible,
        List<string> warnings,
        string stage)
    {
        Exception? last = null;
        for (int attempt = 1; attempt <= 3; attempt++)
        {
            try
            {
                return new ProfileComplianceService().Inspect(path, profile, visible);
            }
            catch (COMException ex) when (IsTransientRpc(ex) && attempt < 3)
            {
                last = ex;
                warnings.Add($"تلاش مجدد RPC موقت Word: مرحله={stage}، تلاش={attempt}، HRESULT=0x{ex.HResult:X8}.");
                StabilizeReleasedCom();
                Thread.Sleep(300 * attempt);
            }
        }

        throw new InvalidOperationException($"{stage} failed after transient-RPC retries.", last);
    }

    private static bool IsTransientRpc(COMException ex)
        => ex.HResult is RpcSServerUnavailable or RpcSCallFailed or RpcEDisconnected or RpcECallRejected;

    private static void StabilizeReleasedCom()
    {
        try
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        catch { }
    }

    private static void TryDeleteFailedOutput(string output, List<string> warnings)
    {
        for (int attempt = 1; attempt <= 4; attempt++)
        {
            try
            {
                if (!File.Exists(output)) return;
                File.Delete(output);
                return;
            }
            catch (Exception ex)
            {
                if (attempt == 4)
                {
                    warnings.Add($"خروجی تأییدنشده حذف نشد: {output} ({ex.Message})");
                    return;
                }
                StabilizeReleasedCom();
                Thread.Sleep(250 * attempt);
            }
        }
    }

    private sealed record TocLocalityIntegrityFingerprint(
        int TocSdtCount,
        string FullVisibleTextSha256,
        string TocVisibleTextSha256,
        string OutsideTocVisibleTextSha256,
        int DocumentParagraphs,
        int TocParagraphs,
        int TocVisibleParagraphs,
        int OutsideTocParagraphs,
        int DocumentTables,
        int OutsideTocTables,
        int DocumentFieldChars,
        int TocFieldChars,
        int OutsideTocFieldChars,
        int DocumentSemanticFields,
        int TocSemanticFields,
        int OutsideTocSemanticFields,
        int DocumentInstrText,
        int TocInstrText,
        int OutsideTocInstrText,
        int DocumentHyperlinks,
        int TocHyperlinks,
        int OutsideTocHyperlinks,
        int DocumentSections,
        int OutsideTocSections,
        int DocumentDrawings,
        int OutsideTocDrawings,
        int TocEnvelopeFieldChars,
        string TocEnvelopeSignature);

    private static List<string> CompareInventory(
        WordAuditSnapshot before,
        WordAuditSnapshot after,
        string inputPath,
        string outputPath,
        string profileId,
        List<string> warnings,
        Action<string>? trace)
    {
        var issues = new List<string>();
        void Check(string name, int a, int b)
        {
            if (a != b) issues.Add($"Inventory changed: {name} {a} -> {b}");
        }

        var paragraphChanged = before.Counts.Paragraphs != after.Counts.Paragraphs;
        var fieldsChanged = before.Counts.Fields != after.Counts.Fields;
        var hyperlinksChanged = before.Counts.Hyperlinks != after.Counts.Hyperlinks;

        // QH01-E247 + QH01-E270 + R8R1/R8R2/R8R3 deep-audit closure:
        // Word-native TOC regeneration may change Paragraphs, Fields and Hyperlinks as derived
        // Word object-model counters. These three counters are accepted only as one indivisible,
        // profile-scoped exception backed by strict persisted-OOXML locality proof.
        // All other Word inventory counters remain strict.
        if (paragraphChanged || fieldsChanged || hyperlinksChanged)
        {
            var locality = VerifyTocLocalizedWordInventoryVariance(
                inputPath, outputPath, before, after, profileId, out var localityEvidence);

            if (locality)
            {
                warnings.Add(
                    $"Word COM inventory variance accepted by TOC-local OOXML guard: " +
                    $"Paragraphs {before.Counts.Paragraphs}->{after.Counts.Paragraphs}; " +
                    $"Fields {before.Counts.Fields}->{after.Counts.Fields}; " +
                    $"Hyperlinks {before.Counts.Hyperlinks}->{after.Counts.Hyperlinks}; " +
                    localityEvidence);
                trace?.Invoke(
                    $"INVENTORY_TOC_LOCALITY_GUARD:PASS:" +
                    $"PARAGRAPHS={before.Counts.Paragraphs}->{after.Counts.Paragraphs}:" +
                    $"FIELDS={before.Counts.Fields}->{after.Counts.Fields}:" +
                    $"HYPERLINKS={before.Counts.Hyperlinks}->{after.Counts.Hyperlinks}:" +
                    localityEvidence);
            }
            else
            {
                if (paragraphChanged)
                    issues.Add($"Inventory changed: Paragraphs {before.Counts.Paragraphs} -> {after.Counts.Paragraphs}");
                if (fieldsChanged)
                    issues.Add($"Inventory changed: Fields {before.Counts.Fields} -> {after.Counts.Fields}");
                if (hyperlinksChanged)
                    issues.Add($"Inventory changed: Hyperlinks {before.Counts.Hyperlinks} -> {after.Counts.Hyperlinks}");
                trace?.Invoke("INVENTORY_TOC_LOCALITY_GUARD:FAIL:" + localityEvidence);
            }
        }

        Check("Tables", before.Counts.Tables, after.Counts.Tables);
        Check("InlineShapes", before.Counts.InlineShapes, after.Counts.InlineShapes);
        Check("FloatingShapes", before.Counts.FloatingShapes, after.Counts.FloatingShapes);
        Check("EndNoteFields", before.Counts.EndNoteFields, after.Counts.EndNoteFields);
        var reconstructionBaselinePresent =
            inputPath.EndsWith(TocReconstructionBaselineSuffix, StringComparison.OrdinalIgnoreCase);

        if (reconstructionBaselinePresent)
        {
            if (before.Counts.TocFields != 0 || after.Counts.TocFields != 1)
                issues.Add($"TOC reconstruction inventory mismatch: TocFields {before.Counts.TocFields} -> {after.Counts.TocFields}");
        }
        else
        {
            Check("TocFields", before.Counts.TocFields, after.Counts.TocFields);
        }

        Check("Sections", before.Counts.Sections, after.Counts.Sections);
        Check("Styles", before.Counts.Styles, after.Counts.Styles);
        Check("Footnotes", before.Counts.Footnotes, after.Counts.Footnotes);
        Check("Endnotes", before.Counts.Endnotes, after.Counts.Endnotes);
        Check("Comments", before.Counts.Comments, after.Counts.Comments);
        Check("Revisions", before.Counts.Revisions, after.Counts.Revisions);
        if (!reconstructionBaselinePresent)
            Check("ContentControls", before.Counts.ContentControls, after.Counts.ContentControls);

        Check("Bookmarks", before.Counts.Bookmarks, after.Counts.Bookmarks);
        Check("Equations", before.Counts.Equations, after.Counts.Equations);
        return issues;
    }

    private static bool VerifyTocLocalizedWordInventoryVariance(
        string inputPath,
        string outputPath,
        WordAuditSnapshot beforeWord,
        WordAuditSnapshot afterWord,
        string profileId,
        out string evidence)
    {
        try
        {
            var before = InspectTocLocalityIntegrity(inputPath);
            var after = InspectTocLocalityIntegrity(outputPath);

            var eligibleProfile =
                string.Equals(profileId, "proposal.governance.fa", StringComparison.OrdinalIgnoreCase)
                || string.Equals(profileId, "book.university.fa", StringComparison.OrdinalIgnoreCase);

            var wordParagraphDelta = afterWord.Counts.Paragraphs - beforeWord.Counts.Paragraphs;
            var wordFieldDelta = afterWord.Counts.Fields - beforeWord.Counts.Fields;
            var wordHyperlinkDelta = afterWord.Counts.Hyperlinks - beforeWord.Counts.Hyperlinks;

            var documentParagraphDelta = after.DocumentParagraphs - before.DocumentParagraphs;
            var tocParagraphDelta = after.TocParagraphs - before.TocParagraphs;

            var documentFieldCharDelta = after.DocumentFieldChars - before.DocumentFieldChars;
            var tocFieldCharDelta = after.TocFieldChars - before.TocFieldChars;

            var documentSemanticFieldDelta = after.DocumentSemanticFields - before.DocumentSemanticFields;
            var tocSemanticFieldDelta = after.TocSemanticFields - before.TocSemanticFields;

            var documentInstrTextDelta = after.DocumentInstrText - before.DocumentInstrText;
            var tocInstrTextDelta = after.TocInstrText - before.TocInstrText;

            var documentHyperlinkDelta = after.DocumentHyperlinks - before.DocumentHyperlinks;
            var tocHyperlinkDelta = after.TocHyperlinks - before.TocHyperlinks;

            var outsideTextEqual = string.Equals(
                before.OutsideTocVisibleTextSha256,
                after.OutsideTocVisibleTextSha256,
                StringComparison.Ordinal);

            var protectedOutsideStructureEqual =
                before.OutsideTocParagraphs == after.OutsideTocParagraphs &&
                before.DocumentTables == after.DocumentTables &&
                before.OutsideTocTables == after.OutsideTocTables &&
                before.OutsideTocFieldChars == after.OutsideTocFieldChars &&
                before.OutsideTocInstrText == after.OutsideTocInstrText &&
                before.OutsideTocHyperlinks == after.OutsideTocHyperlinks &&
                before.DocumentSections == after.DocumentSections &&
                before.OutsideTocSections == after.OutsideTocSections &&
                before.DocumentDrawings == after.DocumentDrawings &&
                before.OutsideTocDrawings == after.OutsideTocDrawings;

            var allRawDeltasAreTocLocal =
                documentParagraphDelta == tocParagraphDelta &&
                documentFieldCharDelta == tocFieldCharDelta &&
                documentSemanticFieldDelta == tocSemanticFieldDelta &&
                documentInstrTextDelta == tocInstrTextDelta &&
                documentHyperlinkDelta == tocHyperlinkDelta;

            var wordParagraphCounterReconciled =
                wordParagraphDelta == tocParagraphDelta;

            var wordHyperlinkCounterReconciled =
                wordHyperlinkDelta == tocHyperlinkDelta;

            // Word's Fields collection counts semantic fields, while OOXML may split one
            // complex field instruction across multiple <w:instrText> nodes. Reconcile
            // against actual field units (fldSimple + fldChar begin), not instrText nodes.
            // Word may additionally expose TOC hyperlinks as wdFieldHyperlink fields.
            var wordFieldCounterReconciled =
                wordFieldDelta == tocSemanticFieldDelta
                || wordFieldDelta == tocSemanticFieldDelta + tocHyperlinkDelta;

            var wordCountersMatchRawTocRewrite =
                wordParagraphCounterReconciled &&
                wordHyperlinkCounterReconciled &&
                wordFieldCounterReconciled;

            // A Word-native TOC update may add/remove stale visible entries and may materialize
            // PAGEREF fields. The safe boundary is spatial, not count-stability: every raw delta
            // must be inside the one TOC, while outside visible text and protected structure are
            // byte/structure stable. Final TOC must remain non-empty.
            var tocShapeIsBoundedToAuthorityRegion =
                before.TocVisibleParagraphs > 0 &&
                after.TocVisibleParagraphs > 0 &&
                after.TocParagraphs > 0 &&
                after.TocInstrText > 0;

            var reconstructionBaseline =
                inputPath.EndsWith(TocReconstructionBaselineSuffix, StringComparison.OrdinalIgnoreCase);
            var tocAuthorityTransitionAccepted = reconstructionBaseline
                ? before.TocSdtCount == 0 && after.TocSdtCount == 1
                : before.TocSdtCount == 1 && after.TocSdtCount == 1;

            var pass =
                eligibleProfile &&
                tocAuthorityTransitionAccepted &&
                outsideTextEqual &&
                protectedOutsideStructureEqual &&
                allRawDeltasAreTocLocal &&
                wordCountersMatchRawTocRewrite &&
                (reconstructionBaseline
                    ? after.TocVisibleParagraphs > 0 && after.TocParagraphs > 0 && after.TocInstrText > 0
                    : tocShapeIsBoundedToAuthorityRegion);

            evidence =
                $"PROFILE={profileId};" +
                $"TOC_AUTHORITY_SCOPES={before.TocSdtCount}->{after.TocSdtCount};" +
                $"VISIBLE_TOC={before.TocVisibleParagraphs}->{after.TocVisibleParagraphs};" +
                $"WORD_P_DELTA={wordParagraphDelta};TOC_P_DELTA={tocParagraphDelta};" +
                $"WORD_F_DELTA={wordFieldDelta};TOC_SEMANTIC_FIELD_DELTA={tocSemanticFieldDelta};TOC_INSTR_DELTA={tocInstrTextDelta};" +
                $"WORD_H_DELTA={wordHyperlinkDelta};TOC_HYPERLINK_DELTA={tocHyperlinkDelta};" +
                $"FIELD_REPRESENTATION_ALLOWED={tocSemanticFieldDelta}|{tocSemanticFieldDelta + tocHyperlinkDelta};" +
                $"RAW_FLDCHAR_DELTA={documentFieldCharDelta};TOC_FLDCHAR_DELTA={tocFieldCharDelta};" +
                $"RAW_SEMANTIC_FIELD_DELTA={documentSemanticFieldDelta};TOC_SEMANTIC_FIELD_DELTA={tocSemanticFieldDelta};" +
                $"HYPERLINK_DELTA={documentHyperlinkDelta}/{tocHyperlinkDelta};" +
                $"OUTSIDE_TEXT_EQUAL={outsideTextEqual};" +
                $"OUTSIDE_STRUCTURE_EQUAL={protectedOutsideStructureEqual};" +
                $"BOUNDARY_MODE={(before.TocEnvelopeSignature.StartsWith("RECOVERED_", StringComparison.Ordinal) ? "RECOVERED_DOCUMENT_LEVEL" : "STANDARD_SDT_OR_OTHER")}->{(after.TocEnvelopeSignature.StartsWith("RECOVERED_", StringComparison.Ordinal) ? "RECOVERED_DOCUMENT_LEVEL" : "STANDARD_SDT_OR_OTHER")};" +
                $"TOC_ENVELOPE_FLDCHAR={before.TocEnvelopeFieldChars}->{after.TocEnvelopeFieldChars};" +
                $"TOC_ENVELOPE_SIGNATURE={before.TocEnvelopeSignature}->{after.TocEnvelopeSignature};" +
                $"OUTSIDE_P={before.OutsideTocParagraphs}->{after.OutsideTocParagraphs};" +
                $"OUTSIDE_FLDCHAR={before.OutsideTocFieldChars}->{after.OutsideTocFieldChars};" +
                $"OUTSIDE_SEM_FIELD={before.OutsideTocSemanticFields}->{after.OutsideTocSemanticFields};" +
                $"OUTSIDE_INSTR={before.OutsideTocInstrText}->{after.OutsideTocInstrText};" +
                $"OUTSIDE_HYPERLINK={before.OutsideTocHyperlinks}->{after.OutsideTocHyperlinks};" +
                $"OUTSIDE_SECTION={before.OutsideTocSections}->{after.OutsideTocSections};" +
                $"OUTSIDE_DRAWING={before.OutsideTocDrawings}->{after.OutsideTocDrawings};" +
                $"RAW_DELTAS_TOC_LOCAL={allRawDeltasAreTocLocal};" +
                $"WORD_COUNTERS_MATCH_TOC={wordCountersMatchRawTocRewrite};" +
                $"TOC_SHAPE_BOUNDED_TO_AUTHORITY={tocShapeIsBoundedToAuthorityRegion}";

            return pass;
        }
        catch (Exception ex)
        {
            evidence = "forensic-error=" + ex.Message;
            return false;
        }
    }

    private static string NormalizeTocIntegrityText(XElement paragraph)
        => string.Join(
            " ",
            paragraph.Descendants(W + "t")
                .Select(x => ((string?)x ?? string.Empty).Trim())
                .Where(x => x.Length > 0))
            .Trim();

    private static bool IsTocIntegrityTitle(string text)
    {
        var compact = new string(
            (text ?? string.Empty)
                .Where(ch => !char.IsWhiteSpace(ch) && ch != '\u200c' && ch != '\u200f' && ch != '\u200e')
                .ToArray());
        return compact.Equals("فهرستمطالب", StringComparison.Ordinal);
    }

    private static bool LooksLikePlainTocIntegrityEntry(XElement paragraph)
    {
        var text = NormalizeTocIntegrityText(paragraph);
        if (string.IsNullOrWhiteSpace(text) || !paragraph.Descendants(W + "tab").Any())
            return false;

        var tail = text.Split(' ', StringSplitOptions.RemoveEmptyEntries).LastOrDefault() ?? string.Empty;
        return tail.Length > 0
               && tail.All(ch =>
                   char.IsDigit(ch)
                   || (ch >= '\u06F0' && ch <= '\u06F9')
                   || (ch >= '\u0660' && ch <= '\u0669'));
    }

    private static int[] FindTocIntegrityParagraphIndexes(XDocument document)
    {
        var paragraphs = document.Descendants(W + "p").ToArray();

        var mainAndEntryIndexes = paragraphs
            .Select((paragraph, index) => new { paragraph, index })
            .Where(x =>
            {
                var instruction = string.Join(
                    " ",
                    x.paragraph.Descendants(W + "instrText")
                        .Select(t => ((string?)t ?? string.Empty).Trim()));

                var main = instruction.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)
                           && (instruction.Length == 3 || char.IsWhiteSpace(instruction[3]));
                var entry =
                    (instruction.Contains("PAGEREF", StringComparison.OrdinalIgnoreCase)
                     && instruction.Contains("_Toc", StringComparison.OrdinalIgnoreCase))
                    || x.paragraph.Descendants(W + "hyperlink")
                        .Select(link => (string?)link.Attribute(W + "anchor"))
                        .Any(anchor =>
                            !string.IsNullOrWhiteSpace(anchor)
                            && anchor.StartsWith("_Toc", StringComparison.OrdinalIgnoreCase));
                return main || entry;
            })
            .Select(x => x.index)
            .ToArray();

        if (mainAndEntryIndexes.Length >= 2)
        {
            var start = mainAndEntryIndexes.Min();
            var end = mainAndEntryIndexes.Max();
            if (end - start <= 250)
                return Enumerable.Range(start, end - start + 1).ToArray();
        }

        var titleIndex = Array.FindIndex(
            paragraphs,
            paragraph => IsTocIntegrityTitle(NormalizeTocIntegrityText(paragraph)));

        if (titleIndex < 0)
            return Array.Empty<int>();

        var degraded = new List<int>();
        for (var index = titleIndex + 1; index < paragraphs.Length && degraded.Count < 120; index++)
        {
            var paragraph = paragraphs[index];
            var text = NormalizeTocIntegrityText(paragraph);
            if (string.IsNullOrWhiteSpace(text))
            {
                if (degraded.Count == 0)
                    continue;
                break;
            }

            if (!LooksLikePlainTocIntegrityEntry(paragraph))
                break;

            degraded.Add(index);
        }

        return degraded.Count >= 2 ? degraded.ToArray() : Array.Empty<int>();
    }

    private static XElement[] FindTocFieldEnvelopeExtensions(XElement tocSdt, out string signature)
    {
        signature = "NONE";
        if (tocSdt.Parent is null)
            return Array.Empty<XElement>();

        static string FieldCharType(XElement field)
            => ((string?)field.Attribute(W + "fldCharType") ?? string.Empty).Trim().ToLowerInvariant();

        static bool IsFieldMarkerOnlyParagraph(XElement paragraph, string expectedType, out XElement? marker)
        {
            marker = null;
            if (paragraph.Name != W + "p")
                return false;

            if (paragraph.Descendants(W + "t").Any(node => !string.IsNullOrWhiteSpace((string?)node)))
                return false;
            if (paragraph.Descendants(W + "instrText").Any(node => !string.IsNullOrWhiteSpace((string?)node)))
                return false;
            if (paragraph.Descendants(W + "hyperlink").Any()
                || paragraph.Descendants(W + "drawing").Any()
                || paragraph.Descendants(W + "object").Any()
                || paragraph.Descendants(W + "pict").Any()
                || paragraph.Descendants(W + "tbl").Any()
                || paragraph.Descendants(W + "sectPr").Any())
                return false;

            var markers = paragraph.Descendants(W + "fldChar").ToArray();
            if (markers.Length != 1 || !FieldCharType(markers[0]).Equals(expectedType, StringComparison.OrdinalIgnoreCase))
                return false;

            // Do not absorb a paragraph that contains any other semantic field material.
            if (paragraph.Descendants(W + "fldSimple").Any())
                return false;

            marker = markers[0];
            return true;
        }

        var insideMarkers = tocSdt.Descendants(W + "fldChar").ToArray();
        var beginCount = insideMarkers.Count(field => FieldCharType(field) == "begin");
        var endCount = insideMarkers.Count(field => FieldCharType(field) == "end");
        var separateCount = insideMarkers.Count(field => FieldCharType(field) == "separate");

        // Word may persist the closing marker of the outer TOC complex field in an immediately
        // adjacent, otherwise-empty paragraph outside the automatic-TOC SDT. Accept only the
        // exact complementary marker needed to balance the field envelope; anything else remains
        // outside authority and therefore fails closed.
        if (beginCount == endCount + 1)
        {
            var next = tocSdt.ElementsAfterSelf().FirstOrDefault();
            if (next is not null
                && IsFieldMarkerOnlyParagraph(next, "end", out var marker)
                && marker is not null)
            {
                signature = $"AFTER_END:BEGIN={beginCount}:SEPARATE={separateCount}:END={endCount}:EXT=1";
                return new[] { marker };
            }
        }

        if (endCount == beginCount + 1)
        {
            var previous = tocSdt.ElementsBeforeSelf().LastOrDefault();
            if (previous is not null
                && IsFieldMarkerOnlyParagraph(previous, "begin", out var marker)
                && marker is not null)
            {
                signature = $"BEFORE_BEGIN:BEGIN={beginCount}:SEPARATE={separateCount}:END={endCount}:EXT=1";
                return new[] { marker };
            }
        }

        signature = $"UNBALANCED_OR_NO_SAFE_ADJACENT_MARKER:BEGIN={beginCount}:SEPARATE={separateCount}:END={endCount}";
        return Array.Empty<XElement>();
    }


    private static XElement[] FindRecoveredTocFieldEnvelopeExtensions(
        XDocument document,
        IReadOnlyList<int> tocIndexes,
        out string signature)
    {
        signature = "RECOVERED_NO_ENVELOPE";

        var paragraphs = document.Descendants(W + "p").ToArray();
        var ordered = tocIndexes
            .Where(index => index >= 0 && index < paragraphs.Length)
            .Distinct()
            .OrderBy(index => index)
            .ToArray();

        if (ordered.Length < 2)
        {
            signature = $"RECOVERED_RANGE_UNAVAILABLE:COUNT={ordered.Length}";
            return Array.Empty<XElement>();
        }

        for (var i = 1; i < ordered.Length; i++)
        {
            if (ordered[i] != ordered[i - 1] + 1)
            {
                signature = $"RECOVERED_RANGE_NONCONTIGUOUS:COUNT={ordered.Length}:FIRST={ordered[0]}:LAST={ordered[^1]}";
                return Array.Empty<XElement>();
            }
        }

        static string FieldCharType(XElement field)
            => ((string?)field.Attribute(W + "fldCharType") ?? string.Empty).Trim().ToLowerInvariant();

        static bool IsFieldMarkerOnlyParagraph(XElement paragraph, string expectedType, out XElement? marker)
        {
            marker = null;

            if (paragraph.Descendants(W + "t").Any(node => !string.IsNullOrWhiteSpace((string?)node)))
                return false;
            if (paragraph.Descendants(W + "instrText").Any(node => !string.IsNullOrWhiteSpace((string?)node)))
                return false;
            if (paragraph.Descendants(W + "hyperlink").Any()
                || paragraph.Descendants(W + "drawing").Any()
                || paragraph.Descendants(W + "object").Any()
                || paragraph.Descendants(W + "pict").Any()
                || paragraph.Descendants(W + "tbl").Any()
                || paragraph.Descendants(W + "sectPr").Any()
                || paragraph.Descendants(W + "fldSimple").Any())
                return false;

            var markers = paragraph.Descendants(W + "fldChar").ToArray();
            if (markers.Length != 1
                || !FieldCharType(markers[0]).Equals(expectedType, StringComparison.OrdinalIgnoreCase))
                return false;

            marker = markers[0];
            return true;
        }

        var tocParagraphs = ordered.Select(index => paragraphs[index]).ToArray();
        var insideMarkers = tocParagraphs.SelectMany(p => p.Descendants(W + "fldChar")).ToArray();
        var beginCount = insideMarkers.Count(field => FieldCharType(field) == "begin");
        var endCount = insideMarkers.Count(field => FieldCharType(field) == "end");
        var separateCount = insideMarkers.Count(field => FieldCharType(field) == "separate");

        if (beginCount == endCount + 1)
        {
            var nextIndex = ordered[^1] + 1;
            if (nextIndex < paragraphs.Length
                && IsFieldMarkerOnlyParagraph(paragraphs[nextIndex], "end", out var marker)
                && marker is not null)
            {
                signature =
                    $"RECOVERED_AFTER_END:BEGIN={beginCount}:SEPARATE={separateCount}:END={endCount}:" +
                    $"FIRST={ordered[0]}:LAST={ordered[^1]}:EXT_INDEX={nextIndex}:EXT=1";
                return new[] { marker };
            }
        }

        if (endCount == beginCount + 1)
        {
            var previousIndex = ordered[0] - 1;
            if (previousIndex >= 0
                && IsFieldMarkerOnlyParagraph(paragraphs[previousIndex], "begin", out var marker)
                && marker is not null)
            {
                signature =
                    $"RECOVERED_BEFORE_BEGIN:BEGIN={beginCount}:SEPARATE={separateCount}:END={endCount}:" +
                    $"FIRST={ordered[0]}:LAST={ordered[^1]}:EXT_INDEX={previousIndex}:EXT=1";
                return new[] { marker };
            }
        }

        signature =
            $"RECOVERED_BALANCED_OR_NO_SAFE_ADJACENT_MARKER:BEGIN={beginCount}:SEPARATE={separateCount}:END={endCount}:" +
            $"FIRST={ordered[0]}:LAST={ordered[^1]}";
        return Array.Empty<XElement>();
    }

    private static TocLocalityIntegrityFingerprint InspectTocLocalityIntegrity(string path)
    {
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var document = LoadXmlEntry(archive, "word/document.xml");

        static int CountSemanticFields(XContainer root)
            => root.Descendants(W + "fldSimple").Count()
               + root.Descendants(W + "fldChar").Count(field =>
                   string.Equals(
                       ((string?)field.Attribute(W + "fldCharType") ?? string.Empty).Trim(),
                       "begin",
                       StringComparison.OrdinalIgnoreCase));

        // R8R23R16: locality authority follows the actual Word content-control boundary when
        // exactly one recognized automatic TOC SDT exists. Paragraph-span inference remains
        // only as a fail-closed fallback for degraded documents without a recognized wrapper.
        var recognizedTocSdts = document
            .Descendants(W + "sdt")
            .Where(IsRecognizedAutomaticTocSdt)
            .Where(sdt => !sdt.Ancestors(W + "sdt").Any(IsRecognizedAutomaticTocSdt))
            .ToArray();

        var paragraphs = document.Descendants(W + "p").ToArray();
        var tocIndexes = FindTocIntegrityParagraphIndexes(document);

        var tocEnvelopeSignature = "NO_TOC_AUTHORITY_ENVELOPE";
        var tocEnvelopeFieldChars = recognizedTocSdts.Length == 1
            ? FindTocFieldEnvelopeExtensions(recognizedTocSdts[0], out tocEnvelopeSignature)
            : tocIndexes.Length > 0
                ? FindRecoveredTocFieldEnvelopeExtensions(document, tocIndexes, out tocEnvelopeSignature)
                : Array.Empty<XElement>();

        XElement[] tocParagraphs;
        var outside = new XDocument(document);

        if (recognizedTocSdts.Length == 1)
        {
            tocParagraphs = recognizedTocSdts[0].Descendants(W + "p").ToArray();

            var outsideRecognized = outside
                .Descendants(W + "sdt")
                .Where(IsRecognizedAutomaticTocSdt)
                .Where(sdt => !sdt.Ancestors(W + "sdt").Any(IsRecognizedAutomaticTocSdt))
                .ToArray();

            if (outsideRecognized.Length != 1)
                throw new InvalidOperationException(
                    $"TOC SDT clone boundary contradiction: source=1 clone={outsideRecognized.Length}.");

            var outsideEnvelopeFieldChars =
                FindTocFieldEnvelopeExtensions(outsideRecognized[0], out var outsideEnvelopeSignature);

            if (outsideEnvelopeFieldChars.Length != tocEnvelopeFieldChars.Length
                || !string.Equals(outsideEnvelopeSignature, tocEnvelopeSignature, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    $"TOC field-envelope clone contradiction: source={tocEnvelopeSignature}/{tocEnvelopeFieldChars.Length} " +
                    $"clone={outsideEnvelopeSignature}/{outsideEnvelopeFieldChars.Length}.");

            foreach (var marker in outsideEnvelopeFieldChars)
                marker.Remove();

            outsideRecognized[0].Remove();
        }
        else
        {
            tocParagraphs = tocIndexes
                .Where(index => index >= 0 && index < paragraphs.Length)
                .Select(index => paragraphs[index])
                .ToArray();

            var outsideEnvelopeSignature = "NO_TOC_AUTHORITY_ENVELOPE";
            var outsideEnvelopeFieldChars = tocIndexes.Length > 0
                ? FindRecoveredTocFieldEnvelopeExtensions(outside, tocIndexes, out outsideEnvelopeSignature)
                : Array.Empty<XElement>();

            var expectedOutsideEnvelopeSignature = outsideEnvelopeSignature;

            if (outsideEnvelopeFieldChars.Length != tocEnvelopeFieldChars.Length
                || !string.Equals(expectedOutsideEnvelopeSignature, tocEnvelopeSignature, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    $"Recovered TOC field-envelope clone contradiction: source={tocEnvelopeSignature}/{tocEnvelopeFieldChars.Length} " +
                    $"clone={expectedOutsideEnvelopeSignature}/{outsideEnvelopeFieldChars.Length}.");

            foreach (var marker in outsideEnvelopeFieldChars)
                marker.Remove();

            var outsideParagraphsFallback = outside.Descendants(W + "p").ToArray();
            foreach (var index in tocIndexes.OrderByDescending(x => x))
            {
                if (index >= 0 && index < outsideParagraphsFallback.Length)
                    outsideParagraphsFallback[index].Remove();
            }
        }

        var logicalTocCount = recognizedTocSdts.Length > 0
            ? recognizedTocSdts.Length
            : (tocParagraphs.Length > 0 ? 1 : 0);

        var tocText = string.Concat(
            tocParagraphs.SelectMany(p => p.Descendants(W + "t"))
                .Select(x => (string?)x ?? string.Empty));

        var tocVisibleParagraphs = tocParagraphs.Count(paragraph =>
            paragraph.Descendants(W + "t")
                .Any(text => !string.IsNullOrWhiteSpace((string?)text)));

        var fullText = string.Concat(document.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));
        var outsideText = string.Concat(outside.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));

        static string TextSha256(string value)
            => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));

        // R8R23R16: a recognized automatic-TOC SDT remains the primary boundary; when Word persists
        // the TOC as a document-level recovered TOC, a deterministic adjacent complementary fldChar closes the
        // outer TOC complex-field envelope. WordprocessingML SDT content can legally contain
        // run-level content directly; Word may also serialize the outer closing marker in an
        // otherwise-empty adjacent paragraph. No visible text or unrelated field material is
        // absorbed into authority. R13-R15 undercounted one fldChar (86 vs document delta 87)
        // while all 29 semantic fields implied exactly 87 field characters. Keep paragraph
        // fallback is retained for recovered document-level TOC authority without a recognized TOC SDT.
        var tocBoundary = recognizedTocSdts.Length == 1 ? recognizedTocSdts[0] : null;

        var tocEnvelopeSemanticBegins = tocEnvelopeFieldChars.Count(field =>
            string.Equals(
                ((string?)field.Attribute(W + "fldCharType") ?? string.Empty).Trim(),
                "begin",
                StringComparison.OrdinalIgnoreCase));

        var tocFieldChars = tocBoundary is not null
            ? tocBoundary.Descendants(W + "fldChar").Count() + tocEnvelopeFieldChars.Length
            : tocParagraphs.Sum(p => p.Descendants(W + "fldChar").Count()) + tocEnvelopeFieldChars.Length;

        var tocSemanticFields = tocBoundary is not null
            ? CountSemanticFields(tocBoundary) + tocEnvelopeSemanticBegins
            : tocParagraphs.Sum(p =>
                p.Descendants(W + "fldSimple").Count()
                + p.Descendants(W + "fldChar").Count(field =>
                    string.Equals(
                        ((string?)field.Attribute(W + "fldCharType") ?? string.Empty).Trim(),
                        "begin",
                        StringComparison.OrdinalIgnoreCase)))
              + tocEnvelopeSemanticBegins;

        var tocInstrText = tocBoundary is not null
            ? tocBoundary.Descendants(W + "instrText").Count()
            : tocParagraphs.Sum(p => p.Descendants(W + "instrText").Count());

        var tocHyperlinks = tocBoundary is not null
            ? tocBoundary.Descendants(W + "hyperlink").Count()
            : tocParagraphs.Sum(p => p.Descendants(W + "hyperlink").Count());

        return new TocLocalityIntegrityFingerprint(
            logicalTocCount,
            TextSha256(fullText),
            TextSha256(tocText),
            TextSha256(outsideText),
            document.Descendants(W + "p").Count(),
            tocParagraphs.Length,
            tocVisibleParagraphs,
            outside.Descendants(W + "p").Count(),
            document.Descendants(W + "tbl").Count(),
            outside.Descendants(W + "tbl").Count(),
            document.Descendants(W + "fldChar").Count(),
            tocFieldChars,
            outside.Descendants(W + "fldChar").Count(),
            CountSemanticFields(document),
            tocSemanticFields,
            CountSemanticFields(outside),
            document.Descendants(W + "instrText").Count(),
            tocInstrText,
            outside.Descendants(W + "instrText").Count(),
            document.Descendants(W + "hyperlink").Count(),
            tocHyperlinks,
            outside.Descendants(W + "hyperlink").Count(),
            document.Descendants(W + "sectPr").Count(),
            outside.Descendants(W + "sectPr").Count(),
            document.Descendants(W + "drawing").Count(),
            outside.Descendants(W + "drawing").Count(),
            tocEnvelopeFieldChars.Length,
            tocEnvelopeSignature);
    }

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }
}
