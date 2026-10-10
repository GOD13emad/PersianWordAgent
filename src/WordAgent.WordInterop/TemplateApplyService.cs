using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using WordAgent.Core;

namespace WordAgent.WordInterop;

public sealed record TemplateApplyInventory(
    int PackageEntries,
    int StoryParts,
    int Paragraphs,
    int Tables,
    int Drawings,
    int FieldInstructions,
    int TextNodes,
    int RevisionElements,
    int SectionProperties,
    int ContentControls,
    int Bookmarks,
    int Hyperlinks,
    int RelationshipsParts,
    int MediaEntries,
    int StyleDefinitions,
    string TextPayloadSha256,
    string FieldInstructionSha256,
    string MediaPayloadSha256);

public sealed record TemplateStyleMapping(
    string SourceStyleId,
    string? SourceStyleName,
    string TargetStyleId,
    string? TargetStyleName,
    string Reason,
    int ParagraphCount);

public sealed record TemplateApplyResult(
    string InputFile,
    string TemplateFile,
    string OutputFile,
    string InputSha256,
    string TemplateSha256,
    string OutputSha256,
    int TemplateStylesImported,
    int TemplateStylesReplaced,
    int ParagraphsRemapped,
    int SectionsNormalized,
    bool ThemeReplaced,
    bool FontTableReplaced,
    bool NoChangesNeeded,
    IReadOnlyList<TemplateStyleMapping> StyleMappings,
    TemplateApplyInventory Before,
    TemplateApplyInventory After,
    IReadOnlyList<string> IntegrityIssues,
    IReadOnlyList<string> Warnings);

public sealed class TemplateApplyService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly string[] SectionPropertyOrder =
    [
        "headerReference","footerReference","footnotePr","endnotePr","type","pgSz","pgMar",
        "paperSrc","pgBorders","lnNumType","pgNumType","cols","formProt","vAlign",
        "noEndnote","titlePg","textDirection","bidi","rtlGutter","docGrid","printerSettings",
        "sectPrChange"
    ];

    private static readonly string[] StylePropertyOrder =
    [
        "name","aliases","basedOn","next","link","autoRedefine","hidden","uiPriority",
        "semiHidden","unhideWhenUsed","qFormat","locked","personal","personalCompose",
        "personalReply","rsid","pPr","rPr","tblPr","trPr","tcPr"
    ];

    private static readonly string[] ParagraphPropertyOrder =
    [
        "pStyle","keepNext","keepLines","pageBreakBefore","framePr","widowControl","numPr",
        "suppressLineNumbers","pBdr","shd","tabs","suppressAutoHyphens","kinsoku","wordWrap",
        "overflowPunct","topLinePunct","autoSpaceDE","autoSpaceDN","bidi","adjustRightInd",
        "snapToGrid","spacing","ind","contextualSpacing","mirrorIndents","suppressOverlap","jc",
        "textDirection","textAlignment","textboxTightWrap","outlineLvl","divId","cnfStyle",
        "rPr","sectPr","pPrChange"
    ];

    private sealed record StyleInfo(
        string StyleId,
        string Type,
        string? Name,
        bool IsDefault,
        int? OutlineLevel);

    private sealed record TemplateSnapshot(
        XDocument Document,
        XDocument Styles,
        byte[]? ThemeBytes,
        byte[]? FontTableBytes,
        IReadOnlyDictionary<string, StyleInfo> StylesById,
        string? DefaultParagraphStyleId,
        IReadOnlyDictionary<int, string> HeadingStyles,
        IReadOnlyDictionary<string, string> CanonicalRoles,
        IReadOnlyDictionary<string, string> ParagraphStyleNames);

    public TemplateApplyResult Apply(
        string inputPath,
        string templatePath,
        string? outputDirectory = null)
    {
        var input = PathPolicy.RequireSupportedWordFile(inputPath);
        var template = PathPolicy.RequireSupportedWordFile(templatePath);

        if (Path.GetFullPath(input).Equals(Path.GetFullPath(template), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Input and template must be different files.");

        PreflightInput(input);

        var before = InspectPackage(input);
        if (before.RevisionElements != 0)
        {
            throw new InvalidOperationException(
                $"Template Apply blocked: tracked-revision XML elements={before.RevisionElements}. Accept/reject revisions first; this mode never rewrites tracked-change markup.");
        }

        var inputHash = Sha256File(input);
        var templateHash = Sha256File(template);
        var templateSnapshot = LoadTemplateSnapshot(template, input);

        var output = PathPolicy.CreateSafeOutputPath(input, outputDirectory, "_WordAgent_TemplateApplied");
        var transaction = CreateTransactionalTempPath(output);
        File.Copy(input, transaction, overwrite: false);
        PauseForTransactionTest();

        var immutableBefore = ReadEntryHashes(input);
        var mutatedEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var warnings = new List<string>
        {
            "Template Apply MVP is deterministic and does not use AI or semantic model inference.",
            "Applied scope: template paragraph/table style definitions, paragraph-style role mapping, page size/margins/columns, and the existing target theme part. Source fontTable and its relationship graph remain authoritative.",
            "Preserved scope: visible text, field instructions, images/media, relationships, source headers/footers, source numbering definitions/instances, source style numbering bindings, direct paragraph numbering bindings, comments, bookmarks, hyperlinks, and source word/fontTable.xml plus its relationship graph. Persian typography/direction is normalized afterward by the shared Persian Foundation.",
            "Template numbering topology is never transplanted implicitly. Template style numbering bindings are stripped unless an existing source style already owns a numbering binding, in which case the source binding is preserved.",
            "Template headers/footers and numbering topology transplantation remain separate explicit future capabilities.",
            "Output is written to an isolated transaction file and promoted to the final path only after all integrity gates pass."
        };

        int stylesImported = 0;
        int stylesReplaced = 0;
        int sourceNumberingBindingsPreserved = 0;
        int templateNumberingBindingsStripped = 0;
        int paragraphsRemapped = 0;
        int sectionsNormalized = 0;
        bool themeReplaced = false;
        bool fontTableReplaced = false;
        var mappingResults = new List<TemplateStyleMapping>();

        try
        {
            using var file = new FileStream(transaction, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            using var archive = new ZipArchive(file, ZipArchiveMode.Update, leaveOpen: false);

            if (archive.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "Digitally signed Office package detected; Template Apply is blocked because package mutation would invalidate the signature.");
            }

            var sourceStyles = LoadXmlEntry(archive, "word/styles.xml");
            var sourceStyleInfos = ParseStyles(sourceStyles);
            var styleMapping = BuildParagraphStyleMapping(sourceStyleInfos, templateSnapshot);

            var styleChanged = MergeTemplateStyles(
                sourceStyles,
                templateSnapshot.Styles,
                ref stylesImported,
                ref stylesReplaced,
                ref sourceNumberingBindingsPreserved,
                ref templateNumberingBindingsStripped);

            if (styleChanged)
            {
                ReplaceXmlEntry(archive, "word/styles.xml", sourceStyles);
                mutatedEntries.Add("word/styles.xml");
            }

            var document = LoadXmlEntry(archive, "word/document.xml");
            var mappingCounts = ApplyParagraphMappings(document, styleMapping);
            paragraphsRemapped = mappingCounts.Values.Sum();

            foreach (var pair in styleMapping)
            {
                if (!mappingCounts.TryGetValue(pair.Key, out var count) || count == 0)
                    continue;

                sourceStyleInfos.TryGetValue(pair.Key, out var sourceInfo);
                templateSnapshot.StylesById.TryGetValue(pair.Value.TargetStyleId, out var targetInfo);

                mappingResults.Add(new TemplateStyleMapping(
                    pair.Key,
                    sourceInfo?.Name,
                    pair.Value.TargetStyleId,
                    targetInfo?.Name,
                    pair.Value.Reason,
                    count));
            }

            var sectionChanges = ApplySectionGeometry(document, templateSnapshot.Document);
            sectionsNormalized = sectionChanges;

            if (paragraphsRemapped > 0 || sectionChanges > 0)
            {
                ReplaceXmlEntry(archive, "word/document.xml", document);
                mutatedEntries.Add("word/document.xml");
            }

            themeReplaced = ReplaceBinaryFromTemplateIfTargetExists(
                archive,
                "word/theme/theme1.xml",
                templateSnapshot.ThemeBytes,
                mutatedEntries);

            // Relationship-safe font policy:
            // Never raw-copy a template fontTable.xml without also transplanting its FontPart
            // relationship graph.  The real Birjand template contains w:embed* r:id values;
            // raw replacement previously created dangling font relationships.  Preserve the
            // source fontTable.xml and its relationships.  Template font choices still flow
            // through imported style/rFonts definitions, while the shared Persian Foundation
            // owns final Persian complex-script typography.
            fontTableReplaced = false;
            warnings.Add(
                "Relationship safety: source word/fontTable.xml and its relationship graph were preserved; raw template fontTable replacement is disabled.");

            warnings.Add(
                $"Numbering safety: source bindings preserved={sourceNumberingBindingsPreserved}; " +
                $"template bindings stripped={templateNumberingBindingsStripped}; source numbering topology remains authoritative.");
            warnings.Add(
                "Template style verification preserves template authority except source numbering bindings and the shared Persian Foundation-owned complex-script font/language properties; those Persian typography properties are intentionally normalized after template application.");
        }
        catch
        {
            TryDelete(transaction);
            throw;
        }

        try
        {
            var foundation = new PersianDocumentFoundationService().ApplyInPlace(
                transaction,
                new PersianDocumentFoundationOptions(ProfileId: null));
            foreach (var entry in foundation.MutatedEntries)
                mutatedEntries.Add(entry);

            warnings.Add(
                $"Persian Foundation after Template Apply: paragraphs={foundation.PersianParagraphs}; " +
                $"bodyJustified={foundation.BodyParagraphsJustified}; headingsRight={foundation.HeadingParagraphsRightAligned}; " +
                $"toc={foundation.TocParagraphsNormalized}; runs={foundation.PersianRunsNormalized}; styles={foundation.StyleDefinitionsNormalized}.");

            if (PersianDocumentFoundationService.RestabilizeExistingAutomaticListsOpenXml(transaction))
            {
                mutatedEntries.Add("word/styles.xml");
                mutatedEntries.Add("word/settings.xml");
                warnings.Add("Automatic-list authority: existing native list styles re-stabilized after template style merge via shared Persian Foundation OpenXML contract.");
            }

            if (EnsureAutomaticTocUpdateOnOpen(transaction))
            {
                mutatedEntries.Add("word/settings.xml");
                warnings.Add("Automatic TOC lifecycle: existing native TOC field detected; word/settings.xml updateFields=true enforced.");
            }
        }
        catch
        {
            TryDelete(transaction);
            throw;
        }

        var noChangesNeeded = mutatedEntries.Count == 0;
        if (noChangesNeeded)
        {
            TryDelete(transaction);
            File.Copy(input, transaction, overwrite: false);
        }

        var after = InspectPackage(transaction);
        var issues = CompareInventory(before, after);
        issues.AddRange(CompareImmutableEntries(transaction, immutableBefore, mutatedEntries));
        issues.AddRange(VerifyTemplateApplication(transaction, templateSnapshot));
        issues.AddRange(VerifyNumberingSafety(input, transaction));

        var outputHash = Sha256File(transaction);
        if (noChangesNeeded && !inputHash.Equals(outputHash, StringComparison.OrdinalIgnoreCase))
            issues.Add("Idempotency mismatch: no template changes were needed but output bytes differ from input.");

        if (!noChangesNeeded && inputHash.Equals(outputHash, StringComparison.OrdinalIgnoreCase))
            issues.Add("Persistence mismatch: template mutations were reported but output SHA-256 equals input.");

        var templateHashAfter = Sha256File(template);
        if (!templateHash.Equals(templateHashAfter, StringComparison.OrdinalIgnoreCase))
            issues.Add("Template input changed unexpectedly.");

        if (issues.Count != 0)
        {
            TryDelete(transaction);
            throw new InvalidOperationException(
                "Template Apply integrity gate failed. " + string.Join(" | ", issues.Take(16)));
        }

        CommitTransactionalOutput(transaction, output);

        return new TemplateApplyResult(
            input,
            template,
            output,
            inputHash,
            templateHash,
            outputHash,
            stylesImported,
            stylesReplaced,
            paragraphsRemapped,
            sectionsNormalized,
            themeReplaced,
            fontTableReplaced,
            noChangesNeeded,
            mappingResults,
            before,
            after,
            issues,
            warnings);
    }

    private static bool EnsureAutomaticTocUpdateOnOpen(string packagePath)
    {
        using var file = new FileStream(packagePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(file, ZipArchiveMode.Update, leaveOpen: false);

        var documentEntry = archive.GetEntry("word/document.xml");
        var settingsEntry = archive.GetEntry("word/settings.xml");
        if (documentEntry is null || settingsEntry is null)
            return false;

        XDocument document;
        using (var stream = documentEntry.Open())
            document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);

        // Word may split one field instruction across multiple w:instrText nodes.
        // A native TOC can therefore appear as an exact standalone "TOC" node followed by
        // separate nodes for the switches. End-of-node after TOC is a valid field-keyword boundary.
        var hasTocField = document
            .Descendants(W + "instrText")
            .Select(node => ((string?)node ?? string.Empty).Trim())
            .Any(instruction => Regex.IsMatch(instruction, @"^TOC(?:\s|\\|$)", RegexOptions.IgnoreCase));

        if (!hasTocField)
            return false;

        XDocument settings;
        using (var stream = settingsEntry.Open())
            settings = XDocument.Load(stream, LoadOptions.PreserveWhitespace);

        var root = settings.Root
            ?? throw new InvalidOperationException("Invalid word/settings.xml: root element is missing.");

        var updateFields = root.Element(W + "updateFields");
        if (updateFields is not null)
        {
            var current = ((string?)updateFields.Attribute(W + "val") ?? string.Empty).Trim();
            if (current.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                current.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                current.Equals("on", StringComparison.OrdinalIgnoreCase))
                return false;
        }
        else
        {
            updateFields = new XElement(W + "updateFields");
            InsertSettingsElementInSchemaOrder(root, updateFields);
        }

        updateFields.SetAttributeValue(W + "val", "true");
        ReplaceXmlEntry(archive, "word/settings.xml", settings);
        return true;
    }

    private static readonly string[] SettingsElementsAfterUpdateFields =
    {
        "hdrShapeDefaults",
        "footnotePr",
        "endnotePr",
        "compat",
        "docVars",
        "rsids",
        "mathPr",
        "uiCompat97To2003",
        "attachedSchema",
        "themeFontLang",
        "clrSchemeMapping",
        "doNotIncludeSubdocsInStats",
        "doNotAutoCompressPictures",
        "forceUpgrade",
        "captions",
        "readModeInkLockDown",
        "smartTagType",
        "schemaLibrary",
        "shapeDefaults",
        "doNotEmbedSmartTags",
        "decimalSymbol",
        "listSeparator"
    };

    private static void InsertSettingsElementInSchemaOrder(XElement settingsRoot, XElement element)
    {
        var successor = settingsRoot.Elements()
            .FirstOrDefault(candidate =>
                SettingsElementsAfterUpdateFields.Contains(
                    candidate.Name.LocalName,
                    StringComparer.Ordinal));

        if (successor is null)
            settingsRoot.Add(element);
        else
            successor.AddBeforeSelf(element);
    }

    private static void PreflightInput(string input)
    {
        using var archive = ZipFile.OpenRead(input);
        if (archive.GetEntry("word/document.xml") is null)
            throw new InvalidOperationException("Invalid input Word package: word/document.xml is missing.");

        if (archive.GetEntry("word/styles.xml") is null)
            throw new InvalidOperationException("Invalid input Word package: word/styles.xml is missing.");

        if (archive.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
        {
            throw new InvalidOperationException(
                "Digitally signed Office package detected; Template Apply is blocked because package mutation would invalidate the signature.");
        }
    }

    private static TemplateSnapshot LoadTemplateSnapshot(string templatePath, string inputPath)
    {
        using var archive = ZipFile.OpenRead(templatePath);

        var document = ReadXmlEntry(archive, "word/document.xml")
            ?? throw new InvalidOperationException("Template is missing word/document.xml.");

        var styles = ReadXmlEntry(archive, "word/styles.xml")
            ?? throw new InvalidOperationException("Template is missing word/styles.xml.");

        using var sourceArchive = ZipFile.OpenRead(inputPath);
        var originalStyles = ReadXmlEntry(sourceArchive, "word/styles.xml") ?? throw new InvalidOperationException("Source styles missing.");
        var bound = TemplateStyleIdentityBinder.Bind(document, styles, originalStyles);
        document = bound.Document;
        styles = bound.Styles;
        var stylesById = ParseStyles(styles);
        var paragraphStyles = stylesById.Values
            .Where(x => x.Type.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        var defaultParagraphStyle = paragraphStyles.FirstOrDefault(x => x.IsDefault)?.StyleId;

        var headingStyles = new Dictionary<int, string>();
        var canonicalRoles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var paragraphStyleNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var style in paragraphStyles)
        {
            if (!string.IsNullOrWhiteSpace(style.Name))
            {
                var normalizedName = NormalizeToken(style.Name);
                if (!string.IsNullOrWhiteSpace(normalizedName) && !paragraphStyleNames.ContainsKey(normalizedName))
                    paragraphStyleNames[normalizedName] = style.StyleId;
            }

            var heading = HeadingLevel(style);
            if (heading is not null && !headingStyles.ContainsKey(heading.Value))
                headingStyles[heading.Value] = style.StyleId;

            var role = CanonicalRole(style);
            if (!string.IsNullOrWhiteSpace(role) && !canonicalRoles.ContainsKey(role))
                canonicalRoles[role] = style.StyleId;
        }

        return new TemplateSnapshot(
            document,
            styles,
            ReadEntryBytes(archive, "word/theme/theme1.xml"),
            ReadEntryBytes(archive, "word/fontTable.xml"),
            stylesById,
            defaultParagraphStyle,
            headingStyles,
            canonicalRoles,
            paragraphStyleNames);
    }

    private sealed record MappingTarget(string TargetStyleId, string Reason);

    private static IReadOnlyDictionary<string, MappingTarget> BuildParagraphStyleMapping(
        IReadOnlyDictionary<string, StyleInfo> sourceStyles,
        TemplateSnapshot template)
    {
        var result = new Dictionary<string, MappingTarget>(StringComparer.OrdinalIgnoreCase);

        foreach (var source in sourceStyles.Values)
        {
            if (!source.Type.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
                continue;

            if (template.StylesById.TryGetValue(source.StyleId, out var sameId) &&
                sameId.Type.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
            {
                result[source.StyleId] = new MappingTarget(source.StyleId, "same-style-id");
                continue;
            }

            if (!string.IsNullOrWhiteSpace(source.Name))
            {
                var normalizedName = NormalizeToken(source.Name);
                if (template.ParagraphStyleNames.TryGetValue(normalizedName, out var byName))
                {
                    result[source.StyleId] = new MappingTarget(byName, "exact-style-name");
                    continue;
                }
            }

            var heading = HeadingLevel(source);
            if (heading is not null && template.HeadingStyles.TryGetValue(heading.Value, out var byHeading))
            {
                result[source.StyleId] = new MappingTarget(byHeading, $"heading-level-{heading.Value}");
                continue;
            }

            var role = CanonicalRole(source);
            if (!string.IsNullOrWhiteSpace(role) && template.CanonicalRoles.TryGetValue(role, out var byRole))
            {
                result[source.StyleId] = new MappingTarget(byRole, "canonical-role-" + role);
                continue;
            }

            if (source.IsDefault && !string.IsNullOrWhiteSpace(template.DefaultParagraphStyleId))
                result[source.StyleId] = new MappingTarget(template.DefaultParagraphStyleId, "default-paragraph-style");
        }

        return result;
    }

    private static Dictionary<string, int> ApplyParagraphMappings(
        XDocument document,
        IReadOnlyDictionary<string, MappingTarget> mapping)
    {
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var paragraph in document.Descendants(W + "p"))
        {
            var pPr = paragraph.Element(W + "pPr");
            if (pPr is null)
                continue;

            var pStyle = pPr.Element(W + "pStyle");
            if (pStyle is null)
                continue;

            var sourceStyleId = (string?)pStyle.Attribute(W + "val");
            if (string.IsNullOrWhiteSpace(sourceStyleId))
                continue;

            if (!mapping.TryGetValue(sourceStyleId, out var target))
                continue;

            if (sourceStyleId.Equals(target.TargetStyleId, StringComparison.OrdinalIgnoreCase))
                continue;

            pStyle.SetAttributeValue(W + "val", target.TargetStyleId);
            counts.TryGetValue(sourceStyleId, out var currentCount);
            counts[sourceStyleId] = currentCount + 1;
        }

        return counts;
    }

    private static bool MergeTemplateStyles(
        XDocument sourceStyles,
        XDocument templateStyles,
        ref int imported,
        ref int replaced,
        ref int sourceNumberingBindingsPreserved,
        ref int templateNumberingBindingsStripped)
    {
        var sourceRoot = sourceStyles.Root
            ?? throw new InvalidOperationException("Source styles.xml has no root element.");
        var templateRoot = templateStyles.Root
            ?? throw new InvalidOperationException("Template styles.xml has no root element.");

        var changed = false;

        var templateDocDefaults = templateRoot.Element(W + "docDefaults");
        if (templateDocDefaults is not null)
        {
            var sourceDocDefaults = sourceRoot.Element(W + "docDefaults");
            var safeDocDefaults = CreateNumberingSafeDocDefaults(
                templateDocDefaults,
                sourceDocDefaults,
                ref sourceNumberingBindingsPreserved,
                ref templateNumberingBindingsStripped);

            if (sourceDocDefaults is null)
            {
                sourceRoot.AddFirst(safeDocDefaults);
                changed = true;
            }
            else if (!XNode.DeepEquals(sourceDocDefaults, safeDocDefaults))
            {
                sourceDocDefaults.ReplaceWith(safeDocDefaults);
                changed = true;
            }
        }

        var sourceById = sourceRoot
            .Elements(W + "style")
            .Where(x => !string.IsNullOrWhiteSpace((string?)x.Attribute(W + "styleId")))
            .ToDictionary(
                x => ((string?)x.Attribute(W + "styleId"))!,
                x => x,
                StringComparer.OrdinalIgnoreCase);

        foreach (var templateStyle in templateRoot.Elements(W + "style"))
        {
            var styleId = (string?)templateStyle.Attribute(W + "styleId");
            if (string.IsNullOrWhiteSpace(styleId))
                continue;

            sourceById.TryGetValue(styleId, out var sourceStyle);
            var safeTemplateStyle = CreateNumberingSafeTemplateStyle(
                templateStyle,
                sourceStyle,
                ref sourceNumberingBindingsPreserved,
                ref templateNumberingBindingsStripped);

            if (sourceStyle is null)
            {
                sourceRoot.Add(safeTemplateStyle);
                imported++;
                changed = true;
                continue;
            }

            if (XNode.DeepEquals(sourceStyle, safeTemplateStyle))
                continue;

            sourceStyle.ReplaceWith(safeTemplateStyle);
            replaced++;
            changed = true;
        }

        return changed;
    }

    private static XElement CreateNumberingSafeDocDefaults(
        XElement templateDocDefaults,
        XElement? sourceDocDefaults,
        ref int sourceNumberingBindingsPreserved,
        ref int templateNumberingBindingsStripped)
    {
        var safe = new XElement(templateDocDefaults);

        var safePPr = safe
            .Element(W + "pPrDefault")
            ?.Element(W + "pPr");
        var templateNumPr = safePPr?.Element(W + "numPr");
        if (templateNumPr is not null)
        {
            templateNumPr.Remove();
            templateNumberingBindingsStripped++;
        }

        var sourceNumPr = sourceDocDefaults
            ?.Element(W + "pPrDefault")
            ?.Element(W + "pPr")
            ?.Element(W + "numPr");

        if (sourceNumPr is not null)
        {
            var pPrDefault = safe.Element(W + "pPrDefault");
            if (pPrDefault is null)
            {
                pPrDefault = new XElement(W + "pPrDefault");
                safe.Add(pPrDefault);
            }

            safePPr = pPrDefault.Element(W + "pPr");
            if (safePPr is null)
            {
                safePPr = new XElement(W + "pPr");
                pPrDefault.Add(safePPr);
            }

            InsertInSchemaOrder(safePPr, new XElement(sourceNumPr), ParagraphPropertyOrder);
            sourceNumberingBindingsPreserved++;
        }

        return safe;
    }

    private static XElement CreateNumberingSafeTemplateStyle(
        XElement templateStyle,
        XElement? sourceStyle,
        ref int sourceNumberingBindingsPreserved,
        ref int templateNumberingBindingsStripped)
    {
        var safe = new XElement(templateStyle);
        var safePPr = safe.Element(W + "pPr");
        var templateNumPr = safePPr?.Element(W + "numPr");

        if (templateNumPr is not null)
        {
            templateNumPr.Remove();
            templateNumberingBindingsStripped++;
        }

        var sourceNumPr = sourceStyle
            ?.Element(W + "pPr")
            ?.Element(W + "numPr");

        if (sourceNumPr is not null)
        {
            if (safePPr is null)
            {
                safePPr = new XElement(W + "pPr");
                InsertInSchemaOrder(safe, safePPr, StylePropertyOrder);
            }

            InsertInSchemaOrder(safePPr, new XElement(sourceNumPr), ParagraphPropertyOrder);
            sourceNumberingBindingsPreserved++;
        }
        else if (safePPr is not null && !safePPr.HasAttributes && !safePPr.Elements().Any())
        {
            safePPr.Remove();
        }

        return safe;
    }

    private static int ApplySectionGeometry(XDocument sourceDocument, XDocument templateDocument)
    {
        var sourceSections = sourceDocument.Descendants(W + "sectPr").ToArray();
        var templateSections = templateDocument.Descendants(W + "sectPr").ToArray();

        if (sourceSections.Length == 0)
            throw new InvalidOperationException("Input document has no section properties.");

        if (templateSections.Length == 0)
            throw new InvalidOperationException("Template document has no section properties.");

        var changedSections = 0;

        for (var i = 0; i < sourceSections.Length; i++)
        {
            var target = sourceSections[i];
            var sourceTemplate = templateSections[Math.Min(i, templateSections.Length - 1)];
            var changed = false;

            foreach (var localName in new[] { "pgSz", "pgMar", "cols" })
            {
                if (ApplySectionChild(target, sourceTemplate, localName))
                    changed = true;
            }

            if (changed)
                changedSections++;
        }

        return changedSections;
    }

    private static bool ApplySectionChild(XElement targetSection, XElement templateSection, string localName)
    {
        var current = targetSection.Element(W + localName);
        var expected = templateSection.Element(W + localName);

        if (expected is null)
        {
            if (current is null)
                return false;

            current.Remove();
            return true;
        }

        if (current is not null && XNode.DeepEquals(current, expected))
            return false;

        var clone = new XElement(expected);
        if (current is not null)
        {
            current.ReplaceWith(clone);
            return true;
        }

        InsertInSchemaOrder(targetSection, clone, SectionPropertyOrder);
        return true;
    }

    private static bool ReplaceBinaryFromTemplateIfTargetExists(
        ZipArchive outputArchive,
        string entryName,
        byte[]? templateBytes,
        ISet<string> mutatedEntries)
    {
        if (templateBytes is null)
            return false;

        var targetEntry = outputArchive.GetEntry(entryName);
        if (targetEntry is null)
            return false;

        using (var stream = targetEntry.Open())
        using (var memory = new MemoryStream())
        {
            stream.CopyTo(memory);
            if (memory.ToArray().AsSpan().SequenceEqual(templateBytes))
                return false;
        }

        targetEntry.Delete();
        var replacement = outputArchive.CreateEntry(entryName, CompressionLevel.Optimal);
        using (var stream = replacement.Open())
            stream.Write(templateBytes, 0, templateBytes.Length);

        mutatedEntries.Add(entryName);
        return true;
    }

    private static IReadOnlyDictionary<string, StyleInfo> ParseStyles(XDocument styles)
    {
        var result = new Dictionary<string, StyleInfo>(StringComparer.OrdinalIgnoreCase);

        foreach (var style in styles.Descendants(W + "style"))
        {
            var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(styleId))
                continue;

            var type = ((string?)style.Attribute(W + "type") ?? string.Empty).Trim();
            var name = ((string?)style.Element(W + "name")?.Attribute(W + "val"))?.Trim();
            var isDefault = IsOnAttribute(style.Attribute(W + "default"));

            int? outline = null;
            var outlineValue = (string?)style.Element(W + "pPr")?.Element(W + "outlineLvl")?.Attribute(W + "val");
            if (int.TryParse(outlineValue, out var parsedOutline))
                outline = parsedOutline;

            result[styleId] = new StyleInfo(styleId, type, name, isDefault, outline);
        }

        return result;
    }

    private static int? HeadingLevel(StyleInfo style)
    {
        if (style.OutlineLevel is int outlineLevel && outlineLevel is >= 0 and <= 8)
            return outlineLevel + 1;

        foreach (var value in new[] { style.StyleId, style.Name })
        {
            if (string.IsNullOrWhiteSpace(value))
                continue;

            var english = Regex.Match(value, @"heading\s*([1-9])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (english.Success && int.TryParse(english.Groups[1].Value, out var englishLevel))
                return englishLevel;

            var persian = Regex.Match(value, @"عنوان\s*([1-9۱-۹])", RegexOptions.CultureInvariant);
            if (persian.Success)
            {
                var digit = NormalizePersianDigit(persian.Groups[1].Value[0]);
                if (digit is >= '1' and <= '9')
                    return digit - '0';
            }
        }

        return null;
    }

    private static string? CanonicalRole(StyleInfo style)
    {
        var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            NormalizeToken(style.StyleId),
            NormalizeToken(style.Name)
        };

        if (tokens.Contains("title") || tokens.Contains("عنوان"))
            return "title";

        if (tokens.Contains("subtitle") || tokens.Contains("زیرعنوان"))
            return "subtitle";

        if (tokens.Contains("normal") || tokens.Contains("bodytext") || tokens.Contains("متنعادی"))
            return "body";

        var heading = HeadingLevel(style);
        return heading is null ? null : "heading" + heading.Value;
    }

    private static string NormalizeToken(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var builder = new StringBuilder(value.Length);
        foreach (var ch in value.Trim().ToLowerInvariant())
        {
            if (char.IsLetterOrDigit(ch))
                builder.Append(ch);
        }

        return builder.ToString();
    }

    private static char NormalizePersianDigit(char value) => value switch
    {
        '۱' => '1',
        '۲' => '2',
        '۳' => '3',
        '۴' => '4',
        '۵' => '5',
        '۶' => '6',
        '۷' => '7',
        '۸' => '8',
        '۹' => '9',
        _ => value
    };

    private static bool IsOnAttribute(XAttribute? attribute)
    {
        if (attribute is null)
            return false;

        var value = attribute.Value.Trim();
        if (string.IsNullOrWhiteSpace(value))
            return true;

        return !value.Equals("0", StringComparison.OrdinalIgnoreCase) &&
               !value.Equals("false", StringComparison.OrdinalIgnoreCase) &&
               !value.Equals("off", StringComparison.OrdinalIgnoreCase);
    }

    private static TemplateApplyInventory InspectPackage(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var storyParts = GetStoryPartNames(archive);

        int paragraphs = 0;
        int tables = 0;
        int drawings = 0;
        int fieldInstructions = 0;
        int textNodes = 0;
        int revisions = 0;
        int sections = 0;
        int contentControls = 0;
        int bookmarks = 0;
        int hyperlinks = 0;

        var textPayload = new StringBuilder();
        var fieldPayload = new StringBuilder();

        foreach (var partName in storyParts)
        {
            var document = LoadXmlEntry(archive, partName);
            paragraphs += document.Descendants(W + "p").Count();
            tables += document.Descendants(W + "tbl").Count();
            drawings += document.Descendants(W + "drawing").Count();
            sections += document.Descendants(W + "sectPr").Count();
            contentControls += document.Descendants(W + "sdt").Count();
            bookmarks += document.Descendants(W + "bookmarkStart").Count();
            hyperlinks += document.Descendants(W + "hyperlink").Count();
            revisions += CountRevisionElements(document);

            foreach (var text in document.Descendants(W + "t"))
            {
                textNodes++;
                textPayload.Append(partName).Append('\u001F').Append((string?)text ?? string.Empty).Append('\u001E');
            }

            foreach (var instr in document.Descendants(W + "instrText"))
            {
                fieldInstructions++;
                fieldPayload.Append(partName).Append('\u001F').Append((string?)instr ?? string.Empty).Append('\u001E');
            }

            foreach (var simple in document.Descendants(W + "fldSimple"))
            {
                fieldInstructions++;
                fieldPayload.Append(partName).Append('\u001F').Append((string?)simple.Attribute(W + "instr") ?? string.Empty).Append('\u001E');
            }
        }

        var styleDefinitions = 0;
        var stylesEntry = archive.GetEntry("word/styles.xml");
        if (stylesEntry is not null)
        {
            using var stream = stylesEntry.Open();
            var styles = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            styleDefinitions = styles.Descendants(W + "style").Count();
        }

        var mediaPayload = new StringBuilder();
        foreach (var media in archive.Entries
                     .Where(e => e.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase) &&
                                 !e.FullName.EndsWith("/", StringComparison.Ordinal))
                     .OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase))
        {
            using var stream = media.Open();
            mediaPayload.Append(media.FullName)
                .Append('\u001F')
                .Append(Sha256Stream(stream))
                .Append('\u001E');
        }

        return new TemplateApplyInventory(
            archive.Entries.Count,
            storyParts.Count,
            paragraphs,
            tables,
            drawings,
            fieldInstructions,
            textNodes,
            revisions,
            sections,
            contentControls,
            bookmarks,
            hyperlinks,
            archive.Entries.Count(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)),
            archive.Entries.Count(e => e.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase) &&
                                       !e.FullName.EndsWith("/", StringComparison.Ordinal)),
            styleDefinitions,
            Sha256Text(textPayload.ToString()),
            Sha256Text(fieldPayload.ToString()),
            Sha256Text(mediaPayload.ToString()));
    }

    private static List<string> CompareInventory(TemplateApplyInventory before, TemplateApplyInventory after)
    {
        var issues = new List<string>();

        void Equal(string name, int a, int b)
        {
            if (a != b)
                issues.Add($"Package inventory changed unexpectedly: {name} {a} -> {b}.");
        }

        Equal(nameof(before.PackageEntries), before.PackageEntries, after.PackageEntries);
        Equal(nameof(before.StoryParts), before.StoryParts, after.StoryParts);
        Equal(nameof(before.Paragraphs), before.Paragraphs, after.Paragraphs);
        Equal(nameof(before.Tables), before.Tables, after.Tables);
        Equal(nameof(before.Drawings), before.Drawings, after.Drawings);
        Equal(nameof(before.FieldInstructions), before.FieldInstructions, after.FieldInstructions);
        Equal(nameof(before.TextNodes), before.TextNodes, after.TextNodes);
        Equal(nameof(before.RevisionElements), before.RevisionElements, after.RevisionElements);
        Equal(nameof(before.SectionProperties), before.SectionProperties, after.SectionProperties);
        Equal(nameof(before.ContentControls), before.ContentControls, after.ContentControls);
        Equal(nameof(before.Bookmarks), before.Bookmarks, after.Bookmarks);
        Equal(nameof(before.Hyperlinks), before.Hyperlinks, after.Hyperlinks);
        Equal(nameof(before.RelationshipsParts), before.RelationshipsParts, after.RelationshipsParts);
        Equal(nameof(before.MediaEntries), before.MediaEntries, after.MediaEntries);

        if (after.StyleDefinitions < before.StyleDefinitions)
            issues.Add($"Style definitions decreased unexpectedly: {before.StyleDefinitions} -> {after.StyleDefinitions}.");

        if (!before.TextPayloadSha256.Equals(after.TextPayloadSha256, StringComparison.OrdinalIgnoreCase))
            issues.Add("Visible Word text payload changed; Template Apply is formatting-only.");

        if (!before.FieldInstructionSha256.Equals(after.FieldInstructionSha256, StringComparison.OrdinalIgnoreCase))
            issues.Add("Field instruction payload changed; fields must be preserved.");

        if (!before.MediaPayloadSha256.Equals(after.MediaPayloadSha256, StringComparison.OrdinalIgnoreCase))
            issues.Add("Media payload changed; images/media must be preserved.");

        return issues;
    }

    private static IEnumerable<string> VerifyTemplateApplication(
        string outputPath,
        TemplateSnapshot template)
    {
        var issues = new List<string>();

        using var archive = ZipFile.OpenRead(outputPath);
        var outputStyles = ReadXmlEntry(archive, "word/styles.xml");
        var outputDocument = ReadXmlEntry(archive, "word/document.xml");

        if (outputStyles is null)
        {
            issues.Add("Output is missing word/styles.xml.");
            return issues;
        }

        if (outputDocument is null)
        {
            issues.Add("Output is missing word/document.xml.");
            return issues;
        }

        var outputById = outputStyles
            .Descendants(W + "style")
            .Where(x => !string.IsNullOrWhiteSpace((string?)x.Attribute(W + "styleId")))
            .ToDictionary(
                x => ((string?)x.Attribute(W + "styleId"))!,
                x => x,
                StringComparer.OrdinalIgnoreCase);

        foreach (var templateStyle in template.Styles.Descendants(W + "style"))
        {
            var styleId = (string?)templateStyle.Attribute(W + "styleId");
            if (string.IsNullOrWhiteSpace(styleId))
                continue;

            if (!outputById.TryGetValue(styleId, out var outputStyle))
            {
                issues.Add($"Template style missing from output: {styleId}");
                continue;
            }

            if (!TemplateStyleEqualsUnderNumberingPolicy(outputStyle, templateStyle))
                issues.Add($"Template style definition mismatch in output outside numbering-safety policy: {styleId}");
        }

        var outputSections = outputDocument.Descendants(W + "sectPr").ToArray();
        var templateSections = template.Document.Descendants(W + "sectPr").ToArray();

        for (var i = 0; i < outputSections.Length; i++)
        {
            var expectedSection = templateSections[Math.Min(i, templateSections.Length - 1)];
            foreach (var localName in new[] { "pgSz", "pgMar", "cols" })
            {
                var actual = outputSections[i].Element(W + localName);
                var expected = expectedSection.Element(W + localName);
                if (!NullableDeepEquals(actual, expected))
                    issues.Add($"Section {i + 1} template geometry mismatch: {localName}");
            }
        }

        return issues;
    }

    private static bool TemplateStyleEqualsUnderNumberingPolicy(
        XElement outputStyle,
        XElement templateStyle)
    {
        var normalizedOutput = new XElement(outputStyle);
        var normalizedTemplate = new XElement(templateStyle);

        normalizedOutput
            .Element(W + "pPr")
            ?.Element(W + "numPr")
            ?.Remove();

        normalizedTemplate
            .Element(W + "pPr")
            ?.Element(W + "numPr")
            ?.Remove();

        StripFoundationOwnedStyleProperties(normalizedOutput);
        StripFoundationOwnedStyleProperties(normalizedTemplate);

        RemoveEmptyParagraphProperties(normalizedOutput);
        RemoveEmptyParagraphProperties(normalizedTemplate);

        return XNode.DeepEquals(normalizedOutput, normalizedTemplate);
    }

    private static void StripFoundationOwnedStyleProperties(XElement style)
    {
        var fullFontFamilyOwned = FoundationOwnsFullFontFamily(style);
        var automaticListLayoutOwned = FoundationOwnsAutomaticListParagraphLayout(style);
        var automaticListEntryTabsOwned = FoundationOwnsAutomaticListEntryTabs(style);
        var rPr = style.Element(W + "rPr");
        var rFonts = rPr?.Element(W + "rFonts");
        if (rFonts is not null)
        {
            var ownedFontAttributes = fullFontFamilyOwned
                ? new[] { "ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme" }
                : new[] { "cs", "cstheme" };

            foreach (var attributeName in ownedFontAttributes)
                rFonts.Attribute(W + attributeName)?.Remove();

            if (!rFonts.HasAttributes)
                rFonts.Remove();
        }

        var lang = rPr?.Element(W + "lang");
        if (lang is not null)
        {
            lang.Attribute(W + "val")?.Remove();
            lang.Attribute(W + "bidi")?.Remove();
            if (!lang.HasAttributes)
                lang.Remove();
        }

        // Keep automatic-list rPr attached until all Foundation-owned typography has been stripped.
        // Non-automatic styles may still remove an empty rPr immediately.
        if (!automaticListLayoutOwned && rPr is not null && !rPr.HasAttributes && !rPr.Elements().Any())
            rPr.Remove();

        if (automaticListLayoutOwned)
        {
            // Foundation owns automatic-list typography as well as paragraph direction/layout.
            // Ignore only those exact properties during template-integrity comparison; all other
            // template style properties remain protected by the final DeepEquals.
            if (rPr is not null)
            {
                rPr.Element(W + "sz")?.Remove();
                rPr.Element(W + "szCs")?.Remove();
                rPr.Element(W + "b")?.Remove();
                rPr.Element(W + "bCs")?.Remove();
                if (!rPr.HasAttributes && !rPr.Elements().Any())
                    rPr.Remove();
            }

            // The shared Persian Foundation deliberately disables automatic style redefinition
            // for native TOC/Table-of-Figures styles so Word field regeneration cannot overwrite
            // the accepted Persian list contract. Treat that suppression as Foundation-owned
            // during Template Apply integrity comparison, just like list bidi/alignment.
            style.Element(W + "autoRedefine")?.Remove();
            var pPr = style.Element(W + "pPr");
            pPr?.Element(W + "bidi")?.Remove();
            pPr?.Element(W + "jc")?.Remove();

            // Automatic-list entry page-number tab stops are wholly Foundation-owned:
            // EnsureDottedPageNumberTab replaces the child tab set with exactly one
            // right-aligned dotted leader at position 7938. Title styles do not enter this lane.
            if (automaticListEntryTabsOwned)
                pPr?.Element(W + "tabs")?.Remove();
        }
    }

    private static bool FoundationOwnsAutomaticListEntryTabs(XElement style)
    {
        var type = ((string?)style.Attribute(W + "type") ?? string.Empty).Trim();
        if (!type.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
            return false;

        var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Trim();
        var name = ((string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty).Trim();
        var outline = ((string?)style.Element(W + "pPr")?.Element(W + "outlineLvl")?.Attribute(W + "val") ?? string.Empty).Trim();
        var info = PersianParagraphSemanticResolver.ClassifyStyle(styleId, name, outline);
        var compactId = styleId.Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty);
        var compactName = name.Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty);
        return info.TocEntry ||
               compactId.Equals("TableofFigures", StringComparison.OrdinalIgnoreCase) ||
               compactName.Equals("tableoffigures", StringComparison.OrdinalIgnoreCase);
    }

    private static bool FoundationOwnsFullFontFamily(XElement style)
    {
        var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Trim();
        if (styleId.Equals("Hyperlink", StringComparison.OrdinalIgnoreCase))
            return true;

        return FoundationOwnsAutomaticListParagraphLayout(style);
    }

    private static bool FoundationOwnsAutomaticListParagraphLayout(XElement style)
    {
        var type = ((string?)style.Attribute(W + "type") ?? string.Empty).Trim();
        if (!type.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
            return false;

        var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Trim();
        var name = ((string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty).Trim();
        var outline = ((string?)style.Element(W + "pPr")?.Element(W + "outlineLvl")?.Attribute(W + "val") ?? string.Empty).Trim();
        var info = PersianParagraphSemanticResolver.ClassifyStyle(styleId, name, outline);
        var compactId = styleId.Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty);
        var compactName = name.Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty);
        return info.TocEntry || info.TocTitle ||
               compactId.Equals("TableofFigures", StringComparison.OrdinalIgnoreCase) ||
               compactName.Equals("tableoffigures", StringComparison.OrdinalIgnoreCase) ||
               compactName.Equals("عنوانفهرست", StringComparison.Ordinal);
    }

    private static void RemoveEmptyParagraphProperties(XElement style)
    {
        var pPr = style.Element(W + "pPr");
        if (pPr is not null && !pPr.HasAttributes && !pPr.Elements().Any())
            pPr.Remove();
    }

    private static bool NullableDeepEquals(XElement? left, XElement? right)
    {
        if (left is null && right is null)
            return true;

        if (left is null || right is null)
            return false;

        return XNode.DeepEquals(left, right);
    }

    private static List<string> GetStoryPartNames(ZipArchive archive)
    {
        var names = archive.Entries
            .Select(x => x.FullName)
            .Where(IsStoryPart)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (!names.Contains("word/document.xml", StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("word/document.xml was not found.");

        return names;
    }

    private static bool IsStoryPart(string name)
    {
        if (name.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("word/footnotes.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("word/endnotes.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static int CountRevisionElements(XDocument document)
        => document.Descendants(W + "ins").Count()
         + document.Descendants(W + "del").Count()
         + document.Descendants(W + "moveFrom").Count()
         + document.Descendants(W + "moveTo").Count();

    private static Dictionary<string, string> ReadEntryHashes(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        return archive.Entries
            .Where(e => !e.FullName.EndsWith("/", StringComparison.Ordinal))
            .ToDictionary(
                e => e.FullName,
                e =>
                {
                    using var stream = e.Open();
                    return Sha256Stream(stream);
                },
                StringComparer.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> CompareImmutableEntries(
        string outputPath,
        IReadOnlyDictionary<string, string> beforeHashes,
        ISet<string> mutatedEntries)
    {
        var issues = new List<string>();
        using var archive = ZipFile.OpenRead(outputPath);

        var afterNames = archive.Entries
            .Where(e => !e.FullName.EndsWith("/", StringComparison.Ordinal))
            .Select(e => e.FullName)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var pair in beforeHashes)
        {
            if (mutatedEntries.Contains(pair.Key))
                continue;

            if (!afterNames.Contains(pair.Key))
            {
                issues.Add($"Immutable package entry disappeared: {pair.Key}");
                continue;
            }

            var entry = archive.GetEntry(pair.Key);
            if (entry is null)
            {
                issues.Add($"Immutable package entry could not be reopened: {pair.Key}");
                continue;
            }

            using var stream = entry.Open();
            var actual = Sha256Stream(stream);
            if (!actual.Equals(pair.Value, StringComparison.OrdinalIgnoreCase))
                issues.Add($"Immutable package entry changed unexpectedly: {pair.Key}");
        }

        return issues;
    }

    private static List<string> VerifyNumberingSafety(string inputPath, string outputPath)
    {
        var issues = new List<string>();

        using var inputArchive = ZipFile.OpenRead(inputPath);
        using var outputArchive = ZipFile.OpenRead(outputPath);

        var inputNumbering = inputArchive.GetEntry("word/numbering.xml");
        var outputNumbering = outputArchive.GetEntry("word/numbering.xml");

        if ((inputNumbering is null) != (outputNumbering is null))
        {
            issues.Add("Numbering safety mismatch: word/numbering.xml presence changed.");
        }
        else if (inputNumbering is not null && outputNumbering is not null)
        {
            using var inputStream = inputNumbering.Open();
            using var outputStream = outputNumbering.Open();
            var inputHash = Sha256Stream(inputStream);
            var outputHash = Sha256Stream(outputStream);
            if (!inputHash.Equals(outputHash, StringComparison.OrdinalIgnoreCase))
                issues.Add("Numbering safety mismatch: source word/numbering.xml changed.");
        }

        var inputStyles = ReadXmlEntry(inputArchive, "word/styles.xml");
        var outputStyles = ReadXmlEntry(outputArchive, "word/styles.xml");
        if (inputStyles is null || outputStyles is null)
        {
            issues.Add("Numbering safety mismatch: styles.xml could not be loaded.");
        }
        else
        {
            var inputRoot = inputStyles.Root;
            var outputRoot = outputStyles.Root;
            if (inputRoot is null || outputRoot is null)
            {
                issues.Add("Numbering safety mismatch: styles.xml root is missing.");
            }
            else
            {
                var inputById = inputRoot
                    .Elements(W + "style")
                    .Where(x => !string.IsNullOrWhiteSpace((string?)x.Attribute(W + "styleId")))
                    .ToDictionary(
                        x => ((string?)x.Attribute(W + "styleId"))!,
                        x => x,
                        StringComparer.OrdinalIgnoreCase);

                var outputById = outputRoot
                    .Elements(W + "style")
                    .Where(x => !string.IsNullOrWhiteSpace((string?)x.Attribute(W + "styleId")))
                    .ToDictionary(
                        x => ((string?)x.Attribute(W + "styleId"))!,
                        x => x,
                        StringComparer.OrdinalIgnoreCase);

                foreach (var pair in inputById)
                {
                    if (!outputById.TryGetValue(pair.Key, out var outputStyle))
                    {
                        issues.Add($"Numbering safety mismatch: source style disappeared: {pair.Key}");
                        continue;
                    }

                    var beforeSignature = NumberingBindingSignature(
                        pair.Value.Element(W + "pPr")?.Element(W + "numPr"));
                    var afterSignature = NumberingBindingSignature(
                        outputStyle.Element(W + "pPr")?.Element(W + "numPr"));

                    if (!beforeSignature.Equals(afterSignature, StringComparison.Ordinal))
                        issues.Add($"Numbering safety mismatch: source style numbering changed: {pair.Key}");
                }

                foreach (var pair in outputById)
                {
                    if (inputById.ContainsKey(pair.Key))
                        continue;

                    if (!string.IsNullOrEmpty(NumberingBindingSignature(
                        pair.Value.Element(W + "pPr")?.Element(W + "numPr"))))
                    {
                        issues.Add($"Numbering safety mismatch: imported style carries template numbering: {pair.Key}");
                    }
                }

                var inputDefaultSignature = NumberingBindingSignature(
                    inputRoot.Element(W + "docDefaults")
                        ?.Element(W + "pPrDefault")
                        ?.Element(W + "pPr")
                        ?.Element(W + "numPr"));
                var outputDefaultSignature = NumberingBindingSignature(
                    outputRoot.Element(W + "docDefaults")
                        ?.Element(W + "pPrDefault")
                        ?.Element(W + "pPr")
                        ?.Element(W + "numPr"));

                if (!inputDefaultSignature.Equals(outputDefaultSignature, StringComparison.Ordinal))
                    issues.Add("Numbering safety mismatch: docDefaults paragraph numbering changed.");
            }
        }

        var inputDocument = ReadXmlEntry(inputArchive, "word/document.xml");
        var outputDocument = ReadXmlEntry(outputArchive, "word/document.xml");
        if (inputDocument is null || outputDocument is null)
        {
            issues.Add("Numbering safety mismatch: document.xml could not be loaded.");
        }
        else
        {
            var inputParagraphs = inputDocument.Descendants(W + "p").ToArray();
            var outputParagraphs = outputDocument.Descendants(W + "p").ToArray();

            if (inputParagraphs.Length != outputParagraphs.Length)
            {
                issues.Add(
                    $"Numbering safety mismatch: paragraph count changed {inputParagraphs.Length}->{outputParagraphs.Length}.");
            }
            else
            {
                for (var i = 0; i < inputParagraphs.Length; i++)
                {
                    var beforeSignature = NumberingBindingSignature(
                        inputParagraphs[i].Element(W + "pPr")?.Element(W + "numPr"));
                    var afterSignature = NumberingBindingSignature(
                        outputParagraphs[i].Element(W + "pPr")?.Element(W + "numPr"));
                    if (!beforeSignature.Equals(afterSignature, StringComparison.Ordinal))
                    {
                        issues.Add($"Numbering safety mismatch: direct paragraph numbering changed at index {i}.");
                        if (issues.Count >= 16)
                            break;
                    }
                }
            }
        }

        return issues;
    }

    private static string NumberingBindingSignature(XElement? numPr)
    {
        if (numPr is null)
            return string.Empty;

        var tokens = new List<string>();
        foreach (var element in numPr.DescendantsAndSelf())
        {
            var attributes = element
                .Attributes()
                .OrderBy(x => x.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(x => x.Name.LocalName, StringComparer.Ordinal)
                .Select(x => $"{x.Name.NamespaceName}#{x.Name.LocalName}={x.Value}");

            tokens.Add(
                $"{element.Name.NamespaceName}#{element.Name.LocalName}[{string.Join(",", attributes)}]");
        }

        return string.Join("/", tokens);
    }

    private static XDocument LoadXmlEntry(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName)
            ?? throw new InvalidOperationException($"Required XML package entry is missing: {entryName}");

        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static XDocument? ReadXmlEntry(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName);
        if (entry is null)
            return null;

        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static byte[]? ReadEntryBytes(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName);
        if (entry is null)
            return null;

        using var stream = entry.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static void ReplaceXmlEntry(ZipArchive archive, string entryName, XDocument document)
    {
        var old = archive.GetEntry(entryName)
            ?? throw new InvalidOperationException($"Cannot replace missing XML entry: {entryName}");
        old.Delete();

        var replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var stream = replacement.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false,
            Indent = false,
            CloseOutput = false
        });
        document.Save(writer);
    }

    private static void InsertInSchemaOrder(XElement parent, XElement child, IReadOnlyList<string> order)
    {
        var targetIndex = IndexOf(order, child.Name.LocalName);
        if (targetIndex < 0)
        {
            parent.Add(child);
            return;
        }

        foreach (var existing in parent.Elements())
        {
            var existingIndex = IndexOf(order, existing.Name.LocalName);
            if (existingIndex >= 0 && existingIndex > targetIndex)
            {
                existing.AddBeforeSelf(child);
                return;
            }
        }

        parent.Add(child);
    }

    private static int IndexOf(IReadOnlyList<string> values, string value)
    {
        for (var i = 0; i < values.Count; i++)
        {
            if (values[i].Equals(value, StringComparison.Ordinal))
                return i;
        }

        return -1;
    }

    private static string Sha256File(string path)
    {
        using var stream = File.OpenRead(path);
        return Sha256Stream(stream);
    }

    private static string Sha256Stream(Stream stream) =>
        Convert.ToHexString(SHA256.HashData(stream));

    private static string Sha256Text(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));


    private static string CreateTransactionalTempPath(string finalOutput)
    {
        var directory = Path.GetDirectoryName(finalOutput)
            ?? throw new InvalidOperationException("Output directory could not be resolved.");
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, ".wordagent-txn-" + Guid.NewGuid().ToString("N") + ".tmp");
    }

    private static void PauseForTransactionTest()
    {
        var raw = Environment.GetEnvironmentVariable("WORDAGENT_TEST_TRANSACTION_PAUSE_MS");
        if (!int.TryParse(raw, out var milliseconds) || milliseconds <= 0)
            return;

        milliseconds = Math.Min(milliseconds, 120000);
        System.Threading.Thread.Sleep(milliseconds);
    }

    private static void CommitTransactionalOutput(string transaction, string finalOutput)
    {
        try
        {
            if (File.Exists(finalOutput))
                throw new IOException("Final output appeared during the transaction; refusing overwrite.");

            File.Move(transaction, finalOutput);
        }
        catch
        {
            TryDelete(transaction);
            throw;
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // Best effort cleanup only.
        }
    }
}
