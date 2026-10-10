using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using System.Xml;
using System.Xml.Linq;
using WordAgent.Core;

namespace WordAgent.WordInterop;

public sealed record PersianDocumentFoundationOptions(
    string? ProfileId = null,
    bool NormalizeBodyAlignment = true,
    bool NormalizeHeadingAlignment = true,
    bool NormalizeToc = true);

public sealed record PersianDocumentFoundationResult(
    string File,
    string? ProfileId,
    int StoryPartsExamined,
    int PersianParagraphs,
    int BodyParagraphsJustified,
    int HeadingParagraphsRightAligned,
    int TitleParagraphsCentered,
    int TocParagraphsNormalized,
    int PersianRunsNormalized,
    int StyleDefinitionsNormalized,
    bool DocDefaultsNormalized,
    IReadOnlyList<string> MutatedEntries);

public sealed class PersianDocumentFoundationService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly string[] PPrOrder =
    [
        "pStyle","keepNext","keepLines","pageBreakBefore","framePr","widowControl","numPr",
        "suppressLineNumbers","pBdr","shd","tabs","suppressAutoHyphens","kinsoku","wordWrap",
        "overflowPunct","topLinePunct","autoSpaceDE","autoSpaceDN","bidi","adjustRightInd",
        "snapToGrid","spacing","ind","contextualSpacing","mirrorIndents","suppressOverlap","jc",
        "textDirection","textAlignment","textboxTightWrap","outlineLvl","divId","cnfStyle",
        "rPr","sectPr","pPrChange"
    ];

    private static readonly string[] RPrOrder =
    [
        "rStyle","rFonts","b","bCs","i","iCs","caps","smallCaps","strike","dstrike","outline",
        "shadow","emboss","imprint","noProof","snapToGrid","vanish","webHidden","color","spacing",
        "w","kern","position","sz","szCs","highlight","u","effect","bdr","shd","fitText",
        "vertAlign","rtl","cs","em","lang","eastAsianLayout","specVanish","oMath","rPrChange"
    ];

    private static readonly string[] StyleOrder =
    [
        "name","aliases","basedOn","next","link","autoRedefine","hidden","uiPriority",
        "semiHidden","unhideWhenUsed","qFormat","locked","personal","personalCompose",
        "personalReply","rsid","pPr","rPr","tblPr","trPr","tcPr"
    ];

    private static readonly string[] TblPrOrder =
    [
        "tblStyle","tblpPr","tblOverlap","bidiVisual","tblStyleRowBandSize","tblStyleColBandSize",
        "tblW","jc","tblCellSpacing","tblInd","tblBorders","shd","tblLayout","tblCellMar","tblLook",
        "tblCaption","tblDescription","tblPrChange"
    ];

    public PersianDocumentFoundationResult ApplyInPlace(
        string path,
        PersianDocumentFoundationOptions? options = null)
    {
        options ??= new PersianDocumentFoundationOptions();
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Persian Foundation path is required.", nameof(path));

        var filePath = Path.GetFullPath(path);
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Persian Foundation input package was not found.", filePath);

        using var preflight = ZipFile.OpenRead(filePath);
        if (preflight.GetEntry("word/document.xml") is null)
            throw new InvalidOperationException("Persian Foundation requires word/document.xml.");
        if (preflight.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException(
                "Digitally signed Office package detected; Persian Foundation is blocked because mutation would invalidate the signature.");
        preflight.Dispose();

        var mutatedEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int storyPartsExamined = 0;
        int persianParagraphs = 0;
        int bodyParagraphsJustified = 0;
        int headingParagraphsRightAligned = 0;
        int titleParagraphsCentered = 0;
        int tocParagraphsNormalized = 0;
        int persianRunsNormalized = 0;
        int styleDefinitionsNormalized = 0;
        bool docDefaultsNormalized = false;

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);

        var styleInfo = new Dictionary<string, PersianSemanticStyleInfo>(StringComparer.OrdinalIgnoreCase);
        var stylesEntry = archive.GetEntry("word/styles.xml");
        if (stylesEntry is not null)
        {
            var styles = LoadXmlEntry(archive, "word/styles.xml");
            styleInfo = BuildPersianSemanticStyleInfo(styles);

            var styleChanged = false;
            if (NormalizeDocumentDefaults(styles))
            {
                docDefaultsNormalized = true;
                styleChanged = true;
            }

            var hasTocStyleDefinitions =
                styleInfo.Values.Any(info => info.TocEntry || info.TocTitle) ||
                styles.Descendants(W + "style").Any(style =>
                {
                    var candidateStyleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Trim();
                    var candidateStyleName = ((string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty).Trim();
                    return IsAutomaticListEntryStyle(candidateStyleId, candidateStyleName) ||
                           IsAutomaticListTitleStyle(candidateStyleId, candidateStyleName);
                });

            foreach (var style in styles.Descendants(W + "style"))
            {
                var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Trim();
                var styleName = ((string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty).Trim();
                var automaticListEntryStyle = IsAutomaticListEntryStyle(styleId, styleName);
                var automaticListTitleStyle = IsAutomaticListTitleStyle(styleId, styleName);
                var normalizeTocHyperlink =
                    options.NormalizeToc &&
                    hasTocStyleDefinitions &&
                    styleId.Equals("Hyperlink", StringComparison.OrdinalIgnoreCase);

                if (!styleInfo.TryGetValue(styleId, out var info))
                {
                    if (!normalizeTocHyperlink)
                        continue;

                    // Word regenerates native TOC entry text inside Hyperlink character-style runs.
                    // If Hyperlink keeps Times New Roman for ascii/hAnsi, a Word-native TOC update
                    // can reintroduce mixed fonts even when TOC1/2/3 are B Nazanin.
                    var hyperlinkRPr = EnsureChild(style, "rPr", StyleOrder);
                    var hyperlinkChanged = false;
                    if (EnsureFullFontFamily(hyperlinkRPr, "B Nazanin")) hyperlinkChanged = true;
                    if (EnsurePersianLanguage(hyperlinkRPr, setPrimaryLanguage: true)) hyperlinkChanged = true;

                    if (hyperlinkChanged)
                    {
                        styleDefinitionsNormalized++;
                        styleChanged = true;
                    }

                    continue;
                }

                string? targetFont = null;
                if (info.TocEntry || info.TocTitle || automaticListEntryStyle || automaticListTitleStyle)
                    targetFont = "B Nazanin";
                else if (info.Title || info.Chapter || info.Heading)
                    targetFont = "B Titr";
                else if (info.Body || info.Caption)
                    targetFont = "B Nazanin";

                var changedThisStyle = false;
                if (targetFont is not null)
                {
                    var rPr = EnsureChild(style, "rPr", StyleOrder);
                    if (info.TocEntry || info.TocTitle || automaticListEntryStyle || automaticListTitleStyle || normalizeTocHyperlink)
                    {
                        if (EnsureFullFontFamily(rPr, "B Nazanin")) changedThisStyle = true;
                    }
                    else if (EnsureComplexScriptFont(rPr, targetFont))
                    {
                        changedThisStyle = true;
                    }
                    if (EnsurePersianLanguage(rPr, setPrimaryLanguage: true)) changedThisStyle = true;
                }

                var automaticListEntryContract = info.TocEntry || automaticListEntryStyle;
                var automaticListTitleContract = info.TocTitle || automaticListTitleStyle;
                if (automaticListEntryContract || automaticListTitleContract)
                {
                    // Native Word field updates can auto-redefine built-in TOC/Table-of-Figures styles
                    // from regenerated result paragraphs. Remove w:autoRedefine so the Persian
                    // B Nazanin + RTL/start style contract survives subsequent Word updates.
                    var autoRedefine = style.Element(W + "autoRedefine");
                    if (autoRedefine is not null)
                    {
                        autoRedefine.Remove();
                        changedThisStyle = true;
                    }

                    var listStylePPr = EnsureChild(style, "pPr", StyleOrder);
                    if (EnsureOn(listStylePPr, "bidi", PPrOrder)) changedThisStyle = true;
                    // Keep automatic-list STYLE DEFINITIONS direction-aware and logical-start.
                    // Concrete Persian list-title PARAGRAPHS are centered later by direct paragraph formatting.
                    // This preserves the accepted Generic TOCHeading integrity contract while still centering
                    // the visible titles: فهرست مطالب / فهرست اشکال / فهرست جداول.
                    if (SetAlignment(listStylePPr, "start")) changedThisStyle = true;

                    var listRPr = EnsureChild(style, "rPr", StyleOrder);
                    var compactStyleId = (styleId ?? string.Empty).Replace(" ", string.Empty);
                    if (automaticListTitleContract)
                    {
                        if (EnsureFontSizeHalfPoints(listRPr, 28)) changedThisStyle = true; // 14pt
                        if (EnsureToggleOn(listRPr, "b")) changedThisStyle = true;
                        if (EnsureToggleOn(listRPr, "bCs")) changedThisStyle = true;
                    }
                    else if (compactStyleId.Equals("TOC1", StringComparison.OrdinalIgnoreCase))
                    {
                        if (EnsureFontSizeHalfPoints(listRPr, 28)) changedThisStyle = true; // chapter: 14pt
                        if (EnsureToggleOn(listRPr, "b")) changedThisStyle = true;
                        if (EnsureToggleOn(listRPr, "bCs")) changedThisStyle = true;
                    }
                    else if (compactStyleId.Equals("TOC2", StringComparison.OrdinalIgnoreCase))
                    {
                        if (EnsureFontSizeHalfPoints(listRPr, 24)) changedThisStyle = true; // first heading: 12pt
                        if (EnsureToggleOn(listRPr, "b")) changedThisStyle = true;
                        if (EnsureToggleOn(listRPr, "bCs")) changedThisStyle = true;
                    }
                    else
                    {
                        if (EnsureFontSizeHalfPoints(listRPr, 24)) changedThisStyle = true; // deeper levels / LOF / LOT
                        if (EnsureToggleOff(listRPr, "b")) changedThisStyle = true;
                        if (EnsureToggleOff(listRPr, "bCs")) changedThisStyle = true;
                    }
                }

                if (changedThisStyle)
                {
                    styleDefinitionsNormalized++;
                    styleChanged = true;
                }
            }

            if (styleChanged)
            {
                ReplaceXmlEntry(archive, "word/styles.xml", styles);
                mutatedEntries.Add("word/styles.xml");
            }
        }

        var thesisProfile = PersianParagraphSemanticResolver.IsThesisOrDissertationProfile(options.ProfileId);

        foreach (var partName in GetStoryPartNames(archive))
        {
            storyPartsExamined++;
            var coverZoneOpen = partName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase);
            var document = LoadXmlEntry(archive, partName);
            var partChanged = false;
            var paragraphIndex = 0;

            // Native Word list regeneration uses section layout as well as paragraph styles.
            // Only current main-story sections containing Persian text are normalized here.
            // Historical sectPrChange snapshots and English-only sections remain untouched.
            if (options.NormalizeToc && partName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase))
                partChanged |= NormalizePersianSectionLayouts(document);

            foreach (var paragraph in document.Descendants(W + "p"))
            {
                paragraphIndex++;
                var text = ParagraphText(paragraph);
                var hasPersian = ContainsPersianScript(text);
                var hasVisibleText = !string.IsNullOrWhiteSpace(text);
                var inTable = paragraph.Ancestors(W + "tbl").Any();
                var inTextBox = paragraph.Ancestors(W + "txbxContent").Any();
                var hasVisual = paragraph.Descendants(W + "drawing").Any() ||
                                paragraph.Descendants(W + "pict").Any();

                // Physical object/layout rules apply even when the host paragraph has no Persian text.
                if (!hasPersian)
                {
                    if (hasVisual || ((inTable || inTextBox) && hasVisibleText))
                    {
                        var layoutPPr = paragraph.Element(W + "pPr");
                        if (layoutPPr is null)
                        {
                            layoutPPr = new XElement(W + "pPr");
                            paragraph.AddFirst(layoutPPr);
                            partChanged = true;
                        }

                        if (SetAlignment(layoutPPr, hasVisual ? "center" : "right"))
                            partChanged = true;
                    }

                    continue;
                }

                persianParagraphs++;
                var pPr = paragraph.Element(W + "pPr");
                if (pPr is null)
                {
                    pPr = new XElement(W + "pPr");
                    paragraph.AddFirst(pPr);
                    partChanged = true;
                }

                if (EnsureOn(pPr, "bidi", PPrOrder))
                    partChanged = true;

                var info = ResolveParagraphPersianSemanticStyleInfo(paragraph, styleInfo);
                var semantic = PersianParagraphSemanticResolver.Resolve(
                    new PersianParagraphSemanticContext(
                        Text: text,
                        Style: info,
                        ParagraphIndex: paragraphIndex,
                        IsMainDocumentPart: partName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase),
                        InTable: inTable,
                        InTextBox: inTextBox,
                        HasVisual: hasVisual,
                        IsFieldDriven: paragraph.Descendants(W + "instrText").Any() ||
                                       paragraph.Descendants(W + "fldChar").Any(),
                        IsThesisProfile: thesisProfile));
                var isTocTitle = semantic.TocTitle;
                var isTocEntry = semantic.TocEntry;
                var paragraphStyleId = GetParagraphStyleId(paragraph);
                var isAutomaticListEntry = isTocEntry || IsAutomaticListEntryStyle(paragraphStyleId, string.Empty);
                var isAutomaticListTitle = !semantic.FieldDriven && (isTocTitle || IsAutomaticListTitleText(text));
                var isTitle = semantic.Title;
                var isChapter = semantic.Chapter;
                var isHeading = semantic.Heading;
                var isCaption = semantic.Caption;
                var isFieldDriven = semantic.FieldDriven;
                var isLabel = semantic.Label;
                // Keep Apply and VerifyVisualContract on the same physical-caption predicate.
                // Owner typography authority: every physical figure/table caption is B Nazanin 12pt regular/non-bold.
                var isPositionOnlyFigureCaption =
                    !isFieldDriven &&
                    LooksLikePositionOnlyFigureCaption(paragraph, text);
                var isPositionOnlyTableCaption =
                    LooksLikePositionOnlyTableCaption(paragraph, text);
                var isPhysicalCaption =
                    info?.Caption == true ||
                    (!isFieldDriven &&
                     LooksLikeExplicitCaptionSyntax(text) &&
                     IsInExpectedCaptionPosition(paragraph, text)) ||
                    isPositionOnlyFigureCaption ||
                    isPositionOnlyTableCaption;
                var isFigureCaption =
                    !isFieldDriven &&
                    IsFigureCaptionPosition(paragraph) &&
                    (info?.Caption == true ||
                     LooksLikeExplicitCaptionSyntax(text) ||
                     isPositionOnlyFigureCaption);
                var currentAlignment = GetAlignment(pPr);
                var explicitThesisChapterTitle = semantic.ExplicitThesisChapterTitle;
                var earlyCoverLine = semantic.EarlyCoverLine;

                var targetFont = isPhysicalCaption ? "B Nazanin" : semantic.TargetFont;

                if (isPhysicalCaption)
                {
                    var paragraphMarkRPr = EnsureChild(pPr, "rPr", PPrOrder);
                    if (EnsureFullFontFamily(paragraphMarkRPr, "B Nazanin"))
                        partChanged = true;
                    if (EnsureFontSizeHalfPoints(paragraphMarkRPr, 24))
                        partChanged = true;
                    if (EnsureToggleOff(paragraphMarkRPr, "b"))
                        partChanged = true;
                    if (EnsureToggleOff(paragraphMarkRPr, "bCs"))
                        partChanged = true;
                    if (EnsurePersianLanguage(paragraphMarkRPr, setPrimaryLanguage: true))
                        partChanged = true;
                }

                if (isAutomaticListEntry && options.NormalizeToc)
                {
                    if (SetAlignment(pPr, "start"))
                        partChanged = true;
                    tocParagraphsNormalized++;
                }
                else if (hasVisual)
                {
                    if (SetAlignment(pPr, "center"))
                        partChanged = true;
                }
                else if (inTable || inTextBox)
                {
                    // Table/text-box content has higher layout authority than semantic Caption/Title/Heading.
                    // Generic/OpenXML staging contract: Persian table/text-box content remains bidi + logical start
                    // until the thesis native Word round-trip completes. The thesis path then reasserts physical
                    // right after Word Save/Close via RestabilizeTableTextBoxVisualRightOpenXml.
                    if (SetAlignment(pPr, "start"))
                        partChanged = true;
                }
                else if (isPhysicalCaption)
                {
                    // Center only real visual-object captions outside table/text-box content:
                    // profile/style-authoritative captions or explicit caption syntax in the expected
                    // physical position (table title above table; figure/chart caption below visual).
                    if (SetAlignment(pPr, "center"))
                        partChanged = true;
                }
                else if (isAutomaticListTitle)
                {
                    // Visible automatic-list title authority must override legacy paragraph-mark formatting.
                    var listTitleMarkRPr = EnsureChild(pPr, "rPr", PPrOrder);
                    if (EnsureFullFontFamily(listTitleMarkRPr, "B Nazanin")) partChanged = true;
                    if (EnsureFontSizeHalfPoints(listTitleMarkRPr, 28)) partChanged = true;
                    if (EnsureToggleOn(listTitleMarkRPr, "b")) partChanged = true;
                    if (EnsureToggleOn(listTitleMarkRPr, "bCs")) partChanged = true;
                    if (EnsurePersianLanguage(listTitleMarkRPr, setPrimaryLanguage: true)) partChanged = true;
                    if (SetAlignment(pPr, "center")) partChanged = true;
                    titleParagraphsCentered++;
                }
                else if (isTitle || earlyCoverLine)
                {
                    if (SetAlignment(pPr, "center"))
                        partChanged = true;
                    titleParagraphsCentered++;
                }
                else if (explicitThesisChapterTitle)
                {
                    if (SetAlignment(pPr, "center"))
                        partChanged = true;
                    titleParagraphsCentered++;
                }
                else if ((isHeading || isChapter) && options.NormalizeHeadingAlignment)
                {
                    // A style name containing Chapter/Fasl is not sufficient to center content.
                    // Only an explicit "فصل ..." title in a thesis/dissertation profile may center.
                    // Ordinary RTL headings use logical start, which resolves to the visual right edge in Word.
                    if (SetAlignment(pPr, "start"))
                        partChanged = true;
                    headingParagraphsRightAligned++;
                }
                else if (options.NormalizeBodyAlignment &&
                         !isFieldDriven &&
                         !isPhysicalCaption &&
                         !isLabel)
                {
                    if (SetAlignment(pPr, "both"))
                        partChanged = true;
                    bodyParagraphsJustified++;
                }
                else if (string.IsNullOrWhiteSpace(currentAlignment) ||
                         currentAlignment.Equals("left", StringComparison.OrdinalIgnoreCase) ||
                         currentAlignment.Equals("right", StringComparison.OrdinalIgnoreCase) ||
                         currentAlignment.Equals("end", StringComparison.OrdinalIgnoreCase))
                {
                    if (SetAlignment(pPr, "start"))
                        partChanged = true;
                }

                var preserveGenericNativeAutomaticListFieldRuns =
                    !thesisProfile &&
                    isFieldDriven &&
                    paragraph.Descendants(W + "instrText").Any(x =>
                        (((string?)x ?? string.Empty).Contains("TOC", StringComparison.OrdinalIgnoreCase) ||
                         ((string?)x ?? string.Empty).Contains("TOF", StringComparison.OrdinalIgnoreCase)));

                if (preserveGenericNativeAutomaticListFieldRuns)
                    continue;

                foreach (var run in paragraph.Descendants(W + "r"))
                {
                    var rPr = run.Element(W + "rPr");
                    if (rPr is null)
                    {
                        rPr = new XElement(W + "rPr");
                        run.AddFirst(rPr);
                        partChanged = true;
                    }

                    var runText = RunText(run);
                    var containsPersian = ContainsPersianScript(runText);
                    var containsLatin = ContainsLatinScript(runText);

                    var runChanged = false;
                    if (isAutomaticListEntry || isAutomaticListTitle || isPhysicalCaption)
                    {
                        if (EnsureFullFontFamily(rPr, "B Nazanin"))
                            runChanged = true;
                    }
                    else if (EnsureComplexScriptFont(rPr, targetFont))
                    {
                        runChanged = true;
                    }
                    if (isAutomaticListTitle)
                    {
                        if (EnsureFontSizeHalfPoints(rPr, 28)) runChanged = true;
                        if (EnsureToggleOn(rPr, "b")) runChanged = true;
                        if (EnsureToggleOn(rPr, "bCs")) runChanged = true;
                    }
                    if (isPhysicalCaption && EnsureFontSizeHalfPoints(rPr, 24))
                        runChanged = true;
                    if (isPhysicalCaption && EnsureToggleOff(rPr, "b"))
                        runChanged = true;
                    if (isPhysicalCaption && EnsureToggleOff(rPr, "bCs"))
                        runChanged = true;
                    if (EnsurePersianLanguage(rPr, setPrimaryLanguage: containsPersian || isPhysicalCaption))
                        runChanged = true;

                    if (containsPersian)
                    {
                        if (EnsureOn(rPr, "rtl", RPrOrder))
                            runChanged = true;
                    }
                    else if (containsLatin)
                    {
                        // A Latin-only technical run inside an RTL Persian paragraph must remain
                        // logically LTR. Upstream/profile overlays may have materialized w:rtl on
                        // every run; the shared Foundation is the final authority and removes it.
                        var rtl = rPr.Element(W + "rtl");
                        if (rtl is not null)
                        {
                            rtl.Remove();
                            runChanged = true;
                        }
                    }

                    if (runChanged)
                    {
                        persianRunsNormalized++;
                        partChanged = true;
                    }
                }
            }

            if (NormalizeTables(document))
                partChanged = true;

            ValidateVisualLayoutContract(document, styleInfo, options, partName);

            if (partChanged)
            {
                ReplaceXmlEntry(archive, partName, document);
                mutatedEntries.Add(partName);
            }
        }

        var result = new PersianDocumentFoundationResult(
            filePath,
            options.ProfileId,
            storyPartsExamined,
            persianParagraphs,
            bodyParagraphsJustified,
            headingParagraphsRightAligned,
            titleParagraphsCentered,
            tocParagraphsNormalized,
            persianRunsNormalized,
            styleDefinitionsNormalized,
            docDefaultsNormalized,
            mutatedEntries.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray());

        // Close the Open XML package before Word opens the same file.  For thesis/dissertation
        // documents that contain native automatic lists, Word itself is the final authority for
        // built-in TOC/Table-of-Figures style persistence.  OpenXML-only style edits were proven
        // insufficient by R17R7R3/R17R7R5: Word regenerated the entries and reset the built-in
        // style properties.  Apply the supported Word object-model style definitions after one
        // native regeneration, without changing any field instruction payload.
        archive.Dispose();
        stream.Dispose();

        // Owner visual authority R17R7R6R4R6R5:
        // - if a visible Persian "فهرست مطالب" exists but the document has no native TOC field,
        //   materialize one exactly once and remove only an immediately-following legacy static TOC block;
        // - for styleless numbered Persian headings, attach outline levels so Word can regenerate the TOC.
        // Generic Persianize must preserve its established formatting-only integrity contract.
        // Native TOC materialization for Generic outputs is invoked only after that integrity gate
        // has PASSed, through MaterializeMainAutomaticTocAfterIntegrity(). Thesis paths retain
        // their existing Word-native lifecycle here.

        if (options.NormalizeToc &&
            HasNativeAutomaticListFields(filePath) &&
            thesisProfile)
        {
            StabilizeNativeAutomaticListsWithWord(filePath);
            // Word 16.0 can disconnect the COM server while mutating RTL reading order on
            // built-in TOC styles.  Keep the Word object model responsible for the stable
            // properties it accepts (AutomaticallyUpdate/font), then restabilize
            // the bidi/style XML only after Word has saved and fully closed.
            RestabilizeAutomaticListStylesOpenXml(filePath);
            // Word Table-of-Figures regeneration copies the numbering symbol run properties
            // for caption levels into the generated label/number prefix. Normalize those
            // numbering-symbol properties after Word closes so later independent updates
            // cannot reintroduce Times New Roman into شکل/جدول prefixes.
            RestabilizeCaptionNumberingSymbolsOpenXml(filePath);
            // Native Word save can rewrite logical start to physical left in table-cell paragraphs.
            // Reassert the owner-facing table/text-box contract after Word closes, using explicit right.
            RestabilizeTableTextBoxVisualRightOpenXml(filePath);
            // Word 16.0 may remove w:updateFields while saving a document after a native field update.
            // Reassert the automatic-update-on-open lifecycle setting only after Word has fully closed,
            // using schema-order-safe Open XML mutation. Field instructions remain immutable.
            RestabilizeUpdateFieldsOnOpenOpenXml(filePath);
        }

        return result;
    }


    public static bool RestabilizeExistingAutomaticListsOpenXml(string filePath)
    {
        // Shared post-merge authority for documents that already own native automatic-list fields.
        // This path never creates a TOC and never opens/saves the document through Word.
        if (!HasNativeAutomaticListFields(filePath))
            return false;

        RestabilizeAutomaticListStylesOpenXml(filePath);
        RestabilizeUpdateFieldsOnOpenOpenXml(filePath);
        return true;
    }

    public static bool MaterializeMainAutomaticTocAfterIntegrity(string filePath)
    {
        // Narrow post-integrity Generic structural mutation: OpenXML only. Never open/save
        // the Generic document through Word here, because owner evidence proved that a Word
        // round-trip can rewrite unrelated package/media/bookmark state.
        var materializedNow = EnsureMainAutomaticTocMaterializedOpenXml(filePath);
        var hasNativeAutomaticList = materializedNow || HasNativeAutomaticListFields(filePath);
        if (!hasNativeAutomaticList)
            return false;

        // Word copies direct character formatting from styleless outline sources into regenerated TOC entries.
        // Promote only those synthetic styleless outline sources to Heading1..3 and remove direct font-family
        // declarations so the TOC1/2/3 B Nazanin style contract remains authoritative after native Update/F9.
        NormalizeStylelessMainTocSourceHeadingFontInheritanceOpenXml(filePath);

        RestabilizeAutomaticListStylesOpenXml(filePath);
        RestabilizeUpdateFieldsOnOpenOpenXml(filePath);
        return materializedNow;
    }

    private static bool EnsureMainAutomaticTocMaterializedOpenXml(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);
        var documentEntry = archive.GetEntry("word/document.xml");
        if (documentEntry is null)
            return false;

        XDocument document;
        using (var reader = new StreamReader(documentEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: false))
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);

        if (document.Descendants(W + "instrText")
            .Select(x => ((string?)x ?? string.Empty).TrimStart())
            .Any(x => x.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)))
            return false;

        var body = document.Root?.Element(W + "body");
        if (body is null)
            return false;

        var paragraphs = body.Elements(W + "p").ToList();
        var tocTitle = paragraphs.FirstOrDefault(p =>
            NormalizeListTitleText(ParagraphText(p)).Equals("فهرستمطالب", StringComparison.Ordinal));
        if (tocTitle is null)
            return false;

        // Convert only a legacy static TOC immediately below the visible title. We intentionally
        // stop at the first blank paragraph so later document content is never consumed.
        var staticRows = new List<XElement>();
        var cursor = tocTitle.NextNode as XElement;
        while (cursor is not null && cursor.Name == W + "p")
        {
            var text = ParagraphText(cursor).Trim();
            if (string.IsNullOrWhiteSpace(text))
                break;
            if (!LooksLikeLegacyStaticTocRow(text))
                break;
            staticRows.Add(cursor);
            cursor = cursor.NextNode as XElement;
        }

        if (staticRows.Count >= 3)
        {
            foreach (var row in staticRows)
                row.Remove();
        }

        // Styleless numbered Persian documents (e.g. owner review Generic04) need outline
        // levels before Word can regenerate a native \o/\u TOC. Existing styled headings are preserved.
        foreach (var paragraph in body.Elements(W + "p"))
        {
            var text = ParagraphText(paragraph).Trim();
            if (string.IsNullOrWhiteSpace(text))
                continue;

            var existingStyle = GetParagraphStyleId(paragraph);
            if (!string.IsNullOrWhiteSpace(existingStyle))
                continue;

            if (!TryGetNumberedPersianHeadingOutlineLevel(text, out var outlineLevel))
                continue;

            var pPr = paragraph.Element(W + "pPr");
            if (pPr is null)
            {
                pPr = new XElement(W + "pPr");
                paragraph.AddFirst(pPr);
            }

            var outline = pPr.Element(W + "outlineLvl");
            if (outline is null)
            {
                outline = new XElement(W + "outlineLvl");
                InsertInSchemaOrder(pPr, outline, PPrOrder);
            }
            outline.SetAttributeValue(W + "val", outlineLevel.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        var fieldParagraph = BuildNativeTocFieldParagraph();
        tocTitle.AddAfterSelf(fieldParagraph);

        ReplaceXmlEntry(archive, "word/document.xml", document);
        return true;
    }

    private static bool NormalizeStylelessMainTocSourceHeadingFontInheritanceOpenXml(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);
        var documentEntry = archive.GetEntry("word/document.xml");
        if (documentEntry is null)
            return false;

        XDocument document;
        using (var reader = new StreamReader(documentEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: false))
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);

        var body = document.Root?.Element(W + "body");
        if (body is null)
            return false;

        var changed = NormalizeMainTocPresentationTitles(document);
        foreach (var paragraph in body.Elements(W + "p"))
        {
            var text = ParagraphText(paragraph).Trim();
            if (string.IsNullOrWhiteSpace(text) || IsAutomaticListTitleText(text))
                continue;
            if (paragraph.Descendants(W + "fldChar").Any() || paragraph.Descendants(W + "instrText").Any())
                continue;

            // Scope deliberately excludes existing styled headings. The regression is specific to
            // Generic04-style fallback headings that WordAgent itself annotated with direct outlineLvl.
            if (!string.IsNullOrWhiteSpace(GetParagraphStyleId(paragraph)))
                continue;

            var pPr = paragraph.Element(W + "pPr");
            var outlineText = ((string?)pPr?.Element(W + "outlineLvl")?.Attribute(W + "val") ?? string.Empty).Trim();
            if (!int.TryParse(
                    outlineText,
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var outlineLevel) ||
                outlineLevel is < 0 or > 2)
            {
                continue;
            }

            pPr ??= new XElement(W + "pPr");
            if (paragraph.Element(W + "pPr") is null)
                paragraph.AddFirst(pPr);

            var pStyle = pPr.Element(W + "pStyle");
            if (pStyle is null)
            {
                pStyle = new XElement(W + "pStyle");
                InsertInSchemaOrder(pPr, pStyle, PPrOrder);
            }
            pStyle.SetAttributeValue(
                W + "val",
                "Heading" + (outlineLevel + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
            changed = true;

            // Preserve size/bold/spacing/language; remove only direct font-family overrides so Word's
            // regenerated TOC text inherits B Nazanin from TOC1/2/3 instead of the source B Titr.
            foreach (var rPr in paragraph.Descendants(W + "rPr").ToList())
            {
                var rFonts = rPr.Element(W + "rFonts");
                if (rFonts is null)
                    continue;
                rFonts.Remove();
                changed = true;
            }

            var paragraphMarkFonts = pPr.Element(W + "rPr")?.Element(W + "rFonts");
            if (paragraphMarkFonts is not null)
            {
                paragraphMarkFonts.Remove();
                changed = true;
            }
        }

        if (!changed)
            return false;

        ReplaceXmlEntry(archive, "word/document.xml", document);
        return true;
    }

    private static bool NormalizeMainTocPresentationTitles(XDocument document)
    {
        var body = document.Root?.Element(W + "body");
        if (body is null) return false;
        bool changed = false;
        foreach (var paragraph in body.Elements(W + "p"))
        {
            // Source presentation titles only: never edit native field results or hyperlink text.
            if (paragraph.Descendants(W + "fldChar").Any() || paragraph.Descendants(W + "instrText").Any()
                || paragraph.Descendants(W + "hyperlink").Any()) continue;
            var compact = NormalizeListTitleText(ParagraphText(paragraph));
            if (!(compact.Equals("فهرستمطالب", StringComparison.Ordinal)
                || compact.StartsWith("فهرستاشکال", StringComparison.Ordinal)
                || compact.StartsWith("فهرستجداول", StringComparison.Ordinal)) || compact.Length > 80) continue;
            var pPr = paragraph.Element(W + "pPr");
            if (pPr is null) { pPr = new XElement(W + "pPr"); paragraph.AddFirst(pPr); changed = true; }
            var pStyle = EnsureChild(pPr, "pStyle", PPrOrder);
            if ((string?)pStyle.Attribute(W + "val") != "TOCHeading") { pStyle.SetAttributeValue(W + "val", "TOCHeading"); changed = true; }
            var outline = EnsureChild(pPr, "outlineLvl", PPrOrder);
            if ((string?)outline.Attribute(W + "val") != "9") { outline.SetAttributeValue(W + "val", "9"); changed = true; }
            changed |= EnsureOn(pPr, "bidi", PPrOrder);
            var jc = EnsureChild(pPr, "jc", PPrOrder);
            if ((string?)jc.Attribute(W + "val") != "center") { jc.SetAttributeValue(W + "val", "center"); changed = true; }
            var properties = new List<XElement> { EnsureChild(pPr, "rPr", PPrOrder) };
            foreach (var run in paragraph.Elements(W + "r"))
            {
                var runProperties = run.Element(W + "rPr");
                if (runProperties is null) { runProperties = new XElement(W + "rPr"); run.AddFirst(runProperties); changed = true; }
                properties.Add(runProperties);
            }
            foreach (var rPr in properties)
            {
                changed |= EnsureFullFontFamily(rPr, "B Nazanin");
                changed |= EnsureFontSizeHalfPoints(rPr, 28);
                changed |= EnsureToggleOn(rPr, "b");
                changed |= EnsureToggleOn(rPr, "bCs");
            }
        }
        return changed;
    }

    private static string NormalizeListTitleText(string text)
        => (text ?? string.Empty)
            .Replace(" ", string.Empty)
            .Replace("\u200c", string.Empty)
            .Replace("\u200f", string.Empty)
            .Replace("\u200e", string.Empty)
            .Trim();

    private static bool LooksLikeLegacyStaticTocRow(string text)
    {
        var compact = (text ?? string.Empty).Trim();
        if (compact.Length < 2)
            return false;

        // Static TOC rows in the real review corpus begin with an Arabic/Persian number
        // and end with the page number. Require both to avoid deleting narrative paragraphs.
        return Regex.IsMatch(
            compact,
            @"^[0-9۰-۹]+(?:\s*[-–—‑]\s*[0-9۰-۹]+){0,3}\s*[-–—‑]?\s*.+[0-9۰-۹]+$",
            RegexOptions.CultureInvariant);
    }

    private static bool TryGetNumberedPersianHeadingOutlineLevel(string text, out int outlineLevel)
    {
        outlineLevel = 0;
        var match = Regex.Match(
            (text ?? string.Empty).Trim(),
            @"^(?<number>[0-9۰-۹]+(?:\s*[-–—‑]\s*[0-9۰-۹]+){0,2})\s*[-–—‑]\s*\S",
            RegexOptions.CultureInvariant);
        if (!match.Success)
            return false;

        var number = match.Groups["number"].Value;
        var segments = Regex.Split(number, @"\s*[-–—‑]\s*")
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToArray();
        if (segments.Length < 1 || segments.Length > 3)
            return false;

        outlineLevel = segments.Length - 1;
        return true;
    }

    private static XElement BuildNativeTocFieldParagraph()
    {
        var pPr = new XElement(W + "pPr",
            new XElement(W + "bidi"),
            new XElement(W + "jc", new XAttribute(W + "val", "start")));

        return new XElement(W + "p",
            pPr,
            new XElement(W + "r",
                new XElement(W + "fldChar",
                    new XAttribute(W + "fldCharType", "begin"),
                    new XAttribute(W + "dirty", "true"))),
            new XElement(W + "r",
                new XElement(W + "instrText",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    " TOC \\o \"1-3\" \\h \\z \\u ")),
            new XElement(W + "r",
                new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
            new XElement(W + "r",
                new XElement(W + "rPr",
                    new XElement(W + "rFonts",
                        new XAttribute(W + "ascii", "B Nazanin"),
                        new XAttribute(W + "hAnsi", "B Nazanin"),
                        new XAttribute(W + "eastAsia", "B Nazanin"),
                        new XAttribute(W + "cs", "B Nazanin")),
                    new XElement(W + "rtl")),
                new XElement(W + "t", "فهرست مطالب")),
            new XElement(W + "r",
                new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));
    }

    private static bool HasNativeAutomaticListFields(string filePath)
    {
        using var package = ZipFile.OpenRead(filePath);
        var entry = package.GetEntry("word/document.xml");
        if (entry is null)
            return false;

        using var entryStream = entry.Open();
        var document = XDocument.Load(entryStream, LoadOptions.PreserveWhitespace);
        return document
            .Descendants(W + "instrText")
            .Select(x => ((string?)x ?? string.Empty).TrimStart())
            .Any(x => x.StartsWith("TOC", StringComparison.OrdinalIgnoreCase));
    }

    private static void StabilizeNativeAutomaticListsWithWord(string filePath)
    {
        var baselineWord = Process.GetProcessesByName("WINWORD");
        try
        {
            if (baselineWord.Length != 0)
                throw new InvalidOperationException(
                    "Native automatic-list stabilization requires Microsoft Word to be closed before processing.");
        }
        finally
        {
            foreach (var process in baselineWord)
                process.Dispose();
        }

        Word.Application? application = null;
        Word.Document? document = null;
        WordAutomationSecurityScope? automationSecurity = null;
        uint ownedWordProcessId = 0;
        var launchStartedUtc = DateTime.UtcNow;
        try
        {
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
                ScreenUpdating = false
            };
            automationSecurity = WordAutomationSecurityScope.ForceDisable(application);

            var launchedWord = Process.GetProcessesByName("WINWORD");
            try
            {
                var ownedCandidates = launchedWord
                    .Where(p =>
                    {
                        try { return p.StartTime.ToUniversalTime() >= launchStartedUtc.AddSeconds(-2); }
                        catch { return false; }
                    })
                    .Select(p => p.Id)
                    .Distinct()
                    .ToArray();
                if (ownedCandidates.Length != 1)
                    throw new InvalidOperationException($"Native automatic-list stabilization could not uniquely identify its launched WINWORD process. candidates={ownedCandidates.Length}");
                ownedWordProcessId = (uint)ownedCandidates[0];
            }
            finally
            {
                foreach (var process in launchedWord) process.Dispose();
            }

            document = application.Documents.Open(
                FileName: filePath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);

            // Force Word to materialize its native result paragraphs first.  We then modify the
            // built-in document styles through Word itself, which is the persistence mechanism
            // documented by Microsoft for automatic TOC formatting.
            UpdateNativeAutomaticLists(document);
            StabilizeAutomaticListStyle(document, Word.WdBuiltinStyle.wdStyleTOC1);
            StabilizeAutomaticListStyle(document, Word.WdBuiltinStyle.wdStyleTOC2);
            StabilizeAutomaticListStyle(document, Word.WdBuiltinStyle.wdStyleTOC3);
            StabilizeAutomaticListStyle(document, Word.WdBuiltinStyle.wdStyleTableOfFigures);
            NormalizeAutomaticListRanges(document);
            document.Save();
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                FinalReleaseComObject(document);
            }

            automationSecurity?.Dispose();
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                FinalReleaseComObject(application);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        var deadline = DateTime.UtcNow.AddSeconds(10);
        while (DateTime.UtcNow < deadline)
        {
            if (ownedWordProcessId == 0 || !IsProcessAlive((int)ownedWordProcessId))
                return;
            Thread.Sleep(250);
        }

        if (ownedWordProcessId != 0)
            TerminateOwnedWordProcess((int)ownedWordProcessId);

        if (ownedWordProcessId == 0 || !IsProcessAlive((int)ownedWordProcessId))
            return;

        throw new InvalidOperationException(
            "Native automatic-list stabilization could not terminate its owned WINWORD process after Save/Close/Quit.");
    }

    private static bool IsProcessAlive(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch (ArgumentException) { return false; }
    }

    private static void TerminateOwnedWordProcess(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                process.WaitForExit(5000);
            }
        }
        catch (ArgumentException) { }
    }

    private static void UpdateNativeAutomaticLists(Word.Document document)
    {
        Word.TablesOfContents? contents = null;
        Word.TablesOfFigures? figures = null;
        try
        {
            contents = document.TablesOfContents;
            for (var index = 1; index <= contents.Count; index++)
            {
                Word.TableOfContents? toc = null;
                try
                {
                    toc = contents[index];
                    toc.Update();
                }
                finally
                {
                    FinalReleaseComObject(toc);
                }
            }

            figures = document.TablesOfFigures;
            for (var index = 1; index <= figures.Count; index++)
            {
                Word.TableOfFigures? table = null;
                try
                {
                    table = figures[index];
                    table.Update();
                }
                finally
                {
                    FinalReleaseComObject(table);
                }
            }
        }
        finally
        {
            FinalReleaseComObject(figures);
            FinalReleaseComObject(contents);
        }
    }

    private static void StabilizeAutomaticListStyle(Word.Document document, Word.WdBuiltinStyle builtInStyle)
    {
        Word.Styles? styles = null;
        Word.Style? style = null;
        Word.Font? font = null;
        try
        {
            styles = document.Styles;
            style = styles[builtInStyle];
            RunWordInteropStep($"STYLE_{builtInStyle}_AUTOMATICALLY_UPDATE_OFF", () => style.AutomaticallyUpdate = false);

            font = style.Font;
            RunWordInteropStep($"STYLE_{builtInStyle}_FONT_NAME", () => font.Name = "B Nazanin");
            RunWordInteropStep($"STYLE_{builtInStyle}_FONT_NAME_BI", () => font.NameBi = "B Nazanin");

            // R17R7R6R2 reproduced Word RPC disconnect 0x800706BE at the built-in
            // TOC1 Alignment setter even with COM ReadingOrder already absent. Do not
            // mutate ParagraphFormat through COM in this path. Alignment + bidi are
            // reasserted after Word Save/Close by RestabilizeAutomaticListStylesOpenXml.
        }
        finally
        {
            FinalReleaseComObject(font);
            FinalReleaseComObject(style);
            FinalReleaseComObject(styles);
        }
    }

    private static void NormalizeAutomaticListRanges(Word.Document document)
    {
        Word.TablesOfContents? contents = null;
        Word.TablesOfFigures? figures = null;
        try
        {
            contents = document.TablesOfContents;
            for (var index = 1; index <= contents.Count; index++)
            {
                Word.TableOfContents? toc = null;
                Word.Range? range = null;
                try
                {
                    toc = contents[index];
                    range = toc.Range;
                    NormalizeAutomaticListRange(range);
                }
                finally
                {
                    FinalReleaseComObject(range);
                    FinalReleaseComObject(toc);
                }
            }

            figures = document.TablesOfFigures;
            for (var index = 1; index <= figures.Count; index++)
            {
                Word.TableOfFigures? table = null;
                Word.Range? range = null;
                try
                {
                    table = figures[index];
                    range = table.Range;
                    NormalizeAutomaticListRange(range);
                }
                finally
                {
                    FinalReleaseComObject(range);
                    FinalReleaseComObject(table);
                }
            }
        }
        finally
        {
            FinalReleaseComObject(figures);
            FinalReleaseComObject(contents);
        }
    }

    private static void NormalizeAutomaticListRange(Word.Range range)
    {
        Word.Font? font = null;
        try
        {
            font = range.Font;
            RunWordInteropStep("LIST_RANGE_FONT_NAME", () => font.Name = "B Nazanin");
            RunWordInteropStep("LIST_RANGE_FONT_NAME_BI", () => font.NameBi = "B Nazanin");

            // No COM ParagraphFormat mutation here. The post-close style contract is the
            // single alignment/RTL authority, and the independent Word re-update gate
            // proves whether that contract survives regeneration.
        }
        finally
        {
            FinalReleaseComObject(font);
        }
    }

    private static void RestabilizeAutomaticListStylesOpenXml(string filePath)
    {
        XDocument styles;
        using (var readStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var readArchive = new ZipArchive(readStream, ZipArchiveMode.Read, leaveOpen: false))
        {
            var readStylesEntry = readArchive.GetEntry("word/styles.xml")
                ?? throw new InvalidOperationException("Automatic-list post-Word restabilization requires word/styles.xml.");
            using var reader = new StreamReader(readStylesEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: false);
            styles = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }

        var stylesRoot = styles.Root
            ?? throw new InvalidOperationException("Automatic-list style stabilization requires a valid styles root.");

        var changed = false;

        // Generic native TOC is materialized with OpenXML only after formatting-integrity PASS.
        // Because Word is intentionally not opened/saved in this path, missing built-in TOC styles
        // must be materialized explicitly so the later Word field update inherits owner typography.
        foreach (var spec in new[]
        {
            new { Id = "TOC1", Name = "toc 1" },
            new { Id = "TOC2", Name = "toc 2" },
            new { Id = "TOC3", Name = "toc 3" }
        })
        {
            var exists = stylesRoot.Elements(W + "style").Any(style =>
            {
                var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty)
                    .Replace(" ", string.Empty).Trim();
                var styleName = ((string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty)
                    .Replace(" ", string.Empty).Trim();
                return styleId.Equals(spec.Id, StringComparison.OrdinalIgnoreCase) ||
                       styleName.Equals(spec.Name.Replace(" ", string.Empty), StringComparison.OrdinalIgnoreCase);
            });

            if (exists)
                continue;

            var style = new XElement(W + "style",
                new XAttribute(W + "type", "paragraph"),
                new XAttribute(W + "styleId", spec.Id),
                new XElement(W + "name", new XAttribute(W + "val", spec.Name)),
                new XElement(W + "basedOn", new XAttribute(W + "val", "Normal")),
                new XElement(W + "next", new XAttribute(W + "val", "Normal")),
                new XElement(W + "uiPriority", new XAttribute(W + "val", "39")),
                new XElement(W + "semiHidden"),
                new XElement(W + "unhideWhenUsed"));

            var extLst = stylesRoot.Elements().FirstOrDefault(x => x.Name.LocalName == "extLst");
            if (extLst is null)
                stylesRoot.Add(style);
            else
                extLst.AddBeforeSelf(style);

            changed = true;
        }

        foreach (var style in styles.Descendants(W + "style"))
        {
            var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Replace(" ", string.Empty).Trim();
            var styleName = ((string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty).Replace(" ", string.Empty).Trim();
            var automaticEntryStyle =
                styleId.Equals("TOC1", StringComparison.OrdinalIgnoreCase) ||
                styleId.Equals("TOC2", StringComparison.OrdinalIgnoreCase) ||
                styleId.Equals("TOC3", StringComparison.OrdinalIgnoreCase) ||
                styleId.Equals("TableofFigures", StringComparison.OrdinalIgnoreCase) ||
                styleName.Equals("TOC1", StringComparison.OrdinalIgnoreCase) ||
                styleName.Equals("TOC2", StringComparison.OrdinalIgnoreCase) ||
                styleName.Equals("TOC3", StringComparison.OrdinalIgnoreCase) ||
                styleName.Equals("TableofFigures", StringComparison.OrdinalIgnoreCase);
            var automaticTitleStyle =
                styleId.Equals("TOCHeading", StringComparison.OrdinalIgnoreCase) ||
                styleName.Equals("TOCHeading", StringComparison.OrdinalIgnoreCase) ||
                styleName.Equals("TOC Heading", StringComparison.OrdinalIgnoreCase);
            if (!automaticEntryStyle && !automaticTitleStyle)
                continue;

            var autoRedefine = style.Element(W + "autoRedefine");
            if (autoRedefine is not null)
            {
                autoRedefine.Remove();
                changed = true;
            }

            var pPr = EnsureChild(style, "pPr", StyleOrder);
            if (EnsureOn(pPr, "bidi", PPrOrder)) changed = true;
            // Keep the built-in TOCHeading STYLE definition logical-start so Generic integrity
            // remains direction-aware. Visible Persian list-title paragraphs are centered directly.
            if (SetAlignment(pPr, automaticTitleStyle ? "start" : "right")) changed = true;

            var rPr = EnsureChild(style, "rPr", StyleOrder);
            if (EnsureFullFontFamily(rPr, "B Nazanin")) changed = true;

            if (automaticTitleStyle)
            {
                if (EnsureFontSizeHalfPoints(rPr, 28)) changed = true; // list title: 14pt
                if (EnsureToggleOn(rPr, "b")) changed = true;
                if (EnsureToggleOn(rPr, "bCs")) changed = true;
            }
            else if (styleId.Equals("TOC1", StringComparison.OrdinalIgnoreCase))
            {
                if (EnsureFontSizeHalfPoints(rPr, 28)) changed = true; // chapter: 14pt
                if (EnsureToggleOn(rPr, "b")) changed = true;
                if (EnsureToggleOn(rPr, "bCs")) changed = true;
                if (EnsureDottedPageNumberTab(pPr, 7938)) changed = true;
            }
            else if (styleId.Equals("TOC2", StringComparison.OrdinalIgnoreCase))
            {
                if (EnsureFontSizeHalfPoints(rPr, 24)) changed = true; // first heading: 12pt
                if (EnsureToggleOn(rPr, "b")) changed = true;
                if (EnsureToggleOn(rPr, "bCs")) changed = true;
                if (EnsureDottedPageNumberTab(pPr, 7938)) changed = true;
            }
            else
            {
                if (EnsureFontSizeHalfPoints(rPr, 24)) changed = true;
                if (EnsureToggleOff(rPr, "b")) changed = true;
                if (EnsureToggleOff(rPr, "bCs")) changed = true;
                if (EnsureDottedPageNumberTab(pPr, 7938)) changed = true;
            }
        }

        if (!changed)
            return;

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);
        var stylesEntry = archive.GetEntry("word/styles.xml")
            ?? throw new InvalidOperationException("Automatic-list post-Word restabilization requires word/styles.xml.");
        stylesEntry.Delete();
        var replacement = archive.CreateEntry("word/styles.xml", CompressionLevel.Optimal);
        using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        styles.Save(writer, SaveOptions.DisableFormatting);
    }




    private static void RestabilizeCaptionNumberingSymbolsOpenXml(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);
        var numberingEntry = archive.GetEntry("word/numbering.xml");
        if (numberingEntry is null)
            return;

        XDocument numbering;
        using (var reader = new StreamReader(numberingEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: false))
            numbering = XDocument.Load(reader, LoadOptions.PreserveWhitespace);

        var changed = false;
        foreach (var level in numbering.Descendants(W + "lvl"))
        {
            var levelText = ((string?)level.Element(W + "lvlText")?.Attribute(W + "val") ?? string.Empty).Trim();
            var compact = levelText.Replace("\u200f", string.Empty)
                                   .Replace("\u200e", string.Empty)
                                   .Replace(" ", string.Empty);
            var captionNumberingLevel =
                compact.StartsWith("شکل", StringComparison.Ordinal) ||
                compact.StartsWith("جدول", StringComparison.Ordinal);
            if (!captionNumberingLevel)
                continue;

            var rPr = level.Element(W + "rPr");
            if (rPr is null)
            {
                rPr = new XElement(W + "rPr");
                level.Add(rPr);
                changed = true;
            }

            if (EnsureFullFontFamily(rPr, "B Nazanin")) changed = true;
            if (EnsureFontSizeHalfPoints(rPr, 24)) changed = true;
            if (EnsureToggleOff(rPr, "b")) changed = true;
            if (EnsureToggleOff(rPr, "bCs")) changed = true;
        }

        if (!changed)
            return;

        ReplaceXmlEntry(archive, "word/numbering.xml", numbering);
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

    private static void RestabilizeUpdateFieldsOnOpenOpenXml(string filePath)
    {
        XDocument settings;
        using (var readStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var readArchive = new ZipArchive(readStream, ZipArchiveMode.Read, leaveOpen: false))
        {
            var settingsEntry = readArchive.GetEntry("word/settings.xml")
                ?? throw new InvalidOperationException("Automatic-list post-Word restabilization requires word/settings.xml.");
            using var reader = new StreamReader(settingsEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: false);
            settings = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }

        var root = settings.Root
            ?? throw new InvalidOperationException("Invalid word/settings.xml: root element is missing.");

        var changed = false;
        var updateFields = root.Element(W + "updateFields");
        if (updateFields is null)
        {
            updateFields = new XElement(W + "updateFields");
            InsertSettingsElementInSchemaOrder(root, updateFields);
            changed = true;
        }

        var current = ((string?)updateFields.Attribute(W + "val") ?? string.Empty).Trim();
        if (!current.Equals("true", StringComparison.OrdinalIgnoreCase))
        {
            updateFields.SetAttributeValue(W + "val", "true");
            changed = true;
        }

        if (!changed)
            return;

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);
        ReplaceXmlEntry(archive, "word/settings.xml", settings);
    }

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

    private static void RestabilizeTableTextBoxVisualRightOpenXml(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);

        foreach (var partName in GetStoryPartNames(archive))
        {
            var story = LoadXmlEntry(archive, partName);
            var changed = false;

            foreach (var paragraph in story.Descendants(W + "p"))
            {
                var text = ParagraphText(paragraph);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                var inTable = paragraph.Ancestors(W + "tbl").Any();
                var inTextBox = paragraph.Ancestors(W + "txbxContent").Any();
                if (!inTable && !inTextBox)
                    continue;

                var hasVisual = paragraph.Descendants(W + "drawing").Any() ||
                                paragraph.Descendants(W + "pict").Any();
                if (hasVisual)
                    continue;

                var pPr = paragraph.Element(W + "pPr");
                if (pPr is null)
                {
                    pPr = new XElement(W + "pPr");
                    paragraph.AddFirst(pPr);
                    changed = true;
                }

                if (ContainsPersianScript(text) && EnsureOn(pPr, "bidi", PPrOrder))
                    changed = true;

                if (SetAlignment(pPr, "right"))
                    changed = true;
            }

            if (changed)
                ReplaceXmlEntry(archive, partName, story);
        }
    }

    private static void RunWordInteropStep(string step, Action action)
    {
        try
        {
            action();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException(
                $"WORD_LIST_STYLE_COM_STEP_FAIL:{step}:HRESULT=0x{ex.HResult:X8}:{ex.Message}",
                ex);
        }
    }

    private static void FinalReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
            return;

        try { Marshal.FinalReleaseComObject(value); } catch { }
    }

    private static Dictionary<string, PersianSemanticStyleInfo> BuildPersianSemanticStyleInfo(XDocument styles)
    {
        var result = new Dictionary<string, PersianSemanticStyleInfo>(StringComparer.OrdinalIgnoreCase);

        foreach (var style in styles.Descendants(W + "style"))
        {
            var type = ((string?)style.Attribute(W + "type") ?? string.Empty).Trim();
            if (!type.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
                continue;

            var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(styleId))
                continue;

            var name = ((string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty).Trim();
            var outline = ((string?)style.Element(W + "pPr")?.Element(W + "outlineLvl")?.Attribute(W + "val") ?? string.Empty).Trim();
            result[styleId] = PersianParagraphSemanticResolver.ClassifyStyle(styleId, name, outline);
        }

        return result;
    }

    private static string GetParagraphStyleId(XElement paragraph)
        => ((string?)paragraph
            .Element(W + "pPr")
            ?.Element(W + "pStyle")
            ?.Attribute(W + "val") ?? string.Empty).Trim();

    private static bool IsAutomaticListEntryStyle(string styleId, string styleName)
    {
        var id = (styleId ?? string.Empty).Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty);
        var name = (styleName ?? string.Empty).Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty);
        return id.Equals("TableofFigures", StringComparison.OrdinalIgnoreCase) ||
               name.Equals("tableoffigures", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAutomaticListTitleStyle(string styleId, string styleName)
    {
        var id = (styleId ?? string.Empty).Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty);
        var name = (styleName ?? string.Empty).Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty);
        return name.Equals("عنوانفهرست", StringComparison.Ordinal) ||
               id.Equals("ListTitle", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAutomaticListTitleText(string text)
    {
        var compact = (text ?? string.Empty).Replace(" ", string.Empty).Replace("‌", string.Empty).Trim();
        return compact.Equals("فهرستمطالب", StringComparison.Ordinal) ||
               compact.Equals("فهرستاشکال", StringComparison.Ordinal) ||
               compact.Equals("فهرستجداول", StringComparison.Ordinal);
    }

    private static PersianSemanticStyleInfo? ResolveParagraphPersianSemanticStyleInfo(
        XElement paragraph,
        IReadOnlyDictionary<string, PersianSemanticStyleInfo> styleInfo)
    {
        var styleId = ((string?)paragraph
            .Element(W + "pPr")
            ?.Element(W + "pStyle")
            ?.Attribute(W + "val") ?? string.Empty).Trim();

        return styleInfo.TryGetValue(styleId, out var info) ? info : null;
    }

    private static bool NormalizeDocumentDefaults(XDocument styles)
    {
        var root = styles.Root;
        if (root is null)
            return false;

        var changed = false;
        var docDefaults = root.Element(W + "docDefaults");
        if (docDefaults is null)
        {
            docDefaults = new XElement(W + "docDefaults");
            root.AddFirst(docDefaults);
            changed = true;
        }

        var rPrDefault = docDefaults.Element(W + "rPrDefault");
        if (rPrDefault is null)
        {
            rPrDefault = new XElement(W + "rPrDefault");
            docDefaults.AddFirst(rPrDefault);
            changed = true;
        }

        var rPr = rPrDefault.Element(W + "rPr");
        if (rPr is null)
        {
            rPr = new XElement(W + "rPr");
            rPrDefault.Add(rPr);
            changed = true;
        }

        if (EnsureComplexScriptFont(rPr, "B Nazanin"))
            changed = true;
        if (EnsurePersianLanguage(rPr, setPrimaryLanguage: true))
            changed = true;

        return changed;
    }

    private static bool EnsureFullFontFamily(XElement rPr, string font)
    {
        var changed = false;
        var rFonts = rPr.Element(W + "rFonts");
        if (rFonts is null)
        {
            rFonts = new XElement(W + "rFonts");
            InsertInSchemaOrder(rPr, rFonts, RPrOrder);
            changed = true;
        }

        foreach (var attributeName in new[] { "ascii", "hAnsi", "eastAsia", "cs" })
        {
            var current = ((string?)rFonts.Attribute(W + attributeName) ?? string.Empty).Trim();
            if (!current.Equals(font, StringComparison.Ordinal))
            {
                rFonts.SetAttributeValue(W + attributeName, font);
                changed = true;
            }
        }

        foreach (var themeAttributeName in new[] { "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme" })
        {
            var theme = rFonts.Attribute(W + themeAttributeName);
            if (theme is not null)
            {
                theme.Remove();
                changed = true;
            }
        }

        return changed;
    }

    private static bool EnsureComplexScriptFont(XElement rPr, string font)
    {
        var changed = false;
        var rFonts = rPr.Element(W + "rFonts");
        if (rFonts is null)
        {
            rFonts = new XElement(W + "rFonts");
            InsertInSchemaOrder(rPr, rFonts, RPrOrder);
            changed = true;
        }

        var current = ((string?)rFonts.Attribute(W + "cs") ?? string.Empty).Trim();
        if (!current.Equals(font, StringComparison.Ordinal))
        {
            rFonts.SetAttributeValue(W + "cs", font);
            changed = true;
        }

        var csTheme = rFonts.Attribute(W + "cstheme");
        if (csTheme is not null)
        {
            csTheme.Remove();
            changed = true;
        }

        return changed;
    }

    private static bool EnsureFontSizeHalfPoints(XElement rPr, int halfPoints)
    {
        var changed = false;
        var expected = halfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture);
        foreach (var localName in new[] { "sz", "szCs" })
        {
            var element = rPr.Element(W + localName);
            if (element is null)
            {
                element = new XElement(W + localName);
                InsertInSchemaOrder(rPr, element, RPrOrder);
                changed = true;
            }

            var current = ((string?)element.Attribute(W + "val") ?? string.Empty).Trim();
            if (!current.Equals(expected, StringComparison.Ordinal))
            {
                element.SetAttributeValue(W + "val", expected);
                changed = true;
            }
        }

        return changed;
    }


    private static bool EnsureToggleOn(XElement rPr, string localName)
    {
        var element = rPr.Element(W + localName);
        if (element is null)
        {
            element = new XElement(W + localName);
            InsertInSchemaOrder(rPr, element, RPrOrder);
            return true;
        }

        var current = ((string?)element.Attribute(W + "val") ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(current) ||
            current.Equals("1", StringComparison.OrdinalIgnoreCase) ||
            current.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            current.Equals("on", StringComparison.OrdinalIgnoreCase))
            return false;

        element.SetAttributeValue(W + "val", "1");
        return true;
    }

    private static bool EnsureDottedPageNumberTab(XElement pPr, int position)
    {
        var changed = false;
        var tabs = pPr.Element(W + "tabs");
        if (tabs is null)
        {
            tabs = new XElement(W + "tabs");
            InsertInSchemaOrder(pPr, tabs, PPrOrder);
            changed = true;
        }

        // One deterministic dotted leader authority for all automatic-list entry styles.
        var existing = tabs.Elements(W + "tab").ToList();
        if (existing.Count != 1 ||
            !((string?)existing[0].Attribute(W + "val") ?? string.Empty).Equals("right", StringComparison.OrdinalIgnoreCase) ||
            !((string?)existing[0].Attribute(W + "leader") ?? string.Empty).Equals("dot", StringComparison.OrdinalIgnoreCase) ||
            !((string?)existing[0].Attribute(W + "pos") ?? string.Empty).Equals(position.ToString(System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal))
        {
            tabs.RemoveNodes();
            tabs.Add(new XElement(W + "tab",
                new XAttribute(W + "val", "right"),
                new XAttribute(W + "leader", "dot"),
                new XAttribute(W + "pos", position.ToString(System.Globalization.CultureInfo.InvariantCulture))));
            changed = true;
        }

        return changed;
    }

    private static bool EnsureToggleOff(XElement rPr, string localName)
    {
        var element = rPr.Element(W + localName);
        if (element is null)
        {
            element = new XElement(W + localName);
            InsertInSchemaOrder(rPr, element, RPrOrder);
        }

        var current = ((string?)element.Attribute(W + "val") ?? string.Empty).Trim();
        if (current.Equals("0", StringComparison.OrdinalIgnoreCase) ||
            current.Equals("false", StringComparison.OrdinalIgnoreCase) ||
            current.Equals("off", StringComparison.OrdinalIgnoreCase))
            return false;

        element.SetAttributeValue(W + "val", "0");
        return true;
    }

    private static bool EnsurePersianLanguage(XElement rPr, bool setPrimaryLanguage)
    {
        var changed = false;
        var lang = rPr.Element(W + "lang");
        if (lang is null)
        {
            lang = new XElement(W + "lang");
            InsertInSchemaOrder(rPr, lang, RPrOrder);
            changed = true;
        }

        if (setPrimaryLanguage)
        {
            var primary = ((string?)lang.Attribute(W + "val") ?? string.Empty).Trim();
            if (!primary.Equals("fa-IR", StringComparison.OrdinalIgnoreCase))
            {
                lang.SetAttributeValue(W + "val", "fa-IR");
                changed = true;
            }
        }

        var bidi = ((string?)lang.Attribute(W + "bidi") ?? string.Empty).Trim();
        if (!bidi.Equals("fa-IR", StringComparison.OrdinalIgnoreCase))
        {
            lang.SetAttributeValue(W + "bidi", "fa-IR");
            changed = true;
        }

        // Self-heal any legacy/output package where a base w:lang was previously
        // appended after an Office extension child such as w14:ligatures.
        if (lang.ElementsBeforeSelf().Any(x => x.Name.Namespace != W))
        {
            lang.Remove();
            InsertInSchemaOrder(rPr, lang, RPrOrder);
            changed = true;
        }

        return changed;
    }

    private static void ValidateVisualLayoutContract(
        XDocument document,
        IReadOnlyDictionary<string, PersianSemanticStyleInfo> styleInfo,
        PersianDocumentFoundationOptions options,
        string partName)
    {
        var thesisProfile = PersianParagraphSemanticResolver.IsThesisOrDissertationProfile(options.ProfileId);
        var paragraphIndex = 0;

        foreach (var table in document.Descendants(W + "tbl"))
        {
            var tblPr = table.Element(W + "tblPr")
                ?? throw new InvalidOperationException("Visual contract: table properties are missing.");

            var tblW = tblPr.Element(W + "tblW");
            var jc = tblPr.Element(W + "jc");
            var layout = tblPr.Element(W + "tblLayout");

            var widthType = ((string?)tblW?.Attribute(W + "type") ?? string.Empty).Trim();
            var widthValue = ((string?)tblW?.Attribute(W + "w") ?? string.Empty).Trim();
            var tableAlignment = ((string?)jc?.Attribute(W + "val") ?? string.Empty).Trim();
            var layoutType = ((string?)layout?.Attribute(W + "type") ?? string.Empty).Trim();

            if (!widthType.Equals("pct", StringComparison.OrdinalIgnoreCase) ||
                !widthValue.Equals("5000", StringComparison.OrdinalIgnoreCase) ||
                !tableAlignment.Equals("center", StringComparison.OrdinalIgnoreCase) ||
                !layoutType.Equals("autofit", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Visual contract: table must be centered and Fit-to-Window (pct=5000, autofit).");
            }
        }

        foreach (var paragraph in document.Descendants(W + "p"))
        {
            paragraphIndex++;
            var text = ParagraphText(paragraph);
            var hasVisibleText = !string.IsNullOrWhiteSpace(text);
            var hasPersian = ContainsPersianScript(text);
            var inTable = paragraph.Ancestors(W + "tbl").Any();
            var inTextBox = paragraph.Ancestors(W + "txbxContent").Any();
            var hasVisual = paragraph.Descendants(W + "drawing").Any() ||
                            paragraph.Descendants(W + "pict").Any();
            var pPr = paragraph.Element(W + "pPr");
            var alignment = pPr is null ? string.Empty : GetAlignment(pPr);

            if (hasVisual &&
                !alignment.Equals("center", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Visual contract: figure/image host paragraph is not centered.");
            }

            if (!hasVisual &&
                hasVisibleText &&
                (inTable || inTextBox))
            {
                var expectedAlignment = hasPersian ? "start" : "right";
                if (!alignment.Equals(expectedAlignment, StringComparison.OrdinalIgnoreCase))
                {
                    var diagnosticText = text.Length <= 120 ? text : text[..120];
                    throw new InvalidOperationException(
                        hasPersian
                            ? $"Visual contract: Persian table/text-box content is not direction-aware visual-right aligned (w:jc=start with bidi) before the thesis Word round-trip. part={partName}; jc={alignment}; text={diagnosticText}"
                            : $"Visual contract: non-Persian table/text-box content is not right aligned. part={partName}; jc={alignment}; text={diagnosticText}");
                }

                // Table/text-box layout authority is terminal for this pre-Word validation pass.
                // Do not re-apply body/heading semantic alignment rules to the same paragraph.
                continue;
            }

            if (!hasPersian)
                continue;

            var info = ResolveParagraphPersianSemanticStyleInfo(paragraph, styleInfo);
            var semantic = PersianParagraphSemanticResolver.Resolve(
                new PersianParagraphSemanticContext(
                    Text: text,
                    Style: info,
                    ParagraphIndex: paragraphIndex,
                    IsMainDocumentPart: partName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase),
                    InTable: inTable,
                    InTextBox: inTextBox,
                    HasVisual: hasVisual,
                    IsFieldDriven: paragraph.Descendants(W + "instrText").Any() ||
                                   paragraph.Descendants(W + "fldChar").Any(),
                    IsThesisProfile: thesisProfile));
            var isTocTitle = semantic.TocTitle;
            var isTocEntry = semantic.TocEntry;
            var paragraphStyleId = GetParagraphStyleId(paragraph);
            var isAutomaticListEntry = isTocEntry || IsAutomaticListEntryStyle(paragraphStyleId, string.Empty);
            var isAutomaticListTitle = !semantic.FieldDriven && (isTocTitle || IsAutomaticListTitleText(text));
            var isTitle = semantic.Title;
            var isChapter = semantic.Chapter;
            var isHeading = semantic.Heading;
            var isCaption = semantic.Caption;
            var isFieldDriven = semantic.FieldDriven;
            var isLabel = semantic.Label;
            // Keep Apply and VerifyVisualContract on the same physical-caption predicate.
                // Owner typography authority: every physical figure/table caption is B Nazanin 12pt regular/non-bold.
            var isPositionOnlyFigureCaption =
                !isFieldDriven &&
                LooksLikePositionOnlyFigureCaption(paragraph, text);
            var isPositionOnlyTableCaption =
                LooksLikePositionOnlyTableCaption(paragraph, text);
            var isPhysicalCaption =
                info?.Caption == true ||
                (!isFieldDriven &&
                 LooksLikeExplicitCaptionSyntax(text) &&
                 IsInExpectedCaptionPosition(paragraph, text)) ||
                isPositionOnlyFigureCaption ||
                isPositionOnlyTableCaption;
            var isFigureCaption =
                !isFieldDriven &&
                IsFigureCaptionPosition(paragraph) &&
                (info?.Caption == true ||
                 LooksLikeExplicitCaptionSyntax(text) ||
                 isPositionOnlyFigureCaption);
            var explicitThesisChapterTitle = semantic.ExplicitThesisChapterTitle;
            var earlyCoverLine = semantic.EarlyCoverLine;

            var preserveGenericNativeAutomaticListFieldRuns =
                !thesisProfile && isFieldDriven &&
                paragraph.Descendants(W + "instrText").Any(x =>
                    (((string?)x ?? string.Empty).Contains("TOC", StringComparison.OrdinalIgnoreCase) ||
                     ((string?)x ?? string.Empty).Contains("TOF", StringComparison.OrdinalIgnoreCase)));

            if ((isAutomaticListEntry || isAutomaticListTitle) && !preserveGenericNativeAutomaticListFieldRuns)
            {
                foreach (var run in paragraph.Descendants(W + "r"))
                {
                    var runText = RunText(run);
                    if (!ContainsPersianScript(runText))
                        continue;

                    var rFonts = run
                        .Element(W + "rPr")
                        ?.Element(W + "rFonts");

                    var csFont = ((string?)rFonts?.Attribute(W + "cs") ?? string.Empty).Trim();
                    var asciiFont = ((string?)rFonts?.Attribute(W + "ascii") ?? string.Empty).Trim();
                    var hAnsiFont = ((string?)rFonts?.Attribute(W + "hAnsi") ?? string.Empty).Trim();

                    if (!csFont.Equals("B Nazanin", StringComparison.Ordinal) ||
                        !asciiFont.Equals("B Nazanin", StringComparison.Ordinal) ||
                        !hAnsiFont.Equals("B Nazanin", StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            "Visual contract: automatic-list run does not explicitly use B Nazanin for cs/ascii/hAnsi.");
                    }
                }
            }

            if (hasVisual || inTable || inTextBox)
                continue;

            if (isPhysicalCaption)
            {
                if (!alignment.Equals("center", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "Visual contract: figure/table caption is not centered.");

                if (isPhysicalCaption)
                {
                    foreach (var run in paragraph.Elements(W + "r"))
                    {
                        if (string.IsNullOrWhiteSpace(RunText(run)))
                            continue;

                        var rPr = run.Element(W + "rPr")
                            ?? throw new InvalidOperationException(
                                "Visual contract: physical figure/table caption run properties are missing.");
                        var rFonts = rPr.Element(W + "rFonts")
                            ?? throw new InvalidOperationException(
                                "Visual contract: physical figure/table caption fonts are missing.");

                        foreach (var slot in new[] { "ascii", "hAnsi", "eastAsia", "cs" })
                        {
                            var font = ((string?)rFonts.Attribute(W + slot) ?? string.Empty).Trim();
                            if (!font.Equals("B Nazanin", StringComparison.Ordinal))
                                throw new InvalidOperationException(
                                    $"Visual contract: physical figure/table caption font slot {slot} is not B Nazanin.");
                        }

                        var sz = ((string?)rPr.Element(W + "sz")?.Attribute(W + "val") ?? string.Empty).Trim();
                        var szCs = ((string?)rPr.Element(W + "szCs")?.Attribute(W + "val") ?? string.Empty).Trim();
                        if (!sz.Equals("24", StringComparison.Ordinal) ||
                            !szCs.Equals("24", StringComparison.Ordinal))
                        {
                            throw new InvalidOperationException(
                                "Visual contract: physical figure/table caption must be 12pt (w:sz=24 and w:szCs=24).");
                        }

                        foreach (var toggleName in new[] { "b", "bCs" })
                        {
                            var toggle = rPr.Element(W + toggleName);
                            var value = ((string?)toggle?.Attribute(W + "val") ?? string.Empty).Trim();
                            var explicitlyOff =
                                value.Equals("0", StringComparison.OrdinalIgnoreCase) ||
                                value.Equals("false", StringComparison.OrdinalIgnoreCase) ||
                                value.Equals("off", StringComparison.OrdinalIgnoreCase);
                            if (!explicitlyOff)
                                throw new InvalidOperationException(
                                    $"Visual contract: physical figure/table caption {toggleName} must be explicitly off (regular/non-bold).");
                        }
                    }
                }
            }
            else if (isAutomaticListEntry)
            {
                if (!alignment.Equals("start", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "Visual contract: RTL automatic-list entry is not direction-aware visual-right aligned (w:jc=start).");
            }
            else if (isAutomaticListTitle || isTitle || earlyCoverLine || explicitThesisChapterTitle)
            {
                if (!alignment.Equals("center", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "Visual contract: cover/title/explicit thesis chapter line is not centered.");
            }
            else if ((isHeading || isChapter) && options.NormalizeHeadingAlignment)
            {
                if (!alignment.Equals("start", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "Visual contract: RTL in-document heading is not direction-aware visual-right aligned (w:jc=start).");
            }
            else if (options.NormalizeBodyAlignment &&
                     !isFieldDriven &&
                     !isPhysicalCaption &&
                     !isLabel)
            {
                if (!alignment.Equals("both", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "Visual contract: Persian body prose is not justified (w:jc=both).");
            }
        }
    }

    private static bool LooksLikeExplicitCaptionSyntax(string text)
    {
        var value = (text ?? string.Empty).Trim();
        var prefixes = new[] { "شکل", "جدول", "نمودار" };
        var prefix = prefixes.FirstOrDefault(p =>
            value.StartsWith(p, StringComparison.Ordinal) &&
            value.Length > p.Length &&
            char.IsWhiteSpace(value[p.Length]));
        if (prefix is null)
            return false;

        var i = prefix.Length;
        while (i < value.Length && char.IsWhiteSpace(value[i])) i++;

        const string numberWord = "شماره";
        if (i < value.Length &&
            value.AsSpan(i).StartsWith(numberWord.AsSpan(), StringComparison.Ordinal))
        {
            i += numberWord.Length;
            while (i < value.Length && char.IsWhiteSpace(value[i])) i++;
        }

        static bool IsDigitLike(char c)
            => (c >= '0' && c <= '9') || (c >= '۰' && c <= '۹');

        static bool IsTokenChar(char c)
            => IsDigitLike(c) ||
               (c >= 'A' && c <= 'Z') ||
               (c >= 'a' && c <= 'z') ||
               (c >= '\u0600' && c <= '\u06FF');

        static bool IsDash(char c)
            => c == '-' || c == '–' || c == '—' || c == '‑';

        static bool IsCaptionSeparator(char c)
            => IsDash(c) || c == ':' || c == '؛' || c == 'ـ';

        var bracketed = false;
        var expectedClose = '\0';
        if (i < value.Length && (value[i] == '(' || value[i] == '['))
        {
            bracketed = true;
            expectedClose = value[i] == '(' ? ')' : ']';
            i++;
            while (i < value.Length && char.IsWhiteSpace(value[i])) i++;
        }

        var digitStart = i;
        while (i < value.Length && IsDigitLike(value[i])) i++;
        if (i == digitStart)
            return false;

        // Compact hierarchical/alphanumeric identifier segments, e.g. ۵-۱ or ۵-الف.
        while (i + 1 < value.Length && IsDash(value[i]) && IsTokenChar(value[i + 1]))
        {
            i++;
            var tokenStart = i;
            while (i < value.Length && IsTokenChar(value[i])) i++;
            if (i == tokenStart)
                return false;
        }

        if (bracketed)
        {
            while (i < value.Length && char.IsWhiteSpace(value[i])) i++;
            if (i >= value.Length || value[i] != expectedClose)
                return false;

            i++;
            if (i >= value.Length)
                return true;

            // The closing bracket itself is an explicit identifier delimiter. Require either
            // whitespace or a conventional caption separator before any following description.
            return char.IsWhiteSpace(value[i]) || IsCaptionSeparator(value[i]);
        }

        // For unbracketed identifiers, narrative text such as "شکل ۵ توزیع فشار را ..."
        // must remain body text. Require an explicit separator after the identifier.
        while (i < value.Length && char.IsWhiteSpace(value[i])) i++;
        return i < value.Length && IsCaptionSeparator(value[i]);
    }

    private static bool IsFigureCaptionPosition(XElement paragraph)
    {
        var previous = paragraph.ElementsBeforeSelf().LastOrDefault();
        return previous?.Name == W + "p" &&
               (previous.Descendants(W + "drawing").Any() ||
                previous.Descendants(W + "pict").Any());
    }

    private static bool LooksLikePositionOnlyFigureCaption(XElement paragraph, string text)
    {
        if (!IsFigureCaptionPosition(paragraph))
            return false;

        var value = (text ?? string.Empty).Trim();
        if (value.Length == 0 || value.Length > 240)
            return false;

        var hasTocBookmark = paragraph
            .Descendants(W + "bookmarkStart")
            .Select(x => ((string?)x.Attribute(W + "name") ?? string.Empty).Trim())
            .Any(name => name.StartsWith("_Toc", StringComparison.OrdinalIgnoreCase));

        return hasTocBookmark;
    }

    private static bool LooksLikePositionOnlyTableCaption(XElement paragraph, string text)
    {
        var next = paragraph.ElementsAfterSelf().FirstOrDefault();
        if (next?.Name != W + "tbl")
            return false;

        var value = (text ?? string.Empty).Trim();
        if (value.Length == 0 || value.Length > 240)
            return false;

        var hasTocBookmark = paragraph
            .Descendants(W + "bookmarkStart")
            .Select(x => ((string?)x.Attribute(W + "name") ?? string.Empty).Trim())
            .Any(name => name.StartsWith("_Toc", StringComparison.OrdinalIgnoreCase));

        return hasTocBookmark;
    }

    private static bool IsInExpectedCaptionPosition(XElement paragraph, string text)
    {
        var previous = paragraph.ElementsBeforeSelf().LastOrDefault();
        var next = paragraph.ElementsAfterSelf().FirstOrDefault();

        static bool IsVisualParagraph(XElement? element)
            => element?.Name == W + "p" &&
               (element.Descendants(W + "drawing").Any() ||
                element.Descendants(W + "pict").Any());

        static bool IsTable(XElement? element)
            => element?.Name == W + "tbl";

        var value = (text ?? string.Empty).TrimStart();

        // Product contract:
        // - table title/caption is above its table;
        // - figure/chart caption is below its visual.
        if (value.StartsWith("جدول ", StringComparison.Ordinal))
            return IsTable(next);

        if (value.StartsWith("شکل ", StringComparison.Ordinal) ||
            value.StartsWith("نمودار ", StringComparison.Ordinal))
            return IsVisualParagraph(previous);

        return false;
    }

    private static bool NormalizeTables(XDocument document)
    {
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

            var tblW = EnsureChild(tblPr, "tblW", TblPrOrder);
            if (SetWordAttribute(tblW, "type", "pct")) changed = true;
            if (SetWordAttribute(tblW, "w", "5000")) changed = true;

            var jc = EnsureChild(tblPr, "jc", TblPrOrder);
            if (SetWordAttribute(jc, "val", "center")) changed = true;

            var layout = EnsureChild(tblPr, "tblLayout", TblPrOrder);
            if (SetWordAttribute(layout, "type", "autofit")) changed = true;

            var tblInd = tblPr.Element(W + "tblInd");
            if (tblInd is not null)
            {
                tblInd.Remove();
                changed = true;
            }
        }

        return changed;
    }

    private static bool SetWordAttribute(XElement element, string localName, string value)
    {
        var attribute = element.Attribute(W + localName);
        var current = ((string?)attribute ?? string.Empty).Trim();
        if (current.Equals(value, StringComparison.OrdinalIgnoreCase))
            return false;

        element.SetAttributeValue(W + localName, value);
        return true;
    }

    private static bool SetAlignment(XElement pPr, string value)
    {
        var jc = pPr.Element(W + "jc");
        if (jc is null)
        {
            InsertInSchemaOrder(pPr, new XElement(W + "jc", new XAttribute(W + "val", value)), PPrOrder);
            return true;
        }

        var current = ((string?)jc.Attribute(W + "val") ?? string.Empty).Trim();
        if (current.Equals(value, StringComparison.OrdinalIgnoreCase))
            return false;

        jc.SetAttributeValue(W + "val", value);
        return true;
    }

    private static string GetAlignment(XElement pPr)
        => ((string?)pPr.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();

    private static bool EnsureOn(XElement parent, string localName, IReadOnlyList<string> order)
    {
        var element = parent.Element(W + localName);
        if (element is null)
        {
            InsertInSchemaOrder(parent, new XElement(W + localName), order);
            return true;
        }

        var val = element.Attribute(W + "val");
        if (val is null)
            return false;

        var raw = ((string?)val ?? string.Empty).Trim();
        if (raw.Equals("0", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("false", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("off", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("1", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("on", StringComparison.OrdinalIgnoreCase))
        {
            element.SetAttributeValue(W + "val", null);
            return true;
        }

        return false;
    }

    private static XElement EnsureChild(
        XElement parent,
        string localName,
        IReadOnlyList<string>? order)
    {
        var child = parent.Element(W + localName);
        if (child is not null)
            return child;

        child = new XElement(W + localName);
        if (order is null)
            parent.Add(child);
        else
            InsertInSchemaOrder(parent, child, order);
        return child;
    }

    private static void InsertInSchemaOrder(
        XElement parent,
        XElement child,
        IReadOnlyList<string> order)
    {
        var targetIndex = IndexOf(order, child.Name.LocalName);
        if (targetIndex < 0)
        {
            parent.Add(child);
            return;
        }

        foreach (var existing in parent.Elements())
        {
            // In CT_RPr, Office extension children (for example w14:ligatures) follow
            // the base WordprocessingML run properties.  A missing base property such
            // as w:lang must therefore be inserted before the first foreign-namespace
            // extension element, not appended after it.
            if (ReferenceEquals(order, RPrOrder) && existing.Name.Namespace != W)
            {
                existing.AddBeforeSelf(child);
                return;
            }

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
            if (values[i].Equals(value, StringComparison.Ordinal))
                return i;
        return -1;
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
        if (name.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) &&
            name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) &&
            name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static XDocument LoadXmlEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name)
            ?? throw new InvalidOperationException("Missing package entry: " + name);
        using var entryStream = entry.Open();
        return XDocument.Load(entryStream, LoadOptions.PreserveWhitespace);
    }

    private static void ReplaceXmlEntry(ZipArchive archive, string name, XDocument document)
    {
        var existing = archive.GetEntry(name);
        existing?.Delete();

        var replacement = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var stream = replacement.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(false),
            OmitXmlDeclaration = false,
            Indent = false,
            CloseOutput = false
        });
        document.Save(writer);
    }

    private static string ParagraphText(XElement paragraph)
        => string.Concat(paragraph.Descendants(W + "t").Select(x => (string?)x ?? string.Empty)).Trim();

    private static string RunText(XElement run)
        => string.Concat(run.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));

    private static bool NormalizePersianSectionLayouts(XDocument document)
    {
        var body = document.Root?.Element(W + "body");
        if (body is null) return false;
        bool changed = false;
        bool sectionHasPersian = false;
        foreach (var paragraph in body.Descendants(W + "p")
                     .Where(p => !p.Ancestors(W + "sectPrChange").Any()
                              && !p.Ancestors(W + "pPrChange").Any()
                              && !p.Ancestors(W + "del").Any()
                              && !p.Ancestors(W + "moveFrom").Any()
                              && !p.Ancestors(W + "txbxContent").Any()))
        {
            var currentText = string.Concat(paragraph.Descendants(W + "t")
                .Where(t => !t.Ancestors().Any(a => a.Name == W + "del"
                    || a.Name == W + "moveFrom" || a.Name == W + "txbxContent"
                    || a.Name == W + "pPrChange" || a.Name == W + "sectPrChange"))
                .Select(t => t.Value));
            sectionHasPersian |= ContainsPersianScript(currentText);
            var section = paragraph.Element(W + "pPr")?.Element(W + "sectPr");
            if (section is null) continue;
            if (sectionHasPersian) changed |= EnsureRtlSectionLayout(section);
            sectionHasPersian = false;
        }
        if (sectionHasPersian)
            foreach (var section in body.Elements(W + "sectPr"))
                changed |= EnsureRtlSectionLayout(section);
        return changed;
    }

    private static bool EnsureRtlSectionLayout(XElement sectionProperties)
    {
        var bidi = sectionProperties.Element(W + "bidi");
        if (bidi is not null)
        {
            var value = bidi.Attribute(W + "val");
            if (value is null) return false;
            value.Remove();
            return true;
        }
        bidi = new XElement(W + "bidi");
        var later = sectionProperties.Elements().FirstOrDefault(x =>
            x.Name == W + "rtlGutter" || x.Name == W + "docGrid" ||
            x.Name == W + "printerSettings" || x.Name.LocalName == "footnoteColumns" ||
            x.Name == W + "sectPrChange");
        if (later is null) sectionProperties.Add(bidi);
        else later.AddBeforeSelf(bidi);
        return true;
    }

    private static bool ContainsPersianScript(string value)
    {
        foreach (var ch in value)
        {
            if ((ch >= '\u0600' && ch <= '\u06FF') ||
                (ch >= '\u0750' && ch <= '\u077F') ||
                (ch >= '\u08A0' && ch <= '\u08FF') ||
                (ch >= '\uFB50' && ch <= '\uFDFF') ||
                (ch >= '\uFE70' && ch <= '\uFEFF'))
            {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsLatinScript(string value)
    {
        foreach (var ch in value)
        {
            if ((ch >= 'A' && ch <= 'Z') ||
                (ch >= 'a' && ch <= 'z'))
            {
                return true;
            }
        }

        return false;
    }

}
