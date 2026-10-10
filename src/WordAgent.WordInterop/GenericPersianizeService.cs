using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using WordAgent.Core;

namespace WordAgent.WordInterop;

public sealed record GenericPersianizeInventory(
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
    int TocStyles,
    string TextPayloadSha256,
    string FieldInstructionSha256);

public sealed record GenericPersianizeResult(
    string InputFile,
    string OutputFile,
    string InputSha256,
    string OutputSha256,
    int ParagraphsExamined,
    int PersianParagraphs,
    int ParagraphsChanged,
    int BodyParagraphsJustified,
    int RunsChanged,
    int TablesExamined,
    int TablesCentered,
    int PersianTablesRtlEnabled,
    int TocStylesChanged,
    bool AutomaticFieldUpdateDisabled,
    bool NoChangesNeeded,
    GenericPersianizeInventory Before,
    GenericPersianizeInventory After,
    IReadOnlyList<string> IntegrityIssues,
    IReadOnlyList<string> Warnings);

public sealed class GenericPersianizeService
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

    private static readonly string[] TblPrOrder =
    [
        "tblStyle","tblpPr","tblOverlap","bidiVisual","tblStyleRowBandSize","tblStyleColBandSize",
        "tblW","jc","tblCellSpacing","tblInd","tblBorders","shd","tblLayout","tblCellMar",
        "tblLook","tblCaption","tblDescription","tblPrChange"
    ];

    public GenericPersianizeResult Apply(string inputPath, string? outputDirectory = null)
    {
        var input = PathPolicy.RequireSupportedWordFile(inputPath);
        Preflight(input);

        var before = InspectPackage(input);
        if (before.RevisionElements != 0)
            throw new InvalidOperationException(
                $"Generic Persianize blocked: tracked-revision XML elements={before.RevisionElements}. Accept/reject revisions first; this mode never rewrites tracked-change markup.");

        var inputHash = Sha256File(input);
        var output = PathPolicy.CreateSafeOutputPath(input, outputDirectory, "_WordAgent_Persianized");
        var transaction = CreateTransactionalTempPath(output);
        File.Copy(input, transaction, overwrite: false);
        PauseForTransactionTest();

        var warnings = new List<string>
        {
            "اصلاح عمومی RTL از پروفایل خانواده سند یا تشخیص موضوع استفاده نمی‌کند؛ برای اصلاح کامل از مسیر هوشمند پروفایل‌محور استفاده شود.",
            "حاشیه صفحه، بخش‌ها، شماره‌گذاری، شکست صفحه، کادر، سایه‌زنی، Field، تصویر، Relationship و متن حفظ می‌شوند؛ Persian Foundation بدنه را B Nazanin/Justify، تیترها را B Titr/راست‌چین، فهرست‌ها را B Nazanin، جدول‌ها را وسط/Fit-to-Window و شکل‌ها را وسط استاندارد می‌کند.",
            "پاراگراف فارسی راست‌به‌چپ می‌شود؛ متن بدنه در سند اصلی، حتی بدون Style و حتی اگر قبلاً Center/Right بوده باشد، با تشخیص متنی عنوان/تیتر/برچسب/Caption/Field به Justify کامل تبدیل می‌شود.",
            "همه جدول‌ها وسط‌چین می‌شوند و ترتیب بصری RTL فقط برای جدول‌های دارای متن فارسی/عربی فعال می‌شود.",
            "دستور و محتوای Field فهرست مطالب حفظ می‌شود؛ عنوان و ورودی‌های فارسی فهرست توسط Persian Foundation با B Nazanin و جهت/چیدمان متناسب استاندارد می‌شوند.",
            "اگر به‌روزرسانی خودکار Field هنگام بازشدن فعال باشد، در خروجی غیرفعال می‌شود؛ Fieldها حذف نمی‌شوند و فقط با درخواست صریح کاربر به‌روزرسانی خواهند شد.",
            "خروجی ابتدا در فایل تراکنشی جدا ساخته می‌شود و فقط پس از عبور از همه گیت‌های یکپارچگی به مسیر نهایی منتقل می‌شود."
        };

        var immutableBefore = ReadEntryHashes(input);
        var mutatedEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        int paragraphsExamined = 0;
        int persianParagraphs = 0;
        int paragraphsChanged = 0;
        int bodyParagraphsJustified = 0;
        int runsChanged = 0;
        int tablesExamined = 0;
        int tablesCentered = 0;
        int persianTablesRtlEnabled = 0;
        int tocStylesChanged = 0;
        bool automaticFieldUpdateDisabled = false;

        try
        {
            using var file = new FileStream(transaction, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            using var archive = new ZipArchive(file, ZipArchiveMode.Update, leaveOpen: false);

            if (archive.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException(
                    "Digitally signed Office package detected; Generic Persianize is blocked because any package mutation would invalidate the signature.");

            var storyParts = GetStoryPartNames(archive);
            var documentHasPersian = false;
            var styleInfo = archive.GetEntry("word/styles.xml") is not null
                ? BuildStyleInfo(LoadXmlEntry(archive, "word/styles.xml"))
                : new Dictionary<string, PersianSemanticStyleInfo>(StringComparer.OrdinalIgnoreCase);

            foreach (var partName in storyParts)
            {
                var document = LoadXmlEntry(archive, partName);
                var partChanged = false;
                var allowBodyJustification = partName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase);
                var paragraphIndex = 0;

                foreach (var paragraph in document.Descendants(W + "p"))
                {
                    paragraphsExamined++;
                    paragraphIndex++;
                    var text = ParagraphText(paragraph);
                    if (!ContainsPersianScript(text)) continue;

                    documentHasPersian = true;
                    persianParagraphs++;
                    var semantic = ResolveSemanticDecision(
                        paragraph,
                        text,
                        styleInfo,
                        paragraphIndex,
                        allowBodyJustification);
                    var bodyJustified = false;
                    var paragraphChanged = EnsurePersianParagraph(
                        paragraph,
                        semantic,
                        allowBodyJustification,
                        ref runsChanged,
                        out bodyJustified);
                    if (bodyJustified) bodyParagraphsJustified++;
                    if (paragraphChanged)
                    {
                        paragraphsChanged++;
                        partChanged = true;
                    }
                }

                foreach (var table in document.Descendants(W + "tbl"))
                {
                    tablesExamined++;
                    var tableChanged = false;

                    if (EnsureTableCentered(table))
                    {
                        tablesCentered++;
                        tableChanged = true;
                    }

                    if (ContainsPersianScript(string.Concat(table.Descendants(W + "t").Select(x => (string?)x ?? string.Empty))) &&
                        EnsureTableRtlVisual(table))
                    {
                        persianTablesRtlEnabled++;
                        tableChanged = true;
                    }

                    if (tableChanged) partChanged = true;
                }

                if (partChanged)
                {
                    ReplaceXmlEntry(archive, partName, document);
                    mutatedEntries.Add(partName);
                }
            }

            if (documentHasPersian && archive.GetEntry("word/styles.xml") is not null)
            {
                var styles = LoadXmlEntry(archive, "word/styles.xml");
                tocStylesChanged = NormalizeTocStyles(styles);
                if (tocStylesChanged > 0)
                {
                    ReplaceXmlEntry(archive, "word/styles.xml", styles);
                    mutatedEntries.Add("word/styles.xml");
                }
            }

            // Field instructions and relationships remain byte-semantically intact.  Only the
            // "update fields automatically when the document opens" setting is normalized OFF.
            // This aligns with update-on-request policy and prevents Word from prompting the user
            // to refresh fields merely because the output copy is opened.
            if (archive.GetEntry("word/settings.xml") is not null)
            {
                var settings = LoadXmlEntry(archive, "word/settings.xml");
                if (DisableAutomaticFieldUpdateOnOpen(settings))
                {
                    ReplaceXmlEntry(archive, "word/settings.xml", settings);
                    mutatedEntries.Add("word/settings.xml");
                    automaticFieldUpdateDisabled = true;
                }
            }
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
                $"Persian Foundation: paragraphs={foundation.PersianParagraphs}; " +
                $"bodyJustified={foundation.BodyParagraphsJustified}; headingsRight={foundation.HeadingParagraphsRightAligned}; " +
                $"titlesCentered={foundation.TitleParagraphsCentered}; toc={foundation.TocParagraphsNormalized}; " +
                $"runs={foundation.PersianRunsNormalized}; styles={foundation.StyleDefinitionsNormalized}.");
        }
        catch
        {
            TryDelete(transaction);
            throw;
        }

        var noChangesNeeded = mutatedEntries.Count == 0;

        // ZipArchiveMode.Update can rewrite ZIP container metadata/central-directory bytes even
        // when no XML entry was replaced. A semantic no-op must be a true byte-level no-op too:
        // restore the transaction from the authoritative input before all post-transaction gates.
        if (noChangesNeeded)
            File.Copy(input, transaction, overwrite: true);

        var after = InspectPackage(transaction);
        var issues = CompareInventory(before, after);
        issues.AddRange(CompareImmutableEntries(transaction, immutableBefore, mutatedEntries));
        issues.AddRange(VerifyGenericRules(transaction));

        var outputHash = Sha256File(transaction);

        if (noChangesNeeded && !string.Equals(inputHash, outputHash, StringComparison.OrdinalIgnoreCase))
            issues.Add("Idempotency/persistence mismatch: semantic no-op was not restored byte-identically from input.");

        if (!noChangesNeeded && string.Equals(inputHash, outputHash, StringComparison.OrdinalIgnoreCase))
            issues.Add("Persistence mismatch: mutations were reported but output SHA-256 is identical to input.");

        if (issues.Count != 0)
        {
            TryDelete(transaction);
            throw new InvalidOperationException(
                "Generic Persianize integrity gate failed. " + string.Join(" | ", issues.Take(12)));
        }

        CommitTransactionalOutput(transaction, output);

        return new GenericPersianizeResult(
            input,
            output,
            inputHash,
            outputHash,
            paragraphsExamined,
            persianParagraphs,
            paragraphsChanged,
            bodyParagraphsJustified,
            runsChanged,
            tablesExamined,
            tablesCentered,
            persianTablesRtlEnabled,
            tocStylesChanged,
            automaticFieldUpdateDisabled,
            noChangesNeeded,
            before,
            after,
            issues,
            warnings);
    }

    private static void Preflight(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        if (archive.GetEntry("word/document.xml") is null)
            throw new InvalidOperationException("Invalid Word package: word/document.xml is missing.");

        if (archive.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException(
                "Digitally signed Office package detected; Generic Persianize is blocked because package mutation would invalidate the signature.");
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

    private static bool EnsurePersianParagraph(
        XElement paragraph,
        PersianParagraphSemanticDecision semantic,
        bool allowBodyJustification,
        ref int runsChanged,
        out bool bodyJustified)
    {
        var changed = false;
        bodyJustified = false;

        var pPr = paragraph.Element(W + "pPr");
        if (pPr is null)
        {
            pPr = new XElement(W + "pPr");
            paragraph.AddFirst(pPr);
            changed = true;
        }

        if (EnsureOn(pPr, "bidi", PPrOrder)) changed = true;

        // Generic and Shared Foundation consume one semantic-role decision. This prevents
        // pre-pass/final-pass role drift and the transient rewrites that previously broke
        // exact-byte idempotency.
        var shouldJustify = allowBodyJustification && semantic.ShouldJustifyBody;
        var forceSemanticRight = semantic.VisualRightSemantic;

        if (semantic.CenteredSemantic)
        {
            var jc = pPr.Element(W + "jc");
            var value = ((string?)jc?.Attribute(W + "val") ?? string.Empty).Trim();
            if (jc is null)
            {
                InsertInSchemaOrder(
                    pPr,
                    new XElement(W + "jc", new XAttribute(W + "val", "center")),
                    PPrOrder);
                changed = true;
            }
            else if (!value.Equals("center", StringComparison.OrdinalIgnoreCase))
            {
                jc.SetAttributeValue(W + "val", "center");
                changed = true;
            }
        }
        else if (NormalizeParagraphAlignment(pPr, shouldJustify, forceSemanticRight))
        {
            changed = true;
        }

        var alignment = ((string?)pPr.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();
        bodyJustified = shouldJustify && alignment.Equals("both", StringComparison.OrdinalIgnoreCase);

        foreach (var run in paragraph.Descendants(W + "r"))
        {
            var runText = string.Concat(run.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));
            if (!ContainsPersianScript(runText)) continue;

            var runChanged = false;
            var rPr = run.Element(W + "rPr");
            if (rPr is null)
            {
                rPr = new XElement(W + "rPr");
                run.AddFirst(rPr);
                changed = true;
                runChanged = true;
            }

            if (EnsureOn(rPr, "rtl", RPrOrder))
            {
                changed = true;
                runChanged = true;
            }

            if (runChanged) runsChanged++;
        }

        return changed;
    }

    private static bool NormalizeParagraphAlignment(XElement pPr, bool shouldJustify, bool forceRight = false)
    {
        var jc = pPr.Element(W + "jc");
        var value = ((string?)jc?.Attribute(W + "val") ?? string.Empty).Trim();

        if (forceRight)
        {
            // The shared Persian Foundation owns the final semantic distinction between
            // centered document/TOC titles (and protected centered chapter roles) versus
            // ordinary in-document headings. RTL semantic-right is encoded as logical w:jc=start.
            // Preserve an existing explicit center here
            // so the Foundation can adjudicate it without a transient right->center
            // rewrite that would break second-pass byte idempotency.
            if (value.Equals("center", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("distribute", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("thaiDistribute", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (jc is null)
            {
                InsertInSchemaOrder(pPr, new XElement(W + "jc", new XAttribute(W + "val", "start")), PPrOrder);
                return true;
            }
            if (!value.Equals("start", StringComparison.OrdinalIgnoreCase))
            {
                jc.SetAttributeValue(W + "val", "start");
                return true;
            }
            return false;
        }

        // Once a paragraph has been semantically classified as body prose, its previous
        // visual alignment is not authoritative: even legacy centered/right body text is
        // normalized to full justification. Heading/caption/field guards are evaluated
        // before this point.
        if (shouldJustify)
        {
            if (jc is null)
            {
                InsertInSchemaOrder(
                    pPr,
                    new XElement(W + "jc", new XAttribute(W + "val", "both")),
                    PPrOrder);
                return true;
            }

            if (!value.Equals("both", StringComparison.OrdinalIgnoreCase))
            {
                jc.SetAttributeValue(W + "val", "both");
                return true;
            }

            return false;
        }

        // Center/distribute remains protected only for semantic non-body content.
        // Other RTL semantic-right content uses logical start so Word resolves it to the visual right edge.
        if (value.Equals("center", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("distribute", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thaiDistribute", StringComparison.OrdinalIgnoreCase))
            return false;

        if (jc is null)
        {
            InsertInSchemaOrder(
                pPr,
                new XElement(W + "jc", new XAttribute(W + "val", "start")),
                PPrOrder);
            return true;
        }

        if (value.Equals("left", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("right", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("end", StringComparison.OrdinalIgnoreCase) ||
            string.IsNullOrWhiteSpace(value))
        {
            jc.SetAttributeValue(W + "val", "start");
            return true;
        }

        // Preserve existing start/both and other explicit semantic alignments.
        return false;
    }

    private static bool NormalizeParagraphAlignment(XElement pPr)
        => NormalizeParagraphAlignment(pPr, shouldJustify: false);

    private static Dictionary<string, PersianSemanticStyleInfo> BuildStyleInfo(XDocument styles)
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
            var outlineLevel = ((string?)style.Element(W + "pPr")?.Element(W + "outlineLvl")?.Attribute(W + "val") ?? string.Empty).Trim();
            result[styleId] = PersianParagraphSemanticResolver.ClassifyStyle(styleId, name, outlineLevel);
        }

        return result;
    }

    private static PersianParagraphSemanticDecision ResolveSemanticDecision(
        XElement paragraph,
        string text,
        IReadOnlyDictionary<string, PersianSemanticStyleInfo> styleInfo,
        int paragraphIndex,
        bool isMainDocumentPart)
    {
        var styleId = ((string?)paragraph
            .Element(W + "pPr")
            ?.Element(W + "pStyle")
            ?.Attribute(W + "val") ?? string.Empty).Trim();

        styleInfo.TryGetValue(styleId, out var info);

        return PersianParagraphSemanticResolver.Resolve(
            new PersianParagraphSemanticContext(
                Text: text,
                Style: info,
                ParagraphIndex: paragraphIndex,
                IsMainDocumentPart: isMainDocumentPart,
                InTable: paragraph.Ancestors(W + "tbl").Any(),
                InTextBox: paragraph.Ancestors(W + "txbxContent").Any(),
                HasVisual: paragraph.Descendants(W + "drawing").Any() ||
                           paragraph.Descendants(W + "pict").Any(),
                IsFieldDriven: paragraph.Descendants(W + "instrText").Any() ||
                               paragraph.Descendants(W + "fldChar").Any(),
                IsThesisProfile: false));
    }

    private static bool LooksLikeMetadataLine(string text)
    {
        var trimmed = text.Trim();
        if (trimmed.Length == 0)
            return false;

        return System.Text.RegularExpressions.Regex.IsMatch(
            trimmed,
            @"^(?:وضعیت(?:\s+نهایی)?|مبنای(?:\s+علمی)?|تاریخ(?:\s+گزارش)?|شماره|کد|نسخه|موضوع|نام|تهیه‌کننده|تهیه کننده|تأیید|تایید|مصوب|پروژه)\s*[:：]",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant);
    }

    private static bool EnsureTableCentered(XElement table)
    {
        var tblPr = table.Element(W + "tblPr");
        var changed = false;
        if (tblPr is null)
        {
            tblPr = new XElement(W + "tblPr");
            table.AddFirst(tblPr);
            changed = true;
        }

        var jc = tblPr.Element(W + "jc");
        if (jc is null)
        {
            InsertInSchemaOrder(tblPr, new XElement(W + "jc", new XAttribute(W + "val", "center")), TblPrOrder);
            return true;
        }

        var value = ((string?)jc.Attribute(W + "val") ?? string.Empty).Trim();
        if (!value.Equals("center", StringComparison.OrdinalIgnoreCase))
        {
            jc.SetAttributeValue(W + "val", "center");
            changed = true;
        }

        return changed;
    }

    private static bool EnsureTableRtlVisual(XElement table)
    {
        var tblPr = table.Element(W + "tblPr");
        if (tblPr is null)
        {
            tblPr = new XElement(W + "tblPr");
            table.AddFirst(tblPr);
        }

        return EnsureOn(tblPr, "bidiVisual", TblPrOrder);
    }

    private static int NormalizeTocStyles(XDocument styles)
    {
        var changedStyles = 0;

        foreach (var style in styles.Descendants(W + "style"))
        {
            var type = ((string?)style.Attribute(W + "type") ?? string.Empty).Trim();
            if (!type.Equals("paragraph", StringComparison.OrdinalIgnoreCase)) continue;

            var styleId = ((string?)style.Attribute(W + "styleId") ?? string.Empty).Trim();
            var styleName = ((string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty).Trim();

            if (!IsTocStyle(styleId, styleName)) continue;

            var changed = false;
            var pPr = style.Element(W + "pPr");
            if (pPr is null)
            {
                pPr = new XElement(W + "pPr");
                var afterPPr = style.Elements().FirstOrDefault(x =>
                    x.Name == W + "rPr" || x.Name == W + "tblPr" || x.Name == W + "trPr" || x.Name == W + "tcPr");
                if (afterPPr is not null) afterPPr.AddBeforeSelf(pPr);
                else style.Add(pPr);
                changed = true;
            }

            if (EnsureOn(pPr, "bidi", PPrOrder)) changed = true;
            if (NormalizeParagraphAlignment(pPr)) changed = true;

            if (changed) changedStyles++;
        }

        return changedStyles;
    }

    private static bool IsTocStyle(string styleId, string styleName)
    {
        static bool Match(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            var compact = value.Replace(" ", string.Empty, StringComparison.Ordinal).Replace("-", string.Empty, StringComparison.Ordinal);
            return compact.StartsWith("TOC", StringComparison.OrdinalIgnoreCase);
        }

        return Match(styleId) || Match(styleName);
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

    private static bool EnsureOn(XElement parent, string localName, IReadOnlyList<string> order)
    {
        var element = parent.Element(W + localName);
        if (element is null)
        {
            InsertInSchemaOrder(parent, new XElement(W + localName), order);
            return true;
        }

        var val = element.Attribute(W + "val");
        if (val is null) return false;

        var raw = ((string?)val ?? string.Empty).Trim();
        if (raw.Equals("0", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("false", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("off", StringComparison.OrdinalIgnoreCase))
        {
            element.SetAttributeValue(W + "val", null);
            return true;
        }

        // Presence is ON; remove redundant true literals to avoid schema-literal drift.
        if (raw.Equals("1", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("on", StringComparison.OrdinalIgnoreCase))
        {
            element.SetAttributeValue(W + "val", null);
            return true;
        }

        return false;
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
            if (values[i].Equals(value, StringComparison.Ordinal))
                return i;
        return -1;
    }

    private static GenericPersianizeInventory InspectPackage(string path)
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
        int tocStyles = 0;

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

        var stylesEntry = archive.GetEntry("word/styles.xml");
        if (stylesEntry is not null)
        {
            using var stream = stylesEntry.Open();
            var styles = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            tocStyles = styles.Descendants(W + "style").Count(style =>
            {
                var id = (string?)style.Attribute(W + "styleId") ?? string.Empty;
                var name = (string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty;
                return IsTocStyle(id, name);
            });
        }

        return new GenericPersianizeInventory(
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
            archive.Entries.Count(e => e.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase) && !e.FullName.EndsWith("/", StringComparison.Ordinal)),
            tocStyles,
            Sha256Text(textPayload.ToString()),
            Sha256Text(fieldPayload.ToString()));
    }

    private static int CountRevisionElements(XDocument document)
        => document.Descendants(W + "ins").Count()
         + document.Descendants(W + "del").Count()
         + document.Descendants(W + "moveFrom").Count()
         + document.Descendants(W + "moveTo").Count();

    private static List<string> CompareInventory(GenericPersianizeInventory before, GenericPersianizeInventory after)
    {
        var issues = new List<string>();

        void Equal(string name, int a, int b)
        {
            if (a != b) issues.Add($"Package inventory changed unexpectedly: {name} {a} -> {b}.");
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
        Equal(nameof(before.TocStyles), before.TocStyles, after.TocStyles);

        if (!before.TextPayloadSha256.Equals(after.TextPayloadSha256, StringComparison.OrdinalIgnoreCase))
            issues.Add("Visible Word text payload changed; Generic Persianize is formatting-only and may not change text.");

        if (!before.FieldInstructionSha256.Equals(after.FieldInstructionSha256, StringComparison.OrdinalIgnoreCase))
            issues.Add("Field instruction payload changed; fields must be preserved byte-semantically.");

        return issues;
    }

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
            if (mutatedEntries.Contains(pair.Key)) continue;
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

    private static IEnumerable<string> VerifyGenericRules(string path)
    {
        var issues = new List<string>();
        using var archive = ZipFile.OpenRead(path);
        var storyParts = GetStoryPartNames(archive);
        var documentHasPersian = false;
        var styleInfo = archive.GetEntry("word/styles.xml") is not null
            ? BuildStyleInfo(LoadXmlEntry(archive, "word/styles.xml"))
            : new Dictionary<string, PersianSemanticStyleInfo>(StringComparer.OrdinalIgnoreCase);

        foreach (var partName in storyParts)
        {
            var document = LoadXmlEntry(archive, partName);
            var allowBodyJustification = partName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase);

            var paragraphIndex = 0;
            foreach (var paragraph in document.Descendants(W + "p"))
            {
                paragraphIndex++;
                var text = ParagraphText(paragraph);
                if (!ContainsPersianScript(text)) continue;
                documentHasPersian = true;

                var pPr = paragraph.Element(W + "pPr");
                if (pPr is null || !IsOn(pPr.Element(W + "bidi")))
                    issues.Add($"{partName}: Persian paragraph is missing RTL paragraph direction.");

                var alignment = ((string?)pPr?.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();
                if (alignment.Equals("left", StringComparison.OrdinalIgnoreCase) ||
                    alignment.Equals("right", StringComparison.OrdinalIgnoreCase) ||
                    alignment.Equals("end", StringComparison.OrdinalIgnoreCase) ||
                    string.IsNullOrWhiteSpace(alignment))
                {
                    issues.Add($"{partName}: Persian RTL paragraph still has physical/ambiguous alignment instead of direction-aware start/center/both.");
                }

                var semantic = ResolveSemanticDecision(
                    paragraph,
                    text,
                    styleInfo,
                    paragraphIndex,
                    allowBodyJustification);
                var shouldBeCenteredSemantic = semantic.CenteredSemantic;
                var shouldJustify = allowBodyJustification && semantic.ShouldJustifyBody;
                var shouldBeHeadingRight = semantic.VisualRightSemantic;

                if (shouldJustify &&
                    !alignment.Equals("both", StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add($"{partName}: Persian body prose is not fully justified.");
                }

                if (shouldBeCenteredSemantic &&
                    !alignment.Equals("center", StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add($"{partName}: Persian document/TOC title or protected chapter role is not centered.");
                }

                if (shouldBeHeadingRight &&
                    !alignment.Equals("start", StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add($"{partName}: Persian RTL in-document heading is not direction-aware visual-right aligned (w:jc=start).");
                }

                foreach (var run in paragraph.Descendants(W + "r"))
                {
                    var runText = string.Concat(run.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));
                    if (!ContainsPersianScript(runText)) continue;
                    if (!IsOn(run.Element(W + "rPr")?.Element(W + "rtl")))
                        issues.Add($"{partName}: Persian run is missing w:rtl.");
                }
            }

            foreach (var table in document.Descendants(W + "tbl"))
            {
                var tblPr = table.Element(W + "tblPr");
                var alignment = ((string?)tblPr?.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();
                if (!alignment.Equals("center", StringComparison.OrdinalIgnoreCase))
                    issues.Add($"{partName}: table is not centered.");

                var tableText = string.Concat(table.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));
                if (ContainsPersianScript(tableText) && !IsOn(tblPr?.Element(W + "bidiVisual")))
                    issues.Add($"{partName}: Persian table is missing RTL visual order.");
            }
        }

        if (documentHasPersian && archive.GetEntry("word/styles.xml") is not null)
        {
            var styles = LoadXmlEntry(archive, "word/styles.xml");
            foreach (var style in styles.Descendants(W + "style"))
            {
                var id = (string?)style.Attribute(W + "styleId") ?? string.Empty;
                var name = (string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty;
                if (!IsTocStyle(id, name)) continue;

                var pPr = style.Element(W + "pPr");
                if (pPr is null || !IsOn(pPr.Element(W + "bidi")))
                    issues.Add($"TOC style '{id}' is missing RTL paragraph direction.");

                var alignment = ((string?)pPr?.Element(W + "jc")?.Attribute(W + "val") ?? string.Empty).Trim();
                if (!alignment.Equals("start", StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add($"TOC style '{id}' is not direction-aware RTL leading-edge aligned (w:jc=start).");
                }
            }
        }

        return issues;
    }

    private static bool IsOn(XElement? element)
    {
        if (element is null) return false;
        var value = ((string?)element.Attribute(W + "val") ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(value)) return true;
        return value.Equals("1", StringComparison.OrdinalIgnoreCase)
            || value.Equals("true", StringComparison.OrdinalIgnoreCase)
            || value.Equals("on", StringComparison.OrdinalIgnoreCase);
    }

    private static string ParagraphText(XElement paragraph)
        => string.Concat(paragraph.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));

    private static bool ContainsPersianScript(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        foreach (var c in value)
        {
            if ((c >= '\u0600' && c <= '\u06FF') ||
                (c >= '\u0750' && c <= '\u077F') ||
                (c >= '\u08A0' && c <= '\u08FF') ||
                (c >= '\uFB50' && c <= '\uFDFF') ||
                (c >= '\uFE70' && c <= '\uFEFF'))
                return true;
        }
        return false;
    }

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

    private static string Sha256File(string path)
    {
        using var stream = File.OpenRead(path);
        return Sha256Stream(stream);
    }

    private static string Sha256Stream(Stream stream)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream));
    }

    private static string Sha256Text(string value)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(value)));
    }


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
            if (File.Exists(path)) File.Delete(path);
        }
        catch
        {
            // Best-effort cleanup only.
        }
    }
}
