using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using WordAgent.Core;

namespace WordAgent.WordInterop;

public sealed class WordAuditService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const int PersistedFastPathParagraphThreshold = 1000;

    public WordAuditSnapshot Inspect(string inputPath, bool visible = false)
    {
        var path = PathPolicy.RequireSupportedWordFile(inputPath);

        // R10R41: Word native compatibility/rendering remains a separate gate.
        // Large read-only audits no longer enumerate thousands of paragraphs through
        // cross-process Word COM. Persisted OOXML is the deterministic audit authority.
        if (ShouldUsePersistedFastPath(path))
            return InspectPersistedFast(path);

        var warnings = new List<string>();
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

            var counts = CollectCounts(doc, warnings);
            var formatting = CollectFormatting(doc, warnings);
            return new WordAuditSnapshot(path, app.Version, counts, formatting, warnings);
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

    private static bool ShouldUsePersistedFastPath(string path)
    {
        try
        {
            using var zip = ZipFile.OpenRead(path);
            var entry = zip.GetEntry("word/document.xml");
            if (entry is null) return false;
            using var stream = entry.Open();
            var document = XDocument.Load(stream, LoadOptions.None);
            return document.Descendants(W + "p").Take(PersistedFastPathParagraphThreshold + 1).Count()
                > PersistedFastPathParagraphThreshold;
        }
        catch
        {
            // Preserve the already-accepted legacy behavior when the fast-path
            // admission cannot be proven from a valid DOCX package.
            return false;
        }
    }

    private static WordAuditSnapshot InspectPersistedFast(string path)
    {
        var warnings = new List<string>
        {
            "WORD_AUDIT_AUTHORITY=PERSISTED_OOXML_FAST_PATH; WORD_NATIVE_COMPATIBILITY_IS_VERIFIED_SEPARATELY."
        };

        var probe = WordRuntime.Probe();
        var wordVersion = probe.Version ?? "UNKNOWN";

        using var zip = ZipFile.OpenRead(path);
        var document = LoadXml(zip, "word/document.xml");
        var styles = TryLoadXml(zip, "word/styles.xml");

        var counts = CollectPersistedCounts(zip, document, styles);
        var formatting = CollectPersistedFormatting(document, styles);

        return new WordAuditSnapshot(path, wordVersion, counts, formatting, warnings);
    }

    private static InventoryCounts CollectPersistedCounts(
        ZipArchive zip,
        XDocument document,
        XDocument? styles)
    {
        var paragraphs = document.Descendants(W + "p").Count()
            + document.Descendants(W + "tr").Count(); // Word main-story row-end paragraph markers.
        var tables = document.Descendants(W + "tbl").Count();
        var inlineShapes = document.Descendants().Count(x => x.Name.LocalName == "inline");
        var anchors = document.Descendants().Count(x => x.Name.LocalName == "anchor");
        var legacyPicts = document.Descendants(W + "pict").Count();
        var floatingShapes = anchors + legacyPicts;

        var instructions = EnumerateFieldInstructions(document).ToArray();
        var fields = instructions.Length;
        var endNoteFields = instructions.Count(IsEndNoteInstruction);
        var tocFields = instructions.Count(IsTocInstruction);

        var sections = Math.Max(1, document.Descendants(W + "sectPr").Count());
        var styleCount = styles?.Descendants(W + "style").Count() ?? 0;

        var footnotes = CountNumberedPartItems(zip, "word/footnotes.xml", "footnote");
        var endnotes = CountNumberedPartItems(zip, "word/endnotes.xml", "endnote");
        var comments = TryLoadXml(zip, "word/comments.xml")?.Descendants(W + "comment").Count() ?? 0;

        var revisions =
            document.Descendants(W + "ins").Count()
            + document.Descendants(W + "del").Count()
            + document.Descendants(W + "moveFrom").Count()
            + document.Descendants(W + "moveTo").Count();

        var contentControls = document.Descendants(W + "sdt").Count();
        var bookmarks = document.Descendants(W + "bookmarkStart").Count();
        var hyperlinks = document.Descendants(W + "hyperlink").Count();
        var equations = document.Descendants().Count(x => x.Name.LocalName == "oMath");

        return new InventoryCounts(
            paragraphs,
            tables,
            inlineShapes,
            floatingShapes,
            fields,
            endNoteFields,
            tocFields,
            sections,
            styleCount,
            footnotes,
            endnotes,
            comments,
            revisions,
            contentControls,
            bookmarks,
            hyperlinks,
            equations);
    }

    private static FormattingMetrics CollectPersistedFormatting(
        XDocument document,
        XDocument? styles)
    {
        var catalog = PersistedParagraphStyleCatalog.Create(styles);
        var paragraphs = document.Descendants(W + "p").ToArray();

        var nonEmpty = 0;
        var rtl = 0;
        var right = 0;
        var center = 0;
        var inlineCentered = 0;

        foreach (var paragraph in paragraphs)
        {
            var text = string.Concat(paragraph.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));
            if (string.IsNullOrWhiteSpace(text))
                continue;

            nonEmpty++;
            var format = ResolveParagraphFormatting(paragraph, catalog);
            if (format.ReadingOrder == "RTL") rtl++;
            if (format.Alignment == "RIGHT") right++;
            if (format.Alignment == "CENTER") center++;

            if (paragraph.Descendants().Any(x => x.Name.LocalName == "inline")
                && format.Alignment == "CENTER")
                inlineCentered++;
        }

        var rtlTables = 0;
        var centeredTables = 0;
        foreach (var table in document.Descendants(W + "tbl"))
        {
            var tblPr = table.Element(W + "tblPr");
            if (ParseOnOff(tblPr?.Element(W + "bidiVisual")) is true)
                rtlTables++;

            var jc = NormalizeAlignment((string?)tblPr?.Element(W + "jc")?.Attribute(W + "val"));
            if (jc == "CENTER")
                centeredTables++;
        }

        return new FormattingMetrics(
            nonEmpty,
            rtl,
            right,
            center,
            rtlTables,
            centeredTables,
            inlineCentered);
    }

    private static PersistedParagraphFormat ResolveParagraphFormatting(
        XElement paragraph,
        PersistedParagraphStyleCatalog catalog)
    {
        var pPr = paragraph.Element(W + "pPr");
        var styleId = (string?)pPr?.Element(W + "pStyle")?.Attribute(W + "val");
        var resolved = catalog.Resolve(styleId);

        var reading = resolved.ReadingOrder;
        var alignment = resolved.Alignment;

        var directBidi = ParseOnOff(pPr?.Element(W + "bidi"));
        if (directBidi is not null)
            reading = directBidi.Value ? "RTL" : "LTR";

        var directAlignment = NormalizeAlignment(
            (string?)pPr?.Element(W + "jc")?.Attribute(W + "val"));
        if (!string.IsNullOrWhiteSpace(directAlignment))
            alignment = directAlignment!;

        return new PersistedParagraphFormat(reading, alignment);
    }

    private static IEnumerable<string> EnumerateFieldInstructions(XDocument document)
    {
        foreach (var simple in document.Descendants(W + "fldSimple"))
        {
            var instruction = (string?)simple.Attribute(W + "instr");
            if (!string.IsNullOrWhiteSpace(instruction))
                yield return instruction!;
        }

        var depth = 0;
        StringBuilder? current = null;
        foreach (var node in document.Descendants()
                     .Where(x => x.Name == W + "fldChar" || x.Name == W + "instrText"))
        {
            if (node.Name == W + "fldChar")
            {
                var type = ((string?)node.Attribute(W + "fldCharType"))?.Trim();
                if (string.Equals(type, "begin", StringComparison.OrdinalIgnoreCase))
                {
                    if (depth == 0)
                        current = new StringBuilder();
                    depth++;
                }
                else if (string.Equals(type, "end", StringComparison.OrdinalIgnoreCase) && depth > 0)
                {
                    depth--;
                    if (depth == 0 && current is not null)
                    {
                        var value = current.ToString();
                        if (!string.IsNullOrWhiteSpace(value))
                            yield return value;
                        current = null;
                    }
                }
                continue;
            }

            if (depth > 0 && current is not null)
            {
                if (current.Length > 0) current.Append(' ');
                current.Append((string?)node ?? string.Empty);
            }
        }
    }

    private static bool IsEndNoteInstruction(string instruction)
        => instruction.Contains("EN.CITE", StringComparison.OrdinalIgnoreCase)
           || instruction.Contains("EN.BIB", StringComparison.OrdinalIgnoreCase)
           || instruction.Contains("EndNote", StringComparison.OrdinalIgnoreCase);

    private static bool IsTocInstruction(string instruction)
        => instruction.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase);

    private static int CountNumberedPartItems(ZipArchive zip, string entryName, string localName)
    {
        var xml = TryLoadXml(zip, entryName);
        if (xml is null) return 0;

        return xml.Descendants()
            .Where(x => x.Name.LocalName == localName)
            .Count(x =>
            {
                var raw = (string?)x.Attribute(W + "id");
                return !int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id)
                       || id >= 0;
            });
    }

    private static XDocument LoadXml(ZipArchive zip, string entryName)
    {
        var entry = zip.GetEntry(entryName)
            ?? throw new InvalidDataException($"DOCX package entry missing: {entryName}");
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static XDocument? TryLoadXml(ZipArchive zip, string entryName)
    {
        var entry = zip.GetEntry(entryName);
        if (entry is null) return null;
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

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

    private sealed record PersistedParagraphFormat(string ReadingOrder, string Alignment);

    private sealed class PersistedParagraphStyleCatalog
    {
        private readonly Dictionary<string, PersistedStyleNode> _styles;
        private readonly PersistedParagraphFormat _defaults;
        private readonly Dictionary<string, PersistedParagraphFormat> _cache =
            new(StringComparer.OrdinalIgnoreCase);

        private PersistedParagraphStyleCatalog(
            Dictionary<string, PersistedStyleNode> styles,
            PersistedParagraphFormat defaults)
        {
            _styles = styles;
            _defaults = defaults;
        }

        public static PersistedParagraphStyleCatalog Create(XDocument? styles)
        {
            var map = new Dictionary<string, PersistedStyleNode>(StringComparer.OrdinalIgnoreCase);
            var defaults = new PersistedParagraphFormat("UNKNOWN", "UNKNOWN");

            var root = styles?.Root;
            if (root is null)
                return new PersistedParagraphStyleCatalog(map, defaults);

            var docDefaults = root.Element(W + "docDefaults");
            defaults = Apply(
                defaults,
                docDefaults?.Element(W + "pPrDefault")?.Element(W + "pPr"));

            foreach (var style in root.Elements(W + "style"))
            {
                var id = (string?)style.Attribute(W + "styleId");
                if (string.IsNullOrWhiteSpace(id)) continue;
                map[id] = new PersistedStyleNode(
                    id,
                    (string?)style.Element(W + "basedOn")?.Attribute(W + "val"),
                    style);
            }

            return new PersistedParagraphStyleCatalog(map, defaults);
        }

        public PersistedParagraphFormat Resolve(string? styleId)
        {
            if (string.IsNullOrWhiteSpace(styleId))
                return _defaults;

            if (_cache.TryGetValue(styleId, out var cached))
                return cached;

            var result = ResolveCore(styleId!, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            _cache[styleId!] = result;
            return result;
        }

        private PersistedParagraphFormat ResolveCore(string styleId, HashSet<string> visiting)
        {
            if (_cache.TryGetValue(styleId, out var cached))
                return cached;
            if (!_styles.TryGetValue(styleId, out var node))
                return _defaults;
            if (!visiting.Add(styleId))
                return _defaults;

            var current = !string.IsNullOrWhiteSpace(node.BasedOn)
                ? ResolveCore(node.BasedOn!, visiting)
                : _defaults;

            current = Apply(current, node.Element.Element(W + "pPr"));
            visiting.Remove(styleId);
            _cache[styleId] = current;
            return current;
        }

        private static PersistedParagraphFormat Apply(
            PersistedParagraphFormat inherited,
            XElement? pPr)
        {
            if (pPr is null) return inherited;

            var reading = inherited.ReadingOrder;
            var alignment = inherited.Alignment;

            var bidi = ParseOnOff(pPr.Element(W + "bidi"));
            if (bidi is not null)
                reading = bidi.Value ? "RTL" : "LTR";

            var jc = NormalizeAlignment((string?)pPr.Element(W + "jc")?.Attribute(W + "val"));
            if (!string.IsNullOrWhiteSpace(jc))
                alignment = jc!;

            return new PersistedParagraphFormat(reading, alignment);
        }

        private sealed record PersistedStyleNode(string Id, string? BasedOn, XElement Element);
    }

    private static InventoryCounts CollectCounts(Word.Document doc, List<string> warnings)
    {
        int endNoteFields = 0;
        int tocFields = 0;
        Word.Fields? fields = null;
        try
        {
            fields = doc.Fields;
            for (int i = 1; i <= fields.Count; i++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                try
                {
                    field = fields[i];
                    code = field.Code;
                    var text = code.Text ?? string.Empty;
                    if (text.Contains("EN.CITE", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("EN.BIB", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("EndNote", StringComparison.OrdinalIgnoreCase))
                        endNoteFields++;
                    if (text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
                        tocFields++;
                }
                catch (Exception ex) { warnings.Add($"Field {i}: {ex.Message}"); }
                finally { ComUtil.Release(code); ComUtil.Release(field); }
            }
        }
        finally { ComUtil.Release(fields); }

        return new InventoryCounts(
            SafeCount(() => doc.Paragraphs.Count),
            SafeCount(() => doc.Tables.Count),
            SafeCount(() => doc.InlineShapes.Count),
            SafeCount(() => doc.Shapes.Count),
            SafeCount(() => doc.Fields.Count),
            endNoteFields,
            tocFields,
            SafeCount(() => doc.Sections.Count),
            SafeCount(() => doc.Styles.Count),
            SafeCount(() => doc.Footnotes.Count),
            SafeCount(() => doc.Endnotes.Count),
            SafeCount(() => doc.Comments.Count),
            SafeCount(() => doc.Revisions.Count),
            SafeCount(() => doc.ContentControls.Count),
            SafeCount(() => doc.Bookmarks.Count),
            SafeCount(() => doc.Hyperlinks.Count),
            SafeCount(() => doc.OMaths.Count));
    }

    private static FormattingMetrics CollectFormatting(Word.Document doc, List<string> warnings)
    {
        int nonEmpty = 0, rtl = 0, right = 0, center = 0;
        Word.Paragraphs? paragraphs = null;
        try
        {
            paragraphs = doc.Paragraphs;
            var count = paragraphs.Count;
            for (int i = 1; i <= count; i++)
            {
                Word.Paragraph? p = null;
                Word.Range? range = null;
                Word.ParagraphFormat? format = null;
                try
                {
                    p = paragraphs[i];
                    range = p.Range;
                    var text = CleanParagraphText(range.Text);
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    nonEmpty++;
                    format = range.ParagraphFormat;
                    if (format.ReadingOrder == Word.WdReadingOrder.wdReadingOrderRtl) rtl++;
                    if (format.Alignment == Word.WdParagraphAlignment.wdAlignParagraphRight) right++;
                    if (format.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter) center++;
                }
                catch (Exception ex)
                {
                    if (warnings.Count < 50) warnings.Add($"Paragraph {i}: {ex.Message}");
                }
                finally { ComUtil.Release(format); ComUtil.Release(range); ComUtil.Release(p); }
            }
        }
        finally { ComUtil.Release(paragraphs); }

        int rtlTables = 0, centeredTables = 0;
        Word.Tables? tables = null;
        try
        {
            tables = doc.Tables;
            var count = tables.Count;
            for (int i = 1; i <= count; i++)
            {
                Word.Table? table = null;
                Word.Rows? rows = null;
                try
                {
                    table = tables[i];
                    if (table.TableDirection == Word.WdTableDirection.wdTableDirectionRtl) rtlTables++;
                    rows = table.Rows;
                    if (rows.Alignment == Word.WdRowAlignment.wdAlignRowCenter) centeredTables++;
                }
                catch (Exception ex)
                {
                    if (warnings.Count < 50) warnings.Add($"Table {i}: {ex.Message}");
                }
                finally { ComUtil.Release(rows); ComUtil.Release(table); }
            }
        }
        finally { ComUtil.Release(tables); }

        int inlineCentered = 0;
        Word.InlineShapes? inlineShapes = null;
        try
        {
            inlineShapes = doc.InlineShapes;
            var count = inlineShapes.Count;
            for (int i = 1; i <= count; i++)
            {
                Word.InlineShape? shape = null;
                Word.Range? range = null;
                Word.ParagraphFormat? format = null;
                try
                {
                    shape = inlineShapes[i];
                    range = shape.Range;
                    format = range.ParagraphFormat;
                    if (format.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter) inlineCentered++;
                }
                catch { }
                finally { ComUtil.Release(format); ComUtil.Release(range); ComUtil.Release(shape); }
            }
        }
        finally { ComUtil.Release(inlineShapes); }

        return new FormattingMetrics(nonEmpty, rtl, right, center, rtlTables, centeredTables, inlineCentered);
    }

    private static int SafeCount(Func<int> getter)
    {
        try { return getter(); }
        catch { return -1; }
    }

    internal static string CleanParagraphText(string? text)
        => (text ?? string.Empty).Replace("\r", string.Empty).Replace("\a", string.Empty).Replace("\u0001", string.Empty).Trim();
}
