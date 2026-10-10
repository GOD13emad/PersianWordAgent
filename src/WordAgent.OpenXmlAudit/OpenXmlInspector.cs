using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using W = DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;
using WordAgent.Core;

namespace WordAgent.OpenXmlAudit;

public sealed class OpenXmlInspector
{
    public OpenXmlAuditSnapshot Inspect(string path)
    {
        path = PathPolicy.RequireSupportedWordFile(path);
        using var doc = WordprocessingDocument.Open(path, false);
        var main = doc.MainDocumentPart;
        var root = main?.Document;
        if (root is null)
            throw new InvalidDataException("The Word package has no main document part.");

        var paragraphs = root.Descendants<W.Paragraph>().Count();
        var tables = root.Descendants<W.Table>().Count();
        var drawings = root.Descendants<W.Drawing>().Count();
        var styles = main?.StyleDefinitionsPart?.Styles?.Elements<W.Style>().Count() ?? 0;

        var instructionTexts = EnumerateAllRootElements(doc)
            .SelectMany(e => e.Descendants<W.FieldCode>())
            .Select(x => x.Text ?? string.Empty)
            .ToList();

        var fieldInstructions = instructionTexts.Count;
        var endNoteMarkers = instructionTexts.Count(x =>
            x.Contains("EN.CITE", StringComparison.OrdinalIgnoreCase) ||
            x.Contains("EN.BIB", StringComparison.OrdinalIgnoreCase) ||
            x.Contains("EndNote", StringComparison.OrdinalIgnoreCase));

        var validator = new OpenXmlValidator(FileFormatVersions.Office2016);
        var errors = validator.Validate(doc).Take(101).ToList();
        var samples = errors.Take(10)
            .Select(e => $"{e.Description} @ {e.Path?.XPath}")
            .ToList();

        return new OpenXmlAuditSnapshot(
            paragraphs,
            tables,
            drawings,
            styles,
            fieldInstructions,
            endNoteMarkers,
            errors.Count,
            samples);
    }

    private static IEnumerable<OpenXmlElement> EnumerateAllRootElements(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document is not null)
            yield return doc.MainDocumentPart.Document;

        foreach (var h in doc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
            if (h.Header is not null) yield return h.Header;

        foreach (var f in doc.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>())
            if (f.Footer is not null) yield return f.Footer;

        if (doc.MainDocumentPart?.FootnotesPart?.Footnotes is not null)
            yield return doc.MainDocumentPart.FootnotesPart.Footnotes;

        if (doc.MainDocumentPart?.EndnotesPart?.Endnotes is not null)
            yield return doc.MainDocumentPart.EndnotesPart.Endnotes;
    }
}


