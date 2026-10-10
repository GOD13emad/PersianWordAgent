using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;

namespace WordAgent.WordInterop;

public sealed record TemplateSectionContract(
    int Index,
    int? PageWidthTwips,
    int? PageHeightTwips,
    string? Orientation,
    int? MarginTopTwips,
    int? MarginRightTwips,
    int? MarginBottomTwips,
    int? MarginLeftTwips,
    int? HeaderDistanceTwips,
    int? FooterDistanceTwips,
    int? GutterTwips,
    int? ColumnCount,
    int? ColumnSpaceTwips,
    bool? EqualWidthColumns);

public sealed record TemplateStyleContract(
    string StyleId,
    string Type,
    string? Name,
    bool IsDefault,
    bool IsQuickFormat,
    string? BasedOn,
    string? Next,
    string? Alignment,
    bool Bidi,
    string? FontAscii,
    string? FontHighAnsi,
    string? FontComplexScript,
    string? FontEastAsia,
    string? SizeHalfPoints,
    string? ComplexScriptSizeHalfPoints,
    bool Bold,
    bool Italic,
    bool RtlRun,
    string? SpaceBeforeTwips,
    string? SpaceAfterTwips,
    string? LineSpacingTwips);

public sealed record TemplateContractResult(
    string SchemaVersion,
    string TemplateFile,
    string TemplateSha256,
    IReadOnlyList<TemplateSectionContract> Sections,
    IReadOnlyList<TemplateStyleContract> Styles,
    int TableStyleCount,
    int HeaderPartCount,
    int FooterPartCount,
    int AbstractNumberingCount,
    int NumberingInstanceCount,
    int TocFieldCount);

public sealed class TemplateContractService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public TemplateContractResult Analyze(string templatePath)
    {
        if (string.IsNullOrWhiteSpace(templatePath))
            throw new ArgumentException("Template path is required.", nameof(templatePath));

        var fullPath = Path.GetFullPath(templatePath);
        if (!File.Exists(fullPath))
            throw new FileNotFoundException("Template DOCX was not found.", fullPath);

        var beforeSha = Sha256(fullPath);

        using var archive = ZipFile.OpenRead(fullPath);

        var document = ReadXml(archive, "word/document.xml")
            ?? throw new InvalidOperationException("Template is missing word/document.xml.");

        var styles = ReadXml(archive, "word/styles.xml");
        var numbering = ReadXml(archive, "word/numbering.xml");

        var sections = document
            .Descendants(W + "sectPr")
            .Select((sectPr, index) => ParseSection(sectPr, index + 1))
            .ToArray();

        var styleContracts = styles is null
            ? Array.Empty<TemplateStyleContract>()
            : styles
                .Descendants(W + "style")
                .Select(ParseStyle)
                .OrderBy(x => x.Type, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.StyleId, StringComparer.OrdinalIgnoreCase)
                .ToArray();

        var tableStyleCount = styleContracts.Count(x =>
            x.Type.Equals("table", StringComparison.OrdinalIgnoreCase));

        var headerPartCount = archive.Entries.Count(x =>
            x.FullName.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) &&
            x.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

        var footerPartCount = archive.Entries.Count(x =>
            x.FullName.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) &&
            x.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

        var abstractNumberingCount = numbering?.Descendants(W + "abstractNum").Count() ?? 0;
        var numberingInstanceCount = numbering?.Descendants(W + "num").Count() ?? 0;

        var tocFieldCount = document
            .Descendants(W + "instrText")
            .Select(x => x.Value)
            .Count(x => x.Contains("TOC", StringComparison.OrdinalIgnoreCase));

        var afterSha = Sha256(fullPath);
        if (!beforeSha.Equals(afterSha, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Template Analyzer modified the template unexpectedly.");

        return new TemplateContractResult(
            SchemaVersion: "wordagent.template.contract/1.0",
            TemplateFile: fullPath,
            TemplateSha256: beforeSha,
            Sections: sections,
            Styles: styleContracts,
            TableStyleCount: tableStyleCount,
            HeaderPartCount: headerPartCount,
            FooterPartCount: footerPartCount,
            AbstractNumberingCount: abstractNumberingCount,
            NumberingInstanceCount: numberingInstanceCount,
            TocFieldCount: tocFieldCount);
    }

    private static TemplateSectionContract ParseSection(XElement sectPr, int index)
    {
        var pageSize = sectPr.Element(W + "pgSz");
        var margins = sectPr.Element(W + "pgMar");
        var columns = sectPr.Element(W + "cols");

        return new TemplateSectionContract(
            Index: index,
            PageWidthTwips: IntAttr(pageSize, "w"),
            PageHeightTwips: IntAttr(pageSize, "h"),
            Orientation: Attr(pageSize, "orient"),
            MarginTopTwips: IntAttr(margins, "top"),
            MarginRightTwips: IntAttr(margins, "right"),
            MarginBottomTwips: IntAttr(margins, "bottom"),
            MarginLeftTwips: IntAttr(margins, "left"),
            HeaderDistanceTwips: IntAttr(margins, "header"),
            FooterDistanceTwips: IntAttr(margins, "footer"),
            GutterTwips: IntAttr(margins, "gutter"),
            ColumnCount: IntAttr(columns, "num"),
            ColumnSpaceTwips: IntAttr(columns, "space"),
            EqualWidthColumns: BoolAttr(columns, "equalWidth"));
    }

    private static TemplateStyleContract ParseStyle(XElement style)
    {
        var paragraphProperties = style.Element(W + "pPr");
        var runProperties = style.Element(W + "rPr");
        var fonts = runProperties?.Element(W + "rFonts");
        var spacing = paragraphProperties?.Element(W + "spacing");

        return new TemplateStyleContract(
            StyleId: Attr(style, "styleId") ?? string.Empty,
            Type: Attr(style, "type") ?? string.Empty,
            Name: Attr(style.Element(W + "name"), "val"),
            IsDefault: BoolAttr(style, "default") == true,
            IsQuickFormat: style.Element(W + "qFormat") is not null,
            BasedOn: Attr(style.Element(W + "basedOn"), "val"),
            Next: Attr(style.Element(W + "next"), "val"),
            Alignment: Attr(paragraphProperties?.Element(W + "jc"), "val"),
            Bidi: IsOn(paragraphProperties?.Element(W + "bidi")),
            FontAscii: Attr(fonts, "ascii"),
            FontHighAnsi: Attr(fonts, "hAnsi"),
            FontComplexScript: Attr(fonts, "cs"),
            FontEastAsia: Attr(fonts, "eastAsia"),
            SizeHalfPoints: Attr(runProperties?.Element(W + "sz"), "val"),
            ComplexScriptSizeHalfPoints: Attr(runProperties?.Element(W + "szCs"), "val"),
            Bold: IsOn(runProperties?.Element(W + "b")),
            Italic: IsOn(runProperties?.Element(W + "i")),
            RtlRun: IsOn(runProperties?.Element(W + "rtl")),
            SpaceBeforeTwips: Attr(spacing, "before"),
            SpaceAfterTwips: Attr(spacing, "after"),
            LineSpacingTwips: Attr(spacing, "line"));
    }

    private static XDocument? ReadXml(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName);
        if (entry is null)
            return null;

        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static string? Attr(XElement? element, string localName) =>
        element?.Attribute(W + localName)?.Value;

    private static int? IntAttr(XElement? element, string localName)
    {
        var value = Attr(element, localName);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result)
            ? result
            : null;
    }

    private static bool? BoolAttr(XElement? element, string localName)
    {
        var value = Attr(element, localName);
        if (value is null)
            return null;

        return value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("on", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsOn(XElement? element)
    {
        if (element is null)
            return false;

        var value = Attr(element, "val");
        if (value is null)
            return true;

        return !value.Equals("0", StringComparison.OrdinalIgnoreCase) &&
               !value.Equals("false", StringComparison.OrdinalIgnoreCase) &&
               !value.Equals("off", StringComparison.OrdinalIgnoreCase);
    }

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
