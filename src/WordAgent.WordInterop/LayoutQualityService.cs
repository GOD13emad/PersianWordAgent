using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using WordAgent.Core;

namespace WordAgent.WordInterop;

public sealed record LayoutTablePlan(
    int TableIndex,
    int Columns,
    int GridWidthTwips,
    int TargetWidthTwips,
    bool IsWide,
    bool NeedsNormalization,
    string CaptionText,
    string Action);

public sealed record LayoutQualityPlan(
    string InputFile,
    string ProfileId,
    bool TocFieldFound,
    int TocEntryParagraphs,
    int TocStylesFound,
    int TocStylesNeedingCompaction,
    bool TocTitleFound,
    bool FirstContentFound,
    string FirstContentText,
    bool FirstContentHasPageBreakBefore,
    bool BodyPageBreakNeeded,
    bool AppendixFound,
    string AppendixTitle,
    int AppendixTables,
    int WideAppendixTables,
    bool WideTailGroupSupported,
    bool LandscapeTailPresent,
    int PendingActions,
    IReadOnlyList<LayoutTablePlan> Tables,
    IReadOnlyList<string> Blockers,
    IReadOnlyList<string> Warnings)
{
    public bool Ready => Blockers.Count == 0;
}

public sealed record LayoutMutationResult(
    string InputFile,
    string OutputFile,
    string ProfileId,
    string InputSha256,
    string OutputSha256,
    LayoutQualityPlan BeforePlan,
    LayoutQualityPlan AfterPlan,
    int TocStylesChanged,
    int BodyPageBreaksChanged,
    int AppendixTablesNormalized,
    int WideTablesExpanded,
    int TablesStyled,
    int SectionBreakParagraphsInserted,
    int TrailingBlankParagraphsRemoved,
    IReadOnlyList<string> Warnings);

public sealed class LayoutQualityService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const int A4PortraitWidthTwips = 11906;
    private const int A4PortraitHeightTwips = 16838;

    public LayoutQualityPlan Inspect(string inputPath, AgentProfile profile)
    {
        var input = PathPolicy.RequireSupportedWordFile(inputPath);
        using var file = new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read, leaveOpen: false);
        var document = LoadXmlEntry(archive, "word/document.xml");
        var styles = LoadXmlEntry(archive, "word/styles.xml");
        return BuildPlan(input, profile, document, styles);
    }

    public LayoutMutationResult Apply(
        string inputPath,
        AgentProfile profile,
        string? outputDirectory = null)
    {
        var input = PathPolicy.RequireSupportedWordFile(inputPath);
        var before = Inspect(input, profile);
        if (!before.Ready)
            throw new InvalidOperationException("Layout mutation blocked: " + string.Join(" | ", before.Blockers));

        var output = PathPolicy.CreateSafeOutputPath(input, outputDirectory, "_WordAgent_Layout");
        File.Copy(input, output, overwrite: false);
        var warnings = new List<string>();
        var inputHash = Sha256(input);

        int tocChanged = 0;
        int pageBreakChanged = 0;
        int tablesChanged = 0;
        int wideExpanded = 0;
        int tablesStyled = 0;
        int sectionBreaksInserted = 0;
        int trailingBlankParagraphsRemoved = 0;

        try
        {
            using var file = new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            using var archive = new ZipArchive(file, ZipArchiveMode.Update, leaveOpen: false);

            if (archive.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Digitally signed Office package detected; layout mutation is blocked because it would invalidate the package signature.");

            var document = LoadXmlEntry(archive, "word/document.xml");
            var styles = LoadXmlEntry(archive, "word/styles.xml");
            var catalog = StyleCatalog.Build(styles);

            tocChanged = CompactTocStyles(styles, catalog, profile.LayoutPolicy.Toc);
            pageBreakChanged = EnsureBodyStartsAfterToc(document, catalog, profile, profile.LayoutPolicy.Toc) ? 1 : 0;

            var appendixResult = NormalizeAppendixTables(document, catalog, profile, profile.LayoutPolicy.AppendixTables, profile.LayoutPolicy.TablePagination, warnings);
            tablesChanged = appendixResult.TablesNormalized;
            wideExpanded = appendixResult.WideTablesExpanded;
            sectionBreaksInserted = appendixResult.SectionBreakParagraphsInserted;

            var paginationChanged = NormalizeAllTablePagination(document, profile.LayoutPolicy.TablePagination);
            if (paginationChanged > 0)
                warnings.Add($"Table pagination standard normalized on {paginationChanged} table(s): compact tables keep together; large tables split only between rows with repeated header and anti-orphan guards.");

            tablesStyled = NormalizeAllTableStyle(document, profile.LayoutPolicy.TableStyle);
            if (tablesStyled > 0)
                warnings.Add($"Table finishing theme '{NormalizeTableTheme(profile.LayoutPolicy.TableStyle.Theme)}' applied to {tablesStyled} table(s). Captions were not restyled.");

            trailingBlankParagraphsRemoved = RemoveSafeTrailingBlankParagraphs(document);
            if (trailingBlankParagraphsRemoved > 0)
                warnings.Add($"Trailing formatting-only blank paragraph(s) removed: {trailingBlankParagraphsRemoved}.");

            // Office 2016 schema: w:cantSplit, w:tblHeader and w:bidiVisual are
            // OnOffOnlyType.  Do not serialize the transitional numeric literal
            // w:val="1" here; presence means ON.  Normalize equivalent legacy
            // literals before persistence, then enforce the schema contract.
            NormalizeOnOffOnlySchemaLiterals(document);
            AssertOnOffOnlySchemaContract(document);

            ReplaceXmlEntry(archive, "word/styles.xml", styles);
            ReplaceXmlEntry(archive, "word/document.xml", document);
        }
        catch
        {
            TryDelete(output);
            throw;
        }

        var outputHash = Sha256(output);
        if (string.Equals(inputHash, outputHash, StringComparison.OrdinalIgnoreCase))
        {
            TryDelete(output);
            throw new InvalidOperationException("Layout persistence gate failed: output is byte-identical to input.");
        }

        var after = Inspect(output, profile);
        if (!after.Ready)
        {
            TryDelete(output);
            throw new InvalidOperationException("Post-layout plan is blocked: " + string.Join(" | ", after.Blockers));
        }

        if (after.PendingActions != 0)
        {
            TryDelete(output);
            throw new InvalidOperationException($"Post-layout idempotency gate failed: PendingActions={after.PendingActions}.");
        }

        warnings.Add("Layout engine: OpenXML-only mutation. Microsoft Word COM write calls=0.");
        warnings.Add("TOC field content is not regenerated or renumbered. Only TOC style spacing and the first body chapter page-break contract are changed.");
        warnings.Add("Table caption text/style is preserved. Table finishing changes table/cell/paragraph presentation only; fields, relationships, media and headers/footers are not regenerated.");

        return new LayoutMutationResult(
            input,
            output,
            profile.Id,
            inputHash,
            outputHash,
            before,
            after,
            tocChanged,
            pageBreakChanged,
            tablesChanged,
            wideExpanded,
            tablesStyled,
            sectionBreaksInserted,
            trailingBlankParagraphsRemoved,
            warnings);
    }

    private static LayoutQualityPlan BuildPlan(
        string input,
        AgentProfile profile,
        XDocument document,
        XDocument styles)
    {
        var blockers = new List<string>();
        var warnings = new List<string>();
        var catalog = StyleCatalog.Build(styles);
        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidOperationException("word/document.xml has no w:body.");

        var directChildren = body.Elements().ToList();
        var paragraphs = body.Descendants(W + "p").ToList();
        var layoutPolicy = profile.LayoutPolicy;
        var tocPolicy = layoutPolicy.Toc;
        var appendixPolicy = layoutPolicy.AppendixTables;
        var paginationPolicy = layoutPolicy.TablePagination;
        var tableStylePolicy = layoutPolicy.TableStyle;

        var tocFieldFound = false;
        var tocEntries = 0;
        var tocStylesFound = 0;
        var tocStylesNeedingCompaction = 0;
        XElement? tocTitle = null;
        XElement? firstContent = null;
        var firstContentHasPageBreak = false;
        var bodyPageBreakNeeded = false;

        if (tocPolicy.Enabled)
        {
            tocFieldFound = document.Descendants(W + "instrText")
                .Any(x => NormalizeText((string?)x).StartsWith("TOC ", StringComparison.OrdinalIgnoreCase)
                       || NormalizeText((string?)x).Equals("TOC", StringComparison.OrdinalIgnoreCase))
                || document.Descendants(W + "fldSimple")
                    .Any(x => NormalizeText((string?)x.Attribute(W + "instr")).StartsWith("TOC", StringComparison.OrdinalIgnoreCase));

            var tocStyleIds = catalog.GetTocStyleIds();
            tocEntries = paragraphs.Count(p => tocStyleIds.Contains(catalog.GetParagraphStyleId(p)));
            tocStylesFound = tocPolicy.CompactStyles ? tocStyleIds.Count : 0;
            tocStylesNeedingCompaction = tocPolicy.CompactStyles
                ? CountTocStylesNeedingCompaction(styles, tocStyleIds, tocPolicy)
                : 0;

            tocTitle = FindTocTitle(paragraphs, catalog, profile, tocPolicy);
            if (tocPolicy.StartBodyOnNewPage && tocTitle is not null)
                firstContent = FindFirstChapterAfter(paragraphs, tocTitle, catalog, profile);

            firstContentHasPageBreak = firstContent is not null &&
                ElementIsOn(firstContent.Element(W + "pPr")?.Element(W + "pageBreakBefore"));
            bodyPageBreakNeeded = tocPolicy.StartBodyOnNewPage &&
                                  firstContent is not null &&
                                  !firstContentHasPageBreak;

            if (!tocFieldFound)
                blockers.Add("Automatic TOC field was not found; layout engine will not guess a table-of-contents boundary.");
            if (tocTitle is null)
                blockers.Add("TOC title paragraph was not found.");
            if (tocPolicy.StartBodyOnNewPage && firstContent is null)
                blockers.Add("First body Chapter after the TOC was not found.");
            if (tocPolicy.CompactStyles && tocStylesFound == 0)
                blockers.Add("TOC paragraph styles were not found in styles.xml.");
        }

        var appendix = appendixPolicy.Enabled
            ? FindAppendixStart(directChildren, catalog, profile)
            : null;

        if (appendixPolicy.Enabled && appendix is null)
            warnings.Add("No appendix heading matched semanticPolicy.appendix.headingPrefixes. Appendix-table layout actions will be skipped.");

        var section = ReadFinalSectionMetrics(body);
        var portraitWidth = Math.Min(section.PageWidthTwips, section.PageHeightTwips);
        var landscapeWidth = Math.Max(section.PageWidthTwips, section.PageHeightTwips);
        if (portraitWidth <= 0) portraitWidth = A4PortraitWidthTwips;
        if (landscapeWidth <= 0) landscapeWidth = A4PortraitHeightTwips;
        var leftMargin = section.LeftMarginTwips > 0 ? section.LeftMarginTwips : 1417;
        var rightMargin = section.RightMarginTwips > 0 ? section.RightMarginTwips : 1417;
        var portraitTextWidth = Math.Max(3000, portraitWidth - leftMargin - rightMargin);
        var landscapeTextWidth = Math.Max(5000, landscapeWidth - leftMargin - rightMargin);

        var tablePlans = new List<LayoutTablePlan>();
        var wideFlags = new List<bool>();
        var appendixTables = new List<XElement>();
        int globalTableIndex = 0;
        bool inAppendix = false;

        foreach (var child in directChildren)
        {
            if (appendix is not null && ReferenceEquals(child, appendix)) inAppendix = true;
            if (child.Name != W + "tbl") continue;
            globalTableIndex++;

            var columns = GetColumnCount(child);
            var gridWidth = GetGridWidth(child);
            var caption = FindNearestCaption(directChildren, child);
            var paginationMode = paginationPolicy.Enabled
                ? GetPaginationMode(child, paginationPolicy)
                : TablePaginationMode.KeepWhole;
            var paginationNormalized = !paginationPolicy.Enabled ||
                IsTablePaginationNormalized(child, caption, paginationMode, paginationPolicy);
            var tableThemeNormalized = !tableStylePolicy.Enabled ||
                IsTableThemeNormalized(child, tableStylePolicy);

            bool wide = false;
            int target = gridWidth;
            bool appendixNormalized = true;
            string layoutAction = string.Empty;

            if (inAppendix && appendixPolicy.Enabled)
            {
                appendixTables.Add(child);
                wide = appendixPolicy.LandscapeWideTables.Enabled &&
                    LayoutPolicyDecisions.IsWideTable(columns, gridWidth, appendixPolicy.WideDetection);
                wideFlags.Add(wide);

                if (appendixPolicy.FitToTextWidth)
                    target = wide ? landscapeTextWidth : portraitTextWidth;

                appendixNormalized = IsTableNormalized(
                    child,
                    target,
                    caption,
                    appendixPolicy,
                    paginationPolicy);

                layoutAction = wide
                    ? "LANDSCAPE_TAIL + FIT_TO_TEXT_WIDTH"
                    : "PORTRAIT_FIT_TO_TEXT_WIDTH";
            }

            var paginationAction = !paginationPolicy.Enabled
                ? "PAGINATION_POLICY_DISABLED"
                : paginationMode == TablePaginationMode.KeepWhole
                    ? "KEEP_WHOLE_TABLE + NO_ROW_SPLIT + KEEP_CAPTION"
                    : "CONTROLLED_MULTI_PAGE + REPEAT_HEADER + NO_ROW_SPLIT + MIN_2_BODY_ROWS + KEEP_CAPTION";

            var themeAction = !tableStylePolicy.Enabled
                ? "TABLE_THEME_DISABLED"
                : $"TABLE_THEME_{NormalizeTableTheme(tableStylePolicy.Theme)} + CELL_ALIGNMENT_INTELLIGENCE";

            var baseAction = string.IsNullOrWhiteSpace(layoutAction)
                ? paginationAction
                : layoutAction + " + " + paginationAction;

            tablePlans.Add(new LayoutTablePlan(
                globalTableIndex,
                columns,
                gridWidth,
                target,
                wide,
                !paginationNormalized || !appendixNormalized || !tableThemeNormalized,
                caption is null ? string.Empty : Preview(ParagraphText(caption), 120),
                baseAction + " + " + themeAction));
        }

        var wideCount = wideFlags.Count(x => x);
        var tailSupported = true;
        if (wideCount > 0)
        {
            var firstWide = wideFlags.FindIndex(x => x);
            for (int i = firstWide; i < wideFlags.Count; i++)
            {
                if (!wideFlags[i])
                {
                    tailSupported = false;
                    break;
                }
            }

            if (!tailSupported)
                blockers.Add("A wide appendix table occurs before a later portrait table. P2.5 supports only a final wide-table tail group to avoid unsafe mid-document section surgery.");

            if (appendixTables.Count > 0 && !HasOnlyBlankContentAfter(directChildren, appendixTables[^1]))
            {
                tailSupported = false;
                blockers.Add("Non-empty content exists after the last wide appendix table. P2.5 will not insert a landscape tail section in this case.");
            }
        }

        var landscapeTailPresent = wideCount == 0 ||
            (IsFinalSectionLandscape(body) &&
             HasLandscapeBreakBeforeFirstWide(directChildren, appendixTables, wideFlags));

        var pending = tocStylesNeedingCompaction;
        if (bodyPageBreakNeeded) pending++;
        pending += tablePlans.Count(x => x.NeedsNormalization);
        if (wideCount > 0 && !landscapeTailPresent) pending++;

        return new LayoutQualityPlan(
            input,
            profile.Id,
            tocFieldFound,
            tocEntries,
            tocStylesFound,
            tocStylesNeedingCompaction,
            tocTitle is not null,
            firstContent is not null,
            firstContent is null ? string.Empty : Preview(ParagraphText(firstContent), 120),
            firstContentHasPageBreak,
            bodyPageBreakNeeded,
            appendix is not null,
            appendix is null ? string.Empty : Preview(ParagraphText(appendix), 160),
            appendixTables.Count,
            wideCount,
            tailSupported,
            landscapeTailPresent,
            pending,
            tablePlans,
            blockers,
            warnings);
    }
    private static int CompactTocStyles(
        XDocument styles,
        StyleCatalog catalog,
        ProfileTocLayoutPolicy policy)
    {
        if (!policy.Enabled || !policy.CompactStyles) return 0;

        int changed = 0;
        foreach (var styleId in catalog.GetTocStyleIds())
        {
            var style = styles.Descendants(W + "style")
                .FirstOrDefault(x => TokensEqual((string?)x.Attribute(W + "styleId"), styleId));
            if (style is null) continue;

            var pPr = style.Element(W + "pPr");
            if (pPr is null)
            {
                pPr = new XElement(W + "pPr");
                var rPr = style.Element(W + "rPr");
                if (rPr is null) style.Add(pPr);
                else rPr.AddBeforeSelf(pPr);
            }

            var spacing = pPr.Element(W + "spacing");
            if (spacing is null)
            {
                spacing = new XElement(W + "spacing");
                InsertParagraphPropertyInSchemaOrder(pPr, spacing);
            }

            var expectedBefore = policy.Spacing.BeforeTwips.ToString();
            var expectedAfter = policy.Spacing.AfterTwips.ToString();
            var expectedLine = policy.Spacing.LineTwips.ToString();
            var expectedRule = policy.Spacing.LineRule;

            var before = (string?)spacing.Attribute(W + "before") ?? expectedBefore;
            var after = (string?)spacing.Attribute(W + "after") ?? expectedAfter;
            var line = (string?)spacing.Attribute(W + "line") ?? expectedLine;
            var rule = (string?)spacing.Attribute(W + "lineRule") ?? expectedRule;

            if (before != expectedBefore ||
                after != expectedAfter ||
                line != expectedLine ||
                !rule.Equals(expectedRule, StringComparison.OrdinalIgnoreCase))
            {
                changed++;
            }

            spacing.SetAttributeValue(W + "before", expectedBefore);
            spacing.SetAttributeValue(W + "after", expectedAfter);
            spacing.SetAttributeValue(W + "line", expectedLine);
            spacing.SetAttributeValue(W + "lineRule", expectedRule);
        }

        return changed;
    }

    private static bool EnsureBodyStartsAfterToc(
        XDocument document,
        StyleCatalog catalog,
        AgentProfile profile,
        ProfileTocLayoutPolicy policy)
    {
        if (!policy.Enabled || !policy.StartBodyOnNewPage) return false;

        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidOperationException("word/document.xml has no w:body.");
        var paragraphs = body.Elements(W + "p").ToList();
        var tocTitle = FindTocTitle(paragraphs, catalog, profile, policy)
            ?? throw new InvalidOperationException("TOC title paragraph was not found during mutation.");
        var first = FindFirstChapterAfter(paragraphs, tocTitle, catalog, profile)
            ?? throw new InvalidOperationException("First body Chapter after the TOC was not found during mutation.");

        var pPr = EnsurePPr(first);
        var pageBreak = pPr.Element(W + "pageBreakBefore");
        if (pageBreak is not null && ElementIsOn(pageBreak)) return false;

        if (pageBreak is null)
        {
            pageBreak = new XElement(W + "pageBreakBefore");
            InsertParagraphPropertyInSchemaOrder(pPr, pageBreak);
        }

        pageBreak.SetAttributeValue(W + "val", "1");
        return true;
    }
    private sealed record AppendixMutationResult(int TablesNormalized, int WideTablesExpanded, int SectionBreakParagraphsInserted);

    private static AppendixMutationResult NormalizeAppendixTables(
        XDocument document,
        StyleCatalog catalog,
        AgentProfile profile,
        ProfileAppendixTableLayoutPolicy layoutPolicy,
        ProfileTablePaginationLayoutPolicy paginationPolicy,
        List<string> warnings)
    {
        if (!layoutPolicy.Enabled)
            return new AppendixMutationResult(0, 0, 0);

        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidOperationException("word/document.xml has no w:body.");
        var children = body.Elements().ToList();
        var appendix = FindAppendixStart(children, catalog, profile);
        if (appendix is null) return new AppendixMutationResult(0, 0, 0);

        var section = ReadFinalSectionMetrics(body);
        var portraitWidth = Math.Min(section.PageWidthTwips, section.PageHeightTwips);
        var landscapeWidth = Math.Max(section.PageWidthTwips, section.PageHeightTwips);
        if (portraitWidth <= 0) portraitWidth = A4PortraitWidthTwips;
        if (landscapeWidth <= 0) landscapeWidth = A4PortraitHeightTwips;
        var leftMargin = section.LeftMarginTwips > 0 ? section.LeftMarginTwips : 1417;
        var rightMargin = section.RightMarginTwips > 0 ? section.RightMarginTwips : 1417;
        var portraitTextWidth = Math.Max(3000, portraitWidth - leftMargin - rightMargin);
        var landscapeTextWidth = Math.Max(5000, landscapeWidth - leftMargin - rightMargin);

        bool inAppendix = false;
        var appendixTables = new List<(XElement Table, XElement? Caption, bool Wide, int Columns, int GridWidth)>();

        foreach (var child in children)
        {
            if (ReferenceEquals(child, appendix)) inAppendix = true;
            if (child.Name != W + "tbl" || !inAppendix) continue;

            var columns = GetColumnCount(child);
            var gridWidth = GetGridWidth(child);
            var wide = layoutPolicy.LandscapeWideTables.Enabled &&
                LayoutPolicyDecisions.IsWideTable(columns, gridWidth, layoutPolicy.WideDetection);

            appendixTables.Add((
                child,
                FindNearestCaption(children, child),
                wide,
                columns,
                gridWidth));
        }

        if (appendixTables.Count == 0)
            return new AppendixMutationResult(0, 0, 0);

        var firstWide = appendixTables.FindIndex(x => x.Wide);
        if (firstWide >= 0)
        {
            for (int i = firstWide; i < appendixTables.Count; i++)
            {
                if (!appendixTables[i].Wide)
                    throw new InvalidOperationException("Wide appendix table is not part of a final wide-table tail group; layout mutation blocked.");
            }

            if (!HasOnlyBlankContentAfter(children, appendixTables[^1].Table))
                throw new InvalidOperationException("Non-empty content exists after the final wide appendix table; layout mutation blocked.");
        }

        int normalized = 0;
        int expanded = 0;

        foreach (var item in appendixTables)
        {
            var targetWidth = item.GridWidth;
            if (layoutPolicy.FitToTextWidth)
                targetWidth = item.Wide ? landscapeTextWidth : portraitTextWidth;

            if (!IsTableNormalized(item.Table, targetWidth, item.Caption, layoutPolicy, paginationPolicy))
                normalized++;

            NormalizeTable(item.Table, targetWidth, item.Caption, layoutPolicy, paginationPolicy);

            if (layoutPolicy.FitToTextWidth)
                ExpandTableGrid(item.Table, targetWidth);

            if (item.Wide) expanded++;
        }

        int sectionBreaks = 0;
        if (firstWide >= 0)
        {
            var currentChildren = body.Elements().ToList();
            var currentTables = appendixTables.Select(x => x.Table).ToList();
            var currentWideFlags = appendixTables.Select(x => x.Wide).ToList();
            var hasBreak = HasLandscapeBreakBeforeFirstWide(currentChildren, currentTables, currentWideFlags);
            var isLandscape = IsFinalSectionLandscape(body);

            if (!hasBreak || !isLandscape)
            {
                var firstWideItem = appendixTables[firstWide];
                var caption = firstWideItem.Caption
                    ?? throw new InvalidOperationException("Wide appendix table caption was not found; landscape section cannot be created safely.");

                InsertLandscapeTailSection(body, caption, section, insertBreak: !hasBreak);
                sectionBreaks = hasBreak ? 0 : 1;
                warnings.Add("Materialized a final A4 landscape appendix-table tail with continuous page numbering.");
            }
            else
            {
                warnings.Add("Final appendix tail is already landscape with a dedicated section break; no additional section break was inserted.");
            }
        }

        return new AppendixMutationResult(normalized, expanded, sectionBreaks);
    }
    private enum TablePaginationMode
    {
        KeepWhole,
        ControlledMultiPage
    }

    private static TablePaginationMode GetPaginationMode(
        XElement table,
        ProfileTablePaginationLayoutPolicy policy)
    {
        var rows = table.Elements(W + "tr").ToList();
        var rowCount = rows.Count;
        var columns = Math.Max(1, GetColumnCount(table));
        var totalChars = rows.Sum(r =>
            NormalizeText(string.Concat(r.Descendants(W + "t").Select(t => (string?)t ?? string.Empty))).Length);

        return LayoutPolicyDecisions.ShouldKeepTableWhole(
            rowCount,
            columns,
            totalChars,
            policy)
            ? TablePaginationMode.KeepWhole
            : TablePaginationMode.ControlledMultiPage;
    }

    private static int NormalizeAllTablePagination(
        XDocument document,
        ProfileTablePaginationLayoutPolicy policy)
    {
        if (!policy.Enabled) return 0;

        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidOperationException("word/document.xml has no w:body.");
        var directChildren = body.Elements().ToList();
        int changed = 0;

        foreach (var table in directChildren.Where(x => x.Name == W + "tbl"))
        {
            var caption = FindNearestCaption(directChildren, table);
            var mode = GetPaginationMode(table, policy);

            if (!IsTablePaginationNormalized(table, caption, mode, policy))
                changed++;

            ApplyTablePaginationStandard(table, caption, mode, policy);
        }

        return changed;
    }

    private static void ApplyTablePaginationStandard(
        XElement table,
        XElement? caption,
        TablePaginationMode mode,
        ProfileTablePaginationLayoutPolicy policy)
    {
        var rows = table.Elements(W + "tr").ToList();
        if (rows.Count == 0) return;

        if (policy.NoRowSplit)
        {
            foreach (var row in rows)
            {
                var trPr = EnsureTrPr(row);
                var cantSplit = trPr.Element(W + "cantSplit");
                if (cantSplit is null)
                {
                    cantSplit = new XElement(W + "cantSplit");
                    InsertRowPropertyInSchemaOrder(trPr, cantSplit);
                }
                SetOnOffOnlyEnabled(cantSplit);
            }
        }

        if (policy.RepeatHeader)
        {
            var firstTrPr = EnsureTrPr(rows[0]);
            var header = firstTrPr.Element(W + "tblHeader");
            if (header is null)
            {
                header = new XElement(W + "tblHeader");
                InsertRowPropertyInSchemaOrder(firstTrPr, header);
            }
            SetOnOffOnlyEnabled(header);
        }

        if (policy.KeepCaptionWithTable && caption is not null)
            EnsureKeepNext(caption);

        if (mode == TablePaginationMode.KeepWhole)
        {
            for (int i = 0; i < rows.Count - 1; i++)
                EnsureKeepNextOnRowParagraphs(rows[i]);
            return;
        }

        if (policy.MinBodyRowsAtStart > 0)
            EnsureKeepNextOnRowParagraphs(rows[0]);

        var additionalStartLinks = Math.Min(
            Math.Max(0, policy.MinBodyRowsAtStart - 1),
            Math.Max(0, rows.Count - 2));

        for (int i = 1; i <= additionalStartLinks; i++)
            EnsureKeepNextOnRowParagraphs(rows[i]);

        if (rows.Count >= 3 && policy.MinBodyRowsAtEnd >= 2)
        {
            var endLinks = Math.Min(policy.MinBodyRowsAtEnd - 1, rows.Count - 2);
            for (int offset = 2; offset <= endLinks + 1; offset++)
                EnsureKeepNextOnRowParagraphs(rows[^offset]);
        }
    }

    private static bool IsTablePaginationNormalized(
        XElement table,
        XElement? caption,
        TablePaginationMode mode,
        ProfileTablePaginationLayoutPolicy policy)
    {
        if (!policy.Enabled) return true;

        var rows = table.Elements(W + "tr").ToList();
        if (rows.Count == 0) return true;

        if (policy.NoRowSplit &&
            rows.Any(r => !ElementIsOn(r.Element(W + "trPr")?.Element(W + "cantSplit"))))
            return false;

        if (policy.RepeatHeader &&
            !ElementIsOn(rows[0].Element(W + "trPr")?.Element(W + "tblHeader")))
            return false;

        if (policy.KeepCaptionWithTable &&
            caption is not null &&
            !ElementIsOn(caption.Element(W + "pPr")?.Element(W + "keepNext")))
            return false;

        if (mode == TablePaginationMode.KeepWhole)
        {
            for (int i = 0; i < rows.Count - 1; i++)
            {
                if (!RowParagraphsHaveKeepNext(rows[i]))
                    return false;
            }
            return true;
        }

        if (policy.MinBodyRowsAtStart > 0 && !RowParagraphsHaveKeepNext(rows[0]))
            return false;

        var additionalStartLinks = Math.Min(
            Math.Max(0, policy.MinBodyRowsAtStart - 1),
            Math.Max(0, rows.Count - 2));

        for (int i = 1; i <= additionalStartLinks; i++)
        {
            if (!RowParagraphsHaveKeepNext(rows[i]))
                return false;
        }

        if (rows.Count >= 3 && policy.MinBodyRowsAtEnd >= 2)
        {
            var endLinks = Math.Min(policy.MinBodyRowsAtEnd - 1, rows.Count - 2);
            for (int offset = 2; offset <= endLinks + 1; offset++)
            {
                if (!RowParagraphsHaveKeepNext(rows[^offset]))
                    return false;
            }
        }

        return true;
    }

    private static void EnsureKeepNextOnRowParagraphs(XElement row)
    {
        foreach (var p in row.Descendants(W + "p"))
            EnsureKeepNext(p);
    }

    private static bool RowParagraphsHaveKeepNext(XElement row)
    {
        var paragraphs = row.Descendants(W + "p").ToList();
        return paragraphs.Count == 0 ||
               paragraphs.All(p => ElementIsOn(p.Element(W + "pPr")?.Element(W + "keepNext")));
    }

    private static void EnsureKeepNext(XElement paragraph)
    {
        var pPr = EnsurePPr(paragraph);
        var keepNext = pPr.Element(W + "keepNext");
        if (keepNext is null)
        {
            keepNext = new XElement(W + "keepNext");
            InsertParagraphPropertyInSchemaOrder(pPr, keepNext);
        }
        keepNext.SetAttributeValue(W + "val", "1");
    }
    private static void NormalizeTable(
        XElement table,
        int targetWidth,
        XElement? caption,
        ProfileAppendixTableLayoutPolicy layoutPolicy,
        ProfileTablePaginationLayoutPolicy paginationPolicy)
    {
        var tblPr = table.Element(W + "tblPr");
        if (tblPr is null)
        {
            tblPr = new XElement(W + "tblPr");
            table.AddFirst(tblPr);
        }

        if (layoutPolicy.RtlVisual)
        {
            var bidiVisual = tblPr.Element(W + "bidiVisual");
            if (bidiVisual is null)
            {
                bidiVisual = new XElement(W + "bidiVisual");
                InsertTablePropertyInSchemaOrder(tblPr, bidiVisual);
            }
            SetOnOffOnlyEnabled(bidiVisual);
        }

        if (layoutPolicy.FitToTextWidth)
            SetTableProperty(tblPr, "tblW", "type", "dxa", ("w", targetWidth.ToString()));

        if (!string.IsNullOrWhiteSpace(layoutPolicy.TableAlignment))
            SetTableProperty(tblPr, "jc", "val", layoutPolicy.TableAlignment.Trim().ToLowerInvariant());

        if (layoutPolicy.FixedLayout)
            SetTableProperty(tblPr, "tblLayout", "type", "fixed");

        if (layoutPolicy.NormalizeCellMargins)
        {
            var cellMar = tblPr.Element(W + "tblCellMar");
            if (cellMar is null)
            {
                cellMar = new XElement(W + "tblCellMar");
                InsertTablePropertyInSchemaOrder(tblPr, cellMar);
            }

            foreach (var side in new[] { "top", "left", "bottom", "right" })
            {
                var element = cellMar.Element(W + side);
                if (element is null)
                {
                    element = new XElement(W + side);
                    cellMar.Add(element);
                }

                element.SetAttributeValue(W + "w", layoutPolicy.CellMarginTwips.ToString());
                element.SetAttributeValue(W + "type", "dxa");
            }
        }

        var rows = table.Elements(W + "tr").ToList();

        if (paginationPolicy.Enabled && paginationPolicy.RepeatHeader && rows.Count > 0)
        {
            var trPr = EnsureTrPr(rows[0]);
            var header = trPr.Element(W + "tblHeader");
            if (header is null)
            {
                header = new XElement(W + "tblHeader");
                InsertRowPropertyInSchemaOrder(trPr, header);
            }
            SetOnOffOnlyEnabled(header);
        }

        if (paginationPolicy.Enabled && paginationPolicy.NoRowSplit)
        {
            foreach (var row in rows)
            {
                var trPr = EnsureTrPr(row);
                var cantSplit = trPr.Element(W + "cantSplit");
                if (cantSplit is null)
                {
                    cantSplit = new XElement(W + "cantSplit");
                    InsertRowPropertyInSchemaOrder(trPr, cantSplit);
                }
                SetOnOffOnlyEnabled(cantSplit);
            }
        }

        if (paginationPolicy.Enabled &&
            paginationPolicy.KeepCaptionWithTable &&
            caption is not null)
        {
            EnsureKeepNext(caption);
        }
    }

    private static bool IsTableNormalized(
        XElement table,
        int targetWidth,
        XElement? caption,
        ProfileAppendixTableLayoutPolicy layoutPolicy,
        ProfileTablePaginationLayoutPolicy paginationPolicy)
    {
        var tblPr = table.Element(W + "tblPr");
        if (tblPr is null) return false;

        if (layoutPolicy.FitToTextWidth)
        {
            var width = ParseInt((string?)tblPr.Element(W + "tblW")?.Attribute(W + "w"));
            var widthType = (string?)tblPr.Element(W + "tblW")?.Attribute(W + "type");

            if (!string.Equals(widthType, "dxa", StringComparison.OrdinalIgnoreCase) ||
                Math.Abs(width - targetWidth) > 6)
                return false;
        }

        if (!string.IsNullOrWhiteSpace(layoutPolicy.TableAlignment) &&
            !string.Equals(
                (string?)tblPr.Element(W + "jc")?.Attribute(W + "val"),
                layoutPolicy.TableAlignment.Trim(),
                StringComparison.OrdinalIgnoreCase))
            return false;

        if (layoutPolicy.FixedLayout &&
            !TokensEqual((string?)tblPr.Element(W + "tblLayout")?.Attribute(W + "type"), "fixed"))
            return false;

        if (layoutPolicy.RtlVisual && !ElementIsOn(tblPr.Element(W + "bidiVisual")))
            return false;

        if (layoutPolicy.NormalizeCellMargins)
        {
            var cellMar = tblPr.Element(W + "tblCellMar");
            if (cellMar is null) return false;

            foreach (var side in new[] { "top", "left", "bottom", "right" })
            {
                var margin = cellMar.Element(W + side);
                if (margin is null ||
                    !TokensEqual((string?)margin.Attribute(W + "type"), "dxa") ||
                    ParseInt((string?)margin.Attribute(W + "w")) != layoutPolicy.CellMarginTwips)
                    return false;
            }
        }

        var rows = table.Elements(W + "tr").ToList();

        if (paginationPolicy.Enabled &&
            paginationPolicy.RepeatHeader &&
            rows.Count > 0 &&
            !ElementIsOn(rows[0].Element(W + "trPr")?.Element(W + "tblHeader")))
            return false;

        if (paginationPolicy.Enabled &&
            paginationPolicy.NoRowSplit &&
            rows.Any(r => !ElementIsOn(r.Element(W + "trPr")?.Element(W + "cantSplit"))))
            return false;

        if (paginationPolicy.Enabled &&
            paginationPolicy.KeepCaptionWithTable &&
            caption is not null &&
            !ElementIsOn(caption.Element(W + "pPr")?.Element(W + "keepNext")))
            return false;

        if (layoutPolicy.FitToTextWidth)
        {
            var gridWidth = GetGridWidth(table);
            if (gridWidth > 0 && Math.Abs(gridWidth - targetWidth) > 20)
                return false;
        }

        return true;
    }
    private sealed record TableThemePalette(
        string HeaderFill,
        string BandFill,
        string AlternateFill,
        string BorderColor);

    private static int NormalizeAllTableStyle(
        XDocument document,
        ProfileTableStyleLayoutPolicy policy)
    {
        if (!policy.Enabled) return 0;

        int changed = 0;
        foreach (var table in document.Descendants(W + "tbl"))
        {
            if (IsTableThemeNormalized(table, policy))
                continue;

            ApplyTableTheme(table, policy);
            changed++;
        }

        return changed;
    }

    private static bool IsTableThemeNormalized(
        XElement table,
        ProfileTableStyleLayoutPolicy policy)
    {
        if (!policy.Enabled) return true;

        var palette = GetTableThemePalette(policy.Theme);
        var tblPr = table.Element(W + "tblPr");
        if (tblPr is null) return false;

        if (policy.CenterTable &&
            !TokensEqual((string?)tblPr.Element(W + "jc")?.Attribute(W + "val"), "center"))
            return false;

        if (policy.RtlVisual && !ElementIsOn(tblPr.Element(W + "bidiVisual")))
            return false;

        var borders = tblPr.Element(W + "tblBorders");
        if (borders is null) return false;
        foreach (var side in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
        {
            var border = borders.Element(W + side);
            if (border is null) return false;
            if (!TokensEqual((string?)border.Attribute(W + "val"), "single")) return false;
            if (!string.Equals((string?)border.Attribute(W + "color"), palette.BorderColor, StringComparison.OrdinalIgnoreCase)) return false;
            if (ParseInt((string?)border.Attribute(W + "sz")) != policy.BorderSize) return false;
        }

        var rows = table.Elements(W + "tr").ToList();
        for (int r = 0; r < rows.Count; r++)
        {
            var expectedFill = r == 0
                ? palette.HeaderFill
                : policy.BandRows
                    ? (r % 2 == 1 ? palette.BandFill : palette.AlternateFill)
                    : palette.AlternateFill;

            foreach (var cell in rows[r].Elements(W + "tc"))
            {
                var tcPr = cell.Element(W + "tcPr");
                if (tcPr is null) return false;

                var vAlign = (string?)tcPr.Element(W + "vAlign")?.Attribute(W + "val");
                if (!TokensEqual(vAlign, policy.VerticalAlignment)) return false;

                var fill = (string?)tcPr.Element(W + "shd")?.Attribute(W + "fill");
                if (!string.Equals(fill, expectedFill, StringComparison.OrdinalIgnoreCase)) return false;

                foreach (var paragraph in cell.Descendants(W + "p"))
                {
                    var text = ParagraphText(paragraph);
                    var expectedAlignment = r == 0
                        ? policy.HeaderAlignment
                        : DecideCellAlignment(text, policy);

                    var pPr = paragraph.Element(W + "pPr");
                    var jc = (string?)pPr?.Element(W + "jc")?.Attribute(W + "val");
                    var expectedJc = EncodePhysicalAlignmentForRtl(expectedAlignment, ContainsPersian(text));
                    if (!TokensEqual(jc, expectedJc)) return false;

                    if (ContainsPersian(text) &&
                        !ElementIsOn(pPr?.Element(W + "bidi")))
                        return false;
                }
            }
        }

        return true;
    }

    private static void ApplyTableTheme(
        XElement table,
        ProfileTableStyleLayoutPolicy policy)
    {
        var palette = GetTableThemePalette(policy.Theme);
        var tblPr = table.Element(W + "tblPr");
        if (tblPr is null)
        {
            tblPr = new XElement(W + "tblPr");
            table.AddFirst(tblPr);
        }

        if (policy.CenterTable)
            SetTableProperty(tblPr, "jc", "val", "center");

        if (policy.RtlVisual)
        {
            var bidiVisual = tblPr.Element(W + "bidiVisual");
            if (bidiVisual is null)
            {
                bidiVisual = new XElement(W + "bidiVisual");
                InsertTablePropertyInSchemaOrder(tblPr, bidiVisual);
            }
            SetOnOffOnlyEnabled(bidiVisual);
        }

        SetTableBorders(tblPr, palette.BorderColor, policy.BorderSize);

        var rows = table.Elements(W + "tr").ToList();
        for (int r = 0; r < rows.Count; r++)
        {
            var fill = r == 0
                ? palette.HeaderFill
                : policy.BandRows
                    ? (r % 2 == 1 ? palette.BandFill : palette.AlternateFill)
                    : palette.AlternateFill;

            foreach (var cell in rows[r].Elements(W + "tc"))
            {
                var tcPr = EnsureTcPr(cell);
                SetCellShading(tcPr, fill);
                SetCellVerticalAlignment(tcPr, policy.VerticalAlignment);

                foreach (var paragraph in cell.Descendants(W + "p"))
                {
                    var text = ParagraphText(paragraph);
                    var alignment = r == 0
                        ? policy.HeaderAlignment
                        : DecideCellAlignment(text, policy);
                    MaterializeCellParagraphAlignment(paragraph, alignment, ContainsPersian(text));
                }
            }
        }
    }

    private static string DecideCellAlignment(string text, ProfileTableStyleLayoutPolicy policy)
    {
        var compact = NormalizeText(text);
        if (compact.Length == 0) return policy.ShortCellAlignment;
        if (LooksNumericOrCode(compact)) return policy.ShortCellAlignment;
        if (ContainsPersian(compact) && compact.Length > policy.LongTextThreshold)
            return policy.LongPersianAlignment;
        return policy.ShortCellAlignment;
    }

    private static bool LooksNumericOrCode(string text)
    {
        int letters = 0;
        int digits = 0;
        foreach (var ch in text)
        {
            if (char.IsLetter(ch)) letters++;
            if (char.IsDigit(ch) || (ch >= '\u06F0' && ch <= '\u06F9') || (ch >= '\u0660' && ch <= '\u0669')) digits++;
        }

        return digits > 0 && letters <= Math.Max(2, digits / 2);
    }

    private static bool ContainsPersian(string? text)
    {
        if (string.IsNullOrEmpty(text)) return false;
        foreach (var ch in text)
        {
            if ((ch >= '\u0600' && ch <= '\u06FF') ||
                (ch >= '\u0750' && ch <= '\u077F') ||
                (ch >= '\u08A0' && ch <= '\u08FF'))
                return true;
        }
        return false;
    }

    private static string EncodePhysicalAlignmentForRtl(string alignment, bool rtl)
    {
        var token = (alignment ?? string.Empty).Trim().ToUpperInvariant();
        return token switch
        {
            "RIGHT" when rtl => "left",
            "LEFT" when rtl => "right",
            "RIGHT" => "right",
            "LEFT" => "left",
            "CENTER" => "center",
            "JUSTIFY" => "both",
            "DISTRIBUTE" => "distribute",
            _ => throw new InvalidOperationException($"Unsupported table-cell alignment: {alignment}")
        };
    }

    private static void MaterializeCellParagraphAlignment(XElement paragraph, string alignment, bool rtl)
    {
        var pPr = EnsurePPr(paragraph);

        var jc = pPr.Element(W + "jc");
        if (jc is null)
        {
            jc = new XElement(W + "jc");
            InsertParagraphPropertyInSchemaOrder(pPr, jc);
        }
        jc.SetAttributeValue(W + "val", EncodePhysicalAlignmentForRtl(alignment, rtl));

        if (rtl)
        {
            var bidi = pPr.Element(W + "bidi");
            if (bidi is null)
            {
                bidi = new XElement(W + "bidi");
                InsertParagraphPropertyInSchemaOrder(pPr, bidi);
            }
            bidi.SetAttributeValue(W + "val", "1");
        }
    }

    private static XElement EnsureTcPr(XElement cell)
    {
        var tcPr = cell.Element(W + "tcPr");
        if (tcPr is not null) return tcPr;
        tcPr = new XElement(W + "tcPr");
        cell.AddFirst(tcPr);
        return tcPr;
    }

    private static void SetCellShading(XElement tcPr, string fill)
    {
        var shd = tcPr.Element(W + "shd");
        if (shd is null)
        {
            shd = new XElement(W + "shd");
            InsertCellPropertyInSchemaOrder(tcPr, shd);
        }
        shd.SetAttributeValue(W + "val", "clear");
        shd.SetAttributeValue(W + "color", "auto");
        shd.SetAttributeValue(W + "fill", fill);
    }

    private static void SetCellVerticalAlignment(XElement tcPr, string alignment)
    {
        var vAlign = tcPr.Element(W + "vAlign");
        if (vAlign is null)
        {
            vAlign = new XElement(W + "vAlign");
            InsertCellPropertyInSchemaOrder(tcPr, vAlign);
        }
        vAlign.SetAttributeValue(W + "val", alignment.Trim().ToLowerInvariant());
    }

    private static void SetTableBorders(XElement tblPr, string color, int size)
    {
        var borders = tblPr.Element(W + "tblBorders");
        if (borders is null)
        {
            borders = new XElement(W + "tblBorders");
            InsertTablePropertyInSchemaOrder(tblPr, borders);
        }

        foreach (var side in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
        {
            var border = borders.Element(W + side);
            if (border is null)
            {
                border = new XElement(W + side);
                borders.Add(border);
            }
            border.SetAttributeValue(W + "val", "single");
            border.SetAttributeValue(W + "sz", size.ToString());
            border.SetAttributeValue(W + "space", "0");
            border.SetAttributeValue(W + "color", color);
        }
    }

    private static string NormalizeTableTheme(string? theme)
        => (theme ?? "LIGHT_BLUE").Trim().ToUpperInvariant().Replace('-', '_');

    private static TableThemePalette GetTableThemePalette(string? theme)
        => NormalizeTableTheme(theme) switch
        {
            "LIGHT_BLUE" => new TableThemePalette("BDD7EE", "DDEBF7", "F7FBFF", "9EADBA"),
            "BLUE" => new TableThemePalette("9DC3E6", "D9EAF7", "EDF4FA", "5B9BD5"),
            "GREEN" => new TableThemePalette("A9D18E", "E2F0D9", "F3F8EF", "70AD47"),
            var other => throw new InvalidOperationException($"Unsupported table theme: {other}")
        };

    private static void ExpandTableGrid(XElement table, int targetWidth)
    {
        var grid = table.Element(W + "tblGrid");
        if (grid is null) return;
        var cols = grid.Elements(W + "gridCol").ToList();
        var widths = cols.Select(x => ParseInt((string?)x.Attribute(W + "w"))).ToArray();
        var total = widths.Sum();
        if (total <= 0 || Math.Abs(total - targetWidth) <= 20) return;

        var scale = (double)targetWidth / total;
        int written = 0;
        for (int i = 0; i < cols.Count; i++)
        {
            var value = i == cols.Count - 1
                ? targetWidth - written
                : Math.Max(1, (int)Math.Round(widths[i] * scale));
            cols[i].SetAttributeValue(W + "w", value.ToString());
            written += value;
        }

        foreach (var tcW in table.Descendants(W + "tcW"))
        {
            if (!TokensEqual((string?)tcW.Attribute(W + "type"), "dxa")) continue;
            var old = ParseInt((string?)tcW.Attribute(W + "w"));
            if (old <= 0) continue;
            tcW.SetAttributeValue(W + "w", Math.Max(1, (int)Math.Round(old * scale)).ToString());
        }
    }

    private static void InsertLandscapeTailSection(XElement body, XElement caption, SectionMetrics current, bool insertBreak)
    {
        var finalSectPr = body.Element(W + "sectPr")
            ?? throw new InvalidOperationException("Final w:sectPr was not found.");

        if (insertBreak)
        {
            var breakParagraph = new XElement(W + "p");
            var pPr = new XElement(W + "pPr");
            breakParagraph.Add(pPr);
            var portraitSectPr = new XElement(finalSectPr);
            EnsureSectionTypeNextPage(portraitSectPr);
            SetSectionOrientation(portraitSectPr, landscape: false, current);
            ContinuePageNumbering(portraitSectPr);
            pPr.Add(portraitSectPr);
            caption.AddBeforeSelf(breakParagraph);
        }

        EnsureSectionTypeNextPage(finalSectPr);
        SetSectionOrientation(finalSectPr, landscape: true, current);
        ContinuePageNumbering(finalSectPr);
    }

    private static void EnsureSectionTypeNextPage(XElement sectPr)
    {
        var type = sectPr.Element(W + "type");
        if (type is null)
        {
            type = new XElement(W + "type");
            InsertSectionPropertyInSchemaOrder(sectPr, type);
        }
        type.SetAttributeValue(W + "val", "nextPage");
    }

    private static void SetSectionOrientation(XElement sectPr, bool landscape, SectionMetrics current)
    {
        var pageSize = sectPr.Element(W + "pgSz");
        if (pageSize is null)
        {
            pageSize = new XElement(W + "pgSz");
            InsertSectionPropertyInSchemaOrder(sectPr, pageSize);
        }

        var portraitWidth = Math.Min(current.PageWidthTwips, current.PageHeightTwips);
        var portraitHeight = Math.Max(current.PageWidthTwips, current.PageHeightTwips);
        if (portraitWidth <= 0) portraitWidth = A4PortraitWidthTwips;
        if (portraitHeight <= 0) portraitHeight = A4PortraitHeightTwips;

        if (landscape)
        {
            pageSize.SetAttributeValue(W + "w", portraitHeight.ToString());
            pageSize.SetAttributeValue(W + "h", portraitWidth.ToString());
            pageSize.SetAttributeValue(W + "orient", "landscape");
        }
        else
        {
            pageSize.SetAttributeValue(W + "w", portraitWidth.ToString());
            pageSize.SetAttributeValue(W + "h", portraitHeight.ToString());
            pageSize.SetAttributeValue(W + "orient", null);
        }
    }

    private static void ContinuePageNumbering(XElement sectPr)
    {
        var pageNumber = sectPr.Element(W + "pgNumType");
        pageNumber?.SetAttributeValue(W + "start", null);
    }

    private static bool IsFinalSectionLandscape(XElement body)
    {
        var sectPr = body.Element(W + "sectPr");
        if (sectPr is null) return false;
        var pgSz = sectPr.Element(W + "pgSz");
        if (pgSz is null) return false;
        var orient = NormalizeText((string?)pgSz.Attribute(W + "orient"));
        var w = ParseInt((string?)pgSz.Attribute(W + "w"));
        var h = ParseInt((string?)pgSz.Attribute(W + "h"));
        return orient.Equals("landscape", StringComparison.OrdinalIgnoreCase) || (w > 0 && h > 0 && w > h);
    }

    private static bool HasLandscapeBreakBeforeFirstWide(
        IReadOnlyList<XElement> directChildren,
        IReadOnlyList<XElement> appendixTables,
        IReadOnlyList<bool> wideFlags)
    {
        var firstWide = wideFlags.ToList().FindIndex(x => x);
        if (firstWide < 0) return true;
        var table = appendixTables[firstWide];
        var caption = FindNearestCaption(directChildren, table);
        if (caption is null) return false;
        var captionIndex = IndexOfReference(directChildren, caption);
        if (captionIndex <= 0) return false;
        for (int i = captionIndex - 1; i >= 0 && i >= captionIndex - 2; i--)
        {
            var candidate = directChildren[i];
            if (candidate.Name != W + "p") continue;
            if (!string.IsNullOrWhiteSpace(ParagraphText(candidate))) continue;
            if (candidate.Element(W + "pPr")?.Element(W + "sectPr") is not null) return true;
        }
        return false;
    }

    private static int CountTocStylesNeedingCompaction(
        XDocument styles,
        IReadOnlyCollection<string> tocStyleIds,
        ProfileTocLayoutPolicy policy)
    {
        if (!policy.Enabled || !policy.CompactStyles) return 0;

        int count = 0;
        foreach (var styleId in tocStyleIds)
        {
            var style = styles.Descendants(W + "style")
                .FirstOrDefault(x => TokensEqual((string?)x.Attribute(W + "styleId"), styleId));
            if (style is null) continue;

            var spacing = style.Element(W + "pPr")?.Element(W + "spacing");
            var expectedBefore = policy.Spacing.BeforeTwips.ToString();
            var expectedAfter = policy.Spacing.AfterTwips.ToString();
            var expectedLine = policy.Spacing.LineTwips.ToString();
            var expectedRule = policy.Spacing.LineRule;

            var before = (string?)spacing?.Attribute(W + "before") ?? expectedBefore;
            var after = (string?)spacing?.Attribute(W + "after") ?? expectedAfter;
            var line = (string?)spacing?.Attribute(W + "line") ?? expectedLine;
            var rule = (string?)spacing?.Attribute(W + "lineRule") ?? expectedRule;

            if (before != expectedBefore ||
                after != expectedAfter ||
                line != expectedLine ||
                !rule.Equals(expectedRule, StringComparison.OrdinalIgnoreCase))
            {
                count++;
            }
        }

        return count;
    }
    private static XElement? FindTocTitle(
        IReadOnlyList<XElement> paragraphs,
        StyleCatalog catalog,
        AgentProfile profile,
        ProfileTocLayoutPolicy policy)
    {
        var aliases = GetAliases(profile, "TocTitle");
        var byStyle = paragraphs.FirstOrDefault(p =>
            aliases.Any(x => TokensEqual(catalog.GetParagraphStyleName(p), x)));

        if (byStyle is not null) return byStyle;

        return paragraphs.FirstOrDefault(p =>
            policy.TitleTexts.Any(x =>
                NormalizeText(ParagraphText(p)).Equals(
                    NormalizeText(x),
                    StringComparison.Ordinal)));
    }
    private static XElement? FindFirstChapterAfter(
        IReadOnlyList<XElement> paragraphs,
        XElement tocTitle,
        StyleCatalog catalog,
        AgentProfile profile)
    {
        var aliases = GetAliases(profile, "Chapter");
        var start = IndexOfReference(paragraphs, tocTitle);
        if (start < 0) return null;
        for (int i = start + 1; i < paragraphs.Count; i++)
        {
            var paragraph = paragraphs[i];
            if (string.IsNullOrWhiteSpace(ParagraphText(paragraph))) continue;
            var style = catalog.GetParagraphStyleName(paragraph);
            if (aliases.Any(x => TokensEqual(style, x))) return paragraph;
        }
        return null;
    }

    private static XElement? FindAppendixStart(
        IReadOnlyList<XElement> directChildren,
        StyleCatalog catalog,
        AgentProfile profile)
    {
        var chapterAliases = GetAliases(profile, "Chapter");
        var appendixPolicy = profile.SemanticPolicy.Appendix;

        foreach (var child in directChildren)
        {
            if (child.Name != W + "p") continue;
            var text = NormalizeText(ParagraphText(child));
            if (!SemanticPolicyMatcher.MatchesAppendixHeading(text, appendixPolicy)) continue;

            var style = catalog.GetParagraphStyleName(child);
            if (chapterAliases.Count == 0 || chapterAliases.Any(x => TokensEqual(style, x)))
                return child;
        }

        return directChildren.FirstOrDefault(x =>
            x.Name == W + "p" &&
            SemanticPolicyMatcher.MatchesAppendixHeading(
                NormalizeText(ParagraphText(x)),
                appendixPolicy));
    }
    private static XElement? FindNearestCaption(IReadOnlyList<XElement> directChildren, XElement table)
    {
        var index = IndexOfReference(directChildren, table);
        if (index <= 0) return null;
        for (int i = index - 1; i >= 0; i--)
        {
            var candidate = directChildren[i];
            if (candidate.Name == W + "tbl") return null;
            if (candidate.Name != W + "p") continue;
            if (string.IsNullOrWhiteSpace(ParagraphText(candidate))) continue;
            return candidate;
        }
        return null;
    }

    private static int RemoveSafeTrailingBlankParagraphs(XDocument document)
    {
        var body = document.Root?.Element(W + "body");
        if (body is null) return 0;

        var removed = 0;
        while (true)
        {
            var candidate = body.Elements().LastOrDefault(x => x.Name != W + "sectPr");
            if (candidate is null || candidate.Name != W + "p") break;
            if (!string.IsNullOrWhiteSpace(ParagraphText(candidate))) break;

            var pPr = candidate.Element(W + "pPr");
            if (pPr?.Element(W + "sectPr") is not null) break;

            // Conservative hygiene rule: remove only a terminal paragraph whose
            // direct content is formatting-only. Any run, bookmark, field,
            // drawing, comment anchor, SDT, proof marker, or other payload blocks
            // removal even when it has no visible text.
            if (candidate.Elements().Any(x => x.Name != W + "pPr")) break;

            candidate.Remove();
            removed++;
        }

        return removed;
    }

    private static bool HasOnlyBlankContentAfter(IReadOnlyList<XElement> directChildren, XElement element)
    {
        var index = IndexOfReference(directChildren, element);
        if (index < 0) return false;
        for (int i = index + 1; i < directChildren.Count; i++)
        {
            var candidate = directChildren[i];
            if (candidate.Name == W + "sectPr") continue;
            if (candidate.Name == W + "p" && string.IsNullOrWhiteSpace(ParagraphText(candidate))) continue;
            return false;
        }
        return true;
    }

    private static int GetColumnCount(XElement table)
    {
        var gridCount = table.Element(W + "tblGrid")?.Elements(W + "gridCol").Count() ?? 0;
        if (gridCount > 0) return gridCount;
        return table.Elements(W + "tr").Select(r => r.Elements(W + "tc").Count()).DefaultIfEmpty(0).Max();
    }

    private static int GetGridWidth(XElement table)
        => table.Element(W + "tblGrid")?.Elements(W + "gridCol")
            .Select(x => ParseInt((string?)x.Attribute(W + "w")))
            .Sum() ?? 0;

    private static SectionMetrics ReadFinalSectionMetrics(XElement body)
    {
        var sectPr = body.Element(W + "sectPr");
        if (sectPr is null)
            return new SectionMetrics(A4PortraitWidthTwips, A4PortraitHeightTwips, 1417, 1417);

        var pgSz = sectPr.Element(W + "pgSz");
        var pgMar = sectPr.Element(W + "pgMar");
        return new SectionMetrics(
            ParseInt((string?)pgSz?.Attribute(W + "w")),
            ParseInt((string?)pgSz?.Attribute(W + "h")),
            ParseInt((string?)pgMar?.Attribute(W + "left")),
            ParseInt((string?)pgMar?.Attribute(W + "right")));
    }

    private sealed record SectionMetrics(int PageWidthTwips, int PageHeightTwips, int LeftMarginTwips, int RightMarginTwips);

    private static void SetTableProperty(XElement tblPr, string localName, string attrName, string attrValue, params (string Name, string Value)[] additional)
    {
        var element = tblPr.Element(W + localName);
        if (element is null)
        {
            element = new XElement(W + localName);
            InsertTablePropertyInSchemaOrder(tblPr, element);
        }
        element.SetAttributeValue(W + attrName, attrValue);
        foreach (var item in additional) element.SetAttributeValue(W + item.Name, item.Value);
    }

    private static XElement EnsurePPr(XElement paragraph)
    {
        var pPr = paragraph.Element(W + "pPr");
        if (pPr is not null) return pPr;
        pPr = new XElement(W + "pPr");
        paragraph.AddFirst(pPr);
        return pPr;
    }

    private static XElement EnsureTrPr(XElement row)
    {
        var trPr = row.Element(W + "trPr");
        if (trPr is not null) return trPr;
        trPr = new XElement(W + "trPr");
        row.AddFirst(trPr);
        return trPr;
    }

    private static void InsertParagraphPropertyInSchemaOrder(XElement pPr, XElement property)
    {
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
        InsertByRank(pPr, property, Rank);
    }

    private static void InsertTablePropertyInSchemaOrder(XElement tblPr, XElement property)
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
        InsertByRank(tblPr, property, Rank);
    }

    private static void InsertRowPropertyInSchemaOrder(XElement trPr, XElement property)
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
        InsertByRank(trPr, property, Rank);
    }

    private static void InsertCellPropertyInSchemaOrder(XElement tcPr, XElement property)
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
        InsertByRank(tcPr, property, Rank);
    }

    private static void InsertSectionPropertyInSchemaOrder(XElement sectPr, XElement property)
    {
        static int Rank(string name) => name switch
        {
            "headerReference" => 10,
            "footerReference" => 20,
            "footnotePr" => 30,
            "endnotePr" => 40,
            "type" => 50,
            "pgSz" => 60,
            "pgMar" => 70,
            "paperSrc" => 80,
            "pgBorders" => 90,
            "lnNumType" => 100,
            "pgNumType" => 110,
            "cols" => 120,
            "formProt" => 130,
            "vAlign" => 140,
            "noEndnote" => 150,
            "titlePg" => 160,
            "textDirection" => 170,
            "bidi" => 180,
            "rtlGutter" => 190,
            "docGrid" => 200,
            "printerSettings" => 210,
            "sectPrChange" => 220,
            _ => 1000
        };
        InsertByRank(sectPr, property, Rank);
    }

    private static void InsertByRank(XElement parent, XElement property, Func<string, int> rank)
    {
        var target = rank(property.Name.LocalName);
        var before = parent.Elements().FirstOrDefault(x => rank(x.Name.LocalName) > target);
        if (before is null) parent.Add(property);
        else before.AddBeforeSelf(property);
    }

    private static string ParagraphText(XElement paragraph)
        => string.Concat(paragraph.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));

    private static string Preview(string? text, int max)
    {
        var value = NormalizeText(text);
        return value.Length <= max ? value : value[..max] + "...";
    }

    private static IReadOnlyList<string> GetAliases(AgentProfile profile, string role)
        => profile.RoleAliases.TryGetValue(role, out var aliases)
            ? aliases.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray()
            : Array.Empty<string>();

    private static int ParseInt(string? value)
        => int.TryParse(value, out var parsed) ? parsed : 0;

    private static void SetOnOffOnlyEnabled(XElement element)
    {
        // For OnOffOnlyType, the canonical ON representation is an empty
        // element (for example <w:cantSplit/> or <w:tblHeader/>).
        // This avoids Office2016 validator failures for w:val="1".
        element.Attribute(W + "val")?.Remove();
    }

    private static void NormalizeOnOffOnlySchemaLiterals(XDocument document)
    {
        var names = new HashSet<XName>
        {
            W + "cantSplit",
            W + "tblHeader",
            W + "bidiVisual"
        };

        foreach (var element in document.Descendants().Where(e => names.Contains(e.Name)))
        {
            var attr = element.Attribute(W + "val");
            if (attr is null) continue;

            // Nullable-flow stabilization only. Runtime/XML behavior is unchanged:
            // the immediately preceding guard proves the attribute is non-null.
            var nonNullAttr = attr!;
            var value = nonNullAttr.Value.Trim().ToLowerInvariant();
            if (value is "1" or "true" or "on")
            {
                nonNullAttr.Remove(); // presence => ON
            }
            else if (value is "0" or "false" or "off")
            {
                nonNullAttr.Value = "off";
            }
        }
    }

    private static void AssertOnOffOnlySchemaContract(XDocument document)
    {
        var names = new HashSet<XName>
        {
            W + "cantSplit",
            W + "tblHeader",
            W + "bidiVisual"
        };

        var invalid = document.Descendants()
            .Where(e => names.Contains(e.Name))
            .Select(e => new { Element = e, Val = (string?)e.Attribute(W + "val") })
            .Where(x => x.Val is not null &&
                        !string.Equals(x.Val, "on", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(x.Val, "off", StringComparison.OrdinalIgnoreCase))
            .Take(8)
            .ToList();

        if (invalid.Count == 0) return;

        var details = string.Join(", ", invalid.Select(x => $"{x.Element.Name.LocalName}={x.Val}"));
        throw new InvalidOperationException("OOXML OnOffOnly schema guard failed: " + details);
    }

    private static bool IsOn(XAttribute? attribute)
    {
        if (attribute is null) return false;
        var value = ((string?)attribute ?? string.Empty).Trim();
        return value.Length == 0 || value is "1" or "true" or "on";
    }

    private static bool ElementIsOn(XElement? element)
    {
        if (element is null) return false;
        var value = element.Attribute(W + "val");
        return value is null || IsOn(value);
    }

    private static int IndexOfReference(IReadOnlyList<XElement> items, XElement target)
    {
        for (int i = 0; i < items.Count; i++)
            if (ReferenceEquals(items[i], target)) return i;
        return -1;
    }

    private static bool TokensEqual(string? a, string? b)
        => string.Equals(NormalizeText(a), NormalizeText(b), StringComparison.OrdinalIgnoreCase);

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

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
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

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }

    private sealed class StyleCatalog
    {
        private readonly Dictionary<string, string> _idToName;
        private readonly HashSet<string> _tocStyleIds;
        private readonly string _defaultParagraphStyleId;
        private readonly string _defaultParagraphStyleName;

        private StyleCatalog(
            Dictionary<string, string> idToName,
            HashSet<string> tocStyleIds,
            string defaultParagraphStyleId,
            string defaultParagraphStyleName)
        {
            _idToName = idToName;
            _tocStyleIds = tocStyleIds;
            _defaultParagraphStyleId = defaultParagraphStyleId;
            _defaultParagraphStyleName = defaultParagraphStyleName;
        }

        public static StyleCatalog Build(XDocument styles)
        {
            var idToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var tocIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string? defaultId = null;
            string? defaultName = null;

            foreach (var style in styles.Descendants(W + "style"))
            {
                var id = (string?)style.Attribute(W + "styleId");
                var name = (string?)style.Element(W + "name")?.Attribute(W + "val");
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name)) continue;
                idToName[id] = name;

                var normalized = NormalizeText(name).ToLowerInvariant();
                if (normalized is "toc 1" or "toc 2" or "toc 3" || id.Equals("TOC1", StringComparison.OrdinalIgnoreCase) || id.Equals("TOC2", StringComparison.OrdinalIgnoreCase) || id.Equals("TOC3", StringComparison.OrdinalIgnoreCase))
                    tocIds.Add(id);

                var type = (string?)style.Attribute(W + "type");
                if (IsOn(style.Attribute(W + "default")) && string.Equals(type, "paragraph", StringComparison.OrdinalIgnoreCase))
                {
                    defaultId = id;
                    defaultName = name;
                }
            }

            defaultId ??= idToName.ContainsKey("Normal") ? "Normal" : idToName.Keys.FirstOrDefault() ?? "Normal";
            defaultName ??= idToName.TryGetValue(defaultId, out var value) ? value : "Normal";
            return new StyleCatalog(idToName, tocIds, defaultId, defaultName);
        }

        public string GetParagraphStyleId(XElement paragraph)
        {
            var id = (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
            return string.IsNullOrWhiteSpace(id) ? _defaultParagraphStyleId : id;
        }

        public string GetParagraphStyleName(XElement paragraph)
        {
            var id = GetParagraphStyleId(paragraph);
            return _idToName.TryGetValue(id, out var name) ? name : id;
        }

        public IReadOnlyCollection<string> GetTocStyleIds() => _tocStyleIds;
    }
}


