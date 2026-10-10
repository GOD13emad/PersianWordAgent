using System.Text.Json.Serialization;

namespace WordAgent.Core;

public sealed record InventoryCounts(
    int Paragraphs,
    int Tables,
    int InlineShapes,
    int FloatingShapes,
    int Fields,
    int EndNoteFields,
    int TocFields,
    int Sections,
    int Styles,
    int Footnotes,
    int Endnotes,
    int Comments,
    int Revisions,
    int ContentControls,
    int Bookmarks,
    int Hyperlinks,
    int Equations);

public sealed record FormattingMetrics(
    int NonEmptyParagraphs,
    int RtlParagraphs,
    int RightAlignedParagraphs,
    int CenteredParagraphs,
    int RtlTables,
    int CenteredTables,
    int InlineImagesCentered);

public sealed record WordAuditSnapshot(
    string FilePath,
    string WordVersion,
    InventoryCounts Counts,
    FormattingMetrics Formatting,
    IReadOnlyList<string> Warnings);

public sealed record OpenXmlAuditSnapshot(
    int Paragraphs,
    int Tables,
    int Drawings,
    int Styles,
    int FieldInstructions,
    int EndNoteMarkers,
    int ValidationErrors,
    IReadOnlyList<string> ValidationSamples);

public sealed record ComplianceSample(
    int ParagraphIndex,
    string StyleName,
    string TextPreview,
    string? ActualFontBi,
    float? ActualSizeBi,
    string ActualReadingOrder,
    string ActualAlignment,
    bool DirectOverrideSuspected,
    string Expected);

public sealed record UnclassifiedParagraphSample(
    int ParagraphIndex,
    string StyleName,
    string TextPreview);

public sealed record RoleComplianceSummary(
    string Role,
    string? ExpectedFontBi,
    float? ExpectedSizeBi,
    string? ExpectedReadingOrder,
    string? ExpectedAlignment,
    int Total,
    int FontBiMatches,
    int SizeBiMatches,
    int ReadingOrderMatches,
    int AlignmentMatches,
    int FullyCompliant,
    int DirectOverridesSuspected,
    IReadOnlyList<ComplianceSample> Samples);

public sealed record StyleComplianceItem(
    string Role,
    string StyleName,
    bool Found,
    string? ActualFontBi,
    float? ActualSizeBi,
    string ActualReadingOrder,
    string ActualAlignment,
    bool FontBiMatch,
    bool SizeBiMatch,
    bool ReadingOrderMatch,
    bool AlignmentMatch)
{
    public bool FullyCompliant => Found && FontBiMatch && SizeBiMatch && ReadingOrderMatch && AlignmentMatch;
}

public sealed record ProfileComplianceSnapshot(
    string ProfileId,
    string ProfileDisplayName,
    int NonEmptyParagraphs,
    int ClassifiedParagraphs,
    int UnclassifiedParagraphs,
    int ParagraphsInTables,
    int DirectOverridesSuspected,
    IReadOnlyList<RoleComplianceSummary> Roles,
    IReadOnlyList<StyleComplianceItem> Styles,
    IReadOnlyList<UnclassifiedParagraphSample> UnclassifiedSamples,
    IReadOnlyList<string> Warnings);

public sealed record CombinedAuditReport(
    string SchemaVersion,
    DateTimeOffset GeneratedAt,
    string InputFile,
    WordAuditSnapshot? Word,
    OpenXmlAuditSnapshot? OpenXml,
    ProfileComplianceSnapshot? Compliance,
    IReadOnlyList<string> Warnings);

public sealed record FixResult(
    string InputFile,
    string OutputFile,
    string ProfileId,
    int StylesChanged,
    int TablesChanged,
    int TableParagraphsChanged,
    int InlineImagesCentered,
    IReadOnlyList<string> Warnings);

public sealed class AgentProfile
{
    [JsonPropertyName("schemaVersion")]
    public string SchemaVersion { get; set; } = string.Empty;

    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("version")]
    public string Version { get; set; } = string.Empty;

    [JsonPropertyName("documentClass")]
    public string DocumentClass { get; set; } = string.Empty;

    [JsonPropertyName("language")]
    public string Language { get; set; } = string.Empty;

    [JsonPropertyName("classification")]
    public ProfileClassification Classification { get; set; } = new();

    [JsonPropertyName("roleAliases")]
    public Dictionary<string, string[]> RoleAliases { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    [JsonPropertyName("roleDefaults")]
    public Dictionary<string, RoleDefaults> RoleDefaults { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    [JsonPropertyName("rules")]
    public string[] Rules { get; set; } = Array.Empty<string>();

    [JsonPropertyName("referencePolicy")]
    public ProfileReferencePolicy ReferencePolicy { get; set; } = new();

    [JsonPropertyName("fieldPolicy")]
    public ProfileFieldPolicy FieldPolicy { get; set; } = new();

    [JsonPropertyName("outputPolicy")]
    public ProfileOutputPolicy OutputPolicy { get; set; } = new();

    [JsonPropertyName("semanticPolicy")]
    public ProfileSemanticPolicy SemanticPolicy { get; set; } = new();

    [JsonPropertyName("layoutPolicy")]
    public ProfileLayoutPolicy LayoutPolicy { get; set; } = new();

    [JsonPropertyName("editingPolicy")]
    public ProfileEditingPolicy EditingPolicy { get; set; } = new();

    [JsonPropertyName("detectionPolicy")]
    public ProfileDetectionPolicy DetectionPolicy { get; set; } = new();
}

public sealed class ProfileClassification
{
    [JsonPropertyName("keywords")]
    public string[] Keywords { get; set; } = Array.Empty<string>();

    // Optional high-authority document-intent / genre markers. Unlike ordinary
    // topic keywords these contribute only after a minimum distinct-match
    // threshold is reached, preventing one generic procurement term from
    // hijacking detection.
    [JsonPropertyName("intentKeywords")]
    public string[] IntentKeywords { get; set; } = Array.Empty<string>();

    // Optional profile-specific markers that suppress the high-authority intent bonus
    // when a conflicting document genre is strongly evidenced. Base keyword evidence
    // remains intact; only the intent bonus is excluded.
    [JsonPropertyName("intentExclusionKeywords")]
    public string[] IntentExclusionKeywords { get; set; } = Array.Empty<string>();

    [JsonPropertyName("negativeKeywords")]
    public string[] NegativeKeywords { get; set; } = Array.Empty<string>();

    [JsonPropertyName("styleSignals")]
    public string[] StyleSignals { get; set; } = Array.Empty<string>();
}

public sealed class RoleDefaults
{
    [JsonPropertyName("fontBi")]
    public string? FontBi { get; set; }

    [JsonPropertyName("sizeBi")]
    public float? SizeBi { get; set; }

    [JsonPropertyName("readingOrder")]
    public string? ReadingOrder { get; set; }

    [JsonPropertyName("alignment")]
    public string? Alignment { get; set; }
}

public sealed class ProfileReferencePolicy
{
    [JsonPropertyName("owner")]
    public string Owner { get; set; } = "EndNote";

    [JsonPropertyName("allowedModes")]
    public string[] AllowedModes { get; set; } = Array.Empty<string>();

    [JsonPropertyName("manualRenumber")]
    public bool ManualRenumber { get; set; }
}

public sealed class ProfileFieldPolicy
{
    [JsonPropertyName("preserve")]
    public bool Preserve { get; set; } = true;

    [JsonPropertyName("updateTocOnRequestOnly")]
    public bool UpdateTocOnRequestOnly { get; set; } = true;
}

public sealed class ProfileOutputPolicy
{
    [JsonPropertyName("overwriteOriginal")]
    public bool OverwriteOriginal { get; set; }

    [JsonPropertyName("backupRequired")]
    public bool BackupRequired { get; set; } = true;
}

public sealed class ProfileSemanticPolicy
{
    [JsonPropertyName("tableRoles")]
    public ProfileTableRoleSemanticPolicy TableRoles { get; set; } = new();

    [JsonPropertyName("invocation")]
    public ProfileInvocationSemanticPolicy Invocation { get; set; } = new();

    [JsonPropertyName("coverSubtitle")]
    public ProfileCoverSubtitleSemanticPolicy CoverSubtitle { get; set; } = new();

    [JsonPropertyName("figureTableCaption")]
    public ProfileFigureTableCaptionSemanticPolicy FigureTableCaption { get; set; } = new();

    [JsonPropertyName("appendix")]
    public ProfileAppendixSemanticPolicy Appendix { get; set; } = new();
}

public sealed class ProfileTableRoleSemanticPolicy
{
    [JsonPropertyName("firstRowRole")]
    public string? FirstRowRole { get; set; }

    [JsonPropertyName("bodyRole")]
    public string? BodyRole { get; set; }
}

public sealed class ProfileInvocationSemanticPolicy
{
    [JsonPropertyName("role")]
    public string? Role { get; set; }

    [JsonPropertyName("exactCompactTexts")]
    public string[] ExactCompactTexts { get; set; } = Array.Empty<string>();
}

public sealed class ProfileCoverSubtitleSemanticPolicy
{
    [JsonPropertyName("role")]
    public string? Role { get; set; }

    [JsonPropertyName("styleNames")]
    public string[] StyleNames { get; set; } = Array.Empty<string>();

    [JsonPropertyName("maxParagraphIndex")]
    public int? MaxParagraphIndex { get; set; }
}

public sealed class ProfileFigureTableCaptionSemanticPolicy
{
    [JsonPropertyName("role")]
    public string? Role { get; set; }

    [JsonPropertyName("prefixes")]
    public string[] Prefixes { get; set; } = Array.Empty<string>();

    [JsonPropertyName("requireParenthesizedTwoPartNumber")]
    public bool RequireParenthesizedTwoPartNumber { get; set; }
}

public sealed class ProfileAppendixSemanticPolicy
{
    [JsonPropertyName("headingPrefixes")]
    public string[] HeadingPrefixes { get; set; } = Array.Empty<string>();
}
public sealed class ProfileLayoutPolicy
{
    [JsonPropertyName("toc")]
    public ProfileTocLayoutPolicy Toc { get; set; } = new();

    [JsonPropertyName("appendixTables")]
    public ProfileAppendixTableLayoutPolicy AppendixTables { get; set; } = new();

    [JsonPropertyName("tablePagination")]
    public ProfileTablePaginationLayoutPolicy TablePagination { get; set; } = new();

    [JsonPropertyName("tableStyle")]
    public ProfileTableStyleLayoutPolicy TableStyle { get; set; } = new();
}

public sealed class ProfileTocLayoutPolicy
{
    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("compactStyles")]
    public bool CompactStyles { get; set; }

    [JsonPropertyName("titleTexts")]
    public string[] TitleTexts { get; set; } = Array.Empty<string>();

    [JsonPropertyName("spacing")]
    public ProfileTocSpacingLayoutPolicy Spacing { get; set; } = new();

    [JsonPropertyName("startBodyOnNewPage")]
    public bool StartBodyOnNewPage { get; set; }
}

public sealed class ProfileTocSpacingLayoutPolicy
{
    [JsonPropertyName("beforeTwips")]
    public int BeforeTwips { get; set; }

    [JsonPropertyName("afterTwips")]
    public int AfterTwips { get; set; }

    [JsonPropertyName("lineTwips")]
    public int LineTwips { get; set; }

    [JsonPropertyName("lineRule")]
    public string LineRule { get; set; } = string.Empty;
}

public sealed class ProfileAppendixTableLayoutPolicy
{
    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("fitToTextWidth")]
    public bool FitToTextWidth { get; set; }

    [JsonPropertyName("rtlVisual")]
    public bool RtlVisual { get; set; }

    [JsonPropertyName("tableAlignment")]
    public string TableAlignment { get; set; } = string.Empty;

    [JsonPropertyName("fixedLayout")]
    public bool FixedLayout { get; set; }

    [JsonPropertyName("normalizeCellMargins")]
    public bool NormalizeCellMargins { get; set; }

    [JsonPropertyName("cellMarginTwips")]
    public int CellMarginTwips { get; set; }

    [JsonPropertyName("wideDetection")]
    public ProfileWideTableDetectionLayoutPolicy WideDetection { get; set; } = new();

    [JsonPropertyName("landscapeWideTables")]
    public ProfileLandscapeWideTableLayoutPolicy LandscapeWideTables { get; set; } = new();
}

public sealed class ProfileWideTableDetectionLayoutPolicy
{
    [JsonPropertyName("minColumns")]
    public int MinColumns { get; set; }

    [JsonPropertyName("minAverageColumnWidthTwips")]
    public int MinAverageColumnWidthTwips { get; set; }
}

public sealed class ProfileLandscapeWideTableLayoutPolicy
{
    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("strategy")]
    public string Strategy { get; set; } = string.Empty;
}

public sealed class ProfileTablePaginationLayoutPolicy
{
    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("compactMaxRows")]
    public int CompactMaxRows { get; set; }

    [JsonPropertyName("compactScoreThreshold")]
    public int CompactScoreThreshold { get; set; }

    [JsonPropertyName("rowPenaltyStart")]
    public int RowPenaltyStart { get; set; }

    [JsonPropertyName("rowPenaltyWeight")]
    public int RowPenaltyWeight { get; set; }

    [JsonPropertyName("columnPenaltyStart")]
    public int ColumnPenaltyStart { get; set; }

    [JsonPropertyName("columnPenaltyWeight")]
    public int ColumnPenaltyWeight { get; set; }

    [JsonPropertyName("noRowSplit")]
    public bool NoRowSplit { get; set; }

    [JsonPropertyName("repeatHeader")]
    public bool RepeatHeader { get; set; }

    [JsonPropertyName("keepCaptionWithTable")]
    public bool KeepCaptionWithTable { get; set; }

    [JsonPropertyName("minBodyRowsAtStart")]
    public int MinBodyRowsAtStart { get; set; }

    [JsonPropertyName("minBodyRowsAtEnd")]
    public int MinBodyRowsAtEnd { get; set; }
}
public sealed class ProfileTableStyleLayoutPolicy
{
    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("theme")]
    public string Theme { get; set; } = "LIGHT_BLUE";

    [JsonPropertyName("bandRows")]
    public bool BandRows { get; set; } = true;

    [JsonPropertyName("centerTable")]
    public bool CenterTable { get; set; } = true;

    [JsonPropertyName("rtlVisual")]
    public bool RtlVisual { get; set; } = true;

    [JsonPropertyName("verticalAlignment")]
    public string VerticalAlignment { get; set; } = "CENTER";

    [JsonPropertyName("headerAlignment")]
    public string HeaderAlignment { get; set; } = "CENTER";

    [JsonPropertyName("shortCellAlignment")]
    public string ShortCellAlignment { get; set; } = "CENTER";

    [JsonPropertyName("longPersianAlignment")]
    public string LongPersianAlignment { get; set; } = "RIGHT";

    [JsonPropertyName("longTextThreshold")]
    public int LongTextThreshold { get; set; } = 44;

    [JsonPropertyName("borderSize")]
    public int BorderSize { get; set; } = 4;
}

public sealed class ProfileEditingPolicy
{
    [JsonPropertyName("mechanicalAuditEnabled")]
    public bool MechanicalAuditEnabled { get; set; } = true;

    [JsonPropertyName("mutationEnabled")]
    public bool MutationEnabled { get; set; }

    [JsonPropertyName("maxSamples")]
    public int MaxSamples { get; set; } = 200;
}

public sealed class ProfileDetectionPolicy
{
    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("priority")]
    public int Priority { get; set; }

    [JsonPropertyName("minScore")]
    public int MinScore { get; set; } = 20;
}



