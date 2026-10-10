using System.Collections.Generic;
using System.Linq;

namespace WordAgent.Core;

public static class ProfileSchema
{
    public const string LegacyCompatibility = "1.0-legacy";
    public const string Current = "2.0";
}

public sealed record ProfileValidationIssue(
    string Severity,
    string Code,
    string Message);

public sealed record ProfileValidationResult(
    string ProfileId,
    string SchemaVersion,
    IReadOnlyList<ProfileValidationIssue> Issues)
{
    public int ErrorCount => Issues.Count(x => x.Severity.Equals("ERROR", StringComparison.OrdinalIgnoreCase));
    public int WarningCount => Issues.Count(x => x.Severity.Equals("WARNING", StringComparison.OrdinalIgnoreCase));
    public bool IsValid => ErrorCount == 0;
}

public static class ProfileValidator
{
    private static readonly HashSet<string> AllowedReadingOrders =
        new(StringComparer.OrdinalIgnoreCase) { "RTL", "LTR" };

    private static readonly HashSet<string> AllowedAlignments =
        new(StringComparer.OrdinalIgnoreCase) { "RIGHT", "LEFT", "CENTER", "JUSTIFY", "DISTRIBUTE" };

    private static readonly HashSet<string> AllowedTableThemes =
        new(StringComparer.OrdinalIgnoreCase) { "LIGHT_BLUE", "BLUE", "GREEN" };

    private static readonly HashSet<string> AllowedVerticalAlignments =
        new(StringComparer.OrdinalIgnoreCase) { "TOP", "CENTER", "BOTTOM" };

    public static ProfileValidationResult Validate(AgentProfile profile)
    {
        ArgumentNullException.ThrowIfNull(profile);

        var issues = new List<ProfileValidationIssue>();

        void Error(string code, string message) =>
            issues.Add(new ProfileValidationIssue("ERROR", code, message));

        void Warn(string code, string message) =>
            issues.Add(new ProfileValidationIssue("WARNING", code, message));

        var schema = string.IsNullOrWhiteSpace(profile.SchemaVersion)
            ? ProfileSchema.LegacyCompatibility
            : profile.SchemaVersion.Trim();

        if (schema.Equals(ProfileSchema.LegacyCompatibility, StringComparison.OrdinalIgnoreCase))
        {
            Warn("PROFILE-SCHEMA-LEGACY-001",
                "Profile has no explicit schemaVersion and is being loaded through the legacy-compatibility path.");
        }
        else if (!schema.Equals(ProfileSchema.Current, StringComparison.OrdinalIgnoreCase))
        {
            Error("PROFILE-SCHEMA-UNSUPPORTED-001",
                $"Unsupported schemaVersion '{schema}'. Supported: {ProfileSchema.Current} or legacy compatibility.");
        }

        if (string.IsNullOrWhiteSpace(profile.Id))
            Error("PROFILE-ID-001", "Profile id is required.");

        if (string.IsNullOrWhiteSpace(profile.DisplayName))
            Error("PROFILE-DISPLAYNAME-001", "Profile displayName is required.");

        if (string.IsNullOrWhiteSpace(profile.Version))
            Error("PROFILE-VERSION-001", "Profile version is required.");

        if (string.IsNullOrWhiteSpace(profile.DocumentClass))
            Error("PROFILE-CLASS-001", "Profile documentClass is required.");

        if (string.IsNullOrWhiteSpace(profile.Language))
            Error("PROFILE-LANGUAGE-001", "Profile language is required.");

        if (profile.Classification is null)
        {
            Error("PROFILE-CLASSIFICATION-001", "classification object must not be null.");
        }
        else
        {
            ValidateStringArray(profile.Classification.Keywords, "classification.keywords", Error);
            ValidateStringArray(profile.Classification.NegativeKeywords, "classification.negativeKeywords", Error);
            ValidateStringArray(profile.Classification.StyleSignals, "classification.styleSignals", Error);
        }

        if (profile.RoleAliases is null)
        {
            Error("PROFILE-ALIASES-001", "roleAliases must not be null.");
        }
        else
        {
            var aliasOwners = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var pair in profile.RoleAliases)
            {
                var role = pair.Key?.Trim() ?? string.Empty;
                if (role.Length == 0)
                {
                    Error("PROFILE-ROLE-NAME-001", "roleAliases contains an empty role name.");
                    continue;
                }

                if (pair.Value is null)
                {
                    Error("PROFILE-ROLE-ALIASES-NULL-001", $"roleAliases['{role}'] must not be null.");
                    continue;
                }

                foreach (var rawAlias in pair.Value)
                {
                    var alias = rawAlias?.Trim() ?? string.Empty;
                    if (alias.Length == 0)
                    {
                        Error("PROFILE-ALIAS-EMPTY-001", $"Role '{role}' contains an empty style alias.");
                        continue;
                    }

                    if (aliasOwners.TryGetValue(alias, out var owner) &&
                        !owner.Equals(role, StringComparison.OrdinalIgnoreCase))
                    {
                        Error("PROFILE-ALIAS-AMBIGUOUS-001",
                            $"Style alias '{alias}' is assigned to both '{owner}' and '{role}'.");
                    }
                    else
                    {
                        aliasOwners[alias] = role;
                    }
                }
            }
        }

        if (profile.RoleDefaults is null)
        {
            Error("PROFILE-DEFAULTS-001", "roleDefaults must not be null.");
        }
        else
        {
            foreach (var pair in profile.RoleDefaults)
            {
                var role = pair.Key?.Trim() ?? string.Empty;
                if (role.Length == 0)
                {
                    Error("PROFILE-DEFAULT-ROLE-001", "roleDefaults contains an empty role name.");
                    continue;
                }

                var defaults = pair.Value;
                if (defaults is null)
                {
                    Error("PROFILE-DEFAULT-NULL-001", $"roleDefaults['{role}'] must not be null.");
                    continue;
                }

                if (defaults.SizeBi is <= 0)
                    Error("PROFILE-SIZE-001", $"Role '{role}' has non-positive sizeBi.");

                if (!string.IsNullOrWhiteSpace(defaults.ReadingOrder) &&
                    !AllowedReadingOrders.Contains(defaults.ReadingOrder.Trim()))
                {
                    Error("PROFILE-READINGORDER-001",
                        $"Role '{role}' has unsupported readingOrder '{defaults.ReadingOrder}'.");
                }

                if (!string.IsNullOrWhiteSpace(defaults.Alignment) &&
                    !AllowedAlignments.Contains(defaults.Alignment.Trim()))
                {
                    Error("PROFILE-ALIGNMENT-001",
                        $"Role '{role}' has unsupported alignment '{defaults.Alignment}'.");
                }
            }
        }

        if (profile.SemanticPolicy is null)
        {
            Error("PROFILE-SEMANTIC-POLICY-001", "semanticPolicy must not be null.");
        }
        else
        {
            var semantic = profile.SemanticPolicy;

            void ValidateSemanticRoleReference(string? role, string label, bool cueConfigured)
            {
                var value = role?.Trim() ?? string.Empty;

                if (value.Length == 0)
                {
                    if (cueConfigured)
                        Error("PROFILE-SEMANTIC-ROLE-001", $"{label}.role is required when semantic cues are configured.");
                    return;
                }

                if (profile.RoleDefaults?.ContainsKey(value) != true)
                    Error("PROFILE-SEMANTIC-ROLE-REF-001", $"{label} references unknown roleDefaults role '{value}'.");
            }

            if (semantic.TableRoles is null)
            {
                Error("PROFILE-SEMANTIC-TABLE-001", "semanticPolicy.tableRoles must not be null.");
            }
            else
            {
                ValidateSemanticRoleReference(
                    semantic.TableRoles.FirstRowRole,
                    "semanticPolicy.tableRoles.firstRowRole",
                    !string.IsNullOrWhiteSpace(semantic.TableRoles.FirstRowRole));

                ValidateSemanticRoleReference(
                    semantic.TableRoles.BodyRole,
                    "semanticPolicy.tableRoles.bodyRole",
                    !string.IsNullOrWhiteSpace(semantic.TableRoles.BodyRole));
            }

            if (semantic.Invocation is null)
            {
                Error("PROFILE-SEMANTIC-INVOCATION-001", "semanticPolicy.invocation must not be null.");
            }
            else
            {
                ValidateStringArray(
                    semantic.Invocation.ExactCompactTexts,
                    "semanticPolicy.invocation.exactCompactTexts",
                    Error);

                var configured = semantic.Invocation.ExactCompactTexts?.Any(x => !string.IsNullOrWhiteSpace(x)) == true;
                ValidateSemanticRoleReference(
                    semantic.Invocation.Role,
                    "semanticPolicy.invocation",
                    configured);
            }

            if (semantic.CoverSubtitle is null)
            {
                Error("PROFILE-SEMANTIC-COVER-001", "semanticPolicy.coverSubtitle must not be null.");
            }
            else
            {
                ValidateStringArray(
                    semantic.CoverSubtitle.StyleNames,
                    "semanticPolicy.coverSubtitle.styleNames",
                    Error);

                if (semantic.CoverSubtitle.MaxParagraphIndex is <= 0)
                    Error("PROFILE-SEMANTIC-COVER-RANGE-001", "semanticPolicy.coverSubtitle.maxParagraphIndex must be > 0 when specified.");

                var configured = semantic.CoverSubtitle.StyleNames?.Any(x => !string.IsNullOrWhiteSpace(x)) == true;
                ValidateSemanticRoleReference(
                    semantic.CoverSubtitle.Role,
                    "semanticPolicy.coverSubtitle",
                    configured);
            }

            if (semantic.FigureTableCaption is null)
            {
                Error("PROFILE-SEMANTIC-CAPTION-001", "semanticPolicy.figureTableCaption must not be null.");
            }
            else
            {
                ValidateStringArray(
                    semantic.FigureTableCaption.Prefixes,
                    "semanticPolicy.figureTableCaption.prefixes",
                    Error);

                var configured = semantic.FigureTableCaption.Prefixes?.Any(x => !string.IsNullOrWhiteSpace(x)) == true;
                ValidateSemanticRoleReference(
                    semantic.FigureTableCaption.Role,
                    "semanticPolicy.figureTableCaption",
                    configured);
            }

            if (semantic.Appendix is null)
            {
                Error("PROFILE-SEMANTIC-APPENDIX-001", "semanticPolicy.appendix must not be null.");
            }
            else
            {
                ValidateStringArray(
                    semantic.Appendix.HeadingPrefixes,
                    "semanticPolicy.appendix.headingPrefixes",
                    Error);
            }
        }
        if (profile.LayoutPolicy is null)
        {
            Error("PROFILE-LAYOUT-POLICY-001", "layoutPolicy must not be null.");
        }
        else
        {
            var layout = profile.LayoutPolicy;

            if (layout.Toc is null)
            {
                Error("PROFILE-LAYOUT-TOC-001", "layoutPolicy.toc must not be null.");
            }
            else
            {
                ValidateStringArray(layout.Toc.TitleTexts, "layoutPolicy.toc.titleTexts", Error);

                if (layout.Toc.Spacing is null)
                {
                    Error("PROFILE-LAYOUT-TOC-SPACING-001", "layoutPolicy.toc.spacing must not be null.");
                }
                else if (layout.Toc.Enabled && layout.Toc.CompactStyles)
                {
                    if (layout.Toc.Spacing.BeforeTwips < 0 || layout.Toc.Spacing.AfterTwips < 0)
                        Error("PROFILE-LAYOUT-TOC-SPACING-RANGE-001", "TOC spacing before/after twips must be >= 0.");

                    if (layout.Toc.Spacing.LineTwips <= 0)
                        Error("PROFILE-LAYOUT-TOC-LINE-001", "TOC spacing lineTwips must be > 0.");

                    var lineRule = layout.Toc.Spacing.LineRule?.Trim() ?? string.Empty;
                    if (lineRule.Length == 0 ||
                        !(lineRule.Equals("auto", StringComparison.OrdinalIgnoreCase) ||
                          lineRule.Equals("exact", StringComparison.OrdinalIgnoreCase) ||
                          lineRule.Equals("atLeast", StringComparison.OrdinalIgnoreCase)))
                    {
                        Error("PROFILE-LAYOUT-TOC-LINERULE-001", $"Unsupported TOC lineRule '{layout.Toc.Spacing.LineRule}'.");
                    }
                }
            }

            if (layout.AppendixTables is null)
            {
                Error("PROFILE-LAYOUT-APPENDIX-001", "layoutPolicy.appendixTables must not be null.");
            }
            else
            {
                var appendix = layout.AppendixTables;

                if (appendix.WideDetection is null)
                    Error("PROFILE-LAYOUT-WIDE-001", "layoutPolicy.appendixTables.wideDetection must not be null.");

                if (appendix.LandscapeWideTables is null)
                    Error("PROFILE-LAYOUT-LANDSCAPE-001", "layoutPolicy.appendixTables.landscapeWideTables must not be null.");

                if (appendix.Enabled)
                {
                    var alignment = appendix.TableAlignment?.Trim() ?? string.Empty;
                    if (alignment.Length == 0 ||
                        !(alignment.Equals("LEFT", StringComparison.OrdinalIgnoreCase) ||
                          alignment.Equals("CENTER", StringComparison.OrdinalIgnoreCase) ||
                          alignment.Equals("RIGHT", StringComparison.OrdinalIgnoreCase)))
                    {
                        Error("PROFILE-LAYOUT-TABLE-ALIGNMENT-001", $"Unsupported appendix tableAlignment '{appendix.TableAlignment}'.");
                    }

                    if (appendix.NormalizeCellMargins && appendix.CellMarginTwips < 0)
                        Error("PROFILE-LAYOUT-CELL-MARGIN-001", "layoutPolicy.appendixTables.cellMarginTwips must be >= 0.");

                    if (appendix.LandscapeWideTables?.Enabled == true)
                    {
                        if (appendix.WideDetection is null ||
                            appendix.WideDetection.MinColumns < 1 ||
                            appendix.WideDetection.MinAverageColumnWidthTwips < 1)
                        {
                            Error("PROFILE-LAYOUT-WIDE-RANGE-001", "Wide-table thresholds must be >= 1 when landscape wide tables are enabled.");
                        }

                        if (!string.Equals(
                                appendix.LandscapeWideTables.Strategy?.Trim(),
                                "FINAL_WIDE_TAIL_ONLY",
                                StringComparison.OrdinalIgnoreCase))
                        {
                            Error("PROFILE-LAYOUT-LANDSCAPE-STRATEGY-001",
                                "Only FINAL_WIDE_TAIL_ONLY is supported in P2.5.");
                        }
                    }
                }
            }

            if (layout.TablePagination is null)
            {
                Error("PROFILE-LAYOUT-PAGINATION-001", "layoutPolicy.tablePagination must not be null.");
            }
            else if (layout.TablePagination.Enabled)
            {
                var pagination = layout.TablePagination;

                if (pagination.CompactMaxRows < 1)
                    Error("PROFILE-LAYOUT-PAGINATION-ROWS-001", "compactMaxRows must be >= 1.");

                if (pagination.CompactScoreThreshold < 1)
                    Error("PROFILE-LAYOUT-PAGINATION-SCORE-001", "compactScoreThreshold must be >= 1.");

                if (pagination.RowPenaltyStart < 0 ||
                    pagination.RowPenaltyWeight < 0 ||
                    pagination.ColumnPenaltyStart < 0 ||
                    pagination.ColumnPenaltyWeight < 0)
                {
                    Error("PROFILE-LAYOUT-PAGINATION-PENALTY-001", "Pagination penalty starts/weights must be >= 0.");
                }

                if (pagination.MinBodyRowsAtStart < 0 || pagination.MinBodyRowsAtEnd < 0)
                    Error("PROFILE-LAYOUT-PAGINATION-FRAGMENT-001", "Minimum body-row fragment values must be >= 0.");
            }
        }

        if (profile.LayoutPolicy?.TableStyle is null)
        {
            Error("PROFILE-LAYOUT-TABLESTYLE-001", "layoutPolicy.tableStyle must not be null.");
        }
        else if (profile.LayoutPolicy.TableStyle.Enabled)
        {
            var style = profile.LayoutPolicy.TableStyle;
            if (!AllowedTableThemes.Contains(style.Theme?.Trim() ?? string.Empty))
                Error("PROFILE-LAYOUT-TABLESTYLE-THEME-001", $"Unsupported table theme '{style.Theme}'.");

            foreach (var item in new[]
            {
                ("headerAlignment", style.HeaderAlignment),
                ("shortCellAlignment", style.ShortCellAlignment),
                ("longPersianAlignment", style.LongPersianAlignment)
            })
            {
                if (!AllowedAlignments.Contains(item.Item2?.Trim() ?? string.Empty))
                    Error("PROFILE-LAYOUT-TABLESTYLE-ALIGN-001", $"Unsupported tableStyle.{item.Item1} '{item.Item2}'.");
            }

            if (!AllowedVerticalAlignments.Contains(style.VerticalAlignment?.Trim() ?? string.Empty))
                Error("PROFILE-LAYOUT-TABLESTYLE-VALIGN-001", $"Unsupported tableStyle.verticalAlignment '{style.VerticalAlignment}'.");

            if (style.LongTextThreshold < 8)
                Error("PROFILE-LAYOUT-TABLESTYLE-THRESHOLD-001", "tableStyle.longTextThreshold must be >= 8.");

            if (style.BorderSize < 0 || style.BorderSize > 96)
                Error("PROFILE-LAYOUT-TABLESTYLE-BORDER-001", "tableStyle.borderSize must be between 0 and 96.");
        }

        if (profile.EditingPolicy is null)
        {
            Error("PROFILE-EDITING-POLICY-001", "editingPolicy must not be null.");
        }
        else
        {
            if (profile.EditingPolicy.MutationEnabled)
                Error("PROFILE-EDITING-MUTATION-001", "editingPolicy.mutationEnabled=true is not supported in F01; editing remains audit-only.");

            if (profile.EditingPolicy.MaxSamples < 1 || profile.EditingPolicy.MaxSamples > 5000)
                Error("PROFILE-EDITING-SAMPLES-001", "editingPolicy.maxSamples must be between 1 and 5000.");
        }

        ValidateStringArray(profile.Rules, "rules", Error);

        if (profile.ReferencePolicy is null)
        {
            Error("PROFILE-REFERENCE-POLICY-001", "referencePolicy must not be null.");
        }
        else
        {
            if (string.IsNullOrWhiteSpace(profile.ReferencePolicy.Owner))
                Error("PROFILE-REFERENCE-OWNER-001", "referencePolicy.owner is required.");

            if (profile.ReferencePolicy.ManualRenumber)
                Error("PROFILE-ENDNOTE-RENUMBER-001",
                    "manualRenumber=true is forbidden; citation numbering must remain reference-manager owned.");

            ValidateStringArray(profile.ReferencePolicy.AllowedModes, "referencePolicy.allowedModes", Error);
        }

        if (profile.FieldPolicy is null)
        {
            Error("PROFILE-FIELD-POLICY-001", "fieldPolicy must not be null.");
        }
        else if (!profile.FieldPolicy.Preserve)
        {
            Error("PROFILE-FIELD-PRESERVE-001",
                "fieldPolicy.preserve=false is forbidden for controlled WordAgent profiles.");
        }

        if (profile.OutputPolicy is null)
        {
            Error("PROFILE-OUTPUT-POLICY-001", "outputPolicy must not be null.");
        }
        else
        {
            if (profile.OutputPolicy.OverwriteOriginal)
                Error("PROFILE-OVERWRITE-001",
                    "outputPolicy.overwriteOriginal=true is forbidden.");

            if (!profile.OutputPolicy.BackupRequired)
                Error("PROFILE-BACKUP-001",
                    "outputPolicy.backupRequired must remain true.");
        }

        if (profile.DetectionPolicy is null)
        {
            Error("PROFILE-DETECTION-POLICY-001", "detectionPolicy must not be null.");
        }
        else
        {
            if (profile.DetectionPolicy.MinScore < 1)
                Error("PROFILE-DETECTION-MINSCORE-001", "detectionPolicy.minScore must be >= 1.");

            if (profile.DetectionPolicy.Enabled)
            {
                var keywordCount = profile.Classification?.Keywords?.Count(x => !string.IsNullOrWhiteSpace(x)) ?? 0;
                var styleCount = profile.Classification?.StyleSignals?.Count(x => !string.IsNullOrWhiteSpace(x)) ?? 0;

                if (keywordCount + styleCount == 0)
                    Error("PROFILE-DETECTION-SIGNALS-001",
                        "Detection-enabled profiles must define at least one keyword or style signal.");
            }
        }

        return new ProfileValidationResult(
            profile.Id ?? string.Empty,
            schema,
            issues);
    }

    private static void ValidateStringArray(
        string[]? values,
        string label,
        Action<string, string> error)
    {
        if (values is null)
        {
            error("PROFILE-ARRAY-NULL-001", $"{label} must not be null.");
            return;
        }

        for (var i = 0; i < values.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(values[i]))
                error("PROFILE-ARRAY-EMPTY-001", $"{label}[{i}] must not be empty.");
        }
    }
}


