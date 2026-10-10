namespace WordAgent.Core;

public static class LayoutPolicyDecisions
{
    public static bool IsWideTable(
        int columns,
        int gridWidthTwips,
        ProfileWideTableDetectionLayoutPolicy policy)
    {
        ArgumentNullException.ThrowIfNull(policy);

        var average = columns > 0 && gridWidthTwips > 0
            ? gridWidthTwips / columns
            : int.MaxValue;

        return columns >= policy.MinColumns ||
               average < policy.MinAverageColumnWidthTwips;
    }

    public static int ComputePaginationScore(
        int rowCount,
        int columns,
        int totalChars,
        ProfileTablePaginationLayoutPolicy policy)
    {
        ArgumentNullException.ThrowIfNull(policy);

        return totalChars
            + Math.Max(0, rowCount - policy.RowPenaltyStart) * policy.RowPenaltyWeight
            + Math.Max(0, columns - policy.ColumnPenaltyStart) * policy.ColumnPenaltyWeight;
    }

    public static bool ShouldKeepTableWhole(
        int rowCount,
        int columns,
        int totalChars,
        ProfileTablePaginationLayoutPolicy policy)
    {
        ArgumentNullException.ThrowIfNull(policy);

        var score = ComputePaginationScore(rowCount, columns, totalChars, policy);
        return rowCount <= policy.CompactMaxRows &&
               score <= policy.CompactScoreThreshold;
    }
}