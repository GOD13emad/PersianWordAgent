namespace WordAgent.Core;

public static class SemanticPolicyMatcher
{
    public static string? ResolveTableRole(int? tableRow, ProfileTableRoleSemanticPolicy? policy)
    {
        if (policy is null) return null;

        if (tableRow == 1 && !string.IsNullOrWhiteSpace(policy.FirstRowRole))
            return policy.FirstRowRole.Trim();

        return string.IsNullOrWhiteSpace(policy.BodyRole)
            ? null
            : policy.BodyRole.Trim();
    }

    public static bool MatchesInvocation(string text, ProfileInvocationSemanticPolicy? policy)
    {
        if (policy is null || policy.ExactCompactTexts is null || policy.ExactCompactTexts.Length == 0)
            return false;

        var compact = CompactInvocation(text);
        if (compact.Length == 0) return false;

        return policy.ExactCompactTexts
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(CompactInvocation)
            .Any(x => string.Equals(x, compact, StringComparison.Ordinal));
    }

    public static bool MatchesCoverSubtitle(
        string? styleName,
        int paragraphIndex,
        ProfileCoverSubtitleSemanticPolicy? policy)
    {
        if (policy is null || string.IsNullOrWhiteSpace(styleName) || policy.StyleNames is null)
            return false;

        if (policy.MaxParagraphIndex is int max && paragraphIndex > max)
            return false;

        return policy.StyleNames.Any(x =>
            !string.IsNullOrWhiteSpace(x) &&
            string.Equals(x.Trim(), styleName.Trim(), StringComparison.OrdinalIgnoreCase));
    }

    public static bool MatchesFigureTableCaption(
        string text,
        ProfileFigureTableCaptionSemanticPolicy? policy)
    {
        if (policy is null || policy.Prefixes is null || policy.Prefixes.Length == 0)
            return false;

        var value = NormalizeVisibleText(text);
        if (value.Length == 0) return false;

        var prefixMatched = policy.Prefixes.Any(prefix =>
            !string.IsNullOrWhiteSpace(prefix) &&
            value.StartsWith(prefix.Trim(), StringComparison.OrdinalIgnoreCase));

        if (!prefixMatched) return false;
        if (!policy.RequireParenthesizedTwoPartNumber) return true;

        var open = value.IndexOf('(');
        var close = open >= 0 ? value.IndexOf(')', open + 1) : -1;
        if (open < 0 || close <= open + 1) return false;

        var token = value[(open + 1)..close]
            .Replace(" ", string.Empty, StringComparison.Ordinal)
            .Replace("ـ", string.Empty, StringComparison.Ordinal);

        var parts = token.Split(
            new[] { '-', '–', '—' },
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        return parts.Length == 2 && parts.All(IsDigitToken);
    }

    public static bool MatchesAppendixHeading(string text, ProfileAppendixSemanticPolicy? policy)
    {
        if (policy is null || policy.HeadingPrefixes is null || policy.HeadingPrefixes.Length == 0)
            return false;

        var value = NormalizeVisibleText(text);
        return policy.HeadingPrefixes.Any(prefix =>
            !string.IsNullOrWhiteSpace(prefix) &&
            value.StartsWith(prefix.Trim(), StringComparison.Ordinal));
    }

    private static string CompactInvocation(string? value)
        => (value ?? string.Empty)
            .Replace(" ", string.Empty, StringComparison.Ordinal)
            .Replace("\u200c", string.Empty, StringComparison.Ordinal)
            .Trim();

    private static string NormalizeVisibleText(string? value)
        => (value ?? string.Empty)
            .Trim()
            .Replace('\u200c', ' ')
            .Replace('\u200f', ' ')
            .Replace('\u202b', ' ');

    private static bool IsDigitToken(string value)
        => !string.IsNullOrWhiteSpace(value) && value.All(char.IsDigit);
}
