using System.Text.RegularExpressions;

namespace WordAgent.WordInterop;

internal sealed record PersianSemanticStyleInfo(
    string StyleId,
    string Name,
    string OutlineLevel,
    bool TocEntry,
    bool TocTitle,
    bool Title,
    bool Chapter,
    bool Heading,
    bool Body,
    bool Caption);

internal sealed record PersianParagraphSemanticContext(
    string Text,
    PersianSemanticStyleInfo? Style,
    int ParagraphIndex,
    bool IsMainDocumentPart,
    bool InTable,
    bool InTextBox,
    bool HasVisual,
    bool IsFieldDriven,
    bool IsThesisProfile);

internal sealed record PersianParagraphSemanticDecision(
    bool TocEntry,
    bool TocTitle,
    bool Title,
    bool Chapter,
    bool Heading,
    bool Body,
    bool Caption,
    bool Label,
    bool FieldDriven,
    bool EarlyCoverLine,
    bool ExplicitThesisChapterTitle,
    bool CenteredSemantic,
    bool VisualRightSemantic,
    bool ShouldJustifyBody,
    string TargetFont);

internal static class PersianParagraphSemanticResolver
{
    private static readonly string[] HeadingPrefixes =
    [
        "گزارش", "چکیده", "مقدمه", "هدف", "اهداف", "دامنه", "روش", "روش‌شناسی",
        "روش شناسی", "یافته", "یافته‌ها", "یافته ها", "نتیجه", "نتایج", "نتیجه‌گیری",
        "نتیجه گیری", "بحث", "جمع‌بندی", "جمع بندی", "پیشنهاد", "پیشنهادها",
        "منابع", "مراجع", "پیوست", "فصل", "بخش", "خلاصه", "شرح"
    ];

    private static readonly string[] LabelPrefixes =
    [
        "تاریخ", "شماره", "کد", "نسخه", "وضعیت", "موضوع", "نام",
        "تهیه‌کننده", "تهیه کننده", "تأیید", "تایید", "مصوب", "پروژه"
    ];

    internal static PersianSemanticStyleInfo ClassifyStyle(
        string styleId,
        string name,
        string outlineLevel)
    {
        styleId = (styleId ?? string.Empty).Trim();
        name = (name ?? string.Empty).Trim();
        outlineLevel = (outlineLevel ?? string.Empty).Trim();

        var compact = Compact(styleId + "|" + name);
        var idToken = Compact(styleId);
        var nameToken = Compact(name);

        var tocTitle =
            compact.Contains("toctitle", StringComparison.Ordinal) ||
            compact.Contains("tocheading", StringComparison.Ordinal) ||
            compact.Contains("عنوانفهرست", StringComparison.Ordinal) ||
            compact.Contains("فهرستمطالب", StringComparison.Ordinal) ||
            compact.Contains("فهرستجداول", StringComparison.Ordinal) ||
            compact.Contains("فهرستاشکال", StringComparison.Ordinal) ||
            compact.Contains("فهرستشکلها", StringComparison.Ordinal) ||
            compact.Contains("فهرستنمودارها", StringComparison.Ordinal);

        var tocEntry =
            !tocTitle &&
            (Regex.IsMatch(compact, @"(?:^|\|)toc[123](?:\||$)") ||
             compact.Contains("toc1", StringComparison.Ordinal) ||
             compact.Contains("toc2", StringComparison.Ordinal) ||
             compact.Contains("toc3", StringComparison.Ordinal));

        // Chapter/Fasl is resolved before title. This permanently prevents styles such as
        // "Chapter Title" from being interpreted as a document/cover title merely because
        // the word "title" is present.
        var chapter =
            compact.Contains("chapter", StringComparison.Ordinal) ||
            compact.Contains("فصل", StringComparison.Ordinal);

        var explicitDocumentTitle =
            compact.Contains("documenttitle", StringComparison.Ordinal) ||
            compact.Contains("covertitle", StringComparison.Ordinal) ||
            compact.Contains("covertext", StringComparison.Ordinal) ||
            compact.Contains("عنوانسند", StringComparison.Ordinal) ||
            compact.Contains("عنوانپایان", StringComparison.Ordinal) ||
            compact.Contains("عنوانپايان", StringComparison.Ordinal) ||
            compact.Contains("متنرویجلد", StringComparison.Ordinal) ||
            compact.Contains("متنرويجلد", StringComparison.Ordinal) ||
            idToken is "title" or "subtitle" ||
            nameToken is "title" or "subtitle" or "عنوان";

        var title = !tocEntry && !tocTitle && !chapter && explicitDocumentTitle;

        var heading =
            chapter ||
            !string.IsNullOrWhiteSpace(outlineLevel) ||
            compact.Contains("heading", StringComparison.Ordinal) ||
            compact.Contains("تیتر", StringComparison.Ordinal) ||
            compact.Contains("تيتر", StringComparison.Ordinal);

        var body =
            compact.Contains("normal", StringComparison.Ordinal) ||
            compact.Contains("bodytext", StringComparison.Ordinal) ||
            compact.Contains("body", StringComparison.Ordinal) ||
            compact.Contains("متنبدنه", StringComparison.Ordinal) ||
            compact.EndsWith("|متن", StringComparison.Ordinal) ||
            compact.StartsWith("متن|", StringComparison.Ordinal);

        var caption =
            compact.Contains("caption", StringComparison.Ordinal) ||
            compact.Contains("زیرنویس", StringComparison.Ordinal) ||
            compact.Contains("زيرنويس", StringComparison.Ordinal) ||
            compact.Contains("بالانویس", StringComparison.Ordinal) ||
            compact.Contains("بالانويس", StringComparison.Ordinal) ||
            compact.Contains("شکلجدول", StringComparison.Ordinal);

        return new PersianSemanticStyleInfo(
            styleId, name, outlineLevel, tocEntry, tocTitle, title, chapter, heading, body, caption);
    }

    internal static PersianParagraphSemanticDecision Resolve(PersianParagraphSemanticContext context)
    {
        var text = context.Text ?? string.Empty;
        var style = context.Style;

        var tocTitle = !context.IsFieldDriven && (style?.TocTitle == true || IsExactTocTitle(text));
        var tocEntry = style?.TocEntry == true;
        var title = style?.Title == true;
        var chapter = style?.Chapter == true;
        var heading =
            style?.Heading == true ||
            LooksLikeStrongHeadingText(text) ||
            (style is null &&
             !context.InTable &&
             !context.InTextBox &&
             LooksLikeHeadingText(text));
        var caption = style?.Caption == true || LooksLikeCaptionText(text);
        var label = LooksLikeStandaloneLabel(text);

        var explicitThesisChapterTitle =
            context.IsThesisProfile &&
            LooksLikeExplicitChapterTitle(text);

        var earlyCoverLine =
            context.IsMainDocumentPart &&
            !context.InTable &&
            !context.InTextBox &&
            LooksLikeEarlyCoverCenteredLine(text, style, context.ParagraphIndex);

        // Physical containers outrank textual cover/title heuristics.
        var centeredSemantic =
            !context.InTable &&
            !context.InTextBox &&
            (context.HasVisual ||
             tocTitle ||
             title ||
             earlyCoverLine ||
             explicitThesisChapterTitle);

        var visualRightSemantic =
            !centeredSemantic &&
            (context.InTable ||
             context.InTextBox ||
             tocEntry ||
             heading ||
             chapter);

        var shouldJustifyBody =
            context.IsMainDocumentPart &&
            !centeredSemantic &&
            !visualRightSemantic &&
            !context.InTable &&
            !context.InTextBox &&
            !context.IsFieldDriven &&
            !caption &&
            !label;

        var targetFont =
            (tocTitle || tocEntry)
                ? "B Nazanin"
                : (title || chapter || heading ? "B Titr" : "B Nazanin");

        return new PersianParagraphSemanticDecision(
            tocEntry,
            tocTitle,
            title,
            chapter,
            heading,
            style?.Body == true,
            caption,
            label,
            context.IsFieldDriven,
            earlyCoverLine,
            explicitThesisChapterTitle,
            centeredSemantic,
            visualRightSemantic,
            shouldJustifyBody,
            targetFont);
    }

    internal static bool IsThesisOrDissertationProfile(string? profileId)
    {
        if (string.IsNullOrWhiteSpace(profileId))
            return false;

        var normalized = profileId.Trim().ToLowerInvariant();
        return normalized.Contains("thesis", StringComparison.Ordinal) ||
               normalized.Contains("dissertation", StringComparison.Ordinal);
    }

    internal static bool LooksLikeExplicitChapterTitle(string text)
    {
        var trimmed = (text ?? string.Empty).Trim();
        if (trimmed.Length == 0 || trimmed.Length > 160)
            return false;

        var words = trimmed.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (words.Length > 16)
            return false;

        return Regex.IsMatch(
            trimmed,
            @"^[\s\u200c\u200f]*فصل\s*(?:[0-9۰-۹]+|اول|دوم|سوم|چهارم|پنجم|ششم|هفتم|هشتم|نهم|دهم)\s*(?:(?:[:：\-–—]\s*.+)?|)$");
    }

    internal static bool IsExactTocTitle(string text)
    {
        var compact = Compact(text);
        return compact.Equals("فهرستمطالب", StringComparison.Ordinal) ||
               compact.Equals("فهرستجداول", StringComparison.Ordinal) ||
               compact.Equals("فهرستاشکال", StringComparison.Ordinal) ||
               compact.Equals("فهرستشکلها", StringComparison.Ordinal) ||
               compact.Equals("فهرستنمودارها", StringComparison.Ordinal);
    }

    internal static bool LooksLikeStrongHeadingText(string text)
    {
        var trimmed = (text ?? string.Empty).Trim();
        if (trimmed.Length == 0 || LooksLikeCaptionText(trimmed))
            return false;

        var words = trimmed.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0 || words.Length > 20)
            return false;

        if (Regex.IsMatch(
                trimmed,
                @"^[\s\u200c\u200f]*(?:[0-9۰-۹]+(?:\s*[\.\-–—:]\s*|(?:\.[0-9۰-۹]+)+\s+))"))
        {
            return true;
        }

        var endsLikeSentence = EndsLikeSentence(trimmed);

        return !endsLikeSentence &&
               HeadingPrefixes.Any(prefix => trimmed.StartsWith(prefix, StringComparison.Ordinal));
    }

    internal static bool LooksLikeHeadingText(string text)
    {
        var trimmed = (text ?? string.Empty).Trim();
        if (trimmed.Length == 0 || LooksLikeCaptionText(trimmed))
            return false;

        var words = trimmed.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0)
            return false;

        if (Regex.IsMatch(
                trimmed,
                @"^[\s\u200c\u200f]*(?:[0-9۰-۹]+(?:\s*[\.\-–—:]\s*|(?:\.[0-9۰-۹]+)+\s+))"))
        {
            return words.Length <= 20;
        }

        if (words.Length <= 18 &&
            HeadingPrefixes.Any(prefix => trimmed.StartsWith(prefix, StringComparison.Ordinal)))
        {
            return true;
        }

        return trimmed.Length <= 90 &&
               words.Length <= 10 &&
               !EndsLikeSentence(trimmed) &&
               !trimmed.EndsWith(":", StringComparison.Ordinal);
    }

    internal static bool LooksLikeStandaloneLabel(string text)
    {
        var trimmed = (text ?? string.Empty).Trim();
        var words = trimmed.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

        if (trimmed.Length <= 55 &&
            words.Length <= 6 &&
            (trimmed.EndsWith(":", StringComparison.Ordinal) ||
             trimmed.EndsWith("：", StringComparison.Ordinal)))
        {
            return true;
        }

        return trimmed.Length <= 65 &&
               words.Length <= 7 &&
               LabelPrefixes.Any(prefix => trimmed.StartsWith(prefix, StringComparison.Ordinal));
    }

    internal static bool LooksLikeCaptionText(string text)
    {
        var trimmed = (text ?? string.Empty).Trim();
        if (trimmed.Length == 0)
            return false;

        return Regex.IsMatch(
            trimmed,
            @"^(?:جدول|شکل|نمودار)\s*(?:شماره\s*)?[\(\[]?[0-9۰-۹]+[\)\]]?\s*(?:[-–—:：\.]\s*|$)");
    }

    private static bool LooksLikeEarlyCoverCenteredLine(
        string text,
        PersianSemanticStyleInfo? style,
        int paragraphIndex)
    {
        if (paragraphIndex > 12)
            return false;

        var trimmed = (text ?? string.Empty).Trim();
        if (trimmed.Length == 0 ||
            LooksLikeCaptionText(trimmed) ||
            LooksLikeStandaloneLabel(trimmed) ||
            IsExactTocTitle(trimmed))
        {
            return false;
        }

        if (Regex.IsMatch(
                trimmed,
                @"^[\s\u200c\u200f]*[0-9۰-۹]+(?:\s*[\.\-–—:]\s*|(?:\.[0-9۰-۹]+)+\s+)"))
        {
            return false;
        }

        var styleCompact = Compact((style?.StyleId ?? string.Empty) + "|" + (style?.Name ?? string.Empty));
        if (styleCompact.Contains("cover", StringComparison.Ordinal) ||
            styleCompact.Contains("متنرویجلد", StringComparison.Ordinal) ||
            styleCompact.Contains("متنرويجلد", StringComparison.Ordinal))
        {
            return true;
        }

        var textCompact = Compact(trimmed);
        if (textCompact.Equals("بسمهتعالی", StringComparison.Ordinal) ||
            textCompact.Equals("بسمهتعالي", StringComparison.Ordinal))
        {
            return true;
        }

        if (style is not null &&
            (style.Body ||
             style.Caption ||
             style.TocEntry ||
             style.TocTitle ||
             style.Chapter ||
             style.Heading))
        {
            return false;
        }

        var words = trimmed.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        return trimmed.Length <= 240 &&
               words.Length <= 30 &&
               !EndsLikeSentence(trimmed);
    }

    private static bool EndsLikeSentence(string text)
        => text.EndsWith(".", StringComparison.Ordinal) ||
           text.EndsWith("؟", StringComparison.Ordinal) ||
           text.EndsWith("!", StringComparison.Ordinal) ||
           text.EndsWith("؛", StringComparison.Ordinal) ||
           text.EndsWith("،", StringComparison.Ordinal);

    private static string Compact(string value)
        => (value ?? string.Empty)
            .Trim()
            .Replace(" ", string.Empty, StringComparison.Ordinal)
            .Replace("-", string.Empty, StringComparison.Ordinal)
            .Replace("_", string.Empty, StringComparison.Ordinal)
            .Replace("\u200c", string.Empty, StringComparison.Ordinal)
            .Replace("\u200f", string.Empty, StringComparison.Ordinal)
            .ToLowerInvariant();
}
