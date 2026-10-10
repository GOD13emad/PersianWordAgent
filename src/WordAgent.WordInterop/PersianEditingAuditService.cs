using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using WordAgent.Core;

namespace WordAgent.WordInterop;

public sealed record PersianEditingIssue(
    int ParagraphIndex,
    string Code,
    string Severity,
    string TextPreview,
    string Suggestion);

public sealed record PersianEditingAuditResult(
    string Schema,
    string InputFile,
    string InputSha256,
    int ParagraphsExamined,
    int GuardedParagraphs,
    int IssueCount,
    int ArabicYehKafParagraphs,
    int RepeatedSpaceParagraphs,
    int SpaceBeforePunctuationParagraphs,
    int MissingSpaceAfterPunctuationParagraphs,
    int SeparatedMiPrefixParagraphs,
    int RepeatedZwnjParagraphs,
    bool MutationPerformed,
    IReadOnlyList<PersianEditingIssue> Issues);

public sealed class PersianEditingAuditService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static readonly Regex RepeatedSpace = new(@"[ \t]{2,}", RegexOptions.Compiled);
    private static readonly Regex SpaceBeforePunctuation = new(@"\s+[،؛؟!,:؛]", RegexOptions.Compiled);
    private static readonly Regex MissingSpaceAfterPunctuation = new(@"[،؛؟!](?=[\u0600-\u06FF])", RegexOptions.Compiled);
    private static readonly Regex SeparatedMi = new(@"(?<![\p{L}\p{N}])(می|نمی)\s+([\u0600-\u06FF]{2,})", RegexOptions.Compiled);
    private static readonly Regex RepeatedZwnj = new(@"\u200C{2,}", RegexOptions.Compiled);

    public PersianEditingAuditResult Inspect(string inputPath, int maxSamples = 250)
    {
        var input = PathPolicy.RequireSupportedWordFile(inputPath);
        using var file = new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read, leaveOpen: false);
        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidOperationException("word/document.xml is missing.");

        XDocument document;
        using (var stream = entry.Open())
            document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);

        int examined = 0;
        int guarded = 0;
        int yk = 0;
        int repeatSpace = 0;
        int beforePunct = 0;
        int afterPunct = 0;
        int mi = 0;
        int zwnj = 0;
        var issues = new List<PersianEditingIssue>();

        var paragraphs = document.Descendants(W + "p").ToList();
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var paragraph = paragraphs[i];
            var text = string.Concat(paragraph.Descendants(W + "t").Select(x => (string?)x ?? string.Empty));
            if (string.IsNullOrWhiteSpace(text))
                continue;

            if (paragraph.Descendants(W + "fldChar").Any() ||
                paragraph.Descendants(W + "instrText").Any() ||
                paragraph.Descendants(W + "fldSimple").Any() ||
                paragraph.Descendants(M + "oMath").Any() ||
                paragraph.Descendants(M + "oMathPara").Any())
            {
                guarded++;
                continue;
            }

            examined++;
            var preview = Preview(text, 180);

            if (text.IndexOf('\u064A') >= 0 || text.IndexOf('\u0643') >= 0)
            {
                yk++;
                Add(issues, maxSamples, i + 1, "ARABIC_YEH_KAF", "SAFE", preview,
                    "حروف عربی ي/ك را در متن فارسی به ی/ک فارسی تبدیل کنید؛ Field/Equationها Guard شده‌اند.");
            }

            if (RepeatedSpace.IsMatch(text))
            {
                repeatSpace++;
                Add(issues, maxSamples, i + 1, "REPEATED_SPACE", "SAFE", preview,
                    "فاصله‌های متوالی را به یک فاصله کاهش دهید.");
            }

            if (SpaceBeforePunctuation.IsMatch(text))
            {
                beforePunct++;
                Add(issues, maxSamples, i + 1, "SPACE_BEFORE_PUNCTUATION", "SAFE", preview,
                    "فاصله قبل از نشانه نگارشی فارسی حذف شود.");
            }

            if (MissingSpaceAfterPunctuation.IsMatch(text))
            {
                afterPunct++;
                Add(issues, maxSamples, i + 1, "MISSING_SPACE_AFTER_PUNCTUATION", "REVIEW", preview,
                    "پس از نشانه نگارشی احتمالاً یک فاصله لازم است؛ قبل از Mutation باید Context بررسی شود.");
            }

            if (SeparatedMi.IsMatch(text))
            {
                mi++;
                Add(issues, maxSamples, i + 1, "MI_PREFIX_SPACE", "SAFE_WITH_CONTEXT", preview,
                    "در فعل‌های فارسی «می/نمی + فاصله» را پس از تأیید واژه به نیم‌فاصله تبدیل کنید.");
            }

            if (RepeatedZwnj.IsMatch(text))
            {
                zwnj++;
                Add(issues, maxSamples, i + 1, "REPEATED_ZWNJ", "SAFE", preview,
                    "نیم‌فاصله تکراری به یک نیم‌فاصله کاهش یابد.");
            }
        }

        return new PersianEditingAuditResult(
            "wordagent.editing.audit/1.0",
            input,
            Sha256(input),
            examined,
            guarded,
            yk + repeatSpace + beforePunct + afterPunct + mi + zwnj,
            yk,
            repeatSpace,
            beforePunct,
            afterPunct,
            mi,
            zwnj,
            false,
            issues);
    }

    public static string FormatPersianReport(PersianEditingAuditResult result)
    {
        var sb = new StringBuilder();
        sb.AppendLine("ممیزی ویراستاری صوری WordAgent");
        sb.AppendLine(new string('=', 60));
        sb.AppendLine("Schema: " + result.Schema);
        sb.AppendLine("فایل: " + result.InputFile);
        sb.AppendLine("SHA-256: " + result.InputSha256);
        sb.AppendLine("Mutation: NO — این مرحله فقط ممیزی است.");
        sb.AppendLine();
        sb.AppendLine("[گیت‌های Guard]");
        sb.AppendLine($"  Paragraphs examined: {result.ParagraphsExamined}");
        sb.AppendLine($"  Field/Equation guarded paragraphs: {result.GuardedParagraphs}");
        sb.AppendLine();
        sb.AppendLine("[یافته‌ها]");
        sb.AppendLine($"  Arabic Yeh/Kaf: {result.ArabicYehKafParagraphs}");
        sb.AppendLine($"  Repeated spaces: {result.RepeatedSpaceParagraphs}");
        sb.AppendLine($"  Space before punctuation: {result.SpaceBeforePunctuationParagraphs}");
        sb.AppendLine($"  Missing space after punctuation: {result.MissingSpaceAfterPunctuationParagraphs}");
        sb.AppendLine($"  Separated می/نمی prefix: {result.SeparatedMiPrefixParagraphs}");
        sb.AppendLine($"  Repeated ZWNJ: {result.RepeatedZwnjParagraphs}");
        sb.AppendLine($"  Total issue instances: {result.IssueCount}");
        sb.AppendLine();
        sb.AppendLine("[نمونه‌ها]");
        foreach (var issue in result.Issues)
        {
            sb.AppendLine($"  P{issue.ParagraphIndex} | {issue.Code} | {issue.Severity}");
            sb.AppendLine("    " + issue.TextPreview);
            sb.AppendLine("    پیشنهاد: " + issue.Suggestion);
        }
        sb.AppendLine();
        sb.AppendLine("[وضعیت]");
        sb.AppendLine("  این نسخه هیچ متن، Citation، Field، Equation یا EndNote را تغییر نمی‌دهد.");
        sb.AppendLine("  Mutation ویراستاری تا مرحله بعدی و پس از Acceptance ممیزی، غیرفعال است.");
        return sb.ToString();
    }

    private static void Add(
        List<PersianEditingIssue> issues,
        int maxSamples,
        int paragraphIndex,
        string code,
        string severity,
        string preview,
        string suggestion)
    {
        if (issues.Count >= maxSamples) return;
        issues.Add(new PersianEditingIssue(paragraphIndex, code, severity, preview, suggestion));
    }

    private static string Preview(string text, int max)
    {
        var value = string.Join(' ', text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).Trim();
        return value.Length <= max ? value : value[..max] + "...";
    }

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
