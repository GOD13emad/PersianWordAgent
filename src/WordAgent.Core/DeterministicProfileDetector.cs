using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordAgent.Core;

public sealed class DeterministicProfileDetector : IProfileDetector
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const int HeadCharacters = 5000;
    private const int StyleSignalWeight = 3;
    private const int IntentKeywordWeight = 20;
    private const int IntentHeadBonus = 10;
    private const int IntentMinimumDistinctMatches = 3;
    private const int IntentExclusionMinimumDistinctMatches = 3;
    private const int AbsoluteAmbiguityMargin = 10;
    private const double RelativeAmbiguityMargin = 0.08;

    // Smart Detection v2 structural priors. These are intentionally deterministic,
    // auditable, and subordinate to explicit profile selection.
    private const int QuestionnaireFrontMatterPrior = 110;
    private const int QuestionnaireTablePrior = 30;
    private const int QuestionnaireDenseTablePrior = 30;
    private const int GovernanceFrontMatterPrior = 120;
    private const int GovernanceTocSectionPrior = 20;
    private const int ThesisFrontMatterPrior = 80;
    private const int ThesisTocPrior = 40;
    private const int ThesisSectionPrior = 30;
    private const int ThesisHeadingPrior = 20;
    private const int ThesisLongDocumentPrior = 20;
    private const int AcademicProposalThesisSignaturePenalty = 40;

    private readonly string _profilesDirectory;

    public DeterministicProfileDetector(string profilesDirectory)
    {
        if (string.IsNullOrWhiteSpace(profilesDirectory))
            throw new ArgumentException("Profiles directory is required.", nameof(profilesDirectory));

        _profilesDirectory = profilesDirectory;
    }

    public async Task<DocumentProfileCandidate> DetectAsync(string path, CancellationToken ct)
    {
        var result = await DetectDetailedAsync(path, ct).ConfigureAwait(false);
        if (result.Status != ProfileDetectionStatus.Detected || string.IsNullOrWhiteSpace(result.ProfileId))
            throw new InvalidOperationException($"Profile detection did not produce a unique result: {result.Status}.");

        var winner = result.RankedCandidates[0];
        return new DocumentProfileCandidate(winner.ProfileId, result.Confidence, winner.Evidence);
    }

    public Task<ProfileDetectionResult> DetectDetailedAsync(string path, CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();
        var input = PathPolicy.RequireSupportedWordFile(path);

        if (!Directory.Exists(_profilesDirectory))
            throw new DirectoryNotFoundException($"Profiles directory not found: {_profilesDirectory}");

        var profiles = Directory.EnumerateFiles(_profilesDirectory, "*.json", SearchOption.TopDirectoryOnly)
            .Select(p => ProfileLoader.Load(_profilesDirectory, Path.GetFileNameWithoutExtension(p)))
            .Where(p => p.DetectionPolicy?.Enabled == true)
            .ToArray();

        if (profiles.Length == 0)
            return Task.FromResult(new ProfileDetectionResult(
                ProfileDetectionStatus.NoMatch, null, 0d, Array.Empty<ProfileDetectionScore>()));

        var signals = ReadDocumentSignals(input);

        var ranked = profiles
            .Select(p => ScoreProfile(p, signals))
            .OrderByDescending(x => x.Score)
            .ThenByDescending(x => x.Priority)
            .ThenBy(x => x.ProfileId, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var eligible = ranked
            .Where(x =>
            {
                var p = profiles.First(p => p.Id.Equals(x.ProfileId, StringComparison.OrdinalIgnoreCase));
                return x.Score >= Math.Max(1, p.DetectionPolicy.MinScore);
            })
            .ToArray();

        if (eligible.Length == 0)
            return Task.FromResult(new ProfileDetectionResult(
                ProfileDetectionStatus.NoMatch, null, 0d, ranked));

        var top = eligible[0];
        var second = eligible.Length > 1 ? eligible[1] : null;

        if (second is not null)
        {
            var margin = top.Score - second.Score;
            var relativeMargin = margin / (double)Math.Max(1, top.Score);

            // A result is ambiguous only when BOTH the absolute and relative gaps are small.
            // The previous max-threshold rule could reject a document even after it cleared the
            // absolute safety margin by one or more points (the real B007 smoke document scored
            // governance 142 vs formal-admin 131: absolute gap 11, relative gap 7.7%).
            // Keeping both independent guards is deterministic while avoiding false ambiguity
            // at a strong absolute separation.
            if (margin < AbsoluteAmbiguityMargin &&
                relativeMargin < RelativeAmbiguityMargin)
            {
                return Task.FromResult(new ProfileDetectionResult(
                    ProfileDetectionStatus.Ambiguous, null, ComputeConfidence(top.Score, second.Score), ranked));
            }
        }

        return Task.FromResult(new ProfileDetectionResult(
            ProfileDetectionStatus.Detected,
            top.ProfileId,
            ComputeConfidence(top.Score, second?.Score ?? 0),
            ranked));
    }

    private static ProfileDetectionScore ScoreProfile(AgentProfile profile, DocumentSignals signals)
    {
        var score = 0;
        var evidence = new List<string>();

        foreach (var raw in profile.Classification.Keywords ?? Array.Empty<string>())
        {
            var keyword = Compact(raw);
            if (keyword.Length == 0 || !signals.FullCompact.Contains(keyword, StringComparison.Ordinal))
                continue;

            var weight = KeywordWeight(raw);
            if (signals.HeadCompact.Contains(keyword, StringComparison.Ordinal))
                weight += 6;

            score += weight;
            evidence.Add($"+keyword:{raw}={weight}");
        }

        // Document-intent markers are deliberately separate from topical keywords.
        // A profile receives this bonus only after a minimum number of distinct
        // intent markers are present. This lets an industrial proposal about
        // governance remain an IndustrialProposal while preventing a single
        // generic term such as "budget" or "client" from hijacking detection.
        var intentMatches = new List<(string Raw, string Compact, int Weight)>();
        var seenIntent = new HashSet<string>(StringComparer.Ordinal);
        foreach (var raw in profile.Classification.IntentKeywords ?? Array.Empty<string>())
        {
            var keyword = Compact(raw);
            if (keyword.Length == 0 ||
                !signals.FullCompact.Contains(keyword, StringComparison.Ordinal) ||
                !seenIntent.Add(keyword))
            {
                continue;
            }

            var weight = IntentKeywordWeight + PhraseBonus(raw);
            if (signals.HeadCompact.Contains(keyword, StringComparison.Ordinal))
                weight += IntentHeadBonus;

            intentMatches.Add((raw, keyword, weight));
        }

        var intentExclusionMatches = new List<string>();
        var seenIntentExclusion = new HashSet<string>(StringComparer.Ordinal);
        foreach (var raw in profile.Classification.IntentExclusionKeywords ?? Array.Empty<string>())
        {
            var keyword = Compact(raw);
            if (keyword.Length == 0 ||
                !signals.FullCompact.Contains(keyword, StringComparison.Ordinal) ||
                !seenIntentExclusion.Add(keyword))
            {
                continue;
            }

            intentExclusionMatches.Add(raw);
            evidence.Add($"~intent-exclusion-marker:{raw}");
        }

        if (intentExclusionMatches.Count >= IntentExclusionMinimumDistinctMatches)
        {
            evidence.Add($"~intent-excluded:{intentExclusionMatches.Count}/{IntentExclusionMinimumDistinctMatches}");
        }
        else if (intentMatches.Count >= IntentMinimumDistinctMatches)
        {
            foreach (var match in intentMatches)
            {
                score += match.Weight;
                evidence.Add($"+intent:{match.Raw}={match.Weight}");
            }
        }
        else if (intentMatches.Count > 0)
        {
            evidence.Add($"~intent-suppressed:{intentMatches.Count}/{IntentMinimumDistinctMatches}");
        }

        foreach (var raw in profile.Classification.NegativeKeywords ?? Array.Empty<string>())
        {
            var keyword = Compact(raw);
            if (keyword.Length == 0 || !signals.FullCompact.Contains(keyword, StringComparison.Ordinal))
                continue;

            var weight = 10 + PhraseBonus(raw);
            if (signals.HeadCompact.Contains(keyword, StringComparison.Ordinal))
                weight += 4;

            score -= weight;
            evidence.Add($"-keyword:{raw}={weight}");
        }

        var matchedStyles = new HashSet<string>(StringComparer.Ordinal);
        foreach (var raw in profile.Classification.StyleSignals ?? Array.Empty<string>())
        {
            var normalized = Normalize(raw);
            if (normalized.Length == 0 ||
                !signals.UsedStyleNames.Contains(normalized) ||
                !matchedStyles.Add(normalized))
            {
                continue;
            }

            score += StyleSignalWeight;
            evidence.Add($"+style:{raw}={StyleSignalWeight}");
        }

        ApplyStructuralPriors(profile, signals, ref score, evidence);

        return new ProfileDetectionScore(
            profile.Id,
            score,
            profile.DetectionPolicy?.Priority ?? 0,
            evidence);
    }

    private static void ApplyStructuralPriors(
        AgentProfile profile,
        DocumentSignals signals,
        ref int score,
        List<string> evidence)
    {
        var id = profile.Id.Trim().ToLowerInvariant();

        if (id.Equals("questionnaire-assessment-fa", StringComparison.OrdinalIgnoreCase))
        {
            if (signals.FrontMatterCompact.Contains(Compact("پرسشنامه"), StringComparison.Ordinal))
            {
                score += QuestionnaireFrontMatterPrior;
                evidence.Add($"+structure:questionnaire-front-matter={QuestionnaireFrontMatterPrior}");
            }

            if (signals.TableCount >= 8)
            {
                score += QuestionnaireTablePrior;
                evidence.Add($"+structure:table-count>={8}={QuestionnaireTablePrior}");
            }

            if (signals.TableCellCount >= 100)
            {
                score += QuestionnaireDenseTablePrior;
                evidence.Add($"+structure:table-cells>={100}={QuestionnaireDenseTablePrior}");
            }
        }
        else if (id.Equals("proposal.governance.fa", StringComparison.OrdinalIgnoreCase))
        {
            var governanceFrontMatter =
                signals.PrimaryTitleCompact.Contains(Compact("حاکمیت داده"), StringComparison.Ordinal)
                || signals.PrimaryTitleCompact.Contains(Compact("راهبری داده"), StringComparison.Ordinal)
                || signals.PrimaryTitleCompact.Contains(Compact("حکمرانی داده"), StringComparison.Ordinal);

            if (governanceFrontMatter)
            {
                score += GovernanceFrontMatterPrior;
                evidence.Add($"+structure:governance-front-matter={GovernanceFrontMatterPrior}");
            }

            if (governanceFrontMatter && signals.TocFieldCount > 0 && signals.SectionCount >= 2)
            {
                score += GovernanceTocSectionPrior;
                evidence.Add($"+structure:governance-toc-sections={GovernanceTocSectionPrior}");
            }
        }
        else if (id.Equals("thesis.dissertation.phd.fa", StringComparison.OrdinalIgnoreCase))
        {
            var doctoralFrontMatter =
                signals.FrontMatterCompact.Contains(Compact("رساله دکتری"), StringComparison.Ordinal)
                || signals.FrontMatterCompact.Contains(Compact("پایان نامه دکتری"), StringComparison.Ordinal)
                || signals.FrontMatterCompact.Contains(Compact("پایان‌نامه دکتری"), StringComparison.Ordinal);

            if (doctoralFrontMatter && signals.TocFieldCount > 0 && signals.SectionCount >= 4)
            {
                score += ThesisFrontMatterPrior;
                evidence.Add($"+structure:doctoral-front-matter-with-thesis-structure={ThesisFrontMatterPrior}");
            }

            if (signals.TocFieldCount > 0)
            {
                score += ThesisTocPrior;
                evidence.Add($"+structure:toc-field={ThesisTocPrior}");
            }

            if (signals.SectionCount >= 4)
            {
                score += ThesisSectionPrior;
                evidence.Add($"+structure:sections>={4}={ThesisSectionPrior}");
            }

            if (signals.HeadingParagraphCount >= 8)
            {
                score += ThesisHeadingPrior;
                evidence.Add($"+structure:heading-paragraphs>={8}={ThesisHeadingPrior}");
            }

            if (signals.FullCharacterCount >= 50000)
            {
                score += ThesisLongDocumentPrior;
                evidence.Add($"+structure:long-document={ThesisLongDocumentPrior}");
            }
        }
        else if (id.Equals("proposal.academic.fa", StringComparison.OrdinalIgnoreCase))
        {
            var strongThesisSignature =
                signals.TocFieldCount > 0
                && signals.SectionCount >= 4
                && signals.HeadingParagraphCount >= 8
                && signals.FrontMatterCompact.Contains(Compact("رساله دکتری"), StringComparison.Ordinal);

            if (strongThesisSignature)
            {
                score -= AcademicProposalThesisSignaturePenalty;
                evidence.Add($"-structure:strong-thesis-signature={AcademicProposalThesisSignaturePenalty}");
            }
        }
    }

    private static int KeywordWeight(string raw) => 8 + PhraseBonus(raw);

    private static int PhraseBonus(string raw)
    {
        var words = Regex.Split(Normalize(raw), @"\s+")
            .Count(x => !string.IsNullOrWhiteSpace(x));

        return Math.Min(8, Math.Max(0, words - 1) * 2);
    }

    private static double ComputeConfidence(int topScore, int secondScore)
    {
        if (topScore <= 0) return 0d;
        if (secondScore <= 0) return Math.Min(0.99d, 0.80d + Math.Min(0.19d, topScore / 1000d));

        var marginRatio = Math.Max(0d, (topScore - secondScore) / (double)topScore);
        return Math.Round(Math.Min(0.99d, 0.60d + marginRatio * 0.40d), 3);
    }

    private static DocumentSignals ReadDocumentSignals(string path)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read, leaveOpen: false);

        var document = LoadXml(archive, "word/document.xml");
        var styles = LoadXml(archive, "word/styles.xml");

        var text = string.Join(" ", document.Descendants(W + "t").Select(x => x.Value));
        var head = text.Length <= HeadCharacters ? text : text[..HeadCharacters];

        var styleMap = styles.Descendants(W + "style")
            .Select(s => new
            {
                Id = (string?)s.Attribute(W + "styleId"),
                Name = (string?)s.Element(W + "name")?.Attribute(W + "val")
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Id) && !string.IsNullOrWhiteSpace(x.Name))
            .GroupBy(x => x.Id!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => Normalize(g.First().Name!), StringComparer.OrdinalIgnoreCase);

        var usedStyleNames = new HashSet<string>(StringComparer.Ordinal);
        var headingParagraphCount = 0;
        foreach (var paragraph in document.Descendants(W + "p"))
        {
            var pStyle = paragraph.Descendants(W + "pStyle").FirstOrDefault();
            if (pStyle is null) continue;

            var id = (string?)pStyle.Attribute(W + "val");
            if (string.IsNullOrWhiteSpace(id)) continue;

            string styleName;
            if (styleMap.TryGetValue(id, out var mappedName) && mappedName.Length > 0)
                styleName = mappedName;
            else
                styleName = Normalize(id);

            usedStyleNames.Add(styleName);
            if (LooksLikeHeadingStyle(styleName))
                headingParagraphCount++;
        }

        var nonEmptyParagraphTexts = document.Descendants(W + "p")
            .Select(p => string.Join("", p.Descendants(W + "t").Select(x => x.Value)).Trim())
            .Where(x => x.Length > 0)
            .ToArray();

        // The primary title block is deliberately narrow. A topical subtitle in paragraph four
        // must not convert an industrial proposal into a governance proposal.
        var primaryTitleText = string.Join(" ", nonEmptyParagraphTexts.Take(3));

        var frontMatterText = string.Join(" ", nonEmptyParagraphTexts.Take(30));

        var instructionText = string.Join(" ", document.Descendants(W + "instrText").Select(x => x.Value));
        var tocFieldCount = Regex.Matches(instructionText, @"\bTOC\b", RegexOptions.IgnoreCase).Count;
        var tableCount = document.Descendants(W + "tbl").Count();
        var tableCellCount = document.Descendants(W + "tc").Count();
        var sectionCount = document.Descendants(W + "sectPr").Count();
        var contentControlCount = document.Descendants(W + "sdt").Count();

        return new DocumentSignals(
            Compact(text),
            Compact(head),
            Compact(primaryTitleText),
            Compact(frontMatterText),
            usedStyleNames,
            tableCount,
            tableCellCount,
            sectionCount,
            tocFieldCount,
            contentControlCount,
            headingParagraphCount,
            text.Length);
    }

    private static bool LooksLikeHeadingStyle(string normalizedStyleName)
    {
        if (string.IsNullOrWhiteSpace(normalizedStyleName))
            return false;

        return normalizedStyleName.Contains("heading", StringComparison.OrdinalIgnoreCase)
               || normalizedStyleName.Contains("تیتر", StringComparison.Ordinal)
               || normalizedStyleName.Contains("عنوان", StringComparison.Ordinal)
               || normalizedStyleName.Contains("فصل", StringComparison.Ordinal)
               || normalizedStyleName.Contains("chapter", StringComparison.OrdinalIgnoreCase);
    }

    private static XDocument LoadXml(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"DOCX package entry missing: {entryName}");

        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    internal static string Normalize(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var normalized = value.Normalize(NormalizationForm.FormKC)
            .Replace('ي', 'ی')
            .Replace('ك', 'ک')
            .Replace('\u200C', ' ')
            .Replace('\u200D', ' ')
            .Replace('\u200E', ' ')
            .Replace('\u200F', ' ')
            .ToLowerInvariant();

        return Regex.Replace(normalized, @"\s+", " ").Trim();
    }

    internal static string Compact(string value)
    {
        var normalized = Normalize(value);
        if (normalized.Length == 0) return string.Empty;

        var sb = new StringBuilder(normalized.Length);
        foreach (var ch in normalized)
        {
            var category = CharUnicodeInfo.GetUnicodeCategory(ch);
            if (char.IsLetterOrDigit(ch) ||
                category == UnicodeCategory.NonSpacingMark ||
                category == UnicodeCategory.SpacingCombiningMark)
            {
                sb.Append(ch);
            }
        }
        return sb.ToString();
    }

    private sealed record DocumentSignals(
        string FullCompact,
        string HeadCompact,
        string PrimaryTitleCompact,
        string FrontMatterCompact,
        HashSet<string> UsedStyleNames,
        int TableCount,
        int TableCellCount,
        int SectionCount,
        int TocFieldCount,
        int ContentControlCount,
        int HeadingParagraphCount,
        int FullCharacterCount);
}
