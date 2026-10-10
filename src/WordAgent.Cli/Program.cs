using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Threading;
using System.IO.Compression;
using System.Xml.Linq;
using System.Reflection;
using WordAgent.Core;
using WordAgent.OpenXmlAudit;
using WordAgent.WordInterop;

namespace WordAgent.Cli;

internal static class Program
{
    private static readonly string Version = typeof(Program).Assembly
        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
        .InformationalVersion
        ?? throw new InvalidOperationException("CLI informational version is missing.");
    private const string DefaultProfile = "formal-admin-technical-proposal-fa";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private static readonly XNamespace WordMain = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private sealed record TocScopeSnapshot(
        int Controls,
        int Paragraphs,
        int VisibleParagraphs,
        int FieldInstructions,
        string OutsideVisibleTextSha256,
        int OutsideParagraphs,
        int OutsideTables,
        int OutsideFieldInstructions,
        int OutsideHyperlinks,
        int OutsideSections,
        int OutsideDrawings);

    private sealed record TocLocalOpenXmlDeltaProof(
        bool Applicable,
        bool Accepted,
        int TotalParagraphDelta,
        int TocParagraphDelta,
        int TotalFieldInstructionDelta,
        int TocFieldInstructionDelta,
        int BeforeVisibleTocParagraphs,
        int AfterVisibleTocParagraphs,
        string Reason);


    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        try
        {
            if (args.Length == 0) return Help();
            var command = args[0].Trim().ToLowerInvariant();
            var options = ParseOptions(args.Skip(1).ToArray());

            return command switch
            {
                "help" or "-h" or "--help" => Help(),
                "version" or "--version" => PrintVersion(),
                "doctor" => Doctor(options),
                "audit" => Audit(options),
                "edit-audit" => EditAudit(options),
                "plan" => Plan(options),
                "layout-plan" => LayoutPlan(options),
                "layout-fix-controlled" => LayoutFixControlled(options),
                "persianize-generic" or "persianize" => PersianizeGeneric(options),
                "academicize-thesis" => AcademicizeThesis(options),
                "template-analyze" => TemplateAnalyze(options),
                "template-apply" => TemplateApply(options),
                "fix" => LegacyFixDisabled(),
                "fix-controlled" => ControlledFix(options),
                "profiles" => ListProfiles(options),
                "profile-validate" => ValidateProfiles(options),
                "detect-profile" => DetectProfile(options),
                "__internal-word-audit-snapshot" => InternalWordAuditSnapshot(options),
                "__internal-profile-compliance-snapshot" => InternalProfileComplianceSnapshot(options),
                _ => Unknown(command)
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("ERROR: " + ex.Message);
            if (Environment.GetEnvironmentVariable("WORDAGENT_DEBUG") == "1")
                Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static int Help()
    {
        Console.WriteLine($"""
WordAgent {Version}

Commands:
  version
  doctor
  profiles [--profiles <dir>]
  profile-validate [--profile <id>] [--profiles <dir>]
  detect-profile --input <file.docx> [--profiles <dir>]
  audit --input <file.docx> [--profile <id>] [--profiles <dir>] [--outdir <dir>] [--visible]
  edit-audit --input <file.docx> [--outdir <dir>]
  plan  --input <file.docx> [--profile <id>] [--profiles <dir>] [--outdir <dir>] [--visible]
  layout-plan --input <file.docx> [--profile <id>] [--profiles <dir>] [--table-theme LIGHT_BLUE|BLUE|GREEN] [--outdir <dir>] [--visible]
  layout-fix-controlled --input <file.docx> --confirm-layout-plan [--profile <id>] [--profiles <dir>] [--table-theme LIGHT_BLUE|BLUE|GREEN] [--outdir <dir>] [--visible]
  fix-controlled --input <file.docx> (--confirm-plan-ready | --confirm-smart-repair) [--profile <id>] [--profiles <dir>] [--outdir <dir>] [--visible]
  persianize-generic --input <file.docx> --confirm-generic [--outdir <dir>] [--visible]
  academicize-thesis --input <file.docx> --confirm-academicize [--profile thesis-dissertation-academic-fa|thesis.dissertation.phd.fa] [--profiles <dir>] [--outdir <dir>]
  template-analyze --template <file.docx> [--outdir <dir>]
  template-apply --input <file.docx> --template <file.docx> --confirm-template-apply [--outdir <dir>] [--visible]
  fix   (legacy mutation is disabled; use fix-controlled after reviewing plan)

Profile selection:
  If --profile is omitted, WordAgent runs deterministic profile detection and fails closed
  on AMBIGUOUS / NOT_DETECTED. --profile remains an explicit override.

Examples:
  WordAgent.exe doctor
  WordAgent.exe audit --input "H:\Docs\sample.docx"
  WordAgent.exe plan --input "H:\Docs\sample.docx"
  WordAgent.exe layout-plan --input "H:\Docs\sample.docx"
  WordAgent.exe layout-fix-controlled --input "H:\Docs\sample.docx" --confirm-layout-plan
  WordAgent.exe fix-controlled --input "H:\Docs\sample.docx" --confirm-plan-ready
  WordAgent.exe persianize-generic --input "H:\Docs\sample.docx" --confirm-generic
  WordAgent.exe template-analyze --template "H:\Templates\journal.docx"
  WordAgent.exe template-apply --input "H:\Docs\article.docx" --template "H:\Templates\journal.docx" --confirm-template-apply
""");
        return 0;
    }

    private static int PrintVersion()
    {
        Console.WriteLine(Version);
        return 0;
    }

    private static int Doctor(Dictionary<string, string?> options)
    {
        var probe = WordRuntime.Probe();
        Console.WriteLine($"WordAgent: {Version}");
        Console.WriteLine($"OS: {Environment.OSVersion}");
        Console.WriteLine($"64-bit process: {Environment.Is64BitProcess}");
        Console.WriteLine("WORD_PROBE_MODE=COM_FREE_NO_ACTIVATION");
        Console.WriteLine($"Microsoft Word: {(probe.Installed ? "OK" : "NOT FOUND")}");
        if (probe.Version is not null) Console.WriteLine($"Word version: {probe.Version}");
        if (probe.Error is not null) Console.WriteLine($"Word error: {probe.Error}");
        return probe.Installed ? 0 : 2;
    }

    private static int ListProfiles(Dictionary<string, string?> options)
    {
        var dir = ResolveProfilesDir(options);
        if (!Directory.Exists(dir))
        {
            Console.Error.WriteLine($"Profiles directory not found: {dir}");
            return 2;
        }
        foreach (var file in Directory.EnumerateFiles(dir, "*.json").OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            Console.WriteLine(Path.GetFileNameWithoutExtension(file));
        return 0;
    }

    private static int DetectProfile(Dictionary<string, string?> options)
    {
        if (!options.TryGetValue("input", out var inputValue) || string.IsNullOrWhiteSpace(inputValue))
        {
            Console.Error.WriteLine("Missing required option: --input <file.docx>");
            return 2;
        }

        var input = PathPolicy.RequireSupportedWordFile(inputValue!);
        var profilesDir = ResolveProfilesDir(options);

        var detector = new DeterministicProfileDetector(profilesDir);
        var result = detector.DetectDetailedAsync(input).GetAwaiter().GetResult();

        var selectedCandidate = result.RankedCandidates
            .FirstOrDefault(candidate =>
                !string.IsNullOrWhiteSpace(result.ProfileId)
                && candidate.ProfileId.Equals(result.ProfileId, StringComparison.OrdinalIgnoreCase))
            ?? result.RankedCandidates.FirstOrDefault();

        var positiveTextEvidence = selectedCandidate?.Evidence
            .Count(evidence => evidence.StartsWith("+keyword:", StringComparison.OrdinalIgnoreCase))
            ?? 0;
        var styleEvidence = selectedCandidate?.Evidence
            .Count(evidence => evidence.Contains("style:", StringComparison.OrdinalIgnoreCase))
            ?? 0;
        var structuralEvidence = selectedCandidate?.Evidence
            .Count(evidence => evidence.StartsWith("+structure:", StringComparison.OrdinalIgnoreCase)
                               || evidence.StartsWith("-structure:", StringComparison.OrdinalIgnoreCase))
            ?? 0;
        var detectionBasis = structuralEvidence > 0
            ? "TEXT_PLUS_STRUCTURE"
            : positiveTextEvidence > 0 ? "TEXT_CONTENT" : "STRUCTURE_STYLE_FALLBACK";
        var top2 = result.RankedCandidates.Take(2).ToArray();
        var genericFallbackPolicy = result.Status == ProfileDetectionStatus.Detected
            ? "NOT_REQUIRED"
            : "PERSIANIZE_GENERIC_REVIEW_REQUIRED";

        Console.WriteLine($"DETECTION_STATUS={result.Status.ToString().ToUpperInvariant()}");
        Console.WriteLine($"PROFILE_ID={result.ProfileId ?? ""}");
        Console.WriteLine($"CONFIDENCE={result.Confidence:0.000}");
        Console.WriteLine($"CONFIDENCE_PERCENT={result.Confidence * 100.0:0.0}");
        Console.WriteLine($"DETECTION_BASIS={detectionBasis}");
        Console.WriteLine($"TEXT_EVIDENCE_COUNT={positiveTextEvidence}");
        Console.WriteLine($"STYLE_EVIDENCE_COUNT={styleEvidence}");
        Console.WriteLine($"STRUCTURAL_EVIDENCE_COUNT={structuralEvidence}");
        Console.WriteLine($"TOP2_COUNT={top2.Length}");
        for (var i = 0; i < top2.Length; i++)
            Console.WriteLine($"TOP2_{i + 1}={top2[i].ProfileId}:{top2[i].Score}");
        Console.WriteLine($"GENERIC_FALLBACK_POLICY={genericFallbackPolicy}");

        foreach (var candidate in result.RankedCandidates.Take(8))
        {
            Console.WriteLine(
                $"CANDIDATE | {candidate.ProfileId} | score={candidate.Score} | priority={candidate.Priority}");

            foreach (var evidence in candidate.Evidence.Take(12))
                Console.WriteLine($"  EVIDENCE | {evidence}");
        }

        return result.Status switch
        {
            ProfileDetectionStatus.Detected => 0,
            ProfileDetectionStatus.Ambiguous => 3,
            _ => 4
        };
    }

    private static string ResolveProfileId(
        Dictionary<string, string?> options,
        string path,
        string profilesDir)
    {
        var requested = options.GetValueOrDefault("profile");
        if (!string.IsNullOrWhiteSpace(requested))
        {
            Console.WriteLine($"PROFILE_RESOLUTION=EXPLICIT:{requested}");
            return requested!;
        }

        var detector = new DeterministicProfileDetector(profilesDir);
        var result = detector.DetectDetailedAsync(path).GetAwaiter().GetResult();

        if (result.Status != ProfileDetectionStatus.Detected
            || string.IsNullOrWhiteSpace(result.ProfileId))
        {
            var ranked = string.Join(
                ",",
                result.RankedCandidates.Take(5).Select(x => $"{x.ProfileId}:{x.Score}"));

            throw new InvalidOperationException(
                $"Automatic profile detection failed closed. STATUS={result.Status}; "
                + $"CONFIDENCE={result.Confidence:0.000}; CANDIDATES={ranked}. "
                + "Generic fallback policy: use persianize-generic --confirm-generic only after reviewing "
                + "the detect-profile Top-2 evidence; otherwise specify --profile explicitly.");
        }

        Console.WriteLine(
            $"PROFILE_RESOLUTION=AUTO_DETECTED:{result.ProfileId}:CONFIDENCE={result.Confidence:0.000}:CONFIDENCE_PERCENT={result.Confidence * 100.0:0.0}");

        return result.ProfileId!;
    }

    private static int ValidateProfiles(Dictionary<string, string?> options)
    {
        var dir = ResolveProfilesDir(options);
        if (!Directory.Exists(dir))
        {
            Console.Error.WriteLine($"Profiles directory not found: {dir}");
            return 2;
        }

        var requested = options.GetValueOrDefault("profile");
        var profileIds = string.IsNullOrWhiteSpace(requested)
            ? Directory.EnumerateFiles(dir, "*.json", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileNameWithoutExtension)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Cast<string>()
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToArray()
            : new[] { requested! };

        if (profileIds.Length == 0)
        {
            Console.Error.WriteLine($"No profile JSON files found under: {dir}");
            return 2;
        }

        var failures = 0;
        foreach (var profileId in profileIds)
        {
            try
            {
                var profile = ProfileLoader.Load(dir, profileId);
                var validation = ProfileValidator.Validate(profile);

                Console.WriteLine(
                    $"PASS | {profile.Id} | schema={profile.SchemaVersion} | profileVersion={profile.Version} | class={profile.DocumentClass}");

                foreach (var issue in validation.Issues.Where(x =>
                    x.Severity.Equals("WARNING", StringComparison.OrdinalIgnoreCase)))
                {
                    Console.WriteLine($"  WARNING | {issue.Code} | {issue.Message}");
                }
            }
            catch (Exception ex)
            {
                failures++;
                Console.Error.WriteLine($"FAIL | {profileId} | {ex.Message}");
            }
        }

        Console.WriteLine($"PROFILE_VALIDATION_ERRORS={failures}");
        return failures == 0 ? 0 : 2;
    }

    private static int Audit(Dictionary<string, string?> options)
    {
        var input = Require(options, "input");
        var path = PathPolicy.RequireSupportedWordFile(input);
        var outDir = options.GetValueOrDefault("outdir") ?? Path.Combine(Path.GetDirectoryName(path)!, "WordAgent_Audit");
        Directory.CreateDirectory(outDir);
        var visible = IsTrue(options, "visible");
        var profilesDir = ResolveProfilesDir(options);
        var profileId = ResolveProfileId(options, path, profilesDir);

        var warnings = new List<string>();
        OpenXmlAuditSnapshot? openXml = null;
        WordAuditSnapshot? word = null;
        ProfileComplianceSnapshot? compliance = null;
        AgentProfile? profile = null;

        try { profile = ProfileLoader.Load(profilesDir, profileId); }
        catch (Exception ex) { warnings.Add("Profile load failed: " + ex.Message); }

        try { openXml = new OpenXmlInspector().Inspect(path); }
        catch (Exception ex) { warnings.Add("OpenXML audit failed: " + ex.Message); }

        var probe = WordRuntime.Probe();
        if (probe.Installed)
        {
            try { word = InspectWordIsolated(path, visible); }
            catch (Exception ex) { warnings.Add("Word audit failed: " + ex.Message); }

            if (profile is not null)
            {
                try { compliance = InspectComplianceIsolated(path, profile.Id, profilesDir, visible); }
                catch (Exception ex) { warnings.Add("Profile compliance audit failed: " + ex.Message); }
            }
        }
        else
        {
            warnings.Add("Microsoft Word Desktop is not available; effective-formatting and profile-compliance audits were skipped.");
        }

        var report = new CombinedAuditReport("wordagent.audit/1.3", DateTimeOffset.Now, path, word, openXml, compliance, warnings);
        var stem = Path.GetFileNameWithoutExtension(path);
        var jsonPath = Path.Combine(outDir, stem + ".audit.json");
        var txtPath = Path.Combine(outDir, stem + ".audit.fa.txt");
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(report, JsonOptions), new UTF8Encoding(false));
        File.WriteAllText(txtPath, FormatPersianReport(report), new UTF8Encoding(false));

        Console.WriteLine("Audit JSON: " + jsonPath);
        Console.WriteLine("Audit report: " + txtPath);
        return warnings.Count == 0 ? 0 : 3;
    }


    private static int EditAudit(Dictionary<string, string?> options)
    {
        var input = Require(options, "input");
        var path = PathPolicy.RequireSupportedWordFile(input);
        var outDir = options.GetValueOrDefault("outdir")
            ?? Path.Combine(Path.GetDirectoryName(path)!, "WordAgent_EditAudit");
        Directory.CreateDirectory(outDir);

        var result = new PersianEditingAuditService().Inspect(path);
        var stem = Path.GetFileNameWithoutExtension(path);
        var jsonPath = Path.Combine(outDir, stem + ".edit-audit.json");
        var txtPath = Path.Combine(outDir, stem + ".edit-audit.fa.txt");

        File.WriteAllText(jsonPath, JsonSerializer.Serialize(result, JsonOptions), new UTF8Encoding(false));
        File.WriteAllText(txtPath, PersianEditingAuditService.FormatPersianReport(result), new UTF8Encoding(false));

        Console.WriteLine("Editing audit JSON: " + jsonPath);
        Console.WriteLine("Editing audit report: " + txtPath);
        Console.WriteLine($"Mechanical issues: {result.IssueCount} | Paragraphs examined: {result.ParagraphsExamined} | Guarded/skipped: {result.GuardedParagraphs}");
        Console.WriteLine("Mutation: NO");
        return 0;
    }


    private static int Plan(Dictionary<string, string?> options)
    {
        var input = Require(options, "input");
        var path = PathPolicy.RequireSupportedWordFile(input);
        var visible = IsTrue(options, "visible");
        var profilesDir = ResolveProfilesDir(options);
        var profileId = ResolveProfileId(options, path, profilesDir);
        var outDir = options.GetValueOrDefault("outdir") ?? Path.Combine(Path.GetDirectoryName(path)!, "WordAgent_Plan");
        Directory.CreateDirectory(outDir);

        var probe = WordRuntime.Probe();
        if (!probe.Installed)
            throw new InvalidOperationException("Microsoft Word Desktop is required for Safe-Fix planning. " + probe.Error);

        var profile = ProfileLoader.Load(profilesDir, profileId);
        var word = InspectWordIsolated(path, visible);
        var compliance = InspectComplianceIsolated(path, profileId, profilesDir, visible);

        OpenXmlAuditSnapshot? openXml = null;
        string? openXmlWarning = null;
        try { openXml = new OpenXmlInspector().Inspect(path); }
        catch (Exception ex) { openXmlWarning = ex.Message; }

        var planText = FormatSafeFixPlan(path, profile, word, compliance, openXml, openXmlWarning);
        var stem = Path.GetFileNameWithoutExtension(path);
        var txtPath = Path.Combine(outDir, stem + ".fixplan.fa.txt");
        File.WriteAllText(txtPath, planText, new UTF8Encoding(false));

        var styleRepairable = compliance.Styles.Count(x => x.Found && !x.FullyCompliant);
        var styleFailures = compliance.Styles.Count(x => !x.Found && !x.FullyCompliant);
        var gateReady = compliance.UnclassifiedParagraphs == 0 && styleFailures == 0;

        Console.WriteLine("Safe-Fix plan: " + txtPath);
        Console.WriteLine("READ-ONLY: YES");
        Console.WriteLine("SAFE_FIX_GATE=" + (gateReady ? "READY" : "BLOCKED"));
        Console.WriteLine("SAFE_FIX_UNCLASSIFIED=" + compliance.UnclassifiedParagraphs);
        Console.WriteLine("SAFE_FIX_STYLE_FAILURES=" + styleFailures);
        Console.WriteLine("SAFE_FIX_STYLE_REPAIRABLE=" + styleRepairable);
        return gateReady ? 0 : 5;
    }

    private static string CanonicalizeRoleAliasFamily(string role)
    {
        if (role.Equals("TocLegacy1", StringComparison.OrdinalIgnoreCase))
            return "Toc1";
        if (role.Equals("TocLegacy2", StringComparison.OrdinalIgnoreCase))
            return "Toc2";
        return role;
    }

    private static string[] ResolveEquivalentRoleAliases(AgentProfile profile, string role)
    {
        var canonical = CanonicalizeRoleAliasFamily(role);
        var aliases = profile.RoleAliases
            .Where(pair => CanonicalizeRoleAliasFamily(pair.Key)
                .Equals(canonical, StringComparison.OrdinalIgnoreCase))
            .SelectMany(pair => pair.Value ?? Array.Empty<string>())
            .Where(alias => !string.IsNullOrWhiteSpace(alias))
            .Select(alias => alias.Trim())
            .ToList();

        // Some semantic roles intentionally reuse a style whose primary alias belongs
        // to another role. CoverSubtitle is the current supported example.
        var coverSubtitleRole = profile.SemanticPolicy.CoverSubtitle.Role;
        if (!string.IsNullOrWhiteSpace(coverSubtitleRole)
            && CanonicalizeRoleAliasFamily(coverSubtitleRole)
                .Equals(canonical, StringComparison.OrdinalIgnoreCase))
        {
            aliases.AddRange(
                (profile.SemanticPolicy.CoverSubtitle.StyleNames ?? Array.Empty<string>())
                .Where(alias => !string.IsNullOrWhiteSpace(alias))
                .Select(alias => alias.Trim()));
        }

        return aliases
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string? ResolvePreferredExistingStyle(
        AgentProfile profile,
        string role,
        IReadOnlyList<string> equivalentAliases,
        ProfileComplianceSnapshot compliance)
    {
        var canonical = CanonicalizeRoleAliasFamily(role);
        var ordered = new List<string>();

        if (profile.RoleAliases.TryGetValue(canonical, out var canonicalAliases))
        {
            ordered.AddRange(
                (canonicalAliases ?? Array.Empty<string>())
                .Where(alias => !string.IsNullOrWhiteSpace(alias))
                .Select(alias => alias.Trim()));
        }

        ordered.AddRange(equivalentAliases.Where(alias => !string.IsNullOrWhiteSpace(alias)));

        foreach (var candidate in ordered.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (compliance.Styles.Any(item =>
                    item.Found
                    && string.Equals(item.StyleName?.Trim(), candidate.Trim(), StringComparison.OrdinalIgnoreCase)))
                return candidate.Trim();
        }

        return null;
    }

    private static bool RoleStyleBelongsToAliasFamily(
        IReadOnlyList<string> equivalentAliases,
        string? styleName)
        => !string.IsNullOrWhiteSpace(styleName)
           && equivalentAliases.Any(alias => string.Equals(
               alias.Trim(),
               styleName.Trim(),
               StringComparison.OrdinalIgnoreCase));

    private static bool HasScheduledStyleRepairForSample(
        ProfileComplianceSnapshot compliance,
        string role,
        string? styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
            return false;

        var canonical = CanonicalizeRoleAliasFamily(role);
        return compliance.Styles.Any(item =>
            item.Found
            && !item.FullyCompliant
            && CanonicalizeRoleAliasFamily(item.Role)
                .Equals(canonical, StringComparison.OrdinalIgnoreCase)
            && string.Equals(
                item.StyleName?.Trim(),
                styleName.Trim(),
                StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsManagedTocRole(string role)
    {
        var canonical = CanonicalizeRoleAliasFamily(role);
        return canonical.Equals("Toc1", StringComparison.OrdinalIgnoreCase)
               || canonical.Equals("Toc2", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasFullyCompliantExactStyleForSample(
        ProfileComplianceSnapshot compliance,
        string role,
        string? styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
            return false;

        var canonical = CanonicalizeRoleAliasFamily(role);
        return compliance.Styles.Any(item =>
            item.Found
            && item.FullyCompliant
            && CanonicalizeRoleAliasFamily(item.Role)
                .Equals(canonical, StringComparison.OrdinalIgnoreCase)
            && string.Equals(
                item.StyleName?.Trim(),
                styleName.Trim(),
                StringComparison.OrdinalIgnoreCase));
    }

    private static string FormatSafeFixPlan(
        string path,
        AgentProfile profile,
        WordAuditSnapshot word,
        ProfileComplianceSnapshot compliance,
        OpenXmlAuditSnapshot? openXml,
        string? openXmlWarning)
    {
        var sb = new StringBuilder();
        sb.AppendLine("برنامه اصلاح امن WordAgent");
        sb.AppendLine(new string('=', 60));
        sb.AppendLine("Schema: wordagent.fixplan/1.0");
        sb.AppendLine("فایل: " + path);
        sb.AppendLine("Profile: " + profile.Id + " | " + profile.DisplayName);
        sb.AppendLine("Mutation: NO — این فرمان فقط برنامه اصلاح را تولید می‌کند.");
        sb.AppendLine();

        var styleRepairable = compliance.Styles.Count(x => x.Found && !x.FullyCompliant);
        var styleFailures = compliance.Styles.Count(x => !x.Found && !x.FullyCompliant);
        var blockers = new List<string>();
        if (compliance.UnclassifiedParagraphs > 0)
            blockers.Add($"Unclassified paragraphs: {compliance.UnclassifiedParagraphs}");
        if (styleFailures > 0)
            blockers.Add($"Profile style-definition failures/missing: {styleFailures}");
        if (styleRepairable > 0)
            sb.AppendLine($"  REPAIRABLE | Existing profile styles requiring deterministic style repair={styleRepairable}");

        sb.AppendLine("[گیت‌های ایمنی]");
        sb.AppendLine($"  {(compliance.UnclassifiedParagraphs == 0 ? "PASS" : "BLOCK")} | Semantic classification | Unclassified={compliance.UnclassifiedParagraphs}");
        sb.AppendLine($"  {(styleFailures == 0 ? "PASS" : "BLOCK")} | Style definitions | Fail/Missing={styleFailures}");
        sb.AppendLine($"  GUARD | TOC fields={word.Counts.TocFields} | باید حفظ شوند و بازتولید/به‌روزرسانی فقط با درخواست صریح انجام شود.");
        sb.AppendLine($"  GUARD | EndNote fields={word.Counts.EndNoteFields} | در صورت وجود، Fieldها نباید Reset/بازشماری دستی شوند.");
        sb.AppendLine($"  GUARD | Revisions={word.Counts.Revisions} | Comments={word.Counts.Comments}");
        if (openXml is not null)
            sb.AppendLine($"  WARN  | OpenXML validation errors={openXml.ValidationErrors} | خطاهای پیش‌موجود به‌طور خودکار تعمیر نمی‌شوند.");
        if (!string.IsNullOrWhiteSpace(openXmlWarning))
            sb.AppendLine("  WARN  | OpenXML audit unavailable: " + openXmlWarning);
        sb.AppendLine();

        int totalCandidates = 0;

        sb.AppendLine("[اقدامات پیشنهادی — بدون اعمال تغییر]");
        foreach (var role in compliance.Roles.OrderBy(x => x.Role, StringComparer.OrdinalIgnoreCase))
        {
            var delta = role.Total - role.FullyCompliant;
            if (delta <= 0) continue;
            totalCandidates += delta;

            var aliases = ResolveEquivalentRoleAliases(profile, role.Role);
            var canonicalStyle = ResolvePreferredExistingStyle(profile, role.Role, aliases, compliance);

            sb.AppendLine($"Role {role.Role}: candidates={delta}/{role.Total}");
            foreach (var sample in role.Samples)
            {
                var styleBelongsToRole = RoleStyleBelongsToAliasFamily(aliases, sample.StyleName);
                var styleRepairScheduled = HasScheduledStyleRepairForSample(
                    compliance,
                    role.Role,
                    sample.StyleName);
                var managedTocResultOverride =
                    IsManagedTocRole(role.Role)
                    && styleBelongsToRole
                    && HasFullyCompliantExactStyleForSample(
                        compliance,
                        role.Role,
                        sample.StyleName);

                string action;
                if (!styleBelongsToRole && !string.IsNullOrWhiteSpace(canonicalStyle))
                    action = $"ASSIGN_STYLE '{canonicalStyle}' + RESET_DIRECT_FORMAT";
                else if (managedTocResultOverride || styleRepairScheduled || sample.DirectOverrideSuspected)
                    action = "RESET_DIRECT_FORMAT_TO_STYLE";
                else
                    action = "APPLY_ROLE_DEFAULTS_WITH_GUARD";

                sb.AppendLine($"  P{sample.ParagraphIndex} | {action}");
                sb.AppendLine($"    Style={sample.StyleName} | FontBi={sample.ActualFontBi ?? "(mixed/unknown)"} | SizeBi={FormatSize(sample.ActualSizeBi)} | {sample.ActualReadingOrder}/{sample.ActualAlignment}");
                sb.AppendLine("    " + sample.TextPreview);
            }
            sb.AppendLine();
        }

        sb.AppendLine("[جمع‌بندی]");
        sb.AppendLine($"  Candidate paragraphs: {totalCandidates}");
        sb.AppendLine($"  Direct-formatting overrides suspected: {compliance.DirectOverridesSuspected}");
        sb.AppendLine($"  Tables: {word.Counts.Tables} | RTL={word.Formatting.RtlTables} | Centered={word.Formatting.CenteredTables}");
        sb.AppendLine($"  TOC: {word.Counts.TocFields} | EndNote: {word.Counts.EndNoteFields}");
        sb.AppendLine($"  Safe-Fix gate: {(blockers.Count == 0 ? "READY_FOR_CONTROLLED_COPY_FIX" : "BLOCKED")}");
        foreach (var blocker in blockers) sb.AppendLine("  BLOCKER: " + blocker);
        sb.AppendLine();
        sb.AppendLine("اصل فایل نباید overwrite شود. مرحله Mutation فقط روی COPY و پس از Audit مجدد مجاز است.");
        return FinalizeFaReport(sb);
    }


    private static int LayoutPlan(Dictionary<string, string?> options)
    {
        var input = Require(options, "input");
        var path = PathPolicy.RequireSupportedWordFile(input);
        var visible = IsTrue(options, "visible");
        var profilesDir = ResolveProfilesDir(options);
        var profileId = ResolveProfileId(options, path, profilesDir);
        var outDir = options.GetValueOrDefault("outdir") ?? Path.Combine(Path.GetDirectoryName(path)!, "WordAgent_LayoutPlan");
        Directory.CreateDirectory(outDir);

        var probe = WordRuntime.Probe();
        if (!probe.Installed)
            throw new InvalidOperationException("Microsoft Word Desktop is required for the layout quality gate. " + probe.Error);

        var profile = ProfileLoader.Load(profilesDir, profileId);
        ApplyTableThemeOverride(profile, options);
        var word = InspectWordIsolated(path, visible);
        var compliance = InspectComplianceIsolated(path, profileId, profilesDir, visible);
        var layout = new LayoutQualityService().Inspect(path, profile);

        OpenXmlAuditSnapshot? openXml = null;
        string? openXmlWarning = null;
        try { openXml = new OpenXmlInspector().Inspect(path); }
        catch (Exception ex) { openXmlWarning = ex.Message; }

        var text = FormatLayoutPlan(path, profile, word, compliance, layout, openXml, openXmlWarning);
        var stem = Path.GetFileNameWithoutExtension(path);
        var txtPath = Path.Combine(outDir, stem + ".layoutplan.fa.txt");
        File.WriteAllText(txtPath, text, new UTF8Encoding(false));

        Console.WriteLine("Layout plan: " + txtPath);
        Console.WriteLine("READ-ONLY: YES");
        Console.WriteLine("Layout pending actions: " + layout.PendingActions);
        return 0;
    }

    private static string FormatLayoutPlan(
        string path,
        AgentProfile profile,
        WordAuditSnapshot word,
        ProfileComplianceSnapshot compliance,
        LayoutQualityPlan layout,
        OpenXmlAuditSnapshot? openXml,
        string? openXmlWarning)
    {
        var sb = new StringBuilder();
        sb.AppendLine("برنامه کیفیت صفحه‌آرایی WordAgent");
        sb.AppendLine(new string('=', 60));
        sb.AppendLine("Schema: wordagent.layoutplan/1.0");
        sb.AppendLine("فایل: " + path);
        sb.AppendLine("Profile: " + profile.Id + " | " + profile.DisplayName);
        sb.AppendLine("Mutation: NO — این فرمان فقط برنامه صفحه‌آرایی را تولید می‌کند.");
        sb.AppendLine();

        var formattingRemaining = compliance.Roles.Sum(x => x.Total - x.FullyCompliant);
        var styleFailures = compliance.Styles.Count(x => !x.Found && !x.FullyCompliant);
        var formattingReady = formattingRemaining == 0 && compliance.UnclassifiedParagraphs == 0 && styleFailures == 0;
        var layoutReady = layout.Ready;

        sb.AppendLine("[گیت‌های ایمنی]");
        sb.AppendLine($"  {(formattingReady ? "PASS" : "BLOCK")} | Formatting baseline | Remaining={formattingRemaining} | Unclassified={compliance.UnclassifiedParagraphs} | StyleFail={styleFailures}");
        sb.AppendLine($"  {(layoutReady ? "PASS" : "BLOCK")} | Layout semantic gate | Blockers={layout.Blockers.Count}");
        sb.AppendLine($"  GUARD | TOC fields={word.Counts.TocFields} | EndNote fields={word.Counts.EndNoteFields} | Revisions={word.Counts.Revisions}");
        if (openXml is not null)
            sb.AppendLine($"  WARN  | OpenXML validation errors={openXml.ValidationErrors} | خطاهای پیش‌موجود به‌طور خودکار تعمیر نمی‌شوند.");
        if (!string.IsNullOrWhiteSpace(openXmlWarning))
            sb.AppendLine("  WARN  | OpenXML audit unavailable: " + openXmlWarning);
        foreach (var blocker in layout.Blockers) sb.AppendLine("  BLOCKER: " + blocker);
        sb.AppendLine();

        sb.AppendLine("[فهرست مطالب]");
        sb.AppendLine($"  TOC field found: {(layout.TocFieldFound ? "YES" : "NO")}");
        sb.AppendLine($"  TOC entry paragraphs: {layout.TocEntryParagraphs}");
        sb.AppendLine($"  TOC styles found: {layout.TocStylesFound} | styles needing compact spacing: {layout.TocStylesNeedingCompaction}");
        sb.AppendLine($"  TOC title found: {(layout.TocTitleFound ? "YES" : "NO")}");
        sb.AppendLine($"  First body chapter: {(layout.FirstContentFound ? layout.FirstContentText : "(not found)")}");
        sb.AppendLine($"  PageBreakBefore on first body chapter: {(layout.FirstContentHasPageBreakBefore ? "YES" : "NO")}");
        sb.AppendLine($"  Action: {(layout.TocStylesNeedingCompaction > 0 ? "COMPACT_TOC_STYLE_SPACING" : "NONE")} + {(layout.BodyPageBreakNeeded ? "PAGE_BREAK_BEFORE_FIRST_BODY_CHAPTER" : "NO_BODY_BREAK_CHANGE")}");
        sb.AppendLine();

        sb.AppendLine("[جدول‌آرایی]");
        sb.AppendLine($"  Enabled: {(profile.LayoutPolicy.TableStyle.Enabled ? "YES" : "NO")} | Theme={profile.LayoutPolicy.TableStyle.Theme} | Long Persian threshold={profile.LayoutPolicy.TableStyle.LongTextThreshold}");
        sb.AppendLine("  Header cells: centered + vertically centered; body cells: short/numeric centered, long Persian prose right-aligned; caption text/style is preserved.");
        sb.AppendLine();

        sb.AppendLine("[پیوست و جدول‌ها]");
        sb.AppendLine($"  Appendix found: {(layout.AppendixFound ? "YES" : "NO")} | {layout.AppendixTitle}");
        sb.AppendLine($"  Appendix tables: {layout.AppendixTables} | Wide tables: {layout.WideAppendixTables}");
        sb.AppendLine($"  Wide-tail supported: {(layout.WideTailGroupSupported ? "YES" : "NO")} | Landscape tail present: {(layout.LandscapeTailPresent ? "YES" : "NO")}");
        sb.AppendLine("  Pagination standard: compact tables stay whole; large tables split only between rows, repeat header, and avoid single-row fragments.");
        foreach (var table in layout.Tables)
        {
            sb.AppendLine($"  T{table.TableIndex} | cols={table.Columns} | grid={table.GridWidthTwips} | target={table.TargetWidthTwips} | wide={(table.IsWide ? "YES" : "NO")} | needs={(table.NeedsNormalization ? "YES" : "NO")}");
            if (!string.IsNullOrWhiteSpace(table.CaptionText)) sb.AppendLine("    Caption: " + table.CaptionText);
            sb.AppendLine("    " + table.Action);
        }
        sb.AppendLine();

        sb.AppendLine("[جمع‌بندی]");
        sb.AppendLine($"  Layout pending actions: {layout.PendingActions}");
        sb.AppendLine($"  Layout gate: {(formattingReady && layoutReady ? "READY_FOR_CONTROLLED_LAYOUT_FIX" : "BLOCKED")}");
        sb.AppendLine("  هدف: فهرست بدون orphan تک‌خطی، شروع متن اصلی از صفحه جدید، Landscape کنترل‌شده برای جدول عریض، و استاندارد جلوگیری از شکستن جدول‌های کوچک بین صفحات.");
        sb.AppendLine("  اصل فایل overwrite نمی‌شود؛ تغییر فقط روی COPY انجام می‌شود.");
        return FinalizeFaReport(sb);
    }

    private static int LayoutFixControlled(Dictionary<string, string?> options)
    {
        if (!IsTrue(options, "confirm-layout-plan"))
            throw new InvalidOperationException("Controlled layout fix requires --confirm-layout-plan after reviewing layout-plan.");

        var input = Require(options, "input");
        var path = PathPolicy.RequireSupportedWordFile(input);
        var visible = IsTrue(options, "visible");
        var profilesDir = ResolveProfilesDir(options);
        var profileId = ResolveProfileId(options, path, profilesDir);
        var outDir = options.GetValueOrDefault("outdir");

        var probe = WordRuntime.Probe();
        if (!probe.Installed)
            throw new InvalidOperationException("Microsoft Word Desktop is required for post-layout verification. " + probe.Error);

        var profile = ProfileLoader.Load(profilesDir, profileId);
        ApplyTableThemeOverride(profile, options);
        var beforeWord = InspectWordIsolated(path, visible);
        var beforeCompliance = InspectComplianceIsolated(path, profileId, profilesDir, visible);
        var formattingRemaining = beforeCompliance.Roles.Sum(x => x.Total - x.FullyCompliant);
        var styleFailures = beforeCompliance.Styles.Count(x => !x.FullyCompliant);
        if (formattingRemaining != 0 || beforeCompliance.UnclassifiedParagraphs != 0 || styleFailures != 0)
            throw new InvalidOperationException($"Layout fix blocked: formatting baseline is not clean. Remaining={formattingRemaining}, Unclassified={beforeCompliance.UnclassifiedParagraphs}, StyleFail={styleFailures}.");

        var service = new LayoutQualityService();
        var beforeLayout = service.Inspect(path, profile);
        if (!beforeLayout.Ready)
            throw new InvalidOperationException("Layout fix blocked: " + string.Join(" | ", beforeLayout.Blockers));

        OpenXmlAuditSnapshot beforeOpenXml;
        try { beforeOpenXml = new OpenXmlInspector().Inspect(path); }
        catch (Exception ex) { throw new InvalidOperationException("Pre-layout OpenXML audit failed; mutation blocked. " + ex.Message, ex); }

        Console.WriteLine("Layout mutation engine: OpenXML direct package edit (Word COM is read-only for verification).");
        var result = service.Apply(path, profile, outDir);

        var afterWord = InspectWordIsolated(result.OutputFile, visible);
        var afterCompliance = InspectComplianceIsolated(result.OutputFile, profileId, profilesDir, visible);
        OpenXmlAuditSnapshot afterOpenXml;
        try { afterOpenXml = new OpenXmlInspector().Inspect(result.OutputFile); }
        catch (Exception ex) { throw new InvalidOperationException("Post-layout OpenXML audit failed. Output is not approved. " + ex.Message, ex); }

        var issues = new List<string>();
        issues.AddRange(CompareLayoutWordIntegrity(beforeWord, afterWord, result.SectionBreakParagraphsInserted, result.TrailingBlankParagraphsRemoved));
        issues.AddRange(CompareLayoutOpenXmlIntegrity(beforeOpenXml, afterOpenXml, result.SectionBreakParagraphsInserted, result.TrailingBlankParagraphsRemoved));

        var afterRemaining = afterCompliance.Roles.Sum(x => x.Total - x.FullyCompliant);
        if (afterRemaining != 0) issues.Add($"Formatting compliance regressed after layout: Remaining={afterRemaining}.");
        if (afterCompliance.UnclassifiedParagraphs != 0) issues.Add($"Semantic classification regressed after layout: Unclassified={afterCompliance.UnclassifiedParagraphs}.");
        if (afterCompliance.DirectOverridesSuspected > beforeCompliance.DirectOverridesSuspected)
            issues.Add($"Direct-formatting overrides increased after layout: {beforeCompliance.DirectOverridesSuspected} -> {afterCompliance.DirectOverridesSuspected}.");
        if (afterCompliance.Styles.Any(x => !x.FullyCompliant)) issues.Add("Profile style-definition compliance regressed after layout.");
        if (result.AfterPlan.PendingActions != 0) issues.Add($"Layout idempotency failed: PendingActions={result.AfterPlan.PendingActions}.");

        var approved = issues.Count == 0;
        var reportDir = Path.Combine(Path.GetDirectoryName(result.OutputFile)!, "WordAgent_LayoutFix");
        Directory.CreateDirectory(reportDir);
        var stem = Path.GetFileNameWithoutExtension(result.OutputFile);
        var reportPath = Path.Combine(reportDir, stem + ".layoutfix.fa.txt");
        File.WriteAllText(reportPath, FormatLayoutFixReport(result, beforeWord, afterWord, beforeCompliance, afterCompliance, beforeOpenXml, afterOpenXml, issues, approved), new UTF8Encoding(false));

        var planDir = Path.Combine(Path.GetDirectoryName(result.OutputFile)!, "WordAgent_LayoutPlan");
        Directory.CreateDirectory(planDir);
        var postPlanPath = Path.Combine(planDir, stem + ".layoutplan.fa.txt");
        File.WriteAllText(postPlanPath, FormatLayoutPlan(result.OutputFile, profile, afterWord, afterCompliance, result.AfterPlan, afterOpenXml, null), new UTF8Encoding(false));

        Console.WriteLine("Layout intermediate output: " + result.OutputFile);
        Console.WriteLine("Layout report: " + reportPath);
        Console.WriteLine("Post-layout plan: " + postPlanPath);
        Console.WriteLine($"TOC styles changed: {result.TocStylesChanged} | Body page break changed: {result.BodyPageBreaksChanged}");
        Console.WriteLine($"Appendix tables normalized: {result.AppendixTablesNormalized} | Wide tables expanded: {result.WideTablesExpanded} | Tables styled: {result.TablesStyled}");
        Console.WriteLine($"Section-break paragraphs inserted: {result.SectionBreakParagraphsInserted}");
        Console.WriteLine($"Safe trailing blank paragraphs removed: {result.TrailingBlankParagraphsRemoved}");
        Console.WriteLine("Layout pending actions after fix: " + result.AfterPlan.PendingActions);
        foreach (var warning in result.Warnings.Where(x => x.StartsWith("Trailing formatting-only blank paragraph", StringComparison.Ordinal)))
            Console.WriteLine("Layout hygiene: " + warning);
        Console.WriteLine("Integrity gate: " + (approved ? "PASS" : "FAIL"));

        if (!approved)
        {
            foreach (var issue in issues.Take(50)) Console.WriteLine("BLOCK: " + issue);
            Console.WriteLine("OUTPUT STATUS: NOT APPROVED — do not use this copy.");
            return 6;
        }

        var auditOptions = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            ["input"] = result.OutputFile,
            ["profile"] = profileId,
            ["profiles"] = profilesDir,
            ["outdir"] = Path.Combine(Path.GetDirectoryName(result.OutputFile)!, "WordAgent_Audit")
        };
        if (visible) auditOptions["visible"] = "true";
        Audit(auditOptions);

        var planOptions = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            ["input"] = result.OutputFile,
            ["profile"] = profileId,
            ["profiles"] = profilesDir,
            ["outdir"] = Path.Combine(Path.GetDirectoryName(result.OutputFile)!, "WordAgent_Plan")
        };
        if (visible) planOptions["visible"] = "true";
        Plan(planOptions);

        // F02: The user-facing final artifact must combine controlled profile/layout
        // finishing with the accepted profile-independent Persian typography pass.
        // This always writes another COPY; the original and the layout intermediate
        // are never overwritten.
        Console.WriteLine("Final Persianization: applying deterministic RTL+Justify pass after controlled layout.");
        var finalPersian = new GenericPersianizeService().Apply(result.OutputFile, outDir);

        OpenXmlAuditSnapshot finalOpenXml;
        try { finalOpenXml = new OpenXmlInspector().Inspect(finalPersian.OutputFile); }
        catch (Exception ex) { throw new InvalidOperationException("Post-final-Persianization OpenXML audit failed. Output is not approved. " + ex.Message, ex); }

        var finalWord = InspectWordIsolated(finalPersian.OutputFile, visible);
        var finalIssues = new List<string>();
        finalIssues.AddRange(finalPersian.IntegrityIssues);
        finalIssues.AddRange(CompareLayoutWordIntegrity(afterWord, finalWord, 0));
        finalIssues.AddRange(CompareLayoutOpenXmlIntegrity(afterOpenXml, finalOpenXml, 0));

        var finalApproved = finalIssues.Count == 0;
        var finalReportDir = Path.Combine(Path.GetDirectoryName(finalPersian.OutputFile)!, "WordAgent_Persianize");
        Directory.CreateDirectory(finalReportDir);
        var finalStem = Path.GetFileNameWithoutExtension(finalPersian.OutputFile);
        var finalReportPath = Path.Combine(finalReportDir, finalStem + ".persianize.fa.txt");
        File.WriteAllText(
            finalReportPath,
            FormatGenericPersianizeReport(finalPersian, afterWord, finalWord, afterOpenXml, finalOpenXml, finalIssues, finalApproved),
            new UTF8Encoding(false));

        Console.WriteLine("Final Persianize report: " + finalReportPath);
        Console.WriteLine($"Final Persian body paragraphs justified: {finalPersian.BodyParagraphsJustified}");
        Console.WriteLine($"Final Persian paragraphs: {finalPersian.PersianParagraphs} | Paragraphs changed: {finalPersian.ParagraphsChanged} | Runs changed: {finalPersian.RunsChanged}");
        Console.WriteLine("Final Persianization integrity gate: " + (finalApproved ? "PASS" : "FAIL"));

        if (!finalApproved)
        {
            foreach (var issue in finalIssues.Take(50)) Console.WriteLine("BLOCK: " + issue);
            Console.WriteLine("Final Persianization: FAIL");
            Console.WriteLine("OUTPUT STATUS: NOT APPROVED — do not use this copy.");
            return 7;
        }

        Console.WriteLine("Layout output: " + finalPersian.OutputFile);
        Console.WriteLine("Final Persianization: PASS");
        Console.WriteLine("OUTPUT STATUS: APPROVED_FOR_VISUAL_REVIEW");
        return 0;
    }

    private static List<string> CompareLayoutWordIntegrity(WordAuditSnapshot before, WordAuditSnapshot after, int insertedBreakParagraphs, int removedSafeTrailingBlankParagraphs = 0)
    {
        var issues = new List<string>();
        void Equal(string name, int a, int b) { if (a != b) issues.Add($"Word inventory changed unexpectedly: {name} {a} -> {b}"); }
        var expectedParagraphs = before.Counts.Paragraphs + insertedBreakParagraphs - removedSafeTrailingBlankParagraphs;
        if (after.Counts.Paragraphs != expectedParagraphs)
            issues.Add($"Word inventory changed unexpectedly: Paragraphs {before.Counts.Paragraphs} -> {after.Counts.Paragraphs}; expected {expectedParagraphs} (= +{insertedBreakParagraphs} layout break -{removedSafeTrailingBlankParagraphs} safe trailing blank paragraph(s)).");
        if (after.Counts.Sections != before.Counts.Sections + insertedBreakParagraphs)
            issues.Add($"Word inventory changed unexpectedly: Sections {before.Counts.Sections} -> {after.Counts.Sections}; expected +{insertedBreakParagraphs} layout section(s).");
        Equal("Tables", before.Counts.Tables, after.Counts.Tables);
        Equal("InlineShapes", before.Counts.InlineShapes, after.Counts.InlineShapes);
        Equal("FloatingShapes", before.Counts.FloatingShapes, after.Counts.FloatingShapes);
        Equal("Fields", before.Counts.Fields, after.Counts.Fields);
        Equal("EndNoteFields", before.Counts.EndNoteFields, after.Counts.EndNoteFields);
        Equal("TocFields", before.Counts.TocFields, after.Counts.TocFields);
        Equal("Styles", before.Counts.Styles, after.Counts.Styles);
        Equal("Footnotes", before.Counts.Footnotes, after.Counts.Footnotes);
        Equal("Endnotes", before.Counts.Endnotes, after.Counts.Endnotes);
        Equal("Comments", before.Counts.Comments, after.Counts.Comments);
        Equal("Revisions", before.Counts.Revisions, after.Counts.Revisions);
        Equal("ContentControls", before.Counts.ContentControls, after.Counts.ContentControls);
        Equal("Bookmarks", before.Counts.Bookmarks, after.Counts.Bookmarks);
        Equal("Hyperlinks", before.Counts.Hyperlinks, after.Counts.Hyperlinks);
        Equal("Equations", before.Counts.Equations, after.Counts.Equations);
        return issues;
    }

    private static List<string> CompareLayoutOpenXmlIntegrity(OpenXmlAuditSnapshot before, OpenXmlAuditSnapshot after, int insertedBreakParagraphs, int removedSafeTrailingBlankParagraphs = 0)
    {
        var issues = new List<string>();
        void Equal(string name, int a, int b) { if (a != b) issues.Add($"OpenXML structure changed unexpectedly: {name} {a} -> {b}"); }
        var expectedParagraphs = before.Paragraphs + insertedBreakParagraphs - removedSafeTrailingBlankParagraphs;
        if (after.Paragraphs != expectedParagraphs)
            issues.Add($"OpenXML paragraph count {before.Paragraphs} -> {after.Paragraphs}; expected {expectedParagraphs} (= +{insertedBreakParagraphs} section-break -{removedSafeTrailingBlankParagraphs} safe trailing blank paragraph(s)).");
        Equal("Tables", before.Tables, after.Tables);
        Equal("Drawings", before.Drawings, after.Drawings);
        Equal("Styles", before.Styles, after.Styles);
        Equal("FieldInstructions", before.FieldInstructions, after.FieldInstructions);
        Equal("EndNoteMarkers", before.EndNoteMarkers, after.EndNoteMarkers);
        if (after.ValidationErrors > before.ValidationErrors)
            issues.Add($"OpenXML validation errors increased: {before.ValidationErrors} -> {after.ValidationErrors}");
        return issues;
    }

    private static string FormatLayoutFixReport(
        LayoutMutationResult result,
        WordAuditSnapshot beforeWord,
        WordAuditSnapshot afterWord,
        ProfileComplianceSnapshot beforeCompliance,
        ProfileComplianceSnapshot afterCompliance,
        OpenXmlAuditSnapshot beforeOpenXml,
        OpenXmlAuditSnapshot afterOpenXml,
        IReadOnlyList<string> issues,
        bool approved)
    {
        var sb = new StringBuilder();
        sb.AppendLine("گزارش اصلاح کنترل‌شده صفحه‌آرایی WordAgent");
        sb.AppendLine(new string('=', 60));
        sb.AppendLine("Schema: wordagent.layoutfix/1.0");
        sb.AppendLine("Input: " + result.InputFile);
        sb.AppendLine("Output: " + result.OutputFile);
        sb.AppendLine("Profile: " + result.ProfileId);
        sb.AppendLine("Input SHA-256: " + result.InputSha256);
        sb.AppendLine("Output SHA-256: " + result.OutputSha256);
        sb.AppendLine("Original overwrite: NO");
        sb.AppendLine();

        sb.AppendLine("[Mutationهای صفحه‌آرایی]");
        sb.AppendLine($"  TOC styles compacted: {result.TocStylesChanged}");
        sb.AppendLine($"  Body page-break changes: {result.BodyPageBreaksChanged}");
        sb.AppendLine($"  Appendix tables normalized: {result.AppendixTablesNormalized}");
        sb.AppendLine($"  Wide tables expanded: {result.WideTablesExpanded}");
        sb.AppendLine($"  Section-break paragraphs inserted: {result.SectionBreakParagraphsInserted}");
        sb.AppendLine($"  Safe trailing blank paragraphs removed: {result.TrailingBlankParagraphsRemoved}");
        sb.AppendLine("  Page numbering on the new landscape tail: CONTINUOUS (no restart).");
        sb.AppendLine("  Table pagination: compact tables KEEP_WHOLE; large tables CONTROLLED_MULTI_PAGE with repeated header and no row split.");
        sb.AppendLine();

        var beforeRemaining = beforeCompliance.Roles.Sum(x => x.Total - x.FullyCompliant);
        var afterRemaining = afterCompliance.Roles.Sum(x => x.Total - x.FullyCompliant);
        sb.AppendLine("[گیت‌های بعد از Layout]");
        sb.AppendLine($"  Formatting remaining: {beforeRemaining} -> {afterRemaining}");
        sb.AppendLine($"  Direct overrides: {beforeCompliance.DirectOverridesSuspected} -> {afterCompliance.DirectOverridesSuspected}");
        sb.AppendLine($"  TOC fields: {beforeWord.Counts.TocFields} -> {afterWord.Counts.TocFields}");
        sb.AppendLine($"  Tables: {beforeWord.Counts.Tables} -> {afterWord.Counts.Tables}");
        sb.AppendLine($"  Sections: {beforeWord.Counts.Sections} -> {afterWord.Counts.Sections}");
        sb.AppendLine($"  Paragraphs: {beforeWord.Counts.Paragraphs} -> {afterWord.Counts.Paragraphs}");
        sb.AppendLine($"  OpenXML validation errors: {beforeOpenXml.ValidationErrors} -> {afterOpenXml.ValidationErrors}");
        sb.AppendLine($"  Layout pending actions: {result.BeforePlan.PendingActions} -> {result.AfterPlan.PendingActions}");
        sb.AppendLine();

        if (issues.Count > 0)
        {
            sb.AppendLine("[موارد مسدودکننده]");
            foreach (var issue in issues) sb.AppendLine("  - " + issue);
            sb.AppendLine();
        }

        if (result.Warnings.Count > 0)
        {
            sb.AppendLine("[هشدارها / یادداشت‌ها]");
            foreach (var warning in result.Warnings) sb.AppendLine("  - " + warning);
            sb.AppendLine();
        }

        sb.AppendLine("[وضعیت]");
        sb.AppendLine("  Integrity gate: " + (approved ? "PASS" : "FAIL"));
        sb.AppendLine("  Output status: " + (approved ? "APPROVED_FOR_VISUAL_REVIEW" : "NOT_APPROVED"));
        sb.AppendLine("  Visual review in Microsoft Word is still mandatory because pagination is a rendered-layout property.");
        return FinalizeFaReport(sb);
    }


    private static int TemplateApply(Dictionary<string, string?> options)
    {
        if (!IsTrue(options, "confirm-template-apply"))
            throw new InvalidOperationException("Template Apply requires --confirm-template-apply. The original input and template are never overwritten.");

        var inputValue = Require(options, "input");
        var templateValue = Require(options, "template");
        var input = PathPolicy.RequireSupportedWordFile(inputValue);
        var template = PathPolicy.RequireSupportedWordFile(templateValue);
        var outDir = options.GetValueOrDefault("outdir");
        var visible = IsTrue(options, "visible");

        var probe = WordRuntime.Probe();
        if (!probe.Installed)
            throw new InvalidOperationException("Microsoft Word Desktop is required for Template Apply verification. " + probe.Error);

        var templateHashBefore = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(template)));

        OpenXmlAuditSnapshot templateOpenXml;
        try { templateOpenXml = new OpenXmlInspector().Inspect(template); }
        catch (Exception ex) { throw new InvalidOperationException("Template OpenXML preflight failed; mutation blocked. " + ex.Message, ex); }

        if (templateOpenXml.ValidationErrors != 0)
        {
            var samples = string.Join(" | ", templateOpenXml.ValidationSamples.Take(8));
            throw new InvalidOperationException(
                $"Template OpenXML preflight failed: ValidationErrors={templateOpenXml.ValidationErrors}. {samples}");
        }

        Console.WriteLine("TEMPLATE_OPENXML_PREFLIGHT=PASS");

        OpenXmlAuditSnapshot beforeOpenXml;
        try { beforeOpenXml = new OpenXmlInspector().Inspect(input); }
        catch (Exception ex) { throw new InvalidOperationException("Pre-Template-Apply OpenXML audit failed; mutation blocked. " + ex.Message, ex); }

        var beforeWord = InspectWordIsolated(input, visible);

        Console.WriteLine("Template Apply mode: deterministic structural/style application.");
        Console.WriteLine("AI semantic mapping: NOT USED");
        Console.WriteLine("Mutation engine: OpenXML direct package edit; Word COM is read-only for verification.");

        var result = new TemplateApplyService().Apply(input, template, outDir);

        var templateHashAfter = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(template)));
        if (!templateHashBefore.Equals(templateHashAfter, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Template file changed unexpectedly during Template Apply.");

        OpenXmlAuditSnapshot afterOpenXml;
        try { afterOpenXml = new OpenXmlInspector().Inspect(result.OutputFile); }
        catch (Exception ex) { throw new InvalidOperationException("Post-Template-Apply OpenXML audit failed. Output is not approved. " + ex.Message, ex); }

        var afterWord = InspectWordIsolated(result.OutputFile, visible);

        var issues = new List<string>();
        issues.AddRange(result.IntegrityIssues);

        void WordEqual(string name, int before, int after)
        {
            if (before != after)
                issues.Add($"Word inventory changed unexpectedly: {name} {before} -> {after}");
        }

        WordEqual("Paragraphs", beforeWord.Counts.Paragraphs, afterWord.Counts.Paragraphs);
        WordEqual("Tables", beforeWord.Counts.Tables, afterWord.Counts.Tables);
        WordEqual("InlineShapes", beforeWord.Counts.InlineShapes, afterWord.Counts.InlineShapes);
        WordEqual("FloatingShapes", beforeWord.Counts.FloatingShapes, afterWord.Counts.FloatingShapes);
        WordEqual("Fields", beforeWord.Counts.Fields, afterWord.Counts.Fields);
        WordEqual("EndNoteFields", beforeWord.Counts.EndNoteFields, afterWord.Counts.EndNoteFields);
        WordEqual("TocFields", beforeWord.Counts.TocFields, afterWord.Counts.TocFields);
        WordEqual("Sections", beforeWord.Counts.Sections, afterWord.Counts.Sections);
        WordEqual("Footnotes", beforeWord.Counts.Footnotes, afterWord.Counts.Footnotes);
        WordEqual("Endnotes", beforeWord.Counts.Endnotes, afterWord.Counts.Endnotes);
        WordEqual("Comments", beforeWord.Counts.Comments, afterWord.Counts.Comments);
        WordEqual("Revisions", beforeWord.Counts.Revisions, afterWord.Counts.Revisions);
        WordEqual("ContentControls", beforeWord.Counts.ContentControls, afterWord.Counts.ContentControls);
        WordEqual("Bookmarks", beforeWord.Counts.Bookmarks, afterWord.Counts.Bookmarks);
        WordEqual("Hyperlinks", beforeWord.Counts.Hyperlinks, afterWord.Counts.Hyperlinks);
        WordEqual("Equations", beforeWord.Counts.Equations, afterWord.Counts.Equations);

        if (afterWord.Counts.Styles < beforeWord.Counts.Styles)
            issues.Add($"Word style inventory decreased unexpectedly: {beforeWord.Counts.Styles} -> {afterWord.Counts.Styles}");

        if (beforeOpenXml.Paragraphs != afterOpenXml.Paragraphs)
            issues.Add($"OpenXML paragraphs changed unexpectedly: {beforeOpenXml.Paragraphs} -> {afterOpenXml.Paragraphs}");
        if (beforeOpenXml.Tables != afterOpenXml.Tables)
            issues.Add($"OpenXML tables changed unexpectedly: {beforeOpenXml.Tables} -> {afterOpenXml.Tables}");
        if (beforeOpenXml.Drawings != afterOpenXml.Drawings)
            issues.Add($"OpenXML drawings changed unexpectedly: {beforeOpenXml.Drawings} -> {afterOpenXml.Drawings}");
        if (beforeOpenXml.FieldInstructions != afterOpenXml.FieldInstructions)
            issues.Add($"OpenXML field instructions changed unexpectedly: {beforeOpenXml.FieldInstructions} -> {afterOpenXml.FieldInstructions}");
        if (beforeOpenXml.EndNoteMarkers != afterOpenXml.EndNoteMarkers)
            issues.Add($"OpenXML EndNote markers changed unexpectedly: {beforeOpenXml.EndNoteMarkers} -> {afterOpenXml.EndNoteMarkers}");
        if (afterOpenXml.Styles < beforeOpenXml.Styles)
            issues.Add($"OpenXML style inventory decreased unexpectedly: {beforeOpenXml.Styles} -> {afterOpenXml.Styles}");
        if (afterOpenXml.ValidationErrors > beforeOpenXml.ValidationErrors)
            issues.Add($"OpenXML validation errors increased: {beforeOpenXml.ValidationErrors} -> {afterOpenXml.ValidationErrors}");

        var approved = issues.Count == 0;
        var reportDir = Path.Combine(Path.GetDirectoryName(result.OutputFile)!, "WordAgent_TemplateApply");
        Directory.CreateDirectory(reportDir);
        var stem = Path.GetFileNameWithoutExtension(result.OutputFile);
        var reportPath = Path.Combine(reportDir, stem + ".template-apply.fa.txt");
        File.WriteAllText(
            reportPath,
            FormatTemplateApplyReport(result, beforeWord, afterWord, beforeOpenXml, afterOpenXml, issues, approved),
            new UTF8Encoding(false));

        Console.WriteLine("Template output: " + result.OutputFile);
        Console.WriteLine("Template apply report: " + reportPath);
        Console.WriteLine($"Template styles imported: {result.TemplateStylesImported} | replaced: {result.TemplateStylesReplaced}");
        Console.WriteLine($"Paragraphs remapped: {result.ParagraphsRemapped} | Sections normalized: {result.SectionsNormalized}");
        Console.WriteLine($"Theme replaced: {result.ThemeReplaced} | Font table replaced: {result.FontTableReplaced}");
        Console.WriteLine("NO_AI_SEMANTIC_MAPPING=YES");
        Console.WriteLine("NO_CHANGES_NEEDED=" + result.NoChangesNeeded);
        Console.WriteLine("TEMPLATE_APPLY_STATUS=" + (approved ? "PASS" : "FAIL"));
        Console.WriteLine("Integrity gate: " + (approved ? "PASS" : "FAIL"));

        if (!approved)
        {
            foreach (var issue in issues.Take(50))
                Console.WriteLine("BLOCK: " + issue);

            Console.WriteLine("OUTPUT STATUS: NOT APPROVED — do not use this copy.");
            return 8;
        }

        Console.WriteLine("OUTPUT STATUS: APPROVED_FOR_VISUAL_REVIEW");
        return 0;
    }

    private static string FormatTemplateApplyReport(
        TemplateApplyResult result,
        WordAuditSnapshot beforeWord,
        WordAuditSnapshot afterWord,
        OpenXmlAuditSnapshot beforeOpenXml,
        OpenXmlAuditSnapshot afterOpenXml,
        IReadOnlyList<string> issues,
        bool approved)
    {
        var sb = new StringBuilder();
        sb.AppendLine("گزارش اعمال تمپلیت WordAgent");
        sb.AppendLine(new string('=', 60));
        sb.AppendLine("Schema: wordagent.template.apply/1.0");
        sb.AppendLine("Input: " + result.InputFile);
        sb.AppendLine("Template: " + result.TemplateFile);
        sb.AppendLine("Output: " + result.OutputFile);
        sb.AppendLine("Input SHA-256: " + result.InputSha256);
        sb.AppendLine("Template SHA-256: " + result.TemplateSha256);
        sb.AppendLine("Output SHA-256: " + result.OutputSha256);
        sb.AppendLine("Original input overwrite: NO");
        sb.AppendLine("Template mutation: NO");
        sb.AppendLine("AI semantic mapping: NOT USED");
        sb.AppendLine();

        sb.AppendLine("[دامنه اعمال قطعی تمپلیت]");
        sb.AppendLine($"  Template styles imported: {result.TemplateStylesImported}");
        sb.AppendLine($"  Template styles replaced: {result.TemplateStylesReplaced}");
        sb.AppendLine($"  Paragraphs remapped by deterministic style role: {result.ParagraphsRemapped}");
        sb.AppendLine($"  Sections normalized to template geometry: {result.SectionsNormalized}");
        sb.AppendLine($"  Theme replaced when target part existed: {result.ThemeReplaced}");
        sb.AppendLine($"  Font table replaced when target part existed: {result.FontTableReplaced}");
        sb.AppendLine($"  No changes needed: {result.NoChangesNeeded}");
        sb.AppendLine();

        if (result.StyleMappings.Count > 0)
        {
            sb.AppendLine("[نگاشت‌های ساختاری]");
            foreach (var mapping in result.StyleMappings.Take(50))
            {
                sb.AppendLine(
                    $"  {mapping.SourceStyleId} ({mapping.SourceStyleName ?? "—"}) -> " +
                    $"{mapping.TargetStyleId} ({mapping.TargetStyleName ?? "—"}) | " +
                    $"{mapping.Reason} | paragraphs={mapping.ParagraphCount}");
            }
            sb.AppendLine();
        }

        sb.AppendLine("[مواردی که در MVP عمداً حفظ می‌شوند]");
        sb.AppendLine("  Visible text / fields / images / relationships / source headers-footers / numbering topology / direct formatting.");
        sb.AppendLine("  Template header/footer transplantation and numbering remap are deferred to later guarded template stages.");
        sb.AppendLine();

        sb.AppendLine("[گیت‌های Integrity]");
        sb.AppendLine($"  Word paragraphs: {beforeWord.Counts.Paragraphs} -> {afterWord.Counts.Paragraphs}");
        sb.AppendLine($"  Word tables: {beforeWord.Counts.Tables} -> {afterWord.Counts.Tables}");
        sb.AppendLine($"  Word fields: {beforeWord.Counts.Fields} -> {afterWord.Counts.Fields}");
        sb.AppendLine($"  Word sections: {beforeWord.Counts.Sections} -> {afterWord.Counts.Sections}");
        sb.AppendLine($"  Word styles: {beforeWord.Counts.Styles} -> {afterWord.Counts.Styles}");
        sb.AppendLine($"  OpenXML styles: {beforeOpenXml.Styles} -> {afterOpenXml.Styles}");
        sb.AppendLine($"  OpenXML validation errors: {beforeOpenXml.ValidationErrors} -> {afterOpenXml.ValidationErrors}");
        sb.AppendLine($"  Text payload SHA preserved: {result.Before.TextPayloadSha256.Equals(result.After.TextPayloadSha256, StringComparison.OrdinalIgnoreCase)}");
        sb.AppendLine($"  Field instruction SHA preserved: {result.Before.FieldInstructionSha256.Equals(result.After.FieldInstructionSha256, StringComparison.OrdinalIgnoreCase)}");
        sb.AppendLine($"  Media payload SHA preserved: {result.Before.MediaPayloadSha256.Equals(result.After.MediaPayloadSha256, StringComparison.OrdinalIgnoreCase)}");
        sb.AppendLine();

        if (result.Warnings.Count > 0)
        {
            sb.AppendLine("[قواعد / مرزهای MVP]");
            foreach (var warning in result.Warnings)
                sb.AppendLine("  - " + warning);
            sb.AppendLine();
        }

        if (issues.Count > 0)
        {
            sb.AppendLine("[موارد مسدودکننده]");
            foreach (var issue in issues)
                sb.AppendLine("  - " + issue);
            sb.AppendLine();
        }

        sb.AppendLine("[وضعیت]");
        sb.AppendLine("  Integrity gate: " + (approved ? "PASS" : "FAIL"));
        sb.AppendLine("  Output status: " + (approved ? "APPROVED_FOR_VISUAL_REVIEW" : "NOT_APPROVED"));
        sb.AppendLine("  Visual review in Microsoft Word remains required before production use.");
        return FinalizeFaReport(sb);
    }

    private static int TemplateAnalyze(Dictionary<string, string?> options)
    {
        var template = Require(options, "template");
        var path = PathPolicy.RequireSupportedWordFile(template);
        var outDir = options.GetValueOrDefault("outdir")
            ?? Path.Combine(Path.GetDirectoryName(path)!, "WordAgent_Template_Contract");
        Directory.CreateDirectory(outDir);

        var result = new TemplateContractService().Analyze(path);
        var stem = Path.GetFileNameWithoutExtension(path);
        var jsonPath = Path.Combine(outDir, stem + ".template-contract.json");
        var reportPath = Path.Combine(outDir, stem + ".template-contract.fa.txt");

        File.WriteAllText(
            jsonPath,
            JsonSerializer.Serialize(result, JsonOptions),
            new UTF8Encoding(false));

        File.WriteAllText(
            reportPath,
            FormatTemplateContractReport(result),
            new UTF8Encoding(false));

        Console.WriteLine("Template Analyzer: read-only deterministic contract extraction.");
        Console.WriteLine("Template contract JSON: " + jsonPath);
        Console.WriteLine("Template contract report: " + reportPath);
        Console.WriteLine($"Template styles: {result.Styles.Count} | Sections: {result.Sections.Count} | Table styles: {result.TableStyleCount}");
        Console.WriteLine($"Headers: {result.HeaderPartCount} | Footers: {result.FooterPartCount} | TOC fields: {result.TocFieldCount}");
        Console.WriteLine("TEMPLATE_MUTATION=NO");
        Console.WriteLine("TEMPLATE_ANALYZE_STATUS=PASS");
        return 0;
    }

    private static string FormatTemplateContractReport(TemplateContractResult result)
    {
        var sb = new StringBuilder();
        sb.AppendLine("گزارش قرارداد تمپلیت WordAgent");
        sb.AppendLine(new string('=', 60));
        sb.AppendLine("Schema: " + result.SchemaVersion);
        sb.AppendLine("Template: " + result.TemplateFile);
        sb.AppendLine("Template SHA-256: " + result.TemplateSha256);
        sb.AppendLine("Mutation: NO");
        sb.AppendLine();
        sb.AppendLine("[ساختار بسته]");
        sb.AppendLine($"  Sections: {result.Sections.Count}");
        sb.AppendLine($"  Styles: {result.Styles.Count}");
        sb.AppendLine($"  Table styles: {result.TableStyleCount}");
        sb.AppendLine($"  Headers: {result.HeaderPartCount}");
        sb.AppendLine($"  Footers: {result.FooterPartCount}");
        sb.AppendLine($"  Abstract numbering definitions: {result.AbstractNumberingCount}");
        sb.AppendLine($"  Numbering instances: {result.NumberingInstanceCount}");
        sb.AppendLine($"  TOC fields: {result.TocFieldCount}");
        sb.AppendLine();

        sb.AppendLine("[سکشن‌ها]");
        foreach (var section in result.Sections)
        {
            sb.AppendLine(
                $"  #{section.Index}: page={section.PageWidthTwips}x{section.PageHeightTwips} twips | orientation={section.Orientation ?? "default"} | " +
                $"margins T/R/B/L={section.MarginTopTwips}/{section.MarginRightTwips}/{section.MarginBottomTwips}/{section.MarginLeftTwips} | " +
                $"columns={section.ColumnCount ?? 1}");
        }
        sb.AppendLine();

        sb.AppendLine("[استایل‌های کلیدی]");
        foreach (var style in result.Styles
                     .Where(x => x.IsDefault || x.IsQuickFormat || x.Type.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
                     .Take(80))
        {
            sb.AppendLine(
                $"  {style.StyleId} | {style.Name ?? "—"} | type={style.Type} | default={style.IsDefault} | " +
                $"basedOn={style.BasedOn ?? "—"} | align={style.Alignment ?? "—"} | bidi={style.Bidi} | " +
                $"font={style.FontComplexScript ?? style.FontAscii ?? style.FontHighAnsi ?? "—"} | size={style.SizeHalfPoints ?? "—"}");
        }

        sb.AppendLine();
        sb.AppendLine("[وضعیت]");
        sb.AppendLine("  READ_ONLY=YES");
        sb.AppendLine("  TEMPLATE_ANALYZE_STATUS=PASS");
        sb.AppendLine("  این قرارداد پایهٔ حالت «اعمال تمپلیت» است؛ در این مرحله هیچ تغییری روی تمپلیت انجام نشده است.");
        return FinalizeFaReport(sb);
    }


    private static int PersianizeGeneric(Dictionary<string, string?> options)
    {
        if (!IsTrue(options, "confirm-generic"))
            throw new InvalidOperationException("Generic Persianize requires --confirm-generic. This mode is profile-independent and always writes a copy.");

        var input = Require(options, "input");
        var path = PathPolicy.RequireSupportedWordFile(input);
        var outDir = options.GetValueOrDefault("outdir");
        var visible = IsTrue(options, "visible");

        if (IsTrue(options, "fast-verify"))
            return PersianizeGenericFast(path, outDir);

        var probe = WordRuntime.Probe();
        if (!probe.Installed)
            throw new InvalidOperationException("Microsoft Word Desktop is required for Generic Persianize verification. " + probe.Error);

        OpenXmlAuditSnapshot beforeOpenXml;
        try { beforeOpenXml = new OpenXmlInspector().Inspect(path); }
        catch (Exception ex) { throw new InvalidOperationException("Pre-Persianize OpenXML audit failed; mutation blocked. " + ex.Message, ex); }

        var beforeWord = InspectWordIsolated(path, visible);

        Console.WriteLine("Generic Persianize mode: profile-independent deterministic formatting.");
        Console.WriteLine("GENERIC_SEMANTIC_POLICY=V4_STYLELESS_CAPTION_BODY_AWARE");
        Console.WriteLine("Profile detection: NOT USED");
        Console.WriteLine("Mutation engine: OpenXML direct package edit; Word COM is read-only for verification.");

        var result = new GenericPersianizeService().Apply(path, outDir);

        OpenXmlAuditSnapshot afterOpenXml;
        try { afterOpenXml = new OpenXmlInspector().Inspect(result.OutputFile); }
        catch (Exception ex) { throw new InvalidOperationException("Post-Persianize OpenXML audit failed. Output is not approved. " + ex.Message, ex); }

        var afterWord = InspectWordIsolated(result.OutputFile, visible);

        var issues = new List<string>();
        issues.AddRange(result.IntegrityIssues);
        issues.AddRange(CompareLayoutWordIntegrity(beforeWord, afterWord, 0));
        issues.AddRange(CompareLayoutOpenXmlIntegrity(beforeOpenXml, afterOpenXml, 0));

        var approved = issues.Count == 0;
        var reportDir = Path.Combine(Path.GetDirectoryName(result.OutputFile)!, "WordAgent_Persianize");
        Directory.CreateDirectory(reportDir);
        var stem = Path.GetFileNameWithoutExtension(result.OutputFile);
        var reportPath = Path.Combine(reportDir, stem + ".persianize.fa.txt");
        File.WriteAllText(
            reportPath,
            FormatGenericPersianizeReport(result, beforeWord, afterWord, beforeOpenXml, afterOpenXml, issues, approved),
            new UTF8Encoding(false));

        Console.WriteLine("Generic output: " + result.OutputFile);
        Console.WriteLine("Generic report: " + reportPath);
        Console.WriteLine($"Persian paragraphs: {result.PersianParagraphs} | Paragraphs changed: {result.ParagraphsChanged} | Body paragraphs justified: {result.BodyParagraphsJustified} | Runs changed: {result.RunsChanged}");
        Console.WriteLine($"Tables examined: {result.TablesExamined} | Tables centered: {result.TablesCentered} | Persian tables RTL-enabled: {result.PersianTablesRtlEnabled}");
        Console.WriteLine($"TOC styles changed: {result.TocStylesChanged}");
        Console.WriteLine("AUTO_FIELD_UPDATE_DISABLED=" + result.AutomaticFieldUpdateDisabled);
        Console.WriteLine("NO_PROFILE_DETECTION=YES");
        Console.WriteLine("NO_CHANGES_NEEDED=" + result.NoChangesNeeded);
        Console.WriteLine("Integrity gate: " + (approved ? "PASS" : "FAIL"));

        if (!approved)
        {
            foreach (var issue in issues.Take(50)) Console.WriteLine("BLOCK: " + issue);
            Console.WriteLine("OUTPUT STATUS: NOT APPROVED — do not use this copy.");
            return 7;
        }

        var genericTocMaterialized = PersianDocumentFoundationService.MaterializeMainAutomaticTocAfterIntegrity(result.OutputFile);
        var genericExactInputRestored = RestoreExactInputWhenPackageContentEquivalent(path, result.OutputFile);
        Console.WriteLine("GENERIC_POST_INTEGRITY_NATIVE_TOC_MATERIALIZED=" + genericTocMaterialized);
        Console.WriteLine("GENERIC_POST_INTEGRITY_WORD_SAVE_USED=NO");
        Console.WriteLine("GENERIC_EXACT_INPUT_RESTORED=" + genericExactInputRestored);
        Console.WriteLine("OUTPUT STATUS: APPROVED_FOR_VISUAL_REVIEW");
        return 0;
    }

    private static bool RestoreExactInputWhenPackageContentEquivalent(string inputPath, string outputPath)
    {
        if (string.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
            return false;

        using var inputZip = ZipFile.OpenRead(inputPath);
        using var outputZip = ZipFile.OpenRead(outputPath);
        if (inputZip.Entries.Count != outputZip.Entries.Count)
            return false;

        var outputEntries = outputZip.Entries.ToDictionary(entry => entry.FullName, StringComparer.Ordinal);
        var bufferA = new byte[81920];
        var bufferB = new byte[81920];
        foreach (var inputEntry in inputZip.Entries)
        {
            if (!outputEntries.TryGetValue(inputEntry.FullName, out var outputEntry) || inputEntry.Length != outputEntry.Length)
                return false;
            using var inputStream = inputEntry.Open();
            using var outputStream = outputEntry.Open();
            while (true)
            {
                var readA = inputStream.Read(bufferA, 0, bufferA.Length);
                var readB = outputStream.Read(bufferB, 0, bufferB.Length);
                if (readA != readB) return false;
                if (readA == 0) break;
                if (!bufferA.AsSpan(0, readA).SequenceEqual(bufferB.AsSpan(0, readB))) return false;
            }
        }

        inputZip.Dispose();
        outputZip.Dispose();
        File.Copy(inputPath, outputPath, overwrite: true);
        return true;
    }

    private static int PersianizeGenericFast(string path, string? outDir)
    {
        var started = DateTimeOffset.UtcNow;
        OpenXmlAuditSnapshot beforeOpenXml;
        try { beforeOpenXml = new OpenXmlInspector().Inspect(path); }
        catch (Exception ex) { throw new InvalidOperationException("Fast pre-Persianize OpenXML audit failed; mutation blocked. " + ex.Message, ex); }

        Console.WriteLine("GENERIC_VERIFY_MODE=FAST_OPENXML");
        Console.WriteLine("WORD_COM_VERIFICATION=SKIPPED_IN_FAST_MODE");
        Console.WriteLine("Profile detection: NOT USED");
        Console.WriteLine("GENERIC_SEMANTIC_POLICY=V4_STYLELESS_CAPTION_BODY_AWARE");
        var result = new GenericPersianizeService().Apply(path, outDir);

        OpenXmlAuditSnapshot afterOpenXml;
        try { afterOpenXml = new OpenXmlInspector().Inspect(result.OutputFile); }
        catch (Exception ex) { throw new InvalidOperationException("Fast post-Persianize OpenXML audit failed. " + ex.Message, ex); }

        var issues = new List<string>();
        issues.AddRange(result.IntegrityIssues);
        issues.AddRange(CompareLayoutOpenXmlIntegrity(beforeOpenXml, afterOpenXml, 0));
        var approved = issues.Count == 0;

        var reportDir = Path.Combine(Path.GetDirectoryName(result.OutputFile)!, "WordAgent_Persianize");
        Directory.CreateDirectory(reportDir);
        var stem = Path.GetFileNameWithoutExtension(result.OutputFile);
        var reportPath = Path.Combine(reportDir, stem + ".persianize.fast.fa.txt");
        var elapsed = DateTimeOffset.UtcNow - started;
        File.WriteAllText(reportPath,
            "WordAgent Generic Persianize FAST\\n"
            + "Profile detection: NOT USED\\n"
            + "Word COM verification: SKIPPED\\n"
            + $"Persian paragraphs: {result.PersianParagraphs}\\n"
            + $"Body paragraphs justified: {result.BodyParagraphsJustified}\\n"
            + $"Semantic policy: V3 styleless text-aware heading/right + body/full-justify rules.\\n"
            + $"Elapsed seconds: {elapsed.TotalSeconds:0.0}\\n"
            + $"Integrity gate: {(approved ? "PASS" : "FAIL")}\\n",
            new UTF8Encoding(false));

        Console.WriteLine("Generic output: " + result.OutputFile);
        Console.WriteLine("Generic report: " + reportPath);
        Console.WriteLine($"Persian paragraphs: {result.PersianParagraphs} | Paragraphs changed: {result.ParagraphsChanged} | Body paragraphs justified: {result.BodyParagraphsJustified} | Runs changed: {result.RunsChanged}");
        Console.WriteLine($"FAST_ELAPSED_SECONDS={elapsed.TotalSeconds:0.0}");
        Console.WriteLine("Integrity gate: " + (approved ? "PASS" : "FAIL"));
        if (!approved)
        {
            foreach (var issue in issues.Take(50)) Console.WriteLine("BLOCK: " + issue);
            Console.WriteLine("OUTPUT STATUS: NOT APPROVED — do not use this copy.");
            return 7;
        }
        var genericTocMaterialized = PersianDocumentFoundationService.MaterializeMainAutomaticTocAfterIntegrity(result.OutputFile);
        var genericExactInputRestored = RestoreExactInputWhenPackageContentEquivalent(path, result.OutputFile);
        Console.WriteLine("GENERIC_POST_INTEGRITY_NATIVE_TOC_MATERIALIZED=" + genericTocMaterialized);
        Console.WriteLine("GENERIC_POST_INTEGRITY_WORD_SAVE_USED=NO");
        Console.WriteLine("GENERIC_EXACT_INPUT_RESTORED=" + genericExactInputRestored);
        Console.WriteLine("OUTPUT STATUS: APPROVED_FOR_VISUAL_REVIEW");
        return 0;
    }

    private static int AcademicizeThesis(Dictionary<string, string?> options)
    {
        if (!IsTrue(options, "confirm-academicize"))
            throw new InvalidOperationException("Academic thesis transformation requires --confirm-academicize.");

        var input = Require(options, "input");
        var path = PathPolicy.RequireSupportedWordFile(input);
        var outDir = options.GetValueOrDefault("outdir");
        var requestedProfile = options.GetValueOrDefault("profile");
        var explicitProfile = !string.IsNullOrWhiteSpace(requestedProfile);
        var selectedProfileId = explicitProfile
            ? requestedProfile!.Trim()
            : "thesis.dissertation.phd.fa";

        var isGenericThesis = selectedProfileId.Equals(
            "thesis-dissertation-academic-fa",
            StringComparison.OrdinalIgnoreCase);
        var isPhdThesis = selectedProfileId.Equals(
            "thesis.dissertation.phd.fa",
            StringComparison.OrdinalIgnoreCase);
        if (!isGenericThesis && !isPhdThesis)
        {
            throw new InvalidOperationException(
                "academicize-thesis only accepts --profile thesis-dissertation-academic-fa or thesis.dissertation.phd.fa.");
        }

        var profilesDir = ResolveProfilesDir(options);
        var profile = ProfileLoader.Load(profilesDir, selectedProfileId);
        if (!profile.Id.Equals(selectedProfileId, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Selected thesis profile authority could not be resolved exactly.");

        Console.WriteLine("ACADEMIC_TARGET=THESIS_DISSERTATION");
        Console.WriteLine("ACADEMIC_SELECTED_PROFILE=" + selectedProfileId);
        Console.WriteLine("ACADEMIC_PROFILE_SELECTION=" + (explicitProfile ? "EXPLICIT" : "DEFAULT_BACKWARD_COMPAT_PHD"));
        Console.WriteLine("SOURCE_PROFILE_DETECTION=BYPASSED_BY_EXPLICIT_USER_TARGET");
        Console.WriteLine("ORIGINAL_OVERWRITE=NO");

        var result = new ThesisAcademicizeService().Apply(path, selectedProfileId, outDir);
        Console.WriteLine("THESIS_PROFILE_APPLIED=" + result.SelectedProfileId);
        Console.WriteLine("THESIS_ROUTE_POLICY=" + result.RoutePolicy);
        Console.WriteLine("THESIS_OUTPUT=" + result.OutputFile);
        Console.WriteLine("TOC_INSERTED=" + result.TocInserted);
        Console.WriteLine("HEADING_PARAGRAPHS_RIGHT_ALIGNED=" + result.HeadingParagraphsRightAligned);
        Console.WriteLine("BODY_PARAGRAPHS_JUSTIFIED=" + result.BodyParagraphsJustified);
        Console.WriteLine($"TABLES={result.Tables};TABLE_CAPTIONS_INSERTED={result.TableCaptionsInserted}");
        Console.WriteLine($"DRAWINGS={result.Drawings};FIGURE_CAPTIONS_INSERTED={result.FigureCaptionsInserted};CHART_CAPTIONS_INSERTED={result.ChartCaptionsInserted}");
        Console.WriteLine("OUTPUT STATUS: APPROVED_FOR_VISUAL_REVIEW");
        return 0;
    }

    private static string FormatGenericPersianizeReport(
        GenericPersianizeResult result,
        WordAuditSnapshot beforeWord,
        WordAuditSnapshot afterWord,
        OpenXmlAuditSnapshot beforeOpenXml,
        OpenXmlAuditSnapshot afterOpenXml,
        IReadOnlyList<string> issues,
        bool approved)
    {
        var sb = new StringBuilder();
        sb.AppendLine("گزارش اصلاح عمومی RTL وردایجنت");
        sb.AppendLine(new string('=', 60));
        sb.AppendLine("Schema: wordagent.persianize.generic/1.0");
        sb.AppendLine("Input: " + result.InputFile);
        sb.AppendLine("Output: " + result.OutputFile);
        sb.AppendLine("Input SHA-256: " + result.InputSha256);
        sb.AppendLine("Output SHA-256: " + result.OutputSha256);
        sb.AppendLine("Original overwrite: NO");
        sb.AppendLine("Profile detection: NOT USED");
        sb.AppendLine("Document-family formatting: NOT USED");
        sb.AppendLine();

        sb.AppendLine("[دامنه تغییر عمومی]");
        sb.AppendLine($"  Persian paragraphs found: {result.PersianParagraphs}/{result.ParagraphsExamined}");
        sb.AppendLine($"  Paragraph direction/alignment changes: {result.ParagraphsChanged}");
        sb.AppendLine($"  Persian run RTL changes: {result.RunsChanged}");
        sb.AppendLine($"  Tables examined: {result.TablesExamined}");
        sb.AppendLine($"  Tables centered: {result.TablesCentered}");
        sb.AppendLine($"  Persian tables RTL-enabled: {result.PersianTablesRtlEnabled}");
        sb.AppendLine($"  TOC styles changed: {result.TocStylesChanged}");
        sb.AppendLine($"  Automatic field update on open disabled: {result.AutomaticFieldUpdateDisabled}");
        sb.AppendLine($"  No changes needed: {result.NoChangesNeeded}");
        sb.AppendLine();

        sb.AppendLine("[مواردی که عمداً تغییر نمی‌کنند]");
        sb.AppendLine("  Fonts / font sizes / margins / sections / columns / numbering / page breaks / table widths / borders / shading.");
        sb.AppendLine("  Text content / field instructions / images / relationships / comments / bookmarks / hyperlinks.");
        sb.AppendLine("  TOC field content is preserved; only generic RTL style properties may change.");
        sb.AppendLine();

        sb.AppendLine("[گیت‌های Integrity]");
        sb.AppendLine($"  Word paragraphs: {beforeWord.Counts.Paragraphs} -> {afterWord.Counts.Paragraphs}");
        sb.AppendLine($"  Word tables: {beforeWord.Counts.Tables} -> {afterWord.Counts.Tables}");
        sb.AppendLine($"  Word fields: {beforeWord.Counts.Fields} -> {afterWord.Counts.Fields}");
        sb.AppendLine($"  TOC fields: {beforeWord.Counts.TocFields} -> {afterWord.Counts.TocFields}");
        sb.AppendLine($"  Sections: {beforeWord.Counts.Sections} -> {afterWord.Counts.Sections}");
        sb.AppendLine($"  OpenXML validation errors: {beforeOpenXml.ValidationErrors} -> {afterOpenXml.ValidationErrors}");
        sb.AppendLine($"  Text payload SHA preserved: {result.Before.TextPayloadSha256.Equals(result.After.TextPayloadSha256, StringComparison.OrdinalIgnoreCase)}");
        sb.AppendLine($"  Field instruction SHA preserved: {result.Before.FieldInstructionSha256.Equals(result.After.FieldInstructionSha256, StringComparison.OrdinalIgnoreCase)}");
        sb.AppendLine();

        if (result.Warnings.Count > 0)
        {
            sb.AppendLine("[قواعد / یادداشت‌ها]");
            foreach (var warning in result.Warnings) sb.AppendLine("  - " + warning);
            sb.AppendLine();
        }

        if (issues.Count > 0)
        {
            sb.AppendLine("[موارد مسدودکننده]");
            foreach (var issue in issues) sb.AppendLine("  - " + issue);
            sb.AppendLine();
        }

        sb.AppendLine("[وضعیت]");
        sb.AppendLine("  Integrity gate: " + (approved ? "PASS" : "FAIL"));
        sb.AppendLine("  Output status: " + (approved ? "APPROVED_FOR_VISUAL_REVIEW" : "NOT_APPROVED"));
        sb.AppendLine("  Visual review in Microsoft Word is still required for rendered pagination and appearance.");
        return FinalizeFaReport(sb);
    }

    private static int LegacyFixDisabled()
    {
        Console.Error.WriteLine("Legacy 'fix' mutation remains disabled in m3.2.");
        Console.Error.WriteLine("Review the READ-ONLY plan, then use: fix-controlled --confirm-plan-ready");
        return 2;
    }


    private static int InternalWordAuditSnapshot(Dictionary<string, string?> options)
    {
        var input = Require(options, "input");
        var outputJson = Require(options, "output-json");
        var visible = IsTrue(options, "visible");
        var path = PathPolicy.RequireSupportedWordFile(input);

        return WriteInternalWordSnapshot(
            "WORD_AUDIT_WORKER",
            outputJson,
            () => new WordAuditService().Inspect(path, visible));
    }

    private static int InternalProfileComplianceSnapshot(Dictionary<string, string?> options)
    {
        var input = Require(options, "input");
        var outputJson = Require(options, "output-json");
        var visible = IsTrue(options, "visible");
        var profileId = options.GetValueOrDefault("profile") ?? DefaultProfile;
        var profilesDir = ResolveProfilesDir(options);
        var path = PathPolicy.RequireSupportedWordFile(input);
        var profile = ProfileLoader.Load(profilesDir, profileId);

        return WriteInternalWordSnapshot(
            "PROFILE_COMPLIANCE_WORKER",
            outputJson,
            () => new ProfileComplianceService().Inspect(path, profile, visible));
    }

    private static int WriteInternalWordSnapshot<T>(
        string label,
        string outputJson,
        Func<T> producer)
    {
        var preexisting = SnapshotWinWordPids();
        if (preexisting.Count != 0)
        {
            throw new InvalidOperationException(
                $"{label}: INTERNAL_WORD_SESSION_GUARD blocked because WINWORD already exists. "
                + $"PIDS={string.Join(",", preexisting.OrderBy(x => x))}. No process was killed.");
        }

        var snapshot = producer();
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputJson))!);
        File.WriteAllText(
            outputJson,
            JsonSerializer.Serialize(snapshot, JsonOptions),
            new UTF8Encoding(false));

        ForceComFinalization();

        var deadline = DateTime.UtcNow.AddSeconds(30);
        int[] remaining;
        do
        {
            remaining = SnapshotWinWordPids()
                .Where(pid => !preexisting.Contains(pid))
                .OrderBy(pid => pid)
                .ToArray();
            if (remaining.Length == 0)
                break;
            Thread.Sleep(250);
        }
        while (DateTime.UtcNow < deadline);

        if (remaining.Length != 0)
        {
            throw new InvalidOperationException(
                $"{label}: INTERNAL_WORD_ORPHAN_GUARD blocked because WINWORD remained after COM finalization. "
                + $"REMAINING={string.Join(",", remaining)}. No process was killed automatically.");
        }

        Console.WriteLine($"INTERNAL_WORD_ORPHAN_GUARD:PASS:REMAINING=0:LABEL={label}");
        return 0;
    }

    private static void ForceComFinalization()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    // R5R2 compile integration:
    // Public command paths use these wrappers; worker commands and RunIsolatedSnapshot<T>
    // are already defined below. Keep all Word COM inspection in the isolated helper process.
    private static WordAuditSnapshot InspectWordIsolated(string path, bool visible)
    {
        var args = new List<string>
        {
            "__internal-word-audit-snapshot",
            "--input",
            path
        };
        if (visible)
            args.Add("--visible");

        return RunIsolatedSnapshot<WordAuditSnapshot>(
            "WORD_AUDIT",
            args);
    }

    private static ProfileComplianceSnapshot InspectComplianceIsolated(
        string path,
        string profileId,
        string profilesDir,
        bool visible)
    {
        var args = new List<string>
        {
            "__internal-profile-compliance-snapshot",
            "--input",
            path,
            "--profile",
            profileId,
            "--profiles",
            profilesDir
        };
        if (visible)
            args.Add("--visible");

        return RunIsolatedSnapshot<ProfileComplianceSnapshot>(
            "PROFILE_COMPLIANCE",
            args,
            timeoutMs: 120000);
    }

    // QH01-E239:
    // Large-grain helper timeout recovery. The parent owns the helper process it starts,
    // so on timeout it may terminate ONLY that helper process. It never kills WINWORD.
    // After helper termination it waits for helper-created WINWORD to exit naturally.
    // If the Word set returns to zero, one fresh isolated retry is allowed.
    // QH01-E240: compile-safe timeout/orphan PID arrays use distinct local names.
    // QH01-E239/QH01-E240/QH01-E241:
    // Large-grain isolated Word helper recovery.
    // - Parent may terminate only the helper process it created.
    // - If zero WINWORD existed before helper start and exactly one new WINWORD appeared
    //   during that helper attempt, that exact PID may also be terminated after the
    //   helper is gone, but only after start-time/process-set ownership is re-verified.
    // - User/pre-existing Word is never terminated.
    // - One fresh retry is allowed only after the Word set returns to zero.
    private static T RunIsolatedSnapshot<T>(
        string label,
        IReadOnlyList<string> argumentVector,
        int timeoutMs = 180000)
    {
        var executable = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(executable))
            throw new InvalidOperationException($"{label}: current Engine executable path is unavailable.");

        const int maxAttempts = 2;

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            var preexistingWordPids = SnapshotWinWordPids();
            if (preexistingWordPids.Count != 0)
            {
                throw new InvalidOperationException(
                    $"{label}: WORD_COM_SESSION_GUARD blocked because WINWORD already exists. "
                    + $"PIDS={string.Join(",", preexistingWordPids.OrderBy(x => x))}. "
                    + "Close Word yourself; WordAgent will never terminate pre-existing Word.");
            }

            Console.WriteLine(
                $"WORD_COM_SESSION_GUARD:PASS:PREEXISTING=0:LABEL={label}:ATTEMPT={attempt}/{maxAttempts}");

            var outputJson = Path.Combine(
                Path.GetTempPath(),
                $"WordAgent_{label}_{attempt}_{Guid.NewGuid():N}.json");

            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = executable,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            if (string.Equals(
                    Path.GetFileNameWithoutExtension(executable),
                    "dotnet",
                    StringComparison.OrdinalIgnoreCase))
            {
                var managedEntryAssembly = Path.Combine(AppContext.BaseDirectory, "WordAgent.dll");
                if (!File.Exists(managedEntryAssembly))
                    throw new InvalidOperationException(
                        $"{label}: framework-dependent dotnet host detected, but managed entry assembly is unavailable at {managedEntryAssembly}.");

                psi.ArgumentList.Add(managedEntryAssembly);
            }

            foreach (var value in argumentVector)
                psi.ArgumentList.Add(value);
            psi.ArgumentList.Add("--output-json");
            psi.ArgumentList.Add(outputJson);

            using var process = new System.Diagnostics.Process { StartInfo = psi };

            try
            {
                if (!process.Start())
                    throw new InvalidOperationException($"{label}: helper process could not start.");

                var helperPid = process.Id;
                var helperStartedUtc = DateTime.UtcNow;
                var stdoutTask = process.StandardOutput.ReadToEndAsync();
                var stderrTask = process.StandardError.ReadToEndAsync();

                if (!process.WaitForExit(timeoutMs))
                {
                    var timedOutStdout = stdoutTask.IsCompletedSuccessfully ? stdoutTask.Result : string.Empty;
                    var timedOutStderr = stderrTask.IsCompletedSuccessfully ? stderrTask.Result : string.Empty;
                    var observedNewWordPids = SnapshotWinWordPids()
                        .Where(pid => !preexistingWordPids.Contains(pid))
                        .OrderBy(pid => pid)
                        .ToArray();

                    Console.WriteLine(
                        $"WORD_COM_HELPER_TIMEOUT_RECOVERY:START:LABEL={label}:ATTEMPT={attempt}/{maxAttempts}:"
                        + $"HELPER_PID={helperPid}:NEW_WINWORD_PIDS={string.Join(",", observedNewWordPids)}");

                    try
                    {
                        if (!process.HasExited)
                            process.Kill();
                    }
                    catch (Exception killEx)
                    {
                        if (!process.HasExited)
                            throw new InvalidOperationException(
                                $"{label}: owned helper could not be terminated after timeout. "
                                + $"HELPER_PID={helperPid}; ERROR={killEx.Message}.",
                                killEx);
                    }

                    if (!process.HasExited && !process.WaitForExit(15000))
                    {
                        throw new TimeoutException(
                            $"{label}: owned helper remained alive after termination request. "
                            + $"HELPER_PID={helperPid}.");
                    }

                    try { process.WaitForExit(); } catch { }
                    try { _ = stdoutTask.GetAwaiter().GetResult(); } catch { }
                    try { _ = stderrTask.GetAwaiter().GetResult(); } catch { }

                    var timeoutRemainingWordPids = SnapshotWinWordPids()
                        .Where(pid => !preexistingWordPids.Contains(pid))
                        .OrderBy(pid => pid)
                        .ToArray();

                    if (timeoutRemainingWordPids.Length != 0)
                    {
                        TerminateVerifiedHelperOwnedWinWord(
                            label,
                            helperPid,
                            helperStartedUtc,
                            preexistingWordPids,
                            observedNewWordPids);
                    }

                    var cleanupDeadline = DateTime.UtcNow.AddSeconds(15);
                    do
                    {
                        timeoutRemainingWordPids = SnapshotWinWordPids()
                            .Where(pid => !preexistingWordPids.Contains(pid))
                            .OrderBy(pid => pid)
                            .ToArray();

                        if (timeoutRemainingWordPids.Length == 0)
                            break;

                        Thread.Sleep(250);
                    }
                    while (DateTime.UtcNow < cleanupDeadline);

                    if (timeoutRemainingWordPids.Length != 0)
                    {
                        throw new TimeoutException(
                            $"{label}: WORD_COM_HELPER_TIMEOUT_RECOVERY_BLOCKED. "
                            + $"ATTEMPT={attempt}/{maxAttempts}; HELPER_PID={helperPid}; "
                            + $"REMAINING_WINWORD_PIDS={string.Join(",", timeoutRemainingWordPids)}. "
                            + "No pre-existing/user Word was terminated. "
                            + $"stdout={TrimDiagnostic(timedOutStdout)} stderr={TrimDiagnostic(timedOutStderr)}");
                    }

                    Console.WriteLine(
                        $"WORD_COM_HELPER_TIMEOUT_RECOVERY:PASS:LABEL={label}:ATTEMPT={attempt}/{maxAttempts}:"
                        + $"HELPER_PID={helperPid}:REMAINING_WINWORD=0");

                    if (attempt < maxAttempts)
                    {
                        Console.WriteLine(
                            $"WORD_COM_HELPER_RETRY:START:LABEL={label}:NEXT_ATTEMPT={attempt + 1}/{maxAttempts}");
                        continue;
                    }

                    throw new TimeoutException(
                        $"{label}: WORD_COM_HELPER_TIMEOUT_AFTER_RETRY after {maxAttempts} attempts "
                        + $"of {timeoutMs} ms. "
                        + $"stdout={TrimDiagnostic(timedOutStdout)} stderr={TrimDiagnostic(timedOutStderr)}");
                }

                process.WaitForExit();
                var stdout = stdoutTask.GetAwaiter().GetResult();
                var stderr = stderrTask.GetAwaiter().GetResult();

                var orphanDeadline = DateTime.UtcNow.AddSeconds(20);
                int[] orphanRemainingWordPids;
                do
                {
                    orphanRemainingWordPids = SnapshotWinWordPids()
                        .Where(pid => !preexistingWordPids.Contains(pid))
                        .OrderBy(pid => pid)
                        .ToArray();

                    if (orphanRemainingWordPids.Length == 0)
                        break;

                    Thread.Sleep(250);
                }
                while (DateTime.UtcNow < orphanDeadline);

                if (orphanRemainingWordPids.Length != 0)
                {
                    TerminateVerifiedHelperOwnedWinWord(
                        label,
                        helperPid,
                        helperStartedUtc,
                        preexistingWordPids,
                        orphanRemainingWordPids);

                    orphanRemainingWordPids = SnapshotWinWordPids()
                        .Where(pid => !preexistingWordPids.Contains(pid))
                        .OrderBy(pid => pid)
                        .ToArray();
                }

                if (orphanRemainingWordPids.Length != 0)
                {
                    throw new InvalidOperationException(
                        $"{label}: WORD_COM_ORPHAN_GUARD blocked because helper left WINWORD running. "
                        + $"REMAINING={string.Join(",", orphanRemainingWordPids)}. "
                        + "No pre-existing/user Word was terminated.");
                }

                Console.WriteLine(
                    $"WORD_COM_ORPHAN_GUARD:PASS:REMAINING=0:LABEL={label}:ATTEMPT={attempt}/{maxAttempts}");

                if (process.ExitCode != 0)
                {
                    var internalOrphanAfterSuccessfulSnapshot =
                        stderr.Contains(
                            $"{label}: INTERNAL_WORD_ORPHAN_GUARD blocked because WINWORD remained after COM finalization.",
                            StringComparison.OrdinalIgnoreCase)
                        && File.Exists(outputJson);

                    if (internalOrphanAfterSuccessfulSnapshot)
                    {
                        var salvagedJson = File.ReadAllText(outputJson, Encoding.UTF8);
                        var salvagedSnapshot = JsonSerializer.Deserialize<T>(salvagedJson, JsonOptions);
                        if (salvagedSnapshot is not null)
                        {
                            Console.WriteLine(
                                $"WORD_COM_HELPER_ORPHAN_SNAPSHOT_SALVAGE:PASS:LABEL={label}:ATTEMPT={attempt}/{maxAttempts}:"
                                + "OUTPUT_JSON=VALID:OUTER_ORPHAN_GUARD=PASS:REMAINING_WINWORD=0");
                            return salvagedSnapshot;
                        }

                        throw new InvalidOperationException(
                            $"{label}: internal orphan cleanup failed after a snapshot file was written, "
                            + "but the snapshot JSON could not be deserialized safely.");
                    }

                    if (attempt < maxAttempts
                        && stderr.Contains("INTERNAL_WORD_ORPHAN_GUARD", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine(
                            $"WORD_COM_HELPER_RETRY:START:LABEL={label}:NEXT_ATTEMPT={attempt + 1}/{maxAttempts}:"
                            + "REASON=INTERNAL_WORD_ORPHAN_GUARD_WITHOUT_VALID_SNAPSHOT");
                        continue;
                    }

                    throw new InvalidOperationException(
                        $"{label} helper exit={process.ExitCode}. "
                        + $"stdout={TrimDiagnostic(stdout)} stderr={TrimDiagnostic(stderr)}");
                }

                if (!File.Exists(outputJson))
                    throw new InvalidOperationException($"{label}: helper snapshot JSON was not created.");

                var json = File.ReadAllText(outputJson, Encoding.UTF8);
                var snapshot = JsonSerializer.Deserialize<T>(json, JsonOptions);
                return snapshot
                    ?? throw new InvalidOperationException(
                        $"{label}: helper snapshot JSON deserialized to null.");
            }
            finally
            {
                try
                {
                    if (File.Exists(outputJson))
                        File.Delete(outputJson);
                }
                catch { }
            }
        }

        throw new InvalidOperationException($"{label}: isolated helper attempts exhausted unexpectedly.");
    }

    private static void TerminateVerifiedHelperOwnedWinWord(
        string label,
        int helperPid,
        DateTime helperStartedUtc,
        HashSet<int> preexistingWordPids,
        IReadOnlyCollection<int> observedNewWordPids)
    {
        if (preexistingWordPids.Count != 0)
        {
            throw new InvalidOperationException(
                $"{label}: verified-owned Word cleanup refused because pre-existing WINWORD was not zero.");
        }

        var observed = observedNewWordPids
            .Distinct()
            .OrderBy(pid => pid)
            .ToArray();

        var currentNew = SnapshotWinWordPids()
            .Where(pid => !preexistingWordPids.Contains(pid))
            .OrderBy(pid => pid)
            .ToArray();

        if (currentNew.Length == 0)
            return;

        if (observed.Length != 1
            || currentNew.Length != 1
            || observed[0] != currentNew[0])
        {
            throw new InvalidOperationException(
                $"{label}: VERIFIED_HELPER_OWNED_WINWORD_CLEANUP_AMBIGUOUS. "
                + $"HELPER_PID={helperPid}; OBSERVED={string.Join(",", observed)}; "
                + $"CURRENT={string.Join(",", currentNew)}. No WINWORD was terminated.");
        }

        var ownedWordPid = currentNew[0];

        try
        {
            using var ownedWord = System.Diagnostics.Process.GetProcessById(ownedWordPid);

            if (!ownedWord.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"{label}: verified-owned cleanup PID {ownedWordPid} is not WINWORD.");
            }

            DateTime wordStartedUtc;
            try
            {
                wordStartedUtc = ownedWord.StartTime.ToUniversalTime();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"{label}: cannot verify start time for candidate helper-owned WINWORD PID={ownedWordPid}. "
                    + "No WINWORD was terminated.",
                    ex);
            }

            if (wordStartedUtc < helperStartedUtc.AddSeconds(-3))
            {
                throw new InvalidOperationException(
                    $"{label}: candidate WINWORD predates helper ownership window. "
                    + $"HELPER_PID={helperPid}; WINWORD_PID={ownedWordPid}; "
                    + $"WORD_START={wordStartedUtc:o}; HELPER_START={helperStartedUtc:o}. "
                    + "No WINWORD was terminated.");
            }

            Console.WriteLine(
                $"WORD_COM_HELPER_OWNED_WINWORD_RECOVERY:START:LABEL={label}:"
                + $"HELPER_PID={helperPid}:WINWORD_PID={ownedWordPid}:"
                + $"OWNERSHIP=ZERO_BASELINE_UNIQUE_NEW_PID_STARTTIME_VERIFIED");

            ownedWord.Kill();

            if (!ownedWord.WaitForExit(15000))
            {
                throw new TimeoutException(
                    $"{label}: verified helper-owned WINWORD did not exit after termination. "
                    + $"WINWORD_PID={ownedWordPid}.");
            }

            Console.WriteLine(
                $"WORD_COM_HELPER_OWNED_WINWORD_RECOVERY:PASS:LABEL={label}:"
                + $"HELPER_PID={helperPid}:WINWORD_PID={ownedWordPid}");
        }
        catch (ArgumentException)
        {
            // Process exited between verification snapshots.
        }
    }

    private static HashSet<int> SnapshotWinWordPids()
    {
        var result = new HashSet<int>();
        System.Diagnostics.Process[] processes;
        try { processes = System.Diagnostics.Process.GetProcessesByName("WINWORD"); }
        catch { return result; }

        foreach (var process in processes)
        {
            try { result.Add(process.Id); }
            catch { }
            finally { process.Dispose(); }
        }

        return result;
    }

    private static string TrimDiagnostic(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "<empty>";

        var normalized = value.Replace("\r", " ").Replace("\n", " ").Trim();
        return normalized.Length <= 1200 ? normalized : normalized[..1200] + "...";
    }

    private static void DeleteUnapprovedOutput(string path)
    {
        for (var attempt = 1; attempt <= 4; attempt++)
        {
            try
            {
                if (!File.Exists(path))
                    return;
                File.Delete(path);
                return;
            }
            catch
            {
                if (attempt == 4)
                    return;
                Thread.Sleep(250 * attempt);
            }
        }
    }

    private static int ControlledFix(Dictionary<string, string?> options)
    {
        var reviewedPlanConfirmed = IsTrue(options, "confirm-plan-ready");
        var integratedSmartConfirmed = IsTrue(options, "confirm-smart-repair");
        if (!reviewedPlanConfirmed && !integratedSmartConfirmed)
            throw new InvalidOperationException(
                "Controlled fix requires either --confirm-plan-ready after reviewing the READ-ONLY plan "
                + "or --confirm-smart-repair from the integrated fail-closed Smart Repair workflow.");

        if (integratedSmartConfirmed)
            Console.WriteLine("SMART_REPAIR_MODE=INTEGRATED_FAST_FAIL_CLOSED");

        var input = Require(options, "input");
        var path = PathPolicy.RequireSupportedWordFile(input);
        var profilesDir = ResolveProfilesDir(options);
        var profileId = ResolveProfileId(options, path, profilesDir);
        var outDir = options.GetValueOrDefault("outdir");
        var visible = IsTrue(options, "visible");

        var probe = WordRuntime.Probe();
        if (!probe.Installed)
            throw new InvalidOperationException("Microsoft Word Desktop is required for Controlled Copy Fix. " + probe.Error);

        var profile = ProfileLoader.Load(profilesDir, profileId);

        OpenXmlAuditSnapshot beforeOpenXml;
        try { beforeOpenXml = new OpenXmlInspector().Inspect(path); }
        catch (Exception ex) { throw new InvalidOperationException("Pre-fix OpenXML audit failed; mutation blocked. " + ex.Message, ex); }

        Console.WriteLine("Mutation engine: OpenXML direct package edit; Word COM is read-only for audit and write-enabled only for the guarded TOC refresh on the controlled copy. Automatic TOC direction compliance is verified from persisted OOXML + TOC style contract.");

        var tocAuthorityBefore = TocOpenXmlComplianceAuthorityService.Inspect(path, profile);
        var tocAbsentBefore =
            tocAuthorityBefore.StandardTocControls == 0
            && tocAuthorityBefore.Paragraphs == 0;
        // QH01-E225 / R10R45R2:
        // This is provisional profile-level reconstruction intent. The exact source-hierarchy
        // semantic authority is owned by ControlledCopyFixService.PrepareAndMutate and may
        // reconcile this state through ControlledMutationTransaction before downstream gates.
        // AR9 adds UniversityBook only; thesis profiles remain on their specialized academic pipeline.
        var tocReconstructionRequested =
            (string.Equals(profile.Id, "proposal.governance.fa", StringComparison.OrdinalIgnoreCase)
             || string.Equals(profile.Id, "book.university.fa", StringComparison.OrdinalIgnoreCase))
            && tocAbsentBefore;
        var tocNotApplicableWhenAbsent =
            tocAbsentBefore
            && !tocReconstructionRequested
            && !profile.Id.Equals("book.university.fa", StringComparison.OrdinalIgnoreCase)
            && !profile.Id.Equals("thesis.dissertation.phd.fa", StringComparison.OrdinalIgnoreCase)
            && !profile.Id.Equals("thesis-dissertation-academic-fa", StringComparison.OrdinalIgnoreCase);

        if (tocAbsentBefore &&
            !tocReconstructionRequested &&
            !tocNotApplicableWhenAbsent)
        {
            throw new InvalidOperationException(
                $"TOC_CREATE_REQUIRED_FOR_PROFILE:{profile.Id}. This profile requires a structural TOC creation pipeline when no TOC exists.");
        }

        if (!tocNotApplicableWhenAbsent &&
            !tocAuthorityBefore.RepairableVisualLayout &&
            !tocReconstructionRequested)
        {
            var issueText = tocAuthorityBefore.Issues.Count == 0
                ? $"controls={tocAuthorityBefore.StandardTocControls}; paragraphs={tocAuthorityBefore.Paragraphs}; mapped={tocAuthorityBefore.MappedParagraphs}; reading={tocAuthorityBefore.ReadingOrderMatches}; alignment={tocAuthorityBefore.AlignmentMatches}; runRtl={tocAuthorityBefore.TextRunsRtlMatches}"
                : string.Join(" | ", tocAuthorityBefore.Issues.Take(20));

            throw new InvalidOperationException(
                "Automatic TOC OpenXML authority preflight is not safely repairable: " + issueText);
        }

        if (tocReconstructionRequested)
        {
            Console.WriteLine(
                "CONTROLLED_STAGE=TOC_OPENXML_AUTHORITY_PRE:RECONSTRUCT_REQUIRED:"
                + "REASON=Selected profile requires semantic controlled-copy TOC reconstruction.");
        }
        else if (tocNotApplicableWhenAbsent)
        {
            Console.WriteLine(
                "CONTROLLED_STAGE=TOC_OPENXML_AUTHORITY_PRE:NOT_APPLICABLE:"
                + $"PROFILE={profile.Id}:REASON=No automatic TOC exists; absence is not a failure for this profile.");
        }
        else
        {
            Console.WriteLine(
                $"CONTROLLED_STAGE=TOC_OPENXML_AUTHORITY_PRE:REPAIRABLE:CONTROLS={tocAuthorityBefore.StandardTocControls}:PARAGRAPHS={tocAuthorityBefore.Paragraphs}:MAPPED={tocAuthorityBefore.MappedParagraphs}:READING={tocAuthorityBefore.ReadingOrderMatches}:ALIGNMENT={tocAuthorityBefore.AlignmentMatches}:RUNRTL={tocAuthorityBefore.TextRunsRtlMatches}");
        }

        Console.WriteLine("CONTROLLED_STAGE=TRANSACTION_PREPARE:START");

        var service = new ControlledCopyFixService();
        ControlledMutationTransaction transaction;
        try
        {
            transaction = service.PrepareAndMutate(
                path,
                profile,
                outDir,
                visible,
                wordInspector: p => InspectWordIsolated(p, visible),
                complianceInspector: (p, _) => InspectComplianceIsolated(p, profileId, profilesDir, visible),
                trace: message => Console.WriteLine("CONTROLLED_STAGE=" + message));
        }
        catch
        {
            Console.WriteLine("CONTROLLED_STAGE=TRANSACTION_PREPARE:FAIL");
            throw;
        }

        // R10R45R2: ControlledCopyFixService owns the semantic source-hierarchy decision.
        // Propagate its exact no-source-hierarchy N/A result through the transaction boundary
        // instead of letting CLI post-gates re-infer authority from profile identity.
        if (transaction.MissingRequiredTocNotApplicableNoSourceHierarchy)
        {
            if (!tocAbsentBefore || !tocReconstructionRequested || tocNotApplicableWhenAbsent)
            {
                DeleteUnapprovedOutput(transaction.OutputFile);
                Console.WriteLine("CONTROLLED_STAGE=TOC_AUTHORITY_CONTRACT:FAIL:ILLEGAL_NA_PROPAGATION");
                throw new InvalidOperationException(
                    $"TOC N/A transaction authority invariant failed: absentBefore={tocAbsentBefore}; reconstructionRequested={tocReconstructionRequested}; priorNotApplicable={tocNotApplicableWhenAbsent}; profile={profile.Id}.");
            }

            tocReconstructionRequested = false;
            tocNotApplicableWhenAbsent = true;
            Console.WriteLine(
                "CONTROLLED_STAGE=TOC_AUTHORITY_CONTRACT:NOT_APPLICABLE_NO_SOURCE_HIERARCHY:PASS:"
                + $"SOURCE=TRANSACTION:PROFILE={profile.Id}");
        }

        Console.WriteLine("CONTROLLED_STAGE=POST_OOXML_GATE:START");
        OpenXmlAuditSnapshot afterOpenXml;
        try
        {
            afterOpenXml = new OpenXmlInspector().Inspect(transaction.OutputFile);
        }
        catch (Exception ex)
        {
            DeleteUnapprovedOutput(transaction.OutputFile);
            Console.WriteLine("CONTROLLED_STAGE=POST_OOXML_GATE:FAIL:OUTPUT_DISCARDED");
            throw new InvalidOperationException("Post-mutation OpenXML audit failed before Word verification. Output was discarded. " + ex.Message, ex);
        }

        var tocReconstructionBaseline =
            transaction.OutputFile + ".WordAgent_TocReconstructionBaseline.docx";
        var integrityBeforePath = File.Exists(tocReconstructionBaseline)
            ? tocReconstructionBaseline
            : path;
        var integrityBeforeOpenXml = File.Exists(tocReconstructionBaseline)
            ? new OpenXmlInspector().Inspect(tocReconstructionBaseline)
            : beforeOpenXml;

        var preWordOpenXmlIssues = CompareOpenXmlIntegrity(
            integrityBeforeOpenXml,
            afterOpenXml,
            integrityBeforePath,
            transaction.OutputFile,
            profileId,
            out var tocLocalDeltaProof,
            out var tocStyleDeltaEvidence);
        Console.WriteLine(
            $"CONTROLLED_STAGE=TOC_LOCAL_OOXML_DELTA:{(tocLocalDeltaProof.Accepted ? "PASS" : "FAIL")}:"
            + $"APPLICABLE={tocLocalDeltaProof.Applicable}:"
            + $"TOTAL_P={tocLocalDeltaProof.TotalParagraphDelta}:TOC_P={tocLocalDeltaProof.TocParagraphDelta}:"
            + $"TOTAL_FI={tocLocalDeltaProof.TotalFieldInstructionDelta}:TOC_FI={tocLocalDeltaProof.TocFieldInstructionDelta}:"
            + $"VISIBLE={tocLocalDeltaProof.BeforeVisibleTocParagraphs}->{tocLocalDeltaProof.AfterVisibleTocParagraphs}:"
            + $"REASON={tocLocalDeltaProof.Reason}");
        Console.WriteLine("CONTROLLED_STAGE=TOC_STYLE_MATERIALIZATION_DELTA:" + tocStyleDeltaEvidence);
        if (preWordOpenXmlIssues.Count != 0)
        {
            DeleteUnapprovedOutput(transaction.OutputFile);
            Console.WriteLine("CONTROLLED_STAGE=POST_OOXML_GATE:FAIL:OUTPUT_DISCARDED");
            throw new InvalidOperationException(
                "Post-mutation OpenXML integrity gate failed before Word verification. Output was discarded. " +
                string.Join(" | ", preWordOpenXmlIssues.Take(12)));
        }
        Console.WriteLine("CONTROLLED_STAGE=POST_OOXML_GATE:PASS");

        Console.WriteLine("CONTROLLED_STAGE=TOC_COM_WRITE_NORMALIZATION:RETIRED");

        var tocAuthorityAfter = TocOpenXmlComplianceAuthorityService.Inspect(transaction.OutputFile, profile);
        if (tocNotApplicableWhenAbsent)
        {
            if (tocAuthorityAfter.StandardTocControls != 0 || tocAuthorityAfter.Paragraphs != 0)
            {
                DeleteUnapprovedOutput(transaction.OutputFile);
                Console.WriteLine("CONTROLLED_STAGE=TOC_OPENXML_AUTHORITY_POST:FAIL:UNEXPECTED_TOC");
                throw new InvalidOperationException(
                    $"Controlled profile path unexpectedly created TOC authority where none existed: controls={tocAuthorityAfter.StandardTocControls}; paragraphs={tocAuthorityAfter.Paragraphs}.");
            }

            Console.WriteLine(
                "CONTROLLED_STAGE=TOC_OPENXML_AUTHORITY_POST:NOT_APPLICABLE:PASS:"
                + $"PROFILE={profile.Id}:CONTROLS=0:PARAGRAPHS=0");
        }
        else
        {
            if (!tocAuthorityAfter.FullyCompliant)
            {
                DeleteUnapprovedOutput(transaction.OutputFile);
                Console.WriteLine("CONTROLLED_STAGE=TOC_OPENXML_AUTHORITY_POST:FAIL:OUTPUT_DISCARDED");
                throw new InvalidOperationException(
                    "Automatic TOC OpenXML authority post-gate failed. Output was discarded. "
                    + string.Join(" | ", tocAuthorityAfter.Issues.Take(20)));
            }

            var tocAuthorityParagraphDelta = tocAuthorityAfter.Paragraphs - tocAuthorityBefore.Paragraphs;
            var tocAuthorityInventoryAccepted = tocReconstructionRequested
                ? tocAuthorityAfter.StandardTocControls == 1
                  && tocAuthorityAfter.Paragraphs > 0
                  && tocLocalDeltaProof.Accepted
                : tocAuthorityAfter.StandardTocControls == tocAuthorityBefore.StandardTocControls
                  && tocAuthorityAfter.Paragraphs > 0
                  && (tocAuthorityParagraphDelta == 0 || tocLocalDeltaProof.Accepted);

            if (!tocAuthorityInventoryAccepted)
            {
                DeleteUnapprovedOutput(transaction.OutputFile);
                Console.WriteLine("CONTROLLED_STAGE=TOC_OPENXML_AUTHORITY_POST:FAIL:INVENTORY");
                throw new InvalidOperationException(
                    $"Automatic TOC authority inventory is not safely explainable: controls {tocAuthorityBefore.StandardTocControls}->{tocAuthorityAfter.StandardTocControls}; semantic paragraphs {tocAuthorityBefore.Paragraphs}->{tocAuthorityAfter.Paragraphs}; raw TOC-local paragraph delta={tocLocalDeltaProof.TocParagraphDelta}; localityAccepted={tocLocalDeltaProof.Accepted}.");
            }

            if (tocAuthorityParagraphDelta != 0)
            {
                Console.WriteLine(
                    $"CONTROLLED_STAGE=TOC_OPENXML_AUTHORITY_POST:COUNT_CHANGE_ACCEPTED:"
                    + $"SEMANTIC={tocAuthorityBefore.Paragraphs}->{tocAuthorityAfter.Paragraphs}:"
                    + $"RAW_TOC_DELTA={tocLocalDeltaProof.TocParagraphDelta}:"
                    + $"REASON=TOC-local scope proof passed; semantic visible-entry count is not required to equal raw OOXML paragraph delta after Word regeneration.");
            }

            Console.WriteLine(
                $"CONTROLLED_STAGE=TOC_OPENXML_AUTHORITY_POST:PASS:CONTROLS={tocAuthorityAfter.StandardTocControls}:PARAGRAPHS={tocAuthorityAfter.Paragraphs}:MAPPED={tocAuthorityAfter.MappedParagraphs}:READING={tocAuthorityAfter.ReadingOrderMatches}:ALIGNMENT={tocAuthorityAfter.AlignmentMatches}:RUNRTL={tocAuthorityAfter.TextRunsRtlMatches}");
        }


        ControlledCopyFixResult result;
        try
        {
            result = service.VerifyTransaction(
                transaction,
                profile,
                visible,
                wordInspector: p => InspectWordIsolated(p, visible),
                complianceInspector: (p, _) => InspectComplianceIsolated(p, profileId, profilesDir, visible),
                trace: message => Console.WriteLine("CONTROLLED_STAGE=" + message));
        }
        catch
        {
            Console.WriteLine("CONTROLLED_STAGE=POST_WORD_VERIFICATION:FAIL");
            throw;
        }

        var openXmlIssues = preWordOpenXmlIssues;
        var allIssues = result.IntegrityIssues.Concat(openXmlIssues).ToArray();
        var approved = result.WordIntegrityPassed
            && result.IntegrityIssues.Count == 0
            && openXmlIssues.Count == 0;

        var outputRoot = Path.GetDirectoryName(result.OutputFile)!;
        var fixDir = Path.Combine(outputRoot, "WordAgent_Fix");
        Directory.CreateDirectory(fixDir);
        var stem = Path.GetFileNameWithoutExtension(result.OutputFile);
        var reportPath = Path.Combine(fixDir, stem + ".controlledfix.fa.txt");
        File.WriteAllText(
            reportPath,
            FormatControlledFixReport(result, beforeOpenXml, afterOpenXml, openXmlIssues, approved),
            new UTF8Encoding(false));

        Console.WriteLine("Controlled output: " + result.OutputFile);
        Console.WriteLine("Fix report: " + reportPath);
        Console.WriteLine($"Candidates: {result.CandidateParagraphs} | Mutated: {result.MutatedParagraphs}");
        Console.WriteLine("Integrity gate: " + (approved ? "PASS" : "FAIL"));

        if (!approved)
        {
            foreach (var issue in allIssues.Take(40))
                Console.WriteLine("BLOCK: " + issue);

            foreach (var role in result.AfterCompliance.Roles
                         .Where(x => x.Total - x.FullyCompliant > 0)
                         .OrderBy(x => x.Role, StringComparer.OrdinalIgnoreCase))
            {
                var remaining = role.Total - role.FullyCompliant;
                Console.WriteLine(
                    $"CONTROLLED_RESIDUAL_ROLE={role.Role}|REMAINING={remaining}|TOTAL={role.Total}|FONT={role.FontBiMatches}|SIZE={role.SizeBiMatches}|READING={role.ReadingOrderMatches}|ALIGNMENT={role.AlignmentMatches}");

                foreach (var sample in role.Samples)
                {
                    Console.WriteLine(
                        $"CONTROLLED_RESIDUAL_SAMPLE=P{sample.ParagraphIndex}|ROLE={role.Role}|STYLE={sample.StyleName}|FONT={sample.ActualFontBi}|SIZE={sample.ActualSizeBi}|READING={sample.ActualReadingOrder}|ALIGNMENT={sample.ActualAlignment}|OVERRIDE={sample.DirectOverrideSuspected}");
                }
            }

            Console.WriteLine("OUTPUT STATUS: NOT APPROVED — do not use this copy.");
            return 5;
        }

        // Reuse the already accepted post-mutation snapshots. Do not reopen Word twice more.
        var auditDir = Path.Combine(outputRoot, "WordAgent_Audit");
        Directory.CreateDirectory(auditDir);
        var auditReport = new CombinedAuditReport(
            "wordagent.audit/1.3",
            DateTimeOffset.Now,
            result.OutputFile,
            result.AfterWord,
            afterOpenXml,
            result.AfterCompliance,
            result.Warnings);
        var auditJsonPath = Path.Combine(auditDir, stem + ".audit.json");
        var auditTextPath = Path.Combine(auditDir, stem + ".audit.fa.txt");
        File.WriteAllText(auditJsonPath, JsonSerializer.Serialize(auditReport, JsonOptions), new UTF8Encoding(false));
        File.WriteAllText(auditTextPath, FormatPersianReport(auditReport), new UTF8Encoding(false));
        Console.WriteLine("Audit JSON: " + auditJsonPath);
        Console.WriteLine("Audit report: " + auditTextPath);

        var planDir = Path.Combine(outputRoot, "WordAgent_Plan");
        Directory.CreateDirectory(planDir);
        var planText = FormatSafeFixPlan(
            result.OutputFile,
            profile,
            result.AfterWord,
            result.AfterCompliance,
            afterOpenXml,
            openXmlWarning: null);
        var planPath = Path.Combine(planDir, stem + ".fixplan.fa.txt");
        File.WriteAllText(planPath, planText, new UTF8Encoding(false));
        Console.WriteLine("Safe-Fix plan: " + planPath);
        Console.WriteLine("READ-ONLY: YES");

        Console.WriteLine("CONTROLLED_STAGE=FOLLOWUP_REPORTS:REUSED_POST_SNAPSHOTS");
        Console.WriteLine("OUTPUT STATUS: APPROVED_FOR_VISUAL_REVIEW");
        return 0;
    }


    private static List<string> CompareOpenXmlIntegrity(
        OpenXmlAuditSnapshot before,
        OpenXmlAuditSnapshot after,
        string beforePath,
        string afterPath,
        string profileId,
        out TocLocalOpenXmlDeltaProof tocProof,
        out string tocStyleDeltaEvidence)
    {
        tocStyleDeltaEvidence = "NOT_APPLICABLE_COUNT_STABLE";
        tocProof = BuildTocLocalOpenXmlDeltaProof(
            beforePath,
            afterPath,
            before,
            after,
            profileId);

        var issues = new List<string>();
        void Check(string name, int a, int b)
        {
            if (a != b) issues.Add($"OpenXML structure changed: {name} {a} -> {b}");
        }

        if (before.Paragraphs != after.Paragraphs && !tocProof.Accepted)
            issues.Add(
                $"OpenXML structure changed outside an accepted TOC-local rewrite: Paragraphs {before.Paragraphs} -> {after.Paragraphs}. {tocProof.Reason}");

        Check("Tables", before.Tables, after.Tables);
        Check("Drawings", before.Drawings, after.Drawings);

        if (before.Styles != after.Styles)
        {
            var tocStyleDeltaAccepted = VerifyTocStyleMaterializationDelta(
                beforePath,
                afterPath,
                tocProof,
                before.Styles,
                after.Styles,
                out tocStyleDeltaEvidence);

            if (!tocStyleDeltaAccepted)
                issues.Add($"OpenXML structure changed: Styles {before.Styles} -> {after.Styles}. {tocStyleDeltaEvidence}");
        }

        if (before.FieldInstructions != after.FieldInstructions && !tocProof.Accepted)
            issues.Add(
                $"OpenXML structure changed outside an accepted TOC-local rewrite: FieldInstructions {before.FieldInstructions} -> {after.FieldInstructions}. {tocProof.Reason}");

        Check("EndNoteMarkers", before.EndNoteMarkers, after.EndNoteMarkers);
        if (after.ValidationErrors > before.ValidationErrors)
            issues.Add($"OpenXML validation errors increased: {before.ValidationErrors} -> {after.ValidationErrors}");
        return issues;
    }

    private sealed record WordStyleIdentity(string Id, string Name);

    private static WordStyleIdentity[] ReadWordStyleIdentities(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry("word/styles.xml")
            ?? throw new InvalidOperationException("word/styles.xml is missing.");

        using var stream = entry.Open();
        var styles = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        return styles
            .Descendants(WordMain + "style")
            .Select(style => new WordStyleIdentity(
                ((string?)style.Attribute(WordMain + "styleId") ?? string.Empty).Trim(),
                ((string?)style.Element(WordMain + "name")?.Attribute(WordMain + "val") ?? string.Empty).Trim()))
            .Where(style => style.Id.Length > 0)
            .GroupBy(style => style.Id, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(style => style.Id, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static bool IsAllowedWordNativeTocMaterializedStyle(WordStyleIdentity style)
    {
        var id = style.Id.Replace(" ", string.Empty, StringComparison.Ordinal).Trim();
        if (id.Equals("Hyperlink", StringComparison.OrdinalIgnoreCase)
            || id.Equals("FollowedHyperlink", StringComparison.OrdinalIgnoreCase)
            || id.Equals("TOCHeading", StringComparison.OrdinalIgnoreCase))
            return true;

        if (id.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
        {
            var suffix = id[3..];
            return suffix.Length == 1 && suffix[0] >= '1' && suffix[0] <= '9';
        }

        return false;
    }

    private static bool VerifyTocStyleMaterializationDelta(
        string beforePath,
        string afterPath,
        TocLocalOpenXmlDeltaProof tocProof,
        int beforeStyleCount,
        int afterStyleCount,
        out string evidence)
    {
        try
        {
            if (!tocProof.Applicable || !tocProof.Accepted)
            {
                evidence = "TOC-local structural proof is not accepted; style-count variance remains strict.";
                return false;
            }

            var before = ReadWordStyleIdentities(beforePath);
            var after = ReadWordStyleIdentities(afterPath);
            var beforeIds = new HashSet<string>(before.Select(style => style.Id), StringComparer.OrdinalIgnoreCase);
            var afterIds = new HashSet<string>(after.Select(style => style.Id), StringComparer.OrdinalIgnoreCase);

            var removed = beforeIds.Where(id => !afterIds.Contains(id)).OrderBy(id => id, StringComparer.OrdinalIgnoreCase).ToArray();
            var added = after.Where(style => !beforeIds.Contains(style.Id)).ToArray();
            var disallowed = added.Where(style => !IsAllowedWordNativeTocMaterializedStyle(style)).ToArray();

            var countDeltaMatches = afterStyleCount - beforeStyleCount == added.Length - removed.Length;
            var pass =
                removed.Length == 0
                && added.Length > 0
                && disallowed.Length == 0
                && countDeltaMatches;

            evidence =
                $"TOC_STYLE_MATERIALIZATION={(pass ? "PASS" : "FAIL")};" +
                $"COUNT={beforeStyleCount}->{afterStyleCount};" +
                $"ADDED={string.Join(",", added.Select(style => style.Id + "[" + style.Name + "]"))};" +
                $"REMOVED={string.Join(",", removed)};" +
                $"DISALLOWED={string.Join(",", disallowed.Select(style => style.Id + "[" + style.Name + "]"))};" +
                $"COUNT_DELTA_MATCH={countDeltaMatches}";
            return pass;
        }
        catch (Exception ex)
        {
            evidence = "TOC style materialization forensic error: " + ex.Message;
            return false;
        }
    }

    private static TocLocalOpenXmlDeltaProof BuildTocLocalOpenXmlDeltaProof(
        string beforePath,
        string afterPath,
        OpenXmlAuditSnapshot before,
        OpenXmlAuditSnapshot after,
        string profileId)
    {
        var applicable =
            string.Equals(profileId, "proposal.governance.fa", StringComparison.OrdinalIgnoreCase)
            || string.Equals(profileId, "book.university.fa", StringComparison.OrdinalIgnoreCase);

        if (!applicable)
        {
            return new TocLocalOpenXmlDeltaProof(
                false,
                false,
                after.Paragraphs - before.Paragraphs,
                0,
                after.FieldInstructions - before.FieldInstructions,
                0,
                0,
                0,
                "Profile is not eligible for Word-native TOC structural-delta policy.");
        }

        TocScopeSnapshot beforeToc;
        TocScopeSnapshot afterToc;
        try
        {
            beforeToc = ReadSingleTocScope(beforePath);
            afterToc = ReadSingleTocScope(afterPath);
        }
        catch (Exception ex)
        {
            return new TocLocalOpenXmlDeltaProof(
                true,
                false,
                after.Paragraphs - before.Paragraphs,
                0,
                after.FieldInstructions - before.FieldInstructions,
                0,
                0,
                0,
                "TOC-local OpenXML proof could not be read: " + ex.Message);
        }

        var totalParagraphDelta = after.Paragraphs - before.Paragraphs;
        var tocParagraphDelta = afterToc.Paragraphs - beforeToc.Paragraphs;
        var totalFieldDelta = after.FieldInstructions - before.FieldInstructions;
        var tocFieldDelta = afterToc.FieldInstructions - beforeToc.FieldInstructions;

        var outsideVisibleTextEqual = string.Equals(
            beforeToc.OutsideVisibleTextSha256,
            afterToc.OutsideVisibleTextSha256,
            StringComparison.Ordinal);

        var outsideStructureEqual =
            beforeToc.OutsideParagraphs == afterToc.OutsideParagraphs
            && beforeToc.OutsideTables == afterToc.OutsideTables
            && beforeToc.OutsideFieldInstructions == afterToc.OutsideFieldInstructions
            && beforeToc.OutsideHyperlinks == afterToc.OutsideHyperlinks
            && beforeToc.OutsideSections == afterToc.OutsideSections
            && beforeToc.OutsideDrawings == afterToc.OutsideDrawings;

        // Word may legitimately regenerate stale TOC result paragraphs and PAGEREF fields.
        // Do not require count stability or non-negative deltas. Instead, fail closed unless
        // every Paragraph/FieldInstruction delta is spatially confined to the one TOC and
        // the complete visible text + protected structure outside that TOC is identical.
        // QH01-E226: locality proof supports both refresh (1->1) and true creation
        // (0->1) while keeping every structural/text delta confined to the TOC scope.
        var authorityTransitionAccepted =
            (beforeToc.Controls == 1 && afterToc.Controls == 1 && beforeToc.VisibleParagraphs > 0)
            || (beforeToc.Controls == 0 && afterToc.Controls == 1 && beforeToc.VisibleParagraphs == 0);

        var accepted =
            authorityTransitionAccepted
            && afterToc.VisibleParagraphs > 0
            && totalParagraphDelta == tocParagraphDelta
            && totalFieldDelta == tocFieldDelta
            && outsideVisibleTextEqual
            && outsideStructureEqual;

        var reason = accepted
            ? "All Paragraph/FieldInstruction deltas are confined to the single TOC; outside-TOC visible text and protected structure are identical."
            : $"Fail-closed TOC locality mismatch: controls={beforeToc.Controls}->{afterToc.Controls}, visible={beforeToc.VisibleParagraphs}->{afterToc.VisibleParagraphs}, totalP={totalParagraphDelta}, tocP={tocParagraphDelta}, totalFI={totalFieldDelta}, tocFI={tocFieldDelta}, outsideText={outsideVisibleTextEqual}, outsideStructure={outsideStructureEqual}.";

        return new TocLocalOpenXmlDeltaProof(
            true,
            accepted,
            totalParagraphDelta,
            tocParagraphDelta,
            totalFieldDelta,
            tocFieldDelta,
            beforeToc.VisibleParagraphs,
            afterToc.VisibleParagraphs,
            reason);
    }

    private static bool IsAutomaticTocMainInstruction(string? value)
    {
        var normalized = string.Join(
            ' ',
            (value ?? string.Empty)
                .Replace('\u200c', ' ')
                .Replace('\u200f', ' ')
                .Replace('\u200e', ' ')
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

        return normalized.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)
               && (normalized.Length == 3 || char.IsWhiteSpace(normalized[3]));
    }

    private static bool HasAutomaticTocEntrySignature(XElement paragraph)
    {
        var instruction = string.Join(
            " ",
            paragraph
                .Descendants(WordMain + "instrText")
                .Select(x => ((string?)x ?? string.Empty).Trim()));

        if (instruction.Contains("PAGEREF", StringComparison.OrdinalIgnoreCase)
            && instruction.Contains("_Toc", StringComparison.OrdinalIgnoreCase))
            return true;

        return paragraph
            .Descendants(WordMain + "hyperlink")
            .Select(link => (string?)link.Attribute(WordMain + "anchor"))
            .Any(anchor =>
                !string.IsNullOrWhiteSpace(anchor)
                && anchor.StartsWith("_Toc", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsRecognizedAutomaticTocSdt(XElement sdt)
    {
        var gallery = (string?)sdt
            .Element(WordMain + "sdtPr")?
            .Element(WordMain + "docPartObj")?
            .Element(WordMain + "docPartGallery")?
            .Attribute(WordMain + "val");

        if (string.Equals(gallery?.Trim(), "Table of Contents", StringComparison.OrdinalIgnoreCase))
            return true;

        var joinedInstruction = string.Join(
            " ",
            sdt.Descendants(WordMain + "instrText")
                .Select(x => ((string?)x ?? string.Empty).Trim()));

        if (IsAutomaticTocMainInstruction(joinedInstruction))
            return true;

        return sdt
            .Descendants(WordMain + "p")
            .Any(HasAutomaticTocEntrySignature);
    }

    private static string NormalizeVisibleParagraphText(XElement paragraph)
        => string.Join(
            " ",
            paragraph.Descendants(WordMain + "t")
                .Select(x => ((string?)x ?? string.Empty).Trim())
                .Where(x => x.Length > 0))
            .Trim();

    private static bool IsExactPersianTocTitleForRecovery(string text)
    {
        var compact = new string(
            (text ?? string.Empty)
                .Where(ch => !char.IsWhiteSpace(ch) && ch != '\u200c' && ch != '\u200f' && ch != '\u200e')
                .ToArray());

        return compact.Equals("فهرستمطالب", StringComparison.Ordinal);
    }

    private static bool LooksLikeDegradedPlainTextTocEntry(XElement paragraph)
    {
        var text = NormalizeVisibleParagraphText(paragraph);
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var hasTab = paragraph.Descendants(WordMain + "tab").Any();
        if (!hasTab)
            return false;

        var tail = text.Split(' ', StringSplitOptions.RemoveEmptyEntries).LastOrDefault() ?? string.Empty;
        return tail.Length > 0
               && tail.All(ch =>
                   char.IsDigit(ch)
                   || (ch >= '\u06F0' && ch <= '\u06F9')
                   || (ch >= '\u0660' && ch <= '\u0669'));
    }

    private static int[] FindDegradedPlainTextTocParagraphIndexes(XDocument document)
    {
        var paragraphs = document.Descendants(WordMain + "p").ToArray();
        var titleIndex = Array.FindIndex(
            paragraphs,
            paragraph => IsExactPersianTocTitleForRecovery(NormalizeVisibleParagraphText(paragraph)));

        if (titleIndex < 0)
            return Array.Empty<int>();

        var result = new List<int>();
        for (var index = titleIndex + 1; index < paragraphs.Length && result.Count < 120; index++)
        {
            var paragraph = paragraphs[index];
            var text = NormalizeVisibleParagraphText(paragraph);

            if (string.IsNullOrWhiteSpace(text))
            {
                if (result.Count == 0)
                    continue;
                break;
            }

            if (!LooksLikeDegradedPlainTextTocEntry(paragraph))
                break;

            result.Add(index);
        }

        return result.Count >= 2 ? result.ToArray() : Array.Empty<int>();
    }

    private static int[] FindRecoveredDocumentLevelTocParagraphIndexes(XDocument document)
    {
        var paragraphs = document.Descendants(WordMain + "p").ToArray();
        var entryIndexes = paragraphs
            .Select((paragraph, index) => new { paragraph, index })
            .Where(x => HasAutomaticTocEntrySignature(x.paragraph))
            .Select(x => x.index)
            .ToArray();

        if (entryIndexes.Length < 2)
            return Array.Empty<int>();

        var start = entryIndexes.Min();
        var end = entryIndexes.Max();

        var mainIndexes = paragraphs
            .Select((paragraph, index) => new
            {
                index,
                instruction = string.Join(
                    " ",
                    paragraph.Descendants(WordMain + "instrText")
                        .Select(x => ((string?)x ?? string.Empty).Trim()))
            })
            .Where(x => IsAutomaticTocMainInstruction(x.instruction))
            .Select(x => x.index)
            .Where(index => index <= start)
            .ToArray();

        if (mainIndexes.Length > 0)
        {
            var nearestMain = mainIndexes.Max();
            if (start - nearestMain <= 3)
                start = nearestMain;
        }

        if (end - start > 250)
            return Array.Empty<int>();

        return Enumerable.Range(start, end - start + 1).ToArray();
    }

    private static TocScopeSnapshot ReadSingleTocScope(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidOperationException("word/document.xml is missing.");

        using var stream = entry.Open();
        var document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);

        var controls = document
            .Descendants(WordMain + "sdt")
            .Where(IsRecognizedAutomaticTocSdt)
            .ToArray();

        XElement[] paragraphs;
        int[] recoveredIndexes = Array.Empty<int>();
        var logicalControls = 0;

        if (controls.Length == 1)
        {
            paragraphs = controls[0].Descendants(WordMain + "p").ToArray();
            logicalControls = 1;
        }
        else if (controls.Length == 0)
        {
            recoveredIndexes = FindRecoveredDocumentLevelTocParagraphIndexes(document);
            if (recoveredIndexes.Length == 0)
                recoveredIndexes = FindDegradedPlainTextTocParagraphIndexes(document);

            if (recoveredIndexes.Length == 0)
            {
                // QH01-E276: a true pre-creation 0-TOC document is still a complete
                // outside-TOC authority snapshot. Returning all-zero outside metrics here
                // made every legitimate 0->1 TOC creation look like unrelated document drift.
                paragraphs = Array.Empty<XElement>();
                logicalControls = 0;
            }
            else
            {
                var allParagraphs = document.Descendants(WordMain + "p").ToArray();
                paragraphs = recoveredIndexes
                    .Where(index => index >= 0 && index < allParagraphs.Length)
                    .Select(index => allParagraphs[index])
                    .ToArray();
                logicalControls = 1;
            }
        }
        else
        {
            // Multiple independent TOC controls remain fail-closed and deliberately do not
            // receive a fabricated outside-authority fingerprint.
            return new TocScopeSnapshot(controls.Length, 0, 0, 0, string.Empty, 0, 0, 0, 0, 0, 0);
        }

        var visibleParagraphs = paragraphs.Count(paragraph =>
            paragraph.Descendants(WordMain + "t")
                .Any(text => !string.IsNullOrWhiteSpace((string?)text)));

        var outside = new XDocument(document);
        if (controls.Length == 1)
        {
            var outsideControls = outside
                .Descendants(WordMain + "sdt")
                .Where(IsRecognizedAutomaticTocSdt)
                .ToArray();
            foreach (var control in outsideControls)
                control.Remove();
        }
        else
        {
            var outsideParagraphs = outside.Descendants(WordMain + "p").ToArray();
            foreach (var index in recoveredIndexes.OrderByDescending(x => x))
            {
                if (index >= 0 && index < outsideParagraphs.Length)
                    outsideParagraphs[index].Remove();
            }
        }

        var outsideText = string.Concat(
            outside.Descendants(WordMain + "t")
                .Select(text => (string?)text ?? string.Empty));
        var outsideTextSha256 = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(outsideText)));

        return new TocScopeSnapshot(
            logicalControls,
            paragraphs.Length,
            visibleParagraphs,
            paragraphs.Sum(p => p.Descendants(WordMain + "instrText").Count()),
            outsideTextSha256,
            outside.Descendants(WordMain + "p").Count(),
            outside.Descendants(WordMain + "tbl").Count(),
            outside.Descendants(WordMain + "instrText").Count(),
            outside.Descendants(WordMain + "hyperlink").Count(),
            outside.Descendants(WordMain + "sectPr").Count(),
            outside.Descendants(WordMain + "drawing").Count());
    }

    private static string FormatControlledFixReport(
        ControlledCopyFixResult result,
        OpenXmlAuditSnapshot beforeOpenXml,
        OpenXmlAuditSnapshot afterOpenXml,
        IReadOnlyList<string> openXmlIssues,
        bool approved)
    {
        var sb = new StringBuilder();
        sb.AppendLine("گزارش اصلاح کنترل‌شده WordAgent");
        sb.AppendLine(new string('=', 60));
        sb.AppendLine("Schema: wordagent.controlledfix/1.1");
        sb.AppendLine("Input: " + result.InputFile);
        sb.AppendLine("Output: " + result.OutputFile);
        sb.AppendLine("Profile: " + result.ProfileId);
        sb.AppendLine("Input SHA-256: " + result.InputSha256);
        sb.AppendLine("Output SHA-256: " + result.OutputSha256);
        sb.AppendLine("Original overwrite: NO");
        sb.AppendLine();

        sb.AppendLine("[گیت قبل از Mutation]");
        sb.AppendLine($"  Semantic classification: Unclassified={result.BeforeCompliance.UnclassifiedParagraphs}");
        sb.AppendLine($"  Style-definition failures: {result.BeforeCompliance.Styles.Count(x => !x.FullyCompliant)}");
        sb.AppendLine($"  Revisions: {result.BeforeWord.Counts.Revisions}");
        sb.AppendLine($"  Candidate paragraphs: {result.CandidateParagraphs}");
        sb.AppendLine();

        sb.AppendLine("[Mutationهای اعمال‌شده]");
        foreach (var action in result.Actions.OrderBy(x => x.ParagraphIndex))
        {
            sb.AppendLine($"  P{action.ParagraphIndex} | Role={action.Role} | {action.Action}");
            sb.AppendLine($"    Style: {action.SourceStyle} -> {action.TargetStyle ?? action.SourceStyle}");
            sb.AppendLine("    " + action.TextPreview);
        }
        sb.AppendLine();

        var remaining = result.AfterCompliance.Roles.Sum(x => x.Total - x.FullyCompliant);
        sb.AppendLine("[گیت بعد از Mutation]");
        sb.AppendLine($"  Mutated paragraphs: {result.MutatedParagraphs}/{result.CandidateParagraphs}");
        sb.AppendLine($"  Remaining non-compliant paragraphs: {remaining}");
        sb.AppendLine($"  Unclassified paragraphs: {result.AfterCompliance.UnclassifiedParagraphs}");
        sb.AppendLine($"  Direct-formatting overrides suspected: {result.AfterCompliance.DirectOverridesSuspected}");
        var profileGatePassed = remaining == 0
            && result.AfterCompliance.UnclassifiedParagraphs == 0
            && result.AfterCompliance.Styles.All(x => x.FullyCompliant);
        sb.AppendLine($"  Word inventory integrity: {(result.WordIntegrityPassed ? "PASS" : "FAIL")}");
        sb.AppendLine($"  Profile compliance gate: {(profileGatePassed ? "PASS" : "FAIL")}");
        sb.AppendLine($"  OpenXML validation errors: {beforeOpenXml.ValidationErrors} -> {afterOpenXml.ValidationErrors}");
        sb.AppendLine($"  OpenXML structural integrity: {(openXmlIssues.Count == 0 ? "PASS" : "FAIL")}");
        sb.AppendLine();

        if (result.IntegrityIssues.Count > 0 || openXmlIssues.Count > 0)
        {
            sb.AppendLine("[موارد مسدودکننده]");
            foreach (var issue in result.IntegrityIssues.Concat(openXmlIssues)) sb.AppendLine("  - " + issue);
            sb.AppendLine();
        }

        if (result.Warnings.Count > 0)
        {
            sb.AppendLine("[هشدارهای موتور اصلاح / Word COM کنترل‌شده]");
            foreach (var warning in result.Warnings) sb.AppendLine("  - " + warning);
            sb.AppendLine();
        }

        sb.AppendLine("[وضعیت]");
        sb.AppendLine("  Integrity gate: " + (approved ? "PASS" : "FAIL"));
        sb.AppendLine("  Output status: " + (approved ? "APPROVED_FOR_VISUAL_REVIEW" : "NOT_APPROVED"));
        sb.AppendLine("  برای پروفایل پیشنهاده حاکمیتی، نتیجه یک TOC با Microsoft Word به‌روزرسانی می‌شود؛ EndNote و Fieldهای خارج از TOC بازتولید یا بازشماری نمی‌شوند.");
        sb.AppendLine("  The original file remains untouched.");
        return FinalizeFaReport(sb);
    }


    private static void ApplyTableThemeOverride(AgentProfile profile, Dictionary<string, string?> options)
    {
        if (!options.TryGetValue("table-theme", out var raw) || string.IsNullOrWhiteSpace(raw))
            return;

        var token = raw.Trim().ToUpperInvariant().Replace('-', '_');
        if (token == "NONE")
        {
            profile.LayoutPolicy.TableStyle.Enabled = false;
            return;
        }

        if (token is not ("LIGHT_BLUE" or "BLUE" or "GREEN"))
            throw new InvalidOperationException($"Unsupported --table-theme '{raw}'. Allowed: LIGHT_BLUE, BLUE, GREEN, NONE.");

        profile.LayoutPolicy.TableStyle.Enabled = true;
        profile.LayoutPolicy.TableStyle.Theme = token;
    }


    private static string ResolveProfilesDir(Dictionary<string, string?> options)
    {
        if (options.TryGetValue("profiles", out var explicitDir) && !string.IsNullOrWhiteSpace(explicitDir))
            return Path.GetFullPath(explicitDir);

        var besideExe = Path.Combine(AppContext.BaseDirectory, "profiles");
        if (Directory.Exists(besideExe)) return besideExe;

        var current = Path.Combine(Environment.CurrentDirectory, "profiles");
        return current;
    }

    private static string FormatPersianReport(CombinedAuditReport r)
    {
        var sb = new StringBuilder();
        sb.AppendLine("گزارش ممیزی WordAgent");
        sb.AppendLine(new string('=', 60));
        sb.AppendLine("فایل: " + r.InputFile);
        sb.AppendLine("زمان: " + r.GeneratedAt.ToString("u"));
        sb.AppendLine("Schema: " + r.SchemaVersion);
        sb.AppendLine();

        if (r.Word is not null)
        {
            var c = r.Word.Counts;
            var f = r.Word.Formatting;
            sb.AppendLine("[ممیزی توسط Microsoft Word]");
            sb.AppendLine($"Word version: {r.Word.WordVersion}");
            sb.AppendLine($"Paragraphs: {c.Paragraphs}");
            sb.AppendLine($"Tables: {c.Tables} | RTL: {f.RtlTables} | Centered: {f.CenteredTables}");
            sb.AppendLine($"Fields: {c.Fields} | EndNote: {c.EndNoteFields} | TOC: {c.TocFields}");
            sb.AppendLine($"Styles: {c.Styles} | Sections: {c.Sections}");
            sb.AppendLine($"Footnotes: {c.Footnotes} | Endnotes: {c.Endnotes} | Equations: {c.Equations}");
            sb.AppendLine($"Hyperlinks: {c.Hyperlinks} | Comments: {c.Comments} | Revisions: {c.Revisions}");
            sb.AppendLine($"Non-empty paragraphs: {f.NonEmptyParagraphs}");
            sb.AppendLine($"RTL paragraphs: {f.RtlParagraphs}");
            sb.AppendLine($"Right-aligned paragraphs: {f.RightAlignedParagraphs}");
            sb.AppendLine($"Centered paragraphs: {f.CenteredParagraphs}");
            sb.AppendLine($"Inline images centered: {f.InlineImagesCentered}/{c.InlineShapes}");
            sb.AppendLine();
        }

        if (r.Compliance is not null)
        {
            var p = r.Compliance;
            sb.AppendLine("[انطباق با پروفایل]");
            sb.AppendLine($"Profile: {p.ProfileId} | {p.ProfileDisplayName}");
            sb.AppendLine($"Non-empty: {p.NonEmptyParagraphs} | Classified: {p.ClassifiedParagraphs} | Unclassified: {p.UnclassifiedParagraphs} | In tables: {p.ParagraphsInTables}");
            sb.AppendLine($"Direct-formatting overrides suspected: {p.DirectOverridesSuspected}");
            sb.AppendLine();

            foreach (var role in p.Roles)
            {
                var pct = role.Total == 0 ? 100d : role.FullyCompliant * 100d / role.Total;
                sb.AppendLine($"Role {role.Role}: {role.FullyCompliant}/{role.Total} fully compliant ({pct:0.0}%)");
                sb.AppendLine($"  FontBi: {role.FontBiMatches}/{role.Total} | SizeBi: {role.SizeBiMatches}/{role.Total} | RTL/LTR: {role.ReadingOrderMatches}/{role.Total} | Alignment: {role.AlignmentMatches}/{role.Total} | Overrides?: {role.DirectOverridesSuspected}");
                sb.AppendLine($"  Expected: FontBi={role.ExpectedFontBi ?? "(any)"}; SizeBi={FormatSize(role.ExpectedSizeBi)}; ReadingOrder={role.ExpectedReadingOrder ?? "(any)"}; Alignment={role.ExpectedAlignment ?? "(any)"}");
                foreach (var sample in role.Samples)
                {
                    sb.AppendLine($"    P{sample.ParagraphIndex} | Style={sample.StyleName} | FontBi={sample.ActualFontBi ?? "(mixed/unknown)"} | SizeBi={FormatSize(sample.ActualSizeBi)} | {sample.ActualReadingOrder}/{sample.ActualAlignment} | Override?={(sample.DirectOverrideSuspected ? "YES" : "NO")}");
                    sb.AppendLine("      " + sample.TextPreview);
                }
            }

            sb.AppendLine();
            sb.AppendLine("[انطباق تعریف Styleها — OOXML با اعمال وراثت]");
            foreach (var style in p.Styles.Where(x => x.Found || !x.FullyCompliant))
            {
                var state = style.FullyCompliant ? "PASS" : style.Found ? "FAIL" : "MISSING";
                sb.AppendLine($"  {state} | Role={style.Role} | Style={style.StyleName} | FontBi={style.ActualFontBi ?? "(unknown)"} | SizeBi={FormatSize(style.ActualSizeBi)} | {style.ActualReadingOrder}/{style.ActualAlignment}");
            }

            if (p.UnclassifiedSamples.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("[نمونه پاراگراف‌های طبقه‌بندی‌نشده]");
                foreach (var sample in p.UnclassifiedSamples)
                    sb.AppendLine($"  P{sample.ParagraphIndex} | Style={sample.StyleName} | {sample.TextPreview}");
            }

            if (p.Warnings.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("[هشدارهای Compliance]");
                foreach (var warning in p.Warnings.Take(30)) sb.AppendLine(" - " + warning);
            }
            sb.AppendLine();
        }

        if (r.OpenXml is not null)
        {
            var x = r.OpenXml;
            sb.AppendLine("[ممیزی ساختاری Open XML]");
            sb.AppendLine($"Paragraphs: {x.Paragraphs} | Tables: {x.Tables} | Drawings: {x.Drawings} | Styles: {x.Styles}");
            sb.AppendLine($"Field instructions: {x.FieldInstructions} | EndNote markers: {x.EndNoteMarkers}");
            sb.AppendLine("Schema validation target: Office2016");
            sb.AppendLine($"Schema validation errors (first 101 max): {x.ValidationErrors}");
            foreach (var sample in x.ValidationSamples) sb.AppendLine(" - " + sample);
            sb.AppendLine();
        }

        if (r.Warnings.Count > 0)
        {
            sb.AppendLine("[هشدارها]");
            foreach (var w in r.Warnings) sb.AppendLine(" - " + w);
        }
        return FinalizeFaReport(sb);
    }

    private static string FinalizeFaReport(StringBuilder sb)
    {
        var normalized = sb.ToString()
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');

        var lines = normalized.Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (lines[i].Length == 0) continue;
            var line = LocalizeFaReportLine(lines[i]);
            lines[i] = line.StartsWith('\u200F') ? line : "\u200F" + line;
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static string LocalizeFaReportLine(string line)
    {
        if (string.IsNullOrEmpty(line)) return line;

        var localized = line
            .Replace("Schema:", "شناسه ساختار:", StringComparison.OrdinalIgnoreCase)
            .Replace("Input:", "ورودی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Output:", "خروجی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Template:", "تمپلیت:", StringComparison.OrdinalIgnoreCase)
            .Replace("Profile:", "پروفایل:", StringComparison.OrdinalIgnoreCase)
            .Replace("Input SHA-256:", "SHA-256 ورودی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Output SHA-256:", "SHA-256 خروجی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Template SHA-256:", "SHA-256 تمپلیت:", StringComparison.OrdinalIgnoreCase)
            .Replace("Original overwrite:", "بازنویسی فایل اصلی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Original input overwrite:", "بازنویسی ورودی اصلی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Template mutation:", "تغییر تمپلیت:", StringComparison.OrdinalIgnoreCase)
            .Replace("Mutation:", "اعمال تغییر:", StringComparison.OrdinalIgnoreCase)
            .Replace("AI semantic mapping:", "نگاشت معنایی هوش مصنوعی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Profile detection:", "تشخیص پروفایل:", StringComparison.OrdinalIgnoreCase)
            .Replace("Document-family formatting:", "قالب‌بندی خانواده سند:", StringComparison.OrdinalIgnoreCase)
            .Replace("Word version:", "نسخه Word:", StringComparison.OrdinalIgnoreCase)
            .Replace("Paragraphs:", "پاراگراف‌ها:", StringComparison.OrdinalIgnoreCase)
            .Replace("Tables:", "جدول‌ها:", StringComparison.OrdinalIgnoreCase)
            .Replace("Fields:", "Fieldها:", StringComparison.OrdinalIgnoreCase)
            .Replace("Styles:", "Styleها:", StringComparison.OrdinalIgnoreCase)
            .Replace("Sections:", "بخش‌ها:", StringComparison.OrdinalIgnoreCase)
            .Replace("Footnotes:", "پاورقی‌ها:", StringComparison.OrdinalIgnoreCase)
            .Replace("Endnotes:", "یادداشت‌های پایانی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Equations:", "فرمول‌ها:", StringComparison.OrdinalIgnoreCase)
            .Replace("Hyperlinks:", "پیوندها:", StringComparison.OrdinalIgnoreCase)
            .Replace("Comments:", "نظرها:", StringComparison.OrdinalIgnoreCase)
            .Replace("Revisions:", "تغییرات رهگیری‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Non-empty paragraphs:", "پاراگراف‌های غیرخالی:", StringComparison.OrdinalIgnoreCase)
            .Replace("RTL paragraphs:", "پاراگراف‌های راست‌به‌چپ:", StringComparison.OrdinalIgnoreCase)
            .Replace("Right-aligned paragraphs:", "پاراگراف‌های راست‌چین:", StringComparison.OrdinalIgnoreCase)
            .Replace("Centered paragraphs:", "پاراگراف‌های وسط‌چین:", StringComparison.OrdinalIgnoreCase)
            .Replace("Inline images centered:", "تصاویر درون‌خطی وسط‌چین:", StringComparison.OrdinalIgnoreCase)
            .Replace("Non-empty:", "غیرخالی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Classified:", "طبقه‌بندی‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Unclassified:", "طبقه‌بندی‌نشده:", StringComparison.OrdinalIgnoreCase)
            .Replace("In tables:", "داخل جدول:", StringComparison.OrdinalIgnoreCase)
            .Replace("Direct-formatting overrides suspected:", "موارد مشکوک قالب‌بندی مستقیم:", StringComparison.OrdinalIgnoreCase)
            .Replace("Role ", "نقش ", StringComparison.OrdinalIgnoreCase)
            .Replace(" fully compliant", " کاملاً منطبق", StringComparison.OrdinalIgnoreCase)
            .Replace("Expected:", "مقدار مورد انتظار:", StringComparison.OrdinalIgnoreCase)
            .Replace("Style=", "Style=", StringComparison.OrdinalIgnoreCase)
            .Replace("Override?=", "Override?=", StringComparison.OrdinalIgnoreCase)
            .Replace("OpenXML validation errors:", "خطاهای اعتبارسنجی OpenXML:", StringComparison.OrdinalIgnoreCase)
            .Replace("OpenXML structural integrity:", "یکپارچگی ساختاری OpenXML:", StringComparison.OrdinalIgnoreCase)
            .Replace("Word inventory integrity:", "یکپارچگی موجودی Word:", StringComparison.OrdinalIgnoreCase)
            .Replace("Profile compliance gate:", "گیت انطباق پروفایل:", StringComparison.OrdinalIgnoreCase)
            .Replace("Integrity gate:", "گیت یکپارچگی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Output status:", "وضعیت خروجی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Candidate paragraphs:", "پاراگراف‌های نامزد:", StringComparison.OrdinalIgnoreCase)
            .Replace("Persian paragraphs found:", "پاراگراف‌های فارسی یافت‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Paragraph direction/alignment changes:", "تغییر جهت/تراز پاراگراف:", StringComparison.OrdinalIgnoreCase)
            .Replace("Persian run RTL changes:", "تغییر RTL در Runهای فارسی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Tables examined:", "جدول‌های بررسی‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Tables centered:", "جدول‌های وسط‌چین‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Persian tables RTL-enabled:", "جدول‌های فارسی با RTL فعال:", StringComparison.OrdinalIgnoreCase)
            .Replace("TOC styles changed:", "Styleهای فهرست اصلاح‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Automatic field update on open disabled:", "به‌روزرسانی خودکار Field هنگام بازشدن غیرفعال:", StringComparison.OrdinalIgnoreCase)
            .Replace("No changes needed:", "نیازی به تغییر نبود:", StringComparison.OrdinalIgnoreCase)
            .Replace("Template styles imported:", "Styleهای تمپلیت واردشده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Template styles replaced:", "Styleهای تمپلیت جایگزین‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Paragraphs remapped by deterministic style role:", "پاراگراف‌های بازنگاشت‌شده با نقش قطعی Style:", StringComparison.OrdinalIgnoreCase)
            .Replace("Sections normalized to template geometry:", "بخش‌های منطبق‌شده با هندسه تمپلیت:", StringComparison.OrdinalIgnoreCase)
            .Replace("Theme replaced when target part existed:", "Theme در صورت وجود بخش هدف جایگزین شد:", StringComparison.OrdinalIgnoreCase)
            .Replace("Font table replaced when target part existed:", "جدول فونت در صورت وجود بخش هدف جایگزین شد:", StringComparison.OrdinalIgnoreCase)
            .Replace("Visual review in Microsoft Word is still mandatory because pagination is a rendered-layout property.",
                "بازبینی بصری در Microsoft Word همچنان الزامی است، چون صفحه‌بندی وابسته به رندر Word است.",
                StringComparison.OrdinalIgnoreCase)
            .Replace("Visual review in Microsoft Word remains required before production use.",
                "پیش از استفاده نهایی، بازبینی بصری در Microsoft Word الزامی است.",
                StringComparison.OrdinalIgnoreCase)
            .Replace("Visual review in Microsoft Word is still required for rendered pagination and appearance.",
                "برای صفحه‌بندی و ظاهر نهایی، بازبینی بصری در Microsoft Word همچنان لازم است.",
                StringComparison.OrdinalIgnoreCase)
            .Replace("The original file remains untouched.", "فایل اصلی بدون تغییر باقی مانده است.", StringComparison.OrdinalIgnoreCase)
            .Replace("برای پروفایل پیشنهاده حاکمیتی، نتیجه یک TOC با Microsoft Word به‌روزرسانی می‌شود؛ EndNote و Fieldهای خارج از TOC بازتولید یا بازشماری نمی‌شوند.",
                "این فرمان محتوای TOC/Field/EndNote را بازتولید یا بازشماری نمی‌کند.",
                StringComparison.OrdinalIgnoreCase)
            .Replace("READY_FOR_CONTROLLED_COPY_FIX", "آماده برای اصلاح کنترل‌شده روی کپی", StringComparison.OrdinalIgnoreCase)
            .Replace("APPROVED_FOR_VISUAL_REVIEW", "تأییدشده برای بازبینی بصری", StringComparison.OrdinalIgnoreCase)
            .Replace("NOT_APPROVED", "تأییدنشده", StringComparison.OrdinalIgnoreCase)
            .Replace("NOT USED", "استفاده نشد", StringComparison.OrdinalIgnoreCase)
            .Replace("MISSING", "گمشده", StringComparison.OrdinalIgnoreCase)
            .Replace("BLOCKED", "مسدود", StringComparison.OrdinalIgnoreCase)
            .Replace("BLOCKER:", "مسدودکننده:", StringComparison.OrdinalIgnoreCase)
            .Replace("BLOCK |", "مسدود |", StringComparison.OrdinalIgnoreCase)
            .Replace("GUARD |", "محافظ |", StringComparison.OrdinalIgnoreCase)
            .Replace("WARN  |", "هشدار |", StringComparison.OrdinalIgnoreCase)
            .Replace("PASS |", "موفق |", StringComparison.OrdinalIgnoreCase)
            .Replace("FAIL |", "ناموفق |", StringComparison.OrdinalIgnoreCase)
            .Replace(": PASS", ": موفق", StringComparison.OrdinalIgnoreCase)
            .Replace(": FAIL", ": ناموفق", StringComparison.OrdinalIgnoreCase)
            .Replace("=PASS", "=موفق", StringComparison.OrdinalIgnoreCase)
            .Replace("=FAIL", "=ناموفق", StringComparison.OrdinalIgnoreCase);

        localized = ReplaceStandaloneAsciiToken(localized, "YES", "بله");
        localized = ReplaceStandaloneAsciiToken(localized, "NO", "خیر");
        return localized;
    }

    private static string ReplaceStandaloneAsciiToken(string input, string token, string replacement)
    {
        if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(token)) return input;

        var result = new StringBuilder(input.Length);
        var index = 0;
        while (index < input.Length)
        {
            var remaining = input.Length - index;
            var matches = remaining >= token.Length
                && string.Compare(input, index, token, 0, token.Length, StringComparison.OrdinalIgnoreCase) == 0;

            if (matches)
            {
                var leftBoundary = index == 0 || !IsReportTokenChar(input[index - 1]);
                var rightIndex = index + token.Length;
                var rightBoundary = rightIndex >= input.Length || !IsReportTokenChar(input[rightIndex]);
                if (leftBoundary && rightBoundary)
                {
                    result.Append(replacement);
                    index += token.Length;
                    continue;
                }
            }

            result.Append(input[index]);
            index++;
        }

        return result.ToString();
    }

    private static bool IsReportTokenChar(char value)
        => char.IsLetterOrDigit(value) || value == '_';

    private static string FormatSize(float? size)
        => size is null ? "(mixed/unknown)" : size.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);

    private static Dictionary<string, string?> ParseOptions(string[] args)
    {
        var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < args.Length; i++)
        {
            var token = args[i];
            if (!token.StartsWith("--", StringComparison.Ordinal)) continue;
            var key = token[2..];
            if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                result[key] = args[++i];
            else
                result[key] = "true";
        }
        return result;
    }

    private static string Require(Dictionary<string, string?> options, string key)
        => options.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new ArgumentException($"--{key} is required.");

    private static bool IsTrue(Dictionary<string, string?> options, string key)
        => options.TryGetValue(key, out var v) && (string.IsNullOrWhiteSpace(v) || v.Equals("true", StringComparison.OrdinalIgnoreCase) || v == "1");

    private static int Unknown(string command)
    {
        Console.Error.WriteLine("Unknown command: " + command);
        return Help() == 0 ? 2 : 2;
    }
}








