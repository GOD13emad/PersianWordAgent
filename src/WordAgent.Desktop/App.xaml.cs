using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace WordAgent.Desktop;

public partial class App : Application
{
    private static readonly string DesktopVersion = ReleaseContract.DesktopVersion;
    private static readonly string DefaultExpectedEngineVersion = ReleaseContract.ExpectedEngineVersion;
    private const int RealTemplateSmokeTimeoutMs = 1800000;
    private static readonly string[] RequiredProfileIds =
    {
        "book.university.fa",
        "formal-admin-technical-proposal-fa",
        "proposal.academic.fa",
        "proposal.governance.fa",
        "proposal.industrial.fa",
        "questionnaire-assessment-fa",
        "thesis-dissertation-academic-fa",
        "thesis.dissertation.phd.fa"
    };

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        if (e.Args.Any(x => x.Equals("--smoke", StringComparison.OrdinalIgnoreCase)))
        {
            var output = GetOption(e.Args, "--smoke-output");
            var code = RunSmoke(e.Args, output);
            Shutdown(code);
            return;
        }

        ShutdownMode = ShutdownMode.OnMainWindowClose;
        var window = new MainWindow();
        MainWindow = window;
        window.Show();
    }

    private static int RunSmoke(string[] args, string? outputPath)
    {
        var lines = new List<string>();
        try
        {
            if (!RuntimeLocator.TryResolve(out var runtime, out var error))
                throw new InvalidOperationException(error);

            var expectedEngineVersion =
                Environment.GetEnvironmentVariable("WORDAGENT_EXPECTED_ENGINE_VERSION")
                ?? DefaultExpectedEngineVersion;

            lines.Add("DESKTOP_VERSION=" + DesktopVersion);
            lines.Add("ENGINE_EXE=" + runtime.EngineExe);
            lines.Add("PROFILES_DIR=" + runtime.ProfilesDirectory);

            var profiles = Directory.Exists(runtime.ProfilesDirectory)
                ? Directory.EnumerateFiles(
                    runtime.ProfilesDirectory,
                    "*.json",
                    SearchOption.TopDirectoryOnly).ToArray()
                : Array.Empty<string>();

            var profileIds = profiles
                .Select(Path.GetFileNameWithoutExtension)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var missingProfiles = RequiredProfileIds
                .Except(profileIds, StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (profileIds.Length != RequiredProfileIds.Length || missingProfiles.Length != 0)
            {
                throw new InvalidOperationException(
                    $"Required profile surface missing. count={profileIds.Length}, missing={string.Join(",", missingProfiles)}");
            }

            var profileValidationResult = RunEngine(
                runtime.EngineExe,
                new[] { "profile-validate", "--profiles", runtime.ProfilesDirectory },
                30000);

            if (profileValidationResult.ExitCode != 0 ||
                !profileValidationResult.StdOut.Contains("PROFILE_VALIDATION_ERRORS=0", StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"Profile validation failed. exit={profileValidationResult.ExitCode}, stdout={profileValidationResult.StdOut.Trim()}, stderr={profileValidationResult.StdErr.Trim()}");
            }

            var versionResult = RunEngine(runtime.EngineExe, new[] { "version" }, 30000);
            var version = FirstNonEmptyLine(versionResult.StdOut);
            if (versionResult.ExitCode != 0 || !version.Equals(expectedEngineVersion, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"Unexpected engine smoke result. exit={versionResult.ExitCode}, version={version}, expected={expectedEngineVersion}, stderr={versionResult.StdErr.Trim()}");
            }

            var helpResult = RunEngine(runtime.EngineExe, new[] { "help" }, 30000);
            if (helpResult.ExitCode != 0 ||
                !helpResult.StdOut.Contains("persianize-generic", StringComparison.Ordinal) ||
                !helpResult.StdOut.Contains("template-analyze", StringComparison.Ordinal) ||
                !helpResult.StdOut.Contains("template-apply", StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Engine help surface is missing required two-mode product commands.");
            }

            lines.Add("ENGINE_VERSION=" + version);
            lines.Add("PROFILE_COUNT=" + profiles.Length);
            lines.Add("PROFILE_VALIDATION=PASS");
            lines.Add("PROFILE_SMART_MODES_READY=True");
            lines.Add("PROFILES_REQUIRED_FOR_GENERIC_TEMPLATE_MODES=False");
            lines.Add("GENERIC_COMMAND_PRESENT=True");
            lines.Add("TEMPLATE_ANALYZE_COMMAND_PRESENT=True");
            lines.Add("TEMPLATE_APPLY_COMMAND_PRESENT=True");

            if (args.Any(x => x.Equals("--smoke-generic", StringComparison.OrdinalIgnoreCase)))
            {
                var input = GetOption(args, "--smoke-input")
                    ?? throw new InvalidOperationException("--smoke-generic requires --smoke-input.");
                var outDir = GetOption(args, "--smoke-outdir")
                    ?? throw new InvalidOperationException("--smoke-generic requires --smoke-outdir.");

                Directory.CreateDirectory(outDir);
                var genericResult = RunEngine(
                    runtime.EngineExe,
                    new[]
                    {
                        "persianize-generic",
                        "--input", input,
                        "--outdir", outDir,
                        "--confirm-generic"
                    },
                    180000);

                if (genericResult.ExitCode != 0 ||
                    !genericResult.StdOut.Contains("NO_PROFILE_DETECTION=YES", StringComparison.Ordinal) ||
                    !genericResult.StdOut.Contains("Integrity gate: PASS", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Desktop generic orchestration smoke failed. exit={genericResult.ExitCode}, stderr={genericResult.StdErr.Trim()}");
                }

                var genericOutput = ExtractPath(genericResult.StdOut, "Generic output: ");
                var genericReport = ExtractPath(genericResult.StdOut, "Generic report: ");
                if (!File.Exists(genericOutput) || !File.Exists(genericReport))
                    throw new InvalidOperationException("Desktop generic orchestration did not produce output/report.");

                lines.Add("DESKTOP_GENERIC_SMOKE=PASS");
                lines.Add("DESKTOP_GENERIC_OUTPUT=" + genericOutput);
                lines.Add("DESKTOP_GENERIC_REPORT=" + genericReport);
            }

            if (args.Any(x => x.Equals("--smoke-template", StringComparison.OrdinalIgnoreCase)))
            {
                var input = GetOption(args, "--smoke-input")
                    ?? throw new InvalidOperationException("--smoke-template requires --smoke-input.");
                var template = GetOption(args, "--smoke-template-file")
                    ?? throw new InvalidOperationException("--smoke-template requires --smoke-template-file.");
                var outDir = GetOption(args, "--smoke-outdir")
                    ?? throw new InvalidOperationException("--smoke-template requires --smoke-outdir.");

                Directory.CreateDirectory(outDir);
                var templateResult = RunEngine(
                    runtime.EngineExe,
                    new[]
                    {
                        "template-apply",
                        "--input", input,
                        "--template", template,
                        "--outdir", outDir,
                        "--confirm-template-apply"
                    },
                    RealTemplateSmokeTimeoutMs);

                if (templateResult.ExitCode != 0 ||
                    !templateResult.StdOut.Contains("NO_AI_SEMANTIC_MAPPING=YES", StringComparison.Ordinal) ||
                    !templateResult.StdOut.Contains("TEMPLATE_APPLY_STATUS=PASS", StringComparison.Ordinal) ||
                    !templateResult.StdOut.Contains("Integrity gate: PASS", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Desktop Template Apply orchestration smoke failed. exit={templateResult.ExitCode}, stderr={templateResult.StdErr.Trim()}");
                }

                var output = ExtractPath(templateResult.StdOut, "Template output: ");
                var report = ExtractPath(templateResult.StdOut, "Template apply report: ");
                if (!File.Exists(output) || !File.Exists(report))
                    throw new InvalidOperationException("Desktop Template Apply orchestration did not produce output/report.");

                lines.Add("DESKTOP_TEMPLATE_APPLY_SMOKE=PASS");
                lines.Add("DESKTOP_TEMPLATE_OUTPUT=" + output);
                lines.Add("DESKTOP_TEMPLATE_REPORT=" + report);
            }

            lines.Add("WORDAGENT_DESKTOP_SMOKE=PASS");
            WriteSmokeOutput(outputPath, lines);
            return 0;
        }
        catch (Exception ex)
        {
            lines.Add("WORDAGENT_DESKTOP_SMOKE=FAIL");
            lines.Add("ERROR=" + ex);
            WriteSmokeOutput(outputPath, lines);
            return 5;
        }
    }

    private static (int ExitCode, string StdOut, string StdErr) RunEngine(
        string engineExe,
        IReadOnlyList<string> arguments,
        int timeoutMs)
    {
        var psi = new ProcessStartInfo(engineExe)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        foreach (var argument in arguments)
            psi.ArgumentList.Add(argument);

        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException("Could not start WordAgent engine.");

        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        if (!process.WaitForExit(timeoutMs))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException("Engine smoke operation timed out.");
        }

        return (
            process.ExitCode,
            stdoutTask.GetAwaiter().GetResult(),
            stderrTask.GetAwaiter().GetResult());
    }

    private static string FirstNonEmptyLine(string text) =>
        text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault()?.Trim() ?? string.Empty;

    private static string ExtractPath(string text, string prefix)
    {
        foreach (var raw in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var line = raw.Trim();
            if (line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return line[prefix.Length..].Trim();
        }

        return string.Empty;
    }

    private static void WriteSmokeOutput(string? outputPath, IReadOnlyCollection<string> lines)
    {
        if (string.IsNullOrWhiteSpace(outputPath))
            return;

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);
            File.WriteAllLines(outputPath, lines, new UTF8Encoding(false));
        }
        catch
        {
            // Smoke exit code remains authoritative if evidence write itself cannot complete.
        }
    }

    private static string? GetOption(string[] args, string option)
    {
        for (var i = 0; i < args.Length - 1; i++)
        {
            if (args[i].Equals(option, StringComparison.OrdinalIgnoreCase))
                return args[i + 1];
        }

        return null;
    }
}
