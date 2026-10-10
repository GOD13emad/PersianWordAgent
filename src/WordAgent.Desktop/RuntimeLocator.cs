using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System.Text;

namespace WordAgent.Desktop;

public sealed record WordAgentRuntime(string Root, string EngineExe, string ProfilesDirectory);

public static class RuntimeLocator
{
    private static readonly string ExpectedEngineVersion = ReleaseContract.ExpectedEngineVersion;
    private const int VersionProbeTimeoutMs = 10000;

    public static bool TryResolve(out WordAgentRuntime runtime, out string error, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var engineOverride = Environment.GetEnvironmentVariable("WORDAGENT_ENGINE_EXE");
        var profilesOverride = Environment.GetEnvironmentVariable("WORDAGENT_PROFILES_DIR");

        if (!string.IsNullOrWhiteSpace(engineOverride))
        {
            if (!File.Exists(engineOverride))
            {
                runtime = EmptyRuntime();
                error = "WORDAGENT_ENGINE_EXE points to a missing file: " + engineOverride;
                return false;
            }

            var engine = Path.GetFullPath(engineOverride);
            var resolvedOverrideProfiles = ResolveProfilesDirectory(engine, profilesOverride);

            if (!TryValidateEngine(engine, out var versionError, cancellationToken))
            {
                runtime = EmptyRuntime();
                error = "Explicit Engine override rejected. " + versionError;
                return false;
            }

            runtime = new WordAgentRuntime(
                Root: Path.GetDirectoryName(engine) ?? string.Empty,
                EngineExe: engine,
                ProfilesDirectory: resolvedOverrideProfiles);
            error = string.Empty;
            return true;
        }

        var root = Path.GetFullPath(AppContext.BaseDirectory);
        var engineCandidate = Path.Combine(root, "dist", "win-x64", "WordAgent.exe");
        if (!File.Exists(engineCandidate))
        {
            runtime = EmptyRuntime();
            error =
                "WordAgent Engine is missing from the release-local path: " + engineCandidate +
                ". Ancestor/current-directory fallback is disabled for release safety. " +
                "Use WORDAGENT_ENGINE_EXE only for an explicit, version-matched override.";
            return false;
        }

        if (!TryValidateEngine(engineCandidate, out var localVersionError, cancellationToken))
        {
            runtime = EmptyRuntime();
            error = "Release-local Engine rejected. " + localVersionError;
            return false;
        }

        var resolvedEngine = Path.GetFullPath(engineCandidate);
        var resolvedReleaseProfiles = ResolveProfilesDirectory(resolvedEngine, profilesOverride);

        runtime = new WordAgentRuntime(
            Root: root,
            EngineExe: resolvedEngine,
            ProfilesDirectory: resolvedReleaseProfiles);
        error = string.Empty;
        return true;
    }

    private static string ResolveProfilesDirectory(string engineExe, string? profilesOverride)
    {
        if (!string.IsNullOrWhiteSpace(profilesOverride) && Directory.Exists(profilesOverride))
            return Path.GetFullPath(profilesOverride);

        var engineDirectory = Path.GetDirectoryName(Path.GetFullPath(engineExe));
        if (string.IsNullOrWhiteSpace(engineDirectory))
            return string.Empty;

        var adjacentProfiles = Path.Combine(engineDirectory, "profiles");
        return Directory.Exists(adjacentProfiles)
            ? Path.GetFullPath(adjacentProfiles)
            : string.Empty;
    }

    private static bool TryValidateEngine(string engineExe, out string error, CancellationToken cancellationToken)
    {
        var psi = new ProcessStartInfo(engineExe)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("version");
        using var process = new Process { StartInfo = psi };
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(VersionProbeTimeoutMs);
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!process.Start()) { error = "Could not start Engine version probe."; return false; }
            var stdoutTask = ReadProbeTextAsync(process.StandardOutput, deadline.Token);
            var stderrTask = ReadProbeTextAsync(process.StandardError, deadline.Token);
            Task.WhenAll(stdoutTask, stderrTask, process.WaitForExitAsync(deadline.Token))
                .WaitAsync(deadline.Token).GetAwaiter().GetResult();
            var version = FirstNonEmptyLine(stdoutTask.GetAwaiter().GetResult());
            if (process.ExitCode != 0)
            {
                error = $"Engine version probe failed. exit={process.ExitCode}, stderr={stderrTask.GetAwaiter().GetResult().Trim()}";
                return false;
            }
            if (!string.Equals(version, ExpectedEngineVersion, StringComparison.Ordinal))
            {
                error = $"Engine version mismatch. expected={ExpectedEngineVersion}, actual={version}";
                return false;
            }
            error = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            try { if (!process.HasExited) { process.Kill(entireProcessTree: true); process.WaitForExit(2000); } } catch { }
            error = deadline.IsCancellationRequested ? "Engine version probe canceled or timed out." : "Engine version probe error: " + ex.Message;
            return false;
        }
    }

    private static async Task<string> ReadProbeTextAsync(StreamReader reader, CancellationToken cancellationToken)
    {
        var text = new StringBuilder();
        var buffer = new char[1024];
        int count;
        while ((count = await reader.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false)) != 0)
        {
            if (text.Length + count > 65536) throw new InvalidDataException("VERSION_PROBE_OUTPUT_LIMIT");
            text.Append(buffer, 0, count);
        }
        return text.ToString();
    }

    private static string FirstNonEmptyLine(string value)
    {
        using var reader = new StringReader(value ?? string.Empty);
        while (reader.ReadLine() is { } line)
        {
            var trimmed = line.Trim();
            if (trimmed.Length != 0)
                return trimmed;
        }

        return string.Empty;
    }

    private static WordAgentRuntime EmptyRuntime() =>
        new(string.Empty, string.Empty, string.Empty);
}
