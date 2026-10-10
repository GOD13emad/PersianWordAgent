using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WordAgent.Desktop;

public sealed record CliRunResult(int ExitCode, string StdOut, string StdErr, TimeSpan Elapsed);

public sealed class CliProcessRunner
{
    private readonly object _sync = new();
    private Process? _active;
    private const int MaxCapturedCharactersPerStream = 2 * 1024 * 1024;

    public bool IsRunning { get { lock (_sync) return _active is not null; } }

    public async Task<CliRunResult> RunAsync(string executable, IReadOnlyList<string> arguments,
        IProgress<string>? progress, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var psi = new ProcessStartInfo(executable)
        {
            UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true,
            CreateNoWindow = true, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
        };
        foreach (var argument in arguments) psi.ArgumentList.Add(argument);
        using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
        lock (_sync)
        {
            if (_active is not null) throw new InvalidOperationException("Another WordAgent operation is already running.");
            _active = process;
        }
        var stdout = new StringBuilder(); var stderr = new StringBuilder(); var watch = Stopwatch.StartNew();
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!process.Start()) throw new InvalidOperationException("Could not start WordAgent engine.");
            using var registration = cancellationToken.Register(() => StopOwned(process));
            var outTask = PumpAsync(process.StandardOutput, stdout, progress, "OUT", process);
            var errTask = PumpAsync(process.StandardError, stderr, progress, "ERR", process);
            try
            {
                await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                StopOwned(process);
                await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5)).ConfigureAwait(false);
                await Task.WhenAll(outTask, errTask).WaitAsync(TimeSpan.FromSeconds(5)).ConfigureAwait(false);
                throw;
            }
            await Task.WhenAll(outTask, errTask).WaitAsync(TimeSpan.FromSeconds(5)).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return new CliRunResult(process.ExitCode, stdout.ToString(), stderr.ToString(), watch.Elapsed);
        }
        finally
        {
            StopOwned(process);
            watch.Stop();
            lock (_sync) { if (ReferenceEquals(_active, process)) _active = null; }
        }
    }

    public void CancelActive() { lock (_sync) { if (_active is not null) StopOwned(_active); } }

    private static void StopOwned(Process process)
    {
        try { if (!process.HasExited) process.Kill(entireProcessTree: true); }
        catch (InvalidOperationException) { }
        catch (System.ComponentModel.Win32Exception) { }
    }

    private static async Task PumpAsync(StreamReader reader, StringBuilder target,
        IProgress<string>? progress, string channel, Process owned)
    {
        var buffer = new char[4096]; var line = new StringBuilder();
        try
        {
            int count;
            while ((count = await reader.ReadAsync(buffer.AsMemory()).ConfigureAwait(false)) != 0)
            {
                if (target.Length + count > MaxCapturedCharactersPerStream)
                    throw new InvalidDataException("ENGINE_OUTPUT_LIMIT_EXCEEDED; operation is not accepted from truncated output.");
                target.Append(buffer, 0, count);
                for (int i = 0; i < count; i++)
                {
                    char c = buffer[i];
                    if (c == '\n') { progress?.Report("[" + channel + "] " + line.ToString().TrimEnd('\r')); line.Clear(); }
                    else
                    {
                        line.Append(c);
                        if (line.Length >= 4096) { progress?.Report("[" + channel + "] " + line); line.Clear(); }
                    }
                }
            }
            if (line.Length > 0) progress?.Report("[" + channel + "] " + line);
        }
        catch { StopOwned(owned); throw; }
    }
}
