using System.Diagnostics;
using Microsoft.Win32;

namespace WordAgent.WordInterop;

public static class WordRuntime
{
    public static (bool Installed, string? Version, string? Error) Probe()
    {
        try
        {
            var path = ResolveWinWordPath();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return (false, null, "WINWORD.EXE was not found by COM-free registry/path discovery.");

            var info = FileVersionInfo.GetVersionInfo(path);
            var version = !string.IsNullOrWhiteSpace(info.FileVersion)
                ? info.FileVersion
                : info.ProductVersion;

            return (true, string.IsNullOrWhiteSpace(version) ? null : version, null);
        }
        catch (Exception ex)
        {
            return (false, null, ex.Message);
        }
    }

    private static string? ResolveWinWordPath()
    {
        foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                try
                {
                    using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                    using var appPath = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE");
                    var value = appPath?.GetValue(null) as string;
                    if (!string.IsNullOrWhiteSpace(value) && File.Exists(value))
                        return Path.GetFullPath(value);
                }
                catch { }
            }
        }

        var candidates = new List<string>();
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

        if (!string.IsNullOrWhiteSpace(programFiles))
        {
            candidates.Add(Path.Combine(programFiles, "Microsoft Office", "root", "Office16", "WINWORD.EXE"));
            candidates.Add(Path.Combine(programFiles, "Microsoft Office", "Office16", "WINWORD.EXE"));
        }

        if (!string.IsNullOrWhiteSpace(programFilesX86))
        {
            candidates.Add(Path.Combine(programFilesX86, "Microsoft Office", "root", "Office16", "WINWORD.EXE"));
            candidates.Add(Path.Combine(programFilesX86, "Microsoft Office", "Office16", "WINWORD.EXE"));
        }

        return candidates.FirstOrDefault(File.Exists);
    }
}
