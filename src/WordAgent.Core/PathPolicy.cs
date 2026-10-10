namespace WordAgent.Core;

public static class PathPolicy
{
    private static readonly string[] SupportedExtensions = [".docx", ".docm", ".dotx", ".dotm"];

    public static string RequireSupportedWordFile(string input)
    {
        var full = Path.GetFullPath(input);
        if (!File.Exists(full))
            throw new FileNotFoundException("Word file not found.", full);
        if (!SupportedExtensions.Contains(Path.GetExtension(full), StringComparer.OrdinalIgnoreCase))
            throw new NotSupportedException("Supported extensions: .docx, .docm, .dotx, .dotm");
        return full;
    }

    public static string CreateSafeOutputPath(string input, string? outputDirectory = null, string suffix = "_WordAgent")
    {
        var full = Path.GetFullPath(input);
        var dir = string.IsNullOrWhiteSpace(outputDirectory) ? Path.GetDirectoryName(full)! : Path.GetFullPath(outputDirectory);
        Directory.CreateDirectory(dir);
        var candidate = Path.Combine(dir, Path.GetFileNameWithoutExtension(full) + suffix + Path.GetExtension(full));
        if (!File.Exists(candidate)) return candidate;

        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        return Path.Combine(dir, Path.GetFileNameWithoutExtension(full) + suffix + "_" + stamp + Path.GetExtension(full));
    }
}
