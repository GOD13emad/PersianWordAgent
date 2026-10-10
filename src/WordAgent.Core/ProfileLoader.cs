using System.Text.Json;
using System.Text.Json.Nodes;

namespace WordAgent.Core;

public static class ProfileLoader
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };


    // Canonical generic base is intentionally NOT a selectable profile.
    // The only legal inheritance target is the fixed internal .bases/generic.fa.json.
    // This rejects arbitrary relative traversal, cycles and owner supplied paths.
    private static string ResolveEffectiveJson(string profilePath, string profilesDirectory)
    {
        var raw = JsonNode.Parse(File.ReadAllText(profilePath))
            as JsonObject ?? throw new InvalidDataException("Profile root must be an object.");
        if (!raw.ContainsKey("extends"))
            return raw.ToJsonString();

        if (raw["extends"] is not JsonValue value ||
            !value.TryGetValue<string>(out var declared) ||
            !string.Equals(declared, ".bases/generic.fa.json", StringComparison.Ordinal))
            throw new InvalidDataException("PROFILE_BASE_UNTRUSTED: only the internal generic base is supported.");

        var directory = Path.GetFullPath(profilesDirectory);
        var expected = Path.Combine(directory, ".bases", "generic.fa.json");
        var baseDirectory = Path.GetFullPath(Path.Combine(directory, ".bases"));
        if (!string.Equals(Path.GetDirectoryName(Path.GetFullPath(expected)), baseDirectory,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("PROFILE_BASE_PATH_ESCAPE");
        if (!File.Exists(expected))
            throw new InvalidDataException("PROFILE_BASE_MISSING");
        if ((File.GetAttributes(expected) & FileAttributes.ReparsePoint) != 0 ||
            (File.GetAttributes(baseDirectory) & FileAttributes.ReparsePoint) != 0)
            throw new InvalidDataException("PROFILE_BASE_REPARSE_REJECTED");

        var generic = JsonNode.Parse(File.ReadAllText(expected))
            as JsonObject ?? throw new InvalidDataException("PROFILE_BASE_INVALID");
        if (generic.ContainsKey("extends") || generic.ContainsKey("id") ||
            generic.ContainsKey("documentClass") || generic.ContainsKey("classification"))
            throw new InvalidDataException("PROFILE_BASE_UNSAFE_CONTENT");

        var overrides = (JsonObject)raw.DeepClone();
        overrides.Remove("extends");
        MergeInto(generic, overrides);
        return generic.ToJsonString();
    }

    // Object properties merge recursively; arrays and scalar values replace entirely.
    // No role or semantic policy is silently dropped from a specialized profile.
    private static void MergeInto(JsonObject target, JsonObject overrides)
    {
        foreach (var entry in overrides)
        {
            if (target[entry.Key] is JsonObject nestedBase && entry.Value is JsonObject nestedOverride)
                MergeInto(nestedBase, nestedOverride);
            else
                target[entry.Key] = entry.Value?.DeepClone();
        }
    }

    public static AgentProfile Load(string profilesDirectory, string profileId)
    {
        if (string.IsNullOrWhiteSpace(profileId))
            throw new ArgumentException("Profile id is required.", nameof(profileId));

        var direct = Path.Combine(profilesDirectory, profileId + ".json");
        string? path = File.Exists(direct)
            ? direct
            : Directory.Exists(profilesDirectory)
                ? Directory.EnumerateFiles(profilesDirectory, "*.json", SearchOption.TopDirectoryOnly)
                    .FirstOrDefault(p => string.Equals(Path.GetFileNameWithoutExtension(p), profileId, StringComparison.OrdinalIgnoreCase))
                : null;

        if (path is null)
            throw new FileNotFoundException($"Profile '{profileId}' was not found under '{profilesDirectory}'.");

        AgentProfile profile;
        try
        {
            profile = JsonSerializer.Deserialize<AgentProfile>(ResolveEffectiveJson(path, profilesDirectory), JsonOptions)
                ?? throw new InvalidDataException($"Could not parse profile '{path}'.");
        }
        catch (JsonException ex)
        {
            throw new InvalidDataException($"Profile JSON is invalid: '{path}': {ex.Message}", ex);
        }

        if (string.IsNullOrWhiteSpace(profile.Id))
            profile.Id = profileId;

        if (string.IsNullOrWhiteSpace(profile.SchemaVersion))
            profile.SchemaVersion = ProfileSchema.LegacyCompatibility;
        else
            profile.SchemaVersion = profile.SchemaVersion.Trim();

        var validation = ProfileValidator.Validate(profile);
        if (!validation.IsValid)
        {
            var errors = validation.Issues
                .Where(x => x.Severity.Equals("ERROR", StringComparison.OrdinalIgnoreCase))
                .Select(x => $"{x.Code}: {x.Message}");

            throw new InvalidDataException(
                $"Profile '{profile.Id}' failed validation: " + string.Join(" | ", errors));
        }

        return profile;
    }
}
