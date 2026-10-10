namespace WordAgent.Core;

public enum RunMode { AuditOnly, SafeFix, NormalizeToStyles, ReferencesOnly, FieldsRefreshOnly }
public enum RunResult { Pass, PassWithWarnings, FailedPreflight, FailedRules, FailedIntegrity, FailedSave }
public enum RuleRisk { Low, Medium, High, ForbiddenCore }

public sealed record DocumentProfileCandidate(string ProfileId, double Confidence, IReadOnlyList<string> Evidence);
public enum ProfileDetectionStatus { Detected, Ambiguous, NoMatch }
public sealed record ProfileDetectionScore(string ProfileId, int Score, int Priority, IReadOnlyList<string> Evidence);
public sealed record ProfileDetectionResult(
    ProfileDetectionStatus Status,
    string? ProfileId,
    double Confidence,
    IReadOnlyList<ProfileDetectionScore> RankedCandidates);
public sealed record RuleTrace(string RuleId, string ObjectId, string Status, object? Before, object? Action, object? After);

public interface IPreflightService { Task<object> InspectAsync(string path, CancellationToken ct); }
public interface IProfileDetector { Task<DocumentProfileCandidate> DetectAsync(string path, CancellationToken ct); }
public interface IRuleEngine { Task<IReadOnlyList<RuleTrace>> AuditAsync(string path, string profileId, CancellationToken ct); }
public interface IIntegrityGuard { Task<object> CompareAsync(string beforePath, string afterPath, CancellationToken ct); }

