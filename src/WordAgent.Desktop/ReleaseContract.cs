using System.Reflection;

namespace WordAgent.Desktop;

internal static class ReleaseContract
{
    public static readonly string DesktopVersion = ResolveVersion();
    public static readonly string ExpectedEngineVersion = DesktopVersion;

    private static string ResolveVersion()
        => typeof(ReleaseContract).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion
            ?? throw new InvalidOperationException("Desktop informational version is missing.");
}
