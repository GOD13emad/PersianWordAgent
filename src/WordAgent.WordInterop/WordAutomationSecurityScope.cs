using Word = Microsoft.Office.Interop.Word;

namespace WordAgent.WordInterop;

/// <summary>
/// Applies Office's msoAutomationSecurityForceDisable policy (numeric value 3)
/// to a dedicated Word.Application before any document is opened.
/// The previous per-application value is restored on disposal; callers immediately
/// Quit the dedicated application afterward, so a failed best-effort restore cannot
/// leak into another Word process.
/// </summary>
internal sealed class WordAutomationSecurityScope : IDisposable
{
    // Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable.
    // Kept numeric to avoid adding a separate Office Core package solely for this enum.
    private const int ForceDisableValue = 3;

    private dynamic? _application;
    private readonly int _previousValue;
    private bool _disposed;

    private WordAutomationSecurityScope(Word.Application application)
    {
        ArgumentNullException.ThrowIfNull(application);

        _application = application;
        dynamic app = application;
        _previousValue = (int)app.AutomationSecurity;
        app.AutomationSecurity = ForceDisableValue;

        var effectiveValue = (int)app.AutomationSecurity;
        if (effectiveValue != ForceDisableValue)
            throw new InvalidOperationException(
                $"Word AutomationSecurity did not enter ForceDisable mode. actual={effectiveValue}");
    }

    public static WordAutomationSecurityScope ForceDisable(Word.Application application)
        => new(application);

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        var app = _application;
        _application = null;
        if (app is null) return;

        try
        {
            app.AutomationSecurity = _previousValue;
        }
        catch
        {
            // The Word.Application is dedicated to this operation and is Quit immediately
            // after this scope. Never weaken security to compensate for a restore failure.
        }
    }
}
