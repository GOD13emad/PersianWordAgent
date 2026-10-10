using System.Runtime.InteropServices;

namespace WordAgent.WordInterop;

internal static class ComUtil
{
    public static void Release(object? com)
    {
        if (com is null) return;
        try
        {
            if (Marshal.IsComObject(com)) Marshal.FinalReleaseComObject(com);
        }
        catch
        {
            // Best-effort COM cleanup. Word is also closed explicitly by callers.
        }
    }
}
