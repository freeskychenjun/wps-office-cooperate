// Cross-process locks shared by the three helpers (word / excel / ppt).
//
// Why this exists: multi-task sessions run several helper processes at once (one daemon per app,
// plus one-shot helpers), and two of them may want the same OS resource at the same time:
//   - the clipboard: Excel's render uses Range.CopyPicture + Clipboard.GetImage, PPT's render uses
//     Slide.Copy + an EMF from the clipboard. Interleaving them yields the wrong picture / a
//     cleared clipboard (measured: cross-app renders produce mixed images).
//   - the attachment-marker JSON files (excel-attached.json / ppt-attached.json): read-modify-write
//     races can drop entries, which then makes a live workbook look "closed" (DOC_CLOSED).
//
// "Local\" = the current login session, which is the right scope for a desktop tool.
// Acquire returns null when the lock cannot be taken in time; callers decide whether to fail
// with a BUSY error or to degrade gracefully.

using System;
using System.Threading;

internal static class WocLocks
{
    public static IDisposable Clipboard(int timeoutMs) { return Acquire("woc-clipboard", timeoutMs); }
    public static IDisposable State(int timeoutMs) { return Acquire("woc-state", timeoutMs); }

    public static IDisposable Acquire(string name, int timeoutMs)
    {
        Mutex m = null;
        try { m = new Mutex(false, @"Local\" + name); }
        catch (Exception) { return null; }
        bool got = false;
        try { got = m.WaitOne(timeoutMs, false); }
        catch (AbandonedMutexException) { got = true; }   // previous owner died: take it over
        catch (Exception) { }
        if (!got)
        {
            try { m.Dispose(); } catch (Exception) { }
            return null;
        }
        return new Holder(m);
    }

    private sealed class Holder : IDisposable
    {
        private Mutex _m;
        public Holder(Mutex m) { _m = m; }
        public void Dispose()
        {
            Mutex m = _m;
            _m = null;
            if (m == null) return;
            try { m.ReleaseMutex(); } catch (Exception) { }
            try { m.Dispose(); } catch (Exception) { }
        }
    }
}
