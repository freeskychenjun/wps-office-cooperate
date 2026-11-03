// v1.32 - clipboard keeper: a render puts a picture on the clipboard, and the user's clipboard is
// not ours to lose.
//
// Excel render = Range.CopyPicture -> Clipboard.GetImage -> Clipboard.Clear (EtHelper.cs) and PPT
// render = Slide.Copy -> CF_ENHMETAFILE -> PNG (WppHelper.cs / Deliver.cs). Every one of those steps
// overwrites the machine-wide clipboard, so a single render used to eat whatever the user had copied
// (text, an image, a file list - it all went away).
//
// Usage:
//
//     using (ClipboardKeeper.Keep())
//     {
//         range.CopyPicture(1, 2);
//         ClipboardKeeper.Sync();      // "the state on the clipboard now is mine"
//         ...                          // more clipboard work, Sync() after each write of ours
//     }
//
// Rules (ported from dsh-office helper/OfficeHelper.cs:5768-5825, MIT):
//   - on entry: read the sequence number, then snapshot every format currently on the clipboard;
//   - Sync(): note the sequence number of the state we just created ourselves;
//   - on exit: put the snapshot back only when the clipboard still holds OUR state, i.e. the
//     sequence number has not moved since the last Sync(). If it moved, somebody else (the user)
//     copied something newer meanwhile - theirs wins and we leave it alone;
//   - the restored data is tagged "not for the clipboard history / no cloud upload / no monitor
//     processing", the same trick dsh-office uses, so a render does not push the user's text into
//     Win+V (which is a real leak in a shared office);
//   - a failed restore never throws: the render itself already succeeded, so failures are traced
//     and the picture simply stays on the clipboard.
//
// Nested Keep() calls share the outer keeper's snapshot: only the outermost one restores.

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

internal sealed class ClipboardKeeper : IDisposable
{
    [DllImport("user32.dll")]
    private static extern uint GetClipboardSequenceNumber();

    private const string NoHistory = "CanIncludeInClipboardHistory";
    private const string NoCloud = "CanUploadToCloudClipboard";
    private const string NoMonitor = "ExcludeClipboardContentFromMonitorProcessing";

    private const int TryCount = 5;
    private const int RetryMs = 50;

    /// <summary>The innermost active keeper (only the outermost owns save + restore).</summary>
    private static ClipboardKeeper _open;

    private DataObject _data;
    private int _formats;
    private uint _entry;
    private uint _ours;
    private bool _touched;
    private bool _owner;
    private bool _disposed;

    /// <summary>Start keeping the clipboard; dispose the returned object to put it back.</summary>
    public static ClipboardKeeper Keep()
    {
        ClipboardKeeper keeper = new ClipboardKeeper();
        if (_open != null) return keeper;          // nested: the outer keeper restores, not us
        keeper._owner = true;
        try { keeper._entry = GetClipboardSequenceNumber(); } catch (Exception) { }
        keeper._ours = keeper._entry;
        keeper._data = Snapshot(out keeper._formats);
        _open = keeper;
        Say("keep: seq " + keeper._entry + ", saved " + keeper._formats + " format(s)");
        return keeper;
    }

    /// <summary>Call right after this process wrote the clipboard: from now on that state is ours,
    /// so a sequence number that moved after this point means the user copied something newer.</summary>
    public static void Sync()
    {
        ClipboardKeeper keeper = _open;
        if (keeper == null) return;
        keeper._touched = true;
        try { keeper._ours = GetClipboardSequenceNumber(); } catch (Exception) { }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        if (!_owner) return;
        if (ReferenceEquals(_open, this)) _open = null;
        if (!_touched) return;                     // we never touched it: nothing to put back
        uint now;
        try { now = GetClipboardSequenceNumber(); }
        catch (Exception) { return; }
        if (now != _ours)
        {
            Say("skip restore: the clipboard moved on (ours " + _ours + ", now " + now + ")");
            return;
        }
        Restore();
    }

    private void Restore()
    {
        DataObject data = _data;
        int formats = _formats;
        _data = null;
        for (int attempt = 0; attempt < TryCount; attempt++)
        {
            try
            {
                if (data == null || formats == 0)
                {
                    Clipboard.Clear();
                    Say("restore: nothing was on the clipboard before, cleared ours");
                    return;
                }
                // Keep it out of Win+V / cloud clipboard / monitor pipeline: the user did not copy
                // this again, it is their older content coming back.
                data.SetData(NoMonitor, false, new MemoryStream(new byte[] { 1, 0, 0, 0 }));
                data.SetData(NoHistory, false, new MemoryStream(BitConverter.GetBytes(0)));
                data.SetData(NoCloud, false, new MemoryStream(BitConverter.GetBytes(0)));
                Clipboard.SetDataObject(data, true);
                Say("restore: " + formats + " format(s) put back");
                return;
            }
            catch (Exception e)
            {
                if (attempt == TryCount - 1)
                {
                    string note = "restore failed after " + TryCount + " attempts, the render picture stays on the clipboard: "
                        + e.GetType().Name + ": " + e.Message;
                    Trace.WriteLine("woc-clip: " + note);
                    Say(note);
                }
                try { Thread.Sleep(RetryMs); } catch (Exception) { }
            }
        }
    }

    /// <summary>Copy of the clipboard's SIMPLE formats (text / html / rtf / csv / bitmap / file list
    /// / png).  Anything else - an OLE object, a PowerPoint slide, an embedded document - is left
    /// alone on purpose: asking such a format for its data re-enters the owning application from
    /// this STA thread and has been measured to kill the helper process outright (2026-11-03:
    /// `slide -Action reuse` leaves the duplicated slide on the clipboard, and the next render's
    /// snapshot crashed with "The helper exited without writing a response").  The three
    /// history-control formats are skipped as well: they are re-added on restore.</summary>
    private static readonly string[] SimpleFormats = new string[] {
        DataFormats.Text, DataFormats.UnicodeText, DataFormats.Html, DataFormats.Rtf,
        DataFormats.CommaSeparatedValue, DataFormats.Bitmap, DataFormats.FileDrop,
        DataFormats.StringFormat, "PNG", "DeviceIndependentBitmap"
    };

    private static bool IsSimple(string format)
    {
        if (format == null) return false;
        for (int i = 0; i < SimpleFormats.Length; i++)
            if (string.Equals(format, SimpleFormats[i], StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static DataObject Snapshot(out int formats)
    {
        formats = 0;
        // Nothing simple on the clipboard? Then there is nothing worth risking a read for.
        bool any = false;
        try { any = Clipboard.ContainsText() || Clipboard.ContainsImage() || Clipboard.ContainsFileDropList(); }
        catch (Exception) { any = false; }
        if (!any) { Say("keep: the clipboard holds nothing simple (text/image/files) - left alone"); return null; }
        for (int attempt = 0; attempt < TryCount; attempt++)
        {
            try
            {
                IDataObject now = Clipboard.GetDataObject();
                if (now == null) return null;
                string[] names = now.GetFormats(false);
                if (names == null || names.Length == 0) return null;
                DataObject copy = new DataObject();
                int skipped = 0;
                foreach (string format in names)
                {
                    if (format == null) continue;
                    if (format == NoHistory || format == NoCloud || format == NoMonitor) continue;
                    if (!IsSimple(format)) { skipped++; continue; }
                    object value = null;
                    try { value = now.GetData(format, false); }
                    catch (Exception) { continue; }
                    if (value == null) continue;
                    try { copy.SetData(format, false, value); formats++; }
                    catch (Exception) { }
                }
                if (skipped > 0) Say("keep: skipped " + skipped + " format(s) that are not safe to read (OLE/slide)");
                return copy;
            }
            catch (Exception)
            {
                // The clipboard is held open by another process (WPS copying, an explorer drag):
                // wait a moment and read it again rather than losing it.
                try { Thread.Sleep(RetryMs); } catch (Exception) { }
            }
        }
        return null;
    }

    /// <summary>WOC_CLIP_DEBUG=1 prints what the keeper decided on stderr (silent by default: the
    /// three wrappers promise one JSON document on stdout).</summary>
    private static void Say(string text)
    {
        Trace.WriteLine("woc-clip: " + text);
        if (Environment.GetEnvironmentVariable("WOC_CLIP_DEBUG") != "1") return;
        try { Console.Error.WriteLine("woc-clip: " + text); }
        catch (Exception) { }
    }
}
