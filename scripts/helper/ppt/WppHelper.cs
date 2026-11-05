// wps-office-cooperate (ppt) helper: drive WPS Presentation (WPP) over COM to read /
// edit / render a .pptx presentation live. Pure ASCII source (Chinese only ever
// travels inside JSON strings), C# 5 (compiled by the in-box csc.exe of .NET
// Framework 4.x), no third-party dependency.
//
//   WppHelper.exe <request.json> <response.json>
//
// The request is {"cmd":"...", ...}; the response is {"ok":true,"result":...} or
// {"ok":false,"error":{"code":"...","message":"..."}}. Files are read and written
// as UTF-8, so no console code page is involved.
//
// Limitations by design (same family as the excel side):
//  - external COM edits do NOT enter the presentation undo stack: snapshot before
//    big batches, verify with render/diff, never promise Ctrl+Z;
//  - render goes through the clipboard (Slide.Copy -> EnhancedMetafile -> PNG
//    written by THIS process, which stays plaintext on disk; Slide.Export is
//    unusable here because files written by wpp.exe are transparently encrypted).

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

internal sealed class Fail : Exception
{
    public readonly string Code;
    public Fail(string code, string message) : base(message) { Code = code; }
    public Fail(string code, string message, object detail) : base(message) { Code = code; Detail = detail; }
    public object Detail;
}

// ---------------------------------------------------------------- JSON

internal sealed class Json
{
    private readonly string s;
    private int i;

    private Json(string text) { s = text; }

    public static object Parse(string text)
    {
        if (text == null) throw new Fail("BAD_ARGS", "no request");
        Json p = new Json(text);
        p.Ws();
        object v = p.Value();
        p.Ws();
        if (p.i < p.s.Length) throw new Fail("BAD_ARGS", "trailing characters at " + p.i);
        return v;
    }

    private void Ws() { while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n')) i++; }

    private object Value()
    {
        if (i >= s.Length) throw new Fail("BAD_ARGS", "unexpected end of JSON");
        char c = s[i];
        if (c == '{') return Obj();
        if (c == '[') return Arr();
        if (c == '"') return Str();
        if (c == 't') { Lit("true"); return true; }
        if (c == 'f') { Lit("false"); return false; }
        if (c == 'n') { Lit("null"); return null; }
        return Num();
    }

    private void Lit(string w)
    {
        if (i + w.Length > s.Length || s.Substring(i, w.Length) != w) throw new Fail("BAD_ARGS", "bad literal at " + i);
        i += w.Length;
    }

    private Dictionary<string, object> Obj()
    {
        Dictionary<string, object> d = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        i++; Ws();
        if (i < s.Length && s[i] == '}') { i++; return d; }
        while (true)
        {
            Ws();
            if (i >= s.Length || s[i] != '"') throw new Fail("BAD_ARGS", "object key expected at " + i);
            string k = Str();
            Ws();
            if (i >= s.Length || s[i] != ':') throw new Fail("BAD_ARGS", "':' expected at " + i);
            i++;
            Ws();
            d[k] = Value();
            Ws();
            if (i >= s.Length) throw new Fail("BAD_ARGS", "unclosed object");
            if (s[i] == ',') { i++; continue; }
            if (s[i] == '}') { i++; return d; }
            throw new Fail("BAD_ARGS", "',' or '}' expected at " + i);
        }
    }

    private List<object> Arr()
    {
        List<object> list = new List<object>();
        i++; Ws();
        if (i < s.Length && s[i] == ']') { i++; return list; }
        while (true)
        {
            Ws();
            list.Add(Value());
            Ws();
            if (i >= s.Length) throw new Fail("BAD_ARGS", "unclosed array");
            if (s[i] == ',') { i++; continue; }
            if (s[i] == ']') { i++; return list; }
            throw new Fail("BAD_ARGS", "',' or ']' expected at " + i);
        }
    }

    private string Str()
    {
        StringBuilder sb = new StringBuilder();
        i++;
        while (true)
        {
            if (i >= s.Length) throw new Fail("BAD_ARGS", "unclosed string");
            char c = s[i++];
            if (c == '"') return sb.ToString();
            if (c != '\\') { sb.Append(c); continue; }
            if (i >= s.Length) throw new Fail("BAD_ARGS", "bad escape");
            char e = s[i++];
            if (e == 'n') sb.Append('\n');
            else if (e == 't') sb.Append('\t');
            else if (e == 'r') sb.Append('\r');
            else if (e == 'b') sb.Append('\b');
            else if (e == 'f') sb.Append('\f');
            else if (e == 'u')
            {
                if (i + 4 > s.Length) throw new Fail("BAD_ARGS", "bad \\u escape");
                sb.Append((char)int.Parse(s.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                i += 4;
            }
            else if (e == '"' || e == '\\' || e == '/') sb.Append(e);
            // Unknown escapes used to be appended verbatim (the backslash dropped), so an unescaped
            // Windows path "C:\Users\A\Desktop\x.pptx" silently became "C:UsersADesktopx.pptx" and
            // open created a brand-new file at that garbage path (reported 2026-11-02). Fail loud.
            else throw new Fail("BAD_ARGS", "invalid JSON escape \\" + e + " in a string - a Windows path must escape its backslashes (C:\\\\dir\\\\file.pptx) or be produced by a JSON serializer");
        }
    }

    private object Num()
    {
        int start = i;
        while (i < s.Length && ("+-0123456789.eE".IndexOf(s[i]) >= 0)) i++;
        if (start == i) throw new Fail("BAD_ARGS", "unexpected character '" + s[start] + "' at " + start);
        string raw = s.Substring(start, i - start);
        if (raw.IndexOf('.') >= 0 || raw.IndexOf('e') >= 0 || raw.IndexOf('E') >= 0)
            return double.Parse(raw, CultureInfo.InvariantCulture);
        long l;
        if (long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out l)) return l;
        return double.Parse(raw, CultureInfo.InvariantCulture);
    }

    // ------------------------------------------------------------ writing

    public static string Write(object o)
    {
        StringBuilder sb = new StringBuilder();
        Write(sb, o);
        return sb.ToString();
    }

    private static void Write(StringBuilder sb, object o)
    {
        if (o == null) { sb.Append("null"); return; }
        if (o is string) { Quote(sb, (string)o); return; }
        if (o is bool) { sb.Append(((bool)o) ? "true" : "false"); return; }
        if (o is int) { sb.Append(((int)o).ToString(CultureInfo.InvariantCulture)); return; }
        if (o is long) { sb.Append(((long)o).ToString(CultureInfo.InvariantCulture)); return; }
        if (o is double) { sb.Append(((double)o).ToString("R", CultureInfo.InvariantCulture)); return; }
        IDictionary<string, object> map = o as IDictionary<string, object>;
        if (map != null)
        {
            sb.Append('{');
            bool first = true;
            foreach (KeyValuePair<string, object> kv in map)
            {
                if (!first) sb.Append(',');
                first = false;
                Quote(sb, kv.Key);
                sb.Append(':');
                Write(sb, kv.Value);
            }
            sb.Append('}');
            return;
        }
        IEnumerable seq = o as IEnumerable;
        if (seq != null)
        {
            sb.Append('[');
            bool first = true;
            foreach (object item in seq)
            {
                if (!first) sb.Append(',');
                first = false;
                Write(sb, item);
            }
            sb.Append(']');
            return;
        }
        Quote(sb, o.ToString());
    }

    private static void Quote(StringBuilder sb, string v)
    {
        sb.Append('"');
        foreach (char c in v)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                default:
                    if (c < ' ') sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                    else sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
    }
}

internal sealed class Bag
{
    private readonly Dictionary<string, object> d;
    public Bag(object o)
    {
        d = o as Dictionary<string, object>;
        if (d == null) throw new Fail("BAD_ARGS", "the request must be a JSON object");
    }
    public bool Has(string k) { return d.ContainsKey(k) && d[k] != null; }
    public string Str(string k, string fallback)
    {
        if (!Has(k)) return fallback;
        object v = d[k];
        return v as string ?? Convert.ToString(v, CultureInfo.InvariantCulture);
    }
    public string Need(string k)
    {
        string v = Str(k, null);
        if (v == null || v.Trim().Length == 0) throw new Fail("BAD_ARGS", "\"" + k + "\" is required");
        return v;
    }
    public object Raw(string k)
    {
        return Has(k) ? d[k] : null;
    }
    public int Int(string k, int fallback)
    {
        if (!Has(k)) return fallback;
        object v = d[k];
        if (v is long) return (int)(long)v;
        if (v is double) return (int)(double)v;
        int n;
        if (int.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), out n)) return n;
        return fallback;
    }
    public double Dbl(string k, double fallback)
    {
        if (!Has(k)) return fallback;
        object v = d[k];
        if (v is long) return (long)v;
        if (v is double) return (double)v;
        double n;
        if (double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out n)) return n;
        return fallback;
    }
    public bool Flag(string k, bool fallback)
    {
        if (!Has(k)) return fallback;
        object v = d[k];
        if (v is bool) return (bool)v;
        string s = Convert.ToString(v, CultureInfo.InvariantCulture);
        return "1".Equals(s) || "true".Equals(s, StringComparison.OrdinalIgnoreCase) || "yes".Equals(s, StringComparison.OrdinalIgnoreCase);
    }
    public List<object> List(string k)
    {
        if (!Has(k)) return new List<object>();
        List<object> l = d[k] as List<object>;
        if (l == null) throw new Fail("BAD_ARGS", "\"" + k + "\" must be an array");
        return l;
    }
    /// <summary>The raw dictionary, for commands that take an open-ended map of values
    /// (slide -Action reuse builds its text/picture swaps out of one).</summary>
    public Dictionary<string, object> Items() { return d; }
}

// ---------------------------------------------------------------- helper

internal static partial class Program
{
    private const string WPS_WPP = "KWPP.Application";

    private static dynamic _app;

    [STAThread]
    private static int Main(string[] args)
    {
        // v1.32: install the OLE message filter before any COM call, so a rejected call while the
        // user is typing or has a dialog open is retried for up to 6 s instead of failing at once
        // (helper/shared/MessageFilter.cs). One registration covers every command: all requests run
        // on this STA thread.
        MessageFilter.Register();
        Console.OutputEncoding = new UTF8Encoding(false);
        // Detached tray daemon mode: same binary, message-loop instead of a COM request.
        if (args != null && args.Length >= 2 && args[0] == "--tray") return TrayDaemon.Run(args[1]);
        // v1.15 resident mode: keep the COM application alive and serve request JSON over a named
        // pipe (same transport as the Word helper - see helper/shared/PipeServer.cs).
        if (args != null && args.Length >= 2 && args[0] == "--serve")
            return PipeServer.Run(args[1], "ppt", HandleRequest, OpenDocumentCount, ReleaseHints);

        string requestPath = args != null && args.Length > 0 ? args[0] : null;
        string responsePath = args != null && args.Length > 1 ? args[1] : null;
        if (requestPath == null)
        {
            Console.Out.Write("{\"ok\":false,\"error\":{\"code\":\"BAD_ARGS\",\"message\":\"usage: WppHelper.exe <request.json> <response.json> | --serve <pipe> | --tray <state>\"}}");
            return 1;
        }
        string text = HandleRequest(File.ReadAllText(requestPath, Encoding.UTF8));
        if (responsePath != null) File.WriteAllText(responsePath, text, new UTF8Encoding(false));
        else Console.Out.Write(text);
        return 0;
    }

    /// <summary>Set by "daemon-stop": the resident loop exits after answering that request.</summary>
    private static bool _daemonStop;

    /// <summary>Extra key/values a pre-flight check wants to report with the current response
    /// (e.g. "the file had been saved with the badge - it was repaired"). Reset per request.</summary>
    private static Dictionary<string, object> _requestNotes;

    private static void BeginRequest()
    {
        _requestNotes = null;
    }

    private static void Note(string key, object value)
    {
        if (_requestNotes == null) _requestNotes = new Dictionary<string, object>();
        _requestNotes[key] = value;
    }

    private static void AttachNotes(object result)
    {
        if (_requestNotes == null || _requestNotes.Count == 0) return;
        Dictionary<string, object> rd = result as Dictionary<string, object>;
        if (rd == null) return;
        foreach (KeyValuePair<string, object> kv in _requestNotes) rd[kv.Key] = kv.Value;
    }

    /// <summary>Open documents in this automation instance (the daemon's exit rule).</summary>
    private static int OpenDocumentCount()
    {
        try
        {
            if (_daemonStop) return 0;
            if (_app == null) return 0;
            return (int)_app.Presentations.Count;
        }
        catch (Exception) { return 0; }
    }

    /// <summary>One request JSON in, one response JSON out (shared by one-shot and resident modes).</summary>
    private static string HandleRequest(string requestJson)
    {
        object response;
        try
        {
            BeginRequest();
            Bag bag = new Bag(Json.Parse(requestJson));
            string cmd = bag.Str("cmd", "").Trim().ToLowerInvariant();
            if (cmd.Length == 0) throw new Fail("BAD_ARGS", "\"cmd\" is required");
            if (cmd != "open" && cmd != "hint" && cmd != "finish" && cmd != "close")
            {
                // v1.36: repair the hidden-window state WPS can leave behind (below) before the
                // takeover mark, so the mark and the repair land in the same request.
                HealWindow();
                EnsureTitleMarker(bag);
                EnsureSlideMarkIntegrity(bag);
            }
            object result = Dispatch(cmd, bag);
            // `open` skips the pre-flight check above (the deck is not open yet) - do it now
            if (cmd == "open") EnsureSlideMarkIntegrity(bag);
            AttachNotes(result);
            response = Ok(cmd, result);
        }
        catch (Fail f)
        {
            response = Err(f.Code, f.Message);
        }
        catch (COMException e)
        {
            string msg = Clean(e.Message);
            bool busy = e.HResult == unchecked((int)0x8001010A)   // RPC_E_SERVERCALL_RETRYFLT
                || e.HResult == unchecked((int)0x80010001)        // RPC_E_CALL_REJECTED
                || e.HResult == unchecked((int)0x800AC472)        // the app is busy (a dialog, a repaint)
                || msg.IndexOf("busy", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("dialog", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("\u5fd9", StringComparison.Ordinal) >= 0            // busy, in Chinese
                || msg.IndexOf("\u5bf9\u8bdd\u6846", StringComparison.Ordinal) >= 0; // dialog, in Chinese
            response = Err(busy ? "BUSY" : "COM_ERROR", msg + " (0x" + e.HResult.ToString("X8", CultureInfo.InvariantCulture) + ")");
            if (!busy) { try { _app = null; } catch (Exception) { } }
        }
        catch (Exception e)
        {
            response = Err("INTERNAL", e.GetType().Name + ": " + Clean(e.Message));
        }

        return Json.Write(response);
    }

    /// <summary>Liveness probe for the resident daemon: no COM, returns at once.</summary>
    private static object Ping()
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["pong"] = true;
        try { d["pid"] = Process.GetCurrentProcess().Id; } catch (Exception) { }
        d["serving"] = _app != null;
        return d;
    }

    private static Dictionary<string, object> Ok(string cmd, object result)
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["ok"] = true;
        d["cmd"] = cmd;
        d["result"] = result;
        return d;
    }

    private static Dictionary<string, object> Err(string code, string message)
    {
        Dictionary<string, object> err = new Dictionary<string, object>();
        err["code"] = code;
        err["message"] = message ?? "";
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["ok"] = false;
        d["error"] = err;
        return d;
    }

    private static object Dispatch(string cmd, Bag a)
    {
        switch (cmd)
        {
            case "ping": return Ping();
            case "daemon-stop": _daemonStop = true; return new Dictionary<string, object> { { "stopped", true } };
            case "doctor": return Doctor(a);
            case "status": return Status(a);
            case "open": return OpenPres(a);
            case "read": return ReadPres(a);
            case "search": return SearchPres(a);
            case "replace": return ReplaceCmd(a);
            case "selection": return SelectionCmd(a);
            case "slide": return SlideCmd(a);
            case "split": return SplitCmd(a);
            case "merge": return MergeCmd(a);
            case "text": return TextCmd(a);
            case "textbox": return TextboxCmd(a);
            case "image": return ImageCmd(a);
            case "picture": return PictureCmd(a);
            case "comment": return CommentCmd(a);
            case "export": return ExportCmd(a);
            case "table": return TableCmd(a);
            case "chart": return ChartCmd(a);
            case "shape": return ShapeCmd(a);
            case "format": return FormatCmd(a);
            case "font": return FontCmd(a);
            case "notes": return NotesCmd(a);
            case "theme": return ThemeCmd(a);
            case "master": return MasterCmd(a);
            case "audit": return AuditCmd(a);
            case "check": return CheckCmd(a);
            case "outline": return OutlineCmd(a);
            case "align": return AlignCmd(a);
            case "zorder": return ZorderCmd(a);
            case "restore": return RestoreCmd(a);
            case "view": return ViewCmd(a);
            case "render": return RenderPres(a);
            case "snapshot": return SnapshotPres(a);
            case "diff": return DiffPres(a);
            case "save": return SavePres(a);
            case "close": return ClosePres(a);
            case "quit": return QuitAll(a);
            case "hint": return HintCmd(a);
            case "finish": return FinishCmd(a);
            // v1.29 layout studio (see Studio.cs)
            case "crop": return CropCmd(a);
            case "smartart": return SmartArtCmd(a);
            case "connect": return ConnectCmd(a);
            case "bg": return BgCmd(a);
            case "anim": return AnimCmd(a);
            default: throw new Fail("BAD_ARGS", "unknown cmd \"" + cmd + "\"; use doctor|status|open|read|search|replace|selection|slide|split|merge|text|textbox|image|picture|comment|export|table|chart|shape|format|font|theme|master|audit|check|outline|align|zorder|notes|view|render|snapshot|diff|restore|save|close|quit|finish|hint|crop|smartart|connect|bg|anim");
        }
    }

    // ------------------------------------------------------------ plumbing

    private static string Clean(string s)
    {
        if (s == null) return "";
        return s.Replace("\r", " ").Replace("\n", " ").Trim();
    }

    private static object TryGet(Func<object> f, object fallback)
    {
        try { return f(); }
        catch (Exception) { return fallback; }
    }

    /// <summary>WPS answers RPC_E_CALL_REJECTED while it is busy (a dialog, a repaint, a
    /// just-created window). Retrying briefly is what a human would do, and it turns a
    /// mysterious failure into a slightly slower success.</summary>
    private static T ComRetry<T>(Func<T> action)
    {
        Exception last = null;
        for (int attempt = 0; attempt < 5; attempt++)
        {
            try { return action(); }
            catch (COMException e)
            {
                last = e;
                bool busy = e.HResult == unchecked((int)0x80010001) || e.HResult == unchecked((int)0x8001010A)
                    || e.HResult == unchecked((int)0x800AC472);
                if (!busy) throw;
            }
            Thread.Sleep(400 + attempt * 300);
        }
        throw last;
    }

    private static bool Installed(string progId)
    {
        try { return Type.GetTypeFromProgID(progId, false) != null; }
        catch (Exception) { return false; }
    }

    private static dynamic Running(string progId)
    {
        try { return Marshal.GetActiveObject(progId); }
        catch (Exception) { return null; }
    }

    /// <summary>
    /// v1.36: is the cached COM object still worth using? A cached Application survives the app it
    /// points at (the user closed WPS, WPS crashed on close, or its frame was replaced).
    /// Presentations.Count is the cheapest liveness probe. Only "the object is gone" HRESULTs
    /// invalidate the cache: busy/rejected calls are transient and must NOT drop a healthy app.
    /// </summary>
    private static bool StillAlive(dynamic app)
    {
        try { int n = (int)app.Presentations.Count; return true; }
        catch (COMException e)
        {
            int hr = e.HResult;
            return !(hr == unchecked((int)0x80010108)    // RPC_E_DISCONNECTED
                  || hr == unchecked((int)0x800401FD)    // CO_E_OBJNOTCONNECTED
                  || hr == unchecked((int)0x800706BA)    // RPC server unavailable
                  || hr == unchecked((int)0x800401E3));  // MK_E_UNAVAILABLE / empty object
        }
        catch (Exception) { return true; }
    }

    private static dynamic App(bool create, bool visible)
    {
        if (_app != null)
        {
            if (StillAlive(_app)) return _app;
            _app = null;   // detached or died underneath us: re-attach below
        }
        dynamic found = Running(WPS_WPP);
        if (found != null) { _app = found; return _app; }
        if (!create) return null;
        if (!Installed(WPS_WPP))
            throw new Fail("NO_APP", "WPS Presentation (KWPP.Application) is not installed on this computer.");
        try { _app = Activator.CreateInstance(Type.GetTypeFromProgID(WPS_WPP, true)); }
        catch (Exception e) { throw new Fail("NO_APP", "Could not start WPS Presentation: " + Clean(e.Message)); }
        try { _app.Visible = visible; } catch (Exception) { }
        return _app;
    }

    private static string FullPathOf(string path)
    {
        try { return Path.GetFullPath(path); }
        catch (Exception) { return path; }
    }

    private static int ProcessCount(string name)
    {
        try { return Process.GetProcessesByName(name).Length; }
        catch (Exception) { return 0; }
    }

    private static dynamic FindOpen(string fullPath)
    {
        dynamic app = App(false, false);
        if (app == null) return null;
        int n = 0;
        try { n = (int)app.Presentations.Count; }
        catch (Exception) { return null; }
        string want = fullPath.ToLowerInvariant();
        for (int k = 1; k <= n; k++)
        {
            dynamic p = null;
            try { p = app.Presentations.Item(k); }
            catch (Exception) { continue; }
            string full = null;
            try { full = (string)p.FullName; }
            catch (Exception) { }
            if (full != null && full.ToLowerInvariant() == want) return p;
        }
        return null;
    }

    private static dynamic RequirePres(Bag a, bool create)
    {
        return RequirePres(a, create, false);
    }

    /// <summary>reopen=true is the `open` command: an explicit open must be able to
    /// bring back a deck the helper remembered and the user closed afterwards (the
    /// DOC_CLOSED guard exists to stop *silent* re-creation in the middle of a batch,
    /// not to block a deliberate reopen).</summary>
    private static dynamic RequirePres(Bag a, bool create, bool reopen)
    {
        string path = a.Need("doc");
        string full = FullPathOf(path);
        dynamic p = FindOpen(full);
        if (p != null) { MarkerSet(full); return p; }
        if (!reopen && MarkerHas(full))
            throw new Fail("DOC_CLOSED",
                "The presentation " + full + " was attached by this skill earlier but is no longer open"
                + " - it was probably closed by the user just now (any unsaved changes were lost with it)."
                + " Do not keep working silently: if the user wants to continue, reopen it explicitly with open -Doc.");
        if (!create) throw new Fail("NOT_OPEN", "The presentation is not open in WPS Presentation: " + full + ". Call open first.");
        dynamic opened = OpenPath(full, a.Flag("show", true), a.Flag("readOnly", false));
        MarkerSet(full);
        return opened;
    }

    // ------------------------------------------------------------ attachment marker

    private static string MarkerPath()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate", "ppt-attached.json");
    }

    private static Dictionary<string, object> MarkerLoad()
    {
        try
        {
            Dictionary<string, object> d = Json.Parse(File.ReadAllText(MarkerPath(), Encoding.UTF8)) as Dictionary<string, object>;
            if (d != null) return d;
        }
        catch (Exception) { }
        return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
    }

    private static void MarkerSave(Dictionary<string, object> m)
    {
        IDisposable g = WocLocks.State(5000);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(MarkerPath()));
            string tmp = MarkerPath() + ".tmp";
            File.WriteAllText(tmp, Json.Write(m), new UTF8Encoding(false));
            File.Copy(tmp, MarkerPath(), true);
            try { File.Delete(tmp); } catch (Exception) { }
        }
        catch (Exception) { }
        finally { if (g != null) g.Dispose(); }
    }

    private static void MarkerSet(string full)
    {
        IDisposable g = WocLocks.State(5000);
        try
        {
            Dictionary<string, object> m = MarkerLoad();
            m[full.ToLowerInvariant()] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            MarkerSave(m);
        }
        finally { if (g != null) g.Dispose(); }
    }

    private static bool MarkerHas(string full)
    {
        return MarkerLoad().ContainsKey(full.ToLowerInvariant());
    }

    private static void MarkerRemove(string full)
    {
        IDisposable g = WocLocks.State(5000);
        try
        {
            Dictionary<string, object> m = MarkerLoad();
            if (m.Remove(full.ToLowerInvariant())) MarkerSave(m);
        }
        finally { if (g != null) g.Dispose(); }
    }

    private static dynamic OpenPath(string full, bool show, bool readOnly)
    {
        dynamic app;
        dynamic p;
        if (!File.Exists(full))
        {
            // PowerPoint's most common task is a NEW deck: opening a missing path creates one
            // (unnamed in memory; the first save -As/<path> gives it that identity).
            app = App(true, show);
            try { p = ComRetry(() => app.Presentations.Add(show ? -1 : 0)); }
            catch (Exception e) { throw new Fail("COM_ERROR", "could not create a new presentation: " + Clean(e.Message)); }
            // Save immediately: an unnamed deck has FullName "演示文稿N" and could never be
            // found again by later commands (each command is a fresh helper process).
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(full));
                ComRetry(() => { p.SaveAs(full, 24); return 0; }); // ppSaveAsOpenXMLPresentation
            }
            catch (Exception e)
            {
                try { p.Close(); } catch (Exception) { }
                throw new Fail("COM_ERROR", "created the deck but could not save it to " + full + ": " + Clean(e.Message));
            }
            return p;
        }
        app = App(true, show);
        try { p = app.Presentations.Open(full, readOnly ? -1 : 0, 0, show ? -1 : 0); }
        catch (COMException e)
        {
            string msg = Clean(e.Message);
            if (msg.IndexOf("read-only", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("read only", StringComparison.OrdinalIgnoreCase) >= 0)
                p = app.Presentations.Open(full, -1, 0, show ? -1 : 0);
            else throw;
        }
        return p;
    }

    // ------------------------------------------------------------ slides / shapes access

    private static dynamic SlideAt(dynamic pres, int index)
    {
        int n = 0;
        try { n = (int)pres.Slides.Count; }
        catch (Exception) { }
        if (index < 1 || index > n)
            throw new Fail("BAD_ARGS", "slide index " + index + " is out of range (1.." + Math.Max(n, 1) + "). Read the deck again if slides were added or deleted.");
        return pres.Slides.Item(index);
    }

    private static string ShapeText(dynamic sh)
    {
        try
        {
            if ((int)sh.HasTextFrame != 0)
            {
                try { if ((int)sh.TextFrame.HasText == 0) return ""; }
                catch (Exception) { }
                return (string)sh.TextFrame.TextRange.Text;
            }
        }
        catch (Exception) { }
        return "";
    }

    /// <summary>True for the agent's own takeover badge (the pill stamped on every slide while a
    /// takeover is active). Internal: read / search / snapshot / audit must never report it as deck
    /// content, otherwise the badge shows up as a "text change" or a "font too small" issue.</summary>
    private static bool IsTakeoverMark(string name)
    {
        return string.Equals(name, MarkShapeName, StringComparison.Ordinal);
    }

    private static object ShapeInfo(dynamic sh, int i, int textLimit)
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["i"] = i;
        d["name"] = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
        d["type"] = Convert.ToString(TryGet(() => (object)((int)sh.Type).ToString(CultureInfo.InvariantCulture), "?"), CultureInfo.InvariantCulture);
        bool isPh = false; int phType = 0;
        try { isPh = (int)sh.IsPlaceholder != 0; } catch (Exception) { }
        if (isPh) phType = (int)TryGet(() => (object)sh.PlaceholderFormat.Type, (object)0);
        d["ph"] = isPh ? (phType == 1 ? "title" : phType == 2 ? "body" : Convert.ToString(phType, CultureInfo.InvariantCulture)) : null;
        d["l"] = Convert.ToString(TryGet(() => (object)Math.Round(Convert.ToDouble(sh.Left, CultureInfo.InvariantCulture), 1), 0), CultureInfo.InvariantCulture);
        d["t"] = Convert.ToString(TryGet(() => (object)Math.Round(Convert.ToDouble(sh.Top, CultureInfo.InvariantCulture), 1), 0), CultureInfo.InvariantCulture);
        d["w"] = Convert.ToString(TryGet(() => (object)Math.Round(Convert.ToDouble(sh.Width, CultureInfo.InvariantCulture), 1), 0), CultureInfo.InvariantCulture);
        d["h"] = Convert.ToString(TryGet(() => (object)Math.Round(Convert.ToDouble(sh.Height, CultureInfo.InvariantCulture), 1), 0), CultureInfo.InvariantCulture);
        string text = ShapeText(sh);
        if (text.Length > textLimit) text = text.Substring(0, textLimit) + "...";
        d["text"] = text.Replace("\r", "\n");
        return d;
    }

    /// <summary>Resolve a shape on a slide: by name (JSON "shape"), by 1-based index, or "title"/"body" placeholder.</summary>
    private static dynamic ResolveShape(dynamic slide, Bag a, string shapeKey, string idxKey, bool need)
    {
        string name = a.Str(shapeKey, null);
        int idx = a.Int(idxKey, 0);
        int count = 0;
        try { count = (int)slide.Shapes.Count; }
        catch (Exception) { }
        if (name != null)
        {
            string want = name.Trim();
            for (int k = 1; k <= count; k++)
            {
                dynamic sh = null;
                try { sh = slide.Shapes.Item(k); }
                catch (Exception) { continue; }
                string n = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                if (string.Equals(n, want, StringComparison.OrdinalIgnoreCase)) return sh;
            }
            if (need) throw new Fail("BAD_ARGS", "no shape named \"" + want + "\" on that slide; read the slide to see shape names");
            return null;
        }
        if (idx > 0)
        {
            if (idx > count) throw new Fail("BAD_ARGS", "shape index " + idx + " out of range (1.." + Math.Max(count, 1) + ")");
            return slide.Shapes.Item(idx);
        }
        // default: title placeholder, else first text shape
        for (int k = 1; k <= count; k++)
        {
            dynamic sh = slide.Shapes.Item(k);
            try { if ((int)sh.IsPlaceholder != 0 && (int)sh.PlaceholderFormat.Type == 1) return sh; }
            catch (Exception) { }
        }
        for (int k = 1; k <= count; k++)
        {
            dynamic sh = slide.Shapes.Item(k);
            if (ShapeText(sh).Length > 0) return sh;
        }
        if (need) throw new Fail("BAD_ARGS", "no text shape found on that slide; pass -Shape <name> or create one with textbox");
        return null;
    }

    // ------------------------------------------------------------ theme
    // A deck keeps its own theme (colours / fonts / footer) between commands: every
    // command is a fresh helper process, so the theme lives in a small JSON file keyed
    // by the deck path rather than in memory.

    private sealed class DeckTheme
    {
        public int Primary = ColorOf("#1F4E79", 0);          // 主色：水利深蓝
        public int Accent = ColorOf("#C99700", 0);           // 强调色：金
        public int Accent2 = ColorOf("#2E75B6", 0);          // second accent (v1.31: two-tone contrast)
        public int Title = ColorOf("#1F4E79", 0);            // 标题色
        public int Text = ColorOf("#333333", 0);             // 正文色
        public int Muted = ColorOf("#808080", 0);            // 页脚/次要
        public int Surface = ColorOf("#F2F5F9", 0);          // card ground (v1.31)
        public int White = ColorOf("#FFFFFF", 0);
        public string Font = "\u5fae\u8f6f\u96c5\u9ed1";         // 微软雅黑
        public string TitleFont = "\u5fae\u8f6f\u96c5\u9ed1";
        public string Footer = "";
        public bool PageNumber = true;
        public bool SkipFooterOnCover = true;
        // v1.31 tokens: card ground + corner radius, and which of the four section looks the
        // section preset uses.  HasCardFill stays false until a theme file names cardFill, so
        // "default = surface" keeps working when surface itself is overridden later.
        public int CardFill = 0;
        public bool HasCardFill = false;
        public double CardRadius = 12;
        public string SectionStyle = "";

        public static DeckTheme Load(dynamic pres)
        {
            string source;
            return LoadWithSource(pres, out source);
        }

        /// <summary>Read the theme: the local theme file first (it is the writable copy), then
        /// the copy stored in the deck's own Tags (v1.32), then the built-in defaults.
        /// <paramref name="source"/> says which one answered, so a command can report it.</summary>
        public static DeckTheme LoadWithSource(dynamic pres, out string source)
        {
            DeckTheme t = new DeckTheme();
            string path = ThemePath(pres);
            if (File.Exists(path))
            {
                source = "file " + path;
                try
                {
                    Bag b = new Bag(Json.Parse(File.ReadAllText(path, Encoding.UTF8)));
                    return FromBag(b, t);
                }
                catch (Exception) { return t; }
            }
            string json = ThemeTagRead(pres);
            if (json != null && json.Trim().Length > 1)
            {
                try
                {
                    Bag b = new Bag(Json.Parse(json));
                    source = "pptx tags (no local theme file on this machine)";
                    return FromBag(b, t);
                }
                catch (Exception) { }
            }
            source = "defaults (no theme file and no wocTheme tag)";
            return t;
        }

        public static DeckTheme FromBag(Bag b, DeckTheme fallback)
        {
            DeckTheme t = new DeckTheme();
            if (fallback != null) t = fallback;
            if (b.Has("primary")) t.Primary = ColorOf(b.Str("primary", null), t.Primary);
            if (b.Has("accent")) t.Accent = ColorOf(b.Str("accent", null), t.Accent);
            if (b.Has("accent2")) t.Accent2 = ColorOf(b.Str("accent2", null), t.Accent2);
            if (b.Has("title")) t.Title = ColorOf(b.Str("title", null), t.Title);
            if (b.Has("text")) t.Text = ColorOf(b.Str("text", null), t.Text);
            if (b.Has("muted")) t.Muted = ColorOf(b.Str("muted", null), t.Muted);
            if (b.Has("surface")) t.Surface = ColorOf(b.Str("surface", null), t.Surface);
            if (b.Has("cardFill"))
            {
                t.CardFill = ColorOf(b.Str("cardFill", null), t.Surface);
                t.HasCardFill = true;
            }
            if (b.Has("cardRadius")) t.CardRadius = Math.Max(0, Math.Min(60, b.Dbl("cardRadius", t.CardRadius)));
            if (b.Has("sectionStyle")) t.SectionStyle = (b.Str("sectionStyle", "") ?? "").Trim().ToLowerInvariant();
            if (b.Has("font")) t.Font = b.Str("font", t.Font);
            if (b.Has("titleFont")) t.TitleFont = b.Str("titleFont", t.TitleFont);
            if (b.Has("footer")) t.Footer = b.Str("footer", t.Footer);
            if (b.Has("pageNumber")) t.PageNumber = b.Flag("pageNumber", t.PageNumber);
            if (b.Has("skipFooterOnCover")) t.SkipFooterOnCover = b.Flag("skipFooterOnCover", t.SkipFooterOnCover);
            return t;
        }

        public Dictionary<string, object> ToBag()
        {
            Dictionary<string, object> d = new Dictionary<string, object>();
            d["primary"] = Hex(Primary);
            d["accent"] = Hex(Accent);
            d["accent2"] = Hex(Accent2);
            d["title"] = Hex(Title);
            d["text"] = Hex(Text);
            d["muted"] = Hex(Muted);
            d["surface"] = Hex(Surface);
            d["cardFill"] = HasCardFill ? Hex(CardFill) : null;
            d["cardRadius"] = CardRadius;
            d["sectionStyle"] = SectionStyle;
            d["font"] = Font;
            d["titleFont"] = TitleFont;
            d["footer"] = Footer;
            d["pageNumber"] = PageNumber;
            d["skipFooterOnCover"] = SkipFooterOnCover;
            return d;
        }

        public static string Hex(int ole)
        {
            int r = ole & 0xFF, g = (ole >> 8) & 0xFF, b = (ole >> 16) & 0xFF;
            return "#" + r.ToString("X2", CultureInfo.InvariantCulture) + g.ToString("X2", CultureInfo.InvariantCulture) + b.ToString("X2", CultureInfo.InvariantCulture);
        }
    }

    private static string ThemeDir()
    {
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate", "ppt-themes");
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static string Sha16(string s)
    {
        using (System.Security.Cryptography.SHA256 sha = System.Security.Cryptography.SHA256.Create())
        {
            byte[] h = sha.ComputeHash(Encoding.UTF8.GetBytes(s));
            StringBuilder sb = new StringBuilder();
            for (int k = 0; k < 8; k++) sb.Append(h[k].ToString("x2", CultureInfo.InvariantCulture));
            return sb.ToString();
        }
    }

    private static string DeckIdentity(dynamic pres)
    {
        string full = Convert.ToString(TryGet(() => (object)pres.FullName, ""), CultureInfo.InvariantCulture);
        if (full == null || full.Length == 0)
            full = Convert.ToString(TryGet(() => (object)pres.Name, "deck"), CultureInfo.InvariantCulture);
        return full.ToLowerInvariant();
    }

    private static string ThemePath(dynamic pres)
    {
        return Path.Combine(ThemeDir(), Sha16(DeckIdentity(pres)) + ".json");
    }

    /// <summary>The tag the theme is mirrored into, so a deck carried to another machine still
    /// knows its own palette (v1.32).</summary>
    private const string ThemeTagName = "wocTheme";

    /// <summary>Read the theme out of the deck's own Tags.  WPS may refuse the whole Tags
    /// collection; a throw here simply means "this deck carries no theme".</summary>
    private static string ThemeTagRead(dynamic pres)
    {
        try
        {
            dynamic tags = pres.Tags;
            if (tags == null) return null;
            int c = (int)tags.Count;
            for (int k = 1; k <= c; k++)
            {
                dynamic t = null;
                try { t = tags.Item(k); } catch (Exception) { continue; }
                string nm = Convert.ToString(TryGet(() => (object)t.Name, ""), CultureInfo.InvariantCulture);
                if (nm == ThemeTagName)
                    return Convert.ToString(TryGet(() => (object)t.Value, ""), CultureInfo.InvariantCulture);
            }
        }
        catch (Exception) { }
        return null;
    }

    /// <summary>Mirror the theme into the deck's Tags.  Returns {supported, readBack, exact} so
    /// the caller can tell "stored in the file too" from "WPS refused it - the local file stays
    /// the only copy" instead of pretending it worked.</summary>
    private static Dictionary<string, object> ThemeTagWrite(dynamic pres, string json, bool dryRun)
    {
        Dictionary<string, object> r = new Dictionary<string, object>();
        if (dryRun)
        {
            // a probe: can we even touch the Tags collection?  (add, read, delete)
            bool ok = false;
            string probe = "{\"probe\":1}";
            try
            {
                dynamic tags = pres.Tags;
                if (tags != null)
                {
                    try { tags.Add(ThemeTagName + "Probe", probe); } catch (Exception) { }
                    int c = (int)tags.Count;
                    for (int k = 1; k <= c; k++)
                    {
                        dynamic t = null;
                        try { t = tags.Item(k); } catch (Exception) { continue; }
                        string nm = Convert.ToString(TryGet(() => (object)t.Name, ""), CultureInfo.InvariantCulture);
                        if (nm == ThemeTagName + "Probe")
                            ok = Convert.ToString(TryGet(() => (object)t.Value, ""), CultureInfo.InvariantCulture) == probe;
                    }
                    try { tags.Delete(ThemeTagName + "Probe"); } catch (Exception) { }
                }
            }
            catch (Exception) { ok = false; }
            r["supported"] = ok;
            r["dryRun"] = true;
            return r;
        }
        bool wrote = false;
        string back = null;
        try
        {
            dynamic tags = pres.Tags;
            if (tags != null)
            {
                try { tags.Add(ThemeTagName, json); wrote = true; }
                catch (Exception)
                {
                    // some builds refuse a duplicate name: drop it, then add again
                    try { tags.Delete(ThemeTagName); } catch (Exception) { }
                    try { tags.Add(ThemeTagName, json); wrote = true; } catch (Exception) { wrote = false; }
                }
                back = ThemeTagRead(pres);
            }
        }
        catch (Exception) { wrote = false; }
        r["supported"] = wrote && back != null && back.Length > 0;
        r["chars"] = back == null ? 0 : back.Length;
        r["exact"] = back == json;
        r["note"] = r["supported"].Equals(true)
            ? "the theme is also stored in the deck (tag wocTheme); save the deck so it travels with it"
            : "WPS refused Presentation.Tags on this build - the theme stays in the local file only";
        return r;
    }


    // ------------------------------------------------------------ colors / text options

    /// <summary>Parse "#RRGGBB" into an OLE colour (0x00BBGGRR: red in the LOW byte).</summary>
    private static int ColorOf(string spec, int fallback)
    {
        if (spec == null) return fallback;
        string s = spec.Trim().TrimStart('#');
        if (s.Length == 6)
        {
            int r = int.Parse(s.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int g = int.Parse(s.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int b = int.Parse(s.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return (b << 16) | (g << 8) | r;   // OLE color = 0x00BBGGRR
        }
        return fallback;
    }

    /// <summary>Apply the shared font options (font/size/bold/italic/color/align) to a TextRange.</summary>
    private static void ApplyFont(dynamic range, Bag a)
    {
        string font = a.Str("font", null);
        if (font != null) { try { range.Font.Name = font; range.Font.NameFarEast = font; } catch (Exception) { } }
        if (a.Has("size")) { try { range.Font.Size = Convert.ToSingle(a.Dbl("size", 18), CultureInfo.InvariantCulture); } catch (Exception) { } }
        if (a.Has("bold")) { try { range.Font.Bold = a.Flag("bold", false) ? -1 : 0; } catch (Exception) { } }
        if (a.Has("italic")) { try { range.Font.Italic = a.Flag("italic", false) ? -1 : 0; } catch (Exception) { } }
        string color = a.Str("color", null);
        if (color != null) { try { range.Font.Color.RGB = ColorOf(color, 0); } catch (Exception) { } }
        string align = a.Str("align", null);
        if (align != null)
        {
            int v = -1; // ppAlignLeft
            string lo = align.ToLowerInvariant();
            if (lo == "center" || lo == "\u5c45\u4e2d") v = 2;
            else if (lo == "right" || lo == "\u5c45\u53f3") v = 3;
            else if (lo == "justify" || lo == "\u4e24\u7aef\u5bf9\u9f50") v = 4;
            try { range.ParagraphFormat.Alignment = v; } catch (Exception) { }
        }
    }

    private static string TextFrom(Bag a)
    {
        string file = a.Str("textFile", null);
        if (file != null)
        {
            if (!File.Exists(file)) throw new Fail("BAD_ARGS", "text file not found: " + file);
            return File.ReadAllText(file, Encoding.UTF8);
        }
        return a.Need("text");
    }

    // ------------------------------------------------------------ doctor / status

    private static object Doctor(Bag a)
    {
        Dictionary<string, object> apps = new Dictionary<string, object>();
        apps["wpsPpt"] = new Dictionary<string, object> { { "progId", WPS_WPP }, { "installed", Installed(WPS_WPP) } };
        apps["msPpt"] = new Dictionary<string, object> { { "progId", "PowerPoint.Application" }, { "installed", Installed("PowerPoint.Application") } };
        dynamic app = Running(WPS_WPP);
        List<object> open = new List<object>();
        if (app != null)
        {
            int n = 0;
            try { n = (int)app.Presentations.Count; }
            catch (Exception) { }
            for (int k = 1; k <= n; k++)
            {
                dynamic p = null;
                try { p = app.Presentations.Item(k); }
                catch (Exception) { continue; }
                open.Add(Convert.ToString(TryGet(() => (object)p.Name, "?"), CultureInfo.InvariantCulture));
            }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["apps"] = apps;
        r["running"] = app != null;
        r["processes"] = new Dictionary<string, object> { { "wpp", ProcessCount("wpp") } };
        r["openPresentations"] = open;
        r["renderNote"] = "render goes through the clipboard (Slide.Copy -> EMF -> PNG); Slide.Export is not used because wpp.exe-written files are transparently encrypted on this machine family.";
        return r;
    }

    private static object Status(Bag a)
    {
        dynamic app = App(false, false);
        List<object> list = new List<object>();
        int count = 0;
        if (app != null)
        {
            try { count = (int)app.Presentations.Count; }
            catch (Exception) { }
            string active = Convert.ToString(TryGet(() => (object)app.ActivePresentation.Name, ""), CultureInfo.InvariantCulture);
            for (int k = 1; k <= count; k++)
            {
                dynamic p = null;
                try { p = app.Presentations.Item(k); }
                catch (Exception) { continue; }
                Dictionary<string, object> d = new Dictionary<string, object>();
                d["name"] = Convert.ToString(TryGet(() => (object)p.Name, "?"), CultureInfo.InvariantCulture);
                d["path"] = Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture);
                d["saved"] = TryGet(() => (object)p.Saved, true);
                d["readOnly"] = TryGet(() => (object)p.ReadOnly, false);
                d["slides"] = TryGet(() => (object)p.Slides.Count, 0);
                d["active"] = Convert.ToString(d["name"], CultureInfo.InvariantCulture) == active;
                list.Add(d);
            }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["running"] = app != null;
        r["presentations"] = list;
        r["processes"] = new Dictionary<string, object> { { "wpp", ProcessCount("wpp") } };
        return r;
    }

    // ------------------------------------------------------------ open / read / search / selection

    private static object OpenPres(Bag a)
    {
        string full = FullPathOf(a.Need("doc"));
        bool existed = FindOpen(full) != null;
        bool fileExisted = File.Exists(full);
        dynamic p = RequirePres(a, true, true);
        if (a.Flag("show", true))
        {
            // An earlier hidden / window-less open leaves the user staring at nothing while we edit.
            try { dynamic app = App(false, false); if (app != null) app.Visible = true; } catch (Exception) { }
            int wins = 0;
            try { wins = (int)p.Windows.Count; } catch (Exception) { }
            if (wins == 0) { try { p.NewWindow(); } catch (Exception) { } }
            try { p.Windows(1).Activate(); } catch (Exception) { }
        }
        if (!a.Flag("noHint", false))
        {
            bool marked = ApplyTitleMarker(p, true);
            try { TitleStateSet(Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture), true); } catch (Exception) { }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)p.Name, "?"), CultureInfo.InvariantCulture);
        r["path"] = Convert.ToString(TryGet(() => (object)p.FullName, full), CultureInfo.InvariantCulture);
        r["readOnly"] = TryGet(() => (object)p.ReadOnly, false);
        r["saved"] = TryGet(() => (object)p.Saved, true);
        r["slides"] = TryGet(() => (object)p.Slides.Count, 0);
        double w = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double h = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        r["slideSize"] = new Dictionary<string, object> { { "w", Math.Round(w, 1) }, { "h", Math.Round(h, 1) }, { "ratio", Math.Round(w / h, 3) } };
        r["how"] = existed ? "reused" : (fileExisted ? "opened" : "created");
        // Opening a path that does not exist creates a deck (that is how "make a new deck" works on
        // this side), but a typo then leaves a junk file behind and the old response looked exactly
        // like a normal open. Say it out loud (reported 2026-11-02: an unescaped Windows path created
        // "C:\UsersAdministratorDesktop新建 PPTX 演示文稿.pptx").
        if (!fileExisted && !existed)
            r["note"] = "the path did not exist, so a NEW deck was created and saved there - check the path if you expected an existing file";
        return r;
    }

    private static object ReadPres(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; }
        catch (Exception) { }
        int from = Math.Max(1, a.Int("from", 1));
        int count = a.Int("count", 0);
        int textLimit = a.Int("textLimit", 160);
        List<object> items = new List<object>();
        int last = count <= 0 ? n : Math.Min(n, from + count - 1);
        for (int k = from; k <= last; k++)
        {
            dynamic s = p.Slides.Item(k);
            Dictionary<string, object> d = new Dictionary<string, object>();
            d["i"] = k;
            d["layout"] = Convert.ToString(TryGet(() => (object)s.Layout, ""), CultureInfo.InvariantCulture);
            List<object> shapes = new List<object>();
            int sc = 0;
            try { sc = (int)s.Shapes.Count; }
            catch (Exception) { }
            bool badge = false;
            for (int j = 1; j <= sc; j++)
            {
                dynamic sh = null;
                try { sh = s.Shapes.Item(j); }
                catch (Exception) { continue; }
                string shName = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                if (IsTakeoverMark(shName)) { badge = true; continue; }
                shapes.Add(ShapeInfo(sh, j, textLimit));
            }
            d["shapes"] = shapes;
            if (badge) d["takeoverBadge"] = true;
            string notes = "";
            try { notes = (string)s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text; }
            catch (Exception) { }
            if (notes == null) notes = "";
            if (notes.Length > textLimit) notes = notes.Substring(0, textLimit) + "...";
            if (notes.Trim().Length > 0) d["notes"] = notes.Replace("\r", "\n");
            items.Add(d);
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)p.Name, "?"), CultureInfo.InvariantCulture);
        r["slides"] = n;
        r["from"] = from;
        r["to"] = last;
        r["items"] = items;
        if (from > 1 || last < n) r["note"] = "showing slides " + from + ".." + last + " of " + n + "; use -From/-Count to page";
        return r;
    }

    private static object SearchPres(Bag a)
    {
        dynamic p = RequirePres(a, true);
        string needle = a.Need("text");
        int max = a.Int("max", 50);
        bool icase = !a.Flag("case", false);
        List<object> hits = new List<object>();
        int n = 0;
        try { n = (int)p.Slides.Count; }
        catch (Exception) { }
        for (int k = 1; k <= n && hits.Count < max; k++)
        {
            dynamic s = p.Slides.Item(k);
            int sc = 0;
            try { sc = (int)s.Shapes.Count; }
            catch (Exception) { }
            for (int j = 1; j <= sc && hits.Count < max; j++)
            {
                dynamic sh = null;
                try { sh = s.Shapes.Item(j); }
                catch (Exception) { continue; }
                string shName = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                if (IsTakeoverMark(shName)) continue;
                string text = ShapeText(sh);
                if (text.Length == 0) continue;
                int pos = icase ? text.ToLowerInvariant().IndexOf(needle.ToLowerInvariant(), StringComparison.Ordinal)
                                : text.IndexOf(needle, StringComparison.Ordinal);
                if (pos < 0) continue;
                int start = Math.Max(0, pos - 30);
                int len = Math.Min(text.Length - start, needle.Length + 60);
                hits.Add(new Dictionary<string, object> {
                    { "slide", k },
                    { "shape", shName },
                    { "snippet", text.Substring(start, len).Replace("\r", "\n") }
                });
            }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["hits"] = hits;
        r["count"] = hits.Count;
        return r;
    }

    private static object SelectionCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic app = App(false, false);
        Dictionary<string, object> r = new Dictionary<string, object>();
        dynamic win = TryGet(() => (object)app.ActiveWindow, null) as object;
        if (win == null)
        {
            r["note"] = "no active window (presentation open without -Show); selection unknown";
            return r;
        }
        int cur = 0;
        try { cur = (int)win.Selection.SlideRange.SlideIndex; }
        catch (Exception) { }
        r["slide"] = cur;
        List<object> sel = new List<object>();
        try
        {
            int c = (int)win.Selection.ShapeRange.Count;
            for (int k = 1; k <= c; k++)
                sel.Add(Convert.ToString(TryGet(() => (object)win.Selection.ShapeRange.Item(k).Name, "?"), CultureInfo.InvariantCulture));
        }
        catch (Exception) { }
        r["selectedShapes"] = sel;
        return r;
    }

    // ------------------------------------------------------------ slide management (+ presets)

    private static object SlideCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        string action = a.Str("action", "add").ToLowerInvariant();
        int n = 0;
        try { n = (int)p.Slides.Count; }
        catch (Exception) { }

        if (action == "add")
        {
            int at = a.Int("at", n + 1);
            if (at < 1 || at > n + 1) throw new Fail("BAD_ARGS", "-At " + at + " out of range (1.." + (n + 1) + ")");
            string preset = a.Str("preset", null);
            int layout = a.Int("layout", 12); // ppLayoutBlank by default
            if (preset != null) layout = 12;
            dynamic s = p.Slides.Add(at, layout);
            Dictionary<string, object> r = new Dictionary<string, object>();
            r["added"] = at;
            r["slides"] = n + 1;
            if (preset != null)
            {
                List<string> problems;
                r["preset"] = BuildPreset(p, s, a, out problems);
                if (problems != null && problems.Count > 0) r["presetNotes"] = problems;
            }
            return r;
        }
        if (action == "dup")
        {
            int index = a.Int("index", 0);
            if (index < 1 || index > n) throw new Fail("BAD_ARGS", "-Index " + index + " out of range (1.." + n + ")");
            int at = a.Int("at", n + 1);
            // Slide.Duplicate() goes through the clipboard and leaves the slide there.  Without this
            // the user's copied text/image is gone AND the next render's clipboard snapshot meets a
            // slide format it must not read (that killed the helper process - 2026-11-03).  Keep the
            // simple content, do the duplicate, then put the user's clipboard back.
            dynamic s = null;
            using (ClipboardKeeper.Keep())
            {
                s = p.Slides.Item(index).Duplicate();
                ClipboardKeeper.Sync();
            }
            int newIdx = (int)TryGet(() => (object)s.SlideIndex, n + 1);
            if (at != newIdx) { try { s.MoveTo(at); newIdx = at; } catch (Exception) { } }
            return new Dictionary<string, object> { { "duplicated", index }, { "added", newIdx }, { "slides", n + 1 } };
        }
        if (action == "move")
        {
            int index = a.Int("index", 0);
            int to = a.Int("to", 0);
            if (index < 1 || index > n) throw new Fail("BAD_ARGS", "-Index out of range");
            if (to < 1 || to > n) throw new Fail("BAD_ARGS", "-To out of range");
            p.Slides.Item(index).MoveTo(to);
            return new Dictionary<string, object> { { "moved", index }, { "to", to }, { "slides", n } };
        }
        if (action == "delete")
        {
            int index = a.Int("index", 0);
            if (index < 1 || index > n) throw new Fail("BAD_ARGS", "-Index out of range");
            p.Slides.Item(index).Delete();
            return new Dictionary<string, object> { { "deleted", index }, { "slides", n - 1 } };
        }
        if (action == "reuse") return KReuse(a);
        throw new Fail("BAD_ARGS", "slide -Action must be add|dup|move|delete|reuse");
    }

    /// <summary>Preset layouts.  The five original ones (title / section / bullets / two / end) keep
    /// their look; v1.31 adds eight finished layouts (cards / stats / timeline / process / compare /
    /// summary / gallery / split) drawn on a 960x540 design canvas and scaled to the deck (see
    /// Kinds.cs).  Both families return {preset, slide, shapes, overflowWarnings, ..}.</summary>
    private static object BuildPreset(dynamic pres, dynamic slide, Bag a, out List<string> problems)
    {
        problems = new List<string>();
        List<string> notes = problems;
        string preset = a.Str("preset", "").ToLowerInvariant();
        DeckTheme th = DeckTheme.Load(pres);
        if (IsFinishedPreset(preset))
        {
            // A layout that fails half way through leaves the slide deleted, not half drawn
            // (measured: a failed picture layout used to leave an orphan page behind).
            try { return BuildFinishedSlide(pres, slide, a, preset, th, out problems); }
            catch (Exception)
            {
                try { slide.Delete(); } catch (Exception) { }
                throw;
            }
        }
        string made = BuildLegacyPreset(pres, slide, a, preset, th, notes);
        return PresetReport(pres, slide, made, problems);
    }

    private static bool IsFinishedPreset(string p)
    {
        return p == "cards" || p == "stats" || p == "timeline" || p == "process"
            || p == "compare" || p == "summary" || p == "gallery" || p == "split"
            || p == "canvas" || p == "grid";
    }

    /// <summary>Preset layouts (borrowed idea from dsh-office's ready-made slides, deliberately simple):
    /// title / section / bullets / two / end. Geometry derives from the deck's slide size.</summary>
    private static string BuildLegacyPreset(dynamic pres, dynamic slide, Bag a, string preset, DeckTheme th, List<string> problems)
    {
        List<string> notes = problems;   // C#5: an out parameter cannot be captured by a lambda
        double W = Convert.ToDouble(TryGet(() => (object)pres.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double H = Convert.ToDouble(TryGet(() => (object)pres.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        string title = a.Str("title", "");
        if (preset != "end" && title.Length == 0) throw new Fail("BAD_ARGS", "presets need -Title (except end)");

        // colour block / rule
        Action<double, double, double, double, int, string> bar =
            delegate(double l, double t, double w, double h, int fill, string barName)
        {
            dynamic shp = slide.Shapes.AddShape(1, (float)l, (float)t, (float)w, (float)h);  // msoShapeRectangle
            try { shp.Fill.Solid(); shp.Fill.ForeColor.RGB = fill; } catch (Exception) { }
            try { shp.Line.Visible = 0; } catch (Exception) { }
            if (barName != null) { try { shp.Name = barName; } catch (Exception) { } }
        };

        // text box: geometry, text, size, bold(-1/0), align, colour, font overrides, bullets
        Action<double, double, double, double, string, float, int, string, int, string, string, bool> box =
            delegate(double l, double t, double w, double h, string text, float size, int bold, string align, int color, string boxName, string fontName, bool bulleted)
        {
            if (text == null || text.Length == 0) return;
            dynamic tb = slide.Shapes.AddTextbox(1, (float)l, (float)t, (float)w, (float)h);
            if (boxName != null) { try { tb.Name = boxName; } catch (Exception) { } }
            dynamic tr = tb.TextFrame.TextRange;
            // \r = a real paragraph mark (a \n stays an in-paragraph line break in WPS, which
            // would give the whole block a single bullet); bullets therefore need \r.
            tr.Text = bulleted ? text.Replace("\\n", "\r").Replace("\n", "\r") : text.Replace("\\n", "\n");
            size = AutoFitSize(text, size, w, h);   // shrink instead of overflowing
            try
            {
                tr.Font.Size = size;
                tr.Font.Bold = bold;
                tr.Font.Color.RGB = color;
                string fn = fontName != null ? fontName : th.Font;
                tr.Font.Name = fn;
                tr.Font.NameFarEast = fn;
            }
            catch (Exception) { }
            int al = align == "center" ? 2 : align == "right" ? 3 : 1;
            try { tr.ParagraphFormat.Alignment = al; } catch (Exception) { }
            try { tb.TextFrame.WordWrap = -1; } catch (Exception) { }
            if (bulleted)
            {
                try
                {
                    // per paragraph: setting ParagraphFormat once only reached the first one
                    int pc = (int)tr.Paragraphs().Count;
                    for (int pi = 1; pi <= pc; pi++)
                    {
                        dynamic pf = tr.Paragraphs(pi).ParagraphFormat;
                        pf.Bullet.Visible = -1;
                        pf.Bullet.Type = 1;          // ppBulletUnnumbered
                        pf.Bullet.Character = 8226;   // bullet dot
                        pf.Bullet.RelativeSize = 0.9;
                        try { pf.SpaceWithin = 1.25f; } catch (Exception) { }
                        try { pf.SpaceAfter = 10f; } catch (Exception) { }
                    }
                }
                catch (Exception) { notes.Add("bullet formatting not accepted by this WPS build; text was written without bullets"); }
            }
        };

        List<object> items = a.List("items");
        StringBuilder lines = new StringBuilder();
        foreach (object it in items)
        {
            if (it is string) { lines.Append((string)it); }
            else
            {
                Bag ib = null;
                try { ib = new Bag(it); }
                catch (Exception) { }
                if (ib != null) lines.Append(ib.Str("text", ""));
                else lines.Append(Convert.ToString(it, CultureInfo.InvariantCulture));
            }
            lines.Append("\n");
        }
        string bodyText = lines.ToString().TrimEnd('\n');
        string subtitle = a.Str("subtitle", null);
        string number = a.Str("number", null);

        if (preset == "title")
        {
            bar(0, H * 0.74, W, H * 0.26, th.Primary, "coverBand");
            box(W * 0.08, H * 0.26, W * 0.84, H * 0.26, title, 40, -1, "center", th.Title, "title", th.TitleFont, false);
            bar(W * 0.42, H * 0.545, W * 0.16, 5, th.Accent, "accentLine");
            box(W * 0.08, H * 0.77, W * 0.84, H * 0.18, subtitle, 20, 0, "center", th.White, "subtitle", th.Font, false);
            return "title";
        }
        if (preset == "section")
        {
            // v1.31: four looks, picked by the theme token sectionStyle (solid|side|band|number).
            // Default (empty) keeps the original left band.
            string style = th.SectionStyle == null ? "" : th.SectionStyle;
            if (style == "solid")
            {
                bar(0, 0, W, H, th.Primary, "sectionGround");
                box(W * 0.08, H * 0.26, W * 0.34, H * 0.20, number, 72, -1, "left", th.White, "number", th.TitleFont, false);
                bar(W * 0.08, H * 0.50, W * 0.10, 5, th.Accent, "accentLine");
                box(W * 0.08, H * 0.55, W * 0.84, H * 0.22, title, 40, -1, "left", th.White, "sectionTitle", th.TitleFont, false);
                return "section";
            }
            if (style == "band")
            {
                bar(0, H * 0.33, W, H * 0.30, th.Primary, "sectionBand");
                box(W * 0.06, H * 0.33, W * 0.16, H * 0.30, number, 44, -1, "center", th.White, "number", th.TitleFont, false);
                box(W * 0.24, H * 0.33, W * 0.70, H * 0.30, title, 34, -1, "left", th.White, "sectionTitle", th.TitleFont, false);
                bar(W * 0.06, H * 0.70, W * 0.10, 5, th.Accent, "accentLine");
                return "section";
            }
            if (style == "number")
            {
                box(W * 0.48, H * 0.10, W * 0.46, H * 0.64, number, 220, -1, "right",
                    KMix(th.Primary, th.White, 0.86), "number", th.TitleFont, false);
                bar(W * 0.08, H * 0.40, W * 0.08, 5, th.Accent, "accentLine");
                box(W * 0.08, H * 0.46, W * 0.38, H * 0.22, title, 36, -1, "left", th.Title, "sectionTitle", th.TitleFont, false);
                return "section";
            }
            bar(0, 0, W * 0.30, H, th.Primary, "sectionBand");
            box(0, H * 0.34, W * 0.30, H * 0.22, number, 54, -1, "center", th.White, "number", th.TitleFont, false);
            box(W * 0.38, H * 0.40, W * 0.56, H * 0.18, title, 36, -1, "left", th.Title, "sectionTitle", th.TitleFont, false);
            bar(W * 0.38, H * 0.605, W * 0.10, 5, th.Accent, "accentLine");
            return "section";
        }
        if (preset == "bullets")
        {
            box(W * 0.06, H * 0.055, W * 0.88, H * 0.13, title, 30, -1, "left", th.Title, "title", th.TitleFont, false);
            bar(W * 0.06, H * 0.195, W * 0.09, 5, th.Accent, "accentLine");
            // body stops at 78% so a table or a picture can still go underneath without collisions
            box(W * 0.08, H * 0.25, W * 0.84, H * 0.50, bodyText, 20, 0, "left", th.Text, "body", th.Font, true);
            return "bullets";
        }
        if (preset == "two")
        {
            box(W * 0.06, H * 0.055, W * 0.88, H * 0.13, title, 30, -1, "left", th.Title, "title", th.TitleFont, false);
            bar(W * 0.06, H * 0.195, W * 0.09, 5, th.Accent, "accentLine");
            string leftText = items.Count > 0 ? Convert.ToString(items[0], CultureInfo.InvariantCulture) : "";
            string rightText = items.Count > 1 ? Convert.ToString(items[1], CultureInfo.InvariantCulture) : "";
            box(W * 0.05, H * 0.26, W * 0.42, H * 0.62, leftText, 18, 0, "left", th.Text, "left", th.Font, true);
            bar(W * 0.495, H * 0.28, 2, H * 0.56, th.Muted, "divider");
            box(W * 0.53, H * 0.26, W * 0.42, H * 0.62, rightText, 18, 0, "left", th.Text, "right", th.Font, true);
            return "two";
        }
        if (preset == "end")
        {
            bar(0, 0, W, H, th.Primary, "endBackground");
            box(W * 0.08, H * 0.34, W * 0.84, H * 0.22, a.Str("title", "\u8c22\u8c22"), 40, -1, "center", th.White, "title", th.TitleFont, false);
            bar(W * 0.42, H * 0.60, W * 0.16, 5, th.Accent, "accentLine");
            box(W * 0.08, H * 0.66, W * 0.84, H * 0.14, subtitle, 18, 0, "center", th.White, "subtitle", th.Font, false);
            return "end";
        }
        throw new Fail("BAD_ARGS", "unknown preset \"" + preset + "\"; use title|section|bullets|two|end"
            + " or a finished layout cards|stats|timeline|process|compare|summary|gallery|split"
            + " (or -Layout <n> for a raw PowerPoint layout)");
    }

    /// <summary>Rough Chinese-text capacity check: shrink the font instead of letting text spill
    /// out of its box (a spilled box is the most common defect in generated decks).</summary>
    private static float AutoFitSize(string text, float size, double w, double h)
    {
        if (text == null || text.Length == 0 || w <= 0 || h <= 0) return size;
        // CJK glyph ≈ 1.0 em wide, Latin ≈ 0.55 em; line height ≈ 1.3 em
        double em = 0;
        foreach (char c in text)
        {
            if (c == '\n' || c == '\r') continue;
            em += c > 0x2E80 ? 1.0 : 0.55;
        }
        double perLine = Math.Max(1.0, w / size);
        double lineCount = Math.Ceiling(em / perLine);
        double needed = lineCount * size * 1.35;
        if (needed <= h) return size;
        float shrunk = (float)Math.Floor(size * (h / needed) * 10) / 10f;
        return Math.Max(9f, shrunk);
    }


    // ------------------------------------------------------------ text batch (anchor-checked, atomic)

    private static object TextCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        List<object> items = a.List("items");
        if (items.Count == 0)
        {
            // single-shot form: -Slide N [-Shape name] (-Text | -TextFile) [-Expect ...]
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["slide"] = a.Int("slide", 0);
            if (a.Str("shape", null) != null) one["shape"] = a.Str("shape", null);
            if (a.Str("expect", null) != null) one["expect"] = a.Str("expect", null);
            string file = a.Str("textFile", null);
            one[file != null ? "textFile" : "text"] = file != null ? file : TextFrom(a);
            items.Add(one);
        }

        // pass 1: resolve + validate everything first (anchor semantics like revise/write)
        List<dynamic> shapePlan = new List<dynamic>();
        List<string> newTexts = new List<string>();
        List<string> oldTexts = new List<string>();
        foreach (object raw in items)
        {
            Bag it = new Bag(raw);
            int slideIdx = it.Int("slide", 0);
            if (slideIdx < 1) throw new Fail("BAD_ARGS", "every text item needs \"slide\" (1-based)");
            dynamic slide = SlideAt(p, slideIdx);
            string shapeName = it.Str("shape", null);
            dynamic sh = null;
            int count = 0;
            try { count = (int)slide.Shapes.Count; }
            catch (Exception) { }
            if (shapeName != null)
            {
                sh = ResolveShape(slide, it, "shape", "shapeIdx", true);
            }
            else
            {
                // default resolution on the item bag: title placeholder, else first text shape
                for (int k = 1; k <= count; k++)
                {
                    dynamic cand = slide.Shapes.Item(k);
                    try { if ((int)cand.IsPlaceholder != 0 && (int)cand.PlaceholderFormat.Type == 1) { sh = cand; break; } }
                    catch (Exception) { }
                }
                if (sh == null)
                    for (int k = 1; k <= count; k++)
                    {
                        dynamic cand = slide.Shapes.Item(k);
                        if (ShapeText(cand).Length > 0) { sh = cand; break; }
                    }
                if (sh == null) throw new Fail("BAD_ARGS", "no text shape on slide " + slideIdx + "; add \"shape\" or create a textbox first");
            }
            string cur = ShapeText(sh);
            string expect = it.Str("expect", null);
            if (expect != null && !cur.StartsWith(expect, StringComparison.Ordinal))
                throw new Fail("ANCHOR_MISSING", "slide " + slideIdx + " shape \"" + shapeName + "\" does not start with \"" + expect + "\" (current: \"" + Clip(cur, 60) + "\"); nothing was changed");
            string newText = null;
            string tf = it.Str("textFile", null);
            if (tf != null)
            {
                if (!File.Exists(tf)) throw new Fail("BAD_ARGS", "text file not found: " + tf);
                newText = File.ReadAllText(tf, Encoding.UTF8);
            }
            else newText = it.Need("text");
            shapePlan.Add(sh);
            oldTexts.Add(cur);
            newTexts.Add(newText);
        }

        // pass 2: apply, roll back on failure
        int done = 0;
        try
        {
            for (int k = 0; k < shapePlan.Count; k++)
            {
                shapePlan[k].TextFrame.TextRange.Text = newTexts[k];
                done++;
            }
        }
        catch (Exception e)
        {
            int rolled = 0;
            for (int k = done - 1; k >= 0; k--)
            {
                try { shapePlan[k].TextFrame.TextRange.Text = oldTexts[k]; rolled++; }
                catch (Exception) { }
            }
            throw new Fail("COM_ERROR", "text batch stopped after " + done + " item(s): " + Clean(e.Message) + "; rolled back " + rolled + " item(s)");
        }

        List<object> applied = new List<object>();
        for (int k = 0; k < shapePlan.Count; k++)
        {
            Bag it = new Bag(items[k]);
            string after = Clip(ShapeText(shapePlan[k]), 90);
            applied.Add(new Dictionary<string, object> {
                { "slide", it.Int("slide", 0) },
                { "shape", Convert.ToString(TryGet(() => (object)shapePlan[k].Name, ""), CultureInfo.InvariantCulture) },
                { "before", Clip(oldTexts[k], 60) },
                { "textAfter", after }
            });
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["applied"] = applied;
        r["count"] = applied.Count;
        r["note"] = "compare every textAfter against what you intended - that is the cheapest way to catch a wrong-shape write";
        return r;
    }

    private static string Clip(string s, int n)
    {
        if (s == null) return "";
        s = s.Replace("\r", " ").Replace("\n", " ");
        return s.Length > n ? s.Substring(0, n) + "..." : s;
    }

    // ------------------------------------------------------------ textbox / image / table / shape

    private static object TextboxCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic slide = SlideAt(p, a.Int("slide", 0));
        double W = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double H = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        double l = a.Dbl("l", W * 0.08), t = a.Dbl("t", H * 0.35), w = a.Dbl("w", W * 0.84), h = a.Dbl("h", H * 0.2);
        dynamic tb = slide.Shapes.AddTextbox(1, (float)l, (float)t, (float)w, (float)h);
        string boxName = a.Str("name", null);
        if (boxName != null) { try { tb.Name = boxName; } catch (Exception) { } }
        dynamic tr = tb.TextFrame.TextRange;
        tr.Text = TextFrom(a).Replace("\\n", "\n");
        ApplyFont(tr, a);
        try { tb.TextFrame.WordWrap = -1; } catch (Exception) { }
        string name = Convert.ToString(TryGet(() => (object)tb.Name, ""), CultureInfo.InvariantCulture);
        return new Dictionary<string, object> {
            { "added", name }, { "slide", a.Int("slide", 0) },
            { "textAfter", Clip(ShapeText(tb), 90) },
            { "note", "render the slide to check the box is where you wanted and the text fits" }
        };
    }

    private static object ImageCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic slide = SlideAt(p, a.Int("slide", 0));
        string path = a.Need("path");
        if (!File.Exists(path)) throw new Fail("BAD_ARGS", "image not found: " + path);
        double W = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double H = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        double w = a.Dbl("w", 0), h = a.Dbl("h", 0);
        object wArg = w > 0 ? (object)(float)w : Type.Missing;
        object hArg = h > 0 ? (object)(float)h : Type.Missing;
        double l = a.Dbl("l", -1), t = a.Dbl("t", -1);
        if (l < 0 && w > 0) l = (W - w) / 2;   // default: centered
        if (t < 0 && h > 0) t = (H - h) / 2;
        if (l < 0) l = W * 0.1;
        if (t < 0) t = H * 0.1;
        // AddPicture refuses a non-ASCII path (this machine's work folder alone contains
        // Chinese: 0x800C0005 / "Illegal characters in path"), so insert through an ASCII copy.
        string insertPath = AsciiTempCopy(path);
        dynamic pic;
        try { pic = slide.Shapes.AddPicture(insertPath, 0, -1, (float)l, (float)t, wArg, hArg); }
        catch (COMException e) { throw new Fail("COM_ERROR", "AddPicture failed: " + Clean(e.Message) + " (if the file is encrypted, open it once in WPS or decrypt first)"); }
        finally { TryDelete(insertPath); }
        string name = Convert.ToString(TryGet(() => (object)pic.Name, ""), CultureInfo.InvariantCulture);
        double fw = Convert.ToDouble(TryGet(() => (object)pic.Width, (object)0.0), CultureInfo.InvariantCulture);
        double fh = Convert.ToDouble(TryGet(() => (object)pic.Height, (object)0.0), CultureInfo.InvariantCulture);
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["added"] = name;
        r["slide"] = a.Int("slide", 0);
        r["size"] = new Dictionary<string, object> { { "w", Math.Round(fw, 1) }, { "h", Math.Round(fh, 1) } };
        if (fw < W * 0.2) r["note"] = "picture occupies less than 20% of the slide width - check with render whether its lettering is still readable";
        return r;
    }

    private static object TableCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic slide = SlideAt(p, a.Int("slide", 0));
        // data: -Data <json file> {"rows":[["h1","h2"],[...]]} or inline -DataFile
        string dataFile = a.Str("dataFile", a.Str("data", null));
        if (dataFile == null) throw new Fail("BAD_ARGS", "table needs -Data <json file with {\"rows\":[[..]]}>");
        if (!File.Exists(dataFile)) throw new Fail("BAD_ARGS", "data file not found: " + dataFile);
        Bag data = new Bag(Json.Parse(File.ReadAllText(dataFile, Encoding.UTF8)));
        List<object> rows = data.List("rows");
        if (rows.Count < 1) throw new Fail("BAD_ARGS", "\"rows\" must have at least one row");
        int cols = 0;
        foreach (object row in rows) cols = Math.Max(cols, ((List<object>)row).Count);
        if (cols < 1) throw new Fail("BAD_ARGS", "table needs at least one column");

        double W = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double H = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        string deckFont = DeckTheme.Load(p).Font;
        double l = a.Dbl("l", W * 0.05), t = a.Dbl("t", H * 0.62), w = a.Dbl("w", W * 0.90), h = a.Dbl("h", Math.Min(H * 0.32, 34 * rows.Count));
        dynamic table = slide.Shapes.AddTable(rows.Count, cols, (float)l, (float)t, (float)w, (float)h).Table;
        for (int ri = 0; ri < rows.Count; ri++)
        {
            List<object> row = (List<object>)rows[ri];
            for (int ci = 0; ci < row.Count && ci < cols; ci++)
            {
                dynamic cell = table.Cell(ri + 1, ci + 1);
                string v = row[ci] == null ? "" : Convert.ToString(row[ci], CultureInfo.InvariantCulture);
                dynamic cr = cell.Shape.TextFrame.TextRange;
                cr.Text = v.Replace("\\n", "\n");
                bool header = ri == 0 && a.Flag("headerRow", true);
                try { cr.Font.Size = (float)a.Dbl("size", header ? 15 : 13); } catch (Exception) { }
                try
                {
                    cr.Font.Name = deckFont;
                    cr.Font.NameFarEast = deckFont;
                }
                catch (Exception) { }
                if (header)
                {
                    try { cr.Font.Bold = -1; } catch (Exception) { }
                    try { cr.Font.Color.RGB = ColorOf("#FFFFFF", 0); } catch (Exception) { }
                }
            }
        }
        string name = Convert.ToString(TryGet(() => (object)table.Parent.Name, ""), CultureInfo.InvariantCulture);
        return new Dictionary<string, object> {
            { "added", name }, { "slide", a.Int("slide", 0) },
            { "rows", rows.Count }, { "cols", cols },
            { "note", "render the slide to check the table fits (too many rows -> shrink -Size or split the table)" }
        };
    }

    private static object ShapeCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic slide = SlideAt(p, a.Int("slide", 0));
        string action = a.Str("action", "add").ToLowerInvariant();
        if (action == "delete")
        {
            string name = a.Need("name");
            int count = 0;
            try { count = (int)slide.Shapes.Count; }
            catch (Exception) { }
            dynamic target = null;
            for (int k = 1; k <= count; k++)
            {
                dynamic cand = slide.Shapes.Item(k);
                if (string.Equals(Convert.ToString(TryGet(() => (object)cand.Name, ""), CultureInfo.InvariantCulture), name, StringComparison.OrdinalIgnoreCase))
                { target = cand; break; }
            }
            if (target == null) throw new Fail("BAD_ARGS", "no shape named \"" + name + "\" on that slide");
            target.Delete();
            return new Dictionary<string, object> { { "deleted", name }, { "slide", a.Int("slide", 0) } };
        }
        if (action != "add") throw new Fail("BAD_ARGS", "shape -Action must be add|delete");

        string type = a.Str("type", "rect").ToLowerInvariant();
        int auto;
        if (type == "rect" || type == "\u77e9\u5f62") auto = 1;                 // msoShapeRectangle
        else if (type == "roundrect" || type == "\u5706\u89d2\u77e9\u5f62") auto = 5;  // msoShapeRoundedRectangle
        else if (type == "ellipse" || type == "\u692d\u5706" || type == "oval") auto = 9; // msoShapeOval
        else if (type == "line" || type == "\u76f4\u7ebf") auto = 9; // handled below
        else throw new Fail("BAD_ARGS", "unknown shape type \"" + type + "\"; use rect|roundrect|ellipse");

        double W = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double H = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        double l = a.Dbl("l", W * 0.35), t = a.Dbl("t", H * 0.4), w = a.Dbl("w", W * 0.3), h = a.Dbl("h", H * 0.2);
        dynamic shp = slide.Shapes.AddShape(auto, (float)l, (float)t, (float)w, (float)h);
        string fill = a.Str("fill", null), line = a.Str("line", null);
        if (fill != null) { try { shp.Fill.ForeColor.RGB = ColorOf(fill, 0xFFFFFF); shp.Fill.Solid(); } catch (Exception) { } }
        if (line != null) { try { shp.Line.ForeColor.RGB = ColorOf(line, 0); } catch (Exception) { } }
        string text = a.Str("text", null);
        if (text != null)
        {
            shp.TextFrame.TextRange.Text = text.Replace("\\n", "\n");
            ApplyFont(shp.TextFrame.TextRange, a);
        }
        string wantName = a.Str("name", null);
        if (wantName != null && wantName.Length > 0) { try { shp.Name = wantName; } catch (Exception) { } }
        string addedName = Convert.ToString(TryGet(() => (object)shp.Name, ""), CultureInfo.InvariantCulture);
        return new Dictionary<string, object> { { "added", addedName }, { "slide", a.Int("slide", 0) }, { "type", type } };
    }

    // ------------------------------------------------------------ format / notes / view

    private static object FormatCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic slide = SlideAt(p, a.Int("slide", 0));
        dynamic sh = ResolveShape(slide, a, "shape", "shapeIdx", true);
        string text = ShapeText(sh);
        if (text.Length == 0) throw new Fail("BAD_ARGS", "that shape has no text; format only works on text shapes (use textbox first)");
        ApplyFont(sh.TextFrame.TextRange, a);
        string fill = a.Str("fill", null);
        if (fill != null) { try { sh.Fill.ForeColor.RGB = ColorOf(fill, 0xFFFFFF); sh.Fill.Solid(); } catch (Exception) { } }
        return new Dictionary<string, object> {
            { "slide", a.Int("slide", 0) },
            { "shape", Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture) },
            { "readBack", new Dictionary<string, object> {
                { "font", Convert.ToString(TryGet(() => (object)sh.TextFrame.TextRange.Font.Name, ""), CultureInfo.InvariantCulture) },
                { "size", TryGet(() => (object)sh.TextFrame.TextRange.Font.Size, null) },
                { "bold", TryGet(() => (object)sh.TextFrame.TextRange.Font.Bold, null) }
            } }
        };
    }

    /// <summary>Unify the font of every run in the deck - groups and table cells included.
    /// Fonts only: colour, size, bold, alignment and geometry are left exactly as they are.</summary>
    private sealed class FontSweep
    {
        public string Font;
        public int Shapes;
        public int Failed;
        public int Groups;
        public int Tables;
        public List<object> Samples = new List<object>();
    }

    private static object FontCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        string font = a.Str("font", null);
        if (font == null || font.Length == 0) throw new Fail("BAD_ARGS", "font needs -Font <name> (e.g. -Font \u5b8b\u4f53)");
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }
        int from = a.Int("from", 1), to = a.Int("to", 0);
        if (from < 1) from = 1;
        if (to <= 0 || to > n) to = n;
        FontSweep sw = new FontSweep();
        sw.Font = font;
        for (int k = from; k <= to; k++)
        {
            dynamic s = null;
            try { s = p.Slides.Item(k); } catch (Exception) { sw.Failed++; continue; }
            int cnt = 0;
            try { cnt = (int)s.Shapes.Count; } catch (Exception) { continue; }
            for (int i = 1; i <= cnt; i++)
            {
                dynamic sh = null;
                try { sh = s.Shapes.Item(i); } catch (Exception) { continue; }
                FontDeep(sh, k, sw);
            }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["font"] = font;
        r["from"] = from;
        r["to"] = to;
        r["shapes"] = sw.Shapes;
        r["groups"] = sw.Groups;
        r["tables"] = sw.Tables;
        r["failed"] = sw.Failed;
        r["samples"] = sw.Samples;
        return r;
    }

    private static void FontDeep(dynamic sh, int slideNo, FontSweep sw)
    {
        int type = 1;
        try { type = Convert.ToInt32(TryGet(() => (object)sh.Type, 1), CultureInfo.InvariantCulture); } catch (Exception) { }
        if (type == 6)
        {
            sw.Groups++;
            int gc = 0;
            try { gc = (int)sh.GroupItems.Count; } catch (Exception) { sw.Failed++; return; }
            for (int j = 1; j <= gc; j++)
            {
                dynamic g = null;
                try { g = sh.GroupItems.Item(j); } catch (Exception) { sw.Failed++; continue; }
                FontDeep(g, slideNo, sw);
            }
            return;
        }
        if (type == 19)
        {
            sw.Tables++;
            int rows = 0, cols = 0;
            try { rows = (int)sh.Table.Rows.Count; cols = (int)sh.Table.Columns.Count; } catch (Exception) { sw.Failed++; return; }
            int cells = 0;
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    try
                    {
                        dynamic cr = sh.Table.Cell(r, c).Shape.TextFrame.TextRange;
                        cr.Font.Name = sw.Font;
                        try { cr.Font.NameFarEast = sw.Font; } catch (Exception) { }
                        cells++;
                    }
                    catch (Exception) { sw.Failed++; }
                }
            }
            if (cells > 0)
            {
                sw.Shapes++;
                if (sw.Samples.Count < 12)
                    sw.Samples.Add(new Dictionary<string, object> {
                        { "slide", slideNo },
                        { "shape", Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture) },
                        { "cells", cells }
                    });
            }
            return;
        }
        string text = ShapeText(sh);
        if (text == null || text.Length == 0) return;
        try
        {
            dynamic tr = sh.TextFrame.TextRange;
            tr.Font.Name = sw.Font;
            try { tr.Font.NameFarEast = sw.Font; } catch (Exception) { }
            sw.Shapes++;
            if (sw.Samples.Count < 12)
                sw.Samples.Add(new Dictionary<string, object> {
                    { "slide", slideNo },
                    { "shape", Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture) },
                    { "chars", text.Length }
                });
        }
        catch (Exception) { sw.Failed++; }
    }

    private static object NotesCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic slide = SlideAt(p, a.Int("slide", 0));
        if (a.Flag("clear", false))
        {
            try { slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = ""; }
            catch (Exception e) { throw new Fail("COM_ERROR", "could not clear notes: " + Clean(e.Message)); }
            return new Dictionary<string, object> { { "cleared", a.Int("slide", 0) } };
        }
        string text = TextFrom(a);
        try { slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text; }
        catch (Exception e) { throw new Fail("COM_ERROR", "could not write notes: " + Clean(e.Message)); }
        return new Dictionary<string, object> { { "slide", a.Int("slide", 0) }, { "notesAfter", Clip(text, 90) } };
    }

    private static object ViewCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic app = App(false, false);
        int slide = a.Int("slide", 0);
        dynamic win = TryGet(() => (object)app.ActiveWindow, null) as object;
        if (win == null) throw new Fail("BAD_ARGS", "no active window; open the deck with -Show first");
        int n = 0;
        try { n = (int)p.Slides.Count; }
        catch (Exception) { }
        if (slide < 1 || slide > n) throw new Fail("BAD_ARGS", "slide out of range");
        try { win.GotoSlide(slide); }
        catch (Exception) { try { win.View.GotoSlide(slide); } catch (Exception e) { throw new Fail("COM_ERROR", "GotoSlide failed: " + Clean(e.Message)); } }
        return new Dictionary<string, object> { { "showing", slide } };
    }

    // ------------------------------------------------------------ render (clipboard EMF -> PNG)

    private static object RenderPres(Bag a)
    {
        dynamic p = RequirePres(a, true);
        string outDir = a.Need("out");
        Directory.CreateDirectory(outDir);
        int n = 0;
        try { n = (int)p.Slides.Count; }
        catch (Exception) { }
        int from = Math.Max(1, a.Int("from", 1));
        int to = a.Int("to", 0);
        if (to <= 0 || to > n) to = n;
        if (from > to) throw new Fail("BAD_ARGS", "-From " + from + " > last slide " + to);
        bool overview = a.Flag("overview", false);
        int width = a.Int("width", 1100);
        if (overview) return RenderOverview(p, outDir, from, to, width);
        if (to - from + 1 > 4) throw new Fail("BAD_ARGS", "render at most 4 slides per call (got " + (to - from + 1) + "); use -Overview to see many slides at once, or page through with -From/-To");

        List<object> files = new List<object>();
        // v1.32: Slide.Copy() overwrites the machine-wide clipboard, so the user's content is saved
        // once for the whole call and put back at the end (helper/shared/ClipboardKeeper.cs).
        using (ClipboardKeeper.Keep())
        {
            for (int k = from; k <= to; k++)
            {
                dynamic slide = p.Slides.Item(k);
                Bitmap bmp;
                {
                    IDisposable clip = WocLocks.Clipboard(20000);
                    if (clip == null) throw new Fail("BUSY", "the system clipboard is busy (another render is running); retry in a moment");
                    try
                    {
                        slide.Copy();
                        ClipboardKeeper.Sync();
                        Thread.Sleep(250);
                        bmp = ClipboardBitmap(width);
                    }
                    finally { clip.Dispose(); }
                }
                if (bmp == null)
                    throw new Fail("COM_ERROR", "could not get slide " + k + " as an image from the clipboard (is WPS showing a dialog?)");
                string png = Path.Combine(outDir, "slide-" + k.ToString(CultureInfo.InvariantCulture) + ".png");
                bmp.Save(png, ImageFormat.Png);
                bmp.Dispose();
                files.Add(png);
            }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slides"] = files;
        r["note"] = "PNG files are written by the helper process (plaintext); verify with the rendered image, not by trusting the command";
        return r;
    }

    /// <summary>Render up to 24 slides as 2x3 contact sheets with page numbers.</summary>
    private static object RenderOverview(dynamic p, string outDir, int from, int to, int width)
    {
        if (to - from + 1 > 24) throw new Fail("BAD_ARGS", "overview covers at most 24 slides per call; page through with -From/-To");
        int cellW = Math.Max(320, Math.Min(720, width / 2));
        List<object> sheets = new List<object>();
        const int perSheet = 6;
        // v1.32: same clipboard etiquette as render - save the user's content once, put it back at the end.
        using (ClipboardKeeper.Keep())
        {
        for (int start = from; start <= to; start += perSheet)
        {
            int end = Math.Min(to, start + perSheet - 1);
            List<int> indexes = new List<int>();
            List<Bitmap> cells = new List<Bitmap>();
            double ratio = 0.5625;
            for (int k = start; k <= end; k++)
            {
                dynamic slide = p.Slides.Item(k);
            Bitmap bmp;
            {
                IDisposable clip = WocLocks.Clipboard(20000);
                if (clip == null) throw new Fail("BUSY", "the system clipboard is busy (another render is running); retry in a moment");
                try
                {
                    slide.Copy();
                    ClipboardKeeper.Sync();
                    Thread.Sleep(220);
                    bmp = ClipboardBitmap(cellW);
                }
                finally { clip.Dispose(); }
            }
                if (bmp == null) throw new Fail("COM_ERROR", "could not get slide " + k + " as an image from the clipboard");
                cells.Add(bmp);
                indexes.Add(k);
                ratio = (double)bmp.Width / bmp.Height;
            }
            int cellH = (int)Math.Round(cellW / ratio);
            int pad = 18, label = 26, cols = 2;
            int rows = (cells.Count + cols - 1) / cols;
            int sheetW = cols * cellW + (cols + 1) * pad;
            int sheetH = rows * (cellH + label) + (rows + 1) * pad;
            string path = Path.Combine(outDir, "overview-" + start.ToString(CultureInfo.InvariantCulture) + "-" + end.ToString(CultureInfo.InvariantCulture) + ".png");
            using (Bitmap sheet = new Bitmap(sheetW, sheetH))
            using (Graphics g = Graphics.FromImage(sheet))
            using (Font f = new Font("\u5fae\u8f6f\u96c5\u9ed1", 12f))
            using (SolidBrush txt = new SolidBrush(Color.FromArgb(60, 60, 60)))
            using (Pen border = new Pen(Color.FromArgb(205, 205, 205)))
            {
                g.Clear(Color.FromArgb(246, 246, 246));
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                for (int k = 0; k < cells.Count; k++)
                {
                    int col = k % cols, row = k / cols;
                    int x = pad + col * (cellW + pad);
                    int y = pad + row * (cellH + label + pad);
                    g.DrawString("\u7b2c " + indexes[k] + " \u9875", f, txt, x, y);
                    g.DrawImage(cells[k], x, y + label, cellW, cellH);
                    g.DrawRectangle(border, x, y + label, cellW, cellH);
                }
                sheet.Save(path, ImageFormat.Png);
            }
            foreach (Bitmap b in cells) b.Dispose();
            sheets.Add(path);
        }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["overview"] = sheets;
        r["slides"] = to - from + 1;
        r["note"] = "one image per 6 slides with page numbers; check overflow, misalignment and empty areas before saving";
        return r;
    }

    /// <summary>Bitmap of the slide currently on the clipboard (EMF preferred).</summary>
    private static Bitmap ClipboardBitmap(int targetWidth)
    {
        for (int attempt = 0; attempt < 6; attempt++)
        {
            Bitmap result = null;
            // EMF via raw Win32: Clipboard.GetData("EnhancedMetafile") crashes some hosts
            // (access violation observed on this machine family when WPS holds the source);
            // GetClipboardData + CopyEnhMetaFile is the safe route.
            try
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    try
                    {
                        if (IsClipboardFormatAvailable(14)) // CF_ENHMETAFILE
                        {
                            IntPtr src = GetClipboardData(14);
                            if (src != IntPtr.Zero)
                            {
                                IntPtr own = CopyEnhMetaFile(src, IntPtr.Zero);
                                if (own != IntPtr.Zero)
                                {
                                    using (Metafile emf = new Metafile(own, false))
                                    {
                                        double ratio = emf.Height <= 0 ? 0.5625 : (double)emf.Width / emf.Height;
                                        int w = Math.Max(320, Math.Min(2400, targetWidth));
                                        int h = Math.Max(1, (int)Math.Round(w / ratio));
                                        Bitmap bmp = new Bitmap(w, h);
                                        using (Graphics g = Graphics.FromImage(bmp))
                                        {
                                            g.Clear(Color.White);
                                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            g.DrawImage(emf, 0, 0, w, h);
                                        }
                                        result = bmp;
                                    }
                                    DeleteEnhMetaFile(own);
                                }
                            }
                        }
                    }
                    finally { CloseClipboard(); }
                }
            }
            catch (Exception) { if (result != null) { result.Dispose(); result = null; } }
            if (result == null)
            {
                try
                {
                    if (Clipboard.ContainsImage())
                    {
                        Image raw = Clipboard.GetImage();
                        if (raw != null)
                        {
                            int w = Math.Max(320, Math.Min(2400, targetWidth));
                            int h = Math.Max(1, (int)Math.Round((double)w * raw.Height / raw.Width));
                            Bitmap bmp = new Bitmap(w, h);
                            using (Graphics g = Graphics.FromImage(bmp))
                            {
                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                g.DrawImage(raw, 0, 0, w, h);
                            }
                            raw.Dispose();
                            result = bmp;
                        }
                    }
                }
                catch (Exception) { }
            }
            if (result != null) return result;
            Thread.Sleep(200);
        }
        return null;
    }
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr hWndNewOwner);
    [DllImport("user32.dll")]
    private static extern bool CloseClipboard();
    [DllImport("user32.dll")]
    private static extern bool IsClipboardFormatAvailable(int format);
    [DllImport("user32.dll")]
    private static extern IntPtr GetClipboardData(int format);
    [DllImport("gdi32.dll")]
    private static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, IntPtr lpszFile);
    [DllImport("gdi32.dll")]
    private static extern bool DeleteEnhMetaFile(IntPtr hemf);

    // ------------------------------------------------------------ theme command

    private static object ThemeCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        string path = ThemePath(p);
        bool reset = a.Flag("reset", false);
        string setPath = a.Str("set", null);

        if (reset)
        {
            if (File.Exists(path)) File.Delete(path);
            try
            {
                dynamic tags = p.Tags;
                if (tags != null) tags.Delete(ThemeTagName);
            }
            catch (Exception) { }
            return new Dictionary<string, object> { { "reset", true }, { "theme", DeckTheme.Load(p).ToBag() } };
        }
        if (setPath != null)
        {
            if (!File.Exists(setPath)) throw new Fail("BAD_ARGS", "theme file not found: " + setPath);
            Bag b = new Bag(Json.Parse(File.ReadAllText(setPath, Encoding.UTF8)));
            DeckTheme t = DeckTheme.FromBag(b, new DeckTheme());
            File.WriteAllText(path, Json.Write(t.ToBag()), new UTF8Encoding(false));
            Dictionary<string, object> r = new Dictionary<string, object>();
            r["set"] = path;
            r["theme"] = t.ToBag();
            // v1.32: keep a copy inside the deck as well, so the palette survives a machine move
            r["tags"] = ThemeTagWrite(p, Json.Write(t.ToBag()), false);
            if (a.Flag("apply", false)) r["applied"] = ApplyThemeAll(p, t);
            return r;
        }
        if (a.Flag("apply", false))
        {
            string src;
            DeckTheme t = DeckTheme.LoadWithSource(p, out src);
            Dictionary<string, object> r = new Dictionary<string, object>();
            r["theme"] = t.ToBag();
            r["themeSource"] = src;
            r["applied"] = ApplyThemeAll(p, t);
            return r;
        }
        string source;
        DeckTheme cur = DeckTheme.LoadWithSource(p, out source);
        return new Dictionary<string, object> {
            { "path", path },
            { "theme", cur.ToBag() },
            { "themeSource", source },
            { "tags", a.Flag("probe", false) ? ThemeTagWrite(p, "", true) : (object)new Dictionary<string, object> { { "note", "run `theme -Probe` to test whether this deck accepts Presentation.Tags" } } },
            { "note", "theme is per deck; set it with -Set <json>, restyle existing slides with -Apply. Defaults: \u6c34\u5229\u6df1\u84dd #1F4E79 + \u91d1 #C99700 + \u5fae\u8f6f\u96c5\u9ed1" }
        };
    }

    private static object ApplyThemeAll(dynamic p, DeckTheme th)
    {
        int n = 0;
        try { n = (int)p.Slides.Count; }
        catch (Exception) { }
        int styled = 0, footers = 0;
        for (int k = 1; k <= n; k++)
        {
            dynamic s = p.Slides.Item(k);
            styled += ApplyThemeToSlide(s, th, k);
            footers += AddFooter(s, th, k, k == 1);
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slides"] = n;
        r["restyledShapes"] = styled;
        r["footers"] = footers;
        return r;
    }

    /// <summary>Dark colour blocks on a slide (the presets' bands). Text sitting on one of
    /// them must be white, otherwise a theme pass turns an end-slide title invisible.</summary>
    private static List<double[]> DarkRects(dynamic slide)
    {
        List<double[]> list = new List<double[]>();
        int count = 0;
        try { count = (int)slide.Shapes.Count; }
        catch (Exception) { }
        for (int k = 1; k <= count; k++)
        {
            dynamic sh = null;
            try { sh = slide.Shapes.Item(k); }
            catch (Exception) { continue; }
            string name = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
            if (!IsDecoration(name) || name.ToLowerInvariant() == "accentline" || name.ToLowerInvariant() == "divider") continue;
            double l = Convert.ToDouble(TryGet(() => (object)sh.Left, (object)0.0), CultureInfo.InvariantCulture);
            double t = Convert.ToDouble(TryGet(() => (object)sh.Top, (object)0.0), CultureInfo.InvariantCulture);
            double w = Convert.ToDouble(TryGet(() => (object)sh.Width, (object)0.0), CultureInfo.InvariantCulture);
            double h = Convert.ToDouble(TryGet(() => (object)sh.Height, (object)0.0), CultureInfo.InvariantCulture);
            if (w > 0 && h > 0) list.Add(new double[] { l, t, w, h });
        }
        return list;
    }

    private static bool OverDark(List<double[]> rects, double cx, double cy)
    {
        foreach (double[] r in rects)
            if (cx >= r[0] && cx <= r[0] + r[2] && cy >= r[1] && cy <= r[1] + r[3]) return true;
        return false;
    }

    /// <summary>Re-colour and re-font the text shapes of one slide according to the theme.</summary>
    private static int ApplyThemeToSlide(dynamic slide, DeckTheme th, int index)
    {
        int count = 0, done = 0;
        try { count = (int)slide.Shapes.Count; }
        catch (Exception) { }
        List<double[]> darks = DarkRects(slide);
        for (int k = 1; k <= count; k++)
        {
            dynamic sh = null;
            try { sh = slide.Shapes.Item(k); }
            catch (Exception) { continue; }
            string text = ShapeText(sh);
            if (text.Length == 0) continue;
            string name = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
            if (IsDecoration(name)) continue;
            bool isTitle = name == "title" || name == "sectionTitle" || name == "number";
            bool isSub = name == "subtitle" || name == "footer" || name == "pageNo";
            if (!isTitle && !isSub)
            {
                double sz = Convert.ToDouble(TryGet(() => (object)sh.TextFrame.TextRange.Font.Size, (object)0.0), CultureInfo.InvariantCulture);
                isTitle = sz >= 26;
            }
            double l = Convert.ToDouble(TryGet(() => (object)sh.Left, (object)0.0), CultureInfo.InvariantCulture);
            double t = Convert.ToDouble(TryGet(() => (object)sh.Top, (object)0.0), CultureInfo.InvariantCulture);
            double w = Convert.ToDouble(TryGet(() => (object)sh.Width, (object)0.0), CultureInfo.InvariantCulture);
            double h = Convert.ToDouble(TryGet(() => (object)sh.Height, (object)0.0), CultureInfo.InvariantCulture);
            bool dark = OverDark(darks, l + w / 2, t + h / 2);
            try
            {
                dynamic tr = sh.TextFrame.TextRange;
                string font = isTitle ? th.TitleFont : th.Font;
                tr.Font.Name = font;
                tr.Font.NameFarEast = font;
                if (dark) tr.Font.Color.RGB = th.White;
                else if (isSub) tr.Font.Color.RGB = th.Muted;
                else if (isTitle) tr.Font.Color.RGB = th.Title;
                else tr.Font.Color.RGB = th.Text;
                done++;
            }
            catch (Exception) { }
        }
        return done;
    }

    /// <summary>Bottom footer line + page number (skipped on the cover when the theme says so).</summary>
    private static int AddFooter(dynamic slide, DeckTheme th, int index, bool isCover)
    {
        if (isCover && th.SkipFooterOnCover) return 0;
        if (th.Footer == null || (th.Footer.Length == 0 && !th.PageNumber)) return 0;
        double W = 960, H = 540;
        // page size is not on the slide object; derive from an existing shape or use defaults
        try
        {
            dynamic shp = slide.Shapes.Item(1);
            W = Math.Max(Convert.ToDouble(shp.Left, CultureInfo.InvariantCulture) * 2 + Convert.ToDouble(shp.Width, CultureInfo.InvariantCulture), 640);
        }
        catch (Exception) { }
        int added = 0;
        List<double[]> darks = DarkRects(slide);
        bool footerDark = OverDark(darks, 40 + W * 0.3, H - 23);
        bool pageDark = OverDark(darks, W - 60, H - 23);
        int footerColor = footerDark ? th.White : th.Muted;
        int pageColor = pageDark ? th.White : th.Muted;
        // replace an old footer/page number so repeated -Apply does not stack them
        int count = 0;
        try { count = (int)slide.Shapes.Count; }
        catch (Exception) { }
        for (int k = count; k >= 1; k--)
        {
            dynamic sh = null;
            try { sh = slide.Shapes.Item(k); }
            catch (Exception) { continue; }
            string name = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
            if (name == "footer" || name == "pageNo") { try { sh.Delete(); } catch (Exception) { } }
        }
        if (!string.IsNullOrEmpty(th.Footer))
        {
            try
            {
                dynamic tb = slide.Shapes.AddTextbox(1, 40f, (float)(H - 34), (float)(W * 0.6), 22f);
                tb.Name = "footer";
                dynamic tr = tb.TextFrame.TextRange;
                tr.Text = th.Footer;
                tr.Font.Size = 10f;
                tr.Font.Color.RGB = footerColor;
                tr.Font.Name = th.Font;
                tr.Font.NameFarEast = th.Font;
                added++;
            }
            catch (Exception) { }
        }
        if (th.PageNumber)
        {
            try
            {
                dynamic tb = slide.Shapes.AddTextbox(1, (float)(W - 90), (float)(H - 34), 50f, 22f);
                tb.Name = "pageNo";
                dynamic tr = tb.TextFrame.TextRange;
                tr.Text = index.ToString(CultureInfo.InvariantCulture);
                tr.Font.Size = 10f;
                tr.Font.Color.RGB = pageColor;
                tr.Font.Name = th.Font;
                tr.Font.NameFarEast = th.Font;
                tr.ParagraphFormat.Alignment = 3; // right
                added++;
            }
            catch (Exception) { }
        }
        return added;
    }

    // ------------------------------------------------------------ audit (mechanical pre-check)

    private static object AuditCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        double W = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double H = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        int minFont = a.Int("minFont", 12);
        int n = 0;
        try { n = (int)p.Slides.Count; }
        catch (Exception) { }
        int from = Math.Max(1, a.Int("from", 1));
        int to = a.Int("to", 0);
        if (to <= 0 || to > n) to = n;

        List<object> issues = new List<object>();
        for (int k = from; k <= to; k++)
        {
            dynamic s = p.Slides.Item(k);
            int count = 0;
            try { count = (int)s.Shapes.Count; }
            catch (Exception) { }
            List<double[]> boxes = new List<double[]>();
            List<string> boxNames = new List<string>();
            bool hasText = false, hasPicture = false, hasTitle = false;
            for (int j = 1; j <= count; j++)
            {
                dynamic sh = null;
                try { sh = s.Shapes.Item(j); }
                catch (Exception) { continue; }
                string name = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                if (IsTakeoverMark(name)) continue;   // the badge is ours, not a layout problem
                int type = (int)TryGet(() => (object)(int)sh.Type, (object)0);
                double l = Convert.ToDouble(TryGet(() => (object)sh.Left, (object)0.0), CultureInfo.InvariantCulture);
                double t = Convert.ToDouble(TryGet(() => (object)sh.Top, (object)0.0), CultureInfo.InvariantCulture);
                double w = Convert.ToDouble(TryGet(() => (object)sh.Width, (object)0.0), CultureInfo.InvariantCulture);
                double h = Convert.ToDouble(TryGet(() => (object)sh.Height, (object)0.0), CultureInfo.InvariantCulture);
                string text = ShapeText(sh);
                if (text.Trim().Length > 0) hasText = true;
                if (type == 13 || type == 11) hasPicture = true;
                if (name == "title" || name == "sectionTitle") hasTitle = true;
                try
                {
                    if ((int)sh.IsPlaceholder != 0 && (int)sh.PlaceholderFormat.Type == 1) hasTitle = true;
                }
                catch (Exception) { }

                // 1) outside the slide frame
                if (l < -2 || t < -2 || (w > 0 && l + w > W + 2) || (h > 0 && t + h > H + 2))
                    issues.Add(Issue(k, name, "outOfFrame", "shape sits outside the slide (" + R(l) + "," + R(t) + " " + R(w) + "x" + R(h) + ")"));

                if (text.Trim().Length > 0)
                {
                    // 2) text taller than its box (the classic overflow)
                    double boundH = Convert.ToDouble(TryGet(() => (object)sh.TextFrame.TextRange.BoundHeight, (object)0.0), CultureInfo.InvariantCulture);
                    if (boundH > 0 && h > 0 && boundH > h + 4)
                        issues.Add(Issue(k, name, "textOverflow", "text needs " + R(boundH) + "pt but the box is " + R(h) + "pt tall"));
                    // 3) too small to read when projected (footer / page number are legitimately small)
                    bool isFurniture = name == "footer" || name == "pageNo";
                    double sz = Convert.ToDouble(TryGet(() => (object)sh.TextFrame.TextRange.Font.Size, (object)0.0), CultureInfo.InvariantCulture);
                    if (!isFurniture && sz >= 1 && sz < minFont)
                        issues.Add(Issue(k, name, "fontTooSmall", R(sz) + "pt < " + minFont + "pt"));
                }

                // 4) picture too small to show its lettering
                if ((type == 13 || type == 11) && w > 0 && w < W * 0.2)
                    issues.Add(Issue(k, name, "imageTooSmall", "picture is only " + R(w / W * 100) + "% of slide width"));

                if (w > 0 && h > 0) { boxes.Add(new double[] { l, t, w, h }); boxNames.Add(name); }
            }

            // 5) pairwise overlap (a picture under a text box is almost always an accident;
            //    colour bands and rules are meant to sit under text, so they are skipped)
            for (int i = 0; i < boxes.Count; i++)
                for (int m = i + 1; m < boxes.Count; m++)
                {
                    if (IsDecoration(boxNames[i]) || IsDecoration(boxNames[m])) continue;
                    double ox = Math.Min(boxes[i][0] + boxes[i][2], boxes[m][0] + boxes[m][2]) - Math.Max(boxes[i][0], boxes[m][0]);
                    double oy = Math.Min(boxes[i][1] + boxes[i][3], boxes[m][1] + boxes[m][3]) - Math.Max(boxes[i][1], boxes[m][1]);
                    if (ox <= 0 || oy <= 0) continue;
                    double inter = ox * oy;
                    double small = Math.Min(boxes[i][2] * boxes[i][3], boxes[m][2] * boxes[m][3]);
                    if (small > 0 && inter / small > 0.30)
                        issues.Add(Issue(k, boxNames[i] + " x " + boxNames[m], "overlap", R(inter / small * 100) + "% of the smaller shape is covered"));
                }

            // 6) empty / untitled slide
            if (!hasText && !hasPicture) issues.Add(Issue(k, "-", "emptySlide", "no text and no picture"));
            else if (hasText && !hasTitle) issues.Add(Issue(k, "-", "noTitle", "no title shape (name title/sectionTitle or a title placeholder)"));
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slides"] = n;
        r["checked"] = new Dictionary<string, object> { { "from", from }, { "to", to } };
        r["issues"] = issues;
        r["count"] = issues.Count;
        r["note"] = "mechanical pre-check only: it cannot judge aesthetics. Render the flagged slides (and -Overview for the whole deck) before telling the user it is ready.";
        return r;
    }

    private static Dictionary<string, object> Issue(int slide, string shape, string kind, string detail)
    {
        return new Dictionary<string, object> {
            { "slide", slide }, { "shape", shape }, { "kind", kind }, { "detail", detail }
        };
    }

    private static double R(double v) { return Math.Round(v, 1); }

    /// <summary>Colour bands / rules created by the presets: they are supposed to be covered.</summary>
    private static bool IsDecoration(string name)
    {
        if (name == null) return false;
        if (IsTakeoverMark(name)) return true;
        string n = name.ToLowerInvariant();
        return n == "coverband" || n == "sectionband" || n == "endbackground" || n == "accentline" || n == "divider";
    }

    // ------------------------------------------------------------ align / z-order

    private static object AlignCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic slide = SlideAt(p, a.Int("slide", 0));
        string mode = a.Need("mode").ToLowerInvariant();
        List<dynamic> shapes = ResolveShapes(slide, a.Str("shapes", null));
        if (shapes.Count == 0) throw new Fail("BAD_ARGS", "no shapes matched; read the slide to get shape names");
        double W = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);

        if (mode == "left") { double v = Min(shapes, "Left"); foreach (dynamic s in shapes) s.Left = (float)v; }
        else if (mode == "right") { double v = Max(shapes, "Left", "Width"); foreach (dynamic s in shapes) s.Left = (float)(v - Convert.ToDouble(s.Width, CultureInfo.InvariantCulture)); }
        else if (mode == "top") { double v = Min(shapes, "Top"); foreach (dynamic s in shapes) s.Top = (float)v; }
        else if (mode == "bottom") { double v = Max(shapes, "Top", "Height"); foreach (dynamic s in shapes) s.Top = (float)(v - Convert.ToDouble(s.Height, CultureInfo.InvariantCulture)); }
        else if (mode == "center-h") { foreach (dynamic s in shapes) s.Left = (float)((W - Convert.ToDouble(s.Width, CultureInfo.InvariantCulture)) / 2); }
        else if (mode == "same-width") { double v = Convert.ToDouble(shapes[0].Width, CultureInfo.InvariantCulture); foreach (dynamic s in shapes) s.Width = (float)v; }
        else if (mode == "same-height") { double v = Convert.ToDouble(shapes[0].Height, CultureInfo.InvariantCulture); foreach (dynamic s in shapes) s.Height = (float)v; }
        else if (mode == "distribute-h" || mode == "distribute-v")
        {
            bool horiz = mode == "distribute-h";
            shapes.Sort((x, y) => Convert.ToDouble(horiz ? x.Left : x.Top, CultureInfo.InvariantCulture)
                .CompareTo(Convert.ToDouble(horiz ? y.Left : y.Top, CultureInfo.InvariantCulture)));
            if (shapes.Count < 3) throw new Fail("BAD_ARGS", "distribute needs at least 3 shapes");
            double first = Convert.ToDouble(horiz ? shapes[0].Left : shapes[0].Top, CultureInfo.InvariantCulture);
            double last = Convert.ToDouble(horiz ? shapes[shapes.Count - 1].Left : shapes[shapes.Count - 1].Top, CultureInfo.InvariantCulture);
            double total = horiz ? Convert.ToDouble(shapes[shapes.Count - 1].Width, CultureInfo.InvariantCulture)
                                  : Convert.ToDouble(shapes[shapes.Count - 1].Height, CultureInfo.InvariantCulture);
            double span = last + total - first;
            double sum = 0;
            foreach (dynamic s in shapes) sum += Convert.ToDouble(horiz ? s.Width : s.Height, CultureInfo.InvariantCulture);
            double gap = (span - sum) / (shapes.Count - 1);
            double cursor = first;
            foreach (dynamic s in shapes)
            {
                if (horiz) { s.Left = (float)cursor; cursor += Convert.ToDouble(s.Width, CultureInfo.InvariantCulture) + gap; }
                else { s.Top = (float)cursor; cursor += Convert.ToDouble(s.Height, CultureInfo.InvariantCulture) + gap; }
            }
        }
        else throw new Fail("BAD_ARGS", "unknown mode \"" + mode + "\"; use left|right|top|bottom|center-h|same-width|same-height|distribute-h|distribute-v");

        List<object> done = new List<object>();
        foreach (dynamic s in shapes)
            done.Add(new Dictionary<string, object> {
                { "name", Convert.ToString(TryGet(() => (object)s.Name, ""), CultureInfo.InvariantCulture) },
                { "l", R(Convert.ToDouble(TryGet(() => (object)s.Left, (object)0.0), CultureInfo.InvariantCulture)) },
                { "t", R(Convert.ToDouble(TryGet(() => (object)s.Top, (object)0.0), CultureInfo.InvariantCulture)) },
                { "w", R(Convert.ToDouble(TryGet(() => (object)s.Width, (object)0.0), CultureInfo.InvariantCulture)) },
                { "h", R(Convert.ToDouble(TryGet(() => (object)s.Height, (object)0.0), CultureInfo.InvariantCulture)) }
            });
        return new Dictionary<string, object> { { "mode", mode }, { "shapes", done }, { "note", "geometry is reported back; render the slide if it matters" } };
    }

    private static double Min(List<dynamic> shapes, string prop)
    {
        double v = double.MaxValue;
        foreach (dynamic s in shapes)
        {
            double cur = Convert.ToDouble(prop == "Left" ? s.Left : s.Top, CultureInfo.InvariantCulture);
            v = Math.Min(v, cur);
        }
        return v == double.MaxValue ? 0 : v;
    }

    private static double Max(List<dynamic> shapes, string prop, string sizeProp)
    {
        double v = double.MinValue;
        foreach (dynamic s in shapes)
        {
            double start = Convert.ToDouble(sizeProp == "Width" ? s.Left : s.Top, CultureInfo.InvariantCulture);
            double size = Convert.ToDouble(sizeProp == "Width" ? s.Width : s.Height, CultureInfo.InvariantCulture);
            v = Math.Max(v, start + size);
        }
        return v == double.MinValue ? 0 : v;
    }

    private static List<dynamic> ResolveShapes(dynamic slide, string list)
    {
        List<dynamic> result = new List<dynamic>();
        int count = 0;
        try { count = (int)slide.Shapes.Count; }
        catch (Exception) { }
        if (list == null || list.Trim().Length == 0 || list.Trim() == "*" || list.Trim().ToLowerInvariant() == "all")
        {
            for (int k = 1; k <= count; k++)
            {
                try
                {
                    dynamic sh0 = slide.Shapes.Item(k);
                    string nm0 = Convert.ToString(TryGet(() => (object)sh0.Name, ""), CultureInfo.InvariantCulture);
                    if (IsTakeoverMark(nm0)) continue;   // "all shapes" never means our own badge
                    result.Add(sh0);
                }
                catch (Exception) { }
            }
            return result;
        }
        foreach (string raw in list.Split(','))
        {
            string want = raw.Trim();
            if (want.Length == 0) continue;
            int idx;
            if (int.TryParse(want, out idx))
            {
                if (idx >= 1 && idx <= count) result.Add(slide.Shapes.Item(idx));
                continue;
            }
            for (int k = 1; k <= count; k++)
            {
                dynamic sh = slide.Shapes.Item(k);
                string name = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                if (string.Equals(name, want, StringComparison.OrdinalIgnoreCase)) { result.Add(sh); break; }
            }
        }
        return result;
    }

    private static object ZorderCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic slide = SlideAt(p, a.Int("slide", 0));
        dynamic sh = ResolveShape(slide, a, "name", "shapeIdx", true);
        string where = a.Str("where", "front").ToLowerInvariant();
        if (where == "front") sh.ZOrder(0);        // msoBringToFront
        else if (where == "back") sh.ZOrder(1);    // msoSendToBack
        else throw new Fail("BAD_ARGS", "-Where must be front|back");
        return new Dictionary<string, object> {
            { "shape", Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture) },
            { "where", where },
            { "index", TryGet(() => (object)(int)sh.ZOrderPosition, null) }
        };
    }

    // ------------------------------------------------------------ restore (text level)

    private static object RestoreCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        string since = a.Str("since", "last");
        string path;
        if (File.Exists(since)) path = since;
        else
        {
            path = SnapFile(p, since);
        }
        if (!File.Exists(path)) throw new Fail("BAD_ARGS", "snapshot not found: " + path + ". Take snapshot first (or pass -Since <file>).");
        Bag snap = new Bag(Json.Parse(File.ReadAllText(path, Encoding.UTF8)));
        // `replace` writes a richer snapshot (groups, table cells, notes, header/footer);
        // it restores through the same slot walker that produced it.
        if (snap.Has("cells")) return RestoreSlots(p, snap, a.Str("shapes", null), path);
        List<object> before = snap.List("slides");
        string only = a.Str("shapes", null);
        HashSet<string> wanted = null;
        if (only != null)
        {
            wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string raw in only.Split(',')) if (raw.Trim().Length > 0) wanted.Add(raw.Trim());
        }

        int n = 0;
        try { n = (int)p.Slides.Count; }
        catch (Exception) { }
        int restored = 0, skipped = 0;
        List<object> details = new List<object>();
        for (int k = 0; k < Math.Min(before.Count, n); k++)
        {
            Dictionary<string, object> snapshotTexts = ShapeTextMap((List<object>)before[k]);
            dynamic s = p.Slides.Item(k + 1);
            int count = 0;
            try { count = (int)s.Shapes.Count; }
            catch (Exception) { }
            for (int j = 1; j <= count; j++)
            {
                dynamic sh = null;
                try { sh = s.Shapes.Item(j); }
                catch (Exception) { continue; }
                string name = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                string key = name;
                if (!snapshotTexts.ContainsKey(key)) continue;
                if (wanted != null && !wanted.Contains(key) && !wanted.Contains((k + 1) + "!" + key)) continue;
                string want = Convert.ToString(snapshotTexts[key], CultureInfo.InvariantCulture);
                string now = ShapeText(sh);
                if (string.Equals(want, now, StringComparison.Ordinal)) continue;
                try
                {
                    sh.TextFrame.TextRange.Text = want;
                    restored++;
                    details.Add(new Dictionary<string, object> { { "slide", k + 1 }, { "shape", name }, { "now", Clip(now, 50) }, { "restored", Clip(want, 50) } });
                }
                catch (Exception) { skipped++; }
            }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["since"] = path;
        r["restored"] = restored;
        r["failed"] = skipped;
        if (details.Count > 0) r["details"] = details;
        r["note"] = "texts only: layout, sizes and shapes added after the snapshot are not reverted; run diff to confirm, then render before saving";
        return r;
    }

    // ------------------------------------------------------------ snapshot / diff

    private static string SnapshotDir()
    {
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate", "ppt-snapshots");
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static List<object> DeckState(dynamic p)
    {
        List<object> slides = new List<object>();
        int n = 0;
        try { n = (int)p.Slides.Count; }
        catch (Exception) { }
        for (int k = 1; k <= n; k++)
        {
            dynamic s = p.Slides.Item(k);
            List<object> shapes = new List<object>();
            int sc = 0;
            try { sc = (int)s.Shapes.Count; }
            catch (Exception) { }
            for (int j = 1; j <= sc; j++)
            {
                dynamic sh = null;
                try { sh = s.Shapes.Item(j); }
                catch (Exception) { continue; }
                string shName = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                if (IsTakeoverMark(shName)) continue;   // never part of the deck baseline
                shapes.Add(new Dictionary<string, object> {
                    { "name", shName },
                    { "text", ShapeText(sh) }
                });
            }
            slides.Add(shapes);
        }
        return slides;
    }

    // Snapshots live under ppt-snapshots\<hash of the deck path>\ so two decks with the same
    // file name in different folders never share a baseline (multi-task sessions hit this).
    private static string SnapDirOf(dynamic p)
    {
        string full = Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture);
        string dir = Path.Combine(SnapshotDir(), DocKey(full));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static string SnapFile(dynamic p, string tag)
    {
        string safe = new string((tag ?? "last").Where(c => char.IsLetterOrDigit(c) || c == '-' || c == '_').ToArray());
        if (safe.Length == 0) safe = "last";
        return Path.Combine(SnapDirOf(p), safe + ".json");
    }

    private static string DocKey(string fullPath)
    {
        try
        {
            using (System.Security.Cryptography.SHA256 sha = System.Security.Cryptography.SHA256.Create())
            {
                byte[] b = sha.ComputeHash(Encoding.UTF8.GetBytes((fullPath ?? "").ToLowerInvariant()));
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 8; i++) sb.Append(b[i].ToString("x2", CultureInfo.InvariantCulture));
                return sb.ToString();
            }
        }
        catch (Exception) { return "nokey"; }
    }

    private static object SnapshotPres(Bag a)
    {
        dynamic p = RequirePres(a, true);
        string tag = a.Str("tag", "last");
        string safe = new string(tag.Where(c => char.IsLetterOrDigit(c) || c == '-' || c == '_').ToArray());
        if (safe.Length == 0) safe = "last";
        List<object> state = DeckState(p);
        string path = SnapFile(p, safe);
        Dictionary<string, object> snap = new Dictionary<string, object>();
        snap["at"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        snap["slides"] = state;
        File.WriteAllText(path, Json.Write(snap), new UTF8Encoding(false));

        int textShapes = 0;
        foreach (object sl in state) foreach (object sh in (List<object>)sl) if (((string)((Dictionary<string, object>)sh)["text"]).Length > 0) textShapes++;

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["path"] = path;
        r["tag"] = tag;
        r["slides"] = state.Count;
        r["textShapes"] = textShapes;
        r["note"] = "text-level snapshot (shape names + texts); deck has no revision layer and COM edits bypass undo, so snapshot before big batches";
        return r;
    }

    private static object DiffPres(Bag a)
    {
        dynamic p = RequirePres(a, true);
        string since = a.Str("since", "last");
        string path;
        if (File.Exists(since)) path = since;
        else
        {
            path = SnapFile(p, since);
        }
        if (!File.Exists(path)) throw new Fail("BAD_ARGS", "snapshot not found: " + path + ". Take snapshot first (or pass -Since <file>).");
        Bag snap = new Bag(Json.Parse(File.ReadAllText(path, Encoding.UTF8)));
        List<object> before = snap.List("slides");
        List<object> now = DeckState(p);

        List<object> changed = new List<object>();
        int added = Math.Max(0, now.Count - before.Count);
        int removed = Math.Max(0, before.Count - now.Count);
        int common = Math.Min(now.Count, before.Count);
        for (int k = 0; k < common; k++)
        {
            Dictionary<string, object> beforeTexts = ShapeTextMap((List<object>)before[k]);
            Dictionary<string, object> nowTexts = ShapeTextMap((List<object>)now[k]);
            foreach (KeyValuePair<string, object> kv in nowTexts)
            {
                string oldT = beforeTexts.ContainsKey(kv.Key) ? (string)beforeTexts[kv.Key] : null;
                string newT = (string)kv.Value;
                if (!string.Equals(oldT, newT, StringComparison.Ordinal))
                    changed.Add(new Dictionary<string, object> {
                        { "slide", k + 1 }, { "shape", kv.Key },
                        { "before", Clip(oldT, 80) }, { "after", Clip(newT, 80) }
                    });
            }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["since"] = path;
        r["summary"] = new Dictionary<string, object> {
            { "slidesBefore", before.Count }, { "slidesNow", now.Count },
            { "slidesAdded", added }, { "slidesRemoved", removed }, { "textChanges", changed.Count }
        };
        r["changed"] = changed;
        r["note"] = "text-level diff: human edits and agent edits both count; added/removed slides change numbering, re-read the deck";
        return r;
    }

    private static Dictionary<string, object> ShapeTextMap(List<object> slideShapes)
    {
        Dictionary<string, object> m = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        foreach (object o in slideShapes)
        {
            Dictionary<string, object> sh = (Dictionary<string, object>)o;
            string name = Convert.ToString(sh["name"], CultureInfo.InvariantCulture);
            string key = name;
            int dup = 2;
            while (m.ContainsKey(key)) key = name + "#" + (dup++);
            m[key] = Convert.ToString(sh["text"], CultureInfo.InvariantCulture);
        }
        return m;
    }

    // ------------------------------------------------------------ save / close / quit / finish

    private static object SavePres(Bag a)
    {
        dynamic p = RequirePres(a, true);
        // never write the takeover badge into the file: strip it, save, stamp it back
        int markCount = 0;
        try { markCount = ApplySlideMark(p, false); } catch (Exception) { }
        try
        {
            return SavePresCore(a, p);
        }
        finally
        {
            if (markCount > 0)
            {
                try { ApplySlideMark(p, true); } catch (Exception) { }
                // the file on disk is clean now: record its stats so a later Ctrl+S is detectable
                try { MarkStateSet(Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture)); }
                catch (Exception) { }
            }
        }
    }

    private static object SavePresCore(Bag a, dynamic p)
    {
        string asPath = a.Str("as", null);
        if (asPath == null)
        {
            string full = Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture);
            if (full.Length == 0) throw new Fail("BAD_ARGS", "this presentation has never been saved; use save -As <path>");
            p.Save();
            return new Dictionary<string, object> { { "saved", full } };
        }
        if (asPath.ToLowerInvariant().EndsWith(".pdf", StringComparison.Ordinal))
        {
            // 32 = ppSaveAsPDF
            p.SaveAs(asPath, 32);
            return new Dictionary<string, object> { { "exported", asPath } };
        }
        int format = 24; // ppSaveAsOpenXMLPresentation (.pptx)
        string low = asPath.ToLowerInvariant();
        if (low.EndsWith(".ppt", StringComparison.Ordinal)) format = 1;
        string oldFull = Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture);
        p.SaveAs(asPath, format);
        MarkerRemove(oldFull);
        MarkerSet(FullPathOf(asPath));
        return new Dictionary<string, object> { { "savedAs", asPath } };
    }

    private static object ClosePres(Bag a)
    {
        string full = FullPathOf(a.Need("doc"));
        dynamic p = FindOpen(full);
        if (p == null) throw new Fail("NOT_OPEN", "The presentation is not open: " + full);
        bool save = a.Flag("save", false);
        bool saved = false;
        // drop the takeover markers BEFORE saving, so `close -Save` never persists them
        try { ApplyTitleMarker(p, false); } catch (Exception) { }
        int markCount = 0;
        try { markCount = ApplySlideMark(p, false); } catch (Exception) { }
        if (save)
        {
            bool dirty = true;
            try { dirty = (int)p.Saved == 0; } catch (Exception) { }
            if (dirty)
            {
                try { ComRetry(() => { p.Save(); return 0; }); saved = true; }
                catch (Exception e)
                {
                    if (markCount > 0) { try { ApplySlideMark(p, true); } catch (Exception) { } }
                    throw new Fail("COM_ERROR", "could not save before closing " + full + ": " + Clean(e.Message));
                }
            }
        }
        try { TitleStateSet(full, false); } catch (Exception) { }
        try { MarkStateClear(full); } catch (Exception) { }
        p.Close();
        MarkerRemove(full);   // otherwise a later `open -Doc` of the same path is refused as DOC_CLOSED
        TrayDaemon.Stop("ppt");   // drop the resident hint icon of this app, if any
        dynamic app = App(false, false);
        int left = 0;
        try { left = (int)app.Presentations.Count; }
        catch (Exception) { left = 0; }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["closed"] = full;
        r["saved"] = saved;
        r["presentationsLeft"] = left;
        if (left == 0)
        {
            try { app.Quit(); }
            catch (Exception) { }
            // v1.36: the automation instance is gone - do not keep a dead COM object cached.
            _app = null;
        }
        r["processes"] = ProcessCount("wpp");
        return r;
    }

    // ------------------------------------------------------------ hint
    // Takeover hint for WPS Presentation. Unlike Writer and Spreadsheets there is no writable
    // PER-DOCUMENT window caption here: ActiveWindow.Caption is read-only, and
    // Application.Caption is APP-wide and is printed as "<deck> - <caption>" - writing it
    // replaces the localized product name ("WPS 演示") with whatever we send, so a safe
    // restore would need persisted per-deck state (measured 2026-10-30). The hint therefore
    // stays a tray balloon - the same agent-agnostic mechanism finish uses - and the response
    // says so explicitly so no caller assumes a title was tagged.

    private static bool ShowBalloon(string title, string text)
    {
        try
        {
            if (text == null) text = "";
            if (text.Length > 240) text = text.Substring(0, 240);
            using (NotifyIcon ni = new NotifyIcon())
            {
                ni.Icon = SystemIcons.Information;
                ni.Visible = true;
                ni.BalloonTipTitle = title;
                ni.BalloonTipText = text;
                ni.ShowBalloonTip(6000);
                DateTime end = DateTime.Now.AddSeconds(4);
                while (DateTime.Now < end) { Application.DoEvents(); Thread.Sleep(50); }
                ni.Visible = false;
            }
            return true;
        }
        catch (Exception) { return false; }
    }

    private static bool PlayAsterisk()
    {
        try { System.Media.SystemSounds.Asterisk.Play(); Thread.Sleep(300); return true; }
        catch (Exception) { return false; }
    }

    // ------------------------------------------------------------ window title marker
    // WPS Presentation gives no writable per-document caption (ActiveWindow.Caption is read-only and
    // Application.Caption does not repaint the frame), but the real frame window - class
    // PP12FrameClass, owned by the shared wps.exe frame host - still accepts SetWindowText.
    // Measured 2026-11-02: the prefix survives COM edits, slide switches and idle; WPS resets the
    // title when the user switches to another tab, so the marker can never show on the wrong deck.
    private const string TitleTag = "\u3014AI\u63a5\u7ba1\u4e2d\u3015";
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool SetWindowText(IntPtr h, string s);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);

    private const string MarkShapeName = "woc-ai-takeover-mark";

    /// <summary>Delete every takeover badge (by name) from a shape collection.</summary>
    private static void DeleteSlideMarks(dynamic shapes)
    {
        try
        {
            int n = (int)shapes.Count;
            for (int k = n; k >= 1; k--)
            {
                dynamic sh = null;
                try { sh = shapes.Item(k); } catch (Exception) { continue; }
                string nm = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                if (string.Equals(nm, MarkShapeName, StringComparison.Ordinal))
                {
                    try { sh.Delete(); } catch (Exception) { }
                }
            }
        }
        catch (Exception) { }
    }

    /// <summary>An in-deck takeover badge: a small pill drawn on EVERY slide, because the WPS tab is
    /// custom-drawn from the file name and cannot show the OS title.
    /// Measured 2026-11-02: a badge on the slide master is hidden by any full-bleed picture
    /// (the deck's cover slide) since master shapes sit behind slide shapes - so the pill is
    /// stamped per slide instead (topmost shape, visible in edit view and in the slide show).
    /// Returns the number of slides carrying the badge.</summary>
    private static int ApplySlideMark(dynamic p, bool on)
    {
        try
        {
            double slideW = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
            double slideH = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
            double markW = Math.Round(slideW * 0.138, 2), markH = Math.Round(slideH * 0.048, 2);
            if (markW < 90) markW = 90;
            if (markH < 18) markH = 18;
            double left = Math.Round(slideW - markW - slideW * 0.0125, 2);
            if (left < 12) left = 12;
            double top = Math.Round(slideH * 0.022, 2);
            if (top < 8) top = 8;

            // remember where the user is looking, so stamping never moves the view
            // (View.CurrentShowPosition is empty on WPS 12; Selection.SlideRange reports the real slide)
            int pos = 0;
            try { pos = (int)p.Windows[1].Selection.SlideRange.SlideIndex; } catch (Exception) { }

            dynamic master = null;
            try { master = p.SlideMaster; } catch (Exception) { }
            if (master != null) DeleteSlideMarks(master.Shapes);   // legacy master badge from the v1.27 draft

            int slides = 0;
            try { slides = (int)p.Slides.Count; } catch (Exception) { return 0; }
            int marked = 0;
            for (int i = 1; i <= slides; i++)
            {
                try
                {
                    dynamic sl = p.Slides[i];
                    dynamic shapes = sl.Shapes;
                    DeleteSlideMarks(shapes);
                    if (!on) continue;
                    dynamic shape = shapes.AddShape(5, left, top, markW, markH);   // msoShapeRoundedRectangle
                    shape.Name = MarkShapeName;
                    // the theme's default shape style can bring a gradient / transparency with it
                    // (measured: the pill washed out on a plain white deck), so force a solid fill
                    try { shape.Fill.Solid(); } catch (Exception) { }
                    try { shape.Fill.Visible = -1; } catch (Exception) { }
                    try { shape.Fill.ForeColor.RGB = 0x227EE6; } catch (Exception) { }   // BGR for orange
                    try { shape.Fill.Transparency = 0; } catch (Exception) { }
                    try { shape.Line.Visible = 0; } catch (Exception) { }
                    try { shape.Shadow.Visible = 0; } catch (Exception) { }
                    dynamic tf = shape.TextFrame;
                    try { tf.MarginLeft = 0; tf.MarginRight = 0; tf.MarginTop = 0; tf.MarginBottom = 0; } catch (Exception) { }
                    try { tf.WordWrap = 0; } catch (Exception) { }
                    try { tf.VerticalAnchor = 3; } catch (Exception) { }          // msoAnchorMiddle
                    dynamic tr = tf.TextRange;
                    tr.Text = TitleTag;
                    try { tr.Font.Size = Math.Round(markH * 0.46, 1); } catch (Exception) { }
                    tr.Font.Bold = -1;
                    tr.Font.Color.RGB = 0xFFFFFF;
                    try { tr.ParagraphFormat.Alignment = 2; } catch (Exception) { }   // msoAlignCenter
                    marked++;
                }
                catch (Exception) { }
            }

            if (pos > 0)
            {
                try { p.Windows[1].View.GotoSlide(pos); } catch (Exception) { }
            }
            try
            {
                dynamic app = p.Application;
                dynamic actWin = app.ActiveWindow;
                if (actWin != null) actWin.Selection.Unselect();
            }
            catch (Exception) { }
            return marked;
        }
        catch (Exception) { return 0; }
    }

    // ------------------------------------------------- slide-mark integrity (v1.28)
    // The badge is deck CONTENT, so a manual Ctrl+S writes it into the file and no COM-level guard
    // can stop that. Instead we remember the file stats we last saw while the badge was on screen
    // and, on the next command, notice that the file was written behind our back: strip the badge,
    // save again, put the badge back and say so (badgeRepaired). Same strip -> save -> re-stamp
    // dance `save`/`export`/`close -Save` already do.

    private static string MarkStateDir()
    {
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                  "wps-office-cooperate", "ppt-slide-mark");
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static string MarkStatePath(string full)
    {
        return Path.Combine(MarkStateDir(), DocKey(full) + ".json");
    }

    private static Dictionary<string, object> MarkStateGet(string full)
    {
        try
        {
            string p = MarkStatePath(full);
            if (!File.Exists(p)) return null;
            return Json.Parse(File.ReadAllText(p, Encoding.UTF8)) as Dictionary<string, object>;
        }
        catch (Exception) { return null; }
    }

    /// <summary>Record the deck file's current stats, i.e. "this is what the file looks like while a
    /// badge is on screen". A later difference means somebody saved the deck with the badge in it.</summary>
    private static void MarkStateSet(string full)
    {
        try
        {
            if (full == null || full.Length == 0) return;
            if (!File.Exists(full)) { MarkStateClear(full); return; }
            FileInfo fi = new FileInfo(full);
            Dictionary<string, object> d = new Dictionary<string, object>();
            d["full"] = full;
            d["at"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            d["mtimeUtc"] = fi.LastWriteTimeUtc.ToString("o", CultureInfo.InvariantCulture);
            d["size"] = fi.Length;
            File.WriteAllText(MarkStatePath(full), Json.Write(d), new UTF8Encoding(false));
        }
        catch (Exception) { }
    }

    private static void MarkStateClear(string full)
    {
        try
        {
            if (full == null || full.Length == 0) return;
            string p = MarkStatePath(full);
            if (File.Exists(p)) File.Delete(p);
        }
        catch (Exception) { }
    }

    /// <summary>If the deck was saved while the badge was in it, remove the badge from the file.</summary>
    private static void EnsureSlideMarkIntegrity(Bag a)
    {
        try
        {
            string doc = a.Str("doc", null);
            if (doc == null || doc.Length == 0) return;
            string full = FullPathOf(doc);
            Dictionary<string, object> st = MarkStateGet(full);
            if (st == null) return;
            if (!File.Exists(full)) return;   // deck may be open-but-never-saved; keep the state

            FileInfo fi = new FileInfo(full);
            string recMtime = "";
            long recSize = -1;
            try { if (st.ContainsKey("mtimeUtc")) recMtime = Convert.ToString(st["mtimeUtc"], CultureInfo.InvariantCulture); }
            catch (Exception) { }
            try { if (st.ContainsKey("size")) recSize = Convert.ToInt64(st["size"], CultureInfo.InvariantCulture); }
            catch (Exception) { }
            bool same = string.Equals(recMtime, fi.LastWriteTimeUtc.ToString("o", CultureInfo.InvariantCulture), StringComparison.Ordinal)
                        && recSize == fi.Length;
            if (same) return;   // the file on disk is the one we recorded: badge still only on screen

            dynamic p = FindOpen(full);
            if (p == null) return;   // closed for now: repair on the next command that touches it

            // Strip whatever badge is still in the in-memory deck. Careful: an external save (the
            // user's Ctrl+S, or our own raw-COM test) can leave this client's object graph out of
            // sync with the file, so "no badge found in memory" does NOT mean the file is clean -
            // the save below is what actually rewrites the file without the badge.
            int had = 0;
            try { had = ApplySlideMark(p, false); } catch (Exception) { }

            bool ro = false;
            try { ro = (int)p.ReadOnly != 0; } catch (Exception) { }
            if (ro)
            {
                try { ApplySlideMark(p, true); } catch (Exception) { }
                Note("badgeDirty", "the deck was saved while the takeover badge was in it, but the file is read-only -"
                    + " run hint -Clear and save again before delivering it");
                return;
            }
            try { ComRetry(() => { p.Save(); return 0; }); }
            catch (Exception e)
            {
                try { ApplySlideMark(p, true); } catch (Exception) { }
                Note("badgeRepairFailed", "stripped the takeover badge but could not save the repaired file: " + Clean(e.Message));
                return;
            }
            int again = 0;
            try { again = ApplySlideMark(p, true); } catch (Exception) { }
            MarkStateSet(full);
            Note("badgeRepaired", "the deck had been saved (Ctrl+S) with the takeover badge; the badge was removed, the file"
                + " re-saved, and the badge put back on screen");
            Note("badgeMarksRemoved", had);
            Note("slideMarkCount", again);
        }
        catch (Exception) { }
    }

    private static IntPtr FrameFor(string presName)
    {
        IntPtr found = IntPtr.Zero;
        string want = presName ?? "";
        EnumWindowsProc cb = delegate(IntPtr h, IntPtr l)
        {
            StringBuilder cls = new StringBuilder(64);
            GetClassName(h, cls, cls.Capacity);
            if (!string.Equals(cls.ToString(), "PP12FrameClass", StringComparison.Ordinal)) return true;
            if (!IsWindowVisible(h)) return true;
            StringBuilder tb = new StringBuilder(1024);
            GetWindowText(h, tb, tb.Capacity);
            string title = tb.ToString();
            if (title.Length == 0) return true;
            if (want.Length > 0 && title.IndexOf(want, StringComparison.OrdinalIgnoreCase) < 0) return true;
            found = h;
            return false;
        };
        EnumWindows(cb, IntPtr.Zero);
        GC.KeepAlive(cb);
        return found;
    }

    /// <summary>Prefix / strip the takeover tag on the real frame title. Returns false when the
    /// deck is not the active tab (WPS has already reset the title to the other document).</summary>
    private static bool ApplyTitleMarker(dynamic p, bool on)
    {
        try
        {
            string name = Convert.ToString(TryGet(() => (object)p.Name, ""), CultureInfo.InvariantCulture);
            if (name.Length == 0) return false;
            IntPtr h = FrameFor(name);
            if (h == IntPtr.Zero) return false;
            StringBuilder tb = new StringBuilder(1024);
            GetWindowText(h, tb, tb.Capacity);
            string title = tb.ToString();
            if (on)
            {
                if (title.StartsWith(TitleTag, StringComparison.Ordinal)) return true;
                SetWindowText(h, TitleTag + " " + title);
                return true;
            }
            if (title.StartsWith(TitleTag + " ", StringComparison.Ordinal))
                SetWindowText(h, title.Substring(TitleTag.Length + 1));
            return true;
        }
        catch (Exception) { return false; }
    }

    private static string TitleStatePath()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate", "ppt-title-hint.json");
    }
    private static bool TitleStateGet(string full)
    {
        try
        {
            Dictionary<string, object> m = Json.Parse(File.ReadAllText(TitleStatePath(), Encoding.UTF8)) as Dictionary<string, object>;
            if (m != null && m.ContainsKey(full.ToLowerInvariant())) return Convert.ToInt32(m[full.ToLowerInvariant()], CultureInfo.InvariantCulture) == 1;
        }
        catch (Exception) { }
        return false;
    }
    private static void TitleStateSet(string full, bool on)
    {
        IDisposable g = WocLocks.State(3000);
        try
        {
            Dictionary<string, object> m = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            try
            {
                Dictionary<string, object> old = Json.Parse(File.ReadAllText(TitleStatePath(), Encoding.UTF8)) as Dictionary<string, object>;
                if (old != null) m = old;
            }
            catch (Exception) { }
            m[full.ToLowerInvariant()] = on ? 1 : 0;
            Directory.CreateDirectory(Path.GetDirectoryName(TitleStatePath()));
            string tmp = TitleStatePath() + ".tmp";
            File.WriteAllText(tmp, Json.Write(m), new UTF8Encoding(false));
            File.Copy(tmp, TitleStatePath(), true);
            try { File.Delete(tmp); } catch (Exception) { }
        }
        catch (Exception) { }
        finally { if (g != null) g.Dispose(); }
    }

    /// <summary>v1.36: repair a hidden Presentation window. WPS's own close path can crash and come
    /// back with the frame chrome up but the document window hidden (Application.Visible false) -
    /// the user sees an empty grey window and calls it "stuck". Every request that may touch the
    /// app repairs that first, so the next agent command un-sticks what the host broke.</summary>
    private static void HealWindow()
    {
        try
        {
            dynamic app = _app;
            if (app == null) return;
            bool visible;
            try { visible = (bool)app.Visible; } catch (Exception) { return; }
            if (visible) return;
            int n;
            try { n = (int)app.Presentations.Count; } catch (Exception) { return; }
            if (n <= 0) return;
            try { app.Visible = true; } catch (Exception) { return; }
            try
            {
                dynamic p = app.ActivePresentation;
                if (p != null)
                {
                    int wins = 0;
                    try { wins = (int)p.Windows.Count; } catch (Exception) { }
                    if (wins == 0) { try { p.NewWindow(); } catch (Exception) { } }
                    try { p.Windows(1).Activate(); } catch (Exception) { }
                }
            }
            catch (Exception) { }
            Console.Error.WriteLine("woc-helper: repaired a hidden presentation window (v1.36)");
        }
        catch (Exception) { }
    }

    /// <summary>Re-apply the marker on later commands: WPS drops it when the user switches tabs.</summary>
    private static void EnsureTitleMarker(Bag a)
    {
        try
        {
            string doc = a.Str("doc", null);
            if (doc == null) return;
            string full = FullPathOf(doc);
            if (!TitleStateGet(full)) return;
            dynamic app = App(false, false);
            if (app == null) return;
            dynamic act = null;
            try { act = app.ActivePresentation; } catch (Exception) { return; }
            string actFull = Convert.ToString(TryGet(() => (object)act.FullName, ""), CultureInfo.InvariantCulture);
            if (!string.Equals(actFull, full, StringComparison.OrdinalIgnoreCase)) return;
            ApplyTitleMarker(act, true);
        }
        catch (Exception) { }
    }

    // ------------------------------------------------------------ hint lease (v1.35)
    // The frame-title marker and the per-slide badge are the VISIBLE hint; the intent registry is
    // TitleState (ppt-title-hint.json). The resident helper calls this when the agent's lease
    // expires and on its way out (shared/HintLease.cs), so a session that simply stops - chat
    // ended, model error, host killed - cannot leave a window branded as taken over.
    // Nothing is written to the deck here: the badge is deck CONTENT, so it is only taken out of
    // memory (and its repair state dropped) - never saved.
    private static void ReleaseHints()
    {
        try { TrayDaemon.Stop("ppt"); } catch (Exception) { }
        dynamic app = null;
        try { app = App(false, false); } catch (Exception) { }
        if (app == null) return;
        int n = 0;
        try { n = (int)app.Presentations.Count; } catch (Exception) { return; }
        for (int k = 1; k <= n; k++)
        {
            try
            {
                dynamic p = app.Presentations.Item(k);
                string full = Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture);
                if (full.Length == 0) continue;
                // TitleState is kept ON deliberately: it is the INTENT, not the mark. The next
                // request re-applies the prefix through EnsureTitleMarker, so a session that only
                // paused comes back branded instead of silently unbranded.
                if (TitleStateGet(full)) ApplyTitleMarker(p, false);
                // The badge is in-memory deck content: strip it, but only when MarkState says a
                // badge was stamped - that snapshot is written by MarkStateSet right after
                // stamping, so null means "no badge in this deck" and the slide walk (expensive:
                // every slide, plus a view move/restore) can be skipped. Dropping the snapshot
                // keeps the next command from "repairing" a badge nobody ever saved. Never save.
                if (MarkStateGet(full) != null)
                {
                    ApplySlideMark(p, false);
                    MarkStateClear(full);   // no badge left, nothing to repair later
                }
            }
            catch (Exception) { }
        }
    }

    private static object HintCmd(Bag a)
    {
        dynamic p = RequirePres(a, false);
        bool clear = a.Flag("clear", false);
        string name = Convert.ToString(TryGet(() => (object)p.Name, ""), CultureInfo.InvariantCulture);
        string text = a.Str("text", null);
        if (text == null)
            text = clear
                ? "\u5df2\u7ed3\u675f\u534f\u4f5c\u300a" + name + "\u300b\u3002"   // 已结束协作《...》。
                : "\u5df2\u63a5\u7ba1\u300a" + name + "\u300b\uff1a\u6539\u52a8\u4f1a\u76f4\u63a5\u5199\u5165\u6f14\u793a\u6587\u7a3f\uff0c\u6539\u4e86\u54ea\u51e0\u9875\u6211\u4f1a\u7528 diff \u544a\u8bc9\u4f60\u3002"; // 已接管《...》：改动会直接写入演示文稿…

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = name;
        r["cleared"] = clear;
        bool titleMarked = false;
        string full = Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture);
        try
        {
            titleMarked = ApplyTitleMarker(p, !clear);
            if (full.Length > 0) TitleStateSet(full, !clear);
        }
        catch (Exception) { }
        r["captionSupported"] = false;   // COM caption stays read-only
        r["titleMarked"] = titleMarked;  // the real marker is a Win32 SetWindowText on the frame
        // v1.28: the slide badge is deck CONTENT (a manual Ctrl+S writes it into the file), so it is
        // opt-in. Default: window title + balloon + optional tray only, zero document changes.
        bool wantSlideMark = !clear && a.Flag("slideMark", false);
        int slideMarkCount = 0;
        try
        {
            // false covers both -Clear and -NoSlideMark: either way no badge may stay behind
            slideMarkCount = ApplySlideMark(p, wantSlideMark);
        }
        catch (Exception) { }
        r["slideMarked"] = slideMarkCount > 0;
        if (slideMarkCount > 0) r["slideMarkCount"] = slideMarkCount;
        try
        {
            if (slideMarkCount > 0) MarkStateSet(full);
            else if (full.Length > 0) MarkStateClear(full);
        }
        catch (Exception) { }
        r["limitation"] = titleMarked
            ? "the per-document caption is read-only in COM, so the marker is written on the real window frame"
              + " (class PP12FrameClass); WPS clears it when you switch to another tab and the next command re-applies it"
            : "no visible frame matched this deck (it may not be the active tab); the tray balloon is the only signal";
        if (slideMarkCount > 0)
            r["slideMarkWarning"] = "the badge is deck content: if the user presses Ctrl+S it is written into the file."
                + " Call finish (or hint -Clear) before the deck is saved or delivered; should the file get saved anyway,"
                + " the next command strips the badge, re-saves and reports badgeRepaired:true";
        if (!clear && slideMarkCount == 0)
            r["inAppHint"] = "nothing was written into the deck (default). The WPS Presentation tab is custom-drawn from the"
                + " file name, so the OS title is only visible in the taskbar / Alt+Tab. Pass -SlideMark to stamp the"
                + " in-app badge on every slide - it is deck content, so it has to be cleared again (finish / hint -Clear)"
                + " before the deck is saved or handed over";

        bool quiet = a.Flag("quiet", false);
        bool toast = !quiet && a.Flag("toast", !clear);
        if (toast) toast = ShowBalloon("WPS \u6f14\u793a AI \u52a9\u624b", text);  // WPS 演示 AI 助手
        r["toast"] = toast;

        bool sound = !quiet && a.Flag("sound", false);
        if (sound) sound = PlayAsterisk();
        r["sound"] = sound;

        // resident tray icon (opt-in) - for Presentation this is the ONLY persistent visual
        // channel: no writable caption, no status bar.
        if (clear) r["trayStopped"] = TrayDaemon.Stop("ppt");
        else if (a.Flag("tray", false))
            r["tray"] = TrayDaemon.Start("ppt", Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture),
                                         "\u3014AI\u3015 " + name, text, a.Str("wrapper", null));
        r["processes"] = ProcessCount("wpp");
        return r;
    }

    private static object FinishCmd(Bag a)
    {
        // Notification only. Never touches the presentation (there is no good
        // "status line" spot on a slide; writing one would pollute the deck).
        string note = a.Str("note", "");
        string timeText = DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        string text = "\u2705 AI \u5df2\u5b8c\u6210 " + timeText + (note.Length > 0 ? " | " + note : "");
        // clear the takeover markers on the deck this call is about - never on whatever tab
        // happens to be active (that would strip another deck's badge / title prefix)
        dynamic target = null;
        try
        {
            dynamic app2 = App(false, false);
            string docArg = a.Str("doc", null);
            if (docArg != null && docArg.Length > 0)
            {
                try { target = FindOpen(FullPathOf(docArg)); } catch (Exception) { }
            }
            if (target == null) { try { target = app2.ActivePresentation; } catch (Exception) { } }
            if (target != null) ApplySlideMark(target, false);
        }
        catch (Exception) { }
        try
        {
            if (target != null)
            {
                ApplyTitleMarker(target, false);
                string tfull = Convert.ToString(TryGet(() => (object)target.FullName, ""), CultureInfo.InvariantCulture);
                TitleStateSet(tfull, false);
                MarkStateClear(tfull);   // no badge left, nothing to repair later
            }
        }
        catch (Exception) { }

        Dictionary<string, object> r = new Dictionary<string, object>();
        bool quiet = a.Flag("quiet", false);
        bool toast = !quiet && a.Flag("toast", true);
        if (toast)
        {
            try
            {
                string t = text;
                if (t.Length > 240) t = t.Substring(0, 240);
                using (NotifyIcon ni = new NotifyIcon())
                {
                    ni.Icon = SystemIcons.Information;
                    ni.Visible = true;
                    ni.BalloonTipTitle = "WPS \u6f14\u793a AI \u52a9\u624b";  // WPS 演示 AI 助手
                    ni.BalloonTipText = t;
                    ni.ShowBalloonTip(6000);
                    DateTime end = DateTime.Now.AddSeconds(4);
                    while (DateTime.Now < end) { Application.DoEvents(); Thread.Sleep(50); }
                    ni.Visible = false;
                }
            }
            catch (Exception) { toast = false; }
        }
        r["toast"] = toast;
        bool sound = !quiet && a.Flag("sound", true);
        if (sound)
        {
            try { System.Media.SystemSounds.Asterisk.Play(); }
            catch (Exception) { sound = false; }
            Thread.Sleep(300);
        }
        r["sound"] = sound;
        r["trayStopped"] = TrayDaemon.Stop("ppt");
        r["note"] = "finish touches nothing in the deck; save before finishing if the work should persist";
        return r;
    }

    private static object QuitAll(Bag a)
    {
        dynamic app = App(false, false);
        if (app == null) return new Dictionary<string, object> { { "processes", ProcessCount("wpp") } };
        try { app.Quit(); }
        catch (Exception) { }
        _app = null;   // v1.36: quit() leaves the cached object dead; drop it now
        return new Dictionary<string, object> { { "processes", ProcessCount("wpp") } };
    }
}
