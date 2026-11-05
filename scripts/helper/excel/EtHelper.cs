// wps-excel-cooperate helper: drive WPS Spreadsheets (ET) over COM to read / write / comment /
// render a .xlsx workbook live. Pure ASCII source (Chinese only ever travels inside JSON
// strings), C# 5 (compiled by the in-box csc.exe of .NET Framework 4.x), no third-party
// dependency.
//
//   EtHelper.exe <request.json> <response.json>
//
// The request is {"cmd":"...", ...}; the response is {"ok":true,"result":...} or
// {"ok":false,"error":{"code":"...","message":"..."}}. Files are read and written as UTF-8,
// so no console code page is involved.

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
}

// ---------------------------------------------------------------- helper

internal static partial class Program
{
    private const string WPS_ET = "KET.Application";
    private const string MS_EXCEL = "Excel.Application";

    private static dynamic _app;

    [STAThread]
    private static int Main(string[] args)
    {
        // v1.32: install the OLE message filter before any COM call, so a rejected call while the
        // user is typing in a cell is retried for up to 6 s instead of failing at once
        // (helper/shared/MessageFilter.cs). One registration covers every command: all requests run
        // on this STA thread.
        MessageFilter.Register();
        Console.OutputEncoding = new UTF8Encoding(false);
        // Detached tray daemon mode: same binary, message-loop instead of a COM request.
        if (args != null && args.Length >= 2 && args[0] == "--tray") return TrayDaemon.Run(args[1]);
        // v1.15 resident mode: keep the COM application alive and serve request JSON over a named
        // pipe (same transport as the Word helper - see helper/shared/PipeServer.cs).
        if (args != null && args.Length >= 2 && args[0] == "--serve")
            return PipeServer.Run(args[1], "excel", HandleRequest, OpenDocumentCount, ReleaseHints);

        string requestPath = args != null && args.Length > 0 ? args[0] : null;
        string responsePath = args != null && args.Length > 1 ? args[1] : null;
        if (requestPath == null)
        {
            Console.Out.Write("{\"ok\":false,\"error\":{\"code\":\"BAD_ARGS\",\"message\":\"usage: EtHelper.exe <request.json> <response.json> | --serve <pipe> | --tray <state>\"}}");
            return 1;
        }
        string text = HandleRequest(File.ReadAllText(requestPath, Encoding.UTF8));
        if (responsePath != null) File.WriteAllText(responsePath, text, new UTF8Encoding(false));
        else Console.Out.Write(text);
        return 0;
    }

    /// <summary>Set by "daemon-stop": the resident loop exits after answering that request.</summary>
    private static bool _daemonStop;

    /// <summary>Open documents in this automation instance (the daemon's exit rule).</summary>
    private static int OpenDocumentCount()
    {
        try
        {
            if (_daemonStop) return 0;
            if (_app == null) return 0;
            return (int)_app.Workbooks.Count;
        }
        catch (Exception) { return 0; }
    }

    /// <summary>One request JSON in, one response JSON out (shared by one-shot and resident modes).</summary>
    private static string HandleRequest(string requestJson)
    {
        object response;
        try
        {
            Bag bag = new Bag(Json.Parse(requestJson));
            string cmd = bag.Str("cmd", "").Trim().ToLowerInvariant();
            if (cmd.Length == 0) throw new Fail("BAD_ARGS", "\"cmd\" is required");
            // v1.35 hint lease: put the takeover mark back before doing anything else (WPS resets
            // the caption on a tab switch, and the resident helper strips it when the agent goes
            // quiet). `open`/`new` apply it themselves; hint/finish/close own the intent registry.
            if (cmd != "open" && cmd != "new" && cmd != "hint" && cmd != "finish" && cmd != "close")
            {
                // v1.36: the same "first thing on contact" slot also repairs the blank-window state
                // WPS can leave behind (see HealWindow).
                HealWindow();
                EnsureHint(bag);
            }
            response = Ok(cmd, Dispatch(cmd, bag));
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
                || e.HResult == unchecked((int)0x800AC472)        // the app is busy (typing in a cell)
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
            case "open": return OpenBook(a);
            case "new": return NewBook(a);
            case "read": return ReadBook(a);
            case "search": return SearchBook(a);
            case "selection": return SelectionBook(a);
            case "rows": return RowsCols(a, true);
            case "cols": return RowsCols(a, false);
            case "write": return WriteCells(a);
            case "fill": return FillFormula(a);
            case "sheet": return SheetCmd(a);
            case "comment": return CommentCmd(a);
            case "render": return RenderBook(a);
            case "chart": return ChartCmd(a);
            case "autofit": return AutofitCmd(a);
            case "format": return FormatRange(a);
            case "snapshot": return SnapshotBook(a);
            case "diff": return DiffBook(a);
            case "restore": return RestoreBook(a);
            case "calc": return CalcBook(a);
            case "save": return SaveBook(a);
            case "close": return CloseBook(a);
            case "quit": return QuitAll(a);
            case "hint": return HintCmd(a);
            case "finish": return FinishCmd(a);
            // Data work and the "finished look" commands live in EtData.cs.
            case "clean": return CmdClean(a);
            case "fillblanks": return CmdFillBlanks(a);
            case "dedupe": return CmdDedupe(a);
            case "deletewhere": return CmdDeleteWhere(a);
            case "sort": return CmdSort(a);
            case "filter": return CmdFilter(a);
            case "copy": return CmdCopy(a);
            case "addcol": return CmdAddColumn(a);
            case "pivot": return CmdPivot(a);
            case "freeze": return CmdFreeze(a);
            case "cformat": return CmdConditionalFormat(a);
            case "validate": return CmdValidate(a);
            case "table": return CmdTable(a);
            case "replace": return CmdReplace(a);
            default: throw new Fail("BAD_ARGS", "unknown cmd \"" + cmd + "\"; use doctor|status|open|new|read|search|selection|write|fill|rows|cols|sheet|comment|render|chart|autofit|format|snapshot|diff|restore|calc|save|close|quit|finish|hint|clean|fillblanks|dedupe|deletewhere|sort|filter|copy|addcol|pivot|freeze|cformat|validate|table");
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
    /// points at (the user closed WPS, WPS crashed on close - a daily occurrence on this machine,
    /// or its frame was replaced). Workbooks.Count is the cheapest liveness probe. Only "the object
    /// is gone" HRESULTs invalidate the cache: busy/rejected calls (typing in a cell, a dialog) are
    /// transient and must NOT throw away a healthy app object.
    /// </summary>
    private static bool StillAlive(dynamic app)
    {
        try { int n = (int)app.Workbooks.Count; return true; }
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

    /// <summary>The spreadsheet app: the running WPS ET instance if there is one, otherwise a new one.</summary>
    private static dynamic App(bool create, bool visible)
    {
        if (_app != null)
        {
            if (StillAlive(_app)) return _app;
            _app = null;   // detached or died underneath us: re-attach below
        }
        dynamic found = Running(WPS_ET);
        if (found != null) { _app = found; return _app; }
        if (!create) return null;
        if (!Installed(WPS_ET))
            throw new Fail("NO_APP", "WPS Spreadsheets (KET.Application) is not installed on this computer (Excel.Application is only used as a fallback for attach, never to create).");
        try { _app = Activator.CreateInstance(Type.GetTypeFromProgID(WPS_ET, true)); }
        catch (Exception e) { throw new Fail("NO_APP", "Could not start WPS Spreadsheets: " + Clean(e.Message)); }
        try { _app.Visible = visible; } catch (Exception) { }
        return _app;
    }

    private static string FullPathOf(string path)
    {
        try { return Path.GetFullPath(path); }
        catch (Exception) { return path; }
    }

    private static dynamic FindOpen(string fullPath)
    {
        dynamic app = App(false, false);
        if (app == null) return null;
        int n = 0;
        try { n = (int)app.Workbooks.Count; }
        catch (Exception) { return null; }
        string want = fullPath.ToLowerInvariant();
        for (int k = 1; k <= n; k++)
        {
            dynamic wb = null;
            try { wb = app.Workbooks.Item(k); }
            catch (Exception) { continue; }
            string full = null;
            try { full = (string)wb.FullName; }
            catch (Exception) { }
            if (full != null && full.ToLowerInvariant() == want) return wb;
        }
        return null;
    }

    private static dynamic RequireBook(Bag a, bool create)
    {
        string path = a.Need("doc");
        string full = FullPathOf(path);
        dynamic wb = FindOpen(full);
        if (wb != null) { MarkerSet(full); return wb; }
        // Attached earlier but gone now: the user (or something else) closed the workbook.
        // Never silently reopen it - a deliberate close may mean "stop" or an intentional
        // discard of unsaved changes. Surface it and let a explicit open re-attach.
        if (MarkerHas(full))
            throw new Fail("DOC_CLOSED",
                "The workbook " + full + " was attached by this skill earlier but is no longer open"
                + " - it was probably closed by the user just now (any unsaved changes were lost with it)."
                + " Do not keep working silently: if the user wants to continue, reopen it explicitly with open -Doc.");
        if (!create) throw new Fail("NOT_OPEN", "The workbook is not open in WPS Spreadsheets: " + full + ". Call open first.");
        dynamic opened = OpenPath(full, a.Flag("show", true), a.Flag("readOnly", false));
        MarkerSet(full);
        return opened;
    }

    // ------------------------------------------------------------ attachment marker
    // Remembers which workbooks this skill has attached, so a workbook that disappears
    // mid-collaboration is reported as DOC_CLOSED instead of being silently reopened.

    private static string MarkerPath()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate", "excel-attached.json");
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
        if (!File.Exists(full)) throw new Fail("NOT_OPEN", "The file does not exist: " + full);
        dynamic app = App(true, show);
        dynamic wb;
        try { wb = app.Workbooks.Open(full, UpdateLinks: 0, ReadOnly: readOnly); }
        catch (COMException e)
        {
            string msg = Clean(e.Message);
            if (msg.IndexOf("read-only", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("read only", StringComparison.OrdinalIgnoreCase) >= 0)
                wb = app.Workbooks.Open(full, UpdateLinks: 0, ReadOnly: true);
            else throw;
        }
        if (show)
        {
            try { app.Visible = true; } catch (Exception) { }
            try { wb.Activate(); } catch (Exception) { }
        }
        return wb;
    }

    private static bool IsReadOnly(dynamic wb)
    {
        object v = TryGet(() => (object)wb.ReadOnly, null);
        return v is bool && (bool)v;
    }

    private static Dictionary<string, object> BookInfo(dynamic wb)
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["name"] = Convert.ToString(TryGet(() => (object)wb.Name, ""), CultureInfo.InvariantCulture);
        d["path"] = Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture);
        d["readOnly"] = IsReadOnly(wb);
        d["saved"] = TryGet(() => (object)wb.Saved, true);
        d["sheets"] = SheetsBrief(wb);
        return d;
    }

    private static List<object> SheetsBrief(dynamic wb)
    {
        List<object> list = new List<object>();
        int n = 0;
        try { n = (int)wb.Worksheets.Count; } catch (Exception) { return list; }
        dynamic active = ActiveSheetSafe(wb);
        string activeName = active == null ? "" : Convert.ToString(TryGet(() => (object)active.Name, ""), CultureInfo.InvariantCulture);
        for (int k = 1; k <= n; k++)
        {
            dynamic ws = null;
            try { ws = wb.Worksheets.Item(k); } catch (Exception) { continue; }
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["index"] = k;
            one["name"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
            string vis = Convert.ToString(TryGet(() => (object)ws.Visible, "-1"), CultureInfo.InvariantCulture);
            one["visible"] = vis == "1" || vis == "-1";
            one["active"] = activeName.Length > 0 && (string)one["name"] == activeName;
            object used = TryGet(() => (object)ws.UsedRange.Address[false, false], null);
            one["usedRange"] = used == null ? "" : Convert.ToString(used, CultureInfo.InvariantCulture);
            list.Add(one);
        }
        return list;
    }

    private static Dictionary<string, object> AppInfo(dynamic app)
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["name"] = Convert.ToString(TryGet(() => (object)app.Name, ""), CultureInfo.InvariantCulture);
        d["version"] = Convert.ToString(TryGet(() => (object)app.Version, ""), CultureInfo.InvariantCulture);
        d["build"] = Convert.ToString(TryGet(() => (object)app.Build, ""), CultureInfo.InvariantCulture);
        d["path"] = Convert.ToString(TryGet(() => (object)app.Path, ""), CultureInfo.InvariantCulture);
        d["visible"] = TryGet(() => (object)(bool)app.Visible, false);
        return d;
    }

    private static int ProcessCount(string name)
    {
        try { return Process.GetProcessesByName(name).Length; }
        catch (Exception) { return -1; }
    }

    private static dynamic SheetByName(dynamic wb, string name)
    {
        int n = 0;
        try { n = (int)wb.Worksheets.Count; } catch (Exception) { }
        for (int k = 1; k <= n; k++)
        {
            dynamic ws = null;
            try { ws = wb.Worksheets.Item(k); } catch (Exception) { continue; }
            string nm = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
            if (string.Equals(nm, name, StringComparison.OrdinalIgnoreCase)) return ws;
        }
        throw new Fail("BAD_ARGS", "No sheet named \"" + name + "\" in this workbook.");
    }

    /// <summary>Resolve the sheet to work on: the -Sheet name, else the active sheet.</summary>
    // wb.ActiveSheet is not trustworthy on WPS 12: it can disagree with the sheet the user is
    // actually looking at, and it flaps between two consecutive commands (reported 2026-11-01 on
    // a 3-sheet workbook: open said Sheet2 active, the next command saw Sheet1). The window's own
    // ActiveSheet tracks what is on screen, so prefer it and fall back only if it is unavailable.
    private static dynamic ActiveSheetSafe(dynamic wb)
    {
        dynamic ws = TryGet(() => (object)wb.Windows(1).ActiveSheet, null);
        if (ws != null) return ws;
        return TryGet(() => (object)wb.ActiveSheet, null);
    }

    private static dynamic TargetSheet(dynamic wb, Bag a)
    {
        string name = a.Str("sheet", null);
        return name == null ? (dynamic)ActiveSheetSafe(wb) : SheetByName(wb, name);
    }

    /// <summary>Split an optional "Sheet!A1:B3" reference into sheet name + plain address.</summary>
    private static void SplitRef(string refText, out string sheetName, out string address)
    {
        sheetName = null; address = refText == null ? null : refText.Trim();
        if (address == null) return;
        int bang = address.IndexOf('!');
        if (bang > 0)
        {
            sheetName = address.Substring(0, bang).Trim('\'');
            address = address.Substring(bang + 1).Trim();
        }
    }

    // ------------------------------------------------------------ cell values

    /// <summary>Read Value2 of a range into a 1-based 2D array of objects (scalars boxed for single cells).</summary>
    private static object[,] ReadValues(dynamic range)
    {
        object raw = range.Value2;
        if (raw == null || raw is string || raw is double || raw is long || raw is int || raw is bool)
        {
            object[,] one = new object[2, 2];
            one[1, 1] = raw;
            return one;
        }
        return (object[,])raw;
    }

    private static object JsonCellValue(object v)
    {
        if (v == null) return "";
        // An error cell comes back as a negative integer; left alone it reads "-2146826281".
        string error = ErrorText(v);
        if (error != null) return error;
        if (v is double) { double d = (double)v; if (d == Math.Floor(d) && Math.Abs(d) < 1e15) return (long)d; return d; }
        if (v is bool) return v;
        return Convert.ToString(v, CultureInfo.InvariantCulture);
    }

    // ------------------------------------------------------------ commands

    private static object Doctor(Bag a)
    {
        Dictionary<string, object> r = new Dictionary<string, object>();
        Dictionary<string, object> apps = new Dictionary<string, object>();
        string[][] table = new string[][] {
            new string[] {"wpsEt", WPS_ET}, new string[] {"excel", MS_EXCEL},
        };
        foreach (string[] pair in table)
        {
            Dictionary<string, object> one = new Dictionary<string, object>();
            bool installed = Installed(pair[1]);
            one["progId"] = pair[1];
            one["installed"] = installed;
            one["running"] = installed && Running(pair[1]) != null;
            apps[pair[0]] = one;
        }
        r["apps"] = apps;
        r["processes"] = new Dictionary<string, object> {
            {"et", ProcessCount("et")}, {"excel", ProcessCount("excel")}
        };
        r["render"] = RenderCapable();

        dynamic app = App(false, false);
        if (app != null)
        {
            r["app"] = AppInfo(app);
            r["openWorkbooks"] = Workbooks(app);
        }
        else
        {
            r["app"] = null;
            r["openWorkbooks"] = new List<object>();
        }
        r["undoNote"] = "External COM edits do NOT enter the spreadsheet undo stack and clear the user's own undo history: verify with render/read before writing, snapshot before big batches.";
        return r;
    }

    private static Dictionary<string, object> RenderCapable()
    {
        Dictionary<string, object> probe = new Dictionary<string, object>();
        try
        {
            dynamic app = App(false, false);
            if (app == null) { probe["renderOk"] = false; probe["note"] = "no running instance to probe"; return probe; }
            dynamic wb = TryGet(() => (object)app.ActiveWorkbook, null);
            if (wb == null) { probe["renderOk"] = true; probe["note"] = "clipboard CopyPicture assumed available; verified when a workbook is open"; return probe; }
            dynamic ws = ActiveSheetSafe(wb);
            if (ws == null) { probe["renderOk"] = false; return probe; }
            string png = Path.Combine(Path.GetTempPath(), "et-render-probe.png");
            bool ok = TryRender(ws, "A1:A1", png);
            try { File.Delete(png); } catch (Exception) { }
            probe["renderOk"] = ok;
        }
        catch (Exception e) { probe["renderOk"] = false; probe["error"] = Clean(e.Message); }
        return probe;
    }

    private static List<object> Workbooks(dynamic app)
    {
        List<object> list = new List<object>();
        int n = 0;
        try { n = (int)app.Workbooks.Count; } catch (Exception) { return list; }
        dynamic active = TryGet(() => (object)app.ActiveWorkbook, null);
        string activeName = active == null ? "" : Convert.ToString(TryGet(() => (object)active.Name, ""), CultureInfo.InvariantCulture);
        for (int k = 1; k <= n; k++)
        {
            dynamic wb = null;
            try { wb = app.Workbooks.Item(k); } catch (Exception) { continue; }
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["name"] = Convert.ToString(TryGet(() => (object)wb.Name, ""), CultureInfo.InvariantCulture);
            one["path"] = Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture);
            one["readOnly"] = IsReadOnly(wb);
            one["saved"] = TryGet(() => (object)wb.Saved, true);
            one["active"] = activeName.Length > 0 && (string)one["name"] == activeName;
            try
            {
                string cap = CaptionOf(wb);
                if (cap != null) one["caption"] = cap;
                one["hinted"] = IsHinted(cap);
            }
            catch (Exception) { }
            list.Add(one);
        }
        return list;
    }

    private static object Status(Bag a)
    {
        Dictionary<string, object> r = new Dictionary<string, object>();
        dynamic app = App(false, false);
        r["running"] = app != null;
        r["processes"] = ProcessCount("et");
        r["workbooks"] = app == null ? new List<object>() : Workbooks(app);
        if (app != null) r["app"] = AppInfo(app);
        return r;
    }

    private static object OpenBook(Bag a)
    {
        string full = FullPathOf(a.Need("doc"));
        bool show = a.Flag("show", true);
        dynamic existing = FindOpen(full);
        dynamic wb;
        string how;
        int etBefore = ProcessCount("et");
        bool otherInstancePossible = false;
        if (existing != null)
        {
            wb = existing; how = "reused";
            if (show)
            {
                try { dynamic app2 = App(false, false); if (app2 != null) app2.Visible = true; } catch (Exception) { }
                try { wb.Activate(); } catch (Exception) { }
            }
        }
        else
        {
            // GetActiveObject only ever reaches the FIRST registered instance. If more than one
            // ET process runs and the reachable instance does not have the workbook, it may be
            // open in another one - opening here would fork the file. Say so.
            dynamic attached = App(false, false);
            otherInstancePossible = attached != null && etBefore > 1;
            wb = OpenPath(full, show, a.Flag("readOnly", false)); how = "opened";
        }
        Dictionary<string, object> info = BookInfo(wb);
        info["how"] = how;
        // Takeover hint: pure window state, applied on attach (see HintCmd).
        if (!a.Flag("noHint", false)) info["hint"] = HintOnTakeover(wb, a);
        else HintStateSet(Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture), null);
        MarkerSet(full);
        if (otherInstancePossible)
        {
            info["multiInstanceWarning"] = "ET \u8fd0\u884c\u4e2d\u6709 " + etBefore
                + " \u4e2a\u8fdb\u7a0b\uff0c\u4f46\u672c skill \u80fd\u63a5\u5230\u7684\u5b9e\u4f8b\u91cc\u6ca1\u6709\u8fd9\u4efd\u5de5\u4f5c\u7c3f\uff1a\u5b83\u53ef\u80fd\u6b63\u5f00\u5728\u53e6\u4e00\u4e2a\u5b9e\u4f8b\u91cc\uff0c\u7ee7\u7eed\u5199\u5165\u4f1a\u4e92\u76f8\u8986\u76d6\u3002";
        }
        return info;
    }

    // SaveAs file-format constant for a path. 51 = xlOpenXMLWorkbook (.xlsx).
    private static int WorkbookFormatFor(string path)
    {
        int format = 51;
        string low = path.ToLowerInvariant();
        if (low.EndsWith(".xls", StringComparison.Ordinal)) format = 56;
        else if (low.EndsWith(".csv", StringComparison.Ordinal)) format = 6;
        else if (low.EndsWith(".xlsm", StringComparison.Ordinal)) format = 52;
        return format;
    }

    // Create a brand-new workbook and save it to -Doc. Before v1.16 the only way to get an
    // empty workbook was open, which refuses paths that do not exist yet, so agents had to
    // hand-roll an openpyxl script and then re-open it (observed 2026-11-01: 22 steps for
    // "create a blank xlsx and open it"). Overwrite deletes the target first: a SaveAs onto an
    // existing path can raise an interactive "already exists" dialog that hangs the helper.
    private static object NewBook(Bag a)
    {
        string full = FullPathOf(a.Need("doc"));
        bool show = a.Flag("show", true);
        bool overwrite = a.Flag("overwrite", false);
        if (File.Exists(full))
        {
            if (!overwrite)
                throw new Fail("EXISTS", "The file already exists: " + full + ". Pass -Overwrite to replace it, or use open to keep it.");
            // Replacing a workbook that is currently open in WPS cannot work: Windows locks the
            // file, and even if the delete succeeded the open copy would save over the new one.
            if (FindOpen(full) != null)
                throw new Fail("IN_USE", "The workbook is open in WPS: " + full + ". Close it first (close -Doc) or create a different file.");
            try { File.Delete(full); }
            catch (Exception e) { throw new Fail("IN_USE", "cannot replace the existing file (is it open elsewhere?): " + Clean(e.Message)); }
        }
        string dir = Path.GetDirectoryName(full);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        dynamic app = App(true, show);
        dynamic wb;
        try { wb = app.Workbooks.Add(); }
        catch (Exception e) { throw new Fail("COM_ERROR", "Workbooks.Add failed: " + Clean(e.Message)); }
        try { wb.SaveAs(full, WorkbookFormatFor(full)); }
        catch (Exception e)
        {
            try { wb.Close(false); } catch (Exception) { }
            throw new Fail("COM_ERROR", "SaveAs failed for " + full + ": " + Clean(e.Message));
        }
        MarkerSet(full);
        Dictionary<string, object> info = BookInfo(wb);
        info["how"] = "created";
        info["note"] = "brand-new workbook saved to disk; write to it like any opened workbook";
        if (!a.Flag("noHint", false)) info["hint"] = HintOnTakeover(wb, a);
        else HintStateSet(Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture), null);
        return info;
    }

    private static object ReadBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string rangeText = a.Str("range", null);
        string sheetName, address;
        if (rangeText != null)
        {
            string refSheet;
            SplitRef(rangeText, out refSheet, out address);
            sheetName = refSheet != null ? refSheet : a.Str("sheet", null);
        }
        else { sheetName = a.Str("sheet", null); address = null; }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["book"] = BookInfo(wb);
        if (sheetName == null && address == null)
            return r; // book overview: sheets are already in book.sheets

        dynamic ws = sheetName == null ? (dynamic)ActiveSheetSafe(wb) : SheetByName(wb, sheetName);
        if (ws == null) throw new Fail("COM_ERROR", "No active sheet in this workbook.");
        dynamic range;
        if (address == null || address.Length == 0)
        {
            range = ws.UsedRange;
            address = Convert.ToString(TryGet(() => (object)range.Address[false, false], ""), CultureInfo.InvariantCulture);
        }
        else range = ws.Range(address);

        int rows = (int)range.Rows.Count, cols = (int)range.Columns.Count;
        object[,] values = ReadValues(range);
        object[,] formulas = ReadValues0(range.Formula);

        // Auto-profile for large blocks so the caller's context does not blow up.
        bool forceProfile = a.Flag("profile", false);
        if (forceProfile || rows > 60 || cols > 15)
        {
            r["profile"] = ProfileOf(range, address, values, rows, cols);
            return r;
        }

        // A date shown as a serial number is useless to a reader: send it as a date.
        bool[] dated = DateColumnsOf(range, values, rows, cols);
        List<object> table = new List<object>();
        for (int i = 1; i <= rows; i++)
        {
            List<object> row = new List<object>();
            for (int j = 1; j <= cols; j++)
            {
                object cellValue = values[i, j];
                object shown = cellValue;
                if (dated[j] && IsNum(cellValue))
                {
                    try
                    {
                        DateTime started = DateTime.FromOADate(Convert.ToDouble(cellValue, CultureInfo.InvariantCulture));
                        shown = (started.Hour == 0 && started.Minute == 0 && started.Second == 0)
                            ? started.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                            : started.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
                    }
                    catch (Exception) { shown = cellValue; }
                }
                object f = formulas[i, j];
                string fs = f == null ? null : Convert.ToString(f, CultureInfo.InvariantCulture);
                if (fs != null && fs.StartsWith("=", StringComparison.Ordinal))
                {
                    Dictionary<string, object> cell = new Dictionary<string, object>();
                    cell["v"] = JsonCellValue(shown);
                    cell["f"] = fs;
                    row.Add(cell);
                }
                else row.Add(JsonCellValue(shown));
            }
            table.Add(row);
        }
        r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
        r["range"] = address;
        r["rows"] = rows;
        r["cols"] = cols;
        List<object> dateCols = new List<object>();
        for (int j = 1; j <= cols; j++) if (dated[j]) dateCols.Add(ColLetter((int)range.Column + j - 1));
        if (dateCols.Count > 0) r["dateColumns"] = dateCols;
        r["values"] = table;
        return r;
    }

    private static object[,] ReadValues0(object raw)
    {
        if (raw == null || raw is string || raw is double || raw is long || raw is int || raw is bool)
        {
            object[,] one = new object[2, 2];
            one[1, 1] = raw;
            return one;
        }
        return (object[,])raw;
    }

    /// <summary>
    /// What is in each column of a block: the kinds of values, the range, the commonest texts and
    /// what looks wrong ("LOOK"). This is how a long table is understood without reading every row.
    /// A date is recognised by the column's display format, never by the value alone, and errors come
    /// back as "#DIV/0!" rather than as a negative number.
    /// </summary>
    private static Dictionary<string, object> ProfileOf(dynamic range, string address, object[,] values, int rows, int cols)
    {
        Dictionary<string, object> p = new Dictionary<string, object>();
        p["range"] = address;
        p["rows"] = rows;
        p["cols"] = cols;
        bool[] dated = DateColumnsOf(range, values, rows, cols);
        int col0 = (int)range.Column;

        // Row 1 is the header row; what is below it is data.
        int emptyRows = 0, duplicateRows = 0;
        Dictionary<string, int> whole = new Dictionary<string, int>();
        for (int i = 2; i <= rows; i++)
        {
            StringBuilder key = new StringBuilder();
            bool any = false;
            for (int j = 1; j <= cols; j++)
            {
                if (!IsBlankValue(values[i, j])) any = true;
                key.Append(TextOf(values[i, j])).Append('\u0001');
            }
            if (!any) { emptyRows++; continue; }
            string k = key.ToString();
            if (whole.ContainsKey(k)) duplicateRows++; else whole[k] = 1;
        }
        p["dataRows"] = Math.Max(0, rows - 1);
        if (emptyRows > 0) p["emptyRows"] = emptyRows;
        if (duplicateRows > 0) p["duplicateRows"] = duplicateRows;

        List<object> colList = new List<object>();
        for (int j = 1; j <= cols; j++)
        {
            Dictionary<string, object> c = new Dictionary<string, object>();
            int numbers = 0, texts = 0, blanks = 0, errors = 0, others = 0;
            int spaced = 0, numberLike = 0, dateLike = 0, negative = 0;
            double sum = 0, min = double.MaxValue, max = double.MinValue;
            List<double> nums = new List<double>();
            Dictionary<string, int> distinct = new Dictionary<string, int>();
            for (int i = 1; i <= rows; i++)
            {
                object v = values[i, j];
                if (v == null || (v is string && ((string)v).Length == 0)) { blanks++; continue; }
                if (ErrorText(v) != null) { errors++; continue; }
                if (v is double)
                {
                    double d = (double)v;
                    numbers++; nums.Add(d); sum += d;
                    if (d < min) min = d;
                    if (d > max) max = d;
                    if (d < 0) negative++;
                    int n;
                    string shownNumber = Plain(d);
                    distinct.TryGetValue(shownNumber, out n);
                    distinct[shownNumber] = n + 1;
                }
                else if (v is string)
                {
                    texts++;
                    string s = (string)v;
                    if (s.StartsWith("=", StringComparison.Ordinal)) others++; // formula text: no distinct
                    else
                    {
                        if (s != Tidy(s)) spaced++;
                        double n2; DateTime d2;
                        if (DateShaped.IsMatch(Tidy(s)) && ParseDate(s, DateTime.Now.Year, out d2)) dateLike++;
                        else if (ParseNumber(s, out n2)) numberLike++;
                        int n;
                        distinct.TryGetValue(s, out n);
                        distinct[s] = n + 1;
                    }
                }
                else others++;
            }
            c["col"] = j;
            c["letter"] = ColLetter(col0 + j - 1);
            string header = TextOf(values[1, j]).Trim();
            if (header.Length > 0) c["header"] = header;
            c["numbers"] = numbers; c["texts"] = texts; c["blanks"] = blanks; c["others"] = others;
            if (errors > 0) c["errors"] = errors;
            if (nums.Count > 0)
            {
                nums.Sort();
                if (dated[j])
                {
                    c["dates"] = numbers;
                    try
                    {
                        c["dateFrom"] = DateTime.FromOADate(min).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                        c["dateTo"] = DateTime.FromOADate(max).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    }
                    catch (Exception) { }
                    c["min"] = JsonCellValue(min);
                    c["max"] = JsonCellValue(max);
                }
                else
                {
                    double median = nums[0];
                    if (nums.Count % 2 == 1) median = nums[nums.Count / 2];
                    else if (nums.Count > 1) median = (nums[nums.Count / 2 - 1] + nums[nums.Count / 2]) / 2;
                    c["min"] = JsonCellValue(min);
                    c["median"] = JsonCellValue(median);
                    c["mean"] = Math.Round(sum / numbers, 2);
                    c["max"] = JsonCellValue(max);
                    c["sum"] = Math.Round(sum, 2);
                    if (negative > 0) c["negative"] = negative;
                    // Values far beyond the bulk of the column: more than three times the spread above
                    // the upper quartile, and only when they are a handful (a real long tail is not a fault).
                    if (nums.Count >= 20)
                    {
                        double q1 = nums[nums.Count / 4], q3 = nums[nums.Count * 3 / 4];
                        double fence = q3 + 3 * Math.Max(q3 - q1, 1e-9);
                        int far = 0;
                        foreach (double n3 in nums) if (n3 > fence) far++;
                        if (far > 0 && far <= nums.Count / 20)
                        {
                            Dictionary<string, object> outlier = new Dictionary<string, object>();
                            outlier["count"] = far;
                            outlier["over"] = Math.Round(fence, 2);
                            c["farAbove"] = outlier;
                        }
                    }
                }
            }
            if (distinct.Count > 0)
            {
                var top = distinct.OrderByDescending(kv => kv.Value).ThenBy(kv => kv.Key).Take(8)
                    .Select(kv => (object)(kv.Key + " x" + kv.Value)).ToList();
                c["topValues"] = top;
                c["distinct"] = distinct.Count;
                if (numbers + texts > distinct.Count) c["repeats"] = numbers + texts - distinct.Count;
            }
            // LOOK: the few things worth fixing before working with this column.
            List<object> look = new List<object>();
            if (spaced > 0) look.Add(spaced.ToString(CultureInfo.InvariantCulture) + " with stray spaces");
            if (numberLike > 0) look.Add(numberLike.ToString(CultureInfo.InvariantCulture) + " number(s) written as text");
            if (dateLike > 0) look.Add(dateLike.ToString(CultureInfo.InvariantCulture) + " date(s) written as text");
            if (numbers > 0 && texts > 0) look.Add("mixed numbers and text");
            if (errors > 0) look.Add(errors.ToString(CultureInfo.InvariantCulture) + " error value(s)");
            if (look.Count > 0) c["look"] = look;
            colList.Add(c);
        }
        p["columns"] = colList;

        // head and tail samples (5 rows each), same cell encoding as the small read
        int head = Math.Min(5, rows), tailRows = Math.Min(5, rows);
        Func<int, int, object> sampleCell = (i, j) =>
        {
            object v = values[i, j];
            if (dated[j] && IsNum(v))
            {
                try { return DateText(Convert.ToDouble(v, CultureInfo.InvariantCulture), true); }
                catch (Exception) { }
            }
            return JsonCellValue(v);
        };
        List<object> headList = new List<object>();
        for (int i = 1; i <= head; i++) { List<object> row = new List<object>(); for (int j = 1; j <= cols; j++) row.Add(sampleCell(i, j)); headList.Add(row); }
        p["sampleHead"] = headList;
        if (rows > head + tailRows)
        {
            List<object> tailList = new List<object>();
            for (int i = rows - tailRows + 1; i <= rows; i++) { List<object> row = new List<object>(); for (int j = 1; j <= cols; j++) row.Add(sampleCell(i, j)); tailList.Add(row); }
            p["sampleTail"] = tailList;
        }
        p["note"] = "large range: column profiles + head/tail samples returned; use -Range to read a specific block in full";
        return p;
    }

    // ------------------------------------------------------------ search / selection

    /// <summary>Find text in cell values (and optionally formulas) across sheets.</summary>
    private static object SearchBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string text = a.Need("text");
        bool inFormulas = true; // values and formulas are both searched, hits say which
        bool ignoreCase = !a.Flag("case", false);
        StringComparison cmp = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        int max = a.Int("max", 50);
        string onlySheet = a.Str("sheet", null);

        List<object> hits = new List<object>();
        int total = 0;
        int nSheets = 0;
        try { nSheets = (int)wb.Worksheets.Count; } catch (Exception) { }
        for (int s = 1; s <= nSheets && total < max; s++)
        {
            dynamic ws = null;
            try { ws = wb.Worksheets.Item(s); } catch (Exception) { continue; }
            string sheetName = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
            if (onlySheet != null && !string.Equals(sheetName, onlySheet, StringComparison.OrdinalIgnoreCase)) continue;
            dynamic used = null;
            try { used = ws.UsedRange; } catch (Exception) { continue; }
            int rows = (int)used.Rows.Count, cols = (int)used.Columns.Count;
            object[,] values = ReadValues(used);
            object[,] formulas = inFormulas ? ReadValues0(used.Formula) : null;
            for (int i = 1; i <= rows && total < max; i++)
                for (int j = 1; j <= cols && total < max; j++)
                {
                    string addr = null;
                    object v = values[i, j];
                    string vs = v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture);
                    bool hit = vs.IndexOf(text, cmp) >= 0;
                    string fs = null;
                    if (!hit && inFormulas)
                    {
                        object f = formulas[i, j];
                        fs = f == null ? null : Convert.ToString(f, CultureInfo.InvariantCulture);
                        if (fs != null && fs.StartsWith("=", StringComparison.Ordinal)) hit = fs.IndexOf(text, cmp) >= 0;
                    }
                    if (!hit) continue;
                    if (addr == null) addr = Convert.ToString(TryGet(() => (object)used.Cells[i, j].Address[false, false], ""), CultureInfo.InvariantCulture);
                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["sheet"] = sheetName;
                    one["cell"] = addr;
                    one["value"] = vs;
                    if (fs != null && fs.StartsWith("=", StringComparison.Ordinal)) one["formula"] = fs;
                    hits.Add(one);
                    total++;
                }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["text"] = text;
        r["hits"] = hits;
        r["count"] = hits.Count;
        if (hits.Count >= max) r["truncated"] = true;
        return r;
    }

    /// <summary>Read what the user currently has selected (the "change THIS part" gesture).</summary>
    private static object SelectionBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic app = App(false, false);
        dynamic sel = TryGet(() => (object)app.Selection, null);
        if (sel == null) throw new Fail("COM_ERROR", "nothing is selected");
        Dictionary<string, object> r = new Dictionary<string, object>();
        object addrProbe = TryGet(() => (object)sel.Address[false, false], null);
        if (addrProbe == null)
        {
            r["selectionType"] = "not a cell range";
            r["note"] = "the current selection is not a cell range";
            return r;
        }
        string sheetName = Convert.ToString(TryGet(() => (object)sel.Worksheet.Name, ""), CultureInfo.InvariantCulture);
        // only meaningful if the selection sits in the workbook we were asked about
        bool inBook = false;
        int nSheets = 0;
        try { nSheets = (int)wb.Worksheets.Count; } catch (Exception) { }
        for (int s = 1; s <= nSheets; s++)
        {
            dynamic ws = null;
            try { ws = wb.Worksheets.Item(s); } catch (Exception) { continue; }
            if (string.Equals(Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture), sheetName, StringComparison.OrdinalIgnoreCase)) { inBook = true; break; }
        }
        if (!inBook) throw new Fail("NOT_OPEN", "the selection is in another workbook (sheet \"" + sheetName + "\"). Ask the user to click into this workbook first.");
        string address = Convert.ToString(TryGet(() => (object)sel.Address[false, false], ""), CultureInfo.InvariantCulture);
        int rows = 0, cols = 0;
        try { rows = (int)sel.Rows.Count; cols = (int)sel.Columns.Count; } catch (Exception) { }
        r["sheet"] = sheetName;
        r["range"] = address;
        r["rows"] = rows;
        r["cols"] = cols;
        r["cells"] = rows * cols;
        if (rows > 0 && cols > 0 && rows * cols <= 400)
        {
            object[,] values = ReadValues(sel);
            List<object> table = new List<object>();
            for (int i = 1; i <= rows; i++)
            {
                List<object> row = new List<object>();
                for (int j = 1; j <= cols; j++) row.Add(JsonCellValue(values[i, j]));
                table.Add(row);
            }
            r["values"] = table;
        }
        else r["note"] = "selection larger than 400 cells: values omitted, use read -Range " + sheetName + "!" + address;
        return r;
    }

    // ------------------------------------------------------------ insert / delete rows & cols

    private static string ColLetter(int index)
    {
        string s = "";
        while (index > 0) { int rem = (index - 1) % 26; s = (char)('A' + rem) + s; index = (index - 1) / 26; }
        return s;
    }

    private static object RowsCols(Bag a, bool rowsMode)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "The workbook is read-only; structural edits are refused.");
        string action = a.Str("action", "insert").ToLowerInvariant();
        if (action != "insert" && action != "delete") throw new Fail("BAD_ARGS", "action must be insert or delete");
        int index = a.Int("index", 0);
        if (index < 1) throw new Fail("BAD_ARGS", "-Index <1-based position> is required");
        int count = a.Int("count", 1);
        if (count < 1) throw new Fail("BAD_ARGS", "-Count must be >= 1");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");
        string what = rowsMode ? "row" : "column";

        string span = rowsMode ? index + ":" + (index + count - 1) : ColLetter(index) + ":" + ColLetter(index + count - 1);
        dynamic block = rowsMode ? ws.Rows(span) : ws.Columns(span);
        PerfState perf = PerfOn();
        try
        {
            if (action == "insert") block.Insert();
            else block.Delete();
        }
        finally { PerfOff(perf); }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
        r["action"] = action;
        r["what"] = what;
        r["at"] = index;
        r["count"] = count;
        r["note"] = "structure changed - existing formulas were adjusted by the app; snapshot before this kind of edit makes rollback possible";
        return r;
    }

    // ------------------------------------------------------------ chart

    private static int ChartTypeOf(string t)
    {
        switch (t)
        {
            case "column": case "combo": return 51;   // xlColumnClustered (combo starts from columns)
            case "bar": return 57;                    // xlBarClustered
            case "stacked": case "stackedcolumn": return 52;
            case "stackedbar": return 58;
            case "line": return 4;                    // xlLine
            case "smooth": case "smoothline": return 4;
            case "pie": return 5;                     // xlPie
            case "doughnut": case "donut": return -4120;
            case "scatter": return -4169;             // xlXYScatter
            case "scatterline": return 74;
            case "area": return 76;                   // xlArea
            case "radar": return -4151;
            default: throw new Fail("BAD_ARGS", "chart type must be column|bar|stacked|stackedbar|line|smooth|pie|doughnut|scatter|scatterline|area|radar|combo, got \"" + t + "\"");
        }
    }

    private static string ChartTypeName(object t)
    {
        // WPS hands back the chart type as int (some builds as double), so accept any numeric
        // IConvertible: only the double branch used to be mapped, which made `chart -Action list`
        // report raw codes ("4", "51") while `chart -Action add` reported names ("line", "column").
        int v;
        if (t != null && !(t is string) && t is IConvertible)
        {
            try { v = Convert.ToInt32(t, CultureInfo.InvariantCulture); }
            catch (Exception) { return Convert.ToString(t, CultureInfo.InvariantCulture); }
            if (v == 51) return "column";
            if (v == 57) return "bar";
            if (v == 52) return "stacked";
            if (v == 58) return "stackedbar";
            if (v == 4 || v == 65) return "line";
            if (v == 5) return "pie";
            if (v == -4120) return "doughnut";
            if (v == -4169 || v == 74) return "scatter";
            if (v == 76) return "area";
            if (v == -4151) return "radar";
            if (v == -4098) return "unknown";
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }
        return Convert.ToString(t, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// The title of a chart. WPS turns the call down while it is still laying out a chart whose series
    /// were just moved to another type or axis: it is asked again, and then by the other way there is
    /// to ask (SetElement). Same retry as dsh-office OfficeHelper.cs:4736-4753.
    /// </summary>
    private static void Titled(dynamic chart, string title)
    {
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                if (attempt == 2) chart.SetElement(2); else chart.HasTitle = true;
                chart.ChartTitle.Text = title;
                return;
            }
            catch (Exception)
            {
                if (attempt >= 3) throw;
                System.Threading.Thread.Sleep(250);
            }
        }
    }

    /// <summary>The secondary value axis, asked for more than once: WPS returns null on the first ask.</summary>
    private static dynamic SecondaryAxis(dynamic chart)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                dynamic axis = chart.Axes(2, 2);
                if (axis != null) return axis;
            }
            catch (Exception) { }
            System.Threading.Thread.Sleep(150);
        }
        return null;
    }

    // Geometry of one ChartObject, in sheet points. Shared by chart list / move / add so the
    // agent can see where a chart actually landed instead of guessing from a screenshot.
    private static Dictionary<string, object> ChartGeometry(dynamic co)
    {
        Dictionary<string, object> g = new Dictionary<string, object>();
        double l = 0, t = 0, w = 0, h = 0;
        try { l = (double)co.Left; t = (double)co.Top; w = (double)co.Width; h = (double)co.Height; }
        catch (Exception) { }
        g["left"] = Math.Round(l, 1);
        g["top"] = Math.Round(t, 1);
        g["width"] = Math.Round(w, 1);
        g["height"] = Math.Round(h, 1);
        return g;
    }

    private static void ChartPlace(dynamic ws, dynamic co, string anchorRef)
    {
        if (anchorRef == null || co == null) return;
        dynamic anchor = ws.Range(anchorRef.Split(':')[0]);
        co.Left = (double)anchor.Left;
        co.Top = (double)anchor.Top;
    }

    private static object ChartCmd(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string action = a.Str("action", "add").ToLowerInvariant();
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");

        if (action == "list")
        {
            // -All scans every sheet: when a chart lands on the wrong sheet (the old wrapper
            // dropped -Sheet, reported 2026-11-01) this is the one command that still finds it.
            bool all = a.Flag("all", false);
            List<object> list = new List<object>();
            int sheetCount = 1;
            if (all) { try { sheetCount = (int)wb.Worksheets.Count; } catch (Exception) { sheetCount = 1; } }
            for (int s = 1; s <= sheetCount; s++)
            {
                dynamic one = ws;
                if (all)
                {
                    try { one = wb.Worksheets.Item(s); } catch (Exception) { continue; }
                }
                string sheetName = Convert.ToString(TryGet(() => (object)one.Name, ""), CultureInfo.InvariantCulture);
                int n = 0;
                try { n = (int)one.ChartObjects().Count; } catch (Exception) { }
                for (int k = 1; k <= n; k++)
                {
                    dynamic co = null;
                    try { co = one.ChartObjects().Item(k); } catch (Exception) { continue; }
                    Dictionary<string, object> e = new Dictionary<string, object>();
                    e["sheet"] = sheetName;
                    e["index"] = k;
                    e["name"] = Convert.ToString(TryGet(() => (object)co.Name, ""), CultureInfo.InvariantCulture);
                    e["type"] = ChartTypeName(TryGet(() => (object)co.Chart.ChartType, null));
                    e["title"] = Convert.ToString(TryGet(() => (object)co.Chart.ChartTitle.Text, ""), CultureInfo.InvariantCulture);
                    foreach (KeyValuePair<string, object> kv in ChartGeometry(co)) e[kv.Key] = kv.Value;
                    list.Add(e);
                }
            }
            return new Dictionary<string, object> {
                { "sheet", all ? "*" : Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture) },
                { "charts", list }, { "count", list.Count } };
        }
        if (action == "delete")
        {
            int idx = a.Int("index", 0);
            if (idx < 1) throw new Fail("BAD_ARGS", "chart -Action delete needs -Index <n> (see chart -Action list)");
            ws.ChartObjects().Item(idx).Delete();
            return new Dictionary<string, object> {
                { "sheet", Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture) },
                { "deleted", idx } };
        }
        if (action == "move")
        {
            int idx = a.Int("index", 0);
            if (idx < 1) throw new Fail("BAD_ARGS", "chart -Action move needs -Index <n> (see chart -Action list)");
            dynamic co = ws.ChartObjects().Item(idx);
            ChartPlace(ws, co, a.Str("anchor", null));
            if (a.Has("width")) { int w = a.Int("width", 0); if (w > 0) co.Width = (double)w; }
            if (a.Has("height")) { int h = a.Int("height", 0); if (h > 0) co.Height = (double)h; }
            Dictionary<string, object> r = new Dictionary<string, object>();
            r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
            r["index"] = idx;
            r["moved"] = true;
            foreach (KeyValuePair<string, object> kv in ChartGeometry(co)) r[kv.Key] = kv.Value;
            return r;
        }
        if (action == "add")
        {
            if (IsReadOnly(wb)) throw new Fail("READONLY", "The workbook is read-only; adding a chart is refused.");
            string dataRef = a.Need("range");
            string type = a.Str("type", "line").ToLowerInvariant();
            int ct = ChartTypeOf(type);
            dynamic data = ws.Range(dataRef);
            double left = 20, top = 20;
            string anchorRef = a.Str("anchor", null);
            if (anchorRef != null)
            {
                dynamic anchor = ws.Range(anchorRef.Split(':')[0]);
                left = (double)anchor.Left; top = (double)anchor.Top;
            }
            int boxW = a.Int("width", 480), boxH = a.Int("height", 300);
            if (boxW < 100) boxW = 100;
            if (boxH < 80) boxH = 80;
            dynamic co = ws.ChartObjects().Add(left, top, (double)boxW, (double)boxH);
            dynamic chart = co.Chart;
            chart.ChartType = ct;
            // Explicit series instead of SetSourceData: on WPS ET a bare single-column
            // SetSourceData guesses the plot direction wrong and explodes into one
            // series per ROW. Column 1 = x (optional), columns 2..N = one series each,
            // row 1 = series names. v1.16: previously only the LAST column became a series,
            // so a "Q-H family" table silently rendered as a single curve.
            int dataRows = (int)data.Rows.Count, dataCols = (int)data.Columns.Count;
            if (dataRows < 2) throw new Fail("BAD_ARGS", "chart -Range must include a header row and at least one data row");
            dynamic xRange = null;
            if (dataCols >= 2)
            {
                string x1 = Convert.ToString(data.Cells[2, 1].Address[false, false], CultureInfo.InvariantCulture);
                string x2 = Convert.ToString(data.Cells[dataRows, 1].Address[false, false], CultureInfo.InvariantCulture);
                xRange = ws.Range(x1 + ":" + x2);
            }
            int firstCol = dataCols >= 2 ? 2 : 1;
            List<object> seriesNames = new List<object>();
            for (int c = firstCol; c <= dataCols; c++)
            {
                string v1 = Convert.ToString(data.Cells[2, c].Address[false, false], CultureInfo.InvariantCulture);
                string v2 = Convert.ToString(data.Cells[dataRows, c].Address[false, false], CultureInfo.InvariantCulture);
                dynamic ser = chart.SeriesCollection().NewSeries();
                ser.Values = ws.Range(v1 + ":" + v2);
                if (xRange != null) { try { ser.XValues = xRange; } catch (Exception) { } }
                object h = TryGet(() => (object)data.Cells[1, c].Text, null);
                string sName = h == null ? "" : Convert.ToString(h, CultureInfo.InvariantCulture);
                if (sName.Length > 0) { try { ser.Name = sName; } catch (Exception) { } }
                seriesNames.Add(sName);
            }
            int nSeries = seriesNames.Count;

            // Per-series type and axis, AFTER every series exists: moving them earlier makes WPS
            // throw away the type it had worked out. combo = the last series becomes a line;
            // -Secondary moves the named ones (name or 1-based position) to the right-hand axis.
            List<object> secondary = a.List("secondary");
            bool[] toSecondary = new bool[nSeries + 1];
            List<object> secondaryApplied = new List<object>();
            List<object> secondaryUnknown = new List<object>();
            foreach (object raw in secondary)
            {
                string key = TextOf(raw).Trim();
                int idx = 0;
                bool found = false;
                if (int.TryParse(key, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx) && idx >= 1 && idx <= nSeries)
                {
                    toSecondary[idx] = true;
                    found = true;
                }
                else
                {
                    for (int k = 1; k <= nSeries; k++)
                    {
                        if (string.Equals(TextOf(seriesNames[k - 1]), key, StringComparison.OrdinalIgnoreCase))
                        {
                            toSecondary[k] = true;
                            found = true;
                            break;
                        }
                    }
                }
                if (found) secondaryApplied.Add(key); else secondaryUnknown.Add(key);
            }
            for (int k = 1; k <= nSeries; k++)
            {
                bool asLine = type == "combo" && (k == nSeries || toSecondary[k]);
                if (!asLine && !toSecondary[k]) continue;
                dynamic ser = chart.SeriesCollection(k);
                if (asLine) { try { ser.ChartType = 65; } catch (Exception) { } }
                if (toSecondary[k]) { try { ser.AxisGroup = 2; } catch (Exception) { } }
            }

            try { chart.HasLegend = seriesNames.Count > 1; } catch (Exception) { }
            string title = a.Str("title", null);
            if (title != null)
            {
                try
                {
                    Titled(chart, title);
                    try { chart.ChartTitle.Font.Size = 13; chart.ChartTitle.Font.Bold = true; } catch (Exception) { }
                }
                catch (Exception) { }
            }
            bool round = type == "pie" || type == "doughnut" || type == "donut" || type == "radar";
            string xTitle = a.Str("xtitle", null), yTitle = a.Str("ytitle", null), y2Title = a.Str("y2title", null);
            if (xTitle != null && !round) { try { chart.Axes(1).HasTitle = true; chart.Axes(1).AxisTitle.Text = xTitle; } catch (Exception) { } }
            if (yTitle != null && !round) { try { chart.Axes(2).HasTitle = true; chart.Axes(2).AxisTitle.Text = yTitle; } catch (Exception) { } }
            bool y2 = false;
            if (y2Title != null)
            {
                dynamic axis = SecondaryAxis(chart);
                if (axis != null)
                {
                    try { axis.HasTitle = true; axis.AxisTitle.Text = y2Title; y2 = true; } catch (Exception) { }
                }
            }
            string numberFormat = a.Str("numberFormat", a.Str("numformat", null));
            if (numberFormat != null) { try { chart.Axes(2).TickLabels.NumberFormat = numberFormat; } catch (Exception) { } }
            if (a.Has("min")) { double v; if (double.TryParse(a.Str("min", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out v)) { try { chart.Axes(2).MinimumScale = v; } catch (Exception) { } } }
            if (a.Has("max")) { double v; if (double.TryParse(a.Str("max", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out v)) { try { chart.Axes(2).MaximumScale = v; } catch (Exception) { } } }

            string labels = a.Str("labels", null);
            bool wantLabels = labels != null && !labels.Equals("false", StringComparison.OrdinalIgnoreCase) && !labels.Equals("none", StringComparison.OrdinalIgnoreCase) && labels != "0";
            if (labels == null && (type == "pie" || type == "doughnut" || type == "donut")) wantLabels = true;
            if (wantLabels)
            {
                try
                {
                    for (int k = 1; k <= nSeries; k++)
                    {
                        dynamic ser = chart.SeriesCollection(k);
                        ser.HasDataLabels = true;
                        if (labels != null && labels.Equals("percent", StringComparison.OrdinalIgnoreCase))
                        {
                            ser.DataLabels().ShowPercentage = true;
                            ser.DataLabels().ShowValue = false;
                            ser.DataLabels().ShowCategoryName = round && type != "radar";
                        }
                        try { ser.DataLabels().Font.Size = 9; } catch (Exception) { }
                    }
                }
                catch (Exception) { }
            }

            string legend = a.Str("legend", null);
            if (legend == null) legend = round ? "right" : nSeries > 1 ? "bottom" : "none";
            if (legend != null)
            {
                try
                {
                    chart.HasLegend = !legend.Equals("none", StringComparison.OrdinalIgnoreCase) && !legend.Equals("false", StringComparison.OrdinalIgnoreCase);
                    if (nSeries > 0 && legend.Length > 0) chart.Legend.Position = legend.StartsWith("r", StringComparison.OrdinalIgnoreCase) ? -4152
                        : legend.StartsWith("t", StringComparison.OrdinalIgnoreCase) ? -4160
                        : legend.StartsWith("l", StringComparison.OrdinalIgnoreCase) ? -4131 : -4107;
                }
                catch (Exception) { }
            }

            // A plain, readable look: one typeface, pale grid lines, bars that are not thin.
            try { chart.ChartArea.Font.Name = "\u5fae\u8f6f\u96c5\u9ed1"; chart.ChartArea.Font.Size = 10; } catch (Exception) { }
            try { chart.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = ColorFrom("#E4E6EA"); } catch (Exception) { }

            string chartName = a.Str("name", null);
            string resultName = chartName != null ? chartName : (title != null ? title : ("AI-" + type + "-" + DateTime.Now.ToString("HHmmss", CultureInfo.InvariantCulture)));
            try { co.Name = resultName; } catch (Exception) { }

            Dictionary<string, object> r = new Dictionary<string, object>();
            r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
            r["type"] = type;
            r["range"] = dataRef;
            r["title"] = title == null ? "" : title;
            r["name"] = Convert.ToString(TryGet(() => (object)co.Name, ""), CultureInfo.InvariantCulture);
            r["seriesCount"] = nSeries;
            r["seriesNames"] = seriesNames;
            r["secondary"] = secondaryApplied;
            if (secondaryUnknown.Count > 0) r["secondaryUnknown"] = secondaryUnknown;
            r["secondaryAxisTitle"] = y2;
            foreach (KeyValuePair<string, object> kv in ChartGeometry(co)) r[kv.Key] = kv.Value;
            // Where it sits on the sheet, so the next chart does not land on top of it.
            string occupancy = "";
            try { occupancy = Convert.ToString(co.TopLeftCell.Address[false, false], CultureInfo.InvariantCulture) + ":" + Convert.ToString(co.BottomRightCell.Address[false, false], CultureInfo.InvariantCulture); }
            catch (Exception) { }
            if (occupancy.Length > 0) r["occupancy"] = occupancy;
            // Charts lying over one another is the commonest fault of a dashboard: say so at once.
            List<object> overlaps = new List<object>();
            try
            {
                double l = (double)co.Left, t = (double)co.Top, rr2 = l + (double)co.Width, bb = t + (double)co.Height;
                int others = (int)ws.ChartObjects().Count;
                string mine = Convert.ToString(TryGet(() => (object)co.Name, ""), CultureInfo.InvariantCulture);
                for (int k = 1; k <= others; k++)
                {
                    dynamic other = null;
                    try { other = ws.ChartObjects().Item(k); } catch (Exception) { continue; }
                    string otherName = Convert.ToString(TryGet(() => (object)other.Name, ""), CultureInfo.InvariantCulture);
                    if (string.Equals(otherName, mine, StringComparison.Ordinal)) continue;
                    double ol = (double)other.Left, ot = (double)other.Top;
                    double or2 = ol + (double)other.Width, ob = ot + (double)other.Height;
                    if (l < or2 - 2 && ol < rr2 - 2 && t < ob - 2 && ot < bb - 2)
                    {
                        overlaps.Add(otherName + " (" + Convert.ToString(TryGet(() => (object)other.TopLeftCell.Address[false, false], ""), CultureInfo.InvariantCulture)
                            + ":" + Convert.ToString(TryGet(() => (object)other.BottomRightCell.Address[false, false], ""), CultureInfo.InvariantCulture) + ")");
                        break;
                    }
                }
            }
            catch (Exception) { }
            if (overlaps.Count > 0) r["overlaps"] = overlaps;
            if (nSeries == 0) r["note"] = "NOTHING is plotted: check the range";
            else if (overlaps.Count > 0) r["hint"] = "this chart OVERLAPS another one: move it with chart -Action move -Index n -Anchor <cell>";
            else r["hint"] = "render -IncludeCharts shows the cells plus every chart in one image; no desktop screenshot needed";
            return r;
        }
        throw new Fail("BAD_ARGS", "chart -Action must be add|list|delete|move");
    }

    // ------------------------------------------------------------ write

    private sealed class WriteTarget
    {
        public dynamic Ws;
        public dynamic Range;
        public string SheetName;
        public string WsName() { return Convert.ToString(TryGet(() => (object)Ws.Name, SheetName ?? ""), CultureInfo.InvariantCulture); }
        public string Address;
        public object NewValue;     // scalar or object[,] block
        public bool IsFormula;
        public string Expect;
        public object OldFormula;   // captured for rollback (Formula restores value+formula)
    }

    /// <summary>Capture the given sheets and store as a numbered batch snapshot. Cheap:
    /// only the sheets the batch touches, values+formulas in one bulk read each.</summary>
    private static Dictionary<string, object> BatchSnapshot(dynamic wb, string cmd, List<string> sheetNames)
    {
        HashSet<string> set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string s in sheetNames) if (!string.IsNullOrEmpty(s)) set.Add(s);
        if (set.Count == 0) set.Add("__none__");
        Dictionary<string, object> cap = CaptureBook(wb, set);
        cap.Remove("__none__");
        string dir = SnapDirOf(wb);
        string safe = "batch";
        int seq = 1;
        string seqFile = Path.Combine(dir, "seq.txt");
        try { seq = int.Parse(File.ReadAllText(seqFile).Trim(), CultureInfo.InvariantCulture) + 1; } catch (Exception) { seq = 1; }
        File.WriteAllText(seqFile, seq.ToString(CultureInfo.InvariantCulture), new UTF8Encoding(false));
        Dictionary<string, object> meta = new Dictionary<string, object>();
        meta["cmd"] = cmd;
        meta["seq"] = (long)seq;
        meta["sheets"] = sheetNames;
        cap["meta"] = meta;
        string path = Path.Combine(dir, "batch" + seq + ".json");
        File.WriteAllText(path, Json.Write(cap), new UTF8Encoding(false));
        int cells = 0;
        foreach (object s in (List<object>)cap["sheets"]) cells += ((Dictionary<string, object>)((Dictionary<string, object>)s)["cells"]).Count;
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["path"] = path;
        r["seq"] = (long)seq;
        r["cells"] = cells;
        return r;
    }

    private static object WriteCells(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "The workbook is read-only (opened ReadOnly or locked elsewhere); writing is refused.");

        List<object> rawItems = a.List("items");
        if (rawItems.Count == 0) throw new Fail("BAD_ARGS", "write needs at least one item");

        List<WriteTarget> targets = new List<WriteTarget>();
        foreach (object raw in rawItems)
        {
            Bag item = new Bag(raw);
            WriteTarget t = new WriteTarget();
            string refText = item.Str("cell", null) ?? item.Str("range", null);
            if (refText == null) throw new Fail("BAD_ARGS", "each write item needs \"cell\" (or \"range\")");
            string refSheet, address;
            SplitRef(refText, out refSheet, out address);
            string sheetName = refSheet != null ? refSheet : item.Str("sheet", null);
            if (sheetName == null) sheetName = a.Str("sheet", null);
            t.SheetName = sheetName;
            t.Address = address;
            t.Ws = sheetName == null ? (dynamic)ActiveSheetSafe(wb) : SheetByName(wb, sheetName);
            if (t.Ws == null) throw new Fail("COM_ERROR", "No active sheet in this workbook.");
            t.Range = t.Ws.Range(address);
            t.Expect = item.Str("expect", null);
            if (item.Has("formula")) { t.NewValue = item.Str("formula", ""); t.IsFormula = true; }
            else if (item.Has("values"))
            {
                List<object> rowsList = item.List("values");
                int nr = rowsList.Count, nc = 1;
                foreach (object rowObj in rowsList)
                {
                    List<object> row = rowObj as List<object>;
                    if (row != null && row.Count > nc) nc = row.Count;
                }
                object[,] block = new object[nr, nc]; // 0-based, exact size: what COM expects on write
                for (int i = 0; i < nr; i++)
                {
                    List<object> row = rowsList[i] as List<object>;
                    for (int j = 0; j < nc; j++)
                    {
                        object v = row != null && j < row.Count ? row[j] : null;
                        if (v is double) block[i, j] = (double)v;
                        else if (v is long) block[i, j] = (double)(long)v;
                        else if (v is bool) block[i, j] = v;
                        else if (v == null) block[i, j] = "";
                        else block[i, j] = Convert.ToString(v, CultureInfo.InvariantCulture);
                    }
                }
                int wantRows = (int)t.Range.Rows.Count, wantCols = (int)t.Range.Columns.Count;
                if (wantRows != nr || wantCols != nc)
                    throw new Fail("BAD_ARGS", "\"values\" block is " + nr + "x" + nc + " but range " + address + " is " + wantRows + "x" + wantCols);
                t.NewValue = block;
            }
            else if (item.Has("value"))
            {
                object v = item.Raw("value");
                if (v is double) t.NewValue = (double)v;
                else if (v is long) t.NewValue = (double)(long)v;
                else if (v is bool) t.NewValue = v;
                else if (v == null) t.NewValue = "";
                else t.NewValue = Convert.ToString(v, CultureInfo.InvariantCulture);
            }
            else throw new Fail("BAD_ARGS", "each write item needs \"value\", \"values\" or \"formula\"");
            targets.Add(t);
        }

    // Pass 1: validate everything before touching anything (anchor checks).
        foreach (WriteTarget t in targets)
        {
            if (t.Expect != null)
            {
                string current = Convert.ToString(TryGet(() => (object)t.Range.Text, "\u0001"), CultureInfo.InvariantCulture);
                if (current != t.Expect)
                    throw new Fail("ANCHOR_MISSING", "Cell " + (t.SheetName == null ? "" : t.SheetName) + "!" + t.Address + " currently shows \"" + Clean(current)
                        + "\", not the expected \"" + Clean(t.Expect) + "\". Read the cell again; if a human just changed it, re-plan.");
            }
        }

        // Auto pre-batch snapshot of the affected sheets: rollback must never depend on
        // the agent remembering to take one. The returned path also gives per-batch
        // attribution (diff -Since <this file> = exactly what this batch changed).
        Dictionary<string, object> batchSnap = BatchSnapshot(wb, "write", targets.Select(t => t.WsName()).Distinct().ToList());
        string batchSnapPath = Convert.ToString(batchSnap["path"], CultureInfo.InvariantCulture);

        // Capture old formulas for rollback, then write.
        foreach (WriteTarget t in targets)
        {
            try
            {
                int rows = (int)t.Range.Rows.Count, cols = (int)t.Range.Columns.Count;
                if (rows == 1 && cols == 1) t.OldFormula = t.Range.Formula;
                else t.OldFormula = t.Range.Formula;
            }
            catch (Exception) { t.OldFormula = null; }
        }

        List<object> applied = new List<object>();
        int written = 0;
        PerfState perf = PerfOn();
        try
        {
            foreach (WriteTarget t in targets)
            {
                object before;
                try { before = t.Range.Text; } catch (Exception) { before = null; }
                if (t.IsFormula) t.Range.Formula = (string)t.NewValue;
                else if (t.NewValue is object[,]) t.Range.Value2 = (object[,])t.NewValue;
                else t.Range.Value2 = t.NewValue;
                written++;
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["sheet"] = t.SheetName;
                one["cell"] = t.Address;
                one["before"] = before == null ? "" : Convert.ToString(before, CultureInfo.InvariantCulture);
                object afterObj;
                try { afterObj = t.Range.Text; } catch (Exception) { afterObj = null; }
                one["after"] = afterObj == null ? "" : Convert.ToString(afterObj, CultureInfo.InvariantCulture);
                if (t.IsFormula) one["formula"] = (string)t.NewValue;
                applied.Add(one);
            }
        }
        catch (Exception e)
        {
            // Roll back what this batch already wrote. Excel has no revision layer; restoring
            // the captured Formulas is the only way back (and it also kills undo history).
            int rolledBack = 0;
            for (int k = applied.Count - 1; k >= 0; k--)
            {
                WriteTarget t = targets[k];
                try { if (t.OldFormula != null) { t.Range.Formula = t.OldFormula; rolledBack++; } }
                catch (Exception) { }
            }
            throw new Fail("COM_ERROR", "write stopped after " + written + " cell(s): " + Clean(e.Message) + "; rolled back " + rolledBack + " cell(s) by restoring captured formulas");
        }
        finally { PerfOff(perf); }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["applied"] = applied;
        r["count"] = applied.Count;
        r["preSnapshot"] = batchSnap;
        r["undoNote"] = "spreadsheet edits cannot be undone with Ctrl+Z; pre-batch snapshot taken (see preSnapshot.path): diff -Since <path> shows exactly this batch's changes, restore -Since <path> undoes it";
        return r;
    }

    // ------------------------------------------------------------ perf guard
    // Writing thousands of cells one COM call at a time makes the app recalculate the whole
    // workbook after EVERY write. Suspending calculation (and screen refresh) for the batch
    // and recalculating once at the end turns minutes into seconds on big sheets.

    private sealed class PerfState { public object Calc; public object Screen; public bool CalcSuspended; public bool ScreenSuspended; }

    private static PerfState PerfOn()
    {
        PerfState st = new PerfState();
        dynamic app = App(false, false);
        if (app == null) return st;
        try { st.Calc = app.Calculation; app.Calculation = -4135; st.CalcSuspended = true; } catch (Exception) { } // xlCalculationManual
        try { st.Screen = app.ScreenUpdating; app.ScreenUpdating = false; st.ScreenSuspended = true; } catch (Exception) { }
        return st;
    }

    private static void PerfOff(PerfState st)
    {
        dynamic app = App(false, false);
        if (app == null) return;
        if (st.ScreenSuspended) { try { app.ScreenUpdating = st.Screen == null ? true : st.Screen; } catch (Exception) { } }
        if (st.CalcSuspended)
        {
            try { app.Calculate(); } catch (Exception) { }
            try { app.Calculation = st.Calc == null ? -4105 : st.Calc; } catch (Exception) { } // xlCalculationAutomatic
        }
    }

    // ------------------------------------------------------------ fill

    /// <summary>Fill one formula down a single-column range: write the top cell once, then
    /// native FillDown for the rest (one COM call, no per-row recalculation storm).</summary>
    private static object FillFormula(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "The workbook is read-only; filling is refused.");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");
        dynamic range = ws.Range(a.Need("range"));
        string formula = a.Need("formula");
        if (!formula.StartsWith("=", StringComparison.Ordinal))
            throw new Fail("BAD_ARGS", "\"formula\" must start with '='");
        int rows = (int)range.Rows.Count, cols = (int)range.Columns.Count;
        if (cols != 1) throw new Fail("BAD_ARGS", "fill supports a single column (got " + cols + "); write a values block for rectangles");
        if (rows < 2) throw new Fail("BAD_ARGS", "range needs at least 2 rows to fill");

        Dictionary<string, object> batchSnap = BatchSnapshot(wb, "fill", new List<string> { Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture) });
        PerfState perf = PerfOn();
        try
        {
            range.Cells[1, 1].Formula = formula;
            range.FillDown();
        }
        finally { PerfOff(perf); }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
        r["range"] = a.Str("range", "");
        r["rows"] = rows;
        r["formula"] = formula;
        r["preSnapshot"] = batchSnap;
        object topAfter = TryGet(() => (object)range.Cells[1, 1].Text, null);
        object bottomAfter = TryGet(() => (object)range.Cells[rows, 1].Text, null);
        r["topAfter"] = topAfter == null ? "" : Convert.ToString(topAfter, CultureInfo.InvariantCulture);
        r["bottomAfter"] = bottomAfter == null ? "" : Convert.ToString(bottomAfter, CultureInfo.InvariantCulture);
        // A formula that cannot compute shows an error straight away: say so instead of letting the
        // caller discover it later (first up to 20 cells, with the error code).
        List<object> errors = ErrorCells(range);
        r["errors"] = errors;
        if (errors.Count > 0) r["note"] = errors.Count.ToString(CultureInfo.InvariantCulture) + " cell(s) show an error: fix the formula or the data it reads";
        return r;
    }

    // ------------------------------------------------------------ sheet ops

    private static object SheetCmd(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string action = a.Str("action", "list").ToLowerInvariant();
        if (action == "list") return new Dictionary<string, object> { { "sheets", SheetsBrief(wb) } };
        if (action == "add")
        {
            string name = a.Need("name");
            dynamic ws = wb.Worksheets.Add();
            try { ws.Name = name; } catch (Exception e) { throw new Fail("BAD_ARGS", "Could not name the new sheet \"" + name + "\": " + Clean(e.Message)); }
            return new Dictionary<string, object> { { "added", name }, { "sheets", SheetsBrief(wb) } };
        }
        if (action == "rename")
        {
            string name = a.Need("name"), newName = a.Need("newName");
            SheetByName(wb, name).Name = newName;
            return new Dictionary<string, object> { { "renamed", name }, { "to", newName }, { "sheets", SheetsBrief(wb) } };
        }
        if (action == "delete")
        {
            string name = a.Need("name");
            if ((int)wb.Worksheets.Count <= 1) throw new Fail("BAD_ARGS", "cannot delete the last sheet of a workbook");
            dynamic ws = SheetByName(wb, name);
            // Deleting a NON-EMPTY sheet makes WPS raise a modal "are you sure" alert. A COM caller
            // cannot answer it, so the call blocks forever (found 2026-11-04: `sheet -Action delete`
            // on a non-empty sheet hung > 90 s until the test process was killed). Suppress alerts
            // for the duration of the delete only, and report how much content was discarded so the
            // caller can tell "deleted an empty scratch sheet" from "threw data away".
            long usedCells = 0;
            try
            {
                object cnt = TryGet(() => (object)ws.UsedRange.CountLarge, null);
                if (cnt == null) cnt = TryGet(() => (object)ws.UsedRange.Count, null);
                if (cnt != null) usedCells = Convert.ToInt64(cnt, CultureInfo.InvariantCulture);
            }
            catch (Exception) { }
            dynamic app = null;
            try { app = wb.Application; } catch (Exception) { }
            // Fall back to the cached application object: some WPS builds refuse .Application.
            if (app == null) { try { app = _app; } catch (Exception) { } }
            bool alertsSaved = false;
            bool alertsPrev = true;
            try
            {
                if (app != null)
                {
                    alertsPrev = (bool)app.DisplayAlerts;
                    app.DisplayAlerts = false;
                    alertsSaved = true;
                }
            }
            catch (Exception) { }
            try { ws.Delete(); }
            finally
            {
                if (alertsSaved) { try { app.DisplayAlerts = alertsPrev; } catch (Exception) { } }
            }
            return new Dictionary<string, object>
            {
                { "deleted", name },
                { "usedCells", usedCells },
                { "note", "the sheet is gone; a non-empty sheet is discarded without the WPS confirmation alert (alerts are suppressed for this one call)" },
                { "sheets", SheetsBrief(wb) }
            };
        }
        if (action == "activate")
        {
            // Needed because several WPS operations (chart add on older builds) follow the active
            // sheet, and because "show the user this sheet" had no primitive: agents used raw COM
            // GetActiveObject + Worksheets.Item(n).Activate() to get here (reported 2026-11-01).
            string name = a.Need("name");
            dynamic target = SheetByName(wb, name);
            try { target.Activate(); }
            catch (Exception e) { throw new Fail("COM_ERROR", "could not activate sheet \"" + name + "\": " + Clean(e.Message)); }
            return new Dictionary<string, object> { { "activated", name }, { "sheets", SheetsBrief(wb) } };
        }
        throw new Fail("BAD_ARGS", "sheet -Action must be list|add|rename|delete|activate");
    }

    // ------------------------------------------------------------ comments

    private static object CommentCmd(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string action = a.Str("action", "add").ToLowerInvariant();
        if (action == "list")
        {
            List<object> list = new List<object>();
            int nSheets = (int)wb.Worksheets.Count;
            for (int s = 1; s <= nSheets; s++)
            {
                dynamic ws = null;
                try { ws = wb.Worksheets.Item(s); } catch (Exception) { continue; }
                string sheetName = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
                int n = 0;
                try { n = (int)ws.Comments.Count; } catch (Exception) { continue; }
                for (int k = 1; k <= n; k++)
                {
                    dynamic cm = null;
                    try { cm = ws.Comments.Item(k); } catch (Exception) { continue; }
                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["sheet"] = sheetName;
                    one["cell"] = Convert.ToString(TryGet(() => (object)cm.Parent.Address[false, false], ""), CultureInfo.InvariantCulture);
                    one["author"] = Convert.ToString(TryGet(() => (object)cm.Author, ""), CultureInfo.InvariantCulture);
                    one["text"] = Convert.ToString(TryGet(() => (object)cm.Text(null, null, true), TryGet(() => (object)cm.Text(), "")), CultureInfo.InvariantCulture);
                    list.Add(one);
                }
            }
            return new Dictionary<string, object> { { "comments", list }, { "count", list.Count } };
        }
        if (action == "add")
        {
            string cell = a.Need("cell");
            string text = a.Need("text");
            dynamic ws = TargetSheet(wb, a);
            if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");
            dynamic range = ws.Range(cell);
            dynamic existing = TryGet(() => (object)range.Comment, null);
            if (existing != null) existing.Delete();
            range.AddComment(text);
            range.Comment.Visible = false;
            return new Dictionary<string, object> { { "added", cell }, { "sheet", Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture) } };
        }
        if (action == "delete")
        {
            if (a.Flag("all", false))
            {
                int deleted = 0;
                int nSheets = (int)wb.Worksheets.Count;
                for (int s = 1; s <= nSheets; s++)
                {
                    dynamic ws = null;
                    try { ws = wb.Worksheets.Item(s); } catch (Exception) { continue; }
                    int n = 0;
                    try { n = (int)ws.Comments.Count; } catch (Exception) { continue; }
                    while (n > 0) { try { ws.Comments.Item(1).Delete(); deleted++; } catch (Exception) { break; } try { n = (int)ws.Comments.Count; } catch (Exception) { break; } }
                }
                return new Dictionary<string, object> { { "deletedAll", deleted } };
            }
            throw new Fail("BAD_ARGS", "comment -Action delete needs -All in v0.1 (list first, then delete all or use write to clear)");
        }
        throw new Fail("BAD_ARGS", "comment -Action must be list|add|delete");
    }

    // ------------------------------------------------------------ render / autofit / format

    private static bool TryRender(dynamic ws, string address, string pngPath)
    {
        try { return TryRenderRange(ws.Range(address), pngPath); }
        catch (Exception) { return false; }
    }

    // CopyPicture one Range into a PNG. CopyPicture renders CELLS ONLY: floating chart objects
    // are not part of the range picture (verified 2026-11-01 on a sheet with five charts), which
    // is why render -IncludeCharts composites the chart images on top instead of trusting the
    // range copy. The helper writes the file itself (System.Drawing), so the PNG stays plaintext
    // even on machines where WPS-written files come out encrypted.
    private static bool TryRenderRange(dynamic range, string pngPath)
    {
        // The clipboard is a machine-wide resource shared with the PPT helper: serialise renders.
        IDisposable clip = WocLocks.Clipboard(15000);
        if (clip == null) throw new Fail("BUSY", "the system clipboard is busy (another render is running); retry in a moment");
        try
        {
            // v1.32: the copy below overwrites the machine-wide clipboard, so keep the user's and put
            // it back on the way out (helper/shared/ClipboardKeeper.cs). Sync() after each write of
            // ours: it tells the keeper which state is "ours", so a clipboard the user repopulated
            // meanwhile is left alone.
            using (ClipboardKeeper.Keep())
            {
                range.CopyPicture(1, 2); // xlScreen, xlBitmap
                ClipboardKeeper.Sync();
                if (!Clipboard.ContainsImage()) return false;
                using (Image img = Clipboard.GetImage())
                {
                    if (img == null) return false;
                    img.Save(pngPath, ImageFormat.Png);
                }
                Clipboard.Clear();
                ClipboardKeeper.Sync();
                return true;
            }
        }
        catch (Exception) { return false; }
        finally { clip.Dispose(); }
    }

    private static object RenderBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        // -Overview: every sheet as a thumbnail in one picture (a workbook read at a glance).
        if (a.Flag("overview", false))
        {
            string overviewDir = a.Need("out");
            Directory.CreateDirectory(overviewDir);
            return RenderOverviewSheets(wb, a, overviewDir, a.Int("width", 1100));
        }
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");
        bool withCharts = a.Flag("includeCharts", false);
        string requested = a.Str("range", null);
        if (requested == null || requested.Length == 0)
            requested = Convert.ToString(TryGet(() => (object)ws.UsedRange.Address[false, false], "A1"), CultureInfo.InvariantCulture);

        string outDir = a.Need("out");
        Directory.CreateDirectory(outDir);
        string sheetName = Convert.ToString(TryGet(() => (object)ws.Name, "Sheet"), CultureInfo.InvariantCulture);
        string safe = new string(sheetName.Where(ch => char.IsLetterOrDigit(ch) || ch == '-' || ch == '_').ToArray());
        if (safe.Length == 0) safe = "sheet";

        Dictionary<string, object> r = new Dictionary<string, object>();
        List<object> chartInfo = new List<object>();
        List<dynamic> picked = new List<dynamic>();
        string address = requested;
        dynamic baseRange = null;
        bool needBase = false;

        if (withCharts)
        {
            // -IncludeCharts grows the view to cover the requested range plus EVERY floating chart
            // on this sheet. Why expanding is enough: Range.CopyPicture(xlScreen, xlBitmap) paints
            // the on-screen appearance of the range, so a chart that overlaps it is ALREADY in the
            // picture - measured 2026-11-02 on this machine: A1:H52 (overlaps two charts) rendered
            // both, A1:H10 (no overlap) rendered none. What a plain render never shows is a chart
            // parked outside the range (e.g. tables in A..D, charts in J onward). Drawing the charts
            // a second time on top would duplicate them, so this path only expands the range.
            dynamic range = ws.Range(requested);
            double x1 = (double)range.Left, y1 = (double)range.Top;
            double x2 = x1 + (double)range.Width, y2 = y1 + (double)range.Height;
            int tlRow = (int)range.Cells[1, 1].Row, tlCol = (int)range.Cells[1, 1].Column;
            dynamic brc0 = range.Cells[(int)range.Rows.Count, (int)range.Columns.Count];
            int brRow = (int)brc0.Row, brCol = (int)brc0.Column;
            int n = 0;
            try { n = (int)ws.ChartObjects().Count; } catch (Exception) { }
            for (int k = 1; k <= n; k++)
            {
                dynamic co = null;
                try { co = ws.ChartObjects().Item(k); } catch (Exception) { continue; }
                double cl, ct, cw, chh;
                try
                {
                    cl = (double)co.Left; ct = (double)co.Top;
                    cw = (double)co.Width; chh = (double)co.Height;
                }
                catch (Exception) { continue; }
                picked.Add(co);
                if (cl < x1) x1 = cl;
                if (ct < y1) y1 = ct;
                if (cl + cw > x2) x2 = cl + cw;
                if (ct + chh > y2) y2 = ct + chh;
                try
                {
                    dynamic tlc = co.TopLeftCell;
                    if ((int)tlc.Row < tlRow) tlRow = (int)tlc.Row;
                    if ((int)tlc.Column < tlCol) tlCol = (int)tlc.Column;
                    dynamic brc = co.BottomRightCell;
                    if ((int)brc.Row > brRow) brRow = (int)brc.Row;
                    if ((int)brc.Column > brCol) brCol = (int)brc.Column;
                }
                catch (Exception) { }
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["index"] = k;
                one["name"] = Convert.ToString(TryGet(() => (object)co.Name, ""), CultureInfo.InvariantCulture);
                one["title"] = Convert.ToString(TryGet(() => (object)co.Chart.ChartTitle.Text, ""), CultureInfo.InvariantCulture);
                one["left"] = Math.Round(cl, 1);
                one["top"] = Math.Round(ct, 1);
                chartInfo.Add(one);
            }
            baseRange = ws.Range(ws.Cells[tlRow, tlCol], ws.Cells[brRow, brCol]);
            address = Convert.ToString(baseRange.Address[false, false], CultureInfo.InvariantCulture);
            needBase = true;
            r["chartsIncluded"] = picked.Count;
            if (picked.Count == 0)
                r["note"] = "no floating charts on this sheet; the plain range image already shows everything";
            else
            {
                r["charts"] = chartInfo;
                r["hint"] = "one image now covers the cells and every chart on this sheet; do not take a desktop screenshot to cross-check";
            }
        }

        string png = Path.Combine(outDir, safe + "-" + address.Replace(":", "_").Replace("!", "_") + ".png");
        bool ok = needBase ? TryRenderRange(baseRange, png) : TryRender(ws, address, png);

        if (!ok)
            throw new Fail("COM_ERROR", "CopyPicture produced no image (range " + address + "). The range may be empty, or WPS refused the copy.");

        // rescale to the requested width (dispose the source before touching its file)
        int wantWidth = a.Int("width", 1100);
        int srcW = 0, srcH = 0;
        try { using (Image img = Image.FromFile(png)) { srcW = img.Width; srcH = img.Height; } }
        catch (Exception) { }
        if (srcW > 0 && wantWidth >= 100 && Math.Abs(srcW - wantWidth) > 2)
        {
            try
            {
                int h = (int)Math.Round((double)srcH * wantWidth / srcW);
                string scaledPath = Path.Combine(Path.GetDirectoryName(png), Path.GetFileNameWithoutExtension(png) + "-w.png");
                using (Image img = Image.FromFile(png))
                using (Bitmap scaled = new Bitmap(img, wantWidth, h))
                {
                    scaled.Save(scaledPath, ImageFormat.Png);
                }
                File.Delete(png);
                png = scaledPath;
            }
            catch (Exception) { }
        }
        try { using (Image img = Image.FromFile(png)) { r["width"] = img.Width; r["height"] = img.Height; } }
        catch (Exception) { }
        r["sheet"] = sheetName;
        r["range"] = address;
        if (withCharts && requested != address) r["requestedRange"] = requested;
        r["png"] = png;
        return r;
    }

    private static object AutofitCmd(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");
        string address = a.Str("range", null);
        dynamic range = address == null ? (dynamic)ws.UsedRange : ws.Range(address);
        int beforeCols = (int)range.Columns.Count;
        range.Columns.AutoFit();
        range.Rows.AutoFit();
        return new Dictionary<string, object> { { "sheet", Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture) }, { "range", address == null ? "usedRange" : address } };
    }

    private static double ParsePoints(string s, string what)
    {
        string t = s.Trim().ToLowerInvariant();
        double v;
        if (t.EndsWith("pt")) t = t.Substring(0, t.Length - 2).Trim();
        if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v))
            throw new Fail("BAD_ARGS", "cannot parse " + what + " \"" + s + "\"");
        return v;
    }

    // COM wants the colour packed as BGR in the low 24 bits (value = R + G*256 + B*65536), not as
    // the ARGB int System.Drawing produces. Getting this backwards silently swaps red and blue:
    // `-Color "#FF0000"` rendered BLUE until 2026-11-02 (found while testing the format batch form;
    // the Word helper already packed BGR correctly - WpsHelper.cs "r + g * 256 + b * 65536").
    private static int ColorFrom(string s)
    {
        string t = s.Trim().TrimStart('#');
        if (t.Length == 6)
        {
            int v;
            if (int.TryParse(t, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v))
            {
                int r = (v >> 16) & 0xFF, g = (v >> 8) & 0xFF, b = v & 0xFF;
                return r | (g << 8) | (b << 16);
            }
        }
        throw new Fail("BAD_ARGS", "color must look like #C00000, got \"" + s + "\"");
    }

    // Apply the format options to one Range. Shared by the single-range and the multi-range
    // (-Range "A1:E1;A5:E5;A9:E9") paths, so a 26-step formatting pass becomes one command.
    private static List<object> ApplyFormat(dynamic range, Bag a)
    {
        List<object> done = new List<object>();
        if (a.Has("font")) { range.Font.Name = a.Str("font", ""); done.Add("font"); }
        if (a.Has("size")) { range.Font.Size = ParsePoints(a.Str("size", ""), "size"); done.Add("size"); }
        if (a.Has("bold")) { range.Font.Bold = a.Flag("bold", true) ? 1 : 0; done.Add("bold"); }
        if (a.Has("italic")) { range.Font.Italic = a.Flag("italic", true) ? 1 : 0; done.Add("italic"); }
        if (a.Has("underline")) { range.Font.Underline = a.Flag("underline", true) ? 2 : -4142; done.Add("underline"); }
        if (a.Has("color")) { range.Font.Color = ColorFrom(a.Str("color", "")); done.Add("color"); }
        if (a.Has("fillColor"))
        {
            string fc = a.Str("fillColor", "").Trim();
            if (fc.Equals("none", StringComparison.OrdinalIgnoreCase) || fc == "\u65e0") range.Interior.ColorIndex = -4142; // xlNone
            else range.Interior.Color = ColorFrom(fc);
            done.Add("fillColor");
        }
        if (a.Has("halign"))
        {
            string h = a.Str("halign", "").ToLowerInvariant();
            int v = h == "left" ? -4131 : h == "center" ? -4108 : h == "right" ? -4152 : h == "justify" ? -4130 : 1 /*general*/;
            range.HorizontalAlignment = v; done.Add("halign");
        }
        if (a.Has("valign"))
        {
            string v0 = a.Str("valign", "").ToLowerInvariant();
            int v = v0 == "top" ? -4160 : v0 == "center" ? -4108 : v0 == "bottom" ? -4107 : -4108;
            range.VerticalAlignment = v; done.Add("valign");
        }
        if (a.Has("wrap")) { range.WrapText = a.Flag("wrap", true) ? 1 : 0; done.Add("wrap"); }
        if (a.Has("numberFormat")) { range.NumberFormat = a.Str("numberFormat", "General"); done.Add("numberFormat"); }
        if (a.Has("colWidth")) { range.ColumnWidth = double.Parse(a.Str("colWidth", "").Trim(), CultureInfo.InvariantCulture); done.Add("colWidth"); }
        if (a.Has("rowHeight")) { range.RowHeight = ParsePoints(a.Str("rowHeight", ""), "rowHeight"); done.Add("rowHeight"); }
        if (a.Has("borders"))
        {
            string b = a.Str("borders", "").ToLowerInvariant();
            // XlBordersIndex 5 = xlDiagonalDown, 6 = xlDiagonalUp. Styling those together with the
            // real edges (7..12) paints one huge X across the whole range instead of a grid
            // (observed 2026-11-01 on a 4x9 table). Diagonals are always cleared, never drawn.
            for (int edge = 5; edge <= 12; edge++)
            {
                try
                {
                    dynamic bd = range.Borders.Item(edge);
                    if (edge == 5 || edge == 6) { bd.LineStyle = -4142; continue; } // xlLineStyleNone
                    if (b == "none") { bd.LineStyle = -4142; }
                    else { bd.LineStyle = 1; bd.Weight = b == "thick" ? 4 : 2; } // xlContinuous
                }
                catch (Exception) { }
            }
            done.Add("borders");
        }
        return done;
    }

    private static object FormatRange(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");
        // Batch form (v1.18): {"items":[{"range":"A1:E1","bold":true,"fillColor":"#D9E1F2"},
        //                                {"sheet":"Sheet2","range":"A5:E5","font":"黑体"}]}
        // Each item carries its own options, so one command can lay out a whole sheet with
        // per-block differences (the semicolon list below only supports ONE option set).
        List<object> batch = a.List("items");
        if (batch.Count > 0)
        {
            List<object> results = new List<object>();
            for (int i = 0; i < batch.Count; i++)
            {
                Bag one = new Bag(batch[i]);
                string rangeText0 = one.Need("range");
                dynamic sheet = ws;
                string sheetName = one.Str("sheet", null);
                if (sheetName != null) sheet = SheetByName(wb, sheetName);
                string[] chunks = rangeText0.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                List<object> applied = new List<object>();
                for (int c = 0; c < chunks.Length; c++)
                {
                    string part0 = chunks[c].Trim();
                    if (part0.Length == 0) continue;
                    List<object> done0 = ApplyFormat(sheet.Range(part0), one);
                    if (done0.Count == 0)
                        throw new Fail("BAD_ARGS", "format items[" + i + "] for \"" + part0 + "\" has no formatting option");
                    applied.Add(new Dictionary<string, object> { { "range", part0 }, { "applied", done0 } });
                }
                Dictionary<string, object> entry = new Dictionary<string, object>();
                entry["sheet"] = Convert.ToString(TryGet(() => (object)sheet.Name, ""), CultureInfo.InvariantCulture);
                entry["ranges"] = applied;
                results.Add(entry);
            }
            return new Dictionary<string, object> { { "items", results }, { "count", results.Count } };
        }
        string rangeText = a.Need("range");
        // Semicolon-separated list: one round trip formats many disjoint blocks (v1.16 - a
        // formatting pass used to cost one command per block, ~0.5-0.8 s each).
        string[] parts = rangeText.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        List<object> perRange = new List<object>();
        for (int i = 0; i < parts.Length; i++)
        {
            string part = parts[i].Trim();
            if (part.Length == 0) continue;
            dynamic range = ws.Range(part);
            List<object> done = ApplyFormat(range, a);
            if (done.Count == 0)
                throw new Fail("BAD_ARGS", "format needs at least one formatting option (font/size/bold/color/fillColor/halign/valign/numberFormat/colWidth/rowHeight/borders/wrap)");
            perRange.Add(new Dictionary<string, object> { { "range", part }, { "applied", done } });
        }
        if (perRange.Count == 0) throw new Fail("BAD_ARGS", "format needs -Range");
        if (perRange.Count == 1) return perRange[0];
        return new Dictionary<string, object> { { "ranges", perRange }, { "count", perRange.Count } };
    }

    // ------------------------------------------------------------ snapshot / diff

    private static string SnapDirOf(dynamic wb)
    {
        string full = Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture);
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate", "excel-snapshots", DocKey(full));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static string SnapPath(dynamic wb)
    {
        return Path.Combine(SnapDirOf(wb), "last.json");
    }

    // Stable per-document key (same idea as the Word side): two workbooks that share a file name
    // in different folders must never share a snapshot baseline (multi-task sessions hit this).
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

    private static Dictionary<string, object> CaptureBook(dynamic wb, HashSet<string> onlySheets)
    {
        Dictionary<string, object> cap = new Dictionary<string, object>();
        List<object> sheets = new List<object>();
        int n = 0;
        try { n = (int)wb.Worksheets.Count; } catch (Exception) { }
        for (int s = 1; s <= n; s++)
        {
            dynamic ws = null;
            try { ws = wb.Worksheets.Item(s); } catch (Exception) { continue; }
            string name = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
            if (onlySheets != null && !onlySheets.Contains(name.ToLowerInvariant())) continue;
            object used = TryGet(() => (object)ws.UsedRange.Address[false, false], null);
            if (used == null) continue;
            string address = Convert.ToString(used, CultureInfo.InvariantCulture);
            dynamic range = ws.UsedRange;
            int rows = (int)range.Rows.Count, cols = (int)range.Columns.Count;
            object[,] values = ReadValues(range);
            object[,] formulas = ReadValues0(range.Formula);
            Dictionary<string, object> cells = new Dictionary<string, object>();
            for (int i = 1; i <= rows; i++)
                for (int j = 1; j <= cols; j++)
                {
                    object v = values[i, j]; object f = formulas[i, j];
                    bool hasV = !(v == null || (v is string && ((string)v).Length == 0));
                    string fs = f == null ? null : Convert.ToString(f, CultureInfo.InvariantCulture);
                    bool hasF = fs != null && fs.StartsWith("=", StringComparison.Ordinal);
                    if (!hasV && !hasF) continue;
                    Dictionary<string, object> cell = new Dictionary<string, object>();
                    if (hasV) cell["v"] = JsonCellValue(v);
                    if (hasF) cell["f"] = fs;
                    cells[(string)range.Cells[i, j].Address[false, false]] = cell;
                }
            Dictionary<string, object> sheetCap = new Dictionary<string, object>();
            sheetCap["name"] = name;
            sheetCap["used"] = address;
            sheetCap["cells"] = cells;
            sheets.Add(sheetCap);
        }
        cap["sheets"] = sheets;
        cap["taken"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        return cap;
    }

    private static string CellKeyOf(string sheet, string cell) { return sheet + "!" + cell; }

    /// <summary>The sheet names a snapshot covers, lower-cased; null = every sheet.</summary>
    private static HashSet<string> SnapshotSheetSet(Dictionary<string, object> cap)
    {
        HashSet<string> set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (object s in (List<object>)cap["sheets"])
            set.Add((string)((Dictionary<string, object>)s)["name"]);
        return set;
    }

    private static object SnapshotBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        HashSet<string> only = null;
        string sheetList = a.Str("sheets", null);
        if (sheetList != null)
        {
            only = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string nm in sheetList.Split(',')) if (nm.Trim().Length > 0) only.Add(nm.Trim());
        }
        Dictionary<string, object> cap = CaptureBook(wb, only);
        string path = a.Str("out", null);
        if (path == null) path = SnapPath(wb);
        string json = Json.Write(cap);
        File.WriteAllText(path, json, new UTF8Encoding(false));
        int cellCount = 0;
        foreach (object s in (List<object>)cap["sheets"]) cellCount += ((Dictionary<string, object>)((Dictionary<string, object>)s)["cells"]).Count;
        return new Dictionary<string, object> { { "path", path }, { "cells", cellCount }, { "taken", cap["taken"] } };
    }

    private sealed class CellState
    {
        public string Sheet, Cell;
        public object V; public string F;
        public bool Present;
    }

    private static Dictionary<string, CellState> FlattenSnapshot(Dictionary<string, object> cap)
    {
        Dictionary<string, CellState> flat = new Dictionary<string, CellState>(StringComparer.OrdinalIgnoreCase);
        foreach (object s in (List<object>)cap["sheets"])
        {
            Dictionary<string, object> sheet = (Dictionary<string, object>)s;
            string sheetName = (string)sheet["name"];
            Dictionary<string, object> cells = (Dictionary<string, object>)sheet["cells"];
            foreach (KeyValuePair<string, object> kv in cells)
            {
                Dictionary<string, object> cell = (Dictionary<string, object>)kv.Value;
                CellState st = new CellState();
                st.Sheet = sheetName; st.Cell = kv.Key; st.Present = true;
                if (cell.ContainsKey("v")) st.V = cell["v"];
                if (cell.ContainsKey("f")) st.F = Convert.ToString(cell["f"], CultureInfo.InvariantCulture);
                flat[CellKeyOf(sheetName, kv.Key)] = st;
            }
        }
        return flat;
    }

    private static bool SameV(object a2, object b)
    {
        string x = a2 == null ? "" : Convert.ToString(a2, CultureInfo.InvariantCulture);
        string y = b == null ? "" : Convert.ToString(b, CultureInfo.InvariantCulture);
        return x == y;
    }

    private static object DiffBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string since = a.Str("since", "last");
        string path = since;
        if (since == "last") path = SnapPath(wb);
        if (!File.Exists(path))
            throw new Fail("BAD_ARGS", "snapshot not found: " + path + ". Take snapshot first (or pass -Since <file>).");
        Dictionary<string, object> capBefore = Json.Parse(File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
        if (capBefore == null) throw new Fail("BAD_ARGS", "snapshot file is not a JSON object: " + path);
        Dictionary<string, CellState> before = FlattenSnapshot(capBefore);

        // compare only the sheets the snapshot covers
        HashSet<string> snapSheets = SnapshotSheetSet(capBefore);
        Dictionary<string, object> cap = CaptureBook(wb, snapSheets);
        Dictionary<string, CellState> after = FlattenSnapshot(cap);

        List<object> changed = new List<object>();
        foreach (KeyValuePair<string, CellState> kv in after)
        {
            CellState was;
            bool existed = before.TryGetValue(kv.Key, out was);
            CellState now = kv.Value;
            if (!existed)
            {
                if (now.V == null || SameV(now.V, "")) continue; // empty new cell: not interesting
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["sheet"] = now.Sheet; one["cell"] = now.Cell; one["change"] = "added";
                one["after"] = now.V; if (now.F != null) one["formula"] = now.F;
                changed.Add(one);
            }
            else if (!SameV(was.V, now.V) || was.F != now.F)
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["sheet"] = now.Sheet; one["cell"] = now.Cell; one["change"] = "changed";
                one["before"] = was.V; one["after"] = now.V;
                if (now.F != null) one["formula"] = now.F;
                changed.Add(one);
            }
        }
        foreach (KeyValuePair<string, CellState> kv in before)
        {
            if (!after.ContainsKey(kv.Key))
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["sheet"] = kv.Value.Sheet; one["cell"] = kv.Value.Cell; one["change"] = "removed";
                one["before"] = kv.Value.V; if (kv.Value.F != null) one["formulaBefore"] = kv.Value.F;
                changed.Add(one);
            }
        }

        int max = a.Int("max", 200);
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["since"] = path;
        r["changes"] = changed.Count > max ? changed.GetRange(0, max) : changed;
        r["totalChanges"] = changed.Count;
        if (changed.Count > max) r["truncated"] = true;
        return r;
    }

    // ------------------------------------------------------------ restore

    /// <summary>Restore the workbook to a snapshot: rewrite every differing cell and (by
    /// default) clear cells that appeared after the snapshot. The workbook is not saved -
    /// the caller decides whether to save, close-without-save, or inspect with diff/render.</summary>
    private static object RestoreBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "The workbook is read-only; restoring is refused.");
        string since = a.Str("since", "last");
        string path = since;
        if (since == "last") path = SnapPath(wb);
        if (!File.Exists(path))
            throw new Fail("BAD_ARGS", "snapshot not found: " + path + ". Take snapshot first (or pass -Since <file>).");

        Dictionary<string, object> capBefore = Json.Parse(File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
        if (capBefore == null) throw new Fail("BAD_ARGS", "snapshot file is not a JSON object: " + path);
        Dictionary<string, CellState> before = FlattenSnapshot(capBefore);
        // restore must not touch sheets the snapshot does not cover
        HashSet<string> snapSheetsR = SnapshotSheetSet(capBefore);
        Dictionary<string, CellState> now = FlattenSnapshot(CaptureBook(wb, snapSheetsR));
        bool keepExtras = a.Flag("keepExtras", false);

        // resolve sheets once
        Dictionary<string, dynamic> sheetMap = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
        List<object> extraSheets = new List<object>();
        int nSheets = 0;
        try { nSheets = (int)wb.Worksheets.Count; } catch (Exception) { }
        for (int s = 1; s <= nSheets; s++)
        {
            dynamic ws = null;
            try { ws = wb.Worksheets.Item(s); } catch (Exception) { continue; }
            string nm = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
            sheetMap[nm] = ws;
        }
        foreach (KeyValuePair<string, dynamic> kv in sheetMap)
        {
            bool inSnap = false;
            foreach (object s in (List<object>)capBefore["sheets"])
                if (string.Equals((string)((Dictionary<string, object>)s)["name"], kv.Key, StringComparison.OrdinalIgnoreCase)) { inSnap = true; break; }
            if (!inSnap) extraSheets.Add(kv.Key);
        }

        // plan: list of {ws, addr, isFormula, payload}
        List<object[]> plan = new List<object[]>();
        List<object> skippedMissingSheet = new List<object>();
        // Selective mode: -Cells "Sheet1!C2,Sheet1!E2" restores ONLY those cells (from the
        // snapshot; a listed cell absent from the snapshot is cleared). Everything else -
        // including extras-clearing - is left alone, so human edits elsewhere survive.
        string cellsSpec = a.Str("cells", null);
        bool selective = cellsSpec != null && cellsSpec.Trim().Length > 0;
        int clearedSelective = 0;
        if (selective)
        {
            string defaultSheet = a.Str("sheet", null);
            foreach (string raw in cellsSpec.Split(','))
            {
                string entry = raw.Trim();
                if (entry.Length == 0) continue;
                string sh, addr;
                SplitRef(entry, out sh, out addr);
                if (sh == null) sh = defaultSheet;
                if (sh == null || addr == null) throw new Fail("BAD_ARGS", "cell entries must look like \"Sheet!C2\" (got \"" + entry + "\")");
                dynamic ws;
                if (!sheetMap.TryGetValue(sh, out ws)) { skippedMissingSheet.Add(CellKeyOf(sh, addr)); continue; }
                string key = CellKeyOf(sh, addr);
                CellState was;
                if (before.TryGetValue(key, out was))
                {
                    if (was.F != null) plan.Add(new object[] { ws, addr, true, was.F });
                    else plan.Add(new object[] { ws, addr, false, was.V });
                }
                else
                {
                    plan.Add(new object[] { ws, addr, false, "" });
                    clearedSelective++;
                }
            }
        }
        else
        foreach (KeyValuePair<string, CellState> kv in before)
        {
            CellState was = kv.Value;
            CellState cur;
            bool exists = now.TryGetValue(kv.Key, out cur);
            if (exists && SameV(was.V, cur.V) && was.F == cur.F) continue;
            dynamic ws;
            if (!sheetMap.TryGetValue(was.Sheet, out ws)) { skippedMissingSheet.Add(kv.Key); continue; }
            if (was.F != null) plan.Add(new object[] { ws, was.Cell, true, was.F });
            else plan.Add(new object[] { ws, was.Cell, false, was.V });
        }
        int cleared = 0;
        if (!keepExtras && !selective)
        {
            foreach (KeyValuePair<string, CellState> kv in now)
            {
                if (before.ContainsKey(kv.Key)) continue;
                if (kv.Value.V == null || SameV(kv.Value.V, "")) continue; // already empty
                dynamic ws;
                if (!sheetMap.TryGetValue(kv.Value.Sheet, out ws)) continue;
                plan.Add(new object[] { ws, kv.Value.Cell, false, "" });
                cleared++;
            }
        }

        if (plan.Count == 0)
        {
            Dictionary<string, object> none = new Dictionary<string, object>();
            none["restored"] = 0; none["cleared"] = 0;
            none["note"] = "workbook already matches the snapshot";
            return none;
        }

        // capture old formulas for rollback, then write all
        object[] old = new object[plan.Count];
        for (int k = 0; k < plan.Count; k++)
        {
            dynamic ws = (dynamic)plan[k][0]; string addr = (string)plan[k][1];
            try { old[k] = ws.Range(addr).Formula; } catch (Exception) { old[k] = null; }
        }
        int written = 0;
        PerfState perf = PerfOn();
        try
        {
            for (int k = 0; k < plan.Count; k++)
            {
                dynamic ws = (dynamic)plan[k][0]; string addr = (string)plan[k][1];
                dynamic range = ws.Range(addr);
                if ((bool)plan[k][2]) range.Formula = (string)plan[k][3];
                else
                {
                    object v = plan[k][3];
                    if (v is long) range.Value2 = (double)(long)v;
                    else if (v is double) range.Value2 = (double)v;
                    else if (v == null) range.Value2 = "";
                    else range.Value2 = Convert.ToString(v, CultureInfo.InvariantCulture);
                }
                written++;
            }
        }
        catch (Exception e)
        {
            int rolledBack = 0;
            for (int k = written - 1; k >= 0; k--)
            {
                dynamic ws = (dynamic)plan[k][0]; string addr = (string)plan[k][1];
                try { if (old[k] != null) { ws.Range(addr).Formula = old[k]; rolledBack++; } }
                catch (Exception) { }
            }
            throw new Fail("COM_ERROR", "restore stopped after " + written + " cell(s): " + Clean(e.Message) + "; rolled back " + rolledBack + " cell(s)");
        }
        finally { PerfOff(perf); }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["since"] = path;
        r["restored"] = plan.Count - cleared;
        r["cleared"] = cleared;
        if (selective) { r["selective"] = true; r["clearedListed"] = clearedSelective; }
        if (skippedMissingSheet.Count > 0) r["skippedMissingSheet"] = skippedMissingSheet;
        if (extraSheets.Count > 0) r["extraSheets"] = extraSheets;
        r["note"] = "workbook restored in memory only - run diff to verify, then save to persist (or close without -Save to discard the restore itself)";
        return r;
    }

    // ------------------------------------------------------------ misc

    private static object CalcBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic app = App(false, false);
        app.Calculate();
        return new Dictionary<string, object> { { "calculated", true } };
    }

    private static object SaveBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string asPath = a.Str("as", null);
        if (asPath == null)
        {
            string full = Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture);
            if (full.Length == 0) throw new Fail("BAD_ARGS", "this workbook has never been saved; use save -As <path>");
            wb.Save();
            return new Dictionary<string, object> { { "saved", full } };
        }
        if (asPath.ToLowerInvariant().EndsWith(".pdf", StringComparison.Ordinal))
        {
            // 0 = xlTypePDF, 0 = xlQualityStandard
            wb.ExportAsFixedFormat(0, asPath, 0, false, false, Type.Missing, Type.Missing, false, true);
            return new Dictionary<string, object> { { "exported", asPath } };
        }
        int format = 51; // xlOpenXMLWorkbook (.xlsx)
        string low = asPath.ToLowerInvariant();
        if (low.EndsWith(".xls", StringComparison.Ordinal)) format = 56;
        else if (low.EndsWith(".csv", StringComparison.Ordinal)) format = 6;
        else if (low.EndsWith(".xlsm", StringComparison.Ordinal)) format = 52;
        string oldFull = Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture);
        wb.SaveAs(asPath, format);
        // identity moves to the new path (quirk: SaveAs renames the workbook)
        MarkerRemove(oldFull);
        MarkerSet(FullPathOf(asPath));
        // Keep a takeover caption consistent with the new name instead of showing the old one.
        try
        {
            string cap = CaptionOf(wb);
            if (IsHinted(cap))
            {
                string keep = cap.StartsWith(HINT_READONLY, StringComparison.Ordinal) ? HINT_READONLY : HINT_TAG;
                TrySetCaption(wb, HintCaption(NameOf(wb), keep));
                // The intent registry is keyed by path, so a save-as must move the entry with the
                // workbook - otherwise the next command would fail to re-brand the renamed file.
                string newFull = Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture);
                if (!string.Equals(oldFull, newFull, StringComparison.OrdinalIgnoreCase))
                {
                    HintStateSet(oldFull, null);
                    HintStateSet(newFull, keep);
                }
            }
        }
        catch (Exception) { }
        // Repoint the tray icon at the new path (it retires itself when the file it watches is
        // no longer open, so a stale path after save-as would make it quit early).
        try
        {
            string cap2 = CaptionOf(wb);
            if (IsHinted(cap2))
                TrayDaemon.UpdateDoc("excel", Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture), cap2);
        }
        catch (Exception) { }
        return new Dictionary<string, object> { { "savedAs", asPath } };
    }

    private static object CloseBook(Bag a)
    {
        string full = FullPathOf(a.Need("doc"));
        dynamic wb = FindOpen(full);
        if (wb == null) throw new Fail("NOT_OPEN", "The workbook is not open: " + full);
        bool save = a.Flag("save", false);
        try { ClearHint(wb); } catch (Exception) { }
        wb.Close(save);
        // Marker is deliberately KEPT: once a workbook is closed, any later command must
        // fail with DOC_CLOSED instead of silently reopening it. Only an explicit open
        // re-attaches. (Convenience yields to the collaboration-safe semantics.)
        // Closing the last workbook kills the automation instance silently (same as Writer).
        dynamic app = App(false, false);
        int left = 0;
        try { left = (int)app.Workbooks.Count; } catch (Exception) { left = 0; }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["closed"] = full;
        r["workbooksLeft"] = left;
        if (left == 0)
        {
            try { app.Quit(); } catch (Exception) { }
            // v1.36: the automation instance is gone - drop the cached COM object instead of
            // keeping a dead one around (the next command re-attaches to whatever WPS is running,
            // or starts a new one).
            _app = null;
        }
        r["processes"] = ProcessCount("et");
        return r;
    }

    // ------------------------------------------------------------ hint
    // Same takeover hint as the Writer side (see WpsHelper.cs): window caption + status bar
    // + tray balloon, all pure view state - nothing is written into the workbook. Measured on
    // WPS 12.0 (2026-10-30): Workbook.ActiveWindow.Caption is writable and lands in the frame
    // title; Application.StatusBar is READ/WRITE here (unlike Writer), and setting it to
    // False hands the bar back to WPS.

    // WPS 12 shows open documents as WINDOW TABS which truncate a caption from the right
    // (~22 half-width cells), so a suffix tag vanishes on long workbook names (measured
    // 2026-10-30 on the Writer side; Spreadsheets uses the same frame/tab). The marker goes
    // in FRONT: "[AI] book.xlsx".
    private const string HINT_TAG = "\u3014AI\u63a5\u7ba1\u4e2d\u3015";                    // 〔AI接管中〕
    private const string HINT_READONLY = "\u3014AI\u63a5\u7ba1\u4e2d\u00b7\u53ea\u8bfb\u3015";   // 〔AI接管中·只读〕

    private static bool IsHinted(string caption)
    {
        if (caption == null) return false;
        return caption.StartsWith(HINT_TAG + " ", StringComparison.Ordinal)
            || caption.StartsWith(HINT_READONLY + " ", StringComparison.Ordinal);
    }

    private static string HintTagFor(dynamic wb, Bag a)
    {
        string tag = HintTag(a);
        // A read-only workbook refuses writes (READONLY); label the hint accordingly.
        if (tag == HINT_TAG && IsReadOnly(wb)) return HINT_READONLY;
        return tag;
    }

    private static string NameOf(dynamic wb)
    {
        return Convert.ToString(TryGet(() => (object)wb.Name, ""), CultureInfo.InvariantCulture);
    }

    private static string HintTag(Bag a)
    {
        string t = a.Str("tag", null);
        if (t == null) return HINT_TAG;
        return t.Trim();
    }

    /// <summary>Caption for a workbook: the marker in front, then the file name (or just the name when cleared).</summary>
    private static string HintCaption(string name, string tag)
    {
        if (tag == null || tag.Length == 0) return name;
        return tag + " " + name;
    }

    /// <summary>
    /// Caption of the workbook's own window. Measured on WPS 12.0 (2026-10-30): Workbook.
    /// ActiveWindow returns null here (it is only populated when the book IS the active one),
    /// while Workbook.Windows(1).Caption reads and writes reliably - so that is the primary
    /// path, with ActiveWindow only as a fallback.
    /// </summary>
    private static string CaptionOf(dynamic wb)
    {
        try { return Convert.ToString((object)wb.Windows.Item(1).Caption, CultureInfo.InvariantCulture); }
        catch (Exception) { }
        try { return Convert.ToString((object)wb.ActiveWindow.Caption, CultureInfo.InvariantCulture); }
        catch (Exception) { return null; }
    }

    private static bool TrySetCaption(dynamic wb, string caption)
    {
        try { wb.Windows.Item(1).Caption = caption; return true; }
        catch (Exception) { }
        try { wb.ActiveWindow.Caption = caption; return true; }
        catch (Exception) { return false; }
    }

    private static bool TrySetStatusBar(dynamic app, string text)
    {
        if (app == null) return false;
        try
        {
            if (text == null || text.Length == 0) app.StatusBar = false;   // hand the bar back to WPS
            else app.StatusBar = text;
            return true;
        }
        catch (Exception) { return false; }
    }

    // ------------------------------------------------------------ hint lease (v1.35)
    // The caption is the VISIBLE hint; this registry is the INTENT ("this workbook is under
    // takeover"). Keeping them apart is what makes the hint a lease: the resident helper removes
    // the caption when the agent goes quiet (shared/HintLease.cs) but leaves the intent alone, so
    // the next request re-applies it with EnsureHint - a paused session comes back branded, an
    // ended session leaves a clean window. Cleared only by hint -Clear / finish / close /
    // open -NoHint, i.e. when the collaboration is explicitly over.
    private static string HintStatePath()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate", "hint-state-excel.json");
    }

    private static string HintStateGet(string full)
    {
        if (full == null || full.Length == 0) return null;
        try
        {
            Dictionary<string, object> m = Json.Parse(File.ReadAllText(HintStatePath(), Encoding.UTF8)) as Dictionary<string, object>;
            if (m != null)
            {
                object v;
                if (m.TryGetValue(full.ToLowerInvariant(), out v) && v != null)
                {
                    string tag = Convert.ToString(v, CultureInfo.InvariantCulture);
                    if (tag != null && tag.Length > 0) return tag;
                }
            }
        }
        catch (Exception) { }
        return null;
    }

    private static void HintStateSet(string full, string tag)
    {
        if (full == null || full.Length == 0) return;
        IDisposable g = WocLocks.State(3000);
        try
        {
            Dictionary<string, object> m = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            try
            {
                Dictionary<string, object> old = Json.Parse(File.ReadAllText(HintStatePath(), Encoding.UTF8)) as Dictionary<string, object>;
                if (old != null) m = old;
            }
            catch (Exception) { }
            string key = full.ToLowerInvariant();
            if (tag == null || tag.Length == 0) m.Remove(key);
            else m[key] = tag;
            string dir = Path.GetDirectoryName(HintStatePath());
            if (dir != null) Directory.CreateDirectory(dir);
            string tmp = HintStatePath() + ".tmp";
            File.WriteAllText(tmp, Json.Write(m), new UTF8Encoding(false));
            File.Copy(tmp, HintStatePath(), true);
            try { File.Delete(tmp); } catch (Exception) { }
        }
        catch (Exception) { }
        finally { if (g != null) g.Dispose(); }
    }

    /// <summary>v1.36: repair a hidden spreadsheet window. WPS's own close path can crash and come
    /// back with the frame chrome up but the document window hidden (Application.Visible false), so
    /// the user stares at an empty grey window: "the file is stuck". Verified on 2026-11-05 - after
    /// that state, `app.Visible = true` + Activate brings the grid back at once. Every request that
    /// is allowed to touch the app repairs this before doing its work, so the next agent command
    /// un-sticks what the host broke (the user does not have to know the trick).</summary>
    private static void HealWindow()
    {
        try
        {
            if (HintLease.HintOff()) { HintLease.Trace("excel", "heal skipped (hints disabled)"); return; }   // v1.37 A/B switch: no automatic UI writes at all
            dynamic app = _app;
            if (app == null) return;
            bool visible;
            try { visible = (bool)app.Visible; } catch (Exception) { return; }
            if (visible) return;
            int n;
            try { n = (int)app.Workbooks.Count; } catch (Exception) { return; }
            if (n <= 0) return;
            try { app.Visible = true; } catch (Exception) { return; }
            try { app.ActiveWorkbook.Activate(); } catch (Exception) { }
            Console.Error.WriteLine("woc-helper: repaired a hidden spreadsheet window (v1.36)");
        }
        catch (Exception) { }
    }

    /// <summary>Re-apply the hint on a workbook that is still under takeover. Called before every
    /// command: WPS drops the caption when the user switches to another tab, and the lease release
    /// takes it off when the agent goes quiet - in both cases the mark belongs back on screen as
    /// soon as the agent touches the workbook again.</summary>
    private static void EnsureHint(Bag a)
    {
        try
        {
            if (HintLease.HintOff()) return;   // v1.37 master switch
            string doc = a.Str("doc", null);
            if (doc == null || doc.Length == 0) return;
            string full = FullPathOf(doc);
            string tag = HintStateGet(full);
            if (tag == null) return;
            dynamic app = App(false, false);
            if (app == null) return;
            dynamic wb = FindOpen(full);
            if (wb == null) return;
            if (IsHinted(CaptionOf(wb))) return;          // still branded: no COM write at all
            TrySetCaption(wb, HintCaption(NameOf(wb), tag));
        }
        catch (Exception) { }
    }

    /// <summary>Take the visible hint off every workbook this instance has open. Called by the
    /// resident helper when the agent's lease expires and on its way out; the intent registry is
    /// kept, so the next request puts the mark back (see EnsureHint).</summary>
    private static void ReleaseHints()
    {
        try { TrayDaemon.Stop("excel"); } catch (Exception) { }
        // v1.37 master switch: with the mark off there is nothing to take off, and this is exactly
        // the write that used to hit WPS at an arbitrary moment (possibly while the user closes).
        if (HintLease.HintOff())
        {
            HintLease.Trace("excel", "release skipped (hints disabled)");
            return;
        }
        HintLease.Trace("excel", "release start");
        dynamic app = null;
        try { app = App(false, false); } catch (Exception) { }
        if (app == null) return;
        int n = 0;
        try { n = (int)app.Workbooks.Count; } catch (Exception) { return; }
        for (int k = 1; k <= n; k++)
        {
            try
            {
                dynamic wb = app.Workbooks.Item(k);
                if (IsHinted(CaptionOf(wb))) TrySetCaption(wb, NameOf(wb));
            }
            catch (Exception) { }
        }
        try { TrySetStatusBar(app, ""); } catch (Exception) { }
        HintLease.Trace("excel", "release done");
    }

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

    private static Dictionary<string, object> HintOnTakeover(dynamic wb, Bag a)
    {
        // v1.37 master switch: see HintLease.HintOff.
        if (HintLease.HintOff())
        {
            Dictionary<string, object> off = new Dictionary<string, object>();
            off["disabled"] = true;
            off["captionSet"] = false;
            off["note"] = "Takeover hints are disabled (hint-off.txt / WOC_HINT_OFF, v1.37): nothing is written into WPS.";
            HintLease.Trace("excel", "open: hint skipped (disabled)");
            return off;
        }
        string tag = HintTagFor(wb, a);
        string name = NameOf(wb);
        string caption = HintCaption(name, tag);
        bool set = TrySetCaption(wb, caption);
        // Record the intent even when the caption could not be set (hidden window): the next
        // command's EnsureHint will retry, and the lease release needs to know what to clear.
        HintStateSet(Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture), tag);
        Dictionary<string, object> h = new Dictionary<string, object>();
        h["tag"] = tag;
        h["caption"] = caption;
        h["captionSet"] = set;
        if (!set)
            h["note"] = "The window caption could not be set (the workbook window is hidden)."
                + " Run `hint -Doc ... -Toast` once the window is up to announce the takeover to the user.";
        return h;
    }

    private static Dictionary<string, object> ClearHint(dynamic wb)
    {
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["trayStopped"] = TrayDaemon.Stop("excel");
        if (wb == null) return r;
        r["captionCleared"] = TrySetCaption(wb, NameOf(wb));
        r["statusBarCleared"] = TrySetStatusBar(App(false, false), "");
        // finish / close are an explicit end of the collaboration: drop the intent too, so a later
        // command does not re-brand a workbook this skill was told to let go.
        HintStateSet(Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture), null);
        return r;
    }

    private static object HintCmd(Bag a)
    {
        dynamic wb = RequireBook(a, false);
        dynamic app = App(false, false);
        bool clear = a.Flag("clear", false);
        string tag = HintTagFor(wb, a);
        string name = NameOf(wb);
        // v1.37 master switch: an explicit hint is still answered, but nothing is written into WPS.
        // -Clear is never blocked (taking an old brand off is always safe).
        if (!clear && HintLease.HintOff())
        {
            Dictionary<string, object> off = new Dictionary<string, object>();
            off["name"] = name;
            off["cleared"] = false;
            off["disabled"] = true;
            off["captionSet"] = false;
            off["statusBar"] = false;
            off["toast"] = false;
            off["note"] = "Takeover hints are disabled (hint-off.txt / WOC_HINT_OFF, v1.37): nothing is written into WPS.";
            HintLease.Trace("excel", "hint skipped (disabled)");
            return off;
        }
        string text = a.Str("text", null);
        if (text == null)
            text = clear
                ? "\u5df2\u7ed3\u675f\u534f\u4f5c\u300a" + name + "\u300b\uff1a\u7a97\u53e3\u6807\u9898\u4e0e\u72b6\u6001\u680f\u5df2\u8fd8\u539f\u3002"
                : "\u5df2\u63a5\u7ba1\u300a" + name + "\u300b\uff1a\u6539\u52a8\u4f1a\u76f4\u63a5\u5199\u5165\u8868\u683c\uff0c\u6211\u4f1a\u5148\u5b58\u5feb\u7167\uff0c\u4f60\u968f\u65f6\u53ef\u4ee5\u53eb\u6211\u56de\u6eda\u3002";

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = name;
        r["cleared"] = clear;

        string caption = HintCaption(name, clear ? "" : tag);
        bool captionSet = TrySetCaption(wb, caption);
        r["caption"] = caption;
        r["captionSet"] = captionSet;
        if (!clear) r["tag"] = tag;
        // The intent registry follows an explicit hint too: -Clear ends the takeover, a plain
        // hint (re)brands the workbook, so a later lease check knows what to release.
        HintStateSet(Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture), clear ? null : tag);

        string bar = a.Str("bar", null);
        if (bar == null) bar = clear ? "" : text;
        bool barSet = TrySetStatusBar(app, bar);
        r["statusBar"] = barSet;
        r["statusBarText"] = bar;

        bool quiet = a.Flag("quiet", false);
        bool toast = !quiet && a.Flag("toast", !clear);
        if (toast) toast = ShowBalloon("WPS \u8868\u683c AI \u52a9\u624b", text);  // WPS 表格 AI 助手
        r["toast"] = toast;

        bool sound = !quiet && a.Flag("sound", false);
        if (sound) sound = PlayAsterisk();
        r["sound"] = sound;

        // resident tray icon (opt-in): tooltip = "marker · current step", right-click restores.
        if (clear) r["trayStopped"] = TrayDaemon.Stop("excel");
        else if (a.Flag("tray", false))
            r["tray"] = TrayDaemon.Start("excel", Convert.ToString(TryGet(() => (object)wb.FullName, ""), CultureInfo.InvariantCulture),
                                         caption, text, a.Str("wrapper", null));
        r["processes"] = ProcessCount("et");
        return r;
    }

    // ------------------------------------------------------------ finish
    // Agent-agnostic completion signal: a status line inside the workbook (the shared
    // canvas, works no matter which agent drove it) + optional native tray balloon and
    // system sound. Deliberately no DSH/plugin dependency - this must work everywhere.

    private static object FinishCmd(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string note = a.Str("note", "");
        bool savedBefore = TryGet(() => (object)wb.Saved, true) is bool && (bool)TryGet(() => (object)wb.Saved, true);
        string timeText = DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        string statusLine = "\u2705 AI \u5df2\u5b8c\u6210 " + timeText                      // ✅ AI 已完成 ...
            + (note.Length > 0 ? " | " + note : "")
            + " | " + (savedBefore ? "\u4fdd\u5b58\u65f6\u70b9\uff1a\u5df2\u5b58" : "\u672a\u4fdd\u5b58");  // 保存时点：已存 / 未保存

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["statusLine"] = statusLine;

        // Status line is OPT-IN: only written when -Cell names a free cell (or -Force
        // overrides). Default finish touches NOTHING in the workbook - notification
        // happens via toast/sound, the workbook stays byte-for-byte as the agent left it.
        if (!a.Flag("quiet", false))
        {
            string cell = a.Str("cell", null);
            if (cell != null)
            {
                dynamic ws = TargetSheet(wb, a);
                if (ws == null) throw new Fail("COM_ERROR", "No target sheet for the status line.");
                dynamic range = ws.Range(cell);
                string cur = Convert.ToString(TryGet(() => (object)range.Text, ""), CultureInfo.InvariantCulture);
                if (cur.Length > 0 && !a.Flag("force", false))
                    throw new Fail("BAD_ARGS", "Status cell " + cell + " already holds \"" + Clean(cur)
                        + "\". Pass -Cell pointing at a free cell (e.g. just below the used range), or -Force to overwrite.");
                range.Value2 = statusLine;
                r["statusCell"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture) + "!" + cell;
            }
        }

        // tray balloon (renders as a toast on Win10/11). The helper is a console app:
        // pump messages for a few seconds so the icon lives long enough to show it.
        // Toast + sound are the DEFAULT completion signals now that the status line is
        // opt-in; -Quiet mutes both for silent runs.
        bool quiet = a.Flag("quiet", false);
        bool toast = !quiet && a.Flag("toast", true);
        if (toast) toast = ShowBalloon("WPS \u8868\u683c AI \u52a9\u624b", statusLine);  // WPS 表格 AI 助手
        r["toast"] = toast;

        bool sound = !quiet && a.Flag("sound", true);
        if (sound) sound = PlayAsterisk();
        r["sound"] = sound;

        // Undo the takeover hint so the window is not left branded "AI working".
        r["hintCleared"] = ClearHint(wb);
        return r;
    }

    private static object QuitAll(Bag a)
    {
        dynamic app = App(false, false);
        if (app == null) return new Dictionary<string, object> { { "processes", ProcessCount("et") } };
        try { app.Quit(); } catch (Exception) { }
        return new Dictionary<string, object> { { "processes", ProcessCount("et") } };
    }
}
