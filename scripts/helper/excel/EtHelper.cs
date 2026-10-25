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
            else sb.Append(e);
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

internal static class Program
{
    private const string WPS_ET = "KET.Application";
    private const string MS_EXCEL = "Excel.Application";

    private static dynamic _app;

    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        string requestPath = args != null && args.Length > 0 ? args[0] : null;
        string responsePath = args != null && args.Length > 1 ? args[1] : null;

        object response;
        try
        {
            if (requestPath == null) throw new Fail("BAD_ARGS", "usage: EtHelper.exe <request.json> <response.json>");
            Bag bag = new Bag(Json.Parse(File.ReadAllText(requestPath, Encoding.UTF8)));
            string cmd = bag.Str("cmd", "").Trim().ToLowerInvariant();
            if (cmd.Length == 0) throw new Fail("BAD_ARGS", "\"cmd\" is required");
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
                || msg.IndexOf("busy", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("\u5bf9\u8bdd\u6846", StringComparison.Ordinal) >= 0;
            response = Err(busy ? "BUSY" : "COM_ERROR", msg + " (0x" + e.HResult.ToString("X8", CultureInfo.InvariantCulture) + ")");
        }
        catch (Exception e)
        {
            response = Err("INTERNAL", e.GetType().Name + ": " + Clean(e.Message));
        }

        string text = Json.Write(response);
        if (responsePath != null) File.WriteAllText(responsePath, text, new UTF8Encoding(false));
        else Console.Out.Write(text);
        return 0;
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
            case "doctor": return Doctor(a);
            case "status": return Status(a);
            case "open": return OpenBook(a);
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
            case "finish": return FinishCmd(a);
            default: throw new Fail("BAD_ARGS", "unknown cmd \"" + cmd + "\"; use doctor|status|open|read|search|selection|write|fill|rows|cols|sheet|comment|render|chart|autofit|format|snapshot|diff|restore|calc|save|close|quit|finish");
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

    /// <summary>The spreadsheet app: the running WPS ET instance if there is one, otherwise a new one.</summary>
    private static dynamic App(bool create, bool visible)
    {
        if (_app != null) return _app;
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
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(MarkerPath()));
            File.WriteAllText(MarkerPath(), Json.Write(m), new UTF8Encoding(false));
        }
        catch (Exception) { }
    }

    private static void MarkerSet(string full)
    {
        Dictionary<string, object> m = MarkerLoad();
        m[full.ToLowerInvariant()] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        MarkerSave(m);
    }

    private static bool MarkerHas(string full)
    {
        return MarkerLoad().ContainsKey(full.ToLowerInvariant());
    }

    private static void MarkerRemove(string full)
    {
        Dictionary<string, object> m = MarkerLoad();
        if (m.Remove(full.ToLowerInvariant())) MarkerSave(m);
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
        dynamic active = TryGet(() => (object)wb.ActiveSheet, null);
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
    private static dynamic TargetSheet(dynamic wb, Bag a)
    {
        string name = a.Str("sheet", null);
        return name == null ? (dynamic)TryGet(() => (object)wb.ActiveSheet, null) : SheetByName(wb, name);
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
            dynamic ws = TryGet(() => (object)wb.ActiveSheet, null);
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
        if (existing != null) { wb = existing; how = "reused"; }
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
        MarkerSet(full);
        if (otherInstancePossible)
        {
            info["multiInstanceWarning"] = "ET \u8fd0\u884c\u4e2d\u6709 " + etBefore
                + " \u4e2a\u8fdb\u7a0b\uff0c\u4f46\u672c skill \u80fd\u63a5\u5230\u7684\u5b9e\u4f8b\u91cc\u6ca1\u6709\u8fd9\u4efd\u5de5\u4f5c\u7c3f\uff1a\u5b83\u53ef\u80fd\u6b63\u5f00\u5728\u53e6\u4e00\u4e2a\u5b9e\u4f8b\u91cc\uff0c\u7ee7\u7eed\u5199\u5165\u4f1a\u4e92\u76f8\u8986\u76d6\u3002";
        }
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

        dynamic ws = sheetName == null ? (dynamic)TryGet(() => (object)wb.ActiveSheet, null) : SheetByName(wb, sheetName);
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
            r["profile"] = ProfileOf(address, values, rows, cols);
            return r;
        }

        List<object> table = new List<object>();
        for (int i = 1; i <= rows; i++)
        {
            List<object> row = new List<object>();
            for (int j = 1; j <= cols; j++)
            {
                object f = formulas[i, j];
                string fs = f == null ? null : Convert.ToString(f, CultureInfo.InvariantCulture);
                if (fs != null && fs.StartsWith("=", StringComparison.Ordinal))
                {
                    Dictionary<string, object> cell = new Dictionary<string, object>();
                    cell["v"] = JsonCellValue(values[i, j]);
                    cell["f"] = fs;
                    row.Add(cell);
                }
                else row.Add(JsonCellValue(values[i, j]));
            }
            table.Add(row);
        }
        r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
        r["range"] = address;
        r["rows"] = rows;
        r["cols"] = cols;
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

    private static Dictionary<string, object> ProfileOf(string address, object[,] values, int rows, int cols)
    {
        Dictionary<string, object> p = new Dictionary<string, object>();
        p["range"] = address;
        p["rows"] = rows;
        p["cols"] = cols;
        List<object> colList = new List<object>();
        for (int j = 1; j <= cols; j++)
        {
            Dictionary<string, object> c = new Dictionary<string, object>();
            int numbers = 0, texts = 0, blanks = 0, others = 0;
            List<double> nums = new List<double>();
            Dictionary<string, int> distinct = new Dictionary<string, int>();
            for (int i = 1; i <= rows; i++)
            {
                object v = values[i, j];
                if (v == null || (v is string && ((string)v).Length == 0)) { blanks++; continue; }
                if (v is double)
                {
                    numbers++; nums.Add((double)v);
                }
                else if (v is string)
                {
                    texts++;
                    string s = (string)v;
                    if (s.StartsWith("=", StringComparison.Ordinal)) others++; // formula text: skip distinct
                    else { int n; distinct.TryGetValue(s, out n); distinct[s] = n + 1; }
                }
                else others++;
            }
            c["col"] = j;
            c["numbers"] = numbers; c["texts"] = texts; c["blanks"] = blanks; c["others"] = others;
            if (nums.Count > 0)
            {
                nums.Sort();
                c["min"] = JsonCellValue(nums[0]);
                c["median"] = JsonCellValue(nums[nums.Count / 2]);
                c["max"] = JsonCellValue(nums[nums.Count - 1]);
            }
            if (distinct.Count > 0)
            {
                var top = distinct.OrderByDescending(kv => kv.Value).ThenBy(kv => kv.Key).Take(8)
                    .Select(kv => (object)(kv.Key + " x" + kv.Value)).ToList();
                c["topValues"] = top;
                c["distinct"] = distinct.Count;
            }
            colList.Add(c);
        }
        p["columns"] = colList;

        // head and tail samples (5 rows each), same cell encoding as the small read
        int head = Math.Min(5, rows), tailRows = Math.Min(5, rows);
        Func<int, int, int, object> cellAt = (i, j, rr) =>
        {
            object v = values[i, j];
            return JsonCellValue(v);
        };
        List<object> headList = new List<object>();
        for (int i = 1; i <= head; i++) { List<object> row = new List<object>(); for (int j = 1; j <= cols; j++) row.Add(cellAt(i, j, rows)); headList.Add(row); }
        p["sampleHead"] = headList;
        if (rows > head + tailRows)
        {
            List<object> tailList = new List<object>();
            for (int i = rows - tailRows + 1; i <= rows; i++) { List<object> row = new List<object>(); for (int j = 1; j <= cols; j++) row.Add(cellAt(i, j, rows)); tailList.Add(row); }
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
            case "column": return 51;        // xlColumnClustered
            case "bar": return 57;           // xlBarClustered
            case "line": return 4;           // xlLine
            case "pie": return 5;            // xlPie
            case "scatter": return -4169;    // xlXYScatter
            case "area": return 76;          // xlArea
            default: throw new Fail("BAD_ARGS", "chart type must be column|bar|line|pie|scatter|area, got \"" + t + "\"");
        }
    }

    private static string ChartTypeName(object t)
    {
        if (t is double) { int v = (int)(double)t; if (v == 51) return "column"; if (v == 57) return "bar"; if (v == 4) return "line"; if (v == 5) return "pie"; if (v == -4169) return "scatter"; if (v == 76) return "area"; }
        return Convert.ToString(t, CultureInfo.InvariantCulture);
    }

    private static object ChartCmd(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        string action = a.Str("action", "add").ToLowerInvariant();
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");

        if (action == "list")
        {
            List<object> list = new List<object>();
            int n = 0;
            try { n = (int)ws.ChartObjects().Count; } catch (Exception) { }
            for (int k = 1; k <= n; k++)
            {
                dynamic co = null;
                try { co = ws.ChartObjects().Item(k); } catch (Exception) { continue; }
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["index"] = k;
                one["name"] = Convert.ToString(TryGet(() => (object)co.Name, ""), CultureInfo.InvariantCulture);
                one["type"] = ChartTypeName(TryGet(() => (object)co.Chart.ChartType, null));
                one["title"] = Convert.ToString(TryGet(() => (object)co.Chart.ChartTitle.Text, ""), CultureInfo.InvariantCulture);
                list.Add(one);
            }
            return new Dictionary<string, object> { { "sheet", Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture) }, { "charts", list }, { "count", list.Count } };
        }
        if (action == "delete")
        {
            int idx = a.Int("index", 0);
            if (idx < 1) throw new Fail("BAD_ARGS", "chart -Action delete needs -Index <n> (see chart -Action list)");
            ws.ChartObjects().Item(idx).Delete();
            return new Dictionary<string, object> { { "deleted", idx } };
        }
        if (action == "add")
        {
            if (IsReadOnly(wb)) throw new Fail("READONLY", "The workbook is read-only; adding a chart is refused.");
            string dataRef = a.Need("range");
            string type = a.Str("type", "line").ToLowerInvariant();
            int ct = ChartTypeOf(type);
            dynamic data = ws.Range(dataRef);
            string anchorRef = a.Str("anchor", null);
            double left = 20, top = 20;
            if (anchorRef != null)
            {
                dynamic anchor = ws.Range(anchorRef.Split(':')[0]);
                left = (double)anchor.Left; top = (double)anchor.Top;
            }
            dynamic co = ws.ChartObjects().Add(left, top, 480, 300);
            dynamic chart = co.Chart;
            chart.ChartType = ct;
            // Explicit series instead of SetSourceData: on WPS ET a bare single-column
            // SetSourceData guesses the plot direction wrong and explodes into one
            // series per ROW. Column 1 = x (optional), column 2 (or 1) = values, row 1 = header.
            int dataRows = (int)data.Rows.Count, dataCols = (int)data.Columns.Count;
            dynamic valuesRange, xRange = null; string seriesName = "";
            int vCol = dataCols >= 2 ? dataCols : 1; // values = LAST column (col 1 is the x/time axis)
            string v1 = Convert.ToString(data.Cells[2, vCol].Address[false, false], CultureInfo.InvariantCulture);
            string v2 = Convert.ToString(data.Cells[dataRows, vCol].Address[false, false], CultureInfo.InvariantCulture);
            valuesRange = ws.Range(v1 + ":" + v2);
            object h = TryGet(() => (object)data.Cells[1, vCol].Text, null);
            seriesName = h == null ? "" : Convert.ToString(h, CultureInfo.InvariantCulture);
            if (dataCols >= 2)
            {
                string x1 = Convert.ToString(data.Cells[2, 1].Address[false, false], CultureInfo.InvariantCulture);
                string x2 = Convert.ToString(data.Cells[dataRows, 1].Address[false, false], CultureInfo.InvariantCulture);
                xRange = ws.Range(x1 + ":" + x2);
            }
            dynamic ser = chart.SeriesCollection().NewSeries();
            ser.Values = valuesRange;
            if (xRange != null) { try { ser.XValues = xRange; } catch (Exception) { } }
            if (seriesName.Length > 0) { try { ser.Name = seriesName; } catch (Exception) { } }
            try { chart.HasLegend = false; } catch (Exception) { }
            string title = a.Str("title", null);
            if (title != null)
            {
                try { chart.HasTitle = true; chart.ChartTitle.Text = title; } catch (Exception) { }
            }
            try { co.Name = title != null ? title : ("AI-" + type + "-" + DateTime.Now.ToString("HHmmss", CultureInfo.InvariantCulture)); } catch (Exception) { }
            Dictionary<string, object> r = new Dictionary<string, object>();
            r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
            r["type"] = type;
            r["range"] = dataRef;
            r["title"] = title == null ? "" : title;
            r["name"] = Convert.ToString(TryGet(() => (object)co.Name, ""), CultureInfo.InvariantCulture);
            r["seriesCount"] = TryGet(() => (object)(int)chart.SeriesCollection().Count, 0);
            r["hint"] = "verify visually with render -Range <the area under the chart> before saving";
            return r;
        }
        throw new Fail("BAD_ARGS", "chart -Action must be add|list|delete");
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
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate", "excel-snapshots");
        Directory.CreateDirectory(dir);
        string name = Convert.ToString(TryGet(() => (object)wb.Name, "book"), CultureInfo.InvariantCulture);
        string safe = new string(name.Where(char.IsLetterOrDigit).ToArray());
        int seq = 1;
        string seqFile = Path.Combine(dir, safe + "-seq.txt");
        try { seq = int.Parse(File.ReadAllText(seqFile).Trim(), CultureInfo.InvariantCulture) + 1; } catch (Exception) { seq = 1; }
        File.WriteAllText(seqFile, seq.ToString(CultureInfo.InvariantCulture), new UTF8Encoding(false));
        Dictionary<string, object> meta = new Dictionary<string, object>();
        meta["cmd"] = cmd;
        meta["seq"] = (long)seq;
        meta["sheets"] = sheetNames;
        cap["meta"] = meta;
        string path = Path.Combine(dir, safe + "-batch" + seq + ".json");
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
            t.Ws = sheetName == null ? (dynamic)TryGet(() => (object)wb.ActiveSheet, null) : SheetByName(wb, sheetName);
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
            SheetByName(wb, name).Delete();
            return new Dictionary<string, object> { { "deleted", name }, { "sheets", SheetsBrief(wb) } };
        }
        throw new Fail("BAD_ARGS", "sheet -Action must be list|add|rename|delete");
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
        try
        {
            dynamic range = ws.Range(address);
            range.CopyPicture(1, 2); // xlScreen, xlBitmap
            if (!Clipboard.ContainsImage()) return false;
            using (Image img = Clipboard.GetImage())
            {
                if (img == null) return false;
                img.Save(pngPath, ImageFormat.Png);
            }
            Clipboard.Clear();
            return true;
        }
        catch (Exception) { return false; }
    }

    private static object RenderBook(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");
        string address = a.Str("range", null);
        if (address == null || address.Length == 0)
            address = Convert.ToString(TryGet(() => (object)ws.UsedRange.Address[false, false], "A1"), CultureInfo.InvariantCulture);

        string outDir = a.Need("out");
        Directory.CreateDirectory(outDir);
        string sheetName = Convert.ToString(TryGet(() => (object)ws.Name, "Sheet"), CultureInfo.InvariantCulture);
        string safe = new string(sheetName.Where(ch => char.IsLetterOrDigit(ch) || ch == '-' || ch == '_').ToArray());
        if (safe.Length == 0) safe = "sheet";
        string png = Path.Combine(outDir, safe + "-" + address.Replace(":", "_").Replace("!", "_") + ".png");

        if (!TryRender(ws, address, png))
            throw new Fail("COM_ERROR", "CopyPicture produced no image (range " + address + "). The range may be empty, or WPS refused the copy.");

        // rescale to the requested width (dispose the source before touching its file)
        int wantWidth = a.Int("width", 1100);
        Dictionary<string, object> r = new Dictionary<string, object>();
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

    private static int ColorFrom(string s)
    {
        string t = s.Trim().TrimStart('#');
        if (t.Length == 6)
        {
            int v;
            if (int.TryParse(t, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v))
                return Color.FromArgb((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF).ToArgb() & 0xFFFFFF;
        }
        throw new Fail("BAD_ARGS", "color must look like #C00000, got \"" + s + "\"");
    }

    private static object FormatRange(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "No target sheet.");
        dynamic range = ws.Range(a.Need("range"));
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
            int style = b == "none" ? -4142 : b == "thick" ? -4137 : 1; // xlContinuous
            double weight = b == "none" ? 2 : b == "thick" ? -4137 : 2;
            for (int edge = 5; edge <= 12; edge++)
            {
                try
                {
                    dynamic bd = range.Borders.Item(edge);
                    if (b == "none") { bd.LineStyle = -4142; }
                    else { bd.LineStyle = 1; bd.Weight = b == "thick" ? 4 : 2; }
                }
                catch (Exception) { }
            }
            done.Add("borders");
        }
        if (done.Count == 0) throw new Fail("BAD_ARGS", "format needs at least one formatting option (font/size/bold/color/fillColor/halign/valign/numberFormat/colWidth/rowHeight/borders/wrap)");
        return new Dictionary<string, object> { { "range", a.Str("range", "") }, { "applied", done } };
    }

    // ------------------------------------------------------------ snapshot / diff

    private static string SnapPath(dynamic wb)
    {
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate", "excel-snapshots");
        Directory.CreateDirectory(dir);
        string name = Convert.ToString(TryGet(() => (object)wb.Name, "book"), CultureInfo.InvariantCulture);
        string safe = new string(name.Where(char.IsLetterOrDigit).ToArray());
        return Path.Combine(dir, safe + "-last.json");
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
        return new Dictionary<string, object> { { "savedAs", asPath } };
    }

    private static object CloseBook(Bag a)
    {
        string full = FullPathOf(a.Need("doc"));
        dynamic wb = FindOpen(full);
        if (wb == null) throw new Fail("NOT_OPEN", "The workbook is not open: " + full);
        bool save = a.Flag("save", false);
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
        }
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
        if (toast)
        {
            try
            {
                string text = statusLine;
                if (text.Length > 240) text = text.Substring(0, 240);
                using (NotifyIcon ni = new NotifyIcon())
                {
                    ni.Icon = SystemIcons.Information;
                    ni.Visible = true;
                    ni.BalloonTipTitle = "WPS \u8868\u683c AI \u52a9\u624b";  // WPS 表格 AI 助手
                    ni.BalloonTipText = text;
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
            Thread.Sleep(300); // let the sound start before the process exits
        }
        r["sound"] = sound;
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
