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
}

// ---------------------------------------------------------------- helper

internal static class Program
{
    private const string WPS_WPP = "KWPP.Application";

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
            if (requestPath == null) throw new Fail("BAD_ARGS", "usage: WppHelper.exe <request.json> <response.json>");
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
            case "open": return OpenPres(a);
            case "read": return ReadPres(a);
            case "search": return SearchPres(a);
            case "selection": return SelectionCmd(a);
            case "slide": return SlideCmd(a);
            case "text": return TextCmd(a);
            case "textbox": return TextboxCmd(a);
            case "image": return ImageCmd(a);
            case "table": return TableCmd(a);
            case "shape": return ShapeCmd(a);
            case "format": return FormatCmd(a);
            case "notes": return NotesCmd(a);
            case "view": return ViewCmd(a);
            case "render": return RenderPres(a);
            case "snapshot": return SnapshotPres(a);
            case "diff": return DiffPres(a);
            case "save": return SavePres(a);
            case "close": return ClosePres(a);
            case "quit": return QuitAll(a);
            case "finish": return FinishCmd(a);
            default: throw new Fail("BAD_ARGS", "unknown cmd \"" + cmd + "\"; use doctor|status|open|read|search|selection|slide|text|textbox|image|table|shape|format|notes|view|render|snapshot|diff|save|close|quit|finish");
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

    private static dynamic App(bool create, bool visible)
    {
        if (_app != null) return _app;
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
        string path = a.Need("doc");
        string full = FullPathOf(path);
        dynamic p = FindOpen(full);
        if (p != null) { MarkerSet(full); return p; }
        if (MarkerHas(full))
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
        dynamic app;
        dynamic p;
        if (!File.Exists(full))
        {
            // PowerPoint's most common task is a NEW deck: opening a missing path creates one
            // (unnamed in memory; the first save -As/<path> gives it that identity).
            app = App(true, show);
            try { p = app.Presentations.Add(show ? -1 : 0); }
            catch (Exception e) { throw new Fail("COM_ERROR", "could not create a new presentation: " + Clean(e.Message)); }
            // Save immediately: an unnamed deck has FullName "演示文稿N" and could never be
            // found again by later commands (each command is a fresh helper process).
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(full));
                p.SaveAs(full, 24); // ppSaveAsOpenXMLPresentation
            }
            catch (Exception e) { throw new Fail("COM_ERROR", "created the deck but could not save it to " + full + ": " + Clean(e.Message)); }
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

    // ------------------------------------------------------------ colors / text options

    private static int ColorOf(string spec, int fallback)
    {
        if (spec == null) return fallback;
        string s = spec.Trim().TrimStart('#');
        if (s.Length == 6)
        {
            int r = int.Parse(s.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int g = int.Parse(s.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int b = int.Parse(s.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return (r << 16) | (g << 8) | b;   // OLE color: 0x00BBGGRR
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
        dynamic p = RequirePres(a, true);
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)p.Name, "?"), CultureInfo.InvariantCulture);
        r["path"] = Convert.ToString(TryGet(() => (object)p.FullName, full), CultureInfo.InvariantCulture);
        r["readOnly"] = TryGet(() => (object)p.ReadOnly, false);
        r["saved"] = TryGet(() => (object)p.Saved, true);
        r["slides"] = TryGet(() => (object)p.Slides.Count, 0);
        double w = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double h = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        r["slideSize"] = new Dictionary<string, object> { { "w", Math.Round(w, 1) }, { "h", Math.Round(h, 1) }, { "ratio", Math.Round(w / h, 3) } };
        r["how"] = existed ? "reused" : "opened";
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
            for (int j = 1; j <= sc; j++)
            {
                dynamic sh = null;
                try { sh = s.Shapes.Item(j); }
                catch (Exception) { continue; }
                shapes.Add(ShapeInfo(sh, j, textLimit));
            }
            d["shapes"] = shapes;
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
                string text = ShapeText(sh);
                if (text.Length == 0) continue;
                int pos = icase ? text.ToLowerInvariant().IndexOf(needle.ToLowerInvariant(), StringComparison.Ordinal)
                                : text.IndexOf(needle, StringComparison.Ordinal);
                if (pos < 0) continue;
                int start = Math.Max(0, pos - 30);
                int len = Math.Min(text.Length - start, needle.Length + 60);
                hits.Add(new Dictionary<string, object> {
                    { "slide", k },
                    { "shape", Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture) },
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
            dynamic s = p.Slides.Item(index).Duplicate();
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
        throw new Fail("BAD_ARGS", "slide -Action must be add|dup|move|delete");
    }

    /// <summary>Preset layouts (borrowed idea from dsh-office's ready-made slides, deliberately simple):
    /// title / section / bullets / two / end. Geometry derives from the deck's slide size.</summary>
    private static string BuildPreset(dynamic pres, dynamic slide, Bag a, out List<string> problems)
    {
        problems = new List<string>();
        string preset = a.Str("preset", "").ToLowerInvariant();
        double W = Convert.ToDouble(TryGet(() => (object)pres.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double H = Convert.ToDouble(TryGet(() => (object)pres.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);
        string title = a.Str("title", "");
        if (preset != "end" && title.Length == 0) throw new Fail("BAD_ARGS", "presets need -Title (except end)");

        Action<double, double, double, double, string, float, int, string, string> box =
            delegate(double l, double t, double w, double h, string text, float size, int bold, string align, string boxName)
        {
            if (text == null || text.Length == 0) return;
            dynamic tb = slide.Shapes.AddTextbox(1, (float)l, (float)t, (float)w, (float)h);
            if (boxName != null) { try { tb.Name = boxName; } catch (Exception) { } }
            dynamic tr = tb.TextFrame.TextRange;
            tr.Text = text.Replace("\\n", "\n");
            try { tr.Font.Size = size; tr.Font.Bold = bold; } catch (Exception) { }
            int al = align == "center" ? 2 : align == "right" ? 3 : 1;
            try { tr.ParagraphFormat.Alignment = al; } catch (Exception) { }
            try { tb.TextFrame.WordWrap = -1; } catch (Exception) { }
        };

        if (preset == "title")
        {
            box(W * 0.1, H * 0.28, W * 0.8, H * 0.3, title, 40, -1, "center", "title");
            box(W * 0.1, H * 0.62, W * 0.8, H * 0.15, a.Str("subtitle", null), 18, 0, "center", "subtitle");
            return "title";
        }
        if (preset == "section")
        {
            string num = a.Str("number", null);
            if (num != null) box(W * 0.1, H * 0.30, W * 0.8, H * 0.2, num, 28, 0, "center", "number");
            box(W * 0.1, num != null ? H * 0.44 : H * 0.36, W * 0.8, H * 0.25, title, 36, -1, "center", "sectionTitle");
            return "section";
        }
        if (preset == "bullets")
        {
            box(W * 0.06, H * 0.06, W * 0.88, H * 0.14, title, 30, -1, "left", "title");
            StringBuilder sb = new StringBuilder();
            foreach (object item in a.List("items")) sb.Append(Convert.ToString(item, CultureInfo.InvariantCulture)).Append("\n");
            box(W * 0.08, H * 0.24, W * 0.84, H * 0.68, sb.ToString().TrimEnd('\n'), 20, 0, "left", "body");
            return "bullets";
        }
        if (preset == "two")
        {
            box(W * 0.06, H * 0.06, W * 0.88, H * 0.14, title, 30, -1, "left", "title");
            List<object> items = a.List("items");
            string left = items.Count > 0 ? Convert.ToString(items[0], CultureInfo.InvariantCulture) : "";
            string right = items.Count > 1 ? Convert.ToString(items[1], CultureInfo.InvariantCulture) : "";
            box(W * 0.05, H * 0.24, W * 0.42, H * 0.68, left, 18, 0, "left", "left");
            box(W * 0.53, H * 0.24, W * 0.42, H * 0.68, right, 18, 0, "left", "right");
            return "two";
        }
        if (preset == "end")
        {
            box(W * 0.1, H * 0.36, W * 0.8, H * 0.25, a.Str("title", "\u8c22\u8c22"), 40, -1, "center", "title");
            box(W * 0.1, H * 0.62, W * 0.8, H * 0.15, a.Str("subtitle", null), 18, 0, "center", "subtitle");
            return "end";
        }
        throw new Fail("BAD_ARGS", "unknown preset \"" + preset + "\"; use title|section|bullets|two|end (or -Layout <n> for a raw PowerPoint layout)");
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
        dynamic pic;
        try { pic = slide.Shapes.AddPicture(path, 0, -1, (float)l, (float)t, wArg, hArg); }
        catch (COMException e) { throw new Fail("COM_ERROR", "AddPicture failed: " + Clean(e.Message) + " (if the file is encrypted, open it once in WPS or decrypt first)"); }
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
        double l = a.Dbl("l", W * 0.08), t = a.Dbl("t", H * 0.25), w = a.Dbl("w", W * 0.84), h = a.Dbl("h", Math.Min(H * 0.6, 36 * rows.Count));
        dynamic table = slide.Shapes.AddTable(rows.Count, cols, (float)l, (float)t, (float)w, (float)h).Table;
        for (int ri = 0; ri < rows.Count; ri++)
        {
            List<object> row = (List<object>)rows[ri];
            for (int ci = 0; ci < row.Count && ci < cols; ci++)
            {
                dynamic cell = table.Cell(ri + 1, ci + 1);
                string v = row[ci] == null ? "" : Convert.ToString(row[ci], CultureInfo.InvariantCulture);
                cell.Shape.TextFrame.TextRange.Text = v.Replace("\\n", "\n");
                if (ri == 0 && a.Flag("headerRow", true))
                {
                    try { cell.Shape.TextFrame.TextRange.Font.Bold = -1; } catch (Exception) { }
                }
                try { cell.Shape.TextFrame.TextRange.Font.Size = (float)a.Dbl("size", 14); } catch (Exception) { }
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
        if (to - from + 1 > 4) throw new Fail("BAD_ARGS", "render at most 4 slides per call (got " + (to - from + 1) + "); page through with -From/-To");
        int width = a.Int("width", 1100);

        List<object> files = new List<object>();
        for (int k = from; k <= to; k++)
        {
            dynamic slide = p.Slides.Item(k);
            slide.Copy();
            Thread.Sleep(250);
            string png = Path.Combine(outDir, "slide-" + k.ToString(CultureInfo.InvariantCulture) + ".png");
            if (!SaveClipboardImage(png, width))
                throw new Fail("COM_ERROR", "could not get slide " + k + " as an image from the clipboard (is WPS showing a dialog?)");
            files.Add(png);
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slides"] = files;
        r["note"] = "PNG files are written by the helper process (plaintext); verify with the rendered image, not by trusting the command";
        return r;
    }

    [STAThread]
    private static bool SaveClipboardImage(string pngPath, int targetWidth)
    {
        for (int attempt = 0; attempt < 6; attempt++)
        {
            Image result = null;
            // EMF via raw Win32: Clipboard.GetData("EnhancedMetafile") crashes some
            // hosts (access violation observed on this machine family when WPS holds
            // the source); GetClipboardData + CopyEnhMetaFile is the safe route.
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
            // fallback: bitmap format on the clipboard (lower resolution, always works)
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
            if (result != null)
            {
                try
                {
                    result.Save(pngPath, ImageFormat.Png);
                    result.Dispose();
                    return true;
                }
                catch (Exception) { result.Dispose(); }
            }
            Thread.Sleep(200);
        }
        return false;
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
                shapes.Add(new Dictionary<string, object> {
                    { "name", Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture) },
                    { "text", ShapeText(sh) }
                });
            }
            slides.Add(shapes);
        }
        return slides;
    }

    private static object SnapshotPres(Bag a)
    {
        dynamic p = RequirePres(a, true);
        string tag = a.Str("tag", "last");
        string safe = new string(tag.Where(c => char.IsLetterOrDigit(c) || c == '-' || c == '_').ToArray());
        if (safe.Length == 0) safe = "last";
        string name = Convert.ToString(TryGet(() => (object)p.Name, "?"), CultureInfo.InvariantCulture);
        string fileTag = (name.Length > 0 ? name.Replace(".pptx", "").Replace(".ppt", "") : "deck") + "-" + safe + ".json";
        List<object> state = DeckState(p);
        string path = Path.Combine(SnapshotDir(), fileTag);
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
            string name = Convert.ToString(TryGet(() => (object)p.Name, "?"), CultureInfo.InvariantCulture);
            string fileTag = (name.Replace(".pptx", "").Replace(".ppt", "")) + "-" + since + ".json";
            path = Path.Combine(SnapshotDir(), fileTag);
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
        p.Close();
        dynamic app = App(false, false);
        int left = 0;
        try { left = (int)app.Presentations.Count; }
        catch (Exception) { left = 0; }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["closed"] = full;
        r["presentationsLeft"] = left;
        if (left == 0)
        {
            try { app.Quit(); }
            catch (Exception) { }
        }
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
        r["note"] = "finish touches nothing in the deck; save before finishing if the work should persist";
        return r;
    }

    private static object QuitAll(Bag a)
    {
        dynamic app = App(false, false);
        if (app == null) return new Dictionary<string, object> { { "processes", ProcessCount("wpp") } };
        try { app.Quit(); }
        catch (Exception) { }
        return new Dictionary<string, object> { { "processes", ProcessCount("wpp") } };
    }
}
