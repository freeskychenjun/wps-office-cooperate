// wps-word-cooperate helper: drive WPS Writer over COM to read / revise / comment / render a .docx.
// Pure ASCII source (Chinese only ever travels inside JSON strings), C# 5 (compiled by the
// in-box csc.exe of .NET Framework 4.x), no third-party dependency.
//
//   WpsHelper.exe <request.json> <response.json>
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
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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
    /// <summary>A copy without the given keys - lets one branch reuse ApplyParaFormat without
    /// re-interpreting a field that already has a different meaning in that command.</summary>
    public Bag Without(params string[] keys)
    {
        Dictionary<string, object> copy = new Dictionary<string, object>(d);
        foreach (string k in keys) copy.Remove(k);
        return new Bag(copy);
    }
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
    private const string WPS_WORD = "KWPS.Application";
    private const string WPS_EXCEL = "KET.Application";
    private const string WPS_PPT = "KWPP.Application";
    private const string MS_WORD = "Word.Application";

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
            if (requestPath == null) throw new Fail("BAD_ARGS", "usage: WpsHelper.exe <request.json> <response.json>");
            Bag bag = new Bag(Json.Parse(File.ReadAllText(requestPath, Encoding.UTF8)));
            string cmd = bag.Str("cmd", "").Trim().ToLowerInvariant();
            if (cmd.Length == 0) throw new Fail("BAD_ARGS", "\"cmd\" is required");
            string docPath = bag.Str("doc", null);
            Dictionary<string, object> live = null;
            object dispatched = null;
            Signature before = null;
            bool isWrite = IsWriteCommand(cmd) && !(cmd == "revisions" && bag.Str("action", "list").Trim().ToLowerInvariant() == "list");
            if (isWrite && docPath != null)
            {
                try { before = TakeSignature(docPath); } catch (Exception) { before = null; }
            }
            dispatched = Dispatch(cmd, bag);
            live = Ok(cmd, dispatched) as Dictionary<string, object>;
            if (live != null && before != null)
            {
                try { AfterWrite(cmd, docPath, before, live); }
                catch (Exception) { }
            }
            // The document had to be reopened this run (it was closed in WPS): say so, and
            // cross-check the journal - if the last recorded state was "unsaved", the close may
            // have thrown away both the agent's and the human's edits.
            if (_reopened && live != null)
            {
                Dictionary<string, object> innerRe = null;
                object payloadRe;
                if (live.TryGetValue("result", out payloadRe)) innerRe = payloadRe as Dictionary<string, object>;
                if (innerRe != null)
                {
                innerRe["reopened"] = true;
                bool unsavedAtClose = false;
                try
                {
                    Dictionary<string, object> prior = docPath == null ? null : LoadState(docPath);
                    if (prior != null && prior.ContainsKey("saved") && prior["saved"] is bool && !(bool)prior["saved"])
                        unsavedAtClose = true;
                }
                catch (Exception) { }
                if (unsavedAtClose) innerRe["unsavedAtClose"] = true;
                innerRe["reopenedHint"] = unsavedAtClose
                    ? "\u8fd9\u4efd\u6587\u6863\u521a\u521a\u4e0d\u5728 WPS \u91cc\uff0c\u5df2\u4ece\u78c1\u76d8\u91cd\u65b0\u6253\u5f00\uff08\u672a\u5f39\u7a97\uff09\u3002"
                      + "\u4e14\u4e0a\u6b21\u8bb0\u5f55\u65f6\u5b83\u5904\u4e8e\u300c\u672a\u4fdd\u5b58\u300d\u72b6\u6001\u2014\u2014\u5982\u679c\u5173\u95ed\u65f6\u6ca1\u6709\u4fdd\u5b58\uff0c"
                      + "\u5173\u95ed\u524d\u7684\u6539\u52a8\uff08\u542b\u672c skill \u7684\u4fee\u8ba2/\u6279\u6ce8\uff09\u53ef\u80fd\u5df2\u4e22\u5931\uff1a\u5148 diff \u5feb\u7167\u5bf9\u8d26\u3001\u786e\u8ba4\u540e\u518d\u7ee7\u7eed\uff0c\u5e76\u544a\u77e5\u7528\u6237\u3002"
                    : "\u8fd9\u4efd\u6587\u6863\u521a\u521a\u4e0d\u5728 WPS \u91cc\uff0c\u5df2\u4ece\u78c1\u76d8\u91cd\u65b0\u6253\u5f00\uff08\u672a\u5f39\u7a97\uff09\u3002"
                      + "\u82e5\u7528\u6237\u662f\u4e3b\u52a8\u5173\u95ed\u7684\uff0c\u5148\u5411\u4ed6\u786e\u8ba4\u662f\u5426\u7ee7\u7eed\u518d\u52a8\u624b\u3002";
                }
            }
            response = (object)live ?? Ok(cmd, dispatched);
        }
        catch (Fail f)
        {
            response = Err(f.Code, f.Message);
        }
        catch (COMException e)
        {
            string msg = Clean(e.Message);
            string low = msg.ToLowerInvariant();
            bool busy = e.HResult == unchecked((int)0x8001010A)
                || low.IndexOf("busy", StringComparison.Ordinal) >= 0
                || msg.IndexOf("\u5fd9", StringComparison.Ordinal) >= 0
                || msg.IndexOf("\u5bf9\u8bdd\u6846", StringComparison.Ordinal) >= 0
                || low.IndexOf("dialog", StringComparison.Ordinal) >= 0;
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
            case "open": return OpenDoc(a);
            case "read": return ReadDoc(a);
            case "search": return SearchDoc(a);
            case "selection": return SelectionDoc(a);
            case "undo": return UndoDoc(a);
            case "revise": return Revise(a);
            case "comment": return Comment(a);
            case "render": return RenderDoc(a);
            case "font": return SetFont(a);
            case "format": return FormatDoc(a);
            case "style": return StyleDoc(a);
            case "page": return PageDoc(a);
            case "table": return TableDoc(a);
            case "captions": return CaptionsDoc(a);
            case "caption": return CaptionDoc(a);
            case "picture": return PictureDoc(a);
            case "toc": return TocDoc(a);
            case "reference": return ReferenceDoc(a);
            case "view": return ViewDoc(a);
            case "fields": return FieldsDoc(a);
            case "cell": return CellDoc(a);
            case "insert": return InsertDoc(a);
            case "revisions": return RevisionsDoc(a);
            case "snapshot": return SnapshotDoc(a);
            case "diff": return DiffDoc(a);
            case "handoff": return HandoffDoc(a);
            case "audit": return AuditDoc(a);
            case "unit": return UnitsDoc(a);
            case "units": return UnitsDoc(a);
            case "save": return SaveDoc(a);
            case "close": return CloseDoc(a);
            case "quit": return QuitAll(a);
            case "finish": return FinishCmd(a);
            default: throw new Fail("BAD_ARGS", "unknown cmd \"" + cmd + "\"; use doctor|status|open|read|search|selection|undo|revise|comment|render|font|format|style|page|table|captions|caption|picture|toc|reference|view|fields|cell|insert|revisions|snapshot|diff|handoff|audit|save|close|quit");
        }
    }

    // ------------------------------------------------------------ plumbing

    // ------------------------------------------------------------ formatting primitives
    // Shared by format / style / caption / table / picture: one place for what a "--bold" means.

    private static readonly Dictionary<string, int> ALIGN = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        {"left", 0}, {"center", 1}, {"right", 2}, {"justify", 3}, {"distribute", 4},
        {"justified", 3}, {"both", 3}, {"\u4e24\u7aef\u5bf9\u9f50", 3}, {"\u5c45\u5de6", 0}, {"\u5c45\u4e2d", 1}, {"\u5c45\u53f3", 2},
    };

    private static int Alignment(string s)
    {
        if (s == null) return -1;
        int v;
        if (ALIGN.TryGetValue(s.Trim(), out v)) return v;
        throw new Fail("BAD_ARGS", "unknown align \"" + s + "\" (left|center|right|justify)");
    }

    /// <summary>Points from a number of centimetres or a string like "1.2cm" / "24pt" / "18".</summary>
    private static double Points(string s)
    {
        if (s == null) return 0;
        string t = s.Trim().ToLowerInvariant();
        double n;
        if (t.EndsWith("cm"))
        {
            if (double.TryParse(t.Substring(0, t.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out n)) return n * 28.3465;
            throw new Fail("BAD_ARGS", "bad length \"" + s + "\"");
        }
        if (t.EndsWith("pt"))
        {
            if (double.TryParse(t.Substring(0, t.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out n)) return n;
            throw new Fail("BAD_ARGS", "bad length \"" + s + "\"");
        }
        if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out n)) return n;
        throw new Fail("BAD_ARGS", "bad length \"" + s + "\" (use 24, \"24pt\" or \"1.2cm\")");
    }

    /// <summary>"1.5" = 1.5 lines, "20pt" = exactly 20 pt, "at least 20pt" = minimum. Also single|1.5|double.</summary>
    private static void SetLineSpacing(dynamic pf, string spec)
    {
        if (spec == null) return;
        string t = spec.Trim().ToLowerInvariant();
        if (t == "single") { pf.LineSpacingRule = 0; return; }
        if (t == "double") { pf.LineSpacingRule = 5; pf.LineSpacing = 24; return; }
        if (t.StartsWith("at least") || t.StartsWith("minimum") || t.StartsWith("\u6700\u5c0f\u503c"))
        {
            string rest = t.Replace("at least", "").Replace("minimum", "").Replace("\u6700\u5c0f\u503c", "").Trim();
            pf.LineSpacingRule = 3; // wdLineSpaceAtLeast
            pf.LineSpacing = (float)Points(rest);
            return;
        }
        if (t.EndsWith("pt"))
        {
            pf.LineSpacingRule = 4; // wdLineSpaceExactly
            pf.LineSpacing = (float)Points(t);
            return;
        }
        double mult;
        if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out mult))
        {
            if (Math.Abs(mult - 1.0) < 0.001) { pf.LineSpacingRule = 0; return; }
            pf.LineSpacingRule = 5; // wdLineSpaceMultiple
            pf.LineSpacing = (float)(12.0 * mult);
            return;
        }
        throw new Fail("BAD_ARGS", "bad lineSpacing \"" + spec + "\" (1.5 | \"20pt\" | \"at least 20pt\" | single | double)");
    }

    /// <summary>Character formatting of one range/style font. Only the fields present in the request are touched.</summary>
    private static List<string> ApplyCharFormat(dynamic font, Bag a)
    {
        List<string> done = new List<string>();
        if (a.Has("font"))
        {
            string name = a.Str("font", "");
            font.NameFarEast = name;
            font.Name = name;
            try { font.NameAscii = name; } catch (Exception) { }
            try { font.NameOther = name; } catch (Exception) { }
            done.Add("font=" + name);
        }
        if (a.Has("latinFont"))
        {
            string name = a.Str("latinFont", "");
            try { font.NameAscii = name; font.Name = name; } catch (Exception) { }
            done.Add("latinFont=" + name);
        }
        if (a.Has("size")) { string sz = a.Str("size", "12"); font.Size = (float)Points(sz); done.Add("size=" + sz); }
        if (a.Has("bold")) { font.Bold = a.Flag("bold", false) ? 1 : 0; done.Add("bold=" + font.Bold); }
        if (a.Has("italic")) { font.Italic = a.Flag("italic", false) ? 1 : 0; done.Add("italic=" + font.Italic); }
        if (a.Has("underline")) { font.Underline = a.Flag("underline", false) ? 1 : 0; done.Add("underline=" + font.Underline); }
        if (a.Has("super")) { bool v = a.Flag("super", false); font.Superscript = v; if (v) font.Subscript = false; done.Add("super=" + v); }
        if (a.Has("sub")) { bool v = a.Flag("sub", false); font.Subscript = v; if (v) font.Superscript = false; done.Add("sub=" + v); }
        if (a.Has("color"))
        {
            string hex = a.Str("color", "#000000").TrimStart('#');
            if (hex.Length != 6) throw new Fail("BAD_ARGS", "color must be #RRGGBB");
            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
            font.Color = r + g * 256 + b * 65536;
            done.Add("color=#" + hex);
        }
        return done;
    }

    /// <summary>Paragraph formatting of one range/style paragraph format.</summary>
    private static List<string> ApplyParaFormat(dynamic pf, Bag a)
    {
        List<string> done = new List<string>();
        if (a.Has("align")) { pf.Alignment = Alignment(a.Str("align", "left")); done.Add("align=" + a.Str("align", "")); }
        if (a.Has("indentChars")) { pf.CharacterUnitLeftIndent = (float)Points(a.Str("indentChars", "0")); done.Add("indentChars=" + a.Str("indentChars", "")); }
        if (a.Has("leftIndent")) { pf.LeftIndent = (float)Points(a.Str("leftIndent", "0")); done.Add("leftIndent=" + a.Str("leftIndent", "")); }
        if (a.Has("firstLineIndent"))
        {
            string raw = a.Str("firstLineIndent", "0");
            double chars;
            if (double.TryParse(raw.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out chars))
            {
                pf.CharacterUnitFirstLineIndent = (float)chars;   // 2 = two Chinese characters
                done.Add("firstLineIndent=" + chars + "\u5b57\u7b26");
            }
            else
            {
                pf.FirstLineIndent = (float)Points(raw);
                done.Add("firstLineIndent=" + raw);
            }
        }
        if (a.Has("lineSpacing")) { SetLineSpacing(pf, a.Str("lineSpacing", null)); done.Add("lineSpacing=" + a.Str("lineSpacing", "")); }
        if (a.Has("spaceBefore")) { pf.SpaceBefore = (float)Points(a.Str("spaceBefore", "0")); done.Add("spaceBefore=" + a.Str("spaceBefore", "")); }
        if (a.Has("spaceAfter")) { pf.SpaceAfter = (float)Points(a.Str("spaceAfter", "0")); done.Add("spaceAfter=" + a.Str("spaceAfter", "")); }
        if (a.Has("pageBreakBefore")) { pf.PageBreakBefore = a.Flag("pageBreakBefore", false) ? 1 : 0; done.Add("pageBreakBefore=" + a.Flag("pageBreakBefore", false)); }
        if (a.Has("noSplitWords")) { bool v = a.Flag("noSplitWords", true); try { pf.WordWrap = v ? 0 : -1; } catch (Exception) { } done.Add("noSplitWords=" + v); }
        return done;
    }

    private static string Clean(string s)
    {
        if (s == null) return "";
        return s.Replace("\r", " ").Replace("\n", " ").Trim();
    }

    private static string ParaText(object range)
    {
        try
        {
            string t = (string)((dynamic)range).Text;
            return CleanPara(t);
        }
        catch (Exception) { return ""; }
    }

    private static string CleanPara(string t)
    {
        if (t == null) return "";
        return t.TrimEnd('\r', '\n', '\a', '\u0007');
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

    /// <summary>The WPS Writer instance: the running one if there is one, otherwise a new one.</summary>
    private static dynamic App(bool create, bool visible)
    {
        if (_app != null) return _app;
        dynamic found = Running(WPS_WORD) ?? Running(MS_WORD);
        if (found != null) { _app = found; return _app; }
        if (!create) return null;
        if (!Installed(WPS_WORD) && !Installed(MS_WORD))
            throw new Fail("NO_APP", "Neither WPS Writer (KWPS.Application) nor Microsoft Word (Word.Application) is installed on this computer.");
        string prog = Installed(WPS_WORD) ? WPS_WORD : MS_WORD;
        try { _app = Activator.CreateInstance(Type.GetTypeFromProgID(prog, true)); }
        catch (Exception e) { throw new Fail("NO_APP", "Could not start WPS Writer: " + Clean(e.Message)); }
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
        try { n = (int)app.Documents.Count; }
        catch (Exception) { return null; }
        string want = fullPath.ToLowerInvariant();
        for (int k = 1; k <= n; k++)
        {
            dynamic d = null;
            try { d = app.Documents.Item(k); }
            catch (Exception) { continue; }
            string full = null;
            try { full = (string)d.FullName; }
            catch (Exception) { }
            if (full != null && full.ToLowerInvariant() == want) return d;
        }
        return null;
    }

    private static dynamic RequireDoc(Bag a, bool create)
    {
        string path = a.Need("doc");
        string full = FullPathOf(path);
        dynamic doc = FindOpen(full);
        if (doc != null) return doc;
        if (!create) throw new Fail("NOT_OPEN", "The document is not open in WPS Writer: " + full + ". Call open first.");
        // The document was CLOSED under us (the user shut it in WPS). Reopen quietly - no
        // window popping back into the user's face - and flag it so Main can stamp the response.
        _reopened = true;
        return OpenPath(full, false, a.Flag("readOnly", false));
    }

    // Set when this run had to reopen a document the user had closed; Main stamps the response.
    private static bool _reopened;

    private static dynamic OpenPath(string full, bool show, bool readOnly)
    {
        if (!File.Exists(full)) throw new Fail("NOT_OPEN", "The file does not exist: " + full);
        dynamic app = App(true, show);
        dynamic doc;
        try { doc = app.Documents.Open(full, false, readOnly, false); }
        catch (COMException e)
        {
            string msg = Clean(e.Message);
            if (msg.IndexOf("read-only", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("read only", StringComparison.OrdinalIgnoreCase) >= 0)
                doc = app.Documents.Open(full, false, true, false);
            else throw;
        }
        if (show)
        {
            try { app.Visible = true; } catch (Exception) { }
            try { doc.Activate(); } catch (Exception) { }
        }
        return doc;
    }

    private static bool IsReadOnly(dynamic doc)
    {
        object v = TryGet(() => (object)doc.ReadOnly, null);
        if (v is bool) return (bool)v;
        return false;
    }

    private static Dictionary<string, object> DocInfo(dynamic doc)
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        d["path"] = Convert.ToString(TryGet(() => (object)doc.FullName, ""), CultureInfo.InvariantCulture);
        d["readOnly"] = IsReadOnly(doc);
        d["saved"] = TryGet(() => (object)doc.Saved, true);
        d["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        d["tables"] = TryGet(() => (object)(int)doc.Tables.Count, 0);
        d["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        d["revisions"] = TryGet(() => (object)(int)doc.Revisions.Count, 0);
        d["comments"] = TryGet(() => (object)(int)doc.Comments.Count, 0);
        return d;
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

    // ------------------------------------------------------------ commands

    private static object Doctor(Bag a)
    {
        Dictionary<string, object> r = new Dictionary<string, object>();
        Dictionary<string, object> apps = new Dictionary<string, object>();
        string[][] table = new string[][] {
            new string[] {"wps", WPS_WORD}, new string[] {"wpsExcel", WPS_EXCEL}, new string[] {"wpsPpt", WPS_PPT},
            new string[] {"msWord", MS_WORD}, new string[] {"msExcel", "Excel.Application"}, new string[] {"msPpt", "PowerPoint.Application"},
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
            {"wps", ProcessCount("wps")}, {"et", ProcessCount("et")}, {"wpp", ProcessCount("wpp")}, {"winword", ProcessCount("WINWORD")}
        };

        dynamic app = App(false, false);
        if (app != null)
        {
            r["app"] = AppInfo(app);
            r["openDocuments"] = Documents(app);
            // Undo grouping: every batch command wraps its edits in one UndoRecord when the
            // host supports it; if not, Ctrl+Z undoes step by step. Say which world we are in.
            object ur = TryGet(() => (object)app.UndoRecord, null);
            r["undoRecord"] = ur != null;
            if (ur == null) r["undoRecordHint"] = "This WPS does not expose UndoRecord: batch commands cannot be undone with a single Ctrl+Z (each edit is one step). Prefer tracked revisions for text changes.";
            dynamic doc = ActiveDoc(app);
            if (doc != null)
            {
                Dictionary<string, object> probe = new Dictionary<string, object>();
                probe["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
                object bytes = TryGet(() => (object)((byte[])doc.Windows.Item(1).Panes.Item(1).Pages.Item(1).EnhMetaFileBits), null);
                byte[] bits = bytes as byte[];
                probe["emfBytes"] = bits == null ? 0 : bits.Length;
                probe["renderOk"] = bits != null && bits.Length > 1000;
                r["render"] = probe;
                r["encryptedAtRest"] = null; // cannot be told from inside the app
            }
        }
        else
        {
            r["app"] = null;
            r["openDocuments"] = new List<object>();
            r["note"] = "WPS Writer is not running; render capability is reported once a document is open.";
        }
        return r;
    }

    private static dynamic ActiveDoc(dynamic app)
    {
        return TryGet(() => (object)app.ActiveDocument, null) as dynamic;
    }

    private static List<object> Documents(dynamic app)
    {
        List<object> list = new List<object>();
        int n = 0;
        try { n = (int)app.Documents.Count; } catch (Exception) { return list; }
        dynamic active = ActiveDoc(app);
        string activeName = active == null ? "" : Convert.ToString(TryGet(() => (object)active.Name, ""), CultureInfo.InvariantCulture);
        for (int k = 1; k <= n; k++)
        {
            dynamic d = null;
            try { d = app.Documents.Item(k); } catch (Exception) { continue; }
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["name"] = Convert.ToString(TryGet(() => (object)d.Name, ""), CultureInfo.InvariantCulture);
            one["path"] = Convert.ToString(TryGet(() => (object)d.FullName, ""), CultureInfo.InvariantCulture);
            one["readOnly"] = IsReadOnly(d);
            one["saved"] = TryGet(() => (object)d.Saved, true);
            one["active"] = activeName.Length > 0 && Convert.ToString(TryGet(() => (object)d.Name, ""), CultureInfo.InvariantCulture) == activeName;
            one["paragraphs"] = TryGet(() => (object)(int)d.Paragraphs.Count, 0);
            list.Add(one);
        }
        return list;
    }

    private static object Status(Bag a)
    {
        Dictionary<string, object> r = new Dictionary<string, object>();
        dynamic app = App(false, false);
        r["running"] = app != null;
        r["processes"] = ProcessCount("wps");
        r["documents"] = app == null ? new List<object>() : Documents(app);
        if (app != null) r["app"] = AppInfo(app);
        return r;
    }

    private static object OpenDoc(Bag a)
    {
        string full = FullPathOf(a.Need("doc"));
        bool show = a.Flag("show", true);
        dynamic existing = FindOpen(full);
        dynamic doc;
        string how;
        int wpsBefore = ProcessCount("wps");
        bool otherInstancePossible = false;
        if (existing != null) { doc = existing; how = "reused"; }
        else
        {
            // GetActiveObject only ever reaches the FIRST registered instance. If WPS is running
            // with more than one process but that instance did not have the document, the user may
            // have it open in another instance - opening it here would fork the file. Say so.
            dynamic attached = App(false, false);
            otherInstancePossible = attached != null && wpsBefore > 1;
            doc = OpenPath(full, show, a.Flag("readOnly", false)); how = "opened";
        }
        Dictionary<string, object> info = DocInfo(doc);
        info["how"] = how;
        if (otherInstancePossible)
        {
            info["multiInstanceWarning"] = "WPS \u8fd0\u884c\u4e2d\u6709 " + wpsBefore
                + " \u4e2a\u8fdb\u7a0b\uff0c\u4f46\u672c skill \u80fd\u63a5\u5230\u7684\u5b9e\u4f8b\u91cc\u6ca1\u6709\u8fd9\u4efd\u6587\u6863\uff1a\u5b83\u53ef\u80fd\u6b63\u5f00\u5728\u53e6\u4e00\u4e2a\u5b9e\u4f8b\u91cc\u3002"
                + "\u5982\u679c\u4f60\u770b\u5230\u4e24\u4e2a WPS \u7a97\u53e3\u6253\u5f00\u4e86\u540c\u4e00\u4e2a\u6587\u4ef6\uff0c\u8bf7\u5148\u5728 WPS \u91cc\u624b\u5de5\u5173\u6389\u53e6\u4e00\u4efd\u518d\u7ee7\u7eed\uff08\u5426\u5219\u540e\u7eed\u5199\u5165/\u4fdd\u5b58\u4f1a\u4e92\u76f8\u8986\u76d6\uff09\u3002";
        }
        return info;
    }

    /// <summary>
    /// Table spans in document order: {start, end, tableNumber}. <c>Table.Range.End</c> is EXCLUSIVE:
    /// the paragraph that follows a table has Start == that End, so membership must be
    /// <c>start &lt;= offset &amp;&amp; offset &lt; end</c>. The old <c>&lt;=</c> test reported the paragraph
    /// right after a table as being inside it (field report 2026-10-26, a real 75-page report where the
    /// prose paragraph 10297 got inTable:true and cost two renders to clear up).
    /// </summary>
    private static List<int[]> TableSpans(dynamic doc)
    {
        List<int[]> spans = new List<int[]>();
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        for (int t = 1; t <= tables; t++)
        {
            try
            {
                dynamic tb = doc.Tables.Item(t);
                spans.Add(new int[] { (int)tb.Range.Start, (int)tb.Range.End, t });
            }
            catch (Exception) { spans.Add(new int[] { -1, -1, t }); }
        }
        return spans;
    }

    /// <summary>1-based table number for a character offset; 0 when the offset is not inside a table.</summary>
    private static int TableAt(List<int[]> spans, int offset)
    {
        if (offset < 0) return 0;
        foreach (int[] sp in spans)
            if (sp[0] >= 0 && offset >= sp[0] && offset < sp[1]) return sp[2];
        return 0;
    }

    /// <summary>One paragraph of a read: number, text, style, list string, and whether it sits in a table.</summary>
    private static Dictionary<string, object> ReadOnePara(dynamic doc, int k, List<int[]> spans, out bool inTable, out int tableNo)
    {
        inTable = false;
        tableNo = 0;
        Dictionary<string, object> one = new Dictionary<string, object>();
        dynamic p = null;
        try { p = doc.Paragraphs.Item(k); } catch (Exception) { return null; }
        one["i"] = k;
        string text = CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));
        one["text"] = text;
        one["style"] = Convert.ToString(TryGet(() => (object)p.Style.NameLocal, ""), CultureInfo.InvariantCulture);
        string listString = Convert.ToString(TryGet(() => (object)p.Range.ListFormat.ListString, ""), CultureInfo.InvariantCulture);
        if (listString != null && listString.Length > 0) one["list"] = listString;
        int start = -1;
        try { start = (int)p.Range.Start; } catch (Exception) { }
        tableNo = TableAt(spans, start);
        if (tableNo > 0) { inTable = true; one["inTable"] = true; one["table"] = tableNo; }
        return one;
    }

    private static object ReadDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        int total = (int)doc.Paragraphs.Count;
        int from = Math.Max(1, a.Int("from", 1));
        int count = a.Int("count", 60);
        bool skipCells = a.Flag("skipTableCells", false);
        string only = a.Str("only", null);
        Regex onlyRe = null;
        if (only != null && only.Trim().Length > 0)
        {
            try { onlyRe = new Regex(only, RegexOptions.Compiled); }
            catch (Exception e) { throw new Fail("BAD_ARGS", "\"only\" is not a usable regular expression: " + Clean(e.Message)); }
        }
        // -SkipTableCells / -Only change what -Count means: it is then "how many paragraphs to RETURN"
        // instead of "how many paragraph numbers to walk", so a 10k-paragraph report that is 90% table
        // cells can be swept in one or two calls instead of 22 (~4.5 min; field report 2026-10-26).
        bool filtered = skipCells || onlyRe != null;
        if (from > total) from = total;
        int want = count > 0 ? count : int.MaxValue;

        List<int[]> spans = TableSpans(doc);

        List<object> items = new List<object>();
        int skippedCells = 0, skippedOther = 0, scanned = 0, to = from - 1;
        for (int k = from; k <= total; k++)
        {
            if (filtered && items.Count >= want) break;
            if (!filtered && count > 0 && k > from + count - 1) break;
            bool inTable;
            int tableNo;
            Dictionary<string, object> one = ReadOnePara(doc, k, spans, out inTable, out tableNo);
            if (one == null) continue;
            scanned++;
            to = k;
            if (skipCells && inTable) { skippedCells++; continue; }
            if (onlyRe != null && !onlyRe.IsMatch(Convert.ToString(one["text"], CultureInfo.InvariantCulture))) { skippedOther++; continue; }
            items.Add(one);
        }
        if (items.Count == 0 && to < from) to = from;

        List<object> tableList = new List<object>();
        for (int t = 1; t <= spans.Count; t++)
        {
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["table"] = t;
            one["rows"] = TryGet(() => (object)(int)doc.Tables.Item(t).Rows.Count, 0);
            one["columns"] = TryGet(() => (object)(int)doc.Tables.Item(t).Columns.Count, 0);
            string firstRow = "";
            try
            {
                dynamic tb = doc.Tables.Item(t);
                StringBuilder sb = new StringBuilder();
                for (int c = 1; c <= Math.Min(6, (int)tb.Columns.Count); c++)
                {
                    if (c > 1) sb.Append(" | ");
                    sb.Append(CleanPara(Convert.ToString(tb.Cell(1, c).Range.Text, CultureInfo.InvariantCulture)));
                }
                firstRow = sb.ToString();
            }
            catch (Exception) { }
            if (firstRow.Length > 0) one["header"] = firstRow;
            tableList.Add(one);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["readOnly"] = IsReadOnly(doc);
        r["paragraphs"] = total;
        r["from"] = from;
        r["to"] = Math.Max(from, to);
        r["returned"] = items.Count;
        if (filtered)
        {
            r["scanned"] = scanned;
            r["filter"] = (skipCells ? "skipTableCells" : "") + (skipCells && onlyRe != null ? "+" : "") + (onlyRe != null ? "only:" + only : "");
            if (skipCells && skippedCells > 0) r["skippedTableCells"] = skippedCells;
            if (onlyRe != null && skippedOther > 0) r["skippedByOnly"] = skippedOther;
        }
        r["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        r["revisions"] = TryGet(() => (object)(int)doc.Revisions.Count, 0);
        r["comments"] = TryGet(() => (object)(int)doc.Comments.Count, 0);
        r["items"] = items;
        r["tables"] = tableList;
        return r;
    }

    private static int IndexOf(string haystack, string needle, int start)
    {
        if (haystack == null || needle == null || needle.Length == 0) return -1;
        return haystack.IndexOf(needle, start, StringComparison.Ordinal);
    }

    private static object Revise(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; nothing was changed. Close it elsewhere and open it again.");
        List<object> items = a.List("items");
        if (items.Count == 0) throw new Fail("BAD_ARGS", "\"items\" is empty");

        int total = (int)doc.Paragraphs.Count;

        // Pass 1: check every item before touching anything, so that a batch is all or nothing.
        List<Dictionary<string, object>> plan = new List<Dictionary<string, object>>();
        foreach (object raw in items)
        {
            Dictionary<string, object> item = raw as Dictionary<string, object>;
            if (item == null) throw new Fail("BAD_ARGS", "each item must be an object");
            Bag ib = new Bag(item);
            int para = ib.Int("para", 0);
            if (para < 1) throw new Fail("BAD_ARGS", "item without \"para\"");
            if (para > total) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not exist (the document has " + total + ").");
            string expect = ib.Str("expect", "");
            string op = ib.Str("op", "replace").Trim().ToLowerInvariant();
            if (op != "replace" && op != "set_text" && op != "insert_after" && op != "delete")
                throw new Fail("BAD_ARGS", "unknown op \"" + op + "\" (replace|set_text|insert_after|delete)");
            dynamic p0 = doc.Paragraphs.Item(para);
            string t0 = CleanPara(Convert.ToString(TryGet(() => (object)p0.Range.Text, ""), CultureInfo.InvariantCulture));
            if (expect.Length > 0 && !t0.StartsWith(expect, StringComparison.Ordinal))
                throw new Fail("ANCHOR_MISSING", "paragraph " + para + " starts with \"" + Head(t0, 40) + "\", not \"" + Head(expect, 40) + "\". Read the document again and renumber.");
            if (op == "replace")
            {
                string find0 = ib.Need("find");
                if (IndexOf(t0, find0, 0) < 0)
                    throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not contain \"" + find0 + "\" (it reads \"" + Head(t0, 40) + "\").");
            }
            Dictionary<string, object> step = new Dictionary<string, object>();
            step["para"] = para;
            step["op"] = op;
            step["item"] = item;
            plan.Add(step);
        }

        // Highest paragraph first: inserts and deletes then leave the numbers still to come unchanged.
        plan.Sort(delegate(Dictionary<string, object> x, Dictionary<string, object> y)
        {
            return ToInt(y, "para").CompareTo(ToInt(x, "para"));
        });

        dynamic app = App(true, true);
        dynamic undo = null;
        try { undo = app.UndoRecord; undo.StartCustomRecord("wps-word-cooperate batch"); } catch (Exception) { undo = null; }

        bool wasTracking = false;
        try { wasTracking = (bool)doc.TrackRevisions; } catch (Exception) { }
        doc.TrackRevisions = true;

        List<object> done = new List<object>();
        int before = (int)doc.Revisions.Count;
        Exception failure = null;
        try
        {
            // Highest paragraph first: inserts and deletes then leave the lower numbers intact.
            List<object> ordered = new List<object>(items);
            ordered.Sort(delegate(object x, object y)
            {
                int px = ToInt((x as Dictionary<string, object>), "para");
                int py = ToInt((y as Dictionary<string, object>), "para");
                return py.CompareTo(px);
            });

            foreach (Dictionary<string, object> step in plan)
            {
                Dictionary<string, object> item = step["item"] as Dictionary<string, object>;
                Bag ib = new Bag(item);
                int para = ToInt(step, "para");
                string op = Convert.ToString(step["op"], CultureInfo.InvariantCulture);
                string expect = ib.Str("expect", "");

                dynamic p = doc.Paragraphs.Item(para);
                string text = CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));

                Dictionary<string, object> rec = new Dictionary<string, object>();
                rec["para"] = para;
                rec["op"] = op;

                if (op == "replace")
                {
                    string find = ib.Need("find");
                    string replace = ib.Str("replace", "");
                    bool all = ib.Flag("all", false);
                    int hits = 0;
                    for (int guard = 0; guard < 500; guard++)
                    {
                        string now = CleanPara(Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(para).Range.Text, ""), CultureInfo.InvariantCulture));
                        int idx = IndexOf(now, find, 0);
                        if (idx < 0) break;
                        int baseStart = (int)doc.Paragraphs.Item(para).Range.Start;
                        int start = baseStart + idx;
                        // A comment anchor or a field marker takes a position that Range.Text does not
                        // count, so the arithmetic lands one character off (measured: the last character
                        // of "find" survived the replacement). Verify and nudge until the range is the
                        // text we mean to replace.
                        string at = Convert.ToString(doc.Range(start, start + find.Length).Text, CultureInfo.InvariantCulture);
                        if (at != find)
                        {
                            bool fixedStart = false;
                            for (int d = 1; d <= 3 && !fixedStart; d++)
                            {
                                foreach (int candidate in new int[] { start - d, start + d })
                                {
                                    if (candidate < baseStart) continue;
                                    string probe = Convert.ToString(doc.Range(candidate, candidate + find.Length).Text, CultureInfo.InvariantCulture);
                                    if (probe == find) { start = candidate; fixedStart = true; break; }
                                }
                            }
                            if (!fixedStart)
                                throw new Fail("ANCHOR_MISSING", "found \"" + find + "\" in paragraph " + para + " but could not place the edit on it (a comment or field marker next to it shifts the text positions). Delete or move that comment first, or edit the paragraph by hand.");
                        }
                        doc.Range(start, start + find.Length).Text = replace;
                        hits++;
                        if (!all) break;
                    }
                    if (hits == 0) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not contain \"" + find + "\" (current text: \"" + Head(text, 40) + "\").");
                    rec["hits"] = hits;
                }
                else if (op == "set_text")
                {
                    string replacement = ib.Str("text", "");
                    dynamic rng = p.Range;
                    int len = (int)rng.End - (int)rng.Start;
                    if (len > 0) rng.MoveEnd(1, -1); // keep the paragraph mark
                    rng.Text = replacement;
                    rec["chars"] = replacement.Length;
                }
                else if (op == "insert_after")
                {
                    string added = ib.Str("text", "");
                    p.Range.InsertParagraphAfter();
                    int np = para + 1;
                    if (np <= (int)doc.Paragraphs.Count) doc.Paragraphs.Item(np).Range.InsertBefore(added);
                    rec["at"] = np;
                }
                else if (op == "delete")
                {
                    p.Range.Delete();
                    rec["deleted"] = true;
                }
                else throw new Fail("BAD_ARGS", "unknown op \"" + op + "\" (replace|set_text|insert_after|delete)");

                // What the paragraph reads now: lets the caller spot a wrong replacement without reading again.
                try
                {
                    int now = para;
                    if (op == "insert_after") now = ToInt(rec, "at");
                    if (now >= 1 && now <= (int)doc.Paragraphs.Count)
                        rec["textAfter"] = Head(CleanPara(Convert.ToString(doc.Paragraphs.Item(now).Range.Text, CultureInfo.InvariantCulture)), 90);
                }
                catch (Exception) { }

                done.Add(rec);
            }
        }
        catch (Exception e) { failure = e; }
        finally
        {
            try { doc.TrackRevisions = wasTracking; } catch (Exception) { }
            if (undo != null) { try { undo.EndCustomRecord(); } catch (Exception) { } }
        }
        if (failure != null)
        {
            // Take our own revisions back out again - and only ours.
            int rolled = 0;
            try
            {
                for (int i = (int)doc.Revisions.Count; i > before; i--)
                {
                    try { doc.Revisions.Item(i).Reject(); rolled++; } catch (Exception) { }
                }
            }
            catch (Exception) { }
            Fail fail = failure as Fail;
            string code = fail != null ? fail.Code : "COM_ERROR";
            string message = fail != null ? fail.Message : Clean(failure.Message);
            throw new Fail(code, message + " (rolled back " + rolled + " revision(s); nothing left changed)");
        }

        int after = (int)doc.Revisions.Count;
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["applied"] = done.Count;
        r["items"] = done;
        r["revisionsBefore"] = before;
        r["revisionsAfter"] = after;
        r["revisionsAdded"] = after - before;
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["saved"] = false;
        r["note"] = "Changes are tracked revisions in the open document; nothing was saved. Look at them with render, then save.";
        return r;
    }

    private static int ToInt(Dictionary<string, object> d, string key)
    {
        if (d == null || !d.ContainsKey(key) || d[key] == null) return 0;
        object v = d[key];
        if (v is long) return (int)(long)v;
        if (v is double) return (int)(double)v;
        int n;
        return int.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), out n) ? n : 0;
    }

    private static string Head(string s, int n)
    {
        if (s == null) return "";
        return s.Length <= n ? s : s.Substring(0, n) + "...";
    }

    private static object Comment(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        // List mode: the anchor of  comment -Delete -Index n. "Delete the 3rd comment" needs a listing.
        string act = a.Str("action", null);
        if (act != null && act.Trim().ToLowerInvariant() == "list")
        {
            int totalC = (int)TryGet(() => (object)(int)doc.Comments.Count, 0);
            List<object> list = new List<object>();
            for (int i = 1; i <= totalC; i++)
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                try
                {
                    dynamic c = doc.Comments.Item(i);
                    one["i"] = i;   // the n of  comment -Delete -Index n
                    one["author"] = Convert.ToString(TryGet(() => (object)c.Author, ""), CultureInfo.InvariantCulture);
                    one["scope"] = Head(CleanPara(Convert.ToString(TryGet(() => (object)c.Scope.Text, ""), CultureInfo.InvariantCulture)), 60);
                    one["text"] = Head(CleanPara(Convert.ToString(TryGet(() => (object)c.Range.Text, ""), CultureInfo.InvariantCulture)), 200);
                }
                catch (Exception e) { one["i"] = i; one["error"] = Clean(e.Message); }
                list.Add(one);
            }
            Dictionary<string, object> lr = new Dictionary<string, object>();
            lr["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
            lr["comments"] = totalC;
            lr["items"] = list;
            lr["saved"] = true;
            lr["note"] = "i \u5c31\u662f comment -Delete -Index \u7684\u7f16\u53f7\uff1bscope=\u6279\u6ce8\u8d34\u5728\u54ea\u6bb5\u6587\u5b57\uff0ctext=\u6279\u6ce8\u5185\u5bb9\u3002\u53ea\u8bfb\u3002";
            return lr;
        }
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no comment was changed.");
        if (a.Flag("delete", false))
        {
            int commentTotal = (int)TryGet(() => (object)(int)doc.Comments.Count, 0);
            List<object> gone = new List<object>();
            bool all = a.Flag("all", false);
            int index = a.Int("index", 0);
            if (!all && (index < 1 || index > commentTotal))
                throw new Fail("BAD_ARGS", "comment -Delete needs -Index N (the document has " + commentTotal + " comment(s)) or -All");
            for (int i = commentTotal; i >= 1; i--)
            {
                if (!all && i != index) continue;
                Dictionary<string, object> one = new Dictionary<string, object>();
                try
                {
                    dynamic c = doc.Comments.Item(i);
                    one["i"] = i;
                    one["scope"] = CleanPara(Convert.ToString(TryGet(() => (object)c.Scope.Text, ""), CultureInfo.InvariantCulture));
                    c.Delete();
                    one["deleted"] = true;
                }
                catch (Exception e) { one["deleted"] = false; one["error"] = Clean(e.Message); }
                gone.Add(one);
            }
            Dictionary<string, object> dr = new Dictionary<string, object>();
            dr["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
            dr["deleted"] = gone;
            dr["comments"] = TryGet(() => (object)(int)doc.Comments.Count, 0);
            dr["saved"] = false;
            return dr;
        }
        // Batch mode: several comments in one call (what an in-place review needs).
        if (a.Has("items"))
        {
            List<object> list = a.List("items");
            if (list.Count == 0) throw new Fail("BAD_ARGS", "\"items\" is empty");
            int paraCount = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
            List<Dictionary<string, object>> plan = new List<Dictionary<string, object>>();
            foreach (object raw in list)
            {
                Dictionary<string, object> item = raw as Dictionary<string, object>;
                if (item == null) throw new Fail("BAD_ARGS", "each item must be {\"para\":N,\"text\":\"...\"}");
                Bag ib = new Bag(item);
                int pIdx = ib.Int("para", 0);
                if (pIdx < 1 || pIdx > paraCount) throw new Fail("ANCHOR_MISSING", "paragraph " + pIdx + " does not exist (the document has " + paraCount + ").");
                string tx = ib.Str("text", null);
                if (tx == null || tx.Trim().Length == 0) throw new Fail("BAD_ARGS", "item for paragraph " + pIdx + " has no \"text\"");
                string an = ib.Str("anchor", null);
                if (an != null && an.Length > 0)
                {
                    string pt = CleanPara(Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(pIdx).Range.Text, ""), CultureInfo.InvariantCulture));
                    if (IndexOf(pt, an, 0) < 0)
                        throw new Fail("ANCHOR_MISSING", "paragraph " + pIdx + " does not contain \"" + an + "\"; nothing was commented.");
                }
                plan.Add(item);
            }
            dynamic app0 = App(true, true);
            dynamic undo0 = null;
            try { undo0 = app0.UndoRecord; undo0.StartCustomRecord("wps-word-cooperate comments"); } catch (Exception) { undo0 = null; }
            List<object> made = new List<object>();
            try
            {
                foreach (Dictionary<string, object> item in plan)
                {
                    Bag ib = new Bag(item);
                    int pIdx = ib.Int("para", 0);
                    string tx = ib.Str("text", "");
                    string an = ib.Str("anchor", null);
                    dynamic rng = doc.Paragraphs.Item(pIdx).Range;
                    if (an != null && an.Length > 0)
                    {
                        string pt = CleanPara(Convert.ToString(rng.Text, CultureInfo.InvariantCulture));
                        int idx = IndexOf(pt, an, 0);
                        int start = (int)rng.Start + idx;
                        rng = doc.Range(start, start + an.Length);
                    }
                    doc.Comments.Add(rng, tx);
                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["para"] = pIdx;
                    one["anchor"] = an == null ? "" : an;
                    one["text"] = tx;
                    made.Add(one);
                }
            }
            catch (Exception e)
            {
                throw new Fail("COM_ERROR", "comments stopped part way: " + Clean(e.Message));
            }
            finally { if (undo0 != null) { try { undo0.EndCustomRecord(); } catch (Exception) { } } }
            Dictionary<string, object> br = new Dictionary<string, object>();
            br["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
            br["added"] = made.Count;
            br["items"] = made;
            br["comments"] = TryGet(() => (object)(int)doc.Comments.Count, 0);
            br["saved"] = false;
            br["note"] = "All comments were checked first: if one anchor is missing nothing is added.";
            return br;
        }
        int para = a.Int("para", 0);
        if (para < 1) throw new Fail("BAD_ARGS", "\"para\" is required");
        string text = a.Str("text", null);
        if (text == null || text.Length == 0) throw new Fail("BAD_ARGS", "\"text\" is required");

        int total = (int)doc.Paragraphs.Count;
        if (para > total) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not exist (the document has " + total + ").");
        dynamic p = doc.Paragraphs.Item(para);
        string paraText = CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));

        dynamic range = p.Range;
        string anchor = a.Str("anchor", null);
        if (anchor != null && anchor.Length > 0)
        {
            int idx = IndexOf(paraText, anchor, 0);
            if (idx < 0) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not contain \"" + anchor + "\".");
            int start = (int)p.Range.Start + idx;
            range = doc.Range(start, start + anchor.Length);
        }
        doc.Comments.Add(range, text);
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["comments"] = TryGet(() => (object)(int)doc.Comments.Count, 0);
        r["para"] = para;
        r["anchor"] = anchor == null ? "" : anchor;
        r["text"] = text;
        return r;
    }

    private static object RenderDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        string dir = a.Need("out");
        int width = Math.Max(320, Math.Min(2400, a.Int("width", 1100)));
        int page = Math.Max(1, a.Int("page", 1));
        int last = Math.Max(page, a.Int("to", page));
        Directory.CreateDirectory(dir);

        dynamic win = doc.Windows.Item(1);
        try { if ((int)win.View.Type != 3) win.View.Type = 3; } catch (Exception) { }
        dynamic pane = win.Panes.Item(1);
        int total = (int)pane.Pages.Count;
        if (total <= 0) throw new Fail("COM_ERROR", "WPS reports 0 pages. Open the document in print view (View.Type = 3) and try again.");
        if (page > total) throw new Fail("BAD_ARGS", "page " + page + " does not exist (the document has " + total + ").");
        last = Math.Min(last, total);

        List<object> files = new List<object>();
        for (int p = page; p <= last; p++)
        {
            byte[] bits = (byte[])pane.Pages.Item(p).EnhMetaFileBits;
            string file = Path.Combine(dir, "page-" + p.ToString(CultureInfo.InvariantCulture) + ".png");
            SaveEmfAsPng(bits, file, width);
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["page"] = p;
            one["path"] = file;
            using (Image img = Image.FromFile(file)) { one["width"] = (int)img.Width; one["height"] = (int)img.Height; }
            one["emfBytes"] = bits.Length;
            files.Add(one);
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["pages"] = total;
        r["files"] = files;
        r["note"] = "Tracked revisions and comments that are visible in print view appear in these pictures.";
        return r;
    }

    private static void SaveEmfAsPng(byte[] emf, string path, int width)
    {
        using (MemoryStream stream = new MemoryStream(emf))
        using (Metafile meta = new Metafile(stream))
        {
            int height = (int)Math.Round(width * (double)meta.Height / meta.Width);
            int factor = Math.Max(2, (int)Math.Ceiling(3000.0 / width));
            using (Bitmap large = new Bitmap(width * factor, height * factor))
            using (Bitmap small = new Bitmap(width, height))
            {
                using (Graphics g = Graphics.FromImage(large))
                {
                    g.Clear(Color.White);
                    g.DrawImage(meta, 0, 0, large.Width, large.Height);
                }
                using (Graphics g = Graphics.FromImage(small))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(large, 0, 0, width, height);
                }
                small.Save(path, ImageFormat.Png);
            }
        }
    }

    /// <summary>Set the East Asian and Latin font of every story (body, headers, footers, text frames).</summary>
    private static object SetFont(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no font was changed.");
        string name = a.Need("name");
        bool keepTracking = a.Flag("track", false);

        dynamic app = App(true, true);
        dynamic undo = null;
        try { undo = app.UndoRecord; undo.StartCustomRecord("wps-word-cooperate font"); } catch (Exception) { undo = null; }

        bool wasTracking = false;
        try { wasTracking = (bool)doc.TrackRevisions; } catch (Exception) { }
        // Formatting is not a revision here: a whole-document font change would bury the text edits.
        if (!keepTracking) { try { doc.TrackRevisions = false; } catch (Exception) { } }

        List<object> done = new List<object>();
        try
        {
            done.Add(ApplyFont(doc.Content, name, "content"));

            int sections = (int)TryGet(() => (object)(int)doc.Sections.Count, 0);
            for (int s = 1; s <= sections; s++)
            {
                for (int h = 1; h <= 3; h++)
                {
                    try { dynamic hrange = doc.Sections.Item(s).Headers.Item(h).Range; done.Add(ApplyFont(hrange, name, "header " + s + "/" + h)); }
                    catch (Exception) { }
                    try { dynamic frange = doc.Sections.Item(s).Footers.Item(h).Range; done.Add(ApplyFont(frange, name, "footer " + s + "/" + h)); }
                    catch (Exception) { }
                }
            }

            int shapes = (int)TryGet(() => (object)(int)doc.Shapes.Count, 0);
            for (int i = 1; i <= shapes; i++)
            {
                try
                {
                    dynamic shape = doc.Shapes.Item(i);
                    if ((int)shape.TextFrame.HasText != 0)
                        done.Add(ApplyFont(shape.TextFrame.TextRange, name, "shape " + i));
                }
                catch (Exception) { }
            }
        }
        finally
        {
            if (!keepTracking) { try { doc.TrackRevisions = wasTracking; } catch (Exception) { } }
            if (undo != null) { try { undo.EndCustomRecord(); } catch (Exception) { } }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["font"] = name;
        r["ranges"] = done;
        r["readBack"] = Convert.ToString(TryGet(() => (object)doc.Content.Font.NameFarEast, ""), CultureInfo.InvariantCulture);
        r["readBackLatin"] = Convert.ToString(TryGet(() => (object)doc.Content.Font.Name, ""), CultureInfo.InvariantCulture);
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["tracked"] = keepTracking;
        r["saved"] = false;
        r["note"] = "Only the text that is there now was changed; text typed later follows its paragraph style. Undo once (Ctrl+Z) takes the whole change back.";
        return r;
    }

    private static Dictionary<string, object> ApplyFont(dynamic range, string name, string where)
    {
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["where"] = where;
        try
        {
            dynamic f = range.Font;
            f.NameFarEast = name;
            f.Name = name;
            try { f.NameAscii = name; } catch (Exception) { }
            try { f.NameOther = name; } catch (Exception) { }
            r["ok"] = true;
        }
        catch (Exception e) { r["ok"] = false; r["error"] = Clean(e.Message); }
        r["chars"] = Convert.ToString(TryGet(() => (object)CleanPara((string)range.Text).Length, 0), CultureInfo.InvariantCulture);
        return r;
    }

    private static object SaveDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        string asPath = a.Str("as", null);
        if (asPath != null && asPath.Length > 0)
        {
            string full = FullPathOf(asPath);
            string dir = Path.GetDirectoryName(full);
            if (dir != null && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
            doc.SaveAs2(full, 16); // wdFormatXMLDocument
        }
        else doc.Save();
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["path"] = Convert.ToString(TryGet(() => (object)doc.FullName, ""), CultureInfo.InvariantCulture);
        r["saved"] = TryGet(() => (object)doc.Saved, true);
        r["revisions"] = TryGet(() => (object)(int)doc.Revisions.Count, 0);
        return r;
    }

    private static object CloseDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, false);
        bool save = a.Flag("save", false);
        string name = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        string full = Convert.ToString(TryGet(() => (object)doc.FullName, ""), CultureInfo.InvariantCulture);
        // Snapshot the state while the doc is still open, then journal the FINAL saved flag
        // after closing - AfterWrite cannot do this (the document is gone by then), and a stale
        // "unsaved" record would cry wolf on the next reopen (unsavedAtClose false positive).
        Signature closing = null;
        try { closing = TakeSignature(full); } catch (Exception) { }
        doc.Close(save ? 1 : 0); // wdSaveChanges : wdDoNotSaveChanges
        if (closing != null)
        {
            closing.Saved = save;
            try { SaveState(full, closing, save ? "close(-Save)" : "close"); } catch (Exception) { }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["closed"] = name;
        r["saved"] = save;
        int left = 0;
        dynamic app = App(false, false);
        if (app != null)
        {
            left = (int)TryGet(() => (object)(int)app.Documents.Count, 0);
            if (left == 0) { try { app.Quit(); } catch (Exception) { } _app = null; System.Threading.Thread.Sleep(700); }
        }
        r["documentsLeft"] = left;
        r["processes"] = ProcessCount("wps");
        return r;
    }

    // ------------------------------------------------------------ finish
    // Agent-agnostic completion signal: tray balloon + system sound, pure .NET Framework
    // (no DSH, no plugins - must work for any agent on any machine). Never touches the
    // document: notification only, per the user's decision on the Excel sibling skill.

    private static object FinishCmd(Bag a)
    {
        string note = a.Str("note", "");
        string docName = a.Str("doc", null);
        string timeText = DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        string statusText = "\u2705 AI \u5df2\u5b8c\u6210 " + timeText                   // ✅ AI 已完成 ...
            + (docName != null ? " | " + Path.GetFileName(docName) : "")
            + (note.Length > 0 ? " | " + note : "");

        bool quiet = a.Flag("quiet", false);
        bool toast = !quiet && a.Flag("toast", true);
        if (toast)
        {
            try
            {
                string text = statusText;
                if (text.Length > 240) text = text.Substring(0, 240);
                using (NotifyIcon ni = new NotifyIcon())
                {
                    ni.Icon = SystemIcons.Information;
                    ni.Visible = true;
                    ni.BalloonTipTitle = "WPS \u6587\u5b57 AI \u52a9\u624b";  // WPS 文字 AI 助手
                    ni.BalloonTipText = text;
                    ni.ShowBalloonTip(6000);
                    DateTime end = DateTime.Now.AddSeconds(4);
                    while (DateTime.Now < end) { Application.DoEvents(); Thread.Sleep(50); }
                    ni.Visible = false;
                }
            }
            catch (Exception) { toast = false; }
        }

        bool sound = !quiet && a.Flag("sound", true);
        if (sound)
        {
            try { System.Media.SystemSounds.Asterisk.Play(); }
            catch (Exception) { sound = false; }
            Thread.Sleep(300);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["statusText"] = statusText;
        r["toast"] = toast;
        r["sound"] = sound;
        r["note"] = "notification only - the document was not touched";
        return r;
    }

    private static object QuitAll(Bag a)
    {
        dynamic app = App(false, false);
        int before = ProcessCount("wps");
        if (app != null)
        {
            int n = (int)TryGet(() => (object)(int)app.Documents.Count, 0);
            for (int k = n; k >= 1; k--)
            {
                try { app.Documents.Item(k).Close(0); } catch (Exception) { }
            }
            try { app.Quit(); } catch (Exception) { }
            _app = null;
        }
        System.Threading.Thread.Sleep(800);
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["processesBefore"] = before;
        r["processesAfter"] = ProcessCount("wps");
        return r;
    }
}
