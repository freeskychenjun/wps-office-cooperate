// LaTeX (a practical subset) -> OMML, the equation markup of .docx files.
//
// WPS Writer only builds a native equation when it is handed a complete Flat OPC package whose
// /word/document.xml carries an <m:oMath> element; a bare <m:oMathPara> is ignored, and
// Document.OMaths.Add(range) is refused outright ("value out of expected range"). So the LaTeX
// has to be translated *here*: Microsoft Word can turn linear text into an equation itself, WPS
// cannot.
//
// The converter is a small recursive-descent parser that writes OMML directly (no MathML in the
// middle). Everything it cannot read throws OmmlFail, and the caller turns that into a
// MATH_UNSUPPORTED error instead of silently writing a wrong equation.
//
// Pure ASCII source: every non-ASCII character is written as a \uXXXX escape.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

internal sealed class OmmlFail : Exception
{
    public OmmlFail(string message) : base(message) { }
}

internal static class Omml
{
    /// <summary>One formula as an &lt;m:oMath&gt; element, or throws OmmlFail.</summary>
    public static string LatexToOmml(string latex)
    {
        if (latex == null) throw new OmmlFail("empty formula");
        TexParser p = new TexParser(latex);
        string body = p.ParseAll();
        if (body.Length == 0) throw new OmmlFail("the formula has no content");
        return "<m:oMath>" + body + "</m:oMath>";
    }

    public static string Esc(string text)
    {
        if (text == null) return "";
        return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
    }
}

/// <summary>LaTeX -> OMML. One instance per formula; not thread safe.</summary>
internal sealed class TexParser
{
    private const string FONT = "<w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr>";
    private readonly string s;
    private int i;

    // greek letters, operators and relations
    private static readonly Dictionary<string, string> SYMBOLS = BuildSymbols();

    // n-ary operators (the big ones that take a lower and an upper limit)
    private static readonly Dictionary<string, string> NARY = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        {"sum", "\u2211"}, {"prod", "\u220f"}, {"coprod", "\u2210"}, {"int", "\u222b"},
        {"iint", "\u222c"}, {"iiint", "\u222d"}, {"oint", "\u222e"}, {"bigcup", "\u22c3"},
        {"bigcap", "\u22c2"}, {"bigvee", "\u22c1"}, {"bigwedge", "\u22c0"},
    };

    // upright function names ("sin", "log", "lim"): written as plain runs
    private static readonly Dictionary<string, string> FUNCS = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        {"sin", "sin"}, {"cos", "cos"}, {"tan", "tan"}, {"cot", "cot"}, {"sec", "sec"}, {"csc", "csc"},
        {"arcsin", "arcsin"}, {"arccos", "arccos"}, {"arctan", "arctan"}, {"sinh", "sinh"}, {"cosh", "cosh"},
        {"tanh", "tanh"}, {"log", "log"}, {"ln", "ln"}, {"exp", "exp"}, {"lg", "lg"},
        {"lim", "lim"}, {"limsup", "lim sup"}, {"liminf", "lim inf"}, {"max", "max"}, {"min", "min"},
        {"sup", "sup"}, {"inf", "inf"}, {"det", "det"}, {"arg", "arg"}, {"deg", "deg"}, {"dim", "dim"},
        {"gcd", "gcd"}, {"Pr", "Pr"},
    };

    // delimiters after \left / \right
    private static readonly Dictionary<string, string> DELIM = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        {"{", "{"}, {"}", "}"}, {"|", "\u2016"}, {"langle", "\u27e8"}, {"rangle", "\u27e9"},
        {"lceil", "\u2308"}, {"rceil", "\u2309"}, {"lfloor", "\u230a"}, {"rfloor", "\u230b"},
        {"vert", "|"}, {"lvert", "|"}, {"rvert", "|"}, {"Vert", "\u2016"}, {"lVert", "\u2016"},
        {"rVert", "\u2016"}, {"lbrack", "["}, {"rbrack", "]"}, {"lbrace", "{"}, {"rbrace", "}"},
        {".", ""}, {"backslash", "\\"}, {"uparrow", "\u2191"}, {"downarrow", "\u2193"},
    };

    // accents: command -> combining mark OMML wants
    private static readonly Dictionary<string, string> ACCENTS = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        {"hat", "\u0302"}, {"widehat", "\u0302"}, {"tilde", "\u0303"}, {"widetilde", "\u0303"},
        {"vec", "\u20d7"}, {"dot", "\u0307"}, {"ddot", "\u0308"}, {"check", "\u030c"},
        {"breve", "\u0306"}, {"acute", "\u0301"}, {"grave", "\u0300"}, {"mathring", "\u030a"},
    };

    private static Dictionary<string, string> BuildSymbols()
    {
        Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.Ordinal);
        string[,] greek = new string[,] {
            {"alpha", "\u03b1"}, {"beta", "\u03b2"}, {"gamma", "\u03b3"}, {"delta", "\u03b4"},
            {"epsilon", "\u03b5"}, {"varepsilon", "\u03b5"}, {"zeta", "\u03b6"}, {"eta", "\u03b7"},
            {"theta", "\u03b8"}, {"vartheta", "\u03d1"}, {"iota", "\u03b9"}, {"kappa", "\u03ba"},
            {"lambda", "\u03bb"}, {"mu", "\u03bc"}, {"nu", "\u03bd"}, {"xi", "\u03be"},
            {"omicron", "\u03bf"}, {"pi", "\u03c0"}, {"varpi", "\u03d6"}, {"rho", "\u03c1"},
            {"varrho", "\u03f1"}, {"sigma", "\u03c3"}, {"varsigma", "\u03c2"}, {"tau", "\u03c4"},
            {"upsilon", "\u03c5"}, {"phi", "\u03c6"}, {"varphi", "\u03d5"}, {"chi", "\u03c7"},
            {"psi", "\u03c8"}, {"omega", "\u03c9"},
            {"Gamma", "\u0393"}, {"Delta", "\u0394"}, {"Theta", "\u0398"}, {"Lambda", "\u039b"},
            {"Xi", "\u039e"}, {"Pi", "\u03a0"}, {"Sigma", "\u03a3"}, {"Upsilon", "\u03a5"},
            {"Phi", "\u03a6"}, {"Psi", "\u03a8"}, {"Omega", "\u03a9"},
        };
        for (int k = 0; k < greek.GetLength(0); k++) d[greek[k, 0]] = greek[k, 1];
        string[,] ops = new string[,] {
            {"le", "\u2264"}, {"leq", "\u2264"}, {"ge", "\u2265"}, {"geq", "\u2265"},
            {"ne", "\u2260"}, {"neq", "\u2260"}, {"approx", "\u2248"}, {"equiv", "\u2261"},
            {"sim", "\u223c"}, {"simeq", "\u2243"}, {"cong", "\u2245"}, {"propto", "\u221d"},
            {"pm", "\u00b1"}, {"mp", "\u2213"}, {"times", "\u00d7"}, {"div", "\u00f7"},
            {"cdot", "\u22c5"}, {"ast", "\u2217"}, {"circ", "\u2218"}, {"bullet", "\u2022"},
            {"infty", "\u221e"}, {"partial", "\u2202"}, {"nabla", "\u2207"}, {"angle", "\u2220"},
            {"to", "\u2192"}, {"rightarrow", "\u2192"}, {"leftarrow", "\u2190"},
            {"Rightarrow", "\u21d2"}, {"Leftarrow", "\u21d0"}, {"leftrightarrow", "\u2194"},
            {"Leftrightarrow", "\u21d4"}, {"mapsto", "\u21a6"},
            {"in", "\u2208"}, {"notin", "\u2209"}, {"ni", "\u220b"},
            {"subset", "\u2282"}, {"subseteq", "\u2286"}, {"supset", "\u2283"}, {"supseteq", "\u2287"},
            {"cup", "\u222a"}, {"cap", "\u2229"}, {"setminus", "\u2216"},
            {"forall", "\u2200"}, {"exists", "\u2203"}, {"nexists", "\u2204"},
            {"emptyset", "\u2205"}, {"varnothing", "\u2205"}, {"perp", "\u22a5"}, {"parallel", "\u2225"},
            {"degree", "\u00b0"}, {"prime", "\u2032"}, {"ldots", "\u2026"}, {"cdots", "\u22ef"},
            {"dots", "\u2026"}, {"vdots", "\u22ee"}, {"ddots", "\u22f1"},
            {"therefore", "\u2234"}, {"because", "\u2235"}, {"mid", "|"},
            {"langle", "\u27e8"}, {"rangle", "\u27e9"},
            {"lbrace", "{"}, {"rbrace", "}"}, {"vert", "|"}, {"Vert", "\u2016"},
            {"hbar", "\u210f"}, {"ell", "\u2113"}, {"Re", "\u211c"}, {"Im", "\u2111"},
            {"aleph", "\u2135"}, {"surd", "\u221a"},
        };
        for (int k = 0; k < ops.GetLength(0); k++) d[ops[k, 0]] = ops[k, 1];
        return d;
    }

    public TexParser(string text) { s = text ?? ""; }

    // ---------------------------------------------------------------- runs

    private static string Run(string text) { return RunStyled(text, 0); }

    /// <summary>style: 0 = mathematical (italic letters), 1 = plain upright, 2 = normal text</summary>
    private static string RunStyled(string text, int style)
    {
        if (text == null || text.Length == 0) return "";
        string props = style == 2 ? "<m:rPr><m:nor/></m:rPr>"
            : style == 1 ? "<m:rPr><m:sty m:val=\"p\"/></m:rPr>" : "";
        return "<m:r>" + props + FONT + "<m:t xml:space=\"preserve\">" + Omml.Esc(text) + "</m:t></m:r>";
    }

    private static string Join(List<string> parts)
    {
        if (parts == null || parts.Count == 0) return "";
        StringBuilder sb = new StringBuilder();
        for (int k = 0; k < parts.Count; k++) sb.Append(parts[k]);
        return sb.ToString();
    }

    // ---------------------------------------------------------------- scanning

    private bool AtEnd { get { return i >= s.Length; } }

    private bool AtCmd(string name)
    {
        if (i >= s.Length || s[i] != '\\') return false;
        string rest = s.Substring(i + 1);
        if (!rest.StartsWith(name, StringComparison.Ordinal)) return false;
        if (name.Length > 0 && char.IsLetter(name[name.Length - 1]))
        {
            // "\\int" must not match "\\intx": a command name ends at a non-letter
            if (rest.Length > name.Length && char.IsLetter(rest[name.Length])) return false;
        }
        return true;
    }

    private void SkipSpace()
    {
        while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n')) i++;
    }

    private string ReadName()
    {
        // assumes s[i] == '\\'
        i++;
        if (AtEnd) throw new OmmlFail("a backslash at the end of the formula");
        StringBuilder sb = new StringBuilder();
        if (char.IsLetter(s[i]))
        {
            while (i < s.Length && char.IsLetter(s[i])) { sb.Append(s[i]); i++; }
        }
        else { sb.Append(s[i]); i++; }
        if (i < s.Length && s[i] == '*') i++;   // \operatorname* / \tag* : star ignored
        return sb.ToString();
    }

    /// <summary>The raw text of a {...} group, braces kept balanced.</summary>
    private string RawGroup()
    {
        SkipSpace();
        if (AtEnd || s[i] != '{') throw new OmmlFail("expected {");
        int depth = 0, start = i + 1;
        while (i < s.Length)
        {
            char c = s[i];
            if (c == '\\') { i += 2; continue; }
            if (c == '{') depth++;
            else if (c == '}') { depth--; if (depth == 0) { string raw = s.Substring(start, i - start); i++; return raw; } }
            i++;
        }
        throw new OmmlFail("unclosed {");
    }

    // ---------------------------------------------------------------- grammar

    public string ParseAll()
    {
        List<string> parts = Seq(new char[0], true);
        return Join(parts);
    }

    /// <summary>
    /// A run of atoms. Stops on one of `stops` ('}', ')' , ']', '&'), on \right, on \\ and on
    /// \end - the callers own those characters.
    /// </summary>
    private List<string> Seq(char[] stops, bool top)
    {
        List<string> parts = new List<string>();
        StringBuilder pending = new StringBuilder();
        while (i < s.Length)
        {
            char c = s[i];
            if (IsStop(c, stops)) break;
            if (c == '\\' && (AtCmd("\\\\") || AtCmd("right") || AtCmd("end") || AtCmd("cr"))) break;
            if (c == ' ' || c == '\t' || c == '\r' || c == '\n') { i++; continue; }
            if (c == '^' || c == '_')
            {
                string baseElem = TakeBase(parts, pending);
                string wrapped = Scripts(baseElem);
                if (wrapped.Length > 0) parts.Add(wrapped);
                continue;
            }
            if (c == '{')
            {
                Flush(parts, pending);
                parts.Add(AttachScripts(Group()));
                continue;
            }
            if (c == '\\')
            {
                Flush(parts, pending);
                string elem = Command();
                if (elem != null && elem.Length > 0) parts.Add(AttachScripts(elem));
                continue;
            }
            if (c == '~') { Flush(parts, pending); parts.Add(Run("\u00a0")); i++; continue; }
            if (c == '&' && !IsStop('&', stops))
            {
                // alignment mark outside a matrix: it only means "next column"
                Flush(parts, pending); i++; continue;
            }
            pending.Append(MapChar(c));
            i++;
        }
        Flush(parts, pending);
        return parts;
    }

    private static bool IsStop(char c, char[] stops)
    {
        if (stops == null) return false;
        for (int k = 0; k < stops.Length; k++) if (stops[k] == c) return true;
        return false;
    }

    private static void Flush(List<string> parts, StringBuilder pending)
    {
        if (pending.Length == 0) return;
        parts.Add(Run(pending.ToString()));
        pending.Length = 0;
    }

    /// <summary>The element a script applies to: the pending text (last character only) or the last atom.</summary>
    private string TakeBase(List<string> parts, StringBuilder pending)
    {
        if (pending.Length > 1)
        {
            parts.Add(Run(pending.ToString(0, pending.Length - 1)));
            string last = pending.ToString(pending.Length - 1, 1);
            pending.Length = 0;
            return Run(last);
        }
        if (pending.Length == 1)
        {
            string only = pending.ToString();
            pending.Length = 0;
            return Run(only);
        }
        if (parts.Count > 0)
        {
            string last = parts[parts.Count - 1];
            parts.RemoveAt(parts.Count - 1);
            return last;
        }
        return "";
    }

    /// <summary>Reads ^ and _ (in any order, repeated) and wraps the base.</summary>
    private string Scripts(string baseElem)
    {
        string sub = null, sup = null;
        int primes = 0;
        while (i < s.Length)
        {
            char c = s[i];
            if (c == '^' || c == '_')
            {
                i++;
                string arg = ArgElem();
                if (c == '^') sup = arg; else sub = arg;
                continue;
            }
            if (c == '\'') { i++; primes++; continue; }
            break;
        }
        if (primes > 0)
        {
            StringBuilder pb = new StringBuilder();
            for (int k = 0; k < primes; k++) pb.Append('\u2032');
            string marks = Run(pb.ToString());
            sup = sup == null ? marks : sup + marks;
        }
        if (sub == null && sup == null) return baseElem;
        if (baseElem.Length == 0) baseElem = "";
        if (sub != null && sup != null)
            return "<m:sSubSup><m:e>" + baseElem + "</m:e><m:sub>" + sub + "</m:sub><m:sup>" + sup + "</m:sup></m:sSubSup>";
        if (sup != null)
            return "<m:sSup><m:e>" + baseElem + "</m:e><m:sup>" + sup + "</m:sup></m:sSup>";
        return "<m:sSub><m:e>" + baseElem + "</m:e><m:sub>" + sub + "</m:sub></m:sSub>";
    }

    private string AttachScripts(string baseElem)
    {
        if (i >= s.Length) return baseElem;
        char c = s[i];
        if (c == '^' || c == '_' || c == '\'') return Scripts(baseElem);
        return baseElem;
    }

    /// <summary>One argument: a {...} group or a single atom, as one OMML fragment.</summary>
    private string ArgElem()
    {
        SkipSpace();
        if (AtEnd) throw new OmmlFail("a script or operator is missing its argument");
        if (s[i] == '{') return Group();
        if (s[i] == '\\') return Command();
        if (s[i] == '^' || s[i] == '_') throw new OmmlFail("double script");
        string one = MapChar(s[i]);
        i++;
        return Run(one);
    }

    private string Group()
    {
        ExpectChar('{');
        List<string> inner = Seq(new char[] { '}' }, false);
        ExpectChar('}');
        return Join(inner);
    }

    private void ExpectChar(char c)
    {
        if (AtEnd || s[i] != c) throw new OmmlFail("expected " + c);
        i++;
    }

    // ---------------------------------------------------------------- atoms

    private static string MapChar(char c)
    {
        // TeX typesets a hyphen as a real minus sign and an asterisk as a binary star
        if (c == '-') return "\u2212";
        if (c == '*') return "\u2217";
        if (c == '\u0000') return "";
        return c.ToString();
    }

    private string Command()
    {
        string name = ReadName();
        // escaped characters
        if (name.Length == 1)
        {
            char ch = name[0];
            if (ch == '%' || ch == '$' || ch == '&' || ch == '#' || ch == '_' || ch == '{' || ch == '}')
                return Run(ch.ToString());
            if (ch == ' ') return Run("\u00a0");
            if (ch == ',') return Run("\u2009");
            if (ch == ':') return Run("\u2005");
            if (ch == ';') return Run("\u2004");
            if (ch == '!') return "";
            if (ch == '|') return Run("\u2016");
            if (ch == '/') return "";
            if (ch == '-') return "";
            if (ch == '\\') return "";   // handled by Seq as a row break
            string sym0;
            if (SYMBOLS.TryGetValue(name, out sym0)) return Run(sym0);
            throw new OmmlFail("unsupported command \"\\" + name + "\"");
        }

        if (name == "frac" || name == "dfrac" || name == "tfrac" || name == "cfrac" || name == "binom")
            return Frac(name == "binom");

        if (name == "sqrt") return Sqrt();

        if (name == "text" || name == "textnormal" || name == "mbox" || name == "textrm")
            return RunStyled(RawGroup(), 2);
        if (name == "mathrm" || name == "rm" || name == "operatorname" || name == "textup" || name == "mathsf")
            return RunStyled(RawGroup(), 1);
        if (name == "mathbf" || name == "bf" || name == "boldsymbol" || name == "bm" || name == "mathtt")
            return RunStyled(RawGroup(), 0);   // bold needs a run property OMML sets on the whole equation
        if (name == "mathit" || name == "it" || name == "mathcal" || name == "mathbb" || name == "cal")
            return RunStyled(RawGroup(), 0);

        if (name == "left") return Delimited();
        if (name == "begin") return Environment();

        if (ACCENTS.ContainsKey(name))
        {
            string chr = ACCENTS[name];
            string inner = ArgElem();
            return "<m:acc><m:accPr><m:chr m:val=\"" + chr + "\"/></m:accPr><m:e>" + inner + "</m:e></m:acc>";
        }
        if (name == "overline" || name == "bar" || name == "underbar")
        {
            string inner = ArgElem();
            string pos = name == "underbar" ? "bot" : "top";
            return "<m:bar><m:barPr><m:pos m:val=\"" + pos + "\"/></m:barPr><m:e>" + inner + "</m:e></m:bar>";
        }
        if (name == "underline")
        {
            string inner = ArgElem();
            return "<m:bar><m:barPr><m:pos m:val=\"bot\"/></m:barPr><m:e>" + inner + "</m:e></m:bar>";
        }

        if (NARY.ContainsKey(name)) return Nary(name);
        if (FUNCS.ContainsKey(name))
        {
            string word = FUNCS[name];
            string known;
            if (SYMBOLS.TryGetValue(name, out known)) word = known;
            string baseElem = RunStyled(word, 1);
            // \lim_{x\to0}: a limit set underneath, not as a subscript to the right
            if ((name == "lim" || name == "limsup" || name == "liminf" || name == "max" || name == "min"
                 || name == "sup" || name == "inf") && i < s.Length && s[i] == '_')
            {
                i++;
                string low = ArgElem();
                if (i < s.Length && s[i] == '^')
                {
                    i++;
                    string up = ArgElem();
                    return "<m:limUpp><m:e><m:limLow><m:e>" + baseElem + "</m:e><m:lim>" + low + "</m:lim></m:limLow></m:e><m:lim>"
                        + up + "</m:lim></m:limUpp>";
                }
                return "<m:limLow><m:e>" + baseElem + "</m:e><m:lim>" + low + "</m:lim></m:limLow>";
            }
            return baseElem;
        }

        string sym;
        if (SYMBOLS.TryGetValue(name, out sym)) return Run(sym);

        if (name == "quad") return Run("\u2003\u2003");
        if (name == "qquad") return Run("\u2003\u2003\u2003\u2003");
        if (name == "thinspace" || name == "enspace" || name == "medspace" || name == "thickspace") return Run("\u2009");
        if (name == "displaystyle" || name == "textstyle" || name == "scriptstyle" || name == "scriptscriptstyle"
            || name == "limits" || name == "nolimits" || name == "notag" || name == "nonumber"
            || name == "left." || name == "mathstrut" || name == "strut" || name == "displaystyle")
            return "";
        if (name == "tag" || name == "label")
        {
            // A \tag inside the converter is ignored: numbers are handled by the caller.
            RawGroup();
            return "";
        }
        if (name == "phantom" || name == "hphantom" || name == "vphantom") { ArgElem(); return ""; }
        if (name == "overbrace" || name == "underbrace")
        {
            string inner = ArgElem();
            string chr = name == "overbrace" ? "\u23de" : "\u23df";
            string pos = name == "overbrace" ? "top" : "bot";
            string vjc = name == "overbrace" ? "bot" : "top";
            return "<m:groupChr><m:groupChrPr><m:chr m:val=\"" + chr + "\"/><m:pos m:val=\"" + pos
                + "\"/><m:vertJc m:val=\"" + vjc + "\"/></m:groupChrPr><m:e>" + inner + "</m:e></m:groupChr>";
        }
        if (name == "not") return Run("\u0338");

        throw new OmmlFail("unsupported command \"\\" + name + "\"");
    }

    private string Frac(bool binom)
    {
        string num = ArgElem();
        SkipSpace();
        string den = ArgElem();
        string frac = binom
            ? "<m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr><m:num>" + num + "</m:num><m:den>" + den + "</m:den></m:f>"
            : "<m:f><m:num>" + num + "</m:num><m:den>" + den + "</m:den></m:f>";
        if (!binom) return frac;
        return "<m:d><m:dPr><m:begChr m:val=\"(\"/><m:endChr m:val=\")\"/></m:dPr><m:e>" + frac + "</m:e></m:d>";
    }

    private string Sqrt()
    {
        string deg = null;
        SkipSpace();
        if (!AtEnd && s[i] == '[')
        {
            int depth = 1, start = ++i;
            while (i < s.Length && depth > 0)
            {
                if (s[i] == '[') depth++;
                else if (s[i] == ']') depth--;
                if (depth == 0) break;
                i++;
            }
            string raw = s.Substring(start, i - start);
            i++;   // ']'
            deg = new TexParser(raw).ParseAll();
        }
        string radicand = ArgElem();
        if (deg == null)
            return "<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/><m:e>" + radicand + "</m:e></m:rad>";
        return "<m:rad><m:deg>" + deg + "</m:deg><m:e>" + radicand + "</m:e></m:rad>";
    }

    private string Nary(string name)
    {
        string sign = NARY[name];
        string sub = null, sup = null;
        // \sum, \sum_{i=1}^{n}, \sum_i^n, \sum\limits_{...}
        while (true)
        {
            SkipSpace();
            if (i < s.Length && s[i] == '\\')
            {
                if (AtCmd("limits") || AtCmd("nolimits")) { ReadName(); continue; }
            }
            if (i < s.Length && (s[i] == '_' || s[i] == '^'))
            {
                char c = s[i];
                i++;
                string arg = ArgElem();
                if (c == '_') sub = arg; else sup = arg;
                continue;
            }
            break;
        }
        // what the operator applies to: the next atom, unless a relation or a comma follows
        string operand = "";
        SkipSpace();
        if (i < s.Length)
        {
            char c = s[i];
            bool boundary = c == '=' || c == '+' || c == '-' || c == ',' || c == ')' || c == ']' || c == '}'
                || (c == '\\' && (AtCmd("right") || AtCmd("end") || AtCmd("\\")));
            // The operand keeps its own scripts: in "\sum x_i^2" the i and the 2 belong to x,
            // not to the operator (they must not end up as a script of the whole n-ary element).
            if (!boundary) operand = AttachScripts(ArgElem());
        }
        return "<m:nary><m:naryPr><m:chr m:val=\"" + sign + "\"/><m:limLoc m:val=\""
            + (sub != null || sup != null ? "undOvr" : "subSup") + "\"/>"
            + (sub == null ? "<m:subHide m:val=\"1\"/>" : "")
            + (sup == null ? "<m:supHide m:val=\"1\"/>" : "")
            + "</m:naryPr><m:sub>" + (sub == null ? "" : sub) + "</m:sub><m:sup>" + (sup == null ? "" : sup)
            + "</m:sup><m:e>" + operand + "</m:e></m:nary>";
    }

    private string ReadDelim() { return ReadDelimAt(); }

    private string ReadDelimAt()
    {
        SkipSpace();
        if (AtEnd) throw new OmmlFail("a delimiter is missing after \\left or \\right");
        char c = s[i];
        if (c == '{')
        {
            string raw = RawGroup().Trim();
            return DelimOf(raw);
        }
        if (c == '\\')
        {
            string name = ReadName();
            string d;
            if (DELIM.TryGetValue(name, out d)) return d;
            if (name.Length == 1) return name;
            throw new OmmlFail("unsupported delimiter \"\\" + name + "\"");
        }
        i++;
        if (c == '.') return "";
        return c.ToString();
    }

    private static string DelimOf(string raw)
    {
        if (raw.Length == 0) return "";
        if (raw == ".") return "";
        if (raw[0] == '\\')
        {
            string name = raw.Substring(1);
            string d;
            if (DELIM.TryGetValue(name, out d)) return d;
            return raw;
        }
        return raw;
    }

    private string Delimited()
    {
        string open = ReadDelimAt();
        List<string> inner = Seq(new char[0], false);
        if (!AtCmd("right")) throw new OmmlFail("\\left without \\right");
        i += 6;   // \right
        string close = ReadDelimAt();
        return "<m:d><m:dPr><m:begChr m:val=\"" + Omml.Esc(open) + "\"/><m:endChr m:val=\""
            + Omml.Esc(close) + "\"/></m:dPr><m:e>" + Join(inner) + "</m:e></m:d>";
    }

    // ---------------------------------------------------------------- environments

    private string Environment()
    {
        string env = RawGroup().Trim();
        List<List<string>> rows = new List<List<string>>();
        List<string> cells = new List<string>();
        while (true)
        {
            List<string> cell = Seq(new char[] { '&' }, false);
            cells.Add(Join(cell));
            if (AtEnd) throw new OmmlFail("\\begin{" + env + "} is never closed");
            if (s[i] == '&') { i++; continue; }
            if (AtCmd("\\\\") || AtCmd("cr"))
            {
                if (AtCmd("\\\\")) i += 2; else i += 3;
                SkipSpace();
                if (!AtEnd && s[i] == '[') { while (i < s.Length && s[i] != ']') i++; if (i < s.Length) i++; }
                rows.Add(cells);
                cells = new List<string>();
                continue;
            }
            if (AtCmd("end"))
            {
                rows.Add(cells);
                i += 4;
                RawGroup();
                break;
            }
            throw new OmmlFail("unexpected text inside \\begin{" + env + "}");
        }

        bool cases = env == "cases" || env == "dcases";
        bool strip = env == "pmatrix" || env == "bmatrix" || env == "Bmatrix" || env == "vmatrix" || env == "Vmatrix";
        int columns = 0;
        foreach (List<string> row in rows) if (row.Count > columns) columns = row.Count;
        if (columns < 1) columns = 1;

        if (env == "aligned" || env == "align" || env == "align*" || env == "gather" || env == "gather*"
            || env == "split" || env == "eqnarray" || env == "eqnarray*" || env == "array")
        {
            // one equation per row, no columns: the alignment mark is a real gap
            StringBuilder arr = new StringBuilder();
            arr.Append("<m:eqArr><m:eqArrPr><m:maxDist m:val=\"1\"/></m:eqArrPr>");
            foreach (List<string> row in rows)
            {
                StringBuilder line = new StringBuilder();
                for (int k = 0; k < row.Count; k++)
                {
                    if (k > 0) line.Append(Run("\u2003"));
                    line.Append(row[k]);
                }
                arr.Append("<m:e>").Append(line).Append("</m:e>");
            }
            arr.Append("</m:eqArr>");
            return arr.ToString();
        }

        StringBuilder mcs = new StringBuilder();
        for (int c = 0; c < columns; c++)
            mcs.Append("<m:mc><m:mcPr><m:count m:val=\"1\"/><m:mcJc m:val=\"")
               .Append(cases && c == 1 ? "left" : "center").Append("\"/></m:mcPr></m:mc>");

        StringBuilder body = new StringBuilder();
        foreach (List<string> row in rows)
        {
            body.Append("<m:mr>");
            for (int c = 0; c < columns; c++)
                body.Append("<m:e>").Append(c < row.Count ? row[c] : "").Append("</m:e>");
            body.Append("</m:mr>");
        }
        string matrix = "<m:m><m:mPr><m:mcs>" + mcs + "</m:mcs></m:mPr>" + body + "</m:m>";

        if (cases)
            return "<m:d><m:dPr><m:begChr m:val=\"{\"/><m:endChr m:val=\"\"/></m:dPr><m:e>" + matrix + "</m:e></m:d>";
        if (!strip) return matrix;
        string beg, end;
        if (env == "pmatrix") { beg = "("; end = ")"; }
        else if (env == "bmatrix") { beg = "["; end = "]"; }
        else if (env == "Bmatrix") { beg = "{"; end = "}"; }
        else if (env == "vmatrix") { beg = "|"; end = "|"; }
        else { beg = "\u2016"; end = "\u2016"; }
        return "<m:d><m:dPr><m:begChr m:val=\"" + Omml.Esc(beg) + "\"/><m:endChr m:val=\""
            + Omml.Esc(end) + "\"/></m:dPr><m:e>" + matrix + "</m:e></m:d>";
    }
}
