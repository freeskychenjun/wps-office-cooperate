// wps-office-cooperate (ppt): `check` - the pre-delivery gate, and `outline` - the
// "what does this deck actually say" export that feeds an AI-written script.
//
// `audit` answers "where might something be wrong" (a clue list, high false-positive
// rate on hand-made decks).  `check` answers "may this be delivered": every finding
// carries a rule name and a severity, the thresholds come from a configurable rule
// library (defaults = the Chinese reporting norms in the feature note), and a
// suppression list keeps known design intent out of the count without hiding it.
//
// Read-only: nothing here writes to the presentation.  That is what makes it safe to
// run on every deck before a review meeting.
//
// Pure ASCII source (Chinese only ever travels inside JSON strings), C# 5.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing.Text;

internal static partial class Program
{
    // severity order: the higher the number, the more it blocks a delivery
    private const int SevInfo = 0;
    private const int SevMinor = 1;
    private const int SevMajor = 2;
    private const int SevBlocker = 3;

    private sealed class RuleThresholds
    {
        public int TitleMin = 28;
        public int BodyMin = 14;
        public int FootMin = 10;
        public int DensityChars = 300;
        public int DensityShapes = 40;
        public double ContrastMin = 4.5;
        public double OverflowFactor = 1.05;
        public int ImageWidthPct = 20;
    }

    private sealed class SuppressRule
    {
        public int Slide;          // 0 = any slide
        public string Shape;       // null = any shape
        public string Rule;        // null = any rule
    }

    private sealed class ShapeProbe
    {
        public int Slide;
        public int Index;
        public string Name = "";
        public int Type;
        public int PhType;
        public bool IsTitle;
        public double L, T, W, H;
        public string Text = "";
        public double FontSize;      // -1 = mixed / unknown
        public string FontName;
        public string FontEa;
        public double BoundHeight;
        public bool HasFill;
        public int FillRgb;
        public int TextRgb = -1;
        public bool IsPicture;
        public string Role = "body";  // title | body | footnote
    }

    private sealed class SlideProbe
    {
        public int No;
        public List<ShapeProbe> Shapes = new List<ShapeProbe>();
        public bool SlideNumberVisible;
        public int Chars;
        public bool HasPicture;
    }

    // ---------------------------------------------------------------- check

    private static object CheckCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int slideCount = 0;
        try { slideCount = (int)p.Slides.Count; } catch (Exception) { }
        if (slideCount == 0) throw new Fail("BAD_ARGS", "this presentation has no slides");

        double slideW = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0), CultureInfo.InvariantCulture);
        double slideH = Convert.ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0), CultureInfo.InvariantCulture);

        string profile = a.Str("profile", "\u6c47\u62a5");   // 汇报
        List<SuppressRule> suppress;
        RuleThresholds th = ProfileThresholds(profile, a.Str("config", null), out suppress);
        HashSet<string> only = ParseRuleList(a.Str("rules", null));
        int minSeverity = ParseSeverity(a.Str("minSeverity", null), SevInfo);

        List<SlideProbe> deck = new List<SlideProbe>();
        for (int k = 1; k <= slideCount; k++)
        {
            dynamic s = null;
            try { s = p.Slides.Item(k); } catch (Exception) { continue; }
            deck.Add(ProbeSlide(s, k));
        }

        List<object> issues = new List<object>();
        List<object> hidden = new List<object>();

        foreach (SlideProbe slide in deck)
        {
            bool isCover = slide.No == 1;
            foreach (ShapeProbe sh in slide.Shapes)
            {
                Add(issues, slide, sh, "outframe", SevBlocker, only, minSeverity,
                    OutOfFrameDetail(sh, slideW, slideH), "move it back inside the page (align/zorder, or drag)");
                Add(issues, slide, sh, "overflow", SevBlocker, only, minSeverity,
                    OverflowDetail(sh, th), "make the box taller, drop a size, or cut words");
                Add(issues, slide, sh, "placeholder", SevBlocker, only, minSeverity,
                    PlaceholderDetail(sh), "delete the placeholder or fill it in");
                Add(issues, slide, sh, "fontsize", SevMajor, only, minSeverity,
                    FontSizeDetail(sh, th), "raise it (format -Size) or move the text off this slide");
                Add(issues, slide, sh, "contrast", SevMinor, only, minSeverity,
                    ContrastDetail(sh, th), "darken the text or lighten the fill");
                Add(issues, slide, sh, "image-small", SevMinor, only, minSeverity,
                    ImageSmallDetail(sh, slideW, th), "enlarge the picture or move its lettering to text");
            }

            // per-slide rules
            if (slide.Chars == 0 && !slide.HasPicture)
                Add(issues, slide, null, "empty", SevMajor, only, minSeverity,
                    "no text and no picture on this slide", "delete the slide or put something useful on it");
            if (slide.Chars > th.DensityChars)
                Add(issues, slide, null, "density", SevMinor, only, minSeverity,
                    slide.Chars + " characters on one slide (limit " + th.DensityChars + ")", "split it: one message per slide");
            if (slide.Shapes.Count > th.DensityShapes)
                Add(issues, slide, null, "density", SevMinor, only, minSeverity,
                    slide.Shapes.Count + " shapes on one slide (limit " + th.DensityShapes + ")", "group or drop the decorative shapes");
            if (isCover)
            {
                if (slide.SlideNumberVisible)
                    Add(issues, slide, null, "pagenum", SevMajor, only, minSeverity,
                        "the cover shows a slide number", "turn HeadersFooters.SlideNumber off on slide 1");
            }
            else if (!slide.SlideNumberVisible)
                Add(issues, slide, null, "pagenum", SevMajor, only, minSeverity,
                    "no visible slide number", "forward with theme -Apply, or set HeadersFooters.SlideNumber.Visible");
        }

        AddTitleConsistency(issues, deck, only, minSeverity);

        // suppression runs last, so the report can list what was hidden
        List<object> kept = new List<object>();
        foreach (object o in issues)
        {
            Dictionary<string, object> issue = (Dictionary<string, object>)o;
            SuppressRule hit = MatchSuppression(issue, suppress);
            if (hit == null) kept.Add(issue);
            else
            {
                issue["suppressedBy"] = SuppressionLabel(hit);
                hidden.Add(issue);
            }
        }

        Dictionary<string, object> summary = new Dictionary<string, object>();
        int[] counts = new int[4];
        foreach (object o in kept)
        {
            Dictionary<string, object> issue = (Dictionary<string, object>)o;
            int sev = SeverityOf(Convert.ToString(issue["severity"], CultureInfo.InvariantCulture));
            if (sev >= 0 && sev < 4) counts[sev]++;
        }
        summary["blocker"] = counts[SevBlocker];
        summary["major"] = counts[SevMajor];
        summary["minor"] = counts[SevMinor];
        summary["info"] = counts[SevInfo];
        summary["total"] = kept.Count;
        summary["suppressed"] = hidden.Count;

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["profile"] = profile;
        r["slides"] = slideCount;
        r["rules"] = new List<object>(new List<string>(only).ToArray());
        r["summary"] = summary;
        r["fonts"] = FontInventory(deck);
        r["limits"] = new List<object>(new object[] { FontLimitNote });
        r["issues"] = kept;
        if (hidden.Count > 0) r["suppressed"] = hidden;
        r["verdict"] = counts[SevBlocker] > 0 ? "not ready: fix the blockers first"
                       : counts[SevMajor] > 0 ? "readable but rough: the majors are what a reviewer notices"
                       : "mechanically clean";
        r["note"] = "mechanical gate only: it cannot judge whether the content is right or the design is good. "
                  + "Render the flagged slides (and -Overview) before telling anyone it is ready.";

        string outPath = a.Str("out", null);
        if (outPath != null && outPath.Length > 0)
        {
            string written = WriteCheckReport(outPath, profile, summary, kept, hidden, slideCount);
            r["out"] = written;
            if (written == null) throw new Fail("BAD_ARGS", "could not write the report to " + outPath);
        }
        return r;
    }

    /// <summary>Every typeface the deck asks for, and whether this machine has it. A
    /// missing one is why "it looked fine on my laptop" becomes "it broke on the
    /// projector": the substitute font is wider, so the text wraps differently.</summary>
    private static List<object> FontInventory(List<SlideProbe> deck)
    {
        Dictionary<string, int> uses = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (SlideProbe slide in deck)
        {
            foreach (ShapeProbe sh in slide.Shapes)
            {
                Bucket(uses, sh.FontName);
                Bucket(uses, sh.FontEa);
            }
        }
        List<object> list = new List<object>();
        foreach (KeyValuePair<string, int> kv in uses)
        {
            list.Add(new Dictionary<string, object> {
                { "name", kv.Key }, { "installed", FontInstalled(kv.Key) }, { "shapes", kv.Value }
            });
        }
        return list;
    }

    private static void Bucket(Dictionary<string, int> uses, string name)
    {
        if (name == null) return;
        string key = name.Trim();
        if (key.Length == 0 || key.StartsWith("+", StringComparison.Ordinal)) return;
        if (!uses.ContainsKey(key)) uses[key] = 0;
        uses[key]++;
    }

    private static void Add(List<object> issues, SlideProbe slide, ShapeProbe sh, string rule, int severity,
        HashSet<string> only, int minSeverity, string detail, string fix)
    {
        if (detail == null) return;
        if (!RuleOn(only, rule)) return;
        if (severity < minSeverity) return;
        Dictionary<string, object> issue = new Dictionary<string, object>();
        issue["slide"] = slide.No;
        issue["shape"] = sh == null ? "-" : sh.Name;
        issue["rule"] = rule;
        issue["severity"] = SeverityName(severity);
        issue["detail"] = detail;
        issue["fix"] = fix;
        issues.Add(issue);
    }

    // ---------------------------------------------------------------- rule details

    private static string OverflowDetail(ShapeProbe sh, RuleThresholds th)
    {
        if (sh.Text.Length == 0 || sh.BoundHeight <= 0 || sh.H <= 0) return null;
        if (sh.BoundHeight <= sh.H * th.OverflowFactor + 2) return null;
        return "text needs " + R(sh.BoundHeight) + "pt but the box is " + R(sh.H) + "pt tall ("
             + R((sh.BoundHeight - sh.H) / sh.H * 100) + "% over)";
    }

    /// <summary>A 2pt tolerance: shapes are dragged by hand and a background band that
    /// bleeds 1pt past the edge is a design choice, not a defect.</summary>
    private static string OutOfFrameDetail(ShapeProbe sh, double slideW, double slideH)
    {
        if (sh.W <= 0 || sh.H <= 0) return null;
        double over = 0;
        string edge = null;
        if (sh.L < -2) { over = -sh.L; edge = "left"; }
        if (sh.T < -2 && -sh.T > over) { over = -sh.T; edge = "top"; }
        if (slideW > 0 && sh.L + sh.W > slideW + 2 && sh.L + sh.W - slideW > over) { over = sh.L + sh.W - slideW; edge = "right"; }
        if (slideH > 0 && sh.T + sh.H > slideH + 2 && sh.T + sh.H - slideH > over) { over = sh.T + sh.H - slideH; edge = "bottom"; }
        if (edge == null) return null;
        return "shape runs " + R(over) + "pt past the " + edge + " edge ("
             + R(sh.L) + "," + R(sh.T) + " " + R(sh.W) + "x" + R(sh.H) + ")";
    }

    private static readonly string[] PlaceholderHints = new string[]
    {
        "\u5355\u51fb\u6b64\u5904\u6dfb\u52a0",   // 单击此处添加
        "Click to add",
        "\u5355\u51fb\u6b64\u5904\u7f16\u8f91"    // 单击此处编辑
    };

    private static string PlaceholderDetail(ShapeProbe sh)
    {
        if (sh.Text.Length == 0) return null;
        foreach (string hint in PlaceholderHints)
            if (sh.Text.IndexOf(hint, StringComparison.OrdinalIgnoreCase) >= 0)
                return "leftover placeholder text: \"" + Clip(sh.Text, 40) + "\"";
        return null;
    }

    private static string FontSizeDetail(ShapeProbe sh, RuleThresholds th)
    {
        if (sh.Text.Trim().Length == 0 || sh.FontSize <= 0) return null;
        int floor = sh.Role == "title" ? th.TitleMin : sh.Role == "footnote" ? th.FootMin : th.BodyMin;
        if (sh.FontSize >= floor) return null;
        return R(sh.FontSize) + "pt " + sh.Role + " text (floor " + floor + "pt)";
    }

    private static HashSet<string> _installedFonts;

    private static bool FontInstalled(string name)
    {
        if (name == null || name.Length == 0) return true;
        if (name.StartsWith("+", StringComparison.Ordinal)) return true;   // theme font reference (minor latin/east asian)
        if (_installedFonts == null)
        {
            _installedFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (InstalledFontCollection col = new InstalledFontCollection())
                {
                    foreach (System.Drawing.FontFamily f in col.Families) _installedFonts.Add(f.Name);
                }
            }
            catch (Exception) { }
        }
        if (_installedFonts.Count == 0) return true;    // could not enumerate: never cry wolf
        return _installedFonts.Contains(name);
    }

    private static string FontFallbackDetail(ShapeProbe sh)
    {
        if (sh.Text.Trim().Length == 0) return null;
        if (!FontInstalled(sh.FontName)) return "font \"" + sh.FontName + "\" is not installed here (it will be substituted on another machine)";
        if (!FontInstalled(sh.FontEa)) return "east-asian font \"" + sh.FontEa + "\" is not installed here";
        return null;
    }

    /// <summary>Why there is no automatic "missing font" rule: measured on WPS
    /// 12.1.0.29161 - a deck whose run declares a typeface this machine does not have
    /// reads back through COM as the SUBSTITUTE (a fixture with
    /// typeface="Definitely-Not-Installed-Font" reported "Arial").  The information is
    /// gone before we can see it, so reporting the font inventory is all that is honest.</summary>
    private const string FontLimitNote =
        "missing-font detection is not possible through COM on WPS: an unknown typeface is "
        + "resolved to an installed one on load, so the deck's real font list is invisible. "
        + "`fonts` reports what WPS resolved to; open the deck on the target machine to be sure.";

    private static string ContrastDetail(ShapeProbe sh, RuleThresholds th)
    {
        if (sh.Text.Trim().Length == 0) return null;
        if (!sh.HasFill || sh.TextRgb < 0 || sh.FillRgb < 0) return null;
        double ratio = ContrastRatio(sh.TextRgb, sh.FillRgb);
        if (ratio >= th.ContrastMin) return null;
        return "text/background contrast " + Math.Round(ratio, 2) + ":1 (floor " + th.ContrastMin + ":1)";
    }

    private static string ImageSmallDetail(ShapeProbe sh, double slideW, RuleThresholds th)
    {
        if (!sh.IsPicture || sh.W <= 0 || slideW <= 0) return null;
        double pct = sh.W / slideW * 100;
        if (pct >= th.ImageWidthPct) return null;
        return "picture is only " + R(pct) + "% of the slide width";
    }

    /// <summary>WCAG relative luminance contrast.</summary>
    private static double ContrastRatio(int rgbA, int rgbB)
    {
        double a = RelativeLuminance(rgbA);
        double b = RelativeLuminance(rgbB);
        double hi = Math.Max(a, b), lo = Math.Min(a, b);
        return (hi + 0.05) / (lo + 0.05);
    }

    private static double RelativeLuminance(int rgb)
    {
        int r = (rgb >> 0) & 0xFF, g = (rgb >> 8) & 0xFF, b = (rgb >> 16) & 0xFF;   // OLE order: BGR
        return 0.2126 * Channel(r) + 0.7152 * Channel(g) + 0.0722 * Channel(b);
    }

    private static double Channel(int v)
    {
        double c = v / 255.0;
        return c <= 0.03928 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
    }

    // ---------------------------------------------------------------- title consistency

    private static void AddTitleConsistency(List<object> issues, List<SlideProbe> deck, HashSet<string> only, int minSeverity)
    {
        if (!RuleOn(only, "title")) return;
        if (minSeverity > SevMajor) return;

        Dictionary<double, int> sizes = new Dictionary<double, int>();
        foreach (SlideProbe slide in deck)
        {
            double size = TitleSize(slide);
            if (size <= 0) continue;
            if (!sizes.ContainsKey(size)) sizes[size] = 0;
            sizes[size]++;
        }
        if (sizes.Count <= 1) return;

        // the most common title size wins; a slide that departs from it by more than 4pt
        // looks like it came from a different deck
        double modal = 0;
        int best = 0;
        foreach (KeyValuePair<double, int> kv in sizes)
            if (kv.Value > best) { best = kv.Value; modal = kv.Key; }
        if (best < 2) return;

        int flagged = 0;
        foreach (SlideProbe slide in deck)
        {
            if (flagged >= 8) break;
            double size = TitleSize(slide);
            if (size <= 0) continue;
            if (Math.Abs(size - modal) <= 4) continue;
            Dictionary<string, object> issue = new Dictionary<string, object>();
            issue["slide"] = slide.No;
            issue["shape"] = "-";
            issue["rule"] = "title";
            issue["severity"] = SeverityName(SevMajor);
            issue["detail"] = "title is " + R(size) + "pt while the rest of the deck uses " + R(modal) + "pt";
            issue["fix"] = "format -Size " + R(modal) + " on this title, or accept the variation on purpose";
            issues.Add(issue);
            flagged++;
        }
    }

    private static double TitleSize(SlideProbe slide)
    {
        foreach (ShapeProbe sh in slide.Shapes)
            if (sh.IsTitle && sh.FontSize > 0) return sh.FontSize;
        return -1;
    }

    // ---------------------------------------------------------------- probing

    private static SlideProbe ProbeSlide(dynamic slide, int no)
    {
        SlideProbe probe = new SlideProbe();
        probe.No = no;
        try { probe.SlideNumberVisible = (int)slide.HeadersFooters.SlideNumber.Visible != 0; }
        catch (Exception) { probe.SlideNumberVisible = false; }

        int count = 0;
        try { count = (int)slide.Shapes.Count; } catch (Exception) { }
        for (int i = 1; i <= count; i++)
        {
            dynamic sh = null;
            try { sh = slide.Shapes.Item(i); } catch (Exception) { continue; }
            ShapeProbe sp = ProbeShape(sh, no, i);
            if (sp == null) continue;
            probe.Shapes.Add(sp);
            // Furniture must not make an otherwise blank slide look occupied: with the
            // slide number switched on WPS adds a numbered placeholder whose text is a
            // field, and counting that would hide a genuinely empty slide.
            if (sp.Role != "footnote" && sp.PhType != 13 && sp.PhType != 14 && sp.PhType != 15 && sp.PhType != 16)
            {
                probe.Chars += sp.Text.Length;
                if (sp.IsPicture) probe.HasPicture = true;
            }
        }
        return probe;
    }

    private static ShapeProbe ProbeShape(dynamic sh, int slideNo, int index)
    {
        ShapeProbe sp = new ShapeProbe();
        sp.Slide = slideNo;
        sp.Index = index;
        sp.Name = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
        sp.Type = Convert.ToInt32(TryGet(() => (object)(int)sh.Type, (object)0), CultureInfo.InvariantCulture);
        sp.L = Convert.ToDouble(TryGet(() => (object)sh.Left, (object)0.0), CultureInfo.InvariantCulture);
        sp.T = Convert.ToDouble(TryGet(() => (object)sh.Top, (object)0.0), CultureInfo.InvariantCulture);
        sp.W = Convert.ToDouble(TryGet(() => (object)sh.Width, (object)0.0), CultureInfo.InvariantCulture);
        sp.H = Convert.ToDouble(TryGet(() => (object)sh.Height, (object)0.0), CultureInfo.InvariantCulture);
        sp.Text = ShapeText(sh);
        sp.IsPicture = (sp.Type == 13 || sp.Type == 11);

        bool isPlaceholder = false;
        try { isPlaceholder = (int)sh.IsPlaceholder != 0; } catch (Exception) { }
        if (isPlaceholder) sp.PhType = Convert.ToInt32(TryGet(() => (object)sh.PlaceholderFormat.Type, (object)0), CultureInfo.InvariantCulture);
        // A title is either a real title/centre-title placeholder or a shape this skill's
        // presets named "title"/"sectionTitle" (hand-built decks often use plain boxes).
        sp.IsTitle = (isPlaceholder && (sp.PhType == 1 || sp.PhType == 3))
                     || string.Equals(sp.Name, "title", StringComparison.OrdinalIgnoreCase)
                     || string.Equals(sp.Name, "sectionTitle", StringComparison.OrdinalIgnoreCase);

        if (sp.Text.Trim().Length > 0)
        {
            try { sp.BoundHeight = Convert.ToDouble(TryGet(() => (object)sh.TextFrame.TextRange.BoundHeight, (object)0.0), CultureInfo.InvariantCulture); }
            catch (Exception) { sp.BoundHeight = 0; }
            sp.FontSize = ReadFontSize(sh);
            sp.FontName = Convert.ToString(TryGet(() => (object)sh.TextFrame.TextRange.Font.Name, null), CultureInfo.InvariantCulture);
            sp.FontEa = Convert.ToString(TryGet(() => (object)sh.TextFrame.TextRange.Font.NameFarEast, null), CultureInfo.InvariantCulture);
            object rgb = TryGet(() => (object)sh.TextFrame.TextRange.Font.Color.RGB, null);
            sp.TextRgb = rgb == null ? -1 : Convert.ToInt32(rgb, CultureInfo.InvariantCulture);
        }

        try
        {
            if ((int)sh.Fill.Visible != 0)
            {
                sp.HasFill = true;
                object rgb = TryGet(() => (object)sh.Fill.ForeColor.RGB, null);
                sp.FillRgb = rgb == null ? -1 : Convert.ToInt32(rgb, CultureInfo.InvariantCulture);
            }
        }
        catch (Exception) { sp.HasFill = false; }

        sp.Role = sp.IsTitle ? "title" : IsFurniture(sp.Name) ? "footnote" : "body";
        return sp;
    }

    private static bool IsFurniture(string name)
    {
        if (name == null) return false;
        return name == "footer" || name == "pageNo" || name == "date"
            || name.IndexOf("\u9875\u811a", StringComparison.Ordinal) >= 0        // 页脚
            || name.IndexOf("\u7f16\u53f7", StringComparison.Ordinal) >= 0;       // 编号
    }

    /// <summary>Font.Size of a range is 9999999 when the range mixes sizes; fall back to
    /// the first character, which is what the reader's eye sees first anyway.</summary>
    private static double ReadFontSize(dynamic sh)
    {
        double size = Convert.ToDouble(TryGet(() => (object)sh.TextFrame.TextRange.Font.Size, (object)0.0), CultureInfo.InvariantCulture);
        if (size > 0 && size < 4000) return size;
        double first = Convert.ToDouble(TryGet(() => (object)sh.TextFrame.TextRange.Characters(1, 1).Font.Size, (object)0.0), CultureInfo.InvariantCulture);
        return first > 0 && first < 4000 ? first : -1;
    }

    // ---------------------------------------------------------------- config / severity

    private static RuleThresholds ProfileThresholds(string profile, string configPath, out List<SuppressRule> suppress)
    {
        RuleThresholds th = new RuleThresholds();
        suppress = new List<SuppressRule>();
        // ASCII aliases exist because a Chinese argument does not survive the trip
        // through `powershell -File` on this machine (PS 5.1 decodes the command line as
        // ANSI, so "投标" arrives as mojibake).  Anything Chinese should travel in JSON.
        string key = profile ?? "";
        string lower = key.ToLowerInvariant();
        if (key.IndexOf("\u8bc4\u5ba1", StringComparison.Ordinal) >= 0 || lower == "review" || lower == "rehearsal")
        {
            th.BodyMin = 16; th.FootMin = 12;                                  // 评审
        }
        else if (key.IndexOf("\u6295\u6807", StringComparison.Ordinal) >= 0 || lower == "bid" || lower == "tender")
        {
            th.TitleMin = 32; th.BodyMin = 18; th.FootMin = 12;                // 投标
        }
        else if (key.IndexOf("\u5185\u90e8", StringComparison.Ordinal) >= 0 || lower == "internal")
        {
            th.TitleMin = 24; th.BodyMin = 12; th.FootMin = 9;                 // 内部
        }
        else if (key.IndexOf("\u6c47\u62a5", StringComparison.Ordinal) >= 0 || lower == "report" || lower.Length == 0)
        {
            // 汇报 = the default (28/14/10)
        }
        else
        {
            throw new Fail("BAD_ARGS", "unknown -Profile \"" + profile
                + "\"; use report|review|bid|internal (ASCII, command-line safe) or the Chinese names inside a JSON call");
        }

        if (configPath == null || configPath.Length == 0) return th;
        if (!File.Exists(configPath)) throw new Fail("BAD_ARGS", "config file not found: " + configPath);
        object parsed;
        try { parsed = Json.Parse(File.ReadAllText(configPath, Encoding.UTF8)); }
        catch (Exception e) { throw new Fail("BAD_ARGS", "config is not valid JSON: " + Clean(e.Message)); }
        Bag cfg = new Bag(parsed);

        RuleThresholds fromFile = new RuleThresholds();
        th.TitleMin = cfg.Int("titleMin", th.TitleMin);
        th.BodyMin = cfg.Int("bodyMin", th.BodyMin);
        th.FootMin = cfg.Int("footMin", th.FootMin);
        th.DensityChars = cfg.Int("densityChars", th.DensityChars);
        th.DensityShapes = cfg.Int("densityShapes", th.DensityShapes);
        th.ContrastMin = cfg.Dbl("contrastMin", th.ContrastMin);
        th.OverflowFactor = cfg.Dbl("overflowFactor", th.OverflowFactor);
        th.ImageWidthPct = cfg.Int("imageWidthPct", th.ImageWidthPct);

        foreach (object o in cfg.List("suppress"))
        {
            Bag entry = new Bag(o);
            SuppressRule rule = new SuppressRule();
            rule.Slide = entry.Int("slide", 0);
            rule.Shape = entry.Str("shape", null);
            rule.Rule = entry.Str("rule", null);
            suppress.Add(rule);
        }
        return th;
    }

    private static HashSet<string> ParseRuleList(string spec)
    {
        HashSet<string> want = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        string[] all = new string[] { "outframe", "overflow", "placeholder", "fontsize",
                                      "contrast", "image-small", "empty", "density", "pagenum", "title" };
        if (spec == null || spec.Trim().Length == 0)
        {
            foreach (string name in all) want.Add(name);
            return want;
        }
        foreach (string raw in spec.Split(','))
        {
            string v = raw.Trim().ToLowerInvariant();
            if (v.Length == 0) continue;
            if (v == "all") { foreach (string name in all) want.Add(name); continue; }
            bool known = false;
            foreach (string name in all) if (name == v) { known = true; break; }
            if (!known) throw new Fail("BAD_ARGS", "unknown rule \"" + raw.Trim() + "\"; use all|" + string.Join("|", all));
            want.Add(v);
        }
        return want;
    }

    private static bool RuleOn(HashSet<string> only, string rule)
    {
        return only.Count == 0 || only.Contains(rule);
    }

    private static int ParseSeverity(string name, int fallback)
    {
        if (name == null || name.Trim().Length == 0) return fallback;
        switch (name.Trim().ToLowerInvariant())
        {
            case "info": return SevInfo;
            case "minor": return SevMinor;
            case "major": return SevMajor;
            case "blocker": return SevBlocker;
            default: throw new Fail("BAD_ARGS", "unknown -MinSeverity \"" + name + "\"; use info|minor|major|blocker");
        }
    }

    private static string SeverityName(int severity)
    {
        switch (severity)
        {
            case SevBlocker: return "blocker";
            case SevMajor: return "major";
            case SevMinor: return "minor";
            default: return "info";
        }
    }

    private static int SeverityOf(string name)
    {
        switch (name)
        {
            case "blocker": return SevBlocker;
            case "major": return SevMajor;
            case "minor": return SevMinor;
            case "info": return SevInfo;
            default: return -1;
        }
    }

    private static SuppressRule MatchSuppression(Dictionary<string, object> issue, List<SuppressRule> suppress)
    {
        if (suppress.Count == 0) return null;
        int slide = Convert.ToInt32(issue["slide"], CultureInfo.InvariantCulture);
        string shape = Convert.ToString(issue["shape"], CultureInfo.InvariantCulture);
        string rule = Convert.ToString(issue["rule"], CultureInfo.InvariantCulture);
        foreach (SuppressRule s in suppress)
        {
            if (s.Slide != 0 && s.Slide != slide) continue;
            if (s.Shape != null && !string.Equals(s.Shape, shape, StringComparison.OrdinalIgnoreCase)) continue;
            if (s.Rule != null && !string.Equals(s.Rule, rule, StringComparison.OrdinalIgnoreCase)) continue;
            return s;
        }
        return null;
    }

    private static string SuppressionLabel(SuppressRule rule)
    {
        List<string> parts = new List<string>();
        if (rule.Slide != 0) parts.Add("slide " + rule.Slide);
        if (rule.Shape != null) parts.Add("shape " + rule.Shape);
        if (rule.Rule != null) parts.Add("rule " + rule.Rule);
        return parts.Count == 0 ? "config" : "config: " + string.Join(", ", parts.ToArray());
    }

    // ---------------------------------------------------------------- report

    private static string WriteCheckReport(string path, string profile, Dictionary<string, object> summary,
        List<object> issues, List<object> hidden, int slideCount)
    {
        string full = FullPathOf(path);
        try
        {
            string dir = Path.GetDirectoryName(full);
            if (dir != null && dir.Length > 0) Directory.CreateDirectory(dir);
        }
        catch (Exception) { }

        bool markdown = full.EndsWith(".md", StringComparison.OrdinalIgnoreCase);
        StringBuilder sb = new StringBuilder();
        if (markdown)
        {
            sb.Append("# \u4ea4\u4ed8\u524d\u68c0\u67e5\u62a5\u544a\n\n");     // 交付前检查报告
            sb.Append("- profile: ").Append(profile).Append("\n");
            sb.Append("- slides: ").Append(slideCount).Append("\n");
            sb.Append("- at: ").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)).Append("\n\n");
            sb.Append("| severity | count |\n|---|---|\n");
            foreach (string name in new string[] { "blocker", "major", "minor", "info", "suppressed" })
                sb.Append("| ").Append(name).Append(" | ").Append(Convert.ToString(summary[name], CultureInfo.InvariantCulture)).Append(" |\n");
            string[] order = new string[] { "blocker", "major", "minor", "info" };
            foreach (string name in order)
            {
                List<object> bucket = new List<object>();
                foreach (object o in issues)
                    if (Convert.ToString(((Dictionary<string, object>)o)["severity"], CultureInfo.InvariantCulture) == name)
                        bucket.Add(o);
                if (bucket.Count == 0) continue;
                sb.Append("\n## ").Append(name).Append(" (").Append(bucket.Count).Append(")\n\n");
                sb.Append("| slide | shape | rule | detail | fix |\n|---|---|---|---|---|\n");
                foreach (object o in bucket)
                {
                    Dictionary<string, object> issue = (Dictionary<string, object>)o;
                    sb.Append("| ").Append(Convert.ToString(issue["slide"], CultureInfo.InvariantCulture))
                      .Append(" | ").Append(Convert.ToString(issue["shape"], CultureInfo.InvariantCulture))
                      .Append(" | ").Append(Convert.ToString(issue["rule"], CultureInfo.InvariantCulture))
                      .Append(" | ").Append(Convert.ToString(issue["detail"], CultureInfo.InvariantCulture))
                      .Append(" | ").Append(Convert.ToString(issue["fix"], CultureInfo.InvariantCulture))
                      .Append(" |\n");
                }
            }
            if (hidden.Count > 0)
            {
                sb.Append("\n## suppressed (").Append(hidden.Count).Append(")\n\n");
                sb.Append("| slide | shape | rule | detail | why |\n|---|---|---|---|---|\n");
                foreach (object o in hidden)
                {
                    Dictionary<string, object> issue = (Dictionary<string, object>)o;
                    sb.Append("| ").Append(Convert.ToString(issue["slide"], CultureInfo.InvariantCulture))
                      .Append(" | ").Append(Convert.ToString(issue["shape"], CultureInfo.InvariantCulture))
                      .Append(" | ").Append(Convert.ToString(issue["rule"], CultureInfo.InvariantCulture))
                      .Append(" | ").Append(Convert.ToString(issue["detail"], CultureInfo.InvariantCulture))
                      .Append(" | ").Append(Convert.ToString(issue["suppressedBy"], CultureInfo.InvariantCulture))
                      .Append(" |\n");
                }
            }
        }
        else
        {
            Dictionary<string, object> doc = new Dictionary<string, object>();
            doc["profile"] = profile;
            doc["slides"] = slideCount;
            doc["at"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            doc["summary"] = summary;
            doc["issues"] = issues;
            if (hidden.Count > 0) doc["suppressed"] = hidden;
            sb.Append(Json.Write(doc));
        }

        try
        {
            File.WriteAllText(full, sb.ToString(), new UTF8Encoding(false));
            return full;
        }
        catch (Exception) { return null; }
    }

    // ---------------------------------------------------------------- outline

    private static object OutlineCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }

        List<object> slides = new List<object>();
        StringBuilder md = new StringBuilder();
        for (int k = 1; k <= n; k++)
        {
            dynamic s = null;
            try { s = p.Slides.Item(k); } catch (Exception) { continue; }
            List<ShapeProbe> shapes = new List<ShapeProbe>();
            int sc = 0;
            try { sc = (int)s.Shapes.Count; } catch (Exception) { }
            for (int i = 1; i <= sc; i++)
            {
                dynamic sh = null;
                try { sh = s.Shapes.Item(i); } catch (Exception) { continue; }
                ShapeProbe sp = ProbeShape(sh, k, i);
                if (sp != null) shapes.Add(sp);
            }
            // reading order: top to bottom, then left to right
            shapes.Sort(delegate (ShapeProbe x, ShapeProbe y)
            {
                int byTop = x.T.CompareTo(y.T);
                return byTop != 0 ? byTop : x.L.CompareTo(y.L);
            });

            string title = "";
            List<string> bullets = new List<string>();
            foreach (ShapeProbe sp in shapes)
            {
                if (sp.Text.Trim().Length == 0) continue;
                if (title.Length == 0 && sp.Role == "title") { title = Clean(sp.Text); continue; }
                if (sp.Role == "footnote") continue;
                foreach (string line in sp.Text.Replace("\r", "\n").Split('\n'))
                {
                    string one = Clean(line);
                    if (one.Length > 0) bullets.Add(one);
                }
            }

            string notes = "";
            try { notes = Clean(Convert.ToString(s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text, CultureInfo.InvariantCulture)); }
            catch (Exception) { notes = ""; }

            Dictionary<string, object> item = new Dictionary<string, object>();
            item["slide"] = k;
            item["title"] = title;
            item["bullets"] = new List<object>(bullets.ToArray());
            if (notes.Length > 0) item["notes"] = notes;
            slides.Add(item);

            md.Append("## ").Append(k).Append(". ").Append(title.Length > 0 ? title : "(\u65e0\u6807\u9898)").Append("\n\n");  // (無标题)
            foreach (string b in bullets) md.Append("- ").Append(b).Append("\n");
            if (notes.Length > 0) md.Append("\n> ").Append(notes.Replace("\n", " ")).Append("\n");
            md.Append("\n");
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slides"] = slides;
        r["count"] = slides.Count;
        string outPath = a.Str("out", null);
        if (outPath != null && outPath.Length > 0)
        {
            string full = FullPathOf(outPath);
            try
            {
                string dir = Path.GetDirectoryName(full);
                if (dir != null && dir.Length > 0) Directory.CreateDirectory(dir);
                if (full.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
                    File.WriteAllText(full, md.ToString(), new UTF8Encoding(false));
                else
                    File.WriteAllText(full, Json.Write(r), new UTF8Encoding(false));
                r["out"] = full;
            }
            catch (Exception e) { throw new Fail("BAD_ARGS", "could not write the outline: " + Clean(e.Message)); }
        }
        r["note"] = "titles + bullets + speaker notes only; writing the actual script is the model's job, and `notes` writes it back";
        return r;
    }
}
