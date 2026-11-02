// wps-office-cooperate (ppt): finished-slide layouts, phase 1 (v1.31).
//
// Eight ready-made layouts behind `slide -Preset`: cards, stats, timeline, process,
// compare, summary, gallery, split.  They are laid out on a 960 x 540 design canvas and
// scaled to whatever page size the deck actually has (4:3, 16:9 or a custom size), so
// nothing here is a hard-coded slide size.
//
// Design rules shared by every layout:
//   * content sits inside x 48..912, y 118..494 (the title band owns y 28..100);
//   * every page carries a title, a page number (a PLAIN TEXT BOX: WPS' own slide-number
//     placeholder shows "1" on every page on this machine) and the theme footer;
//   * text is measured with a conservative CJK/Latin estimator and the font is stepped
//     down (18 -> 16 -> 14 ...) rather than left to spill out of its box; when even the
//     smallest step does not fit, an overflow warning is reported instead of pretending;
//   * pictures are scaled by their own aspect ratio: cover a frame (cut the overflow, the
//     cut taken from the top where heads are) or contain (shown whole, centred).
//
// Pure ASCII source, C# 5 (no string interpolation, no null-conditional operator).

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

internal static partial class Program
{
    // ---------------------------------------------------------------- canvas

    /// <summary>The design canvas: everything below is written for 960 x 540 and mapped
    /// through these scales, so a 4:3 or a custom-size deck gets the same composition.</summary>
    private sealed class KindCanvas
    {
        public double W, H, Sx, Sy, S;

        public KindCanvas(double w, double h)
        {
            W = w <= 1 ? 960 : w;
            H = h <= 1 ? 540 : h;
            Sx = W / 960.0;
            Sy = H / 540.0;
            S = Math.Min(Sx, Sy);
            if (S <= 0.05) S = 1;
        }

        public double X(double v) { return v * Sx; }
        public double Y(double v) { return v * Sy; }
        public double WD(double v) { return v * Sx; }
        public double HT(double v) { return v * Sy; }
        public double FS(double v) { return Math.Max(6.0, v * S); }
    }

    // ---------------------------------------------------------------- text metrics

    /// <summary>Advance width of one character in em units: CJK sits on a full em, Latin
    /// letters between half and two thirds, spaces and narrow punctuation much less.</summary>
    private static double KGlyphEm(char c)
    {
        if (c == '\t') return 0.6;
        if (c == ' ') return 0.30;
        if (c < 128)
        {
            if (c >= 'A' && c <= 'Z') return 0.64;
            if (c >= '0' && c <= '9') return 0.56;
            if (".,:;'|!()[]{}iljtfrI".IndexOf(c) >= 0) return 0.30;
            if ("mwMW".IndexOf(c) >= 0) return 0.85;
            return 0.53;
        }
        return 1.0;
    }

    private static List<string> KSplitLines(string text)
    {
        List<string> lines = new List<string>();
        if (text == null) { lines.Add(""); return lines; }
        string t = text.Replace("\r\n", "\n").Replace('\r', '\n');
        string[] parts = t.Split('\n');
        for (int i = 0; i < parts.Length; i++) lines.Add(parts[i]);
        return lines;
    }

    /// <summary>How many typeset lines the text needs in a box this wide.</summary>
    private static int KLineCount(string text, double widthPt, double sizePt)
    {
        if (widthPt <= 1 || sizePt <= 1) return 1;
        double cap = Math.Max(1.0, widthPt / sizePt);   // characters per line, in em
        int total = 0;
        List<string> lines = KSplitLines(text);
        foreach (string line in lines)
        {
            double em = 0;
            foreach (char c in line) em += KGlyphEm(c);
            total += Math.Max(1, (int)Math.Ceiling(em / cap));
        }
        return Math.Max(1, total);
    }

    /// <summary>Height the text occupies: line count * size * line factor.</summary>
    private static double KNeedHeight(string text, double widthPt, double sizePt)
    {
        return KLineCount(text, widthPt, sizePt) * sizePt * 1.32 + 1.5;
    }

    // ---------------------------------------------------------------- colour bits

    private static int KMix(int a, int b, double share)
    {
        if (share <= 0) return a;
        if (share >= 1) return b;
        int ar = a & 0xFF, ag = (a >> 8) & 0xFF, ab = (a >> 16) & 0xFF;
        int br = b & 0xFF, bg = (b >> 8) & 0xFF, bb = (b >> 16) & 0xFF;
        int r = (int)Math.Round(ar + (br - ar) * share);
        int g = (int)Math.Round(ag + (bg - ag) * share);
        int l = (int)Math.Round(ab + (bb - ab) * share);
        if (r < 0) r = 0; if (r > 255) r = 255;
        if (g < 0) g = 0; if (g > 255) g = 255;
        if (l < 0) l = 0; if (l > 255) l = 255;
        return (l << 16) | (g << 8) | r;
    }

    /// <summary>True when a fill is dark enough that text on it must be white.</summary>
    private static bool KIsDark(int ole)
    {
        int r = ole & 0xFF, g = (ole >> 8) & 0xFF, b = (ole >> 16) & 0xFF;
        return (r * 0.299 + g * 0.587 + b * 0.114) < 150;
    }

    /// <summary>The hairline colour every layout draws with (the theme has no line token).</summary>
    private static int KLineColor(DeckTheme th)
    {
        return KMix(th.Muted, th.White, 0.62);
    }

    private static double KSettle(double top, double room, double used)
    {
        double spare = room - used;
        return spare > 24 ? top + spare * 0.42 : top;
    }

    private static double KImageAspect(string path)
    {
        try
        {
            if (path == null || !File.Exists(path)) return 0;
            using (System.Drawing.Image probe = System.Drawing.Image.FromFile(path))
                if (probe.Height > 0) return (double)probe.Width / probe.Height;
        }
        catch (Exception) { }
        return 0;
    }

    private static string KClip(string s, int n)
    {
        if (s == null) return "";
        string t = s.Replace("\r", " ").Replace("\n", " ");
        return t.Length <= n ? t : t.Substring(0, n) + "...";
    }

    // ---------------------------------------------------------------- one designed page

    /// <summary>Everything a layout needs: the slide, the theme, the canvas, the list of
    /// text boxes placed (for the overflow heuristics) and the notes to report back.</summary>
    private sealed class KindPage
    {
        public dynamic Slide;
        public DeckTheme Th;
        public KindCanvas C;
        public int Index = 1, Total = 1;
        public double BodyTop = 118, BodyBottom = 494, Left = 48, Right = 912;
        public double LastSize;
        public List<string> Warnings = new List<string>();
        public List<object> Drops = new List<object>();
        public List<object> Info = new List<object>();

        public double BodyW { get { return Right - Left; } }
        public double BodyH { get { return BodyBottom - BodyTop; } }

        public int CardGround()
        {
            return Th.HasCardFill ? Th.CardFill : Th.Surface;
        }

        public int ToneColor(int i)
        {
            if (i % 2 == 0) return Th.Primary;
            return Th.Accent;
        }

        public int SafeInk(int fill)
        {
            return KIsDark(fill) ? Th.White : Th.Text;
        }

        // ---- primitives (all coordinates in design units, 960 x 540) ----

        public dynamic Shape(int type, double l, double t, double w, double h, int fill, string name)
        {
            dynamic shp = null;
            try { shp = Slide.Shapes.AddShape(type, (float)C.X(l), (float)C.Y(t), (float)Math.Max(0.5, C.WD(w)), (float)Math.Max(0.5, C.HT(h))); }
            catch (Exception e) { throw new Fail("COM_ERROR", "AddShape failed: " + Clean(e.Message)); }
            if (fill >= 0)
            {
                try { shp.Fill.Solid(); shp.Fill.ForeColor.RGB = fill; } catch (Exception) { }
            }
            else
            {
                try { shp.Fill.Visible = 0; } catch (Exception) { }
            }
            try { shp.Line.Visible = 0; } catch (Exception) { }
            if (name != null) { try { shp.Name = name; } catch (Exception) { } }
            return shp;
        }

        /// <summary>A filled block; rounded corners use the theme card radius.</summary>
        public dynamic Block(double l, double t, double w, double h, int fill, string name)
        {
            double rad = Th.CardRadius;
            if (rad <= 0.5) return Shape(1, l, t, w, h, fill, name);
            dynamic shp = Shape(5, l, t, w, h, fill, name);
            double shortest = Math.Min(C.WD(w), C.HT(h));
            double adj = shortest <= 1 ? 0 : Math.Min(0.5, (rad * C.S) / shortest);
            try { shp.Adjustments[1] = (float)adj; } catch (Exception) { }
            return shp;
        }

        /// <summary>An outlined box (no fill): a thin frame in the theme hairline colour.</summary>
        public dynamic Frame(double l, double t, double w, double h, int line, double weight, string name)
        {
            double rad = Th.CardRadius;
            dynamic shp = Shape(rad <= 0.5 ? 1 : 5, l, t, w, h, -1, name);
            try
            {
                shp.Line.Visible = -1;
                shp.Line.ForeColor.RGB = line;
                shp.Line.Weight = (float)Math.Max(0.5, weight * C.S);
            }
            catch (Exception) { }
            if (rad > 0.5)
            {
                double shortest = Math.Min(C.WD(w), C.HT(h));
                double adj = shortest <= 1 ? 0 : Math.Min(0.5, (rad * C.S) / shortest);
                try { shp.Adjustments[1] = (float)adj; } catch (Exception) { }
            }
            return shp;
        }

        public dynamic Oval(double l, double t, double w, double h, int fill, string name)
        {
            return Shape(9, l, t, w, h, fill, name);
        }

        /// <summary>A hairline rule: a filled rectangle 1pt tall reads cleaner than a line shape
        /// on WPS (a zero-height line is invisible in some builds).</summary>
        public dynamic Rule(double l, double t, double w, int color, string name)
        {
            return Shape(1, l, t, w, Math.Max(0.6, 1.0 / Math.Max(C.Sy, 0.2)), color, name);
        }

        // ---- text ----

        /// <summary>Draw a text box.  <paramref name="ladder"/> is the descending list of type
        /// sizes to try (design points); the first one that fits the box is used.  Segments
        /// between ** marks are set bold in the accent colour.</summary>
        public dynamic Text(double l, double t, double w, double h, string text, double size, double[] ladder,
                            bool bold, string align, int color, string font, string name)
        {
            LastSize = size;
            if (text == null || text.Trim().Length == 0) return null;
            double bw = C.WD(w), bh = C.HT(h);
            double chosen = size;
            if (ladder == null) ladder = new double[] { size };
            bool fits = false;
            for (int i = 0; i < ladder.Length; i++)
            {
                double actual = C.FS(ladder[i]);
                if (KNeedHeight(text, bw, actual) <= bh * 1.02) { chosen = ladder[i]; fits = true; break; }
            }
            if (!fits)
            {
                chosen = ladder[ladder.Length - 1];
                double need = KNeedHeight(text, bw, C.FS(chosen));
                Warnings.Add("possible overflow on shape \"" + (name == null ? "text" : name) + "\": \""
                    + KClip(text, 24) + "\" needs about " + Math.Round(need, 0) + "pt in a "
                    + Math.Round(bh, 0) + "pt box even at " + Math.Round(chosen, 1) + "pt");
            }
            if (chosen < size - 0.01)
            {
                Dictionary<string, object> d = new Dictionary<string, object>();
                d["shape"] = name;
                d["from"] = Math.Round(size, 1);
                d["to"] = Math.Round(chosen, 1);
                Drops.Add(d);
            }
            List<int> starts = new List<int>();
            List<int> lengths = new List<int>();
            string shown = StripMarks(text, starts, lengths);

            dynamic tb = null;
            try { tb = Slide.Shapes.AddTextbox(1, (float)C.X(l), (float)C.Y(t), (float)Math.Max(1, bw), (float)Math.Max(1, bh)); }
            catch (Exception e) { throw new Fail("COM_ERROR", "AddTextbox failed: " + Clean(e.Message)); }
            if (name != null) { try { tb.Name = name; } catch (Exception) { } }
            dynamic tf = tb.TextFrame;
            try { tf.WordWrap = -1; } catch (Exception) { }
            try { tf.MarginLeft = 0f; tf.MarginRight = 0f; tf.MarginTop = 0f; tf.MarginBottom = 0f; } catch (Exception) { }
            try { tf.VerticalAnchor = 1; } catch (Exception) { }
            dynamic tr = tf.TextRange;
            try
            {
                tr.Text = shown;
                tr.Font.Size = (float)C.FS(chosen);
                tr.Font.Bold = bold ? -1 : 0;
                tr.Font.Color.RGB = color;
                string fn = font == null ? Th.Font : font;
                tr.Font.Name = fn;
                tr.Font.NameFarEast = fn;
                tr.ParagraphFormat.Alignment = align == "center" ? 2 : align == "right" ? 3 : 1;
            }
            catch (Exception) { }
            for (int i = 0; i < starts.Count; i++)
            {
                try
                {
                    dynamic part = tr.Characters(starts[i], lengths[i]);
                    part.Font.Bold = -1;
                    part.Font.Color.RGB = Th.Accent;
                }
                catch (Exception) { }
            }
            // Measured on WPS 12.1: writing TextFrame.AutoSize (any value) makes the box collapse
            // to a 29pt strip, so the geometry the layout asked for is written back afterwards and
            // AutoSize is left alone.  Without this every text box reads back 29pt tall and the
            // mechanical overflow check compares against the wrong height.
            try
            {
                tb.Left = (float)C.X(l);
                tb.Top = (float)C.Y(t);
                tb.Width = (float)Math.Max(1, bw);
                tb.Height = (float)Math.Max(1, bh);
            }
            catch (Exception) { }
            LastSize = chosen;
            return tb;
        }

        /// <summary>Vertical anchor for a box that has already been created (1 top, 2 middle, 3 bottom).</summary>
        public void Anchor(dynamic tb, int anchor)
        {
            if (tb == null) return;
            try { tb.TextFrame.VerticalAnchor = anchor; } catch (Exception) { }
        }

        // ---- pictures ----

        /// <summary>Place a picture in a frame.  "cover" crops a COPY of the picture to the frame's
        /// aspect ratio in GDI+ and inserts it at the exact frame size; "contain" shows it whole and
        /// centred.  WPS' PictureFormat.Crop* is deliberately not used: measured on WPS 12.1
        /// (2026-11-02) a value written there is read back as points, not as a 0..1 fraction, and a
        /// later Width/Height write throws the crop away - a 1200x600 picture in a 480x540 frame
        /// ended up 552x540 at x = -36 and `check` reported it running off the page.</summary>
        public dynamic Picture(double l, double t, double w, double h, string path, string fit, string name)
        {
            if (path == null || path.Length == 0) throw new Fail("BAD_ARGS", "this layout needs a picture path");
            if (!File.Exists(path)) throw new Fail("BAD_ARGS", "image not found: " + path);
            double bw = C.WD(w), bh = C.HT(h);
            double ar = KImageAspect(path);
            if (ar <= 0) ar = bw / Math.Max(bh, 1);
            bool cover = "cover".Equals(fit, StringComparison.OrdinalIgnoreCase);
            string source = path;
            string cropped = null;
            if (cover)
            {
                cropped = KCoverCrop(path, bw / Math.Max(bh, 1.0));
                if (cropped != null) source = cropped;
            }
            string tmp = AsciiTempCopy(source);
            dynamic pic = null;
            try { pic = Slide.Shapes.AddPicture(tmp, 0, -1, (float)C.X(l), (float)C.Y(t)); }
            catch (Exception e) { throw new Fail("COM_ERROR", "AddPicture failed: " + Clean(e.Message)); }
            finally { TryDelete(tmp); if (cropped != null) TryDelete(cropped); }
            if (name != null) { try { pic.Name = name; } catch (Exception) { } }
            try { pic.LockAspectRatio = 0; } catch (Exception) { }
            double fw, fh, fl, ft;
            if (cover)
            {
                fw = bw; fh = bh; fl = C.X(l); ft = C.Y(t);
            }
            else
            {
                if (ar > bw / Math.Max(bh, 1)) { fw = bw; fh = bw / ar; } else { fh = bh; fw = bh * ar; }
                fl = C.X(l) + (bw - fw) / 2;
                ft = C.Y(t) + (bh - fh) / 2;
            }
            try { pic.Width = (float)fw; pic.Height = (float)fh; } catch (Exception) { }
            try { pic.Left = (float)fl; pic.Top = (float)ft; } catch (Exception) { }
            double aw = fw, ah = fh;
            try
            {
                aw = Convert.ToDouble(pic.Width, CultureInfo.InvariantCulture);
                ah = Convert.ToDouble(pic.Height, CultureInfo.InvariantCulture);
            }
            catch (Exception) { }
            if (Math.Abs(aw - fw) > 6 || Math.Abs(ah - fh) > 6)
                Warnings.Add("picture \"" + (name == null ? "image" : name) + "\" is " + Math.Round(aw, 0) + "x"
                    + Math.Round(ah, 0) + "pt instead of " + Math.Round(fw, 0) + "x" + Math.Round(fh, 0)
                    + "pt - check the render");
            return pic;
        }

        /// <summary>Crop a temporary copy of the picture to the frame's aspect ratio with GDI+, so
        /// that no COM crop is involved.  A too-tall picture keeps its upper part (where heads are),
        /// a too-wide one is cut evenly on both sides.  Returns a temp PNG, or null when the picture
        /// cannot be read (the caller then falls back to a plain scale).</summary>
        private static string KCoverCrop(string path, double targetAspect)
        {
            try
            {
                if (targetAspect <= 0.01) return null;
                using (System.Drawing.Image src = System.Drawing.Image.FromFile(path))
                {
                    int sw = src.Width, sh = src.Height;
                    if (sw < 1 || sh < 1) return null;
                    double have = (double)sw / sh;
                    int cw = sw, ch = sh, cx = 0, cy = 0;
                    if (have > targetAspect)
                    {
                        cw = (int)Math.Round(sh * targetAspect);
                        if (cw < 1) cw = 1;
                        cx = (sw - cw) / 2;
                    }
                    else if (have < targetAspect)
                    {
                        ch = (int)Math.Round(sw / targetAspect);
                        if (ch < 1) ch = 1;
                        cy = (int)Math.Round((sh - ch) * 0.2);
                        if (cy < 0) cy = 0;
                    }
                    if (cw >= sw && ch >= sh) return null;
                    using (System.Drawing.Bitmap dst = new System.Drawing.Bitmap(cw, ch))
                    {
                        using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(dst))
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                            g.DrawImage(src, new System.Drawing.Rectangle(0, 0, cw, ch),
                                new System.Drawing.Rectangle(cx, cy, cw, ch), System.Drawing.GraphicsUnit.Pixel);
                        }
                        string dir = Path.Combine(Path.GetTempPath(), PptTempDirName);
                        Directory.CreateDirectory(dir);
                        string dest = Path.Combine(dir, "kindcrop-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png");
                        dst.Save(dest, System.Drawing.Imaging.ImageFormat.Png);
                        return dest;
                    }
                }
            }
            catch (Exception) { return null; }
        }

        // ---- page furniture ----

        /// <summary>The header of an ordinary content page: title, accent tick, hairline, and
        /// where the body may start.</summary>
        public double Head(string title, string lang)
        {
            if (title != null && title.Trim().Length > 0)
            {
                Text(48, 28, 864, 52, title, 28, new double[] { 28, 26, 24, 22, 20 }, true, "left", Th.Title, Th.TitleFont, "title");
            }
            // the small accent tick and the hairline under the head band
            Shape(1, 48, 88, 44, 3.5, Th.Accent, null);
            Rule(48, 100, 864, KLineColor(Th), "headRule");
            Foot();
            return 118;
        }

        /// <summary>Footer text at the left, the page number at the right - both plain text boxes
        /// (the WPS slide-number placeholder is not used: it reads "1" on every page here).</summary>
        public void Foot()
        {
            if (Th.Footer != null && Th.Footer.Trim().Length > 0)
                Text(Left, 506, 520, 16, Th.Footer, 10, new double[] { 10 }, false, "left", Th.Muted, Th.Font, "footer");
            if (Th.PageNumber)
                Text(752, 506, 160, 16, PageText(), 10, new double[] { 10 }, false, "right", Th.Muted, Th.Font, "pageNo");
        }

        public string PageText()
        {
            return Index.ToString(CultureInfo.InvariantCulture) + " / " + Total.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Footer that has to live inside a column of the page (the split layout).</summary>
        public void FootIn(double l, double w)
        {
            if (Th.Footer != null && Th.Footer.Trim().Length > 0)
                Text(l, 508, w - 74, 15, Th.Footer, 10, new double[] { 10 }, false, "left", Th.Muted, Th.Font, "footer");
            if (Th.PageNumber)
                Text(l + w - 68, 508, 68, 15, PageText(), 10, new double[] { 10 }, false, "right", Th.Muted, Th.Font, "pageNo");
        }
    }

    /// <summary>Split "a **bold accent** segment" into the visible text plus the character
    /// ranges that carry the mark (1-based, as WPS' Characters() wants them).</summary>
    private static string StripMarks(string text, List<int> starts, List<int> lengths)
    {
        if (text == null) return "";
        if (text.IndexOf("**", StringComparison.Ordinal) < 0) return text;
        StringBuilder sb = new StringBuilder();
        string[] parts = text.Split(new string[] { "**" }, StringSplitOptions.None);
        for (int i = 0; i < parts.Length; i++)
        {
            if (parts[i].Length == 0) continue;
            if (i % 2 == 1)
            {
                starts.Add(sb.Length + 1);
                lengths.Add(parts[i].Length);
            }
            sb.Append(parts[i]);
        }
        return sb.ToString();
    }

    // ---------------------------------------------------------------- item plumbing

    /// <summary>Read the array a layout needs: a top-level key on the request ("cards"), the same
    /// key inside a single-item -Items object ({"cards": [...]}), or the -Items array itself.</summary>
    private static List<object> KindItems(Bag a, string key)
    {
        if (a.Has(key)) return a.List(key);
        List<object> items = a.List("items");
        if (items.Count == 1)
        {
            Bag one = null;
            try { one = new Bag(items[0]); }
            catch (Exception) { }
            if (one != null && one.Has(key)) return one.List(key);
        }
        return items;
    }

    /// <summary>A single object field from the request or from the one-item -Items object.</summary>
    private static Bag KindBag(Bag a, string key)
    {
        if (a.Has(key))
        {
            try { return new Bag(a.Raw(key)); }
            catch (Exception) { }
        }
        List<object> items = a.List("items");
        if (items.Count >= 1)
        {
            Bag one = null;
            try { one = new Bag(items[0]); }
            catch (Exception) { }
            if (one != null && one.Has(key))
            {
                try { return new Bag(one.Raw(key)); }
                catch (Exception) { }
            }
        }
        return null;
    }

    private static string Field(Bag b, string key, string fallback)
    {
        if (b == null) return fallback;
        string v = b.Str(key, null);
        return v == null ? fallback : v;
    }

    /// <summary>The one-item -Items object itself: its fields are the layout's own fields.</summary>
    private static Bag ItemBag(Bag a)
    {
        List<object> items = a.List("items");
        if (items.Count != 1) return null;
        try { return new Bag(items[0]); }
        catch (Exception) { return null; }
    }

    /// <summary>A field that may live on the command line (-Title) or inside the -Items object:
    /// Chinese travels safely through JSON files, not always through the command line.</summary>
    private static string ItemText(Bag a, string key, string fallback)
    {
        string v = a.Str(key, null);
        if (v != null && v.Length > 0) return v;
        Bag one = ItemBag(a);
        if (one != null)
        {
            v = one.Str(key, null);
            if (v != null) return v;
        }
        return fallback;
    }

    private static string PlainOf(object raw)
    {
        if (raw == null) return "";
        if (raw is string) return (string)raw;
        try
        {
            Bag b = new Bag(raw);
            string t = b.Str("text", null);
            if (t != null) return t;
            t = b.Str("head", null);
            if (t != null) return t;
        }
        catch (Exception) { }
        return Convert.ToString(raw, CultureInfo.InvariantCulture);
    }

    // ---------------------------------------------------------------- the eight layouts

    /// <summary>2-6 parallel cards, equal width and equal height, each with an optional big
    /// value, a head and a line of text.</summary>
    private static void KCards(KindPage p, Bag a, List<object> items, string title)
    {
        DeckTheme th = p.Th;
        int n = Math.Max(1, Math.Min(items.Count, 6));
        if (n < 2) p.Warnings.Add("cards expects 2 to 6 cards; only " + items.Count + " were given");
        int perRow = n <= 4 ? n : 3;
        int rows = (n + perRow - 1) / perRow;
        double gap = 18;
        double cw = (p.BodyW - gap * (perRow - 1)) / perRow;
        double ch = Math.Min(rows == 1 ? 218 : 172, (p.BodyH - gap * (rows - 1)) / rows);
        double top = KSettle(p.BodyTop, p.BodyH, rows * ch + gap * (rows - 1));
        int ground = p.CardGround();

        for (int i = 0; i < n; i++)
        {
            Bag card = null;
            try { card = new Bag(items[i]); }
            catch (Exception) { }
            string head = Field(card, "head", null);
            if (head == null) head = PlainOf(items[i]);
            string text = Field(card, "text", null);
            string value = Field(card, "value", null);

            double cx = p.Left + (i % perRow) * (cw + gap);
            double cy = top + (i / perRow) * (ch + gap);
            p.Block(cx, cy, cw, ch, ground, "card" + (i + 1).ToString(CultureInfo.InvariantCulture));
            p.Shape(1, cx + 20, cy + 20, 32, 4, th.Accent, null);

            double ty = cy + 36;
            if (value != null && value.Length > 0)
            {
                p.Text(cx + 20, ty, cw - 40, rows == 1 ? 46 : 40, value, rows == 1 ? 32 : 26,
                       rows == 1 ? new double[] { 32, 28, 24, 20 } : new double[] { 26, 22, 20, 18 },
                       true, "left", th.Accent, th.TitleFont,
                       "cardValue" + (i + 1).ToString(CultureInfo.InvariantCulture));
                ty += rows == 1 ? 52 : 44;
            }
            double headH = rows == 1 ? 46 : 40;
            if (head != null && head.Length > 0)
            {
                p.Text(cx + 20, ty, cw - 40, headH, head, rows == 1 ? 17.5 : 15.5,
                       rows == 1 ? new double[] { 17.5, 16, 15, 14, 13 } : new double[] { 15.5, 14, 13, 12 },
                       true, "left", th.Text, th.Font, "cardHead" + (i + 1).ToString(CultureInfo.InvariantCulture));
                ty += headH + 4;
            }
            double rest = cy + ch - 14 - ty;
            if (text != null && text.Length > 0 && rest > 14)
            {
                p.Text(cx + 20, ty, cw - 40, rest, text, rows == 1 ? 14 : 13,
                       new double[] { rows == 1 ? 14 : 13, 13, 12, 11, 10 },
                       false, "left", th.Muted, th.Font, "cardText" + (i + 1).ToString(CultureInfo.InvariantCulture));
            }
            p.Info.Add(KindInfo("card", i + 1, cx, cy, cw, ch));
        }
    }

    /// <summary>2-4 big figures with a small label: the numbers carry the page.</summary>
    private static void KStats(KindPage p, Bag a, List<object> items, string title)
    {
        DeckTheme th = p.Th;
        int n = Math.Max(1, Math.Min(items.Count, 4));
        if (n < 2) p.Warnings.Add("stats expects 2 to 4 figures; only " + items.Count + " were given");
        double gap = 22;
        double cw = (p.BodyW - gap * (n - 1)) / n;
        int ground = p.CardGround();
        bool valueOnly = true;
        bool hasNote = false;
        for (int i = 0; i < n; i++)
        {
            Bag st = null;
            try { st = new Bag(items[i]); }
            catch (Exception) { }
            if (Field(st, "label", null) != null) valueOnly = false;
            string note0 = Field(st, "text", null);
            if (note0 != null && note0.Length > 0) hasNote = true;
        }
        double ch = Math.Min(p.BodyH - 24, hasNote ? 300 : 234);
        double top = KSettle(p.BodyTop, p.BodyH, ch);
        // one type size for every figure: numbers that shrink one by one would look ragged
        double baseSize = n >= 4 ? 42 : 50;
        double[] sizes = new double[] { baseSize, 38, 34, 30, 26, 22 };
        double shared = sizes[sizes.Length - 1];
        for (int k = 0; k < sizes.Length; k++)
        {
            bool all = true;
            for (int i = 0; i < n; i++)
            {
                Bag st0 = null;
                try { st0 = new Bag(items[i]); }
                catch (Exception) { }
                string v0 = Field(st0, "value", null);
                if (v0 == null) v0 = PlainOf(items[i]);
                string u0 = Field(st0, "unit", null);
                string shown0 = u0 != null && u0.Length > 0 ? v0 + " " + u0 : v0;
                if (KNeedHeight(shown0, p.C.WD(cw - 48), p.C.FS(sizes[k])) > p.C.HT(92) * 1.02) { all = false; break; }
            }
            if (all) { shared = sizes[k]; break; }
        }

        for (int i = 0; i < n; i++)
        {
            Bag st = null;
            try { st = new Bag(items[i]); }
            catch (Exception) { }
            string value = Field(st, "value", null);
            if (value == null) value = PlainOf(items[i]);
            string unit = Field(st, "unit", null);
            string label = Field(st, "label", null);
            string note = Field(st, "text", null);
            double cx = p.Left + i * (cw + gap);

            if (valueOnly)
            {
                // nothing but numbers: no card at all, the figures stand on the page
                p.Shape(1, cx, top + 40, 56, 5, i % 2 == 0 ? th.Primary : th.Accent, null);
                p.Text(cx, top + 56, cw, 110, value, n >= 4 ? 46 : 56, new double[] { n >= 4 ? 46 : 56, 42, 36, 30, 26 },
                       true, "left", i % 2 == 0 ? th.Primary : th.Accent, th.TitleFont,
                       "statValue" + (i + 1).ToString(CultureInfo.InvariantCulture));
                if (label != null)
                    p.Text(cx, top + 176, cw, 40, label, 14, new double[] { 14, 13, 12 }, false, "left", th.Text, th.Font,
                           "statLabel" + (i + 1).ToString(CultureInfo.InvariantCulture));
                if (note != null && note.Length > 0)
                    p.Text(cx, top + 216, cw, ch - 224, note, 11.5, new double[] { 11.5, 10.5, 9.5 }, false, "left", th.Muted, th.Font,
                           "statText" + (i + 1).ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                p.Block(cx, top, cw, ch, ground, "statCard" + (i + 1).ToString(CultureInfo.InvariantCulture));
                p.Shape(1, cx + 24, top + 30, 44, 4, i % 2 == 0 ? th.Primary : th.Accent, null);
                string shown = unit != null && unit.Length > 0 ? value + " " + unit : value;
                dynamic fig = p.Text(cx + 24, top + 48, cw - 48, 92, shown, shared,
                                     new double[] { shared },
                                     true, "left", i % 2 == 0 ? th.Primary : th.Accent, th.TitleFont,
                                     "statValue" + (i + 1).ToString(CultureInfo.InvariantCulture));
                if (unit != null && unit.Length > 0 && fig != null)
                {
                    // the unit is set small next to the figure, in the body colour
                    try
                    {
                        dynamic tail = fig.TextFrame.TextRange.Characters(value.Length + 2, unit.Length);
                        tail.Font.Size = (float)p.C.FS(17);
                        tail.Font.Bold = 0;
                        tail.Font.Color.RGB = th.Text;
                    }
                    catch (Exception) { }
                }
                p.Rule(cx + 24, top + 152, 24, KLineColor(th), null);
                if (label != null)
                    p.Text(cx + 24, top + 166, cw - 48, 52, label, 14, new double[] { 14, 13, 12 }, false, "left", th.Text, th.Font,
                           "statLabel" + (i + 1).ToString(CultureInfo.InvariantCulture));
                if (note != null && note.Length > 0)
                    p.Text(cx + 24, top + 224, cw - 48, ch - 238, note, 11.5, new double[] { 11.5, 10.5, 9.5 },
                           false, "left", th.Muted, th.Font,
                           "statText" + (i + 1).ToString(CultureInfo.InvariantCulture));
            }
            p.Info.Add(KindInfo("stat", i + 1, cx, top, cw, ch));
        }
    }

    /// <summary>A horizontal time line: one rule, one dot per step, the words alternating above
    /// and below the line so two neighbouring steps never collide.</summary>
    private static void KTimeline(KindPage p, Bag a, List<object> items, string title)
    {
        DeckTheme th = p.Th;
        int n = Math.Max(1, Math.Min(items.Count, 5));
        if (items.Count < 3) p.Warnings.Add("timeline reads best with 3 to 5 steps; " + items.Count + " were given");
        double cw = p.BodyW / n;
        double lineY = p.BodyTop + (p.BodyH * 0.5);
        p.Rule(p.Left, lineY, p.BodyW, KLineColor(th), "tlRule");

        double aboveRoom = lineY - 26 - p.BodyTop;
        double belowRoom = p.BodyBottom - 8 - (lineY + 26);

        for (int i = 0; i < n; i++)
        {
            Bag st = null;
            try { st = new Bag(items[i]); }
            catch (Exception) { }
            string when = Field(st, "when", null);
            if (when == null) when = (i + 1).ToString("00", CultureInfo.InvariantCulture);
            string head = Field(st, "head", null);
            if (head == null) head = PlainOf(items[i]);
            string text = Field(st, "text", null);
            double ccx = p.Left + i * cw + cw / 2;
            double bw = Math.Max(60, cw - 18);
            double bx = ccx - bw / 2;
            bool up = (i % 2) == 0;
            double room = up ? aboveRoom : belowRoom;

            // choose the largest type that still fits the one side of the line
            double whenSize = 14, headSize = 16, textSize = 14;
            double[] whenL = new double[] { 14, 13, 12 };
            double[] headL = new double[] { 16, 15, 14, 13 };
            double[] textL = new double[] { 14, 13, 12, 11 };
            double wh = 0, mh = 0, thh = 0, total = 0;
            bool ok = false;
            for (int k = 0; k < 4 && !ok; k++)
            {
                whenSize = whenL[Math.Min(k, whenL.Length - 1)];
                headSize = headL[Math.Min(k, headL.Length - 1)];
                textSize = textL[Math.Min(k, textL.Length - 1)];
                wh = KNeedHeight(when, p.C.WD(bw), p.C.FS(whenSize));
                mh = KNeedHeight(head, p.C.WD(bw), p.C.FS(headSize));
                thh = text == null ? 0 : KNeedHeight(text, p.C.WD(bw), p.C.FS(textSize));
                total = wh + 4 + mh + (thh > 0 ? 4 + thh : 0);
                ok = total <= room;
            }
            if (!ok)
                p.Warnings.Add("timeline step " + (i + 1) + " may not fit: " + Math.Round(total, 0)
                    + "pt of words in " + Math.Round(room, 0) + "pt on " + (up ? "the upper" : "the lower") + " side of the line");

            double y = up ? lineY - 26 - total : lineY + 26;
            p.Text(bx, y, bw, wh + 3, when, whenSize, whenL, true, "center",
                   i == n - 1 ? th.Accent : th.Primary, th.TitleFont,
                   "tlWhen" + (i + 1).ToString(CultureInfo.InvariantCulture));
            y += wh + 4;
            p.Text(bx, y, bw, mh + 3, head, headSize, headL, true, "center", th.Text, th.Font,
                   "tlHead" + (i + 1).ToString(CultureInfo.InvariantCulture));
            y += mh + 4;
            if (text != null && text.Length > 0)
                p.Text(bx, y, bw, thh + 3, text, textSize, textL, false, "center", th.Muted, th.Font,
                       "tlText" + (i + 1).ToString(CultureInfo.InvariantCulture));
            p.Oval(ccx - 7, lineY - 7, 14, 14, i == n - 1 ? th.Accent : th.Primary,
                   "tlDot" + (i + 1).ToString(CultureInfo.InvariantCulture));
            p.Info.Add(KindInfo("step", i + 1, bx, up ? lineY - 26 - total : lineY + 26, bw, total));
        }
    }

    /// <summary>3-5 steps in a row, numbered, with an arrow between each pair.</summary>
    private static void KProcess(KindPage p, Bag a, List<object> items, string title)
    {
        DeckTheme th = p.Th;
        int n = Math.Max(1, Math.Min(items.Count, 5));
        if (items.Count < 3) p.Warnings.Add("process reads best with 3 to 5 steps; " + items.Count + " were given");
        double gap = 26;
        double cw = (p.BodyW - gap * (n - 1)) / n;

        // the card grows with the longest text so the blocks stay equal height
        double bodySize = 14;
        double[] bodyLadder = new double[] { 14, 13, 12, 11 };
        double need = 0;
        for (int i = 0; i < n; i++)
        {
            Bag st = null;
            try { st = new Bag(items[i]); }
            catch (Exception) { }
            string text = Field(st, "text", null);
            double h = KNeedHeight(text == null ? "" : text, p.C.WD(cw - 32), p.C.FS(bodySize));
            if (h > need) need = h;
            string head = Field(st, "head", null);
            double hh = KNeedHeight(head == null ? "" : head, p.C.WD(cw - 32), p.C.FS(16));
            if (hh > 46) need += hh - 46;
        }
        double ch = Math.Min(p.BodyH - 12, Math.Max(192, 148 + need));
        double top = KSettle(p.BodyTop, p.BodyH, ch);

        for (int i = 0; i < n; i++)
        {
            Bag st = null;
            try { st = new Bag(items[i]); }
            catch (Exception) { }
            string head = Field(st, "head", null);
            if (head == null) head = PlainOf(items[i]);
            string text = Field(st, "text", null);
            double cx = p.Left + i * (cw + gap);
            int tone = p.ToneColor(i);
            p.Frame(cx, top, cw, ch, KLineColor(th), 1, "step" + (i + 1).ToString(CultureInfo.InvariantCulture));
            p.Oval(cx + 18, top + 18, 34, 34, tone, "stepNo" + (i + 1).ToString(CultureInfo.InvariantCulture));
            p.Text(cx + 18, top + 18, 34, 34, (i + 1).ToString(CultureInfo.InvariantCulture), 15,
                   new double[] { 15, 13 }, true, "center", p.SafeInk(tone), th.TitleFont,
                   "stepNum" + (i + 1).ToString(CultureInfo.InvariantCulture));
            p.Text(cx + 16, top + 62, cw - 32, 46, head, 16, new double[] { 16, 15, 14, 13 }, true, "left", th.Text, th.Font,
                   "stepHead" + (i + 1).ToString(CultureInfo.InvariantCulture));
            if (text != null && text.Length > 0)
                p.Text(cx + 16, top + 112, cw - 32, ch - 128, text, bodySize, bodyLadder, false, "left", th.Muted, th.Font,
                       "stepText" + (i + 1).ToString(CultureInfo.InvariantCulture));
            if (i < n - 1)
            {
                double ay = top + ch / 2 - 7;
                p.Shape(33, cx + cw + 5, ay, gap - 10, 14, KLineColor(th), null);
            }
            p.Info.Add(KindInfo("step", i + 1, cx, top, cw, ch));
        }
    }

    /// <summary>Two sides side by side, one card each, the points of both numbered alike.</summary>
    private static void KCompare(KindPage p, Bag a, string title)
    {
        DeckTheme th = p.Th;
        Bag left = KindBag(a, "left");
        Bag right = KindBag(a, "right");
        if (left == null && right == null)
            throw new Fail("BAD_ARGS", "compare needs \"left\" and \"right\": {\"left\":{\"head\":..,\"points\":[..]},\"right\":{..}};"
                + " pass them in the -Items JSON object");
        double gap = 24;
        double cw = (p.BodyW - gap) / 2;
        List<string> lPoints = KPointTexts(left);
        List<string> rPoints = KPointTexts(right);
        int most = Math.Max(1, Math.Max(lPoints.Count, rPoints.Count));
        double headH = 36;
        double rowRoom = p.BodyH - headH - 40;
        double row = Math.Min(64, rowRoom / most);
        double ch = Math.Min(p.BodyH, 26 + headH + 26 + most * row + 10);
        double top = KSettle(p.BodyTop, p.BodyH, ch);
        int ground = p.CardGround();

        for (int s = 0; s < 2; s++)
        {
            Bag side = s == 0 ? left : right;
            if (side == null) continue;
            List<string> points = s == 0 ? lPoints : rPoints;
            int tone = s == 0 ? th.Primary : th.Accent;
            double cx = p.Left + s * (cw + gap);
            p.Block(cx, top, cw, ch, ground, "side" + (s + 1).ToString(CultureInfo.InvariantCulture));
            p.Shape(1, cx + 24, top + 26, 40, 4, tone, null);
            p.Text(cx + 24, top + 40, cw - 48, headH, Field(side, "head", ""), 17, new double[] { 17, 16, 15, 14 },
                   true, "left", tone, th.Font, "sideHead" + (s + 1).ToString(CultureInfo.InvariantCulture));
            double py = top + 26 + headH + 22;
            for (int i = 0; i < points.Count && i < most; i++)
            {
                if (py + row > top + ch - 8) break;
                p.Oval(cx + 24, py + 8, 6, 6, tone, null);
                p.Text(cx + 40, py, cw - 64, row - 4, points[i], 14, new double[] { 14, 13, 12, 11 },
                       false, "left", th.Text, th.Font, "sidePoint" + (s + 1) + "-" + (i + 1).ToString(CultureInfo.InvariantCulture));
                py += row;
            }
            p.Info.Add(KindInfo("side", s + 1, cx, top, cw, ch));
        }
        p.Info.Add(KindInfo("divider", 0, p.Left + cw, p.BodyTop + 20, gap, p.BodyH - 40));
    }

    private static List<string> KPointTexts(Bag side)
    {
        List<string> list = new List<string>();
        if (side == null || !side.Has("points")) return list;
        List<object> raw = side.List("points");
        for (int i = 0; i < raw.Count; i++)
        {
            string t = PlainOf(raw[i]);
            if (t != null && t.Length > 0) list.Add(t);
        }
        return list;
    }

    /// <summary>2-4 numbered takeaways: a big number, a head and one sentence each.</summary>
    private static void KSummary(KindPage p, Bag a, List<object> items, string title)
    {
        DeckTheme th = p.Th;
        int n = Math.Max(1, Math.Min(items.Count, 4));
        if (items.Count < 2) p.Warnings.Add("summary expects 2 to 4 points; " + items.Count + " were given");
        double rowH = Math.Min(122, p.BodyH / n);
        double top = KSettle(p.BodyTop, p.BodyH, n * rowH);
        for (int i = 0; i < n; i++)
        {
            Bag pt = null;
            try { pt = new Bag(items[i]); }
            catch (Exception) { }
            string head = Field(pt, "head", null);
            if (head == null) head = PlainOf(items[i]);
            string text = Field(pt, "text", null);
            double y = top + i * rowH;
            p.Text(p.Left, y, 92, 56, (i + 1).ToString("00", CultureInfo.InvariantCulture), 40,
                   new double[] { 40, 36, 30, 26 }, true, "left", i == 0 ? th.Accent : th.Primary, th.TitleFont,
                   "sumNo" + (i + 1).ToString(CultureInfo.InvariantCulture));
            double ty = y + 2;
            if (head != null && head.Length > 0)
            {
                p.Text(p.Left + 100, ty, p.BodyW - 100, 32, head, 18, new double[] { 18, 17, 16, 15 },
                       true, "left", th.Text, th.Font, "sumHead" + (i + 1).ToString(CultureInfo.InvariantCulture));
                ty += 34;
            }
            if (text != null && text.Length > 0)
                p.Text(p.Left + 100, ty, p.BodyW - 100, y + rowH - 12 - ty, text, 14,
                       new double[] { 14, 13, 12, 11 }, false, "left", th.Muted, th.Font,
                       "sumText" + (i + 1).ToString(CultureInfo.InvariantCulture));
            if (i < n - 1) p.Rule(p.Left, y + rowH - 4, p.BodyW, KLineColor(th), null);
            p.Info.Add(KindInfo("point", i + 1, p.Left, y, p.BodyW, rowH));
        }
    }

    /// <summary>2-6 pictures in a grid with captions, each scaled to fit its cell whole
    /// (or cut to fill it when fit is "cover").</summary>
    private static void KGallery(KindPage p, Bag a, List<object> items, string title)
    {
        DeckTheme th = p.Th;
        int n = Math.Max(1, Math.Min(items.Count, 6));
        if (items.Count == 0)
            throw new Fail("BAD_ARGS", "gallery needs \"images\": [{\"path\":\"a.png\",\"caption\":\"..\"}, ..]");
        int perRow = n <= 3 ? n : (n == 4 ? 2 : 3);
        int rows = (n + perRow - 1) / perRow;
        double gap = 18;
        double cw = (p.BodyW - gap * (perRow - 1)) / perRow;
        double ch = (p.BodyH - gap * (rows - 1)) / rows;
        string fit = (a.Str("fit", "contain") ?? "contain").Trim().ToLowerInvariant();
        if (fit != "contain" && fit != "cover") throw new Fail("BAD_ARGS", "-Fit for gallery must be contain or cover");

        // captions too long for one line get two, under every picture alike
        bool twoLines = false;
        for (int i = 0; i < n; i++)
        {
            Bag im = null;
            try { im = new Bag(items[i]); }
            catch (Exception) { }
            string cap = Field(im, "caption", null);
            if (cap != null && cap.Length * 10.5 > cw + 10) twoLines = true;
        }
        double capH = twoLines ? 34 : 20;

        for (int i = 0; i < n; i++)
        {
            Bag im = null;
            try { im = new Bag(items[i]); }
            catch (Exception) { }
            string path = Field(im, "path", null);
            if (path == null) path = Field(im, "image", null);
            if (path == null) path = Field(im, "text", null);
            if (path == null) path = Convert.ToString(items[i], CultureInfo.InvariantCulture);
            string cap = Field(im, "caption", null);
            string itemFit = Field(im, "fit", fit);
            double aspect = Field(im, "aspect", null) == null ? 0 : im.Dbl("aspect", 0);

            int row = i / perRow;
            int inRow = Math.Min(perRow, n - row * perRow);
            double rowW = inRow * cw + (inRow - 1) * gap;
            double cx = p.Left + (p.BodyW - rowW) / 2 + (i % perRow) * (cw + gap);
            double cy = p.BodyTop + row * (ch + gap);
            double cellH = ch - capH - 4;
            if ("cover".Equals(itemFit, StringComparison.OrdinalIgnoreCase))
            {
                p.Picture(cx, cy, cw, cellH, path, "cover", "pic" + (i + 1).ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                // shown whole: the picture sits on the BOTTOM of its cell so every caption in a
                // row starts at the same height, right under its picture
                double ar = KImageAspect(path);
                if (ar <= 0) ar = cw / Math.Max(cellH, 1);
                double iw = ar > cw / Math.Max(cellH, 1) ? cw : cellH * ar;
                double ih = ar > cw / Math.Max(cellH, 1) ? cw / ar : cellH;
                p.Picture(cx + (cw - iw) / 2, cy + cellH - ih, iw, ih, path, "contain",
                          "pic" + (i + 1).ToString(CultureInfo.InvariantCulture));
            }
            if (cap != null && cap.Length > 0)
                p.Text(cx, cy + ch - capH + 2, cw, capH - 2, cap, 10, new double[] { 10 },
                       false, "center", th.Muted, th.Font, "cap" + (i + 1).ToString(CultureInfo.InvariantCulture));
            p.Info.Add(KindInfo("image", i + 1, cx, cy, cw, ch));
            if (aspect < 0) p.Warnings.Add("gallery item " + (i + 1) + ": negative aspect ignored");
        }
        p.Info.Add(KindInfo("mode", 0, 0, 0, 0, 0, fit + (fit == "cover" ? " (cropped to a uniform cell)" : " (shown whole, never cut)")));
    }

    /// <summary>Half the page a picture edge to edge, the words on the other half.</summary>
    private static void KSplit(KindPage p, Bag a, Bag one, string title)
    {
        DeckTheme th = p.Th;
        double W = p.C.W / p.C.Sx;   // design width (960 by construction)
        bool imgRight = !"left".Equals(ItemText(a, "side", Field(one, "side", "right")).Trim().ToLowerInvariant(), StringComparison.Ordinal);
        string path = ItemText(a, "path", Field(one, "image", null));
        if (path == null) path = ItemText(a, "image", null);
        if (path == null || path.Length == 0) throw new Fail("BAD_ARGS", "split needs -Path <image> or \"image\" in the -Items object");
        string text = ItemText(a, "text", Field(one, "text", null));
        List<object> points = new List<object>();
        if (one != null && one.Has("points")) points = one.List("points");
        else points = KindItems(a, "points");

        double imgX = imgRight ? 480 : 0;
        double tx = imgRight ? 56 : 536;
        double tw = 376;
        p.Picture(imgX, 0, 480, 540, path, "cover", "splitImage");

        p.Shape(1, tx, 104, 56, 4, th.Accent, null);
        p.Text(tx, 118, tw, 112, title, 30, new double[] { 30, 26, 24, 22, 20, 18 }, true, "left", th.Title,
               th.TitleFont, "title");
        double y = 240;
        if (text != null && text.Length > 0)
        {
            double thh = KNeedHeight(text, p.C.WD(tw), p.C.FS(13.5));
            thh = Math.Min(thh, 120);
            p.Text(tx, y, tw, thh + 2, text, 14, new double[] { 14, 13, 12, 11 }, false, "left", th.Text,
                   th.Font, "body");
            y += thh + 18;
        }
        int count = Math.Min(points.Count, 4);
        double room = 492 - y;
        double row = count == 0 ? 0 : Math.Min(56, room / count);
        for (int i = 0; i < count; i++)
        {
            p.Shape(1, tx, y + 8, 6, 6, p.ToneColor(i), null);
            p.Text(tx + 18, y, tw - 18, Math.Max(16, row - 6), PlainOf(points[i]), 14,
                   new double[] { 14, 13, 12 }, false, "left", th.Text, th.Font,
                   "point" + (i + 1).ToString(CultureInfo.InvariantCulture));
            y += row;
        }
        p.FootIn(tx, tw);
        p.Info.Add(KindInfo("image", 0, imgX, 0, 480, 540));
        p.Info.Add(KindInfo("text", 0, tx, 100, tw, 400));
        if (Math.Abs(W - 960) > 1) p.Info.Add(KindInfo("note", 0, 0, 0, 0, 0, "design width " + Math.Round(W, 0)));
    }

    private static Dictionary<string, object> KindInfo(string kind, int i, double l, double t, double w, double h)
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["kind"] = kind;
        if (i > 0) d["i"] = i;
        d["l"] = Math.Round(l, 1);
        d["t"] = Math.Round(t, 1);
        d["w"] = Math.Round(w, 1);
        d["h"] = Math.Round(h, 1);
        return d;
    }

    private static Dictionary<string, object> KindInfo(string kind, int i, double l, double t, double w, double h, string note)
    {
        Dictionary<string, object> d = KindInfo(kind, i, l, t, w, h);
        d["note"] = note;
        return d;
    }

    // ---------------------------------------------------------------- entry point

    /// <summary>Build one finished layout and report what was drawn, what had to shrink and what
    /// may still be too wide for its box.  Returns the object `slide -Preset` hands back.</summary>
    private static object BuildFinishedSlide(dynamic pres, dynamic slide, Bag a, string preset, DeckTheme th, out List<string> problems)
    {
        problems = new List<string>();
        double W, H;
        SlideSize(pres, out W, out H);
        KindPage p = new KindPage();
        p.Slide = slide;
        p.Th = th;
        p.C = new KindCanvas(W, H);
        p.Warnings = problems;
        try { p.Index = Convert.ToInt32(TryGet(() => (object)slide.SlideIndex, (object)1), CultureInfo.InvariantCulture); }
        catch (Exception) { p.Index = 1; }
        try { p.Total = Convert.ToInt32(TryGet(() => (object)pres.Slides.Count, (object)p.Index), CultureInfo.InvariantCulture); }
        catch (Exception) { p.Total = p.Index; }
        if (p.Total < p.Index) p.Total = p.Index;

        string title = ItemText(a, "title", "");
        bool isSplit = preset == "split";
        if (!isSplit)
        {
            if (title.Length == 0) throw new Fail("BAD_ARGS", "preset \"" + preset + "\" needs -Title (or \"title\" in the -Items object)");
            p.Head(title, null);
        }
        else if (title.Length == 0) throw new Fail("BAD_ARGS", "preset \"split\" needs -Title (or \"title\" in the -Items object)");

        if (preset == "cards") KCards(p, a, KindItems(a, "cards"), title);
        else if (preset == "stats") KStats(p, a, KindItems(a, "stats"), title);
        else if (preset == "timeline") KTimeline(p, a, KindItems(a, "steps"), title);
        else if (preset == "process") KProcess(p, a, KindItems(a, "steps"), title);
        else if (preset == "compare") KCompare(p, a, title);
        else if (preset == "summary") KSummary(p, a, KindItems(a, "points"), title);
        else if (preset == "gallery") KGallery(p, a, KindItems(a, "images"), title);
        else if (preset == "split") KSplit(p, a, ItemBag(a), title);
        else throw new Fail("BAD_ARGS", "unknown preset: " + preset);

        // what actually landed on the slide (read back from WPS, not what we meant to draw)
        Dictionary<string, object> r = PresetReport(pres, slide, preset, problems);
        r["fontSizeDrops"] = p.Drops;
        r["layout"] = p.Info;
        r["designCanvas"] = "960x540 scaled to " + Math.Round(W, 0) + "x" + Math.Round(H, 0);
        if (preset == "gallery")
            r["fit"] = (a.Str("fit", "contain") ?? "contain").Trim().ToLowerInvariant()
                + ": contain shows every picture whole and centred (nothing is cut); cover crops each picture to its cell";
        r["pageNumber"] = th.PageNumber ? p.PageText() : "off (theme pageNumber=false)";
        return r;
    }

    /// <summary>The object `slide -Preset` returns: the shapes that are really on the page, read
    /// back from WPS, plus a note for every shape that sticks out of the slide.</summary>
    private static Dictionary<string, object> PresetReport(dynamic pres, dynamic slide, string preset, List<string> problems)
    {
        double W, H;
        SlideSize(pres, out W, out H);
        if (problems == null) problems = new List<string>();
        List<object> shapes = new List<object>();
        int count = 0;
        try { count = (int)slide.Shapes.Count; } catch (Exception) { }
        for (int k = 1; k <= count; k++)
        {
            dynamic sh = null;
            try { sh = slide.Shapes.Item(k); } catch (Exception) { continue; }
            string nm = ShapeNameOf(sh);
            if (IsTakeoverMark(nm)) continue;
            Dictionary<string, object> d = new Dictionary<string, object>();
            d["name"] = nm;
            d["type"] = Convert.ToString(TryGet(() => (object)(int)sh.Type, (object)0), CultureInfo.InvariantCulture);
            d["l"] = R(ShapeD(sh, "left", 0));
            d["t"] = R(ShapeD(sh, "top", 0));
            d["w"] = R(ShapeD(sh, "width", 0));
            d["h"] = R(ShapeD(sh, "height", 0));
            string txt = ShapeText(sh);
            if (txt.Length > 0) d["text"] = KClip(txt, 40);
            shapes.Add(d);
            // anything sticking out of the page is reported here (a layout should never do it)
            double l = ToD(d["l"]), t = ToD(d["t"]), w = ToD(d["w"]), h = ToD(d["h"]);
            if (l < -2 || t < -2 || l + w > W + 2 || t + h > H + 2)
                problems.Add("shape \"" + nm + "\" sits outside the page (" + l + "," + t + " " + w + "x" + h + ")");
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["preset"] = preset;
        try { r["slide"] = Convert.ToInt32(TryGet(() => (object)slide.SlideIndex, (object)1), CultureInfo.InvariantCulture); }
        catch (Exception) { r["slide"] = 1; }
        r["shapes"] = shapes;
        r["overflowWarnings"] = problems;
        return r;
    }
}
