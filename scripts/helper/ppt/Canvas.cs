// wps-office-cooperate (ppt): finished-slide layouts, phase 2 (v1.32).
//
// Adds three things on top of the eight ready-made layouts in Kinds.cs:
//   * `slide -Preset canvas`  - free layout: every item carries its own [x,y,w,h] box on the
//     960x540 design canvas, and `edges` join node items with real connectors;
//   * `slide -Preset grid`    - rows of parts with a span; widths come from the spans, row
//     heights from how much text each part needs, and the leftover height goes to the
//     parts that can use it (charts, pictures, panels);
//   * shape-drawn charts (`chart` item): scatter / curve / pie / doughnut painted with
//     polygons, lines and text - not a native chart object (WPS needs an open window for
//     AddChart, refuses more than three series, and a native chart cannot be put into a
//     free-layout box of our choosing).
// It also carries `slide -Action reuse` (duplicate a page, then swap text/pictures by
// shape name).
//
// Geometry contract: everything here is written in design units on a 960 x 540 canvas and
// mapped through KindCanvas, so a 4:3 deck gets the same composition.  Anything that falls
// outside the body area is clamped and reported in overflowWarnings - never silently moved.
//
// Pure ASCII source, C# 5 (no string interpolation, no null-conditional operator).

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

internal static partial class Program
{
    // ---------------------------------------------------------------- reading numbers

    private static double KDbl(object v, double fallback)
    {
        if (v == null) return fallback;
        if (v is long) return (long)v;
        if (v is double) return (double)v;
        if (v is bool) return ((bool)v) ? 1 : 0;
        double d;
        if (double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), NumberStyles.Float,
                CultureInfo.InvariantCulture, out d)) return d;
        return fallback;
    }

    /// <summary>Read a [x,y,w,h] (or [x1,y1,dx,dy] for a line) array; null when it is not one.</summary>
    private static double[] KBoxArr(object raw)
    {
        List<object> arr = raw as List<object>;
        if (arr == null || arr.Count < 4) return null;
        return new double[] { KDbl(arr[0], 0), KDbl(arr[1], 0), KDbl(arr[2], 0), KDbl(arr[3], 0) };
    }

    private static double[] KBoxArr(Bag b, string key)
    {
        if (b == null || !b.Has(key)) return null;
        return KBoxArr(b.Raw(key));
    }

    /// <summary>A descending list of type sizes from <paramref name="top"/> down to the floor,
    /// one point at a time - the ladder Kinds.cs steps a text box down with.</summary>
    private static double[] KLadder(double top, double floor)
    {
        List<double> l = new List<double>();
        double s = top;
        int guard = 0;
        while (s > floor - 0.01 && guard < 40)
        {
            l.Add(Math.Round(s, 1));
            s -= 1.0;
            guard++;
        }
        if (l.Count == 0) l.Add(floor);
        return l.ToArray();
    }

    private static string KNum(double v, int digits)
    {
        return Math.Round(v, digits).ToString(CultureInfo.InvariantCulture);
    }

    private static List<object> KArr(object raw)
    {
        return raw as List<object> ?? new List<object>();
    }

    /// <summary>Rough width of a short label in points (for centring an edge label).</summary>
    private static double KTextWidth(string text, double size)
    {
        double em = 0;
        if (text != null) foreach (char c in text) em += KGlyphEm(c);
        return em * size;
    }

    /// <summary>The single document object of a canvas / grid / reuse request.  The wrapper
    /// hands -Items over as an array; a lone object that carries one of the document keys
    /// (items / canvas / rows / edges / texts / images) is the document itself.</summary>
    private static Bag KDoc(Bag a)
    {
        if (a.Has("canvas"))
        {
            try { return new Bag(a.Raw("canvas")); } catch (Exception) { }
        }
        List<object> items = a.List("items");
        if (items.Count != 1) return null;
        Bag one = null;
        try { one = new Bag(items[0]); } catch (Exception) { return null; }
        string[] keys = new string[] { "items", "canvas", "rows", "edges", "texts", "images" };
        for (int i = 0; i < keys.Length; i++) if (one.Has(keys[i])) return one;
        return null;
    }

    // ---------------------------------------------------------------- primitives

    /// <summary>One straight segment.  Used for arrow items, chart axes, curve segments and
    /// fit lines; a plain line shape covers all of them (no freeform involved).</summary>
    private static dynamic KLine(KindPage p, double x1, double y1, double x2, double y2,
                                 int color, double weight, bool arrow, bool dashed, string name)
    {
        dynamic ln = null;
        try
        {
            ln = p.Slide.Shapes.AddLine((float)p.C.X(x1), (float)p.C.Y(y1),
                                       (float)p.C.X(x2), (float)p.C.Y(y2));
        }
        catch (Exception e) { throw new Fail("COM_ERROR", "AddLine failed: " + Clean(e.Message)); }
        try
        {
            ln.Line.Visible = -1;
            ln.Line.ForeColor.RGB = color;
            ln.Line.Weight = (float)Math.Max(0.4, weight * p.C.S);
        }
        catch (Exception) { }
        if (arrow) { try { ln.Line.EndArrowheadStyle = 2; } catch (Exception) { } }
        if (dashed) { try { ln.Line.DashStyle = 4; } catch (Exception) { } }
        if (name != null) { try { ln.Name = name; } catch (Exception) { } }
        return ln;
    }

    /// <summary>A filled polygon built with BuildFreeform.  Returns null when this WPS build
    /// refuses the freeform builder - the caller then falls back to a shape with adjustments.</summary>
    private static dynamic KFreeform(KindPage p, List<double[]> pts, int fill, string name)
    {
        if (pts == null || pts.Count < 3) return null;
        dynamic shape = null;
        try
        {
            dynamic fb = p.Slide.Shapes.BuildFreeform(0, (float)p.C.X(pts[0][0]), (float)p.C.Y(pts[0][1]));
            for (int i = 1; i < pts.Count; i++)
                fb.AddNodes(0, 0, (float)p.C.X(pts[i][0]), (float)p.C.Y(pts[i][1]),
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // return to the first node so the outline closes and the fill covers the ring
            fb.AddNodes(0, 0, (float)p.C.X(pts[0][0]), (float)p.C.Y(pts[0][1]),
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            shape = fb.ConvertToShape();
        }
        catch (Exception) { return null; }
        if (shape == null) return null;
        try { shape.Fill.Solid(); shape.Fill.ForeColor.RGB = fill; } catch (Exception) { }
        try { shape.Line.Visible = 0; } catch (Exception) { }
        if (name != null) { try { shape.Name = name; } catch (Exception) { } }
        return shape;
    }

    private static int KSeriesColor(DeckTheme th, int i)
    {
        if (i == 0) return th.Primary;
        if (i == 1) return th.Accent;
        if (i == 2) return th.Accent2;
        int[] extra = new int[] {
            KMix(th.Primary, th.White, 0.38), KMix(th.Accent, th.White, 0.38),
            KMix(th.Accent2, th.White, 0.30), KMix(th.Text, th.White, 0.45)
        };
        int k = (i - 3) % extra.Length;
        if (k < 0) k += extra.Length;
        return extra[k];
    }

    // ---------------------------------------------------------------- nodes and items

    /// <summary>A named box the edges can attach to.</summary>
    private sealed class KNode
    {
        public string Id;
        public dynamic Shape;
        public double L, T, W, H;
    }

    /// <summary>Read one box off an item, clamp it into the body area, and report when the
    /// clamp actually moved something (an item drawn in the margins would look like a bug).</summary>
    private static double[] KFitBox(KindPage p, Bag it, string kind, int index, double[] fallback)
    {
        double[] box = KBoxArr(it, "box");
        if (box == null) box = KBoxArr(it, "rect");
        if (box == null) box = fallback;
        if (box == null) return null;
        double l = box[0], t = box[1], w = box[2], h = box[3];
        double nl = l, nt = t, nw = w, nh = h;
        if (nw < 8) nw = 8;
        if (nh < 8) nh = 8;
        if (nl < p.Left) nl = p.Left;
        if (nl > p.Right - 8) nl = p.Right - 8;
        if (nt < p.BodyTop) nt = p.BodyTop;
        if (nt > p.BodyBottom - 8) nt = p.BodyBottom - 8;
        if (nl + nw > p.Right) nw = p.Right - nl;
        if (nt + nh > p.BodyBottom) nh = p.BodyBottom - nt;
        if (Math.Abs(nl - l) > 0.6 || Math.Abs(nt - t) > 0.6 || Math.Abs(nw - w) > 0.6 || Math.Abs(nh - h) > 0.6)
            p.Warnings.Add("item " + (index + 1).ToString(CultureInfo.InvariantCulture) + " (" + kind + ") box ["
                + KNum(l, 0) + "," + KNum(t, 0) + "," + KNum(w, 0) + "," + KNum(h, 0)
                + "] was clamped into the body area -> [" + KNum(nl, 0) + "," + KNum(nt, 0) + ","
                + KNum(nw, 0) + "," + KNum(nh, 0) + "] (body x " + KNum(p.Left, 0) + ".." + KNum(p.Right, 0)
                + ", y " + KNum(p.BodyTop, 0) + ".." + KNum(p.BodyBottom, 0) + ")");
        return new double[] { nl, nt, nw, nh };
    }

    /// <summary>Type sizes and paddings of a card.  Shared by the drawing code and by the
    /// grid's height estimate, so that what is measured is exactly what is drawn.</summary>
    private static void KCardMetrics(out double vs, out double hs, out double padTop, out double padBottom)
    {
        vs = 24; hs = 14; padTop = 26; padBottom = 14;
    }
    private static string KItemKind(Bag it)
    {
        string kind = (it.Str("type", "text") ?? "text").Trim().ToLowerInvariant();
        if (kind == "title") kind = "heading";
        if (kind == "bullet") kind = "bullets";
        if (kind == "picture" || kind == "img") kind = "image";
        if (kind == "rect" || kind == "roundrect" || kind == "circle") kind = "panel";
        return kind;
    }

    private static int SafeTextOn(DeckTheme th, int fill)
    {
        return KIsDark(fill) ? th.White : th.Text;
    }

    /// <summary>Draw one item inside the given box.  Shared by the canvas and the grid: the
    /// canvas reads the box from JSON, the grid computes it from spans and row heights.</summary>
    private static void KDrawItem(KindPage p, Bag it, double l, double t, double w, double h,
                                  Dictionary<string, KNode> nodes, string tag, int index)
    {
        string kind = KItemKind(it);
        DeckTheme th = p.Th;
        string name = tag;
        if (kind == "node")
        {
            string id = it.Str("id", tag);
            if (id == null || id.Length == 0) id = tag;
            name = "node-" + id;
        }

        if (kind == "text")
        {
            double size = KDbl(it.Raw("size"), 14);
            if (size < 6) size = 6;
            if (size > 72) size = 72;
            int color = ColorOf(it.Str("color", null), th.Text);
            string align = (it.Str("align", "left") ?? "left").ToLowerInvariant();
            string valign = (it.Str("valign", "top") ?? "top").ToLowerInvariant();
            dynamic tb = p.Text(l, t, w, h, Field(it, "text", null), size, KLadder(size, 9),
                                it.Flag("bold", false), align, color, Field(it, "font", th.Font), name);
            p.Anchor(tb, valign == "middle" || valign == "center" ? 2 : valign == "bottom" ? 3 : 1);
        }
        else if (kind == "heading")
        {
            double size = KDbl(it.Raw("size"), 19);
            double barH = Math.Min(h - 6 > 8 ? h - 6 : 8, 22);
            p.Shape(1, l, t + 3, 4, barH, th.Accent, null);
            p.Text(l + 12, t, w - 12, h, Field(it, "text", null), size, KLadder(size, 11),
                   true, "left", ColorOf(it.Str("color", null), th.Title), th.TitleFont, name);
        }
        else if (kind == "card")
        {
            int ground = p.CardGround();
            p.Block(l, t, w, h, ColorOf(it.Str("fill", null), ground), name);
            p.Shape(1, l + 18, t + 15, 28, 3.5, th.Accent, null);
            double vs, hs, padTop, padBottom;
            KCardMetrics(out vs, out hs, out padTop, out padBottom);
            double ty = t + padTop;
            string value = Field(it, "value", null);
            if (value != null && value.Length > 0)
            {
                p.Text(l + 18, ty, w - 36, vs * 1.4, value, vs, KLadder(vs, 12), true, "left", th.Accent, th.TitleFont, name + "-v");
                ty += vs * 1.45;
            }
            string head = Field(it, "head", null);
            if (head != null && head.Length > 0)
            {
                p.Text(l + 18, ty, w - 36, hs * 1.45, head, hs, KLadder(hs, 10), true, "left", th.Text, th.Font, name + "-h");
                ty += hs * 1.5;
            }
            string text = Field(it, "text", null);
            double room = t + h - padBottom - ty;
            if (text != null && text.Length > 0 && room >= 12)
                p.Text(l + 18, ty, w - 36, room, text, 12.5, KLadder(12.5, 9), false, "left", th.Muted, th.Font, name + "-t");
        }
        else if (kind == "panel")
        {
            string form = (it.Str("shape", "round") ?? "round").ToLowerInvariant();
            int fill = ColorOf(it.Str("fill", null), th.Surface);
            string text = Field(it, "text", null);
            if (form == "circle")
            {
                p.Oval(l, t, w, h, fill, name);
                if (text != null && text.Length > 0)
                    p.Text(l + 6, t + 6, w - 12, h - 12, text, 14, KLadder(14, 9), false, "center", SafeTextOn(th, fill), th.Font, name + "-t");
            }
            else if (form == "rect")
            {
                p.Shape(1, l, t, w, h, fill, name);
                if (text != null && text.Length > 0)
                    p.Text(l + 12, t + 8, w - 24, h - 16, text, 14, KLadder(14, 9), false, "left", SafeTextOn(th, fill), th.Font, name + "-t");
            }
            else
            {
                p.Block(l, t, w, h, fill, name);
                if (text != null && text.Length > 0)
                    p.Text(l + 16, t + 10, w - 32, h - 20, text, 14, KLadder(14, 9), false, "left", SafeTextOn(th, fill), th.Font, name + "-t");
            }
        }
        else if (kind == "line" || kind == "arrow")
        {
            // the box is a start point plus a delta, not a rectangle
            double x1 = l, y1 = t, x2 = l + w, y2 = t + h;
            bool arrow = kind == "arrow";
            if (it.Has("arrow")) arrow = it.Flag("arrow", true);
            KLine(p, x1, y1, x2, y2, ColorOf(it.Str("color", null), KLineColor(th)),
                  KDbl(it.Raw("weight"), 1.5), arrow, it.Flag("dashed", false), name);
        }
        else if (kind == "image")
        {
            string path = it.Str("path", null);
            if (path == null)
                p.Warnings.Add("item " + (index + 1).ToString(CultureInfo.InvariantCulture) + " (image) has no \"path\"");
            else
            {
                string fit = (it.Str("fit", "cover") ?? "cover").ToLowerInvariant();
                p.Picture(l, t, w, h, path, fit, name);
            }
        }
        else if (kind == "stat")
        {
            string value = Field(it, "value", "");
            string unit = Field(it, "unit", "");
            string label = Field(it, "label", "");
            double vs = Math.Min(48, Math.Max(20, Math.Min(h * 0.5, w * 0.28)));
            double used = vs * 1.4;
            if (label.Length > 0) used += KNeedHeight(label, p.C.WD(w), p.C.FS(13)) / Math.Max(p.C.Sy, 0.05);
            double top = t + ((h - used) > 28 ? (h - used) * 0.34 : 0);
            p.Text(l, top, w, Math.Max(18, vs * 1.35), unit.Length > 0 ? value + " " + unit : value, vs,
                   KLadder(vs, 14), true, "left", ColorOf(it.Str("color", null), th.Accent), th.TitleFont, name + "-v");
            if (label.Length > 0)
                p.Text(l, top + vs * 1.4, w, Math.Max(14, h - (top - t) - vs * 1.4), label, 13, KLadder(13, 9),
                       false, "left", th.Text, th.Font, name + "-l");
        }
        else if (kind == "bullets")
        {
            List<object> pts = KArr(it.Raw("points"));
            if (pts.Count == 0) pts = KArr(it.Raw("items"));
            List<string> lines = new List<string>();
            for (int i = 0; i < pts.Count; i++)
            {
                string one = PlainOf(pts[i]);
                if (one == null) continue;
                one = one.Trim();
                if (one.Length == 0) continue;
                lines.Add("\u2022 " + one);
            }
            if (lines.Count == 0)
                p.Warnings.Add("item " + (index + 1).ToString(CultureInfo.InvariantCulture) + " (bullets) has no points");
            else
            {
                double size = KDbl(it.Raw("size"), 14);
                p.Text(l, t, w, h, string.Join("\r", lines.ToArray()), size, KLadder(size, 9),
                       false, "left", th.Text, th.Font, name);
            }
        }
        else if (kind == "steps")
        {
            List<object> steps = KArr(it.Raw("steps"));
            if (steps.Count == 0) steps = KArr(it.Raw("items"));
            int n = steps.Count;
            if (n == 0)
                p.Warnings.Add("item " + (index + 1).ToString(CultureInfo.InvariantCulture) + " (steps) has no steps");
            else
            {
                double gap = 10;
                double cw = (w - gap * (n - 1)) / n;
                for (int i = 0; i < n; i++)
                {
                    Bag st = null;
                    try { st = new Bag(steps[i]); } catch (Exception) { }
                    double cx = l + i * (cw + gap);
                    double numS = Math.Min(20, Math.Max(12, cw * 0.28));
                    p.Oval(cx, t, numS, numS, i % 2 == 0 ? th.Primary : th.Accent, null);
                    p.Text(cx, t + numS * 0.12, numS, numS * 0.86, (i + 1).ToString(CultureInfo.InvariantCulture),
                           numS * 0.55, new double[] { numS * 0.55 }, true, "center", th.White, th.TitleFont, null);
                    string head = Field(st, "head", null);
                    if (head != null && head.Length > 0)
                        p.Text(cx, t + numS + 5, cw, 16, head, 12.5, KLadder(12.5, 9), true, "left", th.Text, th.Font, name + "-h" + i);
                    string tx = Field(st, "text", null);
                    if (tx != null && tx.Length > 0)
                        p.Text(cx, t + numS + 24, cw, Math.Max(12, h - numS - 26), tx, 11.5, KLadder(11.5, 9),
                               false, "left", th.Muted, th.Font, name + "-t" + i);
                }
            }
        }
        else if (kind == "callout")
        {
            int fill = ColorOf(it.Str("fill", null), KMix(th.Accent, th.White, 0.86));
            p.Block(l, t, w, h, fill, name);
            p.Shape(1, l, t, Math.Max(3, 4), h, th.Accent, null);
            p.Text(l + 16, t + 8, w - 30, h - 16, Field(it, "text", null), 13.5, KLadder(13.5, 9),
                   false, "left", th.Text, th.Font, name + "-t");
        }
        else if (kind == "quote")
        {
            p.Shape(1, l, t, 3, h, th.Accent, null);
            p.Text(l + 16, t, w - 16, h, Field(it, "text", null), 16, KLadder(16, 10), false, "left", th.Text, th.Font, name);
            string by = Field(it, "by", null);
            if (by != null && by.Length > 0)
                p.Text(l + 16, t + h - 18, w - 16, 16, "-- " + by, 11, new double[] { 11, 10 },
                       false, "right", th.Muted, th.Font, name + "-by");
        }
        else if (kind == "node")
        {
            int fill = ColorOf(it.Str("fill", null), th.HasCardFill ? th.CardFill : th.Surface);
            p.Block(l, t, w, h, fill, name);
            dynamic frame = p.Shape(5, l, t, w, h, -1, name + "-frame");
            try
            {
                frame.Line.Visible = -1;
                frame.Line.ForeColor.RGB = th.Accent;
                frame.Line.Weight = (float)Math.Max(0.75, 1.25 * p.C.S);
            }
            catch (Exception) { }
            string text = Field(it, "text", null);
            if (text != null && text.Length > 0)
                p.Text(l + 8, t + 4, w - 16, h - 8, text, 14, KLadder(14, 10), true, "center", SafeTextOn(th, fill), th.Font, name + "-t");
            KNode node = new KNode();
            node.Id = it.Str("id", tag);
            if (node.Id == null || node.Id.Length == 0) node.Id = tag;
            node.Shape = frame;
            node.L = l; node.T = t; node.W = w; node.H = h;
            if (nodes != null && !nodes.ContainsKey(node.Id)) nodes[node.Id] = node;
        }
        else if (kind == "chart")
        {
            Bag ch = null;
            try { ch = new Bag(it.Raw("chart")); }
            catch (Exception) { }
            if (ch == null) ch = it;   // a chart spec may sit straight on the item
            KChartPaint(p, ch, l, t, w, h, name);
        }
        else if (kind == "table")
        {
            p.Warnings.Add("item " + (index + 1).ToString(CultureInfo.InvariantCulture)
                + " (table) is not drawn by canvas/grid; add it with the `table` command on the finished page"
                + " (or use a 'steps' / 'cards' item)");
        }
        else
        {
            string text = Field(it, "text", null);
            p.Warnings.Add("item " + (index + 1).ToString(CultureInfo.InvariantCulture) + ": unknown type \"" + kind
                + "\" was drawn as plain text (known: text, heading, card, panel, line, arrow, image, stat, bullets,"
                + " steps, callout, quote, node, chart)");
            if (text != null && text.Length > 0)
                p.Text(l, t, w, h, text, 14, KLadder(14, 9), false, "left", th.Text, th.Font, name);
        }
    }

    // ---------------------------------------------------------------- shape-drawn charts

    /// <summary>Paint a chart out of shapes: axes, ticks, points / lines / slices, legend and
    /// axis titles.  Deliberately not a native chart object (see the file header).</summary>
    private static void KChartPaint(KindPage p, Bag ch, double l, double t, double w, double h, string name)
    {
        string type = (ch.Str("type", "scatter") ?? "scatter").Trim().ToLowerInvariant();
        if (p.ChartTypes != null) p.ChartTypes.Add(type);
        if (w < 110 || h < 80)
            p.Warnings.Add("chart \"" + name + "\" has only " + KNum(w, 0) + "x" + KNum(h, 0)
                + " design units of room; labels may collide (give it a bigger box)");
        if (type == "scatter") KChartScatter(p, ch, l, t, w, h, name);
        else if (type == "curve" || type == "line") KChartCurve(p, ch, l, t, w, h, name);
        else if (type == "pie" || type == "doughnut" || type == "donut") KChartPie(p, ch, l, t, w, h, name, type != "pie");
        else
        {
            p.Warnings.Add("chart type \"" + type + "\" is not drawn by this build (scatter, curve, pie, doughnut)");
            p.Text(l, t, w, h, "chart type \"" + type + "\" is not supported yet", 12,
                   new double[] { 12, 11 }, false, "center", p.Th.Muted, p.Th.Font, name + "-msg");
        }
    }

    private sealed class KAxis
    {
        public double Min, Max;
        public double P0, P1;      // slide position of min and of max

        public double At(double v)
        {
            double span = Max - Min;
            if (Math.Abs(span) < 1e-9) return (P0 + P1) / 2;
            return P0 + (v - Min) / span * (P1 - P0);
        }
    }

    /// <summary>Five evenly spaced tick values across the data range.</summary>
    private static double[] KTicks(double min, double max)
    {
        double[] t = new double[5];
        for (int i = 0; i < 5; i++) t[i] = min + (max - min) * i / 4.0;
        return t;
    }

    /// <summary>Raw data range.  The axis bounds themselves are rounded by KNiceTicks, so
    /// nothing is padded here; a flat series is widened so it cannot divide by zero.</summary>
    private static void KDataRange(List<double> vals, out double min, out double max)
    {
        min = 0; max = 1;
        if (vals.Count == 0) return;
        min = vals[0]; max = vals[0];
        for (int i = 1; i < vals.Count; i++)
        {
            if (vals[i] < min) min = vals[i];
            if (vals[i] > max) max = vals[i];
        }
        if (max - min < 1e-9)
        {
            double span = Math.Max(Math.Abs(max) * 0.2, 1);
            min -= span / 2;
            max += span / 2;
        }
    }

    /// <summary>Round the axis bounds outward to a 1 / 2 / 5 x 10^k step, so the labels read
    /// like a real chart (0, 500, 1000 ...) instead of 1914.3.  Returns the tick values and,
    /// through lo / hi / digits, the bounds the plot must use and the decimals it needs.</summary>
    private static double[] KNiceTicks(double min, double max, bool nonNegative,
                                      out double lo, out double hi, out int digits)
    {
        if (max - min < 1e-9) { min -= 0.5; max += 0.5; }
        if (max <= min) max = min + 1;
        double raw = (max - min) / 4.0;
        double mag = Math.Pow(10, Math.Floor(Math.Log10(raw)));
        double norm = raw / mag;
        double step = mag * (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10);
        lo = Math.Floor(min / step) * step;
        hi = Math.Ceiling(max / step) * step;
        if (nonNegative && lo < 0) lo = 0;
        int count = (int)Math.Round((hi - lo) / step) + 1;
        if (count > 9)
        {
            step *= 2;
            lo = Math.Floor(min / step) * step;
            hi = Math.Ceiling(max / step) * step;
            if (nonNegative && lo < 0) lo = 0;
            count = (int)Math.Round((hi - lo) / step) + 1;
        }
        if (count < 2) { count = 2; hi = lo + step; }
        digits = step >= 1 ? 0 : (step >= 0.1 ? 1 : 2);
        double[] ticks = new double[count];
        for (int i = 0; i < count; i++) ticks[i] = lo + i * step;
        return ticks;
    }

    private static string KAxisLabel(double v, int digits)
    {
        string s = KNum(v, digits);
        if (s == "-0") s = "0";
        return s;
    }

    /// <summary>Plot frame shared by the cartesian charts: axes, ticks and axis titles.</summary>
    private static void KFrame(KindPage p, double l, double t, double w, double h,
                               string xTitle, string yTitle,
                               double xmin, double xmax, double ymin, double ymax,
                               bool nonNegX, bool nonNegY,
                               out KAxis ax, out KAxis ay, string name)
    {
        DeckTheme th = p.Th;
        double padL = 42, padR = 12, padT = 18, padB = 30;
        double px0 = l + padL, px1 = l + w - padR, py0 = t + padT, py1 = t + h - padB;
        KLine(p, px0, py1, px1, py1, th.Muted, 1.0, false, false, name + "-xaxis");
        KLine(p, px0, py0, px0, py1, th.Muted, 1.0, false, false, name + "-yaxis");

        double xlo, xhi, ylo, yhi;
        int xd, yd;
        double[] xs = KNiceTicks(xmin, xmax, nonNegX, out xlo, out xhi, out xd);
        double[] ys = KNiceTicks(ymin, ymax, nonNegY, out ylo, out yhi, out yd);

        ax = new KAxis();
        ax.Min = xlo; ax.Max = xhi; ax.P0 = px0; ax.P1 = px1;
        ay = new KAxis();
        ay.Min = ylo; ay.Max = yhi; ay.P0 = py1; ay.P1 = py0;

        for (int i = 0; i < xs.Length; i++)
        {
            double x = ax.At(xs[i]);
            KLine(p, x, py1, x, py1 + 3, th.Muted, 0.75, false, false, null);
            // the first and the last label are aligned to the plot edge, otherwise they hang
            // over the y labels / off the chart
            string align = i == 0 ? "left" : (i == xs.Length - 1 ? "right" : "center");
            double lx = i == 0 ? x : (i == xs.Length - 1 ? x - 52 : x - 26);
            p.Text(lx, py1 + 2, 52, 14, KAxisLabel(xs[i], xd), 8.5, new double[] { 8.5 },
                   false, align, th.Muted, th.Font, null);
        }
        for (int i = 0; i < ys.Length; i++)
        {
            double y = ay.At(ys[i]);
            KLine(p, px0 - 3, y, px0, y, th.Muted, 0.75, false, false, null);
            p.Text(l, y - 6.5, padL - 5, 14, KAxisLabel(ys[i], yd), 8.5,
                   new double[] { 8.5 }, false, "right", th.Muted, th.Font, null);
        }
        if (xTitle != null && xTitle.Length > 0)
            p.Text(px0, t + h - 13, px1 - px0, 14, xTitle, 9, new double[] { 9 }, false, "center", th.Text, th.Font, name + "-xtitle");
        if (yTitle != null && yTitle.Length > 0)
            p.Text(l, t - 1, Math.Max(60, Math.Min(200, w - 20)), 14, yTitle, 9, new double[] { 9 },
                   false, "left", th.Text, th.Font, name + "-ytitle");
    }

    private static void KChartScatter(KindPage p, Bag ch, double l, double t, double w, double h, string name)
    {
        DeckTheme th = p.Th;
        List<object> series = KArr(ch.Raw("series"));
        List<List<double[]>> pts = new List<List<double[]>>();
        List<string> names = new List<string>();
        List<double> allX = new List<double>(), allY = new List<double>();
        for (int i = 0; i < series.Count; i++)
        {
            Bag s = null;
            try { s = new Bag(series[i]); } catch (Exception) { }
            List<double[]> list = new List<double[]>();
            List<object> raw = KArr(s == null ? null : s.Raw("points"));
            for (int k = 0; k < raw.Count; k++)
            {
                List<object> pair = raw[k] as List<object>;
                if (pair == null || pair.Count < 2) continue;
                double x = KDbl(pair[0], 0), y = KDbl(pair[1], 0);
                list.Add(new double[] { x, y });
                allX.Add(x); allY.Add(y);
            }
            pts.Add(list);
            names.Add(s == null ? "series " + (i + 1).ToString(CultureInfo.InvariantCulture)
                                : Field(s, "name", "series " + (i + 1).ToString(CultureInfo.InvariantCulture)));
        }
        if (allX.Count == 0)
        {
            p.Warnings.Add("scatter chart \"" + name + "\" has no points");
            return;
        }
        double xmin, xmax, ymin, ymax;
        KDataRange(allX, out xmin, out xmax);
        KDataRange(allY, out ymin, out ymax);
        KAxis ax, ay;
        KFrame(p, l, t, w, h, Field(ch, "xTitle", null), Field(ch, "yTitle", null),
               xmin, xmax, ymin, ymax, xmin >= 0, ymin >= 0, out ax, out ay, name);

        for (int i = 0; i < pts.Count; i++)
        {
            int color = KSeriesColor(th, i);
            for (int k = 0; k < pts[i].Count; k++)
                p.Oval(ax.At(pts[i][k][0]) - 2.6, ay.At(pts[i][k][1]) - 2.6, 5.2, 5.2, color, null);
            // series key along the top of the plot (a scatter has no line end to label)
            double kx = l + w - 8 - (pts.Count - i) * 96;
            if (kx + 20 < l + w)
            {
                p.Shape(1, kx, t + 2, 8, 8, color, null);
                p.Text(kx + 11, t - 1, 78, 14, names[i], 8.5, new double[] { 8.5 }, false, "left", th.Muted, th.Font, null);
            }
        }

        if (!ch.Flag("fit", false)) return;
        if (allX.Count < 2)
        {
            p.Warnings.Add("scatter chart \"" + name + "\": a fit line needs at least 2 points");
            return;
        }
        double sx = 0, sy = 0, sxx = 0, sxy = 0;
        int n = allX.Count;
        for (int i = 0; i < n; i++) { sx += allX[i]; sy += allY[i]; sxx += allX[i] * allX[i]; sxy += allX[i] * allY[i]; }
        double denom = n * sxx - sx * sx;
        if (Math.Abs(denom) < 1e-12)
        {
            p.Warnings.Add("scatter chart \"" + name + "\": every x is the same, no fit line");
            return;
        }
        double slope = (n * sxy - sx * sy) / denom;
        double intercept = (sy - slope * sx) / n;
        double meanY = sy / n;
        double sse = 0, sst = 0;
        for (int i = 0; i < n; i++)
        {
            double pred = intercept + slope * allX[i];
            sse += (allY[i] - pred) * (allY[i] - pred);
            sst += (allY[i] - meanY) * (allY[i] - meanY);
        }
        double r2 = sst < 1e-12 ? 1 : 1 - sse / sst;
        double fx0 = ax.Min, fx1 = ax.Max;
        double fy0 = intercept + slope * fx0, fy1 = intercept + slope * fx1;
        double lo = ay.Min, hi = ay.Max;
        if (fy0 < lo || fy0 > hi || fy1 < lo || fy1 > hi)
        {
            double gx0 = (lo - intercept) / (Math.Abs(slope) < 1e-12 ? 1e-12 : slope);
            double gx1 = (hi - intercept) / (Math.Abs(slope) < 1e-12 ? 1e-12 : slope);
            if (gx0 > gx1) { double tmp = gx0; gx0 = gx1; gx1 = tmp; }
            if (gx0 > ax.Min) fx0 = gx0;
            if (gx1 < ax.Max) fx1 = gx1;
            fy0 = intercept + slope * fx0; fy1 = intercept + slope * fx1;
        }
        KLine(p, ax.At(fx0), ay.At(fy0), ax.At(fx1), ay.At(fy1), th.Accent, 2.0, false, false, name + "-fit");
        string eq = "y = " + KNum(slope, 3) + "x " + (intercept < 0 ? "- " : "+ ") + KNum(Math.Abs(intercept), 3)
            + "   R\u00b2 = " + KNum(r2, 3);
        // the equation sits inside the plot just above the x axis on the right: the top band
        // already carries the axis title and the series key
        p.Text(l + w - 206, ay.P0 - 21, 196, 14, eq, 9.5, new double[] { 9.5 },
               false, "right", th.Accent, th.Font, name + "-eq");
    }

    private static void KChartCurve(KindPage p, Bag ch, double l, double t, double w, double h, string name)
    {
        DeckTheme th = p.Th;
        List<object> series = KArr(ch.Raw("series"));
        List<List<double[]>> pts = new List<List<double[]>>();
        List<string> names = new List<string>();
        List<double> allX = new List<double>(), allY = new List<double>();
        for (int i = 0; i < series.Count; i++)
        {
            Bag s = null;
            try { s = new Bag(series[i]); } catch (Exception) { }
            List<double> xs = new List<double>(), ys = new List<double>();
            List<object> rx = KArr(s == null ? null : s.Raw("x"));
            List<object> ry = KArr(s == null ? null : s.Raw("y"));
            if (rx.Count == 0)
            {
                // also accept points: [[x,y],...] for symmetry with scatter
                List<object> raw = KArr(s == null ? null : s.Raw("points"));
                for (int k = 0; k < raw.Count; k++)
                {
                    List<object> pair = raw[k] as List<object>;
                    if (pair == null || pair.Count < 2) continue;
                    xs.Add(KDbl(pair[0], 0)); ys.Add(KDbl(pair[1], 0));
                }
            }
            else
            {
                for (int k = 0; k < rx.Count; k++) xs.Add(KDbl(rx[k], 0));
                for (int k = 0; k < ry.Count; k++) ys.Add(KDbl(ry[k], 0));
            }
            int m = Math.Min(xs.Count, ys.Count);
            List<double[]> list = new List<double[]>();
            for (int k = 0; k < m; k++)
            {
                list.Add(new double[] { xs[k], ys[k] });
                allX.Add(xs[k]); allY.Add(ys[k]);
            }
            pts.Add(list);
            names.Add(s == null ? "series " + (i + 1).ToString(CultureInfo.InvariantCulture)
                                : Field(s, "name", "series " + (i + 1).ToString(CultureInfo.InvariantCulture)));
        }
        if (allX.Count == 0)
        {
            p.Warnings.Add("curve chart \"" + name + "\" has no points");
            return;
        }
        double xmin, xmax, ymin, ymax;
        KDataRange(allX, out xmin, out xmax);
        KDataRange(allY, out ymin, out ymax);
        KAxis ax, ay;
        KFrame(p, l, t, w, h, Field(ch, "xTitle", null), Field(ch, "yTitle", null),
               xmin, xmax, ymin, ymax, xmin >= 0, ymin >= 0, out ax, out ay, name);

        bool smooth = ch.Flag("smooth", true);
        List<double[]> labelSpots = new List<double[]>();
        for (int i = 0; i < pts.Count; i++)
        {
            int color = KSeriesColor(th, i);
            List<double[]> list = pts[i];
            for (int k = 1; k < list.Count; k++)
            {
                double x1 = ax.At(list[k - 1][0]), y1 = ay.At(list[k - 1][1]);
                double x2 = ax.At(list[k][0]), y2 = ay.At(list[k][1]);
                if (smooth)
                {
                    // two short segments through a midpoint keep the join soft without a freeform
                    double mx = (x1 + x2) / 2, my = (y1 + y2) / 2;
                    KLine(p, x1, y1, mx, my, color, 2.0, false, false, null);
                    KLine(p, mx, my, x2, y2, color, 2.0, false, false, null);
                }
                else
                {
                    KLine(p, x1, y1, x2, y2, color, 2.0, false, false, null);
                }
            }
            if (list.Count <= 12)
                for (int k = 0; k < list.Count; k++)
                    p.Oval(ax.At(list[k][0]) - 2.2, ay.At(list[k][1]) - 2.2, 4.4, 4.4, color, null);
            if (list.Count > 0)
            {
                double ex = ax.At(list[list.Count - 1][0]), ey = ay.At(list[list.Count - 1][1]);
                double ly = ey - 7;
                for (int g = 0; g < labelSpots.Count; g++)
                    if (Math.Abs(labelSpots[g][1] - ly) < 10 && Math.Abs(labelSpots[g][0] - ex) < 60) ly -= 11;
                labelSpots.Add(new double[] { ex, ly });
                double lx = ex + 4;                       // keep the series name inside the chart
                if (lx + 60 > l + w) lx = ex - 64;
                if (ly < t + 4) ly = ey + 6;
                p.Text(lx, ly, 62, 14, names[i], 9, new double[] { 9 }, false, "left", color, th.Font, name + "-label" + i);
            }
        }
        if (pts.Count > 1)
        {
            for (int i = 0; i < pts.Count; i++)
            {
                double kx = l + w - 8 - (pts.Count - i) * 96;
                if (kx + 20 > l + w) break;
                p.Shape(1, kx, t + 2, 8, 8, KSeriesColor(th, i), null);
                p.Text(kx + 11, t - 1, 78, 14, names[i], 8.5, new double[] { 8.5 }, false, "left", th.Muted, th.Font, null);
            }
        }
    }

    private static void KChartPie(KindPage p, Bag ch, double l, double t, double w, double h, string name, bool doughnut)
    {
        DeckTheme th = p.Th;
        List<object> cats = KArr(ch.Raw("categories"));
        List<object> vals = KArr(ch.Raw("values"));
        int n = Math.Min(cats.Count, vals.Count);
        List<string> labels = new List<string>();
        List<double> values = new List<double>();
        double total = 0;
        for (int i = 0; i < n; i++)
        {
            double v = KDbl(vals[i], 0);
            if (v < 0) v = 0;
            labels.Add(Convert.ToString(cats[i], CultureInfo.InvariantCulture));
            values.Add(v);
            total += v;
        }
        if (n == 0 || total <= 0)
        {
            p.Warnings.Add("pie chart \"" + name + "\" needs categories and positive values");
            return;
        }
        bool legend = ch.Flag("legend", true);
        double legendW = legend ? Math.Min(96, w * 0.34) : 0;
        double areaW = w - legendW;
        double d = Math.Min(areaW - 8, h - 8);
        if (d < 40) d = Math.Max(24, Math.Min(areaW, h));
        // centre the pie + legend group in the box: a circle that only fills half the width and
        // then a legend glued to the far edge reads as two unrelated objects
        double groupW = d + (legend ? legendW + 12 : 0);
        double groupX = l + (w - groupW) / 2;
        if (groupX < l) groupX = l;
        double areaX = groupX;
        double cx = areaX + d / 2, cy = t + h / 2, r = d / 2;
        double inner = doughnut ? r * 0.55 : 0;

        double angle = -90;                     // start at 12 o'clock, clockwise
        for (int i = 0; i < n; i++)
        {
            double frac = values[i] / total;
            double sweep = frac * 360;
            int color = KSeriesColor(th, i);
            List<double[]> poly = new List<double[]>();
            if (!doughnut) poly.Add(new double[] { cx, cy });
            int steps = Math.Max(2, (int)Math.Ceiling(Math.Abs(sweep) / 6));
            for (int k = 0; k <= steps; k++)
            {
                double a = (angle + sweep * k / steps) * Math.PI / 180.0;
                poly.Add(new double[] { cx + r * Math.Cos(a), cy + r * Math.Sin(a) });
            }
            if (doughnut)
                for (int k = steps; k >= 0; k--)
                {
                    double a = (angle + sweep * k / steps) * Math.PI / 180.0;
                    poly.Add(new double[] { cx + inner * Math.Cos(a), cy + inner * Math.Sin(a) });
                }
            dynamic slice = KFreeform(p, poly, color, name + "-slice" + i);
            if (slice == null)
            {
                // fallback: the pie shape with adjustments (start / end angle in degrees)
                try
                {
                    slice = p.Shape(142, cx - r, cy - r, 2 * r, 2 * r, color, name + "-slice" + i);
                    try { slice.Adjustments[1] = (float)angle; } catch (Exception) { }
                    try { slice.Adjustments[2] = (float)(angle + sweep); } catch (Exception) { }
                }
                catch (Exception) { }
                if (i == 0)
                    p.Warnings.Add("pie chart \"" + name + "\": BuildFreeform was refused, slices fall back to the pie shape"
                        + (doughnut ? " (drawn as a filled pie - the doughnut hole is lost)" : ""));
            }
            if (frac >= 0.07)
            {
                double mid = (angle + sweep / 2) * Math.PI / 180.0;
                double lr = doughnut ? (r + inner) / 2 : r * 0.62;
                p.Text(cx + lr * Math.Cos(mid) - 20, cy + lr * Math.Sin(mid) - 8, 40, 16,
                       KNum(frac * 100, 0) + "%", 9.5, new double[] { 9.5 }, true, "center", th.White, th.Font, null);
            }
            angle += sweep;
        }
        if (doughnut)
            p.Oval(cx - inner, cy - inner, inner * 2, inner * 2, th.White, name + "-hole");
        if (legend)
        {
            double lxp = areaX + d + 12;
            double rowH = Math.Min(18, (h - 4) / Math.Max(1, n));
            double ly = cy - rowH * n / 2;
            for (int i = 0; i < n; i++)
            {
                p.Shape(1, lxp, ly + rowH / 2 - 4, 8, 8, KSeriesColor(th, i), null);
                p.Text(lxp + 11, ly + rowH / 2 - 8, Math.Max(30, legendW - 14), Math.Max(14, rowH),
                       labels[i] + " " + KNum(values[i], 0), 8.5, new double[] { 8.5 },
                       false, "left", th.Text, th.Font, null);
                ly += rowH;
            }
        }
    }

    // ---------------------------------------------------------------- edges

    /// <summary>Join node items with real connectors (BeginConnect / EndConnect, so the line
    /// follows the box when it moves).  Every edge that fails to attach is reported.</summary>
    private static void KDrawEdges(KindPage p, List<object> edges, Dictionary<string, KNode> nodes)
    {
        if (edges == null || edges.Count == 0) return;
        for (int i = 0; i < edges.Count; i++)
        {
            Bag e = null;
            try { e = new Bag(edges[i]); } catch (Exception) { }
            if (e == null) continue;
            string from = e.Str("from", null);
            string to = e.Str("to", null);
            Dictionary<string, object> rec = new Dictionary<string, object>();
            rec["from"] = from;
            rec["to"] = to;
            if (from == null || to == null)
            {
                rec["connected"] = new Dictionary<string, object> { { "begin", false }, { "end", false } };
                rec["note"] = "edge needs \"from\" and \"to\"";
                p.Warnings.Add("edge " + (i + 1).ToString(CultureInfo.InvariantCulture) + " has no from/to");
                if (p.Edges != null) p.Edges.Add(rec);
                continue;
            }
            KNode a = nodes.ContainsKey(from) ? nodes[from] : null;
            KNode b = nodes.ContainsKey(to) ? nodes[to] : null;
            if (a == null || b == null)
            {
                string missing = a == null ? from : to;
                rec["connected"] = new Dictionary<string, object> { { "begin", false }, { "end", false } };
                rec["note"] = "no node with id \"" + missing + "\"";
                p.Warnings.Add("edge " + from + " -> " + to + ": no node with id \"" + missing
                    + "\" (add a {\"type\":\"node\",\"id\":\"" + missing + "\"} item)");
                if (p.Edges != null) p.Edges.Add(rec);
                continue;
            }

            string kindWord = (e.Str("kind", "straight") ?? "straight").Trim().ToLowerInvariant();
            int kind = 1;
            if (kindWord == "elbow" || kindWord == "angled") kind = 2;
            else if (kindWord == "curve" || kindWord == "curved") kind = 3;

            string autoFrom, autoTo;
            SideAuto(a.Shape, b.Shape, out autoFrom, out autoTo);
            string gotFrom, gotTo;
            int siteFrom = SideSite(e.Str("fromSide", null), out gotFrom);
            int siteTo = SideSite(e.Str("toSide", null), out gotTo);
            if (siteFrom == 0) { siteFrom = SideSite(autoFrom, out gotFrom); }
            if (siteTo == 0) { siteTo = SideSite(autoTo, out gotTo); }

            double[] pa = Anchor(a.Shape, gotFrom);
            double[] pb = Anchor(b.Shape, gotTo);
            dynamic link = null;
            try
            {
                link = p.Slide.Shapes.AddConnector(kind, (float)pa[0], (float)pa[1], (float)pb[0], (float)pb[1]);
            }
            catch (Exception ex)
            {
                rec["connected"] = new Dictionary<string, object> { { "begin", false }, { "end", false } };
                rec["note"] = "AddConnector failed: " + Clean(ex.Message);
                p.Warnings.Add("edge " + from + " -> " + to + ": AddConnector failed (" + Clean(ex.Message) + ")");
                if (p.Edges != null) p.Edges.Add(rec);
                continue;
            }
            bool beginOk = false, endOk = false;
            try
            {
                link.ConnectorFormat.BeginConnect(a.Shape, siteFrom);
                beginOk = ToD(TryGet(() => (object)link.ConnectorFormat.BeginConnected, (object)0)) != 0;
            }
            catch (Exception) { beginOk = false; }
            try
            {
                link.ConnectorFormat.EndConnect(b.Shape, siteTo);
                endOk = ToD(TryGet(() => (object)link.ConnectorFormat.EndConnected, (object)0)) != 0;
            }
            catch (Exception) { endOk = false; }
            try { link.RerouteConnections(); } catch (Exception) { }
            try
            {
                link.Line.Visible = -1;
                link.Line.ForeColor.RGB = ColorOf(e.Str("color", null), p.Th.Muted);
                link.Line.Weight = (float)Math.Max(0.75, KDbl(e.Raw("weight"), 1.25) * p.C.S);
                if (e.Flag("arrow", true)) link.Line.EndArrowheadStyle = 2;
                if (e.Flag("dashed", false)) link.Line.DashStyle = 4;
            }
            catch (Exception) { }
            try { link.Name = "edge" + (i + 1).ToString(CultureInfo.InvariantCulture); } catch (Exception) { }

            Dictionary<string, object> conn = new Dictionary<string, object>();
            conn["begin"] = beginOk;
            conn["end"] = endOk;
            rec["connected"] = conn;
            rec["kind"] = kindWord;
            rec["fromSide"] = gotFrom;
            rec["toSide"] = gotTo;
            if (!beginOk || !endOk)
                p.Warnings.Add("edge " + from + " -> " + to + " is not attached (begin=" + beginOk + ", end=" + endOk
                    + "); the line will not follow the boxes when they move");

            string label = e.Str("label", null);
            if (label != null && label.Length > 0)
            {
                double cl = ToD(TryGet(() => (object)link.Left, (object)0));
                double ct = ToD(TryGet(() => (object)link.Top, (object)0));
                double cw = ToD(TryGet(() => (object)link.Width, (object)0));
                double chh = ToD(TryGet(() => (object)link.Height, (object)0));
                double size = KDbl(e.Raw("labelSize"), 9.5);
                double tw = Math.Min(140, Math.Max(36, KTextWidth(label, size) + 8));
                double sx = Math.Max(p.C.Sx, 0.05), sy = Math.Max(p.C.Sy, 0.05);
                double labX = (cl + cw / 2 - tw / 2) / sx, labY = (ct + chh / 2 - 9) / sy;
                // two edges between the same pair of boxes share a midpoint: nudge the second
                // label down so both stay readable
                for (int g = 0; g < p.LabelSpots.Count; g++)
                {
                    double[] spot = p.LabelSpots[g];
                    if (Math.Abs(spot[0] - labX) < tw && Math.Abs(spot[1] - labY) < 16) labY = spot[1] + 17;
                }
                p.LabelSpots.Add(new double[] { labX, labY });
                dynamic plate = p.Text(labX, labY, tw / sx, 19 / sy,
                       label, size, new double[] { size }, false, "center", p.Th.Text, p.Th.Font,
                       "edgeLabel" + (i + 1).ToString(CultureInfo.InvariantCulture));
                // a white plate behind the words, so the connector does not run through them
                if (plate != null)
                {
                    try
                    {
                        plate.Fill.Solid();
                        plate.Fill.ForeColor.RGB = p.Th.White;
                        plate.Fill.Transparency = 0f;
                        plate.Line.Visible = 0;
                    }
                    catch (Exception) { }
                }
                rec["label"] = label;
            }
            if (p.Edges != null) p.Edges.Add(rec);
        }
    }

    // ---------------------------------------------------------------- canvas preset

    private static void KCanvas(KindPage p, Bag a, Bag doc, string title)
    {
        p.Left = 48;
        p.Right = 912;
        p.BodyTop = (title != null && title.Trim().Length > 0) ? 130 : 40;
        p.BodyBottom = 494;

        List<object> items = new List<object>();
        if (doc != null)
        {
            if (doc.Has("items")) items = doc.List("items");
            else if (doc.Has("canvas")) items = doc.List("canvas");
        }
        else items = a.List("items");
        if (items.Count == 0)
        {
            p.Warnings.Add("canvas has no items (pass {\"title\":\"...\",\"items\":[{type,box,...}],\"edges\":[...]})");
            return;
        }

        Dictionary<string, KNode> nodes = new Dictionary<string, KNode>(StringComparer.Ordinal);
        for (int i = 0; i < items.Count; i++)
        {
            Bag it = null;
            try { it = new Bag(items[i]); }
            catch (Exception)
            {
                p.Warnings.Add("item " + (i + 1).ToString(CultureInfo.InvariantCulture) + " is not an object; skipped");
                continue;
            }
            string kind = KItemKind(it);
            double[] box = KFitBox(p, it, kind, i, null);
            if (box == null)
            {
                p.Warnings.Add("item " + (i + 1).ToString(CultureInfo.InvariantCulture) + " (" + kind + ") has no box [x,y,w,h]; skipped");
                continue;
            }
            KDrawItem(p, it, box[0], box[1], box[2], box[3], nodes, "c" + (i + 1).ToString(CultureInfo.InvariantCulture), i);
        }

        List<object> edges = new List<object>();
        if (doc != null && doc.Has("edges")) edges = doc.List("edges");
        else edges = a.List("edges");
        KDrawEdges(p, edges, nodes);
    }

    // ---------------------------------------------------------------- grid preset

    /// <summary>How tall a part needs to be for its own text, in design units.  Charts and
    /// pictures ask for a floor and then absorb whatever height is left over.</summary>
    private static double KPartNeed(KindPage p, Bag it, double w)
    {
        string kind = KItemKind(it);
        double inner = Math.Max(20, w - 24);
        double say = Math.Max(p.C.Sy, 0.05);
        if (kind == "heading") return 28;
        if (kind == "text")
        {
            double size = KDbl(it.Raw("size"), 14);
            return KNeedHeight(Field(it, "text", ""), p.C.WD(inner), p.C.FS(size)) / say + 6;
        }
        if (kind == "bullets")
        {
            List<object> pts = KArr(it.Raw("points"));
            if (pts.Count == 0) pts = KArr(it.Raw("items"));
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < pts.Count; i++)
            {
                if (i > 0) sb.Append("\r");
                sb.Append("\u2022 ").Append(PlainOf(pts[i]));
            }
            return KNeedHeight(sb.ToString(), p.C.WD(inner), p.C.FS(13)) / say + 6;
        }
        if (kind == "stat") return Field(it, "label", "").Length > 0 ? 78 : 58;
        if (kind == "card")
        {
            double vs, hs, padTop, padBottom;
            KCardMetrics(out vs, out hs, out padTop, out padBottom);
            double need = padTop + padBottom;
            if (Field(it, "value", null) != null) need += vs * 1.45;
            if (Field(it, "head", null) != null) need += hs * 1.5;
            string text = Field(it, "text", null);
            if (text != null && text.Length > 0)
                need += KNeedHeight(text, p.C.WD(inner - 12), p.C.FS(12.5)) / say + 4;
            return need;
        }
        if (kind == "panel")
        {
            string text = Field(it, "text", null);
            if (text == null || text.Length == 0) return 40;
            return KNeedHeight(text, p.C.WD(inner - 16), p.C.FS(14)) / say + 22;
        }
        if (kind == "callout")
            return KNeedHeight(Field(it, "text", ""), p.C.WD(inner - 16), p.C.FS(13.5)) / say + 20;
        if (kind == "quote")
        {
            double need = KNeedHeight(Field(it, "text", ""), p.C.WD(inner - 16), p.C.FS(16)) / say + 10;
            if (Field(it, "by", "").Length > 0) need += 18;
            return need;
        }
        if (kind == "steps")
        {
            List<object> steps = KArr(it.Raw("steps"));
            int n = Math.Max(1, steps.Count);
            double cw = (inner - 10 * (n - 1)) / n;
            double worst = 0;
            for (int i = 0; i < steps.Count; i++)
            {
                Bag st = null;
                try { st = new Bag(steps[i]); } catch (Exception) { }
                double need = 44 + KNeedHeight(Field(st, "text", ""), p.C.WD(cw), p.C.FS(11.5)) / say;
                if (need > worst) worst = need;
            }
            return Math.Max(60, worst);
        }
        if (kind == "node") return 52;
        if (kind == "line" || kind == "arrow") return 10;
        if (kind == "image") return 120;
        if (kind == "chart") return 150;
        if (kind == "stack") return 60;
        return 40;
    }

    /// <summary>True when a part would rather absorb leftover height than dictate its own.</summary>
    private static bool KPartGrow(Bag it)
    {
        if (it == null) return false;
        string kind = KItemKind(it);
        if (kind == "chart" || kind == "image" || kind == "panel" || kind == "stack") return true;
        if (it.Has("grow")) return it.Flag("grow", true);
        return false;
    }

    /// <summary>Split available height for a stacked cell: text-ish rows keep their measured
    /// height, the flexible ones share the rest.</summary>
    private static void KStack(KindPage p, List<object> items, double l, double t, double w, double h,
                               Dictionary<string, KNode> nodes, string tag)
    {
        int n = items.Count;
        if (n == 0) return;
        double gap = 10;
        double[] need = new double[n];
        bool[] grow = new bool[n];
        double total = 0;
        for (int i = 0; i < n; i++)
        {
            Bag it = null;
            try { it = new Bag(items[i]); } catch (Exception) { }
            need[i] = Math.Max(18, KPartNeed(p, it, w));
            grow[i] = KPartGrow(it);
            total += need[i];
        }
        double avail = h - gap * (n - 1);
        if (avail < 20) avail = 20;
        if (total > avail)
        {
            double scale = avail / total;
            for (int i = 0; i < n; i++) need[i] *= scale;
            p.Warnings.Add("stack \"" + tag + "\" needs " + KNum(total, 0) + " design units but has " + KNum(avail, 0)
                + "; its items were squeezed");
        }
        else
        {
            double spare = avail - total;
            int growCount = 0;
            for (int i = 0; i < n; i++) if (grow[i]) growCount++;
            if (growCount > 0)
            {
                for (int i = 0; i < n; i++) if (grow[i]) need[i] += spare / growCount;
            }
            // only a flexible item (chart / picture / panel) takes the spare height: stretching
            // plain text just opens a gap between one line and the next
        }
        double y = t;
        for (int i = 0; i < n; i++)
        {
            Bag it = null;
            try { it = new Bag(items[i]); } catch (Exception) { }
            if (it != null)
                KDrawItem(p, it, l, y, w, need[i], nodes, tag + "-s" + (i + 1).ToString(CultureInfo.InvariantCulture), i);
            y += need[i] + gap;
        }
    }

    private static void KGrid(KindPage p, Bag a, Bag doc, string title)
    {
        p.Left = 48;
        p.Right = 912;
        p.BodyTop = 130;
        p.BodyBottom = 494;

        List<object> rows = new List<object>();
        if (doc != null && doc.Has("rows")) rows = doc.List("rows");
        else rows = a.List("rows");
        if (rows.Count == 0)
        {
            p.Warnings.Add("grid has no rows (pass {\"title\":\"...\",\"rows\":[[part,...],[part]]})");
            return;
        }
        double gutter = 18;
        int rowCount = rows.Count;
        List<List<object>> rowParts = new List<List<object>>();
        List<int> rowSpans = new List<int>();
        double[] need = new double[rowCount];
        bool[] grow = new bool[rowCount];
        double totalNeed = 0;
        for (int r = 0; r < rowCount; r++)
        {
            List<object> parts = KArr(rows[r]);
            rowParts.Add(parts);
            int spans = 0;
            double worst = 0;
            bool anyGrow = false;
            for (int k = 0; k < parts.Count; k++)
            {
                Bag it = null;
                try { it = new Bag(parts[k]); } catch (Exception) { }
                spans += Math.Max(1, (int)KDbl(it == null ? null : it.Raw("span"), 1));
                if (KPartGrow(it)) anyGrow = true;
            }
            if (spans < 1) spans = 1;
            rowSpans.Add(spans);
            double cellW = (p.BodyW - gutter * (Math.Max(1, parts.Count) - 1)) / spans;
            for (int k = 0; k < parts.Count; k++)
            {
                Bag it = null;
                try { it = new Bag(parts[k]); } catch (Exception) { }
                int span = Math.Max(1, (int)KDbl(it == null ? null : it.Raw("span"), 1));
                double needHere = KPartNeed(p, it, cellW * span);
                if (needHere > worst) worst = needHere;
            }
            need[r] = Math.Max(24, worst);
            grow[r] = anyGrow;
            totalNeed += need[r];
        }
        double availRows = p.BodyH - gutter * (rowCount - 1);
        if (availRows < 30) availRows = 30;
        if (totalNeed > availRows)
        {
            double scale = availRows / totalNeed;
            for (int r = 0; r < rowCount; r++) need[r] *= scale;
            p.Warnings.Add("grid rows need " + KNum(totalNeed, 0) + " design units but only " + KNum(availRows, 0)
                + " are free; every row was squeezed (drop a row or shorten the text)");
        }
        else
        {
            double spare = availRows - totalNeed;
            int growCount = 0;
            for (int r = 0; r < rowCount; r++) if (grow[r]) growCount++;
            if (growCount > 0)
            {
                for (int r = 0; r < rowCount; r++) if (grow[r]) need[r] += spare / growCount;
            }
            else
            {
                for (int r = 0; r < rowCount; r++) need[r] += spare * (need[r] / Math.Max(totalNeed, 1));
            }
        }

        Dictionary<string, KNode> nodes = new Dictionary<string, KNode>(StringComparer.Ordinal);
        double y = p.BodyTop;
        for (int r = 0; r < rowCount; r++)
        {
            List<object> parts = rowParts[r];
            int spans = rowSpans[r];
            double cellW = (p.BodyW - gutter * (Math.Max(1, parts.Count) - 1)) / spans;
            double x = p.Left;
            for (int k = 0; k < parts.Count; k++)
            {
                Bag it = null;
                try { it = new Bag(parts[k]); } catch (Exception) { }
                int span = Math.Max(1, (int)KDbl(it == null ? null : it.Raw("span"), 1));
                double cw = cellW * span;
                string tag = "g" + (r + 1).ToString(CultureInfo.InvariantCulture) + "-" + (k + 1).ToString(CultureInfo.InvariantCulture);
                if (it == null)
                    p.Warnings.Add("grid row " + (r + 1).ToString(CultureInfo.InvariantCulture) + " part "
                        + (k + 1).ToString(CultureInfo.InvariantCulture) + " is not an object; skipped");
                else if (KItemKind(it) == "stack")
                    KStack(p, KArr(it.Raw("items")), x, y, cw, need[r], nodes, tag);
                else
                    KDrawItem(p, it, x, y, cw, need[r], nodes, tag, k);
                x += cw + gutter;
            }
            y += need[r] + gutter;
        }

        List<object> edges = new List<object>();
        if (doc != null && doc.Has("edges")) edges = doc.List("edges");
        else edges = a.List("edges");
        KDrawEdges(p, edges, nodes);
    }

    // ---------------------------------------------------------------- slide -Action reuse

    /// <summary>Duplicate a page and swap text / pictures by shape name.  Nothing is guessed:
    /// every key that does not resolve is reported back in "missing".</summary>
    private static object KReuse(Bag a)
    {
        dynamic pres = RequirePres(a, true);
        int n = 0;
        try { n = (int)pres.Slides.Count; } catch (Exception) { }
        int from = a.Int("from", 0);
        if (from < 1)
            throw new Fail("BAD_ARGS", "slide -Action reuse needs -From <source slide number> (1.." + n + "); here -From is the SOURCE page, not a page range");
        if (from > n) throw new Fail("BAD_ARGS", "-From " + from + " is out of range (1.." + n + ")");

        Bag doc = KDoc(a);
        Dictionary<string, string> texts = new Dictionary<string, string>(StringComparer.Ordinal);
        Dictionary<string, string> images = new Dictionary<string, string>(StringComparer.Ordinal);
        if (doc != null)
        {
            Bag t = null, im = null;
            try { t = new Bag(doc.Raw("texts")); } catch (Exception) { }
            try { im = new Bag(doc.Raw("images")); } catch (Exception) { }
            texts = PairsOf(t);
            images = PairsOf(im);
        }
        if (texts.Count == 0 && images.Count == 0)
            throw new Fail("BAD_ARGS", "slide -Action reuse needs -Items {\"texts\":{\"<shape name>\":\"new text\"},\"images\":{\"<shape name>\":\"<new picture path>\"}}");

        dynamic src = pres.Slides.Item(from);
        // Duplicate() uses the clipboard and leaves the slide on it; keep the user's simple clipboard
        // content across it (and keep the slide format away from the next render's clipboard snapshot,
        // which used to kill the helper - 2026-11-03).
        dynamic dup = null;
        using (ClipboardKeeper.Keep())
        {
            dup = src.Duplicate();
            ClipboardKeeper.Sync();
        }
        int newIdx = from + 1;
        try { newIdx = Convert.ToInt32(dup.SlideIndex, CultureInfo.InvariantCulture); } catch (Exception) { }
        List<string> notes = new List<string>();
        int at = a.Int("at", 0);
        int requestedAt = 0;
        if (at > 0)
        {
            requestedAt = at > n + 1 ? n + 1 : at;
            if (requestedAt != newIdx)
            {
                // MoveTo on the object Duplicate() returned is silently ignored by WPS, so the
                // slide has to be fetched back out of the collection and moved that way.  The
                // position is then checked (not assumed): reporting a page that does not exist
                // is worse than reporting the one it really landed on.
                int shapeCount = 0;
                try { shapeCount = (int)dup.Shapes.Count; } catch (Exception) { }
                try { pres.Slides.Item(newIdx).MoveTo(requestedAt); } catch (Exception) { }
                bool landed = false;
                try { landed = (int)pres.Slides.Item(requestedAt).Shapes.Count == shapeCount; }
                catch (Exception) { landed = false; }
                if (landed) newIdx = requestedAt;
                else notes.Add("could not move the copy to position " + requestedAt + "; it stayed at " + newIdx);
            }
        }

        List<object> replacedText = new List<object>();
        List<string> missingText = new List<string>();
        List<object> replacedImg = new List<object>();
        List<string> missingImg = new List<string>();

        Dictionary<string, dynamic> byName = new Dictionary<string, dynamic>(StringComparer.Ordinal);
        int shapes = 0;
        try { shapes = (int)dup.Shapes.Count; } catch (Exception) { }
        for (int k = 1; k <= shapes; k++)
        {
            dynamic sh = null;
            try { sh = dup.Shapes.Item(k); } catch (Exception) { continue; }
            string nm = ShapeNameOf(sh);
            if (nm == null || nm.Length == 0) continue;
            if (!byName.ContainsKey(nm)) byName[nm] = sh;
        }

        int applied = 0;
        foreach (KeyValuePair<string, string> kv in texts)
        {
            if (!byName.ContainsKey(kv.Key)) { missingText.Add(kv.Key); continue; }
            dynamic sh = byName[kv.Key];
            string before = ShapeText(sh);
            bool ok = false;
            try
            {
                sh.TextFrame.TextRange.Text = kv.Value;
                ok = ShapeText(sh) == kv.Value;
            }
            catch (Exception) { ok = false; }
            if (ok) applied++;
            else missingText.Add(kv.Key);
            Dictionary<string, object> rec = new Dictionary<string, object>();
            rec["shape"] = kv.Key;
            rec["before"] = KClip(before, 40);
            rec["after"] = KClip(kv.Value, 40);
            rec["ok"] = ok;
            replacedText.Add(rec);
        }

        foreach (KeyValuePair<string, string> kv in images)
        {
            if (!byName.ContainsKey(kv.Key)) { missingImg.Add(kv.Key); continue; }
            if (kv.Value == null || !File.Exists(kv.Value))
            {
                missingImg.Add(kv.Key);
                Dictionary<string, object> bad = new Dictionary<string, object>();
                bad["shape"] = kv.Key;
                bad["ok"] = false;
                bad["detail"] = "picture not found: " + kv.Value;
                replacedImg.Add(bad);
                continue;
            }
            dynamic pic = byName[kv.Key];
            double cl = ToD(TryGet(() => (object)pic.Left, (object)0));
            double ct = ToD(TryGet(() => (object)pic.Top, (object)0));
            double cw = ToD(TryGet(() => (object)pic.Width, (object)0));
            double chh = ToD(TryGet(() => (object)pic.Height, (object)0));
            string insertPath = AsciiTempCopy(kv.Value);
            try
            {
                // insert the new picture, carry the geometry over, then drop the old one
                dynamic fresh = dup.Shapes.AddPicture(insertPath, 0, -1, (float)cl, (float)ct, (float)cw, (float)chh);
                try { fresh.Name = kv.Key; } catch (Exception) { }
                pic.Delete();
                byName[kv.Key] = fresh;
                Dictionary<string, object> rec = new Dictionary<string, object>();
                rec["shape"] = kv.Key;
                rec["path"] = kv.Value;
                rec["l"] = R(cl); rec["t"] = R(ct); rec["w"] = R(cw); rec["h"] = R(chh);
                rec["ok"] = true;
                replacedImg.Add(rec);
                applied++;
            }
            catch (Exception e)
            {
                Dictionary<string, object> bad = new Dictionary<string, object>();
                bad["shape"] = kv.Key;
                bad["ok"] = false;
                bad["detail"] = Clean(e.Message);
                replacedImg.Add(bad);
                missingImg.Add(kv.Key);
            }
            finally { TryDelete(insertPath); }
        }

        // A reused page is a copy, and this skill writes the page number as a PLAIN TEXT BOX (the WPS
        // slide-number placeholder reads "1" on every page on this machine) - so the copy would keep
        // the source's "1 / 1". Refresh it here; nothing else is renumbered.
        string pageNoAfter = null;
        if (byName.ContainsKey("pageNo"))
        {
            try
            {
                dynamic pn = byName["pageNo"];
                string want = newIdx.ToString(CultureInfo.InvariantCulture) + " / "
                    + (n + 1).ToString(CultureInfo.InvariantCulture);
                pn.TextFrame.TextRange.Text = want;
                pageNoAfter = ShapeText(pn);
            }
            catch (Exception) { }
        }

        Dictionary<string, object> miss = new Dictionary<string, object>();
        miss["texts"] = missingText;
        miss["images"] = missingImg;

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["reused"] = from;
        r["added"] = newIdx;
        if (requestedAt > 0) r["requestedAt"] = requestedAt;
        r["slides"] = n + 1;
        r["texts"] = replacedText;
        r["images"] = replacedImg;
        r["applied"] = applied;
        r["missing"] = miss;
        if (pageNoAfter != null) r["pageNo"] = pageNoAfter;
        r["note"] = "-From is the SOURCE slide here (read/render/export use it as a page range start);"
            + " -At is the insertion position, default right after the source";
        if (notes.Count > 0) r["warnings"] = notes;
        if (missingText.Count > 0 || missingImg.Count > 0)
            r["warning"] = missingText.Count + " text key(s) and " + missingImg.Count
                + " image key(s) did not resolve; `read -From " + newIdx + " -Count 1` shows the real shape names";
        return r;
    }

    /// <summary>Slide index of a slide object, falling back to what the caller thought it was.</summary>
    private static int SlideIndexOf(dynamic slide, int fallback)
    {
        try { return Convert.ToInt32(slide.SlideIndex, CultureInfo.InvariantCulture); }
        catch (Exception) { return fallback; }
    }
    /// <summary>String pairs out of a JSON object, keeping the order they were given in.</summary>
    private static Dictionary<string, string> PairsOf(Bag b)
    {
        Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.Ordinal);
        if (b == null) return d;
        foreach (KeyValuePair<string, object> kv in b.Items())
        {
            string v = kv.Value as string;
            if (v == null) v = Convert.ToString(kv.Value, CultureInfo.InvariantCulture);
            d[kv.Key] = v;
        }
        return d;
    }
}
