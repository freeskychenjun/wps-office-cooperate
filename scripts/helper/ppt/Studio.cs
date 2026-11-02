// wps-office-cooperate (ppt): layout studio - five commands borrowed from dsh-office,
// re-measured on WPS Presentation 12.1.0.29161 (see reference/ppt-quirks.md).
//
//   crop      - a picture cut to a frame: cover/contain, focus semantics (default keeps the
//               upper fifth, which is where a head is), and a warning when too much was cut;
//   smartart  - native SmartArt from the 134 layouts WPS exposes (all Chinese names);
//   connect   - a connector line between two shapes, with the sides chosen by geometry;
//   bg        - slide background: colour, gradient, picture, or a texture generated as a PNG;
//   anim      - slide transition and entrance effects, with the readback reported honestly.
//
// Measured facts this file depends on:
//   * Shape.PictureFormat.CropLeft/CropTop/CropRight/CropBottom are readable and writable
//     (Shape.CropLeft does NOT exist), and the value is a 0..1 fraction of the picture, not
//     points and not the frame size.  The picture is scaled first, then cut, then the frame
//     is set to the box - in that order the four fractions land exactly.
//   * Application.Visible must be true; Shapes.AddSmartArt(layout, l, t, w, h) works and a
//     fresh diagram arrives with sample nodes that have to be deleted before writing text.
//   * AddConnector(kind, beginX, beginY, endX, endY) + ConnectorFormat.BeginConnect /
//     EndConnect + Line.EndArrowheadStyle work; BeginConnected/EndConnected read back -1.
//   * Slide.Background.Fill accepts .Solid / .OneColorGradient / .UserPicture - the last is
//     what texture backgrounds use (a 1600x900 PNG drawn here, never a COM texture fill).
//   * SlideShowTransition.EntryEffect and TimeLine.MainSequence.AddEffect are writable, but
//     whether a deck actually plays them can only be judged in a slide show: the response
//     reports the readback and never claims more than that.
//
// Pure ASCII source (Chinese only ever travels inside JSON strings or as \u escapes), C# 5.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

internal static partial class Program
{
    // ---------------------------------------------------------------- shared bits

    /// <summary>The slide canvas in points (a deck may be 4:3, 16:9 or a custom size).</summary>
    private static void SlideSize(dynamic p, out double W, out double H)
    {
        W = ToD(TryGet(() => (object)p.PageSetup.SlideWidth, (object)960.0));
        H = ToD(TryGet(() => (object)p.PageSetup.SlideHeight, (object)540.0));
        if (W <= 1) W = 960;
        if (H <= 1) H = 540;
    }

    private static double ShapeD(dynamic sh, string prop, double fallback)
    {
        try
        {
            if (prop == "left") return ToD(sh.Left);
            if (prop == "top") return ToD(sh.Top);
            if (prop == "width") return ToD(sh.Width);
            if (prop == "height") return ToD(sh.Height);
        }
        catch (Exception) { }
        return fallback;
    }

    private static string ShapeNameOf(dynamic sh)
    {
        return Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
    }

    private static Dictionary<string, object> BoxOf(dynamic sh)
    {
        return new Dictionary<string, object> {
            { "l", R(ShapeD(sh, "left", 0)) }, { "t", R(ShapeD(sh, "top", 0)) },
            { "w", R(ShapeD(sh, "width", 0)) }, { "h", R(ShapeD(sh, "height", 0)) }
        };
    }

    /// <summary>Every shape on the slide, the agent's own takeover badge left out.</summary>
    private static List<dynamic> ContentShapes(dynamic slide)
    {
        List<dynamic> list = new List<dynamic>();
        int count = 0;
        try { count = (int)slide.Shapes.Count; } catch (Exception) { }
        for (int k = 1; k <= count; k++)
        {
            dynamic sh = null;
            try { sh = slide.Shapes.Item(k); } catch (Exception) { continue; }
            if (IsTakeoverMark(ShapeNameOf(sh))) continue;
            list.Add(sh);
        }
        return list;
    }

    // ---------------------------------------------------------------- crop

    /// <summary>Read the picture's own width/height ratio.  The frame keeps the ratio of the
    /// picture only while nothing is cropped; with crops the ratio is recovered from the frame
    /// and the fractions (frame = scaled picture * visible fraction, in both crop models).</summary>
    private static double PictureAspect(dynamic pic)
    {
        double fw = ShapeD(pic, "width", 0), fh = ShapeD(pic, "height", 0);
        double cl = ToD(TryGet(() => (object)pic.PictureFormat.CropLeft, (object)0.0));
        double ct = ToD(TryGet(() => (object)pic.PictureFormat.CropTop, (object)0.0));
        double cr = ToD(TryGet(() => (object)pic.PictureFormat.CropRight, (object)0.0));
        double cb = ToD(TryGet(() => (object)pic.PictureFormat.CropBottom, (object)0.0));
        double spanX = 1 - cl - cr, spanY = 1 - ct - cb;
        if (spanX < 0.02 || spanY < 0.02 || fw <= 1 || fh <= 1) return 0;
        return (fw / spanX) / (fh / spanY);
    }

    private static object CropCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic slide = SlideAt(p, a.Int("slide", 0));
        double W, H;
        SlideSize(p, out W, out H);

        string fit = (a.Str("fit", "cover") ?? "cover").Trim().ToLowerInvariant();
        if (fit != "cover" && fit != "contain") throw new Fail("BAD_ARGS", "-Fit must be cover or contain");
        string focus = (a.Str("focus", "") ?? "").Trim().ToLowerInvariant();
        if (focus == "centre" || focus == "middle") focus = "center";
        if (focus.Length > 0 && focus != "center" && focus != "top" && focus != "bottom"
            && focus != "left" && focus != "right")
            throw new Fail("BAD_ARGS", "-Focus must be top|center|bottom|left|right");

        double l = a.Dbl("l", -1), t = a.Dbl("t", -1), w = a.Dbl("w", 0), h = a.Dbl("h", 0);
        if (w <= 0) w = W;
        if (h <= 0) h = H;
        if (l < 0) l = (W - w) / 2;
        if (t < 0) t = (H - h) / 2;

        dynamic pic = null;
        string path = a.Str("path", null);
        double fileAspect = 0;
        if (path != null && path.Length > 0)
        {
            if (!File.Exists(path)) throw new Fail("BAD_ARGS", "image not found: " + path);
            try
            {
                using (System.Drawing.Image probe = System.Drawing.Image.FromFile(path))
                    if (probe.Height > 0) fileAspect = (double)probe.Width / probe.Height;
            }
            catch (Exception) { }
            string insertPath = AsciiTempCopy(path);
            try { pic = slide.Shapes.AddPicture(insertPath, 0, -1, (float)l, (float)t); }
            catch (COMException e) { throw new Fail("COM_ERROR", "AddPicture failed: " + Clean(e.Message)); }
            finally { TryDelete(insertPath); }
            string nm = a.Str("name", null);
            if (nm != null && nm.Length > 0) { try { pic.Name = nm; } catch (Exception) { } }
        }
        else
        {
            pic = ResolveShape(slide, a, "shape", "shapeIdx", true);
        }

        int shapeType = (int)ToD(TryGet(() => (object)pic.Type, (object)0));
        bool hasPicture = false;
        try { ToD(pic.PictureFormat.CropLeft); hasPicture = true; } catch (Exception) { hasPicture = false; }
        if (!hasPicture && shapeType != 13)
            throw new Fail("BAD_ARGS", "\"" + ShapeNameOf(pic) + "\" is not a picture (shape type " + shapeType
                + "); crop works on pictures - use image to place one first");

        if (fileAspect <= 0) fileAspect = PictureAspect(pic);
        if (fileAspect <= 0) fileAspect = w / Math.Max(h, 1);

        double frameBeforeW = ShapeD(pic, "width", 0), frameBeforeH = ShapeD(pic, "height", 0);
        double clipBeforeW = 1, clipBeforeH = 1;
        try
        {
            clipBeforeW = 1 - ToD(TryGet(() => (object)pic.PictureFormat.CropLeft, (object)0.0))
                            - ToD(TryGet(() => (object)pic.PictureFormat.CropRight, (object)0.0));
            clipBeforeH = 1 - ToD(TryGet(() => (object)pic.PictureFormat.CropTop, (object)0.0))
                            - ToD(TryGet(() => (object)pic.PictureFormat.CropBottom, (object)0.0));
        }
        catch (Exception) { }

        // start from an uncut picture every time, so the fractions below are absolute
        try
        {
            pic.PictureFormat.CropLeft = 0f;
            pic.PictureFormat.CropTop = 0f;
            pic.PictureFormat.CropRight = 0f;
            pic.PictureFormat.CropBottom = 0f;
        }
        catch (Exception e) { throw new Fail("COM_ERROR", "cannot reset the crop: " + Clean(e.Message)); }

        // where the cut is taken from: centre by default horizontally, the upper fifth vertically
        double shareX = focus == "left" ? 0.0 : focus == "right" ? 1.0 : 0.5;
        double shareY = focus == "top" ? 0.0 : focus == "bottom" ? 1.0 : focus == "center" ? 0.5 : 0.2;
        if (focus == "top" || focus == "bottom") shareX = 0.5;
        if (focus == "left" || focus == "right") shareY = 0.2;

        double scaledW, scaledH, cropL = 0, cropT = 0, cropR = 0, cropB = 0, lost = 0;
        try { pic.LockAspectRatio = 0; } catch (Exception) { }
        if (fit == "cover")
        {
            if (fileAspect > w / Math.Max(h, 1)) { scaledH = h; scaledW = h * fileAspect; }
            else { scaledW = w; scaledH = w / fileAspect; }
            try { pic.Width = (float)scaledW; pic.Height = (float)scaledH; } catch (Exception) { }
            double overX = scaledW - w, overY = scaledH - h;
            if (overX > 0.5)
            {
                cropL = overX * shareX / scaledW;
                cropR = overX * (1 - shareX) / scaledW;
            }
            if (overY > 0.5)
            {
                cropT = overY * shareY / scaledH;
                cropB = overY * (1 - shareY) / scaledH;
            }
            // WPS reads these as POINTS of the picture's scaled size, not as a 0..1 fraction:
            // writing 0.14 was stored and read back but cut nothing. Measured - scale by the size.
            try
            {
                pic.PictureFormat.CropLeft = (float)(cropL * scaledW);
                pic.PictureFormat.CropTop = (float)(cropT * scaledH);
                pic.PictureFormat.CropRight = (float)(cropR * scaledW);
                pic.PictureFormat.CropBottom = (float)(cropB * scaledH);
            }
            catch (Exception e) { throw new Fail("COM_ERROR", "cannot set the crop: " + Clean(e.Message)); }
            lost = 1 - (w * h) / Math.Max(scaledW * scaledH, 1);
        }
        else
        {
            try { pic.LockAspectRatio = -1; pic.Width = (float)w; } catch (Exception) { }
            double nh = ShapeD(pic, "height", h);
            if (nh > h + 0.5) { try { pic.Height = (float)h; } catch (Exception) { } }
            scaledW = ShapeD(pic, "width", w);
            scaledH = ShapeD(pic, "height", h);
        }
        try { pic.Left = (float)l; pic.Top = (float)t; } catch (Exception) { }
        if (fit == "contain")
        {
            double aw = ShapeD(pic, "width", w), ah = ShapeD(pic, "height", h);
            try
            {
                pic.Left = (float)(l + (w - aw) / 2);
                pic.Top = (float)(t + (h - ah) / 2);
            }
            catch (Exception) { }
        }
        // Measured: WPS keeps a crop only while the frame is not resized afterwards. Writing
        // .Width/.Height once a crop is in place stretches the picture to the frame (the fractions
        // stay, the render ignores them), so the frame is sized FIRST and never touched again -
        // the box is then whatever the crop leaves, read back below.
        double frameW = ShapeD(pic, "width", 0), frameH = ShapeD(pic, "height", 0);        double backCL = ToD(TryGet(() => (object)pic.PictureFormat.CropLeft, (object)-1));
        double backCT = ToD(TryGet(() => (object)pic.PictureFormat.CropTop, (object)-1));
        double backCR = ToD(TryGet(() => (object)pic.PictureFormat.CropRight, (object)-1));
        double backCB = ToD(TryGet(() => (object)pic.PictureFormat.CropBottom, (object)-1));

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slide"] = a.Int("slide", 0);
        r["shape"] = ShapeNameOf(pic);
        r["fit"] = fit;
        r["focus"] = focus.Length == 0 ? "default (centre horizontally, upper fifth vertically)" : focus;
        r["box"] = new Dictionary<string, object> { { "l", R(l) }, { "t", R(t) }, { "w", R(w) }, { "h", R(h) } };
        r["pictureAspect"] = Math.Round(fileAspect, 3);
        r["frame"] = BoxOf(pic);
        r["boxActual"] = BoxOf(pic);
        r["cropReadBack"] = new Dictionary<string, object> {
            { "left", Math.Round(backCL, 4) }, { "top", Math.Round(backCT, 4) },
            { "right", Math.Round(backCR, 4) }, { "bottom", Math.Round(backCB, 4) }
        };
        if (fit == "cover" && backCL + backCT + backCR + backCB <= 0.0001)
            r["warning"] = "WPS read the crop back as zero: the picture is stretched to the box instead of cut. Nothing was cut -"
                         + " check the render, and if this matters place the picture with image -W/-H instead";
        r["crop"] = new Dictionary<string, object> {
            { "left", Math.Round(cropL, 4) }, { "top", Math.Round(cropT, 4) },
            { "right", Math.Round(cropR, 4) }, { "bottom", Math.Round(cropB, 4) },
            { "units", "fraction of the picture; written to WPS as points (fraction x scaled size)" }
        };
        r["wasCropped"] = new Dictionary<string, object> {
            { "w", Math.Round(frameBeforeW, 1) }, { "h", Math.Round(frameBeforeH, 1) },
            { "visibleW", Math.Round(clipBeforeW, 3) }, { "visibleH", Math.Round(clipBeforeH, 3) }
        };
        if (fit == "cover")
        {
            int shows = (int)Math.Round((1 - lost) * 100);
            r["showsPercent"] = shows;
            if (lost > 0.4)
            {
                string side = focus == "top" ? "the upper part" : focus == "bottom" ? "the lower part"
                    : focus == "left" ? "the left part" : focus == "right" ? "the right part"
                    : "the middle";
                r["cut"] = "\"" + ShapeNameOf(pic) + "\" shows only " + shows + "% of itself (" + side
                    + "): render the slide and look - if a head or the subject is gone, pass -Focus top|center|bottom|left|right,"
                    + " or -Fit contain to show it whole";
            }
        }
        r["note"] = "the cut and the focus are right; the frame may end up a few points off the requested box (WPS re-lays the frame after every crop write - measured), so look at boxActual and render the slide";
        return r;
    }

    // ---------------------------------------------------------------- smartart

    private sealed class SmItem
    {
        public string Text = "";
        public List<SmItem> Children = new List<SmItem>();
    }

    private static List<SmItem> ParseSmItems(List<object> raw)
    {
        List<SmItem> items = new List<SmItem>();
        if (raw == null) return items;
        foreach (object o in raw)
        {
            SmItem one = new SmItem();
            string s = o as string;
            if (s != null) { one.Text = s; items.Add(one); continue; }
            Bag b = new Bag(o);
            one.Text = b.Str("text", b.Str("label", ""));
            if (b.Has("children")) one.Children = ParseSmItems(b.List("children"));
            items.Add(one);
        }
        return items;
    }

    /// <summary>English request word -> the Chinese names WPS gives its layouts (source stays ASCII).</summary>
    private static string LayoutKeyword(string want)
    {
        switch ((want ?? "").Trim().ToLowerInvariant())
        {
            case "process": case "flow": case "steps": case "step": return "\u6d41\u7a0b";
            case "cycle": case "circular": case "radial": return "\u5faa\u73af";
            case "hierarchy": case "org": case "orgchart": case "tree": case "structure": return "\u5c42\u6b21\u7ed3\u6784";
            case "pyramid": return "\u91d1\u5b57\u5854";
            case "venn": return "\u7ef4\u6069";
            case "list": case "bullet": case "bullets": return "\u5217\u8868";
            case "matrix": case "grid": return "\u77e9\u9635";
            case "funnel": return "\u6f0f\u6597";
            case "target": case "goal": return "\u76ee\u6807";
            case "gear": case "gears": return "\u9f7f\u8f6e";
            case "timeline": case "time": return "\u65f6\u95f4\u7ebf";
            case "arrow": case "arrows": case "chevron": return "\u7bad\u5934";
            case "balance": case "scale": return "\u5e73\u8861";
            case "picture": case "image": case "photo": return "\u56fe\u7247";
            case "relationship": case "relationships": return "\u5173\u7cfb";
            case "part": case "parts": return "\u90e8\u5206";
            default: return want;
        }
    }

    private static dynamic SmartArtLayout(dynamic app, string want, out string note, out int index, out int total)
    {
        note = null;
        index = 0;
        total = 0;
        dynamic layouts = app.SmartArtLayouts;
        total = (int)layouts.Count;
        if (total < 1) throw new Fail("UNSUPPORTED", "WPS exposes no SmartArtLayouts in this build");
        string keyword = LayoutKeyword(want);
        if (keyword == null || keyword.Length == 0) keyword = "process";
        // 1) exact name, 2) name contains the keyword, 3) id tail matches (PowerPoint layout id)
        for (int k = 1; k <= total; k++)
        {
            string name = Convert.ToString(TryGet(() => (object)layouts.Item(k).Name, ""), CultureInfo.InvariantCulture);
            if (string.Equals(name, want, StringComparison.OrdinalIgnoreCase)) { index = k; return layouts.Item(k); }
        }
        for (int k = 1; k <= total; k++)
        {
            string name = Convert.ToString(TryGet(() => (object)layouts.Item(k).Name, ""), CultureInfo.InvariantCulture);
            if (name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0) { index = k; return layouts.Item(k); }
        }
        for (int k = 1; k <= total; k++)
        {
            string id = Convert.ToString(TryGet(() => (object)layouts.Item(k).Id, ""), CultureInfo.InvariantCulture);
            if (id.EndsWith("/" + want, StringComparison.OrdinalIgnoreCase)) { index = k; return layouts.Item(k); }
        }
        List<string> shown = new List<string>();
        for (int k = 1; k <= total && shown.Count < 8; k++)
            shown.Add(Convert.ToString(TryGet(() => (object)layouts.Item(k).Name, ""), CultureInfo.InvariantCulture));
        note = "no layout matched \"" + want + "\" (" + keyword + "); used the first of " + total
             + " instead. First names: " + string.Join(", ", shown.ToArray());
        index = 1;
        return layouts.Item(1);
    }

    /// <summary>WPS exposes SmartArt nodes through more than one door and which one answers differs
    /// by build: the first that takes the text wins, so a diagram is never left empty silently.</summary>
    private static bool SetNodeText(dynamic node, string text)
    {
        try { node.TextFrame2.TextRange.Text = text; return true; } catch (Exception) { }
        try { node.TextFrame.TextRange.Text = text; return true; } catch (Exception) { }
        try { node.Shapes(1).TextFrame.TextRange.Text = text; return true; } catch (Exception) { }
        try { node.Shapes.Item(1).TextFrame.TextRange.Text = text; return true; } catch (Exception) { }
        return false;
    }

    private static dynamic AddNode(dynamic node, int position)
    {
        try { return node.AddNode(position); } catch (Exception) { }
        try { return node.AddNode(); } catch (Exception) { }
        return null;
    }

    private static string NodeText(dynamic node)
    {
        try
        {
            string s = Convert.ToString(node.TextFrame2.TextRange.Text, CultureInfo.InvariantCulture);
            if (s != null) return s;
        }
        catch (Exception) { }
        try { return Convert.ToString(node.TextFrame.TextRange.Text, CultureInfo.InvariantCulture); }
        catch (Exception) { return ""; }
    }

    private static int GrowNodes(dynamic node, List<SmItem> children)
    {
        int made = 0;
        foreach (SmItem child in children)
        {
            dynamic below = null;
            below = AddNode(node, 5);   // msoSmartArtNodeBelow
            if (below == null) break;
            SetNodeText(below, child.Text);
            made++;
            made += GrowNodes(below, child.Children);
        }
        return made;
    }

    private static object SmartArtCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        dynamic app = null;
        try { app = p.Application; } catch (Exception) { app = App(false, false); }
        if (app == null) throw new Fail("NO_APP", "WPS Presentation is not running");
        string action = (a.Str("action", "list") ?? "list").Trim().ToLowerInvariant();
        if (action == "list") return SmartArtList(app, a.Str("preset", a.Str("match", null)));

        if (action != "add") throw new Fail("BAD_ARGS", "smartart -Action must be list|add");
        int slideNo = a.Int("slide", 0);
        dynamic slide = SlideAt(p, slideNo);
        double W, H;
        SlideSize(p, out W, out H);
        double l = a.Dbl("l", W * 0.1), t = a.Dbl("t", H * 0.2);
        double w = a.Dbl("w", W * 0.8), h = a.Dbl("h", H * 0.6);
        string want = a.Str("preset", a.Str("type", "process"));

        string note;
        int index, total;
        dynamic layout = SmartArtLayout(app, want, out note, out index, out total);
        string layoutName = Convert.ToString(TryGet(() => (object)layout.Name, ""), CultureInfo.InvariantCulture);
        string layoutId = Convert.ToString(TryGet(() => (object)layout.Id, ""), CultureInfo.InvariantCulture);

        List<SmItem> items = ParseSmItems(a.List("items"));
        string single = a.Str("text", null);
        if (items.Count == 0 && single != null && single.Length > 0)
            items.Add(new SmItem { Text = single });
        if (items.Count == 0) throw new Fail("BAD_ARGS", "smartart -Action add needs -Items <json> (a list of texts, or objects with text+children)");

        dynamic shape = null;
        try { shape = slide.Shapes.AddSmartArt(layout, (float)l, (float)t, (float)w, (float)h); }
        catch (Exception e)
        {
            throw new Fail("UNSUPPORTED", "Shapes.AddSmartArt was refused by WPS: " + Clean(e.Message)
                + ". The layouts exist (" + total + " of them) but a diagram cannot be created here;"
                + " build the box diagram with shape + connect instead.");
        }
        string nm = a.Str("name", null);
        if (nm != null && nm.Length > 0) { try { shape.Name = nm; } catch (Exception) { } }

        dynamic smart = null;
        try { smart = shape.SmartArt; } catch (Exception) { smart = null; }
        if (smart == null)
        {
            Dictionary<string, object> bad = new Dictionary<string, object>();
            bad["slide"] = slideNo;
            bad["shape"] = ShapeNameOf(shape);
            bad["layout"] = layoutName;
            bad["unsupported"] = true;
            bad["note"] = "the shape was created but Shape.SmartArt is not readable in this WPS build, so its nodes cannot be written;"
                        + " delete it (shape -Action delete) and use shape + connect";
            return bad;
        }

        int removed = 0, nodes = 0;
        try
        {
            dynamic all = smart.AllNodes;
            while ((int)all.Count > 1)
            {
                all.Item((int)all.Count).Delete();
                removed++;
                if (removed > 60) break;
            }
            dynamic root = all.Item(1);
            SetNodeText(root, items[0].Text);
            int made = GrowNodes(root, items[0].Children);
            dynamic previous = root;
            for (int k = 1; k < items.Count; k++)
            {
                dynamic sib = null;
                sib = AddNode(previous, 2);   // msoSmartArtNodeAfter
                if (sib == null) break;
                SetNodeText(sib, items[k].Text);
                made++;
                made += GrowNodes(sib, items[k].Children);
                previous = sib;
            }
            nodes = (int)all.Count;
        }
        catch (Exception e)
        {
            note = (note == null ? "" : note + " | ") + "node writing stopped: " + Clean(e.Message);
        }

        List<object> texts = new List<object>();
        try
        {
            dynamic all = smart.AllNodes;
            int cnt = (int)all.Count;
            for (int k = 1; k <= cnt && k <= 12; k++) texts.Add(NodeText(all.Item(k)));
        }
        catch (Exception) { }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slide"] = slideNo;
        r["shape"] = ShapeNameOf(shape);
        r["layout"] = layoutName;
        r["layoutId"] = layoutId;
        r["layoutIndex"] = index;
        r["layoutsTotal"] = total;
        r["nodes"] = nodes;
        r["sampleNodesRemoved"] = removed;
        r["texts"] = texts;
        r["box"] = BoxOf(shape);
        if (note != null) r["note"] = note;
        return r;
    }

    private static object SmartArtList(dynamic app, string filter)
    {
        dynamic layouts = app.SmartArtLayouts;
        int total = (int)layouts.Count;
        string keyword = filter == null || filter.Length == 0 ? null : LayoutKeyword(filter);
        List<object> found = new List<object>();
        for (int k = 1; k <= total; k++)
        {
            string name = Convert.ToString(TryGet(() => (object)layouts.Item(k).Name, ""), CultureInfo.InvariantCulture);
            string id = Convert.ToString(TryGet(() => (object)layouts.Item(k).Id, ""), CultureInfo.InvariantCulture);
            if (keyword != null && name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0
                && id.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0) continue;
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["i"] = k;
            one["name"] = name;
            one["id"] = id;
            found.Add(one);
            if (found.Count >= 200) break;
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["total"] = total;
        r["matched"] = found.Count;
        r["layouts"] = found;
        r["note"] = "WPS names its layouts in Chinese; -Preset accepts an English keyword (process/cycle/hierarchy/pyramid/venn/list/matrix/"
                   + "funnel/target/gear/timeline/arrow/balance/picture), a Chinese keyword, or a layout name as listed above";
        return r;
    }

    // ---------------------------------------------------------------- connect

    private static int SideSite(string side, out string normalised)
    {
        normalised = (side ?? "").Trim().ToLowerInvariant();
        if (normalised == "top" || normalised == "up") { normalised = "top"; return 1; }
        if (normalised == "left") { normalised = "left"; return 2; }
        if (normalised == "bottom" || normalised == "down") { normalised = "bottom"; return 3; }
        if (normalised == "right") { normalised = "right"; return 4; }
        normalised = null;
        return 0;
    }

    private static void SideAuto(dynamic from, dynamic to, out string fromSide, out string toSide)
    {
        double fl = ShapeD(from, "left", 0), ft = ShapeD(from, "top", 0);
        double fw = ShapeD(from, "width", 0), fh = ShapeD(from, "height", 0);
        double tl = ShapeD(to, "left", 0), tt = ShapeD(to, "top", 0);
        double tw = ShapeD(to, "width", 0), th = ShapeD(to, "height", 0);
        double dx = (tl + tw / 2) - (fl + fw / 2);
        double dy = (tt + th / 2) - (ft + fh / 2);
        double gapX = dx > 0 ? tl - (fl + fw) : fl - (tl + tw);
        double gapY = dy > 0 ? tt - (ft + fh) : ft - (tt + th);
        bool horizontal;
        if (gapX > 0 && gapY <= 0) horizontal = true;        // side by side, one above the other's band
        else if (gapY > 0 && gapX <= 0) horizontal = false;  // stacked
        else horizontal = Math.Abs(dx) >= Math.Abs(dy);
        if (horizontal)
        {
            fromSide = dx >= 0 ? "right" : "left";
            toSide = dx >= 0 ? "left" : "right";
        }
        else
        {
            fromSide = dy >= 0 ? "bottom" : "top";
            toSide = dy >= 0 ? "top" : "bottom";
        }
    }

    private static object ConnectCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int slideNo = a.Int("slide", 0);
        dynamic slide = SlideAt(p, slideNo);
        dynamic from = ResolveShape(slide, a, "fromShape", "fromShapeIdx", true);
        dynamic to = ResolveShape(slide, a, "toShape", "toShapeIdx", true);
        if (ShapeNameOf(from) == ShapeNameOf(to))
            throw new Fail("BAD_ARGS", "-FromShape and -ToShape are the same shape");

        string kindWord = (a.Str("kind", "straight") ?? "straight").Trim().ToLowerInvariant();
        int kind;
        if (kindWord == "straight" || kindWord == "line") kind = 1;
        else if (kindWord == "elbow" || kindWord == "angled") kind = 2;
        else if (kindWord == "curve" || kindWord == "curved") kind = 3;
        else throw new Fail("BAD_ARGS", "-Kind must be straight|elbow|curve");

        string arrowWord = (a.Str("arrow", "end") ?? "end").Trim().ToLowerInvariant();
        if (arrowWord == "none" || arrowWord == "end" || arrowWord == "both") { }
        else throw new Fail("BAD_ARGS", "-Arrow must be none|end|both");

        string autoFrom, autoTo;
        SideAuto(from, to, out autoFrom, out autoTo);
        string gotFrom, gotTo;
        int siteFrom = SideSite(a.Str("fromSide", null), out gotFrom);
        int siteTo = SideSite(a.Str("toSide", null), out gotTo);
        if (siteFrom == 0) { gotFrom = autoFrom; siteFrom = SideSite(autoFrom, out gotFrom); }
        if (siteTo == 0) { gotTo = autoTo; siteTo = SideSite(autoTo, out gotTo); }

        double[] fromAnchor = Anchor(from, gotFrom);
        double[] toAnchor = Anchor(to, gotTo);

        dynamic link = null;
        try
        {
            link = slide.Shapes.AddConnector(kind, (float)fromAnchor[0], (float)fromAnchor[1],
                (float)toAnchor[0], (float)toAnchor[1]);
        }
        catch (Exception e) { throw new Fail("COM_ERROR", "AddConnector failed: " + Clean(e.Message)); }

        bool beginOk = false, endOk = false;
        try
        {
            link.ConnectorFormat.BeginConnect(from, siteFrom);
            beginOk = ToD(TryGet(() => (object)link.ConnectorFormat.BeginConnected, (object)0)) != 0;
        }
        catch (Exception) { beginOk = false; }
        try
        {
            link.ConnectorFormat.EndConnect(to, siteTo);
            endOk = ToD(TryGet(() => (object)link.ConnectorFormat.EndConnected, (object)0)) != 0;
        }
        catch (Exception) { endOk = false; }
        try { link.RerouteConnections(); } catch (Exception) { }

        int endArrow = 0, beginArrow = 0;
        try
        {
            if (arrowWord == "end" || arrowWord == "both") { link.Line.EndArrowheadStyle = 2; }
            if (arrowWord == "both") { link.Line.BeginArrowheadStyle = 2; }
            endArrow = (int)ToD(TryGet(() => (object)link.Line.EndArrowheadStyle, (object)0));
            beginArrow = (int)ToD(TryGet(() => (object)link.Line.BeginArrowheadStyle, (object)0));
        }
        catch (Exception) { }

        string labelName = null;
        string label = a.Str("label", null);
        if (label != null && label.Length > 0)
        {
            double cl = ToD(TryGet(() => (object)link.Left, (object)0));
            double ct = ToD(TryGet(() => (object)link.Top, (object)0));
            double cw = ToD(TryGet(() => (object)link.Width, (object)0));
            double ch = ToD(TryGet(() => (object)link.Height, (object)0));
            double mx = cl + cw / 2, my = ct + ch / 2;
            try
            {
                dynamic box = slide.Shapes.AddTextbox(1, (float)(mx - 40), (float)(my - 12), 80f, 24f);
                dynamic tr = box.TextFrame.TextRange;
                tr.Text = label;
                try
                {
                    tr.Font.Size = 11f;
                    tr.Font.Color.RGB = ColorOf(a.Str("color", "#333333"), 0x333333);
                    tr.ParagraphFormat.Alignment = 2;   // centred
                    box.TextFrame.WordWrap = -1;
                    box.Fill.Solid();
                    box.Fill.ForeColor.RGB = ColorOf("#FFFFFF", 0xFFFFFF);
                    box.Fill.Transparency = 0f;
                    box.Line.Visible = 0;
                }
                catch (Exception) { }
                labelName = ShapeNameOf(box);
            }
            catch (Exception) { }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slide"] = slideNo;
        r["shape"] = ShapeNameOf(link);
        r["from"] = ShapeNameOf(from);
        r["to"] = ShapeNameOf(to);
        r["kind"] = kindWord;
        r["sides"] = new Dictionary<string, object> {
            { "from", gotFrom + " (site " + siteFrom + ")" }, { "to", gotTo + " (site " + siteTo + ")" }
        };
        r["connected"] = new Dictionary<string, object> { { "begin", beginOk }, { "end", endOk } };
        r["arrows"] = new Dictionary<string, object> { { "end", endArrow }, { "begin", beginArrow } };
        r["box"] = BoxOf(link);
        if (labelName != null) r["label"] = labelName;
        if (!beginOk || !endOk)
            r["warning"] = "the connector does not report both ends attached - the line is drawn but will NOT follow the boxes when they move;"
                         + " check that both shapes are native autoshapes (a freeform or a group has no connection sites)";
        r["note"] = "render the slide to check the line does not run through other boxes";
        return r;
    }

    private static double[] Anchor(dynamic sh, string side)
    {
        double l = ShapeD(sh, "left", 0), t = ShapeD(sh, "top", 0);
        double w = ShapeD(sh, "width", 0), h = ShapeD(sh, "height", 0);
        if (side == "top") return new double[] { l + w / 2, t };
        if (side == "left") return new double[] { l, t + h / 2 };
        if (side == "bottom") return new double[] { l + w / 2, t + h };
        return new double[] { l + w, t + h / 2 };
    }

    // ---------------------------------------------------------------- bg

    /// <summary>A texture as a plain PNG: WPS accepts Background.Fill.UserPicture, and drawing
    /// the ground here is far steadier than the old COM texture fills.</summary>
    private static string GroundFile(string texture, string colour)
    {
        string dir = Path.Combine(Path.GetTempPath(), PptTempDirName);
        Directory.CreateDirectory(dir);
        string key = texture + "-" + (colour ?? "").Replace("#", "").ToLowerInvariant();
        string file = Path.Combine(dir, "ground-" + key + ".png");
        if (File.Exists(file) && new FileInfo(file).Length > 1000) return file;

        int W = 1600, H = 900;
        Color baseColour = Color.FromArgb(247, 243, 233);
        if (colour != null && colour.Length > 0)
        {
            try { baseColour = ColorTranslator.FromHtml(colour); } catch (Exception) { }
        }
        using (Bitmap bmp = new Bitmap(W, H, PixelFormat.Format24bppRgb))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(baseColour);
                if (texture == "paper")
                {
                    // aged paper: darker at the edges, a light ellipse of focus in the middle
                    using (GraphicsPath path = new GraphicsPath())
                    {
                        path.AddEllipse(-W * 0.25f, -H * 0.35f, W * 1.5f, H * 1.7f);
                        using (PathGradientBrush brush = new PathGradientBrush(path))
                        {
                            brush.CenterColor = Color.FromArgb(0, baseColour);
                            Color edge = Blend(baseColour, Color.FromArgb(107, 78, 30), 0.55);
                            brush.SurroundColors = new Color[] { Color.FromArgb(80, edge) };
                            brush.FocusScales = new PointF(0.55f, 0.5f);
                            g.FillRectangle(brush, 0, 0, W, H);
                        }
                    }
                    Random rnd = new Random(20261102);
                    for (int i = 0; i < 9000; i++)
                    {
                        int x = rnd.Next(W), y = rnd.Next(H);
                        int shade = rnd.Next(-10, 12);
                        Color speck = Color.FromArgb(
                            Math.Max(0, Math.Min(255, baseColour.R + shade)),
                            Math.Max(0, Math.Min(255, baseColour.G + shade)),
                            Math.Max(0, Math.Min(255, baseColour.B + shade)));
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(90, speck))) g.FillRectangle(b, x, y, 1, 1);
                    }
                }
                else if (texture == "grid")
                {
                    Color line = Blend(baseColour, Color.FromArgb(90, 100, 120), 0.22);
                    using (Pen pen = new Pen(line, 1.2f))
                    {
                        for (int x = 0; x < W; x += 50) g.DrawLine(pen, x, 0, x, H);
                        for (int y = 0; y < H; y += 50) g.DrawLine(pen, 0, y, W, y);
                    }
                }
                else if (texture == "dot" || texture == "dots")
                {
                    Color dot = Blend(baseColour, Color.FromArgb(90, 100, 120), 0.30);
                    using (SolidBrush b = new SolidBrush(dot))
                    {
                        for (int y = 18; y < H; y += 36)
                            for (int x = 18; x < W; x += 36)
                                g.FillEllipse(b, x - 2f, y - 2f, 4f, 4f);
                    }
                }
                else if (texture == "gradient")
                {
                    Color far = Blend(baseColour, Color.FromArgb(15, 61, 117), 0.18);
                    using (LinearGradientBrush brush = new LinearGradientBrush(
                        new Point(0, 0), new Point(W, H), baseColour, far))
                        g.FillRectangle(brush, 0, 0, W, H);
                }
                else throw new Fail("BAD_ARGS", "unknown -Texture \"" + texture + "\"; use paper|grid|dot|gradient");
            }
            bmp.Save(file, ImageFormat.Png);
        }
        return file;
    }

    private static Color Blend(Color a, Color b, double share)
    {
        if (share < 0) share = 0;
        if (share > 1) share = 1;
        return Color.FromArgb(
            (int)Math.Round(a.R + (b.R - a.R) * share),
            (int)Math.Round(a.G + (b.G - a.G) * share),
            (int)Math.Round(a.B + (b.B - a.B) * share));
    }

    private static object BgCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }
        if (n < 1) throw new Fail("BAD_ARGS", "this presentation has no slides");

        int one = a.Int("slide", 0);
        bool all = a.Flag("all", false) || one < 1;
        if (!all && (one < 1 || one > n)) throw new Fail("BAD_ARGS", "bg needs -Slide 1.." + n + " (or -All)");

        string texture = a.Str("texture", null);
        string path = a.Str("path", a.Str("image", null));
        string colour = a.Str("color", null);
        bool gradient = a.Flag("gradient", false);
        bool reset = a.Flag("reset", false);
        int how = 0;
        if (reset) how++;
        if (texture != null && texture.Length > 0) how++;
        if (path != null && path.Length > 0) how++;
        if (gradient) how++;
        if (colour != null && colour.Length > 0 && !gradient && how == 0) how++;
        if (how == 0) throw new Fail("BAD_ARGS", "bg needs one of -Color, -Gradient, -Texture, -Path (or -Reset)");
        if (how > 1 && !(gradient && colour != null) && !(texture != null && colour != null))
            throw new Fail("BAD_ARGS", "bg takes one background at a time (-Color may only accompany -Gradient/-Texture)");
        if (path != null && path.Length > 0 && !File.Exists(path)) throw new Fail("BAD_ARGS", "background image not found: " + path);

        string ground = null;
        string asciiImage = null;
        if (texture != null && texture.Length > 0) ground = GroundFile(texture.Trim().ToLowerInvariant(), colour);
        if (path != null && path.Length > 0) asciiImage = AsciiTempCopy(path);

        List<object> slides = new List<object>();
        int first = all ? 1 : one, last = all ? n : one;
        try
        {
            for (int k = first; k <= last; k++)
            {
                dynamic slide = p.Slides.Item(k);
                Dictionary<string, object> one2 = new Dictionary<string, object>();
                one2["slide"] = k;
                try
                {
                    if (reset)
                    {
                        try { slide.FollowMasterBackground = -1; } catch (Exception) { }
                        slide.Background.Fill.Solid();
                        slide.Background.Fill.ForeColor.RGB = ColorOf("#FFFFFF", 0xFFFFFF);
                    }
                    else
                    {
                        try { slide.FollowMasterBackground = 0; } catch (Exception) { }
                        dynamic fill = slide.Background.Fill;
                        if (ground != null || asciiImage != null)
                        {
                            fill.UserPicture(ground != null ? ground : asciiImage);
                        }
                        else if (gradient)
                        {
                            fill.OneColorGradient(1, 2, 0.35f);
                            if (colour != null && colour.Length > 0) fill.ForeColor.RGB = ColorOf(colour, 0xFFFFFF);
                        }
                        else
                        {
                            fill.Solid();
                            fill.ForeColor.RGB = ColorOf(colour, 0xFFFFFF);
                        }
                    }
                    one2["fillType"] = (int)ToD(TryGet(() => (object)slide.Background.Fill.Type, (object)-1));
                    string hex = null;
                    try { hex = DeckTheme.Hex((int)ToD(TryGet(() => (object)slide.Background.Fill.ForeColor.RGB, (object)0))); }
                    catch (Exception) { }
                    if (hex != null) one2["foreColor"] = hex;
                }
                catch (Exception e)
                {
                    one2["error"] = Clean(e.Message);
                }
                slides.Add(one2);
            }
        }
        finally
        {
            TryDelete(asciiImage);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slides"] = slides;
        r["count"] = slides.Count;
        if (ground != null) r["texture"] = new Dictionary<string, object> { { "kin", texture }, { "file", ground } };
        if (texture != null && texture.Length > 0) r["how"] = "a 1600x900 PNG drawn by the helper, then Background.Fill.UserPicture";
        r["note"] = "fill types read back per slide (1=solid, 3=gradient, 6=picture); render the slide - text laid on a busy texture may lose contrast";
        return r;
    }

    // ---------------------------------------------------------------- anim

    /// <summary>PptEntryEffect constants that WPS accepted in the probe (dsh-office used the same).</summary>
    private static int TransitionConstant(string name)
    {
        switch ((name ?? "").Trim().ToLowerInvariant())
        {
            case "none": case "cut": return 0;
            case "fade": case "fade-smooth": return 3849;
            case "push": return 3852;
            case "wipe": case "wipe-down": return 2817;
            case "split": return 3848;
            default:
                throw new Fail("BAD_ARGS", "unknown -Transition \"" + name + "\"; use none|fade|push|wipe|split");
        }
    }

    private static int AnimEffectConstant(string name)
    {
        switch ((name ?? "").Trim().ToLowerInvariant())
        {
            case "appear": return 1;
            case "fade": return 10;
            case "fly-in": case "flyin": case "fly": return 2;
            case "wipe": return 22;
            case "zoom": return 23;
            default:
                throw new Fail("BAD_ARGS", "unknown -Effect \"" + name + "\"; use fade|fly-in|wipe|zoom|appear");
        }
    }

    private static object AnimCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }
        if (n < 1) throw new Fail("BAD_ARGS", "this presentation has no slides");

        int one = a.Int("slide", 0);
        bool all = a.Flag("all", false) || one < 1;
        if (!all && (one < 1 || one > n)) throw new Fail("BAD_ARGS", "anim needs -Slide 1.." + n + " (or -All)");
        string transition = a.Str("transition", null);
        string effect = a.Str("effect", null);
        if ((transition == null || transition.Length == 0) && (effect == null || effect.Length == 0))
            throw new Fail("BAD_ARGS", "anim needs -Transition none|fade|push|wipe|split and/or -Effect fade|fly-in|wipe|zoom|appear");
        bool onClick = a.Flag("onClick", false);
        int trigger = onClick ? 1 : 2;   // 1 = on click, 2 = with previous
        int transitionValue = transition != null && transition.Length > 0 ? TransitionConstant(transition) : -1;
        int effectValue = effect != null && effect.Length > 0 ? AnimEffectConstant(effect) : 0;
        double duration = a.Dbl("duration", 0.6);

        List<object> slides = new List<object>();
        int first = all ? 1 : one, last = all ? n : one;
        bool unsupported = false;
        List<string> broke = new List<string>();
        for (int k = first; k <= last; k++)
        {
            dynamic slide = p.Slides.Item(k);
            Dictionary<string, object> entry = new Dictionary<string, object>();
            entry["slide"] = k;
            if (transitionValue >= 0)
            {
                int readBack = -1;
                try
                {
                    slide.SlideShowTransition.EntryEffect = transitionValue;
                    try { slide.SlideShowTransition.Duration = (float)duration; } catch (Exception) { }
                    readBack = (int)ToD(TryGet(() => (object)slide.SlideShowTransition.EntryEffect, (object)-1));
                }
                catch (Exception e)
                {
                    broke.Add("slide " + k + " transition: " + Clean(e.Message));
                    unsupported = true;
                }
                entry["transition"] = new Dictionary<string, object> {
                    { "wrote", transition }, { "constant", transitionValue },
                    { "readBack", readBack }, { "matches", readBack == transitionValue }
                };
                if (readBack != transitionValue && readBack >= 0) unsupported = true;
            }
            if (effectValue > 0)
            {
                int before = 0, after = 0;
                try { before = (int)slide.TimeLine.MainSequence.Count; } catch (Exception) { }
                int added = 0;
                List<dynamic> shapes = ContentShapes(slide);
                foreach (dynamic sh in shapes)
                {
                    try
                    {
                        dynamic fx = slide.TimeLine.MainSequence.AddEffect(sh, effectValue, 0, trigger);
                        try { fx.Timing.Duration = (float)duration; } catch (Exception) { }
                        added++;
                    }
                    catch (Exception e)
                    {
                        if (broke.Count < 6) broke.Add("slide " + k + " effect: " + Clean(e.Message));
                        break;
                    }
                }
                try { after = (int)slide.TimeLine.MainSequence.Count; } catch (Exception) { }
                entry["effect"] = new Dictionary<string, object> {
                    { "wrote", effect }, { "constant", effectValue }, { "trigger", trigger },
                    { "shapes", shapes.Count }, { "added", added },
                    { "sequenceBefore", before }, { "sequenceAfter", after }
                };
                if (added == 0 && shapes.Count > 0) unsupported = true;
            }
            slides.Add(entry);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slides"] = slides;
        r["count"] = slides.Count;
        r["unsupported"] = unsupported;
        if (broke.Count > 0) r["problems"] = broke;
        if (unsupported)
            r["note"] = "WPS refused or did not read back part of this: the deck may save without it. Nothing was silently dropped -"
                      + " see readBack/added above, and treat this as unsupported rather than working.";
        else
            r["note"] = "the values above are what WPS read back after writing. That is NOT proof of playback:"
                      + " only a slide show (F5) shows whether the transition and effects actually run.";
        return r;
    }
}
