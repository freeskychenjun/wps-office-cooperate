// wps-office-cooperate (ppt): `master` - 版式、页脚与"套外来模板".
//
// The measured truth about WPS 12.1.0.29161 (each line cost a probe, none is guessed):
//   * Slide.CustomLayout = <layout object> does NOTHING (the read-back keeps the old
//     name, no error).  The only working way to change a slide's layout is the LEGACY
//     constant: slide.Layout = 1..11 (1 -> "标题幻灯片", 2 -> "标题和文本", ...).
//   * Presentation.ApplyTemplate(path) returns success and changes nothing (designs
//     count, slide size and layouts are identical afterwards) - a silent no-op, so it
//     is not used here.
//   * Copy/Paste does NOT carry the source design, which makes it useless for "import
//     their template into my deck"... BUT it is exactly right in the other direction:
//     pasting MY content pages into THEIR template deck makes the content adopt their
//     master and layouts.  That is `rebase`, and it is how 套模板 is implemented.
//   * Per-slide HeadersFooters (footer text, page number, date) write and read back
//     reliably; the master's own HeadersFooters stays empty, so a batch has to walk the
//     slides.  A footer that is not Visible ignores its Text, so visibility is set first.
//
// Pure ASCII source (Chinese only ever travels inside JSON strings), C# 5.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

internal static partial class Program
{
    private static readonly string[] LayoutConstantNames = new string[]
    {
        "", "title slide", "title and text", "section header", "two content", "comparison",
        "title only", "blank", "content with caption", "picture with caption",
        "title and vertical text", "vertical title and text"
    };

    private static object MasterCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }
        string action = (a.Str("action", "list") ?? "list").ToLowerInvariant();

        if (action == "list") return MasterList(p, n);
        if (action == "footer") return MasterFooter(p, n, a);
        if (action == "layout") return MasterLayout(p, n, a);
        if (action == "rebase") return MasterRebase(p, n, a);
        throw new Fail("BAD_ARGS", "master -Action must be list|footer|layout|rebase");
    }

    // ---------------------------------------------------------------- list

    private static object MasterList(dynamic p, int n)
    {
        Dictionary<string, object> r = new Dictionary<string, object>();

        List<object> designs = new List<object>();
        int designCount = 0;
        try { designCount = (int)p.Designs.Count; } catch (Exception) { }
        for (int i = 1; i <= designCount; i++)
        {
            dynamic d = null;
            try { d = p.Designs.Item(i); } catch (Exception) { continue; }
            designs.Add(new Dictionary<string, object> {
                { "index", i },
                { "name", Convert.ToString(TryGet(() => (object)d.Name, ""), CultureInfo.InvariantCulture) },
                { "layouts", Convert.ToInt32(TryGet(() => (object)d.SlideMaster.CustomLayouts.Count, (object)0), CultureInfo.InvariantCulture) }
            });
        }
        r["designs"] = designs;

        List<object> layouts = new List<object>();
        int layoutCount = 0;
        try { layoutCount = (int)p.SlideMaster.CustomLayouts.Count; } catch (Exception) { }
        for (int i = 1; i <= layoutCount; i++)
        {
            dynamic l = null;
            try { l = p.SlideMaster.CustomLayouts.Item(i); } catch (Exception) { continue; }
            layouts.Add(new Dictionary<string, object> {
                { "index", i },
                { "name", Convert.ToString(TryGet(() => (object)l.Name, ""), CultureInfo.InvariantCulture) }
            });
        }
        r["layouts"] = layouts;

        List<object> slides = new List<object>();
        for (int k = 1; k <= n; k++)
        {
            dynamic s = null;
            try { s = p.Slides.Item(k); } catch (Exception) { continue; }
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["slide"] = k;
            one["layout"] = Convert.ToString(TryGet(() => (object)s.CustomLayout.Name, ""), CultureInfo.InvariantCulture);
            string footer = "";
            bool footerOn = false, numberOn = false, dateOn = false;
            try
            {
                dynamic hf = s.HeadersFooters;
                footer = Convert.ToString(TryGet(() => (object)hf.Footer.Text, ""), CultureInfo.InvariantCulture);
                footerOn = Convert.ToInt32(TryGet(() => (object)hf.Footer.Visible, (object)0), CultureInfo.InvariantCulture) != 0;
                numberOn = Convert.ToInt32(TryGet(() => (object)hf.SlideNumber.Visible, (object)0), CultureInfo.InvariantCulture) != 0;
                dateOn = Convert.ToInt32(TryGet(() => (object)hf.DateAndTime.Visible, (object)0), CultureInfo.InvariantCulture) != 0;
            }
            catch (Exception) { }
            one["footer"] = footer;
            one["footerVisible"] = footerOn;
            one["pageNumber"] = numberOn;
            one["date"] = dateOn;
            slides.Add(one);
        }
        r["slides"] = slides;
        r["size"] = new Dictionary<string, object> {
            { "w", R(ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)0.0))) },
            { "h", R(ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)0.0))) }
        };
        r["note"] = "changing a layout works only through -Layout <1..11> (the legacy constants); assigning "
                  + "Slide.CustomLayout does nothing in WPS. Try it on one slide and render before doing the whole deck.";
        return r;
    }

    private static double ToDouble(object value)
    {
        try { return Convert.ToDouble(value, CultureInfo.InvariantCulture); } catch (Exception) { return 0; }
    }

    // ---------------------------------------------------------------- footer / page numbers

    private static object MasterFooter(dynamic p, int n, Bag a)
    {
        HashSet<int> only = ParseSlideSet(a.Str("slides", null), n);
        bool skipCover = a.Flag("skipCover", false);
        bool clear = a.Flag("clear", false);
        string text = a.Str("text", null);
        bool wantNumber = a.Flag("pageNum", false);
        bool wantDate = a.Flag("date", false);
        if (!clear && text == null && !wantNumber && !wantDate)
            throw new Fail("BAD_ARGS", "master -Action footer needs -Text <footer> and/or -PageNum/-Date (or -Clear)");

        List<object> changed = new List<object>();
        List<object> skipped = new List<object>();
        int footers = 0, numbers = 0, dates = 0;

        for (int k = 1; k <= n; k++)
        {
            if (only != null && !only.Contains(k)) continue;
            if (skipCover && k == 1) { skipped.Add(new Dictionary<string, object> { { "slide", k }, { "reason", "cover" } }); continue; }
            dynamic s = null;
            try { s = p.Slides.Item(k); } catch (Exception) { continue; }
            Dictionary<string, object> rec = new Dictionary<string, object>();
            rec["slide"] = k;
            Dictionary<string, object> before = new Dictionary<string, object>();
            try
            {
                dynamic hf = s.HeadersFooters;
                before["footer"] = Convert.ToString(TryGet(() => (object)hf.Footer.Text, ""), CultureInfo.InvariantCulture);
                before["pageNumber"] = Convert.ToInt32(TryGet(() => (object)hf.SlideNumber.Visible, (object)0), CultureInfo.InvariantCulture) != 0;
                rec["before"] = before;
            }
            catch (Exception) { }
            bool touched = false;
            if (clear || text != null)
            {
                // Visibility first: a hidden footer ignores whatever Text you give it.
                try
                {
                    dynamic hf = s.HeadersFooters;
                    if (clear)
                    {
                        hf.Footer.Visible = 0;
                        hf.Footer.Text = "";
                    }
                    else
                    {
                        hf.Footer.Visible = -1;
                        hf.Footer.Text = text;
                    }
                    touched = true;
                    footers++;
                }
                catch (Exception e) { rec["footerError"] = Clean(e.Message); }
            }
            if (wantNumber || clear)
            {
                try
                {
                    s.HeadersFooters.SlideNumber.Visible = clear ? 0 : -1;
                    touched = true;
                    if (!clear) numbers++;
                }
                catch (Exception e) { rec["numberError"] = Clean(e.Message); }
            }
            if (wantDate || clear)
            {
                try
                {
                    s.HeadersFooters.DateAndTime.Visible = clear ? 0 : -1;
                    touched = true;
                    if (!clear) dates++;
                }
                catch (Exception e) { rec["dateError"] = Clean(e.Message); }
            }
            if (!touched) continue;
            try
            {
                dynamic hf = s.HeadersFooters;
                Dictionary<string, object> after = new Dictionary<string, object>();
                after["footer"] = Convert.ToString(TryGet(() => (object)hf.Footer.Text, ""), CultureInfo.InvariantCulture);
                after["pageNumber"] = Convert.ToInt32(TryGet(() => (object)hf.SlideNumber.Visible, (object)0), CultureInfo.InvariantCulture) != 0;
                after["date"] = Convert.ToInt32(TryGet(() => (object)hf.DateAndTime.Visible, (object)0), CultureInfo.InvariantCulture) != 0;
                rec["after"] = after;
            }
            catch (Exception) { }
            changed.Add(rec);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["slides"] = n;
        r["changed"] = changed.Count;
        r["footers"] = footers;
        r["pageNumbers"] = numbers;
        r["dates"] = dates;
        if (changed.Count > 0) r["details"] = changed;
        if (skipped.Count > 0) r["skipped"] = skipped;
        r["note"] = "footers and page numbers are per slide and only show once Visible is on; the master's own "
                  + "HeadersFooters stays empty, so this walks every page. Run render on one page before saving.";
        return r;
    }

    // ---------------------------------------------------------------- layout

    private static object MasterLayout(dynamic p, int n, Bag a)
    {
        HashSet<int> only = ParseSlideSet(a.Str("slides", null), n);
        int constant = a.Int("layout", 0);
        if (constant < 1 || constant > 11)
            throw new Fail("BAD_ARGS", "master -Action layout needs -Layout <1..11> (the legacy layout constants: 1 title slide, "
                + "2 title and text, 3 section header, 4 two content, 5 comparison, 6 title only, 7 blank, 8 content with "
                + "caption, 9 picture with caption, 10 title and vertical text, 11 vertical title and text)");
        bool dryRun = a.Flag("dryRun", false);

        List<object> changed = new List<object>();
        List<object> warnings = new List<object>();
        for (int k = 1; k <= n; k++)
        {
            if (only != null && !only.Contains(k)) continue;
            dynamic s = null;
            try { s = p.Slides.Item(k); } catch (Exception) { continue; }
            string before = Convert.ToString(TryGet(() => (object)s.CustomLayout.Name, ""), CultureInfo.InvariantCulture);
            int shapesBefore = Convert.ToInt32(TryGet(() => (object)s.Shapes.Count, (object)0), CultureInfo.InvariantCulture);
            Dictionary<string, object> rec = new Dictionary<string, object>();
            rec["slide"] = k;
            rec["layoutBefore"] = before;
            rec["shapesBefore"] = shapesBefore;
            if (dryRun) { rec["would"] = constant; changed.Add(rec); continue; }
            try
            {
                // The only form that works in WPS: the legacy constant, not CustomLayout.
                s.Layout = constant;
                string after = Convert.ToString(TryGet(() => (object)s.CustomLayout.Name, ""), CultureInfo.InvariantCulture);
                int shapesAfter = Convert.ToInt32(TryGet(() => (object)s.Shapes.Count, (object)0), CultureInfo.InvariantCulture);
                rec["layoutAfter"] = after;
                rec["shapesAfter"] = shapesAfter;
                if (after == before && constant != 0)
                    rec["note"] = "layout did not change (already that layout, or WPS refused it)";
                if (shapesAfter < shapesBefore)
                    warnings.Add(new Dictionary<string, object> {
                        { "slide", k }, { "detail", "shapes went from " + shapesBefore + " to " + shapesAfter
                            + ": placing content on a layout can remove placeholders that held it" } });
                changed.Add(rec);
            }
            catch (Exception e)
            {
                rec["error"] = Clean(e.Message);
                changed.Add(rec);
            }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["layoutConstant"] = constant;
        r["changed"] = changed.Count;
        if (changed.Count > 0) r["details"] = changed;
        if (warnings.Count > 0) r["warnings"] = warnings;
        r["dryRun"] = dryRun;
        r["use"] = "compare layoutBefore/layoutAfter and shapesBefore/shapesAfter, then render the changed pages: "
                 + "a layout swap re-flows the placeholders it owns.";
        return r;
    }

    // ---------------------------------------------------------------- rebase (套模板)

    private static object MasterRebase(dynamic p, int n, Bag a)
    {
        string template = a.Str("path", null);
        if (template == null || template.Length == 0) throw new Fail("BAD_ARGS", "master -Action rebase needs -Path <template deck>");
        template = FullPathOf(template);
        if (!File.Exists(template)) throw new Fail("BAD_ARGS", "template not found: " + template);
        string outArg = a.Str("out", null);
        if (outArg == null || outArg.Length == 0) throw new Fail("BAD_ARGS", "master -Action rebase needs -Out <new deck>");
        string target = FullPathOf(outArg);
        HashSet<int> only = ParseSlideSet(a.Str("slides", null), n);
        int layoutConstant = a.Int("layout", 0);
        bool keepTemplatePages = a.Flag("keepTemplatePages", false);

        string self = Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture);
        if (string.Equals(template, self, StringComparison.OrdinalIgnoreCase))
            throw new Fail("BAD_ARGS", "the template is the deck itself");

        dynamic app = App(true, false);
        dynamic tpl = null;
        try
        {
            tpl = app.Presentations.Open(template, 0, 0, 0);
            int templatePages = 0;
            try { templatePages = (int)tpl.Slides.Count; } catch (Exception) { }
            double srcW = ToDouble(TryGet(() => (object)p.PageSetup.SlideWidth, (object)0.0));
            double srcH = ToDouble(TryGet(() => (object)p.PageSetup.SlideHeight, (object)0.0));
            double tplW = ToDouble(TryGet(() => (object)tpl.PageSetup.SlideWidth, (object)0.0));
            double tplH = ToDouble(TryGet(() => (object)tpl.PageSetup.SlideHeight, (object)0.0));

            int pasted = 0;
            for (int k = 1; k <= n; k++)
            {
                if (only != null && !only.Contains(k)) continue;
                p.Slides.Item(k).Copy();
                Thread.Sleep(80);
                try { tpl.Slides.Paste(); }
                catch (Exception)
                {
                    Thread.Sleep(350);
                    tpl.Slides.Paste();
                }
                pasted++;
            }
            if (layoutConstant >= 1 && layoutConstant <= 11)
            {
                for (int k = templatePages + 1; k <= templatePages + pasted; k++)
                {
                    try { tpl.Slides.Item(k).Layout = layoutConstant; } catch (Exception) { }
                }
            }
            if (!keepTemplatePages)
            {
                // the template's own sample pages are scaffolding, not content
                for (int k = templatePages; k >= 1; k--) tpl.Slides.Item(k).Delete();
            }
            int finalPages = 0;
            try { finalPages = (int)tpl.Slides.Count; } catch (Exception) { }
            tpl.SaveAs(target, 24);

            Dictionary<string, object> r = new Dictionary<string, object>();
            r["source"] = self;
            r["template"] = template;
            r["out"] = target;
            r["pasted"] = pasted;
            r["templatePagesRemoved"] = keepTemplatePages ? 0 : templatePages;
            r["pages"] = finalPages;
            r["sizeBefore"] = new Dictionary<string, object> { { "w", R(srcW) }, { "h", R(srcH) } };
            r["sizeAfter"] = new Dictionary<string, object> { { "w", R(tplW) }, { "h", R(tplH) } };
            if (Math.Abs(srcW - tplW) > 1 || Math.Abs(srcH - tplH) > 1)
                r["warning"] = "the template is " + R(tplW) + "x" + R(tplH) + " while the source is " + R(srcW) + "x" + R(srcH)
                             + ": shapes keep their absolute position, so content designed for a bigger page can overflow. "
                             + "Render the result before saving anything else.";
            r["note"] = "the pages were pasted into the template, so they adopt its master and layouts (pasting into a blank "
                      + "deck does the opposite and would throw the template away). The source deck is untouched; "
                      + "the new file is written by WPS and therefore encrypted on this machine.";
            return r;
        }
        catch (Exception e)
        {
            throw new Fail("COM_ERROR", "rebase failed: " + Clean(e.Message));
        }
        finally
        {
            if (tpl != null) { try { tpl.Close(); } catch (Exception) { } }
        }
    }
}
