// fillgaps - fill the empty band at the foot of a page with body text: the band a figure or a
// table leaves behind when it does not fit and jumps whole to the next page.
//
// Ported from dsh-office (helper\OfficeHelper.cs: FillGaps + Block / Absorb / Widen / IsCaption)
// and re-measured on WPS 12.1.0.29161, where every primitive it needs works:
//   Range.Information(3) = page, (6) = vertical position in points, (12) = inside a table,
//   Paragraph.OutlineLevel, Section.PageSetup.PageHeight / BottomMargin / TopMargin,
//   Document.Repaginate(), Document.ComputeStatistics(2) = pages,
//   Range.FormattedText (assignable - that is how a paragraph is carried around).
//
//   {"doc":"...","dryRun":false,"maxMoves":8,"minGapPt":0}
//
// The rule (deliberately conservative):
//   * the paragraph just before the block is body text (OutlineLevel == 10),
//   * the block starts on a LATER page than that paragraph,
//   * the empty band below that paragraph is worth at least max(90, 12% of the text height) pt,
//   * the paragraph just after the block is body text, not empty, holds no picture,
//     is not a caption, and does not sit inside a table.
// Then that paragraph is copied to the top of the band and the original is deleted, so the
// figure / table keeps its place and only the text moves up. A paragraph that ends with ":" or
// the full-width colon announces what follows it, so the next one comes along. Text never jumps
// over a heading: a non-body paragraph anywhere in the chain stops the move.

using System;
using System.Collections.Generic;
using System.Globalization;

internal static partial class Program
{
    private const int FG_PAGE = 3;           // wdActiveEndPageNumber
    private const int FG_VPOS = 6;           // wdVerticalPositionRelativeToPage (points)
    private const int FG_INTABLE = 12;       // wdWithInTable
    private const int FG_BODY = 10;          // wdOutlineLevelBodyText

    /// <summary>
    /// fillgaps - see the header. dryRun reports and touches nothing (so it also runs on a
    /// read-only document); every other mode changes the open document and saves nothing.
    /// </summary>
    private static object FillGapsDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        bool dryRun = a.Flag("dryRun", false);
        if (!dryRun && IsReadOnly(doc))
            throw new Fail("READONLY", "The document is open read-only; fillgaps would move paragraphs in it. Use -DryRun to see what it would move.");

        int maxMoves = a.Int("maxMoves", 8);
        if (maxMoves < 1) maxMoves = 1;
        if (maxMoves > 40) maxMoves = 40;
        int minGapGiven = a.Int("minGapPt", 0);

        try { doc.Repaginate(); } catch (Exception) { }

        double[] geo0 = GapGeometry(doc, null);
        double threshold = minGapGiven > 0 ? (double)minGapGiven : Math.Max(90.0, (geo0[0] - geo0[1]) * 0.12);

        int figures = (int)TryGet(() => (object)(int)doc.InlineShapes.Count, 0);
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        List<int> all = new List<int>();
        for (int i = 1; i <= figures; i++) all.Add(i);
        for (int i = 1; i <= tables; i++) all.Add(-i);

        int moved = 0;
        int blocks = 0;
        List<object> moves = new List<object>();
        List<string> notes = new List<string>();
        if (all.Count == 0)
            notes.Add("The document has no inline picture and no table, so there is nothing fillgaps could move text in front of.");

        foreach (int index in all)
        {
            int here = 0;
            bool more = false;
            for (int step = 0; step < maxMoves; step++)
            {
                int start, end;
                if (!GapBlock(doc, index, out start, out end)) break;
                int docEnd = (int)TryGet(() => (object)(int)doc.Content.End, 0);
                if (docEnd <= 2 || start <= 0 || end >= docEnd - 1) break;

                Tr("fillgaps block " + index + " start " + start + " end " + end + " step " + step);

                dynamic before = null;
                try { before = doc.Range(start - 1, start - 1).Paragraphs.Item(1); } catch (Exception) { break; }

                double gap = 0;
                int gapPage = 0;
                int gapAtPara = 0;
                if (!more)
                {
                    int level = (int)TryGet(() => (object)Convert.ToInt32(before.OutlineLevel, CultureInfo.InvariantCulture), -1);
                    if (level != FG_BODY) break;
                    int pageTip = (int)TryGet(() => (object)Convert.ToInt32(doc.Range(start - 1, start - 1).Information(FG_PAGE), CultureInfo.InvariantCulture), 0);
                    int pageHead = (int)TryGet(() => (object)Convert.ToInt32(doc.Range(start, start).Information(FG_PAGE), CultureInfo.InvariantCulture), 0);
                    if (pageHead <= pageTip) break;
                    double[] geo = GapGeometry(doc, before);
                    double vpos = Num(TryGet(() => (object)Convert.ToDouble(doc.Range(start - 1, start - 1).Information(FG_VPOS), CultureInfo.InvariantCulture), null), -1);
                    if (vpos < 0) break;
                    gap = geo[0] - vpos - 20;
                    gapPage = pageTip;
                    gapAtPara = ParaIndexAtOrAfter(doc, start - 1);
                    Tr("fillgaps block " + index + " pages " + pageTip + "/" + pageHead + " gap " + gap + " threshold " + threshold);
                    if (gap < threshold) break;
                }

                // Empty paragraphs around the block are its spacing; a caption belongs to it.
                GapAbsorb(doc, ref start, ref end);
                if (end >= docEnd - 1) break;

                dynamic after = null;
                try { after = doc.Range(end, end).Paragraphs.Item(1); } catch (Exception) { break; }
                string text = CleanPara(Convert.ToString(TryGet(() => (object)after.Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
                if (text.Length == 0) break;
                if ((int)TryGet(() => (object)Convert.ToInt32(after.OutlineLevel, CultureInfo.InvariantCulture), -1) != FG_BODY) break;
                if ((int)TryGet(() => (object)(int)after.Range.InlineShapes.Count, 0) > 0) break;
                if (GapIsCaption(doc, after)) break;
                if (GapInTable(after.Range)) break;

                Dictionary<string, object> rec = new Dictionary<string, object>();
                rec["block"] = index;
                rec["kind"] = index > 0 ? "figure" : "table";
                rec["para"] = ParaIndexAtOrAfter(doc, start);
                rec["gapAtPara"] = gapAtPara;
                rec["page"] = gapPage;
                rec["gap"] = Math.Round(gap, 1);
                rec["text"] = Head(text, 80);

                if (dryRun)
                {
                    // Nothing is written: one report line per block is what a dry run is for.
                    rec["wouldMove"] = true;
                    moves.Add(rec);
                    here++;
                    break;
                }

                // Carry the paragraph to the top of the band, then find the block again (it moved
                // down by exactly what was put in front of it) and delete the paragraph it left.
                try
                {
                    if (!GapCopyAbove(doc, index, start, after))
                    {
                        notes.Add("block " + index + ": could not find a place in front of it for the paragraph");
                        break;
                    }
                }
                catch (Exception e) { notes.Add("block " + index + ": could not copy the paragraph up: " + Clean(e.Message)); break; }
                if (!GapBlock(doc, index, out start, out end)) break;
                GapAbsorb(doc, ref start, ref end);
                if (end >= (int)TryGet(() => (object)(int)doc.Content.End, 0) - 1) break;
                try { doc.Range(end, end).Paragraphs.Item(1).Range.Delete(); }
                catch (Exception e) { notes.Add("block " + index + ": the paragraph was copied up but not removed below it: " + Clean(e.Message)); }

                moved++;
                here++;
                moves.Add(rec);
                more = text.EndsWith(":", StringComparison.Ordinal) || text.EndsWith("\uff1a", StringComparison.Ordinal);
            }
            if (here > 0) blocks++;
        }

        if (moved > 0) { try { doc.Repaginate(); } catch (Exception) { } }

        string summary;
        if (dryRun)
        {
            summary = "\u8bd5\u8dd1\uff1a\u53ea\u62a5\u4e0d\u52a8"
                + " - " + moves.Count + " paragraph(s) could be brought in front of " + blocks + " figure(s)/table(s)"
                + " (minGapPt " + Math.Round(threshold, 1) + ")";
        }
        else if (moved == 0)
        {
            summary = "\u6ca1\u6709\u4e00\u9875\u5728\u56fe\u8868\u4e4b\u524d\u7559\u767d\u5230\u503c\u5f97\u642c\u6bb5"
                + " - no page is left part-empty before a figure or table enough to be worth moving text up"
                + " (minGapPt " + Math.Round(threshold, 1) + ")";
        }
        else
        {
            summary = "\u628a\u6b63\u6587\u6bb5\u63d0\u5230\u56fe\u8868\u4e4b\u524d\uff0c\u586b\u4e0a\u5b83\u4eec\u4e0a\u4e00\u9875\u5e95\u90e8\u7684\u7559\u767d"
                + " - " + moved + " paragraph(s) brought in front of " + blocks + " figure(s)/table(s)"
                + "; \u642c\u5b8c\u8bf7\u76ee\u68c0\uff1a\u56fe/\u8868\u662f\u5426\u8fd8\u5728\u539f\u4f4d\u3001\u6709\u6ca1\u6709\u8d8a\u8fc7\u6807\u9898";
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["dryRun"] = dryRun;
        r["moved"] = moved;
        r["blocks"] = blocks;
        r["minGapPt"] = Math.Round(threshold, 1);
        r["maxMoves"] = maxMoves;
        r["figures"] = figures;
        r["tables"] = tables;
        r["moves"] = moves;
        r["notes"] = notes;
        r["summary"] = summary;
        r["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["tracked"] = false;
        r["saved"] = false;
        r["note"] = "Only body text moves: it is copied above the figure/table and deleted below it, so the figure/table keeps its place. "
            + "Headings are never crossed. Nothing was saved - render the pages named in moves to check them, then save.";
        return r;
    }

    /// <summary>
    /// Put a copy of `para` in front of the block. A paragraph block takes the formatted text at
    /// its own start, but a table does not: assigning there lands INSIDE the first cell (measured
    /// on WPS 12.1 - the text turned up above "cell 1,1"). So for a table the paragraph before it
    /// is split first ("\r" at its paragraph mark, the same trick the insert command uses) and the
    /// empty paragraph that appears between the two is replaced by the text.
    /// </summary>
    private static bool GapCopyAbove(dynamic doc, int index, int start, dynamic para)
    {
        if (index > 0 || start <= 1)
        {
            doc.Range(start, start).FormattedText = para.Range.FormattedText;
            return true;
        }
        // A table's own start is inside its first cell, so split the paragraph before it and use
        // the empty paragraph in between. If WPS put that break inside the cell after all, take it
        // back and leave the block alone rather than dropping the text into a cell.
        int mark = start - 1;
        doc.Range(mark, mark).InsertBefore("\r");
        dynamic probe = doc.Range(mark + 1, mark + 1);
        if (GapInTable(probe))
        {
            try { doc.Range(mark + 1, mark + 2).Delete(); } catch (Exception) { }
            return false;
        }
        // Replace the empty paragraph's *text position*, never its mark: deleting the paragraph
        // mark that sits in front of a table merges the table into that paragraph, which is how
        // the moved text ended up inside the first cell instead of above the table.
        probe.FormattedText = para.Range.FormattedText;
        return true;
    }

    // ------------------------------------------------------------ geometry    /// <summary>bottom (page height - bottom margin) and top margin of the paragraph's section, in points.</summary>
    private static double[] GapGeometry(dynamic doc, dynamic para)
    {
        dynamic setup = null;
        if (para != null)
        {
            try { setup = para.Range.Sections.Item(1).PageSetup; }
            catch (Exception) { setup = null; }
        }
        if (setup == null)
        {
            try { setup = doc.PageSetup; }
            catch (Exception) { setup = null; }
        }
        double pageH = 0, bottomM = 0, topM = 0;
        if (setup != null)
        {
            pageH = Num(TryGet(() => (object)setup.PageHeight, null), 0);
            bottomM = Num(TryGet(() => (object)setup.BottomMargin, null), 0);
            topM = Num(TryGet(() => (object)setup.TopMargin, null), 0);
        }
        if (pageH <= 0) { pageH = 792; bottomM = 72; topM = 72; }
        return new double[] { pageH - bottomM, topM };
    }

    private static bool GapTruthy(object v)
    {
        if (v == null) return false;
        if (v is bool) return (bool)v;
        try { return Convert.ToInt32(v, CultureInfo.InvariantCulture) != 0; }
        catch (Exception) { return false; }
    }

    private static bool GapInTable(dynamic range)
    {
        try { return GapTruthy(range.Information(FG_INTABLE)); }
        catch (Exception) { return false; }
    }

    // ------------------------------------------------------------ block / absorb / widen

    /// <summary>
    /// Where one block sits: a figure (index &gt; 0, with the empty paragraphs before it and its
    /// caption after it) or a table (index &lt; 0, with its caption before it).
    /// </summary>
    private static bool GapBlock(dynamic doc, int index, out int start, out int end)
    {
        start = 0;
        end = 0;
        if (index > 0)
        {
            dynamic para = null;
            try { para = doc.InlineShapes.Item(index).Range.Paragraphs.Item(1); }
            catch (Exception) { return false; }
            if (GapInTable(para.Range)) return false;       // a picture inside a cell is not a block of its own

            dynamic first = para;
            dynamic before = null;
            try { before = para.Previous(); } catch (Exception) { before = null; }
            int guard = 0;
            while (before != null && guard++ < 20)
            {
                string t = CleanPara(Convert.ToString(TryGet(() => (object)before.Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
                int shapes = (int)TryGet(() => (object)(int)before.Range.InlineShapes.Count, 0);
                if (t.Length > 0 || shapes > 0) break;
                first = before;
                try { before = before.Previous(); } catch (Exception) { before = null; }
            }

            dynamic last = para;
            dynamic next = null;
            try { next = para.Next(); } catch (Exception) { next = null; }
            if (next != null && GapIsCaption(doc, next)) last = next;

            start = (int)TryGet(() => (object)(int)first.Range.Start, 0);
            end = (int)TryGet(() => (object)(int)last.Range.End, 0);
            GapWiden(doc, ref start, ref end);
            return start > 0 && end > start;
        }

        dynamic table = null;
        try { table = doc.Tables.Item(-index); }
        catch (Exception) { return false; }
        start = (int)TryGet(() => (object)(int)table.Range.Start, 0);
        end = (int)TryGet(() => (object)(int)table.Range.End, 0);
        if (start > 1)
        {
            dynamic above = null;
            try { above = doc.Range(start - 1, start - 1).Paragraphs.Item(1); }
            catch (Exception) { above = null; }
            if (above != null && GapIsCaption(doc, above))
                start = (int)TryGet(() => (object)(int)above.Range.Start, start);
        }
        GapWiden(doc, ref start, ref end);
        return start > 0 && end > start;
    }

    /// <summary>Figures and tables that follow one another directly move as one block.</summary>
    private static void GapAbsorb(dynamic doc, ref int start, ref int end)
    {
        for (int guard = 0; guard < 4; guard++)
        {
            int last = (int)TryGet(() => (object)(int)doc.Content.End, 0);
            if (last <= 2 || end >= last - 1) return;
            dynamic next = null;
            try { next = doc.Range(end, end).Paragraphs.Item(1); }
            catch (Exception) { return; }
            int to = 0;
            if (GapInTable(next.Range))
                to = (int)TryGet(() => (object)(int)next.Range.Tables.Item(1).Range.End, 0);
            else if ((int)TryGet(() => (object)(int)next.Range.InlineShapes.Count, 0) > 0)
            {
                to = (int)TryGet(() => (object)(int)next.Range.End, 0);
                dynamic caption = null;
                try { caption = next.Next(); } catch (Exception) { caption = null; }
                if (caption != null && GapIsCaption(doc, caption))
                    to = (int)TryGet(() => (object)(int)caption.Range.End, 0);
            }
            else if (GapIsCaption(doc, next))
            {
                dynamic under = null;
                try { under = next.Next(); } catch (Exception) { under = null; }
                if (under != null && GapInTable(under.Range))
                    to = (int)TryGet(() => (object)(int)under.Range.Tables.Item(1).Range.End, 0);
            }
            if (to <= end) return;
            end = to;
            GapWiden(doc, ref start, ref end);
        }
    }

    /// <summary>Empty paragraphs around a figure or table belong to it: they are its spacing.</summary>
    private static void GapWiden(dynamic doc, ref int start, ref int end)
    {
        int last = (int)TryGet(() => (object)(int)doc.Content.End, 0);
        for (int guard = 0; guard < 6 && start > 1; guard++)
        {
            dynamic above = null;
            try { above = doc.Range(start - 1, start - 1).Paragraphs.Item(1); }
            catch (Exception) { return; }
            string t = CleanPara(Convert.ToString(TryGet(() => (object)above.Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
            int shapes = (int)TryGet(() => (object)(int)above.Range.InlineShapes.Count, 0);
            if (t.Length > 0 || shapes > 0 || GapInTable(above.Range)) break;
            start = (int)TryGet(() => (object)(int)above.Range.Start, start);
        }
        for (int guard = 0; guard < 6 && end < last - 1; guard++)
        {
            dynamic below = null;
            try { below = doc.Range(end, end).Paragraphs.Item(1); }
            catch (Exception) { return; }
            string t = CleanPara(Convert.ToString(TryGet(() => (object)below.Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
            int shapes = (int)TryGet(() => (object)(int)below.Range.InlineShapes.Count, 0);
            if (t.Length > 0 || shapes > 0 || GapInTable(below.Range)) break;
            end = (int)TryGet(() => (object)(int)below.Range.End, end);
        }
    }

    /// <summary>
    /// A caption paragraph: the document's own caption parser (the one `captions` and `audit`
    /// use, so the three can never disagree) or a paragraph whose style says Caption / the
    /// Chinese caption style name (\u9898\u6ce8 is matched by its escape in the code below).
    /// </summary>
    private static bool GapIsCaption(dynamic doc, dynamic para)
    {
        try
        {
            string text = CleanPara(Convert.ToString(TryGet(() => (object)para.Range.Text, ""), CultureInfo.InvariantCulture));
            if (ParseCaption(text) != null) return true;
        }
        catch (Exception) { }
        try
        {
            string style = Convert.ToString(TryGet(() => (object)para.Style.NameLocal, ""), CultureInfo.InvariantCulture);
            if (style != null && style.Length > 0
                && (style.IndexOf("Caption", StringComparison.OrdinalIgnoreCase) >= 0
                    || style.IndexOf("\u9898\u6ce8", StringComparison.Ordinal) >= 0)) return true;
        }
        catch (Exception) { }
        return false;
    }
}
