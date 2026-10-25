// v0.5 search: find arbitrary text across the document and get paragraph numbers back.
// read pages with -From/-Count are clumsy for a 1000-paragraph report; this is the locator
// that revise's "expect" anchors and comment's -Para need.
//
//   {"doc":"...","text":"...","style":"optional NameLocal filter","max":50,"caseSensitive":false}
//
// Returns {"hits":[{"para","style","inTable","at","context"}],"paragraphsScanned","totalHits",
//          "truncated"} - "at" is the 0-based character offset inside the paragraph text.

using System;
using System.Collections.Generic;
using System.Globalization;

internal static partial class Program
{
    private static object SearchDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        string needle = a.Need("text");
        if (needle.Trim().Length == 0) throw new Fail("BAD_ARGS", "\"text\" is empty");
        string styleFilter = a.Str("style", null);
        if (styleFilter != null) styleFilter = styleFilter.Trim();
        bool caseSensitive = a.Flag("caseSensitive", false);
        int maxHits = a.Int("max", 50);
        if (maxHits < 1) maxHits = 1;

        // Fast path: one Content.Text fetch + in-memory match (multi-term via '|'),
        // per-hit offset->paragraph conversion. Falls back to the loop below wholesale.
        object fast = SearchFast(doc, needle, styleFilter, caseSensitive, maxHits);
        if (fast != null) return fast;

        int maxPerPara = 3;

        int total = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);

        // Table spans so hits can be flagged inTable (same trick as read).
        List<int[]> spans = new List<int[]>();
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        for (int t = 1; t <= tables; t++)
        {
            try
            {
                dynamic tb = doc.Tables.Item(t);
                spans.Add(new int[] { (int)tb.Range.Start, (int)tb.Range.End });
            }
            catch (Exception) { spans.Add(new int[] { -1, -1 }); }
        }

        List<object> hits = new List<object>();
        int totalHits = 0;
        bool truncated = false;
        StringComparison cmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        for (int k = 1; k <= total; k++)
        {
            dynamic p = null;
            try { p = doc.Paragraphs.Item(k); } catch (Exception) { continue; }
            string text = CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));
            if (text.Length == 0) continue;
            string style = Convert.ToString(TryGet(() => (object)p.Style.NameLocal, ""), CultureInfo.InvariantCulture);
            if (styleFilter != null && styleFilter.Length > 0 && string.Compare(style, styleFilter, StringComparison.OrdinalIgnoreCase) != 0)
                continue;

            int at = IndexOf(text, needle, 0, cmp);
            if (at < 0) continue;

            int start = -1, end = -1;
            try { start = (int)p.Range.Start; end = (int)p.Range.End; } catch (Exception) { }
            bool inTable = false;
            foreach (int[] span in spans)
                if (span[0] >= 0 && start >= span[0] && start <= span[1]) { inTable = true; break; }

            int inThis = 0;
            while (at >= 0)
            {
                totalHits++;
                inThis++;
                if (inThis <= maxPerPara && !truncated)
                {
                    if (hits.Count >= maxHits) truncated = true;
                    else
                    {
                        Dictionary<string, object> one = new Dictionary<string, object>();
                        one["para"] = k;
                        one["style"] = style;
                        if (inTable) one["inTable"] = true;
                        one["at"] = at;
                        one["context"] = ContextOf(text, at, needle.Length);
                        hits.Add(one);
                    }
                }
                at = IndexOf(text, needle, at + 1, cmp);
            }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["text"] = needle;
        if (styleFilter != null) r["style"] = styleFilter;
        r["paragraphs"] = total;
        r["hits"] = hits;
        r["totalHits"] = totalHits;
        if (totalHits > hits.Count) r["truncated"] = true;
        r["saved"] = false;
        r["note"] = "para 就是要喂给 revise(-Para)/comment(-Para)/view(-Para) 的段落号；命中文本是「接受视图」（与 read/diff 同一口径）。";
        return r;
    }

    private static int IndexOf(string haystack, string needle, int start, StringComparison cmp)
    {
        if (haystack == null || needle == null || needle.Length == 0) return -1;
        return haystack.IndexOf(needle, start, cmp);
    }

    /// <summary>About 40 characters around a hit, so the caller can tell one hit from another.</summary>
    private static string ContextOf(string text, int at, int len)
    {
        int from = Math.Max(0, at - 20);
        int take = Math.Min(text.Length - from, (at - from) + len + 20);
        string ctx = text.Substring(from, take);
        if (from > 0) ctx = "\u2026" + ctx;
        if (from + take < text.Length) ctx = ctx + "\u2026";
        return ctx;
    }

    // ------------------------------------------------------------ selection

    /// <summary>
    /// selection - read-only peek at what the human currently has selected / where the cursor is.
    ///   {"doc":"..."}
    /// Returns {"text","paras":[...],"fromPara","toPara","style","inTable","length"}. "paras" are the
    /// same numbers revise/comment/format take, so "改我选中的这段" becomes selection -> revise.
    /// A bare cursor (no selection) reports the paragraph it sits in with length 0.
    /// </summary>
    private static object SelectionDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        dynamic app = App(false, false);
        if (app == null) throw new Fail("NOT_OPEN", "No WPS instance is reachable.");
        dynamic sel = null;
        try { sel = app.Selection; } catch (Exception) { }
        if (sel == null) throw new Fail("COM_ERROR", "This WPS instance does not expose a Selection.");

        string text = "";
        int selStart = -1, selEnd = -1;
        try
        {
            text = CleanPara(Convert.ToString(sel.Text, CultureInfo.InvariantCulture));
            selStart = (int)sel.Range.Start;
            selEnd = (int)sel.Range.End;
        }
        catch (Exception) { }

        // Make sure the selection belongs to OUR document, not another open window.
        try
        {
            int docStart = (int)doc.Range().Start;
            int docEnd = (int)doc.Range().End;
            if (selStart < docStart || selEnd > docEnd)
                throw new Fail("NOT_OPEN", "The current selection is in a DIFFERENT document/window (this skill is driving \"" + doc.Name + "\"). Click inside that document in WPS and try again.");
        }
        catch (Fail) { throw; }
        catch (Exception) { }

        // Which paragraph numbers the selection covers (same numbering read/revise use).
        // Fast: offset -> paragraph in one roundtrip per end; slow: the old linear scan.
        int total = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        List<object> paras = new List<object>();
        int from = 0, to = 0;
        string style = "";
        bool inTable = false;
        int tableIndex = 0, tableRow = 0, tableCol = 0;
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);

        from = ParaAt(doc, selStart, true);
        to = selEnd > selStart ? ParaAt(doc, selEnd - 1, false) : from;
        bool scanned = false;
        if (from == 0 || to == 0 || to < from)
        {
            // host refused or mapping looked wrong: the linear scan still answers
            from = 0; to = 0;
            for (int k = 1; k <= total; k++)
            {
                dynamic p = null;
                try { p = doc.Paragraphs.Item(k); } catch (Exception) { continue; }
                int s = -1, e = -1;
                try { s = (int)p.Range.Start; e = (int)p.Range.End; } catch (Exception) { }
                if (s < 0) continue;
                if (e <= selStart) continue;          // entirely before the selection
                if (s >= selEnd && selEnd > selStart) break; // entirely after a real selection
                if (s >= selEnd && selEnd == selStart && s > selStart) break; // cursor: stop past its paragraph
                if (from == 0)
                {
                    scanned = true;
                    from = k;
                    style = Convert.ToString(TryGet(() => (object)p.Style.NameLocal, ""), CultureInfo.InvariantCulture);
                    for (int t = 1; t <= tables; t++)
                    {
                        try
                        {
                            dynamic tb = doc.Tables.Item(t);
                            if (s >= (int)tb.Range.Start && s <= (int)tb.Range.End)
                            {
                                inTable = true;
                                tableIndex = t;
                                try
                                {
                                    dynamic c = p.Range.Cells.Item(1);
                                    tableRow = (int)c.RowIndex;
                                    tableCol = (int)c.ColumnIndex;
                                }
                                catch (Exception) { }
                                break;
                            }
                        }
                        catch (Exception) { }
                    }
                }
                to = k;
                paras.Add(k);
                if (paras.Count >= 200) break;
            }
        }
        else
        {
            // fast path worked: bound the list, then one paragraph object for style/table info
            for (int k = from; k <= to && paras.Count < 200; k++) paras.Add(k);
            dynamic p0 = null;
            try { p0 = doc.Paragraphs.Item(from); } catch (Exception) { }
            if (p0 != null)
            {
                style = Convert.ToString(TryGet(() => (object)p0.Style.NameLocal, ""), CultureInfo.InvariantCulture);
                int s = -1;
                try { s = (int)p0.Range.Start; } catch (Exception) { }
                for (int t = 1; t <= tables; t++)
                {
                    try
                    {
                        dynamic tb = doc.Tables.Item(t);
                        if (s >= (int)tb.Range.Start && s <= (int)tb.Range.End)
                        {
                            inTable = true;
                            tableIndex = t;
                            try
                            {
                                dynamic c = p0.Range.Cells.Item(1);
                                tableRow = (int)c.RowIndex;
                                tableCol = (int)c.ColumnIndex;
                            }
                            catch (Exception) { }
                            break;
                        }
                    }
                    catch (Exception) { }
                }
            }
        }

        // Inline images: report any image inside the selection, or inside the paragraphs the
        // selection covers ("change this image" -> picture -Image n). Floating shapes are not
        // addressable by picture -Image (InlineShapes only), so they are ignored here.
        List<object> images = new List<object>();
        try
        {
            int inl = (int)doc.InlineShapes.Count;
            int coverStart = selStart, coverEnd = selEnd;
            if (from > 0 && to > 0)
            {
                try
                {
                    coverStart = Math.Min(coverStart, (int)doc.Paragraphs.Item(from).Range.Start);
                    coverEnd = Math.Max(coverEnd, (int)doc.Paragraphs.Item(to).Range.End);
                }
                catch (Exception) { }
            }
            for (int n = 1; n <= inl; n++)
            {
                dynamic sh = null;
                try { sh = doc.InlineShapes.Item(n); } catch (Exception) { continue; }
                try
                {
                    int s0 = (int)sh.Range.Start, e0 = (int)sh.Range.End;
                    if (e0 < coverStart || s0 > coverEnd) continue;
                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["image"] = n;   // the n of  picture -Image n
                    one["type"] = Convert.ToString(TryGet(() => (object)sh.Type, ""), CultureInfo.InvariantCulture);
                    one["widthPt"] = Math.Round((double)(float)sh.Width, 1);
                    one["heightPt"] = Math.Round((double)(float)sh.Height, 1);
                    images.Add(one);
                }
                catch (Exception) { }
            }
        }
        catch (Exception) { }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["text"] = text.Length > 2000 ? text.Substring(0, 2000) : text;
        r["length"] = text.Length;
        r["paras"] = paras;
        r["fromPara"] = from;
        r["toPara"] = to;
        if (style.Length > 0) r["style"] = style;
        if (inTable)
        {
            r["inTable"] = true;
            r["table"] = tableIndex;          // the n of  table -Table n  /  cell -Table n
            if (tableRow > 0) r["row"] = tableRow;
            if (tableCol > 0) r["col"] = tableCol;
        }
        if (images.Count > 0)
        {
            r["images"] = images;             // the n of  picture -Image n
        }
        r["saved"] = false;
        r["note"] = from == 0
            ? "\u6ca1\u627e\u5230\u5149\u6807/\u9009\u533a\u6240\u5728\u6bb5\uff08\u53ef\u80fd\u9009\u5728\u522b\u7684\u7a97\u53e3\uff09\u3002"
            : "\u8fd9\u4e9b\u6bb5\u843d\u53f7\u5c31\u662f read/revise/comment/format \u7528\u7684\u7f16\u53f7\uff1a\u6539\u9009\u4e2d\u6bb5\u7528 revise -Items\uff08expect \u53d6\u9009\u4e2d\u6587\u5f00\u5934\uff09\uff0c\u6392\u7248\u7528 format -From "
            + from + " -To " + to
            + (inTable ? "\u3002\u9009\u4e2d\u5728\u8868\u683c\u91cc\uff1a\u8868\u53f7 " + tableIndex + "\uff08table -Table " + tableIndex + " \u6539\u6846\u7ebf/\u5b57\u53f7\uff0ccell -Table " + tableIndex + " -Row r -Col c \u6539\u5355\u5143\u683c\uff09\uff0c\u5149\u6807\u5728\u7b2c " + tableRow + " \u884c\u7b2c " + tableCol + " \u5217\u3002" : "\u3002\u547d\u4ee4\u672c\u8eab\u4e0d\u6539\u6587\u6863\u3002");
        return r;
    }


    // ------------------------------------------------------------ undo

    /// <summary>
    /// undo — the agent's own Ctrl+Z. Formatting batches (format/style/table/...) are one undo
    /// record by design; this is how the agent backs one out again without a human at the keyboard.
    ///   {"doc":"...","count":1}
    /// Guardrail: Undo also rolls back HUMAN edits made after ours. Only call it right after your
    /// own write and only when no changedOutside warning appeared. Capped at 20 steps.
    /// </summary>
    private static object UndoDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only.");
        int want = a.Int("count", 1);
        if (want < 1) want = 1;
        if (want > 20) want = 20;
        int done = 0;
        for (int i = 0; i < want; i++)
        {
            bool ok = false;
            try { ok = (bool)doc.Undo(); } catch (Exception) { ok = false; }
            if (!ok) break;
            done++;
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["requested"] = want;
        r["undone"] = done;
        r["revisions"] = TryGet(() => (object)(int)doc.Revisions.Count, 0);
        r["saved"] = false;
        r["note"] = done == 0
            ? "\u6ca1\u6709\u53ef\u64a4\u7684\u64cd\u4f5c\uff08\u64a4\u9500\u6808\u5df2\u7a7a\uff09\u3002"
            : "\u64a4\u4e86 " + done + " \u6b65\u3002\u6ce8\u610f\uff1aUndo \u4e0d\u533a\u5206\u4f5c\u8005\uff0c\u5982\u679c\u4f60\u5199\u5165\u4e4b\u540e\u4eba\u4e5f\u6539\u8fc7\uff0c\u8fd9\u4e9b\u6539\u52a8\u4e5f\u4f1a\u4e00\u8d77\u88ab\u64a4\u6389\uff1b\u53ea\u5728\u81ea\u5df1\u521a\u5199\u5b8c\u3001\u4e14\u6ca1\u6709 changedOutside \u65f6\u4f7f\u7528\u3002\u672a\u4fdd\u5b58\uff0c\u786e\u8ba4\u540e\u518d save\u3002";
        return r;
    }
}
