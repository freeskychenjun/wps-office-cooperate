// P3 table of contents, reference lists / citations, and window scrolling.
// Part of the wps-word-cooperate helper (C# 5, compiled together with WpsHelper.cs).
// Pure ASCII source: Chinese travels only inside JSON strings or \u escapes.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

internal static partial class Program
{
    // WdConstants that matter here.
    private const int WD_COLLAPSE_END = 0;
    private const int WD_COLLAPSE_START = 1;
    private const int WD_CHARACTER = 1;
    private const int WD_ALIGN_LEFT = 0;
    private const int WD_ALIGN_CENTER = 1;
    private const int WD_OUTLINE_BODY_TEXT = 10;
    private const int WD_GOTO_PAGE = 1;
    private const int WD_GOTO_ABSOLUTE = 1;
    private const int WD_INFORMATION_ACTIVE_END_PAGE = 3;

    // ------------------------------------------------------------ small shared helpers

    /// <summary>One undo step (Ctrl+Z takes the whole command back). Returns null if WPS will not group.</summary>
    private static dynamic BeginUndo(string label)
    {
        dynamic app = App(true, true);
        dynamic undo = null;
        try { undo = app.UndoRecord; undo.StartCustomRecord(label); } catch (Exception) { undo = null; }
        return undo;
    }

    private static void EndUndo(dynamic undo)
    {
        if (undo == null) return;
        try { undo.EndCustomRecord(); } catch (Exception) { }
    }

    private static bool Tracking(dynamic doc)
    {
        try { return (bool)doc.TrackRevisions; } catch (Exception) { return false; }
    }

    private static void SetTracking(dynamic doc, bool on)
    {
        try { doc.TrackRevisions = on; } catch (Exception) { }
    }

    private static string TocTextAt(dynamic doc, int index)
    {
        try
        {
            string t = Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(index).Range.Text, ""), CultureInfo.InvariantCulture);
            return CleanPara(t);
        }
        catch (Exception) { return ""; }
    }

    private static string DocName(dynamic doc)
    {
        return Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
    }

    /// <summary>Index of the last paragraph that has text (the body's trailing empty paragraph is skipped).</summary>
    private static int LastTextParagraph(dynamic doc)
    {
        int total = (int)doc.Paragraphs.Count;
        int floor = Math.Max(1, total - 30);
        for (int i = total; i >= floor; i--)
            if (TocTextAt(doc, i).Trim().Length > 0) return i;
        return total;
    }

    // ------------------------------------------------------------ toc

    /// <summary>
    /// toc — insert and update a table of contents.
    ///   {"para":N,"title":"\u76ee\u5f55","levels":3}   (para omitted = before the first paragraph)
    /// Returns {"toc":n,"entries":N,"tablesOfContents":N,"atPara"}.
    /// </summary>
    private static object TocDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no table of contents was inserted.");

        int total = (int)doc.Paragraphs.Count;
        int at = a.Int("para", 0);
        if (at < 1) at = 1;                                  // omitted: right at the beginning of the document
        if (at > total) throw new Fail("ANCHOR_MISSING", "paragraph " + at + " does not exist (the document has " + total + ").");

        int levels = a.Int("levels", 3);
        if (levels < 1) levels = 1;
        if (levels > 9) levels = 9;
        string title = a.Str("title", "");
        bool wantTitle = title != null && title.Trim().Length > 0;
        title = wantTitle ? title.Trim() : "";
        bool track = a.Flag("track", false);
        bool numbers = a.Flag("numbers", false);             // keep an inherited heading numbering? no by default
        bool wasTracking = Tracking(doc);

        dynamic undo = BeginUndo("wps-word-cooperate toc");
        List<int> created = new List<int>();
        int tocPara = 0, tocCount = 0, entries = 0, pages = 0, holderRemoved = 0;
        bool titleStyled = false;
        bool tocError = false;
        string preview = "";
        string tocText = "";
        try
        {
            if (!track) SetTracking(doc, false);

            // 1. the title paragraph, inserted before the target (its own paragraph, never numbered)
            if (wantTitle)
            {
                dynamic anchor = doc.Paragraphs.Item(at).Range;
                anchor.InsertParagraphBefore();
                created.Add(at);
                dynamic tp = doc.Paragraphs.Item(at);
                tp.Range.InsertBefore(title);
                try { tp.Format.Alignment = WD_ALIGN_CENTER; } catch (Exception) { }
                try { tp.Range.Font.Bold = 1; } catch (Exception) { }
                if (a.Has("size")) { try { tp.Range.Font.Size = (float)Points(a.Str("size", "16")); } catch (Exception) { } }
                try { tp.Format.SpaceBefore = 0; tp.Format.SpaceAfter = 6; } catch (Exception) { }
                if (!numbers)
                {
                    try { tp.Range.ListFormat.RemoveNumbers(); } catch (Exception) { }
                    try { tp.Format.OutlineLevel = WD_OUTLINE_BODY_TEXT; } catch (Exception) { }
                }
                // A title that inherited the style of the heading it sits before would otherwise show up in the TOC.
                try { tp.Style = FindStyle(doc, "Normal"); titleStyled = true; }
                catch (Exception) { }
                tocPara = at + 1;
            }
            else tocPara = at;

            // 2. an empty paragraph that becomes the TOC (the canonical Word recipe: Add on the empty paragraph)
            doc.Paragraphs.Item(tocPara).Range.InsertParagraphBefore();
            created.Add(tocPara);
            dynamic holder = doc.Paragraphs.Item(tocPara).Range;
            holder.Collapse(WD_COLLAPSE_START);
            doc.TablesOfContents.Add(holder, true, 1, levels);

            tocCount = (int)TryGet(() => (object)(int)doc.TablesOfContents.Count, 0);
            entries = CountTocEntries(doc);
            if (entries == 0)
            {
                // A field that was never updated shows nothing: update it (and every other field) once.
                try { doc.TablesOfContents.Item(tocCount).Update(); } catch (Exception) { }
                try { doc.Fields.Update(); } catch (Exception) { }
                entries = CountTocEntries(doc);
            }
            // WPS leaves the paragraph mark of the empty paragraph the TOC was added to behind:
            // an empty line right under the last entry. Take that mark out again.
            try { holderRemoved = RemoveTocTail(doc, tocPara, entries, tocCount); } catch (Exception) { }
            try { doc.Repaginate(); } catch (Exception) { }
            pages = (int)TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
            tocText = TocText(doc, tocCount);
            if (TocMissing(tocText))
            {
                // WPS/Word put their own error line in place of the entries: there is no entry after all.
                entries = 0;
                tocError = true;
            }
            preview = Head(tocText, 200);
        }
        catch (Exception e)
        {
            // Nothing lasting: take back the paragraphs we created when the TOC itself never got in.
            if (tocCount == 0)
            {
                for (int i = created.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        int idx = created[i];
                        if (idx >= 1 && idx <= (int)doc.Paragraphs.Count) doc.Paragraphs.Item(idx).Range.Delete();
                    }
                    catch (Exception) { }
                }
            }
            if (!track) SetTracking(doc, wasTracking);
            EndUndo(undo);
            Fail fail = e as Fail;
            if (fail != null) throw fail;
            throw new Fail("COM_ERROR", Clean(e.Message));
        }
        if (!track) SetTracking(doc, wasTracking);
        EndUndo(undo);

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = DocName(doc);
        r["toc"] = tocCount;
        r["entries"] = entries;
        r["tablesOfContents"] = tocCount;
        r["atPara"] = tocPara;
        r["title"] = title;
        r["levels"] = levels;
        r["preview"] = preview;
        r["tocError"] = tocError;
        r["tailRemoved"] = holderRemoved;
        r["pages"] = pages;
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["tracked"] = track;
        r["saved"] = false;
        if (tocCount == 0)
            r["note"] = "WPS reported no table-of-contents object after the insert; nothing usable was added. Nothing was saved.";
        else if (tocError || entries == 0)
            r["note"] = "The table of contents field is in the document, but it has no entries: no paragraph uses a heading style (Heading 1-3 / \u6807\u9898 1-3), so the field shows \u201c"
                + "\u9519\u8bef\uff01\u672a\u627e\u5230\u76ee\u5f55\u9879\u3002" + "\u201d / \u201cError! No table of contents entries found.\u201d. "
                + "The field was left in place: give the chapter titles a heading style and update it (F9 / right click > Update Field). Nothing was saved.";
        else
            r["note"] = "The table of contents is a field: its page numbers are only as fresh as the last update (F9 / right click > Update Field). Nothing was saved; Ctrl+Z takes the insert back (WPS records the title formatting, the field insert and the clean-up as separate undo steps, so it may take a couple of presses)."
                + (wantTitle ? " The title paragraph is plain body text (centred, bold) so it does not appear as an entry itself." : " No title paragraph was inserted (title omitted).");
        if (titleStyled == false && wantTitle)
            r["styleNote"] = "The title paragraph keeps the style of its neighbour; its outline level was not reset, so it may be listed in the TOC itself.";
        return r;
    }

    /// <summary>
    /// Removes the empty paragraph mark that the TOC insert leaves right under the last entry.
    /// Only touches a paragraph that is empty and sits directly behind a paragraph inside the TOC field.
    /// </summary>
    private static int RemoveTocTail(dynamic doc, int tocPara, int entries, int tocIndex)
    {
        if (entries <= 0) return 0;
        int total = (int)doc.Paragraphs.Count;
        int cand = tocPara + entries;
        if (cand < tocPara + 1 || cand >= total) return 0;      // never take the document's last paragraph away
        dynamic toc = doc.TablesOfContents.Item(tocIndex);
        int tocStart = (int)toc.Range.Start;
        int prevStart = (int)doc.Paragraphs.Item(cand - 1).Range.Start;
        if (prevStart < tocStart) return 0;                     // the paragraph above is not part of the TOC
        if (TocTextAt(doc, cand).Trim().Length > 0) return 0;  // not empty: somebody else's paragraph
        doc.Paragraphs.Item(cand).Range.Delete();
        return cand;
    }

    /// <summary>Number of non-empty paragraphs inside the TOC's own field result (one per entry).</summary>
    private static int CountTocEntries(dynamic doc)
    {
        int count = 0;
        int n = (int)TryGet(() => (object)(int)doc.TablesOfContents.Count, 0);
        for (int i = 1; i <= n; i++)
        {
            try
            {
                dynamic rng = doc.TablesOfContents.Item(i).Range;
                int paragraphs = (int)rng.Paragraphs.Count;
                for (int k = 1; k <= paragraphs; k++)
                {
                    string t = CleanPara(Convert.ToString(TryGet(() => (object)rng.Paragraphs.Item(k).Range.Text, ""), CultureInfo.InvariantCulture));
                    if (t.Trim().Length > 0) count++;
                }
            }
            catch (Exception) { }
        }
        return count;
    }

    /// <summary>The whole result of the TOC field as one line (tab separated entries, like "chapter 2 1.1 x 2").</summary>
    private static string TocText(dynamic doc, int index)
    {
        string text = Convert.ToString(TryGet(() => (object)doc.TablesOfContents.Item(index).Range.Text, ""), CultureInfo.InvariantCulture);
        if (text == null) return "";
        text = text.Replace("\r", " ").Replace("\n", " ").Replace("\u0007", " ").Trim();
        while (text.IndexOf("  ", StringComparison.Ordinal) >= 0) text = text.Replace("  ", " ");
        return text;
    }

    /// <summary>
    /// WPS/Word show "\u9519\u8bef\uff01\u672a\u627e\u5230\u76ee\u5f55\u9879\u3002" / "Error! No table of contents entries found."
    /// as the field's result when no paragraph uses a heading style: that is not an entry.
    /// </summary>
    private static bool TocMissing(string text)
    {
        if (text == null || text.Length == 0) return true;
        if (text.IndexOf("\u672a\u627e\u5230\u76ee\u5f55\u9879", StringComparison.Ordinal) >= 0) return true;   // 未找到目录项
        if (text.IndexOf("\u9519\u8bef", StringComparison.Ordinal) >= 0) return true;                       // 错误
        if (text.IndexOf("No table of contents entries", StringComparison.OrdinalIgnoreCase) >= 0) return true;
        return false;
    }

    // ------------------------------------------------------------ reference

    /// <summary>
    /// reference — insert a reference list, or a numbered citation at a paragraph.
    ///   list : {"items":["Author. Title[M]. ..."],"para":N,"heading":"\u53c2\u8003\u6587\u732e"}
    ///   cite : {"para":N,"numbers":"1,2,5"|"3-6"}   -> superscript [1,2,5]
    /// Returns {"action","inserted","fields"}.
    /// </summary>
    private static object ReferenceDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; nothing was inserted.");
        string action = a.Str("action", "").Trim().ToLowerInvariant();
        if (action.Length == 0) action = a.Has("items") ? "list" : "cite";
        if (action == "list") return ReferenceList(doc, a);
        if (action == "cite") return ReferenceCite(doc, a);
        throw new Fail("BAD_ARGS", "unknown action \"" + action + "\" (list|cite)");
    }

    private static object ReferenceList(dynamic doc, Bag a)
    {
        List<object> items = a.List("items");
        if (items.Count == 0) throw new Fail("BAD_ARGS", "\"items\" is empty: give one string per reference");

        string heading = a.Str("heading", "");
        bool wantHeading = heading != null && heading.Trim().Length > 0;
        heading = wantHeading ? heading.Trim() : "";
        string headingStyle = a.Str("headingStyle", "");
        bool track = a.Flag("track", false);
        bool wasTracking = Tracking(doc);

        int total = (int)doc.Paragraphs.Count;
        int para = a.Int("para", 0);
        int at;
        if (para < 1) at = LastTextParagraph(doc);           // omitted: the end of the document body
        else
        {
            if (para > total) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not exist (the document has " + total + ").");
            at = para;
        }

        // Already the heading? Then do not write it a second time (re-running the command stays idempotent-ish).
        string anchorText = TocTextAt(doc, at).Trim();
        bool headingExists = wantHeading && anchorText.StartsWith(heading, StringComparison.Ordinal);

        dynamic undo = BeginUndo("wps-word-cooperate reference list");
        int headingAt = 0, first = 0, last = 0, inserted = 0;
        try
        {
            if (!track) SetTracking(doc, false);

            // The whole block (heading paragraph + one paragraph per entry) goes in with a single
            // assignment: one COM round trip, and the paragraph marks split it for us.
            StringBuilder block = new StringBuilder();
            foreach (object raw in items)
            {
                string text = Convert.ToString(raw, CultureInfo.InvariantCulture);
                if (text == null) continue;
                text = Clean(text);                          // one paragraph per entry: no stray line breaks
                if (text.Length == 0) continue;
                inserted++;
                if (block.Length > 0) block.Append("\r");
                block.Append("[").Append(inserted.ToString(CultureInfo.InvariantCulture)).Append("] ").Append(text);
            }
            if (inserted == 0) throw new Fail("BAD_ARGS", "\"items\" has no usable entry (all of them are empty)");

            bool addHeading = wantHeading && !headingExists;
            string whole = addHeading ? (heading + "\r" + block.ToString()) : block.ToString();

            // Insert after the anchor paragraph's text, leaving its paragraph mark where it is.
            dynamic anchor = doc.Paragraphs.Item(at).Range;
            anchor.MoveEnd(WD_CHARACTER, -1);
            anchor.Collapse(WD_COLLAPSE_END);
            anchor.Text = "\r" + whole;

            int lines = 1;
            for (int i = 0; i < whole.Length; i++) if (whole[i] == '\r') lines++;
            int linesAt = at + 1;
            if (addHeading)
            {
                headingAt = linesAt;
                try { doc.Paragraphs.Item(headingAt).Range.Font.Bold = 1; } catch (Exception) { }
                try { doc.Paragraphs.Item(headingAt).Format.Alignment = WD_ALIGN_LEFT; } catch (Exception) { }
                try { doc.Paragraphs.Item(headingAt).Format.SpaceBefore = 12; } catch (Exception) { }
                if (headingStyle.Length > 0)
                {
                    try { doc.Paragraphs.Item(headingAt).Style = FindStyle(doc, headingStyle); }
                    catch (Exception) { }
                }
            }
            first = linesAt + (addHeading ? 1 : 0);
            last = linesAt + lines - 1;
            if (inserted > 0) HangingIndentRange(doc, first, last, 21.0);
            try { doc.Repaginate(); } catch (Exception) { }
        }
        catch (Exception e)
        {
            if (!track) SetTracking(doc, wasTracking);
            EndUndo(undo);
            Fail fail = e as Fail;
            if (fail != null) throw fail;
            throw new Fail("COM_ERROR", Clean(e.Message));
        }
        if (!track) SetTracking(doc, wasTracking);
        EndUndo(undo);

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = DocName(doc);
        r["action"] = "list";
        r["inserted"] = inserted;
        r["fields"] = 0;
        r["heading"] = heading;
        r["headingAt"] = headingAt;
        r["headingExisted"] = headingExists;
        r["first"] = first;
        r["last"] = last;
        r["hangingIndent"] = 21;
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = "The list is plain text: the \u201c[1] \u201d numbers are typed, with a hanging indent (left 21pt, first line -21pt), not an automatic list, so they will not renumber themselves. "
            + (headingStyle.Length > 0 && !headingExists ? "The heading paragraph uses the style \u201c" + headingStyle + "\u201d. " : "The heading paragraph is plain bold text, not a heading style, so it does not enter a table of contents. ")
            + "Nothing was saved; Undo once (Ctrl+Z) takes the whole insert back.";
        if (!track) r["note"] = Convert.ToString(r["note"], CultureInfo.InvariantCulture) + " The inserts are not tracked revisions (pass track:true to make them revisions).";
        return r;
    }

    /// <summary>Hanging indent for a block of paragraphs: keeps wrapped reference lines lined up after the "[n] " label.</summary>
    private static void HangingIndentRange(dynamic doc, int firstPara, int lastPara, double width)
    {
        try
        {
            int start = (int)doc.Paragraphs.Item(firstPara).Range.Start;
            int end = (int)doc.Paragraphs.Item(lastPara).Range.End;
            dynamic pf = doc.Range(start, end).ParagraphFormat;
            // Character-unit indents win over point values in WPS/Word, so clear them first.
            try { pf.CharacterUnitFirstLineIndent = 0; } catch (Exception) { }
            try { pf.CharacterUnitLeftIndent = 0; } catch (Exception) { }
            pf.LeftIndent = (float)width;
            pf.FirstLineIndent = (float)(-width);
        }
        catch (Exception) { }
    }

    private static object ReferenceCite(dynamic doc, Bag a)
    {
        int para = a.Int("para", 0);
        if (para < 1) throw new Fail("BAD_ARGS", "\"para\" is required for action cite");
        int total = (int)doc.Paragraphs.Count;
        if (para > total) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not exist (the document has " + total + ").");

        string numbers = a.Need("numbers");
        List<int> list = ExpandNumbers(numbers);
        if (list.Count == 0) throw new Fail("BAD_ARGS", "no reference number in \"" + numbers + "\" (use \"1,2,5\" or \"3-6\")");
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i < list.Count; i++)
        {
            if (i > 0) sb.Append(",");
            sb.Append(list[i].ToString(CultureInfo.InvariantCulture));
        }
        sb.Append("]");
        string text = sb.ToString();

        bool superscript = a.Flag("super", true);
        bool track = a.Flag("track", false);
        bool wasTracking = Tracking(doc);

        dynamic undo = BeginUndo("wps-word-cooperate citation");
        int atPos = 0;
        string paragraphNow = "";
        try
        {
            if (!track) SetTracking(doc, false);
            dynamic p = doc.Paragraphs.Item(para);
            dynamic rng = p.Range;
            rng.MoveEnd(WD_CHARACTER, -1);                   // leave the paragraph mark alone
            int end = (int)rng.End;
            dynamic insertAt = doc.Range(end, end);
            insertAt.Text = text;
            atPos = end;
            dynamic inserted = doc.Range(end, end + text.Length);
            if (superscript)
            {
                try { inserted.Font.Superscript = true; }
                catch (Exception) { try { inserted.Font.Superscript = 1; } catch (Exception) { } }
            }
            paragraphNow = TocTextAt(doc, para);
        }
        catch (Exception e)
        {
            if (!track) SetTracking(doc, wasTracking);
            EndUndo(undo);
            Fail fail = e as Fail;
            if (fail != null) throw fail;
            throw new Fail("COM_ERROR", Clean(e.Message));
        }
        if (!track) SetTracking(doc, wasTracking);
        EndUndo(undo);

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = DocName(doc);
        r["action"] = "cite";
        r["text"] = text;
        r["fields"] = 0;
        r["para"] = para;
        r["atPos"] = atPos;
        r["superscript"] = superscript;
        r["paragraph"] = Head(paragraphNow, 120);
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = "The citation is superscript plain text at the end of the paragraph, not a cross-reference field: a REF field needs an existing bookmark to point at and would otherwise show \u201cError! Reference source not found.\u201d. "
            + "Nothing was saved; Undo once (Ctrl+Z) takes it back.";
        if (superscript == false) r["note"] = Convert.ToString(r["note"], CultureInfo.InvariantCulture) + " Superscript was switched off (super:false).";
        return r;
    }

    /// <summary>"1,2,5" or "3-6" (also "3\u20136", "3\uff5e6", "\uff0c" separators) -> ordered, de-duplicated numbers.</summary>
    private static List<int> ExpandNumbers(string spec)
    {
        List<int> list = new List<int>();
        HashSet<int> seen = new HashSet<int>();
        if (spec == null) return list;
        string s = spec
            .Replace("\uff0c", ",").Replace("\u3001", ",").Replace("\uff1b", ",").Replace(";", ",")
            .Replace(" ", ",").Replace("\t", ",").Replace("\r", ",").Replace("\n", ",")
            .Replace("\u2013", "-").Replace("\u2014", "-").Replace("\uff0d", "-")
            .Replace("\uff5e", "-").Replace("~", "-");
        string[] parts = s.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string raw in parts)
        {
            string t = raw.Trim();
            if (t.Length == 0) continue;
            int dash = t.IndexOf('-');
            if (dash > 0)
            {
                int from, to;
                if (int.TryParse(t.Substring(0, dash).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out from)
                    && int.TryParse(t.Substring(dash + 1).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out to))
                {
                    if (from > to) { int swap = from; from = to; to = swap; }
                    if (to - from > 500) throw new Fail("BAD_ARGS", "range \"" + t + "\" is too wide (at most 500 numbers)");
                    for (int n = from; n <= to; n++)
                    {
                        if (n < 1) throw new Fail("BAD_ARGS", "reference numbers start at 1 (got \"" + t + "\")");
                        if (seen.Add(n)) list.Add(n);
                    }
                    continue;
                }
            }
            int v;
            if (!int.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out v))
                throw new Fail("BAD_ARGS", "\"" + t + "\" is not a reference number or a range like 3-6");
            if (v < 1) throw new Fail("BAD_ARGS", "reference numbers start at 1 (got " + v + ")");
            if (seen.Add(v)) list.Add(v);
        }
        return list;
    }

    // ------------------------------------------------------------ view

    /// <summary>
    /// view — scroll the window to a paragraph or a page (the only "follow the edit" a skill can do).
    ///   {"para":N} or {"page":N}
    /// Returns {"scrolled","para","page"}.
    /// </summary>
    private static object ViewDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        int para = a.Int("para", 0);
        int page = a.Int("page", 0);
        if (para < 1 && page < 1) throw new Fail("BAD_ARGS", "\"para\" or \"page\" is required");
        if (para > 0) page = 0;                              // paragraph wins when both are given
        bool focus = a.Flag("focus", false);

        int total = (int)doc.Paragraphs.Count;
        if (para > total) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not exist (the document has " + total + ").");
        int pages = (int)TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        if (page > 0 && pages > 0 && page > pages)
            throw new Fail("BAD_ARGS", "page " + page + " does not exist (the document has " + pages + ").");

        dynamic app = App(true, true);

        // Bring this document's window forward only when it is not the one on screen already.
        bool activated = false;
        try
        {
            dynamic active = app.ActiveDocument;
            string had = active == null ? "" : DocName(active);
            if (!string.Equals(had, DocName(doc), StringComparison.OrdinalIgnoreCase))
            {
                doc.Activate();
                activated = true;
            }
        }
        catch (Exception) { }

        dynamic target;
        if (para > 0) target = doc.Paragraphs.Item(para).Range;
        else
        {
            target = doc.GoTo(WD_GOTO_PAGE, WD_GOTO_ABSOLUTE, page);
            if (target == null) throw new Fail("COM_ERROR", "WPS returned no range for page " + page + ".");
        }

        int caretBefore = -1;
        try { caretBefore = (int)app.Selection.Range.Start; } catch (Exception) { caretBefore = -1; }

        bool scrolled = false;
        bool selected = false;
        bool viaWindow = false;
        string error = "";
        try
        {
            dynamic win = TryGet(() => (object)doc.ActiveWindow, null) ?? TryGet(() => (object)doc.Windows.Item(1), null);
            if (win == null) throw new Fail("COM_ERROR", "the document has no window to scroll");
            if (focus)
            {
                target.Select();
                selected = true;
                win.ScrollIntoView(app.Selection.Range);
            }
            else
            {
                win.ScrollIntoView(target);                  // scrolls only; does not move the caret
            }
            scrolled = true;
            viaWindow = true;
        }
        catch (Exception e)
        {
            error = Clean(e.Message);
        }

        if (!scrolled)
        {
            // Fallback for a WPS build without Window.ScrollIntoView: selecting does move the caret.
            try
            {
                target.Select();
                selected = true;
                dynamic win = TryGet(() => (object)doc.ActiveWindow, null) ?? TryGet(() => (object)doc.Windows.Item(1), null);
                if (win != null) win.ScrollIntoView(app.Selection.Range);
                else doc.Windows.Item(1).ScrollIntoView(app.Selection.Range);
                scrolled = true;
                viaWindow = false;
            }
            catch (Exception e)
            {
                throw new Fail("COM_ERROR", "This WPS build refused both ScrollIntoView and Select: " + Clean(e.Message)
                    + (error.Length > 0 ? " (first attempt: " + error + ")" : ""));
            }
        }

        int caretAfter = -1;
        try { caretAfter = (int)app.Selection.Range.Start; } catch (Exception) { caretAfter = -1; }
        bool movedCaret = caretBefore >= 0 && caretAfter >= 0 && caretBefore != caretAfter;

        int landedPage = (int)TryGet(() => (object)(int)target.Information(WD_INFORMATION_ACTIVE_END_PAGE), 0);
        int vertical = (int)TryGet(() => (object)(int)doc.ActiveWindow.VerticalPercentScrolled, -1);

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = DocName(doc);
        r["scrolled"] = scrolled;
        r["para"] = para > 0 ? para : 0;
        r["page"] = landedPage > 0 ? landedPage : page;
        r["requestedPage"] = page;
        r["movedCaret"] = movedCaret;
        r["selected"] = selected;
        r["focus"] = focus;
        r["activated"] = activated;
        r["viaScrollIntoView"] = viaWindow;
        r["verticalPercentScrolled"] = vertical;
        r["pages"] = pages;
        r["saved"] = (bool)TryGet(() => (object)(bool)doc.Saved, true);
        if (focus && selected)
            r["note"] = "focus:true was given, so the text caret was put on the target (Range.Select) and the window scrolled there. The document content is unchanged and nothing was saved.";
        else if (movedCaret)
            r["note"] = "The view scrolled, but the text caret moved with it (this WPS build only scrolls after a Select). The document content is unchanged and nothing was saved.";
        else
            r["note"] = "The view scrolled through Window.ScrollIntoView; the document and the text caret were left alone. Nothing was saved.";
        if (viaWindow == false)
            r["note"] = Convert.ToString(r["note"], CultureInfo.InvariantCulture)
                + " Window.ScrollIntoView failed (" + error + "), so the fallback selected the range.";
        if (activated)
            r["note"] = Convert.ToString(r["note"], CultureInfo.InvariantCulture) + " The document was brought to the front first (it was not the active one).";
        return r;
    }
}
