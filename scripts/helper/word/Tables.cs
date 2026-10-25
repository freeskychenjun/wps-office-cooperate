// P2 tables / captions / pictures for the wps-word-cooperate helper.
// C# 5 (compiled by the in-box csc.exe together with WpsHelper.cs); Chinese only appears in
// user-facing messages and in values that are matched against request text, never against
// style names read from the document.
//
//   table   insert a table, or reformat existing ones
//   caption a numbered caption (SEQ field) for a table or a figure
//   picture insert an inline picture, or resize / crop one that is there

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

internal static partial class Program
{
    // ------------------------------------------------------------ Word object-model values
    // Spelled out because WPS has no interop assembly to name them from.

    private const int BT_TOP = -1;
    private const int BT_LEFT = -2;
    private const int BT_BOTTOM = -3;
    private const int BT_RIGHT = -4;
    private const int BT_HORZ = -5;
    private const int BT_VERT = -6;
    private const int BT_DIAG_DOWN = -7;
    private const int BT_DIAG_UP = -8;

    private const int LS_NONE = 0;        // wdLineStyleNone
    private const int LS_SINGLE = 1;      // wdLineStyleSingle
    private const int LW_050 = 4;         // wdLineWidth050pt
    private const int LW_075 = 6;         // wdLineWidth075pt
    private const int LW_150 = 12;        // wdLineWidth150pt

    private const int PREF_POINTS = 1;    // wdPreferredWidthPoints
    private const int PREF_PERCENT = 2;   // wdPreferredWidthPercent

    private const int ROW_ATLEAST = 1;    // wdRowHeightAtLeast
    private const int ROW_EXACTLY = 2;    // wdRowHeightExactly

    private const int SEQ_FIELD = 12;     // wdFieldSequence
    private const int COLLAPSE_END = 0;   // wdCollapseEnd
    private const int ALIGN_ROW_CENTER = 1;

    // ------------------------------------------------------------ one undo step, no revisions

    /// <summary>
    /// The wrapper every P2 command runs its work in: one Ctrl+Z step for the whole command, and
    /// (unless the request says track:true) revision tracking off, because these are layout edits
    /// and tracked formatting would bury the text revisions. Dispose restores both.
    /// </summary>
    private sealed class P2Edit : IDisposable
    {
        private readonly dynamic doc;
        private readonly bool keepTracking;
        private readonly bool wasTracking;
        private readonly dynamic undo;

        public P2Edit(dynamic document, bool track, string what)
        {
            doc = document;
            keepTracking = track;
            wasTracking = false;
            try { wasTracking = (bool)doc.TrackRevisions; } catch (Exception) { }

            dynamic app = null;
            try { app = App(true, true); } catch (Exception) { }
            dynamic record = null;
            if (app != null) { try { record = app.UndoRecord; record.StartCustomRecord(what); } catch (Exception) { record = null; } }
            undo = record;

            try { doc.TrackRevisions = track; } catch (Exception) { }
        }

        public void Dispose()
        {
            try { doc.TrackRevisions = wasTracking; } catch (Exception) { }
            if (undo != null) { try { undo.EndCustomRecord(); } catch (Exception) { } }
        }
    }

    // ------------------------------------------------------------ small shared helpers

    /// <summary>Dev tracing: appends a step to the file in WPS_HELPER_TRACE, so a call that hangs leaves a trace.</summary>
    private static void Tr(string message)
    {
        try
        {
            string path = Environment.GetEnvironmentVariable("WPS_HELPER_TRACE");
            if (path == null || path.Length == 0) return;
            File.AppendAllText(path, DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) + " " + message + "\r\n", new System.Text.UTF8Encoding(false));
        }
        catch (Exception) { }
    }

    private static double Num(object value, double fallback)
    {
        if (value == null) return fallback;
        try { return Convert.ToDouble(value, CultureInfo.InvariantCulture); } catch (Exception) { return fallback; }
    }

    private static string ParaTextAt(dynamic doc, int n)
    {
        try
        {
            dynamic p = doc.Paragraphs.Item(n);
            return CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));
        }
        catch (Exception) { return ""; }
    }

    /// <summary>Width of the text column of the first section, in points.</summary>
    private static double TextWidth(dynamic doc)
    {
        try
        {
            dynamic ps = doc.PageSetup;
            double w = Num((object)ps.PageWidth, 0) - Num((object)ps.LeftMargin, 0) - Num((object)ps.RightMargin, 0);
            if (w > 72) return w;
        }
        catch (Exception) { }
        return 430;   // A4 with 3.17/2.54 cm margins, as a last resort
    }

    /// <summary>Index of the first paragraph whose start is at or after a character position (paragraph starts are ordered).</summary>
    private static int ParaIndexAtOrAfter(dynamic doc, int position)
    {
        int lo = 1;
        int hi = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 1);
        int answer = hi;
        while (lo <= hi)
        {
            int mid = (lo + hi) / 2;
            int start = (int)TryGet(() => (object)(int)doc.Paragraphs.Item(mid).Range.Start, -1);
            if (start >= position) { answer = mid; hi = mid - 1; }
            else lo = mid + 1;
        }
        return answer;
    }

    /// <summary>Which table (1-based) a paragraph sits in, or 0 for the body.</summary>
    private static int TabOfParagraph(dynamic doc, int para)
    {
        int start = (int)TryGet(() => (object)(int)doc.Paragraphs.Item(para).Range.Start, -1);
        if (start < 0) return 0;
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        for (int t = 1; t <= tables; t++)
        {
            int s = -1, e = -1;
            try { dynamic tb = doc.Tables.Item(t); s = (int)tb.Range.Start; e = (int)tb.Range.End; }
            catch (Exception) { continue; }
            if (start >= s && start < e) return t;
        }
        return 0;
    }

    /// <summary>The 1-based index of an item in a collection of ranges (Tables / InlineShapes).</summary>
    private static int RangeIndexOf(dynamic collection, dynamic item)
    {
        int start;
        try { start = (int)item.Range.Start; } catch (Exception) { return 0; }
        int n = (int)TryGet(() => (object)(int)collection.Count, 0);
        for (int i = 1; i <= n; i++)
        {
            int s = (int)TryGet(() => (object)(int)collection.Item(i).Range.Start, -2);
            if (s == start) return i;
        }
        return 0;
    }

    // ============================================================ table

    /// <summary>
    /// table — insert a table, or reformat an existing one.
    ///   insert : {"para":N,"data":[[...],[...]],"borders":"grid|three-line|none","header":true,
    ///             "widths":"2,3,2","caption":"...","font":"...","size":10.5,"align":"center","nested":false}
    ///   format : {"table":1,"borders":"...","font":"...","size":10.5,"align":"center",
    ///             "autofit":"window|content|fixed","rowHeight":"18pt","rowHeightRule":"exact|atLeast",
    ///             "header":true,"keepTogether":true}
    /// Returns {"name","action","table","rows","columns","atPara"|"tables","saved":false}.
    /// </summary>
    private static object TableDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no table was changed.");

        // "data" means insert; without it the command reformats tables that are already there.
        if (a.Has("data")) return TabInsert(doc, a);
        return TabFormat(doc, a);
    }

    // ------------------------------------------------------------ table: insert

    private static object TabInsert(dynamic doc, Bag a)
    {
        List<object> raw = a.List("data");
        if (raw.Count == 0)
            throw new Fail("BAD_ARGS", "\"data\" is empty; it must be an array of rows, e.g. [[\"head 1\",\"head 2\"],[\"a\",\"b\"]]");

        List<List<string>> cells = new List<List<string>>();
        int columns = 0;
        foreach (object one in raw)
        {
            List<object> row = one as List<object>;
            if (row == null) throw new Fail("BAD_ARGS", "every row of \"data\" must be an array of cell values (row " + (cells.Count + 1) + " is not).");
            if (row.Count == 0) throw new Fail("BAD_ARGS", "row " + (cells.Count + 1) + " of \"data\" is empty.");
            List<string> line = new List<string>();
            foreach (object cell in row)
                line.Add(cell == null ? "" : Convert.ToString(cell, CultureInfo.InvariantCulture));
            if (line.Count > columns) columns = line.Count;
            cells.Add(line);
        }
        foreach (List<string> line in cells)
            while (line.Count < columns) line.Add("");
        int rows = cells.Count;

        int total = (int)doc.Paragraphs.Count;
        int at;
        if (a.Has("para"))
        {
            at = a.Int("para", 0);
            if (at < 1 || at > total)
                throw new Fail("ANCHOR_MISSING", "paragraph " + at + " does not exist (the document has " + total + ").");
        }
        else at = total;   // end of the document

        string expect = a.Str("expect", "");
        if (expect.Length > 0)
        {
            string has = ParaTextAt(doc, at);
            if (!has.StartsWith(expect, StringComparison.Ordinal))
                throw new Fail("ANCHOR_MISSING", "paragraph " + at + " starts with \"" + Head(has, 40) + "\", not \"" + Head(expect, 40) + "\".");
        }

        bool nested = a.Flag("nested", false);
        int inTable = TabOfParagraph(doc, at);
        if (inTable > 0 && !nested)
            throw new Fail("BAD_ARGS", "paragraph " + at + " is inside table " + inTable + "; pass nested:true to put a table in that cell, or pick a paragraph after the table.");

        string bordersSpec = a.Str("borders", "grid");
        string caption = a.Str("caption", null);
        bool header = a.Flag("header", false);
        bool track = a.Flag("track", false);

        // Everything is validated before the document is touched, so a bad argument leaves no half a table behind.
        string borderMode = TabBorderMode(bordersSpec);
        List<double> ratio = a.Has("widths") ? TabWidthValues(a.Str("widths", ""), columns) : new List<double>();
        int tableAlign = a.Has("align") ? TabTableAlign(a.Str("align", "center")) : -1;
        int cellAlign = a.Has("cellAlign") ? Alignment(a.Str("cellAlign", "center")) : -1;

        int tableIndex = 0;
        int outRows = rows;
        int outColumns = columns;
        int atPara = at;
        string captionText = "";
        string captionField = "";
        List<string> applied = new List<string>();
        Dictionary<string, object> borderRead = null;
        List<object> widthsRead = new List<object>();

        using (P2Edit edit = new P2Edit(doc, track, "wps-word-cooperate table"))
        {
            int slot;
            if (nested)
            {
                slot = at;
            }
            else
            {
                // A caption sits above its table, so it is written first and the table follows it.
                if (caption != null && caption.Length > 0)
                {
                    Dictionary<string, object> cap = AddCaption(doc, at, "\u8868", caption, false, a.Str("captionAlign", "center"), a);
                    atPara = ToInt(cap, "para");
                    captionText = Convert.ToString(cap["text"], CultureInfo.InvariantCulture);
                    captionField = Convert.ToString(cap["field"], CultureInfo.InvariantCulture);
                    applied.Add("caption");
                    at = atPara;
                }
                dynamic anchor = doc.Paragraphs.Item(at);
                anchor.Range.InsertParagraphAfter();
                slot = at + 1;
            }

            dynamic holder = doc.Paragraphs.Item(slot);
            int start = (int)holder.Range.Start;
            dynamic target = doc.Range(start, start);
            int paragraphsBefore = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
            dynamic table = TabAdd(doc, target, rows, columns, nested, slot);
            if (nested)
            {
                // WPS answers a nested Tables.Add without an error, hands back a table object - and makes
                // nothing (the cell gains no paragraph and nothing renders). So check the document itself.
                int grew = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0) - paragraphsBefore;
                int madeRows = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
                Tr("nested: paragraphs grew by " + grew + ", table claims " + madeRows + " rows, document tables " + (int)TryGet(() => (object)(int)doc.Tables.Count, 0));
                if (grew < rows)
                    throw new Fail("COM_ERROR", "WPS accepted a table inside the cell but inserted nothing (the cell gained " + grew + " paragraph(s), the table object claims "
                        + madeRows + " rows). A nested table does not work through COM on this WPS build; put the table after the cell, or paste it into the cell by hand.");
            }
            tableIndex = RangeIndexOf(doc.Tables, table);
            if (tableIndex == 0) tableIndex = (int)TryGet(() => (object)(int)doc.Tables.Count, 1);
            outRows = (int)TryGet(() => (object)(int)table.Rows.Count, rows);
            outColumns = (int)TryGet(() => (object)(int)table.Columns.Count, columns);
            if (nested) tableIndex = TabOfParagraph(doc, (int)TryGet(() => (object)(int)table.Range.Start, 0));
            applied.Add("rows=" + outRows + " columns=" + outColumns);

            for (int rowI = 1; rowI <= rows; rowI++)
            {
                for (int colI = 1; colI <= columns; colI++)
                {
                    string value = cells[rowI - 1][colI - 1];
                    if (value.Length == 0) continue;
                    try { table.Cell(rowI, colI).Range.Text = value; }
                    catch (Exception e) { throw new Fail("COM_ERROR", "could not write cell " + rowI + "," + colI + ": " + Clean(e.Message)); }
                }
            }

            List<string> appliedBorders = TabBorders(table, borderMode);
            applied.AddRange(appliedBorders);
            borderRead = TabBorderReadback(table);

            try { table.AutoFitBehavior(1); } catch (Exception) { }   // wdAutoFitContent
            if (ratio.Count > 0)
            {
                applied.AddRange(TabWidths(doc, table, ratio));
            }
            else if (columns > 0)
            {
                // No widths given: fill the text column, so the table never runs off the page.
                try { table.PreferredWidthType = PREF_PERCENT; table.PreferredWidth = 100; table.AutoFitBehavior(2); }
                catch (Exception) { }
            }
            widthsRead = TabWidthList(table);

            List<string> fontDone = ApplyCharFormat(table.Range.Font, a);
            applied.AddRange(fontDone);

            if (tableAlign >= 0)
            {
                try { table.Rows.Alignment = tableAlign; applied.Add("align=" + a.Str("align", "")); }
                catch (Exception e) { throw new Fail("COM_ERROR", "could not align the table: " + Clean(e.Message)); }
            }
            if (cellAlign >= 0)
            {
                try { table.Range.ParagraphFormat.Alignment = cellAlign; applied.Add("cellAlign=" + a.Str("cellAlign", "")); }
                catch (Exception) { }
            }
            if (header)
            {
                applied.AddRange(TabHeaderRow(table));
            }
            if (!nested)
            {
                string tidied = TabTidyAfter(doc, table);
                if (tidied.Length > 0) applied.Add(tidied);
            }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["action"] = "insert";
        r["table"] = tableIndex;
        r["rows"] = outRows;
        r["columns"] = outColumns;
        if (nested) r["nested"] = true;
        r["atPara"] = atPara;
        r["borders"] = bordersSpec;
        r["borderRead"] = borderRead;
        r["widths"] = widthsRead;
        r["header"] = header;
        if (captionText.Length > 0)
        {
            r["caption"] = captionText;
            r["captionField"] = captionField;
        }
        r["applied"] = applied;
        r["tables"] = TryGet(() => (object)(int)doc.Tables.Count, 0);
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = "The table is inserted in the open document and nothing was saved; the caption above it is a SEQ field, so it renumbers itself. "
            + "Undo once (Ctrl+Z) takes the whole command back. Check the result with render before saving.";
        return r;
    }

    // ------------------------------------------------------------ table: format

    private static object TabFormat(dynamic doc, Bag a)
    {
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        if (tables == 0)
            throw new Fail("BAD_ARGS", "the document has no table to format; use table with \"data\" to insert one.");

        List<int> targets = new List<int>();
        if (a.Has("table") && a.Int("table", 0) > 0)
        {
            int n = a.Int("table", 0);
            if (n > tables) throw new Fail("BAD_ARGS", "table " + n + " does not exist (the document has " + tables + ").");
            targets.Add(n);
        }
        else
        {
            for (int t = 1; t <= tables; t++) targets.Add(t);
        }

        bool track = a.Flag("track", false);
        List<object> done = new List<object>();

        // Validated for every target before anything is changed.
        string borderMode = a.Has("borders") ? TabBorderMode(a.Str("borders", "")) : "";
        int tableAlign = a.Has("align") ? TabTableAlign(a.Str("align", "center")) : -1;
        int cellAlign = a.Has("cellAlign") ? Alignment(a.Str("cellAlign", "center")) : -1;
        int autoFit = a.Has("autofit") ? TabAutoFitValue(a.Str("autofit", "")) : -1;
        Dictionary<int, List<double>> ratios = new Dictionary<int, List<double>>();
        foreach (int n in targets)
        {
            if (!a.Has("widths")) break;
            int columns = (int)TryGet(() => (object)(int)doc.Tables.Item(n).Columns.Count, 0);
            ratios[n] = TabWidthValues(a.Str("widths", ""), columns);
        }

        using (P2Edit edit = new P2Edit(doc, track, "wps-word-cooperate table format"))
        {
            foreach (int n in targets)
            {
                dynamic table = doc.Tables.Item(n);
                List<string> applied = new List<string>();
                int rows = (int)table.Rows.Count;

                if (borderMode.Length > 0)
                {
                    Tr("format " + n + ": borders");
                    applied.AddRange(TabBorders(table, borderMode));
                    Tr("format " + n + ": borders done");
                }
                if (autoFit >= 0)
                {
                    Tr("format " + n + ": autofit " + autoFit);
                    try { table.AutoFitBehavior(autoFit); }
                    catch (Exception e) { throw new Fail("COM_ERROR", "WPS refused AutoFitBehavior(" + autoFit + ") on table " + n + ": " + Clean(e.Message)); }
                    applied.Add("autofit=" + a.Str("autofit", ""));
                    Tr("format " + n + ": autofit done");
                }
                Tr("format " + n + ": font");
                List<string> fontDone = ApplyCharFormat(table.Range.Font, a);
                applied.AddRange(fontDone);
                Tr("format " + n + ": font done");

                // Paragraph indents/spacing on the whole table range (every cell paragraph).
                // "align" is deliberately excluded: in this command it already means table
                // alignment (Rows.Alignment) and cell alignment is "cellAlign".
                if (a.Has("firstLineIndent") || a.Has("leftIndent") || a.Has("indentChars")
                    || a.Has("lineSpacing") || a.Has("spaceBefore") || a.Has("spaceAfter"))
                {
                    Tr("format " + n + ": paraFormat");
                    applied.AddRange(ApplyParaFormat(table.Range.ParagraphFormat, a.Without("align", "pageBreakBefore", "noSplitWords")));
                    Tr("format " + n + ": paraFormat done");
                }

                if (tableAlign >= 0)
                {
                    Tr("format " + n + ": align " + tableAlign);
                    try { table.Rows.Alignment = tableAlign; applied.Add("align=" + a.Str("align", "")); }
                    catch (Exception e) { throw new Fail("COM_ERROR", "could not align table " + n + ": " + Clean(e.Message)); }
                    Tr("format " + n + ": align done");
                }
                if (cellAlign >= 0)
                {
                    try { table.Range.ParagraphFormat.Alignment = cellAlign; applied.Add("cellAlign=" + a.Str("cellAlign", "")); }
                    catch (Exception) { }
                }
                if (a.Has("widths"))
                {
                    Tr("format " + n + ": widths");
                    applied.AddRange(TabWidths(doc, table, ratios[n]));
                    Tr("format " + n + ": widths done");
                }
                if (a.Has("rowHeight"))
                {
                    Tr("format " + n + ": rowHeight");
                    applied.AddRange(TabRowHeight(table, a));
                    Tr("format " + n + ": rowHeight done");
                }
                if (a.Flag("header", false))
                {
                    Tr("format " + n + ": header");
                    applied.AddRange(TabHeaderRow(table));
                    Tr("format " + n + ": header done");
                }
                if (a.Flag("keepTogether", false))
                {
                    Tr("format " + n + ": keepTogether");
                    applied.AddRange(TabKeepTogether(table));
                    Tr("format " + n + ": keepTogether done");
                }
                if (applied.Count == 0)
                    applied.Add("(no formatting field given)");

                Tr("format " + n + ": readback");
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["table"] = n;
                one["rows"] = rows;
                one["columns"] = TryGet(() => (object)(int)table.Columns.Count, 0);
                one["widths"] = TabWidthList(table);
                one["rowHeight"] = TryGet(() => (object)Math.Round(Num((object)table.Rows.Item(1).Height, 0), 1), 0);
                one["align"] = TryGet(() => (object)(int)table.Rows.Alignment, -1);
                one["borderRead"] = TabBorderReadback(table);
                one["applied"] = applied;
                done.Add(one);
            }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["action"] = "format";
        r["tables"] = done;
        r["count"] = done.Count;
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = "Formatting is not a revision, so nothing shows as a tracked change; nothing was saved. "
            + "Undo once (Ctrl+Z) takes the whole command back. Check the result with render before saving.";
        return r;
    }

    // ------------------------------------------------------------ table: the pieces

    /// <summary>
    /// Tables.Add at a collapsed range, with the variants WPS accepts. A table in the body takes the
    /// plain five-argument call; a nested table (the range sits inside a cell) needed a different one.
    /// </summary>
    private static dynamic TabAdd(dynamic doc, dynamic target, int rows, int cols, bool nested, int para)
    {
        List<string> errors = new List<string>();
        try { return doc.Tables.Add(target, rows, cols, 1, 1); }   // wdWord9TableBehavior, wdAutoFitContent
        catch (Exception e) { errors.Add("five arguments: " + Clean(e.Message)); Tr("tables.add 5 args failed"); }

        int start = (int)TryGet(() => (object)(int)target.Start, 0);
        try
        {
            dynamic again = doc.Range(start, start);
            return doc.Tables.Add(again, rows, cols);
        }
        catch (Exception e) { errors.Add("three arguments: " + Clean(e.Message)); Tr("tables.add 3 args failed"); }

        if (nested)
        {
            // A table inside a cell: WPS rejects the plain call, so the cell itself is asked next.
            try
            {
                dynamic cells = doc.Paragraphs.Item(para).Range.Cells;
                if ((int)cells.Count > 0)
                {
                    dynamic cellRange = cells.Item(1).Range;
                    Tr("nested: cell range " + (int)cellRange.Start + ".." + (int)cellRange.End);
                    return doc.Tables.Add(cellRange, rows, cols, 1, 1);
                }
            }
            catch (Exception e) { errors.Add("cell of the paragraph: " + Clean(e.Message)); Tr("tables.add cell of the paragraph failed"); }

            try
            {
                dynamic paragraphRange = doc.Paragraphs.Item(para).Range;
                return doc.Tables.Add(paragraphRange, rows, cols, 1, 1);
            }
            catch (Exception e) { errors.Add("paragraph range: " + Clean(e.Message)); Tr("tables.add paragraph range failed"); }

            dynamic cell = TabCellAt(doc, start);
            if (cell != null)
            {
                try
                {
                    dynamic cellRange = cell.Range;
                    return cellRange.Tables.Add(cellRange, rows, cols, 1, 1);
                }
                catch (Exception e) { errors.Add("cell range: " + Clean(e.Message)); Tr("tables.add cell range failed"); }
            }
            try
            {
                dynamic pick = doc.Range(start, start);
                pick.Select();
                return doc.Tables.Add(doc.Selection.Range, rows, cols, 1, 1);
            }
            catch (Exception e) { errors.Add("selected range: " + Clean(e.Message)); Tr("tables.add by selection failed"); }
        }
        throw new Fail("COM_ERROR", "WPS did not add the table (" + string.Join(" | ", errors.ToArray()) + ").");
    }

    /// <summary>The cell that holds a character position, or null.</summary>
    private static dynamic TabCellAt(dynamic doc, int position)
    {
        try
        {
            int tables = (int)doc.Tables.Count;
            Tr("cell search at " + position + " over " + tables + " table(s)");
            for (int t = 1; t <= tables; t++)
            {
                dynamic table = doc.Tables.Item(t);
                int rows = (int)table.Rows.Count;
                for (int r = 1; r <= rows; r++)
                {
                    dynamic row = table.Rows.Item(r);
                    int cells = (int)row.Cells.Count;
                    for (int c = 1; c <= cells; c++)
                    {
                        dynamic cell = row.Cells.Item(c);
                        int s = (int)cell.Range.Start;
                        int e = (int)cell.Range.End;
                        if (position >= s && position < e)
                        {
                            Tr("cell found " + t + "/" + r + "/" + c + " [" + s + "," + e + ")");
                            return cell;
                        }
                    }
                }
            }
            Tr("cell not found");
        }
        catch (Exception e) { Tr("cell search failed: " + Clean(e.Message)); }
        return null;
    }

    /// <summary>Canonical border mode for a request value: "grid" | "three-line" | "none" ("" = leave alone).</summary>
    private static string TabBorderMode(string spec)
    {
        string s = (spec == null ? "" : spec.Trim().ToLowerInvariant());
        if (s.Length == 0) return "";
        if (s == "none" || s == "\u65e0\u6846\u7ebf" || s == "\u65e0") return "none";
        if (s == "grid" || s == "all" || s == "\u5168\u6846\u7ebf" || s == "\u6240\u6709\u6846\u7ebf") return "grid";
        if (s == "three-line" || s == "threeline" || s == "three_line" || s == "3line"
            || s == "\u4e09\u7ebf\u8868" || s == "\u4e09\u7ebf") return "three-line";
        if (s == "thick" || s == "thick-grid" || s == "thickgrid" || s == "thick_grid"
            || s == "\u7c97\u6846\u7ebf" || s == "\u7c97\u8fb9\u6846" || s == "\u5916\u7c97\u5185\u7ec6") return "thick";
        throw new Fail("BAD_ARGS", "unknown borders \"" + spec + "\" (grid | thick | three-line | none)");
    }

    /// <summary>grid | three-line | none, applied to the table and to the header rule of a three-line table.</summary>
    private static List<string> TabBorders(dynamic table, string mode)
    {
        List<string> done = new List<string>();
        if (mode == null || mode.Length == 0) return done;

        int rows = (int)TryGet(() => (object)(int)table.Rows.Count, 1);
        dynamic borders = table.Borders;

        // Wipe every line first: the table style may have put borders in that direct formatting has to override.
        for (int i = BT_TOP; i >= BT_DIAG_UP; i--) TabBorder(borders, i, LS_NONE, LW_050);
        try { borders.InsideLineStyle = LS_NONE; } catch (Exception) { }
        try { borders.OutsideLineStyle = LS_NONE; } catch (Exception) { }

        if (mode == "none")
        {
            done.Add("borders=none");
            return done;
        }

        if (mode == "grid")
        {
            for (int i = BT_TOP; i >= BT_VERT; i--) TabBorder(borders, i, LS_SINGLE, LW_050);
            done.Add("borders=grid (0.5 pt, outer + inner)");
            return done;
        }

        if (mode == "thick")
        {
            // Chinese-report standard: a bold frame around the table, thin rules inside.
            TabBorder(borders, BT_TOP, LS_SINGLE, LW_150);
            TabBorder(borders, BT_LEFT, LS_SINGLE, LW_150);
            TabBorder(borders, BT_BOTTOM, LS_SINGLE, LW_150);
            TabBorder(borders, BT_RIGHT, LS_SINGLE, LW_150);
            TabBorder(borders, BT_HORZ, LS_SINGLE, LW_050);
            TabBorder(borders, BT_VERT, LS_SINGLE, LW_050);
            try { borders.OutsideLineStyle = LS_SINGLE; } catch (Exception) { }
            try { borders.InsideLineStyle = LS_SINGLE; } catch (Exception) { }
            done.Add("borders=thick (1.5 pt outer frame, 0.5 pt inner rules)");
            return done;
        }
        // three-line: outer top + outer bottom thick, a thin rule under the header row, nothing else.
        TabBorder(borders, BT_TOP, LS_SINGLE, LW_150);
        TabBorder(borders, BT_BOTTOM, LS_SINGLE, LW_150);
        try { TabBorder(table.Rows.Item(1).Borders, BT_TOP, LS_SINGLE, LW_150); } catch (Exception) { }
        try { TabBorder(table.Rows.Item(rows).Borders, BT_BOTTOM, LS_SINGLE, LW_150); } catch (Exception) { }
        if (rows > 1)
        {
            try { TabBorder(table.Rows.Item(1).Borders, BT_BOTTOM, LS_SINGLE, LW_075); } catch (Exception) { }
        }
        done.Add("borders=three-line (1.5 pt top/bottom, 0.75 pt under the header, no other lines)");
        return done;
    }

    private static bool TabBorder(dynamic borders, int which, int lineStyle, int lineWidth)
    {
        try
        {
            dynamic b = borders.Item(which);
            b.LineStyle = lineStyle;
            if (lineStyle != LS_NONE) b.LineWidth = lineWidth;
            return true;
        }
        catch (Exception e) { Tr("  border item " + which + " failed: " + Clean(e.Message)); return false; }
    }

    /// <summary>What the six borders of the table really are now: [lineStyle, lineWidth] each.</summary>
    private static Dictionary<string, object> TabBorderReadback(dynamic table)
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        string[] names = new string[] { "top", "left", "bottom", "right", "insideH", "insideV", "diagDown", "diagUp" };
        int[] which = new int[] { BT_TOP, BT_LEFT, BT_BOTTOM, BT_RIGHT, BT_HORZ, BT_VERT, BT_DIAG_DOWN, BT_DIAG_UP };
        for (int i = 0; i < names.Length; i++)
        {
            List<object> pair = new List<object>();
            try
            {
                dynamic b = table.Borders.Item(which[i]);
                pair.Add((int)b.LineStyle);
                pair.Add(TryGet(() => (object)(int)b.LineWidth, 0));
            }
            catch (Exception) { pair.Add(-1); pair.Add(0); }
            d[names[i]] = pair;
        }
        return d;
    }

    /// <summary>"2,3,2" -> the numbers, checked against the column count of the table.</summary>
    private static List<double> TabWidthValues(string spec, int columns)
    {
        List<double> wanted = new List<double>();
        if (spec == null || spec.Trim().Length == 0)
            throw new Fail("BAD_ARGS", "\"widths\" is empty; give one number per column, e.g. \"2,3,2\".");
        foreach (string part in spec.Split(','))
        {
            double v;
            if (!double.TryParse(part.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out v) || v <= 0)
                throw new Fail("BAD_ARGS", "widths must be positive comma-separated numbers, e.g. \"2,3,2\" (got \"" + spec + "\").");
            wanted.Add(v);
        }
        if (columns > 0 && wanted.Count != columns)
            throw new Fail("BAD_ARGS", "widths has " + wanted.Count + " value(s) but the table has " + columns + " column(s).");
        return wanted;
    }

    /// <summary>Column widths in the given ratio: fit to content first, then split the text column.</summary>
    private static List<string> TabWidths(dynamic doc, dynamic table, List<double> wanted)
    {
        List<string> done = new List<string>();
        int columns = (int)TryGet(() => (object)(int)table.Columns.Count, 0);
        if (columns <= 0) throw new Fail("COM_ERROR", "the table reports no columns, so widths cannot be set.");
        if (wanted.Count != columns)
            throw new Fail("BAD_ARGS", "widths has " + wanted.Count + " value(s) but the table has " + columns + " column(s).");

        string spec = string.Join(",", Array.ConvertAll<double, string>(wanted.ToArray(), delegate(double v) { return Convert.ToString(v, CultureInfo.InvariantCulture); }));

        double sum = 0;
        foreach (double v in wanted) sum += v;

        try { table.AutoFitBehavior(1); } catch (Exception) { }   // wdAutoFitContent first, as WPS prefers

        double total = TextWidth(doc);

        // Percent of the table width is what both Word and WPS honour most reliably. The table itself
        // is pinned to the text column afterwards: "100 %" alone means the width auto-fit left behind,
        // which came out about 2 % wider than the text column in the first test.
        bool percentOk = false;
        try
        {
            table.PreferredWidthType = PREF_PERCENT;
            table.PreferredWidth = 100;
            for (int c = 1; c <= columns; c++)
            {
                dynamic col = table.Columns.Item(c);
                col.PreferredWidthType = PREF_PERCENT;
                col.PreferredWidth = (float)Math.Round(wanted[c - 1] / sum * 100.0, 2);
            }
            table.PreferredWidthType = PREF_POINTS;
            table.PreferredWidth = (float)Math.Round(total, 1);
            percentOk = true;
        }
        catch (Exception) { percentOk = false; }

        if (percentOk)
        {
            List<object> read = TabWidthList(table);
            // WPS sometimes drops percentage widths: fall back to points when the columns came out wrong.
            if (WidthsLookRight(read, wanted, sum) && WidthsFitText(read, total))
            {
                done.Add("widths=" + spec + " -> " + WidthsText(read));
                return done;
            }
        }

        try
        {
            table.PreferredWidthType = PREF_POINTS;
            table.PreferredWidth = (float)Math.Round(total, 1);
        }
        catch (Exception) { }
        for (int c = 1; c <= columns; c++)
        {
            double one = Math.Round(total * wanted[c - 1] / sum, 1);
            try
            {
                dynamic col = table.Columns.Item(c);
                col.PreferredWidthType = PREF_POINTS;
                col.PreferredWidth = (float)one;
                col.Width = (float)one;
            }
            catch (Exception) { }
        }
        done.Add("widths=" + spec + " -> " + WidthsText(TabWidthList(table)) + " (points)");
        return done;
    }

    private static List<object> TabWidthList(dynamic table)
    {
        List<object> list = new List<object>();
        int columns = (int)TryGet(() => (object)(int)table.Columns.Count, 0);
        for (int c = 1; c <= columns; c++)
        {
            double w = 0;
            try { w = Math.Round(Num((object)table.Columns.Item(c).Width, 0), 1); } catch (Exception) { }
            list.Add(w);
        }
        return list;
    }

    private static string WidthsText(List<object> widths)
    {
        List<string> parts = new List<string>();
        foreach (object w in widths)
            parts.Add(Convert.ToString(w, CultureInfo.InvariantCulture) + "pt");
        return string.Join(" / ", parts.ToArray());
    }

    private static bool WidthsLookRight(List<object> widths, List<double> wanted, double sum)
    {
        if (widths.Count != wanted.Count) return false;
        double total = 0;
        foreach (object w in widths) total += Num(w, 0);
        if (total <= 0) return false;
        for (int i = 0; i < widths.Count; i++)
        {
            double got = Num(widths[i], 0) / total;
            double want = wanted[i] / sum;
            if (Math.Abs(got - want) > 0.10) return false;
        }
        return true;
    }

    /// <summary>True when the columns together are the width of the text column, not wider (they would run into the margin).</summary>
    private static bool WidthsFitText(List<object> widths, double textWidth)
    {
        double total = 0;
        foreach (object w in widths) total += Num(w, 0);
        return Math.Abs(total - textWidth) <= Math.Max(6, textWidth * 0.02);
    }

    /// <summary>wdAutoFit value for window | content | fixed.</summary>
    private static int TabAutoFitValue(string spec)
    {
        string mode = (spec == null ? "" : spec.Trim().ToLowerInvariant());
        if (mode == "window" || mode == "page" || mode == "\u7a97\u53e3" || mode == "\u9875\u9762") return 2;
        if (mode == "content" || mode == "auto" || mode == "\u5185\u5bb9") return 1;
        if (mode == "fixed" || mode == "\u56fa\u5b9a") return 0;
        throw new Fail("BAD_ARGS", "unknown autofit \"" + spec + "\" (window | content | fixed)");
    }

    /// <summary>wdAlignRow value for an "align" on a table; only left/center/right mean anything there.</summary>
    private static int TabTableAlign(string spec)
    {
        int align = Alignment(spec);
        if (align < 0 || align > 2)
            throw new Fail("BAD_ARGS", "a table can only be aligned left|center|right, not \"" + spec + "\" (use cellAlign for the text inside the cells).");
        return align;
    }

    /// <summary>
    /// Word leaves the paragraph the table was built in behind as an empty line under the table.
    /// Removed here, unless it is the last paragraph of the body (the body has to end with one).
    /// </summary>
    private static string TabTidyAfter(dynamic doc, dynamic table)
    {
        try
        {
            int tableEnd = (int)table.Range.End;
            int total = (int)doc.Paragraphs.Count;
            int next = ParaIndexAtOrAfter(doc, tableEnd);
            if (next < 1 || next > total || next >= total) return "";
            string text = ParaTextAt(doc, next);
            if (text.Trim().Length > 0) return "";
            if (TabOfParagraph(doc, next) > 0) return "";
            doc.Paragraphs.Item(next).Range.Delete();
            return "removed the empty paragraph the table left below it";
        }
        catch (Exception) { return ""; }
    }

    private static List<string> TabRowHeight(dynamic table, Bag a)
    {
        List<string> done = new List<string>();
        double height = Points(a.Str("rowHeight", "18pt"));
        string ruleName = a.Str("rowHeightRule", "atLeast").Trim().ToLowerInvariant();
        int rule;
        if (ruleName == "exact" || ruleName == "exactly" || ruleName == "fixed" || ruleName == "\u56fa\u5b9a\u503c") rule = ROW_EXACTLY;
        else if (ruleName == "auto" || ruleName == "\u81ea\u52a8") rule = 0;
        else rule = ROW_ATLEAST;
        int rows = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
        int ok = 0;
        for (int r = 1; r <= rows; r++)
        {
            try
            {
                dynamic row = table.Rows.Item(r);
                row.Height = (float)Math.Round(height, 1);
                row.HeightRule = rule;
                ok++;
            }
            catch (Exception) { }
        }
        done.Add("rowHeight=" + Math.Round(height, 1) + "pt (" + (rule == ROW_EXACTLY ? "exactly" : rule == ROW_ATLEAST ? "at least" : "auto") + ") on " + ok + " row(s)");
        return done;
    }

    private static List<string> TabHeaderRow(dynamic table)
    {
        List<string> done = new List<string>();
        int rows = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
        if (rows < 1) return done;
        try
        {
            dynamic first = table.Rows.Item(1);
            first.HeadingFormat = -1;                       // repeat on every page
            first.Range.Font.Bold = 1;
            first.Range.ParagraphFormat.Alignment = ALIGN_ROW_CENTER;
            done.Add("header row 1 repeated on each page, bold and centred");
        }
        catch (Exception e) { done.Add("header row: " + Clean(e.Message)); }
        return done;
    }

    private static List<string> TabKeepTogether(dynamic table)
    {
        List<string> done = new List<string>();
        int rows = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
        int kept = 0;
        for (int r = 1; r <= rows; r++)
        {
            try
            {
                dynamic row = table.Rows.Item(r);
                row.AllowBreakAcrossPages = 0;
                kept++;
            }
            catch (Exception) { }
            if (r < rows)
            {
                try { table.Rows.Item(r).Range.ParagraphFormat.KeepWithNext = -1; } catch (Exception) { }
            }
        }
        done.Add("keepTogether: " + kept + " row(s) cannot break across pages, and every row but the last keeps with the next");
        return done;
    }

    // ============================================================ caption

    /// <summary>
    /// caption — a numbered caption for a table or a figure, as a SEQ field (never typed digits).
    ///   {"para":N,"label":"\u8868|\u56fe","title":"...","above":true,"align":"center"}
    /// Returns {"para","text","field","label"} plus where it landed and how many fields were refreshed.
    /// </summary>
    private static object CaptionDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no caption was added.");

        int para = a.Int("para", 0);
        if (para < 1) throw new Fail("BAD_ARGS", "\"para\" is required: the paragraph the caption belongs to (use read to see the numbers).");
        int total = (int)doc.Paragraphs.Count;
        if (para > total)
            throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not exist (the document has " + total + ").");

        string label = CapLabel(a.Str("label", "\u8868"));
        string title = a.Str("title", "").Trim();
        bool above = a.Flag("above", label == "\u8868");   // 表 above, 图 below, as Chinese reports do
        string align = a.Str("align", "center");
        bool track = a.Flag("track", false);

        int anchor = CapAnchor(doc, para, above);

        Dictionary<string, object> caption;
        using (P2Edit edit = new P2Edit(doc, track, "wps-word-cooperate caption"))
        {
            caption = AddCaption(doc, anchor, label, title, above, align, a);
        }

        int renewed = 0;
        // A brand-new SEQ field shows its number at once, but earlier captions may need renumbering
        // once this one exists, so every field with the same identifier is refreshed.
        renewed = CapRenumber(doc, label);

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["para"] = ToInt(caption, "para");
        r["label"] = label;
        r["title"] = title;
        r["text"] = Convert.ToString(caption["text"], CultureInfo.InvariantCulture);
        r["field"] = Convert.ToString(caption["field"], CultureInfo.InvariantCulture);
        r["number"] = Convert.ToString(caption["number"], CultureInfo.InvariantCulture);
        r["above"] = above;
        r["align"] = align;
        r["anchorPara"] = anchor;
        r["fieldsRefreshed"] = renewed;
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = "The number after " + label + " is a SEQ field (type 12), not typed text, so it renumbers when captions change; "
            + "select all and press F9 in WPS if a number looks stale. Nothing was saved; undo once (Ctrl+Z) takes it back.";
        return r;
    }

    /// <summary>Normalised label: 表 / 图, also accepting table/tab/fig/figure.</summary>
    private static string CapLabel(string raw)
    {
        string s = (raw == null ? "" : raw.Trim());
        if (s.Length == 0) return "\u8868";
        string low = s.ToLowerInvariant();
        if (low == "table" || low == "tab" || low == "tbl") return "\u8868";
        if (low == "figure" || low == "fig" || low == "image" || low == "picture") return "\u56fe";
        return s;
    }

    /// <summary>
    /// The paragraph to attach the caption to. A paragraph inside a table is resolved to the table
    /// itself (before it for a table caption, after it for a figure caption), because inserting a
    /// paragraph into a cell would put the caption inside the table.
    /// Returns 0 when the caption has to go in front of everything (a table that starts the body).
    /// </summary>
    private static int CapAnchor(dynamic doc, int para, bool above)
    {
        int tableNo = TabOfParagraph(doc, para);
        if (tableNo <= 0) return para;

        int tableStart = (int)TryGet(() => (object)(int)doc.Tables.Item(tableNo).Range.Start, 0);
        int tableEnd = (int)TryGet(() => (object)(int)doc.Tables.Item(tableNo).Range.End, 0);
        int first = ParaIndexAtOrAfter(doc, tableStart);
        int after = ParaIndexAtOrAfter(doc, tableEnd);

        if (above)
        {
            if (first > 1) return first - 1;   // the body paragraph just above the table
            return 0;                          // the table starts the body: put the caption in front of it
        }
        int last = Math.Max(first, after - 1);
        return last;
    }

    /// <summary>
    /// Writes "<label> <SEQ field> <title>" into its own paragraph next to paragraph `para`
    /// (above: in front of it, otherwise behind it) and formats it. Returns the new paragraph index,
    /// the visible text, the field code and the number the field shows.
    /// </summary>
    private static Dictionary<string, object> AddCaption(dynamic doc, int para, string label, string title, bool above, string align, Bag fmt)
    {
        int index;
        if (para <= 0)
        {
            // In front of everything: ask the body for a paragraph, which lands above a leading table.
            try { doc.Content.InsertParagraphBefore(); }
            catch (Exception e) { throw new Fail("COM_ERROR", "could not add a paragraph at the start of the document: " + Clean(e.Message)); }
            index = 1;
        }
        else
        {
            dynamic neighbour = doc.Paragraphs.Item(para);
            if (above) neighbour.Range.InsertParagraphBefore();
            else neighbour.Range.InsertParagraphAfter();
            index = above ? para : para + 1;
        }
        if (index > (int)doc.Paragraphs.Count) index = (int)doc.Paragraphs.Count;

        dynamic cap = doc.Paragraphs.Item(index);
        int start = (int)cap.Range.Start;
        int end = (int)cap.Range.End;
        int contentEnd = Math.Max(start, end - 1);          // the paragraph mark itself is at end-1
        if (contentEnd > start) doc.Range(start, contentEnd).Delete();

        // The title is written first and the field is then put in front of it: text inserted at the
        // end of a field's result lands *inside* the field, and updating the field would wipe it
        // (measured: "图 1 青山站断面图" lost its title as soon as F9-style renumbering ran).
        string prefix = label + " ";
        string suffix = title.Length > 0 ? " " + title : "";
        dynamic body = doc.Range(start, contentEnd);
        body.Text = prefix + suffix;
        int pos = start + prefix.Length;

        string code = label + " \\* ARABIC";
        string number = "";
        try
        {
            dynamic at = doc.Range(pos, pos);
            dynamic field = at.Fields.Add(at, SEQ_FIELD, code);
            try { field.Update(); } catch (Exception) { }
            code = Convert.ToString(TryGet(() => (object)field.Code.Text, code), CultureInfo.InvariantCulture);
            number = Convert.ToString(TryGet(() => (object)CleanPara(Convert.ToString(field.Result.Text, CultureInfo.InvariantCulture)), ""), CultureInfo.InvariantCulture).Trim();
        }
        catch (Exception e)
        {
            throw new Fail("COM_ERROR", "could not insert the SEQ field for the caption: " + Clean(e.Message));
        }

        dynamic cap2 = doc.Paragraphs.Item(index);
        if (align != null && align.Trim().Length > 0)
        {
            int a1 = Alignment(align);
            if (a1 >= 0) { try { cap2.Range.ParagraphFormat.Alignment = a1; } catch (Exception) { } }
        }
        if (fmt != null)
        {
            try { ApplyCharFormat(cap2.Range.Font, fmt); } catch (Exception) { }
        }
        // A caption above its table must not be left alone at the foot of a page (图/表题不落单).
        if (above) { try { cap2.Range.ParagraphFormat.KeepWithNext = -1; } catch (Exception) { } }

        Dictionary<string, object> d = new Dictionary<string, object>();
        d["para"] = index;
        d["label"] = label;
        d["title"] = title;
        d["text"] = ParaTextAt(doc, index);
        d["field"] = code;
        d["number"] = number;
        return d;
    }

    /// <summary>Updates every SEQ field that belongs to this label so the numbering follows the new caption.</summary>
    private static int CapRenumber(dynamic doc, string label)
    {
        int done = 0;
        try
        {
            dynamic fields = doc.Fields;
            int total = (int)fields.Count;
            for (int i = 1; i <= total; i++)
            {
                dynamic field = null;
                try { field = fields.Item(i); } catch (Exception) { continue; }
                int type = (int)TryGet(() => (object)(int)field.Type, -1);
                if (type != SEQ_FIELD) continue;
                string code = Convert.ToString(TryGet(() => (object)field.Code.Text, ""), CultureInfo.InvariantCulture);
                if (code == null || code.IndexOf(label, StringComparison.Ordinal) < 0) continue;
                try { field.Update(); done++; } catch (Exception) { }
            }
        }
        catch (Exception) { }
        return done;
    }

    // ============================================================ picture

    /// <summary>
    /// picture — insert an inline picture (optionally with a caption below it), or resize / crop one.
    ///   insert : {"path":"...","para":N,"width":"12cm","caption":"...","align":"center"}
    ///   adjust : {"image":1,"width":"12cm","height":"8cm","cropLeft":10,...}   crops are percent 0-100
    /// Returns {"action","image","width","height","caption"} with sizes in points.
    /// </summary>
    private static object PictureDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no picture was changed.");

        string action = a.Str("action", a.Has("path") ? "insert" : "adjust").Trim().ToLowerInvariant();
        if (action == "insert" || action == "add") return PicInsert(doc, a);
        if (action == "adjust" || action == "resize" || action == "crop") return PicAdjust(doc, a);
        throw new Fail("BAD_ARGS", "unknown picture action \"" + action + "\" (insert | adjust)");
    }

    private static object PicInsert(dynamic doc, Bag a)
    {
        string path = a.Need("path");
        string full = FullPathOf(path);
        if (!File.Exists(full)) throw new Fail("BAD_ARGS", "the picture does not exist: " + full);

        int total = (int)doc.Paragraphs.Count;
        int at = total;
        if (a.Has("para"))
        {
            at = a.Int("para", 0);
            if (at < 1 || at > total)
                throw new Fail("ANCHOR_MISSING", "paragraph " + at + " does not exist (the document has " + total + ").");
        }
        string expect = a.Str("expect", "");
        if (expect.Length > 0)
        {
            string has = ParaTextAt(doc, at);
            if (!has.StartsWith(expect, StringComparison.Ordinal))
                throw new Fail("ANCHOR_MISSING", "paragraph " + at + " starts with \"" + Head(has, 40) + "\", not \"" + Head(expect, 40) + "\".");
        }

        bool track = a.Flag("track", false);
        string caption = a.Str("caption", null);
        string align = a.Str("align", "center");
        bool clamped = false;
        double naturalWidthPt = 0;
        double scaleToNatural = 1;
        double width = 0, height = 0;
        int image = 0;
        int atPara = at;
        string captionText = "";

        using (P2Edit edit = new P2Edit(doc, track, "wps-word-cooperate picture"))
        {
            // An empty paragraph is used as it is; otherwise the picture gets a paragraph of its own.
            int host = at;
            string text = ParaTextAt(doc, at);
            if (text.Trim().Length > 0)
            {
                try
                {
                    dynamic p = doc.Paragraphs.Item(at);
                    p.Range.InsertParagraphAfter();
                    host = at + 1;
                }
                catch (Exception e) { throw new Fail("COM_ERROR", "could not add a paragraph for the picture: " + Clean(e.Message)); }
            }
            atPara = host;

            dynamic holder = doc.Paragraphs.Item(host);
            int start = (int)holder.Range.Start;
            dynamic shape = null;
            try { shape = doc.InlineShapes.AddPicture(full, false, true, doc.Range(start, start)); }
            catch (Exception first)
            {
                try
                {
                    dynamic sel = doc.Range(start, start);
                    sel.Select();
                }
                catch (Exception) { }
                try { shape = doc.InlineShapes.AddPicture(full, false, true); }
                catch (Exception second)
                {
                    throw new Fail("COM_ERROR", "WPS did not insert the picture: " + Clean(second.Message) + " (first try: " + Clean(first.Message) + ")");
                }
            }
            if (shape == null) throw new Fail("COM_ERROR", "WPS returned no picture object for " + full);
            image = RangeIndexOf(doc.InlineShapes, shape);
            if (image == 0) image = (int)TryGet(() => (object)(int)doc.InlineShapes.Count, 1);

            double limit = TextWidth(doc);
            double want = a.Has("width") ? Points(a.Str("width", "12cm")) : Num((object)shape.Width, 0);
            if (want <= 0) want = 340;
            if (want > limit) { want = Math.Floor(limit); clamped = true; }

            double ratio = 1;
            double nativeW = Num((object)shape.Width, 0);
            double nativeH = Num((object)shape.Height, 0);
            if (nativeW > 0 && nativeH > 0) ratio = nativeH / nativeW;

            try
            {
                if (a.Has("height") && !a.Has("width"))
                {
                    double h = Points(a.Str("height", "8cm"));
                    if (h > 0) shape.Height = (float)h;
                    width = Num((object)shape.Width, want);
                    height = Num((object)shape.Height, h);
                }
                else
                {
                    double h = want * ratio;
                    shape.LockAspectRatio = -1;      // msoTrue: keep the proportions while resizing
                    shape.Width = (float)want;
                    if (a.Has("height"))
                    {
                        double fixedH = Points(a.Str("height", "8cm"));
                        if (fixedH > 0) { try { shape.LockAspectRatio = 0; } catch (Exception) { } shape.Height = (float)fixedH; }
                    }
                    width = Num((object)shape.Width, want);
                    height = Num((object)shape.Height, h);
                }
            }
            catch (Exception e) { throw new Fail("COM_ERROR", "could not size the picture: " + Clean(e.Message)); }

            int a1 = Alignment(align);
            if (a1 >= 0) { try { doc.Paragraphs.Item(host).Range.ParagraphFormat.Alignment = a1; } catch (Exception) { } }
            try { doc.Paragraphs.Item(host).Format.SpaceBefore = 6f; doc.Paragraphs.Item(host).Format.SpaceAfter = 6f; } catch (Exception) { }

            if (caption != null && caption.Length > 0)
            {
                Dictionary<string, object> cap = AddCaption(doc, host, "\u56fe", caption, false, "center", null);
                captionText = Convert.ToString(cap["text"], CultureInfo.InvariantCulture);
            }

            // How far the picture was shrunk: text drawn small in a big picture becomes unreadable on paper.
            try
            {
                using (System.Drawing.Image img = System.Drawing.Image.FromFile(full))
                    naturalWidthPt = img.Width * 0.75;   // 96 dpi
            }
            catch (Exception) { }
            if (naturalWidthPt > 0) scaleToNatural = width / naturalWidthPt;
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["action"] = "insert";
        r["image"] = image;
        r["width"] = Math.Round(width, 1);
        r["height"] = Math.Round(height, 1);
        r["caption"] = captionText;
        r["path"] = full;
        r["atPara"] = atPara;
        r["clamped"] = clamped;
        if (naturalWidthPt > 0)
        {
            r["naturalWidthPt"] = Math.Round(naturalWidthPt, 1);
            r["scale"] = Math.Round(scaleToNatural, 2);
            if (scaleToNatural < 0.5)
                r["warning"] = "The picture is shown at " + Math.Round(scaleToNatural * 100) + " % of its natural size; text drawn in it prints at about that fraction. Redraw it for the width it will have, or widen it.";
        }
        r["images"] = TryGet(() => (object)(int)doc.InlineShapes.Count, 0);
        r["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = (clamped ? "The picture was made no wider than the text column (" + Math.Round(TextWidth(doc), 1) + " pt), so it cannot run off the page. " : "")
            + (captionText.Length > 0 ? "The caption below it is a SEQ field, so it renumbers itself. " : "")
            + "Nothing was saved; undo once (Ctrl+Z) takes the whole command back. Check it with render before saving.";
        return r;
    }

    private static object PicAdjust(dynamic doc, Bag a)
    {
        int count = (int)TryGet(() => (object)(int)doc.InlineShapes.Count, 0);
        if (count == 0) throw new Fail("BAD_ARGS", "the document has no inline picture to adjust; insert one with picture -Path.");
        int n = a.Int("image", 0);
        if (n < 1) throw new Fail("BAD_ARGS", "\"image\" is required for adjust (the number read reports: 1 for the first inline picture).");
        if (n > count) throw new Fail("BAD_ARGS", "picture " + n + " does not exist (the document has " + count + " inline picture(s)).");

        bool track = a.Flag("track", false);
        double width = 0, height = 0;
        double factorX = 1, factorY = 1;
        Dictionary<string, object> crop = new Dictionary<string, object>();

        using (P2Edit edit = new P2Edit(doc, track, "wps-word-cooperate picture adjust"))
        {
            dynamic shape = doc.InlineShapes.Item(n);

            // WPS (like Word) keeps CropLeft/Top/... in points of the picture at 100 % scale, not in
            // percent and not in points of the size the picture is shown at: a 340 pt wide 800 px
            // picture cropped by 170 pt lost only 96 pt of its displayed width (measured on
            // 12.1.0.29161). So the factor between a crop point and a displayed point is measured on
            // this very picture here, and the percentages are converted with it.
            double w0 = Num((object)shape.Width, 0);
            double h0 = Num((object)shape.Height, 0);
            bool anyCrop = a.Has("cropLeft") || a.Has("cropRight") || a.Has("cropTop") || a.Has("cropBottom");
            if (anyCrop)
            {
                try
                {
                    dynamic pf = shape.PictureFormat;
                    factorX = PicCropFactor(shape, true);
                    factorY = PicCropFactor(shape, false);

                    double cL0 = Num((object)pf.CropLeft, 0);
                    double cR0 = Num((object)pf.CropRight, 0);
                    double cT0 = Num((object)pf.CropTop, 0);
                    double cB0 = Num((object)pf.CropBottom, 0);
                    double allW = (factorX > 0 ? w0 / factorX : w0) + cL0 + cR0;   // the picture width in crop points
                    double allH = (factorY > 0 ? h0 / factorY : h0) + cT0 + cB0;

                    double pL = PicCropShare(a, "cropLeft", allW, cL0);
                    double pR = PicCropShare(a, "cropRight", allW, cR0);
                    double pT = PicCropShare(a, "cropTop", allH, cT0);
                    double pB = PicCropShare(a, "cropBottom", allH, cB0);
                    if (pL + pR > 0.98) { double f = 0.98 / (pL + pR); pL *= f; pR *= f; }
                    if (pT + pB > 0.98) { double f = 0.98 / (pT + pB); pT *= f; pB *= f; }

                    double cL = Math.Round(pL * allW, 2), cR = Math.Round(pR * allW, 2);
                    double cT = Math.Round(pT * allH, 2), cB = Math.Round(pB * allH, 2);
                    pf.CropLeft = (float)cL;
                    pf.CropRight = (float)cR;
                    pf.CropTop = (float)cT;
                    pf.CropBottom = (float)cB;

                    // The factor measured above is only an estimate (WPS rounds the crop in points), so
                    // it is corrected with what the width and height really did, and the crops are set
                    // once more with the corrected picture size.
                    double w1 = Num((object)shape.Width, w0);
                    double h1 = Num((object)shape.Height, h0);
                    double dx = (cL + cR) - (cL0 + cR0);
                    if (Math.Abs(dx) > 0.5)
                    {
                        double kx = (w0 - w1) / dx;
                        if (kx > 0.02 && kx < 50)
                        {
                            double allW2 = w1 / kx + cL + cR;
                            if (Math.Abs(allW2 - allW) > 0.5)
                            {
                                factorX = kx;
                                allW = allW2;
                                cL = Math.Round(pL * allW, 2);
                                cR = Math.Round(pR * allW, 2);
                            }
                        }
                    }
                    double dy = (cT + cB) - (cT0 + cB0);
                    if (Math.Abs(dy) > 0.5)
                    {
                        double ky = (h0 - h1) / dy;
                        if (ky > 0.02 && ky < 50)
                        {
                            double allH2 = h1 / ky + cT + cB;
                            if (Math.Abs(allH2 - allH) > 0.5)
                            {
                                factorY = ky;
                                allH = allH2;
                                cT = Math.Round(pT * allH, 2);
                                cB = Math.Round(pB * allH, 2);
                            }
                        }
                    }
                    pf.CropLeft = (float)cL;
                    pf.CropRight = (float)cR;
                    pf.CropTop = (float)cT;
                    pf.CropBottom = (float)cB;

                    crop["left"] = Math.Round(cL, 1);
                    crop["right"] = Math.Round(cR, 1);
                    crop["top"] = Math.Round(cT, 1);
                    crop["bottom"] = Math.Round(cB, 1);
                    crop["visible"] = Math.Round(Math.Max(0, 1.0 - pL - pR) * 100.0, 1) + "% x " + Math.Round(Math.Max(0, 1.0 - pT - pB) * 100.0, 1) + "%";

                    Dictionary<string, object> read = new Dictionary<string, object>();
                    read["left"] = Math.Round(Num((object)pf.CropLeft, 0), 1);
                    read["right"] = Math.Round(Num((object)pf.CropRight, 0), 1);
                    read["top"] = Math.Round(Num((object)pf.CropTop, 0), 1);
                    read["bottom"] = Math.Round(Num((object)pf.CropBottom, 0), 1);
                    read["pictureWidth"] = Math.Round(allW, 1);
                    read["pictureHeight"] = Math.Round(allH, 1);
                    crop["readBack"] = read;
                }
                catch (Exception e) { throw new Fail("COM_ERROR", "WPS refused the crop: " + Clean(e.Message)); }
            }

            // Sizing comes last, on the cropped frame, and both sides are set explicitly: WPS scaling a
            // cropped picture through LockAspectRatio produced a wrong height in testing.
            try
            {
                double visW = Num((object)shape.Width, w0);
                double visH = Num((object)shape.Height, h0);
                bool wantW = a.Has("width");
                bool wantH = a.Has("height");
                if (wantW)
                {
                    width = Points(a.Str("width", ""));
                    height = visW > 0 ? width * visH / visW : visH;
                }
                if (wantH)
                {
                    height = Points(a.Str("height", ""));
                    if (!wantW) width = visH > 0 ? height * visW / visH : visW;
                }
                if (!wantW && !wantH) { width = visW; height = visH; }
                if (wantW || wantH)
                {
                    try { shape.LockAspectRatio = 0; } catch (Exception) { }
                    shape.Width = (float)Math.Round(width, 1);
                    shape.Height = (float)Math.Round(height, 1);
                }
                width = Num((object)shape.Width, width);
                height = Num((object)shape.Height, height);
            }
            catch (Exception e) { throw new Fail("COM_ERROR", "could not size picture " + n + ": " + Clean(e.Message)); }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["action"] = "adjust";
        r["image"] = n;
        r["width"] = Math.Round(width, 1);
        r["height"] = Math.Round(height, 1);
        r["crop"] = crop;
        r["cropFactor"] = new List<object> { Math.Round(factorX, 3), Math.Round(factorY, 3) };
        r["caption"] = PicCaptionOf(doc, n);
        r["images"] = count;
        r["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = "Sizes are in points (1 cm = 28.3465 pt). Crop values were given as percent (0-100) and are stored by WPS in points of the uncropped picture, "
            + "so they are converted with the factor measured on this picture (cropFactor). "
            + "Nothing was saved; undo once (Ctrl+Z) takes the whole command back. Check it with render before saving.";
        return r;
    }

    /// <summary>
    /// How many points of displayed size one crop point takes off, measured on the picture itself:
    /// WPS stores the crop at the 100 % scale of the image, not at the size it is shown at.
    /// </summary>
    private static double PicCropFactor(dynamic shape, bool horizontal)
    {
        try
        {
            dynamic pf = shape.PictureFormat;
            double before, probe;
            if (horizontal)
            {
                double cur = Num((object)pf.CropLeft, 0);
                before = Num((object)shape.Width, 0);
                if (before <= 6) return 1;
                probe = Math.Max(1.0, Math.Min(20.0, before * 0.10));
                pf.CropLeft = (float)(cur + probe);
                double after = Num((object)shape.Width, before);
                pf.CropLeft = (float)cur;
                double k = (before - after) / probe;
                return (k > 0.02 && k < 50) ? k : 1;
            }
            else
            {
                double cur = Num((object)pf.CropTop, 0);
                before = Num((object)shape.Height, 0);
                if (before <= 6) return 1;
                probe = Math.Max(1.0, Math.Min(20.0, before * 0.10));
                pf.CropTop = (float)(cur + probe);
                double after = Num((object)shape.Height, before);
                pf.CropTop = (float)cur;
                double k = (before - after) / probe;
                return (k > 0.02 && k < 50) ? k : 1;
            }
        }
        catch (Exception) { return 1; }
    }


    /// <summary>The crop of one side as a fraction of the picture: what the request asked for, or what is there now.</summary>
    private static double PicCropShare(Bag a, string key, double uncroppedSize, double current)
    {
        if (a.Has(key))
        {
            double pct = Math.Max(0, Math.Min(100, a.Int(key, 0)));
            return pct / 100.0;
        }
        if (uncroppedSize <= 0) return 0;
        return Math.Max(0, Math.Min(1, current / uncroppedSize));
    }

    /// <summary>The caption paragraph right below an inline picture, if it looks like one.</summary>
    private static string PicCaptionOf(dynamic doc, int image)
    {
        try
        {
            dynamic shape = doc.InlineShapes.Item(image);
            int pos = (int)shape.Range.End;
            int para = ParaIndexAtOrAfter(doc, pos);
            for (int p = para; p <= Math.Min(para + 1, (int)doc.Paragraphs.Count); p++)
            {
                string text = ParaTextAt(doc, p).Trim();
                if (text.Length == 0) continue;
                if (text.StartsWith("\u56fe", StringComparison.Ordinal) || text.StartsWith("Figure", StringComparison.OrdinalIgnoreCase))
                    return text;
                return "";
            }
        }
        catch (Exception) { }
        return "";
    }
}
