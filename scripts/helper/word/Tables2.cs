// Table actions added from field notes (2026-10-25): scan / header-repeat / caption
// (move the caption out of the first row) / rows (delete one row).
//
// Two WPS quirks every function here works around:
//   * table.Rows.Item(n).Delete() / .HeadingFormat / .Range / .Cells throw 0x800A1767
//     (the Row object is half-implemented) - so rows are always reached by SELECTING
//     them and going through app.Selection.Rows.
//   * InsertParagraphAfter on the paragraph that sits right before a table lands
//     INSIDE the table's first cell - so a paragraph before a table is created by
//     inserting "\r" + text at (prevPara.Range.End - 1) instead.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

internal static partial class Program
{
    private const int WD_ROW = 10;        // wdRow
    private const int WD_PAGE = 3;        // wdActiveEndPageNumber
    private const int WD_UNDO_RECORD = 0; // unused marker for readability

    // ------------------------------------------------------------ scan (read only)

    /// <summary>
    /// table -Action scan: one line per table (row/column counts, page span, whether the
    /// caption was made into row 1, header repetition, border weights). Read-only: it can
    /// run on a read-only document, takes no undo record and never saves.
    /// </summary>
    private static object TabScan(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        int total = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        List<int> targets = new List<int>();
        int only = a.Int("table", 0);
        if (only > 0)
        {
            if (only > total) throw new Fail("BAD_ARGS", "table " + only + " does not exist (the document has " + total + ").");
            targets.Add(only);
        }
        else for (int n = 1; n <= total; n++) targets.Add(n);

        dynamic app = App(true, false);
        int selStart = -1, selEnd = -1;
        try { selStart = (int)app.Selection.Start; selEnd = (int)app.Selection.End; } catch (Exception) { }
        try { doc.Repaginate(); } catch (Exception) { }

        List<object> rows = new List<object>();
        int multipage = 0, multipageMissing = 0, captionRows = 0, noThick = 0;
        try
        {
            foreach (int n in targets)
            {
                Tr("scan table " + n);
                dynamic table = doc.Tables.Item(n);
                int tRows = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
                int cols = (int)TryGet(() => (object)(int)table.Columns.Count, 0);
                int[] span = PageSpanOf(doc, table);
                bool split = span[1] > span[0] && span[0] > 0;
                int[] cells = RowCellCounts(table, 2);
                int guess = (cols > 0 && cells.Length > 1 && cells[1] > 0 && cells[1] < cols) ? 2 : 1;
                string first = CleanPara(Convert.ToString(TryGet(() => (object)table.Cell(1, 1).Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
                bool cap = IsCaptionText(first);
                int heading = ReadHeadingFormat(app, table);
                Dictionary<string, object> bd = TabBorderReadback(table);
                int outer = PairWidth(bd, "top");
                int inner = PairWidth(bd, "insideH");
                bool thick = outer >= 12 && inner <= 4;

                if (split) multipage++;
                if (split && heading != -1) multipageMissing++;
                if (cap) captionRows++;
                if (!thick) noThick++;

                Dictionary<string, object> one = new Dictionary<string, object>();
                one["table"] = n;
                one["rows"] = tRows;
                one["columns"] = cols;
                one["pageFrom"] = span[0];
                one["pageTo"] = span[1];
                one["split"] = split;
                one["row1Cells"] = cells.Length > 1 ? cells[1] : 0;
                one["row2Cells"] = cells.Length > 2 ? cells[2] : 0;
                one["headerRowsGuess"] = guess;
                one["firstRowText"] = Head(first, 40);
                one["captionInFirstRow"] = cap;
                one["headingFormat"] = heading;
                one["headerRepeated"] = heading == -1;
                one["outerWidth"] = outer;
                one["innerWidth"] = inner;
                one["thickBorders"] = thick;
                rows.Add(one);
            }
        }
        finally { RestoreSelection(doc, selStart, selEnd); }

        Dictionary<string, object> summary = new Dictionary<string, object>();
        summary["tables"] = targets.Count;
        summary["multipage"] = multipage;
        summary["multipageWithoutHeaderRepeat"] = multipageMissing;
        summary["captionInFirstRow"] = captionRows;
        summary["bordersNotThick"] = noThick;

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["action"] = "scan";
        r["summary"] = summary;
        r["tables"] = rows;
        r["note"] = "read-only. headerRowsGuess=2 means row 1 has fewer cells than columns (a merged two-level header) - check that row 2 is really a header before repeating it.";
        return r;
    }

    // ------------------------------------------------------------ header repeat

    /// <summary>
    /// table -Action header-repeat: set "repeat header rows on every page" on the tables that
    /// straddle a page break. Iterates until stable, because repeating a header pushes pages
    /// and can make a neighbouring table straddle a page break too (measured: 22 -> 6 -> 0).
    /// </summary>
    private static object TabHeaderRepeat(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no header was set.");
        int total = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        if (total == 0) throw new Fail("BAD_ARGS", "the document has no table.");

        int passes = a.Int("passes", 3);
        if (passes < 1) passes = 1;
        if (passes > 6) passes = 6;
        int only = a.Int("table", 0);
        int rowOverride = a.Int("row", 0);
        if (only > 0 && only > total) throw new Fail("BAD_ARGS", "table " + only + " does not exist (the document has " + total + ").");

        dynamic app = App(true, true);
        int selStart = -1, selEnd = -1;
        try { selStart = (int)app.Selection.Start; selEnd = (int)app.Selection.End; } catch (Exception) { }

        List<object> log = new List<object>();
        int changed = 0, passesUsed = 0;
        try
        {
            using (P2Edit edit = new P2Edit(doc, false, "wps-word-cooperate header repeat"))
            {
                for (int pass = 1; pass <= passes; pass++)
                {
                    passesUsed = pass;
                    try { doc.Repaginate(); } catch (Exception) { }

                    List<int> todo = new List<int>();
                    List<string> what = new List<string>();
                    int from = only > 0 ? only : 1, to = only > 0 ? only : total;
                    for (int n = from; n <= to; n++)
                    {
                        dynamic table = doc.Tables.Item(n);
                        int[] span = PageSpanOf(doc, table);
                        bool split = span[1] > span[0] && span[0] > 0;
                        int heading = ReadHeadingFormat(app, table);
                        // In single-table mode an unsplit table is still fixed when asked.
                        if (heading == -1) continue;
                        if (!split && only == 0) continue;
                        todo.Add(n);
                        what.Add(n.ToString(CultureInfo.InvariantCulture) + (split ? "" : "*"));
                    }

                    Dictionary<string, object> entry = new Dictionary<string, object>();
                    entry["pass"] = pass;
                    entry["candidates"] = todo.Count;
                    entry["tables"] = what;
                    log.Add(entry);
                    if (todo.Count == 0) break;

                    foreach (int n in todo)
                    {
                        dynamic table = doc.Tables.Item(n);
                        int cols = (int)TryGet(() => (object)(int)table.Columns.Count, 0);
                        int[] cells = RowCellCounts(table, 2);
                        int rowsToRepeat = rowOverride > 0 ? rowOverride : ((cols > 0 && cells.Length > 1 && cells[1] > 0 && cells[1] < cols) ? 2 : 1);
                        int tRows = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
                        if (rowsToRepeat > tRows) rowsToRepeat = tRows;
                        bool ok = false;
                        string err = null;
                        try
                        {
                            SelectRows(app, table, rowsToRepeat);
                            app.Selection.Rows.HeadingFormat = -1;
                            ok = ReadHeadingFormat(app, table) == -1;
                        }
                        catch (Exception e) { err = Clean(e.Message); }
                        Dictionary<string, object> one = new Dictionary<string, object>();
                        one["table"] = n;
                        one["headerRows"] = rowsToRepeat;
                        one["columns"] = cols;
                        one["headingFormat"] = ReadHeadingFormat(app, table);
                        one["ok"] = ok;
                        if (err != null) one["error"] = err;
                        log.Add(one);
                        if (ok) changed++;
                    }
                }
            }

            // final verification sweep
            try { doc.Repaginate(); } catch (Exception) { }
            int missing = 0;
            for (int n = 1; n <= total; n++)
            {
                dynamic table = doc.Tables.Item(n);
                int[] span = PageSpanOf(doc, table);
                bool split = span[1] > span[0] && span[0] > 0;
                if (!split) continue;
                if (ReadHeadingFormat(app, table) != -1) missing++;
            }

            Dictionary<string, object> r = new Dictionary<string, object>();
            r["action"] = "header-repeat";
            r["mode"] = only > 0 ? "table" : "document";
            r["passesUsed"] = passesUsed;
            r["changed"] = changed;
            r["log"] = log;
            r["multipageStillMissing"] = missing;
            r["note"] = "format-only change (no revision marks). 0 missing = stable; otherwise raise -Passes. Render the continuation page to confirm the header row is there.";
            return r;
        }
        finally { RestoreSelection(doc, selStart, selEnd); }
    }

    // ------------------------------------------------------------ caption out of row 1

    /// <summary>
    /// table -Action caption: a caption that was typed into the first (merged) row is moved
    /// into a paragraph of its own just above the table, then that row is deleted.
    /// Order matters: insert the paragraph first (deleting the row first would destroy the
    /// caption text), apply borders last (deleting a row re-computes the top border).
    /// </summary>
    private static object TabCaptionMove(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; the caption was not moved.");
        int total = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        if (total == 0) throw new Fail("BAD_ARGS", "the document has no table.");

        int n = a.Int("table", 0);
        bool force = a.Flag("force", false);
        if (n < 1)
        {
            // no table given: use the first one whose first row is a caption
            for (int k = 1; k <= total; k++)
            {
                string t = CleanPara(Convert.ToString(TryGet(() => (object)doc.Tables.Item(k).Cell(1, 1).Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
                if (IsCaptionText(t)) { n = k; break; }
            }
            if (n < 1) throw new Fail("BAD_ARGS", "no table starts with a caption in its first row; pass -Table n to force one (or -Force).");
        }
        if (n > total) throw new Fail("BAD_ARGS", "table " + n + " does not exist (the document has " + total + ").");

        dynamic table = doc.Tables.Item(n);
        int rowsBefore = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
        if (rowsBefore < 2) throw new Fail("BAD_ARGS", "table " + n + " has only " + rowsBefore + " row(s); nothing to move.");

        dynamic firstCell = table.Cell(1, 1);
        string raw = Convert.ToString(TryGet(() => (object)firstCell.Range.Paragraphs.Item(1).Range.Text, ""), CultureInfo.InvariantCulture);
        string caption = CleanPara(raw);
        if (!IsCaptionText(caption) && !force)
            throw new Fail("BAD_ARGS", "the first cell of table " + n + " is not a caption (\"" + Head(caption.Trim(), 40) + "\"); pass -Force to move it anyway.");

        // keep the caption's own look, but only if it looks like a caption at all
        string fontName = Convert.ToString(TryGet(() => (object)firstCell.Range.Font.Name, ""), CultureInfo.InvariantCulture);
        string fontFe = Convert.ToString(TryGet(() => (object)firstCell.Range.Font.NameFarEast, ""), CultureInfo.InvariantCulture);
        double fontSize = Convert.ToDouble(TryGet(() => (object)firstCell.Range.Font.Size, (object)0.0), CultureInfo.InvariantCulture);
        int align = Convert.ToInt32(TryGet(() => (object)(int)firstCell.Range.ParagraphFormat.Alignment, (object)1), CultureInfo.InvariantCulture);

        dynamic app = App(true, true);
        int selStart = -1, selEnd = -1;
        try { selStart = (int)app.Selection.Start; selEnd = (int)app.Selection.End; } catch (Exception) { }

        int newPara = 0, rowsAfter = rowsBefore;
        try
        {
            using (P2Edit edit = new P2Edit(doc, false, "wps-word-cooperate caption move"))
            {
                // 1) the paragraph right before the table
                int tableStart = (int)TryGet(() => (object)(int)table.Range.Start, 0);
                int prevCount = 0;
                dynamic prevPara = null;
                if (tableStart > 1)
                {
                    prevCount = (int)TryGet(() => (object)(int)doc.Range(0, tableStart - 1).Paragraphs.Count, 0);
                    if (prevCount > 0) prevPara = doc.Paragraphs.Item(prevCount);
                }

                if (prevPara != null)
                {
                    // insert AT the paragraph mark (End - 1) so the new paragraph can never
                    // land inside the table - the mistake that InsertParagraphAfter makes.
                    int pos = (int)prevPara.Range.End - 1;
                    doc.Range(pos, pos).InsertAfter("\r" + caption);
                    newPara = prevCount + 1;
                }
                else
                {
                    doc.Paragraphs.Item(1).Range.InsertBefore(caption + "\r");
                    newPara = 1;
                }

                int insideTable = TabOfParagraph(doc, newPara);
                if (insideTable > 0)
                    throw new Fail("COM_ERROR", "the caption paragraph landed inside table " + insideTable + " - the table starts at the very beginning of the document; move the caption by hand.");

                // 2) style the new paragraph like the caption was
                dynamic np = doc.Paragraphs.Item(newPara);
                if (fontName != null && fontName.Length > 0)
                {
                    try { np.Range.Font.Name = fontName; } catch (Exception) { }
                    try { if (fontFe != null && fontFe.Length > 0) np.Range.Font.NameFarEast = fontFe; } catch (Exception) { }
                }
                if (fontSize > 0) { try { np.Range.Font.Size = (float)fontSize; } catch (Exception) { } }
                try
                {
                    np.Format.Alignment = align;
                    np.Format.FirstLineIndent = 0;
                    np.Format.LeftIndent = 0;
                    np.Format.RightIndent = 0;
                }
                catch (Exception) { }

                // 3) delete the row that held the caption
                rowsAfter = DeleteRowViaSelection(app, table, 1);
                if (rowsAfter >= rowsBefore)
                    throw new Fail("COM_ERROR", "WPS did not delete row 1 of table " + n + " (still " + rowsAfter + " row(s)); the caption paragraph was inserted, undo and try the row deletion by hand.");
            }
        }
        finally { RestoreSelection(doc, selStart, selEnd); }

        List<string> applied = new List<string>();
        applied.Add("caption paragraph above the table (paragraph " + newPara + ")");
        applied.Add("deleted row 1: " + rowsBefore + " -> " + rowsAfter + " row(s)");

        // 4) borders last: deleting a row recomputes the top border
        string borderMode = a.Has("borders") ? TabBorderMode(a.Str("borders", "")) : "";
        if (borderMode.Length > 0)
        {
            using (P2Edit edit = new P2Edit(doc, false, "wps-word-cooperate caption borders"))
            {
                applied.AddRange(TabBorders(table, borderMode));
            }
        }

        Dictionary<string, object> rslt = new Dictionary<string, object>();
        rslt["action"] = "caption";
        rslt["table"] = n;
        rslt["caption"] = Head(caption.Trim(), 80);
        rslt["newPara"] = newPara;
        rslt["rowsBefore"] = rowsBefore;
        rslt["rowsAfter"] = rowsAfter;
        rslt["firstRowText"] = Head(CleanPara(Convert.ToString(TryGet(() => (object)doc.Tables.Item(n).Cell(1, 1).Range.Text, ""), CultureInfo.InvariantCulture)).Trim(), 40);
        rslt["borderRead"] = TabBorderReadback(table);
        rslt["applied"] = applied;
        rslt["note"] = "format-only, no revision marks. Render the page to confirm the caption now sits outside the frame and the outer border is thicker.";
        return rslt;
    }

    // ------------------------------------------------------------ delete one row

    private static object TabRowDelete(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no row was deleted.");
        int total = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        if (total == 0) throw new Fail("BAD_ARGS", "the document has no table.");
        int n = a.Int("table", 0);
        if (n < 1) throw new Fail("BAD_ARGS", "rows needs -Table n (see table -Action scan for the numbers).");
        if (n > total) throw new Fail("BAD_ARGS", "table " + n + " does not exist (the document has " + total + ").");
        int row = a.Int("row", 1);
        dynamic app = App(true, true);
        int selStart = -1, selEnd = -1;
        try { selStart = (int)app.Selection.Start; selEnd = (int)app.Selection.End; } catch (Exception) { }

        int before, after;
        string firstAfter = "";
        try
        {
            using (P2Edit edit = new P2Edit(doc, false, "wps-word-cooperate delete row"))
            {
                dynamic table = doc.Tables.Item(n);
                before = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
                if (row < 1 || row > before) throw new Fail("BAD_ARGS", "row " + row + " is out of range (the table has " + before + " row(s)).");
                after = DeleteRowViaSelection(app, table, row);
            }
            firstAfter = Head(CleanPara(Convert.ToString(TryGet(() => (object)doc.Tables.Item(n).Cell(1, 1).Range.Text, ""), CultureInfo.InvariantCulture)).Trim(), 40);
        }
        finally { RestoreSelection(doc, selStart, selEnd); }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["action"] = "rows";
        r["table"] = n;
        r["deleted"] = row;
        r["rowsBefore"] = before;
        r["rowsAfter"] = after;
        r["firstRowText"] = firstAfter;
        r["note"] = "backspace-style deletion: the row is gone, its text is gone too. Read Rows.Count back before saving.";
        return r;
    }

    // ------------------------------------------------------------ shared row helpers

    /// <summary>Select row 1 .. row 1+count-1 of a table through the Selection object.</summary>
    private static void SelectRows(dynamic app, dynamic table, int count)
    {
        table.Cell(1, 1).Range.Select();
        if (count > 1) app.Selection.MoveEnd(WD_ROW, count);
        else app.Selection.MoveEnd(WD_ROW, 1);
    }

    /// <summary>-1 = repeating, 0 = not repeating, -999 = could not be read.</summary>
    private static int ReadHeadingFormat(dynamic app, dynamic table)
    {
        try
        {
            table.Cell(1, 1).Range.Select();
            app.Selection.MoveEnd(WD_ROW, 1);
            object v = app.Selection.Rows.HeadingFormat;
            return Convert.ToInt32(v, CultureInfo.InvariantCulture);
        }
        catch (Exception) { return -999; }
    }

    private static int DeleteRowViaSelection(dynamic app, dynamic table, int row)
    {
        int before = (int)TryGet(() => (object)(int)table.Rows.Count, 0);
        table.Cell(row, 1).Range.Select();
        app.Selection.MoveEnd(WD_ROW, 1);
        app.Selection.Rows.Delete();
        return (int)TryGet(() => (object)(int)table.Rows.Count, before);
    }

    private static void RestoreSelection(dynamic doc, int start, int end)
    {
        if (start < 0) return;
        try
        {
            int docEnd = (int)doc.Content.End;
            if (start <= end && end < docEnd) doc.Range(start, end).Select();
            else if (docEnd > 1) doc.Range(docEnd - 1, docEnd - 1).Select();
        }
        catch (Exception) { }
    }

    /// <summary>First and last page a table occupies (needs a fresh Repaginate).</summary>
    private static int[] PageSpanOf(dynamic doc, dynamic table)
    {
        int start = (int)TryGet(() => (object)(int)table.Range.Start, 0);
        int end = (int)TryGet(() => (object)(int)table.Range.End, 0);
        int last = end > start ? end - 1 : start;
        int p1 = Convert.ToInt32(TryGet(() => (object)doc.Range(start, start).Information(WD_PAGE), (object)0), CultureInfo.InvariantCulture);
        int p2 = Convert.ToInt32(TryGet(() => (object)doc.Range(last, last).Information(WD_PAGE), (object)0), CultureInfo.InvariantCulture);
        if (p2 < p1) p2 = p1;
        return new int[] { p1, p2 };
    }

    /// <summary>cells[row] = number of cells that belong to that row, for rows 1..maxRow.</summary>
    private static int[] RowCellCounts(dynamic table, int maxRow)
    {
        int[] counts = new int[maxRow + 1];
        try
        {
            dynamic cells = table.Range.Cells;
            int count = (int)cells.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic c = cells.Item(i);
                int r = (int)c.RowIndex;
                if (r >= 1 && r <= maxRow) counts[r]++;
                else if (r > maxRow) break;      // cells come in document order
            }
        }
        catch (Exception) { }
        return counts;
    }

    /// <summary>"表1 xxx" / "表 1 xxx" / "附表2 ..." typed into a cell.</summary>
    private static bool IsCaptionText(string text)
    {
        if (text == null) return false;
        string s = text.TrimStart(' ', '\u3000', '\t');
        if (s.Length < 2) return false;
        if (s[0] != '\u8868') return false;            // 表
        int i = 1;
        while (i < s.Length && (s[i] == ' ' || s[i] == '\u3000' || s[i] == '\t')) i++;
        return i < s.Length && char.IsDigit(s[i]);
    }

    private static int PairWidth(Dictionary<string, object> readback, string key)
    {
        if (readback == null || !readback.ContainsKey(key)) return 0;
        List<object> pair = readback[key] as List<object>;
        if (pair == null || pair.Count < 2) return 0;
        return Convert.ToInt32(pair[1], CultureInfo.InvariantCulture);
    }
}
