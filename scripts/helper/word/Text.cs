// v0.3 text work: refresh every field, set one table cell, insert paragraphs (clean or tracked).
using System;
using System.Collections.Generic;
using System.Globalization;

internal static partial class Program
{
    /// <summary>
    /// fields — update every field in the document: page numbers, TOC, captions, cross references.
    ///   {"doc":"..."}
    /// Returns {"toc":N,"tocUpdated":[...],"fieldsUpdated":bool,"pages":N}.
    /// </summary>
    private static object FieldsDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; its fields were not updated.");

        List<object> tocs = new List<object>();
        int tocCount = (int)TryGet(() => (object)(int)doc.TablesOfContents.Count, 0);

        bool fieldsOk = true;
        string fieldsError = null;
        try { doc.Fields.Update(); }
        catch (Exception e) { fieldsOk = false; fieldsError = Clean(e.Message); }

        for (int i = 1; i <= tocCount; i++)
        {
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["toc"] = i;
            try
            {
                dynamic toc = doc.TablesOfContents.Item(i);
                toc.Update();
                one["ok"] = true;
            }
            catch (Exception e) { one["ok"] = false; one["error"] = Clean(e.Message); }
            tocs.Add(one);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["fieldsUpdated"] = fieldsOk;
        if (!fieldsOk) r["fieldsError"] = fieldsError;
        r["toc"] = tocCount;
        r["tocUpdated"] = tocs;
        r["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        r["saved"] = false;
        r["note"] = "Field results and the table of contents are refreshed in the open document; nothing was saved.";
        return r;
    }

    /// <summary>
    /// cell — replace the text of one table cell, optionally with formatting.
    ///   {"table":1,"row":2,"col":3,"text":"...","font":"...","size":10.5,"bold":true,"align":"center"}
    ///   {"para":N,"text":"..."}   (para inside a cell; the table/row/col are worked out from it)
    /// Returns {"table","row","col","text","readBack"}.
    /// </summary>
    private static object CellDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; the cell was not changed.");
        string text = a.Str("text", null);
        if (text == null) throw new Fail("BAD_ARGS", "\"text\" is required (pass an empty string to clear the cell)");

        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        if (tables == 0) throw new Fail("BAD_ARGS", "the document has no table");

        dynamic cell = null;
        int tableNo = a.Int("table", 0);
        int row = a.Int("row", 0);
        int col = a.Int("col", 0);

        if (tableNo <= 0)
        {
            int para = a.Int("para", 0);
            if (para <= 0) throw new Fail("BAD_ARGS", "give \"table\" with \"row\" and \"col\", or a \"para\" that sits inside a cell");
            int total = (int)doc.Paragraphs.Count;
            if (para > total) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not exist (the document has " + total + ").");
            dynamic p = doc.Paragraphs.Item(para);
            dynamic inCell = null;
            try { inCell = p.Range.Cells.Item(1); }
            catch (Exception) { inCell = null; }
            if (inCell == null)
                throw new Fail("BAD_ARGS", "paragraph " + para + " is not inside a table cell; read marks those paragraphs with inTable:true, or give table/row/col");
            cell = inCell;
            tableNo = FindTableOfCell(doc, cell);
            row = (int)TryGet(() => (object)(int)cell.RowIndex, 0);
            col = (int)TryGet(() => (object)(int)cell.ColumnIndex, 0);
        }
        else
        {
            if (tableNo > tables) throw new Fail("BAD_ARGS", "table " + tableNo + " does not exist (the document has " + tables + ")");
            if (row <= 0 || col <= 0) throw new Fail("BAD_ARGS", "\"row\" and \"col\" are both required (1-based)");
            dynamic tb = doc.Tables.Item(tableNo);
            int rows = (int)tb.Rows.Count;
            int cols = (int)tb.Columns.Count;
            if (row > rows) throw new Fail("BAD_ARGS", "row " + row + " does not exist (the table has " + rows + ")");
            if (col > cols) throw new Fail("BAD_ARGS", "column " + col + " does not exist (the table has " + cols + ")");
            try { cell = tb.Cell(row, col); }
            catch (Exception e) { throw new Fail("BAD_ARGS", "cell " + row + "," + col + " cannot be used (merged cell?): " + Clean(e.Message)); }
        }

        dynamic app = App(true, true);
        dynamic undo = null;
        try { undo = app.UndoRecord; undo.StartCustomRecord("wps-word-cooperate cell"); } catch (Exception) { undo = null; }
        bool wasTracking = false;
        try { wasTracking = (bool)doc.TrackRevisions; } catch (Exception) { }
        bool track = a.Flag("track", false);
        if (!track) { try { doc.TrackRevisions = false; } catch (Exception) { } }
        try
        {
            cell.Range.Text = text;
            List<string> applied = new List<string>();
            applied.AddRange(ApplyCharFormat(cell.Range.Font, a));
            applied.AddRange(ApplyParaFormat(cell.Range.ParagraphFormat, a));

            Dictionary<string, object> r = new Dictionary<string, object>();
            r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
            r["table"] = tableNo;
            r["row"] = row;
            r["col"] = col;
            r["text"] = text;
            r["applied"] = applied;
            r["readBack"] = CleanPara(Convert.ToString(TryGet(() => (object)cell.Range.Text, ""), CultureInfo.InvariantCulture));
            r["tracked"] = track;
            r["saved"] = false;
            return r;
        }
        finally
        {
            if (!track) { try { doc.TrackRevisions = wasTracking; } catch (Exception) { } }
            if (undo != null) { try { undo.EndCustomRecord(); } catch (Exception) { } }
        }
    }

    private static int FindTableOfCell(dynamic doc, dynamic cell)
    {
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        int want = (int)TryGet(() => (object)(int)cell.Range.Start, -1);
        for (int t = 1; t <= tables; t++)
        {
            int start = (int)TryGet(() => (object)(int)doc.Tables.Item(t).Range.Start, -2);
            int end = (int)TryGet(() => (object)(int)doc.Tables.Item(t).Range.End, -3);
            if (want >= start && want <= end) return t;
        }
        return 0;
    }

    /// <summary>
    /// insert — insert paragraphs (several at once), clean by default, tracked only when asked.
    ///   {"para":N,"where":"after|before|end","items":[{"text":"...","style":"正文","font":"...","size":12}],
    ///    "expect":"...","track":false}
    /// Returns {"inserted":N,"from":i,"to":j,"paragraphs":N}.
    /// </summary>
    private static object InsertDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; nothing was inserted.");
        List<object> items = a.List("items");
        if (items.Count == 0) throw new Fail("BAD_ARGS", "\"items\" is empty; give [{\"text\":\"...\"}, ...]");

        List<string> texts = new List<string>();
        List<Bag> bags = new List<Bag>();
        foreach (object raw in items)
        {
            Dictionary<string, object> item = raw as Dictionary<string, object>;
            if (item == null) throw new Fail("BAD_ARGS", "each item must be an object like {\"text\":\"...\"}");
            Bag ib = new Bag(item);
            string t = ib.Str("text", null);
            if (t == null) throw new Fail("BAD_ARGS", "every item needs \"text\"");
            texts.Add(t);
            bags.Add(ib);
        }

        int total = (int)doc.Paragraphs.Count;
        string where = a.Str("where", "after").Trim().ToLowerInvariant();
        if (where != "after" && where != "before" && where != "end")
            throw new Fail("BAD_ARGS", "unknown where \"" + where + "\" (after|before|end)");

        int at = a.Int("para", 0);
        if (where == "end") at = total;
        if (at < 1 || at > total) throw new Fail("ANCHOR_MISSING", "paragraph " + at + " does not exist (the document has " + total + ").");

        dynamic anchor = doc.Paragraphs.Item(at);
        string expect = a.Str("expect", "");
        if (expect.Length > 0)
        {
            string has = CleanPara(Convert.ToString(TryGet(() => (object)anchor.Range.Text, ""), CultureInfo.InvariantCulture));
            if (!has.StartsWith(expect, StringComparison.Ordinal))
                throw new Fail("ANCHOR_MISSING", "paragraph " + at + " starts with \"" + Head(has, 40) + "\", not \"" + Head(expect, 40) + "\".");
        }

        dynamic app = App(true, true);
        dynamic undo = null;
        try { undo = app.UndoRecord; undo.StartCustomRecord("wps-word-cooperate insert"); } catch (Exception) { undo = null; }
        bool track = a.Flag("track", false);
        bool wasTracking = false;
        try { wasTracking = (bool)doc.TrackRevisions; } catch (Exception) { }
        try { doc.TrackRevisions = track; } catch (Exception) { }

        int from, to;
        try
        {
            if (where == "before") { anchor.Range.InsertParagraphBefore(); from = at; }
            else if (where == "end") { doc.Content.InsertParagraphAfter(); from = (int)doc.Paragraphs.Count; }
            else { anchor.Range.InsertParagraphAfter(); from = at + 1; }
            if (from > (int)doc.Paragraphs.Count) from = (int)doc.Paragraphs.Count;

            string block = string.Join("\r", texts.ToArray());
            doc.Paragraphs.Item(from).Range.InsertBefore(block);
            to = from + texts.Count - 1;

            for (int k = 0; k < bags.Count; k++)
            {
                int index = from + k;
                if (index > (int)doc.Paragraphs.Count) break;
                dynamic p = doc.Paragraphs.Item(index);
                Bag ib = bags[k];
                if (ib.Has("style"))
                {
                    string styleName = ib.Str("style", "");
                    try { p.Style = FindStyle(doc, styleName).NameLocal; }
                    catch (Exception e) { throw new Fail("BAD_ARGS", "cannot apply style \"" + styleName + "\" to the new paragraph: " + Clean(e.Message)); }
                }
                ApplyCharFormat(p.Range.Font, ib);
                ApplyParaFormat(p.Format, ib);
            }
        }
        finally
        {
            if (!track) { try { doc.TrackRevisions = wasTracking; } catch (Exception) { } }
            if (undo != null) { try { undo.EndCustomRecord(); } catch (Exception) { } }
        }

        bool inTable = false;
        try { dynamic p0 = doc.Paragraphs.Item(from); dynamic c0 = p0.Range.Cells.Item(1); inTable = c0 != null; } catch (Exception) { }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["inserted"] = texts.Count;
        r["inTable"] = inTable;
        r["from"] = from;
        r["to"] = to;
        r["where"] = where;
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["revisions"] = TryGet(() => (object)(int)doc.Revisions.Count, 0);
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = (inTable ? "The new paragraphs landed INSIDE a table cell (the anchor paragraph was a cell, or it sat right before a table). " : "")
            + "Paragraph numbers after the insertion moved; read again before the next edit.";
        return r;
    }
}
