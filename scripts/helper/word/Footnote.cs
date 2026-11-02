// footnote - native footnotes: list, insert, edit, delete.
//
// Measured on WPS 12.0/12.1.0.29161: Footnotes.Add(<range>, <reference>, <text>) works and the
// note is a real native footnote (numbering is WPS's own, the reference mark is U+0002).
// FootnoteOptions.NumberingRule is NOT exposed by that WPS build, so the response says
// "numberingSupported": false instead of pretending the switch exists.
//
//   {"cmd":"footnote","doc":"...","action":"list"}
//   {"cmd":"footnote","doc":"...","action":"insert","para":12,"expect":"...","text":"...","find":"word","where":"after"}
//   {"cmd":"footnote","doc":"...","action":"edit","index":2,"expect":"...","text":"..."}
//   {"cmd":"footnote","doc":"...","action":"delete","index":2}
//
// WPS reports Footnote.Index relative to Range.Footnotes, not to Document.Footnotes: the index
// this command returns is ALWAYS the document collection's index, and it is re-read after every
// change (inserting can leave the collection unsorted until WPS repaginates).

using System;
using System.Collections.Generic;
using System.Globalization;

internal static partial class Program
{
    private static object FootnoteDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        string action = a.Str("action", "list").Trim().ToLowerInvariant();
        if (action == "add" || action == "create") action = "insert";
        if (action == "remove") action = "delete";
        if (action == "update") action = "edit";

        if (action == "list") return FootnoteList(doc);
        if (action == "numbering") return FootnoteNumbering(doc, a);
        if (action != "insert" && action != "edit" && action != "delete")
            throw new Fail("BAD_ARGS", "unknown action \"" + action + "\" (list|insert|edit|delete|numbering)");
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no footnote was changed.");

        int before = (int)doc.Footnotes.Count;
        Dictionary<string, object> r;
        if (action == "insert") r = FootnoteInsert(doc, a);
        else if (action == "edit") r = FootnoteEdit(doc, a);
        else r = FootnoteDelete(doc, a);
        r["footnotes"] = TryGet(() => (object)(int)doc.Footnotes.Count, before);
        r["saved"] = false;
        return r;
    }

    private static Dictionary<string, object> FootnoteList(dynamic doc)
    {
        List<object> notes = new List<object>();
        int total = (int)doc.Footnotes.Count;
        bool numberingSupported = false;
        int style = 0, start = 1;
        try
        {
            dynamic opts = doc.Sections.Item(1).Range.FootnoteOptions;
            style = (int)opts.NumberStyle;
            start = (int)opts.StartingNumber;
            // The property this skill would like to set is missing from some WPS builds; touching
            // both spellings tells us which one (if any) this build actually answers on.
            object probe = TryGet(() => (object)opts.NumberingRule, null);
            if (probe == null) probe = TryGet(() => (object)doc.Footnotes.NumberingRule, null);
            numberingSupported = probe != null;
        }
        catch (Exception) { }

        int number = start;
        for (int k = 1; k <= total; k++)
        {
            dynamic note = null;
            try { note = doc.Footnotes.Item(k); } catch (Exception) { continue; }
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["i"] = k;
            dynamic reference = null;
            string mark = "";
            int at = 0;
            try
            {
                reference = note.Reference;
                at = (int)reference.Start;
                mark = Convert.ToString(TryGet(() => (object)reference.Text, ""), CultureInfo.InvariantCulture);
            }
            catch (Exception) { }
            bool automatic = mark.IndexOf('\u0002') >= 0;
            one["label"] = automatic ? FootnoteLabel(number, style) : Clean(mark);
            if (automatic) number++;
            one["auto"] = automatic;
            one["para"] = at > 0 ? Convert.ToInt32(TryGet(() => (object)(int)doc.Range(0, at).Paragraphs.Count, (object)0), CultureInfo.InvariantCulture) : 0;
            one["page"] = reference == null ? 0 : Convert.ToInt32(TryGet(() => (object)(int)reference.Information(3), (object)0), CultureInfo.InvariantCulture);
            one["text"] = NoteText(note);
            notes.Add(one);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["action"] = "list";
        r["count"] = total;
        r["notes"] = notes;
        r["numberingSupported"] = numberingSupported;
        r["numberStyle"] = style;
        r["note"] = numberingSupported
            ? "Indices are Document.Footnotes indices (not Range.Footnotes, which WPS numbers differently)."
            : "WPS 12 does not expose FootnoteOptions.NumberingRule (numberingSupported:false): the note numbers "
              + "follow the style already set in the document; change it in WPS if you must.";
        return r;
    }

    /// <summary>
    /// Numbering rule (continuous / restart each section / restart each page). Word exposes it on
    /// both Document.Footnotes and Range.FootnoteOptions; WPS 12 answers on one of the two only,
    /// so both are probed and the response says plainly what this build supports.
    /// </summary>
    private static Dictionary<string, object> FootnoteNumbering(dynamic doc, Bag a)
    {
        string wanted = a.Str("number", a.Str("rule", "")).Trim().ToLowerInvariant();
        int rule;
        if (wanted == "continuous" || wanted == "0") rule = 0;
        else if (wanted == "section" || wanted == "1") rule = 1;
        else if (wanted == "page" || wanted == "2") rule = 2;
        else throw new Fail("BAD_ARGS", "unknown numbering rule \"" + wanted + "\" (continuous|section|page)");

        bool ok = false;
        string via = "";
        string readBack = "";
        try
        {
            doc.Footnotes.NumberingRule = rule;
            readBack = Convert.ToString(doc.Footnotes.NumberingRule, CultureInfo.InvariantCulture);
            ok = readBack.Trim() == Convert.ToString(rule, CultureInfo.InvariantCulture);
            via = "Document.Footnotes";
        }
        catch (Exception) { }
        if (!ok)
        {
            try
            {
                dynamic opts = doc.Sections.Item(1).Range.FootnoteOptions;
                opts.NumberingRule = rule;
                readBack = Convert.ToString(opts.NumberingRule, CultureInfo.InvariantCulture);
                ok = readBack.Trim() == Convert.ToString(rule, CultureInfo.InvariantCulture);
                via = "Range.FootnoteOptions";
            }
            catch (Exception) { }
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["action"] = "numbering";
        r["numberingSupported"] = ok;
        r["rule"] = rule;
        r["via"] = via;
        r["readBack"] = readBack;
        r["note"] = ok
            ? "The numbering rule is now " + wanted + " (whole document; WPS renumbers by itself)."
            : "This WPS build does not expose FootnoteOptions.NumberingRule nor Document.Footnotes.NumberingRule: "
              + "numbering stays as the document already has it. Set it in the WPS user interface if it must change.";
        return r;
    }

    private static string FootnoteLabel(int number, int style)    {
        if (style <= 0 || style == 5) return Convert.ToString(number, CultureInfo.InvariantCulture);
        if (style == 1 || style == 2)
        {
            int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] letters = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            string label = "";
            int n = number;
            for (int k = 0; k < values.Length; k++) while (n >= values[k]) { label += letters[k]; n -= values[k]; }
            return style == 2 ? label.ToLowerInvariant() : label;
        }
        if ((style == 3 || style == 4) && number > 0)
        {
            string label = "";
            int n = number;
            while (n > 0) { n--; label = (char)('A' + n % 26) + label; n /= 26; }
            return style == 4 ? label.ToLowerInvariant() : label;
        }
        return Convert.ToString(number, CultureInfo.InvariantCulture);
    }

    /// <summary>The note body without the reference mark (U+0002) and without the trailing CR.</summary>
    private static dynamic NoteBody(dynamic note)
    {
        dynamic body = note.Range.Duplicate;
        string text = Convert.ToString(TryGet(() => (object)body.Text, ""), CultureInfo.InvariantCulture);
        if (text.StartsWith("\u0002", StringComparison.Ordinal)) body.Start = (int)body.Start + 1;
        text = Convert.ToString(TryGet(() => (object)body.Text, ""), CultureInfo.InvariantCulture);
        if (text.EndsWith("\r", StringComparison.Ordinal)) body.End = (int)body.End - 1;
        return body;
    }

    private static string NoteText(dynamic note)
    {
        try { return Clean(Convert.ToString((object)NoteBody(note).Text, CultureInfo.InvariantCulture)).Trim(' ', '\r', '\n'); }
        catch (Exception) { return ""; }
    }

    private static Dictionary<string, object> FootnoteInsert(dynamic doc, Bag a)
    {
        int para = a.Int("para", 0);
        int total = (int)doc.Paragraphs.Count;
        if (para < 1 || para > total)
            throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not exist (the document has " + total + ").");
        dynamic p = doc.Paragraphs.Item(para);
        string expect = a.Str("expect", "");
        if (expect.Length > 0)
        {
            string has = CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));
            if (!has.StartsWith(expect, StringComparison.Ordinal))
                throw new Fail("ANCHOR_MISSING", "paragraph " + para + " starts with \"" + Head(has, 40) + "\", not \"" + Head(expect, 40) + "\".");
        }
        string text = a.Str("text", null);
        if (text == null) throw new Fail("BAD_ARGS", "insert needs \"text\" (the footnote body)");
        if (text.Trim().Length == 0) throw new Fail("BAD_ARGS", "a footnote must contain text");
        if (text.IndexOf("$", StringComparison.Ordinal) >= 0)
            throw new Fail("BAD_ARGS", "an equation inside a footnote is not supported yet: put the formula in the body text "
                + "with formula, or write it in plain characters (x\u00b2, \u03b1 \u2264 \u03b2)");

        dynamic spot = p.Range.Duplicate;
        spot.End = (int)spot.End - 1;
        string find = a.Str("find", "");
        if (find.Length > 0)
        {
            string where = a.Str("where", "after").Trim().ToLowerInvariant();
            if (where != "before" && where != "after")
                throw new Fail("BAD_ARGS", "unknown where \"" + where + "\" (before|after)");
            string body = CleanPara(Convert.ToString(TryGet(() => (object)spot.Text, ""), CultureInfo.InvariantCulture));
            int idx = IndexOf(body, find, 0);
            if (idx < 0) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not contain \"" + Head(find, 40) + "\".");
            int at = (int)spot.Start + idx + (where == "before" ? 0 : find.Length);
            spot = doc.Range(at, at);
        }
        else
        {
            if (a.Has("where")) throw new Fail("BAD_ARGS", "where needs find: without find the note goes at the end of the paragraph");
            int at = (int)spot.End;
            spot = doc.Range(at, at);
        }

        int before = (int)doc.Footnotes.Count;
        dynamic note = FootnoteAdd(doc, spot, " ");
        FillNote(note, text);

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["action"] = "insert";
        r["para"] = para;
        r["added"] = (int)doc.Footnotes.Count - before;
        r["count"] = TryGet(() => (object)(int)doc.Footnotes.Count, 0);
        r["text"] = text;
        r["note"] = "The footnote reference mark was added to paragraph " + para + ". WPS may hand back a stale "
            + "collection index right after the insert: run footnote -Action list and use that index.";
        return r;
    }

    private static dynamic FootnoteAdd(dynamic doc, dynamic spot, string text)
    {
        // WPS accepts the three-argument COM signature positionally (range, reference, text).
        try { return doc.Footnotes.Add(spot, "", text); }
        catch (Exception) { }
        try { return doc.Footnotes.Add(spot, Type.Missing, text); }
        catch (Exception) { }
        return doc.Footnotes.Add(spot);
    }

    private static void FillNote(dynamic note, string text)
    {
        dynamic body = NoteBody(note);
        body.Text = text;
        // The mark and the CR are native parts of the note; give the new body them back.
        body = NoteBody(note);
        body.Text = text;
    }

    private static dynamic FindNote(dynamic doc, Bag a)
    {
        int index = a.Int("index", 0);
        int total = (int)doc.Footnotes.Count;
        if (index < 1 || index > total)
            throw new Fail("ANCHOR_MISSING", "there is no footnote " + index + " (the document has " + total + "); run footnote -Action list");
        dynamic note = doc.Footnotes.Item(index);
        string expect = a.Str("expect", "");
        if (expect.Length > 0)
        {
            string actual = NoteText(note);
            if (!actual.StartsWith(expect, StringComparison.Ordinal))
                throw new Fail("ANCHOR_MISMATCH", "footnote " + index + " starts with \"" + Head(actual, 40)
                    + "\", not \"" + Head(expect, 40) + "\"; run footnote -Action list");
        }
        return note;
    }

    private static Dictionary<string, object> FootnoteEdit(dynamic doc, Bag a)
    {
        dynamic note = FindNote(doc, a);
        int index = a.Int("index", 0);
        string text = a.Str("text", null);
        if (text == null) throw new Fail("BAD_ARGS", "edit needs \"text\"");
        if (text.Trim().Length == 0) throw new Fail("BAD_ARGS", "a footnote must contain text");
        FillNote(note, text);
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["action"] = "edit";
        r["index"] = index;
        r["text"] = text;
        r["readBack"] = NoteText(doc.Footnotes.Item(index));
        r["note"] = "Retyped in place: the reference mark and the native numbering are untouched.";
        return r;
    }

    private static Dictionary<string, object> FootnoteDelete(dynamic doc, Bag a)
    {
        dynamic note = FindNote(doc, a);
        int index = a.Int("index", 0);
        string was = NoteText(note);
        note.Delete();
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["action"] = "delete";
        r["index"] = index;
        r["deleted"] = was;
        r["note"] = "WPS renumbers the remaining native footnotes by itself; indices shift down by one.";
        return r;
    }
}
