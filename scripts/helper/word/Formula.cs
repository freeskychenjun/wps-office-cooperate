// formula - LaTeX into a NATIVE WPS Writer equation.
//
// WPS 12 does not turn linear text into an equation the way Microsoft Word does, and
// Document.OMaths.Add(range) is refused ("value out of expected range"). What does work is
// Range.InsertXML(<Flat OPC package>) carrying <m:oMath> in /word/document.xml - measured on
// WPS 12.0/12.1.0.29161 (OMaths 0 -> 1, readable back through OMaths.Item(1).Range.Text).
//
//   {"cmd":"formula","doc":"...","para":12,"expect":"...","latex":"E=mc^2","display":false}
//   {"cmd":"formula","doc":"...","para":12,"latex":"\\frac{a}{b}","display":true,"number":"1"}
//   {"cmd":"formula","doc":"...","items":[{"para":12,"latex":"x^2"}, ...]}
//
// A number ("-Number 3", or \tag{3} inside the LaTeX) is written as the "equation#(3)" array
// Word itself uses: an <m:eqArr> of one line whose "#" sends what follows to the right margin.
//
// Insertion is verified: if Document.OMaths.Count did not grow, what came in is rolled back and
// the call fails with MATH_FAILED, so a formula can never quietly degrade into plain text.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

internal static partial class Program
{
    private const string FLAT_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string FLAT_NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>The Flat OPC package WPS needs in order to accept an equation at all.</summary>
    private static string FlatDocument(string paragraphXml)
    {
        return "<?xml version=\"1.0\" standalone=\"yes\"?>"
            + "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">"
            + "<pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\">"
            + "<pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships></pkg:xmlData></pkg:part>"
            + "<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\">"
            + "<pkg:xmlData><w:document xmlns:w=\"" + FLAT_NS_W + "\" xmlns:m=\"" + FLAT_NS_M + "\"><w:body><w:p>"
            + paragraphXml
            + "</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
    }

    private static readonly Regex TAG = new Regex(@"\\tag\*?\s*\{([^}]*)\}", RegexOptions.CultureInvariant);

    private static object FormulaDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no equation was inserted.");

        List<Bag> jobs = new List<Bag>();
        List<object> items = a.List("items");
        if (items.Count > 0)
        {
            foreach (object raw in items)
            {
                Dictionary<string, object> item = raw as Dictionary<string, object>;
                if (item == null) throw new Fail("BAD_ARGS", "each item must be an object like {\"para\":12,\"latex\":\"x^2\"}");
                jobs.Add(new Bag(item));
            }
        }
        else jobs.Add(a);
        if (jobs.Count > 12) throw new Fail("BAD_ARGS", "at most 12 formulas per call (got " + jobs.Count + "); split the batch");

        int mathsStart = 0;
        try { mathsStart = (int)doc.OMaths.Count; } catch (Exception) { }
        int paragraphsStart = (int)doc.Paragraphs.Count;

        List<object> done = new List<object>();
        List<object> failed = new List<object>();
        foreach (Bag job in jobs)
        {
            try { done.Add(InsertOneFormula(doc, job)); }
            catch (Fail f) { failed.Add(FormulaFailure(job, f.Code, f.Message)); }
            catch (OmmlFail o) { failed.Add(FormulaFailure(job, "MATH_UNSUPPORTED", o.Message)); }
            catch (Exception e) { failed.Add(FormulaFailure(job, "MATH_FAILED", Clean(e.Message))); }
        }

        int mathsEnd = mathsStart;
        try { mathsEnd = (int)doc.OMaths.Count; } catch (Exception) { }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["inserted"] = done.Count;
        r["failed"] = failed;
        r["equations"] = done;
        r["omaths"] = mathsEnd;
        r["omathsAdded"] = mathsEnd - mathsStart;
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["saved"] = false;
        string note = "Native equations, not pictures: WPS accepts them only through InsertXML of a Flat OPC package. "
            + "Read them back with read -Full / the equations array; paragraph numbers after the insertion moved.";
        if (failed.Count > 0) note += " " + failed.Count + " formula(s) were NOT inserted (see failed[]); nothing was left behind for them.";
        r["note"] = note;
        if (done.Count == 0)
        {
            string code = failed.Count > 0 && Convert.ToString(((Dictionary<string, object>)failed[0])["code"], CultureInfo.InvariantCulture) == "MATH_UNSUPPORTED"
                ? "MATH_UNSUPPORTED" : "MATH_FAILED";
            string msg = failed.Count > 0 ? Convert.ToString(((Dictionary<string, object>)failed[0])["message"], CultureInfo.InvariantCulture) : "no formula was inserted";
            throw new Fail(code, msg);
        }
        return r;
    }

    private static Dictionary<string, object> FormulaFailure(Bag job, string code, string message)
    {
        Dictionary<string, object> f = new Dictionary<string, object>();
        f["code"] = code;
        f["message"] = message ?? "";
        f["latex"] = job.Str("latex", "");
        f["para"] = job.Int("para", 0);
        return f;
    }

    /// <summary>The OMML fragment for one job: inline <m:oMath> or a centred <m:oMathPara>.</summary>
    private static string FormulaXml(Bag job, out bool display, out string latex, out string number)
    {
        latex = job.Need("latex").Trim();
        display = job.Flag("display", false);
        // A "$$...$$" spelling means display, "$...$" means inline; the flags win when given.
        if (latex.StartsWith("$$", StringComparison.Ordinal) && latex.EndsWith("$$", StringComparison.Ordinal) && latex.Length > 4)
        {
            latex = latex.Substring(2, latex.Length - 4).Trim();
            if (!job.Has("display")) display = true;
        }
        else if (latex.Length > 2 && latex[0] == '$' && latex[latex.Length - 1] == '$')
            latex = latex.Substring(1, latex.Length - 2).Trim();

        number = job.Str("number", "");
        Match m = TAG.Match(latex);
        if (m.Success)
        {
            if (number.Length == 0) number = m.Groups[1].Value.Trim();
            latex = latex.Remove(m.Index, m.Length).Trim();
        }
        if (latex.Length == 0) throw new Fail("BAD_ARGS", "\"latex\" is empty");

        string omml = Omml.LatexToOmml(latex);
        if (number.Length > 0)
        {
            if (!display) throw new Fail("BAD_ARGS", "a numbered equation must be display (set \"display\":true)");
            string inner = omml.Substring(9, omml.Length - 19);
            omml = "<m:oMath><m:eqArr><m:eqArrPr><m:maxDist m:val=\"1\"/></m:eqArrPr><m:e>"
                + inner
                + "<m:r><m:t>#</m:t></m:r>"
                + "<m:r><m:rPr><m:nor/></m:rPr><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/></w:rPr><m:t>("
                + Omml.Esc(number) + ")</m:t></m:r>"
                + "</m:e></m:eqArr></m:oMath>";
        }
        if (!display) return omml;
        return "<m:oMathPara><m:oMathParaPr><m:jc m:val=\"center\"/></m:oMathParaPr>" + omml + "</m:oMathPara>";
    }

    private static Dictionary<string, object> InsertOneFormula(dynamic doc, Bag job)
    {
        bool display;
        string latex, number;
        string xml = FormulaXml(job, out display, out latex, out number);

        int para = job.Int("para", 0);
        int total = (int)doc.Paragraphs.Count;
        if (para == 0) para = total;
        if (para < 0 || para > total)
            throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not exist (the document has " + total + ").");

        dynamic host = doc.Paragraphs.Item(para);
        string expect = job.Str("expect", "");
        if (expect.Length > 0)
        {
            string has = CleanPara(Convert.ToString(TryGet(() => (object)host.Range.Text, ""), CultureInfo.InvariantCulture));
            if (!has.StartsWith(expect, StringComparison.Ordinal))
                throw new Fail("ANCHOR_MISSING", "paragraph " + para + " starts with \"" + Head(has, 40) + "\", not \"" + Head(expect, 40) + "\".");
        }

        int paragraphsBefore = (int)doc.Paragraphs.Count;
        int mathsBefore = 0;
        try { mathsBefore = (int)doc.OMaths.Count; } catch (Exception) { }
        int contentEndBefore = (int)doc.Content.End;

        // Where the equation goes. A display equation gets a paragraph of its own, an inline one
        // goes at the end of the paragraph (or next to the anchor text that -Anchor names).
        int at;
        string anchor = job.Str("anchor", "");
        if (anchor.Length > 0)
        {
            string where = job.Str("where", "after").Trim().ToLowerInvariant();
            if (where != "before" && where != "after")
                throw new Fail("BAD_ARGS", "unknown where \"" + where + "\" (before|after)");
            string text = CleanPara(Convert.ToString(TryGet(() => (object)host.Range.Text, ""), CultureInfo.InvariantCulture));
            int idx = IndexOf(text, anchor, 0);
            if (idx < 0) throw new Fail("ANCHOR_MISSING", "paragraph " + para + " does not contain \"" + Head(anchor, 40) + "\".");
            at = (int)host.Range.Start + idx + (where == "before" ? 0 : anchor.Length);
        }
        else if (display)
        {
            // a paragraph of its own, right after the anchor paragraph
            host.Range.InsertParagraphAfter();
            at = (int)doc.Paragraphs.Item(Math.Min(para + 1, (int)doc.Paragraphs.Count)).Range.Start;
        }
        else at = (int)host.Range.End - 1;

        dynamic spot = doc.Range(at, at);
        spot.InsertXML(FlatDocument(xml));

        int mathsAfter = 0;
        try { mathsAfter = (int)doc.OMaths.Count; } catch (Exception) { }
        if (mathsAfter <= mathsBefore)
        {
            // WPS took the XML but made no equation: put the document back the way it was.
            try
            {
                int grew = (int)doc.Content.End - contentEndBefore;
                if (grew > 0) doc.Range(at, at + grew).Delete();
            }
            catch (Exception) { }
            throw new Fail("MATH_FAILED", "WPS inserted the XML but built no equation (OMaths stayed at " + mathsBefore
                + "); the insertion was rolled back. Rewrite the LaTeX in plainer form, or report it.");
        }

        // The package comes in as a paragraph of its own. For an inline formula the mark it brought
        // is taken out again so the line stays one line; for a display formula an empty paragraph
        // left behind is removed so no blank line remains.
        try
        {
            int now = (int)doc.Paragraphs.Count;
            if (!display && now > paragraphsBefore)
            {
                dynamic line = doc.Range(at, at).Paragraphs.Item(1).Range;
                int lineEnd = (int)line.End;
                doc.Range(lineEnd - 1, lineEnd).Delete();
            }
            else if (display && now > paragraphsBefore)
            {
                RemoveEmptyParagraphNear(doc, para + 1);
            }
        }
        catch (Exception) { }

        dynamic landed = null;
        try { landed = doc.Range(at, at + 1).Paragraphs.Item(1); }
        catch (Exception)
        {
            try { landed = doc.Paragraphs.Item(Math.Min(para + 1, (int)doc.Paragraphs.Count)); }
            catch (Exception) { }
        }

        if (display && landed != null)
        {
            try
            {
                dynamic f = landed.Format;
                f.Alignment = 1;
                f.CharacterUnitFirstLineIndent = 0;
                f.FirstLineIndent = 0;
                if (Convert.ToDouble(f.SpaceBefore) < 6) f.SpaceBefore = 6;
                if (Convert.ToDouble(f.SpaceAfter) < 6) f.SpaceAfter = 6;
            }
            catch (Exception) { }
        }

        string readBack = "";
        try
        {
            int count = (int)doc.OMaths.Count;
            readBack = Convert.ToString(TryGet(() => (object)doc.OMaths.Item(count).Range.Text, ""), CultureInfo.InvariantCulture);
            readBack = CleanPara(readBack);
        }
        catch (Exception) { }

        Dictionary<string, object> one = new Dictionary<string, object>();
        one["para"] = para;
        one["latex"] = latex;
        one["display"] = display;
        one["number"] = number;
        one["readBack"] = readBack;
        one["omaths"] = TryGet(() => (object)(int)doc.OMaths.Count, 0);
        return one;
    }

    /// <summary>Deletes the first empty paragraph at or after `from` (one the insertion left behind).</summary>
    private static void RemoveEmptyParagraphNear(dynamic doc, int from)
    {
        int total = (int)doc.Paragraphs.Count;
        for (int k = Math.Max(1, from); k <= Math.Min(total, from + 2); k++)
        {
            dynamic p = doc.Paragraphs.Item(k);
            string text = CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));
            int maths = Convert.ToInt32(TryGet(() => (object)(int)p.Range.OMaths.Count, (object)0), CultureInfo.InvariantCulture);
            if (text.Length != 0 || maths > 0) continue;
            dynamic range = p.Range;
            int end = (int)range.End;
            if (end - 1 <= 0) continue;
            doc.Range(end - 1, end).Delete();
            return;
        }
    }
}
