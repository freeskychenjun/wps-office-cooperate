// In-memory block replace for the Excel helper: read the block once, replace locally, write only
// the cells that actually changed, once.
//
// WHY NOT Range.Replace
// ---------------------
// WPS re-scans the text it has just written when Range.Replace runs: replacing "P" with "P-" on a
// cell that holds "P" produces "P--" instead of "P-" (measured on WPS 12.1.0.29161, same finding
// as dsh-office helper/OfficeHelper.cs:5045-5065). The same call also walks the whole sheet one
// cell at a time, so a single row can be hit twice and there is no before/after report.
//
// A local replace is immune by construction: one pass over the ORIGINAL string, the scanner resumes
// at the END of the inserted text, so text we just wrote is never looked at again. That is also why
// the report can list every changed cell with its before/after and why -DryRun is exact.
//
// Discipline, same as everywhere else in this skill: cells that hold formulas are never written, a
// write block is a 0-based object[,] of the exact size, and only cells whose text really changed
// are written back (untouched cells keep their formatting, comments and data validation).
//
// ASCII only: no non-ASCII character appears in this file.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

internal static partial class Program
{
    private sealed class ReplaceHit
    {
        public int R;
        public int C;
        public string Cell;
        public string Before;
        public string After;
    }

    /// <summary>Text of a cell the way a person sees it; numbers use the plain invariant form.</summary>
    private static string ReplaceCellText(object v)
    {
        if (v == null) return "";
        if (v is string) return (string)v;
        if (IsNum(v))
        {
            try { return Plain(Convert.ToDouble(v, CultureInfo.InvariantCulture)); }
            catch (Exception) { return Convert.ToString(v, CultureInfo.InvariantCulture); }
        }
        return Convert.ToString(v, CultureInfo.InvariantCulture);
    }

    private static object CmdReplace(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb))
            throw new Fail("READONLY", "the workbook is read-only; replace is refused (open it for editing first)");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");

        string find = a.Str("find", null);
        if (find == null || find.Length == 0)
            throw new Fail("BAD_ARGS", "replace needs -Find <text>; an empty search text is refused because it would rewrite every cell");

        bool hasNew = a.Has("replace");
        string repl = hasNew ? a.Str("replace", "") : "";
        bool whole = a.Flag("whole", false);
        bool matchCase = a.Flag("case", false) || a.Flag("matchCase", false);
        bool useRegex = a.Flag("regex", false);
        bool dryRun = a.Flag("dryRun", false);
        int cap = a.Int("max", 0);
        if (cap < 0) cap = 0;

        // ---- the block, read once
        string refText = a.Str("range", null);
        dynamic range = (refText == null || refText.Trim().Length == 0) ? ws.UsedRange : SheetRef(wb, ws, refText);
        Block b = new Block();
        b.Ws = ws;
        b.Range = range;
        b.Sheet = Convert.ToString(TryGet(() => (object)range.Worksheet.Name, ""), CultureInfo.InvariantCulture);
        b.Address = Convert.ToString(TryGet(() => (object)range.Address[false, false], ""), CultureInfo.InvariantCulture);
        b.Rows = (int)range.Rows.Count;
        b.Cols = (int)range.Columns.Count;
        b.Row0 = (int)range.Row;
        b.Col0 = (int)range.Column;
        b.Grid = ReadValues(range);
        b.Headers = new string[b.Cols + 1];
        for (int c = 1; c <= b.Cols; c++) b.Headers[c] = TextOf(b.Grid[1, c]).Trim();
        if (b.Rows < 1 || b.Cols < 1) throw new Fail("BAD_ARGS", "the range " + b.Address + " is empty");

        // The formula block is one more COM call; a cell whose formula text starts with "=" holds a
        // formula and is never overwritten with a value.
        bool[,] isFormula = new bool[b.Rows + 1, b.Cols + 1];
        try
        {
            object rawFormula = range.Formula;
            object[,] fb = rawFormula as object[,];
            if (fb != null)
            {
                for (int r = 1; r <= b.Rows; r++)
                    for (int c = 1; c <= b.Cols; c++)
                    {
                        string f = TextOf(fb[r, c]);
                        isFormula[r, c] = f.Length > 0 && f[0] == '=';
                    }
            }
            else if (rawFormula != null && b.Rows == 1 && b.Cols == 1)
            {
                string f = TextOf(rawFormula);
                isFormula[1, 1] = f.Length > 0 && f[0] == '=';
            }
        }
        catch (Exception) { /* no formula information: treat every cell as a value cell */ }

        // ---- the matcher
        Regex rx = null;
        if (useRegex)
        {
            RegexOptions opt = RegexOptions.CultureInvariant;
            if (!matchCase) opt |= RegexOptions.IgnoreCase;
            string pattern = whole ? "\\A(?:" + find + ")\\z" : find;
            try { rx = new Regex(pattern, opt, TimeSpan.FromSeconds(5)); }
            catch (ArgumentException e)
            {
                throw new Fail("BAD_ARGS", "the -Regex pattern is not valid: " + Clean(e.Message)
                    + "; drop -Regex to replace literal text, or fix the pattern");
            }
        }

        // One pass over the original text. The scanner resumes after the inserted text, so what we
        // just wrote is never examined again - the exact opposite of WPS Range.Replace.
        Func<string, string> pass = delegate(string text)
        {
            if (rx != null)
            {
                if (!rx.IsMatch(text)) return null;
                return rx.Replace(text, repl);
            }
            if (whole)
            {
                bool same = matchCase
                    ? string.Equals(text, find, StringComparison.Ordinal)
                    : string.Equals(text, find, StringComparison.OrdinalIgnoreCase);
                return same ? repl : null;
            }
            StringBuilder sb = new StringBuilder(text.Length + 16);
            int i = 0;
            bool hit = false;
            while (i < text.Length)
            {
                int at = matchCase
                    ? text.IndexOf(find, i, StringComparison.Ordinal)
                    : text.IndexOf(find, i, StringComparison.OrdinalIgnoreCase);
                if (at < 0) { sb.Append(text, i, text.Length - i); break; }
                sb.Append(text, i, at - i);
                sb.Append(repl);
                i = at + find.Length;
                hit = true;
            }
            return hit ? sb.ToString() : null;
        };

        // ---- walk the block in memory
        List<int> columns = new List<int>();
        if (a.Has("column")) columns.Add(ColumnOf(b, a.Raw("column")));
        else for (int c = 1; c <= b.Cols; c++) columns.Add(c);

        List<ReplaceHit> hits = new List<ReplaceHit>();
        Dictionary<string, object> byColumn = new Dictionary<string, object>();
        Dictionary<string, int> colKeys = new Dictionary<string, int>(); // column -> used key ordinal
        int scanned = 0, formulaSkipped = 0, nonTextSkipped = 0, numericHits = 0, unchanged = 0;
        bool truncated = false;

        foreach (int c in columns)
        {
            string label = ColLetter(b.Col0 + c - 1);
            string head = b.Headers[c];
            string key = head.Length > 0 ? head : label;
            // Two columns can share a header; keep the counts apart by appending the letter.
            int seen;
            if (colKeys.TryGetValue(key, out seen))
            {
                colKeys[key] = seen + 1;
                key = key + " (" + label + ")";
            }
            else colKeys[key] = 1;

            for (int r = 1; r <= b.Rows; r++)
            {
                scanned++;
                object v = b.Grid[r, c];
                if (isFormula[r, c]) { formulaSkipped++; continue; }
                string text = ReplaceCellText(v);
                if (text.Length == 0) continue;
                if (!(v is string))
                {
                    // Numbers, dates, booleans and error values are never rewritten as text; report
                    // how many of them would have matched so the caller is not left guessing.
                    if (pass(text) != null) numericHits++;
                    nonTextSkipped++;
                    continue;
                }
                string after = pass(text);
                if (after == null) continue;
                if (string.Equals(after, text, StringComparison.Ordinal)) { unchanged++; continue; }
                if (cap > 0 && hits.Count >= cap) { truncated = true; break; }
                ReplaceHit h = new ReplaceHit();
                h.R = r; h.C = c;
                h.Cell = ColLetter(b.Col0 + c - 1) + (b.Row0 + r - 1).ToString(CultureInfo.InvariantCulture);
                h.Before = text;
                h.After = after;
                hits.Add(h);
                int n = 0;
                if (byColumn.ContainsKey(key)) n = (int)byColumn[key];
                byColumn[key] = n + 1;
            }
            if (truncated) break;
        }

        // ---- report
        int sampleLimit = a.Flag("all", false) ? hits.Count : 30;
        List<object> samples = new List<object>();
        for (int k = 0; k < hits.Count && k < sampleLimit; k++)
        {
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["cell"] = hits[k].Cell;
            one["before"] = hits[k].Before;
            one["after"] = hits[k].After;
            samples.Add(one);
        }

        Dictionary<string, object> r2 = new Dictionary<string, object>();
        r2["sheet"] = b.Sheet;
        r2["range"] = b.Address;
        r2["find"] = find;
        r2["replace"] = hasNew ? repl : null;
        r2["whole"] = whole;
        r2["matchCase"] = matchCase;
        r2["regex"] = useRegex;
        r2["dryRun"] = dryRun;
        r2["scanned"] = scanned;
        r2["changed"] = hits.Count;
        r2["wouldChange"] = hits.Count;
        r2["samples"] = samples;
        r2["samplesTruncated"] = hits.Count > samples.Count;
        r2["byColumn"] = byColumn;
        r2["formulaCellsSkipped"] = formulaSkipped;
        r2["nonTextCellsSkipped"] = nonTextSkipped;
        r2["numericMatchesLeftAlone"] = numericHits;
        r2["unchangedMatches"] = unchanged;
        r2["truncated"] = truncated;
        r2["engine"] = "memory-block";

        if (numericHits > 0)
        {
            r2["numericNote"] = numericHits.ToString(CultureInfo.InvariantCulture)
                + " numeric/date cell(s) also contain the search text but were left alone: numbers are never rewritten as text (that would silently turn a number column into text)";
        }

        if (dryRun)
        {
            r2["applied"] = 0;
            r2["note"] = hits.Count == 0
                ? "nothing matches in " + b.Address + ": check -MatchCase, -Whole, the range/sheet, or whether the text is stored as a formula"
                : "dry run: nothing was written; drop -DryRun to apply the same list";
            return r2;
        }
        if (hits.Count == 0)
        {
            r2["applied"] = 0;
            r2["note"] = "nothing matches in " + b.Address + ": check -MatchCase, -Whole, the range/sheet, or whether the text is stored as a formula";
            return r2;
        }

        // ---- one write per row run, only the cells that changed
        Dictionary<string, object> snap = BatchSnapshot(wb, "replace", new List<string> { b.Sheet });
        List<object> written = new List<object>();
        int ok = 0;
        PerfState perf = PerfOn();
        try
        {
            List<ReplaceHit> ordered = new List<ReplaceHit>(hits);
            ordered.Sort(delegate(ReplaceHit x, ReplaceHit y)
            {
                if (x.R != y.R) return x.R.CompareTo(y.R);
                return x.C.CompareTo(y.C);
            });
            int i2 = 0;
            while (i2 < ordered.Count)
            {
                int j = i2;
                while (j + 1 < ordered.Count && ordered[j + 1].R == ordered[i2].R && ordered[j + 1].C == ordered[j].C + 1) j++;
                int len = j - i2 + 1;
                dynamic target = b.Range.Cells[ordered[i2].R, ordered[i2].C];
                if (len > 1) target = target.Resize[1, len];
                if (len == 1)
                {
                    target.Value2 = ordered[i2].After;
                }
                else
                {
                    object[,] block = new object[1, len]; // 0-based, exact size - see excel-quirks.md
                    for (int k = 0; k < len; k++) block[0, k] = ordered[i2 + k].After;
                    target.Value2 = block;
                }
                ok += len;
                for (int k = i2; k <= j; k++)
                {
                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["cell"] = ordered[k].Cell;
                    one["after"] = ordered[k].After;
                    written.Add(one);
                }
                i2 = j + 1;
            }
        }
        catch (Exception e)
        {
            // Put the captured text back: a half-applied replace is worse than a failed one.
            int back = 0;
            foreach (ReplaceHit h in hits)
            {
                try
                {
                    dynamic cell = b.Range.Cells[h.R, h.C];
                    cell.Value2 = h.Before;
                    back++;
                }
                catch (Exception) { }
            }
            throw new Fail("COM_ERROR", "replace stopped after " + ok + " cell(s): " + Clean(e.Message)
                + "; put back " + back + " cell(s) of the " + hits.Count + " planned");
        }
        finally { PerfOff(perf); }

        r2["applied"] = ok;
        r2["written"] = written;
        r2["preSnapshot"] = snap;
        r2["undoNote"] = "spreadsheet edits cannot be undone with Ctrl+Z: pre-batch snapshot taken (preSnapshot.path); diff -Since <path> shows exactly this change, restore -Since <path> undoes it";
        return r2;
    }
}
