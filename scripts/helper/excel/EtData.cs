// Data work for the Excel helper: cleaning, de-duplicating, sorting, filtering, copying,
// adding columns, pivot tables and the "finished look" commands (freeze, conditional formats,
// a drop-down list, a ListObject table), plus the overview composite image of render -Overview.
//
// Discipline, same as everywhere else in this skill: read the block once into memory, compute
// locally, write the block back once. Writing back needs a 0-based object[,] of the exact size
// (a 1-based array is accepted by COM and silently lands nothing - see reference/excel-quirks.md).
// Cells that hold formulas are never overwritten by a value.
//
// ASCII only: every non-ASCII character that has to appear in a result is written as a \u escape.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

internal static partial class Program
{
    // ---------------------------------------------------------------- value helpers

    private static bool IsNum(object v) { return v is double || v is int || v is long || v is decimal || v is float; }

    private static bool IsBlankValue(object v) { return v == null || (v is string && ((string)v).Trim().Length == 0); }

    private static string TextOf(object v) { return v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture); }

    /// <summary>Trim, one kind of space only, no runs of spaces: "2024 / 3 1" becomes readable again.</summary>
    private static string Tidy(string t)
    {
        if (t == null) return "";
        string s = t.Replace('\u00a0', ' ').Replace('\u3000', ' ').Replace("\t", " ").Trim();
        while (s.Contains("  ")) s = s.Replace("  ", " ");
        return s;
    }

    private static readonly Regex NumberPattern = new Regex(
        "^[\\u00a5\\uffe5$\\u20ac\\u00a3]?\\s*([-+]?\\d*\\.?\\d+(?:[eE][-+]?\\d+)?)\\s*(%|\\uff05)?\\s*(?:[^\\d\\s\\-.,A-Za-z%\\uff05]{1,3}|kg|g|km|cm|mm|m|pcs|k|w)?\\z",
        RegexOptions.CultureInvariant);

    /// <summary>A number out of what a person typed: "¥1,299", "89元", "10%", "12 件".</summary>
    private static bool ParseNumber(string text, out double number)
    {
        number = 0;
        string s = Tidy(text).Replace(",", "").Replace("\uff0c", "").Replace(" ", "");
        Match m = NumberPattern.Match(s);
        if (!m.Success) return false;
        if (!double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) return false;
        if (m.Groups[2].Success) number /= 100;
        return true;
    }

    private static readonly Regex DateSolid = new Regex("^(\\d{4})(\\d{2})(\\d{2})\\z", RegexOptions.CultureInvariant);
    private static readonly Regex DateFull = new Regex(
        "^(\\d{4})\\s*[-/.\\u5e74]\\s*(\\d{1,2})\\s*[-/.\\u6708]\\s*(\\d{1,2})\\s*\\u65e5?(?:\\s.*)?\\z",
        RegexOptions.CultureInvariant);
    private static readonly Regex DateMonthDay = new Regex(
        "^(\\d{1,2})\\s*[-/.\\u6708]\\s*(\\d{1,2})\\s*\\u65e5?\\z", RegexOptions.CultureInvariant);
    private static readonly Regex DateShaped = new Regex(
        "^(\\d{8}|\\d{4}\\s*[-/.\\u5e74]\\s*\\d{1,2}\\s*[-/.\\u6708]\\s*\\d{1,2}.*|\\d{1,2}\\s*\\u6708\\s*\\d{1,2}\\s*\\u65e5?)\\z",
        RegexOptions.CultureInvariant);

    /// <summary>A date out of what a person typed: 2025/3/1, 2025.03.01, 20250301, 2025年3月1日, 3月1日 (the year given).</summary>
    private static bool ParseDate(string text, int year, out DateTime date)
    {
        date = DateTime.MinValue;
        string s = Tidy(text);
        Match m = DateSolid.Match(s);
        if (!m.Success) m = DateFull.Match(s);
        int y = 0, mo = 0, d = 0;
        if (m.Success)
        {
            y = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            mo = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            d = int.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
        }
        else
        {
            m = DateMonthDay.Match(s);
            if (m.Success)
            {
                y = year;
                mo = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                d = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            }
        }
        if (m.Success)
        {
            if (y < 1900 || y > 2200 || mo < 1 || mo > 12 || d < 1 || d > DateTime.DaysInMonth(y, mo)) return false;
            date = new DateTime(y, mo, d);
            return true;
        }
        return DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) && date.Year >= 1900;
    }

    /// <summary>The "#DIV/0!" kind of text for a cell that holds an error, else null.</summary>
    private static string ErrorText(object v)
    {
        if (!(v is int) && !(v is long)) return null;
        long l = Convert.ToInt64(v, CultureInfo.InvariantCulture);
        switch (l)
        {
            case -2146826281L: return "#DIV/0!";
            case -2146826246L: return "#N/A";
            case -2146826259L: return "#NAME?";
            case -2146826288L: return "#NULL!";
            case -2146826252L: return "#NUM!";
            case -2146826265L: return "#REF!";
            case -2146826273L: return "#VALUE!";
        }
        return null;
    }

    private static string Plain(double number)
    {
        if (Math.Abs(number - Math.Round(number)) < 1e-9 && Math.Abs(number) < 1e15)
            return Math.Round(number).ToString("0", CultureInfo.InvariantCulture);
        return number.ToString("0.####", CultureInfo.InvariantCulture);
    }

    /// <summary>A JSON scalar as the cell value it should become.</summary>
    private static object ScalarOf(object v)
    {
        if (v == null) return null;
        if (v is double) { double d = (double)v; if (d == Math.Floor(d) && Math.Abs(d) < 1e15) return (long)d; return d; }
        return v;
    }

    private static string Clip(string s, int max)
    {
        if (s == null) return "";
        return s.Length <= max ? s : s.Substring(0, max) + "..";
    }

    // ---------------------------------------------------------------- the block a command works on

    private sealed class Block
    {
        public dynamic Ws;
        public dynamic Range;
        public string Sheet;
        public string Address;
        public int Rows;
        public int Cols;
        public int Row0;
        public int Col0;
        public object[,] Grid;      // 1-based, as COM hands it back
        public string[] Headers;    // 1-based
    }

    private static Block BlockOf(dynamic ws, Bag a)
    {
        string refText = a.Str("range", null);
        dynamic range;
        try { range = (refText == null || refText.Trim().Length == 0) ? ws.UsedRange : ws.Range(refText); }
        catch (Exception) { throw new Fail("BAD_ARGS", "cannot read the range \"" + refText + "\""); }
        Block b = new Block();
        b.Ws = ws;
        b.Range = range;
        b.Sheet = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
        b.Address = Convert.ToString(TryGet(() => (object)range.Address[false, false], ""), CultureInfo.InvariantCulture);
        b.Rows = (int)range.Rows.Count;
        b.Cols = (int)range.Columns.Count;
        b.Row0 = (int)range.Row;
        b.Col0 = (int)range.Column;
        b.Grid = ReadValues(range);
        b.Headers = new string[b.Cols + 1];
        for (int c = 1; c <= b.Cols; c++) b.Headers[c] = TextOf(b.Grid[1, c]).Trim();
        return b;
    }

    /// <summary>A column of the block: by header text, by position (1-based), or by sheet column letter.</summary>
    private static int ColumnOf(Block b, object key)
    {
        string name = TextOf(key).Trim();
        if (name.Length == 0) throw new Fail("BAD_ARGS", "a column is needed: its header, its position or its letter");
        for (int c = 1; c <= b.Cols; c++)
            if (string.Equals(b.Headers[c], name, StringComparison.OrdinalIgnoreCase)) return c;
        int at = 0;
        if (int.TryParse(name, NumberStyles.Integer, CultureInfo.InvariantCulture, out at) && at >= 1 && at <= b.Cols) return at;
        if (name.Length <= 3 && Regex.IsMatch(name, "^[A-Za-z]+$"))
        {
            int letter = 0;
            foreach (char ch in name.ToUpperInvariant()) letter = letter * 26 + (ch - 'A' + 1);
            at = letter - b.Col0 + 1;
            if (at >= 1 && at <= b.Cols) return at;
        }
        List<string> heads = new List<string>();
        for (int c = 1; c <= b.Cols; c++) heads.Add(b.Headers[c].Length == 0 ? "(" + ColLetter(b.Col0 + c - 1) + ")" : b.Headers[c]);
        throw new Fail("ANCHOR_MISSING", "there is no column \"" + name + "\" in " + b.Address + ". Its headers: " + string.Join(", ", heads.ToArray()) + ".");
    }

    private static List<int> ColumnsOf(Block b, Bag a, bool allWhenNone)
    {
        List<int> list = new List<int>();
        List<object> many = a.List("columns");
        if (many.Count > 0) { foreach (object key in many) list.Add(ColumnOf(b, key)); }
        else if (a.Has("column")) list.Add(ColumnOf(b, a.Raw("column")));
        else if (allWhenNone) for (int c = 1; c <= b.Cols; c++) list.Add(c);
        return list;
    }

    private static bool ColumnHasFormulas(Block b, int c)
    {
        try
        {
            object has = b.Range.Columns[c].HasFormula;
            if (has is bool) return (bool)has;
        }
        catch (Exception) { }
        return false;
    }

    /// <summary>The display format of a range, as a string: WPS answers Null (or a 2D array of
    /// formats) when the cells of a multi-column block disagree, so anything else is ignored.</summary>
    private static string FormatText(object raw)
    {
        if (raw == null) return "";
        string one = raw as string;
        if (one != null) return one;
        object[,] block = raw as object[,];
        if (block != null)
        {
            try { return Convert.ToString(block[1, 1], CultureInfo.InvariantCulture); }
            catch (Exception) { return ""; }
        }
        return Convert.ToString(raw, CultureInfo.InvariantCulture);
    }

    private static bool LooksLikeDate(string format)
    {
        if (format == null || format.Length == 0) return false;
        return Regex.IsMatch(format, "[ymd\\u5e74\\u6708\\u65e5]", RegexOptions.IgnoreCase)
            && format.IndexOf("0.0", StringComparison.Ordinal) < 0
            && !string.Equals(format, "General", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Which columns are shown as dates: decided by the display format, never by the value.</summary>
    private static bool[] DateColumnsOf(dynamic range, object[,] grid, int rows, int cols)
    {
        bool[] dated = new bool[cols + 1];
        for (int c = 1; c <= cols; c++)
        {
            bool anyNumber = false;
            for (int r = 1; r <= rows; r++) if (IsNum(grid[r, c])) { anyNumber = true; break; }
            if (!anyNumber) continue;
            // A whole column usually shares one format; when the header disagrees WPS answers Null, so a
            // few data cells are asked instead (3 COM calls per column, not one per cell).
            string format = FormatText(TryGet(() => (object)range.Columns[c].NumberFormat, null));
            if (format.Length > 0)
            {
                dated[c] = LooksLikeDate(format);
                continue;
            }
            int seen = 0, hits = 0;
            int[] probes = new int[] { 2, Math.Max(2, rows / 2), rows };
            int lastProbe = 0;
            for (int k = 0; k < probes.Length; k++)
            {
                int r = probes[k];
                if (r < 1 || r > rows || r == lastProbe) continue;
                lastProbe = r;
                if (!IsNum(grid[r, c])) continue;
                seen++;
                if (LooksLikeDate(FormatText(TryGet(() => (object)range.Cells[r, c].NumberFormat, null)))) hits++;
            }
            dated[c] = seen > 0 && hits * 2 > seen;
        }
        return dated;
    }

    private static string DateText(double serial, bool withTime)
    {
        try
        {
            DateTime d = DateTime.FromOADate(serial);
            if (!withTime || (d.Hour == 0 && d.Minute == 0 && d.Second == 0))
                return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            return d.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        }
        catch (Exception) { return Plain(serial); }
    }

    // ---------------------------------------------------------------- deleting rows

    /// <summary>Delete rows bottom-up, neighbours merged into one piece: one COM call per 18 runs.</summary>
    private static void DeleteRows(dynamic ws, List<int> rows)
    {
        rows.Sort();
        for (int i = rows.Count - 1; i >= 0; )
        {
            StringBuilder address = new StringBuilder();
            int taken = 0;
            while (i >= 0 && taken < 18)
            {
                int last = rows[i], first = last;
                while (i > 0 && rows[i - 1] == first - 1) { i--; first = rows[i]; }
                i--;
                if (address.Length > 0) address.Append(',');
                address.Append(first).Append(':').Append(last);
                taken++;
            }
            ws.Range[address.ToString()].Delete();
        }
    }

    private static string RowList(List<int> rows)
    {
        List<string> parts = new List<string>();
        for (int i = 0; i < rows.Count && i < 12; i++) parts.Add(rows[i].ToString(CultureInfo.InvariantCulture));
        return string.Join(", ", parts.ToArray()) + (rows.Count > 12 ? ", .." : "");
    }

    /// <summary>Cells among a range that show an error, so a bad formula is reported at once.</summary>
    private static List<object> ErrorCells(dynamic range)
    {
        List<object> found = new List<object>();
        try
        {
            object[,] grid = ReadValues(range);
            int row0 = (int)range.Row, col0 = (int)range.Column;
            for (int r = 1; r <= grid.GetLength(0); r++)
                for (int c = 1; c <= grid.GetLength(1); c++)
                {
                    string code = ErrorText(grid[r, c]);
                    if (code == null) continue;
                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["cell"] = ColLetter(col0 + c - 1) + (row0 + r - 1).ToString(CultureInfo.InvariantCulture);
                    one["code"] = code;
                    found.Add(one);
                    if (found.Count >= 20) return found;
                }
        }
        catch (Exception) { }
        return found;
    }

    // ---------------------------------------------------------------- conditions

    private static bool Test(object value, string how, object against, List<object> among)
    {
        string text = Tidy(TextOf(value));
        switch (how)
        {
            case "blank": return IsBlankValue(value);
            case "notblank": case "not blank": return !IsBlankValue(value);
            case "contains": return text.IndexOf(Tidy(TextOf(against)), StringComparison.OrdinalIgnoreCase) >= 0;
            case "in":
                if (among != null) foreach (object one in among)
                    if (string.Equals(text, Tidy(TextOf(one)), StringComparison.OrdinalIgnoreCase)) return true;
                return false;
            case "text": return value is string && !IsBlankValue(value);
            case "error": return ErrorText(value) != null;
        }
        if (IsBlankValue(value)) return false;
        double a = 0, b = 0;
        bool left = true, right = true;
        if (IsNum(value)) a = Convert.ToDouble(value, CultureInfo.InvariantCulture); else left = ParseNumber(text, out a);
        if (IsNum(against)) b = Convert.ToDouble(against, CultureInfo.InvariantCulture); else right = ParseNumber(TextOf(against), out b);
        bool numbers = left && right;
        int order = numbers ? a.CompareTo(b) : string.Compare(text, Tidy(TextOf(against)), StringComparison.OrdinalIgnoreCase);
        switch (how)
        {
            case "=": case "==": case "equals": return order == 0;
            case "<>": case "!=": return order != 0;
            case ">": return numbers && order > 0;
            case "<": return numbers && order < 0;
            case ">=": return numbers && order >= 0;
            case "<=": return numbers && order <= 0;
        }
        throw new Fail("BAD_ARGS", "unknown test \"" + how + "\": use blank, notblank, =, <>, >, <, >=, <=, contains, in, text or error");
    }

    // ---------------------------------------------------------------- clean

    private static object CmdClean(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "the workbook is read-only; cleaning is refused");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        Block b = BlockOf(ws, a);

        string to = a.Str("to", null);
        string casing = a.Str("case", null);
        bool trim = a.Flag("trim", true), noSpaces = a.Flag("noSpaces", false);
        List<object> remove = a.List("remove");
        if (remove.Count == 0 && a.Has("remove")) remove.Add(a.Need("remove"));

        Dictionary<string, object> lookup = null;
        object rawMap = a.Raw("map");
        Dictionary<string, object> map = rawMap as Dictionary<string, object>;
        if (map != null)
        {
            lookup = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, object> pair in map) lookup[Tidy(pair.Key)] = pair.Value;
        }

        bool converts = to != null || lookup != null || casing != null || remove.Count > 0;
        List<int> columns = ColumnsOf(b, a, !converts);
        if (columns.Count == 0)
            throw new Fail("BAD_ARGS", "clean needs -Column/-Columns when it converts, maps or changes the case");
        int year = a.Int("year", 0);
        string numberFormat = a.Str("numberFormat", null);

        List<object> said = new List<object>();
        int totalChanged = 0, totalFailed = 0;
        foreach (int c in columns)
        {
            if (to == "date" && year == 0)
            {
                // Dates given without a year take the year the rest of the column is in.
                Dictionary<int, int> years = new Dictionary<int, int>();
                for (int r = 2; r <= b.Rows; r++)
                {
                    DateTime known = DateTime.MinValue;
                    object v = b.Grid[r, c];
                    if (IsNum(v)) { try { known = DateTime.FromOADate(Convert.ToDouble(v, CultureInfo.InvariantCulture)); } catch (Exception) { } }
                    else if (v is string && !Regex.IsMatch(Tidy((string)v), "^\\d{1,2}\\s*[-/.\\u6708]")) ParseDate((string)v, 1, out known);
                    if (known.Year > 1900) { int n; years.TryGetValue(known.Year, out n); years[known.Year] = n + 1; }
                }
                int best = 0;
                foreach (KeyValuePair<int, int> pair in years) if (pair.Value > best) { best = pair.Value; year = pair.Key; }
                if (year == 0) year = DateTime.Now.Year;
            }

            bool formulas = ColumnHasFormulas(b, c);
            int changed = 0, failed = 0;
            List<string> samples = new List<string>();
            object[,] fresh = new object[b.Rows, 1];
            fresh[0, 0] = b.Grid[1, c];
            for (int r = 2; r <= b.Rows; r++)
            {
                object before = b.Grid[r, c], after = before;
                string s = before as string;
                if (s != null)
                {
                    if (trim) s = Tidy(s);
                    if (noSpaces) s = s.Replace(" ", "");
                    if (remove.Count > 0) foreach (object piece in remove) { string cut = TextOf(piece); if (cut.Length > 0) s = s.Replace(cut, ""); }
                    if (trim) s = s.Trim();
                    if (casing == "upper") s = s.ToUpperInvariant();
                    else if (casing == "lower") s = s.ToLowerInvariant();
                    else if (casing == "proper") s = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(s.ToLowerInvariant());
                    after = s.Length == 0 ? null : (object)s;
                    if (lookup != null)
                    {
                        object mapped;
                        if (lookup.TryGetValue(s, out mapped)) { after = ScalarOf(mapped); s = after as string; }
                    }
                    if (s != null && s.Length > 0)
                    {
                        double number; DateTime date;
                        if (to == "number")
                        {
                            if (ParseNumber(s, out number)) after = number;
                            else { failed++; if (samples.Count < 5) samples.Add(ColLetter(b.Col0 + c - 1) + (b.Row0 + r - 1).ToString(CultureInfo.InvariantCulture) + " \"" + Clip(s, 16) + "\""); }
                        }
                        else if (to == "date")
                        {
                            if (ParseDate(s, year, out date)) after = date.ToOADate();
                            else { failed++; if (samples.Count < 5) samples.Add(ColLetter(b.Col0 + c - 1) + (b.Row0 + r - 1).ToString(CultureInfo.InvariantCulture) + " \"" + Clip(s, 16) + "\""); }
                        }
                    }
                }
                else if (before != null && to == "text") after = TextOf(before);
                else if (to == "date" && IsNum(before)
                    && Convert.ToDouble(before, CultureInfo.InvariantCulture) >= 19000101
                    && Convert.ToDouble(before, CultureInfo.InvariantCulture) <= 22001231)
                {
                    DateTime date;
                    if (ParseDate(Convert.ToInt64(Convert.ToDouble(before, CultureInfo.InvariantCulture)).ToString(CultureInfo.InvariantCulture), year, out date))
                        after = date.ToOADate();
                    else { failed++; if (samples.Count < 5) samples.Add(ColLetter(b.Col0 + c - 1) + (b.Row0 + r - 1).ToString(CultureInfo.InvariantCulture) + " " + TextOf(before)); }
                }
                else if (before != null && lookup != null)
                {
                    object mapped;
                    if (lookup.TryGetValue(TextOf(before), out mapped)) after = ScalarOf(mapped);
                }
                fresh[r - 1, 0] = after;
                if (!object.Equals(before, after)) changed++;
            }

            if (changed > 0)
            {
                if (!formulas)
                {
                    if (to == "text") b.Range.Columns[c].NumberFormat = "@";
                    b.Range.Columns[c].Value2 = fresh;
                }
                else
                {
                    // Formulas stand in this column: only the plain cells that changed are written.
                    for (int r = 2; r <= b.Rows; r++)
                    {
                        if (object.Equals(b.Grid[r, c], fresh[r - 1, 0])) continue;
                        dynamic cell = b.Range.Columns[c].Cells[r, 1];
                        bool hasFormula = false;
                        try { object f = cell.HasFormula; hasFormula = f is bool && (bool)f; } catch (Exception) { }
                        if (hasFormula) continue;
                        cell.Value2 = fresh[r - 1, 0];
                    }
                }
            }
            if (to == "date")
            {
                try { b.Range.Columns[c].Offset[1, 0].Resize[Math.Max(1, b.Rows - 1), 1].NumberFormat = "yyyy-mm-dd"; } catch (Exception) { }
            }
            else if (to == "number" && numberFormat != null)
            {
                try { b.Range.Columns[c].Offset[1, 0].Resize[Math.Max(1, b.Rows - 1), 1].NumberFormat = numberFormat; } catch (Exception) { }
            }

            // What the column holds now, when it is a handful of values: the check that cleaning worked.
            Dictionary<string, int> seen = new Dictionary<string, int>();
            bool few = true;
            for (int r = 2; r <= b.Rows && few; r++)
            {
                object v = fresh[r - 1, 0];
                if (IsBlankValue(v)) continue;
                string shown = v is double ? Plain((double)v) : TextOf(v);
                int n;
                seen.TryGetValue(shown, out n);
                seen[shown] = n + 1;
                if (seen.Count > 16) few = false;
            }
            List<string> now = new List<string>();
            if (few && seen.Count > 0 && to != "number" && to != "date")
                foreach (KeyValuePair<string, int> pair in seen)
                    now.Add("\"" + Clip(pair.Key, 20) + "\" " + pair.Value.ToString(CultureInfo.InvariantCulture));

            Dictionary<string, object> one = new Dictionary<string, object>();
            one["column"] = ColLetter(b.Col0 + c - 1);
            one["header"] = b.Headers[c];
            one["changed"] = changed;
            one["notReadable"] = failed;
            if (samples.Count > 0) one["samples"] = samples;
            if (now.Count > 0) one["valuesNow"] = now;
            said.Add(one);
            totalChanged += changed;
            totalFailed += failed;
        }

        Dictionary<string, object> res = new Dictionary<string, object>();
        res["sheet"] = b.Sheet;
        res["range"] = b.Address;
        res["to"] = to == null ? "" : to;
        res["changed"] = totalChanged;
        res["notReadable"] = totalFailed;
        res["columns"] = said;
        if (totalChanged == 0) res["note"] = "nothing changed: the cells already looked like that";
        return res;
    }

    // ---------------------------------------------------------------- fill blanks

    private static object CmdFillBlanks(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "the workbook is read-only; filling is refused");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        Block b = BlockOf(ws, a);
        List<int> columns = ColumnsOf(b, a, false);
        if (columns.Count == 0) throw new Fail("BAD_ARGS", "fillblanks needs -Column or -Columns");
        if (!a.Has("with")) throw new Fail("BAD_ARGS", "fillblanks needs -With <value|above|mean|median|mode>");

        object withRaw = a.Raw("with");
        string how = withRaw as string;
        List<object> said = new List<object>();
        int totalFilled = 0;
        foreach (int c in columns)
        {
            object fill = ScalarOf(withRaw);
            if (how == "mean" || how == "median" || how == "mode")
            {
                List<double> numbers = new List<double>();
                Dictionary<string, int> seen = new Dictionary<string, int>();
                object commonest = null;
                int most = 0;
                for (int r = 2; r <= b.Rows; r++)
                {
                    object v = b.Grid[r, c];
                    if (IsBlankValue(v)) continue;
                    if (IsNum(v)) numbers.Add(Convert.ToDouble(v, CultureInfo.InvariantCulture));
                    string key = TextOf(v);
                    int n;
                    seen.TryGetValue(key, out n);
                    seen[key] = n + 1;
                    if (seen[key] > most) { most = seen[key]; commonest = v; }
                }
                if (how == "mode") fill = commonest;
                else
                {
                    if (numbers.Count == 0)
                        throw new Fail("BAD_ARGS", "column \"" + b.Headers[c] + "\" has no numbers to take the " + how + " of");
                    numbers.Sort();
                    double sum = 0;
                    foreach (double n in numbers) sum += n;
                    fill = how == "mean"
                        ? Math.Round(sum / numbers.Count, 4)
                        : (numbers.Count % 2 == 1 ? numbers[numbers.Count / 2] : (numbers[numbers.Count / 2 - 1] + numbers[numbers.Count / 2]) / 2);
                }
            }

            bool formulas = ColumnHasFormulas(b, c);
            object[,] fresh = new object[b.Rows, 1];
            fresh[0, 0] = b.Grid[1, c];   // the header row is written back unchanged, never dropped
            List<int> blankRows = new List<int>();
            object last = null;
            for (int r = 2; r <= b.Rows; r++)
            {
                object before = b.Grid[r, c];
                object after = before;
                if (!IsBlankValue(before)) last = before;
                else
                {
                    // An entirely empty row stays empty: it is a divider or a stray, not a missing value.
                    bool wholeRowEmpty = true;
                    for (int k = 1; k <= b.Cols; k++) if (!IsBlankValue(b.Grid[r, k])) { wholeRowEmpty = false; break; }
                    object value = how == "above" ? last : fill;
                    if (!wholeRowEmpty && value != null) { after = value; blankRows.Add(r); }
                }
                fresh[r - 1, 0] = after;
            }
            if (blankRows.Count > 0)
            {
                if (!formulas) b.Range.Columns[c].Value2 = fresh;
                else foreach (int r in blankRows) b.Range.Columns[c].Cells[r, 1].Value2 = fresh[r - 1, 0];
            }
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["column"] = ColLetter(b.Col0 + c - 1);
            one["header"] = b.Headers[c];
            one["filled"] = blankRows.Count;
            if (how == "above") one["with"] = "the cell above";
            else if (fill == null) one["with"] = "nothing";
            else one["with"] = fill is double ? (object)Plain((double)fill) : fill;
            said.Add(one);
            totalFilled += blankRows.Count;
        }
        Dictionary<string, object> res = new Dictionary<string, object>();
        res["sheet"] = b.Sheet;
        res["range"] = b.Address;
        res["filled"] = totalFilled;
        res["columns"] = said;
        return res;
    }

    // ---------------------------------------------------------------- remove duplicates

    private static object CmdDedupe(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "the workbook is read-only; de-duplicating is refused");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        Block b = BlockOf(ws, a);
        List<int> keys = ColumnsOf(b, a, true);
        Dictionary<string, int> seen = new Dictionary<string, int>();
        List<int> gone = new List<int>();
        for (int r = 2; r <= b.Rows; r++)
        {
            StringBuilder key = new StringBuilder();
            bool any = false;
            foreach (int c in keys)
            {
                if (!IsBlankValue(b.Grid[r, c])) any = true;
                key.Append(Tidy(TextOf(b.Grid[r, c])).ToUpperInvariant()).Append('\u0001');
            }
            if (!any) continue;
            string k = key.ToString();
            if (seen.ContainsKey(k)) gone.Add(b.Row0 + r - 1); else seen[k] = r;
        }
        List<string> which = new List<string>();
        foreach (int c in keys) which.Add(b.Headers[c].Length == 0 ? ColLetter(b.Col0 + c - 1) : b.Headers[c]);
        string what = keys.Count == b.Cols ? "whole rows" : "the same " + string.Join(" + ", which.ToArray());
        Dictionary<string, object> res = new Dictionary<string, object>();
        res["sheet"] = b.Sheet;
        res["range"] = b.Address;
        res["key"] = what;
        res["removed"] = gone.Count;
        if (gone.Count == 0) { res["note"] = "no duplicates: nothing deleted"; return res; }
        res["rows"] = RowList(gone);
        DeleteRows(ws, gone);
        res["note"] = "the first of each repeated group was kept; row numbers below the deleted ones have changed";
        return res;
    }

    // ---------------------------------------------------------------- delete where

    private static object CmdDeleteWhere(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "the workbook is read-only; deleting rows is refused");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        Block b = BlockOf(ws, a);
        List<int> gone = new List<int>();
        string said;
        if (!a.Has("column"))
        {
            // No column named: delete the rows that are entirely empty.
            for (int r = 2; r <= b.Rows; r++)
            {
                bool any = false;
                for (int c = 1; c <= b.Cols; c++) if (!IsBlankValue(b.Grid[r, c])) { any = true; break; }
                if (!any) gone.Add(b.Row0 + r - 1);
            }
            said = "entirely empty";
        }
        else
        {
            int c = ColumnOf(b, a.Raw("column"));
            string how = a.Str("is", null);
            List<object> among = a.List("in");
            if (how == null) how = among.Count > 0 ? "in" : (a.Has("value") ? "=" : "blank");
            object against = a.Raw("value");
            for (int r = 2; r <= b.Rows; r++)
                if (Test(b.Grid[r, c], how, against, among)) gone.Add(b.Row0 + r - 1);
            said = "\"" + b.Headers[c] + "\" " + how
                + (against != null ? " " + TextOf(against) : among.Count > 0 ? " [" + among.Count.ToString(CultureInfo.InvariantCulture) + " values]" : "");
        }
        Dictionary<string, object> res = new Dictionary<string, object>();
        res["sheet"] = b.Sheet;
        res["range"] = b.Address;
        res["condition"] = said;
        res["deleted"] = gone.Count;
        if (gone.Count == 0) { res["note"] = "nothing matched: nothing deleted"; return res; }
        res["rows"] = RowList(gone);
        DeleteRows(ws, gone);
        res["note"] = "row numbers below the deleted ones have changed";
        return res;
    }

    // ---------------------------------------------------------------- sort

    private static object CmdSort(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "the workbook is read-only; sorting is refused");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        Block b = BlockOf(ws, a);
        List<object> by = a.List("by");
        string defaultOrder = a.Str("order", "asc");
        if (by.Count == 0)
        {
            if (!a.Has("column")) throw new Fail("BAD_ARGS", "sort needs -By \"column[:desc],column2\" (or -Column and -Order)");
            by.Add(a.Raw("column"));
        }
        dynamic sort = ws.Sort;
        sort.SortFields.Clear();
        List<object> said = new List<object>();
        foreach (object raw in by)
        {
            string spec = TextOf(raw).Trim();
            bool down = defaultOrder.StartsWith("d", StringComparison.OrdinalIgnoreCase);
            int colon = spec.LastIndexOf(':');
            if (colon > 0)
            {
                string tail = spec.Substring(colon + 1).Trim();
                if (tail.StartsWith("d", StringComparison.OrdinalIgnoreCase) || tail.StartsWith("a", StringComparison.OrdinalIgnoreCase))
                {
                    down = tail.StartsWith("d", StringComparison.OrdinalIgnoreCase);
                    spec = spec.Substring(0, colon).Trim();
                }
            }
            int c = ColumnOf(b, spec);
            sort.SortFields.Add(Key: b.Range.Columns[c], SortOn: 0, Order: down ? 2 : 1, DataOption: 0);
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["column"] = b.Headers[c].Length == 0 ? ColLetter(b.Col0 + c - 1) : b.Headers[c];
            one["order"] = down ? "desc" : "asc";
            said.Add(one);
        }
        sort.SetRange(b.Range);
        sort.Header = a.Flag("header", true) ? 1 : 2;
        sort.Apply();
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = b.Sheet;
        r["range"] = b.Address;
        r["by"] = said;
        r["note"] = "sorted in place; formulas and formats moved with their rows";
        return r;
    }

    // ---------------------------------------------------------------- filter

    private static object CmdFilter(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        string action = a.Str("action", "apply").ToLowerInvariant();
        if (action == "clear" || action == "off")
        {
            try { if (Convert.ToString(TryGet(() => (object)ws.FilterMode, "False"), CultureInfo.InvariantCulture).StartsWith("T", StringComparison.OrdinalIgnoreCase)) ws.ShowAllData(); } catch (Exception) { }
            if (!a.Flag("keepArrows", false)) { try { ws.AutoFilterMode = false; } catch (Exception) { } }
            Dictionary<string, object> cleared = new Dictionary<string, object>();
            cleared["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
            cleared["note"] = "all rows shown again";
            return cleared;
        }
        Block b = BlockOf(ws, a);
        int c = ColumnOf(b, a.Raw("column"));
        List<object> values = a.List("in");
        try
        {
            if (values.Count > 0)
            {
                string[] wanted = new string[values.Count];
                for (int i = 0; i < values.Count; i++) wanted[i] = TextOf(values[i]);
                b.Range.AutoFilter(Field: c, Criteria1: wanted, Operator: 7);
            }
            else if (a.Has("criteria2"))
            {
                b.Range.AutoFilter(Field: c, Criteria1: a.Need("criteria"), Operator: a.Str("join", "and") == "or" ? 2 : 1, Criteria2: a.Need("criteria2"));
            }
            else b.Range.AutoFilter(Field: c, Criteria1: a.Need("criteria"));
        }
        catch (Exception error)
        {
            // Measured 2026-11-02: a range covered by a ListObject refuses Range.AutoFilter with
            // E_ACCESSDENIED (0x80070005) on WPS, while sort / delete / fill work there.
            throw new Fail("COM_ERROR", "WPS refused the filter on " + b.Address + " (" + Clean(error.Message)
                + "). If this range is a table (ListObject), Range.AutoFilter does not work there: filter before making the table, or delete the table first");
        }
        int shown = 0;
        try { shown = (int)b.Range.Columns[1].SpecialCells(12).Cells.Count - 1; } catch (Exception) { }
        int total = b.Rows - 1;
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = b.Sheet;
        r["range"] = b.Address;
        r["column"] = b.Headers[c].Length == 0 ? ColLetter(b.Col0 + c - 1) : b.Headers[c];
        r["shown"] = shown;
        r["total"] = total;
        r["note"] = "the other rows are hidden, not deleted; filter -Action clear shows them again";
        return r;
    }

    // ---------------------------------------------------------------- copy

    private static object CmdCopy(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "the workbook is read-only; copying into it is refused");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        string from = a.Need("range");
        string to = a.Need("to");
        dynamic source = SheetRef(wb, ws, from);
        dynamic target = SheetRef(wb, ws, to);
        bool valuesOnly = a.Flag("values", false);
        if (valuesOnly)
        {
            dynamic into = target.Cells[1, 1].Resize[(int)source.Rows.Count, (int)source.Columns.Count];
            // Writing the values over would make Excel read every text again ("20250301" would turn into a
            // number): its own copy keeps each cell as it is, and only the formulas are then replaced.
            source.Copy(into.Cells[1, 1]);
            try { _app.CutCopyMode = 0; } catch (Exception) { }
            try
            {
                dynamic formulas = into.SpecialCells(-4123);
                foreach (dynamic area in formulas.Areas) area.Value2 = area.Value2;
            }
            catch (Exception) { }
            for (int c = 1; c <= (int)source.Columns.Count; c++)
                try { into.Columns[c].ColumnWidth = source.Columns[c].ColumnWidth; } catch (Exception) { }
            Dictionary<string, object> vr = new Dictionary<string, object>();
            vr["sheet"] = b1Name(into);
            vr["at"] = Convert.ToString(TryGet(() => (object)into.Address[false, false], ""), CultureInfo.InvariantCulture);
            vr["from"] = Convert.ToString(TryGet(() => (object)source.Address[false, false], from), CultureInfo.InvariantCulture);
            vr["values"] = true;
            vr["note"] = "copied with Excel's own copy, then formulas were turned into their results: text is not re-read";
            return vr;
        }
        source.Copy(target.Cells[1, 1]);
        try { _app.CutCopyMode = 0; } catch (Exception) { }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = b1Name(target);
        r["at"] = Convert.ToString(TryGet(() => (object)target.Cells[1, 1].Address[false, false], to), CultureInfo.InvariantCulture);
        r["from"] = Convert.ToString(TryGet(() => (object)source.Address[false, false], from), CultureInfo.InvariantCulture);
        r["values"] = false;
        r["note"] = "formulas and formats came along; the source is untouched";
        return r;
    }

    private static string b1Name(dynamic range)
    {
        return Convert.ToString(TryGet(() => (object)range.Worksheet.Name, ""), CultureInfo.InvariantCulture);
    }

    /// <summary>A range that may name its sheet ("Sheet2!A1:C9"); without one it is on the sheet at hand.</summary>
    private static dynamic SheetRef(dynamic wb, dynamic ws, string address)
    {
        if (address == null) throw new Fail("BAD_ARGS", "a range is needed");
        int bang = address.LastIndexOf('!');
        if (bang < 0) return ws.Range(address);
        string sheet = address.Substring(0, bang).Trim('\'', '=');
        dynamic other = null;
        try { other = wb.Worksheets.Item(sheet); } catch (Exception) { other = null; }
        if (other == null) throw new Fail("ANCHOR_MISSING", "there is no sheet \"" + sheet + "\"");
        return other.Range(address.Substring(bang + 1));
    }

    // ---------------------------------------------------------------- add a column

    private static object CmdAddColumn(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "the workbook is read-only; adding a column is refused");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        Block b = BlockOf(ws, a);
        string header = a.Need("header");
        // A column of that name is filled again, not added twice.
        int at = b.Cols + 1;
        bool existing = false;
        for (int c = 1; c <= b.Cols; c++)
            if (string.Equals(b.Headers[c], header, StringComparison.OrdinalIgnoreCase)) { at = c; existing = true; }
        bool inserted = false;
        if (a.Has("after") && !existing)
        {
            at = ColumnOf(b, a.Raw("after")) + 1;
            b.Range.Columns[at].EntireColumn.Insert();
            inserted = true;
        }
        dynamic cells = inserted ? ws.Range(ws.Cells[b.Row0, b.Col0 + at - 1], ws.Cells[b.Row0 + b.Rows - 1, b.Col0 + at - 1]) : b.Range.Columns[at];
        dynamic head = cells.Cells[1, 1];
        head.Value2 = header;
        try { head.Font.Bold = b.Range.Cells[1, 1].Font.Bold; } catch (Exception) { }
        string where = ColLetter((int)head.Column);
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = b.Sheet;
        r["column"] = where;
        r["header"] = header;
        r["filledAgain"] = existing;
        if (b.Rows < 2) { r["note"] = "no data rows to fill"; return r; }
        dynamic body = cells.Cells[2, 1].Resize[b.Rows - 1, 1];
        if (a.Has("formula"))
        {
            string formula = a.Need("formula");
            if (!formula.StartsWith("=", StringComparison.Ordinal)) formula = "=" + formula;
            body.Formula = formula;
        }
        else if (a.Has("value")) body.Value2 = ScalarOf(a.Raw("value"));
        if (a.Has("numberFormat")) { try { body.NumberFormat = a.Need("numberFormat"); } catch (Exception) { } }
        if (a.Flag("values", false)) { try { body.Value2 = body.Value2; } catch (Exception) { } }
        r["range"] = TryGet(() => (object)body.Address[false, false], "");
        List<object> errors = ErrorCells(body);
        r["errors"] = errors;
        if (errors.Count > 0) r["note"] = "some result cells show an error: check the formula and the cells it reads";
        return r;
    }

    // ---------------------------------------------------------------- pivot table

    /// <summary>How many row items a pivot table shows; -1 when the app will not say.</summary>
    private static int RowItemCount(dynamic table)
    {
        try { return (int)table.RowRange.Rows.Count; }
        catch (Exception) { return -1; }
    }

    private const string GroupRecipe =
        "Fallback that always works: excel addcol -Header \"\u6708\u4efd\" -Formula \"=TEXT(<date cell>,\"yyyy-mm\")\" "
        + "(or =YEAR(), or =\"Q\"&ROUNDUP(MONTH(<date cell>)/3,0)), then pivot with -Rows \"\u6708\u4efd\" instead of -Group; "
        + "keep the helper column: the pivot reads it.";

    private static object CmdPivot(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "the workbook is read-only; a pivot table is refused");
        if (a.Has("group"))
        {
            Bag gcheck = new Bag(a.Raw("group"));
            string byCheck = gcheck.Str("by", "month");
            if (byCheck != "month" && byCheck != "quarter" && byCheck != "year")
                throw new Fail("BAD_ARGS", "pivot -Group needs month, quarter or year after the colon (e.g. \"\u65e5\u671f:month\"), not \""
                    + Clean(byCheck) + "\"");
        }
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        dynamic source;
        if (a.Has("source")) source = SheetRef(wb, ws, a.Need("source"));
        else if (a.Has("range")) source = SheetRef(wb, ws, a.Need("range"));
        else source = ws.UsedRange;

        // A pivot table needs a name over every column of its source.
        List<string> heads = new List<string>();
        int srcCols = (int)source.Columns.Count;
        object firstRaw = source.Rows[1].Value2;
        object[,] firstRow = firstRaw as object[,];
        for (int c = 1; c <= srcCols; c++)
        {
            string h = firstRow == null ? (c == 1 ? TextOf(firstRaw) : "") : TextOf(firstRow[1, c]);
            heads.Add(h.Trim());
            if (h.Trim().Length == 0)
                throw new Fail("BAD_ARGS", "the source " + Convert.ToString(TryGet(() => (object)source.Address[false, false], ""), CultureInfo.InvariantCulture)
                    + " has an empty header in column " + ColLetter((int)source.Column + c - 1) + ": a pivot table needs a name over every column");
        }

        dynamic dest;
        string destSheet;
        if (a.Has("to"))
        {
            dynamic to = SheetRef(wb, ws, a.Need("to"));
            dest = to.Cells[1, 1];
            destSheet = b1Name(to);
        }
        else
        {
            dynamic sheets = wb.Worksheets;
            dynamic made = sheets.Add(After: sheets.Item((int)sheets.Count));
            string wanted = a.Str("name", "\u900f\u89c6\u8868");
            try { made.Name = wanted; } catch (Exception) { }
            dest = made.Range["A3"];
            destSheet = Convert.ToString(TryGet(() => (object)made.Name, ""), CultureInfo.InvariantCulture);
        }

        // WPS: Create takes exactly two arguments and CreatePivotTable exactly two (verified 2026-11-02).
        dynamic cache = wb.PivotCaches().Create(1, source);
        dynamic table = cache.CreatePivotTable(dest, a.Str("tableName", a.Str("name", "\u900f\u89c6\u8868")));
        List<string> notes = new List<string>();
        string groupMode = a.Has("group") ? "unknown" : "none";

        Func<object, string> field = delegate(object key)
        {
            string name = TextOf(key).Trim();
            foreach (string head in heads) if (string.Equals(head, name, StringComparison.OrdinalIgnoreCase)) return head;
            throw new Fail("ANCHOR_MISSING", "the source has no column \"" + name + "\". Its headers: " + string.Join(", ", heads.ToArray()) + ".");
        };

        int position = 0;
        foreach (object key in a.List("rows"))
        {
            dynamic f = table.PivotFields(field(key));
            f.Orientation = 1;
            f.Position = ++position;
        }
        position = 0;
        foreach (object key in a.List("columns"))
        {
            dynamic f = table.PivotFields(field(key));
            f.Orientation = 2;
            f.Position = ++position;
        }
        foreach (object key in a.List("filters"))
        {
            dynamic f = table.PivotFields(field(key));
            f.Orientation = 3;
        }

        List<object> values = a.List("values");
        if (values.Count == 0) throw new Fail("BAD_ARGS", "pivot needs -Agg \"field:sum[:caption],field2:count\"");
        List<object> valueFields = new List<object>();
        foreach (object raw in values)
        {
            Bag item = raw is string ? new Bag(new Dictionary<string, object> { { "field", raw } }) : new Bag(raw);
            string name = field(item.Raw("field"));
            string fn = item.Str("fn", "sum").ToLowerInvariant();
            int code = fn == "count" ? -4112 : fn == "average" || fn == "mean" || fn == "avg" ? -4106 : fn == "max" ? -4136 : fn == "min" ? -4139 : -4157;
            string caption = item.Str("name", null);
            if (caption == null)
                caption = (fn == "count" ? "\u8ba1\u6570\uff1a" : fn == "max" ? "\u6700\u5927\uff1a" : fn == "min" ? "\u6700\u5c0f\uff1a"
                    : (fn == "average" || fn == "mean" || fn == "avg") ? "\u5e73\u5747\uff1a" : "\u5408\u8ba1\uff1a") + name;
            dynamic data = table.AddDataField(table.PivotFields(name), caption, code);
            if (item.Has("numberFormat")) { try { data.NumberFormat = item.Need("numberFormat"); } catch (Exception) { } }
            else if (fn != "count") { try { data.NumberFormat = "#,##0.00"; } catch (Exception) { } }
            valueFields.Add(caption);
        }

        if (a.Has("group"))
        {
            // Dates on the rows or columns, gathered by month, quarter or year.
            // WPS 12.1.0.29161 DOES support the native call (measured 2026-11-02: 24 months ->
            // 12 / 4 / 2 row items, totals unchanged), but it returns silently when it cannot group
            // (text dates, mixed types), so the item count is read back instead of trusting it.
            Bag group = new Bag(a.Raw("group"));
            string by = group.Str("by", "month");
            dynamic gf = table.PivotFields(field(group.Raw("field")));
            int itemsBefore = RowItemCount(table);
            try
            {
                dynamic cell = gf.DataRange.Cells[1, 1];
                object[] periods = new object[] { false, false, false, false, by == "month", by == "quarter", by == "year" };
                cell.Group(Start: true, End: true, Periods: periods);
                int itemsAfter = RowItemCount(table);
                if (itemsBefore > 1 && itemsAfter > 0 && itemsAfter >= itemsBefore)
                {
                    groupMode = "failed";
                    notes.Add("the dates were NOT grouped by " + by + " (the call returned without an error, but the pivot still shows "
                        + itemsAfter + " row item(s) instead of fewer): the column probably holds text, not real dates. "
                        + GroupRecipe);
                }
                else groupMode = "native";
            }
            catch (Exception error)
            {
                groupMode = "failed";
                notes.Add("the dates could not be grouped by " + by + " (" + Clean(error.Message)
                    + "): the column must hold real dates only. " + GroupRecipe);
            }
        }
        if (a.Has("sort"))
        {
            Bag sort = new Bag(a.Raw("sort"));
            try
            {
                dynamic f = table.PivotFields(field(sort.Raw("field")));
                string byName = sort.Has("by") ? sort.Need("by") : Convert.ToString(table.DataFields[1].Name, CultureInfo.InvariantCulture);
                f.AutoSort(sort.Str("order", "desc").StartsWith("d", StringComparison.OrdinalIgnoreCase) ? 2 : 1, byName);
            }
            catch (Exception) { notes.Add("the sort was not applied: \"by\" must be the name of a value as it is shown in the table"); }
        }
        try { table.RowAxisLayout(1); } catch (Exception) { }

        dynamic shown2 = table.TableRange1;
        try { shown2.EntireColumn.AutoFit(); } catch (Exception) { }
        // The result goes back with the answer, so it need not be read a second time.
        object[,] grid = ReadValues(shown2);
        List<object> rows = new List<object>();
        int limit = Math.Min(grid.GetLength(0), 40), width = Math.Min(grid.GetLength(1), 14);
        int r0 = (int)shown2.Row;
        for (int r = 1; r <= limit; r++)
        {
            List<object> line = new List<object>();
            line.Add((r0 + r - 1).ToString(CultureInfo.InvariantCulture));
            for (int c = 1; c <= width; c++) line.Add(JsonCellValue(grid[r, c]));
            rows.Add(line);
        }

        Dictionary<string, object> res = new Dictionary<string, object>();
        res["sheet"] = destSheet;
        res["name"] = Convert.ToString(TryGet(() => (object)table.Name, ""), CultureInfo.InvariantCulture);
        res["range"] = Convert.ToString(TryGet(() => (object)shown2.Address[false, false], ""), CultureInfo.InvariantCulture);
        res["rows"] = rows;
        res["valueFields"] = valueFields;
        res["groupMode"] = groupMode;
        if (a.Has("group"))
        {
            Bag g = new Bag(a.Raw("group"));
            Dictionary<string, object> gi = new Dictionary<string, object>();
            gi["field"] = TextOf(g.Raw("field")).Trim();
            gi["by"] = g.Str("by", "month");
            gi["mode"] = groupMode;
            res["group"] = gi;
        }
        res["truncated"] = grid.GetLength(0) > limit || grid.GetLength(1) > width;
        res["notes"] = notes;
        if (notes.Count > 0) res["note"] = string.Join("; ", notes.ToArray());
        return res;
    }

    // ---------------------------------------------------------------- freeze panes

    private static object CmdFreeze(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        string action = a.Str("action", "freeze").ToLowerInvariant();
        dynamic window = null;
        try { window = wb.Windows.Item(1); } catch (Exception) { window = null; }
        if (window == null) throw new Fail("COM_ERROR", "this workbook has no window to freeze");

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);

        if (action == "remove" || action == "clear" || action == "off")
        {
            try { window.FreezePanes = false; } catch (Exception) { }
            try { window.Split = false; } catch (Exception) { }
            r["freezePanes"] = false;
            r["note"] = "panes are no longer frozen";
            return r;
        }

        dynamic cell = ws.Range(a.Str("cell", "A2")).Cells[1, 1];
        int wantRow = (int)cell.Row, wantCol = (int)cell.Column;
        try { ws.Activate(); } catch (Exception) { }
        try { if ((int)window.View != 1) window.View = 1; } catch (Exception) { }
        Exception last = null;
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                try { window.FreezePanes = false; } catch (Exception) { }
                try { window.Split = false; } catch (Exception) { }
                window.ScrollRow = 1;
                window.ScrollColumn = 1;
                window.SplitRow = wantRow - 1;
                window.SplitColumn = wantCol - 1;
                window.FreezePanes = wantRow > 1 || wantCol > 1;
                last = null;
                break;
            }
            catch (Exception error)
            {
                last = error;
                // Excel refuses while its window is not the active one: bring it forward and retry.
                try { window.Activate(); } catch (Exception) { }
                System.Threading.Thread.Sleep(150);
            }
        }
        if (last != null)
            throw new Fail("COM_ERROR", "WPS would not freeze the panes of \"" + Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture)
                + "\" just now (" + Clean(last.Message) + "); the rest of the work can go on, freezing is only a convenience for scrolling");
        r["cell"] = Convert.ToString(TryGet(() => (object)cell.Address[false, false], ""), CultureInfo.InvariantCulture);
        r["splitRow"] = Convert.ToInt32(TryGet(() => (object)(int)window.SplitRow, wantRow - 1), CultureInfo.InvariantCulture);
        r["splitColumn"] = Convert.ToInt32(TryGet(() => (object)(int)window.SplitColumn, wantCol - 1), CultureInfo.InvariantCulture);
        r["freezePanes"] = true;
        return r;
    }

    // ---------------------------------------------------------------- conditional formats

    private static object CmdConditionalFormat(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        dynamic target = ws.Range(a.Need("range"));
        string rule = a.Str("rule", ">").Trim().ToLowerInvariant();
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
        r["range"] = Convert.ToString(TryGet(() => (object)target.Address[false, false], ""), CultureInfo.InvariantCulture);
        r["rule"] = rule;

        if (rule == "clear" || rule == "none")
        {
            target.FormatConditions.Delete();
            r["cleared"] = true;
            return r;
        }

        Func<string, string> eq = delegate(string key)
        {
            object raw = a.Raw(key);
            string v = TextOf(raw);
            if (v.StartsWith("=", StringComparison.Ordinal)) return v;
            return IsNum(raw) ? "=" + v : "=\"" + v.Replace("\"", "\"\"") + "\"";
        };

        dynamic made = null;
        try
        {
            if (rule == "colorscale" || rule == "color scale") { target.FormatConditions.AddColorScale(3); r["count"] = (int)target.FormatConditions.Count; return r; }
            if (rule == "databar" || rule == "data bar") { target.FormatConditions.AddDatabar(); r["count"] = (int)target.FormatConditions.Count; return r; }
            if (rule == "duplicate" || rule == "unique")
            {
                made = target.FormatConditions.AddUniqueValues();
                made.DupeUnique = rule == "duplicate" ? 1 : 0;
            }
            else if (rule == "top" || rule == "bottom")
            {
                made = target.FormatConditions.AddTop10();
                made.TopBottom = rule == "top" ? 1 : 0;
                made.Rank = Math.Max(1, a.Int("value", 10));
            }
            else if (rule == "formula") made = target.FormatConditions.Add(Type: 2, Formula1: eq("value"));
            else if (rule == "blank") made = target.FormatConditions.Add(Type: 10);
            else
            {
                int code = rule == "between" ? 1 : rule == "=" ? 3 : rule == "<>" ? 4 : rule == ">" ? 5
                    : rule == "<" ? 6 : rule == ">=" ? 7 : rule == "<=" ? 8 : 0;
                if (code == 0)
                    throw new Fail("BAD_ARGS", "unknown rule \"" + rule + "\": use >, <, >=, <=, =, <>, between, top, bottom, duplicate, unique, blank, formula, colorscale, databar or clear");
                if (code == 1) made = target.FormatConditions.Add(Type: 1, Operator: code, Formula1: eq("value"), Formula2: eq("value2"));
                else made = target.FormatConditions.Add(Type: 1, Operator: code, Formula1: eq("value"));
            }
        }
        catch (Fail) { throw; }
        catch (Exception error)
        {
            throw new Fail("COM_ERROR", "WPS refused the conditional format \"" + rule + "\" (" + Clean(error.Message) + ")");
        }

        try { made.Interior.Color = ColorFrom(a.Str("fill", a.Str("fillColor", "#FFC7CE"))); } catch (Exception) { }
        try { made.Font.Color = ColorFrom(a.Str("color", "#9C0006")); } catch (Exception) { }
        if (a.Has("bold")) { try { made.Font.Bold = a.Flag("bold", false) ? 1 : 0; } catch (Exception) { } }
        r["count"] = (int)target.FormatConditions.Count;
        return r;
    }

    // ---------------------------------------------------------------- drop-down list

    private static object CmdValidate(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        dynamic target = ws.Range(a.Need("range"));
        List<object> list = a.List("list");
        if (list.Count == 0 && a.Has("list")) list.Add(a.Need("list"));
        if (list.Count == 0) throw new Fail("BAD_ARGS", "validate needs -List \"a,b,c\": the allowed values");
        List<string> items = new List<string>();
        foreach (object item in list) items.Add(TextOf(item));
        string joined = string.Join(",", items.ToArray());
        if (joined.Length > 255)
            throw new Fail("BAD_ARGS", "the list is " + joined.Length.ToString(CultureInfo.InvariantCulture) + " characters; a drop-down list holds at most 255 (write the values to cells and point at them instead)");
        target.Validation.Delete();
        target.Validation.Add(Type: 3, AlertStyle: 1, Operator: 1, Formula1: joined);
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = Convert.ToString(TryGet(() => (object)ws.Name, ""), CultureInfo.InvariantCulture);
        r["range"] = Convert.ToString(TryGet(() => (object)target.Address[false, false], ""), CultureInfo.InvariantCulture);
        r["list"] = items;
        r["note"] = "these cells now take only the listed values";
        return r;
    }

    // ---------------------------------------------------------------- ListObject table

    private static object CmdTable(Bag a)
    {
        dynamic wb = RequireBook(a, true);
        if (IsReadOnly(wb)) throw new Fail("READONLY", "the workbook is read-only; a table is refused");
        dynamic ws = TargetSheet(wb, a);
        if (ws == null) throw new Fail("COM_ERROR", "no target sheet");
        Block b = BlockOf(ws, a);
        dynamic home = ws;
        dynamic table = null;
        int existing = 0;
        try { existing = (int)home.ListObjects.Count; } catch (Exception) { }
        for (int k = 1; k <= existing; k++)
        {
            try
            {
                dynamic one = home.ListObjects.Item(k);
                object hit = _app.Intersect(one.Range, b.Range);
                if (hit != null) { table = one; break; }
            }
            catch (Exception) { }
        }
        bool made = table == null;
        if (made) table = home.ListObjects.Add(1, b.Range, Type.Missing, 1);
        string style = a.Str("style", "TableStyleMedium2");
        try { table.TableStyle = style; } catch (Exception) { }
        if (a.Has("name")) { try { table.Name = a.Need("name"); } catch (Exception error) { throw new Fail("BAD_ARGS", "that table name was refused (" + Clean(error.Message) + "): a table name must be unique, with no spaces"); } }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["sheet"] = b.Sheet;
        r["name"] = Convert.ToString(TryGet(() => (object)table.Name, ""), CultureInfo.InvariantCulture);
        r["range"] = Convert.ToString(TryGet(() => (object)table.Range.Address[false, false], ""), CultureInfo.InvariantCulture);
        r["style"] = style;
        r["created"] = made;
        r["note"] = made
            ? "banded rows and filter arrows; formulas may now use the column names"
            : "these cells were already a table: only its style (and name, if given) changed";
        return r;
    }

    // ---------------------------------------------------------------- render -Overview

    /// <summary>Every sheet of the workbook as one grid of pictures with the sheet names over them.</summary>
    private static object RenderOverviewSheets(dynamic wb, Bag a, string outDir, int width)
    {
        List<dynamic> sheets = new List<dynamic>();
        List<string> names = new List<string>();
        string wanted = a.Str("sheets", a.Str("sheet", null));
        if (wanted != null && wanted.Trim().Length > 0)
        {
            foreach (string piece in wanted.Split(','))
            {
                string one = piece.Trim();
                if (one.Length == 0) continue;
                dynamic found = null;
                try { found = wb.Worksheets.Item(one); } catch (Exception) { found = null; }
                if (found == null) throw new Fail("ANCHOR_MISSING", "there is no sheet \"" + one + "\"");
                sheets.Add(found);
                names.Add(one);
            }
        }
        else
        {
            int count = (int)wb.Worksheets.Count;
            for (int s = 1; s <= count && s <= 12; s++)
            {
                dynamic one = wb.Worksheets.Item(s);
                sheets.Add(one);
                names.Add(Convert.ToString(TryGet(() => (object)one.Name, "Sheet" + s.ToString(CultureInfo.InvariantCulture)), CultureInfo.InvariantCulture));
            }
        }
        if (sheets.Count == 0) throw new Fail("COM_ERROR", "this workbook has no sheet to show");

        int cellW = Math.Max(420, Math.Min(900, width / 2));
        List<Bitmap> cells = new List<Bitmap>();
        List<string> cellNames = new List<string>();
        List<object> entries = new List<object>();
        string scratchDir = Path.Combine(outDir, "overview-parts");
        Directory.CreateDirectory(scratchDir);
        try
        {
            for (int k = 0; k < sheets.Count; k++)
            {
                dynamic one = sheets[k];
                dynamic used = one.UsedRange;
                string address = Convert.ToString(TryGet(() => (object)used.Address[false, false], "A1"), CultureInfo.InvariantCulture);
                string part = Path.Combine(scratchDir, "sheet-" + (k + 1).ToString(CultureInfo.InvariantCulture) + ".png");
                bool ok = false;
                try { ok = TryRenderRange(used, part); } catch (Exception) { ok = false; }
                if (!ok) continue;
                Bitmap bmp = null;
                try { using (Image img = Image.FromFile(part)) bmp = new Bitmap(img, cellW, Math.Max(40, (int)Math.Round((double)img.Height * cellW / img.Width))); }
                catch (Exception) { bmp = null; }
                if (bmp == null) continue;
                cells.Add(bmp);
                cellNames.Add(names[k]);
                Dictionary<string, object> e = new Dictionary<string, object>();
                e["sheet"] = names[k];
                e["range"] = address;
                entries.Add(e);
            }
            if (cells.Count == 0)
                throw new Fail("COM_ERROR", "no sheet could be turned into a picture: every sheet looks empty");

            int pad = 20, label = 30, cols = 2;
            int rows = (cells.Count + cols - 1) / cols;
            int cellH = 0;
            foreach (Bitmap bmp in cells) if (bmp.Height > cellH) cellH = bmp.Height;
            int sheetW = cols * cellW + (cols + 1) * pad;
            int sheetH = rows * (cellH + label) + (rows + 1) * pad;
            string path = Path.Combine(outDir, "overview-" + cells.Count.ToString(CultureInfo.InvariantCulture) + "sheets.png");
            using (Bitmap sheet = new Bitmap(sheetW, sheetH))
            using (Graphics g = Graphics.FromImage(sheet))
            using (Font f = new Font("\u5fae\u8f6f\u96c5\u9ed1", 13f))
            using (SolidBrush ink = new SolidBrush(Color.FromArgb(45, 45, 45)))
            using (Pen border = new Pen(Color.FromArgb(205, 205, 205)))
            {
                g.Clear(Color.FromArgb(246, 246, 246));
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                for (int k = 0; k < cells.Count; k++)
                {
                    int col = k % cols, row = k / cols;
                    int x = pad + col * (cellW + pad);
                    int y = pad + row * (cellH + label + pad);
                    g.DrawString(cellNames[k], f, ink, x, y);
                    g.DrawImage(cells[k], x, y + label, cellW, cells[k].Height);
                    g.DrawRectangle(border, x, y + label, cellW, cells[k].Height);
                }
                sheet.Save(path, ImageFormat.Png);
            }
            Dictionary<string, object> r = new Dictionary<string, object>();
            r["overview"] = path;
            r["sheets"] = entries;
            r["count"] = cells.Count;
            try { using (Image img = Image.FromFile(path)) { r["width"] = img.Width; r["height"] = img.Height; } } catch (Exception) { }
            r["note"] = "one picture with every sheet as a thumbnail; render -Range on one sheet for the legible version";
            return r;
        }
        finally
        {
            foreach (Bitmap bmp in cells) bmp.Dispose();
            try { Directory.Delete(scratchDir, true); } catch (Exception) { }
        }
    }
}
