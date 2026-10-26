// v1.8 captions (2026-10-26): the missing "list" capability behind a table-numbering check.
//
// Field report: checking whether table captions are numbered consistently on a real 75-page report
// took ~6 minutes - 22 `read` calls (4.5 min) to pair 32 tables with 32 captions - while the
// rewrite itself took 2 seconds. The reason was structural: no command produced a list. `table -Action
// scan` had rows/pages/first-row text but no paragraph number, `audit` only reported a paragraph
// number when a caption was MISSING, and `search` mis-addressed paragraphs. So the caller had to
// sweep every paragraph (9000+ of the 10521 were table cells) to pair tables with captions.
//
// This command does that pairing directly and answers "are the table numbers any good?":
// duplicates, gaps inside one section, mixed dot/hyphen formats, mixed spacing after the number,
// captions inside the table's first row vs above it, captions below their table, numbering out of
// document order, and dangling in-text references ("见表3.3-2" pointing at a table that is gone).
//
//   captions {"doc":"...", "label":"\u8868", "max":200, "locate":true}
//
// Cost: one roundtrip per table plus a few paragraphs around it, and ONE Content.Text fetch for the
// reference scan - independent of the paragraph count, so a 10k-paragraph report answers in seconds.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

internal static partial class Program
{
    /// <summary>"表2.4.1  峡江站天然径流成果表" split into its parts.</summary>
    private sealed class CaptionInfo
    {
        public string Label = "";   // 表 / 附表 / 图 / 附图
        public string Number = "";  // 2.4.1 / 3.3-1 / 1
        public string Gap = "";     // whitespace between number and title, verbatim
        public string Title = "";
        public string Format = "";  // dot / hyphen / plain
        public string Text = "";
    }

    private sealed class CaptionRec
    {
        public int Table;           // 0 when the caption is not tied to a table
        public int Para;
        public string Pos = "";     // above / in-table / below
        public string Style = "";
        public CaptionInfo Info;
    }

    /// <summary>
    /// "表 2.4.1  xxx" / "附表3.3-1　xxx" / "图4-1 xxx". The label must be the FIRST thing in the
    /// paragraph (only whitespace may precede it) - that is what separates a caption from a
    /// sentence that merely mentions a table number ("如表3.3-2所示").
    /// </summary>
    private static readonly Regex CaptionRx = new Regex(
        "^[ \u3000\t]*(?<label>附表|附图|表|图)[ \u3000\t]*(?<num>[0-9]+(?:[.\\-\uff0d\u2013\u2014\uff0e][0-9]+)*)(?<gap>[ \u3000\t]*)(?<title>[^\\r\\n]*)$",
        RegexOptions.Compiled);

    /// <summary>Every "表N..." / "图N..." mention anywhere in a line (used for the reference scan).</summary>
    private static readonly Regex RefRx = new Regex(
        "(?<label>附表|附图|表|图)[ \u3000]*(?<num>[0-9]+(?:[.\\-\uff0d\u2013\u2014\uff0e][0-9]+)*)",
        RegexOptions.Compiled);

    private static CaptionInfo ParseCaption(string raw)
    {
        if (raw == null) return null;
        string text = CleanPara(raw).Trim();
        if (text.Length < 2) return null;
        Match m = CaptionRx.Match(text);
        if (!m.Success) return null;
        CaptionInfo c = new CaptionInfo();
        c.Label = m.Groups["label"].Value;
        c.Number = m.Groups["num"].Value;
        c.Gap = m.Groups["gap"].Value;
        c.Title = m.Groups["title"].Value.Trim();
        c.Text = text;
        c.Format = CaptionFormat(c.Number);
        if (c.Title.Length == 0 && c.Number.Length == 0) return null;
        return c;
    }

    /// <summary>dot (2.4.1) / hyphen (3.3-1) / plain (表1) - the separator before the last number.</summary>
    private static string CaptionFormat(string number)
    {
        for (int i = number.Length - 1; i >= 0; i--)
        {
            char ch = number[i];
            if (ch == '.') return "dot";
            if (ch == '-' || ch == '\uff0d' || ch == '\u2013' || ch == '\u2014') return "hyphen";
            if (ch == '\uff0e') return "dot";
        }
        return "plain";
    }

    private static int[] NumberParts(string number)
    {
        List<int> parts = new List<int>();
        int cur = -1;
        foreach (char ch in number)
        {
            if (ch >= '0' && ch <= '9')
            {
                if (cur < 0) cur = 0;
                cur = cur * 10 + (ch - '0');
            }
            else
            {
                if (cur >= 0) { parts.Add(cur); cur = -1; }
            }
        }
        if (cur >= 0) parts.Add(cur);
        return parts.ToArray();
    }

    /// <summary>"3.3-2" -&gt; "3.3" (everything but the last component).</summary>
    private static string NumberPrefix(string number)
    {
        int cut = -1;
        for (int i = number.Length - 1; i >= 0; i--)
        {
            char ch = number[i];
            if (ch == '.' || ch == '-' || ch == '\uff0d' || ch == '\u2013' || ch == '\u2014' || ch == '\uff0e') { cut = i; break; }
        }
        return cut > 0 ? number.Substring(0, cut) : "";
    }

    /// <summary>Compare two numbers component-wise; used for document-order and gap checks.</summary>
    private static int CompareNumbers(string a, string b)
    {
        int[] pa = NumberParts(a), pb = NumberParts(b);
        int n = Math.Min(pa.Length, pb.Length);
        for (int i = 0; i < n; i++)
        {
            if (pa[i] != pb[i]) return pa[i] < pb[i] ? -1 : 1;
        }
        if (pa.Length == pb.Length) return 0;
        return pa.Length < pb.Length ? -1 : 1;
    }

    private static int ParaStartAt(dynamic doc, int k)
    {
        try { return (int)doc.Paragraphs.Item(k).Range.Start; }
        catch (Exception) { return -1; }
    }

    // ParaTextAt(doc, k) is shared with Tables.cs (same partial class).

    /// <summary>
    /// Paragraph number that really contains a character offset. Range(0,offset).Paragraphs.Count is
    /// off by one at boundaries (same reason the paragraph mapping in search needed fixing), so the
    /// candidates are resolved by span - the paragraph whose [Start,End) contains the offset.
    /// </summary>
    private static int ParaAtOffset(dynamic doc, int offset, int total)
    {
        if (offset < 0) return 0;
        int c = (int)TryGet(() => (object)(int)doc.Range(0, offset).Paragraphs.Count, 0);
        int[] probes = new int[] { c, c + 1, c + 2, c - 1 };
        foreach (int cand in probes)
        {
            if (cand < 1 || cand > total) continue;
            int[] sp = ParaSpan(doc, cand);
            if (sp == null) continue;
            if (offset >= sp[0] && offset < sp[1]) return cand;
        }
        return c >= 1 && c <= total ? c : 0;
    }

    /// <summary>Paragraph number of the table's first cell paragraph.</summary>
    private static int TableStartPara(dynamic doc, dynamic table, int total)
    {
        int start = (int)TryGet(() => (object)(int)table.Range.Start, -1);
        if (start < 0) return 0;
        return ParaAtOffset(doc, start, total);
    }

    /// <summary>Paragraph number of the last paragraph inside the table (its last cell's paragraph).</summary>
    private static int TableEndPara(dynamic doc, dynamic table, int total)
    {
        int end = (int)TryGet(() => (object)(int)table.Range.End, -1);
        if (end <= 1) return 0;
        return ParaAtOffset(doc, end - 1, total);
    }

    /// <summary>The caption paragraph above a table: look back up to 4 paragraphs, skipping blanks.</summary>
    private static CaptionRec CaptionAbove(dynamic doc, int startPara, HashSet<int> captionParas)
    {
        int blanks = 0;
        for (int k = startPara - 1; k >= 1 && k >= startPara - 4; k--)
        {
            string text = ParaTextAt(doc, k);
            if (text.Trim().Length == 0) { blanks++; if (blanks > 2) return null; continue; }
            if (captionParas.Contains(k)) return null;             // belongs to the previous table
            CaptionInfo info = ParseCaption(text);
            if (info == null) return null;                          // a real paragraph - stop looking
            CaptionRec rec = new CaptionRec();
            rec.Table = 0;
            rec.Para = k;
            rec.Pos = "above";
            rec.Style = Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(k).Style.NameLocal, ""), CultureInfo.InvariantCulture);
            rec.Info = info;
            return rec;
        }
        return null;
    }

    /// <summary>The caption paragraph below a table (a caption below its table is the wrong side).</summary>
    private static CaptionRec CaptionBelow(dynamic doc, int endPara, HashSet<int> captionParas)
    {
        int blanks = 0;
        for (int k = endPara + 1; k <= endPara + 4; k++)
        {
            string text = ParaTextAt(doc, k);
            if (text.Trim().Length == 0) { blanks++; if (blanks > 2) return null; continue; }
            if (captionParas.Contains(k)) return null;
            CaptionInfo info = ParseCaption(text);
            if (info == null) return null;
            CaptionRec rec = new CaptionRec();
            rec.Para = k;
            rec.Pos = "below";
            rec.Style = Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(k).Style.NameLocal, ""), CultureInfo.InvariantCulture);
            rec.Info = info;
            return rec;
        }
        return null;
    }

    private static bool LabelMatches(string wanted, string actual)
    {
        if (wanted == null || wanted.Length == 0) return true;
        if (actual == null) return false;
        if (actual == wanted) return true;
        // 表 also accepts 附表, 图 also accepts 附图
        if (wanted == "\u8868") return actual == "\u9644\u8868";
        if (wanted == "\u56fe") return actual == "\u9644\u56fe";
        return false;
    }

    private static object CaptionsDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        int total = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        string wanted = a.Str("label", "\u8868");
        if (wanted == null || wanted.Trim().Length == 0) wanted = "\u8868";
        wanted = wanted.Trim();
        int maxIssues = a.Int("max", 200);
        if (maxIssues <= 0) maxIssues = 200;

        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);

        List<CaptionRec> caps = new List<CaptionRec>();
        List<object> tableRows = new List<object>();
        List<int> uncaptioned = new List<int>();
        HashSet<int> captionParas = new HashSet<int>();
        int above = 0, inTable = 0, below = 0;

        for (int n = 1; n <= tables; n++)
        {
            Tr("captions table " + n);
            dynamic table = doc.Tables.Item(n);
            int startPara = TableStartPara(doc, table, total);
            int endPara = TableEndPara(doc, table, total);

            CaptionRec rec = null;

            // 1) the title typed into the first row (merged cell) - very common in Chinese reports
            string firstCell = "";
            try { firstCell = CleanPara(Convert.ToString(table.Cell(1, 1).Range.Text, CultureInfo.InvariantCulture)); }
            catch (Exception) { }
            CaptionInfo head = ParseCaption(firstCell);
            if (head != null && LabelMatches(wanted, head.Label))
            {
                rec = new CaptionRec();
                rec.Para = startPara;
                rec.Pos = "in-table";
                rec.Style = Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(startPara).Style.NameLocal, ""), CultureInfo.InvariantCulture);
                rec.Info = head;
            }
            // 2) the paragraph above the table
            if (rec == null)
            {
                rec = CaptionAbove(doc, startPara, captionParas);
                if (rec != null && !LabelMatches(wanted, rec.Info.Label)) rec = null;
            }
            // 3) below the table: report it, it is the wrong side
            if (rec == null)
            {
                rec = CaptionBelow(doc, endPara > 0 ? endPara : startPara, captionParas);
                if (rec != null && !LabelMatches(wanted, rec.Info.Label)) rec = null;
            }

            if (rec == null)
            {
                uncaptioned.Add(n);
            }
            else
            {
                rec.Table = n;
                caps.Add(rec);
                if (rec.Pos == "above") above++;
                else if (rec.Pos == "in-table") inTable++;
                else below++;
                if (rec.Para > 0) captionParas.Add(rec.Para);
            }

            Dictionary<string, object> row = new Dictionary<string, object>();
            row["table"] = n;
            row["para"] = startPara;
            row["rows"] = TryGet(() => (object)(int)table.Rows.Count, 0);
            row["columns"] = TryGet(() => (object)(int)table.Columns.Count, 0);
            if (rec == null) row["caption"] = null;
            else
            {
                Dictionary<string, object> cap = new Dictionary<string, object>();
                cap["para"] = rec.Para;
                cap["pos"] = rec.Pos;
                cap["label"] = rec.Info.Label;
                cap["number"] = rec.Info.Number;
                cap["format"] = rec.Info.Format;
                cap["gap"] = rec.Info.Gap;
                cap["title"] = rec.Info.Title;
                if (rec.Style != null && rec.Style.Length > 0) cap["style"] = rec.Style;
                row["caption"] = cap;
            }
            tableRows.Add(row);
        }

        // ---------------------------------------------------------------- numbering checks
        List<object> issues = new List<object>();
        bool truncated = false;

        // duplicates: same number twice
        Dictionary<string, List<CaptionRec>> byNumber = new Dictionary<string, List<CaptionRec>>();
        foreach (CaptionRec c in caps)
        {
            List<CaptionRec> list;
            if (!byNumber.TryGetValue(c.Info.Number, out list)) { list = new List<CaptionRec>(); byNumber[c.Info.Number] = list; }
            list.Add(c);
        }
        List<object> duplicates = new List<object>();
        foreach (KeyValuePair<string, List<CaptionRec>> kv in byNumber)
        {
            if (kv.Value.Count < 2) continue;
            List<object> paras = new List<object>();
            List<object> tbls = new List<object>();
            foreach (CaptionRec c in kv.Value) { paras.Add(c.Para); tbls.Add(c.Table); }
            Dictionary<string, object> dup = new Dictionary<string, object>();
            dup["number"] = kv.Key;
            dup["paras"] = paras;
            dup["tables"] = tbls;
            duplicates.Add(dup);
            AddIssue(issues, ref truncated, maxIssues, "caption-duplicate", kv.Value[1].Para, kv.Key,
                Head(kv.Value[1].Info.Text, 60),
                "\u540c\u4e00\u8868\u53f7\u7528\u4e86 " + kv.Value.Count + " \u6b21\uff08\u6bb5 " + JoinInts(paras) + "\uff09\uff1a\u6539\u6389\u540e\u9762\u90a3\u5f20\u8868\u7684\u7f16\u53f7\uff0c\u5e76\u7528 fields \u5237\u65b0\u57df\u5f0f\u9898\u6ce8");
        }

        // gaps inside one section: 2.4.1 -> 2.4.3
        List<object> gaps = new List<object>();
        Dictionary<string, List<CaptionRec>> byPrefix = new Dictionary<string, List<CaptionRec>>();
        List<string> prefixOrder = new List<string>();
        foreach (CaptionRec c in caps)
        {
            string prefix = NumberPrefix(c.Info.Number);
            List<CaptionRec> list;
            if (!byPrefix.TryGetValue(prefix, out list)) { list = new List<CaptionRec>(); byPrefix[prefix] = list; prefixOrder.Add(prefix); }
            list.Add(c);
        }
        foreach (string prefix in prefixOrder)
        {
            List<CaptionRec> list = byPrefix[prefix];
            if (list.Count < 2) continue;
            for (int i = 1; i < list.Count; i++)
            {
                int[] prev = NumberParts(list[i - 1].Info.Number);
                int[] cur = NumberParts(list[i].Info.Number);
                if (prev.Length == 0 || cur.Length == 0) continue;
                if (prev.Length != cur.Length) continue;
                int delta = cur[cur.Length - 1] - prev[prev.Length - 1];
                if (delta == 1) continue;
                if (delta <= 0) continue;                           // 0 = duplicate (own finding), <0 = out of order (own finding)
                List<object> missing = new List<object>();
                for (int v = prev[prev.Length - 1] + 1; v < cur[cur.Length - 1] && missing.Count < 8; v++)
                {
                    string before = list[i - 1].Info.Number;
                    int cut = before.Length;
                    for (int j = before.Length - 1; j >= 0; j--)
                    {
                        char ch = before[j];
                        if (ch >= '0' && ch <= '9') { cut = j; }
                        else break;
                    }
                    missing.Add(before.Substring(0, cut) + v.ToString(CultureInfo.InvariantCulture));
                }
                Dictionary<string, object> gap = new Dictionary<string, object>();
                gap["prefix"] = prefix;
                gap["after"] = list[i - 1].Info.Number;
                gap["before"] = list[i].Info.Number;
                gap["missing"] = missing;
                gap["afterPara"] = list[i - 1].Para;
                gap["beforePara"] = list[i].Para;
                gaps.Add(gap);
                AddIssue(issues, ref truncated, maxIssues, "caption-gap", list[i].Para, list[i].Info.Number,
                    "\u524d\u4e00\u5f20 " + list[i - 1].Info.Number + " \u2192 " + list[i].Info.Number,
                    "\u540c\u4e00\u8282\u5185\u8868\u53f7\u8df3\u53f7\uff08\u7f3a " + JoinInts(missing) + "\uff09\uff1a\u786e\u8ba4\u662f\u5426\u6f0f\u4e86\u8868\uff0c\u6216\u8868\u53f7\u5199\u9519");
            }
        }

        // mixed number formats: 2.4.1 vs 3.3-1
        Dictionary<string, int> formatCount = new Dictionary<string, int>();
        Dictionary<string, List<object>> formatSamples = new Dictionary<string, List<object>>();
        foreach (CaptionRec c in caps)
        {
            int n;
            formatCount.TryGetValue(c.Info.Format, out n);
            formatCount[c.Info.Format] = n + 1;
            List<object> samples;
            if (!formatSamples.TryGetValue(c.Info.Format, out samples)) { samples = new List<object>(); formatSamples[c.Info.Format] = samples; }
            if (samples.Count < 6) samples.Add(c.Info.Number + " (\u6bb5 " + c.Para + ")");
        }
        List<object> formats = new List<object>();
        foreach (KeyValuePair<string, int> kv in formatCount)
        {
            Dictionary<string, object> f = new Dictionary<string, object>();
            f["format"] = kv.Key;
            f["count"] = kv.Value;
            f["samples"] = formatSamples[kv.Key];
            formats.Add(f);
        }
        if (formatCount.Count > 1)
        {
            AddIssue(issues, ref truncated, maxIssues, "caption-format-mixed", 0, null,
                "\u5e76\u5b58\u7684\u8868\u53f7\u683c\u5f0f\uff1a" + DescribeFormats(formatCount),
                "\u540c\u4e00\u4efd\u62a5\u544a\u91cc\u8868\u53f7\u683c\u5f0f\u8981\u7edf\u4e00\uff08\u70b9\u5206 2.4.1 \u6216\u77ed\u6a2a 3.3-1 \u4e8c\u9009\u4e00\uff0c\u4e00\u822c\u8ddf\u672c\u9879\u76ee\u6a21\u677f\u6216\u7b2c\u4e00\u7ae0\u4f53\u4f8b\uff09");
        }

        // spacing between the number and the title: "" / "  " / "\u3000" are style variants worth
        // unifying; a run of >6 spaces is not a style at all but the author aligning the title by
        // hand (measured on a real report: 17-20 spaces), which needs tabs/paragraph alignment.
        Dictionary<string, int> gapCount = new Dictionary<string, int>();
        int alignGaps = 0;
        foreach (CaptionRec c in caps)
        {
            string gap = c.Info.Gap ?? "";
            if (gap.Length > 6) { alignGaps++; continue; }
            string key = gap.Length == 0 ? "" : (gap.Replace(' ', 'S').Replace('\u3000', 'F').Replace('\t', 'T') + "*" + gap.Length);
            int n;
            gapCount.TryGetValue(key, out n);
            gapCount[key] = n + 1;
        }
        List<object> gapsSpacing = new List<object>();
        foreach (KeyValuePair<string, int> kv in gapCount)
        {
            Dictionary<string, object> g = new Dictionary<string, object>();
            g["gap"] = kv.Key;
            g["count"] = kv.Value;
            gapsSpacing.Add(g);
        }
        if (alignGaps > 0)
        {
            Dictionary<string, object> g = new Dictionary<string, object>();
            g["gap"] = "(\u7528 >6 \u4e2a\u7a7a\u683c\u624b\u5de5\u5bf9\u9f50)";
            g["count"] = alignGaps;
            gapsSpacing.Add(g);
        }
        if (gapCount.Count > 1)
        {
            AddIssue(issues, ref truncated, maxIssues, "caption-space-mixed", 0, null,
                "\u9898\u6ce8\u8868\u53f7\u4e0e\u9898\u540d\u4e4b\u95f4\u7684\u7a7a\u683c\u4e0d\u4e00\u81f4\uff1a" + DescribeGaps(gapCount),
                "\u7edf\u4e00\u4e3a\u4e24\u4e2a\u534a\u89d2\u7a7a\u683c\u6216\u4e00\u4e2a\u5168\u89d2\u7a7a\u683c\uff1bformat -Match \u53ef\u6309\u6b63\u5219\u6279\u91cf\u6539\u8fd9\u4e9b\u6bb5");
        }
        if (alignGaps > 0)
        {
            AddIssue(issues, ref truncated, maxIssues, "caption-space-align", 0, null,
                alignGaps + " \u5f20\u9898\u6ce8\u7528\u5927\u91cf\u7a7a\u683c\u628a\u9898\u540d\u9876\u5230\u53f3\u8fb9",
                "\u522b\u7528\u7a7a\u683c\u5bf9\u9f50\uff1a\u6539\u7528\u5236\u8868\u4f4d\uff08Tab\uff09\u6216\u6bb5\u843d\u5bf9\u9f50\uff0c\u5426\u5219\u6362\u5b57\u4f53/\u6539\u5b57\u53f7\u540e\u5fc5\u5b9a\u9519\u4f4d");
        }

        // captions below their table (figure captions belong below, table captions do not)
        foreach (CaptionRec c in caps)
        {
            if (c.Pos != "below") continue;
            AddIssue(issues, ref truncated, maxIssues, "caption-below", c.Para, c.Info.Number, Head(c.Info.Text, 60),
                "\u9898\u6ce8\u5728\u8868\u683c\u4e0b\u65b9\uff08\u8868\u9898\u5e94\u5728\u8868\u4e0a\u65b9\uff09\uff1a\u628a\u8fd9\u6bb5\u79fb\u5230\u8868\u683c\u4e0a\u65b9");
        }

        // mixed caption placement (inside the first row vs above the table)
        if (above > 0 && inTable > 0)
        {
            AddIssue(issues, ref truncated, maxIssues, "caption-position-mixed", 0, null,
                inTable + " \u5f20\u9898\u540d\u5728\u8868\u5185\u7b2c 1 \u884c\uff0c" + above + " \u5f20\u5728\u8868\u5916\u4e0a\u65b9",
                "\u540c\u4e00\u4efd\u6587\u6863\u9898\u6ce8\u4f4d\u7f6e\u8981\u4e00\u81f4\uff1b\u505a\u8fdb\u8868\u683c\u7b2c 1 \u884c\u7684\u7528 table -Action caption \u63d0\u5230\u6846\u5916\uff08\u4f1a\u987a\u5e26\u52a0\u7c97\u5916\u6846\uff09");
        }

        // Numbering out of document order - but not for a duplicated number: that one already has its
        // own finding, and "3.3-4 followed by 3.3-3" is the duplicate talking, not a second problem.
        foreach (string prefix in prefixOrder)
        {
            List<CaptionRec> list = byPrefix[prefix];
            for (int i = 1; i < list.Count; i++)
            {
                if (byNumber.ContainsKey(list[i].Info.Number) && byNumber[list[i].Info.Number].Count > 1) continue;
                if (CompareNumbers(list[i - 1].Info.Number, list[i].Info.Number) > 0)
                {
                    AddIssue(issues, ref truncated, maxIssues, "caption-order", list[i].Para, list[i].Info.Number,
                        list[i - 1].Info.Number + " \u540e\u9762\u8ddf\u7740 " + list[i].Info.Number,
                        "\u8868\u53f7\u672a\u6309\u6587\u6863\u987a\u5e8f\u9012\u589e\uff1a\u786e\u8ba4\u662f\u4e0d\u662f\u8868\u683c\u987a\u5e8f\u8c03\u6574\u540e\u5fd8\u4e86\u6539\u53f7");
                }
            }
        }

        // tables with no caption at all
        foreach (int n in uncaptioned)
        {
            Dictionary<string, object> t = null;
            foreach (object o in tableRows) { Dictionary<string, object> d = (Dictionary<string, object>)o; if (Convert.ToInt32(d["table"], CultureInfo.InvariantCulture) == n) { t = d; break; } }
            int p = t != null ? Convert.ToInt32(t["para"], CultureInfo.InvariantCulture) : 0;
            AddIssue(issues, ref truncated, maxIssues, "caption-missing", p, null,
                "\u8868\u683c " + n + "\uff08\u8d77\u59cb\u6bb5 " + p + "\uff09\u8868\u4e0a\u65b9\u3001\u8868\u5185\u5747\u65e0\u9898\u6ce8",
                "\u8865\u201c\u8868 N\u3000\u9898\u540d\u201d\u5728\u8868\u683c\u4e0a\u65b9\uff1b\u82e5\u8868\u540d\u505a\u5728\u7b2c 1 \u884c\uff0c\u7528 table -Action caption -Table " + n + " \u63d0\u51fa\u6765");
        }

        // ---------------------------------------------------------------- reference scan (one Content.Text)
        List<object> references = new List<object>();
        List<object> unresolved = new List<object>();
        int refTotal = 0;
        string all = WholeText(doc);
        if (all != null)
        {
            Dictionary<string, int> refCount = new Dictionary<string, int>();
            List<string> refOrder = new List<string>();
            foreach (string rawLine in all.Split('\r'))
            {
                string line = rawLine;
                if (line == null) continue;
                string trimmed = line.Trim();
                if (trimmed.Length == 0) continue;
                CaptionInfo self = ParseCaption(trimmed);
                bool captionLine = self != null && LabelMatches(wanted, self.Label);
                int lead = 0;
                while (lead < line.Length && (line[lead] == ' ' || line[lead] == '\u3000' || line[lead] == '\t')) lead++;
                foreach (Match m in RefRx.Matches(line))
                {
                    string lb = m.Groups["label"].Value;
                    if (!LabelMatches(wanted, lb)) continue;
                    string num = m.Groups["num"].Value;
                    // the caption's own label+number is not a reference
                    if (captionLine && m.Index == lead && lb == self.Label && num == self.Number) continue;
                    int c0;
                    refCount.TryGetValue(num, out c0);
                    if (c0 == 0) refOrder.Add(num);
                    refCount[num] = c0 + 1;
                    refTotal++;
                }
            }
            refOrder.Sort(delegate(string x, string y) { return CompareNumbers(x, y); });
            foreach (string num in refOrder)
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["number"] = num;
                one["count"] = refCount[num];
                one["resolved"] = byNumber.ContainsKey(num);
                if (byNumber.ContainsKey(num))
                {
                    List<object> paras = new List<object>();
                    foreach (CaptionRec c in byNumber[num]) paras.Add(c.Para);
                    one["captionParas"] = paras;
                }
                references.Add(one);
                if (!byNumber.ContainsKey(num))
                {
                    unresolved.Add(num);
                    AddIssue(issues, ref truncated, maxIssues, "ref-dangling", 0, num,
                        "\u6b63\u6587\u5f15\u7528\u4e86 " + num + "\uff08" + refCount[num] + " \u5904\uff09\uff0c\u4f46\u6587\u6863\u91cc\u6ca1\u6709\u8fd9\u5f20\u8868\u7684\u9898\u6ce8",
                        "\u6838\u5bf9\u5f15\u7528\u4e0e\u8868\u53f7\uff1a\u6539\u6210\u5b9e\u9645\u5b58\u5728\u7684\u8868\u53f7\uff0c\u6216\u8865\u4e0a\u88ab\u5f15\u7528\u7684\u8868\u683c/\u9898\u6ce8");
                }
            }
        }

        // ---------------------------------------------------------------- result
        List<object> captionList = new List<object>();
        foreach (CaptionRec c in caps)
        {
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["number"] = c.Info.Number;
            one["para"] = c.Para;
            one["pos"] = c.Pos;
            if (c.Table > 0) one["table"] = c.Table;
            one["label"] = c.Info.Label;
            one["format"] = c.Info.Format;
            one["gap"] = c.Info.Gap;
            one["title"] = c.Info.Title;
            if (c.Style != null && c.Style.Length > 0) one["style"] = c.Style;
            captionList.Add(one);
        }

        Dictionary<string, object> positions = new Dictionary<string, object>();
        positions["above"] = above;
        positions["in-table"] = inTable;
        positions["below"] = below;
        positions["none"] = uncaptioned.Count;

        Dictionary<string, object> summary = new Dictionary<string, object>();
        summary["tables"] = tables;
        summary["captions"] = caps.Count;
        summary["positions"] = positions;
        summary["duplicates"] = duplicates;
        summary["sequenceGaps"] = gaps;
        summary["formats"] = formats;
        summary["spaces"] = gapsSpacing;
        summary["uncaptioned"] = uncaptioned;
        summary["references"] = refTotal;
        summary["unresolvedReferences"] = unresolved;

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["label"] = wanted;
        r["paragraphs"] = total;
        r["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        r["summary"] = summary;
        r["tables"] = tableRows;
        r["captions"] = captionList;
        r["references"] = references;
        r["issues"] = issues;
        r["count"] = issues.Count;
        r["truncated"] = truncated;
        r["saved"] = false;
        r["note"] = "\u53ea\u8bfb\u3002caption.pos = above\uff08\u8868\u4e0a\u65b9\uff09/ in-table\uff08\u505a\u5728\u7b2c 1 \u884c\uff09/ below\uff08\u5728\u8868\u4e0b\u65b9\uff0c\u5e94\u4e0a\u79fb\uff09\uff1bpara \u662f\u9898\u6ce8\u6bb5\u8bb8\u53ef\u76f4\u63a5\u5582\u7ed9 revise/comment/format\u3002issues \u91cc\u7684 fix \u5efa\u8bae\u4e0d\u81ea\u52a8\u6267\u884c\uff1a\u8868\u53f7\u91cd\u53f7/\u8df3\u53f7\u5f80\u5f80\u8981\u8ddf\u4f5c\u8005\u786e\u8ba4\u6539\u54ea\u4e00\u4e2a\u3002";
        return r;
    }

    private static string JoinInts(List<object> values)
    {
        List<string> parts = new List<string>();
        foreach (object o in values) parts.Add(Convert.ToString(o, CultureInfo.InvariantCulture));
        return string.Join(", ", parts.ToArray());
    }

    private static string DescribeFormats(Dictionary<string, int> counts)
    {
        List<string> parts = new List<string>();
        foreach (KeyValuePair<string, int> kv in counts)
            parts.Add(kv.Key + "=" + kv.Value.ToString(CultureInfo.InvariantCulture) + "\u5f20");
        return string.Join("\uff0c", parts.ToArray());
    }

    private static string DescribeGaps(Dictionary<string, int> counts)
    {
        List<string> parts = new List<string>();
        foreach (KeyValuePair<string, int> kv in counts)
            parts.Add(kv.Key + "=" + kv.Value.ToString(CultureInfo.InvariantCulture) + "\u5f20");
        return string.Join("\uff0c", parts.ToArray());
    }

    private static void AddIssue(List<object> issues, ref bool truncated, int max, string rule, int para, string number, string sample, string suggest)
    {
        if (issues.Count >= max) { truncated = true; return; }
        Dictionary<string, object> one = new Dictionary<string, object>();
        one["rule"] = rule;
        if (para > 0) one["para"] = para;
        if (number != null && number.Length > 0) one["number"] = number;
        one["sample"] = sample;
        one["suggest"] = suggest;
        issues.Add(one);
    }
}
