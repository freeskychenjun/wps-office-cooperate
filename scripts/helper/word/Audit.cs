// Deterministic checks for an in-place review: the mechanical half of "校审", so the model spends its
// judgement on meaning instead of scanning 600 paragraphs for m3/s and repeated words.
//
//   audit  {"doc":"...","max":200,"rules":"unit,repeat"}
//
// Every finding points at a paragraph number; the model then decides and writes comments with
// `comment -Items` on the same open document - no second skill, no file hand-off.
//
// C# 5: no local functions, no string interpolation (compiled by the in-box csc of .NET Framework 4).

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

internal static partial class Program
{
    private sealed class Finding
    {
        public int Para;
        public string Rule;
        public string Sample;
        public string Suggest;
    }

    /// <summary>Collects findings during the scan; a class because C# 5 has no local functions.</summary>
    private sealed class AuditCollector
    {
        public List<Finding> Found = new List<Finding>();
        public HashSet<string> Wanted;
        public int Max;
        public int Para;
        public bool Truncated;

        public void Add(string rule, string sample, string suggest)
        {
            if (Wanted != null && !Wanted.Contains(rule)) return;
            if (Found.Count >= Max) { Truncated = true; return; }
            Finding f = new Finding();
            f.Para = Para;
            f.Rule = rule;
            f.Sample = Head(sample, 80);
            f.Suggest = suggest;
            Found.Add(f);
        }
    }

    private static object AuditDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        int total = (int)doc.Paragraphs.Count;
        int max = a.Int("max", 200);
        if (max <= 0) max = 200;

        string only = a.Str("rules", null);
        HashSet<string> wanted = null;
        if (only != null && only.Trim().Length > 0)
        {
            wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string one in only.Split(','))
                if (one.Trim().Length > 0) wanted.Add(one.Trim());
        }

        AuditCollector c = new AuditCollector();
        c.Wanted = wanted;
        c.Max = max;

        Regex unit = new Regex(@"[0-9][\s\u3000]?(m3/s|m3|km2|m2)", RegexOptions.Compiled);
        Regex repeat = new Regex(@"([\u4e00-\u9fa5]{2,4})\1", RegexOptions.Compiled);
        Regex doubleSpace = new Regex(@"  ", RegexOptions.Compiled);
        Regex spacePunct = new Regex(@" (，|。|；|：|、|）)", RegexOptions.Compiled);
        Regex fullwidthDigit = new Regex(@"[\uff10-\uff19]", RegexOptions.Compiled);

        int lastHeading = 0;
        int emptyRun = 0;

        for (int i = 1; i <= total; i++)
        {
            if (c.Found.Count >= max) { c.Truncated = true; break; }
            dynamic p = doc.Paragraphs.Item(i);
            string text = CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));
            string style = Convert.ToString(TryGet(() => (object)p.Style.NameLocal, ""), CultureInfo.InvariantCulture);
            c.Para = i;

            if (text.Trim().Length == 0)
            {
                emptyRun++;
                if (emptyRun == 4)
                    c.Add("empty-run", "第 " + (i - 3) + " 段起连续空段", "删除多余空段（排版时也常导致半空页）");
                continue;
            }
            emptyRun = 0;

            Match m;
            m = unit.Match(text);
            if (m.Success) c.Add("unit", m.Value, "单位指数掉下来了（应为上标）。修法：word unit -Apply（按最长单位优先、逐处守卫地批量改成 m³/s、km² 这类 Unicode 上标，已正确的会自动跳过）");

            m = repeat.Match(text);
            if (m.Success) c.Add("repeat-word", m.Value, "疑似重复词，确认后删去一个");

            if (text.IndexOf('\uFF08') >= 0 && text.IndexOf('(') >= 0)
                c.Add("bracket-mix", text, "同一段混用全角括号与半角括号，统一为全角");

            m = doubleSpace.Match(text);
            if (m.Success) c.Add("double-space", text, "连续空格，建议删至一个");

            m = spacePunct.Match(text);
            if (m.Success) c.Add("space-before-punct", m.Value, "全角标点前不应有空格");

            m = fullwidthDigit.Match(text);
            if (m.Success) c.Add("fullwidth-digit", m.Value, "全角数字，正文数字应用半角");

            if (text.EndsWith(" ") || text.EndsWith("\u3000"))
                c.Add("trailing-space", text, "段末多余空格");

            int level = HeadingLevel(style);
            if (level > 0)
            {
                if (lastHeading > 0 && level > lastHeading + 1)
                    c.Add("heading-jump", text, "从标题 " + lastHeading + " 跳到标题 " + level + "，检查是否漏了中间层");
                lastHeading = level;
            }
        }

        // Captions: a table wants "表" above it, a picture wants "图" below it.
        // The judge is the caption PARSER (label must open the paragraph), not a loose "the paragraph
        // contains 表N" regex: a body sentence like "如表3.3-2所示" used to count as a caption (field
        // report 2026-10-26), and a caption typed into the table's first row used to be reported twice
        // (caption-missing AND table-caption-in-row, for the same table).
        try
        {
            int tables = (int)doc.Tables.Count;
            for (int t = 1; t <= tables; t++)
            {
                dynamic tb = doc.Tables.Item(t);
                int startPara = TableStartPara(doc, tb, total);
                bool captionInRow = false;
                try
                {
                    string firstCell = CleanPara(Convert.ToString(TryGet(() => (object)tb.Cell(1, 1).Range.Text, ""), CultureInfo.InvariantCulture)).Trim();
                    captionInRow = IsCaptionText(firstCell);
                    if (captionInRow)
                        c.Add("table-caption-in-row", "表格 " + t + " 第 1 行：" + Head(firstCell, 40),
                              "题名做进了表格第 1 行；用 table -Action caption -Table " + t + " 提到框外（会自动删该行并可顺带加粗边框）");
                }
                catch (Exception) { }

                if (captionInRow) continue;   // the caption exists, it only sits in the wrong place

                bool ok = false;
                for (int k = startPara - 1; k >= 1 && k >= startPara - 3; k--)
                {
                    string above = CleanPara(Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(k).Range.Text, ""), CultureInfo.InvariantCulture));
                    if (above.Trim().Length == 0) continue;      // blank line between caption and table
                    ok = ParseCaption(above) != null;            // first real paragraph decides
                    break;
                }
                c.Para = startPara;
                if (!ok) c.Add("caption-missing", "表格 " + t + "（起始段 " + startPara + "）", "表格前缺少「表 N」题注，表内第 1 行也没有；补题注后可用 word captions 一次核对全篇编号");
            }

            int shapes = (int)doc.InlineShapes.Count;
            for (int s = 1; s <= shapes; s++)
            {
                dynamic sh = doc.InlineShapes.Item(s);
                int at = ParaAtOffset(doc, (int)sh.Range.Start, total);
                bool ok = false;
                // A figure caption may sit 1-3 paragraphs below its picture (measured on a real report:
                // picture / empty paragraph / a stray "/" paragraph / 图2.1-1 xxx). Walk forward a few
                // paragraphs but stop at the next heading, which means the figure's area has ended.
                for (int k = at; k <= Math.Min(total, at + 4); k++)
                {
                    string below = ParaTextAt(doc, k);
                    string style = Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(k).Style.NameLocal, ""), CultureInfo.InvariantCulture);
                    if (below.Trim().Length == 0) continue;
                    CaptionInfo ci = ParseCaption(below);
                    if (ci != null && (ci.Label == "\u56fe" || ci.Label == "\u9644\u56fe")) { ok = true; break; }
                    if (HeadingLevel(style) > 0) break;
                }
                c.Para = at;
                if (!ok) c.Add("caption-missing", "图片 " + s + "（第 " + at + " 段之后）", "图片下方缺少「图 N」题注");
            }
        }
        catch (Exception) { }

        Dictionary<string, object> byRule = new Dictionary<string, object>();
        foreach (Finding f in c.Found)
        {
            int n = byRule.ContainsKey(f.Rule) ? ToInt(byRule, f.Rule) : 0;
            byRule[f.Rule] = n + 1;
        }

        List<object> items = new List<object>();
        foreach (Finding f in c.Found)
        {
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["para"] = f.Para;
            one["rule"] = f.Rule;
            one["sample"] = f.Sample;
            one["suggest"] = f.Suggest;
            items.Add(one);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["paragraphs"] = total;
        r["findings"] = items;
        r["count"] = items.Count;
        r["byRule"] = byRule;
        r["truncated"] = c.Truncated;
        r["saved"] = false;
        r["note"] = "These are the mechanical checks only. Read the flagged paragraphs, decide, then write comments with `comment -Items` on this same open document; do not hand the file to an offline skill while WPS has it open. \u9898\u6ce8\u7f16\u53f7\u672c\u8eab\u7684\u95ee\u9898\uff08\u91cd\u53f7/\u8df3\u53f7/\u683c\u5f0f\u6df7\u7528/\u60ac\u7a7a\u5f15\u7528\uff09\u4e0d\u5728\u8fd9\u91cc\u67e5\uff1a\u8dd1 word captions \u4e00\u6761\u547d\u4ee4\u5c31\u51fa\u6e05\u5355\u3002";
        return r;
    }

    /// <summary>1..4 for 标题 1..4 / Heading 1..4, 1 for 标题, 0 for anything else.</summary>
    private static int HeadingLevel(string style)
    {
        if (style == null) return 0;
        if (style == "\u6807\u9898" || style == "Title") return 1;
        bool cn = style.StartsWith("\u6807\u9898 ");
        bool en = style.StartsWith("Heading ");
        if (!cn && !en) return 0;
        string num = style.Substring(style.LastIndexOf(' ') + 1);
        int n;
        return int.TryParse(num, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? n : 0;
    }
}
