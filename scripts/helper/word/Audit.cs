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
            if (m.Success) c.Add("unit", m.Value, "改用上标写法（m³/s、km²、万m³），或写成公式");

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
        try
        {
            int tables = (int)doc.Tables.Count;
            for (int t = 1; t <= tables; t++)
            {
                dynamic tb = doc.Tables.Item(t);
                int at = (int)doc.Range(0, (int)tb.Range.Start).Paragraphs.Count;
                bool ok = false;
                for (int k = Math.Max(1, at - 3); k <= at && k <= total; k++)
                {
                    string above = CleanPara(Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(k).Range.Text, ""), CultureInfo.InvariantCulture));
                    if (above.Contains("\u8868 ") || above.Contains("\u8868\u3000") || Regex.IsMatch(above, "\u8868\\s*[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]")) { ok = true; break; }
                }
                c.Para = at;
                if (!ok) c.Add("caption-missing", "表格 " + t + "（第 " + at + " 段之前）", "表格前缺少「表 N」题注");
            }

            int shapes = (int)doc.InlineShapes.Count;
            for (int s = 1; s <= shapes; s++)
            {
                dynamic sh = doc.InlineShapes.Item(s);
                int at = (int)doc.Range(0, (int)sh.Range.Start).Paragraphs.Count;
                bool ok = false;
                for (int k = at; k <= Math.Min(total, at + 3); k++)
                {
                    string below = CleanPara(Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(k).Range.Text, ""), CultureInfo.InvariantCulture));
                    if (below.Contains("\u56fe ") || below.Contains("\u56fe\u3000") || Regex.IsMatch(below, "\u56fe\\s*[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]")) { ok = true; break; }
                }
                c.Para = at;
                if (!ok) c.Add("caption-missing", "图片 " + s + "（第 " + at + " 段）", "图片下方缺少「图 N」题注");
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
        r["note"] = "These are the mechanical checks only. Read the flagged paragraphs, decide, then write comments with `comment -Items` on this same open document; do not hand the file to an offline skill while WPS has it open.";
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
