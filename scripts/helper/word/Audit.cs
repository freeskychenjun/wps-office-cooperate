// Deterministic checks for an in-place review: the mechanical half of "校审", so the model spends its
// judgement on meaning instead of scanning 600 paragraphs for m3/s and repeated words.
//
//   audit  {"doc":"...","max":200,"rules":"unit,repeat"}   ("slow":true forces the old loop)
//
// Every finding points at a paragraph number; the model then decides and writes comments with
// `comment -Items` on the same open document - no second skill, no file hand-off.
//
// v1.9 fast path (2026-10-26): the old loop cost 3 COM roundtrips per paragraph
// (Paragraphs.Item + Range.Text + Style.NameLocal) - 31.5k roundtrips on a 10521-paragraph
// report, minutes of wall clock, while the text rules need nothing but the text. The fast path
// fetches Content.Text ONCE and runs the text rules in memory; paragraph numbers are then
// resolved the way `search` does it - native Find for the sample text, span-verified - because
// Content.Text offsets do NOT map onto Range coordinates once tables are present (row-end
// marks; the v0.7 lesson, relearned on the 20-table fixture: offset-verified resolution failed
// on exactly the table-heavy documents). Headings keep real Range offsets straight from
// Find-by-style. The fast path runs only when Revisions.Count == 0 (Content.Text hides deleted
// revision text) and falls back to the old loop wholesale on any refusal - correctness first.
//
// C# 5: no local functions, no string interpolation (compiled by the in-box csc of .NET Framework 4).

using System;
using System.Collections.Generic;
using System.Diagnostics;
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

    // Rules shared by the fast and slow paths; ONE definition so the two can never disagree.
    private static readonly Regex AuditUnitRx = new Regex(@"[0-9][\s\u3000]?(m3/s|m3|km2|m2)", RegexOptions.Compiled);
    private static readonly Regex AuditRepeatRx = new Regex(@"([\u4e00-\u9fa5]{2,4})\1", RegexOptions.Compiled);
    private static readonly Regex AuditDoubleSpaceRx = new Regex(@"  ", RegexOptions.Compiled);
    private static readonly Regex AuditSpacePunctRx = new Regex(@" (，|。|；|：|、|）)", RegexOptions.Compiled);
    private static readonly Regex AuditFullwidthDigitRx = new Regex(@"[\uff10-\uff19]", RegexOptions.Compiled);

    private const string AuditUnitSuggest =
        "\u5355\u4f4d\u6307\u6570\u6389\u4e0b\u6765\u4e86\uff08\u5e94\u4e3a\u4e0a\u6807\uff09\u3002\u4fee\u6cd5\uff1aword unit -Apply\uff08\u6309\u6700\u957f\u5355\u4f4d\u4f18\u5148\u3001\u9010\u5904\u5b88\u536b\u5730\u6279\u91cf\u6539\u6210 m\u00b3/s\u3001km\u00b2 \u8fd9\u7c7b Unicode \u4e0a\u6807\uff0c\u5df2\u6b63\u786e\u7684\u4f1a\u81ea\u52a8\u8df3\u8fc7\uff09";
    private const string AuditRepeatSuggest = "\u7591\u4f3c\u91cd\u590d\u8bcd\uff0c\u786e\u8ba4\u540e\u5220\u53bb\u4e00\u4e2a";
    private const string AuditBracketSuggest = "\u540c\u4e00\u6bb5\u6df7\u7528\u5168\u89d2\u62ec\u53f7\u4e0e\u534a\u89d2\u62ec\u53f7\uff0c\u7edf\u4e00\u4e3a\u5168\u89d2";
    private const string AuditDoubleSpaceSuggest = "\u8fde\u7eed\u7a7a\u683c\uff0c\u5efa\u8bae\u5220\u81f3\u4e00\u4e2a";
    private const string AuditSpacePunctSuggest = "\u5168\u89d2\u6807\u70b9\u524d\u4e0d\u5e94\u6709\u7a7a\u683c";
    private const string AuditFullwidthSuggest = "\u5168\u89d2\u6570\u5b57\uff0c\u6b63\u6587\u6570\u5b57\u5e94\u7528\u534a\u89d2";
    private const string AuditTrailingSuggest = "\u6bb5\u672b\u591a\u4f59\u7a7a\u683c";
    private const string AuditEmptyRunSuggest = "\u5220\u9664\u591a\u4f59\u7a7a\u6bb5\uff08\u6392\u7248\u65f6\u4e5f\u5e38\u5bfc\u81f4\u534a\u7a7a\u9875\uff09";

    /// <summary>A finding located in Content.Text before its paragraph number is resolved.</summary>
    private sealed class RawFind
    {
        public int SortOff;      // block start offset - keeps the paragraph/rule order of the old loop
        public int ResolveOff;   // heading-jump only: a REAL Range offset from Find-by-style
        public int Seq;          // rule order inside one paragraph (unit first ... heading-jump last)
        public string Rule;
        public string Sample;    // text findings: what native Find locates the paragraph by
        public string Suggest;
        public int Para;         // filled at resolution; the emission sort key
        public bool EmptyRun;    // resolved through an anchor block instead
        public string NextText;  // empty-run only: the anchor block's text
        public bool NextAfter;   // anchor sits after the run (normal) or before it (run at doc end)
    }

    /// <summary>Why the fast path bailed out (diagnostics in the response when it falls back).</summary>
    private static string _auditFastWhy;

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

        Stopwatch sw = Stopwatch.StartNew();
        string path = "slow";
        int revs = (int)TryGet(() => (object)doc.Revisions.Count, 1);

        AuditCollector c = NewCollector(wanted, max);
        if (!a.Flag("slow", false) && revs == 0)
        {
            _auditFastWhy = null;
            if (AuditFastScan(doc, total, max, wanted, c)) path = "fast";
            else c = NewCollector(wanted, max);   // refuse -> rerun the whole thing the old way
        }
        if (path == "slow") AuditSlowScan(doc, total, max, c);
        AuditTableCaptions(doc, total, c);

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
        r["path"] = path;
        r["ms"] = sw.ElapsedMilliseconds;
        if (path == "slow" && _auditFastWhy != null) r["fastFail"] = _auditFastWhy;
        r["note"] = "These are the mechanical checks only. Read the flagged paragraphs, decide, then write comments with `comment -Items` on this same open document; do not hand the file to an offline skill while WPS has it open. \u9898\u6ce8\u7f16\u53f7\u672c\u8eab\u7684\u95ee\u9898\uff08\u91cd\u53f7/\u8df3\u53f7/\u683c\u5f0f\u6df7\u7528/\u60ac\u7a7a\u5f15\u7528\uff09\u4e0d\u5728\u8fd9\u91cc\u67e5\uff1a\u8dd1 word captions \u4e00\u6761\u547d\u4ee4\u5c31\u51fa\u6e05\u5355\u3002";
        return r;
    }

    private static AuditCollector NewCollector(HashSet<string> wanted, int max)
    {
        AuditCollector c = new AuditCollector();
        c.Wanted = wanted;
        c.Max = max;
        return c;
    }

    /// <summary>Whether a rule survives the -Rules filter (null = all rules wanted).</summary>
    private static bool AuditRuleWanted(HashSet<string> wanted, string rule)
    {
        return wanted == null || wanted.Contains(rule);
    }

    // ------------------------------------------------------------ fast path

    /// <summary>
    /// One Content.Text fetch, in-memory text rules, Find-by-style headings; paragraph numbers
    /// resolved by native Find of the sample text (monotonic, document order) - the same verified
    /// mechanism `search` uses, immune to the row-end offset drift of Content.Text.
    /// Returns false on ANY refusal; the caller reruns everything through the old loop.
    /// </summary>
    private static bool AuditFastScan(dynamic doc, int total, int max, HashSet<string> wanted, AuditCollector c)
    {
        string all = WholeText(doc);
        if (all == null) { _auditFastWhy = "wholeText"; return false; }

        List<RawFind> raws = new List<RawFind>();
        int off = 0;
        int emptyRun = 0;
        RawFind pendingEmpty = null;   // the newest empty-run raw still waiting for its anchor
        foreach (string block in all.Split('\r'))
        {
            if (raws.Count >= max) { c.Truncated = true; break; }
            string text = CleanPara(block);
            if (text.Trim().Length == 0)
            {
                emptyRun++;
                if (emptyRun == 4 && AuditRuleWanted(wanted, "empty-run"))
                {
                    RawFind f = new RawFind();
                    f.SortOff = off; f.Seq = 5;
                    f.Rule = "empty-run"; f.EmptyRun = true; f.Suggest = AuditEmptyRunSuggest;
                    raws.Add(f);
                    pendingEmpty = f;
                }
            }
            else
            {
                emptyRun = 0;
                if (pendingEmpty != null) { pendingEmpty.NextText = text; pendingEmpty.NextAfter = true; pendingEmpty = null; }
                Match m;
                if (AuditRuleWanted(wanted, "unit"))
                {
                    m = AuditUnitRx.Match(text);
                    if (m.Success) AddRaw(raws, off, 10, "unit", m.Value, AuditUnitSuggest);
                }
                if (AuditRuleWanted(wanted, "repeat-word"))
                {
                    m = AuditRepeatRx.Match(text);
                    if (m.Success) AddRaw(raws, off, 20, "repeat-word", m.Value, AuditRepeatSuggest);
                }
                if (AuditRuleWanted(wanted, "bracket-mix")
                    && text.IndexOf('\uFF08') >= 0 && text.IndexOf('(') >= 0)
                    AddRaw(raws, off, 30, "bracket-mix", text, AuditBracketSuggest);
                if (AuditRuleWanted(wanted, "double-space"))
                {
                    m = AuditDoubleSpaceRx.Match(text);
                    if (m.Success) AddRaw(raws, off, 40, "double-space", text, AuditDoubleSpaceSuggest);
                }
                if (AuditRuleWanted(wanted, "space-before-punct"))
                {
                    m = AuditSpacePunctRx.Match(text);
                    if (m.Success) AddRaw(raws, off, 50, "space-before-punct", m.Value, AuditSpacePunctSuggest);
                }
                if (AuditRuleWanted(wanted, "fullwidth-digit"))
                {
                    m = AuditFullwidthDigitRx.Match(text);
                    if (m.Success) AddRaw(raws, off, 60, "fullwidth-digit", m.Value, AuditFullwidthSuggest);
                }
                if (AuditRuleWanted(wanted, "trailing-space")
                    && (text.EndsWith(" ") || text.EndsWith("\u3000")))
                    AddRaw(raws, off, 70, "trailing-space", text, AuditTrailingSuggest);
            }
            off += block.Length + 1;
        }

        // A run at the very end of the document has no block after it: anchor on the LAST
        // non-empty block instead (the run then sits in the four paragraphs after it).
        if (pendingEmpty != null)
        {
            string tail = null;
            foreach (string block in all.Split('\r'))
            {
                if (CleanPara(block).Trim().Length > 0) tail = CleanPara(block);
            }
            pendingEmpty.NextText = tail;
            pendingEmpty.NextAfter = false;
        }

        // Heading paragraphs: enumerate the styles once (names only), then Find-by-style for every
        // style HeadingLevel recognises - exactly the old NameLocal test, so the two paths agree.
        if (AuditRuleWanted(wanted, "heading-jump"))
        {
            if (!AuditFindHeadings(doc, all, raws)) return false;
        }

        raws.Sort(delegate(RawFind x, RawFind y)
        {
            if (x.SortOff != y.SortOff) return x.SortOff.CompareTo(y.SortOff);
            return x.Seq.CompareTo(y.Seq);
        });

        int contentEnd = 0;
        try { contentEnd = (int)doc.Content.End; }
        catch (Exception) { _auditFastWhy = "contentEnd"; return false; }

        int searchFrom = 0;
        Dictionary<int, int[]> cache = new Dictionary<int, int[]>();
        foreach (RawFind f in raws)
        {
            int para;
            if (f.Rule == "heading-jump")
            {
                para = AuditParaOfOffset(doc, f.ResolveOff, total, cache);   // real Range offset
                if (para <= 0) { _auditFastWhy = "resolve:heading"; return false; }
            }
            else if (f.EmptyRun)
            {
                if (f.NextText == null || f.NextText.Trim().Length == 0) { _auditFastWhy = "resolve:empty-run"; return false; }
                // a local cursor: the anchor paragraph may carry its own finding later in raws,
                // so the shared searchFrom must not be advanced past it
                int from2 = searchFrom;
                int anchor = AuditFindBySample(doc, ref from2, contentEnd, f.NextText, false);
                if (anchor <= 0) { _auditFastWhy = "resolve:empty-run"; return false; }
                para = f.NextAfter ? anchor - 1 : anchor + 4;
                // verify: the four paragraphs para-3..para must really all be empty
                if (para - 3 < 1 || para > total) { _auditFastWhy = "resolve:empty-run-verify"; return false; }
                for (int k = para - 3; k <= para; k++)
                    if (ParaTextAt(doc, k).Trim().Length > 0) { _auditFastWhy = "resolve:empty-run-verify"; return false; }
            }
            else
            {
                para = AuditFindBySample(doc, ref searchFrom, contentEnd, f.Sample, true);
                if (para <= 0) { _auditFastWhy = "resolve:" + f.Rule; return false; }
            }
            f.Para = para;
        }

        // Emit in PARAGRAPH order with the old loop's rule order inside each paragraph, so the
        // fast and slow paths return byte-identical lists.
        raws.Sort(delegate(RawFind x, RawFind y)
        {
            if (x.Para != y.Para) return x.Para.CompareTo(y.Para);
            if (x.Seq != y.Seq) return x.Seq.CompareTo(y.Seq);
            return 0;
        });
        foreach (RawFind f in raws)
        {
            c.Para = f.Para;
            string sample = f.Sample;
            if (f.EmptyRun) sample = "\u7b2c " + (f.Para - 3) + " \u6bb5\u8d77\u8fde\u7eed\u7a7a\u6bb5";
            c.Add(f.Rule, sample, f.Suggest);
        }
        return true;
    }

    private static void AddRaw(List<RawFind> raws, int sortOff, int seq, string rule, string sample, string suggest)
    {
        RawFind f = new RawFind();
        f.SortOff = sortOff;
        f.Seq = seq;
        f.Rule = rule;
        f.Sample = sample;
        f.Suggest = suggest;
        raws.Add(f);
    }

    /// <summary>
    /// Paragraph number of the next occurrence of sample text at/after searchFrom, via native
    /// Find (the scan runs inside WPS, coordinates are real). searchFrom advances past the hit so
    /// findings are matched in document order. A trailing-space sample that Find cannot honour is
    /// retried trimmed. Returns 0 when the text cannot be found (drift/refusal).
    /// </summary>
    private static int AuditFindBySample(dynamic doc, ref int searchFrom, int contentEnd, string sample, bool retryTrimmed)
    {
        if (sample == null) return 0;
        string needle = sample;
        while (true)
        {
            if (needle == null || needle.Length == 0) return 0;
            dynamic cursor;
            try { cursor = doc.Range(searchFrom, contentEnd); }
            catch (Exception) { return 0; }
            bool ok = false;
            try { ok = (bool)cursor.Find.Execute(Head(needle, 200), true, false, false, false, false, true, 0, false, "", 0); }
            catch (Exception) { return 0; }
            if (!ok)
            {
                if (retryTrimmed && needle.Trim().Length > 0 && needle.Trim() != needle)
                {
                    needle = needle.Trim();
                    continue;
                }
                return 0;
            }
            int start = 0;
            try { start = (int)cursor.Start; searchFrom = (int)cursor.End; }
            catch (Exception) { return 0; }
            int spanStart, spanEnd;
            return ParaOfHit(doc, start, Head(needle, 200), StringComparison.Ordinal, out spanStart, out spanEnd);
        }
    }

    /// <summary>Locate every heading paragraph via native Find-by-style; emit heading-jump raws.</summary>
    private static bool AuditFindHeadings(dynamic doc, string all, List<RawFind> raws)
    {
        // 1) which styles are headings - by NameLocal, the same test the slow path applies
        List<int[]> headingStyles = new List<int[]>();   // {style index, level}
        int nStyles = 0;
        try { nStyles = (int)doc.Styles.Count; }
        catch (Exception) { _auditFastWhy = "styles.count"; return false; }
        for (int i = 1; i <= nStyles; i++)
        {
            string name = Convert.ToString(TryGet(() => (object)doc.Styles.Item(i).NameLocal, ""), CultureInfo.InvariantCulture);
            int lvl = HeadingLevel(name);
            if (lvl > 0) headingStyles.Add(new int[] { i, lvl });
        }
        if (headingStyles.Count == 0) return true;   // no heading styles: nothing can jump

        int contentEnd = 0;
        try { contentEnd = (int)doc.Content.End; }
        catch (Exception) { _auditFastWhy = "contentEnd"; return false; }

        // 2) Find every paragraph of each heading style (the scan runs inside WPS)
        List<object[]> hits = new List<object[]>();     // {offset, level, text-from-the-hit-range}
        foreach (int[] se in headingStyles)
        {
            object styleObj = TryGet(() => (object)doc.Styles.Item(se[0]), null);
            if (styleObj == null) continue;
            dynamic cursor = null;
            try
            {
                cursor = doc.Content;
                cursor.Find.ClearFormatting();
                cursor.Find.Style = styleObj;
            }
            catch (Exception) { _auditFastWhy = "find.style"; return false; }
            int guard = 0;
            while (guard++ < 5000)
            {
                bool ok = false;
                try { ok = (bool)cursor.Find.Execute("", false, false, false, false, false, true, 0, true, "", 0); }
                catch (Exception) { _auditFastWhy = "find.exec"; return false; }
                if (!ok) break;
                int s = 0, e = 0;
                try { s = (int)cursor.Start; e = (int)cursor.End; }
                catch (Exception) { _auditFastWhy = "find.range"; return false; }
                if (e <= s) break;
                // the sample comes from the hit range itself - indexing Content.Text with a Range
                // offset drifts once tables are present (same lesson as the empty sample it caused)
                string htext = CleanPara(Convert.ToString(TryGet(() => (object)cursor.Text, ""), CultureInfo.InvariantCulture));
                hits.Add(new object[] { s, se[1], htext });
                if (e >= contentEnd) break;
                try
                {
                    cursor = doc.Range(e, contentEnd);
                    cursor.Find.ClearFormatting();
                    cursor.Find.Style = styleObj;   // Find state is per-Range: re-arm the new one
                }
                catch (Exception) { _auditFastWhy = "find.rearm"; return false; }
            }
        }

        // 3) heading-jump in document order (the old loop's rule)
        hits.Sort(delegate(object[] x, object[] y) { return ((int)x[0]).CompareTo((int)y[0]); });
        int last = 0;
        foreach (object[] h in hits)
        {
            int level = (int)h[1];
            if (last > 0 && level > last + 1)
            {
                RawFind f = new RawFind();
                f.SortOff = (int)h[0];
                f.ResolveOff = (int)h[0];      // a real Range offset from Find-by-style
                f.Seq = 100;
                f.Rule = "heading-jump";
                f.Sample = (string)h[2];
                f.Suggest = "\u4ece\u6807\u9898 " + last + " \u8df3\u5230\u6807\u9898 " + level + "\uff0c\u68c0\u67e5\u662f\u5426\u6f0f\u4e86\u4e2d\u95f4\u5c42";
                raws.Add(f);
            }
            last = level;
        }
        return true;
    }

    /// <summary>
    /// Paragraph number containing a character offset, verified by span
    /// (Range.Start &lt;= offset &lt; Range.End) with a per-call span cache. Returns -1 when no
    /// candidate verifies - used for heading offsets only (they come from Find, not Content.Text).
    /// </summary>
    private static int AuditParaOfOffset(dynamic doc, int offset, int total, Dictionary<int, int[]> cache)
    {
        if (offset < 0) return -1;
        int[] cand = ParaCandidatesAt(doc, offset);
        if (cand == null) return -1;
        List<int> probe = new List<int>();
        foreach (int cc in cand)
        {
            foreach (int p in new int[] { cc - 1, cc, cc + 1, cc + 2 })
                if (p >= 1 && !probe.Contains(p)) probe.Add(p);
        }
        foreach (int p in probe)
        {
            if (p > total) continue;
            int[] sp;
            if (!cache.TryGetValue(p, out sp)) { sp = ParaSpan(doc, p); cache[p] = sp; }
            if (sp != null && offset >= sp[0] && offset < sp[1]) return p;
        }
        return -1;
    }

    // ------------------------------------------------------------ slow path (the v1.8 loop, unchanged)

    private static void AuditSlowScan(dynamic doc, int total, int max, AuditCollector c)
    {
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
                    c.Add("empty-run", "\u7b2c " + (i - 3) + " \u6bb5\u8d77\u8fde\u7eed\u7a7a\u6bb5", AuditEmptyRunSuggest);
                continue;
            }
            emptyRun = 0;

            Match m;
            m = AuditUnitRx.Match(text);
            if (m.Success) c.Add("unit", m.Value, AuditUnitSuggest);

            m = AuditRepeatRx.Match(text);
            if (m.Success) c.Add("repeat-word", m.Value, AuditRepeatSuggest);

            if (text.IndexOf('\uFF08') >= 0 && text.IndexOf('(') >= 0)
                c.Add("bracket-mix", text, AuditBracketSuggest);

            m = AuditDoubleSpaceRx.Match(text);
            if (m.Success) c.Add("double-space", text, AuditDoubleSpaceSuggest);

            m = AuditSpacePunctRx.Match(text);
            if (m.Success) c.Add("space-before-punct", m.Value, AuditSpacePunctSuggest);

            m = AuditFullwidthDigitRx.Match(text);
            if (m.Success) c.Add("fullwidth-digit", m.Value, AuditFullwidthSuggest);

            if (text.EndsWith(" ") || text.EndsWith("\u3000"))
                c.Add("trailing-space", text, AuditTrailingSuggest);

            int level = HeadingLevel(style);
            if (level > 0)
            {
                if (lastHeading > 0 && level > lastHeading + 1)
                    c.Add("heading-jump", text, "\u4ece\u6807\u9898 " + lastHeading + " \u8df3\u5230\u6807\u9898 " + level + "\uff0c\u68c0\u67e5\u662f\u5426\u6f0f\u4e86\u4e2d\u95f4\u5c42");
                lastHeading = level;
            }
        }
    }

    // ------------------------------------------------------------ captions (shared by both paths)

    /// <summary>
    /// Captions: a table wants "表" above it, a picture wants "图" below it.
    /// The judge is the caption PARSER (label must open the paragraph), not a loose "the paragraph
    /// contains 表N" regex: a body sentence like "如表3.3-2所示" used to count as a caption (field
    /// report 2026-10-26), and a caption typed into the table's first row used to be reported twice.
    /// O(tables + images) roundtrips - cheap on any document.
    /// </summary>
    private static void AuditTableCaptions(dynamic doc, int total, AuditCollector c)
    {
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
                        c.Add("table-caption-in-row", "\u8868\u683c " + t + " \u7b2c 1 \u884c\uff1a" + Head(firstCell, 40),
                              "\u9898\u540d\u505a\u8fdb\u4e86\u8868\u683c\u7b2c 1 \u884c\uff1b\u7528 table -Action caption -Table " + t + " \u63d0\u5230\u6846\u5916\uff08\u4f1a\u81ea\u52a8\u5220\u8be5\u884c\u5e76\u53ef\u987a\u5e26\u52a0\u7c97\u8fb9\u6846\uff09");
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
                if (!ok) c.Add("caption-missing", "\u8868\u683c " + t + "\uff08\u8d77\u59cb\u6bb5 " + startPara + "\uff09", "\u8868\u683c\u524d\u7f3a\u5c11\u300c\u8868 N\u300d\u9898\u6ce8\uff0c\u8868\u5185\u7b2c 1 \u884c\u4e5f\u6ca1\u6709\uff1b\u8865\u9898\u6ce8\u540e\u53ef\u7528 word captions \u4e00\u6b21\u6838\u5bf9\u5168\u7bc7\u7f16\u53f7");
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
                if (!ok) c.Add("caption-missing", "\u56fe\u7247 " + s + "\uff08\u7b2c " + at + " \u6bb5\u4e4b\u540e\uff09", "\u56fe\u7247\u4e0b\u65b9\u7f3a\u5c11\u300c\u56fe N\u300d\u9898\u6ce8");
            }
        }
        catch (Exception) { }
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
