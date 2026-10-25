// v0.7 fast locate: fetch doc.Content.Text ONCE (one COM roundtrip, ~0.5 s on a 15k-paragraph
// report) and match in memory; convert hit offsets to paragraph numbers with the
// Range(0,offset).Paragraphs.Count trick (one roundtrip per HIT, not per paragraph).
// The per-paragraph loop remains as the fallback for hosts where Content.Text misbehaves.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

internal static partial class Program
{
    /// <summary>The whole document text in one COM call, or null when the host will not give it.</summary>
    private static string WholeText(dynamic doc)
    {
        try
        {
            string all = Convert.ToString(doc.Content.Text, CultureInfo.InvariantCulture);
            return all ?? null;
        }
        catch (Exception) { return null; }
    }

    /// <summary>
    /// 1-based paragraph candidates for a character offset, via ONE roundtrip:
    /// Range(0,offset).Paragraphs.Count is k when offset sits mid-paragraph k, k-1 when it sits
    /// exactly at paragraph k's start, and 1 for a collapsed range at 0 - so BOTH k and k+1 are
    /// possible answers. Callers disambiguate by checking the paragraph text (cheap, per hit).
    /// Returns null when the host refuses.
    /// </summary>
    private static int[] ParaCandidatesAt(dynamic doc, int offset)
    {
        if (offset < 0) offset = 0;
        try
        {
            int c = (int)doc.Range(0, offset).Paragraphs.Count;
            if (c < 1) return new int[] { 1 };
            return new int[] { c, c + 1 };
        }
        catch (Exception) { return null; }
    }

    /// <summary>Verify a mapped paragraph really contains the needle; false = mapping drifted.</summary>
    private static bool ParaContains(dynamic doc, int para, string needle, StringComparison cmp)
    {
        try
        {
            string t = CleanPara(Convert.ToString(doc.Paragraphs.Item(para).Range.Text, CultureInfo.InvariantCulture));
            return t.IndexOf(needle, cmp) >= 0;
        }
        catch (Exception) { return false; }
    }

    /// <summary>Pick the candidate whose text contains the needle; 0 = neither (drift).</summary>
    private static int ParaOfHit(dynamic doc, int offset, string needle, StringComparison cmp)
    {
        int[] cand = ParaCandidatesAt(doc, offset);
        if (cand == null) return 0;
        foreach (int c in cand)
        {
            if (ParaContains(doc, c, needle, cmp)) return c;
        }
        return 0;
    }

    /// <summary>
    /// Paragraph containing a character offset, chosen from the two candidates by checking
    /// Range.Start/End (prefer the paragraph that STARTS at the offset - matches how a selection
    /// anchored at a boundary reads). 0 = host refused.
    /// </summary>
    private static int ParaAt(dynamic doc, int offset, bool preferStart)
    {
        int[] cand = ParaCandidatesAt(doc, offset);
        if (cand == null) return 0;
        int firstOk = 0;
        foreach (int c in cand)
        {
            try
            {
                dynamic p = doc.Paragraphs.Item(c);
                int s = (int)p.Range.Start, e = (int)p.Range.End;
                if (offset >= s && offset <= e)
                {
                    if (preferStart && s == offset) return c;
                    if (firstOk == 0) firstOk = c;
                }
            }
            catch (Exception) { }
        }
        return firstOk;
    }

    // ------------------------------------------------------------ search (fast)

    /// <summary>
    /// search - the scan runs INSIDE WPS via native Find (no per-paragraph marshalling), so it
    /// works with tracked revisions too: Find returns ranges whose coordinates properly include
    /// revision marks, unlike offsets computed over Content.Text (which hides deleted text and
    /// therefore drifts whenever deletions exist - see ParaCandidatesAt users elsewhere).
    /// -Text may carry several needles separated by '|'. Falls back to the caller's slow loop
    /// only if the host refuses Find outright.
    /// </summary>
    private static object SearchFast(dynamic doc, string text, string styleFilter, bool caseSensitive, int maxHits)
    {
        List<string> needles = new List<string>();
        foreach (string raw in text.Split('|'))
        {
            string n = raw.Trim();
            if (n.Length > 0) needles.Add(n);
        }
        if (needles.Count == 0) throw new Fail("BAD_ARGS", "\"text\" is empty");
        StringComparison cmp = StringComparison.OrdinalIgnoreCase;

        List<object> hits = new List<object>();
        Dictionary<string, int> totals = new Dictionary<string, int>();
        int totalHits = 0;
        const int maxPerPara = 3;
        Dictionary<int, int> perParaCount = new Dictionary<int, int>();
        // context comes from the paragraph text (accept view), same as the slow path reports
        Dictionary<int, string> paraTextCache = new Dictionary<int, string>();

        foreach (string needle in needles)
        {
            dynamic cursor = null;
            try { cursor = doc.Content; } catch (Exception) { return null; }
            while (true)
            {
                bool ok = false;
                try { ok = (bool)cursor.Find.Execute(needle, caseSensitive, false, false, false, false, true, 0, false, "", 0); }
                catch (Exception) { return null; }
                if (!ok) break;

                int foundStart = 0, foundEnd = 0;
                try { foundStart = (int)cursor.Start; foundEnd = (int)cursor.End; }
                catch (Exception) { return null; }

                int para = ParaOfHit(doc, foundStart, needle, cmp);
                if (para == 0) return null;                     // drift/refusal: fall back wholesale

                totalHits++;
                int seen;
                perParaCount.TryGetValue(para, out seen);
                perParaCount[para] = seen + 1;

                if (seen < maxPerPara && hits.Count < maxHits)
                {
                    string ptext;
                    if (!paraTextCache.TryGetValue(para, out ptext))
                    {
                        ptext = "";
                        try { ptext = CleanPara(Convert.ToString(doc.Paragraphs.Item(para).Range.Text, CultureInfo.InvariantCulture)); }
                        catch (Exception) { }
                        paraTextCache[para] = ptext;
                    }
                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["para"] = para;
                    if (needles.Count > 1) one["term"] = needle;
                    int rel = ptext.IndexOf(needle, cmp);
                    one["at"] = rel < 0 ? 0 : rel;
                    one["context"] = rel < 0 ? Head(ptext, 60) : ContextOf(ptext, rel, needle.Length);
                    hits.Add(one);
                }
                int t;
                totals.TryGetValue(needle, out t);
                totals[needle] = t + 1;

                // continue AFTER this hit
                try { cursor = doc.Range(foundEnd, (int)doc.Content.End); }
                catch (Exception) { return null; }
                if (foundEnd >= (int)doc.Content.End) break;
            }
        }

        // Style filter / inTable need per-hit roundtrips; do them only for reported hits.
        if (styleFilter != null && styleFilter.Trim().Length > 0)
        {
            List<object> kept = new List<object>();
            foreach (Dictionary<string, object> h in hits)
            {
                string style = Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item((int)h["para"]).Style.NameLocal, ""), CultureInfo.InvariantCulture);
                if (string.Compare(style, styleFilter.Trim(), StringComparison.OrdinalIgnoreCase) == 0) kept.Add(h);
            }
            hits = kept;
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["text"] = text;
        if (styleFilter != null && styleFilter.Trim().Length > 0) r["style"] = styleFilter;
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["hits"] = hits;
        r["totalHits"] = totalHits;
        if (needles.Count > 1)
        {
            List<object> perTerm = new List<object>();
            foreach (string needle in needles)
            {
                int t;
                totals.TryGetValue(needle, out t);
                Dictionary<string, object> pt = new Dictionary<string, object>();
                pt["term"] = needle; pt["total"] = t;
                perTerm.Add(pt);
            }
            r["terms"] = perTerm;
        }
        if (totalHits > hits.Count) r["truncated"] = true;
        r["saved"] = false;
        r["note"] = "para 就是要喂给 revise(-Para)/comment(-Para)/view(-Para) 的段落号；命中文本是「接受视图」（与 read/diff 同一口径）。-Text 可用 | 分隔多个词一次查完。";
        return r;
    }
}
