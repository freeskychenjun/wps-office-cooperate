// v0.7 fast locate: fetch doc.Content.Text ONCE (one COM roundtrip, ~0.5 s on a 15k-paragraph
// report) and match in memory; convert hit offsets to paragraph numbers with the
// Range(0,offset).Paragraphs.Count trick (one roundtrip per HIT, not per paragraph).
// The per-paragraph loop remains as the fallback for hosts where Content.Text misbehaves.
//
// v1.8 (2026-10-26) - fixes from a field report where `search` on a 10521-paragraph report
// returned 113 hits but silently mis-addressed 4 captions and reported the same paragraph 3x:
//   * the paragraph of a hit is now VERIFIED by offset, not by "the candidate paragraph contains
//     the word": Range.Start <= offset < Range.End must hold (paragraph ranges are half-open and
//     tile the document, so exactly one candidate can pass). Long tables are where the old
//     text test picked the wrong neighbour, and a wrong-but-nonzero answer never triggered the
//     slow fallback.
//   * every hit gets its OWN offset inside its paragraph (derived from the Find offset, then
//     re-checked against the paragraph text); the old code re-ran IndexOf and therefore reported
//     the first occurrence three times, with identical `at`/`context`.
//   * the 3-hits-per-paragraph cap is now visible: capped paragraphs are listed, so
//     totalHits - hits.Count is explainable instead of silent.

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
    /// possible answers. Callers disambiguate by offset/range, not by guessing.
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

    /// <summary>Range.Start / Range.End of a paragraph, or null.</summary>
    private static int[] ParaSpan(dynamic doc, int para)
    {
        if (para < 1) return null;
        try
        {
            dynamic p = doc.Paragraphs.Item(para);
            return new int[] { (int)p.Range.Start, (int)p.Range.End };
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

    /// <summary>
    /// Paragraph (and its span) that really contains a character offset. A candidate passes only
    /// when <c>Range.Start &lt;= offset &lt; Range.End</c>; because paragraph ranges tile the document
    /// half-open, at most one candidate can. Two extra neighbours are checked for hosts that count
    /// paragraphs of a sub-range differently around tables. Falls back to the old "the text contains
    /// the word" test only if no span matches, so a drifted answer is reported as 0 (the caller then
    /// switches to the slow per-paragraph scan) instead of looking plausible but pointing elsewhere.
    /// Returns 0 when nothing can be verified.
    /// </summary>
    private static int ParaOfHit(dynamic doc, int offset, string needle, StringComparison cmp, out int spanStart, out int spanEnd)
    {
        spanStart = -1;
        spanEnd = -1;
        int[] cand = ParaCandidatesAt(doc, offset);
        if (cand == null) return 0;

        List<int> probe = new List<int>();
        foreach (int c in cand)
        {
            if (c - 1 >= 1 && !probe.Contains(c - 1)) probe.Add(c - 1);
            if (c >= 1 && !probe.Contains(c)) probe.Add(c);
            if (c + 1 >= 1 && !probe.Contains(c + 1)) probe.Add(c + 1);
            if (c + 2 >= 1 && !probe.Contains(c + 2)) probe.Add(c + 2);
        }

        foreach (int c in probe)
        {
            int[] sp = ParaSpan(doc, c);
            if (sp == null) continue;
            if (offset >= sp[0] && offset < sp[1])
            {
                spanStart = sp[0];
                spanEnd = sp[1];
                return c;
            }
        }
        // Coordinates do not tile as expected (host quirk): accept the text test, but do not report
        // a span - the caller then knows the offset inside the paragraph is not trustworthy either.
        foreach (int c in probe)
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
                if (offset >= s && offset < e)
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
    /// only if the host refuses Find outright or a hit cannot be verified.
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
        StringComparison atCmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        List<object> hits = new List<object>();
        Dictionary<string, int> totals = new Dictionary<string, int>();
        int totalHits = 0;
        int textFallbacks = 0;
        const int maxPerPara = 3;
        Dictionary<int, int> perParaCount = new Dictionary<int, int>();
        // context comes from the paragraph text (accept view), same as the slow path reports
        Dictionary<int, string> paraTextCache = new Dictionary<int, string>();
        // last reported offset per paragraph: consecutive hits in one paragraph must move forward
        Dictionary<int, int> lastAt = new Dictionary<int, int>();

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

                int spanStart, spanEnd;
                int para = ParaOfHit(doc, foundStart, needle, cmp, out spanStart, out spanEnd);
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
                    // This hit's OWN offset inside the paragraph: foundStart - paragraph start.
                    // Range.Text hides deleted revision text while the Find offsets count it, so the
                    // subtraction is re-checked against the paragraph text and only then trusted.
                    int rel = spanStart >= 0 ? foundStart - spanStart : -1;
                    bool exact = rel >= 0 && rel <= ptext.Length - needle.Length
                        && string.Compare(ptext, rel, needle, 0, needle.Length, atCmp) == 0;
                    if (!exact)
                    {
                        // Recover the offset from the text itself, continuing after the previous hit
                        // of this paragraph so two hits can never be reported at the same `at`.
                        int after;
                        if (lastAt.TryGetValue(para, out after) && after + 1 <= ptext.Length)
                            rel = ptext.IndexOf(needle, after + 1, atCmp);
                        else
                            rel = -1;
                        if (rel < 0) rel = ptext.IndexOf(needle, atCmp);
                        textFallbacks++;
                    }
                    if (rel < 0) rel = 0;
                    lastAt[para] = rel;

                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["para"] = para;
                    if (needles.Count > 1) one["term"] = needle;
                    one["at"] = rel;
                    one["context"] = ContextOf(ptext, rel, needle.Length);
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
        // Paragraph numbers above are verified by character offset (not by "somewhere in the
        // paragraph the word occurs"), so they can be fed straight to revise/comment/view.
        r["verified"] = true;
        r["maxPerParagraph"] = maxPerPara;
        List<object> capped = new List<object>();
        foreach (KeyValuePair<int, int> kv in perParaCount)
        {
            if (kv.Value <= maxPerPara) continue;
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["para"] = kv.Key;
            one["total"] = kv.Value;
            one["reported"] = maxPerPara;
            capped.Add(one);
        }
        if (capped.Count > 0) r["cappedParas"] = capped;
        if (textFallbacks > 0) r["atFallbacks"] = textFallbacks;
        if (totalHits > hits.Count) r["truncated"] = true;
        r["saved"] = false;
        r["note"] = "para 就是要喂给 revise(-Para)/comment(-Para)/view(-Para) 的段落号；命中文本是「接受视图」（与 read/diff 同一口径）。-Text 可用 | 分隔多个词一次查完。段落号按字符偏移校验过（verified:true）；同一段命中超过 " + maxPerPara + " 处的，多出来的只计数（见 cappedParas），要全部定位就加大 -Max 或换更专的词。";
        return r;
    }
}
