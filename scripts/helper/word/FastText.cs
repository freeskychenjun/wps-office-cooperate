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
using System.Diagnostics;
using System.Xml;
using System.Text;
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

    // ------------------------------------------------------------ shared block model (v1.10)

    /// <summary>
    /// Normalize one '\r'-split block of Content.Text into paragraph text. In WPS every table
    /// cell paragraph contributes "text\r BEL" and therefore the NEXT block starts with a leading
    /// BEL - strip both ends (CleanPara only trims the trailing one). Measured on the 20-table /
    /// 352-row fixture (2026-10-26): block k (1-based) IS paragraph k, tables included.
    /// </summary>
    private static string BlockNorm(string raw)
    {
        if (raw == null) return "";
        int s = 0, e = raw.Length;
        while (s < e && (raw[s] == '\a' || raw[s] == '\u0007')) s++;
        while (e > s && (raw[e - 1] == '\r' || raw[e - 1] == '\n' || raw[e - 1] == '\a' || raw[e - 1] == '\u0007')) e--;
        return raw.Substring(s, e - s);
    }

    /// <summary>
    /// Content.Text split into '\r' blocks, verified to be the WPS shape where block k-1 IS
    /// paragraph k: (a) no unaccepted revisions (deleted text hides from Content.Text but still
    /// occupies Range coordinates - the v0.7 lesson), (b) the block count matches the paragraph
    /// count once trailing cell marks / the trailing '\r' are accounted for, (c) spot-checks of the
    /// first / last / middle and evenly spread interior paragraphs agree with Paragraphs.Item(k).
    /// Any mismatch returns null and the callers fall back to per-paragraph reads - correctness first.
    ///
    /// v1.13: the fallback is no longer silent - <paramref name="why"/> says which check failed and
    /// callers surface it as "fastFail". Field evidence (6 real WPS tasks, 2026-10-25..30) showed the
    /// fast path silently falling back on a 7339-paragraph / 23-table report, putting snapshot / diff
    /// / read back on the ~50 s per-paragraph loop with nothing visible to explain why.
    /// </summary>
    private static string[] FastBlocksVerified(dynamic doc, int total, bool forceSlow)
    {
        string why;
        return FastBlocksVerified(doc, total, forceSlow, out why);
    }

    /// <summary>Why the last FastBlocksVerified call fell back (null when it did not).</summary>
    private static string _fastWhy;

    private static string[] FastBlocksVerified(dynamic doc, int total, bool forceSlow, out string why)
    {
        why = null;
        _fastWhy = null;
        if (forceSlow) return Fallback("forced -Slow", out why);
        if (total <= 0) return Fallback("no paragraphs", out why);
        int revs = (int)TryGet(() => (object)(int)doc.Revisions.Count, 1);
        if (revs != 0) return Fallback("revisions=" + revs, out why);
        string all = WholeText(doc);
        if (all == null) return Fallback("Content.Text refused", out why);
        string[] blocks = all.Split('\r');
        int raw = blocks.Length;

        int end = blocks.Length;
        while (end > 1 && MarkOnly(blocks[end - 1])) end--;          // trailing cell marks
        bool sentinel = end > 0 && blocks[end - 1].Length == 0;      // the usual trailing '\r'
        int count = sentinel ? end - 1 : end;
        if (count != total)
            return Fallback("blocks=" + raw + " usable=" + count + " total=" + total, out why);

        string[] shrunk = new string[total];
        Array.Copy(blocks, shrunk, total);
        foreach (int k in SpotParagraphs(total))
        {
            string para;
            try { para = CleanPara(Convert.ToString(doc.Paragraphs.Item(k).Range.Text, CultureInfo.InvariantCulture)); }
            catch (Exception) { return Fallback("spot p" + k + " refused", out why); }
            string block = BlockNorm(shrunk[k - 1]);
            if (para == null || block != para)
                return Fallback("spot p" + k + " differs (block " + Describe(block) + " vs para " + Describe(para) + ")", out why);
        }
        return shrunk;
    }

    private static string[] Fallback(string reason, out string why)
    {
        why = reason;
        _fastWhy = reason;
        return null;
    }

    /// <summary>True when a raw block holds only WPS cell/paragraph marks, no text at all.</summary>
    private static bool MarkOnly(string raw)
    {
        if (raw == null || raw.Length == 0) return false;
        foreach (char c in raw)
            if (c != '\a' && c != '\u0007' && c != '\r' && c != '\n' && c != ' ' && c != '\u00a0' && c != '\t')
                return false;
        return true;
    }

    /// <summary>
    /// Paragraph numbers whose text is cross-checked against the block model: the ends, the middle
    /// and (v1.13) an evenly spread interior set - a single table in the middle of a report used to
    /// pass the old 3-sample check while every paragraph after it was off by one.
    /// </summary>
    private static List<int> SpotParagraphs(int total)
    {
        List<int> ks = new List<int>();
        int extra = total >= 1000 ? 12 : (total >= 200 ? 8 : 3);
        for (int j = 0; j < extra; j++)
        {
            int k = (int)Math.Round(1.0 + (total - 1) * (double)j / Math.Max(1, extra - 1));
            if (k < 1) k = 1;
            if (k > total) k = total;
            if (!ks.Contains(k)) ks.Add(k);
        }
        int mid = Math.Max(1, total / 2);
        if (!ks.Contains(mid)) ks.Add(mid);
        ks.Sort();
        return ks;
    }


    // ------------------------------------------------------------ chunked range reads (v1.13)

    /// <summary>
    /// Paragraph texts via a few BIG Range.Text reads instead of one read per paragraph.
    ///
    /// Why this exists: Content.Text (the v1.7-v1.10 block model) is unusable on some real WPS
    /// documents - measured on a 7334-paragraph / 23-table report: Content.Text produced 7328 blocks
    /// whose text was misaligned from paragraph 2 onwards (table cells come back in a different
    /// order), so snapshot/diff/read silently fell back to ~5 roundtrips per paragraph (snapshot
    /// 45 s, diff 52 s, read 30 s per call). Content.Text also HIDES deleted revision text, so it
    /// can never be trusted while track changes are on - which is exactly when live editing runs.
    ///
    /// This path only uses primitives the slow path already uses (Paragraphs.Item, Range.Start,
    /// Range.End, Range.Text), so it returns the SAME strings the per-paragraph loop produces.
    ///
    /// Safety: a spanning Range.Text does NOT always split into one piece per paragraph (measured:
    /// a paragraph holding only "/" and a paragraph holding only a page break each vanish from the
    /// span text, so the piece count drops). Therefore every chunk is verified twice - the piece
    /// count must equal the paragraph count, and three sampled paragraphs must match their pieces -
    /// and a failing chunk is SPLIT IN HALF and retried, down to 4 paragraphs. Only the few
    /// paragraphs WPS will not render in a span are then read one by one, so the result is exactly
    /// the slow path's result at a fraction of the roundtrips.
    /// </summary>
    private static List<string> DocTextsByChunks(dynamic doc, int total, out string why)
    {
        why = null;
        if (total < 2) return null;
        List<string> texts = new List<string>(total);
        ChunkStats st = new ChunkStats();
        if (!ExtractRange(doc, 1, total, texts, st))
        {
            why = "chunk reads aborted after " + st.Calls + " calls";
            return null;
        }
        if (texts.Count != total)
        {
            why = "chunk total " + texts.Count + " != " + total;
            return null;
        }
        _chunkStats = st.Describe();
        return texts;
    }

    /// <summary>Counters for the chunked extractor, surfaced as "chunkStats" in the response.</summary>
    private class ChunkStats
    {
        public int Calls;
        public int ChunkReads;
        public int ChunkFails;
        public int ParaReads;

        public string Describe()
        {
            return "calls=" + Calls + " bigReads=" + ChunkReads + " split=" + ChunkFails + " singleReads=" + ParaReads;
        }
    }

    private static string _chunkStats;

    private static bool ExtractRange(dynamic doc, int from, int to, List<string> texts, ChunkStats st)
    {
        if (to < from) return true;
        int count = to - from + 1;
        if (count > 4)
        {
            string body = null;
            try
            {
                int s0 = (int)doc.Paragraphs.Item(from).Range.Start;
                int s1 = (int)doc.Paragraphs.Item(to).Range.End;
                st.Calls += 2;
                body = Convert.ToString(doc.Range(s0, s1).Text, CultureInfo.InvariantCulture);
                st.Calls++;
            }
            catch (Exception) { body = null; }
            if (body != null)
            {
                string[] pieces = body.Split('\r');
                int n = pieces.Length;
                if (n == count + 1 && pieces[n - 1].Length == 0) n = count;
                if (n == count && VerifyChunk(doc, from, to, pieces, st))
                {
                    st.ChunkReads++;
                    for (int i = 0; i < count; i++) texts.Add(BlockNorm(pieces[i]));
                    return true;
                }
                st.ChunkFails++;
            }
            int mid = from + count / 2;
            return ExtractRange(doc, from, mid - 1, texts, st) && ExtractRange(doc, mid, to, texts, st);
        }
        for (int k = from; k <= to; k++)
        {
            string t;
            try { t = CleanPara(Convert.ToString(doc.Paragraphs.Item(k).Range.Text, CultureInfo.InvariantCulture)); st.Calls++; }
            catch (Exception) { return false; }
            texts.Add(t);
        }
        st.ParaReads += count;
        return true;
    }

    /// <summary>
    /// A chunk is only trusted when three sampled paragraphs (first / middle / last) match their
    /// pieces, on top of the caller's piece-count check. Drift from a paragraph that a spanning
    /// Range.Text swallows is caught by the count; this catches the rarer case of compensating
    /// drops that keep the count intact.
    /// </summary>
    private static bool VerifyChunk(dynamic doc, int from, int to, string[] pieces, ChunkStats st)
    {
        int[] probes = new int[] { from, from + (to - from) / 2, to };
        foreach (int k in probes)
        {
            int idx = k - from;
            if (idx < 0 || idx >= pieces.Length) return false;
            string para;
            try { para = CleanPara(Convert.ToString(doc.Paragraphs.Item(k).Range.Text, CultureInfo.InvariantCulture)); st.Calls++; }
            catch (Exception) { return false; }
            if (BlockNorm(pieces[idx]) != para) return false;
        }
        return true;
    }

    // ------------------------------------------------------------ textdiag (v1.13)

    /// <summary>
    /// textdiag - why the fast read paths do or do not apply to THIS document, with the raw evidence.
    /// Read-only, one WordOpenXML fetch. It exists because a silent fallback is indistinguishable
    /// from a slow document: a 7339-paragraph report paid ~50 s per snapshot / diff with nothing
    /// in the response to explain it (see the v1.13 entry in SKILL.md).
    /// </summary>
    private static object TextDiag(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        int total = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        int revs = (int)TryGet(() => (object)(int)doc.Revisions.Count, 1);
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        int sumRows = 0;
        for (int t = 1; t <= tables; t++)
            sumRows += (int)TryGet(() => (object)(int)doc.Tables.Item(t).Rows.Count, 0);

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = DocName(doc);
        r["paragraphs"] = total;
        r["revisions"] = revs;
        r["tables"] = tables;
        r["tableRows"] = sumRows;

        // ---- (1) the v1.7-v1.10 block model: Content.Text split on '\r'
        string all = WholeText(doc);
        if (all == null)
        {
            r["contentChars"] = -1;
            r["blocks"] = -1;
            r["blockModel"] = "Content.Text refused";
        }
        else
        {
            string[] blocks = all.Split('\r');
            r["contentChars"] = all.Length;
            r["blocks"] = blocks.Length;
            r["blockDelta"] = blocks.Length - total;
            List<object> tail = new List<object>();
            for (int i = Math.Max(0, blocks.Length - 2); i < blocks.Length; i++)
                tail.Add(Describe(BlockNorm(blocks[i])));
            r["blockTail"] = tail;
            int blockMiss = 0;
            List<object> blockSpots = new List<object>();
            foreach (int k in SpotParagraphs(total))
            {
                string para = "";
                try { para = CleanPara(Convert.ToString(doc.Paragraphs.Item(k).Range.Text, CultureInfo.InvariantCulture)); }
                catch (Exception) { }
                string block = k - 1 < blocks.Length ? BlockNorm(blocks[k - 1]) : "";
                if (block != para) blockMiss++;
                if (blockSpots.Count < 6)
                    blockSpots.Add(new Dictionary<string, object> { { "para", k }, { "match", block == para }, { "block", Describe(block) }, { "paraText", Describe(para) } });
            }
            r["blockSpotMatches"] = (SpotParagraphs(total).Count - blockMiss) + "/" + SpotParagraphs(total).Count;
            r["blockSpotSample"] = blockSpots;
        }

        // ---- (2) the v1.13 path: one WordOpenXML fetch, parsed locally
        Stopwatch sw = Stopwatch.StartNew();
        string xml = null;
        try { xml = Convert.ToString(doc.Content.WordOpenXML, CultureInfo.InvariantCulture); }
        catch (Exception e) { r["xmlError"] = Clean(e.Message); }
        r["xmlFetchMs"] = sw.ElapsedMilliseconds;
        r["xmlChars"] = xml == null ? -1 : xml.Length;
        if (xml != null)
        {
            sw.Restart();
            int wP = -1;
            try
            {
                XmlDocument xd = new XmlDocument();
                xd.XmlResolver = null;
                xd.LoadXml(xml);
                XmlNamespaceManager nm = new XmlNamespaceManager(xd.NameTable);
                nm.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                XmlNodeList ps = xd.SelectNodes("//w:body//w:p", nm);
                wP = ps == null ? -1 : ps.Count;
                r["xmlParagraphs"] = wP;
                r["xmlPlusRows"] = wP + sumRows;
                r["xmlPlusRowsMatches"] = (wP + sumRows) == total;
                List<object> cmp = new List<object>();
                int shown = 0;
                for (int i = 0; i < ps.Count && shown < 8; i++)
                {
                    if (i + 1 > total) break;
                    string pt = "";
                    try { pt = CleanPara(Convert.ToString(doc.Paragraphs.Item(i + 1).Range.Text, CultureInfo.InvariantCulture)); }
                    catch (Exception) { }
                    if (pt.Length == 0) continue;
                    string xt = CleanPara(XmlParaTextOf(ps[i], nm));
                    cmp.Add(new Dictionary<string, object> { { "para", i + 1 }, { "match", xt == pt }, { "xml", Describe(xt) }, { "paraText", Describe(pt) } });
                    shown++;
                }
                r["xmlSample"] = cmp;
            }
            catch (Exception e) { r["xmlParseError"] = Clean(e.Message); }
            r["xmlParseMs"] = sw.ElapsedMilliseconds;
        }

        // ---- (3) verdicts from the real code paths
        string why;
        string[] fast = FastBlocksVerified(doc, total, a.Flag("slow", false), out why);
        r["blockModelFastPath"] = fast != null;
        r["blockModelFail"] = why;
        if (xml != null)
        {
            Stopwatch sw2 = Stopwatch.StartNew();
            string xmlWhy;
            List<string> texts = XmlDocTexts(doc, total, out xmlWhy);
            sw2.Stop();
            r["xmlFastPath"] = texts != null;
            r["xmlFail"] = xmlWhy;
            r["xmlTextsMs"] = sw2.ElapsedMilliseconds;
            if (texts != null) r["xmlNote"] = _xmlNote;
        }
        r["fastFail"] = why;
        return r;
    }

    /// <summary>Paragraph text of one w:p node (accept view, same mapping the fast path uses).</summary>
    private static string XmlParaTextOf(XmlNode p, XmlNamespaceManager nm)
    {
        StringBuilder sb = new StringBuilder();
        XmlNodeList ts = p.SelectNodes(".//w:t", nm);
        if (ts != null)
            foreach (XmlNode t in ts)
                sb.Append(t.InnerText);
        return sb.ToString();
    }

    /// <summary>Short, log-safe rendering of a paragraph for fastFail / textdiag messages.</summary>
    private static string Describe(string s)
    {
        if (s == null) return "null";
        s = CleanPara(s);
        if (s.Length > 24) s = s.Substring(0, 24) + "...";
        return "[" + s.Length + "]'" + s + "'";
    }
}
