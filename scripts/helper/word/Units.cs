// Superscript unit repair (word unit / word units).
//
// Problem this solves: a Chinese engineering report writes "2248m3/s", "15300km2" etc.
// The exponent is sometimes a real superscript and sometimes "fallen down" (a plain digit).
// Three states exist in one document:
//   * unicode  : the text already holds the character 小 ²/³        -> correct, leave it
//   * format   : the digit carries run-level Font.Superscript      -> correct, leave it
//   * broken   : a plain ASCII 2/3                                 -> this is what we fix
//
// Fix = replace the plain digit with the unicode superscript character (same visual
// result, no run-format juggling, survives copy/paste and further text edits).
//
// Why per-hit and not one Replace-All: a blind replace also rewrites English-looking
// decoys ("item3" -> "ite³", "m23" -> "m²3"). Guards per hit:
//   * the character before the unit must not be an ASCII letter  (item3, inner m2 of km2)
//   * the character after the unit must not be an ASCII digit     (m23, km20)
// Longest units are processed first, so "km2" is fixed as a whole and its inner "m2"
// hit is skipped by the previous-char guard.
//
// Find is filtered with Font.Superscript = 0, so already-correct (format) instances never
// surface; unicode instances are different text and never match either.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

internal static partial class Program
{
    // unit -> exponent digit (2 or 3); metric area/volume units of a water report.
    // Only units that are genuinely used here. Anything else that merely LOOKS like a unit
    // (e.g. "km3" for 立方公里, which in these reports is a typo for km2) is reported as a
    // suspect for a human to judge - never silently superscripted, or the typo would come
    // out looking deliberate.
    private static readonly string[][] UNIT_TABLE = new string[][]
    {
        new string[] { "km2", "2" },
        new string[] { "hm2", "2" },
        new string[] { "cm2", "2" },
        new string[] { "mm2", "2" },
        new string[] { "dm2", "2" },
        new string[] { "um2", "2" },
        new string[] { "nm2", "2" },
        new string[] { "m3",  "3" },
        new string[] { "m2",  "2" },
    };

    // tokens that are almost always a typo in this kind of report: reported, never fixed
    private static readonly string[] SUSPECT_TOKENS = new string[] { "km3", "km4", "m4", "hm3" };

    // A unit is only fixed when it hangs off a value ("500km2", "1.2亿m3", "面积（km2）").
    // An isolated token ("km2 = 1", "<m1，m2，m3>") is a variable or label in a template/script
    // legend - measured on a real report: those were damaged by an earlier build of this
    // command, which is exactly what these guards and the fragment report exist to prevent.
    private const string CTX_FIX = null;
    private const string CTX_ISOLATED = "isolated (no value in front - looks like a variable or label)";
    private const string CTX_ASSIGN = "followed by = (assignment or label)";
    private const string CTX_ENGLISH = "letter in front (looks like an English word)";
    private const string CTX_DIGIT = "digit after (the number continues)";
    private const string CTX_SCRIPT = "inside a template/script example ([for], assignment lines, variable lists)";

    // Markers of a template/script example block (a real report had a section documenting a
    // document-template language: [for] ... [/for], k~a = 1, m1 = 1 / n2 = 1 lines).
    private static readonly Regex SCRIPT_TAG = new Regex(@"\[/?(?:for|FOR|resource|Resource|RESOURCE|File|file|if|IF|else)\b[^\]]*\]", RegexOptions.Compiled);
    private static readonly Regex SCRIPT_ASSIGN = new Regex(@"(?:^|\s)[A-Za-z][A-Za-z0-9]*(?:~[A-Za-z])?\s*=\s*[0-9]", RegexOptions.Compiled);
    private static readonly Regex SCRIPT_TILDE = new Regex(@"[A-Za-z]~[A-Za-z]", RegexOptions.Compiled);

    /// <summary>Does this paragraph belong to a template/script example? (Its units are variables.)</summary>
    private static bool IsScriptParagraph(dynamic doc, int start, int end)
    {
        string text = "";
        try
        {
            int from = Math.Max(0, start - 400);
            int to = Math.Min((int)doc.Content.End, end + 400);
            text = Convert.ToString(TryGet(() => (object)doc.Range(from, to).Text, ""), CultureInfo.InvariantCulture);
        }
        catch (Exception) { return false; }
        if (text.Length == 0) return false;
        if (SCRIPT_TAG.IsMatch(text)) return true;                 // [for] / [/resource] / [File] ...
        if (SCRIPT_TILDE.IsMatch(text)) return true;               // k~a
        if (SCRIPT_ASSIGN.Matches(text).Count >= 2) return true;   // two or more "x = 1" assignments
        return false;
    }

    /// <summary>Pages that contain a script/template example marker (for the report).</summary>
    private static List<int> ScriptPagesOf(dynamic doc)
    {
        List<int> pages = new List<int>();
        int docEnd = (int)TryGet(() => (object)doc.Content.End, 0);
        string[] markers = new string[] { "[for]", "[/for]", "[FOR]", "[File]", "[Resource]", "[/resource]" };
        foreach (string marker in markers)
        {
            dynamic rng = null;
            try { rng = doc.Content.Duplicate; } catch (Exception) { continue; }
            int guard = 0;
            while (guard++ < 60)
            {
                try { rng.SetRange((int)rng.Start, docEnd); } catch (Exception) { break; }
                dynamic find = rng.Find;
                find.ClearFormatting();
                find.Text = marker;
                find.MatchCase = true;
                find.MatchWildcards = false;
                find.Forward = true;
                find.Wrap = 0;
                bool hit;
                try { hit = find.Execute(); } catch (Exception) { break; }
                if (!hit) break;
                int at = (int)rng.Start;
                int page = Convert.ToInt32(TryGet(() => (object)doc.Range(at, at).Information(3), (object)0), CultureInfo.InvariantCulture);
                if (page > 0 && !pages.Contains(page)) pages.Add(page);
                if ((int)rng.End <= at) break;
            }
        }
        pages.Sort();
        return pages;
    }

    /// <summary>Decide whether the token at [start,end) is a real unit in its context.</summary>
    private static string UnitContext(string text, int start, int end)
    {
        int len = text.Length;
        char prev = start > 0 ? text[start - 1] : ' ';
        char prev2 = start > 1 ? text[start - 2] : ' ';
        char next = end < len ? text[end] : ' ';
        if (next >= '0' && next <= '9') return CTX_DIGIT;
        if (IsAsciiLetter(prev)) return CTX_ENGLISH;
        int k = end;
        while (k < len && (text[k] == ' ' || text[k] == '\t' || text[k] == '\u3000')) k++;
        char nextNonSpace = k < len ? text[k] : ' ';
        if (nextNonSpace == '=' || nextNonSpace == '\uFF1D') return CTX_ASSIGN;
        if (prev >= '0' && prev <= '9') return CTX_FIX;
        if (prev == '\u4E07' || prev == '\u4EBF') return CTX_FIX;          // 万 / 亿
        if (prev == '(' || prev == '\uFF08') return CTX_FIX;               // ( （
        if (IsCjkIdeograph(prev)) return CTX_FIX;                          // 汉字 right before ("面积km2")
        if ((prev == ' ' || prev == '\u3000') && prev2 >= '0' && prev2 <= '9') return CTX_FIX;  // "500 km2"
        return CTX_ISOLATED;
    }

    /// <summary>Han characters only - NOT CJK punctuation: a comma before the token means it is
    /// a list member ("&lt;m1，m2，m3&gt;"), not a value.</summary>
    private static bool IsCjkIdeograph(char c)
    {
        return (c >= '\u4E00' && c <= '\u9FFF') || (c >= '\u3400' && c <= '\u4DBF') || (c >= '\uF900' && c <= '\uFAFF');
    }

    private static object UnitsDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        bool apply = a.Flag("apply", false);
        bool normalize = a.Flag("normalize", false);

        List<string[]> units = UnitList(a);

        // ---- census (one text read; instant even on 300 pages)
        string text = Convert.ToString(TryGet(() => (object)doc.Content.Text, ""), CultureInfo.InvariantCulture);
        List<int> scriptPages = ScriptPagesOf(doc);
        List<object> report = new List<object>();
        HashSet<string> present = new HashSet<string>(StringComparer.Ordinal);
        int candTotal = 0, uniTotal = 0, dangerLetter = 0, dangerDigit = 0;
        int fixable = 0, isolated = 0, assignment = 0;
        List<object> fragmentSamples = new List<object>();
        Dictionary<string, int> fragmentByToken = new Dictionary<string, int>(StringComparer.Ordinal);
        List<object> dangerSamples = new List<object>();
        foreach (string[] u in units)
        {
            string needle = u[0];
            string exp = u[1];
            string uni = needle.Substring(0, needle.Length - 1) + (exp == "3" ? "\u00B3" : "\u00B2");
            int cand = CountOutside(text, needle, exp);
            int uniN = CountLiteral(text, uni);
            candTotal += cand;
            uniTotal += uniN;
            if (cand > 0) present.Add(needle);   // units absent from the text cost a full scan each

            // classify every candidate: a real unit hangs off a value; an isolated token is a
            // variable/label in a script legend (measured: those must NOT be touched)
            if (cand > 0)
            {
                Regex rxAll = new Regex(@"(?<![A-Za-z])" + Regex.Escape(needle) + @"(?![0-9])");
                foreach (Match mm in rxAll.Matches(text))
                {
                    string why = UnitContext(text, mm.Index, mm.Index + mm.Length);
                    if (why == CTX_FIX) { fixable++; continue; }
                    if (why == CTX_ASSIGN) assignment++;
                    else if (why == CTX_ISOLATED) isolated++;
                    else if (why == CTX_ENGLISH) dangerLetter++;
                    else if (why == CTX_DIGIT) dangerDigit++;
                    if (why == CTX_ISOLATED || why == CTX_ASSIGN)
                    {
                        if (!fragmentByToken.ContainsKey(needle)) fragmentByToken[needle] = 0;
                        fragmentByToken[needle] = fragmentByToken[needle] + 1;
                        if (fragmentSamples.Count < 8) fragmentSamples.Add(DangerSample(text, mm.Index, mm.Length, why));
                    }
                }
            }

            Dictionary<string, object> one = new Dictionary<string, object>();
            one["unit"] = needle;
            one["textHits"] = cand;       // plain-digit hits including already-format-superscript ones
            one["unicodeOk"] = uniN;      // already written with the ²/³ character
            report.Add(one);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["action"] = apply ? "apply" : "scan";
        r["pages"] = TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        r["characters"] = text.Length;
        r["units"] = report;
        r["summary"] = new Dictionary<string, object> {
            { "plainDigitHits", candTotal },
            { "unicodeOk", uniTotal },
            { "unitContext", fixable },
            { "isolated", isolated },
            { "assignment", assignment },
            { "letterPrefixedSkipped", dangerLetter },
            { "digitSuffixedSkipped", dangerDigit }
        };
        if (fragmentSamples.Count > 0)
        {
            List<object> frag = new List<object>();
            foreach (KeyValuePair<string, int> kv in fragmentByToken)
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["token"] = kv.Key;
                one["count"] = kv.Value;
                frag.Add(one);
            }
            r["fragments"] = frag;
            r["fragmentSamples"] = fragmentSamples;
            r["fragmentsNote"] = "isolated unit-shaped tokens: usually variable names or labels in a template/script legend (e.g. \"km2 = 1\" inside a [for] example). They are NOT fixed - adding a superscript there would corrupt the document. Read the samples, then fix real units by hand if any of them is one.";
        }

        // fragments that are not units at all but look like a unit whose base got lost:
        // "n2" / "n3" next to "m2" / "m3" in the same list is a strong hint of that
        List<object> loose = new List<object>();
        foreach (string probe in new string[] { "n2", "n3", "k2", "h2", "c2", "d2", "u2" })
        {
            Regex rxp = new Regex(@"(?<![A-Za-z0-9])" + Regex.Escape(probe) + @"(?![0-9A-Za-z])");
            MatchCollection hits = rxp.Matches(text);
            if (hits.Count == 0) continue;
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["token"] = probe;
            one["count"] = hits.Count;
            one["likely"] = "m" + probe.Substring(1) + " / m" + (probe.Substring(1) == "2" ? "\u00B2" : "\u00B3") + " ?";
            List<object> ctx = new List<object>();
            foreach (Match m in hits) { if (ctx.Count < 3) ctx.Add(DangerSample(text, m.Index, m.Length, "unit-shaped fragment")); }
            one["samples"] = ctx;
            loose.Add(one);
        }
        if (loose.Count > 0)
        {
            r["looseFragments"] = loose;
            r["looseNote"] = "unit-shaped fragments with no unit base (n2, k2 ...): often a typo for m2/m3, or a variable name. Reported only - confirm with the author.";
        }

        if (scriptPages.Count > 0)
        {
            r["scriptPages"] = scriptPages;
            r["scriptPagesNote"] = "pages that contain a template/script example ([for] ... [/for], [File], k~a = 1 assignment lines). Units on those paragraphs are variable names, so they are never fixed; only reported.";
        }

        // Suspects: look like units but are almost certainly a typo (km3 for 立方公里 in a
        // report about areas is a typo for km2). Report them; do not touch them.
        List<object> suspects = new List<object>();
        foreach (string tok in SUSPECT_TOKENS)
        {
            Regex rx = new Regex(@"(?<![A-Za-z0-9])" + Regex.Escape(tok) + @"(?![0-9])");
            MatchCollection hits = rx.Matches(text);
            if (hits.Count == 0) continue;
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["token"] = tok;
            one["count"] = hits.Count;
            one["likely"] = tok.StartsWith("km", StringComparison.Ordinal) ? "km2 / km\u00B2" : "check by hand";
            List<object> ctx = new List<object>();
            foreach (Match m in hits) { if (ctx.Count < 4) ctx.Add(DangerSample(text, m.Index, m.Length, "suspect token")); }
            one["samples"] = ctx;
            suspects.Add(one);
        }
        if (suspects.Count > 0)
        {
            r["suspects"] = suspects;
            r["suspectsNote"] = "these look like typos (e.g. km3 in an area table is probably km2), so they are NOT changed automatically - "
                              + "the exponent would become a superscript 3 and the typo would look deliberate. Confirm with the author, "
                              + "then fix the digit with revise/table cell, or re-run with -Units \"<token>\" only if it really is a cubic unit.";
        }

        if (!apply)
        {
            r["note"] = "scan only. Run the same command with -Apply to fix. A plain-digit hit is only fixed when the digit is not already a superscript; "
                      + "the exact count of broken ones is reported after -Apply (that is when Find can see the formatting).";
            return r;
        }

        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no unit was changed.");

        // ---- apply: per-hit, filtered, guarded
        dynamic app = App(true, true);
        int docEnd = (int)TryGet(() => (object)doc.Content.End, 0);
        int fixedTotal = 0, skipPrev = 0, skipNext = 0, skipAssign = 0, skipIsolated = 0, skipScript = 0, failed = 0;
        List<int> scriptPageHits = new List<int>();
        List<object> perUnit = new List<object>();
        List<object> skipSamples = new List<object>();
        List<object> fixedOnScriptPage = new List<object>();
        bool wasTracking = false;
        try { wasTracking = (bool)doc.TrackRevisions; } catch (Exception) { }
        bool track = a.Flag("track", false);

        using (P2Edit edit = new P2Edit(doc, track, "wps-word-cooperate superscript units"))
        {
            foreach (string[] u in units)
            {
                string needle = u[0];
                if (!present.Contains(needle)) continue;   // no candidate in the text: nothing to scan
                string supCh = u[1] == "3" ? "\u00B3" : "\u00B2";
                int fixedHere = 0;
                dynamic rng = null;
                try { rng = doc.Content.Duplicate; } catch (Exception) { rng = null; }
                if (rng == null) continue;
                while (true)
                {
                    if (fixedTotal > 0 && (fixedTotal % 200) == 0) Tr("units: " + fixedTotal + " fixed so far");
                    try { rng.SetRange((int)rng.Start, docEnd); } catch (Exception) { break; }
                    dynamic find = rng.Find;
                    find.ClearFormatting();
                    find.Font.Superscript = 0;        // broken (non-superscript) only
                    find.Text = needle;
                    find.MatchCase = true;
                    find.MatchWildcards = false;
                    find.Forward = true;
                    find.Wrap = 0;
                    bool hit;
                    try { hit = find.Execute(); } catch (Exception) { break; }
                    if (!hit) break;
                    int s = (int)rng.Start;
                    int e = (int)rng.End;
                    if (e <= s) break;

                    string why = ClassifyLive(doc, s, e, docEnd);
                    if (why != CTX_FIX)
                    {
                        if (why == CTX_ENGLISH) skipPrev++;
                        else if (why == CTX_DIGIT) skipNext++;
                        else if (why == CTX_ASSIGN) skipAssign++;
                        else if (why == CTX_SCRIPT) skipScript++;
                        else skipIsolated++;
                        if (skipSamples.Count < 6) AddSkipSample(doc, s, e, why, skipSamples);
                        try { rng.SetRange(e, docEnd); } catch (Exception) { break; }
                        continue;
                    }

                    try
                    {
                        doc.Range(e - 1, e).Text = supCh;
                        fixedHere++;
                        fixedTotal++;
                        if (scriptPages.Count > 0)
                        {
                            int page = Convert.ToInt32(TryGet(() => (object)doc.Range(s, s).Information(3), (object)0), CultureInfo.InvariantCulture);
                            if (page > 0 && scriptPages.Contains(page) && fixedOnScriptPage.Count < 6)
                                AddSkipSample(doc, s, e, "fixed on a page that also has a script example - check by hand", fixedOnScriptPage);
                        }
                    }
                    catch (Exception) { failed++; }
                    try { rng.SetRange(e, docEnd); } catch (Exception) { break; }
                }
                if (fixedHere > 0)
                {
                    Dictionary<string, object> one = new Dictionary<string, object>();
                    one["unit"] = needle;
                    one["fixed"] = fixedHere;
                    perUnit.Add(one);
                }
            }
        }

        try { doc.TrackRevisions = wasTracking; } catch (Exception) { }

        r["fixed"] = fixedTotal;
        r["fixedPerUnit"] = perUnit;
        r["skipped"] = new Dictionary<string, object> {
            { "letterPrefixed", skipPrev },
            { "digitSuffixed", skipNext },
            { "assignment", skipAssign },
            { "isolated", skipIsolated },
            { "script", skipScript },
            { "failed", failed }
        };
        if (fixedOnScriptPage.Count > 0) r["fixedOnScriptPage"] = fixedOnScriptPage;
        if (skipSamples.Count > 0) r["skippedSamples"] = skipSamples;
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = fixedTotal == 0
            ? "nothing to fix: every unit found was already superscript (unicode or run format)."
            : "the exponent digits were replaced with unicode superscript characters (format-only change, one undo step). Render the pages you changed, then save.";
        return r;
    }

    /// <summary>Same decision as UnitContext, but read from the live document (text offsets and
    /// Range positions do not line up in documents that contain tables).</summary>
    private static string ClassifyLive(dynamic doc, int start, int end, int docEnd)
    {
        // A template/script example wins over everything else: its "units" are variable names.
        if (IsScriptParagraph(doc, start, end)) return CTX_SCRIPT;
        string before = "";
        try { int from = Math.Max(0, start - 2); before = Convert.ToString(doc.Range(from, start).Text ?? "", CultureInfo.InvariantCulture); }
        catch (Exception) { }
        char prev = before.Length > 0 ? before[before.Length - 1] : ' ';
        char prev2 = before.Length > 1 ? before[before.Length - 2] : ' ';
        string after = "";
        try { int to = Math.Min(docEnd, end + 6); after = Convert.ToString(doc.Range(end, to).Text ?? "", CultureInfo.InvariantCulture); }
        catch (Exception) { }
        char next = after.Length > 0 ? after[0] : ' ';
        if (next >= '0' && next <= '9') return CTX_DIGIT;
        if (IsAsciiLetter(prev)) return CTX_ENGLISH;
        int k = 0;
        while (k < after.Length && (after[k] == ' ' || after[k] == '\t' || after[k] == '\u3000')) k++;
        char nextNonSpace = k < after.Length ? after[k] : ' ';
        if (nextNonSpace == '=' || nextNonSpace == '\uFF1D') return CTX_ASSIGN;
        if (prev >= '0' && prev <= '9') return CTX_FIX;
        if (prev == '\u4E07' || prev == '\u4EBF') return CTX_FIX;
        if (prev == '(' || prev == '\uFF08') return CTX_FIX;
        if (IsCjkIdeograph(prev)) return CTX_FIX;
        if ((prev == ' ' || prev == '\u3000') && prev2 >= '0' && prev2 <= '9') return CTX_FIX;
        return CTX_ISOLATED;
    }

    private static void AddSkipSample(dynamic doc, int start, int end, string why, List<object> into)
    {
        string ctx = "";
        try
        {
            int from = Math.Max(0, start - 14);
            int to = end + 10;
            ctx = Convert.ToString(TryGet(() => (object)doc.Range(from, to).Text, ""), CultureInfo.InvariantCulture);
        }
        catch (Exception) { }
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["text"] = ctx.Replace("\r", " ").Replace("\n", " ").Replace("\a", " ").Replace("\u0007", " ").Trim();
        d["why"] = why;
        into.Add(d);
    }

    private static List<string[]> UnitList(Bag a)
    {
        List<string[]> list = new List<string[]>();
        string spec = a.Str("units", null);
        if (spec == null || spec.Trim().Length == 0)
        {
            foreach (string[] u in UNIT_TABLE) list.Add(u);
            return list;
        }
        foreach (string raw in spec.Split(','))
        {
            string one = raw.Trim().ToLowerInvariant();
            if (one.Length < 2) continue;
            char last = one[one.Length - 1];
            if (last != '2' && last != '3')
                throw new Fail("BAD_ARGS", "unit \"" + raw.Trim() + "\" must end with the exponent digit 2 or 3 (e.g. m3, km2).");
            bool known = false;
            foreach (string[] u in UNIT_TABLE) if (u[0] == one) known = true;
            if (!known && one.Length == 2)
                throw new Fail("BAD_ARGS", "unit \"" + raw.Trim() + "\" is not a known unit; give a full base name such as km2 or cm2.");
            list.Add(new string[] { one, last.ToString() });
        }
        if (list.Count == 0) throw new Fail("BAD_ARGS", "\"units\" has no usable unit (e.g. \"m3,km2\").");
        // longest first, so km2 is handled before its inner m2
        list.Sort((x, y) => y[0].Length.CompareTo(x[0].Length));
        return list;
    }

    private static bool IsAsciiLetter(char c)
    {
        return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
    }

    /// <summary>plain-digit candidates of one unit, ignoring tokens glued to a letter or a number.</summary>
    private static int CountOutside(string text, string needle, string exp)
    {
        Regex rx = new Regex(@"(?<![A-Za-z])" + Regex.Escape(needle) + @"(?![0-9])");
        return rx.Matches(text).Count;
    }

    private static int CountLiteral(string text, string literal)
    {
        if (literal.Length == 0) return 0;
        // a letter right before it means it belongs to a longer unit (m² inside km²)
        Regex rx = new Regex(@"(?<![A-Za-z])" + Regex.Escape(literal));
        return rx.Matches(text).Count;
    }

    private static Dictionary<string, object> DangerSample(string text, int index, int length, string why)
    {
        int from = Math.Max(0, index - 12);
        int len = Math.Min(text.Length - from, length + 24);
        string ctx = text.Substring(from, len).Replace("\r", " ").Replace("\n", " ").Replace("\a", " ");
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["text"] = ctx;
        d["why"] = why;
        return d;
    }
}
