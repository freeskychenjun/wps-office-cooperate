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

    private static object UnitsDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        bool apply = a.Flag("apply", false);
        bool normalize = a.Flag("normalize", false);

        List<string[]> units = UnitList(a);

        // ---- census (one text read; instant even on 300 pages)
        string text = Convert.ToString(TryGet(() => (object)doc.Content.Text, ""), CultureInfo.InvariantCulture);
        List<object> report = new List<object>();
        HashSet<string> present = new HashSet<string>(StringComparer.Ordinal);
        int candTotal = 0, uniTotal = 0, dangerLetter = 0, dangerDigit = 0;
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

            // decoys: a letter right before the unit that is NOT a legal unit prefix
            // (k/c/d/h/u/n/m + m forms km2/cm2/... which are units of their own), or a digit after it.
            string letterDanger = @"[A-Za-z-[kcdhunm]]" + Regex.Escape(needle) + @"(?![0-9])";
            foreach (Match m in Regex.Matches(text, letterDanger))
            {
                dangerLetter++;
                if (dangerSamples.Count < 8) dangerSamples.Add(DangerSample(text, m.Index, m.Length, "letter-prefixed (looks like an English word)"));
            }
            foreach (Match m in Regex.Matches(text, @"(?<![A-Za-z0-9])" + Regex.Escape(needle) + @"[0-9]"))
            {
                dangerDigit++;
                if (dangerSamples.Count < 8) dangerSamples.Add(DangerSample(text, m.Index, m.Length, "digit-suffixed (a number continues)"));
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
            { "letterPrefixedSkipped", dangerLetter },
            { "digitSuffixedSkipped", dangerDigit }
        };
        if (dangerSamples.Count > 0) r["skippedSamples"] = dangerSamples;

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
        int fixedTotal = 0, skipPrev = 0, skipNext = 0, failed = 0;
        List<object> perUnit = new List<object>();
        List<object> skipSamples = new List<object>();
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

                    string prev = s > 0 ? Convert.ToString(TryGet(() => (object)doc.Range(s - 1, s).Text, ""), CultureInfo.InvariantCulture) : "";
                    if (prev.Length > 0 && IsAsciiLetter(prev[0]))
                    {
                        skipPrev++;
                        if (skipSamples.Count < 6) AddSkipSample(doc, s, e, "letter before the unit - probably an English word, not a unit", skipSamples);
                        try { rng.SetRange(e, docEnd); } catch (Exception) { break; }
                        continue;
                    }
                    string next = e < docEnd - 1 ? Convert.ToString(TryGet(() => (object)doc.Range(e, e + 1).Text, ""), CultureInfo.InvariantCulture) : "";
                    if (next.Length > 0 && next[0] >= '0' && next[0] <= '9')
                    {
                        skipNext++;
                        if (skipSamples.Count < 6) AddSkipSample(doc, s, e, "a digit follows the unit - the number continues", skipSamples);
                        try { rng.SetRange(e, docEnd); } catch (Exception) { break; }
                        continue;
                    }

                    try
                    {
                        doc.Range(e - 1, e).Text = supCh;
                        fixedHere++;
                        fixedTotal++;
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
        r["skipped"] = new Dictionary<string, object> { { "letterPrefixed", skipPrev }, { "digitSuffixed", skipNext }, { "failed", failed } };
        if (skipSamples.Count > 0) r["skippedSamples"] = skipSamples;
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = fixedTotal == 0
            ? "nothing to fix: every unit found was already superscript (unicode or run format)."
            : "the exponent digits were replaced with unicode superscript characters (format-only change, one undo step). Render the pages you changed, then save.";
        return r;
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
