// wps-office-cooperate (ppt): `replace` - find / replace text across a whole deck.
//
// Why this file exists: the deck has no revision layer, so a batch text edit is the
// only automation the user cannot review afterwards by pressing Ctrl+Z. Everything
// here therefore leans on three rules:
//   1. walk EVERY text container (top-level shapes, groups, table cells, notes,
//      headers/footers, optionally the slide master and its layouts);
//   2. before writing a span, read it back through Characters(start,len) and compare:
//      a span that does not read back exactly as matched is reported, never rewritten;
//   3. hand back a text-level snapshot (same folder as `snapshot`) that `restore
//      -Since prereplace` can roll back.
//
// Indexing facts this file relies on (measured on WPS 12.1.0.29161, not guessed):
//   - TextRange.Characters(Start, Length) is 1-based and counts the paragraph mark
//     as one character, so an index into TextRange.Text maps straight onto it;
//   - assigning .Text on a Characters range keeps Font.Size / .Name / .NameFarEast /
//     .Bold / .Color.RGB of the text around it, and works even when the range spans
//     two differently formatted runs (the new text takes the formatting at the start
//     of the range);
//   - replacements are applied from the BACK, because Characters indexes live text;
//   - HeadersFooters.Footer/Header/DateAndTime expose a plain .Text string (no
//     Characters), and their .Text only sticks once .Visible is -1 - so the footer
//     scope rewrites the whole string and skips hidden ones.
//
// Pure ASCII source (Chinese only ever travels inside JSON strings), C# 5.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

internal static partial class Program
{
    /// <summary>Scopes searched when the caller does not say. The slide master is
    /// deliberately NOT here: a master edit ripples into every slide, so it must be
    /// asked for by name.</summary>
    private static readonly string[] DefaultReplaceScopes = new string[] { "slides", "groups", "tables", "notes", "footer" };

    // ---------------------------------------------------------------- types

    private sealed class ReplaceRule
    {
        public string Label = "";
        public string Find = "";
        public string Repl = "";
        public bool IsRegex;
        public bool CaseSensitive;
        public Regex Compiled;
    }

    private sealed class Span
    {
        public int Start0;
        public int Len;
        public string Repl = "";
    }

    /// <summary>Font of the first character of a replaced span, captured before the
    /// write and put back afterwards.</summary>
    private sealed class FontSnap
    {
        public float Size;
        public string Name;
        public string NameFarEast;
        public int? Bold;
        public int? Italic;
        public int? Underline;
        public int? Rgb;
    }

    private sealed class Counts
    {
        public int Slides, Shapes, Groups, Tables, Cells, Notes, Footers, Master, TextSlots;
    }

    /// <summary>One addressable piece of text in the deck. Either a real TextRange
    /// (Frame) or a plain string property (Plain, used for header/footer/date).</summary>
    private sealed class TextSlot
    {
        public int Slide;          // 1-based; 0 = slide master / layout
        public string Where = "body";
        public string Path = "";   // stable identity inside the deck
        public string Name = "";
        public dynamic Frame;      // shape: .TextFrame.TextRange
        public dynamic Plain;      // header/footer/date object: .Text string

        public string Text()
        {
            if (Plain != null) return Convert.ToString(Plain.Text, CultureInfo.InvariantCulture);
            if (Frame == null) return "";
            try
            {
                dynamic tf = Frame.TextFrame;
                if ((int)tf.HasText == 0) return "";
                return Convert.ToString(tf.TextRange.Text, CultureInfo.InvariantCulture);
            }
            catch (Exception) { return ""; }
        }

        public dynamic FreshRange()
        {
            return Frame.TextFrame.TextRange;
        }

        public void SetWhole(string value)
        {
            if (Plain != null) { Plain.Text = value; return; }
            Frame.TextFrame.TextRange.Text = value;
        }
    }

    // ---------------------------------------------------------------- command

    private static object ReplaceCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        bool dryRun = a.Flag("dryRun", false);
        int max = a.Int("max", 500);
        if (max <= 0) max = 500;

        List<ReplaceRule> rules = BuildReplaceRules(a);      // validates everything before any write
        HashSet<string> scopes = ParseReplaceScopes(a.Str("scope", null));
        HashSet<int> only = ParseSlideList(a.Str("slides", null));

        Counts counts = new Counts();
        List<TextSlot> slots = CollectTextSlots(p, scopes, only, counts);

        string snapPath = null;
        if (!dryRun) snapPath = WriteSlotSnapshot(p, slots);

        List<object> hits = new List<object>();
        List<object> skipped = new List<object>();
        List<object> failed = new List<object>();
        HashSet<string> touched = new HashSet<string>(StringComparer.Ordinal);
        int changed = 0;
        bool capped = false;

        foreach (ReplaceRule rule in rules)
        {
            foreach (TextSlot slot in slots)
            {
                if (changed >= max) { capped = true; break; }
                string text;
                try { text = slot.Text(); } catch (Exception) { continue; }
                if (text.Length == 0) continue;
                List<Span> spans = FindSpans(text, rule, max - changed);
                if (spans.Count == 0) continue;
                ApplySpans(slot, rule, text, spans, dryRun, hits, skipped, failed, touched, ref changed);
            }
            if (capped) break;
        }

        Dictionary<string, object> scanned = new Dictionary<string, object>();
        scanned["slides"] = counts.Slides;
        scanned["shapes"] = counts.Shapes;
        scanned["groups"] = counts.Groups;
        scanned["tables"] = counts.Tables;
        scanned["cells"] = counts.Cells;
        scanned["notes"] = counts.Notes;
        scanned["footers"] = counts.Footers;
        scanned["master"] = counts.Master;
        scanned["textSlots"] = counts.TextSlots;

        List<object> ruleReport = new List<object>();
        foreach (ReplaceRule rule in rules)
            ruleReport.Add(new Dictionary<string, object> {
                { "find", rule.Find }, { "replace", rule.Repl },
                { "regex", rule.IsRegex }, { "case", rule.CaseSensitive }, { "label", rule.Label }
            });

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["scanned"] = scanned;
        r["rules"] = ruleReport;
        r["scope"] = new List<object>(new List<string>(scopes).ToArray());
        r["changed"] = changed;
        r["shapesTouched"] = touched.Count;
        r["applied"] = !dryRun;
        if (capped) r["capped"] = true;
        if (hits.Count > 0) r["hits"] = hits;
        if (skipped.Count > 0) r["skipped"] = skipped;
        if (failed.Count > 0) r["failed"] = failed;
        if (snapPath != null) r["preSnapshot"] = snapPath;
        r["note"] = (dryRun ? "dry run: nothing was written. " : "snapshot taken before the batch; roll back with restore -Since prereplace. ")
                   + "text only - run diff and render before saving";
        if (changed == 0 && skipped.Count == 0 && failed.Count == 0) r["note"] = "no match: check -Case, -Scope and the exact characters (full-width vs half-width)";
        return r;
    }

    // ---------------------------------------------------------------- rules

    private static List<ReplaceRule> BuildReplaceRules(Bag a)
    {
        List<ReplaceRule> rules = new List<ReplaceRule>();
        if (a.Has("items"))
        {
            foreach (object o in a.List("items"))
            {
                Bag r = new Bag(o);
                rules.Add(MakeReplaceRule(
                    r.Str("find", null),
                    r.Has("replace") ? r.Str("replace", "") : "",
                    r.Flag("regex", a.Flag("regex", false)),
                    r.Flag("case", a.Flag("case", false)),
                    r.Str("label", null)));
            }
        }
        else
        {
            rules.Add(MakeReplaceRule(
                a.Str("find", null),
                a.Has("replace") ? a.Str("replace", "") : "",
                a.Flag("regex", false),
                a.Flag("case", false),
                a.Str("label", null)));
        }
        if (rules.Count == 0) throw new Fail("BAD_ARGS", "replace needs -Find <text> or -Items <json>");
        if (rules.Count > 200) throw new Fail("BAD_ARGS", "too many replace rules (" + rules.Count + "); split the batch");
        return rules;
    }

    private static ReplaceRule MakeReplaceRule(string find, string repl, bool isRegex, bool caseSensitive, string label)
    {
        if (find == null || find.Length == 0) throw new Fail("BAD_ARGS", "replace: \"find\" must not be empty (use -Items for several rules)");
        ReplaceRule rule = new ReplaceRule();
        rule.Find = find;
        rule.Repl = repl == null ? "" : repl;
        rule.IsRegex = isRegex;
        rule.CaseSensitive = caseSensitive;
        rule.Label = label == null || label.Length == 0 ? Clip(find, 40) : label;
        if (isRegex)
        {
            try { rule.Compiled = new Regex(find, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase); }
            catch (Exception e) { throw new Fail("BAD_ARGS", "bad -Regex pattern \"" + find + "\": " + Clean(e.Message)); }
        }
        return rule;
    }

    private static HashSet<string> ParseReplaceScopes(string spec)
    {
        HashSet<string> want = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (spec == null || spec.Trim().Length == 0)
        {
            foreach (string s in DefaultReplaceScopes) want.Add(s);
            return want;
        }
        foreach (string raw in spec.Split(','))
        {
            string v = raw.Trim().ToLowerInvariant();
            if (v.Length == 0) continue;
            if (v == "all") { foreach (string s in DefaultReplaceScopes) want.Add(s); want.Add("master"); }
            else if (v == "slides" || v == "body") want.Add("slides");
            else if (v == "groups" || v == "group") want.Add("groups");
            else if (v == "tables" || v == "table") want.Add("tables");
            else if (v == "notes") want.Add("notes");
            else if (v == "footer" || v == "footers" || v == "headerfooter") want.Add("footer");
            else if (v == "master" || v == "layout" || v == "layouts") want.Add("master");
            else throw new Fail("BAD_ARGS", "unknown -Scope value \"" + raw.Trim() + "\"; use all|slides|groups|tables|notes|footer|master");
        }
        return want;
    }

    private static HashSet<int> ParseSlideList(string spec)
    {
        if (spec == null || spec.Trim().Length == 0) return null;
        HashSet<int> list = new HashSet<int>();
        foreach (string raw in spec.Split(','))
        {
            string v = raw.Trim();
            if (v.Length == 0) continue;
            int dash = v.IndexOf('-');
            if (dash > 0)
            {
                int lo, hi;
                if (!int.TryParse(v.Substring(0, dash).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out lo)
                    || !int.TryParse(v.Substring(dash + 1).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out hi))
                    throw new Fail("BAD_ARGS", "bad -Slides range \"" + v + "\"; use e.g. 1-5,8");
                if (hi < lo) { int t = lo; lo = hi; hi = t; }
                for (int i = lo; i <= hi; i++) if (i >= 1) list.Add(i);
            }
            else
            {
                int one;
                if (!int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out one))
                    throw new Fail("BAD_ARGS", "bad -Slides value \"" + v + "\"; use e.g. 1-5,8");
                if (one >= 1) list.Add(one);
            }
        }
        if (list.Count == 0) throw new Fail("BAD_ARGS", "-Slides matched no slide: " + spec);
        return list;
    }

    // ---------------------------------------------------------------- walk the deck

    private static List<TextSlot> CollectTextSlots(dynamic p, HashSet<string> scopes, HashSet<int> only, Counts c)
    {
        List<TextSlot> slots = new List<TextSlot>();

        if (scopes.Contains("master"))
        {
            try
            {
                dynamic m = p.SlideMaster;
                WalkTextShapes(m.Shapes, 0, "master:", "master", false, true, true, true, slots, c);
                int layouts = 0;
                try { layouts = (int)m.CustomLayouts.Count; } catch (Exception) { }
                for (int i = 1; i <= layouts; i++)
                {
                    dynamic lay = null;
                    try { lay = m.CustomLayouts.Item(i); } catch (Exception) { continue; }
                    dynamic shapes = null;
                    try { shapes = lay.Shapes; } catch (Exception) { continue; }
                    WalkTextShapes(shapes, 0, "layout" + i + ":", "master", false, true, true, true, slots, c);
                }
            }
            catch (Exception) { }
        }

        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }

        for (int k = 1; k <= n; k++)
        {
            if (only != null && !only.Contains(k)) continue;
            dynamic s = null;
            try { s = p.Slides.Item(k); } catch (Exception) { continue; }
            c.Slides++;

            try
            {
                WalkTextShapes(s.Shapes, k, "", "body", false,
                    scopes.Contains("slides"), scopes.Contains("groups"), scopes.Contains("tables"), slots, c);
            }
            catch (Exception) { }

            if (scopes.Contains("notes"))
            {
                try
                {
                    dynamic pages = s.NotesPage;
                    WalkTextShapes(pages.Shapes, k, "notes:", "notes", true, true, true, true, slots, c);
                }
                catch (Exception) { }
            }

            if (scopes.Contains("footer"))
            {
                try
                {
                    dynamic hf = s.HeadersFooters;
                    AddPlainSlot(slots, k, "footer:footer", "footer", hf.Footer, c);
                    AddPlainSlot(slots, k, "footer:header", "header", hf.Header, c);
                    AddPlainSlot(slots, k, "footer:date", "date", hf.DateAndTime, c);
                }
                catch (Exception) { }
            }
        }

        // per-scope slot counts, for the report: a "how much did it even look at" line
        foreach (TextSlot slot in slots)
        {
            if (slot.Where == "notes") c.Notes++;
            else if (slot.Where == "footer") c.Footers++;
            else if (slot.Where == "master") c.Master++;
        }
        return slots;
    }

    private static void AddPlainSlot(List<TextSlot> slots, int slide, string path, string name, dynamic obj, Counts c)
    {
        if (obj == null) return;
        bool visible = false;
        try { visible = (int)obj.Visible != 0; } catch (Exception) { visible = false; }
        if (!visible) return;   // hidden header/footer has no text worth touching, and
                                // showing it would silently change the deck's layout
        TextSlot slot = new TextSlot();
        slot.Slide = slide;
        slot.Where = "footer";
        slot.Path = path;
        slot.Name = name;
        slot.Plain = obj;
        slots.Add(slot);
        c.TextSlots++;
    }

    private static void WalkTextShapes(dynamic shapes, int slide, string prefix, string where, bool notesMode,
        bool allowText, bool allowGroup, bool allowTable, List<TextSlot> slots, Counts c)
    {
        int count = 0;
        try { count = (int)shapes.Count; } catch (Exception) { return; }

        for (int i = 1; i <= count; i++)
        {
            dynamic sh = null;
            try { sh = shapes.Item(i); } catch (Exception) { continue; }
            c.Shapes++;
            string name = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
            int type = 0;
            try { type = Convert.ToInt32(TryGet(() => (object)sh.Type, 0), CultureInfo.InvariantCulture); } catch (Exception) { }

            if (type == 6)  // group
            {
                if (!allowGroup) continue;
                c.Groups++;
                dynamic items = null;
                try { items = sh.GroupItems; } catch (Exception) { continue; }
                // A group's contents are their own scope; a text inside a group on the
                // slide master still belongs to the master's report, not to "group".
                string inner = (where == "master") ? "master" : (notesMode ? "notes" : "group");
                WalkTextShapes(items, slide, prefix + name + "/", inner, notesMode, true, true, allowTable, slots, c);
                continue;
            }

            if (type == 19)  // table
            {
                if (!allowTable) continue;
                c.Tables++;
                int rows = 0, cols = 0;
                try { rows = (int)sh.Table.Rows.Count; cols = (int)sh.Table.Columns.Count; }
                catch (Exception) { continue; }
                for (int rr = 1; rr <= rows; rr++)
                {
                    for (int cc = 1; cc <= cols; cc++)
                    {
                        dynamic cell = null;
                        try { cell = sh.Table.Cell(rr, cc).Shape; } catch (Exception) { continue; }
                        c.Cells++;
                        AddShapeSlot(slots, slide, "table", prefix + name + "[" + rr + "," + cc + "]", name + "[" + rr + "," + cc + "]", cell, c);
                    }
                }
                continue;
            }

            if (!allowText) continue;
            if (IsHeaderFooterPlaceholder(sh)) continue;   // reached through the footer scope instead
            AddShapeSlot(slots, slide, where, prefix + name, name, sh, c);
        }
    }

    private static void AddShapeSlot(List<TextSlot> slots, int slide, string where, string path, string name, dynamic sh, Counts c)
    {
        bool hasText = false;
        try
        {
            if ((int)sh.HasTextFrame == 0) return;
            hasText = (int)sh.TextFrame.HasText != 0;
        }
        catch (Exception) { return; }
        if (!hasText) return;
        TextSlot slot = new TextSlot();
        slot.Slide = slide;
        slot.Where = where;
        slot.Path = path;
        slot.Name = name;
        slot.Frame = sh;
        slots.Add(slot);
        c.TextSlots++;
    }

    /// <summary>Date / header / footer / slide-number placeholders are owned by
    /// Slide.HeadersFooters; touching them twice would fight the footer scope.</summary>
    private static bool IsHeaderFooterPlaceholder(dynamic sh)
    {
        bool isPlaceholder = false;
        try { isPlaceholder = (int)sh.IsPlaceholder != 0; } catch (Exception) { return false; }
        if (!isPlaceholder) return false;
        int phType = 0;
        try { phType = Convert.ToInt32(TryGet(() => (object)sh.PlaceholderFormat.Type, 0), CultureInfo.InvariantCulture); } catch (Exception) { }
        return phType == 13 || phType == 14 || phType == 15 || phType == 16;
    }

    // ---------------------------------------------------------------- matching

    private static List<Span> FindSpans(string text, ReplaceRule rule, int room)
    {
        List<Span> spans = new List<Span>();
        if (room <= 0) return spans;

        if (rule.IsRegex)
        {
            foreach (Match m in rule.Compiled.Matches(text))
            {
                if (m.Length == 0) continue;
                Span sp = new Span();
                sp.Start0 = m.Index;
                sp.Len = m.Length;
                sp.Repl = m.Result(rule.Repl);   // supports $1
                spans.Add(sp);
                if (spans.Count >= room) break;
            }
            return spans;
        }

        StringComparison cmp = rule.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        int at = 0;
        while (at <= text.Length - rule.Find.Length)
        {
            int hit = text.IndexOf(rule.Find, at, cmp);
            if (hit < 0) break;
            Span sp = new Span();
            sp.Start0 = hit;
            sp.Len = rule.Find.Length;
            sp.Repl = rule.Repl;
            spans.Add(sp);
            at = hit + rule.Find.Length;
            if (spans.Count >= room) break;
        }
        return spans;
    }

    private static string SpanWindow(string text, Span sp)
    {
        int start = sp.Start0 - 18;
        if (start < 0) start = 0;
        int end = sp.Start0 + sp.Len + 24;
        if (end > text.Length) end = text.Length;
        return text.Substring(start, end - start).Replace("\r", " / ");
    }

    // ---------------------------------------------------------------- write

    private static void ApplySpans(TextSlot slot, ReplaceRule rule, string text, List<Span> spans, bool dryRun,
        List<object> hits, List<object> skipped, List<object> failed, HashSet<string> touched, ref int changed)
    {
        // A match that crosses a paragraph mark would merge two paragraphs; report it
        // instead of silently rejoining the user's layout.
        List<Span> usable = new List<Span>();
        foreach (Span sp in spans)
        {
            if (text.IndexOf('\r', sp.Start0, sp.Len) >= 0)
            {
                Dictionary<string, object> sk = SkipRecord(slot, rule, sp, "cross-paragraph", "the match contains a paragraph mark; replace it by hand or narrow -Find");
                skipped.Add(sk);
                continue;
            }
            usable.Add(sp);
        }
        if (usable.Count == 0) return;

        if (slot.Plain != null)
        {
            // header/footer/date: a plain string, so rebuild it in one go.
            StringBuilder built = new StringBuilder();
            int at = 0;
            foreach (Span sp in usable)
            {
                built.Append(text, at, sp.Start0 - at);
                built.Append(sp.Repl);
                at = sp.Start0 + sp.Len;
            }
            built.Append(text, at, text.Length - at);
            try
            {
                if (!dryRun) slot.Plain.Text = built.ToString();
            }
            catch (Exception e)
            {
                failed.Add(FailRecord(slot, rule, Clean(e.Message)));
                return;
            }
            foreach (Span sp in usable)
            {
                hits.Add(HitRecord(slot, rule, sp, text, dryRun));
                changed++;   // a string slot has no per-span identity to verify; the
                             // rebuild is computed from the same text we just read
            }
            touched.Add(slot.Slide + "|" + slot.Where + "|" + slot.Path);
            return;
        }

        // TextRange slots: back to front, because Characters indexes the live text.
        for (int i = usable.Count - 1; i >= 0; i--)
        {
            Span sp = usable[i];
            string matched = text.Substring(sp.Start0, sp.Len);
            bool verified = false;
            string detail = null;
            List<string> fontProblems = null;
            try
            {
                dynamic sub = slot.FreshRange().Characters(sp.Start0 + 1, sp.Len);
                string readBack = Convert.ToString(sub.Text, CultureInfo.InvariantCulture);
                verified = string.Equals(readBack, matched, StringComparison.Ordinal);
                if (!verified)
                    detail = "Characters(" + (sp.Start0 + 1) + "," + sp.Len + ") reads \"" + Clip(readBack, 40)
                           + "\" instead of the matched \"" + Clip(matched, 40) + "\"";
                else if (!dryRun)
                {
                    // WPS does NOT reliably keep the font of the replaced text when the
                    // span covers whole runs (measured: it takes the following run's
                    // font), so snapshot the first character's font and put it back.
                    FontSnap snap = SnapFont(slot, sp.Start0);
                    sub.Text = sp.Repl;
                    fontProblems = ApplyFontSnapshot(slot, sp.Start0, sp.Repl.Length, snap);
                }
            }
            catch (Exception e)
            {
                verified = false;
                detail = Clean(e.Message);
            }

            if (!verified)
            {
                skipped.Add(SkipRecord(slot, rule, sp, "unverified-span", detail));
                continue;
            }

            Dictionary<string, object> record = HitRecord(slot, rule, sp, text, dryRun);
            if (fontProblems != null && fontProblems.Count > 0)
            {
                record["fontNotKept"] = new List<object>(fontProblems.ToArray());
            }
            hits.Add(record);
            changed++;
            touched.Add(slot.Slide + "|" + slot.Where + "|" + slot.Path);
        }
    }

    // ---------------------------------------------------------------- font of a span

    /// <summary>Read the font of the character at start0 (the first character of the
    /// span, i.e. the one whose look the replacement should inherit).</summary>
    private static FontSnap SnapFont(TextSlot slot, int start0)
    {
        FontSnap snap = new FontSnap();
        dynamic font = null;
        try { font = slot.FreshRange().Characters(start0 + 1, 1).Font; }
        catch (Exception) { return snap; }

        object v = TryGet(() => (object)font.Size, null);
        if (v != null)
        {
            double size = Convert.ToDouble(v, CultureInfo.InvariantCulture);
            if (size > 0 && size < 4000) snap.Size = (float)size;   // 9999999 = mixed
        }
        v = TryGet(() => (object)font.Name, null);
        if (v != null) snap.Name = Convert.ToString(v, CultureInfo.InvariantCulture);
        v = TryGet(() => (object)font.NameFarEast, null);
        if (v != null) snap.NameFarEast = Convert.ToString(v, CultureInfo.InvariantCulture);
        v = TryGet(() => (object)font.Bold, null);
        if (v != null) snap.Bold = Convert.ToInt32(v, CultureInfo.InvariantCulture);
        v = TryGet(() => (object)font.Italic, null);
        if (v != null) snap.Italic = Convert.ToInt32(v, CultureInfo.InvariantCulture);
        v = TryGet(() => (object)font.Underline, null);
        if (v != null) snap.Underline = Convert.ToInt32(v, CultureInfo.InvariantCulture);
        v = TryGet(() => (object)font.Color.RGB, null);
        if (v != null) snap.Rgb = Convert.ToInt32(v, CultureInfo.InvariantCulture);
        return snap;
    }

    /// <summary>Make the replacement look like the text it replaced. Only properties
    /// that actually differ from what WPS produced are written, so a span that survived
    /// the edit untouched is left exactly as WPS wrote it (no pinned explicit values).
    /// Returns the names of the properties WPS refused, so the response can say so.</summary>
    private static List<string> ApplyFontSnapshot(TextSlot slot, int start0, int replLen, FontSnap snap)
    {
        List<string> problems = new List<string>();
        if (replLen <= 0 || snap == null) return problems;
        dynamic font = null;
        try { font = slot.FreshRange().Characters(start0 + 1, replLen).Font; }
        catch (Exception e) { problems.Add("range:" + Clean(e.Message)); return problems; }

        if (snap.Size > 0)
        {
            object raw = TryGet(() => (object)font.Size, null);
            double now = raw == null ? double.NaN : Convert.ToDouble(raw, CultureInfo.InvariantCulture);
            bool differs = !(now > 0 && now < 4000) || Math.Abs(now - snap.Size) > 0.01;
            if (differs)
            {
                bool ok = false;
                // VT_R8 is what every COM server accepts; some refuse VT_R4.
                try { font.Size = (double)snap.Size; ok = true; } catch (Exception) { }
                if (!ok) { try { font.Size = snap.Size; ok = true; } catch (Exception) { } }
                if (!ok) problems.Add("size");
            }
        }
        if (snap.Name != null)
        {
            string now = Convert.ToString(TryGet(() => (object)font.Name, null), CultureInfo.InvariantCulture);
            if (!string.Equals(now, snap.Name, StringComparison.Ordinal))
            {
                try { font.Name = snap.Name; } catch (Exception) { problems.Add("name"); }
            }
        }
        if (snap.NameFarEast != null)
        {
            string now = Convert.ToString(TryGet(() => (object)font.NameFarEast, null), CultureInfo.InvariantCulture);
            if (!string.Equals(now, snap.NameFarEast, StringComparison.Ordinal))
            {
                try { font.NameFarEast = snap.NameFarEast; } catch (Exception) { problems.Add("ea"); }
            }
        }
        if (snap.Bold.HasValue) WriteIfDifferent(font, "Bold", snap.Bold.Value, problems);
        if (snap.Italic.HasValue) WriteIfDifferent(font, "Italic", snap.Italic.Value, problems);
        if (snap.Underline.HasValue) WriteIfDifferent(font, "Underline", snap.Underline.Value, problems);
        if (snap.Rgb.HasValue)
        {
            object raw = TryGet(() => (object)font.Color.RGB, null);
            int now = raw == null ? -1 : Convert.ToInt32(raw, CultureInfo.InvariantCulture);
            if (now != snap.Rgb.Value)
            {
                try { font.Color.RGB = snap.Rgb.Value; } catch (Exception) { problems.Add("color"); }
            }
        }
        return problems;
    }

    private static void WriteIfDifferent(dynamic font, string property, int value, List<string> problems)
    {
        object raw = null;
        if (property == "Bold") raw = TryGet(() => (object)font.Bold, null);
        else if (property == "Italic") raw = TryGet(() => (object)font.Italic, null);
        else raw = TryGet(() => (object)font.Underline, null);
        if (raw != null && Convert.ToInt32(raw, CultureInfo.InvariantCulture) == value) return;
        try
        {
            if (property == "Bold") font.Bold = value;
            else if (property == "Italic") font.Italic = value;
            else font.Underline = value;
        }
        catch (Exception) { problems.Add(property.ToLowerInvariant()); }
    }

    private static Dictionary<string, object> HitRecord(TextSlot slot, ReplaceRule rule, Span sp, string text, bool dryRun)
    {
        Dictionary<string, object> h = new Dictionary<string, object>();
        h["slide"] = slot.Slide;
        h["where"] = slot.Where;
        h["shape"] = slot.Name;
        h["path"] = slot.Path;
        h["start"] = sp.Start0 + 1;
        h["before"] = Clip(text.Substring(sp.Start0, sp.Len), 80);
        h["after"] = Clip(sp.Repl, 80);
        h["context"] = Clean(SpanWindow(text, sp));
        h["rule"] = rule.Label;
        if (dryRun) h["would"] = true;
        return h;
    }

    private static Dictionary<string, object> SkipRecord(TextSlot slot, ReplaceRule rule, Span sp, string reason, string detail)
    {
        Dictionary<string, object> s = new Dictionary<string, object>();
        s["slide"] = slot.Slide;
        s["where"] = slot.Where;
        s["shape"] = slot.Name;
        s["path"] = slot.Path;
        s["reason"] = reason;
        s["detail"] = detail;
        s["rule"] = rule.Label;
        s["at"] = sp.Start0 + 1;
        return s;
    }

    private static Dictionary<string, object> FailRecord(TextSlot slot, ReplaceRule rule, string message)
    {
        Dictionary<string, object> f = new Dictionary<string, object>();
        f["slide"] = slot.Slide;
        f["where"] = slot.Where;
        f["shape"] = slot.Name;
        f["path"] = slot.Path;
        f["rule"] = rule.Label;
        f["message"] = message;
        return f;
    }

    // ---------------------------------------------------------------- snapshot / restore

    private static string SlotSnapshotPath(dynamic p)
    {
        string name = Convert.ToString(TryGet(() => (object)p.Name, "deck"), CultureInfo.InvariantCulture);
        if (name.Length == 0) name = "deck";
        string fileTag = name.Replace(".pptx", "").Replace(".ppt", "") + "-prereplace.json";
        return Path.Combine(SnapshotDir(), fileTag);
    }

    private static string WriteSlotSnapshot(dynamic p, List<TextSlot> slots)
    {
        List<object> cells = new List<object>();
        foreach (TextSlot slot in slots)
        {
            string t;
            try { t = slot.Text(); } catch (Exception) { continue; }
            cells.Add(new Dictionary<string, object> {
                { "slide", slot.Slide }, { "where", slot.Where }, { "path", slot.Path },
                { "name", slot.Name }, { "text", t }
            });
        }
        Dictionary<string, object> snap = new Dictionary<string, object>();
        snap["at"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        snap["kind"] = "text-slots";
        snap["doc"] = Convert.ToString(TryGet(() => (object)p.Name, "?"), CultureInfo.InvariantCulture);
        snap["cells"] = cells;
        string path = SlotSnapshotPath(p);
        try { File.WriteAllText(path, Json.Write(snap), new UTF8Encoding(false)); }
        catch (Exception) { return null; }
        return path;
    }

    /// <summary>Called from RestoreCmd when the snapshot on disk is a text-slot one
    /// (the shape written by `replace`), which covers groups, table cells, notes and
    /// header/footer as well - the plain snapshot only carries top-level shape text.</summary>
    private static object RestoreSlots(dynamic p, Bag snap, string only, string snapPath)
    {
        HashSet<string> wanted = null;
        if (only != null)
        {
            wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string raw in only.Split(',')) if (raw.Trim().Length > 0) wanted.Add(raw.Trim());
        }

        Counts counts = new Counts();
        List<TextSlot> slots = CollectTextSlots(p, ParseReplaceScopes("all"), null, counts);
        Dictionary<string, TextSlot> index = new Dictionary<string, TextSlot>(StringComparer.Ordinal);
        foreach (TextSlot slot in slots)
        {
            string key = slot.Slide + "|" + slot.Where + "|" + slot.Path;
            if (!index.ContainsKey(key)) index[key] = slot;
        }

        int restored = 0, missing = 0, failed = 0;
        List<object> details = new List<object>();
        foreach (object o in snap.List("cells"))
        {
            Bag cell = new Bag(o);
            int slide = cell.Int("slide", 0);
            string where = cell.Str("where", "body");
            string path = cell.Str("path", "");
            string want = cell.Str("text", "");
            string name = cell.Str("name", "");
            if (wanted != null && !wanted.Contains(name) && !wanted.Contains(path)
                && !wanted.Contains(slide.ToString(CultureInfo.InvariantCulture) + "!" + name)) continue;

            TextSlot slot;
            if (!index.TryGetValue(slide + "|" + where + "|" + path, out slot)) { missing++; continue; }
            string now;
            try { now = slot.Text(); } catch (Exception) { failed++; continue; }
            if (string.Equals(now, want, StringComparison.Ordinal)) continue;
            try
            {
                slot.SetWhole(want);
                restored++;
                details.Add(new Dictionary<string, object> {
                    { "slide", slide }, { "where", where }, { "shape", name },
                    { "now", Clip(now, 50) }, { "restored", Clip(want, 50) }
                });
            }
            catch (Exception) { failed++; }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["since"] = snapPath;
        r["restored"] = restored;
        r["failed"] = failed;
        if (missing > 0) r["missing"] = missing;
        if (details.Count > 0) r["details"] = details;
        r["note"] = "text only (whole paragraph rewritten, so run-level formatting inside a restored paragraph is lost); "
                  + "shapes added or removed after the snapshot are not reverted; run diff, then render before saving";
        return r;
    }
}
