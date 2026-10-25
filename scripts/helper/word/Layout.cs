// P1 layout: char/paragraph formatting on a range, style-level formatting, page setup.
// Part of the wps-word-cooperate helper (C# 5, compiled together with WpsHelper.cs).

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

internal static partial class Program
{
    /// <summary>format: apply char + paragraph formatting to one paragraph, a range, one style's paragraphs, or the whole document.</summary>
    private static object FormatDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; no formatting was applied.");

        bool track = a.Flag("track", false);
        bool wasTracking = false;
        try { wasTracking = (bool)doc.TrackRevisions; } catch (Exception) { }

        dynamic app = App(true, true);
        dynamic undo = null;
        try { undo = app.UndoRecord; undo.StartCustomRecord("wps-word-cooperate format"); } catch (Exception) { undo = null; }
        if (!track) { try { doc.TrackRevisions = false; } catch (Exception) { } }

        List<string> applied = new List<string>();
        List<object> matchSamples = null;
        string target;
        int paragraphs = 0;
        try
        {
            if (a.Has("para"))
            {
                int n = a.Int("para", 0);
                int total = (int)doc.Paragraphs.Count;
                if (n < 1 || n > total) throw new Fail("BAD_ARGS", "paragraph " + n + " does not exist (the document has " + total + ").");
                dynamic p = doc.Paragraphs.Item(n);
                string expect = a.Str("expect", "");
                if (expect.Length > 0)
                {
                    string text = CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));
                    if (!text.StartsWith(expect, StringComparison.Ordinal))
                        throw new Fail("ANCHOR_MISSING", "paragraph " + n + " starts with \"" + Head(text, 40) + "\", not \"" + Head(expect, 40) + "\".");
                }
                applied.AddRange(ApplyCharFormat(p.Range.Font, a));
                applied.AddRange(ApplyParaFormat(p.Format, a));
                target = "paragraph " + n;
                paragraphs = 1;
            }
            else if (a.Has("from") || a.Has("to"))
            {
                int total = (int)doc.Paragraphs.Count;
                int from = Math.Max(1, a.Int("from", 1));
                int to = Math.Min(total, a.Int("to", from));
                if (from > to) throw new Fail("BAD_ARGS", "\"from\" is after \"to\"");
                dynamic start = doc.Paragraphs.Item(from).Range.Start;
                dynamic end = doc.Paragraphs.Item(to).Range.End;
                dynamic range = doc.Range((int)start, (int)end);
                applied.AddRange(ApplyCharFormat(range.Font, a));
                applied.AddRange(ApplyParaFormat(range.ParagraphFormat, a));
                target = "paragraphs " + from + "-" + to;
                paragraphs = to - from + 1;
            }
            else if (a.Has("match"))
            {
                // Every paragraph whose TEXT matches a .NET regex - the caption/case-law locator
                // ("^\\s*(图|附图)\\s*\\d" hits all figure captions). One undo record for the lot.
                string pattern = a.Str("match", "");
                if (pattern.Trim().Length == 0) throw new Fail("BAD_ARGS", "\"match\" is empty");
                System.Text.RegularExpressions.Regex rx;
                try { rx = new System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.None); }
                catch (Exception e) { throw new Fail("BAD_ARGS", "bad match regex \"" + pattern + "\": " + Clean(e.Message)); }
                int total = (int)doc.Paragraphs.Count;
                int hits = 0;
                List<object> samples = new List<object>();

                // Fast path: scan the whole text in '\r' chunks. ONLY when the document has no
                // unaccepted revisions - deleted revision text counts into Range coordinates but
                // is hidden in Content.Text, so offsets drift the moment a deletion exists.
                int revisions0 = (int)TryGet(() => (object)(int)doc.Revisions.Count, 1);
                string all = revisions0 == 0 ? WholeText(doc) : null;
                bool fastOk = all != null;
                if (fastOk)
                {
                    int pos = 0;
                    int from = 0;
                    while (from <= all.Length)
                    {
                        int cr = all.IndexOf('\r', from);
                        string chunk = cr < 0 ? all.Substring(from) : all.Substring(from, cr - from);
                        if (rx.IsMatch(chunk))
                        {
                            int[] cand = ParaCandidatesAt(doc, pos);
                            if (cand == null) { fastOk = false; break; }
                            dynamic p = null;
                            string ptext = null;
                            int paraNo = 0;
                            foreach (int pc in cand)
                            {
                                string t = null;
                                try { t = CleanPara(Convert.ToString(doc.Paragraphs.Item(pc).Range.Text, CultureInfo.InvariantCulture)); }
                                catch (Exception) { }
                                if (t != null && rx.IsMatch(t)) { p = doc.Paragraphs.Item(pc); ptext = t; paraNo = pc; break; }
                            }
                            if (p == null) { fastOk = false; break; }   // drift: abandon fast path
                            ApplyCharFormat(p.Range.Font, a);
                            ApplyParaFormat(p.Format, a);
                            hits++;
                            if (samples.Count < 5) samples.Add(paraNo + ": " + Head(ptext, 40));
                            if (hits >= 5000) break;
                        }
                        if (cr < 0) break;
                        pos += chunk.Length + 1;
                        from = cr + 1;
                    }
                }
                if (!fastOk)
                {
                    hits = 0; samples = new List<object>();
                    for (int i = 1; i <= total; i++)
                    {
                        dynamic p = null;
                        try { p = doc.Paragraphs.Item(i); } catch (Exception) { continue; }
                        string text = CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture));
                        if (text.Length == 0 || !rx.IsMatch(text)) continue;
                        ApplyCharFormat(p.Range.Font, a);
                        ApplyParaFormat(p.Format, a);
                        hits++;
                        if (samples.Count < 5) samples.Add(i + ": " + Head(text, 40));
                        if (hits >= 5000) break;
                    }
                }
                if (hits == 0) throw new Fail("BAD_ARGS", "no paragraph text matches \"" + pattern + "\".");
                applied.Add("match=" + pattern);
                target = "match \"" + pattern + "\"";
                paragraphs = hits;
                matchSamples = samples;
            }
            else if (a.Has("style"))
            {
                string want = a.Str("style", "");
                int total = (int)doc.Paragraphs.Count;
                int hits = 0;
                for (int i = 1; i <= total; i++)
                {
                    dynamic p = doc.Paragraphs.Item(i);
                    string has = Convert.ToString(TryGet(() => (object)p.Style.NameLocal, ""), CultureInfo.InvariantCulture);
                    if (!string.Equals(has, want, StringComparison.OrdinalIgnoreCase)) continue;
                    List<string> gotChars = ApplyCharFormat(p.Range.Font, a);
                    List<string> gotPara = ApplyParaFormat(p.Format, a);
                    if (applied.Count == 0) { applied.AddRange(gotChars); applied.AddRange(gotPara); }
                    hits++;
                    if (hits >= 5000) break;
                }
                if (hits == 0) throw new Fail("BAD_ARGS", "no paragraph uses the style \"" + want + "\"; check the name with read (it is the localised name, e.g. \u6807\u9898 1).");
                target = "style " + want;
                paragraphs = hits;
                if (applied.Count == 0) applied.Add("(no formatting field given)");
            }
            else
            {
                dynamic content = doc.Content;
                applied.AddRange(ApplyCharFormat(content.Font, a));
                applied.AddRange(ApplyParaFormat(content.ParagraphFormat, a));
                target = "whole document";
                paragraphs = (int)doc.Paragraphs.Count;
            }
        }
        finally
        {
            if (!track) { try { doc.TrackRevisions = wasTracking; } catch (Exception) { } }
            if (undo != null) { try { undo.EndCustomRecord(); } catch (Exception) { } }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["target"] = target;
        r["paragraphs"] = paragraphs;
        r["applied"] = applied;
        if (matchSamples != null) r["samples"] = matchSamples;
        r["readBack"] = FormatReadback(doc, a.Has("para") ? a.Int("para", 1) : (a.Has("from") ? a.Int("from", 1) : 1));
        r["tracked"] = track;
        r["saved"] = false;
        r["note"] = "Nothing was saved. Undo once (Ctrl+Z) takes the whole change back.";
        return r;
    }

    /// <summary>style: change the style itself, so every paragraph using it follows.</summary>
    private static object StyleDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; the style was not changed.");
        string want = a.Need("style");

        dynamic style = FindStyle(doc, want);
        List<string> applied = new List<string>();
        applied.AddRange(ApplyCharFormat(style.Font, a));
        applied.AddRange(ApplyParaFormat(style.ParagraphFormat, a));

        int stripped = 0;
        if (a.Has("numbering") && !a.Flag("numbering", true))
        {
            int total = (int)doc.Paragraphs.Count;
            for (int i = 1; i <= total; i++)
            {
                dynamic p = doc.Paragraphs.Item(i);
                string has = Convert.ToString(TryGet(() => (object)p.Style.NameLocal, ""), CultureInfo.InvariantCulture);
                if (!string.Equals(has, want, StringComparison.OrdinalIgnoreCase)) continue;
                try { p.Range.ListFormat.RemoveNumbers(); stripped++; } catch (Exception) { }
            }
            applied.Add("numbering removed on " + stripped + " paragraph(s)");
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["style"] = Convert.ToString(TryGet(() => (object)style.NameLocal, ""), CultureInfo.InvariantCulture);
        r["applied"] = applied;
        r["font"] = Convert.ToString(TryGet(() => (object)style.Font.NameFarEast, ""), CultureInfo.InvariantCulture);
        r["size"] = TryGet(() => (object)(double)style.Font.Size, 0);
        r["saved"] = false;
        r["note"] = "Style-level changes are not revisions and cannot be taken back with Ctrl+Z; nothing was saved yet.";
        return r;
    }

    private static dynamic FindStyle(dynamic doc, string name)
    {
        string[] aliases = new string[] {
            "Normal", "\u6b63\u6587", "Heading 1", "\u6807\u9898 1", "Heading 2", "\u6807\u9898 2", "Heading 3", "\u6807\u9898 3",
        };
        string[] tries;
        int index = Array.FindIndex(aliases, x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase));
        if (index >= 0 && index % 2 == 0) tries = new string[] { aliases[index + 1], name };
        else if (index >= 0) tries = new string[] { name, aliases[index - 1] };
        else tries = new string[] { name };

        foreach (string t in tries)
        {
            try { return doc.Styles.Item(t); } catch (Exception) { }
        }
        // Built-in ids as a last resort: Normal -1, Heading 1..3 -2..-4.
        int builtin = -1;
        if (string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase)) builtin = -1;
        else if (string.Equals(name, "Heading 1", StringComparison.OrdinalIgnoreCase)) builtin = -2;
        else if (string.Equals(name, "Heading 2", StringComparison.OrdinalIgnoreCase)) builtin = -3;
        else if (string.Equals(name, "Heading 3", StringComparison.OrdinalIgnoreCase)) builtin = -4;
        if (builtin != -1)
        {
            try { return doc.Styles.Item(builtin); } catch (Exception) { }
        }
        throw new Fail("BAD_ARGS", "no style named \"" + name + "\". Style names are localised (e.g. \u6b63\u6587, \u6807\u9898 1); read returns the exact name of every paragraph's style.");
    }

    /// <summary>page: paper, orientation, margins, header/footer text, page numbers.</summary>
    private static object PageDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);

        // Query mode: no page option given -> report every section's setup ("only change the 2nd section's margins"
        // needs to know how many sections there are and what they look like). Read-only, changes nothing.
        bool anyOption = a.Has("paper") || a.Has("orientation") || a.Has("top") || a.Has("bottom") || a.Has("left")
            || a.Has("right") || a.Has("header") || a.Has("footer") || a.Has("pageNumber") || a.Has("pageNumberStart");
        if (!anyOption)
        {
            List<object> secs = new List<object>();
            int nsec = (int)TryGet(() => (object)(int)doc.Sections.Count, 0);
            for (int s = 1; s <= nsec; s++)
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                try
                {
                    dynamic ps = doc.Sections.Item(s).PageSetup;
                    one["section"] = s;
                    one["paperSizeId"] = TryGet(() => (object)(int)ps.PaperSize, 0);   // 7=A4, 6=A3 ...
                    one["orientation"] = (int)TryGet(() => (object)(int)ps.Orientation, 0) == 1 ? "landscape" : "portrait";
                    one["marginsCm"] = new Dictionary<string, object> {
                        {"top", Math.Round((double)(float)ps.TopMargin / 28.3465, 2)},
                        {"bottom", Math.Round((double)(float)ps.BottomMargin / 28.3465, 2)},
                        {"left", Math.Round((double)(float)ps.LeftMargin / 28.3465, 2)},
                        {"right", Math.Round((double)(float)ps.RightMargin / 28.3465, 2)},
                    };
                    one["headerText"] = Head(CleanPara(Convert.ToString(TryGet(() => (object)doc.Sections.Item(s).Headers.Item(1).Range.Text, ""), CultureInfo.InvariantCulture)), 40);
                    one["footerText"] = Head(CleanPara(Convert.ToString(TryGet(() => (object)doc.Sections.Item(s).Footers.Item(1).Range.Text, ""), CultureInfo.InvariantCulture)), 40);
                }
                catch (Exception e) { one["section"] = s; one["error"] = Clean(e.Message); }
                secs.Add(one);
            }
            Dictionary<string, object> qr = new Dictionary<string, object>();
            qr["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
            qr["sections"] = secs;
            qr["saved"] = true;
            qr["note"] = "\u67e5\u8be2\u6a21\u5f0f\uff1a\u6ca1\u4f20\u4efb\u4f55\u9875\u9762\u9009\u9879\u65f6\u53ea\u62a5\u5404\u8282\u8bbe\u7f6e\u3002\u8981\u6539\u5168\u90e8\u8282\u7528 -All\uff1b\u53ea\u6539\u7b2c\u4e00\u8282\u4e0d\u52a0 -All\uff08\u9ed8\u8ba4\u8865\u4e0a -Top \u7b49\u53c2\u6570\u624d\u4f1a\u5199\uff09\u3002";
            return qr;
        }

        if (IsReadOnly(doc)) throw new Fail("READONLY", "The document is open read-only; page setup was not changed.");

        Dictionary<string, int> papers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            {"A4", 7}, {"A3", 6}, {"A5", 11}, {"B5", 13}, {"Letter", 2}, {"Legal", 4},
        };

        int sections = (int)doc.Sections.Count;
        List<object> done = new List<object>();
        List<string> applied = new List<string>();

        for (int s = 1; s <= sections; s++)
        {
            dynamic section = doc.Sections.Item(s);
            dynamic ps = section.PageSetup;
            List<string> step = new List<string>();

            if (a.Has("paper"))
            {
                string p = a.Str("paper", "A4");
                int id;
                if (!papers.TryGetValue(p, out id)) throw new Fail("BAD_ARGS", "unknown paper \"" + p + "\" (A4|A3|A5|B5|Letter|Legal)");
                ps.PaperSize = id;
                step.Add("paper=" + p);
            }
            if (a.Has("orientation"))
            {
                string o = a.Str("orientation", "portrait").ToLowerInvariant();
                ps.Orientation = (o.StartsWith("land") || o == "\u6a2a\u5411") ? 1 : 0;
                step.Add("orientation=" + o);
            }
            if (a.Has("top")) { ps.TopMargin = (float)Points(a.Str("top", "2.54cm")); step.Add("top=" + a.Str("top", "")); }
            if (a.Has("bottom")) { ps.BottomMargin = (float)Points(a.Str("bottom", "2.54cm")); step.Add("bottom=" + a.Str("bottom", "")); }
            if (a.Has("left")) { ps.LeftMargin = (float)Points(a.Str("left", "3.17cm")); step.Add("left=" + a.Str("left", "")); }
            if (a.Has("right")) { ps.RightMargin = (float)Points(a.Str("right", "2.54cm")); step.Add("right=" + a.Str("right", "")); }

            bool everySection = a.Flag("all", true);
            if (s == 1 || everySection)
            {
                if (a.Has("header"))
                {
                    dynamic hr = section.Headers.Item(1).Range;
                    hr.Text = a.Str("header", "");
                    step.Add("header");
                }
                if (a.Has("footer"))
                {
                    dynamic fr = section.Footers.Item(1).Range;
                    fr.Text = a.Str("footer", "");
                    step.Add("footer");
                }
                if (a.Has("pageNumber") || a.Has("pageNumberStart"))
                {
                    step.Add(PlacePageNumber(section, a));
                }
            }

            Dictionary<string, object> one = new Dictionary<string, object>();
            one["section"] = s;
            one["applied"] = step;
            one["paperSize"] = TryGet(() => (object)(int)ps.PaperSize, 0);
            one["orientation"] = TryGet(() => (object)(int)ps.Orientation, 0);
            one["margins"] = new Dictionary<string, object> {
                {"top", TryGet(() => (object)Math.Round((double)ps.TopMargin, 1), 0)},
                {"bottom", TryGet(() => (object)Math.Round((double)ps.BottomMargin, 1), 0)},
                {"left", TryGet(() => (object)Math.Round((double)ps.LeftMargin, 1), 0)},
                {"right", TryGet(() => (object)Math.Round((double)ps.RightMargin, 1), 0)},
            };
            done.Add(one);
            applied.AddRange(step);
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["sections"] = done;
        r["applied"] = applied;
        r["saved"] = false;
        return r;
    }

    /// <summary>Puts (or removes) a PAGE field in the primary footer and aligns it.</summary>
    private static string PlacePageNumber(dynamic section, Bag a)
    {
        string where = a.Str("pageNumber", "");
        dynamic footer = section.Footers.Item(1);
        dynamic fr = footer.Range;
        try
        {
            for (int i = (int)fr.Fields.Count; i >= 1; i--)
            {
                dynamic f = fr.Fields.Item(i);
                if ((int)f.Type == 33) f.Delete();   // wdFieldPage
            }
        }
        catch (Exception) { }

        if (where.Length > 0 && !where.Equals("none", StringComparison.OrdinalIgnoreCase) && !where.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (CleanPara(Convert.ToString(TryGet(() => (object)fr.Text, ""), CultureInfo.InvariantCulture)).Trim().Length == 0)
                fr.Text = "";
            dynamic at = section.Footers.Item(1).Range;
            at.Collapse(0);                      // wdCollapseEnd
            at.Fields.Add(at, 33, null, false);  // wdFieldPage
            int align = Alignment(where.Equals("false", StringComparison.OrdinalIgnoreCase) ? "center" : where);
            if (align >= 0) { dynamic pf = footer.Range.ParagraphFormat; pf.Alignment = align; }
        }

        if (a.Has("pageNumberStart"))
        {
            int n = a.Int("pageNumberStart", 1);
            try { section.Headers.Item(1).PageNumbers.StartingNumber = n; }
            catch (Exception e) { return "pageNumber " + where + " (starting number not supported: " + Clean(e.Message) + ")"; }
        }
        return "pageNumber=" + (where.Length == 0 ? "kept" : where);
    }

    /// <summary>What the first paragraphs of the document actually look like now.</summary>
    private static Dictionary<string, object> FormatReadback(dynamic doc, int para)
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        try
        {
            int total = (int)doc.Paragraphs.Count; if (para < 1) para = 1; if (para > total) para = total;
            dynamic p = doc.Paragraphs.Item(para);
            dynamic f = p.Range.Font;
            d["paragraph"] = para;
            d["text"] = Head(CleanPara(Convert.ToString(TryGet(() => (object)p.Range.Text, ""), CultureInfo.InvariantCulture)), 30);
            d["font"] = Convert.ToString(TryGet(() => (object)f.NameFarEast, ""), CultureInfo.InvariantCulture);
            d["size"] = TryGet(() => (object)Math.Round((double)f.Size, 1), 0);
            dynamic pf = p.Format;
            d["align"] = TryGet(() => (object)(int)pf.Alignment, 0);
            d["lineSpacingRule"] = TryGet(() => (object)(int)pf.LineSpacingRule, 0);
            d["lineSpacing"] = TryGet(() => (object)Math.Round((double)pf.LineSpacing, 1), 0);
            d["firstLineIndent"] = TryGet(() => (object)Math.Round((double)pf.FirstLineIndent, 1), 0);
        }
        catch (Exception) { }
        return d;
    }
}
