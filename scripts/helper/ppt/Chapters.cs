// wps-office-cooperate (ppt): `split` / `merge` - 分册拆分与多稿合并.
//
// Why this file exists: a 80-page 汇报 routinely has to become 5 分册, and three
// 县 decks routinely have to become one.  Until now the skill could only duplicate,
// move or delete a single slide, so that work was manual.
//
// Two measured facts decide how the code is written (WPS 12.1.0.29161):
//   * Presentation.Slides.InsertFromFile is NOT usable: it empties the target deck.
//   * Copy/Paste does NOT carry the source's design.  A slide copied from a deck whose
//     master carried a marker text arrives in the target with the TARGET's master and
//     layout.  So `split` must not build parts by pasting - it copies the whole file
//     (SaveCopyAs) and deletes the slides outside the range, which keeps the design,
//     the layouts and the notes master exactly as they were.
//   * `merge` is the opposite direction and that is fine: pasted pages taking the
//     target's look is what "merge three county decks into one" usually wants, and
//     `theme -Apply` / `master` can unify the rest.
//
// Pure ASCII source (Chinese only ever travels inside JSON strings), C# 5.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

internal static partial class Program
{
    /// <summary>Ranges like "1-20,21-45,46-80" - one entry per output deck.</summary>
    private sealed class Chunk
    {
        public int From;
        public int To;
        public string Label()
        {
            return From == To
                ? From.ToString(CultureInfo.InvariantCulture)
                : From.ToString(CultureInfo.InvariantCulture) + "-" + To.ToString(CultureInfo.InvariantCulture);
        }
        public int Count() { return To - From + 1; }
    }

    private static List<Chunk> ParseChunks(string spec, int slideCount)
    {
        if (spec == null || spec.Trim().Length == 0)
            throw new Fail("BAD_ARGS", "split needs -Chunks \"1-20,21-45,46-80\" (one range per output deck)");
        List<Chunk> chunks = new List<Chunk>();
        int expectFrom = 1;
        foreach (string raw in spec.Split(','))
        {
            string v = raw.Trim();
            if (v.Length == 0) continue;
            int dash = v.IndexOf('-');
            int lo, hi;
            if (dash > 0)
            {
                if (!int.TryParse(v.Substring(0, dash).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out lo)
                    || !int.TryParse(v.Substring(dash + 1).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out hi))
                    throw new Fail("BAD_ARGS", "bad chunk \"" + v + "\"; use e.g. 1-20,21-45");
            }
            else
            {
                if (!int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out lo))
                    throw new Fail("BAD_ARGS", "bad chunk \"" + v + "\"");
                hi = lo;
            }
            if (hi < lo) { int t = lo; lo = hi; hi = t; }
            if (lo < 1 || hi > slideCount)
                throw new Fail("BAD_ARGS", "chunk \"" + v + "\" is outside 1.." + slideCount);
            if (lo != expectFrom)
                throw new Fail("BAD_ARGS", "chunks must be contiguous and start at 1: expected from " + expectFrom
                    + " but \"" + v + "\" starts at " + lo + " (a gap means slides would be lost silently)");
            Chunk chunk = new Chunk();
            chunk.From = lo;
            chunk.To = hi;
            chunks.Add(chunk);
            expectFrom = hi + 1;
        }
        if (chunks.Count == 0) throw new Fail("BAD_ARGS", "no chunk parsed from -Chunks " + spec);
        if (expectFrom != slideCount + 1)
            throw new Fail("BAD_ARGS", "chunks cover 1.." + (expectFrom - 1) + " but the deck has " + slideCount
                + " slides; add the missing range instead of dropping pages");
        return chunks;
    }

    // ---------------------------------------------------------------- split

    private static object SplitCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }
        if (n == 0) throw new Fail("BAD_ARGS", "this presentation has no slides");

        List<Chunk> chunks = ParseChunks(a.Str("chunks", null), n);
        string outArg = a.Str("out", null);
        if (outArg == null || outArg.Length == 0) throw new Fail("BAD_ARGS", "split needs -Out <folder>");
        string folder = FullPathOf(outArg);
        Directory.CreateDirectory(folder);

        string baseName = a.Str("prefix", null);
        if (baseName == null || baseName.Length == 0)
        {
            baseName = Convert.ToString(TryGet(() => (object)p.Name, "deck"), CultureInfo.InvariantCulture);
            int dot = baseName.LastIndexOf('.');
            if (dot > 0) baseName = baseName.Substring(0, dot);
        }
        string safeBase = AsciiFileName(baseName);

        bool dryRun = a.Flag("dryRun", false);
        dynamic app = App(true, false);
        string scratch = Path.Combine(Path.GetTempPath(), "woc-ppt");
        Directory.CreateDirectory(scratch);
        int total = 0;
        try { total = (int)p.Slides.Count; } catch (Exception) { }

        List<object> files = new List<object>();
        List<object> failed = new List<object>();
        for (int c = 0; c < chunks.Count; c++)
        {
            Chunk chunk = chunks[c];
            string target = Path.Combine(folder, safeBase + "-part" + (c + 1).ToString("00", CultureInfo.InvariantCulture) + ".pptx");
            if (dryRun)
            {
                files.Add(new Dictionary<string, object> {
                    { "out", target }, { "chunk", chunk.Label() }, { "slides", chunk.Count() }, { "would", true } });
                continue;
            }
            string copyPath = Path.Combine(scratch, "split-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".pptx");
            dynamic copy = null;
            try
            {
                // Copy the whole file first: a copy keeps the design, layouts and notes
                // master (pasting into a fresh deck would silently drop all three).
                p.SaveCopyAs(copyPath, 24);
                copy = app.Presentations.Open(copyPath, 0, 0, 0);
                int count = 0;
                try { count = (int)copy.Slides.Count; } catch (Exception) { }
                if (count != total)
                    throw new Fail("COM_ERROR", "the working copy has " + count + " slides but the source has " + total);

                for (int k = count; k > chunk.To; k--) copy.Slides.Item(k).Delete();          // drop the tail
                for (int k = 1; k < chunk.From; k++) copy.Slides.Item(1).Delete();            // drop the head
                int kept = 0;
                try { kept = (int)copy.Slides.Count; } catch (Exception) { }
                if (kept != chunk.Count())
                    throw new Fail("COM_ERROR", "chunk " + chunk.Label() + " left " + kept
                        + " slides, expected " + chunk.Count());

                copy.SaveAs(target, 24);        // ppSaveAsOpenXMLPresentation
                files.Add(new Dictionary<string, object> {
                    { "out", target }, { "chunk", chunk.Label() }, { "slides", kept },
                    { "first", SlideTitle(copy, 1) }, { "last", SlideTitle(copy, kept) }
                });
            }
            catch (Exception e)
            {
                failed.Add(new Dictionary<string, object> {
                    { "chunk", chunk.Label() }, { "reason", Clean(e.Message) } });
            }
            finally
            {
                if (copy != null) { try { copy.Close(); } catch (Exception) { } }
                TryDelete(copyPath);
            }
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["source"] = Convert.ToString(TryGet(() => (object)p.FullName, ""), CultureInfo.InvariantCulture);
        r["slides"] = n;
        r["chunks"] = chunks.Count;
        r["out"] = folder;
        r["files"] = files;
        if (failed.Count > 0) r["failed"] = failed;
        r["dryRun"] = dryRun;
        if (!dryRun) r["note"] = "each part is a copy of the source with the other pages deleted, so the design, layouts and "
                               + "notes master come along. Files written by WPS are transparently encrypted on this machine - "
                               + "decrypt before sending.";
        if (dryRun) r["note"] = "dry run: nothing was written";
        return r;
    }

    /// <summary>Append to the end, retrying once: Paste goes through the clipboard and
    /// a busy clipboard is the one failure that is worth simply repeating.</summary>
    private static void PasteAppend(dynamic part)
    {
        try { part.Slides.Paste(); return; }
        catch (Exception)
        {
            Thread.Sleep(350);
            part.Slides.Paste();
        }
    }

    private static string SlideTitle(dynamic pres, int index)
    {
        try
        {
            dynamic slide = pres.Slides.Item(index);
            return Clip(FirstTextOf(slide), 40);
        }
        catch (Exception) { return ""; }
    }

    private static string FirstTextOf(dynamic slide)
    {
        int count = 0;
        try { count = (int)slide.Shapes.Count; } catch (Exception) { return ""; }
        for (int i = 1; i <= count; i++)
        {
            dynamic sh = null;
            try { sh = slide.Shapes.Item(i); } catch (Exception) { continue; }
            string text = ShapeText(sh);
            if (text != null && text.Trim().Length > 0) return Clean(text);
        }
        return "";
    }

    /// <summary>Keep the file name ASCII: AddPicture-style APIs are not the only thing that
    /// dislikes the Chinese in this work folder, and a part file is easier to pass on.</summary>
    private static string AsciiFileName(string name)
    {
        if (name == null) return "deck";
        StringBuilder sb = new StringBuilder();
        foreach (char ch in name)
        {
            if (ch < 128 && (char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' || ch == '.')) sb.Append(ch);
            else if (ch == ' ') sb.Append('_');
        }
        if (sb.Length == 0) sb.Append("deck");
        if (sb.Length > 60) sb.Length = 60;
        return sb.ToString();
    }

    // ---------------------------------------------------------------- merge

    private static object MergeCmd(Bag a)
    {
        dynamic target = RequirePres(a, true);
        int targetCount = 0;
        try { targetCount = (int)target.Slides.Count; } catch (Exception) { }

        List<string> sources = SourceList(a);
        if (sources.Count == 0) throw new Fail("BAD_ARGS", "merge needs -Sources \"a.pptx,b.pptx\" or -Items <json with sources>");
        if (sources.Count > 60) throw new Fail("BAD_ARGS", "too many sources (" + sources.Count + "); merge in batches");

        int after = a.Int("after", targetCount);
        if (after < 0) after = 0;
        if (after > targetCount) throw new Fail("BAD_ARGS", "-After " + after + " is beyond the target deck (" + targetCount + " slides)");
        bool dryRun = a.Flag("dryRun", false);

        dynamic app = App(true, false);
        List<object> appended = new List<object>();
        List<object> failed = new List<object>();
        int insertAt = after;
        foreach (string raw in sources)
        {
            string path = FullPathOf(raw);
            if (!File.Exists(path))
            {
                failed.Add(new Dictionary<string, object> { { "source", path }, { "reason", "file not found" } });
                continue;
            }
            string self = Convert.ToString(TryGet(() => (object)target.FullName, ""), CultureInfo.InvariantCulture);
            if (string.Equals(path, self, StringComparison.OrdinalIgnoreCase))
            {
                failed.Add(new Dictionary<string, object> { { "source", path }, { "reason", "that is the target deck itself" } });
                continue;
            }
            dynamic src = null;
            try
            {
                src = app.Presentations.Open(path, -1, 0, 0);   // read-only, no window
                int count = 0;
                try { count = (int)src.Slides.Count; } catch (Exception) { }
                if (dryRun)
                {
                    appended.Add(new Dictionary<string, object> {
                        { "source", path }, { "slides", count }, { "would", true }, { "at", insertAt + 1 } });
                    insertAt += count;
                    continue;
                }
                for (int k = 1; k <= count; k++)
                {
                    src.Slides.Item(k).Copy();
                    Thread.Sleep(80);
                    try { target.Slides.Paste(insertAt + 1); }
                    catch (Exception)
                    {
                        Thread.Sleep(350);
                        target.Slides.Paste(insertAt + 1);
                    }
                    insertAt++;
                }
                appended.Add(new Dictionary<string, object> {
                    { "source", path }, { "slides", count }, { "from", insertAt - count + 1 }, { "to", insertAt } });
            }
            catch (Exception e)
            {
                failed.Add(new Dictionary<string, object> { { "source", path }, { "reason", Clean(e.Message) } });
            }
            finally
            {
                if (src != null) { try { src.Close(); } catch (Exception) { } }
            }
        }

        int now = 0;
        try { now = (int)target.Slides.Count; } catch (Exception) { }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["target"] = Convert.ToString(TryGet(() => (object)target.FullName, ""), CultureInfo.InvariantCulture);
        r["slidesBefore"] = targetCount;
        r["slidesNow"] = now;
        r["appended"] = appended;
        if (failed.Count > 0) r["failed"] = failed;
        r["dryRun"] = dryRun;
        if (!dryRun)
            r["note"] = "the pasted pages keep the design they came from; run `theme -Apply` if the whole deck has to look alike, "
                      + "and `check` before delivering. Undo by deleting the appended range, then save.";
        return r;
    }

    /// <summary>Sources as a comma list on the command line, as a JSON array in
    /// {"sources":[...]} (the safe route when the paths are awkward), or as
    /// {"items":[{"source":"..."}]}.</summary>
    private static List<string> SourceList(Bag a)
    {
        List<string> sources = new List<string>();
        string spec = a.Str("sources", null);
        if (spec != null && spec.Trim().Length > 0)
        {
            foreach (string raw in spec.Split(','))
            {
                string v = raw.Trim().Trim('"');
                if (v.Length > 0) sources.Add(v);
            }
        }
        foreach (object o in a.List("sourcesList"))
        {
            string one = o as string;
            if (one == null)
            {
                Bag item = new Bag(o);
                one = item.Str("source", item.Str("path", null));
            }
            if (one != null && one.Trim().Length > 0) sources.Add(one.Trim());
        }
        if (a.Has("items"))
        {
            foreach (object o in a.List("items"))
            {
                if (o is string) { sources.Add((string)o); continue; }
                Bag item = new Bag(o);
                string one = item.Str("source", item.Str("path", null));
                if (one != null) sources.Add(one);
            }
        }
        return sources;
    }
}
