// wps-office-cooperate (ppt): the delivery-and-collaboration trio.
//
//   comment  - put review remarks on the page they belong to (the Word side has had
//              this all along; a deck review without it means "page 7, third box",
//              which is exactly the ambiguity the user is trying to avoid);
//   export   - hand the deck over as PDF / a PNG series / a markdown outline;
//   picture  - list and swap every picture on the deck (unit logo, new renderings).
//
// Measured on WPS 12.1.0.29161 while writing this file:
//   * Slide.Comments.Add(Left, Top, Author, AuthorInitials, Text) - exactly 5 args
//     (3 and 6 both fail). Presentation.Comments stays 0: remarks are stored per slide.
//     The Left/Top we pass do NOT survive: 100,100 reads back as 5,5, so the anchor has
//     to be repeated in the comment text itself (we prefix "@<shape>: ").
//   * Presentation.SaveCopyAs(path, 32) writes a real PDF and leaves the open
//     presentation's identity alone; ExportAsFixedFormat is not usable from here.
//   * Shapes.AddPicture with a non-ASCII path fails outright (this machine's work
//     folder is "D:\DSH工作区\..."), so every insert goes through an ASCII copy in %TEMP%.
//   * Files written by wpp.exe are transparently encrypted on this machine, PDF
//     included - the response says so instead of pretending the file is ready to send.
//
// Pure ASCII source (Chinese only ever travels inside JSON strings), C# 5.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;

internal static partial class Program
{
    // ---------------------------------------------------------------- shared bits

    private const string PptTempDirName = "woc-ppt";

    /// <summary>Copy to an all-ASCII temp path: AddPicture cannot take the Chinese path
    /// this machine's work folder has.  The name is reduced to ASCII; the extension and a
    /// GUID keep it unique so two calls never collide.</summary>
    private static string AsciiTempCopy(string source)
    {
        string dir = Path.Combine(Path.GetTempPath(), PptTempDirName);
        Directory.CreateDirectory(dir);
        string ext = Path.GetExtension(source);
        StringBuilder safe = new StringBuilder();
        foreach (char ch in Path.GetFileNameWithoutExtension(source))
        {
            if (ch < 128 && (char.IsLetterOrDigit(ch) || ch == '-' || ch == '_')) safe.Append(ch);
            else if (ch == ' ') safe.Append('_');
        }
        if (safe.Length == 0) safe.Append("img");
        if (safe.Length > 40) safe.Length = 40;
        string dest = Path.Combine(dir, safe.ToString() + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ext.ToLowerInvariant());
        File.Copy(source, dest, true);
        return dest;
    }

    private static void TryDelete(string path)
    {
        try { if (path != null && File.Exists(path)) File.Delete(path); } catch (Exception) { }
    }

    /// <summary>True when the file on disk starts with the Agile DGS cipher marker.</summary>
    private static bool LooksEncrypted(string path)
    {
        try
        {
            if (!File.Exists(path)) return false;
            using (FileStream fs = File.OpenRead(path))
            {
                if (fs.Length < 8) return false;
                int b0 = fs.ReadByte(), b1 = fs.ReadByte(), b2 = fs.ReadByte();
                return b0 == 0x17 && b1 == 0xDA && b2 == 0x5F;
            }
        }
        catch (Exception) { return false; }
    }

    /// <summary>"1-5,8,10-" style slide list.  null = every slide.</summary>
    private static HashSet<int> ParseSlideSet(string spec, int slideCount)
    {
        if (spec == null || spec.Trim().Length == 0) return null;
        HashSet<int> want = new HashSet<int>();
        foreach (string raw in spec.Split(','))
        {
            string v = raw.Trim();
            if (v.Length == 0) continue;
            int dash = v.IndexOf('-');
            if (dash > 0)
            {
                int lo, hi;
                if (!int.TryParse(v.Substring(0, dash).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out lo))
                    throw new Fail("BAD_ARGS", "bad -Slides range \"" + v + "\"");
                string tail = v.Substring(dash + 1).Trim();
                if (tail.Length == 0) hi = slideCount;
                else if (!int.TryParse(tail, NumberStyles.Integer, CultureInfo.InvariantCulture, out hi))
                    throw new Fail("BAD_ARGS", "bad -Slides range \"" + v + "\"");
                if (hi < lo) { int t = lo; lo = hi; hi = t; }
                for (int i = lo; i <= hi; i++) if (i >= 1 && i <= slideCount) want.Add(i);
            }
            else
            {
                int one;
                if (!int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out one))
                    throw new Fail("BAD_ARGS", "bad -Slides value \"" + v + "\"");
                if (one >= 1 && one <= slideCount) want.Add(one);
            }
        }
        if (want.Count == 0) throw new Fail("BAD_ARGS", "-Slides matched nothing inside 1.." + slideCount + ": " + spec);
        return want;
    }

    private static string SlideSelection(int slideNo, HashSet<int> only)
    {
        return only == null || only.Contains(slideNo) ? null : "slide not in -Slides";
    }

    // ---------------------------------------------------------------- comment

    private static object CommentCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }
        string action = (a.Str("action", "add") ?? "add").ToLowerInvariant();
        if (action == "list") return CommentList(p, n);

        if (action == "delete")
        {
            int index = a.Int("index", 0);
            bool all = a.Flag("clear", false) || a.Flag("all", false);
            int slideNo = a.Int("slide", 0);
            if (!all && index <= 0) throw new Fail("BAD_ARGS", "comment -Action delete needs -Index <n> or -All");
            if (slideNo < 1) throw new Fail("BAD_ARGS", "comment -Action delete needs -Slide <n>");
            dynamic slide = SlideAt(p, slideNo);
            int before = 0;
            try { before = (int)slide.Comments.Count; } catch (Exception) { }
            if (before == 0) return new Dictionary<string, object> { { "slide", slideNo }, { "deleted", 0 } };
            int deleted = 0;
            if (all)
            {
                for (int i = before; i >= 1; i--)
                {
                    try { slide.Comments.Item(i).Delete(); deleted++; } catch (Exception) { }
                }
            }
            else
            {
                if (index > before) throw new Fail("BAD_ARGS", "slide " + slideNo + " has " + before + " comment(s); no index " + index);
                slide.Comments.Item(index).Delete();
                deleted = 1;
            }
            Dictionary<string, object> r = new Dictionary<string, object>();
            r["slide"] = slideNo;
            r["deleted"] = deleted;
            r["left"] = Math.Max(0, before - deleted);
            return r;
        }

        if (action != "add") throw new Fail("BAD_ARGS", "comment -Action must be add|list|delete");

        // add: one item, or a batch that rolls back as a whole
        string author = a.Str("author", "AI");
        List<object> items = new List<object>();
        if (a.Has("items"))
        {
            foreach (object o in a.List("items")) items.Add(o);
        }
        else items.Add(ItemFromBag(a, a.Int("slide", 0)));

        List<object> added = new List<object>();
        try
        {
            foreach (object o in items)
            {
                Bag item = new Bag(o);
                int slideNo = item.Int("slide", 0);
                if (slideNo < 1 || slideNo > n) throw new Fail("BAD_ARGS", "comment item needs slide 1.." + n);
                string text = item.Str("text", null);
                if (text == null && item.Has("textFile"))
                {
                    string file = item.Str("textFile", null);
                    if (!File.Exists(file)) throw new Fail("BAD_ARGS", "comment textFile not found: " + file);
                    text = File.ReadAllText(file, Encoding.UTF8).TrimEnd('\r', '\n');
                }
                if (text == null || text.Trim().Length == 0) throw new Fail("BAD_ARGS", "comment item needs text (or textFile)");

                dynamic slide = SlideAt(p, slideNo);
                double left = -1, top = -1;
                string shapeName = item.Str("shape", null);
                dynamic shape = null;
                if (shapeName != null)
                {
                    shape = FindShapeByName(slide, shapeName);
                    if (shape == null) throw new Fail("BAD_ARGS", "slide " + slideNo + " has no shape \"" + shapeName + "\"");
                    left = Convert.ToDouble(TryGet(() => (object)shape.Left, (object)40.0), CultureInfo.InvariantCulture);
                    top = Convert.ToDouble(TryGet(() => (object)shape.Top, (object)40.0), CultureInfo.InvariantCulture);
                    text = "@" + shapeName + "\uff1a" + text;   // 全角冒号：the anchor WPS keeps is the text, not Left/Top
                }
                if (left < 0) left = 40;
                if (top < 0) top = 40;
                string initials = item.Str("authorInitials", null);
                if (initials == null || initials.Length == 0) initials = author.Length > 0 ? author.Substring(0, 1) : "AI";

                dynamic made = slide.Comments.Add((float)left, (float)top, author, initials, text);
                Dictionary<string, object> rec = new Dictionary<string, object>();
                rec["slide"] = slideNo;
                if (shapeName != null) rec["shape"] = shapeName;
                rec["index"] = Convert.ToInt32(TryGet(() => (object)slide.Comments.Count, (object)0), CultureInfo.InvariantCulture);
                rec["text"] = Clip(Convert.ToString(TryGet(() => (object)made.Text, ""), CultureInfo.InvariantCulture), 120);
                added.Add(rec);
            }
        }
        catch (Exception e)
        {
            // atomic: a half-written batch of review remarks is worse than none
            int rolled = 0;
            foreach (object o in added)
            {
                Dictionary<string, object> rec = (Dictionary<string, object>)o;
                try
                {
                    int slideNo = Convert.ToInt32(rec["slide"], CultureInfo.InvariantCulture);
                    int index = Convert.ToInt32(rec["index"], CultureInfo.InvariantCulture);
                    dynamic slide = p.Slides.Item(slideNo);
                    if (index >= 1 && index <= (int)slide.Comments.Count) slide.Comments.Item(index).Delete();
                    rolled++;
                }
                catch (Exception) { }
            }
            throw new Fail("COMMENT_FAILED", Clean(e.Message)
                + (rolled > 0 ? " (rolled back " + rolled + " comment(s))" : ""));
        }

        Dictionary<string, object> result = new Dictionary<string, object>();
        result["added"] = added;
        result["count"] = added.Count;
        result["author"] = author;
        result["note"] = "WPS stores remarks per slide and pins the pin position itself, so each text "
                       + "carries \"@<shape>: \" to say which box it refers to";
        return result;
    }

    private static Dictionary<string, object> ItemFromBag(Bag a, int slide)
    {
        Dictionary<string, object> d = new Dictionary<string, object>();
        d["slide"] = slide;
        if (a.Has("shape")) d["shape"] = a.Str("shape", null);
        if (a.Has("text")) d["text"] = a.Str("text", null);
        if (a.Has("textFile")) d["textFile"] = a.Str("textFile", null);
        if (a.Has("authorInitials")) d["authorInitials"] = a.Str("authorInitials", null);
        return d;
    }

    private static dynamic FindShapeByName(dynamic slide, string name)
    {
        int count = 0;
        try { count = (int)slide.Shapes.Count; } catch (Exception) { }
        for (int i = 1; i <= count; i++)
        {
            dynamic sh = null;
            try { sh = slide.Shapes.Item(i); } catch (Exception) { continue; }
            string actual = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
            if (string.Equals(actual, name, StringComparison.OrdinalIgnoreCase)) return sh;
        }
        return null;
    }

    private static object CommentList(dynamic p, int n)
    {
        List<object> slides = new List<object>();
        int total = 0;
        for (int k = 1; k <= n; k++)
        {
            dynamic slide = null;
            try { slide = p.Slides.Item(k); } catch (Exception) { continue; }
            int count = 0;
            try { count = (int)slide.Comments.Count; } catch (Exception) { count = 0; }
            if (count == 0) continue;
            List<object> list = new List<object>();
            for (int i = 1; i <= count; i++)
            {
                dynamic c = null;
                try { c = slide.Comments.Item(i); } catch (Exception) { continue; }
                list.Add(new Dictionary<string, object> {
                    { "index", i },
                    { "author", Convert.ToString(TryGet(() => (object)c.Author, ""), CultureInfo.InvariantCulture) },
                    { "text", Clip(Convert.ToString(TryGet(() => (object)c.Text, ""), CultureInfo.InvariantCulture), 200) },
                    { "left", R(Convert.ToDouble(TryGet(() => (object)c.Left, (object)0.0), CultureInfo.InvariantCulture)) },
                    { "top", R(Convert.ToDouble(TryGet(() => (object)c.Top, (object)0.0), CultureInfo.InvariantCulture)) }
                });
            }
            total += list.Count;
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["slide"] = k;
            one["comments"] = list;
            slides.Add(one);
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["total"] = total;
        r["slides"] = slides;
        if (total == 0) r["note"] = "no comments on this deck";
        return r;
    }

    // ---------------------------------------------------------------- export

    private static object ExportCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }
        if (n == 0) throw new Fail("BAD_ARGS", "this presentation has no slides");

        string format = (a.Str("format", "pdf") ?? "pdf").ToLowerInvariant();
        string outArg = a.Str("out", null);
        if (outArg == null || outArg.Length == 0) throw new Fail("BAD_ARGS", "export needs -Out <path or folder>");
        string full = FullPathOf(outArg);
        HashSet<int> only = ParseSlideSet(a.Str("slides", null), n);

        if (format == "pdf")
        {
            string path = full.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? full : full + ".pdf";
            string dir = Path.GetDirectoryName(path);
            if (dir != null && dir.Length > 0) Directory.CreateDirectory(dir);
            if (LooksLikePlainFile(path)) TryDelete(path);
            try { p.SaveCopyAs(path, 32); }   // ppSaveAsPDF; SaveCopyAs keeps the open deck's identity
            catch (Exception e) { throw new Fail("COM_ERROR", "could not export PDF: " + Clean(e.Message)); }
            bool encrypted = LooksEncrypted(path);
            Dictionary<string, object> r = new Dictionary<string, object>();
            r["format"] = "pdf";
            r["out"] = path;
            r["slides"] = only == null ? n : only.Count;
            r["encrypted"] = encrypted;
            if (only != null) r["note"] = "PDF export always covers the whole deck; use -Format png for a page range";
            if (encrypted) r["note"] = "the file on disk is transparently encrypted (17 DA 5F A0 ...): decrypt it "
                                     + "(FastDycrypt) before sending it to anyone";
            return r;
        }

        if (format == "md" || format == "outline" || format == "json")
        {
            Bag outlineBag = new Bag(new Dictionary<string, object> { { "out", full }, { "doc", a.Need("doc") } });
            return OutlineCmd(outlineBag);   // one implementation of "what does the deck say"
        }

        if (format != "png" && format != "jpg" && format != "jpeg")
            throw new Fail("BAD_ARGS", "export -Format must be pdf|png|jpg|md");

        // PNG/JPEG go through the same clipboard->EMF path as render, because WPS would
        // write them transparently encrypted and they could never be read back.
        string folder = full;
        if (folder.EndsWith(".png", StringComparison.OrdinalIgnoreCase) || folder.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
            || folder.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase)) folder = Path.GetDirectoryName(full);
        if (folder == null || folder.Length == 0) folder = Path.GetDirectoryName(full);
        Directory.CreateDirectory(folder);

        int width = a.Int("width", 1600);
        int quality = a.Int("quality", 88);
        bool jpeg = format != "png";
        List<object> files = new List<object>();
        List<object> failed = new List<object>();
        for (int k = 1; k <= n; k++)
        {
            if (only != null && !only.Contains(k)) continue;
            string target = Path.Combine(folder, "slide-" + k.ToString("00", CultureInfo.InvariantCulture) + (jpeg ? ".jpg" : ".png"));
            try
            {
                Bitmap shot = CaptureSlide(p, k, width);
                if (shot == null) { failed.Add(new Dictionary<string, object> { { "slide", k }, { "reason", "clipboard empty" } }); continue; }
                SaveBitmap(shot, target, jpeg, quality);
                shot.Dispose();
                files.Add(target);
            }
            catch (Exception e)
            {
                failed.Add(new Dictionary<string, object> { { "slide", k }, { "reason", Clean(e.Message) } });
            }
        }

        Dictionary<string, object> result = new Dictionary<string, object>();
        result["format"] = jpeg ? "jpg" : "png";
        result["out"] = folder;
        result["files"] = files;
        result["count"] = files.Count;
        if (failed.Count > 0) result["failed"] = failed;
        result["note"] = "bitmaps are written by this helper process, so they stay plaintext (WPS would encrypt them)";
        return result;
    }

    /// <summary>The same clipboard route render uses: WPS would write an exported PNG
    /// transparently encrypted, so the bitmap has to be written by this process.</summary>
    private static Bitmap CaptureSlide(dynamic p, int slideNo, int width)
    {
        dynamic slide = p.Slides.Item(slideNo);
        slide.Copy();
        System.Threading.Thread.Sleep(250);
        return ClipboardBitmap(width);
    }

    private static bool LooksLikePlainFile(string path)
    {
        return File.Exists(path);
    }

    private static void SaveBitmap(Bitmap bmp, string path, bool jpeg, int quality)
    {
        if (!jpeg)
        {
            bmp.Save(path, ImageFormat.Png);
            return;
        }
        ImageCodecInfo codec = null;
        foreach (ImageCodecInfo c in ImageCodecInfo.GetImageEncoders())
            if (c.FormatID == ImageFormat.Jpeg.Guid) { codec = c; break; }
        if (codec == null) { bmp.Save(path, ImageFormat.Jpeg); return; }
        using (EncoderParameters pars = new EncoderParameters(1))
        {
            pars.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)quality);
            bmp.Save(path, codec, pars);
        }
    }

    // ---------------------------------------------------------------- picture

    private static object PictureCmd(Bag a)
    {
        dynamic p = RequirePres(a, true);
        int n = 0;
        try { n = (int)p.Slides.Count; } catch (Exception) { }
        string action = (a.Str("action", "list") ?? "list").ToLowerInvariant();
        HashSet<int> only = ParseSlideSet(a.Str("slides", null), n);

        if (action == "list") return PictureList(p, n, only);
        if (action != "replace") throw new Fail("BAD_ARGS", "picture -Action must be list|replace");

        string newPath = a.Str("new", null);
        if (newPath == null || newPath.Length == 0) throw new Fail("BAD_ARGS", "picture -Action replace needs -New <image>");
        if (!File.Exists(newPath)) throw new Fail("BAD_ARGS", "new picture not found: " + newPath);
        bool dryRun = a.Flag("dryRun", false);
        bool keepSize = a.Flag("keepSize", true);
        bool resetCrop = a.Flag("resetCrop", true);

        MatchSpec spec = MatchSpec.Parse(a.Str("match", null), a.Str("ids", null));

        string insertPath = AsciiTempCopy(newPath);
        List<object> hits = new List<object>();
        List<object> skipped = new List<object>();
        try
        {
            for (int k = 1; k <= n; k++)
            {
                if (only != null && !only.Contains(k)) continue;
                dynamic slide = null;
                try { slide = p.Slides.Item(k); } catch (Exception) { continue; }
                foreach (PictureRef pref in PicturesOf(slide))
                {
                    dynamic pic = pref.Shape;
                    bool grouped = pref.Path.IndexOf('/') >= 0;
                    string name = Convert.ToString(TryGet(() => (object)pic.Name, ""), CultureInfo.InvariantCulture);
                    double l = ToD(TryGet(() => (object)pic.Left, (object)0.0));
                    double t = ToD(TryGet(() => (object)pic.Top, (object)0.0));
                    double w = ToD(TryGet(() => (object)pic.Width, (object)0.0));
                    double h = ToD(TryGet(() => (object)pic.Height, (object)0.0));
                    string link = TryGet(() => (object)pic.LinkFormat.SourceFullName, null) as string;

                    if (!spec.Matches(k, name, l, t, w, h)) continue;

                    if (grouped)
                    {
                        // A child of a group carries coordinates in the group's own child
                        // space, which COM does not expose; guessing the transform would
                        // drop the logo somewhere random, so say so instead.
                        skipped.Add(new Dictionary<string, object> {
                            { "slide", k }, { "shape", pref.Path }, { "reason", "inside a group" },
                            { "detail", "child coordinates are relative to the group; replace this one by hand" } });
                        continue;
                    }

                    Dictionary<string, object> rec = new Dictionary<string, object>();
                    rec["slide"] = k;
                    rec["shape"] = name;
                    rec["before"] = new Dictionary<string, object> {
                        { "l", R(l) }, { "t", R(t) }, { "w", R(w) }, { "h", R(h) }
                    };
                    rec["linked"] = link != null;

                    if (dryRun) { rec["would"] = true; hits.Add(rec); continue; }

                    if (link != null)
                    {
                        // a linked picture: repoint the link, geometry cannot move
                        try
                        {
                            pic.LinkFormat.SourceFullName = newPath;
                            try { pic.LinkFormat.AutoUpdate = -1; } catch (Exception) { }
                            rec["how"] = "relinked";
                            hits.Add(rec);
                        }
                        catch (Exception e)
                        {
                            skipped.Add(new Dictionary<string, object> {
                                { "slide", k }, { "shape", name }, { "reason", "relink failed" }, { "detail", Clean(e.Message) } });
                        }
                        continue;
                    }

                    // embedded: insert the new picture, carry the geometry over, then drop the old one
                    double cl = 0, ct = 0, cr = 0, cb = 0;
                    if (!resetCrop)
                    {
                        cl = ToD(TryGet(() => (object)pic.PictureFormat.CropLeft, (object)0.0));
                        ct = ToD(TryGet(() => (object)pic.PictureFormat.CropTop, (object)0.0));
                        cr = ToD(TryGet(() => (object)pic.PictureFormat.CropRight, (object)0.0));
                        cb = ToD(TryGet(() => (object)pic.PictureFormat.CropBottom, (object)0.0));
                    }
                    try
                    {
                        dynamic fresh = slide.Shapes.AddPicture(insertPath, 0, -1, (float)l, (float)t,
                            keepSize ? (object)(float)w : Type.Missing, keepSize ? (object)(float)h : Type.Missing);
                        if (!keepSize)
                        {
                            // keep the old footprint but scale the new bitmap to fit inside it
                            double nw = ToD(TryGet(() => (object)fresh.Width, (object)w));
                            double nh = ToD(TryGet(() => (object)fresh.Height, (object)h));
                            double ratio = Math.Min(w / Math.Max(nw, 1), h / Math.Max(nh, 1));
                            fresh.Width = (float)(nw * ratio);
                            fresh.Height = (float)(nh * ratio);
                            fresh.Left = (float)(l + (w - nw * ratio) / 2);
                            fresh.Top = (float)(t + (h - nh * ratio) / 2);
                        }
                        if (!resetCrop && (cl != 0 || ct != 0 || cr != 0 || cb != 0))
                        {
                            try
                            {
                                fresh.PictureFormat.CropLeft = (float)cl;
                                fresh.PictureFormat.CropTop = (float)ct;
                                fresh.PictureFormat.CropRight = (float)cr;
                                fresh.PictureFormat.CropBottom = (float)cb;
                            }
                            catch (Exception) { }
                        }
                        string freshName = Convert.ToString(TryGet(() => (object)fresh.Name, ""), CultureInfo.InvariantCulture);
                        pic.Delete();
                        rec["after"] = new Dictionary<string, object> {
                            { "name", freshName },
                            { "l", R(ToD(TryGet(() => (object)fresh.Left, (object)l))) },
                            { "t", R(ToD(TryGet(() => (object)fresh.Top, (object)t))) },
                            { "w", R(ToD(TryGet(() => (object)fresh.Width, (object)w))) },
                            { "h", R(ToD(TryGet(() => (object)fresh.Height, (object)h))) }
                        };
                        rec["how"] = "replaced";
                        hits.Add(rec);
                    }
                    catch (Exception e)
                    {
                        skipped.Add(new Dictionary<string, object> {
                            { "slide", k }, { "shape", name }, { "reason", "insert failed" }, { "detail", Clean(e.Message) } });
                    }
                }
            }
        }
        finally { TryDelete(insertPath); }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["new"] = newPath;
        r["changed"] = dryRun ? 0 : CountReplaced(hits);
        r["wouldChange"] = dryRun ? hits.Count : 0;
        r["dryRun"] = dryRun;
        if (hits.Count > 0) r["hits"] = hits;
        if (skipped.Count > 0) r["skipped"] = skipped;
        r["note"] = "geometry is carried over from the old picture; run render on the changed slides, because a new "
                  + "bitmap with a different aspect ratio looks fine in numbers and wrong on screen";
        return r;
    }

    private static int CountReplaced(List<object> hits)
    {
        int count = 0;
        foreach (object o in hits)
        {
            Dictionary<string, object> rec = (Dictionary<string, object>)o;
            if (rec.ContainsKey("how")) count++;
        }
        return count;
    }

    private static double ToD(object v)
    {
        if (v == null) return 0;
        try { return Convert.ToDouble(v, CultureInfo.InvariantCulture); } catch (Exception) { return 0; }
    }

    /// <summary>Every picture on the slide, groups included (a unit logo very often sits
    /// inside a group with its wordmark).  Path keeps the group chain readable.</summary>
    private static List<PictureRef> PicturesOf(dynamic slide)
    {
        List<PictureRef> found = new List<PictureRef>();
        CollectPictures(slide.Shapes, "", found);
        return found;
    }

    private sealed class PictureRef
    {
        public dynamic Shape;
        public string Path = "";
    }

    private static void CollectPictures(dynamic shapes, string prefix, List<PictureRef> found)
    {
        int count = 0;
        try { count = (int)shapes.Count; } catch (Exception) { return; }
        for (int i = 1; i <= count; i++)
        {
            dynamic sh = null;
            try { sh = shapes.Item(i); } catch (Exception) { continue; }
            int type = Convert.ToInt32(TryGet(() => (object)(int)sh.Type, (object)0), CultureInfo.InvariantCulture);
            if (type == 6)
            {
                string name = Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                dynamic items = null;
                try { items = sh.GroupItems; } catch (Exception) { continue; }
                CollectPictures(items, prefix + name + "/", found);
                continue;
            }
            if (type == 13 || type == 11)
            {
                PictureRef reference = new PictureRef();
                reference.Shape = sh;
                reference.Path = prefix + Convert.ToString(TryGet(() => (object)sh.Name, ""), CultureInfo.InvariantCulture);
                found.Add(reference);
            }
        }
    }

    private static object PictureList(dynamic p, int n, HashSet<int> only)
    {
        List<object> slides = new List<object>();
        int total = 0;
        for (int k = 1; k <= n; k++)
        {
            if (only != null && !only.Contains(k)) continue;
            dynamic slide = null;
            try { slide = p.Slides.Item(k); } catch (Exception) { continue; }
            List<object> pics = new List<object>();
            int idx = 0;
            foreach (PictureRef pref in PicturesOf(slide))
            {
                idx++;
                dynamic pic = pref.Shape;
                string link = TryGet(() => (object)pic.LinkFormat.SourceFullName, null) as string;
                pics.Add(new Dictionary<string, object> {
                    { "n", idx },
                    { "name", Convert.ToString(TryGet(() => (object)pic.Name, ""), CultureInfo.InvariantCulture) },
                    { "path", pref.Path },
                    { "group", pref.Path.IndexOf('/') >= 0 },
                    { "l", R(ToD(TryGet(() => (object)pic.Left, (object)0.0))) },
                    { "t", R(ToD(TryGet(() => (object)pic.Top, (object)0.0))) },
                    { "w", R(ToD(TryGet(() => (object)pic.Width, (object)0.0))) },
                    { "h", R(ToD(TryGet(() => (object)pic.Height, (object)0.0))) },
                    { "linked", link != null },
                    { "link", link }
                });
            }
            if (pics.Count == 0) continue;
            total += pics.Count;
            Dictionary<string, object> one = new Dictionary<string, object>();
            one["slide"] = k;
            one["pictures"] = pics;
            slides.Add(one);
        }
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["total"] = total;
        r["slides"] = slides;
        if (total == 0) r["note"] = "no pictures on this deck";
        else r["note"] = "pick the ones you mean with -Match \"name:图片 3\" | \"size:56x57\" | \"pos:0,0\" | -Ids \"12,3,4\", then re-run with -Action replace -New <file>";
        return r;
    }

    /// <summary>Which pictures to touch: by name, by displayed size (WxH within 1pt),
    /// by position (L,T within 1pt), or by the ordinal shown in -Action list.</summary>
    private sealed class MatchSpec
    {
        public string Name;
        public double W = -1, H = -1;
        public double L = double.NaN, T = double.NaN;
        public HashSet<int> Ids;

        public static MatchSpec Parse(string match, string ids)
        {
            MatchSpec spec = new MatchSpec();
            if (ids != null && ids.Trim().Length > 0)
            {
                spec.Ids = new HashSet<int>();
                foreach (string raw in ids.Split(','))
                {
                    int one;
                    if (int.TryParse(raw.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out one))
                        spec.Ids.Add(one);
                }
            }
            if (match == null || match.Trim().Length == 0) return spec;
            string m = match.Trim();
            int colon = m.IndexOf(':');
            if (colon <= 0) throw new Fail("BAD_ARGS", "-Match must look like \"name:<shape>\" | \"size:<W>x<H>\" | \"pos:<L>,<T>\"");
            string kind = m.Substring(0, colon).Trim().ToLowerInvariant();
            string value = m.Substring(colon + 1).Trim();
            if (kind == "name") spec.Name = value;
            else if (kind == "size")
            {
                string[] parts = value.Split('x', 'X', '*');
                if (parts.Length != 2) throw new Fail("BAD_ARGS", "-Match size wants WxH in points, e.g. size:56x57");
                spec.W = double.Parse(parts[0], CultureInfo.InvariantCulture);
                spec.H = double.Parse(parts[1], CultureInfo.InvariantCulture);
            }
            else if (kind == "pos" || kind == "position")
            {
                string[] parts = value.Split(',');
                if (parts.Length != 2) throw new Fail("BAD_ARGS", "-Match pos wants L,T in points, e.g. pos:0,0");
                spec.L = double.Parse(parts[0], CultureInfo.InvariantCulture);
                spec.T = double.Parse(parts[1], CultureInfo.InvariantCulture);
            }
            else throw new Fail("BAD_ARGS", "unknown -Match kind \"" + kind + "\"; use name|size|pos");
            return spec;
        }

        public bool Matches(int slide, string name, double l, double t, double w, double h)
        {
            bool any = Name != null || W >= 0 || !double.IsNaN(L) || Ids != null;
            if (!any) return true;      // no filter: every picture
            if (Ids != null && Ids.Contains(slide)) return true;
            if (Name != null && string.Equals(Name, name, StringComparison.OrdinalIgnoreCase)) return true;
            if (W >= 0 && Math.Abs(w - W) <= 1 && Math.Abs(h - H) <= 1) return true;
            if (!double.IsNaN(L) && Math.Abs(l - L) <= 1 && Math.Abs(t - T) <= 1) return true;
            return false;
        }
    }
}
