// v0.4 collaboration: see what the human did (revisions), and what changed since a snapshot.
//
//   revisions  list / accept / reject tracked changes
//   snapshot   take a text snapshot of the document
//   diff       what changed since a snapshot (added / removed / rewritten paragraphs)

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

internal static partial class Program
{
    // wdRevisionType that matter here.
    private const int REV_INSERT = 1;
    private const int REV_DELETE = 2;
    private const int REV_PROPERTY = 3;

    // Beyond this many paragraph pairs diff falls back to a positional walk (2000 x 2000).
    private const long DIFF_CELLS_MAX = 4000000;

    // Above this many paragraphs the revision->paragraph lookup probes instead of building a table.
    private const int PARA_TABLE_MAX = 2000;

    /// <summary>
    /// revisions — what is still pending in the document, and let them be accepted or rejected.
    ///   {"doc":"...","action":"list"|"accept"|"reject","index":N|0,"all":true}
    /// list    -> {"total":N,"items":[{"i","type","author","date","para","text","kind"}]}
    /// accept  -> {"action":"accept","index":N|"all","done":N,"revisionsAfter":N}
    /// reject  -> same
    /// type: 1 insert, 2 delete, 3 format (wdRevisionInsert/Delete/Property).
    /// </summary>
    private static object RevisionsDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        string action = a.Str("action", "list").Trim().ToLowerInvariant();
        if (action.Length == 0) action = "list";
        if (action != "list" && action != "accept" && action != "reject")
            throw new Fail("BAD_ARGS", "unknown action \"" + action + "\" (list|accept|reject)");

        if (action == "list") return RevisionsList(doc);

        if (IsReadOnly(doc))
            throw new Fail("READONLY", "The document is open read-only; no revision was " + (action == "accept" ? "accepted" : "rejected") + ".");
        int total = RevisionCount(doc);
        bool all = a.Flag("all", false);
        int index = a.Int("index", 0);
        bool accept = action == "accept";
        if (!all)
        {
            if (index < 1) throw new Fail("BAD_ARGS", "give \"index\" (1-based) or \"all\":true");
            if (index > total)
                throw new Fail("BAD_ARGS", "revision " + index + " does not exist: the document has " + total + " pending revision(s) right now (0 = none).");
        }

        dynamic undo = BeginUndo("wps-word-cooperate " + action + " revisions");
        bool wasTracking = Tracking(doc);
        // Accepting/rejecting must not be recorded as a new change of its own.
        SetTracking(doc, false);
        try
        {
            if (total == 0) { /* nothing to decide; reported below */ }
            else if (all) RevisionDecideAll(doc, accept);
            else
            {
                try
                {
                    if (accept) doc.Revisions.Item(index).Accept();
                    else doc.Revisions.Item(index).Reject();
                }
                catch (COMException e)
                {
                    throw new Fail("COM_ERROR", "WPS could not " + action + " revision " + index + ": " + Clean(e.Message)
                        + ". Check the document again (a revision may have been accepted by hand), then retry.");
                }
            }
        }
        finally
        {
            SetTracking(doc, wasTracking);
            EndUndo(undo);
        }

        int after = RevisionCount(doc);
        int done = Math.Max(0, total - after);
        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = DocName(doc);
        r["action"] = action;
        r["index"] = all ? 0 : index;
        r["all"] = all;
        r["done"] = done;
        r["revisionsBefore"] = total;
        r["revisionsAfter"] = after;
        r["paragraphs"] = TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        r["saved"] = false;
        if (total == 0)
            r["note"] = "文档里没有待处理的修订，什么都没做（已被人接受/拒绝的修订不在集合里）。";
        else if (done == 0)
            r["note"] = "WPS 报告没有修订被" + (accept ? "接受" : "拒绝") + "，集合仍是 " + after + " 条；文档状态可能刚被人改过，重跑 revisions list 看看。";
        else
            r["note"] = "已" + (accept ? "接受" : "拒绝") + " " + done + " 条，还剩 " + after + " 条待处理；"
                + (all ? "" : "段落编号可能变了，") + "一次 Ctrl+Z 可整批撤回。改动没有保存。";
        return r;
    }

    private static int RevisionCount(dynamic doc)
    {
        return (int)TryGet(() => (object)(int)doc.Revisions.Count, 0);
    }

    /// <summary>Accept or reject every pending revision; returns how many actually went away.</summary>
    private static int RevisionDecideAll(dynamic doc, bool accept)
    {
        int before = RevisionCount(doc);
        if (before == 0) return 0;
        try
        {
            if (accept) doc.Revisions.AcceptAll();
            else doc.Revisions.RejectAll();
        }
        catch (Exception) { }
        int done = Math.Max(0, before - RevisionCount(doc));
        if (done > 0) return done;

        // WPS versions without a working AcceptAll/RejectAll: one at a time, from the end.
        for (int i = before; i >= 1; i--)
        {
            try
            {
                if (accept) doc.Revisions.Item(i).Accept();
                else doc.Revisions.Item(i).Reject();
                done++;
            }
            catch (Exception) { }
        }
        return done;
    }

    private static object RevisionsList(dynamic doc)
    {
        int total = RevisionCount(doc);
        int paragraphs = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        int[] starts = null;
        if (total > 0 && paragraphs > 0 && paragraphs <= PARA_TABLE_MAX) starts = ParagraphStarts(doc, paragraphs);

        Dictionary<string, object> byType = new Dictionary<string, object>();
        byType["insert"] = 0;
        byType["delete"] = 0;
        byType["format"] = 0;
        Dictionary<string, object> authors = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        List<object> items = new List<object>();

        for (int i = 1; i <= total; i++)
        {
            dynamic rev = null;
            try { rev = doc.Revisions.Item(i); } catch (Exception) { continue; }
            int type = (int)TryGet(() => (object)(int)rev.Type, 0);
            string kind = RevisionKind(type);
            string author = Convert.ToString(TryGet(() => (object)rev.Author, ""), CultureInfo.InvariantCulture);
            string date = RevisionDate(TryGet(() => (object)rev.Date, null));
            int pos = -1;
            try { pos = (int)rev.Range.Start; } catch (Exception) { pos = -1; }
            int para = 0;
            if (pos >= 0 && paragraphs > 0)
                para = starts != null ? ParaAtOffset(starts, paragraphs, pos) : ParaAtOffsetProbe(doc, paragraphs, pos);

            string text = "";
            try { text = CleanPara(Convert.ToString(rev.Range.Text, CultureInfo.InvariantCulture)); } catch (Exception) { }
            if (text.Length == 0 && pos >= 0)
            {
                int end = -1;
                try { end = (int)rev.Range.End; } catch (Exception) { end = -1; }
                if (end > pos)
                    text = CleanPara(Convert.ToString(TryGet(() => (object)doc.Range(pos, end).Text, ""), CultureInfo.InvariantCulture));
            }

            Dictionary<string, object> one = new Dictionary<string, object>();
            one["i"] = i;
            one["type"] = type;
            one["kind"] = kind;
            one["author"] = author;
            one["date"] = date;
            one["para"] = para;
            one["text"] = Head(text, 80);
            items.Add(one);

            if (!byType.ContainsKey(kind)) byType[kind] = 0;
            byType[kind] = ToInt(byType, kind) + 1;
            if (author == null) author = "";
            if (!authors.ContainsKey(author)) authors[author] = 0;
            authors[author] = ToInt(authors, author) + 1;
        }

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = DocName(doc);
        r["readOnly"] = IsReadOnly(doc);
        r["action"] = "list";
        r["total"] = total;
        r["byType"] = byType;
        r["authors"] = authors;
        r["paragraphs"] = paragraphs;
        r["items"] = items;
        r["saved"] = false;
        r["note"] = total == 0
            ? "文档里没有待处理的修订（已接受/拒绝的修订不在集合里）。"
            : "这 " + total + " 条就是全部待裁决的修订：已被人接受/拒绝的不在集合里。paragraph 编号是修订当前所在的段落。";
        return r;
    }

    /// <summary>Short name for a wdRevisionType value.</summary>
    private static string RevisionKind(int type)
    {
        switch (type)
        {
            case REV_INSERT: return "insert";
            case REV_DELETE: return "delete";
            case REV_PROPERTY: return "format";
            case 4: return "paragraph_number";
            case 5: return "display_field";
            case 8: return "style";
            case 9: return "replace";
            case 10: return "paragraph_property";
            case 11: return "table_property";
            case 12: return "section_property";
            case 13: return "style_definition";
            case 14: return "moved_from";
            case 15: return "moved_to";
            case 16: return "cell_insert";
            case 17: return "cell_delete";
            case 18: return "cell_merge";
            default: return "other";
        }
    }

    private static string RevisionDate(object raw)
    {
        if (raw == null) return "";
        if (raw is DateTime) return ((DateTime)raw).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        return Convert.ToString(raw, CultureInfo.InvariantCulture);
    }

    // ------------------------------------------------------------ revision -> paragraph

    /// <summary>Character offset of every paragraph's start, 1-based (starts[1] is paragraph 1).</summary>
    private static int[] ParagraphStarts(dynamic doc, int total)
    {
        int[] starts = new int[total + 1];
        for (int k = 1; k <= total; k++)
            starts[k] = (int)TryGet(() => (object)(int)doc.Paragraphs.Item(k).Range.Start, k <= 1 ? 0 : starts[k - 1]);
        return starts;
    }

    /// <summary>The 1-based paragraph holding a character offset.</summary>
    private static int ParaAtOffset(int[] starts, int total, int pos)
    {
        if (pos <= starts[1]) return 1;
        int lo = 1, hi = total, best = 1;
        while (lo <= hi)
        {
            int mid = (lo + hi) / 2;
            if (starts[mid] <= pos) { best = mid; lo = mid + 1; }
            else hi = mid - 1;
        }
        return best;
    }

    /// <summary>Same answer without a whole offset table, for very long documents.</summary>
    private static int ParaAtOffsetProbe(dynamic doc, int total, int pos)
    {
        int n = (int)TryGet(() => (object)(int)doc.Range(0, pos).Paragraphs.Count, 1);
        if (n < 1) n = 1;
        if (n > total) n = total;
        int start = (int)TryGet(() => (object)(int)doc.Paragraphs.Item(n).Range.Start, -1);
        int end = (int)TryGet(() => (object)(int)doc.Paragraphs.Item(n).Range.End, -1);
        if (start >= 0 && pos < start) return Math.Max(1, n - 1);
        if (end >= 0 && pos >= end) return Math.Min(total, n + 1);
        return n;
    }

    // ------------------------------------------------------------ snapshot

    /// <summary>
    /// snapshot — remember the text of every paragraph, so diff can say what changed afterwards.
    ///   {"doc":"...","tag":"before-round2"}
    /// Returns {"path","tag","paragraphs","digest","at"}.
    /// Snapshots live under %LOCALAPPDATA%\wps-word-cooperate\snapshots\<hash>\<tag>.json;
    /// "last" always means the newest snapshot of that document.
    /// </summary>
    private static object SnapshotDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        string dir = SnapshotDirOf(doc);
        Directory.CreateDirectory(dir);

        string at = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
        string tag = a.Str("tag", null);
        if (tag != null) tag = tag.Trim();
        if (tag == null || tag.Length == 0 || string.Equals(tag, "last", StringComparison.OrdinalIgnoreCase))
            tag = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
        tag = SafeTag(tag);
        if (tag.Length == 0) throw new Fail("BAD_ARGS", "\"tag\" has no file-name characters left after cleanup");

        List<string> texts = DocTexts(doc);
        string digest = ShortHash(SHA256.Create(), string.Join("\n", texts.ToArray()), 16);
        int paragraphs = texts.Count;
        int page = (int)TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        int revisions = RevisionCount(doc);
        string full = Convert.ToString(TryGet(() => (object)doc.FullName, ""), CultureInfo.InvariantCulture);

        Dictionary<string, object> payload = new Dictionary<string, object>();
        payload["doc"] = full;
        payload["tag"] = tag;
        payload["at"] = at;
        payload["paragraphs"] = paragraphs;
        payload["page"] = page;
        payload["revisions"] = revisions;
        payload["texts"] = texts;

        string path = Path.Combine(dir, tag + ".json");
        bool portable = false;
        string outPath = a.Str("out", null);
        if (outPath != null && outPath.Trim().Length > 0)
        {
            // Portable snapshot: written next to the document (or wherever asked) so another
            // machine / agent / colleague can diff against it. Not visible to "last".
            path = FullPathOf(outPath.Trim());
            string outDir = Path.GetDirectoryName(path);
            if (outDir != null && outDir.Length > 0) Directory.CreateDirectory(outDir);
            portable = true;
        }
        File.WriteAllText(path, Json.Write(payload), new UTF8Encoding(false));

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = DocName(doc);
        r["path"] = path;
        r["tag"] = tag;
        r["at"] = at;
        r["paragraphs"] = paragraphs;
        r["page"] = page;
        r["revisions"] = revisions;
        r["digest"] = digest;
        r["saved"] = false;
        r["note"] = portable
            ? "\u5feb\u7167\u5df2\u6309\u8981\u6c42\u5199\u5230\u6307\u5b9a\u8def\u5f84\uff08\u4fbf\u4e8e\u8de8\u673a\u5668/\u4ea4\u63a5\u65f6\u968f\u6587\u6863\u8d70\uff09\u3002diff \u65f6\u7528 -Since <\u8fd9\u4e2a\u8def\u5f84> \u624d\u80fd\u627e\u5230\u5b83\uff08\u5b83\u4e0d\u53c2\u4e0e last\uff09\u3002"
            : "\u5feb\u7167\u5b58\u5728\u7528\u6237\u76ee\u5f55\u4e0b\uff08\u4e0d\u5199\u6587\u6863\u76ee\u5f55\uff09\uff0c\u6587\u6863\u672c\u8eab\u6ca1\u88ab\u6539\u52a8\u3001\u4e5f\u6ca1\u6709\u4fdd\u5b58\u3002diff -Since last \u5c31\u662f\u8ddf\u8fd9\u4efd\u6bd4\u3002\u8981\u8de8\u673a\u5668/\u4ea4\u63a5\u7ed9\u522b\u4eba\uff0c\u52a0 -Out <\u8def\u5f84> \u5199\u5230\u6587\u6863\u65c1\u8fb9\u3002";
        return r;
    }

    /// <summary>Every paragraph's text, cleaned, in document order.</summary>
    private static List<string> DocTexts(dynamic doc)
    {
        List<string> texts = new List<string>();
        int total = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        if (total <= 0) return texts;

        // Fast path: one COM call instead of one per paragraph. The split must reproduce the
        // paragraph count EXACTLY, and the document must carry NO unaccepted revisions (deleted
        // text hides from Content.Text but still occupies Range coordinates). Table markers add
        // stray \r\a, so most real reports fall through to the loop below (correct, slower).
        // Both snapshot and diff read through here, so the two sides stay comparable.
        int tables = (int)TryGet(() => (object)(int)doc.Tables.Count, 0);
        int revs = (int)TryGet(() => (object)(int)doc.Revisions.Count, 1);
        if (tables == 0 && revs == 0 && total > 200)
        {
            string all = Convert.ToString(TryGet(() => (object)doc.Content.Text, null), CultureInfo.InvariantCulture);
            if (all != null && all.Length > 0)
            {
                List<string> fast = new List<string>();
                int from = 0;
                for (int k = 0; k < all.Length; k++)
                {
                    if (all[k] != '\r') continue;
                    fast.Add(CleanPara(all.Substring(from, k - from)));
                    from = k + 1;
                }
                if (from < all.Length) fast.Add(CleanPara(all.Substring(from)));
                if (fast.Count == total + 1 && fast[total].Length == 0) fast.RemoveAt(total);
                if (fast.Count == total) return fast;
            }
        }

        for (int k = 1; k <= total; k++)
        {
            string t = "";
            try { t = CleanPara(Convert.ToString(doc.Paragraphs.Item(k).Range.Text, CultureInfo.InvariantCulture)); }
            catch (Exception) { }
            texts.Add(t);
        }
        return texts;
    }

    private static string SnapshotRoot()
    {
        string baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (baseDir == null || baseDir.Length == 0) baseDir = Path.GetTempPath();
        return Path.Combine(Path.Combine(baseDir, "wps-office-cooperate"), "word-snapshots");
    }

    /// <summary>%LOCALAPPDATA%\wps-word-cooperate\snapshots\<sha1 of the doc path, 16 hex>.</summary>
    private static string SnapshotDirOf(dynamic doc)
    {
        string full = Convert.ToString(TryGet(() => (object)doc.FullName, ""), CultureInfo.InvariantCulture);
        if (full == null || full.Length == 0)
            throw new Fail("BAD_ARGS", "The document has no file on disk yet; save it first so a snapshot can be tied to it.");
        string key = ShortHash(SHA1.Create(), full.ToLowerInvariant(), 16);
        return Path.Combine(SnapshotRoot(), key);
    }

    private static string SafeTag(string tag)
    {
        char[] bad = Path.GetInvalidFileNameChars();
        StringBuilder sb = new StringBuilder();
        foreach (char c in tag)
        {
            bool ok = true;
            for (int i = 0; i < bad.Length; i++) if (bad[i] == c) { ok = false; break; }
            if (ok) sb.Append(c);
        }
        string s = sb.ToString().Trim().Trim('.').Trim();
        if (s.Length > 60) s = s.Substring(0, 60);
        return s;
    }

    private static string ShortHash(HashAlgorithm algorithm, string text, int hexLength)
    {
        byte[] bytes = algorithm.ComputeHash(Encoding.UTF8.GetBytes(text));
        StringBuilder sb = new StringBuilder(bytes.Length * 2);
        foreach (byte b in bytes) sb.Append(b.ToString("x2", CultureInfo.InvariantCulture));
        string hex = sb.ToString();
        return hex.Length <= hexLength ? hex : hex.Substring(0, hexLength);
    }

    /// <summary>The newest snapshot file of one document (by write time), or null.</summary>
    private static string NewestSnapshot(string dir)
    {
        if (!Directory.Exists(dir)) return null;
        string best = null;
        DateTime bestTime = DateTime.MinValue;
        foreach (string f in Directory.GetFiles(dir, "*.json"))
        {
            DateTime t = File.GetLastWriteTimeUtc(f);
            if (best == null || t > bestTime) { best = f; bestTime = t; }
        }
        return best;
    }

    // ------------------------------------------------------------ diff

    /// <summary>One step of the paragraph alignment: 0 same, 1 rewritten, 2 removed, 3 added.</summary>
    private sealed class DiffOp
    {
        public int Kind;
        public string A;      // text in the snapshot (removed / before)
        public string B;      // text now (added / after)
        public int AIndex;    // 0-based position in the snapshot
        public int BIndex;    // 0-based position now
    }

    /// <summary>
    /// diff — compare the document as it is now with a snapshot (default: the newest one).
    ///   {"doc":"...","since":"<path>|last"}
    /// Returns {"from","tag","summary":{"changed","added","removed","same","revisions","paragraphs"},
    ///          "changed":[{"para","before","after"}],"added":[...],"removed":[...]}
    /// Costs: O(n*m) over paragraphs, capped - a 600-paragraph report is instant.
    /// </summary>
    private static object DiffDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        string dir = SnapshotDirOf(doc);
        string since = a.Str("since", "last");
        if (since != null) since = since.Trim();
        if (since == null || since.Length == 0) since = "last";

        string snapPath = null;
        if (string.Equals(since, "last", StringComparison.OrdinalIgnoreCase))
        {
            snapPath = NewestSnapshot(dir);
            if (snapPath == null)
                throw new Fail("BAD_ARGS", "No snapshot of this document yet. Run snapshot first (snapshot -Doc ... -Tag base), make the change, then diff -Since base.");
        }
        else
        {
            string direct = FullPathOf(since);
            if (File.Exists(direct)) snapPath = direct;
            else
            {
                string byTag = Path.Combine(dir, SafeTag(since) + ".json");
                if (File.Exists(byTag)) snapPath = byTag;
            }
            if (snapPath == null)
                throw new Fail("BAD_ARGS", "Snapshot not found: " + since + " (neither an existing file nor a tag in " + dir + "). Run snapshot first, or pass the path printed by snapshot.");
        }

        Dictionary<string, object> snap = null;
        try { snap = Json.Parse(File.ReadAllText(snapPath, Encoding.UTF8)) as Dictionary<string, object>; }
        catch (Exception e) { throw new Fail("BAD_ARGS", "The snapshot file cannot be read as JSON: " + Clean(e.Message)); }
        if (snap == null || !snap.ContainsKey("texts"))
            throw new Fail("BAD_ARGS", "That file is not a snapshot (it has no \"texts\"): " + snapPath);

        List<string> before = new List<string>();
        List<object> raw = snap["texts"] as List<object>;
        if (raw != null)
            foreach (object o in raw) before.Add(Convert.ToString(o, CultureInfo.InvariantCulture) ?? "");
        List<string> now = DocTexts(doc);

        List<object> changed = new List<object>();
        List<object> added = new List<object>();
        List<object> removed = new List<object>();
        int same = 0;
        bool degraded = (long)before.Count * (long)now.Count > DIFF_CELLS_MAX;
        if (degraded) PositionalDiff(before, now, changed, added, removed, ref same);
        else DiffByLcs(before, now, changed, added, removed, ref same);

        Dictionary<string, object> summary = new Dictionary<string, object>();
        summary["changed"] = changed.Count;
        summary["added"] = added.Count;
        summary["removed"] = removed.Count;
        summary["same"] = same;
        summary["paragraphs"] = now.Count;
        summary["revisions"] = RevisionCount(doc);

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = DocName(doc);
        r["from"] = snapPath;
        r["tag"] = Convert.ToString(TryGet(() => (object)snap["tag"], ""), CultureInfo.InvariantCulture);
        r["at"] = Convert.ToString(TryGet(() => (object)snap["at"], ""), CultureInfo.InvariantCulture);
        r["summary"] = summary;
        r["changed"] = changed;
        r["added"] = added;
        r["removed"] = removed;
        if (degraded) r["degraded"] = true;
        r["saved"] = false;
        r["note"] = "changed 里既可能是 Agent 改的，也可能是人改的：未处理的修订=待裁决（用 revisions 看），已变文本=已生效。"
            + (degraded ? " 段落数超过 " + before.Count + "x" + now.Count + " 上限，已退化为分段位置比对：插入/删除可能被报成改写。" : "");
        return r;
    }

    /// <summary>Paragraph pairs up to +/- 4M: LCS alignment, rewrite runs paired as changed.</summary>
    private static void DiffByLcs(List<string> before, List<string> now, List<object> changed, List<object> added, List<object> removed, ref int same)
    {
        int n = before.Count, m = now.Count;
        int[,] lcs = new int[n + 1, m + 1];
        for (int i = n - 1; i >= 0; i--)
            for (int j = m - 1; j >= 0; j--)
                lcs[i, j] = string.Equals(before[i], now[j], StringComparison.Ordinal)
                    ? lcs[i + 1, j + 1] + 1
                    : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);

        List<DiffOp> ops = new List<DiffOp>();
        int x = 0, y = 0;
        while (x < n && y < m)
        {
            if (string.Equals(before[x], now[y], StringComparison.Ordinal))
            {
                DiffOp op = new DiffOp();
                op.Kind = 0; op.A = before[x]; op.B = now[y]; op.AIndex = x; op.BIndex = y;
                ops.Add(op);
                x++; y++;
            }
            else if (lcs[x + 1, y] >= lcs[x, y + 1])
            {
                DiffOp op = new DiffOp();
                op.Kind = 2; op.A = before[x]; op.AIndex = x; op.BIndex = y;
                ops.Add(op);
                x++;
            }
            else
            {
                DiffOp op = new DiffOp();
                op.Kind = 3; op.B = now[y]; op.AIndex = x; op.BIndex = y;
                ops.Add(op);
                y++;
            }
        }
        while (x < n)
        {
            DiffOp op = new DiffOp();
            op.Kind = 2; op.A = before[x]; op.AIndex = x; op.BIndex = m;
            ops.Add(op);
            x++;
        }
        while (y < m)
        {
            DiffOp op = new DiffOp();
            op.Kind = 3; op.B = now[y]; op.AIndex = n; op.BIndex = y;
            ops.Add(op);
            y++;
        }

        // A run without a match is one rewrite: pair the removals with the additions in order.
        int k = 0;
        while (k < ops.Count)
        {
            if (ops[k].Kind == 0) { same++; k++; continue; }
            int j = k;
            List<DiffOp> dels = new List<DiffOp>();
            List<DiffOp> ins = new List<DiffOp>();
            while (j < ops.Count && ops[j].Kind != 0)
            {
                if (ops[j].Kind == 2) dels.Add(ops[j]); else ins.Add(ops[j]);
                j++;
            }
            int pairs = Math.Min(dels.Count, ins.Count);
            for (int q = 0; q < pairs; q++)
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["para"] = ins[q].BIndex + 1;
                one["before"] = Head(dels[q].A, 120);
                one["after"] = Head(ins[q].B, 120);
                changed.Add(one);
            }
            for (int q = pairs; q < dels.Count; q++)
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["after"] = dels[q].BIndex + 1;
                one["text"] = Head(dels[q].A, 120);
                removed.Add(one);
            }
            for (int q = pairs; q < ins.Count; q++)
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["para"] = ins[q].BIndex + 1;
                one["text"] = Head(ins[q].B, 120);
                added.Add(one);
            }
            k = j;
        }
    }

    /// <summary>Fallback for very long documents: compare position by position, no alignment.</summary>
    private static void PositionalDiff(List<string> before, List<string> now, List<object> changed, List<object> added, List<object> removed, ref int same)
    {
        int n = before.Count, m = now.Count;
        int count = Math.Max(n, m);
        for (int i = 0; i < count; i++)
        {
            bool inBefore = i < n, inNow = i < m;
            if (inBefore && inNow)
            {
                if (string.Equals(before[i], now[i], StringComparison.Ordinal)) { same++; continue; }
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["para"] = i + 1;
                one["before"] = Head(before[i], 120);
                one["after"] = Head(now[i], 120);
                changed.Add(one);
            }
            else if (inBefore)
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["after"] = Math.Min(i + 1, m + 1);
                one["text"] = Head(before[i], 120);
                removed.Add(one);
            }
            else
            {
                Dictionary<string, object> one = new Dictionary<string, object>();
                one["para"] = i + 1;
                one["text"] = Head(now[i], 120);
                added.Add(one);
            }
        }
    }
}
