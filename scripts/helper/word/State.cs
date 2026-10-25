// v0.4 collaboration state: notice when the human changed the document under us, keep a journal
// of what the agent did, and write a handoff note next to the document.
//
//   handoff  {"doc":"...","out":"optional path","note":"optional free text"}
//
// A signature is cheap on purpose (a few dozen COM calls, not every paragraph): paragraph count,
// revision count, the saved flag, and the text of a sample of paragraphs.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

internal static partial class Program
{
    private static readonly string[] WRITE_COMMANDS = new string[] {
        "revise", "comment", "font", "format", "style", "page", "table", "caption", "picture",
        "toc", "reference", "insert", "cell", "fields", "save", "close", "revisions", "undo",
    };

    private static bool IsWriteCommand(string cmd)
    {
        foreach (string c in WRITE_COMMANDS) if (c == cmd) return true;
        return false;
    }

    private sealed class Signature
    {
        public int Paragraphs;
        public int Revisions;
        public bool Saved;
        public string Sample = "";

        public string Describe()
        {
            return Paragraphs + " \u6bb5 / " + Revisions + " \u5904\u4fee\u8ba2 / " + (Saved ? "\u5df2\u4fdd\u5b58" : "\u672a\u4fdd\u5b58");
        }

        public bool SameAs(Signature other)
        {
            if (other == null) return false;
            return Paragraphs == other.Paragraphs && Revisions == other.Revisions
                && Saved == other.Saved && Sample == other.Sample;
        }
    }

    private static string Sha16(string text)
    {
        using (SHA256 sha = SHA256.Create())
        {
            byte[] h = sha.ComputeHash(Encoding.UTF8.GetBytes(text ?? ""));
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < 8; i++) sb.Append(h[i].ToString("x2", CultureInfo.InvariantCulture));
            return sb.ToString();
        }
    }

    private static string StatePath(string docPath)
    {
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                  "wps-office-cooperate", "word-state");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, Sha16(docPath.ToLowerInvariant()) + ".json");
    }

    /// <summary>A cheap fingerprint of the open document: enough to notice a human edit, cheap enough to run twice per command.</summary>
    private static Signature TakeSignature(string docPath)
    {
        if (docPath == null) return null;
        dynamic doc = FindOpen(FullPathOf(docPath));
        if (doc == null) return null;

        Signature s = new Signature();
        s.Paragraphs = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        s.Revisions = (int)TryGet(() => (object)(int)doc.Revisions.Count, 0);
        s.Saved = (bool)TryGet(() => (object)(bool)doc.Saved, false);

        StringBuilder sample = new StringBuilder();
        int step = s.Paragraphs > 60 ? Math.Max(1, s.Paragraphs / 40) : 1;
        for (int i = 1; i <= s.Paragraphs; i += step)
        {
            string t = CleanPara(Convert.ToString(TryGet(() => (object)doc.Paragraphs.Item(i).Range.Text, ""), CultureInfo.InvariantCulture));
            sample.Append(i).Append(':').Append(t).Append('\n');
        }
        s.Sample = Sha16(sample.ToString());
        return s;
    }

    private static Dictionary<string, object> LoadState(string docPath)
    {
        try
        {
            string path = StatePath(docPath);
            if (!File.Exists(path)) return null;
            object parsed = Json.Parse(File.ReadAllText(path, Encoding.UTF8));
            return parsed as Dictionary<string, object>;
        }
        catch (Exception) { return null; }
    }

    private static void SaveState(string docPath, Signature now, string cmd)
    {
        Dictionary<string, object> prior = LoadState(docPath);
        List<object> journal = new List<object>();
        if (prior != null && prior.ContainsKey("journal"))
        {
            List<object> old = prior["journal"] as List<object>;
            if (old != null) journal.AddRange(old);
        }
        if (cmd != null)
        {
        Dictionary<string, object> entry = new Dictionary<string, object>();
        entry["at"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        entry["cmd"] = cmd;
        entry["paragraphs"] = now.Paragraphs;
        entry["revisions"] = now.Revisions;
        entry["saved"] = now.Saved;
        journal.Add(entry);
        while (journal.Count > 200) journal.RemoveAt(0);
        }

        Dictionary<string, object> state = new Dictionary<string, object>();
        state["path"] = docPath;
        state["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        state["paragraphs"] = now.Paragraphs;
        state["revisions"] = now.Revisions;
        state["saved"] = now.Saved;
        state["sample"] = now.Sample;
        state["journal"] = journal;
        File.WriteAllText(StatePath(docPath), Json.Write(state), new UTF8Encoding(false));
    }

    /// <summary>Called by Main around every write command: warns when the document moved under us, then journals this write.</summary>
    private static void AfterWrite(string cmd, string docPath, Signature before, Dictionary<string, object> response)
    {
        if (before == null || docPath == null) return;
        Dictionary<string, object> prior = LoadState(docPath);
        bool movedOutside = false;
        string priorDescription = null;

        if (prior != null)
        {
            Signature was = new Signature();
            was.Paragraphs = ToInt(prior, "paragraphs");
            was.Revisions = ToInt(prior, "revisions");
            was.Saved = prior.ContainsKey("saved") && prior["saved"] is bool && (bool)prior["saved"];
            was.Sample = prior.ContainsKey("sample") ? Convert.ToString(prior["sample"], CultureInfo.InvariantCulture) : "";
            if (!was.SameAs(before)) movedOutside = true;
            priorDescription = was.Describe();
        }

        Signature after = TakeSignature(docPath);
        // A command that changed nothing is not worth a journal line (repeated no-op formats, list calls).
        bool quiet = after != null && after.SameAs(before);
        if (after != null) SaveState(docPath, after, quiet ? null : cmd);

        if (!movedOutside) return;
        Dictionary<string, object> inner = null;
        object payload;
        if (response.TryGetValue("result", out payload)) inner = payload as Dictionary<string, object>;
        if (inner == null) return;
        inner["changedOutside"] = true;
        inner["changedOutsideHint"] = "Before this command the document was already different from what it was when this skill last wrote to it (then: "
            + priorDescription + "; before this command: " + before.Describe() + "). Someone may have edited it in the meantime: run `diff` to see what changed, `revisions` to see what is still pending, then continue.";
    }


    /// <summary>Shows the two strings around the first place they differ, so a three-character edit is visible in a table cell.</summary>
    private static void Delta(string before, string after, out string shownBefore, out string shownAfter)
    {
        before = before ?? "";
        after = after ?? "";
        int n = Math.Min(before.Length, after.Length);
        int i = 0;
        while (i < n && before[i] == after[i]) i++;
        int from = Math.Max(0, i - 20);
        shownBefore = (from > 0 ? "\u2026" : "") + MdCell(Slice(before, from, 60));
        shownAfter = (from > 0 ? "\u2026" : "") + MdCell(Slice(after, from, 60));
    }

    private static string Slice(string s, int from, int len)
    {
        if (s == null) return "";
        if (from >= s.Length) return "";
        return s.Substring(from, Math.Min(len, s.Length - from));
    }
    private static string MdCell(string s)
    {
        if (s == null) return "";
        s = s.Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        return s.Length <= 60 ? s : s.Substring(0, 60) + "\u2026";
    }
    /// <summary>
    /// handoff — a short markdown note next to the document: what this skill did, what is still pending,
    /// and whatever the caller wants the human to know. Deliberately plain text so a person reads it.
    /// </summary>
    private static object HandoffDoc(Bag a)
    {
        dynamic doc = RequireDoc(a, true);
        string full = Convert.ToString(TryGet(() => (object)doc.FullName, ""), CultureInfo.InvariantCulture);
        string note = a.Str("note", null);
        string outPath = a.Str("out", null);
        if (outPath == null || outPath.Trim().Length == 0)
        {
            string dir = Path.GetDirectoryName(full);
            string name = Path.GetFileNameWithoutExtension(full);
            outPath = Path.Combine(dir == null ? "." : dir, name + ".handoff.md");
        }
        else outPath = FullPathOf(outPath);

        int paragraphs = (int)TryGet(() => (object)(int)doc.Paragraphs.Count, 0);
        int pages = (int)TryGet(() => (object)(int)doc.ComputeStatistics(2), 0);
        int revisions = (int)TryGet(() => (object)(int)doc.Revisions.Count, 0);
        int comments = (int)TryGet(() => (object)(int)doc.Comments.Count, 0);
        bool saved = (bool)TryGet(() => (object)(bool)doc.Saved, true);

        Dictionary<string, object> prior = LoadState(full);
        List<object> journal = new List<object>();
        if (prior != null && prior.ContainsKey("journal") && prior["journal"] is List<object>)
            journal.AddRange((List<object>)prior["journal"]);

        // The newest local snapshot, if any: the next session continues the diff chain from there.
        string snapshotPath = null;
        try { snapshotPath = NewestSnapshot(SnapshotDirOf(doc)); } catch (Exception) { snapshotPath = null; }

        StringBuilder sb = new StringBuilder();
        sb.AppendLine("# \u4ea4\u63a5\u5355\uff1a" + Path.GetFileName(full));
        sb.AppendLine();
        sb.AppendLine("- \u6587\u6863\uff1a`" + full + "`");
        sb.AppendLine("- \u751f\u6210\u65f6\u95f4\uff1a" + DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture));
        sb.AppendLine("- \u89c4\u6a21\uff1a" + paragraphs + " \u6bb5 / " + pages + " \u9875");
        sb.AppendLine("- \u72b6\u6001\uff1a" + (saved ? "\u5df2\u4fdd\u5b58" : "**\u5c1a\u672a\u4fdd\u5b58**\uff08\u9700\u8981\u4f60 save\uff09"));
        sb.AppendLine("- \u5f85\u5904\u7406\u4fee\u8ba2\uff1a" + revisions + " \u5904" + (revisions > 0 ? "\uff08\u5728 WPS\u300c\u5ba1\u9605\u300d\u91cc\u9010\u6761\u63a5\u53d7/\u62d2\u7edd\uff09" : ""));
        sb.AppendLine("- \u6279\u6ce8\uff1a" + comments + " \u6761");
        sb.AppendLine("- \u6700\u65b0\u5feb\u7167\uff1a" + (snapshotPath == null ? "\u65e0\uff08\u63a5\u624b\u540e\u5efa\u8bae\u5148 snapshot \u4e00\u4efd\u518d\u52a8\u624b\uff09" : "`" + snapshotPath + "`\uff08diff \u7528 -Since <\u6b64\u8def\u5f84>\uff1b\u8de8\u673a\u5668\u63a5\u529b\u7528 snapshot -Out \u628a\u5b83\u590d\u5236\u5230\u6587\u6863\u65c1\u8fb9\uff09"));
        sb.AppendLine();

        if (note != null && note.Trim().Length > 0)
        {
            sb.AppendLine("## \u8fd9\u4e00\u8f6e\u505a\u4e86\u4ec0\u4e48");
            sb.AppendLine();
            sb.AppendLine(note.Trim());
            sb.AppendLine();
        }

        if (journal.Count > 0)
        {
            sb.AppendLine("## \u64cd\u4f5c\u8bb0\u5f55\uff08\u6700\u8fd1 " + Math.Min(30, journal.Count) + " \u6761\uff09");
            sb.AppendLine();
            sb.AppendLine("| \u65f6\u95f4 | \u547d\u4ee4 | \u6bb5\u6570 | \u4fee\u8ba2 | \u4fdd\u5b58 |");
            sb.AppendLine("|---|---|---|---|---|");
            int start = Math.Max(0, journal.Count - 30);
            for (int i = start; i < journal.Count; i++)
            {
                Dictionary<string, object> e = journal[i] as Dictionary<string, object>;
                if (e == null) continue;
                sb.AppendLine("| " + Convert.ToString(e.ContainsKey("at") ? e["at"] : "", CultureInfo.InvariantCulture)
                    + " | `" + Convert.ToString(e.ContainsKey("cmd") ? e["cmd"] : "", CultureInfo.InvariantCulture) + "`"
                    + " | " + ToInt(e, "paragraphs")
                    + " | " + ToInt(e, "revisions")
                    + " | " + ((e.ContainsKey("saved") && e["saved"] is bool && (bool)e["saved"]) ? "\u662f" : "\u5426") + " |");
            }
            sb.AppendLine();
        }

        // What changed since the last snapshot: the part a person cares about most when taking over.
        try
        {
            Bag db = new Bag(new Dictionary<string, object> { { "doc", full }, { "since", "last" } });
            Dictionary<string, object> dd = DiffDoc(db) as Dictionary<string, object>;
            if (dd != null && dd.ContainsKey("summary"))
            {
                Dictionary<string, object> sm = dd["summary"] as Dictionary<string, object>;
                int changed = sm == null ? 0 : ToInt(sm, "changed");
                int added = sm == null ? 0 : ToInt(sm, "added");
                int removed = sm == null ? 0 : ToInt(sm, "removed");
                sb.AppendLine("## \u81ea\u4e0a\u6b21\u5feb\u7167\u4ee5\u6765\u7684\u53d8\u5316\uff08\u6539 " + changed + " \u6bb5 / \u589e " + added + " / \u5220 " + removed + "\uff09");
                sb.AppendLine();
                List<object> rows = dd.ContainsKey("changed") ? dd["changed"] as List<object> : null;
                if (rows != null && rows.Count > 0)
                {
                    sb.AppendLine("| \u6bb5\u843d | \u6539\u524d | \u6539\u540e |");
                    sb.AppendLine("|---|---|---|");
                    for (int i = 0; i < rows.Count && i < 10; i++)
                    {
                        Dictionary<string, object> row = rows[i] as Dictionary<string, object>;
                        if (row == null) continue;
                        string bshow, ashow;
                        Delta(Convert.ToString(row.ContainsKey("before") ? row["before"] : "", CultureInfo.InvariantCulture),
                              Convert.ToString(row.ContainsKey("after") ? row["after"] : "", CultureInfo.InvariantCulture), out bshow, out ashow);
                        sb.AppendLine("| " + ToInt(row, "para") + " | " + bshow + " | " + ashow + " |");
                    }
                    if (rows.Count > 10) sb.AppendLine("| \u2026 | \u5171 " + rows.Count + " \u6bb5\uff0c\u53ea\u5217 10 \u6bb5 | |");
                    sb.AppendLine();
                }
            }
        }
        catch (Exception) { }
        sb.AppendLine("## \u63a5\u624b\u65f6\u5efa\u8bae\u5148\u505a");
        sb.AppendLine();
        sb.AppendLine("1. `status` \u770b\u8fd9\u4efd\u6587\u6863\u662f\u4e0d\u662f\u8fd8\u5f00\u7740\uff1b");
        sb.AppendLine("2. `revisions` \u770b\u8fd8\u6709\u591a\u5c11\u4fee\u8ba2\u6ca1\u88ab\u5904\u7406\uff1b");
        sb.AppendLine("3. `diff` \u770b\u4e0e\u4e0a\u6b21\u5feb\u7167\u76f8\u6bd4\u4f60\u6539\u4e86\u54ea\u4e9b\u6bb5\uff1b");
        sb.AppendLine("4. `render` \u62bd\u51e0\u9875\u770b\u7248\u5f0f\uff0c\u518d\u51b3\u5b9a\u4e0b\u4e00\u6b65\u3002");
        sb.AppendLine();

        string dirOut = Path.GetDirectoryName(outPath);
        if (dirOut != null && dirOut.Length > 0 && !Directory.Exists(dirOut)) Directory.CreateDirectory(dirOut);
        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(false));

        Dictionary<string, object> r = new Dictionary<string, object>();
        r["name"] = Convert.ToString(TryGet(() => (object)doc.Name, ""), CultureInfo.InvariantCulture);
        r["path"] = outPath;
        r["journalEntries"] = journal.Count;
        r["snapshotPath"] = snapshotPath;
        r["revisions"] = revisions;
        r["comments"] = comments;
        r["saved"] = false;
        r["note"] = "\u4ea4\u63a5\u5355\u5199\u5728\u6587\u6863\u65c1\u8fb9\uff0c\u4eba\u6253\u5f00\u5c31\u80fd\u770b\u3002\u5b83\u4e0d\u6539\u6587\u6863\u672c\u8eab\u3002";
        return r;
    }
}
