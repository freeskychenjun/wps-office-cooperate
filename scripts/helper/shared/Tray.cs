// wps-office-cooperate: resident tray indicator for a takeover hint (v1.12).
//
// Why a tray icon instead of a floating window: a top-most overlay would cover the document
// the user is reading, fight WPS for focus and need DPI/multi-monitor/cleanup handling. A tray
// icon is always available, never covers content, and carries the same "who is editing" answer
// on hover plus a right-click action. See SKILL.md v1.12 for the design note.
//
// Lifecycle: this file is compiled into EACH app helper (word/excel/ppt). `hint -Tray` makes
// the short-lived helper write a state file and spawn a detached copy of ITSELF with
// `--tray <state.json>`; that process shows the icon, watches the state file (tooltip/progress
// updates + the stop flag) and exits on its own. The restore menu entry runs the skill's own
// wrapper (`hint -Clear`) so the daemon needs no COM knowledge - it stays app-agnostic.
//
// Pure ASCII source: Chinese lives in \u escapes, same rule as the rest of the helpers.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

internal static class TrayDaemon
{
    private const int MaxTip = 63;                 // NotifyIcon.Text hard limit

    // ---------------------------------------------------------------- daemon

    /// <summary>Entry point for the detached `--tray <state.json>` process.</summary>
    public static int Run(string statePath)
    {
        try { Application.EnableVisualStyles(); }
        catch (Exception) { }
        Dictionary<string, object> state = LoadState(statePath);
        if (state == null) return 0;
        if (GetBool(state, "stop")) return 0;

        NotifyIcon ni = new NotifyIcon();
        ContextMenuStrip menu = new ContextMenuStrip();
        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        try
        {
            ni.Icon = BuildIcon();
            ni.Text = Tip(state);
            ni.Visible = true;
            ToolStripItem restore = menu.Items.Add("\u7ed3\u675f\u534f\u4f5c\uff1a\u8fd8\u539f\u6807\u9898\u6807\u8bb0");  // 结束协作：还原标题标记
            restore.Click += delegate { Restore(state, ni, menu); };
            ToolStripItem drop = menu.Items.Add("\u4ec5\u5173\u95ed\u6258\u76d8\u56fe\u6807");                            // 仅关闭托盘图标
            drop.Click += delegate { Drop(ni, menu); };
            ni.ContextMenuStrip = menu;
            ni.DoubleClick += delegate { Restore(state, ni, menu); };

            string lastSig = Sig(state);
            int tick = 0;
            int notInUse = 0;
            timer.Interval = 800;
            timer.Tick += delegate
            {
                Dictionary<string, object> now = LoadState(statePath);
                if (now == null || GetBool(now, "stop"))
                {
                    timer.Stop();
                    Drop(ni, menu);
                    return;
                }
                string sig = Sig(now);
                if (sig != lastSig)
                {
                    lastSig = sig;
                    state = now;
                    ni.Text = Tip(now);
                }
                long updated = GetLong(now, "updated");
                int maxMinutes = GetInt(now, "maxMinutes", 240);
                if (maxMinutes > 0 && updated > 0 && NowMs() - updated > (long)maxMinutes * 60L * 1000L)
                {
                    timer.Stop();
                    Drop(ni, menu);
                    return;
                }
                // The user may close the document while the icon is up. Ask the OS (Restart
                // Manager) whether any process still has the file open: two consecutive
                // "not in use" checks (~10 s) retire the icon, so a transient save/rename does
                // not flicker it away (measured 2026-10-30: IN_USE=1 while WPS has the doc,
                // 0 right after the user closes it).
                tick++;
                if (tick % 6 == 0)      // every ~4.8 s
                {
                    string doc = GetStr(now, "doc");
                    if (doc.Length > 0 && File.Exists(doc))
                    {
                        string inUse = FileInUse(doc);          // "1" / "0" / "?" (unknown)
                        if (inUse == "0") notInUse++;
                        else notInUse = 0;
                        if (notInUse >= 2)
                        {
                            timer.Stop();
                            Drop(ni, menu);
                        }
                    }
                    else notInUse = 0;                          // missing path: do not guess
                }
            };
            timer.Start();
            Application.Run();
        }
        catch (Exception) { }
        finally
        {
            try { timer.Dispose(); } catch (Exception) { }
            try { menu.Dispose(); } catch (Exception) { }
            try { ni.Visible = false; ni.Dispose(); } catch (Exception) { }
        }
        return 0;
    }

    private static void Restore(Dictionary<string, object> state, NotifyIcon ni, ContextMenuStrip menu)
    {
        try
        {
            ni.Visible = false;
            List<string> cmd = GetStrings(state, "restore");
            if (cmd.Count > 0)
            {
                ProcessStartInfo psi = new ProcessStartInfo(cmd[0]);
                StringBuilder args = new StringBuilder();
                for (int i = 1; i < cmd.Count; i++)
                {
                    if (args.Length > 0) args.Append(' ');
                    args.Append(Quote(cmd[i]));
                }
                psi.Arguments = args.ToString();
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                try
                {
                    using (Process p = Process.Start(psi))
                    {
                        if (p != null) p.WaitForExit(30000);
                    }
                }
                catch (Exception) { }
            }
        }
        catch (Exception) { }
        Drop(ni, menu);
    }

    private static void Drop(NotifyIcon ni, ContextMenuStrip menu)
    {
        try { ni.Visible = false; } catch (Exception) { }
        try { Application.ExitThread(); } catch (Exception) { }
    }

    // ---------------------------------------------------------------- start / stop (used by hint)

    public static string DefaultStateDir()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate");
    }

    private static string StatePath(string app)
    {
        return Path.Combine(DefaultStateDir(), "tray-" + app + ".json");
    }

    private static string PidPath(string app)
    {
        return Path.Combine(DefaultStateDir(), "tray-" + app + ".pid");
    }

    /// <summary>
    /// Write the tray state and make sure exactly one daemon is showing it. Reuses a live
    /// daemon (it polls the file, so the tooltip just updates); otherwise spawns a detached
    /// copy of this executable in --tray mode.
    /// </summary>
    public static Dictionary<string, object> Start(string app, string doc, string caption, string text, string wrapper)
    {
        Dictionary<string, object> r = new Dictionary<string, object>();
        try
        {
            Directory.CreateDirectory(DefaultStateDir());
            string statePath = StatePath(app);
            string pidPath = PidPath(app);

            Dictionary<string, object> state = new Dictionary<string, object>();
            state["app"] = app;
            state["doc"] = doc;
            state["caption"] = caption;
            state["text"] = text;
            state["updated"] = NowMs();
            state["maxMinutes"] = 240;      // a forgotten icon cleans itself up after 4 h
            state["stop"] = false;
            List<object> restore = new List<object>();
            if (wrapper != null && wrapper.Length > 0)
            {
                restore.Add("powershell");
                restore.Add("-NoProfile");
                restore.Add("-ExecutionPolicy");
                restore.Add("Bypass");
                restore.Add("-File");
                restore.Add(wrapper);
                restore.Add("hint");
                restore.Add("-Doc");
                restore.Add(doc);
                restore.Add("-Clear");
                restore.Add("-Quiet");
            }
            state["restore"] = restore;
            WriteState(statePath, state);

            int existing = ReadPid(pidPath);
            if (existing > 0 && Alive(existing))
            {
                r["started"] = true;
                r["reused"] = true;
                r["pid"] = existing;
                r["stateFile"] = statePath;
                return r;
            }

            string self = Process.GetCurrentProcess().MainModule.FileName;
            // UseShellExecute=true is the point: it starts the daemon WITHOUT inheriting our
            // stdio handles. With UseShellExecute=false the daemon would hold the caller's
            // stdout/stderr pipe open, and any caller that reads until EOF (python
            // subprocess.communicate, PowerShell pipelines) would block for the icon's whole
            // 4-hour life instead of returning after the hint (measured 2026-10-30).
            ProcessStartInfo psi = new ProcessStartInfo(self, "--tray \"" + statePath + "\"");
            psi.UseShellExecute = true;
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            Process p = Process.Start(psi);
            int pid = p == null ? 0 : p.Id;
            try { File.WriteAllText(pidPath, pid.ToString(CultureInfo.InvariantCulture), new UTF8Encoding(false)); }
            catch (Exception) { }
            Thread.Sleep(300);              // let the message loop come up before the caller reports success
            r["started"] = pid > 0;
            r["pid"] = pid;
            r["stateFile"] = statePath;
            // Windows parks every new tray icon in the "show hidden icons" (^) overflow by
            // default and no process can promote itself; say it instead of pretending the icon
            // is on the taskbar. The user drags it out once and it stays.
            r["note"] = "Windows puts a new tray icon in the \"^\" overflow area by default."
                + " Tell the user they can hover it there for the current step, or drag it onto the taskbar to keep it visible."
                + " Right-click restores the title marker; finish/close removes the icon.";
        }
        catch (Exception e)
        {
            r["started"] = false;
            r["error"] = e.Message;
        }
        return r;
    }

    /// <summary>Ask the daemon to exit (it polls the stop flag). Safe when no icon is running.</summary>
    public static bool Stop(string app)
    {
        try
        {
            string statePath = StatePath(app);
            if (!File.Exists(statePath)) return false;
            Dictionary<string, object> state = LoadState(statePath);
            if (state == null)
            {
                try { File.Delete(statePath); } catch (Exception) { }
                return false;
            }
            state["stop"] = true;
            state["updated"] = NowMs();
            WriteState(statePath, state);
            try { File.Delete(PidPath(app)); } catch (Exception) { }
            return true;
        }
        catch (Exception) { return false; }
    }

    // ---------------------------------------------------------------- plumbing

    /// <summary>
    /// Point the running icon at a different file (save -As renames the document). Keeps the
    /// current caption/text; no-op when no tray state exists.
    /// </summary>
    public static bool UpdateDoc(string app, string doc, string caption)
    {
        try
        {
            string statePath = StatePath(app);
            if (!File.Exists(statePath)) return false;
            Dictionary<string, object> state = LoadState(statePath);
            if (state == null) return false;
            state["doc"] = doc;
            if (caption != null && caption.Length > 0) state["caption"] = caption;
            state["updated"] = NowMs();
            WriteState(statePath, state);
            return true;
        }
        catch (Exception) { return false; }
    }

    // ---------------------------------------------------------------- file-in-use (Restart Manager)
    // "Did the user close the document?" - the authoritative answer is the OS: Restart Manager
    // enumerates the processes holding a handle to the file (it powers the "file is in use"
    // dialog). Measured 2026-10-30 on this machine: IN_USE=1 (WPS Office) while the doc is open,
    // IN_USE=0 within a second of closing it. The daemon polls it so a closed document does not
    // leave a stale "AI is editing" icon behind.

    [StructLayout(LayoutKind.Sequential)]
    private struct RmUniqueProcess
    {
        public int dwProcessId;
        public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct RmProcessInfo
    {
        public RmUniqueProcess Process;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string strAppName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string strServiceShortName;
        public int ApplicationType;
        public uint AppStatus;
        public uint TSSessionId;
        [MarshalAs(UnmanagedType.Bool)] public bool bRestartable;
    }

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, string strSessionKey);
    [DllImport("rstrtmgr.dll")]
    private static extern int RmEndSession(uint pSessionHandle);
    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmRegisterResources(uint pSessionHandle, uint nFiles, string[] rgsFilenames,
                                                 uint nApplications, IntPtr rgApplications, uint nServices, string[] rgsServiceNames);
    [DllImport("rstrtmgr.dll")]
    private static extern int RmGetList(uint dwSessionHandle, out uint pnProcInfoNeeded, ref uint pnProcInfo,
                                        [In, Out] RmProcessInfo[] rgAffectedApps, ref uint lpdwRebootReasons);

    /// <summary>Returns "1" (open by some process), "0" (nobody holds it), or "?" (unknown).</summary>
    private static string FileInUse(string path)
    {
        uint handle;
        try
        {
            int rc = RmStartSession(out handle, 0, Guid.NewGuid().ToString());
            if (rc != 0) return "?";
            try
            {
                rc = RmRegisterResources(handle, 1, new string[] { path }, 0, IntPtr.Zero, 0, null);
                if (rc != 0) return "?";
                uint needed = 0, count = 0, reasons = 0;
                RmProcessInfo[] infos = new RmProcessInfo[0];
                rc = RmGetList(handle, out needed, ref count, infos, ref reasons);
                if (rc == 2 /*ERROR_MORE_DATA*/ || needed > 0)
                {
                    infos = new RmProcessInfo[needed];
                    count = needed;
                    rc = RmGetList(handle, out needed, ref count, infos, ref reasons);
                    if (rc != 0) return "?";
                }
                return count == 0 ? "0" : "1";
            }
            finally { RmEndSession(handle); }
        }
        catch (Exception) { return "?"; }
    }

    private static Dictionary<string, object> LoadState(string path)
    {
        try
        {
            if (!File.Exists(path)) return null;
            object o = Json.Parse(File.ReadAllText(path, Encoding.UTF8));
            return o as Dictionary<string, object>;
        }
        catch (Exception) { return null; }
    }

    private static void WriteState(string path, Dictionary<string, object> state)
    {
        string tmp = path + ".tmp";
        File.WriteAllText(tmp, Json.Write(state), new UTF8Encoding(false));
        try { File.Delete(path); } catch (Exception) { }
        File.Move(tmp, path);
    }

    private static int ReadPid(string pidPath)
    {
        try
        {
            if (!File.Exists(pidPath)) return 0;
            int pid;
            if (int.TryParse(File.ReadAllText(pidPath, Encoding.UTF8).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out pid)) return pid;
        }
        catch (Exception) { }
        return 0;
    }

    private static bool Alive(int pid)
    {
        try
        {
            Process p = Process.GetProcessById(pid);
            return p != null && !p.HasExited;
        }
        catch (Exception) { return false; }
    }

    private static string Sig(Dictionary<string, object> state)
    {
        return GetStr(state, "caption") + "\u0001" + GetStr(state, "text") + "\u0001" + GetLong(state, "updated").ToString(CultureInfo.InvariantCulture);
    }

    private static string Tip(Dictionary<string, object> state)
    {
        string caption = GetStr(state, "caption");
        string text = GetStr(state, "text");
        string tip = caption;
        if (text != null && text.Length > 0) tip = (caption.Length > 0 ? caption + " \u00b7 " : "") + text;
        if (tip == null) tip = "";
        tip = tip.Replace("\r", " ").Replace("\n", " ").Trim();
        if (tip.Length == 0) tip = "AI";
        if (tip.Length > MaxTip) tip = tip.Substring(0, MaxTip - 1) + "\u2026";
        return tip;
    }

    /// <summary>Small orange round "AI" badge so the icon is recognisable in the tray.</summary>
    private static Icon BuildIcon()
    {
        Bitmap bmp = new Bitmap(32, 32);
        try
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.AntiAlias;
                g.Clear(Color.Transparent);
                using (Brush fill = new SolidBrush(Color.FromArgb(0xD9, 0x77, 0x06)))
                    g.FillEllipse(fill, 1, 1, 30, 30);
                using (Font f = new Font("Segoe UI", 13f, FontStyle.Bold, GraphicsUnit.Pixel))
                using (Brush fg = new SolidBrush(Color.White))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    g.DrawString("AI", f, fg, new RectangleF(0f, 1f, 32f, 32f), sf);
                }
            }
            IntPtr h = bmp.GetHicon();
            return Icon.FromHandle(h);
        }
        finally
        {
            try { bmp.Dispose(); } catch (Exception) { }
        }
    }

    private static string Quote(string s)
    {
        if (s == null) return "\"\"";
        if (s.IndexOf(' ') < 0 && s.IndexOf('"') < 0) return s;
        return "\"" + s.Replace("\"", "\\\"") + "\"";
    }

    private static long NowMs()
    {
        return (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
    }

    private static string GetStr(Dictionary<string, object> d, string key)
    {
        object v;
        if (d != null && d.TryGetValue(key, out v) && v != null) return Convert.ToString(v, CultureInfo.InvariantCulture);
        return "";
    }

    private static long GetLong(Dictionary<string, object> d, string key)
    {
        object v;
        if (d != null && d.TryGetValue(key, out v) && v != null)
        {
            long n;
            if (v is long) return (long)v;
            if (v is int) return (long)(int)v;
            if (v is double) return (long)(double)v;
            if (long.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out n)) return n;
        }
        return 0;
    }

    private static int GetInt(Dictionary<string, object> d, string key, int fallback)
    {
        long v = GetLong(d, key);
        return v == 0 ? fallback : (int)v;
    }

    private static bool GetBool(Dictionary<string, object> d, string key)
    {
        object v;
        if (d != null && d.TryGetValue(key, out v) && v != null)
        {
            if (v is bool) return (bool)v;
            bool b;
            if (bool.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), out b)) return b;
        }
        return false;
    }

    private static List<string> GetStrings(Dictionary<string, object> d, string key)
    {
        List<string> list = new List<string>();
        object v;
        if (d == null || !d.TryGetValue(key, out v) || v == null) return list;
        IList<object> arr = v as IList<object>;
        if (arr != null)
        {
            foreach (object o in arr) list.Add(Convert.ToString(o, CultureInfo.InvariantCulture));
            return list;
        }
        object[] objs = v as object[];
        if (objs != null)
        {
            foreach (object o in objs) list.Add(Convert.ToString(o, CultureInfo.InvariantCulture));
            return list;
        }
        string single = Convert.ToString(v, CultureInfo.InvariantCulture);
        if (single != null && single.Length > 0) list.Add(single);
        return list;
    }
}
