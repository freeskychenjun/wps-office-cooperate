// wps-office-cooperate: the takeover hint is a LEASE, not a permanent mark (v1.35).
//
// The bug this fixes (reported 2026-11-05): `open` brands the document window "AI is working"
// (Word/Excel caption, PPT frame title, tray icon). Every way to remove that brand used to be
// agent-driven (`finish`, `close`, `hint -Clear`), so an agent that simply STOPS - the user ends
// the chat, the model errors out, the session is killed - left the tab looking permanently taken
// over. Measured on this machine: a deck opened at 16:52 still showed "〔AI接管中〕" at 17:30,
// with the resident helper idle and no agent in sight.
//
// A lease fixes that generically, for every agent (the skill must not depend on any one host):
//
//   * the resident helper renews `hint-lease-<app>.txt` when it starts and after every request;
//   * while it waits for the next request it polls that file, and once the lease is older than
//     WOC_HINT_LEASE_MS (default 120 s) it removes the VISUAL hint from every document that
//     still carries one (see each helper's ReleaseHints);
//   * the helper also releases on its way out (20 min idle / `serve -Action stop`), so even a
//     lost lease cannot leave a branded window behind.
//
// What is deliberately NOT cleared is the INTENT registry (hint-state-<app>.json / PPT's
// ppt-title-hint.json): the next request re-applies the hint, so a session that pauses longer
// than the lease comes back to a correctly branded window instead of a silently unbranded one.
//
// One-shot mode (WOC_NO_DAEMON=1) has no resident process to run the timer; there the hint is
// still only removed by finish / close / hint -Clear (documented).

using System;
using System.Globalization;
using System.IO;
using System.Text;

internal static class HintLease
{
    /// <summary>How long after the last request the hint may stay on screen.</summary>
    private const long DefaultLeaseMs = 120 * 1000;

    /// <summary>Upper bound on a single wait before the lease is re-checked.</summary>
    private const int MaxPollMs = 5000;

    public static long LeaseMs()
    {
        string env = Environment.GetEnvironmentVariable("WOC_HINT_LEASE_MS");
        long v;
        if (env != null && long.TryParse(env, NumberStyles.Integer, CultureInfo.InvariantCulture, out v) && v >= 2000) return v;
        return DefaultLeaseMs;
    }

    /// <summary>Poll cadence: half the lease (so the mark is gone well within one lease period),
    /// clamped to 1-5 s so a test can use a 2-3 s lease without spinning.</summary>
    public static int PollMs()
    {
        long poll = LeaseMs() / 2;
        if (poll < 1000) poll = 1000;
        if (poll > MaxPollMs) poll = MaxPollMs;
        return (int)poll;
    }

    public static string StateDir()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wps-office-cooperate");
    }

    private static string PathFor(string app)
    {
        return Path.Combine(StateDir(), "hint-lease-" + app + ".txt");
    }

    /// <summary>Renew the lease. Called by the resident helper (start + after every request);
    /// failures are ignored - a missing lease only means "the watchdog has nothing to expire".</summary>
    public static void Touch(string app)
    {
        try
        {
            Directory.CreateDirectory(StateDir());
            File.WriteAllText(PathFor(app), NowMs().ToString(CultureInfo.InvariantCulture), new UTF8Encoding(false));
        }
        catch (Exception) { }
    }

    /// <summary>Epoch ms of the last renewal, or 0 when there is none/unreadable.</summary>
    public static long LastSeen(string app)
    {
        try
        {
            string p = PathFor(app);
            if (!File.Exists(p)) return 0;
            long v;
            if (long.TryParse(File.ReadAllText(p, Encoding.UTF8).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out v)) return v;
        }
        catch (Exception) { }
        return 0;
    }

    /// <summary>Has the agent been silent for longer than the lease? False when no lease exists
    /// (one-shot mode, or a helper that never started a daemon).</summary>
    public static bool Stale(string app)
    {
        long last = LastSeen(app);
        if (last <= 0) return false;
        return NowMs() - last > LeaseMs();
    }

    private static long NowMs()
    {
        return (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
    }
}
