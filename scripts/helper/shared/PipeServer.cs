// v1.14 - resident helper: one long-lived process per app, so a command does not pay
// "start .NET exe + attach to WPS + find the document" again and again.
//
// Measured floor on this machine (2026-11-01): a warm `status` cost ~4.8 s, of which ~2.0 s is
// the caller's own PowerShell start, ~2.3 s the wrapper's source-hash + exe start + COM attach,
// and the rest the actual work. The daemon keeps the COM application object alive and serves
// requests over a named pipe, which removes the exe start and the COM attach from every call.
//
// Protocol (both directions): 4-byte little-endian length + UTF-8 JSON. One request per
// connection, one response per request, strictly serialized (WPS automation is single-instance
// anyway, and COM must be used from the thread that created the objects ⇒ the server loop runs on
// the [STAThread] Main thread).
//
// The daemon is intentionally dumb about the API: it receives the SAME request JSON the one-shot
// mode receives and returns the SAME response JSON, so "daemon vs one-shot" is a pure transport
// difference and can be A/B tested byte for byte.

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;

internal static class PipeServer
{
    private const int MaxMessage = 32 * 1024 * 1024;

    /// <summary>
    /// Serve requests until idle, until no document is left open, or until the pipe breaks.
    /// Returns the process exit code.
    ///
    /// v1.35: this is also where the takeover-hint LEASE is enforced. The hint must not outlive
    /// the agent (see shared/HintLease.cs), so the helper renews a lease file per request and,
    /// while waiting in short slices, calls releaseHints() once the lease has expired. The same
    /// release runs on the way out (idle timeout / `serve -Action stop` / last document closed),
    /// so a branded window cannot survive the process either.
    ///
    /// v1.36: an expired lease ends the process too - "the agent went quiet" is the same thing as
    /// "no client should stay attached to WPS". The wrapper restarts a helper on the next command.
    /// </summary>
    public static int Run(string pipeName, string app, Func<string, string> handler, Func<int> openDocuments, Action releaseHints)
    {
        int idleMs = 20 * 60 * 1000;
        string env = Environment.GetEnvironmentVariable("WOC_DAEMON_IDLE_MS");
        int parsed;
        if (env != null && int.TryParse(env, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) && parsed > 2000)
            idleMs = parsed;
        string once = Environment.GetEnvironmentVariable("WOC_DAEMON_ONCE");
        bool exitAfterOne = once == "1";
        int pollMs = HintLease.PollMs();

        Console.Error.WriteLine("woc-daemon: listening on " + pipeName + " (idle " + idleMs
            + " ms, hint lease " + HintLease.LeaseMs() + " ms)");
        bool hadDocument = false;
        bool released = false;
        DateTime lastActivity = DateTime.UtcNow;
        HintLease.Touch(app);
        try
        {
        while (true)
        {
            string request;
            using (NamedPipeServerStream server = new NamedPipeServerStream(
                pipeName, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous))
            {
                // NB: EndWaitForConnection must get the SAME IAsyncResult - passing null makes it
                // throw, the stream gets disposed, and the client sees a broken pipe right after
                // connecting (that bug cost one debugging round).
                IAsyncResult connecting = server.BeginWaitForConnection(null, null);
                // Wait in slices: a request wakes us early, and every slice re-checks the lease so
                // an agent that went away is noticed without waiting for the 20 min idle timeout.
                while (true)
                {
                    int idleLeft = idleMs - (int)(DateTime.UtcNow - lastActivity).TotalMilliseconds;
                    if (idleLeft <= 0)
                    {
                        Console.Error.WriteLine("woc-daemon: idle timeout, exiting");
                        return 0;
                    }
                    int slice = Math.Min(idleLeft, pollMs);
                    if (connecting.AsyncWaitHandle.WaitOne(slice)) break;
                    if (!released && HintLease.Stale(app))
                    {
                        // Release ONCE per lease period: the intent registry stays, so the next
                        // request re-applies the hint and a resumed session is branded again.
                        //
                        // v1.36: release AND leave. The expired lease means the agent's session is
                        // over, so keeping this process - i.e. keeping an automation client
                        // attached to the user's WPS - alive for another 20 idle minutes only adds
                        // one more participant to the user's own save/close path (WPS 12.1 on this
                        // machine already crashes there on its own, see SKILL.md v1.36). The
                        // wrapper starts a fresh helper the next time an agent actually runs a
                        // command, and the kept intent registry re-brands at that moment.
                        released = true;
                        Console.Error.WriteLine("woc-daemon: hint lease expired (no request for "
                            + HintLease.LeaseMs() + " ms) - releasing takeover hints and exiting");
                        try { releaseHints(); }
                        catch (Exception) { Console.Error.WriteLine("woc-daemon: hint release failed"); }
                        return 0;
                    }
                }
                try { server.EndWaitForConnection(connecting); }
                catch (Exception e)
                {
                    Console.Error.WriteLine("woc-daemon: accept failed: " + e.Message);
                    continue;
                }
                try
                {
                    request = ReadMessage(server);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine("woc-daemon: bad request frame: " + e.Message);
                    continue;
                }
                string response;
                if (request == null || request.Length == 0)
                {
                    response = "{\"ok\":false,\"error\":{\"code\":\"BAD_ARGS\",\"message\":\"empty request\"}}";
                }
                else
                {
                    try { response = handler(request); }
                    catch (Exception e) { response = "{\"ok\":false,\"error\":{\"code\":\"INTERNAL\",\"message\":\"" + Escape(e.Message) + "\"}}"; }
                }
                try
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(response ?? "");
                    byte[] len = BitConverter.GetBytes(bytes.Length);
                    server.Write(len, 0, 4);
                    server.Write(bytes, 0, bytes.Length);
                    server.Flush();
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine("woc-daemon: response write failed: " + e.Message);
                }
            }

            lastActivity = DateTime.UtcNow;
            HintLease.Touch(app);   // a served request renews the hint lease
            released = false;
            try { if (openDocuments() > 0) hadDocument = true; }
            catch (Exception) { }
            if (exitAfterOne) return 0;
            if (hadDocument)
            {
                int left = -1;
                try { left = openDocuments(); } catch (Exception) { }
                if (left == 0)
                {
                    Console.Error.WriteLine("woc-daemon: no documents left, exiting");
                    return 0;
                }
            }
        }
        }
        finally
        {
            // Any exit path (idle timeout, serve -Action stop, last document closed, a crash in
            // this thread) must not leave a window branded "AI is working". The lease path above
            // has already released before returning, so do not touch COM on that exit a second time
            // (the app may be half-dead exactly then).
            if (!released)
            {
                try { releaseHints(); } catch (Exception) { }
            }
        }
    }

    private static string ReadMessage(Stream s)
    {
        byte[] len = new byte[4];
        if (!ReadExactly(s, len, 4)) return null;
        int n = BitConverter.ToInt32(len, 0);
        if (n <= 0 || n > MaxMessage) return null;
        byte[] buf = new byte[n];
        if (!ReadExactly(s, buf, n)) return null;
        return Encoding.UTF8.GetString(buf);
    }

    private static bool ReadExactly(Stream s, byte[] buf, int count)
    {
        int got = 0;
        while (got < count)
        {
            int r = s.Read(buf, got, count - got);
            if (r <= 0) return false;
            got += r;
        }
        return true;
    }

    private static string Escape(string s)
    {
        if (s == null) return "";
        // Full JSON string escaping: the old version only handled \\, " and CR/LF, so an exception
        // message carrying a tab or another control character produced a frame the client could not
        // parse (the client would then fall back to the one-shot path and re-run the command).
        StringBuilder b = new StringBuilder(s.Length + 8);
        foreach (char c in s)
        {
            switch (c)
            {
                case '\\': b.Append("\\\\"); break;
                case '"': b.Append("\\\""); break;
                case '\r': b.Append("\\r"); break;
                case '\n': b.Append("\\n"); break;
                case '\t': b.Append("\\t"); break;
                case '\b': b.Append("\\b"); break;
                case '\f': b.Append("\\f"); break;
                default:
                    if (c < 0x20) b.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                    else b.Append(c);
                    break;
            }
        }
        return b.ToString();
    }
}
