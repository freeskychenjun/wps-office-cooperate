// v1.32 - busy retry: the OLE message filter an automation client is supposed to install.
//
// Why this exists: WPS rejects an incoming COM call while the user is typing in a cell, has a
// modal dialog open, or the app is repainting. Without a filter our caller gets that rejection at
// once (0x80010001 RPC_E_CALL_REJECTED, or 0x8001010A RPC_E_SERVERCALL_RETRYFLT) and the command
// fails while the user is in the middle of a keystroke. With a filter the OLE channel is told
// "retry in 150 ms" for up to 6 s, so a short busy window is ridden out instead of turned into an
// error. When 6 s is not enough the call is cancelled on purpose: the COMException then reaches the
// existing error mapping in HandleRequest and comes back as BUSY, which the caller can retry.
//
// Policy ported verbatim from dsh-office (helper/OfficeHelper.cs:115-135, MIT):
//   HandleInComingCall -> 0    (SERVERCALL_ISHANDLED)
//   RetryRejectedCall  -> dwRejectType == 2 (SERVERCALL_RETRYLATER) and dwTickCount < 6000
//                         ? 150 (retry after 150 ms) : -1 (cancel the call)
//   MessagePending     -> 2    (PENDINGMSG_WAITDEFPROCESS)
//
// One registration covers every command: the helpers are [STAThread] and every request runs on the
// Main thread (see PipeServer.cs), and a message filter belongs to the apartment, not to a call.
// Register it before the first COM call - Main does that as its first statement.

using System;
using System.Globalization;
using System.Runtime.InteropServices;

/// <summary>The OLE message filter interface (oleidl.h: IOleMessageFilter, IID 00000016-...).</summary>
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

internal sealed class MessageFilter : IOleMessageFilter
{
    /// <summary>How long to ask OLE to wait before retrying a rejected call (milliseconds).</summary>
    public const int RetryAfterMs = 150;

    /// <summary>How much rejection time to sit through before letting the call fail (milliseconds).</summary>
    public const int GiveUpAfterMs = 6000;

    private const int ServerCallRetryLater = 2;   // SERVERCALL_RETRYLATER
    private const int PendingWaitDefault = 2;     // PENDINGMSG_WAITDEFPROCESS
    private const int ServerCallHandled = 0;      // SERVERCALL_ISHANDLED

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

    /// <summary>Keeps the managed side of the registered filter alive: COM holds a reference to the
    /// callable wrapper, and a collected filter would leave that reference dangling.</summary>
    private static MessageFilter _installed;

    /// <summary>HRESULT of the last registration (0 = S_OK); kept for diagnostics.</summary>
    public static int LastHResult = -1;

    public static bool Installed { get { return _installed != null; } }

    /// <summary>Install the filter on the current (STA) thread. Cheap and idempotent; call it once
    /// in Main, before any COM object is created or attached.</summary>
    public static int Register()
    {
        MessageFilter filter = new MessageFilter();
        IOleMessageFilter old = null;
        int hr;
        try { hr = CoRegisterMessageFilter(filter, out old); }
        catch (Exception e)
        {
            LastHResult = -1;
            Say("register failed: " + e.GetType().Name + ": " + e.Message);
            return -1;
        }
        _installed = filter;
        LastHResult = hr;
        Say("CoRegisterMessageFilter -> 0x" + hr.ToString("X8", CultureInfo.InvariantCulture)
            + " (replaced=" + (old != null ? "yes" : "no") + ")");
        return hr;
    }

    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
    {
        return ServerCallHandled;
    }

    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        if (dwRejectType == ServerCallRetryLater && dwTickCount < GiveUpAfterMs) return RetryAfterMs;
        return -1;
    }

    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
    {
        return PendingWaitDefault;
    }

    /// <summary>WOC_MSGFLT_DEBUG=1 prints the registration HRESULT on stderr (off by default, so a
    /// normal run stays silent - the three wrappers promise one JSON document on stdout).</summary>
    private static void Say(string text)
    {
        if (Environment.GetEnvironmentVariable("WOC_MSGFLT_DEBUG") != "1") return;
        try { Console.Error.WriteLine("woc-msgfilter: " + text); }
        catch (Exception) { }
    }
}
