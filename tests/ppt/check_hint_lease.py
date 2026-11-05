"""v1.35 takeover-hint LEASE acceptance (PowerPoint / WPS Presentation).

The bug being fixed (reported 2026-11-05) was first observed on a presentation: a deck opened by
the agent at 16:52 still carried the "〔AI接管中〕" frame-title prefix at 17:30, because every
cleanup path was agent-driven and the agent was gone.

WPS Presentation has no writable COM caption, so the marker is Win32 `SetWindowText` on the real
frame window (class `PP12FrameClass`). This test reads exactly that string (a tiny compiled probe)
instead of trusting any response field.

Contract under test (default lease 120 s; this test sets WOC_HINT_LEASE_MS=4000):

  * `open` prefixes the frame title and records the intent (ppt-title-hint.json)
  * silence longer than the lease strips the prefix, while the intent is kept
  * the next command re-applies the prefix
  * `serve -Action stop` releases immediately on the way out
  * `finish` ends the collaboration (intent off, no prefix)

    python tests/ppt/check_hint_lease.py
"""
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2]
PPT = ROOT / "scripts" / "ppt.ps1"
SRC = ROOT / "tests" / "ppt" / "out" / "master" / "deck.pptx"
WORK = ROOT / ".tmp"
COPY = WORK / "hint-lease.pptx"
TAG = "〔AI接管中〕"
# The title probe spawns PowerShell + Add-Type (~2 s per call), and probing does NOT renew the
# lease (it is not a skill command). A 4 s lease would therefore let the mark expire *between* two
# probes and make an assertion flaky; 8 s keeps the whole probe sequence inside one lease period
# while still being far below the 120 s default. The silent wait below is 2x the lease.
LEASE_MS = "8000"
SILENT_WAIT_S = 18.0
TITLE_STATE = Path(os.environ.get("LOCALAPPDATA", "")) / "wps-office-cooperate" / "ppt-title-hint.json"
PROBE = WORK / "ppt-frame-titles.ps1"

PROBE_SOURCE = r"""
param([string]$Out)
$sig = @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class WocTitles {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  public static List<string> Titles() {
    List<string> found = new List<string>();
    EnumWindows(delegate(IntPtr h, IntPtr l) {
      StringBuilder cls = new StringBuilder(64);
      GetClassName(h, cls, cls.Capacity);
      if (!String.Equals(cls.ToString(), "PP12FrameClass", StringComparison.Ordinal)) return true;
      if (!IsWindowVisible(h)) return true;
      StringBuilder tb = new StringBuilder(1024);
      GetWindowText(h, tb, tb.Capacity);
      if (tb.Length > 0) found.Add(tb.ToString());
      return true;
    }, IntPtr.Zero);
    return found;
  }
}
'@
Add-Type -TypeDefinition $sig -Language CSharp
$titles = [WocTitles]::Titles()
[System.IO.File]::WriteAllText($Out, ($titles -join "`n"), (New-Object System.Text.UTF8Encoding($false)))
"""


def child_env():
    env = dict(os.environ)
    env["WOC_HINT_LEASE_MS"] = LEASE_MS
    return env


def run(*args, timeout=300):
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(PPT), *args],
        capture_output=True, timeout=timeout, env=child_env(),
    )
    out = p.stdout.decode("utf-8", "replace").strip()
    try:
        return json.loads(out)
    except ValueError:
        raise SystemExit("not JSON from ppt %s: %s / %s"
                         % (args[0], out[:400], p.stderr.decode("utf-8", "replace")[:400]))


def check(label, cond, detail=""):
    print(("  PASS  " if cond else "  FAIL  ") + label + (("  <- " + str(detail)) if detail else ""))
    return bool(cond)


def frame_titles():
    """All visible PP12FrameClass titles (UTF-8, read back from a file to dodge console encoding)."""
    WORK.mkdir(exist_ok=True)
    PROBE.write_text(PROBE_SOURCE, encoding="ascii")
    out = WORK / "ppt-frame-titles.txt"
    if out.exists():
        out.unlink()
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(PROBE), str(out)],
                   capture_output=True, timeout=120, env=child_env())
    try:
        text = out.read_text(encoding="utf-8")
    except Exception:
        return []
    return [t for t in text.splitlines() if t.strip()]


def our_titles():
    return [t for t in frame_titles() if COPY.name.lower() in t.lower()]


def intent_on():
    try:
        data = json.loads(TITLE_STATE.read_text(encoding="utf-8"))
    except Exception:
        return False
    key = str(COPY).lower()
    return int(data.get(key, 0)) == 1


def main():
    if not SRC.exists():
        raise SystemExit("fixture deck missing: %s" % SRC)
    WORK.mkdir(exist_ok=True)
    try:
        shutil.copyfile(SRC, COPY)
    except PermissionError:
        run("close", "-Doc", str(COPY))
        shutil.copyfile(SRC, COPY)
    # Fresh daemon so the small lease actually applies (an earlier session may hold one alive).
    run("serve", "-Action", "reset")
    time.sleep(1.0)

    ok = True
    print("== P1 open prefixes the frame title and records the intent ==")
    run("open", "-Doc", str(COPY))
    titles = our_titles()
    ok &= check("frame title carries the prefix", any(TAG in t for t in titles), titles)
    ok &= check("takeover intent recorded", intent_on())

    print("== P2 a command keeps the lease alive ==")
    run("read", "-Doc", str(COPY), "-From", "1", "-Count", "1")
    titles = our_titles()
    ok &= check("prefix still there after a command", any(TAG in t for t in titles), titles)

    print("== P3 silence longer than the lease strips the prefix (the reported bug) ==")
    time.sleep(SILENT_WAIT_S)   # stay silent: any wrapper call would renew the lease
    titles = our_titles()
    ok &= check("helper removed the prefix on its own", titles and not any(TAG in t for t in titles), titles)
    ok &= check("intent is kept for a resumed session", intent_on())

    print("== P4 the next command re-applies it ==")
    run("read", "-Doc", str(COPY), "-From", "1", "-Count", "1")
    titles = our_titles()
    ok &= check("prefix back after one command", any(TAG in t for t in titles), titles)

    print("== P5 serve stop releases on the way out ==")
    run("serve", "-Action", "stop")
    time.sleep(2.0)
    ok &= check("prefix gone right after serve stop", not any(TAG in t for t in our_titles()), our_titles())

    print("== P6 finish ends the collaboration ==")
    run("hint", "-Doc", str(COPY), "-Quiet")
    run("serve", "-Action", "reset")
    time.sleep(1.0)
    run("finish", "-Doc", str(COPY), "-Quiet")
    ok &= check("finish turned the intent off", intent_on() is False)
    ok &= check("no prefix after finish", not any(TAG in t for t in our_titles()), our_titles())

    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
