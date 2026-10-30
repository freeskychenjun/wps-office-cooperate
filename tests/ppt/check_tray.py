"""v1.12 resident tray-indicator acceptance (PowerPoint).

For Presentation the tray icon is the ONLY persistent "AI is editing" channel (no writable
caption, no status bar), so this is the most important app for it.

    python tests/ppt/check_tray.py
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
COPY = WORK / "tray-check.pptx"
STATE = Path(os.environ["LOCALAPPDATA"]) / "wps-office-cooperate" / "tray-ppt.json"


def run(*args):
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(PPT), *args],
        capture_output=True, timeout=300,
    )
    return json.loads(p.stdout.decode("utf-8", "replace").strip())


def check(label, cond, detail=""):
    print(("  PASS  " if cond else "  FAIL  ") + label + (("  <- " + str(detail)) if detail else ""))
    return bool(cond)


def alive(pid):
    p = subprocess.run(["tasklist", "/FI", "PID eq %d" % pid, "/NH"], capture_output=True, text=True, timeout=30)
    return str(pid) in (p.stdout or "")


def wait_gone(pid, seconds=8):
    end = time.time() + seconds
    while time.time() < end:
        if not alive(pid):
            return True
        time.sleep(0.4)
    return False


EXTERNAL_CLOSE_PS = """param([string]$ProgId, [string]$Path)
$ErrorActionPreference = 'Stop'
$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
$full = [System.IO.Path]::GetFullPath($Path)
if ($ProgId -like '*KWPS*') {
  foreach ($d in $app.Documents) { if ([string]$d.FullName -eq $full) { $d.Close(0) } }
} elseif ($ProgId -like '*KET*') {
  foreach ($w in $app.Workbooks) { if ([string]$w.FullName -eq $full) { $w.Close($false) } }
} else {
  foreach ($p in $app.Presentations) { if ([string]$p.FullName -eq $full) { $p.Close() } }
}
"""


def external_close(progid, path):
    """Simulate the user closing the document in WPS (our own close command would be too kind)."""
    script = WORK / "external-close.ps1"
    if not script.exists():
        script.write_text(EXTERNAL_CLOSE_PS, encoding="utf-8")
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
                    "-ProgId", progid, "-Path", str(path)], capture_output=True, timeout=120)
def main():
    WORK.mkdir(exist_ok=True)
    try:
        shutil.copyfile(SRC, COPY)
    except PermissionError:
        run("close", "-Doc", str(COPY))
        shutil.copyfile(SRC, COPY)
    try:
        if STATE.exists():
            STATE.unlink()
    except OSError:
        pass

    ok = True
    print("== P1 hint -Tray gives Presentation its persistent channel ==")
    run("open", "-Doc", str(COPY))
    r = run("hint", "-Doc", str(COPY), "-Text", "正在改第 3 页", "-Quiet", "-Tray")["result"]
    tray = r.get("tray", {})
    pid = tray.get("pid") or 0
    ok &= check("tray reported started", tray.get("started") is True and pid > 0, tray)
    ok &= check("daemon process alive", alive(pid), pid)
    ok &= check("caption still unsupported (truthful)", r.get("captionSupported") is False, r.get("captionSupported"))

    print("== P2 finish removes the icon ==")
    r = run("finish", "-Doc", str(COPY), "-Quiet")["result"]
    ok &= check("finish reports tray stop", r.get("trayStopped") is True, r.get("trayStopped"))
    ok &= check("daemon exited", wait_gone(pid), pid)

    print("== P3 the icon retires itself when the USER closes the deck ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "正在改第 4 页", "-Quiet", "-Tray")["result"]
    pid2 = (r.get("tray") or {}).get("pid") or 0
    ok &= check("tray up before the close", pid2 > 0 and alive(pid2), pid2)
    external_close("KWPP.Application", COPY)
    ok &= check("icon retired after the user closed the deck", wait_gone(pid2, 25), pid2)

    run("open", "-Doc", str(COPY))
    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
