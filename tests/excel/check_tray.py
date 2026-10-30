"""v1.12 resident tray-indicator acceptance (Excel).

    python tests/excel/check_tray.py
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
EXCEL = ROOT / "scripts" / "excel.ps1"
SRC = ROOT / "tests" / "excel" / "work" / "test.xlsx"
WORK = ROOT / ".tmp"
COPY = WORK / "tray-check.xlsx"
STATE = Path(os.environ["LOCALAPPDATA"]) / "wps-office-cooperate" / "tray-excel.json"


def run(*args):
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(EXCEL), *args],
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
    print("== X1 hint -Tray spawns the daemon ==")
    run("open", "-Doc", str(COPY))
    r = run("hint", "-Doc", str(COPY), "-Text", "正在核对第 3 张表", "-Quiet", "-Tray")["result"]
    tray = r.get("tray", {})
    pid = tray.get("pid") or 0
    ok &= check("tray reported started", tray.get("started") is True and pid > 0, tray)
    ok &= check("daemon process alive", alive(pid), pid)
    ok &= check("state file written", STATE.exists(), STATE)

    print("== X2 finish removes the icon ==")
    r = run("finish", "-Doc", str(COPY), "-Quiet")["result"]
    ok &= check("finish reports tray stop", (r.get("hintCleared") or {}).get("trayStopped") is True, r.get("hintCleared"))
    ok &= check("daemon exited", wait_gone(pid), pid)

    print("== X3 the icon retires itself when the USER closes the workbook ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "正在核对第 4 张表", "-Quiet", "-Tray")["result"]
    pid2 = (r.get("tray") or {}).get("pid") or 0
    ok &= check("tray up before the close", pid2 > 0 and alive(pid2), pid2)
    external_close("KET.Application", COPY)
    ok &= check("icon retired after the user closed the workbook", wait_gone(pid2, 25), pid2)

    run("open", "-Doc", str(COPY))
    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
