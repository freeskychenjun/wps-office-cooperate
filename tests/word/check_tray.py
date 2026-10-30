"""v1.12 resident tray-indicator acceptance (Word).

Covers the real lifecycle: hint -Tray spawns a detached daemon, the state file carries the
tooltip data, the daemon exits on finish, and the restore command baked into the state file
(the same one the tray menu runs) really clears the caption.

    python tests/word/check_tray.py
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
WORD = ROOT / "scripts" / "word.ps1"
SRC = ROOT / "tests" / "word" / "corpus" / "report-300.pristine.docx"
WORK = ROOT / ".tmp"
COPY = WORK / "tray-check.docx"
STATE = Path(os.environ["LOCALAPPDATA"]) / "wps-office-cooperate" / "tray-word.json"
TAG = "〔AI〕"


def run(*args):
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(WORD), *args],
        capture_output=True, timeout=300,
    )
    out = p.stdout.decode("utf-8", "replace").strip()
    try:
        return json.loads(out)
    except ValueError:
        raise SystemExit("not JSON from word %s: %s / %s"
                         % (args[0], out[:400], p.stderr.decode("utf-8", "replace")[:400]))


def check(label, cond, detail=""):
    print(("  PASS  " if cond else "  FAIL  ") + label + (("  <- " + str(detail)) if detail else ""))
    return bool(cond)


def alive(pid):
    try:
        p = subprocess.run(["tasklist", "/FI", "PID eq %d" % pid, "/NH"],
                           capture_output=True, text=True, timeout=30)
        return str(pid) in (p.stdout or "")
    except Exception:
        return False


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
    if not SRC.exists():
        raise SystemExit("corpus missing: %s" % SRC)
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
    print("== T1 hint -Tray spawns the daemon ==")
    run("open", "-Doc", str(COPY))
    r = run("hint", "-Doc", str(COPY), "-Text", "正在核对第 3 节", "-Quiet", "-Tray")["result"]
    tray = r.get("tray", {})
    pid = tray.get("pid") or 0
    ok &= check("tray reported started", tray.get("started") is True and pid > 0, tray)
    ok &= check("daemon process alive", alive(pid), pid)
    ok &= check("state file written", STATE.exists(), STATE)

    print("== T2 state file carries the tooltip + restore command ==")
    st = json.loads(STATE.read_text(encoding="utf-8"))
    ok &= check("state names this document", st.get("doc", "").endswith("tray-check.docx"), st.get("doc"))
    ok &= check("state caption has the marker", TAG in st.get("caption", ""), st.get("caption"))
    ok &= check("state text is the progress sentence", st.get("text") == "正在核对第 3 节", st.get("text"))
    restore = st.get("restore") or []
    ok &= check("restore command points at hint -Clear", "hint" in restore and "-Clear" in restore, restore)
    ok &= check("restore carries the wrapper script", any(str(x).endswith("word.ps1") for x in restore), restore)

    print("== T3 the tray menu action (restore) really clears the caption ==")
    cmd = subprocess.run(restore, capture_output=True, timeout=120)
    ok &= check("restore command exits 0", cmd.returncode == 0, cmd.returncode)
    doc = [d for d in run("status")["result"]["documents"] if d.get("name") == COPY.name]
    ok &= check("caption cleared after restore", bool(doc) and doc[0].get("hinted") is False, doc)
    ok &= check("daemon exited after restore", wait_gone(pid), pid)

    print("== T4 starting again + finish removes the icon ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "正在核对第 4 节", "-Quiet", "-Tray")["result"]
    pid2 = (r.get("tray") or {}).get("pid") or 0
    ok &= check("second tray started", pid2 > 0 and alive(pid2), pid2)
    r = run("finish", "-Doc", str(COPY), "-Quiet")["result"]
    ok &= check("finish reports tray stop", (r.get("hintCleared") or {}).get("trayStopped") is True, r.get("hintCleared"))
    ok &= check("daemon exited after finish", wait_gone(pid2), pid2)

    print("== T5 the icon retires itself when the USER closes the document ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "正在核对第 5 节", "-Quiet", "-Tray")["result"]
    pid3 = (r.get("tray") or {}).get("pid") or 0
    ok &= check("tray up before the close", pid3 > 0 and alive(pid3), pid3)
    external_close("KWPS.Application", COPY)
    ok &= check("icon retired after the user closed the document", wait_gone(pid3, 25), pid3)

    run("open", "-Doc", str(COPY))
    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
