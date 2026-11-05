"""v1.36 hidden-window SELF-HEAL acceptance (Excel / WPS Spreadsheets).

Field bug (reported 2026-11-05): after an agent session took a workbook over and the user closed it,
WPS's own close path (which crashes on this host several times a day) came back with the frame chrome
visible but `Application.Visible = false` and every document child window hidden - the window looked
frozen ("卡死"): tabs and title bar only, grey empty content area.

Fix under test: before executing any document command the helper checks `Application.Visible`; if the
app is hidden while workbooks exist it re-shows the app and activates the active workbook (see
`HealWindow()` in scripts/helper/excel/EtHelper.cs), so the next agent command un-sticks whatever the
host broke. This test hides the live app through a second COM client and expects the next wrapper call
to repair it.

    python tests/excel/check_window_heal.py
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2]
EXCEL = ROOT / "scripts" / "excel.ps1"
BUILDER = ROOT / "tests" / "excel" / "make_workbook.py"
WORK = ROOT / ".tmp"
COPY = WORK / "window-heal.xlsx"

PROBE = r"""
$ErrorActionPreference = 'Stop'
$a = [Runtime.InteropServices.Marshal]::GetActiveObject('KET.Application')
if ($env:WOC_HIDE_APP -eq '1') { $a.Visible = $false; Start-Sleep -Milliseconds 300 }
$vis = [string]$a.Visible
$wins = ''
try { $wins = [string]$a.Windows.Count } catch { $wins = 'err' }
$winVis = ''
try { $winVis = [string]$a.ActiveWindow.Visible } catch { $winVis = 'err' }
[Console]::Out.Write($vis + '|' + $wins + '|' + $winVis)
"""


def run(*args, timeout=300):
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(EXCEL), *args],
        capture_output=True, timeout=timeout, env=dict(os.environ),
    )
    out = p.stdout.decode("utf-8", "replace").strip()
    try:
        return json.loads(out)
    except ValueError:
        raise SystemExit("not JSON from excel %s: %s / %s"
                         % (args[0], out[:400], p.stderr.decode("utf-8", "replace")[:400]))


def probe(hide=False):
    env = dict(os.environ)
    env["WOC_HIDE_APP"] = "1" if hide else "0"
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", PROBE],
        capture_output=True, timeout=120, env=env,
    )
    out = p.stdout.decode("utf-8", "replace").strip()
    if not out:
        raise SystemExit("COM probe failed: " + p.stderr.decode("utf-8", "replace")[:600])
    vis, wins, winvis = out.split("|")
    return {"visible": vis.lower() == "true", "windows": wins, "activeWindowVisible": winvis.lower() == "true"}


def check(label, cond, detail=""):
    print(("  PASS  " if cond else "  FAIL  ") + label + (("  <- " + str(detail)) if detail else ""))
    return bool(cond)


def ensure_fixture():
    src = ROOT / "tests" / "excel" / "work" / "test.xlsx"
    if src.exists():
        return src
    out = Path(tempfile.gettempdir()) / "woc-excel-heal-fixture.xlsx"
    if not out.exists():
        subprocess.run([sys.executable, str(BUILDER), "--out", str(out)], check=True)
    return out


def main():
    WORK.mkdir(exist_ok=True)
    src = ensure_fixture()
    try:
        shutil.copyfile(src, COPY)
    except PermissionError:
        run("close", "-Doc", str(COPY))
        shutil.copyfile(src, COPY)

    run("serve", "-Action", "reset")
    time.sleep(1.0)

    ok = True
    print("== H1 a normal open leaves the app visible ==")
    r = run("open", "-Doc", str(COPY), "-NoHint")
    ok &= check("open ok", r.get("ok") is True, json.dumps(r)[:200])
    ok &= check("app visible after open", probe()["visible"] is True)

    print("== H2 hiding the app simulates the frozen window the user saw ==")
    st = probe(hide=True)
    ok &= check("app is now hidden", st["visible"] is False, st)

    print("== H3 the next command heals it (v1.36) ==")
    r = run("read", "-Doc", str(COPY), "-Sheet", "Data", "-Range", "A1:C3")
    ok &= check("read still succeeds", r.get("ok") is True, json.dumps(r)[:200])
    res = r.get("result") or {}
    ok &= check("read returned data", bool(res.get("values")), str(res.get("values"))[:200])
    st = probe()
    ok &= check("app visible again without user action", st["visible"] is True, st)
    ok &= check("active window restored", st["activeWindowVisible"] is True, st)

    print("== H4 status now reports a visible app ==")
    r = run("status")
    app = (r.get("result") or {}).get("app") or {}
    ok &= check("status app.visible is true", app.get("visible") is True, app)

    run("close", "-Doc", str(COPY))
    run("serve", "-Action", "reset")
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
