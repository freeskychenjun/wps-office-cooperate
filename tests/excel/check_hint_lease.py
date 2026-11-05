"""v1.35 takeover-hint LEASE acceptance (Excel / WPS Spreadsheets).

The bug being fixed (reported 2026-11-05): an agent takes a document over, the user ends the chat /
the agent disappears, and the window stays branded "〔AI接管中〕" forever because every cleanup path
(`finish` / `close` / `hint -Clear`) is agent-driven.

Contract under test (default lease 120 s; this test sets WOC_HINT_LEASE_MS=4000):

  * `open` brands the caption and records the intent (hint-state-excel.json)
  * silence longer than the lease strips the VISIBLE mark (caption back to the plain name), while
    the intent is kept
  * the next command re-applies the mark (a paused session comes back branded)
  * `serve -Action stop` releases immediately on the way out
  * `finish` ends the collaboration (intent entry removed)

    python tests/excel/check_hint_lease.py
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
COPY = WORK / "hint-lease.xlsx"
TAG = "〔AI接管中〕"
LEASE_MS = "4000"
STATE = Path(os.environ.get("LOCALAPPDATA", "")) / "wps-office-cooperate" / "hint-state-excel.json"


def child_env():
    env = dict(os.environ)
    env["WOC_HINT_LEASE_MS"] = LEASE_MS
    return env


def ensure_fixture():
    src = ROOT / "tests" / "excel" / "work" / "test.xlsx"
    if src.exists():
        return src
    out = Path(tempfile.gettempdir()) / "woc-excel-lease-fixture.xlsx"
    if not out.exists():
        subprocess.run([sys.executable, str(BUILDER), "--out", str(out)], check=True)
    return out


def run(*args, timeout=300):
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(EXCEL), *args],
        capture_output=True, timeout=timeout, env=child_env(),
    )
    out = p.stdout.decode("utf-8", "replace").strip()
    try:
        return json.loads(out)
    except ValueError:
        raise SystemExit("not JSON from excel %s: %s / %s"
                         % (args[0], out[:400], p.stderr.decode("utf-8", "replace")[:400]))


def check(label, cond, detail=""):
    print(("  PASS  " if cond else "  FAIL  ") + label + (("  <- " + str(detail)) if detail else ""))
    return bool(cond)


def status_for(name):
    books = run("status")["result"].get("workbooks") or []
    mine = [b for b in books if b.get("name") == name]
    return mine[0] if mine else {}


def intent_has(name):
    try:
        data = json.loads(STATE.read_text(encoding="utf-8"))
    except Exception:
        return False
    return any(str(v) and str(k).endswith(name.lower()) for k, v in data.items())


def main():
    src = ensure_fixture()
    WORK.mkdir(exist_ok=True)
    try:
        shutil.copyfile(src, COPY)
    except PermissionError:
        run("close", "-Doc", str(COPY))
        shutil.copyfile(src, COPY)
    # Fresh daemon so the small lease actually applies (an earlier session may hold one alive).
    run("serve", "-Action", "reset")
    time.sleep(1.0)

    ok = True
    print("== E1 open brands the window and records the intent ==")
    r = run("open", "-Doc", str(COPY))["result"]
    h = r.get("hint") or {}
    ok &= check("caption carries the marker", TAG in str(h.get("caption", "")), h.get("caption"))
    ok &= check("status reports hinted", status_for(COPY.name).get("hinted") is True, status_for(COPY.name))
    ok &= check("takeover intent recorded", intent_has(COPY.name))

    print("== E2 commands keep the lease alive ==")
    run("read", "-Doc", str(COPY), "-Count", "3")
    ok &= check("mark still on after a command", status_for(COPY.name).get("hinted") is True)

    print("== E3 silence longer than the lease strips the visible mark (the reported bug) ==")
    # NB: the status call itself renews the lease, so it must come AFTER the silent wait.
    time.sleep(9.0)
    # v1.36: an expired lease now ENDS the helper process (it used to idle on for 20 minutes), so
    # a user who closes the workbook right after the agent goes away has no automation client left.
    dg = run("serve", "-Action", "status")["result"]
    ok &= check("helper process exited with the lease (v1.36)", dg.get("running") is False, dg)
    st = status_for(COPY.name)
    ok &= check("helper released the mark on its own", st.get("hinted") is False, st.get("hinted"))
    ok &= check("caption back to the plain file name", str(st.get("caption", "")) == COPY.name, st.get("caption"))
    ok &= check("intent is kept for a resumed session", intent_has(COPY.name))

    print("== E4 the next command re-applies it ==")
    run("read", "-Doc", str(COPY), "-Count", "1")
    ok &= check("mark back after one command", status_for(COPY.name).get("hinted") is True)

    print("== E5 serve stop releases on the way out ==")
    run("serve", "-Action", "stop")
    time.sleep(2.0)
    ok &= check("mark gone right after serve stop", status_for(COPY.name).get("hinted") is False)

    print("== E6 finish ends the collaboration ==")
    run("hint", "-Doc", str(COPY), "-Text", "lease test", "-Quiet")
    run("serve", "-Action", "reset")
    time.sleep(1.0)
    run("finish", "-Doc", str(COPY), "-Quiet")
    ok &= check("finish dropped the intent entry", not intent_has(COPY.name), STATE)

    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
