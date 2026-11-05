"""v1.35 takeover-hint LEASE acceptance (Word / WPS Writer).

The bug being fixed (reported 2026-11-05): an agent takes the document over, the user ends the
chat / the agent disappears, and the window stays branded "〔AI接管中〕" forever because every
cleanup path was agent-driven.

Contract under test (default lease is 120 s; this test sets WOC_HINT_LEASE_MS=4000):

  * `open` brands the caption and records the takeover INTENT (hint-state-word.json)
  * while commands keep coming the lease is renewed (mark stays)
  * after the lease expires with no commands the resident helper strips the VISIBLE mark
    (caption back to the plain file name) but keeps the intent
  * the next command re-applies the mark (a paused session comes back branded)
  * `serve -Action stop` releases immediately on the way out
  * `finish` ends the collaboration: intent entry removed, nothing re-applies later

    python tests/word/check_hint_lease.py
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
COPY = WORK / "hint-lease.docx"
TAG = "〔AI接管中〕"
LEASE_MS = "4000"
STATE = Path(os.environ.get("LOCALAPPDATA", "")) / "wps-office-cooperate" / "hint-state-word.json"


def child_env():
    env = dict(os.environ)
    env["WOC_HINT_LEASE_MS"] = LEASE_MS
    return env


def run(*args, timeout=300):
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(WORD), *args],
        capture_output=True, timeout=timeout, env=child_env(),
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


def status_for(name):
    docs = run("status")["result"].get("documents") or []
    mine = [d for d in docs if d.get("name") == name]
    return mine[0] if mine else {}


def intent_has(name):
    try:
        data = json.loads(STATE.read_text(encoding="utf-8"))
    except Exception:
        return False
    return any(str(v) and str(k).endswith(name.lower()) for k, v in data.items())


def wait_quiet(seconds=9.0):
    """Stay completely silent (any wrapper call would renew the lease), then look once."""
    time.sleep(seconds)


def main():
    if not SRC.exists():
        raise SystemExit("corpus missing: %s" % SRC)
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
    print("== L1 open brands the window and records the intent ==")
    r = run("open", "-Doc", str(COPY))["result"]
    h = r.get("hint") or {}
    ok &= check("caption carries the marker", TAG in str(h.get("caption", "")), h.get("caption"))
    ok &= check("status reports hinted", status_for(COPY.name).get("hinted") is True, status_for(COPY.name))
    ok &= check("takeover intent recorded", intent_has(COPY.name))

    print("== L2 commands keep the lease alive ==")
    run("read", "-Doc", str(COPY), "-Count", "1")
    run("read", "-Doc", str(COPY), "-Count", "1", "-From", "2")
    ok &= check("mark still on after back-to-back commands", status_for(COPY.name).get("hinted") is True)

    print("== L3 silence longer than the lease strips the visible mark (the reported bug) ==")
    # NB: the status call itself renews the lease, so it must come AFTER the silent wait.
    wait_quiet()
    st = status_for(COPY.name)
    ok &= check("helper released the mark on its own", st.get("hinted") is False, st.get("hinted"))
    ok &= check("caption back to the plain file name", str(st.get("caption", "")) == COPY.name, st.get("caption"))
    ok &= check("intent is kept for a resumed session", intent_has(COPY.name))

    print("== L4 the next command re-applies it ==")
    run("read", "-Doc", str(COPY), "-Count", "1")
    ok &= check("mark back after one command", status_for(COPY.name).get("hinted") is True)

    print("== L5 serve stop releases on the way out ==")
    run("serve", "-Action", "stop")
    time.sleep(2.0)
    ok &= check("mark gone right after serve stop", status_for(COPY.name).get("hinted") is False)

    print("== L6 finish ends the collaboration (no re-branding later) ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "lease test", "-Quiet")["result"]
    ok &= check("hint re-brands", r.get("caption", "").startswith(TAG), r.get("caption"))
    run("serve", "-Action", "reset")   # kill the daemon; the intent must survive/clear explicitly
    time.sleep(1.0)
    run("finish", "-Doc", str(COPY), "-Quiet")
    ok &= check("finish dropped the intent entry", not intent_has(COPY.name), STATE)
    r = run("hint", "-Doc", str(COPY), "-Clear", "-Quiet")["result"]
    ok &= check("clear reports the plain caption", r.get("caption") == COPY.name, r.get("caption"))

    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
