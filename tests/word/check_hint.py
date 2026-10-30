"""v1.11 takeover-hint acceptance (Word / WPS Writer).

Runs the real skill against a scratch copy of the corpus and asserts the hint contract:
open tags the window caption, status reports it, hint updates the status bar silently,
finish / hint -Clear restore the caption, and nothing touches the document text.

    python tests/word/check_hint.py
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2]
WORD = ROOT / "scripts" / "word.ps1"
SRC = ROOT / "tests" / "word" / "corpus" / "report-300.pristine.docx"
WORK = ROOT / ".tmp"
COPY = WORK / "hint-check.docx"
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


def main():
    if not SRC.exists():
        raise SystemExit("corpus missing: %s" % SRC)
    WORK.mkdir(exist_ok=True)
    try:
        shutil.copyfile(SRC, COPY)
    except PermissionError:
        run("close", "-Doc", str(COPY))
        shutil.copyfile(SRC, COPY)

    ok = True
    print("== H1 open tags the caption ==")
    r = run("open", "-Doc", str(COPY))["result"]
    h = r.get("hint", {})
    cap = str(h.get("caption", ""))
    ok &= check("caption set", h.get("captionSet") is True, h)
    ok &= check("caption carries a marker", TAG in cap, cap)
    ok &= check("caption keeps the file name", COPY.name in cap, cap)
    paras_before = r.get("paragraphs")

    print("== H2 status reports the tag ==")
    docs = run("status")["result"]["documents"]
    mine = [d for d in docs if d.get("name") == COPY.name]
    ok &= check("hinted true", bool(mine) and mine[0].get("hinted") is True, mine)

    print("== H3 hint refreshes quietly ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "正在核对全部段落", "-Quiet")["result"]
    ok &= check("status bar write accepted", r.get("statusBar") is True, r)
    ok &= check("quiet suppresses the balloon", r.get("toast") is False, r)
    ok &= check("tag kept", TAG in str(r.get("caption", "")), r.get("caption"))

    print("== H4 finish restores ==")
    r = run("finish", "-Doc", str(COPY), "-Quiet")["result"]
    ok &= check("finish clears caption", r.get("hintCleared", {}).get("captionCleared") is True, r)
    docs = run("status")["result"]["documents"]
    mine = [d for d in docs if d.get("name") == COPY.name]
    ok &= check("hinted false after finish", bool(mine) and mine[0].get("hinted") is False, mine)

    print("== H5 explicit clear + document untouched ==")
    r = run("hint", "-Doc", str(COPY), "-Clear", "-Quiet")["result"]
    ok &= check("caption back to file name", r.get("caption") == COPY.name, r)
    r = run("read", "-Doc", str(COPY), "-Count", "1")["result"]
    ok &= check("paragraph count unchanged (no content write)", r.get("paragraphs") == paras_before,
                "%s -> %s" % (paras_before, r.get("paragraphs")))
    r = run("revisions", "-Doc", str(COPY))["result"]
    ok &= check("no revisions were created by the hint", (r.get("count") or 0) == 0, r.get("count"))

    run("close", "-Doc", str(COPY))

    print("== H6 read-only label + finish without -Doc ==")
    r = run("open", "-Doc", str(COPY), "-ReadOnly")["result"]
    h = r.get("hint", {})
    cap = str(h.get("caption", ""))
    ok &= check("read-only docs are labelled read-only", "〔AI 只读〕" in cap, h)
    r = run("finish", "-Quiet")["result"]
    ok &= check("finish without -Doc clears the active document", r.get("hintCleared", {}).get("captionCleared") is True, r)
    run("close", "-Doc", str(COPY))

    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
