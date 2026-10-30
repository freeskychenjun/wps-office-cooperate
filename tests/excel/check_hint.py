"""v1.11 takeover-hint acceptance (Excel / WPS Spreadsheets).

    python tests/excel/check_hint.py
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2]
EXCEL = ROOT / "scripts" / "excel.ps1"
SRC = ROOT / "tests" / "excel" / "work" / "test.xlsx"
WORK = ROOT / ".tmp"
COPY = WORK / "hint-check.xlsx"
TAG = "〔AI〕"


def run(*args):
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(EXCEL), *args],
        capture_output=True, timeout=300,
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


def main():
    if not SRC.exists():
        raise SystemExit("workbook missing: %s" % SRC)
    WORK.mkdir(exist_ok=True)
    try:
        shutil.copyfile(SRC, COPY)
    except PermissionError:
        run("close", "-Doc", str(COPY))
        shutil.copyfile(SRC, COPY)

    ok = True
    print("== E1 open tags the window caption ==")
    r = run("open", "-Doc", str(COPY))["result"]
    h = r.get("hint", {})
    cap = str(h.get("caption", ""))
    ok &= check("caption set (Windows(1).Caption)", h.get("captionSet") is True, h)
    ok &= check("caption carries a marker", TAG in cap, cap)
    ok &= check("caption keeps the file name", COPY.name in cap, cap)

    print("== E2 status reports the tag ==")
    books = run("status")["result"]["workbooks"]
    mine = [b for b in books if b.get("name") == COPY.name]
    ok &= check("hinted true", bool(mine) and mine[0].get("hinted") is True, mine)

    print("== E3 hint writes the status bar only ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "正在核对第 3 张表", "-Quiet")["result"]
    ok &= check("status bar write accepted", r.get("statusBar") is True, r)
    ok &= check("quiet suppresses the balloon", r.get("toast") is False, r)

    print("== E4 finish restores ==")
    r = run("finish", "-Doc", str(COPY), "-Quiet")["result"]
    ok &= check("finish clears caption", r.get("hintCleared", {}).get("captionCleared") is True, r)
    books = run("status")["result"]["workbooks"]
    mine = [b for b in books if b.get("name") == COPY.name]
    ok &= check("hinted false after finish", bool(mine) and mine[0].get("hinted") is False, mine)

    print("== E5 explicit clear ==")
    r = run("hint", "-Doc", str(COPY), "-Clear", "-Quiet")["result"]
    ok &= check("caption back to file name", r.get("caption") == COPY.name, r)

    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
