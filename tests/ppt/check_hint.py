"""v1.11 takeover-hint acceptance (PowerPoint / WPS Presentation).

The deck side has no writable per-document caption, so the contract under test is:
hint works, stays silent with -Quiet, and tells the truth about the missing caption.

    python tests/ppt/check_hint.py
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2]
PPT = ROOT / "scripts" / "ppt.ps1"
SRC = ROOT / "tests" / "ppt" / "out" / "master" / "deck.pptx"
WORK = ROOT / ".tmp"
COPY = WORK / "hint-check.pptx"


def run(*args):
    p = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(PPT), *args],
        capture_output=True, timeout=300,
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


def main():
    if not SRC.exists():
        raise SystemExit("deck missing: %s" % SRC)
    WORK.mkdir(exist_ok=True)
    try:
        shutil.copyfile(SRC, COPY)
    except PermissionError:
        run("close", "-Doc", str(COPY))
        shutil.copyfile(SRC, COPY)

    ok = True
    print("== P1 open ==")
    r = run("open", "-Doc", str(COPY))["result"]
    ok &= check("deck opened", r.get("name") == COPY.name, r.get("name"))
    slides = r.get("slides")

    print("== P2 hint is a balloon and says why ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "AI 正在改第 3 页", "-Quiet")["result"]
    ok &= check("quiet suppresses the balloon", r.get("toast") is False, r)
    ok &= check("caption honestly unsupported", r.get("captionSupported") is False, r.get("captionSupported"))
    ok &= check("limitation explained", bool(r.get("limitation")), r.get("limitation"))
    ok &= check("deck named", r.get("name") == COPY.name, r.get("name"))

    print("== P3 the hint did not touch the deck ==")
    r = run("read", "-Doc", str(COPY), "-Count", "1")["result"]
    ok &= check("slide count unchanged", r.get("slides") == slides, "%s -> %s" % (slides, r.get("slides")))

    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
