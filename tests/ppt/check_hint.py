"""Takeover-hint acceptance (PowerPoint / WPS Presentation) - v1.11 contract + v1.27 slide badge.

The deck side has no writable per-document caption (WPS draws its tab from the file name),
so the visible marker is a per-slide badge. Contract under test:

  * hint works, stays silent with -Quiet, tells the truth about the missing COM caption
  * hint stamps one badge per slide by default (slideMarkCount == slide count)
  * the badge never shows up as deck content: read/search/snapshot/diff/check skip it,
    read reports it as takeoverBadge instead
  * save strips it from the file (deliverables stay clean), -Clear removes it from the deck
  * -NoSlideMark leaves the deck byte-identical in shape terms

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
MARK = "woc-ai-takeover-mark"


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


def read_slide(n=1):
    return run("read", "-Doc", str(COPY), "-From", str(n), "-Count", "1")["result"]["items"][0]


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
    pristine_shapes = len(read_slide(1)["shapes"])
    ok &= check("baseline read has no badge", not read_slide(1).get("takeoverBadge"), pristine_shapes)

    print("== P2 hint: quiet balloon, honest caption, badge on every slide ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "AI 正在改第 3 页", "-Quiet")["result"]
    ok &= check("quiet suppresses the balloon", r.get("toast") is False, r.get("toast"))
    ok &= check("caption honestly unsupported", r.get("captionSupported") is False, r.get("captionSupported"))
    ok &= check("limitation explained", bool(r.get("limitation")), r.get("limitation"))
    ok &= check("deck named", r.get("name") == COPY.name, r.get("name"))
    ok &= check("badge stamped on every slide", r.get("slideMarkCount") == slides,
                "%s vs slides=%s" % (r.get("slideMarkCount"), slides))
    ok &= check("slideMarked flag set", r.get("slideMarked") is True, r.get("slideMarked"))

    print("== P3 the badge is not deck content ==")
    s1 = read_slide(1)
    ok &= check("read flags takeoverBadge", s1.get("takeoverBadge") is True, s1.get("takeoverBadge"))
    ok &= check("read shape list unchanged by the badge", len(s1["shapes"]) == pristine_shapes,
                "%s -> %s" % (pristine_shapes, len(s1["shapes"])))
    ok &= check("no mark shape leaks into read",
                not any(sh["name"] == MARK for sh in s1["shapes"]))
    hits = run("search", "-Doc", str(COPY), "-Text", "AI接管中")["result"]
    ok &= check("search ignores the badge text", hits.get("count") == 0, hits.get("count"))
    chk = run("check", "-Doc", str(COPY))["result"]
    bad = [i for i in chk.get("issues", []) if i.get("shape") == MARK]
    ok &= check("check does not report the badge", not bad, len(bad))

    print("== P4 snapshot/diff stay quiet ==")
    run("snapshot", "-Doc", str(COPY), "-Tag", "with-badge")
    d = run("diff", "-Doc", str(COPY), "-Since", "with-badge")["result"]
    ok &= check("diff sees no text change from the badge", d["summary"]["textChanges"] == 0,
                d["summary"])

    print("== P5 -NoSlideMark leaves the deck alone ==")
    r = run("hint", "-Doc", str(COPY), "-NoSlideMark", "-Quiet")["result"]
    ok &= check("-NoSlideMark stamps nothing", r.get("slideMarked") is False, r.get("slideMarked"))
    ok &= check("read has no badge after -NoSlideMark", not read_slide(1).get("takeoverBadge"))

    print("== P6 save strips the badge from the file ==")
    run("hint", "-Doc", str(COPY), "-Quiet")          # stamp again
    run("save", "-Doc", str(COPY))
    run("close", "-Doc", str(COPY))
    r = run("open", "-Doc", str(COPY), "-NoHint")["result"]
    ok &= check("reopened the saved copy", r.get("slides") == slides, r.get("slides"))
    ok &= check("saved file carries no badge", not read_slide(1).get("takeoverBadge"))
    ok &= check("saved file has no mark shape",
                not any(sh["name"] == MARK for sh in read_slide(1)["shapes"]))

    print("== P7 hint -Clear removes the badge ==")
    run("hint", "-Doc", str(COPY), "-Quiet")
    r = run("hint", "-Doc", str(COPY), "-Clear", "-Quiet")["result"]
    ok &= check("-Clear reports no badge left", r.get("slideMarked") is False, r.get("slideMarked"))
    ok &= check("deck has no badge after -Clear", not read_slide(1).get("takeoverBadge"))

    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
