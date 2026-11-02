"""Takeover-hint acceptance (PowerPoint / WPS Presentation) - v1.11 contract + v1.28 badge policy.

The deck side has no writable per-document caption (WPS draws its tab from the file name), so the
only in-app marker is a per-slide badge - and that badge is deck CONTENT. Contract under test:

  * hint works, stays silent with -Quiet, tells the truth about the missing COM caption
  * hint does NOT touch the deck by default: no badge, no new shapes, inAppHint explains why
  * hint -SlideMark is the opt-in: one badge per slide, read reports takeoverBadge, and the
    response warns that Ctrl+S would persist it
  * the badge never shows up as deck content: read/search/snapshot/diff/check skip it
  * save strips it from the file; -Clear / -NoSlideMark remove it from the deck
  * if the user saves with Ctrl+S anyway, the next command repairs the file (badgeRepaired)

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

# The point of the v1.28 repair path is a save we cannot intercept, so the test performs one:
# a raw COM save (WPS UI equivalent of the user pressing Ctrl+S) while the badge is in the deck.
RAW_SAVE = r"""
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('KWPP.Application')
$deck = $null
for ($i = 1; $i -le $app.Presentations.Count; $i++) {
    $pr = $app.Presentations.Item($i)
    if ($pr.Name -eq 'hint-check.pptx') { $deck = $pr; break }
}
if ($null -eq $deck) { 'NO_DECK'; exit 1 }
$deck.Save()
'SAVED'
"""


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


def raw_save():
    p = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", RAW_SAVE],
                       capture_output=True, timeout=120)
    return p.stdout.decode("utf-8", "replace").strip()


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

    print("== P2 default hint keeps the deck untouched ==")
    r = run("hint", "-Doc", str(COPY), "-Text", "AI 正在改第 3 页", "-Quiet")["result"]
    ok &= check("quiet suppresses the balloon", r.get("toast") is False, r.get("toast"))
    ok &= check("caption honestly unsupported", r.get("captionSupported") is False, r.get("captionSupported"))
    ok &= check("limitation explained", bool(r.get("limitation")), r.get("limitation"))
    ok &= check("deck named", r.get("name") == COPY.name, r.get("name"))
    ok &= check("no badge by default", r.get("slideMarked") is False, r.get("slideMarked"))
    ok &= check("inAppHint explains the opt-in", bool(r.get("inAppHint")), r.get("inAppHint"))
    ok &= check("deck has no badge after a plain hint", not read_slide(1).get("takeoverBadge"))

    print("== P3 hint -SlideMark is the opt-in ==")
    r = run("hint", "-Doc", str(COPY), "-SlideMark", "-Quiet")["result"]
    ok &= check("badge stamped on every slide", r.get("slideMarkCount") == slides,
                "%s vs slides=%s" % (r.get("slideMarkCount"), slides))
    ok &= check("slideMarked flag set", r.get("slideMarked") is True, r.get("slideMarked"))
    ok &= check("response warns about Ctrl+S", bool(r.get("slideMarkWarning")), r.get("slideMarkWarning"))
    s1 = read_slide(1)
    ok &= check("read flags takeoverBadge", s1.get("takeoverBadge") is True, s1.get("takeoverBadge"))
    ok &= check("read shape list unchanged by the badge", len(s1["shapes"]) == pristine_shapes,
                "%s -> %s" % (pristine_shapes, len(s1["shapes"])))
    ok &= check("no mark shape leaks into read", not any(sh["name"] == MARK for sh in s1["shapes"]))
    hits = run("search", "-Doc", str(COPY), "-Text", "AI接管中")["result"]
    ok &= check("search ignores the badge text", hits.get("count") == 0, hits.get("count"))
    chk = run("check", "-Doc", str(COPY))["result"]
    bad = [i for i in chk.get("issues", []) if i.get("shape") == MARK]
    ok &= check("check does not report the badge", not bad, len(bad))
    run("snapshot", "-Doc", str(COPY), "-Tag", "with-badge")
    d = run("diff", "-Doc", str(COPY), "-Since", "with-badge")["result"]
    ok &= check("diff sees no text change from the badge", d["summary"]["textChanges"] == 0, d["summary"])

    print("== P4 -NoSlideMark takes an existing badge away ==")
    r = run("hint", "-Doc", str(COPY), "-NoSlideMark", "-Quiet")["result"]
    ok &= check("-NoSlideMark stamps nothing", r.get("slideMarked") is False, r.get("slideMarked"))
    ok &= check("read has no badge after -NoSlideMark", not read_slide(1).get("takeoverBadge"))

    print("== P5 a save we DO control strips the badge ==")
    run("hint", "-Doc", str(COPY), "-SlideMark", "-Quiet")
    run("save", "-Doc", str(COPY))
    run("close", "-Doc", str(COPY))
    r = run("open", "-Doc", str(COPY), "-NoHint")["result"]
    ok &= check("reopened the saved copy", r.get("slides") == slides, r.get("slides"))
    ok &= check("saved file carries no badge", not read_slide(1).get("takeoverBadge"))
    ok &= check("saved file has no mark shape", not any(sh["name"] == MARK for sh in read_slide(1)["shapes"]))

    print("== P6 a manual Ctrl+S is repaired by the next command ==")
    run("hint", "-Doc", str(COPY), "-SlideMark", "-Quiet")
    ok &= check("raw COM save succeeded (user pressed Ctrl+S)", raw_save() == "SAVED")
    r = run("read", "-Doc", str(COPY), "-From", "1", "-Count", "1")["result"]
    ok &= check("next command reports badgeRepaired", bool(r.get("badgeRepaired")), r.get("badgeRepaired"))
    ok &= check("badge still on screen after the repair", r["items"][0].get("takeoverBadge") is True)
    run("close", "-Doc", str(COPY))
    run("open", "-Doc", str(COPY), "-NoHint")
    ok &= check("repaired file on disk has no badge", not read_slide(1).get("takeoverBadge"))
    ok &= check("repaired file has no mark shape", not any(sh["name"] == MARK for sh in read_slide(1)["shapes"]))

    print("== P7 hint -Clear removes the badge ==")
    run("hint", "-Doc", str(COPY), "-SlideMark", "-Quiet")
    r = run("hint", "-Doc", str(COPY), "-Clear", "-Quiet")["result"]
    ok &= check("-Clear reports no badge left", r.get("slideMarked") is False, r.get("slideMarked"))
    ok &= check("deck has no badge after -Clear", not read_slide(1).get("takeoverBadge"))

    run("close", "-Doc", str(COPY))
    print("RESULT:", "ALL PASS" if ok else "FAILURES PRESENT")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
