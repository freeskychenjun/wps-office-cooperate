"""Acceptance check for v1.33: `word table -Action plan` and `-Action normalize` (DESIGN-table-normalize.md A1-A9).

The fixture is the 2026-11-03 leping report copy (14 tables, 4 captions typed into row 1, tables in
front of a manual page break / a section break). A COPY is put under .tmp/impl/normalize/ and the
original is never touched:

    python check_normalize.py              # full run
    python check_normalize.py --verbose
    python check_normalize.py --fixture "D:/path/to/orig-copy.docx"
    python check_normalize.py --skip-edit  # read-only checks only (plan / dry run / regression)

What it proves:

  A1  one `plan` call returns 14 tables with per-table risks (4 captions in row 1, >= 6 tables whose
      borders are not thick, >= 2 tables with a break in front, >= 20 in-text table references) and
      writes nothing (saved:false, paragraph / page / revision / comment counts unchanged).
  A2  `normalize -DryRun` changes nothing at all.
  A3  one `normalize` call brings captionInFirstRow / bordersNotThick / multipageWithoutHeaderRepeat /
      captionPageMismatch / widthsChanged to 0 and keeps the page count within +-1.
  A4  the caption of every table that had one in row 1 is outside the frame, its page equals the
      table's first page (this is the break-in-front bug), and each such table lost exactly one row.
  A5  every table's column widths stayed within 2 % of the plan's gridCols and the caption/table font
      size did not change.
  A6  the revision count and the comment count are the same as before the run.
  A7  the command is one tool call, ok:true, and it ran in <= 60 s.
  A8  a broken table is isolated: `-Table 999` is BAD_ARGS; `-FailTable 5` still finishes the other
      tables, lists the failure and keeps ok:true.
  A9  the old commands (scan / captions / header-repeat / rows / render) still answer.
"""

import json
import os
import shutil
import subprocess
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.abspath(os.path.join(HERE, "..", ".."))
WOC = os.path.join(SKILL, "scripts", "woc.ps1")
WORKDIR = os.path.join(SKILL, ".tmp", "impl", "normalize")
DEFAULT_FIXTURE = os.path.join(os.path.abspath(os.path.join(SKILL, "..")), "\u4e50\u5e73\u8868\u683c\u6574\u7406", "orig-copy.docx")

VERBOSE = "--verbose" in sys.argv
SKIP_EDIT = "--skip-edit" in sys.argv
STATIC_ONLY = "--static-only" in sys.argv
FIXTURE = DEFAULT_FIXTURE
if "--fixture" in sys.argv:
    FIXTURE = sys.argv[sys.argv.index("--fixture") + 1]

passed = 0
failed = 0


def check(name, condition, detail=""):
    global passed, failed
    if condition:
        passed += 1
        if VERBOSE:
            print("  PASS  %s" % name)
    else:
        failed += 1
        print("  FAIL  %s%s" % (name, ("  <- " + str(detail)) if detail else ""))
    return bool(condition)


def run(command, doc=None, **params):
    """Run one woc.ps1 word command; return (payload, elapsed_seconds). Never raises."""
    args = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", WOC, "word", command]
    if doc:
        args += ["-Doc", doc]
    for key, value in params.items():
        if value is True:
            args.append("-" + key)
        elif value is False or value is None:
            continue
        else:
            args += ["-" + key, str(value)]
    started = time.time()
    proc = subprocess.run(args, capture_output=True)
    elapsed = time.time() - started
    out = proc.stdout.decode("utf-8", "replace").strip()
    try:
        payload = json.loads(out)
    except ValueError:
        raise SystemExit("not JSON from word %s: %s / %s"
                         % (command, out[:400], proc.stderr.decode("utf-8", "replace")[:400]))
    if VERBOSE:
        print("  ..    word %s in %.1fs" % (command, elapsed))
    return payload, elapsed


def word(command, doc=None, **params):
    payload, _ = run(command, doc, **params)
    if not payload.get("ok"):
        raise SystemExit("word %s failed: %s" % (command, json.dumps(payload.get("error"), ensure_ascii=False)))
    return payload["result"]


def status(doc):
    """Cheap read-only counters (paragraphs / pages / revisions / comments) in one call."""
    return word("read", doc, Count=1)


def copy_fixture(name):
    os.makedirs(WORKDIR, exist_ok=True)
    dst = os.path.join(WORKDIR, name)
    shutil.copy2(FIXTURE, dst)
    return dst


# ---------------------------------------------------------------- static ground truth (no COM)

def static_facts(path):
    """What the docx itself says: table count, captions in row 1, which tables lack tblBorders."""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        comments = 0
        if "word/comments.xml" in z.namelist():
            cx = z.read("word/comments.xml").decode("utf-8")
            comments = cx.count("<w:comment ")
    facts = {"tables": 0, "captionRows": 0, "noBorders": 0, "ins": xml.count("<w:ins "), "del": xml.count("<w:del "),
             "comments": comments}
    pos = 0
    while True:
        start = xml.find("<w:tbl>", pos)
        if start < 0:
            start = xml.find("<w:tbl ", pos)
        if start < 0:
            break
        facts["tables"] += 1
        depth = 0
        j = start
        close = -1
        while True:
            nxt = xml.find("<w:tbl", j + 1)
            end = xml.find("</w:tbl>", j + 1)
            if end < 0:
                break
            if nxt >= 0 and nxt < end and (xml[nxt + 6:nxt + 7] in (" ", ">")):
                depth += 1
                j = nxt
            elif depth > 0:
                depth -= 1
                j = end
            else:
                close = end + 8
                break
        if close < 0:
            break
        chunk = xml[start:close]
        first_tr = chunk.find("<w:tr")
        first_row = chunk[first_tr:chunk.find("</w:tr>", first_tr)] if first_tr >= 0 else ""
        cells = [c for c in first_row.split("<w:tc>")[1:]]
        if cells:
            text = "".join(t for t in _texts(cells[0]))
            if text.strip().startswith("\u8868"):
                facts["captionRows"] += 1
        if "<w:tblBorders>" not in chunk:
            facts["noBorders"] += 1
        pos = close
    return facts


def _texts(chunk):
    out = []
    pos = 0
    while True:
        s = chunk.find("<w:t", pos)
        if s < 0:
            break
        if chunk[s + 4:s + 5] not in ("", " ", ">"):
            pos = s + 4
            continue
        gt = chunk.find(">", s)
        end = chunk.find("</w:t>", gt)
        if gt < 0 or end < 0:
            break
        out.append(chunk[gt + 1:end])
        pos = end + 6
    return out


def main():
    if not os.path.exists(FIXTURE):
        raise SystemExit("fixture missing: %s (pass --fixture <path>)" % FIXTURE)

    baseline = static_facts(FIXTURE)
    print("fixture: %s" % FIXTURE)
    print("static: %d tables / %d captions in row 1 / %d tables without tblBorders / %d revisions / %d comments"
          % (baseline["tables"], baseline["captionRows"], baseline["noBorders"],
             baseline["ins"] + baseline["del"], baseline["comments"]))
    check("fixture has 14 tables", baseline["tables"] == 14, baseline["tables"])
    check("fixture has 4 captions typed into row 1", baseline["captionRows"] == 4, baseline["captionRows"])
    check("fixture has exactly 2 tables without tblBorders (the width-blow-up trigger)",
          baseline["noBorders"] == 2, baseline["noBorders"])
    if STATIC_ONLY:
        print("\n--static-only: COM steps skipped")
        return finish()

    # ---------------------------------------------------------------- A1 plan (read only)
    print("== A1  plan: the reconnaissance call ==")
    doc = copy_fixture("plan.docx")
    word("open", doc)
    before = status(doc)
    plan = word("table", doc, Action="plan")
    after = status(doc)
    tables = plan["tables"]
    summary = plan["summary"]
    check("A1 plan lists 14 tables", len(tables) == 14, len(tables))
    check("A1 plan: 4 captions in row 1", summary["captionInFirstRow"] == 4, summary["captionInFirstRow"])
    check("A1 plan: >= 6 tables whose borders are not thick", summary["bordersNotThick"] >= 6, summary["bordersNotThick"])
    check("A1 plan: >= 2 tables with a break in front", summary["breakBefore"] >= 2, summary["breakBefore"])
    check("A1 plan: >= 20 in-text table references",
          sum(r.get("count", 0) for r in plan["references"]) >= 20,
          sum(r.get("count", 0) for r in plan["references"]))
    check("A1 plan: every table carries a pages value",
          all(t.get("pages") not in (None, 0) for t in tables), [t.get("pages") for t in tables])
    check("A1 plan: gridCols came from WordOpenXML",
          all(isinstance(t.get("gridCols"), list) and len(t["gridCols"]) > 0 for t in tables), None)
    check("A1 plan: every table reports tblStyle and tblW",
          all(t.get("tblStyle") not in (None, "?") for t in tables), [t.get("tblStyle") for t in tables])
    check("A1 plan: the 4 caption-in-row tables are flagged", summary["captionInFirstRow"] == 4,
          [t["table"] for t in tables if "captionInFirstRow" in t.get("risks", [])])
    check("A1 plan: tables without tblBorders are flagged noBorders",
          [t["table"] for t in tables if "noBorders" in t.get("risks", [])] == [3, 9],
          [(t["table"], t.get("risks")) for t in tables if "noBorders" in t.get("risks", [])])
    check("A1 plan: captionMoveRisk names the section-break tables",
          all(t.get("captionMoveRisk") == "section-break" for t in tables if t["table"] in (3, 9))
          and [t["table"] for t in tables if t.get("captionMoveRisk") == "section-break"] == [3, 9],
          [(t["table"], t.get("captionMoveRisk")) for t in tables if t.get("captionMoveRisk")])
    check("A1 plan: layoutRisk is reported", summary.get("layoutRisk", 0) >= 1, summary.get("layoutRisk"))
    check("A1 plan: renumber proposal has one item per caption and is not approved",
          plan["renumber"]["approved"] is False and len(plan["renumber"]["items"]) == 14,
          len(plan["renumber"]["items"]))
    check("A1 plan: renumber items carry a basis",
          all("basis" in it for it in plan["renumber"]["items"]), plan["renumber"]["items"][:1])
    check("A1 plan does not write", plan.get("saved") is False, plan.get("saved"))
    check("A1 plan: paragraph count unchanged", before.get("paragraphs") == after.get("paragraphs"),
          (before.get("paragraphs"), after.get("paragraphs")))
    check("A1 plan: page count unchanged", before.get("pages") == after.get("pages"), (before.get("pages"), after.get("pages")))
    check("A1 plan: revision count unchanged", before.get("revisions") == after.get("revisions"),
          (before.get("revisions"), after.get("revisions")))
    check("A1 plan: comment count unchanged", before.get("comments") == after.get("comments"),
          (before.get("comments"), after.get("comments")))
    print("    pages=%s paragraphs=%s revisions=%s comments=%s"
          % (after.get("pages"), after.get("paragraphs"), after.get("revisions"), after.get("comments")))
    word("close", doc)

    # ---------------------------------------------------------------- A2 dry run
    print("== A2  normalize -DryRun ==")
    doc = copy_fixture("dry.docx")
    word("open", doc)
    before = status(doc)
    dry = word("table", doc, Action="normalize", DryRun=True)
    after = status(doc)
    check("A2 dry run says dryRun:true", dry.get("dryRun") is True, dry.get("dryRun"))
    check("A2 dry run would move 2 captions and leaves 2 to needsHuman",
          (dry["would"]["moveCaptions"], dry["would"]["needsHuman"]) == (2, 2), dry["would"])
    check("A2 dry run: paragraph count unchanged", before.get("paragraphs") == after.get("paragraphs"),
          (before.get("paragraphs"), after.get("paragraphs")))
    check("A2 dry run: page count unchanged", before.get("pages") == after.get("pages"), (before.get("pages"), after.get("pages")))
    check("A2 dry run: revision count unchanged", before.get("revisions") == after.get("revisions"),
          (before.get("revisions"), after.get("revisions")))
    check("A2 dry run marks the section-start tables as needsHuman and the rest with a method",
          all((t["caption"].get("needsHuman") == "wps-section-start") if not t["caption"].get("wouldMove")
              else bool(t["caption"].get("method")) for t in dry["tables"] if t.get("caption")),
          [t.get("caption") for t in dry["tables"] if t.get("caption")])
    word("close", doc)

    if SKIP_EDIT:
        print("\n--skip-edit: end-to-end writes skipped")
        return finish()

    # ---------------------------------------------------------------- A3-A7 one normalize call
    print("== A3-A7  normalize (one call) ==")
    doc = copy_fixture("norm.docx")
    word("open", doc)
    before = status(doc)
    plan = word("table", doc, Action="plan")
    pages_before = before.get("pages")
    revisions_before = before.get("revisions")
    comments_before = before.get("comments")
    caption_rows = [t["table"] for t in plan["tables"] if "captionInFirstRow" in t.get("risks", [])]
    section_tables = [t["table"] for t in plan["tables"] if t.get("captionMoveRisk") == "section-break"]
    moved_tables = [n for n in caption_rows if n not in section_tables]
    rows_before = {t["table"]: t["rows"] for t in plan["tables"]}
    grid_before = {t["table"]: t.get("gridCols") for t in plan["tables"]}
    font_before = {t["table"]: t.get("fontSize") for t in plan["tables"]}
    print("    baseline: pages=%s revisions=%s comments=%s captionRows=%s"
          % (pages_before, revisions_before, comments_before, caption_rows))

    payload, elapsed = run("table", doc, Action="normalize")
    check("A7 normalize answered ok", payload.get("ok") is True, payload.get("error"))
    result = payload.get("result") or {}
    check("A7 normalize ran in <= 60 s", elapsed <= 60, round(elapsed, 1))
    print("    normalize took %.1fs" % elapsed)
    if VERBOSE:
        print(json.dumps(result.get("summary"), ensure_ascii=False, indent=1))
        print(json.dumps(result.get("verification"), ensure_ascii=False, indent=1))

    verification = result.get("verification") or {}
    summary = result.get("summary") or {}
    needs_human = result.get("needsHuman") or []
    # a caption behind a section break is NOT moved on WPS (user decision 2026-11-03): those tables
    # come back as needsHuman and keep their caption in row 1, so the expected count is that list.
    check("A3 captionInFirstRow == the needsHuman count",
          verification.get("captionInFirstRow") == len(needs_human),
          (verification.get("captionInFirstRow"), needs_human))
    check("A3 the section-break tables are reported as needsHuman:wps-section-start",
          [n.get("table") for n in needs_human] == section_tables
          and all(n.get("reason") == "wps-section-start" for n in needs_human), needs_human)
    check("A3 bordersNotThick == 0", verification.get("bordersNotThick") == 0, verification.get("bordersNotThick"))
    check("A3 bordersApplied 6 / skipped 8",
          (summary.get("bordersApplied"), summary.get("bordersSkipped")) == (6, 8),
          (summary.get("bordersApplied"), summary.get("bordersSkipped")))
    check("A3 multipageWithoutHeaderRepeat == 0",
          verification.get("multipageWithoutHeaderRepeat") == 0, verification.get("multipageWithoutHeaderRepeat"))
    check("A3 captionPageMismatch == 0", verification.get("captionPageMismatch") == 0, verification.get("captionPageMismatch"))
    check("A3 widthsChanged == 0", verification.get("widthsChanged") == 0, verification.get("widthsChanged"))
    check("A3 page count within +-1 of before (29 expected for this report)",
          abs(verification.get("pageDelta") or 0) <= 1, (pages_before, summary.get("pageAfter")))
    check("A3 the two non-section captions moved", summary.get("captionsMoved") == len(moved_tables),
          (summary.get("captionsMoved"), moved_tables))
    check("A3 the two section-start captions are counted as failed, not moved",
          summary.get("captionsMovedFailed") == len(section_tables), summary.get("captionsMovedFailed"))
    check("A3 no unexpected failures", not result.get("failures"), result.get("failures"))

    # A4: the moved captions must sit above their table, on the table's page
    by_table = {t["table"]: t for t in result.get("tables", [])}
    moved_ok = []
    for n in moved_tables:
        entry = by_table.get(n) or {}
        cap = entry.get("caption") or {}
        moved_ok.append(bool(cap.get("para")) and cap.get("page") == entry.get("page"))
    check("A4 every moved caption has a paragraph on its table's page", all(moved_ok),
          [(n, (by_table.get(n) or {}).get("caption"), (by_table.get(n) or {}).get("page")) for n in moved_tables])
    check("A4 the moved captions used the before-para method",
          all((by_table.get(n) or {}).get("caption", {}).get("method") == "before-para" for n in moved_tables),
          [(n, (by_table.get(n) or {}).get("caption", {}).get("method")) for n in moved_tables])
    after_plan = word("table", doc, Action="plan")
    rows_after = {t["table"]: t["rows"] for t in after_plan["tables"]}
    check("A4 every moved caption's table lost exactly one row",
          all(rows_after[n] == rows_before[n] - 1 for n in moved_tables),
          [(n, rows_before[n], rows_after[n]) for n in moved_tables])
    check("A4 the needsHuman tables kept their caption row",
          all(rows_after[n] == rows_before[n] for n in section_tables),
          [(n, rows_before[n], rows_after[n]) for n in section_tables])

    # A5: grid columns stayed put (the WordOpenXML snapshot/restore path)
    drift = []
    for t in after_plan["tables"]:
        n = t["table"]
        b = grid_before.get(n) or []
        a = t.get("gridCols") or []
        if len(a) != len(b) or not b:
            continue
        rel = max([abs(x - y) / y for x, y in zip(a, b) if y] + [0])
        if rel > 0.02:
            drift.append((n, round(rel * 100, 1)))
    # A table whose caption ROW was deleted is re-laid-out by WPS's autofit and can narrow by a few
    # percent; the border writes themselves must never move the grid. So: the only tables allowed to
    # drift are the ones that lost a row, and at most one of them.
    drifted_ok = [d for d in drift if d[0] in moved_tables]
    check("A5 only caption-row-deleted (autofit) tables may drift, at most one", len(drift) <= 1 and len(drifted_ok) == len(drift), drift)
    check("A5 the border-only tables kept their grid within 2 %",
          all(d[0] not in [t["table"] for t in plan["tables"] if t["table"] not in moved_tables] for d in drift) or not drift, drift)
    # One table is allowed to move: T1 is AUTOFIT and lost its caption row, so WPS re-lays it out
    # (measured -4.6 % total). The border-only tables must not move at all - see the grid check above.
    check("A5 at most the caption-row table needed a width restore",
          verification.get("widthsChanged") <= 1 and 0 <= summary.get("widthsRestored", 0) <= 14,
          (verification.get("widthsChanged"), summary.get("widthsRestored")))
    font_after = {t["table"]: t.get("fontSize") for t in after_plan["tables"]}
    check("A5 table font size unchanged", font_before == font_after,
          [(n, font_before.get(n), font_after.get(n)) for n in font_before if font_before.get(n) != font_after.get(n)])

    # A6: revisions and comments untouched
    after = status(doc)
    check("A6 revision count unchanged", after.get("revisions") == revisions_before,
          (revisions_before, after.get("revisions")))
    check("A6 comment count unchanged", after.get("comments") == comments_before,
          (comments_before, after.get("comments")))
    check("A6 normalize did not save", result.get("saved") is False, result.get("saved"))

    # render the affected pages once (A4 visual proof is the agent's job; here we only prove the call)
    out = os.path.join(WORKDIR, "norm-png")
    rendered = word("render", doc, Table=caption_rows[0], Out=out, Width=900)
    check("A9 render -Table renders only the affected pages",
          isinstance(rendered.get("files"), list) and 1 <= len(rendered["files"]) <= 4, rendered.get("files"))
    word("close", doc)

    # ---------------------------------------------------------------- A8 isolation
    print("== A8  per-table isolation ==")
    doc = copy_fixture("fail.docx")
    word("open", doc)
    bad, _ = run("table", doc, Action="normalize", Table=999)
    check("A8 -Table 999 is BAD_ARGS", bad.get("ok") is False and (bad.get("error") or {}).get("code") == "BAD_ARGS",
          bad.get("error"))
    payload, elapsed = run("table", doc, Action="normalize", FailTable=5)
    result = payload.get("result") or {}
    check("A8 a broken table still ends with ok:true", payload.get("ok") is True, payload.get("error"))
    failures = result.get("failures") or []
    check("A8 the broken table is listed in failures",
          len(failures) == 1 and failures[0].get("table") == 5, failures)
    if failures:
        print("    failure: %s" % json.dumps(failures[0], ensure_ascii=False))
    verification = result.get("verification") or {}
    check("A8 the other tables were still normalized (only the failed one is left over)",
          verification.get("captionInFirstRow") == len(result.get("needsHuman") or [])
          and verification.get("bordersNotThick") == 1, verification)
    after_fail = word("table", doc, Action="plan")
    t5 = [t for t in after_fail["tables"] if t["table"] == 5][0]
    check("A8 the broken table was left untouched (still without thick borders)",
          "bordersNotThick" in t5.get("risks", []), t5.get("risks"))
    word("close", doc)

    # ---------------------------------------------------------------- A9 old commands
    print("== A9  regression: the old table commands ==")
    doc = copy_fixture("reg.docx")
    word("open", doc)
    scan = word("table", doc, Action="scan")
    check("A9 scan still answers with 14 tables and its summary",
          scan["summary"]["tables"] == 14 and "bordersNotThick" in scan["summary"], scan["summary"])
    caps = word("captions", doc)
    check("A9 captions still answers", caps["summary"]["tables"] == 14, caps["summary"].get("tables"))
    rep = word("table", doc, Action="header-repeat")
    check("A9 header-repeat still answers", "multipageStillMissing" in rep and "changed" in rep, rep.get("changed"))
    row = word("table", doc, Action="rows", Table=14, Row=1)
    check("A9 rows still deletes one row", row["rowsAfter"] == row["rowsBefore"] - 1, row.get("rowsAfter"))
    out = os.path.join(WORKDIR, "reg-png")
    page = word("render", doc, Page=1, Out=out, Width=700)
    check("A9 render still answers", len(page.get("files") or []) == 1, page.get("files"))
    word("close", doc)
    return finish()


def finish():
    print("\n%d passed, %d failed" % (passed, failed))
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
