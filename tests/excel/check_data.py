#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Independent oracle for the Excel data-layer regression driver (tests/excel).

    python check_data.py --work <work-dir> [--files <dir>] [--skip-file] [--skip-reason TEXT]

The driver (run_data.ps1) only starts WPS, copies the pristine fixture, calls the skill and
stores every raw JSON response under <work>/logs/.  This file decides whether the run was
correct:

  * ground truth is the PRISTINE fixture (<work>/base.xlsx), read with openpyxl - never a
    workbook WPS has touched;
  * every expectation is re-derived here from that fixture (sort order, dedupe survivors,
    clean transformations, fill formula series ...) with plain python;
  * the saved workbooks are read back with openpyxl (formulas with data_only=False, the
    values WPS cached with data_only=True); chart presence comes from the drawing part.
    On this machine WPS writes workbooks transparently encrypted, so the driver decrypts
    the saved files into <work>/plain first and points --files there; an undecryptable file
    is reported as a FAIL naming the encryption header instead of a stack trace.

A case that was not run (subset -Only, or a driver that died before it) is printed as SKIP
with the reason, never silently dropped.  Every line is PASS / FAIL / SKIP; the exit code is
1 when anything failed.
"""
import argparse
import json
import re
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import openpyxl
from openpyxl.utils import get_column_letter

ENCRYPTED_MAGIC = bytes.fromhex("17DA5FA01633CD9A")

WORK = None          # work dir (base.xlsx, logs/, cases/, plain/)
FILES = None         # dir holding the decrypted per-case folders (None = file oracle off)
SKIP_REASON = ""
COUNTS = {"pass": 0, "fail": 0, "skip": 0}
FAILED = []          # (case, label) for the final recap


# ---------------------------------------------------------------- reporting

def check(cid, label, cond, detail=""):
    if cond:
        COUNTS["pass"] += 1
        print("  PASS  %s" % label)
    else:
        COUNTS["fail"] += 1
        FAILED.append((cid, label))
        text = "  FAIL  %s" % label
        if detail:
            text += "  <- " + str(detail)
        print(text)
    return bool(cond)


def skip(cid, label, reason):
    COUNTS["skip"] += 1
    print("  SKIP  %s  (%s)" % (label, reason))


def file_skip(label):
    """The file oracle is off: say so once per check instead of pretending to pass."""
    skip("files", label, "file oracle disabled: " + (SKIP_REASON or "no decrypted files"))


# ---------------------------------------------------------------- comparison helpers

def norm(v):
    """One comparison vocabulary for WPS values and openpyxl values."""
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, str):
        s = v.strip()
        return None if s == "" else s
    if isinstance(v, (int, float)):
        f = float(v)
        if f.is_integer():
            return int(f)
        return round(f, 10)
    return v


def same(a, b):
    return norm(a) == norm(b)


def show(v):
    return repr(v)


# WPS returns the raw chart-type code when a chart is listed (v1.20 helper: ChartTypeName only
# maps doubles, WPS hands back an int, so "4"/"51" come out as strings while `chart -Action add`
# reports the friendly name). Map both spellings so the assertion does not depend on the form;
# the mismatch itself is reported as a finding for the skill, not hidden by this mapping.
CHART_TYPE_NAMES = {
    "4": "line", "51": "column", "5": "pie", "57": "bar", "52": "stacked", "58": "stackedbar",
    "-4120": "doughnut", "-4169": "scatter", "74": "scatter", "76": "area", "-4151": "radar",
}


def chart_type(value):
    text = str(value).strip()
    return CHART_TYPE_NAMES.get(text, text)


def block(ws, r1, c1, r2, c2):
    """Values of a rectangle, rows of normalized values."""
    out = []
    for r in range(r1, r2 + 1):
        out.append([norm(ws.cell(row=r, column=c).value) for c in range(c1, c2 + 1)])
    return out


def formula(ws, row, col):
    return ws.cell(row=row, column=col).value


def check_block(cid, label, ws, expected, r1=1, c1=1):
    """Compare a rectangle of the workbook with expected rows starting at (r1, c1)."""
    problems = []
    for i, want_row in enumerate(expected):
        for j, want in enumerate(want_row):
            got = ws.cell(row=r1 + i, column=c1 + j).value
            if not same(got, want):
                addr = "%s%d" % (get_column_letter(c1 + j), r1 + i)
                problems.append("%s got %s want %s" % (addr, show(got), show(want)))
    ok = not problems
    check(cid, label, ok, "; ".join(problems[:6]) if problems else "")
    return ok


def check_cell(cid, ws, row, col, want, label=None):
    got = ws.cell(row=row, column=col).value
    lab = label or ("%s%d == %s" % (get_column_letter(col), row, show(want)))
    check(cid, lab, same(got, want), "got %s" % show(got))
    return same(got, want)


def check_formula(cid, ws, row, col, want, label=None):
    got = ws.cell(row=row, column=col).value
    lab = label or ("%s%d formula %s" % (get_column_letter(col), row, want))
    check(cid, lab, isinstance(got, str) and got.strip() == want, "got %s" % show(got))
    return isinstance(got, str) and got.strip() == want


def check_identical(cid, label, ws_a, ws_b, max_row, max_col):
    """Two sheets hold the same values AND formulas over the given rectangle."""
    problems = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            a = ws_a.cell(row=r, column=c).value
            b = ws_b.cell(row=r, column=c).value
            if not same(a, b):
                problems.append("%s%d %s vs %s" % (get_column_letter(c), r, show(a), show(b)))
    check(cid, label, not problems, "; ".join(problems[:6]))


# ---------------------------------------------------------------- workspace access

def read_json(path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


class Book:
    """A saved workbook plus the same file opened for cached values."""

    def __init__(self, wb, values, path):
        self.wb = wb
        self.values = values
        self.path = path

    def ws(self, name):
        return self.wb[name]

    def vals(self, name):
        return self.values[name]


def load_book(cid, rel_name):
    """Open <files>/<cid>/<rel_name>; FAIL (once) with a readable reason and return None."""
    if FILES is None:
        file_skip("%s/%s" % (cid, rel_name))
        return None
    path = Path(FILES) / cid / rel_name
    if not path.exists():
        check(cid, "saved workbook %s exists" % rel_name, False,
              "not found under %s (decryption or the case itself failed)" % (Path(FILES) / cid))
        return None
    head = path.read_bytes()[:8]
    if head.startswith(ENCRYPTED_MAGIC):
        check(cid, "saved workbook %s is readable" % rel_name, False,
              "still transparently encrypted (header %s): decryption failed" % head.hex().upper())
        return None
    try:
        wb = openpyxl.load_workbook(str(path))
        vals = openpyxl.load_workbook(str(path), data_only=True)
    except Exception as exc:  # noqa: BLE001 - any unreadable file is a finding
        check(cid, "saved workbook %s is readable" % rel_name, False, "%s: %s" % (type(exc).__name__, exc))
        return None
    check(cid, "saved workbook %s is readable" % rel_name, True)
    return Book(wb, vals, path)


def step_ms(cid, step):
    """Wall time the driver measured for one skill call (0 = not recorded)."""
    path = Path(WORK) / "logs" / ("%s-%s.ms.txt" % (cid, step))
    try:
        return int(path.read_text(encoding="utf-8").strip())
    except Exception:
        return 0


def resp(cid, step):
    path = Path(WORK) / "logs" / ("%s-%s.json" % (cid, step))
    if not path.exists():
        return None
    data = read_json(path)
    if not isinstance(data, dict):
        return None
    return data


def need(cid, step):
    data = resp(cid, step)
    if data is None:
        check(cid, "response %s present" % step, False, "logs/%s-%s.json missing or not JSON" % (cid, step))
    return data


def need_ok(cid, step):
    data = need(cid, step)
    if data is None:
        return None
    err = data.get("error") or {}
    check(cid, "%s ok:true" % step, data.get("ok") is True,
          (err.get("code", "") + " " + err.get("message", "")).strip() if err else "")
    return data.get("result") or {}


def result_of(cid, step):
    data = resp(cid, step)
    if not isinstance(data, dict):
        return {}
    return data.get("result") or {}


# ---------------------------------------------------------------- base fixture access

def base_sheet_rows(base, name):
    ws = base[name]
    return [[norm(c) for c in row] for row in ws.iter_rows(values_only=True)]


# ---------------------------------------------------------------- the cases
#
# Each function receives the pristine base workbook (openpyxl) and asserts one scenario.

def case_01_write(base):
    cid = "01-write"
    print("== %s: write - values, formula, values block ==" % cid)
    r = need_ok(cid, "write")
    if r:
        check(cid, "write reports 3 applied items", r.get("count") == 3, "count=%s" % r.get("count"))
        applied = r.get("applied") or []
        cells = [a.get("cell") for a in applied if isinstance(a, dict)]
        check(cid, "applied cells are E1, E2, F2:G3", cells == ["E1", "E2", "F2:G3"], cells)
        before = [a.get("before") for a in applied if isinstance(a, dict)]
        check(cid, "every applied cell started empty", all((b or "") == "" for b in before), before)
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Data")
    wsv = book.vals("Data")
    check_cell(cid, ws, 1, 5, "computed", "E1 holds the written text")
    check_formula(cid, ws, 2, 5, "=AVERAGE(B2:B4)", "E2 holds the formula")
    avg = (norm(base["Data"].cell(row=2, column=2).value)
           + norm(base["Data"].cell(row=3, column=2).value)
           + norm(base["Data"].cell(row=4, column=2).value)) / 3.0
    check(cid, "E2 cached value is the average (%.4g)" % avg,
          same(wsv.cell(row=2, column=5).value, avg), "got %s" % show(wsv.cell(row=2, column=5).value))
    check_block(cid, "F2:G3 holds the written 2x2 block", ws, [[1, 2], [3, 4]], r1=2, c1=6)
    # The rest of Data must be exactly the fixture.
    check_identical(cid, "Data rows 1..31 unchanged apart from the written cells",
                    ws, base["Data"], 31, 4)


def case_02_write_anchor(base):
    cid = "02-write-anchor"
    print("== %s: write - expected-text mismatch ==" % cid)
    data = need(cid, "write")
    if data is not None:
        err = data.get("error") or {}
        check(cid, "write ok:false", data.get("ok") is False, data.get("ok"))
        check(cid, "error code is ANCHOR_MISSING", err.get("code") == "ANCHOR_MISSING", err.get("code"))
        check(cid, "message names the mismatching cell H1", "H1" in str(err.get("message", "")), err.get("message"))
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Data")
    check_cell(cid, ws, 2, 2, 10, "B2 is still the fixture value (nothing written)")
    check(cid, "H1 is still empty", ws.cell(row=1, column=8).value in (None, ""),
          show(ws.cell(row=1, column=8).value))
    check_identical(cid, "the whole Data sheet still equals the fixture", ws, base["Data"], 31, 4)


def case_03_fill(base):
    cid = "03-fill"
    print("== %s: fill - relative formula down 1999 rows ==" % cid)
    r = need_ok(cid, "fill")
    if r:
        check(cid, "fill reports 1999 rows", r.get("rows") == 1999, r.get("rows"))
        check(cid, "topAfter is 3", str(norm(r.get("topAfter"))) == "3", r.get("topAfter"))
        check(cid, "bottomAfter is 3999", str(norm(r.get("bottomAfter"))) == "3999", r.get("bottomAfter"))
        check(cid, "no error cells", not (r.get("errors") or []), r.get("errors"))
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Fill")
    wsv = book.vals("Fill")
    check_cell(cid, ws, 1, 2, "y", "B1 header is still 'y'")
    check_cell(cid, ws, 2, 1, 1, "A2 is still the fixture value")
    check_cell(cid, ws, 2000, 1, 1999, "A2000 is still the fixture value")
    check_formula(cid, ws, 2, 2, "=A2*2+1", "B2 formula")
    check_formula(cid, ws, 3, 2, "=A3*2+1", "B3 formula (relative, one step down)")
    check_formula(cid, ws, 2000, 2, "=A2000*2+1", "B2000 formula (relative, last row)")
    check(cid, "B2 cached value is 3", same(wsv.cell(row=2, column=2).value, 3),
          show(wsv.cell(row=2, column=2).value))
    check(cid, "B2000 cached value is 3999", same(wsv.cell(row=2000, column=2).value, 3999),
          show(wsv.cell(row=2000, column=2).value))


def case_04_replace(base):
    cid = "04-replace"
    print("== %s: replace - dry-run, single pass, numeric/formula guards, idempotency ==" % cid)
    r = need_ok(cid, "dryrun-number")
    if r:
        check(cid, "numeric dry run: dryRun true, engine memory-block",
              r.get("dryRun") is True and r.get("engine") == "memory-block", (r.get("dryRun"), r.get("engine")))
        check(cid, "numeric dry run: changed 0", r.get("changed") == 0, r.get("changed"))
        check(cid, "numeric dry run: all 5 numeric matches left alone",
              r.get("numericMatchesLeftAlone") == 5, r.get("numericMatchesLeftAlone"))
    r = need_ok(cid, "formula-guard")
    if r:
        check(cid, "formula guard: applied 0", r.get("applied") == 0, r.get("applied"))
        check(cid, "formula guard: at least one formula cell skipped",
              (r.get("formulaCellsSkipped") or 0) >= 1, r.get("formulaCellsSkipped"))
    r = need_ok(cid, "dryrun")
    if r:
        check(cid, "dry run: applied 0, would change 4",
              r.get("applied") == 0 and r.get("changed") == 4, (r.get("applied"), r.get("changed")))
        samples = {(s.get("cell"), s.get("before"), s.get("after")) for s in (r.get("samples") or [])}
        check(cid, "dry run lists PP -> P-P- (single pass)", ("A3", "PP", "P-P-") in samples, sorted(samples)[:6])
    need_ok(cid, "save-dryrun")
    r = need_ok(cid, "apply")
    if r:
        check(cid, "apply: applied and changed are 4", r.get("applied") == 4 and r.get("changed") == 4,
              (r.get("applied"), r.get("changed")))
        samples = {(s.get("cell"), s.get("after")) for s in (r.get("samples") or [])}
        check(cid, "apply samples cover A2/A3/A4/A6",
              {("A2", "P- item"), ("A3", "P-P-"), ("A4", "alpha P- beta"), ("A6", "P-")} <= samples,
              sorted(samples))
    r = need_ok(cid, "case-insensitive")
    if r:
        check(cid, "ALPHA matches the lower-case alpha (default is case-insensitive)",
              r.get("changed") == 1 and r.get("applied") == 0, (r.get("changed"), r.get("applied")))
    r = need_ok(cid, "apply-idem-1")
    if r:
        check(cid, "idempotency step 1 applies 1 cell", r.get("applied") == 1, r.get("applied"))
    r = need_ok(cid, "apply-idem-2")
    if r:
        check(cid, "idempotency step 2 applies nothing", r.get("applied") == 0 and r.get("changed") == 0,
              (r.get("applied"), r.get("changed")))
    need_ok(cid, "save-final")

    dry = load_book(cid, "04-replace-dryrun.xlsx")
    if dry is not None:
        check_identical(cid, "dry-run left the Text sheet exactly as the fixture",
                        dry.ws("Text"), base["Text"], 6, 3)
    final = load_book(cid, "04-replace-final.xlsx")
    if final is None:
        return
    ws = final.ws("Text")
    check_block(cid, "Text column A after the single-pass replace",
                ws,
                [["name"],
                 ["P- item"], ["P-P-"], ["alpha P- beta"], ["no-token-here"], ["P-"]],
                r1=1, c1=1)
    check_block(cid, "Text column B (numbers) is untouched",
                ws, [[c] for c in [norm(base["Text"].cell(row=r, column=2).value) for r in range(1, 7)]],
                r1=1, c1=2)
    check_cell(cid, ws, 2, 3, "0", "C2 text '0' untouched by the alpha replace")
    check_formula(cid, ws, 3, 3, "=B2+B3", "C3 formula untouched")
    check_cell(cid, ws, 5, 3, "omega", "C5 alpha -> omega")
    check_cell(cid, ws, 6, 3, "P", "C6 'P' untouched")


def case_05_sort(base):
    cid = "05-sort"
    print("== %s: sort - descending by qty, formulas travel with rows ==" % cid)
    r = need_ok(cid, "sort")
    if r:
        by = r.get("by") or []
        check(cid, "sort reports qty desc",
              len(by) == 1 and by[0].get("column") == "qty" and by[0].get("order") == "desc", by)
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Sort")
    wsv = book.vals("Sort")
    rows = [[norm(base["Sort"].cell(row=r, column=c).value) for c in (1, 2)] for r in range(2, 10)]
    rows = [row for row in rows if row[0] is not None]
    expected = sorted(rows, key=lambda row: row[1], reverse=True)
    check_cell(cid, ws, 1, 1, "name", "header stays in row 1")
    check_block(cid, "names follow the descending qty order",
                ws, [[row[0]] for row in expected], r1=2, c1=1)
    check_block(cid, "qty values are descending",
                ws, [[row[1]] for row in expected], r1=2, c1=2)
    for i in range(len(expected)):
        row = 2 + i
        check_formula(cid, ws, row, 3, "=B%d*10" % row, "C%d formula followed its row" % row)
        check(cid, "C%d cached value %d" % (row, expected[i][1] * 10),
              same(wsv.cell(row=row, column=3).value, expected[i][1] * 10),
              show(wsv.cell(row=row, column=3).value))


def case_06_deletewhere(base):
    cid = "06-deletewhere"
    print("== %s: deletewhere - rows by text condition ==" % cid)
    r = need_ok(cid, "delete")
    if r:
        check(cid, "3 rows deleted", r.get("deleted") == 3, r.get("deleted"))
        check(cid, "condition names the status column", "status" in str(r.get("condition", "")), r.get("condition"))
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Where")
    rows = [[norm(base["Where"].cell(row=r, column=c).value) for c in (1, 2, 3)] for r in range(2, 10)]
    kept = [row for row in rows if row[1] != "drop"]
    check_block(cid, "surviving rows keep their order", ws, kept, r1=2, c1=1)
    check(cid, "no stale row is left below the survivors",
          ws.cell(row=2 + len(kept), column=1).value in (None, ""),
          show(ws.cell(row=2 + len(kept), column=1).value))


def case_07_fillblanks(base):
    cid = "07-fillblanks"
    print("== %s: fillblanks - above, then a constant ==" % cid)
    r = need_ok(cid, "above")
    if r:
        check(cid, "'above' fills 2 cells", r.get("filled") == 2, r.get("filled"))
        cols = r.get("columns") or []
        check(cid, "'above' works on column B", len(cols) == 1 and cols[0].get("column") == "B", cols)
    r = need_ok(cid, "zero")
    if r:
        check(cid, "'0' fills the remaining 1 cell", r.get("filled") == 1, r.get("filled"))
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Blanks")
    raw = [norm(base["Blanks"].cell(row=r, column=2).value) for r in range(2, 7)]
    carried = []
    last = None
    for value in raw:
        if value is not None:
            last = value
        carried.append(last)
    expected = [0 if value is None else value for value in carried]
    check_block(cid, "qty column: above-fill then the constant 0", ws,
                [[norm(base["Blanks"].cell(row=r, column=1).value), expected[r - 2]] for r in range(2, 7)],
                r1=2, c1=1)
    check_cell(cid, ws, 1, 2, "qty", "header is still 'qty'")


def case_08_dedupe(base):
    cid = "08-dedupe"
    print("== %s: dedupe - first of each site kept ==" % cid)
    r = need_ok(cid, "dedupe")
    if r:
        check(cid, "3 duplicates removed", r.get("removed") == 3, r.get("removed"))
        check(cid, "key is the site column", "site" in str(r.get("key", "")), r.get("key"))
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Dedupe")
    rows = [[norm(base["Dedupe"].cell(row=r, column=c).value) for c in (1, 2)] for r in range(2, 9)]
    seen = set()
    kept = []
    for row in rows:
        if row[0] in seen:
            continue
        seen.add(row[0])
        kept.append(row)
    check_block(cid, "only the first row of each site survives", ws, kept, r1=2, c1=1)


def case_09_clean(base):
    cid = "09-clean"
    print("== %s: clean - trim/case/no-spaces, then text -> number ==" % cid)
    r = need_ok(cid, "upper")
    if r:
        check(cid, "upper+NoSpaces changes 3 name cells", r.get("changed") == 3, r.get("changed"))
    r = need_ok(cid, "number")
    if r:
        check(cid, "to-number changes 4 qty cells", r.get("changed") == 4, r.get("changed"))
        check(cid, "no unreadable value", r.get("notReadable") == 0, r.get("notReadable"))
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Clean")
    names = []
    for r in range(2, 6):
        raw = base["Clean"].cell(row=r, column=1).value
        names.append([re.sub(r"\s+", "", str(raw).strip()).upper()])
    check_block(cid, "name column trimmed, de-spaced and upper-cased", ws, names, r1=2, c1=1)
    numbers = []
    for r in range(2, 6):
        raw = base["Clean"].cell(row=r, column=3).value
        numbers.append([int(float(norm(raw)))])
    check_block(cid, "qty column is numeric now (fixture held text)", ws, numbers, r1=2, c1=3)
    check_cell(cid, ws, 1, 1, "name", "header is still 'name'")


def case_10_table(base):
    cid = "10-table"
    print("== %s: table - create a ListObject, second call is a restyle ==" % cid)
    r = need_ok(cid, "create")
    if r:
        check(cid, "table created", r.get("created") is True, r.get("created"))
        check(cid, "table name TblOne", r.get("name") == "TblOne", r.get("name"))
        check(cid, "table range A1:C7", r.get("range") == "A1:C7", r.get("range"))
    r = need_ok(cid, "again")
    if r:
        check(cid, "second call is not a new table", r.get("created") is False, r.get("created"))
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Table")
    names = list(ws.tables.keys())
    check(cid, "openpyxl sees the table TblOne", "TblOne" in names, names)
    if "TblOne" in names:
        check(cid, "table reference is A1:C7", ws.tables["TblOne"].ref == "A1:C7", ws.tables["TblOne"].ref)
    check_block(cid, "table values are untouched", ws, block(base["Table"], 1, 1, 7, 3))


def case_11_sheet(base):
    cid = "11-sheet"
    print("== %s: sheet - list / add / write / rename / delete ==" % cid)
    r = need_ok(cid, "list")
    if r:
        check(cid, "13 sheets before the case", len(r.get("sheets") or []) == 13, len(r.get("sheets") or []))
    r = need_ok(cid, "add")
    if r:
        check(cid, "add reports NewSheet", r.get("added") == "NewSheet", r.get("added"))
    r = need_ok(cid, "write-new")
    if r:
        check(cid, "writing to the new sheet works", r.get("count") == 1, r.get("count"))
    r = need_ok(cid, "rename")
    if r:
        check(cid, "rename reports Summary -> Summary2",
              r.get("renamed") == "Summary" and r.get("to") == "Summary2", (r.get("renamed"), r.get("to")))
    r = need_ok(cid, "delete")
    if r:
        check(cid, "delete reports NewSheet", r.get("deleted") == "NewSheet", r.get("deleted"))
        check(cid, "delete reports how many used cells the sheet had (non-empty sheet)",
              (r.get("usedCells") or 0) > 0, r.get("usedCells"))
        elapsed = step_ms(cid, "delete")
        # Before v1.34 the Delete() of a non-empty sheet blocked on WPS's modal confirmation
        # forever; the helper now suppresses DisplayAlerts, so a few seconds is the whole budget.
        check(cid, "delete of a non-empty sheet returned quickly", 0 < elapsed < 20000, "%d ms" % elapsed)
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    names = list(book.wb.sheetnames)
    # Summary2 takes the slot Summary had; NewSheet was deleted again.
    expected = [("Summary2" if n == "Summary" else n) for n in base.sheetnames]
    check(cid, "sheet list after add+rename+delete", names == expected, names)
    check(cid, "NewSheet is gone", "NewSheet" not in names, names)
    check_cell(cid, book.ws("Summary2"), 2, 2, 1, "renamed sheet kept its data (Summary!B2 == 1)")


def case_12_rowscols(base):
    cid = "12-rowscols"
    print("== %s: rows/cols - insert, then delete back ==" % cid)
    r = need_ok(cid, "rows-insert")
    if r:
        check(cid, "rows insert reports row/at 2/count 2",
              r.get("what") == "row" and r.get("at") == 2 and r.get("count") == 2, r)
    need_ok(cid, "save-inserted")
    need_ok(cid, "close-inserted")
    need_ok(cid, "open-inserted")
    r = need_ok(cid, "rows-delete")
    if r:
        check(cid, "rows delete reports count 2", r.get("count") == 2, r.get("count"))
    r = need_ok(cid, "cols-insert")
    if r:
        check(cid, "cols insert reports column/at 3", r.get("what") == "column" and r.get("at") == 3, r)
    r = need_ok(cid, "cols-delete")
    if r:
        check(cid, "cols delete reports count 1", r.get("count") == 1, r.get("count"))
    need_ok(cid, "save")

    shifted = load_book(cid, "12-rowscols-shifted.xlsx")
    if shifted is not None:
        ws = shifted.ws("Rows")
        check(cid, "inserted: A2 and A3 are empty",
              ws.cell(row=2, column=1).value in (None, "") and ws.cell(row=3, column=1).value in (None, ""),
              (show(ws.cell(row=2, column=1).value), show(ws.cell(row=3, column=1).value)))
        check_block(cid, "inserted: the four data rows moved down by 2",
                    ws, [[norm(base["Rows"].cell(row=r, column=c).value) for c in (1, 2)] for r in range(2, 6)],
                    r1=4, c1=1)
    final = load_book(cid, "12-rowscols-inserted.xlsx")
    if final is None:
        return
    ws = final.ws("Rows")
    check_block(cid, "after delete, the Rows sheet equals the fixture again",
                ws, block(base["Rows"], 1, 1, 5, 2))
    check(cid, "no leftover value in column C",
          all(ws.cell(row=r, column=3).value in (None, "") for r in range(1, 6)),
          [show(ws.cell(row=r, column=3).value) for r in range(1, 6)])


def case_13_snapshot(base):
    cid = "13-snapshot"
    print("== %s: snapshot - sheet-scoped diff and restore isolation ==" % cid)
    r = need_ok(cid, "snapshot")
    if r:
        check(cid, "snapshot captured cells", (r.get("cells") or 0) > 0, r.get("cells"))
    r = need_ok(cid, "dirty")
    if r:
        check(cid, "two cells dirtied", r.get("count") == 2, r.get("count"))
    r = need_ok(cid, "diff")
    if r:
        changes = r.get("changes") or []
        check(cid, "diff reports exactly 1 change", r.get("totalChanges") == 1, r.get("totalChanges"))
        if changes:
            one = changes[0]
            check(cid, "the change is Data!C2", (one.get("sheet"), one.get("cell")) == ("Data", "C2"),
                  (one.get("sheet"), one.get("cell")))
            check(cid, "before/after are low -> 999",
                  same(one.get("before"), "low") and same(one.get("after"), 999),
                  (one.get("before"), one.get("after")))
    r = need_ok(cid, "restore")
    if r:
        check(cid, "restore put 1 cell back", r.get("restored") == 1, r.get("restored"))
        extra = r.get("extraSheets") or []
        check(cid, "restore lists Summary as out of scope", "Summary" in extra, extra)
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    check_cell(cid, book.ws("Data"), 2, 3, norm(base["Data"].cell(row=2, column=3).value),
               "Data!C2 was restored from the snapshot")
    check_cell(cid, book.ws("Summary"), 2, 2, 888, "Summary!B2 kept the out-of-scope edit")
    check_block(cid, "Data around C2 is the fixture again",
                book.ws("Data"), block(base["Data"], 1, 1, 5, 4))


def case_14_restore(base):
    cid = "14-restore"
    print("== %s: restore -Cells - selective rollback ==" % cid)
    need_ok(cid, "snapshot")
    need_ok(cid, "dirty")
    r = need_ok(cid, "restore-one")
    if r:
        check(cid, "selective restore is flagged", r.get("selective") is True, r.get("selective"))
        check(cid, "F2 cleared (it was not in the snapshot)", r.get("clearedListed") == 1, r.get("clearedListed"))
        check(cid, "only the listed cell was touched", r.get("restored") == 1 and r.get("cleared") == 0,
              (r.get("restored"), r.get("cleared")))
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    ws = book.ws("Data")
    check(cid, "F2 is empty again", ws.cell(row=2, column=6).value in (None, ""),
          show(ws.cell(row=2, column=6).value))
    check_cell(cid, ws, 3, 6, 222, "F3 kept the edit (not listed)")
    check_cell(cid, ws, 4, 6, 333, "F4 kept the edit (not listed)")


def case_15_chart(base):
    cid = "15-chart"
    print("== %s: chart - add two, list, delete one, read the drawing back ==" % cid)
    r = need_ok(cid, "add-line")
    if r:
        check(cid, "line chart on Series with 2 series", r.get("sheet") == "Series" and r.get("seriesCount") == 2,
              (r.get("sheet"), r.get("seriesCount")))
        check(cid, "series names are q1/q2", r.get("seriesNames") == ["q1", "q2"], r.get("seriesNames"))
        check(cid, "chart name is ChartQH", r.get("name") == "ChartQH", r.get("name"))
        check(cid, "chart title is Q-H", r.get("title") == "Q-H", r.get("title"))
    r = need_ok(cid, "add-column")
    if r:
        check(cid, "column chart with 1 series", r.get("seriesCount") == 1, r.get("seriesCount"))
    r = need_ok(cid, "list-2")
    if r:
        charts = r.get("charts") or []
        check(cid, "list reports 2 charts", r.get("count") == 2, r.get("count"))
        check(cid, "list types are line then column",
              [chart_type(c.get("type")) for c in charts] == ["line", "column"], [c.get("type") for c in charts])
    r = need_ok(cid, "delete-1")
    if r:
        check(cid, "delete reports index 1", r.get("deleted") == 1, r.get("deleted"))
    r = need_ok(cid, "list-1")
    if r:
        charts = r.get("charts") or []
        check(cid, "one chart left, and it is the column one",
              r.get("count") == 1 and charts and chart_type(charts[0].get("type")) == "column", charts)
    need_ok(cid, "save")

    book = load_book(cid, cid + ".xlsx")
    if book is None:
        return
    charts = list(book.ws("Series")._charts)
    kinds = [type(c).__name__ for c in charts]
    check(cid, "openpyxl reads exactly one chart from the saved file", len(charts) == 1, kinds)
    check(cid, "the surviving chart is a column/bar chart", kinds and kinds[0] in ("BarChart", "Bar3DChart"), kinds)
    check_block(cid, "Series data is untouched",
                book.ws("Series"), block(base["Series"], 1, 1, 21, 3))


def case_16_xls(base):
    cid = "16-xls"
    print("== %s: .xls round trip ==" % cid)
    r = need_ok(cid, "save-as-xls")
    if r:
        check(cid, "SaveAs produced a .xls", str(r.get("savedAs", "")).lower().endswith(".xls"), r.get("savedAs"))
    r = need_ok(cid, "read")
    if r:
        values = r.get("values") or []
        check(cid, "read back the Data header row",
              values and [norm(v) for v in values[0]] == ["site", "value", "level", "note"], values[:1])
        check(cid, "read back Data!B2/B3 numbers",
              len(values) >= 3 and same(values[1][1], 10) and same(values[2][1], 11), values[1:3])
    r = need_ok(cid, "write")
    if r:
        applied = (r.get("applied") or [{}])[0]
        check(cid, "write with expect succeeded in .xls", r.get("count") == 1, r.get("count"))
        check(cid, "written cell shows 777", str(norm(applied.get("after"))) == "777", applied.get("after"))
    r = need_ok(cid, "fill")
    if r:
        check(cid, "fill wrote 100 rows into .xls", r.get("rows") == 100, r.get("rows"))
        check(cid, "fill topAfter 2 / bottomAfter 200",
              str(norm(r.get("topAfter"))) == "2" and str(norm(r.get("bottomAfter"))) == "200",
              (r.get("topAfter"), r.get("bottomAfter")))
    need_ok(cid, "save")

    if FILES is None:
        file_skip("%s/16-xls.xls" % cid)
        return
    path = Path(FILES) / cid / "16-xls.xls"
    if not path.exists():
        check(cid, "saved .xls exists", False, str(path))
        return
    head = path.read_bytes()[:8]
    check(cid, "saved .xls has the OLE magic (BIFF container)",
          head == bytes.fromhex("D0CF11E0A1B11AE1"), head.hex().upper())
    try:
        import xlrd
        sheet = xlrd.open_workbook(str(path))
    except Exception as exc:  # noqa: BLE001
        check(cid, "xlrd can read the saved .xls", False, "%s: %s" % (type(exc).__name__, exc))
        return
    check(cid, "xlrd can read the saved .xls", True)
    data = sheet.sheet_by_name("Data")
    check(cid, "Data!B2 persisted as 777", str(norm(data.cell_value(1, 1))) == "777", data.cell_value(1, 1))
    fill = sheet.sheet_by_name("Fill")
    check(cid, "Fill!C101 persisted as 200", same(fill.cell_value(100, 2), 200), fill.cell_value(100, 2))
    check(cid, "Fill!C2 persisted as 2", same(fill.cell_value(1, 2), 2), fill.cell_value(1, 2))


CASES = [
    ("01-write", case_01_write),
    ("02-write-anchor", case_02_write_anchor),
    ("03-fill", case_03_fill),
    ("04-replace", case_04_replace),
    ("05-sort", case_05_sort),
    ("06-deletewhere", case_06_deletewhere),
    ("07-fillblanks", case_07_fillblanks),
    ("08-dedupe", case_08_dedupe),
    ("09-clean", case_09_clean),
    ("10-table", case_10_table),
    ("11-sheet", case_11_sheet),
    ("12-rowscols", case_12_rowscols),
    ("13-snapshot", case_13_snapshot),
    ("14-restore", case_14_restore),
    ("15-chart", case_15_chart),
    ("16-xls", case_16_xls),
]


def case_was_run(cid):
    logs = Path(WORK) / "logs"
    if not logs.is_dir():
        return False
    return any(logs.glob("%s-*.json" % cid))


def main():
    global WORK, FILES, SKIP_REASON
    parser = argparse.ArgumentParser(description="Independent oracle for the Excel data regression run")
    parser.add_argument("--work", required=True, help="work dir written by run_data.ps1")
    parser.add_argument("--files", default="", help="dir holding the decrypted workbooks (default: <work>/plain)")
    parser.add_argument("--skip-file", action="store_true", help="skip every file-level assertion")
    parser.add_argument("--skip-reason", default="", help="why the file oracle is off (printed with each SKIP)")
    args = parser.parse_args()

    WORK = Path(args.work)
    SKIP_REASON = args.skip_reason
    if args.skip_file:
        FILES = None
    else:
        FILES = Path(args.files) if args.files else WORK / "plain"

    base_path = WORK / "base.xlsx"
    if not base_path.exists():
        print("FAIL  fixture missing: %s" % base_path)
        return 1
    base = openpyxl.load_workbook(str(base_path))

    print("oracle: pristine fixture %s" % base_path)
    print("oracle: saved workbooks %s" % (FILES if FILES else "(file oracle disabled)"))
    print("")

    ran = 0
    for cid, func in CASES:
        if not case_was_run(cid):
            print("== %s: not run ==" % cid)
            skip(cid, "whole case", "no skill responses recorded (case not selected or the driver stopped early)")
            continue
        ran += 1
        before = (COUNTS["pass"], COUNTS["fail"], COUNTS["skip"])
        try:
            func(base)
        except Exception as exc:  # noqa: BLE001 - a broken check is a finding, not a crash
            check(cid, "oracle ran without an internal error", False, "%s: %s" % (type(exc).__name__, exc))
        after = (COUNTS["pass"], COUNTS["fail"], COUNTS["skip"])
        print("  ---- %s: %d pass, %d fail, %d skip" % (
            cid, after[0] - before[0], after[1] - before[1], after[2] - before[2]))
        print("")

    summary = {
        "pass": COUNTS["pass"], "fail": COUNTS["fail"], "skip": COUNTS["skip"],
        "cases": ran, "files": bool(FILES), "skipReason": SKIP_REASON,
        "failures": ["%s: %s" % (cid, label) for cid, label in FAILED],
    }
    try:
        (WORK / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    except Exception:
        pass

    print("------------------------------------------------------------------")
    line = "RESULT  %d PASS  %d FAIL  %d SKIP  (%d case(s) run%s)" % (
        COUNTS["pass"], COUNTS["fail"], COUNTS["skip"], ran,
        "" if FILES else ", file oracle off")
    print(line)
    if FAILED:
        print("")
        print("failed checks:")
        for cid, label in FAILED:
            print("  - %s: %s" % (cid, label))
    return 1 if COUNTS["fail"] else 0


if __name__ == "__main__":
    sys.exit(main())
