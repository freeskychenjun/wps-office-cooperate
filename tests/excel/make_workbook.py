#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fixture builder for the Excel data-layer regression driver (tests/excel).

    python make_workbook.py --out <work>/base.xlsx

Builds ONE pristine workbook with a sheet per scenario, deterministic content (no clock,
no randomness, no external files).  The driver copies it per case, runs the skill against
WPS, and check_data.py re-derives every expectation from this same pristine file.

Sheet tour (see check_data.py CASES for how each one is exercised):

  Data     4 columns x 30 rows        write / write-anchor / snapshot (C2) / selective restore (F)
  Fill     2000 rows                  fill -Formula down a column (acceptance case 17: B1:B2000)
  Calc     small numeric + formulas   formula persistence after save (part of 01-write)
  Sort     name/qty/total(=B*10)      sort -By qty:desc, formulas must travel with their row
  Where    site/status/qty            deletewhere -Column status -Value drop
  Blanks   item/qty with holes        fillblanks -With above / -With 0
  Dedupe   site/val with repeats      dedupe -Column site (first of each group kept)
  Clean    padded/case-mixed strings  clean -Casing upper -NoSpaces / clean -To number
  Table    a/b/c                      table -Name TblOne (idempotent on the same range)
  Rows     k/v                        rows + cols insert/delete
  Series   x/q1/q2, 20 rows           chart add (2 series) / list / delete
  Summary  key/value                  must survive a Data-only snapshot restore
  Text     name/code/label            replace: dry-run, single pass, numeric/formula guards, idempotency

Pure ASCII on purpose: the generation is plain openpyxl calls and every label is ASCII, so
there is nothing to get wrong across code pages.
"""
import argparse
import sys

import openpyxl
from openpyxl.utils import get_column_letter

LEVELS = ("low", "medium", "high")


def build_data(wb):
    ws = wb.create_sheet("Data")
    ws.append(["site", "value", "level", "note"])
    for i in range(30):
        row = 2 + i
        ws.cell(row=row, column=1, value="S%02d" % (i + 1))
        ws.cell(row=row, column=2, value=10 + i)
        ws.cell(row=row, column=3, value=LEVELS[i % 3])
        if i % 5 == 0:
            ws.cell(row=row, column=4, value="n%d" % (i + 1))


def build_fill(wb):
    ws = wb.create_sheet("Fill")
    ws.append(["x", "y"])
    for i in range(1999):
        ws.cell(row=2 + i, column=1, value=1 + i)


def build_calc(wb):
    ws = wb.create_sheet("Calc")
    ws.append(["a", "b", "sum", "ratio"])
    for i in range(5):
        row = 2 + i
        ws.cell(row=row, column=1, value=1 + i)
        ws.cell(row=row, column=2, value=(1 + i) * 2)
        ws.cell(row=row, column=3, value="=A%d+B%d" % (row, row))
        ws.cell(row=row, column=4, value="=A%d/B%d" % (row, row))


SORT_ROWS = (
    ("alpha", 5), ("bravo", 3), ("charlie", 8), ("delta", 1),
    ("echo", 9), ("foxtrot", 2), ("golf", 7), ("hotel", 4),
)


def build_sort(wb):
    ws = wb.create_sheet("Sort")
    ws.append(["name", "qty", "total"])
    for i, (name, qty) in enumerate(SORT_ROWS):
        row = 2 + i
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=qty)
        ws.cell(row=row, column=3, value="=B%d*10" % row)


WHERE_ROWS = (
    ("R1", "keep", 1), ("R2", "drop", 2), ("R3", "keep", 3), ("R4", "drop", 4),
    ("R5", "hold", 5), ("R6", "keep", 6), ("R7", "drop", 7), ("R8", "keep", 8),
)


def build_where(wb):
    ws = wb.create_sheet("Where")
    ws.append(["site", "status", "qty"])
    for name, status, qty in WHERE_ROWS:
        ws.append([name, status, qty])


def build_blanks(wb):
    ws = wb.create_sheet("Blanks")
    ws.append(["item", "qty"])
    ws.append(["a", None])
    ws.append(["b", 10])
    ws.append(["c", None])
    ws.append(["d", 30])
    ws.append(["e", None])


DEDUPE_ROWS = (
    ("X", 1), ("Y", 2), ("X", 1), ("Z", 3), ("Y", 2), ("X", 1), ("W", 4),
)


def build_dedupe(wb):
    ws = wb.create_sheet("Dedupe")
    ws.append(["site", "val"])
    for name, val in DEDUPE_ROWS:
        ws.append([name, val])


def build_clean(wb):
    ws = wb.create_sheet("Clean")
    ws.append(["name", "city", "qty"])
    ws.append(["  north  ", "  east  ", "10"])
    ws.append(["NORTH", "EAST ", "20"])
    ws.append(["north", "east", "30"])
    ws.append([" south ", " west", "40"])


def build_table(wb):
    ws = wb.create_sheet("Table")
    ws.append(["a", "b", "c"])
    for i in range(6):
        r = 2 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=(i + 1) * 2)
        ws.cell(row=r, column=3, value=(i + 1) * 3)


def build_rows(wb):
    ws = wb.create_sheet("Rows")
    ws.append(["k", "v"])
    for i in range(4):
        ws.append(["r%d" % (i + 1), i + 1])


def build_series(wb):
    ws = wb.create_sheet("Series")
    ws.append(["x", "q1", "q2"])
    for i in range(20):
        x = i + 1
        ws.cell(row=2 + i, column=1, value=x)
        ws.cell(row=2 + i, column=2, value=x * 2)
        ws.cell(row=2 + i, column=3, value=100 - x)


def build_summary(wb):
    ws = wb.create_sheet("Summary")
    ws.append(["key", "value"])
    ws.append(["alpha", 1])
    ws.append(["beta", 2])
    ws.append(["gamma", 3])


TEXT_ROWS = (
    ("P item", 10, "0"),
    ("PP", 20, "=B2+B3"),
    ("alpha P beta", 30, None),
    ("no-token-here", 40, "alpha"),
    ("P", 50, "P"),
)


def build_text(wb):
    ws = wb.create_sheet("Text")
    ws.append(["name", "code", "label"])
    for i, (name, code, label) in enumerate(TEXT_ROWS):
        row = 2 + i
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=code)
        if label is not None:
            ws.cell(row=row, column=3, value=label)


# The first sheet of a fresh openpyxl Workbook is replaced by "Data", so sheet order is exactly
# the order below; check_data.py asserts the order after the sheet add/rename/delete case.
BUILDERS = (
    ("Data", build_data),
    ("Fill", build_fill),
    ("Calc", build_calc),
    ("Sort", build_sort),
    ("Where", build_where),
    ("Blanks", build_blanks),
    ("Dedupe", build_dedupe),
    ("Clean", build_clean),
    ("Table", build_table),
    ("Rows", build_rows),
    ("Series", build_series),
    ("Summary", build_summary),
    ("Text", build_text),
)

SHEET_ORDER = [name for name, _ in BUILDERS]


def build(path):
    wb = openpyxl.Workbook()
    # Drop the default sheet; every scenario sheet is created explicitly in order.
    wb.remove(wb.active)
    for _, builder in BUILDERS:
        builder(wb)
    # Keep column widths sane so a human can read the workbook during a manual run; that is
    # view state only and never asserted.
    for ws in wb.worksheets:
        for col in range(1, min(ws.max_column, 8) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 14
    wb.save(path)


def main():
    parser = argparse.ArgumentParser(description="Build the Excel data-regression fixture workbook")
    parser.add_argument("--out", required=True, help="path of the .xlsx to write")
    parser.add_argument("--list", action="store_true", help="print the sheet tour and exit")
    args = parser.parse_args()
    if args.list:
        print("\n".join("  %-8s <-- built by make_workbook.py" % n for n in SHEET_ORDER))
        return 0
    build(args.out)
    print("wrote %s (%d sheets: %s)" % (args.out, len(SHEET_ORDER), ", ".join(SHEET_ORDER)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
