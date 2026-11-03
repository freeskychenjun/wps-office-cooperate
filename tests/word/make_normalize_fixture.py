#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate the self-contained fixture used by tests/word/check_normalize.py.

The original fixture of that check was a copy of a real 2026-11-03 report that lived
outside the repository (D:\\<workspace>\\...\\orig-copy.docx), so a fresh clone failed on
its very first Word test.  This script replaces it with a document that carries exactly
the *static* properties check_normalize asserts (and nothing else - the end-to-end COM
assertions A1-A9 still need a real report, see check_normalize.py --fixture):

    14 tables
     4 tables with a caption typed into the first cell of row 1 ("表 x-y ...")
     2 tables WITHOUT <w:tblBorders> (the width-blow-up trigger); the other 12 carry
       an explicit tblBorders block, which python-docx never writes by itself
     0 tracked revisions (<w:ins> / <w:del>) and 0 comments (no word/comments.xml)

  python make_normalize_fixture.py              # tests/word/out/normalize-fixture.docx
  python make_normalize_fixture.py --out <path>
  python make_normalize_fixture.py --quiet

The output directory is gitignored (tests/word/out/); tests/clean.ps1 removes it.
"""

import argparse
import os
import sys
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(HERE, "out", "normalize-fixture.docx")

TABLE_COUNT = 14
# 1-based table numbers.  A caption in row 1 is what `word table -Action plan` flags
# captionInFirstRow; the real report has 4 of them.
CAPTION_TABLES = (1, 4, 7, 10)
# ... and exactly these two are written without tblBorders (the report's T3 / T9).
NO_BORDER_TABLES = (3, 9)

BORDER_EDGES = ("top", "left", "bottom", "right", "insideH", "insideV")


def add_borders(table):
    """Append an explicit <w:tblBorders> block (Word/WPS treat a missing block as
    'no borders' and rebuild the layout on edit, which is the bug under test)."""
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in BORDER_EDGES:
        element = OxmlElement("w:" + edge)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "8")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "auto")
        borders.append(element)
    # Schema order: tblBorders must come before tblLayout / tblCellMar / tblLook,
    # and python-docx does write tblLook for every table it creates.
    look = tbl_pr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(borders)
    else:
        tbl_pr.append(borders)


def build(path):
    """Write the fixture to `path` (parent directory is created)."""
    document = Document()
    for style, size in (("Title", 20), ("Heading 1", 11), ("Normal", 10.5)):
        try:
            document.styles[style].font.size = Pt(size)
        except KeyError:
            pass

    document.add_heading("Table normalize acceptance fixture", level=0)
    document.add_paragraph(
        "Synthetic corpus for tests/word/check_normalize.py --static-only. "
        "%d tables, %d captions typed into row 1, %d tables without tblBorders."
        % (TABLE_COUNT, len(CAPTION_TABLES), len(NO_BORDER_TABLES)))

    for number in range(1, TABLE_COUNT + 1):
        document.add_heading("Section %d" % number, level=1)
        paragraph = document.add_paragraph()
        paragraph.add_run(
            "Commentary for table %d: the body text refers to it as table %d-%d."
            % (number, number, number))
        if number in CAPTION_TABLES:
            # The caption is a real row of the table (the defect the plan reports).
            document.add_paragraph("Table %d follows." % number)

        table = document.add_table(rows=4, cols=3)
        if number not in NO_BORDER_TABLES:
            add_borders(table)

        # Row 1: either the typed caption (first cell) or a header row.
        first = table.rows[0].cells
        if number in CAPTION_TABLES:
            first[0].text = "\u8868 %d-%d \u4e3b\u8981\u5de5\u7a0b\u91cf\u8868" % (number, number)
            first[1].text = ""
            first[2].text = ""
        else:
            first[0].text = "Item"
            first[1].text = "Unit"
            first[2].text = "Value"
        # Rows 2..4: plain data, so the table is a normal 4-row table.
        for row_index, row in enumerate(table.rows[1:], start=1):
            cells = row.cells
            cells[0].text = "row %d" % row_index
            cells[1].text = "m3/s" if row_index % 2 else "km2"
            cells[2].text = "%d.%d" % (number, row_index)

        document.add_paragraph(
            "Text after table %d keeps the tables apart and makes the reference "
            "count non-zero." % number)

    out_dir = os.path.dirname(os.path.abspath(path))
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    document.save(path)
    return path


def facts(path):
    """Count the static facts from the saved zip (independent of check_normalize)."""
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
        names = archive.namelist()
    tables = xml.count("<w:tbl>") + xml.count("<w:tbl ")
    borders = xml.count("<w:tblBorders>")
    revisions = xml.count("<w:ins ") + xml.count("<w:del ")
    comments = 0
    if "word/comments.xml" in names:
        comments = archive.read("word/comments.xml").decode("utf-8").count("<w:comment ")
    return {"tables": tables, "noBorders": tables - borders,
            "revisions": revisions, "comments": comments}


def main(argv):
    parser = argparse.ArgumentParser(description="Generate the table-normalize fixture.")
    parser.add_argument("--out", default=DEFAULT_OUT)
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)

    path = build(os.path.abspath(args.out))
    if args.quiet:
        return 0

    counted = facts(path)
    print("fixture: %s" % path)
    print("static: %d tables / %d tables without tblBorders / %d revisions / %d comments"
          % (counted["tables"], counted["noBorders"], counted["revisions"], counted["comments"]))
    expected = (TABLE_COUNT, len(NO_BORDER_TABLES), 0, 0)
    actual = (counted["tables"], counted["noBorders"], counted["revisions"], counted["comments"])
    if actual != expected:
        print("FIXTURE MISMATCH: expected (tables, noBorders, revisions, comments) = %r"
              % (expected,))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
