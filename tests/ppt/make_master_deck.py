#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fixture for `ppt master` (tests/ppt).

source.pptx  4 pages, 16:9, master title stamped "SRC-MARK", pages titled S-1..S-4
template.pptx 2 sample pages, same size, master title stamped "TPL-MARK", and a
             gold band on every page so the difference is visible on screen

  python make_master_deck.py --out <dir>
"""
import argparse
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def stamp_master(prs, text):
    for ph in prs.slide_master.placeholders:
        if ph.placeholder_format.type == 1:
            ph.text = text
            for para in ph.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(12)
            return True
    return False


def build(path, titles, mark, band=None):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    stamp_master(prs, mark)
    blank = prs.slide_layouts[6]
    for text in titles:
        slide = prs.slides.add_slide(blank)
        if band:
            strip = slide.shapes.add_shape(1, Inches(0), Inches(6.2), Inches(13.333), Inches(0.9))
            strip.fill.solid()
            strip.fill.fore_color.rgb = RGBColor.from_string(band)
            strip.line.fill.background()
        box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11), Inches(1.2))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(32)
        run.font.bold = True
        box.name = "title"
        body = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11), Inches(1.0))
        brun = body.text_frame.paragraphs[0].add_run()
        brun.text = "body of " + text
        brun.font.size = Pt(20)
    prs.save(path)


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)
    out = os.path.abspath(args.out)
    if not os.path.isdir(out):
        os.makedirs(out)
    build(os.path.join(out, "source.pptx"), ["S-1", "S-2", "S-3", "S-4"], "SRC-MARK")
    build(os.path.join(out, "template.pptx"), ["TPL-1", "TPL-2"], "TPL-MARK", band="1F4E79")
    print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
