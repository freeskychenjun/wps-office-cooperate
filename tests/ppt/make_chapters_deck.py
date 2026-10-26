#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fixture for `ppt split` / `ppt merge` (tests/ppt).

split-deck.pptx : 12 slides, 16:9, each with a unique title "PAGE-N" and a 24pt body
                  line, plus a picture on every third page (so a lost/mis-scaled page
                  is visible in the count AND in the geometry).
county-a/b/c     : 3 slides each, titled "A-1".."C-3", so a merge can be checked for
                  ORDER, not just for count.

  python make_chapters_deck.py --out <dir>
"""
import argparse
import os
import struct
import sys
import zlib

from pptx import Presentation
from pptx.util import Inches, Pt

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def png(path, w, h, rgb):
    raw = b""
    for _ in range(h):
        raw += b"\x00" + bytes(rgb) * w

    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    data = b"\x89PNG\r\n\x1a\n"
    data += chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
    data += chunk(b"IDAT", zlib.compress(raw, 9))
    data += chunk(b"IEND", b"")
    with open(path, "wb") as fh:
        fh.write(data)


def deck(path, titles, with_pictures, master_mark=""):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    if master_mark:
        # A marker in the MASTER is how a part file proves it kept the design: pasting
        # slides into a fresh deck loses it, copying the file keeps it.
        for ph in prs.slide_master.placeholders:
            if ph.placeholder_format.type == 1:
                ph.text = master_mark
                for para in ph.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(12)
                break
    asset = os.path.join(os.path.dirname(os.path.abspath(path)), "asset.png")
    if not os.path.exists(asset):
        png(asset, 60, 40, (31, 78, 121))
    for index, text in enumerate(titles, start=1):
        slide = prs.slides.add_slide(blank)
        box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.0), Inches(1.0))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(32)
        run.font.bold = True
        box.name = "title"
        body = slide.shapes.add_textbox(Inches(0.8), Inches(1.9), Inches(11.0), Inches(1.0))
        brun = body.text_frame.paragraphs[0].add_run()
        brun.text = "body of " + text
        brun.font.size = Pt(20)
        if with_pictures and index % 3 == 0:
            slide.shapes.add_picture(asset, Inches(1.0), Inches(3.2), Inches(2.0), Inches(1.4))
    prs.save(path)


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)
    out = os.path.abspath(args.out)
    if not os.path.isdir(out):
        os.makedirs(out)
    deck(os.path.join(out, "split-deck.pptx"),
         ["PAGE-%d" % n for n in range(1, 13)], True, master_mark="MASTER-MARK")
    deck(os.path.join(out, "county-a.pptx"), ["A-%d" % n for n in (1, 2, 3)], False, master_mark="MASTER-A")
    deck(os.path.join(out, "county-b.pptx"), ["B-%d" % n for n in (1, 2, 3)], False, master_mark="MASTER-B")
    deck(os.path.join(out, "county-c.pptx"), ["C-%d" % n for n in (1, 2, 3)], False, master_mark="MASTER-C")
    print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
