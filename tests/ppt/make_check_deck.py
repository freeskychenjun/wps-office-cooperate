#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fixture for `ppt check` / `ppt outline` acceptance (tests/ppt).

Every slide seeds exactly one (or a named few) of the defects the rule library is
supposed to catch, so check.py can assert rule by rule instead of "it found
something".  Two slides are deliberately CLEAN: a gate that fires on everything is
worse than no gate.

  python make_check_deck.py --out <dir>

writes <dir>/check-deck.pptx (plus .pristine.pptx)

Seeded defects, one slide each:
  2  overflow      a text box far too short for its text
  3  fontsize      body text at 9pt
  4  placeholder   leftover "Click to add title" style prompt
  5  outframe      a shape half off the page
  6  pagenum       no slide number (the deck turns numbers on everywhere else)
  7  density       >300 characters of body text
  8  contrast      light grey text on a light grey fill
  9  empty         a slide with nothing on it
  10 image-small   a tiny picture
  11 fontfallback  a font name that is certainly not installed
  12 title         title at 18pt while every other title is 32pt
  1  clean cover   proper title, no number (the pagenum rule allows the cover)
  13 clean page    proper title + body at 20pt
"""
import argparse
import os
import shutil
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt, Emu

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

MISSING_FONT = "Definitely-Not-Installed-Font"


def title(slide, text, size=32):
    box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.5), Inches(1.0))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.name = "Arial"
    run.font.bold = True
    # mark it as a title so the role detection can find it
    box.name = "title"
    return box


def body(slide, text, size=20, top=1.8, left=0.8, width=11.5, height=1.2,
         color=None, fill=None, name="body"):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.word_wrap = True
    run = box.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.name = "Arial"
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
    box.name = name
    if fill is not None:
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor.from_string(fill)
    return box


def number_on(slide, value=True):
    """Turn the slide number placeholder on/off for this slide."""
    hf = slide._element.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}clrMapOvr")
    # python-pptx has no <p:hf> API: write the element directly.
    ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    hf_el = slide._element.find(ns + "hf")
    if hf_el is None:
        hf_el = slide._element.makeelement(ns + "hf", {})
        slide._element.append(hf_el)
    hf_el.set("sldNum", "1" if value else "0")


def build(path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # 1 cover: clean, no slide number
    cover = prs.slides.add_slide(blank)
    title(cover, "洪水风险图编制成果汇报")
    body(cover, "江西省水利规划设计研究院  2025年10月", size=20, top=3.0)
    number_on(cover, False)

    # 2 overflow
    s = prs.slides.add_slide(blank)
    title(s, "资料收集与整理")
    box = body(s, "2025年完成修水水系干支流 217 处断面的资料收集、整理与复核工作，"
                  "并对历史洪水痕迹进行了补充调查。", size=20, height=0.4)
    number_on(s)

    # 3 fontsize
    s = prs.slides.add_slide(blank)
    title(s, "编制依据")
    body(s, "依据《洪水风险图编制技术细则》及江西省实施细则。", size=9, top=1.8)
    number_on(s)

    # 4 placeholder left over
    s = prs.slides.add_slide(blank)
    title(s, "技术路线")
    body(s, "Click to add title", size=20, top=1.8)
    number_on(s)

    # 5 out of frame
    s = prs.slides.add_slide(blank)
    title(s, "计算模型")
    shape = s.shapes.add_textbox(Inches(12.4), Inches(2.0), Inches(3.0), Inches(1.0))
    run = shape.text_frame.paragraphs[0].add_run()
    run.text = "600 万元"
    run.font.size = Pt(20)
    shape.name = "outbox"
    number_on(s)

    # 6 page number missing
    s = prs.slides.add_slide(blank)
    title(s, "成果汇总")
    body(s, "共编制风险图 43 幅，覆盖面积 1.2 万平方公里。", size=20, top=1.8)
    number_on(s, False)

    # 7 density
    s = prs.slides.add_slide(blank)
    title(s, "风险区特征")
    body(s, "风险区涉及 12 个县（市、区），" * 30, size=16, top=1.8, height=4.5)
    number_on(s)

    # 8 contrast: light grey on light grey
    s = prs.slides.add_slide(blank)
    title(s, "转移安置方案")
    body(s, "转移人口 3.2 万人，安置点 46 处。", size=20, top=1.8,
         color="DDDDDD", fill="EEEEEE")
    number_on(s)

    # 9 empty
    s = prs.slides.add_slide(blank)
    number_on(s)

    # 10 tiny picture
    s = prs.slides.add_slide(blank)
    title(s, "典型断面示意")
    pic_path = os.path.join(os.path.dirname(os.path.abspath(path)), "_tiny.png")
    make_tiny_png(pic_path)
    s.shapes.add_picture(pic_path, Inches(1.0), Inches(2.0), Inches(0.9), Inches(0.9))
    number_on(s)

    # 11 font that is not installed
    s = prs.slides.add_slide(blank)
    title(s, "附件清单")
    box = body(s, "附表 1 风险图清单；附表 2 转移安置表。", size=20, top=1.8)
    for para in box.text_frame.paragraphs:
        for run in para.runs:
            run.font.name = MISSING_FONT
    number_on(s)

    # 12 title size out of line
    s = prs.slides.add_slide(blank)
    title(s, "下一步工作", size=18)
    body(s, "2026 年汛前完成成果复核与归档。", size=20, top=1.8)
    number_on(s)

    # 13 clean content page
    s = prs.slides.add_slide(blank)
    title(s, "结语")
    body(s, "恳请各位专家批评指正。", size=20, top=1.8)
    number_on(s)

    prs.save(path)


def make_tiny_png(path):
    """A 40x40 PNG: at 0.9in wide that is far below the 20%-of-slide-width floor."""
    import struct
    import zlib

    w = h = 40
    raw = b""
    for y in range(h):
        raw += b"\x00" + bytes([200, 200, 200] * w)

    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(raw, 9))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as fh:
        fh.write(png)


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--mode", default="check", choices=["check"])
    args = ap.parse_args(argv)
    out = os.path.abspath(args.out)
    if not os.path.isdir(out):
        os.makedirs(out)
    deck = os.path.join(out, "check-deck.pptx")
    build(deck)
    shutil.copyfile(deck, os.path.join(out, "check-deck.pristine.pptx"))
    print(deck)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
