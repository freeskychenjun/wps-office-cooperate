#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the fixture deck for `ppt replace` acceptance (tests/ppt).

The deck is deliberately made with python-pptx, i.e. by a completely different
writer than the C# helper, so the expectations are not "the same code checking
itself".  Every text slot holds one or two occurrences of the token, and the
layout of the slots is fixed, so check.py can derive the expected hit count,
the expected per-scope counts and the expected per-character formatting from the
pristine copy of this deck.

  python make_deck.py --out <dir>

writes
  <dir>/replace-deck.pptx            the deck the skill will edit
  <dir>/replace-deck.pristine.pptx   untouched copy, ground truth for check.py

Slots with a hit (16 in total):
  slide 1  cover title / cover subtitle                                   2  body
  slide 2  three-paragraph body / project-number line                     4  body
  slide 3  two textboxes inside one group                                 2  group
  slide 4  two table cells                                                2  table
  slide 5  cross-run token x2 + one mixed-size run                        3  body
  slide 6  notes page                                                     1  notes
  slide 7  one control shape (also the -Case probe)                       1  body
  master   one textbox on the slide master                                1  master
  footer   NOT part of this deck (see --mode footer): it needs COM, and
           python-pptx cannot write <p:hf>, so the driver adds it there and
           asserts it in a separate scenario.
"""
import argparse
import os
import shutil
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml import parse_xml
from pptx.util import Inches, Pt

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

TOKEN = "2024"
NEW = "2025"


def style_run(run, size=None, bold=None, color=None, latin=None, ea=None):
    font = run.font
    if size is not None:
        font.size = Pt(size)
    if bold is not None:
        font.bold = bold
    if latin is not None:
        font.name = latin
    if color is not None:
        font.color.rgb = RGBColor.from_string(color)
    if ea is not None:
        rpr = font._rPr
        el = rpr.find(qn("a:ea"))
        if el is None:
            el = parse_xml("<a:ea %s/>" % nsdecls("a"))
            latin_el = rpr.find(qn("a:latin"))
            if latin_el is not None:
                latin_el.addnext(el)
            else:
                rpr.append(el)
        el.set("typeface", ea)


def add_text(shapes, left, top, width, height, paragraphs,
             size=20, latin="Arial", ea="宋体", bold=None, color=None):
    """paragraphs: list of paragraphs; each paragraph is a list of
    (text, overrides-dict) runs."""
    box = shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    for p_index, runs in enumerate(paragraphs):
        para = tf.paragraphs[0] if p_index == 0 else tf.add_paragraph()
        for text, over in runs:
            run = para.add_run()
            run.text = text
            style_run(run,
                      size=over.get("size", size),
                      bold=over.get("bold", bold),
                      color=over.get("color", color),
                      latin=over.get("latin", latin),
                      ea=over.get("ea", ea))
    return box


def build_footer_deck(path):
    """One blank slide, no token anywhere: the driver sets the footer over COM."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(path)


def build_full_deck(path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ---- slide 1: cover (2 hits)
    s1 = prs.slides.add_slide(blank)
    add_text(s1.shapes, 0.8, 1.2, 11.5, 1.6,
             [[(TOKEN + "年度洪水风险图编制项目", {})]],
             size=40, latin="Arial", ea="黑体", bold=True, color="1F4E79")
    add_text(s1.shapes, 0.8, 3.1, 11.5, 0.9,
             [[("江西省水利规划设计研究院  " + TOKEN + "年10月", {})]],
             size=20, latin="Arial", ea="宋体", color="333333")

    # ---- slide 2: three paragraphs + project number (4 hits)
    s2 = prs.slides.add_slide(blank)
    add_text(s2.shapes, 0.8, 1.0, 11.5, 2.8, [
        [(TOKEN + "年完成资料收集与整理。", {})],
        [(TOKEN + "年开展洪水风险图编制。", {})],
        [(TOKEN + "年10月通过专家评审。", {})],
    ], size=20, latin="Arial", ea="宋体", color="333333")
    add_text(s2.shapes, 0.8, 4.2, 11.5, 0.9,
             [[("项目编号：JX-" + TOKEN + "-001", {})]],
             size=18, latin="Arial", ea="微软雅黑", color="666666")

    # ---- slide 3: one group, two of its three textboxes carry the token (2 hits)
    s3 = prs.slides.add_slide(blank)
    group = s3.shapes.add_group_shape()
    add_text(group.shapes, 0.8, 1.0, 11.5, 0.8,
             [[("第一分册：" + TOKEN + "年成果", {})]],
             size=16, latin="Arial", ea="楷体", color="1F4E79")
    add_text(group.shapes, 0.8, 2.2, 11.5, 0.8,
             [[("第二分册：" + TOKEN + "年成果", {})]],
             size=16, latin="Arial", ea="楷体", color="1F4E79")
    add_text(group.shapes, 0.8, 3.4, 11.5, 0.8,
             [[("第三分册：无相关年号", {})]],
             size=16, latin="Arial", ea="楷体", color="1F4E79")

    # ---- slide 4: table, two cells carry the token (2 hits)
    s4 = prs.slides.add_slide(blank)
    table = s4.shapes.add_table(2, 2, Inches(0.8), Inches(1.2),
                                Inches(7.5), Inches(1.6)).table
    table.cell(0, 0).text = "年度"
    table.cell(0, 1).text = "工作内容"
    table.cell(1, 0).text = TOKEN + "年"
    table.cell(1, 1).text = TOKEN + "年编制"
    for r in range(2):
        for c in range(2):
            for para in table.cell(r, c).text_frame.paragraphs:
                for run in para.runs:
                    style_run(run, size=14, latin="Arial", ea="微软雅黑", color="333333")

    # ---- slide 5: cross-run tokens (2) + one mixed-size token (1)
    s5 = prs.slides.add_slide(blank)
    add_text(s5.shapes, 0.8, 1.0, 11.5, 1.0, [[
        ("评价", {}),
        ("20", {"size": 24, "bold": True, "color": "C00000"}),
        ("24", {"size": 24, "bold": False, "color": "C00000"}),
        ("年度成果", {}),
    ]], size=16, latin="Arial", ea="宋体", color="333333")
    add_text(s5.shapes, 0.8, 2.4, 11.5, 1.0, [[
        ("复核", {}),
        ("20", {"size": 22, "color": "0070C0"}),
        ("24", {"size": 22, "color": "0070C0"}),
        ("年度成果", {}),
    ]], size=16, latin="Arial", ea="宋体", color="333333")
    add_text(s5.shapes, 0.8, 3.8, 11.5, 1.0, [[
        (TOKEN, {"size": 26, "color": "375623", "ea": "楷体"}),
        ("年度结论", {}),
    ]], size=16, latin="Arial", ea="宋体", color="333333")

    # ---- slide 6: body without the token + notes page with it (1 hit)
    s6 = prs.slides.add_slide(blank)
    add_text(s6.shapes, 0.8, 1.0, 11.5, 1.0,
             [[("本页正文不含年号。", {})]],
             size=20, latin="Arial", ea="宋体", color="333333")
    s6.notes_slide.notes_text_frame.text = "备注：" + TOKEN + "年数据来源于水文年鉴。"

    # ---- slide 7: control shape, also the -Case probe (1 hit for the token)
    s7 = prs.slides.add_slide(blank)
    add_text(s7.shapes, 0.8, 1.0, 11.5, 1.0,
             [[("2023年资料与 AA" + TOKEN + "aa 编号", {})]],
             size=18, latin="Arial", ea="宋体", color="333333")

    # ---- slide master: the token in the master title placeholder (1 hit, opt-in
    # scope).  The Blank layout has no title placeholder, so the text shows up only
    # in the master itself - which is exactly what the master scope has to find.
    master_title = None
    for ph in prs.slide_master.placeholders:
        if ph.placeholder_format.type == 1:     # TITLE
            master_title = ph
            break
    if master_title is None:
        raise SystemExit("the default template has no master title placeholder")
    master_title.text = "母版·" + TOKEN + "·标题位"
    for para in master_title.text_frame.paragraphs:
        for run in para.runs:
            style_run(run, size=12, latin="Arial", ea="宋体", color="808080")

    prs.save(path)


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--mode", default="full", choices=["full", "footer"])
    args = ap.parse_args(argv)

    out = os.path.abspath(args.out)
    if not os.path.isdir(out):
        os.makedirs(out)

    if args.mode == "footer":
        deck = os.path.join(out, "footer-deck.pptx")
        build_footer_deck(deck)
        shutil.copyfile(deck, os.path.join(out, "footer-deck.pristine.pptx"))
        print(deck)
        return 0

    deck = os.path.join(out, "replace-deck.pptx")
    build_full_deck(deck)
    shutil.copyfile(deck, os.path.join(out, "replace-deck.pristine.pptx"))
    print(deck)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
