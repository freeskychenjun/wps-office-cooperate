#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fixture for `ppt comment` / `export` / `picture` (tests/ppt).

Five slides:
  1  cover                        one 200x120 picture at a known spot
  2  body + 2 pictures            one plain, one cropped
  3  one group holding 1 picture  (listed by `picture`, must NOT be replaced: a group
                                   child's coordinates are group-local and COM does not
                                   expose the transform)
  4  body, no picture             (anchor for a batch of comments)
  5  clean page                   (outline/export filler)

  python make_deliver_deck.py --out <dir>
  writes <dir>/deliver-deck.pptx (+ .pristine.pptx) and <dir>/swatch.png to insert later
"""
import argparse
import json
import os
import shutil
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


def title(slide, text, size=32):
    box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.5), Inches(1.0))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.name = "Arial"
    run.font.bold = True
    box.name = "title"
    return box


def body(slide, text, size=20, top=1.8, name="body"):
    box = slide.shapes.add_textbox(Inches(0.8), Inches(top), Inches(11.5), Inches(1.2))
    box.text_frame.word_wrap = True
    run = box.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.name = "Arial"
    box.name = name
    return box


def build(path, asset_dir):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    logo = os.path.join(asset_dir, "logo.png")
    png(logo, 200, 120, (20, 78, 121))
    swatch = os.path.join(asset_dir, "swatch.png")
    png(swatch, 240, 160, (201, 151, 0))

    # 1 cover with a picture at a known geometry
    s1 = prs.slides.add_slide(blank)
    title(s1, "洪水风险图编制成果汇报")
    body(s1, "江西省水利规划设计研究院  2025年10月", top=3.0)
    s1.shapes.add_picture(logo, Inches(1.0), Inches(4.4), Inches(2.0), Inches(1.2))

    # 2 two pictures, one cropped
    s2 = prs.slides.add_slide(blank)
    title(s2, "成果展示")
    pic_a = s2.shapes.add_picture(swatch, Inches(1.0), Inches(2.0), Inches(3.0), Inches(2.0))
    pic_a.name = "pic-plain"
    pic_b = s2.shapes.add_picture(swatch, Inches(5.0), Inches(2.0), Inches(3.0), Inches(2.0))
    pic_b.name = "pic-cropped"
    pic_b.crop_left = 0.1
    pic_b.crop_top = 0.05

    # 3 the same logo inside a group (must be listed, not replaced)
    s3 = prs.slides.add_slide(blank)
    title(s3, "组合内图")
    group = s3.shapes.add_group_shape()
    group.shapes.add_picture(logo, Inches(1.2), Inches(2.2), Inches(1.5), Inches(0.9))
    inner = group.shapes.add_textbox(Inches(1.2), Inches(3.4), Inches(4.0), Inches(0.6))
    run = inner.text_frame.paragraphs[0].add_run()
    run.text = "组合内的说明文字"
    run.font.size = Pt(16)

    # 4 comment anchor page
    s4 = prs.slides.add_slide(blank)
    title(s4, "评审意见")
    body(s4, "本章数据来源需要补充说明。", name="note1")
    body(s4, "图表编号与正文不一致。", top=3.2, name="note2")

    # 5 filler
    s5 = prs.slides.add_slide(blank)
    title(s5, "结语")
    body(s5, "恳请各位专家批评指正。")

    prs.save(path)
    return swatch


def write_payloads(out):
    """Chinese review remarks belong in UTF-8 files written by python: a .ps1 driver
    cannot carry them (Windows PowerShell 5.1 reads its own source as ANSI)."""
    comments = {
        "items": [
            {"slide": 4, "shape": "note1", "text": "数据来源需补充：《江西省水文年鉴》2020—2024。"},
            {"slide": 4, "shape": "note2", "text": "图表编号与正文不一致，请统一为 图 3-1。"},
            {"slide": 1, "text": "封面色带建议加深一档，投影时更清楚。"},
        ]
    }
    with open(os.path.join(out, "comments.json"), "w", encoding="utf-8") as fh:
        json.dump(comments, fh, ensure_ascii=False, indent=1)
    bad = {
        "items": [
            {"slide": 4, "shape": "note1", "text": "这一条必须被回滚"},
            {"slide": 4, "shape": "nope", "text": "这个形状不存在"},
        ]
    }
    with open(os.path.join(out, "comments-bad.json"), "w", encoding="utf-8") as fh:
        json.dump(bad, fh, ensure_ascii=False, indent=1)


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)
    out = os.path.abspath(args.out)
    if not os.path.isdir(out):
        os.makedirs(out)
    deck = os.path.join(out, "deliver-deck.pptx")
    asset_dir = os.path.join(out, "assets")
    if not os.path.isdir(asset_dir):
        os.makedirs(asset_dir)
    swatch = build(deck, asset_dir)
    # ASCII file name on purpose: the replacement path reaches the skill on the command
    # line, which mangles Chinese.  The DIRECTORY is still "工作区", so the insert still
    # has to go through the ASCII temp copy.
    replacement = os.path.join(asset_dir, "new-image.png")
    png(replacement, 300, 200, (55, 86, 35))
    shutil.copyfile(deck, os.path.join(out, "deliver-deck.pristine.pptx"))
    write_payloads(out)
    print(swatch)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
