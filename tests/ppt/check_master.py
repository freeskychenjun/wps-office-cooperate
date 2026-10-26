#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Assertions for `ppt master` (tests/ppt).

  python check_master.py --dir <out-master>

What is checked, and why it is worth checking:
  list       the inventory has to describe THIS deck (1 design, 11 layouts, 4 slides)
  footer     a batch must reach every page except the cover, read back as applied, and
             a -Clear must undo it (a footer you cannot remove is a trap)
  layout     WPS only honours the legacy constants; the deck on disk has to show the
             new layout name on the target slide and keep the others alone
  rebase     the point of the whole command: the pages come out wearing the TEMPLATE's
             master ("TPL-MARK"), with their own text intact, and the template's sample
             pages gone
"""
import argparse
import json
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

_failures = []


def fail(message):
    _failures.append(message)


def load(folder, step):
    path = os.path.join(folder, "resp-%s.json" % step)
    if not os.path.exists(path):
        fail("missing response %r" % step)
        return None, {}
    with open(path, "r", encoding="utf-8") as fh:
        raw = fh.read()
    try:
        doc = json.loads(raw)
    except Exception as exc:
        fail("response %r is not JSON: %s" % (step, exc))
        return None, {}
    return doc, doc.get("result", {})


def ok(step, doc):
    if doc is None:
        return False
    if not doc.get("ok"):
        fail("step %r failed: %s" % (step, json.dumps(doc.get("error"), ensure_ascii=False)))
        return False
    return True


def check_list(folder):
    doc, listed = load(folder, "master-list")
    if not ok("master-list", doc):
        return
    if len(listed.get("designs") or []) < 1:
        fail("master list shows no design")
    if len(listed.get("layouts") or []) < 5:
        fail("master list shows %d layouts, expected the deck's set"
             % len(listed.get("layouts") or []))
    slides = listed.get("slides") or []
    if len(slides) != 4:
        fail("master list shows %d slides, expected 4" % len(slides))
    for entry in slides:
        if not entry.get("layout"):
            fail("slide %s has no layout name in the inventory" % entry.get("slide"))
        if entry.get("footerVisible") or entry.get("pageNumber"):
            fail("slide %s already had a footer before the batch" % entry.get("slide"))


def check_footer(folder):
    doc, applied = load(folder, "master-footer")
    if ok("master-footer", doc):
        if applied.get("footers") != 3:
            fail("footer batch touched %r pages, expected 3 (cover skipped)"
                 % applied.get("footers"))
        if applied.get("pageNumbers") != 3:
            fail("page number batch touched %r pages, expected 3" % applied.get("pageNumbers"))
        skipped = applied.get("skipped") or []
        if not any(s.get("reason") == "cover" for s in skipped):
            fail("the cover was not reported as skipped: %r" % skipped)

    doc, listed = load(folder, "master-list2")
    if ok("master-list2", doc):
        slides = {s["slide"]: s for s in listed.get("slides") or []}
        if slides.get(1, {}).get("footerVisible"):
            fail("the cover got a footer even though -SkipCover was given")
        for k in (2, 3, 4):
            entry = slides.get(k) or {}
            if entry.get("footer") != "ACME CORP":
                fail("slide %d footer reads %r" % (k, entry.get("footer")))
            if not entry.get("footerVisible") or not entry.get("pageNumber"):
                fail("slide %d did not get footer+number: %r" % (k, entry))

    doc, cleared = load(folder, "master-unfooter")
    if ok("master-unfooter", doc):
        if cleared.get("changed") != 4:
            fail("-Clear touched %r pages, expected all 4" % cleared.get("changed"))
    doc, listed = load(folder, "master-list3")
    if ok("master-list3", doc):
        for entry in listed.get("slides") or []:
            if entry.get("footerVisible") or entry.get("pageNumber"):
                fail("slide %s still shows a footer after -Clear: %r" % (entry.get("slide"), entry))
                break

    # and the deck saved to disk.  WPS does NOT write <p:hf>: switching a footer on
    # materialises a footer placeholder shape on the slide (and a slide-number one),
    # so that is what has to be asserted.
    plain = os.path.join(folder, "source.plain.pptx")
    if not os.path.exists(plain):
        fail("source.plain.pptx missing (decrypt step)")
        return
    from pptx import Presentation
    prs = Presentation(plain)
    for index, slide in enumerate(prs.slides, start=1):
        footers = []
        has_number = False
        for shape in slide.shapes:
            name = shape.name or ""
            text = shape.text_frame.text.strip() if shape.has_text_frame else ""
            if "页脚" in name or "Footer" in name:
                footers.append(text)
            if "编号" in name or "Slide Number" in name:
                has_number = True
        if index == 1:
            if any(footers) or has_number:
                fail("the saved cover shows a footer/page number: %r / %r" % (footers, has_number))
        else:
            if "ACME CORP" not in footers:
                fail("slide %d has footer placeholders %r, expected 'ACME CORP'" % (index, footers))
            if not has_number:
                fail("slide %d has no slide-number placeholder" % index)


def check_layout(folder):
    doc, dry = load(folder, "layout-dry")
    if ok("layout-dry", doc):
        details = dry.get("details") or []
        if not details:
            fail("layout -DryRun returned no details")
        elif details[0].get("would") != 1:
            fail("layout -DryRun did not report the wanted constant")

    doc, done = load(folder, "layout")
    if ok("layout", doc):
        details = done.get("details") or []
        if len(details) != 1 or details[0].get("slide") != 2:
            fail("layout switch did not touch exactly slide 2: %r" % details)
        else:
            entry = details[0]
            if entry.get("layoutBefore") == entry.get("layoutAfter"):
                fail("layout on slide 2 did not change (%r -> %r)"
                     % (entry.get("layoutBefore"), entry.get("layoutAfter")))
            if not entry.get("layoutAfter"):
                fail("no layoutAfter reported for slide 2")

    doc, bad = load(folder, "layout-bad")
    if doc is None or doc.get("ok"):
        fail("-Layout 42 should be refused")
    else:
        if "1..11" not in json.dumps(doc.get("error"), ensure_ascii=False):
            fail("the out-of-range layout error does not explain the 1..11 range")


def check_rebase(folder):
    doc, done = load(folder, "rebase")
    if not ok("rebase", doc):
        return
    if done.get("pasted") != 4:
        fail("rebase pasted %r pages, expected 4" % done.get("pasted"))
    if done.get("pages") != 4:
        fail("rebased deck has %r pages, expected 4 (template sample pages removed)"
             % done.get("pages"))

    plain = os.path.join(folder, "rebased.plain.pptx")
    if not os.path.exists(plain):
        fail("rebased.plain.pptx missing (decrypt step)")
        return
    from pptx import Presentation
    prs = Presentation(plain)
    titles = []
    for slide in prs.slides:
        text = ""
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                text = shape.text_frame.text.strip()
                break
        titles.append(text)
    if titles != ["S-1", "S-2", "S-3", "S-4"]:
        fail("rebased deck pages are %r, expected S-1..S-4" % titles)
    mark = ""
    for ph in prs.slide_master.placeholders:
        if ph.placeholder_format.type == 1:
            mark = ph.text_frame.text.strip()
            break
    if mark != "TPL-MARK":
        fail("rebased deck kept %r as its master, expected the template's 'TPL-MARK'" % mark)
    # the template's gold band lives on its slides, not its master, so it must be gone
    bands = sum(1 for slide in prs.slides for shape in slide.shapes
                if shape.has_text_frame is False and shape.shape_type is not None
                and int(shape.shape_type) == 1)
    if bands > 0:
        fail("the template's sample decorations came along (%d autoshapes)" % bands)


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True)
    args = ap.parse_args(argv)
    folder = os.path.abspath(args.dir)
    if not os.path.isdir(folder):
        print("no such directory: %s" % folder)
        return 2

    check_list(folder)
    check_footer(folder)
    check_layout(folder)
    check_rebase(folder)

    if _failures:
        print("")
        for line in _failures:
            print("  FAIL  %s" % line)
        print("\n%d check(s) failed" % len(_failures))
        return 1
    print("\nall checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
