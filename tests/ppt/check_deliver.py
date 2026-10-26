#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Assertions for `ppt comment` / `export` / `picture` (tests/ppt).

  python check_deliver.py --dir <out-deliver>
"""
import argparse
import json
import os
import re
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


def expect_ok(step, doc):
    if doc is None:
        return False
    if not doc.get("ok"):
        fail("step %r failed: %s" % (step, json.dumps(doc.get("error"), ensure_ascii=False)))
        return False
    return True


def check_comments(folder):
    doc, added = load(folder, "comment-add")
    if expect_ok("comment-add", doc):
        if added.get("count") != 3:
            fail("comment add reported count=%r, expected 3" % added.get("count"))
        texts = [h.get("text", "") for h in added.get("added") or []]
        if not any("@note1" in t for t in texts):
            fail("a comment anchored to a shape does not carry the @shape prefix: %s" % texts)
        if not any("色带" in t for t in texts):
            fail("the slide-1 comment text did not survive the round trip: %s" % texts)

    doc, listed = load(folder, "comment-list")
    if expect_ok("comment-list", doc):
        if listed.get("total") != 3:
            fail("comment list total=%r, expected 3" % listed.get("total"))
        by_slide = {item["slide"]: len(item["comments"]) for item in listed.get("slides") or []}
        if by_slide.get(4) != 2 or by_slide.get(1) != 1:
            fail("comments landed on the wrong slides: %r" % by_slide)

    # the failing batch must roll back completely
    doc, _ = load(folder, "comment-bad")
    if doc is None or doc.get("ok"):
        fail("a batch naming a missing shape should fail")
    else:
        msg = json.dumps(doc.get("error"), ensure_ascii=False)
        if "rolled back" not in msg:
            fail("the failed batch did not report a rollback: %s" % msg)
    doc, after_bad = load(folder, "comment-list2")
    if expect_ok("comment-list2", doc):
        if after_bad.get("total") != 3:
            fail("after the failed batch there are %r comments, expected the original 3"
                 % after_bad.get("total"))

    doc, deleted = load(folder, "comment-del1")
    if expect_ok("comment-del1", doc):
        if deleted.get("deleted") != 1:
            fail("deleting one comment reported %r" % deleted.get("deleted"))
    doc, deleted_all = load(folder, "comment-delall")
    if expect_ok("comment-delall", doc):
        if deleted_all.get("deleted") != 2:
            fail("deleting all comments on slide 4 reported %r, expected 2" % deleted_all.get("deleted"))
        if deleted_all.get("left") != 0:
            fail("slide 4 still has %r comments" % deleted_all.get("left"))
    doc, final = load(folder, "comment-list3")
    if expect_ok("comment-list3", doc):
        if final.get("total") != 0:
            fail("after deleting everything %r comments remain" % final.get("total"))


def check_pictures(folder):
    doc, listed = load(folder, "picture-list")
    if expect_ok("picture-list", doc):
        # 1 on the cover + 2 on slide 2 + 1 inside the group on slide 3
        if listed.get("total") != 4:
            fail("picture list found %r pictures, expected 4 (cover 1, slide 2 two, group 1)"
                 % listed.get("total"))
        grouped = [p for item in listed.get("slides") or [] for p in item["pictures"] if p.get("group")]
        if len(grouped) != 1:
            fail("expected exactly one grouped picture in the list, got %d" % len(grouped))
        if grouped and "/" not in grouped[0].get("path", ""):
            fail("the grouped picture path does not show the group chain: %r" % grouped[0].get("path"))

    doc, dry = load(folder, "picture-dry")
    if expect_ok("picture-dry", doc):
        if dry.get("changed") != 0 or not dry.get("dryRun"):
            fail("picture -DryRun changed=%r dryRun=%r" % (dry.get("changed"), dry.get("dryRun")))
        if dry.get("wouldChange") != 3:
            fail("picture -DryRun wouldChange=%r, expected 3 (the three top-level pictures)"
                 % dry.get("wouldChange"))

    doc, replaced = load(folder, "picture-replace")
    if expect_ok("picture-replace", doc):
        if replaced.get("changed") != 3:
            fail("picture replace changed=%r, expected 3" % replaced.get("changed"))
        skipped = replaced.get("skipped") or []
        if not any("group" in json.dumps(s, ensure_ascii=False) for s in skipped):
            fail("the grouped picture was not reported as skipped: %r" % skipped)
        for hit in replaced.get("hits") or []:
            before, after = hit.get("before") or {}, hit.get("after") or {}
            for key in ("l", "t", "w", "h"):
                if key in before and key in after and abs(before[key] - after[key]) > 1.0:
                    fail("slide %s %s: geometry %s moved %.1f -> %.1f"
                         % (hit.get("slide"), hit.get("shape"), key, before[key], after[key]))

    # the saved deck: the pictures really were swapped, geometry kept, group untouched
    plain = os.path.join(folder, "deck.plain.pptx")
    if not os.path.exists(plain):
        fail("deck.plain.pptx missing (the driver must decrypt the saved deck)")
    else:
        from pptx import Presentation
        prs = Presentation(plain)
        pics = []
        for index, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                if shape.shape_type is not None and int(shape.shape_type) in (13, 11):
                    pics.append((index, shape))
        if len(pics) != 3:
            fail("the saved deck has %d top-level pictures, expected 3 (the grouped one is"
                 " inside the group, so a top-level count of 2 + group)" % len(pics))
        # the 300x200 source fills the old 3.0x2.0in slot, so the size is kept exactly
        for index, shape in pics:
            if index == 2:
                if abs(shape.width.inches - 3.0) > 0.05 or abs(shape.height.inches - 2.0) > 0.05:
                    fail("slide 2 picture size drifted to %.2fx%.2f in, expected 3.00x2.00"
                         % (shape.width.inches, shape.height.inches))


def check_exports(folder):
    doc, pdf = load(folder, "export-pdf")
    if expect_ok("export-pdf", doc):
        path = pdf.get("out") or ""
        if not os.path.exists(path):
            fail("export pdf did not produce %r" % path)
        if "encrypted" not in pdf:
            fail("export pdf does not report whether the file is encrypted")
        if pdf.get("slides") != 5:
            fail("export pdf reported %r slides, expected 5" % pdf.get("slides"))

    plain = os.path.join(folder, "out.plain.pdf")
    if not os.path.exists(plain):
        fail("out.plain.pdf missing")
    else:
        with open(plain, "rb") as fh:
            data = fh.read()
        if not data.startswith(b"%PDF"):
            fail("the decrypted PDF does not start with %PDF")
        pages = len(re.findall(rb"/Type\s*/Page[^s]", data))
        if pages != 5:
            fail("the PDF has %d page objects, expected 5" % pages)

    doc, pngs = load(folder, "export-png")
    if expect_ok("export-png", doc):
        files = pngs.get("files") or []
        if len(files) != 2:
            fail("export png produced %d files for -Slides 1-2" % len(files))
        for path in files:
            if not os.path.exists(path):
                fail("export png file missing: %s" % path)
                continue
            with open(path, "rb") as fh:
                head = fh.read(8)
            if head != b"\x89PNG\r\n\x1a\n":
                fail("export png wrote something that is not a PNG: %s" % path)
        # plaintext on purpose: WPS would have encrypted it
        for path in files:
            if os.path.exists(path):
                with open(path, "rb") as fh:
                    if fh.read(3) == b"\x17\xda\x5f":
                        fail("export png file is encrypted: %s" % path)

    doc, md = load(folder, "export-md")
    if expect_ok("export-md", doc):
        path = os.path.join(folder, "deck.md")
        if not os.path.exists(path):
            fail("export md did not write deck.md")
        else:
            with open(path, "r", encoding="utf-8") as fh:
                text = fh.read()
            if "洪水风险图" not in text:
                fail("export md did not contain the cover title")


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True)
    args = ap.parse_args(argv)
    folder = os.path.abspath(args.dir)
    if not os.path.isdir(folder):
        print("no such directory: %s" % folder)
        return 2

    check_comments(folder)
    check_pictures(folder)
    check_exports(folder)

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
