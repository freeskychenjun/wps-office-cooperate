# -*- coding: utf-8 -*-
"""Temporary positive self-test: apply edits.json to the corpus with WPS COM
(TrackRevisions=True) and add 2 comments, producing a *revised* copy.

This is NOT the skill helper - it only exists to prove validate.py can pass.
"""
import json
import os
import shutil
import sys
import time

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
CORPUS = os.path.join(HERE, "corpus")
SRC = os.path.join(CORPUS, "report-300.docx")
DST = os.path.join(CORPUS, "_selftest-revised.docx")
EDITS = os.path.join(CORPUS, "edits.json")


def main():
    if os.path.exists(DST):
        os.remove(DST)
    shutil.copy2(SRC, DST)

    items = json.load(open(EDITS, encoding="utf-8"))["items"]
    items.sort(key=lambda it: it["para"])

    app = win32.Dispatch("KWPS.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = 0
    except Exception:
        pass

    log = []
    doc = app.Documents.Open(DST, False, False, False)
    try:
        doc.TrackRevisions = True
        offset = 0
        for it in items:
            live = it["para"] + offset
            para = doc.Paragraphs(live)
            text = para.Range.Text.replace("\r", "").replace("\x07", "")
            if not text.startswith(it["expect"]):
                raise SystemExit("ANCHOR_MISSING para %d: %r" % (live, text[:40]))
            if it["op"] == "replace":
                rng = para.Range
                rng.Find.ClearFormatting()
                rng.Find.Replacement.ClearFormatting()
                ok = rng.Find.Execute(it["find"], True, False, False, False, False,
                                      True, 0, False, it["replace"], 2)
                log.append((live, "replace", bool(ok)))
            elif it["op"] == "set_text":
                r = para.Range
                doc.Range(r.Start, r.End - 1).Text = it["text"]
                log.append((live, "set_text", True))
            elif it["op"] == "insert_after":
                para.Range.InsertParagraphAfter()
                newp = doc.Paragraphs(live + 1)
                r = newp.Range
                doc.Range(r.Start, r.End - 1).Text = it["text"]
                offset += 1
                log.append((live, "insert_after", True))

        doc.Comments.Add(doc.Paragraphs(151).Range, "建议核实孔口地面高程与实测断面的对应关系，并补充测量成果。")
        doc.Comments.Add(doc.Paragraphs(86).Range, "洪峰流量单位应统一为 m³/s，并与水文计算成果表核对。")

        doc.Save()
        log.append(("revisions", int(doc.Revisions.Count), ""))
        log.append(("comments", int(doc.Comments.Count), ""))
    finally:
        doc.Close(0)
        try:
            left = int(app.Documents.Count)
            if left == 0:
                app.Quit()
                time.sleep(1.0)
            log.append(("documents_left", left, ""))
        except Exception as exc:
            log.append(("quit_err", str(exc), ""))
        with open(os.path.join(CORPUS, "_selftest.log"), "w", encoding="utf-8") as fh:
            fh.write("\n".join(repr(x) for x in log))


if __name__ == "__main__":
    main()
