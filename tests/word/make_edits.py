#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Build corpus/edits.json (10 revise items) for the wps-word-cooperate acceptance run.

Contract (DESIGN.md section 3):
  {"items":[{"para":12,"expect":"...","op":"replace","find":"..","replace":".."},
            {"para":15,"expect":"...","op":"set_text","text":"..."},
            {"para":20,"expect":"...","op":"insert_after","text":"..."}]}

`para` is the 1-based WPS Document.Paragraphs index (table cell paragraphs are
counted inline, exactly like WPS does).  It is taken from corpus/meta.json so it
never has to be guessed.

Composition: 5 x replace (3 targets + 2 extras), 3 x set_text, 2 x insert_after.
The two insertions are placed AFTER every other edited paragraph, so paragraph
indices of the earlier items are unaffected no matter how the helper applies the
batch (original numbering or live numbering).
"""

import json
import os
import sys

from docx import Document

HERE = os.path.dirname(os.path.abspath(__file__))
CORPUS = os.path.join(HERE, "corpus")
DOCX = os.path.join(CORPUS, "report-300.docx")
META = os.path.join(CORPUS, "meta.json")
TARGETS = os.path.join(CORPUS, "targets.json")
EDITS = os.path.join(CORPUS, "edits.json")

EXPECT_LEN = 16

INSERT_TEXTS = {
    "ins_1": "本次复核补充了 2020 年以来三次洪水的洪痕调查成果，可作为设计洪水复核的参考依据。",
    "ins_2": "建议建设单位在开工前完成征地移民与临时用地手续，保障工程按计划节点顺利推进。",
}


def prefix(text, n=EXPECT_LEN):
    t = text.strip()
    p = t[:n].rstrip()
    if len(p) < 10:
        raise SystemExit("paragraph too short for expect: %r" % text)
    return p


def main():
    meta = json.load(open(META, encoding="utf-8"))
    tags = meta["tags"]
    targets = json.load(open(TARGETS, encoding="utf-8"))
    doc = Document(DOCX)

    def py_text(tag):
        py = tags[tag]["py"]
        return doc.paragraphs[py].text.strip()

    def wps_para(tag):
        return int(tags[tag]["wps"])

    # ---- 5 x replace ------------------------------------------------------
    rep = {
        "target_A": ("孔口高程", "孔口地面高程"),
        "target_B": ("大概 1234 m3/s", "约 1234 m³/s"),
        "target_C": ("进行了进行", "进行了"),
        "r4": ("河段比较顺直", "河段较为顺直"),
        "r5": ("建议下阶段补充", "建议下阶段进一步补充"),
    }
    items = []
    for tag, (find, repl) in rep.items():
        text = py_text(tag)
        assert text.count(find) == 1, "%s: find must occur exactly once" % tag
        items.append({
            "para": wps_para(tag),
            "expect": prefix(text),
            "op": "replace",
            "find": find,
            "replace": repl,
        })

    # ---- 3 x set_text -----------------------------------------------------
    for tag in ("set_1", "set_2", "set_3"):
        old = py_text(tag)
        new = "经复核，" + old.rstrip("。") + "，相关成果与上一阶段报告核对一致。"
        assert new != old
        items.append({
            "para": wps_para(tag),
            "expect": prefix(old),
            "op": "set_text",
            "text": new,
        })

    # ---- 2 x insert_after -------------------------------------------------
    for tag in ("ins_1", "ins_2"):
        anchor = py_text(tag)
        items.append({
            "para": wps_para(tag),
            "expect": prefix(anchor),
            "op": "insert_after",
            "text": INSERT_TEXTS[tag],
        })

    items.sort(key=lambda it: it["para"])

    # ---- guards -----------------------------------------------------------
    assert len(items) == 10
    assert sum(1 for i in items if i["op"] == "replace") == 5
    assert sum(1 for i in items if i["op"] == "set_text") == 3
    assert sum(1 for i in items if i["op"] == "insert_after") == 2
    assert len({i["para"] for i in items}) == 10, "paragraph numbers must be distinct"

    ins = [i["para"] for i in items if i["op"] == "insert_after"]
    others = [i["para"] for i in items if i["op"] != "insert_after"]
    assert len(ins) == 2 and ins == sorted(ins)
    assert min(ins) > max(others), "insert_after anchors must follow all other edits"

    # expect must be a real prefix of the pristine paragraph text
    for it in items:
        hits = [t for t in targets if t["para"] == it["para"]]
        if hits:
            original = hits[0]["original"]
        else:
            match = [v for v in tags.values() if int(v["wps"]) == it["para"]]
            assert match, "para %d has no tag record" % it["para"]
            original = match[0]["text"].strip()
        assert original.startswith(it["expect"]), \
            "expect is not a true prefix for para %d: %r vs %r" % (it["para"], it["expect"], original[:30])
        assert 10 <= len(it["expect"]) <= 18, "expect length out of range"

    with open(EDITS, "w", encoding="utf-8") as fh:
        json.dump({"items": items}, fh, ensure_ascii=False, indent=2)

    summary = {
        "edits": EDITS,
        "count": len(items),
        "ops": {op: sum(1 for i in items if i["op"] == op)
                for op in ("replace", "set_text", "insert_after")},
        "paras": [i["para"] for i in items],
        "expects": [{"para": i["para"], "op": i["op"], "expect": i["expect"]} for i in items],
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
