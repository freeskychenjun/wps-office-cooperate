"""Acceptance check: word captions + the two field fixes it depends on.

Runs the WPS-backed helper against a COPY of captions-fixture.docx, so the committed
fixture is never touched:

    python check_captions.py            # full run
    python check_captions.py --verbose  # print every assertion

What it proves (each one is a defect from the 2026-10-26 field report):

  C1  `captions` lists every table with its caption, number, position and paragraph.
  C2  it reports the planted duplicate (3.3-3 twice), the planted gap (2.4.1 -> 2.4.3),
      the mixed dot/hyphen formats, the mixed spacing, the table without a caption and
      the dangling reference 2.9.9 -- and does not invent anything else.
  B1  a prose paragraph that directly follows a table is NOT reported inTable:true.
  C3  `read -SkipTableCells` returns exactly the non-table paragraphs, in one call.
  C4  `table -Action scan` gives a `para` that really is the table's first paragraph.
  A1  `search` paragraph numbers are verified: every hit's `at` offset really points at the
      needle inside the paragraph it reports (this is what used to drift on long tables).
  A2  `search -Text 表` finds the caption paragraphs that the field report missed
      (they are the ones the duplicate/gap analysis needs).
"""

import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
WORD_PS1 = os.path.join(SKILL, "scripts", "word.ps1")
FIXTURE = os.path.join(HERE, "captions-fixture.docx")
EXPECTED = os.path.join(HERE, "expected.json")
WORKDIR = os.path.join(SKILL, ".tmp", "captions-check")

VERBOSE = "--verbose" in sys.argv

passed = 0
failed = 0


def check(name, condition, detail=""):
    global passed, failed
    if condition:
        passed += 1
        if VERBOSE:
            print("  PASS  %s" % name)
    else:
        failed += 1
        print("  FAIL  %s%s" % (name, ("  <- " + str(detail)) if detail else ""))
    return bool(condition)


def word(command, doc=None, **params):
    """Run one helper command and return its `result` object (raises on ok=false)."""
    args = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", WORD_PS1, command]
    if doc:
        args += ["-Doc", doc]
    for key, value in params.items():
        if value is True:
            args.append("-" + key)
        elif value is False or value is None:
            continue
        else:
            args += ["-" + key, str(value)]
    proc = subprocess.run(args, capture_output=True)
    out = proc.stdout.decode("utf-8", "replace").strip()
    try:
        payload = json.loads(out)
    except ValueError:
        raise SystemExit("not JSON from word %s: %s / %s" % (command, out[:400], proc.stderr.decode("utf-8", "replace")[:400]))
    if not payload.get("ok"):
        raise SystemExit("word %s failed: %s" % (command, json.dumps(payload.get("error"), ensure_ascii=False)))
    return payload["result"]


def main():
    if not os.path.exists(FIXTURE):
        raise SystemExit("fixture missing: %s (run make_fixture.py first)" % FIXTURE)
    expected = json.load(open(EXPECTED, encoding="utf-8"))

    os.makedirs(WORKDIR, exist_ok=True)
    doc = os.path.join(WORKDIR, "captions-fixture.docx")
    shutil.copy2(FIXTURE, doc)

    print("== open ==")
    info = word("open", doc)
    check("open reports 20 tables", info.get("tables") == expected["tables"], info.get("tables"))

    # ---------------------------------------------------------------- C1 / C2 captions
    print("== C1/C2  captions: the list and the numbering verdict ==")
    cap = word("captions", doc)
    summ = cap["summary"]

    check("C1 tables listed = 20", summ["tables"] == 20, summ["tables"])
    check("C1 captions found = 19", summ["captions"] == 19, summ["captions"])
    check("C1 every table row has a para", all(isinstance(t.get("para"), int) and t["para"] > 0 for t in cap["tables"]),
          [t.get("para") for t in cap["tables"]])
    check("C1 caption positions 13 above / 6 in-table / 1 none",
          (summ["positions"]["above"], summ["positions"]["in-table"], summ["positions"]["none"]) == (13, 6, 1),
          summ["positions"])

    # every caption: label/number/title/pos must match the fixture ground truth
    by_number = {}
    for c in cap["captions"]:
        by_number.setdefault(c["number"], []).append(c)
    mismatches = []
    for want in expected["captions"]:
        got = by_number.get(want["number"], [])
        if not got:
            mismatches.append("%s: missing" % want["number"])
            continue
        hit = None
        for c in got:
            if (c["title"] == want["title"] and c["pos"] == want["pos"]):
                hit = c
                break
        if hit is None:
            mismatches.append("%s: title/pos %s != %s" % (want["number"], [(c["title"], c["pos"]) for c in got],
                                                          (want["title"], want["pos"])))
            continue
        if hit.get("table") != want["table"]:
            mismatches.append("%s: table %s != %s" % (want["number"], hit.get("table"), want["table"]))
        if hit.get("gap") != want["gap"]:
            mismatches.append("%s: gap %r != %r" % (want["number"], hit.get("gap"), want["gap"]))
    check("C1 all 19 captions match the fixture (number/title/pos/gap/table)", not mismatches, mismatches)

    dups = summ["duplicates"]
    check("C2 duplicate 3.3-3 reported once, on tables 14+15",
          len(dups) == 1 and dups[0]["number"] == "3.3-3" and dups[0]["tables"] == [14, 15], dups)

    gaps = summ["sequenceGaps"]
    check("C2 exactly one gap and it is 2.4.1 -> 2.4.3 missing 2.4.2",
          len(gaps) == 1 and gaps[0]["after"] == "2.4.1" and gaps[0]["before"] == "2.4.3"
          and gaps[0]["missing"] == ["2.4.2"], gaps)

    formats = dict((f["format"], f["count"]) for f in summ["formats"])
    check("C2 formats counted as dot=7 / hyphen=12", formats == {"dot": 7, "hyphen": 12}, formats)

    check("C2 table 19 reported as having no caption", summ["uncaptioned"] == [19], summ["uncaptioned"])
    check("C2 dangling reference 2.9.9 reported, nothing else",
          summ["unresolvedReferences"] == ["2.9.9"], summ["unresolvedReferences"])
    resolved = dict((r["number"], r["resolved"]) for r in cap["references"])
    wanted_refs = expected["defects"]["resolvedReferences"]
    check("C2 the four planted references resolve",
          all(resolved.get(n) is True for n in wanted_refs),
          dict((n, resolved.get(n)) for n in wanted_refs))

    rules = [i["rule"] for i in cap["issues"]]
    check("C2 issue rules are exactly the seven planted findings",
          sorted(set(rules)) == sorted(["caption-duplicate", "caption-gap", "caption-format-mixed",
                                        "caption-space-mixed", "caption-position-mixed",
                                        "caption-missing", "ref-dangling"]), rules)
    check("C2 every issue carries a suggestion", all(i.get("suggest") for i in cap["issues"]), rules)

    # ---------------------------------------------------------------- B1 / C3 read
    print("== B1/C3  read: inTable boundary and the cell-skipping sweep ==")
    full = word("read", doc, Count=0)
    items = dict((it["i"], it) for it in full["items"])
    check("B1 full read returns every paragraph", len(items) == full["paragraphs"], (len(items), full["paragraphs"]))

    prose_text = expected["defects"]["proseAfterTable"]["text"]
    after = [(i, it) for i, it in items.items() if it["text"].startswith(prose_text)]
    check("B1 the prose paragraph right after table 13 exists", len(after) == 1, [i for i, _ in after])
    if after:
        idx, it = after[0]
        check("B1 it is NOT flagged inTable (Range.End is exclusive)", not it.get("inTable"), it.get("inTable"))
        check("B1 the paragraph before it is the table's last cell",
              items.get(idx - 1, {}).get("inTable") is True, items.get(idx - 1, {}))

    prose = word("read", doc, Count=0, SkipTableCells=True)
    prose_paras = set(it["i"] for it in prose["items"])
    table_paras = set(i for i, it in items.items() if it.get("inTable"))
    check("C3 -SkipTableCells returns exactly the non-table paragraphs",
          prose_paras == set(items) - table_paras, len(prose_paras))
    check("C3 -SkipTableCells reports what it skipped",
          prose.get("skippedTableCells") == len(table_paras), (prose.get("skippedTableCells"), len(table_paras)))
    check("C3 the prose sweep is one call's worth of output", len(prose_paras) < 400, len(prose_paras))

    only = word("read", doc, Count=0, Only="\u8868\\d")   # paragraphs whose text has 表<digit>
    only_paras = set(it["i"] for it in only["items"])
    expect_only = set(i for i, it in items.items() if _has_table_number(it["text"]))
    check("C3 -Only <regex> selects the same paragraphs a manual scan would",
          only_paras == expect_only, (len(only_paras), len(expect_only)))

    # ---------------------------------------------------------------- C4 table scan
    print("== C4  table -Action scan: para = the table's first paragraph ==")
    scan = word("table", doc, Action="scan")
    bad = []
    for row in scan["tables"]:
        n, para = row["table"], row["para"]
        if items.get(para, {}).get("table") != n:
            bad.append((n, para, items.get(para, {}).get("table")))
    check("C4 every table's para is a paragraph read marks as that table", not bad, bad)

    # ---------------------------------------------------------------- A1 / A2 search
    print("== A1/A2  search: verified paragraph numbers and per-hit offsets ==")
    found = word("search", doc, Text="\u8868", Max=900)
    hits = found["hits"]
    check("A1 mapping is reported as verified", found.get("verified") is True, found.get("verified"))

    wrong = []
    for h in hits:
        text = items.get(h["para"], {}).get("text", "")
        at = h["at"]
        if not (0 <= at < len(text)) or text[at:at + 1] != "\u8868":
            wrong.append((h["para"], at, text[at:at + 4]))
    check("A1 every hit's `at` points at the needle inside the reported paragraph", not wrong, wrong[:6])

    seen_at = {}
    bad_dup = []
    for h in hits:
        key = (h["para"], h["at"])
        if key in seen_at:
            bad_dup.append(key)
        seen_at[key] = 1
    check("A1 two hits in one paragraph never share the same `at`", not bad_dup, bad_dup[:6])

    expect_hit_paras = set(i for i, it in items.items() if "\u8868" in it["text"])
    capped = set(c["para"] for c in found.get("cappedParas", []))
    missing = expect_hit_paras - set(h["para"] for h in hits) - capped
    check("A2 search finds every paragraph containing 表 (minus the capped ones)", not missing, sorted(missing)[:10])

    caption_paras = set(c["para"] for c in cap["captions"])
    check("A2 the four captions the field report missed are all hit",
          caption_paras.issubset(set(h["para"] for h in hits)),
          sorted(caption_paras - set(h["para"] for h in hits)))

    # ---------------------------------------------------------------- audit caption judge
    print("== audit: caption judge (defect D) ==")
    audit = word("audit", doc, Action="caption-missing,table-caption-in-row")
    by_rule = audit["byRule"]
    check("audit reports no spurious caption-missing for a captioned table",
          by_rule.get("caption-missing", 0) == 1, by_rule)
    check("audit reports the 6 in-row captions and only those",
          by_rule.get("table-caption-in-row", 0) == 6, by_rule)

    # ---------------------------------------------------------------- cleanup
    print("== close ==")
    word("close", doc)
    print("\n%d passed, %d failed" % (passed, failed))
    return 0 if failed == 0 else 1


TABLE_NUMBER = re.compile(u"\u8868[ \u3000]*\\d")


def _has_table_number(text):
    """True when the paragraph text contains 表 followed by a digit (whitespace allowed)."""
    return bool(TABLE_NUMBER.search(text or ""))


if __name__ == "__main__":
    sys.exit(main())
