#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Assertions for `ppt check` / `ppt outline` (tests/ppt).

The fixture (make_check_deck.py) puts one defect per rule on its own slide and keeps
two slides clean, so every assertion here is "this rule fired on THIS slide and
nowhere else".  False positives are failures too: a gate nobody trusts gets ignored.

Layout of the seeded deck (1-based slide numbers):
   1  clean cover (no slide number, which is allowed)
   2  overflow
   3  fontsize
   4  placeholder
   5  outframe
   6  pagenum      (the driver turns its number off, the others on)
   7  density
   8  contrast
   9  empty
  10  image-small
  11  (font that does not exist - see the limits note)
  12  title        (18pt against 32pt everywhere else)
  13  clean page

  python check_check.py --dir <out>
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


def load(folder, name):
    path = os.path.join(folder, "resp-%s.json" % name)
    if not os.path.exists(path):
        fail("missing response %r" % name)
        return {}
    with open(path, "r", encoding="utf-8") as fh:
        raw = fh.read()
    try:
        doc = json.loads(raw)
    except Exception as exc:
        fail("response %r is not JSON: %s" % (name, exc))
        return {}
    if not doc.get("ok"):
        fail("step %r failed: %s" % (name, json.dumps(doc.get("error"), ensure_ascii=False)))
        return {}
    return doc.get("result", {})


def rules_on(result, slide):
    return sorted({i["rule"] for i in result.get("issues", []) if i.get("slide") == slide})


def check_seeded_defects(result):
    expect = {
        2: "overflow",
        3: "fontsize",
        4: "placeholder",
        5: "outframe",
        6: "pagenum",
        7: "density",
        8: "contrast",
        9: "empty",
        10: "image-small",
        12: "title",
    }
    for slide, rule in expect.items():
        got = rules_on(result, slide)
        if rule not in got:
            fail("slide %d: rule %r did not fire (got %s)" % (slide, rule, got))

    # fontsize legitimately co-fires on 12: an 18pt title is both out of line with the
    # deck (rule `title`) and below the 28pt title floor (rule `fontsize`).  Two true
    # statements about one shape are not a false positive.
    allowed = {("fontsize", 12)}
    seen = {}
    for issue in result.get("issues", []):
        seen.setdefault(issue["rule"], set()).add(issue["slide"])
    for slide, rule in expect.items():
        slides = seen.get(rule, set())
        stray = sorted(s for s in slides if s != slide and (rule, s) not in allowed)
        if stray:
            fail("rule %r also fired on slides %s (only %d was seeded)" % (rule, stray, slide))


def check_clean_slides(result):
    for slide in (1, 13):
        got = rules_on(result, slide)
        if got:
            fail("clean slide %d reported %s (false positive)" % (slide, got))
    # slide 11's only seeded "defect" is a font WPS hides; nothing may fire there
    got = rules_on(result, 11)
    if got:
        fail("slide 11 reported %s, expected nothing" % got)


def check_severities(result):
    want = {
        "overflow": "blocker",
        "outframe": "blocker",
        "placeholder": "blocker",
        "fontsize": "major",
        "pagenum": "major",
        "empty": "major",
        "title": "major",
        "density": "minor",
        "contrast": "minor",
        "image-small": "minor",
    }
    for issue in result.get("issues", []):
        rule = issue["rule"]
        if rule in want and issue["severity"] != want[rule]:
            fail("rule %r reported severity %r, expected %r" % (rule, issue["severity"], want[rule]))
    for issue in result.get("issues", []):
        if not issue.get("fix"):
            fail("issue on slide %s has no fix hint" % issue.get("slide"))


def check_summary(result):
    summary = result.get("summary", {})
    counted = len(result.get("issues", []))
    total = summary.get("total")
    if total != counted:
        fail("summary.total=%r but %d issues returned" % (total, counted))
    parts = sum(summary.get(k, 0) for k in ("blocker", "major", "minor", "info"))
    if parts != total:
        fail("severity counts do not add up: %r vs total %r" % (summary, total))
    if summary.get("blocker", 0) < 1:
        fail("expected at least one blocker on this fixture")


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True)
    args = ap.parse_args(argv)
    folder = os.path.abspath(args.dir)
    if not os.path.isdir(folder):
        print("no such directory: %s" % folder)
        return 2

    main_run = load(folder, "check")
    if main_run:
        check_seeded_defects(main_run)
        check_clean_slides(main_run)
        check_severities(main_run)
        check_summary(main_run)
        print("  issues on the fixture: %s" % json.dumps(main_run.get("summary"), ensure_ascii=False))
        limits = main_run.get("limits") or []
        if not any("missing-font" in str(x) for x in limits):
            fail("the measured missing-font limitation is not declared in `limits`")
        fonts = main_run.get("fonts") or []
        if not fonts:
            fail("no font inventory in the result")
        if main_run.get("verdict") is None:
            fail("no verdict line")

    # -MinSeverity blocker keeps only the blockers
    blocked = load(folder, "check-blocker")
    if blocked:
        for issue in blocked.get("issues", []):
            if issue["severity"] != "blocker":
                fail("-MinSeverity blocker returned a %r" % issue["severity"])
                break
        if blocked.get("summary", {}).get("major"):
            fail("-MinSeverity blocker still counts majors")
        if len(blocked.get("issues", [])) != blocked.get("summary", {}).get("blocker"):
            fail("-MinSeverity blocker: returned list and count disagree")

    # -Rules narrows the engine
    narrowed = load(folder, "check-rules")
    if narrowed:
        rules = {i["rule"] for i in narrowed.get("issues", [])}
        if not rules.issubset({"overflow", "outframe"}):
            fail("-Rules overflow,outframe returned %s" % sorted(rules))
        if "overflow" not in rules or "outframe" not in rules:
            fail("-Rules overflow,outframe missed one of them: %s" % sorted(rules))

    # -Profile bid (投标) raises the body floor to 18pt, so more slides trip fontsize
    base = {i["slide"] for i in main_run.get("issues", []) if i["rule"] == "fontsize"}
    bid = load(folder, "check-profile")
    if bid:
        if bid.get("profile") not in ("投标", "bid"):
            fail("-Profile was not applied: %r" % bid.get("profile"))
        raised = {i["slide"] for i in bid.get("issues", []) if i["rule"] == "fontsize"}
        if not raised > base:
            fail("-Profile bid should flag more fontsize issues than the default (%s vs %s)" % (raised, base))

    # -Config: a lower body floor must silence slide 3, and the whitelisted shape on
    # slide 5 must move to `suppressed` instead of disappearing (design intent stays
    # visible in the report, it just stops counting against delivery).
    cfg = load(folder, "check-config")
    if cfg:
        open_rules = {i["rule"] for i in cfg.get("issues", []) if i.get("slide") == 3}
        if "fontsize" in open_rules:
            fail("-Config bodyMin=9 did not raise the floor: slide 3 still reports fontsize")
        remaining = [i for i in cfg.get("issues", [])
                     if i.get("slide") == 5 and i.get("shape") == "outbox"]
        if remaining:
            fail("-Config suppression did not hide the whitelisted shape on slide 5")
        hidden = [i for i in cfg.get("suppressed") or []
                  if i.get("slide") == 5 and i.get("shape") == "outbox"]
        if not hidden:
            fail("-Config suppression dropped the issue instead of listing it as suppressed")
        elif not hidden[0].get("suppressedBy"):
            fail("suppressed entry carries no suppressedBy label")
        if cfg.get("summary", {}).get("suppressed") != len(cfg.get("suppressed") or []):
            fail("summary.suppressed does not match the suppressed list")

    # markdown report
    md_path = os.path.join(folder, "report.md")
    if not os.path.exists(md_path):
        fail("check -Out did not write report.md")
    else:
        with open(md_path, "r", encoding="utf-8") as fh:
            md = fh.read()
        for token in ("blocker", "major", "minor", "| slide |"):
            if token not in md:
                fail("report.md lacks %r" % token)
        if "overflow" not in md or "outframe" not in md:
            fail("report.md does not list the seeded rules")

    # outline
    outline = load(folder, "outline")
    if outline:
        slides = outline.get("slides") or []
        if len(slides) < 13:
            fail("outline returned %d slides, expected 13" % len(slides))
        first = slides[0] if slides else {}
        if "洪水风险图" not in (first.get("title") or ""):
            fail("outline did not pick the cover title: %r" % first.get("title"))
        with_text = [s for s in slides if s.get("bullets")]
        if not with_text:
            fail("outline returned no bullets at all")
    ol_path = os.path.join(folder, "outline.md")
    if not os.path.exists(ol_path):
        fail("outline -Out did not write outline.md")
    else:
        with open(ol_path, "r", encoding="utf-8") as fh:
            text = fh.read()
        if "洪水风险图" not in text or not text.lstrip().startswith("##"):
            fail("outline.md does not look like a markdown outline")

    # the exit-code gate
    gate_path = os.path.join(folder, "gate.json")
    if not os.path.exists(gate_path):
        fail("gate.json missing (the driver must record -ExitCode)")
    else:
        with open(gate_path, "r", encoding="utf-8") as fh:
            gate = json.loads(fh.read())
        if gate.get("exitCode") != 3:
            fail("check -ExitCode returned %r, expected 3 (blockers present)" % gate.get("exitCode"))

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
