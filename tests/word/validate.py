#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Independent acceptance validator for the `wps-word-cooperate` skill (COM / WPS Writer).

Usage:
    python validate.py --doc <revised.docx> --targets corpus/targets.json \
        --edits corpus/edits.json [--expect-comments 2] [--png-dir <dir>]

Prints one JSON report to stdout and writes tests/validate-report.json.
Exit code: 0 = pass, 1 = fail, 2 = bad arguments.

Why COM: WPS re-saves docx as ciphertext (header 17DA5FA01633CD9A) on this
machine, so python-docx / zipfile cannot open the revised file.  The original
corpus is plaintext, but the revised copy is not -- everything here goes through
Word-compatible automation.

Tracked changes are judged on a *copy*: <doc>.accept-probe.docx is opened,
Revisions.AcceptAll() is applied there, the text is read, and the copy is
deleted.  The delivered document itself is never modified.
"""

import argparse
import gc
import json
import os
import shutil
import subprocess
import sys
import time
import traceback

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_PATH = os.path.join(SCRIPT_DIR, "validate-report.json")

CLEAN_CHARS = ("\r", "\x07", "\x0b", "\x0c", "\x0e")


# --------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------
def count_wps():
    """Number of running wps.exe processes (the machine-wide discipline check)."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq wps.exe", "/NH"],
            capture_output=True, timeout=20,
        ).stdout
        return out.lower().count(b"wps.exe")
    except Exception:
        return -1


def clean_text(t):
    if not t:
        return ""
    for ch in CLEAN_CHARS:
        t = t.replace(ch, "")
    return t.strip()


def norm(t):
    return " ".join(clean_text(t).split())


def read_paragraphs(doc):
    n = int(doc.Paragraphs.Count)
    texts = []
    for i in range(1, n + 1):
        try:
            texts.append(clean_text(doc.Paragraphs(i).Range.Text))
        except Exception:
            texts.append("")
    return texts


def read_pages(doc):
    try:
        return int(doc.ComputeStatistics(2))
    except Exception:
        try:
            doc.Repaginate()
            return int(doc.ComputeStatistics(2))
        except Exception:
            return None


def read_revisions(doc):
    res = {"total": 0, "by_type": {}, "sample": []}
    try:
        rev = doc.Revisions
        total = int(rev.Count)
    except Exception as exc:
        res["error"] = "%s: %s" % (type(exc).__name__, exc)
        return res
    res["total"] = total
    for i in range(1, total + 1):
        try:
            item = rev.Item(i)
        except Exception:
            try:
                item = rev(i)
            except Exception:
                continue
        try:
            ty = int(item.Type)
        except Exception:
            ty = -1
        key = str(ty)
        res["by_type"][key] = res["by_type"].get(key, 0) + 1
        if len(res["sample"]) < 30:
            prefix = ""
            try:
                prefix = clean_text(item.Range.Text)[:24]
            except Exception:
                pass
            res["sample"].append({"index": i, "type": ty, "text_prefix": prefix})
    return res


def read_comments(doc):
    out = {"count": 0, "items": []}
    try:
        cm = doc.Comments
        cnt = int(cm.Count)
    except Exception as exc:
        out["error"] = "%s: %s" % (type(exc).__name__, exc)
        return out
    out["count"] = cnt
    for i in range(1, cnt + 1):
        scope = ""
        body = ""
        try:
            c = cm.Item(i)
        except Exception:
            try:
                c = cm(i)
            except Exception:
                continue
        try:
            scope = clean_text(c.Scope.Text)
        except Exception:
            try:
                scope = clean_text(c.Range.Text)
            except Exception:
                pass
        try:
            body = clean_text(c.Range.Text)
        except Exception:
            pass
        out["items"].append({
            "index": i,
            "scope_prefix": scope[:20],
            "text_prefix": body[:40],
        })
    return out


def check_png_dir(dirpath):
    from PIL import Image

    results = []
    for name in sorted(os.listdir(dirpath)):
        if not name.lower().endswith(".png"):
            continue
        path = os.path.join(dirpath, name)
        try:
            img = Image.open(path).convert("L")
            hist = img.histogram()
            total = sum(hist) or 1
            nonwhite = sum(hist[:245])  # darker than near-white
            ratio = nonwhite / float(total)
            results.append({
                "file": name,
                "width": img.width,
                "height": img.height,
                "nonwhite_ratio": round(ratio, 6),
                "status": "blank" if ratio < 0.005 else "ok",
            })
        except Exception as exc:
            results.append({
                "file": name,
                "status": "error",
                "error": "%s: %s" % (type(exc).__name__, exc),
            })
    return results


# --------------------------------------------------------------------------
# revision-aware index mapping
# --------------------------------------------------------------------------
def make_index_resolver(items):
    inserts = sorted(it["para"] for it in items if it["op"] == "insert_after")

    def resolve(item):
        """Return 1-based index in the accepted document."""
        n = item["para"]
        shift = sum(1 for p in inserts if p < n)
        if item["op"] == "insert_after":
            return n + shift + 1
        return n + shift

    return resolve


def evaluate_edits(items, target_by_para, accepted):
    resolve = make_index_resolver(items)
    results = []
    for it in items:
        idx = resolve(it)
        if it["op"] == "replace":
            tgt = target_by_para.get(it["para"])
            expected = tgt["original"].replace(tgt["find"], tgt["replace"]) if tgt else None
        else:
            expected = it["text"]
        actual = accepted[idx - 1] if 0 <= idx - 1 < len(accepted) else None
        ok = expected is not None and actual is not None and norm(actual) == norm(expected)
        entry = {
            "para": it["para"],
            "op": it["op"],
            "index_in_accepted": idx,
            "ok": ok,
        }
        if not ok:
            if expected is None:
                entry["error"] = "no targets.json record for this paragraph"
            entry["expected_prefix"] = (expected or "")[:40]
            entry["actual_prefix"] = (actual or "<missing>")[:40]
            if expected is not None:
                for j, txt in enumerate(accepted):
                    if norm(txt) == norm(expected):
                        entry["found_at_index"] = j + 1
                        break
        results.append(entry)
    return results


def evaluate_targets(targets, accepted, insert_paras):
    """targets_hit on the accepted view (tracked view keeps deleted characters)."""
    out = []
    for tgt in targets:
        shift = sum(1 for p in insert_paras if p < tgt["para"])
        idx = tgt["para"] + shift
        actual = accepted[idx - 1] if 0 <= idx - 1 < len(accepted) else ""
        expected = tgt["original"].replace(tgt["find"], tgt["replace"])
        out.append({
            "para": tgt["para"],
            "kind": tgt.get("kind", "target"),
            "find": tgt["find"],
            "replace": tgt["replace"],
            "hit": norm(actual) == norm(expected),
            "text_prefix": actual[:40],
        })
    return out


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="wps-word-cooperate acceptance validator")
    ap.add_argument("--doc", required=True, help="revised docx produced by the AI/helper")
    ap.add_argument("--targets", required=True, help="corpus/targets.json")
    ap.add_argument("--edits", required=True, help="corpus/edits.json")
    ap.add_argument("--expect-comments", type=int, default=None)
    ap.add_argument("--png-dir", default=None)
    args = ap.parse_args()

    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    failures = []
    report = {
        "doc": os.path.abspath(args.doc),
        "targets_file": os.path.abspath(args.targets),
        "edits_file": os.path.abspath(args.edits),
        "paragraphs": None,
        "pages": None,
        "revisions": {"total": 0, "by_type": {}, "sample": []},
        "tracked_diff": {"items": [], "all_ok": False},
        "comments": {"count": 0, "items": []},
        "targets_hit": [],
        "png": [],
        "com": {},
        "verdict": "fail",
        "failures": failures,
    }

    for path in (args.doc, args.targets, args.edits):
        if not os.path.isfile(path):
            report["failures"] = ["file not found: %s" % path]
            write_report(report)
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return 2

    with open(args.targets, encoding="utf-8") as fh:
        targets = json.load(fh)
    with open(args.edits, encoding="utf-8") as fh:
        edits = json.load(fh)["items"]
    target_by_para = {t["para"]: t for t in targets}

    proc_before = count_wps()
    report["com"]["wps_procs_before"] = proc_before

    app = None
    opened = []
    probe_path = os.path.join(
        os.path.dirname(os.path.abspath(args.doc)),
        os.path.splitext(os.path.basename(args.doc))[0] + ".accept-probe.docx",
    )

    try:
        import win32com.client as win32
        import pythoncom  # noqa: F401
    except Exception as exc:
        report["failures"] = ["pywin32 unavailable: %s" % exc]
        report["com"]["wps_procs_after"] = count_wps()
        write_report(report)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 2

    try:
        app = win32.Dispatch("KWPS.Application")
        for attr, value in (("Visible", False), ("DisplayAlerts", 0), ("ScreenUpdating", False)):
            try:
                setattr(app, attr, value)
            except Exception:
                pass
        report["com"]["app_name"] = str(getattr(app, "Name", "?"))
        report["com"]["app_version"] = str(getattr(app, "Version", "?"))
        try:
            report["com"]["app_build"] = str(getattr(app, "Build", "?"))
        except Exception:
            pass

        doc = app.Documents.Open(os.path.abspath(args.doc), False, True, False)
        opened.append(doc)

        report["paragraphs"] = int(doc.Paragraphs.Count)
        report["pages"] = read_pages(doc)
        report["revisions"] = read_revisions(doc)
        report["comments"] = read_comments(doc)
        tracked = read_paragraphs(doc)

        # ---- accepted-view probe (copy on disk, never the original) --------
        accepted = None
        probe_note = None
        try:
            if os.path.exists(probe_path):
                os.remove(probe_path)
            shutil.copy2(args.doc, probe_path)
            probe = app.Documents.Open(probe_path, False, False, False)
            try:
                probe.Revisions.AcceptAll()
                accepted = read_paragraphs(probe)
                report["accepted_paragraphs"] = len(accepted)
            finally:
                try:
                    probe.Close(0)
                except Exception:
                    pass
            # close may be async on WPS; retry deletion briefly
            for attempt in range(6):
                try:
                    if os.path.exists(probe_path):
                        os.remove(probe_path)
                    break
                except Exception:
                    time.sleep(0.4)
            if os.path.exists(probe_path):
                probe_note = "probe copy could not be deleted: %s" % probe_path
        except Exception as exc:
            probe_note = "accept probe failed: %s: %s" % (type(exc).__name__, exc)

        if probe_note:
            report["tracked_diff"]["probe_error"] = probe_note

        if accepted is not None:
            report["tracked_diff"]["items"] = evaluate_edits(edits, target_by_para, accepted)
            report["tracked_diff"]["all_ok"] = all(i["ok"] for i in report["tracked_diff"]["items"])

            # targets_hit is judged on the accepted view (tracked view still
            # contains the deleted characters in Word/WPS)
            ins_items = [it for it in edits if it["op"] == "insert_after"]
            report["targets_hit"] = evaluate_targets(
                targets, accepted, [i["para"] for i in ins_items]
            )

            if len(accepted) != len(tracked) + len(ins_items):
                report["tracked_diff"]["paragraph_count_note"] = (
                    "accepted=%d tracked=%d inserts=%d" % (len(accepted), len(tracked), len(ins_items))
                )
        else:
            report["tracked_diff"]["all_ok"] = False

        if args.png_dir:
            if os.path.isdir(args.png_dir):
                report["png"] = check_png_dir(args.png_dir)
            else:
                report["png"] = [{"error": "png dir not found: %s" % args.png_dir}]

        # ---- verdict -------------------------------------------------------
        if report["revisions"]["total"] < len(edits):
            failures.append(
                "revisions: %d < %d (at least one revision per item; a real run needs ~2 per replace)"
                % (report["revisions"]["total"], len(edits))
            )
        if not report["tracked_diff"]["all_ok"]:
            bad = [i for i in report["tracked_diff"]["items"] if not i["ok"]]
            if report["tracked_diff"].get("probe_error"):
                failures.append("tracked_diff: %s" % report["tracked_diff"]["probe_error"])
            for i in bad[:12]:
                failures.append(
                    "tracked_diff: para %d (%s) does not match accepted text" % (i["para"], i["op"])
                )
            if len(bad) > 12:
                failures.append("tracked_diff: ... and %d more mismatches" % (len(bad) - 12))
        missing = [h for h in report["targets_hit"] if not h["hit"]]
        for h in missing:
            failures.append(
                "targets_hit: para %d not rewritten (%r -> %r)" % (h["para"], h["find"], h["replace"])
            )
        if args.expect_comments is not None and report["comments"]["count"] != args.expect_comments:
            failures.append(
                "comments: %d != expected %d" % (report["comments"]["count"], args.expect_comments)
            )
        for p in report["png"]:
            if p.get("status") == "blank":
                failures.append("png: %s is blank (nonwhite=%.4f%%)" % (p["file"], p["nonwhite_ratio"] * 100))
            elif p.get("status") == "error":
                failures.append("png: %s unreadable (%s)" % (p["file"], p.get("error")))
        if report["pages"] is None or report["pages"] < 1:
            failures.append("pages: could not be computed")

    except Exception as exc:
        failures.append("COM error: %s: %s" % (type(exc).__name__, exc))
        report["com"]["traceback"] = traceback.format_exc()[-1500:]
    finally:
        # close only the documents we opened
        for d in reversed(opened):
            try:
                d.Close(0)  # wdDoNotSaveChanges
            except Exception:
                pass
        quit_done = False
        remaining = None
        try:
            if app is not None:
                # WPS close is asynchronous: Documents.Count can still be 1
                # for a moment after Close(0).  Poll briefly before deciding.
                for _ in range(12):
                    try:
                        remaining = int(app.Documents.Count)
                    except Exception:
                        remaining = None
                        break
                    if remaining == 0:
                        break
                    time.sleep(0.25)
                if remaining == 0:
                    app.Quit()
                    quit_done = True
                    time.sleep(2.0)
        except Exception as exc:
            report["com"]["quit_error"] = "%s: %s" % (type(exc).__name__, exc)
        try:
            if os.path.exists(probe_path):
                os.remove(probe_path)
        except Exception:
            pass
        report["com"]["quit"] = quit_done
        report["com"]["documents_left"] = remaining
        if not quit_done and remaining:
            report["com"]["quit_skipped_because"] = (
                "%d other document(s) still open - not ours to close" % remaining
            )
        report["com"]["wps_procs_after"] = count_wps()
        try:
            del app
        except Exception:
            pass
        gc.collect()

    report["verdict"] = "pass" if not failures else "fail"
    write_report(report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["verdict"] == "pass" else 1


def write_report(report):
    try:
        with open(REPORT_PATH, "w", encoding="utf-8") as fh:
            json.dump(report, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass


if __name__ == "__main__":
    sys.exit(main())
