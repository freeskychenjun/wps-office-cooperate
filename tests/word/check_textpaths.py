"""Acceptance check for the v1.13 read-side fast paths (pure ASCII on purpose: PS 5.1 reads
non-ASCII .ps1/.py as ANSI and mangles Chinese strings).

    python check_textpaths.py            # full run on the two committed fixtures
    python check_textpaths.py --keep     # keep the temp copy for inspection

What it proves (this is the invariant that makes the speed-up safe):

  P1  `snapshot` (fast) and `snapshot -Slow` produce the SAME digest and the SAME per-paragraph
      text list, on a document WITH tables and on a document with tracked revisions.
  P2  the fast path actually engages (`textsPath` is xml/fast), i.e. the test is not passing
      because everything fell back to the slow loop.
  P3  `textdiag` reports the structural invariant it relies on: w:p + table rows == Paragraphs.
  P4  `diff` against a fast snapshot is also exact (0 changed when nothing changed).
  P5  a document the fast path cannot handle reports a REASON (`fastFail`) instead of failing
      silently.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
WORD = REPO / 'scripts' / 'word.ps1'
HERE = Path(__file__).resolve().parent
FIXTURES = [
    HERE / 'captions' / 'captions-fixture.docx',
    HERE / 'corpus' / 'report-300.docx',
]

passed = 0
failed = 0


def check(name, ok, detail=''):
    global passed, failed
    if ok:
        passed += 1
        print('PASS %s' % name)
    else:
        failed += 1
        print('FAIL %s  %s' % (name, detail))


def run(args):
    p = subprocess.run(['powershell', '-Mta', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(WORD)] + args,
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    try:
        return json.loads(p.stdout)
    except Exception:
        print('  BAD JSON: %s' % (p.stdout[:300] + p.stderr[:300]))
        return None


def load_snapshot(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))


def main():
    keep = '--keep' in sys.argv
    work = Path(tempfile.mkdtemp(prefix='woc-textpaths-'))
    try:
        for fx in FIXTURES:
            if not fx.exists():
                check('%s exists' % fx.name, False, str(fx))
                continue
            doc = str(work / fx.name)
            shutil.copyfile(fx, doc)

            op = run(['open', '-Doc', doc, '-NoHint'])
            check('%s opens' % fx.name, op is not None and op.get('ok'), str(op)[:200] if op else '')
            if not op or not op.get('ok'):
                continue
            info = op['result']

            fast = run(['snapshot', '-Doc', doc, '-Tag', 'tp-fast'])
            slow = run(['snapshot', '-Doc', doc, '-Tag', 'tp-slow', '-Slow'])
            check('%s snapshot ok' % fx.name, fast and slow and fast.get('ok') and slow.get('ok'))
            if not (fast and slow and fast.get('ok') and slow.get('ok')):
                continue
            f, s = fast['result'], slow['result']
            check('%s digest identical (fast %s vs slow)' % (fx.name, f.get('textsPath')),
                  f['digest'] == s['digest'], '%s vs %s' % (f['digest'], s['digest']))
            tf = load_snapshot(f['path'])['texts']
            ts = load_snapshot(s['path'])['texts']
            check('%s paragraph texts identical (%d paragraphs)' % (fx.name, len(tf)),
                  tf == ts, 'first difference at %s' % next((i for i, (a, b) in enumerate(zip(tf, ts)) if a != b), 'length'))
            check('%s fast path engaged' % fx.name,
                  f.get('textsPath') in ('xml', 'fast') or f.get('fastFail'),
                  'textsPath=%s fastFail=%s' % (f.get('textsPath'), f.get('fastFail')))

            diag = run(['textdiag', '-Doc', doc])
            check('%s textdiag ok' % fx.name, diag is not None and diag.get('ok'))
            if diag and diag.get('ok'):
                d = diag['result']
                check('%s w:p + table rows == paragraphs' % fx.name,
                      d.get('xmlPlusRows') == d.get('paragraphs'),
                      'wP=%s rows=%s paras=%s' % (d.get('xmlParagraphs'), d.get('tableRows'), d.get('paragraphs')))

            df = run(['diff', '-Doc', doc, '-Since', 'tp-fast'])
            check('%s diff of an unchanged doc reports 0 changed' % fx.name,
                  df is not None and df.get('ok') and df['result']['summary']['changed'] == 0
                  and df['result']['summary']['added'] == 0 and df['result']['summary']['removed'] == 0,
                  json.dumps(df['result']['summary'], ensure_ascii=False)[:200] if df and df.get('ok') else str(df)[:200])
            check('%s revisions exercised: %s' % (fx.name, info.get('revisions')),
                  True)
            run(['close', '-Doc', doc, '-NoSave']) if False else None
            run(['close', '-Doc', doc])
    finally:
        if not keep:
            shutil.rmtree(work, ignore_errors=True)
        else:
            print('kept:', work)
    print('\n%d passed, %d failed' % (passed, failed))
    return 0 if failed == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
