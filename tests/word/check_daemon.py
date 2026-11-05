"""Acceptance check for the resident helper (v1.14 word, v1.15 excel/ppt). Pure ASCII.

Proves that the daemon is a *transport* change only, and that it pays - for every app:

  D1  `serve -Action start|status|stop` works and the pipe is named after the built binary.
  D2  Every hot command returns the SAME JSON through the daemon as through the one-shot path
      (after dropping volatile fields: timings, pids, absolute paths).
  D3  The daemon really is used (per-command wall time drops).
  D4  `close` of the last document makes the daemon exit by itself.
  D5  With the daemon stopped, commands still work (one-shot fallback) - nothing depends on it.

    python check_daemon.py            # all three apps
    python check_daemon.py word       # one app
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SCRIPTS = REPO / 'scripts'

VOLATILE = {'ms', 'path', 'at', 'pid', 'pipe', 'elapsedMs', 'durationMs', 'time', 'startedAt',
            'note', 'changedOutsideHint', 'hint', 'captionedAt', 'textsMs', 'xmlNote', 'chunkStats',
            'taken',   # excel/ppt snapshot timestamp
            # "how": opened vs reused depends on whether an earlier phase left the doc open;
            # multiInstanceWarning depends on how many WPS automation instances happen to run.
            'how', 'multiInstanceWarning'}

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


def make_fixtures(dirp):
    """Create xlsx/pptx fixtures with openpyxl / python-pptx (both are in the bundled deps)."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'summary'
    ws['A1'] = 'item'
    ws['B1'] = 'value'
    for i in range(2, 42):
        ws.cell(row=i, column=1, value='row%d' % i)
        ws.cell(row=i, column=2, value=i * 1.5)
    ws['D1'] = 'total'
    ws['D2'] = '=SUM(B2:B41)'
    ws2 = wb.create_sheet('detail')
    for r in range(1, 21):
        for c in range(1, 6):
            ws2.cell(row=r, column=c, value=r * c)
    wb.create_sheet('empty')
    xlsx = dirp / 'probe.xlsx'
    wb.save(xlsx)

    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation()
    for i in range(1, 7):
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = 'chapter %d' % i
        s.placeholders[1].text = 'body of slide %d used by the resident-transport check' % i
    t = prs.slides.add_slide(prs.slide_layouts[5])
    t.shapes.title.text = 'table slide'
    tbl = t.shapes.add_table(4, 3, Inches(1), Inches(1.6), Inches(6), Inches(2)).table
    for r in range(4):
        for c in range(3):
            tbl.cell(r, c).text = 'r%dc%d' % (r, c)
    pptx = dirp / 'probe.pptx'
    prs.save(pptx)
    return xlsx, pptx


def run(ps, args, env_extra=None, timeout=300):
    env = dict(os.environ)
    if env_extra:
        env.update(env_extra)
    sw = time.time()
    p = subprocess.run(['powershell', '-Mta', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(ps)] + args,
                       capture_output=True, text=True, encoding='utf-8', errors='replace', env=env, timeout=timeout)
    ms = int((time.time() - sw) * 1000)
    try:
        return json.loads(p.stdout), ms
    except Exception:
        print('  BAD JSON (%.0f ms): %s' % (ms, (p.stdout or p.stderr)[:300]))
        return None, ms


def scrub(v):
    if isinstance(v, dict):
        return {k: scrub(x) for k, x in v.items() if k not in VOLATILE}
    if isinstance(v, list):
        return [scrub(x) for x in v]
    if isinstance(v, str) and (':\\' in v or '\\\\' in v):
        return '<path>'
    return v


def kill_strays():
    subprocess.run(['powershell', '-NoProfile', '-Command',
                    "Get-Process WpsHelper,EtHelper,WppHelper -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue;"
                    "Get-CimInstance Win32_Process -Filter \"Name='wps.exe' or Name='et.exe' or Name='wpp.exe'\" |"
                    " Where-Object { $_.CommandLine -match 'Automation' } |"
                    " ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"],
                   capture_output=True, text=True)
    time.sleep(1.0)


def run_app(app, ps, doc, sequence):
    print('\n=== %s ===' % app)
    run(ps, ['serve', '-Action', 'stop'])
    time.sleep(0.6)

    one, t_one = {}, {}
    for name, args in sequence:
        res, ms = run(ps, args, {'WOC_NO_DAEMON': '1'})
        one[name], t_one[name] = res, ms
    check('%s one-shot baseline produced results' % app, all(v is not None for v in one.values()),
          str({k: v is not None for k, v in one.items()}))

    start, _ = run(ps, ['serve', '-Action', 'start'])
    check('%s D1 serve start' % app, start is not None and start.get('ok') and start['result'].get('started'),
          json.dumps(start)[:200] if start else '')
    st, _ = run(ps, ['serve', '-Action', 'status'])
    check('%s D1 serve status running' % app, st is not None and st['result'].get('running'),
          json.dumps(st)[:200] if st else '')

    daemon, t_daemon = {}, {}
    for name, args in sequence:
        res, ms = run(ps, args)
        daemon[name], t_daemon[name] = res, ms

    for name, _ in sequence:
        a, b = one.get(name), daemon.get(name)
        ok = a is not None and b is not None and a.get('ok') == b.get('ok')
        if ok and a.get('ok'):
            ok = scrub(a['result']) == scrub(b['result'])
        check('%s D2 %s identical (daemon vs one-shot)' % (app, name), ok,
              '' if ok else json.dumps({'oneshot': scrub(a), 'daemon': scrub(b)}, ensure_ascii=False)[:500])

    faster = sum(1 for n in t_daemon if n != 'open' and t_daemon[n] < t_one[n] - 500)
    total = len([n for n in t_daemon if n != 'open'])
    check('%s D3 daemon faster on %d/%d non-open commands by >0.5 s' % (app, faster, total),
          faster >= max(1, total - 1), 'one-shot=%s daemon=%s' % (t_one, t_daemon))
    print('    timings one-shot: %s' % t_one)
    print('    timings daemon  : %s' % t_daemon)

    close, _ = run(ps, ['close', '-Doc', doc])
    check('%s close works through the daemon' % app, close is not None and close.get('ok'))
    cres = (close or {}).get('result') or {}
    left = 0
    for k, v in cres.items():
        if str(k).endswith('Left') and isinstance(v, int):
            left = v
    if left:
        # Environment, not a regression: the user keeps their own document open in the same WPS
        # instance, so the test never closes "the last document" and the daemon correctly stays.
        print('  SKIP  %s D4 daemon exit: %d other document(s) still open (user has one open)' % (app, left))
    else:
        gone, st2 = False, None
        for _ in range(10):
            time.sleep(0.8)
            st2, _ = run(ps, ['serve', '-Action', 'status'])
            if st2 is not None and st2['result'].get('running') is False:
                gone = True
                break
        check('%s D4 daemon exits after the last document is closed' % app, gone,
              json.dumps(st2)[:200] if st2 else '')

    res, _ = run(ps, ['open', '-Doc', doc, '-NoHint'])
    check('%s D5 fallback open works' % app, res is not None and res.get('ok'))
    res2, _ = run(ps, ['status'], {'WOC_NO_DAEMON': '1'})
    check('%s D5 fallback status works' % app, res2 is not None and res2.get('ok'))
    run(ps, ['close', '-Doc', doc])


def main():
    only = [a for a in sys.argv[1:] if not a.startswith('-')]
    kill_strays()
    work = Path(tempfile.mkdtemp(prefix='woc-daemon-'))
    try:
        xlsx, pptx = make_fixtures(work)
        word_doc = work / 'captions-fixture.docx'
        shutil.copyfile(REPO / 'tests' / 'word' / 'captions' / 'captions-fixture.docx', word_doc)

        apps = [
            ('word', SCRIPTS / 'word.ps1', str(word_doc), [
                ('open', ['open', '-Doc', str(word_doc), '-NoHint']),
                ('snapshot', ['snapshot', '-Doc', str(word_doc), '-Tag', 'ab']),
                ('read', ['read', '-Doc', str(word_doc), '-SkipTableCells', '-Count', '40']),
                ('search', ['search', '-Doc', str(word_doc), '-Text', '表']),
                ('status', ['status']),
                ('diff', ['diff', '-Doc', str(word_doc), '-Since', 'ab']),
            ]),
            ('excel', SCRIPTS / 'excel.ps1', str(xlsx), [
                ('open', ['open', '-Doc', str(xlsx)]),
                ('read', ['read', '-Doc', str(xlsx)]),
                ('snapshot', ['snapshot', '-Doc', str(xlsx)]),
                ('status', ['status']),
                ('diff', ['diff', '-Doc', str(xlsx)]),
            ]),
            ('ppt', SCRIPTS / 'ppt.ps1', str(pptx), [
                ('open', ['open', '-Doc', str(pptx)]),
                ('read', ['read', '-Doc', str(pptx)]),
                ('snapshot', ['snapshot', '-Doc', str(pptx)]),
                ('status', ['status']),
                ('diff', ['diff', '-Doc', str(pptx)]),
            ]),
        ]
        for app, ps, doc, seq in apps:
            if only and app not in only:
                continue
            run_app(app, ps, doc, seq)
    finally:
        for ps in (SCRIPTS / 'word.ps1', SCRIPTS / 'excel.ps1', SCRIPTS / 'ppt.ps1'):
            run(ps, ['serve', '-Action', 'stop'])
        shutil.rmtree(work, ignore_errors=True)

    print('\n%d passed, %d failed' % (passed, failed))
    return 0 if failed == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
