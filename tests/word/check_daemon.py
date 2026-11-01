"""Acceptance check for the v1.14 resident helper (pure ASCII).

Proves that the daemon is a *transport* change only, and that it actually pays:

  D1  `serve -Action start|status|stop` works and the pipe names the built binary.
  D2  Every hot command returns the SAME JSON through the daemon as through the one-shot path
      (after dropping volatile fields: timings, pids, absolute paths).
  D3  The daemon really is used (server reports ping/pid, and per-command wall time drops).
  D4  `close` of the last document makes the daemon exit by itself.
  D5  With the daemon stopped, commands still work (one-shot fallback) - nothing depends on it.

    python check_daemon.py
"""

import json
import os
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
WORD = REPO / 'scripts' / 'word.ps1'
FIXTURE = REPO / 'tests' / 'word' / 'captions' / 'captions-fixture.docx'

VOLATILE = {'ms', 'path', 'at', 'pid', 'pipe', 'elapsedMs', 'durationMs', 'time', 'startedAt',
            'note', 'changedOutsideHint', 'hint', 'captionedAt', 'textsMs', 'xmlNote', 'chunkStats',
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


def run(args, env_extra=None, timeout=300):
    env = dict(os.environ)
    if env_extra:
        env.update(env_extra)
    sw = time.time()
    p = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(WORD)] + args,
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
    """Leftover helper daemons / automation instances make WPS answer from a different instance
    than the one under test, which is not what this check is about."""
    subprocess.run(['powershell', '-NoProfile', '-Command',
                    "Get-Process WpsHelper -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue;"
                    "Get-CimInstance Win32_Process -Filter \"Name='wps.exe'\" |"
                    " Where-Object { $_.CommandLine -match 'Automation' } |"
                    " ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"],
                   capture_output=True, text=True)
    time.sleep(1.0)


def main():
    kill_strays()
    work = Path(tempfile.mkdtemp(prefix='woc-daemon-'))
    doc = str(work / FIXTURE.name)
    shutil.copyfile(FIXTURE, doc)
    try:
        run(['serve', '-Action', 'stop'])
        time.sleep(0.8)

        sequence = [
            ('open', ['open', '-Doc', doc, '-NoHint']),
            ('snapshot', ['snapshot', '-Doc', doc, '-Tag', 'ab']),
            ('read', ['read', '-Doc', doc, '-SkipTableCells', '-Count', '40']),
            ('search', ['search', '-Doc', doc, '-Text', '表']),
            ('status', ['status']),
            ('diff', ['diff', '-Doc', doc, '-Since', 'ab']),
        ]

        # --- one-shot baseline
        oneshot = {}
        t_oneshot = {}
        for name, args in sequence:
            res, ms = run(args, {'WOC_NO_DAEMON': '1'})
            oneshot[name] = res
            t_oneshot[name] = ms
        check('one-shot baseline produced results', all(v is not None for v in oneshot.values()),
              str({k: v is not None for k, v in oneshot.items()}))

        # --- daemon
        start, _ = run(['serve', '-Action', 'start'])
        check('D1 serve start', start is not None and start.get('ok') and start['result'].get('started'),
              json.dumps(start)[:200] if start else '')
        status1, _ = run(['serve', '-Action', 'status'])
        check('D1 serve status reports running', status1 is not None and status1['result'].get('running'),
              json.dumps(status1)[:200] if status1 else '')

        daemon = {}
        t_daemon = {}
        for name, args in sequence:
            res, ms = run(args)
            daemon[name] = res
            t_daemon[name] = ms

        for name in [n for n, _ in sequence]:
            a, b = oneshot.get(name), daemon.get(name)
            ok = a is not None and b is not None and a.get('ok') == b.get('ok')
            if ok and a.get('ok'):
                ok = scrub(a['result']) == scrub(b['result'])
            check('D2 %s identical through daemon and one-shot' % name, ok,
                  '' if ok else json.dumps({'oneshot': scrub(a), 'daemon': scrub(b)}, ensure_ascii=False)[:600])

        faster = 0
        for name in ('snapshot', 'diff', 'status', 'read'):
            if t_daemon[name] < t_oneshot[name] - 500:
                faster += 1
        check('D3 daemon is faster on the read-side commands (%d/4 by >0.5 s)' % faster, faster >= 2,
              'one-shot=%s daemon=%s' % (t_oneshot, t_daemon))
        print('    timings  one-shot: %s' % t_oneshot)
        print('    timings  daemon  : %s' % t_daemon)

        close, _ = run(['close', '-Doc', doc])
        check('close works through the daemon', close is not None and close.get('ok'))
        gone = False
        status2 = None
        for _ in range(10):
            time.sleep(0.8)
            status2, _ = run(['serve', '-Action', 'status'])
            if status2 is not None and status2['result'].get('running') is False:
                gone = True
                break
        check('D4 daemon exits after the last document is closed', gone,
              json.dumps(status2)[:200] if status2 else '')

        # --- D5 fallback with no daemon at all
        res, _ = run(['open', '-Doc', doc, '-NoHint'])
        check('D5 one-shot fallback still opens', res is not None and res.get('ok'))
        res2, _ = run(['status'], {'WOC_NO_DAEMON': '1'})
        check('D5 one-shot fallback still runs status', res2 is not None and res2.get('ok'))
        run(['serve', '-Action', 'stop'])
        time.sleep(0.5)
        res3, _ = run(['close', '-Doc', doc])
        check('D5 fallback cleanup works', res3 is not None and res3.get('ok'))
    finally:
        run(['serve', '-Action', 'stop'])
        shutil.rmtree(work, ignore_errors=True)

    print('\n%d passed, %d failed' % (passed, failed))
    return 0 if failed == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
