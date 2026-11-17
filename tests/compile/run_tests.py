#!/usr/bin/env python3
"""What VB6 compiles and what it rejects (phase 6.8): each folder cN_name is a small project; VB6's Make EXE of
each, and the output of the made program, are in expected/results.txt (make.bat.txt, renamed make.bat, makes it
in the Win7 VM). lbvc builds each project as Make does (no --debug: no Compile On Demand) and must give the same
compile error (its message) or, when VB6 made the program, the same out.txt.

    python3 tests/compile/run_tests.py [names...]
"""
import os, re, shutil, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
LBVC = os.path.join(HERE, '..', '..', 'compiler', 'lbvc.py')
WORK = os.path.join(HERE, 'work')
os.environ['VBRT_LCID'] = '1032'                 # the Greek Windows 7 of the VM (c14: "3.5" -> 35)


def vb6_results():
    vb, cur = {}, None
    for line in open(os.path.join(HERE, 'expected', 'results.txt'), encoding='cp1252').read().replace('\r', '').split('\n'):
        m = re.match(r'==== (\S+)', line)
        if m:
            cur = m.group(1); vb[cur] = []
        elif cur and line.strip():
            vb[cur].append(line.strip())
    return vb


def summary(lines):
    if any(l.startswith('Build of') and l.endswith('failed.') for l in lines):
        e = [l for l in lines if 'Compile Error' in l]
        return ['COMPILE ERROR: ' + (e[0].split(' : ', 1)[1] if e else '?')]
    return [l for l in lines if not l.startswith('Build of')]


def lbvc_result(n):
    d = os.path.join(WORK, n)
    shutil.rmtree(d, ignore_errors=True)
    shutil.copytree(os.path.join(HERE, n), d)
    r = subprocess.run([sys.executable, LBVC, os.path.join(d, n + '.vbp'), '-o', os.path.join(d, 'prog')],
                       capture_output=True, text=True)
    if r.returncode != 0:
        e = [l for l in (r.stderr + r.stdout).splitlines() if ': error: ' in l]
        msg = e[0].split(': error: ', 1)[1] if e else (r.stderr + r.stdout).strip()
        if msg.count(': ') and not msg.rstrip().endswith("'"):
            msg = msg.rsplit(': ', 1)[0]          # lbvc adds the name: "variable not defined: x"
        return ['COMPILE ERROR: ' + msg]
    subprocess.run([os.path.join(d, 'prog')], cwd=d, capture_output=True, timeout=60,
                   env=dict(os.environ, VBRT_NO_ERRBOX='1'))
    p = os.path.join(d, 'out.txt')
    if not os.path.exists(p):
        return []
    return [l.strip() for l in open(p, encoding='cp1252').read().replace('\r', '').split('\n') if l.strip()]


def main():
    vb = vb6_results()
    names = sorted((d for d in os.listdir(HERE) if re.match(r'c\d+_', d)), key=lambda d: int(d[1:].split('_')[0]))
    if sys.argv[1:]:
        names = [n for n in names if any(a in n for a in sys.argv[1:])]
    ok = bad = 0
    for n in names:
        want = summary(vb.get(n, []))
        got = lbvc_result(n)
        if [g.lower() for g in got] == [w.lower() for w in want]:
            ok += 1
            print('ok  ', n)
        else:
            bad += 1
            print('FAIL', n, '\n   VB6 :', want, '\n   lbvc:', got)
    print(f'{ok} passed, {bad} failed')
    return 1 if bad else 0


sys.exit(main())
