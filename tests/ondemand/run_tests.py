#!/usr/bin/env python3
"""Compile On Demand, as a run of the VB6 IDE (F5): each folder oN_name is a small project whose program writes
expected/<name>.txt; lbvc builds it with --ide (a procedure with a compile error becomes a stub that reports it when
it is called) and its output must be the one of VB6's F5 run (<name>/expected/<name>.txt, from the VM).

    python3 tests/ondemand/run_tests.py [names...]
"""
import glob, os, re, shutil, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
LBVC = os.path.join(HERE, '..', '..', 'compiler', 'lbvc.py')
WORK = os.path.join(HERE, 'work')


def main():
    names = sorted((d for d in os.listdir(HERE) if re.match(r'o\d+_', d)), key=lambda d: int(d[1:].split('_')[0]))
    if sys.argv[1:]:
        names = [n for n in names if any(a in n for a in sys.argv[1:])]
    ok = bad = pending = 0
    for n in names:
        d = os.path.join(WORK, n)
        shutil.rmtree(d, ignore_errors=True)
        shutil.copytree(os.path.join(HERE, n), d, ignore=shutil.ignore_patterns('expected'))
        vbp = glob.glob(os.path.join(d, '*.vbp'))[0]
        r = subprocess.run([sys.executable, LBVC, vbp, '--ide', '-o', os.path.join(d, 'prog')],
                           capture_output=True, text=True)
        if r.returncode != 0:
            bad += 1
            print('FAIL', n, '(compile):', (r.stderr + r.stdout).strip().splitlines()[-1:])
            continue
        p = subprocess.run([os.path.join(d, 'prog')], cwd=d, capture_output=True, text=True, timeout=60,
                           env=dict(os.environ, VBRT_NO_ERRBOX='1', QT_QPA_PLATFORM='offscreen'))
        out = os.path.join(d, 'expected', n + '.txt')
        got = open(out, 'rb').read() if os.path.exists(out) else b''
        # <name>.err: the message of the "Compile error:" box that VB6 showed (it ends the run, writes nothing)
        errf = os.path.join(HERE, n, 'expected', n + '.err')
        msg = open(errf, encoding='utf-8').read().strip() if os.path.exists(errf) else None
        if p.returncode and not (msg and p.returncode == 1 and 'Compile error:' in p.stderr
                                 and msg.lower() in p.stderr.lower()):
            got += f"[exit {p.returncode}] {p.stderr}".encode()
        elif msg and not p.returncode:
            got += b"[no compile error: " + msg.encode() + b"]"
        exp = os.path.join(HERE, n, 'expected', n + '.txt')
        if not os.path.exists(exp):
            pending += 1
            print(f"---- {n}: no VB6 output yet; lbvc wrote work/{n}/expected/{n}.txt")
            continue
        want = open(exp, 'rb').read()
        if got == want:
            ok += 1
            print('ok  ', n)
        else:
            bad += 1
            print('FAIL', n, '\n   VB6 :', want, '\n   lbvc:', got)
    print(f'passed {ok}, failed {bad}, without VB6 output {pending}')
    return 1 if bad else 0


sys.exit(main())
