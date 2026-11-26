#!/usr/bin/env python3
# LinBASIC Studio - form tests: each folder tests/forms/<name>/ is a VB6 project whose
# program writes expected\<name>.txt; lbvc compiles it, runs it without a screen (Qt
# "offscreen"), and the output is compared with the one VB6 wrote (<name>/expected/).
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#   python3 tests/forms/run_tests.py [names...]
import glob, os, shutil, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', '..', 'compiler'))
from lbvc.driver import compile_project, build_exe, find_runtime   # noqa: E402


def oracle_env():
    rt = os.path.join(HERE, 'work', '.xdg')
    os.makedirs(rt, mode=0o700, exist_ok=True)
    os.chmod(rt, 0o700)
    env = dict(os.environ, QT_QPA_PLATFORM='offscreen', XDG_RUNTIME_DIR=rt, VBRT_NO_ERRBOX='1')
    p = os.path.join(HERE, '..', 'lang', 'expected', 'LOCALE.txt')
    if os.path.exists(p):
        lcid, cp = open(p).read().split()[:2]
        env['VBRT_LCID'], env['VBRT_CODEPAGE'] = lcid, cp
    return env


def main():
    names = sys.argv[1:] or sorted(os.path.basename(d) for d in glob.glob(os.path.join(HERE, '*')) if os.path.isdir(d)
                                   and glob.glob(os.path.join(d, '*.vbp')))
    rt = find_runtime()
    passed = failed = pending = 0
    for n in names:
        src = os.path.join(HERE, n)
        vbp = glob.glob(os.path.join(src, '*.vbp'))[0]
        work = os.path.join(HERE, 'work', n)
        shutil.rmtree(work, ignore_errors=True)
        os.makedirs(os.path.join(work, 'expected'))
        for f in os.listdir(src):               # data files of the test (pictures ...)
            fp = os.path.join(src, f)
            if os.path.isfile(fp) and os.path.splitext(f)[1].lower() not in ('.frm', '.frx', '.bas', '.cls', '.vbp', '.txt'):
                shutil.copy(fp, work)
        try:
            csrc, _ = compile_project([vbp])
        except Exception as e:
            print(f"FAIL {n} (compile): {e}")
            failed += 1
            continue
        exe = os.path.join(work, 'prog')
        target = os.environ.get('LBVC_TARGET', 'linux')   # win32 / win64: the .exe runs in Wine
        exe = build_exe(csrc, exe, rt, target=target) or exe
        env = oracle_env()
        help_log = os.path.join(work, 'help.txt')      # what the help of the program opens (VBRT_HELP_LOG)
        env['VBRT_HELP_LOG'] = help_log if target == 'linux' else 'help.txt'
        cmd = [exe]
        if target != 'linux':                             # Wine on a virtual X screen (Xvfb)
            cmd = ['xvfb-run', '-a', '-s', '-screen 0 1024x768x24', 'wine', exe]
            env.pop('QT_QPA_PLATFORM', None)
            env.setdefault('WINEDEBUG', '-all')
        r = subprocess.run(cmd, cwd=work, capture_output=True, text=True, timeout=180, env=env)
        out = os.path.join(work, 'expected', n + '.txt')
        got = open(out, 'rb').read() if os.path.exists(out) else b''
        if r.returncode != 0:
            got += f"[exit {r.returncode}] {r.stderr}".encode()
        exp_path = os.path.join(src, 'expected', n + '.txt')
        if not os.path.exists(exp_path):
            print(f"---- {n}: no VB6 output yet; lbvc wrote work/{n}/expected/{n}.txt")
            pending += 1
            continue
        exp = open(exp_path, 'rb').read()
        hexp_path = os.path.join(src, 'expected', n + '-help.txt')   # LinBASIC's help: no VB6 output of it
        if got == exp and os.path.exists(hexp_path):
            hgot = open(help_log, 'rb').read().replace(b'\r\n', b'\n') if os.path.exists(help_log) else b''
            hexp = open(hexp_path, 'rb').read().replace(b'\r\n', b'\n')
            if hgot != hexp:
                failed += 1
                print(f"FAIL {n} (help)")
                hl, gl = hexp.decode('utf-8').split('\n'), hgot.decode('utf-8', 'replace').split('\n')
                for i in range(max(len(hl), len(gl))):
                    a = hl[i] if i < len(hl) else '<none>'
                    b = gl[i] if i < len(gl) else '<none>'
                    if a != b:
                        print(f"   help line {i + 1}\n     expected: {a!r}\n     lbvc:     {b!r}")
                continue
        if got == exp:
            print(f"ok   {n}")
            passed += 1
            continue
        failed += 1
        print(f"FAIL {n}")
        el, gl = exp.decode('cp1252').split('\r\n'), got.decode('cp1252', 'replace').split('\r\n')
        for i in range(max(len(el), len(gl))):
            a = el[i] if i < len(el) else '<none>'
            b = gl[i] if i < len(gl) else '<none>'
            if a != b:
                print(f"   line {i + 1}\n     VB6:  {a!r}\n     lbvc: {b!r}")
    print(f"passed {passed}, failed {failed}, without VB6 output {pending}")


if __name__ == '__main__':
    main()
