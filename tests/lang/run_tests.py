#!/usr/bin/env python3
# LinBASIC Studio - language tests: compile each tests/lang/t_*.bas with lbvc, run it,
# and compare its output with the output of the same test compiled by VB6 (expected/).
# LBVC_TARGET=win64 (or win32): the tests as Windows .exe files, run in Wine.
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#   python3 tests/lang/run_tests.py [names...]
import json, os, shutil, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', '..', 'compiler'))
from lbvc.driver import compile_project, build_exe, find_runtime   # noqa: E402

HARNESS = '''Attribute VB_Name = "HARNESS"
Option Explicit
Sub Main()
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\\out.txt" For Output As #f
    On Error GoTo bad
    {sub} f
    Close #f
    Exit Sub
bad:
    Print #f, "UNHANDLED"; Err.Number; Err.Description
    Close #f
End Sub
'''


def oracle_env():
    """The regional settings of the Windows that produced expected/ (expected/LOCALE.txt:
    "<LCID> <ANSI code page>"); the lbvc programs run with the same ones."""
    env = dict(os.environ, VBRT_NO_ERRBOX='1')
    p = os.path.join(HERE, 'expected', 'LOCALE.txt')
    if os.path.exists(p):
        lcid, cp = open(p).read().split()[:2]
        env['VBRT_LCID'], env['VBRT_CODEPAGE'] = lcid, cp
    return env


def run(name, spec, rt, work):
    sub, extra = (spec, []) if isinstance(spec, str) else (spec['sub'], spec.get('files', []))
    d = os.path.join(work, name)
    shutil.rmtree(d, ignore_errors=True)
    os.makedirs(d)
    h = os.path.join(d, 'harness.bas')
    with open(h, 'wb') as f:
        f.write(HARNESS.format(sub=sub).replace('\n', '\r\n').encode('cp1252'))
    try:
        csrc, _ = compile_project([os.path.join(HERE, name + '.bas')] + [os.path.join(HERE, x) for x in extra] + [h])
    except Exception as e:
        return 'compile', str(e)
    exe = os.path.join(d, 'prog')
    target = os.environ.get('LBVC_TARGET', 'linux')       # win32 / win64: the .exe runs in Wine
    exe = build_exe(csrc, exe, rt, target=target) or exe
    cmd = [exe] if target == 'linux' else ['wine', exe]
    env = oracle_env()
    if target != 'linux':
        env.setdefault('WINEDEBUG', '-all')
    r = subprocess.run(cmd, cwd=d, capture_output=True, text=True, timeout=120, env=env)
    out = os.path.join(d, 'out.txt')
    got = open(out, 'rb').read() if os.path.exists(out) else b''
    if r.returncode != 0:
        got += f"[exit {r.returncode}] {r.stderr}".encode()
    return 'ran', got


# differences of Wine and of 64-bit Windows that are not LinBASIC's (LBVC_TARGET=win32 / win64): reported, not
# counted as failures
WINE_KNOWN = {
    't_declare': "Wine: Sleep 50 between two GetTickCount can measure less than 40 ms",
    't_declare2': "Wine: lstrlenA(vbNullString) sets LastDllError 87 (Windows: unchanged, 71); win64: a BSTR "
                  "pointer is 8 bytes, so CopyMemory p, b(8), 4 of a VARIANT reads half of it (VB6 is 32-bit)",
}


def main():
    tests = json.load(open(os.path.join(HERE, 'tests.json')))
    target = os.environ.get('LBVC_TARGET', 'linux')
    known = 0
    names = sys.argv[1:] or list(tests)
    rt = find_runtime()
    work = os.path.join(HERE, 'work')
    passed = failed = pending = 0
    for n in names:
        n = n.replace('.bas', '')
        kind, got = run(n, tests[n], rt, work)
        if kind == 'compile':
            print(f"FAIL {n} (compile): {got}")
            failed += 1
            continue
        exp_path = os.path.join(HERE, 'expected', n + '.txt')
        if not os.path.exists(exp_path):
            print(f"---- {n}: no VB6 output yet; lbvc wrote work/{n}/out.txt")
            pending += 1
            continue
        exp = open(exp_path, 'rb').read()
        if got == exp:
            print(f"ok   {n}")
            passed += 1
        elif target != 'linux' and n in WINE_KNOWN:
            known += 1
            print(f"KNOWN {n} ({target}): {WINE_KNOWN[n]}")
        else:
            failed += 1
            print(f"FAIL {n}")
            el, gl = exp.decode('cp1252').split('\r\n'), got.decode('cp1252', 'replace').split('\r\n')
            for i in range(max(len(el), len(gl))):
                a = el[i] if i < len(el) else '<none>'
                b = gl[i] if i < len(gl) else '<none>'
                if a != b:
                    print(f"   line {i + 1}\n     VB6:  {a!r}\n     lbvc: {b!r}")
    print(f"passed {passed}, failed {failed}, without VB6 output {pending}" + (f", known differences {known}" if known else ""))
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
