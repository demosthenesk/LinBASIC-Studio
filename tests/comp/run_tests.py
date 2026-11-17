#!/usr/bin/env python3
# LinBASIC Studio - component tests: each folder tests/comp/<name>/ is a VB6 project group
# (<name>/*.vbg) of ActiveX components and their client; the client's program writes
# expected\<name>.txt. lbvc compiles the client (with the components it uses, as in the IDE),
# runs it, and the output is compared with the one VB6 wrote (<name>/expected/).
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#   python3 tests/comp/run_tests.py [names...]
import glob, os, re, shutil, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', '..', 'compiler'))
sys.path.insert(0, os.path.join(HERE, '..', 'forms'))
from lbvc.driver import compile_project, build_exe, find_runtime, register_component   # noqa: E402
from run_tests import oracle_env                                  # noqa: E402  (tests/forms)


def group_vbps(vbg):
    """every project of a .vbg"""
    text = open(vbg, 'rb').read().decode('cp1252')
    return [os.path.join(os.path.dirname(vbg), v.strip().replace('\\', os.sep))
            for v in re.findall(r'(?im)^(?:Startup)?Project=(.+?)\s*$', text)]


def startup_vbp(vbg):
    """the startup project of a .vbg"""
    text = open(vbg, 'rb').read().decode('cp1252')
    m = re.search(r'(?im)^StartupProject=(.+?)\s*$', text)
    return os.path.join(os.path.dirname(vbg), m.group(1).strip().replace('\\', os.sep))


def main():
    names = sys.argv[1:] or sorted(os.path.basename(d) for d in glob.glob(os.path.join(HERE, '*'))
                                   if os.path.isdir(d) and glob.glob(os.path.join(d, '*.vbg')))
    rt = find_runtime()
    target = os.environ.get('LBVC_TARGET', 'linux')       # win32 / win64: the .exe runs in Wine
    passed = failed = pending = 0
    for n in names:
        src = os.path.join(HERE, n)
        work = os.path.join(HERE, 'work', n)
        shutil.rmtree(work, ignore_errors=True)
        shutil.copytree(src, work, ignore=shutil.ignore_patterns('expected'))
        os.makedirs(os.path.join(work, 'expected'))
        vbp = startup_vbp(glob.glob(os.path.join(work, '*.vbg'))[0])
        os.environ['LBVC_REGISTRY'] = os.path.join(work, 'components.json')   # not the user's
        for other in group_vbps(glob.glob(os.path.join(work, '*.vbg'))[0]):    # Make of the components
            if os.path.abspath(other) != os.path.abspath(vbp):
                register_component(other)
        try:
            csrc, _ = compile_project([vbp])
        except Exception as e:
            print(f"FAIL {n} (compile): {e}")
            failed += 1
            continue
        exe_dir = os.path.dirname(vbp)
        exe = build_exe(csrc, os.path.join(exe_dir, 'prog'), rt, target=target) or os.path.join(exe_dir, 'prog')
        env = oracle_env()
        cmd = [exe]
        if target != 'linux':
            cmd = ['xvfb-run', '-a', '-s', '-screen 0 1024x768x24', 'wine', exe]
            env.pop('QT_QPA_PLATFORM', None)
            env.setdefault('WINEDEBUG', '-all')
        r = subprocess.run(cmd, cwd=exe_dir, capture_output=True, text=True, timeout=180, env=env)
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
