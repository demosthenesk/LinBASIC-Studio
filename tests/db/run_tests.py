#!/usr/bin/env python3
# LinBASIC Studio - database tests (DAO, ADO, the Data control): each folder tests/db/<name>/ is a
# VB6 project whose program writes expected\<name>.txt; lbvc compiles it, runs it, and the output is
# compared with the one VB6 wrote (<name>/expected/). A test that reads BIBLIO.MDB or NWIND.MDB gets
# a copy of the user's own (the VB98 folder of VB6: $VB98_DIR); the
# Microsoft files are never part of LinBASIC Studio.
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#   python3 tests/db/run_tests.py [names...]
import glob, os, re, shutil, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
VB98 = os.environ.get('VB98_DIR', '')
sys.path.insert(0, os.path.join(HERE, '..', '..', 'compiler'))
from lbvc.driver import compile_project, build_exe, find_runtime   # noqa: E402


# in Wine (LBVC_TARGET win32/win64) these differ from VB6 because of Wine, not of LinBASIC: reported as KNOWN
WINE_KNOWN = {
    'd_odbc1': "Wine's odbc32 gives the SQLite ODBC driver of Linux, whose text columns are narrow (adVarChar 200 / "
               "adLongVarChar 201, SQL_VARCHAR 129); the driver of Windows gives them wide (202 / 203 / 130) as VB6 "
               "shows (on Linux LinBASIC makes them wide: odbc_wide)",
}


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


def rpt_normalize(data, drop=()):
    """the lines of a DataReport export ("  |..." in the output) as their visible text: VB6 lays the
    pages out with the metrics of the fonts of Windows (columns, rows, where the pages break), LinBASIC
    Studio with its own; what the report says and in what order must be the same"""
    out = []
    for ln in data.split(b'\r\n'):
        if ln.startswith(b'  bytes='):
            continue
        if ln.startswith(b'  |'):
            t = re.sub(rb'<cr>$', b'', ln[3:])
            t = re.sub(rb'<HR>', b' --- ', t)
            t = re.sub(rb'<[^>]*>', b' ', t)
            t = re.sub(rb'-{3,}', b'---', t)
            t = re.sub(rb'(Header|Page) \d+ of \d+', b'', t)
            t = b' '.join(t.split())
            if t and not any(d.fullmatch(t) for d in drop):
                out.append(b'  |' + t)
            continue
        out.append(ln)
    return b'\r\n'.join(out)


def sqlite_odbc():
    """the ADO tests over ODBC (d_odbc*) need unixODBC with the SQLite ODBC driver"""
    try:
        r = subprocess.run(['odbcinst', '-q', '-d'], capture_output=True, text=True, timeout=20)
    except OSError:
        return False
    names = [l.strip().strip('[]').lower() for l in r.stdout.splitlines()]
    if os.environ.get('LBVC_TARGET', 'linux') == 'win32':
        # 32-bit Wine: its odbc32 needs the 32-bit unixODBC and SQLite ODBC driver (i386 libsqliteodbc)
        if not glob.glob('/usr/lib/i386-linux-gnu/odbc/libsqlite3odbc*.so'):
            return False
    if os.environ.get('LBVC_TARGET', 'linux') != 'linux':
        # Wine: its odbc32 goes to the unixODBC of the computer with the name of the Windows driver,
        # so unixODBC needs that name too (~/.odbcinst.ini: [SQLite3 ODBC Driver] Driver=libsqlite3odbc.so)
        return 'sqlite3 odbc driver' in names
    return any('sqlite3' in n for n in names)


def main():
    names = sys.argv[1:] or sorted(os.path.basename(d) for d in glob.glob(os.path.join(HERE, '*')) if os.path.isdir(d)
                                   and glob.glob(os.path.join(d, '*.vbp')))
    rt = find_runtime()
    passed = failed = pending = known = 0
    for n in names:
        src = os.path.join(HERE, n)
        vbp = glob.glob(os.path.join(src, '*.vbp'))[0]
        work = os.path.join(HERE, 'work', n)
        shutil.rmtree(work, ignore_errors=True)
        os.makedirs(os.path.join(work, 'expected'))
        missing = False
        text = ''.join(open(p, 'rb').read().decode('cp1252', 'replace') for p in glob.glob(os.path.join(src, '*.bas')) + glob.glob(os.path.join(src, '*.frm')))
        for mdb in ('BIBLIO.MDB', 'NWIND.MDB'):
            if mdb.lower() in text.lower():
                p = os.path.join(VB98, mdb) if VB98 else ''
                if p and os.path.exists(p):
                    shutil.copy(p, work)
                else:
                    missing = True
        if ('sqlite3 odbc driver' in text.lower() or os.path.exists(os.path.join(src, 'odbc.ini'))) and not sqlite_odbc():
            print(f"---- {n}: needs unixODBC and the SQLite ODBC driver (sudo apt install unixodbc libsqliteodbc)"
                  + ("; in Wine registered as [SQLite3 ODBC Driver] too" if os.environ.get('LBVC_TARGET', 'linux') != 'linux' else "")
                  + ("; for win32 the 32-bit ones (libsqliteodbc:i386)" if os.environ.get('LBVC_TARGET') == 'win32' else ""))
            pending += 1
            continue
        if missing:
            print(f"---- {n}: needs BIBLIO.MDB / NWIND.MDB of VB6 (set VB98_DIR)")
            pending += 1
            continue
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
        if os.path.exists(os.path.join(work, 'odbc.ini')):   # the data sources of ODBC the test names (user DSNs)
            env['ODBCINI'] = os.path.join(work, 'odbc.ini')
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
        # App.Path in the output: the shared folder of the VM there, the work folder here
        exp = re.sub(rb'\\\\VBoxSvr\\VMShared\\[A-Za-z0-9_]+\\' + re.escape(n.encode()), b'<App.Path>', exp)
        got = got.replace(work.encode(), b'<App.Path>')
        got = got.replace(('Z:' + work.replace('/', '\\')).encode(), b'<App.Path>')   # in Wine
        if n.startswith('d_rpt') or n == 'd_de2':               # the files a DataReport exported: their text, not their layout
            drop = os.path.join(src, 'normalize.txt')   # what is where the pages break (page header / footer)
            drop = [re.compile(l.strip().encode()) for l in open(drop, encoding='utf-8') if l.strip() and not l.startswith('#')] \
                if os.path.exists(drop) else []
            exp, got = rpt_normalize(exp, drop), rpt_normalize(got, drop)
        if got == exp:
            print(f"ok   {n}")
            passed += 1
            continue
        if target != 'linux' and n in WINE_KNOWN:
            known += 1
            print(f"KNOWN {n} ({target}): {WINE_KNOWN[n]}")
            continue
        failed += 1
        print(f"FAIL {n}")
        el, gl = exp.decode('cp1253', 'replace').split('\r\n'), got.decode('cp1253', 'replace').split('\r\n')
        for i in range(max(len(el), len(gl))):
            a = el[i] if i < len(el) else '<none>'
            b = gl[i] if i < len(gl) else '<none>'
            if a != b:
                print(f"   line {i + 1}\n     VB6:  {a!r}\n     lbvc: {b!r}")
    print(f"passed {passed}, failed {failed}, without VB6 output {pending}" + (f", known (Wine) {known}" if known else ""))


if __name__ == '__main__':
    main()
