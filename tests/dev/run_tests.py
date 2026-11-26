#!/usr/bin/env python3
# LinBASIC Studio - device tests (MCI, the Multimedia control, MSComm): each folder tests/dev/<name>/
# is a VB6 project whose program writes expected\<name>.txt; lbvc compiles it, runs it, and the output
# is compared with the one VB6 wrote (<name>/expected/). The sound goes to a GStreamer fakesink
# (VBRT_MCI_SINK); the serial ports COM1 and COM2 are a pair of pseudo terminals made by socat; the Inet
# tests (f_inet*) talk to inet_server.py (HTTP 47120, FTP 47121) on this computer; the MAPI tests (f_mapi*) have
# LinBASIC's provider of Simple MAPI with a store of their own (work/.mapi/<name>: LBMAPI_DIR), f_mapi1 none at all
# (LBMAPI_NONE: VB6 without a mail program).
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#   python3 tests/dev/run_tests.py [names...]
import glob, os, shutil, subprocess, sys, time
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', '..', 'compiler'))
from lbvc.driver import compile_project, build_exe, find_runtime   # noqa: E402


# in Wine (LBVC_TARGET win32/win64) these differ from VB6 because of Wine, not of LinBASIC: reported as KNOWN
WINE_KNOWN = {
    'd_mci': "Wine's MCI (winmm) is not the one of Windows: error numbers and texts, the sequencer, the AVI window",
    'd_mci2': "Wine's MCI (winmm): error texts and numbers, the notify codes",
    'f_mci': "Wine's MCI (winmm): error texts and numbers, the notify codes, the device of a sound file",
    'f_comm': "Wine: COM1/COM2 are the serial ports of the Wine prefix, not the socat pair of the test",
}


def wine_trust(on, keys=None):
    """Wine: WinINet trusts only the store of Windows, so the certificates of ca.pem (made by inet_server.py) are
    put in the Root store of the Wine prefix (HKCU and HKLM ...\\SystemCertificates\\Root\\Certificates\\<SHA-1>,
    the value Blob: the serialized certificate) while the tests run, and taken out after them"""
    import base64, hashlib, struct
    env = dict(os.environ, WINEDEBUG='-all')
    roots = ('HKEY_CURRENT_USER', 'HKEY_LOCAL_MACHINE')
    if not on:
        for k in keys or []:
            for r in roots:
                subprocess.run(['wine', 'reg', 'delete', f'{r}\\Software\\Microsoft\\SystemCertificates\\Root\\Certificates\\{k}',
                                '/f'], env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return []
    tls = os.path.join(HERE, 'work', '.tls')
    subprocess.run([sys.executable, os.path.join(HERE, 'inet_server.py'), '--tls-dir', tls, '--certs-only'],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    ca = os.path.join(tls, 'ca.pem')
    if not os.path.exists(ca):
        return []
    text, keys, lines = open(ca).read(), [], ['REGEDIT4', '']
    for part in text.split('-----BEGIN CERTIFICATE-----')[1:]:
        der = base64.b64decode(part.split('-----END CERTIFICATE-----')[0])
        k = hashlib.sha1(der).hexdigest().upper()
        keys.append(k)
        blob = struct.pack('<III', 32, 1, len(der)) + der        # CERT_CERT_PROP_ID, X509_ASN_ENCODING, size
        hx = ','.join(f'{b:02x}' for b in blob)
        rows = [hx[i:i + 75] for i in range(0, len(hx), 75)]
        value = '\\\n  '.join(rows)
        for r in roots:
            lines += [f'[{r}\\Software\\Microsoft\\SystemCertificates\\Root\\Certificates\\{k}]',
                      f'"Blob"=hex:{value}', '']
    reg = os.path.join(tls, 'wine-root.reg')
    with open(reg, 'w', newline='\r\n') as f:
        f.write('\n'.join(lines) + '\n')
    subprocess.run(['wine', 'reg', 'import', 'Z:' + reg.replace('/', '\\')], env=env,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return keys


def oracle_env():
    rt = os.path.join(HERE, 'work', '.xdg')
    os.makedirs(rt, mode=0o700, exist_ok=True)
    os.chmod(rt, 0o700)
    env = dict(os.environ, QT_QPA_PLATFORM='offscreen', XDG_RUNTIME_DIR=rt, VBRT_NO_ERRBOX='1', VBRT_NO_BROWSER='1')
    env.setdefault('VBRT_MCI_SINK', 'fakesink')
    p = os.path.join(HERE, '..', 'lang', 'expected', 'LOCALE.txt')
    if os.path.exists(p):
        lcid, cp = open(p).read().split()[:2]
        env['VBRT_LCID'], env['VBRT_CODEPAGE'] = lcid, cp
    return env


def sqlite_odbc(target):
    """the RDO tests (f_rdo*) need unixODBC with the SQLite ODBC driver (in Wine registered as
    [SQLite3 ODBC Driver] too: ~/.odbcinst.ini), as tests/db/d_odbc*"""
    try:
        r = subprocess.run(['odbcinst', '-q', '-d'], capture_output=True, text=True, timeout=20)
    except OSError:
        return False
    names = [l.strip().strip('[]').lower() for l in r.stdout.splitlines()]
    if target == 'win32' and not glob.glob('/usr/lib/i386-linux-gnu/odbc/libsqlite3odbc*.so'):
        return False                            # 32-bit Wine: a 32-bit driver (none in Ubuntu 24)
    if target != 'linux':
        return 'sqlite3 odbc driver' in names
    return any('sqlite3' in n for n in names)


def port_pair(work):
    """COM1 <-> COM2: two pseudo terminals joined by socat (None if socat is missing)"""
    if not shutil.which('socat'):
        return None, {}
    a, b = os.path.join(work, 'ttyCOM1'), os.path.join(work, 'ttyCOM2')
    p = subprocess.Popen(['socat', f'pty,raw,echo=0,link={a}', f'pty,raw,echo=0,link={b}'],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(50):
        if os.path.exists(a) and os.path.exists(b):
            break
        time.sleep(0.05)
    return p, {'VBRT_COM1': a, 'VBRT_COM2': b}


def inet_server(work):
    """the HTTP and FTP servers of the Inet tests; server.txt tells the program where they are"""
    with open(os.path.join(work, 'server.txt'), 'w') as f:
        f.write('127.0.0.1\n')
    p = subprocess.Popen([sys.executable, os.path.join(HERE, 'inet_server.py'), '--root', os.path.join(work, 'ftproot'),
                          '--log', os.path.join(work, 'inetserver.log'), '--tls-dir', os.path.join(HERE, 'work', '.tls')],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    import socket
    for port in (47121, 47125):                 # the last server it starts (47125 when openssl is there)
        for _ in range(100):
            try:
                socket.create_connection(('127.0.0.1', port), timeout=0.2).close()
                break
            except OSError:
                time.sleep(0.05)
    return p


def main():
    names = sys.argv[1:] or sorted(os.path.basename(d) for d in glob.glob(os.path.join(HERE, '*')) if os.path.isdir(d)
                                   and glob.glob(os.path.join(d, '*.vbp'))
                                   and not os.path.exists(os.path.join(d, 'run_check.py')))
    names = [n for n in names if n != 'sysinfo']
    rt = find_runtime()
    passed = failed = pending = known = 0
    target = os.environ.get('LBVC_TARGET', 'linux')
    trusted = wine_trust(True) if target != 'linux' and any(n.startswith('f_inet') for n in names) else []
    try:
        passed, failed, pending, known = run_all(names, rt, target)
    finally:
        if trusted:
            wine_trust(False, trusted)
    if not sys.argv[1:] or 'sysinfo' in sys.argv[1:]:   # SysInfo's events: a check of its own, no VB6
        r = subprocess.run([sys.executable, os.path.join(HERE, 'sysinfo', 'run_check.py')], capture_output=True, text=True)
        print(r.stdout.rstrip())
        if r.stdout.startswith('ok'):
            passed += 1
        elif r.stdout.startswith('FAIL'):
            failed += 1
    print(f"passed {passed}, failed {failed}, without VB6 output {pending}" + (f", known (Wine) {known}" if known else ''))


def mapi_store(n, target):
    """the store of the provider of Simple MAPI for a MAPI test (outside its folder: the test writes App.Path and
    the store as names); f_mapi1: no provider"""
    if n == 'f_mapi1':
        return {'LBMAPI_NONE': '1'}
    store = os.path.join(HERE, 'work', '.mapi', n)
    shutil.rmtree(store, ignore_errors=True)
    os.makedirs(store)
    return {'LBMAPI_DIR': store if target == 'linux' else 'Z:' + store.replace('/', '\\')}


def exe_vbp(src):
    """the project of the test: the Standard EXE one when the folder also has the ActiveX control project its form
    uses (f_drep*: RepCtl.vbp next to DRepTest.vbp)"""
    vbps = sorted(glob.glob(os.path.join(src, '*.vbp')))
    for v in vbps:
        for line in open(v, encoding='cp1252', errors='replace'):
            if line.strip().lower().replace(' ', '') == 'type=exe':
                return v
    return vbps[0]


def run_all(names, rt, target):
    passed = failed = pending = known = 0
    for n in names:
        src = os.path.join(HERE, n)
        vbp = exe_vbp(src)
        work = os.path.join(HERE, 'work', n)
        shutil.rmtree(work, ignore_errors=True)
        os.makedirs(os.path.join(work, 'expected'))
        for f in os.listdir(src):               # data files of the test (pictures ...)
            fp = os.path.join(src, f)
            if os.path.isfile(fp) and os.path.splitext(f)[1].lower() not in ('.frm', '.frx', '.bas', '.cls', '.vbp', '.txt'):
                shutil.copy(fp, work)
        if (n.startswith('f_rdo') or n.startswith('f_drep')) and not sqlite_odbc(target):
            print(f"---- {n}: needs unixODBC and the SQLite ODBC driver (sudo apt install unixodbc libsqliteodbc)"
                  + ("; in Wine registered as [SQLite3 ODBC Driver] too" if target != 'linux' else "")
                  + (" -- 32-bit Wine needs a 32-bit driver, which Ubuntu 24 does not have (a limit of the test"
                     " computer, not of LinBASIC Studio: on Windows the 32-bit driver is there)" if target == 'win32' else ""))
            pending += 1
            continue
        try:
            csrc, _ = compile_project([vbp])
        except Exception as e:
            print(f"FAIL {n} (compile): {e}")
            failed += 1
            continue
        exe = os.path.join(work, 'prog')
        exe = build_exe(csrc, exe, rt, target=target) or exe
        env = oracle_env()
        sp, ports = port_pair(work) if n.endswith('comm') else (None, {})
        env.update(ports)
        if n.startswith('f_mapi'):
            env.update(mapi_store(n, target))
        if n.startswith('f_inet'):
            sp = inet_server(work)
            ca = os.path.join(HERE, 'work', '.tls', 'ca.pem')    # HTTPS: the certificates the tests trust
            env['VBRT_CA_FILE'] = ca if target == 'linux' else 'Z:' + ca.replace('/', '\\')
        cmd = [exe]
        if target != 'linux':                             # Wine on a virtual X screen (Xvfb)
            cmd = ['xvfb-run', '-a', '-s', '-screen 0 1024x768x24', 'wine', exe]
            env.pop('QT_QPA_PLATFORM', None)
            env.setdefault('WINEDEBUG', '-all')
        try:
            r = subprocess.run(cmd, cwd=work, capture_output=True, text=True, timeout=180, env=env)
        except subprocess.TimeoutExpired:              # a program that does not end: a failure, the others go on
            print(f"FAIL {n}: it did not end in 180 seconds")
            failed += 1
            continue
        finally:
            if sp:
                sp.kill()
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
    return passed, failed, pending, known


if __name__ == '__main__':
    main()
