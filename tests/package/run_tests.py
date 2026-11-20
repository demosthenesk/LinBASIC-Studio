#!/usr/bin/env python3
"""lbvc --package (phase 9.5): the portable package of a program, as the Package and Deployment Wizard of VB6
(VB-TESTS/DEV-9/pd1) makes its package: the question to recompile when a source is newer than the program
(pd1-05), the package folder made when missing (pd1-09), the list of the files (pd1-10, 11: the program, its
runtime, the files the user adds, unchecked ones left out), the report (pd1-21) and the settings next to the
.vbp (pd1: <project>.PDM, here <project>.lbpkg). The package must run from another folder.

    python3 tests/package/run_tests.py [names...]
"""
import os, shutil, subprocess, sys, tarfile, time, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, '..', '..'))
sys.path.insert(0, os.path.join(ROOT, 'compiler'))
sys.path.insert(0, os.path.join(HERE, '..', 'forms'))
from lbvc import package as P                                     # noqa: E402
from lbvc.driver import TARGETS, mingw_cc                         # noqa: E402
from run_tests import oracle_env                                  # noqa: E402  (tests/forms)

LBVC = os.path.join(ROOT, 'compiler', 'lbvc.py')
WORK = os.path.join(HERE, 'work')


def fresh(name):
    d = os.path.join(WORK, name)
    shutil.rmtree(d, ignore_errors=True)
    shutil.copytree(os.path.join(HERE, name), d)
    return os.path.join(d, name + '.vbp')


def later(path, seconds=5):
    t = time.time() + seconds
    os.utime(path, (t, t))


def names_in(folder):
    out = set()
    for root, _, files in os.walk(folder):
        for n in files:
            out.add(os.path.relpath(os.path.join(root, n), folder).replace(os.sep, '/'))
    return out


def run_moved(app, exe_name, what, **env):
    """the package copied to another folder runs there (App.Path is that folder)"""
    moved = os.path.join(WORK, 'moved-' + what)
    shutil.rmtree(moved, ignore_errors=True)
    shutil.copytree(app, moved, symlinks=True)
    e = oracle_env()
    e.update(env)
    e.pop('LD_LIBRARY_PATH', None)
    e.pop('QT_PLUGIN_PATH', None)
    r = subprocess.run([os.path.join(moved, exe_name)], env=e, capture_output=True, text=True, timeout=60)
    out = os.path.join(moved, 'out.txt')
    return r, (open(out).read().strip() if os.path.exists(out) else None)


def t_console(log):
    """a program without forms: recompile when there is no program, the folder Package made, the archive"""
    vbp = fresh('p_console')
    pdir = os.path.dirname(vbp)
    exe = P.exe_path(vbp, 'linux')
    assert exe == os.path.join(pdir, 'p_console'), exe          # ExeName32 is for Windows: <name> for Linux
    assert sorted(os.path.basename(s) for s in P.stale_sources(vbp, exe)) == ['ModMain.bas', 'p_console.vbp']
    rep = P.run(vbp, 'linux', log=log)
    assert os.path.exists(exe), 'the program was not made'
    app = os.path.join(pdir, 'Package', 'p_console-linux')
    want = {'p_console', 'licenses/RUNTIME-EXCEPTION.txt', 'licenses/GPL-3.0.txt', 'licenses/LICENSES.txt'}
    assert names_in(app) == want, names_in(app)
    tgz = os.path.join(pdir, 'Package', 'p_console-linux.tar.gz')
    with tarfile.open(tgz) as t:
        got = {n for n in t.getnames() if not t.getmember(n).isdir()}
    assert got == {'p_console-linux/' + n for n in want}, got
    assert rep.startswith(f"The portable package of your application has been built as '{tgz}'."), rep
    assert 'Qt 6' not in rep
    assert P.stale_sources(vbp, exe) == []
    note = open(os.path.join(app, 'licenses', 'LICENSES.txt')).read()
    assert note.startswith('Console Package\n') and 'Qt' not in note, note
    r, out = run_moved(app, 'p_console', 'console')
    assert out == 'console ok', (r.returncode, r.stderr, out)
    s = open(os.path.join(pdir, 'p_console.lbpkg'), encoding='utf-8', newline='').read()     # CRLF, as a .PDM
    assert s == ('[Root]\r\nMost Recent Package=Portable Package 1\r\n\r\n[Package|Portable Package 1]\r\n'
                 'Target=linux\r\nFormat=tar.gz\r\nPackageFolder=Package\r\n'), repr(s)


def t_recompile(log):
    """VB6 pd1-05: a source newer than the program: Yes makes it again, No packages the program there is"""
    vbp = fresh('p_console')
    pdir = os.path.dirname(vbp)
    P.run(vbp, 'linux', log=log)
    exe = P.exe_path(vbp, 'linux')
    t0 = os.path.getmtime(exe)
    later(os.path.join(pdir, 'ModMain.bas'))
    assert [os.path.basename(s) for s in P.stale_sources(vbp, exe)] == ['ModMain.bas']
    P.run(vbp, 'linux', recompile='no', log=log)
    assert os.path.getmtime(exe) == t0, 'No: the program must stay as it is'
    lines = []
    P.run(vbp, 'linux', recompile='auto', log=lines.append)
    assert os.path.getmtime(exe) > t0, 'auto: a newer source: the program is made again'
    assert lines[0] == 'lbvc: newer than the program: ModMain.bas', lines
    os.remove(exe)
    try:
        P.run(vbp, 'linux', recompile='no', log=log)
        raise AssertionError('no program, --recompile no: an error')
    except SystemExit as e:
        assert 'not found (make the program first' in str(e), e


def t_forms(log):
    """a program with forms: Qt 6 in lib/ and plugins/ (qt.conf), the LGPL of Qt; added files, an unchecked one
    left out, the settings kept; the package runs with its own Qt"""
    vbp = fresh('p_forms')
    pdir = os.path.dirname(vbp)
    rep = P.run(vbp, 'linux', add=[os.path.join(pdir, 'readme.txt')], remove=['licenses/GPL-3.0.txt'], log=log)
    app = os.path.join(pdir, 'Package', 'p_forms-linux')
    got = names_in(app)
    assert {'p_forms', 'readme.txt', 'licenses/LGPL-3.0.txt', 'licenses/RUNTIME-EXCEPTION.txt', 'qt.conf',
            'lib/libQt6Core.so.6', 'lib/libQt6Gui.so.6', 'lib/libQt6Widgets.so.6', 'lib/libQt6XcbQpa.so.6',
            'plugins/platforms/libqxcb.so', 'plugins/platforms/libqoffscreen.so'} <= got, sorted(got)
    assert any(n.startswith('lib/libicu') for n in got) == ('licenses/ICU-Unicode.txt' in got), sorted(got)
    assert 'licenses/GPL-3.0.txt' not in got, got
    assert 'Left out:\n  licenses/GPL-3.0.txt' in rep, rep
    assert 'needs Qt 6' not in rep, rep
    assert open(os.path.join(app, 'qt.conf')).read() == '[Paths]\nPrefix = .\nPlugins = plugins\n'
    assert not any(os.path.islink(os.path.join(app, n)) for n in got), "links: the files they name"
    s = P.Settings.load(vbp).get()
    assert s['added'] == ['readme.txt'] and s['removed'] == ['licenses/GPL-3.0.txt'], s
    # the next package: the same files without saying them again
    P.run(vbp, 'linux', recompile='no', log=log)
    assert names_in(app) == got
    note = open(os.path.join(app, 'licenses', 'LICENSES.txt')).read()
    assert note.startswith('Forms Package\n') and 'LGPL-3.0' in note and '  libQt6Core.so.6\n' in note, note
    r, out = run_moved(app, 'p_forms', 'forms', QT_DEBUG_PLUGINS='1')
    assert out == 'forms ok Command1', (r.returncode, r.stderr, out)
    moved = os.path.join(WORK, 'moved-forms')
    # the Qt of the package, not the system's: the libraries (ldd) and the platform plugin (QT_DEBUG_PLUGINS)
    libs, _ = P.ldd(os.path.join(moved, 'p_forms'))
    assert os.path.realpath(libs['libQt6Core.so.6']) == os.path.realpath(os.path.join(moved, 'lib', 'libQt6Core.so.6')), libs
    assert os.path.join(moved, 'plugins', 'platforms', 'libqoffscreen.so') in r.stderr, r.stderr[-2000:]


def t_old_exe(log):
    """a program with forms made before lbvc put $ORIGIN/lib in it: it would take the system's Qt -> make it again"""
    vbp = fresh('p_forms')
    P.run(vbp, 'linux', fmt='folder', log=log)
    exe = P.exe_path(vbp, 'linux')
    data = open(exe, 'rb').read()
    assert b'$ORIGIN/lib' in data
    open(exe, 'wb').write(data.replace(b'$ORIGIN/lib', b'$ORIGIN/xyz'))
    try:
        P.run(vbp, 'linux', fmt='folder', recompile='no', log=log)
        raise AssertionError('an old program: an error')
    except SystemExit as e:
        assert 'was made by an older LinBASIC Studio' in str(e), e
    P.run(vbp, 'linux', fmt='folder', recompile='yes', log=log)
    assert b'$ORIGIN/lib' in open(exe, 'rb').read()


def png_info(png):
    """(width, height, rows of RGBA bytes) of a .png made by ico_to_png (8-bit RGBA, filter 0)"""
    import struct, zlib
    assert png[:8] == b'\x89PNG\r\n\x1a\n'
    w, h = struct.unpack_from('>II', png, 16)
    p, idat = 8, b''
    while p < len(png):
        n = struct.unpack_from('>I', png, p)[0]
        if png[p + 4:p + 8] == b'IDAT':
            idat += png[p + 8:p + 8 + n]
        p += 12 + n
    raw = zlib.decompress(idat)
    return w, h, [raw[y * (4 * w + 1) + 1:(y + 1) * (4 * w + 1)] for y in range(h)]


def t_icon(log):
    """the icon of the AppImage: an .ico (DIB, 32 bits with alpha; 4 bits with a palette and the AND mask) as a
    .png, the top row first; a form without Icon: the icon of LinBASIC Studio"""
    import struct
    # 2x2, 32 bits: bottom-up rows, BGRA; top-left red, top-right green, bottom-left blue, bottom-right white
    dib = struct.pack('<IiiHHIIiiII', 40, 2, 4, 1, 32, 0, 0, 0, 0, 0, 0)
    dib += bytes((255, 0, 0, 255, 255, 255, 255, 128)) + bytes((0, 0, 255, 255, 0, 255, 0, 255)) + bytes(8)
    ico = struct.pack('<HHH', 0, 1, 1) + struct.pack('<BBBBHHII', 2, 2, 0, 0, 1, 32, len(dib), 22) + dib
    w, h, rows = png_info(P.ico_to_png(ico))
    assert (w, h) == (2, 2)
    assert rows[0] == bytes((255, 0, 0, 255, 0, 255, 0, 255)), rows
    assert rows[1] == bytes((0, 0, 255, 255, 255, 255, 255, 128)), rows
    # 2x1... 2x2, 4 bits: palette 0 black, 1 yellow; the AND mask makes the bottom-right transparent
    pal = bytes((0, 0, 0, 0)) + bytes((0, 255, 255, 0)) + bytes(4 * 14)
    xor = bytes((0x01, 0, 0, 0)) + bytes((0x10, 0, 0, 0))        # bottom row: 0,1 ; top row: 1,0
    andm = bytes((0x40, 0, 0, 0)) + bytes(4)                     # bottom row: the second pixel
    dib = struct.pack('<IiiHHIIiiII', 40, 2, 4, 1, 4, 0, 0, 0, 0, 0, 0) + pal + xor + andm
    ico = struct.pack('<HHH', 0, 1, 1) + struct.pack('<BBBBHHII', 2, 2, 16, 0, 1, 4, len(dib), 22) + dib
    w, h, rows = png_info(P.ico_to_png(ico))
    assert rows[0] == bytes((255, 255, 0, 255, 0, 0, 0, 255)), rows
    assert rows[1] == bytes((0, 0, 0, 255, 255, 255, 0, 0)), rows
    png = b'\x89PNG\r\n\x1a\n' + b'rest'
    assert P.ico_to_png(png) == png
    vbp = fresh('p_forms')
    own = open(os.path.join(ROOT, 'ide', 'lbstudio', 'res', 'linbasic-studio-256.png'), 'rb').read()
    assert P.project_icon(vbp) == own


def t_appimage(log):
    """--format appimage: one file; the files the user adds next to it; App.Path is the folder of the .AppImage
    (the user's decision, 2026-11-19); it runs from another folder"""
    tool, _ = P.appimage_tools()
    if not tool:
        return 'SKIP (no appimagetool: packaging/build-appimage.sh fetches it into build/tools)'
    vbp = fresh('p_forms')
    pdir = os.path.dirname(vbp)
    rep = P.run(vbp, 'linux', fmt='appimage', add=[os.path.join(pdir, 'readme.txt')], log=log)
    out = os.path.join(pdir, 'Package', 'p_forms-linux')
    got = names_in(out)
    assert {'p_forms-x86_64.AppImage', 'readme.txt', 'licenses/LGPL-3.0.txt', 'licenses/LGPL-2.1.txt',
            'licenses/LICENSES.txt'} <= got, sorted(got)
    assert not any(n.startswith(('lib/', 'plugins/')) for n in got), sorted(got)
    ai = os.path.join(out, 'p_forms-x86_64.AppImage')
    assert os.access(ai, os.X_OK)
    assert rep.startswith(f"The AppImage of your application has been built as '{ai}'."), rep
    moved = os.path.join(WORK, 'moved-appimage')
    shutil.rmtree(moved, ignore_errors=True)
    os.makedirs(moved)
    shutil.copy2(ai, moved)
    e = oracle_env()
    e['APPIMAGE_EXTRACT_AND_RUN'] = '1'          # (no FUSE needed)
    e.pop('LD_LIBRARY_PATH', None)
    r = subprocess.run([os.path.join(moved, 'p_forms-x86_64.AppImage')], env=e, capture_output=True, text=True,
                       timeout=180)
    o = os.path.join(moved, 'out.txt')
    assert os.path.exists(o) and open(o).read().strip() == 'forms ok Command1', (r.returncode, r.stderr[-2000:])
    try:
        P.run(vbp, 'win64', fmt='appimage', log=log)
        raise AssertionError('an AppImage for Windows: an error')
    except SystemExit as e:
        assert 'a program for Linux' in str(e), e


def t_folder_zip(log):
    """--format folder: no archive; zip; another package folder (VB6 pd1-08, 09: made when missing)"""
    vbp = fresh('p_console')
    pdir = os.path.dirname(vbp)
    rep = P.run(vbp, 'linux', fmt='folder', log=log)
    assert not os.path.exists(os.path.join(pdir, 'Package', 'p_console-linux.tar.gz'))
    assert rep.startswith('The portable package of your application has been built in the folder'), rep
    other = os.path.join(WORK, 'elsewhere', 'out')
    shutil.rmtree(os.path.join(WORK, 'elsewhere'), ignore_errors=True)
    P.run(vbp, 'linux', folder=other, fmt='zip', recompile='no', log=log)
    with zipfile.ZipFile(os.path.join(other, 'p_console-linux.zip')) as z:
        assert 'p_console-linux/p_console' in z.namelist(), z.namelist()
    s = P.Settings.load(vbp).get()
    assert s['Format'] == 'zip' and s['PackageFolder'] == other, s     # outside the project: the whole path


def t_cli(log):
    """lbvc p.vbp --package: the report on stdout"""
    vbp = fresh('p_console')
    r = subprocess.run([sys.executable, LBVC, vbp, '--package', '--format', 'folder'], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert r.stdout.startswith('The portable package of your application has been built in the folder'), r.stdout
    assert 'lbvc: making ' in r.stderr, r.stderr
    r = subprocess.run([sys.executable, LBVC, vbp, '--package', '--format', 'cab'], capture_output=True, text=True)
    assert r.returncode == 2 and 'invalid choice' in r.stderr, r.stderr
    dll = fresh('p_dll')
    r = subprocess.run([sys.executable, LBVC, dll, '--package'], capture_output=True, text=True)
    assert r.returncode != 0 and 'an ActiveX DLL has no program of its own to package' in r.stderr, r.stderr


def t_windows(log):
    """win64 (mingw-w64): the .exe and the DLLs of Qt in the folder and in the .zip"""
    try:
        subprocess.run([mingw_cc('win64'), '--version'], capture_output=True, check=True)
    except (OSError, subprocess.CalledProcessError, SystemExit):
        return 'SKIP (no mingw-w64)'
    from lbvc.driver import find_runtime
    if not os.path.isdir(os.path.join(find_runtime(), 'win64', 'qt', 'bin')):
        return 'SKIP (no Qt for win64 in runtime/win64)'
    vbp = fresh('p_forms')
    pdir = os.path.dirname(vbp)
    P.run(vbp, 'win64', log=log)
    assert P.exe_path(vbp, 'win64').endswith('p_forms.exe')
    with zipfile.ZipFile(os.path.join(pdir, 'Package', 'p_forms-win64.zip')) as z:
        got = set(z.namelist())
    for n in ('p_forms.exe', 'platforms/qwindows.dll', 'libstdc++-6.dll', 'licenses/LGPL-3.0.txt',
              'licenses/GCC-Runtime-Library-Exception-3.1.txt', 'licenses/winpthreads-MIT.txt'):
        assert 'p_forms-win64/' + n in got, (n, sorted(got))
    assert any(n.startswith('p_forms-win64/Qt') and n.endswith('Widgets.dll') for n in got), sorted(got)
    assert 'win64' in TARGETS


def t_windows_tcc(log):
    """win64 with the cross Tiny C Compiler of LinBASIC Studio (the AppImage makes Windows programs so, without
    mingw-w64): the .exe uses the runtime DLLs (vbrt.dll, vbforms.dll), they and Qt 5 go into the .zip"""
    from lbvc.driver import find_runtime
    rt = find_runtime()
    if not os.path.exists(os.path.join(rt, 'win64', 'tcc', 'cross-tcc')) or \
            not os.path.exists(os.path.join(rt, 'win64', 'dll', 'vbrt.dll')):
        return 'SKIP (no runtime/win64/tcc/cross-tcc and dll/: packaging/build-tcc.sh, runtime/build-dll-win.sh)'
    old = os.environ.get('LBVC_WIN_CC')
    os.environ['LBVC_WIN_CC'] = 'tcc'
    try:
        vbp = fresh('p_forms')
        pdir = os.path.dirname(vbp)
        P.run(vbp, 'win64', log=log)
        exe = P.exe_path(vbp, 'win64')
        assert b'vbrt.dll' in open(exe, 'rb').read() and b'vbforms.dll' in open(exe, 'rb').read(), \
            'the program of tcc imports the runtime DLLs'
        with zipfile.ZipFile(os.path.join(pdir, 'Package', 'p_forms-win64.zip')) as z:
            got = set(z.namelist())
        for n in ('p_forms.exe', 'vbrt.dll', 'vbforms.dll', 'Qt5Core.dll', 'Qt5Widgets.dll', 'platforms/qwindows.dll',
                  'libstdc++-6.dll', 'licenses/LGPL-3.0.txt'):
            assert 'p_forms-win64/' + n in got, (n, sorted(got))
    finally:
        if old is None:
            os.environ.pop('LBVC_WIN_CC', None)
        else:
            os.environ['LBVC_WIN_CC'] = old


def t_win32_after_win64(log):
    """E8: the IDE's Make for win64 left p_forms.exe (x86-64) in the project folder (ExeName32); the package for
    win32 must not take it: the program for win32 is made in p_forms-win32/ and the .zip has a PE i386 program"""
    from lbvc.driver import find_runtime
    rt = find_runtime()
    for t in ('win32', 'win64'):
        if not os.path.exists(os.path.join(rt, t, 'tcc', 'cross-tcc')) or \
                not os.path.exists(os.path.join(rt, t, 'dll', 'vbrt.dll')):
            return f'SKIP (no runtime/{t}/tcc/cross-tcc and dll/)'
    old = os.environ.get('LBVC_WIN_CC')
    os.environ['LBVC_WIN_CC'] = 'tcc'
    try:
        vbp = fresh('p_forms')
        pdir = os.path.dirname(vbp)
        P.run(vbp, 'win64', log=log)                              # ExeName32="p_forms.exe": the project folder
        made64 = os.path.join(pdir, 'p_forms.exe')
        assert P.program_target(made64) == 'win64'
        assert P.exe_path(vbp, 'win64') == made64
        exe32 = P.exe_path(vbp, 'win32')
        assert exe32 == os.path.join(pdir, 'p_forms-win32', 'p_forms.exe'), exe32
        P.run(vbp, 'win32', log=log)
        assert P.program_target(exe32) == 'win32'
        assert P.program_target(made64) == 'win64', 'the program of the win64 Make is left alone'
        with zipfile.ZipFile(os.path.join(pdir, 'Package', 'p_forms-win32.zip')) as z:
            head = z.read('p_forms-win32/p_forms.exe')[:4096]
        pe = int.from_bytes(head[0x3c:0x40], 'little')
        assert head[pe + 4:pe + 6] == b'\x4c\x01', 'the .exe in p_forms-win32.zip is PE i386'
        bad = os.path.join(pdir, 'p_forms-win32', 'p_forms.exe')
        shutil.copyfile(made64, bad)                              # a wrong program where win32's should be
        later(bad)
        try:
            P.run(vbp, 'win32', recompile='no', log=log)
            raise AssertionError('a program for win64 was packaged for win32')
        except SystemExit as e:
            assert 'win64' in str(e), e
        P.run(vbp, 'win32', log=log)                              # auto: made again
        assert P.program_target(bad) == 'win32'
        assert P.program_target(os.path.join(pdir, 'p_forms.vbp')) is None
        assert P.program_target(os.path.join(pdir, 'missing.exe')) is None
    finally:
        if old is None:
            os.environ.pop('LBVC_WIN_CC', None)
        else:
            os.environ['LBVC_WIN_CC'] = old


TESTS = {'t_console': t_console, 't_recompile': t_recompile, 't_forms': t_forms, 't_old_exe': t_old_exe,
         't_icon': t_icon, 't_appimage': t_appimage, 't_folder_zip': t_folder_zip,
         't_cli': t_cli, 't_windows': t_windows, 't_windows_tcc': t_windows_tcc,
         't_win32_after_win64': t_win32_after_win64}


def main():
    names = sys.argv[1:] or list(TESTS)
    os.makedirs(WORK, exist_ok=True)
    passed = 0
    for n in names:
        print(f'---- {n}', flush=True)
        try:
            r = TESTS[n](lambda s: None)
            print(f'SKIP {n}: {r}' if r else f'PASS {n}')       # ce.sh shows the SKIP lines, not the PASS ones
            passed += 1
        except SystemExit as e:
            print(f'FAIL {n}: SystemExit {e}')
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f'FAIL {n}: {e!r}')
    print(f'{passed}/{len(names)} passed')
    return 0 if passed == len(names) else 1


if __name__ == '__main__':
    sys.exit(main())
