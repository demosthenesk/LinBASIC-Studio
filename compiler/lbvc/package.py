# LinBASIC Studio - lbvc --package: the Package and Deployment Wizard of VB6, for portable programs (phase 9.5)
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#
# VB6 (VB-TESTS/DEV-9/pd1) asks to recompile when a source is newer than the .exe, assembles the package in
# <project folder>\Package (made when missing), lists the files of the package (the program, its runtime,
# the files the user adds; a checkbox each), builds it and shows a report; its settings are kept next to the
# .vbp (<project>.PDM). LinBASIC Studio makes portable packages only (no setup program, the user's decision):
# the folder <Package>/<program>-<target>/ with everything the program needs next to it, and the same in an
# archive: a .zip for Windows, a .tar.gz for Linux; or, for Linux, an AppImage (one file; App.Path is the folder
# of the .AppImage, the user's decision: the files the user adds go next to it). The settings: <project>.lbpkg
# next to the .vbp.
import os, shutil, tarfile, zipfile

from .project import load_project

FORMATS = ('folder', 'zip', 'tar.gz', 'appimage')
DEFAULT_SCRIPT = 'Portable Package 1'
LICENSE_NOTE = 'LICENSES.txt'


def default_format(target):
    return 'tar.gz' if target == 'linux' else 'zip'


def program_name(project):
    """the name of the made program without .exe: ExeName32 of the last Make, else the project's name"""
    exe = (project.get('ExeName32') or '').strip('"').replace('\\', '/')
    return os.path.splitext(os.path.basename(exe))[0] if exe else project.name


def program_target(path):
    """the system a made program is for, from its header: linux (ELF), win32 / win64 (PE i386 / x86-64);
    None when it cannot be told"""
    try:
        with open(path, 'rb') as f:
            head = f.read(4096)
    except OSError:
        return None
    if head[:4] == b'\x7fELF':
        return 'linux'
    if head[:2] == b'MZ' and len(head) >= 0x40:
        pe = int.from_bytes(head[0x3c:0x40], 'little')
        if head[pe:pe + 4] == b'PE\0\0':
            return {0x14c: 'win32', 0x8664: 'win64'}.get(int.from_bytes(head[pe + 4:pe + 6], 'little'))
    return None


def exe_path(vbp, target, project=None):
    """where Make puts the program (as the IDE: ExeName32 in the project folder when it is for this system,
    else <name> for Linux, <name>-<target>/<name>.exe for Windows). An ExeName32 that is there but made for
    the other Windows (win64 for win32, E8) is left alone: <name>-<target>/<name>.exe"""
    project = project or load_project(vbp)
    folder = os.path.dirname(os.path.abspath(vbp))
    exe = (project.get('ExeName32') or '').strip('"').replace('\\', '/')
    if exe and (target == 'linux') == (not exe.lower().endswith('.exe')):
        exe = os.path.normpath(exe if os.path.isabs(exe) else os.path.join(folder, exe))
        if program_target(exe) in (None, target):
            return exe
    name = program_name(project)
    if target == 'linux':
        return os.path.join(folder, name)
    return os.path.join(folder, f'{name}-{target}', name + '.exe')


def project_sources(vbp, project=None):
    """the files the program is made of: the .vbp, its members with their .frx / .ctx, the resource file"""
    project = project or load_project(vbp)
    files = [os.path.abspath(vbp)]
    for kind, _, rel in project.members():
        p = project.resolve(rel)
        files.append(p)
        base = os.path.splitext(p)[0]
        for ext in ('.frx', '.FRX', '.ctx', '.CTX', '.dsx', '.DSX'):
            if os.path.exists(base + ext):
                files.append(base + ext)
                break
    res = project.get('ResFile32')
    if res:
        files.append(project.resolve(res.strip('"')))
    return files


def stale_sources(vbp, exe):
    """VB6 (pd1-05): "The following source files for this project are newer than the executable file":
    the sources newer than the program (all of them when there is no program)"""
    srcs = [s for s in project_sources(vbp) if os.path.exists(s)]
    if not os.path.exists(exe):
        return srcs
    t = os.path.getmtime(exe)
    return [s for s in srcs if os.path.getmtime(s) > t]


# ---------------------------------------------------------------- the settings (<project>.lbpkg)
class Settings:
    """the settings of the packages of a project, by script name (VB6: <project>.PDM, "Script name")"""

    def __init__(self, path):
        self.path = path
        self.recent = None
        self.scripts = {}               # name -> {'Target':.., 'Format':.., 'PackageFolder':.., 'added': [], 'removed': []}

    @classmethod
    def load(cls, vbp):
        s = cls(os.path.splitext(os.path.abspath(vbp))[0] + '.lbpkg')
        if not os.path.exists(s.path):
            return s
        cur = None
        for line in open(s.path, encoding='utf-8-sig').read().splitlines():
            line = line.strip()
            if not line or line.startswith(';'):
                continue
            if line.startswith('[') and line.endswith(']'):
                sec = line[1:-1]
                cur = s.scripts.setdefault(sec.split('|', 1)[1], {'added': [], 'removed': []}) \
                    if sec.startswith('Package|') else None
                continue
            k, _, v = line.partition('=')
            k, v = k.strip(), v.strip()
            if cur is None:
                if k == 'Most Recent Package':
                    s.recent = v
            elif k.startswith('Added') and k[5:].isdigit():
                cur['added'].append(v)
            elif k.startswith('Removed') and k[7:].isdigit():
                cur['removed'].append(v)
            else:
                cur[k] = v
        return s

    def get(self, name=None):
        return self.scripts.get(name or self.recent or DEFAULT_SCRIPT, {'added': [], 'removed': []})

    def put(self, name, values):
        self.scripts[name] = values
        self.recent = name

    def save(self):
        out = ['[Root]', f'Most Recent Package={self.recent or DEFAULT_SCRIPT}']
        for name, v in self.scripts.items():
            out += ['', f'[Package|{name}]']
            for k in ('Target', 'Format', 'PackageFolder'):
                if v.get(k):
                    out.append(f'{k}={v[k]}')
            out += [f'Added{i}={a}' for i, a in enumerate(v.get('added', []), 1)]
            out += [f'Removed{i}={r}' for i, r in enumerate(v.get('removed', []), 1)]
        with open(self.path, 'w', encoding='utf-8', newline='\r\n') as f:
            f.write('\n'.join(out) + '\n')


def rel_to(path, folder):
    """a path in the settings: relative to the project folder when it is inside it"""
    p = os.path.abspath(path)
    try:
        r = os.path.relpath(p, folder)
    except ValueError:                  # another drive (Windows)
        return p
    return p if r.startswith('..') else r


# ---------------------------------------------------------------- the files of the package
class PackageFile:
    """a file of the package: a copy of source, or data (bytes) written by lbvc (qt.conf)"""

    def __init__(self, source, name, kind, checked=True, data=None):
        self.source, self.name, self.kind, self.checked, self.data = source, name, kind, checked, data

    def __repr__(self):
        return f'PackageFile({self.name!r}, {self.kind}, {self.source!r}, {self.checked})'


def license_files(csrc, target, icu=False, appimage=False):
    """[(text, name)]: the licenses that go with a made program (licenses/README.md, "The programs you make");
    icu: the package carries the ICU of Qt (Linux); appimage: the runtime of the AppImage (libfuse: LGPL-2.1)"""
    from .driver import SRC_ROOT, uses_forms, uses_db, TARGETS
    lic = os.path.join(SRC_ROOT, 'licenses')
    out = [(os.path.join(SRC_ROOT, 'RUNTIME-EXCEPTION'), 'RUNTIME-EXCEPTION.txt'),
           (os.path.join(SRC_ROOT, 'LICENSE'), 'GPL-3.0.txt')]
    if uses_forms(csrc):
        out.append((os.path.join(lic, 'LGPL-3.0.txt'), 'LGPL-3.0.txt'))
        out.append((os.path.join(lic, 'LGPL-2.1.txt'), 'LGPL-2.1.txt'))      # CHMLib (the .chm help), linked in
        if target in TARGETS:
            out += [(os.path.join(lic, 'GCC-Runtime-Library-Exception-3.1.txt'), 'GCC-Runtime-Library-Exception-3.1.txt'),
                    (os.path.join(lic, 'winpthreads-MIT.txt'), 'winpthreads-MIT.txt')]
    if icu:
        out.append((os.path.join(lic, 'ICU-Unicode.txt'), 'ICU-Unicode.txt'))
    if appimage:
        if not uses_forms(csrc):
            out.append((os.path.join(lic, 'LGPL-2.1.txt'), 'LGPL-2.1.txt'))
        _, rtf = appimage_tools()
        if rtf and os.path.exists(rtf + '.LICENSE'):
            out.append((rtf + '.LICENSE', 'AppImage-runtime-LICENSE.txt'))
    if uses_db(csrc):
        out += [(os.path.join(lic, 'SQLite-public-domain.txt'), 'SQLite-public-domain.txt')]
    for s, _ in out:
        if not os.path.exists(s):
            raise SystemExit(f"lbvc: {s} not found (the licenses that go with a program: LinBASIC Studio is "
                             f"not complete)")
    return [(s, 'licenses/' + n) for s, n in out]


def license_note(csrc, target, title, qt_libs=(), appimage=False):
    from .driver import uses_forms, uses_db, TARGETS
    lines = [f'{title}', '', 'This program was made with LinBASIC Studio. Its runtime (libvbrt, libvbforms, libvbdb) is',
             'GPL-3.0-or-later with the LinBASIC Studio Runtime Library Exception (licenses/RUNTIME-EXCEPTION.txt):',
             'the program may be given to others under terms of its author\'s choice.']
    if uses_forms(csrc):
        lines += ['', 'Forms: Qt (The Qt Company), LGPL-3.0 (licenses/LGPL-3.0.txt, licenses/GPL-3.0.txt), as shared',
                  'libraries that can be replaced. Source: https://download.qt.io/archive/qt/',
                  'The .chm help: CHMLib (Jed Wing), LGPL-2.1 (licenses/LGPL-2.1.txt), linked in.']
        if target in TARGETS:
            lines += ['The C++ runtime of mingw-w64 that Qt needs: libstdc++ / libgcc (GPL-3.0 with the GCC Runtime',
                      'Library Exception 3.1), winpthreads (MIT).']
        if qt_libs:
            lines += ['The libraries in lib/ (with the plugins of Qt in plugins/) are Qt 6 and the libraries it needs,',
                      'each under its own license: Qt LGPL-3.0, ICU the Unicode license (licenses/ICU-Unicode.txt), the',
                      'xcb / xkbcommon libraries MIT/X11, double-conversion and pcre2 BSD, md4c MIT, libb2 CC0:']
            lines += ['  ' + n for n in qt_libs]
    if uses_db(csrc):
        lines += ['', 'DAO / ADO: SQLite (public domain), linked in.']
    if appimage:
        lines += ['', 'The AppImage: its runtime is of the AppImage project (https://github.com/AppImage/type2-runtime,',
                  'MIT; it contains libfuse, LGPL-2.1, licenses/LGPL-2.1.txt, squashfuse BSD-2-Clause and zstd BSD).']
    return '\n'.join(lines) + '\n'


def package_files(vbp, target, csrc, exe, added=(), removed=(), fmt=None):
    """VB6 "Included Files" (pd1-10, 11): the program, the files of its runtime (Linux: Qt too), the licenses,
    the files the user added; the ones the user unchecked (removed: their names) stay in the list, unchecked"""
    from .driver import find_runtime, runtime_files, uses_forms
    files = [PackageFile(exe, os.path.basename(exe), 'program')]
    rt_files = runtime_files(csrc, find_runtime(), target, exe)
    for src, rel in rt_files:
        files.append(PackageFile(src, rel.replace(os.sep, '/'), 'runtime'))
    if target == 'linux' and uses_forms(csrc):
        for src, rel in linux_qt_files([exe] + [s for s, _ in rt_files]):
            files.append(PackageFile(src, rel, 'qt'))
        files.append(PackageFile(None, 'qt.conf', 'qt', data=QT_CONF))
    icu = any(f.kind == 'qt' and f.name.startswith('lib/libicu') for f in files)
    for src, rel in license_files(csrc, target, icu, fmt == 'appimage'):
        files.append(PackageFile(src, rel, 'license'))
    folder = os.path.dirname(os.path.abspath(vbp))
    for a in added:
        p = a if os.path.isabs(a) else os.path.join(folder, a)
        files.append(PackageFile(os.path.normpath(p), os.path.basename(p), 'added'))
    gone = {r.lower() for r in removed}
    for f in files:
        if f.name.lower() in gone and f.kind != 'program':
            f.checked = False
    return files


# ---------------------------------------------------------------- Qt 6 for Linux
# qt.conf next to the program: Qt takes its plugins from plugins/ (the libraries: the DT_RPATH $ORIGIN/lib of the
# program, driver.PACKAGE_LIBDIR)
QT_CONF = b'[Paths]\nPrefix = .\nPlugins = plugins\n'
# the plugins a program of LinBASIC Studio needs: the window system (xcb; offscreen for tests and servers), OpenGL
# of xcb, the pictures it loads (LoadPicture: .gif, .ico, .jpg), the input of accented letters (compose, ibus)
QT_PLUGINS = ('platforms/libqxcb.so', 'platforms/libqoffscreen.so', 'xcbglintegrations/libqxcb-glx-integration.so',
              'xcbglintegrations/libqxcb-egl-integration.so', 'imageformats/libqgif.so', 'imageformats/libqico.so',
              'imageformats/libqjpeg.so', 'platforminputcontexts/libcomposeplatforminputcontextplugin.so',
              'platforminputcontexts/libibusplatforminputcontextplugin.so')
# the libraries that go with Qt: Qt itself, ICU, and the ones a computer without Qt may not have (Ubuntu:
# libxcb-cursor0 is missing on many); the C library, X11, OpenGL, fontconfig, glib ... are the system's
QT_DEPS = ('libQt6', 'libicu', 'libxcb-cursor', 'libxcb-icccm', 'libxcb-image', 'libxcb-keysyms', 'libxcb-render-util',
           'libxcb-util', 'libxcb-xkb', 'libxkbcommon', 'libdouble-conversion', 'libb2', 'libpcre2-16', 'libmd4c',
           'libts.so')


def ldd(path):
    """({name: file}, [names not found]) of the shared libraries of path, as the program would load them (the
    Qt of the AppImage of LinBASIC Studio when the IDE runs from it: LBVC_QT_LIBDIR)"""
    import subprocess
    env = dict(os.environ)
    qt = os.environ.get('LBVC_QT_LIBDIR')
    if qt:
        env['LD_LIBRARY_PATH'] = qt + (':' + env['LD_LIBRARY_PATH'] if env.get('LD_LIBRARY_PATH') else '')
    r = subprocess.run(['ldd', path], capture_output=True, text=True, env=env)
    out, missing = {}, []
    for line in r.stdout.splitlines():
        name, sep, rest = line.strip().partition(' => ')
        if not sep:
            continue
        p = rest.rsplit(' (', 1)[0].strip()
        if p == 'not found':
            missing.append(name)
        else:
            out[name] = p
    return out, missing


def qt_plugin_dir(core):
    """the plugins of the Qt whose libQt6Core.so.6 is core"""
    import subprocess
    d = os.path.dirname(os.path.realpath(core))
    for c in (os.path.join(d, '..', 'plugins'), os.path.join(d, 'qt6', 'plugins')):
        if os.path.isdir(os.path.join(c, 'platforms')):
            return os.path.normpath(c)
    for cmd in (['qtpaths6', '--query', 'QT_INSTALL_PLUGINS'], ['qtpaths', '--plugin-dir']):
        try:
            p = subprocess.run(cmd, capture_output=True, text=True).stdout.strip()
        except OSError:
            continue
        if p and os.path.isdir(os.path.join(p, 'platforms')):
            return p
    raise SystemExit(f"lbvc: the plugins of the Qt 6 of {core} not found")


def linux_qt_files(objs):
    """[(file, name in the package)]: the Qt 6 a program with forms uses (lib/, plugins/), so that its package
    runs on a Linux without Qt"""
    libs = {}
    for o in objs:
        if o.endswith('.so') or os.access(o, os.X_OK):
            found, missing = ldd(o)
            if missing:
                raise SystemExit(f"lbvc: {os.path.basename(o)} needs {', '.join(missing)}, not found on this computer")
            libs.update(found)
    core = libs.get('libQt6Core.so.6')
    if not core:
        raise SystemExit("lbvc: the program does not use Qt 6 (libQt6Core.so.6): made without forms?")
    pdir = qt_plugin_dir(core)
    plugins = []
    for rel in QT_PLUGINS:
        p = os.path.join(pdir, rel)
        found, missing = ldd(p) if os.path.exists(p) else ({}, ['the file'])
        if not missing:
            plugins.append((p, 'plugins/' + rel))
            libs.update(found)
        elif rel == 'platforms/libqxcb.so':           # without it no window opens
            raise SystemExit(f"lbvc: {p}: {', '.join(missing)} not found (the X11 plugin of Qt 6)")
    out = [(p, 'lib/' + n) for n, p in sorted(libs.items()) if n.startswith(QT_DEPS)]
    return out + plugins


# ---------------------------------------------------------------- AppImage (Linux)
PNG_SIG = b'\x89PNG\r\n\x1a\n'


def ico_to_png(data):
    """the largest picture of an .ico as a .png (an AppImage wants .png / .svg): the .ico of a form (Icon), with
    1, 4, 8, 24 or 32 bits a pixel and its AND mask, or a .png inside it"""
    import struct, zlib
    if data[:8] == PNG_SIG:
        return data
    best = None
    for i in range(struct.unpack_from('<H', data, 4)[0]):
        w, h, _, _, _, bpp, size, off = struct.unpack_from('<BBBBHHII', data, 6 + 16 * i)
        key = ((w or 256) * (h or 256), bpp)
        if best is None or key > best[0]:
            best = (key, size, off)
    img = data[best[2]:best[2] + best[1]]
    if img[:8] == PNG_SIG:
        return img
    hsz, w, bh, _, bpp = struct.unpack_from('<IiiHH', img, 0)
    used = struct.unpack_from('<I', img, 32)[0]
    h = abs(bh) // 2
    p, pal = hsz, []
    if bpp <= 8:
        for k in range(used or (1 << bpp)):
            b, g, r = img[p + 4 * k:p + 4 * k + 3]
            pal.append((r, g, b))
        p += 4 * (used or (1 << bpp))
    stride, mstride = ((w * bpp + 31) // 32) * 4, ((w + 31) // 32) * 4
    xor, mask = img[p:p + stride * h], img[p + stride * h:p + stride * h + mstride * h]
    alpha = bpp == 32 and any(xor[i + 3] for i in range(0, len(xor), 4))
    raw = bytearray()
    for y in range(h):
        src = h - 1 - y                         # a DIB is bottom-up
        raw.append(0)
        for x in range(w):
            if bpp == 32:
                b, g, r, a = xor[src * stride + 4 * x:src * stride + 4 * x + 4]
            elif bpp == 24:
                b, g, r = xor[src * stride + 3 * x:src * stride + 3 * x + 3]
                a = 255
            else:
                byte = xor[src * stride + (x * bpp) // 8]
                idx = (byte >> (8 - bpp - (x * bpp) % 8)) & ((1 << bpp) - 1)
                r, g, b = pal[idx] if idx < len(pal) else (0, 0, 0)
                a = 255
            if not alpha:
                a = 255
                if mask and (mask[src * mstride + x // 8] >> (7 - x % 8)) & 1:
                    a = 0
            raw += bytes((r, g, b, a))

    def chunk(kind, d):
        return struct.pack('>I', len(d)) + kind + d + struct.pack('>I', zlib.crc32(kind + d) & 0xffffffff)
    return PNG_SIG + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 6, 0, 0, 0)) + \
        chunk(b'IDAT', zlib.compress(bytes(raw), 9)) + chunk(b'IEND', b'')


def project_icon(vbp, project=None):
    """the icon of the program as a .png: the Icon of the IconForm of the .vbp (VB6's Make: "Icon"), else of
    the startup form; without one the icon of LinBASIC Studio"""
    import struct
    from .driver import SRC_ROOT
    from .project import load_source
    from .project.source import FrxRef
    from .project.frx import FrxStore
    from .project.fsutil import ci_path
    project = project or load_project(vbp)
    want = (project.get('IconForm') or project.startup or '').strip('"').lower()
    for kind, _, rel in project.members():
        if kind != 'Form':
            continue
        path = project.resolve(rel)
        try:
            sf = load_source(path)
        except (OSError, ValueError):
            continue
        if not sf.root or (sf.name or '').lower() != want:
            continue
        v = sf.root.get('Icon')
        if isinstance(v, FrxRef):
            try:
                fmt, img = FrxStore.load(ci_path(os.path.join(os.path.dirname(path), v.file))).picture(v.offset)
                if fmt == 'ico' or img[:8] == PNG_SIG:
                    return ico_to_png(img)
            except (OSError, ValueError, IndexError, TypeError, struct.error):
                pass
        break
    with open(os.path.join(SRC_ROOT, 'ide', 'lbstudio', 'res', 'linbasic-studio-256.png'), 'rb') as f:
        return f.read()


def appimage_tools():
    """(appimagetool, the type2 runtime or None): LBVC_APPIMAGETOOL / LBVC_APPIMAGE_RUNTIME, runtime/appimage (the
    AppImage of LinBASIC Studio carries them), build/tools (packaging/build-appimage.sh), appimagetool in PATH"""
    from .driver import SRC_ROOT, find_runtime
    tool = os.environ.get('LBVC_APPIMAGETOOL')
    rtf = os.environ.get('LBVC_APPIMAGE_RUNTIME')
    for d in (os.path.join(find_runtime(), 'appimage'), os.path.join(SRC_ROOT, 'build', 'tools')):
        if not tool and os.path.exists(os.path.join(d, 'appimagetool')):
            tool = os.path.join(d, 'appimagetool')
        if not rtf and os.path.exists(os.path.join(d, 'runtime-x86_64')):
            rtf = os.path.join(d, 'runtime-x86_64')
    return tool or shutil.which('appimagetool'), rtf


APPRUN = """#!/bin/sh
# the AppImage of a program made with LinBASIC Studio: App.Path is the folder of the .AppImage file
HERE="$(dirname "$(readlink -f "$0")")"
[ -n "$APPIMAGE" ] && export VBRT_APP_PATH="$(dirname "$APPIMAGE")"
exec "$HERE/usr/bin/{name}" "$@"
"""


def make_appimage(app_files, added, out_dir, name, title, icon, console):
    """an AppImage of the files of the program (app_files: the folder made as for a .tar.gz) -> its path"""
    import subprocess, tempfile
    tool, rtf = appimage_tools()
    if not tool:
        raise SystemExit("lbvc: appimagetool not found (LinBASIC Studio carries it in its AppImage; from the source: "
                         "packaging/build-appimage.sh fetches it into build/tools, or set LBVC_APPIMAGETOOL)")
    with tempfile.TemporaryDirectory() as td:
        appdir = os.path.join(td, 'AppDir')
        shutil.copytree(app_files, os.path.join(appdir, 'usr', 'bin'))
        for a in added:
            p = os.path.join(appdir, 'usr', 'bin', a)
            if os.path.exists(p):
                os.remove(p)                    # next to the .AppImage, where App.Path is
        with open(os.path.join(appdir, 'AppRun'), 'w') as f:
            f.write(APPRUN.format(name=name))
        os.chmod(os.path.join(appdir, 'AppRun'), 0o755)
        safe = title.replace('\n', ' ').strip() or name
        with open(os.path.join(appdir, name + '.desktop'), 'w') as f:
            f.write(f'[Desktop Entry]\nType=Application\nName={safe}\nExec={name}\nIcon={name}\n'
                    f'Terminal={"true" if console else "false"}\nCategories=Utility;\n')
        for n in (name + '.png', '.DirIcon'):
            with open(os.path.join(appdir, n), 'wb') as f:
                f.write(icon)
        out = os.path.join(out_dir, f'{name}-x86_64.AppImage')
        env = {k: v for k, v in os.environ.items() if k not in ('APPIMAGE', 'APPDIR', 'ARGV0', 'OWD')}
        env.update(ARCH='x86_64', APPIMAGE_EXTRACT_AND_RUN='1')
        if 'LD_LIBRARY_PATH_ORIG' in env:       # the IDE of the AppImage: not its Qt for the tools
            env['LD_LIBRARY_PATH'] = env['LD_LIBRARY_PATH_ORIG']
        else:
            env.pop('LD_LIBRARY_PATH', None)
        cmd = [tool, '--no-appstream'] + (['--runtime-file', rtf] if rtf else []) + [appdir, out]
        r = subprocess.run(cmd, capture_output=True, text=True, env=env)
        if r.returncode != 0 or not os.path.exists(out):
            raise SystemExit(f"lbvc: appimagetool failed:\n{r.stdout}{r.stderr}")
    return out


# ---------------------------------------------------------------- building it
def build_package(vbp, target, folder, fmt, files, csrc, title):
    """the folder <folder>/<program>-<target>/ and its archive (an AppImage: the folder has the .AppImage, the
    files the user added and the licenses); -> (folder of the program, archive or None, report)"""
    exe = files[0].source
    name = os.path.splitext(os.path.basename(exe))[0]
    os.makedirs(folder, exist_ok=True)
    app = os.path.join(folder, f'{name}-{target}')
    shutil.rmtree(app, ignore_errors=True)
    os.makedirs(app)
    missing = [f for f in files if f.checked and f.data is None and not os.path.exists(f.source)]
    if missing:
        raise SystemExit('lbvc: not found: ' + ', '.join(f.source for f in missing))
    for f in files:
        if not f.checked:
            continue
        dst = os.path.join(app, f.name.replace('/', os.sep))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if f.data is not None:
            with open(dst, 'wb') as o:
                o.write(f.data)
        else:
            shutil.copy2(f.source, dst)         # a symbolic link (libQt6Core.so.6): the file it names
    qt_libs = [f.name[4:] for f in files if f.checked and f.kind == 'qt' and f.name.startswith('lib/')]
    with open(os.path.join(app, 'licenses', LICENSE_NOTE), 'w', encoding='utf-8', newline='\r\n' if target != 'linux' else '\n') as f:
        f.write(license_note(csrc, target, title, qt_libs, fmt == 'appimage'))
    archive = None
    if fmt == 'appimage':
        from .driver import uses_forms
        added = [f.name for f in files if f.checked and f.kind == 'added']
        staged = app + '.files'
        shutil.rmtree(staged, ignore_errors=True)
        os.rename(app, staged)
        os.makedirs(app)
        try:
            for a in added:
                shutil.copy2(os.path.join(staged, a), os.path.join(app, a))
            shutil.copytree(os.path.join(staged, 'licenses'), os.path.join(app, 'licenses'))
            archive = make_appimage(staged, added, app, name, title,
                                    project_icon(vbp), not uses_forms(csrc))
        finally:
            shutil.rmtree(staged, ignore_errors=True)
    if fmt == 'zip':
        archive = os.path.join(folder, f'{name}-{target}.zip')
        with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(app):
                for n in sorted(names):
                    p = os.path.join(root, n)
                    z.write(p, os.path.relpath(p, folder))
    elif fmt == 'tar.gz':
        archive = os.path.join(folder, f'{name}-{target}.tar.gz')
        with tarfile.open(archive, 'w:gz') as t:
            t.add(app, arcname=os.path.basename(app))
    return app, archive, report(vbp, target, app, archive, files, csrc)


def report(vbp, target, app, archive, files, csrc):
    """VB6 "Packaging Report" (pd1-21)"""
    from .driver import uses_forms
    exe = os.path.basename(files[0].source)
    lines = []
    if archive and archive.endswith('.AppImage'):
        lines.append(f"The AppImage of your application has been built as '{archive}'.")
        lines.append('')
        lines.append('Copy it, with the files next to it, anywhere on the other computer and run it; its App.Path is '
                     'the folder it is in.')
    elif archive:
        lines.append(f"The portable package of your application has been built as '{archive}'.")
        lines.append('')
        lines.append(f"The folder '{app}' has the same files, unpacked.")
    else:
        lines.append(f"The portable package of your application has been built in the folder '{app}'.")
    lines.append('')
    if archive and archive.endswith('.AppImage'):
        pass
    elif target == 'linux':
        lines.append(f'Copy the folder (or unpack the archive) anywhere on the other computer and run {exe} in it.')
        if uses_forms(csrc) and not any(f.checked and f.kind == 'qt' and f.name.startswith('lib/') for f in files):
            lines.append('The program needs Qt 6 on the computer that runs it (Ubuntu: sudo apt install libqt6widgets6).')
    else:
        lines.append(f'Copy the folder (or unpack the archive) anywhere on the other computer and run {exe} in it; '
                     'nothing has to be installed.')
    n = sum(1 for f in files if f.checked)
    lines.append('')
    lines.append(f'Files: {n}')
    for f in files:
        if f.checked:
            lines.append(f'  {f.name}')
    skipped = [f for f in files if not f.checked]
    if skipped:
        lines.append('Left out:')
        lines += [f'  {f.name}' for f in skipped]
    return '\n'.join(lines) + '\n'


def run(vbp, target, folder=None, fmt=None, add=(), remove=(), recompile='auto', script=None, codepage='cp1252',
        log=print):
    """lbvc --package: -> the report. recompile: auto (VB6's question answered Yes when a source is newer than
    the program, or there is none), yes, no"""
    from .driver import compile_project, build_exe, find_runtime, COMPONENT_TYPES
    vbp = os.path.abspath(vbp)
    project = load_project(vbp, codepage)
    if project.type.lower() in ('oledll', 'control'):
        raise SystemExit(f"lbvc: {COMPONENT_TYPES[project.type.lower()]} has no program of its own to package")
    pdir = os.path.dirname(vbp)
    settings = Settings.load(vbp)
    script = script or settings.recent or DEFAULT_SCRIPT
    old = settings.get(script)
    fmt = fmt or old.get('Format') or default_format(target)
    if fmt not in FORMATS:
        raise SystemExit(f"lbvc: --format {fmt}: folder, zip, tar.gz or appimage")
    if fmt == 'appimage' and target != 'linux':
        raise SystemExit("lbvc: --format appimage: a program for Linux (--target linux)")
    folder = folder or old.get('PackageFolder') or 'Package'
    folder = folder if os.path.isabs(folder) else os.path.join(pdir, folder)
    added = list(old.get('added', [])) + [rel_to(a, pdir) for a in add if rel_to(a, pdir) not in old.get('added', [])]
    removed = list(old.get('removed', [])) + [r for r in remove if r not in old.get('removed', [])]
    exe = exe_path(vbp, target, project)
    csrc, _ = compile_project([vbp], codepage)
    stale = stale_sources(vbp, exe)
    made_for = program_target(exe)
    if made_for not in (None, target):          # a program for another system is never packaged (E8)
        if recompile == 'no':
            raise SystemExit(f"lbvc: {exe} is a program for {made_for}, not {target}; make it again "
                             f"(--recompile yes)")
        recompile = 'yes'
    if recompile == 'yes' or (recompile == 'auto' and stale):
        if stale and os.path.exists(exe):
            log("lbvc: newer than the program: " + ', '.join(os.path.relpath(s, pdir) for s in stale))
        log(f"lbvc: making {exe}")
        build_exe(csrc, exe, find_runtime(), target=target)
    elif not os.path.exists(exe):
        raise SystemExit(f"lbvc: {exe} not found (make the program first, or --recompile yes)")
    if target == 'linux':
        from .driver import uses_forms, PACKAGE_LIBDIR
        with open(exe, 'rb') as f:
            if uses_forms(csrc) and PACKAGE_LIBDIR.encode() not in f.read():
                raise SystemExit(f"lbvc: {exe} was made by an older LinBASIC Studio: it would not find the Qt of "
                                 f"its package; make it again (--recompile yes)")
    files = package_files(vbp, target, csrc, exe, added, removed, fmt)
    title = project.get('Title') or project.name
    app, archive, rep = build_package(vbp, target, folder, fmt, files, csrc, title)
    settings.put(script, {'Target': target, 'Format': fmt, 'PackageFolder': rel_to(folder, pdir),
                          'added': added, 'removed': removed})
    settings.save()
    return rep

