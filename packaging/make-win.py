#!/usr/bin/env python3
"""make-win.py - the portable LinBASIC Studio for Windows (Phase 5): one folder that runs from
anywhere (a USB stick too) on Windows 7 and later, 32-bit or 64-bit.

    packaging/make-win.py win64 <python-3.8.10-embed-amd64.zip> <out folder> [--wheels DIR] [--version V]
    packaging/make-win.py win32 <python-3.8.10-embed-win32.zip> <out folder> [--wheels DIR] [--version V]

<out folder>/LinBASIC-Studio-<version>-<target>/
    LinBASIC Studio.exe      the launcher (packaging/win/launcher.c): python\\pythonw.exe -m lbstudio
    lbvc.cmd                 the compiler on the command line
    portable.ini             the settings stay in settings\\ (paths.py)
    python\\                  Python 3.8.10 embeddable (the last Python for Windows 7) + PyQt5 5.15,
                             QScintilla 2.14, pefile in Lib\\site-packages
    ide\\lbstudio, compiler\\, data\\vb6
    runtime\\vbrt.h, runtime\\<target>\\{dll, tcc, qt\\bin, qt\\plugins, lbdbinfo.exe}
                             what the programs are built with: the Tiny C Compiler (LGPL), the runtime
                             as DLLs (build-dll-win.sh), Qt 5 for the forms (build-qt-windows.sh)
    README.md, LICENSE, RUNTIME-EXCEPTION, licenses\\       the licenses of the third-party software (licenses/README.md)

The wheels: --wheels DIR (the .whl files), or pip downloads them (cp38, the target's platform).
The Python zip comes from python.org (python-3.8.10-embed-*.zip)."""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.dirname(HERE)
PLAT = {"win32": "win32", "win64": "win_amd64"}
MINGW = {"win32": "i686-w64-mingw32", "win64": "x86_64-w64-mingw32"}
WHEELS = ["PyQt5==5.15.10", "PyQt5-Qt5==5.15.2", "PyQt5-sip==12.15.0", "QScintilla==2.14.1", "pefile"]
# the Qt of PyQt5 the IDE does not use (QtCore, QtGui, QtWidgets, QtNetwork, QtPrintSupport, Qsci stay)
QT_DROP = ("Qt5Quick", "Qt5Qml", "Qt5WebSockets", "Qt5WebChannel", "Qt5Multimedia", "Qt5Bluetooth", "Qt5Nfc",
           "Qt5Location", "Qt5Positioning", "Qt5Sensors", "Qt5SerialPort", "Qt5RemoteObjects", "Qt5Sql",
           "Qt5Test", "Qt5Designer", "Qt5Help", "Qt5Xml", "Qt5XmlPatterns", "Qt5DBus", "Qt5OpenGL",
           "Qt5TextToSpeech", "Qt5WinExtras", "Qt5Svg", "Qt5WebView", "opengl32sw", "d3dcompiler_47", "libEGL",
           "libGLESv2", "libcrypto", "libssl", "libeay32", "ssleay32")
PYQT_KEEP = ("QtCore", "QtGui", "QtWidgets", "QtNetwork", "QtPrintSupport", "Qsci", "sip")
PLUGINS_KEEP = ("platforms", "imageformats", "styles", "printsupport")


def version():
    with open(os.path.join(SRC, "ide", "lbstudio", "__init__.py"), encoding="utf-8") as f:
        for line in f:
            if line.startswith("VERSION"):
                return line.split("=", 1)[1].strip().strip("\"'")
    return "0.0.0"


def copytree(src, dst):
    shutil.copytree(src, dst, ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".git*"))


def wheels(target, wdir, work):
    if wdir:
        files = glob.glob(os.path.join(wdir, "*.whl"))
    else:
        d = os.path.join(work, "wheels")
        subprocess.run([sys.executable, "-m", "pip", "download", "-q", "--only-binary=:all:", "--no-deps",
                        "--platform", PLAT[target], "--python-version", "38", "--implementation", "cp",
                        "-d", d] + WHEELS, check=True)
        files = glob.glob(os.path.join(d, "*.whl"))
    if not any(os.path.basename(f).lower().startswith("pefile") for f in files):
        d = os.path.join(work, "wheels-pure")
        subprocess.run([sys.executable, "-m", "pip", "download", "-q", "--no-deps", "-d", d, "pefile"], check=True)
        files += glob.glob(os.path.join(d, "*.whl"))
    return files


def trim_qt(site):
    q = os.path.join(site, "PyQt5")
    for f in os.listdir(q):                              # the bindings of the dropped libraries
        if f.endswith((".pyd", ".pyi")) and not f.split(".")[0] in PYQT_KEEP:
            os.remove(os.path.join(q, f))
    for sub in ("uic", "bindings", "lupdate", "pylupdate", "pyrcc"):
        shutil.rmtree(os.path.join(q, sub), ignore_errors=True)
    for f in glob.glob(os.path.join(q, "pylupdate*")) + glob.glob(os.path.join(q, "pyrcc*")):
        os.remove(f) if os.path.isfile(f) else shutil.rmtree(f)
    qt = os.path.join(q, "Qt5") if os.path.isdir(os.path.join(q, "Qt5")) else os.path.join(q, "Qt")
    b = os.path.join(qt, "bin")
    for f in os.listdir(b):
        if f.startswith(QT_DROP):
            os.remove(os.path.join(b, f))
    pl = os.path.join(qt, "plugins")
    for d in os.listdir(pl):
        if d not in PLUGINS_KEEP:
            shutil.rmtree(os.path.join(pl, d))
    for f in glob.glob(os.path.join(pl, "imageformats", "*")):
        if not os.path.basename(f).startswith(("qgif", "qico", "qjpeg")):
            os.remove(f)
    for d in ("qml", "translations", "resources", "qsci"):
        p = os.path.join(qt, d)
        if d == "translations" and os.path.isdir(p):
            for f in os.listdir(p):                       # Qt's own texts in Greek and English stay
                if not f.startswith(("qtbase_el", "qtbase_en", "qt_el", "qt_en")):
                    os.remove(os.path.join(p, f))
        elif d != "translations":
            shutil.rmtree(p, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("target", choices=["win32", "win64"])
    ap.add_argument("python_zip")
    ap.add_argument("out")
    ap.add_argument("--wheels")
    ap.add_argument("--version", default=version())
    a = ap.parse_args()
    t = a.target
    rt = os.path.join(SRC, "runtime", t)
    for need in ("dll/vbrt.dll", "dll/vbrt.def", "tcc/tcc.exe", "qt/bin", "lbdbinfo.exe"):
        if not os.path.exists(os.path.join(rt, need)):
            raise SystemExit(f"runtime/{t}/{need} is missing (packaging/build-tcc.sh, runtime/build-dll-win.sh {t}, "
                             f"runtime/db/build-win.sh {t})")
    name = f"LinBASIC-Studio-{a.version}-{t}"
    dest = os.path.join(os.path.abspath(a.out), name)
    if os.path.exists(dest):
        shutil.rmtree(dest)
    os.makedirs(dest)
    with tempfile.TemporaryDirectory() as work:
        # Python
        py = os.path.join(dest, "python")
        with zipfile.ZipFile(a.python_zip) as z:
            z.extractall(py)
        pth = glob.glob(os.path.join(py, "python3*._pth"))
        if not pth:
            raise SystemExit("not an embeddable Python zip (no python3*._pth)")
        zipname = os.path.basename(pth[0])[:-5] + ".zip"
        with open(pth[0], "w", newline="\r\n") as f:
            f.write(f"{zipname}\n.\nLib\\site-packages\n..\\ide\n..\\compiler\nimport site\n")
        site = os.path.join(py, "Lib", "site-packages")
        os.makedirs(site)
        for w in wheels(t, a.wheels, work):
            with zipfile.ZipFile(w) as z:
                z.extractall(site)
        trim_qt(site)
        # the IDE, the compiler, the data of VB6
        copytree(os.path.join(SRC, "ide", "lbstudio"), os.path.join(dest, "ide", "lbstudio"))
        copytree(os.path.join(SRC, "compiler", "lbvc"), os.path.join(dest, "compiler", "lbvc"))
        shutil.copy(os.path.join(SRC, "compiler", "lbvc.py"), os.path.join(dest, "compiler"))
        copytree(os.path.join(SRC, "data", "vb6"), os.path.join(dest, "data", "vb6"))
        # the runtime: what tcc builds the programs with
        r = os.path.join(dest, "runtime")
        os.makedirs(os.path.join(r, t))
        shutil.copy(os.path.join(SRC, "runtime", "vbrt.h"), r)
        shutil.copy(os.path.join(rt, "lbdbinfo.exe"), os.path.join(r, t))
        os.makedirs(os.path.join(r, t, "dll"))
        for f in glob.glob(os.path.join(rt, "dll", "*.dll")) + glob.glob(os.path.join(rt, "dll", "*.def")):
            shutil.copy(f, os.path.join(r, t, "dll"))
        shutil.copytree(os.path.join(rt, "tcc"), os.path.join(r, t, "tcc"), ignore=shutil.ignore_patterns("cross-tcc"))
        for sub in ("bin", "plugins"):
            shutil.copytree(os.path.join(rt, "qt", sub), os.path.join(r, t, "qt", sub))
        # the launcher, the command line, the settings
        cc = shutil.which(MINGW[t] + "-gcc")
        if cc is None:
            raise SystemExit(f"{MINGW[t]}-gcc not found (sudo apt install mingw-w64)")
        res = os.path.join(work, "launcher-res.o")         # the icon of the .exe
        subprocess.run([MINGW[t] + "-windres", os.path.join(HERE, "win", "launcher.rc"), "-o", res],
                       cwd=os.path.join(HERE, "win"), check=True)
        subprocess.run([cc, "-municode", "-mwindows", "-Os", "-s", os.path.join(HERE, "win", "launcher.c"), res,
                        "-o", os.path.join(dest, "LinBASIC Studio.exe")], check=True)
        with open(os.path.join(dest, "lbvc.cmd"), "w", newline="\r\n") as f:
            f.write('@"%~dp0python\\python.exe" "%~dp0compiler\\lbvc.py" %*\n')
        with open(os.path.join(dest, "portable.ini"), "w", newline="\r\n") as f:
            f.write("[LinBASIC Studio]\n; this file makes the folder portable: the settings are kept in settings\\\n"
                    "; (without it: %APPDATA%\\linbasic-studio)\n")
        for f in ("README.md", "LICENSE", "RUNTIME-EXCEPTION"):
            if os.path.exists(os.path.join(SRC, f)):
                shutil.copy(os.path.join(SRC, f), dest)
        shutil.copytree(os.path.join(SRC, "licenses"), os.path.join(dest, "licenses"))   # the third parties
        shutil.copy(os.path.join(SRC, "runtime", "db", "mdb", "COPYING.LIB"),
                    os.path.join(dest, "licenses", "libmdb-COPYING.LIB"))
    size = sum(os.path.getsize(os.path.join(d, f)) for d, _, fs in os.walk(dest) for f in fs)
    print(f"{dest}  ({size / 1048576:.0f} MB)")


if __name__ == "__main__":
    main()
