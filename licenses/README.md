# Third-party software in LinBASIC Studio

LinBASIC Studio itself is GPL-3.0-or-later (`LICENSE` at the top of the folder); its runtime
(libvbrt, libvbforms, libvbdb) has in addition the LinBASIC Studio Runtime Library Exception
(`RUNTIME-EXCEPTION`). It contains, or its
packages carry, the following software of other authors, each under its own license. The full
texts are in this folder; GPL-3.0 is `LICENSE`.

No Microsoft file is part of LinBASIC Studio: not VB6, not its runtime, not MSDN, not the `.mdb`
samples. The IDE opens the user's own MSDN help when the user owns it.

## In the source and in every package

| Software | Where | License | Text |
|---|---|---|---|
| SQLite 3 | `runtime/db/sqlite`, inside `libvbdb` / `vbdb.dll` | public domain | `SQLite-public-domain.txt` |
| CHMLib (Jed Wing), the .chm help of the programs | `runtime/forms/chm`, inside `libvbforms` / `vbforms.dll` | LGPL-2.1-or-later | `LGPL-2.1.txt` |
| Parts of Linpad++ (editor, styles, languages) | `ide/lbstudio` | GPL-3.0 (same author) | `LICENSE` |

## Linux AppImage

| Software | Where in the AppImage | License | Text |
|---|---|---|---|
| Python 3.12 | `usr/bin`, `usr/lib/python3.12` | PSF-2.0 | `Python-PSF-2.0.txt` |
| PyQt6, PyQt6-QScintilla (Riverbank Computing) | `usr/share/linbasic-studio/lib/PyQt6` | GPL-3.0 | `LICENSE` |
| PyQt6-sip | same | BSD-2-Clause | `sip-BSD-2-Clause.txt` |
| Qt 6 (The Qt Company), from the PyQt6-Qt6 wheel | `.../PyQt6/Qt6/lib`, `plugins` | LGPL-3.0 | `LGPL-3.0.txt` |
| ICU, inside that Qt | `.../PyQt6/Qt6/lib/libicu*` | Unicode (MIT-style) | `ICU-Unicode.txt` |
| pefile | `.../lib/pefile.py` | MIT | `pefile-MIT.txt` |
| Tiny C Compiler (tinycc) | `usr/share/linbasic-studio/runtime/tcc-linux` | LGPL-2.1 | `runtime/tcc-linux/COPYING`, `LGPL-2.1.txt` |
| crt1.o, crti.o, crtn.o, libc_nonshared.a of glibc | `runtime/tcc-linux/crt` | LGPL-2.1 with the linking exception | `LGPL-2.1.txt` |
| Tiny C Compiler (tinycc) for Windows, as a cross compiler (the Windows programs) | `usr/share/linbasic-studio/runtime/<target>/tcc` | LGPL-2.1 | `runtime/<target>/tcc/COPYING`, `LGPL-2.1.txt` |
| Qt 5.15 LTS (qtbase) for Windows, built with mingw-w64 (the forms of the Windows programs) | `usr/share/linbasic-studio/runtime/<target>/qt` | LGPL-3.0 | `LGPL-3.0.txt` |
| libstdc++, libgcc, winpthreads of mingw-w64 (the DLLs Qt for Windows needs) | `usr/share/linbasic-studio/runtime/<target>/qt/bin` | GPL-3.0 with the GCC Runtime Library Exception 3.1; MIT | `LICENSE`, `GCC-Runtime-Library-Exception-3.1.txt`, `winpthreads-MIT.txt` |
| appimagetool and the type2 runtime of the AppImage project (the AppImages of your programs) | `usr/share/linbasic-studio/runtime/appimage` | MIT (the runtime contains libfuse, LGPL-2.1, squashfuse, BSD-2-Clause, and zstd, BSD) | `runtime/appimage/*.LICENSE`, `LGPL-2.1.txt` |
| System libraries of Ubuntu 24.04 that Python and Qt need | `usr/lib` | the licenses of their Ubuntu packages, listed in `usr/lib/LIBRARIES.txt` | — |

## Windows packages (win32, win64)

| Software | Where in the folder | License | Text |
|---|---|---|---|
| Python 3.8.10 (embeddable) | `python` | PSF-2.0 | `python/LICENSE.txt`, `Python-PSF-2.0.txt` |
| PyQt5, QScintilla (Riverbank Computing) | `python/Lib/site-packages/PyQt5` | GPL-3.0 | `LICENSE` |
| PyQt5-sip | same | BSD-2-Clause | `sip-BSD-2-Clause.txt` |
| Qt 5.15 of PyQt5 (the IDE), from the PyQt5-Qt5 wheel | `python/Lib/site-packages/PyQt5/Qt5` | LGPL-3.0 | `LGPL-3.0.txt`, its `dist-info/LICENSE` |
| pefile | `python/Lib/site-packages/pefile.py` | MIT | `pefile-MIT.txt` |
| Tiny C Compiler (tinycc) | `runtime/<target>/tcc` | LGPL-2.1 | `runtime/<target>/tcc/COPYING`, `LGPL-2.1.txt` |
| Qt 5.15 LTS (qtbase), built with mingw-w64 (the programs' forms) | `runtime/<target>/qt` | LGPL-3.0 | `LGPL-3.0.txt` |
| libstdc++, libgcc of GCC (mingw-w64) | `runtime/<target>/qt/bin` | GPL-3.0 with the GCC Runtime Library Exception 3.1 | `LICENSE`, `GCC-Runtime-Library-Exception-3.1.txt` |
| winpthreads of mingw-w64 | `runtime/<target>/qt/bin/libwinpthread-1.dll` | MIT | `winpthreads-MIT.txt` |

## The programs you make

A program made with LinBASIC Studio contains its runtime (`libvbrt`, `libvbforms` for forms,
`libvbdb` for DAO / ADO). The runtime is part of LinBASIC Studio (`LICENSE`) and has the LinBASIC
Studio Runtime Library Exception (`RUNTIME-EXCEPTION`): you may give your program to others under
terms of your choice, open or closed source; only a changed runtime must be given with its source
(GPL). Depending on what it uses, the program also contains these libraries of other authors:

* forms: Qt (LGPL-3.0), loaded as shared libraries (`.so`, `.dll`) that can be replaced;
* the help of a program with forms (.chm files: F1, What's This): CHMLib (LGPL-2.1), linked in. If
  you give such a program to others, give them, or offer, the C source of your program (`lbvc -C`)
  or its object files, so that it can be linked again with a changed CHMLib;
* DAO / ADO: SQLite (public domain), linked in. The Jet .mdb files are read and written by
  LinBASIC Studio's own code (`runtime/db/jet*.c`, part of the runtime);
* Windows: the DLLs of the mingw-w64 C++ runtime that Qt needs (GCC Runtime Library Exception,
  winpthreads MIT);
* the portable package for Linux (`lbvc --package`, Package and Deployment Wizard): Qt 6 and the
  libraries it needs in `lib/` and `plugins/` (Qt LGPL-3.0, ICU Unicode license, the xcb / xkbcommon
  libraries MIT/X11, double-conversion and pcre2 BSD, md4c MIT, libb2 CC0); an AppImage also the
  runtime of the AppImage project (MIT, with libfuse LGPL-2.1). `lbvc --package` puts the licenses into
  the package (`licenses/`, `licenses/LICENSES.txt`).

The source of the LGPL libraries: Qt https://download.qt.io/archive/qt/ (5.15 for Windows,
`packaging/build-qt-windows.sh` builds it), Tiny C Compiler https://github.com/TinyCC/tinycc
(`packaging/build-tcc.sh`, `packaging/build-tcc-linux.sh`), CHMLib (the files used are in
`runtime/forms/chm`).
