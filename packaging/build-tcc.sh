#!/bin/sh
# build-tcc.sh - the Tiny C Compiler (LGPL-2.1) of LinBASIC Studio for Windows (Phase 5):
#   runtime/win32/tcc/, runtime/win64/tcc/   tcc.exe, libtcc.dll, include/, lib/ (the compiler the
#                                            IDE of Windows bundles: it builds the user's programs)
#   runtime/win32/tcc/cross-tcc ...          the same compiler running on Linux (lbvc uses it when
#                                            LBVC_WIN_CC=tcc, to test the Windows path here)
# from the source of tinycc (git, the "mob" branch, or a commit given as $1).
#   packaging/build-tcc.sh [commit]
set -e
here="$(cd "$(dirname "$0")" && pwd)"
rt="$(dirname "$here")/runtime"
work="${TMPDIR:-/tmp}/lb-tcc.$$"
mkdir -p "$work"
trap 'rm -rf "$work"' EXIT
git clone -q https://github.com/TinyCC/tinycc "$work/src"
[ -n "$1" ] && (cd "$work/src" && git checkout -q "$1")
rev=$(cd "$work/src" && git rev-parse --short HEAD)
# the cross compilers (they run here) and the libtcc1.a of each target
(cd "$work/src" && ./configure >/dev/null && make -j"$(nproc)" cross-i386-win32 cross-x86_64-win32 >/dev/null)
for a in i386 x86_64; do
    case $a in i386) k=win32; pre=i686-w64-mingw32- ;; x86_64) k=win64; pre=x86_64-w64-mingw32- ;; esac
    b="$work/$a"
    cp -r "$work/src" "$b"
    (cd "$b" && git clean -fdxq && ./configure --cross-prefix=$pre --targetos=WIN32 --cpu=$a >/dev/null &&
     gcc -DC2STR conftest.c -o c2str.exe && ./c2str.exe include/tccdefs.h tccdefs_.h &&
     make -j"$(nproc)" tcc.exe >/dev/null)
    out="$rt/$k/tcc"
    rm -rf "$out"
    mkdir -p "$out/include" "$out/lib"
    cp "$b/tcc.exe" "$b/libtcc.dll" "$out/"
    cp -r "$work/src/win32/include/." "$out/include/"
    cp "$work/src/include/"*.h "$out/include/"
    cp "$work/src/win32/lib/"*.def "$out/lib/"
    cp "$work/src/$a-win32-libtcc1.a" "$out/lib/libtcc1.a"
    cp "$work/src/$a-win32-libtcc1.a" "$out/lib/$a-win32-libtcc1.a"     # the name cross-tcc asks for
    cp "$work/src/$a-win32-tcc" "$out/cross-tcc"
    cp "$work/src/COPYING" "$out/COPYING"
    echo "tinycc $rev" > "$out/VERSION"
    echo "runtime/$k/tcc: tinycc $rev"
done
