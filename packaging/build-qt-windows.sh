#!/bin/sh
# build-qt-windows.sh - Qt for the Windows programs that File > Make writes: Qt 5.15 LTS (qtbase:
# Core, Gui, Widgets and the Windows plugins) with mingw-w64. Qt 5.15, not Qt 6: Qt 6 needs
# Windows 10, Qt 5.15 runs on Windows 7, 8, 10 and 11, 32-bit and 64-bit.
# Qt is LGPL-3.0: its DLLs are shipped unmodified next to the programs.
#   packaging/build-qt-windows.sh win64|win32 [work folder]
# Needs: mingw-w64, perl, git, make
# Result: runtime/<target>/qt/{bin,lib,include,plugins}, then runtime/<target>/libvbrt.a and libvbforms.a
set -e
HERE=$(cd "$(dirname "$0")/.." && pwd)
A=$1
case "$A" in
    win64) T=x86_64-w64-mingw32 ;;
    win32) T=i686-w64-mingw32 ;;
    *) echo "usage: $0 win32|win64 [work]"; exit 2 ;;
esac
W=${2:-$HOME/.cache/linbasic-studio/qt-build}
QTV=5.15.19-lts-lgpl
mkdir -p "$W/mingwbin"
[ -d "$W/qtbase-$QTV" ] || git clone --depth 1 --branch "v$QTV" https://github.com/qt/qtbase.git "$W/qtbase-$QTV"
# the posix-thread variant of mingw-w64 under the plain names that qmake calls
for a in x86_64 i686; do
    for t in gcc g++; do ln -sf "$(command -v $a-w64-mingw32-$t-posix)" "$W/mingwbin/$a-w64-mingw32-$t"; done
    for t in ar ranlib strip windres objcopy nm ld objdump dlltool; do ln -sf "$(command -v $a-w64-mingw32-$t)" "$W/mingwbin/$a-w64-mingw32-$t"; done
done
B="$W/build-$A"
rm -rf "$B" && mkdir -p "$B" && cd "$B"
"$W/qtbase-$QTV/configure" -xplatform win32-g++ -device-option CROSS_COMPILE="$W/mingwbin/$T-" \
  -prefix "$W/qt-$A" -opensource -confirm-license -release -shared \
  -nomake examples -nomake tests -no-dbus -no-icu -opengl desktop \
  -qt-zlib -qt-libpng -qt-libjpeg -qt-pcre -qt-harfbuzz -qt-freetype \
  -no-sql-sqlite -no-sql-odbc -no-sql-psql -no-sql-mysql
make sub-src-qmake_all
make -j"$(nproc)" -C src sub-widgets
make -j"$(nproc)" -C src sub-platformsupport
make -j"$(nproc)" -C src/plugins/platforms sub-windows
make -j"$(nproc)" -C src sub-3rdparty-libjpeg-make_first
for p in imageformats styles; do make -j"$(nproc)" -C src/plugins sub-$p; done
for d in src/corelib src/gui src/widgets src/plugins/platforms/windows src/plugins/imageformats src/plugins/styles; do make -C "$d" install; done
# the kit: what the programs need (DLLs, plugins) and what libvbforms needs (headers, import libraries)
K="$HERE/runtime/$A/qt"
rm -rf "$K" && mkdir -p "$K/bin" "$K/lib" "$K/plugins"
cp "$W/qt-$A"/bin/Qt5Core.dll "$W/qt-$A"/bin/Qt5Gui.dll "$W/qt-$A"/bin/Qt5Widgets.dll "$K/bin/"
cp "$W/qt-$A"/lib/libQt5Core.a "$W/qt-$A"/lib/libQt5Gui.a "$W/qt-$A"/lib/libQt5Widgets.a "$K/lib/"
cp -r "$W/qt-$A/include" "$K/"
for d in platforms imageformats styles; do [ -d "$W/qt-$A/plugins/$d" ] && cp -r "$W/qt-$A/plugins/$d" "$K/plugins/"; done
for n in libstdc++-6.dll libwinpthread-1.dll libgcc_s_seh-1.dll libgcc_s_dw2-1.dll; do
    p=$($T-gcc-posix -print-file-name=$n); [ -f "$p" ] && cp "$p" "$K/bin/"
done
$T-strip "$K"/bin/*.dll "$K"/plugins/*/*.dll 2>/dev/null || true
sh "$HERE/runtime/build-win.sh" "$A"
sh "$HERE/runtime/forms/build-win.sh" "$A"
echo "runtime/$A ready"
