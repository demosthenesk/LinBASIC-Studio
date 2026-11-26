#!/bin/sh
# build-win.sh - libvbforms.a for Windows (mingw-w64 + Qt 6 for Windows in runtime/<target>/qt)
#   ./build-win.sh win64 | win32
cd "$(dirname "$0")" || exit 1
case "$1" in
    win64) T=x86_64-w64-mingw32 ;;
    win32) T=i686-w64-mingw32 ;;
    *) echo "usage: $0 win32|win64"; exit 2 ;;
esac
CXX=$(command -v $T-g++-posix || command -v $T-g++) || { echo "$T-g++ not found (sudo apt install mingw-w64)"; exit 1; }
QT=../$1/qt
[ -d "$QT/include/QtWidgets" ] || { echo "Qt for $1 not found in runtime/$1/qt (packaging/build-qt-windows.sh)"; exit 1; }
I="-I$QT/include -I$QT/include/QtCore -I$QT/include/QtGui -I$QT/include/QtWidgets"
CC=$(command -v $T-gcc-posix || command -v $T-gcc) || { echo "$T-gcc not found (sudo apt install mingw-w64)"; exit 1; }
$CXX -O2 -g0 -std=c++17 -Wall -DQT_NO_DEBUG -DUNICODE -D_UNICODE -D_USE_MATH_DEFINES $I -c vbforms.cpp -o vbforms-$1.o || exit 1
# CHMLib (chm/, LGPL): the .chm help files of the programs (help.inc), its names lbchm_...
for f in chm_lib lzx; do
    $CC -O2 -g0 -w -DWIN32 -include chm/lbchm.h -c chm/$f.c -o chm-$f-$1.o || exit 1
done
rm -f ../$1/libvbforms.a && $T-ar rcs ../$1/libvbforms.a vbforms-$1.o chm-chm_lib-$1.o chm-lzx-$1.o \
    && rm -f vbforms-$1.o chm-chm_lib-$1.o chm-lzx-$1.o && echo "runtime/$1/libvbforms.a built"
