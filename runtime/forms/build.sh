#!/bin/sh
# build.sh - builds libvbforms.a, the forms runtime (Qt6) of programs compiled by lbvc
cd "$(dirname "$0")" || exit 1
CXX=${CXX:-g++}
CC=${CC:-gcc}
$CXX -O2 -g0 -fPIC -std=c++17 -Wall $(pkg-config --cflags ${LBVC_QT_PKG:-Qt6Widgets}) -c vbforms.cpp -o vbforms.o || exit 1
# CHMLib (chm/, LGPL): the .chm help files of the programs (help.inc), its names lbchm_...
for f in chm_lib lzx; do
    $CC -O2 -g0 -fPIC -w -include chm/lbchm.h -c chm/$f.c -o chm-$f.o || exit 1
done
rm -f libvbforms.a && ar rcs libvbforms.a vbforms.o chm-chm_lib.o chm-lzx.o && rm -f vbforms.o chm-chm_lib.o chm-lzx.o \
    && echo "runtime/forms/libvbforms.a built"
# libvbforms.so: the same runtime as a shared library, for the programs the bundled Tiny C Compiler
# of the AppImage builds (tcc links C, not the C++ of libvbforms.a); it goes next to such a program
$CXX -shared -o libvbforms.so -Wl,--whole-archive libvbforms.a -Wl,--no-whole-archive \
    $(pkg-config --libs ${LBVC_QT_PKG:-Qt6Widgets}) -s || exit 1
echo "runtime/forms/libvbforms.so built"
