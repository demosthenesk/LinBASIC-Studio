#!/bin/sh
# build-win.sh - libvbrt.a for Windows with mingw-w64 (on Linux: sudo apt install mingw-w64)
#   ./build-win.sh win64      runtime/win64/libvbrt.a
#   ./build-win.sh win32      runtime/win32/libvbrt.a
cd "$(dirname "$0")" || exit 1
case "$1" in
    win64) T=x86_64-w64-mingw32 ;;
    win32) T=i686-w64-mingw32 ;;
    *) echo "usage: $0 win32|win64"; exit 2 ;;
esac
CC=$(command -v $T-gcc-posix || command -v $T-gcc) || { echo "$T-gcc not found (sudo apt install mingw-w64)"; exit 1; }
python3 mkerrors.py >/dev/null || exit 1
OUT=$1
mkdir -p "$OUT/obj"
SRCS="core variant decimal intl array io lib object binfile dll varray format settings resource os scripting component mci mci_win serial net inet stdfmt lbmapi mapi addin debug dbgpipe_win"
for f in $SRCS; do $CC -O2 -g0 -Wall -Wno-unused-result -Wno-unused-function -Wno-unused-variable -c $f.c -o "$OUT/obj/$f.o" || exit 1; done
rm -f "$OUT/libvbrt.a" && ${CC%gcc*}ar rcs "$OUT/libvbrt.a" "$OUT"/obj/*.o 2>/dev/null || $T-ar rcs "$OUT/libvbrt.a" "$OUT"/obj/*.o
rm -rf "$OUT/obj"
echo "runtime/$OUT/libvbrt.a built"
