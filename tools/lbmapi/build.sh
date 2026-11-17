#!/bin/sh
# lbmapi.dll (32-bit, for VB6 in the VM): LinBASIC's provider of Simple MAPI as a mail program of Windows.
#   sh tools/lbmapi/build.sh <output folder>
set -e
HERE=$(cd "$(dirname "$0")" && pwd)
OUT=${1:-$HERE}
CC=$(command -v i686-w64-mingw32-gcc-posix || command -v i686-w64-mingw32-gcc) || { echo "i686-w64-mingw32-gcc not found (sudo apt install mingw-w64)"; exit 1; }
mkdir -p "$OUT"
"$CC" -O2 -Wall -shared -o "$OUT/lbmapi.dll" "$HERE/lbmapi_dll.c" "$HERE/../../runtime/lbmapi.c" \
    -Wl,--kill-at -static-libgcc
cp -f "$HERE/install.bat" "$HERE/uninstall.bat" "$OUT/"
echo "$OUT/lbmapi.dll built"
