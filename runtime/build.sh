#!/bin/sh
# build.sh - builds libvbrt.a, the runtime of programs compiled by lbvc
#   ./build.sh          runtime/libvbrt.a (gcc)
#   ./build.sh clean
cd "$(dirname "$0")" || exit 1
if [ "$1" = clean ]; then rm -f ./*.o libvbrt.a errors.inc; exit 0; fi
python3 mkerrors.py >/dev/null || exit 1
CC=${CC:-gcc}
CFLAGS="-O2 -g0 -fPIC -Wall -Wno-unused-result"
SRCS="core variant decimal intl array io lib object binfile dll varray format debug settings resource os scripting component mci serial net inet stdfmt lbmapi mapi addin"
for f in $SRCS; do $CC $CFLAGS -c $f.c -o $f.o || exit 1; done
rm -f libvbrt.a && ar rcs libvbrt.a $(for f in $SRCS; do printf '%s.o ' $f; done) && rm -f ./*.o && echo "runtime/libvbrt.a built"
