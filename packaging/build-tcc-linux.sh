#!/bin/sh
# build-tcc-linux.sh - the Tiny C Compiler (LGPL-2.1) the Linux AppImage of LinBASIC Studio carries,
# so that the programs are built on a computer without gcc (Phase 5): runtime/tcc-linux/
#   tcc                 the compiler (x86_64 Linux), its own headers in include/, lib/libtcc1.a
#   crt/                crt1.o crti.o crtn.o libc_nonshared.a of glibc (LGPL-2.1 with the linking
#                       exception) and small linker scripts for libc, libm, libdl, libstdc++:
#                       what libc6-dev / libstdc++-dev would give, pointing at the libraries every
#                       Linux has (libc.so.6, libm.so.6, libstdc++.so.6, libgcc_s.so.1)
# lbvc uses it when LBVC_LINUX_TCC names this folder (the AppRun of the AppImage sets it). The
# generated C includes only vbrt.h, so no system header is needed.
#   packaging/build-tcc-linux.sh [commit]      (on Ubuntu 24.04 x86_64, with gcc and libc6-dev)
set -e
here="$(cd "$(dirname "$0")" && pwd)"
out="$(dirname "$here")/runtime/tcc-linux"
work="${TMPDIR:-/tmp}/lb-tccl.$$"
mkdir -p "$work"
trap 'rm -rf "$work"' EXIT
git clone -q https://github.com/TinyCC/tinycc "$work/src"
cd "$work/src"
[ -n "$1" ] && git checkout -q "$1"
rev=$(git rev-parse --short HEAD)
./configure --crtprefix='{B}/crt' \
    --libpaths='{B}/lib:{B}/crt:/lib/x86_64-linux-gnu:/usr/lib/x86_64-linux-gnu' \
    --sysincludepaths='{B}/include:/usr/include/x86_64-linux-gnu:/usr/include' >/dev/null
make -j"$(nproc)" >/dev/null
rm -rf "$out"
mkdir -p "$out/include" "$out/lib" "$out/crt"
cp tcc COPYING "$out/"
strip "$out/tcc"
cp include/*.h "$out/include/"
cp libtcc1.a "$out/lib/"
L=/usr/lib/x86_64-linux-gnu
cp $L/crt1.o $L/crti.o $L/crtn.o $L/libc_nonshared.a "$out/crt/"
printf 'GROUP ( /lib/x86_64-linux-gnu/libc.so.6 libc_nonshared.a AS_NEEDED ( /lib64/ld-linux-x86-64.so.2 ) )\n' > "$out/crt/libc.so"
printf 'GROUP ( /lib/x86_64-linux-gnu/libm.so.6 )\n' > "$out/crt/libm.so"
printf 'GROUP ( /lib/x86_64-linux-gnu/libc.so.6 )\n' > "$out/crt/libdl.so"
printf 'GROUP ( /lib/x86_64-linux-gnu/libstdc++.so.6 /lib/x86_64-linux-gnu/libgcc_s.so.1 )\n' > "$out/crt/libstdc++.so"
echo "tinycc $rev (glibc $(ldd --version | head -1 | awk '{print $NF}'))" > "$out/VERSION"
echo "runtime/tcc-linux: tinycc $rev"
