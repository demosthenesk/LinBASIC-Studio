#!/bin/sh
# build-dll-win.sh - the runtime of LinBASIC Studio as DLLs for Windows, for the programs that the
# bundled Tiny C Compiler (tcc) builds on Windows (Phase 5): runtime/<target>/dll/
#   vbrt.dll     libvbrt (the language)            vbrt.def
#   vbdb.dll     libvbdb (DAO / ADO over SQLite)   vbdb.def
#   vbforms.dll  libvbforms (the forms, Qt 5)      vbforms.def
# The .def files are what tcc links with (tcc -impdef). The static libraries are built first
# (build-win.sh, db/build-win.sh, forms/build-win.sh).
#   ./build-dll-win.sh win64 | win32 [a tcc that runs here; default <target>/tcc/cross-tcc of
#                                    packaging/build-tcc.sh]
cd "$(dirname "$0")" || exit 1
case "$1" in
    win64) T=x86_64-w64-mingw32; TCC=${2:-$1/tcc/cross-tcc} ;;
    win32) T=i686-w64-mingw32; TCC=${2:-$1/tcc/cross-tcc} ;;
    *) echo "usage: $0 win32|win64 [tcc]"; exit 2 ;;
esac
CC=$(command -v $T-gcc-posix || command -v $T-gcc) || { echo "$T-gcc not found (sudo apt install mingw-w64)"; exit 1; }
CXX=$(command -v $T-g++-posix || command -v $T-g++)
K=$1
D=$K/dll
mkdir -p "$D"
for f in "$K/libvbrt.a" "$K/libvbdb.a" "$K/libvbforms.a"; do
    [ -f "$f" ] || { echo "$f is missing (build-win.sh, db/build-win.sh, forms/build-win.sh)"; exit 1; }
done
# vbrt.dll: everything of libvbrt; the C runtime of mingw inside (no libgcc DLL)
$CC -shared -o "$D/vbrt.dll" -Wl,--whole-archive "$K/libvbrt.a" -Wl,--no-whole-archive \
    -Wl,--export-all-symbols -static-libgcc -lwinmm -lws2_32 -lwininet -lm -s || exit 1
# vbdb.dll: SQLite, the .mdb import and DAO / ADO, on vbrt.dll
$CC -shared -o "$D/vbdb.dll" -Wl,--whole-archive "$K/libvbdb.a" -Wl,--no-whole-archive \
    "$D/vbrt.dll" -Wl,--export-all-symbols -static-libgcc -static -s || exit 1
# vbforms.dll: the forms on Qt 5 (its DLLs go next to the program), on vbrt.dll
Q=$K/qt/lib
$CXX -shared -o "$D/vbforms.dll" -Wl,--whole-archive "$K/libvbforms.a" -Wl,--no-whole-archive \
    "$D/vbrt.dll" -L"$Q" -lQt5Widgets -lQt5Gui -lQt5Core -lwinmm -lws2_32 -lwininet -Wl,--export-all-symbols \
    -static-libgcc -static-libstdc++ -static -s || exit 1
# what tcc links with
case "$TCC" in /*) ;; */*) TCC="$PWD/$TCC" ;; esac
if command -v "$TCC" >/dev/null 2>&1; then
    for n in vbrt vbdb vbforms; do
        (cd "$D" && "$TCC" -impdef $n.dll -o $n.def) || exit 1
    done
else
    echo "($TCC not found: no .def files; tcc -impdef makes them)"
fi
echo "runtime/$K/dll built"
