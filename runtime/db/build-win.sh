#!/bin/sh
# build-win.sh - runtime/<target>/libvbdb.a for Windows with mingw-w64: DAO / ADO over SQLite, the
# Jet .mdb files read and written (jetread.c, jetwrite.c)
#   ./build-win.sh win64 | win32
cd "$(dirname "$0")" || exit 1
case "$1" in
    win64) T=x86_64-w64-mingw32 ;;
    win32) T=i686-w64-mingw32 ;;
    *) echo "usage: $0 win32|win64"; exit 2 ;;
esac
CC=$(command -v $T-gcc-posix || command -v $T-gcc) || { echo "$T-gcc not found (sudo apt install mingw-w64)"; exit 1; }
OUT=../$1/libvbdb.a
OBJ=obj-$1
mkdir -p $OBJ
SQLITE_FLAGS="-DSQLITE_THREADSAFE=0 -DSQLITE_ENABLE_COLUMN_METADATA -DSQLITE_DEFAULT_FOREIGN_KEYS=0 -DSQLITE_OMIT_LOAD_EXTENSION -DSQLITE_DQS=0 -DSQLITE_MAX_ATTACHED=30"
# (made again when the flags change: their copy in sqlite3.flags)
if [ ! -f $OBJ/sqlite3.o ] || [ sqlite/sqlite3.c -nt $OBJ/sqlite3.o ] || [ "$(cat $OBJ/sqlite3.flags 2>/dev/null)" != "$SQLITE_FLAGS" ]; then
    $CC -O2 -w $SQLITE_FLAGS -c sqlite/sqlite3.c -o $OBJ/sqlite3.o || exit 1
    echo "$SQLITE_FLAGS" > $OBJ/sqlite3.flags
fi
rm -f $OBJ/mdb_*.o                              # (libmdb of mdbtools: no longer used)
for f in jetread jetwrite jetsort jetprops jetquery jetnew jetcp; do
    $CC -O2 -g0 -w -I.. -c $f.c -o $OBJ/$f.o || exit 1
done
for f in jetsql dao daoschema daors mdbimport mdbexport ado bind odbc; do
    $CC -O2 -g0 -w $SQLITE_FLAGS -I.. -c $f.c -o $OBJ/$f.o || exit 1
done
rm -f $OUT && $T-ar rcs $OUT $OBJ/*.o
echo "runtime/$1/libvbdb.a built"
# the tool of the IDE of Windows (the tables and fields of a database at design time)
if [ -f ../$1/libvbrt.a ]; then
    $CC -O2 -w $SQLITE_FLAGS -I.. lbdbinfo.c $OUT ../$1/libvbrt.a -lwinmm -lodbc32 -lole32 -loleaut32 -static -s -o ../$1/lbdbinfo.exe || exit 1
    echo "runtime/$1/lbdbinfo.exe built"
fi
