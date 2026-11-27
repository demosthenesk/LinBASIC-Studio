#!/bin/sh
# build.sh - runtime/libvbdb.a: DAO / ADO over SQLite, Jet .mdb import (libmdb of mdbtools, LGPL);
# linked into the programs that use a database
cd "$(dirname "$0")" || exit 1
CC=${CC:-gcc}
OUT=../libvbdb.a
mkdir -p obj
SQLITE_FLAGS="-DSQLITE_THREADSAFE=0 -DSQLITE_ENABLE_COLUMN_METADATA -DSQLITE_DEFAULT_FOREIGN_KEYS=0 -DSQLITE_OMIT_LOAD_EXTENSION -DSQLITE_DQS=0 -DSQLITE_MAX_ATTACHED=30"
# (made again when the flags change: their copy in sqlite3.flags)
if [ ! -f obj/sqlite3.o ] || [ sqlite/sqlite3.c -nt obj/sqlite3.o ] || [ "$(cat obj/sqlite3.flags 2>/dev/null)" != "$SQLITE_FLAGS" ]; then
    $CC -O2 -fPIC -w $SQLITE_FLAGS -c sqlite/sqlite3.c -o obj/sqlite3.o || exit 1
    echo "$SQLITE_FLAGS" > obj/sqlite3.flags
fi
MDB_FLAGS="-DTLS=_Thread_local -DHAVE_ICONV=1 -DICONV_CONST= -DHAVE_GMTIME_R=1 -DHAVE_STRPTIME=1 -DHAVE_VASPRINTF=1 -D_GNU_SOURCE"
for f in mdb/*.c; do
    o=obj/mdb_$(basename $f .c).o
    if [ ! -f $o ] || [ $f -nt $o ] || [ mdb/mdbtools.h -nt $o ]; then
        $CC -O2 -fPIC -w $MDB_FLAGS -Imdb -c $f -o $o || exit 1
    fi
done
# libmdb's own GLib (mdb/fakeglib.c) gets names of its own (lbmdb_g_...): a program with forms has the real
# GLib too (Qt, the GTK theme) and must not call these instead (a program built by tcc exports its symbols)
# the same for SQLite (lbsq_sqlite3_...): GTK's file dialog and other libraries load the system's libsqlite3
nm --defined-only -g obj/sqlite3.o | awk '$3 ~ /^(lbsq_)?sqlite3/ { n = $3; sub(/^lbsq_/, "", n); print n, "lbsq_" n }' > obj/sqlite.map
objcopy --redefine-syms=obj/sqlite.map obj/sqlite3.o || exit 1
nm --defined-only -g obj/mdb_fakeglib.o | awk '$3 ~ /^(lbmdb_)?g_/ { n = $3; sub(/^lbmdb_/, "", n); print n, "lbmdb_" n }' > obj/glib.map
for o in obj/mdb_*.o; do objcopy --redefine-syms=obj/glib.map "$o" || exit 1; done
for f in jetsql dao daoschema daors mdbimport ado bind odbc; do
    $CC -O2 -g -fPIC -Wall -Wno-unused-function -Wno-unused-variable -Wno-format-truncation $SQLITE_FLAGS -I.. -Imdb -c $f.c -o obj/$f.o || exit 1
    objcopy --redefine-syms=obj/glib.map --redefine-syms=obj/sqlite.map obj/$f.o || exit 1
done
rm -f $OUT && ar rcs $OUT obj/*.o
echo "runtime/libvbdb.a built"
# the tool of the IDE (the tables and fields of a database at design time): a native build only
if [ -z "$LBDB_NO_TOOL" ] && [ -f ../libvbrt.a ]; then
    $CC -O2 -w $SQLITE_FLAGS -I.. -c lbdbinfo.c -o obj/lbdbinfo.tool || exit 1
    objcopy --redefine-syms=obj/sqlite.map obj/lbdbinfo.tool || exit 1
    $CC obj/lbdbinfo.tool $OUT ../libvbrt.a -lm -ldl -o lbdbinfo || exit 1
    rm -f obj/lbdbinfo.tool
fi
