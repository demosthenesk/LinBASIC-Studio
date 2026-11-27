#!/bin/sh
# build32.sh - the SQLite ODBC driver of Christian Werner (sqliteodbc 0.99991: sqlite3odbc.c, sqlite3odbc.h as he
# distributes them, license.terms) built 32-bit, for the tests of the Windows programs in 32-bit Wine
# (LBVC_TARGET=win32 python3 tests/db/run_tests.py): Wine's 32-bit odbc32 goes to the 32-bit unixODBC of the
# system, which needs a 32-bit driver, and Ubuntu has none (no libsqliteodbc:i386). Built as Debian builds the
# 64-bit one (libsqliteodbc 0.99991: no Unicode entry points, odbcinst found with dlopen, the system's SQLite).
# Needs: zig (python3 -m pip install --user ziglang: its cross compiler for i386), the 32-bit libsqlite3 that Wine
# has (/usr/lib/i386-linux-gnu/libsqlite3.so.0), the headers of unixODBC (unixodbc-dev, or taken from its package
# with apt download) and of SQLite (libsqlite3-dev, else those of the runtime).
# Makes i386/libsqlite3odbc.so and i386/odbcinst.ini ([SQLite3], [SQLite3 ODBC Driver]); run_tests.py gives the
# 32-bit programs that file (ODBCSYSINI).
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later (this script; the driver: license.terms).
set -e
cd "$(dirname "$0")"
HERE=$(pwd)
OUT=$HERE/i386
LIB32=/usr/lib/i386-linux-gnu/libsqlite3.so.0
[ -f "$LIB32" ] || { echo "no 32-bit libsqlite3 ($LIB32): Wine's i386 libraries are needed"; exit 1; }
python3 -m ziglang version >/dev/null 2>&1 || { echo "zig is needed: python3 -m pip install --user ziglang"; exit 1; }
mkdir -p "$OUT/include"
# the headers of unixODBC: sizeof(long) of i386 in unixodbc.h
H=/usr/include
TMP=""
if [ ! -f $H/sql.h ]; then
    TMP=$(mktemp -d)
    (cd "$TMP" && apt download unixodbc-dev >/dev/null 2>&1 && dpkg-deb -x unixodbc-dev_*.deb x) || { echo "unixodbc-dev: apt download failed"; exit 1; }
    H=$TMP/x/usr/include
fi
for f in sql.h sqlext.h sqltypes.h sqlucode.h odbcinst.h odbcinstext.h; do cp "$H/$f" "$OUT/include/"; done
UH=$(ls "$H"/*/unixodbc.h "$H"/unixodbc.h 2>/dev/null | head -1)
sed 's/SIZEOF_LONG_INT 8/SIZEOF_LONG_INT 4/' "$UH" > "$OUT/include/unixodbc.h"
if [ -f /usr/include/sqlite3.h ]; then cp /usr/include/sqlite3.h "$OUT/include/"
else cp ../../../runtime/db/sqlite/sqlite3.h "$OUT/include/"; fi
[ -n "$TMP" ] && rm -rf "$TMP"
python3 -m ziglang cc -target x86-linux-gnu.2.17 -shared -fPIC -O2 -w \
    -DWITHOUT_WINTERFACE=1 "-DCOLATTRIBUTE_LAST_ARG_TYPE=SQLLEN *" -DSETSTMTOPTION_LAST_ARG_TYPE=SQLULEN \
    -DDRIVER_VER_INFO=\"$(cat VERSION)\" -DHAVE_UNIXODBC=1 -DHAVE_SQLWCHAR=1 -DHAVE_SQLLEN=1 -DHAVE_SQLULEN=1 \
    -DHAVE_SQLROWCOUNT=1 -DHAVE_SQLSETPOSIROW=1 -DHAVE_SQLROWOFFSET=1 -DHAVE_SQLROWSETSIZE=1 -DHAVE_USLEEP=1 \
    -DHAVE_NANOSLEEP=1 -DHAVE_LOCALECONV=1 -DUSE_DLOPEN_FOR_GPPS=1 \
    -DHAVE_SQLITE3COLUMNTABLENAME=1 -DHAVE_SQLITE3COLUMNDATABASENAME=1 -DHAVE_SQLITE3COLUMNORIGINNAME=1 \
    -DHAVE_SQLITE3LOADEXTENSION=1 -DHAVE_SQLITE3OVERLOADFUNCTION=1 -DHAVE_SQLITE3PREPAREV2=1 \
    -DHAVE_SQLITE3CLEARBINDINGS=1 -DHAVE_SQLITE3CREATEMODULEV2=1 -DHAVE_SQLITE3VFS=1 -DHAVE_SQLITE3PROFILE=1 \
    -DHAVE_SQLITE3STRNICMP=1 -DHAVE_SQLITE3TABLECOLUMNMETADATA=1 -DHAVE_SQLITE3CLOSEV2=1 \
    -I"$OUT/include" sqlite3odbc.c "$LIB32" -ldl -lm -o "$OUT/libsqlite3odbc.so"
cat > "$OUT/odbcinst.ini" <<INI
[SQLite3]
Description=SQLite3 ODBC Driver (32-bit, the tests of LinBASIC Studio)
Driver=$OUT/libsqlite3odbc.so

[SQLite3 ODBC Driver]
Description=SQLite3 ODBC Driver (32-bit, the tests of LinBASIC Studio)
Driver=$OUT/libsqlite3odbc.so
INI
echo "tests/db/sqliteodbc/i386/libsqlite3odbc.so built"
