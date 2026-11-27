"""The tables, queries and fields of a database at design time: the drop-down lists of RecordSource
and DataField in the Properties window. The runtime's tool lbdbinfo reads the database (a Jet .mdb
through its SQLite copy, as the programs read it)."""
import os
import subprocess

from . import paths

_cache = {}


def tool():
    if os.name == "nt":
        cands = (os.path.join(paths.RUNTIME, paths.host_target(), "lbdbinfo.exe"),)
    else:
        cands = (os.path.join(paths.RUNTIME, "db", "lbdbinfo"), os.path.join(paths.RUNTIME, "lbdbinfo"))
    for p in cands:
        if os.path.isfile(p) and os.access(p, os.X_OK):
            return p
    return None


def resolve(name, base_dir=None):
    """DatabaseName as a file of this computer (a relative name: from the project's folder)"""
    if not name:
        return None
    p = name.replace("\\", "/") if not os.path.exists(name) else name
    if not os.path.isabs(p) and base_dir:
        p = os.path.join(base_dir, p)
    return p if os.path.isfile(p) else None


def _run(path, *args):
    t = tool()
    if t is None or path is None:
        return []
    try:
        st = os.stat(path)
    except OSError:
        return []
    key = (path, st.st_mtime, args)
    if key in _cache:
        return _cache[key]
    try:
        r = subprocess.run([t, path, *args], capture_output=True, text=True, timeout=30, **paths.NO_WINDOW)
    except (OSError, subprocess.TimeoutExpired):
        return []
    out = [line.split("\t", 1) for line in r.stdout.splitlines() if "\t" in line] if r.returncode == 0 else []
    _cache[key] = out
    return out


def objects(path):
    """the names of the tables and the queries, as the RecordSource list of VB6 has them"""
    rows = _run(path)
    return [n for k, n in rows if k == "T"] + [n for k, n in rows if k == "Q"]


def fields(path, source):
    if not source:
        return []
    return [n for k, n in _run(path, source) if k == "F"]


# ---------------------------------------------------------------- ADO (the DataEnvironment designer)
class AdoError(Exception):
    pass


def _ado(*args):
    t = tool()
    if t is None:
        raise AdoError("lbdbinfo is not built (runtime/db/build.sh)")
    key = ("ado",) + args
    if key in _cache:
        return _cache[key]
    try:
        r = subprocess.run([t, "-ado", *args], capture_output=True, text=True, timeout=60, **paths.NO_WINDOW)
    except (OSError, subprocess.TimeoutExpired) as e:
        raise AdoError(str(e))
    if r.returncode != 0:
        raise AdoError(r.stderr.strip() or "the connection failed")
    out = [line.split("\t") for line in r.stdout.splitlines() if "\t" in line]
    _cache[key] = out
    return out


def forget():
    """Refresh: the database may have changed"""
    _cache.clear()


def ado_objects(connect, system=False):
    """the tables and views of an ADO connection string: [(T|Q, name)]; system: also the system tables (S) - Show
    System Objects of the Data View"""
    return [(r[0], r[1]) for r in _ado(connect) if system or r[0] != "S"]


def ado_fields(connect, cmdtype, source):
    """the fields of a command: [(name, ADO type, defined size, precision)]"""
    return [(r[1], int(r[2]), int(r[3]), int(r[4])) for r in _ado(connect, str(cmdtype), source) if r[0] == "F" and len(r) >= 5]


# ---------------------------------------------------------------- the Data View and the Query Designer (phase 12.3)
def _unescape(v):
    if v == "\\N":
        return None
    out, i = [], 0
    while i < len(v):
        if v[i] == "\\" and i + 1 < len(v):
            out.append({"t": "\t", "n": "\n", "r": "\r", "\\": "\\"}.get(v[i + 1], v[i + 1]))
            i += 2
        else:
            out.append(v[i])
            i += 1
    return "".join(out)


def ado_query(connect, sql, limit=1000):
    """a query run through ADO: ([(column, ADO type)], [[value or None]], more rows than limit)"""
    t = tool()
    if t is None:
        raise AdoError("lbdbinfo is not built (runtime/db/build.sh)")
    try:
        r = subprocess.run([t, "-ado-sql", connect, sql, str(limit)], capture_output=True, timeout=120,
                           **paths.NO_WINDOW)
    except (OSError, subprocess.TimeoutExpired) as e:
        raise AdoError(str(e))
    out = r.stdout.decode("utf-8", "replace")
    if r.returncode != 0:
        raise AdoError(r.stderr.decode("utf-8", "replace").strip() or "the query failed")
    cols, rows, more = [], [], False
    for line in out.split("\n"):
        if not line:
            continue
        parts = line.split("\t")
        if parts[0] == "C" and len(parts) >= 3:
            cols.append((parts[1], int(parts[2])))
        elif parts[0] == "R":
            rows.append([_unescape(x) for x in parts[1:]])
        elif parts[0] == "M":
            more = True
    return cols, rows, more


def ado_exec(connect, sql):
    """an action query (UPDATE / INSERT / DELETE ...) run through ADO: the number of records it changed (or None)"""
    t = tool()
    if t is None:
        raise AdoError("lbdbinfo is not built (runtime/db/build.sh)")
    try:
        r = subprocess.run([t, "-ado-exec", connect, sql], capture_output=True, timeout=120, **paths.NO_WINDOW)
    except (OSError, subprocess.TimeoutExpired) as e:
        raise AdoError(str(e))
    if r.returncode != 0:
        raise AdoError(r.stderr.decode("utf-8", "replace").strip() or "the statement failed")
    forget()
    for line in r.stdout.decode("utf-8", "replace").splitlines():   # OK<tab>the records affected
        p = line.split("\t")
        if p[0] == "OK" and len(p) > 1 and p[1].lstrip("-").isdigit() and int(p[1]) >= 0:
            return int(p[1])
    return None


def ado_relations(connect):
    """the relations of the database (the joins the Diagram pane draws): [(primary table, its column, foreign table,
    its column)] - those LinBASIC Studio keeps for a Jet database (MSysLB_Relations); none through ODBC"""
    try:
        _cols, rows, _more = ado_query(connect, "SELECT tbl, ftbl, fields FROM MSysLB_Relations", 0)
    except AdoError:
        return []
    out = []
    for pt, ft, fields in rows:
        for pair in (fields or "").split(";"):
            if "=" in pair:
                pc, fc = pair.split("=", 1)
                out.append((pt, pc, ft, fc))
    return out


def ado_keys(connect, table):
    """the columns of the primary key of a table (the key of a row edited in the Results pane)"""
    try:
        _cols, rows, _more = ado_query(connect, "SELECT fields FROM MSysLB_Indexes WHERE isprimary <> 0 AND tbl = '%s'"
                                       % table.replace("'", "''"), 0)
    except AdoError:
        return []
    for (fields,) in rows:
        return [f.lstrip("+-").strip() for f in (fields or "").split(";") if f.strip()]
    return []


def ado_columns(connect, table):
    """the columns of a table in the order of the database"""
    try:
        return [f[0] for f in ado_fields(connect, 2, table)]
    except AdoError:
        return []
