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


def ado_objects(connect):
    """the tables and views of an ADO connection string: [(T|Q, name)]"""
    return [(r[0], r[1]) for r in _ado(connect)]


def ado_fields(connect, cmdtype, source):
    """the fields of a command: [(name, ADO type, defined size, precision)]"""
    return [(r[1], int(r[2]), int(r[3]), int(r[4])) for r in _ado(connect, str(cmdtype), source) if r[0] == "F" and len(r) >= 5]
