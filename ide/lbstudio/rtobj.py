"""The objects of the LinBASIC Studio runtime for the IDE (phase 12.4): DAO, ADO ... through "lbdbinfo -server", so that
the Visual Data Manager works with a database exactly as a program does (the same DAO, the same MSysLB_* tables, the
same SQL of Jet). An object is used as late binding in VB:

    srv = Server()
    dbe = srv.create("DAO.DBEngine.36")
    db = dbe.call("OpenDatabase", "nwind.mdb")
    for td in db.each("TableDefs"): print(td.get("Name"))
"""
import datetime
import decimal
import os
import subprocess
import threading

from . import dbinfo, paths

EMPTY = object()                                # Empty (a Variant with no value)
MISSING = object()                              # an optional argument left out
OLE_EPOCH = datetime.datetime(1899, 12, 30)


class RtError(Exception):
    """a run-time error of the runtime: its Err.Number and Err.Description"""

    def __init__(self, number, description):
        super().__init__(f"{description} ({number})" if number else description)
        self.number = number
        self.description = description


def _esc(s):
    return s.replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n").replace("\r", "\\r")


def _unesc(s):
    out, i = [], 0
    while i < len(s):
        if s[i] == "\\" and i + 1 < len(s):
            out.append({"t": "\t", "n": "\n", "r": "\r"}.get(s[i + 1], s[i + 1]))
            i += 2
        else:
            out.append(s[i])
            i += 1
    return "".join(out)


def to_date(d):
    """a Date of VB (days since 30 Dec 1899) as a datetime"""
    return OLE_EPOCH + datetime.timedelta(days=d)


def from_date(dt):
    if isinstance(dt, datetime.date) and not isinstance(dt, datetime.datetime):
        dt = datetime.datetime(dt.year, dt.month, dt.day)
    return (dt - OLE_EPOCH) / datetime.timedelta(days=1)


class Obj:
    """an object of the runtime"""

    def __init__(self, srv, oid):
        self.srv = srv
        self.id = oid

    def __repr__(self):
        return f"<rtobj {self.id}>"

    def __eq__(self, other):
        return isinstance(other, Obj) and other.id == self.id and other.srv is self.srv

    def __hash__(self):
        return hash(self.id)

    def get(self, name, *args):
        return self.srv.member(self, "G", name, args)

    def let(self, name, value, *args):
        return self.srv.member(self, "L", name, args + (value,))

    def set(self, name, value, *args):
        return self.srv.member(self, "S", name, args + (value,))

    def call(self, name, *args):
        return self.srv.member(self, "M", name, args)

    def typename(self):
        return self.srv.request("T", str(self.id))

    def count(self, coll):
        return self.get(coll).get("Count")

    def each(self, coll):
        """the items of a collection of the object (For Each)"""
        c = self.get(coll)
        n = c.get("Count")
        return [c.get("Item", i) for i in range(n)]

    def release(self):
        if self.id:
            self.srv.request("R", str(self.id))


class Server:
    """one lbdbinfo -server: its objects live as long as it does"""

    def __init__(self, tool=None):
        t = tool or dbinfo.tool()
        if t is None:
            raise RtError(0, "lbdbinfo is not built (runtime/db/build.sh)")
        self.proc = subprocess.Popen([t, "-server"], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                     stderr=subprocess.DEVNULL, cwd=os.getcwd(), **paths.NO_WINDOW)
        self.lock = threading.Lock()

    def close(self):
        p = getattr(self, "proc", None)
        if p is not None and p.poll() is None:
            try:
                p.stdin.write(b"Q\n")
                p.stdin.flush()
                p.wait(5)
            except (OSError, subprocess.TimeoutExpired):
                p.kill()
        self.proc = None

    def __del__(self):
        try:
            self.close()
        except Exception:                       # noqa: BLE001 - (at exit)
            pass

    def request(self, *fields):
        line = "\t".join(_esc(f) for f in fields) + "\n"
        with self.lock:
            if self.proc is None or self.proc.poll() is not None:
                raise RtError(0, "the database tool stopped")
            self.proc.stdin.write(line.encode("utf-8"))
            self.proc.stdin.flush()
            ans = self.proc.stdout.readline().decode("utf-8", "replace").rstrip("\r\n")
        if not ans:
            raise RtError(0, "the database tool stopped")
        parts = ans.split("\t")
        if parts[0] == "ER":
            raise RtError(int(parts[1]) if len(parts) > 1 and parts[1].lstrip("-").isdigit() else 0,
                          _unesc(parts[2]) if len(parts) > 2 else "")
        return self.value(_unesc(parts[1]) if len(parts) > 1 else "e")

    def value(self, v):
        k, _, t = v.partition(":")
        if k == "e":
            return EMPTY
        if k == "n":
            return None
        if k == "s":
            return t
        if k in ("h", "i", "y"):
            return int(t)
        if k in ("f", "d"):
            return float(t)
        if k == "c":
            return decimal.Decimal(t.replace(",", "."))
        if k == "x":
            return decimal.Decimal(t.replace(",", "."))
        if k == "t":
            return to_date(float(t))
        if k == "b":
            return t == "1"
        if k == "o":
            return Obj(self, int(t)) if int(t) else None
        if k == "a":
            return bytes.fromhex(t)
        return t

    @staticmethod
    def arg(v):
        if v is EMPTY:
            return "e"
        if v is MISSING:
            return "m"
        if v is None:
            return "n"
        if isinstance(v, bool):
            return f"b:{1 if v else 0}"
        if isinstance(v, int):
            return f"i:{v}" if -2**31 <= v < 2**31 else f"d:{float(v)!r}"
        if isinstance(v, float):
            return f"d:{v!r}"
        if isinstance(v, decimal.Decimal):
            return f"c:{v}"
        if isinstance(v, (datetime.datetime, datetime.date)):
            return f"t:{from_date(v)!r}"
        if isinstance(v, Obj):
            return f"o:{v.id}"
        if isinstance(v, (bytes, bytearray)):
            return "a:" + bytes(v).hex()
        return "s:" + str(v)

    def create(self, progid):
        return self.request("N", progid)

    def member(self, obj, kind, name, args):
        return self.request("C", str(obj.id if obj is not None else 0), kind, name, *[self.arg(a) for a in args])
