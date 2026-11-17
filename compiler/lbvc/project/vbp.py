# LinBASIC Studio - .vbp / .vbg reader and writer
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""A .vbp is a list of Key=Value lines, optionally followed by [Section] blocks
(add-in settings).  Order and repeated keys matter (Form=, Module=, Reference=...),
so the file is kept as an ordered list of entries and written back unchanged
unless an entry is edited.
"""
import os

DEFAULT_ENCODING = "cp1252"


def _read_text(path, encoding):
    with open(path, "rb") as f:
        data = f.read()
    crlf = b"\r\n" in data
    return data.decode(encoding, errors="surrogateescape"), crlf


def _unquote(v):
    v = v.strip()
    if len(v) >= 2 and v[0] == '"' and v[-1] == '"':
        return v[1:-1].replace('""', '"')
    return v


class Reference:
    """Reference=*\\G{GUID}#major.minor#lcid#path#description"""
    def __init__(self, raw):
        self.raw = raw
        body = raw[2:] if raw.startswith("*\\") else raw
        self.kind = body[:1]                    # 'G' typelib, 'A' project
        parts = body[1:].split("#")
        self.guid = parts[0] if parts else ""
        self.version = parts[1] if len(parts) > 1 else ""
        self.lcid = parts[2] if len(parts) > 2 else ""
        self.path = parts[3] if len(parts) > 3 else ""
        self.description = "#".join(parts[4:]) if len(parts) > 4 else ""

    def __repr__(self):
        return f"Reference({self.description or self.path!r})"


class ObjectRef:
    """Object={GUID}#major.minor#lcid; FILE.OCX"""
    def __init__(self, raw):
        self.raw = raw
        left, _, right = raw.partition(";")
        parts = left.strip().split("#")
        self.guid = parts[0]
        self.version = parts[1] if len(parts) > 1 else ""
        self.lcid = parts[2] if len(parts) > 2 else ""
        self.file = right.strip().strip('"')

    def __repr__(self):
        return f"ObjectRef({self.file!r})"


class Project:
    MEMBER_KEYS = ("Form", "Module", "Class", "UserControl", "PropertyPage",
                   "UserDocument", "Designer", "RelatedDoc", "ResFile32")

    def __init__(self, path=None):
        self.path = path
        self.entries = []        # [key, value] in file order (top part)
        self.sections = []       # [(name, [[key, value], ...])]
        self.crlf = True
        self.encoding = DEFAULT_ENCODING

    # ---- reading -------------------------------------------------------
    @classmethod
    def parse(cls, text, path=None, crlf=True, encoding=DEFAULT_ENCODING):
        p = cls(path)
        p.crlf, p.encoding = crlf, encoding
        target = p.entries
        for line in text.splitlines():
            s = line.strip()
            if not s:
                continue
            if s.startswith("[") and s.endswith("]"):
                target = []
                p.sections.append((s[1:-1], target))
                continue
            key, sep, value = line.partition("=")
            if not sep:
                target.append([line, None])     # keep unknown lines verbatim
            else:
                target.append([key, value])
        return p

    # ---- queries -------------------------------------------------------
    def get(self, key, default=None):
        for k, v in self.entries:
            if v is not None and k.lower() == key.lower():
                return _unquote(v)
        return default

    def get_all(self, key):
        return [v for k, v in self.entries if v is not None and k.lower() == key.lower()]

    def set(self, key, value, quote=False):
        if quote:
            value = '"' + value.replace('"', '""') + '"'
        for e in self.entries:
            if e[1] is not None and e[0].lower() == key.lower():
                e[1] = value
                return
        self.entries.append([key, value])

    @property
    def name(self):
        return self.get("Name", "Project1")

    @property
    def type(self):
        return self.get("Type", "Exe")        # Exe, OleExe, OleDll, Control

    @property
    def startup(self):
        return self.get("Startup", "")          # "Form1", "Sub Main", "(None)"

    @property
    def exe_name(self):
        return self.get("ExeName32", self.name + ".exe")

    def members(self):
        """Yield (kind, name, relative_path) for every source file of the project."""
        for k, v in self.entries:
            if v is None:
                continue
            kind = k.strip()
            if kind not in self.MEMBER_KEYS or kind in ("RelatedDoc", "ResFile32"):
                continue
            v = v.strip()
            if kind in ("Module", "Class") and ";" in v:
                name, _, rel = v.partition(";")
                yield kind, name.strip(), rel.strip()
            else:
                rel = _unquote(v)
                yield kind, os.path.splitext(os.path.basename(rel.replace("\\", "/")))[0], rel

    def references(self):
        return [Reference(v.strip()) for v in self.get_all("Reference")]

    def objects(self):
        return [ObjectRef(v.strip()) for v in self.get_all("Object")]

    def resolve(self, rel):
        """the file of a member; names from Windows are found whatever their case (main.frm -> MAIN.FRM)"""
        from .fsutil import ci_path, find_moved
        base = os.path.dirname(self.path) if self.path else "."
        p = ci_path(os.path.normpath(os.path.join(base, rel.replace("\\", "/"))))
        if not os.path.exists(p):               # C:\My Projects\x.frm, PROJEC~1\FORM1~1.FRM: in the folder of the project
            p = find_moved(base, rel) or p
        return p

    # ---- writing -------------------------------------------------------
    def dumps(self):
        out = []
        for k, v in self.entries:
            out.append(k if v is None else f"{k}={v}")
        for name, items in self.sections:
            out.append("")
            out.append(f"[{name}]")
            for k, v in items:
                out.append(k if v is None else f"{k}={v}")
        nl = "\r\n" if self.crlf else "\n"
        return nl.join(out) + nl

    def save(self, path=None):
        path = path or self.path
        with open(path, "wb") as f:
            f.write(self.dumps().encode(self.encoding, errors="surrogateescape"))


class ProjectGroup:
    """VBGROUP 5.0 / StartupProject=... / Project=..."""
    def __init__(self, path=None):
        self.path = path
        self.header = "VBGROUP 5.0"
        self.entries = []
        self.crlf = True

    @classmethod
    def parse(cls, text, path=None, crlf=True):
        g = cls(path)
        g.crlf = crlf
        lines = [l for l in text.splitlines() if l.strip()]
        if lines and lines[0].upper().startswith("VBGROUP"):
            g.header = lines.pop(0)
        for l in lines:
            k, _, v = l.partition("=")
            g.entries.append([k, v])
        return g

    def projects(self):
        return [v.strip() for k, v in self.entries if k.strip().lower() in ("project", "startupproject")]

    @property
    def startup(self):
        for k, v in self.entries:
            if k.strip().lower() == "startupproject":
                return v.strip()
        return None

    def dumps(self):
        nl = "\r\n" if self.crlf else "\n"
        return nl.join([self.header] + [f"{k}={v}" for k, v in self.entries]) + nl


def load_project(path, encoding=DEFAULT_ENCODING):
    text, crlf = _read_text(path, encoding)
    return Project.parse(text, path, crlf, encoding)


def load_group(path, encoding=DEFAULT_ENCODING):
    text, crlf = _read_text(path, encoding)
    return ProjectGroup.parse(text, path, crlf)
