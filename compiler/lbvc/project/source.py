# LinBASIC Studio - reader/writer for VB6 source files
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""One reader for every VB6 text module:

  .frm  Form          VERSION 5.00 + Object lines + Begin VB.Form ... End + attributes + code
  .ctl  UserControl   same layout, Begin VB.UserControl
  .dob  UserDocument  same layout, Begin VB.UserDocument
  .pag  PropertyPage  same layout, Begin VB.PropertyPage
  .cls  Class         VERSION 1.0 CLASS + BEGIN ... END block + attributes + code
  .bas  Module        attributes + code

The designer section is a tree of Control objects with ordered Property and
PropertyBlock (BeginProperty ... EndProperty) children.  Every item keeps the
line it was read from, so an untouched file is written back byte for byte;
edited items are rewritten in the layout VB6 itself uses.
"""
import re

DEFAULT_ENCODING = "cp1252"
INDENT = "   "
NAME_WIDTH = 16

_FRX_RE = re.compile(r'^(\$?)"([^"]+)":([0-9A-Fa-f]{4,8})$')
_HEX_RE = re.compile(r'^&H([0-9A-Fa-f]+)(&?)$')
_NUM_RE = re.compile(r'^-?(\d+\.?\d*|\.\d+)([Ee][+-]?\d+)?$')


def split_comment(text):
    """Split 'value  'comment' at the first apostrophe outside quotes."""
    q = False
    for i, ch in enumerate(text):
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            return text[:i].rstrip(), text[i + 1:]
    return text.rstrip(), None


def parse_value(text):
    """Turn the right-hand side of a property line into a Python value."""
    t = text.strip()
    m = _FRX_RE.match(t)
    if m:
        return FrxRef(m.group(2), int(m.group(3), 16), m.group(1) == "$")
    if len(t) >= 2 and t[0] == '"' and t[-1] == '"':
        return t[1:-1].replace('""', '"')
    m = _HEX_RE.match(t)
    if m:
        v = int(m.group(1), 16)
        return v - (1 << 32) if v >= 0x80000000 else v
    if _NUM_RE.match(t):
        return float(t) if any(c in t for c in ".Ee") else int(t)
    return t                                   # GUIDs, enum words, anything else


class Raw(str):
    """a value written as it is, without quotes (Shortcut = ^A)"""


def format_value(v):
    if isinstance(v, FrxRef):
        return v.text()
    if isinstance(v, Raw):
        return str.__str__(v)
    if type(v).__name__ == "FrxData":
        return '"(frx)"'                         # never written: the designer gives it an offset first
    if isinstance(v, bool):
        return "-1  'True" if v else "0   'False"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        s = repr(v)
        return s[:-2] if s.endswith(".0") else s
    if isinstance(v, str):
        return '"' + v.replace('"', '""') + '"'
    return str(v)


class FrxRef:
    """A property whose data lives in the companion binary file."""
    __slots__ = ("file", "offset", "is_string")

    def __init__(self, file, offset, is_string=False):
        self.file, self.offset, self.is_string = file, offset, is_string

    def text(self):
        return f'{"$" if self.is_string else ""}"{self.file}":{self.offset:04X}'

    def __eq__(self, o):
        return isinstance(o, FrxRef) and (self.file, self.offset, self.is_string) == (o.file, o.offset, o.is_string)

    def __repr__(self):
        return f"FrxRef({self.text()})"


class Property:
    __slots__ = ("name", "value", "raw_value", "comment", "raw", "_orig")

    def __init__(self, name, value, raw_value=None, comment=None, raw=None):
        self.name = name
        self.value = value
        self.raw_value = raw_value if raw_value is not None else format_value(value)
        self.comment = comment
        self.raw = raw
        self._orig = value

    @property
    def dirty(self):
        return self.raw is None or self.value != self._orig

    def line(self, depth):
        if not self.dirty:
            return self.raw
        rv = format_value(self.value)
        com = ""
        if not isinstance(self.value, bool) and self.comment is not None and self.value == self._orig:
            com = "  '" + self.comment
        n = self.name                   # VB6: "Object." before a name padded as it is (Object.Width           =)
        head = f"Object.{n[7:]:<{NAME_WIDTH}}" if n.startswith("Object.") else f"{n:<{NAME_WIDTH}}"
        return f"{INDENT * depth}{head}=   {rv}{com}"

    def __repr__(self):
        return f"{self.name}={self.value!r}"


class PropertyBlock:
    """BeginProperty Font ... EndProperty  (may nest, may carry a GUID)."""

    def __init__(self, name, guid=None, raw_begin=None, raw_end=None):
        self.name = name
        self.guid = guid
        self.items = []            # Property | PropertyBlock
        self.raw_begin = raw_begin
        self.raw_end = raw_end

    def get(self, name, default=None):
        for it in self.items:
            if isinstance(it, Property) and it.name.lower() == name.lower():
                return it.value
        return default

    def lines(self, depth):
        pad = INDENT * depth
        yield self.raw_begin if self.raw_begin is not None else \
            f"{pad}BeginProperty {self.name}{' ' + self.guid if self.guid else ''} "
        for it in self.items:
            if isinstance(it, Property):
                yield it.line(depth + 1)
            else:
                yield from it.lines(depth + 1)
        yield self.raw_end if self.raw_end is not None else f"{pad}EndProperty"


class Control:
    """Begin <Library.Class> <Name> ... End"""

    def __init__(self, cls, name, raw_begin=None, raw_end=None):
        self.cls = cls                 # 'VB.CommandButton', 'MSComctlLib.ListView'
        self.name = name
        self.items = []                # Property | PropertyBlock  (designer order)
        self.children = []             # nested Controls
        self.order = []                # ('p', item) | ('c', control) — original order
        self.raw_begin = raw_begin
        self.raw_end = raw_end
        self.parent = None

    # -- convenience -------------------------------------------------------
    @property
    def library(self):
        return self.cls.split(".", 1)[0] if "." in self.cls else ""

    @property
    def type_name(self):
        return self.cls.split(".", 1)[-1]

    @property
    def index(self):
        v = self.get("Index")
        return v if isinstance(v, int) else None

    def get(self, name, default=None):
        for it in self.items:
            if isinstance(it, Property) and it.name.lower() == name.lower():
                return it.value
        return default

    def prop(self, name):
        for it in self.items:
            if isinstance(it, Property) and it.name.lower() == name.lower():
                return it
        return None

    def block(self, name):
        for it in self.items:
            if isinstance(it, PropertyBlock) and it.name.lower() == name.lower():
                return it
        return None

    def set(self, name, value):
        p = self.prop(name)
        if p:
            p.value = value
        else:
            p = Property(name, value)
            self.items.append(p)
            # new properties go before the first child control, as VB6 writes them
            pos = next((i for i, (k, _) in enumerate(self.order) if k == "c"), len(self.order))
            self.order.insert(pos, ("p", p))

    def add_child(self, ctl):
        ctl.parent = self
        self.children.append(ctl)
        self.order.append(("c", ctl))

    def walk(self):
        yield self
        for c in self.children:
            yield from c.walk()

    def lines(self, depth):
        pad = INDENT * depth
        yield self.raw_begin if self.raw_begin is not None else f"{pad}Begin {self.cls} {self.name} "
        for kind, it in self.order:
            if kind == "c":
                yield from it.lines(depth + 1)
            elif isinstance(it, Property):
                yield it.line(depth + 1)
            else:
                yield from it.lines(depth + 1)
        yield self.raw_end if self.raw_end is not None else f"{pad}End"

    def __repr__(self):
        return f"Control({self.cls} {self.name}, {len(self.children)} children)"


class SourceFile:
    KIND_BY_EXT = {".frm": "Form", ".ctl": "UserControl", ".dob": "UserDocument",
                   ".pag": "PropertyPage", ".cls": "Class", ".bas": "Module", ".dsr": "Designer"}

    def __init__(self, path=None, kind=None):
        self.path = path
        self.kind = kind
        self.version = None            # 'VERSION 5.00' line
        self.objects = []              # raw 'Object = ...' lines
        self.root = None               # Control (forms etc.)
        self.class_props = None        # PropertyBlock-like list for .cls BEGIN..END
        self.class_raw = None          # (begin_line, end_line)
        self.attributes = []           # [(name, value_text, raw)] header attributes
        self.code = []                 # remaining lines, untouched
        self.crlf = True
        self.encoding = DEFAULT_ENCODING
        self.trailing_newline = True

    # -- queries -------------------------------------------------------------
    @property
    def name(self):
        for n, v, _ in self.attributes:
            if n == "VB_Name":
                return parse_value(v)
        return self.root.name if self.root else None

    def attribute(self, name, default=None):
        for n, v, _ in self.attributes:
            if n.lower() == name.lower():
                return parse_value(split_comment(v)[0])
        return default

    def controls(self):
        return list(self.root.walk()) if self.root else []

    def frx_refs(self):
        out = []
        for c in self.controls():
            stack = list(c.items)
            while stack:
                it = stack.pop()
                if isinstance(it, PropertyBlock):
                    stack.extend(it.items)
                elif isinstance(it.value, FrxRef):
                    out.append((c, it))
        return out

    @property
    def code_text(self):
        return "\n".join(self.code)

    # -- writing ---------------------------------------------------------------
    def lines(self):
        if self.version is not None:
            yield self.version
        if self.kind == "Class" and self.class_props is not None:
            yield self.class_raw[0] if self.class_raw else "BEGIN"
            for p in self.class_props:
                yield p.line(1).replace(INDENT, "  ", 1) if p.dirty else p.raw
            yield self.class_raw[1] if self.class_raw else "END"
        for o in self.objects:
            yield o
        if self.root is not None:
            yield from self.root.lines(0)
        for _, _, raw in self.attributes:
            yield raw
        yield from self.code

    def dumps(self):
        nl = "\r\n" if self.crlf else "\n"
        s = nl.join(self.lines())
        return s + nl if self.trailing_newline else s

    def save(self, path=None):
        path = path or self.path
        with open(path, "wb") as f:
            f.write(self.dumps().encode(self.encoding, errors="surrogateescape"))


class VB6FormatError(Exception):
    def __init__(self, msg, line):
        super().__init__(f"{line}: {msg}")
        self.line = line


def _parse_prop(line, lineno):
    key, sep, rest = line.partition("=")
    if not sep:
        raise VB6FormatError(f"expected 'Name = Value': {line.strip()!r}", lineno)
    raw_value, comment = split_comment(rest.strip())
    return Property(key.strip(), parse_value(raw_value), raw_value, comment, raw=line)


def parse_source(text, path=None, kind=None, crlf=True, encoding=DEFAULT_ENCODING):
    sf = SourceFile(path, kind)
    sf.crlf, sf.encoding = crlf, encoding
    sf.trailing_newline = text.endswith("\n")
    lines = text.split("\n")
    if sf.trailing_newline:
        lines.pop()
    lines = [l[:-1] if l.endswith("\r") else l for l in lines]
    i, n = 0, len(lines)

    if i < n and lines[i].upper().startswith("VERSION"):
        sf.version = lines[i]
        i += 1
        if "CLASS" in sf.version.upper():
            sf.kind = sf.kind or "Class"

    # .cls property block
    if i < n and lines[i].strip().upper() == "BEGIN":
        begin = lines[i]
        i += 1
        sf.class_props = []
        while i < n and lines[i].strip().upper() != "END":
            sf.class_props.append(_parse_prop(lines[i], i + 1))
            i += 1
        if i >= n:
            raise VB6FormatError("BEGIN without END", i)
        sf.class_raw = (begin, lines[i])
        i += 1

    while i < n and lines[i].lstrip().startswith("Object"):
        sf.objects.append(lines[i])
        i += 1

    # designer tree
    if i < n and lines[i].lstrip().startswith("Begin "):
        stack = []
        while i < n:
            line = lines[i]
            s = line.strip()
            lineno = i + 1
            i += 1
            if s.startswith("Begin "):
                parts = s.split()
                if len(parts) < 3:
                    raise VB6FormatError("Begin needs a class and a name", lineno)
                ctl = Control(parts[1], parts[2], raw_begin=line)
                if stack:
                    if not isinstance(stack[-1], Control):
                        raise VB6FormatError("Begin inside BeginProperty", lineno)
                    stack[-1].add_child(ctl)
                else:
                    sf.root = ctl
                stack.append(ctl)
            elif s.startswith("BeginProperty "):
                parts = s.split()
                blk = PropertyBlock(parts[1], parts[2] if len(parts) > 2 else None, raw_begin=line)
                owner = stack[-1]
                owner.items.append(blk)
                if isinstance(owner, Control):
                    owner.order.append(("p", blk))
                stack.append(blk)
            elif s == "EndProperty":
                if not stack or not isinstance(stack[-1], PropertyBlock):
                    raise VB6FormatError("EndProperty without BeginProperty", lineno)
                stack.pop().raw_end = line
            elif s == "End":
                if not stack or not isinstance(stack[-1], Control):
                    raise VB6FormatError("End without Begin", lineno)
                stack.pop().raw_end = line
                if not stack:
                    break
            elif s == "":
                continue
            else:
                if not stack:
                    raise VB6FormatError("property outside Begin...End", lineno)
                p = _parse_prop(line, lineno)
                owner = stack[-1]
                owner.items.append(p)
                if isinstance(owner, Control):
                    owner.order.append(("p", p))
        if stack:
            raise VB6FormatError(f"'{stack[-1].name}' has no End", n)
        sf.kind = sf.kind or {"Form": "Form", "MDIForm": "Form", "UserControl": "UserControl",
                              "UserDocument": "UserDocument",
                              "PropertyPage": "PropertyPage"}.get(sf.root.type_name, "Form")

    # header attributes
    while i < n and lines[i].startswith("Attribute "):
        name, _, val = lines[i][10:].partition("=")
        sf.attributes.append((name.strip(), val.strip(), lines[i]))
        i += 1

    sf.code = lines[i:]
    sf.kind = sf.kind or "Module"
    return sf


def load_source(path, encoding=DEFAULT_ENCODING):
    import os
    with open(path, "rb") as f:
        data = f.read()
    text = data.decode(encoding, errors="surrogateescape")
    kind = SourceFile.KIND_BY_EXT.get(os.path.splitext(path)[1].lower())
    return parse_source(text, path, kind, crlf=b"\r\n" in data, encoding=encoding)
