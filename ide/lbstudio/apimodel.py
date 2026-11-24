"""The data of the API Viewer (phase 9.4), without Qt: the Constants, Declares and Types of a text file in the syntax
of Visual Basic (our LBAPI.TXT, made by tools/mkapi.py from the headers of MinGW-w64, or any file of the kind), the
database the viewer makes of one (SQLite), the text of the selected items with their scope, and where Insert puts
them in a module (the end of its declarations).
"""
import os
import re
import sqlite3

from . import paths

KINDS = ["Constants", "Declares", "Types"]
DEFAULT_FILE = "LBAPI.TXT"
LINUX_FILE = "LINUX.TXT"

_SCOPE = r"(?:(?:Public|Private|Global)\s+)?"
RE_CONST = re.compile(r"^\s*" + _SCOPE + r"Const\s+([A-Za-z]\w*)\b(.*)$", re.I)
RE_DECL = re.compile(r"^\s*" + _SCOPE + r"(Declare\s+(?:Function|Sub)\s+([A-Za-z]\w*)\b.*)$", re.I)
RE_TYPE = re.compile(r"^\s*" + _SCOPE + r"Type\s+([A-Za-z]\w*)\s*$", re.I)
RE_END_TYPE = re.compile(r"^\s*End\s+Type\b", re.I)
RE_PROC = re.compile(r"^\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(?:Sub|Function|Property)\s+\w",
                     re.I)


def api_dir():
    """the folder of our database (the Winapi folder of VB6)"""
    return paths.data("api")


def default_file():
    return os.path.join(api_dir(), DEFAULT_FILE)


def strip_comment(line):
    """the line without its comment (' outside the strings)"""
    q = False
    for i, c in enumerate(line):
        if c == '"':
            q = not q
        elif c == "'" and not q:
            return line[:i]
    return line


class ApiData:
    """the items of a file: kind -> {lower name: (name, text without the scope)}, and their order"""

    def __init__(self):
        self.items = {k: {} for k in KINDS}
        self.order = {k: [] for k in KINDS}
        self.path = None
        self.is_text = True

    def add(self, kind, name, text):
        d = self.items[kind]
        if name.lower() in d:
            return
        d[name.lower()] = (name, text)
        self.order[kind].append(name)

    def names(self, kind):
        """the Available Items: sorted from a text file, the order of the database from a database (VB6)"""
        if self.is_text:
            return sorted(self.order[kind], key=str.lower)
        return list(self.order[kind])

    def label(self, kind, name):
        """the text of a row of Available Items: a constant with its value (WM_SETTEXT = &HC)"""
        n, text = self.items[kind][name.lower()]
        if kind == "Constants":
            m = re.match(r"Const\s+\w+\s*(?:As\s+\w+\s*)?=\s*(.*)$", text, re.I)
            if m:
                return "%s = %s" % (n, m.group(1).strip())
        return n

    def text(self, kind, name, scope="Public"):
        """an item with its scope (Public / Private) as Add puts it into Selected Items"""
        n, text = self.items[kind][name.lower()]
        return "%s %s" % (scope, text)

    def count(self):
        return sum(len(v) for v in self.items.values())

    # ------------------------------------------------------------ a text file
    def load_text(self, path, encoding="cp1252"):
        with open(path, "rb") as f:
            raw = f.read().decode(encoding, errors="replace")
        self.load_lines(raw.replace("\r\n", "\n").replace("\r", "\n").split("\n"))
        self.path = path
        self.is_text = True

    def load_lines(self, lines):
        self.__init__()
        joined = []
        buf = ""
        for line in lines:                      # the continuation lines ( _) into one
            s = strip_comment(line).rstrip()
            if s.endswith(" _"):
                buf += s[:-1]
                continue
            joined.append(buf + s)
            buf = ""
        if buf:
            joined.append(buf)
        typ = None
        for s in joined:
            if not s.strip():
                continue
            if typ is not None:
                if RE_END_TYPE.match(s):
                    self.add("Types", typ[0], "Type %s\n%s\nEnd Type" % (typ[0], "\n".join(typ[1])))
                    typ = None
                else:
                    typ[1].append("    " + s.strip())
                continue
            m = RE_TYPE.match(s)
            if m:
                typ = (m.group(1), [])
                continue
            m = RE_DECL.match(s)
            if m:
                self.add("Declares", m.group(2), re.sub(r"\s+", " ", m.group(1).strip()))
                continue
            m = RE_CONST.match(s)
            if m:
                body = s.strip()
                body = re.sub(r"^(?:(?:Public|Private|Global)\s+)", "", body, flags=re.I)
                self.add("Constants", m.group(1), body)

    # ------------------------------------------------------------ the database (SQLite)
    def save_db(self, path):
        if os.path.exists(path):
            os.remove(path)
        con = sqlite3.connect(path)
        try:
            con.execute("CREATE TABLE Items (Kind TEXT, Name TEXT, Body TEXT, Seq INTEGER)")
            rows = []
            seq = 0
            for k in KINDS:
                for n in self.names(k):
                    seq += 1
                    rows.append((k, n, self.items[k][n.lower()][1], seq))
            con.executemany("INSERT INTO Items VALUES (?, ?, ?, ?)", rows)
            con.commit()
        finally:
            con.close()

    def load_db(self, path):
        con = sqlite3.connect("file:%s?mode=ro" % path, uri=True)
        try:
            rows = con.execute("SELECT Kind, Name, Body FROM Items ORDER BY Seq").fetchall()
        except sqlite3.DatabaseError as e:
            raise ValueError(str(e))
        finally:
            con.close()
        self.__init__()
        for k, n, b in rows:
            if k in self.items:
                self.add(k, n, b)
        self.path = path
        self.is_text = False


def linux_names():
    """the procedures that the runtime carries out itself on Linux (LINUX.TXT of tools/mkapi.py)"""
    out = set()
    p = os.path.join(api_dir(), LINUX_FILE)
    try:
        with open(p, encoding="ascii", errors="replace") as f:
            for line in f:
                s = line.strip()
                if s and not s.startswith("'"):
                    out.add(s.lower())
    except OSError:
        pass
    return out


def declare_entry(text):
    """the name of the procedure in its DLL: the Alias, else the name"""
    m = re.search(r'\bAlias\s+"([^"]+)"', text, re.I)
    if m:
        return m.group(1)
    m = re.match(r"(?:\w+\s+)?Declare\s+(?:Function|Sub)\s+(\w+)", text, re.I)
    return m.group(1) if m else ""


class Selection:
    """Selected Items: the items in the order they were added, their text in Full Text or Line Item view"""

    def __init__(self, data):
        self.data = data
        self.items = []                         # (kind, name)
        self.ranges = []                        # (first line, last line) of each item in the text

    def add(self, kind, name):
        if (kind, name) in self.items:
            return False
        self.items.append((kind, name))
        return True

    def remove_at_line(self, line):
        for i, (a, b) in enumerate(self.ranges):
            if a <= line <= b:
                del self.items[i]
                return True
        return False

    def item_at_line(self, line):
        for i, (a, b) in enumerate(self.ranges):
            if a <= line <= b:
                return i
        return -1

    def clear(self):
        self.items = []
        self.ranges = []

    def render(self, scope="Public", full=True):
        """the text of Selected Items: each item (Full Text) or its name (Line Item); a blank line between the
        items of Full Text"""
        out = []
        self.ranges = []
        for kind, name in self.items:
            if kind not in self.data.items or name.lower() not in self.data.items[kind]:
                text = name
            elif full:
                text = self.data.text(kind, name, scope)
            else:
                text = self.data.items[kind][name.lower()][0]
            if out and full:
                out.append("")
            first = len(out)
            out.extend(text.split("\n"))
            self.ranges.append((first, len(out) - 1))
        return "\n".join(out)

    def code(self, scope="Public"):
        """the items as Insert puts them into a module: Full Text, a blank line between them"""
        return [self.data.text(k, n, scope) for k, n in self.items
                if k in self.data.items and n.lower() in self.data.items[k]]


def declarations_end(lines):
    """the index after the last line of the declarations of a module (before its first procedure), and whether
    the module has declarations"""
    first_proc = len(lines)
    for i, l in enumerate(lines):
        if RE_PROC.match(l) and not re.match(r"^\s*(?:(?:Public|Private|Global)\s+)?Declare\b", l, re.I):
            first_proc = i
            break
    end = first_proc                            # (the blank and comment lines just above a procedure are its own)
    while end > 0 and (not lines[end - 1].strip() or
                       (first_proc < len(lines) and lines[end - 1].strip().startswith("'"))):
        end -= 1
    has = any(l.strip() and not re.match(r"^\s*Option\s", l, re.I) for l in lines[:end])
    return end, has


def insert_into(code, items):
    """the code of a module with the items at the end of its declarations: a blank line before them (after other
    declarations), between them, and before the procedures"""
    lines = code.replace("\r\n", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    end, has = declarations_end(lines)
    block = []
    if end > 0:
        block.append("")
    for i, it in enumerate(items):
        if i:
            block.append("")
        block.extend(it.split("\n"))
    rest = lines[end:]
    while rest and not rest[0].strip():
        rest.pop(0)
    if rest:
        block.append("")
    new = lines[:end] + block + rest
    return "\n".join(new) + "\n", end + (1 if end > 0 else 0)


__all__ = ["ApiData", "Selection", "KINDS", "api_dir", "default_file", "linux_names", "declare_entry",
           "declarations_end", "insert_into"]
