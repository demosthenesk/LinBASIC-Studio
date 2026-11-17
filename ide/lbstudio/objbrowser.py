"""The Object Browser (View > Object Browser, F2), as the one of VB6: the classes and the members of the
libraries a project uses (the VB6 type libraries of data/vb6: VB, VBA, VBRUN, stdole and the References and
Components of the project) and of the project itself (its modules, forms, classes, Enums and Types, from the
analysis of refactor.py).

Above: the library (<All Libraries> or one), Go Back / Go Forward, Copy to Clipboard, View Definition, Help, and
the text to search (Search: * and ? are wildcards; Find Whole Word Only on the context menu). Left: the classes,
<globals> first. Right: the members of the class. Below: the declaration of what is selected, the class and
library it belongs to (links), its help context. Show Hidden Members and Group Members on the context menu.
"""
import fnmatch
import re

from PyQt6.QtCore import Qt, pyqtSignal, QSize
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QFont, QBrush
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QComboBox, QToolButton, QSplitter, QListWidget,
                             QListWidgetItem, QTextBrowser, QTreeWidget, QTreeWidgetItem, QMenu, QApplication,
                             QLabel, QMdiSubWindow, QAbstractItemView)

from . import intellisense as IS

ALL = "<All Libraries>"
GLOBALS = "<globals>"
CLASS_ORDER = {"globals": 0}
MEMBER_ORDER = {"prop": 0, "method": 1, "event": 2, "const": 3}
CLASS_WORD = {"class": "Class", "module": "Module", "enum": "Enum", "type": "Type", "interface": "Class",
              "globals": ""}
BASE_OF_KIND = {"Form": "VB.Form", "MDIForm": "VB.MDIForm", "UserControl": "VB.UserControl",
                "PropertyPage": "VB.PropertyPage", "UserDocument": "VB.UserDocument"}
NOT_SHOWN = ("void", "HRESULT")


class Entry:
    """a class (kind globals class module enum type interface) or a member (kind prop method event const)"""
    __slots__ = ("lib", "cls", "name", "kind", "sig", "help", "hidden", "default", "access", "private",
                 "returns", "value", "params", "sym", "module", "ctl", "parent", "lib_kind", "tl")

    def __init__(self, lib, cls, name, kind, **kw):
        self.lib, self.cls, self.name, self.kind = lib, cls, name, kind
        self.sig = kw.get("sig")                 # the declaration as the details show it
        self.help = kw.get("help")               # help context number (type libraries)
        self.hidden = kw.get("hidden", False)
        self.default = kw.get("default", False)  # the default member of its class
        self.access = kw.get("access")           # 'read-only' / 'write-only' of a property
        self.private = kw.get("private", False)  # a Private member of the project
        self.returns = kw.get("returns")
        self.value = kw.get("value")
        self.params = kw.get("params") or []
        self.sym = kw.get("sym")                 # the project's symbol (refactor.Sym): View Definition
        self.module = kw.get("module")           # the IDE module of a project class
        self.ctl = kw.get("ctl")                 # a control of a form
        self.parent = kw.get("parent")           # members: the class Entry
        self.lib_kind = kw.get("lib_kind", "typelib")
        self.tl = kw.get("tl")                   # classes of a type library: the type's dict

    def key(self):
        return (self.lib, self.name.lower())

    def __repr__(self):
        return f"Entry({self.kind} {self.lib}.{self.cls + '.' if self.cls else ''}{self.name})"


# ====================================================================== the type libraries
def _member_kind(group, events):
    if events:
        return "event"
    if any(m.get("kind") == "const" for m in group):
        return "const"
    if any(m.get("invoke") in ("get", "let", "set") or m.get("kind") == "var" for m in group):
        return "prop"
    return "method"


def _params(m, drop_last=False):
    ps = IS.vb_params(m)
    return ps[:-1] if drop_last and ps else ps


def _hidden_name(n):
    return n.startswith("_") and not n.startswith("_B_")


def tl_members(tl, lib, t, cls_name, events=False, seen=None):
    """the members of a type of a type library: Entry list (properties merged from Get / Let / Set)"""
    out = []
    seen = seen if seen is not None else set()
    if t["kind"] == "coclass":
        for imp in t.get("implements", []):
            if not imp.get("default"):
                continue
            hit = tl.find(imp["name"] if "." in imp["name"] else f"{lib}.{imp['name']}")
            if hit is not None:
                out += tl_members(tl, hit[0], hit[1], cls_name, events=bool(imp.get("source")), seen=seen)
        return out
    if t["kind"] == "enum":
        return [Entry(lib, cls_name, m["name"], "const", value=m.get("value"), help=m.get("helpcontext"),
                      hidden=_hidden_name(m["name"]) or "hidden" in m.get("flags", []), returns=m.get("type"))
                for m in t.get("members", [])]
    if t["kind"] in ("interface", "dispatch"):
        for imp in t.get("implements", []):         # a base interface (not IUnknown / IDispatch)
            n = imp["name"].split(".")[-1]
            if n in ("IUnknown", "IDispatch"):
                continue
            hit = tl.find(imp["name"] if "." in imp["name"] else f"{lib}.{imp['name']}")
            if hit is not None and id(hit[1]) not in seen:
                seen.add(id(hit[1]))
                out += tl_members(tl, hit[0], hit[1], cls_name, events, seen)
    groups = {}
    order = []
    for m in t.get("members", []):
        n = IS._vb_name(m["name"])
        if n.lower() not in groups:
            groups[n.lower()] = []
            order.append(n)
        groups[n.lower()].append(m)
    for n in order:
        g = groups[n.lower()]
        kind = _member_kind(g, events)
        flags = set(f for m in g for f in m.get("flags", []))
        hidden = "hidden" in flags or "restricted" in flags or _hidden_name(g[0]["name"])
        e = Entry(lib, cls_name, n, kind, help=next((m.get("helpcontext") for m in g if m.get("helpcontext")), None),
                  hidden=hidden, default=any(m.get("memid") == 0 for m in g))
        if kind == "const":
            e.value, e.returns = g[0].get("value"), g[0].get("type")
        elif kind == "prop":
            get = next((m for m in g if m.get("invoke") == "get" or m.get("kind") == "var"), None)
            put = next((m for m in g if m.get("invoke") in ("let", "set") or m.get("kind") == "var"), None)
            if get is not None:
                e.params = _params(get) if get.get("kind") != "var" else []
                e.returns = get.get("returns") if get.get("kind") != "var" else get.get("type")
            elif put is not None:
                e.params = _params(put, drop_last=True)
                ps = put.get("params", [])
                e.returns = ps[-1].get("type") if ps else None
            if get is None:
                e.access = "write-only"
            elif put is None:
                e.access = "read-only"
        else:
            m = g[0]
            e.params = _params(m)
            r = m.get("returns")
            e.returns = r if r not in NOT_SHOWN else None
        out.append(e)
    return out


def tl_classes(tl, lib, d):
    """the classes of a type library: its coclasses, modules, Enums, records and the interfaces no coclass has"""
    used = set()
    for t in d["types"]:
        if t["kind"] == "coclass":
            for imp in t.get("implements", []):
                used.add(imp["name"].split(".")[-1].lower())
    out = []
    for t in d["types"]:
        k = t["kind"]
        if k == "alias" or (k in ("interface", "dispatch") and t["name"].lower() in used):
            continue
        flags = t.get("flags", [])
        ck = {"coclass": "class", "module": "module", "enum": "enum", "record": "type"}.get(k, "interface")
        out.append(Entry(lib, None, t["name"], ck, help=t.get("helpcontext"), tl=t,
                         hidden="hidden" in flags or "restricted" in flags or t["name"].startswith("_")))
    return out


def tl_globals(tl, lib, d):
    """<globals> of a type library: the members of its modules, the constants of its Enums, the members of its
    application object (VB.Global: App, Forms, LoadPicture ...)"""
    out = []
    for t in d["types"]:
        k = t["kind"]
        if k == "module":
            for e in tl_members(tl, lib, t, t["name"]):
                if t["name"].startswith("_"):
                    e.cls = t["name"]
                out.append(e)
        elif k == "enum":
            hid = "hidden" in t.get("flags", [])
            for e in tl_members(tl, lib, t, t["name"]):
                e.hidden = e.hidden or hid
                out.append(e)
        elif k == "coclass" and "appobject" in t.get("flags", []):
            out += tl_members(tl, lib, t, t["name"])
    return out


# ====================================================================== the project
def _join_decl(lines, line):
    """the statement that starts on a line (continuations joined), without its comment"""
    i = line
    text = lines[i] if i < len(lines) else ""
    while text.rstrip().endswith(" _") and i + 1 < len(lines):
        i += 1
        text = text.rstrip()[:-1] + " " + lines[i].strip()
    return IS.strip_comment(text).strip()


SCOPE_RE = re.compile(r"^(?:(?:Public|Private|Friend|Global|Static|Dim)\s+)+", re.I)


def _proc_sig(s, lines):
    """Sub / Function / Event / Declare: its header without Public / Private; Property: Property X As T"""
    head = SCOPE_RE.sub("", _join_decl(lines, s.decls[0][0]))
    if s.kind == "proc" and (s.pkind or "").startswith("Property"):
        kinds = set()
        get_head = None
        for ln, _ in s.decls:
            h = SCOPE_RE.sub("", _join_decl(lines, ln))
            m = re.match(r"^Property\s+(Get|Let|Set)\b", h, re.I)
            if m:
                kinds.add(m.group(1).title())
                if m.group(1).lower() == "get":
                    get_head = h
        access = None if "Get" in kinds and len(kinds) > 1 else "read-only" if kinds == {"Get"} else \
            "write-only" if "Get" not in kinds else None
        h = get_head or head
        m = re.match(r"^Property\s+(Get|Let|Set)\s+(\w+)\s*(?:\((.*)\))?\s*(?:As\s+(.+))?$", h, re.I)
        if not m:
            return h, access
        kind, name, params, ret = m.groups()
        params = [p.strip() for p in IS.split_top(params or "") if p.strip()]
        if kind.lower() != "get" and params:          # Let / Set: the last parameter is the value
            last = params.pop()
            lm = re.search(r"\bAs\s+(.+)$", last, re.I)
            ret = lm.group(1) if lm else "Variant"
        sig = f"Property {name}" + (f"({', '.join(params)})" if params else "") + f" As {ret or 'Variant'}"
        return sig, access
    return head, None


def project_entries(analysis, proj_name, vb_tl):
    """(classes, members by class key, globals) of a project, from refactor.Project"""
    classes, members, glob = [], {}, []
    for mi in analysis.mods:
        m = mi.module.m
        lines = mi.lines
        kind = m.kind
        ck = "module" if mi.is_std else "class"
        ce = Entry(proj_name, None, mi.name, ck, lib_kind="project", module=m)
        classes.append(ce)
        mem = []
        for s in mi.syms.values():
            e = _sym_entry(s, proj_name, mi.name, lines)
            if e is None:
                continue
            if s.kind in ("type", "enum"):
                ce2 = Entry(proj_name, None, s.name, s.kind, lib_kind="project", sym=s, private=s.scope == "private",
                            module=m, sig=f"{'Enum' if s.kind == 'enum' else 'Type'} {s.name}")
                classes.append(ce2)
                inner = []
                val = -1
                for f in s.fields.values():
                    fe = _sym_entry(f, proj_name, s.name, lines)
                    if f.kind == "item":
                        txt = _join_decl(lines, f.decls[0][0])
                        vm = re.match(r"^\s*\w+\s*=\s*(.+)$", txt)
                        if vm:
                            v = vm.group(1).strip()
                            try:
                                val = int(v[2:], 16) if v.lower().startswith("&h") else int(v.rstrip("&"))
                            except ValueError:
                                val = v
                        else:
                            val = val + 1 if isinstance(val, int) else val
                        fe.value = val
                        fe.sig = f"Const {f.name} = {val}"
                        if s.scope == "public":
                            glob.append(Entry(proj_name, s.name, f.name, "const", value=val, sig=fe.sig, sym=f,
                                              lib_kind="project", module=m))
                    inner.append(fe)
                members[(proj_name, s.name.lower())] = inner
                continue
            if s.kind == "item":
                continue
            mem.append(e)
            if mi.is_std and s.scope == "public" and s.kind in ("proc", "var", "const", "declare"):
                g = Entry(proj_name, mi.name, s.name, e.kind, sig=e.sig, access=e.access, sym=s,
                          lib_kind="project", module=m)
                glob.append(g)
        for low, cs in mi.controls.items():
            ctl = None
            if m.sf is not None and m.sf.root is not None:
                ctl = next((c for c in m.sf.root.walk() if c.name.lower() == low), None)
            t = getattr(ctl, "type_name", None) or "Control"
            mem.append(Entry(proj_name, mi.name, cs.name, "prop", sig=f"{cs.name} As {t}", returns=t,
                             lib_kind="project", module=m, ctl=ctl, sym=cs))
        base = BASE_OF_KIND.get(getattr(m, "display_kind", None) or kind) or BASE_OF_KIND.get(kind)
        if base and vb_tl is not None:                 # a form: the members of VB.Form too
            hit = vb_tl.find(base)
            if hit is not None:
                own = {x.name.lower() for x in mem}
                for e in tl_members(vb_tl, hit[0], hit[1], mi.name):
                    if e.name.lower() not in own:
                        e.lib, e.lib_kind = proj_name, "project"
                        mem.append(e)
        members[(proj_name, mi.name.lower())] = mem
    return classes, members, glob


def _sym_entry(s, lib, cls, lines):
    if not s.decls:
        return None
    k = s.kind
    priv = s.scope == "private"
    if k == "proc":
        sig, access = _proc_sig(s, lines)
        kind = "prop" if (s.pkind or "").startswith("Property") else "method"
        return Entry(lib, cls, s.name, kind, sig=sig, access=access, sym=s, private=priv, lib_kind="project",
                     returns=s.vtype)
    if k in ("declare", "event"):
        sig, _ = _proc_sig(s, lines)
        return Entry(lib, cls, s.name, "method" if k == "declare" else "event", sig=sig, sym=s, private=priv,
                     lib_kind="project")
    if k == "var":
        sig = ("WithEvents " if s.withevents else "") + f"{s.name} As {s.vtype or 'Variant'}"
        return Entry(lib, cls, s.name, "prop", sig=sig, sym=s, private=priv, lib_kind="project", returns=s.vtype)
    if k == "const":
        ln, col = s.decls[0]
        rest = lines[ln][col + len(s.name):] if ln < len(lines) else ""
        vm = re.match(r"^\s*(?:As\s+\w+\s*)?=\s*(.*)$", IS.strip_comment(rest), re.I)
        val = IS.split_top(vm.group(1))[0].strip() if vm else ""
        return Entry(lib, cls, s.name, "const", sig=f"Const {s.name} = {val}", value=val, sym=s, private=priv,
                     lib_kind="project")
    if k == "field":
        return Entry(lib, cls, s.name, "prop", sig=f"{s.name} As {s.vtype or 'Variant'}", sym=s, private=priv,
                     lib_kind="project", returns=s.vtype)
    if k == "item":
        return Entry(lib, cls, s.name, "const", sym=s, private=priv, lib_kind="project")
    if k in ("type", "enum"):
        return Entry(lib, cls, s.name, k, sym=s, private=priv, lib_kind="project")
    return None


# ====================================================================== the details
def _vb_value(v):
    """the value of a constant as VB6 shows it: Const vbNo = 7, Const vbCrLf = "<CR LF>" (the characters)"""
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, str):
        return '"' + v.replace("\0", "") + '"'
    return str(v)


def _ret(t):
    """As T after a Function / Property; nothing for a Variant, as VB6 shows them (Function Left(String, ...))"""
    t = IS.short(t) if t else None
    return f" As {t}" if t and t.lower() != "variant" else ""


def signature(e):
    """the declaration line of the details pane"""
    if e.kind in CLASS_WORD:
        return f"{CLASS_WORD[e.kind]} {e.name}".strip() if e.kind != "globals" else GLOBALS
    if e.sig:
        return e.sig
    if e.kind == "const":
        return f"Const {e.name} = {_vb_value(e.value)}" if e.value is not None else f"Const {e.name}"
    ps = ", ".join(IS.p_text(p) for p in e.params)
    if e.kind == "event":
        return f"Event {e.name}({ps})"
    if e.kind == "prop":
        return f"Property {e.name}" + (f"({ps})" if ps else "") + _ret(e.returns)
    if e.returns:
        return f"Function {e.name}({ps})" + _ret(e.returns)
    return f"Sub {e.name}({ps})"


class ObjectBrowser(QWidget):
    """the window"""
    closed = pyqtSignal()

    def __init__(self, main):
        super().__init__()
        self.main = main
        self.show_hidden = False
        self.group = False
        self.whole_word = False
        self.history, self.hpos = [], -1
        self._busy = False
        self.libs = {}           # name -> {'kind', 'classes': [Entry], 'globals': [Entry], 'members': {}}
        self.tl = None
        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        lay.setSpacing(2)
        row = QHBoxLayout()
        self.lib_combo = QComboBox()
        self.lib_combo.setMinimumWidth(180)
        self.lib_combo.currentIndexChanged.connect(lambda i: self._lib_changed())
        row.addWidget(self.lib_combo)
        self.b_back = self._button("<", "Go Back", self.back)
        self.b_fwd = self._button(">", "Go Forward", self.forward)
        self.b_copy = self._button("Copy", "Copy to Clipboard", self.copy)
        self.b_def = self._button("Definition", "View Definition", self.view_definition)
        self.b_help = self._button("?", "Help", self.help)
        for b in (self.b_back, self.b_fwd, self.b_copy, self.b_def, self.b_help):
            row.addWidget(b)
        row.addStretch(1)
        lay.addLayout(row)
        row = QHBoxLayout()
        self.search_box = QComboBox()
        self.search_box.setEditable(True)
        self.search_box.setMinimumWidth(180)
        self.search_box.lineEdit().returnPressed.connect(lambda: self.search())
        row.addWidget(self.search_box)
        self.b_search = self._button("Search", "Search", lambda: self.search())
        self.b_results = self._button("Results", "Show/Hide Search Results", self.toggle_results)
        row.addWidget(self.b_search)
        row.addWidget(self.b_results)
        row.addStretch(1)
        lay.addLayout(row)
        self.results = QTreeWidget()
        self.results.setHeaderLabels(["Library", "Class", "Member"])
        self.results.setRootIsDecorated(False)
        self.results.itemActivated.connect(self._result_chosen)
        self.results.itemClicked.connect(self._result_chosen)
        self.results.setVisible(False)
        self.results.setMaximumHeight(140)
        lay.addWidget(self.results)
        split = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.addWidget(QLabel("Classes"))
        self.class_list = QListWidget()
        self.class_list.currentItemChanged.connect(lambda cur, prev: self._class_changed())
        ll.addWidget(self.class_list)
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        self.members_label = QLabel("Members")
        rl.addWidget(self.members_label)
        self.member_list = QListWidget()
        self.member_list.currentItemChanged.connect(lambda cur, prev: self._member_changed())
        self.member_list.itemDoubleClicked.connect(lambda it: self.view_definition())
        rl.addWidget(self.member_list)
        split.addWidget(left)
        split.addWidget(right)
        split.setSizes([220, 380])
        for w in (self.class_list, self.member_list):
            w.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            w.customContextMenuRequested.connect(lambda pos, ww=w: self._menu(ww, pos))
            w.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        lay.addWidget(split, 1)
        self.details = QTextBrowser()
        self.details.setOpenLinks(False)
        self.details.anchorClicked.connect(self._link)
        self.details.setMaximumHeight(90)
        lay.addWidget(self.details)
        self.refresh()

    def _button(self, text, tip, slot):
        b = QToolButton()
        b.setText(text)
        b.setToolTip(tip)
        b.setAutoRaise(True)
        b.clicked.connect(lambda checked=False: slot())
        return b

    # ------------------------------------------------------------ the data
    def refresh(self):
        """the libraries again (the code of the project may have changed)"""
        keep = self.state()
        self.libs = {}
        main = self.main
        projects = list(main.group.projects) if getattr(main, "group", None) else \
            ([main.project] if getattr(main, "project", None) is not None else [])
        extra = []
        for p in projects:
            for lib in IS.IntelliSense(p).tl.libs:
                if lib not in extra:
                    extra.append(lib)
        self.tl = IS.TypeLibs(extra)
        from . import refactorui as RF
        for p in projects:
            try:
                an = RF.analyze(main, p)
            except Exception:                    # noqa: BLE001 - code that does not parse: no project library
                continue
            cls, mem, glob = project_entries(an, p.name, self.tl)
            self.libs[p.name] = {"kind": "project", "classes": cls, "globals": glob, "members": mem}
        for lib, d in self.tl.libs.items():
            self.libs[lib] = {"kind": "typelib", "classes": tl_classes(self.tl, lib, d), "data": d,
                              "globals": None, "members": {}}
        self._busy = True
        self.lib_combo.clear()
        self.lib_combo.addItem(ALL)
        for n in sorted(self.libs, key=str.lower):
            self.lib_combo.addItem(n)
        self._busy = False
        if keep and keep[0] in [self.lib_combo.itemText(i) for i in range(self.lib_combo.count())]:
            self.restore(keep, record=False)
        else:
            self._lib_changed()

    def library_names(self):
        return [self.lib_combo.itemText(i) for i in range(self.lib_combo.count())]

    def _globals(self, lib):
        L = self.libs[lib]
        if L["globals"] is None:
            L["globals"] = tl_globals(self.tl, lib, L["data"])
        return L["globals"]

    def classes(self, lib_sel=None):
        """the classes the left list shows for a library (or <All Libraries>)"""
        lib_sel = lib_sel or self.lib_combo.currentText() or ALL
        libs = list(self.libs) if lib_sel == ALL else [lib_sel] if lib_sel in self.libs else []
        out = [Entry(lib_sel, None, GLOBALS, "globals")]
        for lib in libs:
            out += [c for c in self.libs[lib]["classes"] if self.show_hidden or not c.hidden]
        head, rest = out[:1], out[1:]
        rest.sort(key=lambda c: (c.name.lower(), c.lib.lower()))
        return head + rest

    def members(self, cls):
        """the members of a class Entry"""
        if cls is None:
            return []
        if cls.kind == "globals":
            libs = list(self.libs) if cls.lib == ALL else [cls.lib]
            out = []
            for lib in libs:
                L = self.libs.get(lib)
                if L is None:
                    continue
                out += L["globals"] if L["kind"] == "project" else self._globals(lib)
        else:
            L = self.libs.get(cls.lib)
            if L is None:
                return []
            k = (cls.lib, cls.name.lower())
            if k not in L["members"]:
                L["members"][k] = tl_members(self.tl, cls.lib, cls.tl, cls.name) if cls.tl is not None else []
            out = L["members"][k]
        out = [m for m in out if self.show_hidden or not m.hidden]
        uniq = {}
        for m in out:
            uniq.setdefault((m.name.lower(), m.kind), m)
        res = list(uniq.values())
        if self.group:
            res.sort(key=lambda m: (MEMBER_ORDER.get(m.kind, 9), m.name.lower()))
        else:
            res.sort(key=lambda m: m.name.lower())
        for m in res:
            if m.parent is None:
                m.parent = cls
        return res

    # ------------------------------------------------------------ the lists
    def _lib_changed(self):
        if self._busy:
            return
        self._fill_classes()
        self._record()

    def _fill_classes(self, select=None):
        self._busy = True
        self.class_list.clear()
        lib_sel = self.lib_combo.currentText() or ALL
        for c in self.classes(lib_sel):
            it = QListWidgetItem(icon(c.kind, c.private), c.name)
            it.setData(Qt.ItemDataRole.UserRole, c)
            if c.hidden:
                it.setForeground(QBrush(QColor("#808080")))
            if lib_sel == ALL and c.kind != "globals":
                it.setToolTip(f"{c.lib}.{c.name}")
            self.class_list.addItem(it)
        self._busy = False
        row = 0
        if select is not None:
            for i in range(self.class_list.count()):
                c = self.class_list.item(i).data(Qt.ItemDataRole.UserRole)
                if c.name.lower() == select[1].lower() and (select[0] in (None, c.lib) or c.kind == "globals"):
                    row = i
                    break
        self.class_list.setCurrentRow(row)
        self._class_changed()

    def current_class(self):
        it = self.class_list.currentItem()
        return it.data(Qt.ItemDataRole.UserRole) if it is not None else None

    def current_member(self):
        it = self.member_list.currentItem()
        return it.data(Qt.ItemDataRole.UserRole) if it is not None else None

    def current(self):
        return self.current_member() or self.current_class()

    def _class_changed(self):
        if self._busy:
            return
        c = self.current_class()
        self._busy = True
        self.member_list.clear()
        if c is not None:
            self.members_label.setText(f"Members of '{c.name}'")
            for m in self.members(c):
                it = QListWidgetItem(icon(m.kind, m.private), m.name)
                it.setData(Qt.ItemDataRole.UserRole, m)
                f = it.font()
                if m.default:
                    f.setBold(True)
                if m.private:
                    f.setItalic(True)
                it.setFont(f)
                if m.hidden:
                    it.setForeground(QBrush(QColor("#808080")))
                self.member_list.addItem(it)
        self._busy = False
        self._show_details()
        self._record()

    def _member_changed(self):
        if self._busy:
            return
        self._show_details()
        self._record()

    def member_names(self):
        return [self.member_list.item(i).text() for i in range(self.member_list.count())]

    def class_names(self):
        return [self.class_list.item(i).text() for i in range(self.class_list.count())]

    # ------------------------------------------------------------ the details
    def details_text(self):
        return self.details.toPlainText()

    def _type_link(self, t):
        t = IS.short(t) if t else t
        if not t:
            return ""
        for lib, L in self.libs.items():
            for c in L["classes"]:
                if c.name.lower() == t.lower() and c.kind != "globals":
                    return f'<a href="cls:{lib}.{c.name}">{t}</a>'
        return t

    def _sig_html(self, e):
        """the declaration as VB6 shows it: the name bold, the parameters italic, the types links"""
        esc = lambda t: t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")   # noqa: E731
        sig = signature(e)
        if e.kind in CLASS_WORD:
            word, _, name = sig.rpartition(" ")
            return (esc(word) + " " if word else "") + f"<b>{esc(name)}</b>"
        k = sig.find(e.name)
        if k < 0:
            return esc(sig)
        head, rest = sig[:k], sig[k + len(e.name):]
        html_rest = esc(rest)
        if rest.startswith("(") and not e.sig:
            depth, end = 0, -1
            for n, ch in enumerate(rest):
                depth += ch == "("
                depth -= ch == ")"
                if depth == 0:
                    end = n
                    break
            if end > 0:
                html_rest = "(<i>" + esc(rest[1:end]) + "</i>)" + esc(rest[end + 1:])
        html_rest = re.sub(r"\bAs (\w+)", lambda m: "As " + self._type_link(m.group(1)), html_rest)
        html_rest = html_rest.replace("\r\n", "<br>").replace("\r", "<br>").replace("\n", "<br>")
        return esc(head) + f"<b>{esc(e.name)}</b>" + html_rest

    def _show_details(self):
        e = self.current()
        if e is None:
            self.details.setHtml("")
            return
        ind = "&nbsp;&nbsp;&nbsp;&nbsp;"
        if e.kind == "globals":                    # VB6: the library itself
            lines = []
            for lib in ([e.lib] if e.lib in self.libs else []):
                L = self.libs[lib]
                if L["kind"] == "project":
                    lines = [f"Project <b>{lib}</b>"]
                else:
                    lines = [f"Library <b>{lib}</b>", ind + L["data"].get("file", "")]
            self.details.setHtml("<br>".join(lines))
            return
        lines = [self._sig_html(e)]
        if e.access:
            lines.append(ind + e.access)
        cls = e.parent if e.kind not in CLASS_WORD else None
        if e.kind in CLASS_WORD:
            lines.append(f'{ind}Member of <a href="lib:{e.lib}">{e.lib}</a>')
        else:
            owner = e.cls or (cls.name if cls is not None else "")
            word = "Default member of" if e.default else "Member of"
            if owner and not owner.startswith("_"):
                lines.append(f'{ind}{word} <a href="lib:{e.lib}">{e.lib}</a>.'
                             f'<a href="cls:{e.lib}.{owner}">{owner}</a>')
            else:
                lines.append(f'{ind}{word} <a href="lib:{e.lib}">{e.lib}</a>')
        if e.private:
            lines.append(ind + "Private")
        self.details.setHtml("<br>".join(lines))

    def _link(self, url):
        s = url.toString()
        if s.startswith("lib:"):
            self.set_library(s[4:])
        elif s.startswith("cls:"):
            lib, _, name = s[4:].partition(".")
            self.select_class(name, lib)

    # ------------------------------------------------------------ navigation
    def set_library(self, name):
        i = self.lib_combo.findText(name)
        if i < 0:
            return False
        if self.lib_combo.currentIndex() == i:
            self._fill_classes()
            self._record()
        else:
            self.lib_combo.setCurrentIndex(i)     # (_lib_changed fills the classes)
        return True

    def select_class(self, name, lib=None):
        """the class, in the library chosen or (when it is not there) in its own"""
        found = None
        for i in range(self.class_list.count()):
            c = self.class_list.item(i).data(Qt.ItemDataRole.UserRole)
            if c.name.lower() == name.lower() and (lib is None or c.lib == lib or c.kind == "globals"):
                found = i
                break
        if found is None and lib is not None and lib in self.libs:
            self._busy = True
            self.lib_combo.setCurrentIndex(self.lib_combo.findText(lib))
            self._busy = False
            self._fill_classes((lib, name))
            return self.current_class() is not None and self.current_class().name.lower() == name.lower()
        if found is None:
            return False
        self.class_list.setCurrentRow(found)
        return True

    def select_member(self, name):
        for i in range(self.member_list.count()):
            if self.member_list.item(i).text().lower() == name.lower():
                self.member_list.setCurrentRow(i)
                return True
        return False

    def reveal(self, lib, name, cls=None):
        """Definition on a name of a library (MsgBox, Text1.Text): its class and the member"""
        if cls and lib in self.libs:
            self.set_library(lib)
            if self.select_class(cls, lib):
                return self.select_member(name) or self.current_class().name.lower() == name.lower()
        for c in self.classes(lib):
            if c.kind == "globals":
                continue
            if c.name.lower() == name.lower():
                self.set_library(lib)
                return self.select_class(c.name, lib)
        for c in self.classes(lib):
            if c.kind == "globals":
                continue
            if any(m.name.lower() == name.lower() for m in self.members(c)):
                self.set_library(lib)
                self.select_class(c.name, lib)
                return self.select_member(name)
        return False

    def state(self):
        if not self.lib_combo.count():
            return None
        c, m = self.current_class(), self.current_member()
        return (self.lib_combo.currentText(), c.lib if c else None, c.name if c else None, m.name if m else None)

    def restore(self, st, record=True):
        self._busy = True
        self.lib_combo.setCurrentIndex(max(0, self.lib_combo.findText(st[0])))
        self._busy = False
        self._fill_classes((st[1], st[2]) if st[2] else None)
        if st[3]:
            self.select_member(st[3])
        if record:
            self._record()

    def _record(self):
        if self._busy:
            return
        st = self.state()
        if st is None:
            return
        if 0 <= self.hpos < len(self.history) and self.history[self.hpos][:3] == st[:3] and \
                (self.history[self.hpos][3] is None or st[3] is None or self.history[self.hpos] == st):
            self.history[self.hpos] = st
        else:
            del self.history[self.hpos + 1:]
            self.history.append(st)
            self.hpos = len(self.history) - 1
        self._update_buttons()

    def _update_buttons(self):
        self.b_back.setEnabled(self.hpos > 0)
        self.b_fwd.setEnabled(self.hpos < len(self.history) - 1)
        e = self.current()
        self.b_def.setEnabled(e is not None and e.lib_kind == "project" and (e.sym is not None or e.module is not None))

    def back(self):
        if self.hpos > 0:
            self.hpos -= 1
            self._replay()

    def forward(self):
        if self.hpos < len(self.history) - 1:
            self.hpos += 1
            self._replay()

    def _replay(self):
        st = self.history[self.hpos]
        h, p = self.history, self.hpos
        self.history = []
        self.restore(st, record=False)
        self.history, self.hpos = h, p
        self._update_buttons()

    # ------------------------------------------------------------ commands
    def copy(self):
        e = self.current()
        if e is not None:
            QApplication.clipboard().setText(e.name if e.kind != "globals" else "")
        return e.name if e is not None else None

    def view_definition(self):
        """the declaration of a member of the project (or the code of a module of it)"""
        e = self.current()
        if e is None or e.lib_kind != "project":
            return False
        from . import refactorui as RF
        main = self.main
        if e.ctl is not None and e.module is not None:
            d = main.open_designer(e.module)
            if d is not None and hasattr(d, "select_control"):
                d.select_control(e.ctl)
            return True
        s = e.sym
        if s is not None and s.decls and s.mod is not None:
            ln, col = s.decls[0]
            RF.select_range(main, s.mod.module.m, ln, col, len(s.name))
            return True
        if e.module is not None and e.kind in ("class", "module"):
            main.open_code(e.module, 0)
            return True
        return False

    def help(self):
        e = self.current()
        if e is None:
            return False
        lib = e.lib if e.lib_kind != "project" else "VB"
        if e.help and self.main.help_context_number(e.help, lib):
            return True
        self.main.help_keyword(e.name)
        return False

    def toggle_results(self):
        self.results.setVisible(not self.results.isVisible())

    def search(self, text=None):
        """the classes and members whose names match, in the library chosen: [(lib, class, member)]"""
        text = (text if text is not None else self.search_box.currentText()).strip()
        if text and self.search_box.findText(text) < 0:
            self.search_box.insertItem(0, text)
        self.search_box.setCurrentText(text)
        if not text:
            return []
        pat = text.lower()
        wild = "*" in pat or "?" in pat

        def ok(n):
            n = n.lower()
            if wild:
                return fnmatch.fnmatchcase(n, pat)
            return n == pat if self.whole_word else pat in n
        rows = []
        for c in self.classes():
            if c.kind != "globals" and ok(c.name):
                rows.append((c.lib, c.name, ""))
            for m in self.members(c):
                if ok(m.name):
                    rows.append((m.lib, c.name if c.kind != "globals" else GLOBALS, m.name))
        seen, out = set(), []
        for r in rows:
            if r not in seen:
                seen.add(r)
                out.append(r)
        out.sort(key=lambda r: (r[2].lower() != pat, r[0].lower(), r[1].lower(), r[2].lower()))
        self.results.clear()
        for lib, cls, mem in out:
            it = QTreeWidgetItem([lib, cls, mem])
            self.results.addTopLevelItem(it)
        self.results.setVisible(True)
        return out

    def _result_chosen(self, it, col=0):
        lib, cls, mem = it.text(0), it.text(1), it.text(2)
        if not self.select_class(cls, lib if cls != GLOBALS else None):
            return
        if mem:
            self.select_member(mem)

    def _menu(self, w, pos):
        m = QMenu(self)
        m.addAction("&Copy", self.copy)
        a = m.addAction("View &Definition", self.view_definition)
        a.setEnabled(self.b_def.isEnabled())
        m.addSeparator()
        a = m.addAction("Group &Members", lambda: self.set_group(not self.group))
        a.setCheckable(True)
        a.setChecked(self.group)
        a = m.addAction("Show &Hidden Members", lambda: self.set_show_hidden(not self.show_hidden))
        a.setCheckable(True)
        a.setChecked(self.show_hidden)
        a = m.addAction("&Find Whole Word Only", lambda: setattr(self, "whole_word", not self.whole_word))
        a.setCheckable(True)
        a.setChecked(self.whole_word)
        m.addSeparator()
        m.addAction("He&lp", self.help)
        m.exec(w.mapToGlobal(pos))

    def set_show_hidden(self, on):
        st = self.state()
        self.show_hidden = bool(on)
        self.restore(st, record=False)

    def set_group(self, on):
        st = self.state()
        self.group = bool(on)
        self.restore(st, record=False)

    def showEvent(self, e):
        super().showEvent(e)
        self._update_buttons()


class ObjectBrowserWindow(QMdiSubWindow):
    def __init__(self, main):
        super().__init__()
        self.main = main

    def closeEvent(self, e):
        self.main._objbrowser = None
        super().closeEvent(e)


# ====================================================================== icons (drawn here)
_ICONS = {}
ICON_STYLE = {"globals": ("#5a5a5a", "G"), "class": ("#2b7a78", "C"), "interface": ("#2b7a78", "I"),
              "module": ("#6b6b9a", "M"), "enum": ("#b8860b", "E"), "type": ("#3c6fb6", "T"),
              "prop": ("#4b7f3a", "P"), "method": ("#7a3fa0", "M"), "event": ("#c0782a", "E"),
              "const": ("#3c6fb6", "c")}


def icon(kind, private=False):
    key = (kind, private)
    if key not in _ICONS:
        color, letter = ICON_STYLE.get(kind, ("#808080", "?"))
        pm = QPixmap(QSize(16, 16))
        pm.fill(Qt.GlobalColor.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        c = QColor(color)
        if private:
            c = c.lighter(150)
        p.setBrush(c)
        p.setPen(Qt.PenStyle.NoPen)
        if kind in ("prop", "method", "event", "const"):
            p.drawEllipse(1, 1, 14, 14)
        else:
            p.drawRoundedRect(1, 1, 14, 14, 3, 3)
        f = QFont()
        f.setPixelSize(10)
        f.setBold(True)
        p.setFont(f)
        p.setPen(QColor("white"))
        p.drawText(pm.rect(), Qt.AlignmentFlag.AlignCenter, letter)
        p.end()
        _ICONS[key] = QIcon(pm)
    return _ICONS[key]
