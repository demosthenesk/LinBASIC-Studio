"""The object model of the IDE for the add-ins of the user (phase 12.5): VBIDE (Microsoft Visual Basic 6.0
Extensibility) and the command bars of Office, as an add-in sees them - VBE, its projects, components, code modules,
code panes, windows, command bars and the events of the command bars. Our own code, written from the members of
VBIDE (data/vb6/typelibs/VBIDE.json) and from what VB6 does with them; an add-in reaches these objects through its host
(runtime/addin.c, addinhost.py): a member "Get" of an object is get_<member>, a method call_<member>, "Let" / "Set"
let_<member> (names in lower case, as the compiler gives them)."""
import os
import re

from PyQt6.QtGui import QAction, QIcon, QPixmap
from PyQt6.QtWidgets import QApplication, QMenu, QToolBar


class AddInError(Exception):
    """a run-time error for the add-in: its number and description (Err of the add-in)"""

    def __init__(self, number, description):
        super().__init__(f"{description} ({number})")
        self.number = number
        self.description = description


def not_found():
    return AddInError(9, "Subscript out of range")


class Nothing:
    """Nothing (an object variable not set)"""


NOTHING = Nothing()

E_INVALIDARG = -2147024809                      # VB6's "Invalid procedure call or argument" of VBIDE (oracle evlog)


class Out:
    """a value and the new values of ByRef arguments (Find: StartLine ...; ProcOfLine: ProcKind): {index: value}"""

    def __init__(self, value, back):
        self.value = value
        self.back = back


def ide_changed(vbe):
    """after a member of the add-in changed the IDE: the events of the change for the add-ins (addinhost.sync_all)"""
    from .addinhost import sync_all
    sync_all(vbe.main)

# vbext_ComponentType
CT = {"Module": 1, "Class": 2, "Form": 5, "UserControl": 8, "PropertyPage": 7, "Designer": 11}
CT_KIND = {1: "Module", 2: "Class", 5: "Form", 8: "UserControl", 7: "PropertyPage"}
# vbext_ProjectType
PT = {"Exe": 0, "OleExe": 1, "OleDll": 2, "Control": 3}
# vbext_ProcKind
PK_PROC, PK_LET, PK_SET, PK_GET = 0, 1, 2, 3
# vbext_FileType
FT = {"Form": 0, "Module": 1, "Class": 2, "UserControl": 7, "PropertyPage": 8, "UserDocument": 9, "Designer": 12}
FT_PROJECT, FT_GROUP = 3, 11
# the references every project has, not written in the .vbp (VB6: References(1..3), BuiltIn True)
BUILTIN_REFS = ["{000204EF-0000-0000-C000-000000000046}", "{EA544A21-C82D-11D1-A3E4-00A0C90AEA82}",
                "{FCFB3D2E-A0FA-1068-A738-08002B3371B5}"]


def file_type(m):
    return FT.get(m.kind, 0)


_TYPELIBS = None


def typelibs():
    """the type libraries LinBASIC Studio knows (data/vb6/typelibs/index.json)"""
    global _TYPELIBS
    if _TYPELIBS is None:
        import json
        from . import paths
        try:
            with open(os.path.join(paths.VB6_DATA, "typelibs", "index.json"), encoding="utf-8") as f:
                _TYPELIBS = json.load(f)
        except (OSError, ValueError):
            _TYPELIBS = []
    return _TYPELIBS


def typelib_by_guid(guid):
    g = str(guid).strip().lower()
    return next((t for t in typelibs() if t["guid"].lower() == g), None)


def reference_line(lib):
    """the Reference= line of the .vbp for a library, as VB6 writes it"""
    return f"*\\G{lib['guid']}#{lib['version']}#0#{lib['file']}#{lib.get('library') or lib['short']}"

_HEAD = re.compile(r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(Get|Let|Set))\s+"
                   r"([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
_BLANK = re.compile(r"^\s*(?:'.*|Rem(?:\s.*)?)?$", re.IGNORECASE)


def _first_proc(lines):
    return next((i for i, ln in enumerate(lines) if _HEAD.match(ln)), len(lines))


def declaration_lines(lines):
    """CountOfDeclarationLines: the lines before the first procedure, without the comments and empty lines just
    before it (they belong to the procedure)"""
    i = _first_proc(lines)
    while i > 0 and _BLANK.match(lines[i - 1]):
        i -= 1
    return i


def procedures(lines):
    """[(name, vbext_ProcKind, start line, body line, end line)], lines from 1: a procedure begins after the one
    before it (its comments), as ProcStartLine of VB6"""
    out = []
    prev_end = declaration_lines(lines)
    i = 0
    while i < len(lines):
        m = _HEAD.match(lines[i])
        if not m:
            i += 1
            continue
        kind = {None: PK_PROC, "get": PK_GET, "let": PK_LET, "set": PK_SET}[(m.group(2) or "").lower() or None]
        j = i
        while j < len(lines) and not _END.match(lines[j]):
            j += 1
        out.append((m.group(3), kind, prev_end + 1, i + 1, min(j, len(lines) - 1) + 1))
        prev_end = j + 1
        i = j + 1
    return out


class VBObject:
    """an object of the IDE for the add-ins"""
    TYPENAME = "Object"

    def __init__(self, vbe):
        self.vbe = vbe

    def typename(self):
        return self.TYPENAME

    def get_vbe(self):
        return self.vbe


class Collection(VBObject):
    """a collection of VBIDE: Count, Item (by number from 1 or by name), the default member"""

    def items(self):
        return []

    def name_of(self, it):
        return getattr(it, "name_", None)

    def get_count(self):
        return len(self.items())

    def call_item(self, index):
        its = self.items()
        if isinstance(index, (int, float)) and not isinstance(index, bool):
            i = int(index)
            if 1 <= i <= len(its):
                return its[i - 1]
            raise not_found()
        key = str(index).lower()
        for it in its:
            n = self.name_of(it)
            if n is not None and n.lower() == key:
                return it
        raise not_found()

    get_item = call_item

    def get_parent(self):
        return self.vbe


# ---------------------------------------------------------------- VBE
class VBE(VBObject):
    TYPENAME = "VBE"

    def __init__(self, main, host=None):
        super().__init__(self)
        self.main = main
        self.host = host                        # the add-in that has this VBE (addinhost.AddInHost)
        self._wrap = {}                         # an object of the IDE -> its object for the add-ins (identity)
        self._events = Events(self)

    def form_designer(self, m):
        """the designer of a form / user control (opened as VB6 has it open for a new form)"""
        main = self.main
        sw = getattr(main, "design_windows", {}).get(m)
        if sw is None:
            main.open_designer(m)
            sw = getattr(main, "design_windows", {}).get(m)
        d = sw.widget() if sw is not None else None
        return d if d is not None and hasattr(d, "place_new") else None

    def wrap(self, cls, key, *args):
        """one object for the add-ins for each thing of the IDE (VBInstance.ActiveVBProject Is ...)"""
        k = (cls, id(key)) + tuple(a for a in args if isinstance(a, (str, int)))   # (Events: one object each)
        o = self._wrap.get(k)
        if o is None or getattr(o, "_key", None) is not key:
            o = cls(self, key, *args)
            o._key = key
            self._wrap[k] = o
        return o

    def projects(self):
        if self.host is not None and getattr(self.host, "hide_projects", False):
            return []
        m = self.main
        g = getattr(m, "group", None)
        if g is not None:
            return list(g.projects)
        return [m.project] if getattr(m, "project", None) is not None else []

    def get_fullname(self):
        import sys
        return os.path.abspath(sys.argv[0]) if sys.argv and sys.argv[0] else "lbstudio"

    def get_version(self):
        return "6.00"

    def get_name(self):
        return "Microsoft Visual Basic"

    def get_activevbproject(self):
        p = getattr(self.main, "project", None)
        return self.wrap(VBProject, p) if p is not None and self.projects() else NOTHING

    def let_activevbproject(self, proj):
        """(VB6: no ItemActivated for it, oracle evlog)"""
        if isinstance(proj, VBProject) and proj.p in self.projects():
            from .addinhost import quiet_active, hold, sync_all
            quiet_active(self.main)
            with hold(self.main):
                self.main.set_active(proj.p)
                sel = getattr(self.main, "select_project_in_tree", None)
                if sel is not None:             # (the project, not a component of it: no ItemSelected)
                    sel(proj.p)
            sync_all(self.main)

    def get_vbprojects(self):
        return self.wrap(VBProjects, self.main)

    def get_commandbars(self, *args):
        cb = self.wrap(CommandBars, self.main)
        return cb.call_item(args[0]) if args else cb

    def get_events(self):
        return self._events

    def get_windows(self, *args):
        w = self.wrap(Windows, self.main)
        return w.call_item(args[0]) if args else w

    def get_mainwindow(self):
        return self.wrap(Window, self.main, "main")

    def get_activewindow(self):
        sw = self.main.mdi.activeSubWindow() if hasattr(self.main, "mdi") else None
        return self.wrap(Window, sw) if sw is not None else NOTHING

    def get_codepanes(self, *args):
        c = self.wrap(CodePanes, self.main)
        return c.call_item(args[0]) if args else c

    def get_activecodepane(self):
        cw = self._active_code()
        return self.wrap(CodePane, cw.module) if cw is not None else NOTHING

    def _active_code(self):
        from .codewin import CodeWindow
        sw = self.main.mdi.activeSubWindow() if hasattr(self.main, "mdi") else None
        w = sw.widget() if sw is not None else None
        return w if isinstance(w, CodeWindow) else None

    def get_selectedvbcomponent(self):
        m = getattr(self.main, "selected_module", lambda: None)()
        if m is None:
            cw = self._active_code()
            m = cw.module if cw is not None else None
        return self.wrap(VBComponent, m) if m is not None else NOTHING

    def get_addins(self, *args):
        a = self.wrap(AddIns, self.main)
        return a.call_item(args[0]) if args else a

    def get_lastusedpath(self):
        return self.main.S["last_dir"] or ""

    def let_lastusedpath(self, v):
        self.main.S["last_dir"] = str(v)

    def get_displaymodel(self):
        return 0                                # vbext_dm_SDI? VB6 gives 1 for MDI
    get_readonlymode = lambda self: 0           # noqa: E731

    def call_quit(self):
        self.main.close()


# ---------------------------------------------------------------- projects and components
class VBProjects(Collection):
    TYPENAME = "VBProjects"

    def __init__(self, vbe, main):
        super().__init__(vbe)
        self.main = main

    def items(self):
        return [self.vbe.wrap(VBProject, p) for p in self.vbe.projects()]

    def name_of(self, it):
        return it.p.name

    def get_startproject(self):
        g = getattr(self.main, "group", None)
        p = g.startup if g is not None else getattr(self.main, "project", None)
        return self.vbe.wrap(VBProject, p) if p is not None else NOTHING

    def call_add(self, ptype, exclusive=False):
        """Add(vbext_pt_StandardExe ...): a new project in the group, the active one (VB6: its components
        ItemAdded, then the project's ItemAdded - no ItemActivated)"""
        kind = {0: "Exe", 1: "OleExe", 2: "OleDll", 3: "Control"}.get(int(ptype))
        if kind is None:
            raise AddInError(E_INVALIDARG, "Invalid procedure call or argument")
        main = self.main
        from .addinhost import quiet_active, hold
        with hold(main):
            if not getattr(main, "project", None):
                main.new_project(kind)
            else:
                main.add_project(kind)
        p = main.project
        quiet_active(main)
        ide_changed(self.vbe)
        return self.vbe.wrap(VBProject, p)

    def call_remove(self, proj):
        """Remove: VB6 writes its files that were never saved first (in the current folder, with the events of
        the files: oracle evlog), then the project goes without a question"""
        if not isinstance(proj, VBProject) or proj.p not in self.vbe.projects():
            raise AddInError(E_INVALIDARG, "Invalid procedure call or argument")
        p = proj.p
        main = self.main
        if p.untitled:
            base = os.getcwd()
            p.save(os.path.join(base, p.name + ".vbp"))
        else:
            for m in p.modules:
                if getattr(m, "never_saved", False) or not os.path.exists(m.path):
                    m.save()
        from .addinhost import hold
        with hold(main):
            main.remove_project(p, ask=False)
            sel = getattr(main, "select_project_in_tree", None)
            if sel is not None and main.project is not None:
                sel(main.project)               # (the active project, not a component of it)
        ide_changed(self)


class VBProject(VBObject):
    TYPENAME = "VBProject"

    def __init__(self, vbe, p):
        super().__init__(vbe)
        self.p = p

    def get_name(self):
        return self.p.name

    def let_name(self, v):
        import re
        v = str(v)
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,39}", v):
            raise AddInError(380, "Invalid property value")
        self.p.vbp.set("Name", v, quote=True)
        self.p.dirty = True
        self.vbe.main.fill_tree()
        self.vbe.main.update_title()

    def get_description(self):
        return (self.p.vbp.get("Description") or "").strip('"')

    def let_description(self, v):
        self.p.vbp.set("Description", str(v), quote=True)
        self.p.dirty = True

    def get_filename(self):
        return self.p.path or ""

    def get_buildfilename(self):
        return (self.p.vbp.get("ExeName32") or "").strip('"') if self.p.vbp else ""

    def get_type(self):
        return PT.get(self.p.ptype, 0)

    def get_saved(self):
        return not self.p.any_dirty()

    def get_isdirty(self):
        return self.p.any_dirty()

    def get_vbcomponents(self, *args):
        c = self.vbe.wrap(VBComponents, self.p)
        return c.call_item(args[0]) if args else c

    def get_collection(self):
        return self.vbe.get_vbprojects()

    def get_mode(self):
        st = getattr(getattr(self.vbe.main, "runner", None), "state", "design")
        return {"design": 2, "running": 0, "break": 1}.get(st, 2)   # vbext_vm_Run 0, vbext_vm_Break 1, vbext_vm_Design 2

    def get_references(self):
        return self.vbe.wrap(References, self.p)


class References(Collection):
    TYPENAME = "References"

    def __init__(self, vbe, p):
        super().__init__(vbe)
        self.p = p

    def items(self):
        out = [self.vbe.wrap(Reference, b, "BuiltIn") for b in BUILTIN_REFS]
        for k, v in self.p.vbp.entries:
            if v is not None and k.strip() == "Reference":
                out.append(self.vbe.wrap(Reference, v, k.strip()))
        return out

    def name_of(self, it):
        return it.get_name()

    def get_parent(self):
        return self.vbe.wrap(VBProject, self.p)

    def _add(self, lib):
        line = reference_line(lib)
        have = [x for x in self.items() if x.get_guid().lower() == lib["guid"].lower()]
        if have:
            raise AddInError(32813, "Name conflicts with existing module, project, or object library")
        self.p.add_references([line])
        self.p.dirty = True
        ide_changed(self.vbe)
        return next(x for x in self.items() if x.get_guid().lower() == lib["guid"].lower())

    def call_addfromguid(self, guid, major, minor):
        lib = typelib_by_guid(str(guid))
        if lib is None:
            raise AddInError(-2147319779, "Library not registered.")
        return self._add(lib)

    def call_addfromfile(self, path):
        name = os.path.basename(str(path).replace("\\", "/")).lower()
        lib = next((t for t in typelibs() if t["file"].lower() == name), None)
        if lib is None:
            raise AddInError(48, "Error in loading DLL")
        return self._add(lib)

    def call_remove(self, ref):
        if not isinstance(ref, Reference) or ref.kind != "Reference" or ref.line not in [
                v for k, v in self.p.vbp.entries if v is not None and k.strip() == "Reference"]:
            raise AddInError(E_INVALIDARG, "Invalid procedure call or argument")
        self.p.vbp.entries = [[k, v] for k, v in self.p.vbp.entries if not (k.strip() == "Reference" and v == ref.line)]
        self.p.dirty = True
        ide_changed(self.vbe)


class Reference(VBObject):
    TYPENAME = "Reference"

    def __init__(self, vbe, line, kind):
        super().__init__(vbe)
        self.line, self.kind = line, kind       # kind: Reference (a line of the .vbp) / BuiltIn (a GUID)

    def _parts(self):
        return self.line.split("#") if self.kind != "BuiltIn" else [self.line]

    def _lib(self):
        return typelib_by_guid(self.get_guid())

    def get_guid(self):
        g = self._parts()[0]
        return g[g.find("{"):] if "{" in g else g

    def get_fullpath(self):
        ps = self._parts()
        if len(ps) > 3:
            return ps[3]
        lib = self._lib()
        return lib["file"] if lib else ""

    def get_description(self):
        ps = self._parts()
        if len(ps) > 4:
            return ps[4]
        lib = self._lib()
        return lib.get("library", "") if lib else ""

    def get_name(self):
        """the name of the library (stdole, not stdole2), as VB6 gives it"""
        lib = self._lib()
        if lib:
            return lib["short"]
        return os.path.splitext(os.path.basename(self.get_fullpath().replace("\\", "/")))[0]

    def _version(self):
        ps = self._parts()
        if len(ps) > 1:
            return ps[1]
        lib = self._lib()
        return lib["version"] if lib else "0.0"

    def get_major(self):
        try:
            return int(self._version().split(".")[0])
        except (IndexError, ValueError):
            return 0

    def get_minor(self):
        try:
            return int(self._version().split(".")[1])
        except (IndexError, ValueError):
            return 0

    def get_isbroken(self):
        return False

    def get_builtin(self):
        return self.kind == "BuiltIn"

    def get_collection(self):
        return NOTHING

    def get_type(self):
        """vbext_rk_TypeLib 0, vbext_rk_Project 1 (a project of the group: *\\A)"""
        return 1 if self.line.startswith("*\\A") else 0


class VBComponents(Collection):
    TYPENAME = "VBComponents"

    def __init__(self, vbe, p):
        super().__init__(vbe)
        self.p = p

    def items(self):
        return [self.vbe.wrap(VBComponent, m) for m in self.p.modules]

    def name_of(self, it):
        return it.m.name

    def get_parent(self):
        return self.vbe.wrap(VBProject, self.p)

    def get_vbe(self):
        return self.vbe

    def call_add(self, comptype):
        """Add(vbext_ct_StdModule ...): a new module of the project"""
        kind = CT_KIND.get(int(comptype))
        if kind is None:
            raise AddInError(E_INVALIDARG, "Invalid procedure call or argument")
        main = self.vbe.main
        from .addinhost import hold
        with hold(main):                        # (the events after: the new one is the selected one too)
            m = self.p.add_module(kind)
            main.fill_tree()
            sel = getattr(main, "select_in_tree", None)
            if sel is not None:
                sel(m)
            main.update_title()
        return self.vbe.wrap(VBComponent, m)

    def call_remove(self, comp):
        """Remove: a component never saved is written to its file first (VB6, oracle evlog: DoGetNewFileName ...
        AfterWriteFile, then ItemRemoved, ItemSelected, AfterRemoveFile)"""
        if not isinstance(comp, VBComponent) or comp.m not in self.p.modules:
            raise AddInError(E_INVALIDARG, "Invalid procedure call or argument")
        m = comp.m
        main = self.vbe.main
        sync = getattr(main, "_sync_editors", None)
        if sync is not None:
            sync()
        if getattr(m, "never_saved", False) or not os.path.exists(m.path):
            if self.p.untitled:
                m.save(os.path.join(os.getcwd(), os.path.basename(m.rel)))
            else:
                m.save()
        nxt = _tree_next_module(main, m)
        from .addinhost import file_event, hold
        with hold(main):                        # (ItemRemoved, then the next one in the Project Explorer selected)
            for d in (getattr(main, "code_windows", {}), getattr(main, "design_windows", {})):
                sw = d.get(m)
                if sw is not None:
                    sw.close()
            self.p.remove_module(m)
            main.fill_tree()
            if nxt is not None and hasattr(main, "select_in_tree"):
                main.select_in_tree(nxt)
            elif hasattr(main, "select_project_in_tree"):
                main.select_project_in_tree(self.p)
        ide_changed(self.vbe)
        file_event(main, "AfterRemoveFile", self.p, file_type(m), m.path)

    def call_addfile(self, path, related=False):
        """AddFile (VB6 has no Import): the component comes under the name of a new one (Module1), its own name
        after its code is used (VB6: ItemAdded Module1, AfterAddFile, then ItemRenamed - oracle evlog)"""
        from .vbmodel import KIND_OF_EXT
        path = str(path)
        if not os.path.isfile(path):
            raise AddInError(53, "File not found")
        kind = KIND_OF_EXT.get(os.path.splitext(path)[1].lower())
        main = self.vbe.main
        before = list(self.p.modules)
        temp = self.p.unique_name(kind) if kind in ("Form", "Module", "Class", "UserControl", "PropertyPage") else None
        from .addinhost import sync_all, file_event, hold
        with hold(main):
            if main.project is not self.p:
                main.set_active(self.p)
            main.add_file_path(path, events=False)
            new = [m for m in self.p.modules if m not in before]
            if not new:
                raise AddInError(53, "File not found")
            c = self.vbe.wrap(VBComponent, new[0])
            if temp and temp.lower() != new[0].name.lower():
                c.pending = temp
        sync_all(main)
        file_event(main, "AfterAddFile", self.p, file_type(new[0]), path)
        return c


def _tree_next_module(main, m):
    """the component after m in the Project Explorer (None: none after it)"""
    from PyQt6.QtCore import Qt
    tree = getattr(main, "tree", None)
    if tree is None:
        return None
    order = []

    def walk(x):
        for i in range(x.childCount()):
            c = x.child(i)
            order.append(c.data(0, Qt.ItemDataRole.UserRole))
            walk(c)
    walk(tree.invisibleRootItem())
    if m not in order:
        return None
    for x in order[order.index(m) + 1:]:
        if x is not None and x is not m and hasattr(x, "kind"):
            return x
    return None


class VBComponent(VBObject):
    TYPENAME = "VBComponent"

    def __init__(self, vbe, m):
        super().__init__(vbe)
        self.m = m
        self.pending = None                     # AddFile: the name it shows until its code is used

    def get_name(self):
        return self.pending or self.m.name

    def let_name(self, v):
        self.settle()
        err = self.vbe.main.rename_module(self.m, str(v))
        if err:
            raise AddInError(380 if err == "Invalid property value" else 32813, err)
        self.vbe.main.fill_tree()

    def settle(self):
        """(AddFile) the component takes its own name: ItemRenamed"""
        if self.pending:
            old, self.pending = self.pending, None
            from .addinhost import fire_all
            fire_all(self.vbe.main, "VBComponentsEvents", "ItemRenamed", self.m.project, ("c", self.m), old)

    def get_type(self):
        k = self.m.kind
        if k == "Form" and getattr(self.m, "display_kind", "") == "MDIForm":
            return 5
        return CT.get(k, 11)

    def get_codemodule(self):
        self.settle()
        return self.vbe.wrap(CodeModule, self.m)

    def get_properties(self, *args):
        root = self.m.sf.root if self.m.sf is not None else None
        if root is None or not self.m.has_designer:
            c = self.vbe.wrap(Properties, self.m, None)
        else:
            c = self.vbe.wrap(Properties, root, self.m)
        return c.call_item(args[0]) if args else c

    def get_designer(self):
        if not self.m.has_designer or self.m.kind == "Designer":
            return NOTHING
        d = self.vbe.form_designer(self.m)
        return self.vbe.wrap(VBForm, d.root, self.m) if d is not None else NOTHING

    def get_collection(self):
        return self.vbe.wrap(VBComponents, self.m.project)

    def get_isdirty(self):
        return bool(self.m.dirty)

    def get_saved(self):
        return not self.m.dirty

    def get_filenames(self, index=1):
        return self.m.path if self.m.project.path or os.path.exists(self.m.path) else ""

    def call_activate(self):
        main = self.vbe.main
        if self.m.has_designer:
            main.open_designer(self.m)
        else:
            main.open_code(self.m)

    def call_savefile(self, path=None):
        self.m.save(str(path) if path else None)

    def call_export(self, path):
        self.m.sync()
        self.m.write_form(str(path), permanent=False)

    def get_hasopendesigner(self):
        return self.m in getattr(self.vbe.main, "design_windows", {})


# ---------------------------------------------------------------- the code
class CodeModule(VBObject):
    TYPENAME = "CodeModule"

    def __init__(self, vbe, m):
        super().__init__(vbe)
        self.m = m

    def _cw(self):
        sw = getattr(self.vbe.main, "code_windows", {}).get(self.m)
        return sw.widget() if sw is not None else None

    def _lines(self):
        cw = self._cw()
        text = cw.text() if cw is not None else self.m.code_text()
        lines = text.replace("\r\n", "\n").split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        return lines

    def _put(self, lines):
        text = "\n".join(lines) + ("\n" if lines else "")
        cw = self._cw()
        if cw is not None:
            cw.set_text(text)
        self.m.set_code(text)
        self.m.dirty = True
        upd = getattr(self.vbe.main, "update_tab_title", None)
        if cw is not None and upd is not None:
            upd(cw)

    def get_name(self):
        return self.m.name

    def get_parent(self):
        return self.vbe.wrap(VBComponent, self.m)

    def get_codepane(self):
        return self.vbe.wrap(CodePane, self.m)

    def get_countoflines(self):
        return len(self._lines())

    def get_countofdeclarationlines(self):
        return declaration_lines(self._lines())

    def get_lines(self, start, count):
        ls = self._lines()
        start, count = int(start), int(count)
        if start < 1 or count < 1 or start > len(ls):    # (VB6: Lines(1, 0) is an error too, oracle evlog)
            raise AddInError(E_INVALIDARG, "Invalid procedure call or argument")
        return "\r\n".join(ls[start - 1:start - 1 + count])

    call_lines = get_lines

    def call_insertlines(self, line, text):
        ls = self._lines()
        line = int(line)
        if line < 1:
            raise AddInError(5, "Invalid procedure call or argument")
        new = str(text).replace("\r\n", "\n").split("\n")
        i = min(line - 1, len(ls))
        ls[i:i] = new
        self._put(ls)

    def call_deletelines(self, start, count=1):
        ls = self._lines()
        start, count = int(start), int(count)
        if start < 1 or start + count - 1 > len(ls):
            raise AddInError(5, "Invalid procedure call or argument")
        del ls[start - 1:start - 1 + count]
        self._put(ls)

    def call_replaceline(self, line, text):
        ls = self._lines()
        line = int(line)
        if not 1 <= line <= len(ls):
            raise AddInError(5, "Invalid procedure call or argument")
        ls[line - 1:line] = str(text).replace("\r\n", "\n").split("\n")
        self._put(ls)

    def call_addfromstring(self, text):
        """AddFromString: after the declarations, as VB6 puts it"""
        ls = self._lines()
        i = declaration_lines(ls)
        ls[i:i] = str(text).replace("\r\n", "\n").split("\n")
        self._put(ls)

    def call_addfromfile(self, path):
        with open(str(path), "rb") as f:
            self.call_addfromstring(f.read().decode(self.m.project.encoding, "replace"))

    def call_find(self, target, startline=1, startcol=1, endline=-1, endcol=-1, wholeword=False, matchcase=False,
                  patternsearch=False):
        import re
        ls = self._lines()
        sl, el = max(1, int(startline)), int(endline)
        if el < 0 or el > len(ls):
            el = len(ls)
        pat = str(target) if patternsearch else re.escape(str(target))
        if wholeword:
            pat = r"\b" + pat + r"\b"
        rx = re.compile(pat, 0 if matchcase else re.IGNORECASE)
        for n in range(sl, el + 1):
            mt = rx.search(ls[n - 1], (int(startcol) - 1) if n == sl else 0)
            if mt:                              # where it is: the columns from 1, the end after it (oracle evlog)
                return Out(True, {1: n, 2: mt.start() + 1, 3: n, 4: mt.end() + 1})
        return False

    # the procedures: (name, kind) -> lines
    def _procs(self):
        return procedures(self._lines())

    def _proc(self, name, kind):
        for p in self._procs():
            if p[0].lower() == str(name).lower() and (kind is None or p[1] == int(kind)):
                return p
        raise AddInError(35, "Sub or Function not defined")

    def get_procstartline(self, name, kind):
        return self._proc(name, kind)[2]

    def get_procbodyline(self, name, kind):
        return self._proc(name, kind)[3]

    def get_proccountlines(self, name, kind):
        p = self._proc(name, kind)
        return p[4] - p[2] + 1

    def get_procofline(self, line, kind=None):
        """the procedure of a line, its kind ByRef in the second argument"""
        line = int(line)
        for p in self._procs():
            if p[2] <= line <= p[4]:
                return Out(p[0], {1: p[1]})
        return ""

    call_procofline = get_procofline

    def call_createeventproc(self, event, obj):
        """CreateEventProc: Private Sub <object>_<event>(<its parameters>) ... End Sub at the end, as the code window
        adds it; its line (VB6, oracle evlog: 1 in an empty module); an object or event that is not there: 57017"""
        from . import vbtypes
        from .codewin import uctl_events
        m = self.m
        obj, event = str(obj), str(event)
        tname, arr = None, False
        root_kind = None
        if m.is_form or m.kind in ("UserControl", "PropertyPage"):
            root_kind = m.kind if m.kind in ("UserControl", "PropertyPage") else \
                "MDIForm" if getattr(m, "display_kind", "") == "MDIForm" else "Form"
        if root_kind and obj.lower() == root_kind.lower():
            tname, obj = root_kind, root_kind
        elif m.kind == "Class" and obj.lower() == "class":
            tname, obj = "Class", "Class"
        else:
            for name, t, a in (m.controls() if root_kind else []):
                if name.lower() == obj.lower():
                    tname, arr, obj = t, a, name
                    break
        if tname is None:
            raise AddInError(57017, "Event handler is invalid")
        evs = [("Initialize", ""), ("Terminate", "")] if tname == "Class" else vbtypes.events(tname) or uctl_events(tname)
        ev = next((e for e in evs if e[0].lower() == event.lower()), None)
        if ev is None:
            raise AddInError(57017, "Event handler is invalid")
        name = f"{obj}_{ev[0]}"
        for p in self._procs():
            if p[0].lower() == name.lower():
                return p[3]
        params = ev[1] or ""
        if arr:
            params = "Index As Integer" + (", " + params if params else "")
        ls = self._lines()
        if any(x.strip() for x in ls):
            ls.append("")
        ls += [f"Private Sub {name}({params})", "", "End Sub"]
        self._put(ls)
        return len(ls) - 2


class CodePanes(Collection):
    TYPENAME = "CodePanes"

    def __init__(self, vbe, main):
        super().__init__(vbe)
        self.main = main

    def items(self):
        return [self.vbe.wrap(CodePane, m) for m in getattr(self.main, "code_windows", {})]

    def get_current(self):
        return self.vbe.get_activecodepane()


class CodePane(VBObject):
    TYPENAME = "CodePane"

    def __init__(self, vbe, m):
        super().__init__(vbe)
        self.m = m

    def _cw(self):
        sw = getattr(self.vbe.main, "code_windows", {}).get(self.m)
        return sw.widget() if sw is not None else None

    def get_codemodule(self):
        return self.vbe.wrap(CodeModule, self.m)

    def get_window(self):
        sw = getattr(self.vbe.main, "code_windows", {}).get(self.m)
        return self.vbe.wrap(Window, sw) if sw is not None else NOTHING

    def get_collection(self):
        return self.vbe.get_codepanes()

    def call_show(self):
        self.vbe.main.open_code(self.m)

    def get_topline(self):
        cw = self._cw()
        return cw.editor.firstVisibleLine() + 1 if cw is not None else 1

    def let_topline(self, v):
        cw = self._cw()
        if cw is not None:
            cw.editor.setFirstVisibleLine(max(0, int(v) - 1))

    def get_countofvisiblelines(self):
        cw = self._cw()
        return cw.editor.SendScintilla(cw.editor.SCI_LINESONSCREEN) if cw is not None else 0

    def call_setselection(self, startline, startcol, endline, endcol):
        cw = self.vbe.main.open_code(self.m)
        cw.editor.setSelection(int(startline) - 1, int(startcol) - 1, int(endline) - 1, int(endcol) - 1)

    def get_codepaneview(self):
        return 0


# ---------------------------------------------------------------- windows
class Windows(Collection):
    TYPENAME = "Windows"

    def __init__(self, vbe, main):
        super().__init__(vbe)
        self.main = main

    def items(self):
        mdi = getattr(self.main, "mdi", None)
        return [self.vbe.wrap(Window, sw) for sw in (mdi.subWindowList() if mdi is not None else [])]

    def name_of(self, it):
        return it.get_caption()


class Window(VBObject):
    TYPENAME = "Window"

    def __init__(self, vbe, w, kind=None):
        super().__init__(vbe)
        self.w = w
        self.kind = kind

    def get_caption(self):
        return self.w.windowTitle()

    def get_visible(self):
        return self.w.isVisible()

    def let_visible(self, v):
        self.w.setVisible(bool(v))

    def call_close(self):
        if self.kind != "main":
            self.w.close()

    def call_setfocus(self):
        self.w.activateWindow()
        self.w.setFocus()

    def get_hwnd(self):
        return int(self.w.winId())

    def get_left(self):
        return self.w.x()

    def get_top(self):
        return self.w.y()

    def get_width(self):
        return self.w.width()

    def get_height(self):
        return self.w.height()

    def get_windowstate(self):
        return 1 if self.w.isMinimized() else 2 if self.w.isMaximized() else 0


# ---------------------------------------------------------------- the command bars (Office)
class CommandBars(Collection):
    TYPENAME = "CommandBars"

    def __init__(self, vbe, main):
        super().__init__(vbe)
        self.main = main

    def _menubar(self):
        return self.vbe.wrap(CommandBar, self.main.menuBar(), "Menu Bar")

    def items(self):
        out = [self._menubar()]
        for a in self.main.menuBar().actions():        # each menu is a command bar too (VB6: "Add-Ins", "Tools")
            if a.menu() is not None:
                out.append(self.vbe.wrap(CommandBar, a.menu(), a.text().replace("&", "")))
        for tb in self.main.findChildren(QToolBar):
            out.append(self.vbe.wrap(CommandBar, tb, tb.windowTitle()))
        return out

    def name_of(self, it):
        return it.name_

    def get_activemenubar(self):
        return self._menubar()

    def call_add(self, name="", position=None, menubar=False, temporary=False):
        tb = QToolBar(str(name) or "Custom", self.main)
        tb.setObjectName("addin:" + str(name))
        self.main.addToolBar(tb)
        tb.lb_temporary = bool(temporary)
        return self.vbe.wrap(CommandBar, tb, str(name))


class CommandBar(VBObject):
    TYPENAME = "CommandBar"

    def __init__(self, vbe, w, name):
        super().__init__(vbe)
        self.w = w                              # a QMenuBar, a QMenu or a QToolBar
        self.name_ = name

    def get_name(self):
        return self.name_

    def get_namelocal(self):
        return self.name_

    def get_controls(self, *args):
        c = self.vbe.wrap(CommandBarControls, self.w, self)
        return c.call_item(args[0]) if args else c

    def get_visible(self):
        return self.w.isVisible()

    def let_visible(self, v):
        if isinstance(self.w, QToolBar):
            self.w.setVisible(bool(v))

    def get_enabled(self):
        return self.w.isEnabled()

    def let_enabled(self, v):
        self.w.setEnabled(bool(v))

    def get_type(self):
        return 1 if self.name_ == "Menu Bar" else 2 if isinstance(self.w, QMenu) else 0   # msoBarType*

    def get_builtin(self):
        return not self.w.objectName().startswith("addin:")

    def get_index(self):
        its = self.vbe.get_commandbars().items()
        return next((i + 1 for i, it in enumerate(its) if it is self), 0)

    def call_delete(self):
        if isinstance(self.w, QToolBar) and self.w.objectName().startswith("addin:"):
            self.vbe.main.removeToolBar(self.w)
            self.w.deleteLater()

    def call_findcontrol(self, type_=None, id_=None, tag=None, visible=None, recursive=False):
        for c in self.get_controls().items():
            if tag is not None and c.tag != str(tag):
                continue
            if id_ is not None and c.get_id() != int(id_):
                continue
            return c
        return NOTHING


class CommandBarControls(Collection):
    TYPENAME = "CommandBarControls"

    def __init__(self, vbe, w, bar):
        super().__init__(vbe)
        self.w = w
        self.bar = bar

    def items(self):
        out = []
        for a in self.w.actions():
            if a.isSeparator():
                continue
            out.append(self.vbe.wrap(CommandBarControl, a, self.bar))
        return out

    def name_of(self, it):
        return it.a.text().replace("&", "")

    def get_parent(self):
        return self.bar

    def call_add(self, type_=1, id_=1, parameter=None, before=None, temporary=False):
        """Add(msoControlButton / msoControlPopup, Id, Parameter, Before, Temporary): a new command at the end of
        the bar (or before the control number Before)"""
        t = int(type_) if type_ not in (None, "") else 1
        acts = [a for a in self.w.actions() if not a.isSeparator()]
        at = None
        if before not in (None, "") and 1 <= int(before) <= len(acts):
            at = acts[int(before) - 1]
        if t == 10:                             # msoControlPopup: a submenu
            sub = QMenu("", self.w)
            a = sub.menuAction()
        else:
            a = QAction("", self.w)
        a.lb_addin = True
        a.lb_type = t
        a.lb_temporary = bool(temporary)
        if at is not None:
            self.w.insertAction(at, a)
        else:
            self.w.addAction(a)
        c = self.vbe.wrap(CommandBarControl, a, self.bar)
        if parameter not in (None, ""):
            c.parameter = str(parameter)
        if self.vbe.host is not None:
            self.vbe.host.control_added(c)
            a.triggered.connect(lambda checked=False, c=c, h=self.vbe.host: h.control_clicked(c))
        return c


class CommandBarControl(VBObject):
    TYPENAME = "CommandBarControl"

    def __init__(self, vbe, a, bar):
        super().__init__(vbe)
        self.a = a                              # its QAction
        self.bar = bar
        self.tag = ""
        self.parameter = ""
        self.description = ""
        self.group_sep = None

    def typename(self):
        return "CommandBarPopup" if self.a.menu() is not None else "CommandBarButton"

    def get_caption(self):
        return self.a.text()

    def let_caption(self, v):
        self.a.setText(str(v))

    def get_enabled(self):
        return self.a.isEnabled()

    def let_enabled(self, v):
        self.a.setEnabled(bool(v))

    def get_visible(self):
        return self.a.isVisible()

    def let_visible(self, v):
        self.a.setVisible(bool(v))

    def get_tag(self):
        return self.tag

    def let_tag(self, v):
        self.tag = str(v)

    def get_parameter(self):
        return self.parameter

    def let_parameter(self, v):
        self.parameter = str(v)

    def get_tooltiptext(self):
        return self.a.toolTip()

    def let_tooltiptext(self, v):
        self.a.setToolTip(str(v))

    def get_descriptiontext(self):
        return self.description

    def let_descriptiontext(self, v):
        self.description = str(v)
        self.a.setStatusTip(str(v))

    def get_begingroup(self):
        return self.group_sep is not None

    def let_begingroup(self, v):
        w = self.bar.w
        if bool(v) and self.group_sep is None:
            self.group_sep = w.insertSeparator(self.a)
        elif not bool(v) and self.group_sep is not None:
            w.removeAction(self.group_sep)
            self.group_sep = None

    def get_type(self):
        return getattr(self.a, "lb_type", 10 if self.a.menu() is not None else 1)

    def get_id(self):
        return 1 if getattr(self.a, "lb_addin", False) else abs(hash(self.a.text())) % 30000 + 2

    def get_builtin(self):
        return not getattr(self.a, "lb_addin", False)

    def get_index(self):
        acts = [x for x in self.bar.w.actions() if not x.isSeparator()]
        return acts.index(self.a) + 1 if self.a in acts else 0

    def get_parent(self):
        return self.bar

    def get_state(self):
        return -1 if self.a.isChecked() else 0

    def let_state(self, v):
        self.a.setCheckable(True)
        self.a.setChecked(int(v) != 0)

    def get_style(self):
        return getattr(self, "style_", 0)

    def let_style(self, v):
        self.style_ = int(v)

    def get_faceid(self):
        return 0

    def let_faceid(self, v):
        pass

    def get_commandbar(self):
        """a popup: the command bar of its menu"""
        m = self.a.menu()
        if m is None:
            raise AddInError(438, "Object doesn't support this property or method")
        return self.vbe.wrap(CommandBar, m, self.a.text().replace("&", ""))

    def get_controls(self, *args):
        return self.get_commandbar().get_controls(*args)

    def call_pasteface(self):
        """PasteFace: the picture on the clipboard as the picture of the command"""
        img = QApplication.clipboard().image()
        if img is not None and not img.isNull():
            self.a.setIcon(QIcon(QPixmap.fromImage(img)))

    def call_copyface(self):
        ic = self.a.icon()
        if not ic.isNull():
            QApplication.clipboard().setPixmap(ic.pixmap(16, 16))

    def call_execute(self):
        self.a.trigger()

    def call_setfocus(self):
        pass

    def call_reset(self):
        pass

    def call_delete(self, temporary=False):
        w = self.bar.w
        if self.group_sep is not None:
            w.removeAction(self.group_sep)
            self.group_sep = None
        w.removeAction(self.a)
        if self.vbe.host is not None:
            self.vbe.host.control_deleted(self)


# ---------------------------------------------------------------- events
class Events(VBObject):
    TYPENAME = "Events"

    def __init__(self, vbe):
        super().__init__(vbe)
        self.bar_events = []

    def get_commandbarevents(self, ctl):
        """the Click of a command bar control (WithEvents ... As CommandBarEvents)"""
        if not isinstance(ctl, CommandBarControl):
            raise AddInError(5, "Invalid procedure call or argument")
        for e in self.bar_events:
            if e.ctl is ctl:
                return e
        e = CommandBarEvents(self.vbe, ctl)
        self.bar_events.append(e)
        return e

    call_commandbarevents = get_commandbarevents

    def get_vbprojectsevents(self):
        return self.vbe.wrap(SourceEvents, self.vbe.main, "VBProjectsEvents")

    def get_vbcomponentsevents(self, project=None):
        return self.vbe.wrap(SourceEvents, project if project is not None else self.vbe.main, "VBComponentsEvents")

    def get_referencesevents(self, project=None):
        return self.vbe.wrap(SourceEvents, project if project is not None else self, "ReferencesEvents")

    def get_vbcontrolsevents(self, project=None, form=None):
        """the controls added, removed, renamed on a form (VBForm: its root control)"""
        if not isinstance(form, VBForm):
            raise AddInError(E_INVALIDARG, "Invalid procedure call or argument")
        return self.vbe.wrap(SourceEvents, form.root, "VBControlsEvents")

    def get_selectedvbcontrolsevents(self, project=None, form=None):
        if not isinstance(form, VBForm):
            raise AddInError(E_INVALIDARG, "Invalid procedure call or argument")
        return self.vbe.wrap(SourceEvents, form.root, "SelectedVBControlsEvents")

    def get_filecontrolevents(self, project=None):
        return self.vbe.wrap(SourceEvents, project if project is not None else self, "FileControlEvents")


class CommandBarEvents(VBObject):
    TYPENAME = "CommandBarEvents"

    def __init__(self, vbe, ctl):
        super().__init__(vbe)
        self.ctl = ctl


class SourceEvents(VBObject):
    """the events of the projects, components ...: an object whose events the add-in can take (WithEvents); the IDE
    raises ItemAdded / ItemRemoved / ItemRenamed / ItemActivated / ItemSelected through the host"""

    def __init__(self, vbe, key, name):
        super().__init__(vbe)
        self.TYPENAME = name


# ---------------------------------------------------------------- the add-ins
class AddIns(Collection):
    TYPENAME = "Addins"

    def __init__(self, vbe, main):
        super().__init__(vbe)
        self.main = main

    def items(self):
        from .addinhost import user_addins
        return [self.vbe.wrap(AddIn, self.main, a["progid"]) for a in user_addins(self.main)]

    def name_of(self, it):
        return it.progid

    def call_update(self):
        pass


class AddIn(VBObject):
    TYPENAME = "AddIn"

    def __init__(self, vbe, main, progid):
        super().__init__(vbe)
        self.main = main
        self.progid = progid

    def _info(self):
        from .addinhost import user_addins
        return next((a for a in user_addins(self.main) if a["progid"] == self.progid), {})

    def get_progid(self):
        return self.progid

    def get_description(self):
        return self._info().get("description", "")

    def get_guid(self):
        return ""

    def get_connect(self):
        from .addinhost import loaded_host
        return loaded_host(self.main, self.progid) is not None

    def let_connect(self, v):
        from .addinhost import load_addin, unload_addin
        if bool(v):
            load_addin(self.main, self.progid, 0)
        else:
            unload_addin(self.main, self.progid, 1)

    def get_collection(self):
        return self.vbe.get_addins()

    def get_object(self):
        return NOTHING


# ---------------------------------------------------------------- the designers of forms: VBForm, VBControls, Properties
class Properties(Collection):
    """the properties of a form, a control (as the Properties window has them) or a code module (its Name)"""
    TYPENAME = "Properties"

    def __init__(self, vbe, ctl, m):
        super().__init__(vbe)
        self.ctl = ctl                          # a Control of the .frm, or a Module without a designer
        self.m = m

    def items(self):
        from . import vbprops as V
        if self.m is None:                      # a module / class: its Name
            return [self.vbe.wrap(Property, self.ctl, None, "Name")]
        return [self.vbe.wrap(Property, self.ctl, self.m, p.name) for p in V.props_of(self.ctl.type_name)]

    def name_of(self, it):
        return it.get_name()


class Property(VBObject):
    TYPENAME = "Property"

    def __init__(self, vbe, ctl, m, name):
        super().__init__(vbe)
        self.ctl, self.m, self.pname = ctl, m, name

    def _p(self):
        from . import vbprops as V
        return V.prop_def(self.ctl.type_name, self.pname)

    def get_name(self):
        return "Name" if self.pname == "(Name)" else self.pname

    def get_numindices(self):
        return 0

    def get_value(self):
        from . import vbprops as V
        if self.m is None:
            return self.ctl.name
        p = self._p()
        if p is None:
            raise AddInError(438, "Object doesn't support this property or method")
        d = self.vbe.form_designer(self.m)
        v = V.value_of(self.ctl, p, d.root if d is not None else None)
        if p.kind == "font":
            raise AddInError(438, "Object doesn't support this property or method")
        if p.kind == "color":
            v = int(v) - 0x100000000 if int(v) >= 0x80000000 else int(v)
        return v

    def let_value(self, v):
        from . import vbprops as V
        if self.m is None:
            err = self.vbe.main.rename_module(self.ctl, str(v))
            if err:
                raise AddInError(380, err)
            self.vbe.main.fill_tree()
            return
        p = self._p()
        d = self.vbe.form_designer(self.m)
        if p is None or d is None or p.kind in ("font", "picture", "readonly"):
            raise AddInError(383, "'Item' property cannot be set") if p is not None and p.kind == "readonly" else \
                AddInError(438, "Object doesn't support this property or method")
        try:
            if p.kind == "name":
                val = str(v)
            elif isinstance(v, str) and p.kind not in ("string", "list"):
                val = V.parse_input(p, v)
            elif p.kind == "bool":
                val = bool(v)
            elif p.kind == "color":
                val = int(v) & 0xFFFFFFFF
            elif p.kind in ("enum", "int", "index"):
                val = int(v)
            elif p.kind in ("pos", "scale", "formpos", "num", "single"):
                val = int(v) if float(v) == int(float(v)) else float(v)
            else:
                val = v if not isinstance(v, (int, float)) or isinstance(v, bool) else str(v) if p.kind == "string" else v
        except (TypeError, ValueError):
            raise AddInError(380, "Invalid property value")
        from .addinhost import api_change
        with api_change(self.vbe.main):
            err = d.set_property(self.ctl, p, val)
        if err:
            raise AddInError(380, err)
        ide_changed(self.vbe)

    def get_parent(self):
        return self.vbe.wrap(Properties, self.ctl, self.m)

    def get_collection(self):
        return self.get_parent()


class VBForm(VBObject):
    """the designer of a form (VBComponent.Designer): its controls"""
    TYPENAME = "VBForm"

    def __init__(self, vbe, root, m):
        super().__init__(vbe)
        self.root, self.m = root, m

    def get_vbcontrols(self):
        return self.vbe.wrap(VBControls, self.root, self.m, True)

    def get_selectedvbcontrols(self):
        return self.vbe.wrap(SelectedVBControls, self.root, self.m)

    def get_canpastefromclipboard(self):
        return QApplication.clipboard().mimeData().hasFormat("application/x-lbstudio-controls")

    def call_pastefromclipboard(self, container=None):
        d = self.vbe.form_designer(self.m)
        if d is not None:
            from .addinhost import api_change
            with api_change(self.vbe.main):
                d.paste()
            ide_changed(self.vbe)

    def get_parent(self):
        return self.vbe.wrap(VBComponent, self.m)

    def get_vbe(self):
        return self.vbe


def _controls_under(c, deep):
    out = []
    for ch in reversed(c.children):             # (the newest first in the .frm: VB6's order is the oldest first)
        out.append(ch)
        if deep:
            out += _controls_under(ch, True)
    return out


class VBControls(Collection):
    TYPENAME = "VBControls"

    def __init__(self, vbe, container, m, all_=False):
        super().__init__(vbe)
        self.container, self.m, self.all = container, m, all_

    def items(self):
        return [self.vbe.wrap(VBControl, c, self.m) for c in _controls_under(self.container, self.all)]

    def name_of(self, it):
        return it.ctl.name

    def call_add(self, *args):
        """Add(ProgId, RelativeVBControl, fContainer): VB6 gives Type mismatch when the last two are left out
        (oracle evlog); the new control at the top left of the form or of its container"""
        from PyQt6.QtCore import QRect
        if len(args) < 2:
            raise AddInError(13, "Type mismatch")
        progid, rel = str(args[0]), args[1]
        container = bool(args[2]) if len(args) > 2 and args[2] is not None else False
        d = self.vbe.form_designer(self.m)
        if d is None:
            raise AddInError(E_INVALIDARG, "Invalid procedure call or argument")
        cls = progid[3:] if progid.lower().startswith("vb.") else progid
        item = None
        if isinstance(rel, VBControl) and container:
            item = d.items.get(rel.ctl)
        from .addinhost import api_change
        with api_change(self.vbe.main):
            c = d.place_new(cls, QRect(0, 0, 0, 0), item)
        if c is None:
            raise AddInError(711, f"Invalid class string: {progid}")
        ide_changed(self.vbe)
        return self.vbe.wrap(VBControl, c, self.m)

    def call_remove(self, ctl):
        d = self.vbe.form_designer(self.m)
        if not isinstance(ctl, VBControl) or d is None or ctl.ctl not in d.items:
            raise AddInError(-2147417848, "Method '~' of object '~' failed")
        from .addinhost import api_change
        with api_change(self.vbe.main):
            d.select_control(ctl.ctl)
            d.delete_selection()
        ide_changed(self.vbe)


class SelectedVBControls(Collection):
    TYPENAME = "SelectedVBControls"

    def __init__(self, vbe, root, m):
        super().__init__(vbe)
        self.root, self.m = root, m

    def items(self):
        d = getattr(self.vbe.main, "design_windows", {}).get(self.m)
        d = d.widget() if d is not None else None
        sel = [it.ctl for it in getattr(d, "selection", [])] if d is not None else []
        return [self.vbe.wrap(VBControl, c, self.m) for c in sel]

    def name_of(self, it):
        return it.ctl.name

    def call_clear(self):
        d = self.vbe.form_designer(self.m)
        if d is not None:
            d.select([])


class VBControl(VBObject):
    TYPENAME = "VBControl"

    def __init__(self, vbe, ctl, m):
        super().__init__(vbe)
        self.ctl, self.m = ctl, m

    def get_properties(self, *args):
        c = self.vbe.wrap(Properties, self.ctl, self.m)
        return c.call_item(args[0]) if args else c

    def get_progid(self):
        return self.ctl.cls

    def get_classname(self):
        return self.ctl.type_name

    def get_containedvbcontrols(self):
        return self.vbe.wrap(VBControls, self.ctl, self.m, False)

    def get_collection(self):
        return self.vbe.wrap(VBControls, self.ctl.parent or self.ctl, self.m, True)

    def get_controlobject(self):
        return NOTHING

    def get_container(self):
        par = self.ctl.parent
        if par is None or par.parent is None:
            return NOTHING
        return self.vbe.wrap(VBControl, par, self.m)

    def call_inselection(self):
        d = getattr(self.vbe.main, "design_windows", {}).get(self.m)
        d = d.widget() if d is not None else None
        return d is not None and any(it.ctl is self.ctl for it in getattr(d, "selection", []))

    get_inselection = call_inselection

    def let_inselection(self, v):
        d = self.vbe.form_designer(self.m)
        if d is None:
            return
        sel = [it for it in d.selection if it.ctl is not self.ctl]
        if bool(v) and self.ctl in d.items:
            sel.append(d.items[self.ctl])
        d.select(sel)

    def get_vbe(self):
        return self.vbe
