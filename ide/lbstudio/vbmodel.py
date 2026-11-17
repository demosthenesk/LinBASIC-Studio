"""The open project: the .vbp and its modules, read and written with the project readers
of lbvc (compiler/lbvc/project), so that an untouched file is saved byte for byte."""
import locale
import os
import re
import shutil
import sys

from . import paths

if paths.COMPILER not in sys.path:
    sys.path.insert(0, paths.COMPILER)
from lbvc.project import load_project, load_source  # noqa: E402
from lbvc.project.vbp import Project  # noqa: E402
from lbvc.project.source import parse_source, Property, PropertyBlock, FrxRef  # noqa: E402
from lbvc.project.frx import FrxStore, FrxData, FrxWriter  # noqa: E402

KIND_OF_EXT = {".frm": "Form", ".bas": "Module", ".cls": "Class", ".ctl": "UserControl",
               ".pag": "PropertyPage", ".dob": "UserDocument", ".dsr": "Designer"}
FOLDER_OF_KIND = {"Form": "Forms", "MDIForm": "Forms", "Module": "Modules", "Class": "Class Modules",
                  "UserControl": "User Controls", "PropertyPage": "Property Pages",
                  "UserDocument": "User Documents", "Designer": "Designers"}


def default_codepage():
    """VB6 wrote the files in the ANSI code page of Windows: Greek -> cp1253 ..."""
    lang = (os.environ.get("LC_ALL") or os.environ.get("LC_CTYPE") or os.environ.get("LANG")
            or locale.getlocale()[0] or "")
    table = {"el": "cp1253", "ru": "cp1251", "uk": "cp1251", "bg": "cp1251", "sr": "cp1251",
             "pl": "cp1250", "cs": "cp1250", "sk": "cp1250", "hu": "cp1250", "ro": "cp1250",
             "hr": "cp1250", "sl": "cp1250", "tr": "cp1254", "he": "cp1255", "ar": "cp1256",
             "et": "cp1257", "lv": "cp1257", "lt": "cp1257", "vi": "cp1258", "th": "cp874",
             "ja": "cp932", "zh_CN": "cp936", "ko": "cp949", "zh_TW": "cp950"}
    for key in (lang[:5], lang[:2]):
        if key in table:
            return table[key]
    return "cp1252"


class Module:
    """One source file of the project (form, module, class ...)."""

    def __init__(self, project, kind, name, rel, sf=None):
        self.project = project
        self.kind = kind                    # Form, Module, Class ...
        self.name = name
        self.rel = rel                      # as the .vbp has it (Form1.frm, Module1.bas)
        self.sf = sf
        self.dirty = False
        self.editor_text = None             # the text of an open code window (newer than sf.code)
        self.dsx_data = None                # a DataReport its designer changed: the .dsx to write

    @property
    def path(self):
        return self.project.resolve(self.rel)

    @property
    def is_form(self):
        return self.kind == "Form"

    @property
    def has_designer(self):
        """a form, a user control, a designer (.dsr): it opens in the designer"""
        return self.kind in ("Form", "UserControl", "PropertyPage", "Designer") and self.sf is not None and \
            self.sf.root is not None

    @property
    def designer_kind(self):
        """DataEnvironment / DataReport for a .dsr"""
        if self.kind != "Designer":
            return None
        from .dsrmodel import kind_of
        return kind_of(self.sf)

    # ------------------------------------------------------------ Instancing (class modules)
    INSTANCING = {1: "1 - Private", 2: "2 - PublicNotCreatable", 3: "3 - SingleUse", 4: "4 - GlobalSingleUse",
                  5: "5 - MultiUse", 6: "6 - GlobalMultiUse"}

    def _attr_true(self, name, default=False):
        v = self.sf.attribute(name) if self.sf is not None else None
        return default if v is None else str(v).strip().lower() == "true"

    def instancing(self):
        """VB6's Instancing of a class (or user control): from its attributes"""
        if self.sf is None:
            return 1
        if not self._attr_true("VB_Exposed"):
            return 1
        if not self._attr_true("VB_Creatable", True):
            return 2
        multi = True
        for p in self.sf.class_props or []:
            if p.name.lower() == "multiuse":
                multi = p.value not in (0, "0", False)
        glob = self._attr_true("VB_GlobalNameSpace")
        return (6 if glob else 5) if multi else (4 if glob else 3)

    def _set_attr(self, name, value):
        for i, (n, v, raw) in enumerate(self.sf.attributes):
            if n.lower() == name.lower():
                self.sf.attributes[i] = (n, value, f"Attribute {n} = {value}")
                return
        self.sf.attributes.append((name, value, f"Attribute {name} = {value}"))

    def set_instancing(self, n):
        exposed, creatable = n != 1, n not in (2,)
        if n == 1 and self.kind == "Class":
            creatable = True                    # a private class: VB_Creatable = True, as VB6 writes it
        glob = n in (4, 6)
        self._set_attr("VB_GlobalNameSpace", "True" if glob else "False")
        self._set_attr("VB_Creatable", "True" if creatable else "False")
        self._set_attr("VB_Exposed", "True" if exposed else "False")
        if self.sf.class_props is not None:
            multi = -1 if n in (1, 2, 5, 6) else 0
            for i, p in enumerate(self.sf.class_props):
                if p.name.lower() == "multiuse" and p.value != multi:
                    self.sf.class_props[i] = Property("MultiUse", multi, str(multi), "True" if multi else "False",
                                                      raw=f"  MultiUse = {multi}  '{'True' if multi else 'False'}")
        self.dirty = True

    @property
    def display_kind(self):
        if self.is_form and self.sf is not None and self.sf.root is not None and self.sf.root.type_name == "MDIForm":
            return "MDIForm"
        return self.kind

    def code_offset(self):
        """lines of the file before the code (VERSION, the designer section, Attribute lines)"""
        if self.sf is None:
            return 0
        return len(list(self.sf.lines())) - len(self.sf.code)

    def code_text(self):
        if self.editor_text is not None:
            return self.editor_text
        return "\n".join(self.sf.code) if self.sf is not None else ""

    def set_code(self, text):
        self.editor_text = text

    def sync(self):
        """the editor text into the source file"""
        if self.editor_text is not None and self.sf is not None:
            lines = self.editor_text.replace("\r\n", "\n").split("\n")
            if lines and lines[-1] == "":
                lines.pop()
            self.sf.code = lines

    def save(self, path=None):
        self.sync()
        target = path or self.path
        os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
        self.write_form(target, permanent=True)
        self.dirty = False

    # ------------------------------------------------------------ the .frx
    def frx_props(self):
        """the properties whose data is in the .frx (or will be), in the order of the .frm"""
        out = []

        def items(seq):
            for it in seq:
                if isinstance(it, PropertyBlock):
                    yield from items(it.items)
                elif isinstance(it, Property):
                    yield it

        def walk(c):
            for kind, it in c.order:
                if kind == "c":
                    walk(it)
                else:
                    for p in items([it]):
                        if isinstance(p.value, (FrxRef, FrxData)):
                            out.append(p)
        if self.sf is not None and self.sf.root is not None:
            walk(self.sf.root)
        return out

    def frx_store(self, ref):
        """the .frx a reference points to (read once)"""
        from lbvc.project.fsutil import ci_path
        path = ci_path(os.path.join(os.path.dirname(self.path), ref.file))
        cache = self.__dict__.setdefault("_frx", {})
        st = cache.get(path)
        if st is None:
            offs = [p.value.offset for p in self.frx_props() if isinstance(p.value, FrxRef) and p.value.file == ref.file]
            st = cache[path] = FrxStore.load(path, offs, self.sf.encoding) if os.path.exists(path) else None
        return st

    def frx_data(self, prop):
        """FrxData of a property (from the designer, or read from the .frx), or None"""
        v = prop.value
        if isinstance(v, FrxData):
            return v
        if isinstance(v, FrxRef):
            st = self.frx_store(v)
            return st.data_of(prop.name, v) if st is not None else None
        return None

    def frx_value(self, ctl, name):
        """the value of a .frx property for the designer: bytes of a picture, a list, a string"""
        p = ctl.prop(name)
        if p is None:
            return None
        if not isinstance(p.value, (FrxRef, FrxData)):
            return p.value
        d = self.frx_data(p)
        return d.data if d is not None else None

    def write_form(self, target, permanent):
        """the .frm (and its .frx when the designer changed what is in it) at target; permanent:
        the properties keep their new offsets (a save), else they come back (a copy for F5)"""
        if self.kind == "Designer":
            self.sf.save(target)
            if self.dsx_data is not None:
                from lbvc.project.fsutil import ci_path
                with open(ci_path(os.path.splitext(target)[0] + ".dsx"), "wb") as f:
                    f.write(self.dsx_data)
            return
        props = self.frx_props()
        frx_name = os.path.splitext(os.path.basename(target))[0] + ".frx"
        need = any(isinstance(p.value, FrxData) or p.value.file.lower() != frx_name.lower() for p in props)
        if not need:
            self.sf.save(target)
            return
        w = FrxWriter(self.sf.encoding)
        old = []
        for p in props:
            d = self.frx_data(p) or FrxData("raw", b"")
            off, dollar = w.write(d)
            old.append((p, p.value))
            p.value = FrxRef(frx_name, off, dollar)
        with open(os.path.join(os.path.dirname(target), frx_name), "wb") as f:
            f.write(w.bytes())
        self.sf.save(target)
        if permanent:
            self.__dict__.pop("_frx", None)
        else:
            for p, v in old:
                p.value = v

    def controls(self):
        """[(name, type, is_array)] of a form, the form itself first"""
        if self.sf is None or self.sf.root is None:
            return []
        out, seen = [], set()
        for c in self.sf.root.walk():
            if c is self.sf.root:
                continue
            key = c.name.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append((c.name, c.type_name, c.index is not None))
        return out


class VBProject:
    def __init__(self):
        self.group = None                       # the project group it is in (VBGroup), or None
        self.vbp = None
        self.path = None
        self.modules = []
        self.encoding = default_codepage()
        self.dirty = False

    # ------------------------------------------------------------------ files
    @classmethod
    def open(cls, path, encoding=None):
        p = cls()
        p.encoding = encoding or p.encoding
        p.path = os.path.abspath(path)
        p.vbp = load_project(p.path, p.encoding)
        for kind, name, rel in p.vbp.members():
            full = p.vbp.resolve(rel)
            sf = load_source(full, encoding=p.encoding) if os.path.exists(full) else None
            m = Module(p, kind, sf.name if sf is not None and sf.name else name, rel, sf)
            p.modules.append(m)
        return p

    @classmethod
    def new(cls, encoding=None, ptype="Exe", name="Project1"):
        """a new project, not saved yet: a Standard EXE with Form1, an ActiveX EXE or DLL with Class1,
        an ActiveX Control with UserControl1 (what VB6's New Project makes)"""
        p = cls()
        p.encoding = encoding or p.encoding
        p.vbp = Project.parse(new_vbp_text(ptype, name), None, True, p.encoding)
        p.vbp.encoding = p.encoding
        first = {"Exe": ("Form", "Form1"), "OleExe": ("Class", "Class1"), "OleDll": ("Class", "Class1"),
                 "Control": ("UserControl", "UserControl1")}[ptype]
        p.add_module(first[0], first[1], add_to_vbp=ptype != "Exe")
        # VB6's new project: the line of its first module right after Type=, before the Reference
        # (DEV-8/me2 Project1.vbp, pp1: Form=Form1.frm / Class=Class1; ... / UserControl=...)
        ents = p.vbp.entries
        i = next((i for i, (k, v) in enumerate(ents) if v is not None and k.strip() == p.modules[0].kind), None)
        if i is not None and i > 1:
            ents.insert(1, ents.pop(i))
        p.dirty = False
        for m in p.modules:                     # untouched: closing it asks nothing (Save As writes it)
            m.dirty = False
        return p

    @property
    def ptype(self):
        """Exe, OleExe, OleDll, Control"""
        t = (self.vbp.type if self.vbp else "Exe") or "Exe"
        return {"exe": "Exe", "oleexe": "OleExe", "oledll": "OleDll", "control": "Control"}.get(t.lower(), t)

    @property
    def is_component(self):
        return self.ptype in ("OleExe", "OleDll", "Control")

    def resolve(self, rel):
        base = os.path.dirname(self.path) if self.path else os.getcwd()
        return os.path.normpath(os.path.join(base, rel.replace("\\", "/")))

    @property
    def name(self):
        return self.vbp.name if self.vbp else "Project1"

    @property
    def startup(self):
        return self.vbp.startup if self.vbp else ""

    @property
    def untitled(self):
        return self.path is None

    def any_dirty(self):
        return self.dirty or any(m.dirty for m in self.modules)

    def save(self, path=None):
        if path:
            old_dir = os.path.dirname(self.path) if self.path else None
            self.path = os.path.abspath(path)
            self.vbp.path = self.path
            if old_dir and old_dir != os.path.dirname(self.path):
                for m in self.modules:          # the .frx files go along
                    self._copy_frx(m, old_dir, os.path.dirname(self.path))
                self._copy_res(old_dir, os.path.dirname(self.path))
        for m in self.modules:
            if m.dirty or not os.path.exists(m.path):
                m.save()
        self.vbp.save(self.path)
        self.dirty = False

    def _copy_res(self, old_dir, new_dir):
        """the resource file of the project (ResFile32=, for LoadResString ...) goes along"""
        from lbvc.project.fsutil import ci_path
        res = (self.vbp.get("ResFile32") or "").strip('"').replace("\\", "/")
        if not res or os.path.isabs(res) or re.match(r"^[A-Za-z]:", res):
            return
        src = ci_path(os.path.join(old_dir, res))
        dst = os.path.join(new_dir, res)
        if os.path.isfile(src) and os.path.abspath(src) != os.path.abspath(dst):
            os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
            shutil.copyfile(src, dst)

    def _copy_frx(self, m, old_dir, new_dir):
        """the .frx of a module goes along, under the names its .frm uses ("BUTTON.frx")"""
        from lbvc.project.fsutil import ci_path
        rel_dir = os.path.dirname(m.rel.replace("\\", "/"))
        base = os.path.splitext(m.rel.replace("\\", "/"))[0]
        names = {base + ext for ext in (".frx", ".ctx", ".dsx")}
        for p in m.frx_props():
            if isinstance(p.value, FrxRef):
                names.add(os.path.join(rel_dir, p.value.file))
        for name in names:
            src = ci_path(os.path.join(old_dir, name))
            if os.path.exists(src):
                dst = os.path.join(new_dir, name)
                os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
                if os.path.abspath(src) != os.path.abspath(dst):
                    shutil.copyfile(src, dst)

    def write_copy(self, folder):
        """the project as it is in the editors (not saved yet) into folder: what F5 compiles"""
        os.makedirs(folder, exist_ok=True)
        src_dir = os.path.dirname(self.path) if self.path else None
        for m in self.modules:
            m.sync()
            target = os.path.join(folder, m.rel.replace("\\", "/"))
            os.makedirs(os.path.dirname(target), exist_ok=True)
            if src_dir:
                self._copy_frx(m, src_dir, folder)
            m.write_form(target, permanent=False)
        if src_dir:
            self._copy_res(src_dir, folder)
        vbp = os.path.join(folder, os.path.basename(self.path) if self.path else self.name + ".vbp")
        with open(vbp, "wb") as f:
            f.write(self.vbp.dumps().encode(self.encoding, errors="surrogateescape"))
        return vbp

    # ---------------------------------------------------------------- members
    def mdi_form(self):
        """the MDI form of the project (VB6: one at most; Project > Add MDI Form is grey then)"""
        return next((m for m in self.modules if m.display_kind == "MDIForm"), None)

    def unique_name(self, base):
        names = {m.name.lower() for m in self.modules}
        i = 1
        while f"{base}{i}".lower() in names:
            i += 1
        return f"{base}{i}"

    def add_designer(self, kind, name=None):
        """Project > Add Data Environment / Add Data Report: a .dsr as VB6's templates make it, and the
        references it needs"""
        from . import dsrmodel
        name = name or self.unique_name(kind)
        text = dsrmodel.new_dataenv_text(name) if kind == "DataEnvironment" else dsrmodel.new_report_text(name)
        sf = parse_source(text.replace("\n", "\r\n"), None, "Designer", True, self.encoding)
        m = Module(self, "Designer", name, name + ".Dsr", sf)
        m.dirty = True
        self.modules.append(m)
        self._vbp_add(m)
        self.add_references(dsrmodel.REFERENCES[kind])
        self.dirty = True
        return m

    def add_references(self, lines):
        """Reference= lines the project does not have yet (by their GUID), after its last one"""
        entries = self.vbp.entries
        have = {(v or "").split("#")[0].strip().lower() for k, v in entries if v is not None and k.strip() == "Reference"}
        for line in lines:
            if line.split("#")[0].lower() in have:
                continue
            last = max([i for i, (k, v) in enumerate(entries) if v is not None and k.strip() == "Reference"] or [-1])
            entries.insert(last + 1, ["Reference", line])
            self.dirty = True

    def add_module(self, kind, name=None, add_to_vbp=True):
        if kind in ("DataEnvironment", "DataReport"):
            return self.add_designer(kind, name)
        template = kind
        if kind == "MDIForm":                   # an MDI form is a form (Form= in the .vbp, .frm)
            kind = "Form"
        base = {"Form": "Form", "MDIForm": "MDIForm", "Module": "Module", "Class": "Class",
                "UserControl": "UserControl", "PropertyPage": "PropertyPage"}[template]
        name = name or self.unique_name(base)
        ext = {"Form": ".frm", "Module": ".bas", "Class": ".cls", "UserControl": ".ctl", "PropertyPage": ".pag"}[kind]
        text = {"Form": NEW_FORM, "MDIForm": NEW_MDIFORM, "Module": NEW_MODULE, "Class": NEW_CLASS,
                "UserControl": NEW_UCTL, "PropertyPage": NEW_PAG}[template].replace("{name}", name)
        # ActiveX projects: their classes are MultiUse, their user controls public (VB6's defaults)
        exposed = (kind == "Class" and self.ptype in ("OleExe", "OleDll")) or \
            (kind in ("UserControl", "PropertyPage") and self.ptype == "Control")
        text = text.replace("{exposed}", "True" if exposed else "False")
        sf = parse_source(text.replace("\n", "\r\n"), None, kind, True, self.encoding)
        sf.encoding = self.encoding
        m = Module(self, kind, name, name + ext, sf)
        m.dirty = True
        self.modules.append(m)
        self._vbp_add(m)
        if kind == "Form" and not (self.vbp.get("Startup") or "").strip('"'):
            self.vbp.set("Startup", name, quote=True)
        self.dirty = True
        return m

    # ------------------------------------------------------------------ components (.ocx)
    def components(self):
        """the keys of the control libraries of the project (Object= lines it has and LinBASIC knows)"""
        from .ocxlib import lib_of_object_line
        out = []
        for k, v in self.vbp.entries if self.vbp else []:
            if v is not None and k.strip() == "Object":
                lib = lib_of_object_line(v)
                if lib is not None and lib.key not in out:
                    out.append(lib.key)
        return out

    def used_components(self):
        """the libraries whose controls are on some form"""
        from .ocxlib import control_of
        used = set()
        for m in self.modules:
            if m.sf is None or m.sf.root is None:
                continue
            for c in m.sf.root.walk():
                lc = control_of(c.cls)
                if lc:
                    used.add(lc[0].key)
        return used

    def set_components(self, keys):
        """the Object= lines of the project for these libraries (others it has stay)"""
        from .ocxlib import lib_of_object_line, LIB_BY_KEY
        keys = list(keys)
        entries, have = [], set()
        for k, v in self.vbp.entries:
            if v is not None and k.strip() == "Object":
                lib = lib_of_object_line(v)
                if lib is not None:
                    if lib.key not in keys:
                        continue
                    have.add(lib.key)
            entries.append([k, v])
        last = -1
        for i, (k, v) in enumerate(entries):
            if v is not None and k.strip() in ("Reference", "Object"):
                last = i
        for key in keys:
            if key not in have:
                lib = LIB_BY_KEY[key.lower()]
                last += 1
                entries.insert(last, ["Object", lib.vbp_line])
        if entries != self.vbp.entries:
            self.vbp.entries = entries
            self.dirty = True

    def _vbp_add(self, m):
        entries = self.vbp.entries
        if any(k.strip() == m.kind and v is not None and m.rel.lower() in v.lower() for k, v in entries):
            return
        value = m.rel if m.kind in ("Form", "UserControl", "PropertyPage", "UserDocument", "Designer") else f"{m.name}; {m.rel}"
        # after the last member line, as VB6 writes them
        last = -1
        for i, (k, v) in enumerate(entries):
            if v is not None and k.strip() in ("Reference", "Object", "Form", "Module", "Class", "UserControl",
                                                "PropertyPage", "UserDocument", "Designer"):
                last = i
        entries.insert(last + 1, [m.kind, value])

    def remove_module(self, m):
        self.modules.remove(m)
        self.vbp.entries = [[k, v] for k, v in self.vbp.entries
                            if not (v is not None and k.strip() == m.kind and
                                    (v.strip() == m.rel or v.split(";")[-1].strip() == m.rel))]
        if (self.vbp.get("Startup") or "").strip('"').lower() == m.name.lower():
            self.vbp.set("Startup", "Sub Main", quote=True)
        self.dirty = True

    def module(self, name):
        for m in self.modules:
            if m.name.lower() == name.lower():
                return m
        return None

    def module_by_path(self, path):
        rp = os.path.realpath(path)
        for m in self.modules:
            if os.path.realpath(m.path) == rp:
                return m
        return None


# procedures of a module: Sub / Function / Property, for the Procedure box of the code window
PROC_RE = re.compile(r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)",
                     re.IGNORECASE)


def procedures(text):
    """[(name, kind, line)] in the order of the text"""
    out = []
    for i, line in enumerate(text.split("\n")):
        m = PROC_RE.match(line)
        if m:
            out.append((m.group(2), " ".join(m.group(1).split()).title(), i))
    return out


NEW_VBP = """Type=Exe
Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation
Startup="Form1"
Command32=""
Name="Project1"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
"""

# a new form as VB6 writes it on Windows 7 (DEV-8/me2 Form1.frm): 4800 x 3600 outside, the frame of Windows 7
NEW_FORM = """VERSION 5.00
Begin VB.Form {name} 
   Caption         =   "{name}"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "{name}"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "{name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
"""

# Project > Add MDI Form, as VB6 writes it (DEV-8/me2 MDIForm1.frm): AppWorkspace, no ScaleWidth / ScaleHeight
NEW_MDIFORM = """VERSION 5.00
Begin VB.MDIForm {name} 
   BackColor       =   &H8000000C&
   Caption         =   "{name}"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "{name}"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "{name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
"""

NEW_MODULE = """Attribute VB_Name = "{name}"
"""

NEW_CLASS = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "{name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = {exposed}
"""

NEW_UCTL = """VERSION 5.00
Begin VB.UserControl {name} 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "{name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = {exposed}
"""

NEW_PAG = """VERSION 5.00
Begin VB.PropertyPage {name} 
   Caption         =   "{name}"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "{name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = {exposed}
"""

PROJECT_TYPES = [("Exe", "Standard EXE"), ("OleExe", "ActiveX EXE"), ("OleDll", "ActiveX DLL"),
                 ("Control", "ActiveX Control")]


def new_vbp_text(ptype, name="Project1"):
    """the .vbp VB6 writes for a new project of this type (its member lines come with the modules)"""
    lines = NEW_VBP.replace('Name="Project1"', f'Name="{name}"').split("\n")
    lines[0] = f"Type={ptype}"
    if ptype != "Exe":
        lines = [l for l in lines if not l.startswith("Startup=")]
        lines.insert(2, 'Startup="(None)"')
        lines = [('CompatibleMode="1"' if l.startswith("CompatibleMode=") else
                  "StartMode=1" if l.startswith("StartMode=") else l) for l in lines]   # DEV-8/pp1: ax1, dll1, ocx1
    if ptype in ("OleDll", "Control"):
        lines.insert(len(lines) - 1, "ThreadingModel=1")
    return "\n".join(lines)


class VBGroup:
    """a project group (.vbg): projects edited and run together; the startup project runs on F5,
    the others are its components (ActiveX DLL / EXE / Control)"""

    def __init__(self):
        self.path = None
        self.projects = []
        self.startup = None
        self.dirty = False

    @property
    def name(self):
        return os.path.splitext(os.path.basename(self.path))[0] if self.path else "Group1"

    @property
    def untitled(self):
        return self.path is None

    @classmethod
    def open(cls, path, encoding=None):
        """VBGROUP 5.0, StartupProject=..., Project=... (paths relative to the .vbg)"""
        g = cls()
        g.path = os.path.abspath(path)
        with open(g.path, "rb") as f:
            text = f.read().decode(encoding or default_codepage(), errors="replace")
        from lbvc.project.fsutil import ci_path
        for line in text.splitlines():
            key, sep, val = line.partition("=")
            if not sep or key.strip() not in ("StartupProject", "Project"):
                continue
            p = VBProject.open(ci_path(os.path.join(os.path.dirname(g.path), val.strip().replace("\\", "/"))), encoding)
            g.add(p)
            if key.strip() == "StartupProject":
                g.startup = p
        if g.startup is None and g.projects:
            g.startup = g.projects[0]
        g.dirty = False
        return g

    def add(self, p):
        p.group = self
        self.projects.append(p)
        if self.startup is None and p.ptype in ("Exe", "OleExe"):
            self.startup = p
        self.dirty = True

    def remove(self, p):
        self.projects.remove(p)
        p.group = None
        if self.startup is p:
            self.startup = next((q for q in self.projects if q.ptype in ("Exe", "OleExe")),
                                self.projects[0] if self.projects else None)
        self.dirty = True

    def any_dirty(self):
        return self.dirty or any(p.any_dirty() for p in self.projects)

    def _rel(self, p):
        return os.path.relpath(p.path, os.path.dirname(self.path)).replace("/", "\\")

    def save(self, path=None):
        """the .vbg (the projects save themselves): the startup project as StartupProject="""
        if path:
            self.path = os.path.abspath(path)
        lines = ["VBGROUP 5.0"]
        for p in self.projects:
            lines.append(f"{'StartupProject' if p is self.startup else 'Project'}={self._rel(p)}")
        enc = self.startup.encoding if self.startup is not None else default_codepage()
        with open(self.path, "wb") as f:
            f.write(("\r\n".join(lines) + "\r\n").encode(enc, errors="replace"))
        self.dirty = False

    def write_copy(self, folder, project):
        """every project as it is in the editors into folder, in the same places relative to each other
        (the references *\\A..\\Lib\\Lib.vbp of the startup project find its components); the .vbp of project"""
        saved = [p for p in self.projects if p.path]
        base = os.path.commonpath([os.path.dirname(p.path) for p in saved]) if saved else folder
        out = None
        for p in self.projects:
            sub = os.path.relpath(os.path.dirname(p.path), base) if p.path else p.name
            vbp = p.write_copy(os.path.normpath(os.path.join(folder, sub)))
            if p is project:
                out = vbp
        return out

    def references(self, p):
        """the projects of the group that p refers to (Reference=*\\A / Object=*\\A of its .vbp)"""
        out = []
        for k, v in p.vbp.entries:
            if v is None or k.strip() not in ("Reference", "Object") or not v.strip().startswith("*\\A"):
                continue
            target = os.path.normpath(p.resolve(v.strip()[3:]))
            for q in self.projects:
                if q.path and os.path.normcase(os.path.normpath(q.path)) == os.path.normcase(target):
                    out.append(q)
        return out

    def set_reference(self, p, q, on):
        """Project > References: p refers to the project q of the group (or not any more)"""
        rel = os.path.relpath(q.path, os.path.dirname(p.path)).replace("/", "\\")
        line = "*\\A" + rel
        entries = [e for e in p.vbp.entries
                   if not (e[1] is not None and e[0].strip() == "Reference" and e[1].strip().lower() == line.lower())]
        if on:
            last = max([i for i, (k, v) in enumerate(entries) if v is not None and k.strip() == "Reference"] or [0])
            entries.insert(last + 1, ["Reference", line])
        if entries != p.vbp.entries:
            p.vbp.entries = entries
            p.dirty = True
