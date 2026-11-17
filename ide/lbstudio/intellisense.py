"""IntelliSense of the VB6 code window: List Properties/Methods after a dot, Quick Info and
Parameter Info, List Constants and Complete Word.

What is known about types comes from the VB6 type libraries (data/vb6/typelibs: VB, VBA, VBRUN,
stdole and the libraries the project references) and from the code of the project itself: its
forms and their controls, modules, classes, variables, Type and Enum blocks, Declare statements.
"""
import json
import os
import re

from . import paths

# ---------------------------------------------------------------- type libraries
BASE_LIBS = ["VB", "VBA", "VBRUN", "stdole"]
# libraries of the .vbp Object= / Reference= lines, by the file name they name
LIB_OF_FILE = {"mscomctl.ocx": "MSComctlLib", "comctl32.ocx": "ComctlLib", "mscomct2.ocx": "MSComCtl2",
               "comct332.ocx": "ComCtl3", "comdlg32.ocx": "MSComDlg", "msflxgrd.ocx": "MSFlexGridLib",
               "richtx32.ocx": "RichTextLib", "tabctl32.ocx": "TabDlg", "scrrun.dll": "Scripting",
               "msado15.dll": "ADODB", "mschrt20.ocx": "MSChart20Lib", "picclp32.ocx": "PicClip",
               "mci32.ocx": "MCI", "mscomm32.ocx": "MSCommLib"}

KEYWORDS = ("And As Boolean ByRef Byte ByVal Call Case Close Const Currency Date Declare Dim Do Double Each Else "
            "ElseIf End Enum Erase Error Event Exit False For Friend Function Get Global GoSub GoTo If Implements "
            "In Input Integer Is Let Lib Like Line Load Long Loop LSet Me Mod New Next Not Nothing Object On Open "
            "Option Optional Or ParamArray Preserve Print Private Property Public Put RaiseEvent ReDim Rem Resume "
            "Return RSet Select Set Single Static Step Stop String Sub Then To True Type TypeOf Unload Until "
            "Variant Wend While With WithEvents Write Xor").split()
TYPES = ("Boolean Byte Collection Control Currency Date Double Form Integer Long Object Single String "
         "Variant").split()

# the VB constants that go with an enumerated property of the Properties window
ENUM_OF_PROP = {"appearance": None, "dragmode": "DragModeConstants", "mousepointer": "MousePointerConstants",
                "linkmode": "LinkModeConstants", "formborder": "FormBorderStyleConstants",
                "border": "BorderStyleConstants", "align3": "AlignmentConstants", "align2": "AlignmentConstants",
                "scrollbars": "ScrollBarConstants", "windowstate": "FormWindowStateConstants",
                "checkvalue": "CheckBoxConstants", "fillstyle": "FillStyleConstants",
                "drawstyle": "DrawStyleConstants", "drawmode": "DrawModeConstants",
                "scalemode": "ScaleModeConstants", "combostyle": "ComboBoxConstants",
                "multiselect": "MultiSelectConstants", "shape": "ShapeConstants", "align": "AlignConstants",
                "buttonstyle": "ButtonConstants", "liststyle": "ListBoxConstants",
                "startup": "StartUpPositionConstants", "palettemode": "PaletteModeConstants",
                "oledragmode": "OLEDragConstants", "oledropmode": "OLEDropConstants"}


class Member:
    """one entry of a list: kind is prop, method, event, const, enum, type, module, control, keyword"""
    __slots__ = ("name", "kind", "params", "returns", "help", "lib", "value")

    def __init__(self, name, kind, params=None, returns=None, help=None, lib=None, value=None):
        self.name, self.kind, self.params, self.returns = name, kind, params or [], returns
        self.help, self.lib, self.value = help, lib, value

    def signature(self):
        """the text of Quick Info: MsgBox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], ...) As VbMsgBoxResult"""
        if self.kind == "const":
            return f"{self.name} = {self.value}" if self.value is not None else self.name
        ps = ", ".join(p_text(p) for p in self.params)
        head = f"{self.name}({ps})" if self.params or self.kind in ("method", "func") else self.name
        if self.returns and self.returns not in ("HRESULT", "void"):
            head += f" As {short(self.returns)}"
        return head


def short(t):
    t = str(t).split("\\")[-1]
    return t.split(".")[-1]


ENUM_NAMES = {}            # enum lower -> {value: const name}, filled by TypeLibs


def p_text(p):
    """a parameter as VB6 shows it: Buttons As VbMsgBoxStyle = vbOKOnly (no ByVal, no As Variant),
    ParamArray ArgList() As Variant, Args() As Variant"""
    s = p["name"]
    t = short(p["type"]) if p.get("type") else None
    if p.get("paramarray"):
        return f"ParamArray {s}() As {(t or 'Variant()')[:-2] if (t or '').endswith('()') else 'Variant'}"
    if t and t.endswith("()"):
        s, t = s + "()", t[:-2]
    if t and t.lower() != "variant":
        s += f" As {t}"
    d = p.get("default")
    if d is not None:
        if t and isinstance(d, int):
            d = ENUM_NAMES.get(t.lower(), {}).get(d, d)
        s += f" = {d}"
    if p.get("optional"):
        s = f"[{s}]"
    return s


def _vb_name(n):
    """the names VBA hides behind _B_str_ / _B_var_: Chr$ and Chr"""
    if n.startswith("_B_str_"):
        return n[7:] + "$"
    if n.startswith("_B_var_"):
        return n[7:]
    return n


class TypeLibs:
    """the VB6 type libraries (JSON from tools/rea/dump_typelibs.py)"""
    _cache = {}

    def __init__(self, extra=()):
        self.libs = {}
        for lib in BASE_LIBS + [x for x in extra if x not in BASE_LIBS]:
            d = self._load(lib)
            if d is not None:
                self.libs[lib] = d
        self.types = {}              # lower name -> (lib, type dict)
        for lib, d in self.libs.items():
            for t in d["types"]:
                self.types.setdefault(t["name"].lower(), (lib, t))
                self.types.setdefault(f"{lib}.{t['name']}".lower(), (lib, t))
        for lib, d in self.libs.items():
            for t in d["types"]:
                if t["kind"] == "enum":
                    vals = ENUM_NAMES.setdefault(t["name"].lower(), {})
                    for m in t.get("members", []):
                        vals.setdefault(m.get("value"), m["name"])

    @classmethod
    def _load(cls, lib):
        if lib not in cls._cache:
            path = os.path.join(paths.VB6_DATA, "typelibs", lib + ".json")
            try:
                with open(path, encoding="utf-8") as f:
                    cls._cache[lib] = json.load(f)
            except OSError:
                cls._cache[lib] = None
        return cls._cache[lib]

    def find(self, name):
        return self.types.get(str(name).lower())

    def interface(self, name):
        """a coclass -> its default interface; aliases followed"""
        hit = self.find(name)
        seen = 0
        while hit is not None and seen < 5:
            lib, t = hit
            seen += 1
            if t["kind"] == "coclass":
                imp = [i for i in t.get("implements", []) if i.get("default") and not i.get("source")]
                hit = self.find(f"{lib}.{imp[0]['name']}") if imp else None
                continue
            if t["kind"] == "alias":
                hit = self.find(t.get("type") or t.get("alias") or "")
                continue
            return hit
        return hit

    def events_of(self, name):
        hit = self.find(name)
        if hit and hit[1]["kind"] == "coclass":
            lib, t = hit
            src = [i for i in t.get("implements", []) if i.get("default") and i.get("source")]
            return self.find(f"{lib}.{src[0]['name']}") if src else None
        return None

    def members(self, name, events=False):
        """Member list of a type (properties and methods; events when asked)"""
        hit = self.interface(name)
        out = {}
        if hit is None:
            return []
        lib, t = hit
        if t["kind"] == "enum":
            return [Member(m["name"], "const", value=m.get("value"), help=m.get("helpcontext"), lib=lib)
                    for m in t.get("members", [])]
        for m in t.get("members", []):
            n = m["name"]
            flags = m.get("flags", [])
            if "hidden" in flags or "restricted" in flags or n.startswith("_"):
                if not n.startswith("_B_"):
                    continue
            n = _vb_name(n)
            inv = m.get("invoke")
            if m.get("kind") == "const":
                out[n.lower()] = Member(n, "const", value=m.get("value"), lib=lib, help=m.get("helpcontext"))
                continue
            params = vb_params(m)
            if inv in ("get", "let", "set") or m.get("kind") == "var":
                ret = m.get("returns") if inv == "get" else (m.get("type") or (m["params"][-1]["type"] if m.get("params") else None))
                old = out.get(n.lower())
                if old is None or inv == "get":
                    out[n.lower()] = Member(n, "prop", params if inv == "get" else [], ret, m.get("helpcontext"), lib)
            else:
                ret = m.get("returns")
                out[n.lower()] = Member(n, "method", params, ret, m.get("helpcontext"), lib)
        res = list(out.values())
        if events:
            ev = self.events_of(name)
            if ev:
                for m in ev[1].get("members", []):
                    res.append(Member(m["name"], "event", vb_params(m), None,
                                      m.get("helpcontext"), ev[0]))
        return res

    def globals(self):
        """what the code can name without a dot: VBA functions and constants, the VB Global
        object (App, Screen, Load ...), the enumerations of every library"""
        out = {}
        for lib, d in self.libs.items():
            for t in d["types"]:
                k = t["kind"]
                if k == "module" and "hidden" not in t.get("flags", []) or t["name"] == "_HiddenModule":
                    for mm in self.members(f"{lib}.{t['name']}"):
                        out.setdefault(mm.name.lower(), mm)
                elif k == "enum" and "hidden" not in t.get("flags", []):
                    out.setdefault(t["name"].lower(), Member(t["name"], "enum", lib=lib, help=t.get("helpcontext")))
                    for m in t.get("members", []):
                        if not m["name"].startswith("_"):
                            out.setdefault(m["name"].lower(), Member(m["name"], "const", value=m.get("value"),
                                                                     lib=lib, help=m.get("helpcontext")))
        for mm in self.members("VB.Global"):
            out.setdefault(mm.name.lower(), mm)
        for name, t in (("Err", "VBA.ErrObject"), ("Debug", None)):
            out.setdefault(name.lower(), Member(name, "prop", returns=t or "Debug"))
        return out

    def type_names(self):
        """for As: the creatable classes and the types of the libraries"""
        out = set(TYPES)
        for lib, d in self.libs.items():
            for t in d["types"]:
                if t["kind"] in ("coclass", "enum", "record") and "hidden" not in t.get("flags", []) \
                        and not t["name"].startswith("_"):
                    out.add(t["name"])
            out.add(lib)
        return sorted(out, key=str.lower)


def _param(p):
    fl = p.get("flags", [])
    return {"name": p["name"], "type": p.get("type"), "byval": not p.get("byref", True),
            "optional": "optional" in fl or p.get("opt", False), "paramarray": p.get("paramarray", False),
            "default": p.get("default") if "hasdefault" in fl else None}


def vb_params(m):
    """the parameters of a member of a type library as VB sees them (no retval, no lcid): the last optcount are
    optional (cParamsOpt), optcount -1 makes the last one a ParamArray"""
    ps = [dict(p) for p in m.get("params", []) if "retval" not in p.get("flags", []) and "lcid" not in p.get("flags", [])]
    n = m.get("optcount") or 0
    if n == -1 and ps:
        ps[-1]["paramarray"] = True
    elif n > 0:
        for p in ps[-n:]:
            p["opt"] = True
    return [_param(p) for p in ps]


DEBUG_MEMBERS = [Member("Assert", "method", [{"name": "Expression", "type": "Boolean"}]),
                 Member("Print", "method", [{"name": "OutputList", "type": None, "optional": True}])]


# ---------------------------------------------------------------- the code of the project
DECL_RE = re.compile(r"^\s*(?:(Public|Private|Global|Dim|Static|Friend)\s+)?(?:(WithEvents)\s+)?", re.I)
VAR_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_]*)(\$|%|&|!|#|@)?(?:\s*(\([^)]*\)))?(?:\s+As\s+(New\s+)?([A-Za-z_][A-Za-z0-9_.]*))?", re.I)
PROC_RE = re.compile(r"^\s*(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
                     r"(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set|Event|Declare\s+(?:Sub|Function))\s+"
                     r"([A-Za-z_][A-Za-z0-9_]*)([$%&!#@])?(.*)$", re.I)
SUFFIX_TYPE = {"$": "String", "%": "Integer", "&": "Long", "!": "Single", "#": "Double", "@": "Currency"}


def strip_comment(line):
    q = False
    for i, ch in enumerate(line):
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            return line[:i]
    if re.match(r"^\s*Rem\b", line, re.I):
        return ""
    return line


def split_top(text, sep=","):
    out, depth, cur, q = [], 0, "", False
    for ch in text:
        if ch == '"':
            q = not q
        if not q:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == sep and depth == 0:
                out.append(cur)
                cur = ""
                continue
        cur += ch
    out.append(cur)
    return out


def parse_params(text):
    """'(ByVal a As Long, Optional b As String = "x") As Integer' -> ([params], 'Integer')"""
    text = text.strip()
    m = re.match(r"^(?:Lib\s+\"[^\"]*\"\s*(?:Alias\s+\"[^\"]*\"\s*)?)?\((.*)\)\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?", text, re.I)
    if not m:
        r = re.search(r"As\s+([A-Za-z_][A-Za-z0-9_.]*)\s*$", text, re.I)
        return [], r.group(1) if r else None
    ps = []
    for part in split_top(m.group(1)):
        part = part.strip()
        if not part:
            continue
        pm = re.match(r"^(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?([A-Za-z_][A-Za-z0-9_]*)([$%&!#@])?(\(\))?"
                      r"(?:\s+As\s+([A-Za-z_][A-Za-z0-9_.]*))?(?:\s*=\s*(.*))?$", part, re.I)
        if not pm:
            continue
        ps.append({"name": pm.group(4), "type": pm.group(7) or SUFFIX_TYPE.get(pm.group(5) or "", None),
                   "byval": bool(pm.group(2) and pm.group(2).lower().startswith("byval")),
                   "optional": bool(pm.group(1)), "default": pm.group(8)})
    return ps, m.group(2)


class CodeInfo:
    """what the code of one module declares"""

    def __init__(self, text):
        self.vars = {}           # lower name -> (name, type, scope 'Public'|'Private', is_new)
        self.procs = {}          # lower name -> Member (method / prop / event)
        self.proc_lines = []     # (start, end, name, params text)
        self.types = {}          # UDT lower name -> [Member]
        self.enums = {}          # enum lower name -> (name, [Member], public)
        self.public = set()      # lower names of public procs/vars
        self.consts = {}
        self._parse(text)

    def _parse(self, text):
        lines = text.replace("\r\n", "\n").split("\n")
        # join line continuations
        joined, starts, cur, start = [], [], "", 0
        for i, l in enumerate(lines):
            s = strip_comment(l)
            if not cur:
                start = i
            if s.rstrip().endswith(" _"):
                cur += s.rstrip()[:-1] + " "
                continue
            cur += s
            joined.append(cur)
            starts.append(start)
            cur = ""
        in_proc, block, block_name, block_items, block_public = None, None, None, [], True
        for idx, line in enumerate(joined):
            ln = starts[idx]
            t = line.strip()
            low = t.lower()
            if block:
                if re.match(r"^end\s+(type|enum)\b", low):
                    if block == "type":
                        self.types[block_name.lower()] = block_items
                    else:
                        self.enums[block_name.lower()] = (block_name, block_items, block_public)
                    block = None
                    continue
                if block == "type":
                    m = VAR_RE.match(t)
                    if m:
                        block_items.append(Member(m.group(1), "prop", returns=m.group(5) or SUFFIX_TYPE.get(m.group(2) or "", "Variant")))
                else:
                    m = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)(?:\s*=\s*(.*))?$", t)
                    if m:
                        block_items.append(Member(m.group(1), "const", value=m.group(2)))
                continue
            if in_proc is not None:
                if re.match(r"^end\s+(sub|function|property)\b", low):
                    self.proc_lines.append((in_proc[0], ln, in_proc[1], in_proc[2]))
                    in_proc = None
                continue
            m = re.match(r"^(?:(public|private|global)\s+)?(type|enum)\s+([A-Za-z_][A-Za-z0-9_]*)\s*$", low)
            if m:
                block, block_name, block_items = m.group(2), t.split()[-1], []
                block_public = m.group(1) != "private"
                continue
            m = PROC_RE.match(t)
            if m:
                scope, kind, name, suf, rest = m.groups()
                params, ret = parse_params(rest)
                kind_l = kind.lower()
                if kind_l.startswith("declare"):
                    mk = "method"
                elif kind_l == "event":
                    mk = "event"
                elif kind_l.startswith("property"):
                    mk = "prop"
                else:
                    mk = "method"
                ret = ret or SUFFIX_TYPE.get(suf or "", None)
                old = self.procs.get(name.lower())
                if old is None or kind_l == "property get":
                    self.procs[name.lower()] = Member(name, mk, params if mk != "prop" or kind_l == "property get" else [],
                                                      ret if not kind_l.startswith("sub") else None)
                if (scope or "public").lower() != "private":
                    self.public.add(name.lower())
                if not kind_l.startswith("declare") and kind_l != "event":
                    in_proc = (ln, name, rest)
                continue
            m = re.match(r"^(?:(public|private|global)\s+)?const\s+(.*)$", t, re.I)
            if m:
                for part in split_top(m.group(2)):
                    cm = re.match(r"^\s*([A-Za-z_][A-Za-z0-9_]*)[$%&!#@]?(?:\s+As\s+(\w+))?\s*=\s*(.*)$", part)
                    if cm:
                        self.consts[cm.group(1).lower()] = Member(cm.group(1), "const", value=cm.group(3).strip(), returns=cm.group(2))
                        if (m.group(1) or "private").lower() != "private":
                            self.public.add(cm.group(1).lower())
                continue
            m = re.match(r"^(public|private|global|dim)\s+(withevents\s+)?(.*)$", t, re.I)
            if m and not re.match(r"^(public|private)\s+(sub|function|property|declare|event|type|enum|const)\b", t, re.I):
                for part in split_top(m.group(3)):
                    vm = VAR_RE.match(part.strip())
                    if vm:
                        typ = vm.group(5) or SUFFIX_TYPE.get(vm.group(2) or "", "Variant")
                        self.vars[vm.group(1).lower()] = (vm.group(1), typ, m.group(1).lower(), bool(vm.group(4)))
                        if m.group(1).lower() in ("public", "global"):
                            self.public.add(vm.group(1).lower())
        if in_proc is not None:
            self.proc_lines.append((in_proc[0], len(lines), in_proc[1], in_proc[2]))

    def proc_at(self, line):
        for s, e, name, rest in self.proc_lines:
            if s <= line <= e:
                return s, e, name, rest
        return None


def local_vars(text, line):
    """the parameters and Dim/Static/ReDim of the procedure that has the line"""
    info = CodeInfo(text)
    hit = info.proc_at(line)
    if not hit:
        return {}
    s, e, name, rest = hit
    out = {}
    params, ret = parse_params(rest)
    for p in params:
        out[p["name"].lower()] = (p["name"], p["type"] or "Variant")
    lines = text.replace("\r\n", "\n").split("\n")[s + 1:min(e, line + 1)]
    for l in lines:
        t = strip_comment(l).strip()
        m = re.match(r"^(Dim|Static|ReDim(?:\s+Preserve)?)\s+(.*)$", t, re.I)
        if m:
            for part in split_top(m.group(2)):
                vm = VAR_RE.match(part.strip())
                if vm:
                    out[vm.group(1).lower()] = (vm.group(1), vm.group(5) or SUFFIX_TYPE.get(vm.group(2) or "", "Variant"))
    if ret:
        out[name.lower()] = (name, ret)            # the return value of a Function
    return out


# ---------------------------------------------------------------- the engine
class IntelliSense:
    def __init__(self, project):
        self.project = project
        extra = []
        if project is not None and project.vbp is not None:
            for k, v in project.vbp.entries:
                if v is None or k.strip() not in ("Object", "Reference"):
                    continue
                low = v.lower()
                for f, lib in LIB_OF_FILE.items():
                    if f in low:
                        extra.append(lib)
        self.tl = TypeLibs(extra)
        self._globals = None
        self._infos = {}

    # ------------------------------------------------------------ project
    def info(self, m):
        text = m.code_text()
        key = id(m)
        hit = self._infos.get(key)
        if hit is None or hit[0] != text:
            hit = self._infos[key] = (text, CodeInfo(text))
        return hit[1]

    def globals(self):
        if self._globals is None:
            self._globals = self.tl.globals()
        return self._globals

    def module_named(self, name):
        if self.project is None:
            return None
        return self.project.module(name)

    def controls(self, m):
        """name lower -> (name, type, is_array) of a form"""
        if m is None or not m.is_form:
            return {}
        return {n.lower(): (n, t, arr) for n, t, arr in m.controls()}

    # ------------------------------------------------------------ resolving
    def type_of_name(self, name, m, line, text):
        """the type of a name at a place of the code: ('lib', 'VB.TextBox') / ('module', mod) / ('udt', [..]) ..."""
        low = name.lower()
        if low == "me" and m is not None:
            return ("module", m)
        loc = local_vars(text, line) if text is not None else {}
        if low in loc:
            return self.type_of_typename(loc[low][1], m)
        ci = self.info(m) if m is not None else None
        if ci is not None and low in ci.vars:
            return self.type_of_typename(ci.vars[low][1], m)
        ctls = self.controls(m)
        if low in ctls:
            n, t, arr = ctls[low]
            return ("control", "VB." + t if t in _VB_CONTROLS else t, arr)
        if self.project is not None:
            for mm in self.project.modules:
                ci2 = self.info(mm)
                if mm.kind == "Module" and low in ci2.vars and low in ci2.public:
                    return self.type_of_typename(ci2.vars[low][1], mm)
            mod = self.module_named(name)
            if mod is not None:
                return ("module", mod)
        if low in ("app", "screen", "printer", "clipboard"):
            return ("lib", "VB." + name.capitalize() if low != "app" else "VB.App")
        if low == "err":
            return ("lib", "VBA.ErrObject")
        if low == "debug":
            return ("debug",)
        if low == "forms":
            return ("lib", "Object")
        if low in self.tl.types:
            lib, t = self.tl.types[low]
            if t["kind"] == "enum":
                return ("lib", f"{lib}.{t['name']}")
            if t["kind"] == "module" or low in [x.lower() for x in self.tl.libs]:
                return ("lib", f"{lib}.{t['name']}")
        if name in self.tl.libs:
            return ("libname", name)
        # a user Enum or Type of the project
        for mm in ([m] if m is not None else []) + (self.project.modules if self.project else []):
            ci2 = self.info(mm)
            if low in ci2.enums:
                return ("enum", ci2.enums[low][1])
        return None

    def type_of_typename(self, tname, m):
        if not tname:
            return None
        low = tname.lower()
        if low in ("integer", "long", "string", "boolean", "byte", "single", "double", "currency", "date", "variant"):
            return None
        for mm in ([m] if m is not None else []) + (self.project.modules if self.project else []):
            ci = self.info(mm)
            if low in ci.types:
                return ("udt", ci.types[low])
            if low in ci.enums:
                return ("enum", ci.enums[low][1])
        mod = self.module_named(tname.split(".")[-1])
        if mod is not None and mod.kind in ("Class", "Form"):
            return ("module", mod)
        if low in ("form",):
            return ("lib", "VB.Form")
        if low in ("control", "object"):
            return None
        if low == "collection":
            return ("lib", "VBA.Collection")
        if self.tl.find(tname) is not None:
            return ("lib", tname)
        return None

    def members_of(self, typ, m):
        """the Member list of a resolved type"""
        if typ is None:
            return []
        k = typ[0]
        if k == "lib":
            return self.tl.members(typ[1])
        if k == "control":
            return self.tl.members(typ[1])
        if k == "debug":
            return list(DEBUG_MEMBERS)
        if k == "udt" or k == "enum":
            return list(typ[1])
        if k == "libname":
            out = []
            d = self.tl.libs.get(typ[1])
            for t in d["types"] if d else []:
                if t["kind"] in ("coclass", "enum", "module") and "hidden" not in t.get("flags", []) and not t["name"].startswith("_"):
                    out.append(Member(t["name"], "type" if t["kind"] != "enum" else "enum", lib=typ[1]))
            return out
        if k == "module":
            mod = typ[1]
            ci = self.info(mod)
            out = []
            own = mod is m
            if mod.is_form:
                base = "VB.MDIForm" if mod.display_kind == "MDIForm" else "VB.Form"
                out += self.tl.members(base)
                for n, t, arr in mod.controls():
                    out.append(Member(n, "control", returns=t))
            for low, mem in ci.procs.items():
                if mem.kind == "event":
                    continue
                if own or low in ci.public:
                    out.append(mem)
            for low, (n, t, scope, new) in ci.vars.items():
                if own or low in ci.public:
                    out.append(Member(n, "prop", returns=t))
            if mod.kind == "Module":
                for low, c in ci.consts.items():
                    if low in ci.public:
                        out.append(c)
            return out
        return []

    def member_type(self, typ, name, m):
        """the type of typ.name (for a chain: Text1.Font.Name)"""
        for mem in self.members_of(typ, m):
            if mem.name.lower() == name.lower():
                if mem.kind == "control":
                    t = mem.returns
                    return ("control", "VB." + t if t in _VB_CONTROLS else t, False)
                if mem.returns:
                    r = mem.returns
                    if r in ("stdole.Font", "Font", "IFontDisp", "StdFont"):
                        return ("lib", "stdole.Font")
                    if r in ("stdole.Picture", "Picture", "IPictureDisp", "StdPicture"):
                        return ("lib", "stdole.Picture")
                    return self.type_of_typename(r, typ[1] if typ[0] == "module" else m)
                return None
        return None

    def resolve(self, chain, m, line, text):
        """['Me', 'Text1', 'Font'] -> the type of the last part"""
        if not chain:
            return None
        typ = self.type_of_name(chain[0], m, line, text)
        for part in chain[1:]:
            if typ is None:
                return None
            typ = self.member_type(typ, part, m)
        return typ

    # ------------------------------------------------------------ the lists
    def list_members(self, chain, m, line, text):
        typ = self.resolve(chain, m, line, text)
        mems = self.members_of(typ, m)
        uniq = {}
        for x in mems:
            uniq.setdefault(x.name.lower(), x)
        return sorted(uniq.values(), key=lambda x: x.name.lower())

    def complete_words(self, m, line, text):
        """Complete Word: everything that can be named here"""
        out = dict(self.globals())
        for k in KEYWORDS:
            out.setdefault(k.lower(), Member(k, "keyword"))
        if m is not None:
            ci = self.info(m)
            for low, mem in ci.procs.items():
                out.setdefault(low, mem)
            for low, (n, t, scope, new) in ci.vars.items():
                out.setdefault(low, Member(n, "prop", returns=t))
            for low, c in ci.consts.items():
                out.setdefault(low, c)
            for n, t, arr in m.controls():
                out.setdefault(n.lower(), Member(n, "control", returns=t))
            for low, (n, t) in local_vars(text, line).items():
                out.setdefault(low, Member(n, "prop", returns=t))
        if self.project is not None:
            for mm in self.project.modules:
                out.setdefault(mm.name.lower(), Member(mm.name, "module"))
                ci = self.info(mm)
                if mm.kind == "Module":
                    for low in ci.public:
                        mem = ci.procs.get(low) or ci.consts.get(low)
                        if mem is None and low in ci.vars:
                            mem = Member(ci.vars[low][0], "prop", returns=ci.vars[low][1])
                        if mem is not None:
                            out.setdefault(low, mem)
                for low, (n, items, public) in ci.enums.items():
                    out.setdefault(low, Member(n, "enum"))
                    for c in items:
                        out.setdefault(c.name.lower(), c)
                for low in ci.types:
                    out.setdefault(low, Member(low, "type"))
        return sorted(out.values(), key=lambda x: x.name.lower())

    def type_list(self):
        """after As (and New)"""
        out = [Member(n, "type") for n in self.tl.type_names()]
        if self.project is not None:
            for mm in self.project.modules:
                if mm.kind in ("Class", "Form"):
                    out.append(Member(mm.name, "module"))
                ci = self.info(mm)
                for low, (n, items, public) in ci.enums.items():
                    out.append(Member(n, "enum"))
                for low, items in ci.types.items():
                    out.append(Member(low, "type"))
        uniq = {}
        for x in out:
            uniq.setdefault(x.name.lower(), x)
        return sorted(uniq.values(), key=lambda x: x.name.lower())

    def callable(self, chain, m, line, text):
        """the Member of a function or method call (for Parameter Info)"""
        if not chain:
            return None
        name = chain[-1]
        if len(chain) == 1:
            low = name.lower()
            ci = self.info(m) if m is not None else None
            if ci is not None and low in ci.procs:
                return ci.procs[low]
            if self.project is not None:
                for mm in self.project.modules:
                    ci2 = self.info(mm)
                    if mm.kind == "Module" and low in ci2.procs and low in ci2.public:
                        return ci2.procs[low]
            g = self.globals().get(low)
            if g is not None and g.kind in ("method", "prop"):
                return g
            if m is not None and m.is_form:
                for mem in self.tl.members("VB.Form"):
                    if mem.name.lower() == low:
                        return mem
            return None
        typ = self.resolve(chain[:-1], m, line, text)
        for mem in self.members_of(typ, m):
            if mem.name.lower() == name.lower():
                return mem
        return None

    def constants_for(self, type_name):
        """List Constants: the members of an enumeration (a parameter or a variable of that type)"""
        if not type_name:
            return []
        low = short(type_name).lower()
        if low == "boolean":
            return [Member("False", "const"), Member("True", "const")]
        hit = self.tl.find(short(type_name))
        if hit is not None and hit[1]["kind"] == "enum":
            return [Member(x["name"], "const", value=x.get("value")) for x in hit[1].get("members", [])
                    if not x["name"].startswith("_")]
        if self.project is not None:
            for mm in self.project.modules:
                ci = self.info(mm)
                if low in ci.enums:
                    return list(ci.enums[low][1])
        return []

    def property_constants(self, typ, prop):
        """Text1.Alignment = : the constants of VB that go with the property"""
        from . import vbprops as V
        if typ is None or typ[0] not in ("control", "module", "lib"):
            return []
        tname = short(typ[1]) if typ[0] != "module" else ("MDIForm" if typ[1].display_kind == "MDIForm" else "Form")
        p = V.prop_def(tname, prop)
        if p is None:
            return []
        if p.kind == "bool":
            return self.constants_for("Boolean")
        if p.kind == "enum":
            return self.constants_for(ENUM_OF_PROP.get(p.enum))
        if p.kind == "color":
            return self.constants_for("ColorConstants") + self.constants_for("SystemColorConstants")
        return []


_VB_CONTROLS = {"PictureBox", "Label", "TextBox", "Frame", "CommandButton", "CheckBox", "OptionButton", "ComboBox",
                "ListBox", "HScrollBar", "VScrollBar", "Timer", "DriveListBox", "DirListBox", "FileListBox", "Menu",
                "Shape", "Line", "Image", "Data", "OLE"}


# ---------------------------------------------------------------- reading the line at the caret
IDENT = re.compile(r"[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?")


def chain_before(line_text):
    """'   x = Me.Text1(0).Fo' -> (['Me', 'Text1'], 'Fo'): the object chain before the last dot and
    the part of the member already typed; None when the caret is not after a dot"""
    s = line_text
    m = re.search(r"([A-Za-z_][A-Za-z0-9_]*)?$", s)
    typed = m.group(1) or ""
    i = len(s) - len(typed)
    if i == 0 or s[i - 1] != ".":
        return None
    parts = []
    j = i - 1
    while True:
        j -= 1
        # skip an argument list (Text1(0))
        while j >= 0 and s[j] == " ":
            j -= 1
        if j >= 0 and s[j] == ")":
            depth = 0
            while j >= 0:
                if s[j] == ")":
                    depth += 1
                elif s[j] == "(":
                    depth -= 1
                    if depth == 0:
                        break
                j -= 1
            j -= 1
        k = j
        while k >= 0 and (s[k].isalnum() or s[k] == "_"):
            k -= 1
        word = s[k + 1:j + 1]
        if not word:
            break
        parts.insert(0, word)
        if k >= 0 and s[k] == ".":
            j = k
            continue
        break
    if not parts:
        # With block: .Name -> the With object
        return ["<with>"], typed
    return parts, typed


def call_context(text_before):
    """the call around the caret: (chain of the callee, index of the argument) or None.
    Knows both MsgBox("a", vbOK and MsgBox "a", vbOK (a statement without parentheses)."""
    s = text_before
    depth, q, argi = 0, False, 0
    for i in range(len(s) - 1, -1, -1):
        ch = s[i]
        if ch == '"':
            q = not q
            continue
        if q:
            continue
        if ch == ")":
            depth += 1
        elif ch == "(":
            if depth == 0:
                head = s[:i].rstrip()
                cm = re.search(r"((?:[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?\.)*[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?)$", head)
                if cm:
                    return [p for p in cm.group(1).split(".")], argi
                return None
            depth -= 1
        elif ch == "," and depth == 0:
            argi += 1
    # a statement call: Name arg, arg
    m = re.match(r"^\s*(?:Call\s+)?((?:[A-Za-z_][A-Za-z0-9_]*\.)*[A-Za-z_][A-Za-z0-9_]*[$]?)\s+(?![=(])(.*)$", s)
    if m and m.group(1).split(".")[-1].lower() not in [k.lower() for k in KEYWORDS]:
        rest = m.group(2)
        n = sum(1 for x in split_top(rest)) - 1
        return m.group(1).split("."), n
    return None


def with_object(text, line):
    """the object of the innermost With block that is open at the line"""
    stack = []
    for l in text.replace("\r\n", "\n").split("\n")[:line + 1]:
        t = strip_comment(l).strip()
        m = re.match(r"^With\s+(.+)$", t, re.I)
        if m:
            stack.append(m.group(1).strip())
        elif re.match(r"^End\s+With\b", t, re.I) and stack:
            stack.pop()
    return stack[-1] if stack else None


def split_chain(expr):
    """'Me.Text1(0).Font' -> ['Me', 'Text1', 'Font']"""
    out, depth, cur = [], 0, ""
    for ch in expr:
        if ch == "(":
            depth += 1
            continue
        if ch == ")":
            depth -= 1
            continue
        if depth:
            continue
        if ch == ".":
            out.append(cur.strip())
            cur = ""
        else:
            cur += ch
    out.append(cur.strip())
    return [x for x in out if x]
