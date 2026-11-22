"""The code of the ActiveX Control Interface Wizard (phase 9.2d), without Qt.

What VB6's wizard does (VB6 oracle DEV-9/ax1, ax2): the public members of a user control - standard ones (the
properties, methods and events of the UserControl and of the controls on it) and custom ones - each either mapped
to a member of a control on it (or of the UserControl itself) or kept by the control (a property in a variable
m_<Name> with its default m_def_<Name>, saved in the PropertyBag). Over each member it writes a line that says what
it made, and reads it back when it runs again:

    'MemberInfo=<type>,<run time>,<design time>,<default>     a property the control keeps
    'MemberInfo=<type>                                         a method (5: no type, a Sub)
    'MappingInfo=<control>,<control>,-1,<member>               a mapped member (on the Event line of an event)

<type> is the number of the type in Boolean, Byte, Currency, Date, Double, (none), Font, Integer, Long, Object,
OLE_COLOR, Picture, Single, String, Variant; run time 0 Read/Write, 1 Read Only, 2 Write Only, 3 Not Available;
design time 0 Read/Write, 1 Read Only, 2 Not Available. When it runs again, what did not change stays as it is,
the old code of what changed or went out becomes comments ('), the new code is added. The comments, descriptions and
the summary are our own text (VB6's are Microsoft's).
"""
import re

TYPE_CODES = ["Boolean", "Byte", "Currency", "Date", "Double", "", "Font", "Integer", "Long", "Object", "OLE_COLOR",
              "Picture", "Single", "String", "Variant"]
DATA_TYPES = ["AmbientProperties", "Boolean", "Byte", "ContainedControls", "Control", "Currency", "DataMembers",
              "DataSource", "Date", "Double", "Font", "Hyperlink", "Integer", "IStdDataFormat", "Long", "Object",
              "OLE_COLOR", "ParentControls", "Picture", "Single", "String", "Variant"]
RUN_TIME = ["Read/Write", "Read Only", "Write Only", "Not Available"]
DESIGN_TIME = ["Read/Write", "Read Only", "Not Available"]
# the members a new control starts with (DEV-9/ax1-03)
DEFAULT_SELECTED = ["BackColor", "BackStyle", "BorderStyle", "Click", "DblClick", "Enabled", "Font", "ForeColor",
                    "KeyDown", "KeyPress", "KeyUp", "MouseDown", "MouseMove", "MouseUp", "Refresh"]
OBJECT_TYPES = {"font", "picture", "object", "ambientproperties", "containedcontrols", "control", "datamembers",
                "datasource", "hyperlink", "istddataformat", "parentcontrols"}
COLOR_PROPS = {"backcolor", "forecolor", "fillcolor", "bordercolor", "maskcolor"}
WIZ_COMMENT = "' made by the ActiveX Control Interface Wizard (it reads the next line when it runs again)"
# our own short descriptions of the standard members (VB_Description)
DESCRIPTIONS = {
    "BackColor": "The color of the background.", "ForeColor": "The color of the text and of what is drawn.",
    "BackStyle": "Whether the background is transparent or opaque.", "BorderStyle": "The border of the control.",
    "Enabled": "Whether the control answers the user.", "Font": "The font of the text.",
    "Refresh": "Draws the control again.", "Text": "The text of the control.", "Caption": "The caption of the control.",
    "Click": "The user clicked the control.", "DblClick": "The user double-clicked the control.",
    "KeyDown": "A key went down while the control had the focus.", "KeyUp": "A key went up while the control had the focus.",
    "KeyPress": "The user typed a character.", "MouseDown": "A mouse button went down over the control.",
    "MouseMove": "The mouse moved over the control.", "MouseUp": "A mouse button went up over the control.",
    "Change": "The contents of the control changed.",
}


def is_object(t):
    return (t or "").lower() in OBJECT_TYPES


def type_code(t):
    for i, n in enumerate(TYPE_CODES):
        if n.lower() == (t or "").lower():
            return i
    return 9 if is_object(t) else 14


def _short_type(t):
    t = (t or "").split("\\")[-1].split(".")[-1]
    return {"StdFont": "Font", "IFontDisp": "Font", "StdPicture": "Picture", "IPictureDisp": "Picture",
            "HRESULT": ""}.get(t, t)


# ==================================================================== the members of a control type (VB.json)
class TypeInfo:
    """the properties, methods and events of a VB type (UserControl, TextBox ...) from the type library"""

    def __init__(self, type_name):
        from . import vbtypes
        self.props, self.methods, self.events = {}, {}, {}
        t = vbtypes._types().get("_" + type_name.lower())
        for m in (t or {}).get("members", []):
            n = m.get("name", "")
            if not n or n.startswith("_") or "hidden" in (m.get("flags") or []) or \
                    "restricted" in (m.get("flags") or []):
                continue
            inv = m.get("invoke")
            ps = [p for p in m.get("params", []) if "retval" not in (p.get("flags") or [])]
            if inv in ("get", "let", "set"):
                d = self.props.setdefault(n, {"type": "", "get": False, "let": False, "set": False, "args": []})
                d[inv] = True
                if inv == "get":
                    rv = [p for p in m.get("params", []) if "retval" in (p.get("flags") or [])]
                    d["type"] = _short_type(rv[0]["type"] if rv else m.get("returns"))
                    d["args"] = ps
                elif not d["type"] and ps:
                    d["type"] = _short_type(ps[-1]["type"])
            elif inv == "func":
                rv = [p for p in m.get("params", []) if "retval" in (p.get("flags") or [])]
                self.methods[n] = (ps, _short_type(rv[0]["type"]) if rv else "")
        try:
            evs = vbtypes.events(type_name)
        except Exception:                       # noqa: BLE001 - no type library: no events
            evs = []
        for n, args in evs:
            self.events[n] = args

    def kind(self, name):
        if name in self.props:
            return "property"
        if name in self.methods:
            return "method"
        if name in self.events:
            return "event"
        return None


def args_text(params):
    out = []
    for p in params:
        out.append(("ByVal " if not p.get("byref", True) else "") + f"{p['name']} As {_short_type(p['type'])}")
    return ", ".join(out)


def arg_names(args):
    """the names of the arguments of an argument list (ByVal X As Long, Y -> X, Y)"""
    out = []
    for a in [x.strip() for x in (args or "").split(",") if x.strip()]:
        words = [w for w in a.split() if w.lower() not in ("byval", "byref", "optional", "paramarray")]
        if words:
            out.append(words[0].rstrip("()"))
    return ", ".join(out)


# ==================================================================== the model
class AxMember:
    """a public member of the control"""

    def __init__(self, name, kind, custom=False):
        self.name, self.kind, self.custom = name, kind, custom
        self.map_ctl, self.map_member = "", ""
        self.type = ""                          # property: its type; method: its return type ("" a Sub)
        self.default = "0"
        self.runtime, self.design = 0, 0
        self.args = ""
        self.description = ""

    @property
    def mapped(self):
        return bool(self.map_ctl and self.map_member)

    def info(self):
        """the 'MemberInfo= / 'MappingInfo= text of the member"""
        if self.mapped:
            return f"MappingInfo={self.map_ctl},{self.map_ctl},-1,{self.map_member}"
        if self.kind == "property":
            d = self.default if self.default not in ("", None) else "0"
            d = d.strip('"') if d.startswith('"') and d.endswith('"') else d
            return f"MemberInfo={type_code(self.type or 'Variant')},{self.runtime},{self.design},{d}"
        if self.kind == "method":
            return f"MemberInfo={type_code(self.type) if self.type else 5}"
        return ""

    def signature(self):
        return (self.kind, self.info(), self.args.replace(" ", "").lower(), self.type.lower(), self.description,
                self.default if self.kind == "property" and not self.mapped else "")


class AxModel:
    """the settings of the wizard for one user control"""

    def __init__(self, uc_name, controls, values=None):
        self.uc = uc_name
        self.controls = list(controls)          # [(name, type)] on the user control
        self.values = values or {}             # (control, property) -> its value in the designer
        self.info = {"UserControl": TypeInfo("UserControl")}
        for n, t in self.controls:
            self.info[n] = TypeInfo(t)
        self.available = {}                    # name -> kind
        for src in ["UserControl"] + [n for n, _ in self.controls]:
            ti = self.info[src]
            for d, k in ((ti.props, "property"), (ti.methods, "method"), (ti.events, "event")):
                for n in d:
                    self.available.setdefault(n, k)
        self.selected = []                     # the standard members chosen
        self.custom = []                       # the custom members (names)
        self.members = {}                      # name -> AxMember
        self.old = {}                          # name -> AxMember read from the code (the last run)

    # ------------------------------------------------------------ the lists of the pages
    def available_names(self):
        return sorted((n for n in self.available if n not in self.selected), key=str.lower)

    def all_names(self):
        return sorted(self.selected + self.custom, key=str.lower)

    def kind_of(self, name):
        m = self.members.get(name)
        if m is not None:
            return m.kind
        return self.available.get(name)

    def member(self, name):
        """the member of that name (made with the defaults of its kind the first time)"""
        m = self.members.get(name)
        if m is None:
            kind = self.kind_of(name) or "property"
            m = AxMember(name, kind, custom=name in self.custom)
            self._standard_defaults(m)
            self.members[name] = m
        return m

    def _standard_defaults(self, m):
        ti = self.info["UserControl"]
        if m.kind == "property":
            d = ti.props.get(m.name) or next((self.info[c].props[m.name] for c, _ in self.controls
                                              if m.name in self.info[c].props), None)
            m.type = (d or {}).get("type") or "Variant"
            m.default = "0"
        elif m.kind == "method":
            d = ti.methods.get(m.name) or next((self.info[c].methods[m.name] for c, _ in self.controls
                                                if m.name in self.info[c].methods), None)
            if d:
                m.args, m.type = args_text(d[0]), d[1]
        elif m.kind == "event":
            m.args = ti.events.get(m.name) or next((self.info[c].events[m.name] for c, _ in self.controls
                                                    if m.name in self.info[c].events), "")
        if not m.custom:
            m.description = DESCRIPTIONS.get(m.name, "")

    def select(self, name):
        if name not in self.selected and name in self.available:
            self.selected.append(name)

    def unselect(self, name):
        if name in self.selected:
            self.selected.remove(name)

    def add_custom(self, name, kind):
        m = AxMember(name, kind, custom=True)
        m.type = "Variant" if kind in ("property", "method") else ""   # VB6 DEV-9/ax1: Function ... As Variant
        self.members[name] = m
        if name not in self.custom:
            self.custom.append(name)
        return m

    def remove_custom(self, name):
        if name in self.custom:
            self.custom.remove(name)
        self.members.pop(name, None)

    def map_controls(self):
        return ["(None)"] + sorted([n for n, _ in self.controls], key=str.lower) + ["UserControl"]

    def map_members(self, ctl, kind):
        """the members of a control (or the UserControl) of that kind, for Maps to > Member"""
        ti = self.info.get(ctl)
        if ti is None:
            return []
        d = {"property": ti.props, "method": ti.methods, "event": ti.events}.get(kind, {})
        return sorted(d, key=str.lower)

    def set_mapping(self, name, ctl, member):
        m = self.member(name)
        if ctl in ("", "(None)") or member in ("", "(None)"):
            m.map_ctl = m.map_member = ""
        else:
            m.map_ctl, m.map_member = ctl, member
            src = self.info.get(ctl)
            if src is not None:                 # a mapped member takes the type of what it maps to
                if m.kind == "property" and member in src.props:
                    t = src.props[member]["type"] or "Variant"
                    m.type = "OLE_COLOR" if member.lower() in COLOR_PROPS else t
                elif m.kind == "method" and member in src.methods:
                    m.args, m.type = args_text(src.methods[member][0]), src.methods[member][1]
                elif m.kind == "event" and member in src.events:
                    m.args = src.events[member]

    def unmapped(self):
        return [n for n in self.all_names() if not self.member(n).mapped]


def designer_values(module):
    """(control, property) -> its value on the user control in the designer (and the defaults of the properties
    it does not write): the defaults of ReadProperty of the mapped properties (VB6 DEV-9/ax2: "Text1",
    &H8000000F)"""
    from . import vbprops as V
    out = {}
    root = module.sf.root if module is not None and module.sf is not None else None
    if root is None:
        return out
    for c in root.walk():
        owner = "UserControl" if c is root else c.name
        tname = "UserControl" if c is root else c.type_name
        for p in V.props_of(tname):
            if p.default is not None:
                out[(owner, p.name.lower())] = p.default
        if tname in ("TextBox", "Label", "CommandButton", "CheckBox", "OptionButton", "Frame"):
            out[(owner, "text" if tname == "TextBox" else "caption")] = c.name
        for it in c.items:
            if hasattr(it, "value") and hasattr(it, "name"):
                out[(owner, it.name.lower())] = it.value
    return out


# ==================================================================== reading the code of the last run
_INFO_RE = re.compile(r"^\s*'\s*(MemberInfo|MappingInfo)\s*=\s*(.*?)\s*$", re.IGNORECASE)
_PROC_RE = re.compile(r"^\s*(Public|Private|Friend)?\s*(Sub|Function|Property\s+(Get|Let|Set))\s+([A-Za-z_]\w*)\s*"
                      r"\((.*)\)\s*(?:As\s+([\w.]+))?", re.IGNORECASE)
_EVENT_RE = re.compile(r"^\s*(?:Public\s+)?Event\s+([A-Za-z_]\w*)\s*(?:\((.*?)\))?\s*('.*)?$", re.IGNORECASE)
_END_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
_CONST_RE = re.compile(r"^\s*Const\s+m_def_([A-Za-z_]\w*)\s*=\s*(.*?)\s*$", re.IGNORECASE)
_DIM_RE = re.compile(r"^\s*(?:Dim|Private)\s+m_([A-Za-z_]\w*)\s+As\s+([\w.]+)", re.IGNORECASE)
_ATTR_RE = re.compile(r'^\s*Attribute\s+([A-Za-z_]\w*)\.VB_Description\s*=\s*"(.*)"\s*$', re.IGNORECASE)


def read_code(model, text):
    """the members the wizard made in the code (its last run): model.old, model.selected, model.custom,
    model.members"""
    lines = text.replace("\r\n", "\n").split("\n")
    consts, descs = {}, {}
    for ln in lines:
        m = _CONST_RE.match(ln)
        if m:
            consts[m.group(1).lower()] = m.group(2)
        m = _ATTR_RE.match(ln)
        if m:
            descs[m.group(1).lower()] = m.group(2).replace('""', '"')
    found = {}
    pending = None
    for ln in lines:
        m = _INFO_RE.match(ln)
        if m:
            pending = (m.group(1).lower(), m.group(2))
            continue
        e = _EVENT_RE.match(ln)
        if e and not ln.lstrip().startswith("'"):
            am = AxMember(e.group(1), "event")
            am.args = (e.group(2) or "").strip()
            tail = e.group(3) or ""
            mi = re.search(r"MappingInfo\s*=\s*([^,]+),[^,]*,[^,]*,(\w+)", tail, re.IGNORECASE)
            if mi:
                am.map_ctl, am.map_member = mi.group(1).strip(), mi.group(2).strip()
            found[am.name] = am
            pending = None
            continue
        p = _PROC_RE.match(ln)
        if p and pending is not None:
            name, kind = p.group(4), p.group(2).lower()
            if name not in found:
                am = AxMember(name, "property" if kind.startswith("property") else "method")
                if pending[0] == "mappinginfo":
                    parts = [x.strip() for x in pending[1].split(",")]
                    if len(parts) >= 4:
                        am.map_ctl, am.map_member = parts[0], parts[3]
                else:
                    parts = [x.strip() for x in pending[1].split(",")]
                    if am.kind == "property" and len(parts) >= 3:
                        try:
                            am.runtime, am.design = int(parts[1]), int(parts[2])
                        except ValueError:
                            pass
                        am.default = consts.get(name.lower(), parts[3] if len(parts) > 3 else "0")
                if am.kind == "property":
                    am.type = p.group(6) or "Variant"
                else:
                    am.type = p.group(6) or ""
                    if kind == "function" and not am.type:
                        am.type = "Variant"
                    am.args = (p.group(5) or "").strip()
                found[name] = am
            pending = None
            continue
        if ln.strip() and not ln.lstrip().startswith("'"):
            pending = None if not p else pending
    for n, am in found.items():
        am.description = descs.get(n.lower(), "")
        am.custom = n not in model.available
        model.old[n] = am
        model.members[n] = _copy(am)
        if am.custom:
            if n not in model.custom:
                model.custom.append(n)
        elif n not in model.selected:
            model.selected.append(n)
    return found


def _copy(am):
    c = AxMember(am.name, am.kind, am.custom)
    for k in ("map_ctl", "map_member", "type", "default", "runtime", "design", "args", "description"):
        setattr(c, k, getattr(am, k))
    return c


# ==================================================================== the code
def _q(s):
    return '"' + str(s).replace('"', '""') + '"'


def const_value(m):
    d = (m.default or "").strip()
    t = (m.type or "Variant").lower()
    if t == "string":
        if d.startswith('"'):
            return d
        return _q(d) if d not in ("", "0") else '""'
    if t == "boolean":
        return {"true": "True", "-1": "True", "false": "False"}.get(d.lower(), d or "0")
    return d or "0"


def value_text(v):
    """a value of the designer as code (the default of ReadProperty of a mapped property)"""
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, int):
        if v < 0 or v > 0xFFFF:
            return "&H%X" % (v & 0xFFFFFFFF)
        return str(v)
    if isinstance(v, float):
        return repr(v)
    if v is None:
        return "0"
    return _q(v)


def target(m):
    return "UserControl" if m.map_ctl == "UserControl" else m.map_ctl


def member_blocks(model, m):
    """the code of a member: {"decl": [...], "procs": [...], "init": [...], "read": [...], "write": [...]}"""
    out = {"const": [], "dim": [], "event": [], "procs": [], "init": [], "read": [], "write": []}
    n = m.name
    desc = [f"Attribute {n}.VB_Description = {_q(m.description)}"] if m.description else []
    if m.kind == "event":
        tail = f" '{m.info()}" if m.mapped else ""
        out["event"] = [f"Event {n}({m.args}){tail}"] + desc
        if m.mapped:
            src = target(m)
            out["procs"] = [f"Private Sub {src}_{m.map_member}({m.args})",
                            f"    RaiseEvent {n}" + (f"({arg_names(m.args)})" if arg_names(m.args) else ""),
                            "End Sub", ""]
        return out
    head = [WIZ_COMMENT, "'" + m.info()]
    if m.kind == "method":
        kind = "Function" if m.type else "Sub"
        line = f"Public {kind} {n}({m.args})" + (f" As {m.type}" if m.type else "")
        body = []
        if m.mapped:
            tgt, an = f"{target(m)}.{m.map_member}", arg_names(m.args)
            body = [f"    {n} = {tgt}" + (f"({an})" if an else "")] if kind == "Function" else \
                [f"    {tgt}" + (f" {an}" if an else "")]
        out["procs"] = head + [line] + desc + body + [f"End {kind}", ""]
        return out
    # a property
    t = m.type or "Variant"
    obj = is_object(t)
    if m.mapped:
        tgt = f"{target(m)}.{m.map_member}"
        src = model.info.get(m.map_ctl)
        pd = (src.props.get(m.map_member) if src else None) or {"get": True, "let": True, "set": obj}
        can_set = pd.get("let") or pd.get("set")
        get = [f"Public Property Get {n}() As {t}"] + desc + \
            [f"    {'Set ' if obj else ''}{n} = {tgt}", "End Property", ""]
        procs = head + get
        if can_set:
            procs += [f"Public Property {'Set' if obj else 'Let'} {n}(ByVal New_{n} As {t})",
                      f"    {'Set ' if obj else ''}{tgt} = New_{n}", f'    PropertyChanged "{n}"', "End Property", ""]
            if obj:
                dflt = "Ambient.Font" if t.lower() == "font" else "Nothing"
            else:
                dflt = value_text(model.values.get((m.map_ctl, m.map_member.lower())))
            out["read"] = [f'    {"Set " if obj else ""}{tgt} = PropBag.ReadProperty("{n}", {dflt})']
            out["write"] = [f'    Call PropBag.WriteProperty("{n}", {tgt}, {dflt})']
        out["procs"] = procs
        return out
    var = f"m_{n}"
    if not obj:
        out["const"] = [f"Const m_def_{n} = {const_value(m)}"]
    out["dim"] = [f"Dim {var} As {t}"]
    get_guard = ["    If Ambient.UserMode Then Err.Raise 393"] if m.runtime in (2, 3) else []
    let_guard = []
    if m.runtime in (1, 3):
        let_guard.append("    If Ambient.UserMode Then Err.Raise 382")
    if m.design in (1, 2):
        let_guard.append("    If Ambient.UserMode = False Then Err.Raise 387")
    flags = [f'Attribute {n}.VB_MemberFlags = "400"'] if m.design == 2 else []
    memid = [f"Attribute {n}.VB_UserMemId = -512"] if n.lower() == "font" and t.lower() == "font" else []
    out["procs"] = head + [f"Public Property Get {n}() As {t}"] + desc + memid + flags + get_guard + \
        [f"    {'Set ' if obj else ''}{n} = {var}", "End Property", "",
         f"Public Property {'Set' if obj else 'Let'} {n}(ByVal New_{n} As {t})"] + let_guard + \
        [f"    {'Set ' if obj else ''}{var} = New_{n}", f'    PropertyChanged "{n}"', "End Property", ""]
    if obj:
        dflt = "Ambient.Font" if t.lower() == "font" else "Nothing"
        out["init"] = [f"    Set {var} = {dflt}"]
        out["read"] = [f'    Set {var} = PropBag.ReadProperty("{n}", {dflt})']
        out["write"] = [f'    Call PropBag.WriteProperty("{n}", {var}, {dflt})']
    else:
        out["init"] = [f"    {var} = m_def_{n}"]
        out["read"] = [f'    {var} = PropBag.ReadProperty("{n}", m_def_{n})']
        out["write"] = [f'    Call PropBag.WriteProperty("{n}", {var}, m_def_{n})']
    return out


# ---------------------------------------------------------------- the old code of a member: its lines
def _member_lines(lines, name, old):
    """the line numbers of the code the wizard made for a member (its info line and procedures, its event line,
    the handler of a mapped event, its Const / Dim, its lines in Init / Read / WriteProperties)"""
    low = name.lower()
    hits = set()
    i = 0
    while i < len(lines):
        ln = lines[i]
        if ln.lstrip().startswith("'"):
            i += 1
            continue
        p = _PROC_RE.match(ln)
        if p and p.group(4).lower() == low and (p.group(1) or "Public").lower() != "private":
            j = i
            while j < len(lines) and not _END_RE.match(lines[j]):
                j += 1
            k = i - 1                           # its info line and our comment over it
            while k >= 0 and (_INFO_RE.match(lines[k]) or lines[k].startswith(WIZ_COMMENT) or
                              "WARNING! DO NOT REMOVE" in lines[k]):
                hits.add(k)
                k -= 1
            hits.update(range(i, j + 1))
            i = j + 1
            continue
        e = _EVENT_RE.match(ln)
        if e and e.group(1).lower() == low:
            hits.add(i)
        if old is not None and old.kind == "event" and old.mapped and p and \
                p.group(4).lower() == f"{old.map_ctl}_{old.map_member}".lower():
            j = i
            while j < len(lines) and not _END_RE.match(lines[j]):
                j += 1
            hits.update(range(i, j + 1))
            i = j + 1
            continue
        m = _CONST_RE.match(ln)
        if m and m.group(1).lower() == low:
            hits.add(i)
        m = _DIM_RE.match(ln)
        if m and m.group(1).lower() == low:
            hits.add(i)
        if re.match(rf'^\s*(Set\s+)?m_{re.escape(name)}\s*=', ln, re.IGNORECASE) or \
                re.search(rf'Property\s*\(\s*"{re.escape(name)}"', ln, re.IGNORECASE):
            hits.add(i)
        m = re.match(r'^\s*Attribute\s+([A-Za-z_]\w*)\.', ln)
        if m and m.group(1).lower() == low and (i - 1) in hits:
            hits.add(i)
        i += 1
    return hits


def _proc_span(lines, name):
    for i, ln in enumerate(lines):
        p = _PROC_RE.match(ln)
        if p and p.group(4).lower() == name.lower() and not ln.lstrip().startswith("'"):
            for j in range(i + 1, len(lines)):
                if _END_RE.match(lines[j]):
                    return i, j
    return None


def _first_proc(lines):
    for i, ln in enumerate(lines):
        if _PROC_RE.match(ln) and not ln.lstrip().startswith("'"):
            return i
    return len(lines)


INIT_HEAD = ["' the properties of a new control start with their defaults", "Private Sub UserControl_InitProperties()"]
READ_HEAD = ["' the properties saved with the form that holds the control", "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)"]
WRITE_HEAD = ["' the properties to save with the form that holds the control", "Private Sub UserControl_WriteProperties(PropBag As PropertyBag)"]


def generate(model, old_text):
    """the new code of the user control: what changed since the last run commented out, the new code added"""
    lines = old_text.replace("\r\n", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    names = model.all_names()
    keep, write = set(), []
    for n in names:
        m = model.member(n)
        o = model.old.get(n)
        if o is not None and o.signature() == m.signature():
            keep.add(n)
        else:
            write.append(m)
    # the old code of what changed or went out: comments
    gone = [n for n in model.old if n not in keep]
    commented = set()
    for n in gone:
        commented |= _member_lines(lines, n, model.old.get(n))
    for i in sorted(commented):
        lines[i] = "'" + lines[i]
    if not write:
        return "\n".join(lines) + "\n"
    blocks = [member_blocks(model, m) for m in write]
    const = [x for b in blocks for x in b["const"]]
    dim = [x for b in blocks for x in b["dim"]]
    events = [x for b in blocks for x in b["event"]]
    procs = [x for b in blocks for x in b["procs"]]
    init = [x for b in blocks for x in b["init"]]
    read = [x for b in blocks for x in b["read"]]
    wr = [x for b in blocks for x in b["write"]]
    decl = []
    have = set(lines)
    for title, part in (("' the default values of the properties", const), ("' the values of the properties", dim),
                        ("' the events", events)):
        if part:
            decl += ([] if title in have else [title]) + part
    first = _first_proc(lines)
    k = 0                                       # after the last declaration before the first procedure
    for i in range(first):
        if lines[i].strip() and not lines[i].lstrip().startswith("'"):
            k = i + 1
    if k == 0:                                  # no declarations: before the comments over the first procedure
        k = first
        while k > 0 and (not lines[k - 1].strip() or lines[k - 1].startswith(WIZ_COMMENT) or
                         _INFO_RE.match(lines[k - 1])):
            k -= 1
    lines[k:k] = decl + ([""] if decl and k < len(lines) and lines[k].strip() else [])
    while lines and lines[-1] == "":
        lines.pop()
    if lines:
        lines.append("")
    lines += procs
    for head, add in ((INIT_HEAD, init), (READ_HEAD, read), (WRITE_HEAD, wr)):
        if not add:
            continue
        pname = head[1].split()[2].split("(")[0]
        span = _proc_span(lines, pname)
        if span:
            lines[span[1]:span[1]] = add
        else:
            while lines and lines[-1] == "":
                lines.pop()
            lines += [""] + head + add + ["End Sub"]
    while lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines) + "\n"


def summary(model):
    """the Summary Report after Finish (our own text)"""
    names = model.all_names()
    mapped = [n for n in names if model.member(n).mapped]
    out = [f"The ActiveX Control Interface Wizard has written the interface of {model.uc}.", "",
           "Members: " + (", ".join(names) or "(none)"),
           "Mapped to the controls on it: " + (", ".join(f"{n} -> {model.member(n).map_ctl}."
                                                         f"{model.member(n).map_member}" for n in mapped) or "(none)"),
           "", "What to do next:",
           "1) Close the designer of the user control and place the control on a form of a test project (File > Add "
           "Project > Standard EXE, or a form of this project).",
           "2) Look at its properties in the Properties window; save, close and open the form again to see that "
           "they are kept.",
           "3) Write the code of the methods the wizard left empty, and raise the custom events with RaiseEvent.",
           "4) Members that should follow more than one control (BackColor of the control and of its labels ...) "
           "need code of yours in their Property procedures.",
           "5) Resize the controls on it in UserControl_Resize.",
           "6) The Property Page Wizard makes property pages for the control.",
           "", "Run the wizard again to change the interface: what you change is written again, the old code of it "
           "stays as comments."]
    return "\n".join(out) + "\n"


__all__ = ["AxModel", "AxMember", "TypeInfo", "read_code", "generate", "summary", "DATA_TYPES", "RUN_TIME",
           "DESIGN_TIME", "DEFAULT_SELECTED", "TYPE_CODES"]
