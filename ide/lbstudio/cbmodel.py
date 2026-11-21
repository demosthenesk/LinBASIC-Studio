"""The object model of the Class Builder (phase 9.2c) and the code it writes.

What VB6's Class Builder Utility does (VB6 oracle DEV-9/cb1 ... cb6): a tree of the classes and collections of the
project; the classes it made carry their place in the tree in attributes of their header,

    Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
    Attribute VB_Ext_KEY = "Collection" ,"clsItem"           (a collection; "UserDefinedCollection": Set As Collection)
    Attribute VB_Ext_KEY = "Member0" ,"clsChild"              (a class in this one: a property of its type)
    Attribute VB_Ext_KEY = "Top_Level" ,"Yes"                 (in the root of the tree)

and Update Project writes the new classes and the new members of the old ones; what is already in the code of a
class is not changed (it cannot be deleted, renamed or changed from the Class Builder).  Based On copies the members
of a class into the new one; a class made while another one is selected becomes a member of it (an object made in
its Class_Initialize).  The code is our own (the comments of VB6's are Microsoft's text); the names that other code
uses (mvar<Property>, mCol, Add / Item / Count / Remove / NewEnum, modErrorHandling.RaiseError,
modClassIDGenerator.GetNextClassDebugID) are the ones of VB6's, so that projects of either open in the other.
"""
import re

try:
    from lbvc.lexer import KEYWORDS
except ImportError:                             # the IDE without the compiler on the path
    KEYWORDS = set()

PROJECT_ROOT = "Project1"
# the names the Class Builder refuses ("Invalid Name!": Text, Go, Width in DEV-9/cb1, cb3, cb6; Reset is taken):
# the keywords of the language and those words
RESERVED = {w.upper() for w in KEYWORDS} | {"TEXT", "GO", "WIDTH", "ME"}
NAME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,39}$")

# the Data Type lists (DEV-9/cb1: Property Builder and Add Argument sorted with the classes; Return Data Type in
# this order, then the classes)
BASE_TYPES = ["Boolean", "Byte", "Collection", "Currency", "Date", "Double", "ErrObject", "Integer", "Long",
              "Object", "Single", "StdFont", "StdPicture", "String", "Variant"]
RETURN_TYPES = ["(None)", "Byte", "Boolean", "Integer", "Long", "Single", "Double", "Currency", "Date", "String",
                "Variant", "Object", "Collection", "StdFont", "StdPicture", "ErrObject"]
OBJECT_TYPES = {"object", "collection", "stdfont", "stdpicture", "errobject", "iunknown"}

USER_COLLECTION = "UserDefinedCollection"


def valid_name(name):
    """a name the Class Builder takes: a VB identifier that is not a word of the language"""
    return bool(NAME_RE.match(name or "")) and name.upper() not in RESERVED


class Arg:
    """an argument of a method or an event (Add Argument)"""

    def __init__(self, name, type_="Variant", byval=False, array=False, optional=False, default=""):
        self.name, self.type, self.byval, self.array = name, type_, byval, array
        self.optional, self.default = optional, default

    def text(self):
        s = ("Optional " if self.optional else "") + ("ByVal " if self.byval else "") + self.name
        s += "()" if self.array else ""
        s += " As " + self.type
        if self.optional and self.default != "":
            s += " = " + str(self.default)
        return s

    def copy(self):
        return Arg(self.name, self.type, self.byval, self.array, self.optional, self.default)


class Member:
    """a property, method, event or enum of a class

    property: type, decl ("public": Property Let/Set and Get, "friend": the same as Friend, "var": a Public
    variable), default (the default member); method: type (the return type, "" for a Sub), args, friend,
    default; event: args; enum: values (["x=1", "y"]), public.  existing: it is in the code of the class already
    (VB6: it cannot be deleted or changed from the Class Builder)."""

    def __init__(self, kind, name, type_="", **kw):
        self.kind, self.name, self.type = kind, name, type_
        self.decl = kw.get("decl", "public")
        self.default = kw.get("default", False)
        self.args = kw.get("args", [])
        self.friend = kw.get("friend", False)
        self.values = kw.get("values", [])
        self.public = kw.get("public", True)
        self.description = kw.get("description", "")
        self.helpid = kw.get("helpid", 0)
        self.existing = kw.get("existing", False)
        self.accessors = kw.get("accessors", "")     # of a property in the code: "Get/Let", "Get/Set" ...
        self.arg_text = kw.get("arg_text", None)     # the arguments as the code has them

    def copy(self):
        m = Member(self.kind, self.name, self.type, decl=self.decl, default=self.default,
                   args=[a.copy() for a in self.args], friend=self.friend, values=list(self.values),
                   public=self.public, description=self.description, helpid=self.helpid)
        return m

    # ---------------------------------------------------------------- the columns of the lists
    def prop_declaration(self):
        if self.kind != "property":
            return ""
        if self.existing and self.accessors:
            return self.accessors
        if self.decl == "var":
            return "Var"
        t = (self.type or "Variant").lower()
        if t == "variant":
            return "Get/Let/Set"
        if is_object_type(self.type):
            return "Get/Set"
        return "Get/Let"

    def arguments(self):
        if self.arg_text is not None:
            return self.arg_text
        return ", ".join(a.text() for a in self.args)

    def data_type(self):
        return "" if self.kind in ("event", "enum") else (self.type or "")


class CBClass:
    """a class or a collection of the tree"""

    def __init__(self, name, collection=False, item=""):
        self.name = name
        self.collection = collection         # made by the Collection Builder (Add, Item, Count, Remove, NewEnum)
        self.item = item                     # the class of its items
        self.user_collection = False         # Set As Collection (VB_Ext_KEY "Collection" "UserDefinedCollection")
        self.description, self.helpid = "", 0
        self.members = []
        self.children = []                   # the names of the classes in it (VB_Ext_KEY MemberN)
        self.top_level = True
        self.existing = False                # a module of the project already
        self.foreign = False                 # a module the Class Builder did not make (no SavedWithClassBuilder6)
        self.module = None
        self.old_name = None                 # a renamed module: its name in the project
        self.old_children = []               # the children its code has already
        self.old_description, self.old_helpid = "", 0
        self.old_user_collection = False

    @property
    def is_collection(self):
        return self.collection or self.user_collection

    def member(self, name):
        low = name.lower()
        return next((m for m in self.members if m.name.lower() == low), None)


SCALAR_TYPES = {"boolean", "byte", "currency", "date", "decimal", "double", "integer", "long", "single", "string",
                "variant"}


def is_object_type(t):
    """Object, Collection, StdFont ... and the classes: Set and Property Set"""
    return bool(t) and t.lower() not in SCALAR_TYPES and t.lower() != "(none)"


# ==================================================================== reading the classes of the project
_EXT_RE = re.compile(r'^\s*"([^"]*)"\s*,\s*"([^"]*)"\s*$')


def ext_keys(attributes):
    """[(key, value)] of the VB_Ext_KEY attributes of a header"""
    out = []
    for n, v, _ in attributes:
        if n.strip().lower() == "vb_ext_key":
            m = _EXT_RE.match(v or "")
            if m:
                out.append((m.group(1), m.group(2)))
    return out


def _unquote(v):
    v = (v or "").strip()
    if len(v) >= 2 and v[0] == '"' and v[-1] == '"':
        return v[1:-1].replace('""', '"')
    return v


def _attr(attributes, name):
    for n, v, _ in attributes:
        if n.strip().lower() == name.lower():
            return v
    return None


def logical_lines(text):
    """the lines of the code with the continuations ( _) joined: [(first line number, text)]"""
    out, buf, start = [], "", None
    if isinstance(text, (list, tuple)):
        text = "\n".join(text)
    for i, line in enumerate(text.replace("\r\n", "\n").split("\n")):
        if start is None:
            start = i
        s = line.rstrip()
        if s.endswith(" _"):
            buf += s[:-1]
            continue
        out.append((start, buf + line))
        buf, start = "", None
    if buf:
        out.append((start, buf))
    return out


def _strip_comment(s):
    q = False
    for i, ch in enumerate(s):
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            return s[:i]
    return s


_PROC_RE = re.compile(r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+"
                      r"([A-Za-z_]\w*)\s*\((.*)\)\s*(?:As\s+(?:New\s+)?([\w.]+)(\(\))?)?\s*$", re.IGNORECASE)
_EVENT_RE = re.compile(r"^\s*(?:Public\s+)?Event\s+([A-Za-z_]\w*)\s*(?:\((.*)\))?\s*$", re.IGNORECASE)
_ENUM_RE = re.compile(r"^\s*(?:(Public|Private)\s+)?Enum\s+([A-Za-z_]\w*)\s*$", re.IGNORECASE)
_VAR_RE = re.compile(r"^\s*(Public|Global)\s+(?!Sub\b|Function\b|Property\b|Event\b|Enum\b|Const\b|Declare\b|"
                     r"Type\b|WithEvents\b)(.*)$", re.IGNORECASE)
_ATTR_RE = re.compile(r'^\s*Attribute\s+([A-Za-z_]\w*)\.(VB_\w+)\s*=\s*(.*?)\s*$', re.IGNORECASE)
_END_RE = re.compile(r"^\s*End\s+(Sub|Function|Property|Enum|Type)\b", re.IGNORECASE)


def _split_top(s):
    """a, b(1, 2), c -> [a, b(1, 2), c]"""
    out, depth, cur = [], 0, ""
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    if cur.strip():
        out.append(cur)
    return [x.strip() for x in out]


def parse_args(text):
    """the arguments of a procedure header -> [Arg]"""
    out = []
    for a in _split_top(text or ""):
        if not a:
            continue
        default = ""
        if "=" in a:
            a, default = a.split("=", 1)
            a, default = a.strip(), default.strip()
        words = a.split()
        opt = byval = False
        while words and words[0].lower() in ("optional", "byval", "byref", "paramarray"):
            w = words.pop(0).lower()
            opt = opt or w == "optional"
            byval = byval or w == "byval"
        if not words:
            continue
        name = words[0]
        arr = name.endswith("()")
        name = name[:-2] if arr else name
        typ = "Variant"
        if len(words) >= 3 and words[1].lower() == "as":
            typ = words[-1]
        out.append(Arg(name, typ, byval, arr, opt, default))
    return out


def parse_members(text):
    """the members of the code of a class as the Class Builder shows them: its Public / Friend properties (by
    name, with their accessors), methods, events, enums and Public variables; [Member] (existing)"""
    props, out = {}, []
    attrs = {}
    in_proc = False
    enum = None
    for _, raw in logical_lines(text):
        line = _strip_comment(raw).rstrip()
        m = _ATTR_RE.match(raw)
        if m:
            attrs.setdefault(m.group(1).lower(), {})[m.group(2).lower()] = m.group(3)
            continue
        if enum is not None:
            if re.match(r"^\s*End\s+Enum\b", line, re.IGNORECASE):
                out.append(enum)
                enum = None
            elif line.strip():
                enum.values.append(re.sub(r"\s*=\s*", "=", line.strip()))
            continue
        if in_proc:
            if _END_RE.match(line):
                in_proc = False
            continue
        if not line.strip():
            continue
        m = _PROC_RE.match(line)
        if m:
            in_proc = True
            scope = (m.group(1) or "Public").lower()
            if scope == "private":
                continue
            name, kind = m.group(4), m.group(2).lower()
            args = parse_args(m.group(5))
            ret = m.group(6) or ""
            if kind.startswith("property"):
                acc = m.group(3).capitalize()
                p = props.get(name.lower())
                if p is None:
                    p = Member("property", name, existing=True, decl="friend" if scope == "friend" else "public")
                    p._acc = set()
                    props[name.lower()] = p
                    out.append(p)
                p._acc.add(acc)
                if acc == "Get":
                    p.type = ret or "Variant"
                    p.arg_text = ", ".join(a.text() for a in args)
                elif not p.type and args:
                    p.type = args[-1].type
            else:
                out.append(Member("method", name, ret if kind == "function" else "", existing=True,
                                  args=args, friend=scope == "friend",
                                  arg_text=", ".join(a.text() for a in args)))
            continue
        m = _EVENT_RE.match(line)
        if m:
            args = parse_args(m.group(2))
            out.append(Member("event", m.group(1), existing=True, args=args,
                              arg_text=", ".join(a.text() for a in args)))
            continue
        m = _ENUM_RE.match(line)
        if m:
            enum = Member("enum", m.group(2), existing=True, public=(m.group(1) or "Public").lower() != "private")
            continue
        m = _VAR_RE.match(line)
        if m and not in_proc:
            for part in _split_top(m.group(2)):
                vm = re.match(r"^([A-Za-z_]\w*)(\(\))?\s*(?:As\s+(?:New\s+)?([\w.]+))?", part, re.IGNORECASE)
                if vm:
                    out.append(Member("property", vm.group(1), vm.group(3) or "Variant", decl="var", existing=True,
                                      accessors="Var"))
    for p in props.values():
        acc = p._acc
        p.accessors = "/".join(a for a in ("Get", "Let", "Set") if a in acc)
        if not p.type:
            p.type = "Variant"
        del p._acc
    for mem in out:
        a = attrs.get(mem.name.lower(), {})
        mem.description = _unquote(a.get("vb_description") or a.get("vb_vardescription") or "")
        try:
            mem.helpid = int(a.get("vb_helpid") or a.get("vb_varhelpid") or 0)
        except ValueError:
            mem.helpid = 0
        memid = a.get("vb_usermemid") or a.get("vb_varusermemid")
        mem.default = memid is not None and memid.strip() == "0"
    return out


def read_project(project):
    """the Class Builder's model of the class modules of a project -> Model"""
    model = Model()
    if project is None:
        return model
    model.project_name = project.name if hasattr(project, "name") else PROJECT_ROOT
    for m in project.modules:
        if m.kind != "Class" or m.sf is None:
            continue
        c = CBClass(m.name)
        c.existing, c.module = True, m
        keys = ext_keys(m.sf.attributes)
        kd = {k.lower(): v for k, v in keys}
        c.foreign = "savedwithclassbuilder6" not in kd
        c.top_level = kd.get("top_level", "Yes").lower() != "no"
        members = sorted(((int(k[6:]), v) for k, v in keys if re.match(r"^member\d+$", k, re.IGNORECASE)))
        c.children = [v for _, v in members]
        coll = kd.get("collection")
        if coll and coll.lower() == USER_COLLECTION.lower():
            c.user_collection = c.old_user_collection = True
        elif coll:
            c.collection, c.item = True, coll
            c.children = [v for v in c.children if v.lower() != coll.lower()]   # Member0: the item class
        c.old_children = list(c.children)
        c.description = c.old_description = _unquote(_attr(m.sf.attributes, "VB_Description"))
        try:
            c.helpid = c.old_helpid = int(_attr(m.sf.attributes, "VB_HelpID") or 0)
        except ValueError:
            c.helpid = 0
        c.members = parse_members(m.code_text())
        model.add(c)
    return model


# ==================================================================== the model
class Model:
    """the classes of the Class Builder, the options of its code, its changes since the last Update Project"""

    def __init__(self):
        self.classes = []
        self.project_name = PROJECT_ROOT
        self.debug_code = False             # View > Options: Include Debug code in Initialize and Terminate events
        self.err_raise = False              # and Include Err.Raise in all generated methods
        self.dirty = False

    def add(self, c):
        self.classes.append(c)

    def cls(self, name):
        low = (name or "").lower()
        return next((c for c in self.classes if c.name.lower() == low), None)

    def names(self):
        return [c.name for c in self.classes]

    def parents(self, name):
        low = name.lower()
        return [c for c in self.classes if any(x.lower() == low for x in c.children) or
                (c.collection and c.item.lower() == low)]

    def roots(self):
        """the classes in the root of the tree: Top_Level, or not in another one"""
        out = []
        for c in self.classes:
            if c.top_level or not self.parents(c.name):
                out.append(c)
        return out

    def tree(self):
        """[(class, ref, [subtree])]: ref is True for the item class under a collection (VB6 shows it with the
        icon of a property, without what is in it: DEV-9/cb1-50); a class already on the way from the root is not opened again (VB6 opens it for
        ever: DEV-9/cb2-10)"""
        def node(c, path, ref=False):
            kids = []
            if not ref and c.name.lower() not in path:
                p = path | {c.name.lower()}
                if c.collection and c.item and c.item != USER_COLLECTION:
                    ic = self.cls(c.item)
                    if ic is not None:
                        kids.append(node(ic, p, True))
                for n in sorted(c.children, key=str.lower):
                    k = self.cls(n)
                    if k is not None:
                        kids.append(node(k, p))
            return (c, ref, kids)
        # in the order of the names, as VB6 (DEV-9/cb1-50, cb3-02, cb8)
        return [node(c, frozenset()) for c in sorted(self.roots(), key=lambda x: x.name.lower())]

    def unique(self, base):
        names = {c.name.lower() for c in self.classes}
        i = 1
        while f"{base}{i}".lower() in names:
            i += 1
        return f"{base}{i}"

    def changed(self):
        return self.dirty

    def display_members(self, c):
        """the members of a class in the lists: its own, the properties of the classes in it, and those of a
        new collection (as Update Project will write them)"""
        if c.collection and not c.existing:
            return parse_members("\n".join(collection_code(self, c, Options())))
        out = list(c.members)
        have = {m.name.lower() for m in out}
        for n in c.children:
            if n.lower() not in have:
                out.append(Member("property", n, n, accessors="Get/Set", existing=c.existing and n in c.old_children))
        return out


class Options:
    def __init__(self, debug=False, errors=False):
        self.debug, self.errors = debug, errors


# ==================================================================== the code
CLASS_HEADER = """VERSION 1.0 CLASS
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


def _q(s):
    return '"' + str(s).replace('"', '""') + '"'


def ext_lines(model, c):
    """the VB_Ext_KEY lines of a class"""
    out = ['Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"']
    if c.collection and c.item:
        out.append(f'Attribute VB_Ext_KEY = "Collection" ,{_q(c.item)}')
    elif c.user_collection:
        out.append(f'Attribute VB_Ext_KEY = "Collection" ,"{USER_COLLECTION}"')
    i = 0
    if c.collection and c.item:
        out.append(f'Attribute VB_Ext_KEY = "Member0" ,{_q(c.item)}')
        i = 1
    for n in c.children:
        out.append(f'Attribute VB_Ext_KEY = "Member{i}" ,{_q(n)}')
        i += 1
    out.append('Attribute VB_Ext_KEY = "Top_Level" ,"%s"' % ("Yes" if c.top_level else "No"))
    return out


def header_attributes(model, c):
    out = []
    if c.helpid:
        out.append(f"Attribute VB_HelpID = {int(c.helpid)}")
    if c.description:
        out.append(f"Attribute VB_Description = {_q(c.description)}")
    return out + ext_lines(model, c)


def _err_wrap(cls_name, label, tag, body, kind, opts):
    """the body of a procedure, in On Error GoTo ... RaiseError with View > Options > Include Err.Raise"""
    if not opts.errors:
        return body
    lab = label + "Err"
    ex = {"Sub": "Exit Sub", "Function": "Exit Function", "Property": "Exit Property"}[kind]
    return ([f"    On Error GoTo {lab}"] + body + [f"    {ex}", f"{lab}:",
            f'    RaiseError MyUnhandledError, "{cls_name}:{tag}"'])


def member_attributes(name, m, var=False):
    pre = "VB_Var" if var else "VB_"
    out = []
    if m.description:
        out.append(f"Attribute {name}.{pre}Description = {_q(m.description)}")
    if m.helpid:
        out.append(f"Attribute {name}.{pre}HelpID = {int(m.helpid)}")
    if m.default:
        out.append(f"Attribute {name}.{'VB_VarUserMemId' if var else 'VB_UserMemId'} = 0")
    return out


def declaration_lines(m):
    """the lines of a member in the declarations of the class"""
    if m.kind == "property":
        if m.decl == "var":
            return [f"Public {m.name} As {m.type or 'Variant'}"] + member_attributes(m.name, m, var=True)
        return [f"Private mvar{m.name} As {m.type or 'Variant'}"]
    if m.kind == "event":
        args = ", ".join(a.text() for a in m.args)
        return [f"' {m.name}: RaiseEvent {m.name}" + (f"({', '.join(a.name for a in m.args)})" if m.args else "") +
                " raises it", f"Public Event {m.name}" + (f"({args})" if m.args else "()")] + \
            member_attributes(m.name, m)
    if m.kind == "enum":
        return [f"{'Public' if m.public else 'Private'} Enum {m.name}"] + \
            ["    " + re.sub(r"\s*=\s*", " = ", v) for v in m.values] + ["End Enum"]
    return []


def procedure_lines(cls_name, m, opts):
    """the procedures of a member (a property: Let / Set and Get; a method: its Sub or Function)"""
    out = []
    if m.kind == "property" and m.decl != "var":
        scope = "Friend" if m.decl == "friend" else "Public"
        t = m.type or "Variant"
        n, var = m.name, "mvar" + m.name
        first = True

        def attrs():
            nonlocal first
            if first:
                first = False
                return member_attributes(n, m)
            return []
        if t.lower() == "variant":
            out += [f"{scope} Property Let {n}(ByVal vData As Variant)"] + attrs() + \
                _err_wrap(cls_name, n + "Let", n + " Property Let", [f"    {var} = vData"], "Property", opts) + \
                ["End Property", ""]
            out += [f"{scope} Property Set {n}(ByVal vData As Variant)"] + \
                _err_wrap(cls_name, n + "Set", n + " Property Set", [f"    Set {var} = vData"], "Property", opts) + \
                ["End Property", ""]
            out += [f"{scope} Property Get {n}() As Variant"] + \
                _err_wrap(cls_name, n + "Get", n + " Property Get",
                          [f"    If IsObject({var}) Then", f"        Set {n} = {var}", "    Else",
                           f"        {n} = {var}", "    End If"], "Property", opts) + ["End Property", ""]
        elif is_object_type(t):
            out += [f"{scope} Property Set {n}(ByVal vData As {t})"] + attrs() + \
                _err_wrap(cls_name, n + "Set", n + " Property Set", [f"    Set {var} = vData"], "Property", opts) + \
                ["End Property", ""]
            out += [f"{scope} Property Get {n}() As {t}"] + \
                _err_wrap(cls_name, n + "Get", n + " Property Get", [f"    Set {n} = {var}"], "Property", opts) + \
                ["End Property", ""]
        else:
            out += [f"{scope} Property Let {n}(ByVal vData As {t})"] + attrs() + \
                _err_wrap(cls_name, n + "Let", n + " Property Let", [f"    {var} = vData"], "Property", opts) + \
                ["End Property", ""]
            out += [f"{scope} Property Get {n}() As {t}"] + \
                _err_wrap(cls_name, n + "Get", n + " Property Get", [f"    {n} = {var}"], "Property", opts) + \
                ["End Property", ""]
    elif m.kind == "method":
        scope = "Friend" if m.friend else "Public"
        args = ", ".join(a.text() for a in m.args)
        kind = "Function" if m.type and m.type != "(None)" else "Sub"
        head = f"{scope} {kind} {m.name}({args})" + (f" As {m.type}" if kind == "Function" else "")
        body = ["    ' the code of " + m.name] if opts.errors else []
        out += [head] + member_attributes(m.name, m) + \
            _err_wrap(cls_name, m.name, m.name + " Method", body, kind, opts) + [f"End {kind}", ""]
    return out


def child_lines(cls_name, child, opts):
    """a class in this one: its variable, Property Get / Set (and its New / Nothing in Class_Initialize / Terminate)"""
    var = "mvar" + child
    decl = [f"Private {var} As {child}"]
    procs = [f"Public Property Get {child}() As {child}"] + \
        _err_wrap(cls_name, child + "Get", child + " Property Get", [f"    Set {child} = {var}"], "Property", opts) + \
        ["End Property", "", f"Public Property Set {child}(vData As {child})"] + \
        _err_wrap(cls_name, child + "Set", child + " Property Set", [f"    Set {var} = vData"], "Property", opts) + \
        ["End Property", ""]
    return decl, procs, [f"    Set {var} = New {child}"], [f"    Set {var} = Nothing"]


DEBUG_DECL = ["' View > Options > Include Debug code: 0 here leaves it out of this class", "#Const DebugMode = 1",
              "#If DebugMode Then", "    Private mlClassDebugID As Long", "#End If"]
DEBUG_INIT = ["    #If DebugMode Then", "        mlClassDebugID = GetNextClassDebugID()",
              "        Debug.Print \"'\" & TypeName(Me) & \"' instance \" & mlClassDebugID & \" created\"",
              "    #End If"]
DEBUG_TERM = ["    #If DebugMode Then",
              "        Debug.Print \"'\" & TypeName(Me) & \"' instance \" & CStr(mlClassDebugID) & \" is terminating\"",
              "    #End If"]
DEBUG_PROP = ["#If DebugMode Then", "    ' the number of this instance (the order in which the instances were made)",
              "    Public Property Get ClassDebugID()", "        ClassDebugID = mlClassDebugID",
              "    End Property", "#End If"]


def class_code(model, c, opts):
    """the code of a new class (without its header)"""
    decl, procs, init, term = [], [], [], []
    debug = opts.debug
    if debug:
        decl += DEBUG_DECL
        init += DEBUG_INIT
        term += DEBUG_TERM
    vals = [m for m in c.members if m.kind == "property"]
    if any(m.decl != "var" for m in vals):
        decl.append("' the values of the properties")
    for m in c.members:
        decl += declaration_lines(m)
    for n in c.children:
        d, p, i, t = child_lines(c.name, n, opts)
        decl += d
        procs += p
        init += i
        term = t + term
    for m in reversed(c.members):            # the newest first, as VB6 (DEV-9/cb1)
        procs += procedure_lines(c.name, m, opts)
    out = decl + [""] + procs
    if init:
        out += ["Private Sub Class_Initialize()"] + init + ["End Sub", ""]
    if term:
        out += ["Private Sub Class_Terminate()"] + term + ["End Sub", ""]
    if debug:
        out += DEBUG_PROP
    while out and out[-1] == "":
        out.pop()
    return out


def collection_code(model, c, opts):
    """the code of a new collection of c.item: Add (with the properties of the item class), Item (the default
    member), Count, Remove, NewEnum (For Each)"""
    item = c.item or "Variant"
    ic = model.cls(item)
    params, sets = [], []
    if ic is not None:
        for m in model.display_members(ic):
            if m.kind != "property" or (m.accessors and "Get" == m.accessors):
                continue
            t = m.type or "Variant"
            params.append(f"{m.name} As {t}")
            if t.lower() == "variant":
                sets += [f"    If IsObject({m.name}) Then", f"        Set objNewMember.{m.name} = {m.name}", "    Else",
                         f"        objNewMember.{m.name} = {m.name}", "    End If"]
            elif is_object_type(t):
                sets.append(f"    Set objNewMember.{m.name} = {m.name}")
            else:
                sets.append(f"    objNewMember.{m.name} = {m.name}")
    params.append("Optional sKey As String")
    out = ["' the items of this collection", "Private mCol As Collection", "",
           f"Public Function Add({', '.join(params)}) As {item}",
           "    ' a new item with the values given, under the key sKey if there is one",
           "    Dim objNewMember As " + item, "    Set objNewMember = New " + item] + sets + \
        ["    If Len(sKey) = 0 Then", "        mCol.Add objNewMember", "    Else", "        mCol.Add objNewMember, sKey",
         "    End If", "    Set Add = objNewMember", "    Set objNewMember = Nothing", "End Function", "",
         f"Public Property Get Item(vntIndexKey As Variant) As {item}",
         "Attribute Item.VB_UserMemId = 0",
         "    ' an item by its number (from 1) or its key; the default member: x(1), x(\"key\")",
         "    Set Item = mCol(vntIndexKey)", "End Property", "",
         "Public Property Get Count() As Long", "    Count = mCol.Count", "End Property", "",
         "Public Sub Remove(vntIndexKey As Variant)", "    ' takes out the item of that number or key",
         "    mCol.Remove vntIndexKey", "End Sub", "",
         "Public Property Get NewEnum() As IUnknown",
         "Attribute NewEnum.VB_UserMemId = -4", 'Attribute NewEnum.VB_MemberFlags = "40"',
         "    ' For Each item In <this collection>", "    Set NewEnum = mCol.[_NewEnum]", "End Property", "",
         "Private Sub Class_Initialize()", "    Set mCol = New Collection", "End Sub", "",
         "Private Sub Class_Terminate()", "    Set mCol = Nothing", "End Sub"]
    if c.members:                            # members the user gave the collection too
        extra_d, extra_p = [], []
        for m in c.members:
            extra_d += declaration_lines(m)
        for m in reversed(c.members):
            extra_p += procedure_lines(c.name, m, opts)
        out = extra_d + out[:3] + extra_p + out[3:]
    for n in c.children:
        d, p, i, t = child_lines(c.name, n, opts)
        k = out.index("Private Sub Class_Initialize()")
        out[k + 2:k + 2] = i
        k = out.index("Private Sub Class_Terminate()")
        out[k + 1:k + 1] = t
        out = d + out
        k = out.index("Private Sub Class_Initialize()")
        out[k:k] = p
    return out


def module_text(model, c, opts, exposed=False):
    """the whole text of a new class module"""
    head = CLASS_HEADER.replace("{name}", c.name).replace("{exposed}", "True" if exposed else "False")
    code = collection_code(model, c, opts) if c.collection else class_code(model, c, opts)
    return head + "\n".join(header_attributes(model, c)) + "\n" + "\n".join(code) + "\n"


# ---------------------------------------------------------------- the new members of a class that exists
def _first_proc_line(lines):
    in_enum = in_type = False
    for i, l in enumerate(lines):
        s = _strip_comment(l).strip()
        if re.match(r"^(Public\s+|Private\s+)?Enum\b", s, re.IGNORECASE):
            in_enum = True
        elif re.match(r"^(Public\s+|Private\s+)?Type\b", s, re.IGNORECASE):
            in_type = True
        elif re.match(r"^End\s+(Enum|Type)\b", s, re.IGNORECASE):
            in_enum = in_type = False
        elif not in_enum and not in_type and _PROC_RE.match(s):
            return i
    return len(lines)


def _proc_span(lines, name):
    """(header line, End line) of a procedure of that name"""
    for i, l in enumerate(lines):
        m = _PROC_RE.match(_strip_comment(l).rstrip())
        if m and m.group(4).lower() == name.lower():
            for j in range(i + 1, len(lines)):
                if _END_RE.match(lines[j]):
                    return i, j
    return None


def add_to_code(model, c, code, opts):
    """the code of a class that exists, with its new members (VB6: added after its declarations; its old code is
    not changed) -> new code text"""
    lines = code.replace("\r\n", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    new = [m for m in c.members if not m.existing]
    kids = [n for n in c.children if n not in c.old_children]
    if not new and not kids:
        return code
    decl, procs, init, term = [], [], [], []
    if any(m.kind == "property" and m.decl != "var" for m in new):
        decl.append("' the values of the properties")
    for m in new:
        decl += declaration_lines(m)
    for n in kids:
        d, p, i, t = child_lines(c.name, n, opts)
        decl += d
        procs += p
        init += i
        term = t + term
    for m in reversed(new):
        procs += procedure_lines(c.name, m, opts)
    k = _first_proc_line(lines)
    end = k
    while end > 0 and not lines[end - 1].strip():
        end -= 1
    body = lines[:end] + decl + ([""] if procs else []) + procs
    rest = lines[k:]
    if init:
        span = _proc_span(rest, "Class_Initialize")
        if span:
            rest[span[1]:span[1]] = init
        else:
            body += ["Private Sub Class_Initialize()"] + init + ["End Sub", ""]
    if term:
        span = _proc_span(rest, "Class_Terminate")
        if span:
            rest[span[0] + 1:span[0] + 1] = term
        else:
            body += ["Private Sub Class_Terminate()"] + term + ["End Sub", ""]
    if rest and body and body[-1] != "":
        body.append("")
    return "\n".join(body + rest)


# ---------------------------------------------------------------- Update Project
def _new_module(project, kind, name, ext, text):
    from .vbmodel import Module, parse_source
    sf = parse_source(text.replace("\n", "\r\n"), None, kind, True, project.encoding)
    sf.encoding = project.encoding
    m = Module(project, kind, name, name + ext, sf)
    m.dirty = True
    project.modules.append(m)
    project._vbp_add(m)
    return m


def update_project(model, project, rename=None, set_code=None):
    """Update Project: the new classes as modules, the new members (and classes put in them) of the classes that
    exist added to their code, the renamed ones renamed (rename(module, name)), the VB_Ext_KEY attributes of all;
    the modules of View > Options when new code uses them. set_code(module, text): the new code of a module (an
    open code window); without it, the module's text"""
    opts = Options(model.debug_code, model.err_raise)
    exposed = getattr(project, "ptype", "Exe") in ("OleExe", "OleDll")
    added = []
    for c in list(model.classes):
        if c.existing:
            _update_existing(model, c, opts, rename, set_code)
            continue
        m = _new_module(project, "Class", c.name, ".cls", module_text(model, c, opts, exposed))
        added.append(m)
        c.existing, c.module = True, m
        if c.collection:
            c.members = parse_members(m.code_text())
        for x in c.members:
            x.existing = True
        c.old_children = list(c.children)
        c.old_description, c.old_helpid = c.description, c.helpid
        c.old_user_collection = c.user_collection
        c.foreign = False
    if added:
        for name, text in (("modErrorHandling", ERROR_MODULE if opts.errors else None),
                           ("modClassIDGenerator", DEBUG_MODULE if opts.debug else None)):
            if text and project.module(name) is None:
                _new_module(project, "Module", name, ".bas", text)
    project.dirty = True
    model.dirty = False
    return added


def _update_existing(model, c, opts, rename, set_code):
    m = c.module
    if m is None:
        return
    if c.old_name is not None and m.name != c.name:
        if rename is not None:
            rename(m, c.name)
        else:
            m.name = c.name
        c.old_name = None
    text = m.code_text()
    new = add_to_code(model, c, text, opts)
    if new != text:
        if set_code is not None:
            set_code(m, new)
        else:
            m.editor_text = None
            m.sf.code = new.split("\n")
        m.dirty = True
        for x in c.members:
            x.existing = True
        c.old_children = list(c.children)
    # the place of the class in the tree, its Description and Help Context ID (in its header)
    managed = ("vb_ext_key", "vb_description", "vb_helpid")
    keep = [a for a in m.sf.attributes if a[0].strip().lower() not in managed]
    old = sorted(a[2].strip() for a in m.sf.attributes if a[0].strip().lower() in managed)
    lines = header_attributes(model, c)
    if sorted(lines) != old:
        for line in lines:
            n, _, v = line[len("Attribute "):].partition("=")
            keep.append((n.strip(), v.strip(), line))
        m.sf.attributes = keep
        m.dirty = True
    c.foreign = False


# ---------------------------------------------------------------- the modules of View > Options
ERROR_MODULE = """Attribute VB_Name = "modErrorHandling"
' The errors the classes of the Class Builder raise (View > Options > Include Err.Raise).
' Give your own errors numbers above 512 here; the text of an error is the string of its number in the
' resource file of the project, if it has one.
Public Const MyObjectError1 = 1000
Public Const MyObjectError2 = 1010
Public Const MyObjectErrorN = 1234
Public Const MyUnhandledError = 9999

Private Function ErrorText(ByVal ErrorNum As Long) As String
    On Error GoTo NoText
    ErrorText = LoadResString(ErrorNum)
    Exit Function
NoText:
    ErrorText = "Unexpected error " & ErrorNum
End Function

Public Sub RaiseError(ByVal ErrorNumber As Long, ByVal Source As String)
    ' the error goes to the code that called the class
    Err.Raise vbObjectError + ErrorNumber, Source, ErrorText(ErrorNumber)
End Sub
"""

DEBUG_MODULE = """Attribute VB_Name = "modClassIDGenerator"
' The numbers of the instances of the classes of the Class Builder (View > Options > Include Debug code).
Public Function GetNextClassDebugID() As Long
    Static lastID As Long
    lastID = lastID + 1
    GetNextClassDebugID = lastID
End Function
"""


__all__ = ["Arg", "Member", "CBClass", "Model", "Options", "read_project", "parse_members", "module_text",
           "add_to_code", "update_project", "header_attributes", "ext_keys", "valid_name", "BASE_TYPES", "RETURN_TYPES",
           "ERROR_MODULE", "DEBUG_MODULE", "USER_COLLECTION", "is_object_type"]
