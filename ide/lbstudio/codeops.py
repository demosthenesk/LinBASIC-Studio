"""Edits of the code window that VB6's Edit toolbar and Tools menu make (phase 8.3), on plain text so that they
are the same in the IDE and in the tests: Comment Block / Uncomment Block and Add Procedure.

What VB6 writes was taken from the VB6 IDE itself (VB-TESTS/DEV-8/ed1):
  Comment Block     a ' at the start of every line of the selection (or of the caret's line), before the indent
  Uncomment Block   one ' off every line: the one at the start, or the first one after the indent; the indent
                    stays; a line without a comment is left as it is
  Add Procedure     at the end of the module, after one empty line: "Public Sub DoIt()", "Private Function
                    Calc()", "Public Property Get Value() As Variant" + "Public Property Let Value(ByVal vNewValue
                    As Variant)", "Public Static Sub Cnt()" (All Local variables as Statics), each with an empty
                    line in its body; an Event goes to the declarations: an empty line and "Public Event Ev()"
                    after the last declaration; "Invalid procedure name", "A procedure of that name already
                    exists"
"""
import re

KINDS = ("Sub", "Function", "Property", "Event")

# words that can't name a procedure (VB6: Invalid procedure name)
RESERVED = {w.lower() for w in """
    And As Boolean ByRef Byte ByVal Call Case Close Const Currency Date Decimal Declare DefBool DefByte DefCur
    DefDate DefDbl DefDec DefInt DefLng DefObj DefSng DefStr DefVar Dim Do Double Each Else ElseIf Empty End
    Enum Eqv Erase Error Event Exit False For Friend Function Get Global GoSub GoTo If Imp Implements In
    Integer Is Let Like Long Loop LSet Me Mod New Next Not Nothing Null Object On Open Option Optional Or
    ParamArray Preserve Print Private Property Public RaiseEvent ReDim Rem Resume Return RSet Select Set Single
    Static Step Stop String Sub Then To True Type TypeOf Until Variant Wend While With WithEvents Xor
""".split()}

PROC_RE = re.compile(r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+"
                     r"([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
NAME_RE = re.compile(r"[A-Za-z][A-Za-z0-9_]*\Z")
EVENT_RE = re.compile(r"^\s*(?:(?:Public|Private)\s+)?Event\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
COMMENT_LINE_RE = re.compile(r"^\s*(?:'|Rem(?:\s|$))", re.IGNORECASE)


def procedures(text):
    """[(name, kind, line)] (as vbmodel.procedures, without the IDE)"""
    out = []
    for i, line in enumerate(text.split("\n")):
        m = PROC_RE.match(line)
        if m:
            out.append((m.group(2), " ".join(m.group(1).split()).title(), i))
    return out


# ---------------------------------------------------------------- Comment Block / Uncomment Block
def comment_line(line):
    return "'" + line


def uncomment_line(line):
    """one ' off: at the start, or the first one after the indent (VB6 keeps the indent)"""
    stripped = line.lstrip(" \t")
    if stripped.startswith("'"):
        k = len(line) - len(stripped)
        return line[:k] + stripped[1:]
    return line


def block_lines(first, last, sel_empty, last_col):
    """the lines a block command works on: the selection's lines; a selection that ends at the start of a
    line leaves that line out (as the Indent of VB6)"""
    if not sel_empty and last > first and last_col == 0:
        last -= 1
    return first, last


# ---------------------------------------------------------------- Add Procedure
def name_error(name, text):
    """None, or VB6's message for the name"""
    if not NAME_RE.match(name or "") or len(name) > 255 or name.lower() in RESERVED:
        return "Invalid procedure name"
    low = name.lower()
    if any(n.lower() == low for n, _, _ in procedures(text)):
        return "A procedure of that name already exists"
    for line in text.split("\n"):
        m = EVENT_RE.match(line)
        if m and m.group(1).lower() == low:
            return "A procedure of that name already exists"
    return None


def procedure_lines(name, kind, scope="Public", statics=False):
    """the lines Add Procedure writes (without the empty line before them)"""
    st = "Static " if statics and kind != "Event" else ""
    if kind == "Event":
        return [f"{scope} Event {name}()"]
    if kind == "Property":
        return [f"{scope} {st}Property Get {name}() As Variant", "", "End Property", "",
                f"{scope} {st}Property Let {name}(ByVal vNewValue As Variant)", "", "End Property"]
    end = "End Function" if kind == "Function" else "End Sub"
    return [f"{scope} {st}{kind} {name}()", "", end]


def add_procedure(text, name, kind="Sub", scope="Public", statics=False):
    """(new text, line of the caret: the body of the procedure, or the Event line) as VB6 adds it; text with
    \\n line ends. ValueError with VB6's message for a bad name"""
    err = name_error(name, text)
    if err:
        raise ValueError(err)
    if kind not in KINDS:
        raise ValueError(f"unknown kind {kind}")
    new = procedure_lines(name, kind, scope, statics)
    lines = text.split("\n")
    if kind == "Event":
        # after the last line of the declarations: the lines before the first procedure, without the empty and
        # comment lines just above it (they belong to the procedure, as the separator line of VB6 shows)
        procs = procedures(text)
        first = procs[0][2] if procs else len(lines)
        k = first - 1
        while k >= 0 and (not lines[k].strip() or COMMENT_LINE_RE.match(lines[k])):
            k -= 1
        if not procs:                           # only declarations: after the last of them
            k = len(lines) - 1
            while k >= 0 and not lines[k].strip():
                k -= 1
        out = lines[:k + 1] + [""] + new + lines[k + 1:]
        return "\n".join(out), k + 2
    # at the end of the module: the text ends with a line end, then an empty line, then the procedure
    if text and not text.endswith("\n"):
        text += "\n"
    body = text.count("\n") + 1 + 1             # the empty line, the header, then its body
    if not text:
        body = 2
    return text + "\n" + "\n".join(new) + "\n", body


# ---------------------------------------------------------------- Procedure Attributes
# What the Procedure Attributes dialog of VB6 writes (VB-TESTS/DEV-8/ed2): Attribute lines right under the header of
# the procedure (of a property: its first procedure), in this order:
#   Attribute Caption.VB_Description = "Title Text"
#   Attribute Caption.VB_HelpID = 5
#   Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"      (<property page>;<category>)
#   Attribute Caption.VB_UserMemId = -518                                 (the Procedure ID; (Default) = 0)
#   Attribute Caption.VB_MemberFlags = "200"                              (FUNCFLAGS in lower-case hex)
PROC_IDS_METHOD = [("(None)", None), ("(Default)", 0), ("AboutBox", -552), ("Refresh", -550)]   # (ed2 stages 5, 6)
PROC_IDS_PROPERTY = [("(None)", None), ("(Default)", 0), ("Appearance", -520), ("AutoSize", -500), ("BackColor", -501),
                     ("BackStyle", -502), ("BorderColor", -503), ("BorderStyle", -504), ("BorderWidth", -505),
                     ("Caption", -518), ("DrawMode", -507), ("DrawStyle", -508), ("DrawWidth", -509),
                     ("Enabled", -514), ("FillColor", -510), ("FillStyle", -511), ("Font", -512),
                     ("ForeColor", -513), ("hWnd", -515), ("TabStop", -516), ("Text", -517), ("RightToLeft", -611)]
CATEGORIES = ["(None)", "Appearance", "Behavior", "Data", "DDE", "Font", "List", "Misc", "Position", "Scale", "Text"]
# the check boxes: Hide this member, User Interface Default, Don't show in Property Browser, and Data Binding
FLAGS = {"hide": 0x40, "uidefault": 0x200, "nobrowse": 0x400, "bindable": 0x4, "defaultbind": 0x20,
         "displaybind": 0x10, "requestedit": 0x8, "immediatebind": 0x1000}
_ATTR_KEYS = ("VB_Description", "VB_HelpID", "VB_ProcData.VB_Invoke_Property", "VB_UserMemId", "VB_MemberFlags")
ATTR_RE = re.compile(r'^\s*Attribute\s+([A-Za-z_][A-Za-z0-9_]*)\.(VB_[A-Za-z_.]+)\s*=\s*(.*?)\s*$', re.IGNORECASE)


def _unquote(v):
    v = v.strip()
    if len(v) >= 2 and v[0] == '"' and v[-1] == '"':
        return v[1:-1].replace('""', '"')
    return v


def read_attributes(text, name):
    """{description, helpid, page, category, memid (None: none), flags} of a procedure (by name)"""
    a = {"description": "", "helpid": 0, "page": "", "category": "", "memid": None, "flags": 0}
    low = name.lower()
    for line in text.split("\n"):
        m = ATTR_RE.match(line)
        if not m or m.group(1).lower() != low:
            continue
        key, val = m.group(2).lower(), m.group(3)
        if key == "vb_description":
            a["description"] = _unquote(val)
        elif key == "vb_helpid":
            try:
                a["helpid"] = int(val)
            except ValueError:
                pass
        elif key == "vb_procdata.vb_invoke_property":
            a["page"], _, a["category"] = _unquote(val).partition(";")
        elif key == "vb_usermemid":
            try:
                a["memid"] = int(val)
            except ValueError:
                pass
        elif key == "vb_memberflags":
            try:
                a["flags"] = int(_unquote(val), 16)
            except ValueError:
                pass
    return a


def attribute_lines(name, a):
    out = []
    if a.get("description"):
        out.append(f'Attribute {name}.VB_Description = "' + a["description"].replace('"', '""') + '"')
    if a.get("helpid"):
        out.append(f"Attribute {name}.VB_HelpID = {int(a['helpid'])}")
    page, cat = a.get("page") or "", a.get("category") or ""
    if cat == "(None)":
        cat = ""
    if page == "(None)":
        page = ""
    if page or cat:
        out.append(f'Attribute {name}.VB_ProcData.VB_Invoke_Property = "{page};{cat}"')
    if a.get("memid") is not None:
        out.append(f"Attribute {name}.VB_UserMemId = {int(a['memid'])}")
    if a.get("flags"):
        out.append(f'Attribute {name}.VB_MemberFlags = "{int(a["flags"]):x}"')
    return out


def write_attributes(text, name, a):
    """the text with the attributes of the procedure as VB6 writes them: the old lines of these attributes go,
    the new ones come right under the header of the first procedure of that name"""
    low = name.lower()
    lines = text.split("\n")
    keep = []
    for line in lines:
        m = ATTR_RE.match(line)
        if m and m.group(1).lower() == low and m.group(2).lower() in {k.lower() for k in _ATTR_KEYS}:
            continue
        keep.append(line)
    head = next((i for n, _, i in procedures("\n".join(keep)) if n.lower() == low), None)
    if head is None:
        raise ValueError(f"no procedure {name}")
    real = next(n for n, _, i in procedures("\n".join(keep)) if i == head)
    return "\n".join(keep[:head + 1] + attribute_lines(real, a) + keep[head + 1:])
