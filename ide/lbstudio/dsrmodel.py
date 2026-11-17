"""The designers of VB6 kept in .dsr files: the DataEnvironment and the DataReport, as objects the IDE
edits, and back as the text VB6 writes (the VB98 sample DataRept, the VB6 oracles tests/db/d_de1,
d_rpt1). What the IDE does not know of a .dsr stays as it was."""
import os
import uuid

from lbvc.project.source import Property, PropertyBlock, parse_source

DE_GUID = "{C0E45035-5775-11D0-B388-00A0C9055D8E}"
CLSID_STDPICTURE = bytes.fromhex("0452e30b918fce119de300aa004bb851")   # {0BE35204-8F91-11CE-9DE3-00AA004BB851}
REPORT_GUID = "{78E93846-85FD-11D0-8487-00A0C90DC8A9}"
SECTION_GUID = "{1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9}"
FONT_GUID = "{0BE35203-8F91-11CE-9DE3-00AA004BB851}"
FORMAT_GUID = "{6D835690-900B-11D0-9484-00A0C91110ED}"
# the controls of a report: ItemType, the GUID of their block, the Toolbox name
ITEMS = {"RptLabel": (3, "{1C13A8E1-A0B6-11D0-848E-00A0C90DC8A9}"),
         "RptTextBox": (4, "{1C13A8E2-A0B6-11D0-848E-00A0C90DC8A9}"),
         "RptLine": (5, "{1C13A8E3-A0B6-11D0-848E-00A0C90DC8A9}"),
         "RptShape": (6, "{1C13A8E4-A0B6-11D0-848E-00A0C90DC8A9}"),
         "RptImage": (7, "{1C13A8E5-A0B6-11D0-848E-00A0C90DC8A9}"),
         "RptFunction": (13, "{49FF6930-2B8C-11D1-8DA9-00A0C90FFFC2}")}
ITEM_OF_TYPE = {v[0]: k for k, v in ITEMS.items()}
# SectionCode: what a section is
SEC_NAMES = {1: "Report Header", 2: "Page Header", 3: "Group Header", 4: "Detail", 5: "Group Footer",
             7: "Page Footer", 8: "Report Footer"}
AGG_FUNCS = ["(none)", "Any", "Average", "Count", "Maximum", "Minimum", "Standard Deviation", "Sum"]

REFERENCES = {
    "DataEnvironment": [
        "*\\G{3D5C6BF0-69A3-11D0-B393-00A0C9055D8E}#1.0#0#msderun.dll#Microsoft Data Environment Instance 6.0",
        "*\\G{00000200-0000-0010-8000-00AA006D2EA4}#2.0#0#msado20.tlb#Microsoft ActiveX Data Objects 2.0 Library"],
    "DataReport": [
        "*\\G{642AC760-AAB4-11D0-8494-00A0C90DC8A9}#1.0#0#MSDBRPTR.DLL#Microsoft Data Report Designer v6.0",
        "*\\G{6B263850-900B-11D0-9484-00A0C91110ED}#1.0#0#MSSTDFMT.DLL#Microsoft Data Formatting Object Library"],
}


def kind_of(sf):
    g = (sf.root.type_name if sf is not None and sf.root is not None else "").upper()
    return "DataEnvironment" if g == DE_GUID else "DataReport" if g == REPORT_GUID else None


def _props(block):
    return [it for it in block.items if isinstance(it, Property)]


def _get(block, name, default=None):
    for it in block.items:
        if isinstance(it, Property) and it.name.lower() == name.lower():
            return it.value
    return default


def _blocks(block, prefix):
    return [it for it in block.items if isinstance(it, PropertyBlock) and it.name.lower().startswith(prefix.lower())
            and it.name[len(prefix):].isdigit()]


def _fmt(v):
    if isinstance(v, bool):
        return "-1  'True" if v else "0   'False"
    if isinstance(v, float):
        s = repr(v)
        return s[:-2] if s.endswith(".0") else s
    if isinstance(v, int):
        return str(v)
    return '"' + str(v).replace('"', '""') + '"'


def _line(depth, name, value, comment=None):
    s = f"{'   ' * depth}{name:<16}=   {_fmt(value)}"
    return s + "  '" + comment if comment and not isinstance(value, bool) else s


def font_lines(depth, font):
    pad = "   " * depth
    return [f"{pad}BeginProperty Font {FONT_GUID} ",
            _line(depth + 1, "Name", font.get("name", "Arial")),
            _line(depth + 1, "Size", font.get("size", 8.25)),
            _line(depth + 1, "Charset", font.get("charset", 0)),
            _line(depth + 1, "Weight", font.get("weight", 400)),
            _line(depth + 1, "Underline", bool(font.get("underline", False))),
            _line(depth + 1, "Italic", bool(font.get("italic", False))),
            _line(depth + 1, "Strikethrough", bool(font.get("strikethrough", False))),
            f"{pad}EndProperty"]


def read_font(block):
    f = {"name": "Arial", "size": 8.25, "charset": 0, "weight": 400, "underline": False, "italic": False,
         "strikethrough": False}
    if block is None:
        return f
    for p in _props(block):
        k = p.name.lower()
        v = p.value
        if k in ("underline", "italic", "strikethrough"):
            v = v not in (0, False, "0")
        f[k] = v
    return f


def _replace_designer(sf, text):
    """the Begin ... End of sf as the text says; its code and attributes stay"""
    new = parse_source(text.replace("\n", "\r\n"), None, "Designer", True, sf.encoding)
    sf.root = new.root
    sf.version = new.version or sf.version


# ================================================================ DataEnvironment
class DeConnection:
    def __init__(self, name="Connection1", source="", props=None):
        self.name = name
        self.source = source
        # [(name, value, comment)] kept as they were; a new connection: what VB6's designer writes
        self.extra = props if props is not None else [("Expanded", True, None), ("QuoteChar", 0, None),
                                                      ("SeparatorChar", 0, None)]

    def to_lines(self, i, dispid):
        L = [f"   BeginProperty Connection{i} ", _line(2, "ConnectionName", self.name), _line(2, "ConnDispId", dispid),
             _line(2, "SourceOfData", 3), "   " * 2 + f"{'ConnectionSource':<16}=   {_fmt(self.source)}"]
        for x in self.extra:
            if x[0].lower() not in ("connectionname", "conndispid", "sourceofdata", "connectionsource"):
                L.append(_line(2, *x))
        L.append("   EndProperty")
        return L


# DataType (ADO) and HostType (VB) of a parameter: the names the designer shows
PARAM_TYPES = [(2, "adSmallInt"), (3, "adInteger"), (4, "adSingle"), (5, "adDouble"), (6, "adCurrency"),
               (7, "adDate"), (11, "adBoolean"), (12, "adVariant"), (17, "adUnsignedTinyInt"), (129, "adChar"),
               (130, "adWChar"), (131, "adNumeric"), (200, "adVarChar"), (202, "adVarWChar")]
HOST_TYPES = [(2, "Integer (VT_I2)"), (3, "Long (VT_I4)"), (4, "Single (VT_R4)"), (5, "Double (VT_R8)"),
              (6, "Currency (VT_CY)"), (7, "Date (VT_DATE)"), (8, "String (VT_BSTR)"), (11, "Boolean (VT_BOOL)"),
              (12, "Variant (VT_VARIANT)"), (17, "Byte (VT_UI1)")]
HOST_OF_TYPE = {2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 11: 11, 12: 12, 17: 17, 129: 8, 130: 8, 131: 5, 200: 8, 202: 8}


class Param:
    """a parameter of a command (BeginProperty P<n> of the .dsr)"""
    ORDER = ("RealName", "UserName", "Direction", "Precision", "Scale", "Size", "DataType", "HostType", "Required")

    def __init__(self, name="Param1", datatype=202, size=255):
        self.v = {"RealName": name, "UserName": name, "Direction": 1, "Precision": 0, "Scale": 0, "Size": size,
                  "DataType": datatype, "HostType": HOST_OF_TYPE.get(datatype, 12), "Required": True}
        self.extra = []

    @classmethod
    def read(cls, b):
        p = cls()
        known = {k.lower(): k for k in cls.ORDER}
        for x in _props(b):
            k = known.get(x.name.lower())
            if k:
                p.v[k] = x.value
            else:
                p.extra.append((x.name, x.value, x.comment))
        if isinstance(p.v["Required"], int) and not isinstance(p.v["Required"], bool):
            p.v["Required"] = p.v["Required"] != 0
        return p

    @property
    def name(self):
        return str(self.v["UserName"] or self.v["RealName"])

    def to_lines(self, depth):
        return [_line(depth, k, self.v[k]) for k in self.ORDER] + [_line(depth, *x) for x in self.extra]


class DeCommand:
    def __init__(self, name="Command1"):
        self.name = name
        self.conn = ""
        self.ctype = 1                          # adCmdText 1, adCmdTable 2, adCmdStoredProc 4
        self.text = ""
        self.parent = ""                        # RelateToParent: the parent command
        self.relations = []                     # [(parent field, child field)]
        self.params = []                        # [Param]: the parameters (? of the SQL)
        self.grouping = False
        self.groups = []                        # the fields of the grouping
        self.grouping_name = ""
        self.aggregates = []                    # [(name, aggon, field, function index)]
        self.fields = []                        # [(name, type, size, precision)] as VB6 lists them
        self.extra = []                         # other properties: (name, value)
        self.present = None                     # the properties it was read with (None: a new command)
        self.agg_extra = {}                     # aggregate name -> {property: value} VB6 wrote beyond ours

    @property
    def member(self):
        return (self.grouping_name or self.name + "_Grouping") if self.grouping else self.name

    def to_lines(self, i, dispids, model=None):
        top = not self.parent
        L = [f"   BeginProperty Recordset{i} ", _line(2, "CommandName", self.name),
             _line(2, "CommDispId", dispids[0] if top else -1), _line(2, "RsDispId", dispids[1] if top else -1),
             _line(2, "CommandText", self.text), _line(2, "ActiveConnectionName", self.conn),
             _line(2, "CommandType", self.ctype)]
        if self.ctype == 2:
            L.append(_line(2, "dbObjectType", 1))
        if self.parent:
            L += [_line(2, "RelateToParent", True), _line(2, "ParentCommandName", self.parent)]
        elif self.present is None or "parentcommandname" in self.present:
            L.append(_line(2, "ParentCommandName", ""))      # VB6's designer writes it (oracle d_de2)
        if self.grouping:
            L += [_line(2, "Grouping", True), _line(2, "GroupingName", self.member)]
        L.append(_line(2, "Expanded", True))
        if self.grouping:
            L += [_line(2, "SummaryExpanded", True), _line(2, "DetailExpanded", True)]
        L.append(_line(2, "IsRSReturning", True))
        for x in self.extra:
            L.append(_line(2, *x))
        L.append(_line(2, "NumFields", len(self.fields)))
        for k, (fn, ft, fs, fp) in enumerate(self.fields, 1):
            L += [f"      BeginProperty Field{k} ", _line(3, "Precision", fp), _line(3, "Size", fs), _line(3, "Scale", 255),
                  _line(3, "Type", ft), _line(3, "Name", fn), _line(3, "Caption", fn), "      EndProperty"]
        if self.grouping:
            L.append(_line(2, "GroupingName", self.member))
        L.append(_line(2, "NumGroups", len(self.groups) if self.grouping else 0))
        if self.grouping:
            for k, g in enumerate(self.groups, 1):
                ft = next(((t, s, p) for n, t, s, p in self.fields if n.lower() == g.lower()), (202, 255, 255))
                L += [f"      BeginProperty Grouping{k} ", _line(3, "Precision", ft[2]), _line(3, "Size", ft[1]),
                      _line(3, "Scale", 255), _line(3, "Type", ft[0]), _line(3, "Name", g), _line(3, "Caption", g),
                      "      EndProperty"]
        L.append(_line(2, "ParamCount", len(self.params)))
        for k, prm in enumerate(self.params, 1):
            L += [f"      BeginProperty P{k} "] + prm.to_lines(3) + ["      EndProperty"]
        if self.present is not None and "commandparent" in self.present:
            L.append(_line(2, "CommandParent", self.parent))
        L.append(_line(2, "RelationCount", len(self.relations) if self.parent else 0))
        if self.parent:
            for k, (p, c) in enumerate(self.relations, 1):
                L += [f"      BeginProperty Relation{k} ", _line(3, "ParentField", p), _line(3, "ChildField", c),
                      _line(3, "ParentType", 0),
                      _line(3, "ChildType", 1 if any(c.lower() == x.name.lower() for x in self.params) else 0),
                      "      EndProperty"]
        L.append(_line(2, "AggregateCount", len(self.aggregates)))
        for k, (n, on, f, fn) in enumerate(self.aggregates, 1):
            # the type of the aggregate as MSDataShape gives it: COUNT a Long, SUM / AVG of a Currency
            # a Currency, MIN / MAX the field's type, the rest a Double
            src = model.command(on) if model is not None else None
            ft = next((t for nm, t, _, _ in (src.fields if src else []) if nm.lower() == f.lower()), 5)
            at = 3 if fn == 3 else (6 if ft == 6 else 5) if fn in (2, 7) else ft if fn in (4, 5) else 5
            ex = self.agg_extra.get(n.lower(), {})
            L += [f"      BeginProperty Aggregate{k} ", _line(3, "Name", n), _line(3, "AggOn", on), _line(3, "AggField", f),
                  _line(3, "AggType", 1), _line(3, "AggFunction", fn)]
            L += [_line(3, x, ex[x.lower()]) for x in ("Precision", "Size", "Scale") if x.lower() in ex]
            L += [_line(3, "Type", at), _line(3, "Name", n)]
            if "caption" in ex:
                L.append(_line(3, "Caption", ex["caption"]))
            L += [_line(3, "Control", "TextBox"),
                  _line(3, "ControlGuid", "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"), "      EndProperty"]
        L.append("   EndProperty")
        return L


class DataEnvModel:
    def __init__(self, sf):
        self.sf = sf
        root = sf.root
        self.name = sf.name or root.name
        self.typelib = _get(root, "TypeLibGuid") or "{%s}" % str(uuid.uuid4()).upper()
        self.typeinfo = _get(root, "TypeInfoGuid") or "{%s}" % str(uuid.uuid4()).upper()
        self.conns = []
        for b in _blocks(root, "Connection"):
            self.conns.append(DeConnection(str(_get(b, "ConnectionName", "")), str(_get(b, "ConnectionSource", "") or ""),
                                           [(p.name, p.value, p.comment) for p in _props(b)]))
        self.cmds = []
        known = {"commandname", "commdispid", "rsdispid", "commandtext", "activeconnectionname", "commandtype",
                 "dbobjecttype", "relatetoparent", "parentcommandname", "grouping", "groupingname", "expanded",
                 "isrsreturning", "numfields", "numgroups", "paramcount", "commandparent", "relationcount",
                 "aggregatecount", "summaryexpanded", "detailexpanded"}
        for b in _blocks(root, "Recordset"):
            c = DeCommand(str(_get(b, "CommandName", "")))
            c.conn = str(_get(b, "ActiveConnectionName", "") or "")
            c.ctype = int(_get(b, "CommandType", 1) or 1)
            c.text = str(_get(b, "CommandText", "") or "")
            if _get(b, "RelateToParent", 0) not in (0, False, None):
                c.parent = str(_get(b, "ParentCommandName", "") or "")
            c.grouping = _get(b, "Grouping", 0) not in (0, False, None)
            c.grouping_name = str(_get(b, "GroupingName", "") or "")
            c.groups = [str(_get(g, "Name", "")) for g in _blocks(b, "Grouping")]
            c.relations = [(str(_get(r, "ParentField", "")), str(_get(r, "ChildField", ""))) for r in _blocks(b, "Relation")]
            c.params = [Param.read(x) for x in _blocks(b, "P")]
            c.aggregates = [(str(_get(a, "Name", "")), str(_get(a, "AggOn", "")), str(_get(a, "AggField", "")),
                             int(_get(a, "AggFunction", 0) or 0)) for a in _blocks(b, "Aggregate")]
            c.fields = [(str(_get(f, "Name", "")), int(_get(f, "Type", 202) or 202), int(_get(f, "Size", 255) or 0),
                         int(_get(f, "Precision", 255) or 255)) for f in _blocks(b, "Field")]
            c.extra = [(p.name, p.value, p.comment) for p in _props(b) if p.name.lower() not in known]
            c.present = {p.name.lower() for p in _props(b)}
            for a in _blocks(b, "Aggregate"):
                c.agg_extra[str(_get(a, "Name", "")).lower()] = {p.name.lower(): p.value for p in _props(a)
                                                                 if p.name.lower() in ("precision", "size", "scale", "caption")}
            self.cmds.append(c)

    def children(self, cmd):
        return [c for c in self.cmds if c.parent and c.parent.lower() == cmd.name.lower()]

    def top(self):
        return [c for c in self.cmds if not c.parent]

    def command(self, name):
        return next((c for c in self.cmds if c.name.lower() == name.lower()), None)

    def members(self):
        """the DataMember names a bound control or a report can use"""
        return [c.member for c in self.top()]

    def unique(self, base, names):
        low = {n.lower() for n in names}
        i = 1
        while f"{base}{i}".lower() in low:
            i += 1
        return f"{base}{i}"

    def text(self):
        L = ["VERSION 5.00", f"Begin {DE_GUID} {self.name} ",
             _line(1, "ClientHeight", 5070), _line(1, "ClientLeft", 0), _line(1, "ClientTop", 0), _line(1, "ClientWidth", 5160),
             _line(1, "_ExtentX", 9102), _line(1, "_ExtentY", 8943), _line(1, "FolderFlags", 7),
             _line(1, "TypeLibGuid", self.typelib), _line(1, "TypeInfoGuid", self.typeinfo), _line(1, "TypeInfoCookie", 0),
             _line(1, "Version", 4), _line(1, "NumConnections", len(self.conns))]
        did = 1001
        for i, c in enumerate(self.conns, 1):
            L += c.to_lines(i, did)
            did += 1
        L.append(_line(1, "NumRecordsets", len(self.cmds)))
        for i, c in enumerate(self.cmds, 1):
            ids = (did, did + 1) if not c.parent else (-1, -1)
            if not c.parent:
                did += 2
            L += c.to_lines(i, ids, self)
        L.append("End")
        return "\n".join(L)

    def apply(self):
        _replace_designer(self.sf, self.text())


def new_dataenv_text(name):
    m = DataEnvModel.__new__(DataEnvModel)
    m.sf = None
    m.name = name
    m.typelib = "{%s}" % str(uuid.uuid4()).upper()
    m.typeinfo = "{%s}" % str(uuid.uuid4()).upper()
    m.conns = [DeConnection("Connection1", "")]
    m.cmds = []
    return (m.text() + "\nAttribute VB_Name = \"{name}\"\nAttribute VB_GlobalNameSpace = False\n"
            "Attribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n").replace("{name}", name)


# ================================================================ DataReport
class RptItem:
    def __init__(self, kind, name):
        self.kind = kind                        # RptLabel ...
        self.name = name
        self.left = self.top = 0
        self.width = 1440
        self.height = 240
        self.font = read_font(None)
        self.props = {}                         # caption, datafield, datamember, alignment, functiontype ...
        self.fmt = None                         # DataFormat: {type, format, lcid}
        self.pic = None                         # RptImage: the file of its Picture (kept in the .dsx)
        self.pic_off = 0                        # its place in the .dsx being written
        self.dsx = ""                           # the name of the .dsx ("rptX.dsx")

    def to_lines(self, i):
        t, guid = ITEMS[self.kind]
        L = [_line(2, f"ItemType{i}", t), f"      BeginProperty Item{i} {guid} ", _line(3, "_Version", 393216),
             _line(3, "Name", self.name)]
        if self.left:
            L.append(f"         {'Object.Left':<23}=   {_fmt(int(self.left))}")
        if self.top:
            L.append(f"         {'Object.Top':<23}=   {_fmt(int(self.top))}")
        L.append(f"         {'Object.Width':<23}=   {_fmt(int(self.width))}")
        if self.kind != "RptLine" or self.height:
            L.append(f"         {'Object.Height':<23}=   {_fmt(int(self.height))}")
        if self.pic is not None:
            L.append(f"         {'Picture':<16}=   \"{self.dsx}\":{self.pic_off:04X}")
        if self.kind in ("RptLabel", "RptTextBox", "RptFunction"):
            L += font_lines(3, self.font)
        for k, v in self.props.items():
            if k == "Caption":
                L.append(f"         {'Object.Caption':<23}=   {_fmt(v)}")
            elif k not in ("DataMember",):
                L.append(_line(3, k, v))
        if self.kind in ("RptTextBox", "RptFunction"):
            f = self.fmt or {"type": 0, "format": ""}
            # the properties of the DataFormat as the file had them; a new one: as VB6 SP6 writes it
            props = f.get("props") or [("Type", 0), ("Format", ""), ("HaveTrueFalseNull", 0), ("FirstDayOfWeek", 0),
                                       ("FirstWeekOfYear", 0), ("LCID", 0), ("SubFormatType", 0)]
            L.append(f"         BeginProperty DataFormat {FORMAT_GUID} ")
            for n, v in props:
                v = f.get("type", 0) if n.lower() == "type" else f.get("format", "") if n.lower() == "format" else v
                L.append(_line(4, n, v))
            L.append("         EndProperty")
            L.append(_line(3, "DataMember", self.props.get("DataMember", "")))
        L.append("      EndProperty")
        return L


class RptSection:
    def __init__(self, code, name, height=360):
        self.code = code
        self.name = name
        self.height = height
        self.props = {}                         # KeepTogether, ForcePageBreak, Visible
        self.items = []

    def to_lines(self, i):
        L = [f"   {'SectionCode' + str(i):<16}=   {self.code}", f"   BeginProperty Section{i} {SECTION_GUID} ",
             _line(2, "_Version", 393216), _line(2, "Name", self.name)]
        if self.height:
            L.append(f"      {'Object.Height':<23}=   {int(self.height)}")
        for k, v in self.props.items():
            L.append(_line(2, k, v))
        L.append(_line(2, "NumControls", len(self.items)))
        for j, it in enumerate(self.items):
            L += it.to_lines(j)
        L.append("   EndProperty")
        return L


class ReportModel:
    def __init__(self, sf, dsx_name="", path=None, dsx_data=None):
        self.sf = sf
        self.path = path                        # the .dsr: its .dsx has the pictures
        self._dsx = dsx_data                    # or the .dsx not saved yet
        self.dsx_file = (os.path.splitext(os.path.basename(path))[0] if path else (sf.name or sf.root.name)) + ".dsx"
        b = _get(sf.root, "Bindings")
        if b is not None and getattr(b, "file", None):
            self.dsx_file = b.file                  # the name the .dsr has for it ("rptNwind.dsx")
        root = sf.root
        self.name = sf.name or root.name
        self.caption = str(_get(root, "Caption", self.name) or "")
        self.client = [int(_get(root, n, d) or d) for n, d in (("ClientHeight", 8235), ("ClientLeft", 1650),
                                                                 ("ClientTop", 1545), ("ClientWidth", 6585))]
        self.dver = int(_get(root, "_DesignerVersion", 100684049) or 100684049)
        self.extent = [int(_get(root, "_ExtentX", 11615) or 11615), int(_get(root, "_ExtentY", 14526) or 14526)]
        self.present = {p.name.lower() for p in _props(root)}
        self.width = int(_get(root, "ReportWidth", 9360) or 9360)
        self.gridx = int(_get(root, "GridX", 10) or 10)
        self.gridy = int(_get(root, "GridY", 10) or 10)
        self.margins = [int(_get(root, n, 1440) if _get(root, n) is not None else 1440)
                        for n in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin")]
        self.datamember = str(_get(root, "DataMember", "") or "")
        self.title = str(_get(root, "Title", "") or "")
        self.datasource = dsx_name
        fb = next((b for b in root.items if isinstance(b, PropertyBlock) and b.name.lower() == "font"), None)
        self.font = read_font(fb)
        codes = {}
        for p in _props(root):
            if p.name.lower().startswith("sectioncode"):
                codes[p.name[11:]] = int(p.value)
        self.sections = []
        for b in _blocks(root, "Section"):
            s = RptSection(codes.get(b.name[7:], 4), str(_get(b, "Name", "")), 0)
            for p in _props(b):
                n = p.name.lower()
                if n == "object.height":
                    s.height = int(p.value)
                elif n in ("keeptogether", "forcepagebreak", "visible"):
                    s.props[p.name] = p.value
            types = {p.name[8:]: int(p.value) for p in _props(b) if p.name.lower().startswith("itemtype")}
            for ib in _blocks(b, "Item"):
                it = RptItem(ITEM_OF_TYPE.get(types.get(ib.name[4:], 3), "RptLabel"), str(_get(ib, "Name", "")))
                for p in _props(ib):
                    n = p.name.lower()
                    if n in ("_version", "name"):
                        continue
                    if n.startswith("object."):
                        k = n[7:]
                        if k in ("left", "top", "width", "height"):
                            setattr(it, k, int(p.value))
                        elif k == "caption":
                            it.props["Caption"] = p.value
                        else:
                            it.props[p.name[7:]] = p.value
                    elif n == "picture":
                        it.pic = self._picture(p.value)
                    else:
                        it.props[p.name] = p.value
                if it.kind == "RptLine" and _get(ib, "Object.Height") is None:
                    it.height = 0
                fb = next((x for x in ib.items if isinstance(x, PropertyBlock) and x.name.lower() == "font"), None)
                it.font = read_font(fb) if fb is not None else dict(self.font)
                db = next((x for x in ib.items if isinstance(x, PropertyBlock) and x.name.lower() == "dataformat"), None)
                if db is not None:
                    it.fmt = {"type": int(_get(db, "Type", 0) or 0), "format": str(_get(db, "Format", "") or ""),
                              "props": [(x.name, x.value) for x in _props(db)]}
                s.items.append(it)
            self.sections.append(s)

    def _picture(self, ref):
        """the image file of a Picture kept in the .dsx (as in a .frx, with the CLSID of StdPicture)"""
        import struct
        from lbvc.project.source import FrxRef
        from lbvc.project.fsutil import ci_path
        if not isinstance(ref, FrxRef):
            return None
        data = self._dsx
        if data is None:
            if not self.path:
                return None
            try:
                with open(ci_path(os.path.join(os.path.dirname(self.path), ref.file)), "rb") as f:
                    data = self._dsx = f.read()
            except OSError:
                return None
        p = ref.offset + 4
        if data[p:p + 16] == CLSID_STDPICTURE:
            p += 16
        if data[p:p + 4] != b"lt\0\0" or p + 8 > len(data):
            return None
        n = struct.unpack_from("<I", data, p + 4)[0]
        return bytes(data[p + 8:p + 8 + n]) or None

    def all_names(self):
        return [s.name for s in self.sections] + [i.name for s in self.sections for i in s.items]

    def unique(self, base):
        low = {n.lower() for n in self.all_names()}
        i = 1
        while f"{base}{i}".lower() in low:
            i += 1
        return f"{base}{i}"

    def items(self):
        return [i for s in self.sections for i in s.items]

    def needs_dsx(self):
        return bool(self.datasource) or any(i.pic is not None for i in self.items())

    def dsx_bytes(self):
        """the .dsx: the DataSource, then the pictures (VB6 oracle d_de2); the offsets go to the items"""
        import struct
        out = bytearray(dsx_bytes(self.datasource))
        for it in self.items():
            it.dsx = self.dsx_file
            if it.pic is not None:
                it.pic_off = len(out)
                out += struct.pack("<I", len(it.pic) + 24) + CLSID_STDPICTURE + b"lt\0\0" + struct.pack("<I", len(it.pic)) + it.pic
        return bytes(out)

    def text(self):
        self.dsx_bytes()                        # the offsets of the pictures
        L = ["VERSION 5.00", f"Begin {REPORT_GUID} {self.name} "]
        if self.datasource:
            L.append(f"   {'Bindings':<16}=   \"{self.dsx_file}\":0000")
        L += [_line(1, "Caption", self.caption), _line(1, "ClientHeight", self.client[0]), _line(1, "ClientLeft", self.client[1]),
              _line(1, "ClientTop", self.client[2]), _line(1, "ClientWidth", self.client[3]),
              _line(1, "_ExtentX", self.extent[0]), _line(1, "_ExtentY", self.extent[1]), _line(1, "_Version", 393216),
              _line(1, "_DesignerVersion", self.dver)]
        if self.width != 9360 or "reportwidth" in self.present:
            L.append(_line(1, "ReportWidth", self.width))
        L += font_lines(1, self.font)
        L += [_line(1, "GridX", self.gridx), _line(1, "GridY", self.gridy)]
        for n, v in zip(("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"), self.margins):
            if v != 1440 or n.lower() in self.present or not self.present:
                L.append(_line(1, n, v))
        if self.title:
            L.append(_line(1, "Title", self.title))
        if self.datamember:
            L.append(_line(1, "DataMember", self.datamember))
        L.append(_line(1, "NumSections", len(self.sections)))
        for i, s in enumerate(self.sections):
            L += s.to_lines(i)
        L.append("End")
        return "\n".join(L)

    def apply(self):
        _replace_designer(self.sf, self.text())


def dsx_bytes(name):
    """the .dsx of a DataReport: its DataSource (the DataEnvironment), as VB6 writes it"""
    n = name.encode("cp1252", "replace")[:255]
    return bytes([0xC6, 0xFA, 0x01, 0x00, 7, 0, 0, 0, len(n)]) + n + bytes(6)


def read_dsx(path):
    """the DataSource of a report: the name in its .dsx (path: the .dsr or the .dsx)"""
    base = os.path.splitext(path)[0]
    for cand in (base + ".dsx", base + ".DSX", base + ".Dsx"):
        try:
            data = open(cand, "rb").read()
        except OSError:
            continue
        return data[9:9 + data[8]].decode("cp1252", "replace") if len(data) > 9 else ""
    return ""


def new_report_text(name):
    m = ReportModel.__new__(ReportModel)
    m.name = name
    m.caption = name
    m.client = [8235, 1650, 1545, 6585]
    m.extent = [11615, 14526]
    m.dver = 100684049
    m.width = 9360
    m.gridx = m.gridy = 10
    m.margins = [1440, 1440, 1440, 1440]
    m.datamember = ""
    m.title = ""
    m.datasource = ""
    m.dsx_file = name + ".dsx"
    m.present = set()
    m.font = read_font(None)
    m.sections = [RptSection(2, "Section2", 360), RptSection(4, "Section1", 1440), RptSection(7, "Section3", 360)]
    return (m.text() + "\nAttribute VB_Name = \"{name}\"\nAttribute VB_GlobalNameSpace = False\n"
            "Attribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n").replace("{name}", name)


def dsx_path(m):
    return os.path.splitext(m.path)[0] + ".dsx" if m.path else None
