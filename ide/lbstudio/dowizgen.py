"""The Data Object Wizard (phase 9.2f) without Qt: what its pages decide and the modules it makes (VB6 oracle
DEV-9/do1).

A data class: a class that is a data source (DataSourceBehavior = vbDataSource) over a command of a
DataEnvironment of the project. Its rows are kept in a recordset set apart from the database; a lookup (a command
whose display field stands for a field of the rows: PubID -> the Name of the publisher) adds a descriptor field; the
Insert / Update / Delete commands chosen in the wizard take the changes back to the database. Its name is "rscls"
and the name given on the last page, as VB6 names it.

A user control bound to a data class: Single Record (a label and a TextBox / CheckBox / ComboBox for each field),
Data Grid, ListBox or ComboBox. Its name is "uct", the name given and the kind.

The VB_Ext_KEY attributes are the ones VB6's wizard writes (its classes and ours are read alike: a data class of
VB6's wizard can have a user control of ours and the other way round); the members the data class and the user
control show are VB6's too (InitX, MoveToPk, rsMoveComplete, ManualInitialize, SaveMode...). The code itself, its
comments and the texts are our own.
"""
import re

from .dfwizgen import himetric, vbstr

ENUM_MODULE = "clsDow"
ENUM_NAME = "EnumSaveMode"
CLASS_PREFIX = "rscls"
UC_PREFIX = "uct"
UC_TYPES = ["Single Record", "Data Grid", "ListBox", "ComboBox"]
UC_SUFFIX = {"Single Record": "SingleRecord", "Data Grid": "DataGrid", "ListBox": "ListBox", "ComboBox": "ComboBox"}
UC_TYPE_KEY = {"Single Record": "Single Record", "Data Grid": "DataGrid", "ListBox": "ListBox", "ComboBox": "ComboBox"}
CONTROL_TYPES = ["(None)", "TextBox", "CheckBox", "ComboBox"]

REFERENCES = [
    "*\\G{00000205-0000-0010-8000-00AA006D2EA4}#2.5#0#msado25.tlb#Microsoft ActiveX Data Objects 2.5 Library",
    "*\\G{7C0FFAB0-CD84-11D0-949A-00A0C91110ED}#1.0#0#msdatsrc.tlb#Microsoft Data Source Interfaces",
]
UC_REFERENCES = [
    "*\\G{56BF9020-7A2F-11D0-9482-00A0C91110ED}#1.0#0#MSBind.dll#Microsoft Data Binding Collection",
    "*\\G{6B263850-900B-11D0-9484-00A0C91110ED}#1.0#0#MSSTDFMT.DLL#Microsoft Data Formatting Object Library",
]
OBJ_GRID = '{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX'
LIB_GRID = "MSDataGridLib"

# the ADO types: the conversion of a value written to a field, the numeric ones, Boolean
CONVERT = {2: "CInt", 3: "CLng", 4: "CSng", 5: "CDbl", 6: "CCur", 7: "CDate", 11: "CBool", 16: "CInt", 17: "CByte",
           18: "CLng", 19: "CDbl", 133: "CDate", 135: "CDate", 64: "CDate"}
NUMERIC = {2, 3, 4, 5, 6, 14, 16, 17, 18, 19, 20, 21, 131, 139}
BOOLEAN = 11

NO_DE = "The project has no Data Environment with a command: add one with Project > Add Data Environment, " \
        "make a command and get its fields first."
NO_COMMAND = "Choose the command of the Data Environment the class takes its rows from."
NO_FIELDS = "The command '%s' has no fields yet: open the Data Environment, run the command once (its fields " \
            "appear under it) and start the wizard again."
SPACE_FIELD = "The field '%s' has a space in its name: a data class cannot have a property of that name.\n" \
              "Give the field a name without spaces in the command (... AS %s) and choose it again."
NO_CLASS_NAME = "Give the class a name."
BAD_NAME = "'%s' is not a valid name: a letter first, then letters, digits or _ (at most %d characters)."
NAME_USED = "The project has a module named '%s' already."
NO_DATA_CLASS = "Choose the data class the user control is bound to.\n" \
                "(The list has the classes the Data Object Wizard made: make one first with 'A Class object'.)"
NO_UC_NAME = "Give the user control a name."
UNMAPPED = "These parameters of the command have no field of the same name in the main command:\n%s\n\n" \
           "Their values are not set by the class: change its code, or give them the names of the fields " \
           "(the Parameters tab of the command)."


def valid_name(n, most=40):
    return bool(re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", n or "")) and len(n) <= most


# ---------------------------------------------------------------------------------------------- the project
class Field:
    def __init__(self, name, adotype=202, nullable=True, pk=False):
        self.name = name
        self.type = adotype
        self.nullable = nullable
        self.pk = pk
        self.control = "TextBox"                # a user control: (None) / TextBox / CheckBox / ComboBox


class Lookup:
    """a lookup: the command, its display field, the pairs (field of the rows, field of the lookup command)"""
    def __init__(self, command, display, pairs, nullable=True):
        self.command = command
        self.display = display
        self.pairs = list(pairs)
        self.nullable = nullable

    @property
    def descriptor(self):
        return self.command + self.display

    @property
    def source(self):
        return self.pairs[-1][0] if self.pairs else ""


class Change:
    """an Insert / Update / Delete command and its parameters: [(field of the rows, parameter)]"""
    def __init__(self, command="", pairs=None, unmapped=None):
        self.command = command
        self.pairs = list(pairs or [])
        self.unmapped = list(unmapped or [])


class DataClass:
    """what the pages decide for a data class"""
    def __init__(self):
        self.project = "Project1"
        self.de = ""                            # the DataEnvironment
        self.conn = ""                          # its connection
        self.command = ""                       # the main (Select) command
        self.fields = []                        # [Field]
        self.lookups = []                       # [Lookup]
        self.insert = Change()
        self.update = Change()
        self.delete = Change()
        self.insert_as_update = False
        self.root = ""                          # the name given on the last page
        self.stamp = ""

    @property
    def name(self):
        return CLASS_PREFIX + self.root

    def pk(self):
        return [f for f in self.fields if f.pk]

    def field(self, name):
        return next((f for f in self.fields if f.name.lower() == name.lower()), None)

    def lookup_of(self, field):
        """the lookup whose last field of the rows is this one"""
        return next((lk for lk in self.lookups if lk.source.lower() == field.lower()), None)

    def columns(self):
        """the fields of the class's recordset: each descriptor after the field it stands for"""
        out = []
        for f in self.fields:
            out.append(f.name)
            lk = self.lookup_of(f.name)
            if lk is not None:
                out.append(lk.descriptor)
        return out


class UserControlSpec:
    """what the pages decide for a user control"""
    def __init__(self, dc, kind="Single Record", root=""):
        self.dc = dc                            # the DataClass (read from the class's VB_Ext_KEY)
        self.kind = kind
        self.root = root or dc.root
        self.controls = {}                      # field (lower) -> control type, Single Record

    @property
    def name(self):
        return UC_PREFIX + self.root + UC_SUFFIX[self.kind]

    def control_of(self, f):
        return self.controls.get(f.name.lower(), default_control(self.dc, f))


def default_control(dc, f):
    if f.pk:
        return "(None)"
    if dc.lookup_of(f.name) is not None:
        return "ComboBox"
    if f.type == BOOLEAN:
        return "CheckBox"
    return "TextBox"


def control_choices(dc, f):
    out = ["(None)", "TextBox"]
    if f.type == BOOLEAN:
        out.append("CheckBox")
    if dc.lookup_of(f.name) is not None:
        out.append("ComboBox")
    return out


def data_environments(project):
    """[(module, DataEnvModel)] of the project"""
    from .dsrmodel import DataEnvModel, kind_of
    out = []
    for m in (project.modules if project else []):
        sf = getattr(m, "sf", None)
        if sf is not None and kind_of(sf) == "DataEnvironment":
            try:
                out.append((m, DataEnvModel(sf)))
            except Exception:           # a designer the IDE cannot read: not offered
                continue
    return out


def command_fields(cmd):
    return [Field(n, t) for n, t, _s, _p in cmd.fields]


def check_command(de, cmd):
    """None, or why the command cannot be the main command of a data class"""
    if cmd is None:
        return NO_COMMAND
    if not cmd.fields:
        return NO_FIELDS % cmd.name
    for n, *_ in cmd.fields:
        if " " in n.strip() or not valid_name(n.strip(), 255):
            return SPACE_FIELD % (n, re.sub(r"\W", "", n) or "Field1")
    return None


def new_class(project, de, cmd):
    dc = DataClass()
    dc.project = project.name if project is not None else "Project1"
    dc.de = de.name
    dc.command = cmd.name
    dc.conn = cmd.conn or (de.conns[0].name if de.conns else "")
    dc.fields = command_fields(cmd)
    return dc


def map_params(dc, cmd):
    """the parameters of a command against the fields of the rows (by name, as VB6 maps them): a Change"""
    ch = Change(cmd.name if cmd is not None else "")
    if cmd is None:
        return ch
    for p in cmd.params:
        if p.v.get("Direction", 1) not in (1, 3):
            continue
        f = dc.field(p.name)
        if f is None:
            ch.unmapped.append(p.name)
        else:
            ch.pairs.append((f.name, p.name))
    return ch


def _ext_keys(module):
    """the VB_Ext_KEY attributes of a module: {key lower: (key, value)}"""
    out = {}
    for a in getattr(module.sf, "attributes", None) or []:
        if a[0].strip().lower() != "vb_ext_key":
            continue
        m = re.match(r'\s*"([^"]*)"\s*,\s*"([^"]*)"', a[1] or "")
        if m:
            out[m.group(1).lower()] = (m.group(1), m.group(2))
    return out


def data_classes(project):
    """[(module, DataClass)]: the data classes of the project the wizard (VB6's or ours) made"""
    envs = {m.name.lower(): de for m, de in data_environments(project)}
    out = []
    for m in (project.modules if project else []):
        if m.kind != "Class" or getattr(m, "sf", None) is None:
            continue
        k = _ext_keys(m)
        if k.get("classtype", ("", ""))[1].lower() != "data class":
            continue
        dc = DataClass()
        g = lambda key, d="": k.get(key.lower(), ("", d))[1]
        dc.project = g("VBProjectName", project.name)
        dc.de = g("DEDesignerName")
        dc.conn = g("ConnectionName")
        dc.command = g("SelectCommandName") or g("CommandName")
        dc.root = g("ClassRootName") or m.name[len(CLASS_PREFIX):]
        de = envs.get(dc.de.lower())
        cmd = de.command(dc.command) if de is not None else None
        names = [f[0] for f in cmd.fields] if cmd is not None else \
            [kk[1][0][len("FieldPK"):] for kk in k.items() if kk[0].startswith("fieldpk")]
        types = {f[0].lower(): f[1] for f in cmd.fields} if cmd is not None else {}
        for n in names:
            f = Field(n, types.get(n.lower(), 202), g("FieldNullable" + n, "Yes").lower() == "yes",
                      g("FieldPK" + n, "No").lower() == "yes")
            dc.fields.append(f)
        for i in range(1, int(g("NumFKCommands", "0") or 0) + 1):
            pre = "FKCommand%d" % i
            pairs, j = [], 1
            while g("%s_SELECTFIELDNAME%d" % (pre, j)):
                pairs.append((g("%s_SELECTFIELDNAME%d" % (pre, j)), g("%s_FKFIELDNAME%d" % (pre, j))))
                j += 1
            display = g("%s_FKDESCRIPTOR1" % pre)
            if g(pre) and pairs:
                dc.lookups.append(Lookup(g(pre), display, pairs, g("%s_FKNULLABLE1" % pre, "Yes").lower() == "yes"))
        out.append((m, dc))
    return out


# ---------------------------------------------------------------------------------------------- the text
def _q(s):
    return vbstr(s)


def _ext(key, value):
    return 'Attribute VB_Ext_KEY = "%s" ,"%s"' % (key, value)


def _attrs(name, data_source):
    beh = "1  'vbDataSource" if data_source else "0  'vbNone"
    return ["VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "  Persistable = 0  'NotPersistable",
            "  DataBindingBehavior = 0  'vbNone", "  DataSourceBehavior  = %s" % beh,
            "  MTSTransactionMode  = 0  'NotAnMTSObject", "END",
            'Attribute VB_Name = "%s"' % name, "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = True", "Attribute VB_PredeclaredId = False", "Attribute VB_Exposed = False"]


def enum_text():
    """clsDow: the Enum of the SaveMode property (its name and members are VB6's: the classes use them)"""
    return "\n".join(_attrs(ENUM_MODULE, False) + [
        "' The SaveMode of the data classes and user controls of the Data Object Wizard.",
        "Public Enum EnumSaveMode",
        "    adImmediate = 0 ' a changed row goes to the database when another row becomes the current one",
        "    adBatch = 1     ' the changed rows go to the database when the program calls UpdateBatch",
        "End Enum", ""])


def class_keys(dc):
    L = [_ext("WizardYN", "Yes"), _ext("VBProjectName", dc.project), _ext("DEDesignerName", dc.de),
         _ext("ConnectionName", dc.conn), _ext("CommandName", dc.command), _ext("ClassRootName", dc.root),
         _ext("ClassType", "Data Class"), _ext("SelectCommandName", dc.command)]
    for f in dc.fields:
        L += [_ext("FieldNullable" + f.name, "Yes" if f.nullable else "No"),
              _ext("FieldPK" + f.name, "Yes" if f.pk else "No")]
    for i, lk in enumerate(dc.lookups, 1):
        pre = "FKCommand%d" % i
        L.append(_ext(pre, lk.command))
        for j, (sf, ff) in enumerate(lk.pairs, 1):
            L += [_ext("%s_SELECTFIELDNAME%d" % (pre, j), sf), _ext("%s_SELECTCOMMANDNAME%d" % (pre, j), dc.command),
                  _ext("%s_FKFIELDNAME%d" % (pre, j), ff), _ext("%s_FKCOMMANDNAME%d" % (pre, j), lk.command),
                  _ext("%s_FKDESCRIPTOR%d" % (pre, j), lk.display),
                  _ext("%s_FKNULLABLE%d" % (pre, j), "Yes" if lk.nullable else "No")]
        L.append(_ext("%s_FKPKCount" % pre, " %d" % len(lk.pairs)))
    L.append(_ext("NumFKCommands", str(len(dc.lookups))))
    for key, kind, ch in (("InsertCommand", "Insert", dc.insert), ("UpdateCommand", "Update", dc.update),
                          ("DeleteCommand", "Delete", dc.delete)):
        if key == "UpdateCommand" and dc.insert_as_update:
            continue
        if not ch.command:
            continue
        L.append(_ext(key + "Name", ch.command))
        for j, (f, p) in enumerate(ch.pairs, 1):
            L += [_ext("%s_FIELDNAME%d" % (key, j), f), _ext("%s_SELECTCOMMANDNAME%d" % (key, j), dc.command),
                  _ext("%s_PARAMETERCOMMANDNAME%d" % (key, j), ch.command),
                  _ext("%s_PARAMETERCOMMANDTYPE%d" % (key, j), kind), _ext("%s_PARAMETERNAME%d" % (key, j), p)]
        L.append(_ext("%s_FIELDPARAMETERCOUNT" % key, str(len(ch.pairs))))
        if key == "InsertCommand":
            L.append(_ext("InsertCommand_USEASUPDATE", "True" if dc.insert_as_update else "False"))
    L += [_ext("NumInterfaces", "1"), _ext("UseSelectCommand", "False")]
    return L


def class_text(dc):
    sel, de, cn = dc.command, dc.de, dc.conn
    rsel = "de1.rs" + sel
    numeric = [f for f in dc.fields if f.type in NUMERIC]
    update = dc.insert if dc.insert_as_update else dc.update
    changes = bool(dc.insert.command or update.command or dc.delete.command)
    L = _attrs(dc.name, True) + class_keys(dc)
    L += ["'-------------------------------------------------------------------------------",
          "' %s: a data class made by the Data Object Wizard of LinBASIC Studio" % dc.name,
          "' Its rows: the command %s of %s (%s). They are kept in a recordset of" % (sel, de, cn),
          "' their own, apart from the database: a bound control or a user control shows and",
          "' changes them there%s." % (", and the commands chosen in the wizard take the changes to the database"
                                         if changes else " (no command takes the changes to the database)"),
          "'-------------------------------------------------------------------------------",
          "Option Explicit", "",
          "Public SaveMode As EnumSaveMode        ' adImmediate: a row is saved when it is left; adBatch: by UpdateBatch",
          "", "Private de1 As New %s" % de,
          "Private WithEvents rs As ADODB.Recordset   ' the rows the class shows (Data of GetDataMember)"]
    for lk in dc.lookups:
        L.append("Public rs%s As ADODB.Recordset   ' the rows of the lookup %s" % (lk.command, lk.command))
    L += ["", "Private mbLoaded As Boolean      ' the rows are read",
          "Private mbReady As Boolean       ' GetDataMember has given the rows: changes are saved from now on",
          "Private mbAdding As Boolean      ' AddRecord: the new row is not saved yet",
          "Private mbDeleting As Boolean    ' Delete: the row is going",
          "Private mcolNew As New Collection ' the rows AddRecord added that are not in the database yet (their " \
          "Bookmark)", "",
          "Public Event rsMoveComplete()",
          "Public Event DeleteRecordComplete()",
          "Public Event rsUpdateEvent(vFieldName As Variant)",
          "Public Event ClassError(sProcedureName As String, oErr As ErrObject)", ""]
    # the fields as properties
    for f in dc.fields:
        conv = CONVERT.get(f.type)
        L += ["Public Property Get %s() As Variant" % f.name, "    %s = rs(%s)" % (f.name, _q(f.name)), "End Property",
              "Public Property Let %s(vValue As Variant)" % f.name,
              "    If IsNull(vValue) Then",
              "        rs(%s) = Null" % _q(f.name), "    Else",
              "        rs(%s) = %s" % (_q(f.name), "%s(vValue)" % conv if conv else "vValue"),
              "    End If", "End Property", ""]
    for lk in dc.lookups:
        d = lk.descriptor
        L += ["' the %s of the row of %s that %s stands for" % (lk.display, lk.command, lk.source),
              "Public Property Get %s() As Variant" % d, "    %s = rs(%s)" % (d, _q(d)), "End Property",
              "Public Property Let %s(vValue As Variant)" % d, "    rs(%s) = vValue" % _q(d), "End Property", ""]
    L += ["Public Property Get BOF() As Boolean", "    BOF = rs.BOF", "End Property", "",
          "Public Property Get EOF() As Boolean", "    EOF = rs.EOF", "End Property", "",
          "Public Property Get AbsolutePosition() As Long", "    AbsolutePosition = rs.AbsolutePosition", "End Property",
          "Public Property Let AbsolutePosition(lPosition As Long)", "    rs.AbsolutePosition = lPosition",
          "End Property", "",
          "Public Property Get RecordCount() As Long", "    RecordCount = rs.RecordCount", "End Property", ""]
    # GetDataMember
    L += ["' A bound control (or a BindingCollection) asks for the rows: DataMember is the name of the command.",
          "Private Sub Class_GetDataMember(DataMember As String, Data As Object)",
          "    On Error GoTo Failed", "    mbReady = False",
          "    Select Case LCase$(DataMember)",
          '    Case %s, ""' % _q(sel.lower()),
          "        If Not mbLoaded Then LoadRows",
          "        rs.UpdateBatch                  ' what was read is the state the changes start from",
          "        Set Data = rs"]
    for lk in dc.lookups:
        L += ["    Case %s" % _q(lk.command.lower()), "        If Not mbLoaded Then LoadRows",
              "        Set Data = rs%s" % lk.command]
    L += ["    End Select", "    mbReady = True", "    Exit Sub", "Failed:",
          '    RaiseEvent ClassError("GetDataMember", Err)', "End Sub", ""]
    # LoadRows
    L += ["' The rows of the command, apart from the database.", "Private Sub LoadRows()"]
    if dc.lookups:
        L += ["    Dim oField As ADODB.Field"]
    L += ["    If SaveMode = adImmediate Then",
          "        %s.LockType = adLockOptimistic" % rsel, "    Else",
          "        %s.LockType = adLockBatchOptimistic" % rsel, "    End If",
          "    de1.%s" % sel,
          "    Set %s.ActiveConnection = Nothing" % rsel]
    if not dc.lookups:
        L += ["    Set rs = %s.Clone" % rsel]
    else:
        for lk in dc.lookups:
            L += ["    de1.%s" % lk.command, "    Set de1.rs%s.ActiveConnection = Nothing" % lk.command,
                  "    Set rs%s = de1.rs%s" % (lk.command, lk.command)]
        L += ["    ' a recordset of its own: the fields of the command and the descriptors of the lookups",
              "    Set rs = New ADODB.Recordset", "    rs.CursorLocation = adUseClient", "    rs.CursorType = adOpenStatic",
              "    rs.LockType = %s.LockType" % rsel,
              "    For Each oField In %s.Fields" % rsel,
              "        rs.Fields.Append oField.Name, oField.Type, oField.DefinedSize, oField.Attributes",
              "        Select Case LCase$(oField.Name)"]
        for lk in dc.lookups:
            L += ["        Case %s" % _q(lk.source.lower()),
                  "            rs.Fields.Append %s, rs%s.Fields(%s).Type, rs%s.Fields(%s).DefinedSize, adFldIsNullable" % (
                      _q(lk.descriptor), lk.command, _q(lk.display), lk.command, _q(lk.display))]
        L += ["        End Select", "    Next", "    rs.Open",
              "    If %s.RecordCount > 0 Then %s.MoveFirst" % (rsel, rsel),
              "    Do While Not %s.EOF" % rsel, "        rs.AddNew",
              "        For Each oField In %s.Fields" % rsel,
              "            rs(oField.Name).Value = oField.Value", "        Next"]
        for lk in dc.lookups:
            L.append("        rs(%s).Value = LookupText%s(%s)" % (
                _q(lk.descriptor), lk.command, ", ".join("%s(%s).Value" % (rsel, _q(s)) for s, _ in lk.pairs)))
        L += ["        rs.Update", "        %s.MoveNext" % rsel, "    Loop",
              "    If rs.RecordCount > 0 Then rs.MoveFirst"]
    L += ["    de1.%s.Close" % cn, "    mbLoaded = True", "End Sub", ""]
    for lk in dc.lookups:
        args = ["v%d" % i for i in range(len(lk.pairs))]
        cond = " And ".join("rs%s(%s).Value = %s" % (lk.command, _q(ff), a) for (_s, ff), a in zip(lk.pairs, args))
        L += ["' the %s of the row of %s for these values (\"(None)\" for no row)" % (lk.display, lk.command),
              "Private Function LookupText%s(%s) As Variant" % (lk.command, ", ".join("ByVal %s As Variant" % a for a in args)),
              "    LookupText%s = \"(None)\"" % lk.command,
              "    If rs%s.RecordCount = 0 Then Exit Function" % lk.command,
              "    rs%s.MoveFirst" % lk.command, "    Do While Not rs%s.EOF" % lk.command,
              "        If %s Then" % cond,
              "            If Not IsNull(rs%s(%s).Value) Then LookupText%s = rs%s(%s).Value" % (
                  lk.command, _q(lk.display), lk.command, lk.command, _q(lk.display)),
              "            Exit Function", "        End If", "        rs%s.MoveNext" % lk.command, "    Loop",
              "End Function", ""]
    L += ["Private Sub Class_Terminate()", "    Set rs = Nothing"]
    for lk in dc.lookups:
        L.append("    Set rs%s = Nothing" % lk.command)
    L += ["    Set de1 = Nothing", "End Sub", ""]
    # moving
    L += ["Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, _",
          "        adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)",
          "    If rs.BOF Or rs.EOF Then Exit Sub   ' no current row",
          "    RaiseEvent rsMoveComplete", "End Sub", "",
          "Public Sub Move(lRows As Long)", "    On Error Resume Next", "    rs.Move lRows", "End Sub", "",
          "Public Sub MoveFirst()", "    If rs.RecordCount > 0 Then rs.MoveFirst", "End Sub", "",
          "Public Sub MovePrevious()", "    If rs.RecordCount > 0 Then", "        If Not rs.BOF Then rs.MovePrevious",
          "    End If", "End Sub", "",
          "Public Sub MoveNext()", "    If rs.RecordCount > 0 Then", "        If Not rs.EOF Then rs.MoveNext",
          "    End If", "End Sub", "",
          "Public Sub MoveLast()", "    If rs.RecordCount > 0 Then rs.MoveLast", "End Sub", ""]
    pk = dc.pk()
    if pk:
        args = ", ".join("%s As Variant" % f.name for f in pk)
        cond = " And ".join("rs(%s) = %s" % (_q(f.name), f.name) for f in pk)
        L += ["' the row of these key values becomes the current one (False: there is none)",
              "Public Function MoveToPk(%s) As Boolean" % args,
              "    rs.MoveFirst", "    Do While Not rs.EOF", "        If %s Then" % cond,
              "            MoveToPk = True", "            Exit Function", "        End If", "        rs.MoveNext",
              "    Loop", "End Function", ""]
    # checking a row
    L += ["' False (and a message) when the current row cannot be saved",
          "Public Function ValidateData() As Boolean"]
    for f in dc.fields:
        if not f.nullable:
            L += ["    If IsNull(rs(%s).Value) Then" % _q(f.name),
                  "        MsgBox \"%s must have a value.\"" % f.name, "        Exit Function", "    End If"]
    for f in numeric:
        L += ["    If Not IsNull(rs(%s).Value) Then" % _q(f.name),
              "        If Not IsNumeric(rs(%s).Value) Then" % _q(f.name),
              "            MsgBox \"%s must be a number.\"" % f.name, "            Exit Function", "        End If",
              "    End If"]
    L += ["    ValidateData = True", "End Function", ""]
    # saving
    L += ["Public Sub UpdateBatch()", "    On Error GoTo Failed", "    rs.UpdateBatch", "    Exit Sub", "Failed:",
          '    RaiseEvent ClassError("UpdateBatch", Err)', "End Sub", "",
          "Public Sub Update()", "    On Error GoTo Failed", "    rs.UpdateBatch adAffectCurrent", "    Exit Sub", "Failed:",
          '    RaiseEvent ClassError("Update", Err)', "End Sub", "",
          "' a new empty row at the end, the current one",
          "Public Sub AddRecord()", "    mbAdding = True", "    rs.AddNew", "    rs.Update",
          "    mcolNew.Add True, CStr(rs.Bookmark)", "    mbAdding = False",
          "    rs.MoveLast", "End Sub", "",
          "' the current row was added by AddRecord and is not in the database yet",
          "Private Function IsNewRow() As Boolean", "    Dim v As Variant", "    On Error Resume Next",
          "    v = mcolNew.Item(CStr(rs.Bookmark))", "    IsNewRow = (Err.Number = 0)", "End Function", "",
          "Private Sub ForgetNewRow()", "    On Error Resume Next", "    mcolNew.Remove CStr(rs.Bookmark)", "End Sub", "",
          "Public Sub Delete()", "    On Error GoTo Failed", "    mbDeleting = True", "    rs.Delete adAffectCurrent",
          "    mbDeleting = False", "    MovePrevious", "    If rs.BOF And rs.RecordCount > 0 Then rs.MoveFirst",
          "    RaiseEvent DeleteRecordComplete", "    Exit Sub", "Failed:", "    mbDeleting = False",
          '    RaiseEvent ClassError("Delete", Err)', "End Sub", "",
          "' a field changed (a user control calls it): the descriptors follow, then rsUpdateEvent",
          "Public Sub rsUpdate(vFieldName As Variant)"]
    for lk in dc.lookups:
        L.append("    rs(%s).Value = LookupText%s(%s)" % (
            _q(lk.descriptor), lk.command, ", ".join("rs(%s).Value" % _q(s) for s, _ in lk.pairs)))
    L += ["    RaiseEvent rsUpdateEvent(vFieldName)", "End Sub", ""]
    if changes:
        L += ["' a row changes: the commands take the change to the database (Cancel when they fail)",
              "Private Sub rs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, _",
              "        adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)",
              "    Dim bOK As Boolean",
              "    If Not mbReady Or mbAdding Then Exit Sub",
              "    bOK = True", "    Select Case adReason",
              "    Case adRsnDelete                    ' a row added here and not saved has nothing to delete",
              "        If IsNewRow() Then ForgetNewRow Else bOK = SaveDelete()",
              "    Case adRsnUpdate                    ' the row is left (or Update): a new row is inserted",
              "        If mbDeleting Then Exit Sub     ' (the Update that follows a Delete)",
              "        If rs.EditMode = adEditAdd Or IsNewRow() Then",
              "            bOK = SaveInsert()", "            If bOK Then ForgetNewRow", "        Else",
              "            bOK = SaveUpdate()", "        End If",
              "    End Select", "    If Not bOK Then adStatus = adStatusCancel", "End Sub", ""]
        for fn, ch, check in (("SaveInsert", dc.insert, True), ("SaveUpdate", update, True),
                              ("SaveDelete", dc.delete, False)):
            L += ["Private Function %s() As Boolean" % fn]
            if not ch.command:
                L += ["    %s = True                        ' no command: the change stays in the rows" % fn,
                      "End Function", ""]
                continue
            L += ["    Dim oCommand As ADODB.Command", "    On Error GoTo Failed"]
            if check:
                L += ["    If Not ValidateData() Then", '        RaiseEvent ClassError("ValidateData", Err)',
                      "        Exit Function", "    End If"]
            L += ["    If de1.%s.State = adStateClosed Then de1.%s.Open" % (cn, cn),
                  "    Set oCommand = de1.Commands(%s)" % _q(ch.command)]
            for f, p in ch.pairs:
                L.append("    oCommand.Parameters(%s).Value = rs(%s).Value" % (_q(p), _q(f)))
            L += ["    oCommand.Execute", "    %s = True" % fn, "    Exit Function", "Failed:",
                  '    RaiseEvent ClassError("%s", Err)' % fn, "End Function", ""]
    return "\n".join(L)


# ---------------------------------------------------------------------------------------------- user controls
def _uc_head(name, w, h, objects=()):
    L = ["VERSION 5.00"]
    for o in objects:
        L.append('Object = "%s"' % o)
    L += ["Begin VB.UserControl %s " % name, "   ClientHeight    =   %d" % h, "   ClientLeft      =   0",
          "   ClientTop       =   0", "   ClientWidth     =   %d" % w, "   ScaleHeight     =   %d" % h,
          "   ScaleWidth      =   %d" % w]
    return L


def _ctl(cls, name, props, depth=1):
    pad = "   " * depth
    L = ["%sBegin %s %s " % (pad, cls, name)]
    for k, v in props:
        L.append("%s   %s=   %s" % (pad, k.ljust(16), v))
    L.append(pad + "End")
    return L


def _uc_attrs(spec):
    dc = spec.dc
    return ['Attribute VB_Name = "%s"' % spec.name, "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = True", "Attribute VB_PredeclaredId = False", "Attribute VB_Exposed = False",
            _ext("WizardYN", "Yes"), _ext("VBProjectName", dc.project), _ext("DEDesignerName", dc.de),
            _ext("ConnectionName", dc.conn), _ext("CommandName", dc.command), _ext("ClassType", "User Control"),
            _ext("DataClassName", dc.name), _ext("RootName", spec.name),
            _ext("UserControlType", UC_TYPE_KEY[spec.kind])]


def _uc_comment(spec, what):
    dc = spec.dc
    return ["'-------------------------------------------------------------------------------",
            "' %s: a user control made by the Data Object Wizard of LinBASIC Studio" % spec.name,
            "' %s of the data class %s (%s.%s)." % (what, dc.name, dc.de, dc.command),
            "' Init%s binds it; with ManualInitialize False it is called when the control" % dc.root,
            "' is loaded on a running form.",
            "'-------------------------------------------------------------------------------",
            "Option Explicit", ""]


def _uc_init_head(spec, var):
    dc = spec.dc
    return ["Public Function Init%s(Optional oDataSource0 As Variant) As Object" % dc.root,
            "    Set %s = Nothing" % var,
            "    If IsMissing(oDataSource0) Then",
            "        Set %s = New %s" % (var, dc.name), "    Else", "        Set %s = oDataSource0" % var, "    End If"]


def _uc_props(names, save_mode, init):
    """ReadProperties / WriteProperties: ManualInitialize, the others, SaveMode (defaults when the form has none)"""
    L = ["Private Sub UserControl_ReadProperties(PropBag As PropertyBag)",
         '    ManualInitialize = PropBag.ReadProperty("ManualInitialize", False)']
    for n, d in names:
        L.append('    %s = PropBag.ReadProperty("%s", %s)' % (n, n, d))
    if save_mode:
        L.append('    SaveMode = PropBag.ReadProperty("SaveMode", adImmediate)')
    L += ["    If Not ManualInitialize And Ambient.UserMode Then %s" % init, "End Sub", "",
          "Private Sub UserControl_WriteProperties(PropBag As PropertyBag)",
          '    PropBag.WriteProperty "ManualInitialize", ManualInitialize']
    for n, _d in names:
        L.append('    PropBag.WriteProperty "%s", %s' % (n, n))
    if save_mode:
        L.append('    PropBag.WriteProperty "SaveMode", SaveMode')
    L += ["End Sub", ""]
    return L


def _class_error(var):
    return ["Private Sub %s_ClassError(sProcedureName As String, oErr As ErrObject)" % var,
            "    mbClassError = True",
            "    If oErr.Description <> \"\" Then",
            "        MsgBox sProcedureName & \": \" & oErr.Description", "    Else",
            "        MsgBox sProcedureName", "    End If", "End Sub", ""]


def _key_args(dc):
    pk = dc.pk()
    return pk if pk else dc.fields[:1]


def uc_single_text(spec):
    dc = spec.dc
    var = "o" + dc.root
    shown = [(f, spec.control_of(f)) for f in dc.fields if spec.control_of(f) != "(None)"]
    n = len(shown)
    w, h = 4605, max(480, 120 + 480 * n)
    body = []
    for k in reversed(range(n)):                # VB6 writes the newest control first
        f, kind = shown[k]
        top = 120 + 480 * k
        if kind == "CheckBox":
            body += _ctl("VB.CheckBox", "chk" + f.name, [("Height", 315), ("Left", 1920), ("TabIndex", 2 * k + 1),
                                                         ("Top", top), ("Width", 2500)])
        elif kind == "ComboBox":
            body += _ctl("VB.ComboBox", "cbo" + f.name, [("Height", 315), ("Left", 1920), ("Style", "2  'Dropdown List"),
                                                         ("TabIndex", 2 * k + 1), ("Top", top), ("Width", 2500)])
        else:
            body += _ctl("VB.TextBox", "txt" + f.name, [("Height", 315), ("Left", 1920), ("TabIndex", 2 * k + 1),
                                                        ("Text", _q("txt" + f.name)), ("Top", top), ("Width", 2500)])
    for k in reversed(range(n)):
        f, _ = shown[k]
        body += _ctl("VB.Label", "lbl" + f.name, [("Caption", _q(f.name + ": ")), ("Height", 285), ("Left", 120),
                                                  ("TabIndex", 2 * k), ("Top", 120 + 480 * k), ("Width", 1800)])
    L = _uc_head(spec.name, w, h) + body + ["End"] + _uc_attrs(spec)
    L += _uc_comment(spec, "A record at a time")
    combos = [(f, dc.lookup_of(f.name)) for f, kind in shown if kind == "ComboBox"]
    checks = [f for f, kind in shown if kind == "CheckBox"]
    L += ["Public ManualInitialize As Boolean   ' True: the program calls Init%s itself" % dc.root,
          "Public SaveMode As EnumSaveMode", "Public Event MoveComplete(oDataSource As Object)", "",
          "Private WithEvents %s As %s" % (var, dc.name), "Private oColBind As BindingCollection",
          "Private mbClassError As Boolean", "Private mbShowing As Boolean   ' the lookups are being set: no change"]
    if checks:
        L.append("Private oCheckFormat As StdDataFormat")
    L.append("")
    L += _uc_init_head(spec, var)
    L += ["    %s.SaveMode = SaveMode" % var, "    mbClassError = False",
          "    Set oColBind = New BindingCollection",
          "    oColBind.DataMember = %s" % _q(dc.command),
          "    Set oColBind.DataSource = %s" % var,
          "    If mbClassError Then", '        MsgBox "The rows of %s could not be read."' % dc.name,
          "        Exit Function", "    End If"]
    if checks:
        L += ["    Set oCheckFormat = New StdDataFormat", "    oCheckFormat.Type = fmtCheckbox",
              "    oCheckFormat.FalseValue = 0", "    oCheckFormat.TrueValue = 1"]
    for f, kind in shown:
        if kind == "TextBox":
            L.append('    oColBind.Add txt%s, "Text", %s' % (f.name, _q(f.name)))
        elif kind == "CheckBox":
            L.append('    oColBind.Add chk%s, "Value", %s, oCheckFormat' % (f.name, _q(f.name)))
    for f, lk in combos:
        L.append('    FillLookup cbo%s, %s.rs%s, %s, %s' % (f.name, var, lk.command, _q(lk.display),
                                                         "True" if lk.nullable else "False"))
    L += ["    ShowLookups", "    Set Init%s = %s" % (dc.root, var), "End Function", ""]
    pk = dc.pk()
    if pk:
        args = ", ".join("%s As Variant" % f.name for f in pk)
        L += ["Public Function MoveToPK(%s) As Boolean" % args,
              "    MoveToPK = %s.MoveToPk(%s)" % (var, ", ".join(f.name for f in pk)), "End Function", ""]
    L += ["Private Sub %s_rsMoveComplete()" % var, "    ShowLookups",
          "    RaiseEvent MoveComplete(%s)" % var, "End Sub", ""]
    L += _class_error(var)
    # the lookups
    L += ["' the rows of a lookup in a ComboBox: its display field, ItemData the position of the row",
          "Private Sub FillLookup(cbo As ComboBox, oRS As ADODB.Recordset, sDisplay As String, bNone As Boolean)",
          "    cbo.Clear", "    If bNone Then", '        cbo.AddItem "(None)"', "        cbo.ItemData(cbo.NewIndex) = -1",
          "    End If", "    If oRS.RecordCount = 0 Then Exit Sub", "    oRS.MoveFirst", "    Do While Not oRS.EOF",
          "        cbo.AddItem oRS(sDisplay).Value & \"\"",
          "        cbo.ItemData(cbo.NewIndex) = oRS.AbsolutePosition", "        oRS.MoveNext", "    Loop", "End Sub", "",
          "' the ComboBoxes of the lookups show the current row's",
          "Private Sub ShowLookups()"]
    if combos:
        L += ["    Dim i As Integer", "    mbShowing = True", "    On Error Resume Next"]
        for f, lk in combos:
            L += ["    cbo%s.ListIndex = -1" % f.name,
                  "    For i = 0 To cbo%s.ListCount - 1" % f.name,
                  "        If cbo%s.List(i) = %s.%s & \"\" Then cbo%s.ListIndex = i: Exit For" % (
                      f.name, var, lk.descriptor, f.name), "    Next"]
        L += ["    On Error GoTo 0", "    mbShowing = False"]
    L += ["End Sub", ""]
    for f, lk in combos:
        (sfield, ffield) = lk.pairs[-1]
        L += ["Private Sub cbo%s_Click()" % f.name, "    Dim lPos As Long",
              "    If mbShowing Or cbo%s.ListIndex < 0 Then Exit Sub" % f.name,
              "    lPos = cbo%s.ItemData(cbo%s.ListIndex)" % (f.name, f.name), "    If lPos = -1 Then",
              "        %s.%s = Null" % (var, sfield), "    Else",
              "        %s.rs%s.AbsolutePosition = lPos" % (var, lk.command),
              "        %s.%s = %s.rs%s(%s).Value" % (var, sfield, var, lk.command, _q(ffield)), "    End If",
              "    %s.rsUpdate %s" % (var, _q(sfield)), "End Sub", ""]
    L += ["Public Sub MoveFirst()", "    %s.MoveFirst" % var, "End Sub", "",
          "Public Sub MovePrevious()", "    %s.MovePrevious" % var, "End Sub", "",
          "Public Sub MoveNext()", "    %s.MoveNext" % var, "End Sub", "",
          "Public Sub MoveLast()", "    %s.MoveLast" % var, "End Sub", "",
          "Public Sub AddRecord()", "    %s.Move 0                    ' the row being changed is saved first" % var,
          "    %s.AddRecord" % var, "End Sub", "",
          "Public Sub Delete()", "    %s.Delete" % var, "End Sub", "",
          "Public Sub UpdateBatch()", "    %s.UpdateBatch" % var, "End Sub", ""]
    L += _uc_props([], True, "Init" + dc.root)
    L += ["Private Sub UserControl_Terminate()", "    Set oColBind = Nothing", "    Set %s = Nothing" % var,
          "End Sub", ""]
    return "\n".join(L)


def uc_list_text(spec):
    """ListBox / ComboBox: the first field that is not a key, ItemData the position of the row"""
    dc = spec.dc
    var = "o" + dc.root
    combo = spec.kind == "ComboBox"
    ctl = ("cbo" if combo else "lst") + dc.root
    show = next((f for f in dc.fields if not f.pk), dc.fields[0])
    keys = _key_args(dc)
    ev = "ComboBoxClick" if combo else "ListBoxClick"
    if combo:
        body = _ctl("VB.ComboBox", ctl, [("Height", 315), ("Left", 120), ("Style", "2  'Dropdown List"),
                                         ("TabIndex", 0), ("Top", 120), ("Width", 1000)])
    else:
        body = _ctl("VB.ListBox", ctl, [("Height", 840), ("Left", 120), ("TabIndex", 0), ("Top", 120), ("Width", 1000)])
    L = _uc_head(spec.name, 4800, 3600) + body + ["End"] + _uc_attrs(spec)
    L += _uc_comment(spec, "The %s of the rows in a %s" % (show.name, "ComboBox" if combo else "ListBox"))
    L += ["Public ManualInitialize As Boolean   ' True: the program calls Init%s itself" % dc.root,
          "Public NoneFirst As Boolean          ' a (None) item first",
          "Public Event %s(%s)" % (ev, ", ".join("%s As Variant" % f.name for f in keys)), "",
          "Private WithEvents %s As %s" % (var, dc.name), "Private mbClassError As Boolean", ""]
    L += _uc_init_head(spec, var)
    L += ["    mbClassError = False", "    %s.DataMember = %s" % (ctl, _q(dc.command)),
          "    Set %s.DataSource = %s" % (ctl, var),
          "    If mbClassError Then", '        MsgBox "The rows of %s could not be read."' % dc.name,
          "        Exit Function", "    End If",
          "    %s.Clear" % ctl, "    If NoneFirst Then", '        %s.AddItem "(None)"' % ctl,
          "        %s.ItemData(%s.NewIndex) = -1" % (ctl, ctl), "    End If",
          "    %s.MoveFirst" % var, "    Do While Not %s.EOF" % var,
          "        If Not IsNull(%s.%s) Then" % (var, show.name),
          "            %s.AddItem %s.%s & \"\"" % (ctl, var, show.name),
          "            %s.ItemData(%s.NewIndex) = %s.AbsolutePosition" % (ctl, ctl, var), "        End If",
          "        %s.MoveNext" % var, "    Loop", "    UserControl_Resize",
          "    If %s.ListCount > 0 Then %s.ListIndex = 0" % (ctl, ctl),
          "    Set Init%s = %s" % (dc.root, var), "End Function", ""]
    L += ["Private Sub %s_Click()" % ctl, "    If %s.ListIndex < 0 Then Exit Sub" % ctl,
          "    If %s.ItemData(%s.ListIndex) = -1 Then" % (ctl, ctl),
          "        RaiseEvent %s(%s)" % (ev, ", ".join("Null" for _ in keys)), "    Else",
          "        %s.AbsolutePosition = %s.ItemData(%s.ListIndex)" % (var, ctl, ctl),
          "        RaiseEvent %s(%s)" % (ev, ", ".join("%s.%s" % (var, f.name) for f in keys)), "    End If",
          "End Sub", ""]
    L += _class_error(var)
    L += ["Private Sub UserControl_Resize()", "    On Error Resume Next",
          "    %s.Move UserControl.ScaleLeft, UserControl.ScaleTop, UserControl.ScaleWidth%s" % (
              ctl, "" if combo else ", UserControl.ScaleHeight"), "End Sub", ""]
    L += _uc_props([("NoneFirst", "False")], False, "Init" + dc.root)
    L += ["Private Sub UserControl_Terminate()", "    Set %s = Nothing" % var, "End Sub", ""]
    return "\n".join(L)


def uc_grid_text(spec):
    """Data Grid: the columns of the class's recordset; a lookup shows its descriptor (its field hidden) and is
    changed with a ComboBox over the cell"""
    dc = spec.dc
    var = "o" + dc.root
    grid = "grd" + dc.root
    cols = dc.columns()
    consts = [(c.upper() + "_COLUMN", i) for i, c in enumerate(cols)]
    grid_props = [("Height", 3400), ("Left", 120), ("TabIndex", 0), ("Top", 120), ("Width", 4600),
                  ("_ExtentX", himetric(4600)), ("_ExtentY", himetric(3400)), ("_Version", 393216),
                  ("HeadLines", 1), ("RowHeight", 15)]
    body = _ctl("VB.ComboBox", "cmbEdit", [("Appearance", "0  'Flat"), ("Height", 315), ("Left", 480),
                                           ("Style", "2  'Dropdown List"), ("TabIndex", 1), ("Top", 1320),
                                           ("Visible", "0   'False"), ("Width", 1000)])
    body += _ctl("MSDataGridLib.DataGrid", grid, grid_props)
    L = _uc_head(spec.name, 4800, 3600, [OBJ_GRID]) + body + ["End"] + _uc_attrs(spec)
    L += _uc_comment(spec, "The rows in a DataGrid")
    L += ["Public ManualInitialize As Boolean   ' True: the program calls Init%s itself" % dc.root,
          "Public GridEditable As Boolean       ' the rows can be added, changed and deleted in the grid",
          "Public SaveMode As EnumSaveMode", "Public Event MoveComplete(oDataSource As Object)", "",
          "Private WithEvents %s As %s" % (var, dc.name), "Private mbClassError As Boolean",
          "Private mbLoading As Boolean", "Private mlEditCol As Long      ' the column the ComboBox is over"]
    for c, i in consts:
        L.append("Private Const %s As Long = %d" % (c, i))
    L.append("")
    L += ["Public Function Init%s(Optional oDataSource0 As Variant) As Object" % dc.root,
          "    mbLoading = True", "    Set %s = Nothing" % var, "    If IsMissing(oDataSource0) Then",
          "        Set %s = New %s" % (var, dc.name), "    Else", "        Set %s = oDataSource0" % var, "    End If",
          "    %s.SaveMode = SaveMode" % var, "    mbClassError = False",
          "    %s.DataMember = %s" % (grid, _q(dc.command)), "    Set %s.DataSource = %s" % (grid, var),
          "    If mbClassError Then", '        MsgBox "The rows of %s could not be read."' % dc.name,
          "        mbLoading = False", "        Exit Function", "    End If",
          "    %s.AllowAddNew = GridEditable" % grid, "    %s.AllowDelete = GridEditable" % grid,
          "    %s.AllowUpdate = GridEditable" % grid]
    for c, i in consts:
        L.append("    %s.Columns(%s).Caption = %s" % (grid, c, _q(cols[i])))
    for lk in dc.lookups:
        L.append("    %s.Columns(%s).Width = 0               ' its descriptor is shown" % (
            grid, lk.source.upper() + "_COLUMN"))
    L += ["    UserControl_Resize", "    mbLoading = False", "    Set Init%s = %s" % (dc.root, var),
          "End Function", ""]
    L += ["Private Sub %s_rsMoveComplete()" % var, "    RaiseEvent MoveComplete(%s)" % var, "End Sub", ""]
    L += _class_error(var)
    L += ["Private Sub UserControl_Resize()", "    On Error Resume Next",
          "    %s.Move UserControl.ScaleLeft, UserControl.ScaleTop, UserControl.ScaleWidth, UserControl.ScaleHeight"
          % grid, "End Sub", ""]
    # the lookups: a ComboBox over the cell of the descriptor
    L += ["Private Sub %s_RowColChange(LastRow As Variant, ByVal LastCol As Integer)" % grid,
          "    cmbEdit.Visible = False", "    If mbLoading Or Not GridEditable Then Exit Sub",
          "    mlEditCol = %s.Col" % grid, "    Select Case mlEditCol"]
    for lk in dc.lookups:
        L += ["    Case %s" % (lk.descriptor.upper() + "_COLUMN"),
              "        ShowEdit %s.rs%s, %s, %s" % (var, lk.command, _q(lk.display), "True" if lk.nullable else "False")]
    L += ["    End Select", "End Sub", "",
          "Private Sub ShowEdit(oRS As ADODB.Recordset, sDisplay As String, bNone As Boolean)",
          "    Dim i As Integer", "    cmbEdit.Clear", "    If bNone Then", '        cmbEdit.AddItem "(None)"',
          "        cmbEdit.ItemData(cmbEdit.NewIndex) = -1", "    End If",
          "    If oRS.RecordCount > 0 Then oRS.MoveFirst", "    Do While Not oRS.EOF",
          "        cmbEdit.AddItem oRS(sDisplay).Value & \"\"",
          "        cmbEdit.ItemData(cmbEdit.NewIndex) = oRS.AbsolutePosition", "        oRS.MoveNext", "    Loop",
          "    For i = 0 To cmbEdit.ListCount - 1",
          "        If cmbEdit.List(i) = %s.Columns(mlEditCol).Text Then cmbEdit.ListIndex = i: Exit For" % grid,
          "    Next", "    On Error Resume Next",
          "    cmbEdit.Move %s.Left + %s.Columns(mlEditCol).Left, %s.Top + %s.RowTop(%s.Row), %s.Columns(mlEditCol).Width"
          % (grid, grid, grid, grid, grid, grid),
          "    cmbEdit.Visible = True", "    cmbEdit.ZOrder 0", "    cmbEdit.SetFocus", "End Sub", "",
          "Private Sub cmbEdit_Click()", "    Dim lPos As Long",
          "    If Not cmbEdit.Visible Or cmbEdit.ListIndex < 0 Then Exit Sub",
          "    lPos = cmbEdit.ItemData(cmbEdit.ListIndex)", "    Select Case mlEditCol"]
    for lk in dc.lookups:
        sfield, ffield = lk.pairs[-1]
        L += ["    Case %s" % (lk.descriptor.upper() + "_COLUMN"), "        If lPos = -1 Then",
              "            %s.%s = Null" % (var, sfield), "        Else",
              "            %s.rs%s.AbsolutePosition = lPos" % (var, lk.command),
              "            %s.%s = %s.rs%s(%s).Value" % (var, sfield, var, lk.command, _q(ffield)), "        End If",
              "        %s.rsUpdate %s" % (var, _q(sfield))]
    L += ["    End Select", "    cmbEdit.Visible = False", "    %s.Refresh" % grid, "End Sub", "",
          "Private Sub %s_Scroll(Cancel As Integer)" % grid, "    cmbEdit.Visible = False", "End Sub", "",
          "Private Sub %s_KeyPress(KeyAscii As Integer)" % grid, "    Select Case %s.Col" % grid]
    for lk in dc.lookups:
        L += ["    Case %s" % (lk.descriptor.upper() + "_COLUMN"),
              "        KeyAscii = 0                 ' the descriptor is chosen in the ComboBox"]
    L += ["    End Select", "End Sub", ""]
    L += _uc_props([("GridEditable", "False")], True, "Init" + dc.root)
    L += ["Private Sub UserControl_Terminate()", "    Set %s = Nothing" % var, "End Sub", ""]
    return "\n".join(L)


def uc_text(spec):
    if spec.kind == "Data Grid":
        return uc_grid_text(spec)
    if spec.kind in ("ListBox", "ComboBox"):
        return uc_list_text(spec)
    return uc_single_text(spec)


# ---------------------------------------------------------------------------------------------- the result
class Output:
    def __init__(self):
        self.modules = []                       # (kind "Class" | "UserControl", name, text with \n)
        self.references = list(REFERENCES)
        self.libraries = []


def generate_class(dc, has_enum):
    out = Output()
    if not has_enum:
        out.modules.append(("Class", ENUM_MODULE, enum_text()))
    out.modules.append(("Class", dc.name, class_text(dc)))
    return out


def generate_uc(spec, has_enum):
    out = Output()
    if not has_enum:
        out.modules.append(("Class", ENUM_MODULE, enum_text()))
    out.modules.append(("UserControl", spec.name, uc_text(spec)))
    out.references += UC_REFERENCES
    if spec.kind == "Data Grid":
        out.libraries.append(LIB_GRID)
    return out


ADO_GUIDS = {"{00000200-0000-0010-8000-00aa006d2ea4}", "{00000201-0000-0010-8000-00aa006d2ea4}",
             "{00000205-0000-0010-8000-00aa006d2ea4}", "{00000206-0000-0010-8000-00aa006d2ea4}",
             "{ef53050b-882e-4776-b643-eda472e8e3f2}", "{2a75196c-d9eb-4129-b803-931327f72d5c}",
             "{b691e011-1797-432e-907a-4d8c69339129}"}


def _guid(line):
    return (line or "").split("#")[0].lstrip("*\\G").strip().lower()


def needed_references(project, refs):
    """the references to add: not ADO when the project has one of its versions already (the Data Environment's:
    two ADODB libraries in one project are a name conflict for VB6)"""
    have = {_guid(v) for k, v in (project.vbp.entries if project is not None and project.vbp else [])
            if v is not None and k.strip() == "Reference"}
    return [r for r in refs if not (_guid(r) in ADO_GUIDS and have & ADO_GUIDS)]


def has_enum(project):
    """the project has the Enum EnumSaveMode (clsDow, VB6's or ours)"""
    for m in (project.modules if project else []):
        if m.name.lower() == ENUM_MODULE.lower():
            return True
        try:
            text = m.code_text()
        except Exception:
            continue
        if re.search(r"(?im)^\s*(public\s+)?enum\s+%s\b" % ENUM_NAME, text or ""):
            return True
    return False


def check_class_name(project, root):
    if not (root or "").strip():
        return NO_CLASS_NAME
    name = CLASS_PREFIX + root
    if not valid_name(root) or not valid_name(name):
        return BAD_NAME % (root, 40 - len(CLASS_PREFIX))
    if project is not None and project.module(name) is not None:
        return NAME_USED % name
    return None


def check_uc_name(project, spec):
    if not (spec.root or "").strip():
        return NO_UC_NAME
    if not valid_name(spec.root) or not valid_name(spec.name):
        return BAD_NAME % (spec.root, 40 - len(UC_PREFIX) - len(UC_SUFFIX[spec.kind]))
    if project is not None and project.module(spec.name) is not None:
        return NAME_USED % spec.name
    return None


__all__ = ["DataClass", "Field", "Lookup", "Change", "UserControlSpec", "data_environments", "data_classes",
           "check_command", "new_class", "map_params", "class_text", "uc_text", "enum_text", "generate_class",
           "generate_uc", "has_enum", "check_class_name", "check_uc_name", "control_choices", "default_control",
           "UC_TYPES", "CONTROL_TYPES"]
