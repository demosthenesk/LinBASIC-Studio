"""The forms of the Data Form Wizard (phase 9.2b): the settings of the wizard's pages (DataForm) become the text
of a form (.frm), and for the Class binding a data source class (.cls) too. The code and the wording are LinBASIC
Studio's own; what the forms do follows the forms of the VB6 wizard (VB6 oracle DEV-9/df2, the runtime test
tests/db/f_dfw): one record with its fields, a grid, master / detail (MSDataShape), a hierarchical flex grid,
a chart; bound to an ADO Data control, to an ADO recordset in code, or to a class that is a data source.
The profile of the wizard (*.rwp) is read and written here too.
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
import configparser
import io
import re

# the choices of the pages (their order is the order of the lists; the index is what the profile keeps)
DB_TYPES = ("Access", "Remote(ODBC)")
LAYOUTS = ("Single Record", "Grid (Datasheet)", "Master/Detail", "MS HFlexGrid", "MS Chart")
L_SINGLE, L_GRID, L_MD, L_FLEX, L_CHART = range(5)
BINDINGS = ("ADO Data Control", "ADO Code", "Class")
B_ADODC, B_CODE, B_CLASS = range(3)
GRID_STYLES = ("Classic", "Professional", "Contemporary", "Win32 Application")
AGGREGATES = (("Average", "AVG"), ("Count", "COUNT"), ("Maximum", "MAX"), ("Minimum", "MIN"), ("Sum", "SUM"))
# the styles of the chart: (name, VtChChartType)
CHART_STYLES = (("2D Area", 5), ("2D Bar", 1), ("2D Combination", 9), ("2D Line", 3), ("2D Pie", 14), ("2D Step", 7),
                ("3D Area", 4), ("3D Bar", 0), ("3D Combination", 8), ("3D Line", 2), ("3D Step", 6))
# the captions of the chart: (name, how the code sets it)
CHART_CAPTIONS = (("Title", ".TitleText"), ("Footnote", ".FootnoteText"), ("X Axis Title", ".Plot.Axis(0).AxisTitle.Text"),
                  ("Y Axis Title", ".Plot.Axis(1).AxisTitle.Text"), ("Second Y Axis Title", ".Plot.Axis(2).AxisTitle.Text"),
                  ("Z Axis Title", ".Plot.Axis(3).AxisTitle.Text"))
PROFILE_SECTION = "VB 6 Data Form Wizard"

ADO_REF = "*\\G{00000205-0000-0010-8000-00AA006D2EA4}#2.5#0#msado25.tlb#Microsoft ActiveX Data Objects 2.5 Library"
DATASRC_REF = "*\\G{7C0FFAB0-CD84-11D0-949A-00A0C91110ED}#1.0#0#msdatsrc.tlb#Microsoft Data Source Interfaces"
OBJ_ADODC = '{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX'
OBJ_GRID = '{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX'
OBJ_HFLEX = '{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX'
OBJ_CHART = '{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX'
# the library keys of ocxlib for the Object= lines of the project
LIB_ADODC, LIB_GRID, LIB_HFLEX, LIB_CHART = "MSAdodcLib", "MSDataGridLib", "MSHierarchicalFlexGridLib", "MSChart20Lib"

TW_ROW = 340            # one field (twips)
TW_LABEL_W = 1815
TW_TEXT_L = 2040
TW_TEXT_W = 3375
TW_BTN_W = 1095
TW_BTN_H = 300
TW_FORM_W = 5640


class DataForm:
    """what the pages of the wizard decided"""
    def __init__(self):
        self.dbtype = 0                         # DB_TYPES
        self.database = ""                      # Access: the .mdb file
        self.odbc = {"DSN": "", "UID": "", "PWD": "", "Database": "", "Driver": "", "Server": ""}
        self.name = ""                          # of the form
        self.layout = L_SINGLE
        self.binding = B_ADODC
        self.source = ""                        # the record source (table or query)
        self.fields = []                        # the fields chosen, in their order
        self.sort = ""                          # the column to sort by ("" none)
        # Master/Detail: source / fields / sort above are the master's
        self.detail_source = ""
        self.detail_fields = []
        self.detail_sort = ""
        self.master_key = ""                    # the fields that link the two sources
        self.detail_key = ""
        # the controls (Control Selection)
        self.buttons = {"Add": True, "Update": True, "Delete": True, "Refresh": True, "Close": True}
        self.show_datacontrol = True
        # MS HFlexGrid
        self.grid_outline = False
        self.grid_style = 0                     # GRID_STYLES
        self.col_widths = {}                    # field -> twips (0: hidden)
        self.allow_dragging = True
        self.allow_sorting = True
        # MS Chart
        self.chart_x = ""
        self.chart_z = ""
        self.chart_y = ""
        self.chart_agg = 0                      # AGGREGATES
        self.chart_style = 7                    # CHART_STYLES (3D Bar)
        self.show_legend = False
        self.show_markers = False
        self.stack_series = False
        self.rotation = 60
        self.captions = {}                      # CHART_CAPTIONS name -> text
        self.class_override = ""                # the name of the class when the project has its usual one

    # ------------------------------------------------------------ what a page may allow
    @staticmethod
    def bindings_allowed(layout):
        """the binding types a layout can have (VB6: MS Chart only ADO Code, MS HFlexGrid no Class)"""
        if layout == L_CHART:
            return (B_CODE,)
        if layout == L_FLEX:
            return (B_ADODC, B_CODE)
        return (B_ADODC, B_CODE, B_CLASS)

    def class_name(self):
        return self.class_override or "cls" + ident(self.source or self.name)


# ---------------------------------------------------------------------------------------------- text helpers
def ident(s):
    """a name for VB from a table name ("Order Details" -> OrderDetails)"""
    t = re.sub(r"[^A-Za-z0-9_]", "", s or "")
    if not t:
        return "Data"
    return t if t[0].isalpha() else "T" + t


def vbstr(s):
    return '"' + (s or "").replace('"', '""') + '"'


def br(name):
    """a name in SQL: [Year Born]"""
    return "[" + name.replace("]", "]]") + "]"


def select_sql(source, fields, sort="", extra=()):
    cols = ", ".join(br(f) for f in list(fields) + [f for f in extra if f not in fields]) or "*"
    sql = "SELECT %s FROM %s" % (cols, br(source))
    if sort:
        sql += " ORDER BY " + br(sort)
    return sql


def connect_string(df, shape=False):
    """the ADO connection string of the database of the wizard"""
    if df.dbtype == 0:
        base = "Data Source=%s;" % df.database
        provider = "Microsoft.Jet.OLEDB.4.0"
    else:
        parts = []
        for k in ("DSN", "UID", "PWD", "Database", "Server"):
            v = df.odbc.get(k, "")
            if v:
                parts.append("%s=%s;" % (k, v))
        if df.odbc.get("Driver"):
            parts.append("Driver={%s};" % df.odbc["Driver"])
        base = "".join(parts)
        provider = "MSDASQL"
    if shape:
        return "Provider=MSDataShape;Data Provider=%s;%s" % (provider, base)
    return "Provider=%s;%s" % (provider, base)


def shape_sql(df):
    """Master/Detail: the master's rows with a chapter of the detail's (MSDataShape)"""
    master = select_sql(df.source, df.fields, df.sort, extra=(df.master_key,))
    detail = select_sql(df.detail_source, df.detail_fields, df.detail_sort, extra=(df.detail_key,))
    return "SHAPE {%s} APPEND ({%s} AS Detail RELATE %s TO %s)" % (master, detail, br(df.master_key), br(df.detail_key))


def chart_sql(df):
    agg = AGGREGATES[df.chart_agg][1]
    keys = [df.chart_x] + ([df.chart_z] if df.chart_z else [])
    k = ", ".join(br(f) for f in keys)
    return "SELECT %s, %s(%s) AS ChartValue FROM %s GROUP BY %s ORDER BY %s" % (
        k, agg, br(df.chart_y), br(df.source), k, k)


def himetric(tw):
    return int(round(tw * 2540 / 1440))


def label_of(field):
    return field + ":"


# ---------------------------------------------------------------------------------------------- the .frm
class Frm:
    """the text of a form, control by control"""
    def __init__(self):
        self.objects = []
        self.lines = []
        self.depth = 1

    def obj(self, line):
        if line not in self.objects:
            self.objects.append(line)

    def begin(self, cls, name, props):
        ind = "   " * self.depth
        self.lines.append("%sBegin %s %s" % (ind, cls, name))
        for k, v in props:
            self.lines.append("%s   %s=   %s" % (ind, k.ljust(15), v))
        self.depth += 1

    def end(self):
        self.depth -= 1
        self.lines.append("   " * self.depth + "End")

    def ctl(self, cls, name, props):
        self.begin(cls, name, props)
        self.end()


def _ocx_size(w, h):
    return [("_ExtentX", himetric(w)), ("_ExtentY", himetric(h)), ("_Version", 393216)]


def _form_text(df, frm, client_w, client_h, code, sizable=True):
    head = ["VERSION 5.00"]
    for o in frm.objects:
        head.append('Object = "%s"' % o)
    props = [("Caption", vbstr(df.source)),
             ("ClientHeight", client_h), ("ClientLeft", 60), ("ClientTop", 345), ("ClientWidth", client_w),
             ("KeyPreview", "-1  'True"), ("LinkTopic", vbstr(df.name)),
             ("ScaleHeight", client_h), ("ScaleWidth", client_w), ("StartUpPosition", "3  'Windows Default")]
    if not sizable:
        props[0:0] = [("BorderStyle", "3  'Fixed Dialog")]
        props += [("MaxButton", "0   'False"), ("MinButton", "0   'False")]
    body = ["Begin VB.Form %s" % df.name]
    for k, v in props:
        body.append("   %s=   %s" % (k.ljust(15), v))
    body += frm.lines
    body.append("End")
    attrs = ['Attribute VB_Name = "%s"' % df.name, "Attribute VB_GlobalNameSpace = False",
             "Attribute VB_Creatable = False", "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False"]
    return "\n".join(head + body + attrs) + "\n" + code


def _fields_block(frm, fields, top, datasource=""):
    """the labels and text boxes of the fields (control arrays lblField / txtField)"""
    for i, f in enumerate(fields):
        y = top + i * TW_ROW
        frm.ctl("VB.Label", "lblField", [("Caption", vbstr(label_of(f))), ("Height", 255), ("Index", i),
                                         ("Left", 120), ("TabIndex", 100 + i), ("Top", y + 30), ("Width", TW_LABEL_W)])
    for i, f in enumerate(fields):
        y = top + i * TW_ROW
        p = [("DataField", vbstr(f))]
        if datasource:
            p.append(("DataSource", vbstr(datasource)))
        p += [("Height", 285), ("Index", i), ("Left", TW_TEXT_L), ("TabIndex", i), ("Top", y), ("Width", TW_TEXT_W)]
        frm.ctl("VB.TextBox", "txtField", p)
    return top + len(fields) * TW_ROW


def _buttons(df):
    return [b for b in ("Add", "Update", "Delete", "Refresh", "Close") if df.buttons.get(b)]


def _buttons_block(frm, df, width, top=0):
    """the buttons of the form on a picture box at the bottom (picButtons)"""
    names = _buttons(df)
    if not names:
        return 0
    frm.begin("VB.PictureBox", "picButtons", [("Align", "2  'Align Bottom"), ("Appearance", "0  'Flat"),
                                              ("BorderStyle", "0  'None"), ("Height", TW_BTN_H + 60), ("Left", 0),
                                              ("ScaleHeight", TW_BTN_H + 60), ("ScaleWidth", width), ("TabIndex", 50),
                                              ("TabStop", "0   'False"), ("Top", top), ("Width", width)])
    for i, b in enumerate(names):
        p = [("Caption", vbstr("&" + b)), ("Height", TW_BTN_H), ("Left", 60 + i * (TW_BTN_W + 20)), ("TabIndex", 51 + i),
             ("Top", 30), ("Width", TW_BTN_W)]
        if b == "Close":
            p.insert(0, ("Cancel", "-1  'True"))
        frm.ctl("VB.CommandButton", "cmd" + b, p)
    frm.end()
    return TW_BTN_H + 60


def _nav_block(frm, width, top=0):
    """the moves through the records when there is no ADO Data control: |< < Record n of m > >|"""
    frm.begin("VB.PictureBox", "picNavigation", [("Align", "2  'Align Bottom"), ("Appearance", "0  'Flat"),
                                                 ("BorderStyle", "0  'None"), ("Height", 330), ("Left", 0),
                                                 ("ScaleHeight", 330), ("ScaleWidth", width), ("TabIndex", 60),
                                                 ("TabStop", "0   'False"), ("Top", top), ("Width", width)])
    for i, (n, cap) in enumerate((("cmdFirst", "|<"), ("cmdPrevious", "<"))):
        frm.ctl("VB.CommandButton", n, [("Caption", vbstr(cap)), ("Height", 300), ("Left", 60 + i * 360),
                                        ("TabIndex", 61 + i), ("Top", 15), ("Width", 345)])
    frm.ctl("VB.Label", "lblStatus", [("Alignment", "2  'Center"), ("BorderStyle", "1  'Fixed Single"),
                                      ("Height", 300), ("Left", 780), ("TabIndex", 63), ("Top", 15), ("Width", width - 1560)])
    for i, (n, cap) in enumerate((("cmdNext", ">"), ("cmdLast", ">|"))):
        frm.ctl("VB.CommandButton", n, [("Caption", vbstr(cap)), ("Height", 300), ("Left", width - 765 + i * 360),
                                        ("TabIndex", 64 + i), ("Top", 15), ("Width", 345)])
    frm.end()
    return 330


def _adodc_block(frm, df, width, shape=False, source_sql="", top=0):
    frm.obj(OBJ_ADODC)
    cs = connect_string(df, shape)
    p = [("Align", "2  'Align Bottom"), ("Height", 330), ("Left", 0), ("Top", top), ("Width", width)]
    p += _ocx_size(width, 330)[:2]
    p += [("ConnectMode", 0), ("CursorLocation", 3), ("IsolationLevel", -1), ("ConnectionTimeout", 15),
          ("CommandTimeout", 30), ("CursorType", 3), ("LockType", 3), ("CommandType", 1), ("CursorOptions", 0),
          ("CacheSize", 50), ("MaxRecords", 0), ("BOFAction", 0), ("EOFAction", 0), ("ConnectStringType", 1),
          ("Appearance", 1), ("BackColor", -2147483643), ("ForeColor", -2147483640), ("Orientation", 0),
          ("Enabled", -1)]
    inline = len(cs) < 800 and len(source_sql) < 800
    if inline:
        p += [("Connect", vbstr(cs)), ("OLEDBString", vbstr(cs))]
    p += [("OLEDBFile", '""'), ("DataSourceName", '""'), ("OtherAttributes", '""'), ("UserName", '""'),
          ("Password", '""'), ("RecordSource", vbstr(source_sql) if inline else '""'), ("Caption", '" "')]
    if not df.show_datacontrol and df.layout in (L_SINGLE, L_GRID, L_MD):
        p.insert(0, ("Visible", "0   'False"))
    p.append(("_Version", 393216))
    frm.ctl("MSAdodcLib.Adodc", "datPrimary", p)
    return 330 if df.show_datacontrol else 0, inline


# ---------------------------------------------------------------------------------------------- the code
HEADER = """Option Explicit
' {name}: a data form made by the Data Form Wizard of LinBASIC Studio
' {what}
"""


def _header(df):
    if df.layout == L_MD:
        what = "%s with its %s (%s -> %s)" % (df.source, df.detail_source, df.master_key, df.detail_key)
    else:
        what = "the records of " + df.source
    return HEADER.format(name=df.name, what=what)


def _code_buttons(df, rs, refresh_line):
    """the Click of the buttons of picButtons; rs: how the code reaches the recordset"""
    out = []
    b = df.buttons
    if b.get("Add"):
        out.append("""
Private Sub cmdAdd_Click()
    On Error GoTo Failed
    {rs}.AddNew
    Exit Sub
Failed:
    ShowError
End Sub
""".format(rs=rs))
    if b.get("Update"):
        out.append("""
Private Sub cmdUpdate_Click()
    On Error GoTo Failed
    If {rs}.EditMode <> adEditNone Then {rs}.Update
    Exit Sub
Failed:
    ShowError
End Sub
""".format(rs=rs))
    if b.get("Delete"):
        out.append("""
Private Sub cmdDelete_Click()
    On Error GoTo Failed
    If {rs}.BOF Or {rs}.EOF Then Exit Sub
    If MsgBox("Delete this record?", vbQuestion + vbYesNo, Caption) <> vbYes Then Exit Sub
    {rs}.Delete
    {rs}.MoveNext
    If {rs}.EOF And Not {rs}.BOF Then {rs}.MoveLast
    Exit Sub
Failed:
    ShowError
End Sub
""".format(rs=rs))
    if b.get("Refresh"):
        out.append("""
Private Sub cmdRefresh_Click()
    On Error GoTo Failed
    {line}
    Exit Sub
Failed:
    ShowError
End Sub
""".format(line=refresh_line))
    if b.get("Close"):
        out.append("""
Private Sub cmdClose_Click()
    Unload Me
End Sub
""")
    return "".join(out)


SHOW_ERROR = """
Private Sub ShowError()
    MsgBox Err.Description, vbExclamation, Caption
End Sub
"""


def _code_nav(rs):
    """the moves of picNavigation"""
    return """
Private Sub cmdFirst_Click()
    On Error GoTo Failed
    If Not ({rs}.BOF And {rs}.EOF) Then {rs}.MoveFirst
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo Failed
    If {rs}.BOF And {rs}.EOF Then Exit Sub
    {rs}.MovePrevious
    If {rs}.BOF Then {rs}.MoveFirst
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdNext_Click()
    On Error GoTo Failed
    If {rs}.BOF And {rs}.EOF Then Exit Sub
    {rs}.MoveNext
    If {rs}.EOF Then {rs}.MoveLast
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdLast_Click()
    On Error GoTo Failed
    If Not ({rs}.BOF And {rs}.EOF) Then {rs}.MoveLast
    Exit Sub
Failed:
    ShowError
End Sub
""".format(rs=rs)


def _status_text(pos, count):
    return 'IIf({p} > 0, "Record " & {p} & " of " & {c}, "")'.format(p=pos, c=count)


def _resize_fields(df, has_grid=False, grid="grdData", grid_top="0", bottom=0):
    """Form_Resize: the text boxes as wide as the form, the grid in the rest (above the bars at the bottom)"""
    lines = ["", "Private Sub Form_Resize()", "    Dim t As TextBox", "    On Error Resume Next",
             "    If WindowState = vbMinimized Then Exit Sub"]
    if df.fields and df.layout in (L_SINGLE, L_MD):
        lines += ["    For Each t In txtField", "        t.Width = ScaleWidth - t.Left - 120", "    Next"]
    if has_grid:
        lines += ["    %s.Move 120, %s, ScaleWidth - 240, ScaleHeight - %s - %d" % (grid, grid_top, grid_top, bottom + 120)]
    lines += ["End Sub", ""]
    return "\n".join(lines)


# ---------------------------------------------------------------------------------------------- the layouts
def _bottom_bars(frm, df, w, content_h, shape, sql):
    """the bars aligned at the bottom: the buttons above the moves (the ADO Data control or picNavigation).
    The runtime and VB6 stack them by their places in the .frm: the lower Top is the higher bar"""
    btn_h = TW_BTN_H + 60 if _buttons(df) else 0
    low_h = 330 if (df.binding != B_ADODC or df.show_datacontrol) else 0
    client_h = content_h + btn_h + low_h
    _buttons_block(frm, df, w, client_h - low_h - btn_h)
    if df.binding == B_ADODC:
        _h, inline = _adodc_block(frm, df, w, shape, sql, client_h - 330)
    else:
        inline = True
        _nav_block(frm, w, client_h - 330)
    return btn_h + low_h, inline, client_h


def _gen_single(df):
    frm = Frm()
    w = TW_FORM_W
    top = 120
    ds = "datPrimary" if df.binding == B_ADODC else ""
    sql = select_sql(df.source, df.fields, df.sort)
    bottom, inline, client_h = _bottom_bars(frm, df, w, top + len(df.fields) * TW_ROW + 60, False, sql)
    _fields_block(frm, df.fields, top, ds)
    code = _header(df)
    if df.binding == B_ADODC:
        code += _code_adodc_common(df, sql, inline)
        code += _code_buttons(df, "datPrimary.Recordset", "datPrimary.Refresh")
    elif df.binding == B_CODE:
        code += _code_ado(df, sql, binds_fields=True)
        code += _code_buttons(df, "rs", "rs.Requery")
        code += _code_nav("rs")
    else:
        code += _code_class_form(df)
        code += _code_buttons(df, "src", "src.Requery")
        code += _code_nav("src")
    code += _resize_fields(df)
    code += SHOW_ERROR
    return _form_text(df, frm, w, client_h, code)


def _gen_grid(df):
    frm = Frm()
    w, gh = 6480, 3600
    frm.obj(OBJ_GRID)
    sql = select_sql(df.source, df.fields, df.sort)
    bottom, inline, client_h = _bottom_bars(frm, df, w, 120 + gh + 120, False, sql)
    frm.ctl("MSDataGridLib.DataGrid", "grdData", [("Height", gh), ("Left", 120), ("TabIndex", 0), ("Top", 120),
                                                  ("Width", w - 240)] + _ocx_size(w - 240, gh) +
            [("HeadLines", 1), ("RowHeight", 15)])
    code = _header(df)
    if df.binding == B_ADODC:
        code += _code_adodc_common(df, sql, inline, grid_line="    Set grdData.DataSource = datPrimary")
        code += _code_buttons(df, "datPrimary.Recordset", "datPrimary.Refresh")
    elif df.binding == B_CODE:
        code += _code_ado(df, sql, binds_fields=False, grid_line="    Set grdData.DataSource = rs")
        code += _code_buttons(df, "rs", "rs.Requery")
        code += _code_nav("rs")
    else:
        code += _code_class_form(df, grid_lines=['    grdData.DataMember = "Primary"',
                                                 "    Set grdData.DataSource = src"])
        code += _code_buttons(df, "src", "src.Requery")
        code += _code_nav("src")
    code += _resize_fields(df, True, "grdData", "120", bottom)
    code += SHOW_ERROR
    return _form_text(df, frm, w, client_h, code)


def _gen_md(df):
    frm = Frm()
    w, gh = 6480, 2400
    frm.obj(OBJ_GRID)
    sql = shape_sql(df)
    gtop = 120 + len(df.fields) * TW_ROW + 60
    bottom, inline, client_h = _bottom_bars(frm, df, w, gtop + gh + 120, True, sql)
    ds = "datPrimary" if df.binding == B_ADODC else ""
    _fields_block(frm, df.fields, 120, ds)
    frm.ctl("MSDataGridLib.DataGrid", "grdDetail", [("Height", gh), ("Left", 120), ("TabIndex", 40), ("Top", gtop),
                                                    ("Width", w - 240)] + _ocx_size(w - 240, gh) +
            [("AllowUpdate", "0   'False"), ("HeadLines", 1), ("RowHeight", 15)])
    code = _header(df)
    code += "Private Const GRID_TOP As Single = %d\n" % gtop
    detail = {B_ADODC: 'datPrimary.Recordset("Detail").UnderlyingValue', B_CODE: 'rs("Detail").UnderlyingValue',
              B_CLASS: "src.Detail"}[df.binding]
    if df.binding == B_ADODC:
        code += _code_adodc_common(df, sql, inline, shape=True, grid_line="    BindDetail")
        code += _code_buttons(df, "datPrimary.Recordset", "datPrimary.Refresh: BindDetail")
    elif df.binding == B_CODE:
        code += _code_ado(df, sql, binds_fields=True, shape=True, grid_line="    BindDetail")
        code += _code_buttons(df, "rs", "rs.Requery: BindDetail")
        code += _code_nav("rs")
    else:
        code += _code_class_form(df, grid_lines=["    BindDetail"])
        code += _code_buttons(df, "src", "src.Requery: BindDetail")
        code += _code_nav("src")
    code += """
' the grid shows the records of {d} that go with the current record (the chapter "Detail")
Private Sub BindDetail()
    On Error Resume Next
    Set grdDetail.DataSource = {detail}
End Sub
""".format(d=df.detail_source, detail=detail)
    code += _resize_fields(df, True, "grdDetail", "GRID_TOP", bottom)
    code += SHOW_ERROR
    return _form_text(df, frm, w, client_h, code)


# the looks of the hierarchical flex grid: (BackColor, ForeColor, BackColorFixed, ForeColorFixed, GridColor,
# BackColorBand / alternate rows)
FLEX_STYLES = ((0xFFFFFF, 0x000000, 0x000000, 0xFFFFFF, 0xC0C0C0),
               (0xFFFFFF, 0x000000, 0x800000, 0xFFFFFF, 0xC0C0C0),
               (0xFFFFFF, 0x000000, 0x808000, 0xFFFFFF, 0xC0C0C0),
               (-2147483643, -2147483640, -2147483633, -2147483630, -2147483632))


def _gen_flex(df):
    frm = Frm()
    w, gh = 6480, 3840
    frm.obj(OBJ_HFLEX)
    sql = select_sql(df.source, df.fields, df.sort)
    if df.binding == B_ADODC:
        _adodc_block(frm, df, w, False, sql)
        inline = len(connect_string(df)) < 800 and len(sql) < 800
    else:
        inline = True
    st = FLEX_STYLES[df.grid_style]
    p = [("Height", gh), ("Left", 60), ("TabIndex", 0), ("Top", 60), ("Width", w - 120)] + _ocx_size(w - 120, gh)[:2]
    p += [("BackColor", st[0]), ("ForeColor", st[1]), ("Cols", max(2, len(df.fields))), ("FixedCols", 0),
          ("BackColorFixed", st[2]), ("ForeColorFixed", st[3]), ("GridColor", st[4]), ("WordWrap", "-1  'True"),
          ("AllowBigSelection", "0   'False"), ("FocusRect", 0), ("HighLight", 0), ("AllowUserResizing", 1)]
    if df.grid_outline:
        p += [("MergeCells", 2)]
    p += [("_Version", 393216)]
    frm.ctl("MSHierarchicalFlexGridLib.MSHFlexGrid", "flxData", p)
    frm.ctl("VB.CommandButton", "cmdClose", [("Cancel", "-1  'True"), ("Caption", vbstr("&Close")),
                                             ("Height", TW_BTN_H), ("Left", w - 60 - TW_BTN_W), ("TabIndex", 1),
                                             ("Top", gh + 120), ("Width", TW_BTN_W)])
    client_h = gh + 120 + TW_BTN_H + 60
    code = _header(df)
    code += "Private mSortCol As Long\nPrivate mSortDescending As Boolean\nPrivate mDragCol As Long\n"
    if df.binding == B_ADODC:
        code += "\nPrivate Sub Form_Load()\n    On Error GoTo Failed\n    datPrimary.Visible = False\n"
        if not inline:
            code += "    datPrimary.ConnectionString = %s\n    datPrimary.RecordSource = %s\n    datPrimary.Refresh\n" % (
                vbstr(connect_string(df)), vbstr(sql))
        code += "    Set flxData.DataSource = datPrimary\n    SetUpGrid\n    Exit Sub\nFailed:\n    ShowError\nEnd Sub\n"
    else:
        code += """
Private cn As ADODB.Connection
Private rs As ADODB.Recordset

Private Sub Form_Load()
    On Error GoTo Failed
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open {cs}
    Set rs = New ADODB.Recordset
    rs.Open {sql}, cn, adOpenStatic, adLockReadOnly
    Set flxData.DataSource = rs
    SetUpGrid
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rs.Close
    cn.Close
End Sub
""".format(cs=vbstr(connect_string(df)), sql=vbstr(sql))
    lines = ["", "' the widths of the columns, the bold headers and the bands of the style", "Private Sub SetUpGrid()",
             "    Dim c As Long, r As Long", "    With flxData", "        .Redraw = False"]
    for i, f in enumerate(df.fields):
        wdt = df.col_widths.get(f)
        if wdt is not None:
            lines.append("        .ColWidth(%d) = %d" % (i, wdt))
    if df.grid_outline:
        lines += ["        For c = 0 To .Cols - 1", "            .MergeCol(c) = True", "        Next"]
    lines += ["        .Row = 0", "        For c = 0 To .Cols - 1", "            .Col = c", "            .CellFontBold = True",
              "        Next"]
    if df.grid_style == 1:                      # Professional: every other row gray
        lines += ["        .FillStyle = 1", "        For r = .FixedRows To .Rows - 1 Step 2",
                  "            .Row = r: .Col = 0: .RowSel = r: .ColSel = .Cols - 1",
                  "            .CellBackColor = &HE0E0E0", "        Next", "        .FillStyle = 0"]
    elif df.grid_style == 2:                    # Contemporary: every other column gray
        lines += ["        If .Rows > .FixedRows Then", "            .FillStyle = 1",
                  "            For c = 1 To .Cols - 1 Step 2",
                  "                .Row = .FixedRows: .Col = c: .RowSel = .Rows - 1: .ColSel = c",
                  "                .CellBackColor = &HE0E0E0", "            Next", "            .FillStyle = 0", "        End If"]
    lines += ["        .Row = 0",
              "        .Col = 0                    ' back to the first column (the grid shows the current one)",
              "        .Redraw = True", "    End With", "End Sub", ""]
    code += "\n".join(lines)
    if df.allow_sorting:
        code += """
' a double click on a header sorts by its column (again: the other way)
Private Sub flxData_DblClick()
    With flxData
        If .MouseRow >= .FixedRows Then Exit Sub
        If .MouseCol = mSortCol Then mSortDescending = Not mSortDescending Else mSortDescending = False
        mSortCol = .MouseCol
        .Redraw = False
        .Col = mSortCol
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .Sort = IIf(mSortDescending, 2, 1)      ' generic descending / ascending
        .Redraw = True
    End With
End Sub
"""
    if df.allow_dragging:
        code += """
' a header dragged onto another one moves its column there
Private Sub flxData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mDragCol = -1
    If Button = vbLeftButton And flxData.MouseRow = 0 Then mDragCol = flxData.MouseCol
End Sub

Private Sub flxData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mDragCol < 0 Then Exit Sub
    If flxData.MouseRow = 0 And flxData.MouseCol <> mDragCol Then
        flxData.Redraw = False
        flxData.ColPosition(mDragCol) = flxData.MouseCol
        flxData.Redraw = True
    End If
    mDragCol = -1
End Sub
"""
    code += "\nPrivate Sub cmdClose_Click()\n    Unload Me\nEnd Sub\n"
    code += """
Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = vbMinimized Then Exit Sub
    cmdClose.Move ScaleWidth - cmdClose.Width - 60, ScaleHeight - cmdClose.Height - 60
    flxData.Move 60, 60, ScaleWidth - 120, ScaleHeight - cmdClose.Height - 180
End Sub
"""
    code += SHOW_ERROR
    return _form_text(df, frm, w, client_h, code)


def _gen_chart(df):
    frm = Frm()
    w, gh = 6480, 3840
    frm.obj(OBJ_CHART)
    # the chart on the form, as VB6's wizard puts it: VB6 keeps its state in an OleObjectBlob of the .frx, which the
    # project writes when it saves the form (vbmodel.sync_ole_blobs, lbvc/project/chartblob.py)
    frm.ctl("MSChart20Lib.MSChart", "chtData", [("Height", gh), ("Left", 60), ("TabIndex", 0), ("Top", 60),
                                               ("Width", w - 120)])
    frm.ctl("VB.CommandButton", "cmdClose", [("Cancel", "-1  'True"), ("Caption", vbstr("&Close")),
                                             ("Height", TW_BTN_H), ("Left", w - 60 - TW_BTN_W), ("TabIndex", 1),
                                             ("Top", gh + 120), ("Width", TW_BTN_W)])
    client_h = gh + 120 + TW_BTN_H + 60
    ctype = CHART_STYLES[df.chart_style][1]
    sets = []
    for name, how in CHART_CAPTIONS:
        t = df.captions.get(name, "")
        if t:
            sets.append("        %s = %s" % (how, vbstr(t)))
    code = _header(df)
    code += """
Private Const CONNECT_STRING As String = {cs}
Private Const CHART_SQL As String = {sql}
Private Const HAS_Z As Boolean = {hasz}

Private Sub Form_Load()
    Dim cn As ADODB.Connection, rs As ADODB.Recordset
    On Error GoTo Failed
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open CONNECT_STRING
    Set rs = New ADODB.Recordset
    rs.Open CHART_SQL, cn, adOpenStatic, adLockReadOnly
    FillChart rs
    rs.Close
    cn.Close
    With chtData
        .ChartType = {ctype}
        .ShowLegend = {legend}
        .Stacking = {stack}
        .Plot.View3d.Rotation = {rot}
{sets}
    End With
    ShowMarkers {markers}
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, Caption
End Sub

' the rows of the query in the chart: one row for each value of the X field, one series (column) for each value
' of the Z field (without a Z field: one series)
Private Sub FillChart(rs As ADODB.Recordset)
    Dim xs As New Collection, zs As New Collection, r As Long, c As Long, k As String
    Do Until rs.EOF
        k = KeyOf(rs.Fields(0).Value)
        If Not InCol(xs, k) Then xs.Add k, "k" & k
        If HAS_Z Then
            k = KeyOf(rs.Fields(1).Value)
            If Not InCol(zs, k) Then zs.Add k, "k" & k
        End If
        rs.MoveNext
    Loop
    If xs.Count = 0 Then Exit Sub
    With chtData
        .RowCount = xs.Count
        .ColumnCount = IIf(HAS_Z, zs.Count, 1)
        For r = 1 To xs.Count
            .Row = r
            .RowLabel = xs(r)
            For c = 1 To .ColumnCount
                .Column = c
                .Data = 0
            Next
        Next
        If HAS_Z Then
            For c = 1 To zs.Count
                .Column = c
                .ColumnLabel = zs(c)
            Next
        Else
            .Column = 1
            .ColumnLabel = {ylabel}
        End If
        rs.MoveFirst
        Do Until rs.EOF
            .Row = PosOf(xs, KeyOf(rs.Fields(0).Value))
            .Column = IIf(HAS_Z, PosOf(zs, KeyOf(rs.Fields(1).Value)), 1)
            If Not IsNull(rs!ChartValue) Then .Data = rs!ChartValue
            rs.MoveNext
        Loop
    End With
End Sub

Private Function KeyOf(v As Variant) As String
    If IsNull(v) Then KeyOf = "(Null)" Else KeyOf = CStr(v)
End Function

Private Function InCol(col As Collection, k As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = col("k" & k)
    InCol = (Err.Number = 0)
    Err.Clear
End Function

Private Function PosOf(col As Collection, k As String) As Long
    Dim i As Long
    For i = 1 To col.Count
        If col(i) = k Then PosOf = i: Exit Function
    Next
End Function

Private Sub ShowMarkers(ByVal show As Boolean)
    Dim i As Integer
    On Error Resume Next
    For i = 1 To chtData.Plot.SeriesCollection.Count
        chtData.Plot.SeriesCollection(i).SeriesMarker.Show = show
    Next
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = vbMinimized Then Exit Sub
    cmdClose.Move ScaleWidth - cmdClose.Width - 60, ScaleHeight - cmdClose.Height - 60
    chtData.Move 60, 60, ScaleWidth - 120, ScaleHeight - cmdClose.Height - 180
End Sub
""".format(cs=vbstr(connect_string(df)), sql=vbstr(chart_sql(df)), hasz="True" if df.chart_z else "False",
           ctype=ctype, legend="True" if df.show_legend else "False", stack="True" if df.stack_series else "False",
           rot=int(df.rotation), sets="\n".join(sets), markers="True" if df.show_markers else "False",
           ylabel=vbstr(df.chart_y))
    return _form_text(df, frm, w, client_h, code)


# ---------------------------------------------------------------------------------------------- the bindings
def _code_adodc_common(df, sql, inline, shape=False, grid_line=""):
    """the form of an ADO Data control: its caption shows where the record is"""
    out = "\nPrivate Sub Form_Load()\n    On Error GoTo Failed\n"
    if not inline:
        out += "    datPrimary.ConnectionString = %s\n    datPrimary.RecordSource = %s\n    datPrimary.Refresh\n" % (
            vbstr(connect_string(df, shape)), vbstr(sql))
    if grid_line:
        out += grid_line + "\n"
    out += "    Exit Sub\nFailed:\n    ShowError\nEnd Sub\n"
    out += """
Private Sub datPrimary_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    datPrimary.Caption = {status}
End Sub

Private Sub datPrimary_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    MsgBox Description, vbExclamation, Caption
    fCancelDisplay = True
End Sub
""".format(status=_status_text("pRecordset.AbsolutePosition", "pRecordset.RecordCount"))
    return out


def _code_ado(df, sql, binds_fields, shape=False, grid_line=""):
    """the form of an ADO recordset in code"""
    out = """
Private cn As ADODB.Connection
Private WithEvents rs As ADODB.Recordset

Private Sub Form_Load()
    Dim t As TextBox
    On Error GoTo Failed
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open {cs}
    Set rs = New ADODB.Recordset
    rs.Open {sql}, cn, adOpenStatic, adLockOptimistic
""".format(cs=vbstr(connect_string(df, shape)), sql=vbstr(sql))
    if binds_fields and df.fields:
        out += "    For Each t In txtField\n        Set t.DataSource = rs\n    Next\n"
    if grid_line:
        out += grid_line + "\n"
    out += """    Exit Sub
Failed:
    ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If rs.EditMode <> adEditNone Then rs.CancelUpdate
    rs.Close
    cn.Close
End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    lblStatus.Caption = {status}
End Sub
""".format(status=_status_text("rs.AbsolutePosition", "rs.RecordCount"))
    return out


def _code_class_form(df, grid_lines=()):
    """the form of a data source class: its controls bound through the member "Primary" """
    out = """
Private WithEvents src As {cls}

Private Sub Form_Load()
    Dim t As TextBox
    On Error GoTo Failed
    Set src = New {cls}
""".format(cls=df.class_name())
    if df.fields and df.layout in (L_SINGLE, L_MD):
        out += '    For Each t In txtField\n        t.DataMember = "Primary"\n        Set t.DataSource = src\n    Next\n'
    for line in grid_lines:
        out += line + "\n"
    out += """    Exit Sub
Failed:
    ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set src = Nothing
End Sub

Private Sub src_MoveComplete()
    lblStatus.Caption = {status}
End Sub
""".format(status=_status_text("src.AbsolutePosition", "src.RecordCount"))
    return out


def class_text(df):
    """the data source class of the Class binding: the records as the member "Primary" """
    shape = df.layout == L_MD
    sql = shape_sql(df) if shape else select_sql(df.source, df.fields, df.sort)
    name = df.class_name()
    detail = ""
    if shape:
        detail = """
' Master/Detail: the records of {d} of the current record
Public Property Get Detail() As ADODB.Recordset
    Set Detail = rs("Detail").UnderlyingValue
End Property
""".format(d=df.detail_source)
    return """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "{name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' {name}: the records of {src} for the form {form} (Data Form Wizard of LinBASIC Studio).
' Bound controls: DataMember = "Primary", Set DataSource = an object of this class.

Public Event MoveComplete()

Private cn As ADODB.Connection
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open {cs}
    Set rs = New ADODB.Recordset
    rs.Open {sql}, cn, adOpenStatic, adLockOptimistic
    DataMembers.Add "Primary"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If rs.EditMode <> adEditNone Then rs.CancelUpdate
    rs.Close
    cn.Close
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    If DataMember = "Primary" Then Set Data = rs
End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    RaiseEvent MoveComplete
End Sub

Public Property Get AbsolutePosition() As Long
    AbsolutePosition = rs.AbsolutePosition
End Property

Public Property Get RecordCount() As Long
    RecordCount = rs.RecordCount
End Property

Public Property Get BOF() As Boolean
    BOF = rs.BOF
End Property

Public Property Get EOF() As Boolean
    EOF = rs.EOF
End Property

Public Property Get EditMode() As Long
    EditMode = rs.EditMode
End Property
{detail}
Public Sub AddNew()
    rs.AddNew
End Sub

Public Sub Update()
    rs.Update
End Sub

Public Sub Delete()
    rs.Delete
End Sub

' Refresh: the records again from the database; the bound controls read the member again
Public Sub Requery()
    rs.Requery
    DataMemberChanged "Primary"
End Sub

Public Sub MoveFirst()
    rs.MoveFirst
End Sub

Public Sub MovePrevious()
    rs.MovePrevious
End Sub

Public Sub MoveNext()
    rs.MoveNext
End Sub

Public Sub MoveLast()
    rs.MoveLast
End Sub
""".format(name=name, src=df.source, form=df.name, cs=vbstr(connect_string(df, shape)), sql=vbstr(sql), detail=detail)


# ---------------------------------------------------------------------------------------------- the result
class Output:
    """the modules of a data form and what the project needs for them"""
    def __init__(self):
        self.modules = []                       # (kind "Form" | "Class", name, text with \n)
        self.references = [ADO_REF]
        self.libraries = []                     # ocxlib keys


def generate(df):
    out = Output()
    gen = {L_SINGLE: _gen_single, L_GRID: _gen_grid, L_MD: _gen_md, L_FLEX: _gen_flex, L_CHART: _gen_chart}[df.layout]
    text = gen(df)
    if df.binding == B_CLASS:
        out.modules.append(("Class", df.class_name(), class_text(df)))
        out.references.append(DATASRC_REF)
    out.modules.append(("Form", df.name, text))
    if df.binding == B_ADODC:
        out.libraries.append(LIB_ADODC)
    if df.layout in (L_GRID, L_MD):
        out.libraries.append(LIB_GRID)
    if df.layout == L_FLEX:
        out.libraries.append(LIB_HFLEX)
    if df.layout == L_CHART:
        out.libraries.append(LIB_CHART)
    return out


def check(df):
    """what is missing for Finish: None or the message"""
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,39}", df.name or ""):
        return "The name of the form is not valid."
    if df.layout == L_CHART:
        if not (df.source and df.chart_x and df.chart_y):
            return "Select the fields of the chart."
        axes = [f for f in (df.chart_x, df.chart_z, df.chart_y) if f]
        if len(set(axes)) != len(axes):
            return "A field can be on one axis only. Select a different field for each axis."
        return None
    if not df.source or not df.fields:
        return "Select a record source and at least one field."
    if df.layout == L_MD and not (df.detail_source and df.detail_fields and df.master_key and df.detail_key):
        return "Select the detail record source, its fields and the fields that link the two sources."
    if df.binding not in DataForm.bindings_allowed(df.layout):
        return "This form layout cannot have this binding type."
    return None


# ---------------------------------------------------------------------------------------------- the profile
def save_profile(df, path):
    lines = ["[%s]" % PROFILE_SECTION,
             "DatabaseTypeIndex=%d" % df.dbtype,
             "DatabaseTypeName=%s" % ("Access" if df.dbtype == 0 else "Remote(ODBC)"),
             "DatabaseName=%s" % df.database]
    if df.dbtype == 1:
        for k in ("DSN", "UID", "Database", "Driver", "Server"):    # the password is not kept
            lines.append("ODBC%s=%s" % (k, df.odbc.get(k, "")))
    lines += ["BindType=%d" % df.binding, "FormLayout=%d" % df.layout,
              "AddBtn=%d" % int(df.buttons.get("Add", True)), "DelBtn=%d" % int(df.buttons.get("Delete", True)),
              "RefBtn=%d" % int(df.buttons.get("Refresh", True)), "UptBtn=%d" % int(df.buttons.get("Update", True)),
              "ClsBtn=%d" % int(df.buttons.get("Close", True)), "DataControl=%d" % int(df.show_datacontrol)]
    with open(path, "w", encoding="cp1252", errors="replace", newline="\r\n") as f:
        f.write("\n".join(lines) + "\n")


def load_profile(path, df=None):
    """the settings of a profile into df (a new DataForm without one)"""
    df = df or DataForm()
    cp = configparser.ConfigParser(interpolation=None, strict=False)
    cp.optionxform = str
    with open(path, encoding="cp1252", errors="replace") as f:
        cp.read_file(io.StringIO(f.read()))
    if not cp.has_section(PROFILE_SECTION):
        raise ValueError("not a profile of the Data Form Wizard")
    s = cp[PROFILE_SECTION]

    def num(k, d):
        try:
            return int(s.get(k, d))
        except ValueError:
            return d
    df.dbtype = max(0, min(1, num("DatabaseTypeIndex", 0)))
    df.database = s.get("DatabaseName", "")
    for k in ("DSN", "UID", "Database", "Driver", "Server"):
        df.odbc[k] = s.get("ODBC" + k, df.odbc.get(k, ""))
    df.binding = max(0, min(2, num("BindType", 0)))
    df.layout = max(0, min(4, num("FormLayout", 0)))
    for b, k in (("Add", "AddBtn"), ("Delete", "DelBtn"), ("Refresh", "RefBtn"), ("Update", "UptBtn"),
                 ("Close", "ClsBtn")):
        df.buttons[b] = num(k, 1) != 0
    df.show_datacontrol = num("DataControl", 1) != 0
    return df
