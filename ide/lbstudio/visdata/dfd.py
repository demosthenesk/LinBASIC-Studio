"""Data Form Designer of the Visual Data Manager (VB6 DFD.FRM): a form of the project with a Data control on a
RecordSource and a label and a bound control for each field chosen, with Add / Delete / Refresh / Update / Close.
VB6 puts an OLE container on a Binary field; OLE is not part of LinBASIC Studio, so a bound PictureBox."""
from PyQt6.QtWidgets import QDialog, QLineEdit, QListWidget, QPushButton, QComboBox

from ..rtobj import RtError
from . import core as C
from ..dfwizgen import vbstr

dbBoolean, dbLongBinary, dbText, dbMemo = 1, 11, 10, 12


def form_code():
    """BuildFrmCode: the events of the buttons and of the Data control (as VB6 saves them)"""
    return "\r\n".join([
        "Private Sub cmdAdd_Click()",
        "  Data1.Recordset.AddNew",
        "End Sub",
        "",
        "Private Sub cmdDelete_Click()",
        "  'this may produce an error if you delete the last",
        "  'record or the only record in the recordset",
        "  Data1.Recordset.Delete",
        "  Data1.Recordset.MoveNext",
        "End Sub",
        "",
        "Private Sub cmdRefresh_Click()",
        "  'this is really only needed for multi user apps",
        "  Data1.Refresh",
        "End Sub",
        "",
        "Private Sub cmdUpdate_Click()",
        "  Data1.UpdateRecord",
        "  Data1.Recordset.Bookmark = Data1.Recordset.LastModified",
        "End Sub",
        "",
        "Private Sub cmdClose_Click()",
        "  Unload Me",
        "End Sub",
        "",
        "Private Sub Data1_Error(DataErr As Integer, Response As Integer)",
        "  'This is where you would put error handling code",
        "  'If you want to ignore errors, comment out the next line",
        "  'If you want to trap them, add code here to handle them",
        "  MsgBox \"Data error event hit err:\" & Error$(DataErr)",
        "  Response = 0  'throw away the error",
        "End Sub",
        "",
        "Private Sub Data1_Reposition()",
        "  Screen.MousePointer = vbDefault",
        "  On Error Resume Next",
        "  'This will display the current record position",
        "  'for dynasets and snapshots",
        "  Data1.Caption = \"Record: \" & (Data1.Recordset.AbsolutePosition + 1)",
        "  'for the table object you must set the index property when",
        "  'the recordset gets created and use the following line",
        "  'Data1.Caption = \"Record: \" & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1",
        "End Sub",
        "",
        "Private Sub Data1_Validate(Action As Integer, Save As Integer)",
        "  'This is where you put validation code",
        "  'This event gets called when the following actions occur",
        "  Select Case Action",
        "    Case vbDataActionMoveFirst",
        "    Case vbDataActionMovePrevious",
        "    Case vbDataActionMoveNext",
        "    Case vbDataActionMoveLast",
        "    Case vbDataActionAddNew",
        "    Case vbDataActionUpdate",
        "    Case vbDataActionDelete",
        "    Case vbDataActionFind",
        "    Case vbDataActionBookmark",
        "    Case vbDataActionClose",
        "  End Select",
        "  Screen.MousePointer = vbHourglass",
        "End Sub",
        ""]) + "\r\n"


def form_text(name, caption, fields, dbname, connect, source):
    """the .frm of the form as VB6 saves what BuildForm made (VB6 oracle oracles/12.4/frmPubs.frm): fields =
    [(name, type, size)] in their order; the controls with a window first, the newest first, then the labels"""
    n = len(fields)
    height = 1115 + n * 320                     # (the Height BuildForm gives the form)
    width = 5640
    client_w, client_h = width - 240, height - 565
    windowed, labels = [], []

    def ctl(where, cls, cname, props):
        where.append((cls, cname, sorted(props, key=lambda kv: kv[0].lower())))

    top = 0
    for i, (fname, ftype, fsize) in enumerate(fields):
        ctl(labels, "VB.Label", "lblLabels", [("Caption", vbstr(fname + ":")), ("Height", 255), ("Index", i),
                                              ("Left", 120), ("TabIndex", 2 * i), ("Top", i * 320 + 60),
                                              ("Width", 1815)])
        bind = [("DataField", vbstr(fname)), ("DataSource", '"Data1"')]
        if ftype == dbBoolean:
            top = i * 320 + 40
            ctl(windowed, "VB.CheckBox", "chkFields", bind + [("Height", 285), ("Index", i), ("Left", 2040),
                                                               ("TabIndex", 2 * i + 1), ("Top", top), ("Width", 3375)])
        elif ftype == dbLongBinary:
            top = i * 320 + 40
            ctl(windowed, "VB.PictureBox", "picFields", bind + [("Height", 285), ("Index", i), ("Left", 2040),
                                                                 ("ScaleHeight", 225), ("ScaleWidth", 3315),
                                                                 ("TabIndex", 2 * i + 1), ("Top", top), ("Width", 3375)])
        else:
            props = bind + [("Index", i), ("Left", 2040), ("TabIndex", 2 * i + 1), ("Width", 1935 if ftype < 10 else 3375)]
            if ftype == dbText:
                top = i * 320 + 40
                props += [("Height", 285), ("MaxLength", fsize), ("Top", top)]
            elif ftype == dbMemo:
                top = i * 320 + 30
                props += [("Height", 310), ("MultiLine", "-1  'True"), ("ScrollBars", "2  'Vertical"), ("Top", top)]
            else:
                top = i * 320 + 40
                props += [("Height", 285), ("Top", top)]
            ctl(windowed, "VB.TextBox", "txtFields", props)
    btop = top + 340
    data = [("Align", "2  'Align Bottom"), ("Connect", vbstr(connect)), ("DefaultCursorType", "0  'DefaultCursor"),
            ("DefaultType", "2  'UseODBC"), ("Exclusive", "0   'False"), ("Height", 300), ("Left", 0),
            ("Options", 0), ("ReadOnly", "0   'False"), ("RecordsetType", "1  'Dynaset"),
            ("RecordSource", vbstr(source)), ("Top", client_h - 300), ("Width", client_w)]
    if dbname is not None:
        data.append(("DatabaseName", vbstr(dbname)))
    ctl(windowed, "VB.Data", "Data1", data)
    for k, (cname, cap) in enumerate((("cmdAdd", "&Add"), ("cmdDelete", "&Delete"), ("cmdRefresh", "&Refresh"),
                                      ("cmdUpdate", "&Update"), ("cmdClose", "&Close"))):
        ctl(windowed, "VB.CommandButton", cname, [("Caption", vbstr(cap)), ("Height", 300), ("Left", 120 + k * 1080),
                                                  ("TabIndex", 2 * n + k), ("Top", btop), ("Width", 975)])
    lines = ["VERSION 5.00", f"Begin VB.Form {name} ",
             f"   Caption         =   {vbstr(caption[:32])}",
             f"   ClientHeight    =   {client_h}", "   ClientLeft      =   1170", "   ClientTop       =   450",
             f"   ClientWidth     =   {client_w}", '   LinkTopic       =   "Form1"',
             f"   ScaleHeight     =   {client_h}", f"   ScaleWidth      =   {client_w}"]
    for cls, cname, props in list(reversed(windowed)) + list(reversed(labels)):
        lines.append(f"   Begin {cls} {cname} ")
        for k, v in props:
            lines.append(f"      {k.ljust(16)}=   {v}" if len(k) < 16 else f"      {k}=   {v}")
        lines.append("   End")
    lines.append("End")
    lines += [f'Attribute VB_Name = "{name}"', "Attribute VB_GlobalNameSpace = False",
              "Attribute VB_Creatable = False", "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False"]
    return "\r\n".join(lines) + "\r\n" + form_code()


def add_form(main, name, text):
    """the form into the active project (not saved, as VB6) and opened in its designer"""
    from ..vbmodel import Module, parse_source
    p = main.project
    sf = parse_source(text, None, "Form", True, p.encoding)
    sf.encoding = p.encoding
    m = Module(p, "Form", name, name + ".frm", sf)
    m.dirty = True
    p.modules.append(m)
    p._vbp_add(m)
    p.dirty = True
    main.fill_tree()
    main.update_title()
    main.open_designer(m)
    return m


class DataFormDesigner(QDialog):
    def __init__(self, vd):
        super().__init__(vd)
        self.vd = vd
        self.rs = None
        self.rs_source = ""
        self.setWindowTitle("Data Form Designer")
        P = C.Place(self, 462, 247)             # (VB6 oracle vb_dfd.png)
        self.name = QLineEdit()
        P.label("Form Name (w/o Extension): ", 8, 8, 140)
        P.put(self.name, 184, 8, 122, 19)
        self.source = QComboBox()
        self.source.setEditable(True)
        P.label("RecordSource: ", 7, 36, 74)
        P.put(self.source, 112, 32, 334, 21)
        P.label("Select a Table/QueryDef from the list or enter a SQL statement.", 112, 56, 334, center=True)
        P.line(8, 72, 452)
        self.all, self.sel = QListWidget(), QListWidget()
        P.label("Available Fields: ", 8, 80, 80)
        P.put(self.all, 8, 96, 179, 108)
        P.label("Included Fields: ", 234, 80, 78)
        P.put(self.sel, 234, 96, 179, 108)
        self.b_move = []
        for k, t in enumerate((">", ">>", "<<", "<")):
            b = P.put(QPushButton(t), 194, 93 + 30 * k, 33, 25)
            b.clicked.connect(lambda _c=False, i=k: self.move_fields(i))
            self.b_move.append(b)
        self.b_up, self.b_down = QPushButton("\u25b2"), QPushButton("\u25bc")
        for b in (self.b_up, self.b_down):         # (VB6: the black arrows of DFD.FRX)
            f = b.font()
            f.setPointSizeF(f.pointSizeF() * 1.8)
            b.setFont(f)
        P.put(self.b_up, 419, 114, 36, 36)
        P.put(self.b_down, 419, 153, 36, 36)
        self.b_build = P.put(QPushButton("&Build the Form"), 222, 215, 113, 25)
        self.b_close = P.put(QPushButton("&Close"), 341, 215, 113, 25)
        self.b_up.clicked.connect(lambda: self.move_selected(-1))
        self.b_down.clicked.connect(lambda: self.move_selected(1))
        self.b_build.clicked.connect(self.build)
        self.b_close.clicked.connect(self.reject)
        self.all.itemDoubleClicked.connect(lambda *_: self.move_fields(0))
        self.sel.itemDoubleClicked.connect(lambda *_: self.move_fields(3))
        self.source.editTextChanged.connect(self.source_edited)
        self.source.activated.connect(lambda *_: self.load_fields())
        self.source.lineEdit().editingFinished.connect(self.load_fields)
        try:
            self.source.addItems([td.get("Name") for td in vd.s.tabledefs(False)])
            self.source.addItems([q.get("Name") for q in vd.s.db.each("QueryDefs")])
        except RtError as e:
            vd.show_error(e)
        self.source.setCurrentIndex(-1)
        self.source.setEditText("")

    def source_edited(self, _text):
        if self.rs is not None and self.source.currentText() != self.rs_source:
            self.close_rs()
            self.all.clear()
            self.sel.clear()

    def close_rs(self):
        if self.rs is not None:
            try:
                self.rs.call("Close")
            except RtError:
                pass
        self.rs = None

    def load_fields(self):
        """the fields of the RecordSource (a table, a query or a SQL statement)"""
        text = self.source.currentText()
        if not text:
            return
        if self.rs is not None and self.rs_source == text:
            return
        self.all.clear()
        self.sel.clear()
        self.close_rs()
        try:
            self.rs = self.vd.s.db.call("OpenRecordset", text)
            self.rs_source = text
            self.all.addItems([f.get("Name") for f in self.rs.each("Fields")])
        except RtError as e:
            self.vd.say(e.description)
            return
        if self.all.count():
            self.all.setCurrentRow(0)

    def move_fields(self, how):
        a, s = self.all, self.sel
        if how == 0:
            i = a.currentRow()
            if i < 0:
                return
            s.addItem(a.takeItem(i).text())
            if a.count():
                a.setCurrentRow(min(i, a.count() - 1))
            s.setCurrentRow(s.count() - 1)
        elif how == 1:
            if a.count() == 0:
                return
            while a.count():
                s.addItem(a.takeItem(0).text())
            s.setCurrentRow(0)
        elif how == 2:
            if s.count() == 0:
                return
            while s.count():
                a.addItem(s.takeItem(0).text())
            a.setCurrentRow(a.count() - 1)
        else:
            i = s.currentRow()
            if i < 0:
                return
            a.addItem(s.takeItem(i).text())
            a.setCurrentRow(a.count() - 1)
            if s.count():
                s.setCurrentRow(min(i, s.count() - 1))

    def move_selected(self, d):
        s = self.sel
        i = s.currentRow()
        if i < 0 or not 0 <= i + d < s.count():
            return
        s.insertItem(i + d, s.takeItem(i))
        s.setCurrentRow(i + d)

    def build(self):
        vd = self.vd
        if not self.name.text():
            vd.say("Form Name cannot be blank!", "warn")
            self.name.setFocus()
            return None
        if not self.source.currentText():
            vd.say("You must enter a RecordSource!", "warn")
            return None
        if self.sel.count() == 0:
            vd.say("You must include some Columns!", "warn")
            return None
        try:
            fields = []
            for i in range(self.sel.count()):
                f = self.rs.get("Fields", self.sel.item(i).text())
                fields.append((f.get("Name"), f.get("Type"), f.get("Size")))
            caption = self.rs.get("Name")
            sqldb = vd.s.datatype == C.SQLDB
            text = form_text("frm" + self.name.text(), caption, fields, None if sqldb else vd.s.db.get("Name"),
                             vd.s.db.get("Connect") or "", self.source.currentText())
        except RtError as e:
            vd.say(e.description)
            return None
        name = "frm" + self.name.text()
        if vd.main is not None and getattr(vd.main, "project", None) is not None:
            if vd.main.project.module(name) is not None:
                vd.say(f"Name '{name}' conflicts with existing module, project, or object library", "warn")
                return None
            add_form(vd.main, name, text)
        self.name.setText("")
        self.source.setEditText("")
        self.name.setFocus()
        return text

    def done(self, r):
        self.close_rs()
        super().done(r)
