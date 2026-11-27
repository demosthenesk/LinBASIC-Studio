"""The smaller windows of the Visual Data Manager (VB6 ERRORS.FRM, ATTACH.FRM, NEWATTCH.FRM, CPYSTRU.FRM,
REPLACE.FRM, VBIMEX.FRM, DATATYPE.FRM, IMPTBLS.FRM, EXPNAME.FRM): Errors, Attachments and New Attached Table,
Copy Structure, Global Replace, Import/Export with Data Type, Import Tables and Export."""
import os

from PyQt6.QtWidgets import (QDialog, QWidget, QGridLayout, QHBoxLayout, QVBoxLayout, QLabel, QLineEdit, QListWidget,
                             QPushButton, QComboBox, QCheckBox, QPlainTextEdit, QAbstractItemView, QFileDialog)

from ..rtobj import RtError
from . import core as C

dbAttachExclusive, dbAttachSavePWD = 0x10000, 0x20000
dbText = 10


# ---------------------------------------------------------------- shared with VISDATA.BAS
def strip_connect(name):
    """StripConnect: 'Name -> Microsoft Access' -> 'Name'"""
    i = name.find("->")
    return name[:i - 1] if i > 0 else name


def strip_owner(name):
    i = name.find(".")
    return name[i + 1:] if i >= 0 else name


def table_list(vd):
    """RefreshTables of a list: the tables, attached ones with where they come from"""
    out = []
    for td in vd.s.tabledefs(vd.prefs.get("allow_sys", False)):
        name, a = td.get("Name"), td.get("Attributes")
        if a & C.dbAttachedTable == C.dbAttachedTable:
            con = td.get("Connect") or ""
            out.append(name + (" -> Microsoft Access" if con.startswith(";") else " -> " + con.split(";")[0]))
        elif a & C.dbAttachedODBC == C.dbAttachedODBC:
            out.append(name + " -> ODBC")
        else:
            out.append(name)
    return out


def copy_struct(vd, src, dst, from_name, to_name, indexes):
    """CopyStruct: the fields (name, type, size) and, when asked, the indexes of a table into a new table of the
    target database; a table of that name is deleted when the user says so, or another name is asked for"""
    while True:
        same = [td.get("Name") for td in dst.each("TableDefs") if td.get("Name").upper() == to_name.upper()]
        if not same:
            break
        if vd.ask(C.MSG[3]):
            dst.get("TableDefs").call("Delete", same[0])
            break
        to_name = vd.ask_text(C.MSG[4])
        if not to_name:
            return False
    td_from = src.get("TableDefs", from_name)
    td = dst.call("CreateTableDef", strip_owner(to_name))
    for f in td_from.each("Fields"):
        td.get("Fields").call("Append", td_from.call("CreateField", f.get("Name"), f.get("Type"), f.get("Size")))
    if indexes:
        for x in td_from.each("Indexes"):
            ix = td.call("CreateIndex", x.get("Name"))
            for f in x.each("Fields"):
                nf = ix.call("CreateField", f.get("Name"))
                if f.get("Attributes"):
                    nf.let("Attributes", f.get("Attributes"))
                ix.get("Fields").call("Append", nf)
            ix.let("Unique", x.get("Unique"))
            if vd.s.datatype != C.SQLDB:
                ix.let("Primary", x.get("Primary"))
            td.get("Indexes").call("Append", ix)
    dst.get("TableDefs").call("Append", td)
    return to_name


def copy_rows(vd, rs1, dst, to_name):
    """the rows of a recordset added to a table, committed every 1000 rows (CopyData)"""
    ws = vd.s.ws
    rs2 = dst.call("OpenRecordset", to_name)
    names = [f.get("Name") for f in rs1.each("Fields")]
    ws.call("BeginTrans")
    try:
        n = 0
        while not rs1.get("EOF"):
            rs2.call("AddNew")
            for nm in names:
                rs2.get("Fields", nm).let("Value", rs1.get("Fields", nm).get("Value"))
            rs2.call("Update")
            rs1.call("MoveNext")
            n += 1
            if n == 1000:
                ws.call("CommitTrans")
                ws.call("BeginTrans")
                n = 0
        ws.call("CommitTrans")
    except RtError:
        ws.call("Rollback")
        raise
    finally:
        rs2.call("Close")


def copy_data(vd, src, dst, from_name, to_name):
    try:
        rs1 = src.call("OpenRecordset", from_name)
        try:
            copy_rows(vd, rs1, dst, to_name)
        finally:
            rs1.call("Close")
        return True
    except RtError as e:
        vd.show_error(e)
        return False


def ix_fields(x):
    """the Fields of an Index as a string: "+A;-B" (dbDescending)"""
    return ";".join(("-" if f.get("Attributes") & 1 else "+") + f.get("Name") for f in x.each("Fields"))


def _buttons(*specs):
    row = QHBoxLayout()
    out = []
    for text in specs:
        b = QPushButton(text)
        row.addWidget(b)
        out.append(b)
    return row, out


# ---------------------------------------------------------------- Errors
class ErrorsWindow(QWidget):
    """the Errors collection of the DBEngine: number, source, description"""

    def __init__(self, vd, errs=None):
        super().__init__()
        self.vd = vd
        self.setWindowTitle("Errors")
        lay = QVBoxLayout(self)
        self.lst = QListWidget()
        lay.addWidget(self.lst, 1)
        row, (self.b_refresh, self.b_close) = _buttons("&Refresh", "&Close")
        row.addStretch(1)
        lay.addLayout(row)
        self.b_refresh.clicked.connect(self.refresh)
        self.b_close.clicked.connect(lambda: _close(self))
        self.resize(450, 205)
        self.fill(errs if errs is not None else vd.s.errors())

    def fill(self, errs):
        self.lst.clear()
        self.lst.addItems([f"{n}\t{s}\t{d}" for n, s, d in errs])

    def refresh(self):
        """RefreshErrors: no errors -> the message, the window goes"""
        errs = self.vd.s.errors()
        if not errs:
            self.vd.say(C.MSG[30], "warn")
            _close(self)
            return
        self.fill(errs)


def _close(w):
    p = w.parentWidget()
    (p if p is not None and p.inherits("QMdiSubWindow") else w).close()


# ---------------------------------------------------------------- Attachments
class AttachmentsWindow(QWidget):
    """the attached tables of the database: New (New Attached Table), ReAttach (RefreshLink)"""

    def __init__(self, vd):
        super().__init__()
        self.vd = vd
        self.setWindowTitle("Attachments")
        lay = QVBoxLayout(self)
        self.lst = QListWidget()
        self.lst.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        lay.addWidget(self.lst, 1)
        row, (self.b_new, self.b_reattach, self.b_close) = _buttons("&New", "&ReAttach", "&Close")
        lay.addLayout(row)
        self.b_new.clicked.connect(self.new)
        self.b_reattach.clicked.connect(self.reattach)
        self.b_close.clicked.connect(lambda: _close(self))
        self.lst.itemDoubleClicked.connect(lambda it: self.refresh_link(it.text()))
        self.resize(410, 225)
        try:
            for td in vd.s.db.each("TableDefs"):
                if C.attached(td):
                    self.lst.addItem(self.line(td.get("Name"), td.get("SourceTableName"), td.get("Connect")))
        except RtError as e:
            vd.show_error(e)

    @staticmethod
    def line(name, source, connect):
        return name + " " * max(0, 32 - len(name)) + "\t" + (source or "") + "=>" + (connect or "")

    @staticmethod
    def name_of(text):
        return text.split("\t")[0].strip()

    def refresh_link(self, text):
        try:
            self.vd.s.db.get("TableDefs", self.name_of(text)).call("RefreshLink")
            return True
        except RtError as e:
            self.vd.show_error(e)
            return False

    def reattach(self):
        for it in self.lst.selectedItems():
            self.refresh_link(it.text())
        self.vd.msgbar()

    def new(self):
        d = NewAttachDialog(self.vd, self)
        if not self.vd.quiet:
            d.exec()
        return d

    def add_attachment(self, name, database, kind, table, save_pwd=False, exclusive=False):
        """AddAttachment: a TableDef with SourceTableName and Connect appended"""
        from .sqlwin import dupe_name
        vd = self.vd
        try:
            if dupe_name(vd, name):
                return False
            vd.msgbar("Attaching " + name, True)
            connect = connect_str(kind, database)
            td = vd.s.db.call("CreateTableDef", name)
            td.let("SourceTableName", table)
            td.let("Connect", connect)
            a = (dbAttachSavePWD if save_pwd else 0) | (dbAttachExclusive if exclusive else 0)
            if a:
                td.let("Attributes", a)
            vd.s.db.get("TableDefs").call("Append", td)
        except RtError as e:
            vd.show_error(e)
            return False
        for i in range(self.lst.count()):
            if self.name_of(self.lst.item(i).text()).upper() == name.upper():
                self.lst.takeItem(i)
                break
        self.lst.addItem(self.line(name, table, connect))
        vd.msgbar()
        return True


ATTACH_KINDS = ["Access MDB", "Dbase III", "Dbase IV", "FoxPro 2.0", "FoxPro 2.5", "FoxPro 2.6", "Paradox 3.X",
                "Paradox 4.X", "Excel 5.0", "Excel 4.0", "Excel 3.0", "Text Files", "ODBC"]


def connect_str(kind, database):
    """GetConnectStr: ';database=x' for Access, 'Kind;database=x' for the others"""
    s = ";" if kind in ("", "Access MDB") else kind + ";"
    return s + ("database=" + database if database else "")


class NewAttachDialog(QDialog):
    def __init__(self, vd, attachments):
        super().__init__(vd)
        self.vd, self.att = vd, attachments
        self.added = False
        self.setWindowTitle("New Attached Table")
        g = QVBoxLayout(self)
        self.name, self.database = QLineEdit(), QLineEdit()
        self.connect, self.table = QComboBox(), QComboBox()
        self.connect.addItems(ATTACH_KINDS)
        self.table.setEditable(True)
        for t, w in (("Attachment &Name:", self.name), ("&Database Name:", self.database),
                     ("C&onnect String:", self.connect), ("&Table to Attach:", self.table)):
            lb = QLabel(t)
            lb.setBuddy(w)
            g.addWidget(lb)
            g.addWidget(w)
        row = QHBoxLayout()
        self.save_pwd, self.exclusive = QCheckBox("AttachSave&PWD"), QCheckBox("Attach&Exclusive")
        row.addWidget(self.save_pwd)
        row.addWidget(self.exclusive)
        g.addLayout(row)
        row, (self.b_attach, self.b_cancel) = _buttons("&Attach", "&Cancel")
        g.addLayout(row)
        self.b_attach.setEnabled(False)
        self.name.textChanged.connect(self.changed)
        self.connect.currentIndexChanged.connect(self.changed)
        self.database.textChanged.connect(lambda: self.table.clear())
        self.b_attach.clicked.connect(self.attach)
        self.b_cancel.clicked.connect(self.cancel)
        self.table.showPopup = self._popup

    def changed(self, *_):
        self.b_attach.setEnabled(bool(self.name.text()))
        self.table.clear()

    def tables(self):
        """the tables of the database to attach from (cboTableName_DropDown)"""
        out = []
        try:
            kind = self.connect.currentText()
            ws = self.vd.s.ws
            if self.connect.currentIndex() == 0:
                db = ws.call("OpenDatabase", self.database.text(), False, False)
            else:
                db = ws.call("OpenDatabase", self.database.text(), False, False, kind + ";")
            out = [td.get("Name") for td in db.each("TableDefs") if td.get("Attributes") & C.dbSystemObject == 0]
            db.call("Close")
        except RtError as e:
            self.vd.show_error(e)
        return out

    def _popup(self):
        if self.table.count() == 0:
            self.table.addItems(self.tables())
        QComboBox.showPopup(self.table)

    def attach(self):
        if self.att.add_attachment(self.name.text(), self.database.text(), self.connect.currentText(),
                                   self.table.currentText(), self.save_pwd.isChecked(), self.exclusive.isChecked()):
            self.added = True
            self.name.clear()
            self.table.setEditText("")

    def cancel(self):
        if self.added:
            self.vd.refresh_tables()
        self.reject()


# ---------------------------------------------------------------- Copy Structure
class CopyStructureDialog(QDialog):
    def __init__(self, vd, current=""):
        super().__init__(vd)
        self.vd = vd
        self.setWindowTitle("Copy Structure")
        g = QGridLayout(self)
        g.addWidget(QLabel("Tables:"), 0, 0)
        self.lst = QListWidget()
        self.lst.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        g.addWidget(self.lst, 1, 0, 7, 1)
        g.addWidget(QLabel("Target Database:"), 0, 1)
        self.database = QLineEdit(vd.s.dbname)
        g.addWidget(self.database, 1, 1)
        g.addWidget(QLabel("Target Connect String:"), 2, 1)
        self.connect = QPlainTextEdit()
        self.connect.setMaximumHeight(40)
        g.addWidget(self.connect, 3, 1)
        self.c_indexes, self.c_data = QCheckBox("Copy Indexes"), QCheckBox("Copy Data")
        g.addWidget(self.c_indexes, 4, 1)
        g.addWidget(self.c_data, 5, 1)
        g.addWidget(QLabel("(Note: Export is Faster)"), 6, 1)
        row, (self.b_ok, self.b_close) = _buttons("&OK", "&Close")
        g.addLayout(row, 7, 1)
        self.b_ok.clicked.connect(self.ok)
        self.b_close.clicked.connect(self.close_dialog)
        try:
            self.connect.setPlainText(vd.s.db.get("Connect") or "")
            self.lst.addItems(table_list(vd))
        except RtError as e:
            vd.show_error(e)
        for i in range(self.lst.count()):
            if self.lst.item(i).text() == current:
                self.lst.item(i).setSelected(True)
                break

    def ok(self):
        vd = self.vd
        vd.msgbar("Copying Table(s)", True)
        try:
            target = self.database.text()
            if not target or target == vd.s.dbname:
                dst, different = vd.s.db, False
            else:
                dst = vd.s.ws.call("OpenDatabase", target, False, False, self.connect.toPlainText())
                dst.let("QueryTimeout", vd.prefs.get("query_timeout", 5))
                different = True
        except RtError as e:
            vd.show_error(e)
            return
        for i in range(self.lst.count()):
            it = self.lst.item(i)
            if not it.isSelected():
                continue
            src_name = strip_connect(it.text())
            to_name = src_name if different else vd.ask_text("Enter New Table Name:")
            if not to_name:
                continue
            try:
                done = copy_struct(vd, vd.s.db, dst, src_name, to_name, self.c_indexes.isChecked())
            except RtError as e:
                vd.show_error(e)
                done = False
            if done:
                if self.c_data.isChecked() and not copy_data(vd, vd.s.db, dst, src_name, done):
                    vd.say(src_name + ": " + "Copy of Data was Unsuccessful!")
                vd.say(src_name + ": " + "Copy of Structure was Successful!")
                it.setSelected(False)
            else:
                vd.say(src_name + ": " + "Copy of Data was Unsuccessful!")
        vd.msgbar()

    def close_dialog(self):
        self.vd.refresh_tables()
        self.reject()


# ---------------------------------------------------------------- Global Replace
class GlobalReplaceDialog(QDialog):
    def __init__(self, vd):
        super().__init__(vd)
        self.vd = vd
        self.setWindowTitle("Global Replace")
        g = QGridLayout(self)
        g.addWidget(QLabel("&Table:"), 0, 0)
        g.addWidget(QLabel("&Field:"), 0, 1)
        self.tables, self.fields = QListWidget(), QListWidget()
        g.addWidget(self.tables, 1, 0)
        g.addWidget(self.fields, 1, 1)
        g.addWidget(QLabel("&Replace With:"), 2, 0, 1, 2)
        self.replace = QLineEdit()
        g.addWidget(self.replace, 3, 0, 1, 2)
        g.addWidget(QLabel("Cr&iteria:"), 4, 0, 1, 2)
        self.criteria = QLineEdit()
        g.addWidget(self.criteria, 5, 0, 1, 2)
        row, (self.b_ok, self.b_cancel) = _buttons("&OK", "&Cancel")
        g.addLayout(row, 6, 0, 1, 2)
        self.b_ok.setEnabled(False)
        self.tables.currentTextChanged.connect(self.table_changed)
        self.fields.currentTextChanged.connect(self.enable)
        self.replace.textChanged.connect(self.enable)
        self.b_ok.clicked.connect(self.ok)
        self.b_cancel.clicked.connect(self.reject)
        try:
            self.tables.addItems(table_list(vd))
        except RtError as e:
            vd.show_error(e)

    def enable(self, *_):
        self.b_ok.setEnabled(bool(self.fields.currentItem()) and bool(self.replace.text()))

    def table_changed(self, text):
        self.fields.clear()
        if not text:
            return
        try:
            self.fields.addItems([f.get("Name") for f in
                                  self.vd.s.db.get("TableDefs", strip_connect(text)).each("Fields")])
        except RtError as e:
            self.vd.show_error(e)

    def ok(self):
        vd, db = self.vd, self.vd.s.db
        table, field = strip_connect(self.tables.currentItem().text()), self.fields.currentItem().text()
        vd.msgbar("Replacing Records", True)
        try:
            if db.get("TableDefs", table).get("Fields", field).get("Type") == dbText:
                value = "'" + self.replace.text() + "'"
            else:
                value = self.replace.text()
            where = " where " + self.criteria.text() if self.criteria.text() else ""
            if vd.s.datatype == C.SQLDB:
                db.call("Execute", f"update {table} set {field}={value}{where}", C.dbSQLPassThrough)
            else:
                db.call("Execute", f"update {table} set [{field}]={value}{where}")
            if db.get("RecordsAffected") > 0 and vd.s.trans_pending:
                vd.s.changed = True
        except RtError as e:
            vd.show_error(e)
            return False
        self.accept()
        vd.msgbar()
        return True


# ---------------------------------------------------------------- Import/Export
DATA_TYPES = ["Access MDB", "Dbase IV", "Dbase III", "FoxPro 2.6", "FoxPro 2.5", "FoxPro 2.0", "Paradox 4.X",
              "Paradox 3.X", "Btrieve", "Excel 5.0", "Excel 4.0", "Excel 3.0", "Text", "ODBC"]
DT_MSACCESS, DT_BTRIEVE, DT_TEXT, DT_SQLDB = 0, 8, 12, 13
DT_CONNECT = {1: "dBase IV;", 2: "dBase III;", 3: "FoxPro 2.6;", 4: "FoxPro 2.5;", 5: "FoxPro 2.0;", 6: "Paradox 4.X;",
              7: "Paradox 3.X;", 8: "Btrieve;", 9: "Excel 5.0;", 10: "Excel 5.0;", 11: "Excel 5.0;", 12: "Text;"}
DT_FILTER = {0: ("Select Microsoft Access Database", "Microsoft Access (*.mdb);;All Files (*.*)"),
             6: (None, "Paradox (*.db)"), 7: (None, "Paradox (*.db)"), 8: (None, "Btrieve (FILE.DDF)"),
             12: ("Select Text File", "Text Files (*.txt);;All Files (*.*)")}
for _k in (1, 2, 3, 4, 5):
    DT_FILTER[_k] = (None, "Xbase (*.dbf)")
for _k in (9, 10, 11):
    DT_FILTER[_k] = (None, "Excel (*.xls)")


class DataTypeDialog(QDialog):
    """Data Type: the kind of data source to import from or export to (-1: Cancel)"""

    def __init__(self, parent, caption):
        super().__init__(parent)
        self.setWindowTitle(caption)
        self.kind = -1
        g = QGridLayout(self)
        g.addWidget(QLabel("Data Source:"), 0, 0)
        self.lst = QListWidget()
        self.lst.addItems(DATA_TYPES)
        self.lst.setCurrentRow(0)
        g.addWidget(self.lst, 1, 0, 2, 1)
        ok, cancel = QPushButton("&OK"), QPushButton("&Cancel")
        g.addWidget(ok, 1, 1)
        g.addWidget(cancel, 2, 1)
        ok.clicked.connect(self.ok)
        cancel.clicked.connect(self.reject)
        self.lst.itemDoubleClicked.connect(lambda *_: self.ok())

    def ok(self):
        self.kind = self.lst.currentRow()
        self.accept()


class ImportExportDialog(QDialog):
    def __init__(self, vd):
        super().__init__(vd)
        self.vd = vd
        self.setWindowTitle("Import/Export")
        g = QGridLayout(self)
        self.b_import, self.b_close = QPushButton("&Import"), QPushButton("&Close")
        self.b_export, self.b_export_sql = QPushButton("&Export Table(s)"), QPushButton("Export &SQL Results")
        g.addWidget(self.b_import, 0, 0)
        g.addWidget(self.b_close, 0, 1)
        g.addWidget(self.b_export, 1, 0)
        g.addWidget(self.b_export_sql, 1, 1)
        self.lst = QListWidget()
        self.lst.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        g.addWidget(self.lst, 2, 0, 1, 2)
        self.b_export.setEnabled(False)
        sql = vd.sqlwin.text.toPlainText() if vd.sqlwin is not None else ""
        self.b_export_sql.setEnabled(sql[:7].upper() == "SELECT ")
        self.lst.itemClicked.connect(lambda *_: self.b_export.setEnabled(True))
        self.lst.itemDoubleClicked.connect(lambda *_: self.export_tables())
        self.b_close.clicked.connect(self.reject)
        self.b_import.clicked.connect(lambda: self.import_())
        self.b_export.clicked.connect(lambda: self.export_tables())
        self.b_export_sql.clicked.connect(lambda: self.export_sql())
        try:
            self.lst.addItems([td.get("Name") for td in vd.s.tabledefs(False)])
        except RtError as e:
            vd.show_error(e)
        vd.msgbar()

    # the questions of the dialog (in tests: the answers queue of the VisData window)
    def data_type(self, caption):
        if self.vd.quiet:
            return self.vd.answers.pop(0) if self.vd.answers else -1
        d = DataTypeDialog(self, caption)
        return d.kind if d.exec() else -1

    def file_name(self, kind, save):
        if self.vd.quiet:
            return self.vd.answers.pop(0) if self.vd.answers else ""
        title, filt = DT_FILTER.get(kind, (None, "All Files (*.*)"))
        title = title or "Select File"
        if save:
            return QFileDialog.getSaveFileName(self, title, "", filt)[0]
        return QFileDialog.getOpenFileName(self, title, "", filt)[0]

    def export_name(self, table, target):
        """Export: the name of the table in the target ('' = Cancel)"""
        if self.vd.quiet:
            return self.vd.answers.pop(0) if self.vd.answers else table
        d = QDialog(self)
        d.setWindowTitle("Export")
        g = QGridLayout(d)
        g.addWidget(QLabel(C.MSG[57] + table + " ->"), 0, 0)
        e = QLineEdit(table)
        g.addWidget(e, 1, 0)
        g.addWidget(QLabel(C.MSG[58] + target), 2, 0)
        ok, cancel = QPushButton("&OK"), QPushButton("&Cancel")
        g.addWidget(ok, 0, 1)
        g.addWidget(cancel, 1, 1)
        ok.clicked.connect(d.accept)
        cancel.clicked.connect(d.reject)
        return e.text() if d.exec() else ""

    def open_target(self, kind, path):
        """the database to export to or import from (gExpDB / gImpDB of VisData), and the connect part of the SQL"""
        ws = self.vd.s.ws
        if kind == DT_SQLDB:
            cn = self.vd.odbc_logon()
            if not cn:
                return None, ""
            db = ws.call("OpenDatabase", "", False, False, cn)
            return db, "[" + db.get("Connect") + "]."
        if kind == DT_MSACCESS:
            return ws.call("OpenDatabase", path), "[;database=" + path + "]."
        if kind in (8, 9, 10, 11):
            con = DT_CONNECT[kind]
            return ws.call("OpenDatabase", path, False, False, con), "[" + con + "database=" + path + "]."
        folder = os.path.dirname(path)
        con = DT_CONNECT[kind]
        return ws.call("OpenDatabase", folder, False, False, con), "[" + con + "database=" + folder + "]."

    # ---- Export
    def export_tables(self):
        for i in range(self.lst.count()):
            it = self.lst.item(i)
            if not it.isSelected():
                continue
            kind = self.data_type(C.MSG[57].strip() + " '" + it.text() + "'")
            it.setSelected(False)
            if kind < 0:
                continue
            path = "" if kind == DT_SQLDB else self.file_name(kind, True)
            if kind == DT_SQLDB or path:
                self.export(it.text(), path, kind)
        self.b_export.setEnabled(False)

    def export_sql(self):
        kind = self.data_type("Export SQL")
        if kind < 0:
            return
        path = "" if kind == DT_SQLDB else self.file_name(kind, True)
        if kind == DT_SQLDB or path:
            self.export("", path, kind)

    def export(self, table, path, kind):
        """Export (VISDATA.BAS): SELECT * INTO the table of the other database, then its indexes"""
        vd, db = self.vd, self.vd.s.db
        new, dst, state = "", None, 1
        try:
            vd.msgbar(C.MSG[56] + "'" + table + "'", True)
            dst, connect = self.open_target(kind, path)
            if dst is None:
                vd.msgbar()
                return False
            if kind in (DT_MSACCESS, DT_BTRIEVE, DT_SQLDB, 9, 10, 11):
                new = self.export_name(table, path)
                if not new:
                    vd.msgbar()
                    return False
            else:
                new = os.path.splitext(os.path.basename(path.replace("\\", "/")))[0]
            while True:
                try:
                    if table:
                        db.call("Execute", "select * into " + connect + strip_owner(new) + " from " + strip_owner(table))
                        if kind != DT_TEXT:
                            state = 2
                            vd.msgbar(C.MSG[59] + " '" + new + "'", True)
                            dst.get("TableDefs").call("Refresh")
                            td_to = dst.get("TableDefs", new)
                            for x in db.get("TableDefs", table).each("Indexes"):
                                ix = td_to.call("CreateIndex", x.get("Name"))
                                ix.let("Fields", ix_fields(x))
                                ix.let("Unique", x.get("Unique"))
                                if kind != DT_SQLDB and vd.s.datatype != C.SQLDB:
                                    ix.let("Primary", x.get("Primary"))
                                td_to.get("Indexes").call("Append", ix)
                        vd.msgbar()
                        vd.say(C.MSG[60] + " '" + table + "'")
                    else:
                        sql = vd.sqlwin.text.toPlainText()
                        up = sql.upper()
                        fields = sql[7:up.find("FROM", 7) - 1]
                        frm = " " + sql[up.find("FROM"):]
                        db.call("Execute", "select " + fields + " into " + connect + new + frm)
                        vd.msgbar()
                        vd.say(C.MSG[61])
                    break
                except RtError as e:
                    if e.number == 3010:                # the table exists: overwrite?
                        if vd.ask(C.MSG[62], default=False):
                            dst.get("TableDefs").call("Delete", new)
                            continue
                        vd.msgbar()
                        return False
                    raise
            return True
        except RtError as e:
            if state == 2 and dst is not None:     # the indexes could not be made: the new table goes
                try:
                    dst.get("TableDefs").call("Delete", new)
                except RtError:
                    pass
            vd.show_error(e)
            return False

    # ---- Import
    def import_(self):
        vd = self.vd
        kind = self.data_type("Import")
        if kind < 0:
            return False
        try:
            if kind == DT_SQLDB:
                self.imp_db, self.imp_connect = self.open_target(kind, "")
                return self.import_tables(kind) if self.imp_db is not None else False
            path = self.file_name(kind, False)
            if not path:
                return False
            self.imp_db, self.imp_connect = self.open_target(kind, path)
            if kind in (DT_MSACCESS, DT_BTRIEVE, 9, 10, 11):
                return self.import_tables(kind)
            return self.import_table(kind, os.path.splitext(os.path.basename(path.replace("\\", "/")))[0])
        except RtError as e:
            vd.show_error(e)
            return False

    def import_tables(self, kind):
        """Import Tables: the tables of the source; Import copies the one chosen"""
        names = [td.get("Name") for td in self.imp_db.each("TableDefs")
                 if td.get("Attributes") & C.dbSystemObject == 0]
        if self.vd.quiet:
            pick = self.vd.answers.pop(0) if self.vd.answers else ""
            return self.import_table(kind, pick) if pick else False
        d = QDialog(self)
        d.setWindowTitle("Import Tables")
        g = QGridLayout(d)
        g.addWidget(QLabel("Tables In:" + (self.imp_db.get("Name") or "")), 0, 0)
        lst = QListWidget()
        lst.addItems(names)
        g.addWidget(lst, 1, 0, 2, 1)
        b_imp, b_close = QPushButton("&Import"), QPushButton("&Close")
        b_imp.setEnabled(False)
        g.addWidget(b_imp, 1, 1)
        g.addWidget(b_close, 2, 1)
        lst.itemClicked.connect(lambda *_: b_imp.setEnabled(True))
        b_imp.clicked.connect(lambda: lst.currentItem() and self.import_table(kind, lst.currentItem().text()))
        lst.itemDoubleClicked.connect(lambda it: self.import_table(kind, it.text()))
        b_close.clicked.connect(d.reject)
        d.exec()
        return True

    def import_table(self, kind, name):
        """Import (VISDATA.BAS): SELECT * INTO a new table FROM the table of the other database, then its indexes"""
        from .sqlwin import dupe_name
        vd, db = self.vd, self.vd.s.db
        old = name
        new = old.replace(".", "_", 1) if kind == DT_SQLDB and "." in old else old
        if dupe_name(vd, new):
            t = 1
            while dupe_name(vd, new + str(t)):
                t += 1
            new = new + str(t)
        vd.msgbar(C.MSG[63] + "'" + new + "'", True)
        state = 1
        try:
            db.call("Execute", "select * into " + new + " from " + self.imp_connect + old)
            if kind not in (DT_TEXT, 9, 10, 11):
                state = 2
                vd.msgbar(f"{db.get('RecordsAffected')} Rows Imported, Creating Indexes for '{new}'", True)
                db.get("TableDefs").call("Refresh")
                td_to = db.get("TableDefs", new)
                for x in self.imp_db.get("TableDefs", old).each("Indexes"):
                    ix = td_to.call("CreateIndex", x.get("Name"))
                    ix.let("Fields", ix_fields(x))
                    ix.let("Unique", x.get("Unique"))
                    if kind != DT_SQLDB and vd.s.datatype != C.SQLDB:
                        ix.let("Primary", x.get("Primary"))
                    td_to.get("Indexes").call("Append", ix)
        except RtError as e:
            if state == 2:                          # the new table goes when its indexes could not be made
                try:
                    db.get("TableDefs").call("Delete", new)
                except RtError:
                    pass
            vd.show_error(e)
            return False
        self.lst.addItem(new)
        vd.msgbar()
        vd.say(C.MSG[64] + "'" + new + "'.")
        vd.refresh_tables()
        return True
