"""Table Structure of the Visual Data Manager (VB6 TBLSTRU.FRM, ADDFIELD.FRM, ADDINDEX.FRM): the fields and indexes of
a table (Design...) or of a new one (New Table: Build the Table), Add Field, Remove Field, Add Index, Remove Index,
Print Structure."""
from PyQt6.QtWidgets import (QDialog, QLineEdit, QListWidget, QPushButton, QComboBox, QCheckBox, QRadioButton,
                             QPlainTextEdit)

from ..rtobj import RtError
from . import core as C

TYPES = ["Boolean", "Byte", "Integer", "Long", "Currency", "Single", "Double", "Date/Time", "Text", "Binary", "Memo"]
TYPE_CODES = {"Boolean": 1, "Byte": 2, "Integer": 3, "Long": 4, "Currency": 5, "Single": 6, "Double": 7,
              "Date/Time": 8, "Text": 10, "Binary": 11, "Memo": 12, "Counter": 4}
CODE_NAMES = {1: "Boolean", 2: "Byte", 3: "Integer", 4: "Long", 5: "Currency", 6: "Single", 7: "Double",
              8: "Date/Time", 10: "Text", 11: "Binary", 12: "Memo", 15: "GUID", 20: "Decimal"}
SIZES = {1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 4, 7: 8, 8: 8, 10: 50, 11: 0, 12: 0}
dbFixedField, dbVariableField, dbAutoIncrField = 1, 2, 16


class FieldSpec:
    def __init__(self, name, type_, size, attrs=0, required=False, allowzero=False, defval="", vrule="", vtext="",
                 ord_=0, collating=1033):
        self.name, self.type, self.size, self.attrs = name, type_, size, attrs
        self.required, self.allowzero = required, allowzero
        self.defval, self.vrule, self.vtext, self.ord, self.collating = defval, vrule, vtext, ord_, collating


class IndexSpec:
    def __init__(self, name, fields, primary=False, unique=False, required=False, ignorenulls=False, foreign=False):
        self.name, self.fields = name, fields
        self.primary, self.unique, self.required, self.ignorenulls, self.foreign = primary, unique, required, \
            ignorenulls, foreign


class TableStructure(QDialog):
    def __init__(self, vd, table):
        super().__init__(vd)
        self.vd = vd
        self.table = table                      # None: a new table (Build the Table)
        self.setWindowTitle("Table Structure")
        self.fields, self.indexes = [], []
        P = C.Place(self, 512, 409)             # (VB6 oracles vb_tblstru.png, vb_tblnew.png)
        self.name = QLineEdit(table or "")
        P.label("Table &Name:", 8, 8, 63, buddy=self.name)
        P.put(self.name, 128, 8, 209, 19)
        self.lst_fields = QListWidget()
        P.label("&Field List:", 8, 32, 48, buddy=self.lst_fields)
        P.put(self.lst_fields, 7, 48, 193, 173)
        self.b_addfield = P.put(QPushButton("&Add Field"), 8, 224, 96, 25)
        self.b_rmfield = P.put(QPushButton("&Remove Field"), 103, 224, 96, 25)
        # the properties of the field chosen: Name (Locked); Type, Size, CollatingOrder and the kind of field are
        # shown, not changed (picFieldProps1); AllowZeroLength, Required and the rest (picFieldProps2)
        self.f_name, self.f_type, self.f_size, self.f_coll = QLineEdit(), QLineEdit(), QLineEdit(), QLineEdit()
        self.f_ord, self.f_vtext, self.f_vrule, self.f_def = QLineEdit(), QLineEdit(), QLineEdit(), QLineEdit()
        self.f_fixed, self.f_var, self.f_auto = QCheckBox("FixedLength"), QCheckBox("VariableLength"), \
            QCheckBox("AutoIncrement")
        self.f_azl, self.f_req = QCheckBox("AllowZeroLength"), QCheckBox("Required")
        for t, y in (("Name: ", 32), ("Type: ", 56), ("Size: ", 80), ("CollatingOrder: ", 104),
                     ("OrdinalPosition: ", 157), ("ValidationText: ", 181), ("ValidationRule: ", 205),
                     ("DefaultValue: ", 229)):
            P.label(t, 208, y, 78)
        P.put(self.f_name, 304, 32, 193, 19)
        P.put(self.f_type, 304, 56, 73, 19)
        P.put(self.f_size, 304, 80, 73, 19)
        P.put(self.f_coll, 304, 104, 73, 19)
        P.put(self.f_fixed, 384, 56, 93, 17)
        P.put(self.f_var, 384, 80, 93, 17)
        P.put(self.f_auto, 384, 104, 93, 17)
        P.put(self.f_azl, 384, 128, 113, 17)
        P.put(self.f_ord, 304, 152, 73, 19)
        P.put(self.f_req, 384, 152, 81, 17)
        P.put(self.f_vtext, 304, 176, 193, 19)
        P.put(self.f_vrule, 304, 200, 193, 19)
        P.put(self.f_def, 304, 224, 193, 19)
        self.f_name.setReadOnly(True)
        for w in (self.f_type, self.f_size, self.f_coll):
            w.setReadOnly(True)
            w.setEnabled(table is None)         # (a table of the database: grey)
        for w in (self.f_fixed, self.f_var, self.f_auto):
            w.setEnabled(False)
        for w in (self.f_ord, self.f_vtext, self.f_vrule, self.f_def):
            w.setReadOnly(table is not None)
        P.line(8, 256, 504)
        self.lst_indexes = QListWidget()
        P.label("Inde&x List:", 8, 265, 57, buddy=self.lst_indexes)
        P.put(self.lst_indexes, 8, 281, 193, 56)
        self.b_addindex = P.put(QPushButton("Add &Index"), 8, 341, 96, 25)
        self.b_rmindex = P.put(QPushButton("Re&move Index"), 104, 341, 96, 25)
        self.i_name, self.i_fields = QLineEdit(), QLineEdit()
        self.i_name.setReadOnly(True)
        self.i_fields.setReadOnly(True)
        P.label("Name: ", 216, 265, 34)
        P.put(self.i_name, 312, 265, 193, 19)
        self.i_primary, self.i_unique, self.i_foreign, self.i_req, self.i_ign = (
            QCheckBox(t) for t in ("Primary", "Unique", "Foreign", "Required", "IgnoreNull"))
        for w, x, y in ((self.i_primary, 216, 289), (self.i_unique, 312, 289), (self.i_foreign, 416, 289),
                        (self.i_req, 216, 313), (self.i_ign, 312, 313)):
            w.setEnabled(False)
            P.put(w, x, y, 82, 17)
        P.label("Fields: ", 216, 337, 34)
        P.put(self.i_fields, 264, 337, 241, 19)
        self.b_build = P.put(QPushButton("&Build the Table"), 16, 376, 153, 25)
        self.b_close = P.put(QPushButton("&Close"), 184, 376, 145, 25)
        self.b_print = P.put(QPushButton("&Print Structure"), 344, 376, 153, 25)
        self.b_build.setVisible(table is None)
        self.b_build.setEnabled(False)
        self.b_print.setVisible(table is not None)
        self.name.textChanged.connect(self.enable_build)
        self.lst_fields.currentRowChanged.connect(self.show_field)
        self.lst_indexes.currentRowChanged.connect(self.show_index)
        self.b_addfield.clicked.connect(self.add_field)
        self.b_rmfield.clicked.connect(self.remove_field)
        self.b_addindex.clicked.connect(self.add_index)
        self.b_rmindex.clicked.connect(self.remove_index)
        self.b_build.clicked.connect(self.build)
        self.b_close.clicked.connect(self.close_dialog)
        self.b_print.clicked.connect(self.print_structure)
        if table is not None:
            self.load()

    def td(self):
        return self.vd.s.db.get("TableDefs", self.table)

    def load(self):
        try:
            td = self.td()
            self.fields = [FieldSpec(f.get("Name"), f.get("Type"), f.get("Size"), f.get("Attributes"),
                                     f.get("Required"), f.get("AllowZeroLength"), f.get("DefaultValue"),
                                     f.get("ValidationRule"), f.get("ValidationText"), f.get("OrdinalPosition"))
                           for f in td.each("Fields")]
            from .dialogs import ix_fields
            self.indexes = [IndexSpec(x.get("Name"), ix_fields(x), x.get("Primary"), x.get("Unique"),
                                      x.get("Required"), x.get("IgnoreNulls"), x.get("Foreign"))
                            for x in td.each("Indexes")]
        except RtError as e:
            self.vd.show_error(e)
        self.fill()

    def fill(self):
        self.lst_fields.clear()
        self.lst_fields.addItems([f.name for f in self.fields])
        self.lst_indexes.clear()
        self.lst_indexes.addItems([x.name for x in self.indexes])
        if self.fields:
            self.lst_fields.setCurrentRow(0)
        if self.indexes:
            self.lst_indexes.setCurrentRow(0)
        self.enable_build()

    def enable_build(self, *_):
        """Build the Table: a name and a field"""
        self.b_build.setEnabled(bool(self.name.text().strip()) and bool(self.fields))

    def show_field(self, i):
        if not 0 <= i < len(self.fields):
            return
        f = self.fields[i]
        self.f_name.setText(f.name)
        self.f_type.setText(CODE_NAMES.get(f.type, str(f.type)))
        self.f_size.setText(str(f.size))
        self.f_coll.setText(str(f.collating))
        self.f_ord.setText(str(f.ord))
        self.f_vtext.setText(f.vtext or "")
        self.f_vrule.setText(f.vrule or "")
        self.f_def.setText(f.defval or "")
        self.f_fixed.setChecked(bool(f.attrs & dbFixedField))
        self.f_var.setChecked(bool(f.attrs & dbVariableField))
        self.f_auto.setChecked(bool(f.attrs & dbAutoIncrField))
        self.f_azl.setChecked(bool(f.allowzero))
        self.f_req.setChecked(bool(f.required))

    def show_index(self, i):
        if not 0 <= i < len(self.indexes):
            return
        x = self.indexes[i]
        self.i_name.setText(x.name)
        self.i_fields.setText(x.fields)
        for w, v in ((self.i_primary, x.primary), (self.i_unique, x.unique), (self.i_foreign, x.foreign),
                     (self.i_req, x.required), (self.i_ign, x.ignorenulls)):
            w.setChecked(bool(v))

    # ---------------------------------------------------------------- fields and indexes
    def add_field(self, spec=None):
        """Add Field: a field of the new table, or added to the table at once"""
        if spec is None:
            d = AddFieldDialog(self, self.vd)
            self.vd.msgbar("Enter New Field Parameters, Press 'Close' when finished")
            if self.vd.quiet:
                return None
            d.exec()
            self.vd.msgbar()
            return None
        if self.table is not None:
            try:
                td = self.td()
                f = td.call("CreateField", spec.name, spec.type, spec.size)
                if spec.attrs:
                    f.let("Attributes", spec.attrs)
                f.let("Required", bool(spec.required))
                if spec.type in (10, 12):
                    f.let("AllowZeroLength", bool(spec.allowzero))
                for k, v in (("DefaultValue", spec.defval), ("ValidationRule", spec.vrule),
                             ("ValidationText", spec.vtext)):
                    if v:
                        f.let(k, v)
                td.get("Fields").call("Append", f)
            except RtError as e:
                self.vd.show_error(e)
                return False
        self.fields.append(spec)
        self.fill()
        self.lst_fields.setCurrentRow(len(self.fields) - 1)
        return True

    def remove_field(self):
        i = self.lst_fields.currentRow()
        if i < 0 or not self.vd.ask("Remove Field?"):
            return
        if self.table is not None:
            try:
                self.td().get("Fields").call("Delete", self.fields[i].name)
            except RtError as e:
                self.vd.show_error(e)
                return
        del self.fields[i]
        self.fill()

    def add_index(self, spec=None):
        if spec is None:
            d = AddIndexDialog(self, self.vd, [f.name for f in self.fields])
            self.vd.msgbar("Enter New Index Parameters, Press 'Close' when finished")
            if self.vd.quiet:
                return None
            d.exec()
            self.vd.msgbar()
            return None
        if self.table is not None:
            try:
                td = self.td()
                ix = make_index(td, spec)
                td.get("Indexes").call("Append", ix)
            except RtError as e:
                self.vd.show_error(e)
                return False
        self.indexes.append(spec)
        self.fill()
        return True

    def remove_index(self):
        i = self.lst_indexes.currentRow()
        if i < 0 or not self.vd.ask("Remove Index?"):
            return
        if self.table is not None:
            try:
                self.td().get("Indexes").call("Delete", self.indexes[i].name)
            except RtError as e:
                self.vd.show_error(e)
                return
        del self.indexes[i]
        self.fill()

    def build(self):
        """Build the Table: the new TableDef with its fields and indexes appended"""
        name = self.name.text().strip()
        if not name or not self.fields:
            return False
        vd, db = self.vd, self.vd.s.db
        vd.msgbar("Adding the New Table to the Database", True)
        try:
            td = db.call("CreateTableDef", name)
            for spec in self.fields:
                f = td.call("CreateField", spec.name, spec.type, spec.size)
                if spec.attrs:
                    f.let("Attributes", spec.attrs)
                f.let("Required", bool(spec.required))
                if spec.type in (10, 12):
                    f.let("AllowZeroLength", bool(spec.allowzero))
                for k, v in (("DefaultValue", spec.defval), ("ValidationRule", spec.vrule),
                             ("ValidationText", spec.vtext)):
                    if v:
                        f.let(k, v)
                td.get("Fields").call("Append", f)
            for spec in self.indexes:
                td.get("Indexes").call("Append", make_index(td, spec))
            db.get("TableDefs").call("Append", td)
        except RtError as e:
            vd.show_error(e)
            return False
        vd.msgbar()
        vd.refresh_tables()
        self.table = name
        self.accept()
        return True

    def close_dialog(self):
        if self.table is None and self.fields and not self.vd.ask("Close without saving?"):
            return
        if self.table is not None:
            self.vd.refresh_tables()
        self.reject()

    def structure_text(self):
        out = [f"Table: {self.table or self.name.text()}", "", "Fields:"]
        for f in self.fields:
            out.append(f"  {f.name:<30} {CODE_NAMES.get(f.type, f.type):<10} {f.size}")
        out += ["", "Indexes:"]
        for x in self.indexes:
            flags = " ".join(n for n, v in (("Primary", x.primary), ("Unique", x.unique), ("Required", x.required),
                                            ("IgnoreNulls", x.ignorenulls)) if v)
            out.append(f"  {x.name:<30} {x.fields}  {flags}")
        return "\n".join(out)

    def print_structure(self):
        """Print Structure: the fields and indexes to the printer"""
        from PyQt6.QtGui import QTextDocument
        from PyQt6.QtPrintSupport import QPrintDialog, QPrinter
        self.vd.msgbar("Printing Table Structure", True)
        text = self.structure_text()
        if not self.vd.quiet:
            pr = QPrinter()
            if QPrintDialog(pr, self).exec():
                doc = QTextDocument()
                doc.setPlainText(text)
                doc.print(pr)
        self.vd.msgbar()
        return text


def make_index(td, spec):
    ix = td.call("CreateIndex", spec.name)
    for part in [p for p in spec.fields.split(";") if p.strip()]:
        p = part.strip()
        desc = p.startswith("-")
        f = ix.call("CreateField", p.lstrip("+-"))
        if desc:
            f.let("Attributes", 1)              # dbDescending
        ix.get("Fields").call("Append", f)
    for k, v in (("Primary", spec.primary), ("Unique", spec.unique or spec.primary), ("Required", spec.required),
                 ("IgnoreNulls", spec.ignorenulls)):
        ix.let(k, bool(v))
    return ix


class AddFieldDialog(QDialog):
    """Add Field: Name, Type, Size, FixedField / VariableField, AutoIncrField, AllowZeroLength, Required,
    OrdinalPosition, ValidationText, ValidationRule, DefaultValue; OK adds it and the dialog stays (Close ends)"""

    def __init__(self, ts, vd):
        super().__init__(ts)
        self.ts, self.vd = ts, vd
        self.setWindowTitle("Add Field")
        P = C.Place(self, 408, 249)             # (VB6 oracle vb_addfield.png)
        self.name, self.type, self.size = QLineEdit(), QComboBox(), QLineEdit("50")
        self.type.addItems(TYPES)
        self.fixed, self.var = QRadioButton("FixedField"), QRadioButton("VariableField")
        self.auto, self.azl, self.req = QCheckBox("AutoIncrField"), QCheckBox("AllowZeroLength"), QCheckBox("Required")
        self.ord, self.vtext, self.vrule, self.defval = QLineEdit(), QLineEdit(), QLineEdit(), QLineEdit()
        P.label(" Name: ", 8, 8, 37)
        P.put(self.name, 8, 24, 193, 19)
        P.label(" Type: ", 8, 48, 34)
        P.put(self.type, 8, 64, 113, 21)
        P.label(" Size: ", 8, 88, 29)
        P.put(self.size, 8, 104, 89, 19)
        P.put(self.fixed, 16, 128, 159, 17)
        P.put(self.var, 16, 144, 159, 17)
        P.put(self.auto, 16, 176, 159, 17)
        P.put(self.azl, 16, 200, 159, 17)
        P.put(self.req, 16, 224, 159, 17)
        P.label("OrdinalPosition: ", 208, 8, 78)
        P.put(self.ord, 208, 24, 89, 19)
        P.label("ValidationText: ", 208, 48, 75)
        P.put(self.vtext, 208, 64, 193, 19)
        P.label("ValidationRule: ", 208, 88, 74)
        P.put(self.vrule, 208, 112, 193, 19)
        P.label("DefaultValue: ", 208, 136, 68)
        P.put(self.defval, 208, 152, 193, 19)
        self.b_ok = P.put(QPushButton("&OK"), 232, 184, 145, 25)
        self.b_close = P.put(QPushButton("&Close"), 232, 216, 145, 25)
        self.b_ok.setDefault(True)
        self.b_ok.setEnabled(False)
        self.name.textChanged.connect(lambda t: self.b_ok.setEnabled(bool(t)))
        self.type.currentTextChanged.connect(self.type_changed)
        self.b_ok.clicked.connect(self.ok)
        self.b_close.clicked.connect(self.reject)
        self.defaults()

    def defaults(self):
        """SetDefaults: VariableField, AllowZeroLength, a Text of 50"""
        self.name.clear()
        self.fixed.setChecked(False)
        self.var.setChecked(True)
        self.auto.setChecked(False)
        self.req.setChecked(False)
        self.azl.setChecked(True)
        if self.type.currentText() == "Text":
            self.type_changed("Text")
        self.type.setCurrentText("Text")
        self.size.setText("50")
        for w in (self.vtext, self.vrule, self.defval):
            w.clear()

    def type_changed(self, t):
        """cboFieldType_Click: what a field of the type can have"""
        code = TYPE_CODES.get(t, 10)
        self.size.setText(str(SIZES.get(code, 0)) if code != 10 else "50")
        self.size.setEnabled(code == 10)
        self.azl.setEnabled(code in (10, 12))
        self.fixed.setEnabled(code == 10)
        self.var.setEnabled(code == 10)
        self.auto.setEnabled(code == 4)
        if code != 4:
            self.auto.setChecked(False)
        if code not in (10, 12):
            self.azl.setChecked(False)
        if code != 10:
            for w in (self.fixed, self.var):        # (no option chosen)
                w.setAutoExclusive(False)
                w.setChecked(False)
                w.setAutoExclusive(True)

    def spec(self):
        code = TYPE_CODES.get(self.type.currentText(), 10)
        attrs = (dbFixedField if self.fixed.isChecked() else 0) | (dbVariableField if self.var.isChecked() else 0) | \
            (dbAutoIncrField if self.auto.isChecked() and code == 4 else 0)
        try:
            size = int(self.size.text() or 0)
        except ValueError:
            size = 0
        return FieldSpec(self.name.text().strip(), code, size, attrs, self.req.isChecked(), self.azl.isChecked(),
                         self.defval.text(), self.vrule.text(), self.vtext.text())

    def ok(self):
        name = self.name.text().strip()
        if not name:
            return
        if name.upper() in (f.name.upper() for f in self.ts.fields):
            self.vd.say("'" + name + "'" + " Already exists!")
            self.name.selectAll()
            self.name.setFocus()
            return
        if self.ts.add_field(self.spec()):
            self.defaults()
            self.name.setFocus()


class AddIndexDialog(QDialog):
    """Add Index to <table>: Name, the Indexed Fields (a click on Available Fields adds one), Primary, Unique,
    IgnoreNulls"""

    def __init__(self, ts, vd, fields):
        super().__init__(ts)
        self.ts, self.vd = ts, vd
        self.setWindowTitle("Add Index to " + (ts.name.text() if ts.table is None else ts.table))
        P = C.Place(self, 360, 233)             # (VB6 oracle vb_addindex.png)
        self.name = QLineEdit()
        P.label("Name:", 8, 8, 34)
        P.put(self.name, 8, 24, 185, 19)
        self.indexed = QPlainTextEdit()
        P.label("Indexed Fields:", 8, 48, 77)
        P.put(self.indexed, 8, 64, 185, 35)
        self.avail = QListWidget()
        self.avail.setSortingEnabled(True)
        self.avail.addItems(fields)
        P.label("Available Fields:", 8, 104, 80)
        P.put(self.avail, 8, 120, 185, 95)
        self.primary, self.unique, self.ign = QCheckBox("Primary"), QCheckBox("Unique"), QCheckBox("IgnoreNulls")
        P.put(self.primary, 208, 24, 146, 17)
        P.put(self.unique, 208, 56, 146, 17)
        P.put(self.ign, 208, 88, 146, 17)
        self.primary.setChecked(True)
        self.unique.setChecked(True)
        self.b_ok = P.put(QPushButton("&OK"), 216, 160, 121, 25)
        self.b_close = P.put(QPushButton("&Close"), 216, 192, 121, 25)
        self.b_ok.setDefault(True)
        self.b_ok.setEnabled(False)
        self.avail.itemClicked.connect(self.pick)
        self.name.textChanged.connect(self.enable)
        self.indexed.textChanged.connect(self.enable)
        self.b_ok.clicked.connect(self.ok)
        self.b_close.clicked.connect(self.reject)

    def enable(self, *_):
        self.b_ok.setEnabled(bool(self.name.text()) and bool(self.indexed.toPlainText()))

    def pick(self, it):
        """lstFields_Click: the field added to the Indexed Fields (A;B)"""
        cur = self.indexed.toPlainText()
        self.indexed.setPlainText((cur + ";" if cur else "") + it.text())

    def ok(self):
        name = self.name.text().strip()
        fields = self.indexed.toPlainText().strip()
        if not name or not fields:
            return
        spec = IndexSpec(name, fields, self.primary.isChecked(), self.unique.isChecked(), False, self.ign.isChecked())
        if self.ts.add_index(spec):
            self.name.clear()                   # (SetDefaults)
            self.indexed.clear()
            self.unique.setChecked(True)
            self.name.setFocus()
