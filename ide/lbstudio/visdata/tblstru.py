"""Table Structure of the Visual Data Manager (VB6 TBLSTRU.FRM, ADDFIELD.FRM, ADDINDEX.FRM): the fields and indexes of
a table (Design...) or of a new one (New Table: Build the Table), Add Field, Remove Field, Add Index, Remove Index,
Print Structure."""
from PyQt6.QtWidgets import (QDialog, QGridLayout, QLabel, QLineEdit, QListWidget, QPushButton, QComboBox, QCheckBox,
                             QHBoxLayout, QVBoxLayout, QRadioButton, QGroupBox, QPlainTextEdit)

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
        g = QGridLayout(self)
        g.addWidget(QLabel("Table &Name:"), 0, 0)
        self.name = QLineEdit(table or "")
        g.addWidget(self.name, 0, 1, 1, 3)
        g.addWidget(QLabel("&Field List:"), 1, 0)
        self.lst_fields = QListWidget()
        g.addWidget(self.lst_fields, 2, 0, 6, 1)
        # the properties of the field chosen
        props = QGridLayout()
        self.f_name, self.f_type, self.f_size = QLineEdit(), QComboBox(), QLineEdit()
        self.f_type.addItems(TYPES)
        self.f_coll, self.f_ord, self.f_vtext, self.f_vrule, self.f_def = (QLineEdit() for _ in range(5))
        self.f_fixed, self.f_var, self.f_auto = QCheckBox("FixedLength"), QCheckBox("VariableLength"), \
            QCheckBox("AutoIncrement")
        self.f_azl, self.f_req = QCheckBox("AllowZeroLength"), QCheckBox("Required")
        rows = (("Name: ", self.f_name), ("Type: ", self.f_type), ("Size: ", self.f_size),
                ("CollatingOrder: ", self.f_coll), ("OrdinalPosition: ", self.f_ord),
                ("ValidationText: ", self.f_vtext), ("ValidationRule: ", self.f_vrule),
                ("DefaultValue: ", self.f_def))
        for i, (t, w) in enumerate(rows):
            props.addWidget(QLabel(t), i, 0)
            props.addWidget(w, i, 1)
        checks = QVBoxLayout()
        for w in (self.f_fixed, self.f_var, self.f_auto, self.f_azl, self.f_req):
            checks.addWidget(w)
        props.addLayout(checks, 0, 2, 8, 1)
        g.addLayout(props, 1, 1, 7, 3)
        for w in (self.f_name, self.f_size, self.f_coll, self.f_ord, self.f_vtext, self.f_vrule, self.f_def):
            w.setReadOnly(table is not None)
        row = QHBoxLayout()
        self.b_addfield = QPushButton("&Add Field")
        self.b_rmfield = QPushButton("&Remove Field")
        row.addWidget(self.b_addfield)
        row.addWidget(self.b_rmfield)
        g.addLayout(row, 8, 0)
        g.addWidget(QLabel("Inde&x List:"), 9, 0)
        self.lst_indexes = QListWidget()
        g.addWidget(self.lst_indexes, 10, 0, 3, 1)
        ip = QGridLayout()
        self.i_name, self.i_fields = QLineEdit(), QLineEdit()
        self.i_name.setReadOnly(True)
        self.i_fields.setReadOnly(True)
        ip.addWidget(QLabel("Name: "), 0, 0)
        ip.addWidget(self.i_name, 0, 1, 1, 4)
        ip.addWidget(QLabel("Fields: "), 1, 0)
        ip.addWidget(self.i_fields, 1, 1, 1, 4)
        self.i_primary, self.i_unique, self.i_foreign, self.i_req, self.i_ign = (
            QCheckBox(t) for t in ("Primary", "Unique", "Foreign", "Required", "IgnoreNull"))
        for k, w in enumerate((self.i_primary, self.i_unique, self.i_foreign, self.i_req, self.i_ign)):
            w.setEnabled(False)
            ip.addWidget(w, 2, k)
        g.addLayout(ip, 10, 1, 3, 3)
        row = QHBoxLayout()
        self.b_addindex = QPushButton("Add &Index")
        self.b_rmindex = QPushButton("Re&move Index")
        row.addWidget(self.b_addindex)
        row.addWidget(self.b_rmindex)
        g.addLayout(row, 13, 0)
        row = QHBoxLayout()
        self.b_build = QPushButton("&Build the Table")
        self.b_print = QPushButton("&Print Structure")
        self.b_close = QPushButton("&Close")
        for b in (self.b_build, self.b_close, self.b_print):
            row.addWidget(b)
        g.addLayout(row, 14, 0, 1, 4)
        self.b_build.setVisible(table is None)
        self.b_print.setVisible(table is not None)
        self.name.setReadOnly(False)
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

    def show_field(self, i):
        if not 0 <= i < len(self.fields):
            return
        f = self.fields[i]
        self.f_name.setText(f.name)
        self.f_type.setCurrentText(CODE_NAMES.get(f.type, str(f.type)))
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
        g = QGridLayout(self)
        self.name, self.type, self.size = QLineEdit(), QComboBox(), QLineEdit("50")
        self.type.addItems(TYPES)
        self.type.setCurrentText("Text")
        self.type.currentTextChanged.connect(self.type_changed)
        self.fixed, self.var = QRadioButton("FixedField"), QRadioButton("VariableField")
        self.var.setChecked(True)
        self.auto, self.azl, self.req = QCheckBox("AutoIncrField"), QCheckBox("AllowZeroLength"), QCheckBox("Required")
        self.ord, self.vtext, self.vrule, self.defval = QLineEdit(), QLineEdit(), QLineEdit(), QLineEdit()
        for i, (t, w) in enumerate(((" Name: ", self.name), (" Type: ", self.type), (" Size: ", self.size))):
            g.addWidget(QLabel(t), i, 0)
            g.addWidget(w, i, 1)
        box = QGroupBox()
        bv = QVBoxLayout(box)
        for w in (self.fixed, self.var, self.auto, self.azl, self.req):
            bv.addWidget(w)
        g.addWidget(box, 0, 2, 4, 1)
        for i, (t, w) in enumerate((("OrdinalPosition: ", self.ord), ("ValidationText: ", self.vtext),
                                    ("ValidationRule: ", self.vrule), ("DefaultValue: ", self.defval))):
            g.addWidget(QLabel(t), 3 + i, 0)
            g.addWidget(w, 3 + i, 1)
        row = QVBoxLayout()
        ok, close = QPushButton("&OK"), QPushButton("&Close")
        row.addWidget(ok)
        row.addWidget(close)
        row.addStretch(1)
        g.addLayout(row, 4, 2, 3, 1)
        ok.clicked.connect(self.ok)
        close.clicked.connect(self.reject)
        self.type_changed("Text")

    def type_changed(self, t):
        code = TYPE_CODES.get(t, 10)
        self.size.setText(str(SIZES.get(code, 0)))
        self.size.setEnabled(code == 10)
        self.azl.setEnabled(code in (10, 12))
        self.auto.setEnabled(code == 4)
        self.fixed.setEnabled(code == 10)
        self.var.setEnabled(code == 10)
        self.fixed.setChecked(code != 10)
        self.var.setChecked(code == 10)

    def spec(self):
        code = TYPE_CODES.get(self.type.currentText(), 10)
        attrs = (dbFixedField if self.fixed.isChecked() else dbVariableField if code == 10 else 0) | \
            (dbAutoIncrField if self.auto.isChecked() and code == 4 else 0)
        try:
            size = int(self.size.text() or 0)
        except ValueError:
            size = 0
        return FieldSpec(self.name.text().strip(), code, size, attrs, self.req.isChecked(), self.azl.isChecked(),
                         self.defval.text(), self.vrule.text(), self.vtext.text())

    def ok(self):
        if not self.name.text().strip():
            return
        if self.ts.add_field(self.spec()):
            self.name.clear()
            self.name.setFocus()


class AddIndexDialog(QDialog):
    """Add Index: Name, the Indexed Fields (from Available Fields), Primary, Unique, IgnoreNulls"""

    def __init__(self, ts, vd, fields):
        super().__init__(ts)
        self.ts, self.vd = ts, vd
        self.setWindowTitle("Add Index")
        g = QGridLayout(self)
        g.addWidget(QLabel("Name: "), 0, 0)
        self.name = QLineEdit()
        g.addWidget(self.name, 0, 1)
        g.addWidget(QLabel("Indexed  Fields:"), 1, 0, 1, 2)
        self.indexed = QPlainTextEdit()
        self.indexed.setMaximumHeight(50)
        g.addWidget(self.indexed, 2, 0, 1, 2)
        g.addWidget(QLabel("Available Fields: "), 3, 0, 1, 2)
        self.avail = QListWidget()
        self.avail.addItems(fields)
        self.avail.itemClicked.connect(self.pick)
        g.addWidget(self.avail, 4, 0, 1, 2)
        side = QVBoxLayout()
        self.primary, self.unique, self.ign = QCheckBox("Primary"), QCheckBox("Unique"), QCheckBox("IgnoreNulls")
        self.primary.setChecked(True)
        self.unique.setChecked(True)
        ok, close = QPushButton("&OK"), QPushButton("&Close")
        for w in (ok, close, self.primary, self.unique, self.ign):
            side.addWidget(w)
        side.addStretch(1)
        g.addLayout(side, 0, 2, 5, 1)
        ok.clicked.connect(self.ok)
        close.clicked.connect(self.reject)

    def pick(self, it):
        cur = self.indexed.toPlainText().strip()
        self.indexed.setPlainText((cur + ";" if cur else "") + "+" + it.text())

    def ok(self):
        name = self.name.text().strip()
        fields = self.indexed.toPlainText().strip()
        if not name or not fields:
            return
        spec = IndexSpec(name, fields, self.primary.isChecked(), self.unique.isChecked(), False, self.ign.isChecked())
        if self.ts.add_index(spec):
            self.name.clear()
            self.indexed.clear()
