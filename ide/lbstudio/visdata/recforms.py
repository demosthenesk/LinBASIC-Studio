"""The forms of a Recordset in the Visual Data Manager: without a Data Control (VB6 DYNASNAP.FRM, TABLEOBJ.FRM: Add,
Edit, Delete, Close, Sort, Filter, Move, Find / Seek; Update, Cancel), with a Data Control (DATAFORM.FRM) and the grid
(DATAGRID.FRM). The rows are the Recordset's, through the runtime's DAO."""
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QPushButton, QLabel, QLineEdit,
                             QScrollArea, QScrollBar, QTableWidget, QTableWidgetItem, QComboBox, QDialog,
                             QRadioButton, QListWidget, QAbstractItemView)

from ..rtobj import RtError
from . import core as C

dbText, dbMemo = 10, 12
FIELD_WIDTH = {1: 850, 2: 650, 3: 900, 4: 1100, 5: 1800, 6: 1800, 7: 2200, 8: 2000}   # twips


def record_form(vd, rs, caption, table=False, source=""):
    ft = vd.prefs.get("formtype", C.FORM_NODATACTL)
    if ft == C.FORM_DATAGRID:
        w = GridForm(vd, rs, caption)
    else:
        w = RecordForm(vd, rs, caption, table=table, datactl=ft == C.FORM_DATACTL, source=source)
    vd.add_window(w)
    return w


class RecordForm(QWidget):
    """the fields of the current record, one under the other"""

    def __init__(self, vd, rs, caption, table=False, datactl=False, source=""):
        super().__init__()
        self.vd = vd
        self.rs = rs
        self.table = table
        self.datactl = datactl
        self.source = source
        self.editing = self.adding = False
        self.setWindowTitle(caption)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        lay.setSpacing(2)
        self.b = {}
        # -- the buttons
        self.view_buttons = QWidget()
        g = QGridLayout(self.view_buttons)
        g.setContentsMargins(0, 0, 0, 0)
        g.setSpacing(0)
        if datactl:
            names = [["&Add", "&Update", "&Delete", "&Find", "&Refresh", "&Close"]]
        elif table:
            names = [["&Add", "&Edit", "&Delete", "&Seek", "F&ilter", "&Close"]]
        else:
            names = [["&Add", "&Edit", "&Delete", "&Close"], ["&Sort", "F&ilter", "&Move", "&Find"]]
        for r, row in enumerate(names):
            for c, t in enumerate(row):
                b = QPushButton(t)
                g.addWidget(b, r, c)
                self.b[t.replace("&", "")] = b
        lay.addWidget(self.view_buttons)
        if table:
            row = QHBoxLayout()
            row.addWidget(QLabel("Index:"))
            self.indexes = QComboBox()
            row.addWidget(self.indexes, 1)
            lay.addLayout(row)
            try:                                # Name:Fields:Unique|Non-Unique[:Primary] (TABLEOBJ.FRM)
                from .dialogs import ix_fields
                cur = rs.get("Index")
                for x in vd.s.db.get("TableDefs", source).each("Indexes"):
                    t = x.get("Name") + ":" + ix_fields(x) + (":Unique" if x.get("Unique") else ":Non-Unique") + \
                        (":Primary" if x.get("Primary") else "")
                    self.indexes.addItem(t, x.get("Name"))
                    if x.get("Name") == cur:
                        self.indexes.setCurrentIndex(self.indexes.count() - 1)
            except RtError:
                pass
            self.indexes.currentIndexChanged.connect(lambda i: self.set_index(self.indexes.itemData(i)))
        self.change_buttons = QWidget()
        h = QHBoxLayout(self.change_buttons)
        h.setContentsMargins(0, 0, 0, 0)
        h.addStretch(1)
        for t in ("&Update", "&Cancel") if not datactl else ("C&ancel",):
            b = QPushButton(t)
            h.addWidget(b)
            self.b["ch:" + t.replace("&", "")] = b
        h.addStretch(1)
        self.change_buttons.hide()
        lay.addWidget(self.change_buttons)
        hdr = QHBoxLayout()
        self.hdr_name = QLabel("Field NAme:" if table else "Field Name:")   # (as VB6 writes it, vb_tableform.png)
        hdr.addWidget(self.hdr_name)
        hdr.addWidget(QLabel(" Value:" if datactl else (" Value  (F4=Zoom) " if table else " Value (F4=Zoom)")), 1)
        lay.addLayout(hdr)
        # -- the fields
        area = QScrollArea()
        area.setWidgetResizable(True)
        inner = QWidget()
        self.fg = QGridLayout(inner)
        self.fg.setContentsMargins(4, 0, 4, 0)
        area.setWidget(inner)
        lay.addWidget(area, 1)
        self.labels, self.edits = [], []
        try:
            self.fields = [(f.get("Name"), f.get("Type"), f.get("Size")) for f in rs.each("Fields")]
        except RtError as e:
            vd.show_error(e)
            self.fields = []
        for i, (n, t, sz) in enumerate(self.fields):
            lab = QLabel(n + ":")
            ed = QLineEdit()
            ed.setReadOnly(True)
            if t == dbText and sz:
                ed.setMaxLength(sz)
            if t in FIELD_WIDTH:                # GetFieldWidth (VISDATA.BAS): a number or a date is narrower
                ed.setFixedWidth(round(FIELD_WIDTH[t] / 15 * max(1.0, self.fontMetrics().horizontalAdvance(
                    "ValidationText: ") / 75.0)))
            ed.textEdited.connect(lambda _t, k=i: self._dirty(k))
            self.fg.addWidget(lab, i, 0)
            self.fg.addWidget(ed, i, 1, alignment=Qt.AlignmentFlag.AlignLeft if t in FIELD_WIDTH else Qt.AlignmentFlag(0))
            self.labels.append(lab)
            self.edits.append(ed)
        self.fg.setRowStretch(len(self.fields), 1)
        # Value over the values (lblFieldValue.Left = txtFieldData(0).Left)
        fm = self.fontMetrics()
        w0 = max([fm.horizontalAdvance(lb.text()) for lb in self.labels] + [fm.horizontalAdvance(self.hdr_name.text())])
        sp = self.fg.horizontalSpacing()
        self.hdr_name.setFixedWidth(w0 + self.fg.contentsMargins().left() + (sp if sp >= 0 else 6) + 2)
        self.dirty = set()
        # -- the position
        bottom = QHBoxLayout()
        if datactl or table:
            for t in ("|<", "<"):
                b = QPushButton(t)
                b.setFixedWidth(28)
                bottom.addWidget(b)
                self.b[t] = b
        else:
            self.hscroll = QScrollBar(Qt.Orientation.Horizontal)
            bottom.addWidget(self.hscroll, 2)
        self.status = QLabel("")
        bottom.addWidget(self.status, 1)
        if datactl or table:
            for t in (">", ">|"):
                b = QPushButton(t)
                b.setFixedWidth(28)
                bottom.addWidget(b)
                self.b[t] = b
        lay.addLayout(bottom)
        # -- the commands
        self._connect()
        try:
            if vd.s.datatype == C.MSACCESS and rs.get("Type") != C.dbOpenSnapshot:
                rs.let("LockEdits", True)
            if rs.get("RecordCount") > 0:
                rs.call("MoveLast")
                rs.call("MoveFirst")
            self.nrows = rs.get("RecordCount")
        except RtError as e:
            vd.show_error(e)
            self.nrows = 0
        self.set_scroll()
        self.display()
        self.resize(420, min(120 + 26 * len(self.fields), 480))

    def _connect(self):
        b = self.b
        b["Add"].clicked.connect(self.add)
        if "Edit" in b:
            b["Edit"].clicked.connect(self.edit)
        if "Update" in b:
            b["Update"].clicked.connect(self.update_dc)
        b["Delete"].clicked.connect(self.delete)
        b["Close"].clicked.connect(self.close_form)
        for k, fn in (("Sort", self.sort), ("Filter", self.filter), ("Move", self.move), ("Find", self.find),
                      ("Seek", self.seek), ("Refresh", self.refresh)):
            if k in b:
                b[k].clicked.connect(fn)
        for k, m in (("|<", "MoveFirst"), ("<", "MovePrevious"), (">", "MoveNext"), (">|", "MoveLast")):
            if k in b:
                b[k].clicked.connect(lambda _c=False, mm=m: self.nav(mm))
        if "ch:Update" in b:
            b["ch:Update"].clicked.connect(self.update)
        if "ch:Cancel" in b:
            b["ch:Cancel"].clicked.connect(self.cancel)
        if not self.datactl and not self.table:
            self.hscroll.valueChanged.connect(self.scrolled)

    # ---------------------------------------------------------------- showing
    def set_scroll(self):
        if hasattr(self, "hscroll"):
            self.hscroll.blockSignals(True)
            self.hscroll.setRange(0, max(0, self.nrows - 1))
            self.hscroll.setValue(0)
            self.hscroll.blockSignals(False)

    def display(self):
        """DisplayCurrentRecord: the values, "row/rows", (BOF) / (EOF), [Not Updatable]"""
        rs = self.rs
        try:
            noind = False
            if rs.get("Type") == C.dbOpenTable:
                try:
                    noind = self.vd.s.db.get("TableDefs", self.source).get("Indexes").get("Count") == 0
                except RtError:
                    noind = False
            cur = 0
            if rs.get("RecordCount") > 0:
                cur = int(round(self.nrows * (rs.get("PercentPosition") * 0.01))) + 1
            if self.adding:
                st = f"{self.nrows} Rows" if noind else f"{cur}/{self.nrows}"
            elif rs.get("BOF"):
                st = f"(BOF)/{self.nrows}"
                self.clear_fields()
            elif rs.get("EOF"):
                st = f"(EOF)/{self.nrows}"
                self.clear_fields()
            else:
                st = f"{self.nrows} Rows" if noind else f"{cur}/{self.nrows}"
                for i in range(len(self.fields)):
                    v = rs.get("Fields", i).get("Value")
                    self.edits[i].setText(C.text_of(v))
            if not rs.get("Updatable"):
                st += C.MSG[26]
            self.status.setText(st)
        except RtError as e:
            self.vd.show_error(e)
        self.dirty = set()

    def clear_fields(self):
        for e in self.edits:
            e.setText("")

    def _dirty(self, i):
        self.dirty.add(i)

    def set_editing(self, on):
        self.view_buttons.setVisible(not on)
        self.change_buttons.setVisible(on)
        for e in self.edits:
            e.setReadOnly(not on)
        if hasattr(self, "hscroll"):
            self.hscroll.setEnabled(not on)

    def write_fields(self):
        """the changed values into the record (txtFieldData_LostFocus)"""
        for i in sorted(self.dirty):
            t = self.edits[i].text()
            self.rs.get("Fields", i).let("Value", None if t == "" and self.fields[i][1] != dbText else t)
        self.dirty = set()

    # ---------------------------------------------------------------- the commands
    def add(self):
        try:
            self.rs.call("AddNew")
        except RtError as e:
            self.vd.show_error(e)
            return
        self.status.setText("Add record")
        self.adding = True
        self.clear_fields()
        self.set_editing(True)
        if self.edits:
            self.edits[0].setFocus()

    def edit(self):
        try:
            self.rs.call("Edit")
        except RtError as e:
            self.vd.show_error(e)
            return
        self.status.setText("Edit record")
        self.editing = True
        self.set_editing(True)

    def update(self):
        try:
            self.write_fields()
            self.rs.call("Update")
        except RtError as e:
            self.vd.show_error(e)
            return
        if self.vd.s.trans_pending:
            self.vd.s.changed = True
        if self.adding:
            self.nrows += 1
            self.set_scroll()
            try:
                self.rs.call("MoveLast")
            except RtError:
                pass
        self.editing = self.adding = False
        self.set_editing(False)
        self.display()

    def update_dc(self):
        """the Update of the Data Control form: the changes of the current record (it is always editable)"""
        try:
            if not self.adding:
                self.rs.call("Edit")
            self.write_fields()
            self.rs.call("Update")
        except RtError as e:
            self.vd.show_error(e)
            return
        self.adding = False
        self.set_editing(False)
        self.display()

    def cancel(self):
        try:
            self.rs.call("CancelUpdate")
        except RtError:
            pass
        self.editing = self.adding = False
        self.set_editing(False)
        self.display()

    def delete(self):
        if not self.vd.ask("Delete Current Record?"):
            return
        rs = self.rs
        try:
            rs.call("Delete")
            if self.vd.s.trans_pending:
                self.vd.s.changed = True
            if not rs.get("EOF"):
                rs.call("MoveNext")
                if rs.get("EOF") and rs.get("RecordCount") > 0:
                    rs.call("MoveLast")
            self.nrows = rs.get("RecordCount")
            self.set_scroll()
        except RtError as e:
            self.vd.show_error(e)
        self.display()

    def close_form(self):
        p = self.parent()
        while p is not None and not hasattr(p, "setWidget"):
            p = p.parent()
        (p or self).close()

    def nav(self, method):
        try:
            self.rs.call(method)
            if self.rs.get("EOF") and self.rs.get("RecordCount"):
                self.rs.call("MoveLast")
            elif self.rs.get("BOF") and self.rs.get("RecordCount"):
                self.rs.call("MoveFirst")
        except RtError as e:
            self.vd.show_error(e)
        self.display()

    def scrolled(self, v):
        try:
            self.rs.call("MoveFirst")
            if v:
                self.rs.call("Move", v)
        except RtError as e:
            self.vd.show_error(e)
        self.display()

    def _reopen(self, attr, value, msg):
        rs = self.rs
        try:
            if rs.get("RecordCount") == 0:
                return
            rs.let(attr, value)
            self.vd.msgbar(msg, True)
            rs2 = rs.call("OpenRecordset", rs.get("Type"))
            if rs2.get("RecordCount") > 0:
                rs2.call("MoveLast")
                rs2.call("MoveFirst")
            self.rs = rs2
            self.nrows = rs2.get("RecordCount")
            self.set_scroll()
        except RtError as e:
            self.vd.show_error(e)
        self.vd.msgbar()
        self.display()

    def sort(self):
        s = self.vd.ask_text("Enter Sort Column:")
        if s:
            self._reopen("Sort", s, "Setting New Sort Order")

    def filter(self):
        s = self.vd.ask_text("Enter Filter Expression:")
        if s:
            self._reopen("Filter", s, "Setting New Filter")

    def move(self):
        s = self.vd.ask_text("Enter number of Rows to Move:\n(Use negative value to move backwards)")
        if not s:
            return
        try:
            self.rs.call("Move", int(s))
            if self.rs.get("EOF"):
                self.rs.call("MoveLast")
            elif self.rs.get("BOF"):
                self.rs.call("MoveFirst")
        except (RtError, ValueError) as e:
            self.vd.show_error(e)
        self.display()

    def find(self):
        """Find: a field, an operator and a value (FIND.FRM); FindFirst, FindNext, FindPrevious or FindLast"""
        d = FindDialog(self, [f[0] for f in self.fields])
        vals = self.vd.answers.pop(0) if self.vd.quiet and self.vd.answers else None
        if vals is None:
            if self.vd.quiet or not d.exec():
                return
            vals = d.values()
        i, op, expr = vals[:3]
        how = vals[3] if len(vals) > 3 else "FindFirst"
        n, t, _sz = self.fields[i]
        name = n if " " not in n else f"[{n}]"
        crit = f"{name} {op} '{expr}'" if t in (dbText, dbMemo) else \
            (f"{name} {op} #{expr}#" if t == 8 else f"{name} {op} {expr}")
        try:
            bm = self.rs.get("Bookmark")
            self.rs.call(how, crit)
            if self.rs.get("NoMatch"):
                self.vd.say("Record Not Found", "warn")
                self.rs.let("Bookmark", bm)
        except RtError as e:
            self.vd.show_error(e)
        self.display()

    def seek(self):
        """Seek of a Table type Recordset (SEEK.FRM): an operator and the value of the index"""
        vals = self.vd.answers.pop(0) if self.vd.quiet and self.vd.answers else None
        if vals is None:
            d = SeekDialog(self)
            if self.vd.quiet or not d.exec():
                return
            vals = d.values()
        try:
            bm = self.rs.get("Bookmark")
            self.rs.call("Seek", vals[0], vals[1])
            if self.rs.get("NoMatch"):
                self.vd.say("Record Not Found", "warn")
                self.rs.let("Bookmark", bm)
        except RtError as e:
            self.vd.show_error(e)
        self.display()

    def set_index(self, name):
        try:
            self.rs.let("Index", name)
            self.rs.call("MoveFirst")
        except RtError as e:
            self.vd.show_error(e)
        self.display()

    def refresh(self):
        try:
            self.rs.call("Requery")
            self.nrows = self.rs.get("RecordCount")
        except RtError as e:
            self.vd.show_error(e)
        self.display()


class FindDialog(QDialog):
    """Find Record (FIND.FRM): Fields, Operators, Value or Expression; FindFirst / FindNext / FindPrevious /
    FindLast"""
    OPS = ["=", "<>", ">=", "<=", ">", "<", "Like"]
    HOW = ["FindFirst", "FindNext", "FindPrevious", "FindLast"]

    def __init__(self, parent, fields):
        super().__init__(parent)
        self.setWindowTitle("Find Record")
        P = C.Place(self, 363, 162)             # (VB6 oracle vb_find.png)
        self.lst, self.op, self.expr = QListWidget(), QListWidget(), QLineEdit()
        P.label("&Fields: ", 16, 8, 34, buddy=self.lst)
        P.put(self.lst, 16, 24, 113, 95)
        P.label("&Operators: ", 136, 8, 56, buddy=self.op)
        P.put(self.op, 136, 24, 64, 95)
        P.label("&Value or Expression: ", 208, 8, 101, buddy=self.expr)
        P.put(self.expr, 208, 24, 137, 19)
        self.lst.addItems(fields)
        self.lst.setCurrentRow(0)
        self.op.addItems(self.OPS)
        self.op.setCurrentRow(0)
        self.how = []
        for k, t in enumerate(self.HOW):
            r = P.put(QRadioButton(t), 208, 48 + 16 * k, 137, 17)
            self.how.append(r)
        self.how[0].setChecked(True)
        self.b_ok = P.put(QPushButton("&OK"), 40, 128, 113, 25)
        self.b_cancel = P.put(QPushButton("&Cancel"), 192, 128, 113, 25)
        self.b_ok.setDefault(True)
        self.b_ok.setEnabled(False)
        self.expr.textChanged.connect(lambda t: self.b_ok.setEnabled(bool(t)))
        self.b_ok.clicked.connect(self.accept)
        self.b_cancel.clicked.connect(self.reject)

    def values(self):
        how = next(t for r, t in zip(self.how, self.HOW) if r.isChecked())
        return self.lst.currentRow(), self.OPS[self.op.currentRow()], self.expr.text(), how


class SeekDialog(QDialog):
    """Seek Parameters (SEEK.FRM): the operator and the value of the index"""
    OPS = ["=", ">=", "<=", ">", "<"]

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Seek Parameters")
        P = C.Place(self, 226, 130)             # (VB6 oracle vb_seek.png)
        P.label("Operator: ", 8, 8, 51)
        self.ops = []
        for t, x, w in zip(self.OPS, (8, 48, 96, 144, 184), (33, 41, 41, 33, 33)):
            self.ops.append(P.put(QRadioButton(t), x, 24, w, 17))
        self.ops[0].setChecked(True)
        self.val = QLineEdit()
        P.label("&Value: ", 8, 48, 33, buddy=self.val)
        P.put(self.val, 8, 64, 209, 19)
        self.b_ok = P.put(QPushButton("&OK"), 8, 96, 97, 25)
        self.b_cancel = P.put(QPushButton("&Cancel"), 120, 96, 97, 25)
        self.b_ok.setDefault(True)
        self.b_ok.setEnabled(False)
        self.val.textChanged.connect(lambda t: self.b_ok.setEnabled(bool(t)))
        self.b_ok.clicked.connect(self.accept)
        self.b_cancel.clicked.connect(self.reject)

    def values(self):
        return next(t for r, t in zip(self.ops, self.OPS) if r.isChecked()), self.val.text()


class GridForm(QWidget):
    """DATAGRID.FRM: the rows in a grid (changed in place), the data control under it"""

    def __init__(self, vd, rs, caption):
        super().__init__()
        self.vd = vd
        self.rs = rs
        self.setWindowTitle(caption)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.grid = QTableWidget()
        self.grid.verticalHeader().setDefaultSectionSize(20)
        lay.addWidget(self.grid, 1)
        bar = QHBoxLayout()
        self.b = {}
        for t, m in (("|<", "MoveFirst"), ("<", "MovePrevious")):
            b = QPushButton(t)
            b.setFixedWidth(28)
            b.clicked.connect(lambda _c=False, r=m: self.nav(r))
            bar.addWidget(b)
            self.b[t] = b
        self.caption = QLabel("")
        bar.addWidget(self.caption, 1)
        for t, m in ((">", "MoveNext"), (">|", "MoveLast")):
            b = QPushButton(t)
            b.setFixedWidth(28)
            b.clicked.connect(lambda _c=False, r=m: self.nav(r))
            bar.addWidget(b)
            self.b[t] = b
        lay.addLayout(bar)
        self._filling = False
        self.load()
        self.grid.itemChanged.connect(self.changed)
        self.grid.currentCellChanged.connect(lambda r, c, pr, pc: self.go(r))
        self.resize(560, 360)

    def load(self):
        rs = self.rs
        self._filling = True
        try:
            self.fields = [(f.get("Name"), f.get("Type")) for f in rs.each("Fields")]
            self.grid.setColumnCount(len(self.fields))
            self.grid.setHorizontalHeaderLabels([f[0] for f in self.fields])
            rows = []
            if rs.get("RecordCount") > 0:
                rs.call("MoveFirst")
                while not rs.get("EOF"):
                    rows.append([C.text_of(rs.get("Fields", i).get("Value")) for i in range(len(self.fields))])
                    rs.call("MoveNext")
                rs.call("MoveFirst")
            self.grid.setRowCount(len(rows))
            for r, row in enumerate(rows):
                for c, v in enumerate(row):
                    self.grid.setItem(r, c, QTableWidgetItem(v))
            if not rs.get("Updatable"):
                self.grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        except RtError as e:
            self.vd.show_error(e)
        self._filling = False
        self.show_pos()

    def show_pos(self):
        try:
            self.caption.setText("" if self.rs.get("EOF") or self.rs.get("BOF") else
                                 f"{self.rs.get('AbsolutePosition') + 1}/{self.rs.get('RecordCount')}")
        except RtError:
            pass

    def go(self, r):
        if r < 0 or self._filling:
            return
        try:
            self.rs.let("AbsolutePosition", r)
        except RtError:
            pass
        self.show_pos()

    def nav(self, m):
        try:
            self.rs.call(m)
            if self.rs.get("EOF") and self.rs.get("RecordCount"):
                self.rs.call("MoveLast")
            if self.rs.get("BOF") and self.rs.get("RecordCount"):
                self.rs.call("MoveFirst")
            r = self.rs.get("AbsolutePosition")
            self._filling = True
            self.grid.setCurrentCell(max(0, r), max(0, self.grid.currentColumn()))
            self._filling = False
        except RtError as e:
            self.vd.show_error(e)
        self.show_pos()

    def changed(self, it):
        if self._filling:
            return
        try:
            self.rs.let("AbsolutePosition", it.row())
            self.rs.call("Edit")
            t = it.text()
            self.rs.get("Fields", it.column()).let("Value", None if t == "" and self.fields[it.column()][1] != dbText
                                                   else t)
            self.rs.call("Update")
            if self.vd.s.trans_pending:
                self.vd.s.changed = True
        except RtError as e:
            self.vd.show_error(e)
            try:
                self.rs.call("CancelUpdate")
            except RtError:
                pass
