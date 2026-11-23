"""The designers of the .dsr files: the DataEnvironment (its connections and commands, the hierarchy
of the commands: relations, grouping, aggregates; the fields as ADO gives them) and the DataReport
(its sections, the controls on them, the DataSource / DataMember, Retrieve Structure), as VB6 has
them. The objects are in dsrmodel; every change goes back into the .dsr as VB6 writes it."""
import os
import re

from PyQt6.QtCore import Qt, QRect, pyqtSignal
from PyQt6.QtGui import QColor, QFont, QPainter, QPen
from PyQt6.QtWidgets import (QWidget, QHBoxLayout, QVBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem, QToolBar,
                             QStackedWidget, QFormLayout, QLineEdit, QComboBox, QCheckBox, QPushButton, QLabel,
                             QTabWidget, QPlainTextEdit, QRadioButton, QListWidget, QTableWidget, QTableWidgetItem,
                             QMessageBox, QFileDialog, QScrollArea, QSpinBox, QToolButton, QMenu, QButtonGroup,
                             QAbstractItemView, QHeaderView, QGridLayout)

from . import dsrmodel as M
from . import dbinfo

NAME_RE = re.compile(r"[A-Za-z][A-Za-z0-9_]{0,39}$")


class _Base(QWidget):
    modified = pyqtSignal(object)
    selection_changed = pyqtSignal(object, object)       # not used: the Properties window shows the module

    def __init__(self, main, module):
        super().__init__()
        self.main = main
        self.module = module
        self.status = QLabel("")
        self.status.setStyleSheet("color: #a00000")

    def changed(self):
        self.model.name = self.module.name
        self.model.apply()
        self.module.dirty = True
        self.modified.emit(self.module)

    def say(self, text=""):
        self.status.setText(text)

    def project_dir(self):
        p = self.module.project.path
        return os.path.dirname(p) if p else os.getcwd()


# ====================================================================== DataEnvironment
def design_cs(cs, base_dir):
    """the connection string as this computer can open it at design time: a database file named
    with a Windows path or a relative one is looked for in the folder of the project"""
    m = re.search(r"(Data Source|DBQ)\s*=\s*([^;]*)", cs, re.IGNORECASE)
    if not m:
        return cs
    src = m.group(2).strip().strip('"')
    if os.path.isfile(src):
        return cs
    for cand in (os.path.join(base_dir, src.replace("\\", "/")), os.path.join(base_dir, src.replace("\\", "/").split("/")[-1])):
        if os.path.isfile(cand):
            return cs[:m.start(2)] + cand + cs[m.end(2):]
    return cs


def is_action(sql):
    """a statement that changes the database (INSERT, UPDATE, DELETE, a table made or dropped, SELECT ... INTO)"""
    import re
    t = re.sub(r"'[^']*'|\"[^\"]*\"|\[[^\]]*\]", " ", sql or "").strip().lower()
    first = (t.split() or [""])[0]
    if first in ("insert", "update", "delete", "create", "drop", "alter", "grant", "revoke", "truncate", "exec",
                 "execute", "call", "merge", "replace"):
        return True
    return first == "select" and re.search(r"\binto\b", t.split(" from ")[0]) is not None


def without_params(sql):
    """the ? of a statement as NULL: its fields at design time, without values for its parameters"""
    out, q = [], None
    for ch in sql:
        if q:
            if ch == q:
                q = None
        elif ch in "'\"":
            q = ch
        elif ch == "?":
            ch = "NULL"
        out.append(ch)
    return "".join(out)


CMD_TYPES = [(1, "SQL Statement"), (2, "Table"), (4, "Stored Procedure"), (8, "Unknown")]


class ConnectionPage(QWidget):
    def __init__(self, d):
        super().__init__()
        self.d = d
        self.conn = None
        f = QFormLayout(self)
        self.name = QLineEdit()
        f.addRow("&Name:", self.name)
        h = QHBoxLayout()
        self.cs = QLineEdit()
        self.cs.setMinimumWidth(320)
        h.addWidget(self.cs, 1)
        b = QPushButton("&Build...")
        b.clicked.connect(self.build)
        h.addWidget(b)
        f.addRow("&Connection String:", h)
        h = QHBoxLayout()
        b = QPushButton("&Test Connection")
        b.clicked.connect(self.test)
        h.addWidget(b)
        b = QPushButton("&Apply")
        b.clicked.connect(self.apply)
        h.addWidget(b)
        h.addStretch(1)
        f.addRow("", h)
        f.addRow("", QLabel("Jet: Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<file.mdb>\n"
                            "ODBC: Provider=MSDASQL.1;Data Source=<DSN> or Driver={...};..."))

    def show_conn(self, c):
        self.conn = c
        self.name.setText(c.name)
        self.cs.setText(c.source)

    def build(self):
        path, _ = QFileDialog.getOpenFileName(self, "Data Link Properties - Database", self.d.project_dir(),
                                              "Access Databases (*.mdb *.MDB);;All Files (*)")
        if path:
            self.cs.setText(f"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={path};Persist Security Info=False")

    def test(self):
        try:
            n = len(dbinfo.ado_objects(design_cs(self.cs.text(), self.d.project_dir())))
            QMessageBox.information(self, "Microsoft Data Link", f"Test connection succeeded ({n} tables).")
        except dbinfo.AdoError as e:
            QMessageBox.warning(self, "Microsoft Data Link", f"Test connection failed:\n{e}")

    def apply(self):
        c, d = self.conn, self.d
        new = self.name.text().strip()
        if new != c.name:
            if not NAME_RE.match(new):
                return d.say("Invalid name")
            if any(x is not c and x.name.lower() == new.lower() for x in d.model.conns):
                return d.say(f"The name '{new}' is already used")
            for cmd in d.model.cmds:
                if cmd.conn.lower() == c.name.lower():
                    cmd.conn = new
            c.name = new
        c.source = self.cs.text().strip()
        dbinfo.forget()
        d.say()
        d.changed()
        d.fill_tree(c)


class CommandPage(QWidget):
    """the Properties of a command, as the property pages of VB6 (General, Relation, Grouping,
    Aggregates)"""

    def __init__(self, d):
        super().__init__()
        self.d = d
        self.cmd = None
        v = QVBoxLayout(self)
        self.tabs = QTabWidget()
        v.addWidget(self.tabs, 1)
        # -- General
        g = QWidget()
        f = QFormLayout(g)
        self.name = QLineEdit()
        f.addRow("Command &Name:", self.name)
        self.conn = QComboBox()
        f.addRow("&Connection:", self.conn)
        self.rb_obj = QRadioButton("&Database Object:")
        self.rb_sql = QRadioButton("&SQL Statement:")
        bg = QButtonGroup(self)
        bg.addButton(self.rb_obj)
        bg.addButton(self.rb_sql)
        self.objtype = QComboBox()
        self.objtype.addItems(["Table", "Stored Procedure", "View"])
        f.addRow(self.rb_obj, self.objtype)
        self.objname = QComboBox()
        self.objname.setEditable(True)
        f.addRow("&Object Name:", self.objname)
        self.sql = QPlainTextEdit()
        self.sql.setMaximumHeight(90)
        f.addRow(self.rb_sql, self.sql)
        self.rb_obj.toggled.connect(self._source_kind)
        self.objtype.currentIndexChanged.connect(lambda _: self._fill_objects())
        self.conn.currentIndexChanged.connect(lambda _: self._fill_objects())
        self.tabs.addTab(g, "General")
        # -- Parameters
        pw = QWidget()
        pv = QVBoxLayout(pw)
        self.params = QTableWidget(0, 6)
        self.params.setHorizontalHeaderLabels(["Name", "Direction", "Data Type", "Host Data Type", "Size", "Required"])
        self.params.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.params.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        pv.addWidget(self.params, 1)
        h = QHBoxLayout()
        for text, fn in (("&Add", lambda: self._param_row(M.Param("Param%d" % (self.params.rowCount() + 1)))),
                         ("&Remove", lambda: self.params.removeRow(self.params.currentRow())
                          if self.params.currentRow() >= 0 else None),
                         ("From the &SQL (?)", self._params_from_sql)):
            b = QPushButton(text)
            b.clicked.connect(fn)
            h.addWidget(b)
        h.addStretch(1)
        pv.addLayout(h)
        self.tabs.addTab(pw, "Parameters")
        # -- Relation
        r = QWidget()
        rv = QVBoxLayout(r)
        self.relate = QCheckBox("&Relate to a Parent Command Object")
        rv.addWidget(self.relate)
        h = QHBoxLayout()
        lab = QLabel("&Parent Command:")
        h.addWidget(lab)
        self.parent = QComboBox()
        lab.setBuddy(self.parent)
        h.addWidget(self.parent, 1)
        rv.addLayout(h)
        grid = QGridLayout()
        grid.addWidget(QLabel("Parent Fields"), 0, 0)
        grid.addWidget(QLabel("Child Fields/Parameters"), 0, 1)
        self.pfield = QComboBox()
        self.cfield = QComboBox()
        grid.addWidget(self.pfield, 1, 0)
        grid.addWidget(self.cfield, 1, 1)
        b = QPushButton("&Add")
        b.clicked.connect(self._rel_add)
        grid.addWidget(b, 1, 2)
        rv.addLayout(grid)
        self.rels = QTableWidget(0, 2)
        self.rels.setHorizontalHeaderLabels(["Parent Fields", "Child Fields/Parameters"])
        self.rels.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.rels.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.rels.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        rv.addWidget(self.rels, 1)
        b = QPushButton("Re&move")
        b.clicked.connect(lambda: self.rels.removeRow(self.rels.currentRow()) if self.rels.currentRow() >= 0 else None)
        rv.addWidget(b, 0, Qt.AlignmentFlag.AlignLeft)
        self.relate.toggled.connect(self._relate_toggled)
        self.parent.currentIndexChanged.connect(lambda _: self._fill_pfields())
        self.tabs.addTab(r, "Relation")
        # -- Grouping
        gp = QWidget()
        gv = QVBoxLayout(gp)
        self.grouping = QCheckBox("&Group Command Object")
        gv.addWidget(self.grouping)
        h = QHBoxLayout()
        lab = QLabel("Grouping &Command Name:")
        h.addWidget(lab)
        self.gname = QLineEdit()
        lab.setBuddy(self.gname)
        h.addWidget(self.gname, 1)
        gv.addLayout(h)
        h = QHBoxLayout()
        self.gfields = QListWidget()
        self.gused = QListWidget()
        h.addWidget(self.gfields, 1)
        bv = QVBoxLayout()
        for text, fn in ((">", self._g_add), ("<", self._g_remove)):
            b = QPushButton(text)
            b.setFixedWidth(36)
            b.clicked.connect(fn)
            bv.addWidget(b)
        bv.addStretch(1)
        h.addLayout(bv)
        h.addWidget(self.gused, 1)
        gv.addWidget(QLabel("Fields in Command:                         Fields Used for Grouping:"))
        gv.addLayout(h, 1)
        self.tabs.addTab(gp, "Grouping")
        # -- Aggregates
        ag = QWidget()
        av = QVBoxLayout(ag)
        self.aggs = QTableWidget(0, 4)
        self.aggs.setHorizontalHeaderLabels(["Name", "Function", "Aggregate On", "Field"])
        self.aggs.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.aggs.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        av.addWidget(self.aggs, 1)
        h = QHBoxLayout()
        b = QPushButton("&Add")
        b.clicked.connect(lambda: self._agg_row(("Aggregate%d" % (self.aggs.rowCount() + 1), "", "", 7)))
        h.addWidget(b)
        b = QPushButton("&Remove")
        b.clicked.connect(lambda: self.aggs.removeRow(self.aggs.currentRow()) if self.aggs.currentRow() >= 0 else None)
        h.addWidget(b)
        h.addStretch(1)
        av.addLayout(h)
        self.tabs.addTab(ag, "Aggregates")
        h = QHBoxLayout()
        h.addStretch(1)
        b = QPushButton("&Apply")
        b.clicked.connect(self.apply)
        h.addWidget(b)
        v.addLayout(h)

    # ------------------------------------------------------------------ filling
    def _cs(self):
        c = next((x for x in self.d.model.conns if x.name == self.conn.currentText()), None)
        return design_cs(c.source, self.d.project_dir()) if c is not None and c.source else None

    def _fill_objects(self):
        keep = self.objname.currentText()
        self.objname.clear()
        cs = self._cs()
        if cs:
            try:
                want = "Q" if self.objtype.currentText() == "View" else "T"
                self.objname.addItems([n for k, n in dbinfo.ado_objects(cs) if k == want])
            except dbinfo.AdoError as e:
                self.d.say(str(e))
        self.objname.setEditText(keep)

    def _source_kind(self, obj):
        self.objtype.setEnabled(obj)
        self.objname.setEnabled(obj)
        self.sql.setEnabled(not obj)

    def _relate_toggled(self, on):
        for w in (self.parent, self.pfield, self.cfield, self.rels):
            w.setEnabled(on)

    def _fill_pfields(self):
        self.pfield.clear()
        p = self.d.model.command(self.parent.currentText())
        if p is not None:
            self.pfield.addItems([f[0] for f in p.fields])

    def _rel_add(self):
        p, c = self.pfield.currentText(), self.cfield.currentText()
        if p and c:
            r = self.rels.rowCount()
            self.rels.insertRow(r)
            self.rels.setItem(r, 0, QTableWidgetItem(p))
            self.rels.setItem(r, 1, QTableWidgetItem(c))

    def _g_add(self):
        it = self.gfields.currentItem()
        if it is not None:
            self.gused.addItem(it.text())
            self.gfields.takeItem(self.gfields.row(it))

    def _g_remove(self):
        it = self.gused.currentItem()
        if it is not None:
            self.gfields.addItem(it.text())
            self.gused.takeItem(self.gused.row(it))

    DIRS = ["0 - adParamUnknown", "1 - adParamInput", "2 - adParamOutput", "3 - adParamInputOutput",
            "4 - adParamReturnValue"]

    def _param_row(self, prm):
        r = self.params.rowCount()
        self.params.insertRow(r)
        it = QTableWidgetItem(prm.name)
        it.setData(Qt.ItemDataRole.UserRole, prm)
        self.params.setItem(r, 0, it)
        cb = QComboBox()
        cb.addItems(self.DIRS)
        cb.setCurrentIndex(int(prm.v["Direction"] or 0) if 0 <= int(prm.v["Direction"] or 0) <= 4 else 1)
        self.params.setCellWidget(r, 1, cb)
        for col, table, key in ((2, M.PARAM_TYPES, "DataType"), (3, M.HOST_TYPES, "HostType")):
            cb = QComboBox()
            for n, label in table:
                cb.addItem(label, n)
            i = cb.findData(int(prm.v[key] or 0))
            if i < 0:
                cb.addItem(str(prm.v[key]), int(prm.v[key] or 0))
                i = cb.count() - 1
            cb.setCurrentIndex(i)
            if col == 2:                        # a new data type: its host type with it
                cb.activated.connect(lambda _, r=r, cb=cb: self._host_follows(r, cb.currentData()))
            self.params.setCellWidget(r, col, cb)
        self.params.setItem(r, 4, QTableWidgetItem(str(prm.v["Size"])))
        ck = QCheckBox()
        ck.setChecked(bool(prm.v["Required"]))
        self.params.setCellWidget(r, 5, ck)

    def _host_follows(self, r, datatype):
        cb = self.params.cellWidget(r, 3)
        i = cb.findData(M.HOST_OF_TYPE.get(datatype, 12))
        if i >= 0:
            cb.setCurrentIndex(i)

    def _read_params(self):
        out = []
        for r in range(self.params.rowCount()):
            it = self.params.item(r, 0)
            prm = it.data(Qt.ItemDataRole.UserRole) or M.Param()
            name = it.text().strip() or f"Param{r + 1}"
            if name != prm.name:
                prm.v["UserName"] = name
                if not prm.v["RealName"] or prm.v["RealName"] == prm.name:
                    prm.v["RealName"] = name
            prm.v["Direction"] = self.params.cellWidget(r, 1).currentIndex()
            prm.v["DataType"] = self.params.cellWidget(r, 2).currentData()
            prm.v["HostType"] = self.params.cellWidget(r, 3).currentData()
            try:
                prm.v["Size"] = int(self.params.item(r, 4).text())
            except (ValueError, AttributeError):
                pass
            prm.v["Required"] = self.params.cellWidget(r, 5).isChecked()
            out.append(prm)
        return out

    def _params_from_sql(self):
        """one parameter for every ? of the SQL statement (outside its strings), as the provider tells
        the designer of VB6"""
        n, q = 0, None
        for ch in self.sql.toPlainText():
            if q:
                if ch == q:
                    q = None
            elif ch in "'\"":
                q = ch
            elif ch == "?":
                n += 1
        while self.params.rowCount() < n:
            self._param_row(M.Param("Param%d" % (self.params.rowCount() + 1)))
        while self.params.rowCount() > n:
            self.params.removeRow(self.params.rowCount() - 1)

    def _agg_row(self, a):
        name, on, field, fn = a
        r = self.aggs.rowCount()
        self.aggs.insertRow(r)
        self.aggs.setItem(r, 0, QTableWidgetItem(name))
        cb = QComboBox()
        cb.addItems(M.AGG_FUNCS)
        cb.setCurrentIndex(fn if 0 <= fn < len(M.AGG_FUNCS) else 7)
        self.aggs.setCellWidget(r, 1, cb)
        on_cb = QComboBox()
        on_cb.setEditable(True)
        on_cb.addItems(self._agg_targets())
        on_cb.setEditText(on or (self._agg_targets() or [""])[0])
        self.aggs.setCellWidget(r, 2, on_cb)
        f_cb = QComboBox()
        f_cb.setEditable(True)
        self._agg_fields(f_cb, on_cb.currentText())
        f_cb.setEditText(field)
        on_cb.currentTextChanged.connect(lambda t, f_cb=f_cb: self._agg_fields(f_cb, t))
        self.aggs.setCellWidget(r, 3, f_cb)

    def _agg_targets(self):
        """Aggregate On: the grouping of the command, or one of the commands below it"""
        c = self.cmd
        out = [c.member] if c is not None and c.grouping else []
        stack = list(self.d.model.children(c)) if c is not None else []
        while stack:
            x = stack.pop(0)
            out.append(x.name)
            stack.extend(self.d.model.children(x))
        return out

    def _agg_fields(self, cb, target):
        keep = cb.currentText()
        cb.clear()
        c = self.d.model.command(target) if target else None
        if c is None and self.cmd is not None and target == self.cmd.member:
            c = self.cmd
        if c is not None:
            cb.addItems([f[0] for f in c.fields])
        cb.setEditText(keep)

    def show_cmd(self, c):
        d = self.d
        self.cmd = c
        self.name.setText(c.name)
        self.conn.blockSignals(True)
        self.conn.clear()
        self.conn.addItems([x.name for x in d.model.conns])
        self.conn.setCurrentText(c.conn)
        self.conn.blockSignals(False)
        obj = c.ctype in (2, 4)
        self.rb_obj.setChecked(obj)
        self.rb_sql.setChecked(not obj)
        self._source_kind(obj)
        self.objtype.blockSignals(True)
        self.objtype.setCurrentIndex(1 if c.ctype == 4 else 0)
        self.objtype.blockSignals(False)
        self._fill_objects()
        self.objname.setEditText(c.text if obj else "")
        self.sql.setPlainText(c.text if not obj else "")
        # relation
        self.relate.setChecked(bool(c.parent))
        self.parent.blockSignals(True)
        self.parent.clear()
        below = set()
        stack = [c]
        while stack:
            x = stack.pop()
            below.add(x.name.lower())
            stack.extend(d.model.children(x))
        self.parent.addItems([x.name for x in d.model.cmds if x.name.lower() not in below])
        self.parent.setCurrentText(c.parent)
        self.parent.blockSignals(False)
        self._fill_pfields()
        self.cfield.clear()
        self.cfield.addItems([f[0] for f in c.fields] + [p.name for p in c.params])
        self.rels.setRowCount(0)
        for p, ch in c.relations:
            self.pfield.setCurrentText(p)
            r = self.rels.rowCount()
            self.rels.insertRow(r)
            self.rels.setItem(r, 0, QTableWidgetItem(p))
            self.rels.setItem(r, 1, QTableWidgetItem(ch))
        self._relate_toggled(bool(c.parent))
        # grouping
        self.grouping.setChecked(c.grouping)
        self.gname.setText(c.member if c.grouping else c.name + "_Grouping")
        self.gused.clear()
        self.gused.addItems(c.groups)
        self.gfields.clear()
        self.gfields.addItems([f[0] for f in c.fields if f[0] not in c.groups])
        self.tabs.setTabEnabled(3, not c.parent)
        # parameters
        self.params.setRowCount(0)
        for prm in c.params:
            self._param_row(prm)
        # aggregates
        self.aggs.setRowCount(0)
        for a in c.aggregates:
            self._agg_row(a)

    # ------------------------------------------------------------------ apply
    def apply(self):
        d, c = self.d, self.cmd
        new = self.name.text().strip()
        if not NAME_RE.match(new):
            return d.say("Invalid command name")
        if any(x is not c and x.name.lower() == new.lower() for x in d.model.cmds):
            return d.say(f"The name '{new}' is already used")
        if new != c.name:
            for x in d.model.cmds:
                if x.parent.lower() == c.name.lower():
                    x.parent = new
                x.aggregates = [(n, new if on.lower() == c.name.lower() else on, f, fn) for n, on, f, fn in x.aggregates]
            c.name = new
        c.conn = self.conn.currentText()
        if self.rb_obj.isChecked():
            c.ctype = 4 if self.objtype.currentText() == "Stored Procedure" else 2
            c.text = self.objname.currentText().strip()
        else:
            c.ctype = 1
            c.text = self.sql.toPlainText().strip()
        if self.relate.isChecked() and self.parent.currentText():
            c.parent = self.parent.currentText()
            c.relations = [(self.rels.item(r, 0).text(), self.rels.item(r, 1).text()) for r in range(self.rels.rowCount())]
            c.grouping = False
        else:
            c.parent = ""
            c.relations = []
        if not c.parent:
            c.grouping = self.grouping.isChecked() and self.gused.count() > 0
            c.groups = [self.gused.item(i).text() for i in range(self.gused.count())] if c.grouping else []
            g = self.gname.text().strip()
            c.grouping_name = g if c.grouping and NAME_RE.match(g) else ""
        aggs = []
        for r in range(self.aggs.rowCount()):
            name = (self.aggs.item(r, 0).text() if self.aggs.item(r, 0) else "").strip()
            if not name:
                continue
            aggs.append((name, self.aggs.cellWidget(r, 2).currentText(), self.aggs.cellWidget(r, 3).currentText(),
                         self.aggs.cellWidget(r, 1).currentIndex()))
        c.aggregates = aggs
        if self.rb_sql.isChecked() and self.params.rowCount() == 0:
            self._params_from_sql()
        c.params = self._read_params()
        d.retrieve(c)
        d.changed()
        d.fill_tree(c)


DRAG_ROLE = Qt.ItemDataRole.UserRole + 1
FIELDS_MIME = "application/x-lbstudio-defields"


class FieldTree(QTreeWidget):
    """the tree of the DataEnvironment: its fields and commands are dragged onto a DataReport (VB6)"""

    def __init__(self, d):
        super().__init__()
        self.d = d
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)

    def mimeData(self, items):
        from PyQt6.QtCore import QMimeData
        import json
        md = QMimeData()
        it = items[0] if items else None
        info = it.data(0, DRAG_ROLE) if it is not None else None
        if info:
            md.setData(FIELDS_MIME, json.dumps({"de": self.d.module.name, "member": info[0],
                                                "fields": info[1]}).encode())
            md.setText(", ".join(info[1]))
        return md

    def mimeTypes(self):
        return [FIELDS_MIME]


class DataEnvDesigner(_Base):
    def __init__(self, main, module):
        super().__init__(main, module)
        self.model = M.DataEnvModel(module.sf)
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        tb = QToolBar()
        tb.addAction("Add Connection", self.add_connection)
        tb.addAction("Add Command", lambda: self.add_command(False))
        tb.addAction("Add Child Command", lambda: self.add_command(True))
        tb.addAction("Delete", self.delete)
        tb.addSeparator()
        tb.addAction("Refresh", self.refresh)
        tb.addAction("View Code", lambda: self.main.open_code(self.module))
        v.addWidget(tb)
        sp = QSplitter()
        self.tree = FieldTree(self)
        self.tree.setHeaderHidden(True)
        self.tree.currentItemChanged.connect(self._chosen)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._menu)
        sp.addWidget(self.tree)
        self.pages = QStackedWidget()
        self.pages.addWidget(QLabel("  Select a connection or a command."))
        self.cpage = ConnectionPage(self)
        self.mpage = CommandPage(self)
        self.pages.addWidget(self.cpage)
        self.pages.addWidget(self.mpage)
        sp.addWidget(self.pages)
        sp.setSizes([260, 520])
        v.addWidget(sp, 1)
        v.addWidget(self.status)
        self.surface = self.tree
        self.fill_tree()

    # ------------------------------------------------------------------ the tree
    def fill_tree(self, select=None):
        self.model.name = self.module.name
        self.tree.blockSignals(True)
        self.tree.clear()
        root = QTreeWidgetItem([self.model.name])
        self.tree.addTopLevelItem(root)
        cf = QTreeWidgetItem(["Connections"])
        root.addChild(cf)
        sel = None
        for c in self.model.conns:
            it = QTreeWidgetItem([c.name])
            it.setData(0, Qt.ItemDataRole.UserRole, c)
            cf.addChild(it)
            if c is select:
                sel = it
        mf = QTreeWidgetItem(["Commands"])
        root.addChild(mf)

        def add(parent, c):
            nonlocal sel
            label = c.name + (f" grouped using {c.member}" if c.grouping else "")
            if c.parent:
                label += "  (" + ", ".join(f"{p} TO {ch}" for p, ch in c.relations) + ")"
            it = QTreeWidgetItem([label])
            it.setData(0, Qt.ItemDataRole.UserRole, c)
            parent.addChild(it)
            if c is select:
                sel = it
            aggs = [a[0] for a in c.aggregates if not (c.grouping and a[1].lower() == c.member.lower())]
            it.setData(0, DRAG_ROLE, (c.name, [f[0] for f in c.fields] + aggs))

            def leaf(parent_item, member, name):
                x = QTreeWidgetItem([name])
                x.setData(0, DRAG_ROLE, (member, [name]))
                parent_item.addChild(x)
            if c.grouping:
                g = QTreeWidgetItem([f"Summary fields in {c.member}"])
                g.setData(0, DRAG_ROLE, (c.member, list(c.groups)))
                it.addChild(g)
                for f in c.groups:
                    leaf(g, c.member, f)
                det = QTreeWidgetItem([f"Detail fields in {c.name}"])
                det.setData(0, DRAG_ROLE, (c.name, [f[0] for f in c.fields]))
                it.addChild(det)
                for f in c.fields:
                    leaf(det, c.name, f[0])
            else:
                for f in c.fields:
                    leaf(it, c.name, f[0])
            for a in c.aggregates:
                x = QTreeWidgetItem([f"{a[0]} = {M.AGG_FUNCS[a[3]] if 0 <= a[3] < len(M.AGG_FUNCS) else '?'}"
                                     f"({a[1]}.{a[2]})"])
                x.setData(0, DRAG_ROLE, (c.member if c.grouping and a[1].lower() == c.member.lower() else c.name, [a[0]]))
                it.addChild(x)
            for ch in self.model.children(c):
                add(it, ch)
        for c in self.model.top():
            add(mf, c)
        self.tree.expandToDepth(1)
        self.tree.blockSignals(False)
        if sel is not None:
            self.tree.setCurrentItem(sel)
        else:
            self._chosen(None, None)

    def selected(self):
        it = self.tree.currentItem()
        return it.data(0, Qt.ItemDataRole.UserRole) if it is not None else None

    def _chosen(self, cur, prev):
        o = cur.data(0, Qt.ItemDataRole.UserRole) if cur is not None else None
        self.say()
        if isinstance(o, M.DeConnection):
            self.cpage.show_conn(o)
            self.pages.setCurrentIndex(1)
        elif isinstance(o, M.DeCommand):
            self.mpage.show_cmd(o)
            self.pages.setCurrentIndex(2)
        else:
            self.pages.setCurrentIndex(0)

    def _menu(self, pos):
        menu = QMenu(self)
        menu.addAction("Add &Connection", self.add_connection)
        menu.addAction("Add C&ommand", lambda: self.add_command(False))
        if isinstance(self.selected(), M.DeCommand):
            menu.addAction("Add Child Co&mmand", lambda: self.add_command(True))
        if self.selected() is not None:
            menu.addAction("&Delete", self.delete)
        menu.addAction("&Refresh", self.refresh)
        menu.exec(self.tree.viewport().mapToGlobal(pos))

    # ------------------------------------------------------------------ actions
    def add_connection(self):
        c = M.DeConnection(self.model.unique("Connection", [x.name for x in self.model.conns]), "")
        self.model.conns.append(c)
        self.changed()
        self.fill_tree(c)

    def add_command(self, child):
        p = self.selected() if child else None
        if child and not isinstance(p, M.DeCommand):
            return
        c = M.DeCommand(self.model.unique("Command", [x.name for x in self.model.cmds]))
        c.conn = p.conn if p is not None else (self.model.conns[0].name if self.model.conns else "")
        if p is not None:
            c.parent = p.name
        self.model.cmds.append(c)
        self.changed()
        self.fill_tree(c)

    def delete(self):
        o = self.selected()
        if isinstance(o, M.DeConnection):
            if any(c.conn.lower() == o.name.lower() for c in self.model.cmds):
                return self.say(f"The commands of {o.name} use it")
            self.model.conns.remove(o)
        elif isinstance(o, M.DeCommand):
            gone = []
            stack = [o]
            while stack:
                x = stack.pop()
                gone.append(x)
                stack.extend(self.model.children(x))
            self.model.cmds = [c for c in self.model.cmds if c not in gone]
        else:
            return
        self.changed()
        self.fill_tree()

    def retrieve(self, c):
        """the fields of a command, as ADO gives them (the designer lists them, the reports bind them)"""
        conn = next((x for x in self.model.conns if x.name.lower() == c.conn.lower()), None)
        if conn is None or not conn.source or not c.text:
            return
        if c.ctype == 1 and is_action(c.text):  # it would be run to get them: an action query has none (VB6's
            c.fields = []                       # designer: NumFields 0, DEV-9/do1), and the database stays as it is
            self.say()
            return
        try:
            c.fields = dbinfo.ado_fields(design_cs(conn.source, self.project_dir()), c.ctype,
                                         without_params(c.text) if c.ctype == 1 else c.text)
            self.say()
        except dbinfo.AdoError as e:
            self.say(f"{c.name}: {e}")

    def refresh(self):
        dbinfo.forget()
        for c in self.model.cmds:
            self.retrieve(c)
        self.changed()
        self.fill_tree(self.selected())


# ====================================================================== DataReport
TWIPS = 15.0                                    # twips per pixel (96 dpi)
BAR = 18                                        # the bar of a section
TOOLS = [("Pointer", None), ("RptLabel", "Label"), ("RptTextBox", "Text"), ("RptImage", "Image"),
         ("RptLine", "Line"), ("RptShape", "Shape"), ("RptFunction", "Function")]
FUNCTION_TYPES = ["0 - rptFuncSum", "1 - rptFuncAve", "2 - rptFuncMin", "3 - rptFuncMax", "4 - rptFuncRCnt",
                  "5 - rptFuncVCnt", "6 - rptFuncSDEV", "7 - rptFuncSERR"]


def px(t):
    return int(round(float(t or 0) / TWIPS))


def tw(p):
    return int(round(p * TWIPS))


class ReportCanvas(QWidget):
    """the sections of the report one under the other, their controls on them"""

    def __init__(self, d):
        super().__init__()
        self.d = d
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.drag = None
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(d.context_menu)
        self.setAcceptDrops(True)

    # ------------------------------------------------------------------ fields dropped from a DataEnvironment
    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat(FIELDS_MIME):
            e.acceptProposedAction()

    def dragMoveEvent(self, e):
        s, _ = self.section_at(e.position().toPoint().y())
        if e.mimeData().hasFormat(FIELDS_MIME) and s is not None:
            e.acceptProposedAction()
        else:
            e.ignore()

    def dropEvent(self, e):
        import json
        pos = e.position().toPoint()
        s, body = self.section_at(pos.y())
        if s is None:
            return
        info = json.loads(bytes(e.mimeData().data(FIELDS_MIME)).decode())
        self.d.drop_fields(s, tw(pos.x() - body.left()), tw(pos.y() - body.top()), info["de"], info["member"],
                           info["fields"])
        e.acceptProposedAction()

    # ------------------------------------------------------------------ geometry
    def bands(self):
        """[(section, bar rect, body rect)]"""
        out, y = [], 0
        w = px(self.d.model.width)
        for s in self.d.model.sections:
            bar = QRect(0, y, w, BAR)
            body = QRect(0, y + BAR, w, px(s.height))
            out.append((s, bar, body))
            y += BAR + px(s.height)
        return out

    def sizeHint(self):
        b = self.bands()
        h = (b[-1][2].bottom() + 8) if b else 40
        from PyQt6.QtCore import QSize
        return QSize(px(self.d.model.width) + 8, h)

    def relayout(self):
        self.setFixedSize(self.sizeHint())
        self.update()

    def item_rect(self, body, it):
        x, y = body.left() + px(it.left), body.top() + px(it.top)
        w, h = max(px(it.width), 1), max(px(it.height), 1)
        return QRect(x, y, w, h)

    def hit(self, pos):
        """(kind, section, item): 'item', 'handle' (resize of an item), 'bar' (the bar under the section
        above: its height), 'body', 'width' (the right edge: ReportWidth)"""
        bands = self.bands()
        for i, (s, bar, body) in enumerate(bands):
            if body.contains(pos) or (abs(pos.y() - body.bottom()) <= 3 and i == len(bands) - 1 and 0 <= pos.x() <= body.right()):
                for it in reversed(s.items):
                    r = self.item_rect(body, it)
                    if it is self.d.sel and QRect(r.right() - 3, r.bottom() - 3, 8, 8).contains(pos):
                        return "handle", s, it
                    if r.adjusted(-2, -2, 2, 2).contains(pos):
                        return "item", s, it
                if i == len(bands) - 1 and abs(pos.y() - body.bottom()) <= 3:
                    return "bottom", s, None
                if abs(pos.x() - body.right()) <= 3:
                    return "width", s, None
                return "body", s, None
            if bar.contains(pos):
                if i > 0 and pos.y() - bar.top() <= 4:
                    return "bar", bands[i - 1][0], None
                return "section", s, None
        return None, None, None

    def section_at(self, y):
        for s, bar, body in self.bands():
            if body.top() <= y <= body.bottom() + 1:
                return s, body
        return None, None

    # ------------------------------------------------------------------ painting
    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#808080"))
        m = self.d.model
        step_x = max(1440 // max(m.gridx, 1), 1)
        step_y = max(1440 // max(m.gridy, 1), 1)
        f = None
        for s, bar, body in self.bands():
            p.setFont(self.font())
            p.fillRect(bar, QColor("#d4d0c8") if s is not self.d.sel else QColor("#0a246a"))
            p.setPen(QColor("#000000") if s is not self.d.sel else QColor("#ffffff"))
            p.drawText(bar.adjusted(4, 0, 0, 0), Qt.AlignmentFlag.AlignVCenter,
                       f"{M.SEC_NAMES.get(s.code, 'Section')} ({s.name})")
            p.fillRect(body, QColor("#ffffff"))
            p.setPen(QColor("#c0c0c0"))
            for gy in range(0, int(s.height), step_y):
                for gx in range(0, int(m.width), step_x):
                    p.drawPoint(body.left() + px(gx), body.top() + px(gy))
            p.setClipRect(body)
            for it in s.items:
                self.paint_item(p, body, it, f)
            p.setClipping(False)
        if self.drag and self.drag[0] == "new":
            p.setPen(QPen(QColor("#000000"), 1, Qt.PenStyle.DashLine))
            p.drawRect(QRect(self.drag[1], self.drag[2]).normalized())

    def _image(self, it):
        """the picture of an RptImage, made once for its data"""
        from PyQt6.QtGui import QImage
        key = id(it.pic) if it.pic else None
        if key is None:
            return None
        cache = self.__dict__.setdefault("_imgs", {})
        if key not in cache:
            img = QImage()
            img.loadFromData(it.pic)
            cache.clear() if len(cache) > 32 else None
            cache[key] = img
        return cache[key]

    def paint_item(self, p, body, it, base_font):
        r = self.item_rect(body, it)
        k = it.kind
        font = QFont(it.font.get("name", "Arial"))
        font.setPointSizeF(float(it.font.get("size", 8.25) or 8.25))
        font.setBold(int(it.font.get("weight", 400) or 400) >= 600)
        font.setItalic(bool(it.font.get("italic")))
        font.setUnderline(bool(it.font.get("underline")))
        p.setFont(font)
        al = {0: Qt.AlignmentFlag.AlignLeft, 1: Qt.AlignmentFlag.AlignRight, 2: Qt.AlignmentFlag.AlignHCenter}.get(
            int(it.props.get("Alignment", 0) or 0), Qt.AlignmentFlag.AlignLeft)
        if k == "RptLine":
            p.setPen(QPen(QColor("#000000"), 1))
            p.drawLine(r.left(), r.top(), r.left() + px(it.width), r.top() + px(it.height))
        elif k == "RptShape":
            p.setPen(QPen(QColor("#000000"), 1))
            if int(it.props.get("Shape", 0) or 0) in (2, 3):
                p.drawEllipse(r)
            else:
                p.drawRect(r)
        elif k == "RptImage":
            img = self._image(it)
            if img is not None and not img.isNull():
                if int(it.props.get("SizeMode", 0) or 0) == 1:
                    p.drawImage(r, img)
                else:
                    sz = img.size()
                    if int(it.props.get("SizeMode", 0) or 0) == 2:
                        sz = sz.scaled(r.size(), Qt.AspectRatioMode.KeepAspectRatio)
                    p.save()
                    p.setClipRect(r)
                    p.drawImage(QRect(r.topLeft(), sz), img)
                    p.restore()
            else:
                p.setPen(QPen(QColor("#000000"), 1))
                p.drawRect(r)
                p.drawText(r, Qt.AlignmentFlag.AlignCenter, "Image")
        else:
            if k == "RptLabel":
                text = str(it.props.get("Caption", it.name))
            elif k == "RptTextBox":
                text = str(it.props.get("DataField", "") or "Unbound")
            else:
                ft = int(it.props.get("FunctionType", 0) or 0)
                fn = ["Sum", "Ave", "Min", "Max", "RCnt", "VCnt", "SDEV", "SERR"][ft if 0 <= ft < 8 else 0]
                text = f"={fn}({it.props.get('DataField', '')})"
            if k != "RptLabel" or int(it.props.get("BorderStyle", 0) or 0):
                p.setPen(QPen(QColor("#808080"), 1, Qt.PenStyle.DotLine))
                p.drawRect(r)
            p.setPen(QColor("#000000"))
            p.drawText(r.adjusted(1, 0, -1, 0), al | Qt.AlignmentFlag.AlignTop, text)
        if it is self.d.sel:
            p.setPen(QPen(QColor("#0000ff"), 1))
            p.drawRect(r.adjusted(-1, -1, 0, 0))
            p.fillRect(QRect(r.right() - 2, r.bottom() - 2, 5, 5), QColor("#0000ff"))

    # ------------------------------------------------------------------ the mouse
    def snap(self, t, step):
        return int(round(t / step) * step) if self.d.main is None or getattr(self.d.main, "a_grid", None) is None \
            or self.d.main.a_grid.isChecked() else int(t)

    def mousePressEvent(self, e):
        self.setFocus()
        pos = e.position().toPoint()
        if e.button() != Qt.MouseButton.LeftButton:
            kind, s, it = self.hit(pos)
            self.d.select(it or s or None)
            return
        tool = self.d.tool()
        kind, s, it = self.hit(pos)
        if tool and kind in ("body", "item", "handle"):
            self.drag = ("new", pos, pos, s)
            return
        if kind in ("item", "handle"):
            self.d.select(it)
            self.drag = (kind, pos, (it.left, it.top, it.width, it.height), s, it)
        elif kind in ("bar", "bottom"):
            self.drag = ("height", pos, s.height, s)
        elif kind == "width":
            self.drag = ("width", pos, self.d.model.width)
        elif kind in ("section", "body"):
            self.d.select(s)
        else:
            self.d.select(None)

    def mouseMoveEvent(self, e):
        pos = e.position().toPoint()
        m = self.d.model
        sx, sy = max(1440 // max(m.gridx, 1), 1), max(1440 // max(m.gridy, 1), 1)
        if not self.drag:
            kind = self.hit(pos)[0]
            self.setCursor({"bar": Qt.CursorShape.SizeVerCursor, "bottom": Qt.CursorShape.SizeVerCursor,
                            "width": Qt.CursorShape.SizeHorCursor, "handle": Qt.CursorShape.SizeFDiagCursor}.get(
                kind, Qt.CursorShape.CrossCursor if self.d.tool() else Qt.CursorShape.ArrowCursor))
            return
        d = self.drag
        dx, dy = tw(pos.x() - d[1].x()), tw(pos.y() - d[1].y())
        if d[0] == "new":
            self.drag = (d[0], d[1], pos, d[3])
        elif d[0] == "item":
            it = d[4]
            it.left = max(0, self.snap(d[2][0] + dx, sx))
            it.top = max(0, self.snap(d[2][1] + dy, sy))
        elif d[0] == "handle":
            it = d[4]
            if it.kind == "RptLine":
                it.width = self.snap(d[2][2] + dx, sx)
                it.height = self.snap(d[2][3] + dy, sy)
            else:
                it.width = max(sx, self.snap(d[2][2] + dx, sx))
                it.height = max(sy, self.snap(d[2][3] + dy, sy))
        elif d[0] == "height":
            d[3].height = max(0, self.snap(d[2] + dy, sy))
            self.relayout()
        elif d[0] == "width":
            m.width = max(1440, self.snap(d[2] + dx, sx))
            self.relayout()
        self.update()

    def mouseReleaseEvent(self, e):
        d, self.drag = self.drag, None
        if not d:
            return
        if d[0] == "new":
            self.d.create(d[3], QRect(d[1], d[2]).normalized())
        elif d[0] == "item":
            # into another section when it is dropped on one
            it, s0 = d[4], d[3]
            body0 = next(b for s, _, b in self.bands() if s is s0)
            y = body0.top() + px(it.top)
            s1, body1 = self.section_at(y)
            if s1 is not None and s1 is not s0:
                if it.kind == "RptFunction" and s1.code not in (5, 8):
                    self.d.say("A Function control can be in a group footer or the report footer only")
                    it.left, it.top = d[2][0], d[2][1]
                else:
                    s0.items.remove(it)
                    s1.items.append(it)
                    it.top = max(0, tw(y - body1.top()))
            if (it.left, it.top, it.width, it.height) != d[2] or s1 is not s0:
                self.d.changed()
        elif d[0] in ("handle",):
            if (d[4].left, d[4].top, d[4].width, d[4].height) != d[2]:
                self.d.changed()
        elif d[0] == "height":
            if d[3].height != d[2]:
                self.d.changed()
        elif d[0] == "width":
            if self.d.model.width != d[2]:
                self.d.changed()
        self.update()
        self.d.show_props()

    def mouseDoubleClickEvent(self, e):
        self.d.main.open_code(self.d.module)

    def keyPressEvent(self, e):
        it = self.d.sel
        if e.key() == Qt.Key.Key_Delete:
            self.d.delete()
            return
        if isinstance(it, M.RptItem) and e.key() in (Qt.Key.Key_Left, Qt.Key.Key_Right, Qt.Key.Key_Up, Qt.Key.Key_Down):
            step = 15 if e.modifiers() & Qt.KeyboardModifier.ControlModifier else 1440 // max(self.d.model.gridx, 1)
            dx = {Qt.Key.Key_Left: -step, Qt.Key.Key_Right: step}.get(e.key(), 0)
            dy = {Qt.Key.Key_Up: -step, Qt.Key.Key_Down: step}.get(e.key(), 0)
            if e.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                it.width = max(15, it.width + dx)
                it.height = max(0 if it.kind == "RptLine" else 15, it.height + dy)
            else:
                it.left, it.top = max(0, it.left + dx), max(0, it.top + dy)
            self.d.changed()
            self.update()
            self.d.show_props()
            return
        super().keyPressEvent(e)


class ReportDesigner(_Base):
    def __init__(self, main, module):
        super().__init__(main, module)
        if module.dsx_data is not None:         # changed, not saved yet
            d = module.dsx_data
            dsx = d[9:9 + d[8]].decode("cp1252", "replace") if len(d) > 9 else ""
        else:
            dsx = M.read_dsx(module.path) if module.path else ""
        self.model = M.ReportModel(module.sf, dsx, module.path, module.dsx_data)
        self.sel = None
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        tb = QToolBar()
        self.tools = QButtonGroup(self)
        self.tools.setExclusive(True)
        for i, (kind, label) in enumerate(TOOLS):
            b = QToolButton()
            b.setText(label or "Pointer")
            b.setCheckable(True)
            b.setToolTip(kind)
            b.setChecked(i == 0)
            self.tools.addButton(b, i)
            tb.addWidget(b)
        tb.addSeparator()
        tb.addAction("Retrieve Structure", self.retrieve_structure)
        tb.addAction("View Code", lambda: self.main.open_code(self.module))
        v.addWidget(tb)
        sp = QSplitter()
        self.canvas = ReportCanvas(self)
        sa = QScrollArea()
        sa.setWidget(self.canvas)
        sa.setStyleSheet("QScrollArea { background: #808080; }")
        sp.addWidget(sa)
        pw = QWidget()
        pv = QVBoxLayout(pw)
        pv.setContentsMargins(4, 4, 4, 4)
        self.obj_box = QComboBox()
        self.obj_box.activated.connect(self._obj_chosen)
        pv.addWidget(self.obj_box)
        self.form_host = QWidget()
        self.form = QFormLayout(self.form_host)
        pv.addWidget(self.form_host)
        pv.addStretch(1)
        sp.addWidget(pw)
        sp.setSizes([620, 260])
        v.addWidget(sp, 1)
        v.addWidget(self.status)
        self.surface = self.canvas
        self.canvas.relayout()
        self.show_props()

    def tool(self):
        i = self.tools.checkedId()
        return TOOLS[i][0] if i > 0 else None

    # ------------------------------------------------------------------ selection and properties
    def select(self, o):
        self.sel = o
        self.canvas.update()
        self.show_props()

    def _objects(self):
        out = [(f"{self.model.name} DataReport", None)]
        for s in self.model.sections:
            out.append((f"{s.name} Section", s))
            for it in s.items:
                out.append((f"{it.name} {it.kind}", it))
        return out

    def _obj_chosen(self, i):
        objs = self._objects()
        if 0 <= i < len(objs):
            self.select(objs[i][1])

    def _clear_form(self):
        while self.form.rowCount():
            self.form.removeRow(0)

    def _row(self, label, kind, value, setter, items=()):
        if kind == "str":
            w = QLineEdit(str(value if value is not None else ""))
            w.editingFinished.connect(lambda w=w: setter(w.text()))
        elif kind == "int":
            w = QSpinBox()
            w.setRange(-1000000, 10000000)
            w.setValue(int(value or 0))
            w.editingFinished.connect(lambda w=w: setter(w.value()))
        elif kind == "bool":
            w = QCheckBox()
            w.setChecked(value not in (0, False, None, "0"))
            w.toggled.connect(lambda on: setter(-1 if on else 0))
        elif kind == "enum":
            w = QComboBox()
            w.addItems(items)
            w.setCurrentIndex(int(value or 0) if 0 <= int(value or 0) < len(items) else 0)
            w.activated.connect(lambda i: setter(i))
        elif kind == "picture":                 # (None) / (Bitmap): Load... or clear
            w = QPushButton("(Bitmap)" if value else "(None)")
            m = QMenu(w)
            m.addAction("Load Picture...", lambda: self._load_picture(setter))
            m.addAction("(None)", lambda: setter(None))
            w.setMenu(m)
        else:                                   # combo: a list to choose from or a text to type
            w = QComboBox()
            w.setEditable(True)
            w.addItems(list(items))
            w.setEditText(str(value or ""))
            w.activated.connect(lambda i, w=w: setter(w.currentText()))
            w.lineEdit().editingFinished.connect(lambda w=w: setter(w.currentText()))
        self.form.addRow(label, w)
        return w

    def show_props(self):
        self.obj_box.blockSignals(True)
        self.obj_box.clear()
        objs = self._objects()
        for text, o in objs:
            self.obj_box.addItem(text)
        self.obj_box.setCurrentIndex(next((i for i, (_, o) in enumerate(objs) if o is self.sel), 0))
        self.obj_box.blockSignals(False)
        self._clear_form()
        o, m = self.sel, self.model
        if o is None:
            self._row("Caption", "str", m.caption, lambda v: self._set(m, "caption", v))
            self._row("DataSource", "combo", m.datasource, self._set_source, ["(None)"] + self.data_environments())
            self._row("DataMember", "combo", m.datamember, lambda v: self._set(m, "datamember", v), self.members())
            self._row("Title", "str", m.title, lambda v: self._set(m, "title", v))
            self._row("ReportWidth", "int", m.width, lambda v: self._set(m, "width", max(1440, v)))
            for i, n in enumerate(("LeftMargin", "RightMargin", "TopMargin", "BottomMargin")):
                self._row(n, "int", m.margins[i], lambda v, i=i: self._margin(i, v))
            self._row("GridX", "int", m.gridx, lambda v: self._set(m, "gridx", max(1, v)))
            self._row("GridY", "int", m.gridy, lambda v: self._set(m, "gridy", max(1, v)))
            self._row("Font", "str", f"{m.font.get('name')}, {m.font.get('size')}", lambda v: self._font(m.font, v))
        elif isinstance(o, M.RptSection):
            self._row("(Name)", "str", o.name, lambda v: self._rename(o, v))
            self._row("Height", "int", o.height, lambda v: self._set(o, "height", max(0, v)))
            self._row("KeepTogether", "bool", o.props.get("KeepTogether", 0), lambda v: self._prop(o, "KeepTogether", v))
            self._row("ForcePageBreak", "enum", o.props.get("ForcePageBreak", 0),
                      lambda v: self._prop(o, "ForcePageBreak", v),
                      ["0 - rptPageBreakNone", "1 - rptPageBreakBefore", "2 - rptPageBreakAfter", "3 - rptPageBreakBeforeAndAfter"])
            self._row("Visible", "bool", o.props.get("Visible", -1), lambda v: self._prop(o, "Visible", v, -1))
        else:
            self._row("(Name)", "str", o.name, lambda v: self._rename(o, v))
            for n in ("Left", "Top", "Width", "Height"):
                self._row(n, "int", getattr(o, n.lower()), lambda v, n=n: self._set(o, n.lower(), v))
            if o.kind == "RptLabel":
                self._row("Caption", "str", o.props.get("Caption", ""), lambda v: self._prop(o, "Caption", v, None))
            if o.kind in ("RptTextBox", "RptFunction"):
                self._row("DataMember", "combo", o.props.get("DataMember", ""), lambda v: self._prop(o, "DataMember", v, None),
                          self.members(all_levels=True))
                self._row("DataField", "combo", o.props.get("DataField", ""), lambda v: self._prop(o, "DataField", v, None),
                          self.fields_of(o.props.get("DataMember", "")))
                self._row("DataFormat", "str", (o.fmt or {}).get("format", ""), lambda v: self._format(o, v))
            if o.kind == "RptFunction":
                self._row("FunctionType", "enum", o.props.get("FunctionType", 0), lambda v: self._prop(o, "FunctionType", v),
                          FUNCTION_TYPES)
            if o.kind in ("RptLabel", "RptTextBox", "RptFunction"):
                self._row("Alignment", "enum", o.props.get("Alignment", 0), lambda v: self._prop(o, "Alignment", v),
                          ["0 - rptJustifyLeft", "1 - rptJustifyRight", "2 - rptJustifyCenter"])
                self._row("CanGrow", "bool", o.props.get("CanGrow", 0), lambda v: self._prop(o, "CanGrow", v))
                self._row("Font", "str", f"{o.font.get('name')}, {o.font.get('size')}" +
                          (", Bold" if int(o.font.get("weight", 400) or 400) >= 600 else ""), lambda v: self._font(o.font, v))
            if o.kind in ("RptShape",):
                self._row("Shape", "enum", o.props.get("Shape", 0), lambda v: self._prop(o, "Shape", v),
                          ["0 - rptShpRectangle", "1 - rptShpSquare", "2 - rptShpOval", "3 - rptShpCircle",
                           "4 - rptShpRoundedRectangle", "5 - rptShpRoundedSquare"])
            if o.kind == "RptImage":
                self._row("Picture", "picture", o.pic, lambda v: self._picture(o, v))
                self._row("SizeMode", "enum", o.props.get("SizeMode", 0), lambda v: self._prop(o, "SizeMode", v),
                          ["0 - rptSizeClip", "1 - rptSizeStretch", "2 - rptSizeZoom"])
                self._row("PictureAlignment", "enum", o.props.get("PictureAlignment", 8),
                          lambda v: self._prop(o, "PictureAlignment", v, 8),
                          ["0 - rptPATopLeft", "1 - rptPATop", "2 - rptPATopRight", "3 - rptPARight",
                           "4 - rptPABottomRight", "5 - rptPABottom", "6 - rptPABottomLeft", "7 - rptPALeft",
                           "8 - rptPACenter"])
            if o.kind in ("RptLabel", "RptTextBox", "RptFunction", "RptShape", "RptImage"):
                # a shape has a border unless told (VB6 leaves BorderStyle = 1 out: oracle d_de2)
                bdef = 1 if o.kind == "RptShape" else 0
                self._row("BorderStyle", "enum", o.props.get("BorderStyle", bdef),
                          lambda v: self._prop(o, "BorderStyle", v, bdef),
                          ["0 - rptBSTransparent", "1 - rptBSSolid", "2 - rptBSDashes", "3 - rptBSDots",
                           "4 - rptBSDashDot", "5 - rptBSDashDotDot"])

    def _load_picture(self, setter):
        path, _ = QFileDialog.getOpenFileName(self, "Load Picture", self.project_dir(),
                                              "Pictures (*.bmp *.dib *.jpg *.jpeg *.gif *.ico *.wmf *.emf);;All Files (*)")
        if path:
            try:
                with open(path, "rb") as f:
                    setter(f.read())
            except OSError as e:
                self.say(str(e))

    def _picture(self, o, data):
        if o.pic == data:
            return
        o.pic = data
        self.changed()
        self.canvas.update()
        self.show_props()

    def _set(self, o, attr, v):
        if getattr(o, attr) == v:
            return
        setattr(o, attr, v)
        self.changed()
        self.canvas.relayout()

    def _margin(self, i, v):
        if self.model.margins[i] != v:
            self.model.margins[i] = v
            self.changed()

    def _prop(self, o, name, v, default=0):
        """a property of a control or section: its default value is not written (as VB6)"""
        old = o.props.get(name, default)
        if v == old:
            return
        if v == default and name not in ("Caption",):
            o.props.pop(name, None)
        else:
            o.props[name] = v
        self.changed()
        self.canvas.update()
        if name == "DataMember":
            self.show_props()

    def _format(self, o, v):
        f = o.fmt or {"type": 0, "format": ""}
        if f.get("format", "") == v:
            return
        f["format"] = v
        f["type"] = 1 if v else 0
        o.fmt = f
        self.changed()

    def _font(self, font, text):
        """"Arial, 10, Bold": the name, the size, Bold / Italic / Underline"""
        parts = [x.strip() for x in text.split(",") if x.strip()]
        if not parts:
            return
        new = dict(font)
        new["name"] = parts[0]
        for x in parts[1:]:
            try:
                new["size"] = float(x)
                continue
            except ValueError:
                pass
            k = x.lower()
            if k == "bold":
                new["weight"] = 700
            elif k in ("italic", "underline", "strikethrough"):
                new[k] = True
        if not any(x.lower() == "bold" for x in parts[1:]):
            new["weight"] = 400
        for k in ("italic", "underline", "strikethrough"):
            if not any(x.lower() == k for x in parts[1:]):
                new[k] = False
        if new != font:
            font.clear()
            font.update(new)
            self.changed()
            self.canvas.update()

    def _rename(self, o, v):
        v = v.strip()
        if v == o.name:
            return
        if not NAME_RE.match(v):
            return self.say("Invalid property value")
        if v.lower() in {n.lower() for n in self.model.all_names()} or v.lower() == self.model.name.lower():
            return self.say(f"Name '{v}' is already in use")
        o.name = v
        self.say()
        self.changed()
        self.canvas.update()
        self.show_props()

    def _set_source(self, v):
        v = "" if v == "(None)" else v.strip()
        if v == self.model.datasource:
            return
        self.model.datasource = v
        self.changed()
        self.show_props()

    # ------------------------------------------------------------------ the DataEnvironment
    def data_environments(self):
        return [m.name for m in self.module.project.modules if m.designer_kind == "DataEnvironment"]

    def de_model(self):
        name = self.model.datasource
        m = next((x for x in self.module.project.modules if x.designer_kind == "DataEnvironment"
                  and x.name.lower() == name.lower()), None) if name else None
        if m is None:
            return None
        sw = getattr(self.main, "design_windows", {}).get(m)
        if sw is not None and isinstance(sw.widget(), DataEnvDesigner):
            return sw.widget().model
        return M.DataEnvModel(m.sf)

    def chain(self):
        """the levels of the report's DataMember: [(member name, command)] from the top to the detail"""
        de = self.de_model()
        if de is None or not self.model.datamember:
            return []
        top = next((c for c in de.top() if c.member.lower() == self.model.datamember.lower()), None)
        if top is None:
            return []
        out = [(top.member, top)] if top.grouping else []
        c = top
        while c is not None:
            out.append((c.name, c))
            ch = de.children(c)
            c = ch[0] if ch else None
        return out

    def members(self, all_levels=False):
        """the DataMember list: of the report, the top members of the DataEnvironment; of a control,
        the levels of the report's DataMember and every command below it"""
        de = self.de_model()
        if de is None:
            return []
        chain = self.chain()
        if not all_levels or not chain:
            return de.members()
        out = [name for name, _ in chain]
        stack = list(de.children(chain[0][1]))
        while stack:
            x = stack.pop(0)
            if x.name not in out:
                out.append(x.name)
            stack.extend(de.children(x))
        return out

    def fields_of(self, member):
        """the DataField list of a control: the fields of its command and the aggregates the command
        has (the grouping: its grouping fields and the aggregates on it)"""
        de = self.de_model()
        if de is None or not member:
            return []
        for c in de.cmds:
            if c.grouping and c.member.lower() == member.lower():
                return c.groups + [a[0] for a in c.aggregates if a[1].lower() == member.lower()]
            if c.name.lower() == member.lower():
                return [f[0] for f in c.fields] + [a[0] for a in c.aggregates if a[1].lower() != c.member.lower() or not c.grouping]
        return []

    def retrieve_structure(self):
        """the group sections of the report for the hierarchy of its DataMember (VB6's Retrieve
        Structure): a header and a footer for every level above the detail"""
        chain = self.chain()
        if not chain:
            return self.say("Set the DataSource and the DataMember of the report first")
        if any(s.items for s in self.model.sections if s.code in (3, 4, 5)):
            r = QMessageBox.question(self, "Data Report", "Replace existing report layout?\n"
                                     "The controls of the group and detail sections will be lost.")
            if r != QMessageBox.StandardButton.Yes:
                return
        old = {s.name.lower(): s for s in self.model.sections}
        keep = lambda code: [s for s in self.model.sections if s.code == code]  # noqa: E731
        secs = keep(1) + keep(2)
        for name, _ in chain[:-1]:
            secs.append(M.RptSection(3, f"{name}_Header", 360))
        secs.append(M.RptSection(4, f"{chain[-1][0]}_Detail", 360))
        for name, _ in reversed(chain[:-1]):
            secs.append(M.RptSection(5, f"{name}_Footer", 360))
        secs += keep(7) + keep(8)
        for i, s in enumerate(secs):
            o = old.get(s.name.lower())
            if o is not None and s.code in (3, 4, 5):
                secs[i] = o
        self.model.sections = secs
        self.sel = None
        self.changed()
        self.canvas.relayout()
        self.show_props()

    # ------------------------------------------------------------------ editing
    def create(self, s, rect):
        kind = self.tool()
        if s is None or kind is None:
            return
        if kind == "RptFunction" and s.code not in (5, 8):
            self.say("A Function control can be in a group footer or the report footer only")
            self.tools.button(0).setChecked(True)
            return
        body = next(b for x, _, b in self.canvas.bands() if x is s)
        base = dict(TOOLS)[kind]
        it = M.RptItem(kind, self.model.unique(base))
        it.font = dict(self.model.font)
        it.left = max(0, tw(rect.left() - body.left()))
        it.top = max(0, tw(rect.top() - body.top()))
        w, h = tw(rect.width()), tw(rect.height())
        if w < 60 and h < 60:                   # a click: VB6's default size
            w, h = (1440, 0) if kind == "RptLine" else (1440, 240 if kind in ("RptLabel", "RptTextBox", "RptFunction") else 720)
        if kind == "RptLine" and h < 60:
            h = 0
        it.width, it.height = w, h
        if kind == "RptLabel":
            it.props["Caption"] = it.name
        if kind in ("RptTextBox", "RptFunction"):
            it.props["DataMember"] = ""
        s.items.append(it)
        self.tools.button(0).setChecked(True)
        self.say()
        self.changed()
        self.select(it)

    def drop_fields(self, s, x, y, de, member, fields):
        """fields of a DataEnvironment dropped on a section: for each a label with its name and a text
        box bound to it, one under the other (VB6's drag and drop field mapping); the report takes the
        DataEnvironment as its DataSource when it has none"""
        if not self.model.datasource:
            self._set_source(de)
        elif self.model.datasource.lower() != de.lower():
            return self.say(f"The DataSource of the report is {self.model.datasource}")
        if not self.model.datamember and member in self.members():
            self.model.datamember = member
        gx = 1440 // max(self.model.gridx, 1)
        gy = 1440 // max(self.model.gridy, 1)
        x, y = max(0, round(x / gx) * gx), max(0, round(y / gy) * gy)
        last = None
        for i, f in enumerate(fields):
            top = y + i * 288
            lab = M.RptItem("RptLabel", self.model.unique("Label"))
            lab.font = dict(self.model.font)
            lab.left, lab.top, lab.width, lab.height = x, top, 1440, 240
            lab.props["Caption"] = f + ":"
            s.items.append(lab)
            txt = M.RptItem("RptTextBox", self.model.unique("txt" + re.sub(r"\W", "", f)[:30] + "_"))
            txt.name = txt.name[:-2] if txt.name.endswith("_1") and txt.name[:-2].lower() not in \
                {n.lower() for n in self.model.all_names()} else txt.name
            txt.font = dict(self.model.font)
            txt.left, txt.top, txt.width, txt.height = x + 1440 + 144, top, 1440, 240
            txt.props["DataField"] = f
            txt.props["DataMember"] = member
            s.items.append(txt)
            last = txt
        need = (y + len(fields) * 288) if fields else 0
        if need > s.height:
            s.height = need
        self.changed()
        self.canvas.relayout()
        if last is not None:
            self.select(last)

    def delete(self):
        o = self.sel
        if isinstance(o, M.RptItem):
            for s in self.model.sections:
                if o in s.items:
                    s.items.remove(o)
            self.sel = None
            self.changed()
            self.canvas.update()
            self.show_props()

    def _toggle_pair(self, codes, names):
        have = [s for s in self.model.sections if s.code in codes]
        if have:
            if any(s.items for s in have):
                r = QMessageBox.question(self, "Data Report", "Deleting the sections deletes their controls. Continue?")
                if r != QMessageBox.StandardButton.Yes:
                    return
            self.model.sections = [s for s in self.model.sections if s.code not in codes]
        else:
            head = M.RptSection(codes[0], self.model.unique(names[0]) if names[0].endswith("Section") else names[0], 360)
            foot = M.RptSection(codes[1], names[1], 360)
            secs = self.model.sections
            # the report header first, the page header after it; the footers the other way round
            i = 0 if codes[0] == 1 else len([s for s in secs if s.code == 1])
            secs.insert(i, head)
            j = len(secs) if codes[1] == 8 else len([s for s in secs if s.code != 8])
            secs.insert(j, foot)
        self.sel = None
        self.changed()
        self.canvas.relayout()
        self.show_props()

    def context_menu(self, pos):
        menu = QMenu(self)
        has = {s.code for s in self.model.sections}
        a = menu.addAction("Report &Header/Footer", lambda: self._toggle_pair((1, 8), ("ReportHeader", "ReportFooter")))
        a.setCheckable(True)
        a.setChecked(1 in has)
        a = menu.addAction("&Page Header/Footer", lambda: self._toggle_pair((2, 7), ("PageHeader", "PageFooter")))
        a.setCheckable(True)
        a.setChecked(2 in has)
        menu.addSeparator()
        menu.addAction("&Retrieve Structure", self.retrieve_structure)
        if isinstance(self.sel, M.RptItem):
            menu.addAction("&Delete", self.delete)
        menu.addSeparator()
        menu.addAction("View &Code", lambda: self.main.open_code(self.module))
        menu.exec(self.canvas.mapToGlobal(pos))

    def changed(self):
        if self.model.sf.root is None or not any(getattr(x, "name", "").lower() == "bindings" for x in self.model.sf.root.items):
            self.model.dsx_file = os.path.splitext(os.path.basename(self.module.rel.replace("\\", "/")))[0] + ".dsx"
        super().changed()
        self.module.dsx_data = self.model.dsx_bytes() if self.model.needs_dsx() else None


def designer_for(main, module):
    k = module.designer_kind
    if k == "DataEnvironment":
        return DataEnvDesigner(main, module)
    if k == "DataReport":
        return ReportDesigner(main, module)
    return None
