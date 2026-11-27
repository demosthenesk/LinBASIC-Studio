"""The Database Window of the Visual Data Manager (VB6 DATABASE.FRM): the Properties of the database, its tables
(Fields, Indexes, Properties - a property: Name=Value) and its queries; the menu of the right button; Open (a table, a
query) as VisData opens them (VISDATA.BAS OpenTable / OpenQuery)."""
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QIcon, QPainter, QPen, QPixmap
from PyQt6.QtWidgets import QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget, QAbstractItemView

from ..rtobj import RtError, EMPTY
from . import core as C

PROPERTIES, PROPERTY = "Properties", "Property"
TABLE, QUERY, FIELDS, FIELD, INDEXES, INDEX, ATTACHED = "Table", "Query", "Fields", "Field", "Indexes", "Index", \
    "Attached"
ROLE = Qt.ItemDataRole.UserRole


def _icon(kind):
    pm = QPixmap(16, 16)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setPen(QPen(QColor("#000000")))
    if kind in (TABLE, ATTACHED):
        p.setBrush(QColor("#ffffff"))
        p.drawRect(1, 2, 13, 11)
        p.fillRect(2, 3, 12, 3, QColor("#000080"))
        for y in (8, 11):
            p.drawLine(2, y, 13, y)
        p.drawLine(6, 6, 6, 12)
        if kind == ATTACHED:
            p.setPen(QPen(QColor("#008000"), 2))
            p.drawLine(9, 9, 14, 14)
    elif kind == QUERY:
        p.setBrush(QColor("#ffffff"))
        p.drawRect(1, 4, 9, 9)
        p.drawRect(5, 1, 9, 9)
    elif kind in (FIELD, FIELDS):
        p.setBrush(QColor("#c0c0c0"))
        p.drawRect(2, 5, 11, 5)
    elif kind in (INDEX, INDEXES):
        p.setBrush(QColor("#ffff80"))
        p.drawRect(2, 2, 9, 12)
        p.drawLine(4, 6, 9, 6)
        p.drawLine(4, 9, 9, 9)
    else:                                       # a property
        p.setBrush(QColor("#e0e0e0"))
        p.drawRect(2, 2, 11, 11)
        p.drawLine(4, 6, 11, 6)
        p.drawLine(4, 9, 11, 9)
    p.end()
    return QIcon(pm)


def prop_value(p):
    """GetPropertyValue: the value of a property (its error: the text of it)"""
    try:
        v = p.get("Value")
    except RtError as e:
        return f"<{e.description}>" if e.description else ""
    if v is EMPTY:
        return ""
    if isinstance(v, (bytes, bytearray)):
        return "(Binary)"
    return C.text_of(v)


class DBTree(QTreeWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win
        self.setHeaderHidden(True)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

    def mouseReleaseEvent(self, e):
        super().mouseReleaseEvent(e)
        if e.button() == Qt.MouseButton.RightButton:
            self.win.popup_at(self.itemAt(e.position().toPoint()), e.globalPosition().toPoint())


class DatabaseWindow(QWidget):
    def __init__(self, vd):
        super().__init__()
        self.vd = vd
        self.setWindowTitle("Database Window")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        self.tree = DBTree(self)
        self.tree.itemExpanded.connect(self.node_click)
        self.tree.itemClicked.connect(lambda it, _c: self.node_click(it))
        self.tree.itemDoubleClicked.connect(self.double_click)
        self.tree.itemChanged.connect(self.label_edited)
        lay.addWidget(self.tree)
        self.node2 = None
        self._editing = None
        self.load()

    @property
    def db(self):
        return self.vd.s.db

    # ---------------------------------------------------------------- the tree
    def node(self, parent, text, tag, icon, key=None):
        it = QTreeWidgetItem([text])
        it.setIcon(0, _icon(icon))
        it.setData(0, ROLE, (tag, key))
        if parent is None:
            self.tree.addTopLevelItem(it)
        else:
            parent.addChild(it)
        if tag in (PROPERTIES, FIELDS, INDEXES, FIELD, INDEX, TABLE, QUERY):
            it.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.ShowIndicator)
        return it

    def load(self):
        """LoadDatabase"""
        self.tree.blockSignals(True)
        self.tree.clear()
        self.tree.blockSignals(False)
        if self.db is None:
            return
        self.node(None, PROPERTIES, PROPERTIES, PROPERTY)
        found = False
        for td in self.vd.s.tabledefs(self.vd.prefs.get("allow_sys")):
            name = td.get("Name")
            found = True
            att = C.attached(td)
            t = self.node(None, name, TABLE, ATTACHED if att else TABLE)
            self.node(t, FIELDS, FIELDS, FIELD)
            self.node(t, INDEXES, INDEXES, INDEX)
            self.node(t, PROPERTIES, PROPERTIES, PROPERTY)
            if att:
                cn = td.get("Connect")
                kind = cn.split(";")[0] if ";" in cn else ""
                self.node(t, (kind or C.MSACCESS) + " Table", "AttachType", ATTACHED)
        for qd in self.db.each("QueryDefs"):
            q = self.node(None, qd.get("Name"), QUERY, QUERY)
            self.node(q, PROPERTIES, PROPERTIES, PROPERTY)
        self.vd.a_query.setEnabled(found)
        self.vd.pu_newquery.setVisible(found)

    def tag(self, it):
        d = it.data(0, ROLE) if it is not None else None
        return d[0] if d else None

    def node_click(self, it):
        """tvDatabase_NodeClick: the children of a node read when it is opened"""
        tag = self.tag(it)
        if it.childCount() and tag != TABLE and tag != QUERY:
            return
        try:
            self.tree.blockSignals(True)
            if tag == FIELDS:
                for f in self.db.get("TableDefs", it.parent().text(0)).each("Fields"):
                    self.node(it, f.get("Name"), FIELD, FIELD)
            elif tag == FIELD:
                f = self.db.get("TableDefs", it.parent().parent().text(0)).get("Fields", it.text(0))
                for p in f.each("Properties"):
                    n = p.get("Name")
                    if n != "Value":
                        self.node(it, f"{n}={prop_value(p)}", PROPERTY, PROPERTY)
            elif tag == INDEXES:
                for x in self.db.get("TableDefs", it.parent().text(0)).each("Indexes"):
                    self.node(it, x.get("Name"), INDEX, INDEX)
            elif tag == INDEX:
                x = self.db.get("TableDefs", it.parent().parent().text(0)).get("Indexes", it.text(0))
                for p in x.each("Properties"):
                    self.node(it, f"{p.get('Name')}={prop_value(p)}", PROPERTY, PROPERTY)
            elif tag == PROPERTIES and not it.childCount():
                par = it.parent()
                if par is None:
                    owner = self.db
                elif self.tag(par) == TABLE:
                    owner = self.db.get("TableDefs", par.text(0))
                elif self.tag(par) == QUERY:
                    owner = self.db.get("QueryDefs", par.text(0))
                else:
                    owner = None
                if owner is not None:
                    for p in owner.each("Properties"):
                        v = prop_value(p)[:50]
                        self.node(it, f"{p.get('Name')}={v}", PROPERTY, PROPERTY)
        except RtError as e:
            self.vd.show_error(e)
        finally:
            self.tree.blockSignals(False)
        if it.childCount():
            it.setExpanded(True)
        else:
            it.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.DontShowIndicatorWhenChildless)

    def double_click(self, it, _col):
        self.node2 = it
        if self.tag(it) == PROPERTY:
            self.do_edit()
        else:
            self.do_open()

    def popup_at(self, it, gpos):
        """the menu of the right button: what fits the node (tvDatabase_MouseUp)"""
        self.node2 = it
        vd = self.vd
        tag = self.tag(it)
        show = {"open": False, "edit": False, "copy": False, "design": False}
        en_rename = en_delete = False
        if it is not None:
            if tag == TABLE:
                show.update(open=True, copy=True, design=True)
                en_rename = en_delete = True
            elif tag == QUERY:
                show.update(open=True, design=True)
                en_rename = en_delete = True
            elif tag in (INDEX, FIELD):
                en_rename = en_delete = True
            elif tag == PROPERTY:
                show.update(edit=True)
        vd.pu_open.setVisible(show["open"])
        vd.pu_edit.setVisible(show["edit"])
        vd.pu_copy.setVisible(show["copy"])
        vd.pu_design.setVisible(show["design"])
        vd.pu_rename.setVisible(it is not None and vd.s.datatype == C.MSACCESS)
        vd.pu_delete.setVisible(it is not None)
        vd.pu_bar1.setVisible(it is not None)
        vd.pu_rename.setEnabled(en_rename)
        vd.pu_delete.setEnabled(en_delete)
        vd.pu_removeall.setVisible(tag == TABLE)
        if not vd.quiet:
            vd.popup.exec(gpos)
        return vd.popup

    # ---------------------------------------------------------------- the commands of the menu
    def selected(self):
        return self.node2 if self.node2 is not None else self.tree.currentItem()

    def do_open(self):
        it = self.selected()
        if it is None:
            return None
        if self.tag(it) == TABLE:
            return open_table(self.vd, it.text(0))
        if self.tag(it) == QUERY:
            return open_query(self.vd, it.text(0), False)
        return None

    def do_design(self):
        it = self.selected()
        if it is None:
            return None
        if self.tag(it) == TABLE:
            from .tblstru import TableStructure
            d = TableStructure(self.vd, it.text(0))
            if not self.vd.quiet:
                d.exec()
            return d
        if self.tag(it) == QUERY and self.vd.sqlwin is not None:
            try:
                self.vd.sqlwin.text.setPlainText(self.db.get("QueryDefs", it.text(0)).get("SQL"))
            except RtError as e:
                self.vd.show_error(e)
        return None

    def prop_object(self, it):
        """the Property of a node "Name=Value" (mnuDBPUEdit)"""
        name = it.text(0).split("=", 1)[0]
        par = it.parent()
        gp = par.parent() if par is not None else None
        if gp is None:
            return self.db.get("Properties", name)
        t = self.tag(gp)
        if t == TABLE:
            return self.db.get("TableDefs", gp.text(0)).get("Properties", name)
        if t == QUERY:
            return self.db.get("QueryDefs", gp.text(0)).get("Properties", name)
        if t == FIELDS:
            return self.db.get("TableDefs", gp.parent().text(0)).get("Fields", par.text(0)).get("Properties", name)
        if t == INDEXES:
            return self.db.get("TableDefs", gp.parent().text(0)).get("Indexes", par.text(0)).get("Properties", name)
        return None

    def do_edit(self):
        it = self.selected()
        if it is None or self.tag(it) != PROPERTY:
            return None
        try:
            p = self.prop_object(it)
            if p is None:
                return None
            v = p.get("Value")
            try:
                p.let("Value", v)
            except RtError:
                self.vd.say(f"'{p.get('Name')}' Property is Read Only!", "warn")
                return None
            new = self.vd.ask_text(f"{p.get('Name')}:", "Property", C.text_of(v))
            if new == C.text_of(v) or (self.vd.quiet is False and new == ""):
                return None
            p.let("Value", new)
            it.setText(0, f"{p.get('Name')}={prop_value(p)}")
            if p.get("Name") == "Name" and it.parent() is not None and it.parent().parent() is not None:
                it.parent().parent().setText(0, C.text_of(p.get("Value")))
        except RtError as e:
            self.vd.show_error(e)
        return None

    def do_rename(self):
        it = self.selected()
        if it is None or self.tag(it) not in (TABLE, QUERY, INDEX, FIELD):
            return
        self._editing = (it, it.text(0))
        it.setFlags(it.flags() | Qt.ItemFlag.ItemIsEditable)
        if not self.vd.quiet:
            self.tree.editItem(it, 0)

    def label_edited(self, it, _col):
        """tvDatabase_AfterLabelEdit: the object takes the new name"""
        if self._editing is None or self._editing[0] is not it:
            return
        old = self._editing[1]
        new = it.text(0)
        self._editing = None
        it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
        if new == old:
            return
        try:
            t = self.tag(it)
            if t == TABLE:
                self.db.get("TableDefs", old).let("Name", new)
            elif t == QUERY:
                self.db.get("QueryDefs", old).let("Name", new)
            elif t == INDEX:
                self.db.get("TableDefs", it.parent().parent().text(0)).get("Indexes", old).let("Name", new)
            elif t == FIELD:
                self.db.get("TableDefs", it.parent().parent().text(0)).get("Fields", old).let("Name", new)
        except RtError as e:
            self.vd.say(e.description)
            self.tree.blockSignals(True)
            it.setText(0, old)
            self.tree.blockSignals(False)

    def do_delete(self):
        it = self.selected()
        if it is None:
            return
        t = self.tag(it)
        try:
            if t == TABLE:
                if self.vd.ask("Delete Table?"):
                    self.db.get("TableDefs").call("Delete", it.text(0))
                    self._remove(it)
            elif t == QUERY:
                if self.vd.ask("Delete QueryDef?"):
                    self.db.get("QueryDefs").call("Delete", it.text(0))
                    self._remove(it)
            elif t == FIELD:
                if self.vd.ask("Delete Field?"):
                    self.db.get("TableDefs", it.parent().parent().text(0)).get("Fields").call("Delete", it.text(0))
                    self._remove(it)
            elif t == INDEX:
                if self.vd.ask("Delete Index?"):
                    self.db.get("TableDefs", it.parent().parent().text(0)).get("Indexes").call("Delete", it.text(0))
                    self._remove(it)
        except RtError as e:
            self.vd.show_error(e)

    def _remove(self, it):
        par = it.parent()
        if par is None:
            self.tree.takeTopLevelItem(self.tree.indexOfTopLevelItem(it))
        else:
            par.removeChild(it)
        self.node2 = None

    def do_copy_structure(self):
        from .dialogs import CopyStructureDialog
        it = self.selected()
        d = CopyStructureDialog(self.vd, it.text(0) if it is not None and self.tag(it) == TABLE else "")
        if not self.vd.quiet:
            d.exec()
        return d

    def do_remove_all(self):
        it = self.selected()
        if it is None or self.tag(it) != TABLE:
            return
        name = it.text(0)
        if not self.vd.ask(f"Delete All Records in Table? '{name}'"):
            return
        try:
            q = name if name.startswith("[") else f"[{name}]"
            if self.vd.s.datatype == C.SQLDB:
                self.db.call("Execute", "delete from " + name, C.dbSQLPassThrough)
            else:
                self.db.call("Execute", "delete from " + q)
            n = self.db.get("RecordsAffected")
            if n > 0:
                self.vd.say(f"Rows deleted: {n}", "warn")
                if self.vd.s.trans_pending:
                    self.vd.s.changed = True
        except RtError as e:
            self.vd.show_error(e)

    def do_new_table(self):
        from .tblstru import TableStructure
        d = TableStructure(self.vd, None)
        if not self.vd.quiet:
            d.exec()
        return d


# ---------------------------------------------------------------- opening a table, a query
def open_table(vd, name):
    """OpenTable: a Recordset of the type of the toolbar in a form of the kind of the toolbar"""
    from .recforms import record_form
    s, db = vd.s, vd.s.db
    rstype = vd.prefs.get("rstype", C.RS_DYNASET)
    try:
        att = 0
        if s.datatype == C.MSACCESS:
            att = C.attached(db.get("TableDefs", name))
            if att and rstype == C.RS_TABLE:
                if not vd.ask(C.MSG[10]):
                    return None
                vd.set_rstype(C.RS_DYNASET)
                rstype = C.RS_DYNASET
        msg = (C.MSG[11] if rstype == C.RS_DYNASET else C.MSG[12]) if att else \
            {C.RS_TABLE: C.MSG[13], C.RS_DYNASET: C.MSG[14], C.RS_SNAPSHOT: C.MSG[15],
             C.RS_PASSTHRU: C.MSG[16]}[rstype]
        vd.msgbar(msg, True)
        if rstype == C.RS_TABLE:
            rs = db.call("OpenRecordset", name, C.dbOpenTable)
            cap = "Table:"
        elif rstype == C.RS_DYNASET:
            rs = db.call("OpenRecordset", name, C.dbOpenDynaset)
            cap = "Dynaset:"
        elif rstype == C.RS_SNAPSHOT:
            rs = db.call("OpenRecordset", name, C.dbOpenSnapshot)
            cap = "Snapshot:"
        else:
            rs = db.call("OpenRecordset", name, C.dbOpenSnapshot, C.dbSQLPassThrough)
            cap = "Passthrough Snapshot:"
    except RtError as e:
        vd.show_error(e)
        return None
    w = record_form(vd, rs, cap + name, table=rstype == C.RS_TABLE, source=name)
    vd.msgbar()
    return w


def action_type(sql):
    """ActionQueryType: the kind of an action query ("" for one that returns rows)"""
    t = (sql or "").lstrip().upper()
    for k, n in (("DELETE", "Delete"), ("UPDATE", "Update"), ("INSERT", "Append"), ("SELECT", "")):
        if t.startswith(k):
            if k == "SELECT" and " INTO " in t.split(" FROM ")[0]:
                return "Make Table"
            return n
    if t.startswith(("CREATE", "DROP", "ALTER")):
        return "Data Definition"
    return ""


def open_query(vd, name, temp):
    """OpenQuery: a query that returns rows in a form; an action query run (asked first). The SQL of the SQL
    Statement window (temp): a query without name, run - one that returns rows ("Cannot execute a select query.")
    is opened instead"""
    from .recforms import record_form
    s, db = vd.s, vd.s.db
    rstype = vd.prefs.get("rstype", C.RS_DYNASET)
    tried = False
    try:
        if temp:
            qd = db.call("CreateQueryDef", "", name)
            if vd.ask(C.MSG[17], default=False):
                cn = vd.ask_text(C.MSG[18])
                if cn:
                    qd.let("Connect", cn)
            sql, kind, returns = name, "", False
        else:
            qd = db.get("QueryDefs", name)
            sql = qd.get("SQL")
            kind = action_type(sql)
            returns = qd.get("Type") in (0, 16)
        if returns and rstype == C.RS_TABLE:
            if not vd.ask(C.MSG[19]):
                return None
            vd.set_rstype(C.RS_DYNASET)
            rstype = C.RS_DYNASET
        if not returns:
            if not kind or vd.ask(C.MSG[23] + kind + C.MSG[24]):
                set_params(vd, qd)
                vd.msgbar(C.MSG[25], True)
                try:
                    qd.call("Execute")
                    if db.get("RecordsAffected") > 0 and s.trans_pending:
                        s.changed = True
                    vd.msgbar()
                    return None
                except RtError as e:
                    if e.number not in (3065, 3078):
                        raise
                    tried = True                # a query that returns rows: opened
            else:
                return None
        else:
            set_params(vd, qd)
        if rstype == C.RS_SNAPSHOT:
            vd.msgbar(C.MSG[20], True)
            rs = qd.call("OpenRecordset", C.dbOpenSnapshot)
        else:
            vd.msgbar(C.MSG[21], True)
            rs = qd.call("OpenRecordset", C.dbOpenDynaset)
        cap = C.MSG[22] if len(sql) > 50 else sql
        w = record_form(vd, rs, cap, table=False, source=sql)
        vd.msgbar()
        return w
    except RtError as e:
        del tried
        vd.show_error(e)
    vd.msgbar()
    return None


def set_params(vd, qd):
    """SetQDFParams: the value of each parameter asked"""
    for prm in qd.each("Parameters"):
        v = vd.ask_text(C.MSG[29], f"'{prm.get('Name')}':")
        prm.let("Value", v)
