"""The Query Designer (phase 12.3): the SQL Builder of a DataEnvironment command ("Design: Command1") and the windows
of the Data View ("Run Table: Customers", the Design of a view), as the Visual Database Tools of VB6 have them (VB6
oracles Claude Enviroment/oracles/12.3): four panes one under the other - Diagram (the tables, their columns with
checkboxes, the joins), Grid (Column, Alias, Table, Output, Sort Type, Sort Order, [Group By], Criteria, Or...), SQL and
Results; the Query and Diagram menus of the IDE, View > Table / Zoom / Show Panes, the context menus of the panes, the
Properties of the query. The SQL is made and read back by querymodel; the database is reached through LinBASIC
Studio's ADO (dbinfo: lbdbinfo -ado-sql / -ado-exec)."""
import re

from PyQt6.QtCore import Qt, QRectF, QPointF, QMimeData, pyqtSignal, QEvent, QTimer
from PyQt6.QtGui import QBrush, QColor, QFont, QPainter, QPainterPath, QPen, QPolygonF, QTextOption
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QGraphicsView, QGraphicsScene,
                             QGraphicsRectItem, QGraphicsPathItem, QGraphicsProxyWidget, QGraphicsTextItem,
                             QGraphicsSimpleTextItem, QGraphicsItem, QListWidget, QListWidgetItem, QTableWidget,
                             QTableWidgetItem, QPlainTextEdit, QHeaderView, QAbstractItemView, QStyledItemDelegate,
                             QComboBox, QMenu, QMessageBox, QDialog, QTabWidget, QFormLayout, QLineEdit, QCheckBox,
                             QDialogButtonBox, QInputDialog, QFontDialog, QLabel, QStyle)

from . import dbinfo
from . import querymodel as QM

TABLE_MIME = "application/x-lbstudio-dbtable"     # a table dragged from the Data View: "table name"
PANES = ("diagram", "grid", "sql", "results")
NULL_TEXT = "<NULL>"
ADO_TYPE_NAMES = {2: "smallint", 3: "integer", 4: "real", 5: "float", 6: "currency", 7: "datetime", 11: "bit",
                  17: "tinyint", 72: "uniqueidentifier", 128: "binary", 129: "char", 130: "nchar", 131: "numeric",
                  135: "datetime", 200: "varchar", 201: "text", 202: "nvarchar", 203: "ntext", 204: "varbinary",
                  205: "image"}
NUMERIC_TYPES = {2, 3, 4, 5, 6, 11, 14, 16, 17, 18, 19, 20, 21, 131, 139}
DATE_TYPES = {7, 133, 134, 135}
ZOOMS = [200, 150, 100, 75, 50, 25, 10]
TABLE_VIEWS = ["Column Properties", "Column Names", "Keys", "Name Only", "Custom"]
CUSTOM_PROPS = ["Column Name", "Data Type", "Length", "Precision", "Allow Nulls"]
BOX_W, BOX_H, TITLE_H = 150, 108, 20


def literal(value, adotype):
    """a value typed in the Results pane as a SQL literal of Jet"""
    if value is None or value == "" or value == NULL_TEXT:
        return "NULL"
    if adotype in NUMERIC_TYPES:
        v = value.strip()
        if adotype == 11:
            return "True" if v.lower() in ("true", "-1", "1", "yes") else "False"
        if re.fullmatch(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", v):
            return v
    if adotype in DATE_TYPES:
        return f"#{value.strip()}#"
    return "'" + value.replace("'", "''") + "'"


# ====================================================================== the Diagram pane
class ColumnList(QListWidget):
    """the columns of a table in the diagram: * (All Columns) and the columns with checkboxes; a column of a sort or
    a filter shows its sign at its right (VB6)"""

    def __init__(self, box):
        super().__init__()
        self.box = box
        self.setFrameShape(QListWidget.Shape.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(lambda pos: box.designer.table_menu(box, self.mapToGlobal(pos)))
        self.setItemDelegate(_MarkDelegate(self))

    def mousePressEvent(self, e):
        self.box.designer.diagram.scene().clearSelection()
        self.box.setSelected(True)
        super().mousePressEvent(e)


class _MarkDelegate(QStyledItemDelegate):
    def paint(self, p, opt, idx):
        super().paint(p, opt, idx)
        mark = idx.data(Qt.ItemDataRole.UserRole + 1) or ""
        if mark:
            p.save()
            p.setPen(QColor("#000000"))
            f = QFont(opt.font)
            f.setPointSizeF(max(6.0, f.pointSizeF() - 1))
            p.setFont(f)
            r = opt.rect.adjusted(0, 0, -3, 0)
            p.drawText(r, int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter), mark)
            p.restore()


class TableBox(QGraphicsRectItem):
    """a table of the Diagram pane: its title bar (it moves the table) and its list of columns"""

    def __init__(self, designer, qt):
        super().__init__(0, 0, BOX_W, BOX_H)
        self.designer = designer
        self.qt = qt
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable | QGraphicsItem.GraphicsItemFlag.ItemIsSelectable |
                      QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges)
        self.setBrush(QBrush(QColor("#d4d0c8")))
        self.setPen(QPen(QColor("#808080")))
        self.list = ColumnList(self)
        self.proxy = QGraphicsProxyWidget(self)
        self.proxy.setWidget(self.list)
        self.list.itemChanged.connect(self._checked)
        self.setPos(qt.x, qt.y)
        self.height = getattr(qt, "height", BOX_H)
        self.fill()

    def title(self):
        return self.qt.name + (f" ({self.qt.alias})" if self.qt.alias else "")

    def fill(self):
        d = self.designer
        view = d.table_view
        self.list.blockSignals(True)
        if self.qt.collapsed or view == "Name Only":
            self.proxy.hide()
            self.setRect(0, 0, BOX_W, TITLE_H + 2)
        else:
            self.proxy.show()
            h = max(TITLE_H + 30, self.height)
            self.setRect(0, 0, BOX_W, h)
            self.proxy.setGeometry(QRectF(2, TITLE_H + 1, BOX_W - 4, h - TITLE_H - 3))
            cols = sorted(self.qt.columns, key=str.lower)
            if view == "Keys":
                keys = {k.lower() for k in d.keys_of(self.qt.name)}
                cols = [c for c in cols if c.lower() in keys]
            same = [self.list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.list.count())] == ["*"] + cols
            if not same:
                self.list.clear()
            for n, c in enumerate(["*"] + cols):
                text = "* (All Columns)" if c == "*" else c
                if c != "*" and view in ("Column Properties", "Custom"):
                    info = d.field_info(self.qt.name, c)
                    if info is not None:
                        props = d.custom_props if view == "Custom" else ["Column Name", "Data Type", "Length"]
                        parts = []
                        for pname in props:
                            parts.append({"Column Name": c, "Data Type": ADO_TYPE_NAMES.get(info[1], str(info[1])),
                                          "Length": str(info[2]), "Precision": str(info[3]),
                                          "Allow Nulls": ""}.get(pname, ""))
                        text = "   ".join(x for x in parts if x) or c
                it = self.list.item(n) if same else QListWidgetItem()
                it.setText(text)
                it.setData(Qt.ItemDataRole.UserRole, c)
                it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                on = d.model.checked(self.qt.ref, c) if c != "*" else (
                    d.model.checked(self.qt.ref, "*") or (d.model.all_columns and len(d.model.tables) >= 1))
                it.setCheckState(Qt.CheckState.Checked if on else Qt.CheckState.Unchecked)
                r = d.model.row_of(self.qt.ref, c)
                mark = ""
                if r is not None:
                    if r.criteria.strip() or any(o.strip() for o in r.ors):
                        mark += "▽"         # (a filter)
                    if r.sort:
                        mark += " A↓" if r.sort == "Ascending" else " Z↓"
                    if d.model.group_by and r.group == "Group By" and c != "*":
                        mark += " [≡]"
                it.setData(Qt.ItemDataRole.UserRole + 1, mark.strip())
                if not same:
                    self.list.addItem(it)
        self.list.blockSignals(False)
        self.update()

    def _checked(self, it):
        self.designer.toggle_column(self.qt.ref, it.data(Qt.ItemDataRole.UserRole),
                                    it.checkState() == Qt.CheckState.Checked)

    def paint(self, p, opt, w=None):
        r = self.rect()
        sel = self.isSelected()
        p.setPen(QPen(QColor("#000080") if sel else QColor("#808080"), 2 if sel else 1))
        p.setBrush(QBrush(QColor("#d4d0c8")))
        p.drawRect(r)
        title = QRectF(1, 1, r.width() - 2, TITLE_H - 1)
        p.fillRect(title, QColor("#0a246a") if sel else QColor("#d4d0c8"))
        p.setPen(QColor("#ffffff") if sel else QColor("#000000"))
        p.drawText(title.adjusted(20, 0, -18, 0), int(Qt.AlignmentFlag.AlignVCenter), self.title())
        # (the picture of a table, the minimize button)
        p.setPen(QColor("#000000") if not sel else QColor("#ffffff"))
        p.drawRect(QRectF(5, 5, 11, 9))
        p.drawLine(QPointF(5, 8), QPointF(16, 8))
        p.drawLine(QPointF(9, 5), QPointF(9, 14))
        b = self.button_rect()
        p.fillRect(b, QColor("#d4d0c8"))
        p.setPen(QColor("#000000"))
        p.drawRect(b)
        p.drawLine(QPointF(b.left() + 3, b.bottom() - 3), QPointF(b.right() - 3, b.bottom() - 3))

    def button_rect(self):
        return QRectF(self.rect().width() - 17, 3, 14, 14)

    def mousePressEvent(self, e):
        if self.button_rect().contains(e.pos()):
            self.qt.collapsed = not self.qt.collapsed
            self.fill()
            self.designer.update_joins()
            return
        if e.button() == Qt.MouseButton.RightButton:
            self.designer.diagram.scene().clearSelection()
            self.setSelected(True)
            self.designer.table_menu(self, e.screenPos())
            return
        super().mousePressEvent(e)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            self.qt.x, self.qt.y = int(self.pos().x()), int(self.pos().y())
            self.designer.update_joins()
        return super().itemChange(change, value)

    def row_y(self, column):
        """where a column is (the joins start at its row), in the scene"""
        if self.qt.collapsed or self.designer.table_view == "Name Only" or not self.proxy.isVisible():
            return self.scenePos().y() + TITLE_H / 2
        for i in range(self.list.count()):
            if (self.list.item(i).data(Qt.ItemDataRole.UserRole) or "").lower() == column.lower():
                r = self.list.visualItemRect(self.list.item(i))
                y = TITLE_H + 1 + r.center().y()
                y = max(TITLE_H + 4, min(y, self.rect().height() - 4))
                return self.scenePos().y() + y
        return self.scenePos().y() + TITLE_H / 2


class JoinLine(QGraphicsPathItem):
    """a join: a line from a column of one table to a column of the other, a diamond in its middle (VB6); its
    label (Show Relationship Labels)"""

    def __init__(self, designer, join):
        super().__init__()
        self.designer = designer
        self.join = join
        self.setPen(QPen(QColor("#000000"), 1))
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
        self.setZValue(-1)
        self.label = QGraphicsSimpleTextItem(self)
        self.mid = QPointF()

    def relayout(self):
        d = self.designer
        a, b = d.box_of(self.join.left), d.box_of(self.join.right)
        if a is None or b is None:
            self.hide()
            return
        self.show()
        ar, br = a.sceneBoundingRect(), b.sceneBoundingRect()
        ay, by = a.row_y(self.join.lcol), b.row_y(self.join.rcol)
        if ar.right() < br.left():
            p1, p2 = QPointF(ar.right(), ay), QPointF(br.left(), by)
            s1, s2 = 1, -1
        elif br.right() < ar.left():
            p1, p2 = QPointF(ar.left(), ay), QPointF(br.right(), by)
            s1, s2 = -1, 1
        else:
            p1, p2 = QPointF(ar.right(), ay), QPointF(br.right(), by)
            s1 = s2 = 1
        path = QPainterPath(p1)
        q1, q2 = QPointF(p1.x() + 12 * s1, p1.y()), QPointF(p2.x() + 12 * s2, p2.y())
        path.lineTo(q1)
        self.mid = (q1 + q2) / 2
        path.lineTo(q2)
        path.lineTo(p2)
        dm = 6
        path.addPolygon(QPolygonF([QPointF(self.mid.x() - dm, self.mid.y()), QPointF(self.mid.x(), self.mid.y() - dm),
                                   QPointF(self.mid.x() + dm, self.mid.y()), QPointF(self.mid.x(), self.mid.y() + dm),
                                   QPointF(self.mid.x() - dm, self.mid.y())]))
        # (an outer join: the side of all the rows has a square end)
        for on, pt in ((self.join.all_from(self.join.left), q1), (self.join.all_from(self.join.right), q2)):
            if on:
                path.addRect(QRectF(pt.x() - 3, pt.y() - 3, 6, 6))
        self.setPath(path)
        if d.show_labels:
            self.label.setText(f"{d.model.table(self.join.left).name if d.model.table(self.join.left) else ''}"
                               f"{d.model.table(self.join.right).name if d.model.table(self.join.right) else ''}")
            self.label.setPos(self.mid + QPointF(8, -16))
            self.label.show()
        else:
            self.label.hide()

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.RightButton:
            self.designer.join_menu(self, e.screenPos())
            return
        super().mousePressEvent(e)


class Annotation(QGraphicsTextItem):
    """Diagram > New Text Annotation: a text of its own on the diagram"""

    def __init__(self, text="Text annotation"):
        super().__init__(text)
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable | QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
        self.setTextInteractionFlags(Qt.TextInteractionFlag.TextEditorInteraction)
        self.setTextWidth(160)

    def paint(self, p, opt, w=None):
        p.setPen(QPen(QColor("#808080"), 1, Qt.PenStyle.DashLine))
        p.setBrush(QBrush(QColor("#ffffe1")))
        p.drawRect(self.boundingRect())
        super().paint(p, opt, w)


class DiagramView(QGraphicsView):
    def __init__(self, designer):
        super().__init__()
        self.designer = designer
        self.setScene(QGraphicsScene(self))
        self.scene().setSceneRect(0, 0, 1400, 700)
        self.setBackgroundBrush(QBrush(QColor("#ffffff")))
        self.setRenderHint(QPainter.RenderHint.Antialiasing, False)
        self.setAcceptDrops(True)
        self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
        self.page_breaks = False

    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat(TABLE_MIME):
            e.acceptProposedAction()
        else:
            super().dragEnterEvent(e)

    def dragMoveEvent(self, e):
        if e.mimeData().hasFormat(TABLE_MIME):
            e.acceptProposedAction()
        else:
            super().dragMoveEvent(e)

    def dropEvent(self, e):
        if e.mimeData().hasFormat(TABLE_MIME):
            name = bytes(e.mimeData().data(TABLE_MIME)).decode("utf-8")
            pos = self.mapToScene(e.position().toPoint())
            self.designer.add_table(name, int(pos.x()), int(pos.y()))
            e.acceptProposedAction()
        else:
            super().dropEvent(e)

    def contextMenuEvent(self, e):
        if self.itemAt(e.pos()) is not None:
            return super().contextMenuEvent(e)
        self.designer.pane_menu("diagram", e.globalPos())

    def drawForeground(self, p, rect):
        if not self.page_breaks:
            return
        p.setPen(QPen(QColor("#a0a0a0"), 1, Qt.PenStyle.DashLine))
        pw, ph = 794, 1123                      # (an A4 page at 96 dpi)
        x = pw
        while x < rect.right():
            if x > rect.left():
                p.drawLine(QPointF(x, rect.top()), QPointF(x, rect.bottom()))
            x += pw
        y = ph
        while y < rect.bottom():
            if y > rect.top():
                p.drawLine(QPointF(rect.left(), y), QPointF(rect.right(), y))
            y += ph


# ====================================================================== the Grid pane
class GridDelegate(QStyledItemDelegate):
    """the lists of the grid's cells: Column, Table, Sort Type, Group By"""

    def __init__(self, designer):
        super().__init__()
        self.d = designer

    def createEditor(self, parent, opt, idx):
        key = self.d.grid_keys[idx.column()]
        if key in ("column", "table", "sort", "group"):
            cb = QComboBox(parent)
            cb.setEditable(key == "column")
            cb.addItems(self.d.choices(key, idx.row()))
            return cb
        return super().createEditor(parent, opt, idx)

    def setEditorData(self, ed, idx):
        if isinstance(ed, QComboBox):
            ed.setCurrentText(idx.data() or "")
        else:
            super().setEditorData(ed, idx)

    def setModelData(self, ed, model, idx):
        if isinstance(ed, QComboBox):
            model.setData(idx, ed.currentText())
        else:
            super().setModelData(ed, model, idx)


GRID_COLUMNS = {   # the columns of the Grid pane for each type of query
    "Select": ["column", "alias", "table", "output", "sort", "order", "group", "criteria", "or1", "or2", "or3"],
    "Make Table": ["column", "alias", "table", "output", "sort", "order", "group", "criteria", "or1", "or2", "or3"],
    "Insert": ["column", "alias", "table", "output", "append", "group", "criteria", "or1", "or2", "or3"],
    "Insert Values": ["column", "value"],
    "Update": ["column", "table", "value", "criteria", "or1", "or2", "or3"],
    "Delete": ["column", "table", "criteria", "or1", "or2", "or3"],
}                  # (VB6 12.3/q18, q19, q21, q24, q26)
ONE_TABLE = {      # Change Type with more tables: the table of the query (VB6 12.3/q17, q33, q34)
    "Update": ("Update Table", "An UPDATE query updates rows in only one table.\n\nSelect a table from the list and "
               "click OK. Or press Cancel to return to your original query."),
    "Delete": ("Delete Table", "A DELETE query deletes rows from only one table.\n\nSelect a table from the list and "
               "click OK. Or click Cancel to return to your original query."),
    "Insert Values": ("Insert Values", "An INSERT query inserts a row in only one table.\n\nSelect a table from the "
                      "list and click OK. Or press Cancel to return to your original query."),
}
GRID_TITLES = {"column": "Column", "alias": "Alias", "table": "Table", "output": "Output", "sort": "Sort Type",
               "order": "Sort Order", "group": "Group By", "criteria": "Criteria", "or1": "Or...", "or2": "Or...",
               "or3": "Or...", "value": "New Value", "append": "Append"}


# ====================================================================== the designer
class QueryDesigner(QWidget):
    """a query of a connection: the four panes; on_save(sql) when the SQL goes back to its command"""
    modified = pyqtSignal()

    def __init__(self, main, connect, sql="", title="Query", on_save=None, panes=PANES, base_dir=None,
                 conn_name="", query_name=""):
        super().__init__()
        self.main = main
        self.base_dir = base_dir
        try:
            from .dsrdesign import design_cs
            self.cs = design_cs(connect, base_dir) if base_dir else connect
        except Exception:                       # noqa: BLE001
            self.cs = connect
        self.conn_name = conn_name
        self.title = title
        self.query_name = query_name or title.split(": ", 1)[-1]
        self.on_save = on_save
        self.dirty = False
        self.model = QM.QueryModel()
        self.table_view = "Column Names"
        self.custom_props = ["Column Name", "Data Type", "Length"]
        self.show_labels = False
        self.zoom = 100
        self._cols, self._fields, self._keys, self._rels = {}, {}, {}, None
        self.result_cols = []
        self.result_table = None                # the table of the rows of the Results pane (they can be changed)
        self.result_keys = []
        self.new_row = None
        self.stale = False
        self._filling = False
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.split = QSplitter(Qt.Orientation.Vertical)
        lay.addWidget(self.split)
        # -- Diagram
        self.diagram = DiagramView(self)
        self.split.addWidget(self.diagram)
        self.boxes, self.lines, self.notes = [], [], []
        # -- Grid
        self.grid = QTableWidget(0, 0)
        self.grid.setItemDelegate(GridDelegate(self))
        self.grid.verticalHeader().setDefaultSectionSize(18)
        self.grid.verticalHeader().setFixedWidth(16)
        self.grid.horizontalHeader().setDefaultSectionSize(90)
        self.grid.setEditTriggers(QAbstractItemView.EditTrigger.AllEditTriggers)
        self.grid.itemChanged.connect(self._grid_changed)
        self.grid.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.grid.customContextMenuRequested.connect(lambda pos: self.pane_menu("grid", self.grid.mapToGlobal(pos)))
        self.split.addWidget(self.grid)
        # -- SQL
        self.sql = QPlainTextEdit()
        f = QFont("Courier New")
        f.setStyleHint(QFont.StyleHint.Monospace)
        self.sql.setFont(f)
        self.sql.setWordWrapMode(QTextOption.WrapMode.NoWrap)
        self.sql.textChanged.connect(self._sql_typed)
        self.sql.installEventFilter(self)
        self.sql.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.sql.customContextMenuRequested.connect(lambda pos: self.pane_menu("sql", self.sql.mapToGlobal(pos)))
        self.sql_typed = False
        self.split.addWidget(self.sql)
        # -- Results
        self.results = QTableWidget(0, 0)
        self.results.verticalHeader().setDefaultSectionSize(18)
        self.results.verticalHeader().setFixedWidth(16)
        self.results.itemChanged.connect(self._result_edited)
        self.results.currentCellChanged.connect(self._result_moved)
        self.results.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results.customContextMenuRequested.connect(
            lambda pos: self.pane_menu("results", self.results.mapToGlobal(pos)))
        self.split.addWidget(self.results)
        self.split.setSizes([230, 130, 110, 160])
        self.panes = {"diagram": self.diagram, "grid": self.grid, "sql": self.sql, "results": self.results}
        for k in PANES:
            self.panes[k].setVisible(k in panes)
        if sql.strip():
            self.set_sql(sql)
        else:
            self.refresh()
        self.dirty = False

    # ---------------------------------------------------------------- the database
    def columns_of(self, table):
        k = table.lower()
        if k not in self._cols:
            self._cols[k] = dbinfo.ado_columns(self.cs, table)
        return self._cols[k]

    def field_info(self, table, column):
        k = table.lower()
        if k not in self._fields:
            try:
                self._fields[k] = {f[0].lower(): f for f in dbinfo.ado_fields(self.cs, 2, table)}
            except dbinfo.AdoError:
                self._fields[k] = {}
        return self._fields[k].get(column.lower())

    def keys_of(self, table):
        k = table.lower()
        if k not in self._keys:
            self._keys[k] = dbinfo.ado_keys(self.cs, table)
        return self._keys[k]

    def relations(self):
        if self._rels is None:
            self._rels = dbinfo.ado_relations(self.cs)
        return self._rels

    # ---------------------------------------------------------------- the model to the panes
    def changed(self, from_sql=False):
        if not from_sql:                        # changed in the panes: the criteria in parentheses (VB6 12.3/q15)
            for r in self.model.rows:
                r.bare.clear()
        self.dirty = True
        self.mark_stale()
        self.refresh(keep_sql=from_sql)
        self.modified.emit()

    def refresh(self, keep_sql=False):
        self.fill_diagram()
        self.fill_grid()
        if not keep_sql:
            self._filling = True
            self.sql.setPlainText(self.model.sql())
            self._filling = False
            self.sql_typed = False
        ok = self.model.parsed
        self.diagram.setEnabled(ok)
        self.grid.setEnabled(ok)

    def box_of(self, ref):
        return next((b for b in self.boxes if b.qt.ref.lower() == (ref or "").lower()), None)

    def fill_diagram(self):
        sc = self.diagram.scene()
        old = {b.qt.ref.lower(): b for b in self.boxes}
        keep = []
        for t in self.model.tables:
            b = old.pop(t.ref.lower(), None)
            if b is None or b.qt is not t:
                if b is not None:
                    sc.removeItem(b)
                b = TableBox(self, t)
                sc.addItem(b)
            else:
                b.fill()
            keep.append(b)
        for b in old.values():
            sc.removeItem(b)
        self.boxes = keep
        for ln in self.lines:
            sc.removeItem(ln)
        self.lines = []
        for j in self.model.joins:
            ln = JoinLine(self, j)
            sc.addItem(ln)
            self.lines.append(ln)
        self.update_joins()
        r = sc.itemsBoundingRect().adjusted(-20, -20, 200, 200)
        sc.setSceneRect(QRectF(0, 0, max(1400, r.right()), max(700, r.bottom())))

    def update_joins(self):
        for ln in self.lines:
            ln.relayout()

    def grid_columns(self):
        keys = list(GRID_COLUMNS.get(self.model.qtype, GRID_COLUMNS["Select"]))
        if not self.model.group_by and "group" in keys:
            keys.remove("group")
        return keys

    def fill_grid(self):
        self._filling = True
        g = self.grid
        keys = self.grid_columns()
        if keys != getattr(self, "grid_keys", None):        # (the items stay: one may be in its signal now)
            self.grid_keys = keys
            g.setColumnCount(len(keys))
            g.setHorizontalHeaderLabels([GRID_TITLES[k] for k in keys])
        rows = self.model.rows
        g.setRowCount(max(len(rows) + 1, 8))
        g.setVerticalHeaderLabels([""] * g.rowCount())      # (VB6: row selectors without numbers)
        for r in range(g.rowCount()):
            row = rows[r] if r < len(rows) else None
            for c, key in enumerate(self.grid_keys):
                it = g.item(r, c)
                if it is None:
                    it = QTableWidgetItem()
                    g.setItem(r, c, it)
                if key == "output":
                    it.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled |
                                Qt.ItemFlag.ItemIsSelectable)
                    it.setCheckState(Qt.CheckState.Checked if row is not None and row.output else Qt.CheckState.Unchecked)
                    it.setText("")
                else:
                    it.setFlags(Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                    it.setText(self.cell_text(row, key) if row is not None else "")
        self._filling = False

    def cell_text(self, row, key):
        if key == "column":
            return "*" if row.column == "*" else row.column
        if key == "table":
            t = self.model.table(row.table)
            return t.ref if t is not None else ""
        if key == "order":
            return str(row.order) if row.order else ""
        if key.startswith("or"):
            return row.ors[int(key[2]) - 1]
        return {"alias": row.alias, "sort": row.sort, "group": row.group if self.model.group_by else "",
                "criteria": row.criteria, "value": row.value, "append": row.append}.get(key, "")

    def choices(self, key, r):
        if key == "column":
            out = []
            for t in self.model.tables:
                pre = f"{t.ref}." if self.model.many() else ""
                out += [pre + "*"] + [pre + c for c in sorted(t.columns, key=str.lower)]
            return out
        if key == "table":
            return [t.ref for t in self.model.tables]
        if key == "sort":
            return ["Unsorted", "Ascending", "Descending"]
        if key == "group":
            return QM.GROUP_FUNCS
        return []

    # ---------------------------------------------------------------- the panes to the model
    def toggle_column(self, ref, column, on):
        """a checkbox of a table in the diagram"""
        self.model.check(ref, column, on)
        self.changed()

    def add_table(self, name, x=None, y=None):
        cols = self.columns_of(name)
        if not cols:
            try:
                kinds = {n.lower(): k for k, n in dbinfo.ado_objects(self.cs, True)}
            except dbinfo.AdoError:
                kinds = {}
            if name.lower() not in kinds:
                return None
        t = self.model.add_table(name, cols, x, y, self.relations())
        self.changed()
        return t

    def remove_table(self, ref):
        self.model.remove_table(ref)
        self.changed()

    def _grid_changed(self, it):
        if self._filling:
            return
        r, key = it.row(), self.grid_keys[it.column()]
        rows = self.model.rows
        text = it.text().strip()
        if r >= len(rows):                      # a new line: its column first
            if key == "output" or not text:
                return
            row = QM.Row("", "")
            rows.append(row)
            if key != "column":
                self._set_cell(row, key, text, it)
                self.changed()
                return
        row = rows[r] if r < len(rows) else rows[-1]
        if key == "column" and not text:        # the column taken away: the line goes
            rows.remove(row)
            self.changed()
            return
        self._set_cell(row, key, text, it)
        self.changed()

    def _set_cell(self, row, key, text, it):
        m = self.model
        if key == "column":
            tref, col = None, text
            mm = re.fullmatch(r"(?:(\[[^\]]+\]|[^.\s]+)\.)?(\[[^\]]+\]|[^.\s]+|\*)", text)
            if mm:
                tref = QM._uq(mm.group(1)) if mm.group(1) else (m.tables[0].ref if len(m.tables) == 1 else
                                                               (row.table or None))
                if tref is not None and m.table(tref) is not None:
                    t = m.table(tref)
                    c = QM._uq(mm.group(2))
                    if c == "*" or c.lower() in (x.lower() for x in t.columns):
                        row.table, row.column = t.ref, next((x for x in t.columns if x.lower() == c.lower()), c)
                        return
            row.table, row.column = "", col            # an expression
        elif key == "alias":
            row.alias = text
        elif key == "table":
            t = m.table(text)
            row.table = t.ref if t is not None else ""
        elif key == "output":
            row.output = it.checkState() == Qt.CheckState.Checked
        elif key == "sort":
            m.set_sort(row, "" if text in ("", "Unsorted") else text)
        elif key == "order":
            if text.isdigit() and row.sort:
                m.set_order(row, int(text))
        elif key == "group":
            row.group = text or "Group By"
        elif key == "criteria":
            row.criteria = text
        elif key.startswith("or"):
            row.ors[int(key[2]) - 1] = text
        elif key == "value":
            row.value = text
        elif key == "append":
            row.append = text

    def _sql_typed(self):
        if not self._filling:
            self.sql_typed = True
            self.dirty = True
            self.mark_stale()

    def eventFilter(self, obj, e):
        if obj is self.sql and e.type() == QEvent.Type.FocusOut and self.sql_typed:
            self.apply_sql()
        return super().eventFilter(obj, e)

    def apply_sql(self):
        """the SQL pane read back into the Diagram and Grid panes (they go grey when it cannot be shown in them)"""
        text = self.sql.toPlainText()
        self.sql_typed = False
        self.model.parse(text, self.columns_of)
        if self.model.parsed:
            for t in self.model.tables:
                if not t.columns:
                    t.columns = self.columns_of(t.name)
        self.refresh(keep_sql=True)
        self.modified.emit()

    def set_sql(self, text):
        self._filling = True
        self.sql.setPlainText(text)
        self._filling = False
        self.apply_sql()

    def current_sql(self):
        if self.sql_typed:
            self.apply_sql()
        return self.sql.toPlainText().strip()

    # ---------------------------------------------------------------- Query
    def is_select(self, sql=None):
        t = (sql if sql is not None else self.current_sql()).lstrip().upper()
        if self.model.parsed and self.model.qtype != "Select":
            return False
        return t.startswith("SELECT") and not re.search(r"(?is)^SELECT\b.*\bINTO\b", t) or t.startswith("TRANSFORM")

    def run(self):
        """Query > Run: the rows in the Results pane; an action query is run and says how many rows it changed"""
        sql = self.current_sql()
        if not sql:
            return None
        if not self.panes["results"].isVisible() and self.is_select(sql):
            self.show_pane("results", True)
        try:
            if self.is_select(sql):
                cols, rows, more = dbinfo.ado_query(self.cs, sql, 0)
                self.fill_results(cols, rows)
                return len(rows)
            dbinfo.ado_exec(self.cs, sql)
            self._cols.clear()
            self._fields.clear()
            text = "Query has executed successfully"          # (VB6 12.3/q29)
            self.say(text)
            return text
        except dbinfo.AdoError as e:
            self.say(str(e), error=True)
            return None

    def say(self, text, error=False):
        if getattr(self, "quiet", False):
            self.last_message = text
            return
        from . import APP_NAME
        (QMessageBox.warning if error else QMessageBox.information)(self, APP_NAME, text)

    def fill_results(self, cols, rows):
        self._filling = True
        r = self.results
        self.result_cols = cols
        r.clear()
        r.setColumnCount(len(cols))
        r.setHorizontalHeaderLabels([c[0] for c in cols])
        r.setRowCount(len(rows) + 1)
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                r.setItem(i, j, QTableWidgetItem(NULL_TEXT if v is None else v))
        for j in range(len(cols)):
            r.setItem(len(rows), j, QTableWidgetItem(""))
        r.setVerticalHeaderLabels([""] * (len(rows) + 1))
        r.setVerticalHeaderItem(len(rows), QTableWidgetItem("*"))
        # the rows can be changed: one table, its key among the columns
        self.result_table, self.result_keys = None, []
        if self.model.parsed and len(self.model.tables) == 1 and not self.model.group_by and not self.model.distinct:
            t = self.model.tables[0]
            keys = self.keys_of(t.name)
            names = [c[0].lower() for c in cols]
            if keys and all(k.lower() in names for k in keys):
                self.result_table, self.result_keys = t.name, keys
        ro = self.result_table is None
        for i in range(r.rowCount()):
            for j in range(r.columnCount()):
                it = r.item(i, j)
                if ro:
                    it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
        if ro:
            r.setRowCount(len(rows))
        self.new_row = None
        self.stale = False
        self._filling = False
        r.resizeColumnsToContents()

    def mark_stale(self):
        """the Results are of an older query: grey (VB6)"""
        if self.stale or self.results.rowCount() == 0:
            return
        self.stale = True
        self._filling = True
        for i in range(self.results.rowCount()):
            for j in range(self.results.columnCount()):
                it = self.results.item(i, j)
                if it is not None:
                    it.setForeground(QBrush(QColor("#808080")))
                    it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self._filling = False

    def clear_results(self):
        self.results.clear()
        self.results.setRowCount(0)
        self.results.setColumnCount(0)
        self.result_cols, self.result_table = [], None
        self.stale = False

    def _result_edited(self, it):
        """a value changed in the Results pane: the row is changed in the table (a new row: when the cursor leaves it)"""
        if self._filling or self.result_table is None or self.stale:
            return
        r, c = it.row(), it.column()
        last = self.results.rowCount() - 1
        if r == last:                           # the new row (*)
            self.new_row = r
            return
        old = getattr(it, "_old", None)
        keys = self._row_keys(r, c, old)
        col = self.result_cols[c]
        sql = (f"UPDATE {QM._q(self.result_table)} SET {QM._q(col[0])} = {literal(it.text(), col[1])} WHERE "
               + " AND ".join(keys))
        try:
            dbinfo.ado_exec(self.cs, sql)
        except dbinfo.AdoError as e:
            self.say(str(e), error=True)

    def _row_keys(self, r, changed_col=None, old=None):
        names = [c[0].lower() for c in self.result_cols]
        out = []
        for k in self.result_keys:
            j = names.index(k.lower())
            v = self.results.item(r, j).text()
            if j == changed_col and old is not None:
                v = old
            out.append(f"{QM._q(self.result_cols[j][0])} = {literal(v, self.result_cols[j][1])}"
                       if v != NULL_TEXT else f"{QM._q(self.result_cols[j][0])} IS NULL")
        return out

    def _result_moved(self, row, col, prow, pcol):
        it = self.results.item(row, col)
        if it is not None:
            it._old = it.text()
        if self.new_row is not None and prow == self.new_row and row != prow:
            self.insert_new_row()

    def insert_new_row(self):
        r = self.new_row
        self.new_row = None
        if r is None or self.result_table is None:
            return
        cols, vals = [], []
        for j, c in enumerate(self.result_cols):
            t = self.results.item(r, j).text() if self.results.item(r, j) else ""
            if t:
                cols.append(QM._q(c[0]))
                vals.append(literal(t, c[1]))
        if not cols:
            return
        try:
            dbinfo.ado_exec(self.cs, f"INSERT INTO {QM._q(self.result_table)} ({', '.join(cols)}) "
                                     f"VALUES ({', '.join(vals)})")
        except dbinfo.AdoError as e:
            self.say(str(e), error=True)
            return
        self._filling = True
        self.results.setVerticalHeaderItem(r, QTableWidgetItem(""))
        n = self.results.rowCount()
        self.results.insertRow(n)
        for j in range(self.results.columnCount()):
            self.results.setItem(n, j, QTableWidgetItem(""))
        self.results.setVerticalHeaderItem(n, QTableWidgetItem("*"))
        self._filling = False

    def verify(self):
        """Query > Verify SQL Syntax"""
        sql = self.current_sql()
        try:
            if self.is_select(sql):
                dbinfo.ado_query(self.cs, sql, 1)
            elif not re.match(r"(?is)\s*(INSERT|UPDATE|DELETE|SELECT|CREATE|DROP|ALTER|TRANSFORM|PARAMETERS)\b", sql):
                raise dbinfo.AdoError("Syntax error in SQL statement.")
        except dbinfo.AdoError as e:
            self.say(str(e), error=True)
            return False
        self.say("The SQL syntax has been verified against the data source.")
        return True

    def set_group_by(self, on=None):
        self.model.set_group_by(not self.model.group_by if on is None else on)
        self.changed()

    def change_type(self, qtype):
        """Query > Change Type (VB6 12.3/q16-q26): Update, Delete and Insert Values keep one table (it is asked when
        there are more) and its rows; Delete keeps the rows of the criteria; Insert asks the table to append to
        ("Insert Into Table", the name in `quotes`), Make Table the new table ("Make Table")"""
        m = self.model
        if qtype == m.qtype:
            return
        if qtype in ONE_TABLE:
            if len(m.tables) > 1:
                title, text = ONE_TABLE[qtype]
                name = self.choose_table(title, text, [t.ref for t in m.tables])
                keep = m.table(name) if name else None
                if keep is None:
                    return
                for t in list(m.tables):
                    if t is not keep:
                        m.remove_table(t.ref)
            m.target = m.tables[0].name if m.tables else m.target
            m.group_by = False
            if qtype == "Delete":
                m.rows = [r for r in m.rows if r.criteria.strip() or any(o.strip() for o in r.ors)]
                for r in m.rows:
                    r.output = False
        elif qtype == "Insert":
            choices = []
            try:
                choices = [n for k, n in dbinfo.ado_objects(self.cs) if k == "T"]
            except dbinfo.AdoError:
                pass
            name = self.ask_item("Insert Into Table", "Append to\nTable name:", choices)
            if not name:
                return
            m.target = name if name[0] in "`[" else f"`{name}`"
        elif qtype == "Make Table":
            name = self.ask_text("Make Table", "Make\nTable name:", m.target)
            if not name:
                return
            m.target = name
        m.qtype = qtype
        self.changed()

    def ask_text(self, title, label, text):
        if getattr(self, "answers", None):
            return self.answers.pop(0)
        v, ok = QInputDialog.getText(self, title, label, text=text)
        return v.strip() if ok else ""

    def choose_table(self, title, text, items):
        """the table of an Update / Delete / Insert Values query: a list, OK, Cancel, Help (VB6 12.3/q17)"""
        if getattr(self, "answers", None):
            return self.answers.pop(0)
        d = QDialog(self)
        d.setWindowTitle(title)
        h = QHBoxLayout(d)
        v = QVBoxLayout()
        lab = QLabel(text)
        lab.setWordWrap(True)
        v.addWidget(lab)
        lst = QListWidget()
        lst.addItems(items)
        lst.setCurrentRow(0)
        lst.itemDoubleClicked.connect(lambda _it: d.accept())
        v.addWidget(lst)
        h.addLayout(v, 1)
        bb = QDialogButtonBox(Qt.Orientation.Vertical)
        bb.addButton(QDialogButtonBox.StandardButton.Ok)
        bb.addButton(QDialogButtonBox.StandardButton.Cancel)
        bb.addButton(QDialogButtonBox.StandardButton.Help)
        bb.accepted.connect(d.accept)
        bb.rejected.connect(d.reject)
        h.addWidget(bb, 0, Qt.AlignmentFlag.AlignTop)
        if not d.exec() or lst.currentItem() is None:
            return ""
        return lst.currentItem().text()

    def ask_item(self, title, label, items, editable=True):
        if getattr(self, "answers", None):
            return self.answers.pop(0)
        v, ok = QInputDialog.getItem(self, title, label, items, 0, editable)
        return v.strip() if ok else ""

    def grid_row(self):
        r = self.grid.currentRow()
        return self.model.rows[r] if 0 <= r < len(self.model.rows) else None

    def selected_rows(self):
        """the rows of the grid that are selected, or the columns selected in the diagram"""
        rows = sorted({i.row() for i in self.grid.selectedIndexes()})
        out = [self.model.rows[r] for r in rows if r < len(self.model.rows)]
        if out:
            return out
        for b in self.boxes:
            for it in b.list.selectedItems():
                c = it.data(Qt.ItemDataRole.UserRole)
                row = self.model.row_of(b.qt.ref, c)
                if row is None:
                    row = QM.Row(c, b.qt.ref, output=False)
                    self.model.rows.append(row)
                out.append(row)
        return out

    def add_to_output(self):
        for r in self.selected_rows():
            r.output = True
        self.changed()

    def sort(self, how):
        for r in self.selected_rows():
            self.model.set_sort(r, how)
        self.changed()

    def remove_filter(self):
        for r in self.selected_rows():
            r.criteria = ""
            r.ors = [""] * QM.ORS
        self.model.rows = [r for r in self.model.rows if r.output or r.sort or r.criteria or any(r.ors)]
        self.changed()

    def selected_join(self):
        for ln in self.lines:
            if ln.isSelected():
                return ln.join
        return self.model.joins[0] if len(self.model.joins) == 1 else None

    def all_rows_from(self, join, ref, on):
        join.set_all_from(ref, on)
        self.changed()

    def properties(self):
        d = QDialog(self)
        d.setWindowTitle("Properties")
        v = QVBoxLayout(d)
        tabs = QTabWidget()
        w = QWidget()
        f = QFormLayout(w)
        name = QLineEdit(self.model.name)
        f.addRow("Query name:", name)
        allc = QCheckBox("Output all columns")
        allc.setChecked(self.model.all_columns)
        dist = QCheckBox("DISTINCT values")
        dist.setChecked(self.model.distinct)
        f.addRow(QLabel("Options"))
        f.addRow(allc)
        f.addRow(dist)
        tabs.addTab(w, "Query")
        pw = QWidget()
        pv = QVBoxLayout(pw)
        params = QTableWidget(0, 2)
        params.setHorizontalHeaderLabels(["Parameter Name", "Data Type"])
        params.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        for i, p in enumerate(re.findall(r"\?", self.current_sql())):
            params.insertRow(i)
            params.setItem(i, 0, QTableWidgetItem(f"Param{i + 1}"))
            params.setItem(i, 1, QTableWidgetItem("Text"))
        pv.addWidget(params)
        tabs.addTab(pw, "Parameters")
        v.addWidget(tabs)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(d.accept)
        bb.rejected.connect(d.reject)
        v.addWidget(bb)
        self.last_dialog = d
        if getattr(self, "quiet", False) or not d.exec():
            return d
        self.model.name = name.text().strip()
        self.model.all_columns = allc.isChecked()
        self.model.distinct = dist.isChecked()
        self.changed()
        return d

    # ---------------------------------------------------------------- Diagram
    def new_annotation(self):
        a = Annotation()
        c = self.diagram.mapToScene(self.diagram.viewport().rect().center())
        a.setPos(c)
        self.diagram.scene().addItem(a)
        self.notes.append(a)
        return a

    def set_text_font(self):
        notes = [a for a in self.notes if a.isSelected()] or self.notes[-1:]
        if not notes:
            return
        f, ok = QFontDialog.getFont(notes[0].font(), self)
        if ok:
            for a in notes:
                a.setFont(f)

    def selected_boxes(self):
        return [b for b in self.boxes if b.isSelected()]

    def add_related(self):
        """Diagram > Add Related Tables: the tables the selected ones have relations with"""
        have = {t.name.lower() for t in self.model.tables}
        sel = [b.qt.name.lower() for b in (self.selected_boxes() or self.boxes)]
        add = []
        for pt, pc, ft, fc in self.relations():
            for a, b in ((pt, ft), (ft, pt)):
                if a.lower() in sel and b.lower() not in have and b not in add:
                    add.append(b)
        for n in add:
            self.model.add_table(n, self.columns_of(n), None, None, self.relations())
        if add:
            self.changed()
        return add

    def toggle_labels(self):
        self.show_labels = not self.show_labels
        self.update_joins()

    def arrange(self, selection=False):
        boxes = self.selected_boxes() if selection else self.boxes
        x0 = min((b.qt.x for b in boxes), default=20) if selection else 20
        y0 = min((b.qt.y for b in boxes), default=20) if selection else 20
        for i, b in enumerate(boxes):
            b.setPos(x0 + (i % 4) * (BOX_W + 50), y0 + (i // 4) * (BOX_H + 40))
        self.update_joins()

    def autosize(self):
        for b in self.selected_boxes() or self.boxes:
            n = min(len(b.qt.columns) + 1, 20)
            b.height = TITLE_H + 6 + n * max(16, b.list.sizeHintForRow(0))
            b.qt.height = b.height
            b.fill()
        self.update_joins()

    def set_zoom(self, z):
        if z == "Selection":
            boxes = self.selected_boxes() or self.boxes
            if boxes:
                r = boxes[0].sceneBoundingRect()
                for b in boxes[1:]:
                    r = r.united(b.sceneBoundingRect())
                self.diagram.fitInView(r.adjusted(-20, -20, 20, 20), Qt.AspectRatioMode.KeepAspectRatio)
                self.zoom = int(round(self.diagram.transform().m11() * 100))
            return
        self.zoom = z
        self.diagram.resetTransform()
        self.diagram.scale(z / 100.0, z / 100.0)

    def set_table_view(self, view):
        if view == "Custom" and getattr(self, "_custom_set", False) is False:
            self._custom_set = True
        self.table_view = view
        for b in self.boxes:
            b.fill()
        self.update_joins()

    def modify_custom_view(self):
        d = QDialog(self)
        d.setWindowTitle("Modify Custom View")
        v = QVBoxLayout(d)
        v.addWidget(QLabel("Columns:"))
        boxes = []
        for p in CUSTOM_PROPS:
            cb = QCheckBox(p)
            cb.setChecked(p in self.custom_props)
            cb.setEnabled(p != "Column Name")
            v.addWidget(cb)
            boxes.append(cb)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(d.accept)
        bb.rejected.connect(d.reject)
        v.addWidget(bb)
        if getattr(self, "quiet", False) or not d.exec():
            return d
        self.custom_props = [cb.text() for cb in boxes if cb.isChecked()]
        self.set_table_view("Custom")
        return d

    def page_breaks(self, on=None):
        self.diagram.page_breaks = (not self.diagram.page_breaks) if on is None else on
        self.diagram.viewport().update()

    def show_pane(self, key, on=None):
        w = self.panes[key]
        on = (not w.isVisible()) if on is None else on
        if not on and sum(1 for k in PANES if self.panes[k].isVisible()) <= 1:
            return
        w.setVisible(on)

    def pane_shown(self, key):
        return not self.panes[key].isHidden()

    # ---------------------------------------------------------------- the menus
    def pane_menu(self, key, gpos):
        """the right button of a pane (VB6 sq*)"""
        m = QMenu(self)
        if key == "diagram":
            m.addAction("&Run", self.run)
            m.addSeparator()
            a = m.addAction("&Group By", lambda: self.set_group_by())
            a.setCheckable(True)
            a.setChecked(self.model.group_by)
            m.addSeparator()
            names = {"Select": "Select Query", "Insert": "Insert Query", "Insert Values": "Insert Values Query",
                     "Update": "Update Query", "Delete": "Delete Query", "Make Table": "Make Table Query"}
            for t in QM.QUERY_TYPES:
                a = m.addAction(names[t], lambda tt=t: self.change_type(tt))
                a.setCheckable(True)
                a.setChecked(self.model.qtype == t)
        elif key == "sql":
            m.addAction("&Verify SQL Syntax", self.verify)
            m.addSeparator()
            m.addAction("Cu&t", self.sql.cut)
            m.addAction("&Copy", self.sql.copy)
            m.addAction("&Paste", self.sql.paste)
            m.addAction("Select &All", self.sql.selectAll)
            m.addSeparator()
            a = m.addAction("&Group By", lambda: self.set_group_by())
            a.setCheckable(True)
            a.setChecked(self.model.group_by)
            m.addSeparator()
            names = {"Select": "Select Query", "Insert": "Insert Query", "Insert Values": "Insert Values Query",
                     "Update": "Update Query", "Delete": "Delete Query", "Make Table": "Make Table Query"}
            for t in QM.QUERY_TYPES:
                a = m.addAction(names[t], lambda tt=t: self.change_type(tt))
                a.setCheckable(True)
                a.setChecked(self.model.qtype == t)
        elif key == "results":
            m.addAction("&Run", self.run)
            m.addAction("&Clear Results", self.clear_results)
            m.addSeparator()
            for text, fn in (("&First", lambda: self.go_row(0)), ("&Last", lambda: self.go_row(-1)),
                             ("&Next", lambda: self.go_row(self.results.currentRow() + 1)),
                             ("&Previous", lambda: self.go_row(self.results.currentRow() - 1)),
                             ("Ro&w...", self.ask_row), ("Ne&w", self.go_new)):
                m.addAction(text, fn).setEnabled(self.results.rowCount() > 0 and (text != "Ne&w" or
                                                                                    self.result_table is not None))
            m.addSeparator()
            m.addAction("Cu&t", lambda: self.copy_cells(True))
            m.addAction("&Copy", self.copy_cells)
            m.addAction("&Paste", self.paste_cell).setEnabled(self.result_table is not None)
        elif key == "grid":
            m.addAction("&Run", self.run)
            m.addSeparator()
            m.addAction("Cu&t", lambda: self.copy_grid(True))
            m.addAction("&Copy", self.copy_grid)
            m.addAction("&Delete", self.delete_grid_rows)
            m.addSeparator()
            a = m.addAction("&Group By", lambda: self.set_group_by())
            a.setCheckable(True)
            a.setChecked(self.model.group_by)
        m.addSeparator()
        m.addAction("&Hide Pane", lambda: self.show_pane(key, False))
        m.addSeparator()
        m.addAction("P&roperties", self.properties)
        self.last_menu = m
        if not getattr(self, "quiet", False):
            m.exec(gpos)
        return m

    def table_menu(self, box, gpos):
        m = QMenu(self)
        a = m.addAction("&Column Names", lambda: self._box_view(box, False))
        a.setCheckable(True)
        a.setChecked(not box.qt.collapsed)
        a = m.addAction("&Name Only", lambda: self._box_view(box, True))
        a.setCheckable(True)
        a.setChecked(box.qt.collapsed)
        m.addSeparator()
        m.addAction("&Select All Columns", lambda: self.toggle_column(box.qt.ref, "*", True))
        m.addAction("&Remove", lambda: self.remove_table(box.qt.ref))
        m.addSeparator()
        m.addAction("P&roperties", self.properties)
        self.last_menu = m
        if not getattr(self, "quiet", False):
            m.exec(gpos if not isinstance(gpos, QPointF) else gpos.toPoint())
        return m

    def _box_view(self, box, collapsed):
        box.qt.collapsed = collapsed
        box.fill()
        self.update_joins()

    def join_menu(self, line, gpos):
        j = line.join
        m = QMenu(self)
        for ref in (j.left, j.right):
            a = m.addAction(f"&Select All Rows from {ref}", lambda r=ref: self.all_rows_from(j, r, not j.all_from(r)))
            a.setCheckable(True)
            a.setChecked(j.all_from(ref))
        m.addSeparator()
        m.addAction("&Remove", lambda: (self.model.joins.remove(j), self.changed()))
        self.last_menu = m
        if not getattr(self, "quiet", False):
            m.exec(gpos.toPoint() if isinstance(gpos, QPointF) else gpos)
        return m

    # ---------------------------------------------------------------- the Results pane
    def go_row(self, r):
        n = self.results.rowCount()
        if not n:
            return
        r = n - 1 if r < 0 else min(r, n - 1)
        if self.result_table is not None and r == n - 1 and n > 1:
            r = n - 2
        self.results.setCurrentCell(max(0, r), max(0, self.results.currentColumn()))

    def ask_row(self):
        n = self.results.rowCount() - (1 if self.result_table is not None else 0)
        v, ok = QInputDialog.getInt(self, "Go To Row", "Row:", self.results.currentRow() + 1, 1, max(1, n))
        if ok:
            self.go_row(v - 1)

    def go_new(self):
        if self.result_table is not None:
            self.results.setCurrentCell(self.results.rowCount() - 1, 0)

    def copy_cells(self, cut=False):
        from PyQt6.QtWidgets import QApplication
        it = self.results.currentItem()
        if it is None:
            return
        QApplication.clipboard().setText(it.text())
        if cut and self.result_table is not None:
            it.setText("")

    def paste_cell(self):
        from PyQt6.QtWidgets import QApplication
        it = self.results.currentItem()
        if it is not None and self.result_table is not None:
            it.setText(QApplication.clipboard().text())

    def copy_grid(self, cut=False):
        from PyQt6.QtWidgets import QApplication
        it = self.grid.currentItem()
        if it is not None:
            QApplication.clipboard().setText(it.text())
            if cut:
                it.setText("")

    def delete_grid_rows(self):
        rows = sorted({i.row() for i in self.grid.selectedIndexes()}, reverse=True)
        for r in rows:
            if r < len(self.model.rows):
                del self.model.rows[r]
        self.changed()

    # ---------------------------------------------------------------- closing
    def save(self):
        if self.on_save is not None:
            self.on_save(self.current_sql())
        self.dirty = False

    def can_close(self):
        """the query of a command: its SQL goes back to the command - VB6 asks "Save changes to query 'Germans'?"
        (oracle 12.3/q9)"""
        if not self.dirty or self.on_save is None:
            return True
        ans = getattr(self, "close_answer", None)
        if ans is None:
            from . import APP_NAME
            ans = QMessageBox.question(self, APP_NAME, f"Save changes to query '{self.query_name}'?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No |
                                       QMessageBox.StandardButton.Cancel)
        if ans == QMessageBox.StandardButton.Cancel:
            return False
        if ans == QMessageBox.StandardButton.Yes:
            self.save()
        return True


def run_table(main, connect, table, conn_name="", base_dir=None):
    """Open of a table of the Data View: "Run Table: <name>", its rows in the Results pane"""
    d = QueryDesigner(main, connect, "", title=f"Run Table: {table}", panes=("results",), base_dir=base_dir,
                      conn_name=conn_name)
    d.model.add_table(table, d.columns_of(table), relations=())
    d.model.all_columns = True
    d.refresh()
    d.dirty = False
    return d


# ====================================================================== the Query and Diagram menus of the IDE
class IdeMenus:
    """Query and Diagram in the menu bar (between Run and Tools, as VB6 has them) and Table / Zoom / Show Panes in the
    View menu: their commands act on the Query Designer in front; without one they are grey (VB6)"""

    def __init__(self, mw, qmenu, dmenu, vmenu, vbefore=None):
        self.mw = mw
        self.q, self.d, self.v = {}, {}, {}

        def act(menu, store, key, text, fn, check=False):
            a = menu.addAction(text)
            a.triggered.connect(lambda _c=False, f=fn: self._do(f))
            a.setCheckable(check)
            store[key] = a
            return a
        m = qmenu
        act(m, self.q, "run", "&Run", lambda d: d.run())
        act(m, self.q, "clear", "&Clear Results", lambda d: d.clear_results())
        m.addSeparator()
        act(m, self.q, "verify", "&Verify SQL Syntax", lambda d: d.verify())
        m.addSeparator()
        act(m, self.q, "group", "&Group By", lambda d: d.set_group_by(), True)
        ct = m.addMenu("Change &Type")
        self.q["type"] = ct.menuAction()
        names = {"Select": "&Select", "Insert": "&Insert", "Insert Values": "Insert &Values", "Update": "&Update",
                 "Delete": "&Delete", "Make Table": "&Make Table"}          # (VB6 12.3/q16)
        for t in QM.QUERY_TYPES:
            act(ct, self.q, "type:" + t, names[t], lambda d, tt=t: d.change_type(tt), True)
        m.addSeparator()
        act(m, self.q, "output", "&Add To Output", lambda d: d.add_to_output())
        m.addSeparator()
        act(m, self.q, "asc", "Sort &Ascending", lambda d: d.sort("Ascending"))
        act(m, self.q, "desc", "Sort &Decending", lambda d: d.sort("Descending"))
        m.addSeparator()
        act(m, self.q, "filter", "Remove &Filter", lambda d: d.remove_filter())
        m.addSeparator()
        act(m, self.q, "allA", "&Select All Rows From <Table A>", lambda d: self._all_rows(d, "left"), True)
        act(m, self.q, "allB", "S&elect All Rows From <Table B>", lambda d: self._all_rows(d, "right"), True)
        qmenu.aboutToShow.connect(self.update)
        m = dmenu
        act(m, self.d, "note", "&New Text Annotation", lambda d: d.new_annotation())
        act(m, self.d, "font", "Set &Text Font", lambda d: d.set_text_font())
        m.addSeparator()
        act(m, self.d, "related", "Add &Related Tables", lambda d: d.add_related())
        m.addSeparator()
        act(m, self.d, "labels", "Show Relationship &Labels", lambda d: d.toggle_labels(), True)
        m.addSeparator()
        act(m, self.d, "custom", "&Modify Custom View", lambda d: d.modify_custom_view())
        m.addSeparator()
        act(m, self.d, "pages", "View &Page Breaks", lambda d: d.page_breaks(), True)
        act(m, self.d, "recalc", "Recalculate Page &Breaks", lambda d: d.diagram.viewport().update())
        m.addSeparator()
        act(m, self.d, "arrsel", "Arrange &Selection", lambda d: d.arrange(True))
        act(m, self.d, "arr", "&Arrange Tables", lambda d: d.arrange())
        m.addSeparator()
        act(m, self.d, "autosize", "Auto&size Selected Tables", lambda d: d.autosize())
        dmenu.aboutToShow.connect(self.update)
        # View > Table, Zoom, Show Panes (after Property Pages, VB6)
        def sub(text):
            mm = QMenu(text, vmenu)
            if vbefore is not None:
                vmenu.insertMenu(vbefore, mm)
            else:
                vmenu.addMenu(mm)
            return mm
        self.v_sep = vmenu.insertSeparator(vbefore) if vbefore is not None else vmenu.addSeparator()
        t = sub("T&able")
        self.v["table"] = t.menuAction()
        for v in TABLE_VIEWS:
            act(t, self.v, "table:" + v, v, lambda d, vv=v: d.set_table_view(vv), True)
        z = sub("Zoo&m")
        self.v["zoom"] = z.menuAction()
        for n in ZOOMS:
            act(z, self.v, f"zoom:{n}", f"{n}%", lambda d, nn=n: d.set_zoom(nn), True)
        act(z, self.v, "zoom:sel", "&Selection", lambda d: d.set_zoom("Selection"))
        sp = sub("Show Pa&nes")
        self.v["panes"] = sp.menuAction()
        for k, text in (("diagram", "&Diagram"), ("grid", "&Grid"), ("sql", "&SQL"), ("results", "&Results")):
            act(sp, self.v, "pane:" + k, text, lambda d, kk=k: d.show_pane(kk), True)
        vmenu.aboutToShow.connect(self.update)
        self.update()

    def _do(self, fn):
        d = self.mw.current_query()
        if d is not None:
            fn(d)

    @staticmethod
    def _all_rows(d, side):
        j = d.selected_join()
        if j is not None:
            ref = getattr(j, side)
            d.all_rows_from(j, ref, not j.all_from(ref))

    def update(self):
        d = self.mw.current_query()
        on = d is not None
        for a in list(self.q.values()) + list(self.d.values()):
            a.setEnabled(on)
            if not on and a.isCheckable():
                a.setChecked(False)
        for k in ("table", "zoom", "panes"):
            self.v[k].setVisible(on)
        self.v_sep.setVisible(on)
        self.q["allA"].setText("&Select All Rows From <Table A>")
        self.q["allB"].setText("S&elect All Rows From <Table B>")
        if not on:
            return
        m = d.model
        q = self.q
        q["clear"].setEnabled(d.results.rowCount() > 0)
        q["group"].setChecked(m.group_by)
        q["group"].setEnabled(m.parsed and m.qtype in ("Select", "Make Table", "Insert"))
        for t in QM.QUERY_TYPES:
            q["type:" + t].setChecked(m.qtype == t)
        rows = bool(d.grid_row() is not None or any(b.list.selectedItems() for b in d.boxes))
        for k in ("output", "asc", "desc", "filter"):
            q[k].setEnabled(rows)
        j = d.selected_join()
        for k, side in (("allA", "left"), ("allB", "right")):
            q[k].setEnabled(j is not None)
            if j is not None:
                ref = getattr(j, side)
                q[k].setText(("&Select" if side == "left" else "S&elect") + f" All Rows from {ref}")   # (12.3/q14)
                q[k].setChecked(j.all_from(ref))
        g = self.d
        g["font"].setEnabled(bool(d.notes))
        g["related"].setEnabled(bool(d.boxes))
        g["labels"].setChecked(d.show_labels)
        g["pages"].setChecked(d.diagram.page_breaks)
        g["recalc"].setEnabled(d.diagram.page_breaks)
        g["arrsel"].setEnabled(bool(d.selected_boxes()))
        g["arr"].setEnabled(bool(d.boxes))
        g["autosize"].setEnabled(bool(d.selected_boxes()))
        for v in TABLE_VIEWS:
            self.v["table:" + v].setChecked(d.table_view == v)
        for n in ZOOMS:
            self.v[f"zoom:{n}"].setChecked(d.zoom == n)
        for k in PANES:
            self.v["pane:" + k].setChecked(d.pane_shown(k))
