"""The Properties window of VB6: the object box, the Alphabetic and Categorized tabs, the grid of
properties with the editor of each kind (text, True/False and enumerations in a list, the color
palette, the Font dialog) and the name of the property at the bottom."""

import os

from PyQt6.QtCore import Qt, pyqtSignal, QPoint, QSize
from PyQt6.QtGui import QColor, QFont, QPainter, QIcon, QPixmap
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QComboBox, QTabBar, QTableWidget,
                             QTableWidgetItem, QHeaderView, QAbstractItemView, QLabel, QLineEdit, QToolButton,
                             QFrame, QTabWidget, QGridLayout, QListWidget, QListWidgetItem, QFontDialog,
                             QMessageBox, QColorDialog, QPushButton)

from . import vbprops as V
from . import dbinfo

# the properties with a data binding editor (vbprops: the note of the property)
DATA_NOTES = ("dbfile", "connect", "recordsource", "datasource", "datafield")
# the Connect list of the Data control (the installable ISAM formats of DAO 3.51)
CONNECTS = ["Access", "dBASE III;", "dBASE IV;", "dBASE 5.0;", "Excel 3.0;", "Excel 4.0;", "Excel 5.0;",
            "Excel 8.0;", "FoxPro 2.0;", "FoxPro 2.5;", "FoxPro 2.6;", "FoxPro 3.0;", "Lotus WK1;", "Lotus WK3;",
            "Lotus WK4;", "Paradox 3.X;", "Paradox 4.X;", "Paradox 5.X;", "Text;"]


CUSTOM = V.P("(Custom)", "custom", "", category="Misc")   # the property pages of an ActiveX control (phase 11)


def is_data_source(c):
    """a control the DataSource of a bound control can name: the Data control, an ADO Data control"""
    t = c.type_name
    return t in ("Data", "Adodc") or c.cls.lower().endswith(".adodc")

# the Palette tab of VB6 (RGB)
PALETTE = ["FFFFFF", "FFC0C0", "FFE0C0", "FFFFC0", "C0FFC0", "C0FFFF", "C0C0FF", "FFC0FF",
           "E0E0E0", "FF8080", "FFC080", "FFFF80", "80FF80", "80FFFF", "8080FF", "FF80FF",
           "C0C0C0", "FF0000", "FF8000", "FFFF00", "00FF00", "00FFFF", "0000FF", "FF00FF",
           "808080", "C00000", "C04000", "C0C000", "00C000", "00C0C0", "0000C0", "C000C0",
           "404040", "800000", "804000", "808000", "008000", "008080", "000080", "800080",
           "000000", "400000", "804040", "404000", "004000", "004040", "000040", "400040"]
SYS_RGB = ["C8C8C8", "000000", "99B4D1", "BFCDDB", "F0F0F0", "FFFFFF", "646464", "000000", "000000", "000000",
           "B4B4B4", "F4F7FC", "ABABAB", "3399FF", "FFFFFF", "F0F0F0", "A0A0A0", "6D6D6D", "000000", "434E54",
           "FFFFFF", "696969", "E3E3E3", "000000", "FFFFE1"]


def qcolor(v):
    """a VB6 color as QColor (system colors in their Windows 7 values)"""
    v = int(v) & 0xFFFFFFFF
    if v & 0x80000000:
        i = v & 0xFF
        return QColor("#" + SYS_RGB[i]) if i < len(SYS_RGB) else QColor("#F0F0F0")
    return QColor(v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF)


def swatch(v, size=12):
    pm = QPixmap(size, size)
    pm.fill(qcolor(v))
    p = QPainter(pm)
    p.setPen(QColor("#000000"))
    p.drawRect(0, 0, size - 1, size - 1)
    p.end()
    return QIcon(pm)


class ColorPopup(QFrame):
    """the drop-down of a color property: Palette and System tabs"""
    chosen = pyqtSignal(int)

    def __init__(self, parent, value):
        super().__init__(parent, Qt.WindowType.Popup)
        self.setFrameStyle(QFrame.Shape.Box)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        tabs = QTabWidget()
        lay.addWidget(tabs)
        pal = QWidget()
        g = QGridLayout(pal)
        g.setSpacing(2)
        for i, rgb in enumerate(PALETTE):
            b = QToolButton()
            b.setFixedSize(QSize(20, 20))
            b.setAutoRaise(False)
            b.setStyleSheet(f"background:#{rgb}; border:1px solid #808080;")
            v = int(rgb[4:6] + rgb[2:4] + rgb[0:2], 16)          # RGB -> &H00BBGGRR&
            b.clicked.connect(lambda _=False, v=v: self._pick(v))
            g.addWidget(b, i // 8, i % 8)
        more = QPushButton("Custom...")
        more.clicked.connect(lambda: self._custom(value))
        g.addWidget(more, 6, 0, 1, 8)
        tabs.addTab(pal, "Palette")
        self.sys = QListWidget()
        for i, name in enumerate(V.SYSTEM_COLORS):
            it = QListWidgetItem(swatch(0x80000000 | i), name)
            it.setData(Qt.ItemDataRole.UserRole, 0x80000000 | i)
            self.sys.addItem(it)
            if (int(value) & 0xFFFFFFFF) == (0x80000000 | i):
                self.sys.setCurrentItem(it)
        self.sys.itemClicked.connect(lambda it: self._pick(it.data(Qt.ItemDataRole.UserRole)))
        tabs.addTab(self.sys, "System")
        tabs.setCurrentIndex(1 if int(value) & 0x80000000 else 0)
        self.resize(220, 230)

    def _pick(self, v):
        self.chosen.emit(int(v))
        self.close()

    def _custom(self, value):
        c = QColorDialog.getColor(qcolor(value), self, "Color")
        if c.isValid():
            self._pick(c.red() | (c.green() << 8) | (c.blue() << 16))


class TextPopup(QFrame):
    """the drop-down editor of VB6 for lists and text of many lines"""
    done = pyqtSignal(str)

    def __init__(self, parent, text):
        super().__init__(parent, Qt.WindowType.Popup)
        self.setFrameStyle(QFrame.Shape.Box)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(1, 1, 1, 1)
        from PyQt6.QtWidgets import QPlainTextEdit
        self.edit = QPlainTextEdit(text)
        self.edit.installEventFilter(self)
        lay.addWidget(self.edit)

    def eventFilter(self, obj, e):
        from PyQt6.QtCore import QEvent
        if e.type() == QEvent.Type.KeyPress and e.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if e.modifiers() & Qt.KeyboardModifier.ControlModifier:
                self.edit.insertPlainText("\n")
            else:
                self.done.emit(self.edit.toPlainText())
                self.close()
            return True
        if e.type() == QEvent.Type.KeyPress and e.key() == Qt.Key.Key_Escape:
            self.close()
            return True
        return super().eventFilter(obj, e)


class PropertiesWindow(QWidget):
    """the window itself; main.designer_set_property() makes the changes"""

    def _error(self, text):
        """the errors of the values: in the Output window of the IDE, not in a message box"""
        name = getattr(self.ctl, "name", "") if getattr(self, "ctl", None) is not None else ""
        if hasattr(self.main, "message"):
            self.main.message(f"Properties{' (' + name + ')' if name else ''}: {text}", "error", show=True)
            self.main.statusBar().showMessage(text, 5000)
        else:
            QMessageBox.critical(self, "LinBASIC Studio", text)

    def __init__(self, main):
        super().__init__()
        self.main = main
        self.module = None
        self.designer = None
        self.ctl = None
        self.project = None                  # a project selected in the Project Explorer: (Name) only
        self.rows = []                       # P or a category name per row
        self.editor = None
        self.edit_row = -1
        self._busy = False
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(2)
        self.obj = QComboBox()
        self.obj.activated.connect(self._object_chosen)
        lay.addWidget(self.obj)
        self.tabs = QTabBar()
        self.tabs.addTab("Alphabetic")
        self.tabs.addTab("Categorized")
        self.tabs.currentChanged.connect(lambda _: self.fill_rows())
        lay.addWidget(self.tabs)
        self.grid = QTableWidget(0, 2)
        self.grid.horizontalHeader().setVisible(False)
        self.grid.verticalHeader().setVisible(False)
        self.grid.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.grid.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.grid.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.grid.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.grid.verticalHeader().setDefaultSectionSize(self.fontMetrics().height() + 6)
        self.grid.currentCellChanged.connect(lambda r, c, pr, pc: self._row_changed(r))
        self.grid.cellDoubleClicked.connect(self._double)
        lay.addWidget(self.grid, 1)
        self.desc = QLabel()
        self.desc.setWordWrap(True)
        self.desc.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.desc.setMinimumHeight(self.fontMetrics().height() * 3)
        self.desc.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        lay.addWidget(self.desc)
        self.last_prop = "Caption"

    # ------------------------------------------------------------ what is shown
    def show_object(self, module, designer, ctl):
        """the module (and the control of its designer) whose properties are shown"""
        self.module, self.designer = module, designer
        self.ctl = ctl
        self.project = None
        self._fill_objects()
        self.fill_rows()
        # VB6: the selected object ("Properties - Command1", DEV-8/me8-02), the form / module without one
        self._title(ctl.name if ctl is not None else module.name if module is not None else None)

    def show_project(self, q):
        """a project chosen in the Project Explorer: "Project1 Project" with (Name) (DEV-8/me5 picture 11)"""
        self.module = self.designer = self.ctl = None
        self.project = q
        self._fill_objects()
        self.fill_rows()
        self._title(q.name if q is not None else None)

    def _title(self, name):
        """VB6: "Properties - Project1", "Properties - Form1", "Properties - Command1"; "Properties" with nothing"""
        dock = getattr(self.main, "dock_props", None)
        if dock is not None:
            dock.setWindowTitle(f"Properties - {name}" if name else "Properties")

    def _fill_objects(self):
        self._busy = True
        self.obj.clear()
        m = self.module
        if self.project is not None:
            self.obj.addItem(f"{self.project.name} Project", None)
            self._busy = False
            return
        if m is None or m.sf is None:
            self._busy = False
            return
        if m.sf.root is not None:
            for c in m.sf.root.walk():
                label = c.name + (f"({c.index})" if c.index is not None else "") + " " + c.type_name
                self.obj.addItem(label, c)
                if c is self.ctl:
                    self.obj.setCurrentIndex(self.obj.count() - 1)
        else:
            self.obj.addItem(f"{m.name} {m.kind}", None)
        self._busy = False

    def _object_chosen(self, i):
        c = self.obj.itemData(i)
        self.ctl = c
        if self.designer is not None and c is not None and c.type_name != "Menu":
            self.designer.select_control(c) if c is not self.designer.root else self.designer.select([])
        self.fill_rows()

    def type_name(self):
        if self.ctl is not None:
            return self.ctl.type_name
        return None

    def fill_rows(self):
        self._drop_editor()
        self.grid.setRowCount(0)
        self.rows = []
        m = self.module
        if self.project is not None:             # a project: (Name) only
            self.rows = [V.P("(Name)", "name")]
        elif m is None or m.sf is None:
            self.desc.setText("")
            return
        elif self.ctl is None:                   # a module or class: (Name) and the class attributes
            self.rows = [V.P("(Name)", "name")]
            if m.kind == "Class":
                self.rows.append(V.P("Instancing", "enum", 1, f"instancing_{m.project.ptype.lower()}"))
            if m.sf.class_props:
                for p in m.sf.class_props:
                    if p.name.lower() != "multiuse":
                        self.rows.append(V.P(p.name, "readonly", p.raw_value))
        else:
            props = V.props_of(self.ctl.type_name)
            from .ocxpages import pages_of
            if pages_of(self.ctl.type_name):     # its property pages (phase 11): VB6's (Custom) row
                props = props[:1] + [CUSTOM] + props[1:] if props and props[0].name == "(Name)" else [CUSTOM] + props
            if self.tabs.currentIndex() == 1:
                rows = []
                for cat in V.CATEGORIES:
                    inside = sorted([p for p in props if p.category == cat], key=lambda p: (p.name != "(Name)", p.name.lower()))
                    if inside:
                        rows.append(cat)
                        rows += inside
                self.rows = rows
            else:
                self.rows = props
        self.grid.setRowCount(len(self.rows))
        bold = QFont(self.font())
        bold.setBold(True)
        for i, p in enumerate(self.rows):
            if isinstance(p, str):
                a = QTableWidgetItem(p)
                a.setFont(bold)
                a.setBackground(QColor("#E8E8E8"))
                b = QTableWidgetItem("")
                b.setBackground(QColor("#E8E8E8"))
                self.grid.setItem(i, 0, a)
                self.grid.setItem(i, 1, b)
                continue
            self.grid.setItem(i, 0, QTableWidgetItem(p.name))
            self._set_value_item(i, p)
        self.grid.resizeColumnToContents(0)
        # VB6 keeps the row of the property that was last chosen
        for i, p in enumerate(self.rows):
            if not isinstance(p, str) and p.name == self.last_prop:
                self.grid.setCurrentCell(i, 1)
                break
        else:
            if self.rows:
                self.grid.setCurrentCell(0, 1)

    def value(self, p):
        if p is CUSTOM:
            return ""
        if self.project is not None:
            return self.project.name
        if self.ctl is None:
            if p.name == "Instancing":
                return self.module.instancing()
            return self.module.name if p.kind == "name" else p.default
        form = self.designer.root if self.designer is not None else None
        if p.name == "Public" and self.module is not None and self.module.kind in ("UserControl", "PropertyPage"):
            return 0 if self.module.instancing() == 1 else 1           # VB_Exposed of the .ctl / .pag
        ok, sv = V._special_get(self.ctl, p)    # a DBGrid: its OleObjectBlob; DataSource / RowSource: the Bindings
        if ok:
            return sv
        if p.kind in ("picture", "list", "text", "string") and self.module is not None:
            v = self.module.frx_value(self.ctl, p.name)          # the data of the .frx
            if v is None:
                return [] if p.kind == "list" else (None if p.kind == "picture" else p.default)
            return v
        return V.value_of(self.ctl, p, form)

    def _set_value_item(self, i, p):
        v = self.value(p)
        it = QTableWidgetItem(V.display(p, v))
        if p.kind == "color" and v is not None:
            it.setIcon(swatch(v))
        if p.kind == "readonly" and self.ctl is None:
            it.setText(str(p.default))
        if p.kind == "readonly":
            it.setForeground(QColor("#808080"))
        self.grid.setItem(i, 1, it)

    def refresh_values(self):
        """after a change made elsewhere (the designer)"""
        if self.edit_row >= 0 and self.editor is not None and self.editor.hasFocus():
            return
        cur = self.grid.currentRow()
        for i, p in enumerate(self.rows):
            if not isinstance(p, str):
                self._set_value_item(i, p)
        if 0 <= cur < len(self.rows):
            self._row_changed(cur)

    # ------------------------------------------------------------ editing
    def _drop_editor(self):
        if self.editor is not None:
            ed = self.editor
            self.editor = None
            if 0 <= self.edit_row < self.grid.rowCount() and self.grid.cellWidget(self.edit_row, 1) is ed:
                self.grid.removeCellWidget(self.edit_row, 1)
            ed.hide()                             # removeCellWidget does not always delete it: no stray
            ed.deleteLater()                      # editors left on the grid
        self.edit_row = -1

    def _row_changed(self, r):
        self._drop_editor()
        if not 0 <= r < len(self.rows):
            return
        p = self.rows[r]
        if isinstance(p, str):
            self.desc.setText(f"<b>{p}</b>")
            return
        self.last_prop = p.name
        self.desc.setText(self._description(p))
        if self.ctl is None and p.kind != "name" and p.name != "Instancing":
            return
        ed = self._make_editor(p)
        if ed is not None:
            self.editor = ed
            self.edit_row = r
            self.grid.setCellWidget(r, 1, ed)

    def _description(self, p):
        extra = ""
        if p.kind == "enum":
            extra = ", ".join(f"{n} - {t}" for n, t in V.enum_items(p))
        elif p.kind == "bool":
            extra = "True / False"
        elif p.kind == "picture":
            extra = "... loads a picture (bmp, ico, cur, gif, jpg, wmf); Delete clears it. Saved in the .frx file."
        elif p.kind == "list":
            extra = "Ctrl+Enter starts a new line, Enter ends. Saved in the .frx file."
        elif p.kind == "text":
            extra = "With MultiLine = True the arrow opens the text: Ctrl+Enter starts a new line."
        elif p is CUSTOM:
            extra = "... opens the property pages of the control (Shift+F4)."
        return f"<b>{p.name}</b><br>{extra}"

    def _make_editor(self, p):
        v = self.value(p)
        if p.kind in ("bool", "enum", "shortcut"):
            cb = QComboBox()
            if p.kind == "bool":
                items = [("True", True), ("False", False)]
            elif p.kind == "enum":
                items = [(f"{n} - {t}", n) for n, t in V.enum_items(p)]
            else:
                items = [(t, code) for t, code in V.SHORTCUTS]
            for t, d in items:
                cb.addItem(t, d)
            cur = V.display(p, v)
            cb.setCurrentIndex(max(0, cb.findText(cur)))
            cb.activated.connect(lambda i, p=p, cb=cb: self.commit(p, cb.itemData(i)))
            return cb
        if p.kind == "color":
            return self._color_editor(p, v)
        if p.note in DATA_NOTES and self.ctl is not None:
            return self._data_editor(p, v)
        if p.kind == "cbbands" and self.ctl is not None:      # the bands of a CoolBar (coolbands.py)
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(0)
            lab = QLineEdit(V.display(p, v))
            lab.setReadOnly(True)
            lab.setFrame(False)
            h.addWidget(lab, 1)
            b = QToolButton()
            b.setText("...")
            b.clicked.connect(lambda _=False, p=p: self._coolbands_dialog(p))
            h.addWidget(b)
            return w
        if p.kind == "drepbinds" and self.ctl is not None:    # the RepeaterBindings of a DataRepeater (drepbinds.py)
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(0)
            lab = QLineEdit(V.display(p, v))
            lab.setReadOnly(True)
            lab.setFrame(False)
            h.addWidget(lab, 1)
            b = QToolButton()
            b.setText("...")
            b.clicked.connect(lambda _=False, p=p: self._drepbinds_dialog(p))
            h.addWidget(b)
            return w
        if p.kind == "drepname" and self.ctl is not None:     # the user controls a DataRepeater can repeat
            from .drepbinds import repeated_controls
            cb = QComboBox()
            cb.setEditable(True)
            cb.setFrame(False)
            cb.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
            cb.addItems(sorted(repeated_controls(self._db_dir())))
            cb.setEditText(V.display(p, v))
            cb.activated.connect(lambda i, p=p, cb=cb: self._text_done(p, cb.lineEdit()))
            cb.lineEdit().returnPressed.connect(lambda p=p, cb=cb: self._text_done(p, cb.lineEdit()))
            cb.lineEdit().editingFinished.connect(lambda p=p, cb=cb: self._text_done(p, cb.lineEdit()))
            return cb
        if p is CUSTOM and self.ctl is not None:              # (Custom): the property pages of the control
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(0)
            lab = QLineEdit("")
            lab.setReadOnly(True)
            lab.setFrame(False)
            h.addWidget(lab, 1)
            b = QToolButton()
            b.setText("...")
            b.clicked.connect(lambda _=False: self.main.property_pages())
            h.addWidget(b)
            return w
        if p.kind == "ppages" and self.ctl is not None:       # the PropertyPages of a user control (ppwizard.py)
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(0)
            lab = QLineEdit(V.display(p, v))
            lab.setReadOnly(True)
            lab.setFrame(False)
            h.addWidget(lab, 1)
            b = QToolButton()
            b.setText("...")
            b.clicked.connect(lambda _=False, p=p: self._ppages_dialog(p))
            h.addWidget(b)
            return w
        if p.kind == "dgcolumns" and self.ctl is not None:    # the columns of a DataGrid (dgcolumns.py)
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(0)
            lab = QLineEdit(V.display(p, v))
            lab.setReadOnly(True)
            lab.setFrame(False)
            h.addWidget(lab, 1)
            b = QToolButton()
            b.setText("...")
            b.clicked.connect(lambda _=False, p=p: self._dgcolumns_dialog(p))
            h.addWidget(b)
            return w
        if p.kind == "dataformat" and self.ctl is not None:   # the Format page (dfmtdialog.py)
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(0)
            lab = QLineEdit(V.display(p, v))
            lab.setReadOnly(True)
            lab.setFrame(False)
            h.addWidget(lab, 1)
            b = QToolButton()
            b.setText("...")
            b.clicked.connect(lambda _=False, p=p: self._dataformat_dialog(p))
            h.addWidget(b)
            return w
        if p.kind == "dbgcolumns":               # the columns of a DBGrid: a dialog (dbgcolumns.py)
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(0)
            lab = QLineEdit(V.display(p, v))
            lab.setReadOnly(True)
            lab.setFrame(False)
            h.addWidget(lab, 1)
            b = QToolButton()
            b.setText("...")
            b.clicked.connect(lambda _=False, p=p: self._columns_dialog(p))
            h.addWidget(b)
            return w
        if p.kind in ("font", "picture", "list", "readonly"):
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(0)
            lab = QLineEdit(V.display(p, v))
            lab.setReadOnly(True)
            lab.setFrame(False)
            h.addWidget(lab, 1)
            if p.kind != "readonly":
                b = QToolButton()
                b.setText("...")
                h.addWidget(b)
                if p.kind == "font":
                    b.clicked.connect(lambda _=False, p=p: self._font_dialog(p))
                elif p.kind == "picture":
                    b.clicked.connect(lambda _=False, p=p: self._picture_dialog(p))
                    lab.installEventFilter(self)
                    lab.setProperty("vbprop", p.name)
                else:
                    b.setText("")
                    b.setArrowType(Qt.ArrowType.DownArrow)
                    b.clicked.connect(lambda _=False, p=p, b=b: self._text_popup(p, b))
            return w
        if p.kind == "text" and (self.ctl.get("MultiLine") not in (None, 0, False) or "\n" in str(v)):
            w = QWidget()
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(0)
            le = QLineEdit(V.display(p, v))
            le.setFrame(False)
            le.setReadOnly("\n" in str(v))
            le.returnPressed.connect(lambda p=p, le=le: self._text_done(p, le))
            h.addWidget(le, 1)
            b = QToolButton()
            b.setArrowType(Qt.ArrowType.DownArrow)
            b.clicked.connect(lambda _=False, p=p, b=b: self._text_popup(p, b))
            h.addWidget(b)
            return w
        le = QLineEdit(V.display(p, v))
        le.setFrame(False)
        le.returnPressed.connect(lambda p=p, le=le: self._text_done(p, le))
        le.editingFinished.connect(lambda p=p, le=le: self._text_done(p, le))
        return le

    # ------------------------------------------------------------ the properties of data binding
    def _data_editor(self, p, v):
        """DatabaseName: a file ("..."); Connect, RecordSource, DataSource, DataField: a list to choose
        from or a text to type (the tables and queries of the database, the data sources of the form,
        the fields of the data source)"""
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        if p.note == "dbfile":
            le = QLineEdit(V.display(p, v))
            le.setFrame(False)
            le.returnPressed.connect(lambda p=p, le=le: self._text_done(p, le))
            le.editingFinished.connect(lambda p=p, le=le: self._text_done(p, le))
            h.addWidget(le, 1)
            b = QToolButton()
            b.setText("...")
            b.clicked.connect(lambda _=False, p=p: self._dbfile_dialog(p))
            h.addWidget(b)
            return w
        cb = QComboBox()
        cb.setEditable(True)
        cb.setFrame(False)
        cb.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        for t in self._data_items(p):
            cb.addItem(t)
        cb.setEditText(V.display(p, v))
        cb.activated.connect(lambda i, p=p, cb=cb: self._text_done(p, cb.lineEdit()))
        cb.lineEdit().returnPressed.connect(lambda p=p, cb=cb: self._text_done(p, cb.lineEdit()))
        cb.lineEdit().editingFinished.connect(lambda p=p, cb=cb: self._text_done(p, cb.lineEdit()))
        h.addWidget(cb, 1)
        return w

    def _db_dir(self):
        m = self.module
        path = getattr(m, "path", None) or getattr(getattr(m, "project", None), "path", None)
        return os.path.dirname(path) if path else None

    def _data_items(self, p):
        if p.note == "connect":
            return CONNECTS
        root = self.designer.root if self.designer is not None else None
        if p.note == "recordsource":
            return dbinfo.objects(dbinfo.resolve(str(self.ctl.get("DatabaseName", "") or ""), self._db_dir()))
        if root is None:
            return []
        if p.note == "datasource":
            names = []
            for c in root.walk():
                if c is not self.ctl and is_data_source(c) and c.name not in names:
                    names.append(c.name)
            return sorted(names, key=str.lower)
        if p.note == "datafield":
            src = str(self.ctl.get("DataSource", "") or "")
            ds = next((c for c in root.walk() if src and c.name.lower() == src.lower() and is_data_source(c)), None)
            if ds is None or ds.type_name != "Data":
                return []
            db = dbinfo.resolve(str(ds.get("DatabaseName", "") or ""), self._db_dir())
            return dbinfo.fields(db, str(ds.get("RecordSource", "") or ""))
        return []

    def _dbfile_dialog(self, p):
        from PyQt6.QtWidgets import QFileDialog
        path, _ = QFileDialog.getOpenFileName(
            self, "DatabaseName", self._db_dir() or "",
            "Access Databases (*.mdb *.MDB);;LinBASIC Databases (*.lbdb *.db *.sqlite);;All Files (*)")
        if path:
            self.commit(p, path)

    def _color_editor(self, p, v):
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        le = QLineEdit(V.display(p, v))
        le.setFrame(False)
        le.returnPressed.connect(lambda p=p, le=le: self._text_done(p, le))
        h.addWidget(le, 1)
        b = QToolButton()
        b.setArrowType(Qt.ArrowType.DownArrow)
        b.clicked.connect(lambda _=False, p=p, b=b: self._color_popup(p, b))
        h.addWidget(b)
        return w

    def _color_popup(self, p, button):
        pop = ColorPopup(self, self.value(p))
        pop.chosen.connect(lambda v, p=p: self.commit(p, v))
        pop.move(button.mapToGlobal(QPoint(button.width() - pop.width(), button.height())))
        pop.show()

    def eventFilter(self, obj, e):
        """Delete in a picture property: (None), as in VB6"""
        from PyQt6.QtCore import QEvent
        if e.type() == QEvent.Type.KeyPress and obj.property("vbprop") and e.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            p = V.prop_def(self.ctl.type_name, obj.property("vbprop"))
            if p is not None:
                self.commit(p, None)
            return True
        return super().eventFilter(obj, e)

    def _picture_dialog(self, p):
        from PyQt6.QtWidgets import QFileDialog
        exts = ["bmp", "dib", "gif", "jpg", "jpeg", "wmf", "emf", "ico", "cur"]
        pats = " ".join(f"*.{e} *.{e.upper()}" for e in exts)
        path, _ = QFileDialog.getOpenFileName(self, "Load Picture", "", f"All Picture Files ({pats});;All Files (*)")
        if path:
            with open(path, "rb") as f:
                self.commit(p, f.read())

    def _text_popup(self, p, button):
        """the drop-down text of List, ItemData and a MultiLine Text: Ctrl+Enter a new line, Enter ends"""
        v = self.value(p)
        lines = [str(x) for x in v] if p.kind == "list" else str(v or "").replace("\r\n", "\n").split("\n")
        pop = TextPopup(self, "\n".join(lines))
        pop.done.connect(lambda text, p=p: self._popup_done(p, text))
        w = max(180, self.grid.columnWidth(1))
        pop.resize(w, 140)
        pop.move(button.mapToGlobal(QPoint(button.width() - w, button.height())))
        pop.show()
        pop.edit.setFocus()

    def _popup_done(self, p, text):
        lines = text.split("\n")
        if p.kind == "list":
            while lines and lines[-1] == "":
                lines.pop()
            if p.name == "ItemData":
                try:
                    vals = [int(x.strip() or 0) for x in lines]
                    if any(not -2147483648 <= x <= 2147483647 for x in vals):
                        raise ValueError
                except ValueError:
                    self._error("Invalid property value")
                    return
                self.commit(p, vals)
            else:
                self.commit(p, lines)
        else:
            # VB6 ends the text of the drop-down with a line break (VB6 oracle d_frx: "A\r\nB\r\n")
            self.commit(p, "\r\n".join(lines) + "\r\n")

    def grid_fields(self):
        """the fields of the ADO Data control the DataGrid names in DataSource (Retrieve Fields), or []"""
        import re
        root = self.designer.root if self.designer is not None else None
        dsp = V.prop_def(self.ctl.type_name, "DataSource")
        src = str(self.value(dsp) or "") if dsp is not None else ""
        if root is None or not src:
            return []
        ds = next((c for c in root.walk() if c.name.lower() == src.lower() and c.type_name == "Adodc"), None)
        if ds is None:
            return []
        cs = str(ds.get("OLEDBString", "") or "") or str(ds.get("Connect", "") or "")
        source = str(ds.get("RecordSource", "") or "")
        if not cs or not source:
            return []
        m = re.search(r"(?i)data source\s*=\s*([^;]+)", cs)
        if m and not os.path.isabs(m.group(1).strip()):      # a database next to the project
            full = os.path.join(self._db_dir(), m.group(1).strip())
            if os.path.exists(full):
                cs = cs[:m.start(1)] + full + cs[m.end(1):]
        try:
            return [f[0] for f in dbinfo.ado_fields(cs, int(V._num(ds.get("CommandType", 8))), source)]
        except Exception:
            return []

    def _dgcolumns_dialog(self, p):
        """the columns of the design of a DataGrid; Retrieve Fields: the fields of its ADO Data control"""
        from .dgcolumns import GridColumnsDialog
        dlg = GridColumnsDialog(self, self.ctl.name, self.value(p), self.grid_fields())
        if dlg.exec():
            self.commit(p, dlg.chosen())

    def _drepbinds_dialog(self, p):
        """the RepeaterBindings of a DataRepeater: the bindable properties of its repeated control"""
        from .drepbinds import RepeaterBindingsDialog, repeated_controls, bindable_props
        name = V.drep_name(self.ctl)
        path = next((v for k, v in repeated_controls(self._db_dir()).items() if k.lower() == name.lower()), None)
        dlg = RepeaterBindingsDialog(self, self.ctl.name, self.value(p), bindable_props(path) if path else [])
        if dlg.exec():
            self.commit(p, dlg.chosen())

    def _coolbands_dialog(self, p):
        """the bands of a CoolBar; Child: one of the controls inside it"""
        from .coolbands import CoolBandsDialog
        names = [c.name if c.get("Index") is None else f"{c.name} ({int(V._num(c.get('Index')))})"
                 for kk, c in getattr(self.ctl, "order", []) if kk == "c"]
        dlg = CoolBandsDialog(self, self.ctl.name, self.value(p), names)
        if dlg.exec():
            self.commit(p, dlg.chosen())

    def _ppages_dialog(self, p, answer=None):
        """Connect Property Pages: the pages of the system and the PropertyPage modules of the project"""
        from .ppwizard import ConnectPagesDialog
        pages = [m.name for m in self.module.project.modules if m.kind == "PropertyPage"] \
            if self.module is not None else []
        dlg = ConnectPagesDialog(self, list(self.value(p) or []), pages)
        self._last_ppages_dialog = dlg
        if answer is not None:                  # the tests
            self.commit(p, answer(dlg))
            return dlg
        if dlg.exec():
            self.commit(p, dlg.chosen())
        return dlg

    def _dataformat_dialog(self, p):
        """the DataFormat of a bound control (BeginProperty DataFormat)"""
        from .dfmtdialog import DataFormatDialog
        dlg = DataFormatDialog(self, self.ctl.name, self.value(p))
        if dlg.exec():
            self.commit(p, dlg.chosen())

    def _columns_dialog(self, p):
        """the columns of a DBGrid; Retrieve Fields: the fields of the Data control of its DataSource"""
        from .dbgcolumns import ColumnsDialog
        cols, layout = self.value(p)
        fields = []
        root = self.designer.root if self.designer is not None else None
        dsp = V.prop_def(self.ctl.type_name, "DataSource")
        src = str(self.value(dsp) or "") if dsp is not None else ""
        if root is not None and src:
            ds = next((c for c in root.walk() if c.name.lower() == src.lower() and c.type_name == "Data"), None)
            if ds is not None:
                db = dbinfo.resolve(str(ds.get("DatabaseName", "") or ""), self._db_dir())
                fields = dbinfo.fields(db, str(ds.get("RecordSource", "") or ""))
        dlg = ColumnsDialog(self, self.ctl.name, cols, layout, fields)
        if dlg.exec():
            self.commit(p, dlg.chosen())

    def _font_dialog(self, p):
        f = self.value(p)
        qf = QFont(str(f["Name"]))
        qf.setPointSizeF(float(f["Size"]))
        qf.setBold(int(f["Weight"]) >= 600)
        qf.setItalic(bool(f["Italic"]))
        qf.setUnderline(bool(f["Underline"]))
        qf.setStrikeOut(bool(f["Strikethrough"]))
        nf, ok = QFontDialog.getFont(qf, self, "Font")
        if ok:
            self.commit(p, {"Name": nf.family(), "Size": nf.pointSizeF(), "Charset": None,
                            "Weight": 700 if nf.bold() else 400, "Underline": nf.underline(),
                            "Italic": nf.italic(), "Strikethrough": nf.strikeOut()})

    def _text_done(self, p, le):
        if self.editor is None or self._busy:
            return
        text = le.text()
        if text == V.display(p, self.value(p)):
            return
        try:
            v = V.parse_input(p, text)
        except (ValueError, OverflowError):
            self._error("Invalid property value")
            le.setText(V.display(p, self.value(p)))
            return
        self.commit(p, v)

    def _double(self, r, c):
        """double click on a True/False or enumerated property: the next value (as in VB6)"""
        if not 0 <= r < len(self.rows) or isinstance(self.rows[r], str):
            return
        p = self.rows[r]
        v = self.value(p)
        if p.kind == "bool":
            self.commit(p, not v)
        elif p.kind == "enum":
            ks = [n for n, _ in V.enum_items(p)]
            if v in ks:
                self.commit(p, ks[(ks.index(v) + 1) % len(ks)])
        elif p.kind == "font":
            self._font_dialog(p)
        elif p is CUSTOM:
            self.main.property_pages()

    def commit(self, p, value):
        if self.project is not None:
            q = self.project
            err = self.main.rename_project(q, value)
            if err:
                self._error(err)
            self.show_project(q)
            return
        if self.ctl is None:
            if p.kind == "name":
                self.main.rename_module(self.module, value)
            elif p.name == "Instancing":
                self.module.set_instancing(int(value))
                self.main.update_tree_marks()
                self.refresh_values()
            return
        if p.name == "Public" and self.module is not None and self.module.kind in ("UserControl", "PropertyPage"):
            if value and self.module.project.ptype != "Control":
                self._error("Public user controls and property pages are allowed in ActiveX Control projects only")
                self.refresh_values()
                return
            self.module.set_instancing(5 if value else 1)
            self.main.update_tree_marks()
            self.main.refresh_toolbox()
            self.refresh_values()
            return
        if self.designer is None:
            return
        self._busy = True
        try:
            err = self.designer.set_property(self.ctl, p, value)
        finally:
            self._busy = False
        if err:
            self._error(err)
        self.refresh_values()
