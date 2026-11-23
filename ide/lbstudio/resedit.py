# LinBASIC Studio - the VB 6 Resource Editor (Add-Ins > Add-In Manager > VB 6 Resource Editor)
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""The Resource Editor of VB6 (phase 8.6, VB-TESTS/DEV-8/rs1), on resmodel.ResFile:

  - Add-In Manager: "VB 6 Resource Editor" with Loaded/Unloaded, Load on Startup, Command Line; loaded, it
    gives Tools > Resource Editor, Project > Add New Resource File and a button of the Standard toolbar;
  - the window "VB Resource Editor - <file>": New Resource File, Open Resource File, Save Resource File, Edit
    Properties | Duplicate, Undo, Delete | Edit String Tables, Add Cursor, Add Icon, Add Bitmap, Add Custom
    Resource | Show Help; a tree (the root "Resource File" or the path; folders alphabetically), a status bar
    "(Modified)"; the right button on empty space: New, Open | Save, Save As...; closing it only hides it;
  - "Edit String Tables": Insert New Row, Insert New String Table, Properties | Renumber Rows, Cut, Copy,
    Paste, Undo, Export Text To Clipboard | Delete Row, Delete String Table | Help; columns Id and one per
    language; Enter on the last row adds the next id; Renumber Rows: "Renumber Resource Id's" (Renumber Id's
    from the Base Value, Rebase Id's: + Base Value);
  - Edit Properties: Type (custom resources), Id, Language; bitmaps, icons and cursors with a Preview and
    View Image;
  - a changed resource file and a project that closes: "The resource file in <project> has changed.  Click
    Yes to save the changes, or No to discard." (Yes / No).
"""
import os
import struct

from PyQt6.QtCore import Qt, pyqtSignal, QRectF
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QPen, QFont, QPolygonF
from PyQt6.QtCore import QPointF
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QToolBar, QTreeWidget, QTreeWidgetItem, QLabel,
                             QMenu, QFileDialog, QMessageBox, QDialog, QFormLayout, QLineEdit, QComboBox,
                             QDialogButtonBox, QPushButton, QTableWidget, QTableWidgetItem, QGroupBox,
                             QRadioButton, QCheckBox, QPlainTextEdit, QAbstractItemView,
                             QApplication, QGridLayout, QScrollArea)

from . import resmodel as R
from . import lcid as L

TITLE = "VB 6 Resource Editor"
ADDIN = "VB 6 Resource Editor"
PDW = "Package and Deployment Wizard"       # phase 9.5 (DEV-9/pd1-03: its Description is its name)
APPWIZ = "VB 6 Application Wizard"         # phase 9.2 (DEV-9/aw18: Add-Ins > Application Wizard..., Toolbar Wizard...)
DFWIZ = "VB 6 Data Form Wizard"            # phase 9.2b (DEV-9/df1-01: Add-Ins > Data Form Wizard...)
CBUILDER = "VB 6 Class Builder Utility"   # phase 9.2c (DEV-9/cb1-02: Add-Ins > Class Builder Utility...)
AXWIZ = "VB 6 ActiveX Ctrl Interface Wizard"   # phase 9.2d (DEV-9/ax1: Add-Ins > ActiveX Control Interface Wizard...)
PPWIZ = "VB 6 Property Page Wizard"    # phase 9.2e (DEV-9/pp1: Add-Ins > Property Page Wizard...)
DOWIZ = "VB 6 Data Object Wizard"      # phase 9.2f (DEV-9/do1: Add-Ins > VB 6 Data Object Wizard)
ADDINS = [("DTC Framework Registrar", ""), (PDW, PDW), ("Source Code Control", ""),
          (AXWIZ, "Makes the properties, methods and events of a user control and the code under them "
                  "(ActiveX Control Interface Wizard)"), ("VB 6 ActiveX Doc Migration Wizard", ""),
          ("VB 6 Add-In Toolbar", ""), ("VB 6 API Viewer", ""),
          (APPWIZ, "Makes the start of a new application: its main form, menus, toolbar and standard forms "
                   "(Application Wizard), and toolbars of forms (Toolbar Wizard)"),
          (CBUILDER, "Makes the classes and collections of a project and their properties, methods, events and enums "
                     "(Class Builder Utility)"), (DFWIZ, "Makes forms whose controls show and change the records of a database (Data Form "
                                                            "Wizard)"), (DOWIZ, "Makes classes that are a source of data from the commands of a Data Environment, and user "
                                                            "controls bound to them (Data Object Wizard)"),
          (PPWIZ, "Makes the property pages of a user control (Property Page Wizard)"), (ADDIN, "Visual Basic 6 Resource Editor"),
          ("VB 6 Template Manager", ""), ("VB 6 Wizard Manager", "")]


def short_list(combo):
    """a Language list opens as in VB6, 8 rows with a scroll bar, not over the whole screen"""
    combo.setMaxVisibleItems(8)
    combo.setStyleSheet("QComboBox { combobox-popup: 0; }")


def res_create_msg(path):
    """Project > Add New Resource File with a name that is not there (DEV-8/rs1 stage 6)"""
    return f"{os.path.basename(path)}\nThis file doesn't exist.\n\nCreate the file?"


def changed_msg(project_name):
    return (f"The resource file in {project_name} has changed.  Click Yes to save the changes, "
            "or No to discard.")


# ---------------------------------------------------------------- icons of the buttons
def _icon(name):
    pm = QPixmap(16, 16)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    navy, black = QColor(0, 0, 128), QColor(0, 0, 0)
    if name == "new":
        p.setPen(black)
        p.setBrush(QColor(255, 255, 255))
        p.drawPolygon(QPolygonF([QPointF(3, 1.5), QPointF(10, 1.5), QPointF(13, 4.5), QPointF(13, 14.5),
                                 QPointF(3, 14.5)]))
    elif name == "open":
        p.setPen(black)
        p.setBrush(QColor(240, 200, 80))
        p.drawRect(QRectF(1.5, 5.5, 13, 8))
        p.drawRect(QRectF(1.5, 3.5, 5, 2))
    elif name == "save":
        p.setPen(black)
        p.setBrush(QColor(60, 60, 140))
        p.drawRect(QRectF(1.5, 1.5, 13, 13))
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(4.5, 1.5, 7, 5))
    elif name == "props":
        p.setPen(black)
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(1.5, 2.5, 12, 11))
        for y in (5, 8, 11):
            p.drawLine(4, y, 11, y)
    elif name == "duplicate":
        p.setPen(black)
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(1.5, 1.5, 8, 9))
        p.drawRect(QRectF(5.5, 5.5, 8, 9))
    elif name == "undo":
        p.setPen(QPen(navy, 2))
        p.drawArc(QRectF(3, 4, 10, 9), 30 * 16, 220 * 16)
        p.drawLine(3, 4, 4, 9)
        p.drawLine(3, 9, 8, 8)
    elif name == "delete":
        p.setPen(QPen(black, 2))
        p.drawLine(3, 3, 13, 13)
        p.drawLine(13, 3, 3, 13)
    elif name == "help":
        f = QFont("Sans")
        f.setBold(True)
        f.setPixelSize(14)
        p.setFont(f)
        p.setPen(navy)
        p.drawText(QRectF(0, 0, 16, 16), int(Qt.AlignmentFlag.AlignCenter), "?")
    elif name in ("strings", "string"):
        f = QFont("Sans")
        f.setPixelSize(9)
        f.setBold(True)
        p.setFont(f)
        p.setPen(black)
        p.drawText(QRectF(0, 1, 16, 14), int(Qt.AlignmentFlag.AlignCenter), "abc")
        p.drawLine(1, 14, 15, 14)
    elif name == "cursor":
        p.setPen(black)
        p.setBrush(QColor(255, 255, 255))
        p.drawPolygon(QPolygonF([QPointF(4, 1), QPointF(4, 13), QPointF(7, 10), QPointF(9, 15), QPointF(11, 14),
                                 QPointF(9, 9), QPointF(13, 9)]))
    elif name == "icon":
        p.setPen(black)
        for i in range(4):
            for j in range(4):
                p.setBrush(QColor(255, 255, 255) if (i + j) % 2 else QColor(80, 80, 80))
                p.drawRect(QRectF(1.5 + 3 * i, 1.5 + 3 * j, 3, 3))
    elif name == "bitmap":
        p.setPen(black)
        p.setBrush(QColor(110, 190, 240))
        p.drawRect(QRectF(1.5, 2.5, 13, 11))
        p.setBrush(QColor(40, 150, 60))
        p.drawPolygon(QPolygonF([QPointF(2, 13), QPointF(6, 7), QPointF(9, 11), QPointF(11, 9), QPointF(14, 13)]))
    elif name == "custom":
        p.setPen(black)
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(1.5, 1.5, 13, 13))
        p.drawLine(1, 6, 15, 6)
        p.drawLine(6, 1, 6, 15)
    elif name == "folder":
        p.setPen(QColor(150, 120, 20))
        p.setBrush(QColor(250, 220, 100))
        p.drawRect(QRectF(1.5, 4.5, 13, 9))
        p.drawRect(QRectF(1.5, 2.5, 5, 2))
    elif name == "row":
        p.setPen(black)
        p.drawRect(QRectF(1.5, 4.5, 13, 7))
        p.setPen(QPen(QColor(0, 128, 0), 2))
        p.drawLine(12, 1, 12, 5)
        p.drawLine(10, 3, 14, 3)
    elif name == "column":
        p.setPen(black)
        p.drawRect(QRectF(4.5, 1.5, 7, 13))
        p.setPen(QPen(QColor(0, 128, 0), 2))
        p.drawLine(13, 1, 13, 5)
        p.drawLine(11, 3, 15, 3)
    elif name == "renumber":
        f = QFont("Sans")
        f.setPixelSize(8)
        p.setFont(f)
        p.setPen(black)
        for k, y in enumerate((0, 5, 10)):
            p.drawText(QRectF(0, y, 16, 6), int(Qt.AlignmentFlag.AlignLeft), str(k + 1))
            p.drawLine(6, y + 3, 15, y + 3)
    elif name in ("cut", "copy", "paste", "export"):
        p.setPen(black)
        p.setBrush(QColor(255, 255, 255))
        if name == "cut":
            p.drawEllipse(QRectF(2, 10, 4, 4))
            p.drawEllipse(QRectF(10, 10, 4, 4))
            p.drawLine(5, 10, 11, 1)
            p.drawLine(11, 10, 5, 1)
        elif name == "copy":
            p.drawRect(QRectF(1.5, 1.5, 8, 9))
            p.drawRect(QRectF(6.5, 5.5, 8, 9))
        elif name == "paste":
            p.setBrush(QColor(200, 160, 90))
            p.drawRect(QRectF(2.5, 2.5, 10, 12))
            p.setBrush(QColor(255, 255, 255))
            p.drawRect(QRectF(6.5, 6.5, 8, 8))
        else:
            p.setBrush(QColor(200, 160, 90))
            p.drawRect(QRectF(2.5, 2.5, 11, 12))
            p.drawLine(5, 7, 11, 7)
            p.drawLine(5, 10, 11, 10)
    elif name in ("delrow", "deltable"):
        p.setPen(black)
        if name == "delrow":
            p.drawRect(QRectF(1.5, 4.5, 13, 7))
        else:
            p.drawRect(QRectF(4.5, 1.5, 7, 13))
        p.setPen(QPen(QColor(200, 0, 0), 2))
        p.drawLine(9, 9, 15, 15)
        p.drawLine(15, 9, 9, 15)
    p.end()
    return QIcon(pm)


# ---------------------------------------------------------------- pictures of the resources
def bitmap_file(dib):
    """a .bmp file from the data of a RT_BITMAP (the 14 bytes before it)"""
    if len(dib) < 40:
        return b""
    hsz, = struct.unpack_from("<I", dib, 0)
    bits, comp, used = struct.unpack_from("<H", dib, 14)[0], struct.unpack_from("<I", dib, 16)[0], \
        struct.unpack_from("<I", dib, 32)[0]
    colors = used if used else ((1 << bits) if bits <= 8 else 0)
    off = 14 + hsz + colors * 4 + (12 if comp == 3 and hsz == 40 else 0)
    return b"BM" + struct.pack("<IHHI", 14 + len(dib), 0, 0, off) + dib


def icon_file(it):
    """an .ico (or .cur) file of an icon / a cursor of the resource file, for its preview"""
    cur = it.kind == "cursor"
    imgs = [(e, d[4:] if cur else d) for e, d, _ in it.images]
    out = struct.pack("<HHH", 0, 2 if cur else 1, len(imgs))
    off = 6 + 16 * len(imgs)
    body = b""
    for e, d in imgs:
        w, h = struct.unpack_from("<ii", d, 4) if len(d) >= 12 else (32, 64)
        out += struct.pack("<BBBBHHII", w & 0xFF, (h // 2) & 0xFF, e.get("colors", 0) & 0xFF, 0, 1,
                           struct.unpack_from("<H", d, 14)[0] if len(d) >= 16 else 0, len(d), off + len(body))
        body += d
    return out + body


def pixmap_of(it):
    pm = QPixmap()
    if it.kind == "bitmap":
        pm.loadFromData(bitmap_file(it.data), "BMP")
    elif it.kind in ("icon", "cursor"):
        pm.loadFromData(icon_file(it), "ICO")
    return pm


# ---------------------------------------------------------------- dialogs
class AddInManagerDialog(QDialog):
    """Add-Ins > Add-In Manager... (DEV-8/rs1 pictures 1, 2): the add-ins and their Load Behavior"""

    def __init__(self, main):
        super().__init__(main)
        self.setWindowTitle("Add-In Manager")
        self.main = main
        S = main.S
        self.state = {ADDIN: [bool(main.res_loaded), bool(S["addin_reseditor_startup"]), False],
                      PDW: [bool(main.pdw_loaded), bool(S["addin_pdw_startup"]), False],
                      APPWIZ: [bool(main.appwiz_loaded), bool(S["addin_appwiz_startup"]), False],
                      DFWIZ: [bool(main.dfwiz_loaded), bool(S["addin_dfwiz_startup"]), False],
                      CBUILDER: [bool(main.cbuilder_loaded), bool(S["addin_cbuilder_startup"]), False],
                      AXWIZ: [bool(main.axwiz_loaded), bool(S["addin_axwiz_startup"]), False],
                      PPWIZ: [bool(main.ppwiz_loaded), bool(S["addin_ppwiz_startup"]), False],
                      DOWIZ: [bool(main.dowiz_loaded), bool(S["addin_dowiz_startup"]), False]}
        lay = QGridLayout(self)
        self.list = QTableWidget(len(ADDINS), 2)
        self.list.setHorizontalHeaderLabels(["Available Add-Ins", "Load Behavior"])
        self.list.verticalHeader().hide()
        self.list.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.list.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.list.setShowGrid(False)
        self.list.horizontalHeader().setStretchLastSection(True)
        for i, (n, _) in enumerate(ADDINS):
            self.list.setItem(i, 0, QTableWidgetItem(n))
            self.list.setItem(i, 1, QTableWidgetItem(""))
        self.list.resizeColumnToContents(0)
        lay.addWidget(self.list, 0, 0)
        btns = QVBoxLayout()
        for text, slot in (("OK", self._ok), ("Cancel", self.reject)):
            b = QPushButton(text)
            b.clicked.connect(slot)
            btns.addWidget(b)
        btns.addStretch(1)
        b = QPushButton("Help")
        btns.addWidget(b)
        lay.addLayout(btns, 0, 1)
        bottom = QHBoxLayout()
        dg = QGroupBox("Description")
        dl = QVBoxLayout(dg)
        self.desc = QPlainTextEdit()
        self.desc.setReadOnly(True)
        dl.addWidget(self.desc)
        bottom.addWidget(dg, 1)
        lg = QGroupBox("Load Behavior")
        ll = QVBoxLayout(lg)
        self.cb_loaded = QCheckBox("Loaded/Unloaded")
        self.cb_startup = QCheckBox("Load on Startup")
        self.cb_cmd = QCheckBox("Command Line")
        for c in (self.cb_loaded, self.cb_startup, self.cb_cmd):
            ll.addWidget(c)
            c.toggled.connect(self._changed)
        bottom.addWidget(lg)
        lay.addLayout(bottom, 1, 0, 1, 2)
        self.list.currentCellChanged.connect(lambda r, c, pr, pc: self._select(r))
        self._filling = False
        self._fill_behavior()
        self.list.setCurrentCell(0, 0)
        self.resize(560, 420)

    def _name(self, row):
        return ADDINS[row][0] if 0 <= row < len(ADDINS) else None

    def _fill_behavior(self):
        for i, (n, _) in enumerate(ADDINS):
            st = self.state.get(n)
            text = ""
            if st and st[0]:
                text = "Startup / Loaded" if st[1] else "Loaded"
            elif st and st[1]:
                text = "Startup"
            self.list.item(i, 1).setText(text)

    def _select(self, row):
        n = self._name(row)
        self._filling = True
        self.desc.setPlainText(dict(ADDINS).get(n, "") if n else "")
        st = self.state.get(n, [False, False, False])
        on = n in self.state                    # LinBASIC Studio: the Resource Editor, the Package Wizard, the App Wizard
        for c, v in zip((self.cb_loaded, self.cb_startup, self.cb_cmd), st):
            c.setChecked(v)
            c.setEnabled(on)
        self._filling = False

    def _changed(self):
        if self._filling:
            return
        n = self._name(self.list.currentRow())
        if n in self.state:
            self.state[n] = [self.cb_loaded.isChecked(), self.cb_startup.isChecked(), self.cb_cmd.isChecked()]
            self._fill_behavior()

    def _ok(self):
        loaded, startup, _ = self.state[ADDIN]
        self.main.res_loaded = loaded           # Loaded/Unloaded: now; Load on Startup: the next times
        self.main.S["addin_reseditor_startup"] = startup
        loaded, startup, _ = self.state[PDW]
        self.main.pdw_loaded = loaded
        self.main.S["addin_pdw_startup"] = startup
        loaded, startup, _ = self.state[APPWIZ]
        self.main.appwiz_loaded = loaded
        self.main.S["addin_appwiz_startup"] = startup
        loaded, startup, _ = self.state[DFWIZ]
        self.main.dfwiz_loaded = loaded
        self.main.S["addin_dfwiz_startup"] = startup
        loaded, startup, _ = self.state[CBUILDER]
        self.main.cbuilder_loaded = loaded
        self.main.S["addin_cbuilder_startup"] = startup
        loaded, startup, _ = self.state[AXWIZ]
        self.main.axwiz_loaded = loaded
        self.main.S["addin_axwiz_startup"] = startup
        loaded, startup, _ = self.state[PPWIZ]
        self.main.ppwiz_loaded = loaded
        self.main.S["addin_ppwiz_startup"] = startup
        loaded, startup, _ = self.state[DOWIZ]
        self.main.dowiz_loaded = loaded
        self.main.S["addin_dowiz_startup"] = startup
        self.main.S.save()
        self.accept()


class PropertiesDialog(QDialog):
    """Edit Properties (DEV-8/rs1 pictures 21, 23): Type (custom resources), Id, Language; a picture has a
    Preview and View Image"""

    def __init__(self, parent, it, string_lang=None):
        super().__init__(parent)
        self.it = it
        kind = it.kind if it is not None else "string"
        self.setWindowTitle({"icon": "Edit Icon Properties", "cursor": "Edit Cursor Properties",
                             "bitmap": "Edit Bitmap Properties"}.get(kind, "Edit Properties"))
        outer = QHBoxLayout(self)
        form = QFormLayout()
        self.type_e = self.id_e = None
        if kind == "custom":
            self.type_e = QLineEdit(R.show_name(it.rtype))
            form.addRow("Type:", self.type_e)
        if kind != "string":
            self.id_e = QLineEdit(R.show_name(it.name))
            self.id_e.selectAll()
            form.addRow("Id:", self.id_e)
        self.lang = QComboBox()
        short_list(self.lang)
        cur = string_lang if kind == "string" else it.lang
        for i, n in L.combo_items():
            self.lang.addItem(n, i)
        k = self.lang.findData(cur)
        if k < 0:
            self.lang.insertItem(0, L.name(cur), cur)
            k = 0
        self.lang.setCurrentIndex(k)
        form.addRow("Language:", self.lang)
        left = QVBoxLayout()
        left.addLayout(form)
        left.addStretch(1)
        bb = QHBoxLayout()
        bb.addStretch(1)
        ok, cancel = QPushButton("OK"), QPushButton("Cancel")
        ok.setDefault(True)
        ok.clicked.connect(self._ok)
        cancel.clicked.connect(self.reject)
        bb.addWidget(ok)
        bb.addWidget(cancel)
        if kind in ("bitmap", "icon", "cursor"):
            view = QPushButton("View Image")
            view.clicked.connect(lambda: ImageViewer(self, it).exec())
            bb.addWidget(view)
        left.addLayout(bb)
        outer.addLayout(left, 1)
        if kind in ("bitmap", "icon", "cursor"):
            pv = QVBoxLayout()
            pv.addWidget(QLabel("Preview:"))
            pic = QLabel()
            pic.setFixedSize(120, 110)
            pic.setAlignment(Qt.AlignmentFlag.AlignCenter)
            pic.setStyleSheet("border: 1px solid gray;")
            pm = pixmap_of(it)
            if not pm.isNull():
                pic.setPixmap(pm.scaled(116, 106, Qt.AspectRatioMode.KeepAspectRatio))
            pv.addWidget(pic)
            pv.addStretch(1)
            outer.addLayout(pv)
        self.result_values = None

    def _ok(self):
        vals = {"lang": self.lang.currentData()}
        if self.type_e is not None:
            t = R.norm_name(self.type_e.text())
            if t in ("", None):
                return
            vals["type"] = t
        if self.id_e is not None:
            n = R.norm_name(self.id_e.text())
            if n in ("", None):
                return
            vals["name"] = n
        self.result_values = vals
        self.accept()


class ImageViewer(QDialog):
    def __init__(self, parent, it):
        super().__init__(parent)
        self.setWindowTitle(R.show_name(it.name))
        lay = QVBoxLayout(self)
        sa = QScrollArea()
        lab = QLabel()
        lab.setPixmap(pixmap_of(it))
        sa.setWidget(lab)
        lay.addWidget(sa)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        bb.accepted.connect(self.accept)
        lay.addWidget(bb)
        self.resize(320, 280)


class RenumberDialog(QDialog):
    """Renumber Rows (DEV-8/rs1 stage 5): "Renumber Resource Id's" """

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Renumber Resource Id's")
        lay = QVBoxLayout(self)
        g = QGroupBox("Renumber Settings")
        gl = QFormLayout(g)
        self.renumber = QRadioButton("&Renumber Id's")
        self.rebase = QRadioButton("Re&base Id's")
        self.renumber.setChecked(True)
        gl.addRow(self.renumber)
        gl.addRow(self.rebase)
        self.base = QLineEdit()
        gl.addRow("&Base Value:", self.base)
        lay.addWidget(g)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.value = None

    def _ok(self):
        try:
            self.value = int(self.base.text().strip())
        except ValueError:
            return
        self.accept()


def renumber(ids, base, rebase):
    """the new ids of the selected rows: Renumber Id's from the base in order, Rebase Id's + the base"""
    if rebase:
        return [i + base for i in ids]
    return [base + k for k in range(len(ids))]


# ---------------------------------------------------------------- Edit String Tables
class StringTableEditor(QWidget):
    """the window "Edit String Tables": rows of ids, a column of each language"""
    modified = pyqtSignal()

    def __init__(self, editor):
        super().__init__()
        self.editor = editor
        self.undo_stack = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        self.tb = QToolBar()
        self.acts = {}
        for spec in (
                ("row", "Insert New Row", "row", self.insert_row),
                ("table", "Insert New String Table", "column", self.insert_table),
                ("props", "Properties", "props", self.table_properties), None,
                ("renumber", "Renumber Rows", "renumber", self.renumber_rows),
                ("cut", "Cut", "cut", self.cut), ("copy", "Copy", "copy", self.copy),
                ("paste", "Paste", "paste", self.paste), ("undo", "Undo", "undo", self.undo),
                ("export", "Export Text To Clipboard", "export", self.export), None,
                ("delrow", "Delete Row", "delrow", self.delete_row),
                ("deltable", "Delete String Table", "deltable", self.delete_table), None,
                ("help", "Help", "help", lambda: None)):
            if spec is None:
                self.tb.addSeparator()
                continue
            key, text, icon, slot = spec
            a = self.tb.addAction(_icon(icon), "")
            a.setToolTip(text)
            a.triggered.connect(lambda checked=False, f=slot: f())
            self.acts[key] = a
        lay.addWidget(self.tb)
        self.table = QTableWidget()
        self.table.verticalHeader().hide()
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.itemChanged.connect(self._item_changed)
        self.table.itemSelectionChanged.connect(self.update_actions)
        self.table.installEventFilter(self)
        lay.addWidget(self.table)
        self.langs = []
        self._filling = False
        self.combo = None                       # the Language list in the header of a new table (VB6)
        self.combo_lang = None
        hh = self.table.horizontalHeader()
        hh.sectionResized.connect(lambda *a: self._place_combo())
        hh.geometriesChanged.connect(self._place_combo)
        self.table.horizontalScrollBar().valueChanged.connect(lambda *a: self._place_combo())
        self.load()

    # ---- the table <-> the resource file
    @property
    def res(self):
        return self.editor.res

    def load(self):
        """the tables of the file, or one of the language of the system with the row 101 (DEV-8/rs1 picture 7)"""
        res = self.res
        self.langs = [lg for lg, tab in res.strings.items() if tab] or [res.lang]
        ids = res.string_ids() or [101]
        self._fill(ids)

    def _fill(self, ids):
        self._filling = True
        t = self.table
        t.clear()
        t.setColumnCount(1 + len(self.langs))
        t.setHorizontalHeaderLabels(["Id"] + [L.name(lg) for lg in self.langs])
        t.setRowCount(len(ids))
        for r, sid in enumerate(ids):
            t.setItem(r, 0, QTableWidgetItem(str(sid)))
            for c, lg in enumerate(self.langs):
                t.setItem(r, c + 1, QTableWidgetItem(self.res.strings.get(lg, {}).get(sid, "")))
        t.resizeColumnToContents(0)
        for c in range(len(self.langs)):
            t.setColumnWidth(c + 1, 200)
        self._filling = False
        self._place_combo()
        self.update_actions()

    # ---- the Language list in the header of a new table (DEV-8/rs1 stage 4, picture 25)
    def _place_combo(self):
        """the list stays over the header of its column, wherever the column goes"""
        if self.combo_lang not in self.langs:
            if self.combo is not None:
                self.combo.hide()
            self.combo_lang = None
            return
        hh = self.table.horizontalHeader()
        if self.combo is None:
            self.combo = QComboBox(hh.viewport())
            short_list(self.combo)
            for i, n in L.combo_items():
                self.combo.addItem(n, i)
            self.combo.activated.connect(lambda k: self.pick_lang(self.combo.itemData(k)))
        col = 1 + self.langs.index(self.combo_lang)
        k = self.combo.findData(self.combo_lang)
        if k < 0:
            self.combo.insertItem(0, L.name(self.combo_lang), self.combo_lang)
            k = 0
        self.combo.setCurrentIndex(k)
        self.combo.setGeometry(hh.sectionViewportPosition(col), 0, hh.sectionSize(col), hh.height())
        self.combo.show()
        self.combo.raise_()

    def pick_lang(self, new):
        """a language from the list in the header: the table takes it (one that another table has stays)"""
        old = self.combo_lang
        if old is None or new is None or new == old:
            return
        if new in self.langs:
            self._place_combo()
            return
        self.snapshot()
        rows = self.rows()
        self.langs[self.langs.index(old)] = new
        self.res.strings = {(new if k == old else k): v for k, v in self.res.strings.items()}
        self.combo_lang = new
        self._fill([sid for sid, _ in rows] or [101])
        self.store()
        self.editor.refresh()

    def rows(self):
        """[(id, [texts])] of the table"""
        out = []
        for r in range(self.table.rowCount()):
            try:
                sid = int(self.table.item(r, 0).text())
            except (ValueError, AttributeError):
                continue
            out.append((sid, [(self.table.item(r, c + 1).text() if self.table.item(r, c + 1) else "")
                              for c in range(len(self.langs))]))
        return out

    def store(self):
        """the table into the resource file"""
        res = self.res
        for lg in self.langs:
            res.strings[lg] = {}
        for sid, texts in self.rows():
            for lg, s in zip(self.langs, texts):
                if s:
                    res.strings[lg][sid] = s
        for lg in list(res.strings):
            if lg not in self.langs:
                del res.strings[lg]
        self.editor.set_modified()

    def snapshot(self):
        self.undo_stack.append((list(self.langs), self.rows()))
        del self.undo_stack[:-20]

    def _restore(self, state):
        langs, rows = state
        self.langs = list(langs)
        self._filling = True
        self.table.setRowCount(0)
        self._filling = False
        self.res.strings = {lg: {sid: t[k] for sid, t in rows if t[k]} for k, lg in enumerate(self.langs)}
        self._fill([sid for sid, _ in rows] or [101])
        self.editor.set_modified()

    # ---- editing
    def _item_changed(self, item):
        if self._filling:
            return
        if item.column() == 0:                  # a new id: the rows are sorted again (DEV-8/rs1 stage 4)
            try:
                sid = int(item.text())
            except ValueError:
                return
            rows = sorted(self.rows(), key=lambda x: x[0])
            self._fill([s for s, _ in rows])
            self._restore_texts(rows)
            for r in range(self.table.rowCount()):
                if self.table.item(r, 0).text() == str(sid):
                    self.table.setCurrentCell(r, 0)
                    break
        self.store()
        self.editor.refresh()

    def _restore_texts(self, rows):
        self._filling = True
        by_id = dict(rows)
        for r in range(self.table.rowCount()):
            sid = int(self.table.item(r, 0).text())
            for c in range(len(self.langs)):
                self.table.item(r, c + 1).setText(by_id.get(sid, [""] * len(self.langs))[c])
        self._filling = False

    def eventFilter(self, obj, e):
        from PyQt6.QtCore import QEvent
        if obj is self.table and e.type() == QEvent.Type.KeyPress and \
                e.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and self.table.state() != \
                QAbstractItemView.State.EditingState:
            self.enter_pressed()
            return True
        return super().eventFilter(obj, e)

    def enter_pressed(self):
        """Enter: the next row; on the last row a new row with its id + 1 (DEV-8/rs1 stages 2, 4). Enter that
        ends the writing in a cell only ends it, without a new row: the user's choice (visual check of 8.6)"""
        r, c = self.table.currentRow(), max(self.table.currentColumn(), 1)
        if r == self.table.rowCount() - 1:
            self.add_row(int(self.table.item(r, 0).text()) + 1)
        self.table.setCurrentCell(min(r + 1, self.table.rowCount() - 1), c)

    def add_row(self, sid):
        self._filling = True
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(str(sid)))
        for c in range(len(self.langs)):
            self.table.setItem(r, c + 1, QTableWidgetItem(""))
        self._filling = False
        self.update_actions()
        return r

    def insert_row(self):
        self.snapshot()
        ids = [sid for sid, _ in self.rows()]
        sid = (max(ids) + 1) if ids else 101
        r = self.add_row(sid)
        self.table.setCurrentCell(r, 1)

    def insert_table(self, lang=None):
        """Insert New String Table: a new column at once, with the Language list in its header; the first
        language of the list that no table has (VB6: 1112, DEV-8/rs1 stage 4)"""
        if lang is None:
            lang = next((i for i, _ in L.combo_items() if i not in self.langs), None)
        if lang is None or lang in self.langs:
            return
        self.snapshot()
        rows = self.rows()
        self.langs.append(lang)
        self.combo_lang = lang
        self._fill([sid for sid, _ in rows] or [101])
        self.store()
        self.editor.refresh()

    def current_lang(self):
        c = self.table.currentColumn()
        return self.langs[c - 1] if 1 <= c <= len(self.langs) else (self.langs[0] if self.langs else None)

    def table_properties(self):
        lg = self.current_lang()
        if lg is None:
            return
        dlg = PropertiesDialog(self, None, string_lang=lg)
        if dlg.exec() and dlg.result_values:
            new = dlg.result_values["lang"]
            if new != lg and new not in self.langs:
                self.snapshot()
                rows = self.rows()
                self.langs[self.langs.index(lg)] = new
                self.res.strings = {(new if k == lg else k): v for k, v in self.res.strings.items()}
                self._fill([sid for sid, _ in rows])
                self.editor.set_modified()
                self.editor.refresh()

    def selected_rows(self):
        return sorted({i.row() for i in self.table.selectedIndexes()})

    def renumber_rows(self, base=None, rebase=False):
        rows = self.selected_rows()
        if len(rows) < 2:
            return
        if base is None:
            dlg = RenumberDialog(self)
            if not dlg.exec():
                return
            base, rebase = dlg.value, dlg.rebase.isChecked()
        self.snapshot()
        all_rows = self.rows()
        sel = [all_rows[r][0] for r in rows]
        new = dict(zip(sel, renumber(sel, base, rebase)))
        moved = [(new.get(sid, sid), texts) for sid, texts in all_rows]
        moved.sort(key=lambda x: x[0])
        self._filling = True
        self._fill([sid for sid, _ in moved])
        self._restore_texts(moved)
        self.store()

    def cut(self):
        self.copy()
        self.snapshot()
        for i in self.table.selectedItems():
            if i.column() > 0:
                i.setText("")

    def copy(self):
        it = self.table.currentItem()
        if it is not None:
            QApplication.clipboard().setText(it.text())

    def paste(self):
        it = self.table.currentItem()
        if it is not None and it.column() > 0:
            self.snapshot()
            it.setText(QApplication.clipboard().text())

    def undo(self):
        if self.undo_stack:
            self._restore(self.undo_stack.pop())
        self.update_actions()

    def export(self):
        lines = ["\t".join(["Id"] + [L.name(lg) for lg in self.langs])]
        for sid, texts in self.rows():
            lines.append("\t".join([str(sid)] + texts))
        QApplication.clipboard().setText("\r\n".join(lines))

    def delete_row(self):
        rows = self.selected_rows() or ([self.table.currentRow()] if self.table.currentRow() >= 0 else [])
        if not rows:
            return
        self.snapshot()
        for r in reversed(rows):
            self.table.removeRow(r)
        self.store()
        self.update_actions()

    def delete_table(self):
        lg = self.current_lang()
        if lg is None or len(self.langs) < 1:
            return
        self.snapshot()
        rows = self.rows()
        k = self.langs.index(lg)
        self.langs.remove(lg)
        self.res.strings.pop(lg, None)
        rows = [(sid, t[:k] + t[k + 1:]) for sid, t in rows]
        if not self.langs:
            self.editor.close_strings()
            self.editor.set_modified()
            self.editor.refresh()
            return
        self._fill([sid for sid, _ in rows])
        self._restore_texts(rows)
        self.store()
        self.editor.refresh()

    def update_actions(self):
        if not hasattr(self, "acts"):
            return
        self.acts["renumber"].setEnabled(len(self.selected_rows()) > 1)
        self.acts["undo"].setEnabled(bool(self.undo_stack))


# ---------------------------------------------------------------- the Resource Editor
class ResourceEditor(QWidget):
    """the window "VB Resource Editor - <file>" of the active project"""
    title_changed = pyqtSignal(str)

    def __init__(self, main):
        super().__init__()
        self.main = main
        self.res = R.ResFile(L.system_lcid())
        self.project = None
        self.modified = False
        self.undo_item = None                   # (index, item) of the last Delete
        self.strings_win = None
        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        self.tb = QToolBar()
        self.acts = {}
        for spec in (
                ("new", "New Resource File", "new", self.new_file),
                ("open", "Open Resource File", "open", self.open_file),
                ("save", "Save Resource File", "save", self.save),
                ("props", "Edit Properties", "props", self.edit_properties), None,
                ("dup", "Duplicate", "duplicate", self.duplicate),
                ("undo", "Undo", "undo", self.undo),
                ("delete", "Delete", "delete", self.delete), None,
                ("strings", "Edit String Tables", "strings", self.edit_strings),
                ("cursor", "Add Cursor", "cursor", lambda: self.add_file("cursor")),
                ("icon", "Add Icon", "icon", lambda: self.add_file("icon")),
                ("bitmap", "Add Bitmap", "bitmap", lambda: self.add_file("bitmap")),
                ("custom", "Add Custom Resource", "custom", lambda: self.add_file("custom")), None,
                ("help", "Show Help", "help", lambda: None)):
            if spec is None:
                self.tb.addSeparator()
                continue
            key, text, icon, slot = spec
            a = self.tb.addAction(_icon(icon), "")
            a.setToolTip(text)
            a.triggered.connect(lambda checked=False, f=slot: f())
            self.acts[key] = a
        lay.addWidget(self.tb)
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.itemDoubleClicked.connect(self._open_item)
        self.tree.currentItemChanged.connect(lambda *a: self.update_actions())
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._menu)
        lay.addWidget(self.tree)
        sb = QHBoxLayout()
        self.status_icon = QLabel()
        self.status = QLabel("")
        sb.addWidget(self.status_icon)
        sb.addWidget(self.status, 1)
        lay.addLayout(sb)
        self.last_dir = None
        self.refresh()

    # ---- the file of the project
    def title(self):
        return "VB Resource Editor - " + (self.res.path or "")

    def project_dir(self):
        p = self.project
        return os.path.dirname(p.path) if p is not None and p.path else None

    def bind(self, project):
        """the resource file of a project (ResFile32=), or an empty one"""
        self.project = project
        self.res = R.ResFile(L.system_lcid())
        self.undo_item = None
        path = res_path(project)
        if path and os.path.isfile(path):
            try:
                self.res = R.ResFile.load(path, L.system_lcid())
            except (OSError, ValueError, struct.error):
                pass
        self.modified = False
        self.close_strings()
        self.refresh()

    def set_modified(self, on=True):
        self.modified = on
        self.refresh_status()

    def refresh_status(self):
        self.status.setText("(Modified)" if self.modified else "")
        self.status_icon.setPixmap(_icon("save").pixmap(14, 14) if self.modified else QPixmap())
        self.title_changed.emit(self.title())

    def refresh(self):
        sel = self.selected()
        self.tree.clear()
        root = QTreeWidgetItem([self.res.path or "Resource File"])
        root.setIcon(0, _icon("folder"))
        self.tree.addTopLevelItem(root)
        want = None
        for folder, items in R.tree(self.res):
            f = QTreeWidgetItem([folder])
            f.setIcon(0, _icon("folder"))
            root.addChild(f)
            for cap, obj in items:
                it = QTreeWidgetItem([cap])
                it.setData(0, Qt.ItemDataRole.UserRole, obj)
                kind = obj[0] if isinstance(obj, tuple) else obj.kind
                it.setIcon(0, _icon({"string": "strings"}.get(kind, kind)))
                f.addChild(it)
                if obj is sel or (isinstance(obj, tuple) and obj == sel):
                    want = it
        self.tree.expandAll()
        self.tree.setCurrentItem(want or root)
        self.refresh_status()
        self.update_actions()

    def selected(self):
        it = self.tree.currentItem()
        return it.data(0, Qt.ItemDataRole.UserRole) if it is not None else None

    def select(self, obj):
        for i in range(self.tree.topLevelItemCount()):
            root = self.tree.topLevelItem(i)
            for k in range(root.childCount()):
                f = root.child(k)
                for j in range(f.childCount()):
                    it = f.child(j)
                    d = it.data(0, Qt.ItemDataRole.UserRole)
                    if d is obj or (isinstance(d, tuple) and d == obj):
                        self.tree.setCurrentItem(it)
                        return

    def update_actions(self):
        obj = self.selected()
        item = obj if isinstance(obj, R.Item) else None
        self.acts["props"].setEnabled(obj is not None)
        self.acts["dup"].setEnabled(item is not None)
        self.acts["delete"].setEnabled(obj is not None)
        self.acts["undo"].setEnabled(self.undo_item is not None)

    # ---- the buttons
    def ask_save(self):
        """the question of VB6 when the file changed; True to go on"""
        if not self.modified:
            return True
        name = self.project.name if self.project is not None else "Project1"
        r = QMessageBox.warning(self, TITLE, changed_msg(name),
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            return self.save()
        self.modified = False
        return True

    def new_file(self):
        if not self.ask_save():
            return
        self.res = R.ResFile(L.system_lcid())
        self.undo_item = None
        self.close_strings()
        self.set_modified(False)
        self.refresh()

    def open_file(self, path=None):
        if path is None:
            if not self.ask_save():
                return
            path, _ = QFileDialog.getOpenFileName(self, "Open A Resource File", self.start_dir(),
                                                  "Resource Files (*.RES *.res);;All Files (*)")
            if not path:
                return
        self.res = R.ResFile.load(path, L.system_lcid())
        self.undo_item = None
        self.close_strings()
        self.set_modified(False)
        self.attach(path)
        self.refresh()

    def start_dir(self):
        return self.last_dir or (os.path.dirname(self.res.path) if self.res.path else None) or \
            self.project_dir() or os.path.expanduser("~")

    def save(self):
        if self.res.path is None:
            return self.save_as()
        self.res.save()
        self.set_modified(False)
        self.refresh()
        return True

    def save_as(self, path=None):
        if path is None:
            name = (self.project.name if self.project is not None else "Project1") + ".RES"
            path, _ = QFileDialog.getSaveFileName(self, "Save Resource File As",
                                                  os.path.join(self.start_dir(), name), "Resource File (*.RES)")
            if not path:
                return False
        base, ext = os.path.splitext(path)
        path = base + ".RES" if ext.lower() in (".res", "") else path   # VB6 writes s1.RES
        self.res.save(path)
        self.set_modified(False)
        self.attach(path)
        self.refresh()
        return True

    def attach(self, path):
        """the file becomes the resource file of the project: ResFile32= and Related Documents"""
        p = self.project
        if p is None:
            return
        rel = path
        if p.path:
            try:
                rel = os.path.relpath(path, os.path.dirname(p.path))
            except ValueError:
                rel = path
        if not os.path.isabs(rel) and not rel.startswith(".."):
            rel = rel.replace("/", "\\")
        if (p.vbp.get("ResFile32") or "").strip('"') != rel:
            set_res_file(p, rel)
            if hasattr(self.main, "fill_tree"):
                self.main.fill_tree()
                self.main.update_title()

    def edit_properties(self):
        obj = self.selected()
        if isinstance(obj, tuple):
            self.edit_strings()
            if self.strings_win is not None:
                w = self.strings_win.widget()
                if obj[1] in w.langs:
                    w.table.setCurrentCell(0, w.langs.index(obj[1]) + 1)
                w.table_properties()
            return
        if not isinstance(obj, R.Item):
            return
        dlg = PropertiesDialog(self, obj)
        if dlg.exec() and dlg.result_values:
            self.apply_properties(obj, dlg.result_values)

    def apply_properties(self, it, vals):
        if "type" in vals:
            it.rtype = vals["type"]
        if "name" in vals:
            it.name = vals["name"]
        it.lang = vals.get("lang", it.lang)
        self.set_modified()
        self.refresh()
        self.select(it)

    def duplicate(self):
        obj = self.selected()
        if isinstance(obj, R.Item):
            new = self.res.duplicate(obj)
            self.set_modified()
            self.refresh()
            self.select(new)

    def delete(self):
        obj = self.selected()
        if isinstance(obj, R.Item):
            i = self.res.remove(obj)
            self.undo_item = (i, obj)
        elif isinstance(obj, tuple):
            lg = obj[1]
            self.undo_item = (-1, (lg, dict(self.res.strings.get(lg, {}))))
            self.res.strings.pop(lg, None)
            self.close_strings()
        else:
            return
        self.set_modified()
        self.refresh()

    def undo(self):
        """one step: the last Delete comes back (DEV-8/rs1 stage 4)"""
        if self.undo_item is None:
            return
        i, obj = self.undo_item
        self.undo_item = None
        if i >= 0:
            self.res.items.insert(i, obj)
        else:
            lg, tab = obj
            self.res.strings[lg] = tab
        self.set_modified()
        self.refresh()
        self.select(obj if i >= 0 else ("string", obj[0]))

    def edit_strings(self):
        if self.strings_win is None:
            w = StringTableEditor(self)
            self.strings_win = self.main.add_res_window(w, "Edit String Tables")
        else:
            self.main.show_res_window(self.strings_win)
        self.refresh()

    def close_strings(self):
        if self.strings_win is not None:
            sw, self.strings_win = self.strings_win, None
            self.main.remove_res_window(sw)

    def strings_closed(self):
        """the X of "Edit String Tables": the window closes without a question (DEV-8/rs1 stage 2)"""
        self.strings_win = None
        self.refresh()

    def add_file(self, kind, path=None):
        titles = {"bitmap": ("Open A Bitmap File", "Bitmaps (*.BMP *.bmp)"),
                  "icon": ("Open An Icon File", "Icons (*.ICO *.ico)"),
                  "cursor": ("Open A Cursor File", "Cursors (*.CUR *.cur)"),
                  "custom": ("Open A Custom Resource", "Custom (*.*)")}
        if path is None:
            title, filt = titles[kind]
            path, _ = QFileDialog.getOpenFileName(self, title, self.start_dir(), filt)
            if not path:
                return None
        self.last_dir = os.path.dirname(path)
        with open(path, "rb") as f:
            data = f.read()
        try:
            if kind == "bitmap":
                it = self.res.add_bitmap(data)
            elif kind in ("icon", "cursor"):
                it = self.res.add_icon(data, cursor=kind == "cursor")
            else:
                it = self.res.add_custom(data)
        except (ValueError, struct.error) as e:
            QMessageBox.warning(self, TITLE, f"{os.path.basename(path)}: {e}")
            return None
        self.set_modified()
        self.refresh()
        self.select(it)
        return it

    def _open_item(self, it, col):
        obj = it.data(0, Qt.ItemDataRole.UserRole)
        if isinstance(obj, tuple):
            self.edit_strings()
        elif isinstance(obj, R.Item):
            self.edit_properties()

    def _menu(self, pos):
        it = self.tree.itemAt(pos)
        m = QMenu(self)
        if it is None or it.data(0, Qt.ItemDataRole.UserRole) is None:
            m.addAction("New", self.new_file)
            m.addAction("Open", self.open_file)
            m.addSeparator()
            m.addAction("Save", self.save)
            m.addAction("Save As...", self.save_as)
        else:
            self.tree.setCurrentItem(it)
            m.addAction("Edit Properties", self.edit_properties)
            m.addAction("Duplicate", self.duplicate)
            m.addAction("Delete", self.delete)
        m.exec(self.tree.viewport().mapToGlobal(pos))


# ---------------------------------------------------------------- the .vbp
def res_path(project):
    """the full path of the resource file of a project (ResFile32=), or None"""
    if project is None:
        return None
    rel = (project.vbp.get("ResFile32") or "").strip('"')
    if not rel:
        return None
    rel = rel.replace("\\", "/")
    if os.path.isabs(rel) or not project.path:
        return rel
    try:
        from lbvc.project.fsutil import ci_path
        return ci_path(os.path.join(os.path.dirname(project.path), rel))
    except ImportError:
        return os.path.join(os.path.dirname(project.path), rel)


def set_res_file(project, rel):
    """ResFile32="<file>" where VB6 writes it (after the members and the references, DEV-8/rs1 s1.vbp), or
    no resource file"""
    from . import projprops as P
    if rel:
        P.put(project.vbp, "ResFile32", rel, quote=True)
    else:
        P.drop(project.vbp, "ResFile32")
    project.dirty = True
