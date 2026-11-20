"""The VB Application Wizard (phase 9.2a-3): File > New Project > VB Application Wizard, or Add-Ins > Application
Wizard... when the "VB 6 Application Wizard" add-in is loaded (VB6 oracle DEV-9/aw1, pictures p9/oracle/aw1).

The pages of VB6: Introduction (a profile .rwp to load), Interface Type (MDI / SDI / Explorer Style, the name),
Menus, Customize Toolbar, Resources, Internet Connectivity, Standard Forms (+ Form Templates...), Data Access Forms,
Finished! (a profile to save, View Report). Finish: the new project (appwizgen), the String Table in the Resource
Editor (not saved, as VB6: "stored in memory ... until the project has been saved"), the profile, "Application
Created". The wording of the pages is our own; the questions and the names of the choices are VB6's.
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
import os
import re

from PyQt6.QtCore import Qt, QSize, QRectF, QPointF
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPen, QIcon, QPolygonF, QImage
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit, QListWidget,
                             QListWidgetItem, QPlainTextEdit, QStackedWidget, QRadioButton, QButtonGroup,
                             QCheckBox, QComboBox, QFrame, QMessageBox, QFileDialog, QInputDialog, QToolBar,
                             QGroupBox, QToolButton)

from . import appwizgen as G
from . import paths

TITLE = "Application Wizard"
PAGES = ("Introduction", "Interface Type", "Menus", "Customize Toolbar", "Resources", "Internet Connectivity",
         "Standard Forms", "Data Access Forms", "Finished!")
HINTS = {G.APP_MDI: "Consists of a main window which visually contains child windows.",
         G.APP_SDI: "All windows exist at the same level.",
         G.APP_EXPLORER: "Similar to SDI with main form containing TreeView and ListView controls."}
NONE = "(None)"
PROFILE_FILTER = "Wizard Profile (*.rwp)"


def profiles_dir():
    d = os.path.join(paths.config_dir(), "wizards")
    os.makedirs(d, exist_ok=True)
    return d


# ------------------------------------------------------------------------------------------------ the pictures
def wizard_picture(kind, w=150, h=120):
    """the picture on the left of a page: our own drawings"""
    pm = QPixmap(w, h)
    pm.fill(QColor("#808080"))
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    ink, paper, blue, gold = QColor("#202020"), QColor("#FFFFFF"), QColor("#2F6FD0"), QColor("#F2C230")

    def window(x, y, ww, hh, title=True):
        p.setPen(QPen(ink, 1))
        p.setBrush(paper)
        p.drawRect(QRectF(x, y, ww, hh))
        if title:
            p.setBrush(blue)
            p.drawRect(QRectF(x, y, ww, 7))
    if kind == "mdi":
        window(20, 15, 110, 90)
        window(30, 32, 55, 40)
        window(55, 50, 60, 45)
    elif kind == "sdi":
        window(18, 18, 70, 55)
        window(62, 50, 70, 55)
    elif kind == "explorer":
        window(15, 20, 120, 85)
        p.setPen(QPen(QColor("#A0A0A0"), 1))
        p.drawLine(55, 28, 55, 104)
        p.setPen(QPen(ink, 1))
        for i, y in enumerate((34, 44, 54, 64)):
            p.drawLine(22 + (i % 2) * 6, y, 48, y)
        for y in (34, 50, 66, 82):
            for x in (64, 88, 112):
                p.fillRect(QRectF(x, y, 12, 10), QColor("#F0C060"))
    elif kind == "menus":
        window(20, 18, 110, 85)
        p.setPen(QPen(ink, 1))
        f = p.font()
        f.setPixelSize(8)
        p.setFont(f)
        p.drawText(QRectF(26, 27, 100, 10), "File  Edit  View")
        p.setBrush(QColor("#F0F0F0"))
        p.drawRect(QRectF(26, 38, 50, 52))
        for y in (46, 56, 66, 76, 86):
            p.drawLine(30, y, 70, y)
    elif kind == "toolbar":
        window(15, 30, 120, 60)
        for i in range(6):
            p.setBrush(QColor("#E0E0E0"))
            p.drawRect(QRectF(20 + i * 18, 41, 15, 14))
    elif kind == "resources":
        p.setPen(QPen(ink, 1))
        for i, c in enumerate(("#D0E4FA", "#FAE4C8", "#D8F0D0")):
            p.setBrush(QColor(c))
            p.drawRect(QRectF(30 + i * 12, 20 + i * 12, 70, 60))
        p.drawText(QRectF(56, 66, 60, 20), "abc")
    elif kind == "internet":
        p.setPen(QPen(ink, 1.2))
        p.setBrush(QColor("#9CC8F4"))
        p.drawEllipse(QRectF(35, 20, 80, 80))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawEllipse(QRectF(58, 20, 34, 80))
        p.drawLine(35, 60, 115, 60)
        p.drawLine(40, 40, 110, 40)
        p.drawLine(40, 80, 110, 80)
    elif kind == "forms":
        window(25, 20, 70, 50)
        window(55, 50, 70, 50)
        p.setBrush(QColor("#E0E0E0"))
        p.drawRect(QRectF(95, 85, 24, 9))
    elif kind == "data":
        p.setPen(QPen(ink, 1))
        for i in range(3):
            p.setBrush(QColor("#E8D27A"))
            p.drawRect(QRectF(45, 30 + i * 22, 60, 18))
            p.drawEllipse(QRectF(45, 22 + i * 22, 60, 14))
    elif kind == "finished":
        p.setPen(QPen(ink, 2))
        p.drawLine(45, 20, 45, 105)
        for r in range(4):
            for c in range(5):
                if (r + c) % 2 == 0:
                    p.fillRect(QRectF(46 + c * 12, 22 + r * 10, 12, 10), ink)
                else:
                    p.fillRect(QRectF(46 + c * 12, 22 + r * 10, 12, 10), paper)
    else:                                       # the wand of the wizard
        p.setPen(QPen(ink, 5, cap=Qt.PenCapStyle.RoundCap))
        p.drawLine(40, 100, 95, 45)
        p.setPen(QPen(QColor("#8A6A10"), 1))
        p.setBrush(gold)
        cx, cy = 102, 38
        pts = []
        import math
        for i in range(10):
            r = 18 if i % 2 == 0 else 8
            a = math.pi / 2 + i * math.pi / 5
            pts.append(QPointF(cx + r * math.cos(a), cy - r * math.sin(a)))
        p.drawPolygon(QPolygonF(pts))
    p.end()
    return pm


def wizard_icon(size=16):
    """the icon of Add-Ins > Application Wizard... and of New Project (a wand)"""
    return QIcon(wizard_picture("intro", 64, 64).scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio,
                                                         Qt.TransformationMode.SmoothTransformation))


# ------------------------------------------------------------------------------------------------ helpers
def key_of(caption):
    k = re.sub(r"[^A-Za-z0-9_]", "", caption.replace("&", ""))
    return k or "Item"


def bmp16(path):
    """an image file of the user as a 16 x 16 BMP on the mask colour (the image button of Customize Toolbar)"""
    img = QImage(path)
    if img.isNull():
        return None
    img = img.scaled(16, 16, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
    out = QImage(16, 16, QImage.Format.Format_RGB32)
    out.fill(QColor(0xC0, 0xC0, 0xC0))
    p = QPainter(out)
    p.drawImage((16 - img.width()) // 2, (16 - img.height()) // 2, img)
    p.end()
    from PyQt6.QtCore import QByteArray, QBuffer, QIODevice
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QIODevice.OpenModeFlag.WriteOnly)
    out.convertToFormat(QImage.Format.Format_RGB888).save(buf, "BMP")
    buf.close()
    return bytes(ba)


class ToolbarEditor(QFrame):
    """Customize Toolbar (aw4, aw20): the preview, the buttons there are (left, alphabetically, with [Separator]),
    the buttons of the toolbar (right); > < << up down, a picture of the user; the list is edited in place"""

    def __init__(self, parent, images, reset=None):
        super().__init__(parent)
        self.images = images                    # key -> BMP of a button with a picture of the user
        self.buttons = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.preview = QToolBar()
        self.preview.setIconSize(QSize(16, 16))
        self.preview.setMinimumHeight(28)
        self.preview.setStyleSheet("QToolBar { border: 1px solid palette(mid); }")
        lay.addWidget(self.preview)
        row = QHBoxLayout()
        self.avail = QListWidget()
        self.avail.setIconSize(QSize(16, 16))
        self.chosen = QListWidget()
        self.chosen.setIconSize(QSize(16, 16))
        row.addWidget(self.avail, 1)
        col = QVBoxLayout()
        col.addStretch(1)
        for text, tip, slot in ((">", "Add the button", self.add), ("<", "Remove the button", self.remove),
                                ("<<", "Remove all the buttons", self.clear),
                                ("\u2191", "Move Up", lambda: self.move(-1)),
                                ("\u2193", "Move Down", lambda: self.move(1)),
                                ("Picture...", "A button with a picture of your own", self.picture)):
            b = QPushButton(text)
            b.setToolTip(tip)
            b.clicked.connect(slot)
            col.addWidget(b)
        col.addStretch(1)
        row.addLayout(col)
        row.addWidget(self.chosen, 1)
        lay.addLayout(row, 1)
        if reset is not None:
            r = QHBoxLayout()
            r.addStretch(1)
            b = QPushButton("&Reset")
            b.clicked.connect(reset)
            r.addWidget(b)
            lay.addLayout(r)
        self.avail.itemDoubleClicked.connect(lambda _i: self.add())
        self.chosen.itemDoubleClicked.connect(lambda _i: self.remove())

    def icon(self, key):
        from . import tbicons
        if key in self.images:
            img = QImage()
            img.loadFromData(self.images[key])
            pm = QPixmap.fromImage(img)
            pm.setMask(pm.createMaskFromColor(QColor(0xC0, 0xC0, 0xC0)))
            return QIcon(pm)
        return tbicons.icon(key)

    def set_buttons(self, buttons):
        self.buttons = buttons
        self.fill()

    def fill(self):
        from . import tbicons
        on = {b.key for b in self.buttons if not b.separator}
        arow = self.avail.currentRow()
        self.avail.clear()
        it = QListWidgetItem(G.SEPARATOR)
        it.setData(Qt.ItemDataRole.UserRole, "")
        self.avail.addItem(it)
        for name in sorted(set(tbicons.ALL) | set(self.images), key=str.lower):
            if name in on:
                continue
            it = QListWidgetItem(self.icon(name), name)
            it.setData(Qt.ItemDataRole.UserRole, name)
            self.avail.addItem(it)
        self.avail.setCurrentRow(max(0, min(arow, self.avail.count() - 1)))
        cur = self.chosen.currentRow()
        self.chosen.clear()
        for b in self.buttons:
            self.chosen.addItem(QListWidgetItem(G.SEPARATOR) if b.separator else
                                QListWidgetItem(self.icon(b.key), b.caption))
        self.chosen.setCurrentRow(min(cur, self.chosen.count() - 1))
        self.preview.clear()
        for b in self.buttons:
            if b.separator:
                self.preview.addSeparator()
            else:
                a = self.preview.addAction(self.icon(b.key), b.caption)
                a.setToolTip(b.tip)

    def _at(self):
        return self.chosen.currentRow() + 1 if self.chosen.currentRow() >= 0 else len(self.buttons)

    def add(self):
        it = self.avail.currentItem()
        if it is None:
            return
        key = it.data(Qt.ItemDataRole.UserRole)
        r = self._at()
        self.buttons.insert(r, G.TbButton(key) if key else G.TbButton(""))
        self.fill()
        self.chosen.setCurrentRow(r)

    def remove(self):
        r = self.chosen.currentRow()
        if 0 <= r < len(self.buttons):
            del self.buttons[r]
            self.fill()

    def clear(self):
        del self.buttons[:]
        self.fill()

    def move(self, d):
        r = self.chosen.currentRow()
        tb = self.buttons
        if 0 <= r < len(tb) and 0 <= r + d < len(tb):
            tb[r], tb[r + d] = tb[r + d], tb[r]
            self.fill()
            self.chosen.setCurrentRow(r + d)

    def picture(self):
        path, _ = QFileDialog.getOpenFileName(self, "Picture", os.path.expanduser("~"),
                                              "Pictures (*.bmp *.ico *.png *.gif *.jpg *.jpeg)")
        if not path:
            return
        data = bmp16(path)
        if data is None:
            QMessageBox.warning(self, TITLE, f"{os.path.basename(path)} is not a picture.")
            return
        key = os.path.splitext(os.path.basename(path))[0]
        self.images[key] = data
        if not any(b.key == key for b in self.buttons):
            self.buttons.insert(self._at(), G.TbButton(key))
        self.fill()


def toolbar_wizard_icon(size=16):
    """the icon of Add-Ins > Toolbar Wizard... (a toolbar with buttons)"""
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    s = size / 16.0
    p.scale(s, s)
    p.setPen(QPen(QColor("#202020"), 1))
    p.setBrush(QColor("#E8E8E8"))
    p.drawRect(0, 4, 15, 8)
    for i, c in enumerate(("#2F6FD0", "#D03030", "#2E9A3A")):
        p.fillRect(2 + i * 4, 6, 3, 4, QColor(c))
    p.end()
    return QIcon(pm)


class FormTemplatesDialog(QDialog):
    """Standard Forms > Form Templates... (aw7-02)"""

    def __init__(self, parent, chosen):
        super().__init__(parent)
        self.setWindowTitle("Form Templates")
        lay = QVBoxLayout(self)
        t = QLabel("Choose the form templates to add to your application. You can choose as many as you like.")
        t.setWordWrap(True)
        lay.addWidget(t)
        box = QGroupBox("Form Templates")
        bl = QHBoxLayout(box)
        self.list = QListWidget()
        for name in G.TEMPLATES:
            it = QListWidgetItem(name)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Checked if name in chosen else Qt.CheckState.Unchecked)
            self.list.addItem(it)
        bl.addWidget(self.list, 1)
        col = QVBoxLayout()
        for text, on in (("Select &All", True), ("&Clear All", False)):
            b = QPushButton(text)
            b.clicked.connect(lambda _=False, v=on: self.set_all(v))
            col.addWidget(b)
        col.addStretch(1)
        bl.addLayout(col)
        lay.addWidget(box)
        row = QHBoxLayout()
        row.addStretch(1)
        for text, slot in (("OK", self.accept), ("Cancel", self.reject)):
            b = QPushButton(text)
            b.clicked.connect(slot)
            row.addWidget(b)
        b = QPushButton("Help")
        b.clicked.connect(lambda: parent.main.help_keyword("Application Wizard") if hasattr(parent, "main") else None)
        row.addWidget(b)
        lay.addLayout(row)
        self.resize(420, 260)

    def set_all(self, on):
        for i in range(self.list.count()):
            self.list.item(i).setCheckState(Qt.CheckState.Checked if on else Qt.CheckState.Unchecked)

    def chosen(self):
        return [self.list.item(i).text() for i in range(self.list.count())
                if self.list.item(i).checkState() == Qt.CheckState.Checked]


class ReportDialog(QDialog):
    """View Report: the Summary Report (aw9-02: Close, Save)"""

    def __init__(self, parent, text):
        super().__init__(parent)
        self.setWindowTitle("Application Wizard Summary Report")
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMinMaxButtonsHint)
        text = text.replace("\r\n", "\n")
        self.text = text
        lay = QVBoxLayout(self)
        ed = QPlainTextEdit(text)
        ed.setReadOnly(True)
        lay.addWidget(ed, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        b = QPushButton("&Close")
        b.clicked.connect(self.accept)
        row.addWidget(b)
        b = QPushButton("&Save")
        b.clicked.connect(self.save)
        row.addWidget(b)
        lay.addLayout(row)
        self.resize(560, 420)

    def save(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Report", os.path.expanduser("~/Summary Report.txt"),
                                              "Text Files (*.txt)")
        if path:
            with open(path, "w", encoding="utf-8", newline="") as f:
                f.write(self.text.replace("\n", "\r\n"))


REPORT = ("Application Wizard Instructions:\n\n"
          "The Application Wizard has made the start of your application. The code has comments where your own "
          "code goes: search for \"TODO\" in the project (Edit > Find) to find them.\n\n"
          "What To Do Next\n===============\n\n"
          "- Save this report with \"Save\", or close it with \"Close\".\n"
          "- Write the code of the TODO comments.\n"
          "- Set the name, the version and the icon of the program in Project > Properties.\n"
          "- Run the application in LinBASIC Studio: press F5 or choose Start from the Run menu.\n\n"
          "You can go on with the project as it is, or use the other wizards of LinBASIC Studio to add to it as "
          "they become available (the Data Form Wizard, the Class Builder ...).\n\n"
          "Your choices\n============\n\n")


class AppWizard(QDialog):
    """the pages of the Application Wizard and Finish"""

    def __init__(self, main):
        super().__init__(main)
        self.main = main
        self.opts = G.AppOptions(G.APP_MDI)
        self.images = {}                        # key -> BMP of a button with an image of the user
        self.save_profile = None
        self.created = None                     # the project, after Finish
        outer = QVBoxLayout(self)
        body = QHBoxLayout()
        self.picture = QLabel()
        self.picture.setFixedSize(150, 120)
        self.picture.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(self.picture, 0, Qt.AlignmentFlag.AlignTop)
        self.stack = QStackedWidget()
        body.addWidget(self.stack, 1)
        outer.addLayout(body, 1)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        outer.addWidget(line)
        row = QHBoxLayout()
        self.b_help = QPushButton("Help")
        self.b_help.clicked.connect(lambda: main.help_keyword("Application Wizard"))
        row.addWidget(self.b_help)
        row.addStretch(1)
        self.b_cancel = QPushButton("Cancel")
        self.b_back = QPushButton("< &Back")
        self.b_next = QPushButton("&Next >")
        self.b_finish = QPushButton("&Finish")
        for b in (self.b_cancel, self.b_back, self.b_next, self.b_finish):
            row.addWidget(b)
        self.b_cancel.clicked.connect(self.reject)
        self.b_back.clicked.connect(lambda: self.go(self.stack.currentIndex() - 1))
        self.b_next.clicked.connect(lambda: self.go(self.stack.currentIndex() + 1))
        self.b_finish.clicked.connect(self.finish)
        outer.addLayout(row)
        for build in (self._intro_page, self._type_page, self._menus_page, self._toolbar_page, self._res_page,
                      self._internet_page, self._forms_page, self._data_page, self._finished_page):
            build()
        self.resize(720, 470)
        self.load_options()
        self.go(0)

    # ------------------------------------------------------------ the pages
    def _page(self, text):
        w = QFrame()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(6, 0, 0, 0)
        if text:
            t = QLabel(text)
            t.setWordWrap(True)
            lay.addWidget(t)
        self.stack.addWidget(w)
        return lay

    def _profile_row(self, lay, label, save):
        lab = QLabel(label)
        lab.setWordWrap(True)
        lay.addWidget(lab)
        row = QHBoxLayout()
        combo = QComboBox()
        combo.addItem(NONE)
        for f in sorted(os.listdir(profiles_dir())):
            if f.lower().endswith(".rwp"):
                combo.addItem(os.path.join(profiles_dir(), f))
        row.addWidget(combo, 1)
        b = QPushButton("...")
        b.setFixedWidth(32)
        b.clicked.connect(lambda: self.pick_profile(combo, save))
        row.addWidget(b)
        lay.addLayout(row)
        return combo

    def _intro_page(self):
        lay = self._page("The Application Wizard makes the start of a new application for you: its main form, its "
                         "menus and toolbar, and the forms you choose.\n\nYou can go Back at any time to change "
                         "your choices.\n\nClick Next to begin.")
        lay.addStretch(1)
        self.load_combo = self._profile_row(lay, "From what profile do you want to load your settings?", False)
        self.load_combo.activated.connect(lambda _i: self.load_profile(self.load_combo.currentText()))

    def _type_page(self):
        lay = self._page("What type of interface would you like for your application?")
        self.type_group = QButtonGroup(self)
        for t in (G.APP_MDI, G.APP_SDI, G.APP_EXPLORER):
            b = QRadioButton(G.TYPE_NAMES[t])
            self.type_group.addButton(b, t)
            lay.addWidget(b)
        self.type_group.idClicked.connect(self.type_changed)
        hint = QGroupBox("Hint!")
        hl = QVBoxLayout(hint)
        self.hint = QLabel()
        self.hint.setWordWrap(True)
        hl.addWidget(self.hint)
        lay.addWidget(hint)
        lay.addStretch(1)
        lab = QLabel("What &name do you want for the application?")
        lay.addWidget(lab)
        self.name = QLineEdit("Project1")
        lab.setBuddy(self.name)
        lay.addWidget(self.name)

    def _menus_page(self):
        lay = self._page("Choose the menus and sub menus of your application. You can change them later with the "
                         "Menu Editor.")
        row = QHBoxLayout()
        self.menu_list, mbox = self._list_with_buttons("Menus", self.add_menu, self.del_menu,
                                                       lambda d: self.move_menu(d))
        self.sub_list, sbox = self._list_with_buttons("Sub Menus", self.add_sub, self.del_sub,
                                                      lambda d: self.move_sub(d))
        self.sub_buttons = self._last_buttons
        row.addLayout(mbox, 1)
        row.addLayout(sbox, 1)
        lay.addLayout(row, 1)
        b = QPushButton("&Reset")
        b.clicked.connect(self.reset_menus)
        r = QHBoxLayout()
        r.addStretch(1)
        r.addWidget(b)
        lay.addLayout(r)
        self.menu_list.currentRowChanged.connect(lambda _r: self.fill_subs())
        self.menu_list.itemChanged.connect(self.menu_checked)
        self.sub_list.itemChanged.connect(self.sub_checked)

    def _list_with_buttons(self, title, add, delete, move):
        col = QVBoxLayout()
        col.addWidget(QLabel(title + ":"))
        row = QHBoxLayout()
        lst = QListWidget()
        row.addWidget(lst, 1)
        bc = QVBoxLayout()
        self._last_buttons = []
        for text, tip, slot in (("+", "Add", add), ("X", "Delete", delete), ("↑", "Move Up", lambda: move(-1)),
                                ("↓", "Move Down", lambda: move(1))):
            b = QToolButton()
            b.setText(text)
            b.setToolTip(tip)
            b.setFixedSize(26, 24)
            b.clicked.connect(slot)
            bc.addWidget(b)
            self._last_buttons.append(b)
        bc.addStretch(1)
        row.addLayout(bc)
        col.addLayout(row, 1)
        return lst, col

    def _toolbar_page(self):
        lay = self._page("Make the toolbar: move the buttons you want to the list on the right, put them in order "
                         "with the arrows, and add pictures of your own with the picture button. A double click "
                         "moves a button too.")
        self.tbed = ToolbarEditor(self, self.images, reset=self.tb_reset)
        lay.addWidget(self.tbed, 1)

    def _res_page(self):
        lay = self._page("A resource file keeps the strings of the application apart from its code, so the "
                         "application can be translated to other languages without changing the code.\n\n"
                         "The Resource Editor keeps the resource file in memory until you save it; once saved, it "
                         "is the resource file of the project (Related Documents).")
        lay.addStretch(1)
        lay.addWidget(QLabel("Would you like to use a resource file for the strings in your application?"))
        self.res_group = QButtonGroup(self)
        for i, t in ((1, "&Yes"), (0, "N&o")):
            b = QRadioButton(t)
            self.res_group.addButton(b, i)
            lay.addWidget(b)
        lay.addStretch(1)

    def _internet_page(self):
        lay = self._page("The wizard can add a web browser of its own to your application, with a page that it "
                         "opens first.\n\nNote: the web browser of LinBASIC Studio shows HTML and CSS pages, without "
                         "JavaScript.")
        lay.addStretch(1)
        lay.addWidget(QLabel("Do you want your users to be able to access the Internet from your application?"))
        self.net_group = QButtonGroup(self)
        for i, t in ((1, "&Yes"), (0, "N&o")):
            b = QRadioButton(t)
            self.net_group.addButton(b, i)
            lay.addWidget(b)
        t = QLabel("With Yes, give the page the web browser opens first (for example http://www.example.com):")
        t.setWordWrap(True)
        lay.addWidget(t)
        self.url = QLineEdit(G.DEFAULT_URL)
        lay.addWidget(self.url)
        self.net_group.idClicked.connect(lambda i: self.url.setEnabled(i == 1))
        lay.addStretch(1)

    def _forms_page(self):
        lay = self._page("Would you like to include any of these standard forms in your application?")
        self.form_checks = {}
        for key, text in (("Splash", "&Splash screen at application start up"),
                          ("Login", "&Login dialog to accept an ID and Password"),
                          ("Options", "&Options dialog for custom settings"), ("About", "&About Box")):
            c = QCheckBox(text)
            self.form_checks[key] = c
            lay.addWidget(c)
        lay.addStretch(1)
        row = QHBoxLayout()
        row.addWidget(QLabel("To add form templates too, click here."))
        b = QPushButton("Form &Templates...")
        b.clicked.connect(self.form_templates)
        row.addWidget(b)
        row.addStretch(1)
        lay.addLayout(row)
        self.templates_label = QLabel()
        self.templates_label.setWordWrap(True)
        lay.addWidget(self.templates_label)

    def _data_page(self):
        lay = self._page("Would you like to generate forms based upon tables and queries from your database(s)?")
        lay.addWidget(QLabel("Data Forms:"))
        self.data_list = QListWidget()
        lay.addWidget(self.data_list, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        b = QPushButton("&Create New Form...")
        b.setEnabled(False)                     # the Data Form Wizard: phase 9.2b
        b.setToolTip("The Data Form Wizard (a next stage of LinBASIC Studio)")
        row.addWidget(b)
        lay.addLayout(row)

    def _finished_page(self):
        lay = self._page("The Application Wizard has all the information it needs.\n\nClick Finish to build your "
                         "application.")
        lay.addStretch(1)
        self.save_combo = self._profile_row(lay, "To what profile do you want to save your settings?", True)
        lay.addStretch(1)
        row = QHBoxLayout()
        row.addWidget(QLabel("To see a summary report, click View Report."))
        b = QPushButton("&View Report")
        b.clicked.connect(self.view_report)
        row.addWidget(b)
        row.addStretch(1)
        lay.addLayout(row)

    # ------------------------------------------------------------ moving between the pages
    def go(self, i):
        if self.stack.currentIndex() in (1, 4, 5, 6) and i != self.stack.currentIndex():
            self.read_simple()
        i = max(0, min(i, len(PAGES) - 1))
        self.stack.setCurrentIndex(i)
        self.setWindowTitle(f"{TITLE} - {PAGES[i]}")
        kinds = ["intro", {G.APP_MDI: "mdi", G.APP_SDI: "sdi", G.APP_EXPLORER: "explorer"}[self.opts.app_type],
                 "menus", "toolbar", "resources", "internet", "forms", "data", "finished"]
        self.picture.setPixmap(wizard_picture(kinds[i]))
        self.b_back.setEnabled(i > 0)
        self.b_next.setEnabled(i < len(PAGES) - 1)
        self.b_finish.setEnabled(i > 0)         # VB6 (aw2): Finish from the Interface Type page on
        (self.b_finish if i == len(PAGES) - 1 else self.b_next).setDefault(True)
        if i == 2:
            self.fill_menus()
        if i == 3:
            self.fill_toolbar()

    def read_simple(self):
        """the choices of the pages with buttons, into opts"""
        o = self.opts
        o.name = re.sub(r"[^A-Za-z0-9_]", "", self.name.text()) or "Project1"
        o.resources = self.res_group.checkedId() == 1
        o.internet = self.net_group.checkedId() == 1
        o.url = self.url.text().strip() or G.DEFAULT_URL
        o.forms = [k for k in G.STANDARD_FORMS if self.form_checks[k].isChecked()]

    def load_options(self):
        """the controls of the pages from opts (a new wizard, a profile)"""
        o = self.opts
        self.type_group.button(o.app_type).setChecked(True)
        self.hint.setText(HINTS[o.app_type])
        self.name.setText(o.name)
        self.res_group.button(1 if o.resources else 0).setChecked(True)
        self.net_group.button(1 if o.internet else 0).setChecked(True)
        self.url.setText(o.url)
        self.url.setEnabled(o.internet)
        for k, c in self.form_checks.items():
            c.setChecked(k in o.forms)
        self.show_templates()

    def type_changed(self, t):
        self.opts.set_type(t)
        self.hint.setText(HINTS[t])
        self.picture.setPixmap(wizard_picture({G.APP_MDI: "mdi", G.APP_SDI: "sdi", G.APP_EXPLORER: "explorer"}[t]))

    # ------------------------------------------------------------ profiles
    def pick_profile(self, combo, save):
        if save:
            path, _ = QFileDialog.getSaveFileName(self, "Save Profile", os.path.join(profiles_dir(), "profile.rwp"),
                                                  PROFILE_FILTER)
            if path and not path.lower().endswith(".rwp"):
                path += ".rwp"
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Open Profile", profiles_dir(), PROFILE_FILTER)
            if path and not os.path.isfile(path):
                QMessageBox.warning(self, "Open Profile",
                                    f"{os.path.basename(path)}\nFile not found.\nCheck the file name and try again.")
                return
        if not path:
            return
        if combo.findText(path) < 0:
            combo.addItem(path)
        combo.setCurrentIndex(combo.findText(path))
        if not save:
            self.load_profile(path)

    def load_profile(self, path):
        if not path or path == NONE:
            return
        try:
            with open(path, encoding="cp1252") as f:
                o = G.AppOptions.from_rwp(f.read())
        except OSError as e:
            QMessageBox.warning(self, "Open Profile", f"{os.path.basename(path)}\n{e.strerror}")
            return
        o.name = self.opts.name
        self.opts = o
        self.load_options()

    # ------------------------------------------------------------ Menus
    def fill_menus(self):
        cur = max(0, self.menu_list.currentRow())
        self.menu_list.blockSignals(True)
        self.menu_list.clear()
        for m in self.opts.menus:
            it = QListWidgetItem(m.caption)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Checked if m.on else Qt.CheckState.Unchecked)
            self.menu_list.addItem(it)
        self.menu_list.blockSignals(False)
        self.menu_list.setCurrentRow(min(cur, self.menu_list.count() - 1))
        self.fill_subs()

    def current_menu(self):
        r = self.menu_list.currentRow()
        return self.opts.menus[r] if 0 <= r < len(self.opts.menus) else None

    def fill_subs(self):
        m = self.current_menu()
        cur = max(0, self.sub_list.currentRow())
        self.sub_list.blockSignals(True)
        self.sub_list.clear()
        if m is not None:
            for s in m.items:
                it = QListWidgetItem(s.caption if s.key != G.MRU else "MRU List")
                it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                it.setCheckState(Qt.CheckState.Checked if s.on else Qt.CheckState.Unchecked)
                if not m.on:                    # an unchecked menu: its sub menus gray (aw3: &Tools)
                    it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEnabled)
                self.sub_list.addItem(it)
        self.sub_list.blockSignals(False)
        self.sub_list.setCurrentRow(min(cur, self.sub_list.count() - 1))
        for b in self.sub_buttons:              # an unchecked menu: its buttons gray too (aw3: &Tools)
            b.setEnabled(m is not None and m.on)

    def menu_checked(self, it):
        r = self.menu_list.row(it)
        if 0 <= r < len(self.opts.menus):
            self.opts.menus[r].on = it.checkState() == Qt.CheckState.Checked
            self.fill_subs()

    def sub_checked(self, it):
        m = self.current_menu()
        r = self.sub_list.row(it)
        if m is not None and 0 <= r < len(m.items):
            m.items[r].on = it.checkState() == Qt.CheckState.Checked

    def _ask_caption(self, title):
        text, ok = QInputDialog.getText(self, title, "Caption (& before the access key; [Separator] for a line):")
        return text.strip() if ok and text.strip() else None

    def add_menu(self):
        cap = self._ask_caption("Add Menu")
        if cap is None:
            return
        keys = {m.key.lower() for m in self.opts.menus}
        key, n = key_of(cap), 1
        while key.lower() in keys:
            n += 1
            key = f"{key_of(cap)}{n}"
        r = self.menu_list.currentRow() + 1
        self.opts.menus.insert(r, G.Menu(cap, key, True, []))
        self.fill_menus()
        self.menu_list.setCurrentRow(r)

    def del_menu(self):
        r = self.menu_list.currentRow()
        if 0 <= r < len(self.opts.menus):
            del self.opts.menus[r]
            self.fill_menus()

    def move_menu(self, d):
        r = self.menu_list.currentRow()
        if 0 <= r < len(self.opts.menus) and 0 <= r + d < len(self.opts.menus):
            ms = self.opts.menus
            ms[r], ms[r + d] = ms[r + d], ms[r]
            self.fill_menus()
            self.menu_list.setCurrentRow(r + d)

    def add_sub(self):
        m = self.current_menu()
        if m is None:
            return
        cap = self._ask_caption("Add Sub Menu")
        if cap is None:
            return
        if cap.lower() in ("[separator]", "-"):
            item = G.MenuItem(G.SEPARATOR, G.SEPARATOR)
        else:
            keys = {s.key.lower() for s in m.items}
            key, n = key_of(cap), 1
            while key.lower() in keys:
                n += 1
                key = f"{key_of(cap)}{n}"
            item = G.MenuItem(cap, key)
        r = self.sub_list.currentRow() + 1
        m.items.insert(r, item)
        self.fill_subs()
        self.sub_list.setCurrentRow(r)

    def del_sub(self):
        m = self.current_menu()
        r = self.sub_list.currentRow()
        if m is not None and 0 <= r < len(m.items):
            del m.items[r]
            self.fill_subs()

    def move_sub(self, d):
        m = self.current_menu()
        r = self.sub_list.currentRow()
        if m is not None and 0 <= r < len(m.items) and 0 <= r + d < len(m.items):
            m.items[r], m.items[r + d] = m.items[r + d], m.items[r]
            self.fill_subs()
            self.sub_list.setCurrentRow(r + d)

    def reset_menus(self):
        self.opts.menus = G.default_menus(self.opts.app_type)
        self.fill_menus()

    # ------------------------------------------------------------ Customize Toolbar
    def fill_toolbar(self):
        self.tbed.set_buttons(self.opts.toolbar)

    def tb_reset(self):
        self.opts.toolbar = G.default_toolbar(self.opts.app_type)
        self.fill_toolbar()

    # ------------------------------------------------------------ Standard Forms, report
    def form_templates(self):
        dlg = FormTemplatesDialog(self, self.opts.templates)
        if dlg.exec():
            self.opts.templates = dlg.chosen()
            self.show_templates()

    def show_templates(self):
        t = self.opts.templates
        self.templates_label.setText(("Form templates: " + ", ".join(t)) if t else "")

    def view_report(self):
        self.read_simple()
        ReportDialog(self, REPORT + G.summary(self.opts)).exec()

    # ------------------------------------------------------------ Finish
    def make_pic(self, name, size):
        from . import tbicons
        if name in self.images and size == 16:
            return self.images[name]
        return tbicons.bmp(name, size)

    def finish(self):
        self.read_simple()
        main = self.main
        if not main.confirm_discard() or not main.res_confirm():
            return
        gen = G.generate(self.opts)
        project = G.build_project(gen, self.opts, self.make_pic, main.codepage())
        main.close_all_code()
        main.group = None
        main.project = project
        main.res_rebind()
        main.fill_tree()
        main.update_title()
        main.immediate.clear()
        res = G.result_res(gen)
        if res is not None:                     # the String Table: in the Resource Editor, not saved (aw14)
            main.res_loaded = True
            main.apply_addins()
            ed = main.resource_editor()
            ed.bind(project)
            ed.res = res
            ed.set_modified(True)
            ed.refresh()
            main.show_res_editor()
        main.open_designer(project.module("frmMain"))     # the active window: the main form
        path = self.save_combo.currentText()
        if path and path != NONE:
            try:
                with open(path, "w", encoding="cp1252", newline="") as f:
                    f.write(self.opts.to_rwp())
            except OSError as e:
                QMessageBox.warning(self, "Save Profile", f"{path}\n{e.strerror}")
        self.created = project
        self.accept()
        if not main.S["appwiz_no_created"]:
            CreatedDialog(main).exec()


class CreatedDialog(QDialog):
    """after Finish (aw10-01): "The Application has been created." and "Don't show this dialog in the future." """

    def __init__(self, main):
        super().__init__(main)
        self.main = main
        self.setWindowTitle("Application Created")
        lay = QVBoxLayout(self)
        row = QHBoxLayout()
        pic = QLabel()
        pic.setPixmap(wizard_picture("intro", 48, 48))
        row.addWidget(pic, 0, Qt.AlignmentFlag.AlignTop)
        row.addWidget(QLabel("The application has been created."), 1)
        lay.addLayout(row)
        self.again = QCheckBox("Don't show this dialog in the future.")
        lay.addWidget(self.again)
        b = QPushButton("OK")
        b.setDefault(True)
        b.clicked.connect(self.accept)
        r = QHBoxLayout()
        r.addStretch(1)
        r.addWidget(b)
        lay.addLayout(r)

    def accept(self):
        if self.again.isChecked():
            self.main.S["appwiz_no_created"] = True
            self.main.S.save()
        super().accept()


def open_wizard(main):
    """File > New Project > VB Application Wizard, Add-Ins > Application Wizard..."""
    dlg = AppWizard(main)
    main._appwizard = dlg                       # the tests reach it
    dlg.exec()
    return dlg.created


__all__ = ["AppWizard", "ToolbarEditor", "open_wizard", "wizard_icon", "toolbar_wizard_icon", "wizard_picture", "FormTemplatesDialog", "ReportDialog",
           "CreatedDialog"]
