"""ocxpages.py - the property pages of the ActiveX controls at design time (phase 11): Shift+F4, (Custom) in the
Properties window and Properties in the menu of the control.

VB6 shows the pages a control names (ISpecifyPropertyPages) in the OLE property frame: a tab for each page, OK,
Cancel, Apply and Help. The pages here are LinBASIC Studio's own, laid out as VB6 lays them out: the oracle
VB-TESTS/DEV-11/ppcap opened the pages of every control in VB6 and recorded each box (its kind, place, size, text and
the items of its lists); data/ocxpages.json keeps that (Claude Enviroment/tools/ppspec.py makes it). The stock pages
Font, Color and Picture are the ones of ppdesign.PagesFrame. A page whose boxes are properties of the control (a label
"MousePointer:" and its combo box, a check box "Enabled") works by itself: each box shows the value of the property
of its name and Apply / OK set it as the Properties window sets it. A page that does more (the Buttons of a Toolbar,
the Panels of a StatusBar ...) has a handler of its own (HANDLERS).
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
import json
import os
import re

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (QWidget, QLabel, QPushButton, QCheckBox, QRadioButton, QComboBox, QLineEdit, QListWidget,
                             QGroupBox, QFrame, QTabWidget, QPlainTextEdit, QScrollBar, QTableWidget, QDateTimeEdit,
                             QToolButton, QAbstractItemView)

from . import vbprops as V
from .ppdesign import PagesFrame

_DATA = None
HANDLERS = {}                                   # page key -> handler(dialog, key, page widget, [(item, widget)])
STOCK = ("StdFont", "StdColor", "StdPicture")
# the pages of lists (the Buttons of a Toolbar ...): their boxes are not properties of the control; until a handler
# of their own is written they show what VB6 shows and set nothing
LIST_PAGES = set()                              # (all have their code now: stages 2c, 2d)
# the boxes whose text is not the name of their property
ALIASES = {("SSTab", "Tab Count"): "Tabs", ("SSTab", "Current Tab"): "Tab", ("SSTab", "Orientation"): "TabOrientation",
           ("TreeView", "Path Separator"): "PathSeparator", ("ImageList", "Height"): "ImageHeight",
           ("ImageList", "Width"): "ImageWidth", ("ListView", "Normal"): "Icons", ("ListView", "Small"): "SmallIcons",
           ("ListView", "Column Header"): "ColumnHeaderIcons", ("CommonDialog", "Bold"): "FontBold",
           ("CommonDialog", "Italic"): "FontItalic", ("CommonDialog", "Underline"): "FontUnderline",
           ("CommonDialog", "Strikethru"): "FontStrikethru", ("MSHFlexGrid", "Columns"): "Cols",
           ("MSFlexGrid", "Format"): "FormatString", ("MSHFlexGrid", "Format"): "FormatString",
           ("MSRDC", "DataSource"): "DataSourceName", ("RichTextBox", "FileName: [load from]"): "FileName"}
# a box VB6 enables only while a check box of the page is checked: (type, property) -> the property of the check box
ENABLED_BY = {("MonthView", "MaxSelCount"): "MultiSelect", ("UpDown", "SyncBuddy"): "AutoBuddy"}


def data():
    """data/ocxpages.json: {"controls": {type: {"pages", "frame", "props"}}, "pages": {key: {"title", "size", "items"}}}"""
    global _DATA
    if _DATA is None:
        try:
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "ocxpages.json"),
                      encoding="utf-8") as f:
                _DATA = json.load(f)
        except (OSError, ValueError):
            _DATA = {"controls": {}, "pages": {}}
    return _DATA


def pages_of(type_name):
    """the pages of a control of this type (VB6), or None"""
    return data()["controls"].get(type_name) if type_name else None


def dialog_font():
    """MS Sans Serif 8 (11 pixels) as VB6's pages have it; where it is missing a sans serif of the same width"""
    from PyQt6.QtGui import QFontDatabase
    fams = {f.lower(): f for f in QFontDatabase.families()}
    name = next((fams[n] for n in ("ms sans serif", "microsoft sans serif", "liberation sans", "arimo", "arial",
                                   "helvetica", "nimbus sans") if n in fams), None)
    f = QFont(name) if name else QFont()
    f.setPixelSize(11)
    return f


def plain(text):
    """the text of a box without its & and its colon: "&MousePointer:" -> "MousePointer\""""
    return text.replace("&&", "\0").replace("&", "").replace("\0", "&").strip().rstrip(":").strip()


def enum_value(text):
    """the number of an item "3 - lvwReport" (None for other items)"""
    m = re.match(r"^\s*(-?\d+)\s+-\s", text)
    return int(m.group(1)) if m else None


class UpDown(QWidget):
    """the arrows of an up-down control (msctls_updown32): they change the number of the box on their left"""

    def __init__(self, parent):
        super().__init__(parent)
        self.buddy = None
        self.up, self.down = QToolButton(self), QToolButton(self)
        self.up.setArrowType(Qt.ArrowType.UpArrow)
        self.down.setArrowType(Qt.ArrowType.DownArrow)
        self.up.clicked.connect(lambda: self.step(1))
        self.down.clicked.connect(lambda: self.step(-1))

    def resizeEvent(self, e):
        h = self.height() // 2
        self.up.setGeometry(0, 0, self.width(), h)
        self.down.setGeometry(0, h, self.width(), self.height() - h)

    def step(self, d):
        if isinstance(self.buddy, QLineEdit) and self.buddy.isEnabled():
            try:
                n = int(self.buddy.text() or "0") + d
            except ValueError:
                return
            self.buddy.setText(str(max(0, n)))
            self.buddy.textEdited.emit(self.buddy.text())


def kind_of(it):
    """the kind of a recorded box: label, frame, group, check, radio, push, combo, edit, list, table, updown, scroll,
    date, or None (not shown)"""
    c, st = it["cls"], it.get("style", 0)
    if c == "Static":
        k = st & 0x1F
        if k in (4, 5, 6, 7, 8, 9, 0x10, 0x11, 0x12):
            return "frame"
        return "label"
    if c in ("Button", "VCIBitmapButton"):
        k = st & 0xF
        return {2: "check", 3: "check", 5: "check", 6: "check", 4: "radio", 9: "radio", 7: "group"}.get(k, "push")
    if c == "ThunderRT6CommandButton":
        return "push"
    if c == "ThunderRT6CheckBox":
        return "check"
    if c == "ThunderRT6Frame":
        return "group"
    if c in ("ComboBox", "ThunderRT6ComboBox", "VCIColorCombo", "VCIPatternCombo"):
        return "combo"
    if c in ("Edit", "ThunderRT6TextBox"):
        return "edit"
    if c == "ListBox":
        return "list"
    if c == "SysListView32":
        return "table"
    if c in ("msctls_updown32", "msvb_lib_updown"):
        return "updown"
    if c == "ScrollBar":
        return "scroll"
    if c == "msvb_lib_datetimepicker":
        return "date"
    return None


class OcxPagesDialog(PagesFrame):
    """the Property Pages of an ActiveX control on a form"""

    def __init__(self, main, designer, ctl):
        super().__init__(main, designer, ctl)
        info = pages_of(ctl.type_name) or {"pages": [], "props": {}}
        self.info = info
        self.stock_props = dict(info.get("props", {}))
        pages = data()["pages"]
        self.page_keys = [k for k in info["pages"] if k in pages]
        self.widgets = {}                       # page key -> [(recorded box, widget)]
        self.bound = {}                         # page key -> {property: widget}
        self.unbound = {}                       # page key -> [the texts of the boxes that set nothing yet]
        self.setFont(dialog_font())
        pw = max([pages[k]["size"][0] for k in self.page_keys] or [393])
        ph = max([pages[k]["size"][1] for k in self.page_keys] or [245])
        self.page_size = (pw, ph)
        self.tabs = QTabWidget(self)
        self.tabs.setGeometry(6, 7, pw + 8, ph + 26)
        self.tabs.setUsesScrollButtons(False)
        self.tabs.tabBar().setStyleSheet("QTabBar::tab { padding: 2px 5px; min-width: 0px; }")   # VB6: one row
        for key in self.page_keys:
            self.tabs.addTab(self._page(key, pages[key], (pw, ph)), pages[key]["title"])
        fr = info.get("frame") or [0, 0]
        extra = fr[1] - 29 - (ph + 68)          # VB6: its tabs in two rows (MSChart)
        if extra > 10:
            self._rows(pw, extra)
            ph += extra
        cw, y = pw + 20, ph + 39
        self.b_ok, self.b_cancel = QPushButton("OK", self), QPushButton("Cancel", self)
        self.b_apply, self.b_help = QPushButton("&Apply", self), QPushButton("Help", self)
        for i, b in enumerate((self.b_ok, self.b_cancel, self.b_apply, self.b_help)):
            b.setGeometry(cw - 6 - (4 - i) * 75 - (3 - i) * 6, y, 75, 23)
            b.setAutoDefault(False)
        self.b_ok.setDefault(True)
        self.b_ok.clicked.connect(self.ok)
        self.b_cancel.clicked.connect(self.reject)
        self.b_apply.clicked.connect(self.apply)
        self.b_apply.setEnabled(False)
        self.b_help.clicked.connect(self._help)
        self.setFixedSize(cw, ph + 68)

    def _rows(self, pw, rowh):
        """the tabs in two rows, as Windows shows a tab control too narrow for them: the row of the tab chosen next
        to the page, the tabs as wide as the row"""
        from PyQt6.QtWidgets import QTabBar
        self.tabs.tabBar().hide()
        titles = [self.tabs.tabText(i) for i in range(self.tabs.count())]
        first = list(range((len(titles) + 2) // 2))     # VB6 (MSChart): 5 tabs in the first row, 3 in the other
        groups = [first, [i for i in range(len(titles)) if i not in first]]
        self.row_bars = []
        for g in groups:
            bar = QTabBar(self)
            bar.setExpanding(True)
            bar.setDrawBase(False)
            bar.setUsesScrollButtons(False)
            bar.setStyleSheet("QTabBar::tab { padding: 2px 5px; min-width: 0px; }")
            for i in g:
                bar.addTab(titles[i])
            bar.indexes = g
            bar.tabBarClicked.connect(lambda k, g=g: self.tabs.setCurrentIndex(g[k]))
            self.row_bars.append(bar)
        g = self.tabs.geometry()
        hh = self.tabs.tabBar().sizeHint().height()
        self._row_top, self._rowh = g.y(), rowh
        self.tabs.setGeometry(g.x(), g.y() + rowh + hh - 2, g.width(), g.height() - hh + 2)

        def place(cur):
            low = next(b for b in self.row_bars if cur in b.indexes)
            high = next(b for b in self.row_bars if b is not low)
            high.setGeometry(self.tabs.x(), self._row_top, self.tabs.width(), hh)
            low.setGeometry(self.tabs.x(), self._row_top + self._rowh, self.tabs.width(), hh)
            for b in self.row_bars:
                b.blockSignals(True)
                b.setCurrentIndex(b.indexes.index(cur) if cur in b.indexes else -1)
                b.blockSignals(False)
            low.raise_()
        self.tabs.currentChanged.connect(place)
        place(self.tabs.currentIndex())

    def _help(self):
        h = getattr(self.main, "context_help", None)
        if callable(h):
            h(self.type_name)

    # ------------------------------------------------------------ the pages
    def _page(self, key, spec, size):
        if key == "StdFont":
            return self._font_page(size)
        if key == "StdColor":
            return self._color_page(size)
        if key == "StdPicture":
            return self._picture_page(size)
        w = QWidget()
        w.setFixedSize(*size)
        made = []
        items = spec["items"]
        order = sorted(range(len(items)), key=lambda i: 0 if kind_of(items[i]) in ("group", "frame") else 1)
        for i in order:                         # the frames first: the boxes come over them
            it = items[i]
            k = kind_of(it)
            if k is None:
                continue
            wd = self._widget(w, k, it)
            wd.setGeometry(it["x"], it["y"], it["w"], it["h"])
            if it.get("disabled"):
                wd.setEnabled(False)
            made.append((it, wd))
        made.sort(key=lambda m: items.index(m[0]))
        self._fit(made, size)
        self._buddies(made)
        self.widgets[key] = made
        h = HANDLERS.get(key) or BY_TITLE.get((self.type_name, spec["title"]))
        if h is not None:
            h(self, key, w, made)
        elif (self.type_name, spec["title"]) in LIST_PAGES or (self.type_name, "*") in LIST_PAGES:
            self.bound[key] = {}
            self.unbound[key] = ["(the page of a list: not yet)"]
        else:
            self.auto_bind(key, made)
        return w

    @staticmethod
    def _widget(parent, k, it):
        text = it.get("text", "")
        if k == "label":
            wd = QLabel(text, parent)
            a = it.get("style", 0) & 3
            wd.setAlignment((Qt.AlignmentFlag.AlignHCenter if a == 1 else Qt.AlignmentFlag.AlignRight if a == 2
                             else Qt.AlignmentFlag.AlignLeft) | Qt.AlignmentFlag.AlignTop)
            if it.get("style", 0) & 0x1000:     # SS_SUNKEN
                wd.setFrameShape(QFrame.Shape.Panel)
                wd.setFrameShadow(QFrame.Shadow.Sunken)
            wd.setWordWrap(it["h"] > 20)
            return wd
        if k == "frame":
            wd = QFrame(parent)
            wd.setFrameShape(QFrame.Shape.HLine if it["h"] <= 3 else QFrame.Shape.Box)
            wd.setFrameShadow(QFrame.Shadow.Sunken)
            return wd
        if k == "group":
            return QGroupBox(text.replace("&", ""), parent)
        if k == "check":
            wd = QCheckBox(text, parent)
            wd.setChecked(bool(it.get("check")))
            if it.get("style", 0) & 0x20:       # BS_LEFTTEXT: the box on the right of its text
                wd.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
            return wd
        if k == "radio":
            wd = QRadioButton(text, parent)
            wd.setChecked(bool(it.get("check")))
            wd.setAutoExclusive(False)          # the groups of VB6's pages are their own: the handlers keep them
            return wd
        if k == "push":
            return QPushButton(text, parent)
        if k == "combo":
            wd = QComboBox(parent)
            if it.get("style", 3) & 3 in (1, 2):
                wd.setEditable(True)
            wd.addItems(it.get("items", []))
            if it.get("sel", -1) >= 0:
                wd.setCurrentIndex(it["sel"])
            elif wd.isEditable():
                wd.setEditText(text)
            return wd
        if k == "edit":
            if it.get("style", 0) & 4:          # ES_MULTILINE
                wd = QPlainTextEdit(parent)
                wd.setPlainText(text)
                return wd
            wd = QLineEdit(text, parent)
            if it.get("style", 0) & 0x800:      # ES_READONLY: grey, as Windows draws it
                wd.setReadOnly(True)
                gray(wd)
            return wd
        if k == "list":
            wd = QListWidget(parent)
            wd.addItems(it.get("items", []))
            if it.get("sel", -1) >= 0:
                wd.setCurrentRow(it["sel"])
            return wd
        if k == "table":
            wd = QTableWidget(parent)
            wd.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
            return wd
        if k == "updown":
            return UpDown(parent)
        if k == "scroll":
            return QScrollBar(Qt.Orientation.Horizontal if it["w"] > it["h"] else Qt.Orientation.Vertical, parent)
        if k == "date":
            return QDateTimeEdit(parent)
        return QWidget(parent)

    @staticmethod
    def _fit(made, size):
        """the texts of labels and check boxes in the free room beside them: the fonts of Linux are a little wider than
        MS Sans Serif, and VB6's boxes are only as wide as their texts there"""
        boxes = [(it, wd) for it, wd in made if not isinstance(wd, (QGroupBox, QFrame)) or isinstance(wd, QLabel)]
        for it, wd in made:
            if not isinstance(wd, (QLabel, QCheckBox, QRadioButton)) or wd.sizeHint().width() <= it["w"]:
                continue
            y0, y1 = it["y"], it["y"] + it["h"]
            row = [jt for jt, w2 in boxes if jt is not it and jt["y"] < y1 and jt["y"] + jt["h"] > y0]
            right_aligned = isinstance(wd, QLabel) and it.get("style", 0) & 3 == 2 or \
                isinstance(wd, QCheckBox) and it.get("style", 0) & 0x20
            if right_aligned:
                left = max([jt["x"] + jt["w"] for jt in row if jt["x"] + jt["w"] <= it["x"]] + [0]) + 2
                x1 = it["x"] + it["w"]
                wd.setGeometry(max(left, x1 - wd.sizeHint().width()), it["y"], x1 - max(left, x1 - wd.sizeHint().width()),
                               it["h"])
            else:
                right = min([jt["x"] for jt in row if jt["x"] >= it["x"] + it["w"] - 1] + [size[0]]) - 2
                wd.setGeometry(it["x"], it["y"], max(it["w"], min(wd.sizeHint().width(), right - it["x"])), it["h"])
            f = wd.font()                       # still too narrow: a smaller font, the whole text
            if f.pixelSize() <= 0:              # (a font in points: its size in pixels)
                from PyQt6.QtGui import QFontInfo
                f.setPixelSize(QFontInfo(f).pixelSize())
            while wd.sizeHint().width() > wd.width() and f.pixelSize() > 8:
                f.setPixelSize(f.pixelSize() - 1)
                wd.setFont(f)

    @staticmethod
    def _buddies(made):
        """a label with & goes to the box after it; the arrows of an up-down to the box on their left"""
        for i, (it, wd) in enumerate(made):
            if isinstance(wd, QLabel) and "&" in it.get("text", ""):
                nxt = next((w for _, w in made[i + 1:] if not isinstance(w, (QLabel, QFrame, QGroupBox))), None)
                if nxt is not None:
                    wd.setBuddy(nxt)
            if isinstance(wd, UpDown):
                left = [w for jt, w in made if isinstance(w, QLineEdit) and abs(jt["y"] - it["y"]) <= 4
                        and jt["x"] + jt["w"] <= it["x"] + 2]
                if left:
                    wd.buddy = max(left, key=lambda w: w.geometry().x())

    # ------------------------------------------------------------ the boxes that are properties
    def pairs(self, made):
        """(property name, recorded box, widget) of a page: a label and the combo / edit box beside it, a check box"""
        out, used = [], set()
        boxes = [(jt, w2) for jt, w2 in made if isinstance(w2, (QComboBox, QLineEdit, QDateTimeEdit, QPlainTextEdit))]
        for it, wd in made:
            if isinstance(wd, QLabel) and plain(it.get("text", "")):
                cy = it["y"] + it["h"] / 2
                right = [(jt["x"], jt, w2) for jt, w2 in boxes if id(w2) not in used and jt["x"] > it["x"] and
                         jt["y"] - 4 <= cy <= jt["y"] + jt["h"] + 4]
                under = [(jt["y"], jt, w2) for jt, w2 in boxes if id(w2) not in used and
                         0 < jt["y"] - it["y"] <= 24 and abs(jt["x"] - it["x"]) <= 4]
                pick = min(right, key=lambda t: t[0]) if right else min(under, key=lambda t: t[0]) if under else None
                if pick is not None:
                    used.add(id(pick[2]))
                    out.append((plain(it["text"]), pick[1], pick[2]))
            elif isinstance(wd, QCheckBox):
                out.append((plain(it["text"]), it, wd))
        return out

    def find_prop(self, name):
        name = ALIASES.get((self.type_name, name), name)
        for n in (name, name.replace(" ", "")):
            p = self.prop(n)
            if p is not None:
                return p
        return None

    def auto_bind(self, key, made):
        """each box whose name is a property of the control shows its value and changes it"""
        bound, used = {}, set()
        for name, it, wd in self.pairs(made):
            p = self.find_prop(name)
            if p is None:
                continue
            if self.bind(p, wd):
                bound[p.name] = wd
                used.add(id(wd))
        for (t, name), by in ENABLED_BY.items():
            if t == self.type_name and name in bound and by in bound:
                chk = bound[by]
                bound[name].setEnabled(chk.isChecked())
                chk.toggled.connect(bound[name].setEnabled)
        self.bound[key] = bound
        self.unbound[key] = [plain(it.get("text", "")) or f"{it['cls']} {it['id']}" for it, wd in made
                             if id(wd) not in used and isinstance(wd, (QComboBox, QLineEdit, QCheckBox, QListWidget,
                                                                       QRadioButton, QPlainTextEdit))]

    def bind(self, p, wd):
        """a box shows the value of property p and changes it; False when it cannot"""
        v = self.pending.get(p.name, self.value(p.name))
        name = p.name
        if isinstance(wd, QCheckBox):
            wd.setChecked(bool(v))
            wd.toggled.connect(lambda on, name=name: self.changed(name, on))
            return True
        if isinstance(wd, QComboBox) and not wd.isEditable() and p.kind == "string" and \
                any(wd.itemText(i) == "<None>" for i in range(wd.count())):
            names = self.image_lists()          # the ImageLists of the form (ImageList, Icons ...)
            wd.clear()
            wd.addItems(["<None>"] + names)
            wd.setCurrentIndex(names.index(v) + 1 if v in names else 0)
            wd.activated.connect(lambda i, name=name, wd=wd: self.changed(name, "" if i == 0 else wd.itemText(i)))
            return True
        if isinstance(wd, QComboBox) and not wd.isEditable():
            vals = [enum_value(wd.itemText(i)) for i in range(wd.count())]
            if p.kind == "bool" and wd.count() == 2 and None in vals:
                vals = [t.lower() == "true" for t in (wd.itemText(0), wd.itemText(1))]
            if all(x is None for x in vals) and p.kind == "string":     # a list of names (DataSource ...)
                if v and wd.findText(str(v)) < 0:
                    wd.addItem(str(v))
                wd.setCurrentIndex(wd.findText(str(v)) if v else -1)
                wd.activated.connect(lambda i, name=name, wd=wd: self.changed(name, wd.itemText(i)))
                return True
            if all(x is None for x in vals):
                return False
            if v is not None:
                cur = next((i for i, x in enumerate(vals) if x == (int(v) if not isinstance(v, bool) else v)), -1)
                wd.setCurrentIndex(cur)
            wd.activated.connect(lambda i, name=name, vals=vals: vals[i] is not None and self.changed(name, vals[i]))
            return True
        if isinstance(wd, QPlainTextEdit):
            if p.kind not in ("string", "text"):
                return False
            wd.setPlainText("" if v is None else str(v))
            wd.textChanged.connect(lambda name=name, wd=wd: self.changed(name, wd.toPlainText()))
            return True
        if isinstance(wd, (QLineEdit, QComboBox)):
            le = wd if isinstance(wd, QLineEdit) else wd.lineEdit()
            if p.kind in ("font", "picture", "list"):
                return False
            le.setText(str(int(v) & 0xFFFFFFFF) if p.kind == "color" and v is not None else self.shown(p, v))
            le.textEdited.connect(lambda s, name=name: self.changed(name, s))
            return True
        return False

    def image_lists(self):
        root = self.designer.root if self.designer is not None else None
        return [c.name for c in root.walk() if c.type_name == "ImageList"] if root is not None else []

    @staticmethod
    def shown(p, v):
        """a value in a box of a page: numbers with the decimal separator of the system, as VB6 shows them"""
        from PyQt6.QtCore import QLocale
        if v is None:
            return ""
        if p.kind in ("num", "pos", "scale") and isinstance(v, float):
            t = f"{v:g}" if v != int(v) else str(int(v))
            return t.replace(".", QLocale().decimalPoint())
        if p.kind == "bool":
            return "True" if v else "False"
        return str(v)


def gray(wd):
    """a box that is read only: the colour of a button face behind its text (Windows)"""
    from PyQt6.QtGui import QPalette
    pal = wd.palette()
    pal.setColor(QPalette.ColorRole.Base, pal.color(QPalette.ColorRole.Button))
    wd.setPalette(pal)


# ---------------------------------------------------------------- the pages with code of their own
def by_id(made, wid):
    return next((w for it, w in made if it["id"] == wid), None)


def item_of(made, wid):
    return next((it for it, w in made if it["id"] == wid), None)


def _done(dlg, key, names, made=None, ids=()):
    """the boxes a handler has bound (by their names, or their ids) are no longer in the list of what is to do"""
    gone = set(names)
    for it, wd in made or []:
        if it["id"] in ids:
            gone.add(plain(it.get("text", "")) or f"{it['cls']} {it['id']}")
    dlg.unbound[key] = [n for n in dlg.unbound.get(key, []) if n not in gone]


def raw_value(dlg, name, default=None):
    if name in dlg.pending:
        return dlg.pending[name]
    v = dlg.ctl.get(name)
    return default if v is None else v


def set_raw(dlg, name, value, default=None):
    """a value of the .frm that the Properties window does not show, changed on Apply"""
    def put(v, name=name, default=default):
        if default is not None and v == default:
            V.remove(dlg.ctl, name)
        else:
            V.put(dlg.ctl, name, v)
    dlg.raw_apply[name] = put
    dlg.changed(name, value)


def odbc_drivers():
    """the ODBC drivers of this computer (odbcinst.ini of unixODBC)"""
    import configparser
    out = []
    for f in ("/etc/odbcinst.ini", "/usr/local/etc/odbcinst.ini", os.path.expanduser("~/.odbcinst.ini")):
        c = configparser.ConfigParser(strict=False, interpolation=None)
        try:
            c.read(f)
        except (configparser.Error, OSError):
            continue
        out += [n for n in c.sections() if n != "ODBC Drivers" and n not in out]
    return out


def new_dsn_dialog(parent, path=None, answer=None):
    """a new user data source of ODBC (the New... of the Adodc General page): its name, driver, database,
    server, description -> a section of ~/.odbc.ini; the name, or None. answer: (name, driver, database, server)
    for the tests"""
    import configparser
    from PyQt6.QtWidgets import QDialog, QFormLayout, QComboBox, QDialogButtonBox
    path = path or os.path.expanduser("~/.odbc.ini")
    if answer is None:
        d = QDialog(parent)
        d.setWindowTitle("Create New Data Source")
        f = QFormLayout(d)
        name, drv, db, srv, desc = QLineEdit(), QComboBox(), QLineEdit(), QLineEdit(), QLineEdit()
        drv.setEditable(True)
        drv.addItems(odbc_drivers())
        for t, wd in (("Data Source &Name:", name), ("&Driver:", drv), ("Data&base:", db), ("&Server:", srv),
                      ("D&escription:", desc)):
            f.addRow(t, wd)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(d.accept)
        bb.rejected.connect(d.reject)
        f.addRow(bb)
        if not d.exec() or not name.text().strip():
            return None
        answer = (name.text().strip(), drv.currentText().strip(), db.text().strip(), srv.text().strip(),
                  desc.text().strip())
    n, driver, database, server = answer[:4]
    c = configparser.ConfigParser(strict=False, interpolation=None)
    c.optionxform = str
    try:
        c.read(path)
    except (configparser.Error, OSError):
        pass
    if not c.has_section(n):
        c.add_section(n)
    for k, v in (("Description", answer[4] if len(answer) > 4 else ""), ("Driver", driver), ("Database", database),
                 ("Server", server)):
        if v:
            c.set(n, k, v)
    with open(path, "w") as fh:
        c.write(fh)
    return n


def odbc_dsns():
    """the data sources of ODBC of this computer (odbc.ini)"""
    import configparser
    out = []
    for f in (os.path.expanduser("~/.odbc.ini"), "/etc/odbc.ini"):
        c = configparser.ConfigParser(strict=False, interpolation=None)
        try:
            c.read(f)
        except (configparser.Error, OSError):
            continue
        out += [n for n in c.sections() if n != "ODBC Data Sources" and n not in out]
    return out


def _imagelist_general(dlg, key, w, made):
    """16 x 16, 32 x 32, 48 x 48 or Custom: ImageHeight and ImageWidth (VB6: the boxes only for Custom)"""
    dlg.auto_bind(key, made)
    sizes = {232: 16, 233: 32, 236: 48}
    radios = {i: by_id(made, i) for i in (232, 233, 236, 234)}
    he, we = dlg.bound[key].get("ImageHeight"), dlg.bound[key].get("ImageWidth")
    h, wv = int(dlg.value("ImageHeight") or 0), int(dlg.value("ImageWidth") or 0)
    cur = next((i for i, n in sizes.items() if h == wv == n), 234 if h or wv else 232)
    if not h and not wv:
        he.setText("16")
        we.setText("16")

    def pick(i, user=True):
        for j, rb in radios.items():
            rb.setChecked(j == i)
        he.setEnabled(i == 234)
        we.setEnabled(i == 234)
        if user and i != 234:
            n = sizes[i]
            he.setText(str(n))
            we.setText(str(n))
            dlg.changed("ImageHeight", n)
            dlg.changed("ImageWidth", n)
    for i, rb in radios.items():
        rb.clicked.connect(lambda _=False, i=i: pick(i))
    pick(cur, False)
    _done(dlg, key, [], made, (232, 233, 236, 234))


def _sstab_general(dlg, key, w, made):
    """Current Tab with < and >, its TabCaption (TabCaption(i) of the .frm), Tab Count ..."""
    dlg.auto_bind(key, made)
    b = dlg.bound[key]
    cur_e, cap, count = b.get("Tab"), by_id(made, 113), b.get("Tabs")

    def cur():
        try:
            return max(0, int(cur_e.text()))
        except ValueError:
            return 0

    def show():
        i = cur()
        cap.setText(str(raw_value(dlg, f"TabCaption({i})", f"Tab {i}")))

    def move(d):
        try:
            n = int(count.text())
        except ValueError:
            n = 1
        i = min(max(0, cur() + d), max(0, n - 1))
        cur_e.setText(str(i))
        cur_e.textEdited.emit(str(i))
        show()
    cap.textEdited.connect(lambda t: set_raw(dlg, f"TabCaption({cur()})", t))
    cur_e.textEdited.connect(lambda *_: show())
    by_id(made, 111).clicked.connect(lambda: move(-1))
    by_id(made, 112).clicked.connect(lambda: move(1))
    show()
    _done(dlg, key, ["TabCaption"], made, (113,))


OLE_DAY0 = (1899, 12, 30)
DATE_KEYS = {"Value": "CurrentDate", "MinDate": "MinDate", "MaxDate": "MaxDate"}
DATE_DEFAULTS = {("MonthView", "MinDate"): -53688, ("MonthView", "MaxDate"): 2958465,
                 ("DTPicker", "MinDate"): -109205, ("DTPicker", "MaxDate"): 2958465}


def _dates(dlg, key, w, made):
    """Value, MinDate, MaxDate (and the Time of a DTPicker): CurrentDate / MinDate / MaxDate of the .frm, dates of
    OLE (days since 30 December 1899)"""
    from PyQt6.QtCore import QDate, QTime
    dlg.auto_bind(key, made)
    day0 = QDate(*OLE_DAY0)
    for name, it, wd in dlg.pairs(made):
        if not isinstance(wd, QDateTimeEdit):
            continue
        k = DATE_KEYS.get(name)
        if k is None:
            continue
        default = DATE_DEFAULTS.get((dlg.type_name, k), day0.daysTo(QDate.currentDate()))
        v = float(raw_value(dlg, k, default))
        wd.setDisplayFormat("d/M/yyyy")
        wd.setCalendarPopup(True)
        wd.setDate(day0.addDays(int(v)))
        wd.dateChanged.connect(lambda d, k=k, default=default, wd=wd: set_raw(
            dlg, k, day0.daysTo(d) + (float(raw_value(dlg, k, 0)) % 1 if k == "CurrentDate" else 0),
            default if k != "CurrentDate" else None))
        _done(dlg, key, [name])
    t = by_id(made, 2306)                        # the Time of a DTPicker: the fraction of CurrentDate
    if isinstance(t, QDateTimeEdit):
        v = float(raw_value(dlg, "CurrentDate", day0.daysTo(QDate.currentDate())))
        t.setDisplayFormat("HH:mm:ss")
        t.setTime(QTime(0, 0).addSecs(int(round((v % 1) * 86400))))
        t.timeChanged.connect(lambda tm: set_raw(dlg, "CurrentDate", int(float(raw_value(dlg, "CurrentDate", 0)))
                                                 + QTime(0, 0).secsTo(tm) / 86400))
        _done(dlg, key, ["Time"])


def _adodc_general(dlg, key, w, made):
    """the Source of Connection: a Data Link file (ConnectStringType 2, OLEDBFile), an ODBC data source (3,
    DataSourceName) or a connection string (1, OLEDBString = ConnectionString); Other Attributes"""
    from PyQt6.QtWidgets import QFileDialog
    r_file, r_dsn, r_cs = by_id(made, 2203), by_id(made, 2204), by_id(made, 2202)
    e_file, b_browse, c_dsn, b_new = by_id(made, 2206), by_id(made, 2209), by_id(made, 2207), by_id(made, 2210)
    e_cs, b_build, e_other = by_id(made, 2205), by_id(made, 2208), by_id(made, 2212)
    dlg.bind(dlg.prop("ConnectionString"), e_cs)
    e_file.setText(str(raw_value(dlg, "OLEDBFile", "")))
    c_dsn.clear()
    c_dsn.addItems(odbc_dsns())
    c_dsn.setEditText(str(raw_value(dlg, "DataSourceName", "")))
    e_other.setText(str(raw_value(dlg, "OtherAttributes", "")))
    e_file.textEdited.connect(lambda t: set_raw(dlg, "OLEDBFile", t))
    c_dsn.lineEdit().textEdited.connect(lambda t: set_raw(dlg, "DataSourceName", t))
    c_dsn.activated.connect(lambda i: set_raw(dlg, "DataSourceName", c_dsn.itemText(i)))
    e_other.textEdited.connect(lambda t: set_raw(dlg, "OtherAttributes", t))

    def new_dsn():                              # VB6: the ODBC Data Source Administrator; here: a data source of
        name = new_dsn_dialog(dlg)              # unixODBC in ~/.odbc.ini
        if name:
            c_dsn.clear()
            c_dsn.addItems(odbc_dsns())
            c_dsn.setEditText(name)
            set_raw(dlg, "DataSourceName", name)
    b_new.clicked.connect(new_dsn)

    def kind(n, user=True):
        for rb, k in ((r_cs, 1), (r_file, 2), (r_dsn, 3)):
            rb.setChecked(k == n)
        for wd, k in ((e_file, 2), (b_browse, 2), (c_dsn, 3), (b_new, 3), (e_cs, 1), (b_build, 1)):
            wd.setEnabled(k == n)
        if user:
            set_raw(dlg, "ConnectStringType", n)

    def browse():
        path, _ = QFileDialog.getOpenFileName(dlg, "Organize Data Link Files", "",
                                              "Microsoft Data Link Files (*.udl);;All Files (*)")
        if path:
            e_file.setText(path)
            set_raw(dlg, "OLEDBFile", path)

    def build():
        path, _ = QFileDialog.getOpenFileName(dlg, "Data Link Properties - Database", "",
                                              "Access Databases (*.mdb *.MDB);;SQLite Databases (*.db *.sqlite);;"
                                              "All Files (*)")
        if path:
            cs = f"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={path};Persist Security Info=False"
            e_cs.setText(cs)
            dlg.changed("ConnectionString", cs)
    for rb, k in ((r_cs, 1), (r_file, 2), (r_dsn, 3)):
        rb.clicked.connect(lambda _=False, k=k: kind(k))
    b_browse.clicked.connect(browse)
    b_build.clicked.connect(build)
    kind(int(_num(raw_value(dlg, "ConnectStringType", 1))), False)
    dlg.bound[key] = {"ConnectionString": e_cs}
    dlg.unbound[key] = []


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0


def _adodc_recordsource(dlg, key, w, made):
    """Command Type; a table or stored procedure (adCmdTable, adCmdStoredProc) or the text of a command (adCmdText,
    adCmdUnknown): RecordSource"""
    dlg.auto_bind(key, made)
    ctype, table, text = dlg.bound[key].get("CommandType"), by_id(made, 2243), by_id(made, 2242)
    rs = str(dlg.value("RecordSource") or "")

    def tables():
        try:
            from . import dbinfo
            from .dsrdesign import design_cs
            d = dlg.designer
            cs = str(dlg.pending.get("ConnectionString", dlg.value("ConnectionString")) or "")
            q = d.module.project if d is not None else None
            base = os.path.dirname(q.path) if q is not None and q.path else ""
            return [n for _, n in dbinfo.ado_objects(design_cs(cs, base))]
        except Exception:                       # noqa: BLE001 - no database at design time
            return []

    def mode(user=False):
        v = enum_value(ctype.currentText()) or 8
        is_table = v in (2, 4)
        table.setEnabled(is_table)
        text.setEnabled(not is_table)
        if is_table and table.count() == 0:
            table.addItems(tables())
    table.setEditable(True)
    table.setEditText(rs)
    text.setPlainText(rs)
    table.lineEdit().textEdited.connect(lambda t: dlg.changed("RecordSource", t))
    table.activated.connect(lambda i: dlg.changed("RecordSource", table.itemText(i)))
    text.textChanged.connect(lambda: dlg.changed("RecordSource", text.toPlainText()))
    ctype.activated.connect(lambda *_: mode(True))
    mode()
    _done(dlg, key, [], made, (2243, 2242))


def _msrdc_general(dlg, key, w, made):
    """DataSource (the ODBC data sources of this computer), UserName, Password, SQL"""
    c = by_id(made, 2011)
    c.clear()
    c.addItems(odbc_dsns())
    dlg.auto_bind(key, made)


def _coolbar_general(dlg, key, w, made):
    """a page written in VB (its labels have no windows): ImageList, Orientation, Picture, the colours, the checks"""
    from PyQt6.QtWidgets import QColorDialog, QFileDialog
    from PyQt6.QtGui import QColor
    from PyQt6.QtWidgets import QStyle
    labels = {}
    for text, x, y, buddy in (("&ImageList:", 14, 11, 16392), ("&Orientation:", 14, 35, 16393),
                              ("&Picture:", 14, 59, 16391), ("Emboss&Highlight", 128, 100, 16389),
                              ("Emboss&Shadow", 128, 116, 16388), ("&ForeColor:", 128, 156, 16385),
                              ("&BackColor:", 128, 172, 16386)):
        lab = QLabel(text, w)
        lab.setBuddy(by_id(made, buddy))
        lab.setGeometry(x, y, 112, 14)
        lab.show()
        labels[buddy] = lab
    by_id(made, 16391).setIcon(dlg.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
    il, orient = by_id(made, 16392), by_id(made, 16393)
    il.clear()
    names = dlg.image_lists()
    il.addItems(["(None)"] + names)
    v = str(dlg.value("ImageList") or "")
    il.setCurrentIndex(names.index(v) + 1 if v in names else 0)
    il.activated.connect(lambda i: dlg.changed("ImageList", "" if i == 0 else il.itemText(i)))
    orient.clear()
    orient.addItems(["0 - cc3OrientationHorizontal", "1 - cc3OrientationVertical"])
    dlg.bind(dlg.prop("Orientation"), orient)
    pic = QLabel("(None)" if not dlg.value("Picture") else "(Bitmap)", w)
    pic.setFrameShape(QFrame.Shape.Panel)
    pic.setFrameShadow(QFrame.Shadow.Sunken)
    pic.setGeometry(120, 56, 155, 20)
    pic.show()

    def browse():
        path, _ = QFileDialog.getOpenFileName(dlg, "Load Picture", "", "Pictures (*.bmp *.dib *.gif *.jpg *.ico *.cur "
                                              "*.wmf *.emf);;All Files (*)")
        if path:
            with open(path, "rb") as f:
                dlg.changed("Picture", f.read())
            pic.setText("(Bitmap)")
    by_id(made, 16391).clicked.connect(browse)
    for wid, name in ((16389, "EmbossHighlight"), (16388, "EmbossShadow"), (16385, "ForeColor"), (16386, "BackColor")):
        b = by_id(made, wid)

        def paint(b=b, name=name):
            from .propwin import qcolor
            c = dlg.pending.get(name, dlg.value(name))
            b.setStyleSheet(f"background: {qcolor(int(c)).name()}" if c is not None else "")

        def pick(_=False, b=b, name=name, paint=paint):
            c = QColorDialog.getColor(QColor("white"), dlg)
            if c.isValid():
                dlg.changed(name, c.red() | (c.green() << 8) | (c.blue() << 16))
                paint()
        b.clicked.connect(pick)
        paint()
    checks = {}
    for wid, name in ((16390, "EmbossPicture"), (16396, "BandBorders"), (16395, "FixedOrder"), (16394, "VariantHeight")):
        cb = by_id(made, wid)
        dlg.bind(dlg.prop(name), cb)
        checks[name] = cb
    emb = checks["EmbossPicture"]
    for wid in (16389, 16388):
        for wd in (by_id(made, wid), labels[wid]):
            wd.setEnabled(emb.isChecked())
            emb.toggled.connect(wd.setEnabled)
    dlg.bound[key] = dict(checks, ImageList=il, Orientation=orient)
    dlg.unbound[key] = []


def _updown_buddy(dlg, key, w, made):
    """Buddy Control, AutoBuddy, SyncBuddy (only with AutoBuddy); Buddy Property: the properties of the buddy"""
    dlg.auto_bind(key, made)
    prop_c, lab = by_id(made, 2507), by_id(made, 2508)
    buddy = dlg.bound[key].get("BuddyControl")
    sync = dlg.bound[key].get("SyncBuddy")

    from lbvc.symbols import prop_dispids     # (VB6 writes the DISPID of the property of the buddy)

    def fill():
        name = buddy.text().strip() if buddy is not None else ""
        root = dlg.designer.root if dlg.designer is not None else None
        b = next((c for c in root.walk() if c.name.lower() == name.lower()), None) if root is not None and name \
            else None
        props = (["(Default)"] + sorted(p.name for p in V.props_of(b.type_name)
                                        if p.kind in ("string", "text", "int", "num", "bool", "enum")
                                        and p.name in prop_dispids(b.cls))) if b is not None else []
        prop_c.clear()
        prop_c.addItems(props)
        on = b is not None and sync is not None and sync.isChecked()
        prop_c.setEnabled(on)
        lab.setEnabled(on)
        cur = str(dlg.value("BuddyProperty") or "")
        if cur in props:
            prop_c.setCurrentIndex(props.index(cur))
    if buddy is not None:
        buddy.textEdited.connect(lambda *_: fill())
    if sync is not None:
        sync.toggled.connect(lambda *_: fill())
    prop_c.activated.connect(lambda i: dlg.changed("BuddyProperty", prop_c.itemText(i)))
    fill()
    _done(dlg, key, [], made, (2507,))


def _richtext_general(dlg, key, w, made):
    """FileName [load from] and its ...: a file of the computer"""
    from PyQt6.QtWidgets import QFileDialog
    dlg.auto_bind(key, made)
    e = dlg.bound[key].get("FileName")

    def browse():
        path, _ = QFileDialog.getOpenFileName(dlg, "Open", "", "Rich Text Format (*.rtf);;Text Files (*.txt);;"
                                              "All Files (*)")
        if path and e is not None:
            e.setText(path)
            e.textEdited.emit(path)
    by_id(made, 2105).clicked.connect(browse)


BY_TITLE = {("ImageList", "General"): _imagelist_general, ("SSTab", "General"): _sstab_general,
            ("MonthView", "General"): _dates, ("DTPicker", "General"): _dates,
            ("Adodc", "General"): _adodc_general, ("Adodc", "RecordSource"): _adodc_recordsource,
            ("MSRDC", "General"): _msrdc_general, ("CoolBar", "General"): _coolbar_general,
            ("UpDown", "Buddy"): _updown_buddy, ("RichTextBox", "General"): _richtext_general}
from .ocxlists import PAGES as _LISTS          # noqa: E402 - the pages of lists (stage 2c)
BY_TITLE.update(_LISTS)
from .ocxgrids import PAGES as _GRIDS          # noqa: E402 - the pages of the grids (stage 2c)
BY_TITLE.update(_GRIDS)
from .ocxchart import PAGES as _CHART          # noqa: E402 - the pages of an MSChart (stage 2d)
BY_TITLE.update(_CHART)


__all__ = ["OcxPagesDialog", "pages_of", "data", "HANDLERS", "BY_TITLE", "plain", "enum_value"]
