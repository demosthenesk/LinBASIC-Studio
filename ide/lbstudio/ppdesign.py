"""The Property Pages of a user control on a form at design time (phase 9.2e-2): Shift+F4, or Properties in the
menu of the control (VB6 DEV-9/pp1-16..17, pp2-13..15).

VB6 runs the code of the pages at design time; LinBASIC Studio compiles and does not run code then. The pages of
the system are its own here (Font, Color, Picture), and a page of the project shows its Labels, TextBoxes and
CheckBoxes and moves the values between them and the properties of the control when its code follows the lines the
Property Page Wizard writes (txtX.Text = SelectedControls(0).X, chkX.Value = (SelectedControls(0).X And
vbChecked) and back in ApplyChanges); a page whose code does something else says so.
"""
import re

from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QColor, QFont, QPixmap
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton, QLineEdit,
                             QListWidget, QCheckBox, QTabWidget, QWidget, QFontComboBox, QComboBox, QGroupBox,
                             QColorDialog, QFileDialog, QFrame)

from . import vbprops as V
from . import ppwizgen as G

TPP = 15                                        # twips per pixel
STANDARD_TABS = {"StandardFont": "Font", "StandardColor": "Color", "StandardPicture": "Picture"}
_GET_RE = re.compile(r"^\s*(txt|chk)([A-Za-z_]\w*)\.(Text|Value)\s*=\s*\(?\s*SelectedControls\(0\)\.([A-Za-z_]\w*)"
                     r"(\s+And\s+vbChecked\s*\))?\s*$", re.IGNORECASE)
_SET_RE = re.compile(r"^\s*SelectedControls\(0\)\.([A-Za-z_]\w*)\s*=\s*(?:\(\s*)?(txt|chk)([A-Za-z_]\w*)\.(Text|Value)"
                     r"(\s*=\s*vbChecked\s*\))?\s*$", re.IGNORECASE)
_PROC_RE = re.compile(r"^\s*(Public\s+|Private\s+)?Sub\s+([A-Za-z_]\w*)\s*\(", re.IGNORECASE)
# the palettes of the Color page (the strings of msstkprp.dll, DEV-11): Standard Colors, Windows System Colors
COLORS = [("Black", 0x000000), ("White", 0xFFFFFF), ("Red", 0x0000FF), ("Green", 0x00FF00), ("Blue", 0xFF0000),
          ("Yellow", 0x00FFFF), ("Magenta", 0xFF00FF), ("Cyan", 0xFFFF00), ("Dark Gray", 0x808080),
          ("Gray", 0xC0C0C0), ("Dark Red", 0x000080), ("Dark Green", 0x008000), ("Dark Blue", 0x800000),
          ("Dark Yellow", 0x008080), ("Dark Magenta", 0x800080), ("Dark Cyan", 0x808000)]
SYSTEM_COLORS = [(n, 0x80000000 | i) for n, i in (
    ("Desktop", 1), ("Application Workspace", 12), ("Window Background", 5), ("Window Text", 8), ("Menu Bar", 4),
    ("Menu Text", 7), ("Active Title Bar", 2), ("Inactive Title Bar", 3), ("Active Title Bar Text", 9),
    ("Inactive Title Bar Text", 19), ("Active Border", 10), ("Inactive Border", 11), ("Window Frame", 6),
    ("Scroll Bars", 0), ("Button Face", 15), ("Button Shadow", 16), ("Button Text", 18), ("Button Highlight", 20),
    ("Disabled Text", 17), ("Highlight", 13), ("Highlight Text", 14), ("Button Dark Shadow", 21),
    ("Button Light Shadow", 22), ("ToolTip Text", 23), ("ToolTip", 24))]


def bindings(module):
    """{control name (lower): (property, "text" | "check")} of a page whose SelectionChanged and ApplyChanges are
    the wizard's lines; and whether its code does anything else (then it cannot work here)"""
    code = module.code_text().replace("\r\n", "\n").split("\n") if module is not None else []
    gets, sets, other, proc = {}, {}, False, None
    for line in code:
        s = line.strip()
        m = _PROC_RE.match(line)
        if m:
            proc = m.group(2).lower()
            continue
        if re.match(r"^End\s+Sub\b", s, re.IGNORECASE):
            proc = None
            continue
        if not s or s.startswith("'") or proc is None:
            continue
        if proc == "propertypage_selectionchanged":
            g = _GET_RE.match(line)
            if g:
                gets[(g.group(1) + g.group(2)).lower()] = (g.group(4), "check" if g.group(1).lower() == "chk" else "text")
            else:
                other = True
        elif proc == "propertypage_applychanges":
            g = _SET_RE.match(line)
            if g:
                sets[(g.group(2) + g.group(3)).lower()] = g.group(1)
            else:
                other = True
        elif not re.match(r"^Changed\s*=\s*True$", s, re.IGNORECASE):
            other = True
    out = {k: v for k, v in gets.items() if sets.get(k, "").lower() == v[0].lower()}
    return out, other or len(out) != len(gets) or len(out) != len(sets)


def _int(v, default=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


class PagesFrame(QDialog):
    """what the Property Pages of a control have in common (VB6: OleCreatePropertyFrame): the values of the control,
    the changes until Apply / OK, and the stock pages Font, Color and Picture (msstkprp.dll) as VB6 lays them out
    (phase 11, VB6 oracle DEV-11/ppcap: ocxpages.json); a user control (DesignPagesDialog) and an ActiveX control
    (ocxpages.OcxPagesDialog) use it"""

    def __init__(self, main, designer, ctl):
        super().__init__(main)
        self.main, self.designer, self.ctl = main, designer, ctl
        self.setWindowTitle("Property Pages")
        from PyQt6.QtGui import QPalette      # a disabled box: grey, as Windows draws it (not white)
        pal = self.palette()
        pal.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Base, pal.color(QPalette.ColorRole.Button))
        self.setPalette(pal)
        self.type_name = ctl.type_name
        self.pending = {}                       # property name -> new value (Apply)
        self.messages = []
        self.stock_props = {}                   # StdFont / StdColor / StdPicture -> the properties VB6 lists
        self.raw_apply = {}                     # a value of the .frm that is no property of the Properties window
        self.raw_check = {}                     # name -> the value -> the text of an error (before it is written)

    # ------------------------------------------------------------ the values of the control
    def prop(self, name):
        return V.prop_def(self.type_name, name)

    def value(self, name):
        p = self.prop(name)
        if p is None:
            return None
        form = self.designer.root if self.designer is not None else None
        if p.kind in ("picture", "string") and self.designer is not None:
            v = self.designer.module.frx_value(self.ctl, p.name)
            return p.default if v is None else v
        return V.value_of(self.ctl, p, form)

    def changed(self, name, value):
        self.pending[name] = value
        self.b_apply.setEnabled(True)

    def apply(self):
        """Apply: the values into the control (as the Properties window sets them); False when one is wrong"""
        raw = [(n, v) for n, v in self.pending.items() if n in self.raw_apply]
        for n, v in raw:
            err = self.raw_check[n](v) if n in self.raw_check else None
            if err:
                self.message(err)
                return False
        if raw and self.designer is not None:   # TabCaption(1) ...: written into the .frm as they are
            self.designer.checkpoint()
            for n, v in raw:
                self.raw_apply[n](v)
                del self.pending[n]
            self.designer.build()
            self.designer.changed()
        for name, value in list(self.pending.items()):
            p = self.prop(name)
            if p is None:
                continue
            v = value
            if isinstance(value, str) and p.kind not in ("string", "text", "font", "picture", "list"):
                if p.kind in ("int", "num", "pos", "scale", "formpos"):
                    value = value.replace(",", ".")     # 8,25: the decimal separator of the system
                try:
                    v = V.parse_input(p, value)     # what was typed in a box: as the Properties window reads it
                except (ValueError, OverflowError):
                    self.message("Invalid property value")
                    return False
            elif p.kind == "bool":
                v = bool(value)
            err = self.designer.set_property(self.ctl, p, v) if self.designer is not None else None
            if err:
                self.message(err)
                return False
            del self.pending[name]
        self.b_apply.setEnabled(False)
        props = getattr(self.main, "props", None)
        if props is not None and getattr(props, "ctl", None) is self.ctl:
            try:
                props.refresh_values()
            except Exception:                   # noqa: BLE001 - the window shows another object
                pass
        return True

    def ok(self):
        if self.apply():
            self.accept()

    def message(self, text):
        self.messages.append(text)
        if not getattr(self, "quiet", False):
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Property Pages", text)

    def _typed(self, kinds, stock=None):
        """the properties a stock page lists: the ones VB6 lists for the control (ocxpages.json) that LinBASIC Studio
        knows, else the ones of these kinds"""
        if stock is not None and stock in self.stock_props:
            return [n for n in self.stock_props[stock] if self.prop(n) is not None]
        return [p.name for p in V.props_of(self.type_name) if p.kind in kinds and p.name not in ("DragIcon",)]

    @staticmethod
    def _label(text, buddy, parent=None):
        lab = QLabel(text, parent)
        lab.setBuddy(buddy)                     # the & key, not a & on the screen
        return lab

    def _props_list(self, names, parent=None):
        lst = QListWidget(parent)
        lst.addItems(names)
        if names:
            lst.setCurrentRow(0)
        return lst

    @staticmethod
    def _stock_widget(size):
        w = QWidget()
        w.setFixedSize(*size)
        return w

    def _font_page(self, size=(393, 245)):
        """the Font page of msstkprp.dll (VB6 DEV-11/ppcap: StatusBar-2): Properties, Font, Size, Effects, Sample Text"""
        w = self._stock_widget(size)
        self.font_props = self._props_list(self._typed({"font"}, "StdFont"), w)
        self.font_props.setGeometry(9, 29, 132, max(60, size[1] - 72))
        self._label("&Properties:", self.font_props, w).setGeometry(9, 10, 99, 13)
        self.font_name = QFontComboBox(w)
        self.font_name.setEditable(True)        # a font of Windows (MS Sans Serif) keeps its name
        self.font_name.setGeometry(153, 29, 149, 21)
        self._label("&Font:", self.font_name, w).setGeometry(153, 10, 69, 13)
        self.font_size = QComboBox(w)
        self.font_size.setEditable(True)
        self.font_size.addItems(["8", "10", "12", "14", "18", "24"])
        self.font_size.setGeometry(312, 29, 69, 21)
        self._label("&Size:", self.font_size, w).setGeometry(312, 10, 48, 13)
        eff = QGroupBox("Effects", w)
        eff.setGeometry(153, 57, 228, 65)
        self.font_bold, self.font_italic = QCheckBox("Bold", w), QCheckBox("Italic", w)
        self.font_under, self.font_strike = QCheckBox("Underline", w), QCheckBox("Strikeout", w)
        self.font_bold.setGeometry(171, 76, 63, 16)
        self.font_italic.setGeometry(171, 98, 63, 16)
        self.font_under.setGeometry(261, 76, 84, 16)
        self.font_strike.setGeometry(261, 98, 84, 16)
        QLabel("Sample Text:", w).setGeometry(153, 130, 72, 13)
        self.font_sample = QLabel("AaBbYyZz", w)
        self.font_sample.setFrameShape(QFrame.Shape.Box)
        self.font_sample.setStyleSheet("background: white; color: black")
        self.font_sample.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.font_sample.setGeometry(153, 149, 228, max(40, size[1] - 160))
        self.font_props.currentRowChanged.connect(lambda *_: self._font_show())
        for c in (self.font_bold, self.font_italic, self.font_under, self.font_strike):
            c.toggled.connect(lambda *_: self._font_edit())
        self.font_name.currentTextChanged.connect(lambda *_: self._font_edit())
        self.font_size.currentTextChanged.connect(lambda *_: self._font_edit())
        self._font_show()
        return w

    def _font_show(self):
        it = self.font_props.currentItem()
        self._font_busy = True
        f = (self.pending.get(it.text()) or self.value(it.text())) if it is not None else None
        if f:
            self.font_name.setCurrentFont(QFont(str(f.get("Name", "MS Sans Serif"))))
            self.font_name.setEditText(str(f.get("Name", "MS Sans Serif")))
            self.font_size.setCurrentText(_size_text(f.get("Size", 8.25)))
            self.font_bold.setChecked(_int(f.get("Weight", 400)) >= 700)
            self.font_italic.setChecked(bool(f.get("Italic")))
            self.font_under.setChecked(bool(f.get("Underline")))
            self.font_strike.setChecked(bool(f.get("Strikethrough")))
        self._font_busy = False
        self._font_sample()

    def _font_sample(self):
        qf = QFont(self.font_name.currentText())
        try:
            qf.setPointSizeF(float(self.font_size.currentText().replace(",", ".")))
        except ValueError:
            pass
        qf.setBold(self.font_bold.isChecked())
        qf.setItalic(self.font_italic.isChecked())
        qf.setUnderline(self.font_under.isChecked())
        qf.setStrikeOut(self.font_strike.isChecked())
        self.font_sample.setFont(qf)
        self.font_sample.setText(self.font_name.currentText())

    def _font_edit(self):
        if getattr(self, "_font_busy", False):
            return
        it = self.font_props.currentItem()
        if it is None:
            return
        f = dict(self.pending.get(it.text()) or self.value(it.text()) or {})
        f["Name"] = self.font_name.currentText()
        try:
            f["Size"] = float(self.font_size.currentText().replace(",", "."))
        except ValueError:
            pass
        f["Weight"] = 700 if self.font_bold.isChecked() else 400
        f["Italic"], f["Underline"] = self.font_italic.isChecked(), self.font_under.isChecked()
        f["Strikethrough"] = self.font_strike.isChecked()
        self._font_sample()
        self.changed(it.text(), f)

    def _color_page(self, size=(393, 221)):
        """the Color page of msstkprp.dll (DEV-11/ppcap: ImageList-2): Properties with their colours, Color Set
        (Standard Colors, Windows System Colors), Color Palette, Edit Custom Color..."""
        w = self._stock_widget(size)
        self.color_props = self._props_list(self._typed({"color"}, "StdColor"), w)
        self.color_props.setGeometry(9, 29, 132, max(60, size[1] - 36))
        self._color_icons()
        self._label("&Properties:", self.color_props, w).setGeometry(9, 10, 99, 13)
        self.color_set = QComboBox(w)
        self.color_set.addItems(["Standard Colors", "Windows System Colors"])
        self.color_set.setGeometry(153, 29, 231, 21)
        self._label("Color &Set:", self.color_set, w).setGeometry(153, 10, 87, 13)
        self.color_list = QListWidget(w)
        self.color_list.setGeometry(153, 75, 231, 111)
        self._label("&Color Palette:", self.color_list, w).setGeometry(153, 55, 75, 13)
        b = QPushButton("&Edit Custom Color...", w)
        b.setGeometry(153, 192, 231, 23)
        b.clicked.connect(self._color_custom)
        self.color_set.currentIndexChanged.connect(lambda *_: self._color_fill())
        self.color_list.itemClicked.connect(lambda it: self._color_custom() if it.data(Qt.ItemDataRole.UserRole) is None
                                            else self.set_color(it.data(Qt.ItemDataRole.UserRole)))
        self.color_props.currentRowChanged.connect(lambda *_: self._color_show())
        self._color_show()
        return w

    def _color_show(self):
        """the colour set of the property chosen (a system colour: Windows System Colors) and its colour"""
        it = self.color_props.currentItem()
        v = self.pending.get(it.text(), self.value(it.text())) if it is not None else None
        system = v is not None and int(v) & 0x80000000 != 0
        self.color_set.blockSignals(True)
        self.color_set.setCurrentIndex(1 if system else 0)
        self.color_set.blockSignals(False)
        self._color_fill(v)

    def _color_fill(self, current=None):
        from .propwin import swatch
        from PyQt6.QtWidgets import QListWidgetItem
        self.color_list.clear()
        self.color_list.addItem(QListWidgetItem("<Custom...>"))
        table = COLORS if self.color_set.currentIndex() == 0 else SYSTEM_COLORS
        for name, v in table:
            it = QListWidgetItem(swatch(v), name)
            it.setData(Qt.ItemDataRole.UserRole, v)
            self.color_list.addItem(it)
        if current is not None:
            c = int(current) & 0xFFFFFFFF
            row = next((i + 1 for i, (_, v) in enumerate(table) if v == c), None)
            if row is None and c & 0x80000000 and self.color_set.currentIndex() == 1:
                it = QListWidgetItem(swatch(c), "Unrecogized System Color")    # as msstkprp.dll writes it
                it.setData(Qt.ItemDataRole.UserRole, c)
                self.color_list.addItem(it)
                row = self.color_list.count() - 1
            self.color_list.setCurrentRow(row if row is not None else 0)

    def _color_custom(self):
        c = QColorDialog.getColor(QColor("white"), self)
        if c.isValid():
            self.set_color(c.red() | (c.green() << 8) | (c.blue() << 16))

    def set_color(self, v):
        it = self.color_props.currentItem()
        if it is not None:
            self.changed(it.text(), int(v))
            self._color_icons()

    def _color_icons(self):
        """the color of each property beside its name"""
        from .propwin import swatch
        for i in range(self.color_props.count()):
            it = self.color_props.item(i)
            v = self.pending.get(it.text(), self.value(it.text()))
            if v is not None:
                it.setIcon(swatch(int(v)))

    def _picture_page(self, size=(393, 245)):
        """the Picture page of msstkprp.dll (DEV-11/ppcap: StatusBar-3): Properties, Preview, Browse..., Clear"""
        w = self._stock_widget(size)
        self.pic_props = self._props_list(self._typed({"picture"}, "StdPicture"), w)
        self.pic_props.setGeometry(9, 29, 132, max(60, size[1] - 72))
        self._label("&Properties:", self.pic_props, w).setGeometry(9, 10, 99, 13)
        QLabel("Preview:", w).setGeometry(153, 10, 63, 13)
        self.pic_view = QLabel("(None)", w)
        self.pic_view.setFrameShape(QFrame.Shape.Box)
        self.pic_view.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.pic_view.setGeometry(153, 29, 231, max(40, size[1] - 92))
        b1, self.pic_clear = QPushButton("&Browse...", w), QPushButton("&Clear", w)
        b1.setGeometry(153, size[1] - 56, 105, 23)
        self.pic_clear.setGeometry(279, size[1] - 56, 105, 23)
        b1.clicked.connect(self._pic_browse)
        self.pic_clear.clicked.connect(lambda: self.set_picture(None))
        self.pic_props.currentRowChanged.connect(lambda *_: self._pic_show())
        self._pic_show()
        return w

    def _pic_show(self):
        """the picture of the property chosen; Clear only when it has one (VB6 DEV-9/pp-linux)"""
        it = self.pic_props.currentItem()
        v = self.pending.get(it.text(), self.value(it.text())) if it is not None else None
        self.pic_view.setPixmap(QPixmap())
        if v and isinstance(v, (bytes, bytearray)):
            pm = QPixmap()
            if pm.loadFromData(bytes(v)):
                self.pic_view.setPixmap(pm.scaled(self.pic_view.size() - QSize(4, 4), Qt.AspectRatioMode.KeepAspectRatio))
            else:
                self.pic_view.setText("(Bitmap)")
        else:
            self.pic_view.setText("(Bitmap)" if v else "(None)")
        self.pic_clear.setEnabled(bool(v))

    def _pic_browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Browse Picture", "",
                                              "All Picture Files (*.bmp *.dib *.ico *.cur *.wmf *.emf *.gif *.jpg);;"
                                              "Bitmaps (*.bmp *.dib);;Icons & Cursors (*.ico *.cur);;"
                                              "Metafiles (*.wmf *.emf);;GIF Images (*.gif);;JPEG Images (*.jpg);;"
                                              "All Files (*)")
        if path:
            with open(path, "rb") as f:
                self.set_picture(f.read())

    def set_picture(self, data):
        it = self.pic_props.currentItem()
        if it is not None:
            self.changed(it.text(), data)
            self._pic_show()


def _size_text(v):
    """a size of a font as the Font page shows it: 8,25 with the decimal separator of the system"""
    from PyQt6.QtCore import QLocale
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    t = (str(int(f)) if f == int(f) else f"{f:g}")
    return t.replace(".", QLocale().decimalPoint())


class DesignPagesDialog(PagesFrame):
    """Property Pages: a tab for each page of the user control (its PropertyPages), OK / Cancel / Apply"""

    def __init__(self, main, designer, ctl, uc_module):
        super().__init__(main, designer, ctl)
        self.uc = uc_module
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs, 1)
        data = uc_module.frx_value(uc_module.sf.root, "PropertyPages")
        self.page_names = G.read_pages(data)
        self.page_widgets = {}
        for name in self.page_names:
            w = self._page(name)
            if w is not None:
                self.tabs.addTab(w, self._tab_title(name))
        row = QHBoxLayout()
        row.addStretch(1)
        self.b_ok = QPushButton("OK")
        self.b_cancel = QPushButton("Cancel")
        self.b_apply = QPushButton("&Apply")
        for b in (self.b_ok, self.b_cancel, self.b_apply):
            row.addWidget(b)
        lay.addLayout(row)
        self.b_ok.setDefault(True)
        self.b_ok.clicked.connect(self.ok)
        self.b_cancel.clicked.connect(self.reject)
        self.b_apply.clicked.connect(self.apply)
        self.b_apply.setEnabled(False)

    # ------------------------------------------------------------ the pages
    def _tab_title(self, name):
        if name in STANDARD_TABS:
            return STANDARD_TABS[name]
        m = self._module(name)
        cap = m.sf.root.get("Caption") if m is not None and m.sf is not None and m.sf.root is not None else None
        return str(cap) if isinstance(cap, str) and cap else name

    def _module(self, name):
        return next((m for m in self.uc.project.modules if m.kind == "PropertyPage" and m.name.lower() == name.lower()),
                    None)

    def _page(self, name):
        if name == "StandardFont":
            return self._font_page()
        if name == "StandardColor":
            return self._color_page()
        if name == "StandardPicture":
            return self._picture_page()
        m = self._module(name)
        if m is None:
            return None
        return self._custom_page(m)

    def _custom_page(self, m):
        """the controls of a page of the project, bound to the properties when its code is the wizard's"""
        binds, other = bindings(m)
        w = QWidget()
        w.setObjectName(m.name)
        root = m.sf.root
        self.page_widgets[m.name] = {}
        width = _int(root.get("ScaleWidth", root.get("ClientWidth", 5925)), 5925)
        height = _int(root.get("ScaleHeight", root.get("ClientHeight", 3495)), 3495)
        w.setMinimumSize(width // TPP, height // TPP)
        for c in root.walk():
            if c is root:
                continue
            x, y = _int(c.get("Left", 0)) // TPP, _int(c.get("Top", 0)) // TPP
            cw, ch = max(8, _int(c.get("Width", 1200)) // TPP), max(8, _int(c.get("Height", 330)) // TPP)
            t, cap = c.type_name, c.get("Caption", "")
            if t == "Label":
                wd = QLabel(str(cap or ""), w)
            elif t == "TextBox":
                wd = QLineEdit(w)
            elif t == "CheckBox":
                wd = QCheckBox(str(cap or ""), w)
            else:
                wd = QLabel(f"({t})", w)
                wd.setEnabled(False)
            wd.setGeometry(x, y, cw, ch)
            self.page_widgets[m.name][c.name.lower()] = wd
            b = binds.get(c.name.lower())
            if b is None:
                if t in ("TextBox", "CheckBox"):
                    wd.setEnabled(False)
                continue
            prop, kind = b
            v = self.value(prop)
            if kind == "check" and isinstance(wd, QCheckBox):
                wd.setChecked(bool(v))
                wd.toggled.connect(lambda on, prop=prop: self.changed(prop, on))
            elif isinstance(wd, QLineEdit):
                wd.setText("" if v is None else (V.display(self.prop(prop), v) if self.prop(prop) is not None and
                                                 self.prop(prop).kind == "bool" else str(v)))
                wd.textEdited.connect(lambda s, prop=prop: self.changed(prop, s))
        if other:
            note = QLabel("This page has code LinBASIC Studio cannot run at design time: only the boxes the "
                          "Property Page Wizard binds to properties work here.", w)
            note.setWordWrap(True)
            note.setStyleSheet("color: #800000")
            note.setGeometry(4, max(4, height // TPP - 40), max(200, width // TPP - 8), 36)
        return w


def uc_pages_of(main, designer, ctl):
    """the user control module of a control on a form when it has property pages, else None"""
    from .designer import uctl_of
    uc = uctl_of(ctl.cls) if ctl is not None else None
    if not uc:
        return None
    m = uc[4]
    if m is None or m.sf is None or m.sf.root is None:
        return None
    return m if G.read_pages(m.frx_value(m.sf.root, "PropertyPages")) else None


__all__ = ["DesignPagesDialog", "uc_pages_of", "bindings"]
