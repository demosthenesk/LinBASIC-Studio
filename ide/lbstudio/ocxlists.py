"""ocxlists.py - the property pages of the lists of the ActiveX controls (phase 11, stage 2c): the Tabs of a TabStrip,
the Panels of a StatusBar, the Buttons (and their ButtonMenus) of a Toolbar, the Column Headers of a ListView and the
Images of an ImageList.

Each page has an Index (a box and the arrows of a scroll bar), Insert / Remove and the boxes of the item shown; the
list goes into the control on Apply. The .frm gets the blocks VB6 writes (VB6 oracle DEV-11 s_lists; the corpus of
PSC: Claude Enviroment/tools/blockscan.py, blockorder.py): their GUIDs, the properties of an item in the order of VB6,
a property only when it is not its default ("Object." before Tag, ToolTipText, Width, Visible).
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
from PyQt6.QtWidgets import QLineEdit, QComboBox, QCheckBox, QScrollBar, QListWidget, QListWidgetItem, QPushButton

from lbvc.project.source import Property, PropertyBlock
from lbvc.project.frx import FrxData
from . import vbprops as V

G_BUTTONS, G_BUTTON = "{66833FE8-8583-11D1-B16A-00C0F0283628}", "{66833FEA-8583-11D1-B16A-00C0F0283628}"
G_MENUS, G_MENU = "{66833FEC-8583-11D1-B16A-00C0F0283628}", "{66833FEE-8583-11D1-B16A-00C0F0283628}"
G_PANELS, G_PANEL = "{8E3867A5-8586-11D1-B16A-00C0F0283628}", "{8E3867AB-8586-11D1-B16A-00C0F0283628}"
G_TABS, G_TAB = "{1EFB6598-857C-11D1-B16A-00C0F0283628}", "{1EFB659A-857C-11D1-B16A-00C0F0283628}"
G_IMAGES, G_IMAGE = "{2C247F25-8591-11D1-B16A-00C0F0283628}", "{2C247F27-8591-11D1-B16A-00C0F0283628}"
G_COLUMN = "{BDD1F052-858B-11D1-B16A-00C0F0283628}"


def _num(v, d=0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return d


def _twips(v):
    """a width of the .frm: whole twips, as VB6 writes them (the corpus: no width of an item with a fraction)"""
    return int(round(float(v)))


# ---------------------------------------------------------------- the items <-> the blocks of the .frm
# (name in the item, name in the .frm, kind, default: written only when the value differs; None: always)
ITEMS = {
    "Tab": [("Caption", "Caption", "s", ""), ("Key", "Key", "s", ""), ("Tag", "Object.Tag", "s", ""),
            ("ToolTipText", "Object.ToolTipText", "s", ""), ("ImageVarType", "ImageVarType", "i", None),
            ("Image", "ImageIndex", "i", 0)],
    "Panel": [("Style", "Style", "i", 0), ("Alignment", "Alignment", "i", 0), ("AutoSize", "AutoSize", "i", 0),
              ("Bevel", "Bevel", "i", 1), ("Enabled", "Enabled", "b", True), ("Visible", "Object.Visible", "b", True),
              ("Width", "Object.Width", "w", 1440), ("MinWidth", "MinWidth", "w", 1440), ("Picture", "Picture", "p", None),
              ("Text", "Text", "s", ""), ("TextSave", "TextSave", "s", ""), ("Key", "Key", "s", ""),
              ("Tag", "Object.Tag", "s", ""), ("ToolTipText", "Object.ToolTipText", "s", "")],
    "Button": [("Enabled", "Enabled", "b", True), ("Visible", "Object.Visible", "b", True), ("Caption", "Caption", "s", ""),
               ("Key", "Key", "s", ""), ("Description", "Description", "s", ""),
               ("ToolTipText", "Object.ToolTipText", "s", ""), ("Tag", "Object.Tag", "s", ""),
               ("ImageKey", "ImageKey", "s", ""), ("Image", "ImageIndex", "i", 0), ("Style", "Style", "i", 0),
               ("Width", "Object.Width", "w", 0), ("MixedState", "MixedState", "b", False), ("Value", "Value", "i", 0)],
    "ButtonMenu": [("Enabled", "Enabled", "b", True), ("Visible", "Object.Visible", "b", True), ("Key", "Key", "s", ""),
                   ("Tag", "Object.Tag", "s", ""), ("Text", "Text", "s", "")],
    "ColumnHeader": [("Alignment", "Alignment", "i", 0), ("SubItemIndex", "SubItemIndex", "i", 0),
                     ("Key", "Key", "s", ""), ("Tag", "Object.Tag", "s", ""), ("Text", "Text", "s", ""),
                     ("Width", "Object.Width", "w", 1440), ("Icon", "ImageIndex", "i", 0)],
    "ListImage": [("Picture", "Picture", "p", None), ("Key", "Key", "s", None), ("Tag", "Object.Tag", "s", "")],
}
NEW = {"Tab": {"ImageVarType": 2}, "Panel": {}, "Button": {"Menus": []}, "ButtonMenu": {}, "ColumnHeader": {},
       "ListImage": {}}


def new_item(kind, **kw):
    it = {k: ("" if t == "s" else 0) if d is None and t != "p" else d for k, _, t, d in ITEMS[kind]}
    it.update({k: list(v) if isinstance(v, list) else v for k, v in NEW[kind].items()})    # (a list of its own)
    it.update(kw)
    return it


def read_item(kind, blk, frx=None):
    it = new_item(kind)
    for k, n, t, d in ITEMS[kind]:
        p = next((x for x in blk.items if isinstance(x, Property) and x.name.lower() == n.lower()), None)
        if p is None:
            continue
        v = p.value
        if t == "p":
            it[k] = frx(p) if frx else None
        elif t == "s":
            it[k] = "" if v is None else str(v)
        elif t == "b":
            it[k] = bool(_num(v)) if not isinstance(v, bool) else v
        elif t == "w":
            it[k] = _num(v, d or 0)
        else:
            it[k] = int(_num(v))
    return it


def item_block(kind, name, guid, it):
    b = PropertyBlock(name, guid)
    for k, n, t, d in ITEMS[kind]:
        v = it.get(k)
        if t == "p":
            if v:
                b.items.append(Property(n, FrxData("picture", bytes(v))))
            continue
        if d is not None and (v == d or v is None or (t == "s" and str(v) == str(d)) or
                              (t == "w" and abs(_num(v) - d) < 0.5)):
            continue
        if t == "b":
            b.items.append(Property(n, bool(v)))
        elif t == "w":
            b.items.append(Property(n, _twips(v)))
        elif t == "i":
            b.items.append(Property(n, int(v)))
        else:
            b.items.append(Property(n, str(v)))
    return b


def _frx_reader(dlg):
    mod = dlg.designer.module if dlg.designer is not None else None

    def get(p):
        if mod is None:
            return None
        d = mod.frx_data(p)
        return bytes(d.data) if d is not None and isinstance(d.data, (bytes, bytearray)) else None
    return get


def _block(ctl, name):
    return next((x for x in ctl.items if isinstance(x, PropertyBlock) and x.name.lower() == name.lower()), None)


def _put_block(ctl, blk):
    """a list block in the place VB6 writes it: right after _Version (before the Font)"""
    old = _block(ctl, blk.name)
    if old is not None:
        V.remove(ctl, old.name, block=True)
    ctl.items.append(blk)
    pos = next((i + 1 for i, (k, it) in enumerate(ctl.order) if k == "p" and getattr(it, "name", "") == "_Version"), None)
    if pos is None:
        pos = next((i for i, (k, _) in enumerate(ctl.order) if k == "c"), len(ctl.order))
    ctl.order.insert(pos, ("p", blk))


def read_list(dlg, kind):
    """the items of the list of the control of the dialog"""
    return items_of(dlg.ctl, kind, _frx_reader(dlg))


def items_of(ctl, kind, frx=None):
    """the items of a list of a control of the .frm (Tab, Panel, Button, ColumnHeader, ListImage); frx: the bytes of
    a picture of the .frx (the designer draws without them)"""
    if kind == "ColumnHeader":
        n = int(_num(ctl.get("NumItems", 0)))
        return [read_item(kind, b, frx) for i in range(1, n + 1) for b in [_block(ctl, f"ColumnHeader({i})")] if b]
    outer, inner, count = {"Tab": ("Tabs", "Tab", "NumTabs"), "Panel": ("Panels", "Panel", "NumPanels"),
                           "Button": ("Buttons", "Button", "NumButtons"),
                           "ListImage": ("Images", "ListImage", "NumListImages")}[kind]
    o = _block(ctl, outer)
    if o is None:                               # a new control: VB6 has one tab / one panel
        return [new_item(kind)] if kind in ("Tab", "Panel") else []
    out = []
    for i in range(1, int(_num(o.get(count, 0))) + 1):
        b = next((x for x in o.items if isinstance(x, PropertyBlock) and x.name.lower() == f"{inner}{i}".lower()), None)
        it = read_item(kind, b, frx) if b is not None else new_item(kind)
        if kind == "Button" and b is not None:
            mb = next((x for x in b.items if isinstance(x, PropertyBlock) and x.name.lower() == "buttonmenus"), None)
            it["Menus"] = []
            if mb is not None:
                for j in range(1, int(_num(mb.get("NumButtonMenus", 0))) + 1):
                    m = next((x for x in mb.items if isinstance(x, PropertyBlock) and
                              x.name.lower() == f"buttonmenu{j}"), None)
                    it["Menus"].append(read_item("ButtonMenu", m) if m is not None else new_item("ButtonMenu"))
        out.append(it)
    return out


def write_list(ctl, kind, items):
    """the items into the .frm of the control, as VB6 writes them"""
    if kind == "ColumnHeader":
        for x in list(ctl.items):
            if isinstance(x, PropertyBlock) and x.name.lower().startswith("columnheader("):
                V.remove(ctl, x.name, block=True)
        V.remove(ctl, "NumItems")
        V._append(ctl, Property("NumItems", len(items)))
        for i, it in enumerate(items, 1):
            it = dict(it, SubItemIndex=i - 1)
            V._append(ctl, item_block(kind, f"ColumnHeader({i})", G_COLUMN, it))
        return
    outer, inner, count, go, gi = {"Tab": ("Tabs", "Tab", "NumTabs", G_TABS, G_TAB),
                                   "Panel": ("Panels", "Panel", "NumPanels", G_PANELS, G_PANEL),
                                   "Button": ("Buttons", "Button", "NumButtons", G_BUTTONS, G_BUTTON),
                                   "ListImage": ("Images", "ListImage", "NumListImages", G_IMAGES, G_IMAGE)}[kind]
    if not items and kind in ("Button", "ListImage"):
        old = _block(ctl, outer)
        if old is not None:
            V.remove(ctl, old.name, block=True)
        return
    o = PropertyBlock(outer, go)
    o.items.append(Property(count, len(items)))
    for i, it in enumerate(items, 1):
        b = item_block(kind, f"{inner}{i}", gi, it)
        if kind == "Button" and it.get("Menus"):
            mb = PropertyBlock("ButtonMenus", G_MENUS)
            mb.items.append(Property("NumButtonMenus", len(it["Menus"])))
            for j, m in enumerate(it["Menus"], 1):
                mb.items.append(item_block("ButtonMenu", f"ButtonMenu{j}", G_MENU, m))
            b.items.append(mb)
        o.items.append(b)
    _put_block(ctl, o)


# ---------------------------------------------------------------- the editor of a page
class ListEditor:
    """Index (a box and a scroll bar), Insert / Remove and the boxes of the item shown (fields: id -> (key, kind):
    s text, i number, w width (twips), e "N - name" of a combo box, b check box); every change goes to the dialog:
    the whole list on Apply"""

    def __init__(self, dlg, key, made, items, fields, kind, name, idx=209, scroll=223, ins=3, rem=4, after=None,
                 count=None, raw=True):
        from .ocxpages import by_id
        self.dlg, self.key, self.kind, self.name, self.items = dlg, key, kind, name, items
        self.fields = {by_id(made, i): f for i, f in fields.items() if by_id(made, i) is not None}
        self.idx, self.scroll = by_id(made, idx), by_id(made, scroll)
        self.b_ins, self.b_rem = by_id(made, ins), by_id(made, rem)
        self.count = by_id(made, count) if count else None
        self.after = after                      # what a page does after showing an item (enable boxes ...)
        self.on_edit = None                     # what a page does after a value of an item changed
        self.check = None                       # the items -> the text of an error, before Apply writes them
        self.cur = 0 if items else -1
        self.busy = False
        for w, (k, t) in self.fields.items():
            if isinstance(w, QCheckBox):
                w.toggled.connect(lambda on, k=k: self.edit(k, on))
            elif isinstance(w, QComboBox):
                w.activated.connect(lambda i, w=w, k=k: self.edit(k, _enum(w.itemText(i))))
            elif isinstance(w, QLineEdit):
                w.textEdited.connect(lambda s, k=k, t=t: self.edit(k, s, t))
        if isinstance(self.scroll, QScrollBar):
            self.scroll.valueChanged.connect(lambda v: None if self.busy else self.show(v - 1))
        if isinstance(self.idx, QLineEdit):
            self.idx.textEdited.connect(self._typed)
        if self.b_ins is not None:
            self.b_ins.clicked.connect(lambda: self.insert())
        if self.b_rem is not None:
            self.b_rem.clicked.connect(lambda: self.remove())
        if key is not None and raw:             # a list of the .frm (not a property of the Properties window)
            dlg.raw_apply[name] = lambda v, kind=kind: write_list(dlg.ctl, kind, v)
            dlg.raw_check[name] = lambda v: self.check(v) if self.check is not None else None
        if key is not None:                     # (the ButtonMenus: a part of the list of the Buttons)
            dlg.bound[key] = {name: self.idx}
            dlg.unbound[key] = []
            dlg.__dict__.setdefault("lists", {})[name] = self
        self.show(self.cur)

    def _typed(self, s):
        try:
            i = int(s) - 1
        except ValueError:
            return
        if 0 <= i < len(self.items):
            self.show(i)

    def show(self, i):
        self.busy = True
        self.cur = i if 0 <= i < len(self.items) else (len(self.items) - 1 if self.items else -1)
        it = self.items[self.cur] if self.cur >= 0 else None
        if self.idx is not None:
            self.idx.setText(str(self.cur + 1) if it is not None else "0")
            self.idx.setEnabled(it is not None)
        if isinstance(self.scroll, QScrollBar):
            self.scroll.setRange(1, max(1, len(self.items)))
            self.scroll.setValue(self.cur + 1 if it is not None else 1)
            self.scroll.setEnabled(len(self.items) > 1)
        shown = it if it is not None else (self.new() if callable(getattr(self, "new", None)) else
                                           new_item(self.kind))     # (VB6: the values of a new item, greyed)
        for w, (k, t) in self.fields.items():
            w.setEnabled(it is not None)
            v = shown.get(k)
            if isinstance(w, QCheckBox):
                w.setChecked(bool(v) if v is not None else False)
            elif isinstance(w, QComboBox):
                n = int(v or 0)
                w.setCurrentIndex(next((j for j in range(w.count()) if _enum(w.itemText(j)) == n), 0))
            elif isinstance(w, QLineEdit):
                w.setText("" if v is None else (_shown_twips(v) if t == "w" else str(v)))
        if self.b_rem is not None:
            self.b_rem.setEnabled(it is not None)
        if self.count is not None:
            self.count.setText(str(len(self.items)))
        self.busy = False
        if self.after is not None:
            self.after(self, it)

    def edit(self, k, v, t="s"):
        if self.busy or self.cur < 0:
            return
        if t in ("i", "w"):
            try:
                v = float(str(v).replace(",", ".")) if t == "w" else int(v or 0)
            except ValueError:
                return
        self.items[self.cur][k] = v
        if self.on_edit is not None:
            self.on_edit(self, k, v)
        self.changed()

    def changed(self):
        self.dlg.changed(self.name, [dict(x) for x in self.items])
        if self.after is not None:
            self.after(self, self.items[self.cur] if self.cur >= 0 else None)

    def insert(self, item=None):
        it = item if item is not None else (self.new() if callable(getattr(self, "new", None)) else
                                            new_item(self.kind))
        self.items.insert(self.cur + 1, it)
        self.show(self.cur + 1)
        self.changed()

    def remove(self):
        if self.cur < 0:
            return
        del self.items[self.cur]
        self.show(min(self.cur, len(self.items) - 1))
        self.changed()


def _enum(text):
    from .ocxpages import enum_value
    v = enum_value(text)
    return 0 if v is None else v


def _shown_twips(v):
    from PyQt6.QtCore import QLocale
    f = float(v)
    t = str(int(f)) if f == int(f) else f"{f:.4f}".rstrip("0").rstrip(".")
    return t.replace(".", QLocale().decimalPoint())


# ---------------------------------------------------------------- the pages
def tabs_page(dlg, key, w, made):
    """the Tabs of a TabStrip: Caption, Key, Tag, ToolTipText, Image (an index of its ImageList)"""
    ListEditor(dlg, key, made, read_list(dlg, "Tab"),
               {212: ("Caption", "s"), 214: ("Key", "s"), 277: ("Tag", "s"), 282: ("ToolTipText", "s"),
                219: ("Image", "i")}, "Tab", "(Tabs)", after=_image_box(dlg, made, 219))


def _image_box(dlg, made, wid):
    """the Image box of an item: only when the control has an ImageList (VB6)"""
    from .ocxpages import by_id
    box = by_id(made, wid)

    def after(ed, it):
        if box is not None:
            box.setEnabled(it is not None and bool(dlg.pending.get("ImageList", dlg.value("ImageList"))))
    return after


def panels_page(dlg, key, w, made):
    """the Panels of a StatusBar: Text, ToolTipText, Key, Tag, Minimum Width, Actual Width, Alignment, Style, Bevel,
    AutoSize, Picture, Enabled, Visible"""
    from .ocxpages import by_id
    from PyQt6.QtWidgets import QFileDialog
    items = read_list(dlg, "Panel")

    def after(ed, it):
        pic = it.get("Picture") if it else None
        by_id(made, 232).setEnabled(bool(pic))
        by_id(made, 231).setEnabled(it is not None)
        by_id(made, 218).setText(_shown_twips(it.get("Width", 1440)) if it is not None else "")

    def on_edit(ed, k, v):
        it = ed.items[ed.cur]
        if k == "MinWidth":                     # VB6 oracle s_lists2: the Minimum Width is the width too
            it["Width"] = v
        it["TextSave"] = text_save(it)
    ed = ListEditor(dlg, key, made, items,
                    {212: ("Text", "s"), 282: ("ToolTipText", "s"), 214: ("Key", "s"), 277: ("Tag", "s"),
                     220: ("MinWidth", "w"), 211: ("Alignment", "e"), 206: ("Style", "e"), 217: ("Bevel", "e"),
                     219: ("AutoSize", "e"), 208: ("Enabled", "b"), 215: ("Visible", "b")}, "Panel", "(Panels)",
                    after=after)
    ed.on_edit = on_edit
    by_id(made, 218).setReadOnly(True)
    from .ocxpages import gray
    gray(by_id(made, 218))

    def browse():
        path, _ = QFileDialog.getOpenFileName(dlg, "Load Picture", "", "Pictures (*.bmp *.dib *.gif *.jpg *.ico *.cur "
                                              "*.wmf *.emf);;All Files (*)")
        if path and ed.cur >= 0:
            with open(path, "rb") as f:
                ed.items[ed.cur]["Picture"] = f.read()
            ed.changed()

    def clear():
        if ed.cur >= 0:
            ed.items[ed.cur]["Picture"] = None
            ed.changed()
    by_id(made, 231).clicked.connect(browse)
    by_id(made, 232).clicked.connect(clear)


def text_save(it):
    """what a panel shows at design time, as VB6 keeps it in TextSave"""
    from PyQt6.QtCore import QTime, QDate, QLocale
    st = int(it.get("Style", 0))
    if st == 0:
        return str(it.get("Text", ""))
    if st == 5:                                 # the corpus: "8:19 PM", "11:35 a.m." - no seconds
        import re
        fmt = re.sub(r"[:.]s+", "", QLocale().timeFormat(QLocale.FormatType.ShortFormat))
        return QLocale().toString(QTime.currentTime(), fmt)
    if st == 6:
        return QLocale().toString(QDate.currentDate(), QLocale.FormatType.ShortFormat)
    return {1: "CAPS", 2: "NUM", 3: "INS", 4: "SCRL", 7: "KANA"}.get(st, "")


def buttons_page(dlg, key, w, made):
    """the Buttons of a Toolbar and the ButtonMenus of a button (Style 5 - tbrDropdown)"""
    from .ocxpages import by_id

    def after(ed, it):
        st = int(it.get("Style", 0)) if it else 0
        by_id(made, 222).setEnabled(it is not None and st == 4)         # Width: a placeholder
        by_id(made, 297).setEnabled(it is not None)                     # (Value: of every style, VB6)
        ms = it.setdefault("Menus", []) if it is not None else []
        if menus.items is not ms:               # another button: its first menu
            menus.items = ms
            menus.show(0 if ms else -1)
        _image_box(dlg, made, 219)(ed, it)
    menus = None
    ed = ListEditor(dlg, key, made, read_list(dlg, "Button"),
                    {212: ("Caption", "s"), 220: ("Description", "s"), 214: ("Key", "s"), 276: ("Value", "e"),
                     206: ("Style", "e"), 222: ("Width", "w"), 282: ("ToolTipText", "s"), 219: ("Image", "i"),
                     277: ("Tag", "s"), 263: ("Visible", "b"), 264: ("Enabled", "b"), 265: ("MixedState", "b")},
                    "Button", "(Buttons)")
    menus = _MenuEditor(ed, made)
    ed.after = after
    ed.show(ed.cur)


class _MenuEditor(ListEditor):
    """the ButtonMenus of the button shown: their changes are changes of the list of buttons"""

    def __init__(self, buttons, made):
        self.parent = buttons
        super().__init__(buttons.dlg, None, made, [], {299: ("Text", "s"), 300: ("Key", "s"), 301: ("Tag", "s"),
                                                       302: ("Enabled", "b"), 312: ("Visible", "b")},
                         "ButtonMenu", "(Buttons)", idx=295, scroll=296, ins=297, rem=298)

    def changed(self):
        self.parent.changed()


def columns_page(dlg, key, w, made):
    """the Column Headers of a ListView: Text, Alignment, Width, Key, Tag, Icon Index"""
    from .ocxpages import by_id

    def after(ed, it):                          # VB6: no column, an empty Index and Icon Index
        if it is None:
            by_id(made, 209).setText("")
            by_id(made, 311).setText("")
    ed = ListEditor(dlg, key, made, read_list(dlg, "ColumnHeader"),
                    {212: ("Text", "s"), 211: ("Alignment", "e"), 218: ("Width", "w"), 214: ("Key", "s"),
                     277: ("Tag", "s"), 311: ("Icon", "i")}, "ColumnHeader", "(ColumnHeaders)", after=after)
    ed.check = lambda items: ("The first column in a ListView control must be left aligned"
                              if items and int(items[0].get("Alignment", 0)) != 0 else None)


def images_page(dlg, key, w, made):
    """the Images of an ImageList: Insert Picture... (files of pictures), Remove Picture, the Key and Tag of the image
    chosen, the pictures in the list, Image Count; the first picture sets ImageWidth / ImageHeight when they are 0"""
    from .ocxpages import by_id
    from PyQt6.QtWidgets import QFileDialog
    from PyQt6.QtGui import QPixmap, QIcon
    from PyQt6.QtCore import QSize
    lst = by_id(made, 180)

    def after(ed, it):
        if ed.busy:
            return
        lst.blockSignals(True)
        lst.clear()
        lst.setViewMode(QListWidget.ViewMode.IconMode)
        lst.setIconSize(QSize(32, 32))
        for x in ed.items:
            pm = QPixmap()
            pm.loadFromData(bytes(x.get("Picture") or b""))
            lst.addItem(QListWidgetItem(QIcon(pm), ""))
        lst.setCurrentRow(ed.cur)
        lst.setEnabled(bool(ed.items))
        lst.blockSignals(False)
    ed = ListEditor(dlg, key, made, read_list(dlg, "ListImage"), {214: ("Key", "s"), 277: ("Tag", "s")},
                    "ListImage", "(ListImages)", scroll=None, ins=247, rem=257, count=229, after=after)
    lst.currentRowChanged.connect(lambda r: ed.show(r) if r >= 0 else None)

    def insert():
        paths, _ = QFileDialog.getOpenFileNames(dlg, "Select picture", "",
                                                "Graphics Files (*.bmp *.cur *.ico *.jpg *.gif);;All Files (*)")
        add(paths)

    def add(paths):
        for path in paths:
            with open(path, "rb") as f:
                data = f.read()
            if not dlg.pending.get("ImageWidth", dlg.value("ImageWidth")):
                pm = QPixmap()
                if pm.loadFromData(data):
                    dlg.changed("ImageWidth", pm.width())
                    dlg.changed("ImageHeight", pm.height())
            ed.insert(new_item("ListImage", Picture=data))
    b = by_id(made, 247)
    b.clicked.disconnect()
    b.clicked.connect(insert)
    ed.add_files = add


def bands_page(dlg, key, w, made):
    """the Bands of a CoolBar (a page written in VB: its labels have no windows): Index and its arrows, Insert /
    Remove Band, Child (a control inside the CoolBar), Style, the picture and the colours of the band, Caption, Width,
    MinWidth, MinHeight, Key, Tag, Image, Visible, AllowVertical, NewRow, FixedBackground (vbprops.cool_bands)"""
    from .ocxpages import by_id
    from PyQt6.QtWidgets import QLabel, QFrame, QStyle, QColorDialog, QFileDialog
    from PyQt6.QtGui import QColor
    from . import vbprops as VP
    labels = {}
    for text, x, y, buddy in (("Index:", 4, 9, 16412), ("C&hild:", 4, 39, 16404), ("St&yle:", 4, 63, 16389),
                              ("&Picture:", 18, 111, 16398), ("EmbossHighli&ght", 18, 171, 16387),
                              ("EmbossSha&dow", 18, 187, 16386), ("&ForeColor:", 18, 231, 16394),
                              ("&BackColor:", 18, 247, 16393), ("&Caption:", 191, 95, 16406), ("&Width:", 191, 119, 16403),
                              ("Min&Width:", 191, 143, 16401), ("Min&Height:", 191, 167, 16402),
                              ("&Key:", 191, 191, 16400), ("&Tag:", 191, 215, 16388), ("I&mage:", 191, 239, 16405)):
        lab = QLabel(text, w)
        lab.setBuddy(by_id(made, buddy))
        lab.setGeometry(x, y, 98, 14)
        lab.show()
        labels[buddy] = lab
    pic = QLabel("(None)", w)
    pic.setFrameShape(QFrame.Shape.Panel)
    pic.setFrameShadow(QFrame.Shadow.Sunken)
    pic.setGeometry(18, 125, 132, 20)
    pic.show()
    by_id(made, 16398).setIcon(dlg.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
    left, right = by_id(made, 16391), by_id(made, 16390)
    left.setText("<")
    right.setText(">")
    child, style = by_id(made, 16404), by_id(made, 16389)
    kids = [c.name if c.get("Index") is None else f"{c.name}({int(VP._num(c.get('Index')))})"
            for k, c in dlg.ctl.order if k == "c"]
    child.clear()
    child.addItems(["(None)"] + kids)
    style.clear()
    style.addItems(["0 - cc3BandNormal", "1 - cc3BandFixedSize"])
    colors = {16387: "EmbossHighlight", 16386: "EmbossShadow", 16394: "ForeColor", 16393: "BackColor"}

    def after(ed, it):
        on = it is not None
        left.setEnabled(on and ed.cur > 0)
        right.setEnabled(on and ed.cur < len(ed.items) - 1)
        ed.busy = True
        c = str(it.get("Child", "") or "") if on else ""
        child.setCurrentIndex(next((i for i in range(child.count()) if child.itemText(i).lower() == c.lower()), 0))
        ed.busy = False
        upic = on and bool(it.get("UseCoolbarPicture", True))
        by_id(made, 16398).setEnabled(on and not upic)
        labels[16398].setEnabled(on and not upic)
        pic.setEnabled(on and not upic)
        pic.setText("(Bitmap)" if on and it.get("Picture") else "(None)")
        emb = by_id(made, 16385)
        emb.setEnabled(on and not upic)
        for wid in (16387, 16386):
            by_id(made, wid).setEnabled(on and not upic and bool(it.get("EmbossPicture")))
            labels[wid].setEnabled(by_id(made, wid).isEnabled())
        ucol = on and bool(it.get("UseCoolbarColors", True))
        for wid in (16394, 16393):
            by_id(made, wid).setEnabled(on and not ucol)
            labels[wid].setEnabled(on and not ucol)
        for wid, name in colors.items():
            from .propwin import qcolor
            v = it.get(name) if on else None
            by_id(made, wid).setStyleSheet(f"background: {qcolor(int(v)).name()}" if v is not None else "")
        by_id(made, 16409).setEnabled(on and ed.cur > 0)            # NewRow: not of the first band
        img = by_id(made, 16405)
        img.setEnabled(on and bool(dlg.pending.get("ImageList", dlg.value("ImageList"))))
        labels[16405].setEnabled(img.isEnabled())

    ed = ListEditor(dlg, key, made, [dict(b) for b in (dlg.value("(Bands)") or [])],
                    {16406: ("Caption", "s"), 16403: ("Width", "w"), 16401: ("MinWidth", "w"),
                     16402: ("MinHeight", "w"), 16400: ("Key", "s"), 16388: ("Tag", "s"), 16405: ("Image", "i"),
                     16389: ("Style", "e"), 16397: ("UseCoolbarPicture", "b"), 16385: ("EmbossPicture", "b"),
                     16395: ("UseCoolbarColors", "b"), 16408: ("Visible", "b"), 16407: ("AllowVertical", "b"),
                     16409: ("NewRow", "b"), 16399: ("FixedBackground", "b")},
                    "Band", "(Bands)", idx=16412, scroll=None, ins=16411, rem=16410, after=after, raw=False)
    ed.new = lambda: dict(VP.BAND_DEFAULT)
    child.activated.connect(lambda i: ed.edit("Child", "" if i == 0 else child.itemText(i)))
    left.clicked.connect(lambda: ed.show(ed.cur - 1))
    right.clicked.connect(lambda: ed.show(ed.cur + 1))
    for wid, name in colors.items():
        def pick(_=False, name=name):
            c = QColorDialog.getColor(QColor("white"), dlg)
            if c.isValid():
                ed.edit(name, c.red() | (c.green() << 8) | (c.blue() << 16))
        by_id(made, wid).clicked.connect(pick)

    def browse():
        path, _ = QFileDialog.getOpenFileName(dlg, "Load Picture", "", "Pictures (*.bmp *.dib *.gif *.jpg *.ico *.cur "
                                              "*.wmf *.emf);;All Files (*)")
        if path and ed.cur >= 0:
            from lbvc.project.frx import FrxData
            with open(path, "rb") as f:
                ed.items[ed.cur]["Picture"] = FrxData("picture", f.read())
            ed.changed()
    by_id(made, 16398).clicked.connect(browse)
    ed.show(ed.cur)


def data_fields(dlg):
    """the fields of the ADO Data control the control names in DataSource (as the Properties window finds them for
    Retrieve Fields), or []"""
    import os
    import re
    from . import dbinfo
    from . import vbprops as VP
    d = dlg.designer
    root = d.root if d is not None else None
    src = str(dlg.value("DataSource") or "") if VP.prop_def(dlg.type_name, "DataSource") is not None else ""
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
    mod = d.module
    path = getattr(mod, "path", None) or getattr(getattr(mod, "project", None), "path", None)
    if m and path and not os.path.isabs(m.group(1).strip()):     # a database next to the project
        full = os.path.join(os.path.dirname(path), m.group(1).strip())
        if os.path.exists(full):
            cs = cs[:m.start(1)] + full + cs[m.end(1):]
    try:
        return [f[0] for f in dbinfo.ado_fields(cs, int(VP._num(ds.get("CommandType", 8))), source)]
    except Exception:                           # noqa: BLE001 - no database here
        return []


def repeater_page(dlg, key, w, made):
    """the RepeaterBindings of a DataRepeater: PropertyName (a bindable property of its repeated control), DataField,
    Key; Add, Remove, Clear; the list (vbprops.drep_bindings)"""
    import os
    from .ocxpages import by_id
    from PyQt6.QtWidgets import QTableWidgetItem, QHeaderView
    from .drepbinds import repeated_controls, bindable_props
    from . import vbprops as VP
    pn, df, ky, count, table = (by_id(made, i) for i in (1008, 1009, 1018, 1016, 1010))
    add, rem, clr = by_id(made, 1012), by_id(made, 1013), by_id(made, 1017)
    mod = dlg.designer.module if dlg.designer is not None else None
    path = getattr(mod, "path", None) or getattr(getattr(mod, "project", None), "path", None)
    name = VP.drep_name(dlg.ctl)
    ctl_path = next((v for k, v in repeated_controls(os.path.dirname(path) if path else None).items()
                     if k.lower() == name.lower()), None) if name else None
    for cb, items in ((pn, bindable_props(ctl_path) if ctl_path else []), (df, data_fields(dlg))):
        cb.clear()
        cb.setEditable(True)                    # (a name the lists here do not have can be typed)
        cb.addItems(items)
        cb.setEditText("")
    table.setColumnCount(3)
    table.setHorizontalHeaderLabels(["PropertyName", "DataField", "Key"])
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
    table.verticalHeader().hide()
    from .ocxgrids import Grid_binds
    shared = Grid_binds.of(dlg)                 # (the Format page has the same bindings)
    binds = shared.binds

    def fill(sel=-1):
        table.setRowCount(0)
        for b in binds:
            r = table.rowCount()
            table.insertRow(r)
            for c, k in enumerate(("PropertyName", "DataField", "Key")):
                table.setItem(r, c, QTableWidgetItem(str(b.get(k, ""))))
        if 0 <= sel < len(binds):
            table.selectRow(sel)
        count.setText(f"Repeater&Bindings({len(binds)})")
        clr.setEnabled(bool(binds))
        rem.setEnabled(0 <= sel < len(binds))
        state()

    def state(*_):
        add.setEnabled(bool(pn.currentText().strip()) and bool(df.currentText().strip()))

    def changed():
        shared.changed()

    def do_add():
        p = pn.currentText().strip()
        b = {"PropertyName": p, "DataField": df.currentText().strip(), "Key": ky.text().strip()}
        i = next((j for j, x in enumerate(binds) if x["PropertyName"].lower() == p.lower()), None)
        if i is None:                           # (one binding for each property)
            binds.append(b)
            i = len(binds) - 1
        else:
            binds[i] = b
        fill(i)
        changed()

    def do_remove():
        r = table.currentRow()
        if 0 <= r < len(binds):
            del binds[r]
            fill(min(r, len(binds) - 1))
            changed()

    def do_clear():
        binds.clear()
        fill()
        changed()

    def picked():
        r = table.currentRow()
        rem.setEnabled(0 <= r < len(binds))
        if 0 <= r < len(binds):
            pn.setEditText(binds[r]["PropertyName"])
            df.setEditText(binds[r]["DataField"])
            ky.setText(binds[r].get("Key", ""))
    pn.editTextChanged.connect(state)
    df.editTextChanged.connect(state)
    add.clicked.connect(do_add)
    rem.clicked.connect(do_remove)
    clr.clicked.connect(do_clear)
    table.itemSelectionChanged.connect(picked)
    shared.refresh.append(lambda: fill(table.currentRow()))
    fill()
    dlg.bound[key] = {"(RepeaterBindings)": table}
    dlg.unbound[key] = []


PAGES = {("TabStrip", "Tabs"): tabs_page, ("StatusBar", "Panels"): panels_page, ("Toolbar", "Buttons"): buttons_page,
         ("ListView", "Column Headers"): columns_page, ("ImageList", "Images"): images_page,
         ("CoolBar", "Bands"): bands_page, ("DataRepeater", "RepeaterBindings"): repeater_page}
