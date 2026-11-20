"""The form designer: a form drawn as it looks when it runs, with the controls of the
Toolbox, selection handles, moving and sizing on the grid, copy/paste and control arrays.

The designer edits the designer section of the .frm itself (the Control tree of
lbvc.project.source), so what it saves is what VB6 reads; properties it does not touch are
written back exactly as they were. Sizes are in twips (15 per pixel), the grid is VB6's
120 x 120 twips."""
import copy
import os

from PyQt6.QtCore import Qt, QPoint, QRect, QSize, pyqtSignal, QMimeData, QTimer
from PyQt6.QtGui import QPainter, QColor, QPen, QBrush, QFont, QPalette, QPolygon, QKeySequence
from PyQt6.QtWidgets import (QWidget, QScrollArea, QPushButton, QLabel, QLineEdit, QPlainTextEdit, QCheckBox,
                             QRadioButton, QComboBox, QListWidget, QGroupBox, QFrame, QScrollBar, QApplication,
                             QMessageBox, QRubberBand, QToolButton, QGridLayout, QSizePolicy, QStyleFactory,
                             QVBoxLayout)

from lbvc.project.source import Control, Property, PropertyBlock, parse_source
from . import vbprops as V
from . import ocxlib
from . import formops as F

TPP = 15                    # twips per pixel
GRID = 120                  # twips: the grid of VB6 (Tools > Options > General: Width, Height)

# ---------------------------------------------------------------- the intrinsic controls
# class -> (base name of new controls, default width, default height) in twips
TOOLS = [
    ("PictureBox", "Picture", 1215, 1215), ("Label", "Label", 1215, 495), ("TextBox", "Text", 1215, 495),
    ("Frame", "Frame", 1695, 1215), ("CommandButton", "Command", 1215, 495), ("CheckBox", "Check", 1215, 495),
    ("OptionButton", "Option", 1215, 495), ("ComboBox", "Combo", 1215, 315), ("ListBox", "List", 1215, 1035),
    ("HScrollBar", "HScroll", 1215, 255), ("VScrollBar", "VScroll", 255, 1215), ("Timer", "Timer", 420, 420),
    ("DriveListBox", "Drive", 1215, 315), ("DirListBox", "Dir", 1215, 1215), ("FileListBox", "File", 1215, 1215),
    ("Shape", "Shape", 1215, 1215), ("Line", "Line", 1215, 0), ("Image", "Image", 1215, 1215),
    ("Data", "Data", 1140, 345),
]
TOOL = {t[0]: t for t in TOOLS}
CONTAINERS = ("Frame", "PictureBox", "SSTab", "CoolBar")
# controls of an .ocx that have no window when the program runs: a fixed icon on the form
FIXED = ("Timer", "CommonDialog", "ImageList", "MSComm", "Winsock", "Inet", "MAPISession", "MAPIMessages")


# the user controls the forms can have: of the project, and of the control projects of the group
# lower "project.name" -> (project name, user control name, width, height, module)
USER_CONTROLS = {}
_UC_SHOTS = {}                                   # id(module) -> the picture of its designer
_UC_MAKING = set()


def set_user_controls(entries):
    """[(project name, name, width, height, module)]: what the Toolbox offers and the forms show"""
    USER_CONTROLS.clear()
    V.DYNAMIC.clear()
    for proj, name, w, h, m in entries:
        USER_CONTROLS[f"{proj}.{name}".lower()] = (proj, name, w, h, m)
        V.DYNAMIC[name.lower()] = V.uctl_props(m.code_text())


def uctl_of(cls):
    return USER_CONTROLS.get((cls or "").lower())


def forget_uctl_look(module):
    """the user control changed in its designer: the forms draw it again"""
    _UC_SHOTS.pop(id(module), None)


def uctl_look(main, module):
    """the user control as its designer shows it (without the grid): a picture, made once"""
    pm = _UC_SHOTS.get(id(module))
    if pm is not None:
        return pm
    if id(module) in _UC_MAKING:                # a user control inside itself: VB6 does not allow it
        raise RuntimeError("a user control can't contain itself")
    _UC_MAKING.add(id(module))
    try:
        d = FormDesigner(main, module)
    finally:
        _UC_MAKING.discard(id(module))
    d.show_grid = False
    d.surface.resize(geti(module.sf.root, "ClientWidth", 4800) // TPP, geti(module.sf.root, "ClientHeight", 3600) // TPP)
    pm = d.surface.grab()
    d.deleteLater()
    _UC_SHOTS[id(module)] = pm
    return pm


def tool_spec(cls):
    """(class, base name, width, height) of a Toolbox entry: an intrinsic class (TextBox) or the
    class of an .ocx control with its library (MSComctlLib.TreeView)"""
    if cls in TOOL:
        return TOOL[cls]
    uc = uctl_of(cls)
    if uc:                                       # the size of its designer now
        root = uc[4].sf.root if uc[4].sf is not None and uc[4].sf.root is not None else None
        w = geti(root, "ClientWidth", uc[2]) if root is not None else uc[2]
        h = geti(root, "ClientHeight", uc[3]) if root is not None else uc[3]
        return (f"{uc[0]}.{uc[1]}", uc[1], w, h)
    lc = ocxlib.control_of(cls)
    if lc:
        return (lc[1].cls, lc[1].base, lc[1].w, lc[1].h)
    return (cls, cls, 1215, 495)
NO_TABINDEX = ("Timer", "Shape", "Line", "Image", "Menu", "Data")
CAPTIONED = ("Label", "Frame", "CommandButton", "CheckBox", "OptionButton", "Data")
TEXTED = ("TextBox", "ComboBox")
# the windowless controls: VB6 keeps them under the windowed ones and writes them after them
LIGHT = ("Label", "Image", "Line", "Shape")


def item_height_px(ctl):
    """the height of a list item: tmHeight of the font (13 px for MS Sans Serif 8, 14 for
    Tahoma 9); a control without a font of its own has the font of the form"""
    blk = ctl.block("Font")
    up = ctl.parent
    while blk is None and up is not None:
        if up.parent is None:
            blk = up.block("Font")
        up = up.parent
    if blk is None:
        return 13
    try:
        return max(1, round(float(blk.get("Size", 8.25)) * 4 / 3 * 1.2))
    except (TypeError, ValueError):
        return 13


def integral_height(ctl, kind, h):
    """ListBox (IntegralHeight True) and FileListBox: VB6 makes the height whole items plus the
    border (VB6 oracle d_all: ListBox 1080 -> 1035, FileListBox 1200 -> 1065)"""
    if kind == "ListBox":
        if geti(ctl, "IntegralHeight", -1) == 0:
            return h
        border = 4 if geti(ctl, "Appearance", 1) else 2
    elif kind == "FileListBox":
        border = 6 if geti(ctl, "Appearance", 1) else 4
    else:
        return h
    item = item_height_px(ctl)
    n = max(1, (int(h) // TPP - border) // item)
    return (n * item + border) * TPP


def normalize_children(parent):
    """windowed controls, then the windowless ones, then the menus (each in its own order)"""
    kids = parent.children
    win = [x for x in kids if x.type_name != "Menu" and x.type_name not in LIGHT]
    light = [x for x in kids if x.type_name in LIGHT]
    menus = [x for x in kids if x.type_name == "Menu"]
    parent.children = win + light + menus
    props = [(k, x) for k, x in parent.order if k == "p"]
    parent.order = props + [("c", x) for x in parent.children]


def vb_color(v, default):
    """a VB6 color (&HBBGGRR&, &H800000nn& system color) as QColor"""
    try:
        c = int(v) if v is not None else default
    except (TypeError, ValueError):
        c = default
    c &= 0xFFFFFFFF
    if c & 0x80000000:
        sysc = {0x0F: "#F0F0F0", 0x05: "#FFFFFF", 0x08: "#000000", 0x12: "#000000", 0x0C: "#ABABAB",
                0x0D: "#3399FF", 0x0E: "#FFFFFF", 0x10: "#A0A0A0", 0x11: "#6D6D6D", 0x14: "#FFFFFF",
                0x02: "#99B4D1", 0x09: "#000000", 0x01: "#000000", 0x04: "#F0F0F0", 0x07: "#000000"}
        return QColor(sysc.get(c & 0xFF, "#F0F0F0"))
    return QColor(c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF)


def ctl_font(ctl, default=None):
    """the font of a control from BeginProperty Font (MS Sans Serif 8.25 when there is none)"""
    blk = ctl.block("Font") if ctl is not None else None
    if blk is None:
        return QFont(default) if default is not None else _default_font()
    f = QFont(str(blk.get("Name", "MS Sans Serif")))
    try:
        f.setPixelSize(max(1, round(float(blk.get("Size", 8.25)) * 4 / 3)))
    except (TypeError, ValueError):
        f.setPixelSize(11)
    f.setBold(int(blk.get("Weight", 400) or 400) >= 600)
    f.setItalic(bool(blk.get("Italic", 0)))
    f.setUnderline(bool(blk.get("Underline", 0)))
    f.setStrikeOut(bool(blk.get("Strikethrough", 0)))
    return f


_LIGHT_STYLE = None


def light_palette():
    """the colours the designer draws forms and controls with: always the light ones of Windows (as VB6 shows
    them), whatever the theme of the IDE or of the system; the program follows the system when it runs"""
    pal = QPalette()
    R, G = QPalette.ColorRole, QPalette.ColorGroup
    for role, c in ((R.Window, "#F0F0F0"), (R.WindowText, "#000000"), (R.Base, "#FFFFFF"),
                    (R.AlternateBase, "#F7F7F7"), (R.Text, "#000000"), (R.Button, "#F0F0F0"),
                    (R.ButtonText, "#000000"), (R.BrightText, "#FFFFFF"), (R.Light, "#FFFFFF"),
                    (R.Midlight, "#E3E3E3"), (R.Mid, "#A0A0A0"), (R.Dark, "#A0A0A0"), (R.Shadow, "#696969"),
                    (R.Highlight, "#3399FF"), (R.HighlightedText, "#FFFFFF"), (R.ToolTipBase, "#FFFFE1"),
                    (R.ToolTipText, "#000000"), (R.PlaceholderText, "#808080"), (R.Link, "#0066CC")):
        pal.setColor(role, QColor(c))
    for role in (R.WindowText, R.Text, R.ButtonText):
        pal.setColor(G.Disabled, role, QColor("#6D6D6D"))
    return pal


def light_look(w):
    """a control of the designer (and the widgets inside it) with the light look: the Fusion style, which
    draws with the palette, so that a dark theme of the system does not reach the form"""
    global _LIGHT_STYLE
    if _LIGHT_STYLE is None:
        _LIGHT_STYLE = QStyleFactory.create("Fusion")
    if _LIGHT_STYLE is None:
        return
    for x in [w] + w.findChildren(QWidget):
        x.setStyle(_LIGHT_STYLE)


def _default_font():
    f = QFont(QApplication.font())
    f.setPixelSize(11)
    return f


def set_prop(ctl, name, value):
    """set a property where VB6 writes it: properties in alphabetical order, then the controls"""
    p = ctl.prop(name)
    if p is not None:
        p.value = value
        return
    p = Property(name, value)
    ctl.items.append(p)
    # _ExtentX, _ExtentY, _Version of an .ocx come after the properties of the extender (Width ...)
    sortkey = lambda n: "~" + n.lower() if n.startswith("_") else n.lower()   # noqa: E731
    key = sortkey(name)
    pos = None
    for i, (k, it) in enumerate(ctl.order):
        if k == "c":
            pos = i
            break
        if sortkey(it.name) > key:
            pos = i
            break
    ctl.order.insert(len(ctl.order) if pos is None else pos, ("p", p))


def append_prop(ctl, name, value):
    """a property of an .ocx control itself: after the others, in the order the control writes them"""
    from . import vbprops
    if vbprops._bag_of(ctl):
        vbprops.put(ctl, name, value)
        return
    p = Property(name, value)
    ctl.items.append(p)
    pos = next((i for i, (k, _) in enumerate(ctl.order) if k == "c"), len(ctl.order))
    ctl.order.insert(pos, ("p", p))


def del_prop(ctl, name):
    p = ctl.prop(name)
    if p is None:
        return
    ctl.items.remove(p)
    ctl.order = [(k, it) for k, it in ctl.order if it is not p]


def geti(ctl, name, default=0):
    v = ctl.get(name, default)
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return default


def strip_amp(s):
    out, i = [], 0
    while i < len(s):
        if s[i] == "&" and i + 1 < len(s):
            out.append(s[i + 1])
            i += 2
            continue
        out.append(s[i])
        i += 1
    return "".join(out)


# ---------------------------------------------------------------- the look of the .ocx controls
def mci_glyph(p, b, color):
    """the picture of a button of the Multimedia control, in a 16x16 box"""
    p.setPen(Qt.PenStyle.NoPen)
    p.setBrush(QColor("#A00000") if b == "Record" else color)

    def tri(x, y, w, h, right):
        pts = [QPoint(x, y), QPoint(x + w, y + h // 2), QPoint(x, y + h)] if right else \
              [QPoint(x + w, y), QPoint(x, y + h // 2), QPoint(x + w, y + h)]
        p.drawPolygon(QPolygon(pts))
    if b == "Prev":
        p.drawRect(2, 4, 2, 9); tri(4, 4, 5, 9, False); tri(9, 4, 5, 9, False)
    elif b == "Next":
        tri(2, 4, 5, 9, True); tri(7, 4, 5, 9, True); p.drawRect(12, 4, 2, 9)
    elif b == "Play":
        tri(5, 3, 7, 11, True)
    elif b == "Pause":
        p.drawRect(4, 4, 3, 9); p.drawRect(9, 4, 3, 9)
    elif b == "Back":
        p.drawRect(3, 4, 2, 9); tri(6, 4, 6, 9, False)
    elif b == "Step":
        tri(4, 4, 6, 9, True); p.drawRect(11, 4, 2, 9)
    elif b == "Stop":
        p.drawRect(4, 4, 8, 8)
    elif b == "Record":
        p.drawEllipse(4, 4, 8, 8)
    elif b == "Eject":
        p.drawPolygon(QPolygon([QPoint(8, 3), QPoint(13, 9), QPoint(3, 9)])); p.drawRect(3, 11, 11, 2)


class OcxLook(QWidget):
    """a control of an .ocx at design time, drawn as Windows draws it with its default properties"""

    def __init__(self, ctl, kind, val, images=None):
        super().__init__()
        self.ctl, self.kind, self.val = ctl, kind, val
        self.images = images                    # name of an ImageList -> ({key: QPixmap}, [QPixmap])

    def paintEvent(self, e):
        p = QPainter(self)
        k, c, r = self.kind, self.ctl, self.rect().adjusted(0, 0, -1, -1)
        face, white, dark, blue = QColor("#F0F0F0"), QColor("#FFFFFF"), QColor("#808080"), QColor("#4A90D9")
        p.setFont(self.font())
        fm = p.fontMetrics()

        def sunken(rect, fill=white):
            p.fillRect(rect, fill)
            p.setPen(dark)
            p.drawLine(rect.topLeft(), rect.topRight()); p.drawLine(rect.topLeft(), rect.bottomLeft())
            p.setPen(white if fill != white else QColor("#E0E0E0"))
            p.drawLine(rect.bottomLeft(), rect.bottomRight()); p.drawLine(rect.topRight(), rect.bottomRight())

        if k in ("TreeView", "ListView", "RichTextBox", "Animation"):
            sunken(r)
            if k == "RichTextBox":
                p.setPen(QColor("#000000"))
                p.drawText(r.adjusted(3, 2, -2, -2), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop, c.name)
        elif k == "ProgressBar":
            sunken(r, face)
        elif k == "Slider":
            p.fillRect(r, face)
            y = min(r.height() // 2, 12)
            sunken(QRect(8, y - 2, r.width() - 16, 4), white)
            p.setPen(QColor("#000000"))
            ticks = max(1, int(geti(c, "Max", 10) or 10))
            for i in range(ticks + 1):
                x = 10 + (r.width() - 20) * i // ticks
                p.drawLine(x, y + 10, x, y + 13)
            p.setBrush(face); p.setPen(dark)
            p.drawPolygon(QPolygon([QPoint(6, y - 7), QPoint(14, y - 7), QPoint(14, y + 4), QPoint(10, y + 8), QPoint(6, y + 4)]))
        elif k == "StatusBar":
            p.fillRect(r, face)
            sunken(QRect(2, 2, r.width() - 18, r.height() - 3), face)
            p.setPen(dark)
            for i in range(3):
                p.drawLine(r.right() - 2 - 4 * i, r.bottom(), r.right(), r.bottom() - 2 - 4 * i)
        elif k == "Toolbar":
            p.fillRect(r, face)
            p.setPen(white); p.drawLine(r.topLeft(), r.topRight())
            p.setPen(dark); p.drawLine(r.bottomLeft(), r.bottomRight())
            # its buttons as the VB6 IDE shows them (DEV-9/aw10-02): the pictures of the ImageList, the separators
            keyed, listed = self.images(str(c.get("ImageList", "") or "")) if self.images else ({}, [])
            hm = lambda v, d: max(4, int(round(float(v) * 1440 / 2540 / 15))) if v else d   # noqa: E731
            bw, bh = hm(c.get("ButtonWidth"), 23), hm(c.get("ButtonHeight"), 22)
            bh = min(bh, max(4, r.height() - 3))
            flat = geti(c, "Style", 0) == 1
            x = 2
            for b in toolbar_buttons(c):
                st = b.get("style", 0)
                if st == 3:                     # a separator
                    p.setPen(dark); p.drawLine(x + 3, 3, x + 3, 3 + bh - 4)
                    p.setPen(white); p.drawLine(x + 4, 3, x + 4, 3 + bh - 4)
                    x += 8
                    continue
                if st == 4:                     # a placeholder: its Width
                    x += hm(b.get("width"), bw) if b.get("width") else bw
                    continue
                br = QRect(x, 2, bw, bh)
                if not flat:
                    p.setPen(white); p.drawLine(br.topLeft(), br.topRight()); p.drawLine(br.topLeft(), br.bottomLeft())
                    p.setPen(dark); p.drawLine(br.bottomLeft(), br.bottomRight()); p.drawLine(br.topRight(), br.bottomRight())
                img = b.get("image")
                pm = keyed.get(str(img).lower()) if isinstance(img, str) else \
                    (listed[img - 1] if isinstance(img, int) and 0 < img <= len(listed) else None)
                cap = b.get("caption", "")
                if pm is not None:
                    iy = br.top() + (br.height() - pm.height()) // 2 if not cap else br.top() + 2
                    p.drawPixmap(br.left() + (br.width() - pm.width()) // 2, iy, pm)
                if cap:
                    p.setPen(QColor("#000000"))
                    p.drawText(br.adjusted(1, 0, -1, -1), Qt.AlignmentFlag.AlignHCenter |
                               (Qt.AlignmentFlag.AlignBottom if pm is not None else Qt.AlignmentFlag.AlignVCenter), cap)
                x += bw
                if x > r.width():
                    break
        elif k == "CoolBar":                    # its bands as the runtime lays them out (vbprops.cool_layout)
            p.fillRect(r, face)
            bands = V.cool_bands(c)
            kids = V.cool_children(c)
            borders = bool(geti(c, "BandBorders", -1))
            geo, _ = V.cool_layout(bands, self.width(), bool(geti(c, "FixedOrder", 0)), borders,
                                   bool(geti(c, "VariantHeight", -1)),
                                   has_child=lambda b: str(b.get("Child", "") or "").lower() in kids)
            for b, g in zip(bands, geo):
                if g is None:
                    continue
                x, y, w, h, hd, grip = g
                if grip:
                    for gx in (x + 2, x + 5):
                        p.setPen(white); p.drawLine(gx, y + 2, gx, y + h - 3)
                        p.setPen(dark); p.drawLine(gx + 1, y + 2, gx + 1, y + h - 3)
                cap = str(b.get("Caption", "") or "")
                if cap:
                    p.setPen(QColor("#000000"))
                    p.drawText(QRect(x + (9 if grip else 4), y, max(0, hd), h),
                               Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, cap)
                if borders:
                    p.setPen(dark); p.drawLine(x + w, y, x + w, y + h - 1)
                    p.setPen(white); p.drawLine(x + w + 1, y, x + w + 1, y + h - 1)
                    p.setPen(dark); p.drawLine(x, y + h, x + w + 1, y + h)
                    p.setPen(white); p.drawLine(x, y + h + 1, x + w + 1, y + h + 1)
        elif k in ("TabStrip", "SSTab"):
            p.fillRect(r, face)
            th = int(geti(c, "TabHeight", 520)) // TPP if k == "SSTab" else fm.height() + 8
            n = int(geti(c, "Tabs", 3)) if k == "SSTab" else 1
            per = max(1, int(geti(c, "TabsPerRow", 3))) if k == "SSTab" else 1
            tw = max(30, r.width() // per) if k == "SSTab" else fm.horizontalAdvance("Tab 1") + 16
            p.setPen(dark); p.drawRect(QRect(0, th, r.width(), r.height() - th))
            for i in range(min(n, per)):
                tr = QRect(i * tw, 2 if i else 0, tw - 1, th - (2 if i else 0))
                p.fillRect(tr, face); p.setPen(white); p.drawLine(tr.topLeft(), tr.topRight()); p.drawLine(tr.topLeft(), tr.bottomLeft())
                p.setPen(dark); p.drawLine(tr.topRight(), tr.bottomRight())
                cap = str(c.get(f"TabCaption({i})", f"Tab {i}")) if k == "SSTab" else "Tab 1"
                p.setPen(QColor("#000000")); p.drawText(tr, Qt.AlignmentFlag.AlignCenter, cap)
        elif k in ("ImageCombo", "DTPicker"):
            sunken(r)
            b = QRect(r.right() - 16, 2, 15, r.height() - 3)
            p.fillRect(b, face); p.setPen(dark); p.drawRect(b)
            p.setBrush(QColor("#000000"))
            cx, cy = b.center().x(), b.center().y()
            p.drawPolygon(QPolygon([QPoint(cx - 3, cy - 1), QPoint(cx + 3, cy - 1), QPoint(cx, cy + 2)]))
            if k == "DTPicker":
                from PyQt6.QtCore import QDate
                p.setPen(QColor("#000000"))
                p.drawText(r.adjusted(4, 0, -18, 0), Qt.AlignmentFlag.AlignVCenter, QDate.currentDate().toString("d/M/yyyy"))
        elif k == "UpDown":
            h = r.height() // 2
            for y0, up in ((0, True), (h, False)):
                br = QRect(0, y0, r.width(), h)
                p.fillRect(br, face); p.setPen(dark); p.drawRect(br)
                cx, cy = br.center().x(), br.center().y()
                p.setBrush(QColor("#000000"))
                p.drawPolygon(QPolygon([QPoint(cx - 3, cy + (1 if up else -1)), QPoint(cx + 3, cy + (1 if up else -1)),
                                        QPoint(cx, cy + (-2 if up else 2))]))
        elif k == "FlatScrollBar":
            p.fillRect(r, QColor("#E8E8E8"))
            s = min(r.width(), r.height())
            p.fillRect(QRect(0, 0, s, s), face); p.fillRect(QRect(r.width() - s, 0, s, s), face)
            p.setPen(dark); p.drawRect(r)
        elif k == "MonthView":
            sunken(r)
            p.fillRect(QRect(2, 2, r.width() - 3, fm.height() + 6), blue)
            from PyQt6.QtCore import QDate
            d = QDate.currentDate()
            p.setPen(white); p.drawText(QRect(2, 2, r.width() - 3, fm.height() + 6), Qt.AlignmentFlag.AlignCenter, d.toString("MMMM yyyy"))
            p.setPen(QColor("#000000"))
            cw, top = max(10, (r.width() - 8) // 7), fm.height() + 12
            first = QDate(d.year(), d.month(), 1).dayOfWeek() % 7
            for day in range(1, d.daysInMonth() + 1):
                i = first + day - 1
                p.drawText(QRect(4 + (i % 7) * cw, top + (i // 7) * (fm.height() + 2), cw, fm.height()), Qt.AlignmentFlag.AlignCenter, str(day))
        elif k == "MSFlexGrid":
            sunken(r)
            rows, cols = int(geti(c, "Rows", 2)), int(geti(c, "Cols", 2))
            cw, rh = 64, fm.height() + 4
            for i in range(rows):
                for j in range(cols):
                    cell = QRect(2 + j * cw, 2 + i * rh, cw, rh)
                    if i < int(geti(c, "FixedRows", 1)) or j < int(geti(c, "FixedCols", 1)):
                        p.fillRect(cell, face)
                    p.setPen(QColor("#C0C0C0")); p.drawRect(cell)
        elif k == "MSChart":
            p.fillRect(r, white)
            pa = r.adjusted(30, 10, -10, -24)
            p.setPen(QColor("#C0C0C0"))
            for i in range(5):
                y = pa.bottom() - pa.height() * i // 4
                p.drawLine(pa.left(), y, pa.right(), y)
            p.setPen(QColor("#000000")); p.drawLine(pa.bottomLeft(), pa.bottomRight()); p.drawLine(pa.topLeft(), pa.bottomLeft())
            sample = ((20, 25, 30, 35), (40, 45, 50, 55), (30, 35, 40, 45), (60, 65, 70, 75), (50, 55, 60, 65))
            cols = (QColor(0, 0, 255), QColor(255, 0, 0), QColor(0, 255, 0), QColor(255, 255, 0))
            gw = pa.width() / 5
            for i, row in enumerate(sample):
                bw = gw * 0.8 / 4
                for j, v in enumerate(row):
                    h = pa.height() * v / 80
                    p.fillRect(QRect(int(pa.left() + i * gw + gw * 0.1 + j * bw), int(pa.bottom() - h), max(1, int(bw)), int(h)), cols[j])
                p.drawText(QRect(int(pa.left() + i * gw), pa.bottom() + 2, int(gw), fm.height()), Qt.AlignmentFlag.AlignCenter, f"R{i + 1}")
        elif k == "PictureClip":
            p.fillRect(r, face)
            p.setPen(QPen(dark, 1, Qt.PenStyle.DashLine)); p.drawRect(r)
        elif k == "MMControl":                  # the nine buttons of the Multimedia control
            names = [b for b in ("Prev", "Next", "Play", "Pause", "Back", "Step", "Stop", "Record", "Eject")
                     if str(c.get(f"{b}Visible", True)) not in ("False", "0")]
            vert = str(c.get("Orientation", 0)) == "1"
            total = r.height() + 1 if vert else r.width() + 1
            for i, b in enumerate(names):
                a, z = total * i // len(names), total * (i + 1) // len(names)
                br = QRect(0, a, r.width() + 1, z - a) if vert else QRect(a, 0, z - a, r.height() + 1)
                p.fillRect(br, face)
                p.setPen(white); p.drawLine(br.topLeft(), br.topRight()); p.drawLine(br.topLeft(), br.bottomLeft())
                p.setPen(dark); p.drawLine(br.bottomLeft(), br.bottomRight()); p.drawLine(br.topRight(), br.bottomRight())
                p.save(); p.translate(br.center().x() - 8, br.center().y() - 8)
                mci_glyph(p, b, QColor("#808080"))
                p.restore()
        elif k in ("Adodc", "MSRDC"):           # the ADO Data control, the RemoteData control: as the Data control
            DataLook.paint_on(p, self.rect(), str(c.get("Caption", c.name)), white, QColor("#000000"))
        elif k == "DataRepeater":               # the caption, the row indicators, the rows of the repeated control
            sunken(r)
            top = 2
            if str(c.get("CaptionStyle", 1)) != "0":
                hh = fm.height() + 4
                p.fillRect(QRect(2, 2, r.width() - 3, hh), face)
                p.setPen(QColor("#000000"))
                p.drawText(QRect(6, 2, r.width() - 10, hh), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
                           str(c.get("Caption", c.name)))
                top += hh
            if str(c.get("RowIndicator", -1)) not in ("0", "False"):
                p.fillRect(QRect(2, top, 14, r.height() - top - 1), face)
            name = V.drep_name(c)
            p.setPen(QColor("#808080"))
            y = top + 40
            while y < r.height() - 2:
                p.drawLine(2, y, r.width() - 2, y)
                y += 40
            if name:
                p.drawText(QRect(20, top, r.width() - 24, 40), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, name)
        elif k in ("DataGrid", "DBGrid"):
            sunken(r)
            hh = fm.height() + 4
            p.fillRect(QRect(2, 2, r.width() - 3, hh), face)
            p.fillRect(QRect(2, 2, 14, r.height() - 3), face)
            p.setPen(QColor("#C0C0C0"))
            for j in range(2):
                p.drawLine(16 + (j + 1) * 90, 2, 16 + (j + 1) * 90, r.height() - 2)
            for i in range(2):
                p.drawLine(2, 2 + hh * (i + 1), r.width() - 2, 2 + hh * (i + 1))
        elif k == "WebBrowser":                 # a white page (VB6 shows a page of Internet Explorer there)
            sunken(r)
            p.fillRect(QRect(2, 2, r.width() - 3, r.height() - 3), white)
        elif k == "MaskEdBox":                  # the mask with its prompts, as the control shows it
            sunken(r)
            m, pc, shown, i = str(c.get("Mask", "")), (str(c.get("PromptChar", "_")) or "_")[:1], "", 0
            while i < len(m):
                ch = m[i]
                if ch == "\\" and i + 1 < len(m):
                    shown += m[i + 1]; i += 2; continue
                if ch not in "<>":
                    shown += pc if ch in "#9?Aa&C" else ch
                i += 1
            p.setPen(QColor("#000000")); p.drawText(r.adjusted(3, 0, -2, 0), Qt.AlignmentFlag.AlignVCenter, shown)
        elif k in ("DataList", "DataCombo", "DBList", "DBCombo"):
            sunken(r)
            if k in ("DataCombo", "DBCombo"):
                b = QRect(r.right() - 16, 2, 15, r.height() - 3)
                p.fillRect(b, face); p.setPen(dark); p.drawRect(b)
                p.setPen(QColor("#000000")); p.drawText(r.adjusted(4, 0, -18, 0), Qt.AlignmentFlag.AlignVCenter, str(c.get("Text", c.name)))
            else:
                p.setPen(QColor("#000000")); p.drawText(r.adjusted(4, 2, 0, 0), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop, c.name)
        else:
            p.fillRect(r, face)
            p.setPen(dark); p.drawRect(r)
            p.drawText(r, Qt.AlignmentFlag.AlignCenter, k)
        p.end()


class DataLook(QWidget):
    """the Data control: |< and < on the left, > and >| on the right, the Caption between them"""

    def __init__(self, caption):
        super().__init__()
        self.caption = caption

    def paintEvent(self, e):
        p = QPainter(self)
        p.setFont(self.font())
        self.paint_on(p, self.rect(), self.caption, self.palette().color(QPalette.ColorRole.Base),
                      self.palette().color(QPalette.ColorRole.Text))
        p.end()

    @staticmethod
    def paint_on(p, r, caption, base, text):
        p.fillRect(r, base)
        b = max(8, min(r.height(), r.width() // 4))
        face = QColor("#F0F0F0")
        for i, x in enumerate((0, b, r.width() - 2 * b, r.width() - b)):
            br = QRect(x, 0, b, r.height())
            p.fillRect(br, face)
            p.setPen(QColor("#FFFFFF"))
            p.drawLine(br.left(), br.top(), br.right() - 1, br.top())
            p.drawLine(br.left(), br.top(), br.left(), br.bottom() - 1)
            p.setPen(QColor("#808080"))
            p.drawLine(br.left(), br.bottom(), br.right(), br.bottom())
            p.drawLine(br.right(), br.top(), br.right(), br.bottom())
            cx, cy, a = br.center().x(), br.center().y(), max(2, min(4, b // 5))
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor("#000000"))
            if i < 2:
                tri = QPolygon([QPoint(cx - a // 2 - 1, cy), QPoint(cx + a // 2 + 1, cy - a), QPoint(cx + a // 2 + 1, cy + a)])
            else:
                tri = QPolygon([QPoint(cx + a // 2 + 1, cy), QPoint(cx - a // 2 - 1, cy - a), QPoint(cx - a // 2 - 1, cy + a)])
            p.drawPolygon(tri)
            if i in (0, 3):                       # the bar of First / Last
                bx = cx - a // 2 - 3 if i == 0 else cx + a // 2 + 2
                p.fillRect(QRect(bx, cy - a, 1, 2 * a + 1), QColor("#000000"))
        p.setPen(text)
        p.drawText(QRect(2 * b + 3, 0, r.width() - 4 * b - 6, r.height()),
                   Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, caption)


def toolbar_buttons(c):
    """the buttons of a Toolbar from its Buttons block: [{style, key, image (key or index), caption, width}]"""
    out = []
    blk = c.block("Buttons")
    if blk is None:
        return out
    for it in blk.items:
        if not isinstance(it, PropertyBlock) or not it.name.lower().startswith("button"):
            continue
        b = {"style": 0, "image": None, "caption": ""}
        for q in it.items:
            if not isinstance(q, Property):
                continue
            n = q.name.lower().replace("object.", "")
            if n == "style":
                b["style"] = int(q.value) if isinstance(q.value, (int, float)) else 0
            elif n == "key":
                b["key"] = str(q.value)
            elif n in ("imagekey", "image") and str(q.value):
                b["image"] = str(q.value) if not isinstance(q.value, int) else q.value
            elif n == "imageindex" and isinstance(q.value, int) and q.value:
                b["image"] = q.value
            elif n == "caption":
                b["caption"] = str(q.value)
            elif n == "width":
                b["width"] = q.value
        out.append(b)
    return out


def list_images(module, root, name):
    """the pictures of the ImageList <name> of a form: ({key in lower case: QPixmap}, [QPixmap]), masked"""
    from PyQt6.QtGui import QPixmap, QImage
    keyed, listed = {}, []
    il = next((x for x in root.walk() if x.name.lower() == name.lower() and x.type_name == "ImageList"), None) \
        if name else None
    blk = il.block("Images") if il is not None else None
    if blk is None:
        return keyed, listed
    mask = il.get("MaskColor", 0xC0C0C0)
    mask = int(mask) if isinstance(mask, (int, float)) else 0xC0C0C0
    use = il.get("UseMaskColor", -1)
    for it in blk.items:
        if not isinstance(it, PropertyBlock):
            continue
        key, pm = None, None
        for q in it.items:
            if isinstance(q, Property) and q.name.lower() == "key":
                key = str(q.value)
            elif isinstance(q, Property) and q.name.lower() == "picture":
                d = module.frx_data(q)
                img = QImage()
                if d is not None and isinstance(d.data, (bytes, bytearray)) and img.loadFromData(bytes(d.data)):
                    pm = QPixmap.fromImage(img)
                    if use != 0 and not img.hasAlphaChannel():
                        rgb = QColor(mask & 0xFF, (mask >> 8) & 0xFF, (mask >> 16) & 0xFF)
                        pm.setMask(pm.createMaskFromColor(rgb))
        listed.append(pm)
        if key and pm is not None:
            keyed[key.lower()] = pm
    return keyed, listed


def ocx_visual(ctl, kind, val, images=None):
    return OcxLook(ctl, kind, val, images)


# ---------------------------------------------------------------- one control on the form
class DesignItem(QWidget):
    """a control: a real Qt widget for the look (it gets no mouse), the designer handles the rest"""

    def __init__(self, designer, ctl, parent):
        super().__init__(parent)
        self.designer = designer
        self.ctl = ctl
        self.visual = None
        self.setAttribute(Qt.WidgetAttribute.WA_NoMousePropagation)
        self.rebuild()

    @property
    def type(self):
        return self.ctl.type_name

    def rect_twips(self):
        c = self.ctl
        if self.type == "Line":
            x1, y1, x2, y2 = (geti(c, k) for k in ("X1", "Y1", "X2", "Y2"))
            return QRect(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))
        w, h = tool_spec(c.cls if "." in c.cls and not c.cls.lower().startswith("vb.") else self.type)[2:]
        if self.type == "Timer":
            return QRect(geti(c, "Left"), geti(c, "Top"), 420, 420)
        if self.type in FIXED:                   # CommonDialog, ImageList: an icon of 32 x 32 pixels (MSComm 38)
            return QRect(geti(c, "Left"), geti(c, "Top"), w, h)
        return QRect(geti(c, "Left"), geti(c, "Top"), geti(c, "Width", w), geti(c, "Height", h))

    def place(self):
        r = self.rect_twips()
        if self.type == "Line":
            self.setGeometry(r.x() // TPP - 2, r.y() // TPP - 2, max(r.width() // TPP, 1) + 5, max(r.height() // TPP, 1) + 5)
        else:
            self.setGeometry(r.x() // TPP, r.y() // TPP, max(1, r.width() // TPP), max(1, r.height() // TPP))
        if self.visual is not None:
            self.visual.setGeometry(0, 0, self.width(), self.height())

    def rebuild(self):
        """(re)make the widget that shows the control, from its properties"""
        if self.visual is not None:
            self.visual.deleteLater()
            self.visual = None
        c, t = self.ctl, self.type
        val = self.val
        cap = str(val("Caption", c.name))
        w = None
        if t == "CommandButton":
            w = QPushButton(strip_amp(cap))
        elif t == "Label":
            w = QLabel(strip_amp(cap))
            al = geti(c, "Alignment")
            w.setAlignment({1: Qt.AlignmentFlag.AlignRight, 2: Qt.AlignmentFlag.AlignHCenter}.get(al, Qt.AlignmentFlag.AlignLeft) | Qt.AlignmentFlag.AlignTop)
            w.setWordWrap(True)
            if geti(c, "BorderStyle"):
                w.setFrameStyle(QFrame.Shape.Panel | QFrame.Shadow.Sunken)
        elif t == "TextBox":
            if geti(c, "MultiLine"):
                w = QPlainTextEdit(str(val("Text", c.name)).replace("\r\n", "\n"))
            else:
                w = QLineEdit(str(val("Text", c.name)))
                w.setCursorPosition(0)
                if c.get("PasswordChar"):
                    w.setEchoMode(QLineEdit.EchoMode.Password)
        elif t == "Frame":
            w = QGroupBox(strip_amp(cap))
        elif t == "CheckBox":
            w = QCheckBox(strip_amp(cap))
            w.setCheckState({1: Qt.CheckState.Checked, 2: Qt.CheckState.PartiallyChecked}.get(geti(c, "Value"), Qt.CheckState.Unchecked))
        elif t == "OptionButton":
            w = QRadioButton(strip_amp(cap))
            w.setChecked(bool(geti(c, "Value")))
        elif t == "ComboBox":
            w = QComboBox()
            w.setEditable(True)
            w.setEditText(str(val("Text", c.name)) if geti(c, "Style") != 2 else "")
        elif t == "ListBox":
            w = QListWidget()
            w.addItems(val("List", None) or [c.name])     # VB6 shows the items, or the name
        elif t == "PictureBox":
            w = QLabel()
            w.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            pm = self.pixmap("Picture")
            if pm is not None:
                w.setPixmap(pm)
            if geti(c, "BorderStyle", 1):
                w.setFrameStyle(QFrame.Shape.Panel | QFrame.Shadow.Sunken)
                w.setLineWidth(2)
        elif t in ("HScrollBar", "VScrollBar"):
            w = QScrollBar(Qt.Orientation.Horizontal if t == "HScrollBar" else Qt.Orientation.Vertical)
        elif t == "DriveListBox":
            w = QComboBox()
            w.addItem("c: [System]")
        elif t == "DirListBox":
            w = QListWidget()
            w.addItems(["c:\\", "  Program Files", "  Windows"])
        elif t == "FileListBox":
            w = QListWidget()
        elif t == "Data":
            w = DataLook(strip_amp(cap))
        elif ocxlib.control_of(c.cls) and t not in FIXED:
            w = ocx_visual(c, t, val, lambda n: list_images(self.designer.module, self.designer.root, n))
        elif uctl_of(c.cls):                    # a user control: as its designer has it
            w = QLabel()
            w.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            try:
                w.setPixmap(uctl_look(self.designer.main, uctl_of(c.cls)[4]))
            except Exception:                   # noqa: BLE001 - a user control that can't be drawn: its name
                w.setText(c.name)
        elif "." in c.cls and not c.cls.lower().startswith("vb.") and t not in FIXED:
            w = QLabel(f"{c.name}\n({c.cls})")     # a control the IDE does not know: a box with its name
            w.setAlignment(Qt.AlignmentFlag.AlignCenter)
            w.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Plain)
        if w is not None:
            w.setParent(self)
            light_look(w)                       # light at design time (the program follows the system)
            w.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
            w.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            w.setFont(ctl_font(c, self.designer.form_font()))
            pal = w.palette()
            white = t in ("TextBox", "ComboBox", "ListBox", "DriveListBox", "DirListBox", "FileListBox", "TreeView",
                          "ListView", "RichTextBox", "ImageCombo", "DTPicker", "MonthView", "Animation", "Data")
            back = vb_color(c.get("BackColor"), 0x80000005 if white else 0x8000000F)
            fore = vb_color(c.get("ForeColor"), 0x80000008 if white else 0x80000012)
            if white:
                pal.setColor(QPalette.ColorRole.Base, back)
                pal.setColor(QPalette.ColorRole.Text, fore)
            elif t == "CommandButton" and geti(c, "Style") == 1:
                pal.setColor(QPalette.ColorRole.Button, back)   # a Graphical button shows its BackColor
            elif t != "CommandButton":
                w.setBackgroundRole(QPalette.ColorRole.Window)   # not the Dark of the designer's scroll area
                w.setForegroundRole(QPalette.ColorRole.WindowText)
                pal.setColor(QPalette.ColorRole.Window, back)
                pal.setColor(QPalette.ColorRole.WindowText, fore)
                w.setAutoFillBackground(not (t == "Label" and geti(c, "BackStyle", 1) == 0))
            w.setPalette(pal)
            w.show()
            w.lower()
        self.visual = w
        self.place()
        self.update()

    def val(self, name, default=None):
        """a property, with the data of the .frx (List, multi-line Text, pictures)"""
        v = self.designer.module.frx_value(self.ctl, name)
        return default if v is None else v

    def pixmap(self, name):
        from PyQt6.QtGui import QPixmap
        data = self.designer.module.frx_value(self.ctl, name)
        if not data or not isinstance(data, (bytes, bytearray)):
            return None
        pm = QPixmap()
        return pm if pm.loadFromData(bytes(data)) else None

    def paintEvent(self, e):
        t, c = self.type, self.ctl
        p = QPainter(self)
        if t == "Timer":
            p.fillRect(self.rect(), QColor("#F0F0F0"))
            p.setPen(QColor("#808080"))
            p.drawRect(self.rect().adjusted(0, 0, -1, -1))
            p.setPen(QPen(QColor("#202020"), 2))
            r = self.rect().adjusted(6, 6, -6, -6)
            p.drawEllipse(r)
            ctr = r.center()
            p.drawLine(ctr, QPoint(ctr.x(), r.top() + 4))
            p.drawLine(ctr, QPoint(r.right() - 5, ctr.y()))
        elif t == "Shape":
            bw = max(1, geti(c, "BorderWidth", 1))
            pen = QPen(vb_color(c.get("BorderColor"), 0x80000008), bw)
            pen.setStyle({2: Qt.PenStyle.DashLine, 3: Qt.PenStyle.DotLine, 4: Qt.PenStyle.DashDotLine,
                          5: Qt.PenStyle.DashDotDotLine}.get(geti(c, "BorderStyle", 1), Qt.PenStyle.SolidLine))
            if geti(c, "BorderStyle", 1) == 0:
                pen = QPen(Qt.PenStyle.NoPen)
            p.setPen(pen)
            fs = geti(c, "FillStyle", 1)
            p.setBrush(QBrush(vb_color(c.get("FillColor"), 0), Qt.BrushStyle.SolidPattern) if fs == 0 else Qt.BrushStyle.NoBrush)
            r = self.rect().adjusted(bw // 2, bw // 2, -(bw + 1) // 2, -(bw + 1) // 2)
            shp = geti(c, "Shape")
            if shp in (1, 3, 5):
                s = min(r.width(), r.height())
                r = QRect(r.center().x() - s // 2, r.center().y() - s // 2, s, s)
            if geti(c, "BackStyle") == 1:
                p.fillRect(r, vb_color(c.get("BackColor"), 0x80000005))
            if shp in (2, 3):
                p.drawEllipse(r)
            elif shp in (4, 5):
                p.drawRoundedRect(r, min(r.width(), r.height()) / 6, min(r.width(), r.height()) / 6)
            else:
                p.drawRect(r)
        elif t == "Line":
            x1, y1, x2, y2 = (geti(c, k) // TPP for k in ("X1", "Y1", "X2", "Y2"))
            ox, oy = self.x(), self.y()
            p.setPen(QPen(vb_color(c.get("BorderColor"), 0x80000008), max(1, geti(c, "BorderWidth", 1))))
            p.drawLine(x1 - ox, y1 - oy, x2 - ox, y2 - oy)
        elif t == "Image":
            pm = self.pixmap("Picture")
            if pm is not None:
                if geti(c, "Stretch"):
                    p.drawPixmap(self.rect(), pm)
                else:
                    p.drawPixmap(0, 0, pm)
            pen = QPen(QColor("#808080"), 1, Qt.PenStyle.DashLine)
            p.setPen(pen)
            p.drawRect(self.rect().adjusted(0, 0, -1, -1))
        elif t in FIXED:                       # CommonDialog, ImageList: the Toolbox picture
            p.fillRect(self.rect(), QColor("#F0F0F0"))
            p.setPen(QColor("#808080"))
            p.drawRect(self.rect().adjusted(0, 0, -1, -1))
            p.drawPixmap((self.width() - 24) // 2, (self.height() - 24) // 2, tool_icon(t).pixmap(24, 24))
        elif self.visual is None:
            p.fillRect(self.rect(), QColor("#E0E0E0"))
            p.setPen(QColor("#404040"))
            p.drawRect(self.rect().adjusted(0, 0, -1, -1))
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, f"{c.name}\n{c.cls}")

    # ------------------------------------------------------------ mouse
    def mousePressEvent(self, e):
        self.designer.surface.item_pressed(self, e)

    def mouseMoveEvent(self, e):
        self.designer.surface.item_moved(self, e)

    def mouseReleaseEvent(self, e):
        self.designer.surface.item_released(self, e)

    def mouseDoubleClickEvent(self, e):
        self.designer.open_event(self.ctl)


# ---------------------------------------------------------------- the client area of the form
class Handle(QWidget):
    SIZE = 6

    def __init__(self, surface, item, pos, primary):
        super().__init__(surface)
        self.surface, self.item, self.pos_name, self.primary = surface, item, pos, primary
        self.resize(self.SIZE, self.SIZE)
        cur = {"nw": Qt.CursorShape.SizeFDiagCursor, "se": Qt.CursorShape.SizeFDiagCursor,
               "ne": Qt.CursorShape.SizeBDiagCursor, "sw": Qt.CursorShape.SizeBDiagCursor,
               "n": Qt.CursorShape.SizeVerCursor, "s": Qt.CursorShape.SizeVerCursor,
               "e": Qt.CursorShape.SizeHorCursor, "w": Qt.CursorShape.SizeHorCursor}[pos]
        self.setCursor(cur)

    def solid(self):
        """the handles of the last control chosen are solid, the others hollow; with Lock Controls all are
        hollow (DEV-8/fd1 stage 20, picture 29)"""
        return self.primary and not getattr(self.surface.d, "locked", False)

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#000080") if self.solid() else QColor("#FFFFFF"))
        p.setPen(QColor("#000080"))
        p.drawRect(self.rect().adjusted(0, 0, -1, -1))

    def mousePressEvent(self, e):
        self.surface.start_resize(self, e)

    def mouseMoveEvent(self, e):
        self.surface.do_resize(self, e)

    def mouseReleaseEvent(self, e):
        self.surface.end_resize(self, e)


class Surface(QWidget):
    """the client area: grid, controls, rubber band, placing new controls"""

    def __init__(self, designer):
        super().__init__()
        self.d = designer
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setAutoFillBackground(True)
        self.setBackgroundRole(QPalette.ColorRole.Window)
        self.setForegroundRole(QPalette.ColorRole.WindowText)
        self.band = QRubberBand(QRubberBand.Shape.Rectangle, self)
        self.band_origin = None
        self.drag = None                        # moving: (start pos, {item: start rect twips})
        self.resizing = None
        self.handles = []

    # ------------------------------------------------------------ painting
    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), self.palette().color(QPalette.ColorRole.Window))
        pm = self.d.form_pixmap("Picture")
        if pm is not None:
            p.drawPixmap(0, 0, pm)
        if self.d.show_grid:
            p.setPen(QColor(0, 0, 0, 160))
            sx, sy = max(2, F.px(self.d.grid_w)), max(2, F.px(self.d.grid_h))
            for y in range(0, self.height(), sy):
                for x in range(0, self.width(), sx):
                    p.drawPoint(x, y)

    # ------------------------------------------------------------ helpers
    def to_surface(self, w, pt):
        return w.mapTo(self, pt)

    def container_at(self, pos):
        """the Frame / PictureBox under pos (the deepest), or None for the form"""
        best = None
        for it in self.d.items.values():
            if it.type in CONTAINERS:
                tl = it.mapTo(self, QPoint(0, 0))
                if QRect(tl, it.size()).contains(pos):
                    if best is None or best.isAncestorOf(it):
                        best = it
        return best

    def snap(self, v, vertical=False):
        """to the nearest grid line (a half up, as VB6) when Align Controls to Grid is on"""
        if not self.d.align_grid:
            return v
        return F.snap(v, self.d.grid_h if vertical else self.d.grid_w)

    # ------------------------------------------------------------ the mouse on the form
    def mousePressEvent(self, e):
        self.setFocus()
        if e.button() == Qt.MouseButton.RightButton and not self.d.current_tool():
            self.d.select([])                   # the right button on the form: the form, and its menu (me8-01)
            self.d.context_menu(e.globalPosition().toPoint(), None)
            return
        if e.button() != Qt.MouseButton.LeftButton:
            return
        self.band_origin = e.position().toPoint()
        self.band.setGeometry(QRect(self.band_origin, QSize()))
        self.band.show()

    def mouseMoveEvent(self, e):
        if self.band_origin is not None:
            self.band.setGeometry(QRect(self.band_origin, e.position().toPoint()).normalized())

    def mouseReleaseEvent(self, e):
        self.finish_band(e.position().toPoint())

    def finish_band(self, pos):
        if self.band_origin is None:
            return
        r = QRect(self.band_origin, pos).normalized()
        self.band.hide()
        self.band_origin = None
        tool = self.d.current_tool()
        if tool:
            self.d.place_new(tool, r, self.container_at(r.topLeft()))
            return
        if r.width() < 3 and r.height() < 3:
            self.d.select([])
            return
        # the rubber band selects the controls it touches (in the same container)
        hit = []
        for it in self.d.items.values():
            tl = it.mapTo(self, QPoint(0, 0))
            if QRect(tl, it.size()).intersects(r) and it.parentWidget() is self:
                hit.append(it)
        self.d.select(hit)

    def mouseDoubleClickEvent(self, e):
        self.d.open_event(None)                 # the form: Form_Load

    # ------------------------------------------------------------ the mouse on a control
    def item_pressed(self, item, e):
        self.setFocus()
        if e.button() == Qt.MouseButton.RightButton and not self.d.current_tool():
            if item not in self.d.selection:    # the right button on a control: it (the selection stays when
                self.d.select([item])           # it is one of them), and its menu (DEV-8/me8-02, 07)
            self.d.context_menu(e.globalPosition().toPoint(), item)
            return
        if e.button() != Qt.MouseButton.LeftButton:
            return
        tool = self.d.current_tool()
        if tool and item.type in CONTAINERS:   # drawing a new control inside a Frame
            self.band_origin = item.mapTo(self, e.position().toPoint())
            self.band.setGeometry(QRect(self.band_origin, QSize()))
            self.band.show()
            return
        mods = e.modifiers()
        if mods & (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.ShiftModifier):
            sel = list(self.d.selection)
            if item in sel:
                sel.remove(item)
            else:
                sel.append(item)
            self.d.select(sel, primary=item if item in sel else None)
            return
        if item not in self.d.selection:
            self.d.select([item])
        else:
            self.d.select(self.d.selection, primary=item)
        if self.d.locked:                       # Format > Lock Controls: the mouse does not move them
            self.drag = None
            return
        self.drag = (item.mapTo(self, e.position().toPoint()), {it: it.rect_twips() for it in self.d.selection}, False)

    def item_moved(self, item, e):
        if self.band_origin is not None:
            self.band.setGeometry(QRect(self.band_origin, item.mapTo(self, e.position().toPoint())).normalized())
            return
        if not self.drag:
            return
        start, rects, moved = self.drag
        pos = item.mapTo(self, e.position().toPoint())
        dx, dy = (pos.x() - start.x()) * TPP, (pos.y() - start.y()) * TPP
        if not moved and abs(dx) < 3 * TPP and abs(dy) < 3 * TPP:
            return
        if not moved:
            self.d.checkpoint()
        self.drag = (start, rects, True)
        prim = self.d.primary
        r0 = rects.get(prim, next(iter(rects.values())))
        nx, ny = self.snap(r0.x() + dx), self.snap(r0.y() + dy, True)
        dx, dy = nx - r0.x(), ny - r0.y()
        for it, r in rects.items():
            self.d.move_item(it, r.x() + dx, r.y() + dy)
        self.place_handles()

    def item_released(self, item, e):
        if self.band_origin is not None:
            self.finish_band(item.mapTo(self, e.position().toPoint()))
            return
        if self.drag and self.drag[2]:
            self.d.changed()
        self.drag = None

    # ------------------------------------------------------------ handles
    def handle_spots(self, it):
        """(handle name, x, y, enabled) of the handles of a selected control"""
        tl = it.mapTo(self, QPoint(0, 0))
        r = QRect(tl, it.size())
        s = Handle.SIZE
        if it.type == "Line":                   # a handle at the end of the line (X2, Y2)
            o = it.parentWidget().mapTo(self, QPoint(0, 0)) if it.parentWidget() is not self else QPoint(0, 0)
            return [("se", o.x() + geti(it.ctl, "X2") // TPP - s // 2, o.y() + geti(it.ctl, "Y2") // TPP - s // 2, True)]
        if it.type in FIXED:                    # a Timer has no size
            return [("se", r.right() - 2, r.bottom() - 2, False)]
        return [("nw", r.left() - s, r.top() - s, True), ("n", r.center().x() - s // 2, r.top() - s, True),
                ("ne", r.right() + 1, r.top() - s, True), ("e", r.right() + 1, r.center().y() - s // 2, True),
                ("se", r.right() + 1, r.bottom() + 1, True), ("s", r.center().x() - s // 2, r.bottom() + 1, True),
                ("sw", r.left() - s, r.bottom() + 1, True), ("w", r.left() - s, r.center().y() - s // 2, True)]

    def place_handles(self):
        # the same controls selected (a resize or a move going on): the handles move; new handles
        # would take the mouse away from the one being dragged
        want = [(it, name) for it in self.d.selection for name, _, _, _ in self.handle_spots(it)]
        if [(h.item, h.pos_name) for h in self.handles] == want:
            spots = {(it, name): (x, y) for it in self.d.selection for name, x, y, _ in self.handle_spots(it)}
            for h in self.handles:
                h.primary = h.item is self.d.primary
                h.move(*spots[(h.item, h.pos_name)])
                h.raise_()
                h.update()
            return
        for h in self.handles:
            h.deleteLater()
        self.handles = []
        for it in self.d.selection:
            for name, x, y, enabled in self.handle_spots(it):
                h = Handle(self, it, name, it is self.d.primary)
                h.move(x, y)
                if not enabled:
                    h.setCursor(Qt.CursorShape.ArrowCursor)
                    h.setEnabled(False)
                h.show()
                h.raise_()
                self.handles.append(h)

    def start_resize(self, h, e):
        if self.d.locked:                       # Format > Lock Controls: the handles do not size them
            self.resizing = None
            return
        self.d.checkpoint()
        c = h.item.ctl
        self.resizing = (h.mapTo(self, e.position().toPoint()), h.item.rect_twips(), geti(c, "X2"), geti(c, "Y2"))

    def do_resize(self, h, e):
        if not self.resizing:
            return
        start, r, x2, y2 = self.resizing
        pos = h.mapTo(self, e.position().toPoint())
        dx, dy = (pos.x() - start.x()) * TPP, (pos.y() - start.y()) * TPP
        it = h.item
        if it.type == "Line":                   # the handle moves the end of the line (X2, Y2)
            c = it.ctl
            set_prop(c, "X2", self.snap(x2 + dx))
            set_prop(c, "Y2", self.snap(y2 + dy, True))
            it.place()
            it.update()
            self.place_handles()
            return
        x, y, w, hh = r.x(), r.y(), r.width(), r.height()
        n = h.pos_name
        if "e" in n:
            w = max(TPP, self.snap(x + w + dx) - x)
        if "s" in n:
            hh = max(TPP, self.snap(y + hh + dy, True) - y)
        if "w" in n:
            nx = self.snap(x + dx)
            w = max(TPP, w + x - nx)
            x = nx
        if "n" in n:
            ny = self.snap(y + dy, True)
            hh = max(TPP, hh + y - ny)
            y = ny
        if it.type == "ComboBox":
            hh = r.height()                     # VB6: the height of a ComboBox is fixed
        self.d.set_rect(it, x, y, w, hh)
        self.place_handles()

    def end_resize(self, h, e):
        if self.resizing:
            self.d.changed()
        self.resizing = None

    # ------------------------------------------------------------ keys
    def keyPressEvent(self, e):
        k, mods = e.key(), e.modifiers()
        sel = self.d.selection
        if k == Qt.Key.Key_Delete:
            self.d.delete_selection()
        elif e.matches(QKeySequence.StandardKey.Copy):
            self.d.copy_selection()
        elif e.matches(QKeySequence.StandardKey.Cut):
            self.d.copy_selection()
            self.d.delete_selection()
        elif e.matches(QKeySequence.StandardKey.Paste):
            self.d.paste()
        elif e.matches(QKeySequence.StandardKey.Undo):
            self.d.undo()
        elif e.matches(QKeySequence.StandardKey.SelectAll):
            self.d.select([it for it in self.d.items.values() if it.parentWidget() is self])
        elif k in (Qt.Key.Key_Left, Qt.Key.Key_Right, Qt.Key.Key_Up, Qt.Key.Key_Down) and sel:
            self.d.arrow_key(k, mods)
        elif k == Qt.Key.Key_Escape:
            self.d.select([])
        else:
            super().keyPressEvent(e)


class FormFrame(QWidget):
    """the window around the client area: caption bar, borders, menu bar (Windows 7 look),
    and the handles that size the form"""

    def __init__(self, designer):
        super().__init__()
        self.d = designer
        self.surface = Surface(designer)
        self.surface.setParent(self)
        self.sizing = None

    def frame(self):
        if self.d.root.type_name in ("UserControl", "PropertyPage"):   # its client area only
            return 0, 0, 0
        bs = geti(self.d.root, "BorderStyle", 2)
        if bs == 0:
            return 0, 0, 0
        border, top = V.FRAMES.get(bs, V.FRAMES[2])[:2]
        menu = 20 if self.d.menus() else 0
        return border, top - border, menu

    def layout_frame(self):
        b, cap, menu = self.frame()
        cw = geti(self.d.root, "ClientWidth", 4800) // TPP
        ch = geti(self.d.root, "ClientHeight", 3600) // TPP
        self.surface.setGeometry(b, b + cap + menu, cw, ch)
        self.setFixedSize(cw + 2 * b + 10, ch + 2 * b + cap + menu + 10)
        self.update()

    def paintEvent(self, e):
        p = QPainter(self)
        b, cap, menu = self.frame()
        r = QRect(0, 0, self.surface.width() + 2 * b, self.surface.height() + 2 * b + cap + menu)
        if b:
            p.fillRect(r, QColor("#BFCDDB") if not self.d.active else QColor("#99B4D1"))
            p.setPen(QColor("#6D7A88"))
            p.drawRect(r.adjusted(0, 0, -1, -1))
            p.setPen(QColor("#000000"))
            f = QFont(self.font())
            f.setPixelSize(12)
            p.setFont(f)
            title = str(self.d.root.get("Caption", self.d.root.name))
            p.drawText(QRect(b + 22, b, r.width() - 2 * b - 100, cap), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, title)
            # the icon and the buttons of the caption bar
            icon = self.d.form_pixmap("Icon")
            if icon is not None:
                p.drawPixmap(QRect(b + 2, b + 3, 16, 16), icon)
            else:
                p.fillRect(QRect(b + 2, b + 3, 16, 16), QColor("#5A7FB0"))
            x = r.width() - b - 2
            btns = []
            if geti(self.d.root, "ControlBox", -1):
                btns.append("x")
                if geti(self.d.root, "MaxButton", -1) or geti(self.d.root, "MinButton", -1):
                    btns += ["□", "_"]
            for t in btns:
                w = 40 if t == "x" else 26
                x -= w
                br = QRect(x, b + 1, w, cap - 4)
                p.fillRect(br, QColor("#C75050") if t == "x" else QColor("#DDE4EE"))
                p.setPen(QColor("#FFFFFF") if t == "x" else QColor("#000000"))
                p.drawText(br, Qt.AlignmentFlag.AlignCenter, t)
            if menu:
                mr = QRect(b, b + cap, r.width() - 2 * b, menu)
                p.fillRect(mr, QColor("#F0F0F0"))
                p.setPen(QColor("#000000"))
                f.setPixelSize(12)
                p.setFont(f)
                for m, rect in self.menu_rects():
                    p.drawText(rect, Qt.AlignmentFlag.AlignCenter, strip_amp(str(m.get("Caption", ""))))
        # the sizing handles of the form (right, bottom, corner)
        sr = self.surface.geometry()
        p.setPen(QColor("#000080"))
        for pt in (QPoint(sr.right() + b + 1, sr.center().y()), QPoint(sr.center().x(), sr.bottom() + b + 1),
                   QPoint(sr.right() + b + 1, sr.bottom() + b + 1)):
            p.fillRect(QRect(pt.x(), pt.y() - 3, 6, 6), QColor("#FFFFFF"))
            p.drawRect(QRect(pt.x(), pt.y() - 3, 5, 5))

    def menu_rects(self):
        """the visible top-level menus and their places on the menu bar"""
        b, cap, menu = self.frame()
        if not menu:
            return []
        f = QFont(self.font())
        f.setPixelSize(12)
        from PyQt6.QtGui import QFontMetrics
        fm = QFontMetrics(f)
        out, x = [], b + 6
        for m in self.d.menus():
            if not V.value_of(m, V.prop_def("Menu", "Visible")):
                continue
            w = fm.horizontalAdvance(strip_amp(str(m.get("Caption", "")))) + 14
            out.append((m, QRect(x, b + cap, w, menu)))
            x += w
        return out

    def drop_menu(self, m, rect):
        """VB6: a click on a menu of the designer drops it; choosing an item opens its Click code"""
        from PyQt6.QtWidgets import QMenu

        def fill(qm, parent):
            for c in parent.children:
                if c.type_name != "Menu":
                    continue
                cap = str(c.get("Caption", ""))
                if cap == "-":
                    qm.addSeparator()
                    continue
                kids = [k for k in c.children if k.type_name == "Menu"]
                text = cap
                sc = c.get("Shortcut")
                if sc:
                    text += "\t" + V.TEXT_OF_SHORTCUT.get(str(sc).upper(), str(sc))
                if kids:
                    sub = qm.addMenu(text)
                    fill(sub, c)
                else:
                    a = qm.addAction(text)
                    a.setCheckable(bool(V.value_of(c, V.prop_def("Menu", "Checked"))))
                    a.setChecked(a.isCheckable())
                    a.setEnabled(bool(V.value_of(c, V.prop_def("Menu", "Enabled"))))
                    a.triggered.connect(lambda _=False, c=c: self.d.open_event(c))
        qm = QMenu(self)
        fill(qm, m)
        if qm.isEmpty():
            self.d.open_event(m)
            return
        qm.exec(self.mapToGlobal(rect.bottomLeft()))

    def _spot(self, pos):
        sr = self.surface.geometry()
        b = self.frame()[0]
        right = abs(pos.x() - (sr.right() + b + 3)) <= 5
        bottom = abs(pos.y() - (sr.bottom() + b + 1)) <= 5
        return ("e" if right else "") + ("s" if bottom else "") if (right or bottom) else None

    def mouseMoveEvent(self, e):
        pos = e.position().toPoint()
        if self.sizing:
            spot, start, cw, ch = self.sizing
            dx, dy = (pos.x() - start.x()) * TPP, (pos.y() - start.y()) * TPP
            ncw = max(TPP * 60, cw + dx) if "e" in spot else cw
            nch = max(TPP * 20, ch + dy) if "s" in spot else ch
            self.d.set_client_size(int(ncw), int(nch))
            return
        spot = self._spot(pos)
        self.setCursor({"e": Qt.CursorShape.SizeHorCursor, "s": Qt.CursorShape.SizeVerCursor,
                        "es": Qt.CursorShape.SizeFDiagCursor}.get(spot or "", Qt.CursorShape.ArrowCursor))

    def mousePressEvent(self, e):
        pos = e.position().toPoint()
        for m, rect in self.menu_rects():
            if rect.contains(pos):
                self.drop_menu(m, rect)
                return
        spot = self._spot(pos)
        if spot:
            self.d.checkpoint()
            self.sizing = (spot, pos, geti(self.d.root, "ClientWidth", 4800), geti(self.d.root, "ClientHeight", 3600))
        else:
            self.d.select([])

    def mouseReleaseEvent(self, e):
        if self.sizing:
            self.sizing = None
            self.d.changed()


# ---------------------------------------------------------------- the designer window
class FormDesigner(QScrollArea):
    selection_changed = pyqtSignal(object)
    modified = pyqtSignal(object)
    geometry_changed = pyqtSignal()             # a control moved / resized (the Standard toolbar shows where)

    def __init__(self, main, module):
        super().__init__()
        self.main = main
        self.module = module
        self.items = {}                         # Control -> DesignItem
        self.selection = []
        self.primary = None
        self.show_grid = True
        self.align_grid = True
        self.grid_w = self.grid_h = GRID
        self.grid_settings(getattr(main, "S", None))
        self.active = True
        self.undo_stack = []
        self.setWidgetResizable(False)
        self.setBackgroundRole(QPalette.ColorRole.Dark)
        self.frame_w = FormFrame(self)
        self.surface = self.frame_w.surface
        self.setWidget(self.frame_w)
        self.resolve_db_frx()
        self.build()

    @property
    def root(self):
        return self.module.sf.root

    def resolve_db_frx(self):
        """the OleObjectBlob of a DBGrid and the Bindings of a DBGrid / DBList / DBCombo as bytes: the
        Properties window reads and changes them (vbprops.BLOB_PROPS, lbvc/project/oleblob.py)"""
        from lbvc.project.frx import FrxData
        if self.module.sf is None or self.root is None:
            return
        for c in self.root.walk():
            if c.type_name not in ("DBGrid", "DBList", "DBCombo", "DataGrid"):
                continue
            for name in ("OleObjectBlob", "Bindings"):
                p = c.prop(name)
                if p is not None and not isinstance(p.value, FrxData):
                    d = self.module.frx_data(p)
                    if d is not None and isinstance(d.data, (bytes, bytearray)):
                        p.value = FrxData("raw", bytes(d.data))

    def menus(self):
        return [c for c in self.root.children if c.type_name == "Menu"]

    def form_pixmap(self, name):
        """Picture or Icon of the form from the .frx"""
        from PyQt6.QtGui import QPixmap
        data = self.module.frx_value(self.root, name)
        if not data or not isinstance(data, (bytes, bytearray)):
            return None
        pm = QPixmap()
        return pm if pm.loadFromData(bytes(data)) else None

    def form_font(self):
        return ctl_font(self.root)

    def current_tool(self):
        return self.main.toolbox.current() if getattr(self.main, "toolbox", None) else None

    # ------------------------------------------------------------ (re)building
    def build(self):
        for it in list(self.items.values()):
            it.deleteLater()
        self.items = {}
        self.selection, self.primary = [], None
        root = self.root
        pal = light_palette()                   # the controls inherit it: light whatever the theme
        pal.setColor(QPalette.ColorRole.Window, vb_color(root.get("BackColor"), 0x8000000C if root.type_name ==
                                                         "MDIForm" else 0x8000000F))   # MDI: AppWorkspace
        self.surface.setPalette(pal)
        self.surface.setFont(self.form_font())
        self._build_children(root, self.surface)
        self.frame_w.layout_frame()
        self.surface.place_handles()

    def _build_children(self, ctl, parent_widget):
        # the first control of the .frm is on top: make them last to first
        for c in reversed([c for c in ctl.children if c.type_name != "Menu"]):
            it = DesignItem(self, c, parent_widget)
            it.show()
            self.items[c] = it
            if c.type_name in CONTAINERS:
                self._build_children(c, it)

    # ------------------------------------------------------------ selection
    def select(self, items, primary=None):
        self.selection = list(items)
        self.primary = primary if primary in self.selection else (self.selection[0] if self.selection else None)
        self.surface.place_handles()
        self.selection_changed.emit(self.primary.ctl if self.primary else self.root)

    def select_control(self, ctl):
        it = self.items.get(ctl)
        self.select([it] if it else [])

    # ------------------------------------------------------------ changing the form
    def checkpoint(self):
        """for Undo: the designer section as it is now"""
        self.undo_stack.append(copy.deepcopy(self.module.sf.root))
        del self.undo_stack[:-50]

    def undo(self):
        if not self.undo_stack:
            return
        self.module.sf.root = self.undo_stack.pop()
        self.build()
        self.changed()

    def changed(self):
        self.module.dirty = True
        self.modified.emit(self.module)

    def move_item(self, it, x, y):
        c = it.ctl
        if it.type == "Line":
            r = it.rect_twips()
            dx, dy = x - r.x(), y - r.y()
            for k, d in (("X1", dx), ("X2", dx), ("Y1", dy), ("Y2", dy)):
                set_prop(c, k, geti(c, k) + d)
        else:
            set_prop(c, "Left", int(x))
            set_prop(c, "Top", int(y))
        it.place()
        it.update()
        self.geometry_changed.emit()

    def set_rect(self, it, x, y, w, h):
        c = it.ctl
        set_prop(c, "Left", int(x))
        set_prop(c, "Top", int(y))
        set_prop(c, "Width", int(w))
        h = integral_height(c, it.type, h)
        set_prop(c, "Height", int(h))
        if it.type == "PictureBox":            # the scale of a PictureBox follows its size
            self.rescale(c)
        lc = ocxlib.control_of(c.cls)
        if lc:                                 # an .ocx keeps its size in HIMETRIC too
            px = getattr(lc[1], "extent", "_")
            set_prop(c, px + "ExtentX", ocxlib.extent(int(w)))
            set_prop(c, px + "ExtentY", ocxlib.extent(int(h)))
        it.place()
        self.geometry_changed.emit()

    def set_client_size(self, cw, ch):
        root = self.root
        set_prop(root, "ClientWidth", cw)
        set_prop(root, "ClientHeight", ch)
        self.rescale(root)
        if self.realign():
            self.build()
        self.frame_w.layout_frame()
        self.geometry_changed.emit()

    def realign(self):
        """the controls with an Align of their own (PictureBox, Data): on an edge of the form, the
        whole width or height of it, as VB6 places them; True if one moved"""
        cw, ch = geti(self.root, "ClientWidth", 4800), geti(self.root, "ClientHeight", 3600)
        moved = False
        for c in self.root.children:
            al = geti(c, "Align") if c.type_name in ("PictureBox", "Data", "CoolBar") else 0
            if al not in (1, 2, 3, 4):
                continue
            w, h = geti(c, "Width", 1215), geti(c, "Height", 345)
            x, y = 0, 0
            if al in (1, 2):
                w = cw
                y = 0 if al == 1 else max(0, ch - h)
            else:
                h = ch
                x = 0 if al == 3 else max(0, cw - w)
            for k, v in (("Height", h), ("Left", x), ("Top", y), ("Width", w)):
                if geti(c, k, None) != v:
                    V.put(c, k, v)
                    moved = True
            if c.type_name == "CoolBar" and moved:   # its bands follow the new width
                V.cool_place(c)
        return moved

    # ------------------------------------------------------------ the grid, Lock Controls, the keys
    def grid_settings(self, S):
        """Tools > Options > General: Form Grid Settings (Show Grid, Width, Height, Align Controls to Grid)"""
        if S is None:
            return
        try:
            self.show_grid = bool(S["show_grid"])
            self.align_grid = bool(S["align_grid"])
            self.grid_w = max(TPP * 2, int(S["grid_width"]))
            self.grid_h = max(TPP * 2, int(S["grid_height"]))
        except (KeyError, TypeError, ValueError):
            pass
        if getattr(self, "surface", None) is not None:
            self.surface.update()

    @property
    def locked(self):
        """Format > Lock Controls: LockControls = -1 'True of the form (DEV-8/fd1 s4)"""
        return geti(self.root, "LockControls", 0) != 0

    def toggle_lock(self):
        self.checkpoint()
        if self.locked:
            V.remove(self.root, "LockControls")
        else:
            V.put(self.root, "LockControls", True)
        self.surface.place_handles()
        self.changed()

    def arrow_key(self, k, mods):
        """an arrow alone selects the next control that way; Ctrl+arrow moves, Shift+arrow sizes the selected
        controls by the grid (a pixel without Align Controls to Grid), locked or not (DEV-8/fd1 stages 10-15)"""
        ctrl = bool(mods & Qt.KeyboardModifier.ControlModifier)
        shift = bool(mods & Qt.KeyboardModifier.ShiftModifier)
        dx = {Qt.Key.Key_Left: -1, Qt.Key.Key_Right: 1}.get(k, 0)
        dy = {Qt.Key.Key_Up: -1, Qt.Key.Key_Down: 1}.get(k, 0)
        sel = [it for it in self.selection if it.type != "Line"]
        if not ctrl and not shift:
            cur = self.primary or (self.selection[0] if self.selection else None)
            if cur is None:
                return
            sibs = [it for it in self.items.values() if it.parentWidget() is cur.parentWidget()]
            rects = [tuple(it.rect_twips().getRect()) for it in sibs]
            way = {Qt.Key.Key_Left: "left", Qt.Key.Key_Right: "right", Qt.Key.Key_Up: "up", Qt.Key.Key_Down: "down"}[k]
            j = F.neighbor(rects, sibs.index(cur), way)
            if j is not None:
                self.select([sibs[j]])
            return
        if not sel:
            return
        self.checkpoint()
        for it in sel:
            x, y, w, h = it.rect_twips().getRect()
            if shift:
                if it.type in FIXED:
                    continue
                w = F.nudge(w, dx, self.grid_w, self.align_grid, size=True) if dx else w
                h = F.nudge(h, dy, self.grid_h, self.align_grid, size=True) if dy and it.type != "ComboBox" else h
                self.set_rect(it, x, y, w, h)
            else:
                self.move_item(it, F.nudge(x, dx, self.grid_w, self.align_grid) if dx else x,
                               F.nudge(y, dy, self.grid_h, self.align_grid) if dy else y)
        self.surface.place_handles()
        self.changed()

    def format_state(self):
        """what the Format menu has on for the selection (DEV-8/fd1 stages 1, 19): with the form nothing but
        Lock Controls; one control: Align to Grid, Size to Grid, Center in Form, Order; Align, Make Same Size
        and Increase / Decrease / Remove with two; Make Equal with three"""
        n = len([it for it in self.selection if it.type != "Line"])
        return {"one": n >= 1, "two": n >= 2, "three": n >= 3, "lock": True, "locked": self.locked}

    def spacing(self, vertical, how):
        """Format > Horizontal / Vertical Spacing (formops.spacing)"""
        sel = [it for it in self.selection if it.type != "Line"]
        if len(sel) < 2:
            return
        self.checkpoint()
        rects = [tuple(it.rect_twips().getRect()) for it in sel]
        for it, r in zip(sel, F.spacing(rects, vertical, how, self.grid_h if vertical else self.grid_w)):
            self.move_item(it, r[0], r[1])
        self.surface.place_handles()
        self.changed()

    def size_to_grid(self):
        """Format > Size to Grid: Left, Top and the right edge on the grid; the Height stays (VB6)"""
        sel = [it for it in self.selection if it.type != "Line" and it.type not in FIXED]
        if not sel:
            return
        self.checkpoint()
        for it in sel:
            x, y, w, h = F.size_to_grid(tuple(it.rect_twips().getRect()), self.grid_w, self.grid_h)
            self.set_rect(it, x, y, w, h)
        self.surface.place_handles()
        self.changed()

    # ------------------------------------------------------------ Format menu
    def align(self, how):
        """Format > Align (lefts, centers, rights, tops, middles, bottoms, to grid) to the primary control"""
        sel = [it for it in self.selection if it.type != "Line"]
        if not sel or self.primary is None:
            return
        self.checkpoint()
        p = self.primary.rect_twips()
        for it in sel:
            r = it.rect_twips()
            x, y = r.x(), r.y()
            if how == "lefts": x = p.x()
            elif how == "centers": x = p.x() + (p.width() - r.width()) // 2
            elif how == "rights": x = p.x() + p.width() - r.width()
            elif how == "tops": y = p.y()
            elif how == "middles": y = p.y() + (p.height() - r.height()) // 2
            elif how == "bottoms": y = p.y() + p.height() - r.height()
            elif how == "grid": x, y = F.snap(x, self.grid_w), F.snap(y, self.grid_h)
            self.move_item(it, x, y)
        self.surface.place_handles()
        self.changed()

    def same_size(self, how):
        sel = [it for it in self.selection if it.type != "Line" and it.type not in FIXED]
        if not sel or self.primary is None:
            return
        self.checkpoint()
        p = self.primary.rect_twips()
        for it in sel:
            r = it.rect_twips()
            self.set_rect(it, r.x(), r.y(), p.width() if how in ("width", "both") else r.width(),
                          p.height() if how in ("height", "both") and it.type != "ComboBox" else r.height())
        self.surface.place_handles()
        self.changed()

    def center_in_form(self, horizontal):
        sel = [it for it in self.selection if it.type != "Line"]
        if not sel:
            return
        self.checkpoint()
        box = sel[0].rect_twips()
        for it in sel[1:]:
            box = box.united(it.rect_twips())
        pw = geti(self.root, "ScaleWidth", geti(self.root, "ClientWidth", 4800))
        ph = geti(self.root, "ScaleHeight", geti(self.root, "ClientHeight", 3600))
        dx = (pw - box.width()) // 2 - box.x() if horizontal else 0
        dy = 0 if horizontal else (ph - box.height()) // 2 - box.y()
        for it in sel:
            r = it.rect_twips()
            self.move_item(it, r.x() + dx, r.y() + dy)
        self.surface.place_handles()
        self.changed()

    def z_order_state(self):
        """(Bring to Front, Send to Back) of the menu of a control: on when they change something: an other
        control of the same container in front of (behind) a selected one (DEV-8/me8-02, 06, 07, 09)"""
        sel = [it.ctl for it in self.selection]
        front = back = False
        for c in sel:
            p = c.parent
            if p is None:
                continue
            kids = [x for x in p.children if x.type_name != "Menu"]
            i = kids.index(c) if c in kids else -1
            if any(x not in sel for x in kids[:i]):
                front = True
            if any(x not in sel for x in kids[i + 1:]):
                back = True
        return front, back

    def context_menu(self, gpos, item):
        """the right button on the form or on a control (the menu: MainWindow.designer_menu)"""
        if hasattr(self.main, "designer_menu"):
            self.main.designer_menu(self, gpos, item)

    def picture_target(self):
        """where a picture of the clipboard goes: the form (nothing selected) or the one selected PictureBox /
        Image (VB6, DEV-8/me8-16, me9-03, 05); None for the other controls (a CommandButton: grey, me8-02)"""
        if not self.selection:
            ctl = self.root
        elif len(self.selection) == 1 and self.selection[0].type in ("PictureBox", "Image"):
            ctl = self.selection[0].ctl
        else:
            return None
        return ctl if V.prop_def(ctl.type_name, "Picture") is not None else None

    def can_paste(self):
        """Paste of the menus: controls copied (on the form and on every control), or a picture (the form, a
        PictureBox, an Image: their Picture, DEV-8/me8-01, 02, 05, 08, 16, me9-02..05)"""
        md = QApplication.clipboard().mimeData()
        if md is None:
            return False
        if md.hasFormat("application/x-lbstudio-controls"):
            return True
        return md.hasImage() and self.picture_target() is not None

    def paste_picture(self):
        """a picture on the clipboard into the Picture of the form / PictureBox / Image (VB6): a .bmp in the
        .frx; an Image without Stretch takes the size of the picture (me9-05), a PictureBox keeps its size"""
        from PyQt6.QtCore import QBuffer, QIODevice
        img = QApplication.clipboard().image()
        ctl = self.picture_target()
        if img.isNull() or ctl is None:
            return False
        buf = QBuffer()
        buf.open(QIODevice.OpenModeFlag.WriteOnly)
        img.save(buf, "BMP")
        err = self.set_property(ctl, V.prop_def(ctl.type_name, "Picture"), bytes(buf.data()))
        return err is None

    def user_controls_used(self):
        """the modules of the user controls of the project that are on this form"""
        out = []
        for c in self.all_controls():
            uc = uctl_of(c.cls)
            if uc is not None and uc[4] not in out:
                out.append(uc[4])
        return out

    def update_user_controls(self):
        """Update UserControls (the menu of the form): the user controls drawn again from their designers"""
        mods = self.user_controls_used()
        for m in mods:
            forget_uctl_look(m)
        if mods:
            keep = [it.ctl for it in self.selection]
            self.build()
            self.select([self.items[c] for c in keep if c in self.items])

    def z_order(self, front):
        """Format > Order: Bring to Front (Ctrl+J) / Send to Back (Ctrl+K)"""
        if not self.selection:
            return
        self.checkpoint()
        keep = [it.ctl for it in self.selection]
        for c in keep:
            p = c.parent
            rest = [x for x in p.children if x is not c]
            p.children = ([c] + rest) if front else (rest + [c])
            normalize_children(p)           # a Label stays under the windowed controls
        self.build()
        self.select([self.items[c] for c in keep if c in self.items])
        self.changed()

    # ------------------------------------------------------------ new controls
    def all_controls(self):
        return [c for c in self.root.walk() if c is not self.root]

    def form_names(self):
        """the names of the project's forms: a control may not take one (the user's decision, Phase 9 / DEV-9
        nm-01..04: VB6 accepts it and then Form1 in the code is the control, not the form)"""
        names = {self.root.name.lower()}
        proj = getattr(getattr(self, "main", None), "project", None)
        for m in getattr(proj, "modules", None) or []:
            if m.kind == "Form":
                names.add(m.name.lower())
        return names

    def new_name(self, base):
        names = {c.name.lower() for c in self.all_controls()} | self.form_names()
        i = 1
        while f"{base}{i}".lower() in names:
            i += 1
        return f"{base}{i}"

    def next_tabindex(self):
        return max([geti(c, "TabIndex", -1) for c in self.all_controls() if c.get("TabIndex") is not None] + [-1]) + 1

    def _insert(self, parent, ctl):
        """a new control goes on top of its layer: VB6 writes the newest control first, the
        windowed controls before the windowless ones (Label, Image, Line, Shape), menus last"""
        ctl.parent = parent
        parent.children.insert(0, ctl)
        normalize_children(parent)

    def place_new(self, cls, rect_px, container_item):
        uc = uctl_of(cls)
        if uc and uc[4] is self.module:         # a user control on itself: VB6 does not allow it
            self.main.toolbox.reset()
            return None
        name = self.new_name(tool_spec(cls)[1])
        parent = container_item.ctl if container_item is not None else self.root
        origin = container_item.mapTo(self.surface, QPoint(0, 0)) if container_item is not None else QPoint(0, 0)
        w0, h0 = tool_spec(cls)[2], tool_spec(cls)[3]
        if rect_px.width() < 4 and rect_px.height() < 4:
            x, y, w, h = self.surface.snap((rect_px.x() - origin.x()) * TPP), self.surface.snap((rect_px.y() - origin.y()) * TPP, True), w0, h0
        else:
            x = self.surface.snap((rect_px.x() - origin.x()) * TPP)
            y = self.surface.snap((rect_px.y() - origin.y()) * TPP, True)
            w = max(self.grid_w, self.surface.snap(rect_px.width() * TPP)) if self.align_grid else rect_px.width() * TPP
            h = max(self.grid_h, self.surface.snap(rect_px.height() * TPP, True)) if self.align_grid else rect_px.height() * TPP
        self.checkpoint()
        c = self.make_control(cls, name, x, y, w, h)
        self._insert(parent, c)
        self.main.toolbox.reset()
        self.build()
        self.select_control(c)
        self.changed()
        return c

    def make_control(self, cls, name, x, y, w, h):
        lc = ocxlib.control_of(cls)
        if lc:
            return self.make_ocx(lc, name, x, y, w, h)
        uc = uctl_of(cls)
        if uc:
            return self.make_uctl(uc, name, x, y, w, h)
        c = Control("VB." + cls, name)
        c.parent = self.root                          # for the font of the form (set again by _insert)
        props = {}
        if cls in CAPTIONED:
            props["Caption"] = name
        if cls in TEXTED:
            props["Text"] = name
        if cls == "ComboBox":
            h = 315
        if cls == "Line":
            props.update({"X1": x, "X2": x + w, "Y1": y, "Y2": y + h})
        elif cls == "Timer":
            props.update({"Left": x, "Top": y})
        else:
            props.update({"Height": h, "Left": x, "Top": y, "Width": w})
        if cls not in NO_TABINDEX:
            props["TabIndex"] = self.next_tabindex()
        if cls == "PictureBox":
            props.update({"ScaleHeight": h - 60, "ScaleWidth": w - 60})
        if cls in ("ListBox", "FileListBox"):
            props["Height"] = integral_height(c, cls, h)
        for k in sorted(props, key=str.lower):
            set_prop(c, k, props[k])
        if cls == "Data":                             # the properties VB6 always writes, with their comments
            for p in V.CLASSES["Data"]:
                if p.always and p.kind != "pos":
                    V.apply(c, p, p.default)
        return c

    def make_uctl(self, uc, name, x, y, w, h):
        """a user control: Begin Project.Name, the size of its extender and its _ExtentX/_ExtentY, as VB6
        writes them; one of another project of the group: the form names that project at its top"""
        proj, ucname, _, _, m = uc
        c = Control(f"{proj}.{ucname}", name)
        c.parent = self.root
        for k, v in (("Height", h), ("Left", x), ("TabIndex", self.next_tabindex()), ("Top", y), ("Width", w)):
            set_prop(c, k, v)
        set_prop(c, "_ExtentX", ocxlib.extent(w))
        set_prop(c, "_ExtentY", ocxlib.extent(h))
        if m.project is not self.module.project and m.project.path and self.module.path:
            rel = os.path.relpath(m.project.path, os.path.dirname(self.module.path)).replace("/", "\\")
            line = f'Object = "*\\A{rel}"'
            sf = self.module.sf
            if sf is not None and not any(o.strip().lower() == line.lower() for o in sf.objects):
                sf.objects.append(line)
        return c

    def make_ocx(self, lc, name, x, y, w, h):
        """a control of an .ocx: the size of the extender, _ExtentX/_ExtentY/_Version and the defaults
        of the control, as VB6 writes them; the form names the library at its top"""
        lib, oc = lc
        c = Control(f"{lib.key}.{oc.cls}", name)
        c.parent = self.root
        if oc.invisible:                        # an icon: its own size (CommonDialog 480, MSComm 570)
            w, h = oc.w, oc.h
            set_prop(c, "Left", x)
            set_prop(c, "Top", y)
        else:
            if oc.cls in ("Toolbar", "StatusBar", "CoolBar"):    # aligned to the form: its whole width
                w = geti(self.root, "ClientWidth", 4800)
                x = 0
                y = 0 if oc.cls == "Toolbar" else max(0, geti(self.root, "ClientHeight", 3600) - h)
            for k, v in (("Height", h), ("Left", x), ("TabIndex", self.next_tabindex()), ("Top", y), ("Width", w)):
                set_prop(c, k, v)
        if oc.cls == "DBGrid":                  # its properties: in an OleObjectBlob of the .frx (oleblob.py)
            from lbvc.project.oleblob import new_blob
            from lbvc.project.frx import FrxData
            set_prop(c, "OleObjectBlob", FrxData("raw", new_blob(name, w, h, x, y)))
        elif getattr(oc, "extent", "_") != "_":   # the WebBrowser: ExtentX / ExtentY after Width, no _Version
            append_prop(c, oc.extent + "ExtentX", ocxlib.extent(w))
            append_prop(c, oc.extent + "ExtentY", ocxlib.extent(h))
        else:
            set_prop(c, "_ExtentX", ocxlib.extent(w))
            set_prop(c, "_ExtentY", ocxlib.extent(h))
            if getattr(oc, "streamid", None) is not None:
                set_prop(c, "_StreamID", oc.streamid)
            set_prop(c, "_Version", oc.version)
        for k, v in oc.extra:
            v = {"$NAME": name, "$W": w, "$H": h}.get(v, v) if isinstance(v, str) else v
            append_prop(c, k, v)                # Caption = "Adodc1", Text = "DataCombo1", _CBWidth = 4800
        sf = self.module.sf
        if sf is not None and not any(ocxlib.lib_of_object_line(o) is lib for o in sf.objects):
            sf.objects.append(lib.frm_line)
        return c

    def add_default(self, cls):
        """double click on a Toolbox button: a control of the default size in the middle"""
        w, h = tool_spec(cls)[2], tool_spec(cls)[3]
        cw, ch = geti(self.root, "ClientWidth", 4800), geti(self.root, "ClientHeight", 3600)
        x, y = self.surface.snap((cw - w) // 2), self.surface.snap((ch - h) // 2, True)
        self.checkpoint()
        c = self.make_control(cls, self.new_name(tool_spec(cls)[1]), x, y, w, h)
        self._insert(self.root, c)
        self.build()
        self.select_control(c)
        self.changed()

    def delete_selection(self):
        if not self.selection:
            return
        self.checkpoint()
        for it in self.selection:
            c = it.ctl
            p = c.parent
            if p is not None and c in p.children:
                p.children.remove(c)
                p.order = [(k, x) for k, x in p.order if x is not c]
        self.close_tab_order()
        self.build()
        self.select([])
        self.changed()

    def close_tab_order(self):
        """VB6 closes up the tab order when controls go (DEV-8/fd1 t2: Command2 1 after Command1 0 went)"""
        ctls = sorted([c for c in self.root.walk() if c is not self.root and c.prop("TabIndex") is not None],
                      key=lambda c: geti(c, "TabIndex"))
        for i, c in enumerate(ctls):
            if geti(c, "TabIndex", -1) != i:
                V.put(c, "TabIndex", i)

    # ------------------------------------------------------------ clipboard
    def copy_selection(self):
        if not self.selection:
            return
        lines = []
        for it in self.selection:
            lines += list(it.ctl.lines(1))
        md = QMimeData()
        md.setData("application/x-lbstudio-controls", "\r\n".join(lines).encode("utf-8"))
        md.setText("\r\n".join(lines))
        QApplication.clipboard().setMimeData(md)

    def paste(self):
        md = QApplication.clipboard().mimeData()
        if not md.hasFormat("application/x-lbstudio-controls"):
            if md.hasImage():
                self.paste_picture()
            return
        text = bytes(md.data("application/x-lbstudio-controls")).decode("utf-8")
        sf = parse_source("VERSION 5.00\r\nBegin VB.Form Clip \r\n" + text + "\r\nEnd\r\n", None, "Form")
        target = self.primary.ctl if self.primary is not None and self.primary.type in CONTAINERS else self.root
        self.checkpoint()
        new = []
        forms = self.form_names()
        for c in list(sf.root.children):
            existing = [x for x in self.all_controls() if x.name.lower() == c.name.lower()]
            if c.name.lower() in forms:                 # the name of a form: a name of its own, no question
                c.name = self.new_name(TOOL.get(c.type_name, (0, c.type_name))[1])
                c.raw_begin = None
                if c.type_name in CAPTIONED and c.prop("Caption") is not None:
                    set_prop(c, "Caption", c.name)
            elif existing:
                # VB6 (DEV-8/me10): an array of that name already: the next Index, no question; else "create a
                # control array?" (Yes / No / Help, Esc does nothing); No, or another type: a name of its own
                same = all(x.type_name == c.type_name for x in existing)
                is_array = any(geti(x, "Index", -1) >= 0 for x in existing)
                if same and (is_array or self.ask_control_array(c.name)):
                    self.join_array(existing, c)
                else:
                    c.name = self.new_name(TOOL.get(c.type_name, (0, c.type_name))[1])
                    c.raw_begin = None
                    if c.type_name in CAPTIONED and c.prop("Caption") is not None:
                        set_prop(c, "Caption", c.name)
            set_prop(c, "Left", 0) if c.type_name != "Line" else None
            set_prop(c, "Top", 0) if c.type_name != "Line" else None
            if c.get("TabIndex") is not None:
                set_prop(c, "TabIndex", self.next_tabindex())
            self._clear_raw(c)
            self._insert(target, c)
            new.append(c)
        self.build()
        self.select([self.items[c] for c in new if c in self.items])
        self.changed()

    def ask_control_array(self, name):
        """VB6's question (DEV-8/me10-01): "You already have a control named 'X'.  Do you want to create a control
        array?" Yes / No / Help; Esc and the close button do nothing; Help opens the help and asks again"""
        from . import APP_NAME

        class Box(QMessageBox):
            def reject(self):                   # Esc / the close button: VB6 keeps asking
                pass
        while True:
            box = Box(QMessageBox.Icon.Warning, APP_NAME,
                      f"You already have a control named '{name}'.  Do you want to create a control array?", parent=self)
            yes = box.addButton(QMessageBox.StandardButton.Yes)
            box.addButton(QMessageBox.StandardButton.No)
            hlp = box.addButton(QMessageBox.StandardButton.Help)
            box.setDefaultButton(yes)
            self._array_box = box
            box.exec()
            if box.clickedButton() is hlp:
                if hasattr(self.main, "help_keyword"):
                    self.main.help_keyword("control arrays")
                continue
            return box.clickedButton() is yes

    def join_array(self, existing, c):
        """c becomes the next element of the array of the controls named as it (the first one element 0)"""
        idx = [geti(x, "Index", -1) for x in existing]
        if all(i < 0 for i in idx):
            for x in existing:
                set_prop(x, "Index", 0)
            idx = [0]
        set_prop(c, "Index", max(idx) + 1)

    def _clear_raw(self, c):
        """a pasted control is written in the layout of VB6 at its new depth"""
        c.raw_begin = c.raw_end = None
        for it in c.items:
            if isinstance(it, Property):
                it.raw = None
            elif isinstance(it, PropertyBlock):
                it.raw_begin = it.raw_end = None
                for p in it.items:
                    p.raw = None
        for ch in c.children:
            self._clear_raw(ch)


    # ------------------------------------------------------------ the Properties window
    UNITS = {1: (1, 1), 2: (20, 20), 3: (15, 15), 4: (120, 240), 5: (1440, 1440), 6: (1440 / 25.4, 1440 / 25.4),
             7: (1440 / 2.54, 1440 / 2.54)}

    def inner_twips(self, ctl):
        """the inside of a form or picture box in twips (what ScaleWidth/ScaleHeight measure)"""
        if ctl is self.root:
            return geti(ctl, "ClientWidth", 4680), geti(ctl, "ClientHeight", 3195)
        b = 0
        if geti(ctl, "BorderStyle", 1):
            b = 2 * (2 if geti(ctl, "Appearance", 1) else 1) * TPP
        return geti(ctl, "Width", 1215) - b, geti(ctl, "Height", 1215) - b

    @staticmethod
    def _single(x):
        x = float(f"{x:.7g}")
        return int(x) if x == int(x) else x

    def rescale(self, ctl):
        """ScaleWidth/ScaleHeight from the size, in the units of ScaleMode (not for a User scale)"""
        mode = geti(ctl, "ScaleMode", 1)
        if mode == 0:
            return
        ux, uy = self.UNITS.get(mode, (1, 1))
        w, h = self.inner_twips(ctl)
        V.put(ctl, "ScaleHeight", self._single(h / uy))
        V.put(ctl, "ScaleWidth", self._single(w / ux))

    def set_property(self, ctl, p, value):
        """a change made in the Properties window; an error message or None"""
        if p.kind == "name":
            return self.rename(ctl, str(value).strip())
        if p.kind == "readonly":
            return None
        root = self.root
        t = ctl.type_name
        if p.kind == "index":
            if value is not None and any(c is not ctl and c.name.lower() == ctl.name.lower() and c.index == value
                                         for c in root.walk()):
                return "Invalid property value"
        self.checkpoint()
        if p.kind == "font":
            f = dict(value)
            if f.get("Charset") is None:
                f["Charset"] = V.CHARSET_OF_CODEPAGE.get(self.module.project.encoding, 0)
            default = V.font_of(root) if ctl is not root else V.font_of(Control("VB.Form", "x"))
            V.set_font(ctl, f, default)
            if t in ("ListBox", "FileListBox"):
                V.put(ctl, "Height", integral_height(ctl, t, geti(ctl, "Height", 1035)))
        elif p.kind == "formpos":
            x, y, w, h = V.outer_rect(ctl)
            v = int(round(float(value)))
            if p.name == "Left": x = v
            elif p.name == "Top": y = v
            elif p.name == "Width": w = max(TPP * 20, v)
            else: h = max(0, v)
            V.set_outer_rect(ctl, x, y, w, h)
            self.rescale(ctl)
        elif p.kind == "pos" and t != "Line" and t not in FIXED and ctl is not root:
            it = self.items.get(ctl)
            r = it.rect_twips() if it else QRect(geti(ctl, "Left"), geti(ctl, "Top"), geti(ctl, "Width"), geti(ctl, "Height"))
            x, y, w, h = r.x(), r.y(), r.width(), r.height()
            v = int(round(float(value)))
            if p.name == "Left": x = v
            elif p.name == "Top": y = v
            elif p.name == "Width": w = max(0, v)
            else: h = max(0, v)
            if it is not None:
                self.set_rect(it, x, y, w, h)
            else:
                for k, n in (("Left", x), ("Top", y), ("Width", w), ("Height", h)):
                    V.put(ctl, k, n)
        elif p.kind == "scale":
            V.apply(ctl, p, value)
            V.apply(ctl, V.prop_def(t, "ScaleMode"), 0)       # VB6: a scale of your own is a User scale
        elif p.name == "TabIndex":
            self.set_tabindex(ctl, int(value))
        elif p.name == "ScaleMode":
            V.apply(ctl, p, value)
            if value != 0:
                V.remove(ctl, "ScaleLeft")
                V.remove(ctl, "ScaleTop")
                self.rescale(ctl)
        elif t in ("Form", "MDIForm") and p.name == "BorderStyle":
            outer = V.outer_rect(ctl)                 # Windows keeps the window, the client area changes
            V.apply(ctl, p, value)
            V.set_outer_rect(ctl, *outer)
            self.rescale(ctl)
            self._border_effects(ctl, int(value))
        elif p.name in ("List", "ItemData") and t in ("ListBox", "ComboBox"):
            V.apply(ctl, p, value)
            self._sync_itemdata(ctl)
        elif t == "Image" and (p.name == "Picture" or p.name == "Stretch"):
            V.apply(ctl, p, value)
            self._image_autosize(ctl)
        else:
            V.apply(ctl, p, value)
            if t == "ComboBox" and p.name == "Style" and value == 2:
                V.remove(ctl, "Text")                 # a Dropdown List has no Text (VB6 oracle d_props)
            if t == "TextBox" and p.name == "PasswordChar":
                V.apply(ctl, V.prop_def("TextBox", "IMEMode"), 3 if value else 0)
            if t == "ListBox" and p.name == "IntegralHeight" and value:
                V.put(ctl, "Height", integral_height(ctl, t, geti(ctl, "Height", 1035)))
            if t == "PictureBox" and p.name in ("BorderStyle", "Appearance"):
                self.rescale(ctl)
            if p.name in ("Align", "Height", "Width"):
                self.realign()
                if t == "PictureBox":
                    self.rescale(ctl)
            if t == "CoolBar" and p.name in ("BandBorders", "FixedOrder", "VariantHeight", "Width", "Orientation"):
                V.cool_place(ctl)               # its bands and their children as the runtime lays them out
        self.build()
        self.select_control(ctl) if ctl is not root else self.select([])
        self.changed()
        return None

    def _sync_itemdata(self, ctl):
        """VB6 keeps an ItemData for every item of the List (0 when none was given)"""
        items = self.module.frx_value(ctl, "List") or []
        data = list(self.module.frx_value(ctl, "ItemData") or [])
        p = V.prop_def(ctl.type_name, "ItemData")
        if not items:
            V.apply(ctl, p, [])
            return
        data = (data + [0] * len(items))[:len(items)]
        if self.module.frx_value(ctl, "ItemData") != data:
            V.apply(ctl, p, data)

    def _image_autosize(self, ctl):
        """an Image that does not stretch takes the size of its picture"""
        if geti(ctl, "Stretch"):
            return
        data = self.module.frx_value(ctl, "Picture")
        if not data:
            return
        from PyQt6.QtGui import QImage
        img = QImage()
        if img.loadFromData(bytes(data)):
            V.put(ctl, "Height", img.height() * TPP)
            V.put(ctl, "Width", img.width() * TPP)

    def set_tabindex(self, ctl, n):
        """VB6 keeps the tab order whole: the others move up or down to make room"""
        ctls = sorted([c for c in self.root.walk() if c.prop("TabIndex") is not None], key=lambda c: geti(c, "TabIndex"))
        if ctl in ctls:
            ctls.remove(ctl)
        n = max(0, min(n, len(ctls)))
        ctls.insert(n, ctl)
        for i, c in enumerate(ctls):
            if geti(c, "TabIndex", -1) != i or c is ctl:
                V.put(c, "TabIndex", i)

    def _border_effects(self, form, bs):
        """VB6 changes other properties with BorderStyle: a Fixed Dialog has no Min/Max buttons
        and no taskbar button, a ToolWindow no taskbar button"""
        if bs == 3:
            for name in ("MaxButton", "MinButton", "ShowInTaskbar"):
                V.apply(form, V.prop_def("Form", name), False)
        elif bs in (4, 5):
            V.apply(form, V.prop_def("Form", "ShowInTaskbar"), False)

    def rename(self, ctl, new):
        import re
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,39}", new or ""):
            return "Invalid property value"
        if new == ctl.name:
            return None
        if ctl is self.root:
            return self.main.rename_module(self.module, new)
        if new.lower() in self.form_names():
            return f"Ambiguous name detected: {new}"
        if new.lower() != ctl.name.lower():
            existing = [c for c in self.root.walk() if c.name.lower() == new.lower()]
            if existing:
                is_array = any(c is not self.root and geti(c, "Index", -1) >= 0 for c in existing)
                if any(c.type_name != ctl.type_name for c in existing):
                    # VB6 (DEV-8/me10-09, 11): an array of another type / a control of another type
                    return "Control must be same type as rest of array" if is_array else f"'{new}' is a control name"
                if ctl.index is not None:       # a whole array into an other name: not VB6's case (me10)
                    return f"You already have a control named '{new}'."
                if not is_array and not self.ask_control_array(new):
                    return None                 # No: the name stays
                # VB6 (DEV-8/me10-05, 10): Yes, or an array of that name already (no question): the next element
                self.checkpoint()
                self.join_array(existing, ctl)
                ctl.name = existing[0].name
                ctl.raw_begin = None
                self.build()
                self.select_control(ctl)
                self.changed()
                return None
        self.checkpoint()
        old = ctl.name
        targets = [c for c in self.root.walk() if c.name.lower() == old.lower()] if ctl.index is not None else [ctl]
        for c in targets:
            c.name = new
            c.raw_begin = None
        self.build()
        self.select_control(ctl)
        self.changed()
        return None

    # ------------------------------------------------------------ menus
    def menu_editor(self):
        from .menueditor import MenuEditor
        others = [c.name for c in self.root.walk() if c is not self.root and c.type_name != "Menu"]
        dlg = MenuEditor(self, self.root, others)
        if dlg.exec() and dlg.result_menus is not None:
            self.set_menus(dlg.result_menus)

    def set_menus(self, menus):
        self.checkpoint()
        root = self.root
        outer = V.outer_rect(root)                    # adding a menu bar keeps the window
        root.children = [c for c in root.children if c.type_name != "Menu"]
        for m in menus:
            m.parent = root
        root.children += menus
        normalize_children(root)
        V.set_outer_rect(root, *outer)
        self.rescale(root)
        self.build()
        self.changed()

    # ------------------------------------------------------------ code
    def open_event(self, ctl):
        """double click: the default event procedure of the control (or Form_Load)"""
        self.main.open_event_procedure(self.module, ctl)


# ---------------------------------------------------------------- the Toolbox
def tool_icon(cls, size=24):
    from PyQt6.QtGui import QPixmap, QIcon
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setPen(QPen(QColor("#303030"), 1))
    r = QRect(3, 5, size - 6, size - 10)
    if cls is None:                            # the pointer
        poly = QPolygon([QPoint(8, 4), QPoint(8, 19), QPoint(12, 15), QPoint(15, 21), QPoint(17, 20), QPoint(14, 14), QPoint(19, 14)])
        p.setBrush(QColor("#FFFFFF"))
        p.drawPolygon(poly)
    elif cls == "PictureBox":
        p.setBrush(QColor("#FFFFFF")); p.drawRect(r)
        p.setBrush(QColor("#4A90D9")); p.drawPolygon(QPolygon([QPoint(5, 17), QPoint(11, 10), QPoint(19, 17)]))
    elif cls == "Label":
        f = QFont(); f.setPixelSize(16); f.setBold(True); p.setFont(f); p.drawText(QRect(0, 0, size, size), Qt.AlignmentFlag.AlignCenter, "A")
    elif cls == "TextBox":
        p.setBrush(QColor("#FFFFFF")); p.drawRect(r)
        f = QFont(); f.setPixelSize(10); p.setFont(f); p.drawText(r, Qt.AlignmentFlag.AlignCenter, "ab|")
    elif cls == "Frame":
        p.drawRect(r.adjusted(0, 3, 0, 0)); p.fillRect(QRect(6, 5, 9, 6), QColor("#F0F0F0"))
        f = QFont(); f.setPixelSize(8); p.setFont(f); p.drawText(QRect(6, 3, 12, 8), Qt.AlignmentFlag.AlignCenter, "xy")
    elif cls == "CommandButton":
        p.setBrush(QColor("#E8E8E8")); p.drawRoundedRect(r, 2, 2)
    elif cls == "CheckBox":
        p.setBrush(QColor("#FFFFFF")); p.drawRect(QRect(6, 6, 12, 12))
        p.setPen(QPen(QColor("#000000"), 2)); p.drawPolyline(QPolygon([QPoint(8, 12), QPoint(11, 15), QPoint(16, 8)]))
    elif cls == "OptionButton":
        p.setBrush(QColor("#FFFFFF")); p.drawEllipse(QRect(6, 6, 12, 12))
        p.setBrush(QColor("#000000")); p.drawEllipse(QRect(10, 10, 4, 4))
    elif cls in ("ComboBox", "DriveListBox"):
        p.setBrush(QColor("#FFFFFF")); p.drawRect(QRect(3, 8, 18, 8))
        p.setBrush(QColor("#C0C0C0")); p.drawRect(QRect(15, 8, 6, 8))
        if cls == "DriveListBox":
            p.setBrush(QColor("#808080")); p.drawRect(QRect(5, 10, 6, 4))
    elif cls in ("ListBox", "FileListBox", "DirListBox"):
        p.setBrush(QColor("#FFFFFF")); p.drawRect(r)
        for i in range(3):
            p.drawLine(6, 8 + 3 * i, 16, 8 + 3 * i)
        if cls == "DirListBox":
            p.setBrush(QColor("#E8C34A")); p.drawRect(QRect(12, 13, 7, 5))
    elif cls == "HScrollBar":
        p.setBrush(QColor("#E0E0E0")); p.drawRect(QRect(2, 8, 20, 8)); p.setBrush(QColor("#A0A0A0")); p.drawRect(QRect(9, 9, 6, 6))
    elif cls == "VScrollBar":
        p.setBrush(QColor("#E0E0E0")); p.drawRect(QRect(8, 2, 8, 20)); p.setBrush(QColor("#A0A0A0")); p.drawRect(QRect(9, 9, 6, 6))
    elif cls == "Timer":
        p.setBrush(QColor("#FFFFFF")); p.drawEllipse(QRect(4, 4, 16, 16)); p.drawLine(12, 12, 12, 6); p.drawLine(12, 12, 16, 12)
    elif cls == "Shape":
        p.setBrush(QColor("#E05050")); p.drawRect(QRect(3, 9, 10, 10)); p.setBrush(QColor("#5080E0")); p.drawEllipse(QRect(10, 4, 11, 11))
    elif cls == "Line":
        p.setPen(QPen(QColor("#000000"), 2)); p.drawLine(4, 19, 20, 5)
    elif cls == "Image":
        p.setBrush(QColor("#9AD0F5")); p.drawRect(r); p.setBrush(QColor("#3A8F3A")); p.drawPolygon(QPolygon([QPoint(4, 18), QPoint(10, 11), QPoint(14, 15), QPoint(16, 13), QPoint(20, 18)]))
    elif cls == "Data":                        # the four arrows of the Data control
        p.setBrush(QColor("#FFFFFF")); p.drawRect(QRect(2, 7, 20, 10))
        p.setBrush(QColor("#000000"))
        p.drawPolygon(QPolygon([QPoint(4, 12), QPoint(8, 9), QPoint(8, 15)]))
        p.drawPolygon(QPolygon([QPoint(20, 12), QPoint(16, 9), QPoint(16, 15)]))
    elif cls == "UserControl":                 # a user control: a box with a small form in it
        p.setBrush(QColor("#F0F0F0")); p.drawRect(QRect(3, 3, size - 6, size - 6))
        p.setBrush(QColor("#4A90D9")); p.drawRect(QRect(6, 6, size - 12, 4))
        p.setBrush(QColor("#E8E8E8")); p.drawRect(QRect(8, 13, size - 16, 6))
    else:
        ocx_icon(p, cls, size)
    p.end()
    return QIcon(pm)


def ocx_icon(p, cls, size=24):
    """the Toolbox picture of a control of an .ocx (drawn: no picture of Microsoft's)"""
    white, face, blue = QColor("#FFFFFF"), QColor("#E8E8E8"), QColor("#4A90D9")
    if cls == "CommonDialog":
        p.setBrush(face); p.drawRect(QRect(3, 4, 18, 16)); p.setBrush(blue); p.drawRect(QRect(3, 4, 18, 4))
        p.setBrush(white); p.drawRect(QRect(13, 14, 6, 4))
    elif cls == "TabStrip" or cls == "SSTab":
        p.setBrush(face); p.drawRect(QRect(3, 9, 18, 11))
        p.setBrush(white); p.drawRect(QRect(3, 5, 7, 5)); p.setBrush(face); p.drawRect(QRect(10, 6, 7, 4))
        if cls == "SSTab":
            p.setBrush(blue); p.drawRect(QRect(6, 13, 9, 3))
    elif cls == "Toolbar":
        p.setBrush(face); p.drawRect(QRect(2, 8, 20, 9))
        for x in (4, 10, 16):
            p.setBrush(blue); p.drawRect(QRect(x, 10, 4, 5))
    elif cls == "StatusBar":
        p.setBrush(face); p.drawRect(QRect(2, 13, 20, 7)); p.drawLine(10, 14, 10, 19); p.drawLine(16, 14, 16, 19)
    elif cls == "CoolBar":                     # two rows of bands with their grippers
        p.setBrush(face); p.drawRect(QRect(2, 5, 20, 14)); p.drawLine(2, 12, 22, 12)
        p.setBrush(blue)
        for x, y in ((4, 7), (13, 7), (4, 14)):
            p.drawRect(QRect(x, y, 1, 3))
        p.setBrush(white); p.drawRect(QRect(7, 7, 4, 3)); p.drawRect(QRect(16, 7, 4, 3)); p.drawRect(QRect(7, 14, 9, 3))
    elif cls == "ProgressBar":
        p.setBrush(white); p.drawRect(QRect(2, 9, 20, 7))
        for x in (3, 7, 11):
            p.fillRect(QRect(x, 10, 3, 5), blue)
    elif cls in ("TreeView", "ListView"):
        p.setBrush(white); p.drawRect(QRect(3, 3, 18, 18))
        if cls == "TreeView":
            p.drawLine(7, 7, 7, 17); p.drawLine(7, 11, 11, 11); p.drawLine(7, 17, 11, 17)
            for y in (5, 9, 15):
                p.fillRect(QRect(11 if y > 5 else 5, y, 4, 4), blue)
        else:
            for y in (6, 11, 16):
                p.fillRect(QRect(5, y, 4, 3), blue); p.drawLine(11, y + 1, 18, y + 1)
    elif cls in ("Adodc", "MSRDC"):
        p.setBrush(white); p.drawRect(QRect(2, 7, 20, 10))
        p.setBrush(QColor("#000000"))
        p.drawPolygon(QPolygon([QPoint(4, 12), QPoint(8, 9), QPoint(8, 15)]))
        p.drawPolygon(QPolygon([QPoint(20, 12), QPoint(16, 9), QPoint(16, 15)]))
        p.setBrush(blue); p.drawRect(QRect(10, 9, 4, 6))
    elif cls == "DataRepeater":                # three rows of the same control
        p.setBrush(white); p.drawRect(QRect(2, 2, 20, 20))
        for y in (4, 10, 16):
            p.fillRect(QRect(4, y, 3, 4), face); p.fillRect(QRect(9, y + 1, 11, 2), blue)
    elif cls in ("DataGrid", "DBGrid"):
        p.setBrush(white); p.drawRect(QRect(2, 3, 20, 18))
        p.fillRect(QRect(3, 4, 19, 4), face); p.fillRect(QRect(3, 4, 4, 17), face)
        for y in (12, 16):
            p.drawLine(3, y, 21, y)
        p.drawLine(13, 4, 13, 20)
    elif cls in ("DataList", "DataCombo", "DBList", "DBCombo"):
        combo = cls in ("DataCombo", "DBCombo")
        if combo:
            p.setBrush(white); p.drawRect(QRect(2, 8, 20, 8)); p.setBrush(face); p.drawRect(QRect(15, 8, 7, 8))
        else:
            p.setBrush(white); p.drawRect(QRect(3, 3, 18, 18))
            for y in (7, 11, 15):
                p.drawLine(6, y, 15, y)
        # DataList / DataCombo (OLE DB): a blue mark; DBList / DBCombo (the Data control): a yellow one
        p.fillRect(QRect(4, 18 if combo else 17, 6, 3), blue if cls.startswith("Data") else QColor("#C0A000"))
    elif cls == "ImageList":
        for k, c in enumerate((QColor("#E05050"), QColor("#50A050"), blue)):
            p.setBrush(c); p.drawRect(QRect(3 + k * 4, 3 + k * 4, 10, 10))
    elif cls == "Slider":
        p.drawLine(3, 12, 21, 12)
        for x in range(3, 22, 4):
            p.drawLine(x, 17, x, 19)
        p.setBrush(face); p.drawPolygon(QPolygon([QPoint(9, 7), QPoint(14, 7), QPoint(14, 14), QPoint(11, 17), QPoint(9, 14)]))
    elif cls in ("ImageCombo", "DTPicker"):
        p.setBrush(white); p.drawRect(QRect(2, 8, 20, 8)); p.setBrush(face); p.drawRect(QRect(16, 8, 6, 8))
        p.fillRect(QRect(4, 10, 4, 4), blue if cls == "ImageCombo" else QColor("#E05050"))
    elif cls == "Animation":
        p.setBrush(white); p.drawRect(QRect(3, 5, 18, 14)); p.setBrush(blue)
        p.drawPolygon(QPolygon([QPoint(9, 8), QPoint(16, 12), QPoint(9, 16)]))
    elif cls == "UpDown":
        p.setBrush(face); p.drawRect(QRect(8, 3, 9, 9)); p.drawRect(QRect(8, 12, 9, 9)); p.setBrush(QColor("#000000"))
        p.drawPolygon(QPolygon([QPoint(10, 9), QPoint(15, 9), QPoint(12, 6)])); p.drawPolygon(QPolygon([QPoint(10, 15), QPoint(15, 15), QPoint(12, 18)]))
    elif cls == "MonthView":
        p.setBrush(white); p.drawRect(QRect(3, 3, 18, 18)); p.fillRect(QRect(4, 4, 17, 4), blue)
        for y in (10, 14, 18):
            for x in (5, 9, 13, 17):
                p.drawPoint(x, y)
    elif cls == "FlatScrollBar":
        p.setBrush(white); p.drawRect(QRect(2, 8, 20, 8)); p.fillRect(QRect(9, 9, 6, 6), QColor("#A0A0A0"))
    elif cls == "MSFlexGrid":
        p.setBrush(white); p.drawRect(QRect(2, 3, 20, 18)); p.fillRect(QRect(3, 4, 19, 4), face); p.fillRect(QRect(3, 4, 5, 17), face)
        for y in (8, 12, 16):
            p.drawLine(2, y, 21, y)
        for x in (8, 15):
            p.drawLine(x, 3, x, 20)
    elif cls == "RichTextBox":
        p.setBrush(white); p.drawRect(QRect(2, 4, 20, 16))
        f = QFont(); f.setPixelSize(9); f.setBold(True); p.setFont(f); p.drawText(QRect(3, 4, 10, 16), Qt.AlignmentFlag.AlignCenter, "A")
        f.setBold(False); f.setItalic(True); p.setFont(f); p.setPen(QColor("#E05050")); p.drawText(QRect(11, 4, 10, 16), Qt.AlignmentFlag.AlignCenter, "a")
    elif cls == "MSChart":
        p.drawLine(4, 3, 4, 20); p.drawLine(4, 20, 21, 20)
        for x, h, c in ((6, 8, blue), (11, 13, QColor("#E05050")), (16, 6, QColor("#50A050"))):
            p.fillRect(QRect(x, 20 - h, 4, h), c)
    elif cls == "MMControl":
        p.setBrush(face); p.drawRect(QRect(2, 7, 20, 10)); p.setBrush(QColor("#000000"))
        p.drawPolygon(QPolygon([QPoint(5, 9), QPoint(10, 12), QPoint(5, 15)])); p.drawRect(QRect(13, 9, 2, 6)); p.drawRect(QRect(17, 9, 2, 6))
    elif cls == "MSComm":
        p.setBrush(QColor("#303030")); p.drawRoundedRect(QRect(4, 6, 16, 9), 2, 2)
        p.setBrush(QColor("#FFD040"))
        for x in (6, 9, 12, 15):
            p.drawEllipse(QRect(x, 8, 2, 2))
        p.drawLine(12, 15, 12, 20); p.drawLine(12, 20, 20, 20)
    elif cls == "Winsock":                      # two computers and the line between them
        for x in (2, 13):
            p.setBrush(white); p.drawRect(QRect(x, 4, 9, 7)); p.fillRect(QRect(x + 1, 5, 8, 6), blue)
            p.drawLine(x + 4, 11, x + 4, 13); p.drawLine(x + 2, 13, x + 7, 13)
        pen = p.pen(); p.setPen(QColor("#E05050"))
        p.drawLine(6, 14, 6, 19); p.drawLine(6, 19, 17, 19); p.drawLine(17, 19, 17, 14)
        p.setPen(pen)
    elif cls == "WebBrowser":                   # a page with a globe
        p.setBrush(white); p.drawRect(QRect(2, 3, 20, 18)); p.fillRect(QRect(3, 4, 19, 3), blue)
        p.setBrush(QColor("#7FB8E8")); p.drawEllipse(QRect(7, 9, 10, 10))
        p.setBrush(Qt.BrushStyle.NoBrush); p.drawEllipse(QRect(10, 9, 4, 10)); p.drawLine(7, 14, 17, 14)
    elif cls == "Inet":                         # a globe
        p.setBrush(QColor("#7FB8E8")); p.drawEllipse(QRect(3, 3, 18, 18))
        p.setBrush(Qt.BrushStyle.NoBrush); p.drawEllipse(QRect(8, 3, 8, 18))
        p.drawLine(3, 12, 21, 12); p.drawLine(5, 7, 19, 7); p.drawLine(5, 17, 19, 17)
    elif cls in ("MAPISession", "MAPIMessages"):    # an envelope (a key on it: the session)
        p.setBrush(white); p.drawRect(QRect(2, 6, 20, 13))
        p.drawLine(2, 6, 12, 13); p.drawLine(12, 13, 22, 6)
        if cls == "MAPISession":
            p.setBrush(QColor("#FFD040")); p.drawEllipse(QRect(14, 13, 6, 6))
    elif cls == "MaskEdBox":                    # a box with ##
        p.setBrush(white); p.drawRect(QRect(2, 7, 20, 10))
        f = QFont(); f.setPixelSize(9); f.setBold(True); p.setFont(f)
        p.setPen(blue); p.drawText(QRect(2, 7, 20, 10), Qt.AlignmentFlag.AlignCenter, "##")
    elif cls == "PictureClip":
        p.setBrush(white); p.drawRect(QRect(3, 3, 18, 18)); p.drawLine(12, 3, 12, 20); p.drawLine(3, 12, 20, 12)
        p.fillRect(QRect(4, 4, 8, 8), blue)
    else:
        p.setBrush(face); p.drawRect(QRect(3, 3, 18, 18))
        f = QFont(); f.setPixelSize(8); p.setFont(f); p.drawText(QRect(3, 3, 18, 18), Qt.AlignmentFlag.AlignCenter, (cls or "?")[:3])


TOOL_MIME = "application/x-lbstudio-tool"


class ToolTab:
    """a tab of the user in the Toolbox; VB6 does not check the names: two tabs may have one name, "General"
    too, or none (DEV-8/me7-04, 05), so a tab is this object, not its name"""

    def __init__(self, name):
        self.name = name

    def __repr__(self):
        return f"ToolTab({self.name!r})"


class _TabButton(QPushButton):
    """a tab of the Toolbox: a button (General, MyTab ...), as VB6's (DEV-8/me6); a tool dragged over it opens
    the tab, dropped on it goes into the tab. tab None: General"""

    def __init__(self, box, tab):
        super().__init__("General" if tab is None else tab.name)
        self.box = box
        self.tab = tab
        self.setAcceptDrops(True)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setMinimumHeight(self.fontMetrics().height() + 8)     # (a tab without a name keeps its height)
        self.clicked.connect(lambda: box.set_tab(self.tab))
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(lambda pos: box.tab_menu(self.tab, self.mapToGlobal(pos)))

    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat(TOOL_MIME):
            e.acceptProposedAction()
            if self.box.cur is not self.tab:   # (VB6: the tab opens under the tool)
                QTimer.singleShot(0, lambda: self.box.set_tab(self.tab))

    def dropEvent(self, e):
        cls = bytes(e.mimeData().data(TOOL_MIME)).decode()
        self.box.move_tool(cls, self.tab)
        e.acceptProposedAction()


class Toolbox(QWidget):
    """the VB6 Toolbox: the tab General and the user's tabs (DEV-8/me6, me7), each with its tools (the intrinsic
    controls, those of the project's components (.ocx) and its user controls). Add Tab... puts the new tab right
    after the open one; a tool dragged onto a tab moves there; a deleted tab gives its tools back to General, in
    their places; the controls Components adds go into the open tab. The tabs are LinBASIC Studio's settings
    (toolbox_tabs), not the project's: they stay after the IDE is closed, as in VB6.
    A tab is given as a ToolTab, None (General), or a name (the first tab of that name; "General": General)."""

    def __init__(self, main):
        super().__init__()
        self.main = main
        self.lay = QVBoxLayout(self)
        self.lay.setContentsMargins(0, 0, 0, 0)
        self.lay.setSpacing(1)
        self.top = QVBoxLayout()
        self.top.setSpacing(1)
        self.panel = QWidget()
        self.panel.setAcceptDrops(True)
        self.panel.dragEnterEvent = self._panel_drag
        self.panel.dropEvent = self._panel_drop
        self.grid = QGridLayout(self.panel)
        self.grid.setSpacing(1)
        self.grid.setContentsMargins(2, 2, 2, 2)
        self.bottom = QVBoxLayout()
        self.bottom.setSpacing(1)
        self.lay.addLayout(self.top)
        self.lay.addWidget(self.panel, 1)
        self.lay.addLayout(self.bottom)
        self.buttons = {}                       # every tool there is (None: the Pointer), in all the tabs
        self.entries = []                       # (cls, tip) in VB6's order: General shows them in it
        self.components = None
        self.adding = False                     # Project > Components: the new controls go into the open tab
        self.tabs = []                          # the user's tabs (ToolTab), in order
        self.tab_of = {}                        # tool -> its ToolTab (not there: General)
        self.cur = None                         # the open tab (None: General)
        self.tab_buttons = []
        self._drag_from = None
        self._load()
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred)
        self.set_components([])

    # ------------------------------------------------------------ the tabs
    def tab(self, t):
        """a ToolTab from a ToolTab, None or a name"""
        if t is None or isinstance(t, ToolTab):
            return t if t is None or t in self.tabs else None
        if t == "General":
            return None
        return next((x for x in self.tabs if x.name == t), None)

    @property
    def cur_tab(self):
        """the name of the open tab"""
        return "General" if self.cur is None else self.cur.name

    def names(self):
        return [t.name for t in self.tabs]

    def _settings(self):
        return getattr(self.main, "S", None)

    def _load(self):
        S = self._settings()
        self.tabs, self.tab_of = [], {}
        for d in (S["toolbox_tabs"] if S is not None else []) or []:
            t = ToolTab(str(d.get("name", "")))
            self.tabs.append(t)
            for cls in d.get("tools", []):
                self.tab_of[cls] = t

    def _save(self):
        S = self._settings()
        if S is None:
            return
        S["toolbox_tabs"] = [{"name": t.name, "tools": [c for c, x in self.tab_of.items() if x is t]}
                             for t in self.tabs]
        S.save()

    def tab_tools(self, t=None):
        """the tools a tab shows: the Pointer, then its tools in VB6's order"""
        t = self.tab(t)
        if t is None:                           # General (a tool of a tab that is no more: General)
            return [None] + [c for c, _ in self.entries if c is not None and self.tab_of.get(c) not in self.tabs]
        return [None] + [c for c, _ in self.entries if c is not None and self.tab_of.get(c) is t]

    # ------------------------------------------------------------ the tools of the project
    def set_components(self, keys):
        keys = list(keys)
        if keys == self.components:
            return
        before = {c for c, _ in self.entries}
        self.components = keys
        for b in self.buttons.values():
            b.deleteLater()
        self.buttons = {}
        entries = [(None, "Pointer")] + [(t[0], t[0]) for t in TOOLS]
        for key in keys:
            if key.startswith("uc:"):           # a user control of the project or of the group
                uc = uctl_of(key[3:])
                if uc:
                    entries.append((f"{uc[0]}.{uc[1]}", f"{uc[1]} ({uc[0]})"))
                continue
            lib = ocxlib.LIB_BY_KEY.get(key.lower())
            if lib is not None:
                entries += [(f"{lib.key}.{c.cls}", f"{c.cls} ({lib.file})") for c in lib.controls]
        self.entries = entries
        if self.adding and self.cur is not None:        # (DEV-8/me6-19: CommonDialog into the open tab Saved)
            new = [c for c, _ in entries if c is not None and c not in before]
            for c in new:
                self.tab_of[c] = self.cur
            if new:
                self._save()
        for cls, tip in entries:
            b = QToolButton(self.panel)
            b.setIcon(tool_icon("UserControl" if cls and uctl_of(cls) else cls.split(".")[-1] if cls else None))
            b.setIconSize(QSize(24, 24))
            b.setToolTip(tip)
            b.setCheckable(True)
            b.setAutoRaise(True)
            b.clicked.connect(lambda checked=False, c=cls: self._clicked(c))
            b.mouseDoubleClickEvent = lambda e, c=cls: self._double(c)
            if cls is not None:
                b.installEventFilter(self)
                b.setProperty("tool", cls)
            b.hide()
            self.buttons[cls] = b
        self.fill()
        self.reset()

    def fill(self):
        """the tab buttons (the open one and those before it on top, the others below) and the tools of the
        open tab (DEV-8/me6-05, 06, 09)"""
        if self.cur is not None and self.cur not in self.tabs:
            self.cur = None
        for lay in (self.top, self.bottom):
            while lay.count():
                w = lay.takeAt(0).widget()
                if w is not None:
                    w.deleteLater()
        self.tab_buttons = []
        order = [None] + self.tabs
        cur = order.index(self.cur)
        for i, t in enumerate(order):
            tb = _TabButton(self, t)
            (self.top if i <= cur else self.bottom).addWidget(tb)
            self.tab_buttons.append(tb)
        while self.grid.count():
            self.grid.takeAt(0)
        for b in self.buttons.values():
            b.hide()
        tools = self.tab_tools(self.cur)
        for i, cls in enumerate(tools):
            b = self.buttons.get(cls)
            if b is not None:
                self.grid.addWidget(b, i // 2, i % 2)
                b.show()
        for r in range(self.grid.rowCount() + 2):
            self.grid.setRowStretch(r, 0)
        self.grid.setRowStretch(len(tools) // 2 + 1, 1)
        dock = getattr(self.main, "dock_toolbox", None)
        if dock is not None:
            dock.setWindowTitle(self.cur_tab)

    def set_tab(self, t):
        """open a tab: its button goes up, its tools show"""
        t = self.tab(t)
        if t is self.cur:
            return
        self.cur = t
        self.reset()
        self.fill()

    # ------------------------------------------------------------ Add / Delete / Rename / Move Up, Down
    def _ask_name(self, text=""):
        from PyQt6.QtWidgets import QInputDialog
        from . import APP_NAME
        name, ok = QInputDialog.getText(self, APP_NAME, "New Tab Name:", QLineEdit.EchoMode.Normal, text)
        return name if ok else None

    def add_tab(self, name=None):
        """Add Tab...: "New Tab Name:"; any name, also one there is or none (me7-04, 05); the tab goes right after
        the open one, which stays open (me6-09, me7-05, 06)"""
        if name is None:
            name = self._ask_name()
            if name is None:                    # Cancel
                return None
        t = ToolTab(name)
        self.tabs.insert(0 if self.cur is None else self.tabs.index(self.cur) + 1, t)
        self._save()
        self.fill()
        return t

    def delete_tab(self, t):
        """Delete Tab: no question, also with tools in it; its tools go back to General, General opens
        (me6-13, 16)"""
        t = self.tab(t)
        if t is None:
            return
        self.tabs.remove(t)
        self.tab_of = {c: x for c, x in self.tab_of.items() if x is not t}
        self.cur = None
        self._save()
        self.reset()
        self.fill()

    def rename_tab(self, t, new=None):
        """Rename Tab...: the same dialog with the name in it, selected (me6-12)"""
        t = self.tab(t)
        if t is None:
            return
        if new is None:
            new = self._ask_name(t.name)
            if new is None:
                return
        t.name = new
        self._save()
        self.fill()

    def move_tab(self, t, step):
        """Move Up / Move Down: among the user's tabs; General is always first (me6-08, 10, 11)"""
        t = self.tab(t)
        if t is None:
            return
        i = self.tabs.index(t)
        j = i + step
        if not 0 <= j < len(self.tabs):
            return
        self.tabs[i], self.tabs[j] = self.tabs[j], self.tabs[i]
        self._save()
        self.fill()

    def move_tool(self, cls, t):
        """a tool dragged onto a tab: it moves there and the tab opens (me6-14); onto General: back to its place"""
        if cls is None or cls not in self.buttons:
            return
        t = self.tab(t)
        if t is None:
            self.tab_of.pop(cls, None)
        else:
            self.tab_of[cls] = t
        self._save()
        self.cur = t
        self.reset()
        self.fill()

    def tab_menu(self, t, gpos=None):
        """the menu of a tab button: Add Tab..., Delete Tab, Rename Tab... | Move Up, Move Down (General: only
        Add Tab..., me6-03, 08)"""
        from PyQt6.QtWidgets import QMenu
        t = self.tab(t)
        menu = QMenu(self)
        user = t is not None
        i = self.tabs.index(t) if user else -1
        menu.addAction("&Add Tab...", lambda: self.add_tab())
        menu.addAction("&Delete Tab", lambda: self.delete_tab(t)).setEnabled(user)
        menu.addAction("&Rename Tab...", lambda: self.rename_tab(t)).setEnabled(user)
        menu.addSeparator()
        menu.addAction("Move &Up", lambda: self.move_tab(t, -1)).setEnabled(user and i > 0)
        menu.addAction("Move Do&wn", lambda: self.move_tab(t, 1)).setEnabled(user and i < len(self.tabs) - 1)
        self._last_menu = menu
        if gpos is not None:
            menu.exec(gpos)
        return menu

    # ------------------------------------------------------------ dragging a tool to a tab
    def eventFilter(self, obj, e):
        t = e.type()
        if t == e.Type.MouseButtonPress and e.button() == Qt.MouseButton.LeftButton:
            self._drag_from = (obj, e.position().toPoint())
        elif t == e.Type.MouseMove and self._drag_from is not None and self._drag_from[0] is obj and \
                e.buttons() & Qt.MouseButton.LeftButton:
            if (e.position().toPoint() - self._drag_from[1]).manhattanLength() >= QApplication.startDragDistance():
                self._drag_from = None
                from PyQt6.QtGui import QDrag
                drag = QDrag(obj)
                md = QMimeData()
                md.setData(TOOL_MIME, obj.property("tool").encode())
                drag.setMimeData(md)
                drag.setPixmap(obj.icon().pixmap(24, 24))
                drag.exec(Qt.DropAction.MoveAction)
                obj.setDown(False)              # (the release went to the drop: the button would stay down)
                return True
        elif t == e.Type.MouseButtonRelease:
            self._drag_from = None
        return False

    def _panel_drag(self, e):
        if e.mimeData().hasFormat(TOOL_MIME):
            e.acceptProposedAction()

    def _panel_drop(self, e):
        self.move_tool(bytes(e.mimeData().data(TOOL_MIME)).decode(), self.cur)
        e.acceptProposedAction()

    # ------------------------------------------------------------ choosing a tool
    def _clicked(self, cls):
        for c, b in self.buttons.items():
            b.setChecked(c == cls)

    def _double(self, cls):
        if cls is not None:
            d = self.main.current_designer()
            if d is not None:
                d.add_default(cls)
        self.reset()

    def current(self):
        for c, b in self.buttons.items():
            if b.isChecked():
                return c
        return None

    def reset(self):
        self._clicked(None)
