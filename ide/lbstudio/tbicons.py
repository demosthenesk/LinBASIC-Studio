"""The pictures of the toolbar buttons of the Application Wizard and the Toolbar Wizard (phase 9.2): our own,
drawn here (VB6's are Microsoft's), 16 x 16 on the mask colour of the ImageList (&HC0C0C0, MaskColor 12632256,
as the VB6 wizards write it: VB6 oracle DEV-9/aw1 frmMain), saved as BMP for the .frx. The names are the ones of
the Customize Toolbar page (VB6 oracle DEV-9 aw4): New, Open ... Word Underline; the Web Browser form has Back,
Forward, Stop, Refresh, Home, Search at 24 x 24 (aw3 frmBrowser).
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
from PyQt6.QtCore import Qt, QRect, QPoint, QByteArray, QBuffer, QIODevice, QRectF
from PyQt6.QtGui import QImage, QPainter, QColor, QPen, QPolygon, QFont

MASK = QColor(0xC0, 0xC0, 0xC0)
INK, BLUE, RED, GREEN, YELLOW, WHITE = (QColor("#202020"), QColor("#2F6FD0"), QColor("#D03030"),
                                        QColor("#2E9A3A"), QColor("#F2C230"), QColor("#FFFFFF"))
GRAY, PAPER = QColor("#707070"), QColor("#FAFAFA")

# the buttons of the Customize Toolbar page, as VB6 names them (aw4); the toolbar of MDI / SDI (aw3) and of
# Explorer Style (aw17)
from .appwizgen import TOOLBAR_MDI, TOOLBAR_EXPLORER  # noqa: E402,F401  (the lists of the wizard)
ALL = ["Align Left", "Align Right", "Arc", "Back", "Bold", "Button", "Camera", "Center", "Copy", "Cut", "Delete",
       "Disconnect Net Drive", "Double Underline", "Drawing", "Ellipse", "Find", "Forward", "Freeform", "Help",
       "Help What's This", "Italic", "Justify", "Line", "Line Style", "Macro", "Map Network Drive", "New", "Open",
       "Paste", "Print", "Properties", "Rectangle", "Redo", "Save", "Small Caps", "Sort Ascending",
       "Sort Descending", "Spell Check", "Strike Through", "Sum", "Tab Center", "Tab Decimal", "Tab Left",
       "Tab Right", "Underline", "Undo", "Up One Level", "View Details", "View Large Icons", "View List",
       "View Small Icons", "Word Underline"]
BROWSER = ["Back", "Forward", "Stop", "Refresh", "Home", "Search"]


def _pen(p, c=INK, w=1.0):
    p.setPen(QPen(c, w))


def _page(p, x=3, y=1, w=10, h=14, fold=True):
    _pen(p)
    p.setBrush(PAPER)
    if fold:
        p.drawPolygon(QPolygon([QPoint(x, y), QPoint(x + w - 4, y), QPoint(x + w, y + 4), QPoint(x + w, y + h),
                                QPoint(x, y + h)]))
        p.drawLine(x + w - 4, y, x + w - 4, y + 4)
        p.drawLine(x + w - 4, y + 4, x + w, y + 4)
    else:
        p.drawRect(x, y, w, h)


def _lines(p, rows, x0=4, x1=12, align="left"):
    _pen(p, GRAY)
    for i, y in enumerate(rows):
        w = (x1 - x0) if i % 2 == 0 else (x1 - x0) - 3
        if align == "left":
            p.drawLine(x0, y, x0 + w, y)
        elif align == "right":
            p.drawLine(x1 - w, y, x1, y)
        elif align == "center":
            m = (x0 + x1) // 2
            p.drawLine(m - w // 2, y, m + w // 2, y)
        else:
            p.drawLine(x0, y, x1, y)


def _text(p, s, rect, bold=False, italic=False, under=False, strike=False, size=11, color=INK):
    f = QFont("DejaVu Sans")
    f.setPixelSize(size)
    f.setBold(bold)
    f.setItalic(italic)
    f.setUnderline(under)
    f.setStrikeOut(strike)
    p.setFont(f)
    _pen(p, color)
    p.drawText(QRect(*rect), Qt.AlignmentFlag.AlignCenter, s)


def _arrow(p, left, color=GREEN, y=8):
    _pen(p, color.darker(140))
    p.setBrush(color)
    if left:
        p.drawPolygon(QPolygon([QPoint(2, y), QPoint(8, y - 6), QPoint(8, y - 3), QPoint(14, y - 3), QPoint(14, y + 3),
                                QPoint(8, y + 3), QPoint(8, y + 6)]))
    else:
        p.drawPolygon(QPolygon([QPoint(14, y), QPoint(8, y - 6), QPoint(8, y - 3), QPoint(2, y - 3), QPoint(2, y + 3),
                                QPoint(8, y + 3), QPoint(8, y + 6)]))


def _curve_arrow(p, back):
    _pen(p, BLUE, 2)
    p.setBrush(Qt.BrushStyle.NoBrush)
    if back:
        p.drawArc(QRectF(4, 4, 9, 9), 90 * 16, -230 * 16)
        p.setBrush(BLUE)
        p.drawPolygon(QPolygon([QPoint(2, 6), QPoint(7, 3), QPoint(7, 9)]))
    else:
        p.drawArc(QRectF(3, 4, 9, 9), 90 * 16, 230 * 16)
        p.setBrush(BLUE)
        p.drawPolygon(QPolygon([QPoint(14, 6), QPoint(9, 3), QPoint(9, 9)]))


def _drive(p, net_ok):
    _pen(p)
    p.setBrush(QColor("#B8B8B8"))
    p.drawRect(2, 3, 12, 6)
    p.fillRect(QRect(3, 6, 3, 2), GREEN)
    _pen(p, INK)
    p.drawLine(8, 9, 8, 12)
    p.drawLine(3, 12, 13, 12)
    if not net_ok:
        _pen(p, RED, 2)
        p.drawLine(10, 10, 14, 14)
        p.drawLine(14, 10, 10, 14)


def _view(p, kind):
    _pen(p, BLUE)
    p.setBrush(BLUE)
    if kind == "large":
        for x in (2, 9):
            for y in (2, 9):
                p.drawRect(x, y, 4, 4)
    elif kind == "small":
        for x in (2, 9):
            for y in (2, 7, 12):
                p.drawRect(x, y, 2, 2)
                _pen(p, GRAY)
                p.drawLine(x + 3, y + 1, x + 5, y + 1)
                _pen(p, BLUE)
    elif kind == "list":
        for y in (2, 6, 10, 14):
            p.drawRect(2, y - 1, 2, 2)
            _pen(p, GRAY)
            p.drawLine(6, y, 13, y)
            _pen(p, BLUE)
    else:                                       # details
        p.setBrush(Qt.BrushStyle.NoBrush)
        _pen(p, INK)
        p.drawRect(1, 2, 14, 12)
        p.fillRect(QRect(2, 3, 13, 2), QColor("#9AB8E8"))
        _pen(p, GRAY)
        for y in (7, 10, 13):
            p.drawLine(2, y, 14, y)
        p.drawLine(6, 3, 6, 13)


def _tab(p, kind):
    _lines(p, (13,), 1, 15, "full")
    _pen(p, INK, 1)
    p.drawLine(8, 3, 8, 11)
    if kind == "left":
        p.drawLine(8, 11, 13, 11)
    elif kind == "right":
        p.drawLine(3, 11, 8, 11)
    elif kind == "center":
        p.drawLine(4, 11, 12, 11)
    else:
        p.drawLine(4, 11, 12, 11)
        p.drawPoint(11, 8)
        p.drawPoint(11, 9)


def _sort(p, up):
    _text(p, "A" if up else "Z", (1, 0, 9, 8), bold=True, size=8)
    _text(p, "Z" if up else "A", (1, 8, 9, 8), bold=True, size=8)
    _pen(p, BLUE, 1)
    p.drawLine(12, 2, 12, 14)
    p.setBrush(BLUE)
    p.drawPolygon(QPolygon([QPoint(10, 11), QPoint(14, 11), QPoint(12, 14)]))


def draw(p, name):
    """one button picture on a 16 x 16 painter"""
    n = name
    if n == "New":
        _page(p)
    elif n == "Open":
        _pen(p, QColor("#8A6A10"))
        p.setBrush(YELLOW)
        p.drawPolygon(QPolygon([QPoint(1, 4), QPoint(6, 4), QPoint(7, 5), QPoint(13, 5), QPoint(13, 13), QPoint(1, 13)]))
        p.setBrush(QColor("#F7DA7A"))
        p.drawPolygon(QPolygon([QPoint(3, 8), QPoint(15, 8), QPoint(13, 13), QPoint(1, 13)]))
    elif n == "Save":
        _pen(p)
        p.setBrush(BLUE)
        p.drawRect(2, 2, 12, 12)
        p.fillRect(QRect(4, 3, 8, 5), WHITE)
        p.fillRect(QRect(5, 10, 6, 4), QColor("#C8C8C8"))
    elif n == "Print":
        _pen(p)
        p.setBrush(PAPER)
        p.drawRect(4, 1, 8, 5)
        p.setBrush(QColor("#A8A8A8"))
        p.drawRect(1, 6, 14, 6)
        p.setBrush(PAPER)
        p.drawRect(4, 10, 8, 5)
        p.fillRect(QRect(12, 7, 2, 1), GREEN)
    elif n == "Cut":
        _pen(p, INK, 1)
        p.drawLine(5, 1, 10, 9)
        p.drawLine(11, 1, 6, 9)
        p.setBrush(Qt.BrushStyle.NoBrush)
        _pen(p, RED, 1.5)
        p.drawEllipse(QRectF(2, 9, 5, 5))
        p.drawEllipse(QRectF(9, 9, 5, 5))
    elif n == "Copy":
        _page(p, 1, 1, 8, 10, False)
        _page(p, 6, 5, 8, 10, False)
        _lines(p, (8, 10, 12), 7, 13)
    elif n == "Paste":
        _pen(p, QColor("#6A4A10"))
        p.setBrush(QColor("#C89A50"))
        p.drawRect(2, 2, 10, 13)
        p.fillRect(QRect(5, 1, 4, 3), GRAY)
        _page(p, 7, 6, 8, 9, False)
    elif n == "Bold":
        _text(p, "B", (0, 0, 16, 16), bold=True, size=13)
    elif n == "Italic":
        _text(p, "I", (0, 0, 16, 16), italic=True, bold=True, size=13)
    elif n == "Underline":
        _text(p, "U", (0, -1, 16, 15), size=12)
        _pen(p, INK)
        p.drawLine(4, 14, 12, 14)
    elif n == "Double Underline":
        _text(p, "D", (0, -2, 16, 14), size=11)
        _pen(p, INK)
        p.drawLine(4, 12, 12, 12)
        p.drawLine(4, 14, 12, 14)
    elif n == "Word Underline":
        _text(p, "W", (0, -2, 16, 14), size=11)
        _pen(p, INK)
        p.drawLine(2, 13, 14, 13)
    elif n == "Strike Through":
        _text(p, "abc", (0, 0, 16, 16), size=8, strike=True)
    elif n == "Small Caps":
        _text(p, "Aʙᴄ", (0, 0, 16, 16), size=9)
    elif n == "Align Left":
        _lines(p, (3, 6, 9, 12), 2, 14, "left")
    elif n == "Center":
        _lines(p, (3, 6, 9, 12), 2, 14, "center")
    elif n == "Align Right":
        _lines(p, (3, 6, 9, 12), 2, 14, "right")
    elif n == "Justify":
        _lines(p, (3, 6, 9, 12), 2, 14, "full")
    elif n == "Arc":
        _pen(p, INK, 1.5)
        p.drawArc(QRectF(2, 3, 12, 12), 0, 180 * 16)
    elif n == "Ellipse":
        _pen(p, INK)
        p.setBrush(QColor("#9AB8E8"))
        p.drawEllipse(QRectF(1.5, 4, 13, 8))
    elif n == "Rectangle":
        _pen(p, INK)
        p.setBrush(QColor("#9AB8E8"))
        p.drawRect(2, 4, 12, 8)
    elif n == "Line":
        _pen(p, INK, 1.5)
        p.drawLine(2, 14, 14, 2)
    elif n == "Line Style":
        for y, w in ((3, 1), (7, 2), (12, 3)):
            _pen(p, INK, w)
            p.drawLine(2, y, 14, y)
    elif n == "Freeform":
        _pen(p, INK, 1.2)
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawPolyline(QPolygon([QPoint(2, 12), QPoint(5, 4), QPoint(8, 10), QPoint(11, 3), QPoint(14, 12)]))
    elif n == "Drawing":
        p.setBrush(QColor("#9AB8E8"))
        _pen(p)
        p.drawRect(1, 7, 7, 7)
        p.setBrush(YELLOW)
        p.drawEllipse(QRectF(7, 2, 8, 8))
    elif n == "Button":
        _pen(p)
        p.setBrush(QColor("#E0E0E0"))
        p.drawRect(1, 4, 14, 8)
        _pen(p, WHITE)
        p.drawLine(2, 5, 13, 5)
        _text(p, "ok", (1, 4, 14, 8), size=7)
    elif n == "Camera":
        _pen(p)
        p.setBrush(QColor("#404040"))
        p.drawRect(1, 5, 14, 9)
        p.drawRect(5, 3, 5, 2)
        p.setBrush(QColor("#9AB8E8"))
        p.drawEllipse(QRectF(5, 6.5, 6, 6))
    elif n == "Delete":
        _pen(p, RED, 2.2)
        p.drawLine(3, 3, 13, 13)
        p.drawLine(13, 3, 3, 13)
    elif n == "Find":
        _pen(p, INK, 1.5)
        p.setBrush(QColor("#CFE2FA"))
        p.drawEllipse(QRectF(2, 2, 8, 8))
        _pen(p, QColor("#6A4A10"), 2.5)
        p.drawLine(9, 9, 14, 14)
    elif n in ("Back", "Forward"):
        _arrow(p, n == "Back")
    elif n == "Up One Level":
        _pen(p, QColor("#8A6A10"))
        p.setBrush(YELLOW)
        p.drawRect(1, 5, 14, 9)
        _pen(p, INK, 1.5)
        p.drawLine(8, 12, 8, 7)
        p.drawLine(8, 7, 5, 10)
        p.drawLine(8, 7, 11, 10)
    elif n == "Help":
        _pen(p, BLUE)
        p.setBrush(BLUE)
        p.drawEllipse(QRectF(1, 1, 14, 14))
        _text(p, "?", (1, 1, 14, 14), bold=True, size=11, color=WHITE)
    elif n == "Help What's This":
        _pen(p, INK)
        p.setBrush(WHITE)
        p.drawPolygon(QPolygon([QPoint(1, 1), QPoint(1, 12), QPoint(4, 9), QPoint(6, 14), QPoint(8, 13), QPoint(6, 8),
                                QPoint(10, 8)]))
        _text(p, "?", (8, 1, 8, 12), bold=True, size=10, color=BLUE)
    elif n == "Macro":
        _page(p, 2, 1, 11, 14, False)
        _text(p, "▶", (2, 1, 11, 14), size=8, color=GREEN)
    elif n == "Map Network Drive":
        _drive(p, True)
    elif n == "Disconnect Net Drive":
        _drive(p, False)
    elif n == "Properties":
        _page(p, 1, 1, 11, 13, False)
        _lines(p, (4, 7, 10), 3, 10)
        _pen(p, INK)
        p.setBrush(QColor("#9AB8E8"))
        p.drawRect(9, 9, 6, 6)
    elif n == "Redo":
        _curve_arrow(p, False)
    elif n == "Undo":
        _curve_arrow(p, True)
    elif n == "Sort Ascending":
        _sort(p, True)
    elif n == "Sort Descending":
        _sort(p, False)
    elif n == "Spell Check":
        _text(p, "abc", (0, -2, 16, 10), size=7, bold=True)
        _pen(p, GREEN, 2)
        p.drawLine(3, 11, 6, 14)
        p.drawLine(6, 14, 13, 7)
    elif n == "Sum":
        _text(p, "Σ", (0, 0, 16, 16), bold=True, size=13)
    elif n == "Tab Left":
        _tab(p, "left")
    elif n == "Tab Right":
        _tab(p, "right")
    elif n == "Tab Center":
        _tab(p, "center")
    elif n == "Tab Decimal":
        _tab(p, "decimal")
    elif n == "View Large Icons":
        _view(p, "large")
    elif n == "View Small Icons":
        _view(p, "small")
    elif n == "View List":
        _view(p, "list")
    elif n == "View Details":
        _view(p, "details")
    elif n == "Stop":
        _pen(p, RED.darker(130))
        p.setBrush(RED)
        p.drawPolygon(QPolygon([QPoint(5, 1), QPoint(11, 1), QPoint(15, 5), QPoint(15, 11), QPoint(11, 15), QPoint(5, 15),
                                QPoint(1, 11), QPoint(1, 5)]))
        _pen(p, WHITE, 2)
        p.drawLine(5, 8, 11, 8)
    elif n == "Refresh":
        _pen(p, GREEN, 2)
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawArc(QRectF(3, 3, 10, 10), 30 * 16, 300 * 16)
        p.setBrush(GREEN)
        p.drawPolygon(QPolygon([QPoint(10, 1), QPoint(15, 5), QPoint(10, 7)]))
    elif n == "Home":
        _pen(p)
        p.setBrush(RED)
        p.drawPolygon(QPolygon([QPoint(8, 1), QPoint(15, 8), QPoint(1, 8)]))
        p.setBrush(PAPER)
        p.drawRect(3, 8, 10, 7)
        p.fillRect(QRect(7, 11, 3, 4), QColor("#8A6A10"))
    elif n == "Search":
        _pen(p, INK, 1.5)
        p.setBrush(QColor("#CFE2FA"))
        p.drawEllipse(QRectF(1, 1, 9, 9))
        _pen(p, INK, 2.5)
        p.drawLine(9, 9, 14, 14)
    else:                                       # a button of the user's own: a frame and its first letter
        _pen(p)
        p.setBrush(QColor("#E0E0E0"))
        p.drawRect(1, 1, 14, 14)
        _text(p, (name or "?")[:1], (1, 1, 14, 14), bold=True, size=9)


# the pictures of the standard forms (appwizgen): the logo of frmSplash, the icon of frmAbout, the bulb of frmTip;
# drawn at their own size, on white (Logo, Tip) or the face colour (AppIcon)
PICTURES = {"Logo": QColor("#F0F0F0"), "AppIcon": QColor("#F0F0F0"), "Tip": WHITE}


def _draw_picture(p, name, size):
    s = size / 32.0
    p.scale(s, s)
    if name == "Logo":                          # a window on a blue tile, with a gold star
        p.setPen(QPen(QColor("#1D4E9E"), 1.2))
        p.setBrush(QColor("#3A78D8"))
        p.drawRoundedRect(QRectF(1, 1, 30, 30), 5, 5)
        p.setPen(QPen(INK, 0.8))
        p.setBrush(WHITE)
        p.drawRect(QRectF(6, 8, 20, 16))
        p.setBrush(QColor("#1D4E9E"))
        p.drawRect(QRectF(6, 8, 20, 4))
        p.setPen(QPen(GRAY, 0.8))
        for y in (15, 18, 21):
            p.drawLine(QPoint(9, y), QPoint(18, y))
        p.setPen(QPen(QColor("#8A6A10"), 0.6))
        p.setBrush(YELLOW)
        star = [QPoint(23, 13), QPoint(24, 16), QPoint(27, 16), QPoint(25, 18), QPoint(26, 21), QPoint(23, 19),
                QPoint(20, 21), QPoint(21, 18), QPoint(19, 16), QPoint(22, 16)]
        p.drawPolygon(QPolygon(star))
    elif name == "AppIcon":                     # a window of a program
        p.setPen(QPen(INK, 1))
        p.setBrush(WHITE)
        p.drawRect(QRectF(2, 4, 27, 23))
        p.setBrush(BLUE)
        p.drawRect(QRectF(2, 4, 27, 5))
        p.setBrush(RED)
        p.drawRect(QRectF(24, 5, 3, 3))
        p.setPen(QPen(GRAY, 1))
        for y in (13, 17, 21):
            p.drawLine(QPoint(6, y), QPoint(24 if y != 17 else 19, y))
    elif name == "Tip":                         # a bulb
        p.setPen(QPen(QColor("#8A6A10"), 1))
        p.setBrush(YELLOW)
        p.drawEllipse(QRectF(8, 3, 16, 16))
        p.setBrush(QColor("#B0B0B0"))
        p.setPen(QPen(GRAY, 1))
        p.drawRect(QRectF(12, 19, 8, 7))
        for y in (21, 23):
            p.drawLine(QPoint(12, y), QPoint(20, y))
        p.setPen(QPen(QColor("#E0A000"), 1.2))
        for a, b in (((16, 0), (16, 1)), ((4, 11), (6, 11)), ((26, 11), (28, 11)), ((7, 4), (8, 5)),
                     ((25, 4), (24, 5))):
            p.drawLine(QPoint(*a), QPoint(*b))


def image(name, size=16):
    img = QImage(size, size, QImage.Format.Format_RGB32)
    img.fill(PICTURES.get(name, MASK))
    p = QPainter(img)
    if name in PICTURES:
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        _draw_picture(p, name, size)
        p.end()
        return img
    p.setRenderHint(QPainter.RenderHint.Antialiasing, False)
    p.setRenderHint(QPainter.RenderHint.TextAntialiasing, False)
    if size != 16:
        p.scale(size / 16.0, size / 16.0)
    draw(p, name)
    p.end()
    _halo(img)
    return img


HALO = QColor(0xE8, 0xE8, 0xE8)


def _halo(img):
    """a light edge round the dark strokes (B, I, U, the lines of Align ...): the picture keeps its look on
    a light toolbar and can still be seen on a dark one (the toolbar of a program in a dark theme)"""
    w, h = img.width(), img.height()
    mask = MASK.rgb() & 0xFFFFFF

    def dark(x, y):
        c = QColor(img.pixel(x, y))
        return (c.rgb() & 0xFFFFFF) != mask and c.red() * 299 + c.green() * 587 + c.blue() * 114 < 125000
    edge = []
    for y in range(h):
        for x in range(w):
            if (img.pixel(x, y) & 0xFFFFFF) != mask:
                continue
            if any(0 <= x + dx < w and 0 <= y + dy < h and dark(x + dx, y + dy)
                   for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1))):
                edge.append((x, y))
    for x, y in edge:
        img.setPixel(x, y, HALO.rgb())


def bmp(name, size=16):
    """the picture as the bytes of a .bmp (what the .frx of an ImageList keeps)"""
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QIODevice.OpenModeFlag.WriteOnly)
    image(name, size).convertToFormat(QImage.Format.Format_RGB888).save(buf, "BMP")
    buf.close()
    return bytes(ba)


def icon(name, size=16):
    """for the lists of the Customize Toolbar page"""
    from PyQt6.QtGui import QPixmap, QIcon
    img = image(name, size).convertToFormat(QImage.Format.Format_ARGB32)
    m = MASK.rgb()
    for y in range(img.height()):
        for x in range(img.width()):
            if img.pixel(x, y) == m:
                img.setPixel(x, y, 0)
    return QIcon(QPixmap.fromImage(img))


__all__ = ["TOOLBAR_MDI", "TOOLBAR_EXPLORER", "ALL", "BROWSER", "PICTURES", "draw", "image", "bmp", "icon"]
