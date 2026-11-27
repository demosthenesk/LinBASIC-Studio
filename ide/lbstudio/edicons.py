"""The pictures of the Edit toolbar (phase 8.3) and of the bookmarks in the code window: drawn here, in the spirit
of VB6's (nothing of VB6 is copied)."""
from PyQt6.QtCore import Qt, QPoint, QRectF
from PyQt6.QtGui import QColor, QFont, QIcon, QPainter, QPen, QPixmap, QPolygon

BOOKMARK = QColor(0, 200, 220)                  # the cyan of VB6's bookmarks
BOOKMARK_EDGE = QColor(0, 110, 130)


def bookmark_pixmap(size=14):
    """the bookmark of the Margin Indicator Bar: a cyan rounded rectangle"""
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setPen(QPen(BOOKMARK_EDGE, 1))
    p.setBrush(BOOKMARK)
    p.drawRoundedRect(QRectF(1.5, 3.5, size - 3, size - 7), 3, 3)
    p.end()
    return pm


def _lines(p, x0=2, n=4, short=()):
    p.setPen(QPen(QColor(40, 40, 40), 1))
    for i in range(n):
        y = 3 + i * 3
        p.drawLine(x0, y, 14 - (4 if i in short else 0), y)


PE_ICONS = ("view_code", "view_object", "toggle_folders", "properties", "print", "project", "folder", "folder_open",
            "form", "mdiform", "module", "class", "usercontrol", "proppage", "userdoc", "designer", "reldoc")


def _window(p, x, y, w, h, title=QColor(0, 0, 128)):
    """a small window: white, a frame and a title bar"""
    p.setPen(QPen(QColor(40, 40, 40), 1))
    p.setBrush(QColor(255, 255, 255))
    p.drawRect(QRectF(x + 0.5, y + 0.5, w, h))
    p.fillRect(QRectF(x + 1, y + 1, w - 1, 2), title)


def _pe_icon(p, name):
    """the pictures of the Project Explorer (phase 8.7D, DEV-8/me5): its buttons, its folders and modules, and
    the items of its menus"""
    ink = QColor(40, 40, 40)
    if name == "view_code":                     # a page of code
        _window(p, 1, 1, 13, 13)
        p.setPen(QPen(QColor(0, 0, 160), 1))
        for i, w in enumerate((8, 6, 9, 5)):
            p.drawLine(3, 6 + i * 2, 3 + w, 6 + i * 2)
    elif name == "view_object":                 # a form with two controls on it
        _window(p, 1, 2, 13, 11)
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor(212, 208, 200))
        p.drawRect(QRectF(3.5, 6.5, 5, 2))
        p.drawRect(QRectF(3.5, 10.5, 9, 1))
    elif name in ("toggle_folders", "folder", "folder_open"):
        p.setPen(QPen(QColor(150, 120, 30), 1))
        p.setBrush(QColor(250, 220, 110))
        p.drawRect(QRectF(1.5, 3.5, 5, 2))
        if name == "folder_open":
            p.drawRect(QRectF(1.5, 5.5, 12, 8))
            p.setBrush(QColor(255, 240, 160))
            p.drawPolygon(QPolygon([QPoint(3, 8), QPoint(15, 8), QPoint(13, 14), QPoint(1, 14)]))
        else:
            p.drawRect(QRectF(1.5, 5.5, 13, 8))
    elif name == "properties":                  # a card of properties and a hand
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(1.5, 4.5, 11, 10))
        p.fillRect(QRectF(2, 5, 10, 2), QColor(0, 0, 128))
        for i in range(3):
            p.drawLine(3, 9 + i * 2, 10, 9 + i * 2)
        p.setBrush(QColor(120, 120, 120))
        p.drawRect(QRectF(9.5, 1.5, 5, 3))
    elif name == "print":                       # a printer
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(4.5, 1.5, 7, 5))
        p.setBrush(QColor(190, 190, 190))
        p.drawRect(QRectF(1.5, 6.5, 13, 6))
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(4.5, 10.5, 7, 4))
        p.fillRect(QRectF(11, 8, 2, 1), QColor(220, 180, 0))
    elif name == "project":                     # a form and a module, one over the other
        _window(p, 0, 4, 10, 9)
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor(0, 128, 128))
        p.drawEllipse(QRectF(8.5, 1.5, 6, 6))
        p.setBrush(QColor(200, 0, 0))
        p.drawEllipse(QRectF(11, 9, 4, 4))
    elif name in ("form", "mdiform", "usercontrol", "proppage", "userdoc", "designer", "reldoc"):
        if name == "mdiform":
            _window(p, 3, 1, 12, 9)
        _window(p, 1, 3 if name == "mdiform" else 2, 12, 11)
        p.setPen(QPen(ink, 1))
        if name == "usercontrol":               # a control on a form
            p.setBrush(QColor(212, 208, 200))
            p.drawRect(QRectF(3.5, 7.5, 7, 4))
        elif name == "proppage":                # a tab
            p.drawLine(3, 7, 7, 7)
            p.drawLine(7, 7, 7, 9)
            p.drawLine(3, 9, 11, 9)
        elif name == "userdoc":                 # a page with a folded corner
            p.drawLine(9, 7, 11, 9)
        elif name in ("designer", "reldoc"):    # lines
            for i in range(3):
                p.drawLine(3, 7 + i * 2, 10, 7 + i * 2)
    elif name == "module":                      # three coloured boxes joined
        p.setPen(QPen(ink, 1))
        for (x, y), c in zip(((1, 7), (6, 2), (10, 9)), (QColor(200, 0, 160), QColor(0, 140, 60), QColor(0, 0, 200))):
            p.setBrush(c)
            p.drawRect(QRectF(x + 0.5, y + 0.5, 4, 4))
        p.drawLine(5, 9, 8, 6)
        p.drawLine(9, 7, 11, 10)
    elif name == "class":                       # a box with a gold edge
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor(230, 220, 250))
        p.drawRect(QRectF(2.5, 3.5, 11, 10))
        p.setBrush(QColor(220, 180, 0))
        p.drawPolygon(QPolygon([QPoint(2, 3), QPoint(6, 1), QPoint(15, 1), QPoint(13, 3)]))
        p.setPen(QPen(QColor(120, 0, 120), 1.4))
        p.drawLine(5, 8, 11, 8)


def icon(name):
    pm = QPixmap(16, 16)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    f = QFont("Sans")
    f.setPixelSize(10)
    f.setBold(True)
    p.setFont(f)
    if name in ("comment", "uncomment"):
        _lines(p, 6, 4, (1, 3))
        p.setPen(QPen(QColor(0, 128, 0), 2))
        for i in range(4):
            p.drawLine(2, 2 + i * 3, 2, 3 + i * 3)          # the ' of every line
        if name == "uncomment":
            p.setPen(QPen(QColor(200, 30, 30), 1.6))
            p.drawLine(0, 13, 6, 1)
    elif name in ("indent", "outdent"):
        _lines(p, 7, 4)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(0, 80, 160))
        pts = [QPoint(1, 4), QPoint(5, 7), QPoint(1, 10)] if name == "indent" else \
              [QPoint(5, 4), QPoint(1, 7), QPoint(5, 10)]
        p.drawPolygon(QPolygon(pts))
    elif name.startswith("bm_"):
        p.setPen(QPen(BOOKMARK_EDGE, 1))
        p.setBrush(BOOKMARK)
        p.drawRoundedRect(QRectF(1.5, 4.5, 9, 7), 2, 2)
        p.setPen(Qt.PenStyle.NoPen)
        if name == "bm_next" or name == "bm_prev":
            p.setBrush(QColor(0, 80, 160))
            pts = [QPoint(11, 4), QPoint(16, 4), QPoint(13, 9)] if name == "bm_next" else \
                  [QPoint(11, 12), QPoint(16, 12), QPoint(13, 7)]
            p.drawPolygon(QPolygon(pts))
        elif name == "bm_clear":
            p.setPen(QPen(QColor(200, 30, 30), 1.8))
            p.drawLine(9, 9, 15, 15)
            p.drawLine(15, 9, 9, 15)
    elif name == "lock":                        # Format > Lock Controls: a padlock (navy, as VB6's)
        p.setPen(QPen(QColor(0, 0, 128), 1.6))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawArc(QRectF(4.5, 1.5, 7, 8), 0, 180 * 16)
        p.drawLine(5, 6, 5, 8)
        p.drawLine(11, 6, 11, 8)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(0, 0, 128))
        p.drawRect(QRectF(3, 7.5, 10, 7))
    elif name == "breakpoint":
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(128, 0, 0))
        p.drawEllipse(QRectF(3, 3, 10, 10))
    elif name in ("step_into", "step_over", "step_out"):     # the Debug toolbar: lines and a blue arrow
        _lines(p, 8, 4, (1,))
        p.setPen(QPen(QColor(0, 0, 160), 1.6))
        if name == "step_into":
            p.drawLine(2, 2, 6, 2)
            p.drawLine(2, 2, 2, 8)
            p.drawLine(2, 8, 6, 8)
            tip = [QPoint(6, 5), QPoint(9, 8), QPoint(6, 11)]
        elif name == "step_over":
            p.drawLine(5, 2, 2, 2)
            p.drawLine(2, 2, 2, 13)
            p.drawLine(2, 13, 5, 13)
            tip = [QPoint(5, 10), QPoint(8, 13), QPoint(5, 16)]
        else:
            p.drawLine(4, 13, 4, 5)
            tip = [QPoint(1, 6), QPoint(4, 2), QPoint(7, 6)]
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(0, 0, 160))
        p.drawPolygon(QPolygon(tip))
    elif name in ("locals", "immediate", "watch", "quick_watch", "call_stack"):
        if name == "call_stack":                # three frames, one over the other
            p.setPen(QPen(QColor(60, 60, 60), 1))
            p.setBrush(QColor(255, 255, 255))
            for i in range(3):
                p.drawRect(QRectF(1.5 + i * 4, 1.5 + i * 4, 7, 5))
        elif name == "quick_watch":             # the glasses
            p.setPen(QPen(QColor(0, 100, 120), 1.5))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawEllipse(QRectF(1.5, 7.5, 5, 5))
            p.drawEllipse(QRectF(8.5, 7.5, 5, 5))
            p.drawLine(6, 9, 9, 9)
            p.drawLine(13, 8, 15, 3)
        else:                                   # a window with a title bar
            p.setPen(QPen(QColor(40, 40, 40), 1))
            p.setBrush(QColor(255, 255, 255))
            p.drawRect(QRectF(1.5, 2.5, 13, 11))
            p.fillRect(QRectF(2, 3, 12, 2), QColor(0, 0, 128))
            _lines(p, 4, 0)
            p.setPen(QPen(QColor(40, 40, 40), 1))
            for i in range(3):
                p.drawLine(4, 7 + i * 2, 11 - (3 if i == 1 else 0), 7 + i * 2)
            if name == "locals":
                p.setPen(QPen(QColor(200, 160, 0), 1.5))
                p.drawEllipse(QRectF(0.5, 8.5, 3, 6))
            elif name == "watch":
                p.setPen(QPen(QColor(0, 100, 120), 1.4))
                p.drawEllipse(QRectF(2.5, 10.5, 4, 4))
                p.drawEllipse(QRectF(8.5, 10.5, 4, 4))
    elif name in ("front", "back"):             # Bring to Front / Send to Back: two squares
        dark, light = QColor(70, 70, 70), QColor(170, 170, 170)
        first, second = (light, dark) if name == "front" else (dark, light)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(first)
        p.drawRect(QRectF(1, 1, 9, 9))
        p.setBrush(second)
        p.drawRect(QRectF(6, 6, 9, 9))
    elif name in ("align", "center", "same_size"):
        p.setPen(QPen(QColor(30, 30, 30), 1))
        p.setBrush(QColor(255, 255, 255))
        if name == "align":                     # Align Lefts: a line on the left, two boxes
            p.drawLine(2, 1, 2, 15)
            p.drawRect(QRectF(3.5, 2.5, 8, 4))
            p.drawRect(QRectF(3.5, 9.5, 11, 4))
        elif name == "center":                  # Center in Form: a box on a middle line
            p.drawRect(QRectF(1.5, 1.5, 13, 13))
            p.drawRect(QRectF(4.5, 5.5, 7, 5))
            p.drawLine(8, 2, 8, 5)
            p.drawLine(8, 11, 8, 14)
        else:                                   # Make Same Size: two boxes of one width
            p.drawRect(QRectF(1.5, 1.5, 13, 5))
            p.setPen(QPen(QColor(30, 30, 30), 1, Qt.PenStyle.DotLine))
            p.drawRect(QRectF(1.5, 9.5, 13, 5))
    elif name in ("position", "size"):          # the Standard toolbar: where the control is, how big
        p.setPen(QPen(QColor(40, 40, 40), 1))
        p.setBrush(QColor(255, 255, 255))
        if name == "position":                  # a box and its top-left corner
            p.setPen(QPen(QColor(40, 40, 40), 1, Qt.PenStyle.DotLine))
            p.drawRect(QRectF(4.5, 4.5, 10, 10))
            p.setPen(QPen(QColor(0, 0, 0), 1.5))
            p.drawLine(1, 4, 7, 4)
            p.drawLine(4, 1, 4, 7)
        else:                                   # a box and the arrows of its width and height
            p.setPen(QPen(QColor(40, 40, 40), 1, Qt.PenStyle.DotLine))
            p.drawRect(QRectF(1.5, 1.5, 10, 10))
            p.setPen(QPen(QColor(0, 0, 0), 1.5))
            p.drawLine(6, 14, 15, 14)
            p.drawLine(14, 6, 14, 15)
    elif name == "cut":                         # the menus of the designer / Watches (DEV-8/me8): scissors
        p.setPen(QPen(QColor(0, 0, 128), 1.3))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawLine(5, 1, 10, 10)
        p.drawLine(11, 1, 6, 10)
        p.drawEllipse(QRectF(2.5, 10.5, 4, 4))
        p.drawEllipse(QRectF(9.5, 10.5, 4, 4))
    elif name == "copy":                        # two pages, one over the other
        p.setPen(QPen(QColor(0, 0, 128), 1))
        for x, y in ((1.5, 1.5), (6.5, 5.5)):
            p.setBrush(QColor(255, 255, 255))
            p.drawRect(QRectF(x, y, 8, 9))
            for i in range(3):
                p.drawLine(int(x) + 2, int(y) + 3 + i * 2, int(x) + 6, int(y) + 3 + i * 2)
    elif name == "paste":                       # a clipboard and a page
        p.setPen(QPen(QColor(90, 60, 20), 1))
        p.setBrush(QColor(200, 160, 90))
        p.drawRect(QRectF(1.5, 2.5, 10, 12))
        p.setBrush(QColor(160, 160, 160))
        p.drawRect(QRectF(4.5, 1.5, 4, 3))
        p.setPen(QPen(QColor(0, 0, 128), 1))
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(7.5, 6.5, 7, 8))
    elif name == "align_grid":                  # Align to Grid: a box on the crossing lines of the grid
        p.setPen(QPen(QColor(30, 30, 30), 1))
        p.drawLine(4, 1, 4, 15)
        p.drawLine(1, 4, 15, 4)
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(4.5, 4.5, 8, 8))
        p.drawLine(12, 1, 12, 15)
        p.drawLine(1, 12, 15, 12)
    elif name == "menu_editor":                 # a menu bar and its drop-down
        p.setPen(QPen(QColor(30, 30, 30), 1))
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(1.5, 1.5, 13, 3))
        p.drawRect(QRectF(5.5, 4.5, 8, 10))
        for i in range(3):
            p.drawLine(7, 7 + i * 3, 12, 7 + i * 3)
    elif name == "references":                  # Object Browser > References...: two books
        p.setPen(QPen(QColor(40, 40, 40), 1))
        p.setBrush(QColor(120, 120, 200))
        p.drawRect(QRectF(2.5, 3.5, 5, 11))
        p.setBrush(QColor(200, 120, 120))
        p.drawRect(QRectF(8.5, 1.5, 5, 13))
    elif name == "help":                        # a yellow question mark
        f = QFont("Sans")
        f.setPixelSize(14)
        f.setBold(True)
        p.setFont(f)
        p.setPen(QPen(QColor(200, 150, 0), 1))
        p.drawText(QRectF(0, 0, 16, 16), int(Qt.AlignmentFlag.AlignCenter), "?")
    elif name == "form_layout":                 # the Standard toolbar of VB6: a screen with a small window
        p.setPen(QPen(QColor(0, 0, 0), 1))
        p.setBrush(QColor(0, 0, 0))
        p.drawRect(QRectF(1.5, 2.5, 13, 10))
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(4.5, 5.5, 6, 4))
        p.drawLine(5, 14, 11, 14)
    elif name == "toolbox":                     # a hammer and a wrench
        p.setPen(QPen(QColor(90, 60, 20), 2))
        p.drawLine(4, 13, 11, 4)
        p.setPen(QPen(QColor(60, 60, 60), 2))
        p.drawLine(9, 2, 13, 6)
        p.drawLine(4, 4, 12, 13)
    elif name == "data_view":                   # a yellow cylinder (a database)
        p.setPen(QPen(QColor(0, 0, 0), 1))
        p.setBrush(QColor(255, 230, 0))
        p.drawRect(QRectF(2.5, 4.5, 11, 9))
        p.drawEllipse(QRectF(2.5, 2.5, 11, 4))
        p.drawEllipse(QRectF(2.5, 11.5, 11, 4))
    elif name == "component_manager":           # boxes joined
        p.setPen(QPen(QColor(0, 0, 0), 1))
        for (x, y), c in zip(((2, 9), (9, 9), (6, 2)), (QColor(0, 150, 0), QColor(0, 0, 200), QColor(200, 0, 0))):
            p.setBrush(c)
            p.drawEllipse(QRectF(x + 0.5, y + 0.5, 5, 5))
        p.drawLine(5, 10, 8, 6)
        p.drawLine(11, 10, 9, 6)
    elif name == "add_project":                 # a page with a star (Add Standard EXE Project)
        p.setPen(QPen(QColor(0, 0, 0), 1))
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(QRectF(3.5, 4.5, 10, 10))
        p.setPen(QPen(QColor(220, 160, 0), 2))
        p.drawLine(3, 1, 3, 5)
        p.drawLine(1, 3, 5, 3)
    elif name in PE_ICONS:
        _pe_icon(p, name)
    else:                                       # the IntelliSense commands: a letter
        text = {"list_members": "M", "list_constants": "C", "quick_info": "i", "param_info": "P",
                "complete_word": "W"}.get(name, "?")
        p.setPen(QPen(QColor(0, 80, 160), 1))
        p.setBrush(QColor(230, 240, 255))
        p.drawRoundedRect(QRectF(1.5, 1.5, 13, 13), 3, 3)
        p.drawText(QRectF(1, 1, 14, 14), int(Qt.AlignmentFlag.AlignCenter), text)
    p.end()
    return QIcon(pm)
