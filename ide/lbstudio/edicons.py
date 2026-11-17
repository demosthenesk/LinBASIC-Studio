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
    else:                                       # the IntelliSense commands: a letter
        text = {"list_members": "M", "list_constants": "C", "quick_info": "i", "param_info": "P",
                "complete_word": "W"}.get(name, "?")
        p.setPen(QPen(QColor(0, 80, 160), 1))
        p.setBrush(QColor(230, 240, 255))
        p.drawRoundedRect(QRectF(1.5, 1.5, 13, 13), 3, 3)
        p.drawText(QRectF(1, 1, 14, 14), int(Qt.AlignmentFlag.AlignCenter), text)
    p.end()
    return QIcon(pm)
