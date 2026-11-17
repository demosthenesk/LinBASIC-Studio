"""View > Color Palette: the colours of VB6's palette window (DEV-8/me2 pictures 7, 8).

On the left the colours of the selected control: the outer square is its BackColor, the inner one its
ForeColor; a click on one of them chooses which of the two a colour of the palette changes, and "Aa" shows
the text in the two colours. Then 48 colours (3 rows of 16) and a row of 16 colours of your own (white until
Define Colors... gives them one). Default puts back the colour the control has when it is made.
"""

from PyQt6.QtCore import Qt, QRect, pyqtSignal
from PyQt6.QtGui import QColor, QPainter, QPen, QFont
from PyQt6.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QPushButton, QColorDialog

from . import vbprops as V

# the colours of VB6's palette, row by row (read from the window of VB6, DEV-8/me2 picture 7)
PALETTE = [
    0xFFFFFF, 0xE0E0E0, 0xFFC0C0, 0xFFE0C0, 0xFFFFC0, 0xC0FFC0, 0xC0FFFF, 0xC0C0FF,
    0xFFC0FF, 0xC00000, 0xC04000, 0xC0C000, 0x00C000, 0x00C0C0, 0x0000C0, 0xC000C0,
    0xC0C0C0, 0x404040, 0xFF8080, 0xFFC080, 0xFFFF80, 0x80FF80, 0x80FFFF, 0x8080FF,
    0xFF80FF, 0x800000, 0x804000, 0x808000, 0x008000, 0x008080, 0x000080, 0x800080,
    0x808080, 0x000000, 0xFF0000, 0xFF8000, 0xFFFF00, 0x00FF00, 0x00FFFF, 0x0000FF,
    0xFF00FF, 0x400000, 0x804040, 0x404000, 0x004000, 0x004040, 0x000040, 0x400040,
]
COLS, ROWS = 16, 4
CELL = 17


def rgb_to_vb(rgb):
    """0xRRGGBB (as on the screen) -> the OLE_COLOR of VB6 (&H00BBGGRR&)"""
    return ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)


def vb_to_qcolor(v):
    """an OLE_COLOR (a system colour too, &H80000005&) -> the colour that is seen"""
    from .propwin import qcolor
    return qcolor(v)


class Sample(QWidget):
    """the two squares (BackColor around ForeColor) and "Aa" under them"""
    picked = pyqtSignal(str)                   # "BackColor" / "ForeColor"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(40, 60)
        self.back = QColor("#F0F0F0")
        self.fore = QColor("#000000")
        self.target = "BackColor"

    def outer(self):
        return QRect(1, 1, 37, 37)

    def inner(self):
        return QRect(12, 12, 15, 15)

    def mousePressEvent(self, e):
        p = e.position().toPoint()
        if self.inner().contains(p):
            self.target = "ForeColor"
        elif self.outer().contains(p) or p.y() > 40:
            self.target = "BackColor"
        self.update()
        self.picked.emit(self.target)

    def paintEvent(self, e):
        p = QPainter(self)
        o, i = self.outer(), self.inner()
        p.fillRect(o, self.back)
        p.fillRect(i, self.fore)
        # the part that a colour of the palette changes has a white and a black frame, as in VB6
        sel = o if self.target == "BackColor" else i.adjusted(-2, -2, 2, 2)
        p.setPen(QPen(QColor("#000000")))
        p.drawRect(sel.adjusted(0, 0, -1, -1))
        p.setPen(QPen(QColor("#FFFFFF")))
        p.drawRect(sel.adjusted(1, 1, -2, -2))
        r = QRect(1, 43, 37, 15)
        p.fillRect(r, self.back)
        p.setPen(QColor("#808080"))
        p.drawRect(r.adjusted(0, 0, -1, -1))
        f = QFont(self.font())
        f.setBold(True)
        p.setFont(f)
        p.setPen(self.fore)
        p.drawText(r, Qt.AlignmentFlag.AlignCenter, "Aa")
        p.end()


class Grid(QWidget):
    """the 64 boxes"""
    clicked = pyqtSignal(int)

    def __init__(self, palette, parent=None):
        super().__init__(parent)
        self.colors = palette                  # 64 colours 0xRRGGBB
        self.current = -1
        self.setFixedSize(COLS * CELL + 3, ROWS * CELL + 3)

    def cell_rect(self, i):
        r, c = divmod(i, COLS)
        return QRect(2 + c * CELL, 2 + r * CELL, CELL - 2, CELL - 2)

    def index_at(self, pos):
        c, r = (pos.x() - 1) // CELL, (pos.y() - 1) // CELL
        return r * COLS + c if 0 <= c < COLS and 0 <= r < ROWS else -1

    def mousePressEvent(self, e):
        i = self.index_at(e.position().toPoint())
        if i >= 0:
            self.current = i
            self.update()
            self.clicked.emit(i)

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#808080"))
        for i, rgb in enumerate(self.colors):
            p.fillRect(self.cell_rect(i), QColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF))
        if self.current >= 0:
            r = self.cell_rect(self.current).adjusted(-1, -1, 0, 0)
            p.setPen(QColor("#FFFFFF"))
            p.drawRect(r)
            p.setPen(QColor("#000000"))
            p.drawRect(r.adjusted(-1, -1, 1, 1))
        p.end()


class ColorPalette(QWidget):
    """the window; main gives it the designer and the control (show_for) and makes the changes"""

    def __init__(self, main):
        super().__init__()
        self.main = main
        custom = list(main.S["custom_colors"] or [])
        custom = (custom + [0xFFFFFF] * COLS)[:COLS]
        lay = QHBoxLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(8)
        self.sample = Sample()
        lay.addWidget(self.sample, 0, Qt.AlignmentFlag.AlignTop)
        self.grid = Grid(PALETTE + custom)
        lay.addWidget(self.grid, 0, Qt.AlignmentFlag.AlignTop)
        right = QVBoxLayout()
        self.b_default = QPushButton("Default")
        self.b_define = QPushButton("Define Colors...")
        self.b_define.setEnabled(False)        # VB6: for a box of the last row (the colours of your own)
        for b in (self.b_default, self.b_define):
            b.setMinimumWidth(105)
            right.addWidget(b)
        right.addStretch(1)
        lay.addLayout(right)
        self.sample.picked.connect(lambda t: self.refresh())
        self.grid.clicked.connect(self.cell_clicked)
        self.b_default.clicked.connect(self.set_default)
        self.b_define.clicked.connect(self.define_color)
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.refresh()

    # ---------------------------------------------------------------- what is shown
    @property
    def designer(self):
        """the designer in use (the palette changes nothing with a code window in front)"""
        return self.main.current_designer()

    @property
    def ctl(self):
        d = self.designer
        if d is None:
            return None
        return d.primary.ctl if d.primary else d.root

    def show_for(self, designer=None, ctl=None):
        """the selection or the window in front has changed"""
        if self.isVisible():
            self.refresh()

    def _targets(self):
        """the controls a colour goes to: the selected ones (the form when none is)"""
        d = self.designer
        if d is None:
            return []
        ctls = [it.ctl for it in d.selection] if d.selection else [d.root]
        return [c for c in ctls if V.prop_def(c.type_name, self.sample.target) is not None]

    def _value(self, ctl, name):
        p = V.prop_def(ctl.type_name, name)
        return None if p is None else V.value_of(ctl, p)

    def refresh(self):
        c = self.ctl
        for name, attr in (("BackColor", "back"), ("ForeColor", "fore")):
            v = self._value(c, name) if c is not None else None
            setattr(self.sample, attr, vb_to_qcolor(v) if v is not None else
                    QColor("#F0F0F0" if name == "BackColor" else "#000000"))
        v = self._value(c, self.sample.target) if c is not None else None
        cur, self.grid.current = self.grid.current, -1
        if v is not None and not int(v) & 0x80000000:
            rgb = rgb_to_vb(int(v))            # the same swap turns BGR back into RGB
            if 0 <= cur < len(self.grid.colors) and self.grid.colors[cur] == rgb:
                self.grid.current = cur        # the box clicked (white is in the first and in the last row)
            elif rgb in self.grid.colors:
                self.grid.current = self.grid.colors.index(rgb)
        self.b_define.setEnabled(self.grid.current >= (ROWS - 1) * COLS)
        self.sample.update()
        self.grid.update()

    # ---------------------------------------------------------------- the changes
    def _apply(self, value_of_ctl):
        d = self.designer
        ctls = self._targets()
        if d is None or not ctls:
            return
        name = self.sample.target
        d.checkpoint()
        for c in ctls:
            V.apply(c, V.prop_def(c.type_name, name), value_of_ctl(c))
        keep = [c for c in ([it.ctl for it in d.selection] if d.selection else [])]
        primary = d.primary.ctl if d.primary else None
        d.build()
        items = [d.items[c] for c in keep if c in d.items]
        d.select(items, d.items.get(primary))
        d.changed()

    def cell_clicked(self, i):
        rgb = self.grid.colors[i]
        self.grid.current = i                  # the box clicked stays marked (white is in two rows)
        self.b_define.setEnabled(i >= (ROWS - 1) * COLS)
        self._apply(lambda c: rgb_to_vb(rgb))

    def set_default(self):
        name = self.sample.target
        self._apply(lambda c: V.prop_def(c.type_name, name).default)

    def define_color(self):
        i = self.grid.current
        if i < (ROWS - 1) * COLS:
            return
        rgb = self.grid.colors[i]
        col = QColorDialog.getColor(QColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF), self, "Color")
        if not col.isValid():
            return
        rgb = (col.red() << 16) | (col.green() << 8) | col.blue()
        self.grid.colors[i] = rgb
        self.main.S["custom_colors"] = self.grid.colors[(ROWS - 1) * COLS:]
        self.main.S.save()
        self._apply(lambda c: rgb_to_vb(rgb))
