"""View > Form Layout Window (phase 8.4), as the VB6 IDE has it (VB6 oracle DEV-8/fd1 stages 17-19): the screen
with the forms whose designer is open, each where it will start.

  StartUpPosition 0 Manual (not written): at ClientLeft / ClientTop; 1 'CenterOwner and 2 'CenterScreen: in the
  middle; 3 'Windows Default: at the top left.
  Dragging a form: StartUpPosition becomes Manual (the line goes) and ClientLeft / ClientTop take the new place.
  The right button: Resolution Guides (640x480, 800x600, 1024x768, 1280x1024), Startup Position > Manual /
  Center Owner / Center Screen / Windows Default, Dockable, Hide. Resolution Guides is on at first; without
  Dockable the window goes among the code and form windows and does not dock (stage 20).
"""
from PyQt6.QtCore import Qt, QRect, QRectF, QPoint, pyqtSignal
from PyQt6.QtGui import QPainter, QColor, QPen, QFont, QGuiApplication, QAction, QActionGroup
from PyQt6.QtWidgets import QWidget, QMenu

from . import vbprops as V

TPP = 15
STARTUP = [(0, "&Manual"), (1, "Center &Owner"), (2, "Center &Screen"), (3, "&Windows Default")]
GUIDES = [(640, 480), (800, 600), (1024, 768), (1280, 1024)]


def _geti(ctl, name, default=0):
    try:
        v = ctl.get(name)
        return int(v) if v is not None else default
    except (TypeError, ValueError):
        return default


class FormLayoutView(QWidget):
    hide_requested = pyqtSignal()
    dockable_toggled = pyqtSignal(bool)

    def __init__(self, main):
        super().__init__()
        self.main = main
        S = getattr(main, "S", None)
        self.guides = bool(S["layout_guides"]) if S is not None else True
        self.drag = None                        # (designer, press point, window rect in twips at the press)
        self.setMinimumSize(160, 110)
        self.setMouseTracking(False)

    # ------------------------------------------------------------ what it shows
    def screen_twips(self):
        scr = QGuiApplication.primaryScreen()
        g = scr.geometry() if scr is not None else QRect(0, 0, 1920, 1080)
        return g.width() * TPP, g.height() * TPP

    def designers(self):
        """the forms whose designer is open (VB6: a form whose window is closed is not shown)"""
        from .designer import FormDesigner
        out = []
        for m, sw in list(getattr(self.main, "design_windows", {}).items()):
            d = sw.widget() if sw is not None else None
            if isinstance(d, FormDesigner) and m.kind == "Form" and d.root is not None:
                out.append(d)
        return out

    def window_twips(self, d):
        """the window of the form on the screen, in twips"""
        root = d.root
        b, cap, menu = d.frame_w.frame()
        cw, ch = _geti(root, "ClientWidth", 4800), _geti(root, "ClientHeight", 3600)
        w, h = cw + 2 * b * TPP, ch + (2 * b + cap + menu) * TPP
        sw, sh = self.screen_twips()
        sp = _geti(root, "StartUpPosition", 0)
        if sp in (1, 2):
            x, y = (sw - w) // 2, (sh - h) // 2
        elif sp == 3:
            x, y = 0, 0
        else:
            x = _geti(root, "ClientLeft", 120) - b * TPP
            y = _geti(root, "ClientTop", 450) - (b + cap + menu) * TPP
        return QRect(x, y, w, h)

    def _scale(self):
        sw, sh = self.screen_twips()
        r = self.rect().adjusted(8, 6, -8, -10)
        k = min(r.width() / sw, r.height() / sh) if sw and sh else 1.0
        screen = QRectF(r.x() + (r.width() - sw * k) / 2, r.y() + (r.height() - sh * k) / 2, sw * k, sh * k)
        return k, screen

    def to_widget(self, rt):
        k, screen = self._scale()
        return QRectF(screen.x() + rt.x() * k, screen.y() + rt.y() * k, rt.width() * k, rt.height() * k)

    def designer_at(self, pos):
        for d in reversed(self.designers()):
            if self.to_widget(self.window_twips(d)).contains(pos.x(), pos.y()):
                return d
        return None

    # ------------------------------------------------------------ painting
    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#F0F0F0"))
        k, screen = self._scale()
        p.fillRect(screen.adjusted(-4, -4, 4, 4), QColor("#A0A0A0"))
        p.fillRect(screen, QColor("#000000"))
        if self.guides:
            pen = QPen(QColor("#FFFFFF"))
            pen.setStyle(Qt.PenStyle.DashLine)
            f = QFont(self.font())
            f.setPixelSize(9)
            p.setFont(f)
            for gw, gh in GUIDES:
                rr = QRectF(screen.x(), screen.y(), gw * TPP * k, gh * TPP * k).intersected(screen)
                p.setPen(pen)
                p.drawLine(rr.topRight(), rr.bottomRight())
                p.drawLine(rr.bottomLeft(), rr.bottomRight())
                p.setPen(QColor("#FFFFFF"))
                p.drawText(QPoint(int(screen.x()) + 2, int(rr.bottom()) - 2), f"{gw}x{gh}")
        f = QFont(self.font())
        f.setPixelSize(9)
        p.setFont(f)
        for d in self.designers():
            r = self.to_widget(self.window_twips(d))
            p.fillRect(r, QColor("#FFFFFF"))
            cap = QRectF(r.x(), r.y(), r.width(), max(3.0, min(8.0, r.height() / 4)))
            p.fillRect(cap, QColor("#99B4D1"))
            p.setPen(QColor("#404040"))
            p.drawRect(r)
            p.setPen(QColor("#000000"))
            p.drawText(r.adjusted(2, cap.height(), -1, 0), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignBottom,
                       d.root.name)

    # ------------------------------------------------------------ dragging a form
    def mousePressEvent(self, e):
        if e.button() != Qt.MouseButton.LeftButton:
            return
        d = self.designer_at(e.position())
        if d is not None:
            self.drag = (d, e.position(), self.window_twips(d))

    def mouseMoveEvent(self, e):
        if self.drag is None:
            return
        d, start, r = self.drag
        k, _ = self._scale()
        dx, dy = (e.position().x() - start.x()) / k, (e.position().y() - start.y()) / k
        self._place(d, r.x() + dx, r.y() + dy, final=False)

    def mouseReleaseEvent(self, e):
        if self.drag is None:
            return
        d, start, r = self.drag
        self.drag = None
        k, _ = self._scale()
        dx, dy = (e.position().x() - start.x()) / k, (e.position().y() - start.y()) / k
        if abs(dx) < 2 * TPP and abs(dy) < 2 * TPP:
            return
        self._place(d, r.x() + dx, r.y() + dy, final=True)

    def _place(self, d, x, y, final):
        """the window at (x, y) twips: Manual, ClientLeft / ClientTop (DEV-8/fd1 l3, l5)"""
        root = d.root
        b, cap, menu = d.frame_w.frame()
        if final:
            d.checkpoint()
        V.apply(root, V.prop_def("Form", "StartUpPosition"), 0)
        V.put(root, "ClientLeft", int(round(x / TPP)) * TPP + b * TPP)
        V.put(root, "ClientTop", int(round(y / TPP)) * TPP + (b + cap + menu) * TPP)
        if final:
            d.changed()
        self.update()

    def set_startup(self, d, n):
        d.checkpoint()
        V.apply(d.root, V.prop_def("Form", "StartUpPosition"), n)
        d.changed()
        self.update()

    # ------------------------------------------------------------ the right button
    def context_menu(self, pos):
        """the menu of the right button (built for the tests too)"""
        m = QMenu(self)
        a = m.addAction("Resolution &Guides")
        a.setCheckable(True)
        a.setChecked(self.guides)
        a.toggled.connect(self._guides)
        d = self.designer_at(pos)
        if d is not None:
            sm = m.addMenu("Startup &Position")
            grp = QActionGroup(sm)
            cur = _geti(d.root, "StartUpPosition", 0)
            for n, text in STARTUP:
                act = QAction(text, sm)
                act.setCheckable(True)
                act.setChecked(n == cur)
                grp.addAction(act)
                act.triggered.connect(lambda checked=False, nn=n, dd=d: self.set_startup(dd, nn))
                sm.addAction(act)
        m.addSeparator()
        dock = m.addAction("Doc&kable")         # off: a window among the code and form windows that does
        dock.setCheckable(True)                 # not dock when dragged (DEV-8/fd1 stage 20, picture 31)
        S = getattr(self.main, "S", None)
        dock.setChecked(bool(S["layout_dockable"]) if S is not None else True)
        dock.toggled.connect(self.dockable_toggled.emit)
        m.addAction("&Hide", self.hide_requested.emit)
        return m

    def _guides(self, on):
        self.guides = on
        S = getattr(self.main, "S", None)
        if S is not None:
            S["layout_guides"] = on
            S.save()
        self.update()

    def contextMenuEvent(self, e):
        self.context_menu(e.pos()).exec(e.globalPos())
