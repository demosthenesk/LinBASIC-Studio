"""ocxchart.py - the property pages of an MSChart (phase 11, stage 2d): Chart, Axis, Axis Grid, Series, Series Color,
Backdrop, Text and Fonts. They change the chart in its OleObjectBlob (lbvc/project/chartblob.py: ChartEdit) where
VB6 changes it: the VB6 oracle DEV-11 s_chart2 (the chart in a PropertyBag after each box of the pages, read with
chartblob.decode; Claude Enviroment/tools/chartdiff.py). The boxes of colours and patterns of the pages (VCIColorCombo,
VCIPatternCombo) and the boxes of the widths and styles of lines draw what they hold here.
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
import struct

from PyQt6.QtCore import Qt, QSize, QRect, QPoint
from PyQt6.QtGui import QColor, QPixmap, QPainter, QPen, QBrush, QIcon
from PyQt6.QtWidgets import (QWidget, QComboBox, QToolButton, QMenu, QColorDialog, QLabel, QFrame, QCheckBox,
                             QLineEdit, QPushButton, QListWidget, QPlainTextEdit, QHBoxLayout)

from lbvc.project.chartblob import ChartEdit, ChartError

# ChartType of the list of the Chart page: (2D, 3D); Pie and X Y have no 3D (s_chart2)
TYPES = [(1, 0), (3, 2), (5, 4), (7, 6), (9, 8), (14, None), (16, None)]
TYPE_NAMES = ["Bar/Pictograph", "Line", "Area", "Step", "Combination", "Pie", "X Y (Scatter)"]
WIDTHS = [0, 1, 2, 3, 4, 5, 6, 8, 10, 12]               # points (the blob: twentieths, s_chart2: 0, 20 ... 160)
PEN_STYLES = ["None", "Solid", "Dashed", "Dotted", "Dash Dot", "Dash Dot Dot", "Ditted", "Dash Dit", "Dash Dash Dit"]
QT_PENS = [Qt.PenStyle.NoPen, Qt.PenStyle.SolidLine, Qt.PenStyle.DashLine, Qt.PenStyle.DotLine,
           Qt.PenStyle.DashDotLine, Qt.PenStyle.DashDotDotLine, Qt.PenStyle.DotLine, Qt.PenStyle.DashDotLine,
           Qt.PenStyle.DashDotDotLine]
# the patterns of a brush: (Style, Index) - VtBrushStyle Solid 1, Pattern 2 (VtBrushPatterns 0-17), Hatched 3
PATTERNS = [(1, 0)] + [(2, i) for i in range(18)] + [(3, i) for i in range(6)]
QT_PATTERNS = [Qt.BrushStyle.SolidPattern, Qt.BrushStyle.Dense1Pattern, Qt.BrushStyle.Dense2Pattern,
               Qt.BrushStyle.Dense3Pattern, Qt.BrushStyle.Dense4Pattern, Qt.BrushStyle.Dense5Pattern,
               Qt.BrushStyle.HorPattern, Qt.BrushStyle.VerPattern, Qt.BrushStyle.FDiagPattern,
               Qt.BrushStyle.BDiagPattern, Qt.BrushStyle.CrossPattern, Qt.BrushStyle.DiagCrossPattern,
               Qt.BrushStyle.HorPattern, Qt.BrushStyle.VerPattern, Qt.BrushStyle.FDiagPattern,
               Qt.BrushStyle.BDiagPattern, Qt.BrushStyle.CrossPattern, Qt.BrushStyle.Dense6Pattern,
               Qt.BrushStyle.Dense7Pattern, Qt.BrushStyle.HorPattern, Qt.BrushStyle.VerPattern,
               Qt.BrushStyle.FDiagPattern, Qt.BrushStyle.BDiagPattern, Qt.BrushStyle.CrossPattern,
               Qt.BrushStyle.DiagCrossPattern]
STANDARD = [0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0, 0x808080, 0xFF0000,
            0x00FF00, 0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF]      # RGB


def _q(rgb):
    return QColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF)


def _rgb(raw, off):
    """a VtColor of the blob -> 0xRRGGBB, its Automatic"""
    return (raw[off + 2] << 16) | (raw[off + 3] << 8) | raw[off + 4], bool(raw[off + 5] & 0x80)


class ColorBox(QWidget):
    """a box of a colour with its arrow (VCIColorCombo and its VCIBitmapButton): the standard colours, Automatic
    (where the part has one), Other..."""

    def __init__(self, parent, automatic=False):
        super().__init__(parent)
        self.rgb, self.auto, self.can_auto = 0, False, automatic
        self.changed = None
        self.btn = QToolButton(self)
        self.btn.setArrowType(Qt.ArrowType.DownArrow)
        self.btn.clicked.connect(self.menu)

    def resizeEvent(self, e):
        self.btn.setGeometry(self.width() - 16, 2, 15, self.height() - 4)

    def set(self, rgb, auto=False):
        self.rgb, self.auto = rgb, auto
        self.update()

    def paintEvent(self, e):
        p = QPainter(self)
        r = QRect(0, 0, self.width() - 17, self.height() - 1)
        p.fillRect(r, self.palette().base())
        p.setPen(QColor("#808080"))
        p.drawRect(r)
        inner = r.adjusted(2, 2, -2, -2)
        p.fillRect(inner, _q(self.rgb) if self.isEnabled() else self.palette().button().color())
        if self.auto:
            p.setPen(QColor("#ffffff") if self.rgb < 0x808080 else QColor("#000000"))
            p.drawText(inner, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, "Automatic")

    def menu(self):
        m = QMenu(self)
        if self.can_auto:
            m.addAction("Automatic").triggered.connect(lambda: self._pick(self.rgb, True))
        for c in STANDARD:
            pm = QPixmap(24, 12)
            pm.fill(_q(c))
            m.addAction(QIcon(pm), "").triggered.connect(lambda _=False, c=c: self._pick(c, False))
        m.addAction("Other...").triggered.connect(self._other)
        m.exec(self.mapToGlobal(QPoint(0, self.height())))

    def _other(self):
        c = QColorDialog.getColor(_q(self.rgb), self)
        if c.isValid():
            self._pick((c.red() << 16) | (c.green() << 8) | c.blue(), False)

    def _pick(self, rgb, auto):
        self.set(rgb, auto)
        if self.changed is not None:
            self.changed(rgb, auto)


def _line_icon(width_pt=1, style=1, w=60, h=12):
    pm = QPixmap(w, h)
    pm.fill(QColor(0, 0, 0, 0))
    p = QPainter(pm)
    pen = QPen(QColor("#000000"), max(1, int(round(width_pt))))
    pen.setStyle(QT_PENS[style] if 0 <= style < len(QT_PENS) else Qt.PenStyle.SolidLine)
    p.setPen(pen)
    if style:
        p.drawLine(2, h // 2, w - 2, h // 2)
    p.end()
    return QIcon(pm)


def _fill_widths(cb):
    cb.clear()
    for k, wpt in enumerate(WIDTHS):
        cb.addItem("(hairline)" if wpt == 0 else f"{wpt} point")


def _fill_styles(cb):
    cb.clear()
    cb.setIconSize(QSize(56, 12))
    for k, name in enumerate(PEN_STYLES):
        cb.addItem(_line_icon(1, k), "" if k else name)


def _fill_patterns(cb, fore=0x000000, back=0xFFFFFF, none=False):
    """the patterns, drawn in the colours of the part (VB6: its fill colour on its pattern colour)"""
    i = cb.currentIndex()
    cb.blockSignals(True)
    cb.clear()
    cb.setIconSize(QSize(max(40, cb.width() - 24), 12))
    if none:
        cb.addItem("None")
    for k, (st, ix) in enumerate(PATTERNS):
        pm = QPixmap(cb.iconSize())
        pm.fill(_q(back))
        p = QPainter(pm)
        p.fillRect(pm.rect(), QBrush(_q(fore), QT_PATTERNS[k]))
        p.end()
        cb.addItem(QIcon(pm), "")
    cb.setCurrentIndex(max(0, i))
    cb.blockSignals(False)


def _width_index(twentieths):
    pt = float(twentieths) / 20
    return min(range(len(WIDTHS)), key=lambda i: abs(WIDTHS[i] - pt))


class Chart:
    """the chart while its pages are open (one for the dialog): its blob as the pages change it"""

    def __init__(self, dlg):
        from . import vbprops as V
        self.dlg = dlg
        blob = V._frx_bytes(dlg.ctl, "OleObjectBlob")
        if not blob:
            from lbvc.project.chartblob import new_blob
            blob = new_blob(int(V._num(dlg.ctl.get("Width", 4095))), int(V._num(dlg.ctl.get("Height", 3015))))
        self.ed = ChartEdit(blob)
        self.refresh = []
        dlg.tabs.currentChanged.connect(lambda _: [f() for f in self.refresh])

    @staticmethod
    def of(dlg):
        c = dlg.__dict__.get("chart")
        if c is None:
            c = dlg.__dict__["chart"] = Chart(dlg)
        return c

    def changed(self):
        from . import vbprops as V
        from lbvc.project.frx import FrxData
        dlg = self.dlg
        dlg.raw_apply["(Chart)"] = lambda v: V.put(dlg.ctl, "OleObjectBlob", FrxData("raw", v))
        dlg.changed("(Chart)", self.ed.blob())

    # -------------------------------------------------- the parts of the pages
    def f(self, path):
        return self.ed.part(path).f

    def series(self):
        return [x for x in self.f('plot')['series'].f['items'] if x is not None]

    def textblock(self, name):
        """Title, Footnote, X Axis Title, Y Axis Title, Second Y Axis Title"""
        if name in ("title", "footnote"):
            return self.f(name)['tb']
        return self.f('axis(%d)' % {"x": 0, "y": 1, "y2": 2}[name])['title']

    def backdrop(self, name):
        """Chart, Title, Footnote, Legend, Plot"""
        if name == "chart":
            return self.ed.root.f['backdrop']
        return self.f(name)['backdrop']


def _ids(made, *ids):
    from .ocxpages import by_id
    return [by_id(made, i) for i in ids]


def _swap(made, w, wid, automatic=False):
    """the VCIColorCombo of this id (and the VCIBitmapButton beside it) -> a ColorBox in its place"""
    from .ocxpages import by_id
    old = by_id(made, wid)
    g = old.geometry()
    old.hide()
    for it, wd in made:                         # its arrow (a VCIBitmapButton) at its right
        if it["cls"] == "VCIBitmapButton" and abs(it["y"] - g.y()) <= 4 and g.x() < it["x"] < g.x() + g.width():
            wd.hide()
    box = ColorBox(w, automatic)
    box.setGeometry(g.x(), g.y(), g.width(), max(g.height(), 20))
    box.setEnabled(old.isEnabled())
    box.show()
    return box


def _done(dlg, key):
    dlg.bound[key] = {"(Chart)": None}
    dlg.unbound[key] = []


# ---------------------------------------------------------------- Chart
def chart_page(dlg, key, w, made):
    c = Chart.of(dlg)
    d2, d3, lst, legend, stack, markers, rows = _ids(made, 9320, 9321, 9322, 9398, 9399, 1324, 1325)
    busy = [False]

    def show():
        busy[0] = True
        lr = c.f('list')['raw']
        ct = c.ed.get(lr, 32, 'h')
        is3 = any(t3 == ct for _, t3 in TYPES)
        d2.setChecked(not is3)
        d3.setChecked(is3)
        lst.clear()
        names = [n for n, (t2, t3) in zip(TYPE_NAMES, TYPES) if not is3 or t3 is not None]
        lst.addItems(names)
        lst.setCurrentRow(next((i for i, (t2, t3) in enumerate([t for t in TYPES if not is3 or t[1] is not None])
                                if ct in (t2, t3)), 0))
        legend.setChecked(bool(c.ed.get(c.f('legend')['loc'], 6, 'h')))
        ss = c.series()
        stack.setChecked(any(c.ed.get(s.f['raw2'], 44, 'h') for s in ss))
        markers.setChecked(any(c.ed.get(s.f['raw3'], 32, 'h') for s in ss))
        rows.setChecked(bool(c.ed.get(c.ed.root.f['inrow'], 0, 'h')))
        busy[0] = False

    def set_type(*_):
        if busy[0]:
            return
        types = [t for t in TYPES if not d3.isChecked() or t[1] is not None]
        i = max(0, lst.currentRow())
        t2, t3 = types[min(i, len(types) - 1)]
        ct = t3 if d3.isChecked() else t2
        c.ed.put(c.f('list')['raw'], 32, 'h', ct)
        for a in (0, 1, 3):                     # AxisScale.PercentBasis: 0 for X Y, 4 for the others (s_chart2)
            c.ed.put(c.f('axis(%d)' % a)['t2'], 2, 'h', 0 if ct == 16 else 4)
        c.changed()
        show()

    def set_stack(on):
        if busy[0]:
            return
        for k, s in enumerate(c.series()):     # (s_chart2: the series after the first on it, in their order)
            c.ed.put(s.f['raw2'], 42, 'h', 0 if on else k)
            c.ed.put(c.series()[k].f['raw2'], 44, 'h', k if on else 0)
        c.changed()

    def set_flag(path, off, on):
        if busy[0]:
            return
        if path == 'series':
            for k in range(len(c.series())):
                c.ed.put(c.series()[k].f['raw3'], off, 'h', -1 if on else 0)
        else:
            raw = c.f(path)['loc'] if path == 'legend' else c.ed.root.f['inrow']
            c.ed.put(raw, off, 'h', -1 if on else 0)
        c.changed()
    d2.toggled.connect(lambda on: set_type() if on else None)
    d3.toggled.connect(lambda on: set_type() if on else None)
    lst.currentRowChanged.connect(set_type)
    legend.toggled.connect(lambda on: set_flag('legend', 6, on))
    stack.toggled.connect(set_stack)
    markers.toggled.connect(lambda on: set_flag('series', 32, on))
    rows.toggled.connect(lambda on: set_flag('inrow', 0, on))
    c.refresh.append(show)
    show()
    _done(dlg, key)


# ---------------------------------------------------------------- Axis
def axis_page(dlg, key, w, made):
    c = Chart.of(dlg)
    axis, width, show_scale, auto, mn, mx, per_label, per_tick = _ids(made, 1300, 9310, 1301, 9211, 9207, 9212,
                                                                     9214, 9217)
    color = _swap(made, w, 9311)
    _fill_widths(width)
    labels = [wd for it, wd in made if it["cls"] == "Static" and it.get("text", "").startswith(("Mi&n", "&Max", "&Div",
                                                                                                "Divisi"))]
    busy = [False]

    def ax():
        return c.f('axis(%d)' % max(0, axis.currentIndex()))

    def show(*_):
        busy[0] = True
        a = ax()
        width.setCurrentIndex(_width_index(c.ed.get(a['pens'], 14, 'f')))
        color.set(*_rgb(a['pens'], 8))
        show_scale.setChecked(not c.ed.get(a['h'], 0, 'h'))
        has = bool(a['h'][6] & 1)
        auto.setChecked(not has)
        value = axis.currentIndex() in (1, 2)
        for wd in (per_label, per_tick, *labels):
            wd.setEnabled(has)
        for wd in (mn, mx):
            wd.setEnabled(has and value)
        sc = a['scale']
        mn.setText(_num(c.ed.get(sc, 0, 'd')) if sc else "0")
        mx.setText(_num(c.ed.get(sc, 8, 'd')) if sc else "1")
        per_label.setText(str(c.ed.get(sc, 16, 'h')) if sc else "1")
        per_tick.setText(str(c.ed.get(sc, 18, 'h')) if sc else "1")
        busy[0] = False

    def edit(fn):
        if busy[0]:
            return
        fn()
        c.changed()
        show()

    def set_auto(on):
        if busy[0]:
            return
        k = max(0, axis.currentIndex())
        c.ed.set_auto_scale(k, on)
        if not on:                              # the values VB6 gives a scale of its own (s_chart2)
            sc = c.f('axis(%d)' % k)['scale']
            c.ed.put(sc, 8, 'd', 100.0 if k in (1, 2) else 1.0)
            c.ed.put(c.f('axis(%d)' % k)['scale'], 16, 'h', 1 if k in (1, 2) else 5)
        c.changed()
        show()

    def typed(off, fmt, s):
        if busy[0] or not ax()['scale']:
            return
        try:
            v = float(s.replace(",", ".")) if fmt == 'd' else int(s)
        except ValueError:
            return
        c.ed.put(ax()['scale'], off, fmt, v)
        c.changed()
    axis.currentIndexChanged.connect(show)
    width.activated.connect(lambda i: edit(lambda: c.ed.put(ax()['pens'], 14, 'f', WIDTHS[i] * 20.0)))
    color.changed = lambda rgb, a: edit(lambda: c.ed.put_color(ax()['pens'], 8, _bgr(rgb), a))
    show_scale.toggled.connect(lambda on: edit(lambda: c.ed.put(ax()['h'], 0, 'h', 0 if on else -1)))
    auto.toggled.connect(set_auto)
    mn.textEdited.connect(lambda s: typed(0, 'd', s))
    mx.textEdited.connect(lambda s: typed(8, 'd', s))
    per_label.textEdited.connect(lambda s: typed(16, 'h', s))
    per_tick.textEdited.connect(lambda s: typed(18, 'h', s))
    c.refresh.append(show)
    show()
    _done(dlg, key)


def _num(v):
    t = f"{float(v):.6f}".rstrip("0").rstrip(".")
    return t or "0"


def _bgr(rgb):
    """0xRRGGBB -> the value put_color takes (R in the low byte)"""
    return ((rgb >> 16) & 0xFF) | (rgb & 0xFF00) | ((rgb & 0xFF) << 16)


# ---------------------------------------------------------------- Axis Grid
def grid_page(dlg, key, w, made):
    c = Chart.of(dlg)
    axis, s1, w1, s2, w2 = _ids(made, 1300, 9208, 9117, 1341, 9301)
    c1, c2 = _swap(made, w, 1305), _swap(made, w, 1306)
    for s in (s1, s2):
        _fill_styles(s)
    for x in (w1, w2):
        _fill_widths(x)
    busy = [False]

    def pens():
        return c.f('axis(%d)' % max(0, axis.currentIndex()))['pens']

    def show(*_):
        busy[0] = True
        p = pens()
        for base, s, wd, col in ((22, s1, w1, c1), (44, s2, w2, c2)):
            s.setCurrentIndex(max(0, min(c.ed.get(p, base + 2, 'h'), s.count() - 1)))
            wd.setCurrentIndex(_width_index(c.ed.get(p, base + 14, 'f')))
            col.set(*_rgb(p, base + 8))
        busy[0] = False

    def edit(fn):
        if not busy[0]:
            fn()
            c.changed()
    for base, s, wd, col in ((22, s1, w1, c1), (44, s2, w2, c2)):
        s.activated.connect(lambda i, base=base: edit(lambda: c.ed.put(pens(), base + 2, 'h', i)))
        wd.activated.connect(lambda i, base=base: edit(lambda: c.ed.put(pens(), base + 14, 'f', WIDTHS[i] * 20.0)))
        col.changed = lambda rgb, a, base=base: edit(lambda: c.ed.put_color(pens(), base + 8, _bgr(rgb), a))
    axis.currentIndexChanged.connect(show)
    c.refresh.append(show)
    show()
    _done(dlg, key)


# ---------------------------------------------------------------- Series
STAT_FLAGS = {9368: 1, 1312: 2, 1313: 4, 1314: 8, 1315: 16}     # StatLine.Flag: Minimum, Maximum, Mean, SD, Regression
SERIES_TYPES = [0, 5, 7, 9]


def _is_line(c, s):
    """a series drawn as a line: its Edge/Line is its Pen (s_chart3), a bar or an area its EdgePen (s_chart2)"""
    ct = c.ed.get(c.f('list')['raw'], 32, 'h')
    if ct in (8, 9):
        return c.ed.get(s['tail'], 42, 'h') in (5, 9)
    return ct in (2, 3, 6, 7, 16)


def _series_names(c):
    out = []
    for k, s in enumerate(c.series()):
        lt = s.f['dpl'].f['text'] if s.f['dpl'] is not None else None
        out.append(lt.f['s'] if lt is not None else f"C{k + 1}")
    return out


def series_page(dlg, key, w, made):
    c = Chart.of(dlg)
    combo, stype, hide, excl, y2, smark = _ids(made, 1318, 31003, 9132, 31004, 31005, 9367)
    stats = dict(zip(STAT_FLAGS, _ids(made, *STAT_FLAGS)))
    busy = [False]

    def se():
        ss = c.series()
        return ss[max(0, min(combo.currentIndex(), len(ss) - 1))].f

    def show(*_):
        busy[0] = True
        i = max(0, combo.currentIndex())
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(_series_names(c))
        combo.setCurrentIndex(min(i, combo.count() - 1))
        combo.blockSignals(False)
        s = se()
        ct = c.ed.get(c.f('list')['raw'], 32, 'h')
        combo_chart = ct in (8, 9)
        stype.clear()
        stype.addItems(["Bar", "Line", "Area", "Step"] if combo_chart else [TYPE_NAMES[next(
            (k for k, t in enumerate(TYPES) if ct in t), 0)].split("/")[0]])
        if combo_chart:                         # SeriesType: Bar 0, Line 5, Area 7, Step 9 (s_chart3, s_chart4)
            st = c.ed.get(s['tail'], 42, 'h')
            stype.setCurrentIndex(SERIES_TYPES.index(st) if st in SERIES_TYPES else 0)
        stype.setEnabled(combo_chart)
        line = ct in (2, 3, 6, 7, 16) or combo_chart
        flag = c.ed.get(s['raw2'], 32, 'h')
        for wid, cb in stats.items():
            cb.setEnabled(line)
            cb.setChecked(bool(flag & STAT_FLAGS[wid]))
        hide.setChecked(bool(c.ed.get(s['raw2'], 36, 'h')))
        excl.setChecked(bool(c.ed.get(s['raw2'], 34, 'h')))
        y2.setChecked(bool(c.ed.get(s['raw3'], 0, 'h')))
        smark.setChecked(bool(c.ed.get(s['raw3'], 32, 'h')))
        busy[0] = False

    def put(raw, off, v):
        if busy[0]:
            return
        c.ed.put(se()[raw], off, 'h', v)
        c.changed()

    def stat(wid, on):
        if busy[0]:
            return
        f = c.ed.get(se()['raw2'], 32, 'h')
        if on and c.ed.get(se()['raw2'], 80, 'f') < 0:      # the first statistic line: 30 wide (s_chart3)
            c.ed.put(se()['raw2'], 80, 'f', 30.0)
        put('raw2', 32, (f | STAT_FLAGS[wid]) if on else (f & ~STAT_FLAGS[wid]))
    combo.currentIndexChanged.connect(show)
    stype.activated.connect(lambda i: put('tail', 42, SERIES_TYPES[i]))
    hide.toggled.connect(lambda on: put('raw2', 36, -1 if on else 0))
    excl.toggled.connect(lambda on: put('raw2', 34, -1 if on else 0))
    y2.toggled.connect(lambda on: put('raw3', 0, -1 if on else 0))
    smark.toggled.connect(lambda on: put('raw3', 32, -1 if on else 0))
    for wid, cb in stats.items():
        cb.toggled.connect(lambda on, wid=wid: stat(wid, on))
    c.refresh.append(show)
    show()
    _done(dlg, key)


# ---------------------------------------------------------------- Series Color
def series_color_page(dlg, key, w, made):
    c = Chart.of(dlg)
    combo, stype, pattern, estyle, ewidth = _ids(made, 1318, 31003, 31001, 31002, 9121)
    fill, pcolor, ecolor = _swap(made, w, 1308), _swap(made, w, 1307), _swap(made, w, 1309, automatic=True)
    _fill_patterns(pattern)
    _fill_styles(estyle)
    _fill_widths(ewidth)
    busy = [False]

    def se():
        ss = c.series()
        return ss[max(0, min(combo.currentIndex(), len(ss) - 1))].f

    def show(*_):
        busy[0] = True
        i = max(0, combo.currentIndex())
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(_series_names(c))
        combo.setCurrentIndex(min(i, combo.count() - 1))
        combo.blockSignals(False)
        ct = c.ed.get(c.f('list')['raw'], 32, 'h')
        stype.clear()
        stype.addItem(TYPE_NAMES[next((k for k, t in enumerate(TYPES) if ct in t), 0)].split("/")[0])
        b = se()['raw']
        fill.set(*_rgb(b, 46 + 2))
        pcolor.set(*_rgb(b, 46 + 8))
        _fill_patterns(pattern, _rgb(b, 46 + 2)[0], _rgb(b, 46 + 8)[0])
        st, ix = c.ed.get(b, 46, 'h'), c.ed.get(b, 46 + 14, 'h')
        pattern.setCurrentIndex(PATTERNS.index((st, ix)) if (st, ix) in PATTERNS else 0)
        e = 0 if _is_line(c, se()) else 22
        estyle.setCurrentIndex(max(0, min(c.ed.get(b, e + 2, 'h'), estyle.count() - 1)))
        ewidth.setCurrentIndex(_width_index(c.ed.get(b, e + 14, 'f')))
        ecolor.set(*_rgb(b, e + 8))
        busy[0] = False

    def edit(fn):
        if not busy[0]:
            fn()
            c.changed()

    def set_pattern(i):
        st, ix = PATTERNS[i]
        c.ed.put(se()['raw'], 46, 'h', st)
        c.ed.put(se()['raw'], 46 + 14, 'h', ix)
    combo.currentIndexChanged.connect(show)
    fill.changed = lambda rgb, a: edit(lambda: c.ed.put_color(se()['raw'], 46 + 2, _bgr(rgb), a))
    pcolor.changed = lambda rgb, a: edit(lambda: c.ed.put_color(se()['raw'], 46 + 8, _bgr(rgb), a))
    def edge():
        return 0 if _is_line(c, se()) else 22

    def set_width(i):
        c.ed.put(se()['raw'], edge() + 14, 'f', WIDTHS[i] * 20.0)
        if edge() == 0:                         # a line: its markers too (s_chart4)
            c.ed.put(se()['raw3'], 6 + 14, 'f', WIDTHS[i] * 20.0)
    ecolor.changed = lambda rgb, a: edit(lambda: c.ed.put_color(se()['raw'], edge() + 8, _bgr(rgb), a))
    pattern.activated.connect(lambda i: edit(lambda: set_pattern(i)))
    estyle.activated.connect(lambda i: edit(lambda: c.ed.put(se()['raw'], edge() + 2, 'h', i)))
    ewidth.activated.connect(lambda i: edit(lambda: set_width(i)))
    c.refresh.append(show)
    show()
    _done(dlg, key)


# ---------------------------------------------------------------- Backdrop
BACKDROPS = ["chart", "title", "footnote", "legend", "plot"]


def backdrop_page(dlg, key, w, made):
    c = Chart.of(dlg)
    combo, pattern, border, bwidth, shadow = _ids(made, 9108, 1327, 1329, 1330, 1332)
    fill, pcolor, bcolor = _swap(made, w, 1326), _swap(made, w, 1328), _swap(made, w, 1331)
    _fill_patterns(pattern, none=True)
    flabels = [wd for it, wd in made if it["cls"] == "Static" and plain_text(it) in ("Color:", "Pattern Color:")]
    busy = [False]

    def bd():
        return c.backdrop(BACKDROPS[max(0, combo.currentIndex())]).f

    def show(*_):
        busy[0] = True
        b = bd()
        fr, fl = b['raw'], b['fill'].f
        fstyle = c.ed.get(fl['style'], 0, 'h')    # Fill.Style: 0 none, 1 a brush
        st, ix = c.ed.get(fl['raw'], 2, 'h'), c.ed.get(fl['raw'], 2 + 14, 'h')
        pattern.setCurrentIndex(0 if not fstyle else 1 + (PATTERNS.index((st, ix)) if (st, ix) in PATTERNS else 0))
        fill.set(*_rgb(fl['raw'], 2 + 2))
        pcolor.set(*_rgb(fl['raw'], 2 + 8))
        _fill_patterns(pattern, _rgb(fl['raw'], 2 + 2)[0], _rgb(fl['raw'], 2 + 8)[0], none=True)
        pattern.setCurrentIndex(0 if not fstyle else 1 + (PATTERNS.index((st, ix)) if (st, ix) in PATTERNS else 0))
        on = bool(fstyle)
        fill.setEnabled(on)
        pcolor.setEnabled(on and st != 1)
        for lab in flabels:
            lab.setEnabled(on)
        border.setCurrentIndex(max(0, min(c.ed.get(fr, 48, 'h'), 4)))
        bwidth.setText(_num(c.ed.get(fr, 44, 'f') / 20) + "pt")     # (twentieths of a point: 3pt is 60)
        bcolor.set(*_rgb(fr, 32))
        shadow.setChecked(bool(c.ed.get(fr, 28, 'h')))
        busy[0] = False

    def edit(fn):
        if not busy[0]:
            fn()
            c.changed()
            show()

    def set_pattern(i):
        fl = bd()['fill'].f
        if i == 0:
            c.ed.put(fl['style'], 0, 'h', 0)
            return
        st, ix = PATTERNS[i - 1]
        c.ed.put(bd()['fill'].f['style'], 0, 'h', 1)
        c.ed.put(bd()['fill'].f['raw'], 2, 'h', st)
        c.ed.put(bd()['fill'].f['raw'], 2 + 14, 'h', ix)

    def width_typed(s):
        if busy[0]:
            return
        try:
            v = float(s.lower().replace("pt", "").strip().replace(",", "."))
        except ValueError:
            return
        c.ed.put(bd()['raw'], 44, 'f', v * 20)
        c.changed()
    combo.currentIndexChanged.connect(show)
    pattern.activated.connect(lambda i: edit(lambda: set_pattern(i)))
    fill.changed = lambda rgb, a: edit(lambda: c.ed.put_color(bd()['fill'].f['raw'], 2 + 2, _bgr(rgb), a))
    pcolor.changed = lambda rgb, a: edit(lambda: c.ed.put_color(bd()['fill'].f['raw'], 2 + 8, _bgr(rgb), a))
    bcolor.changed = lambda rgb, a: edit(lambda: c.ed.put_color(bd()['raw'], 32, _bgr(rgb), a))
    border.activated.connect(lambda i: edit(lambda: c.ed.put(bd()['raw'], 48, 'h', i)))
    bwidth.textEdited.connect(width_typed)
    shadow.toggled.connect(lambda on: edit(lambda: c.ed.put(bd()['raw'], 28, 'h', 1 if on else 0)))
    c.refresh.append(show)
    show()
    _done(dlg, key)


def plain_text(it):
    return it.get("text", "").replace("&", "")


# ---------------------------------------------------------------- Text
TEXTS = ["title", "footnote", "x", "y", "y2"]
HORZ = {9241: 0, 9242: 2, 9243: 1}              # Left, Center, Right (VtHorizontalAlignment)
VERT = {9244: 0, 9245: 2, 9246: 1}              # Top, Center, Bottom
ORIENT = {1320: 0, 1321: 2, 1333: 1, 1322: 3}   # Horizontal, Up, Vertical (letters down), Down (s_chart2)


def text_page(dlg, key, w, made):
    c = Chart.of(dlg)
    combo, text = _ids(made, 1319, 9262)
    btn = dict(zip((*HORZ, *VERT, *ORIENT), _ids(made, *HORZ, *VERT, *ORIENT)))
    for wid, b in btn.items():
        b.setCheckable(True)
        b.setText("")
        b.setIcon(_text_icon(wid, b.size()))
        b.setIconSize(b.size() - QSize(4, 4))
    busy = [False]

    def tb():
        return c.textblock(TEXTS[max(0, combo.currentIndex())])

    def show(*_):
        busy[0] = True
        t = tb()
        s = t.f['text'].f['s'] if t.f['text'] is not None else ""
        if text.toPlainText() != s:
            text.setPlainText(s)
        lay = t.f['layout']
        h, v, o = c.ed.get(lay, 4, 'h'), c.ed.get(lay, 6, 'h'), c.ed.get(lay, 8, 'h')
        for wid, b in btn.items():
            b.setChecked((wid in HORZ and HORZ[wid] == h) or (wid in VERT and VERT[wid] == v) or
                         (wid in ORIENT and ORIENT[wid] == o))
        busy[0] = False

    def typed():
        if busy[0]:
            return
        s = text.toPlainText()
        c.ed.set_string(tb(), 'text', s)
        k = max(0, combo.currentIndex())
        if k < 2:                               # a title or a footnote with a text is shown (s_chart2)
            c.ed.put(c.f(TEXTS[k])['loc'], 6, 'h', -1 if s else 0)
        c.changed()

    def press(wid):
        if busy[0]:
            return
        off, v = (4, HORZ[wid]) if wid in HORZ else (6, VERT[wid]) if wid in VERT else (8, ORIENT[wid])
        c.ed.put(tb().f['layout'], off, 'h', v)
        c.changed()
        show()
    combo.currentIndexChanged.connect(show)
    text.textChanged.connect(typed)
    for wid, b in btn.items():
        b.clicked.connect(lambda _=False, wid=wid: press(wid))
    c.refresh.append(show)
    show()
    _done(dlg, key)


def _text_icon(wid, size):
    """what the buttons of the Text page show: the lines of an alignment, the word Text turned"""
    from PyQt6.QtGui import QFont
    w, h = max(8, size.width() - 4), max(8, size.height() - 4)
    pm = QPixmap(w, h)
    pm.fill(QColor(0, 0, 0, 0))
    p = QPainter(pm)
    p.setPen(QColor("#000000"))
    if wid in HORZ:                             # four lines, to the left / the middle / the right
        for k, ln in enumerate((w - 12, w - 18, w - 12, w - 18)):
            x = {0: 3, 2: (w - ln) // 2, 1: w - 3 - ln}[HORZ[wid]]
            p.drawLine(x, 4 + k * 3, x + ln, 4 + k * 3)
    elif wid in VERT:                           # a line at the top / the middle / the bottom of a dotted box
        p.setPen(QPen(QColor("#808080"), 1, Qt.PenStyle.DotLine))
        for y in (3, h // 2, h - 4):
            p.drawLine(3, y, w - 3, y)
        p.setPen(QPen(QColor("#000000"), 2))
        y = {0: 3, 2: h // 2, 1: h - 4}[VERT[wid]]
        p.drawLine(3, y, w - 3, y)
    else:                                       # the word Text: across, down letter by letter, up, down
        f = QFont()
        f.setBold(True)
        f.setPixelSize(11)
        p.setFont(f)
        o = ORIENT[wid]
        if o == 0:
            p.drawText(QRect(0, 0, w, h), Qt.AlignmentFlag.AlignCenter, "Text")
        elif o == 1:
            f.setPixelSize(max(7, (h - 2) // 4))
            p.setFont(f)
            p.drawText(QRect(0, 0, w, h), Qt.AlignmentFlag.AlignCenter, "T\ne\nx\nt")
        else:
            p.translate(w / 2, h / 2)
            p.rotate(-90 if o == 2 else 90)
            p.drawText(QRect(-h // 2, -w // 2, h, w), Qt.AlignmentFlag.AlignCenter, "Text")
    p.end()
    return QIcon(pm)


# ---------------------------------------------------------------- Fonts
def _fonts(c, k):
    """the VtFont objects of the item k of the Fonts page: Title, Footnote, Legend, Axis Labels, Axis Title"""
    if k == 0:
        return [c.f('title')['tb'].f['font']]
    if k == 1:
        return [c.f('footnote')['tb'].f['font']]
    if k == 2:
        return [c.f('legend')['font']]
    out = []
    for a in range(4):
        ax = c.f('axis(%d)' % a)
        if k == 3:
            out += [lb.f['font'] for lb in ax['labels'].f['items'] if lb is not None and 'font' in lb.f]
        else:
            out.append(ax['title'].f['font'])
    return [f for f in out if f is not None]


def fonts_page(dlg, key, w, made):
    from PyQt6.QtGui import QFontDatabase, QFont
    c = Chart.of(dlg)
    combo, name, style, size, strike, under, sample = _ids(made, 1317, 5128, 5129, 5130, 5131, 5132, 5133)
    color = _swap(made, w, 5137)
    name.clear()
    name.addItems(QFontDatabase.families())
    name.setEditable(True)
    size.setEditable(True)
    busy = [False]

    def fonts():
        return _fonts(c, max(0, combo.currentIndex()))

    def show(*_):
        busy[0] = True
        fs = fonts()
        if not fs:
            busy[0] = False
            return
        fo = fs[0].f
        nm = fo['name'].f['s'] if fo['name'] is not None else "MS Sans Serif"
        name.setEditText(nm)
        b = fo['raw']
        sz = c.ed.get(b, 0, 'f')
        size.setEditText(str(int(sz)) if sz - int(sz) < 0.3 else _num(sz))      # (VB6: 8.25 shows as 8)
        st = c.ed.get(b, 4, 'h')
        style.setCurrentIndex({0: 0, 1: 1, 2: 2, 3: 3}.get(st & 3, 0))
        ef = c.ed.get(b, 6, 'h')
        strike.setChecked(bool(ef & 0x100))     # VtFontEffect: StrikeThrough 256, Underline 512
        under.setChecked(bool(ef & 0x200))
        color.set(*_rgb(b, 8))
        f = QFont(nm)
        f.setPointSizeF(max(1.0, c.ed.get(b, 0, 'f')))
        f.setBold(bool(st & 1))
        f.setItalic(bool(st & 2))
        f.setStrikeOut(bool(ef & 0x100))
        f.setUnderline(bool(ef & 0x200))
        sample.setFont(f)
        busy[0] = False

    def each(fn):
        if busy[0]:
            return
        try:                                    # VB6 writes the size the page shows with any change (8.25 -> 8)
            sz = float(size.currentText().replace(",", "."))
        except ValueError:
            sz = None
        for i in range(len(fonts())):
            fn(fonts()[i])
            if sz is not None:
                c.ed.put(fonts()[i].f['raw'], 0, 'f', sz)
        c.changed()
        show()

    def set_name(s):
        each(lambda fo: c.ed.set_string(fo, 'name', s))

    def set_size(s):
        try:
            v = float(s.replace(",", "."))
        except ValueError:
            return
        each(lambda fo: c.ed.put(fo.f['raw'], 0, 'f', v))

    def set_bits(off, mask, on):
        each(lambda fo: c.ed.put(fo.f['raw'], off, 'h',
                                 (c.ed.get(fo.f['raw'], off, 'h') | mask) if on else
                                 (c.ed.get(fo.f['raw'], off, 'h') & ~mask)))

    def set_style(i):
        each(lambda fo: c.ed.put(fo.f['raw'], 4, 'h', (c.ed.get(fo.f['raw'], 4, 'h') & ~3) | i))
    combo.currentIndexChanged.connect(show)
    name.activated.connect(lambda i: set_name(name.itemText(i)))
    name.lineEdit().editingFinished.connect(lambda: set_name(name.currentText()))
    style.activated.connect(set_style)
    size.activated.connect(lambda i: set_size(size.itemText(i)))
    size.lineEdit().editingFinished.connect(lambda: set_size(size.currentText()))
    strike.toggled.connect(lambda on: set_bits(6, 0x100, on))
    under.toggled.connect(lambda on: set_bits(6, 0x200, on))
    color.changed = lambda rgb, a: each(lambda fo: c.ed.put_color(fo.f['raw'], 8, _bgr(rgb), a))
    c.refresh.append(show)
    show()
    _done(dlg, key)


PAGES = {("MSChart", "Chart"): chart_page, ("MSChart", "Axis"): axis_page, ("MSChart", "Axis Grid"): grid_page,
         ("MSChart", "Series"): series_page, ("MSChart", "Series Color"): series_color_page,
         ("MSChart", "Backdrop"): backdrop_page, ("MSChart", "Text"): text_page, ("MSChart", "Fonts"): fonts_page}
