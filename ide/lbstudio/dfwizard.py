"""The Data Form Wizard (phase 9.2b): Add-Ins > Data Form Wizard... when the "VB 6 Data Form Wizard" add-in is loaded,
and Create New Form... of the Data Access Forms page of the Application Wizard (VB6 oracle DEV-9/df1, df2: the
pages, their questions and the names of the choices are VB6's; the wording, the pictures and the code of the
forms are our own: dfwizgen).

Pages: Introduction (a profile .rwp), Database Type (Access / Remote(ODBC)), Database or Connect Information, Form
(name, Form Layout, Binding Type), Record Source (Master Record Source, Detail Record Source and Record Source
Relation for Master/Detail), the pages of the MS HFlexGrid (Select Grid Type, Set Appearance Style, Set Column
Settings, Set Application UI Options) or of the MS Chart (Select Chart Fields, Select Chart Style, Select Appearance
Properties, Create Chart Captions), Control Selection, Finished! (a profile to save). Finish: the form (and its
class) in the project, "Data Form Created".
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
import math
import os
import re
import subprocess

from PyQt6.QtCore import Qt, QRectF, QPointF, pyqtSignal
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPen, QIcon, QPolygonF, QBrush
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton, QLineEdit,
                             QListWidget, QListWidgetItem, QStackedWidget, QRadioButton, QButtonGroup, QCheckBox,
                             QComboBox, QFrame, QMessageBox, QFileDialog, QGroupBox, QWidget, QTableWidget,
                             QTableWidgetItem, QSlider, QPlainTextEdit, QAbstractItemView)

from . import dfwizgen as G
from . import dbinfo
from .appwizard import profiles_dir, wizard_picture

TITLE = "Data Form Wizard"
NONE = "(None)"
PROFILE_FILTER = "Wizard Profile (*.rwp)"
PAGE_TITLES = {"intro": "Introduction", "dbtype": "Database Type", "database": "Database",
               "connect": "Connect Information", "form": "Form", "source": "Record Source",
               "master": "Master Record Source", "detail": "Detail Record Source",
               "relation": "Record Source Relation", "gridtype": "Select Grid Type",
               "gridstyle": "Set Appearance Style", "columns": "Set Column Settings",
               "uioptions": "Set Application UI Options", "chartfields": "Select Chart Fields",
               "chartstyle": "Select Chart Style", "chartlook": "Select Appearance Properties",
               "captions": "Create Chart Captions", "controls": "Control Selection", "finished": "Finished!"}
# the pages without the picture on the left (their preview is on the right)
NO_PICTURE = {"gridstyle", "columns", "uioptions", "chartfields", "chartstyle", "chartlook", "captions"}
GRID_HINTS = ("Standard: the records one under the other, as in a datasheet.",
              "Outline: equal values next to each other are merged into one cell, so the records look grouped.")


def page_keys(df):
    """the pages in their order for the choices so far"""
    keys = ["intro", "dbtype", "database" if df.dbtype == 0 else "connect", "form"]
    if df.layout == G.L_MD:
        keys += ["master", "detail", "relation"]
    else:
        keys += ["source"]
    if df.layout == G.L_FLEX:
        keys += ["gridtype", "gridstyle", "columns", "uioptions"]
    elif df.layout == G.L_CHART:
        keys += ["chartfields", "chartstyle", "chartlook", "captions"]
    else:
        keys += ["controls"]
    return keys + ["finished"]


# ------------------------------------------------------------------------------------------------ the pictures
def df_picture(kind, w=150, h=120):
    """the picture on the left of a page: our own drawings"""
    if kind == "finished":
        return wizard_picture("finished", w, h)
    pm = QPixmap(w, h)
    pm.fill(QColor("#808080"))
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    ink, paper, blue = QColor("#202020"), QColor("#FFFFFF"), QColor("#2F6FD0")

    def can(x, y, ww, hh, top):
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor(top).darker(130))
        p.drawRect(QRectF(x, y + hh * 0.15, ww, hh * 0.7))
        p.drawEllipse(QRectF(x, y + hh * 0.7, ww, hh * 0.3))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawRect(QRectF(x + 0.5, y + hh * 0.7, ww - 1, hh * 0.15))
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor(top))
        p.drawEllipse(QRectF(x, y, ww, hh * 0.3))

    def window(x, y, ww, hh):
        p.setPen(QPen(ink, 1))
        p.setBrush(paper)
        p.drawRect(QRectF(x, y, ww, hh))
        p.setBrush(blue)
        p.drawRect(QRectF(x, y, ww, 7))

    def rows(x, y, ww, n, step=8):
        p.setPen(QPen(ink, 1))
        for i in range(n):
            p.drawLine(QPointF(x, y + i * step), QPointF(x + ww * (0.5 + 0.5 * ((i * 37) % 10) / 10), y + i * step))

    def table(x, y, ww, hh, nr=4, nc=4):
        p.setPen(QPen(ink, 1))
        p.setBrush(paper)
        p.drawRect(QRectF(x, y, ww, hh))
        p.fillRect(QRectF(x + 1, y + 1, ww - 1, hh / nr - 1), QColor("#C8C8C8"))
        for i in range(1, nr):
            p.drawLine(QPointF(x, y + i * hh / nr), QPointF(x + ww, y + i * hh / nr))
        for j in range(1, nc):
            p.drawLine(QPointF(x + j * ww / nc, y), QPointF(x + j * ww / nc, y + hh))

    if kind in ("intro", "dbs"):
        can(30, 20, 46, 46, "#7FD8D8")
        can(78, 34, 46, 46, "#E8E8E8")
        if kind == "dbs":
            can(56, 62, 40, 40, "#E87FE0")
        else:
            window(40, 66, 70, 44)
            rows(48, 82, 50, 3)
    elif kind == "dbfile":
        can(70, 18, 52, 52, "#E8E8E8")
        window(22, 52, 62, 50)
        p.setPen(QPen(ink, 1))
        for i, x in enumerate((30, 38, 38, 46)):
            p.fillRect(QRectF(x, 64 + i * 9, 6, 5), QColor("#E0C040"))
            p.drawLine(QPointF(x + 7, 66 + i * 9), QPointF(76, 66 + i * 9))
    elif kind == "connect":
        can(80, 16, 46, 46, "#7FD8D8")
        table(24, 64, 54, 40)
        p.setPen(QPen(QColor("#F0D020"), 2))
        p.drawLine(QPointF(78, 80), QPointF(98, 80))
        p.drawLine(QPointF(98, 80), QPointF(98, 62))
    elif kind.startswith("form"):
        lay = int(kind[4:] or 0)
        window(30, 18, 90, 84)
        if lay == G.L_SINGLE:
            rows(42, 36, 60, 5)
        elif lay == G.L_GRID:
            table(36, 30, 78, 52, 5, 4)
        elif lay == G.L_MD:
            rows(42, 32, 50, 3)
            table(36, 58, 78, 30, 3, 4)
        elif lay == G.L_FLEX:
            table(36, 30, 78, 60, 6, 4)
            p.fillRect(QRectF(36.5, 40.5, 19, 39), QColor("#C8C8C8"))
        else:
            p.setPen(QPen(ink, 1))
            p.setBrush(QColor("#F0E040"))
            p.drawPie(QRectF(44, 36, 46, 46), 0, 120 * 16)
            p.setBrush(QColor("#4060E0"))
            p.drawPie(QRectF(44, 36, 46, 46), 120 * 16, 140 * 16)
            p.setBrush(QColor("#A040A0"))
            p.drawPie(QRectF(44, 36, 46, 46), 260 * 16, 100 * 16)
        p.setBrush(QColor("#D0D0D0"))
        for i in range(4):
            p.drawRect(QRectF(36 + i * 20, 92, 16, 6))
    elif kind == "source":
        table(18, 16, 56, 44, 3, 2)
        p.setPen(QPen(QColor("#00C8D8"), 6))
        p.drawLine(QPointF(30, 60), QPointF(30, 86))
        p.drawLine(QPointF(30, 86), QPointF(58, 86))
        window(60, 58, 74, 50)
        rows(68, 74, 50, 2)
        table(66, 90, 62, 14, 1, 4)
    elif kind == "relation":
        table(20, 30, 50, 66, 5, 2)
        table(84, 18, 50, 66, 5, 2)
        p.fillRect(QRectF(21, 57, 49, 12), QColor("#40D8E0"))
        p.fillRect(QRectF(85, 32, 49, 12), QColor("#40D8E0"))
        p.setPen(QPen(ink, 1))
        p.drawLine(QPointF(70, 63), QPointF(84, 38))
    elif kind == "controls":
        window(32, 20, 86, 76)
        rows(42, 38, 50, 4)
        p.setBrush(QColor("#D0D0D0"))
        for i in range(4):
            p.drawRect(QRectF(38 + i * 20, 84, 16, 7))
        p.setPen(QPen(QColor("#F0D020"), 3))
        p.drawLine(QPointF(14, 88), QPointF(34, 88))
    p.end()
    return pm


def dfwizard_icon(size=16):
    """the icon of Add-Ins > Data Form Wizard..."""
    pm = df_picture("intro").copy(26, 10, 100, 100).scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio,
                                                            Qt.TransformationMode.SmoothTransformation)
    return QIcon(pm)


# ------------------------------------------------------------------------------------------------ the database
def odbc_names(flag):
    """the DSNs (-s) or the drivers (-d) of unixODBC"""
    try:
        r = subprocess.run(["odbcinst", "-q", flag], capture_output=True, text=True, timeout=10)
    except (OSError, subprocess.TimeoutExpired):
        return []
    return [line.strip().strip("[]") for line in r.stdout.splitlines() if line.strip().startswith("[")]


class Schema:
    """the record sources and their fields, for the database of the wizard"""
    def __init__(self, df):
        self.df = df

    def sources(self):
        df = self.df
        if df.dbtype == 0:
            return dbinfo.objects(df.database)
        return [n for _k, n in dbinfo.ado_objects(G.connect_string(df))]

    def fields(self, source):
        df = self.df
        if not source:
            return []
        if df.dbtype == 0:
            return dbinfo.fields(df.database, source)
        return [f[0] for f in dbinfo.ado_fields(G.connect_string(df), 2, source)]


# ------------------------------------------------------------------------------------------------ the widgets
class SourcePicker(QWidget):
    """Record Source: the source, Available Fields -> Selected Fields, Column to Sort By"""
    changed = pyqtSignal()

    def __init__(self, schema):
        super().__init__()
        self.schema = schema
        lay = QGridLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(QLabel("Record Source:"), 0, 2, 1, 2)
        self.source = QComboBox()
        lay.addWidget(self.source, 1, 2, 1, 2)
        lay.addWidget(QLabel("&Available Fields:"), 2, 0)
        lay.addWidget(QLabel("&Selected Fields:"), 2, 2)
        self.avail = QListWidget()
        self.chosen = QListWidget()
        for lw in (self.avail, self.chosen):
            lw.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        lay.addWidget(self.avail, 3, 0)
        lay.addWidget(self.chosen, 3, 2)
        mid = QVBoxLayout()
        self.b_add, self.b_all, self.b_del, self.b_none = (QPushButton(t) for t in (">", ">>", "<", "<<"))
        for b in (self.b_add, self.b_all, self.b_del, self.b_none):
            b.setFixedWidth(32)
            mid.addWidget(b)
        mid.addStretch(1)
        lay.addLayout(mid, 3, 1)
        side = QVBoxLayout()
        self.b_up, self.b_down = QPushButton("↑"), QPushButton("↓")
        for b in (self.b_up, self.b_down):
            b.setFixedWidth(32)
            side.addWidget(b)
        side.addStretch(1)
        lay.addLayout(side, 3, 3)
        lay.addWidget(QLabel("Column to Sort By:"), 4, 2)
        self.sort = QComboBox()
        lay.addWidget(self.sort, 5, 2)
        self.source.activated.connect(lambda _i: self.source_chosen())
        self.b_add.clicked.connect(lambda: self.move(self.avail, self.chosen, False))
        self.b_all.clicked.connect(lambda: self.move(self.avail, self.chosen, True))
        self.b_del.clicked.connect(lambda: self.move(self.chosen, self.avail, False))
        self.b_none.clicked.connect(lambda: self.move(self.chosen, self.avail, True))
        self.avail.itemDoubleClicked.connect(lambda _it: self.move(self.avail, self.chosen, False))
        self.chosen.itemDoubleClicked.connect(lambda _it: self.move(self.chosen, self.avail, False))
        self.b_up.clicked.connect(lambda: self.shift(-1))
        self.b_down.clicked.connect(lambda: self.shift(1))
        self.all_fields = []

    def fill_sources(self, names, current=""):
        self.source.blockSignals(True)
        self.source.clear()
        self.source.addItems(names)
        self.source.setCurrentIndex(names.index(current) if current in names else -1)
        self.source.blockSignals(False)

    def source_chosen(self):
        """a new source: all its fields chosen (VB6 df1-13)"""
        self.set_state(self.source.currentText(), None, "")
        self.changed.emit()

    def set_state(self, source, fields, sort):
        """source, the fields chosen (None: all of them), the sort column"""
        if source and self.source.findText(source) < 0:
            self.source.addItem(source)
        self.source.setCurrentIndex(self.source.findText(source) if source else -1)
        try:
            self.all_fields = self.schema.fields(source) if source else []
        except dbinfo.AdoError as e:
            QMessageBox.warning(self, TITLE, str(e))
            self.all_fields = []
        chosen = list(self.all_fields) if fields is None else [f for f in fields if f in self.all_fields]
        self.avail.clear()
        self.avail.addItems([f for f in self.all_fields if f not in chosen])
        self.chosen.clear()
        self.chosen.addItems(chosen)
        self.fill_sort(sort)

    def fill_sort(self, sort=None):
        cur = self.sort.currentText() if sort is None else sort
        self.sort.clear()
        self.sort.addItem(NONE)
        self.sort.addItems(self.fields())
        i = self.sort.findText(cur)
        self.sort.setCurrentIndex(i if i >= 0 else 0)

    def move(self, src, dst, everything):
        items = [src.item(i) for i in range(src.count())] if everything else src.selectedItems()
        names = [it.text() for it in items]
        for it in items:
            src.takeItem(src.row(it))
        if dst is self.chosen:
            dst.addItems(names)
        else:                                   # back among the available ones in the order of the source
            have = [dst.item(i).text() for i in range(dst.count())] + names
            dst.clear()
            dst.addItems([f for f in self.all_fields if f in have])
        self.fill_sort()
        self.changed.emit()

    def shift(self, d):
        r = self.chosen.currentRow()
        if r < 0 or not 0 <= r + d < self.chosen.count():
            return
        it = self.chosen.takeItem(r)
        self.chosen.insertItem(r + d, it)
        self.chosen.setCurrentRow(r + d)
        self.changed.emit()

    def fields(self):
        return [self.chosen.item(i).text() for i in range(self.chosen.count())]

    def sort_field(self):
        t = self.sort.currentText()
        return "" if t == NONE else t


def style_grid(table, style, fields, rows=5):
    """the preview of a style of the hierarchical flex grid"""
    st = G.FLEX_STYLES[style]

    def col(v, sys_default):
        if v < 0:
            return QColor(sys_default)
        return QColor(v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF)
    back, fore = col(st[0], "#FFFFFF"), col(st[1], "#000000")
    bfix, ffix = col(st[2], "#D4D0C8"), col(st[3], "#000000")
    table.clear()
    table.setRowCount(rows)
    table.setColumnCount(max(1, len(fields)))
    table.horizontalHeader().hide()
    table.verticalHeader().hide()
    for c, f in enumerate(fields):
        it = QTableWidgetItem(f)
        it.setBackground(QBrush(bfix))
        it.setForeground(QBrush(ffix))
        f2 = it.font()
        f2.setBold(True)
        it.setFont(f2)
        table.setItem(0, c, it)
        for r in range(1, rows):
            it = QTableWidgetItem("")
            it.setBackground(QBrush(back if style != 1 or r % 2 else QColor("#D8D8D8")))
            it.setForeground(QBrush(fore))
            table.setItem(r, c, it)


class ChartPreview(QWidget):
    """a sample chart of a style (our own drawing)"""
    def __init__(self):
        super().__init__()
        self.setMinimumSize(220, 200)
        self.style = 7
        self.legend = False
        self.markers = False
        self.stack = False
        self.rotation = 60

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        r = QRectF(self.rect()).adjusted(20, 14, -20 - (40 if self.legend else 0), -20)
        name, _t = G.CHART_STYLES[self.style]
        data = ((30, 55, 40), (70, 35, 60), (45, 80, 30), (90, 50, 75))
        colors = (QColor("#E02020"), QColor("#20B020"), QColor("#3050E0"))
        ink = QColor("#202020")
        if "Pie" in name:
            tot = sum(d[0] for d in data)
            a = 0
            pie = QRectF(r.center().x() - min(r.width(), r.height()) / 2, r.top(), min(r.width(), r.height()),
                         min(r.width(), r.height()))
            for i, d in enumerate(data):
                span = int(360 * 16 * d[0] / tot)
                p.setBrush(QColor.fromHsv((i * 80) % 360, 200, 230))
                p.setPen(QPen(ink, 1))
                p.drawPie(pie, a, span)
                a += span
            p.end()
            return
        three = name.startswith("3D")
        dx = 10 * math.cos(math.radians(self.rotation)) if three else 0
        dy = -8 if three else 0
        p.setPen(QPen(ink, 1))
        p.drawLine(r.bottomLeft(), r.bottomRight())
        p.drawLine(r.bottomLeft(), r.topLeft())
        n = len(data)
        cw = r.width() / n
        series = 1 if self.stack else 3
        for s in range(3):
            pts = []
            for i, d in enumerate(data):
                base = sum(dd for dd in d[:s]) if self.stack else 0
                v = d[s] * (0.5 if self.stack else 1)
                x0 = r.left() + i * cw + (cw * 0.15) + (0 if self.stack else s * cw * 0.7 / series / 3)
                yv = r.bottom() - (base * 0.5 + v) / 110 * r.height() if self.stack else r.bottom() - v / 110 * r.height()
                y0 = r.bottom() - base * 0.5 / 110 * r.height() if self.stack else r.bottom()
                pts.append(QPointF(x0 + cw * 0.3, yv))
                if "Bar" in name or ("Combination" in name and s == 0):
                    bw = cw * (0.7 if self.stack else 0.22)
                    if three:
                        p.setBrush(colors[s].darker(150))
                        p.drawPolygon(QPolygonF([QPointF(x0 + bw, yv), QPointF(x0 + bw + dx, yv + dy),
                                                 QPointF(x0 + bw + dx, y0 + dy), QPointF(x0 + bw, y0)]))
                    p.setBrush(colors[s])
                    p.drawRect(QRectF(x0, yv, bw, y0 - yv))
            if "Line" in name or ("Combination" in name and s > 0):
                p.setPen(QPen(colors[s], 2))
                p.drawPolyline(QPolygonF(pts))
                p.setPen(QPen(ink, 1))
            elif "Area" in name or "Step" in name:
                poly = [QPointF(r.left(), r.bottom())]
                for i, q in enumerate(pts):
                    if "Step" in name:
                        poly += [QPointF(r.left() + i * cw, q.y()), QPointF(r.left() + (i + 1) * cw, q.y())]
                    else:
                        poly.append(QPointF(r.left() + (i + 0.5) * cw, q.y()))
                poly.append(QPointF(r.right(), r.bottom()))
                p.setBrush(colors[s])
                p.drawPolygon(QPolygonF(poly))
            if self.markers:
                p.setBrush(QColor("#FFFFFF"))
                for q in pts:
                    p.drawRect(QRectF(q.x() - 2, q.y() - 2, 4, 4))
        if self.legend:
            for s in range(3):
                y = r.top() + 10 + s * 16
                p.fillRect(QRectF(r.right() + 16, y, 10, 10), colors[s])
                p.drawText(QPointF(r.right() + 30, y + 9), "C%d" % (s + 1))
        p.end()


def chart_icon(type_name, size=16):
    """the small picture of a style in the list"""
    pm = QPixmap(size, size)
    pm.fill(QColor(0, 0, 0, 0))
    p = QPainter(pm)
    ink = QColor("#202060")
    if "Pie" in type_name:
        p.setBrush(QColor("#F0D020"))
        p.setPen(QPen(ink, 1))
        p.drawEllipse(QRectF(1, 1, size - 3, size - 3))
        p.setBrush(QColor("#2040D0"))
        p.drawPie(QRectF(1, 1, size - 3, size - 3), 0, 100 * 16)
    elif "Line" in type_name:
        p.setPen(QPen(ink, 1))
        p.drawPolyline(QPolygonF([QPointF(1, 13), QPointF(5, 7), QPointF(9, 10), QPointF(14, 3)]))
    else:
        for i, h in enumerate((8, 12, 6)):
            p.fillRect(QRectF(1 + i * 5, size - 1 - h, 4, h), QColor("#C02020") if "Step" in type_name else ink)
    p.end()
    return QIcon(pm)


# ------------------------------------------------------------------------------------------------ the wizard
class DataFormWizard(QDialog):
    """the pages of the Data Form Wizard and Finish. collect: the Application Wizard's Create New Form... (Finish
    gives the settings, the Application Wizard makes the forms with its project)"""

    def __init__(self, main, collect=False, df=None):
        super().__init__(main)
        self.main = main
        self.collect = collect
        self.df = df or G.DataForm()
        self.schema = Schema(self.df)
        self.result_df = None
        self.created = None                     # the form module after Finish
        outer = QVBoxLayout(self)
        body = QHBoxLayout()
        self.picture = QLabel()
        self.picture.setFixedSize(150, 120)
        body.addWidget(self.picture, 0, Qt.AlignmentFlag.AlignTop)
        self.stack = QStackedWidget()
        body.addWidget(self.stack, 1)
        outer.addLayout(body, 1)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        outer.addWidget(line)
        row = QHBoxLayout()
        b = QPushButton("Help")
        b.clicked.connect(lambda: main.help_keyword("Data Form Wizard"))
        row.addWidget(b)
        row.addStretch(1)
        self.b_cancel = QPushButton("Cancel")
        self.b_back = QPushButton("< &Back")
        self.b_next = QPushButton("&Next >")
        self.b_finish = QPushButton("&Finish")
        for b in (self.b_cancel, self.b_back, self.b_next, self.b_finish):
            row.addWidget(b)
        self.b_cancel.clicked.connect(self.reject)
        self.b_back.clicked.connect(lambda: self.step(-1))
        self.b_next.clicked.connect(lambda: self.step(1))
        self.b_finish.clicked.connect(self.finish)
        outer.addLayout(row)
        self.pages = {}
        for key, build in (("intro", self._intro_page), ("dbtype", self._dbtype_page), ("database", self._database_page),
                           ("connect", self._connect_page), ("form", self._form_page), ("source", self._source_page),
                           ("master", self._master_page), ("detail", self._detail_page),
                           ("relation", self._relation_page), ("gridtype", self._gridtype_page),
                           ("gridstyle", self._gridstyle_page), ("columns", self._columns_page),
                           ("uioptions", self._uioptions_page), ("chartfields", self._chartfields_page),
                           ("chartstyle", self._chartstyle_page), ("chartlook", self._chartlook_page),
                           ("captions", self._captions_page), ("controls", self._controls_page),
                           ("finished", self._finished_page)):
            w = QFrame()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(6, 0, 0, 0)
            build(lay)
            self.pages[key] = self.stack.addWidget(w)
        for pk in (self.src, self.msrc, self.dsrc):
            pk.changed.connect(self.picker_changed)
        self.resize(640, 460)
        self.key = "intro"
        self.load_settings()
        self.show_page("intro")

    # ------------------------------------------------------------ the pages
    @staticmethod
    def _text(lay, text):
        t = QLabel(text)
        t.setWordWrap(True)
        lay.addWidget(t)
        return t

    def _profile_row(self, lay, label, save):
        self._text(lay, label)
        row = QHBoxLayout()
        combo = QComboBox()
        combo.addItem(NONE)
        for f in sorted(os.listdir(profiles_dir())):
            if f.lower().endswith(".rwp"):
                combo.addItem(os.path.join(profiles_dir(), f))
        row.addWidget(combo, 1)
        b = QPushButton("...")
        b.setFixedWidth(32)
        b.clicked.connect(lambda: self.pick_profile(combo, save))
        row.addWidget(b)
        lay.addLayout(row)
        return combo

    def _intro_page(self, lay):
        self._text(lay, "The Data Form Wizard makes a form whose controls show and change the records of a database: "
                        "an Access file or an ODBC data source.")
        self.load_combo = self._profile_row(lay, "From what profile do you want to load your settings?", False)
        self.load_combo.activated.connect(lambda _i: self.load_profile(self.load_combo.currentText()))
        lay.addStretch(1)

    def _dbtype_page(self, lay):
        self._text(lay, "Select a database format from the list.")
        lay.addStretch(1)
        self.dbtype_list = QListWidget()
        self.dbtype_list.addItems(G.DB_TYPES)
        self.dbtype_list.setMaximumHeight(160)
        lay.addWidget(self.dbtype_list)

    def _database_page(self, lay):
        self._text(lay, "Click Browse... to select the database file.")
        lay.addStretch(1)
        lay.addWidget(QLabel("&Database Name:"))
        self.db_name = QLineEdit()
        lay.addWidget(self.db_name)
        row = QHBoxLayout()
        row.addStretch(1)
        b = QPushButton("B&rowse...")
        b.clicked.connect(self.browse_db)
        row.addWidget(b)
        lay.addLayout(row)
        lay.addStretch(1)

    def _connect_page(self, lay):
        self._text(lay, "Give the information of the connection.")
        lay.addStretch(1)
        g = QGroupBox("ODBC Connect Data")
        gl = QGridLayout(g)
        self.odbc = {}
        for i, (k, label) in enumerate((("DSN", "&DSN:"), ("UID", "&UID:"), ("PWD", "&PWD:"), ("Database", "Data&base:"),
                                        ("Driver", "Dri&ver:"), ("Server", "&Server:"))):
            gl.addWidget(QLabel(label), i, 0)
            if k in ("DSN", "Driver"):
                w = QComboBox()
                w.setEditable(k == "Driver")
                if k == "DSN":
                    w.addItem(NONE)
                w.addItems(odbc_names("-s" if k == "DSN" else "-d"))
                if k == "Driver":
                    w.setCurrentIndex(-1)
            else:
                w = QLineEdit()
                if k == "PWD":
                    w.setEchoMode(QLineEdit.EchoMode.Password)
            self.odbc[k] = w
            gl.addWidget(w, i, 1)
        lay.addWidget(g)

    def _form_page(self, lay):
        self._text(lay, "Select the type of the form and how its controls reach the data.")
        lay.addStretch(1)
        lay.addWidget(QLabel("&What name do you want for the form?"))
        self.form_name = QLineEdit()
        lay.addWidget(self.form_name)
        lay.addStretch(1)
        row = QHBoxLayout()
        col = QVBoxLayout()
        col.addWidget(QLabel("Form Layout"))
        self.layout_list = QListWidget()
        self.layout_list.addItems(G.LAYOUTS)
        self.layout_list.setMaximumHeight(110)
        col.addWidget(self.layout_list)
        row.addLayout(col, 1)
        g = QGroupBox("Binding Type")
        gl = QVBoxLayout(g)
        self.bind_group = QButtonGroup(self)
        for i, t in enumerate(("ADO &Data Control", "ADO C&ode", "C&lass")):
            b = QRadioButton(t)
            self.bind_group.addButton(b, i)
            gl.addWidget(b)
        row.addWidget(g)
        lay.addLayout(row)
        self.layout_list.currentRowChanged.connect(self.layout_changed)

    def _source_page(self, lay):
        self._text(lay, "Select a record source, then the fields that the form shows.")
        self.src = SourcePicker(self.schema)
        lay.addWidget(self.src, 1)

    def _master_page(self, lay):
        self._text(lay, "Select the record source and the fields of the Master part of the form. The Master part "
                        "shows one record at a time.")
        self.msrc = SourcePicker(self.schema)
        lay.addWidget(self.msrc, 1)

    def _detail_page(self, lay):
        self._text(lay, "Select the record source and the fields of the Detail part of the form. The Detail part "
                        "shows, in a grid, the records that go with the record of the Master part.")
        self.dsrc = SourcePicker(self.schema)
        lay.addWidget(self.dsrc, 1)

    def _relation_page(self, lay):
        self._text(lay, "Select a field of each record source: the two fields link the sources.")
        lay.addStretch(1)
        row = QGridLayout()
        row.addWidget(QLabel("Master"), 0, 0)
        row.addWidget(QLabel("Detail"), 0, 1)
        self.rel_master = QListWidget()
        self.rel_detail = QListWidget()
        row.addWidget(self.rel_master, 1, 0)
        row.addWidget(self.rel_detail, 1, 1)
        lay.addLayout(row)

    def _gridtype_page(self, lay):
        self._text(lay, "Select the type of grid you want to use.")
        lay.addStretch(1)
        g = QGroupBox("Grid type")
        gl = QVBoxLayout(g)
        self.gridtype_group = QButtonGroup(self)
        for i, t in enumerate(("&Standard", "&Outline")):
            b = QRadioButton(t)
            self.gridtype_group.addButton(b, i)
            gl.addWidget(b)
        self.gridtype_group.idClicked.connect(lambda i: self.grid_hint.setPlainText(GRID_HINTS[i]))
        lay.addWidget(g)
        lay.addWidget(QLabel("Description:"))
        self.grid_hint = QPlainTextEdit()
        self.grid_hint.setReadOnly(True)
        self.grid_hint.setMaximumHeight(60)
        lay.addWidget(self.grid_hint)

    def _gridstyle_page(self, lay):
        self._text(lay, "Choose a grid style.")
        lay.addWidget(QLabel("&Grid style:"))
        self.gridstyle_list = QListWidget()
        self.gridstyle_list.addItems(G.GRID_STYLES)
        self.gridstyle_list.setMaximumSize(180, 80)
        self.gridstyle_list.currentRowChanged.connect(self.show_style)
        lay.addWidget(self.gridstyle_list)
        self.style_table = QTableWidget()
        self.style_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        lay.addWidget(self.style_table, 1)

    def _columns_page(self, lay):
        self._text(lay, "Drag a column by its header to change the order of the columns. Drag the border of a "
                        "header to change the width of its column; a column made as narrow as nothing is hidden.")
        self.cols_table = QTableWidget()
        self.cols_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        hh = self.cols_table.horizontalHeader()
        hh.setSectionsMovable(True)
        hh.setMinimumSectionSize(0)
        self.cols_table.verticalHeader().hide()
        lay.addWidget(self.cols_table, 1)

    def _uioptions_page(self, lay):
        self._text(lay, "Select which User Interface options you would like.")
        g = QGroupBox("Options")
        gl = QVBoxLayout(g)
        self.cb_drag = QCheckBox("Allow Column &Dragging")
        self.cb_sort = QCheckBox("Allow User &Sorting")
        gl.addWidget(self.cb_drag)
        gl.addWidget(self.cb_sort)
        lay.addWidget(g)
        self._text(lay, "Column dragging: a header dragged onto another one moves its column there. User sorting: a "
                        "double click on a header sorts the records by that column.")
        lay.addStretch(1)

    def _chartfields_page(self, lay):
        row = QHBoxLayout()
        col = QVBoxLayout()
        self._text(col, "Select the fields to group by on the X and Z axes. The Z axis is optional.")
        self.cx, self.cz, self.cy, self.cagg = QComboBox(), QComboBox(), QComboBox(), QComboBox()
        for label, w in (("&X Axis:", self.cx), ("&Z Axis:", self.cz)):
            col.addWidget(QLabel(label))
            col.addWidget(w)
        self._text(col, "Select the field to summarize on the &Y axis.")
        col.addWidget(self.cy)
        self._text(col, "Select the function that summarizes the values of the Y axis.")
        self.cagg.addItems([a for a, _f in G.AGGREGATES])
        col.addWidget(self.cagg)
        col.addStretch(1)
        row.addLayout(col, 1)
        self.preview1 = ChartPreview()
        row.addWidget(self.preview1, 1)
        lay.addLayout(row, 1)

    def _chartstyle_page(self, lay):
        row = QHBoxLayout()
        col = QVBoxLayout()
        self._text(col, "Select the style of chart you want.")
        col.addWidget(QLabel("Chart &style:"))
        self.chartstyle_list = QListWidget()
        for name, _t in G.CHART_STYLES:
            self.chartstyle_list.addItem(QListWidgetItem(chart_icon(name), name))
        self.chartstyle_list.currentRowChanged.connect(self.chart_look_changed)
        col.addWidget(self.chartstyle_list, 1)
        row.addLayout(col, 1)
        self.preview2 = ChartPreview()
        row.addWidget(self.preview2, 1)
        lay.addLayout(row, 1)

    def _chartlook_page(self, lay):
        row = QHBoxLayout()
        col = QVBoxLayout()
        self._text(col, "Select the appearance options that you want.")
        self.cb_legend = QCheckBox("Show &legend")
        self.cb_markers = QCheckBox("Show &markers")
        self.cb_stack = QCheckBox("Stac&k series")
        for c in (self.cb_legend, self.cb_markers, self.cb_stack):
            col.addWidget(c)
            c.toggled.connect(lambda _v: self.chart_look_changed())
        col.addStretch(1)
        self._text(col, "Move the slider to the degree of &rotation you want.")
        self.rot_label = QLabel()
        self.rot_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        col.addWidget(self.rot_label)
        self.rot = QSlider(Qt.Orientation.Horizontal)
        self.rot.setRange(0, 360)
        self.rot.valueChanged.connect(lambda _v: self.chart_look_changed())
        col.addWidget(self.rot)
        row.addLayout(col, 1)
        self.preview3 = ChartPreview()
        row.addWidget(self.preview3, 1)
        lay.addLayout(row, 1)

    def _captions_page(self, lay):
        row = QHBoxLayout()
        col = QVBoxLayout()
        self._text(col, "Select a caption and type its text.")
        col.addWidget(QLabel("&Titles and footnote:"))
        self.cap_list = QListWidget()
        self.cap_list.addItems([n for n, _h in G.CHART_CAPTIONS])
        self.cap_list.currentRowChanged.connect(self.caption_chosen)
        col.addWidget(self.cap_list)
        col.addWidget(QLabel("Te&xt:"))
        self.cap_text = QLineEdit()
        self.cap_text.textEdited.connect(self.caption_typed)
        col.addWidget(self.cap_text)
        col.addStretch(1)
        row.addLayout(col, 1)
        self.preview4 = ChartPreview()
        row.addWidget(self.preview4, 1)
        lay.addLayout(row, 1)

    def _controls_page(self, lay):
        self._text(lay, "Select the controls to place on the form.")
        lay.addStretch(1)
        g = QGroupBox("Available Controls")
        gl = QGridLayout(g)
        self.btn_checks = {}
        for i, b in enumerate(("Add", "Update", "Delete", "Refresh", "Close")):
            c = QCheckBox("&%s Button" % b)
            self.btn_checks[b] = c
            gl.addWidget(c, i, 0)
        self.cb_datacontrol = QCheckBox("Show Data &Control")
        gl.addWidget(self.cb_datacontrol, 5, 0)
        b = QPushButton("Select &All")
        b.clicked.connect(lambda: self.all_controls(True))
        gl.addWidget(b, 0, 1)
        b = QPushButton("Clea&r All")
        b.clicked.connect(lambda: self.all_controls(False))
        gl.addWidget(b, 1, 1)
        lay.addWidget(g)

    def _finished_page(self, lay):
        self._text(lay, "The Data Form Wizard has all the information it needs.\n\nClick Finish to make the form.")
        self.save_combo = self._profile_row(lay, "To what profile do you want to save your settings?", True)
        lay.addStretch(1)

    # ------------------------------------------------------------ settings <-> pages
    def load_settings(self):
        """the controls of the pages from df (a new wizard, a profile)"""
        df = self.df
        self.dbtype_list.setCurrentRow(df.dbtype)
        self.db_name.setText(df.database)
        for k, w in self.odbc.items():
            v = df.odbc.get(k, "")
            if isinstance(w, QComboBox):
                if w.isEditable():
                    w.setEditText(v)
                else:
                    i = w.findText(v)
                    w.setCurrentIndex(i if i >= 0 else 0)
            else:
                w.setText(v)
        self.form_name.setText(df.name)
        self.layout_list.setCurrentRow(df.layout)
        self.layout_changed(df.layout)
        self.bind_group.button(df.binding).setChecked(True)
        for b, c in self.btn_checks.items():
            c.setChecked(df.buttons.get(b, True))
        self.cb_datacontrol.setChecked(df.show_datacontrol)
        self.gridtype_group.button(1 if df.grid_outline else 0).setChecked(True)
        self.grid_hint.setPlainText(GRID_HINTS[1 if df.grid_outline else 0])
        self.gridstyle_list.setCurrentRow(df.grid_style)
        self.cb_drag.setChecked(df.allow_dragging)
        self.cb_sort.setChecked(df.allow_sorting)
        self.cagg.setCurrentIndex(df.chart_agg)
        self.chartstyle_list.setCurrentRow(df.chart_style)
        self.cb_legend.setChecked(df.show_legend)
        self.cb_markers.setChecked(df.show_markers)
        self.cb_stack.setChecked(df.stack_series)
        self.rot.setValue(int(df.rotation))
        self.cap_list.setCurrentRow(0)

    def read_page(self, key):
        """the choices of a page into df"""
        df = self.df
        if key == "dbtype":
            df.dbtype = max(0, self.dbtype_list.currentRow())
        elif key == "database":
            df.database = self.db_name.text().strip()
        elif key == "connect":
            for k, w in self.odbc.items():
                v = (w.currentText() if isinstance(w, QComboBox) else w.text()).strip()
                df.odbc[k] = "" if v == NONE else v
        elif key == "form":
            df.name = self.form_name.text().strip()
            df.layout = max(0, self.layout_list.currentRow())
            b = self.bind_group.checkedId()
            df.binding = b if b in G.DataForm.bindings_allowed(df.layout) else G.DataForm.bindings_allowed(df.layout)[0]
        elif key in ("source", "master"):
            pk = self.src if key == "source" else self.msrc
            df.source, df.fields, df.sort = pk.source.currentText(), pk.fields(), pk.sort_field()
        elif key == "detail":
            df.detail_source, df.detail_fields, df.detail_sort = (self.dsrc.source.currentText(), self.dsrc.fields(),
                                                                  self.dsrc.sort_field())
        elif key == "relation":
            m, d = self.rel_master.currentItem(), self.rel_detail.currentItem()
            df.master_key = m.text() if m else ""
            df.detail_key = d.text() if d else ""
        elif key == "gridtype":
            df.grid_outline = self.gridtype_group.checkedId() == 1
        elif key == "gridstyle":
            df.grid_style = max(0, self.gridstyle_list.currentRow())
        elif key == "columns":
            hh = self.cols_table.horizontalHeader()
            order = sorted(range(self.cols_table.columnCount()), key=hh.visualIndex)
            names = [self.cols_table.horizontalHeaderItem(i).text() for i in order]
            df.col_widths = {self.cols_table.horizontalHeaderItem(i).text(): hh.sectionSize(i) * 15 for i in order}
            if sorted(names) == sorted(df.fields):
                df.fields = names
        elif key == "uioptions":
            df.allow_dragging, df.allow_sorting = self.cb_drag.isChecked(), self.cb_sort.isChecked()
        elif key == "chartfields":
            df.chart_x = self.cx.currentText()
            df.chart_z = "" if self.cz.currentText() == NONE else self.cz.currentText()
            df.chart_y = self.cy.currentText()
            df.chart_agg = max(0, self.cagg.currentIndex())
        elif key == "chartstyle":
            df.chart_style = max(0, self.chartstyle_list.currentRow())
        elif key == "chartlook":
            df.show_legend, df.show_markers = self.cb_legend.isChecked(), self.cb_markers.isChecked()
            df.stack_series, df.rotation = self.cb_stack.isChecked(), self.rot.value()
        elif key == "controls":
            for b, c in self.btn_checks.items():
                df.buttons[b] = c.isChecked()
            df.show_datacontrol = self.cb_datacontrol.isChecked()

    def enter_page(self, key):
        """what a page shows from the choices before it"""
        df = self.df
        try:
            if key == "source":
                self.src.fill_sources(self.schema.sources(), df.source)
                self.src.set_state(df.source, df.fields or None, df.sort)
            elif key == "master":
                self.msrc.fill_sources(self.schema.sources(), df.source)
                self.msrc.set_state(df.source, df.fields or None, df.sort)
            elif key == "detail":
                self.dsrc.fill_sources(self.schema.sources(), df.detail_source)
                self.dsrc.set_state(df.detail_source, df.detail_fields or None, df.detail_sort)
            elif key == "relation":
                for lw, src, cur in ((self.rel_master, df.source, df.master_key),
                                     (self.rel_detail, df.detail_source, df.detail_key)):
                    names = self.schema.fields(src)
                    lw.clear()
                    lw.addItems(names)
                    if cur in names:
                        lw.setCurrentRow(names.index(cur))
                if not (df.master_key or df.detail_key):     # the same name on both sides: chosen
                    common = [f for f in self.schema.fields(df.source) if f in self.schema.fields(df.detail_source)]
                    if common:
                        self.rel_master.setCurrentRow([self.rel_master.item(i).text()
                                                       for i in range(self.rel_master.count())].index(common[0]))
                        self.rel_detail.setCurrentRow([self.rel_detail.item(i).text()
                                                       for i in range(self.rel_detail.count())].index(common[0]))
            elif key == "chartfields":
                names = df.fields
                for w, cur, none in ((self.cx, df.chart_x, False), (self.cz, df.chart_z, True), (self.cy, df.chart_y, False)):
                    w.clear()
                    if none:
                        w.addItem(NONE)
                    w.addItems(names)
                    i = w.findText(cur) if cur else (0 if none else -1)
                    w.setCurrentIndex(i if i >= 0 else 0)
                if not df.chart_y and len(names) > 1:
                    self.cy.setCurrentIndex(1)
            elif key == "gridstyle":
                style_grid(self.style_table, max(0, self.gridstyle_list.currentRow()), df.fields)
            elif key == "columns":
                self.cols_table.clear()
                self.cols_table.setColumnCount(len(df.fields))
                self.cols_table.setRowCount(4)
                self.cols_table.setHorizontalHeaderLabels(df.fields)
                for i, f in enumerate(df.fields):
                    if f in df.col_widths:
                        self.cols_table.setColumnWidth(i, df.col_widths[f] // 15)
            elif key == "controls":
                self.cb_datacontrol.setEnabled(df.binding == G.B_ADODC)
            if key in ("chartfields", "chartstyle", "chartlook", "captions"):
                self.chart_look_changed()
        except dbinfo.AdoError as e:
            QMessageBox.warning(self, TITLE, str(e))

    def check_page(self, key):
        """None, or why Next cannot leave the page"""
        df = self.df
        if key == "database":
            if not df.database or not os.path.isfile(df.database):
                return "Select a database file that exists."
            if not self.schema.sources():
                return "No tables or queries were found in this database."
        elif key == "connect":
            try:
                if not self.schema.sources():
                    return "No tables or views were found with this connection."
            except dbinfo.AdoError as e:
                return "The connection failed:\n" + str(e)
        elif key == "form":
            if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,39}", df.name or ""):
                return "Type a name for the form: a letter, then letters, digits or _."
            p = getattr(self.main, "project", None)
            if not self.collect and p is not None and p.module(df.name) is not None:
                return "There is already a module with the name %s in the project." % df.name
        elif key in ("source", "master"):
            if not df.source or not df.fields:
                return "Select a record source and at least one field."
            if key == "source" and df.layout == G.L_CHART and len(df.fields) < 2:
                return "Select at least two fields: one to group by and one to summarize."
        elif key == "detail":
            if not df.detail_source or not df.detail_fields:
                return "Select a record source and at least one field."
        elif key == "relation":
            if not (df.master_key and df.detail_key):
                return "Select a field of each record source."
        elif key == "chartfields":
            axes = [f for f in (df.chart_x, df.chart_z, df.chart_y) if f]
            if len(set(axes)) != len(axes):
                return "A field can be on one axis only. Select a different field for each axis."
        return None

    def show_page(self, key):
        self.key = key
        keys = page_keys(self.df)
        self.stack.setCurrentIndex(self.pages[key])
        self.setWindowTitle("%s - %s" % (TITLE, PAGE_TITLES[key]))
        pic = {"intro": "intro", "dbtype": "dbs", "database": "dbfile", "connect": "connect",
               "form": "form%d" % self.df.layout, "source": "source", "master": "source", "detail": "source",
               "relation": "relation", "gridtype": "form%d" % G.L_FLEX, "controls": "controls",
               "finished": "finished"}.get(key)
        self.picture.setVisible(key not in NO_PICTURE)
        if pic:
            self.picture.setPixmap(df_picture(pic))
        i = keys.index(key)
        self.b_back.setEnabled(i > 0)
        self.b_next.setEnabled(i < len(keys) - 1)
        self.update_finish()
        (self.b_finish if i == len(keys) - 1 else self.b_next).setDefault(True)
        self.enter_page(key)

    def update_finish(self):
        """Finish: when the form can be made (VB6 df2-02: from the Record Source page on)"""
        keys = page_keys(self.df)
        after = keys.index(self.key) >= keys.index("form") if self.key in keys else False
        self.b_finish.setEnabled(after and G.check(self.df) is None)

    def step(self, d):
        keys = page_keys(self.df)
        self.read_page(self.key)
        if d > 0:
            why = self.check_page(self.key)
            if why:
                QMessageBox.warning(self, TITLE, why)
                return
        keys = page_keys(self.df)               # the layout may have changed the pages
        i = keys.index(self.key) if self.key in keys else 0
        self.show_page(keys[max(0, min(len(keys) - 1, i + d))])

    def go_to(self, key):
        """the tests: straight to a page (the pages before it read)"""
        self.read_page(self.key)
        self.show_page(key)

    # ------------------------------------------------------------ events of the pages
    def picker_changed(self):
        self.read_page(self.key)
        self.update_finish()

    def browse_db(self):
        start = os.path.dirname(self.db_name.text()) or os.path.expanduser("~")
        path, _ = QFileDialog.getOpenFileName(self, "Access Database", start,
                                              "Access Databases (*.mdb *.MDB);;All Files (*)")
        if path:
            self.db_name.setText(path)

    def layout_changed(self, row):
        allowed = G.DataForm.bindings_allowed(max(0, row))
        for i in range(3):
            self.bind_group.button(i).setEnabled(i in allowed)
        if self.bind_group.checkedId() not in allowed:
            self.bind_group.button(allowed[0]).setChecked(True)
        if self.key == "form":
            self.picture.setPixmap(df_picture("form%d" % max(0, row)))

    def show_style(self, row):
        if row >= 0:
            style_grid(self.style_table, row, self.df.fields)

    def chart_look_changed(self, *_a):
        st = max(0, self.chartstyle_list.currentRow())
        self.rot_label.setText("%d°" % self.rot.value())
        for pv in (self.preview1, self.preview2, self.preview3, self.preview4):
            pv.style, pv.legend, pv.markers = st, self.cb_legend.isChecked(), self.cb_markers.isChecked()
            pv.stack, pv.rotation = self.cb_stack.isChecked(), self.rot.value()
            pv.update()

    def caption_chosen(self, row):
        if 0 <= row < len(G.CHART_CAPTIONS):
            self.cap_text.setText(self.df.captions.get(G.CHART_CAPTIONS[row][0], ""))

    def caption_typed(self, text):
        row = self.cap_list.currentRow()
        if 0 <= row < len(G.CHART_CAPTIONS):
            self.df.captions[G.CHART_CAPTIONS[row][0]] = text

    def all_controls(self, on):
        for c in self.btn_checks.values():
            c.setChecked(on)
        if self.cb_datacontrol.isEnabled():
            self.cb_datacontrol.setChecked(on)

    # ------------------------------------------------------------ profiles
    def pick_profile(self, combo, save):
        if save:
            path, _ = QFileDialog.getSaveFileName(self, "Save Profile", os.path.join(profiles_dir(), "dataform.rwp"),
                                                  PROFILE_FILTER)
            if path and not path.lower().endswith(".rwp"):
                path += ".rwp"
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Open Profile", profiles_dir(), PROFILE_FILTER)
        if not path:
            return
        if combo.findText(path) < 0:
            combo.addItem(path)
        combo.setCurrentIndex(combo.findText(path))
        if not save:
            self.load_profile(path)

    def load_profile(self, path):
        if not path or path == NONE:
            return
        try:
            G.load_profile(path, self.df)
        except (OSError, ValueError) as e:
            QMessageBox.warning(self, "Open Profile", "%s\n%s" % (os.path.basename(path), getattr(e, "strerror", None) or e))
            return
        self.load_settings()

    # ------------------------------------------------------------ Finish
    def finish(self):
        self.read_page(self.key)
        why = G.check(self.df)
        if why:
            QMessageBox.warning(self, TITLE, why)
            return
        path = self.save_combo.currentText()
        if path and path != NONE:
            try:
                G.save_profile(self.df, path)
            except OSError as e:
                QMessageBox.warning(self, "Save Profile", "%s\n%s" % (path, e.strerror))
        if self.collect:
            self.result_df = self.df
            self.accept()
            return
        self.created = add_to_project(self.main, self.df)
        if self.created is None:
            return
        self.accept()
        if not self.main.S["dfwiz_no_created"]:
            CreatedDialog(self.main, self.df.name).exec()


def add_to_project(main, df, project=None, open_it=True):
    """the modules of a data form into the project (not saved, as VB6), the references and components it needs"""
    from .vbmodel import Module, parse_source
    p = project or main.project
    if p is None:
        QMessageBox.warning(main, TITLE, "There is no project to add the form to.")
        return None
    if p.module(df.name) is not None:            # a module of the form's name already (the Application Wizard's)
        base, n = df.name, 1
        while p.module(df.name) is not None:
            n += 1
            df.name = "%s%d" % (base, n)
    if df.binding == G.B_CLASS:                 # a module of the class's name already: a number after it
        df.class_override = ""
        base, n = df.class_name(), 1
        while p.module(df.class_name()) is not None:
            n += 1
            df.class_override = "%s%d" % (base, n)
    out = G.generate(df)
    created = None
    for kind, name, text in out.modules:
        ext = ".frm" if kind == "Form" else ".cls"
        sf = parse_source(text.replace("\n", "\r\n"), None, kind, True, p.encoding)
        sf.encoding = p.encoding
        m = Module(p, kind, name, name + ext, sf)
        m.dirty = True
        p.modules.append(m)
        p._vbp_add(m)
        if kind == "Form":
            created = m
    p.add_references(out.references)
    have = p.components()
    p.set_components(have + [k for k in out.libraries if k not in have])
    p.dirty = True
    if main is not None and project is None:
        main.fill_tree()
        main.update_title()
        if open_it and created is not None:
            main.open_designer(created)
    return created


class CreatedDialog(QDialog):
    """after Finish (VB6 df1-17): the form has been made, "Don't show this dialog in the future." """

    def __init__(self, main, name):
        super().__init__(main)
        self.main = main
        self.setWindowTitle("Data Form Created")
        lay = QVBoxLayout(self)
        row = QHBoxLayout()
        pic = QLabel()
        pic.setPixmap(wizard_picture("intro", bg=None).copy(26, 10, 100, 100).scaled(
            48, 48, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        row.addWidget(pic, 0, Qt.AlignmentFlag.AlignTop)
        row.addWidget(QLabel("The form '%s' has been made and added to your project." % name), 1)
        lay.addLayout(row)
        self.again = QCheckBox("Don't show this dialog in the future.")
        lay.addWidget(self.again)
        b = QPushButton("OK")
        b.setDefault(True)
        b.clicked.connect(self.accept)
        r = QHBoxLayout()
        r.addStretch(1)
        r.addWidget(b)
        lay.addLayout(r)

    def accept(self):
        if self.again.isChecked():
            self.main.S["dfwiz_no_created"] = True
            self.main.S.save()
        super().accept()


def open_wizard(main, collect=False):
    """Add-Ins > Data Form Wizard...; collect: the Application Wizard's Create New Form... (the settings)"""
    dlg = DataFormWizard(main, collect)
    main._dfwizard = dlg                        # the tests reach it
    dlg.exec()
    return dlg.result_df if collect else dlg.created


__all__ = ["DataFormWizard", "open_wizard", "add_to_project", "dfwizard_icon", "page_keys"]
