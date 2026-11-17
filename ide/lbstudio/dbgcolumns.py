"""dbgcolumns.py - the columns of a DBGrid at design time (phase 7.6), in one dialog: what the property
pages of the DBGrid in VB6 (Columns, Layout) and its "Retrieve Fields" do. Each column has a Caption, a
DataField, a Width, an Alignment, a NumberFormat and a DefaultValue; Retrieve Fields makes one column for
each field of the Data control the grid is bound to; "Default layout" gives the two blank columns of a
new grid, whose columns come from the data source at run time (_LayoutType 0). The columns are kept in the
OleObjectBlob of the .frx (lbvc/project/oleblob.py: get_columns, set_columns)."""

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QPushButton,
                             QComboBox, QHeaderView, QLabel, QAbstractItemView, QMessageBox)

ALIGNMENTS = [(0, "0 - dbgLeft"), (1, "1 - dbgRight"), (2, "2 - dbgCenter"), (3, "3 - dbgGeneral")]
HEADS = ["Caption", "DataField", "Width", "Alignment", "NumberFormat", "DefaultValue"]
TWIPS_DEFAULT = 1350


def him_to_twips(him):
    return int(round(int(him) * 1440 / 2540))


def twips_to_him(tw):
    return int(round(int(tw) * 2540 / 1440))


class ColumnsDialog(QDialog):
    """cols: [{'Caption', 'DataField', 'DefaultValue', 'NumberFormat', 'Width' (HIMETRIC), 'Alignment'}];
    layout: the _LayoutType (0 the default layout); fields: the fields of the data source (Retrieve Fields)"""

    def __init__(self, parent, name, cols, layout, fields):
        super().__init__(parent)
        self.setWindowTitle(f"{name} - Columns")
        self.fields = list(fields or [])
        self.default_layout = not layout
        v = QVBoxLayout(self)
        self.info = QLabel()
        v.addWidget(self.info)
        self.table = QTableWidget(0, len(HEADS))
        self.table.setHorizontalHeaderLabels(HEADS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        v.addWidget(self.table, 1)
        row = QHBoxLayout()
        for text, fn in (("Insert", self.insert), ("Delete", self.delete), ("Up", lambda: self.move(-1)),
                         ("Down", lambda: self.move(1)), ("Retrieve Fields", self.retrieve),
                         ("Default layout", self.default)):
            b = QPushButton(text)
            b.clicked.connect(fn)
            if text == "Retrieve Fields":
                b.setEnabled(bool(self.fields))
                b.setToolTip("one column for each field of the Data control (DataSource)")
            row.addWidget(b)
        row.addStretch(1)
        ok, cancel = QPushButton("OK"), QPushButton("Cancel")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        row.addWidget(ok)
        row.addWidget(cancel)
        v.addLayout(row)
        for c in cols:
            self._add_row(self.table.rowCount(), c)
        self.table.itemChanged.connect(self._edited)
        self._show_info()
        self.resize(720, 360)

    def _show_info(self):
        self.info.setText("The default layout: at run time the columns are the fields of the data source."
                          if self.default_layout else
                          "A layout of its own: the grid keeps these columns (ReBind too).")

    def _add_row(self, r, c):
        blocked = self.table.blockSignals(True)
        self.table.insertRow(r)
        vals = [c.get("Caption", ""), c.get("DataField", ""), str(him_to_twips(c.get("Width") or twips_to_him(TWIPS_DEFAULT))),
                None, c.get("NumberFormat", ""), c.get("DefaultValue", "")]
        for k, x in enumerate(vals):
            if k == 3:
                cb = QComboBox()
                for n, t in ALIGNMENTS:
                    cb.addItem(t, n)
                cb.setCurrentIndex(int(c.get("Alignment", 3)) if str(c.get("Alignment", 3)).isdigit() else 3)
                cb.currentIndexChanged.connect(self._edited)
                self.table.setCellWidget(r, k, cb)
            else:
                self.table.setItem(r, k, QTableWidgetItem(str(x or "")))
        self.table.blockSignals(blocked)

    def _edited(self, *_):
        if self.default_layout:
            self.default_layout = False
            self._show_info()

    def insert(self):
        r = self.table.currentRow()
        r = self.table.rowCount() if r < 0 else r + 1
        self._add_row(r, {"Alignment": 3})
        self.table.selectRow(r)
        self._edited()

    def delete(self):
        r = self.table.currentRow()
        if r >= 0:
            self.table.removeRow(r)
            self._edited()

    def move(self, d):
        r = self.table.currentRow()
        if r < 0 or not 0 <= r + d < self.table.rowCount():
            return
        cols = self.columns()
        cols[r], cols[r + d] = cols[r + d], cols[r]
        self._fill(cols)
        self.table.selectRow(r + d)
        self._edited()

    def retrieve(self):
        if self.table.rowCount() and QMessageBox.question(self, "Retrieve Fields",
                                                           "Replace the columns with the fields of the data source?") \
                != QMessageBox.StandardButton.Yes:
            return
        self._fill([{"Caption": f, "DataField": f, "Alignment": 3} for f in self.fields])
        self._edited()

    def default(self):
        self._fill([{"Alignment": 3}, {"Alignment": 3}])
        self.default_layout = True
        self._show_info()

    def _fill(self, cols):
        self.table.setRowCount(0)
        for c in cols:
            self._add_row(self.table.rowCount(), c)

    def columns(self):
        out = []
        for r in range(self.table.rowCount()):
            def t(k):
                it = self.table.item(r, k)
                return it.text() if it is not None else ""
            try:
                tw = max(0, int(float(t(2) or TWIPS_DEFAULT)))
            except ValueError:
                tw = TWIPS_DEFAULT
            cb = self.table.cellWidget(r, 3)
            out.append({"Caption": t(0), "DataField": t(1), "Width": twips_to_him(tw),
                        "Alignment": cb.currentData() if cb is not None else 3, "NumberFormat": t(4),
                        "DefaultValue": t(5)})
        return out

    def chosen(self):
        """(columns, _LayoutType)"""
        if self.default_layout or not self.table.rowCount():
            return [{"Alignment": 3}, {"Alignment": 3}], 0
        return self.columns(), 4
