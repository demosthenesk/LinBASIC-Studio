"""dgcolumns.py - the columns of a DataGrid (MSDataGridLib) at design time (phase 7.7): what the property pages of the
DataGrid in VB6 (Columns, Format, Layout) and its "Retrieve Fields" do. Each column has a Caption, a DataField, a
DataFormat (the Format page: dfmtdialog.py), a Width, an Alignment, Locked and Visible. They are written in the .frm as
VB6 writes them: ColumnCount, BeginProperty ColumnNN (DataField, Caption, BeginProperty DataFormat) and the layout in
BeginProperty Split0 / ColumnNN (Alignment, ColumnWidth, Locked, Object.Visible) - vbprops.grid_columns / set_grid_columns.
At run time the grid has these columns until it is bound to its data source; then one column for each field, unless
the program called HoldFields (VB6 oracles f_dfmt2 .. f_dfmt4)."""

from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QPushButton,
                             QComboBox, QHeaderView, QLabel, QAbstractItemView, QMessageBox, QWidget, QToolButton,
                             QLineEdit, QCheckBox)

ALIGNMENTS = [(0, "0 - dbgLeft"), (1, "1 - dbgRight"), (2, "2 - dbgCenter"), (3, "3 - dbgGeneral")]
HEADS = ["Caption", "DataField", "Width", "Alignment", "Locked", "Visible", "DataFormat"]


def _fmt_text(v):
    if not v:
        return ""
    from .dfmtdialog import category_of
    f = str(v.get("Format", "") or "")
    return category_of(v) + (f" ({f})" if f else "")


class GridColumnsDialog(QDialog):
    """cols: [{'Caption', 'DataField', 'Width' (twips or None), 'Alignment', 'Locked', 'Visible', 'DataFormat' (dict or
    None)}]; fields: the fields of the data source (Retrieve Fields)"""

    def __init__(self, parent, name, cols, fields):
        super().__init__(parent)
        self.setWindowTitle(f"{name} - Columns")
        self.fields = list(fields or [])
        self.formats = []                       # the DataFormat of each row
        v = QVBoxLayout(self)
        v.addWidget(QLabel("The columns of the design: the grid shows them until it is bound to its data source "
                           "(HoldFields keeps them)."))
        self.table = QTableWidget(0, len(HEADS))
        self.table.setHorizontalHeaderLabels(HEADS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        v.addWidget(self.table, 1)
        row = QHBoxLayout()
        for text, fn in (("Insert", self.insert), ("Delete", self.delete), ("Up", lambda: self.move(-1)),
                         ("Down", lambda: self.move(1)), ("Retrieve Fields", self.retrieve), ("Clear", self.clear)):
            b = QPushButton(text)
            b.clicked.connect(fn)
            if text == "Retrieve Fields":
                b.setEnabled(bool(self.fields))
                b.setToolTip("one column for each field of the data source (DataSource)")
            row.addWidget(b)
        row.addStretch(1)
        ok, cancel = QPushButton("OK"), QPushButton("Cancel")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        row.addWidget(ok)
        row.addWidget(cancel)
        v.addLayout(row)
        self._fill(cols)
        self.resize(820, 360)

    def _add_row(self, r, c):
        self.table.insertRow(r)
        self.formats.insert(r, dict(c["DataFormat"]) if c.get("DataFormat") else None)
        w = c.get("Width")
        for k, x in enumerate((c.get("Caption", ""), c.get("DataField", ""), "" if w is None else str(w))):
            self.table.setItem(r, k, QTableWidgetItem(str(x or "")))
        cb = QComboBox()
        for n, t in ALIGNMENTS:
            cb.addItem(t, n)
        cb.setCurrentIndex(max(0, min(3, int(c.get("Alignment", 0) or 0))))
        self.table.setCellWidget(r, 3, cb)
        for k, key, dflt in ((4, "Locked", False), (5, "Visible", True)):
            ch = QCheckBox()
            ch.setChecked(bool(c.get(key, dflt)))
            self.table.setCellWidget(r, k, ch)
        self._format_cell(r)

    def _format_cell(self, r):
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        lab = QLineEdit(_fmt_text(self.formats[r]))
        lab.setReadOnly(True)
        lab.setFrame(False)
        h.addWidget(lab, 1)
        b = QToolButton()
        b.setText("...")
        b.clicked.connect(lambda _=False, w=w: self._format_dialog(w))
        h.addWidget(b)
        self.table.setCellWidget(r, 6, w)

    def _format_dialog(self, w):
        for r in range(self.table.rowCount()):
            if self.table.cellWidget(r, 6) is w:
                from .dfmtdialog import DataFormatDialog
                dlg = DataFormatDialog(self, self._text(r, 0) or f"Column{r:02d}", self.formats[r])
                if dlg.exec():
                    self.set_format(r, dlg.chosen())
                return

    def set_format(self, r, value):
        self.formats[r] = value
        self._format_cell(r)

    def _text(self, r, k):
        it = self.table.item(r, k)
        return it.text() if it is not None else ""

    def insert(self):
        r = self.table.currentRow()
        r = self.table.rowCount() if r < 0 else r + 1
        self._add_row(r, {})
        self.table.selectRow(r)

    def delete(self):
        r = self.table.currentRow()
        if r >= 0:
            self.table.removeRow(r)
            del self.formats[r]

    def move(self, d):
        r = self.table.currentRow()
        if r < 0 or not 0 <= r + d < self.table.rowCount():
            return
        cols = self.columns()
        cols[r], cols[r + d] = cols[r + d], cols[r]
        self._fill(cols)
        self.table.selectRow(r + d)

    def retrieve(self):
        if self.table.rowCount() and QMessageBox.question(self, "Retrieve Fields",
                                                           "Replace the columns with the fields of the data source?") \
                != QMessageBox.StandardButton.Yes:
            return
        self._fill([{"Caption": f, "DataField": f} for f in self.fields])

    def clear(self):
        self._fill([])

    def _fill(self, cols):
        self.table.setRowCount(0)
        self.formats = []
        for c in cols:
            self._add_row(self.table.rowCount(), c)

    def columns(self):
        out = []
        for r in range(self.table.rowCount()):
            t = self._text(r, 2).strip().replace(",", ".")
            try:
                width = float(t) if t else None
            except ValueError:
                width = None
            out.append({"Caption": self._text(r, 0), "DataField": self._text(r, 1), "Width": width,
                        "Alignment": self.table.cellWidget(r, 3).currentData(),
                        "Locked": self.table.cellWidget(r, 4).isChecked(),
                        "Visible": self.table.cellWidget(r, 5).isChecked(),
                        "DataFormat": self.formats[r]})
        return out

    def chosen(self):
        return self.columns()
