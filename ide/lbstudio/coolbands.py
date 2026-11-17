"""coolbands.py - the bands of a CoolBar (ComCtl3) at design time (phase 7.4): what the "Bands" property page of
the CoolBar does in VB6. Each band has a Caption, a Key, a Child (a control inside the CoolBar), a Width, a MinWidth
and a MinHeight (twips), NewRow, Visible, a Style (0 - cc3BandNormal, 1 - cc3BandFixedSize), AllowVertical, an Image
(of the ImageList), a Tag and the colours (UseCoolbarColors, BackColor, ForeColor). They are written in the .frm as
the VB6 IDE writes them: BandCount, then Caption1, Child1, MinWidth1 ... BandStyle1, BandTag1 (vbprops.cool_bands /
set_cool_bands); the children then take their places in their bands as the runtime puts them (vbprops.cool_place)."""

from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QPushButton,
                             QComboBox, QHeaderView, QLabel, QAbstractItemView, QCheckBox)

HEADS = ["Caption", "Key", "Child", "Width", "MinWidth", "MinHeight", "NewRow", "Visible", "Style", "AllowVertical",
         "Image", "Tag"]
STYLES = [(0, "0 - cc3BandNormal"), (1, "1 - cc3BandFixedSize")]
CHECKS = {6: "NewRow", 7: "Visible", 9: "AllowVertical"}
TEXTS = {0: "Caption", 1: "Key", 10: "Image", 11: "Tag"}  # Image: an index or a key of the ImageList
NUMS = {3: "Width", 4: "MinWidth", 5: "MinHeight"}


class CoolBandsDialog(QDialog):
    """bands: [dict] (vbprops.cool_bands); children: the names of the controls inside the CoolBar"""

    def __init__(self, parent, name, bands, children):
        super().__init__(parent)
        self.setWindowTitle(f"{name} - Bands")
        self.children = [""] + list(children or [])
        self.extra = []                         # the properties of each band that the table does not show
        v = QVBoxLayout(self)
        v.addWidget(QLabel("The bands of the CoolBar, in their order. Child: a control inside the CoolBar; sizes in "
                           "twips."))
        self.table = QTableWidget(0, len(HEADS))
        self.table.setHorizontalHeaderLabels(HEADS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        v.addWidget(self.table, 1)
        row = QHBoxLayout()
        for text, fn in (("Insert Band", self.insert), ("Remove Band", self.delete), ("Up", lambda: self.move(-1)),
                         ("Down", lambda: self.move(1))):
            b = QPushButton(text)
            b.clicked.connect(fn)
            row.addWidget(b)
        row.addStretch(1)
        ok, cancel = QPushButton("OK"), QPushButton("Cancel")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        row.addWidget(ok)
        row.addWidget(cancel)
        v.addLayout(row)
        self._fill(bands)
        self.resize(980, 340)

    def _add_row(self, r, b):
        from .vbprops import BAND_DEFAULT
        b = dict(BAND_DEFAULT, **(b or {}))
        self.table.insertRow(r)
        self.extra.insert(r, b)
        for k, key in TEXTS.items():
            self.table.setItem(r, k, QTableWidgetItem(str(b.get(key) or "")))
        for k, key in NUMS.items():
            x = b.get(key, 0) or 0
            self.table.setItem(r, k, QTableWidgetItem(str(int(x) if float(x) == int(x) else x)))
        cb = QComboBox()
        cb.addItems(self.children)
        child = str(b.get("Child") or "")
        if child and child not in self.children:
            cb.addItem(child)
        cb.setCurrentText(child)
        self.table.setCellWidget(r, 2, cb)
        st = QComboBox()
        for n, t in STYLES:
            st.addItem(t, n)
        st.setCurrentIndex(1 if int(b.get("Style", 0) or 0) == 1 else 0)
        self.table.setCellWidget(r, 8, st)
        for k, key in CHECKS.items():
            ch = QCheckBox()
            ch.setChecked(bool(b.get(key)))
            self.table.setCellWidget(r, k, ch)

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
            del self.extra[r]

    def move(self, d):
        r = self.table.currentRow()
        if r < 0 or not 0 <= r + d < self.table.rowCount():
            return
        bands = self.bands()
        bands[r], bands[r + d] = bands[r + d], bands[r]
        self._fill(bands)
        self.table.selectRow(r + d)

    def _fill(self, bands):
        self.table.setRowCount(0)
        self.extra = []
        for b in bands:
            self._add_row(self.table.rowCount(), b)

    def bands(self):
        out = []
        for r in range(self.table.rowCount()):
            b = dict(self.extra[r])
            for k, key in TEXTS.items():
                b[key] = self._text(r, k)
            for k, key in NUMS.items():
                t = self._text(r, k).strip().replace(",", ".")
                try:
                    x = float(t) if t else 0
                except ValueError:
                    x = b.get(key, 0) or 0
                x = max(0, x)
                if key in ("Width", "MinWidth", "MinHeight"):
                    x = int(x) // 15 * 15            # whole pixels, as the CoolBar keeps them (f_cool: 2000 -> 1995)
                b[key] = int(x) if float(x) == int(x) else x
            img = b["Image"].strip()
            b["Image"] = int(img) if img.lstrip("-").isdigit() else (img or 0)
            b["Child"] = self.table.cellWidget(r, 2).currentText()
            b["Style"] = self.table.cellWidget(r, 8).currentData()
            for k, key in CHECKS.items():
                b[key] = self.table.cellWidget(r, k).isChecked()
            out.append(b)
        return out

    def chosen(self):
        return self.bands()
