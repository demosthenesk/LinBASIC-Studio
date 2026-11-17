"""The DataRepeater at design time (phase 7.5): the user controls it can repeat (RepeatedControlName: the public user
controls of the ActiveX control projects next to the project and of the registry of components) and the page of its
RepeaterBindings (PropertyName, DataField, Key), as the RepeaterBindings tab of the VB6 property pages.
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
import glob
import os
import re

from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QPushButton,
                             QDialogButtonBox, QLabel, QHeaderView, QComboBox)


def _vbp_controls(vbp):
    """(project name, {control name: .ctl path}) of an ActiveX control project, or (None, {})"""
    try:
        text = open(vbp, "rb").read().decode("cp1252", "replace")
    except OSError:
        return None, {}
    if not re.search(r"(?im)^\s*Type\s*=\s*Control\s*$", text):
        return None, {}
    m = re.search(r'(?im)^\s*Name\s*=\s*"?([^"\r\n]+)"?\s*$', text)
    name = m.group(1).strip() if m else os.path.splitext(os.path.basename(vbp))[0]
    out = {}
    for rel in re.findall(r"(?im)^\s*UserControl\s*=\s*(.+?)\s*$", text):
        path = os.path.join(os.path.dirname(vbp), rel.strip().strip('"').replace("\\", "/"))
        out[os.path.splitext(os.path.basename(path))[0]] = path
    return name, out


def repeated_controls(folder):
    """{"Proj.Ctl": the .ctl (or None)}: what RepeatedControlName can name"""
    out = {}
    if folder and os.path.isdir(folder):
        for v in sorted(glob.glob(os.path.join(folder, "*.[vV][bB][pP]"))):
            name, ctls = _vbp_controls(v)
            for c, path in ctls.items():
                out[f"{name}.{c}"] = path
    try:
        from lbvc.driver import load_registry
        for e in load_registry().values():
            if str(e.get("type", "")).lower() != "control":
                continue
            name, ctls = _vbp_controls(e.get("vbp", ""))
            for c in e.get("controls", []):
                out.setdefault(f"{e.get('name', name)}.{c}", ctls.get(c))
    except Exception:                           # noqa: BLE001 - no registry: the controls next to the project only
        pass
    return out


def bindable_props(ctl_path):
    """the properties of a user control marked data bound (Attribute X.VB_MemberFlags with "bindable" 0x4)"""
    out = []
    try:
        text = open(ctl_path, "rb").read().decode("cp1252", "replace")
    except (OSError, TypeError):
        return out
    for name, flags in re.findall(r'(?im)^\s*Attribute\s+(\w+)\.VB_MemberFlags\s*=\s*"([0-9A-Fa-f]+)"', text):
        if int(flags, 16) & 0x4 and name not in out:
            out.append(name)
    return out


class RepeaterBindingsDialog(QDialog):
    """the RepeaterBindings of a DataRepeater: PropertyName (one of the bindable properties of the repeated control),
    DataField, Key"""

    def __init__(self, parent, name, binds, props=(), fields=()):
        super().__init__(parent)
        self.setWindowTitle(f"{name} - RepeaterBindings")
        self.props, self.fields = list(props), list(fields)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("The properties of the repeated control and the fields they show:"))
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["PropertyName", "DataField", "Key"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        lay.addWidget(self.table, 1)
        row = QHBoxLayout()
        add, rem = QPushButton("&Add"), QPushButton("&Remove")
        add.clicked.connect(lambda: self.add_row({}))
        rem.clicked.connect(self.remove_row)
        row.addWidget(add)
        row.addWidget(rem)
        row.addStretch(1)
        lay.addLayout(row)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        for d in binds or []:
            self.add_row(d)
        self.resize(480, 300)

    def _combo(self, items, text):
        cb = QComboBox()
        cb.setEditable(True)
        cb.addItems(items)
        cb.setEditText(text)
        return cb

    def add_row(self, d):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setCellWidget(r, 0, self._combo(self.props, str(d.get("PropertyName", ""))))
        self.table.setCellWidget(r, 1, self._combo(self.fields, str(d.get("DataField", ""))))
        self.table.setItem(r, 2, QTableWidgetItem(str(d.get("Key", ""))))

    def remove_row(self):
        r = self.table.currentRow()
        if r >= 0:
            self.table.removeRow(r)

    def chosen(self):
        out = []
        for r in range(self.table.rowCount()):
            p = self.table.cellWidget(r, 0).currentText().strip()
            f = self.table.cellWidget(r, 1).currentText().strip()
            k = self.table.item(r, 2).text().strip() if self.table.item(r, 2) else ""
            if p:
                out.append({"PropertyName": p, "DataField": f, "Key": k})
        return out
