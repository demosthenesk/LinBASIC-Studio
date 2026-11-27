"""The Add-In designer (Connect of an add-in project, phase 12.5), as VB6 shows it (oracles 12.5/connect-general.png,
connect-advanced.png, connect-loadbehavior.png): General - Addin Display Name, Addin Description, Application,
Application Version, Initial Load Behavior, "Addin is command-line safe (Does not put up any UI)"; Advanced - Satellite
DLL Name, Registry Key for Additional Addin Data, Addin Specific Data (Name, Data; New Value, Edit Value, Delete Value).
The values are the properties of the .Dsr (DisplayName, Description, AppName, AppVer, LoadName, LoadBehavior,
CmdLineSupport ...); Make registers the add-in with them."""
from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import (QWidget, QTabWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPlainTextEdit,
                             QComboBox, QCheckBox, QFrame, QTreeWidget, QTreeWidgetItem, QPushButton, QInputDialog)

APPS = ["Visual Basic"]
VERSIONS = {"Visual Basic": ["Visual Basic 6.0"]}
# Initial Load Behavior: its text and LoadBehavior (1 Startup, 4 Command Line)
BEHAVIORS = [("Command Line / Startup", 5), ("Command Line", 4), ("Startup", 1), ("None", 0)]


class AddInDesigner(QWidget):
    modified = pyqtSignal()

    def __init__(self, main, module):
        super().__init__()
        self.main = main
        self.module = module
        self.root = module.sf.root
        self._filling = True
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs)
        # ---- General
        g = QWidget()
        v = QVBoxLayout(g)
        self.name = QLineEdit(str(self.root.get("DisplayName", "") or ""))
        self.desc = QPlainTextEdit(str(self.root.get("Description", "") or ""))
        self.app = QComboBox()
        self.app.addItems(APPS)
        self.app.setCurrentText(str(self.root.get("AppName", "Visual Basic") or "Visual Basic"))
        self.ver = QComboBox()
        self.ver.addItems([""] + VERSIONS["Visual Basic"])
        ver = str(self.root.get("AppVer", "") or "")
        self.ver.setCurrentIndex(1 if ver else 0)
        self.load = QComboBox()
        self.cmdsafe = QCheckBox("Addin is &command-line safe (Does not put up any UI)")
        self.cmdsafe.setChecked(int(self.root.get("CmdLineSupport", 0) or 0) != 0)
        for text, w in (("Addin Display &Name", self.name), ("Addin &Description", self.desc)):
            lb = QLabel(text)
            lb.setBuddy(w)
            v.addWidget(lb)
            v.addWidget(w)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        v.addWidget(line)
        for text, w in (("A&pplication", self.app), ("Application &Version", self.ver),
                        ("Initial &Load Behavior", self.load)):
            lb = QLabel(text)
            lb.setBuddy(w)
            v.addWidget(lb)
            v.addWidget(w)
        v.addWidget(self.cmdsafe)
        v.addStretch(1)
        self.tabs.addTab(g, "General")
        # ---- Advanced
        a = QWidget()
        av = QVBoxLayout(a)
        self.sat = QLineEdit(str(self.root.get("SatName", "") or ""))
        self.regkey = QLineEdit(str(self.root.get("RegKey", "") or ""))
        for text, w in (("Satellite DLL &Name", self.sat), ("Registry &Key for Additional Addin Data", self.regkey)):
            lb = QLabel(text)
            lb.setBuddy(w)
            av.addWidget(lb)
            av.addWidget(w)
        lb = QLabel("Addin &Specific Data")
        self.data = QTreeWidget()
        self.data.setHeaderLabels(["Name", "Data"])
        self.data.setRootIsDecorated(False)
        lb.setBuddy(self.data)
        av.addWidget(lb)
        av.addWidget(self.data, 1)
        self.data.addTopLevelItem(QTreeWidgetItem(["(Default)", "(value not set)"]))
        row = QHBoxLayout()
        for text, fn in (("Ne&w Value", self.new_value), ("&Edit Value", self.edit_value),
                         ("&Delete Value", self.delete_value)):
            b = QPushButton(text)
            b.clicked.connect(fn)
            row.addWidget(b)
        row.addStretch(1)
        av.addLayout(row)
        self.tabs.addTab(a, "Advanced")
        self.fill_behaviors()
        lb_val = int(self.root.get("LoadBehavior", 0) or 0)
        for i, (_, n) in enumerate(BEHAVIORS):
            if n == lb_val:
                self.load.setCurrentIndex(i)
        for w in (self.name, self.sat, self.regkey):
            w.textEdited.connect(self.changed)
        self.desc.textChanged.connect(self.changed)
        self.app.currentIndexChanged.connect(self.changed)
        self.ver.currentIndexChanged.connect(self.version_changed)
        self.load.currentIndexChanged.connect(self.changed)
        self.cmdsafe.toggled.connect(self.changed)
        self._filling = False

    def fill_behaviors(self):
        """the Load Behaviors of the version of the application (VB6: none before a version is chosen)"""
        cur = self.load.currentText()
        self.load.blockSignals(True)
        self.load.clear()
        if self.ver.currentText():
            self.load.addItems([t for t, _ in BEHAVIORS])
            if cur:
                self.load.setCurrentText(cur)
        self.load.blockSignals(False)

    def version_changed(self):
        self.fill_behaviors()
        self.changed()

    def changed(self, *_):
        if self._filling:
            return
        r = self.root
        r.set("DisplayName", self.name.text())
        r.set("Description", self.desc.toPlainText())
        r.set("AppName", self.app.currentText())
        r.set("AppVer", "Visual Basic 98 (ver 6.0)" if self.ver.currentText() else "")
        t = self.load.currentText()
        r.set("LoadName", t)
        r.set("LoadBehavior", dict(BEHAVIORS).get(t, 0))
        r.set("CmdLineSupport", -1 if self.cmdsafe.isChecked() else 0)
        if self.sat.text():
            r.set("SatName", self.sat.text())
        if self.regkey.text():
            r.set("RegKey", self.regkey.text())
        self.module.dirty = True
        self.modified.emit()

    # Addin Specific Data
    def new_value(self):
        name, ok = QInputDialog.getText(self, "New Value", "Name:")
        if ok and name:
            self.data.addTopLevelItem(QTreeWidgetItem([name, ""]))

    def edit_value(self):
        it = self.data.currentItem()
        if it is None or it.text(0) == "(Default)" and False:
            return
        val, ok = QInputDialog.getText(self, "Edit Value", it.text(0), text=it.text(1))
        if ok:
            it.setText(1, val)

    def delete_value(self):
        it = self.data.currentItem()
        if it is not None and it.text(0) != "(Default)":
            self.data.takeTopLevelItem(self.data.indexOfTopLevelItem(it))
