"""Data Link Properties (phase 12.3): the dialog of Windows' OLE DB that the Data View of VB6 opens for Add a Data Link
and Modify (VB6 oracle Claude Enviroment/oracles/12.3, dv4, dv5): Provider (the providers LinBASIC Studio's ADO has),
Connection (Jet: the database, the user and password; ODBC: a data source name or a connection string), Advanced (the
connect timeout, the access permissions), All (the properties of the connection string); Test Connection."""
import os
import re

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QDialog, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QListWidget,
                             QListWidgetItem, QPushButton, QLineEdit, QCheckBox, QRadioButton, QComboBox, QFileDialog,
                             QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QInputDialog,
                             QGroupBox, QStackedWidget, QSpinBox)

PROVIDERS = [("Microsoft.Jet.OLEDB.3.51", "Microsoft Jet 3.51 OLE DB Provider"),
             ("Microsoft.Jet.OLEDB.4.0", "Microsoft Jet 4.0 OLE DB Provider"),
             ("MSDASQL.1", "Microsoft OLE DB Provider for ODBC Drivers"),
             ("MSDataShape.1", "MSDataShape")]
PERMISSIONS = ["Read", "ReadWrite", "Share Deny None", "Share Deny Read", "Share Deny Write", "Share Exclusive",
               "Write"]


def parse(cs):
    """a connection string as [(name, value)] in its order"""
    out = []
    for part in re.findall(r'(?:[^;"]|"[^"]*")+', cs or ""):
        if "=" in part:
            k, v = part.split("=", 1)
            v = v.strip()
            if len(v) >= 2 and v[0] == v[-1] == '"':
                v = v[1:-1]
            out.append((k.strip(), v))
    return out


def get(props, name, default=""):
    return next((v for k, v in props if k.lower() == name.lower()), default)


def build(props):
    out = []
    for k, v in props:
        if v is None:
            continue
        out.append(f'{k}="{v}"' if ";" in v and not v.startswith('"') else f"{k}={v}")
    return ";".join(out)


class DataLinkDialog(QDialog):
    def __init__(self, parent=None, connect="", base_dir=None):
        super().__init__(parent)
        self.setWindowTitle("Data Link Properties")
        self.base_dir = base_dir
        props = parse(connect)
        self.extra = [(k, v) for k, v in props if k.lower() not in (
            "provider", "data source", "user id", "password", "persist security info", "extended properties",
            "initial catalog", "connect timeout", "mode")]
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs)
        # -- Provider
        w = QWidget()
        v = QVBoxLayout(w)
        v.addWidget(QLabel("Select the data you want to connect to:"))
        self.providers = QListWidget()
        hdr = QListWidgetItem("OLE DB Provider(s)")
        hdr.setFlags(Qt.ItemFlag.NoItemFlags)
        self.providers.addItem(hdr)
        for key, text in PROVIDERS:
            it = QListWidgetItem(text)
            it.setData(Qt.ItemDataRole.UserRole, key)
            self.providers.addItem(it)
        v.addWidget(self.providers, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        self.b_next = QPushButton("&Next >>")
        self.b_next.clicked.connect(lambda: self.tabs.setCurrentIndex(1))
        row.addWidget(self.b_next)
        v.addLayout(row)
        self.tabs.addTab(w, "Provider")
        # -- Connection: Jet / ODBC
        self.conn = QStackedWidget()
        self.conn.addWidget(self._jet_page())
        self.conn.addWidget(self._odbc_page())
        self.tabs.addTab(self.conn, "Connection")
        # -- Advanced
        w = QWidget()
        v = QVBoxLayout(w)
        gb = QGroupBox("Other")
        gl = QGridLayout(gb)
        gl.addWidget(QLabel("Connect timeout:"), 0, 0)
        self.timeout = QSpinBox()
        self.timeout.setRange(0, 9999)
        self.timeout.setSpecialValueText(" ")
        gl.addWidget(self.timeout, 0, 1)
        gl.addWidget(QLabel("seconds."), 0, 2)
        gl.addWidget(QLabel("Access permissions:"), 1, 0)
        self.perms = QListWidget()
        for p in PERMISSIONS:
            it = QListWidgetItem(p)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Unchecked)
            self.perms.addItem(it)
        gl.addWidget(self.perms, 1, 1, 1, 2)
        v.addWidget(gb)
        v.addStretch(1)
        self.tabs.addTab(w, "Advanced")
        # -- All
        w = QWidget()
        v = QVBoxLayout(w)
        v.addWidget(QLabel("These are the initialization properties for this type of data. To edit a value, select a "
                           "property, then choose Edit Value below."))
        self.all = QTableWidget(0, 2)
        self.all.setHorizontalHeaderLabels(["Name", "Value"])
        self.all.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.all.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.all.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.all.verticalHeader().hide()
        v.addWidget(self.all, 1)
        row = QHBoxLayout()
        self.b_edit = QPushButton("&Edit Value...")
        self.b_edit.clicked.connect(self.edit_value)
        row.addWidget(self.b_edit)
        row.addStretch(1)
        v.addLayout(row)
        self.tabs.addTab(w, "All")
        # -- OK, Cancel, Help
        row = QHBoxLayout()
        row.addStretch(1)
        for text, fn in (("OK", self.accept), ("Cancel", self.reject), ("Help", lambda: None)):
            b = QPushButton(text)
            b.clicked.connect(fn)
            row.addWidget(b)
            setattr(self, "b_" + text.lower(), b)
        lay.addLayout(row)
        self.tabs.currentChanged.connect(self._tab)
        self.providers.currentRowChanged.connect(self._provider)
        # the values
        prov = get(props, "Provider", "Microsoft.Jet.OLEDB.4.0")
        keys = [k.lower() for k, _ in PROVIDERS]
        p = prov.lower()
        idx = keys.index(p) if p in keys else next((i for i, k in enumerate(keys) if p.startswith(k.split(".")[0].lower()
                                                                                                     + ".")), 1)
        self.providers.setCurrentRow(idx + 1)
        self.db.setText(get(props, "Data Source") if self.is_jet() else "")
        self.user.setText(get(props, "User ID", "Admin" if self.is_jet() else ""))
        self.odbc_user.setText(get(props, "User ID"))
        pw = get(props, "Password", None)
        self.blank.setChecked(pw in (None, ""))
        self.password.setText(pw or "")
        self.save_pw.setChecked(get(props, "Persist Security Info", "False").lower() == "true")
        if not self.is_jet():
            ext = get(props, "Extended Properties")
            if ext:
                self.rb_cs.setChecked(True)
                self.cs.setText(ext)
            else:
                self.rb_dsn.setChecked(True)
                self.dsn.setCurrentText(get(props, "Data Source"))
            self.catalog.setCurrentText(get(props, "Initial Catalog"))
        t = get(props, "Connect Timeout")
        if t.isdigit():
            self.timeout.setValue(int(t))
        mode = get(props, "Mode")
        for i in range(self.perms.count()):
            if self.perms.item(i).text().replace(" ", "").lower() in mode.replace(" ", "").lower().split("|"):
                self.perms.item(i).setCheckState(Qt.CheckState.Checked)
        self.resize(380, 470)

    def _jet_page(self):
        w = QWidget()
        g = QGridLayout(w)
        g.addWidget(QLabel("Specify the following to connect to Access data:"), 0, 0, 1, 3)
        g.addWidget(QLabel("1. Select or enter a database name:"), 1, 0, 1, 3)
        self.db = QLineEdit()
        g.addWidget(self.db, 2, 0, 1, 2)
        b = QPushButton("...")
        b.setFixedWidth(24)
        b.clicked.connect(self.browse)
        g.addWidget(b, 2, 2)
        g.addWidget(QLabel("2. Enter information to log on to the database:"), 3, 0, 1, 3)
        g.addWidget(QLabel("User name:"), 4, 0)
        self.user = QLineEdit("Admin")
        g.addWidget(self.user, 4, 1, 1, 2)
        self.pw_label = QLabel("Password:")
        g.addWidget(self.pw_label, 5, 0)
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.EchoMode.Password)
        g.addWidget(self.password, 5, 1, 1, 2)
        self.blank = QCheckBox("&Blank password")
        self.save_pw = QCheckBox("Allow &saving password")
        g.addWidget(self.blank, 6, 1)
        g.addWidget(self.save_pw, 6, 2)
        self.blank.toggled.connect(lambda on: (self.password.setEnabled(not on), self.pw_label.setEnabled(not on)))
        g.setRowStretch(7, 1)
        b = QPushButton("&Test Connection")
        b.clicked.connect(self.test)
        g.addWidget(b, 8, 2)
        self.b_test = b
        return w

    def _odbc_page(self):
        from .ocxpages import odbc_dsns
        w = QWidget()
        g = QGridLayout(w)
        g.addWidget(QLabel("Specify the following to connect to ODBC data:"), 0, 0, 1, 3)
        g.addWidget(QLabel("1. Specify the source of data:"), 1, 0, 1, 3)
        self.rb_dsn = QRadioButton("Use &data source name")
        self.rb_cs = QRadioButton("Use c&onnection string")
        self.rb_dsn.setChecked(True)
        g.addWidget(self.rb_dsn, 2, 0, 1, 3)
        self.dsn = QComboBox()
        self.dsn.setEditable(True)
        self.dsn.addItems(odbc_dsns())
        g.addWidget(self.dsn, 3, 0, 1, 2)
        b = QPushButton("&Refresh")
        b.clicked.connect(lambda: (self.dsn.clear(), self.dsn.addItems(odbc_dsns())))
        g.addWidget(b, 3, 2)
        g.addWidget(self.rb_cs, 4, 0, 1, 3)
        self.cs = QLineEdit()
        g.addWidget(self.cs, 5, 0, 1, 2)
        b = QPushButton("B&uild...")
        b.clicked.connect(self.build_cs)
        g.addWidget(b, 5, 2)
        g.addWidget(QLabel("2. Enter information to log on to the server:"), 6, 0, 1, 3)
        g.addWidget(QLabel("User name:"), 7, 0)
        self.odbc_user = QLineEdit()
        g.addWidget(self.odbc_user, 7, 1, 1, 2)
        g.addWidget(QLabel("Password:"), 8, 0)
        self.odbc_pw = QLineEdit()
        self.odbc_pw.setEchoMode(QLineEdit.EchoMode.Password)
        g.addWidget(self.odbc_pw, 8, 1, 1, 2)
        g.addWidget(QLabel("3. Enter the initial catalog to use:"), 9, 0, 1, 3)
        self.catalog = QComboBox()
        self.catalog.setEditable(True)
        g.addWidget(self.catalog, 10, 0, 1, 3)
        g.setRowStretch(11, 1)
        b = QPushButton("&Test Connection")
        b.clicked.connect(self.test)
        g.addWidget(b, 12, 2)
        self.rb_dsn.toggled.connect(lambda on: (self.dsn.setEnabled(on), self.cs.setEnabled(not on)))
        self.cs.setEnabled(False)
        return w

    def is_jet(self):
        return self.provider().lower().startswith("microsoft.jet")

    def provider(self):
        it = self.providers.currentItem()
        k = it.data(Qt.ItemDataRole.UserRole) if it is not None else None
        return k or "Microsoft.Jet.OLEDB.4.0"

    def _provider(self, _row):
        self.conn.setCurrentIndex(0 if self.is_jet() else 1)

    def _tab(self, i):
        if i == 3:
            self.fill_all()

    def browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Access Database", self.base_dir or "",
                                              "Access Database (*.mdb);;All Files (*.*)")
        if path:
            self.db.setText(path)

    def build_cs(self):
        from .ocxpages import new_dsn_dialog
        name = new_dsn_dialog(self)
        if name:
            self.cs.setText(f"DSN={name}")

    def connect_string(self):
        """the connection string the choices make (VB6's for Jet 4.0: Provider, Data Source, Persist Security Info)"""
        prov = self.provider()
        props = [("Provider", prov)]
        if self.is_jet():
            props.append(("Data Source", self.db.text().strip()))
            user = self.user.text().strip()
            if user and user != "Admin":
                props.append(("User ID", user))
            if not self.blank.isChecked() and self.password.text():
                props.append(("Password", self.password.text()))
        else:
            if self.rb_cs.isChecked():
                props.append(("Extended Properties", self.cs.text().strip()))
            else:
                props.append(("Data Source", self.dsn.currentText().strip()))
            if self.odbc_user.text().strip():
                props.append(("User ID", self.odbc_user.text().strip()))
            if self.odbc_pw.text():
                props.append(("Password", self.odbc_pw.text()))
            if self.catalog.currentText().strip():
                props.append(("Initial Catalog", self.catalog.currentText().strip()))
        props.insert(1 if not self.is_jet() else len(props), ("Persist Security Info",
                                                              "True" if self.save_pw.isChecked() else "False"))
        if self.timeout.value():
            props.append(("Connect Timeout", str(self.timeout.value())))
        mode = [self.perms.item(i).text().replace(" ", "") for i in range(self.perms.count())
                if self.perms.item(i).checkState() == Qt.CheckState.Checked]
        if mode:
            props.append(("Mode", "|".join(mode)))
        props += self.extra
        return build(props)

    def fill_all(self):
        props = parse(self.connect_string())
        self.all.setRowCount(0)
        for k, v in props:
            r = self.all.rowCount()
            self.all.insertRow(r)
            self.all.setItem(r, 0, QTableWidgetItem(k))
            self.all.setItem(r, 1, QTableWidgetItem(v))

    def edit_value(self):
        r = self.all.currentRow()
        if r < 0:
            return
        name = self.all.item(r, 0).text()
        val, ok = QInputDialog.getText(self, "Edit Property Value", f"Property Name: {name}\n\nProperty Value:",
                                       text=self.all.item(r, 1).text())
        if not ok:
            return
        low = name.lower()
        if low == "data source" and self.is_jet():
            self.db.setText(val)
        elif low == "user id":
            (self.user if self.is_jet() else self.odbc_user).setText(val)
        elif low == "connect timeout" and val.isdigit():
            self.timeout.setValue(int(val))
        else:
            self.extra = [(k, v) for k, v in self.extra if k.lower() != low] + [(name, val)]
        self.fill_all()

    def test(self):
        """Test Connection: the connection opened through LinBASIC Studio's ADO"""
        from . import dbinfo
        cs = self.connect_string()
        try:
            from .dsrdesign import design_cs
            cs = design_cs(cs, self.base_dir)
        except Exception:                       # noqa: BLE001 - (the string as it is)
            pass
        try:
            dbinfo.forget()
            dbinfo.ado_objects(cs)
            ok, text = True, "Test connection succeeded."
        except dbinfo.AdoError as e:
            ok, text = False, f"Test connection failed because of an error in initializing provider. {e}"
        self.last_test = (ok, text)
        if getattr(self, "quiet", False):
            return ok
        QMessageBox.information(self, "Microsoft Data Link", text) if ok else \
            QMessageBox.warning(self, "Microsoft Data Link", text)
        return ok
