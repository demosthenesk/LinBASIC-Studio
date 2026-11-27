"""The window of the Visual Data Manager (VB6 VDMDI.FRM): its menus, its toolbar (the type of Recordset, the kind of
form, the transactions), its status bar, opening and closing a database, the recent databases, the preferences."""
import os

from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QActionGroup, QBrush, QColor, QIcon, QPainter, QPen, QPixmap
from PyQt6.QtWidgets import (QMainWindow, QMdiArea, QMenu, QToolBar, QLabel, QFileDialog, QMessageBox, QInputDialog,
                             QDialog, QLineEdit, QComboBox, QPushButton, QMdiSubWindow, QGroupBox)

from ..rtobj import RtError
from . import core as C

TITLE = "VisData"


def _icon(kind):
    """small pictures of our own for the toolbar"""
    pm = QPixmap(16, 16)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setPen(QPen(QColor("#000000")))
    if kind in ("Table", "Dynaset", "Snapshot", "PassThrough"):
        p.setBrush(QColor("#ffffff"))
        p.drawRect(1, 2, 13, 11)
        p.fillRect(2, 3, 12, 3, QColor({"Table": "#000080", "Dynaset": "#008000", "Snapshot": "#800000",
                                        "PassThrough": "#808000"}[kind]))
        for y in (8, 11):
            p.drawLine(2, y, 13, y)
        if kind == "Snapshot":
            p.drawEllipse(9, 7, 6, 6)
    elif kind in ("DataControl", "NoDataControl"):
        p.setBrush(QColor("#ffffff"))
        p.drawRect(1, 1, 13, 13)
        for y in (4, 7):
            p.drawLine(3, y, 12, y)
        if kind == "DataControl":
            p.fillRect(2, 10, 12, 3, QColor("#808080"))
        else:
            p.setPen(QPen(QColor("#ff0000"), 2))
            p.drawLine(2, 2, 13, 13)
    elif kind == "DBGrid":
        p.setBrush(QColor("#ffffff"))
        p.drawRect(1, 1, 13, 13)
        for v in (5, 9):
            p.drawLine(1, v, 14, v)
            p.drawLine(v, 1, v, 14)
    elif kind == "BeginTrans":
        p.setBrush(QColor("#00c000"))
        p.drawEllipse(2, 4, 11, 8)
    elif kind == "Rollback":
        p.setBrush(QColor("#c00000"))
        p.drawEllipse(2, 4, 11, 8)
    elif kind == "Commit":
        p.setBrush(QColor("#0000c0"))
        p.drawEllipse(2, 4, 11, 8)
    elif kind == "VisData":                     # a database under a magnifying glass
        p.setBrush(QColor("#00c0c0"))
        p.drawRect(1, 5, 9, 9)
        p.setBrush(QColor("#ffffff"))
        p.drawEllipse(6, 1, 8, 8)
        p.setPen(QPen(QColor("#000000"), 2))
        p.drawLine(12, 9, 15, 14)
    elif kind in ("SQL", "QRY"):                # the letters on a cyan card
        p.fillRect(0, 1, 16, 14, QColor("#00ffff"))
        p.setPen(QPen(QColor("#000080")))
        f = p.font()
        f.setPixelSize(9 if kind == "QRY" else 10)
        f.setBold(True)
        p.setFont(f)
        p.drawText(pm.rect(), Qt.AlignmentFlag.AlignCenter, kind)
        if kind == "QRY":
            for y, c in ((13, "#ff0000"), (14, "#00c000")):
                p.setPen(QPen(QColor(c)))
                p.drawLine(0, y, 15, y)
    p.end()
    return QIcon(pm)


class VisDataWindow(QMainWindow):
    def __init__(self, main):
        super().__init__()
        self.main = main
        self.setWindowTitle(TITLE)
        self.setWindowIcon(_icon("VisData"))
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        prefs = dict(C.DEFAULT_PREFS)
        prefs.update(main.S.get("visdata", {}) or {})
        self.prefs = prefs
        self.s = C.Session()
        self.quiet = False                      # (the tests: no message boxes)
        self.messages = []
        self.answers = []
        self.mdi = QMdiArea()
        self.mdi.setBackground(QBrush(QColor("#808080")))
        self.setCentralWidget(self.mdi)
        self._menus()
        self._toolbar()
        self.status = QLabel(C.MSG[5])
        self.user = QLabel("User: admin")
        self.statusBar().addWidget(self.status, 1)
        self.statusBar().addPermanentWidget(self.user)
        self.dbwin = self.sqlwin = None
        self.hide_db_tools()
        self.fill_mru()
        self.resize(900, 640)
        if prefs.get("open_last") and prefs.get("mru"):
            name, tag = prefs["mru"][0]
            self.open_db(name, tag, silent=True)

    # ---------------------------------------------------------------- menus
    def _act(self, menu, text, fn, check=None):
        a = menu.addAction(text)
        a.triggered.connect(lambda _c=False: fn())
        if check is not None:
            a.setCheckable(True)
            a.setChecked(check)
        return a

    def _menus(self):
        mb = self.menuBar()
        m = mb.addMenu("&File")
        self.m_file = m
        o = m.addMenu("&Open DataBase...")
        self._act(o, "&Microsoft Access...", lambda: self.open_local(C.MSACCESS))
        sub = o.addMenu("&Dbase")
        for t, k in (("&5.0...", C.DBASE5), ("I&V...", C.DBASE4), ("&III...", C.DBASE3)):
            self._act(sub, t, lambda kk=k: self.open_local(kk))
        sub = o.addMenu("&FoxPro")
        for t, k in (("&3.0...", C.FOX30), ("2.&6...", C.FOX26), ("2.&5...", C.FOX25), ("2.&0...", C.FOX20)):
            self._act(sub, t, lambda kk=k: self.open_local(kk))
        sub = o.addMenu("&Paradox")
        for t, k in (("&5.0...", C.PDOX5), ("&4.X...", C.PDOX4), ("&3.X...", C.PDOX3)):
            self._act(sub, t, lambda kk=k: self.open_local(kk))
        self._act(o, "&Btrieve...", lambda: self.open_local(C.BTRIEVE))
        self._act(o, "&Excel...", lambda: self.open_local(C.EXCEL50))
        self._act(o, "&Text Files...", lambda: self.open_local(C.TEXTFILES))
        self._act(o, "&ODBC...", self.open_odbc)
        n = m.addMenu("&New...")
        acc = n.addMenu("&Microsoft Access")
        self._act(acc, "Version &2.0 MDB...", lambda: self.new_mdb(20))
        self._act(acc, "Version &7.0 MDB...", lambda: self.new_mdb(30))
        sub = n.addMenu("&Dbase")
        for t, k in (("&5.0...", C.DBASE5), ("I&V...", C.DBASE4), ("&III...", C.DBASE3)):
            self._act(sub, t, lambda kk=k: self.new_isam(kk))
        sub = n.addMenu("&FoxPro")
        for t, k in (("&3.0...", C.FOX30), ("2.&6...", C.FOX26), ("2.&5...", C.FOX25), ("2.&0...", C.FOX20)):
            self._act(sub, t, lambda kk=k: self.new_isam(kk))
        sub = n.addMenu("&Paradox")
        for t, k in (("&5.0...", C.PDOX5), ("&4.X...", C.PDOX4), ("&3.X...", C.PDOX3)):
            self._act(sub, t, lambda kk=k: self.new_isam(kk))
        self._act(n, "&Btrieve...", lambda: self.new_isam(C.BTRIEVE))
        self._act(n, "&ODBC...", self.new_odbc)
        self._act(n, "&Text Files...", lambda: self.new_isam(C.TEXTFILES))
        self.a_close = self._act(m, "&Close", self.close_db)
        m.addSeparator()
        self.a_impexp = self._act(m, "&Import/Export...", self.import_export)
        self._act(m, "&Workspace...", self.workspace)
        self._act(m, "&Errors...", self.show_errors)
        m.addSeparator()
        cm = m.addMenu("Co&mpact MDB...")
        self._act(cm, "&7.0 MDB...", lambda: self.compact(30))
        self._act(cm, "&2.0 MDB...", lambda: self.compact(20))
        self._act(m, "&Repair MDB...", self.repair)
        self.sep_mru = m.addSeparator()
        self.a_mru = []
        for i in range(8):
            a = m.addAction(f"&{i + 1}")
            a.triggered.connect(lambda _c=False, ii=i: self.open_mru(ii))
            a.setVisible(False)
            self.a_mru.append(a)
        m.addSeparator()
        self._act(m, "E&xit", self.close)
        u = mb.addMenu("&Utility")
        self.m_util = u
        self.a_query = self._act(u, "&Query Builder...", self.query_builder)
        self._act(u, "Data &Form Designer...", self.data_form_designer)
        self._act(u, "&Global Replace...", self.global_replace)
        self.sep_util = u.addSeparator()
        self.a_attach = self._act(u, "&Attachments..", self.attachments)
        self.a_groups = self._act(u, "&Groups/Users...", self.groups_users)
        self.a_sysdb = self._act(u, "&SYSTEM.MD?...", self.system_md)
        u.addSeparator()
        p = u.addMenu("&Preferences")
        self.a_openlast = self._act(p, "&Open Last DataBase on Startup", self.toggle_open_last,
                                    bool(self.prefs.get("open_last")))
        self.a_allowsys = self._act(p, "&Include System Tables", self.toggle_allow_sys,
                                    bool(self.prefs.get("allow_sys")))
        p.addSeparator()
        self._act(p, "&Query Timeout Value...", self.query_timeout)
        self._act(p, "&Login Timeout Value...", self.login_timeout)
        w = mb.addMenu("&Window")
        self._act(w, "&Tile", self.mdi.tileSubWindows)
        self._act(w, "&Cascade", self.mdi.cascadeSubWindows)
        self._act(w, "&Arrange Icons", lambda: None)
        h = mb.addMenu("&Help")
        self._act(h, "&Topics...", self.help_topics)
        h.addSeparator()
        self._act(h, "&About...", self.about)
        # the menu of the right button of the Database Window
        self.popup = QMenu(self)
        pm = self.popup
        self.pu_open = self._act(pm, "&Open", lambda: self.dbwin and self.dbwin.do_open())
        self.pu_design = self._act(pm, "&Design...", lambda: self.dbwin and self.dbwin.do_design())
        self.pu_edit = self._act(pm, "&Edit", lambda: self.dbwin and self.dbwin.do_edit())
        self.pu_rename = self._act(pm, "&Rename", lambda: self.dbwin and self.dbwin.do_rename())
        self.pu_delete = self._act(pm, "De&lete", lambda: self.dbwin and self.dbwin.do_delete())
        self.pu_bar1 = pm.addSeparator()
        self.pu_copy = self._act(pm, "Copy Structure...", lambda: self.dbwin and self.dbwin.do_copy_structure())
        self.pu_removeall = self._act(pm, "Remove All Records", lambda: self.dbwin and self.dbwin.do_remove_all())
        self._act(pm, "Refresh List", self.refresh_tables)
        pm.addSeparator()
        self.pu_newtable = self._act(pm, "New &Table", lambda: self.dbwin and self.dbwin.do_new_table())
        self.pu_newquery = self._act(pm, "New &Query", self.query_builder)

    def _toolbar(self):
        tb = QToolBar()
        tb.setIconSize(QSize(16, 16))
        tb.setMovable(False)
        self.addToolBar(tb)
        self.toolbar = tb
        self.tb = {}
        g1 = QActionGroup(self)
        for key, tip, val in (("Table", "Table type Recordset", C.RS_TABLE),
                              ("Dynaset", "Dynaset type Recordset", C.RS_DYNASET),
                              ("Snapshot", "Snapshot type Recordset", C.RS_SNAPSHOT),
                              ("PassThrough", "Passthrough type Recordset", C.RS_PASSTHRU)):
            a = tb.addAction(_icon(key), tip)
            a.setToolTip(tip)
            a.setCheckable(True)
            a.setChecked(self.prefs.get("rstype") == val)
            a.triggered.connect(lambda _c=False, v=val: self.set_rstype(v))
            g1.addAction(a)
            self.tb[key] = a
        self.tb["PassThrough"].setVisible(False)
        tb.addSeparator()
        g2 = QActionGroup(self)
        for key, tip, val in (("DataControl", "Use Data Control on New Form", C.FORM_DATACTL),
                              ("NoDataControl", "Don't Use Data Control on New Form", C.FORM_NODATACTL),
                              ("DBGrid", "Use DBGrid Control on New Form", C.FORM_DATAGRID)):
            a = tb.addAction(_icon(key), tip)
            a.setToolTip(tip)
            a.setCheckable(True)
            a.setChecked(self.prefs.get("formtype") == val)
            a.triggered.connect(lambda _c=False, v=val: self.set_formtype(v))
            g2.addAction(a)
            self.tb[key] = a
        tb.addSeparator()
        for key, tip, fn in (("BeginTrans", "Begin a Transaction", self.begin_trans),
                             ("Rollback", "Rollback current Transaction", self.rollback),
                             ("Commit", "Commit current Transaction", self.commit)):
            a = tb.addAction(_icon(key), tip)
            a.setToolTip(tip)
            a.triggered.connect(lambda _c=False, f=fn: f())
            self.tb[key] = a

    # ---------------------------------------------------------------- messages
    def msgbar(self, text="", pause=False):
        self.status.setText(C.MSG[5] if not text else text + (C.MSG[6] if pause else ""))

    def say(self, text, icon="info"):
        self.messages.append(text)
        if self.quiet:
            return
        if icon == "warn":
            QMessageBox.warning(self, TITLE, text)
        else:
            QMessageBox.information(self, TITLE, text)

    def ask(self, text, title=TITLE, default=True):
        """a Yes / No question (default: the button that has the focus)"""
        self.messages.append(text)
        if self.quiet:
            return self.answers.pop(0) if self.answers else default
        b = QMessageBox.StandardButton
        return QMessageBox.question(self, title, text, b.Yes | b.No, b.Yes if default else b.No) == b.Yes

    def ask_text(self, label, title=TITLE, default=""):
        self.messages.append(label)
        if self.quiet:
            return self.answers.pop(0) if self.answers else default
        v, ok = QInputDialog.getText(self, title, label, text=default)
        return v if ok else ""

    def show_error(self, e):
        """VisData's ShowError: the error, its number; the Errors collection offered when it holds it"""
        self.msgbar()
        num = e.number if isinstance(e, RtError) else 0
        desc = e.description if isinstance(e, RtError) else str(e)
        text = "The following Error occurred:\n\n" + desc + "\n" + C.MSG[8] + str(num)
        errs = self.s.errors() if self.s.dbe is not None else []
        if errs and errs[0][0] == num:
            if self.ask(text + "\n\n" + C.MSG[9]):
                self.show_errors()
        else:
            self.say(text)

    # ---------------------------------------------------------------- the tools of an open database
    def hide_db_tools(self):
        self.a_close.setEnabled(False)
        self.a_impexp.setEnabled(False)
        self.m_util.menuAction().setEnabled(False)
        for a in (self.sep_util, self.a_attach, self.a_groups, self.a_sysdb):
            a.setVisible(False)
        for k in ("BeginTrans", "Commit", "Rollback"):
            self.tb[k].setEnabled(False)

    def show_db_tools(self):
        from .dbwindow import DatabaseWindow
        from .sqlwin import SQLWindow
        self.a_close.setEnabled(True)
        self.a_impexp.setEnabled(True)
        self.m_util.menuAction().setEnabled(True)
        self.tb["BeginTrans"].setEnabled(True)
        self.tb["Commit"].setEnabled(False)
        self.tb["Rollback"].setEnabled(False)
        access = self.s.datatype == C.MSACCESS
        for a in (self.sep_util, self.a_attach, self.a_groups, self.a_sysdb):
            a.setVisible(access)
        self.pu_rename.setVisible(access)
        if self.s.datatype == C.SQLDB:
            if self.prefs.get("rstype") == C.RS_TABLE:
                self.set_rstype(C.RS_DYNASET)
            self.tb["PassThrough"].setVisible(True)
            self.tb["Table"].setVisible(False)
        else:
            if self.prefs.get("rstype") == C.RS_PASSTHRU:
                self.set_rstype(C.RS_DYNASET)
            self.tb["PassThrough"].setVisible(False)
            self.tb["Table"].setVisible(True)
        if self.dbwin is None:
            self.dbwin = DatabaseWindow(self)
            sw = self.add_window(self.dbwin, (0, 0, 255, 255))
            sw.destroyed.connect(lambda *_: setattr(self, "dbwin", None))
        if self.sqlwin is None:
            self.sqlwin = SQLWindow(self)
            sw = self.add_window(self.sqlwin, (258, 0, 357, 200))
            sw.destroyed.connect(lambda *_: setattr(self, "sqlwin", None))
        self.sqlwin.b_save.setVisible(access)

    def add_window(self, w, geom=None):
        sw = QMdiSubWindow()
        sw.setWidget(w)
        sw.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        sw.setWindowTitle(w.windowTitle())
        t = w.windowTitle()                     # the icon of the window, as VisData has it
        sw.setWindowIcon(_icon("SQL" if t == "SQL Statement" else "QRY" if t == "Query Builder" else
                               "Table" if t.startswith("Table:") else "VisData"))
        w.windowTitleChanged.connect(sw.setWindowTitle)
        self.mdi.addSubWindow(sw)
        if w.minimumSize() == w.maximumSize():  # a fixed window of VisData (BorderStyle 3: no Min / Max)
            sw.setWindowFlags(Qt.WindowType.SubWindow | Qt.WindowType.WindowTitleHint |
                              Qt.WindowType.WindowSystemMenuHint | Qt.WindowType.WindowCloseButtonHint)
            sw.adjustSize()
        if geom:
            sw.setGeometry(*geom)
        sw.show()
        size = getattr(w, "vb_size", None)
        if size is not None and not geom:       # the MDI window around the size of the form in VB6
            sw.resize(size.width() + sw.width() - w.width(), size.height() + sw.height() - w.height())
        return sw

    def refresh_tables(self):
        self.msgbar(C.MSG[7], True)
        try:
            if self.dbwin is not None:
                self.dbwin.load()
        except RtError as e:
            self.show_error(e)
        self.msgbar()

    # ---------------------------------------------------------------- opening and closing
    def open_local(self, kind, path=None, silent=False):
        """Open DataBase > Microsoft Access... (and the ISAMs: Jet's own error, the runtime has none)"""
        titles = {C.MSACCESS: (C.MSG[36], "Microsoft Access MDBs (*.mdb)"),
                  C.DBASE3: (C.MSG[37], "Dbase DBs (*.dbf)"), C.DBASE4: (C.MSG[37], "Dbase DBs (*.dbf)"),
                  C.DBASE5: (C.MSG[37], "Dbase DBs (*.dbf)"), C.FOX20: (C.MSG[38], "FoxPro DBs (*.dbf)"),
                  C.FOX25: (C.MSG[38], "FoxPro DBs (*.dbf)"), C.FOX26: (C.MSG[38], "FoxPro DBs (*.dbf)"),
                  C.FOX30: (C.MSG[38], "FoxPro DBs (*.dbf)"), C.PDOX3: (C.MSG[39], "Paradox DBs (*.db)"),
                  C.PDOX4: (C.MSG[39], "Paradox DBs (*.db)"), C.PDOX5: (C.MSG[39], "Paradox DBs (*.db)"),
                  C.EXCEL50: (C.MSG[40], "Excel Files (*.xls)"), C.BTRIEVE: (C.MSG[41], "Btrieve DBs (FILE.DDF)"),
                  C.TEXTFILES: (C.MSG[42], "Text Files (*.txt)")}
        if path is None and not silent:
            title, flt = titles.get(kind, titles[C.MSACCESS])
            if self.quiet:
                path = self.answers.pop(0) if self.answers else ""
            else:
                path, _ = QFileDialog.getOpenFileName(self, title, os.path.dirname(self.s.dbname or ""),
                                                      f"{flt};;All Files (*.*)")
            if not path:
                return False
            if kind not in (C.MSACCESS, C.BTRIEVE, C.EXCEL50):
                path = os.path.dirname(path)            # (an ISAM database is its directory)
        return self.open_db(path, kind)

    def open_db(self, name, kind=C.MSACCESS, silent=False, connect=None):
        self.msgbar(C.MSG[43], True)
        try:
            self.s.start()
            cn = connect if connect is not None else ("" if kind == C.MSACCESS else kind)
            db = self.s.ws.call("OpenDatabase", name, False, bool(self.s.readonly), cn)
        except RtError as e:
            self.show_error(e)
            return False
        if self.s.db is not None:
            if not self.close_db():
                db.call("Close")
                self.say(C.MSG[35], "warn")
                return False
        self.s.db = db
        self.s.dbname = name
        self.s.datatype = kind if kind != C.SQLDB or connect is None else C.SQLDB
        self.setWindowTitle(f"{TITLE}:{name}")
        self.show_db_tools()
        self.refresh_tables()
        try:
            db.let("QueryTimeout", int(self.prefs.get("query_timeout", 5)))
        except RtError:
            pass
        self.add_mru(name, C.MSACCESS if not db.get("Connect") else db.get("Connect"))
        if self.s.datatype != C.MSACCESS:
            self.msgbar(C.MSG[44])
        else:
            self.msgbar()
        return True

    def close_db(self):
        """CloseCurrentDB: the pending transaction first"""
        if self.s.db is None:
            return True
        if self.s.changed:
            if self.ask(C.MSG[32]):
                self.s.ws.call("CommitTrans")
                self.s.changed = False
            elif self.ask(C.MSG[33]):
                self.s.ws.call("Rollback")
                self.s.changed = False
            else:
                self.say(C.MSG[34], "warn")
                return False
        self.setWindowTitle(TITLE)
        self.hide_db_tools()
        self.s.trans_pending = False
        self.s.dbname = ""
        self.s.readonly = False
        try:
            self.s.db.call("Close")
        except RtError:
            pass
        self.s.db = None
        for sw in list(self.mdi.subWindowList()):
            sw.close()
        self.dbwin = self.sqlwin = None
        return True

    def open_odbc(self):
        """Open DataBase > ODBC...: the ODBC Logon of VisData"""
        cn = self.odbc_logon()
        if not cn:
            return False
        ok = self.open_db("", C.SQLDB, connect=cn)
        if ok:
            self.s.odbc_source = cn
        return ok

    def odbc_logon(self):
        """the ODBC Logon of VisData: the Connect of a data source ("ODBC;DSN=...;"), "" = Cancel"""
        from ..ocxpages import odbc_dsns, odbc_drivers
        d = QDialog(self)
        d.setWindowTitle("ODBC Logon")
        P = C.Place(d, 298, 212)                # (VB6 oracle vb_odbc.png)
        fra = P.put(QGroupBox("Connect Values:"), 8, 8, 282, 161)
        dsn = QComboBox()
        dsn.setEditable(True)
        dsn.addItems(sorted(odbc_dsns(), key=str.upper))
        dsn.setEditText("")
        uid, pwd, dbn, drv, srv = QLineEdit(), QLineEdit(), QLineEdit(), QComboBox(), QLineEdit()
        pwd.setEchoMode(QLineEdit.EchoMode.Password)
        drv.addItems(sorted(odbc_drivers(), key=str.upper))
        drv.setCurrentIndex(-1)
        labels = []
        for t, w, y, h in (("&DSN:", dsn, 16, 21), ("&UID:", uid, 40, 20), ("&Password:", pwd, 62, 20),
                           ("Data&base:", dbn, 84, 20), ("Dri&ver:", drv, 106, 21), ("&Server:", srv, 129, 22)):
            labels.append(P.label(t, 9, y + 3, 50, buddy=w, parent=fra))
            P.put(w, 75, y, 201, h, fra)

        def dsn_changed(text):
            """cboDSNList_Change: Driver and Server for a connection without a DSN"""
            on = not text or text == "(None)"
            for w in (drv, srv, labels[4], labels[5]):
                w.setEnabled(on)
        dsn.editTextChanged.connect(dsn_changed)
        dsn_changed("")
        b_reg = P.put(QPushButton("&Register"), 8, 177, 96, 30)
        b_ok = P.put(QPushButton("&OK"), 116, 177, 84, 30)
        b_cancel = P.put(QPushButton("&Cancel"), 205, 177, 84, 30)
        b_ok.setDefault(True)
        b_ok.clicked.connect(d.accept)
        b_cancel.clicked.connect(d.reject)
        b_reg.clicked.connect(lambda: self.register_dsn(dsn.currentText().strip(), drv.currentText().strip()))
        if self.quiet:
            return self.answers.pop(0) if self.answers else ""
        if not d.exec():
            return ""
        parts = ["ODBC"]
        if dsn.currentText().strip():
            parts.append("DSN=" + dsn.currentText().strip())
        if drv.currentText().strip():
            parts.append("DRIVER={" + drv.currentText().strip() + "}")
        for k, w in (("SERVER", srv), ("UID", uid), ("PWD", pwd), ("DATABASE", dbn)):
            if w.text().strip():
                parts.append(f"{k}={w.text().strip()}")
        return ";".join(parts) + ";"

    def register_dsn(self, dsn, driver):
        """Register (ODBCLOGN.FRM): DBEngine.RegisterDatabase of the DSN with the driver chosen"""
        if not dsn:
            self.say("Please enter a DSN!")
            return False
        if not driver:
            self.say("Please select a Driver!")
            return False
        try:
            self.s.start()
            self.s.dbe.call("RegisterDatabase", dsn, driver, False, "")
        except RtError as e:
            self.show_error(e)
            return False
        return True

    def new_mdb(self, version):
        """New... > Microsoft Access > Version 2.0 / 7.0 MDB...: a new database, opened"""
        if self.quiet:
            path = self.answers.pop(0) if self.answers else ""
        else:
            path, _ = QFileDialog.getSaveFileName(self, C.MSG[55], "", "Microsoft Access MDBs (*.mdb);;All Files (*.*)")
        if not path:
            return False
        if not os.path.splitext(path)[1]:
            path += ".mdb"
        try:
            self.s.start()
            if os.path.exists(path):
                if not self.ask(f"{path} already exists. Replace it?"):
                    return False
                os.remove(path)
            db = self.s.ws.call("CreateDatabase", path, ";LANGID=0x0409;CP=1252;COUNTRY=0",
                                16 if version == 20 else 32)
            db.call("Close")
        except (RtError, OSError) as e:
            self.show_error(e)
            return False
        return self.open_db(path, C.MSACCESS)

    def new_isam(self, kind):
        """New... > Dbase / FoxPro / Paradox / Btrieve / Text Files: Jet's ISAMs (not in the runtime)"""
        d = self.ask_text(C.MSG[47])
        if not d:
            return False
        return self.open_db(d, kind)

    def new_odbc(self):
        drv = self.ask_text("Enter Driver Name:", "Driver Name", "SQL Server")
        if not drv:
            return
        self.msgbar("Enter New Database Parameters")
        self.open_odbc()

    def compact(self, version):
        if self.quiet:
            src = self.answers.pop(0) if self.answers else ""
        else:
            src, _ = QFileDialog.getOpenFileName(self, C.MSG[48], "", "Microsoft Access MDBs (*.mdb);;All Files (*.*)")
        if not src:
            return
        if self.quiet:
            dst = self.answers.pop(0) if self.answers else ""
        else:
            dst, _ = QFileDialog.getSaveFileName(self, C.MSG[51], "", "Microsoft Access MDBs (*.mdb);;All Files (*.*)")
        if not dst:
            return
        if os.path.abspath(dst) == os.path.abspath(src) and self.s.db is not None and \
                os.path.abspath(self.s.dbname) == os.path.abspath(src):
            self.close_db()
        try:
            self.s.start()
            if os.path.exists(dst):
                os.remove(dst)
            self.msgbar(C.MSG[53] + src, True)
            self.s.dbe.call("CompactDatabase", src, dst, ";LANGID=0x0409;CP=1252;COUNTRY=0",
                            16 if version == 20 else 32)
        except (RtError, OSError) as e:
            self.show_error(e)
            return
        self.msgbar()
        if self.ask(C.MSG[54]):
            self.open_db(dst, C.MSACCESS)

    def repair(self):
        if self.quiet:
            path = self.answers.pop(0) if self.answers else ""
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Open Microsoft Access Database to Repair", "",
                                                  "Microsoft Access MDBs (*.mdb);;All Files (*.*)")
        if not path:
            return
        try:
            self.s.start()
            self.msgbar("Repairing " + path, True)
            self.s.dbe.call("RepairDatabase", path)
        except RtError as e:
            self.show_error(e)
            return
        self.msgbar()
        if self.ask("Open Repaired Database?"):
            if self.s.db is not None:
                self.close_db()
            self.open_db(path, C.MSACCESS)

    # ---------------------------------------------------------------- the recent databases
    def fill_mru(self):
        mru = self.prefs.get("mru") or []
        for i, a in enumerate(self.a_mru):
            if i < len(mru):
                a.setText(f"&{i + 1} {mru[i][0]}")
                a.setVisible(True)
            else:
                a.setVisible(False)
        self.sep_mru.setVisible(bool(mru))

    def add_mru(self, name, tag):
        mru = [m for m in (self.prefs.get("mru") or []) if m[0].upper() != name.upper()]
        mru.insert(0, [name, tag])
        self.prefs["mru"] = mru[:8]
        self.save_prefs()
        self.fill_mru()

    def open_mru(self, i):
        mru = self.prefs.get("mru") or []
        if i >= len(mru):
            return
        name, tag = mru[i]
        if tag == C.MSACCESS:
            self.open_db(name, C.MSACCESS)
        elif str(tag).upper().startswith("ODBC"):
            self.open_db(name, C.SQLDB, connect=tag)
        else:
            self.open_db(name, tag)

    def save_prefs(self):
        self.main.S["visdata"] = dict(self.prefs)
        self.main.S.save()

    # ---------------------------------------------------------------- the toolbar
    def set_rstype(self, v):
        self.prefs["rstype"] = v
        for k, val in (("Table", C.RS_TABLE), ("Dynaset", C.RS_DYNASET), ("Snapshot", C.RS_SNAPSHOT),
                       ("PassThrough", C.RS_PASSTHRU)):
            self.tb[k].setChecked(val == v)
        self.save_prefs()

    def set_formtype(self, v):
        self.prefs["formtype"] = v
        for k, val in (("DataControl", C.FORM_DATACTL), ("NoDataControl", C.FORM_NODATACTL),
                       ("DBGrid", C.FORM_DATAGRID)):
            self.tb[k].setChecked(val == v)
        self.save_prefs()

    def begin_trans(self):
        try:
            self.s.ws.call("BeginTrans")
        except RtError as e:
            self.show_error(e)
            return
        self.s.trans_pending = True
        self.tb["BeginTrans"].setEnabled(False)
        self.tb["Commit"].setEnabled(True)
        self.tb["Rollback"].setEnabled(True)

    def commit(self):
        try:
            self.s.ws.call("CommitTrans")
        except RtError as e:
            self.show_error(e)
            return
        self._trans_done()

    def rollback(self):
        if not self.ask("All changes will be gone, Rollback anyway?"):
            return
        try:
            self.s.ws.call("Rollback")
        except RtError as e:
            self.show_error(e)
            return
        self._trans_done()

    def _trans_done(self):
        self.s.trans_pending = False
        self.s.changed = False
        self.tb["BeginTrans"].setEnabled(True)
        self.tb["Commit"].setEnabled(False)
        self.tb["Rollback"].setEnabled(False)

    def check_trans(self, msg=C.MSG[1]):
        """CheckTransPending: no structure changes inside a transaction"""
        if self.s.trans_pending:
            self.say(msg, "warn")
            return True
        return False

    # ---------------------------------------------------------------- the other commands
    def workspace(self):
        d = QDialog(self)
        d.setWindowTitle("Login")
        P = C.Place(d, 188, 132)                # (VB6 oracle vb_login.png)
        user, pwd = QLineEdit(), QLineEdit()
        pwd.setEchoMode(QLineEdit.EchoMode.Password)
        P.label("Login Name: ", 8, 8, 62)
        P.put(user, 8, 24, 169, 19)
        P.label("Password: ", 8, 48, 53)
        P.put(pwd, 8, 64, 169, 19)
        b_ok = P.put(QPushButton("&OK"), 8, 96, 81, 25)
        b_cancel = P.put(QPushButton("&Cancel"), 96, 96, 81, 25)
        b_ok.setDefault(True)
        b_ok.clicked.connect(d.accept)
        b_cancel.clicked.connect(d.reject)
        if self.quiet or not d.exec():
            return
        try:
            self.s.start()
            ws = self.s.dbe.call("CreateWorkspace", "", user.text().strip() or "admin", pwd.text())
            self.s.ws = ws
        except RtError as e:
            self.show_error(e)
            return
        self.user.setText(f" User: {self.s.username()} ")

    def show_errors(self):
        from .dialogs import ErrorsWindow
        errs = self.s.errors() if self.s.dbe is not None else []
        if not errs:
            self.say(C.MSG[30], "warn")
            return None
        w = ErrorsWindow(self, errs)
        self.add_window(w)
        return w

    def toggle_open_last(self):
        self.prefs["open_last"] = self.a_openlast.isChecked()
        self.save_prefs()

    def toggle_allow_sys(self):
        self.prefs["allow_sys"] = self.a_allowsys.isChecked()
        self.save_prefs()
        if self.s.db is not None:
            self.refresh_tables()

    def query_timeout(self):
        if self.s.db is None:
            self.say("No Database Open", "warn")
            return
        try:
            cur = self.s.db.get("QueryTimeout")
        except RtError:
            cur = self.prefs.get("query_timeout", 5)
        v = self.ask_text("Query Timeout (in seconds):", TITLE, str(cur))
        if not v.strip():
            return
        try:
            n = int(float(v))
            self.s.db.let("QueryTimeout", n)
            self.prefs["query_timeout"] = n
            self.save_prefs()
        except (ValueError, RtError) as e:
            self.show_error(e)

    def login_timeout(self):
        v = self.ask_text("Login Timeout (in seconds):", TITLE, str(self.prefs.get("login_timeout", 20)))
        if not v.strip():
            return
        try:
            n = int(float(v))
            if n >= 0:
                self.prefs["login_timeout"] = n
                self.save_prefs()
                self.s.start()
                self.s.dbe.let("LoginTimeout", n)
        except (ValueError, RtError) as e:
            self.show_error(e)

    def groups_users(self):
        """the security of a Jet workgroup: not in the runtime (decision of 2026-11-27)"""
        self.say("No Users found, try 'Utility/System MD?'!\n\nThe security of a Jet workgroup (users, groups, "
                 "permissions, SYSTEM.MDW) is not part of LinBASIC Studio.", "warn")

    def system_md(self):
        self.say("The security of a Jet workgroup (SYSTEM.MD?) is not part of LinBASIC Studio.", "warn")

    def attachments(self):
        from .dialogs import AttachmentsWindow
        w = AttachmentsWindow(self)
        self.add_window(w)
        return w

    def import_export(self):
        from .dialogs import ImportExportDialog
        d = ImportExportDialog(self)
        if not self.quiet:
            d.exec()
        return d

    def query_builder(self):
        from .querybld import QueryBuilder
        w = QueryBuilder(self)
        self.add_window(w)
        return w

    def data_form_designer(self):
        from .dfd import DataFormDesigner
        if getattr(self.main, "project", None) is None:
            self.say("This function requires an active project!", "warn")
            return None
        d = DataFormDesigner(self)
        if not self.quiet:
            d.exec()
        return d

    def global_replace(self):
        from .dialogs import GlobalReplaceDialog
        d = GlobalReplaceDialog(self)
        if not self.quiet:
            d.exec()
        return d

    def help_topics(self):
        self.main.help_contents()

    def about(self):
        self.say("Visual Data Manager\n\nLinBASIC Studio's own, after the VisData of Visual Basic 6: the databases "
                 "of the programs through the runtime's DAO.")

    def closeEvent(self, e):
        if self.s.db is not None and not self.close_db():
            e.ignore()
            return
        self.save_prefs()
        self.s.stop()
        super().closeEvent(e)
