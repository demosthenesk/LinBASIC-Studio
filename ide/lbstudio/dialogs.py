"""Find/Replace, Options and Project Properties dialogs."""
import os

from PyQt6.QtWidgets import (QDialog, QGridLayout, QLabel, QComboBox, QCheckBox, QPushButton, QLineEdit,
                             QDialogButtonBox, QFormLayout, QHBoxLayout, QFileDialog, QSpinBox, QVBoxLayout,
                             QRadioButton, QGroupBox, QMessageBox, QPlainTextEdit, QWidget)
from PyQt6.QtCore import Qt
from PyQt6.Qsci import QsciScintilla as S

from . import vbmodel
from . import APP_NAME


class FindDialog(QDialog):
    """Edit > Find / Replace of VB6: Current Procedure / Current Module / Current Project"""

    def __init__(self, main):
        super().__init__(main)
        self.main = main
        self.setWindowTitle("Find")
        g = QGridLayout(self)
        g.addWidget(QLabel("Find What:"), 0, 0)
        self.what = QComboBox()
        self.what.setEditable(True)
        self.what.setMinimumWidth(320)
        g.addWidget(self.what, 0, 1)
        self.with_label = QLabel("Replace With:")
        g.addWidget(self.with_label, 1, 0)
        self.repl = QComboBox()
        self.repl.setEditable(True)
        g.addWidget(self.repl, 1, 1)
        box = QGroupBox("Search")
        bl = QVBoxLayout(box)
        self.r_module = QRadioButton("Current Module")
        self.r_project = QRadioButton("Current Project")
        self.r_module.setChecked(True)
        bl.addWidget(self.r_module)
        bl.addWidget(self.r_project)
        g.addWidget(box, 2, 0, 1, 2)
        self.whole = QCheckBox("Find Whole Word Only")
        self.case = QCheckBox("Match Case")
        self.up = QCheckBox("Search Up")
        g.addWidget(self.whole, 3, 0, 1, 2)
        g.addWidget(self.case, 4, 0, 1, 2)
        g.addWidget(self.up, 5, 0, 1, 2)
        col = QVBoxLayout()
        self.b_find = QPushButton("&Find Next")
        self.b_find.setDefault(True)
        self.b_repl = QPushButton("&Replace")
        self.b_all = QPushButton("Replace &All")
        b_cancel = QPushButton("Cancel")
        for b in (self.b_find, self.b_repl, self.b_all, b_cancel):
            col.addWidget(b)
        col.addStretch()
        g.addLayout(col, 0, 2, 6, 1)
        self.b_find.clicked.connect(self.find_next)
        self.b_repl.clicked.connect(self.replace_one)
        self.b_all.clicked.connect(self.replace_all)
        b_cancel.clicked.connect(self.hide)

    def show_find(self, replace):
        self.setWindowTitle("Replace" if replace else "Find")
        for w in (self.with_label, self.repl, self.b_repl, self.b_all):
            w.setVisible(replace)
        ed = self.main.current_editor()
        if ed is not None and ed.sel_text() and "\n" not in ed.sel_text():
            self.what.setEditText(ed.sel_text())
        self.show()
        self.raise_()
        self.what.setFocus()
        self.what.lineEdit().selectAll()

    def _remember(self, box):
        t = box.currentText()
        if t and box.findText(t) < 0:
            box.insertItem(0, t)

    def _editors(self):
        """the code windows to search: the current one, then (Current Project) the others"""
        cur = self.main.current_cw()
        out = [cur] if cur else []
        if self.r_project.isChecked():
            for m in self.main.project.modules:
                if cur is None or m is not cur.module:
                    out.append(m)
        return out

    def find_next(self):
        text = self.what.currentText()
        if not text:
            return False
        self._remember(self.what)
        targets = self._editors()
        for i, t in enumerate(targets):
            cw = t if not isinstance(t, vbmodel.Module) else None
            if cw is None:
                if text.lower() not in t.code_text().lower() if not self.case.isChecked() else text not in t.code_text():
                    continue
                cw = self.main.open_code(t)
                cw.editor.SendScintilla(S.SCI_GOTOPOS, 0)
            ed = cw.editor
            if ed.findFirst(text, False, self.case.isChecked(), self.whole.isChecked(), i == 0 and not self.r_project.isChecked(),
                            not self.up.isChecked()):
                return True
        QMessageBox.information(self, "Find", "The specified region has been searched." if self.r_project.isChecked()
                                else "Search string not found.")
        return False

    def replace_one(self):
        ed = self.main.current_editor()
        if ed is None:
            return
        self._remember(self.repl)
        if ed.hasSelectedText() and ed.selectedText().lower() == self.what.currentText().lower():
            ed.replace(self.repl.currentText())
        self.find_next()

    def replace_all(self):
        n = 0
        text, rep = self.what.currentText(), self.repl.currentText()
        if not text:
            return
        mods = self.main.project.modules if self.r_project.isChecked() else []
        cws = [self.main.open_code(m) for m in mods] if mods else [self.main.current_cw()]
        for cw in cws:
            if cw is None:
                continue
            ed = cw.editor
            ed.beginUndoAction()
            ed.SendScintilla(S.SCI_GOTOPOS, 0)
            while ed.findFirst(text, False, self.case.isChecked(), self.whole.isChecked(), False, True):
                ed.replace(rep)
                n += 1
            ed.endUndoAction()
        QMessageBox.information(self, "Replace", f"{n} replacement(s) were made.")


class OptionsDialog(QDialog):
    """Tools > Options, in VB6's tabs (DEV-8/me4 picture 1): Editor, General, Docking, Environment (without
    Editor Format and Advanced: the colours of the code are the theme's, the rest is not LinBASIC's)"""

    def __init__(self, main):
        from PyQt6.QtWidgets import QTabWidget
        super().__init__(main)
        self.main = main
        S = main.S
        self.setWindowTitle("Options")
        outer = QVBoxLayout(self)
        self.tabs = QTabWidget()
        outer.addWidget(self.tabs)

        def page(title):
            w = QWidget()
            f = QFormLayout(w)
            self.tabs.addTab(w, title)
            return f
        # Editor
        f = page("Editor")
        self.auto_list = QCheckBox("Auto List Members")
        self.auto_list.setChecked(bool(S["auto_list_members"]))
        self.auto_info = QCheckBox("Auto Quick Info")
        self.auto_info.setChecked(bool(S["auto_quick_info"]))
        f.addRow("Code Settings:", self.auto_list)
        f.addRow("", self.auto_info)
        # General: Form Grid Settings (DEV-8/fd1 stage 7), Error Trapping (the code window's Toggle changes it for
        # one session)
        f = page("General")
        gb = QGroupBox("Form Grid Settings")
        gl = QGridLayout(gb)
        self.show_grid = QCheckBox("Show &Grid")
        self.show_grid.setChecked(bool(S["show_grid"]))
        gl.addWidget(self.show_grid, 0, 0, 1, 2)
        gl.addWidget(QLabel("Grid Units: Twips"), 1, 0, 1, 2)
        self.grid_w, self.grid_h = QSpinBox(), QSpinBox()
        for i, (sb, text, key) in enumerate(((self.grid_w, "&Width:", "grid_width"), (self.grid_h, "&Height:", "grid_height")), 2):
            sb.setRange(30, 1200)
            try:
                sb.setValue(int(S[key]))
            except (TypeError, ValueError):
                sb.setValue(120)
            lab = QLabel(text)                  # &Width: the & is a mnemonic (a buddy), not shown
            lab.setBuddy(sb)
            gl.addWidget(lab, i, 0)
            gl.addWidget(sb, i, 1)
        self.align_grid = QCheckBox("Align C&ontrols to Grid")
        self.align_grid.setChecked(bool(S["align_grid"]))
        gl.addWidget(self.align_grid, 4, 0, 1, 2)
        f.addRow(gb)
        self.trap = []
        from PyQt6.QtWidgets import QButtonGroup
        self.trap_group = QButtonGroup(self)
        try:
            cur = int(S["error_trapping"])
        except (TypeError, ValueError):
            cur = 1
        tb = QGroupBox("Error Trapping")
        tl = QVBoxLayout(tb)
        for n, text in enumerate(("Break on &All Errors", "Break in Class &Module", "Break on &Unhandled Errors")):
            rb = QRadioButton(text)
            rb.setChecked(n == cur)
            self.trap.append(rb)
            self.trap_group.addButton(rb, n)
            tl.addWidget(rb)
        f.addRow(tb)
        # Docking: the tool windows that dock (VB6: all but the Object Browser)
        f = page("Docking")
        db = QGroupBox("Dockable")
        dl = QVBoxLayout(db)
        self.dock_boxes = {}
        for key, title in getattr(main, "TOOL_WINDOWS", []):
            cb = QCheckBox(title)
            cb.setChecked(main.is_dockable(key))
            self.dock_boxes[key] = cb
            dl.addWidget(cb)
        f.addRow(db)
        # Environment: what LinBASIC Studio needs to know
        f = page("Environment")
        self.theme = QComboBox()
        self.theme.addItems(["system", "light", "dark"])     # system: as Ubuntu / Windows is
        self.theme.setCurrentText(S["dark_mode"] if S["dark_mode"] in ("system", "light", "dark") else "system")
        f.addRow("Theme:", self.theme)
        self.cp = QComboBox()
        self.cp.setEditable(True)
        self.cp.addItems(["", "cp1252", "cp1253", "cp1250", "cp1251", "cp1254", "cp1255", "cp1256", "cp1257", "cp932", "cp936"])
        self.cp.setCurrentText(S["codepage"])
        self.cp.setToolTip(f"Empty: from the language of the system ({vbmodel.default_codepage()})")
        f.addRow("Code page of VB6 files:", self.cp)
        row = QHBoxLayout()
        self.msdn = QLineEdit(S["msdn_path"])
        b = QPushButton("...")
        b.clicked.connect(self._browse)
        row.addWidget(self.msdn)
        row.addWidget(b)
        f.addRow("MSDN Library (your own copy):", row)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel |
                              QDialogButtonBox.StandardButton.Help)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        bb.helpRequested.connect(lambda: main.help_keyword("Options dialog box") if hasattr(main, "help_keyword") else None)
        outer.addWidget(bb)

    def _browse(self):
        path = QFileDialog.getExistingDirectory(self, "MSDN Library folder (the .CHM files)", self.msdn.text() or os.path.expanduser("~"))
        if path:
            self.msdn.setText(path)

    def _ok(self):
        S = self.main.S
        S["dark_mode"] = self.theme.currentText()
        S["codepage"] = self.cp.currentText().strip()
        S["msdn_path"] = self.msdn.text().strip()
        S["auto_list_members"] = self.auto_list.isChecked()
        S["auto_quick_info"] = self.auto_info.isChecked()
        S["error_trapping"] = next(n for n, rb in enumerate(self.trap) if rb.isChecked())
        S["show_grid"] = self.show_grid.isChecked()
        S["grid_width"] = self.grid_w.value()
        S["grid_height"] = self.grid_h.value()
        S["align_grid"] = self.align_grid.isChecked()
        dock = dict(S["dockable"] or {})
        for key, cb in self.dock_boxes.items():
            if key == "layout":
                S["layout_dockable"] = cb.isChecked()
            else:
                dock[key] = cb.isChecked()
        S["dockable"] = dock
        self.accept()


class AdvancedOptimizationsDialog(QDialog):
    """Compile > Advanced Optimizations... (VB6, DEV-8/pp1 stage 2): the warning and six check boxes"""

    def __init__(self, parent, values, help=None):
        super().__init__(parent)
        from . import projprops
        self.setWindowTitle("Advanced Optimizations")
        lay = QVBoxLayout(self)
        warn = QLabel("Warning: enabling the following optimizations may prevent\ncorrect execution of your program")
        lay.addWidget(warn)
        self.boxes = {}
        for text, key in projprops.ADVANCED:
            cb = QCheckBox(text)
            cb.setChecked(bool(values.get(key)))
            self.boxes[key] = cb
            lay.addWidget(cb)
        row = QHBoxLayout()
        for text, slot in (("OK", self.accept), ("Cancel", self.reject), ("Help", help)):
            b = QPushButton(text)
            if slot is not None:
                b.clicked.connect(slot)
            if text == "OK":
                b.setDefault(True)
            row.addWidget(b)
        lay.addLayout(row)

    def values(self):
        return {k: cb.isChecked() for k, cb in self.boxes.items()}


class ProjectPropertiesDialog(QDialog):
    """Project > <Project> Properties (VB6, DEV-8/pp1): General, Make, Compile, Component, Debugging; OK, Cancel,
    Help. It opens on General; what is on follows the Project Type (projprops.enabled), the lines it writes follow
    the pages that were shown (projprops.write)"""
    PAGES = ["general", "make", "compile", "component", "debugging"]

    def __init__(self, main, project):
        super().__init__(main)
        import copy
        from PyQt6.QtWidgets import QTabWidget, QListWidget, QButtonGroup
        from . import projprops as P
        self.P = P
        self.main, self.p = main, project
        self.messages = []                      # (kind, text) of the messages shown, for the tests
        mods = [(m.kind, m.name) for m in project.modules]
        self.s = P.read(project.vbp, mods, project.path)
        self.old = copy.deepcopy(self.s)
        self.visited = {"general"}
        self._busy = False
        self.setWindowTitle(f"{project.name} - Project Properties")
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs)
        # ---------------------------------------------------------------- General
        g = QWidget()
        gl = QGridLayout(g)
        gl.addWidget(QLabel("Project Type:"), 0, 0)
        gl.addWidget(QLabel("Startup Object:"), 0, 1)
        self.ptype = QComboBox()
        for key, text in P.TYPES:
            self.ptype.addItem(text, key)
        gl.addWidget(self.ptype, 1, 0)
        self.startup = QComboBox()
        gl.addWidget(self.startup, 1, 1)
        gl.addWidget(QLabel("Project Name:"), 2, 0)
        self.pname = QLineEdit(self.s["name"])
        gl.addWidget(self.pname, 3, 0)
        gl.addWidget(QLabel("Help File Name:"), 4, 0)
        gl.addWidget(QLabel("Project Help\nContext ID:"), 4, 1)
        hr = QHBoxLayout()
        self.helpfile = QLineEdit(self.s["helpfile"])
        hr.addWidget(self.helpfile)
        hb = QPushButton("...")
        hb.setFixedWidth(28)
        hb.clicked.connect(self._browse_help)
        hr.addWidget(hb)
        gl.addLayout(hr, 5, 0)
        self.helpctx = QLineEdit(str(self.s["helpctx"]))
        gl.addWidget(self.helpctx, 5, 1)
        gl.addWidget(QLabel("Project Description:"), 6, 0)
        self.desc = QLineEdit(self.s["desc"])
        gl.addWidget(self.desc, 7, 0, 1, 2)
        cl = QVBoxLayout()
        self.unattended = QCheckBox("Unattended Execution")
        self.upgrade = QCheckBox("Upgrade ActiveX Controls")
        self.license = QCheckBox("Require License Key")
        self.retained = QCheckBox("Retained In Memory")
        for cb, key in ((self.unattended, "unattended"), (self.upgrade, "upgrade"), (self.license, "license"),
                        (self.retained, "retained")):
            cb.setChecked(bool(self.s[key]))
            cl.addWidget(cb)
        gl.addLayout(cl, 8, 0)
        tg = QGroupBox("Threading Model")
        tl = QGridLayout(tg)
        self.threading = QComboBox()
        self.threading.addItem("Apartment Threaded", 1)
        self.threading.addItem("Single Threaded", 0)
        tl.addWidget(self.threading, 0, 0, 1, 3)
        self.per_object = QRadioButton("Thread per Object")
        self.pool = QRadioButton("Thread Pool")
        self.thread_group = QButtonGroup(self)
        self.thread_group.addButton(self.per_object)
        self.thread_group.addButton(self.pool)
        (self.per_object if self.s["per_object"] else self.pool).setChecked(True)
        tl.addWidget(self.per_object, 1, 0, 1, 3)
        tl.addWidget(self.pool, 2, 0)
        self.threads = QSpinBox()
        self.threads.setRange(1, 9999)
        self.threads.setValue(self.s["threads"])
        tl.addWidget(self.threads, 2, 1)
        tl.addWidget(QLabel("threads"), 2, 2)
        gl.addWidget(tg, 8, 1)
        gl.setRowStretch(9, 1)
        self.tabs.addTab(g, "General")
        # ---------------------------------------------------------------- Make
        mk = QWidget()
        ml = QGridLayout(mk)
        vg = QGroupBox("Version Number")
        vl = QGridLayout(vg)
        self.major, self.minor, self.revision = QSpinBox(), QSpinBox(), QSpinBox()
        for i, (sb, text, key) in enumerate(((self.major, "Major:", "major"), (self.minor, "Minor:", "minor"),
                                             (self.revision, "Revision:", "revision"))):
            sb.setRange(0, 9999)
            sb.setValue(self.s[key])
            vl.addWidget(QLabel(text), 0, i)
            vl.addWidget(sb, 1, i)
        self.autoinc = QCheckBox("Auto Increment")
        self.autoinc.setChecked(self.s["autoinc"])
        vl.addWidget(self.autoinc, 2, 0, 1, 3)
        ml.addWidget(vg, 0, 0)
        ag = QGroupBox("Application")
        al = QGridLayout(ag)
        al.addWidget(QLabel("Title:"), 0, 0)
        self.title = QLineEdit(self.s["title"])
        al.addWidget(self.title, 0, 1, 1, 2)
        al.addWidget(QLabel("Icon:"), 1, 0)
        self.icon = QComboBox()
        self.icon.addItems(self.s["forms"])
        if self.s["iconform"] in self.s["forms"]:
            self.icon.setCurrentText(self.s["iconform"])
        self.icon.setEnabled(bool(self.s["forms"]))
        al.addWidget(self.icon, 1, 1)
        self.icon_view = QLabel()
        self.icon_view.setFixedSize(32, 32)
        al.addWidget(self.icon_view, 1, 2)
        self.icon.currentTextChanged.connect(self._show_icon)
        self._show_icon()
        ml.addWidget(ag, 0, 1)
        ig = QGroupBox("Version Information")
        il = QGridLayout(ig)
        il.addWidget(QLabel("Type:"), 0, 0)
        il.addWidget(QLabel("Value:"), 0, 1)
        self.info_list = QListWidget()
        self.info_list.addItems([t for t, _ in P.VERSION_INFO])
        self.info_list.setMaximumHeight(60)
        il.addWidget(self.info_list, 1, 0)
        self.info_value = QPlainTextEdit()
        self.info_value.setMaximumHeight(60)
        il.addWidget(self.info_value, 1, 1)
        ml.addWidget(ig, 1, 0, 1, 2)
        ml.addWidget(QLabel("Command Line Arguments:"), 2, 0)
        self.command = QLineEdit(self.s["command"])
        ml.addWidget(self.command, 2, 1)
        ml.addWidget(QLabel("Conditional Compilation Arguments:"), 3, 0)
        self.condcomp = QLineEdit(self.s["condcomp"])
        ml.addWidget(self.condcomp, 3, 1)
        self.remove_unused = QCheckBox("Remove information about unused ActiveX Controls")
        self.remove_unused.setChecked(self.s["remove_unused"])
        ml.addWidget(self.remove_unused, 4, 0, 1, 2)
        ml.setRowStretch(5, 1)
        self.tabs.addTab(mk, "Make")
        self.info_list.currentRowChanged.connect(self._info_row)
        self.info_value.textChanged.connect(self._info_text)
        self.info_list.setCurrentRow(0)
        # ---------------------------------------------------------------- Compile
        cp = QWidget()
        cpl = QVBoxLayout(cp)
        self.pcode = QRadioButton("Compile to P-Code")
        self.native = QRadioButton("Compile to Native Code")
        (self.pcode if self.s["pcode"] else self.native).setChecked(True)
        cpl.addWidget(self.pcode)
        cpl.addWidget(self.native)
        ng = QGroupBox()
        nl = QGridLayout(ng)
        self.fast = QRadioButton("Optimize for Fast Code")
        self.small = QRadioButton("Optimize for Small Code")
        self.noopt = QRadioButton("No Optimization")
        self.opt_group = QButtonGroup(self)
        for i, rb in enumerate((self.fast, self.small, self.noopt)):
            self.opt_group.addButton(rb, i)
            nl.addWidget(rb, i, 0)
        (self.fast, self.small, self.noopt)[min(max(self.s["opt"], 0), 2)].setChecked(True)
        self.pentium = QCheckBox("Favor Pentium Pro(tm)")
        self.pentium.setChecked(self.s["pentium"])
        self.symdebug = QCheckBox("Create Symbolic Debug Info")
        self.symdebug.setChecked(self.s["symdebug"])
        nl.addWidget(self.pentium, 0, 1)
        nl.addWidget(self.symdebug, 1, 1)
        self.adv_btn = QPushButton("Advanced Optimizations...")
        self.adv_btn.clicked.connect(self.advanced_optimizations)
        nl.addWidget(self.adv_btn, 3, 0)
        self.native_group = ng
        cpl.addWidget(ng)
        dr = QHBoxLayout()
        self.dllbase_label = QLabel("DLL Base Address:")
        dr.addWidget(self.dllbase_label)
        self.dllbase = QLineEdit(self.s["dllbase"])
        dr.addWidget(self.dllbase)
        dr.addStretch()
        cpl.addLayout(dr)
        cpl.addStretch()
        self.tabs.addTab(cp, "Compile")
        # ---------------------------------------------------------------- Component
        co = QWidget()
        col = QGridLayout(co)
        sg = QGroupBox("Start Mode")
        sl = QVBoxLayout(sg)
        self.standalone = QRadioButton("Standalone")
        self.activex = QRadioButton("ActiveX Component")
        self.start_group = QButtonGroup(self)
        self.start_group.addButton(self.standalone, 0)
        self.start_group.addButton(self.activex, 1)
        sl.addWidget(self.standalone)
        sl.addWidget(self.activex)
        col.addWidget(sg, 0, 0)
        rg = QGroupBox("Remote Server")
        rl = QVBoxLayout(rg)
        self.remote = QCheckBox("Remote Server Files")
        self.remote.setChecked(self.s["remote"])
        rl.addWidget(self.remote)
        col.addWidget(rg, 1, 0)
        vcg = QGroupBox("Version Compatibility")
        vcl = QGridLayout(vcg)
        self.no_compat = QRadioButton("No Compatibility")
        self.proj_compat = QRadioButton("Project Compatibility")
        self.bin_compat = QRadioButton("Binary Compatibility")
        self.compat_group = QButtonGroup(self)
        for i, rb in enumerate((self.no_compat, self.proj_compat, self.bin_compat)):
            self.compat_group.addButton(rb, i)
            vcl.addWidget(rb, i, 0, 1, 2)
        self.compat_path = QLineEdit(self.s["compat_path"])
        vcl.addWidget(self.compat_path, 3, 0)
        self.compat_browse = QPushButton("...")
        self.compat_browse.setFixedWidth(28)
        self.compat_browse.clicked.connect(self._browse_compat)
        vcl.addWidget(self.compat_browse, 3, 1)
        self.compat_box = vcg
        col.addWidget(vcg, 2, 0, 1, 2)
        col.setRowStretch(3, 1)
        col.setColumnStretch(1, 1)
        self.tabs.addTab(co, "Component")
        # ---------------------------------------------------------------- Debugging
        db = QWidget()
        dl = QVBoxLayout(db)
        wg = QGroupBox("When this project starts")
        wl = QGridLayout(wg)
        self.dbg_wait = QRadioButton("Wait for components to be created")
        self.dbg_comp = QRadioButton("Start component:")
        self.dbg_prog = QRadioButton("Start program:")
        self.dbg_url = QRadioButton("Start browser with URL:")
        self.debug_group = QButtonGroup(self)
        for i, rb in enumerate((self.dbg_wait, self.dbg_comp, self.dbg_prog, self.dbg_url)):
            self.debug_group.addButton(rb, i)
        wl.addWidget(self.dbg_wait, 0, 0, 1, 3)
        wl.addWidget(self.dbg_comp, 1, 0)
        self.dbg_comp_list = QComboBox()
        self.dbg_comp_list.addItems(self.s["components"])
        if self.s["debug_component"] in self.s["components"]:
            self.dbg_comp_list.setCurrentText(self.s["debug_component"])
        wl.addWidget(self.dbg_comp_list, 1, 1, 1, 2)
        wl.addWidget(self.dbg_prog, 2, 0, 1, 3)
        self.dbg_prog_path = QLineEdit(self.s["debug_program"])
        wl.addWidget(self.dbg_prog_path, 3, 0, 1, 2)
        self.dbg_prog_browse = QPushButton("...")
        self.dbg_prog_browse.setFixedWidth(28)
        self.dbg_prog_browse.clicked.connect(self._browse_program)
        wl.addWidget(self.dbg_prog_browse, 3, 2)
        wl.addWidget(self.dbg_url, 4, 0, 1, 3)
        self.dbg_url_text = QLineEdit(self.s["debug_url"])
        wl.addWidget(self.dbg_url_text, 5, 0, 1, 3)
        self.use_existing = QCheckBox("Use existing browser")
        self.use_existing.setChecked(self.s["use_existing"])
        wl.addWidget(self.use_existing, 6, 0, 1, 3)
        self.debug_box = wg
        dl.addWidget(wg)
        dl.addStretch()
        self.tabs.addTab(db, "Debugging")
        # ---------------------------------------------------------------- buttons
        row = QHBoxLayout()
        row.addStretch()
        self.ok_btn = QPushButton("OK")
        self.ok_btn.setDefault(True)
        self.ok_btn.clicked.connect(self._ok)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        self.help_btn = QPushButton("Help")
        if hasattr(main, "help_contents"):
            self.help_btn.clicked.connect(lambda: main.help_contents())
        for b in (self.ok_btn, cancel, self.help_btn):
            row.addWidget(b)
        lay.addLayout(row)
        # ---------------------------------------------------------------- values and what is on
        self.ptype.setCurrentIndex(max(0, self.ptype.findData(self.s["type"])))
        self._show_type_values()
        self.tabs.currentChanged.connect(lambda i: self.visited.add(self.PAGES[i]))
        self.ptype.currentIndexChanged.connect(self._type_changed)
        for w in (self.unattended, self.pcode, self.native, self.per_object, self.pool, self.no_compat,
                  self.proj_compat, self.bin_compat, self.dbg_wait, self.dbg_comp, self.dbg_prog, self.dbg_url):
            w.toggled.connect(self._refresh)
        self.tabs.setCurrentIndex(0)
        self._refresh()

    # ------------------------------------------------------------------ the values of the Project Type
    def _show_type_values(self):
        """Startup Object, Start Mode, Version Compatibility, Debugging as the values are (a new Project Type
        changes them)"""
        prev = self._busy
        self._busy = True
        try:
            items = self.P.startup_items(self.s["type"], self.s["forms"])
            self.startup.clear()
            self.startup.addItems(items)
            self.startup.setCurrentText(self.s["startup"] if self.s["startup"] in items else "Sub Main")
            (self.activex if self.s["startmode"] else self.standalone).setChecked(True)
            (self.no_compat, self.proj_compat, self.bin_compat)[min(max(self.s["compat"], 0), 2)].setChecked(True)
            (self.dbg_wait, self.dbg_comp, self.dbg_prog, self.dbg_url)[min(max(self.s["debug"], 0), 3)].setChecked(True)
            self.threading.setCurrentIndex(self.threading.findData(self.s["threading"]))
        finally:
            self._busy = prev

    def _type_changed(self):
        if self._busy:
            return
        self.collect()
        self.P.change_type(self.s, self.ptype.currentData())
        if self.s["debug"] == 1 and not self.s["components"]:
            self.s["debug"] = 0
        self._show_type_values()
        self._refresh()

    def _refresh(self):
        if self._busy:
            return
        self.collect()
        on = self.P.enabled(self.s)
        self.unattended.setEnabled(on["unattended"])
        self.license.setEnabled(on["license"])
        self.retained.setEnabled(on["retained"])
        self.threading.setEnabled(on["threading"])
        if on["threading"]:
            self.threading.setCurrentIndex(self.threading.findData(self.s["threading"]))
        else:
            self.threading.setCurrentIndex(-1)  # an EXE: the list is empty and grey
        self.per_object.setEnabled(on["thread_mode"])
        self.pool.setEnabled(on["thread_mode"])
        self.threads.setEnabled(on["threads"])
        self.native_group.setEnabled(on["native"])
        self.dllbase.setEnabled(on["dllbase"])
        self.dllbase_label.setEnabled(on["dllbase"])
        self.standalone.setEnabled(on["startmode"])
        self.activex.setEnabled(on["startmode"])
        self.remote.setEnabled(on["remote"])
        self.compat_box.setEnabled(on["compat"])
        self.compat_path.setEnabled(on["compat_path"])
        self.compat_browse.setEnabled(on["compat_path"])
        self.debug_box.setEnabled(on["debugging"])
        self.dbg_comp.setEnabled(on["start_component"])
        self.dbg_comp_list.setEnabled(on["start_component"] and self.dbg_comp.isChecked())
        self.dbg_prog_path.setEnabled(self.dbg_prog.isChecked())
        self.dbg_prog_browse.setEnabled(self.dbg_prog.isChecked())
        self.dbg_url_text.setEnabled(self.dbg_url.isChecked())

    # ------------------------------------------------------------------ Make: the icon of the form
    def _show_icon(self):
        """the Icon of the form chosen (from its .frx), or the icon VB6 gives a form that has none"""
        from PyQt6.QtGui import QPixmap, QPainter, QColor
        name = self.icon.currentText()
        pm = None
        m = self.p.module(name) if name else None
        if m is not None and m.sf is not None and m.sf.root is not None:
            data = m.frx_value(m.sf.root, "Icon")
            if isinstance(data, (bytes, bytearray)):
                pm = QPixmap()
                if not pm.loadFromData(bytes(data)):
                    pm = None
        if pm is None and name:
            pm = QPixmap(32, 32)
            pm.fill(QColor(0, 0, 0, 0))
            g = QPainter(pm)
            g.fillRect(3, 5, 26, 22, QColor("#FFFFFF"))
            g.fillRect(3, 5, 26, 6, QColor("#1C5FB0"))
            g.setPen(QColor("#000000"))
            g.drawRect(3, 5, 25, 21)
            g.end()
        self.icon_view.setPixmap(pm.scaled(32, 32) if pm is not None else QPixmap())

    # ------------------------------------------------------------------ Make: the Version Information
    def _info_row(self, row):
        if row < 0:
            return
        key = self.P.VERSION_INFO[row][1]
        self._busy_info = True
        self.info_value.setPlainText(self.s["info"][key])
        self._busy_info = False

    def _info_text(self):
        if getattr(self, "_busy_info", False):
            return
        row = self.info_list.currentRow()
        if row >= 0:
            self.s["info"][self.P.VERSION_INFO[row][1]] = self.info_value.toPlainText()

    # ------------------------------------------------------------------ the values of the dialog
    def collect(self):
        """the widgets into the values (projprops)"""
        s = self.s
        if self.startup.currentText():
            s["startup"] = self.startup.currentText()
        s["name"] = self.pname.text()
        s["helpfile"] = self.helpfile.text()
        s["helpctx"] = self.helpctx.text().strip() or "0"
        s["desc"] = self.desc.text()
        s["unattended"] = self.unattended.isChecked() and self.unattended.isEnabled()
        s["upgrade"] = self.upgrade.isChecked()
        s["license"] = self.license.isChecked()
        s["retained"] = self.retained.isChecked() and self.unattended.isChecked()
        if self.threading.currentIndex() >= 0:
            s["threading"] = self.threading.currentData()
        s["per_object"] = self.per_object.isChecked()
        s["threads"] = self.threads.value()
        s["major"], s["minor"], s["revision"] = self.major.value(), self.minor.value(), self.revision.value()
        s["autoinc"] = self.autoinc.isChecked()
        s["title"] = self.title.text()
        if self.icon.currentText():
            if s["iconform"] != self.icon.currentText():
                s["iconform_set"] = True
            s["iconform"] = self.icon.currentText()
        s["command"] = self.command.text()
        s["condcomp"] = self.condcomp.text()
        s["remove_unused"] = self.remove_unused.isChecked()
        s["pcode"] = self.pcode.isChecked()
        s["opt"] = max(0, self.opt_group.checkedId())
        s["pentium"] = self.pentium.isChecked()
        s["symdebug"] = self.symdebug.isChecked()
        s["dllbase"] = self.dllbase.text()
        s["startmode"] = 1 if self.activex.isChecked() else 0
        s["remote"] = self.remote.isChecked()
        s["compat"] = max(0, self.compat_group.checkedId())
        s["compat_path"] = self.compat_path.text().strip()
        s["debug"] = max(0, self.debug_group.checkedId())
        s["debug_component"] = self.dbg_comp_list.currentText()
        s["debug_program"] = self.dbg_prog_path.text().strip()
        s["debug_url"] = self.dbg_url_text.text().strip()
        s["use_existing"] = self.use_existing.isChecked()
        return s

    def advanced_optimizations(self):
        dlg = AdvancedOptimizationsDialog(self, self.s["advanced"],
                                          getattr(self.main, "help_contents", None))
        if dlg.exec():
            self.s["advanced"] = dlg.values()

    # ------------------------------------------------------------------ browse
    def _dir(self):
        return os.path.dirname(self.p.path) if self.p.path else os.path.expanduser("~")

    def _browse_help(self):
        f, _ = QFileDialog.getOpenFileName(self, "Help File", self._dir(), "Help Files (*.hlp *.chm);;All Files (*)")
        if f:
            self.helpfile.setText(self._rel(f))

    def _browse_compat(self):
        f, _ = QFileDialog.getOpenFileName(self, "Compatible ActiveX Component", self._dir(),
                                           "ActiveX Components (*.exe *.dll *.ocx);;All Files (*)")
        if f:
            self.compat_path.setText(self._rel(f))

    def _browse_program(self):
        f, _ = QFileDialog.getOpenFileName(self, "Start Program", self._dir(), "All Files (*)")
        if f:
            self.dbg_prog_path.setText(f)

    def _rel(self, f):
        """a file in the folder of the project by its name (VB6: CompatibleEXE32="ocx4.ocx")"""
        d = self._dir()
        return os.path.relpath(f, d) if os.path.dirname(os.path.abspath(f)) == os.path.abspath(d) else f

    # ------------------------------------------------------------------ OK
    def message(self, kind, text):
        self.messages.append((kind, text))
        (QMessageBox.critical if kind == "error" else QMessageBox.warning)(self, APP_NAME, text)

    def _ok(self):
        P = self.P
        s = self.collect()
        bad = P.compat_problem(s, os.path.dirname(self.p.path) if self.p.path else None)
        changed = (s["compat"], s["compat_path"]) != (self.old["compat"], self.old["compat_path"])
        if bad is not None and changed:
            self.message("error", P.compat_error(bad))
            self.tabs.setCurrentIndex(self.PAGES.index("component"))
            return False
        start_mode = P.write(self.p.vbp, self.old, s, self.visited)
        if start_mode:
            self.message("warning", P.START_MODE_MSG)
        self.p.dirty = True
        self.accept()
        return True


class MakeDialog(QDialog):
    """File > Make <Project>: the executable and its system (Linux, Windows 32-bit, Windows 64-bit).
    A Windows program goes into a folder of its own with the DLLs it needs (portable)."""
    TARGETS = [("linux", "Linux x86_64 (executable)"), ("win32", "Windows 32-bit (.exe)"),
               ("win64", "Windows 64-bit (.exe)")]

    @classmethod
    def targets(cls):
        """the systems this IDE can make programs for: on Windows the ones whose runtime and Tiny C
        Compiler it carries (runtime/<target>/dll, tcc); on Linux all of them (mingw-w64 for Windows)"""
        if os.name != "nt":
            return cls.TARGETS
        from . import paths
        return [(k, t) for k, t in cls.TARGETS if k != "linux" and
                os.path.exists(os.path.join(paths.RUNTIME, k, "tcc", "tcc.exe"))] or \
            [(k, t) for k, t in cls.TARGETS if k == paths.host_target()]

    NOT_HERE = ("Programs for Windows need the Windows runtime of LinBASIC Studio (runtime/win32, runtime/win64: "
                "its Tiny C Compiler, DLLs and Qt for Windows) or mingw-w64; this LinBASIC Studio has neither.")

    @classmethod
    def available(cls, target):
        """(can this IDE make programs for target?, why not): on Linux a Windows program needs the Tiny C Compiler
        of LinBASIC Studio for Windows with its DLLs (runtime/<target>/tcc/cross-tcc, dll/: the AppImage carries
        them), or mingw-w64 and the Windows runtime (runtime/<target>, or build-win.sh to make it: the source);
        without them shown, but disabled (the user's decision, 2026-11-19)"""
        if os.name == "nt" or target == "linux":
            return True, ""
        import shutil
        from . import paths
        tri = {"win64": "x86_64-w64-mingw32", "win32": "i686-w64-mingw32"}.get(target, "")
        cc = any(shutil.which(f"{tri}-gcc{s}") for s in ("-posix", ""))
        kit = os.path.exists(os.path.join(paths.RUNTIME, target, "libvbrt.a")) or \
            os.path.exists(os.path.join(paths.RUNTIME, "build-win.sh"))
        tcc = os.path.exists(os.path.join(paths.RUNTIME, target, "tcc", "cross-tcc")) and \
            os.path.exists(os.path.join(paths.RUNTIME, target, "dll", "vbrt.dll"))
        ok = (cc and kit) or tcc
        return ok, ("" if ok else cls.NOT_HERE)

    def __init__(self, parent, project_dir, exe_name, target="linux"):
        from PyQt6.QtWidgets import QRadioButton, QButtonGroup, QHBoxLayout, QVBoxLayout, QGroupBox, QDialogButtonBox
        super().__init__(parent)
        self.setWindowTitle("Make Project")
        self.project_dir, self.exe_name = project_dir, exe_name
        lay = QVBoxLayout(self)
        box = QGroupBox("Target system")
        bl = QVBoxLayout(box)
        self.group = QButtonGroup(self)
        self.buttons = {}
        for key, text in self.targets():
            rb = QRadioButton(text)
            self.group.addButton(rb)
            self.buttons[key] = rb
            bl.addWidget(rb)
            rb.toggled.connect(lambda on, k=key: on and self._target_changed(k))
            ok, why = self.available(key)
            rb.setEnabled(ok)                   # the AppImage: Windows shown, disabled
            rb.setToolTip(why)
        lay.addWidget(box)
        lay.addWidget(QLabel("File name:"))
        row = QHBoxLayout()
        self.path = QLineEdit()
        self.path.setMinimumWidth(420)
        row.addWidget(self.path, 1)
        b = QPushButton("Browse...")
        b.clicked.connect(self._browse)
        row.addWidget(b)
        lay.addLayout(row)
        self.note = QLabel()
        self.note.setWordWrap(True)
        lay.addWidget(self.note)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        if target not in self.buttons or not self.buttons[target].isEnabled():
            target = next(k for k, b in self.buttons.items() if b.isEnabled())
        self.buttons[target].setChecked(True)
        self._target_changed(target)

    def default_path(self, target):
        if target == "linux":
            return os.path.join(self.project_dir, self.exe_name)
        return os.path.join(self.project_dir, f"{self.exe_name}-{target}", self.exe_name + ".exe")

    def _target_changed(self, target):
        self.target_key = target
        cur = self.path.text().strip()
        if not cur or cur in {self.default_path(k) for k, _ in self.targets()}:
            self.path.setText(self.default_path(target))
        if target == "linux":
            self.note.setText("A native program for Linux (x86_64). It needs Qt 6 on the computer that runs it "
                              "when it has forms.")
        else:
            self.note.setText("A portable Windows program: the .exe and the DLLs it needs (Qt for the forms) "
                              "go into its own folder; copy the whole folder." +
                              (" Built by the Tiny C Compiler of LinBASIC Studio." if os.name == "nt" else
                               " Needs mingw-w64 on this computer: sudo apt install mingw-w64"))

    def _browse(self):
        from PyQt6.QtWidgets import QFileDialog
        flt = "Executable (*)" if self.target_key == "linux" else "Windows program (*.exe)"
        p, _ = QFileDialog.getSaveFileName(self, "Make Project", self.path.text(), flt)
        if p:
            if self.target_key != "linux" and not p.lower().endswith(".exe"):
                p += ".exe"
            self.path.setText(p)

    def result_values(self):
        p = self.path.text().strip()
        if self.target_key != "linux" and not p.lower().endswith(".exe"):
            p += ".exe"
        return self.target_key, p


class ComponentsDialog(QDialog):
    """Project > Components: the control libraries (.ocx) of the project, as VB6 lists them"""

    def __init__(self, main, project):
        super().__init__(main)
        from PyQt6.QtWidgets import QListWidget, QListWidgetItem
        from PyQt6.QtCore import Qt
        from .ocxlib import LIBS
        self.project = project
        self.setWindowTitle("Components")
        v = QVBoxLayout(self)
        v.addWidget(QLabel("Controls"))
        self.list = QListWidget()
        self.list.setMinimumSize(420, 260)
        have, self.used = set(project.components()), project.used_components()
        for lib in LIBS:
            it = QListWidgetItem(lib.title)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Checked if lib.key in have else Qt.CheckState.Unchecked)
            it.setData(Qt.ItemDataRole.UserRole, lib.key)
            self.list.addItem(it)
        v.addWidget(self.list)
        self.where = QLabel()
        v.addWidget(self.where)
        self.list.currentItemChanged.connect(self._show)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.list.setCurrentRow(0)

    def _show(self, it, _prev=None):
        from PyQt6.QtCore import Qt
        from .ocxlib import LIB_BY_KEY
        if it is None:
            return
        lib = LIB_BY_KEY[it.data(Qt.ItemDataRole.UserRole).lower()]
        self.where.setText(f"Location: {lib.file}  -  {', '.join(c.cls for c in lib.controls)}")

    def checked(self):
        from PyQt6.QtCore import Qt
        return [self.list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.list.count())
                if self.list.item(i).checkState() == Qt.CheckState.Checked]

    def _ok(self):
        from .ocxlib import LIB_BY_KEY
        keys = self.checked()
        busy = [k for k in self.used if k not in keys]
        if busy:
            QMessageBox.warning(self, "Components",
                                f"Can't remove control or reference; in use: {LIB_BY_KEY[busy[0].lower()].title}")
            return
        self.project.set_components(keys)
        self.accept()


class NewProjectDialog(QDialog):
    """File > New Project: the type of the project (VB6's New Project dialog)"""

    def __init__(self, parent, title="New Project", existing=False):
        from PyQt6.QtWidgets import QListWidget, QListWidgetItem, QVBoxLayout, QDialogButtonBox, QStyle
        from PyQt6.QtCore import QSize, Qt
        super().__init__(parent)
        self.setWindowTitle(title)
        self.existing_path = None               # Add Project: an existing .vbp instead
        lay = QVBoxLayout(self)
        self.list = QListWidget()
        self.list.setViewMode(QListWidget.ViewMode.IconMode)
        self.list.setIconSize(QSize(32, 32))
        self.list.setGridSize(QSize(120, 80))
        self.list.setWordWrap(True)             # "VB Application Wizard" on two lines, as in VB6 (aw1-01)
        self.list.setMovement(QListWidget.Movement.Static)
        self.list.setMinimumSize(520, 120)
        icons = {"Exe": QStyle.StandardPixmap.SP_ComputerIcon, "OleExe": QStyle.StandardPixmap.SP_DriveNetIcon,
                 "OleDll": QStyle.StandardPixmap.SP_FileIcon, "Control": QStyle.StandardPixmap.SP_DialogApplyButton}
        for key, text in vbmodel.PROJECT_TYPES:
            it = QListWidgetItem(self.style().standardIcon(icons[key]), text)
            it.setData(Qt.ItemDataRole.UserRole, key)
            self.list.addItem(it)
        if not existing:                        # VB6 (DEV-9/aw1-01): after ActiveX Control (not in Add Project)
            from .appwizard import wizard_icon
            it = QListWidgetItem(wizard_icon(32), "VB Application Wizard")
            it.setData(Qt.ItemDataRole.UserRole, "AppWizard")
            self.list.addItem(it)
        # an add-in of LinBASIC Studio (phase 12.5; VB6's Addin, a template of Projects: in Add Project too)
        it = QListWidgetItem(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView), "Addin")
        it.setData(Qt.ItemDataRole.UserRole, "Addin")
        self.list.addItem(it)
        self.list.setCurrentRow(0)
        self.list.itemDoubleClicked.connect(lambda _: self.accept())
        lay.addWidget(self.list)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        if existing:
            b = bb.addButton("&Existing Project...", QDialogButtonBox.ButtonRole.ActionRole)
            b.clicked.connect(self._existing)
        lay.addWidget(bb)

    def _existing(self):
        from PyQt6.QtWidgets import QFileDialog
        path, _ = QFileDialog.getOpenFileName(self, "Add Project", os.path.expanduser("~"), "Project Files (*.vbp)")
        if path:
            self.existing_path = path
            self.accept()

    def ptype(self):
        from PyQt6.QtCore import Qt
        it = self.list.currentItem()
        return it.data(Qt.ItemDataRole.UserRole) if it is not None else "Exe"


class ReferencesDialog(QDialog):
    """Project > References: the projects of the group and the registered components (Make) that the
    project uses; a reference is Reference=*\\A<path of the .vbp>, as VB6 writes it in a group"""

    def __init__(self, main, project, group):
        from PyQt6.QtWidgets import QListWidget, QListWidgetItem, QVBoxLayout, QDialogButtonBox
        from PyQt6.QtCore import Qt
        super().__init__(main)
        self.p, self.g = project, group
        self.setWindowTitle(f"References - {project.name}")
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Available References:"))
        self.list = QListWidget()
        self.list.setMinimumSize(460, 260)
        lay.addWidget(self.list)
        self.items = []
        have = {os.path.normcase(os.path.normpath(project.resolve(v.strip()[3:])))
                for k, v in project.vbp.entries
                if v is not None and k.strip() == "Reference" and v.strip().startswith("*\\A")}
        seen = set()
        cands = []
        for q in (group.projects if group is not None else []):
            if q is not project and q.path and q.ptype in ("OleDll", "OleExe", "Control"):
                cands.append((q.name, q.path, "project of the group"))
        try:
            import sys
            from . import paths
            if paths.COMPILER not in sys.path:
                sys.path.insert(0, paths.COMPILER)
            from lbvc.driver import load_registry
            for e in load_registry().values():
                if e.get("vbp") and os.path.exists(e["vbp"]):
                    cands.append((e.get("name", "?"), e["vbp"], "registered (Make)"))
        except Exception:                       # noqa: BLE001 - no registry: the group only
            pass
        for name, path, where in cands:
            key = os.path.normcase(os.path.normpath(path))
            if key in seen or (project.path and key == os.path.normcase(os.path.normpath(project.path))):
                continue
            seen.add(key)
            it = QListWidgetItem(f"{name}    ({where}: {path})")
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Checked if key in have else Qt.CheckState.Unchecked)
            self.list.addItem(it)
            self.items.append((it, path))
        if not self.items:
            lay.addWidget(QLabel("No ActiveX projects: add one to the group (File > Add Project) or Make one."))
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _ok(self):
        from PyQt6.QtCore import Qt
        p = self.p
        for it, path in self.items:
            on = it.checkState() == Qt.CheckState.Checked
            rel = os.path.relpath(path, os.path.dirname(p.path)).replace("/", "\\")
            line = "*\\A" + rel
            entries = [e for e in p.vbp.entries if not (e[1] is not None and e[0].strip() == "Reference" and
                       os.path.normcase(os.path.normpath(p.resolve(e[1].strip()[3:]))) == os.path.normcase(os.path.normpath(path))
                       and e[1].strip().startswith("*\\A"))]
            if on:
                last = max([i for i, (k, v) in enumerate(entries) if v is not None and k.strip() == "Reference"] or [0])
                entries.insert(last + 1, ["Reference", line])
            if entries != p.vbp.entries:
                p.vbp.entries = entries
                p.dirty = True
        self.accept()



class AddProcedureDialog(QDialog):
    """Tools > Add Procedure (VB6, DEV-8/ed1): Name, Type (Sub, Function, Property, Event), Scope (Public,
    Private), All Local variables as Statics. apply(name, kind, scope, statics) adds the procedure and gives None,
    or VB6's message: "Invalid procedure name" keeps the dialog open, "A procedure of that name already exists"
    closes it"""

    def __init__(self, parent, apply):
        super().__init__(parent)
        self.apply = apply
        self.setWindowTitle("Add Procedure")
        grid = QGridLayout(self)
        row = QHBoxLayout()
        lab = QLabel("&Name:")
        row.addWidget(lab)
        self.name = QLineEdit()
        lab.setBuddy(self.name)
        row.addWidget(self.name)
        grid.addLayout(row, 0, 0)
        btns = QVBoxLayout()
        self.ok = QPushButton("OK")
        self.ok.setDefault(True)
        self.ok.clicked.connect(self._ok)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        btns.addWidget(self.ok)
        btns.addWidget(cancel)
        btns.addStretch()
        grid.addLayout(btns, 0, 1, 3, 1)
        tg = QGroupBox("Type")
        tl = QGridLayout(tg)
        self.kinds = {}
        for k, (text, r, c) in {"Sub": ("&Sub", 0, 0), "Property": ("&Property", 0, 1),
                                "Function": ("&Function", 1, 0), "Event": ("&Event", 1, 1)}.items():
            rb = QRadioButton(text)
            self.kinds[k] = rb
            tl.addWidget(rb, r, c)
        self.kinds["Sub"].setChecked(True)
        grid.addWidget(tg, 1, 0)
        sg = QGroupBox("Scope")
        sl = QHBoxLayout(sg)
        self.public = QRadioButton("Pu&blic")
        self.private = QRadioButton("Pri&vate")
        self.public.setChecked(True)
        sl.addWidget(self.public)
        sl.addWidget(self.private)
        grid.addWidget(sg, 2, 0)
        self.statics = QCheckBox("&All Local variables as Statics")
        grid.addWidget(self.statics, 3, 0, 1, 2)
        self.name.setFocus()

    def kind(self):
        return next(k for k, rb in self.kinds.items() if rb.isChecked())

    def scope(self):
        return "Private" if self.private.isChecked() else "Public"

    def set_values(self, name, kind="Sub", scope="Public", statics=False):
        self.name.setText(name)
        self.kinds[kind].setChecked(True)
        (self.private if scope == "Private" else self.public).setChecked(True)
        self.statics.setChecked(statics)

    def _ok(self):
        err = self.apply(self.name.text().strip(), self.kind(), self.scope(), self.statics.isChecked())
        if err is None:
            self.accept()
            return
        QMessageBox.critical(self, APP_NAME, err)
        if err.startswith("Invalid"):
            self.name.setFocus()
            self.name.selectAll()
        else:
            self.reject()


class ProcedureAttributesDialog(QDialog):
    """Tools > Procedure Attributes (VB6, DEV-8/ed2): Name, Description, Project Help File, Help Context ID and, after
    Advanced >>, Procedure ID, Use this Page in Property Browser, Property Category, the Attributes and the Data
    Binding check boxes. procs: [(name, kind)] (kind Sub, Function or Property); read(name) -> attributes;
    apply(name, attributes) writes them (codeops.read_attributes / write_attributes). class_like: a class, a form
    or a control: the property options are on for a Property; in a standard module only Procedure ID and Hide this
    member are (as VB6)."""

    BINDING = [("bindable", "Property is data bound"), ("defaultbind", "This property binds to DataField"),
               ("displaybind", "Show in DataBindings collection at design time"),
               ("requestedit", "Property will call CanPropertyChange before changing"),
               ("immediatebind", "Update immediate")]

    def __init__(self, parent, procs, read, apply, start=None, help_file="", class_like=True, pages=()):
        super().__init__(parent)
        from . import codeops
        self.C = codeops
        self.procs, self.read, self.apply_fn = list(procs), read, apply
        self.class_like, self.pages = class_like, list(pages)
        self.cur = None
        self._loading = False
        self.setWindowTitle("Procedure Attributes")
        g = QGridLayout(self)
        top = QHBoxLayout()
        lab = QLabel("&Name:")
        top.addWidget(lab)
        self.name = QComboBox()
        lab.setBuddy(self.name)
        for n, _ in self.procs:
            self.name.addItem(n)
        top.addWidget(self.name, 1)
        g.addLayout(top, 0, 0, 1, 2)
        lab = QLabel("&Description:")
        g.addWidget(lab, 1, 0, 1, 2)
        self.desc = QPlainTextEdit()
        lab.setBuddy(self.desc)
        self.desc.setFixedHeight(48)
        g.addWidget(self.desc, 2, 0, 1, 2)
        g.addWidget(QLabel("Project Help File:"), 3, 0)
        lab_help = QLabel("&Help Context ID:")
        g.addWidget(lab_help, 3, 1)
        self.help_file = QLineEdit(help_file)
        self.help_file.setReadOnly(True)
        self.help_file.setEnabled(False)
        g.addWidget(self.help_file, 4, 0)
        self.helpid = QLineEdit("0")
        lab_help.setBuddy(self.helpid)
        g.addWidget(self.helpid, 4, 1)
        col = QVBoxLayout()
        self.ok = QPushButton("OK")
        self.ok.setDefault(True)
        self.ok.clicked.connect(self._ok)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        self.apply_btn = QPushButton("&Apply")
        self.apply_btn.setEnabled(False)
        self.apply_btn.clicked.connect(self._apply)
        self.adv_btn = QPushButton("Ad&vanced >>")
        self.adv_btn.clicked.connect(self.show_advanced)
        for b in (self.ok, cancel, self.apply_btn, self.adv_btn):
            col.addWidget(b)
        col.addStretch()
        g.addLayout(col, 0, 2, 5, 1)
        # Advanced
        self.adv = QWidget()
        a = QGridLayout(self.adv)
        a.setContentsMargins(0, 6, 0, 0)
        labs = [QLabel("Procedure &ID:"), QLabel("Use this &Page in\nProperty Browser:"), QLabel("Property &Category:")]
        for i, lab in enumerate(labs):
            a.addWidget(lab, 0, i, Qt.AlignmentFlag.AlignBottom)
        self.procid = QComboBox()
        self.page = QComboBox()
        self.page.addItems(["(None)"] + self.pages)
        self.category = QComboBox()
        for lab, w in zip(labs, (self.procid, self.page, self.category)):
            lab.setBuddy(w)                     # the & is a mnemonic, not shown
        self.category.setEditable(True)            # a category of one's own, too
        self.category.addItems(codeops.CATEGORIES)
        a.addWidget(self.procid, 1, 0)
        a.addWidget(self.page, 1, 1)
        a.addWidget(self.category, 1, 2)
        ag = QGroupBox("Attributes")
        al = QGridLayout(ag)
        self.flag_boxes = {"hide": QCheckBox("Hide this &member"), "uidefault": QCheckBox("&User Interface Default"),
                           "nobrowse": QCheckBox("Don't show in Property &Browser")}
        al.addWidget(self.flag_boxes["hide"], 0, 0)
        al.addWidget(self.flag_boxes["uidefault"], 0, 1)
        al.addWidget(self.flag_boxes["nobrowse"], 1, 0, 1, 2)
        a.addWidget(ag, 2, 0, 1, 3)
        bg = QGroupBox("Data Binding")
        bl = QVBoxLayout(bg)
        self.binding_group = bg
        for i, (k, text) in enumerate(self.BINDING):
            cb = QCheckBox(text)
            if i:
                cb.setContentsMargins(18, 0, 0, 0)
                cb.setStyleSheet("margin-left: 18px")
            self.flag_boxes[k] = cb
            bl.addWidget(cb)
        a.addWidget(bg, 3, 0, 1, 3)
        g.addWidget(self.adv, 5, 0, 1, 3)
        self.adv.hide()
        # changes enable Apply
        self.desc.textChanged.connect(self._changed)
        self.helpid.textChanged.connect(self._changed)
        for w in (self.procid, self.page, self.category):
            w.currentIndexChanged.connect(self._changed)
        self.category.editTextChanged.connect(self._changed)
        for cb in self.flag_boxes.values():
            cb.toggled.connect(self._changed)
        self.flag_boxes["bindable"].toggled.connect(self._enable)
        self.name.currentIndexChanged.connect(self._name_changed)
        i = next((k for k, (n, _) in enumerate(self.procs) if start and n.lower() == start.lower()), 0)
        self.name.setCurrentIndex(i)
        self._load(self.name.currentText())

    def show_advanced(self):
        self.adv.show()
        self.adv_btn.setEnabled(False)          # as VB6: the button goes grey

    def _kind(self, name):
        return next((k for n, k in self.procs if n == name), "Sub")

    def _enable(self):
        prop = self._kind(self.cur) == "Property" and self.class_like
        for w in (self.page, self.category, self.flag_boxes["uidefault"], self.flag_boxes["nobrowse"],
                  self.binding_group):
            w.setEnabled(prop)
        bound = prop and self.flag_boxes["bindable"].isChecked()
        for k, _ in self.BINDING[1:]:
            self.flag_boxes[k].setEnabled(bound)

    def _load(self, name):
        self._loading = True
        self.cur = name
        at = self.read(name)
        self.desc.setPlainText(at["description"])
        self.helpid.setText(str(at["helpid"]))
        ids = self.C.PROC_IDS_PROPERTY if self._kind(name) == "Property" else self.C.PROC_IDS_METHOD
        self.procid.clear()
        for text, v in ids:
            self.procid.addItem(text, v)
        k = next((i for i, (_, v) in enumerate(ids) if v == at["memid"]), -1)
        if k < 0 and at["memid"] is not None:   # a number of its own (from the file): kept
            self.procid.addItem(str(at["memid"]), at["memid"])
            k = self.procid.count() - 1
        self.procid.setCurrentIndex(max(0, k))
        self.page.setCurrentIndex(max(0, self.page.findText(at["page"])) if at["page"] else 0)
        self.category.setCurrentText(at["category"] or "(None)")
        for key, cb in self.flag_boxes.items():
            cb.setChecked(bool(at["flags"] & self.C.FLAGS[key]))
        self._enable()
        self._loading = False
        self.apply_btn.setEnabled(False)

    def _changed(self, *a):
        if not self._loading:
            self.apply_btn.setEnabled(True)

    def values(self):
        try:
            hid = int(self.helpid.text().strip() or "0")
        except ValueError:
            hid = 0
        flags = 0
        for key, cb in self.flag_boxes.items():
            if cb.isChecked() and cb.isEnabled():
                flags |= self.C.FLAGS[key]
        prop = self._kind(self.cur) == "Property" and self.class_like
        cat = self.category.currentText().strip() if prop else ""
        page = self.page.currentText() if prop else ""
        return {"description": self.desc.toPlainText().replace("\n", " ").strip(), "helpid": hid,
                "page": "" if page == "(None)" else page, "category": "" if cat == "(None)" else cat,
                "memid": self.procid.currentData(), "flags": flags}

    def _apply(self):
        if self.cur is not None and self.apply_btn.isEnabled():
            self.apply_fn(self.cur, self.values())
        self.apply_btn.setEnabled(False)

    def _name_changed(self, i):
        if self._loading or i < 0:
            return
        self._apply()                           # what was set for the previous procedure is kept
        self._load(self.name.itemText(i))

    def _ok(self):
        self._apply()
        self.accept()


class PrintDialog(QDialog):
    """File > Print (VB6, DEV-8/pr1): Printer, Range (Selection, Current Module, Current Project), Print What (Form
    Image, Code, Form As Text), Print Quality, Print to File; OK, Cancel, Setup..., Help. Selection is on (and
    chosen) when the code window has a selection; Form Image and Form As Text when a form is the active module
    or Current Project is chosen"""

    def __init__(self, parent, printer_name, has_selection=False, form_active=False, setup=None, help=None):
        super().__init__(parent)
        self.form_active = form_active
        self.setWindowTitle(f"Print - {parent.project.name}" if getattr(parent, "project", None) else "Print")
        g = QGridLayout(self)
        g.addWidget(QLabel("Printer:"), 0, 0)
        self.printer_label = QLabel(printer_name)
        g.addWidget(self.printer_label, 0, 1, 1, 2)
        rg = QGroupBox("Range")
        rl = QVBoxLayout(rg)
        self.r_sel = QRadioButton("&Selection")
        self.r_mod = QRadioButton("Current &Module")
        self.r_proj = QRadioButton("Current &Project")
        for rb in (self.r_sel, self.r_mod, self.r_proj):
            rl.addWidget(rb)
        self.r_sel.setEnabled(has_selection)
        (self.r_sel if has_selection else self.r_mod).setChecked(True)
        g.addWidget(rg, 1, 0, 1, 2)
        wg = QGroupBox("Print What")
        wl = QVBoxLayout(wg)
        self.c_image = QCheckBox("Form &Image")
        self.c_code = QCheckBox("&Code")
        self.c_text = QCheckBox("Form &As Text")
        self.c_code.setChecked(True)
        for cb in (self.c_image, self.c_code, self.c_text):
            wl.addWidget(cb)
        g.addWidget(wg, 1, 2)
        row = QHBoxLayout()
        lab = QLabel("Print &Quality:")
        row.addWidget(lab)
        self.quality = QComboBox()
        lab.setBuddy(self.quality)
        self.quality.addItems(["High", "Medium", "Low", "Draft"])
        row.addWidget(self.quality)
        self.to_file = QCheckBox("Print to Fi&le")
        row.addWidget(self.to_file)
        g.addLayout(row, 2, 0, 1, 3)
        col = QVBoxLayout()
        ok = QPushButton("OK")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        self.setup_btn = QPushButton("&Setup...")
        if setup is not None:
            self.setup_btn.clicked.connect(lambda: (setup(), self.printer_label.setText(self._name())))
        hb = QPushButton("&Help")
        if help is not None:
            hb.clicked.connect(help)
        for b in (ok, cancel):
            col.addWidget(b)
        col.addSpacing(12)
        col.addWidget(self.setup_btn)
        col.addStretch()
        col.addWidget(hb)
        g.addLayout(col, 0, 3, 3, 1)
        self._printer_name = printer_name
        for rb in (self.r_sel, self.r_mod, self.r_proj):
            rb.toggled.connect(self._enable)
        self._enable()

    def _name(self):
        return self.printer_label.text()

    def _enable(self):
        forms = self.r_proj.isChecked() or (self.r_mod.isChecked() and self.form_active)
        for cb in (self.c_image, self.c_text):
            cb.setEnabled(forms)

    def values(self):
        rng = "selection" if self.r_sel.isChecked() else "project" if self.r_proj.isChecked() else "module"
        what = set()
        if self.c_code.isChecked():
            what.add("code")
        if self.c_image.isChecked() and self.c_image.isEnabled():
            what.add("image")
        if self.c_text.isChecked() and self.c_text.isEnabled():
            what.add("text")
        return {"range": rng, "what": what, "quality": self.quality.currentText(), "to_file": self.to_file.isChecked()}


class PropertyPagesDialog(QDialog):
    """View > Property Pages (Shift+F4): the pages of the selected objects; the controls of LinBASIC Studio
    have none of their own, so it shows what VB6 shows for them (DEV-8/me2 pictures 11, 12)"""

    def __init__(self, parent=None):
        from PyQt6.QtWidgets import QTabWidget
        super().__init__(parent)
        self.setWindowTitle("Property Pages")
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        page = QWidget()
        pl = QVBoxLayout(page)
        self.message = QLabel("There are no selected objects with property pages.")
        self.message.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pl.addWidget(self.message)
        self.tabs.addTab(page, "None")
        self.tabs.setMinimumSize(320, 190)
        lay.addWidget(self.tabs, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        self.b_ok = QPushButton("OK")
        self.b_cancel = QPushButton("Cancel")
        self.b_apply = QPushButton("Apply")
        self.b_apply.setEnabled(False)          # nothing changed
        for b in (self.b_ok, self.b_cancel, self.b_apply):
            row.addWidget(b)
        lay.addLayout(row)
        self.b_ok.setDefault(True)
        self.b_ok.clicked.connect(self.accept)
        self.b_cancel.clicked.connect(self.reject)


class MakeGroupDialog(QDialog):
    """File > Make Project Group...: the projects of the group to build (DEV-8/me2 picture 27)"""

    def __init__(self, parent, names):
        from PyQt6.QtWidgets import QListWidget, QAbstractItemView
        super().__init__(parent)
        self.setWindowTitle(APP_NAME)                 # VB6: the name of the program ("Microsoft Visual Basic")
        lay = QGridLayout(self)
        lay.addWidget(QLabel("Build the following projects:"), 0, 0)
        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.list.addItems(names)
        for i in range(self.list.count()):
            self.list.item(i).setSelected(True)     # VB6: all of them at first
        lay.addWidget(self.list, 1, 0)
        col = QVBoxLayout()
        self.b_build = QPushButton("Build")
        self.b_cancel = QPushButton("Cancel")
        self.b_help = QPushButton("Help")
        for b in (self.b_build, self.b_cancel, self.b_help):
            col.addWidget(b)
        col.addStretch(1)
        lay.addLayout(col, 0, 1, 2, 1)
        self.default_opts = QCheckBox("Use Default Build Options")
        self.default_opts.setChecked(True)
        lay.addWidget(self.default_opts, 2, 0, 1, 2)
        self.b_build.setDefault(True)
        self.b_build.clicked.connect(self.accept)
        self.b_cancel.clicked.connect(self.reject)
        self.b_help.clicked.connect(lambda: parent.help_contents() if hasattr(parent, "help_contents") else None)
        self.list.itemSelectionChanged.connect(
            lambda: self.b_build.setEnabled(bool(self.list.selectedItems())))
        self.resize(400, 290)

    def chosen(self):
        """the rows chosen, in the order of the group"""
        return [i for i in range(self.list.count()) if self.list.item(i).isSelected()]


class AddMDIFormDialog(QDialog):
    """Project > Add MDI Form: New (the MDI Form) or Existing (a .frm), and "Don't show this dialog in the
    future" (DEV-8/me2 picture 1)"""

    def __init__(self, parent, start_dir):
        from PyQt6.QtWidgets import QTabWidget, QListWidget, QListWidgetItem, QListView
        from PyQt6.QtGui import QIcon
        super().__init__(parent)
        self.setWindowTitle("Add MDI Form")
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        new = QWidget()
        nl = QHBoxLayout(new)
        self.items = QListWidget()
        self.items.setViewMode(QListView.ViewMode.IconMode)
        self.items.setIconSize(self.items.iconSize().expandedTo(self.items.iconSize()) * 2)
        it = QListWidgetItem(QIcon.fromTheme("window-new"), "MDI Form")
        self.items.addItem(it)
        self.items.setCurrentItem(it)
        self.items.itemDoubleClicked.connect(lambda *_: self._open())
        nl.addWidget(self.items, 1)
        self.tabs.addTab(new, "New")
        self.files = QFileDialog(self, "Existing", start_dir, "Form Files (*.frm);;All Files (*)")
        self.files.setOption(QFileDialog.Option.DontUseNativeDialog, True)
        self.files.setWindowFlags(Qt.WindowType.Widget)
        self.files.setFileMode(QFileDialog.FileMode.ExistingFile)
        self.files.accepted.connect(self._open)
        self.files.rejected.connect(self.reject)
        self.tabs.addTab(self.files, "Existing")
        row = QHBoxLayout()
        row.addWidget(self.tabs, 1)
        col = QVBoxLayout()
        self.b_open = QPushButton("Open")
        self.b_cancel = QPushButton("Cancel")
        self.b_help = QPushButton("Help")
        for b in (self.b_open, self.b_cancel, self.b_help):
            col.addWidget(b)
        col.addStretch(1)
        row.addLayout(col)
        lay.addLayout(row, 1)
        self.dont_show = QCheckBox("Don't show this dialog in the future")
        lay.addWidget(self.dont_show)
        self.b_open.setDefault(True)
        self.b_open.clicked.connect(self._open)
        self.b_cancel.clicked.connect(self.reject)
        self.b_help.clicked.connect(lambda: parent.help_contents() if hasattr(parent, "help_contents") else None)
        self.tabs.currentChanged.connect(self._tab_changed)
        self.path = None                        # Existing: the .frm chosen
        self.resize(520, 380)

    def _tab_changed(self, i):
        """Existing: the file browser has its own Open / Cancel (as VB6's Existing tab)"""
        for b in (self.b_open, self.b_cancel, self.b_help):
            b.setVisible(i == 0)

    def _open(self):
        if self.tabs.currentIndex() == 1:
            sel = self.files.selectedFiles()
            if not sel or not os.path.isfile(sel[0]):
                return
            self.path = sel[0]
        self.accept()
