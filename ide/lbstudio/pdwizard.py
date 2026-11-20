"""Add-Ins > Package and Deployment Wizard (phase 9.5): the wizard of VB6 (VB-TESTS/DEV-9/pd1) for the
portable packages of LinBASIC Studio (lbvc --package, compiler/lbvc/package.py).

VB6 (pd1): an unsaved project -> "...may not work properly if you do not save the project..." (Yes / No /
Cancel, pd1-01) before the wizard opens; the wizard: "Active project:", Package / Deploy / Manage Scripts, Close,
Help (pd1-02); Package: a source newer than the program -> "Do you want to recompile?" (pd1-05), then the pages
Package Type, Package Folder (made when missing, pd1-09), Included Files (checkbox, Name / Source, Add...),
Cab Options, Installation Title, Start Menu Items, Install Locations, Shared Files, Finished! (Script name),
the Packaging Report (Save Report, Close).

LinBASIC Studio (the user's decision, 2026-11-19: portable packages only): Package only (no Deploy / Manage
Scripts); Package Type chooses the system (Linux, Windows 64 / 32-bit), so the question to recompile comes when
it is left (the program depends on the system); Cab Options becomes Package Format (zip / tar.gz / AppImage /
folder); no Installation Title, Start Menu Items, Install Locations, Shared Files (a setup program's pages)."""
import os

from PyQt6.QtCore import Qt, QProcess, QProcessEnvironment, QEventLoop, QDir, QSize, QPointF
from PyQt6.QtGui import QFileSystemModel, QIcon, QPixmap, QPainter, QColor, QPen, QPolygonF
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
                             QListWidget, QPlainTextEdit, QStackedWidget, QTreeWidget, QTreeWidgetItem, QTreeView,
                             QRadioButton, QGroupBox, QMessageBox, QFileDialog, QFrame, QApplication,
                             QAbstractItemView, QInputDialog, QStyle, QToolButton)

from . import paths

TITLE = "Package and Deployment Wizard"
SAVE_FIRST = ("The Package and Deployment Wizard may not work properly if you do not save the project. "
              "Do you want to save the project before proceeding?")
TYPES = {"linux": ("Portable Package for Linux",
                   "Use to create a folder with the program and everything it needs next to it (Qt too), to copy to "
                   "another Linux computer and run there; as it is, in a .tar.gz archive or as an AppImage."),
         "win64": ("Portable Package for Windows (64-bit)",
                   "Use to create a folder with the program and the DLLs it needs, to copy to another Windows "
                   "computer and run there without a setup program; as it is or in a .zip archive."),
         "win32": ("Portable Package for Windows (32-bit)",
                   "Use to create a folder with the program and the DLLs it needs, for 32-bit Windows too; as it "
                   "is or in a .zip archive.")}
FORMAT_TEXT = {"tar.gz": "Archive (.tar.gz)", "zip": "Archive (.zip)",
               "appimage": "AppImage (one file; App.Path is the folder of the .AppImage)", "folder": "Folder only"}


def package_icon(size=48):
    """a closed box (our own drawing): the Package button"""
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    s = size / 48.0
    pen = QPen(QColor(90, 60, 20), max(1.0, 1.5 * s) if size >= 32 else 1.0)
    p.setPen(pen)

    def poly(color, pts):
        p.setBrush(QColor(*color))
        p.drawPolygon(QPolygonF([QPointF(x * s, y * s) for x, y in pts]))
    poly((222, 184, 120), [(6, 16), (24, 8), (42, 16), (24, 24)])          # the top
    poly((196, 152, 88), [(6, 16), (24, 24), (24, 44), (6, 36)])           # the left side
    poly((176, 132, 70), [(24, 24), (42, 16), (42, 36), (24, 44)])         # the right side
    p.setPen(QPen(QColor(240, 225, 190), 3 * s))                           # the tape
    p.drawLine(QPointF(15 * s, 12 * s), QPointF(33 * s, 20 * s))
    p.end()
    return QIcon(pm)


def recompile_text(exe, sources):
    """VB6 (pd1-05)"""
    return ("The following source files for this project are newer than the executable file '%s':\n\n%s\n\n"
            "Do you want to recompile? Choose No to use the existing executable." % (exe, "\n".join(sources)))


def folder_text(folder):
    """VB6 (pd1-09)"""
    return f"Cannot find '{folder}' to use as the package folder. Do you want to create it?"


class Wizard(QDialog):
    """the main window of the wizard (pd1-02)"""

    def __init__(self, main):
        super().__init__(main)
        self.main = main
        self.project = main.project
        self.setWindowTitle(TITLE)
        # VB6 (pd1-02): "Active project:" and its path on top, the big buttons on the left with their names under
        # them, what each does on their right, Close and Help at the bottom
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Active project:"))
        self.active = QLineEdit(self.project.path or "")
        self.active.setReadOnly(True)
        lay.addWidget(self.active)
        row = QHBoxLayout()
        self.b_package = QToolButton()
        self.b_package.setText("&Package")
        self.b_package.setIcon(package_icon())
        self.b_package.setIconSize(QSize(48, 48))
        self.b_package.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.b_package.setAutoRaise(True)
        self.b_package.setFixedWidth(80)
        self.b_package.clicked.connect(self.package)
        row.addWidget(self.b_package, 0, Qt.AlignmentFlag.AlignTop)
        t = QLabel("Bundle this project into a portable package: a folder, or one file (a .zip, a .tar.gz or an "
                   "AppImage), with everything the program needs next to it.")
        t.setWordWrap(True)
        row.addWidget(t, 1, Qt.AlignmentFlag.AlignVCenter)
        lay.addSpacing(8)
        lay.addLayout(row)
        lay.addStretch(1)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        lay.addWidget(line)
        row = QHBoxLayout()
        row.addStretch(1)
        self.b_close = QPushButton("Close")
        self.b_close.clicked.connect(self.reject)
        b_help = QPushButton("Help")
        b_help.clicked.connect(lambda: main.help_keyword("Package and Deployment Wizard"))
        row.addWidget(self.b_close)
        row.addWidget(b_help)
        lay.addLayout(row)
        self.resize(460, 260)
        self.pages = None

    def package(self):
        self.pages = PackageWizard(self.main, self.project)
        self.pages.exec()


class PackageWizard(QDialog):
    """Package: the pages of the wizard and the build (lbvc --package)"""
    PAGES = ("Package Type", "Package Folder", "Included Files", "Package Format", "Finished!")

    def __init__(self, main, project):
        super().__init__(main)
        from lbvc import package as P
        self.P = P
        self.main = main
        self.project = project
        self.vbp = project.path
        self.pdir = os.path.dirname(self.vbp)
        self.settings = P.Settings.load(self.vbp)
        self.script = self.settings.recent or P.DEFAULT_SCRIPT
        old = self.settings.get(self.script)
        self.added = list(old.get("added", []))
        self.removed = list(old.get("removed", []))
        self.old = old
        self.files = []
        self.csrc = None
        self.report = None
        self.archive_ok = False
        outer = QVBoxLayout(self)
        body = QHBoxLayout()
        self.picture = QLabel()
        self.picture.setFixedSize(150, 120)
        self.picture.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.picture.setStyleSheet("background: #808080;")
        body.addWidget(self.picture, 0, Qt.AlignmentFlag.AlignTop)
        self.stack = QStackedWidget()
        body.addWidget(self.stack, 1)
        outer.addLayout(body, 1)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        outer.addWidget(line)
        row = QHBoxLayout()
        self.b_help = QPushButton("Help")
        self.b_help.clicked.connect(lambda: main.help_keyword("Package and Deployment Wizard"))
        row.addWidget(self.b_help)
        row.addStretch(1)
        self.b_cancel = QPushButton("Cancel")
        self.b_back = QPushButton("< &Back")
        self.b_next = QPushButton("&Next >")
        self.b_finish = QPushButton("&Finish")
        for b in (self.b_cancel, self.b_back, self.b_next, self.b_finish):
            row.addWidget(b)
        self.b_cancel.clicked.connect(self.reject)
        self.b_back.clicked.connect(lambda: self.go(self.stack.currentIndex() - 1))
        self.b_next.clicked.connect(self.next)
        self.b_finish.clicked.connect(self.finish)
        outer.addLayout(row)
        self._type_page()
        self._folder_page()
        self._files_page()
        self._format_page()
        self._finished_page()
        self.resize(680, 440)
        self.go(0)

    # ------------------------------------------------------------ the pages
    def _page(self, text):
        w = QFrame()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(6, 0, 0, 0)
        t = QLabel(text)
        t.setWordWrap(True)
        lay.addWidget(t)
        self.stack.addWidget(w)
        return lay

    def _type_page(self):
        from .dialogs import MakeDialog
        lay = self._page("Choose the type of package you want to create.")
        lay.addStretch(1)
        lab = QLabel("Package &type:")          # the & of a label with a buddy is a mnemonic, not shown
        lay.addWidget(lab)
        self.types = QListWidget()
        lab.setBuddy(self.types)
        self.targets = [k for k, _ in MakeDialog.targets()]
        for k in self.targets:
            self.types.addItem(TYPES[k][0])
            ok, why = MakeDialog.available(k)
            if not ok:                          # the AppImage: the Windows packages shown, disabled
                it = self.types.item(self.types.count() - 1)
                it.setFlags(it.flags() & ~(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable))
                it.setToolTip(why)
        lay.addWidget(self.types)
        lab = QLabel("&Description:")
        lay.addWidget(lab)
        self.type_desc = QPlainTextEdit()
        lab.setBuddy(self.type_desc)
        self.type_desc.setReadOnly(True)
        self.type_desc.setMaximumHeight(60)
        lay.addWidget(self.type_desc)
        self.types.currentRowChanged.connect(lambda r: self.type_desc.setPlainText(TYPES[self.targets[r]][1])
                                             if 0 <= r < len(self.targets) else None)
        want = self.old.get("Target") or paths.host_target()
        if want not in self.targets or not MakeDialog.available(want)[0]:
            want = next((k for k in self.targets if MakeDialog.available(k)[0]), self.targets[0])
        self.types.setCurrentRow(self.targets.index(want))

    def _folder_page(self):
        lay = self._page("Choose the folder where your package will be assembled.")
        lay.addStretch(1)
        lab = QLabel("&Package folder:")
        lay.addWidget(lab)
        f = self.old.get("PackageFolder") or "Package"
        self.folder = QLineEdit(os.path.normpath(f if os.path.isabs(f) else os.path.join(self.pdir, f)))
        lab.setBuddy(self.folder)
        lay.addWidget(self.folder)
        row = QHBoxLayout()
        self.fs = QFileSystemModel(self)
        self.fs.setFilter(QDir.Filter.AllDirs | QDir.Filter.NoDotAndDotDot)
        self.fs.setRootPath("/")
        self.tree = QTreeView()
        self.tree.setModel(self.fs)
        for c in (1, 2, 3):
            self.tree.hideColumn(c)
        self.tree.setHeaderHidden(True)
        self.tree.clicked.connect(lambda ix: self.folder.setText(self.fs.filePath(ix)))
        # VB6 (pd1-08): the tree opened at the folder (here: the nearest one that is there), as the model loads
        self.fs.directoryLoaded.connect(lambda _p: self.show_folder())
        row.addWidget(self.tree, 1)
        col = QVBoxLayout()
        col.addStretch(1)
        b = QPushButton("New &Folder...")
        b.clicked.connect(self.new_folder)
        col.addWidget(b)
        row.addLayout(col)
        lay.addLayout(row, 1)

    def _files_page(self):
        lay = self._page("The files in the list below will be included in your package. Click Add to include "
                         "additional files. Clear the checkbox to the left of the file name to remove a file from "
                         "the package.")
        lab = QLabel("&Files:")
        lay.addWidget(lab)
        row = QHBoxLayout()
        self.list = QTreeWidget()
        lab.setBuddy(self.list)
        self.list.setHeaderLabels(["Name", "Source"])
        self.list.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self.list.setRootIsDecorated(False)
        self.list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        row.addWidget(self.list, 1)
        col = QVBoxLayout()
        b = QPushButton("&Add...")
        b.clicked.connect(self.add_file)
        col.addWidget(b)
        col.addStretch(1)
        row.addLayout(col)
        lay.addLayout(row, 1)

    def _format_page(self):
        lay = self._page("You can make the package one archive file, to give to others as it is, or keep it as a "
                         "folder. Choose the appropriate option below.")
        lay.addStretch(1)
        self.fmt_box = QGroupBox("Package format")
        self.fmt_lay = QVBoxLayout(self.fmt_box)
        self.fmt_buttons = {}
        for k in ("tar.gz", "zip", "appimage", "folder"):
            b = QRadioButton(FORMAT_TEXT[k])
            self.fmt_buttons[k] = b
            self.fmt_lay.addWidget(b)
        lay.addWidget(self.fmt_box)
        lay.addStretch(1)

    def _finished_page(self):
        lay = self._page("The wizard has finished collecting information needed to build this package. Enter the "
                         "name under which to save the settings for this session, then click Finish to create the "
                         "package.")
        lay.addStretch(1)
        lab = QLabel("&Script name:")
        lay.addWidget(lab)
        self.script_name = QLineEdit(self.script)
        lab.setBuddy(self.script_name)
        lay.addWidget(self.script_name)

    # ------------------------------------------------------------ moving between them
    def target(self):
        r = self.types.currentRow()
        return self.targets[r] if 0 <= r < len(self.targets) else paths.host_target()

    def go(self, i):
        i = max(0, min(i, len(self.PAGES) - 1))
        self.stack.setCurrentIndex(i)
        self.setWindowTitle(f"{TITLE} - {self.PAGES[i]}")
        icons = (QStyle.StandardPixmap.SP_FileDialogDetailedView, QStyle.StandardPixmap.SP_DirIcon,
                 QStyle.StandardPixmap.SP_FileIcon, QStyle.StandardPixmap.SP_DriveHDIcon,
                 QStyle.StandardPixmap.SP_DialogApplyButton)
        self.picture.setPixmap(self.style().standardIcon(icons[i]).pixmap(64, 64))
        self.b_back.setEnabled(i > 0)
        self.b_next.setEnabled(i < len(self.PAGES) - 1)
        self.b_finish.setEnabled(i == len(self.PAGES) - 1)
        (self.b_finish if i == len(self.PAGES) - 1 else self.b_next).setDefault(True)

    def next(self):
        i = self.stack.currentIndex()
        if i == 0 and not self.leave_type():
            return
        if i == 1 and not self.leave_folder():
            return
        if i == 2:
            self.leave_files()
        self.go(i + 1)
        if i + 1 == 1:
            self.show_folder()
        if i + 1 == 2:
            self.fill_files()
        if i + 1 == 3:
            self.fill_formats()

    def ask(self, text, buttons):
        """a question of the wizard (the tests answer it)"""
        B = QMessageBox.StandardButton
        r = QMessageBox.question(self, TITLE, text, buttons)
        return {B.Yes: "yes", B.No: "no", B.Cancel: "cancel"}.get(r, "cancel")

    def leave_type(self):
        """the program for the chosen system: compiled again when a source is newer (VB6 pd1-05)"""
        P = self.P
        target = self.target()
        self.exe = P.exe_path(self.vbp, target)
        stale = P.stale_sources(self.vbp, self.exe)
        if not os.path.exists(self.exe) or P.program_target(self.exe) not in (None, target):
            return self.compile(target)
        if stale:
            B = QMessageBox.StandardButton
            a = self.ask(recompile_text(self.exe, stale), B.Yes | B.No | B.Cancel)
            if a == "cancel":
                return False
            if a == "yes":
                return self.compile(target)
        return True

    def leave_folder(self):
        f = os.path.normpath(os.path.expanduser(self.folder.text().strip()))
        if not f:
            return False
        if not os.path.isabs(f):
            f = os.path.join(self.pdir, f)
        if not os.path.isdir(f):
            B = QMessageBox.StandardButton
            if self.ask(folder_text(f), B.Yes | B.No) != "yes":
                return False
            try:
                os.makedirs(f)
            except OSError as e:
                QMessageBox.warning(self, TITLE, str(e))
                return False
        self.folder.setText(f)
        return True

    def show_folder(self):
        """the folder tree opened at the package folder, or at the nearest folder above it that exists"""
        p = os.path.abspath(os.path.expanduser(self.folder.text().strip() or self.pdir))
        while p and not os.path.isdir(p) and os.path.dirname(p) != p:
            p = os.path.dirname(p)
        ix = self.fs.index(p)
        if not ix.isValid():
            return
        up = ix.parent()
        while up.isValid():
            self.tree.expand(up)
            up = up.parent()
        if self.tree.currentIndex() != ix:
            self.tree.setCurrentIndex(ix)
        self.tree.scrollTo(ix, QAbstractItemView.ScrollHint.PositionAtCenter)

    def new_folder(self):
        base = self.folder.text().strip() or self.pdir
        name, ok = QInputDialog.getText(self, "New Folder", "Folder name:")
        if ok and name.strip():
            p = os.path.join(base if os.path.isdir(base) else os.path.dirname(base), name.strip())
            os.makedirs(p, exist_ok=True)
            self.folder.setText(p)

    def fill_files(self):
        """VB6 Included Files (pd1-10, 11): the program first, its runtime, the licenses, the files added"""
        P = self.P
        try:
            if self.csrc is None:
                from lbvc.driver import compile_project
                self.csrc, _ = compile_project([self.vbp], self.main.codepage())
            fmt = self.old.get("Format")
            self.files = P.package_files(self.vbp, self.target(), self.csrc, self.exe, self.added, self.removed,
                                         fmt)
        except (SystemExit, Exception) as e:    # noqa: BLE001 - shown to the user
            QMessageBox.warning(self, TITLE, str(e))
            self.files = []
        self.list.clear()
        for f in self.files:
            it = QTreeWidgetItem([f.name, f.source or "(made by the wizard)"])
            flags = it.flags() | Qt.ItemFlag.ItemIsUserCheckable
            if f.kind == "program":             # the program itself is always in the package: checked, fixed
                flags &= ~Qt.ItemFlag.ItemIsUserCheckable
            it.setFlags(flags)
            for c in (0, 1):
                it.setToolTip(c, it.text(c))
            it.setCheckState(0, Qt.CheckState.Checked if f.checked else Qt.CheckState.Unchecked)
            it.setData(0, Qt.ItemDataRole.UserRole, f)
            self.list.addTopLevelItem(it)
        # VB6 (pd1-10, 11): the names narrow, the sources beside them (cut in the middle when long)
        self.list.setColumnWidth(0, min(self.list.sizeHintForColumn(0) + 24, 240))

    def leave_files(self):
        self.removed = []
        for i in range(self.list.topLevelItemCount()):
            it = self.list.topLevelItem(i)
            f = it.data(0, Qt.ItemDataRole.UserRole)
            f.checked = it.checkState(0) == Qt.CheckState.Checked or f.kind == "program"
            if not f.checked:
                self.removed.append(f.name)
        self.added = [self.P.rel_to(f.source, self.pdir) for f in self.files if f.kind == "added" and f.checked]

    def add_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Add File", self.pdir, "All Files (*)")
        if not path:
            return
        self.leave_files()
        rel = self.P.rel_to(path, self.pdir)
        if rel not in self.added:
            self.added.append(rel)
        self.fill_files()

    def fill_formats(self):
        """Linux: tar.gz, AppImage (with appimagetool), folder; Windows: zip, folder"""
        linux = self.target() == "linux"
        tool, _ = self.P.appimage_tools()
        for k, b in self.fmt_buttons.items():
            b.setVisible(k in (("tar.gz", "appimage", "folder") if linux else ("zip", "folder")))
            b.setEnabled(k != "appimage" or bool(tool))
        b = self.fmt_buttons["appimage"]
        b.setToolTip("" if tool else "appimagetool is not found (LinBASIC Studio carries it in its AppImage)")
        want = self.old.get("Format") or self.P.default_format(self.target())
        if self.fmt_buttons[want].isHidden() or not self.fmt_buttons[want].isEnabled():
            want = self.P.default_format(self.target())
        self.fmt_buttons[want].setChecked(True)

    def format(self):
        for k, b in self.fmt_buttons.items():
            if b.isChecked() and not b.isHidden():
                return k
        return self.P.default_format(self.target())

    # ------------------------------------------------------------ the work of lbvc
    def lbvc(self, args, what):
        """lbvc in its own process, the wizard waiting (the IDE stays drawn) -> (exit code, its output)"""
        from .runner import lbvc_command
        cmd = lbvc_command()
        p = QProcess(self)
        env = QProcessEnvironment.systemEnvironment()
        env.insert("LBVC_RUNTIME", paths.RUNTIME)
        p.setProcessEnvironment(env)
        loop = QEventLoop()
        p.finished.connect(lambda *a: loop.quit())
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        self.main.statusBar().showMessage(what)
        try:
            p.start(cmd[0], cmd[1:] + args)
            if p.waitForStarted(10000):
                loop.exec()
        finally:
            QApplication.restoreOverrideCursor()
            self.main.statusBar().clearMessage()
        out = bytes(p.readAllStandardOutput()).decode("utf-8", "replace")
        err = bytes(p.readAllStandardError()).decode("utf-8", "replace")
        return p.exitCode(), out, err

    def compile(self, target):
        args = [self.vbp, "-o", self.exe, "--codepage", self.main.codepage(), "--target", target]
        code, out, err = self.lbvc(args, f"Compiling {os.path.basename(self.exe)}...")
        if code != 0 or not os.path.exists(self.exe):
            QMessageBox.warning(self, TITLE, (err + out).strip() or "The compiler failed.")
            return False
        self.csrc = None
        return True

    def finish(self):
        """VB6 Finish (pd1-20): the settings under the script name, the package, the report (pd1-21)"""
        P = self.P
        name = self.script_name.text().strip() or P.DEFAULT_SCRIPT
        folder = self.folder.text().strip()
        self.settings.put(name, {"Target": self.target(), "Format": self.format(),
                                 "PackageFolder": P.rel_to(folder, self.pdir),
                                 "added": list(self.added), "removed": list(self.removed)})
        self.settings.save()
        args = [self.vbp, "--package", "--script", name, "--recompile", "no", "--target", self.target(),
                "--codepage", self.main.codepage()]
        code, out, err = self.lbvc(args, "Building the package...")
        if code != 0:
            QMessageBox.warning(self, TITLE, (err + out).strip() or "The package could not be built.")
            return
        self.report = out
        self.accept()
        rep = ReportDialog(self.main, os.path.basename(self.vbp), out)
        self.main._pdw_report = rep             # (the tests)
        rep.exec()


class ReportDialog(QDialog):
    """VB6 "<project>.vbp - Packaging Report" (pd1-21): the report, Save Report, Close"""

    def __init__(self, parent, vbp_name, text):
        super().__init__(parent)
        self.setWindowTitle(f"{vbp_name} - Packaging Report")
        lay = QVBoxLayout(self)
        self.text = QPlainTextEdit(text)
        self.text.setReadOnly(True)
        lay.addWidget(self.text, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        b = QPushButton("&Save Report")
        b.clicked.connect(self.save)
        c = QPushButton("Close")
        c.clicked.connect(self.accept)
        row.addWidget(b)
        row.addWidget(c)
        lay.addLayout(row)
        self.resize(520, 380)

    def save(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Report", "", "Text Files (*.txt)")
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.text.toPlainText())


def open_wizard(main):
    """Add-Ins > Package and Deployment Wizard...: an unsaved project first (VB6 pd1-01)"""
    p = main.project
    if p is None:
        return None
    if p.untitled or p.any_dirty():
        B = QMessageBox.StandardButton
        r = main.pdw_ask(SAVE_FIRST, B.Yes | B.No | B.Cancel)
        if r == "cancel":
            return None
        if r == "yes" and not main.save_project():
            return None
        if p.untitled:
            QMessageBox.information(main, TITLE, "The project must be saved before it can be packaged.")
            return None
    w = Wizard(main)
    main._pdw = w                               # (the tests)
    w.exec()
    return w
