"""The API Viewer (phase 9.4): Add-Ins > API Viewer when the "VB 6 API Viewer" add-in is loaded, and on its own
(python3 -m lbstudio.apiviewer, without Insert), as VB6's (oracle DEV-9 api-w, ODHGIES 9.4).

The window of VB6's: File (Load Text File..., Load Database File..., Convert Text to Database..., Exit), Edit (Copy),
View (Line Item / Full Text, Load Last File), Help (About API Viewer...); API Type (Constants, Declares, Types), the
first letters of the word looked for, Available Items, Add, Declare Scope (Public / Private: all the selected items),
Selected Items, Remove, Clear, Insert (into the selected module, at the end of its declarations), Copy. Our database
is LBAPI.TXT (tools/mkapi.py, from the headers of MinGW-w64); the texts of the messages, the icon and the About box
are our own.
"""
import os
import sys

from PyQt6.QtCore import Qt, QRectF, QPointF, QEvent
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPen, QIcon, QAction, QActionGroup, QKeySequence
from PyQt6.QtWidgets import (QMainWindow, QWidget, QGridLayout, QVBoxLayout, QLabel, QComboBox, QLineEdit,
                             QListWidget, QPushButton, QGroupBox, QRadioButton, QPlainTextEdit, QFileDialog,
                             QMessageBox, QApplication, QAbstractItemView)

from . import apimodel as A

TITLE = "API Viewer"
ADDIN = "VB 6 API Viewer"
LOAD_TEXT_TITLE = "Select a Text API File"
LOAD_DB_TITLE = "Select a Database API File"
CONVERT_TITLE = "Select a Name for the New Database"
COPIED = "All entries in Selected Items list copied to clipboard..."
SEARCH_HINT = "Search by typing directly into the list box..."
LINUX_OK = "Linux: the runtime of LinBASIC Studio carries out this procedure itself."
ABOUT = ("API Viewer - shows the Constants, Declares and Types of the Windows API (32 bits) and copies them into "
         "the code of a project of LinBASIC Studio or to the clipboard.\n\n"
         "Its database, LBAPI.TXT, is made by tools/mkapi.py from the headers and the import libraries of "
         "MinGW-w64 (public domain).\n\nLinBASIC Studio")


def insert_msg(module_name):
    return "This will Insert the selected items into: %s. Continue?" % module_name


NO_MODULE = "There is no module to insert the selected items into: select one in the Project window."
LOAD_FAILED = "The file %s could not be read:\n%s"


def apiviewer_icon(size=16):
    """the icon of Add-Ins > API Viewer: our own drawing (a page of declarations and a magnifying glass)"""
    pm = QPixmap(32, 32)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    p.setPen(QPen(QColor("#404040"), 1.2))
    p.setBrush(QColor("#FFFFFF"))
    p.drawRect(QRectF(3, 2, 18, 24))
    p.setPen(QPen(QColor("#2F6FD0"), 1.6))
    for y in (7, 11, 15, 19):
        p.drawLine(QPointF(6, y), QPointF(18 if y != 15 else 14, y))
    p.setPen(QPen(QColor("#202020"), 2.6))
    p.drawLine(QPointF(22, 22), QPointF(29, 29))
    p.setPen(QPen(QColor("#202020"), 1.6))
    p.setBrush(QColor(170, 220, 255, 170))
    p.drawEllipse(QRectF(12, 11, 12, 12))
    p.end()
    return QIcon(pm.scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio,
                           Qt.TransformationMode.SmoothTransformation))


class ApiViewer(QMainWindow):
    """the window; main: the main window of the IDE (Insert), or None on its own"""

    def __init__(self, main=None, settings=None, quiet=False):
        super().__init__(None)
        self.main = main
        self.S = settings if settings is not None else getattr(main, "S", None)
        self.quiet = quiet
        self.messages = []
        self.data = A.ApiData()
        self.sel = A.Selection(self.data)
        self.loaded = False
        self.linux = A.linux_names()
        self.setWindowTitle(TITLE)
        self.setWindowIcon(apiviewer_icon(32))
        self._menus()
        cw = QWidget()
        self.setCentralWidget(cw)
        g = QGridLayout(cw)
        g.setContentsMargins(6, 6, 6, 6)
        g.setHorizontalSpacing(6)
        g.setVerticalSpacing(2)
        lt = QLabel("A&PI Type:")
        self.kind = QComboBox()
        self.kind.addItems(A.KINDS)
        self.kind.setCurrentText("Declares")
        self.kind.setFixedWidth(200)
        lt.setBuddy(self.kind)
        g.addWidget(lt, 0, 0)
        g.addWidget(self.kind, 1, 0, Qt.AlignmentFlag.AlignLeft)
        lf = QLabel("T&ype the first few letters of the word you are looking for:")
        self.find = QLineEdit()
        lf.setBuddy(self.find)
        g.addWidget(lf, 2, 0)
        g.addWidget(self.find, 3, 0)
        la = QLabel("Available &Items:")
        self.avail = QListWidget()
        self.avail.setFixedHeight(106)
        self.avail.setUniformItemSizes(True)
        la.setBuddy(self.avail)
        g.addWidget(la, 4, 0)
        g.addWidget(self.avail, 5, 0)
        side = QVBoxLayout()
        side.setSpacing(4)
        self.b_add = QPushButton("&Add")
        side.addWidget(self.b_add)
        self.scope_box = QGroupBox("Declare Scope")
        sb = QVBoxLayout(self.scope_box)
        sb.setContentsMargins(6, 2, 6, 2)
        self.rb_public = QRadioButton("Pu&blic")
        self.rb_private = QRadioButton("Pri&vate")
        self.rb_public.setChecked(True)
        sb.addWidget(self.rb_public)
        sb.addWidget(self.rb_private)
        side.addWidget(self.scope_box)
        side.addStretch(1)
        g.addLayout(side, 5, 1)
        ls = QLabel("&Selected Items:")
        self.selected = QPlainTextEdit()
        self.selected.setReadOnly(True)
        self.selected.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        ls.setBuddy(self.selected)
        g.addWidget(ls, 6, 0)
        g.addWidget(self.selected, 7, 0)
        side2 = QVBoxLayout()
        side2.setSpacing(4)
        self.b_remove = QPushButton("&Remove")
        self.b_clear = QPushButton("C&lear")
        self.b_insert = QPushButton("I&nsert")
        self.b_copy = QPushButton("&Copy")
        for b in (self.b_remove, self.b_clear, self.b_insert, self.b_copy):
            b.setFixedWidth(126)
            side2.addWidget(b)
        self.b_add.setFixedWidth(126)
        self.scope_box.setFixedWidth(126)
        side2.addStretch(1)
        g.addLayout(side2, 7, 1)
        g.setColumnStretch(0, 1)
        g.setRowStretch(7, 1)
        self.b_insert.setVisible(main is not None)   # (the API Viewer on its own has no Insert: VB6's APILOAD)
        self.statusBar().showMessage("")
        self.resize(465, 430)
        self.kind.currentTextChanged.connect(self.fill)
        self.find.textChanged.connect(self.search)
        self.avail.currentRowChanged.connect(self.avail_changed)
        self.avail.itemDoubleClicked.connect(lambda _it: self.add())
        self.b_add.clicked.connect(self.add)
        self.rb_public.toggled.connect(lambda _on: self.render())
        self.b_remove.clicked.connect(self.remove)
        self.b_clear.clicked.connect(self.clear)
        self.b_insert.clicked.connect(self.insert)
        self.b_copy.clicked.connect(self.copy)
        self.selected.cursorPositionChanged.connect(self.update_buttons)
        self.picked = False                     # an item of Selected Items chosen by a click (or the keys): Remove
        self.selected.viewport().installEventFilter(self)
        self.selected.installEventFilter(self)
        self.update_buttons()
        if self.S is not None and self.S.get("apiviewer_load_last", False):
            last = self.S.get("apiviewer_last", "")
            if last and os.path.isfile(last):
                self.load(last)

    # ------------------------------------------------------------ the menus
    def _menus(self):
        mb = self.menuBar()
        m = mb.addMenu("&File")
        self.a_load_text = m.addAction("Load &Text File...", self.load_text_file)
        self.a_load_db = m.addAction("Load &Database File...", self.load_db_file)
        self.a_convert = m.addAction("&Convert Text to Database...", self.convert)
        m.addSeparator()
        m.addAction("E&xit", self.close)
        m = mb.addMenu("&Edit")
        self.a_copy = m.addAction("&Copy", self.copy_edit)
        self.a_copy.setShortcut(QKeySequence("Ctrl+C"))
        self.a_copy.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        m = mb.addMenu("&View")
        grp = QActionGroup(self)
        self.a_line = QAction("&Line Item", self, checkable=True)
        self.a_full = QAction("&Full Text", self, checkable=True)
        self.a_full.setChecked(True)
        for a in (self.a_line, self.a_full):
            grp.addAction(a)
            m.addAction(a)
            a.triggered.connect(lambda _c=False: self.render())
        m.addSeparator()
        self.a_load_last = QAction("Load Last &File", self, checkable=True)
        self.a_load_last.setChecked(bool(self.S.get("apiviewer_load_last", False)) if self.S is not None else False)
        self.a_load_last.triggered.connect(self.toggle_load_last)
        m.addAction(self.a_load_last)
        m = mb.addMenu("&Help")
        m.addAction("&About API Viewer...", self.about)

    def message(self, text, kind="info", buttons=None):
        self.messages.append(text)
        if self.quiet:
            return QMessageBox.StandardButton.Yes
        box = QMessageBox(self)
        box.setWindowTitle(TITLE)
        box.setText(text)
        box.setIcon({"info": QMessageBox.Icon.Information, "warn": QMessageBox.Icon.Warning,
                     "ask": QMessageBox.Icon.Question}[kind])
        if buttons:
            box.setStandardButtons(buttons)
        return QMessageBox.StandardButton(box.exec())

    def about(self):
        self.message(ABOUT)

    def toggle_load_last(self, on):
        if self.S is not None:
            self.S["apiviewer_load_last"] = bool(on)
            self.S.save()

    # ------------------------------------------------------------ files
    def _dialog_dir(self):
        last = self.S.get("apiviewer_last", "") if self.S is not None else ""
        if last and os.path.isdir(os.path.dirname(last)):
            return os.path.dirname(last)
        return A.api_dir()

    def load_text_file(self, path=None):
        if not path:
            path, _ = QFileDialog.getOpenFileName(self, LOAD_TEXT_TITLE, self._dialog_dir(), "Text (*.TXT *.txt)")
        if path:
            self.load(path)

    def load_db_file(self, path=None):
        if not path:
            path, _ = QFileDialog.getOpenFileName(self, LOAD_DB_TITLE, self._dialog_dir(),
                                                  "API Database (*.db *.DB)")
        if path:
            self.load(path, db=True)

    def load(self, path, db=None):
        if db is None:
            db = not path.lower().endswith(".txt")
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            if db:
                self.data.load_db(path)
            else:
                self.data.load_text(path)
        except (OSError, ValueError, UnicodeError) as e:
            QApplication.restoreOverrideCursor()
            self.message(LOAD_FAILED % (path, e), "warn")
            return False
        QApplication.restoreOverrideCursor()
        self.loaded = True
        self.sel.clear()
        self.setWindowTitle("%s - %s" % (TITLE, os.path.abspath(path)))
        if self.S is not None:
            self.S["apiviewer_last"] = os.path.abspath(path)
            self.S.save()
        self.fill()
        self.render()
        self.statusBar().showMessage(SEARCH_HINT if db else "")
        return True

    def convert(self, path=None):
        """File > Convert Text to Database...: the loaded text file as a database next to it"""
        if not self.loaded or not self.data.is_text:
            return False
        if not path:
            start = os.path.splitext(self.data.path)[0] + ".db"
            path, _ = QFileDialog.getSaveFileName(self, CONVERT_TITLE, start, "API Database (*.db)")
        if not path:
            return False
        try:
            self.data.save_db(path)
        except (OSError, Exception) as e:     # sqlite3.Error and the file system
            self.message(LOAD_FAILED % (path, e), "warn")
            return False
        return self.load(path, db=True)

    # ------------------------------------------------------------ the lists
    def current_kind(self):
        return self.kind.currentText()

    def fill(self, *_):
        self.avail.clear()
        if self.loaded:
            k = self.current_kind()
            self.avail.addItems([self.data.label(k, n) for n in self.data.names(k)])
            self._names = self.data.names(k)
        else:
            self._names = []
        self.search(self.find.text())
        self.update_buttons()

    def search(self, text):
        """the first item that starts with the letters (as VB6: the list goes to it, at its top)"""
        if not self._names:
            return
        t = text.strip().lower()
        if not t:
            return                              # (nothing chosen yet: Add stays off, as VB6's)
        row = -1
        for i, n in enumerate(self._names):
            if n.lower().startswith(t):
                row = i
                break
        if row >= 0:
            self.avail.setCurrentRow(row)
            self.avail.scrollToItem(self.avail.item(row), QAbstractItemView.ScrollHint.PositionAtTop)

    def avail_changed(self, row):
        self.update_buttons()
        if row < 0 or not self.loaded or self.current_kind() != "Declares":
            return
        name = self._names[row]
        text = self.data.items["Declares"][name.lower()][1]
        entry = A.declare_entry(text).lower()
        if entry in self.linux or name.lower() in self.linux:
            self.statusBar().showMessage(LINUX_OK)
        elif self.statusBar().currentMessage() == LINUX_OK:
            self.statusBar().showMessage("")

    def selected_name(self):
        r = self.avail.currentRow()
        return self._names[r] if 0 <= r < len(self._names) else None

    def scope(self):
        return "Private" if self.rb_private.isChecked() else "Public"

    # ------------------------------------------------------------ the buttons
    def add(self):
        n = self.selected_name()
        if n is None:
            return False
        self.sel.add(self.current_kind(), n)
        self.render()
        return True

    def render(self):
        self.picked = False
        full = self.a_full.isChecked()
        self.selected.setPlainText(self.sel.render(self.scope(), full))
        self.update_buttons()

    def cursor_line(self):
        return self.selected.textCursor().blockNumber()

    def remove(self):
        if self.sel.remove_at_line(self.cursor_line()):
            self.render()
            return True
        return False

    def clear(self):
        self.sel.clear()
        self.render()

    def selected_code(self):
        return "\r\n\r\n".join(self.sel.code(self.scope()))

    def copy(self):
        QApplication.clipboard().setText(self.selected_code() if self.sel.items else "")
        self.statusBar().showMessage(COPIED)

    def copy_edit(self):
        """Edit > Copy (Ctrl+C): the text chosen in Selected Items or in the box of the letters, else all"""
        w = QApplication.focusWidget()
        if w is self.selected and self.selected.textCursor().hasSelection():
            self.selected.copy()
        elif w is self.find and self.find.hasSelectedText():
            self.find.copy()
        elif self.sel.items:
            self.copy()

    def target_module(self):
        if self.main is None:
            return None
        f = getattr(self.main, "selected_module", None)
        return f() if f else None

    def insert(self):
        """Insert: the selected items at the end of the declarations of the selected module (its code window
        opens), after "This will Insert the selected items into: <module>. Continue?" """
        if not self.sel.items or self.main is None:
            return False
        m = self.target_module()
        if m is None:
            self.message(NO_MODULE, "warn")
            return False
        yes = QMessageBox.StandardButton.Yes
        if self.message(insert_msg(m.name), "ask", yes | QMessageBox.StandardButton.No) != yes:
            return False
        main = self.main
        if hasattr(main, "_sync_editors"):
            main._sync_editors()
        new, line = A.insert_into(m.code_text(), self.sel.code(self.scope()))
        cw = main.open_code(m) if hasattr(main, "open_code") else None
        if cw is not None and hasattr(main, "_set_code_text"):
            main._set_code_text(cw, new)
            m.set_code(cw.text())
            if hasattr(cw, "goto_line"):
                cw.goto_line(line + 1)
        else:
            m.set_code(new)
            m.dirty = True
        return True

    def update_buttons(self):
        on = self.loaded
        has_sel = bool(self.sel.items)
        self.b_add.setEnabled(on and self.avail.currentRow() >= 0)
        self.scope_box.setEnabled(on)
        self.b_remove.setEnabled(has_sel and self.picked and self.sel.item_at_line(self.cursor_line()) >= 0)
        self.b_clear.setEnabled(has_sel)
        self.b_insert.setEnabled(has_sel)
        self.b_copy.setEnabled(has_sel)
        self.a_convert.setEnabled(on and self.data.is_text)

    def eventFilter(self, obj, e):
        if e.type() in (QEvent.Type.MouseButtonRelease, QEvent.Type.KeyRelease) and \
                obj in (self.selected, self.selected.viewport()):
            self.picked = True
            self.update_buttons()
        return super().eventFilter(obj, e)

    def focusInEvent(self, e):
        super().focusInEvent(e)
        self.update_buttons()


def open_viewer(main):
    """Add-Ins > API Viewer: the window (one), in front"""
    v = getattr(main, "_apiviewer", None)
    if v is None or not v.isVisible():
        v = ApiViewer(main)
        main._apiviewer = v
    v.show()
    v.raise_()
    v.activateWindow()
    return v


def main_standalone(argv=None):
    """the API Viewer on its own (VB6: APILOAD.EXE), without Insert"""
    from .settings import Settings
    app = QApplication.instance() or QApplication(sys.argv if argv is None else argv)
    v = ApiViewer(None, Settings())
    v.show()
    args = (sys.argv if argv is None else argv)[1:]
    if args:
        v.load(args[0])
    return app.exec()


if __name__ == "__main__":
    sys.exit(main_standalone())

__all__ = ["ApiViewer", "open_viewer", "apiviewer_icon", "insert_msg", "TITLE", "ADDIN"]
