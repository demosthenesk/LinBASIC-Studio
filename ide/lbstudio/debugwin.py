"""The windows at the bottom of the IDE: Output (the messages of the compiler and the errors of the
program), and those of the debugger: Locals, Watches, Call Stack, Add Watch and the Immediate window
that runs ? expression and statements while the program is in break mode."""

from PyQt6.QtCore import Qt, QPoint, QRectF, pyqtSignal
from PyQt6.QtGui import QColor, QTextCharFormat, QTextCursor, QIcon, QPixmap, QPainter, QPen, QPolygon
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTreeWidget, QTreeWidgetItem, QToolButton,
                             QDialog, QListWidget, QPushButton, QLineEdit, QDialogButtonBox, QFormLayout,
                             QPlainTextEdit, QHeaderView, QAbstractItemView, QGroupBox, QComboBox, QRadioButton,
                             QGridLayout, QApplication)

# the kinds of a watch (the Watch Type of VB6's Add Watch)
WATCH_EXPR, WATCH_TRUE, WATCH_CHANGE = 0, 1, 2
ALL_PROCS = "(All Procedures)"
ALL_MODS = "(All Modules)"


class OutputWindow(QPlainTextEdit):
    """Output: what the compiler and the program report, one message a line; a double click on a
    line with a place in the code (Form1.frm(12, 5): ...) opens the code at it"""
    goto = pyqtSignal(str, int, int)            # file of the project, line of the file (1-based), column

    COLORS = {"error": QColor(200, 30, 30), "warning": QColor(190, 110, 0), "ok": QColor(30, 130, 50),
              "head": None, "info": None}

    def __init__(self):
        super().__init__()
        self.setReadOnly(True)
        self.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.setUndoRedoEnabled(False)
        self.links = {}                          # block number -> (file, line, column)

    def clear(self):
        super().clear()
        self.links = {}

    def add(self, text, kind="info", where=None):
        """kind: info, head, ok, error; where: (file, line, column) for a double click"""
        cur = QTextCursor(self.document())
        cur.movePosition(QTextCursor.MoveOperation.End)
        fmt = QTextCharFormat()
        color = self.COLORS.get(kind)
        if color is not None:
            fmt.setForeground(color)
        if kind == "head":
            fmt.setFontWeight(700)
        for i, line in enumerate(text.splitlines() or [""]):
            if not self.document().isEmpty():
                cur.insertBlock()
            cur.insertText(line, fmt)
            if where and i == 0:
                self.links[cur.blockNumber()] = where
        self.setTextCursor(cur)
        self.ensureCursorVisible()

    def mouseDoubleClickEvent(self, e):
        n = self.cursorForPosition(e.position().toPoint()).blockNumber()
        if n in self.links:
            self.goto.emit(*self.links[n])
            return
        super().mouseDoubleClickEvent(e)


class ImmediateEdit(QPlainTextEdit):
    """the Immediate window: Debug.Print output, and the line with the caret runs on Enter"""
    command = pyqtSignal(str)

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and not e.modifiers() & Qt.KeyboardModifier.ShiftModifier:
            cur = self.textCursor()
            text = cur.block().text()
            cur.movePosition(cur.MoveOperation.EndOfBlock)
            self.setTextCursor(cur)
            if text.strip():
                self.command.emit(text)
            else:
                super().keyPressEvent(e)
            return
        super().keyPressEvent(e)

    def answer(self, text):
        """the result under the line that was run (VB6 inserts it on the next line)"""
        cur = self.textCursor()
        cur.movePosition(cur.MoveOperation.EndOfBlock)
        cur.insertText("\n" + text if text is not None else "\n")
        if text is not None:
            cur.insertText("\n")
        self.setTextCursor(cur)
        self.ensureCursorVisible()


def _tree(cols):
    t = QTreeWidget()
    t.setColumnCount(len(cols))
    t.setHeaderLabels(cols)
    t.setRootIsDecorated(True)
    t.setUniformRowHeights(True)
    t.header().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
    t.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    return t


class LocalsWindow(QWidget):
    """Locals: the variables of the procedure (and those of its module under Me)"""
    stack_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(1)
        top = QHBoxLayout()
        self.where = QLabel("<Ready>")
        top.addWidget(self.where, 1)
        b = QToolButton()
        b.setText("...")
        b.setToolTip("Call Stack")
        b.clicked.connect(self.stack_requested.emit)
        top.addWidget(b)
        lay.addLayout(top)
        self.tree = _tree(["Expression", "Value", "Type"])
        lay.addWidget(self.tree)

    def clear(self, ready=True):
        self.tree.clear()
        if ready:
            self.where.setText("<Ready>")

    def copy(self):
        """Copy of the menu: the line (Expression, Value, Type)"""
        it = self.tree.currentItem()
        if it is not None:
            QApplication.clipboard().setText("\t".join(it.text(i) for i in range(self.tree.columnCount())))

    def can_collapse_parent(self):
        it = self.tree.currentItem()
        return it is not None and it.parent() is not None

    def collapse_parent(self):
        """Collapse Parent (VB6, DEV-8/me8-19): the line that holds the current one is closed and chosen"""
        it = self.tree.currentItem()
        if it is not None and it.parent() is not None:
            par = it.parent()
            par.setExpanded(False)
            self.tree.setCurrentItem(par)

    def fill(self, where, rows, module_label):
        self.where.setText(where)
        self.tree.clear()
        mnode = None
        for r in rows:
            if len(r) < 4:
                continue
            kind, name, typ, val = r[0], r[1], r[2], r[3]
            if kind == "mvar":
                if mnode is None:
                    mnode = QTreeWidgetItem([module_label, "", ""])
                    self.tree.insertTopLevelItem(0, mnode)
                mnode.addChild(QTreeWidgetItem([name, val, typ]))
            else:
                self.tree.addTopLevelItem(QTreeWidgetItem([name, val, typ]))
        for i in range(3):
            self.tree.resizeColumnToContents(i)


def watch_icon(kind):
    """the picture of a watch in the Watches window: glasses (Watch Expression), a raised hand (Break When Value
    Is True), a delta (Break When Value Changes); drawn here, in the spirit of VB6's"""
    pm = QPixmap(16, 16)
    pm.fill(QColor(0, 0, 0, 0))
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    if kind == WATCH_EXPR:
        p.setPen(QPen(QColor(40, 40, 40), 1.4))
        p.setBrush(QColor(170, 210, 240))
        p.drawEllipse(QRectF(1.5, 5.5, 5.5, 5.5))
        p.drawEllipse(QRectF(9, 5.5, 5.5, 5.5))
        p.drawLine(7, 8, 9, 8)
    elif kind == WATCH_TRUE:
        p.setPen(QPen(QColor(120, 60, 0), 1))
        p.setBrush(QColor(250, 200, 120))
        for i, x in enumerate((4, 6.5, 9, 11.5)):                   # the fingers
            p.drawRoundedRect(QRectF(x - 1, 2 + (i in (0, 3)) * 1.5, 2.2, 7), 1, 1)
        p.drawRoundedRect(QRectF(3, 7, 10, 7), 2.5, 2.5)             # the palm
    else:
        p.setPen(QPen(QColor(0, 70, 140), 1.3))
        p.setBrush(QColor(120, 180, 240))
        p.drawPolygon(QPolygon([QPoint(8, 2), QPoint(14, 14), QPoint(2, 14)]))
    p.end()
    return QIcon(pm)


def watch_context(w):
    """the Context column: Module1.Main, Module1, or (All)"""
    if w.get("mod") and w.get("proc"):
        return f"{w['mod']}.{w['proc']}"
    return w.get("mod") or "(All)"


class WatchWindow(QWidget):
    """Watches: expressions evaluated at every break"""
    edit_requested = pyqtSignal(int)
    delete_requested = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.tree = _tree(["Expression", "Value", "Type", "Context"])
        self.tree.setRootIsDecorated(False)
        self.hit = -1                           # the watch of the last break (select)
        self.tree.itemDoubleClicked.connect(lambda it, c: self.edit_requested.emit(self.tree.indexOfTopLevelItem(it)))
        lay.addWidget(self.tree)

    def keyPressEvent(self, e):
        it = self.tree.currentItem()
        if it is not None and e.key() == Qt.Key.Key_Delete:
            self.delete_requested.emit(self.tree.indexOfTopLevelItem(it))
            return
        super().keyPressEvent(e)

    def fill(self, watches):
        """watches: [(expression, value, type, context[, kind])]"""
        cur = self.hit if self.hit >= 0 else self.current()
        self.tree.clear()
        for w in watches:
            it = QTreeWidgetItem([str(x) for x in w[:4]])
            it.setIcon(0, watch_icon(w[4] if len(w) > 4 else WATCH_EXPR))
            self.tree.addTopLevelItem(it)
        for i in range(4):
            self.tree.resizeColumnToContents(i)
        if 0 <= cur < self.tree.topLevelItemCount():
            self.tree.setCurrentItem(self.tree.topLevelItem(cur))

    def select(self, i):
        """the watch that made the program stop (-1: none); it stays selected when the values come"""
        self.hit = i
        if 0 <= i < self.tree.topLevelItemCount():
            self.tree.setCurrentItem(self.tree.topLevelItem(i))

    def current(self):
        it = self.tree.currentItem()
        return self.tree.indexOfTopLevelItem(it) if it is not None else -1

    def copy(self):
        """Copy of the menu: the line (Expression, Value, Type, Context)"""
        it = self.tree.currentItem()
        if it is not None:
            QApplication.clipboard().setText("\t".join(it.text(i) for i in range(self.tree.columnCount())))


class CallStackDialog(QDialog):
    def __init__(self, parent, frames):
        """frames: [(module, procedure, line)], the newest first"""
        super().__init__(parent)
        self.setWindowTitle("Call Stack")
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Project.Module.Function"))
        self.list = QListWidget()
        for mod, proc, line in frames:
            self.list.addItem(f"{mod}.{proc}")
        self.list.setCurrentRow(0)
        self.list.itemDoubleClicked.connect(lambda _: self.accept())
        lay.addWidget(self.list)
        row = QHBoxLayout()
        show = QPushButton("&Show")
        show.clicked.connect(self.accept)
        close = QPushButton("&Close")
        close.clicked.connect(self.reject)
        row.addStretch()
        row.addWidget(show)
        row.addWidget(close)
        lay.addLayout(row)
        self.resize(420, 260)

    def chosen(self):
        return self.list.currentRow()


class AddWatchDialog(QDialog):
    """Add Watch / Edit Watch, as VB6's: the Expression, its Context (a procedure of a module, a module, or all
    of them) and the Watch Type (Watch Expression, Break When Value Is True, Break When Value Changes)

    watch: {"expr", "kind", "mod", "proc"}; modules: [(module name, [procedure names])]; project: its name"""

    def __init__(self, parent, watch=None, modules=(), project="", title="Add Watch"):
        super().__init__(parent)
        w = dict(watch or {})
        self.setWindowTitle(title)
        self.modules = list(modules)
        lay = QVBoxLayout(self)
        f = QFormLayout()
        self.expr = QLineEdit(w.get("expr", ""))
        self.expr.setMinimumWidth(360)
        f.addRow("&Expression:", self.expr)
        lay.addLayout(f)

        ctx = QGroupBox("Context")
        g = QGridLayout(ctx)
        self.proc_box = QComboBox()
        self.mod_box = QComboBox()
        self.mod_box.addItem(ALL_MODS)
        for name, _ in self.modules:
            self.mod_box.addItem(name)
        g.addWidget(QLabel("&Procedure:"), 0, 0)
        g.addWidget(self.proc_box, 0, 1)
        g.addWidget(QLabel("&Module:"), 1, 0)
        g.addWidget(self.mod_box, 1, 1)
        g.addWidget(QLabel("Project:"), 2, 0)
        g.addWidget(QLabel(project), 2, 1)
        g.setColumnStretch(1, 1)
        lay.addWidget(ctx)
        self.mod_box.currentIndexChanged.connect(lambda _: self._fill_procs())
        mi = self.mod_box.findText(w.get("mod") or ALL_MODS, Qt.MatchFlag.MatchFixedString)
        self.mod_box.setCurrentIndex(max(0, mi))
        self._fill_procs()
        pi = self.proc_box.findText(w.get("proc") or ALL_PROCS, Qt.MatchFlag.MatchFixedString)
        self.proc_box.setCurrentIndex(max(0, pi))

        wt = QGroupBox("Watch Type")
        v = QVBoxLayout(wt)
        self.kinds = [QRadioButton("&Watch Expression"), QRadioButton("Break When Value Is &True"),
                      QRadioButton("Break When Value &Changes")]
        for rb in self.kinds:
            v.addWidget(rb)
        self.kinds[w.get("kind", WATCH_EXPR) if w.get("kind", WATCH_EXPR) in (0, 1, 2) else 0].setChecked(True)
        lay.addWidget(wt)

        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _fill_procs(self):
        self.proc_box.clear()
        self.proc_box.addItem(ALL_PROCS)
        mod = self.mod_box.currentText()
        for name, procs in self.modules:
            if name == mod:
                self.proc_box.addItems(procs)
        self.proc_box.setEnabled(mod != ALL_MODS)

    def expression(self):
        return self.expr.text().strip()

    def set_kind(self, kind):
        self.kinds[kind].setChecked(True)

    def set_context(self, mod="", proc=""):
        self.mod_box.setCurrentIndex(max(0, self.mod_box.findText(mod or ALL_MODS, Qt.MatchFlag.MatchFixedString)))
        self.proc_box.setCurrentIndex(max(0, self.proc_box.findText(proc or ALL_PROCS, Qt.MatchFlag.MatchFixedString)))

    def watch(self):
        mod = self.mod_box.currentText()
        mod = "" if mod == ALL_MODS else mod
        proc = self.proc_box.currentText() if mod else ""
        return {"expr": self.expression(), "kind": next(i for i, rb in enumerate(self.kinds) if rb.isChecked()),
                "mod": mod, "proc": "" if proc == ALL_PROCS else proc}
