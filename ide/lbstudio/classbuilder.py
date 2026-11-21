"""The Class Builder Utility (phase 9.2c): Add-Ins > Class Builder Utility... when the "VB 6 Class Builder Utility"
add-in is loaded, and Project > Add Class Module > VB Class Builder (VB6 oracle DEV-9/cb1 ... cb6).

The window of VB6's: the classes and collections of the project in a tree (Project1 at the root), the members of
the selected one in the Properties / Methods / Events / All lists; File (New > Class..., Collection...; Update
Project; Exit), Edit (Cut, Copy, Paste; Delete, Rename; Properties...), View (ToolBar; Options...), Help; the
toolbar (Add New Class, Add New Collection | Property, Method, Event, Enum to Current Class | Properties | Cut,
Copy, Paste); the builders of a class, a collection, a property, a method (Add Argument), an event and an enum.
The model and the code are in cbmodel.py; the texts and pictures are our own.
"""
import os

from PyQt6.QtCore import Qt, QRectF, QPointF
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPen, QIcon, QPolygonF, QKeySequence, QAction
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit, QComboBox,
                             QPushButton, QTabWidget, QWidget, QTreeWidget, QTreeWidgetItem, QListWidget,
                             QSplitter, QMenuBar, QToolBar, QMenu, QMessageBox, QGroupBox, QRadioButton,
                             QCheckBox, QPlainTextEdit, QInputDialog, QAbstractItemView,
                             QToolButton)

from . import cbmodel as C

TITLE = "Class Builder"
ADDIN = "VB 6 Class Builder Utility"
TABS = [("Properties", ["Name", "Data Type", "Arguments", "Prop Declaration"], ("property",)),
        ("Methods", ["Name", "Data Type", "Arguments"], ("method",)),
        ("Events", ["Name", "Arguments"], ("event",)),
        ("All", ["Name", "Data Type", "Arguments", "Prop Declaration"], ("property", "method", "event", "enum"))]


# ================================================================ the pictures (our own)
def _pm(size=16):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.scale(size / 16.0, size / 16.0)
    return pm, p


def _box(p, x, y, w, h, fill, pen=QColor(40, 40, 40)):
    p.setPen(QPen(pen, 1))
    p.setBrush(fill)
    p.drawRect(QRectF(x + .5, y + .5, w, h))


def picture(kind, size=16):
    """class, collection, property, method, event, enum, project, ref, the toolbar's add-*, props, cut, copy,
    paste, builder"""
    pm, p = _pm(size)
    teal, gold, purple = QColor(0, 128, 128), QColor(230, 190, 40), QColor(130, 60, 170)
    green, white, red = QColor(40, 150, 60), QColor(255, 255, 255), QColor(200, 40, 40)
    base = kind[4:] if kind.startswith("add-") else kind
    if base == "class":
        _box(p, 2, 3, 11, 10, white)
        p.setBrush(purple)
        p.setPen(Qt.PenStyle.NoPen)
        p.drawRect(QRectF(3, 4, 10, 3))
        _box(p, 5, 8, 3, 3, teal)
    elif base == "collection":
        for i, x in enumerate((1, 6, 11)):
            _box(p, x, 9, 4, 4, purple if i != 1 else teal)
        p.setPen(QPen(QColor(40, 40, 40), 1))
        p.drawLine(QPointF(8, 4), QPointF(8, 8))
        p.drawLine(QPointF(3, 8), QPointF(13, 8))
        _box(p, 5, 1, 6, 4, red)
    elif base in ("property", "ref"):
        _box(p, 1, 4, 12, 9, white)
        p.setPen(QPen(QColor(60, 60, 60), 1))
        for y in (7, 10):
            p.drawLine(QPointF(3, y), QPointF(11, y))
        _box(p, 9, 1, 6, 5, gold if base == "property" else teal)
    elif base == "method":
        p.setPen(QPen(QColor(20, 90, 30), 1))
        p.setBrush(green)
        p.drawPolygon(QPolygonF([QPointF(2, 9), QPointF(8, 4), QPointF(14, 9), QPointF(8, 13)]))
        p.setPen(QPen(QColor(60, 60, 60), 1))
        p.drawLine(QPointF(1, 13), QPointF(5, 13))
    elif base == "event":
        p.setPen(QPen(QColor(120, 90, 0), 1))
        p.setBrush(gold)
        p.drawPolygon(QPolygonF([QPointF(9, 1), QPointF(4, 9), QPointF(8, 9), QPointF(6, 15), QPointF(12, 6),
                                 QPointF(8, 6), QPointF(10, 1)]))
    elif base == "enum":
        _box(p, 1, 5, 9, 8, white)
        _box(p, 6, 2, 9, 8, white)
        p.setPen(QPen(QColor(60, 60, 60), 1))
        for y in (5, 7):
            p.drawLine(QPointF(8, y), QPointF(13, y))
    elif base == "project":
        _box(p, 1, 5, 9, 9, white)
        _box(p, 6, 1, 9, 8, QColor(240, 220, 150))
        _box(p, 3, 8, 3, 3, purple)
    elif base == "props":
        _box(p, 2, 2, 11, 12, white)
        p.setPen(QPen(QColor(60, 60, 60), 1))
        for y in (5, 8, 11):
            p.drawLine(QPointF(4, y), QPointF(11, y))
    elif base == "cut":
        p.setPen(QPen(QColor(40, 40, 40), 1.4))
        p.drawLine(QPointF(5, 1), QPointF(10, 10))
        p.drawLine(QPointF(11, 1), QPointF(6, 10))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawEllipse(QRectF(3, 10, 4, 4))
        p.drawEllipse(QRectF(9, 10, 4, 4))
    elif base == "copy":
        _box(p, 1, 1, 8, 10, white)
        _box(p, 6, 5, 8, 10, white)
    elif base == "paste":
        _box(p, 2, 3, 10, 12, QColor(190, 150, 90))
        _box(p, 5, 1, 4, 3, QColor(160, 160, 160))
        _box(p, 7, 7, 8, 8, white)
    elif base == "builder":
        p.setPen(QPen(QColor(40, 40, 40), 1))
        _box(p, 1, 9, 6, 6, purple)
        _box(p, 9, 9, 6, 6, teal)
        _box(p, 5, 1, 6, 6, white)
        p.drawLine(QPointF(8, 7), QPointF(4, 9))
        p.drawLine(QPointF(8, 7), QPointF(12, 9))
    if kind.startswith("add-"):                    # a small plus: a new one
        p.setPen(QPen(QColor(0, 110, 0), 2))
        p.drawLine(QPointF(12, 11), QPointF(12, 16))
        p.drawLine(QPointF(9.5, 13.5), QPointF(14.5, 13.5))
    p.end()
    return pm


def icon(kind, size=16):
    return QIcon(picture(kind, size))


def builder_icon(size=16):
    """the icon of Add-Ins > Class Builder Utility..."""
    return icon("builder", size)


# ================================================================ the builders
class _Builder(QDialog):
    """a builder with the tabs Properties / Attributes (Description, Project Help File, Help Context ID)"""

    def __init__(self, parent, title, help_file="", readonly=False):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.readonly = readonly
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        self.props = QWidget()
        self.form = QVBoxLayout(self.props)
        self.tabs.addTab(self.props, "Properties")
        at = QWidget()
        al = QVBoxLayout(at)
        al.addWidget(QLabel("&Description:"))
        self.desc = QPlainTextEdit()
        self.desc.setMaximumHeight(90)
        al.addWidget(self.desc)
        al.addWidget(QLabel("Project Help File:"))
        self.help_file = QLineEdit(help_file)
        self.help_file.setEnabled(False)
        al.addWidget(self.help_file)
        al.addWidget(QLabel("Help &Context ID:"))
        self.helpid = QLineEdit()
        self.helpid.setMaximumWidth(150)
        al.addWidget(self.helpid)
        al.addStretch(1)
        self.tabs.addTab(at, "Attributes")
        lay.addWidget(self.tabs)
        row = QHBoxLayout()
        row.addStretch(1)
        self.b_ok = QPushButton("OK")
        self.b_cancel = QPushButton("Cancel")
        self.b_ok.setDefault(True)
        self.b_ok.clicked.connect(self._ok)
        self.b_cancel.clicked.connect(self.reject)
        row.addWidget(self.b_ok)
        row.addWidget(self.b_cancel)
        lay.addLayout(row)
        self.messages = []
        self.quiet = getattr(parent, "quiet", False)

    def name_edit(self, text):
        self.form.addWidget(QLabel("&Name:"))
        self.name = QLineEdit(text)
        self.name.selectAll()
        self.form.addWidget(self.name)

    def message(self, text, title=None):
        self.messages.append(text)
        if not self.quiet:
            QMessageBox.information(self, title or self.windowTitle(), text)

    def attributes(self):
        try:
            hid = int(self.helpid.text().strip() or 0)
        except ValueError:
            hid = 0
        return self.desc.toPlainText().strip(), hid

    def set_attributes(self, desc, hid):
        self.desc.setPlainText(desc or "")
        self.helpid.setText(str(hid) if hid else "")

    def lock(self):
        """a member that is in the code already: only to look at (VB6 DEV-9/cb3-06)"""
        for kind in (QLineEdit, QComboBox, QCheckBox, QRadioButton, QPlainTextEdit, QPushButton, QListWidget,
                     QToolButton):
            for w in self.findChildren(kind):
                if w not in (self.b_ok, self.b_cancel):
                    w.setEnabled(False)

    def check_name(self, taken=()):
        n = self.name.text().strip()
        if not C.valid_name(n):
            self.message("Invalid Name!")
            return None
        if n.lower() in {t.lower() for t in taken}:
            self.message(f"'{n}' already exists!", TITLE)
            return None
        return n

    def _ok(self):
        if self.readonly or self.apply():
            self.accept()

    def apply(self):
        return True


class ClassBuilderDialog(_Builder):
    """Class Module Builder: Name, Based On ((New Class) or a class: its members are copied)"""

    def __init__(self, parent, model, c=None, help_file="", taken=(), readonly=False):
        super().__init__(parent, "Class Module Builder", help_file, readonly)
        self.model, self.c, self.taken = model, c, taken
        self.name_edit(c.name if c else model.unique("Class"))
        self.form.addWidget(QLabel("&Based On:"))
        self.based = QComboBox()
        self.based.addItem("(New Class)")
        for k in model.classes:
            if not k.is_collection and k is not c:
                self.based.addItem(k.name)
        self.based.setEnabled(c is None)
        self.form.addWidget(self.based)
        self.form.addStretch(1)
        if c is not None:
            self.set_attributes(c.description, c.helpid)
        self.result_class = None
        if readonly:
            self.lock()
        self.resize(360, 330)

    def apply(self):
        n = self.check_name(self.taken)
        if n is None:
            return False
        c = self.c or C.CBClass(n)
        c.name = n
        c.description, c.helpid = self.attributes()
        if self.c is None and self.based.currentIndex() > 0:
            base = self.model.cls(self.based.currentText())
            if base is not None:                 # Based On: the members of the class (DEV-9/cb2: clsDerived)
                c.members = [m.copy() for m in base.members]
        self.result_class = c
        return True


class CollectionBuilderDialog(_Builder):
    """Collection Builder: Name, Based On ((New Collection) or a collection), Collection Of: an existing class or a
    new one (New Class Properties...: its Class Module Builder)"""

    def __init__(self, parent, model, c=None, help_file="", taken=(), readonly=False):
        super().__init__(parent, "Collection Builder", help_file, readonly)
        self.model, self.c, self.taken = model, c, taken
        top = QHBoxLayout()
        left = QVBoxLayout()
        left.addWidget(QLabel("&Name:"))
        self.name = QLineEdit(c.name if c else model.unique("Collection"))
        self.name.selectAll()
        left.addWidget(self.name)
        left.addWidget(QLabel("&Based On:"))
        self.based = QComboBox()
        self.based.addItem("(New Collection)")
        for k in model.classes:
            if k.collection and k is not c:
                self.based.addItem(k.name)
        self.based.setEnabled(c is None)
        left.addWidget(self.based)
        left.addStretch(1)
        top.addLayout(left, 1)
        g = QGroupBox("Collection Of")
        gl = QVBoxLayout(g)
        self.rb_existing = QRadioButton("&Existing Class")
        self.rb_new = QRadioButton("Ne&w Class")
        self.existing = QListWidget()
        for k in model.classes:
            if not k.is_collection:
                self.existing.addItem(k.name)
        self.new_name = QLineEdit(model.unique("Class"))
        self.b_newprops = QPushButton("New Class P&roperties...")
        self.b_newprops.clicked.connect(self.new_class_properties)
        gl.addWidget(self.rb_existing)
        gl.addWidget(self.existing)
        gl.addWidget(self.rb_new)
        gl.addWidget(self.new_name)
        gl.addWidget(self.b_newprops)
        top.addWidget(g, 1)
        self.form.addLayout(top)
        self.rb_existing.toggled.connect(self._mode)
        (self.rb_existing if self.existing.count() else self.rb_new).setChecked(True)
        self._mode()
        if c is not None:
            self.set_attributes(c.description, c.helpid)
            items = self.existing.findItems(c.item, Qt.MatchFlag.MatchExactly)
            if items:
                self.existing.setCurrentItem(items[0])
            for w in (self.rb_existing, self.rb_new, self.existing, self.new_name, self.b_newprops):
                w.setEnabled(False)
        self.new_class = None                    # the class of New Class (with its Class Module Builder's settings)
        self.result_class = None
        if readonly:
            self.lock()
        self.resize(560, 360)

    def _mode(self):
        ex = self.rb_existing.isChecked()
        self.existing.setEnabled(ex)
        self.new_name.setEnabled(not ex)
        self.b_newprops.setEnabled(not ex)

    def new_class_properties(self):
        """New Class Properties...: the Class Module Builder of the new class (DEV-9/cb2-03)"""
        c = self.new_class or C.CBClass(self.new_name.text().strip() or self.model.unique("Class"))
        c.name = self.new_name.text().strip() or c.name
        d = ClassBuilderDialog(self, self.model, c, self.help_file.text(), self._taken_new())
        d.quiet = self.quiet
        if self.quiet:
            self._last_class_dialog = d
            return d
        if d.exec():
            self.new_class = d.result_class
            self.new_name.setText(d.result_class.name)
        return d

    def _taken_new(self):
        return list(self.taken) + [self.name.text().strip()]

    def apply(self):
        n = self.name.text().strip()
        if not C.valid_name(n):
            self.message("Invalid Name!")
            return False
        if n.lower() in {t.lower() for t in self.taken}:
            self.message(f"'{n}' already exists!", TITLE)
            return False
        c = self.c or C.CBClass(n, True)
        c.name = n
        c.description, c.helpid = self.attributes()
        if self.c is None:
            if self.based.currentIndex() > 0:
                base = self.model.cls(self.based.currentText())
                if base is not None:
                    c.item = base.item
            if self.rb_existing.isChecked():
                it = self.existing.currentItem()
                if it is None:
                    self.message("Select the class of the items of the collection in the list of the classes.",
                                 "Collection Builder")
                    return False
                c.item = it.text()
                self.new_class = None
            else:
                cn = self.new_name.text().strip()
                if not C.valid_name(cn):
                    self.message("Invalid Name!")
                    return False
                if cn.lower() in {t.lower() for t in self._taken_new()}:
                    self.message(f"'{cn}' already exists!", TITLE)
                    return False
                nc = self.new_class or C.CBClass(cn)
                nc.name = cn
                nc.top_level = False
                if nc.member("Key") is None:        # the new class of a collection: Key (DEV-9/cb2-04)
                    nc.members.insert(0, C.Member("property", "Key", "String", decl="var"))
                self.new_class = nc
                c.item = cn
        self.result_class = c
        return True


def _type_combo(model, kind):
    cb = QComboBox()
    if kind == "return":
        cb.addItems(C.RETURN_TYPES)
        cb.addItems(sorted(model.names(), key=str.lower))
    else:
        cb.addItems(sorted(C.BASE_TYPES + model.names(), key=str.lower))
    cb.setEditable(False)
    return cb


class PropertyBuilderDialog(_Builder):
    """Property Builder: Name, Data Type, Declaration (Public Property (Let, Get, Set), Friend Property (Let, Get,
    Set), Public Variable), Default Property?"""

    def __init__(self, parent, model, cls, m=None, help_file="", readonly=False):
        super().__init__(parent, "Property Builder", help_file, readonly)
        self.model, self.cls, self.m = model, cls, m
        self.name_edit(m.name if m else "")
        self.form.addWidget(QLabel("&Data Type:"))
        self.type = _type_combo(model, "property")
        self.type.setCurrentText(m.type if m and m.type else "Variant")
        self.form.addWidget(self.type)
        g = QGroupBox("Declaration")
        gl = QVBoxLayout(g)
        self.rb_public = QRadioButton("&Public Property (Let, Get, Set)")
        self.rb_friend = QRadioButton("&Friend Property (Let, Get, Set)")
        self.rb_var = QRadioButton("Public &Variable")
        self.cb_default = QCheckBox("De&fault Property?")
        for w in (self.rb_public, self.rb_friend, self.rb_var):
            gl.addWidget(w)
        gl.addSpacing(8)
        gl.addWidget(self.cb_default)
        self.form.addWidget(g)
        self.form.addStretch(1)
        decl = m.decl if m else "public"
        {"public": self.rb_public, "friend": self.rb_friend, "var": self.rb_var}[decl].setChecked(True)
        if m is not None:
            self.cb_default.setChecked(m.default)
            self.set_attributes(m.description, m.helpid)
        self.result_member = None
        if readonly:
            self.lock()
        self.resize(360, 380)

    def apply(self):
        taken = [x.name for x in self.model.display_members(self.cls) if x is not self.m]
        n = self.check_name(taken)
        if n is None:
            return False
        m = self.m or C.Member("property", n)
        m.name, m.type = n, self.type.currentText()
        m.decl = "var" if self.rb_var.isChecked() else "friend" if self.rb_friend.isChecked() else "public"
        m.default = self.cb_default.isChecked()
        m.description, m.helpid = self.attributes()
        self.result_member = m
        return True


class ArgumentDialog(QDialog):
    """Add Argument: Name, ByVal, Data Type, Array, Optional and its Default Value (DEV-9/cb1-26)"""

    def __init__(self, parent, model, arg=None):
        super().__init__(parent)
        self.setWindowTitle("Add Argument")
        self.quiet = getattr(parent, "quiet", False)
        self.messages = []
        lay = QGridLayout(self)
        lay.addWidget(QLabel("&Name:"), 0, 0)
        self.name = QLineEdit(arg.name if arg else "")
        lay.addWidget(self.name, 1, 0)
        self.byval = QCheckBox("&ByVal")
        lay.addWidget(self.byval, 2, 0)
        lay.addWidget(QLabel("&Data Type:"), 3, 0)
        self.type = _type_combo(model, "arg")
        self.type.setCurrentText(arg.type if arg else "Variant")
        lay.addWidget(self.type, 4, 0)
        self.array = QCheckBox("&Array")
        lay.addWidget(self.array, 5, 0)
        g = QGroupBox()
        gl = QVBoxLayout(g)
        self.optional = QCheckBox("&Optional")
        gl.addWidget(self.optional)
        gl.addWidget(QLabel("Default &Value: (quote strings as needed)"))
        self.default = QLineEdit()
        gl.addWidget(self.default)
        lay.addWidget(g, 6, 0)
        col = QVBoxLayout()
        self.b_ok = QPushButton("OK")
        self.b_ok.setDefault(True)
        self.b_cancel = QPushButton("Cancel")
        self.b_ok.clicked.connect(self._ok)
        self.b_cancel.clicked.connect(self.reject)
        col.addWidget(self.b_ok)
        col.addWidget(self.b_cancel)
        col.addStretch(1)
        lay.addLayout(col, 0, 1, 7, 1)
        self.byval.toggled.connect(self._states)
        self.optional.toggled.connect(self._states)
        if arg is not None:
            self.byval.setChecked(arg.byval)
            self.array.setChecked(arg.array)
            self.optional.setChecked(arg.optional)
            self.default.setText(str(arg.default))
        self._states()
        self.result_arg = None

    def _states(self):
        """an array is not passed ByVal; the default value only for an Optional argument"""
        if self.byval.isChecked():
            self.array.setChecked(False)
        self.array.setEnabled(not self.byval.isChecked())
        self.default.setEnabled(self.optional.isChecked())

    def _ok(self):
        n = self.name.text().strip()
        if not C.valid_name(n):
            self.messages.append("Invalid Name!")
            if not self.quiet:
                QMessageBox.information(self, "Add Argument", "Invalid Name!")
            return
        self.result_arg = C.Arg(n, self.type.currentText(), self.byval.isChecked(), self.array.isChecked(),
                                self.optional.isChecked(),
                                self.default.text().strip() if self.optional.isChecked() else "")
        self.accept()


class _ArgsMixin:
    """the Arguments list and its buttons (+ add, x remove, up, down)"""

    def args_box(self, args):
        self.args = [a.copy() for a in args]
        self.form.addWidget(QLabel("&Arguments:"))
        row = QHBoxLayout()
        self.arg_list = QListWidget()
        row.addWidget(self.arg_list, 1)
        col = QVBoxLayout()
        self.arg_buttons = []
        for text, tip, slot in (("+", "Add an argument", lambda: self.add_arg()),
                                ("×", "Remove the argument", self.del_arg),
                                ("↑", "Move up", lambda: self.move_arg(-1)),
                                ("↓", "Move down", lambda: self.move_arg(1))):
            b = QToolButton()
            b.setText(text)
            b.setToolTip(tip)
            b.clicked.connect(slot)
            col.addWidget(b)
            self.arg_buttons.append(b)
        col.addStretch(1)
        row.addLayout(col)
        self.form.addLayout(row)
        self.fill_args()

    def fill_args(self):
        self.arg_list.clear()
        for a in self.args:
            self.arg_list.addItem(a.text())

    def add_arg(self, arg=None):
        if arg is None:
            d = ArgumentDialog(self, self.model)
            self._arg_dialog = d
            if self.quiet or not d.exec():
                return d
            arg = d.result_arg
        if any(a.name.lower() == arg.name.lower() for a in self.args):
            self.message(f"'{arg.name}' already exists!", TITLE)
            return None
        self.args.append(arg)
        self.fill_args()
        self.arg_list.setCurrentRow(len(self.args) - 1)
        return arg

    def del_arg(self):
        r = self.arg_list.currentRow()
        if 0 <= r < len(self.args):
            del self.args[r]
            self.fill_args()

    def move_arg(self, d):
        r = self.arg_list.currentRow()
        k = r + d
        if 0 <= r < len(self.args) and 0 <= k < len(self.args):
            self.args[r], self.args[k] = self.args[k], self.args[r]
            self.fill_args()
            self.arg_list.setCurrentRow(k)


class MethodBuilderDialog(_Builder, _ArgsMixin):
    """Method Builder: Name, Arguments, Return Data Type, Declare as Friend?, Default Method?"""

    def __init__(self, parent, model, cls, m=None, help_file="", readonly=False):
        super().__init__(parent, "Method Builder", help_file, readonly)
        self.model, self.cls, self.m = model, cls, m
        self.name_edit(m.name if m else "")
        self.args_box(m.args if m else [])
        self.form.addWidget(QLabel("&Return Data Type:"))
        self.ret = _type_combo(model, "return")
        self.ret.setCurrentText(m.type if m and m.type else "(None)")
        self.form.addWidget(self.ret)
        self.cb_friend = QCheckBox("Declare as &Friend?")
        self.cb_default = QCheckBox("De&fault Method?")
        self.form.addWidget(self.cb_friend)
        self.form.addWidget(self.cb_default)
        if m is not None:
            self.cb_friend.setChecked(m.friend)
            self.cb_default.setChecked(m.default)
            self.set_attributes(m.description, m.helpid)
        self.result_member = None
        if readonly:
            self.lock()
        self.resize(380, 420)

    def apply(self):
        taken = [x.name for x in self.model.display_members(self.cls) if x is not self.m]
        n = self.check_name(taken)
        if n is None:
            return False
        m = self.m or C.Member("method", n)
        m.name = n
        r = self.ret.currentText()
        m.type = "" if r == "(None)" else r
        m.args = self.args
        m.friend, m.default = self.cb_friend.isChecked(), self.cb_default.isChecked()
        m.description, m.helpid = self.attributes()
        self.result_member = m
        return True


class EventBuilderDialog(_Builder, _ArgsMixin):
    """Event Builder: Name, Arguments"""

    def __init__(self, parent, model, cls, m=None, help_file="", readonly=False):
        super().__init__(parent, "Event Builder", help_file, readonly)
        self.model, self.cls, self.m = model, cls, m
        self.name_edit(m.name if m else "")
        self.args_box(m.args if m else [])
        self.form.addStretch(1)
        if m is not None:
            self.set_attributes(m.description, m.helpid)
        self.result_member = None
        if readonly:
            self.lock()
        self.resize(360, 330)

    def apply(self):
        taken = [x.name for x in self.model.display_members(self.cls) if x is not self.m]
        n = self.check_name(taken)
        if n is None:
            return False
        m = self.m or C.Member("event", n)
        m.name, m.args = n, self.args
        m.description, m.helpid = self.attributes()
        self.result_member = m
        return True


class EnumBuilderDialog(QDialog):
    """Enum Builder: Name, Public, Values (+: Add Enum Value "x=1"); OK with a value at least (DEV-9/cb1-37)"""

    def __init__(self, parent, model, cls, m=None, readonly=False):
        super().__init__(parent)
        self.setWindowTitle("Enum Builder")
        self.model, self.cls, self.m = model, cls, m
        self.quiet = getattr(parent, "quiet", False)
        self.messages = []
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("&Name:"))
        self.name = QLineEdit(m.name if m else "")
        lay.addWidget(self.name)
        self.public = QCheckBox("&Public")
        self.public.setChecked(m.public if m else True)
        lay.addWidget(self.public)
        lay.addWidget(QLabel("&Values:"))
        row = QHBoxLayout()
        self.values = QListWidget()
        for v in (m.values if m else []):
            self.values.addItem(v)
        row.addWidget(self.values, 1)
        col = QVBoxLayout()
        for text, slot in (("+", lambda: self.add_value()), ("×", self.del_value), ("↑", lambda: self.move(-1)),
                           ("↓", lambda: self.move(1))):
            b = QToolButton()
            b.setText(text)
            b.clicked.connect(slot)
            col.addWidget(b)
        col.addStretch(1)
        row.addLayout(col)
        lay.addLayout(row)
        br = QHBoxLayout()
        br.addStretch(1)
        self.b_ok = QPushButton("OK")
        self.b_cancel = QPushButton("Cancel")
        self.b_ok.clicked.connect(self._ok)
        self.b_cancel.clicked.connect(self.reject)
        br.addWidget(self.b_ok)
        br.addWidget(self.b_cancel)
        lay.addLayout(br)
        self.readonly = readonly
        self._states()
        self.result_member = None
        if readonly:
            for w in (self.name, self.public, self.values):
                w.setEnabled(False)
        self.resize(260, 330)

    def _states(self):
        self.b_ok.setEnabled(self.readonly or self.values.count() > 0)

    def add_value(self, text=None):
        if text is None:
            text, ok = QInputDialog.getText(self, "Add Enum Value:", "Add New Value (eg. x=1)")
            if not ok:
                return
        text = (text or "").strip()
        name = text.split("=")[0].strip()
        if not text or not C.valid_name(name):
            self.messages.append("Invalid Name!")
            if not self.quiet:
                QMessageBox.information(self, "Enum Builder", "Invalid Name!")
            return
        self.values.addItem(text.replace(" ", ""))
        self._states()

    def del_value(self):
        r = self.values.currentRow()
        if r >= 0:
            self.values.takeItem(r)
            self._states()

    def move(self, d):
        r = self.values.currentRow()
        k = r + d
        if r >= 0 and 0 <= k < self.values.count():
            it = self.values.takeItem(r)
            self.values.insertItem(k, it)
            self.values.setCurrentRow(k)

    def _ok(self):
        if self.readonly:
            self.accept()
            return
        n = self.name.text().strip()
        taken = [x.name for x in self.model.display_members(self.cls) if x is not self.m]
        if not C.valid_name(n):
            self.messages.append("Invalid Name!")
            if not self.quiet:
                QMessageBox.information(self, "Enum Builder", "Invalid Name!")
            return
        if n.lower() in {t.lower() for t in taken}:
            self.messages.append(f"'{n}' already exists!")
            if not self.quiet:
                QMessageBox.information(self, TITLE, f"'{n}' already exists!")
            return
        m = self.m or C.Member("enum", n)
        m.name = n
        m.public = self.public.isChecked()
        m.values = [self.values.item(i).text() for i in range(self.values.count())]
        self.result_member = m
        self.accept()


class OptionsDialog(QDialog):
    """View > Options...: Code Generation (DEV-9/cb1-47)"""

    def __init__(self, parent, debug, errors):
        super().__init__(parent)
        self.setWindowTitle("Class Builder Options")
        lay = QVBoxLayout(self)
        g = QGroupBox("Code Generation")
        gl = QVBoxLayout(g)
        self.cb_debug = QCheckBox("Include &Debug code in Initialize and Terminate events")
        self.cb_err = QCheckBox("Include &Err.Raise in all generated methods")
        self.cb_debug.setChecked(debug)
        self.cb_err.setChecked(errors)
        gl.addWidget(self.cb_debug)
        gl.addWidget(self.cb_err)
        lay.addWidget(g)
        row = QHBoxLayout()
        row.addStretch(1)
        b = QPushButton("OK")
        b.setDefault(True)
        b.clicked.connect(self.accept)
        c = QPushButton("Cancel")
        c.clicked.connect(self.reject)
        row.addWidget(b)
        row.addWidget(c)
        lay.addLayout(row)


class ForeignDialog(QDialog):
    """the project has class modules the Class Builder did not make (DEV-9/cb2-11)"""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle(TITLE)
        lay = QVBoxLayout(self)
        t = QLabel("Some class modules of this project were not made by the Class Builder.\n\n"
                   "They are shown at the root of the tree: the Class Builder does not know where they belong "
                   "in your object model, so place them as your design needs.\n\n"
                   "When you update the project, the Class Builder records the object model in the class modules, "
                   "so that it finds it the next time.")
        t.setWordWrap(True)
        lay.addWidget(t)
        self.again = QCheckBox("&Do not show this dialog in the future")
        lay.addWidget(self.again)
        b = QPushButton("OK")
        b.setDefault(True)
        b.clicked.connect(self.accept)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(b)
        row.addStretch(1)
        lay.addLayout(row)
        self.resize(420, 220)


# ================================================================ the window
EXISTING_MEMBER = ("'{name}' is already in the code of the class.  The Class Builder does not change the code you "
                   "have: {what} it in the code window of {cls}.")
EXISTING_CLASS = ("'{name}' is a class module of the project already.  The Class Builder does not remove modules: "
                  "use Project > Remove {name}.")


class ClassBuilder(QDialog):
    """the window of the Class Builder"""

    def __init__(self, main, project=None, quiet=False):
        super().__init__(main)
        self.main = main
        self.project = project if project is not None else getattr(main, "project", None)
        self.quiet = quiet                           # the tests: the messages in self.messages, no message box
        self.messages = []
        self.setWindowTitle(TITLE)
        self.setWindowIcon(builder_icon(32))
        self.model = C.read_project(self.project)
        S = getattr(main, "S", None)
        self.model.debug_code = bool(S["cb_debug_code"]) if S is not None else False
        self.model.err_raise = bool(S["cb_err_raise"]) if S is not None else False
        self.help_file = (self.project.vbp.get("HelpFile") or "").strip('"') if self.project is not None else ""
        self.clip = None                             # ("member", Member, class name, cut) / ("class", name, cut)
        self.updated = False
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.menubar = QMenuBar()
        lay.setMenuBar(self.menubar)
        self._build_actions()
        self.toolbar = QToolBar()
        for a in (self.a_new_class, self.a_new_coll, None, self.a_new_prop, self.a_new_method, self.a_new_event,
                  self.a_new_enum, None, self.a_props, None, self.a_cut, self.a_copy, self.a_paste):
            if a is None:
                self.toolbar.addSeparator()
            else:
                self.toolbar.addAction(a)
        lay.addWidget(self.toolbar)
        split = QSplitter()
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(4, 0, 0, 4)
        ll.addWidget(QLabel("Classes:"))
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._tree_menu)
        self.tree.currentItemChanged.connect(lambda *_: self.class_changed())
        self.tree.itemChanged.connect(self._tree_renamed)
        ll.addWidget(self.tree)
        split.addWidget(left)
        self.tabs = QTabWidget()
        self.lists = []
        for title, cols, kinds in TABS:
            t = QTreeWidget()
            t.setColumnCount(len(cols))
            t.setHeaderLabels(cols)
            t.setRootIsDecorated(False)
            t.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
            t.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            t.customContextMenuRequested.connect(lambda pos, t=t: self._list_menu(t, pos))
            t.itemDoubleClicked.connect(lambda *_: self.properties())
            t.itemChanged.connect(self._member_renamed)
            t.currentItemChanged.connect(lambda *_: self._states())
            for i, w in enumerate((140, 90, 120, 110)[:len(cols)]):
                t.setColumnWidth(i, w)
            t.kinds = kinds
            self.lists.append(t)
            self.tabs.addTab(t, title)
        self.tabs.setCurrentIndex(3)
        self.tabs.currentChanged.connect(lambda *_: self._states())
        split.addWidget(self.tabs)
        split.setStretchFactor(1, 1)
        split.setSizes([180, 440])
        lay.addWidget(split, 1)
        self._filling = False
        self.fill_tree()
        self.resize(680, 380)

    # ------------------------------------------------------------ menus and toolbar
    def _action(self, menu, text, slot, key=None, pic=None, tip=None):
        a = QAction(text, self)
        if key:
            a.setShortcut(QKeySequence(key))
        if pic:
            a.setIcon(icon(pic))
        if tip:
            a.setToolTip(tip)
        a.triggered.connect(lambda checked=False: slot())
        if menu is not None:
            menu.addAction(a)
        return a

    def _build_actions(self):
        mb = self.menubar
        m = mb.addMenu("&File")
        nm = m.addMenu("&New")
        self.a_new_class = self._action(nm, "&Class...", self.new_class, pic="add-class", tip="Add New Class")
        self.a_new_coll = self._action(nm, "C&ollection...", self.new_collection, pic="add-collection",
                                       tip="Add New Collection")
        m.addSeparator()
        self.a_update = self._action(m, "&Update Project", self.update_project, "Ctrl+S")
        m.addSeparator()
        self._action(m, "E&xit", self.reject)
        m = mb.addMenu("&Edit")
        self.a_cut = self._action(m, "Cu&t", self.cut, "Ctrl+X", "cut", "Cut")
        self.a_copy = self._action(m, "&Copy", self.copy, "Ctrl+C", "copy", "Copy")
        self.a_paste = self._action(m, "&Paste", self.paste, "Ctrl+V", "paste", "Paste")
        m.addSeparator()
        self.a_delete = self._action(m, "&Delete", self.delete)
        self.a_rename = self._action(m, "Re&name", self.rename)
        m.addSeparator()
        self.a_props = self._action(m, "P&roperties...", self.properties, pic="props",
                                    tip="Show Properties for the current object")
        m = mb.addMenu("&View")
        self.a_toolbar = self._action(m, "&ToolBar", self.toggle_toolbar)
        self.a_toolbar.setCheckable(True)
        self.a_toolbar.setChecked(True)
        m.addSeparator()
        self._action(m, "&Options...", self.options)
        m = mb.addMenu("&Help")
        self._action(m, "&Help Topics", self.help)
        # the members of the current class: the toolbar and New in the menus of the right button
        self.a_new_prop = self._action(None, "&Property...", self.new_property, pic="add-property",
                                       tip="Add New Property to Current Class")
        self.a_new_method = self._action(None, "&Method...", self.new_method, pic="add-method",
                                         tip="Add New Method to Current Class")
        self.a_new_event = self._action(None, "&Event...", self.new_event, pic="add-event",
                                        tip="Add New Event to Current Class")
        self.a_new_enum = self._action(None, "E&num...", self.new_enum, pic="add-enum",
                                       tip="Add New Enum to Current Class")
        self.a_set_coll = self._action(None, "Set As Co&llection", self.set_as_collection)

    def toggle_toolbar(self):
        self.toolbar.setVisible(self.a_toolbar.isChecked())

    # ------------------------------------------------------------ the tree and the lists
    def fill_tree(self, select=None):
        cur = select or (self.current_class().name if self.current_class() else None)
        self._filling = True
        self.tree.clear()
        root = QTreeWidgetItem([self.model.project_name])
        root.setIcon(0, icon("project"))
        root.cls = None
        self.tree.addTopLevelItem(root)
        found = [root]

        def add(parent, node):
            c, ref, kids = node
            it = QTreeWidgetItem([c.name])
            it.cls, it.ref = c, ref
            it.setIcon(0, icon("ref" if ref else "collection" if c.is_collection else "class"))
            if not ref:
                it.setFlags(it.flags() | Qt.ItemFlag.ItemIsEditable)
            parent.addChild(it)
            if cur and c.name.lower() == cur.lower() and len(found) == 1:
                found.append(it)
            for k in kids:
                add(it, k)
        for node in self.model.tree():
            add(root, node)
        self.tree.expandAll()
        self._filling = False
        self.tree.setCurrentItem(found[-1])
        self.class_changed()

    def current_class(self):
        it = self.tree.currentItem()
        return getattr(it, "cls", None) if it is not None else None

    def current_list(self):
        return self.lists[self.tabs.currentIndex()]

    def current_member(self):
        it = self.current_list().currentItem()
        return getattr(it, "member", None) if it is not None else None

    def class_changed(self):
        c = self.current_class()
        self._filling = True
        for t in self.lists:
            t.clear()
            if c is None:
                continue
            for m in sorted(self.model.display_members(c), key=lambda x: x.name.lower()):
                if m.kind not in t.kinds:
                    continue
                vals = {"Name": m.name, "Data Type": m.data_type(), "Arguments": m.arguments(),
                        "Prop Declaration": m.prop_declaration()}
                cols = [t.headerItem().text(i) for i in range(t.columnCount())]
                it = QTreeWidgetItem([vals[k] for k in cols])
                it.setIcon(0, icon(m.kind))
                it.member = m
                if not (c.collection and not c.existing):
                    it.setFlags(it.flags() | Qt.ItemFlag.ItemIsEditable)
                t.addTopLevelItem(it)
        self._filling = False
        self._states()

    def _states(self):
        c = self.current_class()
        on = c is not None
        for a in (self.a_new_prop, self.a_new_method, self.a_new_event, self.a_new_enum):
            a.setEnabled(on)
        sel = on or self.current_member() is not None
        for a in (self.a_props, self.a_cut, self.a_copy, self.a_delete, self.a_rename):
            a.setEnabled(sel)
        self.a_paste.setEnabled(self.clip is not None)
        self.a_set_coll.setEnabled(on and not (c.collection if c else True))
        self.a_set_coll.setText("Set As &Class" if on and c.user_collection else "Set As Co&llection")

    def _focus_member(self, target=None):
        """the Edit commands act on the member of the list if the list has the focus (or the right button was
        pressed on it), else on the class of the tree; target: "member" / "class" (the tests)"""
        target = target or self._menu_target
        if target is not None:
            return target == "member"
        return self.current_list().hasFocus() and self.current_member() is not None

    def _on_ref(self):
        """the item class under a collection: the class itself is elsewhere in the tree"""
        it = self.tree.currentItem()
        return it is not None and getattr(it, "ref", False)

    _menu_target = None

    def _tree_menu(self, pos):
        it = self.tree.itemAt(pos)
        if it is not None:
            self.tree.setCurrentItem(it)
        self._menu_target = "class"
        m = QMenu(self)
        nm = m.addMenu("&New")
        for a in (self.a_new_class, self.a_new_coll, self.a_new_prop, self.a_new_method, self.a_new_event,
                  self.a_new_enum):
            nm.addAction(a)
        m.addSeparator()
        for a in (self.a_cut, self.a_copy, self.a_paste):
            m.addAction(a)
        m.addSeparator()
        m.addAction(self.a_delete)
        m.addAction(self.a_rename)
        m.addSeparator()
        m.addAction(self.a_props)
        m.addAction(self.a_set_coll)
        self._states()
        m.exec(self.tree.viewport().mapToGlobal(pos))
        self._menu_target = None

    def _list_menu(self, t, pos):
        it = t.itemAt(pos)
        if it is not None:
            t.setCurrentItem(it)
        self._menu_target = "member"
        m = QMenu(self)
        nm = m.addMenu("&New")
        for a in (self.a_new_prop, self.a_new_method, self.a_new_event, self.a_new_enum):
            nm.addAction(a)
        m.addSeparator()
        for a in (self.a_cut, self.a_copy, self.a_paste):
            m.addAction(a)
        m.addSeparator()
        m.addAction(self.a_delete)
        m.addAction(self.a_rename)
        m.addSeparator()
        m.addAction(self.a_props)
        self._states()
        m.exec(t.viewport().mapToGlobal(pos))
        self._menu_target = None

    # ------------------------------------------------------------ messages
    def message(self, text, title=TITLE):
        self.messages.append(text)
        if not self.quiet:
            QMessageBox.information(self, title, text)

    def run(self, dlg):
        """a builder: shown (its result) or, in the tests, given back to be filled"""
        self.last_dialog = dlg
        dlg.quiet = self.quiet
        if self.quiet:
            return None
        return dlg if dlg.exec() else None

    def taken_names(self, but=None):
        """the names a new class cannot have: the modules of the project and the classes of the tree"""
        names = [c.name for c in self.model.classes if c is not but]
        if self.project is not None:
            names += [m.name for m in self.project.modules if not (but and but.module is m)]
        return names

    # ------------------------------------------------------------ New
    def new_class(self, dlg=None):
        if dlg is None:
            dlg = ClassBuilderDialog(self, self.model, None, self.help_file, self.taken_names())
            if self.run(dlg) is None:
                return dlg if self.quiet else None
        return self.add_class(dlg.result_class)

    def add_class(self, c, parent=None):
        """a new class into the tree: in the class selected (a member of it, VB6 DEV-9/cb2: clsChild) or at the
        root"""
        parent = parent if parent is not None else self.current_class()
        if parent is not None and parent.collection:
            parent = None                        # a collection holds its items only
        if parent is not None:
            c.top_level = False
            parent.children.append(c.name)
        self.model.add(c)
        self.model.dirty = True
        self.fill_tree(c.name)
        return c

    def new_collection(self, dlg=None):
        if dlg is None:
            dlg = CollectionBuilderDialog(self, self.model, None, self.help_file, self.taken_names())
            if self.run(dlg) is None:
                return dlg if self.quiet else None
        c = dlg.result_class
        if dlg.new_class is not None and self.model.cls(dlg.new_class.name) is None:
            self.model.add(dlg.new_class)
        return self.add_class(c)

    def _new_member(self, cls, make):
        c = self.current_class()
        if c is None:
            return None
        dlg = make(c)
        if self.run(dlg) is None:
            return dlg if self.quiet else None
        return self.add_member(c, dlg.result_member)

    def add_member(self, c, m):
        if m.default:                            # one default member in a class
            for x in c.members:
                x.default = False if x is not m else x.default
        if c.member(m.name) is None:
            c.members.append(m)
        self.model.dirty = True
        self.class_changed()
        return m

    def new_property(self):
        return self._new_member(None, lambda c: PropertyBuilderDialog(self, self.model, c, None, self.help_file))

    def new_method(self):
        return self._new_member(None, lambda c: MethodBuilderDialog(self, self.model, c, None, self.help_file))

    def new_event(self):
        return self._new_member(None, lambda c: EventBuilderDialog(self, self.model, c, None, self.help_file))

    def new_enum(self):
        return self._new_member(None, lambda c: EnumBuilderDialog(self, self.model, c, None))

    # ------------------------------------------------------------ Edit
    def properties(self, target=None):
        """the builder of the member or class selected (of one in the code already: only to look at)"""
        c = self.current_class()
        m = self.current_member() if self._focus_member(target) else None
        if m is not None and c is not None:
            ro = m.existing or (c.collection and not c.existing)
            mk = {"property": PropertyBuilderDialog, "method": MethodBuilderDialog, "event": EventBuilderDialog}
            if m.kind == "enum":
                dlg = EnumBuilderDialog(self, self.model, c, m, readonly=ro)
            else:
                dlg = mk[m.kind](self, self.model, c, m, self.help_file, readonly=ro)
            if self.run(dlg) is not None and not ro:
                self.model.dirty = True
                self.class_changed()
            return dlg
        if c is None:
            return None
        cls = CollectionBuilderDialog if c.collection else ClassBuilderDialog
        dlg = cls(self, self.model, c, self.help_file, self.taken_names(c))
        old = c.name
        if self.run(dlg) is not None:
            self._renamed_class(c, old)
            self.model.dirty = True
            self.fill_tree(c.name)
        return dlg

    def delete(self, target=None):
        if self._focus_member(target):
            c, m = self.current_class(), self.current_member()
            if m is None or c is None:
                return False
            if m.existing or m not in c.members:
                self.message(EXISTING_MEMBER.format(name=m.name, what="delete", cls=c.name))
                return False
            c.members.remove(m)
            self.model.dirty = True
            self.class_changed()
            return True
        c = self.current_class()
        if c is None or self._on_ref():
            return False
        if c.existing:
            self.message(EXISTING_CLASS.format(name=c.name))
            return False
        self._remove_class(c)
        self.model.dirty = True
        self.fill_tree()
        return True

    def _remove_class(self, c):
        self.model.classes.remove(c)
        low = c.name.lower()
        for k in self.model.classes:
            k.children = [n for n in k.children if n.lower() != low]
            if k.collection and k.item.lower() == low:
                k.item = ""

    def rename(self, target=None):
        """Rename: the name of the class (tree) or of the member (list) becomes editable"""
        if self._focus_member(target):
            t = self.current_list()
            it = t.currentItem()
            if it is not None:
                t.editItem(it, 0)
            return
        it = self.tree.currentItem()
        if it is not None and getattr(it, "cls", None) is not None:
            self.tree.editItem(it, 0)

    def rename_class(self, c, new):
        """a class gets a new name (a module of the project: at Update Project, as its VB_Name)"""
        new = (new or "").strip()
        if new == c.name:
            return True
        if not C.valid_name(new):
            self.message("Invalid Name!")
            return False
        if new.lower() in {n.lower() for n in self.taken_names(c)}:
            self.message(f"'{new}' already exists!")
            return False
        old = c.name
        c.name = new
        self._renamed_class(c, old)
        self.model.dirty = True
        self.fill_tree(new)
        return True

    def _renamed_class(self, c, old):
        if old == c.name:
            return
        if c.existing and c.old_name is None:
            c.old_name = old
        low = old.lower()
        for k in self.model.classes:
            k.children = [c.name if n.lower() == low else n for n in k.children]
            if k.collection and k.item.lower() == low:
                k.item = c.name

    def rename_member(self, c, m, new):
        new = (new or "").strip()
        if new == m.name:
            return True
        if m.existing or m not in c.members:
            self.message(EXISTING_MEMBER.format(name=m.name, what="rename", cls=c.name))
            return False
        if not C.valid_name(new):
            self.message("Invalid Name!")
            return False
        if new.lower() in {x.name.lower() for x in self.model.display_members(c) if x is not m}:
            self.message(f"'{new}' already exists!")
            return False
        m.name = new
        self.model.dirty = True
        self.class_changed()
        return True

    def _tree_renamed(self, it, col):
        if self._filling or getattr(it, "cls", None) is None:
            return
        c = it.cls
        if not self.rename_class(c, it.text(0)):
            self._filling = True
            it.setText(0, c.name)
            self._filling = False

    def _member_renamed(self, it, col):
        if self._filling or col != 0 or getattr(it, "member", None) is None:
            return
        m = it.member
        if not self.rename_member(self.current_class(), m, it.text(0)):
            self._filling = True
            it.setText(0, m.name)
            self._filling = False

    def cut(self, target=None):
        return self.copy(cut=True, target=target)

    def copy(self, cut=False, target=None):
        if self._focus_member(target):
            c, m = self.current_class(), self.current_member()
            if m is None:
                return None
            if cut and (m.existing or m not in c.members):
                self.message(EXISTING_MEMBER.format(name=m.name, what="cut", cls=c.name))
                return None
            self.clip = ("member", m.copy(), c.name, cut)
            if cut:
                c.members.remove(m)
                self.model.dirty = True
                self.class_changed()
        else:
            c = self.current_class()
            if c is None or self._on_ref():
                return None
            if cut and c.existing:
                self.message(EXISTING_CLASS.format(name=c.name))
                return None
            self.clip = ("class", c.name, cut)
        self._states()
        return self.clip

    def paste(self):
        """a member into the current class ("'X' already exists!": DEV-9/cb2-07); a class cut: into the current
        class (or the root); a class copied: a new class with its members"""
        if self.clip is None:
            return None
        target = self.current_class()
        if self.clip[0] == "member":
            if target is None:
                return None
            m = self.clip[1].copy()
            if any(x.name.lower() == m.name.lower() for x in self.model.display_members(target)):
                self.message(f"'{m.name}' already exists!")
                return None
            target.members.append(m)
            self.model.dirty = True
            self.class_changed()
            return m
        _, name, cut = self.clip
        c = self.model.cls(name)
        if c is None:
            return None
        if cut:
            if target is c:
                return None
            for k in self.model.classes:
                k.children = [n for n in k.children if n.lower() != c.name.lower()]
            if target is not None and not target.collection:
                target.children.append(c.name)
                c.top_level = False
            else:
                c.top_level = True
            self.clip = None
            self.model.dirty = True
            self.fill_tree(c.name)
            return c
        n = self.model.unique(c.name)
        k = C.CBClass(n, c.collection, c.item)
        k.members = [m.copy() for m in c.members]
        k.description, k.helpid = c.description, c.helpid
        return self.add_class(k)

    def set_as_collection(self):
        """Set As Collection: the class is marked a collection (VB_Ext_KEY "Collection" "UserDefinedCollection",
        DEV-9/cb4), its code stays as it is"""
        c = self.current_class()
        if c is None or c.collection:
            return
        c.user_collection = not c.user_collection
        self.model.dirty = True
        self.fill_tree(c.name)

    # ------------------------------------------------------------ View, Help
    def options(self, debug=None, errors=None):
        if debug is None:
            d = OptionsDialog(self, self.model.debug_code, self.model.err_raise)
            self.last_dialog = d
            if self.quiet or not d.exec():
                return d
            debug, errors = d.cb_debug.isChecked(), d.cb_err.isChecked()
        self.model.debug_code, self.model.err_raise = debug, errors
        S = getattr(self.main, "S", None)
        if S is not None:
            S["cb_debug_code"], S["cb_err_raise"] = bool(debug), bool(errors)
            S.save()
        return None

    def help(self):
        if hasattr(self.main, "help_keyword"):
            self.main.help_keyword("Class Builder")

    # ------------------------------------------------------------ Update Project
    def update_project(self):
        """the new classes, the new members of the old ones and the place of each in the tree into the project
        (the old code is not changed): cbmodel.update_project"""
        if self.project is None:
            return False
        main = self.main

        def rename(m, new):
            err = main.rename_module(m, new) if hasattr(main, "rename_module") else None
            if err:
                self.message(err)

        def set_code(m, text):
            sw = getattr(main, "code_windows", {}).get(m)
            cw = sw.widget() if sw is not None else None
            if cw is not None and hasattr(main, "_set_code_text"):
                main._set_code_text(cw, text)
                m.set_code(cw.text())
            else:
                m.editor_text = None
                m.sf.code = text.split("\n")
        C.update_project(self.model, self.project, rename, set_code)
        self.updated = True
        if hasattr(main, "fill_tree"):
            main.fill_tree()
            main.update_title()
        self.class_changed()
        return True

    # ------------------------------------------------------------ the end
    def ask_update(self):
        """Exit with changes: "Update Project with changes?" -> yes / no / cancel"""
        if self.quiet:
            return getattr(self, "answer", "no")
        B = QMessageBox.StandardButton
        r = QMessageBox.question(self, TITLE, "Update Project with changes?", B.Yes | B.No | B.Cancel)
        return {B.Yes: "yes", B.No: "no"}.get(r, "cancel")

    _may_close = False

    def may_close(self):
        """File > Exit, Esc, the close button: with changes, "Update Project with changes?" (DEV-9/cb3-09)"""
        if not self.model.dirty:
            return True
        self.messages.append("Update Project with changes?")
        a = self.ask_update()
        if a == "cancel":
            return False
        if a == "yes":
            self.update_project()
        return True

    def closeEvent(self, ev):
        if not self._may_close and not self.may_close():
            ev.ignore()
            return
        self._may_close = True
        super().closeEvent(ev)                  # (QDialog: reject)
        self._may_close = False

    def reject(self):
        if not self._may_close and not self.may_close():
            return
        self._may_close = True
        super().reject()
        self._may_close = False

    def foreign_message(self):
        """the classes the Class Builder did not make: once, unless "Do not show this dialog in the future" """
        if not any(c.foreign for c in self.model.classes):
            return None
        S = getattr(self.main, "S", None)
        if S is not None and S["cb_no_foreign_msg"]:
            return None
        self.messages.append("foreign")
        d = ForeignDialog(self)
        self.last_dialog = d
        if not self.quiet:
            d.exec()
            if d.again.isChecked() and S is not None:
                S["cb_no_foreign_msg"] = True
                S.save()
        return d


def open_builder(main):
    """Add-Ins > Class Builder Utility..., Project > Add Class Module > VB Class Builder"""
    if getattr(main, "project", None) is None:
        return None
    if hasattr(main, "_sync_editors"):
        main._sync_editors()
    dlg = ClassBuilder(main)
    main._classbuilder = dlg                     # the tests reach it
    dlg.show()
    dlg.foreign_message()
    dlg.exec()
    return dlg


# ================================================================ Project > Add Class Module
DATA_SOURCE = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "{name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = {exposed}
Option Explicit
' A data source: the controls whose DataSource is an object of this class get their records from it.
' Keep the recordsets here, for instance:
'     Private rsMain As ADODB.Recordset       (Project > References: Microsoft ActiveX Data Objects)
' and open them in Class_Initialize.

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    ' the recordset of the member a control asks for ("" is the default one), for instance:
    '     Select Case DataMember
    '     Case ""
    '         Set Data = rsMain
    '     End Select
End Sub
"""

COMPLEX_CONSUMER = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "{name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = {exposed}
Option Explicit
' A consumer of the rows of a data source (a data control, a data environment, a data source class).
' DataSource and DataMember say where the rows come from; keep the source and the member here and
' take the rows from it, for instance into
'     Private rsRows As ADODB.Recordset       (Project > References: Microsoft ActiveX Data Objects)
Private mSource As Object
Private mMember As String

Public Property Get DataSource() As Object
    Set DataSource = mSource
End Property

Public Property Set DataSource(ByVal objDataSource As Object)
    Set mSource = objDataSource
    PropertyChanged "DataSource"
End Property

Public Property Get DataMember() As String
    DataMember = mMember
End Property

Public Property Let DataMember(ByVal Member As String)
    mMember = Member
    PropertyChanged "DataMember"
End Property
"""

TEMPLATES = [("Class Module", "class"), ("VB Class Builder", "builder"), ("Complex Data Consumer", "consumer"),
             ("Data Source", "source")]


def template_picture(kind, size=32):
    if kind == "class":
        return picture("class", size)
    if kind == "builder":
        return picture("builder", size)
    pm, p = _pm(size)
    _box(p, 2, 3, 9, 10, QColor(255, 255, 255))
    p.setBrush(QColor(130, 60, 170))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawRect(QRectF(3, 4, 8, 3))
    p.setPen(QPen(QColor(40, 40, 40), 1))
    p.setBrush(QColor(240, 220, 150) if kind == "source" else QColor(0, 128, 128))
    p.drawEllipse(QRectF(9, 8, 6, 3))
    p.drawRect(QRectF(9, 9.5, 6, 4))
    p.drawEllipse(QRectF(9, 12, 6, 3))
    if kind == "consumer":
        p.setPen(QPen(QColor(200, 40, 40), 1.5))
        p.drawLine(QPointF(12, 7), QPointF(9, 4))
    p.end()
    return pm


class AddClassModuleDialog(QDialog):
    """Project > Add Class Module: New (Class Module, VB Class Builder, Complex Data Consumer, Data Source) or
    Existing (a .cls), "Don't show this dialog in the future" (VB6 DEV-9/cb1-01)"""

    def __init__(self, parent, start_dir):
        from PyQt6.QtWidgets import QListWidgetItem, QListView, QFileDialog
        super().__init__(parent)
        self.setWindowTitle("Add Class Module")
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        new = QWidget()
        nl = QHBoxLayout(new)
        self.items = QListWidget()
        self.items.setViewMode(QListView.ViewMode.IconMode)
        self.items.setIconSize(self.items.iconSize() * 2)
        self.items.setMovement(QListView.Movement.Static)
        self.items.setWordWrap(True)
        self.items.setGridSize(self.items.gridSize().expandedTo(self.items.iconSize() * 3))
        for text, kind in TEMPLATES:
            it = QListWidgetItem(QIcon(template_picture(kind)), text)
            it.setData(Qt.ItemDataRole.UserRole, kind)
            self.items.addItem(it)
        self.items.setCurrentRow(0)
        self.items.itemDoubleClicked.connect(lambda *_: self._open())
        nl.addWidget(self.items, 1)
        self.tabs.addTab(new, "New")
        self.files = QFileDialog(self, "Existing", start_dir, "Class Files (*.cls);;All Files (*)")
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
        self.tabs.currentChanged.connect(lambda i: [b.setVisible(i == 0) for b in (self.b_open, self.b_cancel,
                                                                                   self.b_help)])
        self.path = None
        self.kind = "class"
        self.resize(540, 380)

    def _open(self):
        if self.tabs.currentIndex() == 1:
            sel = self.files.selectedFiles()
            if not sel or not os.path.isfile(sel[0]):
                return
            self.path = sel[0]
        else:
            it = self.items.currentItem()
            self.kind = it.data(Qt.ItemDataRole.UserRole) if it is not None else "class"
        self.accept()


def add_template(project, kind, name=None):
    """a class module of a template of Add Class Module (Data Source, Complex Data Consumer): our own code"""
    from .vbmodel import Module, parse_source
    base = {"source": "clsDataSource", "consumer": "clsComplexDataConsumer"}[kind]
    if name is None:
        name = base
        i = 1
        while project.module(name) is not None:
            i += 1
            name = f"{base}{i}"
    exposed = project.ptype in ("OleExe", "OleDll")
    text = (DATA_SOURCE if kind == "source" else COMPLEX_CONSUMER).replace("{name}", name).replace(
        "{exposed}", "True" if exposed else "False")
    sf = parse_source(text.replace("\n", "\r\n"), None, "Class", True, project.encoding)
    sf.encoding = project.encoding
    m = Module(project, "Class", name, name + ".cls", sf)
    m.dirty = True
    project.modules.append(m)
    project._vbp_add(m)
    project.dirty = True
    return m


__all__ = ["ClassBuilder", "open_builder", "builder_icon", "ADDIN", "AddClassModuleDialog", "add_template",
           "ClassBuilderDialog", "CollectionBuilderDialog", "PropertyBuilderDialog", "MethodBuilderDialog",
           "EventBuilderDialog", "EnumBuilderDialog", "ArgumentDialog", "OptionsDialog", "picture"]
