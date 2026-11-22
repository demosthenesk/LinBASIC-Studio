"""The ActiveX Control Interface Wizard (phase 9.2d): Add-Ins > ActiveX Control Interface Wizard... when the
"VB 6 ActiveX Ctrl Interface Wizard" add-in is loaded (VB6 oracle DEV-9/ax1, ax2).

The pages of VB6's: Introduction (Skip this screen in the future), Select a Control (a project with more than one
user control), Select Interface Members (Available names / Selected names, > >> < <<), Create Custom Interface
Members (New... / Edit... / Delete: Add Custom Member), Set Mapping (Public Name -> Maps to Control, Member), Set
Attributes (the unmapped members: Data Type, Default Value, Run Time, Design Time, Arguments, Description),
Finished! (View Summary Report). The code is axwizgen's; the texts and pictures are our own.
"""
from PyQt6.QtCore import Qt, QRectF, QPointF
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPen, QIcon, QPolygonF
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton, QLineEdit,
                             QListWidget, QStackedWidget, QFrame, QCheckBox, QComboBox, QGroupBox,
                             QRadioButton, QMessageBox, QPlainTextEdit, QTreeWidget, QTreeWidgetItem,
                             QAbstractItemView, QFileDialog)

from . import axwizgen as G
from .appwizard import wizard_picture

TITLE = "ActiveX Control Interface Wizard"
ADDIN = "VB 6 ActiveX Ctrl Interface Wizard"
PAGE_TITLES = {"intro": "Introduction", "control": "Select a Control", "members": "Select Interface Members",
               "custom": "Create Custom Interface Members", "mapping": "Set Mapping", "attributes": "Set Attributes",
               "finished": "Finished!"}
KIND_TEXT = {"property": "Property", "method": "Method", "event": "Event"}


def ax_picture(kind, w=150, h=120):
    """the picture on the left of a page: our own drawings"""
    if kind == "finished":
        return wizard_picture("finished", w, h)
    pm = QPixmap(w, h)
    pm.fill(QColor("#808080"))
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    ink, paper, blue, gold, green = QColor("#202020"), QColor("#FFFFFF"), QColor("#2F6FD0"), QColor("#F2C230"), \
        QColor("#30A040")

    def window(x, y, ww, hh):
        p.setPen(QPen(ink, 1))
        p.setBrush(paper)
        p.drawRect(QRectF(x, y, ww, hh))
        p.setBrush(blue)
        p.drawRect(QRectF(x, y, ww, 7))

    def bolt(x, y):
        p.setPen(QPen(QColor("#806000"), 1))
        p.setBrush(gold)
        p.drawPolygon(QPolygonF([QPointF(x + 5, y), QPointF(x, y + 8), QPointF(x + 4, y + 8), QPointF(x + 2, y + 15),
                                 QPointF(x + 9, y + 5), QPointF(x + 5, y + 5), QPointF(x + 7, y)]))

    def method(x, y):
        p.setPen(QPen(QColor("#205020"), 1))
        p.setBrush(green)
        p.drawPolygon(QPolygonF([QPointF(x, y + 5), QPointF(x + 6, y), QPointF(x + 12, y + 5), QPointF(x + 6, y + 9)]))

    def prop(x, y):
        p.setPen(QPen(ink, 1))
        p.setBrush(paper)
        p.drawRect(QRectF(x, y, 12, 9))
        p.drawLine(QPointF(x + 2, y + 3), QPointF(x + 10, y + 3))
        p.drawLine(QPointF(x + 2, y + 6), QPointF(x + 10, y + 6))
    window(40, 22, 80, 64)
    p.setPen(QPen(ink, 1, Qt.PenStyle.DotLine))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawRect(QRectF(46, 34, 68, 46))
    if kind in ("intro", "members"):
        prop(14, 18)
        method(14, 44)
        bolt(16, 66)
    elif kind == "control":
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor("#D0D0D0"))
        p.drawRect(QRectF(82, 58, 34, 30))
        prop(92, 66)
    elif kind == "custom":
        p.setPen(QPen(ink, 1))
        p.setBrush(gold)
        p.drawEllipse(QRectF(14, 10, 26, 18))
        method(20, 15)
        bolt(70, 46)
    elif kind in ("mapping", "attributes"):
        prop(12, 30)
        method(12, 52)
        bolt(14, 70)
        p.setPen(QPen(ink, 1))
        for i in range(4):
            p.drawLine(QPointF(52, 44 + i * 9), QPointF(80, 44 + i * 9))
            p.drawRect(QRectF(86, 41 + i * 9, 20, 6))
    p.end()
    return pm


def axwizard_icon(size=16):
    """the icon of Add-Ins > ActiveX Control Interface Wizard..."""
    pm = ax_picture("custom").copy(8, 6, 116, 90).scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio,
                                                         Qt.TransformationMode.SmoothTransformation)
    return QIcon(pm)


class CustomMemberDialog(QDialog):
    """Add Custom Member / Edit Custom Member: Name, Type (Property / Method / Event)"""

    def __init__(self, parent, name="", kind="property", title="Add Custom Member"):
        super().__init__(parent)
        self.setWindowTitle(title)
        lay = QHBoxLayout(self)
        left = QVBoxLayout()
        row = QHBoxLayout()
        row.addWidget(QLabel("&Name:"))
        self.name = QLineEdit(name)
        row.addWidget(self.name, 1)
        left.addLayout(row)
        g = QGroupBox("Type")
        gl = QVBoxLayout(g)
        self.kinds = {}
        for k, text in (("property", "&Property"), ("method", "&Method"), ("event", "&Event")):
            rb = QRadioButton(text)
            gl.addWidget(rb)
            self.kinds[k] = rb
        self.kinds.get(kind, self.kinds["property"]).setChecked(True)
        left.addWidget(g)
        lay.addLayout(left, 1)
        col = QVBoxLayout()
        b = QPushButton("OK")
        b.setDefault(True)
        b.clicked.connect(self.accept)
        c = QPushButton("Cancel")
        c.clicked.connect(self.reject)
        h = QPushButton("&Help")
        h.clicked.connect(lambda: parent.help() if hasattr(parent, "help") else None)
        for x in (b, c, h):
            col.addWidget(x)
        col.addStretch(1)
        lay.addLayout(col)

    def kind(self):
        return next(k for k, rb in self.kinds.items() if rb.isChecked())


class SummaryDialog(QDialog):
    """after Finish with View Summary Report: the report (Save / Close)"""

    def __init__(self, parent, text):
        super().__init__(parent)
        self.setWindowTitle(TITLE + " Summary")
        lay = QHBoxLayout(self)
        self.text = QPlainTextEdit(text)
        self.text.setReadOnly(True)
        lay.addWidget(self.text, 1)
        col = QVBoxLayout()
        s = QPushButton("&Save")
        s.clicked.connect(self.save)
        c = QPushButton("&Close")
        c.clicked.connect(self.accept)
        col.addWidget(s)
        col.addWidget(c)
        col.addStretch(1)
        lay.addLayout(col)
        self.resize(720, 460)

    def save(self, path=None):
        if path is None:
            path, _ = QFileDialog.getSaveFileName(self, "Save Summary Report", "ctlwiz.txt", "Text Files (*.txt)")
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.text.toPlainText())
        return path


class AxWizard(QDialog):
    """the pages of the wizard and Finish"""

    def __init__(self, main, project=None, quiet=False):
        super().__init__(main)
        self.main = main
        self.project = project if project is not None else getattr(main, "project", None)
        self.quiet = quiet
        self.messages = []
        self.ucs = [m for m in (self.project.modules if self.project else []) if m.kind == "UserControl"]
        self.uc = self.ucs[0] if self.ucs else None
        self.model = None
        self.done_text = None
        outer = QVBoxLayout(self)
        body = QHBoxLayout()
        self.picture = QLabel()
        self.picture.setFixedSize(150, 120)
        body.addWidget(self.picture, 0, Qt.AlignmentFlag.AlignTop)
        self.stack = QStackedWidget()
        body.addWidget(self.stack, 1)
        outer.addLayout(body, 1)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        outer.addWidget(line)
        row = QHBoxLayout()
        b = QPushButton("&Help")
        b.clicked.connect(self.help)
        row.addWidget(b)
        row.addStretch(1)
        self.b_cancel = QPushButton("Cancel")
        self.b_back = QPushButton("< &Back")
        self.b_next = QPushButton("&Next >")
        self.b_finish = QPushButton("&Finish")
        for b in (self.b_cancel, self.b_back, self.b_next, self.b_finish):
            row.addWidget(b)
        self.b_cancel.clicked.connect(self.reject)
        self.b_back.clicked.connect(lambda: self.step(-1))
        self.b_next.clicked.connect(lambda: self.step(1))
        self.b_finish.clicked.connect(self.finish)
        outer.addLayout(row)
        self.pages = {}
        for key, build in (("intro", self._intro_page), ("control", self._control_page),
                           ("members", self._members_page), ("custom", self._custom_page),
                           ("mapping", self._mapping_page), ("attributes", self._attributes_page),
                           ("finished", self._finished_page)):
            w = QFrame()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(6, 0, 0, 0)
            build(lay)
            self.pages[key] = self.stack.addWidget(w)
        self.resize(640, 460)
        S = getattr(main, "S", None)
        skip = bool(S["axwiz_skip_intro"]) if S is not None else False
        self.key = "intro"
        self.show_page(self.page_keys()[1] if skip and len(self.page_keys()) > 1 else "intro")

    # ------------------------------------------------------------ pages
    @staticmethod
    def _text(lay, text):
        t = QLabel(text)
        t.setWordWrap(True)
        lay.addWidget(t)
        return t

    def _intro_page(self, lay):
        self._text(lay, "The ActiveX Control Interface Wizard writes the interface of a user control: its "
                        "properties, methods and events, and the code under them.\n\n"
                        "Put the controls that make up the user control on it before you start: the wizard can "
                        "map members of your control to members of them.\n\n"
                        "The Property Page Wizard makes the property pages of the control.\n\n"
                        "You can go Back at any time to change what you chose. Click Next to begin.")
        lay.addStretch(1)
        self.cb_skip = QCheckBox("&Skip this screen in the future.")
        lay.addWidget(self.cb_skip)

    def _control_page(self, lay):
        self._text(lay, "Choose the user control whose properties, methods and events you want to write.")
        self.control_list = QListWidget()
        for m in self.ucs:
            self.control_list.addItem(f"{self.project.name}:{m.name}")
        self.control_list.setCurrentRow(0)
        lay.addWidget(self.control_list, 1)

    def _members_page(self, lay):
        self._text(lay, "These are the names of properties, methods and events your control can have. The ones on "
                        "the right are chosen; at first they are the standard members most controls have.\n\n"
                        "Click Next to add members of your own.")
        grid = QGridLayout()
        grid.addWidget(QLabel("&Available names:"), 0, 0)
        grid.addWidget(QLabel("&Selected names:"), 0, 2)
        self.avail = self._two_column_list()
        self.chosen = self._two_column_list()
        grid.addWidget(self.avail, 1, 0)
        grid.addWidget(self.chosen, 1, 2)
        col = QVBoxLayout()
        col.addStretch(1)
        for text, slot in ((">", lambda: self.move_members(1, False)), (">>", lambda: self.move_members(1, True)),
                           ("<", lambda: self.move_members(-1, False)), ("<<", lambda: self.move_members(-1, True))):
            b = QPushButton(text)
            b.setFixedWidth(40)
            b.clicked.connect(slot)
            col.addWidget(b)
        col.addStretch(1)
        grid.addLayout(col, 1, 1)
        lay.addLayout(grid, 1)
        self.avail.itemDoubleClicked.connect(lambda *_: self.move_members(1, False))
        self.chosen.itemDoubleClicked.connect(lambda *_: self.move_members(-1, False))

    @staticmethod
    def _two_column_list():
        t = QTreeWidget()
        t.setColumnCount(2)
        t.setHeaderHidden(True)
        t.setRootIsDecorated(False)
        t.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        t.setColumnWidth(0, 140)
        return t

    def _custom_page(self, lay):
        self._text(lay, "Add the properties, methods and events that are your control's own: those you did not "
                        "choose on the page before.")
        row = QHBoxLayout()
        left = QVBoxLayout()
        left.addWidget(QLabel("&My Custom Members:"))
        self.custom_list = self._two_column_list()
        self.custom_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        left.addWidget(self.custom_list, 1)
        row.addLayout(left, 1)
        col = QVBoxLayout()
        self.b_new = QPushButton("Ne&w...")
        self.b_edit = QPushButton("&Edit...")
        self.b_delete = QPushButton("&Delete")
        self.b_new.clicked.connect(lambda: self.new_custom())
        self.b_edit.clicked.connect(lambda: self.edit_custom())
        self.b_delete.clicked.connect(self.delete_custom)
        for b in (self.b_new, self.b_edit, self.b_delete):
            col.addWidget(b)
        col.addStretch(1)
        row.addLayout(col)
        lay.addLayout(row, 1)
        self.custom_list.currentItemChanged.connect(lambda *_: self._custom_states())

    def _mapping_page(self, lay):
        self._text(lay, "A member of your control can do what a member of a control on it (or of the UserControl) "
                        "does. Choose one or more names in the Public Name list, then the control and the member "
                        "they map to.\n\nWhen the mapping is done, click Next.")
        row = QHBoxLayout()
        left = QVBoxLayout()
        left.addWidget(QLabel("&Public Name:"))
        self.map_names = self._two_column_list()
        left.addWidget(self.map_names, 1)
        row.addLayout(left, 1)
        g = QGroupBox("Maps to")
        gl = QVBoxLayout(g)
        gl.addWidget(QLabel("&Control"))
        self.map_ctl = QComboBox()
        gl.addWidget(self.map_ctl)
        gl.addWidget(QLabel("&Member"))
        self.map_member = QComboBox()
        gl.addWidget(self.map_member)
        gl.addStretch(1)
        row.addWidget(g, 1)
        lay.addLayout(row, 1)
        self.map_names.itemSelectionChanged.connect(self.mapping_selected)
        self.map_ctl.activated.connect(lambda *_: self.mapping_control_changed())
        self.map_member.activated.connect(lambda *_: self.mapping_member_changed())

    def _attributes_page(self, lay):
        self._text(lay, "Set the attributes of the members that are not mapped; a member you do not change keeps "
                        "its defaults.\n\nChoose a name and change its values, for as many members as you need. "
                        "When you are done, click Next.")
        row = QHBoxLayout()
        left = QVBoxLayout()
        left.addWidget(QLabel("&Public Name:"))
        self.attr_names = self._two_column_list()
        self.attr_names.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        left.addWidget(self.attr_names, 1)
        left.addWidget(QLabel("&Description"))
        self.attr_desc = QPlainTextEdit()
        self.attr_desc.setMaximumHeight(50)
        left.addWidget(self.attr_desc)
        row.addLayout(left, 1)
        g = QGroupBox("Attribute Information")
        gl = QGridLayout(g)
        gl.addWidget(QLabel("Data &Type"), 0, 0)
        gl.addWidget(QLabel("Default &Value"), 0, 1)
        self.attr_type = QComboBox()
        self.attr_type.addItems(G.DATA_TYPES)
        self.attr_default = QLineEdit()
        gl.addWidget(self.attr_type, 1, 0)
        gl.addWidget(self.attr_default, 1, 1)
        gl.addWidget(QLabel("&Run Time"), 2, 0)
        gl.addWidget(QLabel("Desi&gn Time"), 2, 1)
        self.attr_run = QComboBox()
        self.attr_run.addItems(G.RUN_TIME)
        self.attr_design = QComboBox()
        self.attr_design.addItems(G.DESIGN_TIME)
        gl.addWidget(self.attr_run, 3, 0)
        gl.addWidget(self.attr_design, 3, 1)
        gl.addWidget(QLabel("&Arguments"), 4, 0)
        self.attr_args = QLineEdit()
        gl.addWidget(self.attr_args, 5, 0, 1, 2)
        row.addWidget(g, 1)
        lay.addLayout(row, 1)
        self.attr_names.currentItemChanged.connect(lambda cur, prev: self.attribute_selected(cur, prev))

    def _finished_page(self, lay):
        self._text(lay, "The wizard knows all it needs to write the interface of your control.\n\n"
                        "Click Finish to write it.")
        lay.addStretch(1)
        lay.addWidget(QLabel("Do you want to:"))
        self.cb_summary = QCheckBox("&View Summary Report")
        self.cb_summary.setChecked(True)
        lay.addWidget(self.cb_summary)

    # ------------------------------------------------------------ the order of the pages
    def page_keys(self):
        keys = ["intro"]
        if len(self.ucs) > 1:
            keys.append("control")
        return keys + ["members", "custom", "mapping", "attributes", "finished"]

    def show_page(self, key):
        self.key = key
        keys = self.page_keys()
        self.stack.setCurrentIndex(self.pages[key])
        self.setWindowTitle(f"{TITLE} - {PAGE_TITLES[key]}")
        self.picture.setPixmap(ax_picture(key))
        i = keys.index(key)
        self.b_back.setEnabled(i > 0)
        self.b_next.setEnabled(i < len(keys) - 1)
        self.b_finish.setEnabled(i >= keys.index("members"))
        (self.b_finish if i == len(keys) - 1 else self.b_next).setDefault(True)
        self.enter_page(key)

    def step(self, d):
        keys = self.page_keys()
        self.read_page(self.key)
        if d > 0:
            why = self.check_page(self.key)
            if why:
                self.message(why)
                return
        i = keys.index(self.key)
        self.show_page(keys[max(0, min(len(keys) - 1, i + d))])

    def go_to(self, key):
        """the tests: straight to a page"""
        self.read_page(self.key)
        self.show_page(key)

    def message(self, text):
        self.messages.append(text)
        if not self.quiet:
            QMessageBox.warning(self, TITLE, text)

    def help(self):
        if hasattr(self.main, "help_keyword"):
            self.main.help_keyword("ActiveX Control Interface Wizard")

    def check_page(self, key):
        if key in ("intro", "control") and self.uc is None:
            return "The project has no user control: add one with Project > Add User Control first."
        return None

    # ------------------------------------------------------------ the model of the user control
    def load_control(self, uc):
        """the members of a user control: those the wizard wrote the last time, else the standard ones"""
        self.uc = uc
        ctls = [(n, t.split(".")[-1]) for n, t, _ in uc.controls()] if uc is not None else []
        self.model = G.AxModel(uc.name, ctls, G.designer_values(uc))
        found = G.read_code(self.model, uc.code_text())
        if not found:
            for n in G.DEFAULT_SELECTED:
                self.model.select(n)
        return self.model

    def read_page(self, key):
        S = getattr(self.main, "S", None)
        if key == "intro" and S is not None and S["axwiz_skip_intro"] != self.cb_skip.isChecked():
            S["axwiz_skip_intro"] = self.cb_skip.isChecked()
            S.save()
        if key in ("intro", "control"):
            uc = self.ucs[max(0, self.control_list.currentRow())] if key == "control" and self.ucs else self.uc
            if uc is not None and (self.model is None or self.model.uc != uc.name):
                self.load_control(uc)
        if key == "attributes":
            self.attribute_selected(None, self.attr_names.currentItem())

    def enter_page(self, key):
        if self.model is None and self.uc is not None and key not in ("intro", "control"):
            self.load_control(self.uc)
        if key == "intro":
            S = getattr(self.main, "S", None)
            self.cb_skip.setChecked(bool(S["axwiz_skip_intro"]) if S is not None else False)
        elif key == "members":
            self.fill_members()
        elif key == "custom":
            self.fill_custom()
        elif key == "mapping":
            self.fill_names(self.map_names, self.model.all_names())
            self.map_ctl.clear()
            self.map_ctl.addItems(self.model.map_controls())
            self.mapping_selected()
        elif key == "attributes":
            self.fill_names(self.attr_names, self.model.unmapped())
            if self.attr_names.topLevelItemCount():
                self.attr_names.setCurrentItem(self.attr_names.topLevelItem(0))

    # ------------------------------------------------------------ Select Interface Members
    def fill_names(self, tree, names):
        tree.clear()
        for n in names:
            it = QTreeWidgetItem([n, KIND_TEXT.get(self.model.kind_of(n), "")])
            tree.addTopLevelItem(it)

    def fill_members(self):
        self.fill_names(self.avail, self.model.available_names())
        self.fill_names(self.chosen, sorted(self.model.selected, key=str.lower))

    def move_members(self, d, everything, names=None):
        src = self.avail if d > 0 else self.chosen
        if names is None:
            items = [src.topLevelItem(i) for i in range(src.topLevelItemCount())] if everything else \
                src.selectedItems()
            names = [it.text(0) for it in items]
        for n in names:
            if d > 0:
                self.model.select(n)
            else:
                self.model.unselect(n)
        self.fill_members()

    # ------------------------------------------------------------ Create Custom Interface Members
    def fill_custom(self):
        self.custom_list.clear()
        for n in sorted(self.model.custom, key=str.lower):
            self.custom_list.addTopLevelItem(QTreeWidgetItem([n, KIND_TEXT[self.model.member(n).kind]]))
        self._custom_states()

    def _custom_states(self):
        on = self.custom_list.currentItem() is not None
        self.b_edit.setEnabled(on)
        self.b_delete.setEnabled(on)

    def check_custom_name(self, name, but=None):
        from .cbmodel import NAME_RE, KEYWORDS
        if not NAME_RE.match(name or "") or name.upper() in {w.upper() for w in KEYWORDS} | {"GO", "ME"}:
            return "Invalid Name!"
        low = name.lower()
        if any(n.lower() == low for n in self.model.selected + self.model.custom if n != but):
            return f"'{name}' is already a member of the control."
        if low in {n.lower() for n in self.model.available} and low != (but or "").lower():
            return f"'{name}' is a standard member: choose it in the list of the page before."
        return None

    def new_custom(self, name=None, kind="property"):
        if name is None:
            d = CustomMemberDialog(self)
            self._last_custom_dialog = d
            if self.quiet or not d.exec():
                return d
            name, kind = d.name.text().strip(), d.kind()
        why = self.check_custom_name(name)
        if why:
            self.message(why)
            return None
        m = self.model.add_custom(name, kind)
        self.fill_custom()
        return m

    def edit_custom(self, name=None, kind=None):
        it = self.custom_list.currentItem()
        if it is None:
            return None
        old = it.text(0)
        if name is None:
            d = CustomMemberDialog(self, old, self.model.member(old).kind, "Edit Custom Member")
            self._last_custom_dialog = d
            if self.quiet or not d.exec():
                return d
            name, kind = d.name.text().strip(), d.kind()
        why = self.check_custom_name(name, old) if name.lower() != old.lower() else None
        if why:
            self.message(why)
            return None
        m = self.model.member(old)
        self.model.remove_custom(old)
        n = self.model.add_custom(name, kind or m.kind)
        if n.kind == m.kind:
            for k in ("type", "default", "runtime", "design", "args", "description"):
                setattr(n, k, getattr(m, k))
        self.fill_custom()
        return n

    def delete_custom(self):
        it = self.custom_list.currentItem()
        if it is not None:
            self.model.remove_custom(it.text(0))
            self.fill_custom()

    # ------------------------------------------------------------ Set Mapping
    def mapping_selected(self):
        names = [it.text(0) for it in self.map_names.selectedItems()]
        self.map_ctl.setEnabled(bool(names))
        self.map_member.setEnabled(bool(names))
        if not names:
            self.map_member.clear()
            return
        m = self.model.member(names[0])
        ctl = m.map_ctl or "(None)"
        i = self.map_ctl.findText(ctl)
        self.map_ctl.setCurrentIndex(max(0, i))
        self._fill_map_members(m.kind)
        j = self.map_member.findText(m.map_member or "(None)")
        self.map_member.setCurrentIndex(max(0, j))

    def _fill_map_members(self, kind):
        self.map_member.clear()
        self.map_member.addItem("(None)")
        self.map_member.addItems(self.model.map_members(self.map_ctl.currentText(), kind))

    def mapping_control_changed(self):
        names = [it.text(0) for it in self.map_names.selectedItems()]
        if not names:
            return
        kind = self.model.member(names[0]).kind
        self._fill_map_members(kind)
        # a member of the same name, when the control has one (VB6 picks it)
        for n in names:
            j = self.map_member.findText(n)
            if j > 0:
                self.map_member.setCurrentIndex(j)
                break
        self.mapping_member_changed()

    def mapping_member_changed(self):
        ctl, member = self.map_ctl.currentText(), self.map_member.currentText()
        for it in self.map_names.selectedItems():
            n = it.text(0)
            mem = member
            if len(self.map_names.selectedItems()) > 1 and self.map_member.findText(n) > 0:
                mem = n                         # several names to one control: each to its member of that name
            self.model.set_mapping(n, ctl, mem)

    def set_mapping(self, name, ctl, member):
        """the tests: Public Name -> Maps to Control, Member"""
        self.model.set_mapping(name, ctl, member)

    # ------------------------------------------------------------ Set Attributes
    def attribute_selected(self, cur, prev):
        if prev is not None:                    # the values of the member left
            m = self.model.member(prev.text(0))
            self.store_attributes(m)
        if cur is None:
            return
        m = self.model.member(cur.text(0))
        prop, method = m.kind == "property", m.kind == "method"
        self.attr_type.setEnabled(prop or method)
        self.attr_default.setEnabled(prop)
        self.attr_run.setEnabled(prop)
        self.attr_design.setEnabled(prop)
        self.attr_args.setEnabled(not prop)
        t = m.type or ("Variant" if method else "")
        if t and self.attr_type.findText(t) < 0:
            self.attr_type.addItem(t)
        self.attr_type.setCurrentText(t or "Variant")
        self.attr_default.setText(m.default if prop else "")
        self.attr_run.setCurrentIndex(m.runtime)
        self.attr_design.setCurrentIndex(m.design)
        self.attr_args.setText(m.args)
        self.attr_desc.setPlainText(m.description)

    def store_attributes(self, m):
        if m.kind == "property":
            m.type, m.default = self.attr_type.currentText(), self.attr_default.text().strip()
            m.runtime, m.design = self.attr_run.currentIndex(), self.attr_design.currentIndex()
        elif m.kind == "method":
            t = self.attr_type.currentText()
            if m.custom or t != (m.type or "Variant"):
                m.type = t
            m.args = self.attr_args.text().strip()
        else:
            m.args = self.attr_args.text().strip()
        m.description = self.attr_desc.toPlainText().strip()

    def set_attributes(self, name, **kw):
        """the tests: the attributes of a member (type, default, runtime, design, args, description)"""
        m = self.model.member(name)
        for k, v in kw.items():
            setattr(m, k, v)
        cur = self.attr_names.currentItem()
        if cur is not None and cur.text(0) == name:   # the page shows the new values
            self.attribute_selected(cur, None)

    # ------------------------------------------------------------ Finish
    def finish(self):
        self.read_page(self.key)
        if self.model is None:
            if self.uc is None:
                self.message("The project has no user control: add one with Project > Add User Control first.")
                return None
            self.load_control(self.uc)
        uc = self.uc
        main = self.main
        if hasattr(main, "_sync_editors"):
            main._sync_editors()
        new = G.generate(self.model, uc.code_text())
        sw = getattr(main, "code_windows", {}).get(uc)
        cw = sw.widget() if sw is not None else None
        if cw is not None and hasattr(main, "_set_code_text"):
            main._set_code_text(cw, new.rstrip("\n") + "\n")
            uc.set_code(cw.text())
        else:
            uc.editor_text = None
            uc.sf.code = new.rstrip("\n").split("\n")
        uc.dirty = True
        if self.project is not None:
            self.project.dirty = True
        self.done_text = new
        if hasattr(main, "update_title"):
            main.update_title()
        if self.cb_summary.isChecked():
            d = SummaryDialog(self, G.summary(self.model))
            self.summary_dialog = d
            if not self.quiet:
                d.exec()
        self.accept()
        return new


def open_wizard(main):
    """Add-Ins > ActiveX Control Interface Wizard..."""
    p = getattr(main, "project", None)
    if p is None:
        return None
    if not any(m.kind == "UserControl" for m in p.modules):
        QMessageBox.information(main, TITLE, "The project has no user control: add one with Project > Add User "
                                             "Control first, with the controls that make it up.")
        return None
    if hasattr(main, "_sync_editors"):
        main._sync_editors()
    dlg = AxWizard(main)
    main._axwizard = dlg                         # the tests reach it
    dlg.exec()
    return dlg


__all__ = ["AxWizard", "open_wizard", "axwizard_icon", "ADDIN", "CustomMemberDialog", "SummaryDialog"]
