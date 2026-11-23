"""The Data Object Wizard (phase 9.2f): Add-Ins > VB 6 Data Object Wizard when the "VB 6 Data Object Wizard"
add-in is loaded (VB6 oracle DEV-9/do1).

The pages of VB6's: Introduction (Skip this screen in the future), Create Object (a Class object to which other
objects can bind data / a User Control object bound to an existing Class object). A class: Select Data Environment
Command (a tree: the project, its Data Environments, their connections and commands), Define Class Field Information
(Nullable? / Primary Key?), Define Lookup Table Information, Define and Map Insert / Update / Delete Data Command (the
parameters mapped to the fields by name), Finished! (the name of the class). A user control: Select the Data Class to
Use (the tree with the data classes under their commands), Select User Control Type, Map Class Properties to a
Control Type, Enter User Control Name. The code is dowizgen's; the texts and pictures are our own.
"""
from PyQt6.QtCore import Qt, QRectF, QPointF
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPen, QIcon
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
                             QStackedWidget, QFrame, QCheckBox, QRadioButton, QMessageBox, QTreeWidget,
                             QTreeWidgetItem, QTableWidget, QTableWidgetItem, QComboBox, QListWidget, QGroupBox,
                             QAbstractItemView, QButtonGroup, QHeaderView)

from . import dowizgen as G
from .appwizard import wizard_picture

TITLE = "Data Object Wizard"
ADDIN = "VB 6 Data Object Wizard"
MESSAGE_TITLE = "Visual Basic Data Object Wizard"
PAGE_TITLES = {"intro": "Introduction", "create": "Create Object", "command": "Select Data Environment Command",
               "fields": "Define Class Field Information", "lookup": "Define Lookup Table Information",
               "insert": "Define and Map Insert Data Command", "update": "Define and Map Update Data Command",
               "delete": "Define and Map Delete Data Command", "finished": "Finished!",
               "dataclass": "Select the Data Class to Use", "uctype": "Select User Control Type",
               "map": "Map Class Properties to a Control Type", "ucname": "Enter User Control Name"}
CLASS_PAGES = ["intro", "create", "command", "fields", "lookup", "insert", "update", "delete", "finished"]
UC_PAGES = ["intro", "create", "dataclass", "uctype", "map", "ucname"]
NONE = "(None)"


def do_picture(kind, w=150, h=120):
    """the picture on the left of a page: our own drawings (a database, its rows, a class)"""
    if kind in ("finished", "ucname"):
        return wizard_picture("finished", w, h)
    pm = QPixmap(w, h)
    pm.fill(QColor("#808080"))
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    ink, paper, blue, teal, gold = QColor("#202020"), QColor("#FFFFFF"), QColor("#2F6FD0"), QColor("#38B0C0"), \
        QColor("#F2C230")

    def drum(x, y, ww=46, hh=44):
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor("#C8C8C8"))
        p.drawRect(QRectF(x, y + 6, ww, hh - 12))
        p.drawEllipse(QRectF(x, y + hh - 12, ww, 12))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawRect(QRectF(x + 1, y + 6, ww - 1, hh - 12))
        p.setPen(QPen(ink, 1))
        p.drawLine(QPointF(x, y + 6), QPointF(x, y + hh - 6))
        p.drawLine(QPointF(x + ww, y + 6), QPointF(x + ww, y + hh - 6))
        p.setBrush(teal)
        p.drawEllipse(QRectF(x, y, ww, 12))

    def rows(x, y, n=4, ww=56):
        p.setPen(QPen(ink, 1))
        p.setBrush(paper)
        p.drawRect(QRectF(x, y, ww, 8 + n * 11))
        p.setBrush(blue)
        p.drawRect(QRectF(x, y, ww, 7))
        for i in range(n):
            yy = y + 10 + i * 11
            p.drawText(QRectF(x + 2, yy - 1, 10, 10), Qt.AlignmentFlag.AlignCenter, str(i + 1))
            p.drawLine(QPointF(x + 14, yy + 4), QPointF(x + ww - 4, yy + 4))
    f = p.font()
    f.setPixelSize(8)
    p.setFont(f)
    if kind in ("insert", "update", "delete"):
        rows(14, 46)
        p.setPen(QPen(ink, 1))
        p.setBrush(paper)
        p.drawRect(QRectF(82, 14, 54, 36))
        for i in range(3):
            p.drawText(QRectF(86, 16 + i * 11, 48, 10), "@ " + "xxxxxx"[: 6 - i])
        p.setPen(QPen(teal, 5))
        p.drawLine(QPointF(40, 44), QPointF(40, 30))
        p.drawLine(QPointF(40, 30), QPointF(76, 30))
    elif kind in ("create", "uctype"):
        for i in range(3):
            p.setPen(QPen(ink, 1))
            p.setBrush(QColor("#E8E8E8"))
            p.drawRect(QRectF(16 + i * 10, 14 + i * 24, 80, 26))
            p.setBrush(paper)
            p.drawRect(QRectF(50 + i * 10, 20 + i * 24, 40, 6))
            p.drawRect(QRectF(50 + i * 10, 30 + i * 24, 40, 6))
        rows(92, 50, 3, 44)
    elif kind == "fields":
        rows(40, 20, 5, 70)
        p.setPen(QPen(ink, 1))
        p.setBrush(gold)
        p.drawEllipse(QRectF(24, 70, 16, 16))
        p.drawRect(QRectF(30, 84, 4, 22))
        p.drawRect(QRectF(34, 98, 6, 3))
    elif kind == "lookup":
        rows(14, 20, 4, 56)
        rows(84, 50, 3, 52)
        p.setPen(QPen(gold, 3))
        p.drawLine(QPointF(70, 40), QPointF(84, 62))
    else:                                           # intro, command, dataclass, map
        drum(16, 14)
        rows(46, 40, 4, 60)
        if kind in ("dataclass", "intro"):
            p.setPen(QPen(ink, 1))
            p.setBrush(gold)
            p.drawRect(QRectF(96, 84, 34, 24))
            p.drawLine(QPointF(100, 92), QPointF(126, 92))
    p.end()
    return pm


def dowizard_icon(size=16):
    """the icon of Add-Ins > VB 6 Data Object Wizard"""
    pm = do_picture("intro").copy(10, 8, 126, 104).scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio,
                                                          Qt.TransformationMode.SmoothTransformation)
    return QIcon(pm)


class DOWizard(QDialog):
    """the pages of the wizard and Finish"""

    def __init__(self, main, project=None, quiet=False):
        super().__init__(main)
        self.main = main
        self.project = project if project is not None else getattr(main, "project", None)
        self.quiet = quiet
        self.messages = []
        self.envs = G.data_environments(self.project)
        self.dc = None                          # the class being made
        self.spec = None                        # the user control being made
        self.made = []
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
        b = QPushButton("Help")
        b.clicked.connect(self.help)
        row.addWidget(b)
        row.addStretch(1)
        self.b_cancel = QPushButton("Cancel")
        self.b_back = QPushButton("<&Back")
        self.b_next = QPushButton("&Next>")
        self.b_finish = QPushButton("&Finish")
        for b in (self.b_cancel, self.b_back, self.b_next, self.b_finish):
            row.addWidget(b)
        self.b_cancel.clicked.connect(self.reject)
        self.b_back.clicked.connect(lambda: self.step(-1))
        self.b_next.clicked.connect(lambda: self.step(1))
        self.b_finish.clicked.connect(self.finish)
        outer.addLayout(row)
        self.pages = {}
        for key in PAGE_TITLES:
            w = QFrame()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(6, 0, 0, 0)
            getattr(self, "_page_" + key)(lay)
            self.pages[key] = self.stack.addWidget(w)
        self.setFixedSize(660, 460)
        S = getattr(main, "S", None)
        skip = bool(S["dowiz_skip_intro"]) if S is not None else False
        if S is not None and S["dowiz_default_uc"]:
            self.rb_uc.setChecked(True)
        self.key = "intro"
        self.show_page("create" if skip else "intro")

    # ------------------------------------------------------------ the pages
    @staticmethod
    def _text(lay, text):
        t = QLabel(text)
        t.setWordWrap(True)
        lay.addWidget(t)
        return t

    def _page_intro(self, lay):
        self._text(lay, "The Data Object Wizard makes classes that are a source of data for other objects, and user "
                        "controls bound to those classes.\n\n"
                        "Their rows come from the commands of a Data Environment of your project: make it and its "
                        "commands before you start the wizard.\n\n"
                        "You can go Back at any time to change what you chose. Click Next to begin.")
        lay.addStretch(1)
        self.cb_skip = QCheckBox("&Skip this screen in the future.")
        lay.addWidget(self.cb_skip)

    def _page_create(self, lay):
        self._text(lay, "Choose what the wizard makes.")
        lay.addStretch(1)
        self.rb_class = QRadioButton("A &Class object to which other objects can bind data")
        self.rb_uc = QRadioButton("A &User Control object bound to an existing Class object")
        self.rb_class.setChecked(True)
        grp = QButtonGroup(self)
        grp.addButton(self.rb_class)
        grp.addButton(self.rb_uc)
        lay.addWidget(self.rb_class)
        lay.addSpacing(16)
        lay.addWidget(self.rb_uc)
        lay.addStretch(1)
        self._text(lay, "Note: a user control is bound to a class this wizard made: make the class first.")

    @staticmethod
    def _tree(lay):
        t = QTreeWidget()
        t.setHeaderHidden(True)
        lay.addWidget(t, 1)
        return t

    def _page_command(self, lay):
        self._text(lay, "Choose the command of a Data Environment the class takes its rows from.")
        self.cmd_tree = self._tree(lay)

    def _page_fields(self, lay):
        self._text(lay, "Tell the wizard which fields can be empty (Null) and which make the primary key of a row. "
                        "Answer as the database is: the class checks a row with them before it saves it.")
        self.field_table = QTableWidget(0, 3)
        self.field_table.setHorizontalHeaderLabels(["Field Name", "Nullable?", "Primary Key?"])
        self.field_table.verticalHeader().hide()
        self.field_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        lay.addWidget(self.field_table, 1)

    def _page_lookup(self, lay):
        self._text(lay, "A lookup shows a field of another command in place of a field of the rows (the name of "
                        "the publisher in place of its number). Choose them and Add.")
        row = QHBoxLayout()
        g = QGroupBox("Define lookup")
        gl = QVBoxLayout(g)
        self.lk_source = QComboBox()
        self.lk_command = QComboBox()
        self.lk_display = QComboBox()
        self.lk_on = QListWidget()
        for text, w in (("&Source field:", self.lk_source), ("Lookup &command:", self.lk_command),
                        ("&Display field:", self.lk_display), ("Lookup &on field:", self.lk_on)):
            lab = QLabel(text)
            lab.setBuddy(w)
            gl.addWidget(lab)
            gl.addWidget(w)
        self.lk_command.currentTextChanged.connect(self._lookup_fields)
        row.addWidget(g, 2)
        col = QVBoxLayout()
        col.addStretch(1)
        self.b_lk_add = QPushButton("&Add >")
        self.b_lk_remove = QPushButton("< &Remove")
        self.b_lk_add.clicked.connect(lambda: self.add_lookup())
        self.b_lk_remove.clicked.connect(lambda: self.remove_lookup())
        col.addWidget(self.b_lk_add)
        col.addWidget(self.b_lk_remove)
        col.addStretch(1)
        row.addLayout(col)
        right = QVBoxLayout()
        right.addWidget(QLabel("Selected lookup commands:"))
        self.lk_table = QTableWidget(0, 3)
        self.lk_table.setHorizontalHeaderLabels(["Lookup Field(s)", "Lookup Command", "Display Field"])
        self.lk_table.verticalHeader().hide()
        self.lk_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.lk_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        right.addWidget(self.lk_table, 1)
        row.addLayout(right, 3)
        lay.addLayout(row, 1)

    def _change_page(self, lay, kind):
        self._text(lay, "Choose the command of the Data Environment that %s a row in the database (or (None): the "
                        "change stays in the class's rows). Its parameters get the values of the fields of the "
                        "same name." % {"insert": "adds", "update": "changes", "delete": "deletes"}[kind])
        if kind == "update":
            self.cb_insert_as_update = QCheckBox("Use &Insert command for Update")
            lay.addWidget(self.cb_insert_as_update)
            self.cb_insert_as_update.toggled.connect(lambda on: self.chg_combo["update"].setEnabled(not on))
        lab = QLabel("&%s:" % kind.capitalize())
        combo = QComboBox()
        lab.setBuddy(combo)
        lay.addWidget(lab)
        lay.addWidget(combo)
        table = QTableWidget(0, 3)
        table.setHorizontalHeaderLabels(["Source Fields", "Maps To", "%s Command Parameters" % kind.capitalize()])
        table.verticalHeader().hide()
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        lay.addWidget(table, 1)
        self.chg_combo[kind] = combo
        self.chg_table[kind] = table
        combo.currentTextChanged.connect(lambda _t, k=kind: self.choose_change(k))

    def _page_insert(self, lay):
        self.chg_combo, self.chg_table = {}, {}
        self._change_page(lay, "insert")

    def _page_update(self, lay):
        self._change_page(lay, "update")

    def _page_delete(self, lay):
        self._change_page(lay, "delete")

    def _name_page(self, lay, question, attr):
        lab = QLabel(question)
        lab.setWordWrap(True)
        lay.addWidget(lab)
        e = QLineEdit()
        lab.setBuddy(e)
        lay.addWidget(e)
        setattr(self, attr, e)
        lay.addStretch(1)
        self._text(lay, "The wizard knows all it needs.\n\nClick Finish to make it.")
        lay.addStretch(1)
        cb = QCheckBox("&Save current settings as default")
        lay.addWidget(cb)
        return cb

    def _page_finished(self, lay):
        self.cb_default_class = self._name_page(lay, "What would you like to &name your Class object?", "class_name")

    def _page_dataclass(self, lay):
        self._text(lay, "Choose the class the user control is bound to (its rows are the control's).")
        self.class_tree = self._tree(lay)

    def _page_uctype(self, lay):
        self._text(lay, "What kind of user control do you want?")
        lay.addStretch(1)
        self.uc_radios = {}
        grp = QButtonGroup(self)
        for text, kind in (("&Single Record", "Single Record"), ("&Data Grid", "Data Grid"), ("&ListBox", "ListBox"),
                           ("&ComboBox", "ComboBox")):
            rb = QRadioButton(text)
            grp.addButton(rb)
            lay.addWidget(rb)
            self.uc_radios[kind] = rb
        self.uc_radios["Single Record"].setChecked(True)
        lay.addStretch(1)

    def _page_map(self, lay):
        self._text(lay, "Choose the control that shows each property of the class on the user control; (None): "
                        "the property is not shown.")
        self.map_table = QTableWidget(0, 2)
        self.map_table.setHorizontalHeaderLabels(["Property Name", "Control Type"])
        self.map_table.verticalHeader().hide()
        self.map_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        lay.addWidget(self.map_table, 1)

    def _page_ucname(self, lay):
        self.cb_default_uc = self._name_page(lay, "Enter the name of the User Control object to make:\n\n"
                                                  "User Control &Name:", "uc_name")

    # ------------------------------------------------------------ the order of the pages
    def page_keys(self):
        return UC_PAGES if self.rb_uc.isChecked() else CLASS_PAGES

    def show_page(self, key):
        self.key = key
        keys = self.page_keys()
        self.stack.setCurrentIndex(self.pages[key])
        self.setWindowTitle("%s - %s" % (TITLE, PAGE_TITLES[key]))
        self.picture.setPixmap(do_picture(key))
        i = keys.index(key)
        self.b_back.setEnabled(i > 0)
        self.b_next.setEnabled(i < len(keys) - 1)
        self.b_finish.setEnabled(i == len(keys) - 1)
        (self.b_finish if i == len(keys) - 1 else self.b_next).setDefault(True)
        self.enter_page(key)

    def step(self, d):
        keys = self.page_keys()
        if d > 0:
            why = self.read_page(self.key)
            if why:
                self.message(why)
                return
        i = keys.index(self.key)
        self.show_page(keys[max(0, min(len(keys) - 1, i + d))])

    def message(self, text, warning=True):
        self.messages.append(text)
        if not self.quiet:
            (QMessageBox.warning if warning else QMessageBox.information)(self, MESSAGE_TITLE, text)

    def help(self):
        if hasattr(self.main, "help_keyword"):
            self.main.help_keyword("Data Object Wizard")

    # ------------------------------------------------------------ entering a page
    def enter_page(self, key):
        if key == "command":
            self._fill_tree(self.cmd_tree, False)
        elif key == "fields":
            self._fill_fields()
        elif key == "lookup":
            self._fill_lookup()
        elif key in ("insert", "update", "delete"):
            self._fill_change(key)
        elif key == "dataclass":
            self._fill_tree(self.class_tree, True)
        elif key == "map":
            self._fill_map()
        elif key == "ucname" and self.spec is not None and not self.uc_name.text():
            self.uc_name.setText(self.spec.dc.root)

    def _fill_tree(self, tree, with_classes):
        """the project, its Data Environments, their connections and commands (and the data classes under them)"""
        keep = tree.currentItem().data(0, Qt.ItemDataRole.UserRole) if tree.currentItem() else None
        tree.clear()
        root = QTreeWidgetItem([self.project.name if self.project else "Project1"])
        tree.addTopLevelItem(root)
        classes = G.data_classes(self.project) if with_classes else []
        for m, de in self.envs:
            di = QTreeWidgetItem(root, [de.name])
            for cn in de.conns:
                ci = QTreeWidgetItem(di, [cn.name])
                for cmd in de.top():
                    if (cmd.conn or de.conns[0].name).lower() != cn.name.lower():
                        continue
                    it = QTreeWidgetItem(ci, [cmd.name])
                    if not with_classes:
                        it.setData(0, Qt.ItemDataRole.UserRole, ("command", de.name, cmd.name))
                    for cm, dc in classes:
                        if dc.de.lower() == de.name.lower() and dc.command.lower() == cmd.name.lower():
                            k = QTreeWidgetItem(it, [cm.name])
                            k.setData(0, Qt.ItemDataRole.UserRole, ("class", cm.name))
        tree.expandAll()
        if keep:
            for it in self._items(tree):
                if it.data(0, Qt.ItemDataRole.UserRole) == keep:
                    tree.setCurrentItem(it)

    @staticmethod
    def _items(tree):
        out, stack = [], [tree.topLevelItem(i) for i in range(tree.topLevelItemCount())]
        while stack:
            it = stack.pop(0)
            out.append(it)
            stack += [it.child(i) for i in range(it.childCount())]
        return out

    def _yes_no(self, table, r, c, on, enabled=True):
        cb = QComboBox()
        cb.addItems(["Yes", "No"])
        cb.setCurrentIndex(0 if on else 1)
        cb.setEnabled(enabled)
        table.setCellWidget(r, c, cb)
        return cb

    def _fill_fields(self):
        t = self.field_table
        t.setRowCount(len(self.dc.fields))
        for r, f in enumerate(self.dc.fields):
            it = QTableWidgetItem(f.name)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            t.setItem(r, 0, it)
            self._yes_no(t, r, 1, f.nullable)
            self._yes_no(t, r, 2, f.pk)

    def de_model(self):
        return next((de for m, de in self.envs if de.name.lower() == self.dc.de.lower()), None)

    def _fill_lookup(self):
        de = self.de_model()
        self.lk_source.clear()
        self.lk_source.addItems([f.name for f in self.dc.fields])
        self.lk_command.blockSignals(True)
        self.lk_command.clear()
        self.lk_command.addItems([c.name for c in de.top() if c.fields and c.name.lower() != self.dc.command.lower()])
        self.lk_source.setCurrentIndex(-1)       # nothing chosen yet (VB6 DEV-9/do1)
        self.lk_command.setCurrentIndex(-1)
        self.lk_command.blockSignals(False)
        self._lookup_fields()
        self._fill_lookup_table()

    def _lookup_fields(self, *_):
        de = self.de_model()
        cmd = de.command(self.lk_command.currentText()) if de is not None else None
        names = [f[0] for f in cmd.fields] if cmd is not None else []
        self.lk_display.clear()
        self.lk_display.addItems(names)
        self.lk_display.setCurrentIndex(-1)
        self.lk_on.clear()
        self.lk_on.addItems(names)
        src = self.lk_source.currentText().lower()
        for i, n in enumerate(names):
            if n.lower() == src:
                self.lk_on.setCurrentRow(i)

    def _fill_lookup_table(self):
        t = self.lk_table
        t.setRowCount(len(self.dc.lookups))
        for r, lk in enumerate(self.dc.lookups):
            for c, v in enumerate((", ".join(s for s, _ in lk.pairs), lk.command, lk.display)):
                t.setItem(r, c, QTableWidgetItem(v))
        self.b_lk_remove.setEnabled(bool(self.dc.lookups))

    def add_lookup(self, source=None, command=None, display=None, on=None):
        """Add >: the lookup of the combo boxes (the tests give them)"""
        if command is not None:
            self.lk_command.setCurrentText(command)
            self._lookup_fields()
        if source is not None:
            self.lk_source.setCurrentText(source)
        if display is not None:
            self.lk_display.setCurrentText(display)
        if on is not None:
            items = self.lk_on.findItems(on, Qt.MatchFlag.MatchFixedString)
            if items:
                self.lk_on.setCurrentItem(items[0])
        src, cmd, disp = self.lk_source.currentText(), self.lk_command.currentText(), self.lk_display.currentText()
        cur = self.lk_on.currentItem()
        if not (src and cmd and disp and cur is not None):
            self.message("Choose the source field, the lookup command, its display field and the field it is "
                         "looked up on.")
            return None
        if self.dc.lookup_of(src) is not None:
            self.message("The field '%s' has a lookup already." % src)
            return None
        f = self.dc.field(src)
        lk = G.Lookup(cmd, disp, [(f.name, cur.text())], f.nullable)
        self.dc.lookups.append(lk)
        self._fill_lookup_table()
        return lk

    def remove_lookup(self, row=None):
        r = self.lk_table.currentRow() if row is None else row
        if 0 <= r < len(self.dc.lookups):
            del self.dc.lookups[r]
        self._fill_lookup_table()

    def _fill_change(self, kind):
        de = self.de_model()
        combo = self.chg_combo[kind]
        ch = self._change(kind)
        combo.blockSignals(True)
        combo.clear()
        combo.addItem(NONE)
        combo.addItems([c.name for c in de.top() if c.name.lower() != self.dc.command.lower() and not c.fields])
        combo.setCurrentText(ch.command or NONE)
        combo.blockSignals(False)
        if kind == "update":
            self.cb_insert_as_update.setEnabled(bool(self.dc.insert.command))
            if not self.dc.insert.command:
                self.cb_insert_as_update.setChecked(False)
            combo.setEnabled(not self.cb_insert_as_update.isChecked())
        self._fill_change_table(kind)

    def _change(self, kind):
        return {"insert": self.dc.insert, "update": self.dc.update, "delete": self.dc.delete}[kind]

    def choose_change(self, kind, name=None):
        """a command chosen in the combo box of an Insert / Update / Delete page: its parameters mapped"""
        combo = self.chg_combo[kind]
        if name is not None:
            combo.setCurrentText(name)
            return self._change(kind)
        text = combo.currentText()
        de = self.de_model()
        ch = G.map_params(self.dc, de.command(text)) if text and text != NONE else G.Change()
        setattr(self.dc, kind, ch)
        self._fill_change_table(kind)
        if ch.unmapped:
            self.message(G.UNMAPPED % "\n".join(ch.unmapped))
        return ch

    def _fill_change_table(self, kind):
        ch = self._change(kind)
        t = self.chg_table[kind]
        t.setRowCount(len(ch.pairs))
        for r, (f, p) in enumerate(ch.pairs):
            for c, v in enumerate((f, "->", p)):
                t.setItem(r, c, QTableWidgetItem(v))

    def _fill_map(self):
        dc = self.spec.dc
        t = self.map_table
        t.setRowCount(len(dc.fields))
        for r, f in enumerate(dc.fields):
            it = QTableWidgetItem(f.name)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            t.setItem(r, 0, it)
            cb = QComboBox()
            cb.addItems(G.control_choices(dc, f))
            cb.setCurrentText(self.spec.control_of(f))
            t.setCellWidget(r, 1, cb)

    def set_control(self, field, kind):
        """the tests: the control of a property on the Map page"""
        for r in range(self.map_table.rowCount()):
            if self.map_table.item(r, 0).text().lower() == field.lower():
                self.map_table.cellWidget(r, 1).setCurrentText(kind)

    def set_field(self, field, nullable=None, pk=None):
        """the tests: Nullable? / Primary Key? of a field"""
        for r in range(self.field_table.rowCount()):
            if self.field_table.item(r, 0).text().lower() == field.lower():
                if nullable is not None:
                    self.field_table.cellWidget(r, 1).setCurrentIndex(0 if nullable else 1)
                if pk is not None:
                    self.field_table.cellWidget(r, 2).setCurrentIndex(0 if pk else 1)

    def select_item(self, tree, *path):
        """the tests: the item of a tree by its text (the last one of that text)"""
        found = None
        for it in self._items(tree):
            if it.text(0) == path[-1]:
                found = it
        if found is not None:
            tree.setCurrentItem(found)
        return found

    # ------------------------------------------------------------ leaving a page (Next): None or why not
    def read_page(self, key):
        S = getattr(self.main, "S", None)
        if key == "intro":
            if S is not None and S["dowiz_skip_intro"] != self.cb_skip.isChecked():
                S["dowiz_skip_intro"] = self.cb_skip.isChecked()
                S.save()
            return None
        if key == "create":
            if not self.envs or not any(de.top() for _m, de in self.envs):
                return G.NO_DE
            return None
        if key == "command":
            it = self.cmd_tree.currentItem()
            data = it.data(0, Qt.ItemDataRole.UserRole) if it is not None else None
            if not data or data[0] != "command":
                return G.NO_COMMAND
            de = next(d for _m, d in self.envs if d.name == data[1])
            cmd = de.command(data[2])
            why = G.check_command(de, cmd)
            if why:
                return why
            if self.dc is None or self.dc.de != de.name or self.dc.command != cmd.name:
                self.dc = G.new_class(self.project, de, cmd)
            return None
        if key == "fields":
            for r, f in enumerate(self.dc.fields):
                f.nullable = self.field_table.cellWidget(r, 1).currentIndex() == 0
                f.pk = self.field_table.cellWidget(r, 2).currentIndex() == 0
            return None
        if key == "update":
            self.dc.insert_as_update = self.cb_insert_as_update.isChecked() and bool(self.dc.insert.command)
            return None
        if key == "finished":
            return G.check_class_name(self.project, self.class_name.text().strip())
        if key == "dataclass":
            it = self.class_tree.currentItem()
            data = it.data(0, Qt.ItemDataRole.UserRole) if it is not None else None
            if not data or data[0] != "class":
                return G.NO_DATA_CLASS
            dc = dict((m.name.lower(), d) for m, d in G.data_classes(self.project))[data[1].lower()]
            if self.spec is None or self.spec.dc.name.lower() != dc.name.lower():
                self.spec = G.UserControlSpec(dc)
                self.uc_name.setText(dc.root)
            return None
        if key == "uctype":
            self.spec.kind = next(k for k, rb in self.uc_radios.items() if rb.isChecked())
            return None
        if key == "map":
            for r, f in enumerate(self.spec.dc.fields):
                self.spec.controls[f.name.lower()] = self.map_table.cellWidget(r, 1).currentText()
            return None
        if key == "ucname":
            self.spec.root = self.uc_name.text().strip()
            return G.check_uc_name(self.project, self.spec)
        return None

    # ------------------------------------------------------------ Finish
    def finish(self):
        why = self.read_page(self.key)
        if why:
            self.message(why)
            return None
        S = getattr(self.main, "S", None)
        if self.rb_uc.isChecked():
            if self.cb_default_uc.isChecked() and S is not None:
                S["dowiz_default_uc"] = True
                S["dowiz_default_uctype"] = self.spec.kind
                S.save()
            out = G.generate_uc(self.spec, G.has_enum(self.project))
        else:
            self.dc.root = self.class_name.text().strip()
            if self.cb_default_class.isChecked() and S is not None:
                S["dowiz_default_uc"] = False
                S.save()
            out = G.generate_class(self.dc, G.has_enum(self.project))
        self.made = add_to_project(self.main, self.project, out)
        self.accept()
        return self.made


def add_to_project(main, p, out):
    """the modules into the project (not saved, as VB6), the references and components they need"""
    from .cbmodel import _new_module
    made = []
    for kind, name, text in out.modules:
        ext = ".ctl" if kind == "UserControl" else ".cls"
        made.append(_new_module(p, kind, name, ext, text))
    p.add_references(G.needed_references(p, out.references))
    if out.libraries:
        have = p.components()
        p.set_components(have + [k for k in out.libraries if k not in have])
    p.dirty = True
    if main is not None and hasattr(main, "fill_tree"):
        main.fill_tree()
        main.update_title()
        if hasattr(main, "refresh_toolbox"):
            main.refresh_toolbox()
    return made


def open_wizard(main):
    """Add-Ins > VB 6 Data Object Wizard"""
    p = getattr(main, "project", None)
    if p is None:
        return None
    if hasattr(main, "_sync_editors"):
        main._sync_editors()
    dlg = DOWizard(main)
    main._dowizard = dlg                        # the tests reach it
    S = getattr(main, "S", None)
    if S is not None and S["dowiz_default_uc"]:
        kind = S["dowiz_default_uctype"]
        if kind in dlg.uc_radios:
            dlg.uc_radios[kind].setChecked(True)
    dlg.exec()
    return dlg


__all__ = ["DOWizard", "open_wizard", "add_to_project", "dowizard_icon", "do_picture"]
