"""The Property Page Wizard (phase 9.2e): Add-Ins > Property Page Wizard... when the "VB 6 Property Page Wizard"
add-in is loaded (VB6 oracle DEV-9/pp1, pp2), and Connect Property Pages (the PropertyPages property of a user
control).

The pages of VB6's: Introduction (Skip this screen in the future), Select a Control (a project with more than one
user control), Select the Property Pages (the pages with check boxes, Add / Rename: "Property Page Name", the arrows
for their order), Add Properties (Available Properties, a tab for each page, > >> < <<; the pages of the system
cannot be changed), Finished! (a Summary Report: Yes / No); then "Property Page Created" (Don't show this dialog in
the future). The code is ppwizgen's; the texts and pictures are our own.
"""
from PyQt6.QtCore import Qt, QRectF, QPointF
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPen, QIcon
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton, QLineEdit,
                             QListWidget, QListWidgetItem, QStackedWidget, QFrame, QCheckBox, QRadioButton,
                             QMessageBox, QTabWidget, QAbstractItemView, QButtonGroup)

from . import ppwizgen as G
from .appwizard import wizard_picture
from .axwizard import SummaryDialog

TITLE = "Property Page Wizard"
ADDIN = "VB 6 Property Page Wizard"
PAGE_TITLES = {"intro": "Introduction", "control": "Select a Control", "pages": "Select the Property Pages",
               "props": "Add Properties", "finished": "Finished!"}
NO_UC = "The project has no user control: add one with Project > Add User Control first, with its properties."


def pp_picture(kind, w=150, h=120):
    """the picture on the left of a page: our own drawings (property pages one behind the other)"""
    if kind == "finished":
        return wizard_picture("finished", w, h)
    pm = QPixmap(w, h)
    pm.fill(QColor("#808080"))
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    ink, paper, blue = QColor("#202020"), QColor("#FFFFFF"), QColor("#2F6FD0")

    def page(x, y, ww=70, hh=52, rows=3):
        p.setPen(QPen(ink, 1))
        p.setBrush(paper)
        p.drawRect(QRectF(x, y, ww, hh))
        p.setBrush(blue)
        p.drawRect(QRectF(x, y, ww, 7))
        p.setPen(QPen(QColor("#606060"), 1))
        for i in range(rows):
            yy = y + 14 + i * 12
            p.drawLine(QPointF(x + 6, yy), QPointF(x + 22, yy))
            p.setBrush(paper)
            p.drawRect(QRectF(x + 26, yy - 4, ww - 34, 7))
    n = 3 if kind in ("intro", "pages", "control") else 1
    for i in range(n):
        page(20 + i * 16, 18 + i * 14)
    if kind == "props":
        p.setPen(QPen(ink, 1))
        p.setBrush(QColor("#F2C230"))
        for i in range(3):
            p.drawRect(QRectF(104, 30 + i * 16, 22, 10))
        p.drawLine(QPointF(100, 50), QPointF(92, 50))
    p.end()
    return pm


def ppwizard_icon(size=16):
    """the icon of Add-Ins > Property Page Wizard..."""
    pm = pp_picture("pages").copy(14, 12, 104, 96).scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio,
                                                         Qt.TransformationMode.SmoothTransformation)
    return QIcon(pm)


class PageNameDialog(QDialog):
    """Add / Rename: "Property Page Name" (DEV-9/pp1-04: "Enter new page name:", OK / Cancel)"""

    def __init__(self, parent, name):
        super().__init__(parent)
        self.setWindowTitle("Property Page Name")
        lay = QGridLayout(self)
        lab = QLabel("&Enter new page name:")
        lay.addWidget(lab, 0, 0, Qt.AlignmentFlag.AlignTop)
        col = QVBoxLayout()
        ok = QPushButton("OK")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        col.addWidget(ok)
        col.addWidget(cancel)
        col.addStretch(1)
        lay.addLayout(col, 0, 1)
        self.name = QLineEdit(name)
        self.name.selectAll()
        lab.setBuddy(self.name)
        lay.addWidget(self.name, 1, 0, 1, 2)


class CreatedDialog(QDialog):
    """after Finish: "Property Page Created" (DEV-9/pp1-12)"""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Property Page Created")
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("The property pages have been made."))
        self.dont_show = QCheckBox("&Don't show this dialog in the future.")
        lay.addWidget(self.dont_show)
        ok = QPushButton("OK")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        lay.addWidget(ok, 0, Qt.AlignmentFlag.AlignHCenter)


class PPWizard(QDialog):
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
        self.summary_dialog = None
        self.created_dialog = None
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
                           ("pages", self._pages_page), ("props", self._props_page),
                           ("finished", self._finished_page)):
            w = QFrame()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(6, 0, 0, 0)
            build(lay)
            self.pages[key] = self.stack.addWidget(w)
        self.setFixedSize(640, 440)
        S = getattr(main, "S", None)
        skip = bool(S["ppwiz_skip_intro"]) if S is not None else False
        self.key = "intro"
        self.show_page(self.page_keys()[1] if skip else "intro")

    # ------------------------------------------------------------ pages
    @staticmethod
    def _text(lay, text):
        t = QLabel(text)
        t.setWordWrap(True)
        lay.addWidget(t)
        return t

    @staticmethod
    def _label(text, buddy):
        lab = QLabel(text)
        lab.setBuddy(buddy)
        return lab

    def _intro_page(self, lay):
        self._text(lay, "The Property Page Wizard makes the property pages of a user control: the pages its "
                        "properties are set in at design time (Shift+F4 on the control on a form).\n\n"
                        "You can go Back at any time to change what you chose. Click Next to begin.")
        lay.addStretch(1)
        self.cb_skip = QCheckBox("&Skip this screen in the future.")
        lay.addWidget(self.cb_skip)

    def _control_page(self, lay):
        self._text(lay, "Choose the user control whose property pages you want to make.")
        self.control_list = QListWidget()
        for m in self.ucs:
            self.control_list.addItem(f"{self.project.name}:{m.name}")
        self.control_list.setCurrentRow(0)
        lay.addWidget(self.control_list, 1)

    def _pages_page(self, lay):
        self._text(lay, "Choose the property pages of your user control: the checked ones are its pages. Add makes "
                        "a new page; the arrows set the order of the pages.")
        row = QHBoxLayout()
        col = QVBoxLayout()
        col.addStretch(1)
        self.b_add = QPushButton("&Add")
        self.b_rename = QPushButton("&Rename")
        self.b_add.clicked.connect(lambda: self.add_page())
        self.b_rename.clicked.connect(lambda: self.rename_page())
        col.addWidget(self.b_add)
        col.addWidget(self.b_rename)
        row.addLayout(col)
        self.page_list = QListWidget()
        self.page_list.itemChanged.connect(self._page_checked)
        self.page_list.currentRowChanged.connect(lambda *_: self._pages_states())
        row.addWidget(self.page_list, 1)
        arrows = QVBoxLayout()
        arrows.addStretch(1)
        self.b_up = QPushButton("↑")
        self.b_down = QPushButton("↓")
        for b, d in ((self.b_up, -1), (self.b_down, 1)):
            b.setFixedWidth(32)
            b.clicked.connect(lambda _=False, d=d: self.move_page(d))
            arrows.addWidget(b)
        arrows.addStretch(1)
        row.addLayout(arrows)
        lay.addLayout(row, 1)

    def _props_page(self, lay):
        self._text(lay, "Choose the properties each page shows: click the tab of a page, choose a property and a "
                        "move button.\nThe pages of the system ('Standard') cannot be changed.")
        grid = QGridLayout()
        self.avail = QListWidget()
        self.avail.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tabs = QTabWidget()
        grid.addWidget(self._label("A&vailable Properties:", self.avail), 0, 0)
        grid.addWidget(QLabel("Available Property Pages:"), 0, 2)
        grid.addWidget(self.avail, 1, 0)
        col = QVBoxLayout()
        col.addStretch(1)
        self.move_buttons = []
        for text, d, every in ((">", 1, False), (">>", 1, True), ("<", -1, False), ("<<", -1, True)):
            b = QPushButton(text)
            b.setFixedWidth(36)
            b.clicked.connect(lambda _=False, d=d, every=every: self.move_props(d, every))
            col.addWidget(b)
            self.move_buttons.append(b)
        col.addStretch(1)
        grid.addLayout(col, 1, 1)
        grid.addWidget(self.tabs, 1, 2)
        grid.setColumnStretch(0, 2)
        grid.setColumnStretch(2, 3)
        lay.addLayout(grid, 1)
        self.tabs.currentChanged.connect(lambda *_: self._props_states())
        self.avail.itemDoubleClicked.connect(lambda *_: self.move_props(1, False))

    def _finished_page(self, lay):
        self._text(lay, "The wizard knows all it needs to make the property pages.\n\nClick Finish to make them.")
        lay.addStretch(1)
        self._text(lay, "Do you want to see a Summary Report with what to do after the wizard?")
        self.rb_yes = QRadioButton("&Yes")
        self.rb_no = QRadioButton("N&o")
        self.rb_yes.setChecked(True)
        grp = QButtonGroup(self)
        grp.addButton(self.rb_yes)
        grp.addButton(self.rb_no)
        lay.addWidget(self.rb_yes)
        lay.addWidget(self.rb_no)

    # ------------------------------------------------------------ the order of the pages
    def page_keys(self):
        keys = ["intro"]
        if len(self.ucs) > 1:
            keys.append("control")
        return keys + ["pages", "props", "finished"]

    def show_page(self, key):
        self.key = key
        keys = self.page_keys()
        self.stack.setCurrentIndex(self.pages[key])
        self.setWindowTitle(f"{TITLE} - {PAGE_TITLES[key]}")
        self.picture.setPixmap(pp_picture(key))
        i = keys.index(key)
        self.b_back.setEnabled(i > 0)
        self.b_next.setEnabled(i < len(keys) - 1)
        self.b_finish.setEnabled(key in ("props", "finished"))      # VB6: from Add Properties on
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
            self.main.help_keyword("Property Page Wizard")

    def check_page(self, key):
        if self.uc is None:
            return NO_UC
        return None

    # ------------------------------------------------------------ the model
    def load_control(self, uc):
        """the pages of a user control (its PropertyPages) and its properties"""
        self.uc = uc
        data = uc.frx_value(uc.sf.root, "PropertyPages") if uc.sf is not None and uc.sf.root is not None else None
        if hasattr(data, "data"):
            data = data.data
        mods = {m.name.lower(): m for m in self.project.modules if m.kind == "PropertyPage"}
        self.model = G.PPModel(uc.name, G.uc_properties(uc.code_text()), G.read_pages(data), mods)
        return self.model

    def read_page(self, key):
        S = getattr(self.main, "S", None)
        if key == "intro" and S is not None and S["ppwiz_skip_intro"] != self.cb_skip.isChecked():
            S["ppwiz_skip_intro"] = self.cb_skip.isChecked()
            S.save()
        if key in ("intro", "control"):
            uc = self.ucs[max(0, self.control_list.currentRow())] if key == "control" and self.ucs else self.uc
            if uc is not None and (self.model is None or self.model.uc != uc.name):
                self.load_control(uc)

    def enter_page(self, key):
        if self.model is None and self.uc is not None and key not in ("intro", "control"):
            self.load_control(self.uc)
        if key == "intro":
            S = getattr(self.main, "S", None)
            self.cb_skip.setChecked(bool(S["ppwiz_skip_intro"]) if S is not None else False)
        elif key == "pages":
            self.fill_pages()
        elif key == "props":
            self.fill_props()

    # ------------------------------------------------------------ Select the Property Pages
    def fill_pages(self, current=None):
        self._filling = True
        self.page_list.clear()
        for pg in self.model.pages:
            it = QListWidgetItem(pg.name)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Checked if pg.checked else Qt.CheckState.Unchecked)
            self.page_list.addItem(it)
        self._filling = False
        names = self.model.page_names()
        if names:
            self.page_list.setCurrentRow(names.index(current) if current in names else 0)
        self._pages_states()

    def _page_checked(self, it):
        if getattr(self, "_filling", False):
            return
        pg = self.model.page(it.text())
        if pg is not None:
            pg.checked = it.checkState() == Qt.CheckState.Checked

    def _pages_states(self):
        it = self.page_list.currentItem()
        pg = self.model.page(it.text()) if it is not None and self.model is not None else None
        self.b_rename.setEnabled(pg is not None and not pg.standard and not pg.existing)
        row = self.page_list.currentRow()
        self.b_up.setEnabled(row > 0)
        self.b_down.setEnabled(0 <= row < self.page_list.count() - 1)

    def check_page_name(self, name, but=None):
        return self.model.check_name(name, but)

    def add_page(self, name=None):
        if name is None:
            d = PageNameDialog(self, self.model.new_name())
            self.last_name_dialog = d
            if self.quiet or not d.exec():
                return d
            name = d.name.text().strip()
        why = self.check_page_name(name)
        if why:
            self.message(why)
            return None
        pg = self.model.add_page(name)
        self.fill_pages(pg.name)
        return pg

    def rename_page(self, name=None):
        it = self.page_list.currentItem()
        if it is None:
            return None
        pg = self.model.page(it.text())
        if pg is None or pg.standard or pg.existing:
            return None
        if name is None:
            d = PageNameDialog(self, pg.name)
            self.last_name_dialog = d
            if self.quiet or not d.exec():
                return d
            name = d.name.text().strip()
        why = self.check_page_name(name, pg)
        if why:
            self.message(why)
            return None
        self.model.rename_page(pg.name, name)
        self.fill_pages(name)
        return pg

    def select_page(self, name):
        """the tests: a page of the list chosen"""
        names = self.model.page_names()
        self.page_list.setCurrentRow(names.index(name))

    def move_page(self, d):
        it = self.page_list.currentItem()
        if it is None:
            return
        name = it.text()
        self.model.move_page(name, d)
        self.fill_pages(name)

    def check(self, name, on=True):
        """the tests: the check box of a page"""
        names = self.model.page_names()
        self.page_list.item(names.index(name)).setCheckState(Qt.CheckState.Checked if on else Qt.CheckState.Unchecked)

    # ------------------------------------------------------------ Add Properties
    def fill_props(self, tab=None):
        cur = tab or (self.tabs.tabText(self.tabs.currentIndex()) if self.tabs.count() else None)
        self.avail.clear()
        self.avail.addItems(self.model.available())
        self.tabs.blockSignals(True)
        while self.tabs.count():
            w = self.tabs.widget(0)
            self.tabs.removeTab(0)
            w.deleteLater()
        self.tab_lists = {}
        for pg in self.model.selected_pages():
            lst = QListWidget()
            lst.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
            lst.addItems(self.model.page_props(pg.name))
            lst.itemDoubleClicked.connect(lambda *_: self.move_props(-1, False))
            self.tabs.addTab(lst, pg.name)
            self.tab_lists[pg.name] = lst
        names = [self.tabs.tabText(i) for i in range(self.tabs.count())]
        if names:
            self.tabs.setCurrentIndex(names.index(cur) if cur in names else 0)
        self.tabs.blockSignals(False)
        self._props_states()

    def current_tab(self):
        return self.tabs.tabText(self.tabs.currentIndex()) if self.tabs.count() else None

    def select_tab(self, name):
        names = [self.tabs.tabText(i) for i in range(self.tabs.count())]
        self.tabs.setCurrentIndex(names.index(name))

    def _props_states(self):
        pg = self.model.page(self.current_tab()) if self.model is not None and self.current_tab() else None
        on = pg is not None and not pg.standard
        for b in self.move_buttons:
            b.setEnabled(on)

    def move_props(self, d, every, names=None):
        tab = self.current_tab()
        pg = self.model.page(tab) if tab else None
        if pg is None or pg.standard:
            return
        lst = self.avail if d > 0 else self.tab_lists.get(tab)
        if names is None:
            items = [lst.item(i) for i in range(lst.count())] if every else lst.selectedItems()
            names = [it.text() for it in items]
        for n in names:
            if d > 0:
                self.model.add_prop(tab, n)
            else:
                self.model.remove_prop(tab, n)
        self.fill_props(tab)

    # ------------------------------------------------------------ Finish
    def finish(self):
        self.read_page(self.key)
        if self.uc is None:
            self.message(NO_UC)
            return None
        if self.model is None:
            self.load_control(self.uc)
        main = self.main
        if hasattr(main, "_sync_editors"):
            main._sync_editors()
        for pg in self.model.custom_pages():   # a page written again: its windows closed first
            if pg.existing and pg.module is not None and \
                    [p.lower() for p in pg.props] != [p.lower() for p in pg.old_props]:
                for wins in (getattr(main, "code_windows", {}), getattr(main, "design_windows", {})):
                    sw = wins.get(pg.module)
                    if sw is not None:
                        sw.close()

        def set_code(m, text):
            sw = getattr(main, "code_windows", {}).get(m)
            cw = sw.widget() if sw is not None else None
            if cw is not None and hasattr(main, "_set_code_text"):
                main._set_code_text(cw, text.rstrip("\n") + "\n")
                m.set_code(cw.text())
            else:
                m.editor_text = None
                m.sf.code = text.rstrip("\n").split("\n")

        def new_module(kind, name, ext, text):
            from .cbmodel import _new_module
            return _new_module(self.project, kind, name, ext, text)
        made = G.apply(self.model, self.project, self.uc, set_code, new_module)
        self.made = made
        d = getattr(main, "design_windows", {}).get(self.uc)
        if d is not None and hasattr(d.widget(), "build"):
            d.widget().build()
        for f in ("fill_tree", "update_title"):
            if hasattr(main, f):
                getattr(main, f)()
        if hasattr(main, "props") and hasattr(main.props, "refresh_values"):
            try:
                main.props.refresh_values()
            except Exception:                   # noqa: BLE001 - the Properties window of another object
                pass
        S = getattr(main, "S", None)
        if S is None or not S["ppwiz_no_created"]:
            cd = CreatedDialog(self)
            self.created_dialog = cd
            if not self.quiet:
                cd.exec()
            if S is not None and cd.dont_show.isChecked():
                S["ppwiz_no_created"] = True
                S.save()
        if self.rb_yes.isChecked():
            sd = SummaryDialog(self, G.summary(self.model))
            sd.setWindowTitle(TITLE + " Summary Report")
            self.summary_dialog = sd
            if not self.quiet:
                sd.exec()
        self.accept()
        return made


class ConnectPagesDialog(QDialog):
    """the PropertyPages property of a user control: "Connect Property Pages" (DEV-9/pp1-17): the pages of the
    system and of the project with check boxes, their order with the arrows"""

    def __init__(self, parent, chosen, project_pages):
        super().__init__(parent)
        self.setWindowTitle("Connect Property Pages")
        lay = QGridLayout(self)
        self.list = QListWidget()
        lay.addWidget(QLabel("&Available Property Pages:"), 0, 0)
        lay.addWidget(self.list, 1, 0, 3, 1)
        others = [n for n in project_pages if n.lower() not in {c.lower() for c in chosen}]
        std = [n for n in G.STANDARD + ["StandardDataFormat"] if n not in chosen]
        for n, on in [(n, True) for n in chosen] + [(n, False) for n in others + std]:
            it = QListWidgetItem(n)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Checked if on else Qt.CheckState.Unchecked)
            self.list.addItem(it)
        self.list.setCurrentRow(0)
        arrows = QVBoxLayout()
        self.b_up = QPushButton("↑")
        self.b_down = QPushButton("↓")
        self.b_up.clicked.connect(lambda: self.move(-1))
        self.b_down.clicked.connect(lambda: self.move(1))
        arrows.addStretch(1)
        arrows.addWidget(self.b_up)
        arrows.addWidget(QLabel("Page Order"))
        arrows.addWidget(self.b_down)
        arrows.addStretch(1)
        lay.addLayout(arrows, 1, 1, 3, 1)
        col = QVBoxLayout()
        ok = QPushButton("OK")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        col.addWidget(ok)
        col.addWidget(cancel)
        col.addStretch(1)
        lay.addLayout(col, 1, 2, 3, 1)

    def move(self, d):
        r = self.list.currentRow()
        j = r + d
        if r < 0 or not 0 <= j < self.list.count():
            return
        it = self.list.takeItem(r)
        self.list.insertItem(j, it)
        self.list.setCurrentRow(j)

    def chosen(self):
        return [self.list.item(i).text() for i in range(self.list.count())
                if self.list.item(i).checkState() == Qt.CheckState.Checked]


def open_wizard(main):
    """Add-Ins > Property Page Wizard..."""
    p = getattr(main, "project", None)
    if p is None:
        return None
    if not any(m.kind == "UserControl" for m in p.modules):
        QMessageBox.information(main, TITLE, NO_UC)
        return None
    if hasattr(main, "_sync_editors"):
        main._sync_editors()
    dlg = PPWizard(main)
    main._ppwizard = dlg
    dlg.exec()
    return dlg


__all__ = ["PPWizard", "open_wizard", "ppwizard_icon", "ADDIN", "PageNameDialog", "CreatedDialog",
           "ConnectPagesDialog"]
