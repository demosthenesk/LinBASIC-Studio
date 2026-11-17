"""The help window: the pages of the user's MSDN Library, its keyword index, Back/Forward and the
See Also / Example links of the pages."""
import os
import re

from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QDesktopServices, QKeySequence
from PyQt6.QtWidgets import (QMainWindow, QTextBrowser, QSplitter, QWidget, QVBoxLayout, QLineEdit, QListWidget,
                             QDialog, QDialogButtonBox, QLabel, QListWidgetItem, QToolBar, QMessageBox, QTabWidget,
                             QTreeWidget, QTreeWidgetItem, QComboBox, QPushButton, QHBoxLayout, QCheckBox, QMenu,
                             QToolButton, QProgressDialog, QApplication, QStyle)

from . import msdn


class TopicsFound(QDialog):
    """the Topics Found dialog of HTML Help: a keyword with more than one page"""

    def __init__(self, parent, topics):
        super().__init__(parent)
        self.setWindowTitle("Topics Found")
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Click a topic, then click Display."))
        self.list = QListWidget()
        for t in topics:
            it = QListWidgetItem(f"{t[0]}    ({t[1].name.upper()}.CHM)")
            it.setData(Qt.ItemDataRole.UserRole, t)
            self.list.addItem(it)
        self.list.setCurrentRow(0)
        self.list.itemDoubleClicked.connect(lambda _: self.accept())
        lay.addWidget(self.list)
        bb = QDialogButtonBox()
        bb.addButton("&Display", QDialogButtonBox.ButtonRole.AcceptRole)
        bb.addButton(QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.resize(460, 300)

    def chosen(self):
        it = self.list.currentItem()
        return it.data(Qt.ItemDataRole.UserRole) if it else None


class HelpWindow(QMainWindow):
    """the MSDN Library viewer of VB6 (DEV-8/me3): Hide, Locate, Previous, Next, Back, Forward, Stop, Refresh,
    Home, Print; the tabs Contents, Index and Search on the left, the page on the right"""
    TAB_CONTENTS, TAB_INDEX, TAB_SEARCH = 0, 1, 2

    def __init__(self, main, library):
        super().__init__(main)
        self.lib = library
        self.setWindowTitle("MSDN Library - Visual Basic 6")
        self.resize(1000, 700)
        self.back, self.fwd = [], []
        self.current = None                  # (ChmFile, local)
        tb = QToolBar()
        tb.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.addToolBar(tb)
        SP = QStyle.StandardPixmap
        st = self.style()

        def act(text, slot, icon, key=None):
            a = tb.addAction(st.standardIcon(icon), text)
            a.triggered.connect(lambda checked=False, f=slot: f())
            if key:
                a.setShortcut(QKeySequence(key))
            return a
        self.a_hide = act("Hide", self.toggle_side, SP.SP_TitleBarShadeButton)
        self.a_locate = act("Locate", self.locate, SP.SP_FileDialogContentsView)
        self.a_prev = act("Previous", lambda: self.step_toc(-1), SP.SP_ArrowUp)
        self.a_next = act("Next", lambda: self.step_toc(1), SP.SP_ArrowDown)
        self.a_back = act("Back", self.go_back, SP.SP_ArrowBack, "Alt+Left")
        self.a_fwd = act("Forward", self.go_forward, SP.SP_ArrowForward, "Alt+Right")
        self.a_stop = act("Stop", lambda: None, SP.SP_BrowserStop)
        self.a_refresh = act("Refresh", self.refresh_page, SP.SP_BrowserReload)
        self.a_home = act("Home", self.go_home, SP.SP_DirHomeIcon)
        self.a_print = act("Print", self.print_page, SP.SP_DialogSaveButton)
        split = QSplitter()
        self.split = split
        self.tabs = QTabWidget()
        # Contents
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.itemExpanded.connect(self._expand)
        self.tree.itemClicked.connect(self._toc_clicked)
        self.tree.itemActivated.connect(self._toc_clicked)
        self.tabs.addTab(self.tree, "&Contents")
        # Index
        side = QWidget()
        sl = QVBoxLayout(side)
        sl.setContentsMargins(2, 2, 2, 2)
        sl.addWidget(QLabel("Type in the keyword to find:"))
        self.find = QLineEdit()
        self.find.textEdited.connect(self._filter)
        self.find.returnPressed.connect(self._display_keyword)
        sl.addWidget(self.find)
        self.index = QListWidget()
        self.index.itemActivated.connect(lambda _: self._display_keyword())
        self.index.itemDoubleClicked.connect(lambda _: self._display_keyword())
        sl.addWidget(self.index)
        self.tabs.addTab(side, "I&ndex")
        # Search
        sw = QWidget()
        sv = QVBoxLayout(sw)
        sv.setContentsMargins(2, 2, 2, 2)
        sv.addWidget(QLabel("Type in the word(s) to search for:"))
        row = QHBoxLayout()
        self.words = QComboBox()
        self.words.setEditable(True)
        self.words.lineEdit().returnPressed.connect(self.list_topics)
        row.addWidget(self.words, 1)
        ops = QToolButton()
        ops.setArrowType(Qt.ArrowType.RightArrow)
        menu = QMenu(ops)
        for op in ("AND", "OR", "NEAR", "NOT"):
            menu.addAction(op, lambda o=op: self._add_operator(o))
        ops.setMenu(menu)
        ops.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        row.addWidget(ops)
        sv.addLayout(row)
        row = QHBoxLayout()
        self.b_list = QPushButton("&List Topics")
        self.b_list.clicked.connect(self.list_topics)
        self.b_display = QPushButton("&Display")
        self.b_display.clicked.connect(self._display_found)
        row.addWidget(self.b_list)
        row.addStretch(1)
        row.addWidget(self.b_display)
        sv.addLayout(row)
        row = QHBoxLayout()
        topic_label = QLabel("Select &topic:")
        row.addWidget(topic_label)
        self.found_label = QLabel("Found: 0")
        row.addStretch(1)
        row.addWidget(self.found_label)
        sv.addLayout(row)
        self.found = QTreeWidget()
        self.found.setRootIsDecorated(False)
        self.found.setHeaderLabels(["Title", "Location", "Rank"])
        self.found.itemDoubleClicked.connect(lambda *_: self._display_found())
        self.found.itemActivated.connect(lambda *_: self._display_found())
        topic_label.setBuddy(self.found)        # Alt+T: the list (the & is a mnemonic, not shown)
        sv.addWidget(self.found, 1)
        self.previous_results = QCheckBox("Search &previous results")
        self.similar = QCheckBox("&Match similar words")
        self.similar.setChecked(True)
        self.titles_only = QCheckBox("Search titles o&nly")
        for c in (self.previous_results, self.similar, self.titles_only):
            sv.addWidget(c)
        self.tabs.addTab(sw, "&Search")
        self.results = []                    # the topics of the last search
        split.addWidget(self.tabs)
        self.view = QTextBrowser()
        # the pages of MSDN are made for a white page (black text, blue links): light whatever the theme of the
        # IDE, or the text without a colour of its own (<pre> code) is light grey on white with a dark theme
        from .designer import light_palette
        self.view.setPalette(light_palette())
        self.view.viewport().setAutoFillBackground(True)
        self.view.setOpenLinks(False)
        self.view.anchorClicked.connect(self._link)
        split.addWidget(self.view)
        split.setSizes([280, 720])
        self.setCentralWidget(split)
        self._keywords = None
        self._toc_filled = False
        self._update_actions()

    # ------------------------------------------------------------ the toolbar
    def toggle_side(self):
        """Hide / Show: the tabs on the left"""
        vis = not self.tabs.isVisible()
        self.tabs.setVisible(vis)
        self.a_hide.setText("Hide" if vis else "Show")

    def refresh_page(self):
        if self.current is not None:
            self.show_page(*self.current, record=False)

    def go_home(self):
        node = self.lib.contents()[0] if self.lib.contents() else None
        page = self.lib.resolve(node) if node is not None else None
        if page:
            self.show_page(*page)

    def print_page(self):
        from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
        printer = QPrinter()
        if QPrintDialog(printer, self).exec():
            self.view.print(printer)

    # ------------------------------------------------------------ Contents
    def _fill_toc(self):
        if self._toc_filled:
            return
        self._toc_filled = True
        self.tree.clear()
        for node in self.lib.contents():
            self.tree.addTopLevelItem(self._item(node))

    def _item(self, node):
        it = QTreeWidgetItem([node.name])
        it.setData(0, Qt.ItemDataRole.UserRole, node)
        icon = QStyle.StandardPixmap.SP_DirIcon if node.children else QStyle.StandardPixmap.SP_FileIcon
        it.setIcon(0, self.style().standardIcon(icon))
        if node.children:
            it.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.ShowIndicator)
        return it

    def _expand(self, it):
        if it.childCount() == 0:             # the children when the book opens (the tree is large)
            node = it.data(0, Qt.ItemDataRole.UserRole)
            for ch in node.children:
                it.addChild(self._item(ch))

    def _toc_clicked(self, it, *_):
        node = it.data(0, Qt.ItemDataRole.UserRole)
        page = self.lib.resolve(node)
        if page is not None:
            self.show_page(*page)
        elif node.children:
            it.setExpanded(not it.isExpanded())

    def show_contents(self):
        self._fill_toc()
        self.tabs.setCurrentIndex(self.TAB_CONTENTS)
        if self.tree.topLevelItemCount():
            top = self.tree.topLevelItem(0)
            top.setExpanded(True)
            self.tree.setCurrentItem(top)
            if self.current is None:
                self._toc_clicked(top)
        self.show()
        self.raise_()
        self.activateWindow()

    def _toc_path(self, chm, local):
        """the nodes from the root to the page (for Locate)"""
        want = (chm.name, local.split("#")[0].lower().lstrip("/"))

        def walk(nodes, path):
            for n in nodes:
                page = self.lib.resolve(n)
                if page is not None and (page[0].name, page[1].split("#")[0].lower().lstrip("/")) == want:
                    return path + [n]
                got = walk(n.children, path + [n])
                if got:
                    return got
            return None
        return walk(self.lib.contents(), [])

    def locate(self):
        """Locate: the page in the Contents tab"""
        if self.current is None:
            return
        path = self._toc_path(*self.current)
        if not path:
            return
        self._fill_toc()
        self.tabs.setCurrentIndex(self.TAB_CONTENTS)
        parent = None
        for node in path:
            items = [parent.child(i) for i in range(parent.childCount())] if parent is not None else \
                [self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount())]
            parent = next((x for x in items if x.data(0, Qt.ItemDataRole.UserRole) is node), None)
            if parent is None:
                return
            if node is not path[-1]:
                parent.setExpanded(True)
        self.tree.setCurrentItem(parent)
        self.tree.scrollToItem(parent)

    def _toc_pages(self):
        if getattr(self, "_pages", None) is None:
            self._pages = []

            def walk(nodes):
                for n in nodes:
                    page = self.lib.resolve(n)
                    if page is not None:
                        self._pages.append(page)
                    walk(n.children)
            walk(self.lib.contents())
        return self._pages

    def step_toc(self, d):
        """Previous / Next: the page before / after this one in the Contents"""
        if self.current is None:
            return
        pages = self._toc_pages()
        key = (self.current[0].name, self.current[1].split("#")[0].lower())
        for i, (f, loc) in enumerate(pages):
            if (f.name, loc.split("#")[0].lower()) == key:
                j = i + d
                if 0 <= j < len(pages):
                    self.show_page(*pages[j])
                return

    # ------------------------------------------------------------ Search
    def _add_operator(self, op):
        le = self.words.lineEdit()
        le.setText(le.text().rstrip() + f" {op} ")
        le.setFocus()

    def show_search(self, words=""):
        self.tabs.setCurrentIndex(self.TAB_SEARCH)
        if words:
            self.words.setEditText(words)
        self.words.setFocus()
        self.show()
        self.raise_()
        self.activateWindow()

    def list_topics(self):
        """List Topics: the pages with the words (the first search makes the index of the pages, once)"""
        q = self.words.currentText().strip()
        if not q:
            return
        if self.words.findText(q) < 0:
            self.words.insertItem(0, q)
        dlg = None
        if getattr(self.lib, "_fts", None) is None:
            dlg = QProgressDialog("Indexing the MSDN Library (once)...", "Cancel", 0, len(self.lib.files), self)
            dlg.setWindowTitle("MSDN Library")
            dlg.setMinimumDuration(300)

        def progress(i, n, name):
            if dlg is None:
                return True
            dlg.setValue(i)
            dlg.setLabelText(f"Indexing the MSDN Library (once): {name.upper()}.CHM")
            QApplication.processEvents()
            return not dlg.wasCanceled()
        within = self.results if self.previous_results.isChecked() else None
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            self.results = self.lib.search(q, titles_only=self.titles_only.isChecked(),
                                           similar=self.similar.isChecked(), within=within, progress=progress)
        finally:
            QApplication.restoreOverrideCursor()
            if dlg is not None:
                dlg.close()
        self.found.clear()
        for title, f, local, rank in self.results:
            it = QTreeWidgetItem([title, f.title, str(rank)])
            it.setData(0, Qt.ItemDataRole.UserRole, (f, local))
            self.found.addTopLevelItem(it)
        self.found_label.setText(f"Found: {len(self.results)}")
        if self.results:
            self.found.setCurrentItem(self.found.topLevelItem(0))
        else:
            QMessageBox.information(self, "MSDN Library", "No topics found.")

    def _display_found(self):
        it = self.found.currentItem()
        if it is not None:
            self.show_page(*it.data(0, Qt.ItemDataRole.UserRole))

    def showEvent(self, e):
        super().showEvent(e)
        self._ensure_index()

    # ------------------------------------------------------------ the index
    def _ensure_index(self):
        if self._keywords is None:
            self._keywords = self.lib.all_keywords()
            self.index.addItems(self._keywords)

    def _filter(self, text):
        self._ensure_index()
        t = text.lower()
        for i, kw in enumerate(self._keywords):
            if kw.lower().startswith(t):
                self.index.setCurrentRow(i)
                self.index.scrollToItem(self.index.item(i), QListWidget.ScrollHint.PositionAtTop)
                return

    def _display_keyword(self):
        it = self.index.currentItem()
        if it is not None:
            self.show_topics(self.lib.by_keyword(it.text()) or [], exact=it.text())

    def show_index(self, word=""):
        self._ensure_index()
        self.tabs.setCurrentIndex(self.TAB_INDEX)
        self.find.setText(word)
        self._filter(word)
        self.show()
        self.raise_()
        self.activateWindow()

    # ------------------------------------------------------------ pages
    def show_topics(self, topics, exact=None):
        """one page: shown; more: the Topics Found dialog"""
        if exact:
            ex = [t for t in topics if t[0].lower() == exact.lower()]
            topics = ex or topics
        if not topics:
            QMessageBox.information(self, "MSDN Library", "The topic was not found in your MSDN Library.")
            return False
        if len(topics) > 1:
            dlg = TopicsFound(self, topics)
            if not dlg.exec():
                return False
            t = dlg.chosen()
        else:
            t = topics[0]
        return self.show_page(t[1], t[2])

    def show_page(self, chm, local, record=True):
        path = chm.page(local)
        if path is None:
            QMessageBox.information(self, "MSDN Library", f"{local} is not in {chm.name.upper()}.CHM.")
            return False
        if record and self.current is not None:
            self.back.append(self.current)
            self.fwd.clear()
        self.current = (chm, local)
        with open(path, "rb") as f:
            data = f.read()
        m = re.search(rb'charset=([A-Za-z0-9_-]+)', data[:600])
        enc = m.group(1).decode() if m else "windows-1252"
        if enc.lower().replace("_", "-") in ("iso8859-1", "iso-8859-1", "latin1", "latin-1", "us-ascii", "ascii"):
            enc = "windows-1252"                # as the browsers read it: the ’ “ ” of MSDN's pages are cp1252
        try:
            html = data.decode(enc, errors="replace")
        except LookupError:
            html = data.decode("cp1252", errors="replace")
        self.view.document().setBaseUrl(QUrl.fromLocalFile(path))
        self.view.setSearchPaths([os.path.dirname(path)])
        self.view.setHtml(msdn.rewrite(html))
        frag = local.split("#")[1] if "#" in local else None
        if frag:
            self.view.scrollToAnchor(frag)
        t = re.search(r"<title>(.*?)</title>", html, re.I | re.S)
        self.setWindowTitle((t.group(1).strip() + " - " if t else "") + "MSDN Library - Visual Basic 6")
        self._update_actions()
        self.show()
        self.raise_()
        self.activateWindow()
        return True

    def go_back(self):
        if self.back:
            self.fwd.append(self.current)
            self.show_page(*self.back.pop(), record=False)

    def go_forward(self):
        if self.fwd:
            self.back.append(self.current)
            self.show_page(*self.fwd.pop(), record=False)

    def _update_actions(self):
        self.a_back.setEnabled(bool(self.back))
        self.a_fwd.setEnabled(bool(self.fwd))

    def _link(self, url):
        s = url.toString()
        low = s.lower()
        chm = self.current[0] if self.current else None
        if low.startswith("alink:") or low.startswith("klink:"):
            names = [x for x in s[6:].split(";") if x]
            topics = self.lib.by_alink(names, chm) if low.startswith("alink:") else \
                [t for n in names for t in self.lib.by_keyword(n)]
            self.show_topics(topics)
            return
        m = re.match(r"^(?:ms-its:|mk:@msitstore:)(?:.*[\\/])?([^\\/]+?\.chm)::/?(.*)$", s, re.I)
        if m:
            f = self.lib.file(m.group(1))
            if f is not None:
                self.show_page(f, m.group(2))
            return
        if url.scheme() in ("http", "https", "mailto", "ftp"):
            QDesktopServices.openUrl(url)
            return
        # a page of the same help file
        if chm is not None:
            base = os.path.dirname(self.current[1].replace("\\", "/"))
            target = url.path() if url.isLocalFile() or url.scheme() == "file" else s
            if url.isLocalFile():
                root = os.path.realpath(chm.pages_dir)
                full = os.path.realpath(url.toLocalFile())
                if full.startswith(root):
                    local = os.path.relpath(full, root)
                    self.show_page(chm, local + (("#" + url.fragment()) if url.fragment() else ""))
                    return
            local = os.path.normpath(os.path.join(base, target)).replace("\\", "/")
            if url.fragment() and not target:
                self.view.scrollToAnchor(url.fragment())
                return
            self.show_page(chm, local + (("#" + url.fragment()) if url.fragment() else ""))
