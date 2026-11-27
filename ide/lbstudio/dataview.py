"""The Data View window (phase 12.3; View > Data View Window, VB6 oracles Claude Enviroment/oracles/12.3 dv*): its
toolbar (Refresh | Add a Data Link | Add a Data Environment), the Data Links of the user and the connections of the
DataEnvironments of the project; a connection -> Tables, Views, Stored Procedures -> the columns. Open (a double click)
shows the rows of a table ("Run Table: <name>"); Design opens a view in the Query Designer; a table is dragged on its
Diagram pane. The Data Links are kept in the settings of the IDE (VB6: in the registry of the user)."""
from PyQt6.QtCore import Qt, QMimeData, QSize
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QPen, QDrag
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QToolBar, QTreeWidget, QTreeWidgetItem, QMenu, QMessageBox,
                             QDialog, QFormLayout, QLabel, QDialogButtonBox, QAbstractItemView)

from . import dbinfo
from .querydesign import TABLE_MIME

ROLE = Qt.ItemDataRole.UserRole                 # (kind, data)
KIND_LINKS, KIND_DES, KIND_LINK, KIND_DECONN, KIND_FOLDER, KIND_TABLE, KIND_VIEW, KIND_PROC, KIND_FIELD = range(9)
FOLDERS = ["Tables", "Views", "Stored Procedures"]


def _icon(kind):
    """small pictures of our own for the tree and the toolbar"""
    pm = QPixmap(16, 16)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    pen = QPen(QColor("#000000"))
    p.setPen(pen)
    if kind == "folder":
        p.setBrush(QColor("#f0d060"))
        p.drawRect(1, 4, 13, 9)
        p.drawRect(1, 2, 5, 2)
    elif kind in ("table", "view"):
        p.setBrush(QColor("#ffffff"))
        p.drawRect(1, 2, 13, 11)
        p.fillRect(2, 3, 12, 3, QColor("#000080") if kind == "table" else QColor("#008080"))
        for y in (8, 11):
            p.drawLine(2, y, 13, y)
        p.drawLine(6, 6, 6, 12)
    elif kind == "proc":
        p.setBrush(QColor("#ffffff"))
        p.drawRect(2, 1, 11, 13)
        for y in (4, 7, 10):
            p.drawLine(4, y, 11, y)
    elif kind == "field":
        p.setBrush(QColor("#c0c0c0"))
        p.drawRect(3, 5, 10, 5)
    elif kind in ("link", "deconn"):
        p.setBrush(QColor("#c0c0c0"))
        p.drawRect(2, 2, 9, 7)
        p.drawLine(6, 9, 6, 12)
        p.drawLine(3, 13, 13, 13)
        if kind == "deconn":
            p.setBrush(QColor("#ffff00"))
            p.drawEllipse(9, 8, 6, 6)
    elif kind == "refresh":
        p.setPen(QPen(QColor("#008000"), 2))
        p.drawArc(3, 3, 10, 10, 30 * 16, 280 * 16)
        p.drawLine(12, 2, 12, 6)
        p.drawLine(12, 6, 8, 6)
    elif kind == "add_link":
        p.setBrush(QColor("#c0c0c0"))
        p.drawRect(1, 4, 8, 6)
        p.drawLine(5, 10, 5, 13)
        p.setPen(QPen(QColor("#000000"), 2))
        p.drawLine(12, 2, 12, 8)
        p.drawLine(9, 5, 15, 5)
    elif kind == "add_de":
        p.setBrush(QColor("#ffffff"))
        p.drawRect(1, 3, 10, 11)
        p.setBrush(QColor("#ffff00"))
        p.drawEllipse(4, 7, 5, 5)
        p.setPen(QPen(QColor("#000000"), 2))
        p.drawLine(13, 1, 13, 7)
        p.drawLine(10, 4, 16, 4)
    p.end()
    return QIcon(pm)


class DataTree(QTreeWidget):
    """the tree; a table or a view is dragged onto the Diagram pane of a Query Designer"""

    def __init__(self, view):
        super().__init__()
        self.view = view
        self.setHeaderHidden(True)
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
        self.setEditTriggers(QAbstractItemView.EditTrigger.EditKeyPressed)

    def startDrag(self, actions):
        it = self.currentItem()
        k = it.data(0, ROLE)[0] if it is not None and it.data(0, ROLE) else None
        if k not in (KIND_TABLE, KIND_VIEW):
            return
        md = QMimeData()
        md.setData(TABLE_MIME, it.text(0).encode("utf-8"))
        md.setText(it.text(0))
        drag = QDrag(self)
        drag.setMimeData(md)
        drag.exec(Qt.DropAction.CopyAction)


class DataView(QWidget):
    def __init__(self, main):
        super().__init__()
        self.main = main
        self.show_system = {}                   # the connections with Show System Objects: {key: True}
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        tb = QToolBar()
        tb.setIconSize(QSize(16, 16))
        self.a_refresh = tb.addAction(_icon("refresh"), "Refresh", self.refresh)
        self.a_add_link = tb.addAction(_icon("add_link"), "Add a Data Link", self.add_link)
        self.a_add_de = tb.addAction(_icon("add_de"), "Add a Data Environment", self.add_data_environment)
        v.addWidget(tb)
        self.toolbar = tb
        self.tree = DataTree(self)
        self.tree.itemExpanded.connect(self._expand)
        self.tree.itemDoubleClicked.connect(self._double)
        self.tree.itemChanged.connect(self._renamed)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.context_menu)
        v.addWidget(self.tree, 1)
        self.fill()

    # ---------------------------------------------------------------- the tree
    def links(self):
        return list(self.main.S.get("data_links", []) or [])

    def save_links(self, links):
        self.main.S["data_links"] = links
        self.main.S.save()

    def de_connections(self):
        """[(designer name, connection name, its connection string)] of the project's DataEnvironments"""
        from .dowizgen import data_environments
        out = []
        proj = getattr(self.main, "project", None)
        for m, de in data_environments(proj):
            sw = getattr(self.main, "design_windows", {}).get(m)
            if sw is not None and hasattr(sw.widget(), "model") and hasattr(sw.widget().model, "conns"):
                de = sw.widget().model
            for c in de.conns:
                out.append((m.name, c.name, c.source))
        return out

    def project_dir(self):
        import os
        p = getattr(getattr(self.main, "project", None), "path", None)
        return os.path.dirname(p) if p else os.getcwd()

    def fill(self):
        t = self.tree
        t.blockSignals(True)
        open_ = self._open_paths()
        t.clear()
        root = QTreeWidgetItem(["Data Links"])
        root.setIcon(0, _icon("folder"))
        root.setData(0, ROLE, (KIND_LINKS, None))
        t.addTopLevelItem(root)
        for ln in self.links():
            self._conn_item(root, ln["name"], ln["connect"], KIND_LINK)
        root.setExpanded(True)
        de = QTreeWidgetItem(["Data Environment Connections"])
        de.setIcon(0, _icon("folder"))
        de.setData(0, ROLE, (KIND_DES, None))
        t.addTopLevelItem(de)
        for dname, cname, src in self.de_connections():
            self._conn_item(de, cname, src, KIND_DECONN, dname)
        de.setExpanded(True)
        t.blockSignals(False)
        self._reopen(open_)

    def _conn_item(self, parent, name, connect, kind, designer=None):
        it = QTreeWidgetItem([name])
        it.setIcon(0, _icon("link" if kind == KIND_LINK else "deconn"))
        it.setData(0, ROLE, (kind, {"name": name, "connect": connect, "designer": designer}))
        if kind == KIND_LINK:
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsEditable)
        it.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.ShowIndicator)
        parent.addChild(it)
        return it

    def _open_paths(self):
        out = set()

        def walk(it, path):
            for i in range(it.childCount()):
                ch = it.child(i)
                p = path + (ch.text(0),)
                if ch.isExpanded():
                    out.add(p)
                walk(ch, p)
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            walk(top, (top.text(0),))
        return out

    def _reopen(self, paths):
        def walk(it, path):
            for i in range(it.childCount()):
                ch = it.child(i)
                p = path + (ch.text(0),)
                if p in paths:
                    ch.setExpanded(True)
                    walk(ch, p)
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            walk(top, (top.text(0),))

    def conn_of(self, it):
        """the connection item an item is under"""
        while it is not None:
            d = it.data(0, ROLE)
            if d and d[0] in (KIND_LINK, KIND_DECONN):
                return it
            it = it.parent()
        return None

    def connect_of(self, it):
        c = self.conn_of(it)
        if c is None:
            return ""
        from .dsrdesign import design_cs
        return design_cs(c.data(0, ROLE)[1]["connect"] or "", self.project_dir())

    def _key(self, conn_item):
        d = conn_item.data(0, ROLE)
        return (d[0], d[1].get("designer"), d[1]["name"])

    def _expand(self, it):
        d = it.data(0, ROLE)
        if not d or it.childCount():
            return
        kind = d[0]
        self.tree.blockSignals(True)
        try:
            if kind in (KIND_LINK, KIND_DECONN):
                for f in FOLDERS:
                    x = QTreeWidgetItem([f])
                    x.setIcon(0, _icon("folder"))
                    x.setData(0, ROLE, (KIND_FOLDER, f))
                    x.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.ShowIndicator)
                    it.addChild(x)
            elif kind == KIND_FOLDER:
                self._fill_folder(it)
            elif kind in (KIND_TABLE, KIND_VIEW):
                cs = self.connect_of(it)
                try:
                    flds = dbinfo.ado_fields(cs, 2, it.text(0))
                except dbinfo.AdoError as e:
                    flds = []
                    self.main.statusBar().showMessage(str(e), 5000)
                for f in flds:
                    x = QTreeWidgetItem([f[0]])
                    x.setIcon(0, _icon("field"))
                    x.setData(0, ROLE, (KIND_FIELD, f))
                    it.addChild(x)
        finally:
            self.tree.blockSignals(False)
        if not it.childCount():
            it.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.DontShowIndicatorWhenChildless)

    def _fill_folder(self, it):
        cs = self.connect_of(it)
        conn = self.conn_of(it)
        sysobj = self.show_system.get(self._key(conn), False)
        folder = it.data(0, ROLE)[1]
        try:
            objs = dbinfo.ado_objects(cs, sysobj)
        except dbinfo.AdoError as e:
            self.main.statusBar().showMessage(str(e), 5000)
            return
        if folder == "Tables":
            names = sorted([n for k, n in objs if k in ("T", "S")], key=str.lower)
            kind, icon = KIND_TABLE, "table"
        elif folder == "Views":
            names = sorted([n for k, n in objs if k == "Q"], key=str.lower)
            kind, icon = KIND_VIEW, "view"
        else:
            names = self.procedures(cs)
            kind, icon = KIND_PROC, "proc"
        for n in names:
            x = QTreeWidgetItem([n])
            x.setIcon(0, _icon(icon))
            x.setData(0, ROLE, (kind, n))
            if kind != KIND_PROC:
                x.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.ShowIndicator)
            it.addChild(x)

    def procedures(self, cs):
        """the stored procedures of a Jet database: its queries that are not views (with PARAMETERS, the action
        queries)"""
        try:
            _c, rows, _m = dbinfo.ado_query(cs, "SELECT name, sql, type FROM MSysLB_Queries", 0)
        except dbinfo.AdoError:
            return []
        out = []
        for name, sql, typ in rows:
            s = (sql or "").lstrip().upper()
            if str(typ) != "0" or s.startswith("PARAMETERS"):
                out.append(name)
        return sorted(out, key=str.lower)

    # ---------------------------------------------------------------- the commands
    def refresh(self):
        dbinfo.forget()
        self.fill()

    def add_link(self, answer=None):
        """Add a Data Link: Data Link Properties; the new link DataLinkN, its name ready to be changed (VB6)"""
        from .datalink import DataLinkDialog
        if answer is None:
            dlg = DataLinkDialog(self, "", self.project_dir())
            if not dlg.exec():
                return None
            answer = dlg.connect_string()
        links = self.links()
        n = 1
        names = {ln["name"].lower() for ln in links}
        while f"datalink{n}" in names:
            n += 1
        name = f"DataLink{n}"
        links.append({"name": name, "connect": answer})
        self.save_links(links)
        self.fill()
        it = self.find(KIND_LINK, name)
        if it is not None:
            self.tree.setCurrentItem(it)
            if not getattr(self, "quiet", False):
                self.tree.editItem(it, 0)
        return name

    def add_data_environment(self):
        """Add a Data Environment: a new DataEnvironment designer in the project (Project > Add Data Environment)"""
        if getattr(self.main, "project", None) is not None:
            self.main.add_module("DataEnvironment")
        self.fill()

    def find(self, kind, name):
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            for j in range(top.childCount()):
                ch = top.child(j)
                d = ch.data(0, ROLE)
                if d and d[0] == kind and ch.text(0).lower() == name.lower():
                    return ch
        return None

    def _renamed(self, it, col):
        d = it.data(0, ROLE)
        if not d or d[0] != KIND_LINK:
            return
        old, new = d[1]["name"], it.text(0).strip()
        if new == old:
            return
        links = self.links()
        if not new or any(ln["name"].lower() == new.lower() and ln["name"] != old for ln in links):
            self.tree.blockSignals(True)
            it.setText(0, old)
            self.tree.blockSignals(False)
            return
        for ln in links:
            if ln["name"] == old:
                ln["name"] = new
        self.save_links(links)
        d[1]["name"] = new
        it.setData(0, ROLE, d)

    def rename(self, it):
        self.tree.editItem(it, 0)

    def remove(self, it):
        d = it.data(0, ROLE)
        if d[0] == KIND_LINK:
            self.save_links([ln for ln in self.links() if ln["name"] != d[1]["name"]])
        elif d[0] == KIND_DECONN:
            sw = self._de_window(d[1]["designer"])
            if sw is not None:
                model = sw.widget().model
                if any(c.conn.lower() == d[1]["name"].lower() for c in model.cmds):
                    QMessageBox.warning(self, "Data View", f"The commands of {d[1]['name']} use it.")
                    return
                model.conns = [c for c in model.conns if c.name.lower() != d[1]["name"].lower()]
                sw.widget().changed()
                sw.widget().fill_tree()
        self.fill()

    def _de_window(self, designer):
        proj = getattr(self.main, "project", None)
        m = next((x for x in (proj.modules if proj else []) if x.name.lower() == (designer or "").lower()), None)
        if m is None:
            return None
        self.main.open_designer(m)
        return self.main.design_windows.get(m)

    def modify(self, it):
        """Modify...: Data Link Properties of the connection"""
        from .datalink import DataLinkDialog
        d = it.data(0, ROLE)
        dlg = DataLinkDialog(self, d[1]["connect"], self.project_dir())
        if not dlg.exec():
            return
        cs = dlg.connect_string()
        if d[0] == KIND_LINK:
            links = self.links()
            for ln in links:
                if ln["name"] == d[1]["name"]:
                    ln["connect"] = cs
            self.save_links(links)
        else:
            sw = self._de_window(d[1]["designer"])
            if sw is not None:
                for c in sw.widget().model.conns:
                    if c.name.lower() == d[1]["name"].lower():
                        c.source = cs
                sw.widget().changed()
                sw.widget().fill_tree()
        dbinfo.forget()
        self.fill()

    def toggle_system(self, it):
        conn = self.conn_of(it)
        k = self._key(conn)
        self.show_system[k] = not self.show_system.get(k, False)
        for i in range(conn.childCount()):      # the folders are read again
            f = conn.child(i)
            f.takeChildren()
            if f.isExpanded():
                self._expand(f)

    def properties(self, it):
        d = it.data(0, ROLE)
        dlg = QDialog(self)
        f = QFormLayout(dlg)
        if d[0] in (KIND_LINK, KIND_DECONN):
            dlg.setWindowTitle(f"{d[1]['name']} Properties")
            f.addRow("Name:", QLabel(d[1]["name"]))
            if d[0] == KIND_DECONN:
                f.addRow("Data Environment:", QLabel(d[1]["designer"]))
            lab = QLabel(d[1]["connect"])
            lab.setWordWrap(True)
            f.addRow("Connection String:", lab)
        elif d[0] in (KIND_TABLE, KIND_VIEW):
            dlg.setWindowTitle(f"{it.text(0)} Properties")
            f.addRow("Name:", QLabel(it.text(0)))
            f.addRow("Type:", QLabel("Table" if d[0] == KIND_TABLE else "View"))
            try:
                n = len(dbinfo.ado_fields(self.connect_of(it) or "", 2, it.text(0)))
            except dbinfo.AdoError:
                n = 0
            f.addRow("Columns:", QLabel(str(n)))
        elif d[0] == KIND_FIELD:
            from .querydesign import ADO_TYPE_NAMES
            fl = d[1]
            dlg.setWindowTitle(f"{fl[0]} Properties")
            f.addRow("Name:", QLabel(fl[0]))
            f.addRow("Data Type:", QLabel(ADO_TYPE_NAMES.get(fl[1], str(fl[1]))))
            f.addRow("Length:", QLabel(str(fl[2])))
            f.addRow("Precision:", QLabel(str(fl[3])))
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        bb.rejected.connect(dlg.reject)
        f.addRow(bb)
        self.last_dialog = dlg
        if not getattr(self, "quiet", False):
            dlg.exec()
        return dlg

    def open_item(self, it):
        """Open: the rows of a table or a view ("Run Table: <name>")"""
        from .querydesign import run_table
        d = it.data(0, ROLE)
        if d[0] not in (KIND_TABLE, KIND_VIEW):
            return None
        conn = self.conn_of(it)
        w = run_table(self.main, self.connect_of(it), it.text(0), conn.text(0))
        w.run()
        return self.main.add_query_window(w)

    def design(self, it):
        """Design of a view: its SQL in the Query Designer"""
        from .querydesign import QueryDesigner
        cs = self.connect_of(it)
        try:
            _c, rows, _m = dbinfo.ado_query(cs, "SELECT sql FROM MSysLB_Queries WHERE name = '%s'"
                                            % it.text(0).replace("'", "''"), 0)
            sql = (rows[0][0] or "").replace("\r\n", "\n").strip().rstrip(";").strip() if rows else ""
        except dbinfo.AdoError:
            sql = ""
        w = QueryDesigner(self.main, cs, sql or f"SELECT *\nFROM {it.text(0)}", title=f"Design: {it.text(0)}")
        return self.main.add_query_window(w)

    def _double(self, it, col):
        d = it.data(0, ROLE)
        if d and d[0] in (KIND_TABLE, KIND_VIEW):
            self.open_item(it)

    def menu_items(self, it):
        """the context menu of an item (VB6 dv*): [(text, slot, enabled[, icon[, checked]]) or None]"""
        d = it.data(0, ROLE) if it is not None else None
        kind = d[0] if d else KIND_LINKS
        R = (lambda: self.refresh())
        if kind in (KIND_LINKS, KIND_DES):
            return [("&Refresh", R, True)], [("Add a Data &Link...", lambda: self.add_link(), True)]
        if kind in (KIND_LINK, KIND_DECONN):
            sysobj = self.show_system.get(self._key(it), False)
            return ([("Re&name", lambda: self.rename(it), kind == KIND_LINK), ("Re&move", lambda: self.remove(it), True),
                     ("&Modify...", lambda: self.modify(it), True), None, ("&Refresh", R, True)],
                    [("Add a Data &Link...", lambda: self.add_link(), True), None,
                     ("Show S&ystem Objects", lambda: self.toggle_system(it), True, None, sysobj), None,
                     ("P&roperties...", lambda: self.properties(it), True)])
        if kind == KIND_FOLDER:
            sysobj = self.show_system.get(self._key(self.conn_of(it)), False)
            return [("&Refresh", R, True)], [("Show S&ystem Objects", lambda: self.toggle_system(it), True, None,
                                              sysobj)]
        if kind in (KIND_TABLE, KIND_VIEW):
            first = [("&Open", lambda: self.open_item(it), True)]
            if kind == KIND_VIEW:
                first.append(("&Design", lambda: self.design(it), True))
            return first + [None, ("&Refresh", R, True)], [("P&roperties...", lambda: self.properties(it), True)]
        if kind == KIND_FIELD:
            return [("&Refresh", R, True)], [("P&roperties...", lambda: self.properties(it), True)]
        return [("&Refresh", R, True)], []

    def context_menu(self, pos):
        """the right button: the item's own commands, Dockable and Hide (as the other tool windows), then the rest"""
        it = self.tree.itemAt(pos)
        if it is not None:
            self.tree.setCurrentItem(it)
        before, after = self.menu_items(it)
        m = QMenu(self)

        def add(items):
            for e in items:
                if e is None:
                    m.addSeparator()
                    continue
                a = m.addAction(e[0])
                a.triggered.connect(lambda _c=False, fn=e[1]: fn())
                a.setEnabled(bool(e[2]))
                if len(e) > 4 and e[4] is not None:
                    a.setCheckable(True)
                    a.setChecked(bool(e[4]))
        add(before)
        m.addSeparator()
        main = self.main
        a = m.addAction("Doc&kable", lambda: main.set_dockable("dataview", not main.is_dockable("dataview")))
        a.setCheckable(True)
        a.setChecked(main.is_dockable("dataview"))
        m.addAction("&Hide", lambda: main.dock_dataview.hide())
        if after:
            m.addSeparator()
            add(after)
        self.last_menu = m
        main.shortcut_menu("Data View", m)
        if not getattr(self, "quiet", False):
            m.exec(self.tree.viewport().mapToGlobal(pos))
        return m
