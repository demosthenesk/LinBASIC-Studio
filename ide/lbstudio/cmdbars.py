"""View > Toolbars > Customize... as VB6 has it (phase 12.2): the command bars of the IDE - its toolbars, its menu bar
and menus, the Shortcut Menus - changed by the user and kept in the settings (S["cmdbars"]).

VB6 oracles (Claude Enviroment/oracles/12.2): the dialog with Toolbars (Menu Bar, Standard, Edit, Debug, Form Editor,
Shortcut Menus and the user's own; New..., Rename..., Delete, Reset...), Commands (the categories of the commands, the
commands with their pictures, Description, Modify Selection) and Options (Large icons, Show ScreenTips on toolbars, Show
shortcut keys in ScreenTips, Menu animations); while it is open the bars are customized: a command is dragged from the
list onto a bar or into a menu, a button or a menu item is dragged to another place (with Ctrl: copied) or off the bars
(deleted), the right button (or Modify Selection) gives Reset, Delete, Name, the button image commands, the styles and
Begin a Group. data/commandbars.json: the lists of VB6.

The model: every command is a QAction of the IDE with an id ("File/Open Project...", the menu and its text without the
& and the keys; the names that change, "Save Form1", are kept as VB6 names them, "Save <File Name>"); a commmand of VB6
that LinBASIC Studio does not have is a grey placeholder. A bar is a list of items {"id", "g" (Begin a Group), "s"
(the style: 0 Default, 1 Text Only (Always), 2 Text Only (in Menus), 3 Image and Text), "n" (its name), "i" (its
image, a PNG in base64)}; "menu:<path>" is a menu (a popup), "w:<key>" a widget of a bar (the position of the
Standard toolbar, the lists of the Form Editor toolbar). Only what the user changed is kept.
"""
import base64
import copy
import json
import os
import random
import re

from PyQt6.QtCore import Qt, QSize, QPoint, QRect, QEvent, QObject, QMimeData, QByteArray, QBuffer, QIODevice, QTimer
from PyQt6.QtGui import QAction, QActionGroup, QIcon, QPixmap, QPainter, QColor, QDrag, QImage, QPen, QCursor
from PyQt6.QtWidgets import (QApplication, QDialog, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
                             QLabel, QPushButton, QListWidget, QListWidgetItem, QCheckBox, QComboBox, QLineEdit,
                             QToolBar, QToolButton, QMenu, QMenuBar, QMessageBox, QWidgetAction, QDialogButtonBox,
                             QGroupBox, QToolTip, QFrame, QAbstractItemView)

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = json.load(open(os.path.join(HERE, "data", "commandbars.json"), encoding="utf-8"))
MIME = "application/x-lbstudio-command"
TITLE = "Microsoft Visual Basic"                # (the title of VB6's messages; ours: LinBASIC Studio)
APP = "LinBASIC Studio"

# the names VB6 gives the commands whose text changes (the attribute of the action of MainWindow)
DYNAMIC = {"a_save_file": "File/Save <File Name>", "a_save_file_as": "File/Save <File Name> As...",
           "a_make": "File/Make <Project>...", "a_remove": "Project/Remove <Item>",
           "a_proj_props": "Project/<Project Name> Properties...", "a_undo": "Edit/Undo", "a_redo": "Edit/Redo",
           "a_break": "Run/Break"}
# a command of VB6 that has an other name in LinBASIC Studio
SAME = {"Help/About Microsoft Visual Basic...": "About LinBASIC Studio...", "Edit/Remove": None}
# the pictures of the commands that have none in the menus (edicons / tbicons)
PICTURES = {"File/New Project": ("tb", "New"), "File/Open Project...": ("tb", "Open"),
            "File/Save Project Group": ("tb", "Save"), "File/Save Project Group As...": ("tb", "Save"),
            "File/Print...": ("ed", "print"), "Edit/Undo": ("tb", "Undo"), "Edit/Redo": ("tb", "Redo"),
            "Edit/Cut": ("ed", "cut"), "Edit/Copy": ("ed", "copy"), "Edit/Paste": ("ed", "paste"),
            "Edit/Delete": ("tb", "Delete"), "Edit/Find...": ("tb", "Find"), "Edit/Find Next": ("tb", "Find"),
            "Edit/Comment Block": ("ed", "comment"), "Edit/Uncomment Block": ("ed", "uncomment"),
            "View/Code": ("ed", "view_code"), "View/Object": ("ed", "view_object"),
            "View/Project Explorer": ("ed", "project"), "View/Properties Window": ("ed", "properties"),
            "View/Property Pages": ("ed", "proppage"), "Project/Add Form": ("ed", "form"),
            "Project/Add MDI Form": ("ed", "mdiform"), "Project/Add Module": ("ed", "module"),
            "Project/Add Class Module": ("ed", "class"), "Project/Add User Control": ("ed", "usercontrol"),
            "Project/Add Property Page": ("ed", "proppage"), "Project/Add User Document": ("ed", "userdoc"),
            "Project/References...": ("ed", "references"), "Tools/Menu Editor...": ("ed", "menu_editor"),
            "Help/Contents...": ("ed", "help"), "Format/to Grid": ("ed", "align_grid"),
            "Format/Size to Grid": ("ed", "align_grid"), "View/Form Layout Window": ("ed", "form_layout"),
            "View/Toolbox": ("ed", "toolbox")}
STYLES = ["Default Style", "Text Only (Always)", "Text Only (in Menus)", "Image and Text"]


def norm(text):
    """the name of a command: its text without the & of its key, the keys after a tab and the spaces"""
    t = (text or "").split("\t")[0]
    t = t.replace("&&", "\0").replace("&", "").replace("\0", "&")
    return t.strip()


def png64(icon, size=16):
    if icon is None or icon.isNull():
        return None
    pm = icon.pixmap(size, size)
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QIODevice.OpenModeFlag.WriteOnly)
    pm.toImage().save(buf, "PNG")
    buf.close()
    return base64.b64encode(bytes(ba)).decode()


def icon64(s):
    if not s:
        return QIcon()
    img = QImage.fromData(base64.b64decode(s), "PNG")
    return QIcon(QPixmap.fromImage(img))


def picture(spec):
    if spec is None:
        return QIcon()
    if spec[0] == "ed":
        from . import edicons
        return edicons.icon(spec[1])
    from . import tbicons
    return tbicons.icon(spec[1])


# ================================================================ the commands
class Registry:
    """the commands of the IDE by their ids, and the categories of the Commands tab"""

    def __init__(self, mw):
        self.mw = mw
        self.actions = {}                       # id -> QAction
        self.ids = {}                           # id(QAction) -> id
        self.menus = {}                         # "menu:File", "menu:Format/Align" -> QMenu
        self.menu_paths = {}                    # id(QMenu) -> "menu:..."
        self.placeholders = set()
        self.widgets = {}                       # "w:pos" -> the QWidgetAction of a bar
        self.cats = {}
        dyn = {}
        for attr, cid in DYNAMIC.items():
            a = getattr(mw, attr, None)
            if a is not None:
                dyn[id(a)] = cid
        self.dyn = dyn
        self.top = []                           # the menus of the menu bar, in their order
        for a in mw.menuBar().actions():
            m = a.menu()
            if m is not None:
                t = norm(a.text())
                self.top.append(t)
                self._walk(m, t)
        self._categories()

    def add(self, cid, a):
        if cid in self.actions and self.actions[cid] is not a:
            n = 2
            while f"{cid} ({n})" in self.actions:
                n += 1
            cid = f"{cid} ({n})"
        self.actions.setdefault(cid, a)
        self.ids.setdefault(id(a), cid)
        return self.ids[id(a)]

    def _walk(self, menu, path):
        mid = "menu:" + path
        self.menus[mid] = menu
        self.menu_paths[id(menu)] = mid
        for a in menu.actions():
            if a.isSeparator():
                continue
            if a.menu() is not None:
                self._walk(a.menu(), path + "/" + norm(a.text()))
                continue
            top = path.split("/")[0]
            cid = self.dyn.get(id(a)) or f"{top}/{norm(a.text())}"
            if path.count("/") and top == "Format" and path.split("/")[1] in ("Horizontal Spacing", "Vertical Spacing"):
                cid += "|" + path.split("/")[1][0]
            self.add(cid, a)

    def find(self, cat, name):
        """the action of a command of VB6 (its category and name)"""
        key = name.split("|")[0]
        same = SAME.get(f"{cat}/{name}", key)
        if same is None:
            return None
        for cid, a in self.actions.items():
            top, _, rest = cid.partition("/")
            if top != cat:
                continue
            base = rest.split("|")[0]
            if base == same and ("|" not in name or rest.endswith("|" + name.split("|")[1])):
                return cid
        return None

    def _categories(self):
        mw = self.mw
        bars = [tb for tb in mw.findChildren(QToolBar)]
        loose = {}                              # the actions of the toolbars only (Comment Block ...)
        for tb in bars:
            for a in tb.actions():
                if not a.isSeparator() and id(a) not in self.ids and norm(a.text()):
                    loose.setdefault(norm(a.text()), a)
        for cat in DATA["categories"]:
            if cat == "Built-in Menus":
                self.cats[cat] = ["menu:" + t for t in self.top]
                continue
            lst = []
            for name in DATA["commands"].get(cat, []):
                cid = self.find(cat, name)
                if cid is None and name.split("|")[0] in loose:
                    cid = self.add(f"{cat}/{name}", loose.pop(name.split("|")[0]))
                if cid is None:                 # a command of VB6 LinBASIC Studio does not have: grey
                    a = QAction(name.split("|")[0], mw)
                    a.setEnabled(False)
                    a.setObjectName("cmdbars-placeholder")
                    cid = self.add(f"{cat}/{name}", a)
                    self.placeholders.add(cid)
                lst.append(cid)
            seen = set(lst)                     # then the commands of LinBASIC Studio in that menu
            for cid in self.actions:
                if cid.split("/")[0] == cat and cid not in seen:
                    lst.append(cid)
                    seen.add(cid)
            self.cats[cat] = lst
        for cid, spec in PICTURES.items():
            a = self.actions.get(cid)
            if a is not None and a.icon().isNull():
                a.setIcon(picture(spec))

    def action(self, cid):
        return self.actions.get(cid)

    def id_of(self, a):
        return self.ids.get(id(a))

    def name(self, cid):
        """the name of a command as the Commands list shows it"""
        if cid.startswith("menu:"):
            return cid[5:].split("/")[-1]
        a = self.actions.get(cid)
        rest = cid.partition("/")[2].split("|")[0]
        if cid in self.dyn.values() or a is None:
            return re.sub(r" \(\d+\)$", "", rest)
        return norm(a.text()) or rest


# ================================================================ the model
def default_options():
    return {"large": False, "tips": True, "keys": False, "anim": 0}


def item(cid, group=False, style=0, name=None, icon=None):
    d = {"id": cid}
    if group:
        d["g"] = True
    if style:
        d["s"] = style
    if name is not None:
        d["n"] = name
    if icon:
        d["i"] = icon
    return d


class CommandBars(QObject):
    """the bars of a MainWindow: their defaults, the user's changes (S["cmdbars"]), the Customize dialog"""
    BUILTIN = ["Menu Bar", "Standard", "Edit", "Debug", "Form Editor", "Shortcut Menus"]

    def __init__(self, mw):
        super().__init__(mw)
        self.mw = mw
        self.reg = Registry(mw)
        self.bars = {"Standard": mw.tb_standard, "Edit": mw.tb_edit, "Debug": mw.tb_debug,
                     "Form Editor": mw.tb_form_editor}
        self.defaults = {}
        self.default_vis = {}
        for name, tb in self.bars.items():
            self.defaults[name] = self.snapshot(tb, name)
        self.defaults["Menu Bar"] = [item("menu:" + t) for t in self.reg.top]
        self.menu_defaults = {}
        for mid, m in self.reg.menus.items():
            self.menu_defaults[mid] = self.menu_snapshot(m)
        self.shortcut_defaults = {}
        for group, menus in DATA["shortcut_menus"].items():
            for name, lst in menus.items():
                self.shortcut_defaults[name] = self.shortcut_snapshot(lst)
        self.proxies = []
        self.custom = {}                        # the user's own toolbars: name -> QToolBar
        self.customizing = None                 # the Customize dialog while it is open
        self.selected = None                    # (bar or menu, index) of the item chosen in customize mode
        self.clip_icon = None                   # Copy Button Image
        self.shortcut_bar = None
        self.filter = CustomizeFilter(self)
        self.tip_filter = TipFilter(self)
        mw.menuBar().installEventFilter(self.tip_filter)
        for m in self.reg.menus.values():
            m.aboutToShow.connect(lambda m=m: self._menu_shown(m))
        self.apply()

    # ---------------------------------------------------------------- the settings
    @property
    def S(self):
        return self.mw.S

    def state(self):
        st = self.S.get("cmdbars") or {}
        st = copy.deepcopy(st) if isinstance(st, dict) else {}
        st.setdefault("bars", {})
        st.setdefault("menus", {})
        st.setdefault("shortcut", {})
        st.setdefault("custom", [])
        o = default_options()
        o.update(st.get("options") or {})
        st["options"] = o
        return st

    def save(self, st):
        self.S["cmdbars"] = st
        self.S.save()

    def options(self):
        return self.state()["options"]

    def set_option(self, key, value):
        st = self.state()
        st["options"][key] = value
        self.save(st)
        self.apply_options()

    # ---------------------------------------------------------------- snapshots of the defaults
    def snapshot(self, tb, name):
        items, group = [], False
        for a in tb.actions():
            if a.isSeparator():
                group = True
                continue
            w = tb.widgetForAction(a)
            cid = self.reg.id_of(a)
            if isinstance(a, QWidgetAction) or (w is not None and not isinstance(w, QToolButton)) or \
                    (cid is None and w is not None and isinstance(w, QToolButton) and w.menu() is not None
                     and w.defaultAction() is None):
                key = f"w:{name}:{len(self.reg.widgets)}"
                self.reg.widgets[key] = a
                items.append(item(key, group))
            elif cid is not None:
                items.append(item(cid, group))
            else:
                cid = self.reg.add(f"Toolbar/{name}/{norm(a.text())}", a)
                items.append(item(cid, group))
            group = False
        return items

    def menu_snapshot(self, m):
        items, group = [], False
        for a in m.actions():
            if a.isSeparator():
                group = True
                continue
            if a.menu() is not None:
                mid = self.reg.menu_paths.get(id(a.menu()))
                if mid is None:
                    continue
                items.append(item(mid, group))
            else:
                cid = self.reg.id_of(a)
                if cid is None:
                    continue
                items.append(item(cid, group))
            group = False
        return items

    def shortcut_snapshot(self, lst):
        items, group = [], False
        for t in lst:
            if t == "|":
                group = True
                continue
            items.append(item("text:" + t, group))
            group = False
        return items

    # ---------------------------------------------------------------- the items of a bar
    def items(self, name):
        st = self.state()
        if name in st["bars"]:
            return st["bars"][name]["items"]
        if name in st["custom"]:
            return []
        return copy.deepcopy(self.defaults.get(name, []))

    def set_items(self, name, items):
        st = self.state()
        b = st["bars"].setdefault(name, {})
        b["items"] = items
        self.save(st)
        self.rebuild(name)

    def menu_items(self, mid):
        st = self.state()
        if mid in st["menus"]:
            return st["menus"][mid]
        if mid in st["shortcut"]:
            return st["shortcut"][mid]
        if mid.startswith("short:"):
            return copy.deepcopy(self.shortcut_defaults.get(mid[6:], []))
        return copy.deepcopy(self.menu_defaults.get(mid, []))

    def set_menu_items(self, mid, items):
        st = self.state()
        (st["shortcut"] if mid.startswith("short:") else st["menus"])[mid] = items
        self.save(st)
        m = self.popup(mid)
        if m is not None and not mid.startswith("short:"):
            self._fill(m, mid)

    def popup(self, mid):
        if mid.startswith("short:"):
            return self.short_menus.get(mid) if hasattr(self, "short_menus") else None
        return self.reg.menus.get(mid) or self.user_menus().get(mid)

    def user_menus(self):
        if not hasattr(self, "_user_menus"):
            self._user_menus = {}
        return self._user_menus

    def bar_names(self):
        return self.BUILTIN + list(self.state()["custom"])

    def toolbar(self, name):
        if name == "Shortcut Menus":
            return self.shortcut_bar
        return self.bars.get(name) or self.custom.get(name)

    # ---------------------------------------------------------------- applying the model
    def apply(self):
        st = self.state()
        for name in st["custom"]:
            if name not in self.custom:
                self.make_custom(name)
        for name in list(self.custom):
            if name not in st["custom"]:
                tb = self.custom.pop(name)
                self.mw.removeToolBar(tb)
                tb.deleteLater()
        for name in self.bars:
            if name in st["bars"]:
                self.rebuild(name)
        for name in self.custom:
            self.rebuild(name)
        if "Menu Bar" in st["bars"]:
            self.rebuild("Menu Bar")
        self.apply_options()

    def make_custom(self, name):
        tb = QToolBar(name, self.mw)
        tb.setObjectName("Custom:" + name)
        tb.setIconSize(QSize(16, 16))
        self.mw.addToolBar(tb)
        self.custom[name] = tb
        tb.installEventFilter(self.tip_filter)
        return tb

    def resolve(self, it):
        """(QAction or QMenu, its kind) of an item"""
        cid = it["id"]
        if cid.startswith("w:"):
            return self.reg.widgets.get(cid), "widget"
        if cid.startswith("menu:") or cid.startswith("user:"):
            return self.popup(cid), "menu"
        if cid.startswith("text:"):
            return None, "text"
        return self.reg.action(cid), "action"

    def proxy(self, a, it, parent, in_menu):
        """the action an item shows: the command itself, or a copy with the item's name, picture or style"""
        style = it.get("s", 0)
        name, ic = it.get("n"), it.get("i")
        no_icon = style == 1 or (in_menu and style == 2 and False)
        if name is None and not ic and not no_icon:
            return a
        p = QAction(parent)
        p.setObjectName("cmdbars-proxy")

        def sync():
            p.setText(name if name is not None else a.text())
            p.setIcon(QIcon() if no_icon else (icon64(ic) if ic else a.icon()))
            p.setEnabled(a.isEnabled())
            p.setVisible(a.isVisible())
            p.setCheckable(a.isCheckable())
            p.setChecked(a.isChecked())
            p.setToolTip(norm(name) if name is not None else a.toolTip())
        sync()
        a.changed.connect(sync)
        p.triggered.connect(lambda _c=False: a.trigger())
        p.target = a
        self.proxies.append(p)
        return p

    def rebuild(self, name):
        if name == "Menu Bar":
            self.rebuild_menubar()
            return
        tb = self.toolbar(name)
        if tb is None:
            return
        for a in list(tb.actions()):
            tb.removeAction(a)
            if a.objectName() == "cmdbars-proxy" or getattr(a, "cmdbars_sep", False):
                a.deleteLater()
        tb.cmdbars_items = []
        for it in self.items(name):
            a, kind = self.resolve(it)
            if a is None:
                continue
            if it.get("g") and tb.actions():
                s = tb.addSeparator()
                s.cmdbars_sep = True
            if kind == "widget":
                tb.addAction(a)
                obj = a
            elif kind == "menu":
                b = QToolButton(tb)
                b.setText(it.get("n") or norm(a.title()))
                b.setMenu(a)
                b.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
                b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
                obj = tb.addWidget(b)
                obj.setObjectName("cmdbars-proxy")
            else:
                obj = self.proxy(a, it, tb, False)
                tb.addAction(obj)
                w = tb.widgetForAction(obj)
                if isinstance(w, QToolButton):
                    st = it.get("s", 0)
                    w.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly if st == 1 or obj.icon().isNull()
                                         else Qt.ToolButtonStyle.ToolButtonTextBesideIcon if st == 3
                                         else Qt.ToolButtonStyle.ToolButtonIconOnly)
            tb.cmdbars_items.append((obj, it))
        if self.customizing is not None:        # the new buttons are customized too
            for b in tb.findChildren(QToolButton):
                b.installEventFilter(self.filter)
        self.apply_options()

    def rebuild_menubar(self):
        mb = self.mw.menuBar()
        for a in list(mb.actions()):
            mb.removeAction(a)
        for it in self.items("Menu Bar"):
            m, kind = self.resolve(it)
            if kind == "menu" and m is not None:
                if it.get("n"):
                    m.setTitle(it["n"])
                mb.addMenu(m)
            elif kind == "action" and m is not None:
                mb.addAction(self.proxy(m, it, mb, True))

    def _menu_shown(self, m):
        mid = self.reg.menu_paths.get(id(m))
        if mid is not None and mid in self.state()["menus"]:
            self._fill(m, mid)

    def _fill(self, m, mid):
        """a menu the user changed: its items in their order (the ones the IDE added since, at the end)"""
        items = self.menu_items(mid)
        known = set()
        for it in items:
            a, kind = self.resolve(it)
            if a is not None:
                known.add(id(a.menuAction() if kind == "menu" else a))
        extra = [a for a in m.actions() if not a.isSeparator() and id(a) not in known and
                 a.objectName() != "cmdbars-proxy" and self.reg.id_of(a) is None and
                 (a.menu() is None or id(a.menu()) not in self.reg.menu_paths)]
        for a in list(m.actions()):
            m.removeAction(a)
        for it in items:
            a, kind = self.resolve(it)
            if a is None:
                continue
            if it.get("g") and m.actions():
                m.addSeparator()
            if kind == "menu":
                m.addMenu(a)
            else:
                m.addAction(self.proxy(a, it, m, True))
        if extra:
            m.addSeparator()
            for a in extra:
                m.addAction(a)

    # ---------------------------------------------------------------- the options
    def apply_options(self):
        o = self.options()
        size = QSize(32, 32) if o.get("large") else QSize(16, 16)
        for tb in list(self.bars.values()) + list(self.custom.values()):
            tb.setIconSize(size)
            tb.installEventFilter(self.tip_filter)
        app = QApplication.instance()
        anim = int(o.get("anim", 0) or 0)
        app.setEffectEnabled(Qt.UIEffect.UI_AnimateMenu, anim in (1, 2, 3))
        app.setEffectEnabled(Qt.UIEffect.UI_FadeMenu, False)

    def tip_text(self, a):
        o = self.options()
        if not o.get("tips", True):
            return None
        t = norm(a.toolTip() or a.text())
        if o.get("keys") and not a.shortcut().isEmpty():
            t += f" ({a.shortcut().toString()})"
        return t

    # ---------------------------------------------------------------- the menu of the right button on the bars
    def popup_menu(self):
        """VB6: the toolbars with a check when shown, then Customize... (the Menu Bar and the Shortcut Menus are not
        in it)"""
        m = QMenu(self.mw)
        for name in ["Debug", "Edit", "Form Editor", "Standard"] + list(self.state()["custom"]):
            tb = self.toolbar(name)
            if tb is None:
                continue
            a = m.addAction(name)
            a.setCheckable(True)
            a.setChecked(tb.isVisible())
            a.toggled.connect(lambda on, t=tb: t.setVisible(on))
        m.addSeparator()
        m.addAction("&Customize...", self.customize)
        return m

    # ---------------------------------------------------------------- the toolbars of the user
    def new_toolbar(self, name):
        st = self.state()
        st["custom"].append(name)
        st["bars"][name] = {"items": []}
        self.save(st)
        tb = self.make_custom(name)
        tb.show()
        return tb

    def rename_toolbar(self, old, new):
        st = self.state()
        st["custom"] = [new if n == old else n for n in st["custom"]]
        st["bars"][new] = st["bars"].pop(old, {"items": []})
        self.save(st)
        tb = self.custom.pop(old)
        tb.setWindowTitle(new)
        tb.setObjectName("Custom:" + new)
        self.custom[new] = tb

    def delete_toolbar(self, name):
        st = self.state()
        st["custom"] = [n for n in st["custom"] if n != name]
        st["bars"].pop(name, None)
        self.save(st)
        tb = self.custom.pop(name, None)
        if tb is not None:
            self.mw.removeToolBar(tb)
            tb.deleteLater()

    def reset_toolbar(self, name):
        st = self.state()
        st["bars"].pop(name, None)
        if name == "Menu Bar":
            st["menus"] = {}
        if name == "Shortcut Menus":
            st["shortcut"] = {}
        self.save(st)
        if name == "Menu Bar":
            for mid, m in self.reg.menus.items():
                self._restore_menu(m, mid)
        if name == "Shortcut Menus":
            self.fill_shortcut_bar()
        self.rebuild(name)

    def _restore_menu(self, m, mid):
        for a in list(m.actions()):
            if a.objectName() == "cmdbars-proxy" or a.isSeparator():
                m.removeAction(a)
        for a in list(m.actions()):
            m.removeAction(a)
        for it in self.menu_defaults.get(mid, []):
            a, kind = self.resolve(it)
            if a is None:
                continue
            if it.get("g") and m.actions():
                m.addSeparator()
            m.addMenu(a) if kind == "menu" else m.addAction(a)

    def unique_name(self, base="Custom"):
        names = set(self.bar_names())
        n = 1
        while f"{base} {n}" in names:
            n += 1
        return f"{base} {n}"

    # ---------------------------------------------------------------- the Shortcut Menus toolbar
    def make_shortcut_bar(self):
        if self.shortcut_bar is None:
            tb = QToolBar("Shortcut Menus", self.mw)
            tb.setObjectName("ShortcutMenus")
            self.mw.addToolBar(tb)
            tb.hide()
            self.shortcut_bar = tb
            self.short_menus = {}
            self.fill_shortcut_bar()
        return self.shortcut_bar

    def fill_shortcut_bar(self):
        tb = self.shortcut_bar
        if tb is None:
            return
        for a in list(tb.actions()):
            tb.removeAction(a)
            a.deleteLater()
        for group, menus in DATA["shortcut_menus"].items():
            top = QMenu(group, tb)
            for name in menus:
                sub = top.addMenu(name)
                mid = "short:" + name
                self.short_menus[mid] = sub
                self.fill_short(sub, mid)
            b = QToolButton(tb)
            b.setText(group)
            b.setMenu(top)
            b.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
            tb.addWidget(b)

    def fill_short(self, m, mid):
        m.clear()
        for it in self.menu_items(mid):
            if it.get("g") and m.actions():
                m.addSeparator()
            cid = it["id"]
            if cid.startswith("text:"):
                t = cid[5:]
                a = m.addAction(t.rstrip(" >"))
                a.cmdbars_text = t
                if t.endswith(" >"):
                    a.setMenu(QMenu(m))
            else:
                a, kind = self.resolve(it)
                if a is None:
                    continue
                m.addMenu(a) if kind == "menu" else m.addAction(self.proxy(a, it, m, True))

    def shortcut(self, name, menu):
        """a context menu of the IDE (named as the Shortcut Menus toolbar names it), as the user changed it: its
        items found by their names, the commands added, the ones taken away left out"""
        mid = "short:" + name
        st = self.state()
        if mid not in st["shortcut"]:
            return menu
        have = {}
        for a in menu.actions():
            if not a.isSeparator():
                have.setdefault(norm(a.text()).lower().replace("continue", "start"), a)
        items = st["shortcut"][mid]
        for a in list(menu.actions()):
            menu.removeAction(a)
        for it in items:
            cid = it["id"]
            if cid.startswith("text:"):
                a = have.get(norm(cid[5:]).rstrip(" >").lower())
            else:
                a, kind = self.resolve(it)
                if kind == "menu" and a is not None:
                    a = a.menuAction()
                elif a is not None:
                    a = self.proxy(a, it, menu, True)
            if a is None:
                continue
            if it.get("g") and menu.actions():
                menu.addSeparator()
            if it.get("n"):
                a.setText(it["n"])
            menu.addAction(a)
        return menu

    # ---------------------------------------------------------------- the Customize dialog
    def customize(self):
        if self.customizing is not None:
            self.customizing.raise_()
            return self.customizing
        dlg = CustomizeDialog(self)
        self.customizing = dlg
        self.begin_customize()
        dlg.finished.connect(lambda _r: self.end_customize())
        dlg.show()
        return dlg

    def all_targets(self):
        """the widgets the user edits in customize mode"""
        out = [self.mw.menuBar()] + list(self.bars.values()) + list(self.custom.values())
        if self.shortcut_bar is not None:
            out.append(self.shortcut_bar)
        return out

    def begin_customize(self):
        self.make_shortcut_bar()
        for w in self.all_targets():
            w.installEventFilter(self.filter)
            w.setAcceptDrops(True)
            for b in w.findChildren(QToolButton):
                b.installEventFilter(self.filter)
        for m in list(self.reg.menus.values()) + list(self.short_menus.values()) + list(self.user_menus().values()):
            m.installEventFilter(self.filter)
            m.setAcceptDrops(True)

    def end_customize(self):
        for w in self.all_targets():
            w.removeEventFilter(self.filter)
            for b in w.findChildren(QToolButton):
                b.removeEventFilter(self.filter)
        for m in list(self.reg.menus.values()) + list(self.short_menus.values()) + list(self.user_menus().values()):
            m.removeEventFilter(self.filter)
        if self.shortcut_bar is not None:
            self.shortcut_bar.hide()            # VB6: the Shortcut Menus toolbar only while customizing
        self.select(None)
        self.customizing = None
        self.filter.hide_marker()

    # ---------------------------------------------------------------- the items (customize mode)
    def where(self, w):
        """(the name of the bar or the id of the menu, its kind) of a widget being customized"""
        if w is self.mw.menuBar():
            return "Menu Bar", "bar"
        for name in self.bar_names():
            if self.toolbar(name) is w:
                return name, "bar"
        if isinstance(w, QMenu):
            mid = self.reg.menu_paths.get(id(w))
            if mid is None and hasattr(self, "short_menus"):
                mid = next((k for k, v in self.short_menus.items() if v is w), None)
            if mid is None:
                mid = next((k for k, v in self.user_menus().items() if v is w), None)
            if mid is not None:
                return mid, "menu"
        return None, None

    def get_items(self, key, kind):
        return self.items(key) if kind == "bar" else self.menu_items(key)

    def put_items(self, key, kind, items):
        if kind == "bar":
            self.set_items(key, items)
        else:
            self.set_menu_items(key, items)
            if key.startswith("short:"):
                self.fill_short(self.short_menus[key], key)

    def insert(self, key, kind, index, it):
        items = self.get_items(key, kind)
        index = max(0, min(index, len(items)))
        items.insert(index, it)
        self.put_items(key, kind, items)
        return index

    def remove(self, key, kind, index):
        items = self.get_items(key, kind)
        if 0 <= index < len(items):
            it = items.pop(index)
            if index < len(items) and it.get("g") and not items[index].get("g"):
                items[index]["g"] = True if index > 0 else items[index].get("g", False)
            self.put_items(key, kind, items)
            return it
        return None

    def move(self, src, dst, copy_it=False):
        """src, dst: (key, kind, index); an item dragged from one place to an other (Ctrl: a copy)"""
        items = self.get_items(src[0], src[1])
        if not 0 <= src[2] < len(items):
            return
        it = copy.deepcopy(items[src[2]])
        if it["id"].startswith("w:") and copy_it:
            copy_it = False                     # a widget is in one place only
        same = (src[0], src[1]) == (dst[0], dst[1])
        if same and not copy_it:
            idx = dst[2]
            items.pop(src[2])
            if idx > src[2]:
                idx -= 1
            items.insert(max(0, min(idx, len(items))), it)
            self.put_items(src[0], src[1], items)
            return
        if not copy_it:
            self.remove(*src)
        self.insert(dst[0], dst[1], dst[2], it)

    def modify(self, key, kind, index, **change):
        items = self.get_items(key, kind)
        if not 0 <= index < len(items):
            return
        it = items[index]
        for k, v in change.items():
            if v is None or v is False or v == 0 and k == "s":
                it.pop(k, None)
            else:
                it[k] = v
        self.put_items(key, kind, items)

    def reset_item(self, key, kind, index):
        """Reset: the name, picture and style of the command"""
        self.modify(key, kind, index, n=None, i=None, s=None)

    def select(self, sel):
        self.selected = sel
        if self.customizing is not None:
            self.customizing.selection_changed()
        self.filter.show_selection()

    def item_at(self, key, kind, index):
        items = self.get_items(key, kind)
        return items[index] if 0 <= index < len(items) else None

    def item_name(self, it):
        if it is None:
            return ""
        if it.get("n") is not None:
            return it["n"]
        cid = it["id"]
        if cid.startswith("text:"):
            return cid[5:].rstrip(" >")
        a, kind = self.resolve(it)
        if kind == "menu" and a is not None:
            return a.title()
        if kind == "widget":
            return ""
        return a.text() if a is not None else self.reg.name(cid)

    def item_icon(self, it):
        if it is None:
            return QIcon()
        if it.get("i"):
            return icon64(it["i"])
        a, kind = self.resolve(it)
        return a.icon() if kind == "action" and a is not None else QIcon()


# ================================================================ customize mode: the mouse on the bars and menus
class CustomizeFilter(QObject):
    """while the Customize dialog is open: a click chooses an item (it does not run it), a drag moves it (Ctrl:
    copies it), off the bars deletes it; a command of the Commands list is dropped onto a bar or into a menu; the
    right button gives the Modify Selection menu"""

    def __init__(self, cb):
        super().__init__(cb)
        self.cb = cb
        self.press = None
        self.marker = None
        self.frame = None

    # ---- where an item is
    def items_of(self, w):
        """[(QRect in w, index)] of the items of a bar or a menu"""
        cb = self.cb
        key, kind = cb.where(w)
        out = []
        if kind == "bar" and isinstance(w, QToolBar):
            for i, (obj, it) in enumerate(self.bar_entries(w, key)):
                ww = w.widgetForAction(obj)
                if ww is not None and ww.isVisible():
                    out.append((ww.geometry(), i))
        elif kind == "bar":                     # the menu bar
            acts = [a for a in w.actions()]
            for i, a in enumerate(acts):
                r = w.actionGeometry(a)
                if r.isValid():
                    out.append((r, i))
        elif kind == "menu":
            i = 0
            for a in w.actions():
                if a.isSeparator():
                    continue
                out.append((w.actionGeometry(a), i))
                i += 1
        return out

    def bar_entries(self, tb, key):
        ent = getattr(tb, "cmdbars_items", None)
        if ent is None:                         # a bar not rebuilt yet: its default items
            ent = []
            items = self.cb.items(key)
            acts = [a for a in tb.actions() if not a.isSeparator()]
            for a, it in zip(acts, items):
                ent.append((a, it))
        return ent

    def index_at(self, w, pos):
        for r, i in self.items_of(w):
            if r.contains(pos):
                return i
        return -1

    def drop_index(self, w, pos):
        """where a dropped item goes: before the item under the mouse (after it, past its middle)"""
        ents = self.items_of(w)
        if not ents:
            return 0, QRect(0, 0, 2, w.height())
        vertical = isinstance(w, QMenu) or (isinstance(w, QToolBar) and w.orientation() == Qt.Orientation.Vertical)
        for r, i in ents:
            if (r.contains(pos)) or (not vertical and pos.x() < r.left()) or (vertical and pos.y() < r.top()):
                after = (pos.y() > r.center().y()) if vertical else (pos.x() > r.center().x())
                if r.contains(pos) and after:
                    return i + 1, (QRect(r.left(), r.bottom(), r.width(), 2) if vertical
                                   else QRect(r.right(), r.top(), 2, r.height()))
                return i, (QRect(r.left(), r.top() - 1, r.width(), 2) if vertical
                           else QRect(r.left() - 1, r.top(), 2, r.height()))
        r, i = ents[-1]
        return i + 1, (QRect(r.left(), r.bottom(), r.width(), 2) if vertical else QRect(r.right(), r.top(), 2, r.height()))

    # ---- the marks
    def show_marker(self, w, rect):
        if self.marker is None or self.marker.parent() is not w:
            self.hide_marker()
            self.marker = QFrame(w)
            self.marker.setStyleSheet("background: black")
        self.marker.setGeometry(rect)
        self.marker.show()
        self.marker.raise_()

    def hide_marker(self):
        if self.marker is not None:
            self.marker.hide()
            self.marker.deleteLater()
            self.marker = None

    def show_selection(self):
        if self.frame is not None:
            self.frame.hide()
            self.frame.deleteLater()
            self.frame = None
        sel = self.cb.selected
        if sel is None:
            return
        w, index = sel
        for r, i in self.items_of(w):
            if i == index:
                self.frame = QFrame(w)
                self.frame.setStyleSheet("border: 2px solid black; background: transparent")
                self.frame.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
                self.frame.setGeometry(r)
                self.frame.show()
                break

    # ---- the events
    def target(self, obj):
        """the bar or menu of a widget (a button of a toolbar: its toolbar)"""
        w = obj
        while w is not None and not isinstance(w, (QToolBar, QMenuBar, QMenu)):
            w = w.parentWidget()
        return w

    def eventFilter(self, obj, ev):
        t = ev.type()
        if t not in (QEvent.Type.MouseButtonPress, QEvent.Type.MouseButtonRelease, QEvent.Type.MouseMove,
                     QEvent.Type.MouseButtonDblClick, QEvent.Type.DragEnter, QEvent.Type.DragMove,
                     QEvent.Type.DragLeave, QEvent.Type.Drop, QEvent.Type.ContextMenu):
            return False
        w = self.target(obj)
        if w is None:
            return False
        key, kind = self.cb.where(w)
        if key is None:
            return False
        if t in (QEvent.Type.DragEnter, QEvent.Type.DragMove, QEvent.Type.DragLeave, QEvent.Type.Drop):
            return self.drag_event(w, key, kind, obj, ev)
        if t == QEvent.Type.ContextMenu:
            i = self.index_at(w, ev.pos() if obj is w else obj.mapTo(w, ev.pos()))
            if i >= 0:
                self.cb.select((w, i))
                self.modify_menu(w, key, kind, i).exec(ev.globalPos())
            return True
        pos = ev.position().toPoint()
        if obj is not w:
            pos = obj.mapTo(w, pos)
        if t == QEvent.Type.MouseButtonPress:
            i = self.index_at(w, pos)
            if ev.button() == Qt.MouseButton.RightButton:     # (the menu: on the ContextMenu event after it)
                if i >= 0:
                    self.cb.select((w, i))
                return True
            if i < 0:
                return False
            self.cb.select((w, i))
            self.press = (w, key, kind, i, ev.globalPosition().toPoint())
            it = self.cb.item_at(key, kind, i)
            if kind == "bar" and isinstance(w, QMenuBar):   # a menu of the menu bar opens: its items are customized
                return False
            if it is not None and (it["id"].startswith("menu:") or it["id"].startswith("user:")) and \
                    isinstance(w, QToolBar):
                return False
            if isinstance(w, QMenu):
                a = [x for x in w.actions() if not x.isSeparator()][i] if i < len(
                    [x for x in w.actions() if not x.isSeparator()]) else None
                if a is not None and a.menu() is not None:
                    w.setActiveAction(a)
            return True
        if t == QEvent.Type.MouseMove and self.press is not None:
            if (ev.globalPosition().toPoint() - self.press[4]).manhattanLength() >= \
                    QApplication.startDragDistance():
                w0, key0, kind0, i0, _ = self.press
                self.press = None
                self.start_drag(w0, key0, kind0, i0)
            return True
        if t in (QEvent.Type.MouseButtonRelease, QEvent.Type.MouseButtonDblClick):
            self.press = None
            return not isinstance(w, QMenuBar)
        return False

    def start_drag(self, w, key, kind, i):
        it = self.cb.item_at(key, kind, i)
        if it is None:
            return
        md = QMimeData()
        md.setData(MIME, json.dumps({"item": it, "src": [key, kind, i]}).encode())
        drag = QDrag(w)
        drag.setMimeData(md)
        ic = self.cb.item_icon(it)
        if not ic.isNull():
            drag.setPixmap(ic.pixmap(16, 16))
        res = drag.exec(Qt.DropAction.MoveAction | Qt.DropAction.CopyAction, Qt.DropAction.MoveAction)
        self.close_panel()
        if res == Qt.DropAction.IgnoreAction:   # dropped off the bars: deleted (VB6)
            self.cb.remove(key, kind, i)
            self.cb.select(None)
        self.hide_marker()

    def drag_event(self, w, key, kind, obj, ev):
        t = ev.type()
        if t == QEvent.Type.DragLeave:          # (it has no data)
            self.hide_marker()
            return False
        if not ev.mimeData().hasFormat(MIME):
            return False
        pos = ev.position().toPoint()
        if obj is not w:
            pos = obj.mapTo(w, pos)
        idx, rect = self.drop_index(w, pos)
        if t in (QEvent.Type.DragEnter, QEvent.Type.DragMove):
            ctrl = bool(ev.modifiers() & Qt.KeyboardModifier.ControlModifier)
            ev.setDropAction(Qt.DropAction.CopyAction if ctrl else Qt.DropAction.MoveAction)
            ev.accept()
            self.show_marker(w, rect)
            if isinstance(w, QMenuBar):          # over a menu of the menu bar: it opens, to drop into it (a
                a = w.actionAt(pos)              # popup menu cannot open while the mouse drags: its items in a
                if a is not None and a.menu() is not None:      # panel of their own, as VB6 shows them)
                    mid = self.cb.reg.menu_paths.get(id(a.menu())) or \
                        next((k for k, v in self.cb.user_menus().items() if v is a.menu()), None)
                    if mid is not None:
                        self.open_panel(mid, w.mapToGlobal(w.actionGeometry(a).bottomLeft()))
            return True
        # the drop
        self.hide_marker()
        data = json.loads(bytes(ev.mimeData().data(MIME)).decode())
        it = data["item"]
        src = data.get("src")
        ctrl = bool(ev.modifiers() & Qt.KeyboardModifier.ControlModifier)
        if src is None:                         # from the Commands list: a new item
            self.cb.insert(key, kind, idx, it)
            ev.setDropAction(Qt.DropAction.CopyAction)
        else:
            self.cb.move(tuple(src), (key, kind, idx), copy_it=ctrl)
            ev.setDropAction(Qt.DropAction.CopyAction if ctrl else Qt.DropAction.MoveAction)
            ev.accept()
            self._moved = True
            self.cb.select(None)
            return True
        ev.accept()
        self.cb.select(None)
        return True

    # ---- a menu while dragging
    def open_panel(self, mid, gpos):
        p = getattr(self, "panel", None)
        if p is not None and p.mid == mid and p.isVisible():
            return p
        if p is not None:
            p.hide()
            p.deleteLater()
        self.panel = MenuPanel(self.cb, mid)
        self.panel.move(gpos)
        self.panel.show()
        return self.panel

    def close_panel(self):
        p = getattr(self, "panel", None)
        if p is not None:
            p.hide()
            p.deleteLater()
            self.panel = None

    # ---- Modify Selection
    def modify_menu(self, w, key, kind, i):
        return modify_menu(self.cb, w, key, kind, i)


def modify_menu(cb, w, key, kind, i, parent=None):
    """Reset, Delete | Name: | Copy / Paste / Reset / Edit / Change Button Image | the four styles | Begin a Group"""
    it = cb.item_at(key, kind, i) or {}
    m = QMenu(parent or cb.mw)
    m.addAction("&Reset", lambda: (cb.reset_item(key, kind, i), cb.select((w, i))))
    m.addAction("&Delete", lambda: (cb.remove(key, kind, i), cb.select(None)))
    m.addSeparator()
    box = QWidget()
    hl = QHBoxLayout(box)
    hl.setContentsMargins(22, 2, 6, 2)
    lab = QLabel("&Name:")
    ed = QLineEdit(cb.item_name(it))
    lab.setBuddy(ed)
    hl.addWidget(lab)
    hl.addWidget(ed)
    wa = QWidgetAction(m)
    wa.setDefaultWidget(box)
    m.addAction(wa)

    def named():
        cb.modify(key, kind, i, n=ed.text())
        m.close()
    ed.returnPressed.connect(named)
    m.cmdbars_name = ed
    m.addSeparator()
    a_copy = m.addAction("&Copy Button Image", lambda: setattr(cb, "clip_icon", png64(cb.item_icon(it))))
    a_paste = m.addAction("&Paste Button Image", lambda: cb.modify(key, kind, i, i=cb.clip_icon))
    a_reset = m.addAction("Rese&t Button Image", lambda: cb.modify(key, kind, i, i=None))
    a_edit = m.addAction("&Edit Button Image...", lambda: edit_image(cb, key, kind, i))
    faces = m.addMenu("Change &Button Image")
    face_grid(faces, lambda name: cb.modify(key, kind, i, i=png64(face_icon(name))))
    plain = it.get("id", "").startswith("text:") or it.get("id", "").startswith("w:")
    for a in (a_copy, a_reset, a_edit, faces.menuAction()):
        a.setEnabled(not plain)
    a_paste.setEnabled(not plain and bool(cb.clip_icon))
    m.addSeparator()
    group = QActionGroup(m)
    group.setExclusionPolicy(QActionGroup.ExclusionPolicy.None_)   # (VB6: a check, not a dot)
    for n, text in enumerate(["Defa&ult Style", "Te&xt Only (Always)", "Text &Only (in Menus)", "Image &and Text"]):
        a = m.addAction(text, lambda n=n: cb.modify(key, kind, i, s=n))
        a.setCheckable(True)
        a.setChecked(it.get("s", 0) == n)
        group.addAction(a)
    m.addSeparator()
    a = m.addAction("Begin a &Group", lambda: cb.modify(key, kind, i, g=not it.get("g", False)))
    a.setCheckable(True)
    a.setChecked(bool(it.get("g")))
    a.setEnabled(i > 0)
    return m


# ================================================================ the pictures of Change Button Image
def face_icon(name, size=16):
    """42 pictures of our own for the buttons (VB6 has a grid of 6 x 7)"""
    img = QImage(16, 16, QImage.Format.Format_ARGB32)
    img.fill(0)
    p = QPainter(img)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, False)
    draw_face(p, name)
    p.end()
    if size != 16:
        img = img.scaled(size, size)
    return QIcon(QPixmap.fromImage(img))


def draw_face(p, name):
    K, W, Y, R, B, G, N = (QColor(0, 0, 0), QColor(255, 255, 255), QColor(255, 255, 0), QColor(255, 0, 0),
                           QColor(0, 0, 128), QColor(128, 128, 128), QColor(128, 128, 0))
    p.setPen(K)

    def poly(pts, fill):
        from PyQt6.QtGui import QPolygon
        p.setBrush(fill)
        p.drawPolygon(QPolygon([QPoint(x, y) for x, y in pts]))

    if name in ("smile", "frown"):
        p.setBrush(Y if name == "smile" else W)
        p.drawEllipse(1, 1, 13, 13)
        p.drawPoint(5, 6)
        p.drawPoint(10, 6)
        if name == "smile":
            p.drawArc(4, 5, 8, 6, 200 * 16, 140 * 16)
        else:
            p.drawArc(4, 9, 8, 6, 20 * 16, 140 * 16)
    elif name == "piggybank":
        p.setBrush(G)
        p.drawEllipse(2, 4, 12, 8)
        p.drawRect(4, 11, 1, 2)
        p.drawRect(11, 11, 1, 2)
        p.drawLine(7, 4, 9, 4)
    elif name == "inkbottle":
        p.setBrush(G)
        p.drawRect(4, 7, 8, 7)
        p.setBrush(K)
        p.drawRect(6, 2, 4, 5)
    elif name in ("disk-out", "disk-in"):
        p.setBrush(N)
        p.drawRect(2, 2, 12, 12)
        p.setBrush(W)
        p.drawRect(4, 2, 8, 5)
        p.setPen(QPen(B, 2))
        if name == "disk-out":
            p.drawLine(7, 9, 13, 3)
        else:
            p.drawLine(13, 3, 8, 8)
    elif name == "microphone":
        p.setBrush(G)
        p.drawEllipse(9, 1, 5, 5)
        p.drawLine(10, 6, 3, 13)
        p.drawLine(2, 14, 6, 14)
    elif name == "speaker":
        p.setBrush(Y)
        p.drawRect(2, 6, 3, 4)
        poly([(5, 6), (9, 2), (9, 14), (5, 10)], Y)
        p.drawArc(9, 4, 5, 8, -60 * 16, 120 * 16)
    elif name == "bell":
        poly([(4, 12), (5, 5), (8, 2), (11, 5), (12, 12)], Y)
        p.drawLine(3, 12, 13, 12)
        p.drawPoint(8, 14)
    elif name == "note":
        p.setBrush(K)
        p.drawEllipse(4, 10, 4, 4)
        p.drawLine(8, 12, 8, 2)
        p.drawLine(8, 2, 11, 5)
    elif name == "telephone":
        p.setBrush(G)
        p.drawRoundedRect(2, 6, 12, 8, 2, 2)
        p.drawArc(2, 2, 12, 8, 0, 180 * 16)
        p.setBrush(W)
        p.drawEllipse(6, 8, 4, 4)
    elif name == "balloon":
        p.setBrush(W)
        p.drawEllipse(1, 2, 14, 9)
        poly([(4, 10), (3, 14), (7, 10)], W)
    elif name == "camera":
        p.setBrush(G)
        p.drawRect(1, 5, 10, 7)
        poly([(11, 7), (15, 5), (15, 12), (11, 10)], G)
        p.drawEllipse(2, 1, 4, 4)
    elif name == "calculator":
        p.setBrush(G)
        p.drawRect(2, 1, 12, 14)
        p.setBrush(W)
        p.drawRect(4, 3, 8, 3)
        for y in (8, 11):
            for x in (4, 7, 10):
                p.drawRect(x, y, 1, 1)
    elif name in ("heart", "diamond"):
        p.setPen(R)
        if name == "heart":
            poly([(8, 14), (2, 7), (2, 4), (4, 2), (6, 2), (8, 4), (10, 2), (12, 2), (14, 4), (14, 7)], R)
        else:
            poly([(8, 1), (14, 8), (8, 15), (2, 8)], R)
    elif name in ("spade", "club"):
        p.setBrush(K)
        if name == "spade":
            poly([(8, 1), (14, 8), (12, 11), (8, 9), (4, 11), (2, 8)], K)
        else:
            p.drawEllipse(5, 1, 6, 6)
            p.drawEllipse(1, 5, 6, 6)
            p.drawEllipse(9, 5, 6, 6)
        poly([(8, 8), (11, 14), (5, 14)], K)
    elif name == "can":
        p.setBrush(B)
        p.drawRect(3, 4, 10, 10)
        p.setBrush(R)
        p.drawRect(3, 2, 10, 3)
    elif name == "mug":
        p.setBrush(W)
        p.drawRect(3, 4, 8, 10)
        p.drawArc(9, 6, 5, 5, -90 * 16, 180 * 16)
    elif name == "trash":
        p.setBrush(W)
        p.drawRect(3, 4, 10, 11)
        p.drawLine(2, 3, 14, 3)
        for x in (6, 8, 10):
            p.drawLine(x, 6, x, 13)
    elif name == "pencil":
        p.setPen(QPen(Y, 3))
        p.drawLine(4, 3, 12, 13)
        p.setPen(K)
        p.drawLine(3, 2, 4, 3)
    elif name == "square":
        p.setPen(G)
        p.setBrush(W)
        p.drawRect(2, 2, 12, 12)
    elif name == "footprints":
        p.setBrush(G)
        p.drawEllipse(2, 2, 4, 6)
        p.drawEllipse(9, 7, 4, 6)
    elif name in ("left", "right", "up", "down"):
        pts = {"left": [(2, 8), (8, 2), (8, 6), (14, 6), (14, 10), (8, 10), (8, 14)],
               "right": [(14, 8), (8, 2), (8, 6), (2, 6), (2, 10), (8, 10), (8, 14)],
               "up": [(8, 2), (2, 8), (6, 8), (6, 14), (10, 14), (10, 8), (14, 8)],
               "down": [(8, 14), (2, 8), (6, 8), (6, 2), (10, 2), (10, 8), (14, 8)]}[name]
        p.setPen(B)
        poly(pts, B)
    elif name == "pushpin":
        p.setBrush(W)
        p.drawEllipse(8, 1, 6, 6)
        p.drawLine(9, 7, 2, 14)
    elif name == "keyboard":
        p.setBrush(N)
        p.drawRect(1, 6, 14, 8)
        for x in range(3, 14, 2):
            p.drawPoint(x, 9)
            p.drawPoint(x, 11)
        p.drawLine(8, 6, 11, 2)
    elif name == "book":
        p.setBrush(W)
        p.drawRect(1, 3, 7, 11)
        p.drawRect(8, 3, 7, 11)
        for y in (6, 9, 12):
            p.drawLine(3, y, 6, y)
            p.drawLine(10, y, 13, y)
    elif name == "sound":
        p.setBrush(Y)
        p.drawRect(2, 6, 3, 4)
        for r in (3, 6, 9):
            p.drawArc(5 - r // 2, 8 - r, r * 2, r * 2, -45 * 16, 90 * 16)
    elif name == "key":
        p.setBrush(W)
        p.drawEllipse(1, 5, 6, 6)
        p.drawLine(7, 8, 15, 8)
        p.drawLine(12, 8, 12, 11)
        p.drawLine(14, 8, 14, 11)
    elif name == "gears":
        p.setBrush(G)
        p.drawEllipse(1, 1, 8, 8)
        p.drawEllipse(7, 7, 8, 8)
        p.setBrush(W)
        p.drawEllipse(4, 4, 2, 2)
        p.drawEllipse(10, 10, 2, 2)
    elif name == "scales":
        p.drawLine(8, 2, 8, 13)
        p.drawLine(2, 4, 14, 4)
        p.drawLine(5, 14, 11, 14)
        p.drawArc(1, 4, 5, 5, 180 * 16, 180 * 16)
        p.drawArc(10, 4, 5, 5, 180 * 16, 180 * 16)
    elif name == "hourglass":
        p.setBrush(W)
        poly([(3, 2), (13, 2), (8, 8), (13, 14), (3, 14), (8, 8)], W)
        p.drawLine(2, 1, 14, 1)
        p.drawLine(2, 15, 14, 15)
    elif name == "person":
        p.setBrush(K)
        p.drawEllipse(6, 1, 4, 4)
        p.drawRect(6, 6, 4, 9)
    elif name == "head":
        p.setBrush(W)
        p.drawEllipse(3, 1, 10, 12)
        p.setBrush(K)
        p.drawChord(3, 1, 10, 12, 30 * 16, 150 * 16)
    elif name == "runner":
        p.setBrush(K)
        p.drawEllipse(9, 1, 3, 3)
        p.drawLine(10, 4, 7, 9)
        p.drawLine(7, 9, 11, 13)
        p.drawLine(7, 9, 3, 13)
        p.drawLine(9, 6, 13, 7)
        p.drawLine(9, 6, 5, 6)
    elif name == "eye":
        p.setBrush(W)
        p.drawEllipse(1, 4, 14, 8)
        p.setBrush(K)
        p.drawEllipse(6, 5, 5, 5)
    elif name == "eightball":
        p.setBrush(K)
        p.drawEllipse(1, 1, 14, 14)
        p.setBrush(W)
        p.drawEllipse(5, 4, 6, 7)
    elif name == "question":
        p.setPen(QPen(Y.darker(130), 2))
        p.drawArc(4, 1, 8, 7, -60 * 16, 240 * 16)
        p.drawLine(8, 8, 8, 10)
        p.drawPoint(8, 13)


def face_grid(menu, picked):
    """the 6 x 7 pictures of Change Button Image, in a grid"""
    box = QWidget()
    g = QGridLayout(box)
    g.setContentsMargins(4, 4, 4, 4)
    g.setSpacing(2)
    for n, name in enumerate(DATA["faces"]):
        b = QToolButton()
        b.setIcon(face_icon(name))
        b.setAutoRaise(True)
        b.setFixedSize(23, 22)
        b.setToolTip(name)
        b.setObjectName("face-" + name)
        b.clicked.connect(lambda _c=False, nm=name: (picked(nm), menu.close()))
        g.addWidget(b, n // 6, n % 6)
    wa = QWidgetAction(menu)
    wa.setDefaultWidget(box)
    menu.addAction(wa)
    menu.cmdbars_faces = box


# ================================================================ the Button Editor
COLORS = ["#FFFFFF", "#FFFF00", "#00FF00", "#00FFFF", "#C0C0C0", "#808000", "#008000", "#008080",
          "#0000FF", "#FF00FF", "#FF0000", "#808080", "#000080", "#800080", "#800000", "#000000"]


class PixelGrid(QWidget):
    """the Picture of the Button Editor: 16 x 16 cells"""
    CELL = 8

    def __init__(self, editor):
        super().__init__()
        self.ed = editor
        self.setFixedSize(16 * self.CELL + 3, 16 * self.CELL + 3)

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("black"))
        c = self.CELL
        for y in range(16):
            for x in range(16):
                v = self.ed.pixels[y][x]
                r = QRect(2 + x * c, 2 + y * c, c - 1, c - 1)
                if v is None:
                    p.fillRect(r, QColor("white"))
                    p.setPen(QColor(192, 192, 192))
                    p.drawLine(r.bottomLeft(), r.topRight())
                else:
                    p.fillRect(r, QColor(v))
        p.end()

    def cell(self, pos):
        x, y = (pos.x() - 2) // self.CELL, (pos.y() - 2) // self.CELL
        return (x, y) if 0 <= x < 16 and 0 <= y < 16 else None

    def mousePressEvent(self, e):
        c = self.cell(e.position().toPoint())
        if c is None:
            return
        x, y = c
        cur = self.ed.color
        self.paint_value = None if self.ed.pixels[y][x] == cur else cur   # VB6: a click on a cell of the same colour
        self.ed.set_pixel(x, y, self.paint_value)                          # clears it

    def mouseMoveEvent(self, e):
        c = self.cell(e.position().toPoint())
        if c is not None and e.buttons() & Qt.MouseButton.LeftButton:
            self.ed.set_pixel(c[0], c[1], getattr(self, "paint_value", self.ed.color))


class ButtonEditor(QDialog):
    """Edit Button Image...: Picture, Colors (16 and Erase), Move, Preview; OK, Cancel, Clear (VB6 oracle r3)"""

    def __init__(self, parent, icon):
        super().__init__(parent)
        self.setWindowTitle("Button Editor")
        self.color = "#000000"
        img = icon.pixmap(16, 16).toImage().convertToFormat(QImage.Format.Format_ARGB32) if not icon.isNull() \
            else None
        self.pixels = [[None] * 16 for _ in range(16)]
        if img is not None and not img.isNull():
            for y in range(min(16, img.height())):
                for x in range(min(16, img.width())):
                    c = QColor.fromRgba(img.pixel(x, y))
                    self.pixels[y][x] = c.name().upper() if c.alpha() > 127 else None
        lay = QGridLayout(self)
        pic = QVBoxLayout()
        pic.addWidget(QLabel("Picture"))
        self.grid = PixelGrid(self)
        pic.addWidget(self.grid)
        lay.addLayout(pic, 0, 0)
        cols = QGroupBox("Colors")
        cg = QGridLayout(cols)
        self.color_buttons = []
        for n, c in enumerate(COLORS):
            b = QToolButton()
            b.setFixedSize(16, 16)
            b.setCheckable(True)
            b.setAutoExclusive(True)
            b.setStyleSheet(f"QToolButton {{ background: {c}; border: 1px solid gray }} "
                            "QToolButton:checked { border: 2px solid black }")
            b.clicked.connect(lambda _c=False, c=c: setattr(self, "color", c))
            b.setChecked(c == self.color)
            cg.addWidget(b, n // 4, n % 4)
            self.color_buttons.append(b)
        er = QHBoxLayout()
        er.addWidget(QLabel("Erase:"))
        self.erase = QToolButton()
        self.erase.setFixedSize(16, 16)
        self.erase.setCheckable(True)
        self.erase.setAutoExclusive(True)
        self.erase.clicked.connect(lambda: setattr(self, "color", None))
        er.addWidget(self.erase)
        cg.addLayout(er, 4, 0, 1, 4)
        lay.addWidget(cols, 0, 1)
        btns = QVBoxLayout()
        self.b_ok = QPushButton("OK")
        self.b_cancel = QPushButton("Cancel")
        self.b_clear = QPushButton("Clear")
        self.b_ok.clicked.connect(self.accept)
        self.b_cancel.clicked.connect(self.reject)
        self.b_clear.clicked.connect(self.clear)
        btns.addWidget(self.b_ok)
        btns.addWidget(self.b_cancel)
        btns.addSpacing(8)
        btns.addWidget(self.b_clear)
        btns.addStretch(1)
        lay.addLayout(btns, 0, 2)
        mv = QGroupBox("Move")
        mg = QGridLayout(mv)
        self.moves = {}
        for name, r, c, dx, dy in (("up", 0, 1, 0, -1), ("left", 1, 0, -1, 0), ("right", 1, 2, 1, 0),
                                   ("down", 2, 1, 0, 1)):
            b = QToolButton()
            b.setArrowType({"up": Qt.ArrowType.UpArrow, "down": Qt.ArrowType.DownArrow, "left": Qt.ArrowType.LeftArrow,
                            "right": Qt.ArrowType.RightArrow}[name])
            b.clicked.connect(lambda _c=False, dx=dx, dy=dy: self.shift(dx, dy))
            mg.addWidget(b, r, c)
            self.moves[name] = b
        lay.addWidget(mv, 1, 0)
        pv = QGroupBox("Preview")
        pl = QHBoxLayout(pv)
        self.preview = QLabel()
        self.preview.setFixedSize(24, 24)
        pl.addWidget(self.preview)
        pl.addStretch(1)
        lay.addWidget(pv, 1, 1)
        self.update_all()

    def set_pixel(self, x, y, v):
        self.pixels[y][x] = v
        self.update_all()

    def clear(self):
        self.pixels = [[None] * 16 for _ in range(16)]
        self.update_all()

    def bounds(self):
        xs = [x for y in range(16) for x in range(16) if self.pixels[y][x] is not None]
        ys = [y for y in range(16) for x in range(16) if self.pixels[y][x] is not None]
        return (min(xs), min(ys), max(xs), max(ys)) if xs else None

    def shift(self, dx, dy):
        old = self.pixels
        new = [[None] * 16 for _ in range(16)]
        for y in range(16):
            for x in range(16):
                if old[y][x] is not None and 0 <= x + dx < 16 and 0 <= y + dy < 16:
                    new[y + dy][x + dx] = old[y][x]
        self.pixels = new
        self.update_all()

    def update_all(self):
        b = self.bounds()                       # VB6: an arrow is grey when the picture is at that edge
        self.moves["up"].setEnabled(b is not None and b[1] > 0)
        self.moves["left"].setEnabled(b is not None and b[0] > 0)
        self.moves["right"].setEnabled(b is not None and b[2] < 15)
        self.moves["down"].setEnabled(b is not None and b[3] < 15)
        self.preview.setPixmap(self.icon().pixmap(16, 16))
        self.grid.update()

    def image(self):
        img = QImage(16, 16, QImage.Format.Format_ARGB32)
        img.fill(0)
        for y in range(16):
            for x in range(16):
                if self.pixels[y][x] is not None:
                    img.setPixelColor(x, y, QColor(self.pixels[y][x]))
        return img

    def icon(self):
        return QIcon(QPixmap.fromImage(self.image()))


def edit_image(cb, key, kind, i):
    it = cb.item_at(key, kind, i)
    dlg = ButtonEditor(cb.customizing or cb.mw, cb.item_icon(it))
    cb._button_editor = dlg
    if dlg.exec():
        cb.modify(key, kind, i, i=png64(dlg.icon()))
    return dlg


# ================================================================ the dialog
class MenuPanel(QFrame):
    """the items of a menu, open while a command is dragged over its name on the menu bar: the place where the
    command is dropped into the menu"""

    def __init__(self, cb, mid):
        super().__init__(None, Qt.WindowType.ToolTip | Qt.WindowType.FramelessWindowHint)
        self.cb, self.mid = cb, mid
        self.setFrameShape(QFrame.Shape.Box)
        self.setAcceptDrops(True)
        self.rows = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        lay.setSpacing(0)
        for it in cb.menu_items(mid):
            if it.get("g") and self.rows:
                ln = QFrame()
                ln.setFrameShape(QFrame.Shape.HLine)
                lay.addWidget(ln)
            row = QLabel("   " + norm(cb.item_name(it)) + ("  \u25b8" if it["id"].startswith(("menu:", "user:")) else ""))
            row.setMinimumWidth(160)
            row.setFixedHeight(20)
            lay.addWidget(row)
            self.rows.append(row)
        self.mark = QFrame(self)
        self.mark.setStyleSheet("background: black")
        self.mark.hide()
        self.adjustSize()

    def index_at(self, y):
        for i, r in enumerate(self.rows):
            g = r.geometry()
            if y < g.center().y():
                return i, g.top()
        return len(self.rows), (self.rows[-1].geometry().bottom() if self.rows else 2)

    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat(MIME):
            e.acceptProposedAction()

    def dragMoveEvent(self, e):
        if not e.mimeData().hasFormat(MIME):
            return
        _i, y = self.index_at(e.position().toPoint().y())
        self.mark.setGeometry(2, y - 1, self.width() - 4, 2)
        self.mark.show()
        ctrl = bool(e.modifiers() & Qt.KeyboardModifier.ControlModifier)
        e.setDropAction(Qt.DropAction.CopyAction if ctrl else Qt.DropAction.MoveAction)
        e.accept()

    def dropEvent(self, e):
        data = json.loads(bytes(e.mimeData().data(MIME)).decode())
        idx, _y = self.index_at(e.position().toPoint().y())
        src = data.get("src")
        ctrl = bool(e.modifiers() & Qt.KeyboardModifier.ControlModifier)
        if src is None:
            self.cb.insert(self.mid, "menu", idx, data["item"])
            e.setDropAction(Qt.DropAction.CopyAction)
        else:
            self.cb.move(tuple(src), (self.mid, "menu", idx), copy_it=ctrl)
            e.setDropAction(Qt.DropAction.CopyAction if ctrl else Qt.DropAction.MoveAction)
        e.accept()
        QTimer.singleShot(0, self.cb.filter.close_panel)


class CommandList(QListWidget):
    """the Commands of a category: dragged onto a bar or into a menu"""

    def __init__(self, cb):
        super().__init__()
        self.cb = cb
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
        self.setIconSize(QSize(16, 16))

    def startDrag(self, actions):
        it = self.currentItem()
        if it is None:
            return
        cid = it.data(Qt.ItemDataRole.UserRole)
        md = QMimeData()
        md.setData(MIME, json.dumps({"item": item(cid)}).encode())
        drag = QDrag(self)
        drag.setMimeData(md)
        if not it.icon().isNull():
            drag.setPixmap(it.icon().pixmap(16, 16))
        drag.exec(Qt.DropAction.CopyAction)
        self.cb.filter.close_panel()


class CustomizeDialog(QDialog):
    def __init__(self, cb):
        super().__init__(cb.mw)
        self.cb = cb
        self.setWindowTitle("Customize")
        self.setModal(False)
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs)
        self.tabs.addTab(self._toolbars_tab(), "Tool&bars")
        self.tabs.addTab(self._commands_tab(), "&Commands")
        self.tabs.addTab(self._options_tab(), "&Options")
        row = QHBoxLayout()
        row.addStretch(1)
        self.b_close = QPushButton("Close")
        self.b_close.clicked.connect(self.close)
        row.addWidget(self.b_close)
        lay.addLayout(row)
        self.resize(366, 336)

    def closeEvent(self, e):
        super().closeEvent(e)
        self.done(0)

    # ---------------------------------------------------------------- Toolbars
    def _toolbars_tab(self):
        w = QWidget()
        g = QGridLayout(w)
        lab = QLabel("Tool&bars:")
        g.addWidget(lab, 0, 0)
        self.bar_list = QListWidget()
        lab.setBuddy(self.bar_list)
        g.addWidget(self.bar_list, 1, 0, 5, 1)
        self.b_new = QPushButton("&New...")
        self.b_rename = QPushButton("&Rename...")
        self.b_delete = QPushButton("&Delete")
        self.b_reset = QPushButton("R&eset...")
        for n, b in enumerate((self.b_new, self.b_rename, self.b_delete, self.b_reset)):
            g.addWidget(b, 1 + n, 1)
        g.setRowStretch(5, 1)
        self.b_new.clicked.connect(self.new_bar)
        self.b_rename.clicked.connect(self.rename_bar)
        self.b_delete.clicked.connect(self.delete_bar)
        self.b_reset.clicked.connect(self.reset_bar)
        self.bar_list.currentRowChanged.connect(lambda _r: self.bar_buttons())
        self.bar_list.itemChanged.connect(self.bar_checked)
        self.fill_bars()
        return w

    def fill_bars(self, current=None):
        cb = self.cb
        self.bar_list.blockSignals(True)
        self.bar_list.clear()
        for name in cb.bar_names():
            it = QListWidgetItem(name)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            tb = cb.toolbar(name)
            on = True if name == "Menu Bar" else (tb is not None and tb.isVisible())
            it.setCheckState(Qt.CheckState.Checked if on else Qt.CheckState.Unchecked)
            self.bar_list.addItem(it)
        names = cb.bar_names()
        self.bar_list.setCurrentRow(names.index(current) if current in names else 0)
        self.bar_list.blockSignals(False)
        self.bar_buttons()

    def current_bar(self):
        it = self.bar_list.currentItem()
        return it.text() if it is not None else None

    def bar_buttons(self):
        name = self.current_bar()
        builtin = name in self.cb.BUILTIN
        self.b_rename.setEnabled(name is not None and not builtin)
        self.b_delete.setEnabled(name is not None and not builtin)
        self.b_reset.setEnabled(name is not None and builtin)

    def bar_checked(self, it):
        name = it.text()
        on = it.checkState() == Qt.CheckState.Checked
        if name == "Menu Bar":                  # VB6: the menu bar is always there
            if not on:
                self.bar_list.blockSignals(True)
                it.setCheckState(Qt.CheckState.Checked)
                self.bar_list.blockSignals(False)
            return
        if name == "Shortcut Menus":
            self.cb.make_shortcut_bar()
            self.cb.begin_customize()
        tb = self.cb.toolbar(name)
        if tb is not None:
            tb.setVisible(on)

    def ask_name(self, title, text):
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        v = QVBoxLayout(dlg)
        lab = QLabel("&Toolbar name:")
        ed = QLineEdit(text)
        ed.selectAll()
        lab.setBuddy(ed)
        v.addWidget(lab)
        v.addWidget(ed)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        dlg.resize(270, 120)
        self._name_dialog = (dlg, ed)
        if getattr(self, "answer", None) is not None:       # (the tests)
            ed.setText(self.answer)
            return self.answer
        return ed.text().strip() if dlg.exec() and ed.text().strip() else None

    def new_bar(self):
        name = self.ask_name("New Toolbar", self.cb.unique_name())
        if not name:
            return
        if name in self.cb.bar_names():
            self.message(f"A toolbar named '{name}' already exists. Type another name.", warn=True)
            return
        tb = self.cb.new_toolbar(name)
        tb.installEventFilter(self.cb.filter)
        tb.setAcceptDrops(True)
        self.fill_bars(name)

    def rename_bar(self):
        old = self.current_bar()
        name = self.ask_name("Rename Toolbar", old)
        if not name or name == old:
            return
        if name in self.cb.bar_names():
            self.message(f"A toolbar named '{name}' already exists. Type another name.", warn=True)
            return
        self.cb.rename_toolbar(old, name)
        self.fill_bars(name)

    def message(self, text, warn=False, ask=False):
        if getattr(self, "answers", None) is not None:      # (the tests)
            self.last_message = text
            return self.answers.pop(0) if self.answers else True
        icon = QMessageBox.Icon.Warning
        box = QMessageBox(icon, APP, text, QMessageBox.StandardButton.Ok |
                          (QMessageBox.StandardButton.Cancel if ask else QMessageBox.StandardButton.NoButton), self)
        return box.exec() == QMessageBox.StandardButton.Ok

    def delete_bar(self):
        name = self.current_bar()
        if self.message(DATA["messages"]["delete"].format(name), ask=True):
            self.cb.delete_toolbar(name)
            self.fill_bars()

    def reset_bar(self):
        name = self.current_bar()
        if self.message(DATA["messages"]["reset"].format(name), ask=True):
            self.cb.reset_toolbar(name)
            self.fill_bars(name)

    # ---------------------------------------------------------------- Commands
    def _commands_tab(self):
        w = QWidget()
        g = QGridLayout(w)
        l1 = QLabel("Cate&gories:")
        l2 = QLabel("Comman&ds:")
        self.cat_list = QListWidget()
        self.cmd_list = CommandList(self.cb)
        l1.setBuddy(self.cat_list)
        l2.setBuddy(self.cmd_list)
        g.addWidget(l1, 0, 0)
        g.addWidget(l2, 0, 1)
        g.addWidget(self.cat_list, 1, 0)
        g.addWidget(self.cmd_list, 1, 1)
        g.setColumnStretch(1, 2)
        g.setColumnStretch(0, 1)
        self.sel_label = QLabel("Selected command:")
        g.addWidget(self.sel_label, 2, 0, 1, 2)
        row = QHBoxLayout()
        self.b_desc = QPushButton("Descr&iption")
        self.b_modify = QPushButton("&Modify Selection ▾")
        row.addWidget(self.b_desc)
        row.addWidget(self.b_modify)
        g.addLayout(row, 3, 0, 1, 2)
        self.cat_list.addItems(DATA["categories"])
        self.cat_list.currentRowChanged.connect(self.fill_commands)
        self.cmd_list.currentRowChanged.connect(lambda _r: self.selection_changed())
        self.b_desc.clicked.connect(self.description)
        self.b_modify.clicked.connect(self.modify_selection)
        self.cat_list.setCurrentRow(0)
        return w

    def fill_commands(self, row):
        reg = self.cb.reg
        if not hasattr(self, "blank"):
            pm = QPixmap(16, 16)
            pm.fill(Qt.GlobalColor.transparent)
            self.blank = QIcon(pm)
        self.cmd_list.clear()
        if row < 0:
            return
        cat = DATA["categories"][row]
        for cid in reg.cats.get(cat, []):
            name = reg.name(cid)
            it = QListWidgetItem(name + ("  ▸" if cid.startswith("menu:") else ""))
            a = reg.action(cid)
            it.setIcon(a.icon() if a is not None and not a.icon().isNull() else self.blank)   # VB6: the names
            it.setSizeHint(QSize(0, 22))                                                     # in one column
            it.setData(Qt.ItemDataRole.UserRole, cid)
            self.cmd_list.addItem(it)
        self.selection_changed()

    def selection_changed(self):
        cur = self.cmd_list.currentItem() if hasattr(self, "cmd_list") else None
        if not hasattr(self, "b_desc"):
            return
        self.b_desc.setEnabled(cur is not None)
        self.sel_label.setEnabled(cur is not None or self.cb.selected is not None)
        self.b_modify.setEnabled(self.cb.selected is not None)

    def description(self):
        cur = self.cmd_list.currentItem()
        if cur is None:
            return None
        cid = cur.data(Qt.ItemDataRole.UserRole)
        a = self.cb.reg.action(cid)
        text = (a.statusTip() or a.toolTip()) if a is not None else ""
        if not text or norm(text) == norm(a.text() if a is not None else ""):
            text = describe(self.cb.reg.name(cid), cid)
        self.last_description = text
        QToolTip.showText(self.b_desc.mapToGlobal(QPoint(0, self.b_desc.height())), text, self.b_desc)
        return text

    def modify_selection(self):
        sel = self.cb.selected
        if sel is None:
            return None
        w, i = sel
        key, kind = self.cb.where(w)
        m = modify_menu(self.cb, w, key, kind, i, self)
        self._modify_menu = m
        if getattr(self, "answers", None) is None:
            m.exec(self.b_modify.mapToGlobal(QPoint(0, self.b_modify.height())))
        return m

    # ---------------------------------------------------------------- Options
    def _options_tab(self):
        w = QWidget()
        v = QVBoxLayout(w)
        o = self.cb.options()
        self.c_large = QCheckBox("&Large icons")
        self.c_tips = QCheckBox("Show Screen&Tips on toolbars")
        self.c_keys = QCheckBox("Show s&hortcut keys in ScreenTips")
        self.c_large.setChecked(bool(o["large"]))
        self.c_tips.setChecked(bool(o["tips"]))
        self.c_keys.setChecked(bool(o["keys"]))
        self.c_keys.setEnabled(bool(o["tips"]))
        v.addSpacing(16)
        for c in (self.c_large, self.c_tips, self.c_keys):
            v.addWidget(c)
        v.addSpacing(30)
        row = QHBoxLayout()
        lab = QLabel("&Menu animations:")
        self.anim = QComboBox()
        self.anim.addItems(DATA["animations"])
        self.anim.setCurrentIndex(int(o["anim"]))
        lab.setBuddy(self.anim)
        row.addWidget(lab)
        row.addWidget(self.anim)
        row.addStretch(1)
        v.addLayout(row)
        v.addStretch(1)
        self.c_large.toggled.connect(lambda on: self.cb.set_option("large", on))
        self.c_tips.toggled.connect(self._tips)
        self.c_keys.toggled.connect(lambda on: self.cb.set_option("keys", on))
        self.anim.currentIndexChanged.connect(lambda i: self.cb.set_option("anim", i))
        return w

    def _tips(self, on):
        self.cb.set_option("tips", on)
        self.c_keys.setEnabled(on)


def describe(name, cid):
    """the Description of a command that has no text of its own"""
    if cid.startswith("menu:"):
        return f"Shows the {name} menu."
    return name.rstrip(".") if name else ""


# ================================================================ the ScreenTips
class TipFilter(QObject):
    """Show ScreenTips on toolbars, Show shortcut keys in ScreenTips; Random menu animations"""

    def __init__(self, cb):
        super().__init__(cb)
        self.cb = cb

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.Type.ToolTip and isinstance(obj, QToolBar):
            a = obj.actionAt(ev.pos())
            if a is None:
                return False
            t = self.cb.tip_text(a)
            if t is None:
                QToolTip.hideText()
                ev.ignore()
                return True
            QToolTip.showText(ev.globalPos(), t, obj)
            return True
        if ev.type() == QEvent.Type.MouseButtonPress and isinstance(obj, QMenuBar):
            if int(self.cb.options().get("anim", 0) or 0) == 1:      # Random: unfolding or sliding, or none
                QApplication.instance().setEffectEnabled(Qt.UIEffect.UI_AnimateMenu, random.random() < 0.67)
        return False
