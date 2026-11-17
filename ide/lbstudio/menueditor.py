"""Tools > Menu Editor (Ctrl+E) of VB6: the menus of a form as an indented list, edited one item at
a time, written back as Begin VB.Menu blocks at the end of the form."""
import copy

from PyQt6.QtWidgets import (QDialog, QGridLayout, QLabel, QLineEdit, QComboBox, QCheckBox, QPushButton,
                             QListWidget, QHBoxLayout, QMessageBox, QToolButton, QVBoxLayout)

from lbvc.project.source import Control
from . import vbprops as V

FIELDS = ("caption", "name", "index", "shortcut", "helpid", "negotiate", "checked", "enabled", "visible",
          "windowlist")


def _entry(level=0, ctl=None):
    e = {"level": level, "ctl": ctl, "caption": "", "name": "", "index": None, "shortcut": "", "helpid": 0,
         "negotiate": 0, "checked": False, "enabled": True, "visible": True, "windowlist": False}
    if ctl is not None:
        g = lambda n, d: V.value_of(ctl, V.prop_def("Menu", n))     # noqa: E731
        e.update(caption=str(ctl.get("Caption", "")), name=ctl.name, index=ctl.index,
                 shortcut=str(ctl.get("Shortcut", "") or ""), helpid=g("HelpContextID", 0),
                 negotiate=g("NegotiatePosition", 0), checked=g("Checked", False), enabled=g("Enabled", True),
                 visible=g("Visible", True), windowlist=g("WindowList", False))
    return e


def flatten(root):
    out = []

    def walk(c, level):
        for m in c.children:
            if m.type_name == "Menu":
                out.append(_entry(level, m))
                walk(m, level + 1)
    walk(root, 0)
    return out


def validate(entries, other_names=()):
    """the checks of the VB6 Menu Editor; (row, message) or None"""
    real = [(i, e) for i, e in enumerate(entries) if e["caption"] or e["name"]]
    prev_level = -1
    names = {}
    shortcuts = set()
    winlist = 0
    for k, (i, e) in enumerate(real):
        if not e["name"]:
            return i, "Menu control must have a name"
        if e["level"] > prev_level + 1:
            return i, "Menu item skipped a level"
        prev_level = e["level"]
        has_sub = k + 1 < len(real) and real[k + 1][1]["level"] > e["level"]
        sep = e["caption"] == "-"
        if sep and (e["checked"] or not e["enabled"] or e["shortcut"]):
            return i, "Separator may not be checked or disabled, or have a shortcut key"
        if sep and e["windowlist"]:
            return i, "Separator may not be the Window List menu"
        if e["checked"] and (e["level"] == 0 or has_sub):
            return i, "Can't put check mark here"
        if e["shortcut"] and (e["level"] == 0 or has_sub):
            return i, "Can't assign shortcut key to a top level menu"
        if e["shortcut"]:
            if e["shortcut"].upper() in shortcuts:
                return i, "Shortcut key already assigned"
            shortcuts.add(e["shortcut"].upper())
        if e["windowlist"]:
            winlist += 1
            if winlist > 1:
                return i, "Can't have more than one Window List menu"
        if e["negotiate"] and e["level"] != 0:
            return i, "Only Top Level Menus can have non-zero NegotiatePosition"
        key = e["name"].lower()
        if key in {n.lower() for n in other_names}:
            return i, f"Ambiguous name detected: {e['name']}"
        if key in names:
            first = names[key]
            if e["index"] is None or first[1]["index"] is None:
                return i, "Menu control array element must have an index"
            if e["index"] <= first[2]:
                return i, "Menu control array indexes must be in ascending order"
            if real[k - 1][1]["name"].lower() != key and not _same_sub(real, first[0], k):
                return i, "Menu control array elements must be contiguous and within the same submenu"
            names[key] = (first[0], first[1], e["index"])
        else:
            names[key] = (k, e, e["index"] if e["index"] is not None else -1)
    # a submenu needs a visible item
    for k, (i, e) in enumerate(real):
        subs = []
        for j in range(k + 1, len(real)):
            if real[j][1]["level"] <= e["level"]:
                break
            if real[j][1]["level"] == e["level"] + 1:
                subs.append(real[j][1])
        if subs and not any(s["visible"] for s in subs):
            return i, "At least one submenu item must be visible"
    return None


def _same_sub(real, a, b):
    """the elements a..b of an array are siblings (the ones between are inside them)"""
    lev = real[a][1]["level"]
    return all(real[j][1]["level"] > lev or real[j][1]["name"].lower() == real[a][1]["name"].lower()
               for j in range(a + 1, b))


def build(entries):
    """the entries -> Menu controls (a tree), keeping the other properties of the old ones"""
    tops, stack = [], []
    for e in entries:
        if not (e["caption"] or e["name"]):
            continue
        if e["ctl"] is not None:
            c = copy.deepcopy(e["ctl"])
            c.children = []
            c.order = [(k, x) for k, x in c.order if k == "p"]
            if c.name != e["name"]:
                c.name = e["name"]
                c.raw_begin = None
        else:
            c = Control("VB.Menu", e["name"])
        for prop, key in (("Caption", "caption"), ("Checked", "checked"), ("Enabled", "enabled"),
                          ("HelpContextID", "helpid"), ("Index", "index"), ("NegotiatePosition", "negotiate"),
                          ("Shortcut", "shortcut"), ("Visible", "visible"), ("WindowList", "windowlist")):
            V.apply(c, V.prop_def("Menu", prop), e[key])
        while stack and stack[-1][0] >= e["level"]:
            stack.pop()
        if stack:
            stack[-1][1].add_child(c)
        else:
            tops.append(c)
        stack.append((e["level"], c))
    for c in tops:
        V.clear_raw(c)
    return tops


class MenuEditor(QDialog):
    def __init__(self, parent, root, other_names=()):
        super().__init__(parent)
        self.setWindowTitle("Menu Editor")
        self.root = root
        self.other_names = other_names
        self.entries = flatten(root) + [_entry(0)]
        self._busy = False
        g = QGridLayout(self)
        self.caption = QLineEdit()
        self.name = QLineEdit()
        self.index = QLineEdit()
        self.index.setMaximumWidth(60)
        self.shortcut = QComboBox()
        for t, code in V.SHORTCUTS:
            self.shortcut.addItem(t, code)
        self.helpid = QLineEdit("0")
        self.helpid.setMaximumWidth(80)
        self.negotiate = QComboBox()
        for n, t in V.ENUMS["negotiate"]:
            self.negotiate.addItem(f"{n} - {t}", n)
        self.checked = QCheckBox("&Checked")
        self.enabled = QCheckBox("&Enabled")
        self.visible = QCheckBox("&Visible")
        self.windowlist = QCheckBox("&WindowList")
        for text, w, r, c, span in (("Ca&ption:", self.caption, 0, 0, 3), ("Na&me:", self.name, 1, 0, 3),
                                     ("Inde&x:", self.index, 2, 0, 1), ("&Shortcut:", self.shortcut, 2, 2, 1),
                                     ("&HelpContextID:", self.helpid, 3, 0, 1),
                                     ("Neg&otiatePosition:", self.negotiate, 3, 2, 1)):
            lab = QLabel(text)
            lab.setBuddy(w)
            g.addWidget(lab, r, c)
            g.addWidget(w, r, c + 1, 1, span)
        checks = QHBoxLayout()
        for w in (self.checked, self.enabled, self.visible, self.windowlist):
            checks.addWidget(w)
        g.addLayout(checks, 4, 0, 1, 4)
        row = QHBoxLayout()
        for text, slot in (("←", lambda: self._level(-1)), ("→", lambda: self._level(1)),
                           ("↑", lambda: self._move(-1)), ("↓", lambda: self._move(1))):
            b = QToolButton()
            b.setText(text)
            b.clicked.connect(slot)
            row.addWidget(b)
        row.addStretch()
        for text, slot in (("&Next", self._next), ("&Insert", self._insert), ("Dele&te", self._delete)):
            b = QPushButton(text)
            b.clicked.connect(slot)
            row.addWidget(b)
        g.addLayout(row, 5, 0, 1, 4)
        self.list = QListWidget()
        self.list.setMinimumHeight(180)
        g.addWidget(self.list, 6, 0, 1, 4)
        side = QVBoxLayout()
        ok = QPushButton("OK")
        ok.setDefault(True)
        cancel = QPushButton("Cancel")
        ok.clicked.connect(self._ok)
        cancel.clicked.connect(self.reject)
        side.addWidget(ok)
        side.addWidget(cancel)
        side.addStretch()
        g.addLayout(side, 0, 4, 3, 1)
        self.caption.textEdited.connect(lambda t: self._field("caption", t))
        self.name.textEdited.connect(lambda t: self._field("name", t))
        self.index.textEdited.connect(self._index_edited)
        self.helpid.textEdited.connect(lambda t: self._field("helpid", int(t) if t.strip().isdigit() else 0))
        self.shortcut.activated.connect(lambda i: self._field("shortcut", self.shortcut.itemData(i)))
        self.negotiate.activated.connect(lambda i: self._field("negotiate", self.negotiate.itemData(i)))
        for key in ("checked", "enabled", "visible", "windowlist"):
            getattr(self, key).toggled.connect(lambda on, key=key: self._field(key, on))
        self.list.currentRowChanged.connect(self._show)
        self._fill()
        self.list.setCurrentRow(0)
        self.result_menus = None

    # ------------------------------------------------------------ the list
    def _label(self, e):
        text = e["caption"] if e["caption"] else ""
        sc = V.TEXT_OF_SHORTCUT.get(e["shortcut"].upper(), "") if e["shortcut"] else ""
        return "····" * e["level"] + text + (f"\t{sc}" if sc else "")

    def _fill(self):
        cur = self.list.currentRow()
        self._busy = True
        self.list.clear()
        for e in self.entries:
            self.list.addItem(self._label(e))
        self._busy = False
        if 0 <= cur < self.list.count():
            self.list.setCurrentRow(cur)

    def _show(self, r):
        if self._busy or not 0 <= r < len(self.entries):
            return
        e = self.entries[r]
        self._busy = True
        self.caption.setText(e["caption"])
        self.name.setText(e["name"])
        self.index.setText("" if e["index"] is None else str(e["index"]))
        self.helpid.setText(str(e["helpid"]))
        self.shortcut.setCurrentIndex(max(0, self.shortcut.findData(e["shortcut"])))
        self.negotiate.setCurrentIndex(max(0, self.negotiate.findData(e["negotiate"])))
        for key in ("checked", "enabled", "visible", "windowlist"):
            getattr(self, key).setChecked(bool(e[key]))
        self._busy = False

    def _field(self, key, value):
        if self._busy:
            return
        r = self.list.currentRow()
        if not 0 <= r < len(self.entries):
            return
        self.entries[r][key] = value
        self.list.item(r).setText(self._label(self.entries[r]))

    def _index_edited(self, t):
        t = t.strip()
        self._field("index", int(t) if t.isdigit() else None)

    def _level(self, d):
        r = self.list.currentRow()
        if 0 <= r < len(self.entries):
            e = self.entries[r]
            e["level"] = max(0, min(e["level"] + d, 5))
            self.list.item(r).setText(self._label(e))

    def _move(self, d):
        r = self.list.currentRow()
        j = r + d
        if 0 <= r < len(self.entries) and 0 <= j < len(self.entries):
            self.entries[r], self.entries[j] = self.entries[j], self.entries[r]
            self._fill()
            self.list.setCurrentRow(j)

    def _next(self):
        r = self.list.currentRow()
        if r == len(self.entries) - 1:
            if not (self.entries[r]["caption"] or self.entries[r]["name"]):
                self.caption.setFocus()
                return
            self.entries.append(_entry(self.entries[r]["level"]))
            self._fill()
        self.list.setCurrentRow(r + 1)
        self.caption.setFocus()

    def _insert(self):
        r = max(0, self.list.currentRow())
        self.entries.insert(r, _entry(self.entries[r]["level"] if r < len(self.entries) else 0))
        self._fill()
        self.list.setCurrentRow(r)
        self.caption.setFocus()

    def _delete(self):
        r = self.list.currentRow()
        if 0 <= r < len(self.entries) - 1 or (r == len(self.entries) - 1 and len(self.entries) > 1):
            del self.entries[r]
            if not self.entries:
                self.entries.append(_entry(0))
            self._fill()
            self.list.setCurrentRow(min(r, len(self.entries) - 1))

    def _ok(self):
        bad = validate(self.entries, self.other_names)
        if bad:
            row, msg = bad
            self.list.setCurrentRow(row)
            QMessageBox.critical(self, "Menu Editor", msg)
            return
        self.result_menus = build(self.entries)
        self.accept()
