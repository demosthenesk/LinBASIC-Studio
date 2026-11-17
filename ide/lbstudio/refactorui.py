"""The refactoring commands of the IDE (Phase 5): Definition, Find References, Rename Symbol, Extract
Procedure, on the project as the code windows have it (refactor.py binds the names)."""
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QBrush, QColor
from PyQt6.QtWidgets import QTreeWidget, QTreeWidgetItem, QInputDialog, QMessageBox, QLineEdit

from . import refactor as R


class Unit:
    """a module as the analysis sees it: its name, kind, the names of its controls"""

    def __init__(self, m):
        self.m = m
        self.name = m.name
        self.kind = "Form" if m.kind in ("Form", "UserControl", "PropertyPage", "UserDocument") else \
            "Class" if m.kind in ("Class", "Designer") else "Module"
        self.controls_list = []
        if m.kind in ("Form", "UserControl", "PropertyPage", "UserDocument") and m.sf is not None and m.sf.root is not None:
            seen = set()
            for c in m.sf.root.walk():
                if c is not m.sf.root and c.name.lower() not in seen:
                    seen.add(c.name.lower())
                    self.controls_list.append(c.name)


def module_text(main, m):
    sw = main.code_windows.get(m)
    if sw is not None:
        return sw.widget().text().replace("\r\n", "\n")
    return m.code_text().replace("\r\n", "\n")


def analyze(main, project=None):
    """the analysis of the project the IDE works on (the code windows' texts)"""
    p = project or main.project
    units = [Unit(m) for m in p.modules if m.sf is not None]
    return R.Project([(u, module_text(main, u.m)) for u in units])


def caret(main):
    """(code window, line, column) of the caret: 0-based, the column in characters"""
    cw = main.current_cw()
    if cw is None:
        return None, 0, 0
    ed = cw.editor
    p = ed.pos()
    line = ed.line_from_pos(p)
    raw = ed.get_bytes(ed.line_start(line), p)
    return cw, line, len(raw.decode("utf-8", errors="replace"))


def symbol_here(main):
    cw, line, col = caret(main)
    if cw is None:
        return None, None, None
    pr = analyze(main)
    s = pr.symbol_at(cw.module.name, line, col)
    if s is None and col > 0:
        s = pr.symbol_at(cw.module.name, line, col - 1)
    return pr, s, cw


def _byte_pos(ed, line, col):
    text = ed.get_bytes(ed.line_start(line), ed.line_end(line)).decode("utf-8", errors="replace")
    return ed.line_start(line) + len(text[:col].encode("utf-8"))


def select_range(main, m, line, col, length):
    cw = main.open_code(m)
    ed = cw.editor
    a = _byte_pos(ed, line, col)
    b = _byte_pos(ed, line, col + length)
    ed.goto_line(line)
    ed.set_sel(a, b)
    ed.setFocus()


# ====================================================================== the References tab
class ReferencesWindow(QTreeWidget):
    """the places of a symbol: double click opens the code there"""
    goto = pyqtSignal(object, int, int, int)

    def __init__(self):
        super().__init__()
        self.setHeaderLabels(["Module", "Line", "Code"])
        self.setRootIsDecorated(False)
        self.setColumnWidth(0, 160)
        self.setColumnWidth(1, 60)
        self.itemActivated.connect(self._go)

    def show_refs(self, sym, refs):
        self.clear()
        for r in refs:
            text = r.mod.lines[r.line].strip() if r.line < len(r.mod.lines) else ""
            it = QTreeWidgetItem([r.mod.name, str(r.line + 1), text + ("" if r.certain else "   (possible: an object of unknown type)")])
            it.setData(0, Qt.ItemDataRole.UserRole, (r.mod.module.m, r.line, r.col, r.length))
            if not r.certain:
                for c in range(3):
                    it.setForeground(c, QBrush(QColor(128, 128, 128)))
            if r.decl:
                f = it.font(2)
                f.setBold(True)
                it.setFont(2, f)
            self.addTopLevelItem(it)

    def _go(self, it, col):
        d = it.data(0, Qt.ItemDataRole.UserRole)
        if d:
            self.goto.emit(*d)


# ====================================================================== the commands
def definition(main):
    pr, s, cw = symbol_here(main)
    if s is None:
        # a name of a library (MsgBox, Text1.Text): VB6 shows it in the Object Browser
        if hasattr(main, "browse_word") and main.browse_word():
            return
        main.statusBar().showMessage("Identifier under cursor is not recognized", 4000)
        return
    if s.kind == "module":
        m = s.mod.module.m
        if m.has_designer:
            main.open_designer(m)
        else:
            main.open_code(m, 0)
        return
    if s.kind == "control":
        m = s.mod.module.m
        d = main.open_designer(m)
        if d is not None and hasattr(d, "select_control"):
            ctl = next((c for c in m.sf.root.walk() if c.name.lower() == s.name.lower()), None)
            if ctl is not None:
                d.select_control(ctl)
        return
    if s.decls:
        main.last_positions = getattr(main, "last_positions", [])
        _, line, col = caret(main)
        main.last_positions.append((cw.module, line, col))
        ln, col = s.decls[0]
        select_range(main, s.mod.module.m, ln, col, len(s.name))


def last_position(main):
    st = getattr(main, "last_positions", [])
    if st:
        m, line, col = st.pop()
        select_range(main, m, line, col, 0)


def find_references(main):
    pr, s, cw = symbol_here(main)
    if s is None:
        main.statusBar().showMessage("Identifier under cursor is not recognized", 4000)
        return
    refs = pr.references(s)
    main.references.show_refs(s, refs)
    sure = sum(1 for r in refs if r.certain)
    main.bottom_tabs.setCurrentWidget(main.references)
    main.dock_bottom.show()
    main.statusBar().showMessage(f"{s.where()}: {sure} references" +
                                 (f", {len(refs) - sure} possible" if len(refs) > sure else ""), 6000)


def rename(main, new=None):
    pr, s, cw = symbol_here(main)
    if s is None:
        main.statusBar().showMessage("Identifier under cursor is not recognized", 4000)
        return False
    if s.kind == "proc":
        owner = R.event_proc_owner(s)
        if owner is not None and new is None:
            QMessageBox.information(main, "Rename", f"{s.name} is an event procedure of {owner}: rename {owner} "
                                                    f"to rename its event procedures.")
            return False
    if new is None:
        new, ok = QInputDialog.getText(main, "Rename", f"New name for {s.where()}:", QLineEdit.EchoMode.Normal, s.name)
        if not ok:
            return False
    new = new.strip()
    err = pr.conflicts(s, new)
    if err:
        QMessageBox.warning(main, "Rename", err)
        return False
    if new == s.name:
        return False
    return apply_rename(main, pr, s, new)


def apply_rename(main, pr, s, new):
    edits = R.rename_edits(pr, s, new)
    # outside the code: the module (its VB_Name, the .vbp), the control (the .frm)
    if s.kind == "module":
        m = s.mod.module.m
        err = main.rename_module(m, new)
        if err:
            QMessageBox.warning(main, "Rename", err)
            return False
    elif s.kind == "control":
        m = s.mod.module.m
        sw = main.design_windows.get(m)
        d = sw.widget() if sw is not None else None
        ctl = next((c for c in m.sf.root.walk() if c is not m.sf.root and c.name.lower() == s.name.lower()), None)
        if d is not None and hasattr(d, "rename") and ctl is not None:
            err = d.rename(ctl, new)
            if err:
                QMessageBox.warning(main, "Rename", err)
                return False
        else:
            for c in m.sf.root.walk():
                if c is not m.sf.root and c.name.lower() == s.name.lower():
                    c.name = new
                    c.raw_begin = None
            m.dirty = True
        for c in m.sf.root.walk():                  # DataSource = "Data1"
            for it in c.items:
                if getattr(it, "name", "").lower() == "datasource" and isinstance(getattr(it, "value", None), str) \
                        and it.value.lower() == s.name.lower():
                    it.value = new
    n = 0
    for mi, ed_list in edits.items():
        m = mi.module.m
        n += len(ed_list)
        sw = main.code_windows.get(m)
        if sw is not None:
            ed = sw.widget().editor
            ed.beginUndoAction()
            for line, col, length, text in sorted(set(ed_list), key=lambda e: (e[0], e[1]), reverse=True):
                ed.replace_range(_byte_pos(ed, line, col), _byte_pos(ed, line, col + length), text)
            ed.endUndoAction()
            m.set_code(sw.widget().text())
        else:
            m.set_code(R.apply_edits(module_text(main, m), ed_list))
        m.dirty = True
    for m, sw in main.code_windows.items():
        sw.widget().fill_objects()
    main.project.dirty = True
    main.update_tree_marks()
    main.fill_properties()
    main.statusBar().showMessage(f"Renamed {s.name} to {new}: {n} places", 6000)
    return True


def extract(main, name=None):
    cw = main.current_cw()
    if cw is None:
        return False
    ed = cw.editor
    a, b = ed.sel_range()
    l0, l1 = ed.line_from_pos(a), ed.line_from_pos(b)
    if b > a and b == ed.line_start(l1) and l1 > l0:
        l1 -= 1                                     # the selection ends at the start of the next line
    if name is None:
        name, ok = QInputDialog.getText(main, "Extract Procedure", "Name of the new Sub:", QLineEdit.EchoMode.Normal,
                                        "NewProcedure")
        if not ok:
            return False
    pr = analyze(main)
    try:
        text, at = R.extract_procedure(pr, cw.module.name, l0, l1, name.strip())
    except R.ExtractError as e:
        QMessageBox.warning(main, "Extract Procedure", str(e))
        return False
    if "\r\n" in cw.text():
        text = text.replace("\n", "\r\n")
    ed.beginUndoAction()
    ed.replace_range(0, ed.length(), text)
    ed.endUndoAction()
    cw.module.set_code(cw.text())
    cw.module.dirty = True
    cw.fill_objects()
    main.update_tree_marks()
    cw.goto_line(at)
    return True
