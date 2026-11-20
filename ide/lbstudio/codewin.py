"""The code window of a module: the Object and Procedure boxes of VB6 above the editor of
Linpad++ (Scintilla with the VB lexer)."""

import re

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QComboBox, QSplitter
from PyQt6 import sip
from PyQt6.Qsci import QsciScintilla as S

from .editor import Editor
from . import vbmodel, vbtypes

GENERAL = "(General)"
DECLARATIONS = "(Declarations)"


class CodeWindow(QWidget):
    modified_changed = pyqtSignal(object, bool)

    def __init__(self, main, module):
        super().__init__()
        self.main = main
        self.module = module
        self._filling = False
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(2)
        top = QHBoxLayout()
        top.setContentsMargins(2, 2, 2, 0)
        self.obj_box = QComboBox()
        self.proc_box = QComboBox()
        for b in (self.obj_box, self.proc_box):
            b.setMinimumWidth(180)
            b.setMaxVisibleItems(24)
            top.addWidget(b, 1)
        lay.addLayout(top)
        # Window > Split: two panes of the same text, one over the other (DEV-8/me2 pictures 5, 6)
        self.splitter = QSplitter(Qt.Orientation.Vertical)
        self.splitter.setHandleWidth(5)
        self.splitter.splitterMoved.connect(self._splitter_moved)
        lay.addWidget(self.splitter, 1)
        self.editor = Editor(main, 0, self.splitter)          # the pane with the caret (the one in use)
        self.first = self.editor                               # the pane that stays when the split goes
        self.second = None
        self.splitter.addWidget(self.editor)
        main.setup_code_editor(self.editor)
        self.set_text(module.code_text())
        self.obj_box.activated.connect(self._object_chosen)
        self.proc_box.activated.connect(self._proc_chosen)
        self._hook_pane(self.editor)
        self.fill_objects()
        self.assist = CodeAssist(self)

    # ---------------------------------------------------------------- Split
    def panes(self):
        return [e for e in (self.first, self.second) if e is not None]

    def _hook_pane(self, ed):
        ed.cursorPositionChanged.connect(lambda *_, e=ed: self._follow_caret() if e is self.editor else None)
        ed.focused.connect(self._pane_focused)

    def _pane_focused(self, ed):
        if ed is not self.editor and ed in self.panes():
            self.editor = ed
            self._follow_caret()

    def is_split(self):
        return self.second is not None

    def set_split(self, on):
        """Window > Split: the window in two panes of the same text; again, back to one pane"""
        if on == self.is_split():
            return
        if on:
            ed = Editor(self.main, 1, self.splitter)
            ed.setDocument(self.first.document())              # the same document: one text, one Undo
            self.main.setup_code_editor(ed)
            self.assist.setup_view(ed)
            self.splitter.insertWidget(0, ed)
            self.second = ed
            self._hook_pane(ed)
            h = max(2, self.splitter.height())
            self.splitter.setSizes([h // 2, h - h // 2])
            line = self.first.SendScintilla(S.SCI_GETFIRSTVISIBLELINE)
            ed.SendScintilla(S.SCI_SETFIRSTVISIBLELINE, line)
            ed.SendScintilla(S.SCI_GOTOPOS, self.first.pos())
            self.first.ensureCursorVisible()
        else:
            gone = self.second
            if self.editor is gone:                   # the place of the pane in use goes to the one that stays
                self.first.SendScintilla(S.SCI_SETFIRSTVISIBLELINE, gone.SendScintilla(S.SCI_GETFIRSTVISIBLELINE))
                self.first.set_sel(gone.SendScintilla(S.SCI_GETANCHOR), gone.pos())
            self.editor = self.first
            self.second = None
            # deleted now: with deleteLater the pane would wait (without a parent) for the event loop, and its
            # timer could still call a main window that is gone (the crash of t_theme)
            gone._ui_timer.stop()
            gone.hide()
            sip.delete(gone)
            self.first.setFocus()

    def _splitter_moved(self, pos, index):
        """the bar dragged to the top or the bottom: one pane again, as in VB6"""
        if self.second is not None and 0 in self.splitter.sizes():
            if self.splitter.sizes()[1] == 0:          # the first pane is gone from sight: keep the second's place
                self.first.SendScintilla(S.SCI_SETFIRSTVISIBLELINE,
                                         self.second.SendScintilla(S.SCI_GETFIRSTVISIBLELINE))
                self.first.set_sel(self.second.SendScintilla(S.SCI_GETANCHOR), self.second.pos())
            self.editor = self.first
            self.set_split(False)
            self.splitter.setSizes([max(1, self.splitter.height())])

    # ----------------------------------------------------------------- text
    def set_text(self, text):
        ed = self.editor
        data = text.replace("\r\n", "\n").encode("utf-8")
        # a code window opened while the program runs or waits in break mode is read-only (Edit and Continue): the
        # text goes in all the same (it was empty, and the "reset your project" question came, DEV-9 df3)
        ro = ed.SendScintilla(S.SCI_GETREADONLY)
        if ro:
            ed.SendScintilla(S.SCI_SETREADONLY, False)
        ed.SendScintilla(S.SCI_SETUNDOCOLLECTION, False)
        ed.SendScintilla(S.SCI_CLEARALL)
        if data:
            ed.SendScintilla(S.SCI_APPENDTEXT, len(data), data)
        if ro:
            ed.SendScintilla(S.SCI_SETREADONLY, True)
        ed.SendScintilla(S.SCI_SETUNDOCOLLECTION, True)
        ed.SendScintilla(S.SCI_EMPTYUNDOBUFFER)
        ed.SendScintilla(S.SCI_SETSAVEPOINT)
        ed.SendScintilla(S.SCI_GOTOPOS, 0)

    def text(self):
        return self.editor.all_text()

    def mark_saved(self):
        self.editor.SendScintilla(S.SCI_SETSAVEPOINT)

    def goto_line(self, line, col=0):
        """line: 0-based line of the code (not of the file)"""
        ed = self.editor
        line = max(0, min(line, ed.line_count() - 1))
        ed.goto_line(line)
        p = ed.line_start(line) + max(0, col)
        e = ed.line_end(line)
        ed.set_sel(ed.line_start(line), e)     # VB6 selects the line with the error
        ed.setFocus()
        return p

    # -------------------------------------------------------------- the boxes
    def fill_objects(self):
        self._filling = True
        self.obj_box.clear()
        self.obj_box.addItem(GENERAL)
        m = self.module
        if m.is_form or m.kind in ("UserControl", "PropertyPage"):
            self.obj_box.addItem(m.kind if m.kind in ("UserControl", "PropertyPage") else
                                 "MDIForm" if m.display_kind == "MDIForm" else "Form")
            for name, tname, arr in m.controls():
                self.obj_box.addItem(name, (tname, arr))
        elif m.kind == "Class":
            self.obj_box.addItem("Class")
        elif m.kind == "Designer" and m.sf is not None and m.sf.root is not None:
            from . import dsrmodel
            k = dsrmodel.kind_of(m.sf)
            if k == "DataEnvironment":
                de = dsrmodel.DataEnvModel(m.sf)
                self.obj_box.addItem("DataEnvironment", ("DataEnvironment", False))
                for c in de.conns:
                    self.obj_box.addItem(c.name, ("ADODB.ConnectionEvents", False))
                for c in de.top():
                    self.obj_box.addItem("rs" + c.member, ("ADODB.RecordsetEvents", False))
            elif k == "DataReport":
                self.obj_box.addItem("DataReport", ("DataReport", False))
        self._filling = False
        self._fill_procs(GENERAL)
        self._follow_caret()

    def _object_type(self, obj):
        if obj in ("Form", "MDIForm", "UserControl", "PropertyPage"):
            return obj, False
        if obj == "Class":
            return "Class", False
        i = self.obj_box.findText(obj)
        d = self.obj_box.itemData(i) if i >= 0 else None
        return d if d else (None, False)

    def _existing(self):
        return {name.lower(): line for name, kind, line in vbmodel.procedures(self.text())}

    def _fill_procs(self, obj):
        self._filling = True
        self.proc_box.clear()
        have = self._existing()
        bold = QFont(self.proc_box.font())
        bold.setBold(True)
        if obj == GENERAL:
            self.proc_box.addItem(DECLARATIONS)
            objs = {self.obj_box.itemText(i).lower() for i in range(1, self.obj_box.count())}
            for name, kind, line in vbmodel.procedures(self.text()):
                head = name.split("_", 1)[0].lower()
                if "_" in name and head in objs:
                    continue                    # event procedures are under their object
                self.proc_box.addItem(name)
        else:
            tname, _ = self._object_type(obj)
            evs = [("Initialize", ""), ("Terminate", "")] if tname == "Class" else vbtypes.events(tname or "")
            if not evs and tname:
                evs = uctl_events(tname)
            for ev, params in sorted(evs, key=lambda e: e[0].lower()):
                self.proc_box.addItem(ev, params)
                key = f"{('Class' if tname == 'Class' else obj)}_{ev}".lower()
                if key in have:
                    self.proc_box.setItemData(self.proc_box.count() - 1, bold, Qt.ItemDataRole.FontRole)
        self._filling = False

    def _object_chosen(self, i):
        obj = self.obj_box.itemText(i)
        self._fill_procs(obj)
        if obj == GENERAL:
            self.editor.goto_line(0)
            self.editor.setFocus()
            return
        # VB6: choosing an object opens its default event (Click, Load ...)
        tname, _ = self._object_type(obj)
        default = {"Form": "Load", "MDIForm": "Load", "Class": "Initialize", "UserControl": "Initialize", "PropertyPage": "SelectionChanged", "Timer": "Timer",
                   "TextBox": "Change", "HScrollBar": "Change", "VScrollBar": "Change",
                   "DirListBox": "Change", "DriveListBox": "Change", "Image": "Click"}.get(tname, "Click")
        j = self.proc_box.findText(default)
        if j >= 0:
            self.proc_box.setCurrentIndex(j)
            self._proc_chosen(j)

    def _proc_chosen(self, i):
        if self._filling:
            return
        obj = self.obj_box.currentText()
        proc = self.proc_box.itemText(i)
        if obj == GENERAL:
            if proc == DECLARATIONS:
                self.editor.goto_line(0)
            else:
                line = self._existing().get(proc.lower())
                if line is not None:
                    self._goto_body(line)
            self.editor.setFocus()
            return
        tname, arr = self._object_type(obj)
        prefix = "Class" if tname == "Class" else obj
        name = f"{prefix}_{proc}"
        line = self._existing().get(name.lower())
        if line is None:
            params = self.proc_box.itemData(i) or ""
            if arr:
                params = "Index As Integer" + (", " + params if params else "")
            line = self._insert_proc(f"Private Sub {name}({params})")
        self._goto_body(line)
        self.editor.setFocus()

    def _insert_proc(self, header):
        """a new event procedure at the end of the module, as VB6 adds it"""
        ed = self.editor
        text = self.text()
        add = ("" if text.endswith("\n") or not text else "\n") + ("\n" if text.strip() else "") + header + "\n\nEnd Sub\n"
        ed.SendScintilla(S.SCI_APPENDTEXT, len(add.encode("utf-8")), add.encode("utf-8"))
        return ed.line_count() - 4 if ed.line_count() >= 4 else 0

    def _goto_body(self, header_line):
        ed = self.editor
        line = min(header_line + 1, ed.line_count() - 1)
        ed.goto_line(line)
        p = ed.line_end(line)
        ed.set_sel(p, p)

    def _follow_caret(self):
        """the boxes show the procedure that has the caret"""
        if self._filling:
            return
        line = self.editor.line_from_pos(self.editor.pos())
        cur = None
        procs = vbmodel.procedures(self.text())
        for name, kind, l0 in procs:
            if l0 <= line:
                cur = name
        if cur is None:
            obj, proc = GENERAL, DECLARATIONS
        else:
            head, _, rest = cur.partition("_")
            i = self.obj_box.findText(head, Qt.MatchFlag.MatchFixedString)
            if rest and i > 0:
                obj, proc = self.obj_box.itemText(i), rest
            elif rest and head.lower() in ("form", "mdiform", "class") and self.obj_box.findText(head, Qt.MatchFlag.MatchFixedString) > 0:
                obj, proc = head, rest
            else:
                obj, proc = GENERAL, cur
        if self.obj_box.currentText() != obj:
            self.obj_box.setCurrentIndex(max(0, self.obj_box.findText(obj, Qt.MatchFlag.MatchFixedString)))
            self._fill_procs(self.obj_box.currentText())
        j = self.proc_box.findText(proc, Qt.MatchFlag.MatchFixedString)
        self._filling = True
        if j >= 0:
            self.proc_box.setCurrentIndex(j)
        self._filling = False


# ---------------------------------------------------------------- IntelliSense in the code window
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPolygon  # noqa: E402
from PyQt6.QtCore import QPoint  # noqa: E402
from . import intellisense as IS  # noqa: E402

IMG = {"prop": 1, "method": 2, "event": 3, "const": 4, "enum": 5, "module": 6, "control": 7, "type": 8, "keyword": 9}


def _icon(kind):
    """small pictures of the lists, in the spirit of VB6 (they are drawn, not copied)"""
    pm = QPixmap(14, 14)
    pm.fill(QColor(0, 0, 0, 0))
    p = QPainter(pm)
    if kind == "prop":
        p.fillRect(3, 2, 8, 10, QColor("#FFFFFF"))
        p.setPen(QColor("#404040"))
        p.drawRect(3, 2, 8, 10)
        p.drawLine(5, 5, 9, 5)
        p.drawLine(5, 8, 9, 8)
    elif kind == "method":
        p.setBrush(QColor("#7A2E9A"))
        p.setPen(QColor("#4A1060"))
        p.drawPolygon(QPolygon([QPoint(7, 1), QPoint(13, 5), QPoint(7, 9), QPoint(1, 5)]))
        p.drawRect(1, 6, 0, 0)
    elif kind == "event":
        p.setBrush(QColor("#F0C000"))
        p.setPen(QColor("#806000"))
        p.drawPolygon(QPolygon([QPoint(8, 0), QPoint(3, 8), QPoint(7, 8), QPoint(5, 13), QPoint(11, 5), QPoint(7, 5)]))
    elif kind in ("const", "enum"):
        p.setBrush(QColor("#808080") if kind == "const" else QColor("#D0A000"))
        p.setPen(QColor("#404040"))
        p.drawRect(2, 4, 9, 6)
    elif kind in ("module", "type"):
        p.setBrush(QColor("#4070C0"))
        p.setPen(QColor("#203060"))
        p.drawRect(2, 2, 10, 10)
    elif kind == "control":
        p.setBrush(QColor("#E0E0E0"))
        p.setPen(QColor("#404040"))
        p.drawRect(1, 3, 12, 8)
    p.end()
    return pm


class CodeAssist:
    """IntelliSense of one code window"""

    def __init__(self, cw):
        self.cw = cw
        self.setup_view(cw.editor)
        self.tip = None                                            # the Member of the call tip

    def setup_view(self, ed):
        """the pictures and the list settings of a pane (each pane of a split window has its own)"""
        for kind, n in IMG.items():
            ed.registerImage(n, _icon(kind))
        ed.SendScintilla(S.SCI_AUTOCSETIGNORECASE, True)
        ed.SendScintilla(S.SCI_AUTOCSETCASEINSENSITIVEBEHAVIOUR, 1)
        ed.SendScintilla(S.SCI_AUTOCSETORDER, 1)                   # Scintilla sorts the list
        ed.SendScintilla(S.SCI_AUTOCSETFILLUPS, 0, b" .(=,)")
        ed.SendScintilla(S.SCI_AUTOCSETMAXHEIGHT, 9)
        ed.SendScintilla(S.SCI_AUTOCSETAUTOHIDE, False)

    @property
    def engine(self):
        return self.cw.main.intellisense()

    def _where(self):
        ed = self.cw.editor
        p = ed.pos()
        line = ed.line_from_pos(p)
        before = ed.get_range(ed.line_start(line), p)
        return ed, p, line, before

    def _show(self, items, typed=""):
        ed = self.cw.editor
        if not items:
            return False
        words = " ".join(f"{m.name}?{IMG.get(m.kind, 9)}" for m in items if " " not in m.name)
        data = words.encode("utf-8")
        ed.SendScintilla(S.SCI_AUTOCSHOW, len(typed.encode("utf-8")), data)
        return True

    # ------------------------------------------------------------ the commands of VB6
    def list_members(self, typed_only=False):
        """after a dot (or Ctrl+J): the properties and methods of the object"""
        ed, p, line, before = self._where()
        got = IS.chain_before(before)
        text = ed.all_text()
        if got is None:
            if typed_only:
                return False
            return self.complete_word(insert_single=False)
        chain, typed = got
        if chain == ["<with>"]:
            w = IS.with_object(text, line)
            if not w:
                return False
            chain = IS.split_chain(w)
        items = self.engine.list_members(chain, self.cw.module, line, text)
        return self._show(items, typed)

    def complete_word(self, insert_single=True):
        """Ctrl+Space: the names that begin with what is typed (one name is written at once)"""
        ed, p, line, before = self._where()
        m = re.search(r"[A-Za-z_][A-Za-z0-9_]*$", before)
        typed = m.group(0) if m else ""
        items = self.engine.complete_words(self.cw.module, line, ed.all_text())
        if typed:
            hits = [x for x in items if x.name.lower().startswith(typed.lower())]
            if len(hits) == 1 and insert_single:
                ed.replace_range(p - len(typed.encode("utf-8")), p, hits[0].name)
                ed.goto_pos(p - len(typed.encode("utf-8")) + len(hits[0].name.encode("utf-8")))
                return True
        return self._show(items, typed)

    def list_types(self):
        ed, p, line, before = self._where()
        m = re.search(r"[A-Za-z_][A-Za-z0-9_.]*$", before)
        typed = m.group(0) if m else ""
        return self._show(self.engine.type_list(), typed.split(".")[-1])

    def list_constants(self):
        """Ctrl+Shift+J, or by itself after "x = " and in an argument of an enumerated type"""
        ed, p, line, before = self._where()
        text = ed.all_text()
        m = re.search(r"[A-Za-z_][A-Za-z0-9_]*$", before)
        typed = m.group(0) if m else ""
        head = before[:len(before) - len(typed)]
        am = re.search(r"((?:[A-Za-z_][A-Za-z0-9_]*(?:\([^()]*\))?\.)*[A-Za-z_][A-Za-z0-9_]*)\s*(?:=|<>)\s*$", head)
        items = []
        if am:
            chain = IS.split_chain(am.group(1))
            eng = self.engine
            if len(chain) >= 2:
                typ = eng.resolve(chain[:-1], self.cw.module, line, text)
                items = eng.property_constants(typ, chain[-1])
                if not items:
                    t = eng.member_type(typ, chain[-1], self.cw.module) if typ else None
                    items = eng.constants_for(t[1] if t and t[0] == "lib" else None)
            else:
                loc = IS.local_vars(text, line)
                ci = eng.info(self.cw.module)
                tn = loc.get(chain[0].lower(), (None, None))[1] or (ci.vars.get(chain[0].lower()) or (None, None))[1]
                items = eng.constants_for(tn)
        else:
            ctx = IS.call_context(head)
            if ctx is not None:
                mem = self.engine.callable(ctx[0], self.cw.module, line, text)
                if mem is not None and ctx[1] < len(mem.params):
                    items = self.engine.constants_for(mem.params[ctx[1]].get("type"))
        return self._show(items, typed)

    def quick_info(self):
        """Ctrl+I: the declaration of the name at the caret"""
        ed, p, line, before = self._where()
        a, b = ed.word_at(p)
        word = ed.get_range(a, b)
        if not word:
            return False
        lb = ed.get_range(ed.line_start(line), a)
        got = IS.chain_before(lb + word)
        text = ed.all_text()
        chain = (got[0] + [word]) if got else [word]
        if chain and chain[0] == "<with>":
            w = IS.with_object(text, line)
            chain = (IS.split_chain(w) if w else []) + [word]
        mem = self.engine.callable(chain, self.cw.module, line, text)
        if mem is None:
            g = self.engine.globals().get(word.lower())
            mem = g
        if mem is None:
            return False
        self._calltip(a, mem.signature(), None)
        return True

    def parameter_info(self, quiet=False):
        """Ctrl+Shift+I, and by itself after ( , and a space after a procedure name"""
        ed, p, line, before = self._where()
        ctx = IS.call_context(before)
        if ctx is None:
            if not quiet:
                self.cancel_tip()
            return False
        chain, argi = ctx
        text = ed.all_text()
        if chain and chain[0] == "":
            return False
        mem = self.engine.callable(chain, self.cw.module, line, text)
        if mem is None or (not mem.params and mem.kind != "method"):
            self.cancel_tip()
            return False
        self.tip = mem
        sig = mem.signature()
        hl = None
        if argi < len(mem.params):
            ptxt = IS.p_text(mem.params[argi])
            start = sig.find(ptxt, len(mem.name))
            if start >= 0:
                hl = (start, start + len(ptxt))
        elif mem.params and mem.params[-1]["name"].lower() in ("outputlist", "paramarray"):
            hl = None
        self._calltip(p, sig, hl)
        return True

    def _calltip(self, pos, text, hl):
        ed = self.cw.editor
        data = text.encode("utf-8")
        same = ed.SendScintilla(S.SCI_CALLTIPACTIVE) and getattr(self, "_tip_text", None) == text
        if not same:
            if ed.SendScintilla(S.SCI_AUTOCACTIVE):
                return                      # a call tip would close the list of constants
            ed.SendScintilla(S.SCI_CALLTIPSHOW, pos, data)
            self._tip_text = text
        if not hl:
            ed.SendScintilla(S.SCI_CALLTIPSETHLT, 0, 0)
        if hl:
            a = len(text[:hl[0]].encode("utf-8"))
            b = len(text[:hl[1]].encode("utf-8"))
            ed.SendScintilla(S.SCI_CALLTIPSETHLT, a, b)

    def cancel_tip(self):
        self.tip = None
        self.cw.editor.SendScintilla(S.SCI_CALLTIPCANCEL)

    # ------------------------------------------------------------ typing
    def char_added(self, ch):
        """after Scintilla has finished with the key (a list shown inside the notification of
        a fill-up character would be completed by that same character)"""
        from PyQt6.QtCore import QTimer
        QTimer.singleShot(0, lambda: self._char_added(ch))

    def _char_added(self, ch):
        S_ = self.cw.main.S
        c = chr(ch) if isinstance(ch, int) else ch
        if not S_["auto_list_members"] and c in ". =":
            return
        if not S_["auto_quick_info"] and c in "(,)":
            return
        ed, p, line, before = self._where()
        if self._in_string_or_comment(before):
            return
        if c == ".":
            self.list_members(typed_only=True)
            return
        if c in "(,":
            self.parameter_info()
            return
        if c == ")":
            self.parameter_info(quiet=True) or self.cancel_tip()
            return
        if c == " ":
            low = before.lower()
            if re.search(r"\b(as|new)\s$", low) and not re.search(r"\bnew\s+as\s$", low):
                self.list_types()
                return
            if re.search(r"(=|<>)\s$", before):
                if self.list_constants():
                    return
            if self.parameter_info(quiet=True):
                ctx = IS.call_context(before)
                if ctx is not None:
                    self.list_constants()
            return

    def _in_string_or_comment(self, before):
        q = False
        for i, ch in enumerate(before):
            if ch == '"':
                q = not q
            elif ch == "'" and not q:
                return True
        return q or bool(re.match(r"^\s*Rem\s", before, re.I))

    def caret_moved(self):
        if self.tip is not None:
            ed, p, line, before = self._where()
            if IS.call_context(before) is None:
                self.cancel_tip()
            else:
                self.parameter_info(quiet=True)


def uctl_events(type_name):
    """the events of a user control on a form: its own (Event lines of the .ctl) and its Extender's"""
    import re
    from .designer import USER_CONTROLS
    m = next((e[4] for e in USER_CONTROLS.values() if e[1].lower() == type_name.lower()), None)
    if m is None:
        return []
    out = [("DragDrop", "Source As Control, X As Single, Y As Single"),
           ("DragOver", "Source As Control, X As Single, Y As Single, State As Integer"),
           ("GotFocus", ""), ("LostFocus", ""), ("Validate", "Cancel As Boolean")]
    for line in m.code_text().splitlines():
        ev = re.match(r"^\s*(?:Public\s+)?Event\s+([A-Za-z_]\w*)\s*(?:\((.*)\))?", line, re.IGNORECASE)
        if ev:
            out = [e for e in out if e[0].lower() != ev.group(1).lower()] + [(ev.group(1), (ev.group(2) or "").strip())]
    return out

