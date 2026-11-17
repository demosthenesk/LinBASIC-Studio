"""Refactoring of a VB6 project (Phase 5): Find References, Rename Symbol, Extract Procedure.

The code of every module is read as tokens with their places in the text (line, column); the
declarations make the symbols (variables, constants, procedures, parameters, types and their fields,
enums and their members, Declare, Event, the modules, the controls of the forms); every name of the
code is then bound as the compiler binds it: the locals of its procedure, the module, the public
names of the standard modules, the members of an object whose type the code says (Me, a form, a
variable As Class1 / As MyType, With ...). The event procedures name their object (Command1_Click,
mObj_Changed, IFoo_Method of Implements): renaming the object renames them. A member of an object
whose type is not known (a Variant, an Object) is a "possible" reference: listed, not renamed.

No Qt here: the IDE gives the modules (their code) and applies the edits."""
import re

from lbvc.lexer import KEYWORDS

_ID = re.compile(r'[A-Za-z_\u0080-￿][A-Za-z0-9_\u0080-￿]*')
_NUM = re.compile(r'(&[hH][0-9A-Fa-f]+&?|&[oO]?[0-7]+&?|(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?[%&!#@]?)')
_DATE = re.compile(r'#[^#\n]*[/:\-][^#\n]*#')
NAME_RE = re.compile(r'[A-Za-z][A-Za-z0-9_]{0,254}$')
SCOPE_WORDS = {"PUBLIC", "PRIVATE", "GLOBAL", "FRIEND", "DIM", "STATIC"}
EVENT_OWNERS = {"form", "mdiform", "usercontrol", "class", "propertypage", "userdocument", "dataenvironment",
                "datareport"}


class T:
    """a token: kind (id kw op str num nl), its text and its place (0-based line and column)"""
    __slots__ = ("kind", "text", "line", "col", "end")

    def __init__(self, kind, text, line, col, end):
        self.kind, self.text, self.line, self.col, self.end = kind, text, line, col, end

    @property
    def up(self):
        return self.text.upper()

    def __repr__(self):
        return f"T({self.kind} {self.text!r} {self.line}:{self.col})"


def tokens(text):
    """the tokens of the code with their places; a line continued with ' _' gives no 'nl'"""
    out = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    stmt_start = True
    for ln, s in enumerate(lines):
        p, n = 0, len(s)
        cont = False
        while p < n:
            ch = s[p]
            if ch in " \t":
                p += 1
                continue
            if ch == "'":
                break
            if ch == "_" and s[p + 1:].strip() == "" and p > 0 and s[p - 1] in " \t":
                cont = True
                break
            if ch == '"':
                q = p + 1
                while q < n:
                    if s[q] == '"':
                        if q + 1 < n and s[q + 1] == '"':
                            q += 2
                            continue
                        break
                    q += 1
                out.append(T("str", s[p:q + 1], ln, p, q + 1))
                p = q + 1
                stmt_start = False
                continue
            if ch == "[":
                q = s.find("]", p)
                q = n - 1 if q < 0 else q
                out.append(T("id", s[p + 1:q], ln, p + 1, q))
                p = q + 1
                stmt_start = False
                continue
            if ch == "#":
                m = _DATE.match(s, p)
                prev = out[-1] if out and out[-1].line == ln else None
                if m and not (prev and prev.kind in ("id", "kw") and prev.up in ("PRINT", "WRITE", "INPUT", "GET", "PUT",
                                                                                   "CLOSE", "AS", "SEEK", "LINE", "LOCK",
                                                                                   "UNLOCK", "WIDTH")):
                    out.append(T("num", m.group(0), ln, p, m.end()))
                    p = m.end()
                    stmt_start = False
                    continue
            if ch.isdigit() or ch == "&" and p + 1 < n and (s[p + 1] in "hHoO" or s[p + 1].isdigit()) or \
                    (ch == "." and p + 1 < n and s[p + 1].isdigit() and not (out and out[-1].kind == "id" and out[-1].end == p)):
                m = _NUM.match(s, p)
                if m and m.end() > p:
                    out.append(T("num", m.group(0), ln, p, m.end()))
                    p = m.end()
                    stmt_start = False
                    continue
            m = _ID.match(s, p)
            if m:
                word = m.group(0)
                e = m.end()
                if word.upper() == "REM" and stmt_start:
                    break
                prev = out[-1] if out else None
                after_dot = prev is not None and prev.kind == "op" and prev.text in (".", "!") and prev.end == p
                kind = "kw" if word.upper() in KEYWORDS and not after_dot else "id"
                out.append(T(kind, word, ln, p, e))
                if e < n and s[e] in "%&!#@$" and not (s[e] == "!" and e + 1 < n and (s[e + 1].isalpha() or s[e + 1] == "[")) \
                        and not (s[e] == "&" and e + 1 < n and s[e + 1] in "hHoO"):
                    e += 1                      # the type suffix stays outside the name
                p = e
                stmt_start = False
                continue
            two = s[p:p + 2]
            if two in ("<>", "<=", ">=", ":="):
                out.append(T("op", two, ln, p, p + 2))
                p += 2
                stmt_start = False
                continue
            out.append(T("op", ch, ln, p, p + 1))
            stmt_start = ch == ":"
            p += 1
        if not cont:
            out.append(T("nl", "\n", ln, n, n))
            stmt_start = True
    return out


def statements(toks):
    """the statements: lists of token indexes, split at the ends of lines and at ':' (not a label's)"""
    st, cur = [], []
    for i, t in enumerate(toks):
        if t.kind == "nl":
            if cur:
                st.append(cur)
            cur = []
        elif t.kind == "op" and t.text == ":" and not (len(cur) == 1 and toks[cur[0]].kind in ("id", "num")):
            if cur:
                st.append(cur)
            cur = []
        else:
            cur.append(i)
    if cur:
        st.append(cur)
    return st


class Sym:
    """a declared name"""

    def __init__(self, kind, name, mod, proc=None, scope="private", vtype=None, decl=None, parent=None):
        self.kind = kind            # var const proc param type field enum item declare event module control label
        self.name = name
        self.mod = mod              # its ModInfo (a module, the control's form ...)
        self.proc = proc            # its Proc for locals and parameters
        self.scope = scope          # local private public
        self.vtype = vtype          # the type its declaration says (As ...)
        self.decls = [decl] if decl else []     # (line, col) of the name in its declarations
        self.parent = parent        # the type of a field, the enum of a member
        self.withevents = False
        self.pkind = None           # Sub Function Property Get/Let/Set

    def where(self):
        if self.kind == "module":
            return self.name
        own = self.mod.name if self.mod is not None else ""
        if self.proc is not None:
            return f"{own}.{self.proc.name}.{self.name}"
        if self.parent is not None:
            return f"{own}.{self.parent.name}.{self.name}"
        return f"{own}.{self.name}"

    def __repr__(self):
        return f"Sym({self.kind} {self.where()})"


class Proc:
    def __init__(self, name, sym, first, last):
        self.name, self.sym = name, sym
        self.first, self.last = first, last     # token indexes of its header and of its End
        self.locals = {}                         # lower name -> Sym
        self.line0 = self.line1 = 0              # its lines (0-based, the End line included)


class Ref:
    """a place of a symbol in the code: the module, line, column, length; certain: bound for sure"""
    __slots__ = ("sym", "mod", "line", "col", "length", "certain", "decl")

    def __init__(self, sym, mod, line, col, length, certain=True, decl=False):
        self.sym, self.mod, self.line, self.col, self.length = sym, mod, line, col, length
        self.certain, self.decl = certain, decl

    def __repr__(self):
        return f"Ref({self.sym.name} {self.mod.name}:{self.line}:{self.col}{'' if self.certain else ' ?'})"


class ModInfo:
    """one module of the project: its tokens, its symbols, its procedures"""

    def __init__(self, module, text):
        self.module = module                     # the IDE's Module (or any object with name/kind)
        self.name = module.name
        self.kind = getattr(module, "kind", "Module")
        self.text = text
        self.toks = tokens(text)
        self.syms = {}                           # lower name -> Sym (module level)
        self.procs = []
        self.controls = {}                       # lower name -> Sym
        self.implements = []                     # the names after Implements
        self.sym = None                          # the module itself
        self.refs = []                           # the references found in its code

    @property
    def is_std(self):
        return self.kind == "Module"

    @property
    def lines(self):
        return self.text.replace("\r\n", "\n").split("\n")


def _type_after_as(toks, k, end):
    """the type name after As at k (As New X, As X.Y): (name, index after it)"""
    if k < end and toks[k].kind == "kw" and toks[k].up == "AS":
        k += 1
        if k < end and toks[k].kind == "kw" and toks[k].up == "NEW":
            k += 1
        if k < end:
            name = toks[k].text
            k += 1
            while k + 1 < end and toks[k].kind == "op" and toks[k].text == "." and toks[k + 1].kind in ("id", "kw"):
                name = toks[k + 1].text     # Lib.Type: the last part
                k += 2
            return name, k
    return None, k


def _skip_parens(toks, k, end):
    if k < end and toks[k].kind == "op" and toks[k].text == "(":
        d = 0
        while k < end:
            if toks[k].kind == "op" and toks[k].text == "(":
                d += 1
            elif toks[k].kind == "op" and toks[k].text == ")":
                d -= 1
                if d == 0:
                    return k + 1
            k += 1
    return k


class Project:
    """the analysis of a project: modules -> symbols -> references"""

    def __init__(self, modules):
        """modules: [(module, code text)]; a module has .name, .kind and, for a form, .controls_list
        (names of its controls)"""
        self.mods = [ModInfo(m, t) for m, t in modules]
        self.by_name = {mi.name.lower(): mi for mi in self.mods}
        self.globals = {}                        # lower -> [Sym]: modules, public names of .bas, enum items
        for mi in self.mods:
            mi.sym = Sym("module", mi.name, mi, scope="public")
            self.globals.setdefault(mi.name.lower(), []).append(mi.sym)
        for mi in self.mods:
            self._declare(mi)
        for mi in self.mods:
            for s in mi.syms.values():
                public = s.scope == "public" and (mi.is_std or s.kind in ("type", "enum", "item") or
                                                  (s.kind == "declare"))
                if s.kind == "item" and s.parent is not None and s.parent.scope == "public":
                    public = True
                if public:
                    self.globals.setdefault(s.name.lower(), []).append(s)
        for mi in self.mods:
            self._bind(mi)

    # ----------------------------------------------------------------- declarations
    def _add(self, table, s):
        old = table.get(s.name.lower())
        if old is not None and old.kind == s.kind == "proc":   # Property Get / Let / Set
            old.decls += s.decls
            return old
        table.setdefault(s.name.lower(), s)
        return table[s.name.lower()]

    def _declare(self, mi):
        toks = mi.toks
        for c in getattr(mi.module, "controls_list", []) or []:
            mi.controls.setdefault(c.lower(), Sym("control", c, mi, scope="public"))
        sts = statements(toks)
        block = None                              # ('type'|'enum', Sym) while inside one
        proc = None
        for st in sts:
            words = [toks[i].up if toks[i].kind in ("kw", "id") else toks[i].text for i in st]
            if block is not None:
                if len(words) >= 2 and words[0] == "END" and words[1] in ("TYPE", "ENUM"):
                    block = None
                    continue
                kind, parent = block
                t = toks[st[0]]
                if t.kind in ("id", "kw"):
                    if kind == "type":
                        vt, _ = _type_after_as(toks, _skip_parens(toks, st[0] + 1, st[-1] + 1), st[-1] + 1)
                        s = Sym("field", t.text, mi, scope=parent.scope, vtype=vt, decl=(t.line, t.col), parent=parent)
                        parent.fields[t.text.lower()] = s
                    else:
                        s = Sym("item", t.text, mi, scope=parent.scope, vtype=parent.name, decl=(t.line, t.col), parent=parent)
                        parent.fields[t.text.lower()] = s
                        self._add(mi.syms, s)
                continue
            if proc is not None:
                if len(words) >= 2 and words[0] == "END" and words[1] in ("SUB", "FUNCTION", "PROPERTY"):
                    proc.last = st[-1]
                    proc.line1 = toks[st[-1]].line
                    proc = None
                    continue
                self._local_decl(mi, proc, st)
                continue
            # module level
            i = 0
            scope = None
            while i < len(st) and words[i] in SCOPE_WORDS - {"DIM"}:
                scope = words[i]
                i += 1
            if i >= len(st):
                continue
            w = words[i]
            if w in ("SUB", "FUNCTION", "PROPERTY"):
                pk = w.title()
                j = i + 1
                if w == "PROPERTY" and j < len(st):
                    pk += " " + words[j].title()
                    j += 1
                if j >= len(st):
                    continue
                nt = toks[st[j]]
                s = Sym("proc", nt.text, mi, scope="private" if scope == "PRIVATE" else "public", decl=(nt.line, nt.col))
                s.pkind = pk
                end = st[-1] + 1
                k2 = _skip_parens(toks, st[j] + 1, end)
                s.vtype, _ = _type_after_as(toks, k2, end)
                s = self._add(mi.syms, s)
                proc = Proc(nt.text, s, st[0], None)
                proc.line0 = toks[st[0]].line
                mi.procs.append(proc)
                self._params(mi, proc, st[j] + 1, end)
            elif w == "DECLARE":
                j = i + 1
                if j < len(st) and words[j] in ("SUB", "FUNCTION"):
                    j += 1
                if j < len(st):
                    nt = toks[st[j]]
                    s = Sym("declare", nt.text, mi, scope="private" if scope == "PRIVATE" else "public",
                            decl=(nt.line, nt.col))
                    k2 = st[j] + 1
                    while k2 <= st[-1] and not (toks[k2].kind == "op" and toks[k2].text == "("):
                        k2 += 1
                    k2 = _skip_parens(toks, k2, st[-1] + 1)
                    s.vtype, _ = _type_after_as(toks, k2, st[-1] + 1)
                    self._add(mi.syms, s)
            elif w == "EVENT" and i + 1 < len(st):
                nt = toks[st[i + 1]]
                self._add(mi.syms, Sym("event", nt.text, mi, scope="public", decl=(nt.line, nt.col)))
            elif w in ("TYPE", "ENUM") and i + 1 < len(st):
                nt = toks[st[i + 1]]
                s = Sym(w.lower(), nt.text, mi, scope="private" if scope == "PRIVATE" else "public",
                        decl=(nt.line, nt.col))
                s.fields = {}
                s = self._add(mi.syms, s)
                block = (w.lower(), s)
            elif w == "IMPLEMENTS" and i + 1 < len(st):
                mi.implements.append(toks[st[i + 1]].text)
            elif w == "CONST":
                self._consts(mi, None, st, i + 1, "private" if scope in (None, "PRIVATE") else "public")
            elif w == "DIM" or (scope is not None and w not in ("SUB", "FUNCTION", "PROPERTY")):
                j = i + 1 if w == "DIM" else i
                sc = "public" if scope in ("PUBLIC", "GLOBAL") else "private"
                self._vars(mi, None, st, j, sc)

    def _vars(self, mi, proc, st, j, scope):
        """name[(..)] [As [New] T], ... from index j of the statement"""
        toks = mi.toks
        end = st[-1] + 1
        k = st[j] if j < len(st) else end
        withev = False
        while k < end:
            t = toks[k]
            if t.kind == "kw" and t.up == "WITHEVENTS":
                withev = True
                k += 1
                continue
            if t.kind == "kw" and t.up in ("SHARED", "PRESERVE"):
                k += 1
                continue
            if t.kind not in ("id", "kw"):
                k += 1
                continue
            name = t
            k = _skip_parens(toks, k + 1, end)
            vt, k = _type_after_as(toks, k, end)
            if proc is not None:
                if name.text.lower() not in proc.locals:
                    proc.locals[name.text.lower()] = Sym("var", name.text, mi, proc, "local", vt, (name.line, name.col))
            else:
                s = Sym("var", name.text, mi, scope=scope, vtype=vt, decl=(name.line, name.col))
                s.withevents = withev
                self._add(mi.syms, s)
            # to the next ','
            d = 0
            while k < end:
                if toks[k].kind == "op" and toks[k].text == "(":
                    d += 1
                elif toks[k].kind == "op" and toks[k].text == ")":
                    d -= 1
                elif toks[k].kind == "op" and toks[k].text == "," and d == 0:
                    k += 1
                    break
                k += 1

    def _consts(self, mi, proc, st, j, scope):
        toks = mi.toks
        end = st[-1] + 1
        k = st[j] if j < len(st) else end
        expect = True
        d = 0
        while k < end:
            t = toks[k]
            if expect and t.kind in ("id", "kw"):
                vt, k2 = _type_after_as(toks, k + 1, end)
                s = Sym("const", t.text, mi, proc, "local" if proc else scope, vt, (t.line, t.col))
                if proc is not None:
                    proc.locals.setdefault(t.text.lower(), s)
                else:
                    self._add(mi.syms, s)
                expect = False
                k = k2
                continue
            if t.kind == "op" and t.text == "(":
                d += 1
            elif t.kind == "op" and t.text == ")":
                d -= 1
            elif t.kind == "op" and t.text == "," and d == 0:
                expect = True
            k += 1

    def _params(self, mi, proc, k, end):
        toks = mi.toks
        if not (k < end and toks[k].kind == "op" and toks[k].text == "("):
            return
        close = _skip_parens(toks, k, end) - 1
        k += 1
        expect = True
        d = 0
        while k < close:
            t = toks[k]
            if expect and t.kind == "kw" and t.up in ("OPTIONAL", "BYVAL", "BYREF", "PARAMARRAY"):
                k += 1
                continue
            if expect and t.kind in ("id", "kw"):
                k2 = _skip_parens(toks, k + 1, close)
                vt, k2 = _type_after_as(toks, k2, close)
                proc.locals[t.text.lower()] = Sym("param", t.text, mi, proc, "local", vt, (t.line, t.col))
                expect = False
                k = k2
                continue
            if t.kind == "op" and t.text == "(":
                d += 1
            elif t.kind == "op" and t.text == ")":
                d -= 1
            elif t.kind == "op" and t.text == "," and d == 0:
                expect = True
            k += 1

    def _local_decl(self, mi, proc, st):
        toks = mi.toks
        a = toks[st[0]]
        if a.kind == "id" and len(st) == 2 and toks[st[1]].kind == "op" and toks[st[1]].text == ":":
            proc.locals.setdefault(a.text.lower(), Sym("label", a.text, mi, proc, "local", None, (a.line, a.col)))
            return
        if a.kind != "kw":
            return
        if a.up in ("DIM", "STATIC"):
            self._vars(mi, proc, st, 1, "local")
        elif a.up == "CONST":
            self._consts(mi, proc, st, 1, "local")
        elif a.up == "REDIM":
            j = 1
            if j < len(st) and toks[st[j]].kind == "kw" and toks[st[j]].up == "PRESERVE":
                j += 1
            if j < len(st):
                t = toks[st[j]]
                if self._lookup_plain(mi, proc, t.text) is None:     # ReDim of a new name declares it
                    self._vars(mi, proc, st, j, "local")

    # ----------------------------------------------------------------- binding
    def _lookup_plain(self, mi, proc, name):
        low = name.lower()
        if proc is not None and low in proc.locals:
            return proc.locals[low]
        if low in mi.syms:
            return mi.syms[low]
        if low in mi.controls:
            return mi.controls[low]
        for s in self.globals.get(low, []):
            return s
        return None

    def module_of_type(self, tname):
        """a type name -> the module (a class, a form) or the Type of a UDT"""
        if not tname:
            return None
        low = tname.lower()
        if low in self.by_name:
            return self.by_name[low]
        for s in self.globals.get(low, []):
            if s.kind == "type":
                return s
        return None

    def _type_target(self, s, mi=None):
        """what the members of a symbol are: a ModInfo, a Type Sym, or None (not known)"""
        if s is None:
            return None
        if s.kind == "module":
            return s.mod
        if s.kind in ("var", "param", "const", "field", "proc", "declare"):
            t = self.module_of_type(s.vtype)
            if t is None and s.vtype and mi is not None:    # a Private Type of the module
                ts = mi.syms.get(s.vtype.lower())
                if ts is not None and ts.kind == "type":
                    return ts
            return t
        return None

    def _member(self, target, name):
        """the member of a module / a Type named name"""
        low = name.lower()
        if isinstance(target, ModInfo):
            s = target.syms.get(low)
            if s is not None and s.kind not in ("type", "enum"):
                return s
            return target.controls.get(low)
        if isinstance(target, Sym) and target.kind == "type":
            return target.fields.get(low)
        return None

    def _bind(self, mi):
        toks = mi.toks
        self._bind_toks = toks                    # the base expressions of members read them
        proc_at = {}
        for p in mi.procs:
            proc_at[p.first] = p
        cur = None
        with_stack = []
        sts = statements(toks)
        # the event procedures: Object_Event, Interface_Member
        for p in mi.procs:
            self._event_proc(mi, p)
        for st in sts:
            if st[0] in proc_at:
                cur = proc_at[st[0]]
            words = [toks[i].up for i in st[:2]]
            if cur is not None and cur.last is not None and st[0] > cur.last:
                cur = None
            if words[0] == "WITH" and len(st) > 1:
                with_stack.append(self._expr_target(mi, cur, st[1:], with_stack))
            elif words[:2] == ["END", "WITH"] and with_stack:
                with_stack.pop()
            for n, i in enumerate(st):
                t = toks[i]
                if t.kind != "id" and not (t.kind == "kw" and n > 0 and toks[st[n - 1]].kind == "op"
                                           and toks[st[n - 1]].text in (".", "!")):
                    continue
                prev = toks[st[n - 1]] if n > 0 else None
                if cur is not None and i == cur.first:
                    continue
                if prev is not None and prev.kind == "op" and prev.text in (".", "!") and prev.end == t.col \
                        and prev.line == t.line:
                    if prev.text == "!":
                        continue                       # rs!Field: a field of a recordset
                    # the base: the tokens before the '.', or the With
                    base = self._base_tokens(st, n - 1)
                    if base:
                        target = self._expr_target(mi, cur, base, with_stack)
                    else:
                        target = with_stack[-1] if with_stack else None
                    s = self._member(target, t.text) if target is not None else None
                    if s is not None:
                        mi.refs.append(Ref(s, mi, t.line, t.col, len(t.text), True, (t.line, t.col) in s.decls))
                    else:
                        mi.refs.append(Ref(None, mi, t.line, t.col, len(t.text), False))
                        mi.refs[-1].sym = t.text.lower()        # by name only
                    continue
                s = self._lookup_plain(mi, cur, t.text)
                if s is not None:
                    mi.refs.append(Ref(s, mi, t.line, t.col, len(t.text), True, (t.line, t.col) in s.decls))
        # the declarations of the module level that the statements did not reach (Type fields)
        seen = {(r.line, r.col) for r in mi.refs}
        for table in [mi.syms] + [s.fields for s in mi.syms.values() if hasattr(s, "fields")]:
            for s in table.values():
                for (ln, col) in s.decls:
                    if (ln, col) not in seen:
                        mi.refs.append(Ref(s, mi, ln, col, len(s.name), True, True))
                        seen.add((ln, col))

    def _base_tokens(self, st, dot):
        """the token indexes of the expression before the '.' at position dot of the statement"""
        out = []
        k = dot - 1
        mi_toks = self._bind_toks
        while k >= 0:
            t = mi_toks[st[k]]
            if t.kind == "op" and t.text == ")":
                d = 0
                while k >= 0:
                    tt = mi_toks[st[k]]
                    if tt.kind == "op" and tt.text == ")":
                        d += 1
                    elif tt.kind == "op" and tt.text == "(":
                        d -= 1
                    out.insert(0, st[k])
                    k -= 1
                    if d == 0:
                        break
                continue
            if t.kind in ("id", "kw") and (t.kind == "id" or t.up in ("ME",) or
                                          (k > 0 and mi_toks[st[k - 1]].kind == "op" and mi_toks[st[k - 1]].text == ".")):
                out.insert(0, st[k])
                k -= 1
                if k >= 0 and mi_toks[st[k]].kind == "op" and mi_toks[st[k]].text in (".", "!") and \
                        mi_toks[st[k]].end == mi_toks[st[k + 1]].col:
                    out.insert(0, st[k])
                    k -= 1
                    continue
                break
            break
        return out

    def _expr_target(self, mi, proc, idx, with_stack):
        """the members' owner of a simple expression a.b(1).c / Me / .x (With)"""
        toks = mi.toks
        target = None
        k = 0
        first = True
        while k < len(idx):
            t = toks[idx[k]]
            if t.kind == "op" and t.text == ".":
                if first:
                    target = with_stack[-1] if with_stack else None
                first = False
                k += 1
                continue
            if t.kind == "op" and t.text == "(":
                d = 0
                while k < len(idx):
                    tt = toks[idx[k]]
                    if tt.kind == "op" and tt.text == "(":
                        d += 1
                    elif tt.kind == "op" and tt.text == ")":
                        d -= 1
                        if d == 0:
                            break
                    k += 1
                k += 1
                continue
            if t.kind == "kw" and t.up == "NEW":
                k += 1
                continue
            if t.kind in ("id", "kw"):
                if first:
                    if t.up == "ME":
                        target = mi
                    else:
                        target = self._type_target(self._lookup_plain(mi, proc, t.text), mi)
                    first = False
                else:
                    if target is None:
                        return None
                    target = self._type_target(self._member(target, t.text), mi)
                k += 1
                continue
            break
        return target

    def _event_proc(self, mi, p):
        """Command1_Click, mObj_Changed, IFoo_Method: the object (and the event / member) of the name"""
        name = p.name
        if "_" not in name or not p.sym.decls:
            return
        ln, col = p.sym.decls[0]
        for cut in [i for i, ch in enumerate(name) if ch == "_"]:
            obj, ev = name[:cut], name[cut + 1:]
            low = obj.lower()
            if low in EVENT_OWNERS:
                return
            s = mi.controls.get(low)
            if s is None:
                v = mi.syms.get(low)
                if v is not None and v.kind == "var" and v.withevents:
                    s = v
            if s is not None:
                mi.refs.append(Ref(s, mi, ln, col, len(obj), True, False))
                if s.kind == "var":                     # the event of the class
                    cls = self.module_of_type(s.vtype)
                    if isinstance(cls, ModInfo):
                        e = cls.syms.get(ev.lower())
                        if e is not None and e.kind == "event":
                            mi.refs.append(Ref(e, mi, ln, col + cut + 1, len(ev), True, False))
                return
            if any(i.lower() == low for i in mi.implements):
                iface = self.by_name.get(low)
                if iface is not None:
                    mi.refs.append(Ref(iface.sym, mi, ln, col, len(obj), True, False))
                    m = iface.syms.get(ev.lower())
                    if m is not None:
                        mi.refs.append(Ref(m, mi, ln, col + cut + 1, len(ev), True, False))
                return

    # ----------------------------------------------------------------- queries
    def module(self, name):
        return self.by_name.get(name.lower())

    def symbol_at(self, mod_name, line, col):
        """the symbol whose name is at (line, col) of a module (0-based)"""
        mi = self.module(mod_name)
        if mi is None:
            return None
        best = None
        for r in mi.refs:
            if r.line == line and r.col <= col <= r.col + r.length:
                if isinstance(r.sym, Sym):
                    if best is None or r.length < best.length:
                        best = r
        if best is not None:
            return best.sym
        # a name nothing binds (a member of a Variant ...): nothing
        return None

    def references(self, sym, possible=True):
        """[Ref] of a symbol in every module: certain ones, and members of unknown objects named the same"""
        out = []
        member_like = sym.kind in ("proc", "var", "const", "field", "event", "control") and sym.proc is None \
            and not (sym.mod is not None and sym.mod.is_std and sym.kind != "field")
        for mi in self.mods:
            for r in mi.refs:
                if r.sym is sym:
                    out.append(r)
                elif possible and member_like and isinstance(r.sym, str) and r.sym == sym.name.lower():
                    out.append(Ref(sym, mi, r.line, r.col, r.length, False))
        out.sort(key=lambda r: (r.mod.name.lower(), r.line, r.col))
        return out

    def conflicts(self, sym, new):
        """why sym cannot be named new (None: it can)"""
        if not NAME_RE.match(new or ""):
            return "Invalid name: it must start with a letter, then letters, digits or _"
        if new.upper() in KEYWORDS:
            return f"'{new}' is a reserved word"
        low = new.lower()
        if low == sym.name.lower():
            return None
        mi = sym.mod
        if sym.kind in ("param", "var", "const", "label") and sym.proc is not None:
            if low in sym.proc.locals:
                return f"'{new}' is already declared in {sym.proc.name}"
            return None
        if sym.kind == "module":
            if low in self.by_name:
                return f"Name '{new}' conflicts with existing module, project, or object library"
            return None
        if sym.kind == "field":
            if low in sym.parent.fields:
                return f"'{new}' is already a member of {sym.parent.name}"
            return None
        if sym.kind == "control":
            if low in mi.controls or low in mi.syms:
                return f"'{new}' is already the name of a control or a member of {mi.name}"
            return None
        if low in mi.syms or low in mi.controls:
            return f"'{new}' is already declared in {mi.name}"
        if sym.scope == "public" and mi.is_std:
            for s in self.globals.get(low, []):
                if s is not sym:
                    return f"'{new}' is already a public name ({s.where()})"
        return None


def event_proc_owner(sym):
    """a procedure named as an event procedure of a control / WithEvents variable of its module: the
    object's symbol (renaming the procedure alone breaks the event)"""
    if sym.kind != "proc" or "_" not in sym.name:
        return None
    mi = sym.mod
    for cut in [i for i, ch in enumerate(sym.name) if ch == "_"]:
        low = sym.name[:cut].lower()
        if low in EVENT_OWNERS:
            return sym.name[:cut]
        if low in mi.controls:
            return mi.controls[low].name
        v = mi.syms.get(low)
        if v is not None and v.kind == "var" and v.withevents:
            return v.name
    return None


def rename_edits(project, sym, new):
    """{ModInfo: [(line, col, length, text)]} for renaming sym to new (the certain references)"""
    edits = {}
    for r in project.references(sym, possible=False):
        edits.setdefault(r.mod, []).append((r.line, r.col, r.length, new))
    return edits


def apply_edits(text, edits):
    """the text with the edits (line, col, length, new text) made, from the end"""
    lines = text.replace("\r\n", "\n").split("\n")
    for ln, col, length, new in sorted(set(edits), key=lambda e: (e[0], e[1]), reverse=True):
        s = lines[ln]
        lines[ln] = s[:col] + new + s[col + length:]
    return "\n".join(lines)


# ====================================================================== Extract Procedure
class ExtractError(Exception):
    pass


def extract_procedure(project, mod_name, line0, line1, name):
    """the lines line0..line1 (0-based, whole lines) of a procedure become a new Private Sub name with
    the locals they use as its parameters (ByRef: what it changes reaches the caller, as before); its
    call takes their place. -> the new text of the module"""
    mi = project.module(mod_name)
    if mi is None:
        raise ExtractError("No such module")
    if not NAME_RE.match(name or "") or name.upper() in KEYWORDS:
        raise ExtractError("Invalid procedure name")
    if name.lower() in mi.syms or name.lower() in mi.controls:
        raise ExtractError(f"'{name}' is already declared in {mi.name}")
    proc = next((p for p in mi.procs if p.line0 < line0 and line1 < p.line1), None)
    if proc is None:
        raise ExtractError("Select whole lines inside a procedure (not its first or last line)")
    toks = [t for t in mi.toks if line0 <= t.line <= line1]
    # the control flow must stay inside
    for i, t in enumerate(toks):
        if t.kind == "kw" and t.up == "EXIT" and i + 1 < len(toks) and toks[i + 1].up in ("SUB", "FUNCTION", "PROPERTY"):
            raise ExtractError("The selection leaves the procedure (Exit Sub/Function)")
        if t.kind == "kw" and t.up in ("GOTO", "GOSUB", "RETURN", "RESUME"):
            raise ExtractError("The selection jumps (GoTo, GoSub, Resume): it cannot be moved")
        if t.kind == "id" and t.text.lower() == proc.name.lower() and proc.sym.pkind in ("Function", "Property Get"):
            raise ExtractError(f"The selection sets the value of {proc.name}")
    depth = 0
    for st in statements(toks):
        a = toks[st[0]].up if toks[st[0]].kind in ("kw", "id") else ""
        b = toks[st[1]].up if len(st) > 1 and toks[st[1]].kind in ("kw", "id") else ""
        if a in ("FOR", "DO", "WHILE", "SELECT", "WITH"):
            depth += 1
        elif a == "IF" and toks[st[-1]].up == "THEN":
            depth += 1
        elif a in ("NEXT", "LOOP", "WEND") or (a == "END" and b in ("IF", "SELECT", "WITH")):
            depth -= 1
        if depth < 0:
            raise ExtractError("The selection is not a whole block (an End/Next/Loop without its start)")
    if depth != 0:
        raise ExtractError("The selection is not a whole block (a start without its End/Next/Loop)")
    if any(r.line >= line0 and r.line <= line1 and isinstance(r.sym, Sym) and r.sym.kind == "label" for r in mi.refs):
        raise ExtractError("The selection has labels")
    # the locals it uses
    used, inside = [], set()
    for r in mi.refs:
        s = r.sym
        if not isinstance(s, Sym) or s.proc is not proc or s.kind not in ("var", "param", "const"):
            continue
        if line0 <= r.line <= line1:
            if s.decls and line0 <= s.decls[0][0] <= line1:
                inside.add(s)
            elif s not in used:
                used.append(s)
    for s in inside:                            # declared in it: not used after it
        if any(r.sym is s and r.line > line1 for r in mi.refs):
            raise ExtractError(f"'{s.name}' is declared in the selection and used after it")
    used.sort(key=lambda s: s.decls[0] if s.decls else (0, 0))
    params = []
    for s in used:
        t = s.vtype or "Variant"
        arr = False
        if s.decls:
            ln, col = s.decls[0]
            rest = mi.lines[ln][col + len(s.name):]
            arr = rest.lstrip().startswith("(")
        if s.kind == "const":
            params.append(f"ByVal {s.name} As {t}")
        else:
            params.append(f"{s.name}{'()' if arr else ''} As {t}")
    lines = mi.lines
    body = lines[line0:line1 + 1]
    indent = re.match(r"[ \t]*", body[0]).group(0) if body else "    "
    # the body one level in, as a new Sub
    strip = min((len(re.match(r"[ \t]*", ln_).group(0)) for ln_ in body if ln_.strip()), default=0)
    new_body = ["    " + ln_[strip:] if ln_.strip() else "" for ln_ in body]
    call = indent + name + ((" " + ", ".join(s.name for s in used)) if used else "")
    sub = ["", f"Private Sub {name}({', '.join(params)})"] + new_body + ["End Sub"]
    out = lines[:line0] + [call] + lines[line1 + 1:proc.line1 + 1] + sub + lines[proc.line1 + 1:]
    shift = len(body) - 1
    return "\n".join(out), proc.line1 - shift + 2      # the text, the line of the new Sub
