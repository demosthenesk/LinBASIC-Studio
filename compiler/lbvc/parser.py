# LinBASIC Studio - lbvc parser: tokens -> syntax tree
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
from .lexer import Tok, VBError, tokenize, SUFFIX_TYPE


class N:
    """Syntax tree node: kind, line, and named fields."""

    def __init__(self, kind, line=0, **kw):
        self.kind = kind
        self.line = line
        self.__dict__.update(kw)

    def __repr__(self):
        d = {k: v for k, v in self.__dict__.items() if k not in ('kind', 'line')}
        return f"N({self.kind}@{self.line} {d})"


TYPE_WORDS = {'INTEGER': 'Integer', 'LONG': 'Long', 'SINGLE': 'Single', 'DOUBLE': 'Double',
              'CURRENCY': 'Currency', 'BOOLEAN': 'Boolean', 'BYTE': 'Byte', 'STRING': 'String',
              'VARIANT': 'Variant', 'DATE': 'Date', 'OBJECT': 'Object', 'ANY': 'Any'}
DEFTYPES = {'DEFINT': 'Integer', 'DEFLNG': 'Long', 'DEFSNG': 'Single', 'DEFDBL': 'Double', 'DEFCUR': 'Currency',
            'DEFBOOL': 'Boolean', 'DEFBYTE': 'Byte', 'DEFSTR': 'String', 'DEFVAR': 'Variant', 'DEFDATE': 'Date',
            'DEFOBJ': 'Object'}
CMP_OPS = ('=', '<>', '<', '>', '<=', '>=')


# Compile On Demand (a run of the IDE, --debug): a syntax error inside a procedure (a block not closed, End Sub
# inside an If ...) does not stop the parse: the procedure gets the error (proc.parse_error) and becomes a stub that
# reports it when it is called, as VB6 compiles a procedure only then (the rest of the program runs)
RECOVER = False


class Parser:
    def __init__(self, text, filename, module_name):
        self.recover = RECOVER
        self.filename = filename
        self.toks = tokenize(text, filename)
        self.pos = 0
        self.module = N('module', 0, name=module_name, file=filename, items=[], options={}, attrs=[])

    # ---------------------------------------------------------------- helpers
    @property
    def t(self):
        return self.toks[self.pos]

    def peek(self, k=1):
        return self.toks[min(self.pos + k, len(self.toks) - 1)]

    def next(self):
        t = self.toks[self.pos]
        self.pos += 1
        return t

    def err(self, msg, tok=None):
        tok = tok or self.t
        raise VBError(msg, tok.line, tok.col, self.filename)

    def expect_op(self, op):
        if not self.t.is_op(op):
            self.err(f"expected: {op}")
        return self.next()

    def expect_kw(self, *kws):
        if not self.t.is_kw(*kws):
            self.err(f"expected: {' or '.join(k.title() for k in kws)}")
        return self.next()

    def accept_op(self, op):
        if self.t.is_op(op):
            self.pos += 1
            return True
        return False

    def accept_kw(self, kw):
        if self.t.is_kw(kw):
            self.pos += 1
            return True
        return False

    def accept_word(self, w):
        if self.t.is_word(w):
            self.pos += 1
            return True
        return False

    def at_end(self):
        return self.t.kind in ('nl', 'eof') or self.t.is_op(':') or self.t.is_kw('ELSE')

    def end_stmt(self):
        if self.t.is_op(':') or self.t.kind == 'nl':
            self.pos += 1
        elif self.t.kind == 'eof' or self.t.is_kw('ELSE'):
            pass
        else:
            self.err("expected: end of statement")

    def skip_nl(self):
        while self.t.kind == 'nl' or self.t.is_op(':'):
            self.pos += 1

    def member_word(self, what):
        """a name after a dot: any word (VB.Global, Database.Function)"""
        t = self.t
        if t.kind not in ('id', 'kw'):
            self.err(f"expected: {what}")
        self.pos += 1
        return t.text or t.value.title()

    def ident(self, what="identifier"):
        t = self.t
        if t.kind == 'id' or (t.kind == 'kw' and t.value in ('STRING', 'DATE', 'ERROR', 'INPUT', 'OUTPUT',
                                                              'APPEND', 'BINARY', 'RANDOM', 'READ', 'WRITE',
                                                              'LINE', 'GET', 'PUT', 'SEEK', 'LOCK', 'STEP',
                                                              'TO', 'PRINT', 'EMPTY', 'OBJECT', 'MODULE', 'LIB',
                                                              'TEXT', 'COMPARE', 'EXPLICIT', 'BASE', 'ACCESS',
                                                              'SHARED', 'LEN', 'RESET', 'NAME', 'PROPERTY')):
            self.pos += 1                          # (Dim Property As String: VB6 lets a variable be called Property)
            return (t.text or t.value), t.suffix
        self.err(f"expected: {what}")

    # ---------------------------------------------------------------- module
    def parse_module(self):
        items = self.module.items
        while True:
            self.skip_nl()
            t = self.t
            if t.kind == 'eof':
                break
            if t.kind == 'id' and t.value.upper() == 'ATTRIBUTE':
                while self.t.kind not in ('nl', 'eof'):
                    self.pos += 1
                continue
            items.append(self.module_item())
        return self.module

    def module_item(self):
        t = self.t
        if t.is_kw('OPTION'):
            return self.option_stmt()
        if t.kind == 'kw' and t.value in DEFTYPES:
            return self.deftype_stmt()
        scope = None
        if t.is_kw('PUBLIC', 'PRIVATE', 'GLOBAL', 'FRIEND'):
            scope = self.next().value.lower()
        static = self.accept_kw('STATIC')
        t = self.t
        if t.is_kw('SUB', 'FUNCTION') or (t.is_kw('PROPERTY') and self.peek().is_kw('GET', 'LET', 'SET')):
            return self.proc(scope or 'public', static)
        if t.is_kw('CONST'):
            self.next()
            n = self.const_decl(scope or 'private')
            self.end_stmt()
            return n
        if t.is_kw('ENUM'):
            return self.enum_def(scope or 'public')
        if t.is_kw('TYPE'):
            return self.type_def(scope or 'public')
        if t.is_kw('DECLARE'):
            return self.declare_stmt(scope or 'public')
        if t.is_kw('EVENT'):
            line = self.next().line
            name, _ = self.ident("event name")
            params = []
            if self.accept_op('('):
                if not self.t.is_op(')'):
                    while True:
                        params.append(self.param())
                        if not self.accept_op(','):
                            break
                self.expect_op(')')
            self.end_stmt()
            if any(p.optional for p in params):
                self.err("events cannot have Optional parameters")
            return N('event', line, scope=scope or 'public', name=name, params=params)
        if t.is_kw('IMPLEMENTS'):
            line = self.next().line
            name, _ = self.ident("interface name")
            while self.t.is_op('.') and not self.t.space_before:
                self.next()
                name += '.' + self.ident("interface name")[0]
            self.end_stmt()
            return N('implements', line, name=name)
        if t.is_kw('DIM') or scope or static:
            if t.is_kw('DIM'):
                self.next()
            n = self.var_decl(scope or 'private', line=t.line)
            self.end_stmt()
            return n
        self.err("statement not valid outside a procedure")

    def option_stmt(self):
        line = self.next().line
        if self.accept_word('EXPLICIT'):
            self.module.options['explicit'] = True
        elif self.accept_word('BASE'):
            t = self.next()
            if t.kind != 'num' or t.value not in (0, 1):
                self.err("expected: 0 or 1", t)
            self.module.options['base'] = t.value
        elif self.accept_word('COMPARE'):
            if self.accept_word('TEXT'):
                self.module.options['compare'] = 'text'
            elif self.accept_word('BINARY'):
                self.module.options['compare'] = 'binary'
            elif self.accept_word('DATABASE'):
                self.module.options['compare'] = 'binary'
            else:
                self.err("expected: Text or Binary")
        elif self.accept_word('PRIVATE'):
            self.accept_word('MODULE')
        else:
            self.err("expected: Explicit, Base, Compare or Private Module")
        self.end_stmt()
        return N('option', line)

    def deftype_stmt(self):
        t = self.next()
        ranges = []
        while True:
            a, _ = self.ident("letter")
            b = a
            if self.accept_op('-'):
                b, _ = self.ident("letter")
            ranges.append((a[0].upper(), b[0].upper()))
            if not self.accept_op(','):
                break
        self.end_stmt()
        return N('deftype', t.line, type=DEFTYPES[t.value], ranges=ranges)

    def type_ref(self):
        """after As: a type name, maybe 'String * n' or 'New X'"""
        new = self.accept_kw('NEW')
        t = self.t
        if t.kind in ('kw', 'id') and t.value.upper() in TYPE_WORDS:
            self.next()
            name = TYPE_WORDS[t.value.upper()]
            fixlen = None
            if name == 'String' and self.accept_op('*'):
                fixlen = self.postfix(self.primary())          # a number or a constant
            return N('type', t.line, name=name, new=new, fixlen=fixlen)
        if t.kind == 'kw' and t.value in ('PROPERTY', 'GLOBAL'):   # As DAO.Property, As VB.Global
            self.next()
            name = t.text or t.value.title()
        else:
            name, _ = self.ident("type name")
        while self.accept_op('.'):
            nt = self.t
            if nt.kind not in ('id', 'kw'):
                self.err("expected: type name")
            self.next()
            name += '.' + (nt.text or nt.value.title())
        return N('type', t.line, name=name, new=new, fixlen=None)

    def var_decl(self, scope, line):
        vars_ = []
        while True:
            withevents = self.accept_kw('WITHEVENTS')
            t = self.t
            name, suffix = self.ident("variable name")
            dims = None
            if self.accept_op('('):
                dims = []
                if not self.t.is_op(')'):
                    while True:
                        dims.append(self.bound_pair())
                        if not self.accept_op(','):
                            break
                self.expect_op(')')
            typ = None
            if self.accept_kw('AS'):
                typ = self.type_ref()
            if withevents and (typ is None or typ.new or dims is not None):
                self.err("WithEvents needs a class type: WithEvents x As ClassName")
            vars_.append(N('var', t.line, name=name, suffix=suffix, dims=dims, type=typ, withevents=withevents))
            if not self.accept_op(','):
                break
        return N('dim', line, scope=scope, vars=vars_)

    def bound_pair(self):
        a = self.expr()
        if self.accept_kw('TO'):
            return (a, self.expr())
        return (None, a)

    def const_decl(self, scope):
        line = self.t.line
        items = []
        while True:
            t = self.t
            name, suffix = self.ident("constant name")
            typ = self.type_ref() if self.accept_kw('AS') else None
            self.expect_op('=')
            items.append(N('constitem', t.line, name=name, suffix=suffix, type=typ, value=self.expr()))
            if not self.accept_op(','):
                break
        return N('const', line, scope=scope, items=items)

    def enum_def(self, scope):
        line = self.next().line
        name, _ = self.ident("enum name")
        self.end_stmt()
        members = []
        while True:
            self.skip_nl()
            if self.t.is_kw('END'):
                self.next()
                self.expect_kw('ENUM')
                self.end_stmt()
                break
            t = self.t
            mname, _ = self.ident("enum member")
            val = self.expr() if self.accept_op('=') else None
            members.append(N('enummember', t.line, name=mname, value=val))
            self.end_stmt()
        return N('enum', line, scope=scope, name=name, members=members)

    def type_def(self, scope):
        """Type Name / fields / End Type"""
        line = self.next().line
        name, _ = self.ident("type name")
        self.end_stmt()
        fields = []
        while True:
            self.skip_nl()
            if self.t.is_kw('END') and not self.peek().is_kw('AS') and not self.peek().is_op('('):
                self.next()
                self.expect_kw('TYPE')
                self.end_stmt()
                break
            t = self.t
            if t.kind == 'kw':                             # Next As Integer, End As Long: a keyword as a field name
                self.pos += 1
                fname, suffix = (t.text or t.value), t.suffix
            else:
                fname, suffix = self.ident("field name")
            dims = None
            if self.accept_op('('):
                dims = []
                if not self.t.is_op(')'):
                    while True:
                        dims.append(self.bound_pair())
                        if not self.accept_op(','):
                            break
                self.expect_op(')')
            typ = self.type_ref() if self.accept_kw('AS') else None
            fields.append(N('field', t.line, name=fname, suffix=suffix, dims=dims, type=typ))
            self.end_stmt()
        if not fields:
            self.err("a Type needs at least one member")
        return N('type', line, scope=scope, name=name, fields=fields)

    def declare_stmt(self, scope):
        """Declare Sub|Function Name Lib "lib" [Alias "entry"] [(params)] [As type]"""
        line = self.next().line
        k = self.next()
        if not k.is_kw('SUB', 'FUNCTION'):
            self.err("expected: Sub or Function", k)
        kind = k.value.lower()
        name, suffix = self.ident("procedure name")
        if not self.accept_kw('LIB'):
            self.err("expected: Lib")
        lt = self.next()
        if lt.kind != 'str':
            self.err("expected: library name in quotes", lt)
        alias = None
        if self.accept_word('ALIAS'):
            at = self.next()
            if at.kind != 'str':
                self.err("expected: name in quotes after Alias", at)
            alias = at.value
        params = []
        if self.accept_op('('):
            if not self.t.is_op(')'):
                while True:
                    params.append(self.param())
                    if not self.accept_op(','):
                        break
            self.expect_op(')')
        rtype = self.type_ref() if self.accept_kw('AS') else None
        if kind == 'sub' and rtype is not None:
            self.err("Sub procedures do not return a value")
        self.end_stmt()
        return N('proc', line, pkind=kind, name=name, suffix=suffix, scope=scope, static=False, params=params,
                 rtype=rtype, body=[], lib=lt.value, alias=alias or name)

    def attribute_stmt(self):
        """Attribute X.VB_UserMemId = 0 (inside a procedure of a class)"""
        line = self.next().line
        target, _ = self.ident("attribute target")
        attr = ''
        if self.accept_op('.'):
            attr, _ = self.ident("attribute name")
            while self.accept_op('.'):          # X.VB_ProcData.VB_Invoke_Property
                more, _ = self.ident("attribute name")
                attr += '.' + more
        else:
            attr, target = target, ''
        self.expect_op('=')
        v = self.expr()
        self.end_stmt()
        val = getattr(v, 'value', None)
        if v.kind == 'un' and v.op == '-' and isinstance(getattr(v.e, 'value', None), (int, float)):
            val = -v.e.value                     # VB_UserMemId = -4 (NewEnum)
        self.module.attrs.append((target, attr, val))
        return N('nop', line)

    # ---------------------------------------------------------------- procedures
    def proc(self, scope, static):
        t = self.next()
        kind = t.value.lower()
        if kind == 'property':
            k = self.next()
            if not (k.is_kw('GET', 'LET', 'SET') or k.is_word('GET', 'LET', 'SET')):
                self.err("expected: Get, Let or Set", k)
            kind = 'property ' + k.value.lower()
        name, suffix = self.ident("procedure name")
        params = []
        if self.accept_op('('):
            if not self.t.is_op(')'):
                while True:
                    params.append(self.param())
                    if not self.accept_op(','):
                        break
            self.expect_op(')')
        rtype = None
        if self.accept_kw('AS'):
            rtype = self.type_ref()
            if self.t.is_op('(') and self.peek().is_op(')'):
                self.next(); self.next()
                rtype.isarray = True
            if kind in ('sub', 'property let', 'property set'):
                self.err("Sub procedures do not return a value" if kind == 'sub' else
                         "Property Let/Set procedures do not return a value")
        self.end_stmt()
        if kind in ('property let', 'property set') and not params:
            self.err(f"{kind.title()} needs a parameter for the value", t)
        end_word = 'SUB' if kind == 'sub' else 'FUNCTION' if kind == 'function' else 'PROPERTY'
        if self.recover:
            start = self.pos
            try:
                body = self.block(lambda: self.t.is_kw('END') and self.peek().is_kw(end_word))
            except VBError as e:
                self.pos = start
                self.resync_proc(end_word)
                return N('proc', t.line, pkind=kind, name=name, suffix=suffix, scope=scope, static=static,
                         params=params, rtype=rtype, body=[], parse_error=e)
        else:
            body = self.block(lambda: self.t.is_kw('END') and self.peek().is_kw(end_word))
        self.next()
        self.next()
        self.end_stmt()
        return N('proc', t.line, pkind=kind, name=name, suffix=suffix, scope=scope, static=static,
                 params=params, rtype=rtype, body=body)

    def resync_proc(self, end_word):
        """after a syntax error in a procedure: on to its End Sub / End Function / End Property at the start of a
        line (past it), or to the next procedure (the IDE starts one at each Sub / Function / Property line), or to
        the end of the file"""
        toks, i = self.toks, self.pos
        while toks[i].kind != 'eof':
            if i == 0 or toks[i - 1].kind == 'nl':
                if toks[i].is_kw('END') and toks[i + 1].is_kw(end_word):
                    i += 2
                    while toks[i].kind not in ('nl', 'eof'):
                        i += 1
                    break
                j = i
                while toks[j].is_kw('PRIVATE', 'PUBLIC', 'FRIEND', 'STATIC'):
                    j += 1
                if toks[j].is_kw('SUB', 'FUNCTION') or (toks[j].is_kw('PROPERTY') and
                                                       (toks[j + 1].is_kw('GET', 'LET', 'SET') or
                                                        toks[j + 1].is_word('GET', 'LET', 'SET'))):
                    if j > i or toks[j].is_kw('SUB', 'FUNCTION', 'PROPERTY'):
                        break
            i += 1
        self.pos = i

    def param(self):
        optional = self.accept_kw('OPTIONAL')
        byval = False
        if self.accept_kw('BYVAL'):
            byval = True
        else:
            self.accept_kw('BYREF')
        paramarray = self.accept_kw('PARAMARRAY')
        t = self.t
        name, suffix = self.ident("parameter name")
        isarray = False
        if self.accept_op('('):
            self.expect_op(')')
            isarray = True
        typ = self.type_ref() if self.accept_kw('AS') else None
        default = None
        if self.accept_op('='):
            if not optional:
                self.err("only Optional parameters can have a default value")
            default = self.expr()
        if paramarray:
            if not isarray or byval or optional or (typ is not None and typ.name != 'Variant'):
                self.err("ParamArray must be an array of Variant: ParamArray name() As Variant")
        return N('param', t.line, name=name, suffix=suffix, byval=byval, optional=optional, isarray=isarray,
                 type=typ, default=default, paramarray=paramarray)

    def block(self, stop):
        """statements until stop() is true at the start of a statement"""
        out = []
        while True:
            self.skip_nl()
            if self.t.kind == 'eof':
                self.err("unexpected end of file")
            if stop():
                return out
            out.extend(self.statement())

    # ---------------------------------------------------------------- statements
    def statement(self):
        """one statement (may return several nodes: a label and the statement after it)"""
        t = self.t
        out = []
        # line numbers and labels
        if t.kind == 'num' and isinstance(t.value, int) and self.toks[self.pos - 1].kind == 'nl' or \
                (t.kind == 'num' and self.pos == 0):
            self.next()
            out.append(N('label', t.line, name=str(t.value)))
            if self.at_end():
                self.accept_op(':')
                return out
            t = self.t
            # 110 End If: a line number before the end of a block (the block sees its end)
            if t.is_kw('ELSE', 'ELSEIF', 'LOOP', 'NEXT', 'WEND', 'CASE') or \
                    (t.is_kw('END') and self.peek().is_kw('IF', 'SELECT', 'WITH', 'SUB', 'FUNCTION', 'PROPERTY')):
                return out
        if (t.kind == 'id' or (t.kind == 'kw' and t.value in ('ERROR', 'OUTPUT', 'BINARY', 'RANDOM', 'APPEND', 'STEP',
                                                               'READ', 'LINE', 'WRITE', 'INPUT', 'GET', 'PUT'))) and \
                self.peek().is_op(':') and not t.suffix and \
                (t.kind != 'id' or t.value.upper() not in ('DOEVENTS', 'BEEP', 'RANDOMIZE', 'CLS', 'REFRESH')) and \
                (self.pos == 0 or self.toks[self.pos - 1].kind == 'nl' or out):   # (300 GetKeyError: too)
            self.next()
            self.next()
            out.append(N('label', t.line, name=t.text))
            return out
        out.append(self.simple_statement())
        return out

    # keywords that are keywords only inside Open / For / Line Input ...: elsewhere names (outPut "x",
    # Binary = True, line = LCase(s), Step = 2, Random)
    SOFT_ALWAYS = ('OUTPUT', 'BINARY', 'RANDOM', 'APPEND', 'STEP', 'READ')
    SOFT_BEFORE = ('LINE', 'WRITE', 'INPUT', 'GET', 'PUT', 'LOCK', 'SEEK', 'ERROR', 'LIB', 'DATE')

    def soft_name(self):
        """a keyword used as a name at the start of a statement becomes a name"""
        t, nx = self.t, self.peek()
        if t.kind != 'kw':
            return
        if t.value in self.SOFT_ALWAYS or \
                (t.value in self.SOFT_BEFORE and (nx.is_op('=') or (nx.is_op('.') and not nx.space_before)) or
                 (t.value == 'LIB' and nx.is_op('(')) or       # Lib(i).Caption: a control array called Lib
                 (t.value == 'LINE' and not nx.is_kw('INPUT') and not nx.is_op('(', '-') and not nx.is_kw('STEP'))):
            t.kind, t.value = 'id', t.text or t.value.title()

    def simple_statement(self):
        t = self.t
        line = t.line
        if t.kind == 'kw' and t.value in ('PRINT', 'LINE', 'ME'):
            g = self.graphics_stmt()
            if g is not None:
                return g
        self.soft_name()
        t = self.t
        if t.is_kw('PROPERTY') and not self.peek().is_kw('GET', 'LET', 'SET'):
            return self.assign_or_call()           # Property = x: a variable called Property
        if t.kind == 'kw':
            k = t.value
            m = getattr(self, 'st_' + k.lower(), None)
            if m:
                return m()
        if t.kind == 'id' and t.value.upper() == 'ATTRIBUTE' and not self.peek().is_op('='):
            return self.attribute_stmt()
        if t.is_op('.') or (t.is_op('!') and self.peek().kind in ('id', 'kw')):
            if t.is_op('.'):
                g = self.graphics_stmt()        # .picTarget.Line (x, y)-(x2, y2) inside With
                if g is not None:
                    return g
            return self.assign_or_call()
        if t.kind == 'id':
            up = t.value.upper()
            if up == 'DEBUG' and self.peek().is_op('.') and self.peek(2).is_kw('PRINT'):
                self.pos += 3
                n = N('print', line, file=None, items=self.print_items(), debug=True)
                self.end_stmt()
                return n
            if up in ('MID', 'MIDB') and self.peek().is_op('(') and self._is_mid_assignment():
                return self.st_mid()
            if up == 'NAME' and not self.peek().is_op('='):
                self.next()
                a = self.expr()
                self.expect_kw('AS')
                b = self.expr()
                self.end_stmt()
                return N('call', line, target=N('name', line, name='Name', suffix=None), args=[a, b], paren=False)
            if up == 'RANDOMIZE' and not self.peek().is_op('='):
                self.next()
                e = None if self.at_end() else self.expr()
                self.end_stmt()
                return N('randomize', line, seed=e)
        g = self.graphics_stmt()
        if g is not None:
            return g
        if t.is_kw('STEP') and self.peek().is_op('=', '('):   # a property named Step: Step = 1
            t.kind, t.value = 'id', t.text or 'Step'
        if t.kind in ('id',) or t.is_kw('ME'):
            return self.assign_or_call()
        self.err("syntax error")

    # ---------------------------------------------------------------- graphics methods
    GMETHODS = ('PSET', 'LINE', 'CIRCLE', 'SCALE', 'PRINT')

    def graphics_stmt(self):
        """[object.]PSet / Line / Circle / Scale / Print with their special syntax"""
        start = self.pos
        k = self.pos
        path = []
        t = self.toks[k]
        # object path: [.]name[(index)](.name[(index)])* .Method   (Pic(1).Line, .picTarget.Line)
        complex_path = False
        if t.is_op('.') and self.toks[k + 1].kind in ('id', 'kw'):
            complex_path = True
            k += 1
            t = self.toks[k]
        while t.kind == 'id' or t.is_kw('ME') or (complex_path and t.kind == 'kw' and path):
            j = k + 1
            if self.toks[j].is_op('(') and not self.toks[j].space_before:
                depth = 0
                while True:
                    if self.toks[j].is_op('('):
                        depth += 1
                    elif self.toks[j].is_op(')'):
                        depth -= 1
                        if depth == 0:
                            break
                    elif self.toks[j].kind in ('nl', 'eof'):
                        return None
                    j += 1
                j += 1
                if not (self.toks[j].is_op('.') and not self.toks[j].space_before):
                    break
                complex_path = True
            if not (self.toks[j].is_op('.') and not self.toks[j].space_before):
                break
            path.append(t)
            k = j + 1
            t = self.toks[k]
        if not (t.kind in ('id', 'kw') and t.value.upper() in self.GMETHODS) or (complex_path and not path):
            return None
        name = t.value.upper()
        nxt = self.toks[k + 1]
        if not path:
            # unqualified: only the forms that cannot be anything else
            if name == 'PRINT' and (nxt.is_op('#')):
                return None
            if name == 'LINE' and not (nxt.is_op('(', '-') or nxt.is_kw('STEP')):
                return None
            if name in ('PSET', 'CIRCLE') and not (nxt.is_op('(') or nxt.is_kw('STEP')):
                return None
            if name == 'SCALE' and not (nxt.is_op('(') or nxt.kind in ('nl', 'eof') or nxt.is_op(':')):
                return None
        elif name not in ('PRINT',) and nxt.is_op('='):
            return None                         # obj.Scale = ... is not a method call
        line = t.line
        obj = None
        if complex_path:                        # the object: parsed as an expression up to the '.' of the method
            saved = self.toks
            self.toks = saved[start:k - 1] + [Tok('eof', None, line, 1)]
            self.pos = 0
            try:
                obj = self.postfix(self.primary())
            finally:
                self.toks = saved
        else:
            for i, pt in enumerate(path):
                nm = N('name', pt.line, name=pt.text or pt.value, suffix=pt.suffix)
                obj = nm if obj is None else N('member', pt.line, base=obj, name=pt.text or pt.value, suffix=pt.suffix)
        self.pos = k + 1
        if name == 'PRINT':
            items = self.print_items()
            self.end_stmt()
            return N('gmethod', line, obj=obj, name='print', items=items)
        if name == 'PSET':
            st, x, y = self.gpoint()
            color = self.expr() if self.accept_op(',') else None
            self.end_stmt()
            return N('gmethod', line, obj=obj, name='pset', step=st, x=x, y=y, color=color)
        if name == 'LINE':
            p1 = None
            if not self.t.is_op('-'):
                p1 = self.gpoint()
            self.expect_op('-')
            p2 = self.gpoint()
            color, box = None, 0
            if self.accept_op(','):
                if not self.t.is_op(','):
                    color = self.expr()
                if self.accept_op(','):
                    w = self.next()
                    word = (w.text or w.value or '').upper()
                    if word not in ('B', 'BF'):
                        self.err("expected: B or BF", w)
                    box = 2 if word == 'BF' else 1
            self.end_stmt()
            return N('gmethod', line, obj=obj, name='line', p1=p1, p2=p2, color=color, box=box)
        if name == 'CIRCLE':
            st, x, y = self.gpoint()
            self.expect_op(',')
            r = self.expr()
            extra = []
            while self.accept_op(','):
                extra.append(None if self.t.is_op(',') or self.at_end() else self.expr())
            self.end_stmt()
            return N('gmethod', line, obj=obj, name='circle', step=st, x=x, y=y, r=r, extra=extra)
        # Scale [(x1, y1)-(x2, y2)]
        p1 = p2 = None
        if self.t.is_op('('):
            p1 = self.gpoint()
            self.expect_op('-')
            p2 = self.gpoint()
        self.end_stmt()
        return N('gmethod', line, obj=obj, name='scale', p1=p1, p2=p2)

    def gpoint(self):
        """[Step] (x, y)"""
        step = self.accept_kw('STEP')
        self.expect_op('(')
        x = self.expr()
        self.expect_op(',')
        y = self.expr()
        self.expect_op(')')
        return step, x, y

    def _is_mid_assignment(self):
        depth, k = 0, self.pos + 1
        while k < len(self.toks):
            tk = self.toks[k]
            if tk.is_op('('):
                depth += 1
            elif tk.is_op(')'):
                depth -= 1
                if depth == 0:
                    return self.toks[k + 1].is_op('=')
            elif tk.kind in ('nl', 'eof'):
                return False
            k += 1
        return False

    def st_mid(self):
        byte = self.t.value.upper() == 'MIDB'
        line = self.next().line
        self.expect_op('(')
        target = self.postfix(self.primary())
        self.expect_op(',')
        start = self.expr()
        length = self.expr() if self.accept_op(',') else None
        self.expect_op(')')
        self.expect_op('=')
        value = self.expr()
        self.end_stmt()
        return N('midstmt', line, target=target, start=start, length=length, value=value, byte=byte)

    def assign_or_call(self):
        line = self.t.line
        target = self.postfix(self.primary(), stmt_start=True)
        if self.accept_op('='):
            value = self.expr()
            self.end_stmt()
            return N('assign', line, target=target, value=value, set=False)
        if target.kind == 'apply' and self.at_end():
            # Foo(x) without Call: the parentheses belong to the argument, which is passed by value
            if len(target.args) > 1:
                self.err("expected: = (use Call to pass several arguments in parentheses)")
            self.end_stmt()
            args = [N('paren', line, e=a) if a is not None and a.kind != 'named' else a for a in target.args]
            return N('call', line, target=target.base, args=args, paren=True)
        args = []
        if target.kind == 'apply' and self.t.is_op(','):
            # Foo (a), b : the parenthesised first argument is passed by value
            args = [N('paren', line, e=a) if a is not None else None for a in target.args]
            target = target.base
            self.next()
            args += self.arg_list(stop_at_end=True)
        elif not self.at_end():
            args = self.arg_list(stop_at_end=True)
        self.end_stmt()
        return N('call', line, target=target, args=args, paren=False)

    def arg_list(self, stop_at_end=False):
        args = []
        while True:
            if self.t.is_op(',') or (stop_at_end and self.at_end()) or self.t.is_op(')'):
                args.append(None)                     # omitted argument
            else:
                args.append(self.argument())
            if not self.accept_op(','):
                break
        if args == [None]:
            return []
        return args

    def argument(self):
        t = self.t
        if t.kind in ('id', 'kw') and self.peek().is_op(':='):      # Append:=True: a keyword as a name too
            self.pos += 2
            return N('named', t.line, name=t.text or t.value.title(), e=self.expr())
        if self.accept_kw('BYVAL'):
            return N('paren', t.line, e=self.expr())
        if t.is_op('#'):                              # Input(n, #1), EOF(#1)
            self.next()
            return N('filenum', t.line, e=self.expr())
        return self.expr()

    def st_call(self):
        line = self.next().line
        target = self.postfix(self.primary())
        self.end_stmt()
        if target.kind == 'apply':
            b = target.base
            if b.kind == 'name' and b.name.upper() == 'RANDOMIZE' and len(target.args) <= 1:
                return N('randomize', line, seed=target.args[0] if target.args else None)   # Call Randomize(Timer)
            return N('call', line, target=target.base, args=target.args, paren=True)
        if target.kind == 'name' and target.name.upper() == 'RANDOMIZE':
            return N('randomize', line, seed=None)
        return N('call', line, target=target, args=[], paren=True)

    def st_let(self):
        self.next()
        return self.assign_or_call()

    def st_lset(self, right=False):
        line = self.next().line
        target = self.postfix(self.primary(), stmt_start=True)
        self.expect_op('=')
        value = self.expr()
        self.end_stmt()
        return N('lset', line, target=target, value=value, right=right)

    def st_rset(self):
        return self.st_lset(right=True)

    def st_set(self):
        line = self.next().line
        target = self.postfix(self.primary(), stmt_start=True)
        self.expect_op('=')
        value = self.expr()
        self.end_stmt()
        return N('assign', line, target=target, value=value, set=True)

    def _getput(self, kind):
        line = self.next().line
        self.accept_op('#')
        f = self.expr()
        self.expect_op(',')
        rec = None if self.t.is_op(',') else self.expr()
        self.expect_op(',')
        target = self.expr() if kind == 'put' else self.postfix(self.primary())   # Put #1, , "PNG" & x
        self.end_stmt()
        return N(kind, line, file=f, rec=rec, target=target)

    def st_get(self):
        return self._getput('get')

    def st_put(self):
        return self._getput('put')

    def st_seek(self):
        line = self.next().line
        self.accept_op('#')
        f = self.expr()
        self.expect_op(',')
        pos = self.expr()
        self.end_stmt()
        return N('seek', line, file=f, pos=pos)

    def st_raiseevent(self):
        line = self.next().line
        name, _ = self.ident("event name")
        args = []
        if self.accept_op('('):
            args = [] if self.t.is_op(')') else self.arg_list()
            self.expect_op(')')
        self.end_stmt()
        return N('raiseevent', line, name=name, args=args)

    def st_with(self):
        line = self.next().line
        e = self.expr()
        self.end_stmt()
        body = self.block(lambda: self.t.is_kw('END') and self.peek().is_kw('WITH'))
        self.next()
        self.next()
        self.end_stmt()
        return N('with', line, e=e, body=body)

    def st_me(self):
        return self.assign_or_call()

    def st_dim(self):
        line = self.next().line
        n = self.var_decl('dim', line)
        self.end_stmt()
        return n

    def st_static(self):
        line = self.next().line
        n = self.var_decl('static', line)
        self.end_stmt()
        return n

    def st_const(self):
        self.next()
        n = self.const_decl('local')
        self.end_stmt()
        return n

    def st_private(self):
        self.err("Private is not valid inside a procedure")

    st_public = st_global = st_private

    def st_redim(self):
        line = self.next().line
        preserve = self.accept_kw('PRESERVE')
        items = []
        while True:
            t = self.t
            target = None
            if t.is_op('.'):                    # ReDim .Data(n) inside With
                self.next()
                nt = self.t
                name, suffix = self.ident("member name")
                target = N('member', nt.line, base=N('withref', t.line), name=name, suffix=suffix)
            else:
                name, suffix = self.ident("array name")
            # a field of a Type: ReDim a.b(n), ReDim a(i).b(n), ReDim .b(n)
            while (self.t.is_op('.') and not self.t.space_before) or (self.t.is_op('(') and self._paren_then_dot()):
                if target is None:
                    target = N('name', t.line, name=name, suffix=suffix)
                if self.t.is_op('('):
                    self.next()
                    args = [self.expr()]
                    while self.accept_op(','):
                        args.append(self.expr())
                    self.expect_op(')')
                    target = N('apply', t.line, base=target, args=args)
                    continue
                self.next()
                nt = self.t
                mname, msuf = self.ident("member name")
                target = N('member', nt.line, base=target, name=mname, suffix=msuf)
            self.expect_op('(')
            dims = []
            while True:
                dims.append(self.bound_pair())
                if not self.accept_op(','):
                    break
            self.expect_op(')')
            typ = self.type_ref() if self.accept_kw('AS') else None
            items.append(N('redimitem', t.line, name=name, suffix=suffix, dims=dims, type=typ, target=target))
            if not self.accept_op(','):
                break
        self.end_stmt()
        return N('redim', line, preserve=preserve, items=items)

    def _paren_then_dot(self):
        """the ( ... ) here is followed by .member (an index on the way to the array, not its bounds)"""
        k, depth = self.pos, 0
        while k < len(self.toks):
            t = self.toks[k]
            if t.is_op('('):
                depth += 1
            elif t.is_op(')'):
                depth -= 1
                if depth == 0:
                    nx = self.toks[k + 1]
                    return nx.is_op('.') and not nx.space_before
            elif t.kind in ('nl', 'eof'):
                return False
            k += 1
        return False

    def st_erase(self):
        line = self.next().line
        names = []
        while True:
            names.append(self.postfix(self.primary()))
            if not self.accept_op(','):
                break
        self.end_stmt()
        return N('erase', line, names=names)

    def st_if(self):
        line = self.next().line
        cond = self.expr()
        self.expect_kw('THEN')
        if self.t.kind == 'nl' or (self.t.is_op(':') and False):
            # block If
            self.end_stmt()
            then = self.block(lambda: self.t.is_kw('ELSEIF', 'ELSE') or
                              (self.t.is_kw('END') and self.peek().is_kw('IF')))
            elifs, els = [], None
            while self.t.is_kw('ELSEIF'):
                self.next()
                c = self.expr()
                self.expect_kw('THEN')
                if self.at_end():                   # ElseIf x Then y: the statement may follow on the line
                    self.end_stmt()
                b = self.block(lambda: self.t.is_kw('ELSEIF', 'ELSE') or
                               (self.t.is_kw('END') and self.peek().is_kw('IF')))
                elifs.append((c, b))
            if self.accept_kw('ELSE'):
                self.end_stmt() if not self.at_end() or self.t.kind == 'nl' or self.t.is_op(':') else None
                els = self.block(lambda: self.t.is_kw('END') and self.peek().is_kw('IF'))
            self.next()
            self.next()
            self.end_stmt()
            return N('if', line, cond=cond, then=then, elifs=elifs, els=els)
        # single-line If: statements separated by ':' up to Else / end of line
        then = self.inline_block()
        els = None
        if self.accept_kw('ELSE'):
            els = self.inline_block()
        if self.t.kind == 'nl' and not getattr(self, 'inline_depth', 0):   # If a Then If b Then x: the outer ends the line
            self.next()
        return N('if', line, cond=cond, then=then, elifs=[], els=els)

    def inline_block(self):
        out = []
        while self.t.is_op(':'):                # If x Then: y / Else: y
            self.next()
        if self.t.kind == 'num' and isinstance(self.t.value, int):    # If x Then 100 (GoTo 100)
            t = self.next()
            return [N('goto', t.line, label=str(t.value))]
        while True:
            if self.t.kind in ('nl', 'eof') or self.t.is_kw('ELSE'):
                return out
            out.append(self.simple_statement_inline())
            while self.t.is_op(':'):
                self.next()

    def simple_statement_inline(self):
        # like simple_statement, but leaves ':' / Else / newline for the single-line If
        save_end = self.end_stmt
        self.end_stmt = self._end_inline
        self.inline_depth = getattr(self, 'inline_depth', 0) + 1
        try:
            return self.simple_statement()
        finally:
            self.end_stmt = save_end
            self.inline_depth -= 1

    def _end_inline(self):
        if not (self.t.kind in ('nl', 'eof') or self.t.is_op(':') or self.t.is_kw('ELSE')):
            self.err("expected: end of statement")

    def st_select(self):
        line = self.next().line
        self.expect_kw('CASE')
        subject = self.expr()
        self.end_stmt()
        cases, els = [], None
        self.skip_nl()
        while True:
            self.skip_nl()
            if self.t.is_kw('END'):
                self.next()
                self.expect_kw('SELECT')
                self.end_stmt()
                break
            ct = self.expect_kw('CASE')
            stop = lambda: self.t.is_kw('CASE') or (self.t.is_kw('END') and self.peek().is_kw('SELECT'))
            if self.accept_kw('ELSE'):
                self.end_stmt()
                els = self.block(stop)
                continue
            conds = []
            while True:
                if self.accept_kw('IS'):
                    op = self.next()
                    if not op.is_op(*CMP_OPS):
                        self.err("expected: comparison operator", op)
                    conds.append(('is', op.value, self.expr()))
                else:
                    a = self.expr()
                    if self.accept_kw('TO'):
                        conds.append(('range', a, self.expr()))
                    else:
                        conds.append(('val', a))
                if not self.accept_op(','):
                    break
            self.end_stmt()
            cases.append((conds, self.block(stop), ct.line))
        return N('select', line, subject=subject, cases=cases, els=els)

    def st_for(self):
        line = self.next().line
        if self.accept_kw('EACH'):
            var = self.postfix(self.primary())
            self.expect_kw('IN')
            coll = self.expr()
            self.end_stmt()
            body = self.block(lambda: self.t.is_kw('NEXT'))
            self.next()
            if not self.at_end():
                self.loop_var()
                if self.t.is_op(','):
                    self.t.kind, self.t.value, self.t.text = 'kw', 'NEXT', 'Next'
                    return N('foreach', line, var=var, coll=coll, body=body)
            self.end_stmt()
            return N('foreach', line, var=var, coll=coll, body=body)
        var = self.postfix(self.primary())
        self.expect_op('=')
        start = self.expr()
        self.expect_kw('TO')
        end = self.expr()
        step = self.expr() if self.accept_kw('STEP') else None
        self.end_stmt()
        body = self.block(lambda: self.t.is_kw('NEXT'))
        self.next()
        # Next [i] [, j]: "Next i, j" closes two loops
        if not self.at_end():
            self.loop_var()
            if self.t.is_op(','):               # Next i, j: this Next closes i, a Next j is left for the outer For
                self.t.kind, self.t.value, self.t.text = 'kw', 'NEXT', 'Next'
                return N('for', line, var=var, start=start, end=end, step=step, body=body)
        self.end_stmt()
        return N('for', line, var=var, start=start, end=end, step=step, body=body)

    def loop_var(self):
        """the variable after Next: a name, or a member (Next Visual.Y), or an element (Next a(1))"""
        self.ident("loop variable")
        while True:
            if self.t.is_op('.') and not self.t.space_before:
                self.next()
                self.next()
            elif self.t.is_op('('):
                depth = 0
                while True:
                    t = self.next()
                    if t.is_op('('):
                        depth += 1
                    elif t.is_op(')'):
                        depth -= 1
                        if depth == 0:
                            break
                    elif t.kind in ('nl', 'eof'):
                        break
            else:
                return

    def st_do(self):
        line = self.next().line
        pre = None
        if self.t.is_kw('WHILE', 'UNTIL'):
            k = self.next().value.lower()
            pre = (k, self.expr())
        self.end_stmt()
        body = self.block(lambda: self.t.is_kw('LOOP'))
        self.next()
        post = None
        if self.t.is_kw('WHILE', 'UNTIL'):
            k = self.next().value.lower()
            post = (k, self.expr())
        self.end_stmt()
        if pre and post:
            self.err("a Do loop can have While/Until at the start or at the end, not both")
        return N('do', line, pre=pre, post=post, body=body)

    def st_while(self):
        line = self.next().line
        cond = self.expr()
        self.end_stmt()
        body = self.block(lambda: self.t.is_kw('WEND'))
        self.next()
        self.end_stmt()
        return N('do', line, pre=('while', cond), post=None, body=body)

    def st_exit(self):
        line = self.next().line
        k = self.next()
        if not k.is_kw('SUB', 'FUNCTION', 'FOR', 'DO', 'PROPERTY'):
            self.err("expected: Do, For, Sub, Function or Property", k)
        self.end_stmt()
        return N('exit', line, what=k.value.lower())

    def st_goto(self):
        line = self.next().line
        t = self.next()
        if t.kind not in ('id', 'num') and not (t.kind == 'kw' and t.text):
            self.err("expected: label", t)
        self.end_stmt()
        return N('goto', line, label=str(t.text if t.kind in ('id', 'kw') else t.value))

    def st_gosub(self):
        line = self.next().line
        t = self.next()
        if t.kind not in ('id', 'num') and not (t.kind == 'kw' and t.text):
            self.err("expected: label", t)
        self.end_stmt()
        return N('gosub', line, label=str(t.text if t.kind in ('id', 'kw') else t.value))

    def st_return(self):
        line = self.next().line
        self.end_stmt()
        return N('return', line)

    def st_on(self):
        line = self.next().line
        if self.t.kind == 'id' and str(self.t.value).upper() == 'LOCAL' and self.peek().is_kw('ERROR'):
            self.next()                            # On Local Error: the old spelling of On Error
        if self.accept_kw('ERROR'):
            if self.accept_kw('RESUME'):
                self.expect_kw('NEXT')
                self.end_stmt()
                return N('onerror', line, mode='resumenext', label=None)
            self.expect_kw('GOTO')
            t = self.next()
            if t.kind == 'num' and t.value == 0:
                mode, label = 'goto0', None
            elif t.is_op('-') and self.t.kind == 'num' and self.t.value == 1:
                self.next()
                mode, label = 'goto-1', None
            elif t.kind in ('id', 'num') or (t.kind == 'kw' and t.text):     # On Error GoTo Error
                mode, label = 'goto', str(t.text if t.kind in ('id', 'kw') else t.value)
            else:
                self.err("expected: label", t)
            self.end_stmt()
            return N('onerror', line, mode=mode, label=label)
        e = self.expr()
        k = self.next()
        if not k.is_kw('GOTO', 'GOSUB'):
            self.err("expected: GoTo or GoSub", k)
        labels = []
        while True:
            t = self.next()
            if t.kind not in ('id', 'num') and not (t.kind == 'kw' and t.text):
                self.err("expected: label", t)
            labels.append(str(t.text if t.kind in ('id', 'kw') else t.value))
            if not self.accept_op(','):
                break
        self.end_stmt()
        return N('ongoto', line, e=e, labels=labels, gosub=k.value == 'GOSUB')

    def st_resume(self):
        line = self.next().line
        if self.accept_kw('NEXT'):
            target = 'next'
        elif self.at_end():
            target = ''
        else:
            t = self.next()
            target = str(t.text if t.kind in ('id', 'kw') else t.value)
            if target == '0':
                target = ''
        self.end_stmt()
        return N('resume', line, target=target)

    def st_end(self):
        line = self.next().line
        if not self.at_end():
            self.err(f"unexpected End {self.t.value}")
        self.end_stmt()
        return N('end', line)

    def st_stop(self):
        line = self.next().line
        self.end_stmt()
        return N('stop', line)

    def st_error(self):
        line = self.next().line
        e = self.expr()
        self.end_stmt()
        return N('errorstmt', line, number=e)

    def file_number(self):
        self.accept_op('#')
        return self.expr()

    def st_print(self):
        line = self.next().line
        if not self.t.is_op('#'):
            self.err("method not valid without suitable object: Print")
        self.next()
        f = self.expr()
        items = []
        if self.accept_op(','):
            items = self.print_items()
        self.end_stmt()
        return N('print', line, file=f, items=items)

    def print_items(self):
        items = []
        while not self.at_end():
            if self.accept_op(';'):
                items.append((';',))
            elif self.accept_op(','):
                items.append((',',))
            elif self.t.is_word('SPC') and self.peek().is_op('('):
                self.pos += 2
                items.append(('spc', self.expr()))
                self.expect_op(')')
            elif self.t.is_word('TAB'):
                self.next()
                if self.accept_op('('):
                    items.append(('tab', self.expr()))
                    self.expect_op(')')
                else:
                    items.append(('tab', None))
            else:
                items.append(('e', self.expr()))
        return items

    def st_write(self):
        line = self.next().line
        if not self.t.is_op('#'):
            self.err("expected: #")
        self.next()
        f = self.expr()
        exprs = []
        if self.accept_op(','):
            while not self.at_end():
                if self.t.is_op(','):              # Write #1, , , "x": an empty field
                    exprs.append(N('empty', self.t.line))
                    self.next()
                    continue
                exprs.append(self.expr())
                if not self.accept_op(',') and not self.accept_op(';'):
                    break
        self.end_stmt()
        return N('write', line, file=f, exprs=exprs)

    def st_open(self):
        line = self.next().line
        path = self.expr()
        mode = 'random'
        if self.accept_kw('FOR'):
            m = self.next()
            if not m.is_word('INPUT', 'OUTPUT', 'APPEND', 'BINARY', 'RANDOM'):
                self.err("expected: Input, Output, Append, Binary or Random", m)
            mode = m.value.lower()
        if self.accept_word('ACCESS'):
            while self.t.is_word('READ', 'WRITE'):
                self.next()
        while self.t.is_word('SHARED', 'LOCK', 'READ', 'WRITE'):
            self.next()
        self.expect_kw('AS')
        f = self.file_number()
        reclen = None
        if self.accept_word('LEN'):
            self.expect_op('=')
            reclen = self.expr()
        self.end_stmt()
        return N('open', line, path=path, mode=mode, file=f, reclen=reclen)

    def st_close(self):
        line = self.next().line
        files = []
        while not self.at_end():
            files.append(self.file_number())
            if not self.accept_op(','):
                break
        self.end_stmt()
        return N('close', line, files=files)

    def st_line(self):
        line = self.next().line
        if not self.accept_kw('INPUT'):
            self.err("syntax error")
        if not self.t.is_op('#'):
            self.err("expected: #")
        self.next()
        f = self.expr()
        self.expect_op(',')
        target = self.postfix(self.primary())
        self.end_stmt()
        return N('lineinput', line, file=f, target=target)

    def st_input(self):
        line = self.next().line
        if not self.t.is_op('#'):
            self.err("expected: #")
        self.next()
        f = self.expr()
        self.expect_op(',')
        targets = []
        while True:
            targets.append(self.postfix(self.primary()))
            if not self.accept_op(','):
                break
        self.end_stmt()
        return N('input', line, file=f, targets=targets)

    def st_rem(self):
        self.next()
        while self.t.kind not in ('nl', 'eof'):
            self.next()
        return N('nop', 0)

    # ---------------------------------------------------------------- expressions
    def expr(self):
        return self.imp_expr()

    def _binary(self, sub, ops, kw=True):
        left = sub()
        while True:
            t = self.t
            if (kw and t.kind == 'kw' and t.value in ops) or (t.kind == 'op' and t.value in ops):
                self.next()
                left = N('bin', t.line, op=t.value, l=left, r=sub())
            else:
                return left

    def imp_expr(self):
        return self._binary(self.eqv_expr, ('IMP',))

    def eqv_expr(self):
        return self._binary(self.xor_expr, ('EQV',))

    def xor_expr(self):
        return self._binary(self.or_expr, ('XOR',))

    def or_expr(self):
        return self._binary(self.and_expr, ('OR',))

    def and_expr(self):
        return self._binary(self.not_expr, ('AND',))

    def not_expr(self):
        if self.t.is_kw('NOT'):
            t = self.next()
            return N('un', t.line, op='NOT', e=self.not_expr())
        return self.cmp_expr()

    def cmp_expr(self):
        left = self.cat_expr()
        while True:
            t = self.t
            if t.is_op(*CMP_OPS):
                self.next()
                left = N('bin', t.line, op=t.value, l=left, r=self.cat_expr())
            elif t.is_kw('LIKE'):
                self.next()
                left = N('bin', t.line, op='LIKE', l=left, r=self.cat_expr())
            elif t.is_kw('IS'):
                self.next()
                left = N('bin', t.line, op='IS', l=left, r=self.cat_expr())
            else:
                return left

    def cat_expr(self):
        return self._binary(self.add_expr, ('&',), kw=False)

    def add_expr(self):
        return self._binary(self.mod_expr, ('+', '-'), kw=False)

    def mod_expr(self):
        return self._binary(self.idiv_expr, ('MOD',))

    def idiv_expr(self):
        return self._binary(self.mul_expr, ('\\',), kw=False)

    def mul_expr(self):
        return self._binary(self.neg_expr, ('*', '/'), kw=False)

    def neg_expr(self):
        t = self.t
        if t.is_op('-'):
            self.next()
            return N('un', t.line, op='-', e=self.neg_expr())
        if t.is_op('+'):
            self.next()
            return self.neg_expr()
        return self.pow_expr()

    def pow_expr(self):
        left = self.postfix(self.primary())
        while self.t.is_op('^'):
            t = self.next()
            # the exponent may carry its own sign: 2 ^ -1
            if self.t.is_op('-'):
                s = self.next()
                right = N('un', s.line, op='-', e=self.postfix(self.primary()))
            else:
                right = self.postfix(self.primary())
            left = N('bin', t.line, op='^', l=left, r=right)
        return left

    def primary(self):
        t = self.t
        if t.kind == 'num':
            self.next()
            return N('num', t.line, value=t.value, type=t.suffix)
        if t.kind == 'str':
            self.next()
            return N('str', t.line, value=t.value)
        if t.kind == 'date':
            self.next()
            return N('date', t.line, text=t.value)
        if t.is_op('('):
            self.next()
            e = self.expr()
            self.expect_op(')')
            return N('paren', t.line, e=e)
        if t.is_kw('TRUE', 'FALSE'):
            self.next()
            return N('bool', t.line, value=t.value == 'TRUE')
        if t.is_kw('NOT'):                         # x = Not y, 0# + Not (Not b): Not as an operand
            self.next()
            return N('un', t.line, op='NOT', e=self.not_expr())
        if t.is_kw('EMPTY', 'NULL', 'NOTHING'):
            self.next()
            return N(t.value.lower(), t.line)
        if t.is_kw('NEW'):
            self.next()
            name, _ = self.ident("class name")
            while self.t.is_op('.') and not self.t.space_before:
                self.next()
                n2 = self.member_word("class name")
                name += '.' + n2
            return N('new', t.line, cls=name)
        if t.is_kw('TYPEOF'):
            self.next()
            e = self.postfix(self.primary())
            self.expect_kw('IS')
            name, _ = self.ident("class name")
            while self.t.is_op('.') and not self.t.space_before:
                self.next()
                n2 = self.member_word("class name")
                name += '.' + n2
            return N('typeof', t.line, e=e, cls=name)
        if t.is_kw('ADDRESSOF'):                   # AddressOf Proc / AddressOf Module.Proc
            self.next()
            name, _ = self.ident("procedure name")
            mod = None
            if self.t.is_op('.') and not self.t.space_before:
                self.next()
                mod, (name, _) = name, self.ident("procedure name")
            return N('addressof', t.line, name=name, mod=mod)
        if t.kind == 'id' or (t.kind == 'kw' and t.value in ('STRING', 'DATE', 'ERROR', 'INPUT', 'ME', 'LEN', 'SEEK',
                                                              'OUTPUT', 'BINARY', 'RANDOM', 'APPEND', 'STEP', 'READ',
                                                              'WRITE', 'LINE', 'LOCK', 'GET', 'PUT', 'LIB', 'PROPERTY')):
            self.next()
            return N('name', t.line, name=t.text or t.value, suffix=t.suffix)
        if t.is_op('!') and self.peek().kind in ('id', 'kw') and not self.peek().space_before:   # !Field inside With
            self.next()
            nt = self.next()
            return N('apply', t.line, base=N('withref', t.line),
                     args=[N('str', nt.line, value=nt.text or str(nt.value))])
        if t.is_op('.'):                           # .member inside With
            self.next()
            nt = self.t
            if nt.kind not in ('id', 'kw'):
                self.err("expected: member name")
            self.next()
            return N('member', nt.line, base=N('withref', t.line), name=nt.text or nt.value.title(), suffix=nt.suffix)
        self.err("expected: expression")

    def postfix(self, e, stmt_start=False):
        while True:
            t = self.t
            if t.is_op('('):
                if stmt_start and e.kind in ('name', 'member') and self._paren_starts_argument():
                    return e                          # Me.Move (W - w) / 2, y: the ( begins the first argument
                self.next()
                args = [] if self.t.is_op(')') else self.arg_list()
                self.expect_op(')')
                e = N('apply', t.line, base=e, args=args)
            elif t.is_op('.') and not t.space_before:
                self.next()
                nt = self.t
                if nt.kind not in ('id', 'kw'):
                    self.err("expected: member name")
                self.next()
                e = N('member', nt.line, base=e, name=nt.text or nt.value.title(), suffix=nt.suffix)
            elif t.is_op('!') and not t.space_before and self.peek().kind in ('id', 'kw'):
                # rs!Name is rs("Name"): the default member with the name as a string
                self.next()
                nt = self.next()
                e = N('apply', t.line, base=e, args=[N('str', nt.line, value=nt.text or str(nt.value))])
            else:
                return e


    CONT_OPS = ('+', '-', '*', '/', '\\', '^', '&', '<', '>', '<=', '>=', '<>')
    CONT_KWS = ('MOD', 'AND', 'OR', 'XOR', 'EQV', 'IMP', 'LIKE', 'IS')

    def _paren_starts_argument(self):
        """at the ( after the name of a statement: the expression goes on after the matching )"""
        depth, k = 0, self.pos
        while k < len(self.toks):
            tk = self.toks[k]
            if tk.is_op('('):
                depth += 1
            elif tk.is_op(')'):
                depth -= 1
                if depth == 0:
                    nt = self.toks[k + 1]
                    return (nt.kind == 'op' and nt.value in self.CONT_OPS) or \
                        (nt.kind == 'kw' and nt.value in self.CONT_KWS)
            elif tk.kind in ('nl', 'eof'):
                return False
            k += 1
        return False


def parse_module(text, filename, module_name):
    return Parser(text, filename, module_name).parse_module()
