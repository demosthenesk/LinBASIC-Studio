# LinBASIC Studio - conditional compilation of VB6 (#Const, #If ... #End If)
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""#If / #ElseIf / #Else / #End If and #Const, before the parser.  The constants are those of
the .vbp (CondComp="DEBUGMODE = 1 : TRACE = 0"), of #Const lines, and Win16 = False, Win32 = True
as in VB6.  A name that is not defined is Empty (0).  The lines that are left out (and the #
lines) become empty lines, so the line numbers of the file stay the same."""
import re

from .lexer import VBError

_TOK = re.compile(r'\s*(?:(?P<num>&H[0-9A-Fa-f]+&?|\d+(?:\.\d*)?(?:[eE][-+]?\d+)?)|(?P<str>"(?:[^"]|"")*")|'
                  r'(?P<id>[A-Za-z_][A-Za-z0-9_]*)|(?P<op><>|<=|>=|[-+*/\\^&=<>()]))')


def parse_condcomp(text):
    """CondComp of a .vbp -> {name lower: value}"""
    out = {}
    for part in (text or "").split(":"):
        if "=" in part:
            k, v = part.split("=", 1)
            try:
                out[k.strip().lower()] = evaluate(v, {})
            except VBError:
                pass
    return out


def evaluate(expr, consts):
    """the value of a constant expression, with the precedence of VB: comparisons before Not,
    And, Or/Xor"""
    toks = []
    pos = 0
    expr = expr.strip()
    while pos < len(expr):
        m = _TOK.match(expr, pos)
        if not m or m.end() == pos:
            if expr[pos:].strip() == "":
                break
            raise VBError(f"invalid expression in #If: {expr}", 0, 0, "")
        pos = m.end()
        if m.group('num'):
            t = m.group('num')
            toks.append(('v', int(t[2:].rstrip('&'), 16) if t.upper().startswith('&H') else
                         (float(t) if any(c in t for c in '.eE') else int(t))))
        elif m.group('str'):
            toks.append(('v', m.group('str')[1:-1].replace('""', '"')))
        elif m.group('id'):
            w = m.group('id').lower()
            if w in ('and', 'or', 'xor', 'not', 'mod', 'eqv', 'imp'):
                toks.append(('op', w))
            elif w in ('true', 'false'):
                toks.append(('v', -1 if w == 'true' else 0))
            else:
                toks.append(('v', consts.get(w, 0)))
        else:
            toks.append(('op', m.group('op')))
    pos = [0]

    def peek():
        return toks[pos[0]] if pos[0] < len(toks) else ('end', None)

    def take(*ops):
        t = peek()
        if t[0] == 'op' and t[1] in ops:
            pos[0] += 1
            return t[1]
        return None

    def b(x):
        return -1 if x else 0

    def num(x):
        if isinstance(x, str):
            try:
                return float(x) if '.' in x else int(x)
            except ValueError:
                raise VBError(f"type mismatch in #If: {expr}", 0, 0, "")
        return x

    def or_():
        v = and_()
        while True:
            op = take('or', 'xor', 'eqv', 'imp')
            if not op:
                return v
            w = and_()
            v = {'or': lambda: num(v) | num(w), 'xor': lambda: num(v) ^ num(w),
                 'eqv': lambda: ~(num(v) ^ num(w)), 'imp': lambda: (~num(v)) | num(w)}[op]()

    def and_():
        v = not_()
        while take('and'):
            v = num(v) & num(not_())
        return v

    def not_():
        if take('not'):
            return ~num(not_())
        return cmp_()

    def cmp_():
        v = cat()
        while True:
            op = take('=', '<>', '<', '>', '<=', '>=')
            if not op:
                return v
            w = cat()
            if isinstance(v, str) != isinstance(w, str):
                v, w = num(v), num(w)
            v = b({'=': v == w, '<>': v != w, '<': v < w, '>': v > w, '<=': v <= w, '>=': v >= w}[op])

    def cat():
        v = add()
        while take('&'):
            v = str(v) + str(add())
        return v

    def add():
        v = mod()
        while True:
            op = take('+', '-')
            if not op:
                return v
            w = mod()
            v = (v + w) if op == '+' and isinstance(v, str) and isinstance(w, str) else \
                (num(v) + num(w) if op == '+' else num(v) - num(w))

    def mod():
        v = idiv()
        while take('mod'):
            v = int(num(v)) % int(num(idiv()))
        return v

    def idiv():
        v = mul()
        while take('\\'):
            v = int(num(v)) // int(num(mul()))
        return v

    def mul():
        v = unary()
        while True:
            op = take('*', '/')
            if not op:
                return v
            w = unary()
            v = num(v) * num(w) if op == '*' else num(v) / num(w)

    def unary():
        if take('-'):
            return -num(unary())
        if take('+'):
            return num(unary())
        return power()

    def power():
        v = atom()
        while take('^'):
            v = num(v) ** num(atom())
        return v

    def atom():
        t = peek()
        if t[0] == 'v':
            pos[0] += 1
            return t[1]
        if take('('):
            v = or_()
            if not take(')'):
                raise VBError(f"expected ) in #If: {expr}", 0, 0, "")
            return v
        raise VBError(f"invalid expression in #If: {expr}", 0, 0, "")

    try:
        v = or_()
    except ZeroDivisionError:
        raise VBError(f"division by zero in #If: {expr}", 0, 0, "")
    if pos[0] != len(toks):
        raise VBError(f"invalid expression in #If: {expr}", 0, 0, "")
    return v


def _truth(v):
    if isinstance(v, str):
        try:
            return float(v) != 0
        except ValueError:
            return bool(v)
    return bool(v)


def preprocess(lines, consts, first_line, filename):
    """lines of code -> the same number of lines, the left out ones empty"""
    consts = {'win16': 0, 'win32': -1, **{k.lower(): v for k, v in consts.items()}}
    out = []
    stack = []                     # (active, taken, parent_active)
    active = True
    skip = 0
    for i, line in enumerate(lines):
        if skip:                   # the continuation lines of a #If (kept empty: the line numbers stay)
            skip -= 1
            out.append("")
            continue
        s = line.strip()
        low = s.lower()
        ln = first_line + i
        if s.startswith('#'):
            j = i                  # #If a Or _ / b Then: one directive over several lines
            while re.search(r'\s_\s*$', s) and j + 1 < len(lines):
                j += 1
                s = re.sub(r'\s_\s*$', ' ', s) + lines[j].strip()
            skip = j - i
            low = s.lower()
            body = s[1:].strip()
            lb = body.lower()
            m = re.match(r'^const\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$', body, re.I)
            if m:
                if active:
                    consts[m.group(1).lower()] = evaluate(_no_comment(m.group(2)), consts)
            elif lb.startswith('if ') or lb.startswith('if('):
                cond = re.sub(r'\s+then\s*$', '', _no_comment(body[2:]), flags=re.I)
                take = active and _truth(_eval(cond, consts, ln, filename))
                stack.append((active, take))
                active = take
            elif lb.startswith('elseif'):
                if not stack:
                    raise VBError("#ElseIf without #If", ln, 1, filename)
                parent, taken = stack[-1]
                cond = re.sub(r'\s+then\s*$', '', _no_comment(body[6:]), flags=re.I)
                take = parent and not taken and _truth(_eval(cond, consts, ln, filename))
                stack[-1] = (parent, taken or take)
                active = take
            elif lb.startswith('else'):
                if not stack:
                    raise VBError("#Else without #If", ln, 1, filename)
                parent, taken = stack[-1]
                active = parent and not taken
                stack[-1] = (parent, True)
            elif re.match(r'^end\s*if\b', lb):
                if not stack:
                    raise VBError("#End If without #If", ln, 1, filename)
                active = stack.pop()[0]
            else:
                raise VBError(f"invalid # statement: {s}", ln, 1, filename)
            out.append("")
            continue
        out.append(line if active else "")
        _ = low
    if stack:
        raise VBError("#If without #End If", first_line + len(lines) - 1, 1, filename)
    return out


def _no_comment(t):
    q = False
    for i, ch in enumerate(t):
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            return t[:i]
    return t


def _eval(cond, consts, ln, filename):
    try:
        return evaluate(cond, consts)
    except VBError as e:
        raise VBError(str(e).split(': ', 1)[-1], ln, 1, filename)
