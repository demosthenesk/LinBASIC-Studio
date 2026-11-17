# LinBASIC Studio - lbvc lexer: VB6 source text -> tokens
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
import re


class VBError(Exception):
    def __init__(self, msg, line=0, col=0, file=None):
        super().__init__(msg)
        self.msg, self.line, self.col, self.file = msg, line, col, file

    def __str__(self):
        return f"{self.file or ''}:{self.line}:{self.col}: error: {self.msg}"


KEYWORDS = set("""
ADDRESSOF AND ANY APPEND AS BINARY BOOLEAN BYREF BYTE BYVAL CALL CASE CLOSE CONST CURRENCY DATE DECLARE
DEFBOOL DEFBYTE DEFCUR DEFDATE DEFDBL DEFINT DEFLNG DEFOBJ DEFSNG DEFSTR DEFVAR DIM DO DOUBLE EACH ELSE ELSEIF
EMPTY END ENUM EQV ERASE ERROR EVENT EXIT FALSE FOR FRIEND FUNCTION GET GLOBAL GOSUB GOTO IF IMP IMPLEMENTS
IN INPUT INTEGER IS LET LIB LIKE LINE LOCK LONG LOOP LSET ME MOD NEW NEXT NOT NOTHING NULL ON OPEN OPTION
OPTIONAL OR OUTPUT PARAMARRAY PRESERVE PRINT PRIVATE PROPERTY PUBLIC PUT RAISEEVENT RANDOM READ REDIM REM
RESUME RETURN RSET SEEK SELECT SET SINGLE STATIC STEP STOP STRING SUB THEN TO TRUE TYPE TYPEOF UNTIL
VARIANT WEND WHILE WITH WITHEVENTS WRITE XOR
""".split())
# words that are keywords only in some places and otherwise ordinary names
SOFT = {"BINARY", "TEXT", "COMPARE", "EXPLICIT", "BASE", "MODULE", "ACCESS", "SHARED", "LEN", "OBJECT",
        "INPUT", "OUTPUT", "APPEND", "RANDOM", "READ", "WRITE", "LINE", "ERROR", "STRING", "DATE", "PRINT",
        "GET", "PUT", "SEEK", "LOCK", "RESET", "NAME", "STEP", "TO"}

SUFFIX_TYPE = {'%': 'Integer', '&': 'Long', '!': 'Single', '#': 'Double', '@': 'Currency', '$': 'String'}


class Tok:
    __slots__ = ("kind", "value", "line", "col", "suffix", "text", "space_before")

    def __init__(self, kind, value, line, col, suffix=None, text=None, space_before=False):
        self.kind = kind            # 'num' 'str' 'date' 'id' 'kw' 'op' 'nl' 'eof'
        self.value = value          # number value / string / upper-case keyword or op / identifier
        self.line, self.col = line, col
        self.suffix = suffix        # type suffix of identifiers and numbers
        self.text = text            # original spelling (identifiers)
        self.space_before = space_before

    def is_kw(self, *names):
        return self.kind == 'kw' and self.value in names

    def is_op(self, *ops):
        return self.kind == 'op' and self.value in ops

    def is_word(self, *names):
        """keyword or identifier with this (upper-case) name"""
        return self.kind in ('kw', 'id') and self.value.upper() in names

    def __repr__(self):
        return f"Tok({self.kind},{self.value!r},{self.line}:{self.col})"


_NUM = re.compile(r'(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?([%&!#@]?)')
_HEX = re.compile(r'&([hH][0-9A-Fa-f]+|[oO][0-7]+|[0-7]+)(&?)')
_ID = re.compile(r'[A-Za-z_\u0080-￿][A-Za-z0-9_\u0080-￿]*')
_DATE = re.compile(r'#([^#\r\n]+)#')
FILE_HASH_AFTER = {"PRINT", "WRITE", "INPUT", "GET", "PUT", "AS", "CLOSE", "LOCK", "UNLOCK", "SEEK", "WIDTH",
                   "RESET", "LINE"}


def number_type(is_float, suffix, value):
    if suffix:
        return SUFFIX_TYPE[suffix]
    if is_float:
        return 'Double'
    if -32768 <= value <= 32767:
        return 'Integer'
    if -2147483648 <= value <= 2147483647:
        return 'Long'
    return 'Double'


def tokenize(text, filename=None):
    """Tokens of a VB6 module (code part only).  Line continuations ' _' are joined."""
    toks = []
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    i = 0
    while i < len(lines):
        start_line = i + 1
        line = lines[i]
        # join continuation lines: a line ending in ' _' (outside strings / comments)
        while _continues(line) and i + 1 < len(lines):
            head, nxt = line.rstrip()[:-1], lines[i + 1]
            lw = re.search(r'([A-Za-z_][A-Za-z0-9_]*)\s*$', head)
            if nxt.lstrip().startswith('.') and re.search(r'[A-Za-z0-9_)\]]\s*$', head) and not (
                    lw and lw.group(1).upper() in KEYWORDS):
                line = head.rstrip() + nxt.lstrip()      # .IpAddress _ / .IpAddressString: one member chain
            else:
                line = head + ' ' + nxt
            i += 1
        _tok_line(line, start_line, toks, filename)
        toks.append(Tok('nl', '\n', start_line, len(line) + 1))
        i += 1
    toks.append(Tok('eof', None, len(lines) + 1, 1))
    return toks


def _continues(line):
    s = line.rstrip()
    if not s.endswith('_') or len(s) < 2 or s[-2] not in ' \t':
        return False
    q = False
    for ch in s[:-1]:
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            return True                         # a comment that ends in ' _' goes on to the next line, as in VB6
    if not q and re.match(r'\s*(\d+\s+)?rem\b', s, re.I):
        return True
    return not q


def _tok_line(s, ln, toks, filename):
    p, n = 0, len(s)
    first = len(toks)
    while p < n:
        ch = s[p]
        if ch in ' \t':
            p += 1
            continue
        col = p + 1
        sp = p > 0 and s[p - 1] in ' \t'
        if ch == "'":
            break
        if ch == '"':
            q = p + 1
            buf = []
            while True:
                if q >= n:
                    raise VBError("expected: \"", ln, col, filename)
                if s[q] == '"':
                    if q + 1 < n and s[q + 1] == '"':
                        buf.append('"')
                        q += 2
                        continue
                    break
                buf.append(s[q])
                q += 1
            toks.append(Tok('str', ''.join(buf), ln, col, space_before=sp))
            p = q + 1
            continue
        if ch == '&':
            m = _HEX.match(s, p)
            if m:
                body = m.group(1)
                if body[0] in 'hH':
                    v = int(body[1:], 16)
                    bits = 16 if (not m.group(2) and v <= 0xFFFF) else 32
                else:
                    v = int(body[1:] if body[0] in 'oO' else body, 8)
                    bits = 16 if (not m.group(2) and v <= 0xFFFF) else 32
                if bits == 16:
                    if v >= 0x8000:
                        v -= 0x10000
                    t = 'Integer'
                else:
                    v &= 0xFFFFFFFF
                    if v >= 0x80000000:
                        v -= 0x100000000
                    t = 'Long'
                toks.append(Tok('num', v, ln, col, suffix=t, space_before=sp))
                p = m.end()
                continue
        if ch.isdigit() or (ch == '.' and p + 1 < n and s[p + 1].isdigit()):
            # a '.' right after a name or ')' is member access, not a number
            if ch == '.' and toks and len(toks) > first and toks[-1].kind in ('id', 'kw') and not sp:
                pass
            else:
                m = _NUM.match(s, p)
                txt, exp, suf = m.group(1), m.group(2), m.group(3)
                is_float = '.' in txt or bool(exp)
                if suf == '#' and m.end() < n and s[m.end()] not in ' \t,;):+-*/\\^=<>&':
                    suf = ''          # e.g. a date literal follows, not a suffix
                    end = m.end() - 1
                else:
                    end = m.end()
                lit = txt + (exp.replace('d', 'e').replace('D', 'e') if exp else '')
                # (922337203685477@: an integer stays exact; a float has only 53 bits)
                v = float(lit) if is_float or suf in ('!', '#') else int(lit)
                if suf == '@' and is_float:   # 922337203685477.5807@: exact (a Fraction), not 53 bits
                    from fractions import Fraction
                    v = Fraction(lit)
                if not is_float and not suf and v > 2147483647:
                    v = float(v)
                t = number_type(is_float, suf, v)
                if t in ('Integer', 'Long') and isinstance(v, float):
                    v = int(v)
                toks.append(Tok('num', v, ln, col, suffix=t, space_before=sp))
                p = end
                continue
        if ch == '#':
            prev = toks[-1] if len(toks) > first else None
            head = None                               # first token of the current statement
            for t in reversed(toks[first:]):
                if t.is_op(':'):
                    break
                head = t
            fileish = prev is not None and (prev.kind in ('kw', 'id') and prev.value.upper() in FILE_HASH_AFTER
                                            or (prev.is_op(',') and head is not None and head.is_word('CLOSE')))
            if not fileish:
                m = _DATE.match(s, p)
                if m and re.search(r'[/:\-]|\b(AM|PM)\b', m.group(1), re.I):
                    toks.append(Tok('date', m.group(1).strip(), ln, col, space_before=sp))
                    p = m.end()
                    continue
            toks.append(Tok('op', '#', ln, col, space_before=sp))
            p += 1
            continue
        if ch == '[':
            q = s.find(']', p)
            if q < 0:
                raise VBError("expected: ]", ln, col, filename)
            toks.append(Tok('id', s[p + 1:q], ln, col, text=s[p + 1:q], space_before=sp))
            p = q + 1
            continue
        m = _ID.match(s, p)
        if m:
            word = m.group(0)
            p = m.end()
            suffix = None
            if p < n and s[p] in '%&!#@$' and not (s[p] == '&' and p + 1 < n and s[p + 1] in 'hHoO'):
                # 'a& b' is concatenation only when written with a space; 'a&' is a suffix
                if not (s[p] == '!' and p + 1 < n and (s[p + 1].isalpha() or s[p + 1] == '[')):
                    suffix = s[p]
                    p += 1
            up = word.upper()
            if up == 'REM' and not suffix and not (toks and toks[-1].kind == 'op' and toks[-1].value in ('!', '.')
                                                   and not sp):
                break                              # (rs!Rem, x.Rem: a name, not a comment)
            if up in KEYWORDS and not suffix:
                toks.append(Tok('kw', up, ln, col, text=word, space_before=sp))
            elif up in ('MID', 'LEFT', 'RIGHT', 'CHR', 'STR', 'UCASE', 'LCASE', 'TRIM', 'LTRIM', 'RTRIM',
                        'SPACE', 'STRING', 'HEX', 'OCT', 'INPUT', 'DATE', 'TIME', 'ERROR', 'COMMAND',
                        'ENVIRON', 'CURDIR', 'DIR', 'FORMAT', 'CHRB', 'CHRW', 'MIDB', 'LEFTB', 'RIGHTB') and suffix == '$':
                toks.append(Tok('id', word, ln, col, suffix='$', text=word, space_before=sp))
            else:
                toks.append(Tok('id', word, ln, col, suffix=suffix, text=word, space_before=sp))
            continue
        two = s[p:p + 2]
        if two in ('<>', '<=', '>=', '=<', '=>', '><', ':='):
            v = {'=<': '<=', '=>': '>=', '><': '<>'}.get(two, two)
            toks.append(Tok('op', v, ln, col, space_before=sp))
            p += 2
            continue
        if ch in '+-*/\\^&=<>(),.:;!':
            toks.append(Tok('op', ch, ln, col, space_before=sp))
            p += 1
            continue
        raise VBError(f"invalid character {ch!r}", ln, col, filename)
