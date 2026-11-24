#!/usr/bin/env python3
# LinBASIC Studio - writes the database of the API Viewer (ide/lbstudio/data/api/LBAPI.TXT): the Declares,
# Constants and Types of the Windows API in the syntax of Visual Basic 6 (32 bits: a pointer is a Long).
# Everything comes from the headers and the import libraries of MinGW-w64 (public domain), read through its
# preprocessor: the procedures that the libraries export with __stdcall (and their DLL), their parameters as the
# headers give them, the numeric and text #defines and enum values, the structures. Our own mapping of the C
# types to the types of VB6; nothing comes from the WIN32API.TXT of Microsoft.
# It also writes ide/lbstudio/data/api/LINUX.TXT: the procedures that the runtime of LinBASIC Studio carries out
# itself on Linux (from runtime/dll.c, runtime/winapi.inc, runtime/forms/*).
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#   python3 tools/mkapi.py [--include DIR] [--lib DIR] [--cc CC] [--check] [--psc DIR]
#     --check     lbvc --check of the whole file as one module (every line must compile)
#     --psc DIR   the Declares of the .bas/.frm/.cls/.ctl files under DIR that the database does not have
import argparse, os, re, struct, subprocess, sys, tempfile, collections

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, '..'))
OUTDIR = os.path.join(ROOT, 'ide', 'lbstudio', 'data', 'api')

# the DLLs, in the order that decides when two libraries export the same name; the name of the Lib clause
DLLS = [('kernel32', 'kernel32'), ('user32', 'user32'), ('gdi32', 'gdi32'), ('advapi32', 'advapi32'),
        ('shell32', 'shell32'), ('comdlg32', 'comdlg32'), ('comctl32', 'comctl32'), ('winmm', 'winmm'),
        ('version', 'version'), ('winspool', 'winspool.drv'), ('shlwapi', 'shlwapi'), ('wininet', 'wininet'),
        ('mpr', 'mpr'), ('msimg32', 'msimg32'), ('imm32', 'imm32'), ('lz32', 'lz32'), ('psapi', 'psapi'),
        ('wsock32', 'wsock32'), ('ole32', 'ole32'), ('oleaut32', 'oleaut32'), ('netapi32', 'netapi32'),
        ('shfolder', 'shfolder'), ('avicap32', 'avicap32'), ('msvfw32', 'msvfw32'), ('uxtheme', 'uxtheme')]

HEADERS = ['windows.h', 'commctrl.h', 'commdlg.h', 'shellapi.h', 'shlobj.h', 'mmsystem.h', 'wininet.h',
           'winspool.h', 'shlwapi.h', 'psapi.h', 'imm.h', 'lzexpand.h', 'winnetwk.h', 'winver.h', 'lm.h',
           'tlhelp32.h', 'richedit.h', 'dbt.h', 'vfw.h', 'uxtheme.h']

# Windows XP (as the programs of VB6 expect the structures: OPENFILENAME, NONCLIENTMETRICS ...); the headers whose
# constants and structures go into the database (the C library and the COM interfaces do not)
SKIP_FILES = re.compile(r'(^_|^psdk_inc/|^sec_api/|intrin|^sal\.h$|^specstrings|^(stdio|stdlib|stddef|stdint|stdarg|'
                        r'wchar|wctype|crtdefs|corecrt\w*|vadefs|math|time|locale|setjmp|signal|float|assert|'
                        r'mem|search|io|fcntl|process|conio|direct|dos|share|tchar|mbstring|mbctype)\.h$|^malloc\.h$|'
                        r'^excpt\.h$|^ctype\.h$|^string\.h$|^errno\.h$|^limits\.h$|^swprintf\.inl$|^stralign\.h$|'
                        r'^(objidl|objidlbase|oaidl|ocidl|unknwn|unknwnbase|wtypes|wtypesbase|propidl|oleidl|urlmon|'
                        r'servprov|msxml|objbase|combaseapi|ole2|oleauto|rpc\w*|shobjidl\w*|shtypes|structuredquery|'
                        r'propsys|propkeydef|ocidl|docobj|exdisp|mshtml\w*|comcat|guiddef|basetyps|cguid|'
                        r'commoncontrols|knownfolders|shlguid|prsht|imagehlp|dbghelp|winsock2?|ws2\w*|'
                        r'mswsock|qos|inaddr|in6addr)\.h$)', re.I)

VB_RESERVED = set('''addressof and any append as binary boolean byref byte byval call case close const currency date
declare defbool defbyte defcur defdate defdbl defint deflng defobj defsng defstr defvar dim do double each else elseif
empty end enum eqv erase error event exit false for friend function get global gosub goto if imp implements in input
integer is let lib like line lock long loop lset me mod new next not nothing null on open option optional or output
paramarray preserve print private property public put raiseevent random read redim rem resume return rset seek
select set single static step stop string sub then to true type typeof until variant wend while with withevents
write xor object'''.split())


# ------------------------------------------------------------------ the import libraries
def ar_symbols(path):
    """the names of the symbol table of an ar archive (the first member "/")"""
    with open(path, 'rb') as f:
        d = f.read()
    if d[:8] != b'!<arch>\n':
        return []
    h = d[8:68]
    if h[:16].strip() != b'/':
        return []
    size = int(h[48:58])
    body = d[68:68 + size]
    n = struct.unpack('>I', body[:4])[0]
    return [s.decode('latin-1') for s in body[4 + 4 * n:].split(b'\0')[:n]]


def exports(libdir):
    """name -> (Lib name, argument bytes) of the __stdcall procedures of the DLLs"""
    out = {}
    for lib, dll in DLLS:
        p = os.path.join(libdir, 'lib%s.a' % lib)
        if not os.path.exists(p):
            print('mkapi: no %s' % p, file=sys.stderr)
            continue
        for s in ar_symbols(p):
            m = re.match(r'^_([A-Za-z]\w*)@(\d+)$', s)
            if m and m.group(1) not in out:
                out[m.group(1)] = (dll, int(m.group(2)))
    return out


# ------------------------------------------------------------------ the preprocessed headers
TOKEN = re.compile(r'\s+|(?P<s>L?"(?:\\.|[^"\\])*")|(?P<c>L?\'(?:\\.|[^\'\\])*\')|'
                   r'(?P<n>0[xX][0-9A-Fa-f]+\w*|\d+\.?\d*(?:[eE][-+]?\d+)?\w*|\.\d+\w*)|(?P<i>[A-Za-z_]\w*)|'
                   r'(?P<p>\.\.\.|<<|>>|->|&&|\|\||[-+*/%&|^~!<>=?:;,.(){}\[\]#])')


def tokens(text):
    out = []
    for m in TOKEN.finditer(text):
        if m.lastgroup:
            out.append(m.group(m.lastgroup))
    return out



def preprocess(cc, incdir):
    src = ''.join('#include <%s>\n' % h for h in HEADERS)
    with tempfile.TemporaryDirectory() as t:
        p = os.path.join(t, 'all.h')
        with open(p, 'w') as f:
            f.write(src)
        cmd = [cc, '-E', '-dD', '-DWINVER=0x0501', '-D_WIN32_WINNT=0x0501', '-D_WIN32_IE=0x0600', '-DNOMINMAX', p]
        if incdir:
            cmd[1:1] = ['-I', incdir]
        r = subprocess.run(cmd, capture_output=True, text=True, errors='replace')
        if r.returncode:
            sys.exit('mkapi: %s failed:\n%s' % (' '.join(cmd), r.stderr[-2000:]))
        return r.stdout



def statements(text):
    """(file, tokens) of the top-level statements; the bodies of inline functions are left out"""
    out = []
    cur = '?'
    toks = []
    tfile = None
    depth = 0
    paren = 0
    skipping = 0
    for line in text.split('\n'):
        if line.startswith('#'):
            m = re.match(r'# \d+ "([^"]*)"', line)
            if m:
                f = m.group(1).replace('\\', '/')
                i = f.find('/include/')
                cur = f[i + 9:] if i >= 0 else os.path.basename(f)
            continue
        for t in tokens(line):
            if skipping:
                if t == '{':
                    skipping += 1
                elif t == '}':
                    skipping -= 1
                    if not skipping:
                        toks = []
                continue
            if not toks:
                tfile = cur
            if t == '{':
                # a function body: "... ) {" at the top level, or after attributes
                if depth == 0 and paren == 0 and toks and is_func_head(toks):
                    skipping = 1
                    continue
                depth += 1
            elif t == '}':
                depth -= 1
            elif t == '(':
                paren += 1
            elif t == ')':
                paren -= 1
            toks.append(t)
            if t == ';' and depth == 0 and paren == 0:
                out.append((tfile, toks))
                toks = []
    return out


def is_func_head(toks):
    t = strip_attrs(toks)
    if not t or t[-1] != ')':
        return False
    if t[0] in ('typedef', 'struct', 'union', 'enum') and '=' not in t:
        # "struct x f(...) {" is still a function; "typedef struct {" is not (the brace is not after ")")
        return t[0] != 'typedef'
    return True


def strip_attrs(toks):
    """the tokens without __attribute__((...)), __declspec(...), __asm__(...) and the storage words"""
    out = []
    i = 0
    n = len(toks)
    while i < n:
        t = toks[i]
        if t in ('__attribute__', '__declspec', '__asm__', '__asm', '__pragma', '_Pragma') and i + 1 < n and toks[i + 1] == '(':
            d = 0
            j = i + 1
            while j < n:
                if toks[j] == '(':
                    d += 1
                elif toks[j] == ')':
                    d -= 1
                    if d == 0:
                        break
                j += 1
            i = j + 1
            continue
        if t in ('extern', 'static', '__inline', '__inline__', 'inline', '__forceinline', '__extension__',
                 '__restrict', '__restrict__', 'restrict', '__unaligned', '__ptr32', '__ptr64', 'register',
                 '__stdcall', '__cdecl', '_cdecl', '__fastcall'):
            i += 1
            continue
        out.append(t)
        i += 1
    return out


def has_stdcall(toks):
    s = ' '.join(toks)
    return '__stdcall__' in s


# ------------------------------------------------------------------ C types
# a type: ('prim', name) | ('name', typedef) | ('ptr', type, const) | ('arr', type, n) | ('func', ret, params, std)
#         | ('su', key) | ('enum',)
PRIMS = {'void', 'char', 'short', 'int', 'long', 'float', 'double', 'signed', 'unsigned', '_Bool', '__int64',
         '__int32', '__int16', '__int8', 'wchar_t', '__int128', 'const', 'volatile'}


class CParser:
    def __init__(self):
        self.typedefs = {}          # name -> type
        self.typedef_file = {}
        self.structs = {}           # key -> {'kind', 'tag', 'members': [(name, type, bits)] or None, 'file'}
        self.anon = 0
        self.funcs = {}             # name -> (ret, params, file)
        self.enum_consts = []       # (file, name, value)
        self.enum_counter = {}

    # -- declaration specifiers
    def specs(self, t, i, file):
        words = []
        base = None
        while i < len(t):
            x = t[i]
            if x in PRIMS:
                words.append(x)
                i += 1
            elif x in ('struct', 'union'):
                i, base = self.struct_spec(t, i, file)
            elif x == 'enum':
                i, base = self.enum_spec(t, i, file)
            elif re.match(r'[A-Za-z_]\w*$', x) and base is None and not [w for w in words if w not in ('const', 'volatile')] \
                    and x in self.typedefs:
                base = ('name', x)
                i += 1
            else:
                break
        if base is None:
            w = [x for x in words if x not in ('const', 'volatile', 'signed')]
            if not w:
                w = ['int'] if words else []
            if not w:
                return i, None
            if 'char' in w:
                base = ('prim', 'char')
            elif 'short' in w or '__int16' in w:
                base = ('prim', 'short')
            elif w.count('long') >= 2 or '__int64' in w:
                base = ('prim', 'int64')
            elif 'double' in w:
                base = ('prim', 'double')
            elif 'float' in w:
                base = ('prim', 'float')
            elif 'void' in w:
                base = ('prim', 'void')
            elif '_Bool' in w or '__int8' in w:
                base = ('prim', 'char')
            elif 'wchar_t' in w:
                base = ('prim', 'short')
            elif '__int128' in w:
                base = ('prim', 'int128')
            else:
                base = ('prim', 'int')
        return i, base

    def struct_spec(self, t, i, file):
        kind = t[i]
        i += 1
        tag = None
        if i < len(t) and re.match(r'[A-Za-z_]\w*$', t[i]):
            tag = t[i]
            i += 1
        key = (kind, tag) if tag else None
        if i < len(t) and t[i] == '{':
            j = match(t, i, '{', '}')
            body = t[i + 1:j]
            if key is None:
                self.anon += 1
                key = (kind, '#%d' % self.anon)
            members = self.members(body, file)
            self.structs[key] = {'kind': kind, 'tag': tag, 'members': members, 'file': file}
            i = j + 1
        elif key is not None and key not in self.structs:
            self.structs[key] = {'kind': kind, 'tag': tag, 'members': None, 'file': file}
        if key is None:
            self.anon += 1
            key = (kind, '#%d' % self.anon)
            self.structs[key] = {'kind': kind, 'tag': None, 'members': None, 'file': file}
        return i, ('su', key)

    def enum_spec(self, t, i, file):
        i += 1
        if i < len(t) and re.match(r'[A-Za-z_]\w*$', t[i]):
            i += 1
        if i < len(t) and t[i] == '{':
            j = match(t, i, '{', '}')
            self.enumerators(t[i + 1:j], file)
            i = j + 1
        return i, ('prim', 'int')

    def enumerators(self, body, file):
        val = -1
        for item in split_commas(body):
            item = strip_attrs(item)
            if not item:
                continue
            name = item[0]
            if len(item) > 2 and item[1] == '=':
                v = evaluate(item[2:], self)
                if v is None:
                    val = None
                    continue
                val = v
            else:
                if val is None:
                    continue
                val += 1
            self.enum_consts.append((file, name, val, False))
            self.enum_counter[name] = val

    def members(self, body, file):
        out = []
        for st in split_semis(body):
            st = strip_attrs(st)
            if not st:
                continue
            i, base = self.specs(st, 0, file)
            if base is None:
                return None
            rest = st[i:]
            if not rest:
                if base[0] == 'su':                 # an anonymous struct or union inside
                    out.append((None, base, None))
                continue
            for d in split_commas(rest):
                bits = None
                if ':' in d:
                    k = d.index(':')
                    bits = d[k + 1:]
                    d = d[:k]
                name, ty = self.declarator(d, base)
                if ty is None:
                    return None
                out.append((name, ty, bits))
        return out

    # -- declarators
    def declarator(self, d, base):
        """(name, type) of a declarator applied to base"""
        d = [x for x in d if x not in ('const', 'volatile', '__restrict', 'restrict')]
        ptrs = 0
        i = 0
        while i < len(d) and d[i] == '*':
            ptrs += 1
            i += 1
        ty = base
        for _ in range(ptrs):
            ty = ('ptr', ty)
        rest = d[i:]
        if not rest:
            return None, ty
        if rest[0] == '(':
            j = match(rest, 0, '(', ')')
            inner = rest[1:j]
            after = rest[j + 1:]
            ty2 = self.suffixes(after, ty)
            if ty2 is None:
                return None, None
            if inner and inner[0] != '*' and all(re.match(r'[A-Za-z_]\w*$', x) for x in inner[:1]) and len(inner) == 1:
                return inner[0], ty2
            return self.declarator(inner, ty2)
        name = None
        if re.match(r'[A-Za-z_]\w*$', rest[0]):
            name = rest[0]
            rest = rest[1:]
        ty = self.suffixes(rest, ty)
        return name, ty

    def suffixes(self, rest, ty):
        if not rest:
            return ty
        if rest[0] == '[':
            j = match(rest, 0, '[', ']')
            n = evaluate(rest[1:j], self) if j > 1 else 0
            inner = self.suffixes(rest[j + 1:], ty)
            if inner is None:
                return None
            return ('arr', inner, n)
        if rest[0] == '(':
            j = match(rest, 0, '(', ')')
            params = self.params(rest[1:j])
            return ('func', ty, params)
        return ty

    def params(self, toks):
        out = []
        toks = strip_attrs(toks)
        if toks in ([], ['void']):
            return out
        for p in split_commas(toks):
            p = strip_attrs(p)
            if p == ['...']:
                out.append(('...', None))
                continue
            i, base = self.specs(p, 0, None)
            if base is None:
                out.append((None, None))
                continue
            name, ty = self.declarator(p[i:], base)
            if ty is not None and ty[0] == 'arr':
                ty = ('ptr', ty[1])
            if ty is not None and ty[0] == 'func':
                ty = ('ptr', ty)
            out.append((name, ty))
        return out

    # -- statements
    def statement(self, file, toks):
        std = has_stdcall(toks)
        t = strip_attrs(toks)
        if not t or t[-1] != ';':
            return
        t = t[:-1]
        typedef = False
        if t and t[0] == 'typedef':
            typedef = True
            t = t[1:]
        i, base = self.specs(t, 0, file)
        if base is None:
            return
        rest = t[i:]
        for d in split_commas(rest):
            if not d:
                continue
            name, ty = self.declarator(d, base)
            if not name or ty is None:
                continue
            if typedef:
                if name not in self.typedefs:
                    self.typedefs[name] = ty
                    self.typedef_file[name] = file
            elif ty[0] == 'func' and std and len(rest) == len(d):
                self.funcs.setdefault(name, (ty[1], ty[2], file))

    # -- resolution
    def resolve(self, ty, seen=None):
        while ty and ty[0] == 'name':
            ty = self.typedefs.get(ty[1])
        return ty


def match(t, i, a, b):
    d = 0
    for j in range(i, len(t)):
        if t[j] == a:
            d += 1
        elif t[j] == b:
            d -= 1
            if d == 0:
                return j
    return len(t) - 1


def split_commas(t):
    out, cur, d = [], [], 0
    for x in t:
        if x in '([{':
            d += 1
        elif x in ')]}':
            d -= 1
        if x == ',' and d == 0:
            out.append(cur)
            cur = []
        else:
            cur.append(x)
    out.append(cur)
    return out


def split_semis(t):
    out, cur, d = [], [], 0
    for x in t:
        if x in '([{':
            d += 1
        elif x in ')]}':
            d -= 1
        cur.append(x)
        if x == ';' and d == 0:
            out.append(cur[:-1])
            cur = []
    if cur:
        out.append(cur)
    return out


# ------------------------------------------------------------------ constant expressions
class Consts:
    def __init__(self, defines, parser):
        self.raw = {}
        for f, n, v in defines:
            self.raw.setdefault(n, (f, v))
        self.parser = parser
        self.memo = {}

    def value(self, name, depth=0):
        if name in self.memo:
            return self.memo[name]
        self.memo[name] = None
        if depth > 40 or name not in self.raw:
            v = self.parser.enum_counter.get(name)
            self.memo[name] = (v, False) if v is not None else None
            return self.memo[name]
        f, text = self.raw[name]
        r = const_value(tokens(text), self, depth)
        self.memo[name] = r
        return r


def const_value(toks, consts, depth=0):
    """(value, hex) of a numeric constant, ('str', text) of a text one, or None"""
    toks = [x for x in toks if x not in ('__MSABI_LONG', '__extension__')]
    if not toks:
        return None
    if all(x.startswith('"') for x in toks):
        s = ''
        for x in toks:
            body = x[1:-1]
            if re.search(r'\\(?![\\"])', body):
                return None
            s += body.replace('\\"', '"').replace('\\\\', '\\')
        if any(ord(c) < 32 or ord(c) > 126 for c in s):
            return None
        return ('str', s)
    hexa = any(re.match(r'0[xX]', x) for x in toks)
    if len(toks) == 1 and re.match(r'[A-Za-z_]\w*$', toks[0]):
        r = consts.value(toks[0], depth + 1) if consts else None
        if r and r[0] == 'str':
            return r
        if r:
            return r
        return None
    hx = [False]
    v = evaluate(toks, consts.parser if consts else None, consts, depth, hx)
    if v is None:
        return None
    return (v, hexa or hx[0])


def evaluate(toks, parser, consts=None, depth=0, hexref=None):
    """the value of a C integer expression, or None (hexref[0]: a constant it uses is written in hex)"""
    if hexref is None:
        hexref = [False]
    out = []
    i = 0
    n = len(toks)
    while i < n:
        x = toks[i]
        if x == '(':
            # a cast: "(" type names "*"... ")" followed by an operand
            j = match(toks, i, '(', ')')
            inner = [y for y in toks[i + 1:j] if y not in ('const', 'volatile')]
            if inner and all(y == '*' or y in PRIMS or (parser and y in parser.typedefs) or y in ('struct', 'union')
                             for y in inner) and j + 1 < n and toks[j + 1] not in (')', '|', '&', '+', '*', '/', '<<', '>>', '^', ',', '?', ':', '%', '==', '!=', '<', '>'):
                if '*' in inner:
                    return None             # a pointer value (HWND_TOPMOST ...): not a constant of VB
                i = j + 1
                continue
            out.append('(')
            i += 1
            continue
        if re.match(r'(0[xX][0-9A-Fa-f]+|\d+)[uUlL]*$', x) or re.match(r'0[xX][0-9A-Fa-f]+(i64|ui64|ll|ull|LL|ULL)$', x):
            m = re.match(r'(0[xX][0-9A-Fa-f]+|\d+)', x)
            s = m.group(1)
            v = int(s, 16) if s[:2].lower() == '0x' else (int(s, 8) if len(s) > 1 and s[0] == '0' else int(s))
            out.append(str(v))
        elif x.startswith("'"):
            body = x[1:-1]
            if len(body) == 1:
                out.append(str(ord(body)))
            else:
                return None
        elif re.match(r'[A-Za-z_]\w*$', x):
            if x == 'sizeof':
                return None
            r = consts.value(x, depth + 1) if consts else None
            if r is not None and r[0] != 'str' and r[1]:
                hexref[0] = True
            if r is None and parser and x in parser.enum_counter:
                r = (parser.enum_counter[x], False)
            if r is None or r[0] == 'str':
                return None
            if r[1]:
                hexref[0] = True
            out.append('(%d)' % r[0])
        elif x in ('|', '&', '^', '~', '<<', '>>', '+', '-', '*', '(', ')'):
            out.append(x)
        elif x == '/':
            out.append('//')
        elif x == '%':
            out.append('%')
        else:
            return None
        i += 1
    expr = ' '.join(out)
    if not expr or not re.match(r'^[\d\s()|&^~<>+\-*/%]+$', expr):
        return None
    try:
        v = eval(expr, {'__builtins__': {}}, {})
    except Exception:
        return None
    if not isinstance(v, int):
        return None
    return v


# ------------------------------------------------------------------ to Visual Basic
class VB:
    def __init__(self, p, exp):
        self.p = p
        self.exp = exp
        self.tnames = {}            # struct key -> VB name
        self.types = {}             # VB name -> lines (None: cannot be written)
        self.order = []
        self.skipped = collections.Counter()
        self.ansi_only = set()

    # the VB name of a struct: its typedef (without the A of an ANSI/Unicode pair), else its tag
    def name_structs(self):
        by_key = collections.defaultdict(list)
        for n, ty in self.p.typedefs.items():
            if ty[0] == 'su':
                by_key[ty[1]].append(n)
            elif ty[0] == 'name':
                r = ty
                while r and r[0] == 'name':
                    r = self.p.typedefs.get(r[1])
        for key, names in by_key.items():
            names = [n for n in names if not n.startswith('_')] or names
            best = None
            for n in names:
                if n.endswith('W') and n[:-1] + 'A' in self.p.typedefs:
                    continue
                best = n
                break
            if best is None:
                continue
            if best.endswith('A') and best[:-1] + 'W' in self.p.typedefs:
                best = best[:-1]
            self.tnames[key] = best
        for key, s in self.p.structs.items():
            if key not in self.tnames and s['tag'] and not s['tag'].startswith('_'):
                self.tnames[key] = s['tag']

    def kind(self, ty):
        """a short description of a resolved type: 'void' 'byte' 'int' 'long' 'cur' 'single' 'double' 'str'
        'wstr' 'pvoid' ('ptr', x) ('su', key) ('arr', ...) 'func' None"""
        r = ty
        while r and r[0] == 'name':
            if r[1] in ('LARGE_INTEGER', 'ULARGE_INTEGER'):
                return 'cur'
            if is_handle_name(r[1]):
                rr = self.p.resolve(r)
                if rr and rr[0] == 'ptr':
                    return 'long'               # HANDLE, HWND, HKEY ...: a number for VB
            r = self.p.typedefs.get(r[1])
        if r is None:
            return None
        if r[0] == 'prim':
            return {'void': 'void', 'char': 'byte', 'short': 'int', 'int': 'long', 'long': 'long', 'int64': 'cur',
                    'float': 'single', 'double': 'double'}.get(r[1])
        if r[0] == 'ptr':
            t = r[1]
            tt = t
            while tt and tt[0] == 'name':
                tt = self.p.typedefs.get(tt[1])
            if tt is None:
                return 'long'
            if tt[0] == 'prim' and tt[1] == 'char':
                return 'str' if self.is_char(t) else 'pvoid'   # LPSTR: a text; LPBYTE: a buffer (As Any)
            if tt[0] == 'prim' and tt[1] == 'short' and t[0] == 'name' and t[1] in ('WCHAR', 'wchar_t', 'OLECHAR'):
                return 'wstr'
            if tt[0] == 'prim' and tt[1] == 'void':
                return 'pvoid'
            if tt[0] == 'func':
                return 'long'
            if tt[0] == 'su' and (tt[1][1] or '').endswith('__'):
                return 'long'                   # DECLARE_HANDLE: a pointer to a structure of nothing
            return ('ptr', t)
        if r[0] == 'su':
            return ('su', r[1])
        if r[0] == 'arr':
            return ('arr', r[1], r[2])
        if r[0] == 'func':
            return 'func'
        return None

    VBT = {'byte': 'Byte', 'int': 'Integer', 'long': 'Long', 'cur': 'Currency', 'single': 'Single',
           'double': 'Double'}

    def struct_vb(self, key):
        """the VB name of a struct that can be written as a Type (written on the way), or None"""
        name = self.tnames.get(key)
        if not name:
            return None
        if name in self.types:
            return name if self.types[name] is not None else None
        self.types[name] = None                 # (a structure that contains itself through a pointer is fine)
        s = self.p.structs.get(key)
        if not s or s['members'] is None or s['kind'] != 'struct' or not s['members']:
            return None
        flat = self.flatten(s['members'])
        if flat is None:
            return None
        lines = []
        used = set()
        for mname, mty in flat:
            k = self.kind(mty)
            vbn = vb_ident(mname)
            if vbn.lower() in used:
                return None
            used.add(vbn.lower())
            if k in self.VBT:
                lines.append('%s As %s' % (vbn, self.VBT[k]))
            elif k == 'str':
                lines.append('%s As String' % vbn)
            elif k in ('wstr', 'pvoid', 'long') or (isinstance(k, tuple) and k[0] == 'ptr'):
                lines.append('%s As Long' % vbn)
            elif isinstance(k, tuple) and k[0] == 'su':
                inner = self.struct_vb(k[1])
                if inner is None:
                    return None
                lines.append('%s As %s' % (vbn, inner))
            elif isinstance(k, tuple) and k[0] == 'arr':
                et, cnt = k[1], k[2]
                if not isinstance(cnt, int) or cnt <= 0:
                    return None
                ek = self.kind(et)
                if ek == 'byte' and self.is_char(et):
                    lines.append('%s As String * %d' % (vbn, cnt))
                elif ek in self.VBT:
                    lines.append('%s(0 To %d) As %s' % (vbn, cnt - 1, self.VBT[ek]))
                elif ek in ('str', 'wstr', 'pvoid', 'long') or (isinstance(ek, tuple) and ek[0] == 'ptr'):
                    lines.append('%s(0 To %d) As Long' % (vbn, cnt - 1))
                elif isinstance(ek, tuple) and ek[0] == 'su':
                    inner = self.struct_vb(ek[1])
                    if inner is None:
                        return None
                    lines.append('%s(0 To %d) As %s' % (vbn, cnt - 1, inner))
                else:
                    return None
            else:
                return None
        self.types[name] = lines
        self.order.append(name)
        return name

    def flatten(self, members):
        """the members as a Type of VB has them: an anonymous structure inline, a union as its largest member,
        the bit fields of one storage unit as one number (fBitFields)"""
        out = []
        nbits = 0
        unit = 0
        group = 0
        for mname, mty, bits in members:
            if bits is not None:
                w = evaluate(bits, self.p)
                sz = self.size(mty)[0]
                if w is None or not sz:
                    return None
                if unit and unit == sz and nbits + w <= sz * 8:
                    nbits += w
                    continue
                group += 1
                unit, nbits = sz, w
                out.append(('fBitFields' if group == 1 else 'fBitFields%d' % group,
                            {1: ('prim', 'char'), 2: ('prim', 'short'), 4: ('prim', 'int'), 8: ('prim', 'int64')}[sz]))
                continue
            unit = 0
            r = self.p.resolve(mty)
            if r and r[0] == 'su':
                s2 = self.p.structs.get(r[1])
                if s2 is None or s2['members'] is None:
                    return None
                if s2['kind'] == 'union':
                    best, bsize = None, -1
                    for v in s2['members']:
                        if v[2] is not None:
                            continue
                        z = self.size(v[1])[0]
                        if z > bsize:
                            best, bsize = v, z
                    if best is None:
                        return None
                    sub = self.flatten([best])
                    if sub is None:
                        return None
                    if mname is not None and best[0] is None:
                        pass                        # (a named union of anonymous structures: their members)
                    out.extend(sub)
                    continue
                if mname is None:                   # an anonymous structure: its members
                    sub = self.flatten(s2['members'])
                    if sub is None:
                        return None
                    out.extend(sub)
                    continue
            if mname is None:
                return None
            out.append((mname, mty))
        return out

    def size(self, ty, depth=0):
        """(size, alignment) in bytes for 32-bit Windows"""
        r = self.p.resolve(ty)
        if r is None or depth > 30:
            return (0, 1)
        if r[0] == 'prim':
            z = {'char': 1, 'short': 2, 'int': 4, 'long': 4, 'int64': 8, 'float': 4, 'double': 8, 'void': 0,
                 'int128': 16}.get(r[1], 4)
            return (z, max(1, min(z, 8)))
        if r[0] in ('ptr', 'func'):
            return (4, 4)
        if r[0] == 'arr':
            z, a = self.size(r[1], depth + 1)
            return ((r[2] or 0) * z if isinstance(r[2], int) else 0, a)
        if r[0] == 'su':
            st = self.p.structs.get(r[1])
            if not st or not st['members']:
                return (0, 1)
            if st['kind'] == 'union':
                zs = [self.size(m[1], depth + 1) for m in st['members']]
                z = max(x[0] for x in zs)
                a = max(x[1] for x in zs)
                return ((z + a - 1) // a * a, a)
            off, al = 0, 1
            for m in st['members']:
                z, a = self.size(m[1], depth + 1)
                off = (off + a - 1) // a * a + z
                al = max(al, a)
            return ((off + al - 1) // al * al, al)
        return (4, 4)

    def is_char(self, ty):
        """char (CHAR, TCHAR in ANSI) and not BYTE / UCHAR: a text"""
        while ty and ty[0] == 'name':
            if ty[1] in ('BYTE', 'UCHAR', 'BOOLEAN', 'UINT8', 'INT8', 'u_char'):
                return False
            if ty[1] in ('CHAR', 'TCHAR', 'CCHAR'):
                return True
            ty = self.p.typedefs.get(ty[1])
        return bool(ty and ty == ('prim', 'char'))

    def param(self, name, ty, idx, out):
        k = self.kind(ty)
        nm = vb_ident(name) if name else 'p%d' % (idx + 1)
        if k in self.VBT:
            out.append('ByVal %s As %s' % (nm, self.VBT[k]))
        elif k == 'str':
            out.append('ByVal %s As String' % nm)
        elif k == 'pvoid':
            out.append('%s As Any' % nm)
        elif k in ('long', 'wstr', 'func'):
            out.append('ByVal %s As Long' % nm)
        elif isinstance(k, tuple) and k[0] == 'ptr':
            pk = self.kind(k[1])
            if pk in self.VBT:
                out.append('%s As %s' % (nm, self.VBT[pk]))
            elif pk in ('long', 'str', 'wstr', 'pvoid', 'func') or (isinstance(pk, tuple) and pk[0] == 'ptr'):
                out.append('%s As Long' % nm)
            elif isinstance(pk, tuple) and pk[0] == 'su':
                tn = self.struct_vb(pk[1])
                out.append('%s As %s' % (nm, tn) if tn else '%s As Any' % nm)
            else:
                out.append('%s As Any' % nm)
        elif isinstance(k, tuple) and k[0] == 'su':
            # a small structure by value (POINT): its two halves, as VB6 programs pass it
            s = self.p.structs.get(k[1])
            mem = s and s['members']
            if mem and len(mem) == 2 and all(m[0] and self.kind(m[1]) == 'long' for m in mem):
                for m in mem:
                    out.append('ByVal %s As Long' % vb_ident('%s%s' % (nm, m[0].capitalize())))
            else:
                return False
        else:
            return False
        return True

    def declare(self, name, ret, params):
        dll, nbytes = self.exp[name]
        rk = self.kind(ret)
        vbname = name
        alias = None
        if name.endswith('A') and name[:-1] + 'W' in self.exp:
            vbname = name[:-1]
            alias = name
        if vbname.lower() in VB_RESERVED:
            alias = alias or name
            vbname = vbname + '_'
        plist = []
        if params and params[-1][0] == '...':
            return None, 'variadic'
        for i, (pn, pt) in enumerate(params):
            if pt is None:
                return None, 'parameter'
            if not self.param(pn, pt, i, plist):
                return None, 'parameter type'
        seen = set()
        for i, pp in enumerate(plist):          # two parameters with the same name (VB: an error)
            nm = pp.split()[1] if pp.startswith('ByVal') else pp.split()[0]
            if nm.lower() in seen:
                plist[i] = pp.replace(' %s ' % nm, ' %s%d ' % (nm, i + 1), 1) if pp.startswith('ByVal') \
                    else '%s%d%s' % (nm, i + 1, pp[len(nm):])
            seen.add(nm.lower())
        size = 0
        for pp in plist:
            size += 8 if pp.startswith('ByVal') and (pp.endswith('Double') or pp.endswith('Currency')) else 4
        if size != nbytes:
            return None, 'argument bytes %d, not %d' % (size, nbytes)
        head = 'Declare Sub' if rk == 'void' else 'Declare Function'
        s = '%s %s Lib "%s"' % (head, vbname, dll)
        if alias:
            s += ' Alias "%s"' % alias
        s += ' (%s)' % ', '.join(plist)
        if rk != 'void':
            if rk in self.VBT:
                s += ' As %s' % self.VBT[rk]
            elif rk in ('str', 'wstr', 'pvoid', 'long', 'func') or (isinstance(rk, tuple) and rk[0] == 'ptr'):
                s += ' As Long'
            else:
                return None, 'return type'
        return (vbname, s), None


EXTRA = [
    ('RtlMoveMemory', 'Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)'),
    ('CopyMemory', 'Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)'),
    ('RtlZeroMemory', 'Declare Sub RtlZeroMemory Lib "kernel32" (Destination As Any, ByVal Length As Long)'),
    ('ZeroMemory', 'Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)'),
    ('RtlFillMemory', 'Declare Sub RtlFillMemory Lib "kernel32" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)'),
    ('FillMemory', 'Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)'),
]


def is_handle_name(n):
    return n in ('HANDLE', 'HGLOBAL', 'HLOCAL', 'HMODULE', 'HINSTANCE', 'HGDIOBJ', 'HCURSOR', 'HICON', 'HWND', 'HDC',
                 'HKEY', 'HMENU', 'HRGN', 'HFONT', 'HBITMAP', 'HBRUSH', 'HPEN', 'HPALETTE', 'HINTERNET', 'HRSRC',
                 'HMETAFILE', 'HENHMETAFILE', 'HACCEL', 'HDROP', 'HIMAGELIST', 'HTREEITEM', 'HPROPSHEETPAGE',
                 'SC_HANDLE', 'HCRYPTPROV') or bool(re.match(r'^H[A-Z0-9_]{2,}$', n))


def vb_ident(n):
    n = n.lstrip('_') or 'x'
    if not re.match(r'[A-Za-z]', n):
        n = 'x' + n
    if n.lower() in VB_RESERVED:
        n += '_'
    return n


def vb_const(v, hexa):
    if not -2 ** 31 <= v < 2 ** 32:
        return None
    u = v & 0xFFFFFFFF
    s = u - 2 ** 32 if u >= 2 ** 31 else u
    if hexa and s >= 0 and s > 9:
        return '&H%X%s' % (s, '&' if 0x8000 <= s <= 0xFFFF else '')
    if hexa and s < 0:
        return '&H%X' % u
    return str(s)


# ------------------------------------------------------------------ Linux: what the runtime carries out itself
def linux_names():
    names = set()
    for rel in ('runtime/dll.c', 'runtime/winapi.inc', 'runtime/forms/vbforms.cpp', 'runtime/forms/winapi_gui.inc'):
        p = os.path.join(ROOT, rel)
        if not os.path.exists(p):
            continue
        text = open(p, encoding='utf-8', errors='replace').read()
        for m in re.finditer(r'\{\s*"\w+"\s*,\s*"([A-Za-z]\w*)"\s*,\s*\w+\s*\}', text):
            names.add(m.group(1))
        for m in re.finditer(r'\{\s*"([A-Za-z]\w*)"\s*,\s*[a-z]_\w+\s*\}', text):
            names.add(m.group(1))
    return names


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description='the database of the API Viewer of LinBASIC Studio')
    ap.add_argument('--include', default=None, help='the include folder of MinGW-w64 (default: the one of the compiler)')
    ap.add_argument('--lib', default=None, help='the lib folder of i686-w64-mingw32 (libuser32.a ...)')
    ap.add_argument('--cc', default='i686-w64-mingw32-gcc')
    ap.add_argument('--out', default=OUTDIR)
    ap.add_argument('--check', action='store_true')
    ap.add_argument('--psc', default=None)
    a = ap.parse_args()
    libdir = a.lib
    if not libdir:
        for c in ('/usr/i686-w64-mingw32/lib', '/usr/lib/gcc/i686-w64-mingw32', '/mingw32/i686-w64-mingw32/lib'):
            if os.path.exists(os.path.join(c, 'libuser32.a')):
                libdir = c
                break
    if not libdir:
        sys.exit('mkapi: the import libraries of i686-w64-mingw32 were not found (--lib)')
    exp = exports(libdir)
    text = preprocess(a.cc, a.include)
    p = CParser()
    for f, toks in statements(text):
        try:
            p.statement(f, toks)
        except (IndexError, RecursionError):
            pass
    defines = []                                # the #defines with their files
    cur = '?'
    for line in text.split('\n'):
        if line.startswith('# '):
            m = re.match(r'# \d+ "([^"]*)"', line)
            if m:
                f = m.group(1).replace('\\', '/')
                i = f.find('/include/')
                cur = f[i + 9:] if i >= 0 else os.path.basename(f)
        elif line.startswith('#define '):
            m = re.match(r'#define\s+([A-Za-z_]\w*)(\(?)(.*)$', line)
            if m and not m.group(2):
                defines.append((cur, m.group(1), m.group(3).strip()))
    vb = VB(p, exp)
    vb.name_structs()

    # Declares: the exported procedures that the headers declare (ANSI; the W twins are left out)
    decls = {}
    reasons = collections.Counter()
    for name in sorted(exp):
        if name.endswith('W') and name[:-1] + 'A' in exp:
            continue
        if name not in p.funcs:
            reasons['not in the headers'] += 1
            continue
        ret, params, f = p.funcs[name]
        r, why = vb.declare(name, ret, params)
        if r is None:
            reasons[why.split(' ')[0] if why.startswith('argument') else why] += 1
            continue
        if r[0].lower() not in decls:
            decls[r[0].lower()] = r

    # well-known procedures that the headers give only as macros of the C library (RtlMoveMemory: memmove)
    for vbname, line in EXTRA:
        dll_entry = re.search(r'Alias "(\w+)"', line)
        entry = dll_entry.group(1) if dll_entry else vbname
        if entry in exp and vbname.lower() not in decls:
            decls[vbname.lower()] = (vbname, line)

    # Types: the structures of the API headers, and those the Declares use
    for key, s in list(p.structs.items()):
        if s['kind'] == 'struct' and s['members'] and not SKIP_FILES.search(os.path.basename(s['file'] or '?')) \
                and not (s['file'] or '').startswith(('psdk_inc', 'sec_api')):
            n = vb.tnames.get(key)
            if n and not n.startswith('_') and not (n.endswith('W') and n[:-1] + 'A' in p.typedefs):
                vb.struct_vb(key)

    # Constants
    consts = Consts(defines, p)
    cvals = {}
    names_used = {n.lower() for n in vb.order} | set(decls)
    for f, n, v in defines:
        if n.startswith('_') or SKIP_FILES.search(os.path.basename(f)) or f.startswith(('psdk_inc', 'sec_api')) \
                or '/' in f and not f.startswith(('ddk/',)) and f.split('/')[0] in ('gl', 'GL', 'KHR', 'directx', 'gdiplus'):
            continue
        if n.lower() in VB_RESERVED or n.lower() in names_used or n.lower() in cvals:
            continue
        if re.match(r'^(WINAPI|CALLBACK|APIENTRY|STDMETHOD|DECLSPEC|DECLARE|WINBASEAPI|WINUSERAPI|NTSYSAPI|PURE|THIS|'
                    r'FAR|NEAR|IN|OUT|OPTIONAL|CONST|VOID|TRUE|FALSE|NULL)$', n):
            continue
        r = consts.value(n)
        if r is None:
            continue
        if r[0] == 'str':
            cvals[n.lower()] = (n, '"%s"' % r[1].replace('"', '""'))
        else:
            s = vb_const(r[0], r[1])
            if s is not None:
                cvals[n.lower()] = (n, s)
    for f, n, v, _h in p.enum_consts:
        if n.startswith('_') or SKIP_FILES.search(os.path.basename(f or '?')) or n.lower() in VB_RESERVED \
                or n.lower() in cvals or n.lower() in names_used:
            continue
        s = vb_const(v, False)
        if s is not None:
            cvals[n.lower()] = (n, s)

    # the file
    os.makedirs(a.out, exist_ok=True)
    out = []
    out.append("' LBAPI.TXT - the Windows API for Visual Basic 6 (32 bits), the database of the API Viewer of")
    out.append("' LinBASIC Studio. Generated by tools/mkapi.py from the headers and the import libraries of")
    out.append("' MinGW-w64 (public domain); do not edit. Declares: %d, Constants: %d, Types: %d."
               % (len(decls), len(cvals), len(vb.order)))
    out.append('')
    out.append("' ---------------------------------------------------------------- Types")
    for n in vb.order:
        out.append('Public Type %s' % n)
        for l in vb.types[n]:
            out.append('    ' + l)
        out.append('End Type')
        out.append('')
    out.append("' ---------------------------------------------------------------- Constants")
    for k in sorted(cvals):
        n, v = cvals[k]
        out.append('Public Const %s = %s' % (n, v))
    out.append('')
    out.append("' ---------------------------------------------------------------- Declares")
    for k in sorted(decls):
        out.append('Public ' + decls[k][1])
    out.append('')
    path = os.path.join(a.out, 'LBAPI.TXT')
    with open(path, 'w', newline='\r\n', encoding='ascii', errors='replace') as fh:
        fh.write('\n'.join(out))
    lin = linux_names()
    with open(os.path.join(a.out, 'LINUX.TXT'), 'w', newline='\r\n') as fh:
        fh.write("' LINUX.TXT - the Windows API procedures that the runtime of LinBASIC Studio carries out itself on\n")
        fh.write("' Linux (generated by tools/mkapi.py from runtime/dll.c, runtime/winapi.inc, runtime/forms)\n")
        for n in sorted(lin, key=str.lower):
            fh.write(n + '\n')
    print('mkapi: %s' % path)
    print('  Declares %d, Constants %d, Types %d; procedures of the libraries left out: %s'
          % (len(decls), len(cvals), len(vb.order), ', '.join('%s %d' % kv for kv in reasons.most_common())))
    print('  Linux: %d procedures of the runtime' % len(lin))

    if a.psc:
        psc_report(a.psc, decls)
    if a.check:
        return check(path)
    return 0


def psc_report(folder, decls):
    used = collections.Counter()
    for dp, dn, fn in os.walk(folder):
        for f in fn:
            if f.lower().endswith(('.bas', '.frm', '.cls', '.ctl')):
                try:
                    t = open(os.path.join(dp, f), encoding='latin-1').read()
                except OSError:
                    continue
                for m in re.finditer(r'Declare\s+(?:Function|Sub)\s+(\w+)\s+Lib\s+"([^"]+)"(?:\s+Alias\s+"([^"]+)")?',
                                     t, re.I):
                    lib = re.sub(r'\.dll$', '', m.group(2).lower())
                    entry = m.group(3) or m.group(1)
                    used[(lib, entry.lower())] += 1
    have = set()
    for k, (vbname, s) in decls.items():
        m = re.search(r'Lib "([^"]+)"(?: Alias "([^"]+)")?', s)
        lib = m.group(1).lower()
        entry = (m.group(2) or vbname).lower()
        have.add((lib, entry))
        have.add((lib, entry.rstrip('_')))
    win = {d for _l, d in DLLS} | {'winspool.drv'}
    tot = sum(c for (l, e), c in used.items() if l in win)
    miss = [((l, e), c) for (l, e), c in used.items() if l in win and (l, e) not in have and (l, e + 'a') not in have]
    got = tot - sum(c for _k, c in miss)
    print('  PSC: %d Declares of the Windows DLLs in %s, %d (%.1f%%) in the database' % (tot, folder, got, 100.0 * got / max(tot, 1)))
    for (l, e), c in sorted(miss, key=lambda x: -x[1])[:60]:
        print('    missing %5d  %s %s' % (c, l, e))


def check(path):
    lbvc = os.path.join(ROOT, 'compiler', 'lbvc.py')
    with tempfile.TemporaryDirectory() as t:
        bas = os.path.join(t, 'LBAPI.bas')
        with open(path, encoding='ascii') as f:
            body = f.read()
        with open(bas, 'w', newline='\r\n') as f:
            f.write('Attribute VB_Name = "LBAPI"\n' + body.replace('\r\n', '\n'))
        r = subprocess.run([sys.executable, lbvc, '--check', bas], capture_output=True, text=True)
        print('  lbvc --check: %s' % ('ok' if r.returncode == 0 else 'FAILED'))
        if r.returncode:
            print((r.stdout + r.stderr)[-6000:])
        return r.returncode


if __name__ == '__main__':
    sys.exit(main())
