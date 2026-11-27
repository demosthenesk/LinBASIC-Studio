# LinBASIC Studio - lbvc: the VBA functions and statements, and the Err / App / Debug objects
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""Each builtin compiles its arguments with the ProcGen and returns an E (C code + VB type).
Functions named like Left$ return String; without the $ VB returns a Variant holding a
String, which here is a String too (Null arguments are not supported yet)."""
import json, os
from .lexer import VBError


def _E():
    from .cgen import E
    return E


_VBA_PARAMS = None


def vba_params(name):
    """the parameter names of a function or statement of the VBA library (MsgBox: Prompt, Buttons, ...)"""
    global _VBA_PARAMS
    if _VBA_PARAMS is None:
        import json, os
        _VBA_PARAMS = {}
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'data', 'vb6', 'typelibs', 'VBA.json')
        try:
            for t in json.load(open(p, encoding='utf-8'))['types']:
                for m in t.get('members', []):
                    ps = [x['name'].lower() for x in m.get('params', []) if 'retval' not in (x.get('flags') or [])]
                    _VBA_PARAMS.setdefault(m['name'].lower().lstrip('_'), ps)
        except (OSError, ValueError, KeyError):
            pass
    return _VBA_PARAMS.get(name.lower().rstrip('$'))


def _args(pg, node, args, lo, hi, name):
    args = list(args)
    if any(a is not None and a.kind == 'named' for a in args):
        # MsgBox Prompt:=s, Title:=t: the arguments in the order of the VBA library
        order = vba_params(name)
        if not order:
            pg.err(f"named arguments are not supported for {name} yet", node)
        out = []
        for a in args:
            if a is not None and a.kind == 'named':
                break
            out.append(a)
        for a in args[len(out):]:
            if a is None or a.kind != 'named':
                pg.err("named arguments must come after the others", node)
            if a.name.lower() not in order:
                pg.err(f"named argument not found: {a.name}", node)
            i = order.index(a.name.lower())
            while len(out) <= i:
                out.append(None)
            if out[i] is not None:
                pg.err(f"argument {a.name} given twice", node)
            out[i] = a.e
        args = out
    while args and args[-1] is None:
        args.pop()
    if len(args) < lo or len(args) > hi:
        pg.err(f"wrong number of arguments: {name}", node)
    return args + [None] * (hi - len(args))


def _x(pg, a, t, node):
    return pg.conv(pg.expr(a), t, node)


def _opt(pg, a, t, node, default):
    return default if a is None else _x(pg, a, t, node)


FUNCS = {}


def func(*names):
    def deco(f):
        for n in names:
            FUNCS[n.lower()] = f
        return f
    return deco


def simple(name, params, rt, cfn, dollar=None):
    """a function whose arguments are converted to the given types"""
    def h(pg, node, args):
        E = _E()
        a = _args(pg, node, args, len(params), len(params), name)
        return E(f"{cfn}({', '.join(_x(pg, x, t, node) for x, t in zip(a, params))})", rt)
    FUNCS[name.lower()] = h
    if dollar:
        FUNCS[name.lower() + '$'] = h


for _n, _p, _r, _c in [
        ('Len', None, None, None),
        ('LCase', ['String'], 'String', 'vb_lcase'), ('UCase', ['String'], 'String', 'vb_ucase'),
        ('LTrim', ['String'], 'String', 'vb_ltrim'), ('RTrim', ['String'], 'String', 'vb_rtrim'),
        ('Trim', ['String'], 'String', 'vb_trim'), ('Space', ['Long'], 'String', 'vb_space'),
        ('Chr', ['Long'], 'String', 'vb_chr'), ('ChrW', ['Long'], 'String', 'vb_chrw'),
        ('Asc', ['String'], 'Integer', 'vb_asc'), ('AscW', ['String'], 'Integer', 'vb_ascw'),
        ('StrReverse', ['String'], 'String', 'vb_strreverse'),
        ('Val', ['String'], 'Double', 'vb_val'),
        ('Sqr', ['Double'], 'Double', 'vb_sqr'), ('Log', ['Double'], 'Double', 'vb_log'),
        ('Exp', ['Double'], 'Double', 'vb_exp'), ('Sin', ['Double'], 'Double', 'vb_sin'),
        ('Cos', ['Double'], 'Double', 'vb_cos'), ('Tan', ['Double'], 'Double', 'vb_tan'),
        ('Atn', ['Double'], 'Double', 'vb_atn'),
        ('DateSerial', ['Long', 'Long', 'Long'], 'Date', 'vb_dateserial'),
        ('TimeSerial', ['Long', 'Long', 'Long'], 'Date', 'vb_timeserial'),
        ('Year', ['Date'], 'Integer', 'vb_year'), ('Month', ['Date'], 'Integer', 'vb_month'),
        ('Day', ['Date'], 'Integer', 'vb_day'), ('Hour', ['Date'], 'Integer', 'vb_hour'),
        ('Minute', ['Date'], 'Integer', 'vb_minute'), ('Second', ['Date'], 'Integer', 'vb_second'),
        ('EOF', ['Long'], 'Boolean', 'vb_eof'), ('LOF', ['Long'], 'Long', 'vb_lof'),
        ('Loc', ['Long'], 'Long', 'vb_loc'), ('Seek', ['Long'], 'Long', 'vb_seekpos'),
        ('FileLen', ['String'], 'Long', 'vb_filelen'),
        ('VarType', ['Variant'], 'Long', 'vb_vartype'), ('TypeName', ['Variant'], 'String', 'vb_typename')]:
    if _p is not None:
        simple(_n, _p, _r, _c, dollar=_r == 'String')


@func('len')
def f_len(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'Len')[0]
    x = pg.expr(a)
    if x.ptr is None and x.code is not None and x.t not in ('String', 'Variant') and _vba_variant_call(pg, a):
        # Len(Hour(Now)), Len(TimeSerial(..)): a function of VBA that gives a Variant (typed here): the length of
        # its text, as of a Variant (LedClock, UXC)
        return len_of(pg, _E()(pg.conv(x, 'Variant', a), 'Variant'), node)
    return len_of(pg, x, node)


def _vba_variant_call(pg, a):
    """a call of a function of the VBA library that gives a Variant (not a procedure of the program)"""
    from .symbols import vba_variant_funcs
    base = a.base if a.kind == 'apply' else a
    if base.kind != 'name' or getattr(base, 'suffix', None) is not None:
        return False
    if base.name.lower() not in vba_variant_funcs():
        return False
    return pg._lookup_quiet(base) is None


def len_of(pg, x, node):
    E = _E()
    if x.t == 'Variant':
        return E(f"(vb_i32)vb_len(vb_v2s({x.code}))", 'Long')
    if x.t == 'String':
        return E(f"vb_len({x.code})", 'Long')
    from .symbols import is_udt, is_obj
    if is_udt(x.t):
        return E(f"((vb_i32){pg.c.udt_len(x.t)})", 'Long')
    if is_obj(x.t):                     # Len(Text1): the default property (Text)
        return E(f"(vb_i32)vb_len(vb_v2s(vb_late_o({x.code}, \"\", VBM_GET, 0, 0)))", 'Long')
    if x.ptr is None:                    # Len(12345): VB6 wants a variable (the size of its type)
        pg.err("variable required - can't assign to this expression", node)
    size = {'Integer': 2, 'Long': 4, 'Single': 4, 'Double': 8, 'Currency': 8, 'Boolean': 2, 'Byte': 1, 'Date': 8}
    return E(f"((vb_i32){size[x.t]})", 'Long')


@func('lenb')
def f_lenb(pg, node, args):
    """LenB: the bytes of a string (two in each character); of other values as Len"""
    a = _args(pg, node, args, 1, 1, 'LenB')[0]
    x = pg.expr(a)
    if x.t == 'String':
        return _E()(f"vb_lenb({x.code})", 'Long')
    if x.t == 'Variant':
        return _E()(f"vb_lenb(vb_v2s({x.code}))", 'Long')
    from .symbols import is_udt
    if is_udt(x.t):                     # a Type: its size in memory (Unicode strings, aligned to 4)
        return _E()(f"((vb_i32){pg.c.udt_lenb(x.t)})", 'Long')
    return len_of(pg, x, node)


@func('varptr')
def f_varptr(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'VarPtr')[0]
    x = pg.expr(a)
    from .symbols import ctype, is_udt, is_obj
    if x.ptr is None and x.code is not None and x.arr is None and not is_udt(x.t):
        # VarPtr(104), VarPtr(AllocStr(n)): VB6 passes a temporary copy, whose address it gives
        tmp = pg.tmp(x.t)
        pg.pre.append(pg.store_code(_E()(tmp, x.t, ptr=f"&{tmp}"), x, node) if not is_obj(x.t)
                      else f"vb_oset(&{tmp}, {x.code});")
        pg.free_after(tmp, x.t, pg.cleanup)
        x = _E()(tmp, x.t, ptr=f"&{tmp}")
    if x.ptr is None:
        pg.err("VarPtr needs a variable", node)
    if ctype(x.t) == 'vbs':                    # a String: CopyMemory of its 4 bytes copies the String (dll.c)
        return _E()(f"vb_varptr_str({x.ptr})", 'Long')
    if getattr(x, 'aref', None):              # VarPtr(Vert(0)) for an API that reads the array from there
        return _E()(f"vb_varptr_elem({x.ptr}, (vb_i32)sizeof({ctype(x.t)}), {x.aref})", 'Long')
    return _E()(f"vb_varptr({x.ptr}, (vb_i32)sizeof({ctype(x.t)}))", 'Long')


@func('strptr')
def f_strptr(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'StrPtr')[0]
    return _E()(f"vb_strptr({_x(pg, a, 'String', node)})", 'Long')


@func('objptr')
def f_objptr(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'ObjPtr')[0]
    return _E()(f"vb_varptr(vb_oid({_x(pg, a, 'Object', node)}), (vb_i32)sizeof(vbobjh))", 'Long')


@func('left', 'left$')
def f_left(pg, node, args):
    a = _args(pg, node, args, 2, 2, 'Left')
    return _E()(f"vb_left({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Long', node)})", 'String')


@func('right', 'right$')
def f_right(pg, node, args):
    a = _args(pg, node, args, 2, 2, 'Right')
    return _E()(f"vb_right({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Long', node)})", 'String')


@func('mid', 'mid$')
def f_mid(pg, node, args):
    a = _args(pg, node, args, 2, 3, 'Mid')
    return _E()(f"vb_mid({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Long', node)}, "
                f"{_opt(pg, a[2], 'Long', node, '-1')})", 'String')


@func('instr')
def f_instr(pg, node, args):
    a = _args(pg, node, args, 2, 4, 'InStr')
    if a[2] is None:                       # InStr(string1, string2)
        start, s1, s2, cmp_ = '1', a[0], a[1], None
    else:
        start, s1, s2, cmp_ = _x(pg, a[0], 'Long', node), a[1], a[2], a[3]
    c = _opt(pg, cmp_, 'Long', node, str(pg.compare_text))
    return _E()(f"vb_instr({start}, {_x(pg, s1, 'String', node)}, {_x(pg, s2, 'String', node)}, {c})", 'Long')


@func('instrrev')
def f_instrrev(pg, node, args):
    a = _args(pg, node, args, 2, 4, 'InStrRev')
    return _E()(f"vb_instrrev({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'String', node)}, "
                f"{_opt(pg, a[2], 'Long', node, '-1')}, {_opt(pg, a[3], 'Long', node, str(pg.compare_text))})", 'Long')


@func('string', 'string$')
def f_string(pg, node, args):
    a = _args(pg, node, args, 2, 2, 'String')
    ch = pg.expr(a[1])
    n = _x(pg, a[0], 'Long', node)
    if ch.t == 'String':
        return _E()(f"vb_string_s({n}, {ch.code})", 'String')
    return _E()(f"vb_string_n({n}, {pg.conv(ch, 'Long', node)})", 'String')


# the byte functions: a String is a sequence of bytes for them (an odd number too: VBS_ODD in the runtime)
@func('ascb')
def f_ascb(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'AscB')
    return _E()(f"((vb_u8)vb_ascb({_x(pg, a[0], 'String', node)}))", 'Byte')   # (VB6: a Byte)


@func('chrb', 'chrb$')
def f_chrb(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'ChrB')
    return _E()(f"vb_chrb({_x(pg, a[0], 'Long', node)})", 'String')


@func('leftb', 'leftb$')
def f_leftb(pg, node, args):
    a = _args(pg, node, args, 2, 2, 'LeftB')
    return _E()(f"vb_leftb({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Long', node)})", 'String')


@func('rightb', 'rightb$')
def f_rightb(pg, node, args):
    a = _args(pg, node, args, 2, 2, 'RightB')
    return _E()(f"vb_rightb({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Long', node)})", 'String')


@func('midb', 'midb$')
def f_midb(pg, node, args):
    a = _args(pg, node, args, 2, 3, 'MidB')
    return _E()(f"vb_midb({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Long', node)}, "
                f"{_opt(pg, a[2], 'Long', node, '-1')})", 'String')


@func('instrb')
def f_instrb(pg, node, args):
    a = _args(pg, node, args, 2, 4, 'InStrB')
    if a[2] is None:
        start, s1, s2, cmp_ = '1', a[0], a[1], None
    else:
        start, s1, s2, cmp_ = _x(pg, a[0], 'Long', node), a[1], a[2], a[3]
    c = _opt(pg, cmp_, 'Long', node, str(pg.compare_text))
    return _E()(f"vb_instrb({start}, {_x(pg, s1, 'String', node)}, {_x(pg, s2, 'String', node)}, {c})", 'Long')


@func('cverr')
def f_cverr(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'CVErr')
    return _E()(f"vb_cverr({_x(pg, a[0], 'Long', node)})", 'Variant')


@func('iserror')
def f_iserror(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'IsError')
    return _E()(f"((vb_bool)(vb_vartype({_x(pg, a[0], 'Variant', node)}) == 10 ? -1 : 0))", 'Boolean')


def _fin(name, n_req, params):
    def h(pg, node, args):
        a = _args(pg, node, args, n_req, len(params), name)
        xs = [_opt(pg, x, 'Double', node, '0') if t == 'd' else _opt(pg, x, 'Long', node, '0') for x, t in zip(a, params)]
        return _E()(f"vb_{name.lower()}({', '.join(xs)})", 'Double')
    FUNCS[name.lower()] = h


for _n, _r, _p in [('Pmt', 3, 'dddds'), ('FV', 3, 'dddds'), ('PV', 3, 'dddds'), ('NPer', 3, 'dddds'),
                   ('IPmt', 4, 'ddddds'), ('PPmt', 4, 'ddddds')]:
    _fin(_n, _r, _p)


@func('sln', 'syd', 'ddb')
def f_depreciation(pg, node, args):
    name = node.name.lower()
    n = {'sln': (3, 3), 'syd': (4, 4), 'ddb': (4, 5)}[name]
    a = _args(pg, node, args, n[0], n[1], node.name)
    xs = [_x(pg, x, 'Double', node) for x in a[:n[0]]]
    if name == 'ddb':
        xs.append(_opt(pg, a[4], 'Double', node, '2.0'))
    return _E()(f"vb_{name}({', '.join(xs)})", 'Double')


@func('rate')
def f_rate(pg, node, args):
    a = _args(pg, node, args, 3, 6, 'Rate')
    return _E()(f"vb_rate({_x(pg, a[0], 'Double', node)}, {_x(pg, a[1], 'Double', node)}, "
                f"{_x(pg, a[2], 'Double', node)}, {_opt(pg, a[3], 'Double', node, '0.0')}, "
                f"{_opt(pg, a[4], 'Long', node, '0')}, {_opt(pg, a[5], 'Double', node, '0.1')})", 'Double')


@func('npv')
def f_npv(pg, node, args):
    a = _args(pg, node, args, 2, 2, 'NPV')
    rate = _x(pg, a[0], 'Double', node)
    ref, vt = _arr(pg, a[1], node)
    return _E()(f"vb_npv({rate}, {ref}, {vt})", 'Double')


@func('irr')
def f_irr(pg, node, args):
    a = _args(pg, node, args, 1, 2, 'IRR')
    ref, vt = _arr(pg, a[0], node)
    return _E()(f"vb_irr({ref}, {vt}, {_opt(pg, a[1], 'Double', node, '0.1')})", 'Double')


@func('mirr')
def f_mirr(pg, node, args):
    a = _args(pg, node, args, 3, 3, 'MIRR')
    ref, vt = _arr(pg, a[0], node)
    return _E()(f"vb_mirr({ref}, {vt}, {_x(pg, a[1], 'Double', node)}, {_x(pg, a[2], 'Double', node)})", 'Double')


@func('partition')
def f_partition(pg, node, args):
    a = _args(pg, node, args, 4, 4, 'Partition')
    return _E()(f"vb_partition({', '.join(_x(pg, x, 'Long', node) for x in a)})", 'String')


@func('fileattr')
def f_fileattr(pg, node, args):
    a = _args(pg, node, args, 2, 2, 'FileAttr')
    return _E()(f"vb_fileattr({_x(pg, a[0], 'Long', node)}, {_x(pg, a[1], 'Long', node)})", 'Long')


@func('replace')
def f_replace(pg, node, args):
    a = _args(pg, node, args, 3, 6, 'Replace')
    return _E()(f"vb_replace({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'String', node)}, "
                f"{_x(pg, a[2], 'String', node)}, {_opt(pg, a[3], 'Long', node, '1')}, "
                f"{_opt(pg, a[4], 'Long', node, '-1')}, {_opt(pg, a[5], 'Long', node, str(pg.compare_text))})", 'String')


@func('strcomp')
def f_strcomp(pg, node, args):
    a = _args(pg, node, args, 2, 3, 'StrComp')
    return _E()(f"(vb_i16)vb_strcomp({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'String', node)}, "
                f"{_opt(pg, a[2], 'Long', node, str(pg.compare_text))})", 'Integer')


@func('str', 'str$')
def f_str(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'Str')[0]
    x = pg.expr(a)
    fn = {'Integer': 'vb_str_i', 'Long': 'vb_str_i', 'Byte': 'vb_str_i', 'Boolean': 'vb_str_i',
          'Single': 'vb_str_f', 'Currency': 'vb_str_c', 'Variant': 'vb_str_v'}.get(x.t)
    if fn:
        return _E()(f"{fn}({x.code})", 'String')
    return _E()(f"vb_str_d({pg.conv(x, 'Double', node)})", 'String')


@func('hex', 'hex$', 'oct', 'oct$')
def f_hex(pg, node, args):
    name = node.name.lower().rstrip('$')
    a = _args(pg, node, args, 1, 1, name)[0]
    x = pg.expr(a)
    small = x.t in ('Integer', 'Byte', 'Boolean')
    v = pg.conv(x, 'Long', node)
    return _E()(f"vb_{name}((vb_i64){v}, {2 if small else 4})", 'String')


def conversion(name, t):
    def h(pg, node, args):
        a = _args(pg, node, args, 1, 1, name)[0]
        x = pg.expr(a)
        if x.t == t:
            return x.__class__(pg.conv(x, t, node), t)    # (an extended value: rounded)
        if t == 'Single' and x.t in ('Double', 'Date') or (t == 'Single' and getattr(x, 'ext', False)):
            # CSng: error 6 when too big; P-code stores Inf first (vb_rf), native code does not (vb_csng)
            c = f"({x.code})" if getattr(x, 'ext', False) else f"((vb_ld){pg.conv(x, 'Double', node)})"
            return _E()(f"vb_rf({c})" if not pg.native else f"vb_csng((double){c})", 'Single')
        return _E()(pg.conv(x, t, node), t)
    FUNCS[name.lower()] = h


for _n, _t in [('CStr', 'String'), ('CInt', 'Integer'), ('CLng', 'Long'), ('CSng', 'Single'),
               ('CDbl', 'Double'), ('CCur', 'Currency'), ('CBool', 'Boolean'), ('CByte', 'Byte'),
               ('CVar', 'Variant'), ('CDate', 'Date'), ('CVDate', 'Date')]:
    conversion(_n, _t)


@func('cdec')
def f_cdec(pg, node, args):
    """CDec: a Variant of subtype Decimal (96-bit integer and a scale: decimal.c)"""
    a = _args(pg, node, args, 1, 1, 'CDec')[0]
    return _E()(f"vb_cdec({_x(pg, a, 'Variant', node)})", 'Variant')


@func('fix', 'int', 'abs', 'sgn')
def f_numfun(pg, node, args):
    name = node.name.lower()
    a = _args(pg, node, args, 1, 1, name)[0]
    x = pg.expr(a)
    E = _E()
    from .symbols import is_obj
    if x.t == 'Variant':                   # the same subtype (a Decimal stays a Decimal), Null stays Null
        if name == 'sgn':
            return E(f"vb_v_sgn({x.code})", 'Integer')
        return E(f"vb_v_intfix({x.code}, {({'int': 0, 'fix': 1, 'abs': 2})[name]})", 'Variant')
    if is_obj(x.t):                        # Int(Text1): the value of the default property
        x = E(pg.conv(x, 'Double', node), 'Double')
    if x.t == 'String':
        x = E(pg.conv(x, 'Double', node), 'Double')
    if name == 'sgn':
        return E(f"((vb_i16)vb_sgn((double)({pg.conv(x, 'Double') if x.t == 'Currency' else x.code})))", 'Integer')
    if x.t in ('Integer', 'Long', 'Byte', 'Boolean'):
        t = 'Integer' if x.t == 'Boolean' else x.t
        if name == 'abs' and t != 'Byte':
            w = 'w' if pg.noovf else ''          # Remove Integer Overflow Checks: Abs(-32768) = -32768
            return E(f"{'vb_abs_i' if t == 'Integer' else 'vb_abs_l'}{w}({x.code})", t)
        return E(x.code, t)
    if x.t == 'Currency':
        d = f"vb_cy2d({x.code})"
        f = {'fix': 'vb_fix', 'int': 'vb_int', 'abs': 'fabs'}[name]
        return E(f"vb_ccur({f}({d}))", 'Currency')
    f = {'fix': 'vb_fix', 'int': 'vb_int', 'abs': 'fabs'}[name]
    if name == 'abs':
        return E(f"fabs({x.code})", x.t)
    return E(f"{f}({x.code})", 'Single' if x.t == 'Single' else 'Double' if x.t != 'Date' else 'Date')


@func('round')
def f_round(pg, node, args):
    a = _args(pg, node, args, 1, 2, 'Round')
    x = pg.expr(a[0])
    if x.t == 'Variant':                   # Round of a Decimal: exact
        return _E()(f"vb_v_round({x.code}, {_opt(pg, a[1], 'Long', node, '0')})", 'Variant')
    return _E()(f"vb_round({pg.conv(x, 'Double', node)}, {_opt(pg, a[1], 'Long', node, '0')})", 'Double')


@func('rnd')
def f_rnd(pg, node, args):
    a = _args(pg, node, args, 0, 1, 'Rnd')
    if a[0] is None:
        return _E()("((float)vb_rnd(1, 0))", 'Single')
    return _E()(f"((float)vb_rnd({_x(pg, a[0], 'Double', node)}, 1))", 'Single')


@func('timer')
def f_timer(pg, node, args):
    _args(pg, node, args, 0, 0, 'Timer')
    return _E()("((float)vb_timer())", 'Single')


@func('now')
def f_now(pg, node, args):
    _args(pg, node, args, 0, 0, 'Now')
    return _E()("vb_v_date(vb_now())", 'Variant')


@func('date', 'date$')
def f_date(pg, node, args):
    _args(pg, node, args, 0, 0, 'Date')
    if node.suffix == '$':                     # Date$: mm-dd-yyyy
        return _E()(f"vb_format(vb_v_date((vb_date)(vb_i32)vb_now()), {pg.c.strlit('mm-dd-yyyy')}, 1, 1, 1)", 'String')
    return _E()("vb_v_date((vb_date)(vb_i32)vb_now())", 'Variant')   # VBA: a Variant (Date)


@func('time', 'time$')
def f_time(pg, node, args):
    _args(pg, node, args, 0, 0, 'Time')
    if node.suffix == '$':                     # Time$: hh:mm:ss
        return _E()(f"vb_format(vb_v_date(vb_now()), {pg.c.strlit('hh:mm:ss')}, 1, 1, 1)", 'String')
    return _E()("vb_v_date(vb_now() - (vb_i32)vb_now())", 'Variant')


@func('weekday')
def f_weekday(pg, node, args):
    a = _args(pg, node, args, 1, 2, 'Weekday')
    return _E()(f"(vb_i16)vb_weekday({_x(pg, a[0], 'Date', node)}, {_opt(pg, a[1], 'Long', node, '1')})", 'Integer')


@func('isnumeric')
def f_isnumeric(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'IsNumeric')[0]
    return _E()(f"vb_isnumeric({_x(pg, a, 'Variant', node)})", 'Boolean')


@func('isempty', 'isnull', 'ismissing')
def f_isx(pg, node, args):
    name = node.name.lower()
    a = _args(pg, node, args, 1, 1, name)[0]
    vt = {'isempty': 'VT_EMPTY', 'isnull': 'VT_NULL', 'ismissing': 'VT_ERROR'}[name]
    x = pg.expr(a)
    if x.t != 'Variant':
        return _E()("((vb_bool)0)", 'Boolean')
    if name != 'ismissing':                 # IsNull(rs!Name): the default value of an object (VB6)
        return _E()(f"((vb_bool)-(vb_vartype_def({x.code}) == {vt}))", 'Boolean')
    return _E()(f"((vb_bool)-(vb_vartype({x.code}) == {vt}))", 'Boolean')


def _all(pg, items):
    """IIf, Choose and Switch evaluate ALL their arguments (VB6), from left to right: [(code, C type)] ->
    the codes, kept in temporaries when an argument calls something"""
    if all(n is None or pg.pure(n) for n, _ in items):
        return [c for _, c in items]
    out = []
    for n, c in items:
        t = pg.tmp('Variant' if c[1] == 'v' else 'Long')
        pg.pre.append(f"{t} = {c[0]};")
        out.append((t, c[1]))
    return out


@func('iif')
def f_iif(pg, node, args):
    a = _args(pg, node, args, 3, 3, 'IIf')
    c, x, y = [k[0] for k in _all(pg, [(a[0], (pg.cond(a[0]), 'l')), (a[1], (_x(pg, a[1], 'Variant', node), 'v')),
                                       (a[2], (_x(pg, a[2], 'Variant', node), 'v'))])]
    return _E()(f"(({c}) ? {x} : {y})", 'Variant')


@func('lbound', 'ubound')
def f_bound(pg, node, args):
    name = node.name.lower()
    a = _args(pg, node, args, 1, 2, name)
    ref, _ = _arr(pg, a[0], node)
    return _E()(f"vb_{name}({ref}, {_opt(pg, a[1], 'Long', node, '1')})", 'Long')


@func('isobject')
def f_isobject(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'IsObject')[0]
    x = pg.expr(a)
    from .symbols import is_obj
    if is_obj(x.t):
        return _E()("((vb_bool)-1)", 'Boolean')
    if x.t != 'Variant':
        return _E()("((vb_bool)0)", 'Boolean')
    return _E()(f"vb_isobject({x.code})", 'Boolean')


@func('msgbox')
def f_msgbox(pg, node, args):
    a = _args(pg, node, args, 1, 5, 'MsgBox')
    title = _opt(pg, a[2], 'String', node, None) or 'vb_stmp(0)'
    if (len(a) > 3 and a[3] is not None) or (len(a) > 4 and a[4] is not None):   # HelpFile, Context: the Help button
        hf = _opt(pg, a[3], 'String', node, None) if len(a) > 3 else None
        ctx = _opt(pg, a[4], 'Long', node, '0') if len(a) > 4 else '0'
        return _E()(f"(vb_i16)vb_msgbox_help({_x(pg, a[0], 'String', node)}, {_opt(pg, a[1], 'Long', node, '0')}, "
                    f"{title}, {hf or 'vb_stmp(0)'}, {ctx})", 'Integer')
    return _E()(f"(vb_i16)vb_msgbox({_x(pg, a[0], 'String', node)}, {_opt(pg, a[1], 'Long', node, '0')}, {title})",
                'Integer')


@func('inputbox', 'inputbox$')
def f_inputbox(pg, node, args):
    a = _args(pg, node, args, 1, 7, 'InputBox')
    t = _opt(pg, a[1], 'String', node, None) or 'vb_stmp(0)'
    d = _opt(pg, a[2], 'String', node, None) or 'vb_stmp(0)'
    return _E()(f"vb_inputbox({_x(pg, a[0], 'String', node)}, {t}, {d})", 'String')


@func('input', 'input$', 'inputb', 'inputb$')
def f_input(pg, node, args):
    """Input(n, #f): n characters of the file; InputB(n, #f): n bytes as a byte string"""
    a = _args(pg, node, args, 2, 2, node.name)
    f = a[1].e if a[1] is not None and a[1].kind == 'filenum' else a[1]
    fn = 'vb_inputb' if node.name.lower().startswith('inputb') else 'vb_input_n'
    return _E()(f"{fn}({_x(pg, a[0], 'Long', node)}, {_x(pg, f, 'Long', node)})", 'String')


@func('freefile')
def f_freefile(pg, node, args):
    _args(pg, node, args, 0, 1, 'FreeFile')
    return _E()("(vb_i16)vb_freefile()", 'Integer')


@func('dir', 'dir$')
def f_dir(pg, node, args):
    a = _args(pg, node, args, 0, 2, 'Dir')
    if a[0] is None:
        return _E()("vb_dir(vb_stmp(0), 0, 0)", 'String')
    return _E()(f"vb_dir({_x(pg, a[0], 'String', node)}, {_opt(pg, a[1], 'Long', node, '0')}, 1)", 'String')


@func('curdir', 'curdir$')
def f_curdir(pg, node, args):
    a = _args(pg, node, args, 0, 1, 'CurDir')
    if a and a[0] is not None:
        return _E()(f"vb_curdir_d({_x(pg, a[0], 'String', node)})", 'String')
    return _E()("vb_curdir()", 'String')


@func('command', 'command$')
def f_command(pg, node, args):
    _args(pg, node, args, 0, 0, 'Command')
    return _E()("vb_command()", 'String')


@func('doevents')
def f_doevents(pg, node, args):
    _args(pg, node, args, 0, 0, 'DoEvents')
    return _E()("(vb_doevents(), (vb_i16)0)", 'Integer')


@func('erl')
def f_erl(pg, node, args):
    _args(pg, node, args, 0, 0, 'Erl')
    return _E()("vb_erl()", 'Long')


@func('error', 'error$')
def f_error(pg, node, args):
    """Error(n): the text of the run-time error n; Error: the one of Err"""
    a = _args(pg, node, args, 0, 1, 'Error')
    if not a or a[0] is None:
        return _E()("vb_err_description()", 'String')
    return _E()(f"vb_error_str({_x(pg, a[0], 'Long', node)})", 'String')


@func('isdate')
def f_isdate(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'IsDate')[0]
    return _E()(f"vb_isdate({_x(pg, a, 'Variant', node)})", 'Boolean')


@func('imestatus')
def f_imestatus(pg, node, args):
    _args(pg, node, args, 0, 0, 'IMEStatus')
    return _E()("((vb_i32)0)", 'Long')          # vbIMENoOp: no input method on this system


@func('getsetting', 'getsetting$')
def f_getsetting(pg, node, args):
    a = _args(pg, node, args, 3, 4, 'GetSetting')
    s = [_x(pg, x, 'String', node) for x in a[:3]]
    return _E()(f"vb_getsetting({s[0]}, {s[1]}, {s[2]}, {_opt(pg, a[3], 'String', node, 'vb_stmp(0)')})", 'String')


@func('getallsettings')
def f_getallsettings(pg, node, args):
    a = _args(pg, node, args, 2, 2, 'GetAllSettings')
    return _E()(f"vb_getallsettings({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'String', node)})", 'Variant')


@func('filedatetime')
def f_filedatetime(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'FileDateTime')
    return _E()(f"vb_v_date(vb_filedatetime({_x(pg, a[0], 'String', node)}))", 'Variant')


@func('getattr')
def f_getattr(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'GetAttr')
    return _E()(f"(vb_i16)vb_getattr({_x(pg, a[0], 'String', node)})", 'Integer')


@func('shell')
def f_shell(pg, node, args):
    a = _args(pg, node, args, 1, 2, 'Shell')
    return _E()(f"vb_v_d(vb_shell({_x(pg, a[0], 'String', node)}, {_opt(pg, a[1], 'Long', node, '2')}))", 'Variant')


@func('environ', 'environ$')
def f_environ(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'Environ')
    x = pg.expr(a[0])
    if x.t == 'String':
        return _E()(f"vb_environ_s({x.code})", 'String')
    if x.t == 'Variant':
        v = x.code
        return _E()(f"(vb_vartype({v}) == VT_BSTR ? vb_environ_s(vb_v2s({v})) : vb_environ_n(vb_v2l({v})))", 'String')
    return _E()(f"vb_environ_n({_x(pg, a[0], 'Long', node)})", 'String')


def lookup(name, suffix=None):
    key = name.lower() + ('$' if suffix == '$' else '')
    f = FUNCS.get(key)
    if f is None and suffix == '$':             # Replace$, LCase$ ...: VB6 takes the $ on its String functions
        f = FUNCS.get(name.lower())
    return f


# ---------------------------------------------------------------- statements that look like calls
STMTS = {}


def stmt(*names):
    def deco(f):
        for n in names:
            STMTS[n.lower()] = f
        return f
    return deco


def _one_path(fn):
    def h(pg, st, args):
        a = _args(pg, st, args, 1, 1, fn)
        pg.emit(f"vb_{fn}({_x(pg, a[0], 'String', st)});")
    return h


for _n in ('kill', 'mkdir', 'rmdir', 'chdir'):
    STMTS[_n] = _one_path(_n)
STMTS['chdrive'] = _one_path('chdrive')


@stmt('setattr')
def s_setattr(pg, st, args):
    a = _args(pg, st, args, 2, 2, 'SetAttr')
    pg.emit(f"vb_setattr({_x(pg, a[0], 'String', st)}, {_x(pg, a[1], 'Long', st)});")


@stmt('appactivate')
def s_appactivate(pg, st, args):
    a = _args(pg, st, args, 1, 2, 'AppActivate')
    pg.emit(f"vb_appactivate({_x(pg, a[0], 'String', st)}, {_opt(pg, a[1], 'Boolean', st, '0')});")


@stmt('reset')
def s_reset(pg, st, args):
    _args(pg, st, args, 0, 0, 'Reset')
    pg.emit("vb_close(0);")


@stmt('name')
def s_name(pg, st, args):
    pg.emit(f"vb_name({_x(pg, args[0], 'String', st)}, {_x(pg, args[1], 'String', st)});")


@stmt('savesetting')
def s_savesetting(pg, st, args):
    a = _args(pg, st, args, 4, 4, 'SaveSetting')
    s = [_x(pg, x, 'String', st) for x in a]
    pg.emit(f"vb_savesetting({s[0]}, {s[1]}, {s[2]}, {s[3]});")


@stmt('deletesetting')
def s_deletesetting(pg, st, args):
    a = _args(pg, st, args, 1, 3, 'DeleteSetting')
    if a[1] is None:                                 # DeleteSetting "App": the whole application
        pg.emit(f"vb_deletesetting({_x(pg, a[0], 'String', st)}, vb_stmp(0), vb_stmp(0), -1);")
        return
    k = _opt(pg, a[2], 'String', st, None)
    pg.emit(f"vb_deletesetting({_x(pg, a[0], 'String', st)}, {_x(pg, a[1], 'String', st)}, "
            f"{k or 'vb_stmp(0)'}, {1 if k else 0});")


@stmt('filecopy')
def s_filecopy(pg, st, args):
    a = _args(pg, st, args, 2, 2, 'FileCopy')
    pg.emit(f"vb_filecopy({_x(pg, a[0], 'String', st)}, {_x(pg, a[1], 'String', st)});")


@stmt('beep')
def s_beep(pg, st, args):
    _args(pg, st, args, 0, 0, 'Beep')
    pg.emit("vb_beep();")


@stmt('lbaddin_hostrun')
def s_lbaddin_hostrun(pg, st, args):
    """the host of an add-in (runtime/addin.c): talks with the IDE until it unloads the add-in (phase 12.5; the
    start of an add-in made as a program, designers.addin_host_source)"""
    a = _args(pg, st, args, 1, 1, 'LbAddIn_HostRun')
    pg.emit(f"vb_addin_run({_x(pg, a[0], 'Object', st)});")


@stmt('doevents')
def s_doevents(pg, st, args):
    pg.emit("vb_doevents();")


@stmt('sendkeys')
def s_sendkeys(pg, st, args):
    a = _args(pg, st, args, 1, 2, 'SendKeys')
    pg.emit(f"vb_sendkeys({_x(pg, a[0], 'String', st)}, {_opt(pg, a[1], 'Boolean', st, '0')});")


@func('rgb')
def f_rgb(pg, node, args):
    a = _args(pg, node, args, 3, 3, 'RGB')
    return _E()(f"vb_rgb({_x(pg, a[0], 'Long', node)}, {_x(pg, a[1], 'Long', node)}, {_x(pg, a[2], 'Long', node)})", 'Long')


@func('qbcolor')
def f_qbcolor(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'QBColor')
    return _E()(f"vb_qbcolor({_x(pg, a[0], 'Long', node)})", 'Long')


def lookup_stmt(name):
    return STMTS.get(name.lower())


# ---------------------------------------------------------------- Err, App
def lookup_member(obj, name):
    if obj == 'err':
        m = {'number': ("vb_err_number()", 'Long'), 'description': ("vb_err_description()", 'String'),
             'source': ("vb_err_source()", 'String'), 'lastdllerror': ("vb_lastdllerr", 'Long'),
             'helpfile': ("vb_err_helpfile()", 'String'), 'helpcontext': ("vb_err_helpcontext()", 'Long')}.get(name)
        if m:
            return lambda pg, node, args: _E()(*m)
    if obj == 'app':
        m = {'path': ("vb_app_path()", 'String'), 'previnstance': ("((vb_bool)0)", 'Boolean'),
             'taskvisible': ("((vb_bool)-1)", 'Boolean'), 'startmode': ("((vb_i16)0)", 'Integer'),
             'threadid': ("((vb_i32)1)", 'Long'), 'hinstance': ("((vb_i32)0x400000)", 'Long'),
             'logmode': ("((vb_i32)0)", 'Long'), 'unattendedapp': ("((vb_bool)0)", 'Boolean')}.get(name)
        if m:
            return lambda pg, node, args: _E()(*m)
        if name in ('major', 'minor', 'revision'):
            return lambda pg, node, args: _E()(f"((vb_i16){int(pg.c.app.get(name, 1 if name == 'major' else 0))})", 'Integer')
        if name in ('title', 'exename', 'comments', 'companyname', 'filedescription', 'legalcopyright',
                    'legaltrademarks', 'productname', 'helpfile'):
            def f(pg, node, args):
                v = pg.c.app.get(name)
                if v is None and name == 'title' and pg.c.project_name:
                    v = pg.c.project_name               # VB6: the Name of the project (as Err.Source)
                if v is None:
                    d = "vb_app_exename()" if name in ('title', 'exename') else "vb_stmp(0)"
                else:
                    d = pg.c.strlit(v)
                if name in ('title', 'helpfile'):     # App.Title = ... changes them at run time
                    return _E()(f"vb_app_str({1 if name == 'title' else 2}, {d})", 'String')
                return _E()(d, 'String')
            return f
    return None


def lookup_member_set(obj, name):
    if obj == 'app' and name in ('title', 'helpfile'):
        k = 1 if name == 'title' else 2
        return lambda pg, st, value: pg.emit(f"vb_app_set_str({k}, {pg.conv(value, 'String', st)});")
    if obj == 'app' and name == 'taskvisible':
        return lambda pg, st, value: pg.emit(f"(void)({pg.conv(value, 'Boolean', st)});")
    if obj == 'err':
        fn = {'number': ('vb_err_set_number', 'Long'), 'description': ('vb_err_set_description', 'String'),
              'source': ('vb_err_set_source', 'String'), 'helpcontext': ('vb_err_set_helpcontext', 'Long'),
              'helpfile': ('vb_err_set_helpfile', 'String')}.get(name)
        if fn:
            return lambda pg, st, value: pg.emit(f"{fn[0]}({pg.conv(value, fn[1], st)});")
    return None


def lookup_method(obj, name):
    if obj == 'err' and name == 'raise':
        def h(pg, st, args):
            pos = [a for a in args if a is None or a.kind != 'named']
            named = {a.name.lower(): a.e for a in args if a is not None and a.kind == 'named'}
            vals = {}
            for i, key in enumerate(('number', 'source', 'description', 'helpfile', 'helpcontext')):
                v = pos[i] if i < len(pos) else named.get(key)
                vals[key] = v
            if vals['number'] is None:
                pg.err("argument not optional: Number", st)
            n = _x(pg, vals['number'], 'Long', st)
            src = _x(pg, vals['source'], 'String', st) if vals['source'] is not None else 'vb_stmp(0)'
            desc = _x(pg, vals['description'], 'String', st) if vals['description'] is not None else 'vb_stmp(0)'
            ctx = _x(pg, vals['helpcontext'], 'Long', st) if vals['helpcontext'] is not None else '0'
            flags = (1 if vals['source'] is not None else 0) | (2 if vals['description'] is not None else 0) | \
                (4 if vals['helpcontext'] is not None else 0)
            if vals['helpfile'] is not None:    # Err.HelpFile (the Help button of a MsgBox of the error)
                hf = _x(pg, vals['helpfile'], 'String', st)
                pg.emit(f"vb_err_raise_hf({n}, {src}, {desc}, {hf}, {ctx}, {flags | 8});")
            else:
                pg.emit(f"vb_err_raise({n}, {src}, {desc}, {ctx}, {flags});")
        return h
    if obj == 'err' and name == 'clear':
        return lambda pg, st, args: pg.emit("vb_err_clear(); vb_lastdllerr = 0;")   # (On Error keeps LastDllError)
    if obj == 'debug' and name == 'assert':     # a made program leaves Debug.Assert out (VB6 too)
        return lambda pg, st, args: None
    if obj == 'app' and name in ('logevent', 'startlogging'):   # the event log of Windows NT: the arguments only
        def h(pg, st, args):
            for x in args:
                if x is not None:
                    pg.emit(f"(void)({_x(pg, x.e if x.kind == 'named' else x, 'Variant', st)});")
        return h
    return None


# names of the VBA library (for 'not supported yet' messages)
def _known():
    here = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(here, '..', '..', 'data', 'vb6', 'typelibs', 'VBA.json')
    out = set()
    if os.path.exists(p):
        for t in json.load(open(p, encoding='utf-8'))['types']:
            if t['kind'] == 'module':
                for m in t.get('members', []):
                    n = m['name']
                    if n.startswith('_B_'):
                        n = n.split('_', 3)[3]
                    out.add(n.lower())
    return out


KNOWN_VBA = _known()


# ---------------------------------------------------------------- arrays as values
def _arr(pg, a, node):
    """(C pointer to the vbarr, C expression of the element VarType) of an array argument:
    an array variable, an array inside a Type, a function result, or a Variant holding one"""
    from .symbols import is_arr, vt_of
    x = pg.expr(a)
    if x.arr is not None and x.code is None:
        et = x.t[2:] if is_arr(x.t) else x.t
        return x.arr, vt_of(et)
    if x.t == 'Variant':
        t = pg.tmp('Variant')
        pg.pre.append(f"vb_vset(&{t}, {x.code});")
        pg.post.append(f"vb_vfree(&{t});")
        return f"vb_v_arrptr({t})", f"({t}.vt & 0xFFF)"
    pg.err("type mismatch: an array is needed", node)


def _arr_result(pg, call, et):
    t = pg.tmp('A:' + et)
    pg.pre.append(f"{t} = {call};")
    pg.post.append(f"vb_afree(&{t});")
    return _E()(None, 'A:' + et, arr=f"&{t}")


@func('split')
def f_split(pg, node, args):
    a = _args(pg, node, args, 1, 4, 'Split')
    d = _x(pg, a[1], 'String', node) if a[1] is not None else 'vb_stmp(0)'
    return _arr_result(pg, f"vb_split({_x(pg, a[0], 'String', node)}, {d}, {1 if a[1] is not None else 0}, "
                           f"{_opt(pg, a[2], 'Long', node, '-1')}, {_opt(pg, a[3], 'Long', node, str(pg.compare_text))})",
                       'String')


@func('join')
def f_join(pg, node, args):
    a = _args(pg, node, args, 1, 2, 'Join')
    ref, vt = _arr(pg, a[0], node)
    d = _x(pg, a[1], 'String', node) if a[1] is not None else 'vb_stmp(0)'
    return _E()(f"vb_join({ref}, {vt}, {d}, {1 if a[1] is not None else 0})", 'String')


@func('filter')
def f_filter(pg, node, args):
    a = _args(pg, node, args, 2, 4, 'Filter')
    ref, vt = _arr(pg, a[0], node)
    return _arr_result(pg, f"vb_filter({ref}, {vt}, {_x(pg, a[1], 'String', node)}, "
                           f"{_opt(pg, a[2], 'Boolean', node, '-1')}, {_opt(pg, a[3], 'Long', node, str(pg.compare_text))})",
                       'String')


@func('array')
def f_array(pg, node, args):
    vals = [_x(pg, a, 'Variant', node) if a is not None else 'vb_v_missing()' for a in args]
    if not vals:
        return _E()("vb_array(0, 0)", 'Variant')
    return _E()(f"vb_array({len(vals)}, (vbvar[]){{ {', '.join(vals)} }})", 'Variant')


@func('isarray')
def f_isarray(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'IsArray')[0]
    x = pg.expr(a)
    if x.arr is not None and x.code is None:
        return _E()("((vb_bool)-1)", 'Boolean')
    if x.t == 'Variant':
        return _E()(f"vb_isarray({x.code})", 'Boolean')
    return _E()("((vb_bool)0)", 'Boolean')


@func('choose')
def f_choose(pg, node, args):
    if len(args) < 2:
        pg.err("argument not optional: Choose", node)
    i = pg.tmp('Long')
    pg.pre.append(f"{i} = {_x(pg, args[0], 'Long', node)};")
    vals = [k[0] for k in _all(pg, [(a, (_x(pg, a, 'Variant', node), 'v')) for a in args[1:]])]
    code = "vb_v_null()"
    for k in range(len(args) - 1, 0, -1):
        code = f"({i} == {k} ? {vals[k - 1]} : {code})"
    return _E()(code, 'Variant')


@func('switch')
def f_switch(pg, node, args):
    if len(args) < 2 or len(args) % 2:
        pg.err("Switch needs pairs of expression and value", node)
    items = []
    for k in range(0, len(args), 2):
        items.append((args[k], (pg.cond(args[k]), 'l')))
        items.append((args[k + 1], (_x(pg, args[k + 1], 'Variant', node), 'v')))
    vals = [k[0] for k in _all(pg, items)]
    code = "vb_v_null()"
    for k in range(len(args) - 2, -1, -2):
        code = f"({vals[k]} ? {vals[k + 1]} : {code})"
    return _E()(code, 'Variant')


@func('strconv')
def f_strconv(pg, node, args):
    a = _args(pg, node, args, 2, 3, 'StrConv')
    if len(a) > 2 and a[2] is not None:          # StrConv(s, conv, LocaleID)
        return _E()(f"vb_strconv_l({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Long', node)}, "
                    f"{_x(pg, a[2], 'Long', node)}, 1)", 'String')
    return _E()(f"vb_strconv({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Long', node)})", 'String')


@func('dateadd')
def f_dateadd(pg, node, args):
    a = _args(pg, node, args, 3, 3, 'DateAdd')
    return _E()(f"vb_dateadd({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Double', node)}, "
                f"{_x(pg, a[2], 'Date', node)})", 'Date')


@func('datediff')
def f_datediff(pg, node, args):
    a = _args(pg, node, args, 3, 5, 'DateDiff')
    return _E()(f"vb_datediff({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Date', node)}, "
                f"{_x(pg, a[2], 'Date', node)}, {_opt(pg, a[3], 'Long', node, '1')})", 'Long')


@func('datepart')
def f_datepart(pg, node, args):
    a = _args(pg, node, args, 2, 4, 'DatePart')
    return _E()(f"(vb_i32)vb_datepart({_x(pg, a[0], 'String', node)}, {_x(pg, a[1], 'Date', node)}, "
                f"{_opt(pg, a[2], 'Long', node, '1')}, {_opt(pg, a[3], 'Long', node, '1')})", 'Integer')


@func('datevalue')
def f_datevalue(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'DateValue')[0]
    return _E()(f"vb_datevalue({_x(pg, a, 'String', node)})", 'Date')


@func('timevalue')
def f_timevalue(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'TimeValue')[0]
    return _E()(f"vb_timevalue({_x(pg, a, 'String', node)})", 'Date')


@func('format', 'format$')
def f_format(pg, node, args):
    a = _args(pg, node, args, 1, 4, 'Format')
    f = _x(pg, a[1], 'String', node) if a[1] is not None else 'vb_stmp(0)'
    return _E()(f"vb_format({_x(pg, a[0], 'Variant', node)}, {f}, {1 if a[1] is not None else 0}, "
                f"{_opt(pg, a[2], 'Long', node, '1')}, {_opt(pg, a[3], 'Long', node, '1')})", 'String')


def _fmtnum(kind, name):
    def h(pg, node, args):
        a = _args(pg, node, args, 1, 5, name)
        return _E()(f"vb_formatnumber({_x(pg, a[0], 'Variant', node)}, {_opt(pg, a[1], 'Long', node, '-1')}, "
                    f"{_opt(pg, a[2], 'Long', node, '-2')}, {_opt(pg, a[3], 'Long', node, '-2')}, "
                    f"{_opt(pg, a[4], 'Long', node, '-2')}, {kind})", 'String')
    FUNCS[name.lower()] = h


_fmtnum(0, 'FormatNumber')
_fmtnum(1, 'FormatCurrency')
_fmtnum(2, 'FormatPercent')


@func('formatdatetime')
def f_formatdatetime(pg, node, args):
    a = _args(pg, node, args, 1, 2, 'FormatDateTime')
    return _E()(f"vb_formatdatetime({_x(pg, a[0], 'Date', node)}, {_opt(pg, a[1], 'Long', node, '0')})", 'String')


@func('monthname')
def f_monthname(pg, node, args):
    a = _args(pg, node, args, 1, 2, 'MonthName')
    return _E()(f"vb_monthname({_x(pg, a[0], 'Long', node)}, {_opt(pg, a[1], 'Boolean', node, '0')})", 'String')


@func('weekdayname')
def f_weekdayname(pg, node, args):
    a = _args(pg, node, args, 1, 3, 'WeekdayName')
    return _E()(f"vb_weekdayname({_x(pg, a[0], 'Long', node)}, {_opt(pg, a[1], 'Boolean', node, '0')}, "
                f"{_opt(pg, a[2], 'Long', node, '0')})", 'String')


# ---------------------------------------------------------------- forms: Load, Unload, Forms
@stmt('load')
def s_load(pg, st, args):
    a = _args(pg, st, args, 1, 1, 'Load')[0]
    if a.kind == 'paren':
        a = a.e
    sy = pg.lookup(a.base.name, a.base) if a.kind == 'apply' and a.base.kind == 'name' else None
    if sy is not None and sy.kind == 'var' and sy.array:   # Load frms(i): a form in an array of forms
        pg.emit(f"vbf_load({_x(pg, a, 'Object', st)});")
        return
    if a.kind == 'apply' and len(a.args) == 1 and a.args[0] is not None:
        # Load Text1(3): a new element of a control array
        arr = pg.expr(a.base)
        pg.emit(f"vbf_load_elem({pg.conv(arr, 'Object', st)}, {_x(pg, a.args[0], 'Long', st)});")
        return
    pg.emit(f"vbf_load({_x(pg, a, 'Object', st)});")


@stmt('unload')
def s_unload(pg, st, args):
    a = _args(pg, st, args, 1, 1, 'Unload')[0]
    if a.kind == 'paren':
        a = a.e
    if a.kind == 'name':                        # Unload DataReport1: a designer's default instance goes
        key = a.name.lower()
        ci = pg.c.classes.get(key)
        if ci is not None and getattr(ci, 'predeclared', False) and pg.lookup(a.name, st) is not None and \
                pg.lookup(a.name, st).kind == 'formdefault':
            from .cgen import cident
            k = cident(key)
            if 'lb_unload' in ci.members:
                pg.emit(f'if (G_FORM_{k}) vb_late_o(G_FORM_{k}, "lb_unload", VBM_METHOD, 0, 0);')
            pg.emit(f'vb_ofree(&G_FORM_{k});')
            return
    pg.emit(f"vbf_unload({_x(pg, a, 'Object', st)});")


@func('lb_report_new')
def f_lb_report_new(pg, node, args):
    pg.c.uses_forms = True
    a = _args(pg, node, args, 2, 2, 'lb_report_new')
    return _E()(f"vbf_report_new({_x(pg, a[0], 'Object', node)}, {_x(pg, a[1], 'String', node)})", 'Object')


@func('forms')
def f_forms(pg, node, args):
    if args:
        return pg.default_member(_E()("vbf_forms()", 'Object'), args, node)
    return _E()("vbf_forms()", 'Object')


@func('screen')
def f_screen(pg, node, args):
    pg.c.uses_forms = True
    if args:
        return pg.default_member(_E()("vbf_screen()", 'Object'), args, node)
    return _E()("vbf_screen()", 'Object')


@func('printer')
def f_printer(pg, node, args):
    pg.c.uses_forms = True
    return _E()("vbf_printer()", 'Object')


@func('printers')
def f_printers(pg, node, args):
    pg.c.uses_forms = True
    if args:
        return pg.default_member(_E()("vbf_printers()", 'Object'), args, node)
    return _E()("vbf_printers()", 'Object')


@func('clipboard')
def f_clipboard(pg, node, args):
    pg.c.uses_forms = True
    return _E()("vbf_clipboard()", 'Object')


@func('callbyname')
def f_callbyname(pg, node, args):
    a = _args(pg, node, args, 3, 64, 'CallByName')
    while a and a[-1] is None:
        a.pop()
    obj = _x(pg, a[0], 'Object', node)
    name = _x(pg, a[1], 'String', node)
    ct = _x(pg, a[2], 'Long', node)
    rest = [('vb_v_missing()' if x is None else _x(pg, x, 'Variant', node)) for x in a[3:]]
    arr = f"(vbvar[]){{ {', '.join(rest)} }}" if rest else '0'
    return _E()(f"vb_callbyname({obj}, {name}, {ct}, {len(rest)}, {arr})", 'Variant')


@func('licenses')
def f_licenses(pg, node, args):
    pg.c.uses_forms = True
    if args:
        return pg.default_member(_E()("vbf_licenses()", 'Object'), args, node)
    return _E()("vbf_licenses()", 'Object')


@func('getobject')
def f_getobject(pg, node, args):
    """GetObject(, class): a new object of the class (no program runs one to be found here);
    GetObject(file): a file opened by its program cannot be had (error 432)"""
    a = _args(pg, node, args, 0, 2, 'GetObject')
    if a[1] is None:
        f = _x(pg, a[0], 'String', node) if a[0] is not None else 'vb_stmp(0)'
        return _E()(f"((void)({f}), vb_raise(432), (vbobj)0)", 'Object')
    if a[0] is not None:
        pg.pre.append(f"(void)({_x(pg, a[0], 'String', node)});")
    return _E()(f"vb_createobject({_x(pg, a[1], 'String', node)}, vb_v_missing())", 'Object')


@func('createobject')
def f_createobject(pg, node, args):
    a = _args(pg, node, args, 1, 2, 'CreateObject')
    srv = 'vb_v_missing()' if len(a) < 2 or a[1] is None else _x(pg, a[1], 'Variant', node)
    return _E()(f"vb_createobject({_x(pg, a[0], 'String', node)}, {srv})", 'Object')


@func('loadpicture')
def f_loadpicture(pg, node, args):
    pg.c.uses_forms = True
    a = _args(pg, node, args, 0, 5, 'LoadPicture')      # (FileName, Size, ColorDepth, X, Y): the file only
    if a[0] is None:
        return _E()("vb_loadpicture(vb_stmp(0), 0)", 'Object')
    return _E()(f"vb_loadpicture({_x(pg, a[0], 'String', node)}, 1)", 'Object')


@func('loadresstring', 'loadresstring$')
def f_loadresstring(pg, node, args):
    a = _args(pg, node, args, 1, 1, 'LoadResString')
    return _E()(f"vb_loadresstring({_x(pg, a[0], 'Long', node)})", 'String')


@func('loadresdata')
def f_loadresdata(pg, node, args):
    a = _args(pg, node, args, 2, 2, 'LoadResData')
    return _E()(f"vb_loadresdata({_x(pg, a[0], 'Variant', node)}, {_x(pg, a[1], 'Variant', node)})", 'Variant')


@func('loadrespicture')
def f_loadrespicture(pg, node, args):
    pg.c.uses_forms = True
    a = _args(pg, node, args, 2, 2, 'LoadResPicture')
    return _E()(f"vb_loadrespicture({_x(pg, a[0], 'Variant', node)}, {_x(pg, a[1], 'Long', node)})", 'Object')


@stmt('savepicture')
def s_savepicture(pg, st, args):
    a = _args(pg, st, args, 2, 2, 'SavePicture')
    pg.emit(f"vb_savepicture({_x(pg, a[0], 'Object', st)}, {_x(pg, a[1], 'String', st)});")
