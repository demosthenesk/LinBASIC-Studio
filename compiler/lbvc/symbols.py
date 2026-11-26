# LinBASIC Studio - lbvc: types, symbols and the VBA constants
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
import json, os, re

NUMERIC = ('Byte', 'Integer', 'Long', 'Single', 'Currency', 'Double')
RANK = {'Byte': 1, 'Boolean': 2, 'Integer': 2, 'Long': 3, 'Single': 4, 'Currency': 5, 'Double': 6, 'Date': 7}
CT = {'Integer': 'vb_i16', 'Long': 'vb_i32', 'Single': 'float', 'Double': 'double', 'Currency': 'vb_cy',
      'Boolean': 'vb_bool', 'Byte': 'vb_u8', 'String': 'vbs', 'Variant': 'vbvar', 'Date': 'vb_date'}
ZERO = {'vbs': '{0, 0, 0}', 'vbvar': '{0}'}
ARRAY_KIND = {'String': 'VBA_STR', 'Variant': 'VBA_VAR'}
SUPPORTED_TYPES = set(CT)


def is_num(t):
    return t in RANK


def is_obj(t):
    """an object type: Object (late bound) or 'O:<class key>'"""
    return bool(t) and (t == 'Object' or t.startswith('O:'))


def is_udt(t):
    """a user-defined type: 'U:<type key>'"""
    return bool(t) and t.startswith('U:')


def is_arr(t):
    """an array value: 'A:<element type>' (functions returning arrays, Split ...)"""
    return bool(t) and t.startswith('A:')


VT_OF = {'Integer': 'VT_I2', 'Long': 'VT_I4', 'Single': 'VT_R4', 'Double': 'VT_R8', 'Currency': 'VT_CY',
         'Date': 'VT_DATE', 'String': 'VT_BSTR', 'Boolean': 'VT_BOOL', 'Variant': 'VT_VARIANT', 'Byte': 'VT_UI1'}


def vt_of(t):
    """VarType constant of an element type"""
    if t == 'Object' or (t or '').startswith('O:'):
        return 'VT_DISPATCH'
    return VT_OF.get(t, 'VT_VARIANT')


def ctype(t):
    if is_arr(t):
        return 'vbarr'
    if is_udt(t):
        return 'U_' + t[2:]
    if is_obj(t):
        return 'vbobj'
    return CT[t]


def zero(t):
    if is_udt(t) or is_arr(t):
        return '{0}'
    if is_obj(t):
        return '0'
    return ZERO.get(CT[t], '0')


def akind(t):
    """element kind of an array of t (runtime VBA_*)"""
    if is_udt(t):
        return 'VBA_UDT'
    if is_obj(t):
        return 'VBA_OBJ'
    return ARRAY_KIND.get(t, 'VBA_PLAIN')


def known_type(t):
    return t in CT or is_obj(t) or is_udt(t) or is_arr(t)


def cident(name):
    """C identifier for a (case-insensitive) VB name"""
    out = []
    for ch in name.lower():
        if ch.isalnum() and ch.isascii() or ch == '_':
            out.append(ch)
        else:
            out.append('_u%04x' % ord(ch))
    return ''.join(out)


class Sym:
    def __init__(self, kind, name, **kw):
        self.kind = kind                 # 'var' 'const' 'proc' 'enum'
        self.name = name
        self.__dict__.update(kw)

    def __repr__(self):
        return f"Sym({self.kind} {self.name})"


_VBA_CONSTS = None


def vba_constants():
    """name (lower) -> (value, type) for every constant of the VBA, VBRUN and MSComDlg libraries"""
    global _VBA_CONSTS
    if _VBA_CONSTS is not None:
        return _VBA_CONSTS
    here = os.path.dirname(os.path.abspath(__file__))
    base = os.path.join(here, '..', '..', 'data', 'vb6', 'typelibs')
    out = {}
    for lib in ('VBA', 'VBRUN', 'MSComDlg', 'Scripting', 'MSComctlLib', 'MSComCtl2', 'MSFlexGridLib', 'TabDlg',
                'RichTextLib', 'MSChart20Lib', 'MCI', 'MSCommLib', 'PicClip', 'DAO', 'ADODB', 'MSBind', 'MSAdodcLib', 'MSDataGridLib', 'MSDataListLib', 'MSDataReportLib', 'MSWinsockLib', 'InetCtlsObjects', 'MSMask', 'MSDBGrid', 'MSDBCtls', 'StdFormat', 'ComCtl3', 'RDO', 'MSRDC', 'MSDataRepeaterLib', 'MSMAPI', 'SHDocVw', 'stdole'):   # MSComDlg: CommonDialog (cdl...); Scripting: ForReading ...
        p = os.path.join(base, lib + '.json')
        if not os.path.exists(p):
            continue
        for t in json.load(open(p, encoding='utf-8'))['types']:
            if t['kind'] not in ('enum', 'module'):
                continue
            for m in t.get('members', []):
                if m['kind'] != 'const' or m['name'].startswith('_'):
                    continue
                v, ty = m.get('value'), m.get('type', 'Long')
                if lib == 'stdole' and m['name'].lower() in out:   # stdole (Checked, Default ...): after the others
                    continue
                if isinstance(v, str):
                    out[m['name'].lower()] = (v, 'String')
                elif isinstance(v, bool):
                    out[m['name'].lower()] = (-1 if v else 0, 'Boolean')
                elif isinstance(v, float):
                    out[m['name'].lower()] = (v, 'Double')
                elif v is not None:
                    out[m['name'].lower()] = (v, 'Integer' if ty == 'Integer' else 'Long')
    _VBA_CONSTS = out
    return out


def lib_const_names(lib):
    """the lower names of the constants of one type library (Scripting: Normal, Directory ...)"""
    here = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(here, '..', '..', 'data', 'vb6', 'typelibs', lib + '.json')
    out = set()
    if os.path.exists(p):
        for t in json.load(open(p, encoding='utf-8'))['types']:
            if t['kind'] in ('enum', 'module'):
                out |= {m['name'].lower() for m in t.get('members', []) if m['kind'] == 'const'}
    return out


_LIB_ENUMS = None


def lib_enums():
    """the enum types of the libraries (AlignmentConstants, VbMsgBoxResult ...): their lower names, also
    as Library.Name; As one of them is a Long"""
    global _LIB_ENUMS
    if _LIB_ENUMS is not None:
        return _LIB_ENUMS
    here = os.path.dirname(os.path.abspath(__file__))
    base = os.path.join(here, '..', '..', 'data', 'vb6', 'typelibs')
    out = set()
    for lib in ('VBA', 'VBRUN', 'VB', 'stdole', 'MSComDlg', 'Scripting', 'MSComctlLib', 'MSComCtl2', 'MSFlexGridLib',
                'TabDlg', 'RichTextLib', 'MSChart20Lib', 'MCI', 'MSCommLib', 'PicClip', 'DAO', 'ADODB', 'MSBind', 'MSAdodcLib', 'MSDataGridLib', 'MSDataListLib', 'MSDataReportLib', 'MSWinsockLib', 'InetCtlsObjects', 'MSMask', 'MSDBGrid', 'MSDBCtls', 'StdFormat', 'ComCtl3', 'RDO', 'MSRDC', 'MSDataRepeaterLib', 'MSMAPI', 'SHDocVw'):
        p = os.path.join(base, lib + '.json')
        if not os.path.exists(p):
            continue
        for t in json.load(open(p, encoding='utf-8'))['types']:
            if t['kind'] == 'enum' and not t['name'].startswith('_'):
                out.add(t['name'].lower())
                out.add(f"{lib.lower()}.{t['name'].lower()}")
    _LIB_ENUMS = out
    return out


_DATE_RE = re.compile(r'^\s*(?:(\d{1,4})[/-](\d{1,2})[/-](\d{1,4}))?\s*(?:(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?\s*$', re.I)


def date_literal(text):
    """#m/d/yyyy hh:mm:ss AM# -> days since 30 Dec 1899 (VB date literals are always month/day/year)"""
    import datetime
    m = _DATE_RE.match(text)
    if not m or not (m.group(1) or m.group(4)):
        return None
    v = 0.0
    if m.group(1):
        a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a > 31:
            y, mo, d = a, b, c                       # yyyy-mm-dd
        else:
            mo, d, y = a, b, c
        if y < 30:
            y += 2000
        elif y < 100:
            y += 1900
        v = (datetime.date(y, mo, d) - datetime.date(1899, 12, 30)).days
    if m.group(4):
        h, mi, s = int(m.group(4)), int(m.group(5)), int(m.group(6) or 0)
        ap = (m.group(7) or '').upper()
        if ap == 'PM' and h < 12:
            h += 12
        if ap == 'AM' and h == 12:
            h = 0
        t = (h * 3600 + mi * 60 + s) / 86400.0
        v = v + t if v >= 0 else v - t
    return float(v)


_VBA_VARIANT = None


def vba_variant_funcs():
    """the lower names of the functions of the VBA library that give a Variant (Hour, Year, TimeSerial ...): lbvc
    types some of them (Integer, Date), VB6 gives Len of them the length of the text (Len(Hour(Now)) = 2)"""
    global _VBA_VARIANT
    if _VBA_VARIANT is not None:
        return _VBA_VARIANT
    here = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(here, '..', '..', 'data', 'vb6', 'typelibs', 'VBA.json')
    out = set()
    try:
        for t in json.load(open(p, encoding='utf-8')).get('types', []):
            for m in t.get('members', []):
                if m.get('kind') == 'func' and m.get('returns') == 'Variant':
                    out.add(m['name'].lower())
    except (OSError, ValueError):
        pass
    _VBA_VARIANT = out
    return out


_VB_PROPS = None


def vb_prop_type(cls, prop):
    """the type of a property of an intrinsic control of VB6 (cls: TextBox, Label ...) from the VB library: Text of
    a TextBox is a String; None if not known"""
    global _VB_PROPS
    if _VB_PROPS is None:
        _VB_PROPS = {}
        here = os.path.dirname(os.path.abspath(__file__))
        p = os.path.join(here, '..', '..', 'data', 'vb6', 'typelibs', 'VB.json')
        try:
            for t in json.load(open(p, encoding='utf-8')).get('types', []):
                if t.get('kind') != 'interface' or not t['name'].startswith('_'):
                    continue
                for m in t.get('members', []):
                    ps = [q for q in m.get('params', []) if 'retval' not in q.get('flags', [])]
                    if m.get('invoke') == 'get' and not ps:
                        _VB_PROPS[(t['name'][1:].lower(), m['name'].lower())] = m.get('returns')
        except (OSError, ValueError):
            pass
    return _VB_PROPS.get((cls.lower(), prop.lower()))


_DISPIDS = None


def prop_dispids(cls):
    """{property name: its DISPID} of a control from the type libraries of data/vb6/typelibs (cls: "TextBox" or
    "MSComctlLib.Slider"): what an UpDown keeps in BuddyProperty (the corpus of PSC: Text of a TextBox 65547,
    Caption of a Label 65537; 0 the default property)"""
    global _DISPIDS
    if _DISPIDS is None:
        _DISPIDS = {}
        here = os.path.dirname(os.path.abspath(__file__))
        base = os.path.join(here, '..', '..', 'data', 'vb6', 'typelibs')
        try:
            names = sorted(os.listdir(base))
        except OSError:
            names = []
        for f in names:
            if not f.endswith('.json'):
                continue
            try:
                d = json.load(open(os.path.join(base, f), encoding='utf-8'))
            except (OSError, ValueError):
                continue
            if not isinstance(d, dict):
                continue
            lib = d.get('library', f[:-5])
            for t in d.get('types', []):
                if t.get('kind') not in ('interface', 'dispinterface') or not t['name'].startswith('_'):
                    continue
                m = {}
                for x in t.get('members', []):
                    if x.get('invoke') in ('get', 'let', 'prop') and x['name'] != '_Default' and x.get('memid') is not None:
                        m.setdefault(x['name'], int(x['memid']))
                key = t['name'][1:].lower()
                _DISPIDS.setdefault(key, m)
                _DISPIDS.setdefault(f"{lib}.{t['name'][1:]}".lower(), m)
    c = (cls or '').lower()
    return _DISPIDS.get(c) or _DISPIDS.get(c.split('.')[-1].replace('vb.', '')) or {}


_PARAMS = None


def member_params():
    """lower member name -> [(interface lower, [parameter names in order])] from every type library in
    data/vb6/typelibs: the order of named arguments (Item:=..., Key:=...) in late-bound calls"""
    global _PARAMS
    if _PARAMS is not None:
        return _PARAMS
    here = os.path.dirname(os.path.abspath(__file__))
    base = os.path.join(here, '..', '..', 'data', 'vb6', 'typelibs')
    out = {}
    for fn in sorted(os.listdir(base)) if os.path.isdir(base) else []:
        if not fn.endswith('.json') or fn == 'index.json':
            continue
        try:
            types = json.load(open(os.path.join(base, fn), encoding='utf-8')).get('types', [])
        except (OSError, ValueError):
            continue
        for t in types:
            if t.get('kind') not in ('dispatch', 'interface'):
                continue
            for m in t.get('members', []):
                ps = [p['name'] for p in m.get('params', []) if 'retval' not in p.get('flags', [])]
                if m.get('invoke') in ('let', 'set') and ps:
                    ps = ps[:-1]                 # the value of a Property Let is not a named argument
                lst = out.setdefault(m['name'].lower(), [])
                entry = (t['name'].lower(), ps)
                if entry not in lst:
                    lst.append(entry)
    _PARAMS = out
    return out
