"""oleblob.py - the OleObjectBlob of a DBGrid (DBGRID32.OCX) in a .frx (phase 7.6).

VB6 keeps the design-time state of a DBGrid as one record of the .frx, not as lines of the .frm:

  [u32 'LB\\r\\0'][u32 n: the bytes after this 24-byte head][u32 Left][u32 Top][u32 Width][u32 Height] (twips)
  an OLE property set: [u16 0xFFFE][u16 0][u32 os][CLSID][u32 sections = 1][FMTID][u32 offset]
  its section right after it: [u32 size][u32 count] count x ([u32 id][u32 offset from the section]),
  the values in that order, the dictionary (id 0) last: [u32 n] n x ([u32 id][u32 len][name\\0]).

The offset of the section counts from the start of the .frx (VB6 writes the whole .frx as one stream), so
a blob written at another place of a .frx has other offsets: relocate() moves them (FrxWriter does it).

Values: VT_I2 2, VT_I4 3, VT_R4 4, VT_R8 5, VT_BOOL 0xB (each 4 bytes), VT_LPSTR 0x1E ([u32 len with the
NUL][bytes], padded to 4), VT_BLOB 0x41 ([u32 size with these 4 bytes][bytes]), 0x46 a font ([u32 size][CLSID StdFont][the
font]), 0x104B a list of property sets ([u32 count] count x ([u32 size][a set])): Columns (CLSID
00028C07), Splits (00028C0C). A Split keeps the widths of its columns in its _ColumnProps blob ("BigRed").

Checked against a .frx that VB6 saved (tests/dev/f_dbgrid: the default layout, two blank columns): parse()
and build() give it again byte for byte. decode() gives the design-time properties as the .frm would have
them: Name -> value, Columns.Count, Columns(i).Caption ..., Split0.MarqueeStyle ..., Font.Name ...
"""
import struct

VT_I2, VT_I4, VT_R4, VT_R8, VT_BOOL, VT_LPSTR, VT_BLOB, VT_FONT, VT_SETS = 2, 3, 4, 5, 0xB, 0x1E, 0x41, 0x46, 0x104B
MAGIC = b'LB\r\0'
HEAD = 24
CLSID_GRID = bytes.fromhex('008c0200000000000000000000000046')     # {00028C00-0000-0000-0000-000000000046}
CLSID_COLUMN = bytes.fromhex('078c0200000000000000000000000046')
CLSID_SPLIT = bytes.fromhex('0c8c0200000000000000000000000046')
FMTID = bytes.fromhex('21088ffb64011b1084ed08002b2ec713')
CLSID_FONT = bytes.fromhex('0352e30b918fce119de300aa004bb851')   # StdFont


def _u32(d, o):
    return struct.unpack_from('<I', d, o)[0]


def _i32(d, o):
    return struct.unpack_from('<i', d, o)[0]


def _pad4(b):
    return b + b'\0' * (-len(b) % 4)


class PSet:
    """a property set: its values by id in the order of the section, the names in the order of the
    dictionary"""

    def __init__(self, clsid, os_=0x00020A04, fmtid=FMTID):
        self.clsid, self.os, self.fmtid = clsid, os_, fmtid
        self.order = []                         # the ids as the section lists them (0: the dictionary)
        self.values = {}                        # id -> (vt, value)
        self.names = []                         # [(id, name)] as the dictionary lists them

    def pid(self, name):
        for k, n in self.names:
            if n.lower() == name.lower():
                return k
        return None

    def get(self, name, default=None):
        k = self.pid(name)
        return self.values[k][1] if k is not None and k in self.values else default

    def set(self, name, value):
        """a value of a property the set has (its type stays)"""
        k = self.pid(name)
        if k is None or k not in self.values:
            raise KeyError(name)
        vt, _ = self.values[k]
        self.values[k] = (vt, value)

    def items(self):
        """(name, vt, value) in the order of the section"""
        nm = dict(self.names)
        for k in self.order:
            if k:
                vt, v = self.values[k]
                yield nm.get(k, f'#{k}'), vt, v


# ---------------------------------------------------------------- reading
def font(b):
    """a StdFont as it is saved (after its CLSID): [u8 version][u16 charset][u8 flags][u16 weight]
    [u32 size * 10000][u8 len][name]"""
    if len(b) < 11:
        return {}
    charset = struct.unpack_from('<H', b, 1)[0]
    flags = b[3]
    weight = struct.unpack_from('<H', b, 4)[0]
    size = _u32(b, 6) / 10000.0
    n = b[10]
    name = b[11:11 + n].decode('cp1252', 'replace')
    return {'Name': name, 'Size': size, 'Charset': charset, 'Weight': weight,
            'Underline': -1 if flags & 4 else 0, 'Italic': -1 if flags & 2 else 0,
            'Strikethrough': -1 if flags & 8 else 0}


def _read_set(d, base):
    if d[base:base + 2] != b'\xfe\xff':
        raise ValueError('no property set')
    ps = PSet(bytes(d[base + 8:base + 24]), _u32(d, base + 4), bytes(d[base + 28:base + 44]))
    sec = base + 48                             # the offset at base + 44 counts from the start of the .frx
    count = _u32(d, sec + 4)
    pairs = [(_i32(d, sec + 8 + 8 * i), _u32(d, sec + 12 + 8 * i)) for i in range(count)]
    for pid, off in pairs:
        ps.order.append(pid)
        o = sec + off
        if pid == 0:
            n = _u32(d, o)
            o += 4
            for _ in range(n):
                k, ln = _i32(d, o), _u32(d, o + 4)
                ps.names.append((k, d[o + 8:o + 8 + ln].rstrip(b'\0').decode('cp1252', 'replace')))
                o += 8 + ln
            continue
        vt = _u32(d, o)
        if vt == VT_I2:
            v = struct.unpack_from('<h', d, o + 4)[0]
        elif vt == VT_I4:
            v = _i32(d, o + 4)
        elif vt == VT_R4:
            v = struct.unpack_from('<f', d, o + 4)[0]
        elif vt == VT_R8:
            v = struct.unpack_from('<d', d, o + 4)[0]
        elif vt == VT_BOOL:
            v = -1 if struct.unpack_from('<h', d, o + 4)[0] else 0
        elif vt == VT_LPSTR:
            ln = _u32(d, o + 4)
            v = d[o + 8:o + 8 + ln].split(b'\0')[0].decode('cp1252', 'replace')
        elif vt == VT_BLOB:                     # (its length counts its own 4 bytes)
            ln = _u32(d, o + 4)
            v = bytes(d[o + 8:o + 4 + ln])
        elif vt == VT_FONT:
            ln = _u32(d, o + 4)
            v = bytes(d[o + 8:o + 8 + ln])
        elif vt == VT_SETS:
            n = _u32(d, o + 4)
            p = o + 8
            v = []
            for _ in range(n):
                size = _u32(d, p)
                v.append(_read_set(d, p + 4))
                nx = p + 4 + size
                if d[nx + 4:nx + 6] != b'\xfe\xff' and d[nx + 4 + (-nx % 4):nx + 6 + (-nx % 4)] == b'\xfe\xff':
                    nx += -nx % 4               # (a set padded to 4 bytes)
                p = nx
        else:
            raise ValueError(f'type {vt:#x}')
        ps.values[pid] = (vt, v)
    return ps


def parse(data):
    """the OleObjectBlob record -> (head: the 4 numbers after the length, PSet of the grid)"""
    d = bytes(data)
    if d[:4] != MAGIC:
        raise ValueError('not an OleObjectBlob')
    head = struct.unpack_from('<4i', d, 8)
    return list(head), _read_set(d, HEAD)


# ---------------------------------------------------------------- writing
def _value(vt, v, at):
    """the bytes of a value; at: where it goes in the stream (the offsets of nested sets)"""
    if vt == VT_I2:
        return struct.pack('<Ihh', vt, v, 0)
    if vt == VT_I4:
        return struct.pack('<Ii', vt, v)
    if vt == VT_R4:
        return struct.pack('<If', vt, v)
    if vt == VT_R8:
        return struct.pack('<Id', vt, v)
    if vt == VT_BOOL:
        return struct.pack('<Ihh', vt, -1 if v else 0, 0)
    if vt == VT_LPSTR:
        b = str(v).encode('cp1252', 'replace') + b'\0'
        return struct.pack('<II', vt, len(b)) + _pad4(b)
    if vt == VT_BLOB:
        return struct.pack('<II', vt, len(v) + 4) + _pad4(bytes(v))
    if vt == VT_FONT:
        return struct.pack('<II', vt, len(v)) + _pad4(bytes(v))
    if vt == VT_SETS:
        out = bytearray(struct.pack('<II', vt, len(v)))
        for s in v:                             # (no padding between the sets: VB6, a grid of 14 columns)
            b = _write_set(s, at + len(out) + 4)
            out += struct.pack('<I', len(b)) + b
        return _pad4(bytes(out))
    raise ValueError(f'type {vt:#x}')


def _write_set(ps, at):
    """a property set at the stream position at"""
    head = struct.pack('<HHI', 0xFFFE, 0, ps.os) + ps.clsid + struct.pack('<I', 1) + ps.fmtid + struct.pack('<I', at + 48)
    sec_at = at + 48
    n = len(ps.order)
    body = bytearray()
    offs = []
    for pid in ps.order:
        off = 8 + 8 * n + len(body)
        offs.append(off)
        if pid == 0:
            dct = bytearray(struct.pack('<I', len(ps.names)))
            for k, name in ps.names:
                b = name.encode('cp1252', 'replace') + b'\0'
                dct += struct.pack('<iI', k, len(b)) + b
            body += dct
        else:
            vt, v = ps.values[pid]
            body += _value(vt, v, sec_at + off)
    sec = struct.pack('<II', 8 + 8 * n + len(body), n) + b''.join(struct.pack('<iI', k, o) for k, o in zip(ps.order, offs))
    return head + sec + bytes(body)


def build(head, ps, at=0):
    """the OleObjectBlob record of a grid, for the .frx offset at"""
    body = _write_set(ps, at + HEAD)
    return MAGIC + struct.pack('<I4i', len(body), *head) + body


def relocate(data, at):
    """the blob again for another place of the .frx (its offsets count from the start of the .frx)"""
    try:
        head, ps = parse(data)
    except (ValueError, IndexError, struct.error):
        return bytes(data)
    return build(head, ps, at)


# ---------------------------------------------------------------- the properties
def column_widths(blob):
    """the widths (HIMETRIC) of the columns in a Split's _ColumnProps: "BigRed" [u8][u8][u32 n], then for
    each column [u32 number, from 1][u32 k] k x ([u32 4][u32 id][u32 value][u32]); id 0x519 is the width"""
    out = {}
    if not blob.startswith(b'BigRed') or len(blob) < 12:
        return out
    n = _u32(blob, 8)
    o = 12
    for _ in range(n):
        if o + 8 > len(blob):
            break
        num, k = _u32(blob, o), _u32(blob, o + 4)
        o += 8
        for _ in range(k):
            if o + 16 > len(blob):
                break
            pid, val = _u32(blob, o + 4), _i32(blob, o + 8)
            if pid == 0x519:
                out[num - 1] = val
            o += 16
    return out


ALIGN_READ = {0: 0, 1: 2, 2: 1}                 # the record 0x425 -> Alignment (dbgLeft 0, dbgRight 1, dbgCenter 2)
ALIGN_WRITE = {0: 0, 1: 2, 2: 1, 3: 0x10}


def column_aligns(blob):
    """the Alignment of the columns from their record 0x425 in the _ColumnProps: 16 dbgGeneral, 0 dbgLeft,
    2 dbgRight, 1 dbgCenter (VB6 oracle f_dbgw: 0 -> 0, 1 -> 2, 2 -> 1, 3 and 4 -> 3); 0x2000 another flag"""
    out = {}
    if not blob.startswith(b'BigRed') or len(blob) < 12:
        return out
    n = _u32(blob, 8)
    o = 12
    for _ in range(n):
        if o + 8 > len(blob):
            break
        num, k = _u32(blob, o), _u32(blob, o + 4)
        o += 8
        for _ in range(k):
            if o + 16 > len(blob):
                break
            pid, val = _u32(blob, o + 4), _i32(blob, o + 8)
            if pid == 0x425:
                out[num - 1] = 3 if val & 0x10 else ALIGN_READ.get(val & 0xF, 3)
            o += 16
    return out


def _plain(vt, v):
    if vt == VT_FONT:
        return font(v[16:]) if len(v) > 16 else {}
    return v


def decode(data):
    """the design-time properties of the grid: [(name, value)], value an int, a float or a str"""
    try:
        head, g = parse(data)
    except (ValueError, IndexError, struct.error):
        return []
    out = []
    cols, splits = [], []
    for k, vt, v in g.items():
        if k == 'Columns' and vt == VT_SETS:
            cols = v
            continue
        if k == 'Splits' and vt == VT_SETS:
            splits = v
            continue
        if k.startswith('_') and k not in ('_ExtentX', '_ExtentY', '_Version', '_LayoutType'):
            continue
        v = _plain(vt, v)
        if isinstance(v, dict):                 # Font, HeadFont
            for fk, fv in v.items():
                out.append((f'{k}.{fk}', fv))
        elif not isinstance(v, (bytes, list)):
            out.append((k, v))
    out.append(('Columns.Count', len(cols)))
    for i, c in enumerate(cols):
        for k in ('Caption', 'DataField', 'NumberFormat', 'DefaultValue'):
            x = c.get(k)
            if x is not None and not isinstance(x, (bytes, list)):
                out.append((f'Columns({i}).{k}', x))
    out.append(('Splits.Count', len(splits)))
    for i, s in enumerate(splits):
        for k, vt, v in s.items():
            if k == '_ColumnProps' and isinstance(v, bytes):
                for ci, w in sorted(column_widths(v).items()):
                    if ci < len(cols):
                        out.append((f'Columns({ci})._Width', w))
                for ci, a in sorted(column_aligns(v).items()):
                    if ci < len(cols):
                        out.append((f'Columns({ci}).Alignment', a))
            elif not k.startswith('_') and not isinstance(v, (bytes, list)):
                out.append((f'Split{i}.{k}', v))
    return out


# the properties of the grid the Properties window edits: those of the grid, and of its Split (one)
GRID_PROPS = ('AllowAddNew', 'AllowArrows', 'AllowDelete', 'AllowUpdate', 'Appearance', 'BackColor', 'BorderStyle',
              'Caption', 'ColumnHeaders', 'DataMode', 'DefColWidth', 'Enabled', 'ForeColor', 'HeadLines',
              'MarqueeUnique', 'RowDividerStyle', 'TabAcrossSplits', 'TabAction', 'WrapCellPointer')
SPLIT_PROPS = ('AllowRowSizing', 'MarqueeStyle', 'RecordSelectors', 'ScrollBars')


def font_bytes(f):
    """a StdFont as a value (0x46): its CLSID, then the font"""
    flags = (2 if f.get('Italic') else 0) | (4 if f.get('Underline') else 0) | (8 if f.get('Strikethrough') else 0)
    name = str(f.get('Name', 'MS Sans Serif')).encode('cp1252', 'replace')[:255]
    return CLSID_FONT + struct.pack('<BHBHIB', 1, int(f.get('Charset', 0)), flags, int(f.get('Weight', 400)),
                                    int(round(float(f.get('Size', 8.25)) * 10000)), len(name)) + name


def get_prop(data, name):
    """a property of the grid from its blob (None: it has none)"""
    try:
        head, g = parse(data)
    except (ValueError, IndexError, struct.error):
        return None
    if name in SPLIT_PROPS:
        sp = g.get('Splits') or []
        return sp[0].get(name) if sp else None
    v = g.get(name)
    if name in ('Font', 'HeadFont') and isinstance(v, bytes):
        return font(v[16:])
    return v


def set_prop(data, name, value):
    """the blob with a property of the grid changed (Font, HeadFont: a dict as font() gives)"""
    head, g = parse(data)
    if name in SPLIT_PROPS:
        sp = g.get('Splits') or []
        if sp:
            sp[0].set(name, int(value))
    elif name in ('Font', 'HeadFont'):
        g.set(name, font_bytes(value))
    elif name in ('Width', 'Height'):
        head[2 if name == 'Width' else 3] = int(value)
        g.set('_ExtentX' if name == 'Width' else '_ExtentY', int(round(int(value) * 2540 / 1440)))
    elif name in ('Left', 'Top'):
        head[0 if name == 'Left' else 1] = int(value)
    else:
        vt = g.values[g.pid(name)][0]
        g.set(name, str(value) if vt == VT_LPSTR else float(value) if vt in (VT_R4, VT_R8) else int(value))
    return build(head, g)


def _column(i, caption='', field='', default='', numfmt=''):
    c = PSet(CLSID_COLUMN)
    c.order = [5, 8, 9, 0x11, 0]
    c.values = {5: (VT_LPSTR, caption), 8: (VT_LPSTR, field), 9: (VT_LPSTR, default), 0x11: (VT_LPSTR, numfmt)}
    c.names = [(0, f'Column{i}'), (5, 'Caption'), (8, 'DataField'), (9, 'DefaultValue'), (0x11, 'NumberFormat')]
    return c


# the Layout page of a column (VB6 oracle DEV-11 s_dbgbag, s_dbgbag2: the stream of the grid in a PropertyBag after
# each change): DividerStyle -> (0x501, 0x4FF the colour of the line, the pixels the width gains), Locked: the flag
# 0x2000 of 0x425 (with the Alignment), AllowSizing 0x507, Visible 0x42B, WrapText 0x4D4, Button 0x4C8
DIVIDERS = {0: (0, None, -1), 1: (1, 0x000000, 0), 2: (1, 0x808080, 0), 3: (2, None, 1), 4: (3, None, 1), 5: (4, None, 0)}
LAYOUT_FLAGS = (("AllowSizing", 0x507, True), ("Visible", 0x42B, True), ("WrapText", 0x4D4, False),
                ("Button", 0x4C8, False))


def column_records(blob):
    """the records of each column of a _ColumnProps: {column (from 0): {id: value}}"""
    out = {}
    if not blob.startswith(b'BigRed') or len(blob) < 12:
        return out
    n = _u32(blob, 8)
    o = 12
    for _ in range(n):
        if o + 8 > len(blob):
            break
        num, k = _u32(blob, o), _u32(blob, o + 4)
        o += 8
        r = out.setdefault(num - 1, {})
        for _ in range(k):
            if o + 16 > len(blob):
                break
            r[_u32(blob, o + 4)] = _i32(blob, o + 8)
            o += 16
    return out


def divider_of(rec):
    """the DividerStyle of a column from its records"""
    kind, color = rec.get(0x501, 1), rec.get(0x4FF, 0x808080)
    if kind == 1:
        return 1 if (color & 0xFFFFFF) == 0 else 2
    return {0: 0, 2: 3, 3: 4, 4: 5}.get(kind, 2)


def _align_raw(al, lay):
    """the record 0x425 of an Alignment: 0x10 a column whose Alignment was never set (VB6 oracle DEV-11 s_dbgal: the
    page shows 0 - Left for it); one set on the Layout page: 0 Left, 2 Right, 1 Center, 4 General"""
    if not lay or (int(al) == 3 and lay.get('AlignFlag', 'DividerStyle' not in lay)):
        return ALIGN_WRITE.get(int(al), 0x10)
    return {0: 0, 1: 2, 2: 1, 3: 4}.get(int(al), 4)


def him_of_twips(tw):
    """a width typed in twips as the grid keeps it: whole pixels, in HIMETRIC (s_dbgbag2: 1000 -> 1773)"""
    return int(round(round(float(tw) / 15) * 2540 / 96))


def _column_props(n, him, aligns=None, layout=None):
    """the _ColumnProps of a Split: n columns of the width him (a list: one for each), the other values as VB6
    writes them for a new grid: 0x501 1, 0x4FF the divider color, 0x4EE 1, 0x507 1, 0x425 16, 0x42B 1, 0x4D4 0,
    0x4C8 0, 0x484 0; layout: a dict for each column (Locked, AllowSizing, Visible, WrapText, Button,
    DividerStyle) or None"""
    widths = him if isinstance(him, (list, tuple)) else [him] * n
    out = bytearray(b'BigRed\x01\x01' + struct.pack('<I', n))
    for i in range(n):
        al = aligns[i] if aligns and i < len(aligns) and aligns[i] is not None else 3
        lay = (layout[i] if layout and i < len(layout) else None) or {}
        kind, color, _ = DIVIDERS.get(int(lay.get('DividerStyle', 2)), DIVIDERS[2])
        flags = {pid: (1 if lay.get(k, d) else 0) for k, pid, d in LAYOUT_FLAGS}
        recs = [(0x519, int(widths[i])), (0x501, kind), (0x4FF, 0x808080 if color is None else color), (0x4EE, 1),
                (0x507, flags[0x507]),
                (0x425, _align_raw(al, lay) | (0x2000 if lay.get('Locked') else 0)),
                (0x42B, flags[0x42B]), (0x4D4, flags[0x4D4]), (0x4C8, flags[0x4C8]), (0x484, 0)]
        out += struct.pack('<II', i + 1, len(recs))
        for k, v in recs:
            out += struct.pack('<IIiI', 4, k, v, 0)
    return bytes(out)


COLUMN_KEYS = ('Caption', 'DataField', 'DefaultValue', 'NumberFormat')
DEFAULT_COL_WIDTH = 2381                        # HIMETRIC (1350 twips)


def get_columns(data):
    """the columns of the grid: [{'Caption', 'DataField', 'DefaultValue', 'NumberFormat', 'Width' (HIMETRIC)}]"""
    try:
        head, g = parse(data)
    except (ValueError, IndexError, struct.error):
        return []
    sp = g.get('Splits') or []
    widths = column_widths(sp[0].get('_ColumnProps') or b'') if sp else {}
    aligns = column_aligns(sp[0].get('_ColumnProps') or b'') if sp else {}
    recs = column_records(sp[0].get('_ColumnProps') or b'') if sp else {}
    out = []
    for i, c in enumerate(g.get('Columns') or []):
        d = {k: c.get(k, '') or '' for k in COLUMN_KEYS}
        d['Width'] = widths.get(i, DEFAULT_COL_WIDTH)
        d['Alignment'] = aligns.get(i, 3)
        r = recs.get(i, {})
        d['Locked'] = bool(r.get(0x425, 0) & 0x2000)
        for k, pid, dflt in LAYOUT_FLAGS:
            d[k] = bool(r.get(pid, 1 if dflt else 0))
        d['DividerStyle'] = divider_of(r)
        d['AlignFlag'] = bool(r.get(0x425, 0x10) & 0x10)    # (never set: the Layout page shows 0 - Left)
        out.append(d)
    return out


def set_columns(data, cols, layout=4):
    """the blob with these columns (as get_columns gives them); the widths and the alignments go into the
    _ColumnProps of the Split. layout: _LayoutType, 4 a layout made at design time (the corpus: every grid with
    columns of its own), 0 the default one (two blank columns; at run time the columns are the fields)"""
    head, g = parse(data)
    sets = [_column(i, *(str(c.get(k, '') or '') for k in COLUMN_KEYS)) for i, c in enumerate(cols)]
    g.set('Columns', sets)
    g.set('_LayoutType', int(layout))
    sp = g.get('Splits') or []
    if sp:
        sp[0].set('_ColumnProps', _column_props(len(cols), [int(c.get('Width') or DEFAULT_COL_WIDTH) for c in cols],
                                                [c.get('Alignment', 3) for c in cols], cols))
    return build(head, g)


SPLIT_PAGE = ('Locked', 'AllowFocus', 'AllowSizing', 'AllowRowSizing', 'RecordSelectors', 'ScrollGroup', 'ScrollBars',
              'MarqueeStyle', 'Size', 'SizeMode')


def get_split(data):
    """the properties of Split0 (the Splits page): {name: value}"""
    try:
        head, g = parse(data)
    except (ValueError, IndexError, struct.error):
        return {}
    sp = g.get('Splits') or []
    return {k: sp[0].get(k) for k in SPLIT_PAGE if sp and sp[0].get(k) is not None}


def set_split(data, values):
    """the blob with these properties of Split0 (booleans: -1 / 0, Size: a float when its type is one)"""
    head, g = parse(data)
    sp = g.get('Splits') or []
    if not sp:
        return bytes(data)
    for k, v in values.items():
        pid = sp[0].pid(k)
        if pid is None or pid not in sp[0].values:
            continue
        vt = sp[0].values[pid][0]
        sp[0].set(k, (-1 if v else 0) if vt == VT_BOOL else float(v) if vt in (VT_R4, VT_R8) else int(v))
    return build(head, g)


def new_blob(name='DBGrid1', width=6495, height=3015, left=0, top=0):
    """the OleObjectBlob of a new DBGrid (the default layout: two blank columns), as VB6 makes one"""
    g = PSet(CLSID_GRID)
    font = {'Name': 'MS Sans Serif', 'Size': 8.25, 'Charset': 0, 'Weight': 400}
    sp = PSet(CLSID_SPLIT)
    sp.order = [0x206, 1, 3, 0x1F, 4, 5, 7, 6, 0xF, 0x10, 0x11, 0x203, 0]
    sp.values = {0x206: (VT_I4, 0), 1: (VT_BOOL, 0), 3: (VT_I2, 1), 0x1F: (VT_I4, 6), 4: (VT_BOOL, 0),
                 5: (VT_BOOL, -1), 7: (VT_I4, 0), 6: (VT_I2, 1), 0xF: (VT_BOOL, -1), 0x10: (VT_BOOL, -1),
                 0x11: (VT_I4, 4), 0x203: (VT_BLOB, _column_props(2, DEFAULT_COL_WIDTH))}
    sp.names = [(0, 'Split0'), (5, 'AllowFocus'), (0xF, 'AllowRowSizing'), (4, 'AllowSizing'), (1, 'Locked'),
                (0x1F, 'MarqueeStyle'), (0x10, 'RecordSelectors'), (0x11, 'ScrollBars'), (3, 'ScrollGroup'),
                (6, 'Size'), (7, 'SizeMode'), (0x203, '_ColumnProps'), (0x206, '_UserFlags')]
    vals = [(0x7D3, VT_I4, int(round(width * 2540 / 1440)), '_ExtentX'),
            (0x7D4, VT_I4, int(round(height * 2540 / 1440)), '_ExtentY'),
            (0x7D7, VT_I4, 0x10000, '_Version'), (0x200, VT_I2, 0, '_LayoutType'),
            (0x10, VT_I4, -2147483647, '_RowHeight'), (8, VT_SETS, [_column(0), _column(1)], 'Columns'),
            (0x23, VT_SETS, [sp], 'Splits'), (1, VT_BOOL, 0, 'AllowDelete'), (2, VT_BOOL, 0, 'AllowAddNew'),
            (4, VT_BOOL, -1, 'AllowUpdate'), (-520, VT_I4, 1, 'Appearance'), (-504, VT_I4, 1, 'BorderStyle'),
            (7, VT_BOOL, -1, 'ColumnHeaders'), (0x25, VT_I4, 0, 'DataMode'), (0xA, VT_R4, 0.0, 'DefColWidth'),
            (-514, VT_BOOL, -1, 'Enabled'), (0xB, VT_FONT, font_bytes(font), 'HeadFont'),
            (0xC, VT_R4, 1.0, 'HeadLines'), (0x4A, VT_BOOL, -1, 'MarqueeUnique'), (0xF, VT_I2, 2, 'RowDividerStyle'),
            (-501, VT_I4, -2147483643, 'BackColor'), (-513, VT_I4, -2147483640, 'ForeColor'),
            (-518, VT_LPSTR, '', 'Caption'), (-512, VT_FONT, font_bytes(font), 'Font'),
            (0x2F, VT_BOOL, -1, 'AllowArrows'), (0x31, VT_BOOL, 0, 'TabAcrossSplits'), (0x32, VT_I4, 0, 'TabAction'),
            (0x33, VT_BOOL, 0, 'WrapCellPointer')]
    g.order = [k for k, _, _, _ in vals] + [0]
    g.values = {k: (vt, v) for k, vt, v, _ in vals}
    names = {k: n for k, _, _, n in vals}
    dict_order = [2, 0x2F, 1, 4, -520, -501, -504, -518, 7, 8, 0x25, 0xA, -514, -512, -513, 0xB, 0xC, 0x4A, 0xF,
                  0x23, 0x31, 0x32, 0x33, 0x7D3, 0x7D4, 0x200, 0x10, 0x7D7]
    g.names = [(0, name)] + [(k, names[k]) for k in dict_order]
    return build([int(left), int(top), int(width), int(height)], g)


# ---------------------------------------------------------------- Bindings
# the data bindings of an .ocx control in the .frx (Bindings = "Form1.frx":0000): records
# [u32 0x0001FAC6][u32 dispid][u8 n][name][6 zero bytes]; DBGrid DataSource 9, DBList RowSource 3,
# DBCombo RowSource 16 (VB6-saved forms: tests/dev/f_dbgrid)
BIND_MAGIC = b'\xc6\xfa\x01\x00'
DISPID_DATASOURCE = {'DBGrid': 9, 'DataGrid': 42}
DISPID_ROWSOURCE = {'DBList': 3, 'DBCombo': 16}


def bindings(data):
    """[(dispid, name)]"""
    out = []
    d = bytes(data or b'')
    i = d.find(BIND_MAGIC)
    while 0 <= i and i + 9 <= len(d):
        k, n = _u32(d, i + 4), d[i + 8]
        out.append((k, d[i + 9:i + 9 + n].decode('cp1252', 'replace')))
        i = d.find(BIND_MAGIC, i + 9 + n)
    return out


def bindings_bytes(pairs):
    out = bytearray()
    for k, name in pairs:
        b = name.encode('cp1252', 'replace')[:255]
        out += BIND_MAGIC + struct.pack('<IB', k, len(b)) + b + b'\0' * 6
    return bytes(out)


def bindings_set(data, dispid, name):
    """the Bindings with dispid bound to name ('' takes it away); b'' when none is left"""
    pairs = [(k, n) for k, n in bindings(data) if k != dispid]
    if name:
        pairs.append((dispid, name))
    return bindings_bytes(pairs)


if __name__ == '__main__':
    import sys
    raw = open(sys.argv[1], 'rb').read()
    off = int(sys.argv[2], 16) if len(sys.argv) > 2 else 0x14
    end = raw.find(b'\xc6\xfa\x01\x00', off + 4)
    rec = raw[off:end if end > 0 else len(raw)]
    for k, v in decode(rec):
        print(f'{k} = {v!r}')
    print('round trip:', build(*parse(rec), at=off) == rec)
