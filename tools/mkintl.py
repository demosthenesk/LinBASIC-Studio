#!/usr/bin/env python3
# LinBASIC Studio - writes runtime/intl_tables.inc: the ANSI code pages of Windows (single and double byte),
# the case pairs of Unicode and the decompositions used by text comparison. The mappings come from the code
# page codecs of Python and the Unicode Character Database (unicodedata), with the differences that VB6 SP6
# showed on Windows (tests/lang/t_intl.bas): the private-use areas, the undefined bytes, the euro of 936.
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#   python3 tools/mkintl.py
import codecs, os, unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'runtime', 'intl_tables.inc')

SBCS = [874, 1250, 1251, 1252, 1253, 1254, 1255, 1256, 1257, 1258]
DBCS = [932, 936, 949, 950]

# undefined single bytes: Windows gives them a private-use character (or the byte itself)
SB_PUA = {
    1253: {0xAA: 0xF8F9, 0xD2: 0xF8FA, 0xFF: 0xF8FB},
    1255: {**{b: 0xF88D + i for i, b in enumerate(range(0xD9, 0xE0))}, 0xFB: 0xF894, 0xFC: 0xF895, 0xFF: 0xF896},
    874: {**{b: 0xF8C1 + i for i, b in enumerate(range(0xDB, 0xDF))}, **{b: 0xF8C5 + i for i, b in enumerate(range(0xFC, 0x100))}},
}
DB_SINGLE = {932: {0x80: 0x80, 0xA0: 0xF8F0, 0xFD: 0xF8F1, 0xFE: 0xF8F2, 0xFF: 0xF8F3},
             936: {0x80: 0x20AC, 0xFF: 0xF8F5}, 949: {0x80: 0x80, 0xFF: 0xF8F7}, 950: {0x80: 0x80, 0xFF: 0xF8F8}}
DEFCHAR = {932: 0x30FB}
LEADS = {932: list(range(0x81, 0xA0)) + list(range(0xE0, 0xFD)), 936: range(0x81, 0xFF), 949: range(0x81, 0xFF),
         950: range(0x81, 0xFF)}
SB_FIX = {1255: {0xCA: 0x05BA}}                 # (newer than Python's table)                         # a broken double-byte character (others: '?')


def user_areas(cp):
    """the user-defined areas of the double-byte code pages: (lead, trail) -> private use, in Windows order"""
    out = {}
    def area(leads, trails, start):
        u = start
        for l in leads:
            for t in trails:
                out[(l, t)] = u
                u += 1
        return u
    if cp == 932:
        area(range(0xF0, 0xFA), list(range(0x40, 0x7F)) + list(range(0x80, 0xFD)), 0xE000)
    elif cp == 936:
        u = area(range(0xAA, 0xB0), range(0xA1, 0xFF), 0xE000)
        u = area(range(0xF8, 0xFF), range(0xA1, 0xFF), u)
        area(range(0xA1, 0xA8), list(range(0x40, 0x7F)) + list(range(0x80, 0xA1)), u)
    elif cp == 949:
        u = area([0xC9], range(0xA1, 0xFF), 0xE000)
        area([0xFE], range(0xA1, 0xFF), u)
    elif cp == 950:
        tr = list(range(0x40, 0x7F)) + list(range(0xA1, 0xFF))
        u = area(range(0xFA, 0xFF), tr, 0xE000)
        u = area(range(0x8E, 0xA1), tr, u)
        u = area(range(0x81, 0x8E), tr, u)
        u = area([0xC6], [t for t in tr if t >= 0xA1], u)
        area([0xC7, 0xC8], tr, u)
    return out


def sbcs_table(cp):
    t = []
    for b in range(0x80, 0x100):
        try:
            u = ord(bytes([b]).decode('cp%d' % cp))
        except UnicodeDecodeError:
            u = SB_FIX.get(cp, {}).get(b) or SB_PUA.get(cp, {}).get(b, b)
        t.append(u)
    return t


def dbcs_tables(cp):
    name = 'cp%d' % cp
    ua = user_areas(cp)
    leads, single = [], []
    for b in range(0x80, 0x100):
        if b in DB_SINGLE[cp]:
            single.append(DB_SINGLE[cp][b])
            continue
        try:
            single.append(ord(bytes([b]).decode(name)))
            continue
        except UnicodeDecodeError:
            pass
        if b in LEADS[cp]:
            leads.append(b)
            single.append(0)
        else:
            single.append(b)
    rows = {}
    for l in leads:
        row = []
        for t in range(0x40, 0xFF):
            u = ua.get((l, t)) or _dec(name, l, t)
            row.append(u or 0)
        rows[l] = row
    return single, rows


def _dec(name, l, t):
    try:
        s = bytes([l, t]).decode(name)
        return ord(s) if len(s) == 1 and ord(s) < 0x10000 else None
    except UnicodeDecodeError:
        return None


def case_pairs():
    """lower -> upper where the mapping goes both ways (Windows keeps ς, ı, ǅ ...)"""
    pairs = []
    for c in range(0x10000):
        ch = chr(c)
        u = ch.upper()
        if len(u) != 1 or u == ch:
            continue
        if u.lower() != ch:
            continue
        if 0x1C90 <= ord(u) <= 0x1CBF or 0x13A0 <= ord(u) <= 0x13FF or 0xAB70 <= c <= 0xABBF:
            continue                            # (newer than the tables of Windows)
        pairs.append((c, ord(u)))
    return pairs


def decomps():
    """character -> (base characters, marks) for text comparison: NFKD, only where it says something"""
    out = []
    ranges = [(0xA0, 0x250), (0x1E00, 0x2000), (0x370, 0x400), (0x400, 0x530), (0x2070, 0x20A0), (0x2150, 0x2190),
              (0x2460, 0x24FF), (0x3040, 0x3100), (0xFB00, 0xFB07), (0xFF00, 0xFFF0)]
    for lo, hi in ranges:
        for c in range(lo, hi):
            d = unicodedata.normalize('NFKD', chr(c))
            if d == chr(c):
                continue
            base = ''.join(x for x in d if not unicodedata.combining(x))
            marks = ''.join(x for x in d if unicodedata.combining(x))
            if not base or len(base) > 3 or len(marks) > 3 or any(ord(x) > 0xFFFF for x in d):
                continue
            if base == ' ' and marks:
                continue
            out.append((c, [ord(x) for x in base], [ord(x) for x in marks]))
    out.sort()
    return out


def arr(name, ctype, vals, per=16):
    lines = [f"static const {ctype} {name}[{len(vals)}] = {{"]
    for i in range(0, len(vals), per):
        lines.append('    ' + ','.join(f"0x{v:X}" for v in vals[i:i + per]) + ',')
    lines.append('};')
    return lines


def main():
    L = ['/* intl_tables.inc - generated by tools/mkintl.py (Python codecs + unicodedata ' +
         unicodedata.unidata_version + '); do not edit */']
    for cp in SBCS:
        L += arr(f"sb{cp}", 'vb_u16', sbcs_table(cp))
    for cp in DBCS:
        single, rows = dbcs_tables(cp)
        L += arr(f"db{cp}_single", 'vb_u16', single)
        leads = sorted(rows)
        idx = [0] * 128
        flat = []
        for k, l in enumerate(leads):
            idx[l - 0x80] = k + 1
            flat += rows[l]
        L += arr(f"db{cp}_lead", 'unsigned char', idx)
        L += arr(f"db{cp}_rows", 'vb_u16', flat)
        # characters with more than one code: the one Windows writes (Python's encoder), when not the lowest
        codes = {}
        for l in leads:
            for t, u in enumerate(rows[l]):
                if u:
                    codes.setdefault(u, []).append((l << 8) | (t + 0x40))
        pref = []
        for u, cs in sorted(codes.items()):
            if len(cs) < 2:
                continue
            try:
                e = chr(u).encode('cp%d' % cp)
            except UnicodeEncodeError:
                continue
            if len(e) == 2 and ((e[0] << 8) | e[1]) != min(cs) and ((e[0] << 8) | e[1]) in cs:
                pref += [u, (e[0] << 8) | e[1]]
        L += arr(f"db{cp}_pref", 'vb_u16', pref or [0, 0])
        L.append(f"#define DB{cp}_NPREF {len(pref) // 2}")
    L.append('static const cpdef cpdefs[] = {')
    for cp in SBCS:
        L.append(f"    {{ {cp}, 0, sb{cp}, 0, 0, 0x3F, 0, 0 }},")
    for cp in DBCS:
        L.append(f"    {{ {cp}, 1, db{cp}_single, db{cp}_lead, db{cp}_rows, 0x{DEFCHAR.get(cp, 0x3F):X}, db{cp}_pref, DB{cp}_NPREF }},")
    L.append('    { 0, 0, 0, 0, 0, 0, 0, 0 }')
    L.append('};')
    pairs = case_pairs()
    L += arr('case_lower', 'vb_u16', [a for a, b in pairs])
    L += arr('case_upper', 'vb_u16', [b for a, b in pairs])
    by_upper = sorted(pairs, key=lambda p: p[1])
    L += arr('case_u2', 'vb_u16', [b for a, b in by_upper])
    L += arr('case_l2', 'vb_u16', [a for a, b in by_upper])
    L.append(f"#define NCASE {len(pairs)}")
    ds = decomps()
    flat = []
    for c, base, marks in ds:
        flat += [c, len(base) | (len(marks) << 4)] + base + marks
    L += arr('decomp', 'vb_u16', flat)
    L.append(f"#define NDECOMP {len(ds)}")
    # Latin letters with marks: the letter alone (the best fit of Windows)
    base = []
    for c in range(0xC0, 0x250):
        d = unicodedata.normalize('NFD', chr(c))
        b = d[0] if d[0] != chr(c) and ord(d[0]) < 0x80 else '\0'
        base.append(ord(b))
    L += arr('latin_base', 'unsigned char', base)
    open(OUT, 'w').write('\n'.join(L) + '\n')
    print(f"{OUT}: {len(SBCS) + len(DBCS)} code pages, {len(pairs)} case pairs, {len(ds)} decompositions")


if __name__ == '__main__':
    main()
