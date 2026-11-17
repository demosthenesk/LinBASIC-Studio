#!/usr/bin/env python3
# LinBASIC Studio - the object model of MSChart (MSChart20Lib) as a C table for the forms runtime
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""Usage:  python3 tools/rea/gen_chart_model.py  ->  runtime/forms/ocx_chart_model.inc

The parts of an MSChart (Plot, Axis, Series, VtColor ...) are objects of the runtime whose
properties are kept by path ("plot.axis(1).valuescale.maximum"); this table says, for each
interface, which members give another part, which give a value (and its kind) and which are
methods. Only names and kinds are taken from the type library."""
import json, os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, '..', '..')


def kind_of(ret, ifaces, enums):
    if ret in ifaces:
        return 'o'
    if ret == 'Boolean':
        return 'b'
    if ret in ('Single', 'Double'):
        return 'f'
    if ret == 'String':
        return 's'
    if ret in ('Integer', 'Long') or ret in enums:
        return 'i'
    return 'v'


def main():
    d = json.load(open(os.path.join(ROOT, 'data', 'vb6', 'typelibs', 'MSChart20Lib.json'), encoding='utf-8'))
    types = d['types']
    ifaces = {t['name'] for t in types if t['kind'] == 'dispatch' and t['name'].startswith('IVc')}
    enums = {t['name'] for t in types if t['kind'] == 'enum'}
    tname = {}                                   # IVcPlot -> Plot (TypeName)
    for t in types:
        if t['kind'] == 'coclass':
            for im in t.get('implements', []):
                if im.get('default') and im['name'] in ifaces:
                    tname[im['name']] = t['name']
    rows, defs = [], []
    for t in types:
        if t['name'] not in ifaces:
            continue
        seen = {}
        dm = None
        for m in t['members']:
            if m['name'].startswith('_'):
                continue
            inv = m.get('invoke')
            key = m['name'].lower()
            if m.get('memid') == 0:
                dm = key
            e = seen.setdefault(key, {'name': key, 'kind': None, 'ret': '', 'args': 0, 'let': False})
            if inv == 'get':
                ret = m.get('returns', '')
                e['kind'] = kind_of(ret, ifaces, enums)
                e['ret'] = ret if e['kind'] == 'o' else ''
                e['args'] = len([p for p in m.get('params', []) if 'retval' not in p.get('flags', [])])
            elif inv in ('let', 'set'):
                e['let'] = True
                if e['kind'] is None:
                    e['kind'] = 'v'
            elif inv == 'func':
                if m.get('returns') not in (None, 'HRESULT', 'void'):
                    e['kind'] = kind_of(m.get('returns'), ifaces, enums)
                    e['args'] = len([p for p in m.get('params', []) if 'retval' not in p.get('flags', [])])
                else:
                    e['kind'] = 'm'
        for e in seen.values():
            rows.append(f'    {{ "{t["name"]}", "{e["name"]}", \'{e["kind"]}\', {1 if e["let"] else 0}, {e["args"]}, '
                        f'"{e["ret"]}" }},')
        defs.append(f'    {{ "{t["name"]}", "{tname.get(t["name"], t["name"][3:])}", "{dm or ""}" }},')
    out = ['/* ocx_chart_model.inc - made by tools/rea/gen_chart_model.py from the MSChart20Lib type library',
           ' * (names and kinds only). Do not edit. */',
           'struct ChartMember { const char *iface, *name; char kind; int let, args; const char *ret; };',
           '/* kind: o a part (ret: its interface), b Boolean, i whole number, f Single/Double, s String,',
           ' * v Variant, m a method */',
           'static const ChartMember chart_members[] = {'] + rows + ['};',
           'struct ChartIface { const char *iface, *tname, *defmember; };',
           'static const ChartIface chart_ifaces[] = {'] + defs + ['};', '']
    p = os.path.join(ROOT, 'runtime', 'forms', 'ocx_chart_model.inc')
    open(p, 'w', encoding='utf-8').write('\n'.join(out))
    print(p, len(rows), 'members', len(defs), 'interfaces')


if __name__ == '__main__':
    main()
