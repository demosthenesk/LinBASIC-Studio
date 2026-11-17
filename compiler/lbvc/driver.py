# LinBASIC Studio - lbvc driver: VB6 project -> C -> native executable
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
import argparse, os, shutil, subprocess, sys, tempfile

from .lexer import VBError
from .parser import parse_module
from .cgen import Compiler, CompInfo
from .project import load_project, load_source

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_ROOT = os.path.normpath(os.path.join(HERE, '..', '..'))
NO_WINDOW = {'creationflags': 0x08000000} if os.name == 'nt' else {}   # CREATE_NO_WINDOW: tcc from the IDE


def host_target():
    """the system lbvc runs on: what a program is built for by default (the IDE of Windows: win32/win64)"""
    if os.name != 'nt':
        return 'linux'
    import struct
    return 'win64' if struct.calcsize('P') == 8 else 'win32'


def find_runtime():
    rt = os.environ.get('LBVC_RUNTIME') or os.path.join(SRC_ROOT, 'runtime')
    if os.name == 'nt':                         # the IDE of Windows: vbrt.h, <target>/dll and <target>/tcc
        if not os.path.exists(os.path.join(rt, 'vbrt.h')):
            raise SystemExit("lbvc: vbrt.h not found (set LBVC_RUNTIME)")
        return rt
    lib = os.path.join(rt, 'libvbrt.a')
    if not os.path.exists(lib) and os.path.exists(os.path.join(rt, 'build.sh')):
        subprocess.run(['sh', os.path.join(rt, 'build.sh')], check=True, stdout=subprocess.DEVNULL)
    if not os.path.exists(lib):
        raise SystemExit("lbvc: libvbrt.a not found (set LBVC_RUNTIME)")
    return rt


def load_module(path, encoding, consts=None):
    """(module name, syntax tree) of a .bas file; consts: the CondComp constants of the project"""
    sf = load_source(path, encoding=encoding)
    if sf.kind not in ('Module', 'Class', 'Form', 'UserControl'):
        what = {'UserControl': 'user controls', 'UserDocument': 'user documents',
                'PropertyPage': 'property pages'}.get(sf.kind, sf.kind)
        raise VBError(f"{what} are not supported yet", 1, 1, path)
    if sf.kind == 'Form' and sf.root is not None and sf.root.type_name not in ('Form', 'MDIForm', 'UserControl'):
        raise VBError(f"{sf.root.type_name} is not supported yet", 1, 1, path)
    name = sf.name or os.path.splitext(os.path.basename(path))[0]
    hl = header_lines(path, encoding)
    code = sf.code
    if any(l.lstrip().startswith('#') for l in code):
        from .condcomp import preprocess
        code = preprocess(code, consts or {}, hl + 1, path)
    text = '\n' * hl + '\n'.join(code)    # keep the line numbers of the file
    ast = parse_module(text, path, name)
    ast.is_class = sf.kind in ('Class', 'Form', 'UserControl')
    ast.vb_attrs = {n.lower(): v.strip('"').strip().lower() for n, v, _ in sf.attributes}
    # MultiUse = 0 in the header of a class: SingleUse (an ActiveX EXE: a copy of the server for each object)
    ast.multiuse = not any(p.name.lower() == 'multiuse' and p.value in (0, '0') for p in (sf.class_props or []))
    ast.is_form = sf.kind in ('Form', 'UserControl')       # a user control: a form inside another form
    ast.is_uctl = sf.root is not None and sf.root.type_name == 'UserControl'
    ast.form_root = sf.root if ast.is_form else None
    ast.form_dir = os.path.dirname(os.path.abspath(path))
    ast.form_enc = encoding
    return name, ast


def load_designer(path, encoding, consts=None):
    """the modules of a designer (.dsr): a DataEnvironment becomes a class with a default instance"""
    from .designers import designer_kind, dataenv_source, datareport_source
    sf = load_source(path, encoding=encoding)
    kind = designer_kind(sf)
    if kind not in ('DataEnvironment', 'DataReport'):
        raise VBError(f"the designer {sf.root.type_name if sf.root else ''} ({kind or 'unknown'}) is not supported yet", 1, 1, path)
    hl = header_lines(path, encoding)
    out = []
    mods = dataenv_source(sf) if kind == 'DataEnvironment' else datareport_source(sf, path)
    for name, (decls, user, procs), attrs in mods:
        if user:                                # the code of the .dsr keeps its line numbers
            if len(decls) > hl:
                hl2 = len(decls)
                lines = decls + user + procs
            else:
                lines = decls + [''] * (hl - len(decls)) + user + procs
        else:
            lines = decls + user + procs
        code = lines
        if any(l.lstrip().startswith('#') for l in code):
            from .condcomp import preprocess
            code = preprocess(code, consts or {}, 1, path)
        ast = parse_module('\n'.join(code), path, name)
        ast.is_class = True
        ast.vb_attrs = {n.lower(): str(v).strip('"').strip().lower() for n, v in attrs}
        ast.multiuse = True
        ast.is_form = False
        ast.is_uctl = False
        ast.form_root = None
        ast.form_dir = os.path.dirname(os.path.abspath(path))
        ast.form_enc = encoding
        out.append((name, ast))
    return out


def header_lines(path, encoding):
    """number of lines before the code: VERSION, BEGIN ... END block, Attribute lines"""
    with open(path, 'rb') as f:
        lines = f.read().decode(encoding, 'replace').splitlines()
    n, depth = 0, 0
    for ln in lines:
        u = ln.strip().upper()
        if depth or u.startswith('VERSION ') or u.startswith('BEGIN') or u.startswith('OBJECT ='):
            if u.startswith('BEGIN') or u.startswith('BEGINPROPERTY'):
                depth += 1
            elif u in ('END', 'ENDPROPERTY'):
                depth -= 1
            n += 1
            continue
        if u.startswith('ATTRIBUTE '):
            n += 1
            continue
        break
    return n


WARNINGS = []                                   # of the last compile_project


class CSource(str):
    """the C of a program, with the C units of the components compiled into it (.units)"""
    units = ()
    infos = ()
    component = False


COMPONENT_TYPES = {'oledll': 'an ActiveX DLL', 'oleexe': 'an ActiveX EXE', 'control': 'an ActiveX Control'}


# ---------------------------------------------------------------- the registry of components
def registry_path():
    """components.json: the ActiveX DLL / EXE projects made with LinBASIC Studio (Make registers them)"""
    p = os.environ.get('LBVC_REGISTRY')
    if p:
        return p
    if os.path.exists(os.path.join(SRC_ROOT, 'portable.ini')):     # the portable package of Windows
        return os.path.join(SRC_ROOT, 'settings', 'components.json')
    if os.name == 'nt':
        base = os.environ.get('APPDATA') or os.path.expanduser('~')
    else:
        base = os.environ.get('XDG_CONFIG_HOME') or os.path.join(os.path.expanduser('~'), '.config')
    return os.path.join(base, 'linbasic-studio', 'components.json')


def load_registry():
    import json
    try:
        with open(registry_path(), encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def register_component(vbp_path, encoding='cp1252'):
    """Make of an ActiveX project: the registry names it (as VB6 registers the .dll / .exe)"""
    import json
    pr = load_project(vbp_path, encoding)
    reg = load_registry()
    classes, controls = [], []
    for kind, mname, rel in pr.members():
        if kind == 'Class':
            classes.append(mname)
        elif kind == 'UserControl':
            controls.append(mname)
    reg[pr.name.lower()] = {'name': pr.name, 'vbp': os.path.abspath(vbp_path), 'type': pr.type, 'classes': classes,
                            'controls': controls,
                            'exe': os.path.basename((pr.get('ExeName32') or '').strip('"').replace('\\', '/')) or
                            pr.name + {'oledll': '.dll', 'control': '.ocx'}.get(pr.type.lower(), '.exe')}
    os.makedirs(os.path.dirname(registry_path()), exist_ok=True)
    with open(registry_path(), 'w', encoding='utf-8') as f:
        json.dump(reg, f, indent=1, ensure_ascii=False)
    return pr.name


def _project_info(vbp, encoding):
    try:
        pr = load_project(vbp, encoding)
        return pr.name, pr.type.lower()
    except Exception:                           # noqa: BLE001 - a .vbp that cannot be read is no component
        return None, None


def find_components(pr, vbp_path, encoding):
    """the .vbp of the components a project uses: its project references (Reference=*\\A..., in a group),
    the projects named in its code (CreateObject("Proj.Class"), New / As Proj.Class) that the registry
    knows or that are in the same folder"""
    import re, glob
    out = []
    base = os.path.dirname(os.path.abspath(vbp_path))
    from .project.fsutil import ci_path
    ocx_files = set()                           # Object={GUID}#1.0#0; CtlPlus.ocx: an ActiveX control project
    for k, v in pr.entries:
        if v is None or k.strip() not in ('Reference', 'Object'):
            continue
        m = re.match(r'\s*\*\\A(.+?)\s*$', v)
        if m:
            p = ci_path(os.path.join(base, m.group(1).replace('\\', '/')))
            if os.path.exists(p):
                out.append(os.path.abspath(p))
        elif k.strip() == 'Object' and ';' in v:
            ocx_files.add(v.rsplit(';', 1)[1].strip().lower())
    names, bare = set(), set()
    for kind, mname, rel in pr.members():
        try:
            text = open(pr.resolve(rel), 'rb').read().decode(encoding, 'replace')
        except OSError:
            continue
        # a form with the controls of a control project: Object = "*\ACtlPlus.vbp", Begin CtlPlus.ShapeLabel
        for rp in re.findall(r'(?im)^Object\s*=\s*"\*\\A([^"]+)"', text):
            p = ci_path(os.path.join(os.path.dirname(pr.resolve(rel)), rp.replace('\\', '/')))
            if os.path.exists(p):
                out.append(os.path.abspath(p))
        names |= {a.lower() for a in re.findall(r'(?im)^\s*Begin\s+([A-Za-z_]\w*)\.\w+\s+\w+', text)} - {'vb'}
        text = re.sub(r"'[^\n]*", '', text)
        # "Proj.Class" in any string (CreateObject("..."), a constant with the ProgID ...)
        names |= {a.lower() for a in re.findall(r'"\s*([A-Za-z_]\w*)\.[A-Za-z_]\w*\s*"', text)}
        names |= {a.lower() for a in re.findall(r'(?i)\b(?:new|as)\s+([A-Za-z_]\w*)\.\w+', text)}
        # a class named without its project (As CoffeeMonitor): a component that has it
        bare |= {a.lower() for a in re.findall(r'(?i)\b(?:new|as|implements)\s+([A-Za-z_]\w*)\b(?!\.)', text)}
    names.discard(pr.name.lower())
    own = {m.lower() for _, m, _ in pr.members()}
    bare -= own
    reg = load_registry()
    local, local_classes = {}, {}               # the component projects next to this one
    for v in glob.glob(os.path.join(base, '*.[vV][bB][pP]')):
        if os.path.abspath(v) == os.path.abspath(vbp_path):
            continue
        try:
            lp = load_project(v, encoding)
        except Exception:                       # noqa: BLE001 - a .vbp that cannot be read is no component
            continue
        if lp.type.lower() not in COMPONENT_TYPES:
            continue
        local[lp.name.lower()] = os.path.abspath(v)
        if os.path.basename(lp.exe_name).lower() in ocx_files:
            out.append(os.path.abspath(v))
        for kind, mname, _ in lp.members():
            if kind == 'Class':
                local_classes.setdefault(mname.lower(), os.path.abspath(v))
    for e in reg.values():
        if e.get('exe', '').lower() in ocx_files and os.path.exists(e.get('vbp', '')):
            out.append(os.path.abspath(e['vbp']))
    for n in sorted(names):
        e = reg.get(n)
        if e and os.path.exists(e.get('vbp', '')):
            out.append(os.path.abspath(e['vbp']))
        elif n in local:
            out.append(local[n])
    for n in sorted(bare):
        if n in local_classes:
            out.append(local_classes[n])
            continue
        for e in reg.values():
            if n in {c.lower() for c in e.get('classes', [])} and os.path.exists(e.get('vbp', '')):
                out.append(os.path.abspath(e['vbp']))
                break
    seen, res = set(), []
    for p in out:
        if p not in seen and os.path.abspath(p) != os.path.abspath(vbp_path):
            seen.add(p)
            res.append(p)
    return res


# the type libraries a project references (Reference=*\G{GUID}...): DAO, ADO
TYPELIB_GUIDS = {
    '{00025e01-0000-0000-c000-000000000046}': 'dao',      # DAO 3.6 (and 3.5, 3.51)
    '{00025e04-0000-0000-c000-000000000046}': 'dao',      # DAO 2.5/3.0 compatibility
    '{00000200-0000-0010-8000-00aa006d2ea4}': 'adodb',    # ADO 2.0
    '{00000201-0000-0010-8000-00aa006d2ea4}': 'adodb',    # ADO 2.1
    '{00000205-0000-0010-8000-00aa006d2ea4}': 'adodb',    # ADO 2.5
    '{00000206-0000-0010-8000-00aa006d2ea4}': 'adodb',    # ADO 2.6
    '{ef53050b-882e-4776-b643-eda472e8e3f2}': 'adodb',    # ADO 2.7
    '{2a75196c-d9eb-4129-b803-931327f72d5c}': 'adodb',    # ADO 2.8
    '{b691e011-1797-432e-907a-4d8c69339129}': 'adodb',    # ADO 6.0
    '{56bf9020-7a2f-11d0-9482-00a0c91110ed}': 'msbind',   # Microsoft Data Binding Collection
    '{6b263850-900b-11d0-9484-00a0c91110ed}': 'stdformat',   # Microsoft Data Formatting Object Library
    '{420b2830-e718-11cf-893d-00a0c9054228}': 'scripting',   # Microsoft Scripting Runtime (its constants)
    '{ee008642-64a8-11ce-920f-08002b369a33}': 'rdo',      # Microsoft Remote Data Object 2.0
}


class TypeLibs(set):
    """the type libraries of a project; .order: as the project lists them (Recordset of DAO or of ADO:
    the one listed first, as VB6 resolves it)"""
    order = ()


def project_typelibs(pr):
    out = TypeLibs()
    order = []
    for r in pr.references():
        lib = TYPELIB_GUIDS.get(r.guid.lower())
        if lib:
            out.add(lib)
            if lib not in order:
                order.append(lib)
    out.order = tuple(order)
    return out


def uses_db(csrc):
    return '/* lbvc:db */' in csrc or any('/* lbvc:db */' in u for _, u in getattr(csrc, 'units', ()))


ORDER_MODES = ('pcode', 'fast', 'small', 'noopt')
ORDER_FLAGS = ('overflow', 'bounds', 'flpoint', 'unrounded')
PCODE = ('pcode', False, False, False, False)


def parse_order(spec):
    """--order MODE[,overflow][,bounds][,flpoint][,unrounded] -> (mode, Remove Integer Overflow Checks,
    Remove Array Bounds Checks, Remove Floating Point Error Checks, Allow Unrounded Floating Point Operations)"""
    parts = [x.strip().lower() for x in spec.split(',') if x.strip()]
    if not parts or parts[0] not in ORDER_MODES or any(x not in ORDER_FLAGS for x in parts[1:]):
        raise ValueError(f"invalid --order: {spec} (pcode, fast, small or noopt, then ,overflow ,bounds ,flpoint "
                         f",unrounded)")
    mode = 'fast' if parts[0] == 'small' else parts[0]
    if mode == 'pcode':                         # P-code has none of the advanced optimizations (VB6)
        return PCODE
    return (mode,) + tuple(f in parts[1:] for f in ORDER_FLAGS)


def project_order(pr):
    """the code VB6 makes with the settings of the project (Compile tab): P-code, native Fast/Small or No
    Optimization, and the advanced optimizations that change what a program does"""
    def flag(key):
        return (pr.get(key, '0') or '0').strip() not in ('0', '')
    if (pr.get('CompilationType', '0') or '0').strip() == '-1':
        return PCODE
    mode = 'noopt' if (pr.get('OptimizationType', '0') or '0').strip() == '2' else 'fast'
    return (mode, flag('OverflowCheck'), flag('BoundsCheck'), flag('FlPointCheck'), flag('UnroundedFP'))


def compile_project(inputs, encoding='cp1252', debug=False, as_component=False, _done=None, order=None, ide=False):
    """inputs: one .vbp, or .bas files -> CSource (the C of the program; .units: the C of the
    components it uses) and the project name. order: (mode, no overflow checks, no bounds checks) of
    the evaluation order (see evalorder.py); None: the settings of the .vbp (P-code for .bas files).
    ide: a run of the IDE (as VB6's Run): the order of P-code and Debug.Print; else a made program, which has no
    Debug.Print (VB6 leaves it out, with its arguments) unless debug"""
    modules, name, startup = [], 'Project1', 'Sub Main'
    from . import parser as _parser
    _parser.RECOVER = bool(debug or ide)          # Compile On Demand: a syntax error stays in its procedure
    if ide and order is None:
        order = PCODE
    order_set = order
    typelibs = set()
    app = {}
    comps, units = [], []
    missing = []
    _done = {} if _done is None else _done        # vbp -> (CompInfo, units) of components already compiled
    component = False
    if len(inputs) == 1 and inputs[0].lower().endswith('.vbp'):
        pr = load_project(inputs[0], encoding)
        name = pr.name
        exe = os.path.splitext(os.path.basename(pr.exe_name))[0]
        app = {'title': pr.get('Title', name), 'exename': exe,
               'major': int(pr.get('MajorVer', '1') or 1), 'minor': int(pr.get('MinorVer', '0') or 0),
               'revision': int(pr.get('RevisionVer', '0') or 0), 'comments': pr.get('VersionComments', ''),
               'companyname': pr.get('VersionCompanyName', ''), 'filedescription': pr.get('VersionFileDescription', ''),
               'legalcopyright': pr.get('VersionLegalCopyright', ''), 'legaltrademarks': pr.get('VersionLegalTrademarks', ''),
               'productname': pr.get('VersionProductName', ''), 'helpfile': pr.get('HelpFile', '')}
        startup = pr.startup or 'Sub Main'
        if order_set is None:
            order = project_order(pr)
        typelibs = project_typelibs(pr)
        ptype = pr.type.lower()
        if ptype not in ('exe',) and ptype not in COMPONENT_TYPES:
            raise VBError(f"this project is a {pr.type} project, which LinBASIC Studio does not compile", 0, 0, inputs[0])
        component = as_component or ptype in ('oledll', 'control')
        res = pr.get('ResFile32', '').strip('"')
        if res and not component:
            from .project.resfile import read_res
            rp = pr.resolve(res)
            if not os.path.exists(rp):
                if not (ide or debug):
                    raise VBError(f"resource file not found: {res}", 0, 0, inputs[0])
                # a run in the IDE: VB6 says so and goes on without it (LoadResString ... then fail: 326)
                missing.append(f"{inputs[0]}:0:0: warning: resource file not found: {res} (the project is loaded without it)")
            else:
                app['res'] = read_res(rp)
        from .condcomp import parse_condcomp
        consts = parse_condcomp(pr.get('CondComp', ''))
        for kind, mname, rel in pr.members():
            path = pr.resolve(rel)
            if kind == 'PropertyPage':              # design time only: the program does not need it
                continue
            if kind == 'Designer':
                modules.extend(load_designer(path, encoding, consts))
                continue
            if kind not in ('Module', 'Class', 'Form', 'UserControl'):
                raise VBError(f"{kind} {mname} is not supported yet", 0, 0, path)
            if (ide or debug) and not os.path.isfile(path):
                # a run in the IDE: VB6 says the file is not found and goes on loading the project without it
                missing.append(f"{inputs[0]}:0:0: warning: file not found: {rel} (the project is loaded without it)")
                continue
            modules.append(load_module(path, encoding, consts))
        # the components it uses: each compiled as a C unit of its own (its names cannot clash)
        for cv in find_components(pr, inputs[0], encoding):
            key = os.path.abspath(cv)
            if key not in _done:
                _done[key] = None                     # (a component that uses itself through another)
                # a run (IDE, --debug): the component Compile On Demand too, as VB6 runs a project group; one that
                # cannot be compiled at all (a type of a library this computer does not have: SSubTimer.EMsgResponse)
                # is left out, its controls empty boxes, as VB6 loads a form whose control does not load
                try:
                    csrc, cname = compile_project([cv], encoding, debug, True, _done, order_set, ide)
                except VBError as e:
                    if not (ide or debug):
                        raise
                    missing.append(f"{cv}:0:0: warning: the component could not be compiled ({e}): it is left out")
                    continue
                _done[key] = (list(csrc.infos) + [csrc.info], list(csrc.units) + [(cname, str(csrc))])
            if _done[key] is None:
                continue
            infos, cunits = _done[key]
            for info in infos:                    # with the components it uses itself (registered too)
                if all(info.name != c.name for c in comps):
                    comps.append(info)
            for u in cunits:
                if all(u[0] != x[0] for x in units):
                    units.append(u)
        # Implements of an interface of a component: the class is compiled into this program too
        have = {m.lower() for m, _ in modules}
        import re
        for mname, ast in list(modules):
            text = '\n'.join(open(ast.file, 'rb').read().decode(encoding, 'replace').splitlines())
            for iname in re.findall(r'(?im)^\s*implements\s+(?:\w+\.)?(\w+)', text):
                if iname.lower() in have:
                    continue
                for ci in comps:
                    c = ci.classes.get(iname.lower())
                    if c is not None and getattr(c, 'src_file', None):
                        m = load_module(c.src_file, encoding, consts)
                        m[1].vb_owner = ci.name           # the class of the component, not of this project
                        modules.append(m)
                        have.add(iname.lower())
                        break
    else:
        for p in inputs:
            modules.append(load_module(p, encoding))
        name = os.path.splitext(os.path.basename(inputs[0]))[0]
    comp = Compiler(modules, name, startup, app, debug=debug and not component, component=component,
                    components=comps)
    comp.ptype = ptype if len(inputs) == 1 and inputs[0].lower().endswith('.vbp') else 'exe'
    comp.typelibs = typelibs
    comp.from_vbp = len(inputs) == 1 and inputs[0].lower().endswith('.vbp')
    comp.order = order or PCODE
    comp.keep_debug = bool(debug or ide)
    comp.on_demand = bool(debug or ide)          # Compile On Demand, as a run in the VB6 IDE
    comp.on_demand_warn = not ide
    csrc = CSource(comp.compile())
    csrc.units = tuple(units)
    csrc.infos = tuple(comps)
    csrc.component = component
    if component:
        files = {m.lower(): a.file for m, a in modules}
        exposed = {}
        for k, ci in comp.classes.items():
            if ci.exposed and (not ci.is_form or ci.is_uctl):     # its user controls too
                ci.src_file = files.get(k)
                exposed[k] = ci
        csrc.info = CompInfo(name, comp.reg_name, exposed)
        csrc.info.consts = {key: (sy.value, sy.t) for k in exposed for key, sy in comp.mscopes.get(k, {}).items()
                            if sy.kind == 'const' and getattr(sy, 'enum_member', False) and sy.public}
        csrc.info.enums = {key for k in exposed for key, sy in comp.mscopes.get(k, {}).items()
                           if sy.kind == 'enum' and getattr(sy, 'public', True)}         # As enmMouseEvent in the program that uses it: a Long
        csrc.info.uses_forms = '/* lbvc:forms */' in csrc
        # its Public Types (of its public classes and user controls, and those it has from its own components): the
        # program that uses it gets a copy of each (Dim c As Comp.ChartItem, a Variant holding one, an event)
        csrc.info.types = {k: it for k, it in comp.udts.items()
                           if getattr(it, 'public_obj', False) or getattr(it, 'imported', False)}
    global WARNINGS
    WARNINGS = missing + [f"{f}:0:0: warning: {m}" for f, _, m in comp.warnings]
    return csrc, name


def _stale(lib, sources):
    """the library is missing or older than one of its sources"""
    if not os.path.exists(lib):
        return True
    t = os.path.getmtime(lib)
    return any(os.path.exists(s) and os.path.getmtime(s) > t for s in sources)


def ensure_runtime(rt):
    """build libvbrt.a (and libvbforms.a) again when the sources of the runtime are newer"""
    import glob
    if os.access(rt, os.W_OK) and _stale(os.path.join(rt, 'libvbrt.a'),
                                         glob.glob(os.path.join(rt, '*.c')) + glob.glob(os.path.join(rt, '*.h'))):
        subprocess.run(['sh', os.path.join(rt, 'build.sh')], check=True, stdout=subprocess.DEVNULL)


TARGETS = {'win64': 'x86_64-w64-mingw32', 'win32': 'i686-w64-mingw32'}


def mingw_cc(target, cxx=False):
    """the mingw-w64 compiler of a Windows target (posix threads: what Qt for Windows needs)"""
    t = TARGETS[target]
    tool = 'g++' if cxx else 'gcc'
    for name in (f"{t}-{tool}-posix", f"{t}-{tool}"):
        p = shutil.which(name)
        if p:
            return p
    raise SystemExit(f"lbvc: {t}-{tool} not found - install mingw-w64 (sudo apt install mingw-w64)")


def windows_kit(rt, target):
    """runtime/<target>: libvbrt.a, libvbforms.a and qt/ (bin, lib, plugins) of Qt for Windows"""
    kit = os.path.join(rt, target)
    lib = os.path.join(kit, 'libvbrt.a')
    if _stale(lib, [os.path.join(rt, f) for f in os.listdir(rt) if f.endswith(('.c', '.h', '.inc'))]) and os.access(rt, os.W_OK):
        r = subprocess.run(['sh', os.path.join(rt, 'build-win.sh'), target], capture_output=True, text=True)
        if r.returncode != 0:
            sys.stderr.write(r.stdout + r.stderr)
            raise SystemExit(f"lbvc: the runtime for {target} could not be built")
    if not os.path.exists(lib):
        raise SystemExit(f"lbvc: {lib} not found")
    return kit


def qt_major(kit):
    """the Qt of a Windows kit: 5 (Qt 5.15, Windows 7 and later) or 6"""
    return 5 if os.path.exists(os.path.join(kit, 'qt', 'bin', 'Qt5Core.dll')) else 6


def windows_dlls(kit, target, exe, forms):
    """the DLLs next to the .exe, so that the folder runs on any Windows (portable)"""
    dest = os.path.dirname(os.path.abspath(exe))
    names = []
    qt = os.path.join(kit, 'qt')
    if forms:
        q = f"Qt{qt_major(kit)}"
        for n in (f'{q}Core.dll', f'{q}Gui.dll', f'{q}Widgets.dll'):
            names.append((os.path.join(qt, 'bin', n), n))
        for sub, n in (('platforms', 'qwindows.dll'), ('imageformats', 'qgif.dll'), ('imageformats', 'qico.dll'),
                       ('imageformats', 'qjpeg.dll'), ('styles', 'qwindowsvistastyle.dll')):
            names.append((os.path.join(qt, 'plugins', sub, n), os.path.join(sub, n)))
    # the C++ runtime of mingw (Qt needs it; the program itself is linked with -static)
    if forms:
        for n in ('libstdc++-6.dll', 'libwinpthread-1.dll',
                  'libgcc_s_seh-1.dll' if target == 'win64' else 'libgcc_s_dw2-1.dll'):
            p = os.path.join(qt, 'bin', n)             # the kit has them (stripped)
            if not os.path.exists(p):
                p = subprocess.run([mingw_cc(target), f'-print-file-name={n}'], capture_output=True, text=True).stdout.strip()
            names.append((p, n))
    for src, rel in names:
        if os.path.exists(src):
            d = os.path.join(dest, os.path.dirname(rel))
            os.makedirs(d, exist_ok=True)
            shutil.copyfile(src, os.path.join(dest, rel))
        elif rel.startswith(('Qt5', 'Qt6', 'platforms')) or rel.startswith('lib'):
            raise SystemExit(f"lbvc: {src} not found (the Qt for {target} of LinBASIC Studio is missing)")


def write_units(csrc, td):
    """prog.c and the C files of the components compiled into the program"""
    c = os.path.join(td, 'prog.c')
    with open(c, 'w') as f:
        f.write(csrc)
    files = [c]
    for i, (_, src) in enumerate(getattr(csrc, 'units', ())):
        p = os.path.join(td, f'comp{i}.c')
        with open(p, 'w') as f:
            f.write(src)
        files.append(p)
    return files


def uses_forms(csrc):
    return '/* lbvc:forms */' in csrc or any('/* lbvc:forms */' in u for _, u in getattr(csrc, 'units', ()))


def use_tcc(target):
    """the Windows programs are built by the bundled Tiny C Compiler on Windows (the IDE of Windows),
    or here when LBVC_WIN_CC=tcc (to test that path); else by mingw-w64, the runtime linked in"""
    return os.name == 'nt' or os.environ.get('LBVC_WIN_CC') == 'tcc'


def windows_tcc(rt, target):
    """the tcc of a target and its folder (include/, lib/): runtime/<target>/tcc (packaging/build-tcc.sh)"""
    d = os.environ.get('LBVC_TCC_DIR') or os.path.join(rt, target, 'tcc')
    exe = os.path.join(d, 'tcc.exe' if os.name == 'nt' else 'cross-tcc')
    if not os.path.exists(exe):
        raise SystemExit(f"lbvc: {exe} not found (the Tiny C Compiler for {target}: packaging/build-tcc.sh)")
    return exe, d


def build_exe_windows_tcc(csrc, out, rt, target):
    """a Windows .exe with tcc: the program alone, the runtime in DLLs next to it (vbrt.dll, vbdb.dll,
    vbforms.dll: runtime/<target>/dll, build-dll-win.sh), and the DLLs of Qt for a program with forms"""
    kit = os.path.join(rt, target)
    dlld = os.path.join(kit, 'dll')
    tcc, tdir = windows_tcc(rt, target)
    forms, db = uses_forms(csrc), uses_db(csrc)
    if not out.lower().endswith('.exe'):
        out += '.exe'
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    libs = ['vbrt'] + (['vbdb'] if db else []) + (['vbforms'] if forms else [])
    for n in libs:
        if not os.path.exists(os.path.join(dlld, n + '.dll')):
            raise SystemExit(f"lbvc: {os.path.join(dlld, n + '.dll')} not found (runtime/build-dll-win.sh {target})")
    with tempfile.TemporaryDirectory() as td:
        cfiles = write_units(csrc, td)
        cmd = [tcc, '-B' + tdir, '-DVB_DLL_IMPORT', '-w', '-I', rt] + cfiles + ['-L', dlld] + \
            [f'-l{n}' for n in libs] + (['-Wl,-subsystem=gui'] if forms else []) + ['-o', out]
        r = subprocess.run(cmd, capture_output=True, text=True, **NO_WINDOW)
        if r.returncode != 0:
            sys.stderr.write(r.stdout + r.stderr)
            raise SystemExit("lbvc: C compiler failed (this is a bug in lbvc)")
    dest = os.path.dirname(os.path.abspath(out))
    for n in libs:
        shutil.copyfile(os.path.join(dlld, n + '.dll'), os.path.join(dest, n + '.dll'))
    windows_dlls(kit, target, out, forms)
    return out


def build_exe_windows(csrc, out, rt, target):
    """a Windows .exe with mingw-w64: a console program, or a GUI program with its DLLs (Qt)"""
    if use_tcc(target):
        return build_exe_windows_tcc(csrc, out, rt, target)
    kit = windows_kit(rt, target)
    forms = uses_forms(csrc)
    if not out.lower().endswith('.exe'):
        out += '.exe'
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    dblibs = []
    if uses_db(csrc):                          # DAO / ADO: the database runtime for Windows
        dl = os.path.join(kit, 'libvbdb.a')
        dsrc = [os.path.join(rt, 'db', f) for f in os.listdir(os.path.join(rt, 'db')) if f.endswith(('.c', '.h', '.inc'))]
        if _stale(dl, dsrc) and os.access(kit, os.W_OK):
            r = subprocess.run(['sh', os.path.join(rt, 'db', 'build-win.sh'), target], capture_output=True, text=True)
            if r.returncode != 0:
                sys.stderr.write(r.stdout + r.stderr)
                raise SystemExit(f"lbvc: the database runtime for {target} could not be built")
        if not os.path.exists(dl):
            raise SystemExit(f"lbvc: {dl} not found (the database runtime for {target} is missing)")
        dblibs = [dl]
    with tempfile.TemporaryDirectory() as td:
        cfiles = write_units(csrc, td)
        if forms:
            fl = os.path.join(kit, 'libvbforms.a')
            fsrc = [os.path.join(rt, 'forms', f) for f in os.listdir(os.path.join(rt, 'forms'))
                    if f.endswith(('.cpp', '.h', '.inc'))] + [os.path.join(rt, 'vbrt.h')]
            if _stale(fl, fsrc) and os.path.isdir(os.path.join(kit, 'qt', 'include')) and os.access(kit, os.W_OK):
                r = subprocess.run(['sh', os.path.join(rt, 'forms', 'build-win.sh'), target], capture_output=True, text=True)
                if r.returncode != 0:
                    sys.stderr.write(r.stdout + r.stderr)
                    raise SystemExit(f"lbvc: the forms runtime for {target} could not be built")
            if not os.path.exists(fl):
                raise SystemExit(f"lbvc: {fl} not found (the forms runtime for {target} is missing)")
            qt = os.path.join(kit, 'qt', 'lib')
            objs = []
            for c in cfiles:
                obj = c[:-2] + '.o'
                r = subprocess.run([mingw_cc(target), '-O1', '-w', '-I', rt, '-c', c, '-o', obj], capture_output=True, text=True)
                if r.returncode != 0:
                    sys.stderr.write(r.stderr)
                    raise SystemExit("lbvc: C compiler failed (this is a bug in lbvc)")
                objs.append(obj)
            cmd = [mingw_cc(target, cxx=True)] + objs + [fl] + dblibs + [os.path.join(kit, 'libvbrt.a'), '-lwinmm', '-lws2_32', '-lwininet', '-L', qt,
                   f'-lQt{qt_major(kit)}Widgets', f'-lQt{qt_major(kit)}Gui', f'-lQt{qt_major(kit)}Core',
                   '-mwindows', '-s', '-o', out]
        else:
            cmd = [mingw_cc(target), '-O1', '-w', '-I', rt] + cfiles + dblibs + [os.path.join(kit, 'libvbrt.a'), '-lwinmm', '-lws2_32', '-lwininet', '-lm',
                   '-static', '-s', '-o', out]
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0:
            sys.stderr.write(r.stderr)
            raise SystemExit("lbvc: C compiler failed (this is a bug in lbvc)")
    windows_dlls(kit, target, out, forms)
    return out


def linux_tcc():
    """the Tiny C Compiler the AppImage of LinBASIC Studio carries (packaging/build-tcc-linux.sh):
    the programs are built without a C compiler on the computer. LBVC_LINUX_TCC names its folder."""
    d = os.environ.get('LBVC_LINUX_TCC')
    if d and os.path.exists(os.path.join(d, 'tcc')):
        return [os.path.join(d, 'tcc'), '-B' + d]
    return None


def qt_link_flags():
    """the Qt 6 of a program with forms: LBVC_QT_LIBDIR (the Qt of the AppImage), else pkg-config"""
    d = os.environ.get('LBVC_QT_LIBDIR')
    if d and os.path.isdir(d):
        return [os.path.join(d, f'libQt6{n}.so.6') for n in ('Widgets', 'Gui', 'Core')] + ['-Wl,-rpath=' + d]
    return subprocess.run(['pkg-config', '--libs', os.environ.get('LBVC_QT_PKG', 'Qt6Widgets')],
                          capture_output=True, text=True).stdout.split()


def forms_so_for_tcc(rt, out):
    """a program with forms built by tcc: runtime/forms/libvbforms.so (packaging/build-appimage.sh)
    goes next to it (rpath $ORIGIN), Qt 6 from the AppImage while the IDE runs it, else the system's"""
    so = os.path.join(rt, 'forms', 'libvbforms.so')
    if not os.path.exists(so):
        raise SystemExit(f"lbvc: {so} not found (the forms runtime for tcc)")
    dest = os.path.dirname(os.path.abspath(out))
    os.makedirs(dest, exist_ok=True)
    shutil.copyfile(so, os.path.join(dest, 'libvbforms.so'))
    rpath = '$ORIGIN' + (':' + os.environ['LBVC_QT_LIBDIR'] if os.environ.get('LBVC_QT_LIBDIR') else '')
    return [os.path.join(dest, 'libvbforms.so'), '-Wl,-rpath=' + rpath]


def build_exe(csrc, out, rt, cc=None, optimize=True, debug_build=False, target='linux'):
    if target in TARGETS:
        return build_exe_windows(csrc, out, rt, target)
    tcc = None if (cc or os.environ.get('CC')) else linux_tcc()
    cc = cc or os.environ.get('CC') or shutil.which('gcc') or shutil.which('tcc') or 'cc'
    if not tcc and not shutil.which(cc) and not os.path.exists(cc):
        raise SystemExit("lbvc: no C compiler (gcc) on this computer: sudo apt install gcc")
    if not tcc:
        ensure_runtime(rt)
    with tempfile.TemporaryDirectory() as td:
        cfiles = write_units(csrc, td)
        libs = [os.path.join(rt, 'libvbrt.a'), '-lm', '-ldl']
        if uses_db(csrc):                          # DAO / ADO: SQLite and the database runtime
            dl = os.path.join(rt, 'libvbdb.a')
            dsrc = [os.path.join(rt, 'db', f) for f in os.listdir(os.path.join(rt, 'db')) if f.endswith(('.c', '.h', '.inc'))]
            if not tcc and _stale(dl, dsrc) and os.access(os.path.join(rt, 'db'), os.W_OK):
                subprocess.run(['sh', os.path.join(rt, 'db', 'build.sh')], check=True, stdout=subprocess.DEVNULL)
            libs = [dl] + libs
        if uses_forms(csrc):                       # forms: the Qt6 runtime
            fl = os.path.join(rt, 'forms', 'libvbforms.a')
            fsrc = [os.path.join(rt, 'forms', f) for f in os.listdir(os.path.join(rt, 'forms'))
                    if f.endswith(('.cpp', '.h', '.inc'))] + [os.path.join(rt, 'vbrt.h')]
            if not tcc and _stale(fl, fsrc) and os.access(os.path.join(rt, 'forms'), os.W_OK):
                subprocess.run(['sh', os.path.join(rt, 'forms', 'build.sh')], check=True, stdout=subprocess.DEVNULL)
            if tcc:                                 # tcc links the C++ of the forms as a shared library
                libs = forms_so_for_tcc(rt, out) + libs
            else:
                libs = [fl] + libs + qt_link_flags() + ['-lstdc++']
        if tcc:
            if uses_forms(csrc):                    # libvbforms.so calls libvbrt: all of it, exported
                i = libs.index(os.path.join(rt, 'libvbrt.a'))
                libs[i:i + 1] = ['-Wl,--whole-archive', libs[i], '-Wl,--no-whole-archive']
                libs.insert(0, '-rdynamic')
            cmd = tcc + ['-w', '-I', rt] + cfiles + libs + ['-o', out]
        else:
            cmd = [cc, '-O1' if optimize else '-O0', '-w', '-fPIC', '-I', rt] + cfiles + libs + ['-o', out]
            if debug_build:
                cmd[1] = '-O0'                      # quick to compile; the program waits for the IDE anyway
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0:
            sys.stderr.write(r.stderr)
            raise SystemExit("lbvc: C compiler failed (this is a bug in lbvc)")


def main(argv=None):
    ap = argparse.ArgumentParser(prog='lbvc', description='LinBASIC Studio VB6 compiler')
    ap.add_argument('inputs', nargs='+', help='a .vbp project, or .bas modules')
    ap.add_argument('-o', '--output', help='executable to write')
    ap.add_argument('-C', '--emit-c', action='store_true', help='write the generated C and stop')
    ap.add_argument('--codepage', default='cp1252', help='code page of the source files (default cp1252)')
    ap.add_argument('--check', action='store_true', help='check syntax and types only')
    ap.add_argument('--debug', action='store_true', help='with the hooks of the debugger of LinBASIC Studio')
    ap.add_argument('--target', choices=['linux', 'win32', 'win64'], default=host_target(),
                    help='the system of the executable: linux, win32 or win64 (.exe; with mingw-w64 on Linux, '
                         'the bundled tcc on Windows); default: the system lbvc runs on')
    ap.add_argument('--order', default=None, metavar='MODE[,overflow][,bounds][,flpoint][,unrounded]',
                    help="the code of VB6 to follow (order of evaluation, checks): pcode (as the IDE runs it), fast, "
                         "small or noopt (native code), with ,overflow ,bounds ,flpoint for Remove Integer Overflow / "
                         "Array Bounds / Floating Point Error Checks and ,unrounded for Allow Unrounded Floating Point "
                         "Operations; default: the Compile settings of the .vbp (pcode for .bas files)")
    ap.add_argument('--ide', action='store_true',
                    help='a run of the IDE (as Run in VB6): the order of P-code, and Debug.Print (a made program has '
                         'no Debug.Print, as in VB6)')
    a = ap.parse_args(argv)
    try:
        order = parse_order(a.order) if a.order else None
    except ValueError as e:
        ap.error(str(e))
    if os.environ.get('LBVC_DEBUG') == '1' and a.target == host_target():   # the tests: with the hooks of the debugger
        a.debug = True
    if a.target != host_target() and not (a.target in TARGETS and use_tcc(a.target)):
        a.debug = False                         # the debugger: a program that runs here
    try:
        csrc, name = compile_project(a.inputs, a.codepage, a.debug, order=order, ide=a.ide)
        for w in WARNINGS:
            sys.stderr.write(w + '\n')
    except VBError as e:
        sys.stderr.write(str(e) + '\n')
        return 1
    except OSError as e:
        sys.stderr.write(f"lbvc: {e}\n")
        return 1
    if a.check:
        return 0
    if csrc.component:                          # an ActiveX DLL: no program of its own; Make registers it
        with tempfile.TemporaryDirectory() as td:
            rt = find_runtime()
            check = (linux_tcc() + ['-c', '-o', os.path.join(td, 'x.o')] if linux_tcc() else
                     [os.environ.get('CC') or 'gcc', '-fsyntax-only']) if os.name != 'nt' else \
                [windows_tcc(rt, host_target())[0], '-B' + windows_tcc(rt, host_target())[1], '-c', '-o', os.path.join(td, 'x.o')]
            for c in write_units(csrc, td):
                r = subprocess.run(check + ['-w', '-I', rt, c],
                                   capture_output=True, text=True)
                if r.returncode != 0:
                    sys.stderr.write(r.stderr)
                    raise SystemExit("lbvc: C compiler failed (this is a bug in lbvc)")
        nm = register_component(a.inputs[0], a.codepage)
        sys.stderr.write(f"lbvc: component {nm} registered ({registry_path()}); the programs that use it "
                         f"compile it into themselves\n")
        return 0
    if a.inputs[0].lower().endswith('.vbp') and load_project(a.inputs[0], a.codepage).type.lower() in COMPONENT_TYPES:
        nm = register_component(a.inputs[0], a.codepage)    # an ActiveX EXE: a program, and a component
        sys.stderr.write(f"lbvc: component {nm} registered ({registry_path()})\n")
    out = a.output or os.path.join(os.path.dirname(os.path.abspath(a.inputs[0])), name)
    if a.emit_c:
        with open(out if out.endswith('.c') else out + '.c', 'w') as f:
            f.write(csrc)
        for i, (cn, u) in enumerate(csrc.units):
            with open((out[:-2] if out.endswith('.c') else out) + f'-{cn}.c', 'w') as f:
                f.write(u)
        return 0
    build_exe(csrc, out, find_runtime(), debug_build=a.debug, target=a.target)
    return 0
