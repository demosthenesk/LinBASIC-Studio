# LinBASIC Studio - lbvc: name/type resolution and C code generation
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""One pass over the syntax tree of all modules of a project:

  * collects the module-level symbols (constants, enums, variables, procedures),
  * resolves every name and computes the static type of every expression,
  * inserts the VB conversions between types,
  * writes C that calls the libvbrt runtime.
"""
import os
import re
from .lexer import VBError
from .parser import N
from .symbols import (NUMERIC, RANK, CT, ARRAY_KIND, SUPPORTED_TYPES, Sym, ctype, zero, cident, is_num,
                      vba_constants, date_literal, is_obj, is_udt, akind, known_type, is_arr, vt_of)
from . import builtins as B
from .evalorder import Model as OrderModel


_LC_SEPS = {1033: ('en_US', '.', ','), 2057: ('en_GB', '.', ','), 1032: ('el_GR', ',', '.'), 1031: ('de_DE', ',', '.'),
            1036: ('fr_FR', ',', ' '), 1040: ('it_IT', ',', '.'), 3082: ('es_ES', ',', '.'), 1043: ('nl_NL', ',', '.'),
            2070: ('pt_PT', ',', ' '), 1046: ('pt_BR', ',', '.'), 1049: ('ru_RU', ',', ' '), 1045: ('pl_PL', ',', ' '),
            1055: ('tr_TR', ',', '.'), 1053: ('sv_SE', ',', ' '), 1041: ('ja_JP', '.', ','), 2052: ('zh_CN', '.', ','),
            1042: ('ko_KR', '.', ',')}


def locale_numsep():
    """(decimal separator, grouping symbol) of the regional settings, as the runtime takes them (core.c lc_table:
    VBRT_LCID, else LC_ALL / LC_NUMERIC / LANG); VB6 converts a String constant with them when it compiles"""
    try:
        lcid = int(os.environ.get('VBRT_LCID', '0'))
    except ValueError:
        lcid = 0
    if lcid in _LC_SEPS:
        return _LC_SEPS[lcid][1:]
    if os.name == 'nt':                          # the regional settings of Windows (the runtime: vbos_user_lcid)
        try:
            import ctypes
            lcid = ctypes.windll.kernel32.GetUserDefaultLCID()
            if lcid in _LC_SEPS:
                return _LC_SEPS[lcid][1:]
        except Exception:                        # noqa: BLE001
            pass
    loc = os.environ.get('LC_ALL') or os.environ.get('LC_NUMERIC') or os.environ.get('LANG') or ''
    for tag, dec, grp in _LC_SEPS.values():
        if loc.startswith(tag):
            return dec, grp
    for tag, dec, grp in _LC_SEPS.values():
        if loc[:2] and loc.startswith(tag[:2]):
            return dec, grp
    return '.', ','


def cstr_lit(text):
    """a C string literal of text (UTF-8, every byte outside printable ASCII as an octal escape)"""
    out = []
    for b in str(text).encode('utf-8'):
        c = chr(b)
        if c in '\\"?':
            out.append('\\' + c)
        elif 32 <= b < 127:
            out.append(c)
        else:
            out.append('\\%03o' % b)
    return '"' + ''.join(out) + '"'

VARIANT = 'Variant'


class E:
    """a compiled expression: C code, VB type, and (for variables) the C pointer to it.
    arr: for an array inside a Type used without an index, the C pointer to its vbarr.
    prop: a public variable of an object (passed ByRef as a copy, like a property)."""
    __slots__ = ('code', 't', 'ptr', 'const', 'arr', 'prop', 'wev', 'fix', 'ext', 'aref')

    def __init__(self, code, t, ptr=None, const=None, arr=None, prop=False, wev=None, fix=None, ext=False):
        self.code, self.t, self.ptr, self.const, self.arr, self.prop = code, t, ptr, const, arr, prop
        self.ext = ext                   # a Single / Double in extended precision (vb_ld), not rounded yet
        self.fix = fix                   # String * n: n
        self.wev = wev                   # WithEvents variable: (handler table, C code of the listener)
        self.aref = None                 # an element of an array: the C pointer to its vbarr (VarPtr)


# the .ocx libraries whose controls the forms runtime has (ocx.inc); their objects are late bound
OCX_LIBS = {'mscomctllib', 'comctllib', 'mscomctl2', 'comct2lib', 'msflexgridlib', 'tabdlg', 'richtextlib',
            'picclip', 'mschart20lib', 'mscomdlg', 'mci', 'mscommlib', 'msadodclib', 'msdatagridlib',
            'msdatalistlib', 'msdatareportlib', 'stdformat', 'sysinfolib', 'mswinsocklib', 'inetctlsobjects',
            'msmask', 'msdbgrid', 'msdbctls', 'comctl3', 'msrdc', 'msdatarepeaterlib', 'msmapi',
            'mshierarchicalflexgridlib', 'shdocvw', 'shdocvwctl'}
OCX_TYPES = {'node', 'nodes', 'listitem', 'listitems', 'listsubitem', 'listsubitems', 'columnheader',
             'columnheaders', 'panel', 'panels', 'button', 'buttons', 'buttonmenu', 'buttonmenus', 'tab', 'tabs',
             'listimage', 'listimages', 'comboitem', 'comboitems', 'imagelist', 'treeview', 'listview', 'toolbar',
             'statusbar', 'progressbar', 'slider', 'tabstrip', 'imagecombo', 'updown', 'dtpicker', 'monthview',
             'flatscrollbar', 'animation', 'msflexgrid', 'sstab', 'richtextbox', 'pictureclip', 'mschart',
             'commondialog', 'adodc', 'datagrid', 'datalist', 'datacombo', 'column', 'columns', 'split', 'splits',
             'selbookmarks', 'mmcontrol', 'mscomm', 'sysinfo', 'winsock', 'inet', 'maskedbox', 'dbgrid', 'dblist',
             'dbcombo', 'coolbar', 'band', 'bands', 'msrdc', 'datarepeater', 'repeaterbindings', 'repeaterbinding',
             'mapisession', 'mapimessages', 'mshflexgrid', 'webbrowser'}

# the classes of the Data Formatting Object Library (runtime/stdfmt.c, MSSTDFMT.DLL), by their names in VB
STDFMT_TYPES = {'stddataformat': 'stddataformat', 'stddataformat2': 'stddataformat', 'istddataformatdisp': 'stddataformat',
                'dataformat': 'stddataformat', 'idataformatdisp': 'stddataformat',
                'stddataformats': 'stddataformats', 'stddataformats2': 'stddataformats',
                'istddataformatsdisp': 'stddataformats', 'dataformats': 'stddataformats', 'idataformatsdisp': 'stddataformats',
                'stddatavalue': 'stddatavalue', 'stddatavalue2': 'stddatavalue', 'istddatavaluedisp': 'stddatavalue'}
STDFMT_NEW = {'stddataformat', 'stddataformats'}

# the classes of the Scripting runtime (runtime/scripting.c), by their names in VB
SCRIPTING_TYPES = {'filesystemobject': 'fso', 'dictionary': 'dictionary', 'textstream': 'textstream',
                   'file': 'file', 'files': 'files', 'folder': 'folder', 'folders': 'folders',
                   'drive': 'drive', 'drives': 'drives',
                   # the interfaces of the type library (As Scripting.IDictionary, IFileSystem ...): the same objects
                   'idictionary': 'dictionary', 'ifilesystem': 'fso', 'ifilesystem3': 'fso', 'itextstream': 'textstream',
                   'ifile': 'file', 'ifilecollection': 'files', 'ifolder': 'folder', 'ifoldercollection': 'folders',
                   'idrive': 'drive', 'idrivecollection': 'drives'}


# the classes of DAO (runtime/db/dao.c), by their names in VB; the collections are one class
DAO_TYPES = {'dbengine': 'dbengine', 'privdbengine': 'dbengine', 'workspace': 'workspace', 'database': 'database',
             'tabledef': 'tabledef', 'field': 'field', 'index': 'index', 'relation': 'relation', 'querydef': 'querydef',
             'parameter': 'parameter', 'recordset': 'recordset', 'error': 'error', 'property': 'property'}
DAO_COLLS = {'workspaces', 'databases', 'tabledefs', 'fields', 'indexes', 'relations', 'querydefs', 'parameters',
             'recordsets', 'errors', 'properties', 'containers', 'documents', 'users', 'groups', 'connections',
             'indexfields'}
DAO_NEW = {'tabledef', 'field', 'index', 'relation', 'querydef', 'dbengine'}
# the classes of ADO (runtime/db/ado.c)
ADO_TYPES = {'connection': 'connection', 'recordset': 'recordset', 'command': 'command', 'field': 'field',
             'fields': 'fields', 'parameter': 'parameter', 'parameters': 'parameters', 'error': 'error',
             'errors': 'errors', 'properties': 'properties', 'property': 'properties',
             '_connection': 'connection', '_recordset': 'recordset', '_command': 'command', '_parameter': 'parameter'}
ADO_NEW = {'connection', 'recordset', 'command', 'parameter'}
# the classes of RDO 2.0 (runtime/db/rdo.inc); the collections are one class
RDO_TYPES = {'rdoengine': 'engine', '_rdoengine': 'engine', 'rdoenvironment': 'environment',
             '_rdoenvironment': 'environment', 'rdoconnection': 'connection', '_rdoconnection': 'connection',
             'rdoquery': 'query', 'rdopreparedstatement': 'query', 'rdoresultset': 'resultset',
             '_rdoresultset': 'resultset', 'rdocolumn': 'column', '_rdocolumn': 'column', 'rdoparameter': 'parameter',
             'rdotable': 'table', 'rdoerror': 'error', 'rdoenvironments': 'coll', 'rdoconnections': 'coll',
             'rdoqueries': 'coll', 'rdopreparedstatements': 'coll', 'rdoresultsets': 'coll', 'rdocolumns': 'coll',
             'rdoparameters': 'coll', 'rdotables': 'coll', 'rdoerrors': 'coll'}
RDO_NEW = {'connection', 'query'}
# the members of DBEngine a program uses without "DBEngine." (it is the application object of DAO)
DAO_GLOBALS = {'opendatabase', 'createdatabase', 'workspaces', 'errors', 'begintrans', 'committrans', 'rollback',
               'compactdatabase', 'repairdatabase', 'createworkspace', 'idle', 'setoption', 'registerdatabase',
               'version', 'inipath', 'systemdb', 'defaultuser', 'defaultpassword', 'logintimeout', 'defaulttype',
               'opendatabase'}


class ClassInfo:
    def __init__(self, name, mk):
        self.name, self.mk = name, mk
        self.fields = []                 # Sym (storage 'field')
        self.members = {}                # lower name -> Sym (field / proc / prop)
        self.init = self.term = None
        self.defmember = None
        self.enummember = None            # NewEnum (VB_UserMemId = -4): For Each
        self.bindable = []               # the properties marked data bound (VB_MemberFlags with "bindable"), as written
        self.events = {}                 # lower name -> N('event') with .index .psyms
        self.implements = []             # (class key, node) of Implements lines
        self.implementors = []           # keys of the classes that implement this one
        self.dispatch = {}               # C name of a member proc -> Sym of its interface dispatcher
        self.is_form = False
        self.controls = {}               # form: lower control name -> (name, is control array)
        self.ctl_cls = {}                # form: lower control name -> its class (VB.TextBox, Proj.UserCtl)
        self.form_root = None            # form: the designer tree of the .frm
        self.form_dir = self.form_enc = None
        self.is_mdi = False
        self.is_uctl = False             # a user control (.ctl): a form inside the form that has it
        self.predeclared = False         # a class with VB_PredeclaredId = True (a DataEnvironment): its default instance
        self.exposed = self.creatable = self.global_ns = False   # Instancing (VB_Exposed, VB_Creatable ...)
        self.progid = None               # an exposed class: "project.class" (one class in every unit)
        self.single_ast = False          # MultiUse = 0 (SingleUse / GlobalSingleUse of an ActiveX EXE)
        self.comp = None                 # a class of a component (ActiveX DLL/EXE) the program uses: CompInfo


def udt_free_code(t, ref):
    """C statement freeing the value at ref (a pointer expression) of type t, or None"""
    if t == 'String':
        return f"vb_sfree({ref});"
    if t == 'Variant':
        return f"vb_vfree({ref});"
    if is_obj(t):
        return f"vb_ofree({ref});"
    if is_udt(t):
        return f"U_{t[2:]}_free({ref});"
    if is_arr(t):
        return f"vb_afree({ref});"
    return None


class CompInfo:
    """a component (ActiveX DLL / EXE project) compiled into the program: its name, the C function that
    registers it, and its public classes (lower name -> ClassInfo)"""

    def __init__(self, name, reg, classes):
        self.name, self.reg, self.classes = name, reg, classes


class Compiler:
    def __init__(self, modules, project_name='Project1', startup='Sub Main', app=None, debug=False,
                 component=False, components=None):
        self.modules = modules           # [(module name, syntax tree)]
        self.component = component       # compiled as a component of another program: no main
        self.components = components or []   # the components this program uses (CompInfo)
        self.debug = debug               # lbvc --debug: the hooks of the debugger of the IDE
        self.modid = {m.lower(): i for i, (m, _) in enumerate(modules)}
        self.app = app or {}             # App.Title, App.Major ... from the .vbp
        self.uses_forms = False          # Screen, Clipboard, LoadPicture: the forms runtime
        self.ptype = 'exe'               # the type of the project (exe, oleexe, oledll, control)
        self.typelibs = set()            # the type libraries it references: 'dao', 'adodb'
        self.uses_db = False             # a Data control, CreateObject("DAO.DBEngine..."): libvbdb
        self.global_cnames = []          # the module variables and Static locals: a copy of the server swaps them
        self.static_cnames = []
        self.warnings = []               # (file, message): what compiles but will not work
        self.project_name = project_name
        self.startup = startup
        self.globals = {}                # public symbols of all modules
        self.mscopes = {}                # module name (lower) -> {name: Sym}
        self.strings = {}                # literal -> C name
        self.out = []
        self.protos = []
        self.inits = []
        self.consts = dict(vba_constants())
        self.comp_enums = set()          # the Public Enums of the components it uses (lower names)
        for comp in self.components:     # the members of the Public Enums of the components it uses
            for k, v in getattr(comp, 'consts', {}).items():
                self.consts.setdefault(k, v)
            self.comp_enums |= getattr(comp, 'enums', set())
        self.frx_cache = {}
        self.module_names = {m.lower() for m, _ in modules}
        self.udts = {}                   # type key -> N('type') with .key .mk .ftypes
        self.udt_names = {}              # (module key or '', lower name) -> type key
        self.enum_names = set()          # (module key or '', lower name) of the Enum types
        self.classes = {}                # class key -> ClassInfo
        self.class_support = []          # C code of class descriptors and late-binding thunks
        self.addressof = {}              # AddressOf: C name of the thunk -> its code
        self.keep_debug = True           # Debug.Print: in the IDE (Run) yes; in a made program (Make) no
        self.on_demand = False           # Compile On Demand (IDE runs, --debug): a procedure with a compile error
        self.on_demand_warn = False      # becomes a stub that reports it when called; warn (not in the IDE)
        self.order = ('pcode', False, False, False, False)   # the code VB6 makes (Compile tab; evalorder.py):
        #   mode (pcode/fast/noopt), Remove Integer Overflow Checks, Remove Array Bounds Checks,
        #   Remove Floating Point Error Checks, Allow Unrounded Floating Point Operations (P-code: none of them)

    # ---------------------------------------------------------------- errors
    def err(self, msg, node=None, mod=None):
        raise VBError(msg, getattr(node, 'line', 0), 1, (mod or self.cur_mod_file))

    # ---------------------------------------------------------------- types
    def deftype(self, mod, name):
        return self.mdeftypes[mod].get(name[0].upper(), VARIANT)

    def resolve_type(self, mod, name, node, new=False):
        """VB type of 'As name' in module mod"""
        if name in SUPPORTED_TYPES:
            if new:
                self.err("invalid use of New keyword", node)
            return name
        low = name.lower()
        if low.startswith(('vba.', 'vbrun.')):
            low = low.split('.', 1)[1]
        if low.startswith(self.project_name.lower() + '.'):   # ThisProject.Class: a class of the project itself
            low = low[len(self.project_name) + 1:]
        if new and low in ('stdfont', 'stdole.stdfont', 'stdpicture', 'stdole.stdpicture', 'propertybag', 'vb.propertybag'):
            self.uses_forms = True               # New StdFont / New StdPicture: objects of the forms runtime
            return 'O:' + low.split('.')[-1]
        if low in ('object', 'iunknown', 'stdole.iunknown', 'idispatch', 'stdole.idispatch'):
            if new:
                self.err("invalid use of New keyword", node)
            return 'Object'
        if low in self.classes and not self.udt_names.get((mod, low)):   # (a Private Type of the module wins)
            return 'O:' + low
        lib, _, short = low.rpartition('.')
        if lib and short in self.classes and self.classes[short].progid == low:
            return 'O:' + short                  # Lib.IFace: the copy of a component's interface in this program
        if low.startswith('vb.'):
            low = low[3:]
        if low.startswith('stdole.'):
            low = low[7:]
        from .symbols import lib_enums
        if low in lib_enums() and low not in self.classes:   # As AlignmentConstants: a Long
            if new:
                self.err("invalid use of New keyword", node)
            return 'Long'
        if low in OLE_TYPES:                    # OLE_COLOR ... (stdole): numbers
            if new:
                self.err("invalid use of New keyword", node)
            return OLE_TYPES[low]
        if low in VB_OBJECT_TYPES:              # Control, Form, TextBox ...: late bound
            if new:
                self.err("invalid use of New keyword", node)
            return 'Object'
        if low == 'collection':
            return 'O:collection'
        if (mod, low) in self.enum_names or ('', low) in self.enum_names or low in self.comp_enums or \
                (low.count('.') == 1 and (('', low.split('.')[1]) in self.enum_names
                                          or low.split('.')[1] in self.comp_enums
                                          or tuple(low.split('.')) in self.enum_names)):
            if new:
                self.err("invalid use of New keyword", node)
            return 'Long'
        key = self.udt_names.get((mod, low)) or self.udt_names.get(('', low))
        if not key and low.count('.') == 1:      # Module1.MEMORYSTATUS: a Type of that module
            m1, t1 = low.split('.')
            key = self.udt_names.get((m1, t1)) or (m1 in self.module_names and self.udt_names.get(('', t1)))
        if key:
            if new:
                self.err("invalid use of New keyword", node)
            return 'U:' + key
        lib, _, short = low.rpartition('.')
        if short in STDFMT_TYPES and (lib == 'stdformat' or (not lib and 'stdformat' in self.typelibs)):
            t = STDFMT_TYPES[short]                # StdDataFormat ... (runtime/stdfmt.c)
            if new and t not in STDFMT_NEW:
                self.err("invalid use of New keyword", node)
            return 'O:stdformat.' + t
        if lib in OCX_LIBS or (not lib and short in OCX_TYPES):
            if new:                              # the objects of the .ocx controls: late bound
                self.err("invalid use of New keyword", node)
            return 'Object'
        if self.ext_class(low):                  # a class of a component the program uses: late bound
            return 'Object'
        t = SCRIPTING_TYPES.get(low[10:] if low.startswith('scripting.') else low)
        if t:                                    # the Scripting runtime (scripting.c)
            if new and t not in ('fso', 'dictionary'):
                self.err("invalid use of New keyword", node)
            return 'O:scripting.' + t
        order = getattr(self.typelibs, 'order', ()) or tuple(sorted(self.typelibs))
        ado_first = 'adodb' in self.typelibs and (lib == 'adodb' or (not lib and (
            'dao' not in self.typelibs or order.index('adodb') < order.index('dao'))))
        if ado_first and short in ADO_TYPES:     # ADO (runtime/db/ado.c)
            t = ADO_TYPES[short]
            if new and t not in ADO_NEW:
                self.err("invalid use of New keyword", node)
            return 'O:adodb.' + t
        if 'rdo' in self.typelibs and (not lib or lib == 'rdo') and short in RDO_TYPES:   # RDO (runtime/db/rdo.inc)
            t = RDO_TYPES[short]
            if new and t not in RDO_NEW:
                self.err("invalid use of New keyword", node)
            return 'O:rdo.' + t
        if 'msbind' in self.typelibs and (not lib or lib == 'msbind') and short in ('bindingcollection', 'binding'):
            if new and short != 'bindingcollection':
                self.err("invalid use of New keyword", node)
            return 'O:msbind.' + short
        if 'dao' in self.typelibs and (not lib or lib == 'dao'):   # DAO (runtime/db)
            if short in DAO_TYPES:
                if new and short not in DAO_NEW:
                    self.err("invalid use of New keyword", node)
                return 'O:dao.' + DAO_TYPES[short]
            if short in DAO_COLLS:
                if new:
                    self.err("invalid use of New keyword", node)
                return 'O:dao.coll'
        self.err(f"user-defined type not defined: {name}", node)

    def ext_of(self, t, tnode):
        """As Lib.Class (a class of a component, late bound in C): its ClassInfo, for the members VB6
        checks when it compiles (early binding)"""
        if t != 'Object' or tnode is None:
            return None
        e = self.ext_class(tnode.name.lower())
        return e[1] if e else None

    def autonew_of(self, t, tnode):
        """the class of Dim x As New C: a type, or 'P:proj.class' for a class of a component"""
        ext = self.ext_class(tnode.name.lower()) if t == 'Object' and tnode is not None else None
        return f"P:{ext[0].name}.{ext[1].name}" if ext else t

    def ext_class(self, low):
        """'proj.class' or 'class' of a component the program uses -> (CompInfo, ClassInfo), or None"""
        proj, _, cls = low.rpartition('.')
        for ci in self.components:
            if proj and proj != ci.name.lower():
                continue
            if cls in ci.classes and (proj or cls not in self.classes):
                return ci, ci.classes[cls]
        return None

    def type_of_decl(self, mod, name, suffix, tnode, node, allow_any=False):
        if tnode is not None:
            if suffix:
                self.err("a variable with a type suffix cannot also have As <type>", node)
            if tnode.name == 'Any':
                if not allow_any:
                    self.err("As Any is only valid in a Declare statement", node)
                return 'Any'
            return self.resolve_type(mod, tnode.name, node, tnode.new)
        if suffix:
            from .lexer import SUFFIX_TYPE
            return SUFFIX_TYPE[suffix]
        return self.deftype(mod, name)

    def fixlen_of(self, mod, tnode, node, local=None):
        """n of 'As String * n', or None"""
        if tnode is None or getattr(tnode, 'fixlen', None) is None:
            return None
        v, t = self.eval_const(tnode.fixlen, mod, local)
        n = int(self.coerce_const(v, t, 'Long', node))
        if not 1 <= n <= 65526:
            self.err("invalid length for a fixed-length string", node)
        return n

    def udt_len(self, t):
        """Len() of a Type: the bytes Put writes for it (variable-length strings count 4)"""
        total = 0
        for f in self.udts[t[2:]].fields:
            if f.dims is not None and not f.fixed_dims:
                total += 4
                continue
            n = 1
            for lo, hi in (f.fixed_dims or []):
                n *= hi - lo + 1
            if f.fixlen:
                one = f.fixlen
            elif is_udt(f.t):
                one = self.udt_len(f.t)
            else:
                one = {'Integer': 2, 'Boolean': 2, 'Long': 4, 'Single': 4, 'Double': 8, 'Currency': 8, 'Date': 8,
                       'Byte': 1, 'String': 4, 'Variant': 16}.get(f.t, 4)
            total += one * n
        return total

    def udt_lenb(self, t, align=False):
        """LenB() of a Type: its size in the memory of VB6 - fixed strings in Unicode, fields at multiples
        of their size up to 4, the whole a multiple of its largest alignment"""
        off, al = 0, 1
        for f in self.udts[t[2:]].fields:
            n = 1
            for lo, hi in (f.fixed_dims or []):
                n *= hi - lo + 1
            if f.dims is not None and not f.fixed_dims:
                one, a = 4, 4                   # a dynamic array: a pointer to its SAFEARRAY
            elif f.fixlen:
                one, a = f.fixlen * 2, 2
            elif is_udt(f.t):
                one, a = self.udt_lenb(f.t, True)
            else:
                one = {'Integer': 2, 'Boolean': 2, 'Long': 4, 'Single': 4, 'Double': 8, 'Currency': 8, 'Date': 8,
                       'Byte': 1, 'String': 4, 'Variant': 16}.get(f.t, 4)
                a = min(one, 4)
            al = max(al, a)
            off = (off + a - 1) // a * a + one * n
        size = (off + al - 1) // al * al
        return (size, al) if align else size

    @staticmethod
    def vbk(t, fix=None):
        """runtime VBK_ kind of a type (Get/Put)"""
        if fix:
            return 'VBK_FSTR'
        if is_udt(t):
            return 'VBK_UDT'
        if is_obj(t):
            return 'VBK_OBJ'
        return {'Integer': 'VBK_I2', 'Long': 'VBK_I4', 'Single': 'VBK_R4', 'Double': 'VBK_R8', 'Currency': 'VBK_CY',
                'Date': 'VBK_DATE', 'Boolean': 'VBK_BOOL', 'Byte': 'VBK_U8', 'String': 'VBK_STR',
                'Variant': 'VBK_VAR'}[t]

    # ---------------------------------------------------------------- literals
    def non_scripting_consts(self):
        """the constants of the libraries other than Scripting (a name both have stays)"""
        from .symbols import lib_const_names
        out = set()
        for lib in ('VBA', 'VBRUN', 'stdole', 'MSComDlg'):
            out |= lib_const_names(lib)
        return out

    def app_object(self):
        """App as a value: the object of the runtime, told what the .vbp says when the program starts"""
        if not getattr(self, '_app_obj', False):
            self._app_obj = True
            info = dict(self.app)
            if info.get('title') is None and self.project_name:
                info['title'] = self.project_name
            for k in ('title', 'helpfile', 'major', 'minor', 'revision', 'comments', 'companyname', 'filedescription',
                      'legalcopyright', 'legaltrademarks', 'productname'):
                v = info.get(k)
                if v is not None:
                    self.inits.append(f'vb_app_info("{k}", {self.strlit(str(v))});')
        return 'vb_app_obj()'

    def strlit(self, s):
        if s not in self.strings:
            self.strings[s] = f"S{len(self.strings)}"
        name = self.strings[s]
        return f"vb_slit({name}, {len(s)})"

    def const_code(self, v, t):
        if t == 'String':
            return self.strlit(v)
        if t == 'Boolean':
            return '((vb_bool)-1)' if v else '((vb_bool)0)'
        if t == 'Currency':
            n = int(round(v * 10000))
            return f"((vb_cy){n}LL)" if n != -2**63 else "((vb_cy)(-9223372036854775807LL-1))"
        if t in ('Double', 'Date'):
            r = repr(float(v))
            return f"({r})" if 'e' in r or '.' in r or 'inf' in r else f"({r}.0)"
        if t == 'Single':
            return f"({float(v)!r}f)"
        if t == 'Integer':
            return f"((vb_i16){int(v)})"
        if t == 'Byte':
            return f"((vb_u8){int(v)})"
        if t == 'Long':
            return f"((vb_i32){int(v)})" if int(v) != -2147483648 else "((vb_i32)(-2147483647-1))"
        if t == 'Variant':
            return 'vb_v_empty()' if v is None else 'vb_v_obj((vbobj)0)' if v == 'Nothing' else 'vb_v_null()'
        if is_obj(t):                           # an Optional object parameter: Nothing
            return '((vbobj)0)'
        raise AssertionError(t)

    INTRINSIC = {'commandbutton', 'label', 'textbox', 'checkbox', 'optionbutton', 'listbox', 'combobox', 'timer',
                 'frame', 'picturebox', 'image', 'shape', 'line', 'hscrollbar', 'vscrollbar', 'drivelistbox',
                 'dirlistbox', 'filelistbox', 'menu'}

    SUPPORTED_OCX = {'mscomdlg.commondialog', 'picclip.pictureclip'} | {
        f'{lib}.{t}' for lib in ('mscomctllib', 'comctllib')
        for t in ('imagelist', 'progressbar', 'slider', 'statusbar', 'toolbar', 'tabstrip', 'treeview', 'listview',
                  'imagecombo')} | {f'{lib}.{t}' for lib in ('mscomctl2', 'comct2lib')
                                     for t in ('updown', 'dtpicker', 'monthview', 'flatscrollbar')} | {
        'tabdlg.sstab', 'msflexgridlib.msflexgrid', 'mshierarchicalflexgridlib.mshflexgrid',
        'richtextlib.richtextbox', 'mschart20lib.mschart', 'mscomctl2.animation', 'comct2lib.animation',
        'msadodclib.adodc', 'msdatagridlib.datagrid', 'msdatalistlib.datalist', 'msdatalistlib.datacombo',
        'mci.mmcontrol', 'mscommlib.mscomm', 'sysinfolib.sysinfo', 'mswinsocklib.winsock', 'inetctlsobjects.inet',
        'msmask.maskedbox', 'msdbgrid.dbgrid', 'msdbctls.dblist', 'msdbctls.dbcombo', 'comctl3.coolbar',
        'msrdc.msrdc', 'msdatarepeaterlib.datarepeater', 'msmapi.mapisession', 'msmapi.mapimessages',
        'shdocvwctl.webbrowser', 'shdocvw.webbrowser'}
    DB_OCX = {'msadodclib.adodc', 'msdatagridlib.datagrid', 'msdatalistlib.datalist', 'msdatalistlib.datacombo',
              'msdbgrid.dbgrid', 'msdbctls.dblist', 'msdbctls.dbcombo', 'msrdc.msrdc', 'msdatarepeaterlib.datarepeater'}

    def check_control(self, ast, c):
        """a warning for a control that the forms runtime does not have (it stays an empty box)"""
        lib, _, cls = c.cls.rpartition('.')
        if c.cls.lower() == 'vb.data' or c.cls.lower() in self.DB_OCX:   # the database runtime (libvbdb)
            self.uses_db = True
            if c.cls.lower() == 'vb.data':
                return
        if lib.lower() == 'vb' and cls.lower() in self.INTRINSIC:
            return
        if c.cls.lower() in self.SUPPORTED_OCX or c.cls.lower() in self.uctl_names:
            return
        if (ast.file, c.cls.lower()) in {(f, k) for f, k, _ in self.warnings}:
            return
        what = {'vb.data': "a Data control (databases: DAO/Jet)", 'vb.ole': "an OLE container"}.get(
            c.cls.lower(), f"an ActiveX control ({c.cls})")
        self.warnings.append((ast.file, c.cls.lower(),
                              f"{c.name} is {what}, which LinBASIC Studio does not support yet: it stays an empty box"))

    # ---------------------------------------------------------------- pass 1: module symbols
    def collect_types(self):
        """pass 0: names of classes and user-defined types, then the fields of the types"""
        # the user controls a form can have: of this project, and of the control projects it uses
        self.uctl_names = {f"{self.project_name}.{m}".lower() for m, a in self.modules if getattr(a, 'is_uctl', False)}
        for comp in self.components:
            self.uctl_names |= {f"{comp.name}.{c.name}".lower() for c in comp.classes.values() if c.is_uctl}
        for mname, ast in self.modules:
            if getattr(ast, 'is_class', False):
                ci = self.classes[mname.lower()] = ClassInfo(mname, mname.lower())
                va = getattr(ast, 'vb_attrs', {})
                if not getattr(ast, 'is_form', False) or getattr(ast, 'is_uctl', False):
                    ci.exposed = va.get('vb_exposed') == 'true'
                    ci.creatable = ci.exposed and va.get('vb_creatable') == 'true'
                    ci.global_ns = ci.exposed and va.get('vb_globalnamespace') == 'true'
                    ci.single_ast = ci.creatable and not getattr(ast, 'multiuse', True)
                    if ci.exposed:            # its identity across the units of a program (Implements)
                        owner = getattr(ast, 'vb_owner', None) or self.project_name
                        ci.progid = f"{owner}.{mname}".lower()
                if not getattr(ast, 'is_form', False) and va.get('vb_predeclaredid') == 'true':
                    ci.predeclared = True
                if getattr(ast, 'is_form', False):
                    ci.is_form = True
                    ci.is_mdi = ast.form_root.type_name == 'MDIForm'
                    ci.is_uctl = ast.form_root.type_name == 'UserControl'
                    ci.form_root = ast.form_root
                    ci.form_dir = getattr(ast, 'form_dir', None)
                    ci.form_enc = getattr(ast, 'form_enc', None)
                    for c in ast.form_root.walk():
                        if c is ast.form_root:
                            continue
                        self.check_control(ast, c)
                        key = c.name.lower()
                        arr = c.index is not None
                        if key in ci.controls and ci.controls[key][1] != arr:
                            self.cur_mod_file = ast.file
                            self.err(f"control {c.name}: an array element and a single control with the same name")
                        ci.controls[key] = (c.name, arr)
                        ci.ctl_cls[key] = c.cls
        for mname, ast in self.modules:          # the Enum types: As Colors is a Long
            for it in ast.items:
                if it.kind == 'enum':
                    self.enum_names.add(('' if it.scope in ('public', 'global') else mname.lower(), it.name.lower()))
        imported = self.import_comp_types()
        for mname, ast in self.modules:
            self.cur_mod_file = ast.file
            mk = mname.lower()
            for it in ast.items:
                if it.kind == 'type':
                    # a Public Type in a class: VB6 allows it in the public classes of ActiveX projects; here it
                    # is visible to the whole program as in a standard module. It can go in a Variant (VT_RECORD,
                    # its name "Project.Type" identifies it in the components too)
                    it.public_obj = it.scope in ('public', 'global') and getattr(ast, 'is_class', False)
                    it.rec_name = f"{self.project_name}.{it.name}"
                    it.imported = False
                    key = cident(mk) + '_' + cident(it.name)
                    scope = '' if it.scope in ('public', 'global') else mk
                    if (scope, it.name.lower()) in self.udt_names:
                        if scope:
                            self.err(f"ambiguous name detected: {it.name}", it)
                        # the same Public Type in two modules (POINTAPI ...): the first one is used
                    else:
                        self.udt_names[(scope, it.name.lower())] = key
                    it.key, it.mk = key, mk
                    self.udts[key] = it
        for low, hk in imported:                 # As ChartItem: a Type of a component, unless the program has one
            self.udt_names.setdefault(('', low), hk)
        for key, it in self.udts.items():
            if getattr(it, 'imported', False):  # (resolved in its component)
                continue
            self.cur_mod_file = dict((m.lower(), a.file) for m, a in self.modules)[it.mk]
            seen = set()
            for f in it.fields:
                if f.name.lower() in seen:
                    self.err(f"duplicate declaration in current scope: {f.name}", f)
                seen.add(f.name.lower())
                f.t = self.type_of_decl(it.mk, f.name, f.suffix, f.type, f) if (f.type or f.suffix) else 'Variant'
                f.cname = 'f_' + cident(f.name)
                f.fixed_dims, f.fixlen = None, None         # resolve_udt_sizes, once the constants are known
                f.autonew = self.autonew_of(f.t, f.type) if (f.type is not None and f.type.new) else None
                if getattr(it, 'public_obj', False) and f.type is not None and (f.type.new or f.type.fixlen is not None):
                    self.err("Fixed-length strings and use of the 'new' qualifier are not allowed for fields in a "
                             "public user defined type defined in an object module", f)
        # order: a Type used inside another comes first; a Type cannot contain itself
        order, state = [], {}

        def visit(k, node):
            if state.get(k) == 1:
                self.err("a Type cannot contain itself", node)
            if state.get(k) == 2:
                return
            state[k] = 1
            for f in self.udts[k].fields:
                if is_udt(f.t) and f.dims is None:
                    visit(f.t[2:], f)
            state[k] = 2
            order.append(k)
        for k, it in self.udts.items():
            visit(k, it)
        self.udt_order = order

    def import_comp_types(self):
        """the Public Types of the components this program uses (their CompInfo.types, resolved there): a copy in
        this program (U_x<component>_<key>, the same fields, so the same layout), named "Component.Type"; gives the
        bare names to register once the program's own Types are known"""
        import copy
        bare = []
        for comp in self.components:
            types = getattr(comp, 'types', None) or {}
            kmap = {ck: 'x' + cident(comp.name.lower()) + '_' + ck for ck in types}
            ok = {}
            for ck, it in types.items():         # (a field of a Type the component does not give: left out)
                ok[ck] = all(not is_udt(f.t) or f.t[2:] in kmap for f in it.fields)
            for ck, it in types.items():
                hk = kmap[ck]
                if not ok[ck] or hk in self.udts:
                    continue
                n = copy.copy(it)
                n.fields = []
                for f in it.fields:
                    g = copy.copy(f)
                    if is_udt(f.t):
                        g.t = 'U:' + kmap[f.t[2:]]
                    elif is_obj(f.t) and f.t not in ('Object', 'O:collection', 'O:stdfont', 'O:stdpicture'):
                        g.t, g.autonew = 'Object', None       # a class of the component: late bound here
                    n.fields.append(g)
                n.key, n.mk, n.imported, n.public_obj = hk, None, True, True
                n.rec_name = getattr(it, 'rec_name', None) or f"{comp.name}.{it.name}"
                self.udts[hk] = n
                self.udt_names.setdefault(('', f"{comp.name.lower()}.{it.name.lower()}"), hk)
                bare.append((it.name.lower(), hk))
        return bare

    def rec_ok(self, t):
        """a Type that can go in a Variant: a Public Type of a public object module (a class, a user control), its
        own or a component's"""
        it = self.udts.get(t[2:]) if is_udt(t) else None
        return bool(it is not None and getattr(it, 'public_obj', False))

    # ---------------------------------------------------------------- Implements, events
    @staticmethod
    def same_signature(a, b):
        if len(a.params) != len(b.params) or (a.rt or None) != (b.rt or None):
            return False
        return all(x.t == y.t and bool(x.byval) == bool(y.byval) and bool(x.array) == bool(y.array) and
                   bool(x.optional) == bool(y.optional) for x, y in zip(a.params, b.params))

    def implements(self, a, b):
        """class key a implements class key b"""
        return b in (k for k, _ in self.classes[a].implements) if a in self.classes else False

    @staticmethod
    def cparams(sym):
        """C parameter declarations and argument names of a procedure (without self)"""
        decls, names = [], []
        for i, p in enumerate(sym.params):
            n = f"p{i}"
            if p.array:
                decls.append(f"vbarr *{n}")
            elif p.byval:
                decls.append(f"{ctype(p.t)} {n}")
            else:
                decls.append(f"{ctype(p.t)} *{n}")
            names.append(n)
        return decls, names

    def link_classes(self):
        files = dict((m.lower(), a.file) for m, a in self.modules)
        sup = self.class_support
        for key, ci in self.classes.items():
            self.cur_mod_file = files[key]
            # Implements: every public member of the interface needs <Interface>_<member>
            for ik, node in ci.implements:
                iface = self.classes[ik]
                iface.implementors.append(key)
                for low, m in iface.members.items():
                    if m.kind == 'var':
                        if m.public:
                            self.err(f"Implements {iface.name}: public variables in an interface are not "
                                     f"supported yet ({m.name})", node)
                        continue
                    impl = ci.members.get(f"{iface.name.lower()}_{low}")
                    slots = [('proc', m)] if m.kind == 'proc' else \
                        [(sl, getattr(m, sl)) for sl in ('get', 'let', 'set') if getattr(m, sl) is not None]
                    for sl, p in slots:
                        if not p.public:
                            continue
                        ip = impl if sl == 'proc' else (getattr(impl, sl, None) if impl is not None and impl.kind == 'prop' else None)
                        if ip is None or ip.kind != 'proc' or ip.pkind != p.pkind:
                            self.err(f"object module needs to implement '{p.name}' for interface '{iface.name}'",
                                     node)
                        if not self.same_signature(p, ip):
                            self.err(f"procedure declaration does not match description of event or procedure "
                                     f"having the same name: {iface.name}_{p.name}", ip.node)
                        p.impls = getattr(p, 'impls', []) + [(key, ip)]
            # WithEvents variables: handler tables
            for f in ci.fields:
                if not getattr(f, 'withevents', False):
                    continue
                if getattr(f, 'wev_ctl', False):
                    pre = f.name.lower() + '_'
                    f.wev_handlers = [(low[len(pre):], m) for low, m in ci.members.items()
                                      if low.startswith(pre) and m.kind == 'proc' and m.pkind == 'sub']
                    f.htable = f"HT_{cident(key)}_{f.cname}"
                    self.protos.append(f"static const vb_member {f.htable}_m[{len(f.wev_handlers) + 1}];")
                    f.htable = f"((const vb_fn *){f.htable}_m)"
                    continue
                ext = getattr(f, 'wev_ext', None)
                src = ext or self.classes[f.t[2:]]
                if not src.events:
                    self.err(f"object does not source automation events: {src.name}", f)
                hs = []
                for ev in sorted(src.events.values(), key=lambda e: e.index):
                    h = ci.members.get(f"{f.name.lower()}_{ev.name.lower()}")
                    if h is not None:
                        same = (len(ev.sym.params) == len(h.params) if ext else self.same_signature(ev.sym, h))
                        if h.kind != 'proc' or h.pkind != 'sub' or not same:
                            self.err(f"procedure declaration does not match description of event or procedure "
                                     f"having the same name: {f.name}_{ev.name}", h.node if h.kind == 'proc' else None)
                        hs.append(f"(vb_fn){h.cname}")
                    else:
                        hs.append("0")
                f.htable = f"HT_{cident(key)}_{f.cname}"
                self.protos.append(f"static const vb_fn {f.htable}[{len(hs)}];")
                sup.append(f"static const vb_fn {f.htable}[{len(hs)}] = {{ {', '.join(hs)} }};")
            # RaiseEvent: one function per event, calling the handlers of every listener
            for ev in ci.events.values():
                decls, names = self.cparams(ev.sym)
                proto = f"static void {ev.sym.cname}({', '.join(['vbobj vb_self'] + decls)})"
                self.protos.append(proto + ';')
                ftype = f"void (*)({', '.join(['vbobj'] + decls)})"
                call = f"(({ftype})h)({', '.join(['ls[i]'] + names)});"
                host = ''

                def by_name(fire):
                    """the arguments as Variants (ByRef ones back after), to handlers found by the event's name"""
                    pg = ThunkGen(self, ci).pg
                    vals, back = [], []
                    for i, (p, nm) in enumerate(zip(ev.sym.params, names)):
                        if p.array:
                            vals.append('vb_v_empty()')
                            continue
                        x = E(nm, p.t) if p.byval else E(f"(*{nm})", p.t, ptr=nm)
                        vals.append(f"vb_vset(&av[{i}], {pg.conv(x, 'Variant')});")
                        if not p.byval:
                            back.append(f"if (!vb_errpend) {pg.store_code(x, E(f'av[{i}]', 'Variant'), None)}")
                    k = len(vals)
                    return (f" {{ vbvar av[{max(k, 1)}] = {{ {{ 0 }} }}; " + ' '.join(v for v in vals if v.startswith('vb_vset'))
                            + f' {fire}(vb_self, "{c_utf8(ev.name.lower())}", {k}, av); '
                            + ' '.join(back) + f" for (i = 0; i < {k}; i++) vb_vfree(&av[i]); }}")
                if ci.is_form and not ci.is_uctl:   # WithEvents f As Form1: the listeners have tables of handlers by name
                    sup.append(proto + " { vb_i32 i;" + by_name('vb_fire_named') + " }")
                    continue
                if ci.is_uctl:                  # a user control: the event goes to the form that has it too
                    host = by_name('vbf_uctl_event')
                sup.append(proto + " { vbobj *ls; const vb_fn **ts; vb_i32 i, n = vb_sinks_begin(vb_self, &ls, &ts); "
                           f"for (i = 0; i < n && !vb_errpend; i++) {{ vb_fn h = ts[i][{ev.index}]; if (h) {call} }} "
                           "vb_sinks_end(ls, ts, n);" + host + " }")
        # interface dispatchers: a call through a variable of the interface type goes to the
        # implementation of the object's class
        for key, ci in self.classes.items():
            # a public class of a component may be implemented by a class of the program that uses it
            if not ci.implementors and not (self.component and ci.exposed):
                continue
            procs = []
            for m in ci.members.values():
                if m.kind == 'proc':
                    procs.append(m)
                elif m.kind == 'prop':
                    procs += [getattr(m, sl) for sl in ('get', 'let', 'set') if getattr(m, sl) is not None]
            pg = None
            for p in procs:
                if not p.public or (not getattr(p, 'impls', None) and not (self.component and ci.exposed)
                                    and not ci.implementors):
                    continue
                d = Sym('proc', p.name, pkind=p.pkind, rt=p.rt, params=p.params, node=p.node, module=p.module,
                        cname='IF_' + p.cname, cls=key, public=True)
                ci.dispatch[p.cname] = d
                decls, names = self.cparams(p)
                rt = ctype(p.rt) if p.rt else 'void'
                proto = f"static {rt} {d.cname}({', '.join(['vbobj vb_self'] + decls)})"
                self.protos.append(proto + ';')
                args = ', '.join(['o'] + names)
                body = ["vbobj o = vb_oid(vb_self);"]
                for ik, ip in getattr(p, 'impls', []):
                    call = f"{ip.cname}({args})"
                    body.append(f"if (o && o->cls == &CLS_{cident(ik)}) "
                                + (f"return {call};" if p.rt else f"{{ {call}; return; }}"))
                # an object of another program or component: its implementation by name ("iface.member")
                if pg is None:
                    pg = ThunkGen(self, ci).pg
                late = self.late_impl(pg, ci, p, names)
                if late:
                    body.append(f"if (o && o->cls != &CLS_{cident(key)}) {{ {late} }}")
                body.append(f"return {p.cname}({args});" if p.rt else f"{p.cname}({args});")
                sup.append(proto + " { " + ' '.join(body) + " }")

    @staticmethod
    def ao_body(call, s, post=''):
        """the body of an AddressOf thunk: the call, its temporaries released, the result as a Long"""
        if s.rt in ('Long', 'Integer', 'Byte', 'Boolean'):
            return f"vb_i32 r = (vb_i32){call}; {post} vb_release(m); return r;"
        if s.rt:
            return f"(void){call}; {post} vb_release(m); return 0;"
        return f"{call}; {post} vb_release(m); return 0;"

    def late_impl(self, pg, ci, p, names):
        """C code calling member p of interface ci late bound, as "<interface>.<member>", or None"""
        if any(q.array for q in p.params) or is_udt(p.rt or '') or (p.rt and is_arr(p.rt)):
            return None
        vals = []
        for q, n in zip(p.params, names):
            if is_udt(q.t):
                return None
            vals.append(pg.conv(E(n if q.byval else f"(*{n})", q.t), 'Variant'))
        kind = {'property get': 'VBM_GET', 'property let': 'VBM_LET', 'property set': 'VBM_SET'}.get(p.pkind, 'VBM_METHOD')
        argv = f"{len(vals)}, (vbvar[]){{ {', '.join(vals)} }}" if vals else "0, 0"
        call = f'vb_late_o(o, "{c_utf8(ci.name.lower())}.{c_utf8(p.name.lower())}", {kind}, {argv})'
        if p.rt:
            return f"return {pg.conv(E(call, 'Variant'), p.rt)};"
        return f"(void){call}; return;"

    def res_code(self):
        """the resource file of the project as C data (LoadResString, LoadResData, LoadResPicture)"""
        res = self.app.get('res') or []
        if not res:
            return []
        out, rows = [], []
        for i, (t, n, data) in enumerate(res):
            out.append(f"static const unsigned char VB_RES_{i}[] = {{{','.join(str(x) for x in data) or '0'}}};")
            tid, tname = (t, '0') if isinstance(t, int) else (0, f'"{c_utf8(t)}"')
            nid, nname = (n, '0') if isinstance(n, int) else (0, f'"{c_utf8(n)}"')
            rows.append(f"{{{tid}, {tname}, {nid}, {nname}, VB_RES_{i}, {len(data)}}}")
        out.append(f"static const vb_res VB_RES[] = {{ {', '.join(rows)} }};")
        return out + ['']

    def interface_vars(self):
        """Public x As T in a class another class Implements: VB6 makes it a property of the interface
        (the implementing class has Property Get / Let / Set Iface_x); here it becomes such a property,
        kept in a private variable of its own"""
        from .parser import parse_module
        impl = {it.name.lower().rpartition('.')[2] for _, ast in self.modules for it in ast.items
                if it.kind == 'implements'}
        for mname, ast in self.modules:
            mk = mname.lower()
            if mk not in impl or mk not in self.classes:
                continue
            items, extra = [], []
            for it in ast.items:
                if it.kind != 'dim' or it.scope not in ('public', 'global'):
                    items.append(it)
                    continue
                keep = []
                for v in it.vars:
                    if v.dims is not None or getattr(v, 'withevents', False) or (v.type is not None and (
                            v.type.new or v.type.fixlen)):
                        keep.append(v)
                        continue
                    t = self.type_of_decl(mk, v.name, v.suffix, v.type, v)
                    tn = v.type.name if v.type is not None else {None: 'Variant', '%': 'Integer', '&': 'Long',
                                                                 '!': 'Single', '#': 'Double', '@': 'Currency',
                                                                 '$': 'String'}.get(v.suffix, 'Variant')
                    n, f = v.name, f"lbIf_{v.name}"
                    src = [f"Private {f} As {tn}"]
                    if t == 'Variant':
                        src += [f"Public Property Get {n}() As {tn}",
                                f"If IsObject({f}) Then Set {n} = {f} Else {n} = {f}", "End Property"]
                    elif is_obj(t) or t == 'Object':
                        src += [f"Public Property Get {n}() As {tn}", f"Set {n} = {f}", "End Property"]
                    else:
                        src += [f"Public Property Get {n}() As {tn}", f"{n} = {f}", "End Property"]
                    # the signatures VB6 wants of the implementing class (compile-6.8 c17-c25): a Variant has
                    # Let and Set ByRef, an object Set ByVal, the others Let ByVal
                    bv = '' if t == 'Variant' else 'ByVal '
                    if not (is_obj(t) or t == 'Object'):
                        src += [f"Public Property Let {n}({bv}lbV As {tn})", f"{f} = lbV", "End Property"]
                    if is_obj(t) or t in ('Object', 'Variant'):
                        src += [f"Public Property Set {n}({bv}lbV As {tn})", f"Set {f} = lbV", "End Property"]
                    gen = parse_module('\r\n'.join(src) + '\r\n', ast.file, mname).items
                    for g in gen:
                        g.line = v.line
                    items.append(gen[0])
                    extra += gen[1:]
                if keep:
                    it.vars = keep
                    items.append(it)
            ast.items = items + extra

    def collect_consts(self):
        """the module-level constants of all the modules first, in the order their values become
        known: Const A = B where B is a Global Const of a module that comes later in the project"""
        done = set()
        progress = True
        while progress:
            progress = False
            for mname, ast in self.modules:
                self.cur_mod_file = ast.file
                mk = mname.lower()
                for it in ast.items:
                    if it.kind == 'enum' and id(it) not in done:
                        # Enum members, with the constants: Const A = EnumMember, Optional x As E = M before the Enum
                        vals, nxt, local = [], 0, {}
                        try:
                            for m in it.members:     # (a member can use the ones before it: C = A Or B)
                                if m.value is not None:
                                    v, t = self.eval_const(m.value, mk, local)
                                    nxt = int(self.coerce_const(v, t, 'Long', m))
                                vals.append(nxt)
                                local[m.name.lower()] = Sym('const', m.name, value=nxt, t='Long')
                                nxt += 1
                        except VBError:
                            continue
                        for m, v in zip(it.members, vals):
                            self.declare(mk, it.scope, Sym('const', m.name, value=v, t='Long', enum_member=True,
                                                           public=it.scope != 'private',
                                                           decl_line=getattr(m, 'line', 0)), m)
                        self.declare(mk, it.scope, Sym('enum', it.name, public=it.scope != 'private'), it)
                        done.add(id(it))
                        progress = True
                        continue
                    if it.kind != 'const':
                        continue
                    for c in it.items:
                        if id(c) in done:
                            continue
                        try:
                            v, t = self.eval_const(c.value, mk)
                        except VBError:
                            continue
                        tt = self.type_of_decl(mk, c.name, c.suffix, c.type, c) if (c.type or c.suffix) else t
                        v = self.coerce_const(v, t, tt, c)
                        self.declare(mk, it.scope, Sym('const', c.name, value=v, t=tt, decl_line=getattr(c, 'line', 0)), c)
                        done.add(id(c))
                        progress = True
        return done

    def resolve_udt_sizes(self):
        """the bounds of the fixed arrays and the lengths of the fixed strings of the Type fields:
        a(MAXN) As Long, s As String * LEN_NAME, with constants of any module"""
        files = dict((m.lower(), a.file) for m, a in self.modules)
        for key, it in self.udts.items():
            if getattr(it, 'imported', False):
                continue
            self.cur_mod_file = files[it.mk]
            for f in it.fields:
                f.fixed_dims = [self.dim_bounds(d, it.mk, f) for d in f.dims] if f.dims else None
                f.fixlen = self.fixlen_of(it.mk, f.type, f)   # (a(9) As String * 8: an array of them too)

    def collect(self):
        self.mdeftypes = {}
        for mname, ast in self.modules:
            self.mscopes.setdefault(mname.lower(), {})
        self.collect_types()
        for mname, ast in self.modules:
            mk = mname.lower()
            self.mscopes.setdefault(mk, {})
            self.mdeftypes[mk] = {}
            for it in ast.items:
                if it.kind == 'deftype':
                    for a, b in it.ranges:
                        for c in range(ord(a), ord(b) + 1):
                            self.mdeftypes[mk][chr(c)] = it.type
        self.interface_vars()
        done = self.collect_consts()
        self.resolve_udt_sizes()
        for mname, ast in self.modules:
            self.cur_mod_file = ast.file
            mk = mname.lower()
            scope = self.mscopes.setdefault(mk, {})
            for it in ast.items:
                if it.kind in ('event', 'implements') and mk not in self.classes:
                    self.err(f"{'Event' if it.kind == 'event' else 'Implements'} is only valid in a class module", it)
            for it in ast.items:
                if it.kind == 'const':
                    for c in it.items:
                        if id(c) in done:
                            continue
                        v, t = self.eval_const(c.value, mk)
                        tt = self.type_of_decl(mk, c.name, c.suffix, c.type, c) if (c.type or c.suffix) else t
                        v = self.coerce_const(v, t, tt, c)
                        self.declare(mk, it.scope, Sym('const', c.name, value=v, t=tt), c)
                elif it.kind == 'enum' and id(it) not in done:
                    nxt = 0
                    for m in it.members:
                        if m.value is not None:
                            v, t = self.eval_const(m.value, mk)
                            nxt = int(self.coerce_const(v, t, 'Long', m))
                        self.declare(mk, it.scope, Sym('const', m.name, value=nxt, t='Long', enum_member=True, public=it.scope != 'private'), m)
                        nxt += 1
                    self.declare(mk, it.scope, Sym('enum', it.name, public=it.scope != 'private'), it)
                elif it.kind == 'dim':
                    cls = self.classes.get(mk)
                    for v in it.vars:
                        t = self.type_of_decl(mk, v.name, v.suffix, v.type, v)
                        public = it.scope in ('public', 'global')
                        if cls:
                            if public and (v.dims is not None or is_udt(t)):
                                self.err("arrays and user-defined types are not allowed as Public members "
                                         "of class modules", v)
                            sym = Sym('var', v.name, t=t, cname='f_' + cident(v.name), array=v.dims is not None,
                                      ndims=len(v.dims) if v.dims else None, byref=False, storage='field',
                                      cls=mk, public=public)
                        else:
                            cname = f"G_{cident(mname)}_{cident(v.name)}"
                            sym = Sym('var', v.name, t=t, cname=cname, array=v.dims is not None,
                                      ndims=len(v.dims) if v.dims else None, byref=False, storage='global')
                        sym.fixed_dims = [self.dim_bounds(d, mk, v) for d in v.dims] if v.dims else None
                        sym.fixlen = self.fixlen_of(mk, v.type, v)
                        sym.autonew = self.autonew_of(t, v.type) if (v.type is not None and v.type.new) else None
                        sym.ext = self.ext_of(t, v.type)
                        sym.withevents = getattr(v, 'withevents', False)
                        if sym.withevents:
                            if not cls:
                                self.err("WithEvents is only valid in a class module", v)
                            tn = v.type.name.lower() if v.type is not None else ''
                            tn = tn[3:] if tn.startswith('vb.') else tn
                            tlib, _, tshort = tn.rpartition('.')
                            if t == 'Object' and ((tn in VB_OBJECT_TYPES and tn not in ('control', 'stdpicture', 'picture'))
                                                  or tlib in OCX_LIBS or (not tlib and tshort in OCX_TYPES)):
                                sym.wev_ctl = True       # WithEvents mText As TextBox / As Form / As TreeView: the events of
                                                         # a control or a form, by name (StdFont: none come)
                            elif t.startswith('O:') and t[2:] in self.classes and self.classes[t[2:]].is_form:
                                sym.wev_ctl = True       # WithEvents f As Form1: the events of the form
                            elif t in ('O:adodb.recordset', 'O:adodb.connection', 'O:stdformat.stddataformat'):
                                sym.wev_ctl = True       # the events of ADO: by name, as a control's
                                sym.wev_rt = True
                            elif t == 'Object' and self.ext_class(tn):
                                sym.wev_ext = self.ext_class(tn)[1]   # a class of a component: its events
                            elif not (t.startswith('O:') and t[2:] in self.classes):
                                self.err("WithEvents needs a class of this project", v)
                        if cls:
                            self.declare(mk, 'private', sym, v)
                            cls.fields.append(sym)
                            cls.members[v.name.lower()] = sym
                        else:
                            self.declare(mk, 'public' if public else 'private', sym, v)
                elif it.kind == 'proc':
                    rt = None
                    if it.pkind in ('function', 'property get'):
                        rt = self.type_of_decl(mk, it.name, it.suffix, it.rtype, it)
                        if getattr(it.rtype, 'isarray', False):
                            rt = 'A:' + rt
                    params = []
                    dll = getattr(it, 'lib', None) is not None
                    if dll and mk in self.classes and it.scope != 'private':
                        self.err("Declare statements are not allowed as Public members of class modules", it)
                    for p in it.params:
                        pt = self.type_of_decl(mk, p.name, p.suffix, p.type, p, allow_any=dll)
                        dv = None
                        if p.default is not None:
                            self._const_conv = True       # Optional x As Long = CLng(vbBlue): VB6 takes it (c4)
                            try:
                                v, t = self.eval_const(p.default, mk)
                            finally:
                                self._const_conv = False
                            dv = (self.coerce_const(v, t, pt, p), pt)
                        if p.optional and p.isarray:
                            self.err("an Optional parameter cannot be an array", p)
                        if p.byval and is_udt(pt) and not p.isarray:
                            self.err("user-defined types cannot be passed ByVal", p)
                        params.append(Sym('param', p.name, t=pt, byval=p.byval, optional=p.optional,
                                          default=dv, array=p.isarray, paramarray=getattr(p, 'paramarray', False),
                                          ext=self.ext_of(pt, p.type)))
                        if params[-1].paramarray and p is not it.params[-1]:
                            self.err("ParamArray must be the last parameter", p)
                    seen_opt = False
                    letval = params[-1] if it.pkind in ('property let', 'property set') and params else None
                    for p in params:
                        if p.optional:
                            seen_opt = True
                        elif seen_opt and p is not letval:   # Property Let P(Optional i, v): v is the value
                            self.err("parameters after an Optional parameter must also be Optional", it)
                    cls = self.classes.get(mk)
                    prefix = {'sub': 'P', 'function': 'P', 'property get': 'PG', 'property let': 'PL',
                              'property set': 'PS'}[it.pkind]
                    sym = Sym('proc', it.name, pkind=it.pkind, rt=rt, params=params, node=it, module=mk,
                              cname=f"{prefix}_{cident(mname)}_{cident(it.name)}", cls=mk if cls else None,
                              public=it.scope != 'private')
                    it.sym = sym
                    if dll:
                        sym.dll = True
                        sym.cls = None
                        sym.cname = f"DD_{cident(mname)}_{cident(it.name)}"
                        self.protos.append(f'static vb_declare {sym.cname} = {{ "{c_utf8(it.lib)}", '
                                           f'"{c_utf8(it.alias)}", 0, 0 }};')
                    scope = 'private' if (it.scope == 'private' or cls) else 'public'
                    if it.pkind.startswith('property'):
                        if it.pkind in ('property let', 'property set') and (
                                params[-1].optional):
                            self.err("the value parameter of a Property Let/Set cannot be Optional or an array", it)
                        key = it.name.lower()
                        prop = self.mscopes[mk].get(key)
                        if prop is not None and prop.kind == 'enum':   # Property Get Align() As Align
                            prop = None
                        if prop is None:
                            prop = Sym('prop', it.name, get=None, let=None, set=None, cls=mk if cls else None,
                                       public=False, module=mk)
                            self.declare(mk, scope, prop, it)
                        elif prop.kind != 'prop':
                            self.err(f"ambiguous name detected: {it.name}", it)
                        slot = it.pkind.split()[1]
                        if getattr(prop, slot) is not None:
                            self.err(f"ambiguous name detected: {it.name}", it)
                        setattr(prop, slot, sym)
                        prop.public = prop.public or sym.public
                        if cls:
                            cls.members[key] = prop
                    else:
                        self.declare(mk, scope, sym, it)
                        if cls:
                            cls.members[it.name.lower()] = sym
                            low = it.name.lower()
                            pfx = 'usercontrol' if cls.is_uctl else 'form' if cls.is_form else 'class'
                            init_name, term_name = pfx + '_initialize', pfx + '_terminate'
                            if low == init_name and it.pkind == 'sub' and not params:
                                cls.init = sym
                            if low == term_name and it.pkind == 'sub' and not params:
                                cls.term = sym
            if mk in self.classes:
                for it in ast.items:
                    if it.kind == 'event':
                        ci = self.classes[mk]
                        if it.name.lower() in ci.events:     # (an Event and a Sub may have the same name)
                            self.err(f"ambiguous name detected: {it.name}", it)
                        it.index = len(ci.events)
                        it.psyms = [Sym('param', p.name, t=self.type_of_decl(mk, p.name, p.suffix, p.type, p),
                                        byval=p.byval, optional=False, default=None, array=p.isarray)
                                    for p in it.params]
                        it.sym = Sym('proc', it.name, pkind='sub', rt=None, params=it.psyms, node=it, module=mk,
                                     cname=f"RE_{cident(mk)}_{cident(it.name)}", cls=mk, public=False)
                        ci.events[it.name.lower()] = it
                    elif it.kind == 'implements':
                        iname = it.name
                        if '.' in iname:          # Implements Lib.IFace: the copy of the interface in this program
                            q, _, last = iname.partition('.')
                            if last.lower() in self.classes and (q.lower() == self.project_name.lower() or any(
                                    q.lower() == c.name.lower() for c in self.components)):
                                iname = last
                        t = self.resolve_type(mk, iname, it)
                        if not (t.startswith('O:') and t[2:] in self.classes) or t[2:] == mk:
                            self.err(f"bad interface for Implements: {it.name}", it)
                        self.classes[mk].implements.append((t[2:], it))
                for target, attr, val in ast.attrs:
                    if attr.lower() == 'vb_usermemid' and val == 0 and target:
                        self.classes[mk].defmember = target.lower()
                    if attr.lower() == 'vb_usermemid' and val == -4 and target:     # NewEnum: For Each
                        self.classes[mk].enummember = target.lower()
                    # Procedure Attributes > "Property is data bound" (FUNCFLAG_FBINDABLE 0x4): what a DataRepeater binds
                    if attr.lower() == 'vb_memberflags' and target and isinstance(val, str):
                        try:
                            flags = int(val, 16)
                        except ValueError:
                            flags = 0
                        b = self.classes[mk].bindable
                        if flags & 0x4 and all(x.lower() != target.lower() for x in b):
                            b.append(target)

    @staticmethod
    def tolerated(old, new):
        """the first of two declarations with the same name that VB6 lets stand side by side: an Enum and a
        member (Property Get Align() As Align), the members of two Enums (None = 0), the same API Declared
        in two modules; None when the names clash"""
        if old.kind == 'enum' and new.kind != 'enum':
            return new
        if new.kind == 'enum':
            return old
        if getattr(old, 'enum_member', False) and getattr(new, 'enum_member', False):
            return old
        if getattr(old, 'dll', False) and getattr(new, 'dll', False):
            return old
        return None

    def declare(self, mk, scope, sym, node):
        key = sym.name.lower()
        ms = self.mscopes[mk]
        if key in ms:
            keep = self.tolerated(ms[key], sym)
            if keep is None:
                self.err(f"ambiguous name detected: {sym.name}", node)
            if keep is not sym:
                return
        if mk in self.classes and sym.kind in ('var', 'proc', 'prop'):
            scope = 'private'            # members of a class are reached through an object
        ms[key] = sym
        if scope in ('public', 'global'):
            old = self.globals.get(key)
            if old is not None and old is not sym:
                keep = self.tolerated(old, sym) if old != 'ambiguous' else None
                self.globals[key] = keep if keep is not None else 'ambiguous'
            else:
                self.globals[key] = sym

    def dim_bounds(self, d, mk, node, local=None):
        lb, ub = d
        base = self.mopts[mk].get('base', 0) if hasattr(self, 'mopts') else 0
        lo = int(self.coerce_const(*self.eval_const(lb, mk, local), 'Long', node)) if lb is not None else base
        hi = int(self.coerce_const(*self.eval_const(ub, mk, local), 'Long', node))
        if hi < lo:
            self.err("range has no values", node)
        return lo, hi

    # ---------------------------------------------------------------- constant expressions
    def eval_const(self, e, mk, local=None):
        k = e.kind
        if k == 'num':
            return e.value, e.type
        if k == 'str':
            return e.value, 'String'
        if k == 'bool':
            return (-1 if e.value else 0), 'Boolean'
        if k == 'date':
            v = date_literal(e.text)
            if v is None:
                self.err(f"invalid date literal #{e.text}#", e)
            return v, 'Date'
        if k == 'paren':
            return self.eval_const(e.e, mk, local)
        if k == 'null':                          # Const c = Null, Optional v As Variant = Null
            return 'Null', 'Variant'
        if k == 'nothing':                       # Optional o As Object = Nothing
            return 'Nothing', 'Object'
        if k == 'empty':
            return None, 'Variant'
        if k == 'member' and getattr(e.base, 'kind', '') in ('name', 'member') and not e.suffix:
            # EnumName.Member, Library.Constant: the constant itself
            s = self.lookup_const(e.name, mk, local, getattr(e, 'line', 0))
            if s is not None:
                return s
        if k == 'name':
            s = self.lookup_const(e.name, mk, local, getattr(e, 'line', 0))
            if s is None:
                self.err(f"constant expression required: {e.name}", e)
            return s
        if k == 'apply' and getattr(e.base, 'kind', '') == 'name' and len(e.args) == 1 and e.args[0] is not None \
                and e.args[0].kind != 'named':
            # Optional c As Long = CLng(vbBlue): VB6 takes a conversion of a constant (compile-6.8 c4)
            to = {'clng': 'Long', 'cint': 'Integer', 'cbyte': 'Byte', 'cdbl': 'Double', 'csng': 'Single',
                  'ccur': 'Currency', 'cbool': 'Boolean', 'cstr': 'String'}.get(e.base.name.lower())
            if to and getattr(self, '_const_conv', False) and not e.base.suffix \
                    and self.lookup_const(e.base.name, mk, local) is None:
                v, t = self.eval_const(e.args[0], mk, local)
                return self.coerce_const(v, t, to, e), to
        if k == 'un':
            v, t = self.eval_const(e.e, mk, local)
            if e.op == '-':
                return -v, ('Integer' if t == 'Boolean' else t)
            if e.op == 'NOT':
                return ~int(v), t
        if k == 'bin':
            a, ta = self.eval_const(e.l, mk, local)
            b, tb = self.eval_const(e.r, mk, local)
            op = e.op
            if op == '&':
                return self.const_str(a, ta) + self.const_str(b, tb), 'String'
            if op == '+' and ta == 'String' and tb == 'String':
                return a + b, 'String'
            if ta == 'String' or tb == 'String':
                self.err("constant expression required", e)
            t = max(ta, tb, key=lambda x: RANK.get(x, 0))
            if t in ('Boolean', 'Byte'):
                t = 'Integer'
            if op in ('+', '-', '*'):
                v = a + b if op == '+' else a - b if op == '-' else a * b
                if t == 'Integer' and not -32768 <= v <= 32767:
                    t = 'Long'
                if t == 'Long' and not -2147483648 <= v <= 2147483647:
                    t = 'Double'
                return v, t
            if op == '/':
                return a / b, 'Double'
            if op == '\\':
                return int(round(a) // round(b)) if (round(a) // round(b)) >= 0 else -(abs(round(a)) // abs(round(b))), t if t in ('Integer', 'Long') else 'Long'
            if op == 'MOD':
                q = int(round(a)) % int(round(b))
                return (q if a >= 0 else -(abs(int(round(a))) % abs(int(round(b))))), t if t in ('Integer', 'Long') else 'Long'
            if op == '^':
                return float(a) ** b, 'Double'
            if op in ('AND', 'OR', 'XOR'):
                x, y = int(a), int(b)
                return (x & y if op == 'AND' else x | y if op == 'OR' else x ^ y), t
            if op in ('=', '<>', '<', '>', '<=', '>='):
                r = {'=': a == b, '<>': a != b, '<': a < b, '>': a > b, '<=': a <= b, '>=': a >= b}[op]
                return (-1 if r else 0), 'Boolean'
        self.err("constant expression required", e)

    def const_str(self, v, t):
        if t == 'String':
            return v
        if t == 'Boolean':
            return 'True' if v else 'False'
        if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return str(v)

    def coerce_const(self, v, t, to, node):
        if to == t or to == 'Variant' or (t == 'Object' and is_obj(to)):
            return v
        if t == 'Object':
            self.err("type mismatch", node)
        if t == 'Variant':                       # Null / Empty to a typed constant or parameter
            if v is None:
                return '' if to == 'String' else 0
            self.err("invalid use of Null", node)
        if to == 'String':
            return self.const_str(v, t)
        if t == 'String':                        # Const C As Long = "&H80000017": as CLng("&H80000017")
            dec, grp = locale_numsep()           # the regional separators (Greek: "3.5" = 35, compile-6.8 c14)
            txt = v.strip().replace(grp, '').replace(dec, '.') if not v.strip()[:2].upper() in ('&H', '&O') else v.strip()
            try:
                if txt[:2].upper() in ('&H', '&O'):
                    n = int(txt[2:].rstrip('&'), 16 if txt[1] in 'hH' else 8)
                    lim = 1 << (16 if len(txt[2:].rstrip('&')) <= 4 and not txt.endswith('&') else 32)
                    v = n - lim if n >= lim // 2 else n
                else:
                    v = float(txt)
            except ValueError:
                self.err("type mismatch", node)
        if to in ('Integer', 'Long', 'Byte'):
            r = int(round(v))  # Python rounds half to even, like VB
            lim = {'Integer': (-32768, 32767), 'Long': (-2147483648, 2147483647), 'Byte': (0, 255)}[to]
            if not lim[0] <= r <= lim[1]:
                self.err("overflow", node)
            return r
        if to == 'Boolean':
            return -1 if v else 0
        return float(v)

    def lookup_const(self, name, mk, local, ref_line=0):
        key = name.lower()
        if local and key in local and local[key].kind == 'const':
            s = local[key]
            return s.value, s.t
        own = self.mscopes[mk].get(key)
        if own is not None and ref_line and getattr(own, 'decl_line', 0) > ref_line:
            return None                          # a constant of this module declared further down: VB6 does not
                                                 # know it yet ("Constant expression required", compile-6.8 c29/c31)
        s = own or (self.globals.get(key) if self.globals.get(key) != 'ambiguous' else None)
        if s is not None and s.kind == 'const':
            return s.value, s.t
        if key in self.consts:
            return self.consts[key]
        return None

    # ---------------------------------------------------------------- program
    def compile(self):
        if getattr(self, 'from_vbp', False) and 'scripting' not in self.typelibs:
            # a project without a reference to the Scripting Runtime: Normal, Directory ... are not constants
            from .symbols import lib_const_names
            sc = lib_const_names('Scripting') - self.non_scripting_consts()
            for k in sc:
                self.consts.pop(k, None)
        self.mopts = {name.lower(): ast.options for name, ast in self.modules}
        self.collect()
        self.link_classes()
        body = []
        for mname, ast in self.modules:
            self.cur_mod_file = ast.file
            mk = mname.lower()
            for s in self.mscopes[mk].values():
                if s.kind == 'var' and s.storage == 'global':
                    self.global_var(s, body)
        for mname, ast in self.modules:
            self.cur_mod_file = ast.file
            for it in ast.items:
                if it.kind == 'proc' and not getattr(it.sym, 'dll', False):
                    pe = getattr(it, 'parse_error', None)
                    if pe is not None:              # a syntax error in it (parser.RECOVER): a stub, as below
                        if not self.on_demand:
                            raise pe
                        if self.on_demand_warn:
                            self.warnings.append((pe.file or ast.file, pe.line,
                                                  f"compile error when called: {pe.msg} (line {pe.line})"))
                        body.extend(ProcGen(self, mname.lower(), ast, it.sym).gen(stub=pe))
                        continue
                    if not self.on_demand:
                        body.extend(ProcGen(self, mname.lower(), ast, it.sym).gen())
                        continue
                    # Compile On Demand (a run of the IDE, --debug): VB6 compiles a procedure when it is first
                    # called, so a compile error in a procedure the program never calls does not stop the run;
                    # the procedure becomes a stub that reports the compile error when it is called
                    np, na = len(self.protos), set(self.addressof)
                    try:
                        body.extend(ProcGen(self, mname.lower(), ast, it.sym).gen())
                    except VBError as e:
                        del self.protos[np:]
                        for k in set(self.addressof) - na:
                            del self.addressof[k]
                        if self.on_demand_warn:
                            self.warnings.append((e.file or ast.file, e.line,
                                                  f"compile error when called: {e.msg} (line {e.line})"))
                        body.extend(ProcGen(self, mname.lower(), ast, it.sym).gen(stub=e))
        if self.ctx:
            self.protos.append("static vb_ctxdef LB_CTX;   /* the copies of the server (SingleUse) */")
        for key in self.classes:
            self.class_code(key)
        main = self.entry_point()
        if self.ctx:
            body += self.ctx_code()
        head = ['/* generated by lbvc (LinBASIC Studio) - do not edit */', '#include "vbrt.h"', '']
        if self.uses_forms or any(ci.is_form for ci in self.classes.values()):
            head.insert(1, '/* lbvc:forms */')
        if self.needs_db():
            head.insert(1, '/* lbvc:db */')
        for s, cname in self.strings.items():
            units = unescape_bytes(s).encode('utf-16-le')
            vals = [str(int.from_bytes(units[i:i + 2], 'little')) for i in range(0, len(units), 2)]
            head.append(f"static const vb_wchar {cname}[] = {{{', '.join(vals + ['0'])}}};")
        head.append('')
        head += self.res_code()
        head += self.types_code()
        return '\n'.join(head + self.protos + [''] + body + self.class_support + list(self.addressof.values())
                         + main) + '\n'

    def global_var(self, s, body):
        self.global_cnames.append(s.cname)
        if s.array:
            body.append(f"static vbarr {s.cname};")
            if s.fixed_dims:
                self.inits.extend(self.adim_code(f"&{s.cname}", s.t, s.fixed_dims, fixed=True,
                                                 fixlen=getattr(s, 'fixlen', None)))
        else:
            body.append(f"static {ctype(s.t)} {s.cname} = {zero(s.t)};")
            if is_udt(s.t):
                self.inits.append(f"U_{s.t[2:]}_init(&{s.cname});")
            if getattr(s, 'fixlen', None):
                self.inits.append(f"vb_sfix_init(&{s.cname}, {s.fixlen});")

    @property
    def ctx(self):
        """an ActiveX EXE with SingleUse classes: each of their objects a copy of the server of its own"""
        return self.component and self.ptype == 'oleexe' and any(
            ci.single_ast and ci.exposed for ci in self.classes.values())

    def ctx_code(self):
        """the globals of a copy of the server: swapped with a buffer when a call goes to an object of
        another copy; a new copy starts with them empty, the initialisations, and its Sub Main"""
        names = self.global_cnames + self.static_cnames
        out = ["static void LB_swap(void *p)", "{", "    char *b = (char *)p; unsigned long o = 0; (void)b; (void)o;"]
        for n in names:                          # byte by byte: the generated C needs no header for it
            out.append(f"    {{ char *g = (char *)&{n}; unsigned long i; for (i = 0; i < sizeof({n}); i++) "
                       f"{{ char t = g[i]; g[i] = b[o + i]; b[o + i] = t; }} o += sizeof({n}); }}")
        out.append("}")
        size = ' + '.join(f"sizeof({n})" for n in names) or '1'
        out.append(f"static unsigned long LB_size(void) {{ return (unsigned long)({size}); }}")
        out += ["static void LB_init(void)", "{"] + ['    ' + i for i in self.inits] + ["}"]
        main = self.component_main()
        out.append(f"static vb_ctxdef LB_CTX = {{ LB_size, LB_swap, LB_init, {main.cname if main else '0'}, 0 }};")
        return out

    def component_main(self):
        st = self.startup.strip().strip('"').lower()
        if st != 'sub main':
            return None
        main = self.globals.get('main')
        if not (isinstance(main, Sym) and main.kind == 'proc'):
            main = next((self.mscopes[mk].get('main') for mk in self.mscopes
                         if isinstance(self.mscopes[mk].get('main'), Sym)), None)
        return main

    def adim_code(self, ref, t, bounds, fixed, fixlen=None):
        """C statements dimensioning the array at ref with constant bounds [(lo, hi)]"""
        out = []
        if is_udt(t):
            out.append(f"({ref})->udt = &UD_{t[2:]};")
        b = ', '.join(f"{lo}, {hi}" for lo, hi in bounds)
        out.append(f"vb_adim({ref}, sizeof({ctype(t)}), {akind(t)}, {1 if fixed else 0}, {len(bounds)}, {b});")
        if fixlen:                               # Dim a(9) As String * 8: each element is 8 characters
            out.append(f"vb_afix_init({ref}, {fixlen});")
        return out

    def cls_ref(self, t):
        """C pointer to the vb_class of an object type (0 for Object)"""
        if t == 'Object':
            return '0'
        if t == 'O:collection' and 'collection' not in self.classes:
            return '&vb_cls_collection'
        if t in ('O:stdfont', 'O:stdpicture', 'O:propertybag') and t[2:] not in self.classes:
            return f"vbf_cls_{t[2:]}()"
        if t.startswith('O:scripting.'):
            return f"&vb_cls_{t[12:]}"
        if t.startswith('O:dao.'):
            return f"&vb_cls_dao_{t[6:]}"
        if t.startswith('O:adodb.'):
            return f"&vb_cls_ado_{t[8:]}"
        if t.startswith('O:rdo.'):
            return f"&vb_cls_rdo_{t[6:]}"
        if t.startswith('O:msbind.'):
            return f"&vb_cls_{t[9:]}"
        if t.startswith('O:stdformat.'):
            return f"&vb_cls_{t[12:]}"
        if t.startswith('P:'):                  # a class of a component (Dim x As New Proj.Class)
            return f'vb_component_class("{c_utf8(t[2:])}")'
        return f"&CLS_{cident(t[2:])}"

    # ---------------------------------------------------------------- user-defined types
    def field_init(self, ref, f):
        """C statements setting up field f at ref (pointer to the field)"""
        if f.dims:
            return self.adim_code(ref, f.t, f.fixed_dims, fixed=True, fixlen=getattr(f, 'fixlen', None))
        if f.dims is not None and is_udt(f.t):
            return [f"({ref})->udt = &UD_{f.t[2:]};"]
        if is_udt(f.t):
            return [f"U_{f.t[2:]}_init({ref});"]
        if getattr(f, 'fixlen', None):
            return [f"vb_sfix_init({ref}, {f.fixlen});"]
        return []

    def field_free(self, ref, f):
        if f.dims is not None:
            return [f"vb_afree({ref});"]
        c = udt_free_code(f.t, ref)
        return [c] if c else []

    def field_copy(self, d, s, f):
        """copy field f from s to d (C lvalues)"""
        if f.dims is not None:
            return f"vb_acopy(&{d}, &{s});"
        t = f.t
        if t == 'String':
            return f"vb_sset(&{d}, {s});"
        if t == 'Variant':
            return f"vb_vset(&{d}, {s});"
        if is_obj(t):
            return f"vb_oset(&{d}, {s});"
        if is_udt(t):
            return f"U_{t[2:]}_copy(&{d}, &{s});"
        return f"{d} = {s};"

    def types_code(self):
        out = []
        for k in self.udt_order:
            out.append(f"typedef struct U_{k} U_{k};")
        for k in self.udt_order:
            it = self.udts[k]
            out.append(f"struct U_{k} {{")
            for f in it.fields:
                out.append(f"    {'vbarr' if f.dims is not None else ctype(f.t)} {f.cname};")
            out.append("};")
        for k in self.udt_order:
            out.append(f"static void U_{k}_init(void *p_); static void U_{k}_free(void *p_); "
                       f"static void U_{k}_copy(void *d_, const void *s_);")
            out.append(f"static const vb_udt UD_{k};")
        for k in self.udt_order:
            it = self.udts[k]
            fl = []
            for f in it.fields:
                arr = 0 if f.dims is None else (1 if f.fixed_dims else 2)
                fl.append(f"{{ {self.vbk(f.t, f.fixlen)}, {f.fixlen or 0}, VB_OFFSETOF(U_{k}, {f.cname}), {arr}, "
                          f"{'&UD_' + f.t[2:] if is_udt(f.t) else '0'} }}")
            out.append(f"static const vb_ufield UF_{k}[] = {{ {', '.join(fl)} }};")
            rn, fn = '0', '0'
            if getattr(it, 'public_obj', False):     # it can go in a Variant: its name, and the names of its fields
                rn, fn = f'"{c_utf8(it.rec_name)}"', f"UN_{k}"
                out.append(f"static const char *const UN_{k}[] = {{ "
                           f"{', '.join(chr(34) + c_utf8(f.name) + chr(34) for f in it.fields)} }};")
            out.append(f"static const vb_udt UD_{k} = {{ sizeof(U_{k}), U_{k}_init, U_{k}_free, U_{k}_copy, "
                       f"UF_{k}, {len(it.fields)}, {rn}, {fn} }};")
        for k in self.udt_order:
            it = self.udts[k]
            init = [c for f in it.fields for c in self.field_init(f"&p->{f.cname}", f)]
            free = [c for f in it.fields for c in self.field_free(f"&p->{f.cname}", f)]
            copy = [self.field_copy(f"d->{f.cname}", f"s->{f.cname}", f) for f in it.fields]
            out.append(f"static void U_{k}_init(void *p_) {{ U_{k} *p = (U_{k} *)p_; (void)p; {' '.join(init)} }}")
            out.append(f"static void U_{k}_free(void *p_) {{ U_{k} *p = (U_{k} *)p_; (void)p; {' '.join(free)} }}")
            out.append(f"static void U_{k}_copy(void *d_, const void *s_) {{ U_{k} *d = (U_{k} *)d_; "
                       f"const U_{k} *s = (const U_{k} *)s_; {' '.join(copy)} }}")
        # classes: instance structs and descriptors
        for key, ci in self.classes.items():
            k = cident(key)
            out.append(f"typedef struct C_{k} {{ vbobjh h;")
            for f in ci.fields:
                out.append(f"    {'vbarr' if f.array else ctype(f.t)} {f.cname};")
            out.append(f"}} C_{k};")
            out.append(f"static C_{k} D_{k};                   /* stands in for Nothing after error 91 */")
            out.append(f"static const vb_class CLS_{k};")
            if ci.is_form:
                out.append(f"static const vbf_formdef FD_{k};")
            if (ci.is_form and not ci.is_uctl) or ci.predeclared:
                out.append(f"static vbobj G_FORM_{k};           /* the default instance: {ci.name} */")
        out.append('')
        return out

    # ---------------------------------------------------------------- forms
    def frm_props(self, ci, ctl):
        """design-time properties of a control as C strings: name, "s<text>" / "n<number>" / "r<raw>" /
        "l<item>\\037<item>..." (List, ItemData from the .frx); BeginProperty blocks as Font.Name ..."""
        from .project.source import Property, PropertyBlock, FrxRef
        from .project.frx import FrxStore
        out = []

        def value(it):
            v = it.value
            if isinstance(v, FrxRef):
                from .project.fsutil import ci_path
                path = ci_path(os.path.join(ci.form_dir or '.', v.file))
                if not os.path.exists(path):
                    return None
                key = (path, ci.form_enc)
                st = self.frx_cache.get(key)
                if st is None:
                    offs = [x.value.offset for c in ci.form_root.walk() for x in c.items
                            if isinstance(x, Property) and isinstance(x.value, FrxRef) and x.value.file == v.file]
                    st = self.frx_cache[key] = FrxStore.load(path, offs, encoding=ci.form_enc or 'cp1252')
                d = st.decode(it.name, v)
                if isinstance(d, list):
                    return 'l' + '\037'.join(str(x) for x in d)
                if isinstance(d, str):
                    return 's' + d
                if isinstance(d, tuple) and len(d) == 2 and d[1]:
                    return 'x' + d[1].hex()     # a picture: its file (BMP, ICO, GIF, JPG ...) in hex
                if isinstance(d, (bytes, bytearray)) and it.name.lower() == 'bindings':
                    return 'x' + bytes(d).hex()   # the data binding of an .ocx control (MSFlexGrid)
                return None
            if isinstance(v, str) and it.raw_value is not None and it.raw_value.strip().startswith('"'):
                return 's' + v
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                return 'n' + (repr(v) if isinstance(v, float) else str(v))
            return 'r' + str(v)

        def blob(it):
            """the OleObjectBlob of a DBGrid: its design-time properties (project/oleblob.py)"""
            from .project.oleblob import decode
            from .project.fsutil import ci_path
            v = it.value
            path = ci_path(os.path.join(ci.form_dir or '.', v.file))
            if not os.path.exists(path):
                return
            key = (path, ci.form_enc)
            st = self.frx_cache.get(key)
            if st is None:
                offs = [x.value.offset for c in ci.form_root.walk() for x in c.items
                        if isinstance(x, Property) and isinstance(x.value, FrxRef) and x.value.file == v.file]
                st = self.frx_cache[key] = FrxStore.load(path, offs, encoding=ci.form_enc or 'cp1252')
            for k, x in decode(st.raw(v.offset)):
                if isinstance(x, str):
                    val = 's' + x
                elif isinstance(x, float):
                    val = 'n' + repr(x)
                else:
                    val = 'n' + str(x)
                out.extend([f'"{c_utf8(k)}"', f'"{c_utf8(val)}"'])

        def walk(items, prefix):
            for it in items:
                if isinstance(it, PropertyBlock):
                    walk(it.items, prefix + it.name + '.')
                elif isinstance(it, Property) and it.name.lower() == 'oleobjectblob' and isinstance(it.value, FrxRef):
                    blob(it)
                elif isinstance(it, Property):
                    val = value(it)
                    if val is not None:
                        out.extend([f'"{c_utf8(prefix + it.name)}"', f'"{c_utf8(val)}"'])
        walk(ctl.items, '')
        return out

    def form_def(self, ci, th):
        k = cident(ci.mk)
        sup = self.class_support
        root = ci.form_root
        ctls = [c for c in root.walk() if c is not root]
        pos = {id(c): i for i, c in enumerate(ctls)}
        sup.append(f"static const char *const FP_{k}[] = {{ {', '.join(self.frm_props(ci, root) + ['0'])} }};")
        rows = []
        for i, c in enumerate(ctls):
            sup.append(f"static const char *const CP_{k}_{i}[] = {{ {', '.join(self.frm_props(ci, c) + ['0'])} }};")
            parent = pos.get(id(c.parent), -1) if c.parent is not None else -1
            rows.append(f'{{ "{c_utf8(c.cls)}", "{c_utf8(c.name)}", {c.index if c.index is not None else -1}, '
                        f'{parent}, CP_{k}_{i} }}')
        sup.append(f"static const vbf_ctldef CD_{k}[] = {{ {', '.join(rows) or '{ 0, 0, 0, 0, 0 }'} }};")
        # event procedures: Form_Load, Command1_Click ...
        evs = []
        for low, m in ci.members.items():
            if m.kind != 'proc' or m.pkind != 'sub' or '_' not in low:
                continue
            obj = low.rsplit('_', 1)[0]
            if obj == ('mdiform' if ci.is_mdi else 'usercontrol' if ci.is_uctl else 'form') or obj in ci.controls:
                evs.append((low, th.proc(m)))
        tbl = ', '.join(f'{{ "{c_utf8(n)}", VBM_METHOD, {fn} }}' for n, fn in evs) or '{ 0, 0, 0 }'
        sup.append(f"static const vb_member EV_{k}[] = {{ {tbl} }};")
        for f in ci.fields:                     # WithEvents variables of controls: event name -> handler
            if getattr(f, 'wev_ctl', False):
                rows = [f'{{ "{c_utf8(n)}", VBM_METHOD, {th.proc(m)} }}' for n, m in f.wev_handlers] + ['{ 0, 0, 0 }']
                sup.append(f"static const vb_member HT_{k}_{f.cname}_m[{len(rows)}] = {{ {', '.join(rows)} }};")
        sup.append(f'static const vbf_formdef FD_{k} = {{ "{c_utf8(ci.name)}", FP_{k}, {len(ctls)}, CD_{k}, EV_{k}, {len(evs)}, {1 if ci.is_mdi else 2 if ci.is_uctl else 0} }};')

    # ---------------------------------------------------------------- classes
    def class_code(self, key):
        ci = self.classes[key]
        k = cident(key)
        out = self.class_support
        fld = lambda f: f"&((C_{k} *)o)->{f.cname}"
        init = []
        for f in ci.fields:
            if f.array:
                if f.fixed_dims:
                    init += self.adim_code(fld(f), f.t, f.fixed_dims, fixed=True, fixlen=getattr(f, 'fixlen', None))
                elif is_udt(f.t):
                    init.append(f"({fld(f)})->udt = &UD_{f.t[2:]};")
            elif is_udt(f.t):
                init.append(f"U_{f.t[2:]}_init({fld(f)});")
            elif getattr(f, 'fixlen', None):
                init.append(f"vb_sfix_init({fld(f)}, {f.fixlen});")
        if ci.init:
            init.append(f"{ci.init.cname}(o);")
        if self.ctx and not ci.is_form:          # the object belongs to a copy of the server
            single = 1 if ci.single_ast and ci.exposed else 0
            init = [f"void *lb_sv = vb_ctx_begin(o, &LB_CTX, {single});"] + init + ["vb_ctx_end(&LB_CTX, lb_sv);"]
        out.append(f"static void CI_{k}(vbobj o) {{ {' '.join(init)} }}")
        if ci.term:
            if self.ctx and not ci.is_form:
                out.append(f"static void CT_{k}(vbobj o) {{ void *lb_sv = vb_ctx_enter(o, &LB_CTX); {ci.term.cname}(o); "
                           f"vb_ctx_end(&LB_CTX, lb_sv); }}")
            else:
                out.append(f"static void CT_{k}(vbobj o) {{ {ci.term.cname}(o); }}")
        clear = [f"vbf_form_clear(o);"] if ci.is_form else []
        for f in ci.fields:
            if getattr(f, 'withevents', False):
                clear.append(f"vb_wev_set(o, {fld(f)}, 0, 0);")
            elif f.array:
                clear.append(f"vb_afree({fld(f)});")
            else:
                c = udt_free_code(f.t, fld(f))
                if c:
                    clear.append(c)
        out.append(f"static void CC_{k}(vbobj o) {{ {' '.join(clear)} }}")
        # late binding: one thunk per public member
        members = []
        th = ThunkGen(self, ci)
        if not ci.is_form:                       # WithEvents variables of controls, forms, ADO objects: event name -> handler
            for f in ci.fields:
                if getattr(f, 'wev_ctl', False):
                    rows = [f'{{ "{c_utf8(n)}", VBM_METHOD, {th.proc(m)} }}' for n, m in f.wev_handlers] + ['{ 0, 0, 0 }']
                    out.append(f"static const vb_member HT_{k}_{f.cname}_m[{len(rows)}] = {{ {', '.join(rows)} }};")
        for low, m in ci.members.items():
            if m.kind == 'var':
                if not m.public:
                    continue
                members.append((low, 'VBM_GET', th.field(m, 'get')))
                members.append((low, 'VBM_SET' if is_obj(m.t) else 'VBM_LET', th.field(m, 'set' if is_obj(m.t) else 'let')))
                if m.t == 'Variant':
                    members.append((low, 'VBM_SET', th.field(m, 'set')))
            elif m.kind == 'proc':
                if not m.public or (is_udt(m.rt or '') and not self.rec_ok(m.rt)):   # a Public Type: a Variant
                    continue
                members.append((low, 'VBM_METHOD', th.proc(m)))
            elif m.kind == 'prop':
                for slot, kind in (('get', 'VBM_GET'), ('let', 'VBM_LET'), ('set', 'VBM_SET')):
                    p = getattr(m, slot)
                    if p is not None and p.public and not (is_udt(p.rt or '') and not self.rec_ok(p.rt)):
                        members.append((low, kind, th.proc(p)))
        # a data source (DataSourceBehavior = vbDataSource): its GetDataMember, for the runtime's binders
        gdm = ci.members.get('class_getdatamember') or ci.members.get('usercontrol_getdatamember')
        if gdm is not None and gdm.kind == 'proc':
            members.append(('lb_getdatamember', 'VBM_METHOD', th.proc(gdm)))
            # DataMembers / DataMemberChanged of a class that is a data source: the runtime's (object.c)
            if not ci.is_uctl and 'datamembers' not in ci.members:
                members.append(('datamembers', 'VBM_GET', 'vb_ds_datamembers'))
            if not ci.is_uctl and 'datamemberchanged' not in ci.members:
                members.append(('datamemberchanged', 'VBM_METHOD', 'vb_ds_datamemberchanged'))
        # Implements: the implementations by "<interface>.<member>", for interface calls of other programs
        for ik, _ in ci.implements:
            iface = self.classes[ik]
            for low, m in iface.members.items():
                slots = [('proc', m)] if m.kind == 'proc' else \
                    [(sl, getattr(m, sl)) for sl in ('get', 'let', 'set') if getattr(m, sl, None) is not None] if m.kind == 'prop' else []
                for sl, p in slots:
                    ip = next((x for kk, x in getattr(p, 'impls', []) if kk == key), None)
                    if ip is None or not p.public or is_udt(ip.rt or ''):
                        continue
                    kd = {'proc': 'VBM_METHOD', 'get': 'VBM_GET', 'let': 'VBM_LET', 'set': 'VBM_SET'}[sl]
                    members.append((f"{iface.name.lower()}.{low}", kd, th.proc(ip)))
        # the data-bound properties: entries of no kind ("lb_bind:<Name>"), for the DataRepeater (ocx_drep.inc)
        for b in getattr(ci, 'bindable', []):
            if b.lower() in ci.members:
                members.append((f"lb_bind:{b}", '0', '0'))
        tbl = ', '.join(f'{{ "{c_utf8(n)}", {kd}, {fn} }}' for n, kd, fn in members) or '{ 0, 0, 0 }'
        out.append(f"static const vb_member MB_{k}[] = {{ {tbl} }};")
        dm = f'"{c_utf8(ci.defmember)}"' if ci.defmember else '0'
        ifl = '0'
        if ci.implements:
            out.append(f"static const vb_class *const IFL_{k}[] = {{ "
                       f"{', '.join('&CLS_' + cident(i) for i, _ in ci.implements)}, 0 }};")
            ifl = f"IFL_{k}"
        form = ', 0, 0'
        if ci.is_form:
            self.form_def(ci, th)
            form = f', vbf_form_member, &FD_{k}'
        em = getattr(ci, 'enummember', None)
        form += f', "{c_utf8(em)}"' if em else ', 0'
        form += f', "{c_utf8(ci.progid)}"' if ci.progid else ', 0'
        out.append(f'static const vb_class CLS_{k} = {{ "{c_utf8(ci.name)}", sizeof(C_{k}), CI_{k}, '
                   f'{"CT_" + k if ci.term else "0"}, CC_{k}, MB_{k}, {len(members)}, {dm}, 0, 0, {ifl}{form} }};')
        out.append('')

    @property
    def reg_name(self):
        """the C function a component registers itself with"""
        return f"lbvc_comp_{cident(self.project_name)}"

    def component_entry(self):
        """a component: no main; a function that registers its public classes with the runtime, and
        runs its Sub Main before its first object is made"""
        st = self.startup.strip().strip('"').lower()
        main = None
        if st == 'sub main':
            main = self.globals.get('main')
            if not (isinstance(main, Sym) and main.kind == 'proc'):
                main = next((self.mscopes[mk].get('main') for mk in self.mscopes
                             if isinstance(self.mscopes[mk].get('main'), Sym)), None)
            if main is None:
                raise VBError("Sub Main is not defined", 0, 0, self.project_name)
        rows = []
        for key, ci in self.classes.items():
            if ci.is_form or not ci.exposed:
                continue
            flags = ' | '.join(f for f, on in (('VBC_CREATABLE', ci.creatable), ('VBC_GLOBAL', ci.global_ns),
                                               ('VBC_SINGLE', self.ctx and ci.single_ast)) if on) or '0'
            rows.append(f'{{ "{c_utf8(key)}", &CLS_{cident(key)}, {flags} }}')
        out = []
        init = '0'
        if main is not None and not self.ctx:    # (with SingleUse classes each copy runs it: LB_CTX)
            out.append(f"static void lbvc_comp_main(void) {{ {main.cname}(); }}")
            init = 'lbvc_comp_main'
        out += [f'void {self.reg_name}(void)', '{']
        out += ['    ' + i for i in self.inits]
        for key, ci in self.classes.items():
            if ci.is_form and ci.is_mdi:
                out.append(f'    vbf_mdi_register(&G_FORM_{cident(key)}, &CLS_{cident(key)});')
            if ci.is_uctl:
                out.append(f'    vbf_uctl_register("{c_utf8(self.project_name.lower())}.{c_utf8(key)}", &CLS_{cident(key)});')
        out.append(f"    static const vb_progclass T[] = {{ {', '.join(rows) or '{ 0, 0, 0 }'} }};")
        out.append(f'    vb_register_component("{c_utf8(self.project_name)}", T, {len(rows)}, {init});')
        out.append('}')
        return out

    def needs_db(self):
        return bool({'dao', 'adodb', 'msbind', 'rdo'} & self.typelibs) or self.uses_db or \
            any(s.lower().startswith(('dao.dbengine', 'dao.privatedbengine', 'adodb.', 'rdo.', 'msrdo20.')) for s in self.strings)

    def check_init(self):
        """the run-time switches of the code VB6 makes: Remove Floating Point Error Checks"""
        return ['    vb_fpchecks = 0;'] if self.order[0] != 'pcode' and self.order[3] else []

    def entry_point(self):
        if self.component:
            return self.component_entry()
        st = self.startup.strip().strip('"')
        if st.lower() == '(none)':               # an ActiveX EXE started on its own, with no startup object
            return [f'void {c.reg}(void);' for c in self.components] + \
                   ['int main(int argc, char **argv)', '{', '    vb_init(argc, argv);'] + self.check_init() + \
                   [f'    {c.reg}();' for c in self.components] + ['    ' + i for i in self.inits] + \
                   ['    vb_end();', '    return 0;', '}']
        if st.lower() in ('sub main', ''):
            cands = [s for s in (self.globals.get('main'),) if isinstance(s, Sym) and s.kind == 'proc']
            if not cands:
                for mk in self.mscopes:
                    s = self.mscopes[mk].get('main')
                    if s is not None and s.kind == 'proc':
                        cands.append(s)
            if not cands:
                raise VBError("Sub Main is not defined", 0, 0, self.project_name)
            main = cands[0]
            if main.pkind != 'sub' or main.params:
                raise VBError("Main must be a Sub without parameters", main.node.line, 1, self.project_name)
        else:
            key = st.lower()
            if key not in self.classes or not self.classes[key].is_form:
                raise VBError(f"startup object not found: {st}", 0, 0, self.project_name)
            main = None
        forms = any(ci.is_form for ci in self.classes.values())
        pre = ['    vb_db_init();'] if self.needs_db() else []
        pre.insert(0, f'    vb_set_project("{c_utf8(str(self.app.get("title") or self.project_name).strip(chr(34)))}");')
        for key, ci in self.classes.items():
            if ci.is_form and ci.is_mdi:
                pre.append(f'    vbf_mdi_register(&G_FORM_{cident(key)}, &CLS_{cident(key)});')
            if ci.is_uctl:
                pre.append(f'    vbf_uctl_register("{c_utf8(self.project_name.lower())}.{c_utf8(key)}", &CLS_{cident(key)});')
        if main is not None:
            start = [f'    {main.cname}();', '    if (vb_errpend) vb_unhandled();']
        else:
            k = cident(st.lower())
            start = [f'    vbf_start(vbf_default(&G_FORM_{k}, &CLS_{k}));', '    if (vb_errpend) vb_unhandled();']
        if forms:
            start.append('    vbf_run();')
            # the program ends: the default instances go, so Form_Terminate runs (f_basic)
            for key, ci in self.classes.items():
                if ci.is_form and not ci.is_uctl:
                    start.append(f'    vb_ofree(&G_FORM_{cident(key)});')
        for key, ci in self.classes.items():       # the default instances of classes (DataEnvironment)
            if ci.predeclared:
                start.append(f'    vb_ofree(&G_FORM_{cident(key)});')
        dbg = []
        if self.debug:
            names = ', '.join(f'"{c_utf8(m)}"' for m, _ in self.modules)
            rows = []
            for mname, _ in self.modules:
                for s in self.mscopes[mname.lower()].values():
                    if s.kind == 'var' and s.storage == 'global' and s.t is not None:
                        vt = (f"VT_ARRAY | {36 if is_udt(s.t) else vt_of(s.t)}" if s.array else
                              "36" if is_udt(s.t) else "VT_DISPATCH" if is_obj(s.t) else vt_of(s.t))
                        rows.append(f'{{"{c_utf8(s.name)}", {vt}, (void *)&{s.cname}}}')
            for key, ci in self.classes.items():
                if ci.is_form and not ci.is_uctl:
                    rows.append(f'{{"{c_utf8(ci.name)}", VT_DISPATCH, (void *)&G_FORM_{cident(key)}}}')
            kinds = ''
            for mname, _ in self.modules:         # Break in Class Module / on Unhandled Errors: the class modules
                ci = self.classes.get(mname.lower())
                kinds += 'M' if ci is None else 'U' if ci.is_uctl else 'F' if ci.is_form else 'C'
            dbg = [f'    static const char *const vb_dbg_mods[] = {{ {names} }};',
                   f'    static const vb_dbgvar vb_dbg_g[] = {{ {", ".join(rows) or "{0, 0, 0}"} }};',
                   f'    vb_dbg_kinds("{kinds}");',
                   f'    vb_dbg_init(vb_dbg_mods, {len(self.modules)}, vb_dbg_g, {len(rows)});']
        if self.app.get('res'):
            dbg.append(f"    vb_res_init(VB_RES, {len(self.app['res'])});")
        if self.uses_forms and not forms:
            dbg.append("    vbf_use();")
        regs = [f'    {c.reg}();' for c in self.components]   # the components the program uses
        if any(getattr(c, 'uses_forms', False) for c in self.components) and not forms and not self.uses_forms:
            regs.append("    vbf_use();")
        protos = [f'void {c.reg}(void);' for c in self.components]
        return protos + ['int main(int argc, char **argv)', '{', '    vb_init(argc, argv);'] + self.check_init() + regs + dbg + \
               ['    ' + i for i in self.inits] + pre + start + ['    vb_end();', '    return 0;', '}']


class ProcGen:
    """C code for one procedure"""

    def __init__(self, comp, mk, ast, sym):
        self.c = comp
        self.mk = mk
        self.ast = ast
        self.sym = sym
        self.node = sym.node
        self.locals = {}
        self.decls = []               # declarations at the top of the C function
        self.cleanup = []             # code at exit (free strings, variants, arrays)
        self.lines = []
        self.ntmp = 0
        self.nstmt = 0
        self.dbg_lines = []                      # (line, statement) of --debug: Set Next Statement
        self.labels = {}              # VB label -> index
        self.goto_labels = set()
        self.loops = []               # stack of (kind, exit label)
        self.explicit = ast.options.get('explicit', False)
        self.compare_text = 1 if ast.options.get('compare') == 'text' else 0
        self.uses_eh = False
        self.withs = []               # stack of With targets (E)
        self.prologue = []
        self.cls = getattr(sym, 'cls', None)
        self.handler = has_onerror(self.node.body)   # On Error in the procedure: the optimizer of VB6 keeps the order
        self.sched = {}               # id(node) -> [node, E, uses]: calls evaluated before (evalorder.py)
        self.xdepth = 0               # depth of expr(): 0 for an expression of the statement itself
        self.sched_ok = not os.environ.get('LBVC_ORDER_OFF')   # (to find a bug of the scheduler)
        self.dry = 0                  # > 0: compiling only to learn a type (dry_type)
        self._dtc = {}

    def err(self, msg, node=None):
        raise VBError(msg, getattr(node, 'line', 0) or self.node.line, 1, self.ast.file)

    # the code VB6 makes (Compile tab): native code and its advanced optimizations (P-code has none)
    @property
    def native(self):
        return self.c.order[0] != 'pcode'

    @property
    def noovf(self):                             # Remove Integer Overflow Checks
        return self.native and self.c.order[1]

    @property
    def nobounds(self):                          # Remove Array Bounds Checks
        return self.native and self.c.order[2]

    @property
    def noflpt(self):                            # Remove Floating Point Error Checks
        return self.native and self.c.order[3]

    @property
    def unrounded(self):                         # Allow Unrounded Floating Point Operations
        return self.native and self.c.order[4]

    def tmp(self, t, init=True):
        self.ntmp += 1
        name = f"t{self.ntmp}"
        self.decls.append(f"{ctype(t) if known_type(t) else t} {name}" + (f" = {zero(t)};" if init and known_type(t)
                                                                      else " = 0;" if init and t == 'int' else ";"))
        return name

    def emit(self, s):
        self.lines.append(s)

    # ---------------------------------------------------------------- procedure
    def gen(self, stub=None):
        s = self.sym
        cparams = ['vbobj vb_self'] if self.cls else []
        prologue = self.prologue
        for p in s.params:
            pc = 'a_' + cident(p.name)
            if p.array:
                cparams.append(f"vbarr *{pc}")
                self.locals[p.name.lower()] = Sym('var', p.name, t=p.t, cname=pc, array=True, ndims=None,
                                                  byref=True, storage='param')
            elif p.byval:
                if p.t in ('String', 'Variant') or is_obj(p.t):
                    cparams.append(f"{ctype(p.t)} {pc}_in")
                    self.decls.append(f"{ctype(p.t)} {pc} = {zero(p.t)};")
                    prologue.append({'String': f"vb_sset(&{pc}, {pc}_in);", 'Variant': f"vb_vset(&{pc}, {pc}_in);"}
                                    .get(p.t, f"vb_oset(&{pc}, {pc}_in);"))
                    self.cleanup.append(udt_free_code(p.t, '&' + pc))
                else:
                    cparams.append(f"{ctype(p.t)} {pc}")
                self.locals[p.name.lower()] = Sym('var', p.name, t=p.t, cname=pc, array=False, byref=False,
                                                  storage='param')
            else:
                cparams.append(f"{ctype(p.t)} *{pc}")
                self.locals[p.name.lower()] = Sym('var', p.name, t=p.t, cname=pc, array=False, byref=True,
                                                  storage='param', optional=p.optional)
            self.locals[p.name.lower()].ismissing_ok = p.optional and p.t == 'Variant'
            self.locals[p.name.lower()].ext = getattr(p, 'ext', None)
        # Out of stack space (28): an endless recursion is an error the program can trap, as in VB6
        prologue.append("{ char vb_sc; if ((char *)&vb_sc < vb_stack_lim) { vb_raise(28); goto vb_exit; } }")
        if self.cls:
            prologue.append("if (!vb_self) { vb_raise(91); goto vb_exit; }")
        rt = s.rt
        if rt:
            self.decls.append(f"{ctype(rt)} vb_ret = {zero(rt)};")
            if is_udt(rt):
                prologue.insert(0, f"U_{rt[2:]}_init(&vb_ret);")
            self.locals[s.name.lower()] = Sym('var', s.name, t=rt, cname='vb_ret', array=False, byref=False,
                                              storage='ret')
        proto = f"static {ctype(rt) if rt else 'void'} {s.cname}({', '.join(cparams) or 'void'})"
        self.c.protos.append(proto + ';')
        if stub is None:
            self.collect_labels(self.node.body)
            self.block(self.node.body)
        else:                                    # Compile On Demand: the compile error, when it is called
            self.lines = [f"vb_compile_error({cstr_lit(stub.file or self.ast.file or '')}, "
                          f"{int(stub.line or 0)}, {cstr_lit(stub.msg)});", "goto vb_exit;"]
        out = [proto, '{', '    vb_i32 vb_m0 = vb_mark();',
               '    vb_i32 vb_ehmode = 0, vb_ehlabel = 0, vb_inh = 0, vb_cur = 0, vb_rs = 0, vb_line = 0;',
               '    (void)vb_ehmode; (void)vb_ehlabel; (void)vb_inh; (void)vb_cur; (void)vb_rs; (void)vb_line;']
        out += ['    ' + d for d in self.decls]
        if self.c.debug:
            out += ['    ' + d for d in self.debug_frame()]
            prologue.insert(0, "if (vb_dbg_on) vb_dbg_enter(&vb_dbgf);")
        out += ['    ' + p for p in prologue]
        out += ['    ' + l if not l.endswith(':;') else l for l in self.lines]
        out += ['vb_exit:', '    if (!vb_errpend && vb_inh) vb_err_clear();']
        if self.c.debug:
            out.append('    if (vb_dbg_on) vb_dbg_leave(&vb_dbgf);')
        out += ['    ' + c for c in self.cleanup]
        out.append('    vb_release(vb_m0);')
        if rt == 'String':
            out += ['    { vbs r = vb_scopy(vb_ret); vb_sfree(&vb_ret); return r; }']
        elif rt == 'Variant':
            out += ['    { vbvar r = vb_vtmp(vb_ret); vb_vfree(&vb_ret); return r; }']
        elif is_obj(rt):
            out += ['    return vb_otmp(vb_ret);']           # the reference of vb_ret goes to the caller
        elif rt:
            out.append('    return vb_ret;')
        else:
            out.append('    return;')
        # error dispatch
        out += ['vb_eh:',
                ('    if (vb_inh || vb_ehmode == 0) { if (vb_dbg_on) { vb_dbg_error(&vb_dbgf); if (vb_dbgf.jump) '
                 '{ vb_rs = vb_dbgf.jump; vb_dbgf.jump = 0; goto vb_resume_same; } } goto vb_exit; }'
                 if self.c.debug else '    if (vb_inh || vb_ehmode == 0) goto vb_exit;'),
                # Break on All Errors: an error the procedure handles stops too (then F5 runs the statement again)
                *(['    if (vb_dbg_on) { vb_dbg_trap(&vb_dbgf); if (vb_dbgf.jump) '
                   '{ vb_rs = vb_dbgf.jump; vb_dbgf.jump = 0; goto vb_resume_same; } }'] if self.c.debug else []),
                '    vb_err_take(vb_line);',
                '    vb_rs = vb_cur;',
                '    if (vb_ehmode == 1) goto vb_resume_next;',
                '    vb_inh = 1;',
                '    switch (vb_ehlabel) {']
        for name, idx in sorted(self.labels.items(), key=lambda x: x[1]):
            out.append(f'    case {idx}: goto L{idx};')
        out += ['    }', '    goto vb_exit;',
                'vb_resume_next:', '    switch (vb_rs) {']
        out += [f'    case {k}: goto R{k};' for k in range(1, self.nstmt + 1)]
        out += ['    }', '    goto vb_exit;', 'vb_resume_same:', '    switch (vb_rs) {']
        out += [f'    case {k}: goto S{k};' for k in range(1, self.nstmt + 1)]
        out += ['    }', '    goto vb_exit;']
        if getattr(self, 'gs_declared', False):
            out += ['vb_gsret:', '    switch (vb_gs[--vb_gsn]) {']
            out += [f'    case {k}: goto G{k};' for k in range(1, getattr(self, 'ngosub', 0) + 1)]
            out += ['    }', '    goto vb_exit;']
        out += ['}', '']
        return out

    def debug_frame(self):
        """lbvc --debug: the variables of the procedure for the Locals window and the Immediate"""
        def entry(name, t, ptr, array):
            if array:
                vt = f"VT_ARRAY | {36 if is_udt(t) else vt_of(t)}"
            elif is_udt(t):
                vt = "36"
            elif is_obj(t):
                vt = "VT_DISPATCH"
            else:
                vt = vt_of(t)
            return f'{{"{c_utf8(name)}", {vt}, (void *)({ptr})}}'
        rows = []
        for key, s in self.locals.items():
            if s.kind != 'var' or s.t is None:
                continue
            if s.array:
                ptr = s.cname if s.byref else f"&{s.cname}"
            else:
                ptr = s.cname if s.byref else f"&{s.cname}"
            rows.append(entry(s.name, s.t[2:] if s.array and is_arr(s.t) else s.t, ptr, s.array))
        mrows = []
        for key, s in self.c.mscopes.get(self.mk, {}).items():
            if s.kind != 'var' or s.t is None:
                continue
            if s.storage == 'field':
                if not self.cls:
                    continue
                ptr = f"vb_self ? (void *)&((C_{cident(s.cls)} *)vb_self)->{s.cname} : 0"
            elif s.storage == 'global':
                ptr = f"&{s.cname}"
            else:
                continue
            mrows.append(entry(s.name, s.t, ptr, s.array))
        mid = self.c.modid.get(self.mk, -1)
        name = c_utf8(self.sym.name)
        first = {}                               # a line: its first statement (Set Next Statement)
        for ln, n in self.dbg_lines:
            first.setdefault(ln, n)
        pairs = ", ".join(f"{ln}, {n}" for ln, n in sorted(first.items())) or "0, 0"
        return [f"vb_dbgvar vb_dbgv[] = {{ {', '.join(rows) or '{0, 0, 0}'} }};",
                f"vb_dbgvar vb_dbgm[] = {{ {', '.join(mrows) or '{0, 0, 0}'} }};",
                f"static const vb_i32 vb_dbgl[] = {{ {pairs} }};",
                f'vb_dbgframe vb_dbgf = {{ "{name}", {mid}, 0, {len(rows)}, vb_dbgv, {len(mrows)}, vb_dbgm, 0, 0, '
                f'&vb_ehmode, &vb_inh, {"vb_self" if self.cls else "0"}, vb_dbgl, {len(first)}, &vb_cur, 0 }};',
                "(void)vb_dbgv; (void)vb_dbgm;"]

    def collect_labels(self, stmts):
        for st in stmts:
            if st.kind == 'label':
                key = st.name.lower()
                if key in self.labels:
                    self.err(f"label defined more than once: {st.name}", st)
                self.labels[key] = len(self.labels) + 1
            for attr in ('then', 'els', 'body'):
                sub = getattr(st, attr, None)
                if isinstance(sub, list):
                    self.collect_labels(sub)
            if st.kind == 'if':
                for _, b in st.elifs:
                    self.collect_labels(b)
            if st.kind == 'select':
                for _, b, _ in st.cases:
                    self.collect_labels(b)

    # ---------------------------------------------------------------- names
    def lookup(self, name, node, create=True):
        key = name.lower()
        if getattr(node, 'lib_qual', False):     # VB.Load, VBA.Left: the library's, not a name of the program
            if key in self.c.consts:             # VBA.vbNullString, VBRUN.vbLeftButton, DAO.dbText: its constant
                v, t = self.c.consts[key]
                return Sym('const', name, value=v, t=t)
            return None
        qm = getattr(node, 'qual_mod', None)
        if qm:                                   # Module1.Name
            s = self.c.mscopes[qm].get(key)
            if s is not None:
                return s
        s = self.locals.get(key)
        if s:
            return s
        if key == 'me':
            if not self.cls:
                self.err("invalid use of Me keyword", node)
            return Sym('me', 'Me', t='O:' + self.cls)
        s = self.c.mscopes[self.mk].get(key)
        if s:
            return s
        if self.cls and self.c.classes[self.cls].is_form and key in self.c.classes[self.cls].controls:
            return Sym('ctl', name, t='Object')
        if key == 'usercontrol' and self.cls and self.c.classes[self.cls].is_uctl:
            return Sym('me', 'UserControl', t='O:' + self.cls)       # UserControl.Enabled: Me
        if key in self.c.classes and ((self.c.classes[key].is_form and not self.c.classes[key].is_uctl) or
                                      self.c.classes[key].predeclared):
            return Sym('formdefault', self.c.classes[key].name, t='O:' + key, key=key)
        g = self.c.globals.get(key)
        if g == 'ambiguous':
            self.err(f"ambiguous name detected: {name}", node)
        if g:
            return g
        if key in self.c.module_names:
            if B.lookup(name, getattr(node, 'suffix', None)) is not None:
                return None                      # Exp(x) inside a class called Exp: the function of VBA
            self.err(f"expected variable or procedure, not module: {name}", node)
        if key in self.c.consts:
            v, t = self.c.consts[key]
            return Sym('const', name, value=v, t=t)
        return None

    def implicit_var(self, name, suffix, node):
        if self.explicit:
            self.err(f"variable not defined: {name}", node)
        t = self.c.type_of_decl(self.mk, name, suffix, None, node)
        return self.new_local(name, t, node)

    def new_local(self, name, t, node, array=False, ndims=None, static=False):
        key = name.lower()
        if key in self.locals and self.locals[key].storage != 'ret':
            self.err(f"duplicate declaration in current scope: {name}", node)
        cname = ('st_' if static else 'v_') + cident(name)
        if static:
            cname = f"st_{cident(self.mk)}_{cident(self.sym.name)}_{cident(name)}"
            self.c.static_cnames.append(cname)
        sym = Sym('var', name, t=t, cname=cname, array=array, ndims=ndims, byref=False,
                  storage='static' if static else 'local')
        self.locals[key] = sym
        decl_target = self.c.protos if static else self.decls
        if array:
            decl_target.append(f"static vbarr {cname};" if static else f"vbarr {cname} = {{0}};")
            if not static:
                self.cleanup.append(f"vb_afree(&{cname});")
            if is_udt(t):
                self.prologue.append(f"{cname}.udt = &UD_{t[2:]};")
        else:
            decl_target.append(f"static {ctype(t)} {cname} = {zero(t)};" if static else f"{ctype(t)} {cname} = {zero(t)};")
            if is_udt(t):
                if static:
                    self.c.protos.append(f"static int {cname}_set;")
                    self.prologue.append(f"if (!{cname}_set) {{ {cname}_set = 1; U_{t[2:]}_init(&{cname}); }}")
                else:
                    self.prologue.append(f"U_{t[2:]}_init(&{cname});")
            if not static:
                c = udt_free_code(t, '&' + cname)
                if c:
                    self.cleanup.append(c)
        sym.autonew = None
        return sym

    # ---------------------------------------------------------------- conversions
    def conv(self, e, to, node=None):
        """code of expression e converted to type 'to'"""
        f = e.t
        c = e.code
        if c == 'vb_err_number()' and (to == 'Object' or is_obj(to)):
            return 'vb_err_obj()'                # ShowErrorMsg Err: the Err object itself
        if e.ext and c is not None:              # an extended value: rounded to its type (error 6 when too big)
            c = f"vb_rf({c})" if f == 'Single' else f"vb_rd({c})"
            if f == to:
                return c
        if e.arr is not None and c is None:     # a whole array
            et = f[2:] if is_arr(f) else f
            if to == 'Variant':
                return f"vb_v_arr({e.arr}, {vt_of(et)})"
            if to == 'String' and et == 'Byte':  # s = b
                return f"vb_barr2s({e.arr})"
            self.err("type mismatch: an array cannot be used here", node)
        if f == to:
            return c
        if is_arr(f) and to == 'Variant' and e.ptr is not None:
            # the array a Function returns, by its name inside it (Call Fill(FuncName, ...) to a Variant): FreeImage
            return f"vb_v_arr({e.ptr}, {vt_of(f[2:])})"
        if f == 'void':
            self.err("expected function or variable", node)
        if is_udt(f) and to == 'Variant' and self.c.rec_ok(f):
            # a Public Type of a public class into a Variant: a copy of it (VT_RECORD)
            return f"vb_v_rec(&{self.udt_addr_value(e)}, &UD_{f[2:]})"
        if is_udt(f) or is_udt(to):
            self.err("type mismatch" + ("" if not (to == 'Variant' or f == 'Variant') else
                                        ": only user-defined types of public object modules can go into a Variant"),
                     node)
        if is_obj(to):
            if is_obj(f):
                if to == 'Object':
                    return c
                if f == 'Object':
                    return f"vb_ocast({c}, {self.c.cls_ref(to)})"
                if self.c.implements(f[2:], to[2:]):            # class -> its interface: a view
                    return f"vb_oview({c}, {self.c.cls_ref(to)})"
                if self.c.implements(to[2:], f[2:]):           # interface -> class: checked when it runs
                    return f"vb_ocast({c}, {self.c.cls_ref(to)})"
                self.err("type mismatch", node)
            if f == 'Variant':
                oc = f"vb_v2obj({c})"
                return oc if to == 'Object' else f"vb_ocast({oc}, {self.c.cls_ref(to)})"
            self.err("type mismatch (object required)", node)
        if is_obj(f):
            if to == 'Variant':
                return f"vb_v_obj({c})"
            if f != 'Object' and f[2:] in self.c.classes and not self.c.classes[f[2:]].defmember:
                self.err("type mismatch (an object has no value here)", node)
            return self.conv(E(f"vb_late_o({c}, \"\", VBM_GET, 0, 0)", 'Variant'), to, node)
        if to == 'Variant':
            return {'Integer': f"vb_v_i({c})", 'Long': f"vb_v_l({c})", 'Single': f"vb_v_f({c})",
                    'Double': f"vb_v_d({c})", 'Currency': f"vb_v_c({c})", 'Boolean': f"vb_v_bool({c})",
                    'Byte': f"vb_v_byte({c})", 'String': f"vb_v_s({c})", 'Date': f"vb_v_date({c})"}[f]
        if f == 'Variant':
            return {'Integer': f"vb_v2i({c})", 'Long': f"vb_v2l({c})", 'Single': f"vb_v2f({c})",
                    'Double': f"vb_v2d({c})", 'Currency': f"vb_v2c({c})", 'Boolean': f"vb_v2bool({c})",
                    'Byte': f"vb_v2byte({c})", 'String': f"vb_v2s({c})", 'Date': f"vb_v2date({c})"}[to]
        if to == 'String':
            return {'Integer': f"vb_cstr_i({c})", 'Long': f"vb_cstr_i({c})", 'Byte': f"vb_cstr_i({c})",
                    'Single': f"vb_cstr_f({c})", 'Double': f"vb_cstr_d({c})", 'Currency': f"vb_cstr_c({c})",
                    'Boolean': f"vb_cstr_b({c})", 'Date': f"vb_cstr_date({c})"}[f]
        if f == 'String':
            if to == 'Boolean':
                return f"vb_s2bool({c})"
            if to == 'Date':
                return f"vb_s2date({c})"
            d = E(f"vb_s2d({c})", 'Double')
            return self.conv(d, to, node)
        # numeric (Boolean, Date included) -> numeric
        if f == 'Currency':
            if to == 'Boolean':
                return f"(({c}) != 0 ? (vb_bool)-1 : (vb_bool)0)"
            c, f = f"vb_cy2d({c})", 'Double'
            if to == 'Double':
                return c
        if to == 'Integer':
            if f == 'Long' and self.noovf:     # Remove Integer Overflow Checks: the low 16 bits
                return f"((vb_i16)({c}))"
            return f"vb_cint((double)({c}))" if f in ('Long', 'Single', 'Double', 'Date') else f"((vb_i16)({c}))"
        if to == 'Long':
            return f"vb_clng((double)({c}))" if f in ('Single', 'Double', 'Date') else f"((vb_i32)({c}))"
        if to == 'Byte':
            if f in ('Integer', 'Long', 'Boolean') and self.noovf:
                return f"((vb_u8)({c}))"
            return f"vb_cbyte((double)({c}))"
        if to == 'Single':
            if f in ('Double', 'Date'):         # P-code: error 6 after the store; native code: no check (VB6)
                return f"vb_rf((vb_ld)({c}))" if not self.native else f"((float)({c}))"
            return f"((float)({c}))"
        if to in ('Double', 'Date'):
            return f"((double)({c}))"
        if to == 'Currency':
            return f"vb_ccur((double)({c}))"
        if to == 'Boolean':
            return f"(({c}) != 0 ? (vb_bool)-1 : (vb_bool)0)"
        self.err(f"type mismatch ({f} to {to})", node)

    def cv(self, e, to, node=None):
        return E(self.conv(e, to, node), to)

    # ---------------------------------------------------------------- expressions
    VBA_QUALS = {'vba', 'vb', 'strings', 'datetime', 'math', 'interaction', 'filesystem', 'conversion', 'information',
                 'financial', 'vbrun'}
    LIB_QUALS = {'vba', 'vb', 'vbrun', 'stdole', 'dao', 'adodb', 'msbind', 'rdo', 'msrdc', 'msadodclib', 'msdatagridlib',
                 'msdatalistlib', 'msdatareportlib', 'mscomdlg', 'scripting', 'mscomctllib', 'mscomctl2',
                 'msflexgridlib', 'mshierarchicalflexgridlib', 'tabdlg', 'richtextlib', 'mschart20lib', 'mci',
                 'mscommlib', 'picclip', 'comctl3', 'sysinfolib', 'msdbctls', 'msdbgrid', 'stdformat', 'mswinsocklib', 'inetctlsobjects',
                 'msmask', 'msdatarepeaterlib', 'msmapi'}

    def unqual(self, e):
        """VBA.Mid$(...), VB.Screen, DateTime.Now, VBA.FileSystem.Kill ...: the library and module names
        before a name of VBA or VB are left out (when the program has no name of its own like them)"""
        if e is None or e.kind not in ('member', 'apply'):
            return e
        if e.kind == 'apply':
            b = self.unqual(e.base)
            return e if b is e.base else N('apply', e.line, **{**{k: v for k, v in e.__dict__.items()
                                                                   if k not in ('kind', 'line')}, 'base': b})
        base = self.unqual(e.base) if e.base.kind == 'member' else e.base
        if base.kind == 'name' and getattr(base, 'suffix', None) is None:
            key = base.name.lower()                 # Module1.Proc: a name of a standard module
            if key in self.c.module_names and key not in self.c.classes and key not in self.locals \
                    and self.c.mscopes.get(key, {}).get(e.name.lower()) is not None:
                return N('name', e.line, name=e.name, suffix=e.suffix, qual_mod=key)
        if base.kind == 'name' and base.name.lower() in self.VBA_QUALS and base.suffix is None:
            key = base.name.lower()
            if self.lookup(base.name, base) is None and key not in self.c.classes and key not in self.c.module_names:
                self.check_lib_member(key, e)
                return N('name', e.line, name=e.name, suffix=e.suffix, lib_qual=True)
        if base.kind == 'name' and base.suffix is None and base.name.lower() == 'global' and (
                getattr(base, 'lib_qual', False) or self.lookup(base.name, base) is None):
            return N('name', e.line, name=e.name, suffix=e.suffix, lib_qual=True)    # VB.Global.Forms: Forms
        if base.kind == 'name' and base.suffix is None and e.suffix is None:
            # DAO.dbText, MSComctlLib.ListViewConstants.lvwReport, MouseButtonConstants.vbLeftButton:
            # a type library or one of its Enums before one of its constants (or Library.Enum.Constant)
            from .symbols import lib_enums
            key, low = base.name.lower(), e.name.lower()
            qual = key in self.LIB_QUALS or (key in lib_enums() and '.' not in key)
            nxt = low in self.c.consts or (key in self.LIB_QUALS and low in lib_enums())
            if qual and nxt and key in self.LIB_QUALS and not getattr(base, 'lib_qual', False):
                self.check_lib_member(key, e)
            if qual and nxt and (getattr(base, 'lib_qual', False) or (
                    self.lookup(base.name, base) is None and key not in self.c.classes
                    and key not in self.c.module_names and ('', key) not in self.c.enum_names
                    and (self.mk, key) not in self.c.enum_names)):
                return N('name', e.line, name=e.name, suffix=e.suffix, lib_qual=True)
        if base is not e.base:
            return N('member', e.line, base=base, name=e.name, suffix=e.suffix)
        return e

    _LIB_STEMS = None

    def check_lib_member(self, lib, e):
        """VBA.ColorConstants, VBA.vbLeftButton: an Enum or a constant of another library -> VB6's compile
        error (ColorConstants and vbLeftButton are VBRUN's)"""
        from .symbols import lib_enums, lib_const_names
        if ProcGen._LIB_STEMS is None:
            base = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'data', 'vb6', 'typelibs')
            ProcGen._LIB_STEMS = {f[:-5].lower(): f[:-5] for f in os.listdir(base) if f.endswith('.json')} \
                if os.path.isdir(base) else {}
        stem = ProcGen._LIB_STEMS.get(lib)
        if stem is None:
            return
        low = e.name.lower()
        if low in lib_enums() and f"{lib}.{low}" not in lib_enums():
            self.err(f"method or data member not found: {e.name}", e)
        if low in self.c.consts and low not in lib_enums() and low not in lib_const_names(stem) and \
                lib in ('vba', 'vbrun', 'vb', 'stdole'):
            self.err(f"method or data member not found: {e.name}", e)

    def expr(self, e, ltr=False):
        hit = self.sched.get(id(e))
        if hit is not None and hit[0] is e:      # a call already evaluated into a temporary
            if not self.dry:
                hit[2] += 1
            return hit[1]
        if self.xdepth == 0 and self.sched_ok and not self.dry:
            r = self.scheduled(e, ltr)
            if r is not None:
                return r
        if e.kind in ('member', 'apply'):
            e = self.unqual(e)
        k = e.kind
        m = getattr(self, 'x_' + k, None)
        if m is None:
            self.err(f"unsupported expression ({k})", e)
        self.xdepth += 1
        try:
            return m(e)
        finally:
            self.xdepth -= 1

    # ---------------------------------------------------------------- order of evaluation (evalorder.py)
    ORDER_TYPES = ('Integer', 'Long', 'Byte', 'Boolean', 'Single', 'Double', 'Currency', 'Date', 'Variant', 'String')
    # functions of VBA whose calls are events: they do something, or read what other calls change
    SIDE_FX = {'rnd', 'inputbox', 'msgbox', 'dir', 'dir$', 'eof', 'lof', 'loc', 'seek', 'input', 'input$', 'inputb',
               'inputb$', 'freefile', 'shell', 'doevents', 'createobject', 'getobject', 'loadpicture', 'timer',
               'now', 'getsetting', 'getallsettings', 'callbyname', 'curdir', 'curdir$', 'filelen', 'filedatetime',
               'getattr', 'err', 'erl', 'environ', 'environ$', 'loadresstring', 'loadresdata', 'loadrespicture'}
    ORDER_OPS = {'=': '=', '<>': '=', '<': '=', '>': '=', '<=': '=', '>=': '=', 'LIKE': 'Like', 'IS': 'Like',
                 'AND': 'And', 'OR': 'Or', 'XOR': 'Xor', 'EQV': 'Xor', 'IMP': 'Xor', 'MOD': 'Mod'}
    CONST_KINDS = ('num', 'str', 'date', 'bool', 'empty', 'null', 'nothing', 'pre', 'addressof', 'filenum')

    def order_model(self):
        mode, noovf = self.c.order[0], self.c.order[1]
        return OrderModel('adv' if mode == 'fast' and noovf else mode, self.handler, self.otype)

    def otype(self, node):
        if node is None:                         # a missing argument
            return 'Variant'
        t = self.dry_type(node)
        return t if t in self.ORDER_TYPES else 'Variant'

    def snapshot(self):
        lists = [self.pre, self.post, self.lines, self.decls, self.cleanup, self.prologue, self.c.protos,
                 self.c.class_support, self.c.static_cnames, self.c.global_cnames]
        # the AddressOf thunks too: one made while trying keeps its name in addressof but loses its prototype
        # (protos), and the real compile then finds it made and leaves the C without the prototype
        return [(x, len(x)) for x in lists] + [(self.locals, dict(self.locals)),
                                               (self.c.addressof, dict(self.c.addressof))]

    @staticmethod
    def rollback(snap):
        for x, n in snap:
            if isinstance(x, dict):              # (an implicit variable made while trying)
                x.clear()
                x.update(n)
            else:
                del x[n:]

    def dry_type(self, node):
        """the type of an expression, compiled without keeping its code"""
        hit = self._dtc.get(id(node))
        if hit is not None and hit[0] is node:
            return hit[1]
        snap, depth = self.snapshot(), self.xdepth
        self.dry += 1
        self.xdepth = 1
        try:
            t = self.expr(node).t
        except Exception:
            t = None
        finally:
            self.rollback(snap)
            self.dry -= 1
            self.xdepth = depth
        self._dtc[id(node)] = (node, t)
        return t

    def _leaf(self, node, evs, method):
        evs.append(node)
        return ('leaf', len(evs) - 1, node, method)

    def _lookup_quiet(self, nm):
        try:
            return self.lookup(nm.name, nm, create=False)
        except VBError:
            return None

    def ev_tree(self, e, evs):
        """the tree of evalorder.py for an expression: its calls that are events go to evs"""
        if e is None:
            return ('const', None)
        if id(e) in self.sched:                  # evaluated already (a part of the statement)
            return ('var', e)
        k = e.kind
        if k == 'named':
            return self.ev_tree(e.e, evs)
        if k == 'paren':
            return self.ev_tree(e.e, evs)
        if k == 'un':
            return ('not' if e.op == 'NOT' else 'neg', self.ev_tree(e.e, evs))
        if k == 'bin':
            return ('bin', self.ORDER_OPS.get(e.op, e.op), self.ev_tree(e.l, evs), self.ev_tree(e.r, evs))
        if k in self.CONST_KINDS:
            return ('const', e)
        if k == 'name':
            return self.ev_name(e, e, evs)
        if k == 'member':
            u = self.unqual(e)
            if u.kind == 'name':
                return self.ev_name(e, u, evs)
            if u.base.kind == 'withref':
                return ('var', e) if self.withs and is_udt(self.withs[-1].t) else self._leaf(e, evs, True)
            if self.ev_count(u.base):
                return self._leaf(e, evs, True)     # a.b().c: the whole chain is one event
            if u.base.kind == 'name' and self._lookup_quiet(u.base) is None:
                return ('var', e)                # App.Path, Err.Number, an Enum's member
            bt = self.dry_type(u.base)
            if bt is None or is_udt(bt):         # Enum.Member, a field of a Type
                return ('var', e)
            return self._leaf(e, evs, True)
        if k == 'apply':
            return self.ev_apply(e, evs)
        if k == 'withref':
            return ('var', e) if self.withs and is_udt(self.withs[-1].t) else self._leaf(e, evs, True)
        if k == 'new':
            return self._leaf(e, evs, False)
        if k == 'typeof':
            return ('call', 'typeof', [self.ev_tree(e.e, evs)], e, None, False)
        return ('var', e)

    def ev_count(self, e):
        evs = []
        self.ev_tree(e, evs)
        return len(evs)

    def ev_name(self, e, nm, evs):
        sy = self._lookup_quiet(nm)
        if sy is None:
            low = nm.name.lower() + (nm.suffix or '' if nm.suffix == '$' else '')
            if B.lookup(nm.name, nm.suffix) is not None:
                return self._leaf(e, evs, False) if low in self.SIDE_FX or nm.name.lower() in self.SIDE_FX \
                    else ('var', e)
            if self.form_builtin(nm.name, nm):
                return self._leaf(e, evs, True)
            return ('var', e)
        if sy.kind in ('proc', 'prop'):
            return self._leaf(e, evs, False)
        if sy.kind == 'const':
            return ('const', e)
        return ('var', e)

    def ev_apply(self, e, evs):
        u = self.unqual(e)
        base = u.base
        args = list(u.args or [])
        def call(name, ev, method):
            ts = [self.ev_tree(a, evs) for a in args]
            return ('call', name, ts, e, self._leaf(e, evs, method)[1] if ev else None, method)
        if base.kind == 'name':
            sy = self._lookup_quiet(base)
            if sy is not None and sy.kind == 'var' and sy.array:
                return call('()', False, False) if args else ('var', e)
            if sy is not None and sy.kind == 'var' and sy.storage == 'ret':
                return call(base.name.lower(), True, False)
            if sy is not None and sy.kind in ('proc', 'prop'):
                return call(base.name.lower(), True, False)
            if sy is not None and sy.kind in ('var', 'me', 'formdefault') and is_obj(sy.t):
                return call(base.name.lower(), True, True)
            if sy is not None:                   # a Variant holding an array, a control array, ...
                return call('()', False, False)
            low = base.name.lower()
            if B.lookup(base.name, base.suffix) is not None:
                if low in ('abs', 'cdbl', 'clng', 'cint', 'cdate', 'len') and len(args) == 1:
                    return call(low, False, False)
                return call(low, low in self.SIDE_FX or low + '$' in self.SIDE_FX, False)
            if self.form_builtin(base.name, base) or self.comp_global(base.name) is not None:
                return call(low, True, True)
            return ('var', e)
        if base.kind == 'member':
            if self.ev_count(base.base) if base.base.kind != 'withref' else False:
                return self._leaf(e, evs, True)
            bb = base.base
            if bb.kind == 'withref':
                if self.withs and is_udt(self.withs[-1].t):
                    return call('()', False, False)
                return call(base.name.lower(), True, True)
            if bb.kind == 'name' and self._lookup_quiet(bb) is None:
                return call(base.name.lower(), True, True)  # Printer.TextWidth(s), Clipboard.GetText()
            bt = self.dry_type(bb)
            if bt is not None and is_udt(bt):    # t.Arr(i): an element of an array in a Type
                return call('()', False, False)
            return call(base.name.lower(), True, True)
        return self._leaf(e, evs, True)

    def hoistable(self, x):
        return (x.code is not None and x.arr is None and not is_udt(x.t) and not is_arr(x.t) and x.t != 'void'
                and known_type(x.t))

    def hoist_events(self, nodes, order, keys):
        """evaluate the calls nodes[i], i in order, into temporaries; False: one cannot be"""
        for i in order:
            node = nodes[i]
            p0 = len(self.pre)
            self.xdepth += 1
            try:
                x = self.expr(node)
            finally:
                self.xdepth -= 1
            if not self.hoistable(x):
                return False
            t = self.tmp(x.t)
            self.pre.append(f"{t} = {x.code};")
            # an error in a call ends the statement: VB6 does not make the calls after it
            self.pre[p0:] = ["if (!vb_errpend) { " + " ".join(self.pre[p0:]) + " }"]
            self.sched[id(node)] = [node, E(t, x.t), 0]
            keys.append(id(node))
        return True

    def scheduled(self, e, ltr=False):
        """an expression with two calls or more: its calls are evaluated first, into temporaries, in the
        order of VB6 (evalorder.py), then the expression with them; None: compile it as it is"""
        evs = []
        tree = self.ev_tree(e, evs)
        if len(evs) < 2:
            return None
        self._dtc = {}
        m = self.order_model()
        order = m.leaves(tree) if ltr else m.events(tree)
        snap, keys = self.snapshot(), []
        try:
            if self.hoist_events(evs, order, keys):
                self.xdepth += 1
                try:
                    r = self.expr(e)
                finally:
                    self.xdepth -= 1
                if all(self.sched[k][2] for k in keys):
                    return r
            self.rollback(snap)                  # a call the code takes another way: as it is
            return None
        except BaseException:
            self.rollback(snap)
            raise
        finally:
            for k in keys:
                self.sched.pop(k, None)
            self._dtc = {}

    def stmt_hoist(self, e, keys, ltr=False, args=None):
        """evaluate the calls of e (or of the arguments args, as those of a call) into temporaries now, for the
        rest of the statement"""
        evs = []
        tree = ('call', '()', [self.ev_tree(a, evs) for a in args], None, None, False) if args is not None \
            else self.ev_tree(e, evs)
        if not evs:
            return True
        self._dtc = {}
        m = self.order_model()
        ok = self.hoist_events(evs, m.leaves(tree) if ltr else m.events(tree), keys)
        self._dtc = {}
        return ok

    def array_reads(self, e, out):
        """the elements of arrays read in e, a(i): VB6 native code computes their indices first"""
        if e is None or not hasattr(e, 'kind'):
            return out
        if e.kind == 'apply' and e.base.kind == 'name' and e.args:
            sy = self._lookup_quiet(e.base)
            if sy is not None and sy.kind == 'var' and sy.array:
                out.append(e)
        for k, v in e.__dict__.items():
            if isinstance(v, N):
                self.array_reads(v, out)
            elif isinstance(v, list):
                for x in v:
                    if isinstance(x, N):
                        self.array_reads(x, out)
        return out

    def x_num(self, e):
        return E(self.c.const_code(e.value, e.type), e.type, const=(e.value, e.type))

    def x_str(self, e):
        return E(self.c.strlit(e.value), 'String', const=(e.value, 'String'))

    def x_bool(self, e):
        return E(self.c.const_code(-1 if e.value else 0, 'Boolean'), 'Boolean', const=(-1 if e.value else 0, 'Boolean'))

    def x_date(self, e):
        v = date_literal(e.text)
        if v is None:
            self.err(f"invalid date literal #{e.text}#", e)
        return E(self.c.const_code(v, 'Date'), 'Date')

    def x_empty(self, e):
        return E('vb_v_empty()', 'Variant')

    def x_null(self, e):
        return E('vb_v_null()', 'Variant')

    def x_filenum(self, e):
        return self.expr(e.e)                     # #1 as an argument: the file number

    def x_paren(self, e):
        r = self.expr(e.e)
        return E(r.code, r.t, const=r.const, ext=r.ext)   # no pointer: (x) is a value

    def var_ref(self, s):
        """E for a scalar variable symbol"""
        if s.kind == 'me':
            return E('vb_self', s.t)
        if s.kind == 'ctl':
            return E(f'vbf_ctl(vb_self, "{c_utf8(s.name.lower())}")', 'Object')
        if s.kind == 'formdefault':
            k = cident(s.key)
            fn = 'vb_odefault' if self.c.classes[s.key].predeclared else 'vbf_default'
            return E(f"{fn}(&G_FORM_{k}, &CLS_{k})", s.t, ptr=f"&G_FORM_{k}")
        if getattr(s, 'storage', None) == 'field':
            code = f"((C_{cident(s.cls)} *)vb_self)->{s.cname}"
            ptr = '&' + code
            if getattr(s, 'withevents', False):
                return E(code, s.t, ptr=ptr, wev=(s.htable, 'vb_self'))
        elif s.byref:
            code, ptr = f"(*{s.cname})", s.cname
        else:
            code, ptr = s.cname, f"&{s.cname}"
        if getattr(s, "autonew", None) and not s.array:
            return E(f"vb_oauto({ptr}, {self.c.cls_ref(s.autonew)})", s.t, ptr=ptr)
        return E(code, s.t, ptr=ptr, fix=getattr(s, 'fixlen', None))

    def arr_ref(self, s):
        """C pointer to the vbarr of an array variable symbol"""
        if getattr(s, 'storage', None) == 'field':
            return f"&((C_{cident(s.cls)} *)vb_self)->{s.cname}"
        return s.cname if s.byref else f"&{s.cname}"

    def array_temp(self, x, et, node):
        """a temporary array (pointer) with the value of an expression: an array function result or a
        Variant holding an array, for an array parameter"""
        self.ntmp += 1
        a = f"t{self.ntmp}"
        self.decls.append(f"vbarr {a} = {{0}};")
        if x.arr is not None and x.code is None:
            xt = x.t[2:] if is_arr(x.t) else x.t
            if xt != et:
                self.err("type mismatch: arrays of different types", node)
            self.pre.append(f"vb_acopy(&{a}, {x.arr});")
        elif x.t == 'Variant':
            self.pre.append(f"vb_v2arr({x.code}, &{a}, {vt_of(et)});")
        elif x.t == 'String' and et == 'Byte':
            self.pre.append(f"vb_s2barr({x.code}, &{a});")
        else:
            self.err("expected array", node)
        self.cleanup.append(f"vb_afree(&{a});")
        return f"&{a}"

    def locked_array(self, a):
        """the vbarr of a(i) passed ByRef when a is a dynamic array variable (else None)"""
        if a.kind != 'apply' or not a.args or a.base.kind != 'name':
            return None
        sy = self.lookup_quiet(a.base.name) if hasattr(self, 'lookup_quiet') else None
        if sy is None:
            try:
                sy = self.lookup(a.base.name, a.base, create=False)
            except Exception:
                return None
        if sy is None or sy.kind != 'var' or not sy.array or getattr(sy, 'fixed_dims', None) or self.is_fixed_local(sy):
            return None
        return self.arr_ref(sy)

    def array_target(self, e, node):
        """(C pointer to vbarr, element type, dims or None) of an array expression: a name or an
        array inside a Type (p.Scores)"""
        self.target_fixlen = None                # (String * n elements: the length, for ReDim)
        if e.kind == 'name':
            sy = self.lookup(e.name, e)
            if sy is None or sy.kind != 'var' or not sy.array:
                self.err(f"expected array: {e.name}", node)
            self.target_fixlen = getattr(sy, 'fixlen', None)
            return self.arr_ref(sy), sy.t, sy.ndims
        if e.kind == 'member':
            x = self.member(e, None, e)
            if x.arr is not None:
                self.target_fixlen = x.fix
                return x.arr, x.t, None
        if e.kind == 'apply' and not e.args:     # Erase a()
            return self.array_target(e.base, node)
        self.err("expected array", node)

    def x_name(self, e):
        s = self.lookup(e.name, e)
        if s is None:
            fm = self.form_builtin(e.name, e) if not e.suffix else None
            if fm:                                   # Left inside a form is Me.Left, not the function
                return self.late(E('vb_self', 'O:' + self.cls), fm, [], e, 'VBM_GET')
            b = B.lookup(e.name, e.suffix)
            if b:
                return b(self, e, [])
            if e.name.lower() in ('err',):
                return E('vb_err_number()', 'Long')
            if e.name.lower() == 'app':              # ShowAbout App, Me: App as an object
                return E(self.c.app_object(), 'Object')
            if e.name.lower() == 'global' and not e.suffix:     # Set g = VB.Global
                return E('vb_global_obj()', 'Object')
            if e.name.lower() == 'dbengine' and 'dao' in self.c.typelibs:
                return E('vb_dao_dbengine()', 'O:dao.dbengine')
            if e.name.lower() == 'rdoengine' and 'rdo' in self.c.typelibs:
                return E('vb_rdo_engine()', 'O:rdo.engine')
            g = self.comp_global(e.name)
            if g is not None:
                return self.late(g, e.name.lower(), [], e, 'VBM_GET')
            fm = self.form_builtin(e.name, e)
            if fm:
                return self.late(E('vb_self', 'O:' + self.cls), fm, [], e, 'VBM_GET')
            s = self.implicit_var(e.name, e.suffix, e)
        if s.kind == 'const':
            if s.t == 'String' and s.value == '' and s.name.lower() == 'vbnullstring':
                return E("((vbs){ 0, 0, 0 })", 'String', const=('', 'String'))   # a null pointer (StrPtr = 0)
            return E(self.c.const_code(s.value, s.t), s.t, const=(s.value, s.t))
        if s.kind in ('me', 'ctl', 'formdefault'):
            return self.var_ref(s)
        if s.kind == 'proc':
            if s.pkind == 'sub':
                self.err(f"expected function or variable: {s.name} is a Sub", e)
            return self.call_proc(s, [], e)
        if s.kind == 'prop':
            return self.prop_get(s, [], e)
        if s.kind == 'var':
            self.check_suffix(s, e)
            if s.array:
                return E(None, 'A:' + s.t, arr=self.arr_ref(s))     # the whole array (assignment, Join ...)
            return self.var_ref(s)
        self.err(f"invalid use of {s.name}", e)

    FORM_BUILTINS = {'caption', 'show', 'hide', 'visible', 'enabled', 'left', 'top', 'width', 'height', 'scalewidth',
                     'scaleheight', 'controls', 'name', 'tag', 'refresh', 'cls', 'move', 'setfocus', 'zorder', 'hwnd',
                     'windowstate', 'backcolor', 'forecolor', 'activecontrol', 'mousepointer', 'icon', 'keypreview',
                     'font', 'fontname', 'fontsize', 'fontbold', 'fontitalic', 'fontunderline', 'fontstrikethru',
                     'currentx', 'currenty', 'scalemode', 'scaleleft', 'scaletop', 'point', 'textwidth', 'textheight',
                     'drawwidth', 'drawstyle', 'drawmode', 'fillstyle', 'fillcolor', 'autoredraw', 'fonttransparent',
                     'image', 'picture', 'borderstyle', 'popupmenu', 'printform', 'scalex', 'scaley', 'appearance',
                     'activeform', 'arrange', 'autoshowchildren', 'mdichild', 'minbutton', 'maxbutton', 'controlbox',
                     'showintaskbar', 'moveable', 'startupposition', 'paint', 'icon', 'whatsthisbutton',
                     'validatecontrols', 'mousepointer', 'drag', 'dragmode', 'hdc', 'paintpicture', 'mouseicon', 'count'}
    UCTL_BUILTINS = {'ambient', 'extender', 'propertychanged', 'canpropertychange', 'size', 'parent',
                     'parentcontrols', 'backstyle', 'hasdc', 'clipcontrols', 'accesskeys', 'eventsfrozen',
                     'containedcontrols', 'canGetFocus', 'forwardfocus', 'controlcontainer', 'maskcolor',
                     'maskpicture', 'palette', 'palettemode', 'toolboxbitmap', 'windowless', 'invisibleatruntime',
                     'hyperlink', 'righttoleft', 'scalemode', 'olemousepointer'}

    def form_builtin(self, name, node=None):
        """an unqualified built-in member of the form inside its own code (Caption = "x", Hide)"""
        if getattr(node, 'lib_qual', False):     # VBA.Left is the function, not Me.Left
            return None
        if self.cls and self.c.classes[self.cls].is_form and name.lower() in self.FORM_BUILTINS:
            return name.lower()
        if self.cls and self.c.classes[self.cls].is_uctl and name.lower() in {b.lower() for b in self.UCTL_BUILTINS}:
            return name.lower()
        if self.cls and name.lower() in ('datamembers', 'datamemberchanged'):   # a class that is a data source
            ci = self.c.classes[self.cls]
            if not ci.is_form and not ci.is_uctl and 'class_getdatamember' in ci.members:
                return name.lower()
        return None

    def prop_get(self, prop, args, node, self_code=None):
        if prop.get is None:
            self.err(f"invalid use of property: {prop.name} has no Property Get", node)
        return self.call_proc(prop.get, args, node, self_code=self_code)

    def check_suffix(self, s, e):
        from .lexer import SUFFIX_TYPE
        if e.suffix and SUFFIX_TYPE[e.suffix] != s.t:
            self.err(f"type-declaration character does not match declared data type: {e.name}{e.suffix}", e)

    def order_args(self, args):
        """VB evaluates the arguments of a call from left to right, C in any order: with two or more
        arguments that call something, all but the last of them are evaluated first, into temporaries"""
        if not args:
            return args
        exprs = [(a.e if a is not None and a.kind == 'named' else a) for a in args]
        imp = [i for i, x in enumerate(exprs) if x is not None and not self.pure(x)]
        if not imp or imp == [len(args) - 1]:
            return args
        out = list(args)
        for i in (imp if imp[-1] != len(args) - 1 else imp[:-1]):
            x = self.expr(exprs[i], ltr=True)     # (an argument but the last: its calls from left to right)
            if (x.code is not None and x.arr is None and not is_udt(x.t) and not is_arr(x.t) and x.t != 'void'
                    and known_type(x.t) and x.ptr is None):     # (not a variable: it goes ByRef)
                t = self.tmp(x.t)
                self.pre.append(f"{t} = {self.conv(x, x.t)};")
                x = E(t, x.t)
            p = _Pre(x)
            p.orig = exprs[i]                     # (what it was: ByVal Text1.Text for a Declare, ctl_prop_type)
            if args[i].kind in ('paren', 'filenum'):     # ByVal x, #1: the mark stays
                p = N(args[i].kind, args[i].line, e=p)
            out[i] = N('named', args[i].line, name=args[i].name, e=p) if args[i].kind == 'named' else p
        return out

    def x_apply(self, e):
        base = e.base
        if e.args:
            e = N('apply', e.line, base=base, args=self.order_args(e.args))
        if base.kind == 'name':
            s = self.lookup(base.name, base)
            if s is not None and s.kind == 'var' and s.array:
                if not e.args:                       # a(): the whole array (UBound(a()), f(a()) ...)
                    return self.expr(base)
                return self.array_elem(s, e.args, e)
            if s is not None and s.kind == 'var' and s.storage == 'ret':
                # recursive call: the function's own name with arguments
                return self.call_proc(self.sym, e.args, e)
            if s is not None and s.kind == 'proc':
                if s.pkind == 'sub':
                    self.err(f"expected function or variable: {s.name} is a Sub", e)
                return self.call_proc(s, e.args, e)
            if s is not None and s.kind == 'prop':
                if e.args and s.get is not None and not s.get.params and (
                        is_obj(s.get.rt or '') or s.get.rt in ('Object', 'Variant')):
                    # Menus(i) with Property Get Menus() As Collection: the default member of what it gives
                    return self.default_member(self.prop_get(s, [], e), e.args, e)
                return self.prop_get(s, e.args, e)
            if s is not None and s.kind in ('var', 'me', 'ctl', 'formdefault') and (is_obj(s.t) or s.t == 'Variant'):
                return self.default_member(self.var_ref(s), e.args, e)
            if s is None:
                b = B.lookup(base.name, base.suffix)
                if b:
                    return b(self, base, e.args)
                fm = self.form_builtin(base.name, base)
                if fm:
                    return self.late(E('vb_self', 'O:' + self.cls), fm, e.args, e, 'VBM_GET')
                g = self.comp_global(base.name)
                if g is not None:
                    return self.late(g, base.name.lower(), e.args, e, 'VBM_GET')
                if not self.explicit and base.name.lower() not in B.KNOWN_VBA:
                    self.err(f"Sub, Function or array not defined: {base.name}", e)
                self.err(f"{base.name} is not supported yet" if base.name.lower() in B.KNOWN_VBA
                         else f"Sub or Function not defined: {base.name}", e)
            self.err(f"expected array or function: {base.name}", e)
        if base.kind == 'member':
            return self.member(base, e.args, e)
        if base.kind in ('apply', 'withref'):    # f(1)(2), !Field in With: default member of the object
            x = self.expr(base)
            if x.code is None and x.arr is not None and e.args:   # Split(s, ",")(2): an element of the result
                return self.elem_ref(x.arr, x.t[2:] if is_arr(x.t) else x.t, None, e.args, e)
            return self.default_member(x, e.args, e)
        self.err("expected: function, array or method", e)

    def default_needs_args(self, t):
        """the default member of class t takes an argument that is not Optional (Item of a Collection)"""
        if t in ('O:collection', 'O:scripting.dictionary'):
            return True
        ci = self.c.classes.get(t[2:]) if t.startswith('O:') else None
        if ci is None or not ci.defmember:
            return False
        m = ci.members.get(ci.defmember.lower())
        if m is None:
            return False
        procs = [m] if m.kind == 'proc' else [p for p in (getattr(m, 'let', None), getattr(m, 'get', None)) if p]
        if not procs:
            return False
        p = procs[0]
        params = p.params[:-1] if getattr(p, 'pkind', '') in ('property let', 'property set') else p.params
        return any(not getattr(a, 'optional', False) and not getattr(a, 'paramarray', False) for a in params)

    def default_member(self, obj, args, node, kind='VBM_GET', value=None):
        """obj(args): the default member (Item of a Collection, VB_UserMemId = 0 of a class)"""
        t = obj.t
        if obj.code is None and obj.arr is not None and args and kind in ('VBM_GET', 'VBM_METHOD'):
            return self.elem_ref(obj.arr, t[2:] if is_arr(t) else t, None, args, node)   # f()(i): an element
        if t not in ('Object', 'Variant') and t.startswith('O:') and t[2:] in self.c.classes:
            ci = self.c.classes[t[2:]]
            if not ci.defmember and ci.is_form and args:   # Form2!Combo1, Form2("Text1"): a control of the form
                ctl = self.late(obj, 'controls', args, node, 'VBM_GET')
                if kind in ('VBM_GET', 'VBM_METHOD'):
                    return ctl
                return self.late(ctl, '', [], node, kind, value)
            if not ci.defmember:
                self.err(f"{ci.name} has no default member", node)
            return self.member_of(obj, ci.defmember, args, node, kind, value)
        if t == 'Variant':                      # v(i): an element of an array in the Variant (or a default member)
            vals = ['vb_v_missing()' if a is None else self.conv(self.expr(a), 'Variant', a) for a in (args or [])]
            argv = f"{len(vals)}, (vbvar[]){{ {', '.join(vals)} }}" if vals else "0, 0"
            if kind in ('VBM_GET', 'VBM_METHOD'):
                return E(f"vb_v_index({obj.code}, {argv})", 'Variant')
            if obj.ptr is None:                  # !Field = v in With .Recordset: the object's default member
                return self.late(obj, '', args, node, kind, value)
            return E(f"vb_v_setindex({obj.ptr}, {self.conv(value, 'Variant', node)}, {argv}, "
                     f"{1 if kind == 'VBM_SET' else 0})", 'void')
        return self.late(obj, '', args, node, kind, value)

    def x_member(self, e):
        return self.member(e, None, e)

    def with_target(self, node):
        if not self.withs:
            self.err("invalid or unqualified reference (. outside With)", node)
        return self.withs[-1]

    def member_base(self, m):
        if m.base.kind == 'withref':
            return self.with_target(m)
        return self.expr(m.base)

    def comp_global(self, name, proj=None):
        """a member of a GlobalMultiUse class of a component, used without an object (Twice(2) or
        Lib.Twice(2)): the expression of that class's one object, or None"""
        low = name.lower()
        if proj is None and 'dao' in self.c.typelibs and low in DAO_GLOBALS and not self.c.classes.get(low):
            return E('vb_dao_dbengine()', 'O:dao.dbengine')
        for comp in self.c.components:
            if proj is not None and comp.name.lower() != proj.lower():
                continue
            for ci in comp.classes.values():
                if getattr(ci, 'global_ns', False) and low in ci.members:
                    return E(f'vb_component_global("{c_utf8(comp.name.lower())}.{c_utf8(ci.name.lower())}")', 'Object')
        return None

    # members of the Extender of a control (Counter1.Left ...): not the user control's own
    EXTENDER = {'name', 'index', 'tag', 'parent', 'left', 'top', 'width', 'height', 'visible', 'move', 'zorder',
                'setfocus', 'tabindex', 'tabstop', 'container', 'tooltiptext', 'dragicon', 'dragmode', 'drag',
                'helpcontextid', 'whatsthishelpid', 'cancel', 'default', 'causesvalidation', 'object',
                'datafield', 'datamember', 'datasource', 'datachanged', 'dataformat', 'showwhatsthis'}

    def static_class(self, base):
        """(ClassInfo, is a control) of an expression whose class VB6 knows when it compiles, though it
        is late bound in C: a user control on the form, a variable As Lib.Class / As UserCtl"""
        name = base.base if base.kind == 'apply' else base
        if name.kind != 'name':
            return None
        sy = self.lookup(name.name, name)
        if sy is None:
            return None
        if sy.kind == 'ctl' and self.cls:
            cls = self.c.classes[self.cls].ctl_cls.get(name.name.lower())
            if cls and cls.lower() in getattr(self.c, 'uctl_names', ()):
                proj, _, short = cls.lower().rpartition('.')
                ci = self.c.classes.get(short) if proj == self.c.project_name.lower() else None
                if ci is None:
                    for comp in self.c.components:
                        if comp.name.lower() == proj:
                            ci = comp.classes.get(short)
                return (ci, True) if ci is not None else None
            return None
        if sy.kind == 'var' and base.kind == 'name' and getattr(sy, 'ext', None) is not None:
            return sy.ext, False
        return None

    def ctl_prop_type(self, node):
        """Text1.Text, Me.Text1.Text, Form1.Label1.Caption: the type of the property of an intrinsic control (VB
        library), or None"""
        if node is not None and node.kind == 'pre':
            node = getattr(node, 'orig', None)
        if node is not None and node.kind == 'paren':
            node = node.e
        if node is None or node.kind != 'member':
            return None
        b = node.base
        if b.kind == 'member' and b.base.kind in ('name', 'me'):      # Me.Text1 / Form1.Text1
            owner = self.cls if b.base.kind == 'me' or b.base.name.lower() == 'me' else b.base.name.lower()
            ci, cname = self.c.classes.get(owner), b.name.lower()
        elif b.kind == 'name':
            sy = self._lookup_quiet(b)
            if sy is None or sy.kind != 'ctl' or not self.cls:
                return None
            ci, cname = self.c.classes.get(self.cls), b.name.lower()
        else:
            return None
        if ci is None or not getattr(ci, 'is_form', False):
            return None
        cls = (getattr(ci, 'ctl_cls', {}) or {}).get(cname)
        if not cls or not cls.lower().startswith('vb.'):
            return None
        from .symbols import vb_prop_type
        return vb_prop_type(cls[3:], node.name)

    def check_member(self, m, node):
        """VB6: Method or data member not found, for a class it binds early"""
        st = self.static_class(m.base)
        if st is None:
            return
        ci, is_ctl = st
        low = m.name.lower()
        sy = ci.members.get(low)
        if sy is not None and getattr(sy, 'public', False) and sy.kind in ('var', 'proc', 'prop'):
            return
        if is_ctl and (low in self.EXTENDER or low in ('count', 'ubound', 'lbound', 'item')):
            return                               # (Count ...: of a control array of user controls)
        self.err(f"method or data member not found: {m.name}", node)

    def member(self, m, args, node, kind='VBM_GET', value=None):
        self.check_member(m, node)
        if m.base.kind == 'name' and m.base.name.lower() == 'usercontrol' and self.cls \
                and self.c.classes[self.cls].is_uctl:
            sy = self.lookup(m.base.name, m.base)
            if sy is not None and sy.kind == 'me':   # UserControl.Enabled: the built-in member, not the class's own
                r = self.late(E('vb_self', 'Object'), m.name.lower(), args, node, kind, value)
                r.code = r.code.replace('vb_late_o(', 'vbf_uc_member(', 1)
                return r
        if m.base.kind == 'name' and m.base.name.lower() == 'error' and not m.base.suffix \
                and self.lookup(m.base.name, m.base) is None and self.comp_global(m.name, m.base.name) is None:
            # Error.Number: Error() gives a Variant (a String), so VB6 makes the program and the member is
            # error 424 (Object required) when it runs (compile-6.8 c16)
            return E("(vb_raise(424), vb_v_empty())", 'Variant')
        if m.base.kind == 'name' and self.lookup(m.base.name, m.base) is None:
            g = self.comp_global(m.name, m.base.name)
            if g is not None:
                return self.late(g, m.name.lower(), args, node, kind, value)
        if m.base.kind == 'name' and kind == 'VBM_GET' and not args:
            # EnumName.Member: the constant (an Enum of the program or of a type library)
            sy = self.lookup(m.base.name, m.base)
            bl = m.base.name.lower()
            is_enum = (sy is not None and sy.kind == 'enum') or (sy is None and (
                (self.mk, bl) in self.c.enum_names or ('', bl) in self.c.enum_names or bl in self.c.comp_enums))
            if is_enum:
                cv = self.c.lookup_const(m.name, self.mk, None)
                if cv is not None:
                    return E(self.c.const_code(cv[0], cv[1]), cv[1], const=cv)
        if m.base.kind == 'name' and kind == 'VBM_GET':
            obj = m.base.name.lower()
            if self.lookup(m.base.name, m.base) is None:
                h = B.lookup_member(obj, m.name.lower())
                if h:
                    return h(self, node, args or [])
                if obj == 'err' and m.name.lower() not in ('number', 'description', 'source', 'helpfile', 'helpcontext',
                                                           'lastdllerror', 'clear', 'raise'):
                    # Err.numer: VB6 compiles it (the Err object is late bound) and it is error 438 when it runs
                    return E("(vb_raise(438), vb_v_empty())", 'Variant')
                if obj in ('err', 'app', 'debug'):
                    self.err(f"{self.describe(m)} is not supported yet", node)
        return self.member_of(self.member_base(m), m.name, args, node, kind, value)

    def member_of(self, b, name, args, node, kind='VBM_GET', value=None):
        """b.name(args): read (VBM_GET / VBM_METHOD), or assign value (VBM_LET / VBM_SET)"""
        t = b.t
        low = name.lower()
        if is_udt(t):
            it = self.c.udts[t[2:]]
            f = next((f for f in it.fields if f.name.lower() == low), None)
            if f is None:
                self.err(f"method or data member not found: {name}", node)
            if b.ptr is None:
                self.err("invalid use of a Type value (assign it to a variable first)", node)
            code = f"({b.code}).{f.cname}"
            if f.dims is not None:
                ref = f"&{code}"
                if not args:                         # r.field or r.field(): the whole array
                    return E(None, f.t, arr=ref, fix=getattr(f, 'fixlen', None))
                x = self.elem_ref(ref, f.t, len(f.dims) if f.dims else None, args, node)
                if getattr(f, 'fixlen', None):       # an element of name(9) As String * 35: padded, cut
                    x.fix = f.fixlen
                return x
            if args and (f.t == 'Variant' or is_obj(f.t)):   # t.Items(0) = x: the array in a Variant field
                return self.default_member(E(code, f.t, ptr=f"&{code}"), args, node, kind, value)
            if args:
                self.err(f"{name} is not an array", node)
            if getattr(f, 'autonew', None):        # c As New Collection in a Type: made when it is used
                return E(f"vb_oauto(&{code}, {self.c.cls_ref(f.autonew)})", f.t, ptr=f"&{code}")
            return E(code, f.t, ptr=f"&{code}", fix=f.fixlen)
        if is_obj(t) and t != 'Object' and t[2:] in self.c.classes:
            ci = self.c.classes[t[2:]]
            m = ci.members.get(low)
            if m is None and ci.is_form:            # built-in members and controls of a form
                return self.late(b, low, args, node, kind, value)
            if m is None and 'lb_inner' in ci.members:   # a designer (DataReport): the object of the runtime inside
                inner = self.member_of(b, 'lb_inner', None, node)
                return self.late(inner, low, args, node, kind, value)
            if m is None:
                self.err(f"method or data member not found: {name}", node)
            if not getattr(m, 'public', False) and self.cls != t[2:]:
                self.err(f"method or data member not found: {name} (Private)", node)
            if m.kind == 'var':
                code = f"((C_{cident(t[2:])} *)vb_ofields({b.code}, &D_{cident(t[2:])}))->{m.cname}"
                if m.array:
                    if args is None:
                        return E(None, m.t, arr=f"&{code}")
                    return self.elem_ref(f"&{code}", m.t, m.ndims, args, node)
                if args:
                    return self.default_member(E(code, m.t), args, node, kind, value)
                if getattr(m, 'autonew', None):
                    return E(f"vb_oauto(&{code}, {self.c.cls_ref(m.autonew)})", m.t, ptr=f"&{code}", prop=True)
                if getattr(m, 'withevents', False):
                    return E(code, m.t, ptr=f"&{code}", prop=True, wev=(m.htable, b.code))
                return E(code, m.t, ptr=f"&{code}", prop=True, fix=getattr(m, 'fixlen', None))
            disp = lambda p: ci.dispatch.get(p.cname, p) if p is not None else None
            if m.kind == 'prop':
                if kind in ('VBM_GET', 'VBM_METHOD'):
                    if m.get is None:
                        self.err(f"invalid use of property: {m.name} has no Property Get", node)
                    if args and not m.get.params and (is_obj(m.get.rt or '') or m.get.rt == 'Variant'):
                        # obj.Commands(1): the default member; reg.Data(i): an element of the array it gives
                        return self.default_member(self.call_proc(disp(m.get), [], node, self_code=b.code), args, node, kind, value)
                    return self.call_proc(disp(m.get), args or [], node, self_code=b.code)
                p = m.let if kind == 'VBM_LET' else m.set
                if args and m.get is not None and not m.get.params and is_obj(m.get.rt or '') and (
                        p is None or len(p.params) - 1 != len(args)):
                    # x.Attributes("color") = v: the default member of the object Attributes gives
                    got = self.call_proc(disp(m.get), [], node, self_code=b.code)
                    return self.default_member(got, args, node, kind, value)
                if p is None and kind == 'VBM_LET' and m.get is not None and (
                        is_obj(m.get.rt or '') or m.get.rt in ('Object', 'Variant')):
                    # x.Font = "Arial" with Get / Set only: the default member of what Get gives (Font.Name)
                    got = self.call_proc(disp(m.get), args or [], node, self_code=b.code)
                    return self.late(got, '', [], node, 'VBM_LET', value)
                if p is None:
                    self.err(f"invalid use of property: {name} has no Property {'Let' if kind == 'VBM_LET' else 'Set'}",
                             node)
                return self.call_proc(disp(p), list(args or []) + [_Pre(value)], node, as_stmt=True, self_code=b.code)
            if m.kind == 'proc':
                if kind in ('VBM_LET', 'VBM_SET'):
                    self.err(f"function call on left-hand side of assignment: {name}", node)
                if args and not m.params and (is_obj(m.rt or '') or m.rt == 'Variant') and kind != 'VBM_METHOD':
                    return self.default_member(self.call_proc(disp(m), [], node, self_code=b.code), args, node, kind, value)
                return self.call_proc(disp(m), args or [], node, as_stmt=(kind == 'VBM_METHOD'), self_code=b.code)
            self.err(f"invalid use of {name}", node)
        if is_obj(t) or t == 'Variant':
            return self.late(b, low, args, node, kind, value)
        self.err(f"invalid qualifier: {self.describe_t(t)}", node)

    def describe_t(self, t):
        if is_udt(t):
            return self.c.udts[t[2:]].name
        return t[2:] if t.startswith('O:') else t

    def late(self, b, name, args, node, kind, value=None):
        """late-bound call (Object, Variant, Collection)"""
        vals = []
        args = list(args or [])
        if any(a is not None and a.kind == 'named' for a in args):
            args = self.order_named(b.t, name, args, node)
        back = []                                # ByRef: plain variables get the callee's value back
        for i, a in enumerate(args):
            if a is None:
                vals.append('vb_v_missing()')
                continue
            x = self.expr(a)
            vals.append(self.conv(x, 'Variant', a))
            if a.kind == 'name' and x.ptr is not None and not x.prop and not x.fix and x.t != 'Variant' \
                    and not is_udt(x.t) and x.arr is None:
                sy = self.lookup(a.name, a)
                if sy is not None and sy.kind == 'var' and not sy.array:
                    back.append((i, x, a))
            elif a.kind == 'name' and x.ptr is not None and x.t == 'Variant' and not x.prop:
                sy = self.lookup(a.name, a)
                if sy is not None and sy.kind == 'var' and not sy.array:
                    back.append((i, x, a))
            elif a.kind == 'name' and x.ptr is not None and is_udt(x.t) and self.c.rec_ok(x.t) and not x.prop:
                sy = self.lookup(a.name, a)      # XChart1.AddItem item: a Type ByRef, its new value back
                if sy is not None and sy.kind == 'var' and not sy.array:
                    back.append((i, x, a))
            elif a.kind == 'name' and x.arr is not None:
                # a dynamic array: the callee may give another one back (Winsock GetData b, vbArray + vbByte)
                sy = self.lookup(a.name, a)
                if sy is not None and sy.kind == 'var' and sy.array and not getattr(sy, 'fixed_dims', None) \
                        and not self.is_fixed_local(sy):
                    back.append((i, x, a))
        if value is not None:
            vals.append(self.conv(value, 'Variant', node))
        argv = f"{len(vals)}, (vbvar[]){{ {', '.join(vals)} }}" if vals else "0, 0"
        if back:
            self.ntmp += 1
            arr = f"t{self.ntmp}"
            self.decls.append(f"vbvar {arr}[{len(vals)}] = {{ {{ 0 }} }};")
            for i, v in enumerate(vals):
                self.pre.append(f"vb_vset(&{arr}[{i}], {v});")
            for i, x, a in back:
                if x.arr is not None:
                    et = x.t[2:] if is_arr(x.t) else x.t
                    self.post.append(f"if (!vb_errpend && vb_isarray({arr}[{i}])) vb_v2arr({arr}[{i}], {x.arr}, {vt_of(et)});")
                    continue
                self.post.append(f"if (!vb_errpend) {self.store_code(x, E(f'{arr}[{i}]', 'Variant'), a)}")
            for i in range(len(vals)):
                self.post.append(f"vb_vfree(&{arr}[{i}]);")
            argv = f"{len(vals)}, {arr}"
        lname = c_utf8(name)
        if is_obj(b.t):
            return E(f"vb_late_o({b.code}, \"{lname}\", {kind}, {argv})", 'Variant')
        return E(f"vb_late({b.code}, \"{lname}\", {kind}, {argv})", 'Variant')

    # the interface of the objects whose members the type libraries describe
    LATE_IFACES = {'O:dao.recordset': 'recordset', 'O:dao.database': 'database', 'O:dao.workspace': 'workspace',
                   'O:dao.querydef': '_querydef', 'O:dao.tabledef': '_tabledef', 'O:dao.field': '_field',
                   'O:dao.dbengine': '_dbengine', 'O:dao.index': '_index', 'O:dao.relation': '_relation',
                   'O:adodb.recordset': 'recordset15', 'O:adodb.connection': 'connection15',
                   'O:adodb.command': 'command15', 'O:adodb.field': 'field20', 'O:adodb.fields': 'fields',
                   'O:adodb.parameter': '_parameter', 'O:adodb.parameters': 'parameters',
                   'O:rdo.engine': '_rdoengine', 'O:rdo.environment': '_rdoenvironment',
                   'O:rdo.connection': '_rdoconnection', 'O:rdo.query': 'rdopreparedstatement',
                   'O:rdo.resultset': '_rdoresultset', 'O:rdo.column': '_rdocolumn', 'O:rdo.table': 'rdotable',
                   'O:collection': '_collection', 'O:scripting.fso': 'ifilesystem3', 'O:scripting.dictionary': 'idictionary',
                   'O:scripting.textstream': 'itextstream', 'O:scripting.file': 'ifile', 'O:scripting.folder': 'ifolder',
                   'O:scripting.drive': 'idrive', 'O:scripting.files': 'ifilecollection',
                   'O:scripting.folders': 'ifoldercollection', 'O:scripting.drives': 'idrivecollection'}

    def order_named(self, t, name, args, node):
        """Add Item:=x, Key:=k on an object of the runtime: the arguments in the order of the type library"""
        from .symbols import member_params
        names = [a.name.lower() for a in args if a is not None and a.kind == 'named']
        cands = member_params().get(name.lower(), [])
        iface = self.LATE_IFACES.get(t)
        ext = self.ext_class(t[2:].lower()) if t.startswith('P:') else None
        eci = ext[1] if ext is not None else None
        if eci is None:                          # x.Method(A:=1) on x As Lib.Class, a user control: early bound
            tgt = getattr(node, 'target', None) if getattr(node, 'kind', '') == 'call' else node
            if getattr(tgt, 'kind', '') == 'apply':
                tgt = tgt.base
            base = getattr(tgt, 'base', None) if getattr(tgt, 'kind', '') == 'member' else None
            st = self.static_class(base) if base is not None and base.kind in ('name', 'apply') else None
            eci = st[0] if st else None
        if eci is not None:                      # a class of a component: the parameters of its member
            m = eci.members.get(name.lower())
            p = m if m is not None and m.kind == 'proc' else getattr(m, 'get', None) if m is not None else None
            if p is not None:
                ps = [q.name for q in p.params]
                bad = next((n for n in names if n not in [q.lower() for q in ps]), None)
                if bad is not None:
                    self.err(f"named argument not found: {bad}", node)
                cands, iface = [('', ps)], 'component'
        if iface:
            cands = [c for c in cands if c[0] == iface or (iface == 'ifilesystem3' and c[0] == 'ifilesystem')] or cands
        fits = [ps for _, ps in cands if all(n in [p.lower() for p in ps] for n in names)]
        if fits and any([p.lower() for p in f] != [p.lower() for p in fits[0]] for f in fits[1:]) and not iface:
            # x.ListItems.Add Key:=...: the collection's name chooses the interface (IListItems)
            tgt = getattr(node, 'target', None) if getattr(node, 'kind', '') == 'call' else node
            if getattr(tgt, 'kind', '') == 'apply':      # Set n = X.Nodes.Add(Key:=...)
                tgt = tgt.base
            base = getattr(tgt, 'base', None) if getattr(tgt, 'kind', '') == 'member' else None
            if getattr(base, 'kind', '') == 'withref' and getattr(self, 'with_nodes', None):
                base = self.with_nodes[-1]               # With lvw.ListItems: .Add Text:=...
            if getattr(base, 'kind', '') == 'apply':     # .ListItems.Add(...) on With lvwFreq.ListItems.Add(...)
                base = base.base
            hint = (getattr(base, 'name', '') or '').lower() if base is not None else ''
            hint = {'listimages': 'images', 'buttons': 'buttons'}.get(hint, hint)   # ImageList.ListImages: IImages
            if hint:
                pick = [ps for ifc, ps in cands if ifc in (hint, 'i' + hint, '_' + hint, 'i' + hint.rstrip('s'))
                        and all(n in [p.lower() for p in ps] for n in names)]
                if pick and all([p.lower() for p in f] == [p.lower() for p in pick[0]] for f in pick[1:]):
                    fits = pick
        if not fits or any([p.lower() for p in f] != [p.lower() for p in fits[0]] for f in fits[1:]) and not iface:
            self.err(f"named arguments: the parameters of {name} are not known", node)
        order = [p.lower() for p in fits[0]]
        out = []
        for a in args:
            if a is not None and a.kind == 'named':
                break
            out.append(a)
        for a in args[len(out):]:
            if a is None or a.kind != 'named':
                self.err("named arguments must come after the others", node)
            i = order.index(a.name.lower())
            if i < len(out) and out[i] is not None:
                self.err(f"argument {a.name} given twice", node)
            while len(out) <= i:
                out.append(None)
            out[i] = a.e
        return out

    def describe(self, e):
        if e.kind == 'name':
            return e.name
        if e.kind == 'member':
            return self.describe(e.base) + '.' + e.name
        if e.kind == 'withref':
            return ''
        return '...'

    def array_elem(self, s, args, node):
        x = self.elem_ref(self.arr_ref(s), s.t, s.ndims, args, node)
        if getattr(s, 'autonew', None):
            return E(f"vb_oauto({x.ptr}, {self.c.cls_ref(s.autonew)})", s.t, ptr=x.ptr)
        if getattr(s, 'fixlen', None):           # an element of an array of String * n
            x.fix = s.fixlen
        return x

    def elem_ref(self, ref, t, ndims, args, node):
        if any(a is None or a.kind == 'named' for a in args):
            self.err("invalid array index", node)
        if ndims is not None and len(args) != ndims:
            self.err("wrong number of dimensions", node)
        idx = [self.conv(self.expr(a), 'Long', a) for a in args]
        ct = ctype(t)
        nc = '_nc' if self.nobounds else ''      # Remove Array Bounds Checks
        if len(idx) == 1:
            p = f"(({ct} *)vb_aelem1{nc}({ref}, {idx[0]}))"
        else:
            p = f"(({ct} *)vb_aelem{nc}({ref}, {len(idx)}, {', '.join(idx)}))"
        x = E(f"(*{p})", t, ptr=p)
        x.aref = ref                             # (VarPtr(a(0)): the elements after it too)
        return x

    def x_addressof(self, e):
        """AddressOf Proc: a number that stands for the procedure (callbacks: SetTimer ...); the runtime
        calls it through a thunk that takes the arguments as Long values"""
        s = None
        if e.mod:
            s = self.c.mscopes.get(e.mod.lower(), {}).get(e.name.lower())
        else:
            s = self.lookup(e.name, e, create=False)
        if s is None or s.kind != 'proc' or getattr(s, 'cls', None) or getattr(s, 'dll', False):
            self.err(f"AddressOf needs a Sub or Function of a standard module: {e.name}", e)
        th = 'AO_' + s.cname
        if th not in self.c.addressof:
            args, pre, post = [], [], []
            k = 0                                # the slot of argv (a ByVal Currency takes two)
            for p in s.params:
                i = k
                k += 1
                v = f"(argc > {i} ? argv[{i}] : 0)"
                if is_udt(p.t) and not p.array and not p.byval:
                    # ByRef r As RECT: the caller's memory, copied in before the call and back after it
                    u = p.t[2:]
                    pre.append(f"U_{u} a{i} = {{0}}; U_{u}_init(&a{i}); vb_cb_udt_in(&UD_{u}, &a{i}, {v});")
                    post.append(f"vb_cb_udt_out(&UD_{u}, &a{i}, {v}); U_{u}_free(&a{i});")
                    args.append(f"&a{i}")
                    continue
                if (is_obj(p.t) or p.t == 'Object') and not p.array:
                    # ByVal lParam As Object (an ObjPtr): the object itself, without a reference of its own
                    o = f"vb_cb_obj({v})"
                    if p.byval:
                        args.append(o)
                    else:
                        pre.append(f"vbobj a{i} = {o};")
                        args.append(f"&a{i}")
                    continue
                if p.t == 'Currency' and not p.array:
                    k += 1 if p.byval else 0     # a ByVal LARGE_INTEGER: two Longs
                    c = (f"((vb_cy)(((vb_u64)(vb_u32)(argc > {i} ? argv[{i}] : 0)) | "
                         f"((vb_u64)(vb_u32)(argc > {i + 1} ? argv[{i + 1}] : 0) << 32)))") if p.byval else \
                        f"(vb_cy)0"
                    if p.byval:
                        args.append(c)
                    else:
                        pre.append(f"vb_cy a{i} = 0; {{ char *q = (char *)vb_ptr_resolve({v}, 8); if (q) memcpy(&a{i}, q, 8); else vb_errpend = 0; }}")
                        args.append(f"&a{i}")
                    continue
                if p.array or p.t not in ('Long', 'Integer', 'Byte', 'Boolean', 'Single', 'Double', 'Variant', 'String'):
                    self.err(f"AddressOf {s.name}: a parameter of type {p.t} cannot come from a callback", e)
                if p.t == 'Variant':
                    v = f"vb_v_l({v})"
                elif p.t == 'String':
                    v = "vb_stmp(0)"
                else:
                    v = f"({ctype(p.t)}){v}"
                if p.byval:
                    args.append(v)
                else:
                    pre.append(f"{ctype(p.t)} a{i} = {v};")
                    args.append(f"&a{i}")
            call = f"{s.cname}({', '.join(args)})"
            self.c.addressof[th] = (f"static vb_i32 {th}(vb_i32 argc, vb_i32 *argv) {{ vb_i32 m = vb_mark(); "
                                    f"{' '.join(pre)} {self.c.ao_body(call, s, ' '.join(post))} }}")
            self.c.protos.append(f"static vb_i32 {th}(vb_i32 argc, vb_i32 *argv);")
        return E(f"vb_addressof({th})", 'Long')

    def x_nothing(self, e):
        return E('((vbobj)0)', 'Object')

    def x_withref(self, e):
        return self.with_target(e)

    def x_new(self, e):
        if e.cls.lower().startswith(self.c.project_name.lower() + '.'):
            e = N('new', e.line, cls=e.cls[len(self.c.project_name) + 1:])
        ext = self.c.ext_class(e.cls.lower())
        if ext and e.cls.lower() not in self.c.classes:  # New Proj.Class: a class of a component
            return E(f'vb_component_new("{c_utf8(ext[0].name)}.{c_utf8(ext[1].name)}")', 'Object')
        t = self.c.resolve_type(self.mk, e.cls, e, new=True)
        if not t.startswith('O:'):
            self.err(f"invalid use of New keyword: {e.cls}", e)
        return E(f"vb_onew({self.c.cls_ref(t)})", t)

    def x_typeof(self, e):
        x = self.expr(e.e)
        low = e.cls.lower()
        low = low[3:] if low.startswith('vb.') else low
        if low in VB_OBJECT_TYPES and low not in self.c.classes:
            return E(f'vb_typeof_name({self.conv(x, "Object", e)}, "{low}")', 'Boolean')
        t = self.c.resolve_type(self.mk, e.cls, e)
        if not is_obj(t):
            self.err(f"expected: class name ({e.cls})", e)
        return E(f"vb_typeof({self.conv(x, 'Object', e)}, {self.c.cls_ref(t)})", 'Boolean')

    def x_un(self, e):
        a = self.expr(e.e)
        if e.op == 'NOT' and a.arr is not None and a.code is None:
            # Not arr: VB6 gives the complement of the array's SAFEARRAY pointer (-1: not dimensioned yet)
            return E(f"((vb_i32)~(({a.arr})->nd ? vb_varptr({a.arr}, (vb_i32)sizeof(vbarr)) : 0))", 'Long')
        if is_obj(a.t) and a.code is not None:  # -Text1, Not Check1: the value of the default property
            a = E(f'vb_late_o({self.conv(a, "Object", e)}, "", VBM_GET, 0, 0)', 'Variant')
        if e.op == '-':
            t = a.t
            if a.const and t in ('Integer', 'Long', 'Double', 'Single', 'Currency') and \
                    not (t == 'Integer' and a.const[0] == -32768):
                return E(self.c.const_code(-a.const[0], t), t, const=(-a.const[0], t))
            if t == 'Variant':
                return E(f"vb_v_neg({a.code})", 'Variant')
            if t == 'String':
                return E(f"(-vb_s2d({a.code}))", 'Double')
            if a.ext:
                return E(f"(-({a.code}))", t, ext=True)
            if t in ('Integer', 'Boolean', 'Byte', 'Long') and self.native:   # native code: no overflow check
                return self.neg_wrap(E(a.code, 'Long' if t == 'Long' else 'Integer'))
            if t in ('Integer', 'Boolean', 'Byte'):
                return E(f"vb_sub_i(0, {a.code})", 'Integer')
            if t == 'Long':
                return E(f"vb_sub_l(0, {a.code})", 'Long')
            if t == 'Currency':
                return E(f"vb_sub_c(0, {a.code})", 'Currency')
            return E(f"(-({a.code}))", t)
        # Not
        t = a.t
        if t == 'Variant':
            return E(f"vb_v_not({a.code})", 'Variant')
        if t == 'Boolean':
            return E(f"((vb_bool)~({a.code}))", 'Boolean')
        if t in ('Integer', 'Byte'):
            return E(f"((vb_i16)~({a.code}))", 'Integer') if t == 'Integer' else E(f"((vb_u8)~({a.code}))", 'Byte')
        if t == 'Long':
            return E(f"((vb_i32)~({a.code}))", 'Long')
        return E(f"((vb_i32)~{self.conv(a, 'Long', e)})", 'Long')

    def num_type(self, ta, tb):
        """result type of + - * on two numeric types"""
        if 'Date' in (ta, tb):
            return 'Date'
        ra, rb = RANK[ta], RANK[tb]
        t = ta if ra >= rb else tb
        if t == 'Boolean':
            t = 'Integer'
        if {ta, tb} == {'Single', 'Long'}:
            t = 'Double'
        if ta == 'Byte' and tb == 'Byte':
            t = 'Byte'
        elif t == 'Byte':
            t = 'Integer'
        return t

    def pure(self, e):
        """an expression without calls (literals, variables, operators on them): it can be evaluated in any order"""
        k = e.kind
        if k in ('num', 'str', 'date', 'bool', 'empty', 'null', 'nothing', 'pre'):
            return True
        hit = self.sched.get(id(e))
        if hit is not None and hit[0] is e:
            return True
        if k == 'paren':
            return self.pure(e.e)
        if k == 'un':
            return self.pure(e.e) if hasattr(e, 'e') else False
        if k == 'bin':
            return self.pure(e.l) and self.pure(e.r)
        if k == 'apply' and e.base.kind == 'name' and not e.args and getattr(e.base, 'suffix', None) is None:
            sy = self._lookup_quiet(e.base)      # a(): the whole array (an argument ByRef: not into a temporary,
            return sy is not None and sy.kind == 'var' and sy.array     # which would pass a copy - GamaSort)
        if k == 'apply' and e.base.kind == 'name' and e.args and getattr(e.base, 'suffix', None) is None:
            sy = self._lookup_quiet(e.base)      # a(i): an element of an array
            return (sy is not None and sy.kind == 'var' and sy.array
                    and all(a is not None and a.kind != 'named' and self.pure(a) for a in e.args))
        if k == 'name':
            try:
                sy = self.lookup(e.name, e, create=False)
            except Exception:
                return False
            return sy is not None and getattr(sy, 'kind', '') in ('var', 'const')
        return False

    def x_bin(self, e):
        op = e.op
        a = self.expr(e.l)
        # VB evaluates from left to right; C does not promise it for the arguments of a function:
        # when both sides call something, the left one is kept in a temporary first
        if (op != "IS" and a.code and a.arr is None and not is_udt(a.t) and not is_arr(a.t) and a.t != "void"
                and known_type(a.t) and not self.pure(e.l) and not self.pure(e.r)):
            t = self.tmp(a.t)
            self.pre.append(f"{t} = {self.conv(a, a.t)};")
            a = E(t, a.t)
        b = self.expr(e.r)
        if op == 'IS':
            for x in (a, b):
                if not (is_obj(x.t) or x.t == 'Variant'):
                    self.err("type mismatch: Is compares objects", e)
            return E(f"((vb_bool)-(vb_oid({self.conv(a, 'Object', e)}) == vb_oid({self.conv(b, 'Object', e)})))", 'Boolean')
        if is_udt(a.t) or is_udt(b.t):
            self.err("type mismatch: a Type cannot be used in an expression", e)
        if is_obj(a.t):
            a = self.cv(a, 'Variant', e)
        if is_obj(b.t):
            b = self.cv(b, 'Variant', e)
        if op == 'LIKE':
            return E(f"vb_like({self.conv(a, 'String', e)}, {self.conv(b, 'String', e)}, {self.compare_text})", 'Boolean')
        if op in ('AND', 'OR', 'XOR', 'EQV', 'IMP'):
            return self.logic(op, a, b, e)
        if op in ('=', '<>', '<', '>', '<=', '>='):
            return self.compare(op, a, b, e)
        if op == '&':
            if a.t == 'Variant' or b.t == 'Variant':
                return E(f"vb_v_cat({self.conv(a, 'Variant')}, {self.conv(b, 'Variant')})", 'Variant')
            return E(f"vb_scat({self.conv(a, 'String', e)}, {self.conv(b, 'String', e)})", 'String')
        if a.t == 'Variant' or b.t == 'Variant':
            fn = {'+': 'add', '-': 'sub', '*': 'mul', '/': 'div', '\\': 'idiv', 'MOD': 'mod', '^': 'pow'}[op]
            return E(f"vb_v_{fn}({self.conv(a, 'Variant')}, {self.conv(b, 'Variant')})", 'Variant')
        if op == '+' and a.t == 'String' and b.t == 'String':
            return E(f"vb_scat({a.code}, {b.code})", 'String')
        if a.t == 'String':
            a = self.cv(a, 'Double', e)
        if b.t == 'String':
            b = self.cv(b, 'Double', e)
        if op == '^':
            return E(f"vb_pow({self.conv(a, 'Double')}, {self.conv(b, 'Double')})", 'Double')
        if op == '/':
            # Single only when a Single meets a Single/Integer/Byte (Integer / Integer is a Double)
            if all(x.t in ('Byte', 'Integer', 'Single', 'Boolean') for x in (a, b)) and 'Single' in (a.t, b.t):
                return self.fop(a, '/', b, 'Single')
            return self.fop(a, '/', b, 'Double')
        if op in ('\\', 'MOD'):
            small = all(x.t in ('Byte', 'Integer', 'Boolean') for x in (a, b))
            if (op == '\\' and self.native and self.c.order[0] == 'fast' and b.const is not None
                    and b.const[0] == -1 and a.t in ('Integer', 'Long', 'Byte', 'Boolean')):
                # x \ -1 in native Fast code: a negation, without the overflow check (VB6)
                return self.neg_wrap(E(self.conv(a, 'Long' if not small else 'Integer'), 'Long' if not small else 'Integer'))
            fn = 'vb_idiv' if op == '\\' else 'vb_mod'
            r = f"{fn}({self.conv(a, 'Long')}, {self.conv(b, 'Long')})"
            if a.t == 'Byte' and b.t == 'Byte':
                return E(f"((vb_u8){r})", 'Byte')
            if small:
                return E(f"vb_cint((double){r})", 'Integer')
            return E(r, 'Long')
        t = self.num_type(a.t, b.t)
        fn = {'+': 'add', '-': 'sub', '*': 'mul'}[op]
        if self.noovf and t in ('Integer', 'Long', 'Byte') or (self.noovf and t == 'Currency' and op != '*'):
            # Remove Integer Overflow Checks: two's complement, as the processor does it
            A, B = self.conv(a, t), self.conv(b, t)
            sym = {'+': '+', '-': '-', '*': '*'}[op]
            if t == 'Integer':
                return E(f"((vb_i16)((vb_i32)({A}) {sym} (vb_i32)({B})))", t)
            if t == 'Byte':
                return E(f"((vb_u8)((vb_i32)({A}) {sym} (vb_i32)({B})))", t)
            if t == 'Long':
                return E(f"((vb_i32)((vb_u32)({A}) {sym} (vb_u32)({B})))", t)
            return E(f"((vb_cy)((vb_u64)({A}) {sym} (vb_u64)({B})))", t)
        if t == 'Date':
            if op == '-' and a.t == 'Date' and b.t == 'Date':
                return E(f"vb_chkd((double)({a.code}) - (double)({b.code}))", 'Double')
            sym = {'+': '+', '-': '-', '*': '*'}[op]
            return E(f"vb_chkd({self.conv(a, 'Double')} {sym} {self.conv(b, 'Double')})",
                     'Double' if op == '*' else 'Date')
        if t == 'Integer':
            return E(f"vb_{fn}_i({self.conv(a, 'Integer')}, {self.conv(b, 'Integer')})", 'Integer')
        if t == 'Long':
            return E(f"vb_{fn}_l({self.conv(a, 'Long')}, {self.conv(b, 'Long')})", 'Long')
        if t == 'Byte':
            return E(f"vb_{fn}_b({a.code}, {b.code})", 'Byte')
        if t == 'Currency':
            return E(f"vb_{fn}_c({self.conv(a, 'Currency')}, {self.conv(b, 'Currency')})", 'Currency')
        return self.fop(a, op, b, t)

    def ld(self, x):
        """an operand of a floating-point operation, in extended precision"""
        if x.ext:
            return f"({x.code})"
        return f"((vb_ld){self.conv(x, 'Double')})"

    def fop(self, a, op, b, t):
        """a + b, a - b, a * b, a / b of Single / Double: on the FPU of VB6, in extended precision; the value is
        rounded when it is stored, passed or compared"""
        fn = 'vb_fxd' if op == '/' else 'vb_fx'
        return E(f"{fn}({self.ld(a)} {op} {self.ld(b)})", t, ext=True)

    def neg_wrap(self, a):
        """-a of an Integer / Long without the overflow check: -(-32768) = -32768 (native code of VB6)"""
        if a.t == 'Long':
            return E(f"((vb_i32)(0u - (vb_u32)({a.code})))", 'Long')
        return E(f"((vb_i16)-(vb_i32)({a.code}))", 'Integer')

    def logic(self, op, a, b, e):
        if a.t == 'Variant' or b.t == 'Variant':
            return E(f"vb_v_{op.lower()}({self.conv(a, 'Variant')}, {self.conv(b, 'Variant')})", 'Variant')
        if a.t == 'Boolean' and b.t == 'Boolean':
            t, ct = 'Boolean', 'vb_bool'
        elif all(x.t in ('Boolean', 'Byte', 'Integer') for x in (a, b)):
            t, ct = ('Byte', 'vb_u8') if a.t == b.t == 'Byte' else ('Integer', 'vb_i16')
        else:
            t, ct = 'Long', 'vb_i32'
        x, y = self.conv(a, t, e), self.conv(b, t, e)
        expr = {'AND': f"({x}) & ({y})", 'OR': f"({x}) | ({y})", 'XOR': f"({x}) ^ ({y})",
                'EQV': f"~(({x}) ^ ({y}))", 'IMP': f"(~({x})) | ({y})"}[op]
        return E(f"(({ct})({expr}))", t)

    def compare(self, op, a, b, e):
        opi = {'=': 0, '<>': 1, '<': 2, '>': 3, '<=': 4, '>=': 5}[op]
        if a.t == 'Variant' or b.t == 'Variant':
            flags = self.compare_text | (2 if a.t != 'Variant' else 0) | (4 if b.t != 'Variant' else 0)
            return E(f"vb_v_cmp({self.conv(a, 'Variant')}, {self.conv(b, 'Variant')}, {opi}, {flags})", 'Variant')
        if a.t == 'String' and b.t == 'String':
            return E(f"((vb_bool)-(vb_scmp({a.code}, {b.code}, {self.compare_text}) {op_c(op)} 0))", 'Boolean')
        if a.t == 'String':
            a = self.cv(a, 'Double', e)
        if b.t == 'String':
            b = self.cv(b, 'Double', e)
        t = 'Double'
        if a.t == 'Currency' and b.t == 'Currency':
            t = 'Currency'
        elif a.t in ('Integer', 'Long', 'Byte', 'Boolean') and b.t in ('Integer', 'Long', 'Byte', 'Boolean'):
            t = 'Long'
        if t == 'Double' and 'Date' not in (a.t, b.t):
            # VB6 rounds each side to its type (a Single to a Single), unless Allow Unrounded Floating Point
            A, B = (self.ld(a), self.ld(b)) if self.unrounded else (self.conv(a, t), self.conv(b, t))
            if self.noflpt or self.unrounded:
                # NaN (no checks): the flags of the FPU say "equal" and "below". (Both sides go into variables
                # first: the Tiny C Compiler compares a long double that a function returns wrong.)
                self.ntmp += 1
                k = self.ntmp
                nan = ('-1' if op in ('=', '<', '<=') else '0') if self.noflpt else '0'
                return E(f"((vb_bool)({{ vb_ld _a{k} = {A}, _b{k} = {B}; (_a{k} != _a{k} || _b{k} != _b{k}) ? {nan} "
                         f": -(_a{k} {op_c(op)} _b{k}); }}))", 'Boolean')
            return E(f"((vb_bool)-({A} {op_c(op)} {B}))", 'Boolean')
        return E(f"((vb_bool)-({self.conv(a, t)} {op_c(op)} {self.conv(b, t)}))", 'Boolean')

    def cond(self, e):
        """C condition (int) of a VB expression used in If/While/Until"""
        x = self.expr(e)
        if is_obj(x.t):
            x = self.cv(x, 'Variant', e)
        if is_udt(x.t):
            self.err("type mismatch", e)
        if x.t == 'Variant':
            return f"vb_v_truth({x.code})"
        if x.t == 'String':
            return f"vb_s2bool({x.code})"
        if x.t == 'Currency':
            return f"(({x.code}) != 0)"
        return f"(({x.code}) != 0)"

    # ---------------------------------------------------------------- calls
    def call_dll(self, s, args, node):
        """call of a Declare procedure: vb_dllcall with one vb_dllarg per argument"""
        if any(a is not None and a.kind == 'named' for a in args):
            self.err("named arguments are not supported with Declare procedures yet", node)
        if len(args) > len(s.params):
            self.err(f"wrong number of arguments: {s.name}", node)
        out = []
        for i, p in enumerate(s.params):
            a = args[i] if i < len(args) else None
            if a is None:
                if not p.optional:
                    self.err(f"argument not optional: {p.name}", node)
                dv = E(self.c.const_code(*p.default), p.default[1]) if p.default else E('0', 'Long')
                x = dv
            elif a.kind == 'paren':             # ByVal s: the String itself (the procedure writes into it)
                x = self.expr(a.e)
            else:
                x = self.expr(a)
            byval = p.byval or (p.t == 'Any' and a is not None and a.kind == 'paren')
            if p.array:
                ref, _, _ = self.array_target(a, a)
                out.append(f"{{ VBD_PTR, 0, 0, (void *)(({ref})->data), 0 }}")
                continue
            if p.t == 'Any' and is_udt(x.t):
                byval = False                   # ByVal t (a Type) for As Any: its address all the same
            if byval:
                t = x.t if p.t == 'Any' else p.t
                if t == 'Variant' and p.t == 'Any':
                    pt = self.ctl_prop_type(a.e if a.kind == 'paren' else a)
                    if pt in ('String', 'Long', 'Integer', 'Boolean', 'Byte', 'Single', 'Double'):
                        # ByVal Text1.Text: VB6 binds the property of an intrinsic control early, its type (String)
                        # goes, not a Variant (WizzForm Sender, Sound Association: was error 49)
                        x, t = E(self.conv(x, pt, a), pt), pt
                if t == 'Variant' and p.t == 'Any':
                    # VB6 pushes the 16 bytes of the VARIANT: "Bad DLL calling convention" (error 49)
                    self.pre.append(f"(void)({self.conv(x, 'Variant', a)});")
                    out.append("{ VBD_BADCC, 0, 0, 0, 0 }")
                    continue
                if t == 'Variant':
                    t = 'Long'
                if t == 'String':
                    if x.ptr is not None and x.t == 'String' and not x.prop:
                        ptr = x.ptr
                    else:
                        tmp = self.tmp('String')
                        self.pre.append(f"vb_sset(&{tmp}, {self.conv(x, 'String', a)});")
                        self.post.append(f"vb_sfree(&{tmp});")
                        ptr = f"&{tmp}"
                    out.append(f"{{ VBD_STR, 0, 0, 0, {ptr} }}")
                elif t in ('Double', 'Single', 'Date'):
                    out.append(f"{{ VBD_DBL, 0, (double)({self.conv(x, t, a)}), 0, 0 }}")
                elif t == 'Currency':
                    out.append(f"{{ VBD_INT, (vb_i64)({self.conv(x, t, a)}), 0, 0, 0 }}")
                elif is_obj(t):
                    out.append(f"{{ VBD_INT, (vb_i64)vb_varptr(vb_oid({self.conv(x, 'Object', a)}), "
                               f"(vb_i32)sizeof(vbobjh)), 0, 0, 0 }}")
                else:
                    out.append(f"{{ VBD_INT, (vb_i64)({self.conv(x, t, a)}), 0, 0, 0 }}")
                continue
            # ByRef: the address of the variable (a copy for an expression)
            t = x.t if p.t == 'Any' else p.t
            if is_udt(t) and x.t == t and x.ptr is not None and not x.prop and t[2:] in self.c.udts:
                # a Type: in the layout of VB6 (ANSI strings), made and read back by the runtime
                out.append(f"{{ VBD_UDT, 0, 0, (void *)({x.ptr}), (vbs *)&UD_{t[2:]} }}")
                continue
            if (p.t == 'Any' or is_obj(p.t) or p.t == 'Object') and is_obj(x.t) and x.ptr is not None and not x.prop:
                # (IPic As IPicture of OleCreatePictureIndirect: the object the procedure makes)
                # an object variable: its ObjPtr, and what the procedure leaves becomes the object (VB6 trick)
                out.append(f"{{ VBD_OREF, 0, 0, (void *)({x.ptr}), 0 }}")
                continue
            # a String: the address of an ANSI BSTR (also for As Any); a Variant: a VARIANT
            def arg(ref):
                if t == 'String':
                    return f"{{ VBD_SREF, 0, 0, 0, {ref} }}"
                return f"{{ {'VBD_VAR' if t == 'Variant' else 'VBD_PTR'}, 0, 0, (void *)({ref}), 0 }}"
            if x.ptr is not None and (x.t == t) and not x.prop and not (t == 'String' and x.fix):
                out.append(arg(x.ptr))
                continue
            tmp = self.tmp(t)
            self.pre.append(self.store_code(E(tmp, t, ptr=f"&{tmp}"), x, a))
            if x.ptr is not None and not x.prop:
                self.post.append(f"if (!vb_errpend) {self.store_code(x, E(tmp, t), a)}")
            self.free_after(tmp, t, self.post)
            out.append(arg(f"&{tmp}"))
        argv = f"{len(out)}, (vb_dllarg[]){{ {', '.join(out)} }}" if out else "0, 0"
        rt = s.rt
        if rt is None:
            return E(f"(void)vb_dllcall(&{s.cname}, {argv})", 'void')
        if rt in ('Double', 'Single', 'Date'):
            return E(f"(({ctype(rt)})vb_dllcall_d(&{s.cname}, {argv}))", rt)
        if rt in ('Long', 'Integer', 'Byte', 'Boolean', 'Currency'):
            return E(f"(({ctype(rt)})vb_dllcall(&{s.cname}, {argv}))", rt)
        if rt == 'Variant':
            return E(f"vb_v_l((vb_i32)vb_dllcall(&{s.cname}, {argv}))", rt)
        if rt == 'String':
            return E(f"vb_dllcall_s(&{s.cname}, {argv})", rt)
        self.err(f"Declare functions returning {rt} are not supported yet", node)

    def call_proc(self, s, args, node, as_stmt=False, self_code=None):
        if getattr(s, 'dll', False):
            return self.call_dll(s, args, node)
        if getattr(s, 'cls', None) and self_code is None:
            if self.cls != s.cls:
                self.err(f"{s.name} is a member of class {self.c.classes[s.cls].name}: use object.{s.name}", node)
            self_code = 'vb_self'
        params = s.params
        pos, named = [], {}
        for a in args:
            if a is not None and a.kind == 'named':
                named[a.name.lower()] = a.e
            else:
                if named:
                    self.err("expected: named argument", node)
                pos.append(a)
        cargs, post = [], []
        if params and getattr(params[-1], 'paramarray', False):
            rest = pos[len(params) - 1:]
            pos = pos[:len(params) - 1]
            vals = []
            for a in rest:
                if a is None:
                    vals.append('vb_v_missing()')
                else:
                    x = self.expr(a)
                    vals.append(self.conv(x, 'Variant', a))
            t = self.tmp('A:Variant')
            self.pre.append(f"vb_aparam(&{t}, {len(vals)}, " + (f"(vbvar[]){{ {', '.join(vals)} }});" if vals else "0);"))
            post.append(f"vb_afree(&{t});")
            params = params[:-1]
            pa_arg = f"&{t}"
        else:
            pa_arg = None
        if len(pos) > len(params):
            self.err(f"wrong number of arguments: {s.name}", node)
        for i, p in enumerate(params):
            a = pos[i] if i < len(pos) else named.pop(p.name.lower(), None)
            if a is None:
                if not p.optional:
                    self.err(f"argument not optional: {p.name}", node)
                if p.default is not None:
                    dv = E(self.c.const_code(*p.default), p.default[1])
                elif p.t == 'Variant':
                    dv = E('vb_v_missing()', 'Variant')
                else:
                    dv = E(self.c.const_code(*default_of(p.t)), p.t)
                if p.byval:
                    cargs.append(self.conv(dv, p.t))
                else:
                    t = self.tmp(p.t)
                    self.pre.append(self.store_code(E(t, p.t, ptr=f"&{t}"), dv, node))
                    self.free_after(t, p.t, post)
                    cargs.append(f"&{t}")
                continue
            if p.array:
                if a.kind == 'apply' and not a.args:           # Fill s(): the array named with ()
                    a = a.base
                if a.kind in ('name', 'member'):
                    ref, et, _ = self.array_target(a, a)
                    if et != p.t:
                        self.err(f"ByRef argument type mismatch: {self.describe(a)}", a)
                    cargs.append(ref)
                    continue
                if a.kind == 'pre' and a.value.arr is not None and a.value.code is None:
                    # an array evaluated first (order_args: another argument calls something): the array itself,
                    # ByRef, not a copy - Wav_BuildHeader udt(i).bytes, rate, ... (a UDT field), as GamaSort t()
                    x = a.value
                    xt = x.t[2:] if is_arr(x.t) else x.t
                    if xt != p.t:
                        self.err(f"ByRef argument type mismatch: {self.describe(a)}", a)
                    cargs.append(x.arr)
                    continue
                cargs.append(self.array_temp(self.expr(a), p.t, a))   # Inv(Inv(m)), StrConv(s, vbFromUnicode)
                continue
            x = self.expr(a)
            if x.t == 'void':
                self.err(f"expected function or variable: {self.describe(a)}", a)
            if p.byval:
                cargs.append(self.conv(x, p.t, a))
                continue
            if x.ptr is not None and x.t == p.t and not x.prop and not x.fix:
                lk = self.locked_array(a)
                if lk:                          # an element of a dynamic array: locked while the call runs
                    self.pre.append(f"vb_alock({lk});")
                    post.append(f"vb_aunlock({lk});")
                cargs.append(x.ptr)
                continue
            if is_udt(p.t) and x.t == 'Variant' and self.c.rec_ok(p.t):
                # a Variant holding the Type (VT_RECORD) to a ByRef parameter of that Type: a copy, back after
                k = p.t[2:]
                t = self.tmp(p.t)
                self.pre.append(f"U_{k}_init(&{t}); vb_v2rec({self.conv(x, 'Variant', a)}, &UD_{k}, &{t});")
                if x.ptr is not None and a.kind not in ('paren', 'pre') and not x.prop:
                    post.append(f"if (!vb_errpend) vb_vset({x.ptr}, vb_v_rec(&{t}, &UD_{k}));")
                self.free_after(t, p.t, post)
                cargs.append(f"&{t}")
                continue
            if is_udt(p.t):
                if x.t != p.t:
                    self.err(f"ByRef argument type mismatch: {self.describe(a)}", a)
                t = self.tmp(p.t)
                self.pre.append(f"U_{p.t[2:]}_copy(&{t}, &{self.udt_value(x)});")
                self.free_after(t, p.t, post)
                cargs.append(f"&{t}")
                continue
            if x.ptr is not None and not x.prop and not x.fix and a.kind not in ('paren', 'pre') and p.t != 'Variant' \
                    and not (is_obj(p.t) and is_obj(x.t)):
                self.err(f"ByRef argument type mismatch: {self.describe(a)}", a)
            if is_arr(x.t) and p.t == 'Variant' and x.ptr is not None and a.kind not in ('paren', 'pre'):
                # the array of a Function, by its name inside it, to a ByRef Variant: the array goes in a Variant
                # and comes back from it (FreeImage_ZLibCompressEx(FreeImage_ZLibCompressVB, ...))
                t = self.tmp('Variant')
                self.pre.append(f"vb_vset(&{t}, {self.conv(x, 'Variant', a)});")
                post.append(f"if (!vb_errpend) vb_v2arr({t}, {x.ptr}, {vt_of(x.t[2:])});")
                self.free_after(t, 'Variant', post)
                cargs.append(f"&{t}")
                continue
            # a value (or a variable of another type passed to a Variant): pass a copy
            t = self.tmp(p.t)
            self.pre.append(self.store_code(E(t, p.t, ptr=f"&{t}"), x, a))
            if x.ptr is not None and a.kind not in ('paren', 'pre') and not x.prop:
                # ByRef Variant parameter with a typed variable: copy the result back
                post.append(f"if (!vb_errpend) {self.store_code(x, E(t, p.t), a)}")
            self.free_after(t, p.t, post)
            cargs.append(f"&{t}")
        if named:
            self.err(f"named argument not found: {next(iter(named))}", node)
        if pa_arg is not None:
            cargs.append(pa_arg)
        if self_code is not None:
            cargs.insert(0, self_code)
        call = f"{s.cname}({', '.join(cargs)})"
        if s.rt and is_arr(s.rt):
            r = self.tmp(s.rt)
            self.pre.append(f"{r} = {call};")
            self.post.extend(post)
            self.post.append(f"vb_afree(&{r});")
            return E(None, s.rt, arr=f"&{r}")
        if s.rt and is_udt(s.rt):
            # a Type returned by value: keep it in a temporary until the end of the statement
            r = self.tmp(s.rt)
            self.pre.append(f"{r} = {call};")
            self.post.extend(post)
            self.post.append(f"U_{s.rt[2:]}_free(&{r});")
            return E(r, s.rt, ptr=f"&{r}")
        if post:
            if s.rt and not as_stmt:
                r = self.tmp(s.rt)
                self.pre.append(f"{r} = {call};")
                self.post.extend(post)
                return E(r, s.rt)
            self.post.extend(post)
        return E(call, s.rt or 'void')

    def free_after(self, t, ty, post):
        # temporaries that may receive an owned value from the callee
        c = udt_free_code(ty, '&' + t)
        if c:
            post.append(c)

    def udt_value(self, x):
        """C lvalue holding the Type value x"""
        if x.ptr is not None:
            return f"(*{x.ptr})"
        return x.code

    def udt_addr_value(self, x):
        """as udt_value, also for a Type a function gives (kept in a temporary first)"""
        if x.ptr is not None or not x.code or not re.search(r'\w\(', x.code):
            return self.udt_value(x)
        t = self.tmp(x.t)
        self.pre.append(f"{t} = {x.code};")
        self.free_after(t, x.t, self.cleanup)
        return t

    # ---------------------------------------------------------------- assignment
    def store_code(self, target, value, node):
        """C statement storing value into the variable expression target"""
        t = target.t
        if t == 'String' and target.fix:
            return f"vb_sset_fix({target.ptr}, {self.conv(value, 'String', node)}, {target.fix});"
        if t == 'String':
            return f"vb_sset({target.ptr}, {self.conv(value, 'String', node)});"
        if t == 'Variant':
            return f"vb_vset({target.ptr}, {self.conv(value, 'Variant', node)});"
        if is_obj(t):
            return f"vb_oset({target.ptr}, {self.conv(value, t, node)});"
        if is_udt(t):
            if value.t == 'Variant' and self.c.rec_ok(t):     # oItem = cItems(1): the Type in the Variant
                return f"vb_v2rec({self.conv(value, 'Variant', node)}, &UD_{t[2:]}, {target.ptr});"
            if value.t != t:
                self.err("type mismatch", node)
            return f"U_{t[2:]}_copy({target.ptr}, &{self.udt_value(value)});"
        return f"*{target.ptr} = {self.conv(value, t, node)};"

    def lvalue(self, e):
        if e.kind == 'name':
            s = self.lookup(e.name, e)
            if s is not None and s.kind == 'formdefault':
                return self.var_ref(s)
            if s is None:
                s = self.implicit_var(e.name, e.suffix, e)
            if s.kind == 'const':
                self.err(f"assignment to constant not permitted: {e.name}", e)
            if s.kind == 'proc':
                self.err(f"function call on left-hand side of assignment: {e.name}", e)
            if s.kind != 'var':
                self.err(f"invalid assignment target: {e.name}", e)
            if s.array:
                self.err(f"can't assign to array: {e.name}", e)
            self.check_suffix(s, e)
            return self.var_ref(s)
        if e.kind == 'apply' and e.base.kind == 'name':
            s = self.lookup(e.base.name, e.base)
            if s is not None and s.kind == 'var' and s.array:
                return self.array_elem(s, e.args, e)
        if e.kind == 'apply' and e.base.kind == 'member':
            return None
        if e.kind == 'member':
            return None
        self.err("invalid assignment target", e)

    # ---------------------------------------------------------------- statements
    def block(self, stmts):
        for st in stmts:
            self.statement(st)

    def statement(self, st):
        k = st.kind
        if k == 'label':
            idx = self.labels[st.name.lower()]
            self.emit(f"L{idx}:;")
            if st.name.isdigit():
                self.emit(f"vb_line = {st.name};")
            return
        if k == 'nop':
            return
        self.nstmt += 1
        n = self.nstmt
        self.pre, self.post = [], []
        pre, post = self.pre, self.post         # this statement's own (the statements of a block have theirs)
        self.emit(f"S{n}:;")
        self.emit(f"vb_cur = {n};")
        if self.c.debug and getattr(st, 'line', 0):
            self.emit(f"VB_DBG({st.line});")
            self.dbg_lines.append((st.line, n))
        save = self.lines
        self.lines = []
        m = getattr(self, 's_' + k, None)
        if m is None:
            self.err(f"unsupported statement ({k})", st)
        m(st)
        body = self.lines
        self.lines = save
        # temporaries prepared for the calls of this statement come first
        for p in pre:
            self.emit(p)
        self.lines.extend(body)
        for p in post:
            self.emit(p)
        if k not in ('if', 'for', 'do', 'select', 'with', 'foreach'):
            self.emit("vb_release(vb_m0);")
            self.emit("if (vb_errpend) goto vb_eh;")
        self.emit(f"R{n}:;")

    def check(self):
        """release temporaries and jump to the handler after an expression evaluated inside a block statement"""
        for p in self.post:
            self.emit(p)
        del self.post[:]
        self.emit("vb_release(vb_m0);")
        self.emit("if (vb_errpend) goto vb_eh;")

    def flush_pre(self):
        for p in self.pre:
            self.emit(p)
        del self.pre[:]

    def sub_target(self, b):
        """b is a Sub (or Debug.Assert / Debug.Print), so b (x) = y is a call, not an assignment"""
        if b.kind == 'member' and b.base.kind == 'name' and b.base.name.lower() == 'debug' \
                and b.name.lower() in ('assert', 'print') and self.lookup(b.base.name, b.base) is None:
            return True
        if b.kind == 'name':
            s = self.lookup(b.name, b)
            return s is not None and s.kind == 'proc' and getattr(s, 'pkind', '') == 'sub'
        return False

    def s_assign(self, st):
        tg = self.unqual(st.target)              # Module1.Count = 0
        st.target = tg
        if tg.kind == 'name' and not st.set and tg.name.lower() in ('date', 'time') and self.lookup(tg.name, tg) is None:
            # Date = d / Time = t: VB6 sets the clock of the system; here the value is only checked (13)
            self.emit(f"(void)({self.conv(self.expr(st.value), 'Date', st)});")
            return
        if tg.kind == 'apply' and len(tg.args or []) == 1 and tg.args[0] is not None and tg.args[0].kind != 'named' \
                and self.sub_target(tg.base):
            # Debug.Assert (n Mod 4) = 0: a call with a comparison; MySub (a) = b is VB6's compile error
            # "Expected Function or variable" (compile-6.8 c5)
            if tg.base.kind == 'name':
                self.err("expected Function or variable", tg)
            arg = N('bin', st.line, op='=', l=N('paren', st.line, e=tg.args[0]), r=st.value)
            return self.s_call(N('call', st.line, target=tg.base, args=[arg], paren=False))
        if tg.kind == 'apply' and not tg.args and tg.base.kind in ('name', 'member'):
            tg = st.target = tg.base            # parts() = Split(s, ":"): the whole array
        kind = 'VBM_SET' if st.set else 'VBM_LET'
        # property of this module / class used by name: Name = v
        if tg.kind in ('name', 'apply') and (tg.base if tg.kind == 'apply' else tg).kind == 'name':
            nm = tg.base if tg.kind == 'apply' else tg
            sy = self.lookup(nm.name, nm)
            if sy is not None and sy.kind == 'prop':
                p = sy.let if not st.set else sy.set
                args = list(tg.args) if tg.kind == 'apply' else []
                if p is None and not st.set and sy.get is not None and (
                        is_obj(sy.get.rt or '') or sy.get.rt in ('Object', 'Variant')):
                    # Font = Parent.Font.Name with Get / Set only: the default member of what Get gives
                    got = self.call_proc(sy.get, args, st)
                    r = self.late(got, '', [], st, 'VBM_LET', self.expr(st.value))
                    self.emit(f"(void)({r.code});")
                    return
                if p is None:
                    self.err(f"invalid use of property: {sy.name} has no Property {'Set' if st.set else 'Let'}", st)
                r = self.call_proc(p, args + [_Pre(self.expr(st.value))], st, as_stmt=True)
                self.emit(r.code + ';')
                return
            if sy is not None and sy.kind == 'me' and tg.kind == 'apply' and not st.set and \
                    self.c.classes[self.cls].is_form:
                # Me!txtOut = s, Me("Text1") = s: the default property of that control of the form
                ctl = self.late(E('vb_self', 'Object'), 'controls', list(tg.args), st, 'VBM_GET')
                r = self.late(ctl, '', [], st, 'VBM_LET', self.expr(st.value))
                self.emit(f"(void)({r.code});")
                return
            if sy is not None and sy.kind == 'me':
                self.err("invalid use of Me keyword", st)
            if sy is not None and sy.kind == 'ctl' and not st.set:   # Text1 = "x": the default property
                r = self.default_member(self.var_ref(sy), list(tg.args) if tg.kind == 'apply' else [], st,
                                        kind='VBM_LET', value=self.expr(st.value))
                self.emit(f"(void)({r.code});" if r.t != 'void' else r.code + ';')
                return
            if sy is None and tg.kind == 'name' and nm.name.lower() == 'printer' and st.set:
                # Set Printer = Printers(i): the program's printer stays the one it prints to
                self.emit(f"(void)({self.conv(self.expr(st.value), 'Object', st)});")
                return
            if sy is None and tg.kind == 'name' and nm.name.lower() == 'err' and not st.set:   # Err = 0: Err.Number
                B.lookup_member_set('err', 'number')(self, st, self.expr(st.value))
                return
            if sy is None and tg.kind == 'name' and self.form_builtin(nm.name):
                r = self.late(E('vb_self', 'O:' + self.cls), nm.name.lower(), [], st, kind, self.expr(st.value))
                self.emit(f"(void)({r.code});")
                return
        if tg.kind == 'member' or (tg.kind == 'apply' and tg.base.kind == 'member'):
            m = tg if tg.kind == 'member' else tg.base
            args = None if tg.kind == 'member' else tg.args
            if m.base.kind == 'name' and self.lookup(m.base.name, m.base) is None and not st.set:
                h = B.lookup_member_set(m.base.name.lower(), m.name.lower())
                if h:
                    h(self, st, self.expr(st.value))
                    return
            value = self.expr(st.value)
            r = self.member(m, args, st, kind=kind, value=value)
            if r.code is None and r.arr is not None and not args:   # t.Data = Split(s, ","): an array in a Type
                self.assign_array(r.arr, r.t[2:] if is_arr(r.t) else r.t, value, st)
                return
            if r.code is None:
                self.err("can't assign to array", st)
            if r.t == 'void' or (r.t == 'Variant' and r.ptr is None and r.code.startswith(('vb_late', 'vb_late_o', 'vbf_uc_member'))):
                self.emit(f"(void)({r.code});" if r.t != 'void' else r.code + ';')
                return
            if r.ptr is None:
                self.err("invalid assignment target", st)
            self.assign_to(r, value, st)
            return
        if (st.set and st.value.kind != 'nothing' and tg.kind == 'apply' and tg.base.kind == 'name' and self.cls
                and tg.base.name.lower() in getattr(self.c.classes.get(self.cls), 'controls', {})
                and self.lookup(tg.base.name, tg.base) is not None and self.lookup(tg.base.name, tg.base).kind == 'ctl'):
            # Set Image1(i) = pic: the default property of the element (Picture)
            self.s_assign(N('assign', st.line, target=N('member', st.line, base=tg, name='Picture', suffix=None),
                            value=st.value, set=True))
            return
        if (st.set and st.value.kind != 'nothing' and tg.kind == 'name' and self.cls
                and tg.name.lower() in getattr(self.c.classes.get(self.cls), 'controls', {})
                and str(self.c.classes[self.cls].ctl_cls.get(tg.name.lower(), '')).lower() in (
                    'vb.image', 'vb.picturebox')
                and self.lookup(tg.name, tg) is not None and self.lookup(tg.name, tg).kind == 'ctl'):
            # Set Picture1 = LoadPicture(f): the default property of the control (Picture)
            self.s_assign(N('assign', st.line, target=N('member', st.line, base=tg, name='Picture', suffix=None),
                            value=st.value, set=True))
            return
        if (st.set and st.value.kind == 'nothing' and tg.kind in ('apply', 'name') and self.cls
                and (tg.base.name if tg.kind == 'apply' and tg.base.kind == 'name' else getattr(tg, 'name', '')).lower()
                in getattr(self.c.classes.get(self.cls), 'controls', {})):
            return                               # Set Image1(i) = Nothing: VB6 accepts it and nothing happens
        if tg.kind in ('apply', 'name') and (tg.base if tg.kind == 'apply' else tg).kind == 'name':
            nm = tg.base if tg.kind == 'apply' else tg
            if self.lookup(nm.name, nm) is None:
                g = self.comp_global(nm.name)    # ShowHeaderIcon(h) = True: a property of a GlobalMultiUse class
                if g is not None:
                    r = self.late(g, nm.name.lower(), list(tg.args) if tg.kind == 'apply' else [], st, kind,
                                  self.expr(st.value))
                    self.emit(f"(void)({r.code});")
                    return
        if tg.kind == 'apply' and tg.base.kind == 'name':
            sy = self.lookup(tg.base.name, tg.base)
            if sy is not None and sy.kind in ('var', 'me') and not sy.array and (is_obj(sy.t) or sy.t == 'Variant'):
                value = self.expr(st.value)
                r = self.default_member(self.var_ref(sy), tg.args, st, kind=kind, value=value)
                self.emit(f"(void)({r.code});" if r.t != 'void' else r.code + ';')
                return
        if tg.kind == 'apply' and tg.base.kind in ('apply', 'withref'):   # arr(i)("key") = v, !Field = v in With
            obj = self.expr(tg.base)
            if is_obj(obj.t) or obj.t == 'Variant':
                r = self.default_member(obj, tg.args, st, kind=kind, value=self.expr(st.value))
                self.emit(f"(void)({r.code});" if r.t != 'void' else r.code + ';')
                return
        if tg.kind == 'name' and not st.set:
            sy = self.lookup(tg.name, tg)
            if sy is not None and sy.kind == 'var' and sy.array:
                self.assign_array(self.arr_ref(sy), sy.t, self.expr(st.value), st,
                                  fixed=bool(getattr(sy, 'fixed_dims', None)) or self.is_fixed_local(sy))
                return
        if tg.kind == 'apply' and tg.base.kind == 'name' and tg.args and self.assign_elem(tg, st):
            return
        target = self.lvalue(tg)
        if target is None:
            self.err(f"invalid assignment target ({self.describe(tg)})", st)
        self.assign_to(target, self.expr(st.value), st)

    def assign_elem(self, tg, st):
        """a(i) = v with calls in i: P-code (and native code without the array bounds checks) evaluates v, then
        i; native code the indices of the arrays read in v, then i, then v"""
        sy = self._lookup_quiet(tg.base)
        if sy is None or sy.kind != 'var' or not sy.array:
            return False
        if all(self.ev_count(a) == 0 for a in tg.args if a is not None):
            return False
        mode, nobounds = self.c.order[0], self.c.order[2]
        snap, keys = self.snapshot(), []
        try:
            if mode == 'pcode' or nobounds:
                ok = self.stmt_hoist(st.value, keys)
            else:
                ok = all(self.stmt_hoist(None, keys, args=list(r.args)) for r in self.array_reads(st.value, []))
            ok = ok and self.stmt_hoist(None, keys, args=list(tg.args))
            if ok:
                target = self.lvalue(tg)
                value = self.expr(st.value)
                if target is not None and all(self.sched[k][2] for k in keys):
                    self.assign_to(target, value, st)
                    return True
            self.rollback(snap)
            return False
        except BaseException:
            self.rollback(snap)
            raise
        finally:
            for k in keys:
                self.sched.pop(k, None)

    def is_fixed_local(self, sy):
        return getattr(sy, 'fixed_local', False)

    def assign_array(self, ref, et, value, st, fixed=False):
        """array = array / Variant / function result"""
        if fixed:
            self.err("can't assign to array", st)
        vt = value.t[2:] if is_arr(value.t) else value.t
        if is_arr(value.t) and vt != et and 'Variant' in (vt, et):
            # a() As Variant = Split(...): VB6 makes the program, the assignment is error 13 when it runs
            # (compile-6.8 c13)
            if value.code is not None:
                self.emit(f"(void)({value.code});")
            self.emit("vb_raise(13);")
            return
        if value.arr is not None and value.code is None:
            if vt != et:
                self.err("type mismatch: arrays of different types", st)
            self.emit(f"vb_acopy({ref}, {value.arr});")
        elif value.t == 'Variant':
            self.emit(f"vb_v2arr({value.code}, {ref}, {vt_of(et)});")
        elif value.t == 'String' and et == 'Byte':      # b = "text" / StrConv(s, vbFromUnicode)
            self.emit(f"vb_s2barr({value.code}, {ref});")
        else:
            self.err("type mismatch: an array is needed", st)

    def assign_to(self, target, value, st):
        if is_arr(target.t):                    # Function F() As T(): F = array
            self.assign_array(target.ptr, target.t[2:], value, st)
            return
        if value.t == 'void':
            self.err("expected function or variable", st)
        if st.set:
            if not (is_obj(target.t) or target.t == 'Variant'):
                self.err("Set needs an object variable (or a Variant)", st)
            if not (is_obj(value.t) or value.t == 'Variant'):
                self.err("type mismatch: Set needs an object", st)
            if value.code == 'vb_global_obj()':
                # Set g = VB.Global: VB6 raises 13 (t_ph68: VB.Global is no Automation object; TypeName "Unknown")
                self.emit("vb_raise(13);")
                return
            # an error in the value (On Error Resume Next): the variable keeps its object (VB6 oracle d_ado2)
            if target.t == 'Variant':
                self.emit(f"{{ vbobj lb_so = {self.conv(value, 'Object', st)}; if (!vb_errpend) vb_vset({target.ptr}, vb_v_obj(lb_so)); }}")
            elif target.wev:
                self.emit(f"{{ vbobj lb_so = {self.conv(value, target.t, st)}; if (!vb_errpend) vb_wev_set({target.wev[1]}, {target.ptr}, lb_so, {target.wev[0]}); }}")
            else:
                self.emit(f"{{ vbobj lb_so = {self.conv(value, target.t, st)}; if (!vb_errpend) vb_oset({target.ptr}, lb_so); }}")
            return
        if is_obj(target.t):
            if getattr(st, 'value', None) is not None and getattr(st.value, 'kind', '') == 'nothing':
                self.err("invalid use of object: use Set to assign an object", st)
            # a default member that needs an argument (Item of a Collection): VB6's compile error
            # "Argument not optional" (compile-6.8 c1: F = New Collection in Function F() As Collection)
            if self.default_needs_args(target.t):
                self.err("argument not optional", st)
            # obj = value (Let) on an object variable: its default property, at run time (VB6: error 91 if
            # Nothing, 438 if the class has none)
            self.emit(f'(void)vb_late_o({target.code}, "", VBM_LET, 1, (vbvar[]){{ {self.conv(value, "Variant", st)} }});')
            return
        if target.t == 'Variant':               # v = Label1 (Let): the value of the default property, not the object
            if is_obj(value.t):
                value = E(f'vb_late_o({self.conv(value, "Object", st)}, "", VBM_GET, 0, 0)', 'Variant')
            elif value.t == 'Variant' and value.const is None:
                value = E(f"vb_v_letval({value.code})", 'Variant')
        if target.t in ('String', 'Variant'):
            # an error in the value: the variable keeps what it had (VB6)
            ct = 'vbs' if target.t == 'String' else 'vbvar'
            v = self.conv(value, target.t, st)
            self.emit(f"{{ {ct} lb_sv = {v}; if (!vb_errpend) {self.store_code(target, E('lb_sv', target.t), st)} }}")
        elif is_udt(target.t):
            self.emit(self.store_code(target, value, st))
        else:
            value = self.byte_quirk(target, value)
            tmp = self.tmp(target.t)
            self.emit(f"{tmp} = {self.conv(value, target.t, st)};")
            # a floating-point error: VB6 stores the result (Inf, NaN), then raises the error
            guard = "!vb_errpend || vb_fpstore" if target.t in ('Single', 'Double') else "!vb_errpend"
            self.emit(f"if ({guard}) *{target.ptr} = {tmp};")

    def byte_quirk(self, target, value):
        """Byte + Byte stored in an Integer / Long: native Fast code, in a procedure without On Error, makes it
        without the overflow check (VB6: 255 + 255 -> 254)"""
        if not (self.c.order[0] == 'fast' and not self.handler and target.t in ('Integer', 'Long')
                and value.t == 'Byte' and value.code):
            return value
        for f in ('vb_add_b(', 'vb_sub_b(', 'vb_mul_b('):
            c = value.code
            if c.startswith(f):
                depth = 0
                for i, ch in enumerate(c):
                    depth += ch == '('
                    depth -= ch == ')'
                    if depth == 0 and ch == ')':
                        break
                if i == len(c) - 1:
                    return E(f[:-1] + 'w(' + c[len(f):], 'Byte')
        return value

    def s_call(self, st):
        tgt = self.unqual(st.target)
        if tgt is not st.target:
            st = N(st.kind, st.line, **{**{k: v for k, v in st.__dict__.items() if k not in ('kind', 'line')}, 'target': tgt})
        if st.args:
            st = N(st.kind, st.line, **{**{k: v for k, v in st.__dict__.items() if k not in ('kind', 'line')},
                                        'args': self.order_args(st.args)})
        if tgt.kind == 'name':
            s = self.lookup(tgt.name, tgt)
            if s is not None and s.kind == 'proc':
                r = self.call_proc(s, st.args, st, as_stmt=True)
                if r.code is not None:               # (a Function returning an array: its call is done already)
                    self.emit(r.code + ';')
                return
            if s is None:
                b = B.lookup_stmt(tgt.name)
                if b:
                    b(self, st, st.args)
                    return
                fm = self.form_builtin(tgt.name, tgt)
                if fm:
                    r = self.late(E('vb_self', 'O:' + self.cls), fm, st.args, st, 'VBM_METHOD')
                    self.emit(f"(void)({r.code});")
                    return
                if tgt.name.lower() == 'propertychanged' and self.cls and len(st.args or []) == 1:
                    # a class bound by a BindingCollection: the binding learns that the property changed
                    a = self.conv(self.expr(st.args[0] if st.args[0].kind != 'named' else st.args[0].e), 'String', st)
                    self.emit(f"vb_prop_changed(vb_self, {a});")
                    return
                b = B.lookup(tgt.name, tgt.suffix)
                if b:
                    r = b(self, tgt, st.args)
                    self.emit(f"(void)({r.code});" if r.t != 'void' else r.code + ';')
                    return
                g = self.comp_global(tgt.name)
                if g is not None:
                    r = self.late(g, tgt.name.lower(), st.args, st, 'VBM_METHOD')
                    self.emit(f"(void)({r.code});")
                    return
                self.err(f"Sub or Function not defined: {tgt.name}", st)
            if s.kind == 'prop' and s.get is not None:
                r = self.prop_get(s, st.args, st)
                self.emit(f"(void)({r.code});")
                return
            if s.kind == 'var' and s.storage == 'ret' and self.sym is not None and self.sym.kind == 'proc':
                # QuickSort a, lo, hi / Call Walk(x) inside Function QuickSort: a recursive call, as in VB6
                r = self.call_proc(self.sym, st.args, st, as_stmt=True)
                if r.code is not None:
                    self.emit(r.code + ';')
                return
            self.err(f"expected Sub, Function or Property: {tgt.name}", st)
        if tgt.kind == 'member':
            if tgt.base.kind == 'name' and self.lookup(tgt.base.name, tgt.base) is None:
                h = B.lookup_method(tgt.base.name.lower(), tgt.name.lower())
                if h:
                    h(self, st, st.args)
                    return
            r = self.member(tgt, st.args, st, kind='VBM_METHOD')
            if r.code is None and r.arr is not None:     # a method returning an array: called already
                return
            if r.code is None:
                self.err("expected Sub, Function or Property", st)
            self.emit(f"(void)({r.code});" if r.t != 'void' else r.code + ';')
            return
        if tgt.kind == 'apply':
            r = self.expr(N('apply', st.line, base=tgt, args=st.args)) if st.args else self.expr(tgt)
            self.emit(f"(void)({r.code});" if r.t != 'void' else r.code + ';')
            return
        self.err(f"expected Sub, Function or Property ({self.describe(tgt)})", st)

    def s_dim(self, st):
        static = st.scope == 'static' or self.sym.node.static
        for v in st.vars:
            t = self.c.type_of_decl(self.mk, v.name, v.suffix, v.type, v)
            fixlen = self.c.fixlen_of(self.mk, v.type, v, self.locals)
            if v.dims is None:
                sym = self.new_local(v.name, t, v, static=static)
                sym.ext = self.c.ext_of(t, v.type)
                if v.type is not None and v.type.new:
                    sym.autonew = self.c.autonew_of(t, v.type)
                if fixlen:
                    sym.fixlen = fixlen
                    if static:
                        self.c.protos.append(f"static int {sym.cname}_set;")
                        self.c.static_cnames.append(f"{sym.cname}_set")
                        self.prologue.append(f"if (!{sym.cname}_set) {{ {sym.cname}_set = 1; vb_sfix_init(&{sym.cname}, {fixlen}); }}")
                    else:
                        self.prologue.append(f"vb_sfix_init(&{sym.cname}, {fixlen});")
                continue
            sym = self.new_local(v.name, t, v, array=True, ndims=len(v.dims) if v.dims else None, static=static)
            sym.fixlen = fixlen
            if v.type is not None and v.type.new:
                sym.autonew = self.c.autonew_of(t, v.type)   # Dim a(3) As New C: each element is made when it is used
            sym.fixed_local = bool(v.dims)
            if v.dims:
                bounds = [self.c.dim_bounds(d, self.mk, v, self.locals) for d in v.dims]
                code = ' '.join(self.c.adim_code(f"&{sym.cname}", t, bounds, fixed=True, fixlen=fixlen))
                if static:
                    self.prologue.append(f"if (!{sym.cname}.data) {{ {code} }}")
                else:
                    self.prologue.append(code)

    def s_const(self, st):
        for c in st.items:
            v, t = self.c.eval_const(c.value, self.mk, self.locals)
            tt = self.c.type_of_decl(self.mk, c.name, c.suffix, c.type, c) if (c.type or c.suffix) else t
            self.locals[c.name.lower()] = Sym('const', c.name, value=self.c.coerce_const(v, t, tt, c), t=tt)

    def s_redim(self, st):
        base = self.ast.options.get('base', 0)
        for it in st.items:
            if it.target is not None:
                tg = self.unqual(it.target)      # ReDim Preserve Module1.arr(n)
                if tg.kind == 'name':
                    s = self.lookup(tg.name, tg)
                    if s is not None and s.kind == 'var' and s.array:
                        ref = self.arr_ref(s)
                        t = s.t
                        bounds = []
                        for lb, ub in it.dims:
                            lo = self.conv(self.expr(lb), 'Long', lb) if lb is not None else str(base)
                            bounds += [lo, self.conv(self.expr(ub), 'Long', ub)]
                        if is_udt(t):
                            self.emit(f"({ref})->udt = &UD_{t[2:]};")
                        self.emit(f"vb_aredim({ref}, sizeof({ctype(t)}), {akind(t)}, {1 if st.preserve else 0}, "
                                  f"{len(it.dims)}, {', '.join(bounds)});")
                        continue
                if tg.kind == 'member':          # ReDim t.Items(n): a Variant field becomes an array
                    x = self.member(tg, None, tg)
                    if x.code is not None and x.t == 'Variant' and x.ptr is not None and x.arr is None:
                        self.redim_variant(None, it, st.preserve, v=x)
                        continue
                ref, t, _ = self.array_target(it.target, it)
                fixlen = self.target_fixlen
                bounds = []
                for lb, ub in it.dims:
                    lo = self.conv(self.expr(lb), 'Long', lb) if lb is not None else str(base)
                    bounds += [lo, self.conv(self.expr(ub), 'Long', ub)]
                if is_udt(t):
                    self.emit(f"({ref})->udt = &UD_{t[2:]};")
                self.emit(f"vb_aredim({ref}, sizeof({ctype(t)}), {akind(t)}, {1 if st.preserve else 0}, "
                          f"{len(it.dims)}, {', '.join(bounds)});")
                if fixlen:                       # ReDim g.Member(n): Member() As String * 20 in a Type
                    self.emit(f"vb_afix_init({ref}, {fixlen});")
                continue
            s = self.lookup(it.name, it)
            if s is None:
                t = self.c.type_of_decl(self.mk, it.name, it.suffix, it.type, it)
                s = self.new_local(it.name, t, it, array=True, ndims=len(it.dims))
            elif s.kind == 'var' and not s.array and s.t == 'Variant':
                self.redim_variant(s, it, st.preserve)   # Dim v: ReDim v(10) makes it an array
                continue
            elif s.kind != 'var' or not s.array:
                self.err(f"expected array: {it.name}", it)
            elif s.fixed_dims if hasattr(s, 'fixed_dims') else False:
                self.err(f"array already dimensioned: {it.name}", it)
            if it.type is not None and self.c.type_of_decl(self.mk, it.name, it.suffix, it.type, it) != s.t:
                self.err(f"can't change data types of array elements: {it.name}", it)
            if s.ndims is not None and len(it.dims) != s.ndims and st.preserve:
                pass
            bounds = []
            for lb, ub in it.dims:
                lo = self.conv(self.expr(lb), 'Long', lb) if lb is not None else str(base)
                hi = self.conv(self.expr(ub), 'Long', ub)
                bounds += [lo, hi]
            ref = self.arr_ref(s)
            if is_udt(s.t):
                self.emit(f"({ref})->udt = &UD_{s.t[2:]};")
            self.emit(f"vb_aredim({ref}, sizeof({ctype(s.t)}), {akind(s.t)}, "
                      f"{1 if st.preserve else 0}, {len(it.dims)}, {', '.join(bounds)});")
            if getattr(s, 'fixlen', None):
                self.emit(f"vb_afix_init({ref}, {s.fixlen});")

    def redim_variant(self, s, it, preserve, v=None):
        base = self.ast.options.get('base', 0)
        et = self.c.type_of_decl(self.mk, it.name or 'v', it.suffix, it.type, it) if it.type is not None else 'Variant'
        if is_udt(et) or is_obj(et) and et != 'Object':
            et = 'Variant'
        self.ntmp += 1
        a = f"t{self.ntmp}"
        self.decls.append(f"vbarr {a} = {{0}};")
        v = v or self.var_ref(s)
        bounds = []
        for lb, ub in it.dims:
            lo = self.conv(self.expr(lb), 'Long', lb) if lb is not None else str(base)
            bounds += [lo, self.conv(self.expr(ub), 'Long', ub)]
        if preserve:
            self.emit(f"if (vb_isarray(*{v.ptr})) vb_v2arr(*{v.ptr}, &{a}, {vt_of(et)});")
        self.emit(f"vb_aredim(&{a}, sizeof({ctype(et)}), {akind(et)}, {1 if preserve else 0}, "
                  f"{len(it.dims)}, {', '.join(bounds)});")
        self.emit(f"vb_vset({v.ptr}, vb_v_arr(&{a}, {vt_of(et)}));")
        self.emit(f"vb_afree(&{a});")

    def s_erase(self, st):
        for nm in st.names:
            nm = self.unqual(nm)                 # Erase Module1.arr
            if nm.kind == 'name':
                sy = self.lookup(nm.name, nm)
                if sy is not None and sy.kind == 'var' and not sy.array and sy.t == 'Variant':
                    x = self.var_ref(sy)                 # Erase v: a Variant holding an array is emptied
                    self.emit(f"vb_v_erase({x.ptr});")
                    continue
            elif nm.kind == 'member':
                x = self.member(nm, None, nm)
                if x.code is not None and x.t == 'Variant' and x.ptr is not None:   # Erase b(1).Items: a Variant field
                    self.emit(f"vb_v_erase({x.ptr});")
                    continue
            ref, _, _ = self.array_target(nm, nm)
            self.emit(f"vb_aerase({ref});")
            if self.target_fixlen:              # a fixed array of String * n: n times Chr(0) in each (VB6 l_fixarr)
                self.emit(f"vb_afix_init({ref}, {self.target_fixlen});")

    def gosub_jump(self, label, node):
        key = label.lower()
        if key not in self.labels:
            self.err(f"label not defined: {label}", node)
        if not getattr(self, 'gs_declared', False):
            self.gs_declared = True
            self.ngosub = 0
            self.decls.append("vb_i32 vb_gs[256], vb_gsn = 0;")
        self.ngosub += 1
        return self.ngosub, f"if (vb_gsn >= 256) vb_raise(28); else {{ vb_gs[vb_gsn++] = {self.ngosub}; goto L{self.labels[key]}; }}"

    def s_gosub(self, st):
        k, code = self.gosub_jump(st.label, st)
        self.emit(code)
        self.emit(f"G{k}:;")

    def s_return(self, st):
        if not getattr(self, 'gs_declared', False):
            self.gs_declared = True
            self.ngosub = 0
            self.decls.append("vb_i32 vb_gs[256], vb_gsn = 0;")
        self.uses_return = True
        self.emit("if (vb_gsn <= 0) vb_raise(3); else goto vb_gsret;")

    def s_ongoto(self, st):
        v = self.tmp('Long')
        self.emit(f"{v} = {self.conv(self.expr(st.e), 'Long', st)};")
        self.emit(f"if ({v} < 0 || {v} > 255) vb_raise(5); else switch ({v}) {{")
        k = None
        for i, lab in enumerate(st.labels, 1):
            if st.gosub:
                k, code = self.gosub_jump(lab, st)
                self.ngosub -= 1                    # every case comes back to the same place
                code = code.replace(f"vb_gs[vb_gsn++] = {k};", f"vb_gs[vb_gsn++] = {k};")
                self.emit(f"case {i}: {code} break;")
            else:
                key = lab.lower()
                if key not in self.labels:
                    self.err(f"label not defined: {lab}", st)
                self.emit(f"case {i}: goto L{self.labels[key]};")
        self.emit("}")
        if st.gosub:
            self.ngosub += 1
            self.emit(f"G{self.ngosub}:;")

    def s_raiseevent(self, st):
        ci = self.c.classes.get(self.cls) if self.cls else None
        ev = ci.events.get(st.name.lower()) if ci else None
        if ev is None:
            self.err(f"event not found: {st.name}", st)
        r = self.call_proc(ev.sym, st.args, st, as_stmt=True)
        self.emit(r.code + ';')

    # ---------------------------------------------------------------- With, For Each
    def s_with(self, st):
        uc = (st.e.kind == 'name' and st.e.name.lower() == 'usercontrol' and self.cls and self.c.classes[self.cls].is_uctl
              and getattr(self.lookup(st.e.name, st.e), 'kind', None) == 'me')
        if uc or (st.e.kind == 'name' and st.e.name.lower() in ('err', 'app', 'debug') and self.lookup(st.e.name, st.e) is None):
            # With Err / With App: .Number ... are Err.Number ...; With UserControl: the built-in members
            name = st.e

            def sub(n):
                if isinstance(n, list):
                    return [sub(x) for x in n]
                if isinstance(n, tuple):
                    return tuple(sub(x) for x in n)
                if not isinstance(n, N):
                    return n
                if n.kind == 'withref':
                    return N('name', n.line, name=name.name, suffix=None)
                d = {k: v for k, v in n.__dict__.items() if k not in ('kind', 'line')}
                if n.kind == 'with':                 # a With inside: its own .names stay
                    d['e'] = sub(d['e'])
                    return N('with', n.line, **d)
                return N(n.kind, n.line, **{k: sub(v) for k, v in d.items()})
            self.block(sub(st.body))
            return
        x = self.expr(st.e)
        self.flush_pre()
        if is_udt(x.t):
            if x.ptr is None:
                w = self.tmp(x.t)
                self.emit(f"{w} = {x.code};")
                tgt = E(w, x.t, ptr=f"&{w}")
            else:
                w = self.tmp(f"{ctype(x.t)} *", init=False)
                self.emit(f"{w} = {x.ptr};")
                tgt = E(f"(*{w})", x.t, ptr=w)
            free = None
        elif is_obj(x.t) or x.t == 'Variant':
            w = self.tmp(x.t)
            self.emit(self.store_code(E(w, x.t, ptr=f"&{w}"), x, st) if x.t == 'Variant'
                      else f"vb_oset(&{w}, {x.code});")
            free = udt_free_code(x.t, '&' + w)
            self.cleanup.append(free)
            tgt = E(w, x.t)
        else:
            self.err("With needs an object or a user-defined type", st)
        self.check()
        self.withs.append(tgt)
        self.with_nodes = getattr(self, 'with_nodes', []) + [st.e]   # (what the With names: named arguments)
        self.block(st.body)
        self.withs.pop()
        self.with_nodes = self.with_nodes[:-1]
        if free:
            self.emit(free)

    def s_foreach(self, st):
        var = self.lvalue(st.var)
        if var is None or not (var.t == 'Variant' or is_obj(var.t)):
            self.err("For Each control variable must be a Variant or an object", st)
        n = self.nstmt
        lbl = f"F{n}"
        i = self.tmp('vb_i32')
        coll = self.unqual(st.coll)               # For Each f In VB.Forms
        if coll.kind == 'apply' and not coll.args and coll.base.kind in ('name', 'member'):
            coll = coll.base                       # For Each v In arr(): the whole array
        arr = None
        if coll.kind == 'name':
            sy = self.lookup(coll.name, coll)
            if sy is not None and sy.kind == 'var' and sy.array:
                arr = (self.arr_ref(sy), sy.t)
        elif coll.kind == 'member':
            x = self.member(coll, None, coll)
            if x.arr is not None:
                arr = (x.arr, x.t)
        set_ = is_obj(var.t)
        if arr is not None:
            ref, et = arr
            if is_udt(et):
                self.err("For Each over an array of user-defined types is not allowed", st)
            self.check()
            self.emit(f"for ({i} = 0; {i} < ({ref})->n; {i}++) {{")
            elem = E(f"(*(({ctype(et)} *)vb_aelem_lin({ref}, {i})))", et)
        else:
            x = self.expr(coll)
            # an array a function gives (For Each v In Split(s, " ")): its elements, through a Variant holding it
            if not (is_obj(x.t) or x.t == 'Variant' or (x.arr is not None and x.code is None)):
                self.err("For Each needs an array or a collection", st)
            c = self.tmp('Variant')
            self.flush_pre()                       # (the call that makes the array: Split ...)
            self.emit(f"vb_vset(&{c}, {self.conv(x, 'Variant', st)});")
            self.emit(f"vb_each_begin(&{c});")
            self.cleanup.append(f"vb_vfree(&{c});")
            self.emit(f"(void)vb_each_count({c});")     # not a collection (438, 424, 91): an error here
            # VB6 (compile-6.8 c10): with Resume Next an error of For Each goes on INTO the body (once, the
            # variable unchanged), and Next is then error 92 (For loop not initialized)
            bad = self.tmp('int')
            self.nstmt += 1
            m = self.nstmt
            for p in self.post:
                self.emit(p)
            del self.post[:]
            self.emit("vb_release(vb_m0);")
            self.emit(f"{bad} = 0;")
            self.emit(f"if (vb_errpend) {{ {bad} = 1; vb_cur = {m}; goto vb_eh; }}")
            self.emit(f"for ({i} = 0; {bad} || {i} < vb_each_count({c}); {i}++) {{")
            elem = E(f"vb_each_item({c}, {i})", 'Variant')
        self.pre, self.post = [], []
        if arr is None:
            self.emit(f"if (!{bad}) {{")
        if set_:
            self.emit(f"vb_oset({var.ptr}, {self.conv(elem, var.t, st)});")
        else:
            self.emit(self.store_code(var, elem, st))
        self.check()
        if arr is None:
            self.emit("}")
            self.emit(f"S{m}:;")
            self.emit(f"R{m}:;")
        self.loops.append(('for', lbl))
        self.block(st.body)
        self.loops.pop()
        if arr is None:
            self.emit(f"if ({bad}) {{ vb_cur = {n}; vb_raise(92); goto vb_eh; }}")
        self.emit("}")
        # VB6: at the end of the loop the variable is Empty (Nothing); Exit For keeps it (oracle f_inet)
        if set_:
            self.emit(f"vb_oset({var.ptr}, (vbobj)0);")
        else:
            self.emit(f"vb_vset({var.ptr}, vb_v_empty());")
        self.emit(f"{lbl}:;")

    def cond_stmt(self):
        """the condition of If / ElseIf / Do / Loop is a statement of its own for Resume and Resume Next: VB6 goes on
        with the next line, so an error in the condition of If, ElseIf, Do While/Until goes into the block, and one in
        Loop While/Until goes on after the loop (robot 6.8: Jumble, SkinDemo looped on Loop Until EOF(n) of a file
        that was not open, because Resume Next went back before the condition)"""
        self.nstmt += 1
        m = self.nstmt
        self.emit(f"S{m}:;")
        self.emit(f"vb_cur = {m};")
        return m

    def s_if(self, st):
        m = self.cond_stmt()
        c = self.cond(st.cond)
        self.flush_pre()
        f = self.tmp('int')
        self.emit(f"{f} = {c};")
        self.check()
        self.emit(f"if ({f}) {{")
        self.emit(f"R{m}:;")
        self.block(st.then)
        self.emit("}")
        n_close = 0
        for cnd, body in st.elifs:
            self.emit("else {")
            n_close += 1
            self.pre, self.post = [], []
            m = self.cond_stmt()
            c = self.cond(cnd)
            self.flush_pre()
            self.emit(f"{f} = {c};")
            self.check()
            self.emit(f"if ({f}) {{")
            self.emit(f"R{m}:;")
            self.block(body)
            self.emit("}")
        if st.els:
            self.emit("else {")
            self.block(st.els)
            self.emit("}")
        for _ in range(n_close):
            self.emit("}")

    def s_select(self, st):
        subj = self.expr(st.subject)
        self.flush_pre()
        t = subj.t if subj.t in ('String', 'Variant') or is_num(subj.t) else 'Variant'
        v = self.tmp(t)
        if t in ('String', 'Variant'):
            self.emit(self.store_code(E(v, t, ptr=f"&{v}"), subj, st))
            self.cleanup.append(f"vb_sfree(&{v});" if t == 'String' else f"vb_vfree(&{v});")
        else:
            self.emit(f"{v} = {self.conv(subj, t, st)};")
        self.check()
        sv = E(v, t)
        f = self.tmp('int')
        depth = 0
        for i, (conds, body, line) in enumerate(st.cases):
            node = N('case', line)
            if i:
                self.emit("else {")
                depth += 1
            self.pre, self.post = [], []
            tests = []
            for cnd in conds:
                if cnd[0] == 'val':
                    tests.append(self.compare('=', sv, self.expr(cnd[1]), node))
                elif cnd[0] == 'is':
                    tests.append(self.compare(cnd[1], sv, self.expr(cnd[2]), node))
                else:
                    lo = self.compare('>=', sv, self.expr(cnd[1]), node)
                    hi = self.compare('<=', sv, self.expr(cnd[2]), node)
                    tests.append(E(f"({self.truth(lo)} && {self.truth(hi)})", 'Boolean'))
            self.flush_pre()
            self.emit(f"{f} = {' || '.join(self.truth(x) for x in tests)};")
            self.check()
            self.emit(f"if ({f}) {{")
            self.block(body)
            self.emit("}")
        if st.els is not None:
            if st.cases:
                self.emit("else {")
                depth += 1
            self.block(st.els)
            if st.cases:
                pass
            else:
                pass
        self.emit("}" * depth)
        if t in ('String', 'Variant'):
            self.emit(f"vb_sfree(&{v});" if t == 'String' else f"vb_vfree(&{v});")

    def truth(self, x):
        if x.t == 'Variant':
            return f"vb_v_truth({x.code})"
        return f"(({x.code}) != 0)"

    def s_for(self, st):
        var = self.lvalue(st.var)
        if var is None and st.var.kind in ('member', 'apply'):
            x = self.expr(st.var)               # For v.Y = ...: a field of a Type
            if x.ptr is not None and not x.prop:
                var = x
        if var is None or var.t not in RANK and var.t != 'Variant':
            self.err("For loop control variable must be numeric", st)
        t = var.t
        lt = 'Double' if t == 'Variant' else t
        # P-code evaluates the start, the limit and the step; native code the limit, the step, then the start
        te, ts = self.tmp(lt), self.tmp(lt)
        tv = self.tmp(t) if t != 'Variant' else None
        for part in (('start', 'end', 'step') if self.c.order[0] == 'pcode' else ('end', 'step', 'start')):
            if part == 'step':
                step = self.expr(st.step) if st.step is not None else E('((vb_i16)1)', 'Integer')
                self.flush_pre()
                self.emit(f"{ts} = {self.conv(step, lt, st)};")
            elif part == 'end':
                end = self.expr(st.end)
                self.flush_pre()
                self.emit(f"{te} = {self.conv(end, lt, st)};")
            else:
                start = self.expr(st.start)
                self.flush_pre()
                if t == 'Variant':
                    self.emit(f"vb_vset({var.ptr}, {self.conv(start, 'Variant', st)});")
                else:
                    self.emit(f"{tv} = {self.conv(start, t, st)};")
        if tv is not None:
            self.emit(f"if (!vb_errpend) *{var.ptr} = {tv};")
        self.check()
        n = self.nstmt
        cur = f"vb_v2d(*{var.ptr})" if t == 'Variant' else f"(*{var.ptr})"
        lbl = f"F{n}"
        # a label inside: GoTo can come in without the For statement (VB6 P-code: Overflow at Next, 6)
        fi = self.tmp('Long') if self.has_label(st.body) else None   # (Long: declared = 0)
        if fi:
            self.emit(f"{fi} = 1;")
        self.emit(f"while ({ts} >= 0 ? {cur} <= {te} : {cur} >= {te}) {{")
        self.loops.append(('for', lbl))
        self.block(st.body)
        self.loops.pop()
        if fi:
            self.emit(f"vb_cur = {n};")
            self.emit(f"if (!{fi}) {{ vb_raise(6); goto vb_eh; }}")
        if t == 'Variant':
            self.emit(f"vb_vset({var.ptr}, vb_v_add(*{var.ptr}, {self.conv(E(ts, lt), 'Variant')}));")
        else:
            add = self.x_bin_code('+', E(f"(*{var.ptr})", t), E(ts, lt))
            # an error at Next is the For statement's: Resume Next goes on after the loop (VB6)
            self.emit(f"vb_cur = {n};")
            self.emit(f"{{ {ctype(t)} lb_nx = {self.conv(add, t, st)}; if (!vb_errpend) *{var.ptr} = lb_nx; }}")
        self.emit("if (vb_errpend) goto vb_eh;")
        self.emit("}")
        self.emit(f"{lbl}:;")
        if fi:
            self.emit(f"{fi} = 0;")

    def has_label(self, stmts):
        for st in stmts:
            if st.kind == 'label':
                return True
            for attr in ('then', 'els', 'body'):
                sub = getattr(st, attr, None)
                if isinstance(sub, list) and self.has_label(sub):
                    return True
            if st.kind == 'if' and any(self.has_label(b) for _, b in st.elifs):
                return True
            if st.kind == 'select' and any(self.has_label(b) for _, b, _ in st.cases):
                return True
        return False

    def x_bin_code(self, op, a, b):
        return self.x_bin(N('bin', 0, op=op, l=_Pre(a), r=_Pre(b)))

    def x_pre(self, e):
        return e.value

    def s_do(self, st):
        n = self.nstmt
        lbl = f"D{n}"
        self.emit("for (;;) {")
        if st.pre:
            self.pre, self.post = [], []
            m = self.cond_stmt()
            c = self.cond(st.pre[1])
            self.flush_pre()
            f = self.tmp('int')
            self.emit(f"{f} = {c};")
            self.check()
            self.emit(f"if ({'!' if st.pre[0] == 'while' else ''}{f}) break;")
            self.emit(f"R{m}:;")
        self.loops.append(('do', lbl))
        self.block(st.body)
        self.loops.pop()
        mpost = None
        if st.post:
            self.pre, self.post = [], []
            mpost = self.cond_stmt()
            c = self.cond(st.post[1])
            self.flush_pre()
            f = self.tmp('int')
            self.emit(f"{f} = {c};")
            self.check()
            self.emit(f"if ({'!' if st.post[0] == 'while' else ''}{f}) break;")
        self.emit("}")
        self.emit(f"{lbl}:;")
        if mpost:
            self.emit(f"R{mpost}:;")

    def s_exit(self, st):
        w = st.what
        if w in ('sub', 'function', 'property'):
            # VB6 (compile-6.8 c2, c3) makes a program with Exit Function / Exit Sub in a Property
            if w != self.sym.pkind and not self.sym.pkind.startswith('property'):
                self.err(f"Exit {w.title()} not allowed in {self.sym.pkind.title()}", st)
            self.emit("goto vb_exit;")
            return
        for kind, lbl in reversed(self.loops):
            if kind == w:
                self.emit(f"goto {lbl};")
                return
        self.err(f"Exit {w.title()} not within {w.title()}", st)

    def s_goto(self, st):
        key = st.label.lower()
        if key not in self.labels:
            self.err(f"label not defined: {st.label}", st)
        self.emit(f"goto L{self.labels[key]};")

    def s_onerror(self, st):
        self.emit("vb_err_clear();")
        if st.mode == 'resumenext':
            self.emit("vb_ehmode = 1;")
        elif st.mode == 'goto0':
            self.emit("vb_ehmode = 0;")
        elif st.mode == 'goto-1':
            self.emit("vb_inh = 0;")
        else:
            key = st.label.lower()
            if key not in self.labels:
                self.err(f"label not defined: {st.label}", st)
            self.emit(f"vb_ehmode = 2; vb_ehlabel = {self.labels[key]};")

    def s_resume(self, st):
        self.emit("if (!vb_inh) { vb_raise(20); goto vb_eh; }")
        self.emit("vb_inh = 0; vb_err_clear();")
        if st.target == 'next':
            self.emit("goto vb_resume_next;")
        elif st.target == '':
            self.emit("goto vb_resume_same;")
        else:
            key = st.target.lower()
            if key not in self.labels:
                self.err(f"label not defined: {st.target}", st)
            self.emit(f"goto L{self.labels[key]};")

    def s_end(self, st):
        self.emit("vb_end();")

    def s_stop(self, st):
        self.emit("vb_dbg_stop(&vb_dbgf);" if self.c.debug else "vb_stop();")

    def s_errorstmt(self, st):
        self.emit(f"vb_err_raise({self.conv(self.expr(st.number), 'Long', st)}, vb_stmp(0), vb_stmp(0), 0, 0);")

    def s_randomize(self, st):
        if st.seed is None:
            self.emit("vb_randomize(0, 0);")
        else:
            self.emit(f"vb_randomize({self.conv(self.expr(st.seed), 'Double', st)}, 1);")

    def s_print(self, st):
        if getattr(st, 'debug', False) and not self.c.keep_debug:
            return                               # Debug.Print: a made program has none (VB6), its arguments either
        f = '0' if st.file is None else self.conv(self.expr(st.file), 'Long', st)
        self.emit(f"vb_pr_begin({f});")
        newline = self.pr_items(st)
        self.emit(f"vb_pr_end({1 if newline else 0});")

    def s_gmethod(self, st):
        """graphics methods of forms and PictureBoxes: Print, PSet, Line, Circle, Scale"""
        if st.obj is None:
            if not (self.cls and self.c.classes[self.cls].is_form):
                self.err(f"method not valid without suitable object: {st.name.capitalize()}", st)
            b = E('vb_self', 'O:' + self.cls)
        else:
            b = self.expr(st.obj)
            if not (is_obj(b.t) or b.t == 'Variant'):
                self.err(f"invalid qualifier: {self.describe_t(b.t)}", st)
        call = 'vb_late_o' if is_obj(b.t) else 'vb_late'
        if st.name == 'print':
            self.emit("vb_pr_begin(-2);")
            nl = self.pr_items(st)
            self.emit(f"vb_pr_end({1 if nl else 0});")
            self.emit(f"{{ vb_i32 nl_; vbs s_ = vb_pr_take(&nl_); "
                      f"{call}({b.code}, \"$print\", VBM_METHOD, 2, (vbvar[]){{ vb_v_s(s_), vb_v_bool(nl_ ? -1 : 0) }}); }}")
            return
        V = lambda x: 'vb_v_missing()' if x is None else self.conv(self.expr(x), 'Variant', st)
        S = lambda f: f"vb_v_bool({-1 if f else 0})"
        if st.name == 'pset':
            vals = [S(st.step), V(st.x), V(st.y), V(st.color)]
        elif st.name == 'line':
            p1 = st.p1 or (False, None, None)
            vals = [S(st.p1 is not None), S(p1[0]), V(p1[1]), V(p1[2]), S(st.p2[0]), V(st.p2[1]), V(st.p2[2]),
                    V(st.color), f"vb_v_i({st.box})"]
        elif st.name == 'circle':
            ex = list(st.extra) + [None] * (4 - len(st.extra))
            vals = [S(st.step), V(st.x), V(st.y), V(st.r)] + [V(x) for x in ex[:4]]
        else:   # scale
            if st.p1 is None:
                vals = []
            else:
                vals = [V(st.p1[1]), V(st.p1[2]), V(st.p2[1]), V(st.p2[2])]
        argv = f"{len(vals)}, (vbvar[]){{ {', '.join(vals)} }}" if vals else "0, 0"
        self.emit(f"{call}({b.code}, \"${st.name}\", VBM_METHOD, {argv});")

    def pr_items(self, st):
        newline = True
        for it in st.items:
            newline = it[0] not in (';', ',')
            if it[0] == ';':
                continue
            if it[0] == ',':
                self.emit("vb_pr_comma();")
            elif it[0] == 'spc':
                self.emit(f"vb_pr_spc({self.conv(self.expr(it[1]), 'Long', st)});")
            elif it[0] == 'tab':
                self.emit(f"vb_pr_tab({self.conv(self.expr(it[1]), 'Long', st) if it[1] is not None else '0'});")
            else:
                x = self.expr(it[1])
                if is_obj(x.t):
                    x = self.cv(x, 'Variant', st)
                if is_udt(x.t):
                    self.err("type mismatch: a Type cannot be printed", st)
                fn = {'String': 'vb_pr_s', 'Integer': 'vb_pr_i', 'Long': 'vb_pr_i', 'Byte': 'vb_pr_i',
                      'Single': 'vb_pr_f', 'Double': 'vb_pr_d', 'Currency': 'vb_pr_c', 'Boolean': 'vb_pr_bool',
                      'Date': 'vb_pr_date', 'Variant': 'vb_pr_v'}.get(x.t)
                if fn is None:
                    self.err("invalid use of a procedure in Print", st)
                self.emit(f"{fn}({x.code});")
        return newline

    def s_write(self, st):
        f = self.conv(self.expr(st.file), 'Long', st)
        self.emit(f"vb_wr_begin({f});")
        for x in st.exprs:
            self.emit(f"vb_wr_v({self.conv(self.expr(x), 'Variant', st)});")
        self.emit("vb_wr_end();")

    def s_open(self, st):
        mode = {'input': 'VBF_INPUT', 'output': 'VBF_OUTPUT', 'append': 'VBF_APPEND', 'binary': 'VBF_BINARY',
                'random': 'VBF_RANDOM'}[st.mode]
        rl = self.conv(self.expr(st.reclen), 'Long', st) if st.reclen is not None else '0'
        self.emit(f"vb_open({self.conv(self.expr(st.path), 'String', st)}, {mode}, "
                  f"{self.conv(self.expr(st.file), 'Long', st)}, {rl});")

    def getput_target(self, st):
        """(code of the call arguments for vb_get/vb_put) of the target variable"""
        tg = st.target
        if tg.kind == 'apply' and not tg.args and tg.base.kind in ('name', 'member'):   # Get #1, , b()
            tg = tg.base
        arr = None
        if tg.kind == 'name':
            sy = self.lookup(tg.name, tg)
            if sy is not None and sy.kind == 'var' and sy.array:
                arr = (self.arr_ref(sy), sy.t, not getattr(sy, 'fixed_dims', None), getattr(sy, 'fixlen', None))
        elif tg.kind == 'member':
            x = self.member(tg, None, tg)
            if x.arr is not None:
                arr = (x.arr, x.t, True, x.fix)
        if arr is not None:
            ref, t, dynamic, fixlen = arr       # (an array of String * n: each element n bytes, as VB6 writes it)
            udt = f"&UD_{t[2:]}" if is_udt(t) else '0'
            return 'arr', f"{self.c.vbk(t, fixlen)}, {fixlen or 0}, {udt}, {ref}, {1 if dynamic else 0}"
        x = self.expr(tg)
        if x.code is None and x.arr is not None and st.kind == 'put':    # Put #1, , ReadFile(f): the array it gives
            t = x.t[2:] if is_arr(x.t) else x.t
            udt = f"&UD_{t[2:]}" if is_udt(t) else '0'
            return 'arr', f"{self.c.vbk(t)}, 0, {udt}, {x.arr}, 1"
        if (x.ptr is None or x.prop) and st.kind == 'put' and not is_udt(x.t) and not is_obj(x.t) and \
                x.arr is None and x.t in ('String', 'Variant', 'Integer', 'Long', 'Byte', 'Boolean', 'Single',
                                          'Double', 'Currency', 'Date'):
            tmp = self.tmp(x.t)                 # Put #1, , "PNG" & s: the value of an expression
            self.emit(self.store_code(E(tmp, x.t, ptr=f"&{tmp}"), x, st))
            self.free_after(tmp, x.t, self.cleanup)
            x = E(tmp, x.t, ptr=f"&{tmp}")
        if x.ptr is None or x.prop:
            self.err("Get/Put needs a variable", st)
        if is_obj(x.t):
            self.err("Get/Put cannot write objects", st)
        udt = f"&UD_{x.t[2:]}" if is_udt(x.t) else '0'
        return 'val', f"{self.c.vbk(x.t, x.fix)}, {x.fix or 0}, {udt}, {x.ptr}"

    def s_get(self, st):
        self.getput(st, 'get')

    def s_put(self, st):
        self.getput(st, 'put')

    def getput(self, st, op):
        f = self.conv(self.expr(st.file), 'Long', st)
        rec = f"vb_recno({self.conv(self.expr(st.rec), 'Long', st)})" if st.rec is not None else '0'
        kind, rest = self.getput_target(st)
        self.emit(f"vb_{op}{'_arr' if kind == 'arr' else ''}({f}, {rec}, {rest});")

    def s_seek(self, st):
        self.emit(f"vb_seek({self.conv(self.expr(st.file), 'Long', st)}, {self.conv(self.expr(st.pos), 'Long', st)});")

    def s_close(self, st):
        if not st.files:
            self.emit("vb_close(0);")
        for f in st.files:
            self.emit(f"vb_close({self.conv(self.expr(f), 'Long', st)});")

    def assign_read(self, tg, tmp, t, st):
        """Input # / Line Input # into a property or a field (Text1.Text, x(i).Name): as an assignment"""
        self.s_assign(N('assign', st.line, target=tg, value=_Pre(E(tmp, t)), set=False))

    def input_target(self, tg):
        """the variable Input # / Line Input # reads into, or None: through an assignment (a property, a
        field, v(i) of a Variant that holds an array)"""
        if tg.kind == 'apply' and tg.base.kind == 'name':
            sy = self.lookup(tg.base.name, tg.base)
            if sy is not None and sy.kind == 'var' and not sy.array and sy.t == 'Variant':
                return None
        return self.lvalue(tg)

    def s_lineinput(self, st):
        t = self.input_target(st.target)
        f = self.conv(self.expr(st.file), 'Long', st)
        if t is None:
            tmp = self.tmp('String')
            self.emit(f"vb_line_input({f}, &{tmp});")
            self.emit("if (!vb_errpend) {")
            self.assign_read(st.target, tmp, 'String', st)
            self.emit("}")
            self.cleanup.append(f"vb_sfree(&{tmp});")
            return
        if t.t == 'String' and not t.fix:
            self.emit(f"vb_line_input({f}, {t.ptr});")
        else:
            tmp = self.tmp('String')
            self.emit(f"vb_line_input({f}, &{tmp});")
            self.emit(f"if (!vb_errpend) {self.store_code(t, E(tmp, 'String'), st)}")
            self.cleanup.append(f"vb_sfree(&{tmp});")

    def s_input(self, st):
        f = self.conv(self.expr(st.file), 'Long', st)
        for tg in st.targets:
            t = self.input_target(tg)
            if t is None:                       # a property or a field: a Variant, then assigned
                tmp = self.tmp('Variant')
                self.emit(f"vb_input_v({f}, &{tmp});")
                self.emit("if (!vb_errpend) {")
                self.assign_read(tg, tmp, 'Variant', st)
                self.emit("}")
                self.cleanup.append(f"vb_vfree(&{tmp});")
                continue
            if t.t == 'String':
                self.emit(f"vb_input_s({f}, {t.ptr});")
            elif t.t == 'Variant':
                self.emit(f"vb_input_v({f}, {t.ptr});")
            else:
                d = self.tmp('Double')
                self.emit(f"vb_input_d({f}, &{d});")
                self.emit(f"if (!vb_errpend) {self.store_code(t, E(d, 'Double'), st)}")

    def s_lset(self, st):
        """LSet / RSet s = v: v in the length of s, left or right aligned; LSet a = b of two Types:
        the bytes of b over those of a (in the layout of VB6)"""
        t = self.expr(st.target)
        v = self.expr(st.value)
        if is_udt(t.t):
            if not is_udt(v.t) or t.ptr is None:
                self.err("type mismatch", st)
            if st.right:
                self.err("RSet works with strings only", st)
            src = v.ptr if v.ptr is not None else f"&{self.udt_value(v)}"
            self.emit(f"vb_udt_lset((void *)({t.ptr}), &UD_{t.t[2:]}, (const void *)({src}), &UD_{v.t[2:]});")
            return
        if t.t != 'String' or t.ptr is None:
            self.err("LSet / RSet needs a String variable", st)
        self.emit(f"vb_slset({t.ptr}, {self.conv(v, 'String', st)}, {1 if st.right else 0});")

    def s_midstmt(self, st):
        keys, snap = [], self.snapshot()        # Mid$(s, a, b) = c: a, b, then c
        try:
            if self.stmt_hoist(None, keys, args=[st.start, st.length, st.value]):
                return self.midstmt(st)
            self.rollback(snap)
        finally:
            for k in keys:
                self.sched.pop(k, None)
        return self.midstmt(st)

    def midstmt(self, st):
        t = self.lvalue(st.target)
        if t is None:                            # Mid(udt.Field, 1) = s: a member
            t = self.expr(st.target)
        if t.t == 'Variant' and t.ptr is not None:   # Mid(v, 2) = "x": a Variant holding a String
            w = self.tmp('String')
            self.emit(f"vb_sset(&{w}, vb_v2s(*{t.ptr}));")
            n = self.conv(self.expr(st.length), 'Long', st) if st.length is not None else '-1'
            fn = 'vb_midb_stmt' if getattr(st, 'byte', False) else 'vb_mid_stmt'
            self.emit(f"{fn}(&{w}, {self.conv(self.expr(st.start), 'Long', st)}, {n}, "
                      f"{self.conv(self.expr(st.value), 'String', st)});")
            self.emit(f"vb_vset({t.ptr}, vb_v_s({w}));")
            self.free_after(w, 'String', self.cleanup)
            return
        if t.t not in ('String',) or t.ptr is None:
            self.err("Mid statement needs a String variable", st)
        n = self.conv(self.expr(st.length), 'Long', st) if st.length is not None else '-1'
        fn = 'vb_midb_stmt' if getattr(st, 'byte', False) else 'vb_mid_stmt'
        self.emit(f"{fn}({t.ptr}, {self.conv(self.expr(st.start), 'Long', st)}, {n}, "
                  f"{self.conv(self.expr(st.value), 'String', st)});")


class ThunkGen:
    """late-binding entry points of the public members of a class:
    vbvar TH_x(vbobj self, vb_i32 argc, vbvar *argv)"""

    def __init__(self, comp, ci):
        self.c = comp
        self.ci = ci
        self.k = cident(ci.mk)
        self.n = 0
        ast = dict((m.lower(), a) for m, a in comp.modules)[ci.mk]
        pg = ProcGen.__new__(ProcGen)
        pg.c, pg.ast, pg.mk = comp, ast, ci.mk
        pg.node = ast
        self.pg = pg

    def name(self, base):
        self.n += 1
        return f"TH_{self.k}_{self.n}_{base}"

    def field(self, f, mode):
        fn = self.name(f.cname)
        fld = f"((C_{self.k} *)self)->{f.cname}"
        pg = self.pg
        if mode == 'get':
            val = f"vb_oauto(&{fld}, {self.c.cls_ref(f.autonew)})" if getattr(f, 'autonew', None) else fld
            body = [f"if (argc) {{ vb_raise(450); return vb_v_empty(); }}",
                    f"return vb_vtmp({pg.conv(E(val, f.t), 'Variant')});"]
        else:
            arg = E('argv[0]', 'Variant')
            if mode == 'set':
                st = (f"vb_vset(&{fld}, vb_v_obj(vb_v2obj(argv[0])));" if f.t == 'Variant'
                      else f"vb_oset(&{fld}, {pg.conv(arg, f.t)});")
            else:
                st = pg.store_code(E(fld, f.t, ptr=f"&{fld}"), arg, None)
            body = ["if (argc != 1) { vb_raise(450); return vb_v_empty(); }", st, "return vb_v_empty();"]
        self.c.class_support.append(f"static vbvar {fn}(vbobj self, vb_i32 argc, vbvar *argv) {{ "
                                    f"(void)argv; {' '.join(body)} }}")
        return fn

    def proc(self, m):
        fn = self.name(m.cname)
        m = self.ci.dispatch.get(m.cname, m)        # an interface: go to the implementation
        pg = self.pg
        pg.node = m.node
        decls, pre, post, cargs = [], [], [], ['self']
        params = list(m.params)
        pa = params and getattr(params[-1], 'paramarray', False)
        if pa:                                      # ParamArray: the arguments from there on, as an array
            params = params[:-1]
        nreq = sum(1 for p in params if not p.optional)
        for i, p in enumerate(params):
            if p.array or (is_udt(p.t) and not self.c.rec_ok(p.t)) or (is_udt(m.rt or '') and not self.c.rec_ok(m.rt)):
                self.c.class_support.append(f"static vbvar {fn}(vbobj self, vb_i32 argc, vbvar *argv) {{ "
                                            f"(void)self; (void)argc; (void)argv; vb_raise(13); return vb_v_empty(); }}")
                return fn
            present = f"(argc > {i} && !vb_ismissing(argv[{i}]))"
            if is_udt(p.t):                     # a Public Type (ByRef): from the Variant, its new value back
                k, t = p.t[2:], f"a{i}"
                decls.append(f"U_{k} {t} = {{0}};")     # (zeroed: _init sets up only its arrays)
                pre.append(f"U_{k}_init(&{t}); if ({present}) vb_v2rec(argv[{i}], &UD_{k}, &{t});")
                post.append(f"if (argc > {i}) vb_vset(&argv[{i}], vb_v_rec(&{t}, &UD_{k})); U_{k}_free(&{t});")
                cargs.append(f"&{t}")
                continue
            if p.default is not None:
                dflt = self.c.const_code(*p.default)
                dflt_t = p.default[1]
            elif p.t == 'Variant':
                dflt, dflt_t = 'vb_v_missing()', 'Variant'
            else:
                v, dflt_t = default_of(p.t)
                dflt = '((vbobj)0)' if is_obj(p.t) else self.c.const_code(v, dflt_t)
            dv = pg.conv(E(dflt, dflt_t), p.t) if not is_obj(p.t) else dflt
            val = f"({present} ? {pg.conv(E(f'argv[{i}]', 'Variant'), p.t)} : {dv})"
            if p.byval:
                cargs.append(val)
            elif p.t == 'Variant':
                t = f"a{i}"
                decls.append(f"vbvar {t} = {{0}};")
                pre.append(f"if ({present}) {t}p = &argv[{i}]; else {{ vb_vset(&{t}, {dflt}); {t}p = &{t}; }}")
                decls.append(f"vbvar *{t}p;")
                post.append(f"vb_vfree(&{t});")
                cargs.append(f"{t}p")
            else:
                t = f"a{i}"
                decls.append(f"{ctype(p.t)} {t} = {zero(p.t)};")
                pre.append(pg.store_code(E(t, p.t, ptr=f"&{t}"), E(val, p.t), None))
                # ByRef: the new value goes back to the caller's argument (Cancel = True in an event)
                post.append(f"if (argc > {i}) vb_vset(&argv[{i}], {pg.conv(E(t, p.t), 'Variant')});")
                c = udt_free_code(p.t, '&' + t)
                if c:
                    post.append(c)
                cargs.append(f"&{t}")
        if pa:
            n = len(params)
            decls.append("vbarr pa = { 0 };")
            pre.append(f"vb_aparam(&pa, argc > {n} ? argc - {n} : 0, argc > {n} ? argv + {n} : 0);")
            post.append("vb_afree(&pa);")
            cargs.append("&pa")
        call = f"{m.cname}({', '.join(cargs)})"
        too_many = "" if pa else f" || argc > {len(params)}"
        body = decls + [f"if (argc < {nreq}{too_many}) {{ vb_raise(450); return vb_v_empty(); }}"] + pre
        if self.c.ctx and not self.ci.is_form:   # the call runs in the copy of the server the object belongs to
            body.insert(len(decls), "void *lb_sv = vb_ctx_enter(self, &LB_CTX);")
            body[len(decls) + 1] = body[len(decls) + 1].replace("{ vb_raise(450); return vb_v_empty(); }",
                                                                  "{ vb_raise(450); vb_ctx_end(&LB_CTX, lb_sv); return vb_v_empty(); }")
            post = post + ["vb_ctx_end(&LB_CTX, lb_sv);"]
        if self.c.component:                    # an error going back to the client: Source = this project
            post = [f'if (vb_errpend) vb_err_source_default("{c_utf8(self.c.project_name)}");'] + post
        if m.rt and is_arr(m.rt):               # a Function returning an array: a Variant holding it
            body.append(f"vbarr r = {call};")
            body += post
            body.append(f"{{ vbvar v = vb_v_arr(&r, {vt_of(m.rt[2:])}); vb_afree(&r); return v; }}")
        elif m.rt and is_udt(m.rt):             # a Function giving a Public Type: a Variant holding it
            body.append(f"{ctype(m.rt)} r = {call};")
            body += post
            body.append(f"{{ vbvar v = vb_v_rec(&r, &UD_{m.rt[2:]}); U_{m.rt[2:]}_free(&r); return v; }}")
        elif m.rt:
            body.append(f"{ctype(m.rt)} r = {call};")
            body += post
            body.append(f"return {pg.conv(E('r', m.rt), 'Variant')};")
        else:
            body.append(call + ';')
            body += post
            body.append("return vb_v_empty();")
        self.c.class_support.append(f"static vbvar {fn}(vbobj self, vb_i32 argc, vbvar *argv) {{ "
                                    f"(void)argv; {' '.join(body)} }}")
        return fn


def has_onerror(stmts):
    """does the procedure contain an On Error statement"""
    stack = list(stmts)
    while stack:
        x = stack.pop()
        if isinstance(x, list):
            stack.extend(x)
        elif isinstance(x, N):
            if x.kind == 'onerror':
                return True
            stack.extend(v for v in x.__dict__.values() if isinstance(v, (list, N)))
    return False


class _Pre(N):
    """an already compiled expression inside a synthesized node"""

    def __init__(self, e):
        super().__init__('pre', 0, value=e)


def op_c(op):
    return {'=': '==', '<>': '!=', '<': '<', '>': '>', '<=': '<=', '>=': '>='}[op]


def default_of(t):
    if t == 'String':
        return '', 'String'
    if t == 'Variant':
        return None, 'Variant'
    return 0, t


# the object types of the VB library: forms and the intrinsic controls
VB_OBJECT_TYPES = {'control', 'form', 'mdiform', 'commandbutton', 'label', 'textbox', 'checkbox', 'optionbutton', 'listbox',
                   'combobox', 'timer', 'frame', 'picturebox', 'image', 'shape', 'line', 'hscrollbar', 'vscrollbar',
                   'drivelistbox', 'dirlistbox', 'filelistbox', 'menu', 'data', 'ole', 'stdfont', 'stdpicture',
                   'picture', 'font', 'propertybag', 'ambientproperties', 'usercontrol', 'parentcontrols',
                   'containedcontrols', 'extender', 'datamembers', 'databindings', 'databinding', 'licenses',
                   'licenseinfo', 'dataobject', 'dataobjectfiles', 'ipicturedisp', 'ipicture', 'ifontdisp', 'ifont',
                   'errobject', 'printer', 'asyncproperty', 'vbcontrolextender', 'screen', 'app', 'clipboard', 'global',
                   'propertypage', 'userdocument', 'selectedcontrols', 'hyperlink'}

# the numeric types of stdole that user controls use (OLE_COLOR ...)
OLE_TYPES = {'ole_color': 'Long', 'ole_handle': 'Long', 'ole_tristate': 'Long', 'ole_optexclusive': 'Variant',
             'ole_cancelbool': 'Boolean', 'ole_enabledefaultbool': 'Boolean',
             **{f'ole_{a}{b}_pixels': 'Long' for a in 'xy' for b in ('pos', 'size')},
             **{f'ole_{a}{b}_himetric': 'Long' for a in 'xy' for b in ('pos', 'size')},
             **{f'ole_{a}{b}_container': 'Single' for a in 'xy' for b in ('pos', 'size')}}


def unescape_bytes(text):
    """the bytes the code page of the source does not define (kept as surrogates when the file was
    read): the character of the same number, as Windows reads them (0x8D -> U+008D)"""
    return re.sub('[\udc80-\udcff]', lambda m: chr(ord(m.group()) - 0xDC00), text)


def c_utf8(text):
    """C string literal contents (UTF-8, escaped) for text"""
    out = []
    for b in unescape_bytes(text).encode('utf-8'):
        if 32 <= b < 127 and chr(b) not in '"\\?':
            out.append(chr(b))
        else:
            out.append('\\%03o' % b)
    return ''.join(out)
