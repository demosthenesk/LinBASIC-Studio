# LinBASIC Studio - the designers of VB6 (.dsr) as modules the compiler knows:
#   DataEnvironment {C0E45035-5775-11D0-B388-00A0C9055D8E}: a class with a default instance (deX), written
#     here in VB from the connections and commands of the .dsr, with the code of the .dsr after it: its
#     Connection objects (WithEvents), one method and one rsX recordset for every command at the top
#     of a hierarchy, the SHAPE command of MSDataShape of a hierarchy (relations, aggregates,
#     grouping), Commands / Connections / Recordsets, GetDataMember for the bound controls and the
#     DataReport. As VB6 does it (VB6 oracle tests/db/d_de1).
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.

DE_GUID = '{C0E45035-5775-11D0-B388-00A0C9055D8E}'
REPORT_GUID = '{78E93846-85FD-11D0-8487-00A0C90DC8A9}'
ADDIN_GUID = '{AC0714F6-3D04-11D1-AE7D-00A0C90F26F4}'     # the Add-In designer (Connect of an add-in project)

# AggFunction of the .dsr (3 COUNT: the VB98 sample DataRept)
AGG_FUNCS = {0: '', 1: 'ANY', 2: 'AVG', 3: 'COUNT', 4: 'MAX', 5: 'MIN', 6: 'STDEV', 7: 'SUM'}


def designer_kind(sf):
    g = (sf.root.type_name if sf.root is not None else '').upper()
    if g == DE_GUID:
        return 'DataEnvironment'
    if g == REPORT_GUID:
        return 'DataReport'
    if g == ADDIN_GUID:
        return 'AddIn'
    return None


def _props(block):
    """the properties of a PropertyBlock / Control: lower name -> value (the first of a name)"""
    out = {}
    for it in block.items:
        if type(it).__name__ == 'Property':
            out.setdefault(it.name.lower(), it.value)
    return out


def _blocks(block, prefix):
    return [it for it in block.items if type(it).__name__ == 'PropertyBlock' and it.name.lower().startswith(prefix.lower())
            and it.name[len(prefix):].isdigit()]


def _s(v):
    """a VB string literal"""
    return '"' + str('' if v is None else v).replace('"', '""') + '"'


class DeCommand:
    def __init__(self, blk):
        p = _props(blk)
        self.p = p
        self.name = str(p.get('commandname', ''))
        self.text = str(p.get('commandtext', ''))
        self.type = int(p.get('commandtype', 1) or 1)
        self.conn = str(p.get('activeconnectionname', ''))
        self.parent = str(p.get('parentcommandname', '') or '') if int(p.get('relatetoparent', 0) or 0) else ''
        self.grouping = int(p.get('grouping', 0) or 0) != 0
        self.grouping_name = str(p.get('groupingname', '') or self.name + '_Grouping')
        self.dispid = int(p.get('commdispid', -1) or -1)
        self.returns = int(p.get('isrsreturning', -1) or 0) != 0
        self.groups = [str(_props(b).get('name', '')) for b in _blocks(blk, 'Grouping')]
        self.relations = [(str(_props(b).get('parentfield', '')), str(_props(b).get('childfield', ''))) for b in _blocks(blk, 'Relation')]
        # ChildType 1: the child field is a parameter of the command (VB6 oracle d_de2)
        self.relparam = [int(_props(b).get('childtype', 0) or 0) == 1 for b in _blocks(blk, 'Relation')]
        # the parameters of the command: (name, real name, direction, size, ADO type, precision, scale)
        self.params = []
        for b in _blocks(blk, 'P'):
            pp = _props(b)
            nm = str(pp.get('username', '') or pp.get('realname', '') or f'Param{len(self.params) + 1}')
            self.params.append((nm, str(pp.get('realname', '') or nm), int(pp.get('direction', 1) or 1),
                                int(pp.get('size', 0) or 0), int(pp.get('datatype', 12) or 12),
                                int(pp.get('precision', 0) or 0), int(pp.get('scale', 0) or 0)))
            self.hosttypes = getattr(self, 'hosttypes', []) + [int(pp.get('hosttype', 12) or 12)]
        self.aggs = []
        for b in _blocks(blk, 'Aggregate'):
            ap = _props(b)
            self.aggs.append((str(ap.get('name', '')), str(ap.get('aggon', '')), str(ap.get('aggfield', '')),
                              int(ap.get('aggtype', 1) or 0), int(ap.get('aggfunction', 0) or 0)))
        self.cursortype = p.get('cursortype')
        self.locktype = p.get('locktype')
        self.cursorloc = p.get('cursorlocation')
        self.children = []

    def sql(self):
        if self.type == 2:                      # a table
            return 'SELECT * FROM `' + self.text + '`'
        if self.type == 4:                      # a stored procedure
            return '{call ' + self.text + '}'
        return self.text

    def hierarchical(self):
        return bool(self.children or self.aggs)

    def param_index(self, name):
        """a relation to a parameter of the command: RELATE ... TO PARAMETER n"""
        for i, prm in enumerate(self.params):
            if name.lower() in (prm[0].lower(), prm[1].lower()):
                return i
        return -1


def shape_text(c):
    """the SHAPE command of MSDataShape for the command c and its children, as the DataEnvironment
    writes it (VB6 oracle d_de1)"""
    def node(x):
        if not x.hierarchical():
            return '{' + x.sql() + '}  AS ' + x.name
        parts = []
        for k in x.children:
            inner = '(' + node(k) + ') AS ' + k.name if k.hierarchical() else node(k)
            rel = ', '.join(f"'{p}' TO  PARAMETER {max(k.param_index(q), 0)}" if isp else f"'{p}' TO '{q}'"
                            for (p, q), isp in zip(k.relations, k.relparam))
            parts.append(f'({inner} RELATE {rel}) AS {k.name}')
        for name, on, fld, _typ, fn in x.aggs:
            parts.append(f"{AGG_FUNCS.get(fn, '')}({on}.'{fld}') AS {name}")
        return ' SHAPE {' + x.sql() + '}  AS ' + x.name + ' APPEND ' + ', '.join(parts)
    if c.grouping:
        by = ', '.join(f"'{g}'" for g in c.groups)
        return ' SHAPE (' + node(c) + ') COMPUTE ' + c.name + ' BY ' + by
    if c.hierarchical():
        return node(c)
    return c.sql()


def dataenv_source(sf):
    """the VB code of the class of a DataEnvironment .dsr, and the class of its collections:
    [(module name, code lines, attributes)]"""
    name = sf.name
    root = sf.root
    conns = []
    for b in _blocks(root, 'Connection'):
        p = _props(b)
        conns.append((str(p.get('connectionname', '')), str(p.get('connectionsource', '') or ''), p))
    cmds = [DeCommand(b) for b in _blocks(root, 'Recordset')]
    byname = {c.name.lower(): c for c in cmds}
    for c in cmds:
        if c.parent and c.parent.lower() in byname:
            byname[c.parent.lower()].children.append(c)
    top = [c for c in cmds if not c.parent]
    user = list(sf.code)
    opts = []
    for i, ln in enumerate(user):              # Option ... first; the code keeps its line numbers
        if ln.strip().lower().startswith('option '):
            opts.append(ln)
            user[i] = ''
    lowcode = '\n'.join(user).lower()

    def has(proc):
        import re
        return re.search(r'(?m)^\s*(?:private\s+|public\s+)?sub\s+' + proc.lower() + r'\s*\(', lowcode) is not None

    coll = f'lbde_{name}_Coll'
    L = list(opts)
    L.append("' ---- made by LinBASIC Studio from the DataEnvironment designer")
    for cn, _, _ in conns:
        L.append(f'Public WithEvents {cn} As ADODB.Connection')
    for c in top:
        L.append(f'Public WithEvents rs{c.grouping_name if c.grouping else c.name} As ADODB.Recordset')
    for c in top:
        L.append(f'Private lb_cmd_{c.name} As ADODB.Command')
    L += [f'Private lb_commands As {coll}', f'Private lb_connections As {coll}', f'Private lb_recordsets As {coll}',
          'Private lb_ready As Boolean', '']
    decls, L = L, []
    L += ['Private Sub Class_Initialize()',
          f'    Set lb_commands = New {coll}', f'    Set lb_connections = New {coll}', f'    Set lb_recordsets = New {coll}']
    for cn, src, p in conns:
        L.append(f'    Set {cn} = New ADODB.Connection')
        L.append(f'    {cn}.ConnectionString = {_s(src)}')
        if p.get('connectiontimeout') is not None:
            L.append(f'    {cn}.ConnectionTimeout = {p.get("connectiontimeout")}')
        if p.get('commandtimeout') is not None:
            L.append(f'    {cn}.CommandTimeout = {p.get("commandtimeout")}')
        L.append(f'    lb_connections.lb_add {cn}, {_s(cn)}')
    for c in top:
        member = c.grouping_name if c.grouping else c.name
        L.append(f'    Set lb_cmd_{c.name} = New ADODB.Command')
        L.append(f'    lb_cmd_{c.name}.CommandText = {_s(shape_text(c))}')
        L.append(f'    lb_cmd_{c.name}.CommandType = {1 if (c.grouping or c.hierarchical()) else c.type}')
        L.append(f'    lb_cmd_{c.name}.Name = {_s(member)}')
        if not (c.grouping or c.hierarchical()):
            for nm, _real, dirn, size, typ, prec, scale in c.params:
                L.append(f'    lb_cmd_{c.name}.Parameters.Append lb_cmd_{c.name}.CreateParameter({_s(nm)}, {typ}, {dirn}, {size})')
                if prec or scale:
                    L.append(f'    lb_cmd_{c.name}.Parameters({_s(nm)}).Precision = {prec}')
                    L.append(f'    lb_cmd_{c.name}.Parameters({_s(nm)}).NumericScale = {scale}')
        L.append(f'    Set rs{member} = New ADODB.Recordset')
        # the recordset keeps the designer's cursor until the program changes it: the command opens it as
        # it is (the Data Object Wizard's classes set LockType before de.Command: VB6 oracle do1)
        L.append(f'    rs{member}.CursorLocation = {c.cursorloc if c.cursorloc is not None else 3}')
        L.append(f'    rs{member}.CursorType = {c.cursortype if c.cursortype is not None else 3}')
        L.append(f'    rs{member}.LockType = {c.locktype if c.locktype is not None else 1}')
        L.append(f'    lb_commands.lb_add lb_cmd_{c.name}, {_s(member)}')
        L.append(f'    lb_recordsets.lb_add rs{member}, {_s(member)}')
    if has('DataEnvironment_Initialize'):
        L.append('    DataEnvironment_Initialize')
    for cn, _, _ in conns:                      # the connections go through MSDataShape
        L.append(f'    {cn}.ConnectionString = lb_shape_cs({cn}.ConnectionString)')
        L.append(f'    {cn}.CursorLocation = 3')
    L += ['End Sub', '']
    L += ['Private Sub Class_Terminate()']
    if has('DataEnvironment_Terminate'):
        L.append('    DataEnvironment_Terminate')
    L += ['End Sub', '']
    L += ['Private Function lb_shape_cs(ByVal cs As String) As String',
          '    Dim parts() As String, i As Integer, prov As String, rest As String, k As String',
          '    If LCase$(Left$(cs, 21)) = "provider=msdatashape" Or InStr(1, cs, "provider=msdatashape", 1) > 0 Then lb_shape_cs = cs: Exit Function',
          '    parts = Split(cs, ";")',
          '    For i = 0 To UBound(parts)',
          '        k = LCase$(Trim$(Left$(parts(i), InStr(parts(i) & "=", "=") - 1)))',
          '        If k = "provider" Then',
          '            prov = Trim$(Mid$(parts(i), InStr(parts(i), "=") + 1))',
          '        ElseIf Trim$(parts(i)) <> "" Then',
          '            rest = rest & ";" & parts(i)',
          '        End If',
          '    Next',
          '    If prov = "" Then prov = "MSDASQL"',
          '    If LCase$(Left$(prov, 15)) = "microsoft.jet.o" Then rest = ";Persist Security Info=True" & rest',
          '    lb_shape_cs = "Provider=MSDataShape.1" & rest & ";Data Provider=" & UCase$(prov)',
          'End Function', '']
    # the connections open when a command is first needed
    L += ['Public Sub lb_prepare()', '    If lb_ready Then Exit Sub', '    lb_ready = True']
    for cn, _, _ in conns:
        L.append(f'    If {cn}.State = 0 Then {cn}.Open')
    for c in top:
        cn = c.conn or (conns[0][0] if conns else '')
        if cn:
            L.append(f'    Set lb_cmd_{c.name}.ActiveConnection = {cn}')
    L += ['End Sub', '']
    for c in top:
        member = c.grouping_name if c.grouping else c.name
        cn = c.conn or (conns[0][0] if conns else '')
        # the parameters of the command are the method's, of their HostType (VB6 oracle d_de2: "x" for a
        # Currency gives 13 Type mismatch); their values go to the command for this Open only
        import re as _re
        ins = [(i, prm) for i, prm in enumerate(c.params) if prm[2] in (1, 3)] \
            if not (c.grouping or c.hierarchical()) else []
        names = [prm[0] if _re.fullmatch(r'[A-Za-z]\w{0,39}', prm[0]) else f'lb_p{i}' for i, prm in ins]
        sig = ', '.join(f'Optional ByVal {n} As Variant' for n in names)
        conv = {2: 'CInt', 3: 'CLng', 4: 'CSng', 5: 'CDbl', 6: 'CCur', 7: 'CDate', 8: 'CStr', 11: 'CBool', 17: 'CByte'}
        L += [f'Public Sub {member}({sig})',
              '    Dim lb_i As Integer, lb_old() As Variant, lb_has() As Boolean',
              ]
        if ins:
            L += [f'    ReDim lb_old({len(ins) - 1}), lb_has({len(ins) - 1})']
            hts = getattr(c, 'hosttypes', [])
            for k, (n, (i, prm)) in enumerate(zip(names, ins)):
                f = conv.get(hts[i] if i < len(hts) else 12)
                L.append(f'    If Not IsMissing({n}) Then lb_has({k}) = True' + (f': {n} = {f}({n})' if f else ''))
        L += ['    lb_prepare',
              f'    If {cn}.State = 0 Then Err.Raise 3704, "ADODB.Connection", "Operation is not allowed when the object is closed."',
              f'    If rs{member}.State <> 0 Then',
              f'        {cn}.lb_willexecute lb_cmd_{c.name}.CommandText',
              '        Err.Raise 3705, "ADODB.Connection", "Operation is not allowed when the object is open."',
              '    End If',
              ]
        for k, (n, (i, prm)) in enumerate(zip(names, ins)):
            L += [f'    If lb_has({k}) Then',
                  f'        lb_old({k}) = lb_cmd_{c.name}.Parameters({i}).Value',
                  f'        lb_cmd_{c.name}.Parameters({i}).Value = {n}',
                  '    End If']
        L += [f'    rs{member}.Open lb_cmd_{c.name}']
        for k, (n, (i, prm)) in enumerate(zip(names, ins)):
            L.append(f'    If lb_has({k}) Then lb_cmd_{c.name}.Parameters({i}).Value = lb_old({k})')
        L += ['End Sub', '']
    L += ['Public Property Get Commands() As Object', '    Set Commands = lb_commands', 'End Property',
          'Public Property Get Connections() As Object', '    Set Connections = lb_connections', 'End Property',
          'Public Property Get Recordsets() As Object', '    Set Recordsets = lb_recordsets', 'End Property', '']
    # the bound controls and the DataReport: DataSource = deX, DataMember = a command
    # DataMember "": the first command (a DataReport asks for it when it is made: VB6 oracle d_de2)
    L += ['Private Sub Class_GetDataMember(DataMember As String, Data As Object)',
          '    Set Data = lb_member(DataMember)',
          'End Sub', '',
          'Public Function lb_member(ByVal DataMember As String) As Object']
    for k, c in enumerate(top):
        member = c.grouping_name if c.grouping else c.name
        cond = f'LCase$(DataMember) = {_s(member.lower())}' + (' Or DataMember = ""' if k == 0 else '')
        L += [f'    If {cond} Then',
              f'        If rs{member}.State = 0 Then {member}',
              f'        Set lb_member = rs{member}',
              '        Exit Function',
              '    End If']
    L += ['End Function']
    attrs = [(n, v) for n, v, _ in sf.attributes]
    # the collections: Item opens the connections first (VB6 oracle d_de1)
    C = ['Option Explicit',
         'Private lb_c As New Collection',
         'Public Sub lb_add(o As Object, ByVal k As String)',
         '    If k = "" Then lb_c.Add o Else lb_c.Add o, k',
         'End Sub',
         'Public Function Item(ByVal Index As Variant) As Object',
         'Attribute Item.VB_UserMemId = 0',
         f'    {name}.lb_prepare',
         '    Set Item = lb_c.Item(Index)',
         'End Function',
         'Public Property Get Count() As Long',
         '    Count = lb_c.Count',
         'End Property',
         'Public Function NewEnum() As IUnknown',
         'Attribute NewEnum.VB_UserMemId = -4',
         f'    {name}.lb_prepare',
         '    Set NewEnum = lb_c.[_NewEnum]',
         'End Function']
    cattrs = [('VB_Name', f'"{coll}"'), ('VB_PredeclaredId', 'False'), ('VB_Exposed', 'False'), ('VB_Creatable', 'False')]
    return [(name, (decls, user, L), attrs), (coll, (C, [], []), cattrs)]


# ---------------------------------------------------------------- DataReport
# ItemType of the controls in the .dsr (the VB98 sample DataRept and VB6 oracle d_rpt1)
RPT_ITEMS = {3: 'RptLabel', 4: 'RptTextBox', 5: 'RptLine', 13: 'RptFunction', 6: 'RptShape', 7: 'RptImage'}
_US, _RS = '\x1f', '\x1e'                       # the fields of a record, a new line inside a value


def _flat(block, pre=''):
    """all the properties of a block and of its inner blocks (Font, DataFormat): name -> value"""
    out = {}
    for it in block.items:
        k = type(it).__name__
        if k == 'Property':
            n = it.name.lower()
            if n.startswith('object.'):
                n = n[7:]
            out.setdefault(pre + n, it.value)
        elif k == 'PropertyBlock' and it.name.lower() in ('font', 'dataformat'):
            out.update(_flat(it, pre + it.name.lower() + '.'))
    return out


def _rec(kind, d):
    parts = [kind]
    for k, v in d.items():
        if isinstance(v, bool):
            v = -1 if v else 0
        parts.append(f"{k}={str(v).replace(chr(13), '').replace(chr(10), _RS)}")
    return _US.join(parts)


def report_pictures(sf, path):
    """the Pictures of the RptImage controls, kept in the .dsx as in a .frx: [(control name, file bytes)]"""
    import os
    from .project.source import FrxRef
    from .project.frx import FrxStore
    from .project.fsutil import ci_path
    out = []
    for b in _blocks(sf.root, 'Section'):
        for c in _blocks(b, 'Item'):
            for it in c.items:
                if type(it).__name__ == 'Property' and it.name.lower() == 'picture' and isinstance(it.value, FrxRef):
                    f = ci_path(os.path.join(os.path.dirname(path or '.'), it.value.file))
                    try:
                        d = FrxStore.load(f, [it.value.offset]).decode('Picture', it.value)
                    except (OSError, ValueError):
                        continue
                    if isinstance(d, tuple) and len(d) == 2 and d[1]:
                        out.append((str(_props(c).get('name', '')), bytes(d[1])))
    return out


def report_definition(sf):
    """the report as the runtime reads it (runtime/forms/report.inc): one record per line,
    R (the report), S (a section), C (a control)"""
    root = sf.root
    lines = [_rec('R', {k: v for k, v in _flat(root).items() if not k.startswith('sectioncode') and k not in
                        ('numsections', 'bindings', '_version', '_designerversion', '_extentx', '_extenty')})]
    codes = {}
    for it in root.items:
        if type(it).__name__ == 'Property' and it.name.lower().startswith('sectioncode'):
            codes[it.name[11:]] = it.value
    for b in _blocks(root, 'Section'):
        sp = _flat(b)
        sd = {'code': codes.get(b.name[7:], 4)}
        sd.update({k: v for k, v in sp.items() if not k.startswith('itemtype') and k not in ('numcontrols', '_version')})
        lines.append(_rec('S', sd))
        types = {}
        for it in b.items:
            if type(it).__name__ == 'Property' and it.name.lower().startswith('itemtype'):
                types[it.name[8:]] = it.value
        for c in _blocks(b, 'Item'):
            cd = {'type': RPT_ITEMS.get(int(types.get(c.name[4:], 3)), 'RptLabel')}
            cd.update({k: v for k, v in _flat(c).items() if k not in ('_version', 'picture')})
            lines.append(_rec('C', cd))
    return '\n'.join(lines)


def _dsx_source(path):
    """the DataSource of a report: the DataEnvironment named in its .dsx"""
    import os
    for cand in (os.path.splitext(path)[0] + '.dsx', os.path.splitext(path)[0] + '.DSX'):
        try:
            data = open(cand, 'rb').read()
        except OSError:
            continue
        if len(data) > 9:
            n = data[8]
            name = data[9:9 + n].decode('cp1252', 'replace')
            # (a .dsx that starts with the Icon of the report - Icon = "X.dsx":0000 - has no DataSource there:
            # its bytes are not a name - robot 2026-11-12: six reports of PSC)
            import re
            if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', name):
                return name
    return ''


REPORT_EVENTS = {   # DataReport_<event>: its parameters, as the runtime passes them
    'initialize': 0, 'terminate': 0, 'activate': 0, 'deactivate': 0, 'resize': 0,
    'queryclose': 2, 'error': 4, 'asyncprogress': 4, 'processingtimeout': 4}


def datareport_source(sf, path):
    """the VB code of the class of a DataReport .dsr: the report object of the runtime inside, its members
    reached through it (lb_inner), its events sent to the DataReport_ procedures of the code"""
    import re
    name = sf.name
    user = list(sf.code)
    opts = []
    for i, ln in enumerate(user):
        if ln.strip().lower().startswith('option '):
            opts.append(ln)
            user[i] = ''
    low = '\n'.join(user).lower()

    def has(ev):
        return re.search(r'(?m)^\s*(?:private\s+|public\s+)?sub\s+datareport_' + ev + r'\s*\(', low) is not None

    definition = report_definition(sf)
    src = _dsx_source(path)
    decls = list(opts) + ["' ---- made by LinBASIC Studio from the DataReport designer", 'Private lb_r As Object', '']
    L = ['Private Sub Class_Initialize()', '    Dim lb_d As String, lb_p() As String']
    for rec in definition.split('\n'):          # one statement per record: short expressions
        L.append('    lb_d = lb_d & ' + ' & '.join(_vbstr_chunks(rec)[:-0 or None]) + ' & Chr$(10)')
    L.append('    Set lb_r = lb_report_new(Me, lb_d)')
    import base64
    for who, data in report_pictures(sf, path):  # the pictures: base64 in pieces, joined once
        b64 = base64.b64encode(data).decode('ascii')
        parts = [b64[i:i + 1000] for i in range(0, len(b64), 1000)]
        L.append(f'    ReDim lb_p({len(parts) - 1}) As String')
        for i, part in enumerate(parts):
            L.append(f'    lb_p({i}) = "{part}"')
        L.append(f'    lb_r.lb_picture {_s(who)}, Join(lb_p, "")')
    if src:
        L.append(f'    Set lb_r.DataSource = {src}')
        # the report asks its DataEnvironment for the default member when it is made (VB6 oracle d_de2)
        L += ['    On Error Resume Next', f'    {src}.lb_member ""', '    Err.Clear', '    On Error GoTo 0']
    if has('initialize'):
        L.append('    DataReport_Initialize')
    L += ['End Sub', '', 'Private Sub Class_Terminate()']
    if has('terminate'):
        L.append('    DataReport_Terminate')
    L += ['    lb_r.lb_close', 'End Sub', '',
          'Public Property Get lb_inner() As Object', '    Set lb_inner = lb_r', 'End Property', '',
          'Public Sub lb_unload()', '    lb_r.lb_close', 'End Sub', '']
    # the events: the runtime calls lb_event, the code's procedures are Private
    L += ['Public Sub lb_event(ByVal lb_name As String, lb_a1 As Variant, lb_a2 As Variant, lb_a3 As Variant, lb_a4 As Variant)',
          '    Dim lb_b As Boolean, lb_i As Integer, lb_j As Integer',
          '    Select Case lb_name']
    for ev, n in REPORT_EVENTS.items():
        if ev in ('initialize', 'terminate') or not has(ev):
            continue
        L.append(f'    Case "{ev}"')
        if ev == 'queryclose':
            L += ['        lb_i = lb_a1', '        lb_j = lb_a2', '        DataReport_QueryClose lb_i, lb_j', '        lb_a1 = lb_i']
        elif ev == 'error':
            L += ['        lb_b = lb_a4', '        DataReport_Error lb_a1, lb_a2, lb_a3, lb_b', '        lb_a4 = lb_b']
        elif ev == 'processingtimeout':
            L += ['        lb_b = lb_a2', '        DataReport_ProcessingTimeout lb_a1, lb_b, lb_a3, lb_a4', '        lb_a2 = lb_b']
        elif ev == 'asyncprogress':
            L.append('        DataReport_AsyncProgress lb_a1, lb_a2, lb_a3, lb_a4')
        else:
            L.append(f'        DataReport_{ev.capitalize()}')
    L += ['    End Select', 'End Sub']
    attrs = [(n, v) for n, v, _ in sf.attributes]
    return [(name, (decls, user, L), attrs)]


def _vbstr_chunks(s):
    """a long text as VB string pieces (a line of VB has a limit): "..." & Chr$(10) & ..."""
    out = []
    for line in s.split('\n'):
        parts = line.split(_US)
        piece = []
        for i, p in enumerate(parts):
            if i:
                piece.append('Chr$(31)')
            if p:
                piece.append(_s(p.replace(_RS, '\n')).replace('\n', '" & Chr$(30) & "'))
        out.append(' & '.join(piece) if piece else '""')
        out.append('Chr$(10)')
    return out[:-1] if out else ['""']


# ---------------------------------------------------------------- the Add-In designer (phase 12.5)
# the events of AddinInstance (AddInDesignerObjects) and of IDTExtensibility (VBIDE), as the IDE calls them through
# the add-in host (runtime/addin.c): (event, parameters of the wrapper, the arguments of the event procedure)
ADDIN_EVENTS = [
    ('OnConnection', 'ByVal Application As Object, ByVal ConnectMode As Long, ByVal AddInInst As Object',
     'Application, ConnectMode, AddInInst, lb_custom'),
    ('OnDisconnection', 'ByVal RemoveMode As Long', 'RemoveMode, lb_custom'),
    ('OnStartupComplete', '', 'lb_custom'),
    ('OnAddInsUpdate', '', 'lb_custom'),
    ('OnBeginShutdown', '', 'lb_custom'),
]


def addin_info(sf):
    """the properties of the designer: DisplayName, Description, LoadBehavior, CmdLineSupport ..."""
    p = _props(sf.root) if sf.root is not None else {}
    return {'displayname': str(p.get('displayname', sf.name) or sf.name), 'description': str(p.get('description', '') or ''),
            'appname': str(p.get('appname', '') or ''), 'appver': str(p.get('appver', '') or ''),
            'loadname': str(p.get('loadname', '') or ''), 'loadbehavior': int(p.get('loadbehavior', 0) or 0),
            'cmdlinesupport': int(p.get('cmdlinesupport', 0) or 0) != 0}


def addin_source(sf):
    """the class of the Connect designer: its code, and a public procedure for each event the IDE raises
    (lbaddin_onconnection ...) that calls AddinInstance_<event> and IDTExtensibility_<event> of the code.
    Implements IDTExtensibility: the IDE calls those procedures itself (no interface of a library here)"""
    import re
    name = sf.name
    user = list(sf.code)
    opts = []
    for i, ln in enumerate(user):              # Option ... first; the code keeps its line numbers
        low = ln.strip().lower()
        if low.startswith('option '):
            opts.append(ln)
            user[i] = ''
        elif re.match(r'implements\s+(vbide\.)?idtextensibility2?\s*$', low):
            user[i] = ''
    lowcode = '\n'.join(user).lower()

    def has(proc):
        return re.search(r'(?m)^\s*(?:private\s+|public\s+|friend\s+)?sub\s+' + proc.lower() + r'\s*\(', lowcode) is not None

    decls = list(opts) + ["' ---- made by LinBASIC Studio from the Add-In designer"]
    procs = ['']
    for ev, params, args in ADDIN_EVENTS:
        procs.append(f'Public Sub lbaddin_{ev}({params})')
        procs.append('    Dim lb_custom() As Variant')
        for pre in ('AddinInstance', 'IDTExtensibility', 'IDTExtensibility2'):
            if has(f'{pre}_{ev}'):
                procs.append(f'    {pre}_{ev} {args}')
        procs.append('End Sub')
    attrs = [(n, v) for n, v, _ in sf.attributes]
    return [(name, (decls, user, procs), attrs)]


def addin_host_source(connect):
    """the start of an add-in made as a program: the class of the designer, then the host of the runtime (addin.c),
    which talks with the IDE until it unloads the add-in"""
    return ['Attribute VB_Name = "LbAddInHost"', 'Option Explicit', 'Sub Main()',
            f'    Dim lb_connect As {connect}', f'    Set lb_connect = New {connect}',
            '    LbAddIn_HostRun lb_connect', 'End Sub']
