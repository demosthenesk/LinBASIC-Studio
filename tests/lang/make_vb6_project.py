#!/usr/bin/env python3
# LinBASIC Studio - writes the VB6 project that produces tests/lang/expected/ (the "oracle"):
# LangTests.vbp + RunAll.bas + the test modules, into a folder to copy to the Windows VM. P-Code, as lbvc
# compiles the .bas tests (the order of evaluation and the checks of the VB6 IDE).
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
#   python3 tests/lang/make_vb6_project.py <out dir>
import json, os, shutil, sys
HERE = os.path.dirname(os.path.abspath(__file__))


def tests():
    """name -> (Sub, [extra files]) from tests.json"""
    out = {}
    for n, v in json.load(open(os.path.join(HERE, 'tests.json'))).items():
        out[n] = (v, []) if isinstance(v, str) else (v['sub'], v.get('files', []))
    return out


def module_name(path):
    for line in open(path, encoding='cp1252'):
        if line.startswith('Attribute VB_Name'):
            return line.split('"')[1]
    raise SystemExit(f"{path}: no Attribute VB_Name")


def main(out):
    os.makedirs(out, exist_ok=True)
    ts = tests()
    cases = '\n'.join(f'    Case {i}: {sub} f' for i, (n, (sub, _)) in enumerate(ts.items()))
    runs = '\n'.join(f'    RunOne "{n}", {i}' for i, n in enumerate(ts))
    runall = f'''Attribute VB_Name = "RunAll"
Option Explicit

Private Sub RunOne(ByVal testName As String, ByVal which As Integer)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\\expected\\" & testName & ".txt" For Output As #f
    On Error GoTo bad
    Select Case which
{cases}
    End Select
    Close #f
    Exit Sub
bad:
    Print #f, "UNHANDLED"; Err.Number; Err.Description
    Close #f
End Sub

Sub Main()
    On Error Resume Next
    MkDir App.Path & "\\expected"
    On Error GoTo 0
{runs}
    MsgBox "Done: " & {len(ts)} & " tests written to the expected folder."
End Sub
'''
    open(os.path.join(out, 'RunAll.bas'), 'wb').write(runall.replace('\n', '\r\n').encode('cp1252'))
    members = ['Module=RunAll; RunAll.bas']
    seen = set()
    for n, (sub, extra) in ts.items():
        for fn in [n + '.bas'] + extra:
            if fn in seen:
                continue
            seen.add(fn)
            shutil.copy(os.path.join(HERE, fn), os.path.join(out, fn))
            kind = 'Class' if fn.lower().endswith('.cls') else 'Module'
            members.append(f'{kind}={module_name(os.path.join(HERE, fn))}; {fn}')
    vbp = ['Type=Exe',
           r'Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\system32\stdole2.tlb#OLE Automation',
           r'Reference=*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\system32\scrrun.dll#Microsoft Scripting Runtime'] + \
        members + ['Startup="Sub Main"', 'ExeName32="LangTests.exe"', 'Command32=""', 'Name="LangTests"',
                   'HelpContextID="0"', 'CompatibleMode="0"', 'MajorVer=1', 'MinorVer=0', 'RevisionVer=0',
                   'AutoIncrementVer=0', 'ServerSupportFiles=0', 'CompilationType=-1', 'OptimizationType=0',
                   'FavorPentiumPro(tm)=0', 'CodeViewDebugInfo=0', 'NoAliasing=0', 'BoundsCheck=0',
                   'OverflowCheck=0', 'FlPointCheck=0', 'FDIVCheck=0', 'UnroundedFP=0', 'StartMode=0',
                   'Unattended=0', 'Retained=0', 'ThreadPerObject=0', 'MaxNumberOfThreads=1']
    open(os.path.join(out, 'LangTests.vbp'), 'wb').write(('\r\n'.join(vbp) + '\r\n').encode('cp1252'))
    print(f"{len(ts)} tests -> {out}")


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'work', 'vb6'))
