# gen6.py - writes Checks.bas: what the Advanced Optimizations of VB6 change in the RESULTS of a program
# (Remove Integer Overflow Checks, Remove Array Bounds Checks, Remove Floating Point Error Checks, Allow
# Unrounded Floating Point Operations). Every case runs twice: H = inline in a procedure with On Error Resume
# Next; N = in a function of its own without On Error (an error goes up to the caller).
INT = [
    ('int add', 'i = KI(32767): i = i + 1', 'i'),
    ('int sub', 'i = KI(-32768): i = i - 1', 'i'),
    ('int mul', 'i = KI(200): i = i * 200', 'i'),
    ('int neg', 'i = KI(-32768): i = -i', 'i'),
    ('int abs', 'i = KI(-32768): i = Abs(i)', 'i'),
    ('int idiv -1', 'i = KI(-32768): i = i \\ -1', 'i'),
    ('int add var', 'i = KI(30000): i = i + KI(30000)', 'i'),
    ('int*int to long', 'i = KI(300): l = 0: l = i * i', 'l'),
    ('int+1 to long', 'i = KI(32767): l = 0: l = i + 1', 'l'),
    ('int+1 to variant', 'i = KI(32767): v = Empty: v = i + 1', 'v'),
    ('hex add', 'i = &H7FFF: i = i + KI(1)', 'i'),
    ('long add', 'l = KL(2147483647): l = l + 1', 'l'),
    ('long sub', 'l = KL(-2147483647) - 1: l = l - 1', 'l'),
    ('long mul', 'l = KL(65536): l = l * 65536', 'l'),
    ('long neg', 'l = KL(-2147483647) - 1: l = -l', 'l'),
    ('long abs', 'l = KL(-2147483647) - 1: l = Abs(l)', 'l'),
    ('long idiv -1', 'l = KL(-2147483647) - 1: l = l \\ -1', 'l'),
    ('long+int', 'l = KL(2147483647): l = l + KI(1)', 'l'),
    ('byte add', 'b = KB(255): b = b + 1', 'b'),
    ('byte sub', 'b = KB(0): b = b - 1', 'b'),
    ('byte mul', 'b = KB(16): b = b * 16', 'b'),
    ('byte+byte to int', 'b = KB(255): i = 0: i = b + b', 'i'),
    ('long to int', 'i = 0: l = KL(40000): i = l', 'i'),
    ('CInt(long)', 'i = 0: l = KL(40000): i = CInt(l)', 'i'),
    ('CInt(double)', 'i = 0: d = KD(40000): i = CInt(d)', 'i'),
    ('CInt(32767.5)', 'i = 0: d = KD(32767.5): i = CInt(d)', 'i'),
    ('double to int', 'i = 0: d = KD(40000): i = d', 'i'),
    ('int to byte', 'b = 0: i = KI(300): b = i', 'b'),
    ('CByte(int)', 'b = 0: i = KI(300): b = CByte(i)', 'b'),
    ('-1 to byte', 'b = 0: i = KI(-1): b = i', 'b'),
    ('double to long', 'l = 0: d = KD(3000000000#): l = d', 'l'),
    ('CLng(double)', 'l = 0: d = KD(3000000000#): l = CLng(d)', 'l'),
    ('int / 0.5 to int', 'i = KI(32767): i = i / 0.5', 'i'),
    ('int ^ 16 to int', 'i = KI(2): i = i ^ 16', 'i'),
    ('currency mul', 'c = KC(922337203685477@): c = c * 10', 'c'),
    ('currency add', 'c = KC(922337203685477@): c = c + c', 'c'),
    ('currency to int', 'i = 0: c = KC(40000): i = c', 'i'),
    ('double to currency', 'c = 0: d = KD(1E+300): c = d', 'c'),
    ('For Integer to 32767', 'n = 0: For i = KI(32765) To 32767: n = n + 1: If n > 10 Then Exit For\n    Next', 'n & "," & i'),
    ('For Byte to 255', 'n = 0: For b = KB(253) To 255: n = n + 1: If n > 10 Then Exit For\n    Next', 'n & "," & b'),
]
BOUNDS = [
    ('fixed g(10)', 'Call FillAll: x = 0: x = g(KL(10))', 'x'),
    ('fixed g(11)', 'Call FillAll: x = 0: x = g(KL(11))', '"-"'),
    ('fixed g(0)', 'Call FillAll: x = 0: x = g(KL(0))', '"-"'),
    ('dynamic dd(11)', 'ReDim dd(1 To 10): x = 0: x = dd(KL(11))', '"-"'),
    ('dynamic dd(0)', 'ReDim dd(1 To 10): x = 0: x = dd(KL(0))', '"-"'),
    ('m(4,1) of 3x3', 'Call FillAll: x = 0: x = m(KL(4), 1)', 'x'),
    ('m(0,2) of 3x3', 'Call FillAll: x = 0: x = m(KL(0), 2)', 'x'),
    ('m(1,4) of 3x3', 'Call FillAll: x = 0: x = m(1, KL(4))', '"-"'),
    ('m(4,1)=99 then m(1,2)', 'Call FillAll: m(KL(4), 1) = 99', 'm(1, 2)'),
    ('ReDim mm(3,3): mm(4,1)', 'ReDim mm(1 To 3, 1 To 3): mm(1, 2) = 12: x = 0: x = mm(KL(4), 1)', 'x'),
    ('byte bb(4) of 0 To 3', 'x = 0: x = bb(KL(4))', '"-"'),
    ('Variant Array(1,2,3)(5)', 'v = Array(1, 2, 3): x = 0: x = v(KL(5))', '"-"'),
    ('Variant holding Long array (11)', 'ReDim dd(1 To 10): v = dd: x = 0: x = v(KL(11))', '"-"'),
]
FLOAT = [
    ('1 / 0', 'd = 0: z = KD(0): d = 1 / z', 'd'),
    ('0 / 0', 'd = 0: z = KD(0): d = z / z', 'd'),
    ('-1 / 0', 'd = 0: z = KD(0): d = -1 / z', 'd'),
    ('1e300 * 1e300', 'd = KD(1E+300): d = d * d', 'd'),
    ('1e308 + 1e308', 'd = 0: d = KD(1E+308) + KD(1E+308)', 'd'),
    ('double 1e300 to single', 's = 0: d = KD(1E+300): s = d', 's'),
    ('CSng(1e300)', 's = 0: d = KD(1E+300): s = CSng(d)', 's'),
    ('single 3e38 * 10', 's = KS(3E+38): s = s * 10', 's'),
    ('single 1 / 0', 's = 0: s = KS(1) / KS(0)', 's'),
    ('10 ^ 400', 'd = 0: d = KD(10) ^ 400', 'd'),
    ('Exp(1000)', 'd = 0: d = Exp(KD(1000))', 'd'),
    ('int 1 / int 0', 'd = 0: d = KI(1) / KI(0)', 'd'),
    ('int 1 \\ int 0', 'i = 0: i = KI(1) \\ KI(0)', 'i'),
    ('int 1 Mod int 0', 'i = 0: i = KI(1) Mod KI(0)', 'i'),
    ('Variant 1 / 0', 'v = Empty: v = CVar(KD(1)) / CVar(KD(0))', 'v'),
    ('1e300*1e300 to Integer', 'i = 0: i = KD(1E+300) * KD(1E+300)', 'i'),
    ('0/0 = 0/0', 'bo = False: z = KD(0): bo = (z / z) = (z / z)', 'bo'),
    ('Sqr(-1)', 'd = 0: d = Sqr(KD(-1))', 'd'),
    ('Log(0)', 'd = 0: d = Log(KD(0))', 'd'),
]
UNROUND = [
    ('(1e16 + 1) - 1e16', 'd1 = KD(1E+16): d2 = KD(1): d = (d1 + d2) - d1', 'd'),
    ('(1e16 + 1) > 1e16', 'd1 = KD(1E+16): d2 = KD(1): bo = (d1 + d2) > d1', 'bo'),
    ('single (2^24 + 1) - 2^24', 's1 = KS(16777216): s2 = KS(1): s = (s1 + s2) - s1', 's'),
    ('single (2^24 + 1) > 2^24', 's1 = KS(16777216): s2 = KS(1): bo = (s1 + s2) > s1', 'bo'),
    ('single 0.1 * 3 = 0.3!', 'bo = (KS(0.1) * KS(3) = 0.3!)', 'bo'),
    ('0.1 + 0.2 = 0.3', 'bo = (KD(0.1) + KD(0.2) = 0.3)', 'bo'),
    ('single s = 1/3 then s = 1/3', 's = KS(1) / KS(3): bo = (s = KS(1) / KS(3))', 'bo'),
    ('double d = 1/3 then d = 1/3', 'd = KD(1) / KD(3): bo = (d = KD(1) / KD(3))', 'bo'),
    ('CSng(1/3) = single 1/3', 'bo = (CSng(KD(1) / KD(3)) = KS(1) / KS(3))', 'bo'),
    ('single 1/3*3', 's = KS(1) / KS(3) * KS(3)', 's'),
    ('single sum 0.1 x10', 's = 0: For n = 1 To 10: s = s + KS(0.1): Next', 's'),
    ('single (1/3)*3 - 1 in double', 'd = CDbl(KS(1) / KS(3) * KS(3)) - 1', 'd'),
]
DIMS = ('    Dim i As Integer, l As Long, b As Byte, c As Currency, d As Double, s As Single, z As Double\n'
        '    Dim v As Variant, n As Long, x As Long, bo As Boolean, d1 As Double, d2 As Double, s1 As Single, s2 As Single\n'
        '    Dim dd() As Long, mm() As Long, bb(0 To 3) As Byte')
out = ['Attribute VB_Name = "Checks"', 'Option Explicit', "' Generated by gen6.py",
       'Private mF As Integer', 'Private g(1 To 10) As Long', 'Private m(1 To 3, 1 To 3) As Long',
       'Private Function KI(ByVal v As Integer) As Integer', '    KI = v', 'End Function',
       'Private Function KL(ByVal v As Long) As Long', '    KL = v', 'End Function',
       'Private Function KB(ByVal v As Byte) As Byte', '    KB = v', 'End Function',
       'Private Function KD(ByVal v As Double) As Double', '    KD = v', 'End Function',
       'Private Function KS(ByVal v As Single) As Single', '    KS = v', 'End Function',
       'Private Function KC(ByVal v As Currency) As Currency', '    KC = v', 'End Function',
       'Private Sub FillAll()', '    Dim i As Long, j As Long',
       '    For i = 1 To 10: g(i) = i * 10: Next',
       '    For i = 1 To 3: For j = 1 To 3: m(i, j) = i * 10 + j: Next: Next', 'End Sub',
       'Private Sub Rep(ByVal nm As String, ByVal v As Variant, ByVal e As Long)',
       '    Dim t As String',
       '    On Error Resume Next',
       '    t = "?"',
       '    t = CStr(v)',
       '    Print #mF, nm & " = " & t & " err" & e',
       'End Sub']
n = 0
groups = (('Int', INT), ('Bounds', BOUNDS), ('Float', FLOAT), ('Unround', UNROUND))
for gname, cases in groups:
    # N: one function per case, without On Error
    for k, (name, setup, res) in enumerate(cases):
        out += [f'Private Function N_{gname}{k}() As Variant', DIMS, '    ' + setup, f'    N_{gname}{k} = {res}',
                'End Function']
    out += [f'Private Sub T_{gname}_N()', '    Dim rv As Variant', '    On Error Resume Next']
    for k, (name, setup, res) in enumerate(cases):
        out += ['    Err.Clear: rv = "(none)"', f'    rv = N_{gname}{k}()', f'    Rep "N {name}", rv, Err.Number']
    out += ['End Sub']
    # H: inline in a procedure with On Error Resume Next
    out += [f'Private Sub T_{gname}_H()', DIMS, '    On Error Resume Next']
    for k, (name, setup, res) in enumerate(cases):
        out += ['    Err.Clear', '    ' + setup, f'    Rep "H {name}", {res}, Err.Number']
    out += ['End Sub']
out += ['Private Sub T_Undim()', '    Dim ee() As Long, x As Long', '    On Error Resume Next',
        '    Err.Clear: x = -1', '    x = ee(KL(1))', '    Rep "H undimensioned ee(1)", x, Err.Number', 'End Sub']
out += ['Sub Main()', '    FillAll', '    mF = FreeFile',
        '    Open App.Path & "\\" & App.EXEName & ".txt" For Output As #mF']
for gname, _ in groups:
    out += [f'    T_{gname}_N', f'    T_{gname}_H']
out += ['    Close #mF',
        "    ' (the last one may crash the program: the file is closed before it)",
        '    mF = FreeFile', '    Open App.Path & "\\" & App.EXEName & ".txt" For Append As #mF',
        '    T_Undim', '    Close #mF',
        '    MsgBox "Done: " & App.EXEName & ".txt"', 'End Sub']
open('Checks.bas', 'wb').write(('\r\n'.join(out) + '\r\n').encode('cp1252'))
