Attribute VB_Name = "ModDeclare2"
Option Explicit
' Phase 6.4: Declare ... As String (a BSTR of ANSI bytes, converted to Unicode), Strings, Variants and
' Types passed to As Any, ByRef As String, and Err.LastDllError.

Private Type TMix
    a As Integer
    s As String * 4
    l As Long
End Type

Private Type TDbl
    b As Byte
    d As Double
    s As String
End Type

Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal psz As String, ByVal cblen As Long) As String
Private Declare Function SysAllocStringByteLenP Lib "oleaut32" Alias "SysAllocStringByteLen" (ByVal psz As Long, ByVal cblen As Long) As String
Private Declare Function SysAllocString Lib "oleaut32" (ByVal psz As Long) As String
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal psz As Long, ByVal cch As Long) As String
Private Declare Function SysStringByteLen Lib "oleaut32" (ByVal bstr As String) As Long
Private Declare Function SysStringLen Lib "oleaut32" (ByVal bstr As String) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (pbstr As String, ByVal psz As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal s As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub SetLastError Lib "kernel32" (ByVal e As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal f As String) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal p As String, ByVal sa As Long) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal p As String) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal p As String) As Long
Private Declare Function NoSuchFn2 Lib "kernel32" () As Long
Private Declare Function NoDllFn2 Lib "nosuchlib456.dll" () As Long

Private Function Codes(ByVal s As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(s)
        r = r & " " & AscW(Mid$(s, i, 1))
    Next
    Codes = Len(s) & ":" & r
End Function

Private Function Bytes(b() As Byte, ByVal n As Long) As String
    Dim i As Long, r As String
    For i = 0 To n - 1
        r = r & " " & b(i)
    Next
    Bytes = Mid$(r, 2)
End Function

Private Sub SetInSub()
    On Error Resume Next
    SetLastError 58
End Sub

Private Sub Trace(ByVal what As String)
    Dim g As Integer
    g = FreeFile
    Open App.Path & "\trace64.txt" For Append As #g
    Print #g, what
    Close #g
End Sub

Private Sub PartLastErr(f As Integer)
    Dim r As String, s As String, t As String, n As Long, p As Long, i As Integer
    Dim b(31) As Byte, b2(31) As Byte, v As Variant
    ' ---- Err.LastDllError: GetLastError, read after every Declare call
    Print #f, "lasterror"
    Dim buf As String, x As Long
    buf = String$(64, 0)
    SetLastError 1234
    Print #f, Err.LastDllError
    x = 1 + 2
    Print #f, Err.LastDllError
    SetLastError 0
    Print #f, Err.LastDllError
    SetLastError &H80070005
    Print #f, Err.LastDllError
    n = GetEnvironmentVariable("LBVC_NOPE_Y", buf, 64)
    Print #f, n; Err.LastDllError
    SetLastError 99
    Call SetEnvironmentVariable("LBVC_T2", "v")
    n = GetEnvironmentVariable("LBVC_T2", buf, 64)
    Print #f, n; Err.LastDllError
    n = GetFileAttributes(App.Path & "\no_such_file_x.txt")
    Print #f, n; Err.LastDllError
    n = GetFileAttributes(App.Path & "\no_such_dir_x\f.txt")
    Print #f, n; Err.LastDllError
    SetLastError 99
    n = GetFileAttributes(App.Path)
    Print #f, (n And 16); Err.LastDllError
    Dim dn As String
    dn = App.Path & "\lbtmpdir64"
    Call RemoveDirectory(dn)
    SetLastError 99
    n = CreateDirectory(dn, 0)
    Print #f, n <> 0; Err.LastDllError
    n = CreateDirectory(dn, 0)
    Print #f, n; Err.LastDllError
    SetLastError 99
    n = RemoveDirectory(dn)
    Print #f, n <> 0; Err.LastDllError
    n = RemoveDirectory(dn)
    Print #f, n; Err.LastDllError
    n = DeleteFile(App.Path & "\no_such_file_x.txt")
    Print #f, n; Err.LastDllError
    SetLastError 59
    x = lstrlenA("abc")
    Print #f, x; Err.LastDllError
    SetLastError 60
    Err.Clear
    Print #f, "clear"; Err.LastDllError
    SetLastError 61
    On Error Resume Next
    Print #f, "on error"; Err.LastDllError
    SetLastError 62
    Err.Raise 5
    Print #f, "raise"; Err.Number; Err.LastDllError
    Err.Clear
    SetLastError 63
    SetInSub
    Print #f, "sub"; Err.LastDllError
    n = NoSuchFn2
    Print #f, "entry"; Err.Number; Err.LastDllError
    Err.Clear
    n = NoDllFn2
    Print #f, "dll"; Err.Number; Err.LastDllError
    Err.Clear
    SetLastError 64
    n = GetTickCount
    Print #f, "tick"; Err.LastDllError
    ' which calls set it to 0 before the call (a String argument?)
    Dim l As Long, m As TMix, b1(3) As Byte, w As String
    w = "abc"
    SetLastError 70
    n = MulDiv(1, 2, 3)
    Print #f, "long args"; Err.LastDllError
    SetLastError 71
    n = lstrlenA(vbNullString)
    Print #f, "null string"; Err.LastDllError
    SetLastError 72
    CopyMemory l, ByVal 0&, 0
    Print #f, "as any"; Err.LastDllError
    SetLastError 73
    CopyMemory b1(0), ByVal w, 1
    Print #f, "as any byval string"; Err.LastDllError
    SetLastError 74
    CopyMemory b1(0), m, 2
    Print #f, "as any type"; Err.LastDllError
    SetLastError 75
    w = SysAllocStringByteLenP(0, 0)
    Print #f, "as string"; Err.LastDllError
    SetLastError 76
    n = SysStringLen(w)
    Print #f, "byval string"; Err.LastDllError
End Sub

Private Sub PartType(f As Integer)
    Dim r As String, s As String, t As String, n As Long, p As Long, i As Integer
    Dim b(31) As Byte, b2(31) As Byte, v As Variant
    ' ---- Types to As Any: ANSI strings, natural alignment
    Print #f, "any-type"
    Dim m As TMix, m2 As TMix, td As TDbl, ma(1) As TMix
    m.a = &H102: m.s = "wx" & ChrW(&H3B1) & "z": m.l = &H3040506
    For i = 0 To 31: b(i) = 255: Next
    CopyMemory b(0), m, 12
    Print #f, b(0); b(1); b(2); b(3); b(4); b(5); "|"; b(8); b(9); b(10); b(11)
    Print #f, Len(m); LenB(m)
    b(0) = 7: b(1) = 0: b(2) = 97: b(3) = 98: b(4) = 225: b(5) = 99
    b(8) = 1: b(9) = 0: b(10) = 0: b(11) = 0
    CopyMemory m2, b(0), 12
    Print #f, m2.a; Codes(m2.s); m2.l
    ma(1) = m
    CopyMemory b(0), ma(1), 12
    Print #f, b(0); b(2); b(8)
    td.b = 9: td.d = 1: td.s = "hey"
    For i = 0 To 31: b(i) = 0: Next
    CopyMemory b(0), td, 24
    Print #f, Len(td); LenB(td); b(0)
    Dim dOff As Long, sOff As Long
    If b(15) = 63 Then
        dOff = 8: sOff = 16
    ElseIf b(11) = 63 Then
        dOff = 4: sOff = 12
    ElseIf b(8) = 63 Then
        dOff = 1: sOff = 9
    End If
    Print #f, "double at"; dOff
    If sOff > 0 Then
        CopyMemory p, b(sOff), 4
        Print #f, "string member"; p <> 0      ' (its ANSI copy is freed after the call)
    End If
End Sub

Private Sub PartNum(f As Integer)
    Dim r As String, s As String, t As String, n As Long, p As Long, i As Integer
    Dim b(31) As Byte, b2(31) As Byte, v As Variant
    ' ---- numbers and Variants to As Any
    Print #f, "any-num"
    Dim d As Double, sg As Single, c As Currency, bo As Boolean, dt As Date, l As Long
    d = 1
    CopyMemory b(0), d, 8
    Print #f, Bytes(b, 8)
    sg = -2
    CopyMemory b(0), sg, 4
    Print #f, Bytes(b, 4)
    c = 1
    CopyMemory b(0), c, 8
    Print #f, Bytes(b, 8)
    bo = True
    CopyMemory b(0), bo, 2
    Print #f, Bytes(b, 2)
    dt = #1/2/1900#
    CopyMemory b(0), dt, 8
    Print #f, Bytes(b, 8)
    l = 0
    CopyMemory l, ByVal 0&, 0
    Print #f, l
    v = CLng(&H11223344)
    CopyMemory b(0), v, 16
    Print #f, b(0); b(1); b(8); b(9); b(10); b(11)
    v = 2.5
    CopyMemory b(0), v, 16
    Print #f, b(0); b(1); b(14); b(15)
    v = "AB"
    CopyMemory b(0), v, 16
    Print #f, b(0); b(1)
    CopyMemory p, b(8), 4
    CopyMemory b2(0), ByVal p, 4
    Print #f, Bytes(b2, 4)
    i = 0
    CopyMemory i, ByVal VarPtr(v), 2
    Print #f, i
    v = Empty
    CopyMemory b(0), v, 2
    Print #f, b(0); b(1)
End Sub

Private Sub PartRet(f As Integer)
    Dim r As String, s As String, t As String, n As Long, p As Long, i As Integer
    Dim b(31) As Byte, b2(31) As Byte, v As Variant
    ' ---- As String: the BSTR the procedure returns holds ANSI bytes; VB converts them, then frees it
    Print #f, "ret"
    Print #f, Codes(SysAllocStringByteLen("hello", 5))
    Print #f, Codes(SysAllocStringByteLen("abcdef", 3))
    Print #f, Codes(SysAllocStringByteLen("a" & vbNullChar & "b", 3))
    Print #f, Codes(SysAllocStringByteLen(ChrW(&H3B1) & ChrW(&H3B2) & "z", 3))
    Print #f, Codes(SysAllocString(StrPtr("hi")))
    Print #f, Codes(SysAllocStringLen(StrPtr("abc"), 3))
    Print #f, Codes(SysAllocStringLen(StrPtr("abc"), 0))
    r = SysAllocStringByteLenP(0, 0)
    Print #f, Len(r); StrPtr(r) <> 0
    r = SysAllocString(0)
    Print #f, Len(r); StrPtr(r) = 0
    Print #f, StrConv(SysAllocString(StrPtr("hi")), vbFromUnicode)
    v = SysAllocStringByteLen("var", 3)
    Print #f, VarType(v); v
End Sub

Private Sub PartByVal(f As Integer)
    Dim r As String, s As String, t As String, n As Long, p As Long, i As Integer
    Dim b(31) As Byte, b2(31) As Byte, v As Variant
    ' ---- ByVal As String: an ANSI BSTR (with its length), NULL for vbNullString
    Print #f, "byval"
    Print #f, SysStringByteLen("abc"); SysStringLen("abc"); SysStringLen("abcd")
    Print #f, SysStringByteLen(vbNullString); SysStringByteLen("")
    Print #f, SysStringByteLen("a" & vbNullChar & "b"); lstrlenA("a" & vbNullChar & "b")
    Print #f, SysStringByteLen(ChrW(&H3B1) & ChrW(&H4E2D)); lstrlenA(vbNullString)
End Sub

Private Sub PartAnyStr(f As Integer)
    Dim r As String, s As String, t As String, n As Long, p As Long, i As Integer
    Dim b(31) As Byte, b2(31) As Byte, v As Variant
    ' ---- Strings to As Any
    Print #f, "any-str"
    Trace "s1"
    s = "abcd"
    CopyMemory b(0), ByVal s, 4
    Trace "s2"
    Print #f, Bytes(b, 4)
    Trace "s3"
    s = "x" & ChrW(&H3B1) & ChrW(&H4E2D)
    CopyMemory b(0), ByVal s, 3
    Trace "s4"
    Print #f, Bytes(b, 3)
    Trace "s5"
    CopyMemory b(0), ByVal StrPtr("AB"), 4
    Trace "s6"
    Print #f, Bytes(b, 4)
    Trace "s7"
    CopyMemory b(0), ByVal "xyz", 3
    Trace "s8"
    Print #f, Bytes(b, 3)
    Trace "s9"
    s = "abcd"
    p = 0
    CopyMemory p, s, 4
    Trace "s10"
    If p = &H64636261 Then
        Print #f, "byref string: the text"
        Trace "s11"
    Else
        Print #f, "byref string: a pointer"
        Trace "s12"
    End If
    Dim fs As String * 4
    fs = "wxyz"
    p = 0
    CopyMemory p, fs, 4
    Trace "s13"
    If p = &H7A797877 Then
        Print #f, "byref fixed: the text"
        Trace "s14"
    Else
        Print #f, "byref fixed: a pointer"
        Trace "s15"
    End If
    b(0) = 225: b(1) = 226: b(2) = 65
    s = String$(3, "-")
    CopyMemory ByVal s, b(0), 3
    Trace "s18"
    Print #f, Codes(s)
    Trace "s19"
    t = String$(2, "-")
    CopyMemory ByVal t, ByVal ChrW$(&H4E2D) & "a", 2
    Trace "s20"
    Print #f, Codes(t)
    Trace "s21"
    s = "wxyz"
    CopyMemory ByVal s, ByVal "AB", 2
    Trace "s22"
    Print #f, s
    Trace "s23"
    s = "abcd"
    CopyMemory ByVal s, ByVal "x" & vbNullChar, 2
    Trace "s24"
    Print #f, Codes(s)
    Trace "s25"
End Sub

Public Sub T_Declare2(f As Integer)
    ' (the parts in this order: the strings to As Any made the VB6 program hang; trace64.txt says where)
    PartLastErr f
    PartType f
    PartNum f
    PartRet f
    PartByVal f
    PartAnyStr f
    PartVarByVal f
End Sub

Private Sub PartVarByVal(f As Integer)
    ' a Variant ByVal to As Any: VB6 pushes the 16 bytes of the VARIANT (the Strings part found it)
    Dim t As String, v As Variant
    Print #f, "byval-variant"
    t = "--"
    v = "ab"
    On Error Resume Next
    Trace "v1"
    CopyMemory ByVal t, ByVal v, 2
    Trace "v2"
    Print #f, Err.Number; t
    Err.Clear
    v = 5
    CopyMemory ByVal t, ByVal v, 2
    Print #f, Err.Number; t
End Sub
