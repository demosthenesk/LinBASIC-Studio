Attribute VB_Name = "ModDeclare"
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As String) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function NoSuchFn Lib "kernel32" () As Long
Private Declare Function NoDllFn Lib "nosuchlib123.dll" () As Long

Public Sub T_Declare(f As Integer)
    Dim t1 As Long, t2 As Long
    t1 = GetTickCount
    Sleep 50
    t2 = GetTickCount
    Print #f, (t2 - t1) >= 40
    Print #f, lstrlenA("hello"); lstrlenW(StrPtr("wide!"))
    Dim buf As String, n As Long
    Call SetEnvironmentVariable("LBVC_TEST", "abc123")
    buf = String$(64, 0)
    n = GetEnvironmentVariable("LBVC_TEST", buf, 64)
    Print #f, n; "["; Left$(buf, n); "]"
    n = GetEnvironmentVariable("LBVC_NOPE_X", buf, 64)
    Print #f, n
    Dim a As Long, b As Long
    a = &H12345678
    CopyMemory b, a, 4
    Print #f, Hex$(b)
    Dim arr(3) As Byte
    CopyMemory arr(0), a, 4
    Print #f, Hex$(arr(0)); Hex$(arr(3))
    Dim i As Integer
    CopyMemory i, ByVal StrPtr("AB"), 2
    Print #f, i
    Print #f, MulDiv(100, 3, 7); MulDiv(7, 5, 2); MulDiv(-7, 5, 2); MulDiv(1, 2, 0)
    Print #f, GetCurrentProcessId() <> 0
    Dim e As String
    Print #f, VarPtr(a) <> 0; StrPtr(e) = 0; StrPtr("x") <> 0
    Print #f, lstrcpy(buf, "copied") <> 0; "["; Left$(buf, 7); "]"
    On Error Resume Next
    n = NoSuchFn
    Print #f, "err"; Err.Number; Err.Description
    Err.Clear
    n = NoDllFn
    Print #f, "err"; Err.Number; Err.Description
    Err.Clear
End Sub
