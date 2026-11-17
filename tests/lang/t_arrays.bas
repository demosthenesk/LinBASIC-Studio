Attribute VB_Name = "ModArrays"
Option Explicit
Private Function Total(a() As Long) As Long
    Dim i As Integer
    For i = LBound(a) To UBound(a)
        Total = Total + a(i)
    Next
End Function

Public Sub T_Arrays(f As Integer)
    Dim a(5) As Long, i As Integer
    Dim m(1 To 3, 0 To 1) As Integer
    Dim d() As String
    For i = 0 To 5: a(i) = i * i: Next
    Print #f, LBound(a); UBound(a); a(5); Total(a)
    m(2, 1) = 7
    Print #f, LBound(m, 1); UBound(m, 1); LBound(m, 2); UBound(m, 2); m(2, 1); m(1, 0)
    ReDim d(2)
    d(0) = "x": d(1) = "y": d(2) = "z"
    ReDim Preserve d(4)
    d(4) = "w"
    Print #f, UBound(d); d(0); d(2); "["; d(3); "]"; d(4)
    ReDim d(1 To 2)
    Print #f, LBound(d); "["; d(1); "]"
    Erase a
    Print #f, a(5); UBound(a)
    Dim v() As Variant
    ReDim v(1)
    v(0) = "s": v(1) = 2.5
    Print #f, v(0); v(1); TypeName(v(1))
    On Error Resume Next
    i = a(6)
    Print #f, "bad index"; Err.Number
    Err.Clear
    Erase d
    i = UBound(d)
    Print #f, "erased"; Err.Number
End Sub
