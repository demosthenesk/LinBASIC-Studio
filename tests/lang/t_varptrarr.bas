Attribute VB_Name = "ModVarPtrArr"
Option Explicit
' Phase 7 (robot: RVTVBIMG, GradientFill hDC, VarPtr(Vert(0)), 5, ...): VarPtr of an array element given to an API
' that reads (or writes) the elements after it; VarPtr arithmetic over the elements of an array.

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type TVert
    x As Long
    y As Long
    c As Integer
    d As Integer
End Type

Public Sub T_VarPtrArr(f As Integer)
    Dim a(4) As Long, b(4) As Long, i As Long
    For i = 0 To 4
        a(i) = (i + 1) * 11
    Next
    CopyMemory b(0), ByVal VarPtr(a(0)), 20
    Print #f, "long:";
    For i = 0 To 4
        Print #f, b(i);
    Next
    Print #f,
    Erase b
    CopyMemory b(0), ByVal VarPtr(a(2)), 12
    Print #f, "middle:";
    For i = 0 To 4
        Print #f, b(i);
    Next
    Print #f,
    Print #f, "step:"; VarPtr(a(1)) - VarPtr(a(0)); VarPtr(a(4)) - VarPtr(a(0)); (VarPtr(a(2)) = VarPtr(a(2)))

    Dim v(2) As TVert, u(2) As TVert
    For i = 0 To 2
        v(i).x = i * 10 + 1
        v(i).y = i * 10 + 2
        v(i).c = i * 10 + 3
        v(i).d = i * 10 + 4
    Next
    CopyMemory u(0), ByVal VarPtr(v(0)), 3 * LenB(v(0))
    Print #f, "type:"; LenB(v(0));
    For i = 0 To 2
        Print #f, u(i).x; u(i).y; u(i).c; u(i).d;
    Next
    Print #f,
    Print #f, "type step:"; VarPtr(v(2)) - VarPtr(v(0))

    ' into the array: the elements from a(1) on
    b(0) = 7: b(1) = 8: b(2) = 9
    CopyMemory ByVal VarPtr(a(1)), b(0), 12
    Print #f, "write:";
    For i = 0 To 4
        Print #f, a(i);
    Next
    Print #f,

    ' 2 dimensions: the elements in memory order
    Dim m(1, 2) As Integer, n(5) As Integer, j As Long
    For i = 0 To 1
        For j = 0 To 2
            m(i, j) = i * 10 + j
        Next
    Next
    CopyMemory n(0), ByVal VarPtr(m(0, 0)), 12
    Print #f, "2d:";
    For i = 0 To 5
        Print #f, n(i);
    Next
    Print #f,
End Sub
