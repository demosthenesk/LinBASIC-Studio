Attribute VB_Name = "ModBstrPtr"
Option Explicit
' Phase 6.8: CopyMemory of the 4-byte BSTR pointer of String variables (ByVal VarPtr(s)), as the sorts of the
' PSC projects move Strings without copying them (robot: BrionWordSearch, TriQuickSortString of mdlArray.bas).

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub SwapStr(a As String, b As String)
    Dim t As Long
    CopyMemory t, ByVal VarPtr(a), 4
    CopyMemory ByVal VarPtr(a), ByVal VarPtr(b), 4
    CopyMemory ByVal VarPtr(b), t, 4
End Sub

Public Sub T_BstrPtr(f As Integer)
    Dim a As String, b As String, i As Long, j As Long, s() As String, t As String, z As Long, p As Long
    a = "first": b = "second"
    SwapStr a, b
    Print #f, a; "|"; b; "|"; Len(a); Len(b)

    s = Split("pear apple fig banana cherry date", " ")
    For i = LBound(s) + 1 To UBound(s)              ' an insertion sort that moves the pointers
        CopyMemory ByVal VarPtr(t), ByVal VarPtr(s(i)), 4
        j = i
        Do While j > LBound(s)
            If s(j - 1) <= t Then Exit Do
            CopyMemory ByVal VarPtr(s(j)), ByVal VarPtr(s(j - 1)), 4
            j = j - 1
        Loop
        CopyMemory ByVal VarPtr(s(j)), ByVal VarPtr(t), 4
    Next
    z = 0
    CopyMemory ByVal VarPtr(t), ByVal VarPtr(z), 4  ' t lets go of the last String without freeing it
    Print #f, Join(s, ","); "|"; Len(t)
    s(0) = s(0) & "!"                               ' the Strings are whole and each is in one place
    s(5) = UCase$(s(5))
    Print #f, s(0); "|"; s(5); "|"; UBound(s)

    CopyMemory p, ByVal VarPtr(b), 4               ' the 4 bytes of a String: its StrPtr
    Print #f, (p = StrPtr(b)); (p <> 0)
End Sub
