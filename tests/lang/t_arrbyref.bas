Attribute VB_Name = "ModArrByRef"
Option Explicit
' Phase 7 (robot: GamaSort, an insertion sort that never ended): a whole array a() as an argument goes ByRef,
' also when the call has other arguments that call something (the order of evaluation puts those arguments into
' temporaries first: a() went with them, so the procedure changed a copy; DTMF Tones: Wav_BuildHeader udt(i).bytes,
' rate, ... - a field of a UDT, the same).

Private Type TBuf
    b() As Byte
    n As Long
End Type

Private mCalls As Long

Private Function Nxt() As Long
    mCalls = mCalls + 1
    Nxt = mCalls
End Function

Private Sub SetAll(a() As Long, ByVal v As Long, ByVal w As Long)
    Dim i As Long
    For i = LBound(a) To UBound(a)
        a(i) = v * 10 + w
    Next
End Sub

Private Function SetFirst(a() As Long, ByVal v As Long) As Long
    a(0) = v
    SetFirst = a(0) + a(1)
End Function

Private Sub Fill(t() As Long, ByVal lo As Long, ByVal up As Long)
    Dim m
    If up - lo < 1 Then
        t(lo) = lo * lo + Nxt() * 0
        Exit Sub
    End If
    m = (lo + up) \ 2
    Call Fill(t(), lo, m)
    Call Fill(t(), m + 1, up)
End Sub

Private Sub Grow(a() As Long, ByVal n As Long)
    ReDim Preserve a(n)
    a(n) = n
End Sub

Private Sub Swap(ByRef a As Long, ByRef b As Long)
    Dim t As Long
    t = a: a = b: b = t
End Sub

Private Sub FillBytes(x() As Byte, ByVal v As Long, ByVal w As Long)
    ReDim x(0 To 2)
    x(0) = v
    x(2) = w
End Sub

Private Function AddByte(x() As Byte, ByVal v As Long) As Long
    ReDim Preserve x(UBound(x) + 1)
    x(UBound(x)) = v
    AddByte = UBound(x)
End Function

Private Function Show(a() As Long) As String
    Dim i As Long, s As String
    For i = LBound(a) To UBound(a)
        s = s & " " & a(i)
    Next
    Show = s
End Function

Public Sub T_ArrByRef(f As Integer)
    Dim a(3) As Long, d() As Long, v As Variant, r As Long
    Call SetAll(a(), Nxt(), Nxt())
    Print #f, "call:"; Show(a())
    SetAll a(), Nxt(), Nxt()
    Print #f, "statement:"; Show(a())
    r = SetFirst(a(), Nxt()) + Nxt()
    Print #f, "function:"; r; Show(a())
    Fill a(), 0, 3
    Print #f, "recursion:"; Show(a())
    v = 1
    Swap a(v), a(v + 2)
    Print #f, "swap:"; Show(a())
    ReDim d(1)
    Grow d(), Nxt()
    Print #f, "redim:"; UBound(d); d(UBound(d))
    Dim u(1) As TBuf
    FillBytes u(1).b, Nxt(), 7
    Print #f, "udt field:"; UBound(u(1).b); u(1).b(0); u(1).b(2)
    r = AddByte(u(1).b, Nxt())
    Print #f, "udt field, function:"; r; u(1).b(r)
    Print #f, "calls:"; mCalls
End Sub
