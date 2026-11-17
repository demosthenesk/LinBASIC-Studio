Attribute VB_Name = "ModFin"
Option Explicit

Private Function F8(ByVal x As Double) As String
    F8 = Format$(x, "0.00000000")
End Function

Private Sub TryDDB(f As Integer, ByVal c As Double, ByVal s As Double, ByVal l As Double, ByVal p As Double, ByVal fa As Double)
    Dim r As Double
    On Error GoTo bad
    r = DDB(c, s, l, p, fa)
    Print #f, "DDB"; c; s; l; p; fa; "="; F8(r)
    Exit Sub
bad:
    Print #f, "DDB"; c; s; l; p; fa; "error"; Err.Number
End Sub

Private Sub TrySYD(f As Integer, ByVal c As Double, ByVal s As Double, ByVal l As Double, ByVal p As Double)
    Dim r As Double
    On Error GoTo bad
    r = SYD(c, s, l, p)
    Print #f, "SYD"; c; s; l; p; "="; F8(r)
    Exit Sub
bad:
    Print #f, "SYD"; c; s; l; p; "error"; Err.Number
End Sub

Private Sub TryPartition(f As Integer, ByVal n As Long, ByVal a As Long, ByVal b As Long, ByVal i As Long)
    Dim r As String
    On Error GoTo bad
    r = Partition(n, a, b, i)
    Print #f, "Partition"; n; a; b; i; "["; r; "]"
    Exit Sub
bad:
    Print #f, "Partition"; n; a; b; i; "error"; Err.Number
End Sub

Public Sub T_Fin(f As Integer)
    Dim x As Variant, p As Long, v(4) As Double, w(5) As Double, e(2) As Double, z() As Double, h As Integer
    ' depreciation
    Print #f, "SLN "; F8(SLN(10000, 1000, 5)); " "; F8(SLN(2500, 500, 24))
    For p = 1 To 6
        TryDDB f, 2400, 300, 10, p, 2
        TrySYD f, 30000, 7500, 10, p
    Next
    TryDDB f, 2400, 300, 10, 2.5, 1.5
    TryDDB f, 1000, 100, 2, 1, 2
    TryDDB f, 1000, 100, 2, 2, 2
    TryDDB f, 1000, 100, 1, 1, 2
    TryDDB f, 0, 100, 5, 1, 2
    TryDDB f, 1000, 100, 5, 6, 2
    TryDDB f, 1000, -1, 5, 1, 2
    TryDDB f, 1000, 100, 5, 0, 2
    TryDDB f, 1000, 100, 5, 1, 0
    TryDDB f, 1000, 900, 5, 3, 2
    TrySYD f, 1000, 100, 5, 0
    TrySYD f, 1000, -5, 5, 1
    On Error Resume Next
    x = SLN(1000, 100, 0)
    Print #f, "SLN life 0"; Err.Number
    Err.Clear
    On Error GoTo 0
    ' NPV, IRR, MIRR
    v(0) = -70000: v(1) = 22000: v(2) = 25000: v(3) = 28000: v(4) = 31000
    Print #f, "NPV "; F8(NPV(0.0625, v)); " "; F8(NPV(0, v)); " "; F8(NPV(-0.5, v))
    Print #f, "IRR "; F8(IRR(v)); " "; F8(IRR(v, 0.5))
    w(0) = -120000: w(1) = 39000: w(2) = 30000: w(3) = 21000: w(4) = 37000: w(5) = 46000
    Print #f, "IRR "; F8(IRR(w)); " MIRR "; F8(MIRR(w, 0.1, 0.12)); " "; F8(MIRR(v, 0.05, 0.05))
    e(0) = 100: e(1) = 200: e(2) = 300
    On Error Resume Next
    x = IRR(e)
    Print #f, "IRR all positive"; Err.Number
    Err.Clear
    x = MIRR(e, 0.1, 0.1)
    Print #f, "MIRR no payment"; Err.Number
    Err.Clear
    x = NPV(-1, v)
    Print #f, "NPV rate -1"; Err.Number
    Err.Clear
    x = IRR(v, -1)
    Print #f, "IRR guess -1"; Err.Number
    Err.Clear
    x = IRR(z)
    Print #f, "IRR not dimensioned"; Err.Number
    Err.Clear
    On Error GoTo 0
    ' Rate
    Print #f, "Rate "; F8(Rate(48, -250, 10000)); " "; F8(Rate(10, -100, 800, 0, 1)); " "; F8(Rate(12, -1000, 0, 15000))
    Print #f, "Rate "; F8(Rate(360, -1073.64, 200000, 0, 0, 0.01))
    On Error Resume Next
    x = Rate(0, -100, 1000)
    Print #f, "Rate nper 0"; Err.Number
    Err.Clear
    On Error GoTo 0
    ' Partition
    TryPartition f, 267, 0, 500, 50
    TryPartition f, 5, 0, 99, 5
    TryPartition f, 99, 0, 99, 5
    TryPartition f, 100, 0, 99, 5
    TryPartition f, -3, 0, 99, 5
    TryPartition f, 7, 1, 9, 1
    TryPartition f, 0, 1, 9, 1
    TryPartition f, 12, 10, 1000, 7
    TryPartition f, 3, 5, 5, 1
    TryPartition f, 3, -1, 5, 1
    TryPartition f, 3, 0, 5, 0
    Print #f, TypeName(Partition(1, 0, 9, 2))
    ' FileAttr
    h = FreeFile
    Open App.Path & "\fin1.tmp" For Output As #h
    Print #f, "FileAttr output"; FileAttr(h, 1)
    Close #h
    Open App.Path & "\fin1.tmp" For Append As #h
    Print #f, "FileAttr append"; FileAttr(h, 1)
    Close #h
    Open App.Path & "\fin1.tmp" For Input As #h
    Print #f, "FileAttr input"; FileAttr(h, 1)
    Close #h
    Open App.Path & "\fin1.tmp" For Binary As #h
    Print #f, "FileAttr binary"; FileAttr(h, 1)
    Close #h
    Open App.Path & "\fin1.tmp" For Random As #h Len = 10
    Print #f, "FileAttr random"; FileAttr(h, 1)
    On Error Resume Next
    x = FileAttr(h, 3)
    Print #f, "FileAttr 3"; Err.Number
    Err.Clear
    Close #h
    x = FileAttr(h, 1)
    Print #f, "FileAttr closed"; Err.Number
    Err.Clear
    On Error GoTo 0
    Kill App.Path & "\fin1.tmp"
End Sub
