Attribute VB_Name = "ModEvents"
Option Explicit
Public gEv As String

Private Function Describe(s As IShape) As String
    Describe = s.Name & "=" & s.Area
End Function

Public Sub T_Events(f As Integer)
    Dim sq As CSquare, ci As CCircle, sh As IShape, o As Object
    Set sq = New CSquare
    sq.Side = 3
    Set ci = New CCircle
    ci.Radius = 1
    Set sh = sq
    Print #f, sh.Name; sh.Area; " "; Describe(ci); " "; TypeName(sh)
    If TypeOf sh Is IShape Then Print #f, "sh is IShape"
    If TypeOf sh Is CSquare Then Print #f, "sh is CSquare"
    If TypeOf ci Is IShape Then Print #f, "ci is IShape"
    If Not TypeOf ci Is CSquare Then Print #f, "ci is not CSquare"
    sh.Grow 2
    Print #f, sq.Side; sh Is sq
    Set o = sh
    Print #f, o.Area; o.Name; TypeName(o); o Is sq
    Set o = sq
    Print #f, o.Side; TypeName(o)
    On Error Resume Next
    Print #f, o.Area
    Print #f, "err"; Err.Number
    Err.Clear
    Set o = sh
    Print #f, o.Side
    Print #f, "err"; Err.Number
    Err.Clear
    Dim v As Variant
    Set v = sh
    Print #f, v.Name; TypeName(v)
    Set sh = o
    Print #f, "err"; Err.Number; sh.Name
    Err.Clear
    Dim cs As CCircle
    Set cs = sh
    Print #f, "err"; Err.Number; cs Is Nothing
    Err.Clear
    Set sh = New IShape
    Print #f, "err"; Err.Number; sh.Area; "["; sh.Name; "]"
    Err.Clear
    On Error GoTo 0
    Dim shapes(1) As IShape, i As Integer, total As Double
    Set shapes(0) = sq
    Set shapes(1) = ci
    For i = 0 To 1
        total = total + shapes(i).Area
    Next
    Print #f, total

    Dim src As CSource, snk As CSink, snk2 As CSink, n As Long
    Set src = New CSource
    Set snk = New CSink
    snk.Tag = "a"
    gEv = ""
    n = src.Run(2)
    Print #f, n; "["; gEv; "]"
    snk.Attach src
    gEv = ""
    n = src.Run(5)
    Print #f, n; "["; gEv; "]"
    snk.Detach
    gEv = ""
    n = src.Run(4)
    Print #f, n; "["; gEv; "]"
    Set snk2 = New CSink
    snk2.Tag = "b"
    snk.Attach src
    snk2.Attach src
    gEv = ""
    n = src.Run(1)
    Print #f, n; "["; gEv; "]"
    Set snk = Nothing
    gEv = ""
    n = src.Run(1)
    Print #f, n; "["; gEv; "]"
End Sub
