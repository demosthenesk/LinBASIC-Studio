Attribute VB_Name = "ModClass"
Option Explicit
Public gLog As String

Private Sub Bump(c As CCounter)
    c.Increment 10
End Sub

Private Sub BumpVal(ByVal c As CCounter)
    c.Increment 100
    Set c = Nothing
End Sub

Private Function Build(ByVal n As Long) As CNode
    Dim head As CNode, nd As CNode, i As Long
    For i = n To 1 Step -1
        Set nd = New CNode
        nd.Value = i
        Set nd.NextNode = head
        Set head = nd
    Next
    Set Build = head
End Function

Public Sub T_Class(f As Integer)
    Dim c As CCounter, d As CCounter
    gLog = ""
    Set c = New CCounter
    c.Name = "c1"
    c.Increment
    c.Increment 5
    Bump c
    Print #f, c.Name; c.Count; c.Twice; gLog
    BumpVal c
    c.Tag = 42
    Print #f, c.Count; c.Tag; (c Is Nothing)
    Set d = c
    d.Increment
    Print #f, c.Count; (c Is d); (c Is Nothing); TypeName(c)
    If TypeOf c Is CCounter Then Print #f, "TypeOf ok"
    Set d = Nothing
    Print #f, "after d"; gLog
    Set c = Nothing
    Print #f, "after c"; gLog

    gLog = ""
    Dim lst As CNode, p As CNode, s As String
    Set lst = Build(3)
    Set p = lst
    Do While Not p Is Nothing
        s = s & p.Value & ","
        Set p = p.NextNode
    Loop
    Print #f, s; gLog
    Set lst = Nothing
    Print #f, "freed"; gLog

    Dim o As Object, v As Variant
    gLog = ""
    Set o = New CCounter
    o.Name = "late"
    o.Increment 3
    o.Tag = "lt"
    Print #f, o.Name; o.Count; o.Tag; TypeName(o)
    Set v = o
    v.Increment
    Print #f, v.Count; IsObject(v); TypeName(v); VarType(v)
    Set o = Nothing
    Print #f, "v holds"; gLog
    v = Empty
    Print #f, "v empty"; gLog

    On Error Resume Next
    c.Increment
    Print #f, "err"; Err.Number
    Err.Clear
    Set o = New CCounter
    o.NoSuchMethod
    Print #f, "err"; Err.Number
    Err.Clear
    Set o = Nothing
    On Error GoTo 0

    Dim a As New CCounter
    gLog = ""
    a.Increment
    Print #f, a.Count; gLog; (a Is Nothing)
    Set a = Nothing
    Print #f, a.Count; gLog

    gLog = ""
    Dim w As CCounter
    Set w = New CCounter
    With w
        .Name = "w"
        .Increment 2
        Print #f, .Count; .Self.Count; .Self Is w
    End With
    Set w = Nothing
    Print #f, gLog

    Dim col As New Collection, it As Variant
    col.Add "one"
    col.Add 2, "two"
    col.Add 3.5, "three", 1
    Print #f, col.Count; col(1); col("two"); col.Item(3)
    For Each it In col
        Print #f, it;
    Next
    Print #f,
    col.Remove "two"
    col.Remove 1
    Print #f, col.Count; col(1)
    Dim arr(2) As Long, x As Variant
    arr(0) = 5: arr(1) = 6: arr(2) = 7
    For Each x In arr
        Print #f, x;
    Next
    Print #f,
    On Error Resume Next
    x = col("missing")
    Print #f, "err"; Err.Number
    Err.Clear
    col.Add 1, "one2"
    col.Add 1, "ONE2"
    Print #f, "err"; Err.Number
    Err.Clear
    x = col(9)
    Print #f, "err"; Err.Number
End Sub
