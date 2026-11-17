Attribute VB_Name = "ModUdt"
Option Explicit

Private Type TPoint
    X As Long
    Y As Long
End Type

Private Type TPerson
    Name As String
    Age As Integer
    Tags(1 To 3) As String
    Pos As TPoint
    Extra As Variant
    Scores() As Double
End Type

Private Sub MovePt(p As TPoint, ByVal dx As Long)
    p.X = p.X + dx
    p.Y = p.Y * 2
End Sub

Private Function MakePoint(ByVal x As Long, ByVal y As Long) As TPoint
    MakePoint.X = x
    MakePoint.Y = y
End Function

Private Function Describe(p As TPerson) As String
    Describe = p.Name & "(" & p.Age & ") " & p.Tags(1) & "/" & p.Tags(2) & "/" & p.Tags(3) & " @" & p.Pos.X & "," & p.Pos.Y
End Function

Public Sub T_Udt(f As Integer)
    Dim a As TPoint, b As TPoint
    a.X = 3: a.Y = 4
    b = a
    b.X = 10
    Print #f, a.X; a.Y; b.X; b.Y
    MovePt a, 5
    Print #f, a.X; a.Y
    a = MakePoint(7, 8)
    Print #f, a.X; a.Y
    Dim p As TPerson, q As TPerson
    p.Name = "Ann": p.Age = 30
    p.Tags(1) = "x": p.Tags(2) = "y": p.Tags(3) = "z"
    p.Pos = a
    p.Extra = 1.5
    ReDim p.Scores(2)
    p.Scores(2) = 9.5
    q = p
    q.Name = "Bob": q.Tags(2) = "w": q.Pos.X = 0: q.Scores(2) = 1
    Print #f, Describe(p); "|"; Describe(q)
    Print #f, p.Extra; q.Extra; p.Scores(2); q.Scores(2); UBound(q.Scores); LBound(p.Tags); UBound(p.Tags)
    Dim arr(1 To 3) As TPoint, i As Integer
    For i = 1 To 3
        arr(i).X = i: arr(i).Y = i * i
    Next
    Print #f, arr(2).X; arr(3).Y
    Dim dyn() As TPerson
    ReDim dyn(1)
    dyn(0).Name = "first": dyn(1).Name = "second": dyn(1).Pos.Y = 42
    ReDim Preserve dyn(2)
    dyn(2) = dyn(1)
    dyn(2).Name = "third"
    Print #f, dyn(0).Name; " "; dyn(1).Name; " "; dyn(2).Name; dyn(2).Pos.Y
    With p
        .Age = .Age + 1
        Print #f, .Name; .Age; .Pos.X
        With .Pos
            .Y = 99
        End With
    End With
    Print #f, p.Pos.Y; Len(p.Name)
    Erase dyn
    Print #f, "erased"
End Sub
