Attribute VB_Name = "ModFiles"
Option Explicit
Public Sub T_Files(f As Integer)
    Dim g As Integer, s As String, a As String, n As Double, v As Variant, p As String, cnt As Integer
    p = App.Path & "\t_files.tmp"
    g = FreeFile
    Open p For Output As #g
    Print #g, "line one"
    Print #g, "two"; 2; "x"
    Write #g, "a,b", 1.5, True, -3
    Write #g,
    Print #g, "last";
    Close #g
    Print #f, "len"; FileLen(p)
    g = FreeFile
    Open p For Input As #g
    Do While Not EOF(g)
        Line Input #g, s
        cnt = cnt + 1
        Print #f, cnt; "["; s; "]"
    Loop
    Close #g
    Open p For Input As #g
    Line Input #g, s
    Line Input #g, s
    Input #g, a, n, v
    Print #f, "input"; "["; a; "]"; n; v; TypeName(v)
    Close #g
    Open p For Append As #g
    Print #g, "appended"
    Close #g
    Print #f, "len"; FileLen(p)
    Kill p
    On Error Resume Next
    Open p For Input As #g
    Print #f, "missing"; Err.Number
    Close #g
    Print #f, "closed"; Err.Number
End Sub
