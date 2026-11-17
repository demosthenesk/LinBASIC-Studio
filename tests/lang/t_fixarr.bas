Attribute VB_Name = "ModFixArr"
Option Explicit
' Phase 7 (robot: Dope Wars, DragonSource): arrays of fixed-length strings inside a Type - name(3) As String * 6,
' Member() As String * 4 - their elements padded and cut, ReDim, copy of the Type, Erase, Len, Get / Put; and
' Get / Put of a whole array of String * n.

Private Type HighScores
    pname(3) As String * 6
    score(3) As Currency
    n As Long
End Type

Private Type GuildRec
    Name As String * 8
    Member() As String * 4
End Type

Private Function Codes(ByVal s As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(s)
        r = r & " " & Asc(Mid$(s, i, 1))
    Next
    Codes = "(" & Len(s) & ")" & r
End Function

Public Sub T_FixArr(f As Integer)
    Dim h As HighScores, h2 As HighScores, h3 As HighScores, g As GuildRec, i As Long
    Dim a(2) As String * 3, b(2) As String * 3, fn As Integer, p As String
    Print #f, "new:"; Codes(h.pname(0))
    h.pname(1) = "Al"
    h.pname(2) = "Bartholomew"
    h.score(2) = 12.5
    Print #f, "padded: [" & h.pname(1) & "]"; Len(h.pname(1))
    Print #f, "cut: [" & h.pname(2) & "]"; Len(h.pname(2))
    h2 = h
    h.pname(1) = "Zed"
    Print #f, "copy: [" & h2.pname(1) & "] [" & h2.pname(2) & "]"; h2.score(2)
    Print #f, "len of the type:"; Len(h2)
    For i = 0 To 3
        h.pname(i) = "p" & i
    Next
    Print #f, "loop: [" & h.pname(0) & h.pname(3) & "]"
    ReDim g.Member(2)
    Print #f, "redim:"; Codes(g.Member(0))
    g.Member(1) = "abcdefgh"
    g.Name = "Knights"
    Print #f, "member: [" & g.Member(1) & "] [" & g.Name & "]"
    ReDim Preserve g.Member(4)
    Print #f, "preserve: [" & g.Member(1) & "]"; Codes(g.Member(4))
    Erase h.pname
    Print #f, "erase field:"; Codes(h.pname(1))
    a(1) = "q"
    Erase a
    Print #f, "erase array:"; Codes(a(1))
    p = App.Path & "\fixarr.bin"
    If Dir(p) <> "" Then Kill p
    fn = FreeFile
    Open p For Binary As #fn
    Put #fn, 1, h2
    Print #f, "put type:"; LOF(fn)
    Get #fn, 1, h3
    Print #f, "get type: [" & h3.pname(1) & "] [" & h3.pname(2) & "]"; h3.score(2)
    Close #fn
    Kill p
    a(0) = "x": a(1) = "yy": a(2) = "zzzz"
    fn = FreeFile
    Open p For Binary As #fn
    Put #fn, 1, a
    Print #f, "put array:"; LOF(fn)
    Get #fn, 1, b
    Print #f, "get array: [" & b(0) & b(1) & b(2) & "]"
    Close #fn
    Kill p
End Sub
