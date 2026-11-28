Attribute VB_Name = "ModJetSort"
Option Explicit
' Phase 14: the sort orders of Jet 4 (CreateDatabase with dbLangGreek, dbLangCzech ...): the texts of a table in the
' order of their index (Seek, MoveNext on the table), in ORDER BY, compared (=, <), the CollatingOrder; in the session
' and again from the file (the copy LinBASIC keeps, file.mdb.lbdb, deleted first); CompactDatabase into another order.

Private Sub NoCopy(ByVal p As String)
    On Error Resume Next
    Kill p & ".lbdb"
    Err.Clear
End Sub

Private Function U(ByVal s As String) As String
    Dim i As Integer
    For i = 1 To Len(s)
        If AscW(Mid$(s, i, 1)) < 128 And AscW(Mid$(s, i, 1)) >= 32 Then U = U & Mid$(s, i, 1) Else U = U & "<" & Hex$(AscW(Mid$(s, i, 1)) And &HFFFF&) & ">"
    Next
End Function

Private Function W(ByVal hexs As String) As String
    Dim i As Integer
    For i = 1 To Len(hexs) Step 4
        W = W & ChrW(CLng("&H" & Mid$(hexs, i, 4)))
    Next
End Function

Private Sub Show(db As Database, ByVal title As String)
    Dim rs As Recordset, s As String, n As Long
    On Error Resume Next
    Out "-- " & title & " CollatingOrder=" & db.CollatingOrder & " field=" & db.TableDefs("T").Fields("S").CollatingOrder
    Set rs = db.OpenRecordset("T", dbOpenTable)
    rs.Index = "IS"
    rs.MoveFirst
    s = ""
    Do Until rs.EOF
        s = s & " " & U(rs!S)
        rs.MoveNext
    Loop
    Out "  index:" & s
    Set rs = db.OpenRecordset("SELECT S FROM T ORDER BY S, ID")
    s = ""
    Do Until rs.EOF
        s = s & " " & U(rs!S)
        rs.MoveNext
    Loop
    Out "  order by:" & s
    Set rs = db.OpenRecordset("T", dbOpenTable)
    rs.Index = "IS"
    s = ""
    n = 0
    Set rs = db.OpenRecordset("SELECT S FROM T")
    Dim t As Recordset
    Set t = db.OpenRecordset("T", dbOpenTable)
    t.Index = "IS"
    Do Until rs.EOF
        t.Seek "=", rs!S
        If t.NoMatch Then s = s & " " & U(rs!S) Else n = n + 1
        rs.MoveNext
    Loop
    Out "  seek found " & n & " missing:" & s
    Chk "show"
End Sub

Private Sub One(ByVal name As String, ByVal lang As String, ByVal words As String)
    Dim p As String, db As Database, a() As String, i As Integer
    On Error Resume Next
    p = App.Path & "\s_" & name & ".mdb"
    Kill p
    NoCopy p
    Err.Clear
    Set db = DBEngine.CreateDatabase(p, lang, dbVersion40)
    Chk "create " & name
    db.Execute "CREATE TABLE T (ID LONG, S TEXT(50))"
    db.Execute "CREATE INDEX IS ON T (S)"
    a = Split(words, " ")
    For i = 0 To UBound(a)
        db.Execute "INSERT INTO T VALUES (" & i & ", '" & W(a(i)) & "')"
    Next
    Chk "rows " & name
    Out "== " & name
    Show db, "session"
    Out "  'ch' > 'h': " & db.OpenRecordset("SELECT Count(*) FROM T WHERE S > '" & W("0068") & "'").Fields(0).Value
    db.Close
    NoCopy p
    Set db = DBEngine.OpenDatabase(p)
    Show db, "read again"
    db.Close
    NoCopy p
End Sub

Sub Main()
    Dim c As String, g As String, cz As String, el As String, sl As String, jp As String, da As String
    OpenOut "d_jetsort"
    On Error Resume Next
    ' words as UTF-16 hex: General / Czech / Greek / Slovenian / Japanese / Danish
    g = "0063006f002d006f0070 0063006f006f0070 0043006f002d006f0070 00e9 0065 0045 00650065 0061006200630064 0041004200430044 0062 00e6 0061006500 007a"
    cz = "006300680061 006800610061 0063006100610061 0043006800610061 0069006100610061 010d00610061 0072006100610061 015900610061 007300610061 016100610061"
    el = "03b103b203b3 039103b103b1 03ac03b103b1 03b203b103b1 03c903b103b1 03ce03b103b1 0061006100610061 03c3 03c2"
    sl = "007a0061 017e0061 00730061 01610061 00630061 010d00610061 01070061 00640061"
    jp = "304b 30ab 304c 304b30fc 304b304b 3041 3042 304b309d 6f22 0061"
    da = "00610061 00e50062 007a 00e6 00f8 0061 00410041"
    One "general", dbLangGeneral, g
    One "czech", dbLangCzech, cz
    One "greek", dbLangGreek, el
    One "slovenian", dbLangSlovenian, sl
    One "japanese", dbLangJapanese, jp
    One "danish", dbLangNorwDan, da
    ' a General database compacted into Czech: the order of the new file
    Dim p As String, q As String, db As Database, i As Integer, a() As String
    p = App.Path & "\s_gen.mdb": q = App.Path & "\s_gen_cz.mdb"
    Kill p: Kill q: NoCopy p: NoCopy q
    Err.Clear
    Set db = DBEngine.CreateDatabase(p, dbLangGeneral, dbVersion40)
    db.Execute "CREATE TABLE T (ID LONG, S TEXT(50))"
    db.Execute "CREATE INDEX IS ON T (S)"
    a = Split(cz, " ")
    For i = 0 To UBound(a)
        db.Execute "INSERT INTO T VALUES (" & i & ", '" & W(a(i)) & "')"
    Next
    db.Close
    DBEngine.CompactDatabase p, q, dbLangCzech
    Chk "compact into Czech"
    NoCopy p: NoCopy q
    Out "== general compacted into czech"
    Set db = DBEngine.OpenDatabase(q)
    Show db, "compacted"
    db.Close
    NoCopy q
    CloseOut
End Sub
