Attribute VB_Name = "ModJetDec"
Option Explicit
' Phase 14: Decimal (Jet 4). DAO cannot make a Decimal field (DDL DECIMAL: 3292, NUMERIC is a Double,
' CreateField dbDecimal: 3259); ADO (Jet OLE DB, ANSI SQL) makes them with a precision and a scale. DAO then writes
' them (SQL, AddNew, Edit) and reads them, in the session and again from the file (the copy LinBASIC keeps,
' file.mdb.lbdb, deleted first); a Jet 4 database cannot be compacted into Jet 3 (3301). The values have up to 15
' digits (LinBASIC: more are kept exactly, but computed as a Double).

Private Sub NoCopy(ByVal p As String)
    On Error Resume Next
    Kill p & ".lbdb"
    Err.Clear
End Sub

Private Function Big() As Variant
    Big = CDec(98765) + CDec(1) / CDec(4)       ' 98765.25
End Function

Private Sub Rows(db As Database)
    Dim rs As Recordset, s As String
    On Error Resume Next
    Set rs = db.OpenRecordset("SELECT * FROM D ORDER BY ID")
    If Err.Number Then Chk "open D": Exit Sub
    Do Until rs.EOF
        Out "  " & V(rs!ID.Value) & " " & V(rs!X.Value) & " " & V(rs!Y.Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT Sum(X) AS S, Max(X) AS M, Count(*) AS N FROM D")
    If Err.Number Then Chk "sum": Exit Sub
    Out "  sum=" & V(rs!S.Value) & " max=" & V(rs!M.Value) & " n=" & V(rs!N.Value)
    rs.Close
    Set rs = db.OpenRecordset("SELECT ID FROM D WHERE X > 50000 ORDER BY ID")
    If Err.Number Then Chk "big": Exit Sub
    s = ""
    Do Until rs.EOF
        s = s & " " & V(rs!ID.Value)
        rs.MoveNext
    Loop
    rs.Close
    Out "  big:" & s
    Chk "rows"
End Sub

Private Sub AdoInfo(ByVal path As String)
    Dim cn As New ADODB.Connection, rs As ADODB.Recordset, f As ADODB.Field
    On Error Resume Next
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path
    Set rs = cn.Execute("SELECT * FROM D")
    For Each f In rs.Fields
        Out "  ado " & f.Name & " type=" & f.Type & " prec=" & f.Precision & " scale=" & f.NumericScale
    Next
    rs.Close
    cn.Close
    Chk "ado info"
End Sub

Sub Main()
    Dim path As String, db As Database, rs As Recordset, td As TableDef, cn As New ADODB.Connection
    OpenOut "d_jetdec"
    On Error Resume Next
    path = App.Path & "\dec.mdb"
    Kill path
    NoCopy path
    Err.Clear
    Set db = DBEngine.CreateDatabase(path, dbLangGeneral, dbVersion40)
    db.Execute "CREATE TABLE D1 (X DECIMAL(28,4))"
    ChkAlways "DAO DDL DECIMAL(28,4)"
    db.Execute "CREATE TABLE D2 (X DECIMAL)"
    ChkAlways "DAO DDL DECIMAL"
    db.Execute "CREATE TABLE D3 (X NUMERIC)"
    ChkAlways "DAO DDL NUMERIC"
    db.Execute "CREATE TABLE D4 (X NUMERIC(10,2))"
    ChkAlways "DAO DDL NUMERIC(10,2)"
    Out "D3.X type=" & db.TableDefs("D3").Fields("X").Type & " size=" & db.TableDefs("D3").Fields("X").Size
    Set td = db.CreateTableDef("D5")
    td.Fields.Append td.CreateField("X", dbDecimal)
    ChkAlways "DAO CreateField dbDecimal"
    db.Close
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path
    cn.Execute "CREATE TABLE D (ID INT CONSTRAINT PK PRIMARY KEY, X DECIMAL(28,4), Y DECIMAL(10,2))"
    Chk "ADO DDL"
    cn.Execute "INSERT INTO D VALUES (1, 1234567890.1234, 12.5)"
    Chk "ADO insert 1"
    cn.Execute "INSERT INTO D VALUES (2, -0.0001, -99999999.99)"
    Chk "ADO insert 2"
    cn.Close
    Out "== ADO made"
    AdoInfo path
    Set db = DBEngine.OpenDatabase(path)
    Out "X type=" & db.TableDefs("D").Fields("X").Type & " size=" & db.TableDefs("D").Fields("X").Size
    db.Execute "INSERT INTO D VALUES (3, 7, 8)", dbFailOnError
    Chk "DAO insert 3"
    Set rs = db.OpenRecordset("D", dbOpenDynaset)
    rs.AddNew
    rs!ID = 4
    rs!X = Big()
    rs!Y = CDec(1) / CDec(8)
    rs.Update
    Chk "addnew 4"
    rs.FindFirst "ID = 1"
    rs.Edit
    rs!X = rs!X + CDec(9)
    rs.Update
    Chk "edit 1"
    rs.Close
    Out "== session"
    Rows db
    db.Close
    NoCopy path
    Out "== read again"
    Set db = DBEngine.OpenDatabase(path)
    Chk "open"
    Rows db
    db.Close
    AdoInfo path
    NoCopy path
    Out "== to Jet 3"
    Kill App.Path & "\dec3.mdb": Kill App.Path & "\plain4.mdb": Kill App.Path & "\plain3.mdb"
    NoCopy App.Path & "\dec3.mdb": NoCopy App.Path & "\plain4.mdb": NoCopy App.Path & "\plain3.mdb"
    Err.Clear
    DBEngine.CompactDatabase path, App.Path & "\dec3.mdb", , dbVersion30
    ChkAlways "decimal database to Jet 3"
    Set db = DBEngine.CreateDatabase(App.Path & "\plain4.mdb", dbLangGeneral, dbVersion40)
    db.Execute "CREATE TABLE P (ID LONG, S TEXT(10))"
    db.Close
    DBEngine.CompactDatabase App.Path & "\plain4.mdb", App.Path & "\plain3.mdb", , dbVersion30
    ChkAlways "plain database to Jet 3"
    Out "plain3 there=" & (Dir(App.Path & "\plain3.mdb") <> "")
    CloseOut
End Sub
