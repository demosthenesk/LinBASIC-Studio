Attribute VB_Name = "ModOdbc1"
Option Explicit

Private Sub ShowFields(rs As ADODB.Recordset)
    Dim f As ADODB.Field
    For Each f In rs.Fields
        Out "  field " & f.Name & " type=" & f.Type & " size=" & f.DefinedSize & " attr=" & f.Attributes & _
            " prec=" & f.Precision & " scale=" & f.NumericScale
    Next
End Sub

Private Sub ShowErrors(cn As ADODB.Connection)
    Dim e As ADODB.Error
    For Each e In cn.Errors
        Out "    error " & e.Number & " native=" & e.NativeError & " state=" & e.SQLState & " [" & e.Source & "] " & e.Description
    Next
End Sub

Private Function Count(cn As ADODB.Connection) As String
    Dim r As ADODB.Recordset
    Set r = cn.Execute("SELECT COUNT(*) FROM People")
    Count = V(r.Fields(0).Value)
    r.Close
End Function

Sub Main()
    Dim cn As ADODB.Connection, rs As ADODB.Recordset, cmd As ADODB.Command, p As String, n As Long, i As Integer
    Dim s As String, gr As Variant
    OpenOut "d_odbc1"
    On Error Resume Next
    p = App.Path & "\odbc1.db"
    If Dir(p) <> "" Then Kill p

    Out "== connection"
    Set cn = New ADODB.Connection
    cn.Open "Driver={SQLite3 ODBC Driver};Database=" & p
    Chk "open"
    Out "state=" & cn.State & " provider=" & cn.Provider & " cursorloc=" & cn.CursorLocation & " errors=" & cn.Errors.Count
    Out "dbms=" & cn.Properties("DBMS Name").Value
    Chk "props"
    Out "readonly=" & cn.Properties("Read-Only Data Source").Value & " sqlsupport=" & cn.Properties("SQL Support").Value
    Chk "props2"

    Out "== ddl and dml"
    cn.Execute "CREATE TABLE People (ID INTEGER PRIMARY KEY AUTOINCREMENT, Name VARCHAR(20) NOT NULL, Age SMALLINT, " & _
        "Born DATE, Salary NUMERIC(10,2), Rate DOUBLE, Active BIT, Notes TEXT, UNIQUE(Name))", n
    Chk "create"
    Out "create affected=" & n
    cn.Execute "INSERT INTO People (Name, Age, Born, Salary, Rate, Active, Notes) VALUES ('Anna', 31, '1993-04-05', 1234.5, 0.25, 1, 'first')", n
    Out "insert affected=" & n
    cn.Execute "INSERT INTO People (Name, Age, Born, Salary, Rate, Active) VALUES ('Bill', 25, '1999-12-31', 99.99, 1, 0)"
    cn.Execute "INSERT INTO People (Name, Age, Active, Notes) VALUES ('Cleo', 42, 1, 'third')"
    cn.Execute "INSERT INTO People (Name, Age, Active) VALUES ('Dino', 37, 0)"
    Chk "inserts"
    cn.Execute "UPDATE People SET Age = Age + 1 WHERE Age > 30", n
    Out "update affected=" & n
    cn.Execute "DELETE FROM People WHERE Name = 'Nobody'", n
    Out "delete affected=" & n

    Out "== forward-only (Execute)"
    Set rs = cn.Execute("SELECT * FROM People ORDER BY ID")
    Chk "execute"
    Out TypeName(rs) & " state=" & rs.State & " cursortype=" & rs.CursorType & " locktype=" & rs.LockType & _
        " cursorloc=" & rs.CursorLocation & " count=" & rs.RecordCount & " bof=" & rs.BOF & " eof=" & rs.EOF & " fields=" & rs.Fields.Count
    ShowFields rs
    DumpRest rs, "rows"
    rs.MoveFirst
    Chk "movefirst forward-only"
    rs.MovePrevious
    Chk "moveprevious forward-only"
    Out "supports bookmark=" & rs.Supports(adBookmark) & " addnew=" & rs.Supports(adAddNew) & " movep=" & rs.Supports(adMovePrevious)
    rs.Close

    Out "== client static"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM People ORDER BY ID", cn, adOpenStatic, adLockOptimistic
    Chk "open client"
    Out "cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " count=" & rs.RecordCount & " pos=" & Pos(rs)
    rs.Find "Name = 'Cleo'"
    Out "find " & Pos(rs)
    rs.Filter = "Age > 30"
    Out "filter count=" & rs.RecordCount & " " & Pos(rs)
    rs.Filter = adFilterNone
    rs.Sort = "Name DESC"
    rs.MoveFirst
    Out "sort " & Pos(rs)
    rs.Sort = ""
    rs.MoveFirst
    rs!Age = 50
    rs!Notes = "edited"
    rs.Update
    Chk "update client"
    Out "after update " & Pos(rs) & " age=" & V(rs!Age)
    rs.AddNew
    rs!Name = "Eve"
    rs!Age = 22
    rs.Update
    Chk "addnew client"
    Out "after addnew id=" & V(rs!ID) & " name=" & V(rs!Name) & " count=" & rs.RecordCount
    rs.MoveFirst
    rs.Move 1
    rs.Delete
    Chk "delete client"
    Out "after delete count=" & rs.RecordCount
    rs.Close
    Set rs = cn.Execute("SELECT ID, Name, Age, Notes FROM People ORDER BY ID")
    DumpRest rs, "server after client changes"
    rs.Close

    Out "== keyset (server)"
    Set rs = New ADODB.Recordset
    rs.Open "People", cn, adOpenKeyset, adLockOptimistic, adCmdTable
    Chk "open keyset"
    Out "cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " cursorloc=" & rs.CursorLocation & " count=" & rs.RecordCount & " source=" & rs.Source
    Out "supports bookmark=" & rs.Supports(adBookmark) & " addnew=" & rs.Supports(adAddNew) & " update=" & rs.Supports(adUpdate) & " delete=" & rs.Supports(adDelete)
    rs.MoveLast
    Out "last " & Pos(rs)
    rs!Age = 33
    rs.Update
    Chk "update keyset"
    rs.CancelUpdate
    rs.MoveFirst
    Out "first " & Pos(rs)
    rs!Age = 34
    rs.Update
    Chk "update keyset first"
    rs.AddNew Array("Name", "Age"), Array("Finn", 44)
    Chk "addnew keyset"
    Out "after addnew " & V(rs!Name) & " id=" & V(rs!ID) & " count=" & rs.RecordCount
    rs.MoveFirst
    rs.Delete
    Chk "delete keyset"
    rs.Close
    Set rs = cn.Execute("SELECT ID, Name, Age FROM People ORDER BY ID")
    DumpRest rs, "server after keyset changes"
    rs.Close
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM People ORDER BY ID", cn, adOpenStatic, adLockOptimistic
    rs.MoveLast
    rs!Born = #1/2/2003#
    rs!Salary = 5.25
    rs.Update
    Chk "client update of nulls"
    Out "client " & Pos(rs) & " born=" & V(rs!Born) & " salary=" & V(rs!Salary)
    rs.Close

    Out "== errors"
    cn.Execute "SELEC 1"
    ChkAlways "syntax"
    ShowErrors cn
    cn.Execute "SELECT * FROM Nope"
    ChkAlways "no table"
    ShowErrors cn
    cn.Execute "INSERT INTO People (Name) VALUES ('Anna')"
    ChkAlways "unique"
    ShowErrors cn
    cn.Execute "INSERT INTO People (Age) VALUES (1)"
    ChkAlways "not null"
    ShowErrors cn
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM People ORDER BY ID", cn, adOpenStatic, adLockOptimistic
    rs!Name = "Bill"
    rs.Update
    ChkAlways "unique via recordset"
    rs.CancelUpdate
    rs.Close
    Out "errors count after=" & cn.Errors.Count

    Out "== transactions"
    Out "count=" & Count(cn)
    n = cn.BeginTrans
    Out "begintrans=" & n
    cn.Execute "INSERT INTO People (Name, Age) VALUES ('Temp', 1)"
    Out "count in trans=" & Count(cn)
    cn.RollbackTrans
    Out "count after rollback=" & Count(cn)
    cn.BeginTrans
    cn.Execute "INSERT INTO People (Name, Age) VALUES ('Kept', 2)"
    cn.CommitTrans
    Out "count after commit=" & Count(cn)
    cn.CommitTrans
    ChkAlways "commit without begin"

    Out "== command"
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "SELECT Name, Age FROM People WHERE Age > ? ORDER BY Name"
    cmd.Parameters.Append cmd.CreateParameter("a", adInteger, adParamInput, , 30)
    Set rs = cmd.Execute
    Chk "cmd execute"
    DumpRest rs, "age > 30"
    rs.Close
    cmd.Parameters(0).Value = 40
    Set rs = cmd.Execute
    DumpRest rs, "age > 40"
    rs.Close
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "INSERT INTO People (Name, Age, Salary) VALUES (?, ?, ?)"
    cmd.Execute n, Array("Gus", 61, 10.5)
    Chk "cmd insert"
    Out "cmd insert affected=" & n
    cmd.CommandText = "UPDATE People SET Age = ? WHERE Name = ?"
    cmd.Parameters.Refresh
    Chk "refresh"
    Out "params=" & cmd.Parameters.Count
    cmd.Execute n, Array(62, "Gus")
    Chk "cmd update"
    Out "cmd update affected=" & n

    Out "== types"
    Set rs = cn.Execute("SELECT ID, Name, Age, Born, Salary, Rate, Active, Notes FROM People WHERE Name = 'Anna'")
    For i = 0 To rs.Fields.Count - 1
        Out "  " & rs.Fields(i).Name & " = " & V(rs.Fields(i).Value)
    Next
    rs.Close
    Set rs = cn.Execute("SELECT COUNT(*) AS N, SUM(Age) AS S, AVG(Rate) AS A, MAX(Name) AS M FROM People")
    ShowFields rs
    For i = 0 To rs.Fields.Count - 1
        Out "  " & rs.Fields(i).Name & " = " & V(rs.Fields(i).Value)
    Next
    rs.Close

    Out "== schema"
    Set rs = cn.OpenSchema(adSchemaTables)
    Chk "schema tables"
    Do Until rs.EOF
        If Left$(rs!TABLE_NAME, 7) <> "sqlite_" Then Out "  table " & rs!TABLE_NAME & " " & rs!TABLE_TYPE
        rs.MoveNext
    Loop
    rs.Close
    Set rs = cn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "People"))
    Chk "schema columns"
    Do Until rs.EOF
        Out "  column " & rs!COLUMN_NAME & " pos=" & rs!ORDINAL_POSITION & " type=" & rs!DATA_TYPE & " nullable=" & rs!IS_NULLABLE
        rs.MoveNext
    Loop
    rs.Close

    Out "== getrows and getstring"
    Set rs = cn.Execute("SELECT Name, Age FROM People ORDER BY Name")
    gr = rs.GetRows(3)
    Out "getrows " & V(gr) & " " & gr(0, 0) & "," & gr(1, 0) & " " & gr(0, 2)
    Out "getstring [" & Replace(Replace(rs.GetString(adClipString, 2, "|", "/"), vbCr, "<cr>"), vbLf, "<lf>") & "]"
    rs.Close

    cn.Close
    Out "closed state=" & cn.State
    cn.Open "Driver={No Such Driver};Database=x"
    ChkAlways "bad driver"
    cn.Open "DSN=NoSuchDsn"
    ChkAlways "bad dsn"
    Out "end"
    CloseOut
End Sub
