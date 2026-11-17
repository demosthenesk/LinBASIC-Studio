Attribute VB_Name = "ModSrv"
Option Explicit

Private Sub ShowFields(rs As ADODB.Recordset)
    Dim f As ADODB.Field
    For Each f In rs.Fields
        Out "  field " & f.Name & " type=" & f.Type & " size=" & f.DefinedSize & " attr=" & f.Attributes & " prec=" & f.Precision & " scale=" & f.NumericScale
    Next
End Sub

Private Sub One(ByVal cs As String, ByVal idcol As String)
    Dim cn As ADODB.Connection, rs As ADODB.Recordset, cmd As ADODB.Command, n As Long
    On Error Resume Next
    Set cn = New ADODB.Connection
    cn.Open cs
    Chk "open"
    Out "state=" & cn.State & " provider=" & cn.Provider & " dbms=" & cn.Properties("DBMS Name").Value
    cn.Execute "DROP TABLE lbpeople"
    Err.Clear
    cn.Execute "CREATE TABLE lbpeople (ID " & idcol & ", Name VARCHAR(20) NOT NULL UNIQUE, Age SMALLINT, Born DATE, Salary NUMERIC(10,2), Rate FLOAT, Notes TEXT)", n
    Chk "create"
    cn.Execute "INSERT INTO lbpeople (Name, Age, Born, Salary, Rate, Notes) VALUES ('Anna', 31, '1993-04-05', 1234.5, 0.25, 'first')", n
    Out "insert affected=" & n
    cn.Execute "INSERT INTO lbpeople (Name, Age) VALUES ('Bill', 25)"
    cn.Execute "INSERT INTO lbpeople (Name, Age, Born) VALUES ('Cleo', 42, '2001-02-03')"
    cn.Execute "UPDATE lbpeople SET Age = Age + 1 WHERE Age > 30", n
    Chk "dml"
    Out "update affected=" & n
    Set rs = cn.Execute("SELECT * FROM lbpeople ORDER BY ID")
    Chk "select"
    Out "cursortype=" & rs.CursorType & " count=" & rs.RecordCount
    ShowFields rs
    DumpRest rs, "rows"
    rs.Close
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM lbpeople ORDER BY ID", cn, adOpenStatic, adLockOptimistic
    Chk "client"
    rs!Age = 50
    rs!Born = #1/2/2003#
    rs.Update
    Chk "client update"
    rs.AddNew
    rs!Name = "Dino"
    rs!Age = 37
    rs!Salary = 10.5
    rs.Update
    Chk "client addnew"
    rs.MoveFirst
    rs.Move 1
    rs.Delete
    Chk "client delete"
    rs.Filter = "Age > 40"
    Out "filter count=" & rs.RecordCount
    rs.Close
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM lbpeople ORDER BY ID", cn, adOpenKeyset, adLockOptimistic
    Chk "keyset"
    rs.MoveLast
    rs!Age = 38
    rs.Update
    Chk "keyset update of a row with Nulls (MSDASQL: fails)"
    rs.CancelUpdate
    rs.MoveFirst
    rs!Age = 51
    rs.Update
    Chk "keyset update"
    rs.AddNew Array("Name", "Age"), Array("Eve", 22)
    Chk "keyset addnew"
    Out "keyset new id=" & V(rs!ID.Value) & " count=" & rs.RecordCount
    rs.Requery
    Chk "requery"
    Out "after requery count=" & rs.RecordCount
    rs.Close
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "SELECT Name, Age FROM lbpeople WHERE Age > ? ORDER BY Name"
    cmd.Parameters.Append cmd.CreateParameter("a", adInteger, adParamInput, , 30)
    Set rs = cmd.Execute
    Chk "cmd"
    DumpRest rs, "age > 30"
    rs.Close
    cmd.CommandText = "UPDATE lbpeople SET Salary = ? WHERE Name = ?"
    cmd.Parameters.Refresh
    Chk "refresh"
    Out "params=" & cmd.Parameters.Count
    cmd.Execute n, Array(99.5, "Eve")
    Chk "cmd update"
    Out "cmd update affected=" & n
    cn.BeginTrans
    cn.Execute "DELETE FROM lbpeople"
    cn.RollbackTrans
    Chk "trans"
    cn.Execute "INSERT INTO lbpeople (Name) VALUES ('Anna')"
    Chk "unique"
    cn.Execute "SELEC 1"
    Chk "syntax"
    Set rs = cn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "lbpeople"))
    Chk "schema"
    Do Until rs.EOF
        Out "  column " & rs!COLUMN_NAME & " pos=" & rs!ORDINAL_POSITION & " type=" & rs!DATA_TYPE & " nullable=" & rs!IS_NULLABLE
        rs.MoveNext
    Loop
    rs.Close
    Set rs = cn.Execute("SELECT ID, Name, Age, Born, Salary FROM lbpeople ORDER BY Name")
    DumpRest rs, "server at end"
    rs.Close
    cn.Close
End Sub

Sub Main()
    OpenOut "srv"
    On Error Resume Next
    Out "== MySQL / MariaDB"
    One "Driver={MySQL ODBC 3.51 Driver};Server=127.0.0.1;Port=3306;Database=lbtest;UID=lb;PWD=lb;Option=3", "INTEGER PRIMARY KEY AUTO_INCREMENT"
    Out "== PostgreSQL"
    One "Driver={PostgreSQL Unicode};Server=127.0.0.1;Port=5432;Database=lbtest;UID=postgres;PWD=x", "SERIAL PRIMARY KEY"
    Out "end"
    CloseOut
End Sub
