Attribute VB_Name = "ModOdbc2"
Option Explicit

Private Sub Server(cn As ADODB.Connection, ByVal title As String)
    Dim r As ADODB.Recordset
    Set r = cn.Execute("SELECT ID, Name, Age, Born, Salary, Notes FROM T ORDER BY ID")
    DumpRest r, title
    r.Close
End Sub

Private Function EM(rs As ADODB.Recordset) As String
    EM = "edit=" & rs.EditMode
    If Err.Number <> 0 Then EM = "edit=?" & Err.Number: Err.Clear
End Function

Sub Main()
    Dim cn As ADODB.Connection, rs As ADODB.Recordset, p As String
    OpenOut "d_odbc2"
    On Error Resume Next
    p = App.Path & "\odbc2.db"
    If Dir(p) <> "" Then Kill p
    Set cn = New ADODB.Connection
    cn.Open "Driver={SQLite3 ODBC Driver};Database=" & p
    Chk "open"
    cn.Execute "CREATE TABLE T (ID INTEGER PRIMARY KEY AUTOINCREMENT, Name VARCHAR(20), Age SMALLINT, Born DATE, Salary NUMERIC(10,2), Notes TEXT)"
    cn.Execute "INSERT INTO T (Name, Age, Born, Salary, Notes) VALUES ('Ann', 30, '2000-01-01', 10.5, 'n1')"
    cn.Execute "INSERT INTO T (Name, Age) VALUES ('Bob', 40)"
    cn.Execute "INSERT INTO T (Name, Age, Born, Salary, Notes) VALUES ('Cat', 50, '2001-02-03', 20, 'n3')"
    cn.Execute "INSERT INTO T (Name, Age, Born, Salary) VALUES ('Dan', 60, '2002-03-04', 30)"
    cn.Execute "INSERT INTO T (Name, Age, Salary, Notes) VALUES ('Eli', 70, 40, 'n5')"
    Chk "setup"

    Out "== client"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM T ORDER BY ID", cn, adOpenStatic, adLockOptimistic
    Chk "open client"
    Out "update criteria=" & rs.Properties("Update Criteria").Value
    Chk "update criteria"
    rs.Move 1
    Out "c1 at " & Pos(rs)
    rs!Born = #1/2/2003#
    Chk "c1 born"
    Out "c1 " & EM(rs)
    rs!Salary = 5.25
    Chk "c1 salary"
    rs.Update
    Chk "c1 update"
    Out "c1 " & EM(rs) & " born=" & V(rs!Born.Value) & " salary=" & V(rs!Salary.Value)
    rs.CancelUpdate
    Chk "c1 cancel"
    rs.MoveFirst
    rs!Salary = 11.5
    rs.Update
    Chk "c2 salary of a row with values"
    Out "c2 " & EM(rs) & " salary=" & V(rs!Salary.Value)
    rs.CancelUpdate
    rs.MoveFirst
    rs.Move 2
    rs!Born = #5/6/2004#
    rs.Update
    Chk "c3 date of a row with values"
    Out "c3 " & Pos(rs) & " " & EM(rs) & " born=" & V(rs!Born.Value)
    rs.CancelUpdate
    cn.Execute "UPDATE T SET Notes = 'srv' WHERE ID = 4"
    rs.MoveFirst
    rs.Move 3
    rs!Age = 61
    rs.Update
    Chk "c4 another column changed on the server"
    Out "c4 " & Pos(rs) & " " & EM(rs)
    rs.CancelUpdate
    cn.Execute "UPDATE T SET Age = 71 WHERE ID = 5"
    rs.MoveLast
    rs!Age = 72
    rs.Update
    Chk "c5 the same column changed on the server"
    Out "c5 " & Pos(rs) & " " & EM(rs)
    rs.CancelUpdate
    rs.MoveFirst
    rs!Notes = "long"
    rs.Update
    Chk "c6 notes"
    rs.CancelUpdate
    rs.MoveFirst
    rs!Age = 33
    rs.Update
    Chk "c7 first again"
    rs.CancelUpdate
    rs.AddNew
    rs!Name = "Fay"
    rs!Age = 80
    rs.Update
    Chk "c8 addnew"
    Out "c8 " & EM(rs) & " id=" & V(rs!ID.Value) & " born=" & V(rs!Born.Value) & " count=" & rs.RecordCount & " pos=" & rs.AbsolutePosition
    rs.CancelUpdate
    rs.Close
    Server cn, "server after client"

    Out "== keyset"
    Set rs = New ADODB.Recordset
    rs.Open "T", cn, adOpenKeyset, adLockOptimistic, adCmdTable
    Chk "open keyset"
    rs!Age = 31
    rs.Update
    Chk "k1 first"
    Out "k1 " & EM(rs) & " age=" & V(rs!Age.Value)
    rs.CancelUpdate
    rs!Age = 32
    rs.Update
    Chk "k2 first again"
    Out "k2 " & EM(rs) & " age=" & V(rs!Age.Value)
    rs.CancelUpdate
    rs.MoveNext
    Out "k3 at " & rs!Name.Value
    rs.Delete
    Chk "k3 delete a row with nulls"
    rs.MoveNext
    Out "k4 at " & rs!Name.Value
    rs.Delete
    Chk "k4 delete a row without nulls"
    rs.MoveNext
    Out "k5 at " & rs!Name.Value
    rs.Delete
    Chk "k5 delete a row with a null long text only"
    rs.AddNew Array("Name", "Age"), Array("Gil", 90)
    Chk "k6 addnew"
    Out "k6 " & EM(rs) & " id=" & V(rs!ID.Value) & " count=" & rs.RecordCount & " pos=" & rs.AbsolutePosition & " eof=" & rs.EOF
    rs!Age = 91
    rs.Update
    Chk "k7 update the new row"
    rs.CancelUpdate
    rs.MoveLast
    Out "k8 last " & rs!Name.Value & " count=" & rs.RecordCount
    rs.Requery
    Chk "k9 requery"
    Out "k9 count=" & rs.RecordCount
    rs.Close
    Server cn, "server after keyset"

    Out "== other cursors"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM T", cn, adOpenForwardOnly, adLockOptimistic
    Chk "forward optimistic"
    Out "fwd cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " count=" & rs.RecordCount & " addnew=" & rs.Supports(adAddNew)
    rs!Notes = "fwd"
    rs.Update
    Chk "fwd update"
    rs.Close
    rs.Open "SELECT * FROM T", cn, adOpenDynamic, adLockOptimistic
    Chk "dynamic"
    Out "dyn cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " count=" & rs.RecordCount
    rs.Close
    rs.Open "SELECT * FROM T", cn, adOpenKeyset, adLockPessimistic
    Chk "pessimistic"
    Out "pess cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " count=" & rs.RecordCount
    rs.Close
    rs.Open "SELECT * FROM T", cn, adOpenStatic, adLockReadOnly
    Chk "static readonly"
    Out "ro cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " count=" & rs.RecordCount
    rs!Age = 1
    Chk "ro assign"
    rs.Close
    rs.Open "T", cn
    Chk "bare name"
    Out "bare state=" & rs.State & " source=" & rs.Source
    rs.Close
    rs.Open "SELECT T.Name, U.Name FROM T, T AS U WHERE T.ID = U.ID", cn, adOpenKeyset, adLockOptimistic
    Chk "join"
    Out "join fields=" & rs.Fields(0).Name & "," & rs.Fields(1).Name & " attr=" & rs.Fields(0).Attributes
    rs.Fields(0).Value = "Zed"
    rs.Update
    Chk "join update"
    rs.Close
    cn.CursorLocation = adUseClient
    Set rs = cn.Execute("SELECT * FROM T")
    Out "cn client execute cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " count=" & rs.RecordCount
    rs.Close
    Set rs = New ADODB.Recordset
    rs.Open "SELECT Name, Age * 2 AS Twice FROM T", cn, adOpenStatic, adLockOptimistic
    Out "expr attr=" & rs.Fields(0).Attributes & "," & rs.Fields(1).Attributes
    rs!Twice = 5
    Chk "expr assign"
    rs.Update
    Chk "expr update"
    rs.CancelUpdate
    rs.Close
    Server cn, "server at end"
    cn.Close
    Out "end"
    CloseOut
End Sub
