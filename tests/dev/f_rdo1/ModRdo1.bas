Attribute VB_Name = "ModRdo1"
Option Explicit
' RDO 2.0 (LinBASIC Studio phase 7.8): the objects of MSRDO20.DLL on the SQLite ODBC driver (as d_odbc1 of ADO):
' rdoEngine, rdoEnvironment, rdoConnection, rdoResultset (types, locks, cursor drivers), rdoColumn, rdoQuery and
' rdoParameter, rdoTable, rdoError, transactions, the collections

Private CS As String

Private Sub Cols(rs As rdoResultset)
    Dim c As rdoColumn
    On Error Resume Next
    For Each c In rs.rdoColumns
        WriteLog "  col " & c.Name & " type=" & c.Type & " size=" & c.Size & " attr=" & c.Attributes & " req=" & c.Required & _
            " zl=" & c.AllowZeroLength & " ord=" & c.OrdinalPosition & " src=" & c.SourceTable & "." & c.SourceColumn & _
            " upd=" & c.Updatable & " chunk=" & c.ChunkRequired & " key=" & c.KeyColumn
        Chk "col " & c.Name
    Next
End Sub

Private Sub RsInfo(ByVal what As String, rs As rdoResultset)
    On Error Resume Next
    WriteLog what & ": type=" & rs.Type & " lock=" & rs.LockType & " rowcount=" & rs.RowCount & " upd=" & rs.Updatable & _
        " bm=" & rs.Bookmarkable & " restart=" & rs.Restartable & " trans=" & rs.Transactions & " edit=" & rs.EditMode & _
        " bof=" & rs.BOF & " eof=" & rs.EOF & " name=[" & rs.Name & "]"
    Chk what & " info"
End Sub

Private Function Rows(rs As rdoResultset) As String
    Dim s As String
    On Error Resume Next
    Do Until rs.EOF
        s = s & " " & rs(0) & ":" & rs!Name
        rs.MoveNext
    Loop
    Chk "rows"
    Rows = s
End Function

Private Sub Fill(cn As rdoConnection)
    On Error Resume Next
    cn.Execute "DELETE FROM People"
    cn.Execute "INSERT INTO People (ID, Name, Age, Born, Salary, Rate, Active, Notes) VALUES (1, 'Anna', 31, '1993-04-05', 1234.5, 0.25, 1, 'first')"
    cn.Execute "INSERT INTO People (ID, Name, Age, Born, Salary, Rate, Active) VALUES (2, 'Bill', 25, '1999-12-31', 99.99, 1, 0)"
    cn.Execute "INSERT INTO People (ID, Name, Age, Active, Notes) VALUES (3, 'Cleo', 42, 1, 'third')"
    cn.Execute "INSERT INTO People (ID, Name, Age, Active) VALUES (4, 'Dino', 37, 0)"
    Chk "fill"
End Sub

Sub Main()
    Dim en As rdoEnvironment, en2 As rdoEnvironment, cn As rdoConnection, cn2 As rdoConnection, rs As rdoResultset
    Dim q As rdoQuery, q2 As rdoQuery, t As rdoTable, p As String, i As Integer, x As Variant, s As String, bm As Variant
    Dim drv As Variant, lk As Variant, ty As Variant
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    On Error Resume Next
    p = App.Path & "\rdo1.db"
    If Dir(p) <> "" Then Kill p
    CS = "Driver={SQLite3 ODBC Driver};Database=" & p

    WriteLog "== engine"
    WriteLog "version=" & rdoEngine.rdoVersion & " cursor=" & rdoEngine.rdoDefaultCursorDriver & " login=" & _
        rdoEngine.rdoDefaultLoginTimeout & " user=[" & rdoEngine.rdoDefaultUser & "] pwd=[" & rdoEngine.rdoDefaultPassword & _
        "] errthr=" & rdoEngine.rdoDefaultErrorThreshold & " locale=" & rdoEngine.rdoLocaleID & " envs=" & _
        rdoEngine.rdoEnvironments.Count & " errs=" & rdoEngine.rdoErrors.Count
    Chk "engine"
    Set en = rdoEngine.rdoEnvironments(0)
    WriteLog "env0 name=[" & en.Name & "] user=[" & en.UserName & "] cursor=" & en.CursorDriver & " login=" & en.LoginTimeout & _
        " cns=" & en.rdoConnections.Count
    Chk "env0"
    Set en2 = rdoEngine.rdoCreateEnvironment("Second", "bob", "secret")
    WriteLog "env2 name=[" & en2.Name & "] user=[" & en2.UserName & "] envs=" & rdoEngine.rdoEnvironments.Count & _
        " item=" & rdoEngine.rdoEnvironments("Second").Name
    Chk "env2"
    en2.Close
    WriteLog "after close envs=" & rdoEngine.rdoEnvironments.Count
    Chk "env2 close"

    WriteLog "== connections"
    Set cn = en.OpenConnection("", rdDriverNoPrompt, False, CS)
    Chk "open"
    WriteLog "name=[" & Mid(cn.Name, InStrRev(cn.Name, "\") + 1) & "] upd=" & cn.Updatable & " trans=" & cn.Transactions & _
        " ver=" & cn.Version & " cursor=" & cn.CursorDriver & " qtime=" & cn.QueryTimeout & " login=" & cn.LoginTimeout & _
        " cns=" & en.rdoConnections.Count & " exec=" & cn.StillExecuting & " connecting=" & cn.StillConnecting
    Chk "cn props"
    WriteLog "connect starts=" & (Left(cn.Connect, Len("Driver={SQLite3 ODBC Driver};Database=")) = "Driver={SQLite3 ODBC Driver};Database=")
    Set cn2 = en.OpenConnection("", rdDriverNoPrompt, False, "Driver={No Such Driver};Database=x")
    Chk "bad driver"
    WriteLog "cn2 nothing=" & (cn2 Is Nothing)
    Set cn2 = New rdoConnection
    WriteLog "new cn2 name=[" & cn2.Name & "] connect=[" & cn2.Connect & "] cursor=" & cn2.CursorDriver
    Chk "new"
    cn2.Connect = CS
    cn2.CursorDriver = rdUseOdbc
    cn2.EstablishConnection rdDriverNoPrompt, True
    Chk "establish"
    WriteLog "established upd=" & cn2.Updatable & " cns=" & en.rdoConnections.Count
    cn2.Close
    WriteLog "closed cns=" & en.rdoConnections.Count
    Chk "close"

    WriteLog "== execute"
    cn.Execute "CREATE TABLE People (ID INTEGER PRIMARY KEY, Name VARCHAR(20) NOT NULL, Age SMALLINT, Born DATE, " & _
        "Salary NUMERIC(10,2), Rate DOUBLE, Active BIT, Notes TEXT, UNIQUE(Name))"
    Chk "create"
    WriteLog "create affected=" & cn.RowsAffected
    Fill cn
    cn.Execute "UPDATE People SET Age = Age + 1 WHERE Age > 30"
    WriteLog "update affected=" & cn.RowsAffected
    cn.Execute "DELETE FROM People WHERE Name = 'Nobody'", rdExecDirect
    WriteLog "delete affected=" & cn.RowsAffected
    cn.Execute "INSERT INTO People (ID, Name) VALUES (1, 'Dup')"
    Chk "dup key"
    cn.Execute "SELEC nonsense"
    Chk "syntax"
    cn.Execute "SELECT * FROM Nowhere"
    Chk "no table"
    WriteLog "errs after=" & rdoEngine.rdoErrors.Count

    WriteLog "== resultsets"
    For Each drv In Array(rdUseIfNeeded, rdUseOdbc, rdUseServer, rdUseClientBatch, rdUseNone)
        For Each ty In Array(rdOpenForwardOnly, rdOpenKeyset, rdOpenDynamic, rdOpenStatic)
            For Each lk In Array(rdConcurReadOnly, rdConcurValues, rdConcurLock, rdConcurRowVer, rdConcurBatch)
                Set cn2 = New rdoConnection
                cn2.Connect = CS
                cn2.CursorDriver = drv
                cn2.EstablishConnection rdDriverNoPrompt
                Set rs = cn2.OpenResultset("SELECT ID, Name FROM People ORDER BY ID", ty, lk)
                If Err.Number <> 0 Then
                    WriteLog "drv " & drv & " type " & ty & " lock " & lk & " ! " & Err.Number & " " & Err.Description
                    Err.Clear
                Else
                    s = ""
                    s = Rows(rs)
                    WriteLog "drv " & drv & " type " & ty & " lock " & lk & " -> type=" & rs.Type & " lock=" & rs.LockType & _
                        " rowcount=" & rs.RowCount & " upd=" & rs.Updatable & " bm=" & rs.Bookmarkable & " rows=" & s
                    Chk "info"
                    rs.Close
                End If
                cn2.Close
                Err.Clear
            Next
        Next
    Next

    WriteLog "== columns and values"
    Set rs = cn.OpenResultset("SELECT * FROM People ORDER BY ID", rdOpenKeyset, rdConcurValues)
    Chk "open"
    RsInfo "keyset", rs
    Cols rs
    For i = 0 To rs.rdoColumns.Count - 1
        WriteLog "  value " & rs.rdoColumns(i).Name & "=" & V(rs.rdoColumns(i).Value)
    Next
    WriteLog "  by name=" & V(rs("Name")) & " bang=" & V(rs!Age) & " item=" & V(rs.rdoColumns.Item("Salary").Value)
    WriteLog "  nope=" & V(rs!Nope)
    Chk "nope"
    WriteLog "  col 99=" & V(rs.rdoColumns(99).Value)
    Chk "col 99"
    rs.MoveLast
    RsInfo "movelast", rs
    WriteLog "  abs=" & rs.AbsolutePosition & " pct=" & rs.PercentPosition
    Chk "abs"
    rs.MoveFirst
    rs.Move 2
    WriteLog "  move 2 -> " & rs!Name & " abs=" & rs.AbsolutePosition
    bm = rs.Bookmark
    rs.MoveFirst
    rs.Bookmark = bm
    WriteLog "  bookmark -> " & rs!Name
    Chk "bookmark"
    rs.MovePrevious: rs.MovePrevious: rs.MovePrevious
    RsInfo "before first", rs
    rs.MovePrevious
    Chk "move before bof"
    rs.Close
    WriteLog "  closed name=" & rs.Name
    Chk "closed rs"

    WriteLog "== edits"
    Set rs = cn.OpenResultset("SELECT * FROM People ORDER BY ID", rdOpenKeyset, rdConcurValues)
    rs.Edit
    WriteLog "edit mode=" & rs.EditMode
    rs!Age = 50
    rs!Notes = Null
    rs.Update
    Chk "update"
    WriteLog "after update mode=" & rs.EditMode & " age=" & V(rs!Age) & " notes=" & V(rs!Notes)
    rs.AddNew
    WriteLog "addnew mode=" & rs.EditMode & " id=" & V(rs!ID)
    rs!ID = 5
    rs!Name = "Eve"
    rs!Age = 28
    rs.Update
    Chk "addnew"
    rs.Bookmark = rs.LastModified
    WriteLog "lastmodified -> " & V(rs!Name)
    Chk "lastmodified"
    rs.MoveFirst
    rs.Edit
    rs!Name = "Bill"
    rs.Update
    Chk "update dup"
    WriteLog "after dup mode=" & rs.EditMode
    rs.CancelUpdate
    rs.Edit
    rs!Age = 99
    rs.CancelUpdate
    WriteLog "cancelupdate age=" & V(rs!Age) & " mode=" & rs.EditMode
    rs.MoveNext
    rs.Delete
    Chk "delete"
    WriteLog "after delete name=" & V(rs!Name)
    Chk "after delete"
    rs.MoveNext
    WriteLog "next name=" & V(rs!Name)
    rs.Requery
    Chk "requery"
    WriteLog "requery rows=" & Rows(rs)
    rs.Close
    Set rs = cn.OpenResultset("SELECT * FROM People", rdOpenForwardOnly, rdConcurReadOnly)
    rs.Edit
    Chk "edit readonly"
    rs.MoveFirst
    Chk "movefirst forward"
    rs.Close
    Set rs = cn.OpenResultset("SELECT ID, Name, Age FROM People ORDER BY ID", rdOpenStatic, rdConcurReadOnly)
    x = rs.GetRows(2)
    WriteLog "getrows ub=" & UBound(x, 1) & "," & UBound(x, 2) & " first=" & x(1, 0) & " second=" & x(1, 1) & " now=" & V(rs!Name)
    s = rs.GetClipString(10, ";", "|", "<null>")
    WriteLog "clip=" & s
    Chk "clip"
    rs.Close

    WriteLog "== queries"
    Set q = cn.CreateQuery("Older", "SELECT Name FROM People WHERE Age > ? ORDER BY ID")
    Chk "create query"
    WriteLog "q name=" & q.Name & " sql=" & q.SQL & " params=" & q.rdoParameters.Count & " type=" & q.Type & " queries=" & _
        cn.rdoQueries.Count & " prepared=" & q.Prepared & " maxrows=" & q.MaxRows & " rowset=" & q.RowsetSize & _
        " cursor=" & q.CursorType & " lock=" & q.LockType & " timeout=" & q.QueryTimeout
    Chk "q props"
    WriteLog "param name=" & q.rdoParameters(0).Name & " dir=" & q.rdoParameters(0).Direction & " type=" & q.rdoParameters(0).Type
    Chk "param"
    q.rdoParameters(0) = 30
    Set rs = q.OpenResultset(rdOpenStatic)
    WriteLog "older than 30:" & Rows(rs)
    rs.Close
    q(0) = 40
    Set rs = q.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    WriteLog "older than 40:" & Rows(rs)
    rs.Close
    Set q2 = New rdoQuery
    Set q2.ActiveConnection = cn
    q2.SQL = "UPDATE People SET Age = ? WHERE Name = ?"
    q2(0) = 77
    q2(1) = "Anna"
    q2.Execute
    Chk "q2 execute"
    WriteLog "q2 affected=" & q2.RowsAffected & " params=" & q2.rdoParameters.Count & " queries=" & cn.rdoQueries.Count
    Set rs = cn.OpenResultset("SELECT Age FROM People WHERE Name = 'Anna'")
    WriteLog "anna age=" & V(rs!Age) & " type=" & rs.Type & " lock=" & rs.LockType
    rs.Close
    q.Close
    WriteLog "after close queries=" & cn.rdoQueries.Count
    Set rs = cn.OpenResultset("{call Nothing}")
    Chk "call"

    WriteLog "== tables"
    cn.rdoTables.Refresh
    Chk "tables refresh"
    WriteLog "tables=" & cn.rdoTables.Count
    For Each t In cn.rdoTables
        WriteLog "  table " & t.Name & " type=" & t.Type & " cols=" & t.rdoColumns.Count & " upd=" & t.Updatable & " rowcount=" & t.RowCount
        Chk "table"
    Next
    Set rs = cn.rdoTables("People").OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    Chk "table rs"
    WriteLog "table rs rows=" & Rows(rs)
    rs.Close

    WriteLog "== transactions"
    cn.BeginTrans
    cn.Execute "DELETE FROM People"
    cn.RollbackTrans
    Chk "rollback"
    Set rs = cn.OpenResultset("SELECT COUNT(*) FROM People")
    WriteLog "after rollback count=" & V(rs(0))
    rs.Close
    cn.BeginTrans
    cn.Execute "DELETE FROM People WHERE ID = 4"
    cn.CommitTrans
    Chk "commit"
    Set rs = cn.OpenResultset("SELECT COUNT(*) FROM People")
    WriteLog "after commit count=" & V(rs(0))
    rs.Close
    cn.CommitTrans
    Chk "commit without begin"
    en.BeginTrans
    cn.Execute "DELETE FROM People WHERE ID = 1"
    en.RollbackTrans
    Chk "env rollback"
    Set rs = cn.OpenResultset("SELECT COUNT(*) FROM People")
    WriteLog "after env rollback count=" & V(rs(0))
    rs.Close

    WriteLog "== more results"
    Set rs = cn.OpenResultset("SELECT ID FROM People WHERE ID = 1; SELECT Name FROM People WHERE ID = 2")
    Chk "multi"
    WriteLog "first=" & V(rs(0))
    WriteLog "more=" & rs.MoreResults
    Chk "more"
    WriteLog "second=" & V(rs(0))
    rs.Close

    WriteLog "== collections and close"
    Set rs = cn.OpenResultset("SELECT * FROM People")
    WriteLog "resultsets=" & cn.rdoResultsets.Count & " cns=" & en.rdoConnections.Count
    cn.Close
    WriteLog "after cn close resultsets name=" & rs.Name
    Chk "rs after cn close"
    cn.Execute "SELECT 1"
    Chk "closed cn"
    WriteLog "cns=" & en.rdoConnections.Count
    WriteLog "== end"
End Sub
