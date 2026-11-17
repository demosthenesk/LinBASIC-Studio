Attribute VB_Name = "ModAdo1"
Option Explicit

Private Sub ShowFields(rs As ADODB.Recordset)
    Dim f As ADODB.Field
    For Each f In rs.Fields
        Out "  field " & f.Name & " type=" & f.Type & " size=" & f.DefinedSize & " attr=" & f.Attributes & _
            " prec=" & f.Precision & " scale=" & f.NumericScale
    Next
End Sub

Sub Main()
    Dim cn As ADODB.Connection, rs As ADODB.Recordset, rs2 As ADODB.Recordset, f As ADODB.Field
    Dim p As String, n As Long, s As String, i As Integer, a As Variant, bm As Variant, er As ADODB.Error
    OpenOut "d_ado1"
    On Error Resume Next
    p = MakeDb("ado1.mdb")
    Chk "makedb"

    Out "== connection"
    Set cn = New ADODB.Connection
    Out TypeName(cn) & " state=" & cn.State & " cursorloc=" & cn.CursorLocation & " mode=" & cn.Mode & _
        " timeout=" & cn.ConnectionTimeout & " cmdtimeout=" & cn.CommandTimeout & " isolation=" & cn.IsolationLevel
    Set rs = cn.Execute("SELECT * FROM People")
    ChkAlways "execute closed"
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & p
    Chk "open"
    Out "state=" & cn.State & " provider=" & cn.Provider & " errors=" & cn.Errors.Count & " defaultdb=" & V(cn.DefaultDatabase)
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & p
    ChkAlways "open again"

    Out "== execute"
    Set rs = cn.Execute("SELECT * FROM People ORDER BY ID")
    Chk "execute"
    Out TypeName(rs) & " state=" & rs.State & " cursortype=" & rs.CursorType & " locktype=" & rs.LockType & _
        " cursorloc=" & rs.CursorLocation & " count=" & rs.RecordCount & " bof=" & rs.BOF & " eof=" & rs.EOF & _
        " fields=" & rs.Fields.Count & " source=" & rs.Source
    ShowFields rs
    Out "pos " & rs.AbsolutePosition & " editmode=" & rs.EditMode
    Chk "pos"
    DumpRest rs, "people"
    Out "eof " & rs.EOF & " bof " & rs.BOF
    rs.MoveNext
    ChkAlways "movenext at eof"
    a = rs!Name
    ChkAlways "field at eof"
    rs.MovePrevious
    ChkAlways "moveprevious forward-only"
    rs.MoveFirst
    Chk "movefirst forward-only"
    Out "after movefirst " & Pos(rs)
    Chk "pos"
    rs!Name = "X"
    ChkAlways "edit read-only"
    rs.Close
    Out "closed state=" & rs.State
    a = rs.EOF
    ChkAlways "eof of closed"
    cn.Execute "UPDATE People SET Age = Age + 1 WHERE Age > 30", n
    Chk "update"
    Out "affected " & n
    cn.Execute "DELETE FROM Nope"
    ChkAlways "missing table"
    Out "errors " & cn.Errors.Count
    For Each er In cn.Errors
        Out "  error " & er.Number & " [" & er.Description & "] src=" & er.Source & " state=" & er.SQLState & " native=" & er.NativeError
    Next
    cn.Execute "SELEC 1"
    ChkAlways "syntax"
    cn.Execute "SELECT Nope FROM People"
    ChkAlways "missing column"
    cn.Execute "INSERT INTO People (Name, Age) VALUES (Null, 3)"
    ChkAlways "null required"
    cn.Execute "INSERT INTO People (ID, Name) VALUES (1, 'Dup')"
    ChkAlways "duplicate key"
    cn.Errors.Clear
    Out "errors after clear " & cn.Errors.Count
    Set rs = cn.Execute("SELECT Count(*) AS N, Max(Age) AS M, Avg(Age), Sum(Salary) FROM People")
    Chk "aggregate"
    ShowFields rs
    DumpRs rs, "aggregate"
    Set rs = cn.Execute("SELECT Name, Age FROM People WHERE Name LIKE 'B%' ORDER BY Name")
    DumpRs rs, "like %"
    Set rs = cn.Execute("SELECT Name, Age FROM People WHERE Name LIKE 'B*' ORDER BY Name")
    DumpRs rs, "like *"
    Set rs = cn.Execute("People", , adCmdTable)
    Chk "cmdtable"
    Out "table source=" & rs.Source & " count=" & rs.RecordCount
    Set rs = cn.Execute("Adults")
    Chk "query by name"
    DumpRs rs, "adults"
    Set rs = cn.Execute("SELECT * FROM People WHERE Age > ?")
    ChkAlways "parameter not given"

    Out "== keyset"
    Set rs = New ADODB.Recordset
    Out "new state=" & rs.State & " cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " cursorloc=" & rs.CursorLocation & _
        " maxrecords=" & rs.MaxRecords & " cachesize=" & rs.CacheSize & " source=[" & rs.Source & "] active=" & V(rs.ActiveConnection)
    a = rs.RecordCount
    ChkAlways "count of closed"
    rs.Open "People", cn, adOpenKeyset, adLockOptimistic, adCmdTable
    Chk "open keyset"
    Out "state=" & rs.State & " cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " count=" & rs.RecordCount & _
        " " & Pos(rs) & " editmode=" & rs.EditMode & " source=" & rs.Source & " active=" & TypeName(rs.ActiveConnection)
    Out "supports addnew=" & rs.Supports(adAddNew) & " bookmark=" & rs.Supports(adBookmark) & " moveprev=" & rs.Supports(adMovePrevious) & _
        " update=" & rs.Supports(adUpdate) & " delete=" & rs.Supports(adDelete) & " find=" & rs.Supports(adFind)
    rs.MoveLast
    Out "last " & Pos(rs)
    rs.Move -2
    Out "move -2 " & Pos(rs)
    rs.Move 10
    Out "move 10 " & Pos(rs)
    ChkAlways "move past end"
    rs.MoveFirst
    rs.MovePrevious
    Out "before first " & Pos(rs) & " bof=" & rs.BOF
    rs.MoveFirst
    bm = rs.Bookmark
    Out "bookmark " & TypeName(bm) & " " & VarType(bm)
    rs.AbsolutePosition = 3
    Out "absolute 3 " & Pos(rs)
    rs.Bookmark = bm
    Out "back to bookmark " & Pos(rs)
    rs.Find "Name = 'Cleo'"
    Out "find cleo " & Pos(rs)
    rs.Find "Name = 'Anna'"
    Out "find anna from cleo " & Pos(rs)
    rs.MoveFirst
    rs.Find "Name LIKE 'B*'"
    Out "find like * " & Pos(rs)
    rs.MoveFirst
    rs.Find "Name LIKE 'D%'"
    Out "find like % " & Pos(rs)
    rs.Find "Age > 30", 1
    Out "find skip 1 " & Pos(rs)
    rs.MoveLast
    rs.Find "Age < 30", 0, adSearchBackward
    Out "find backward " & Pos(rs)
    rs.Find "Age > 30 AND Age < 40"
    ChkAlways "find with and"
    rs.MoveFirst
    rs.Find "Nope = 1"
    ChkAlways "find bad field"
    rs.Filter = "Age > 30"
    Out "filter count=" & rs.RecordCount & " " & Pos(rs) & " filter=" & rs.Filter
    DumpRs rs, "filtered"
    rs.Filter = "Name = 'Bill' OR Name = 'Bert'"
    Out "filter or count=" & rs.RecordCount
    rs.Filter = adFilterNone
    Out "no filter count=" & rs.RecordCount & " " & Pos(rs) & " filter=" & rs.Filter
    rs.Sort = "Age"
    ChkAlways "sort server cursor"

    Out "== editing"
    rs.MoveFirst
    rs!Name = "Anna2"
    Out "after assign editmode=" & rs.EditMode & " value=" & V(rs!Name.Value) & " original=" & V(rs!Name.OriginalValue) & " underlying=" & V(rs!Name.UnderlyingValue)
    rs.MoveNext
    Out "moved editmode=" & rs.EditMode & " " & Pos(rs)
    Set rs2 = cn.Execute("SELECT Name FROM People WHERE ID = 1")
    Out "saved by move: " & V(rs2!Name.Value)
    rs.AddNew
    Out "addnew editmode=" & rs.EditMode & " count=" & rs.RecordCount & " id=" & V(rs!ID.Value) & " name=" & V(rs!Name.Value) & " eof=" & rs.EOF
    rs!Name = "Eve"
    rs!Age = 20
    Out "before update id=" & V(rs!ID.Value)
    rs.Update
    Chk "update"
    Out "updated editmode=" & rs.EditMode & " count=" & rs.RecordCount & " " & Pos(rs) & " id=" & V(rs!ID.Value)
    rs.AddNew Array("Name", "Age"), Array("Fred", 50)
    Chk "addnew with values"
    Out "addnew values editmode=" & rs.EditMode & " " & Pos(rs) & " id=" & V(rs!ID.Value) & " count=" & rs.RecordCount
    rs.Update "Age", 51
    Chk "update with value"
    Out "update value " & V(rs!Age.Value) & " editmode=" & rs.EditMode
    rs!Name = Null
    rs.Update
    ChkAlways "update null required"
    Out "after failed update editmode=" & rs.EditMode
    rs.CancelUpdate
    Out "cancelled editmode=" & rs.EditMode & " " & Pos(rs)
    rs!Age = "abc"
    ChkAlways "type mismatch"
    rs.CancelUpdate
    rs!Name = String(30, "x")
    ChkAlways "too long"
    rs.Update
    ChkAlways "too long update"
    rs.CancelUpdate
    rs.MoveFirst
    rs.Find "Name = 'Dino'"
    rs.Delete
    Chk "delete"
    Out "deleted editmode=" & rs.EditMode & " eof=" & rs.EOF & " bof=" & rs.BOF & " count=" & rs.RecordCount
    a = rs!Name
    ChkAlways "field of deleted"
    rs.MoveNext
    Out "after delete next " & Pos(rs)
    rs.Requery
    Out "requery count=" & rs.RecordCount & " " & Pos(rs)
    DumpRs rs, "after edits"
    rs.Close

    Out "== client cursor"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM People", cn, adOpenKeyset, adLockOptimistic
    Chk "open client"
    Out "cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " count=" & rs.RecordCount & " " & Pos(rs)
    ShowFields rs
    rs.Sort = "Age DESC"
    Chk "sort"
    Out "sort=" & rs.Sort & " " & Pos(rs)
    rs.MoveFirst
    s = ""
    Do Until rs.EOF
        s = s & " " & rs!Name & "/" & rs!Age
        rs.MoveNext
    Loop
    Out "sorted" & s
    rs.Sort = "Name"
    rs.MoveFirst
    Out "by name first " & Pos(rs)
    rs.Sort = ""
    rs.MoveFirst
    a = rs.GetRows(2)
    Chk "getrows"
    Out "getrows " & V(a) & " " & UBound(a, 1) & "x" & UBound(a, 2) & " " & V(a(1, 0)) & " " & V(a(1, 1)) & " " & Pos(rs)
    a = rs.GetRows(adGetRowsRest, , Array("Name", "Age"))
    Chk "getrows fields"
    Out "getrows rest " & UBound(a, 1) & "x" & UBound(a, 2) & " " & V(a(0, 0)) & " eof=" & rs.EOF
    rs.MoveFirst
    s = rs.GetString(adClipString, 2, ",", ";", "(null)")
    Chk "getstring"
    Out "getstring [" & s & "]"
    Set rs2 = rs.Clone
    Chk "clone"
    rs2.MoveLast
    Out "clone " & Pos(rs2) & " count=" & rs2.RecordCount & " orig " & Pos(rs)
    Set rs.ActiveConnection = Nothing
    Chk "disconnect"
    Out "disconnected state=" & rs.State & " count=" & rs.RecordCount
    rs.MoveFirst
    rs!Age = 77
    rs.Update
    Chk "update disconnected"
    Out "disconnected update " & V(rs!Age.Value) & " status=" & rs.Status
    rs.Close

    Out "== fields of a row"
    Set rs = cn.Execute("SELECT * FROM People WHERE ID = 1")
    For Each f In rs.Fields
        Out "  " & f.Name & " = " & V(f.Value) & " actual=" & f.ActualSize
    Next
    Chk "fields"
    a = rs.Fields(99).Value
    ChkAlways "field 99"
    a = rs.Fields("Nope").Value
    ChkAlways "field nope"
    Out "field by index " & rs.Fields(1).Name & " default " & V(rs("Name"))

    Out "== transactions"
    n = cn.BeginTrans
    Out "level " & n
    cn.Execute "UPDATE People SET Age = 0"
    cn.RollbackTrans
    Set rs = cn.Execute("SELECT Sum(Age) FROM People")
    Out "after rollback " & V(rs(0).Value)
    cn.BeginTrans
    cn.Execute "UPDATE People SET Age = Age + 1 WHERE ID = 2"
    cn.CommitTrans
    Set rs = cn.Execute("SELECT Age FROM People WHERE ID = 2")
    Out "after commit " & V(rs(0).Value)
    cn.CommitTrans
    ChkAlways "commit without begin"

    Out "== schema"
    Set rs = cn.OpenSchema(adSchemaTables)
    Chk "schema"
    Do Until rs.EOF
        If rs!TABLE_TYPE = "TABLE" Or rs!TABLE_TYPE = "VIEW" Then Out "  " & rs!TABLE_NAME & " " & rs!TABLE_TYPE
        rs.MoveNext
    Loop
    Set rs = cn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "Pets"))
    Chk "columns"
    Do Until rs.EOF
        Out "  " & rs!COLUMN_NAME & " " & rs!ORDINAL_POSITION & " " & rs!DATA_TYPE
        rs.MoveNext
    Loop

    cn.Close
    Out "closed state=" & cn.State
    cn.Close
    ChkAlways "close again"
    Out "end"
    CloseOut
End Sub
