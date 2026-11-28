Attribute VB_Name = "ModJet"
Option Explicit
' Phase 14: the .mdb file of a program. A database made (CreateDatabase, Jet 4 and Jet 3) with its tables, fields,
' indexes, a relation, queries, properties and rows; closed and opened again from the file itself (the copy LinBASIC
' keeps next to it, file.mdb.lbdb, deleted first); changed (ALTER TABLE, rows, an index, a query, properties), closed
' and opened again; CompactDatabase to the same and to the other version.

Private Sub NoCopy(ByVal p As String)
    ' LinBASIC works on a copy of a Jet file next to it (p.lbdb): without it the .mdb itself is read again
    On Error Resume Next
    Kill p & ".lbdb"
    Err.Clear
End Sub

Private Function Hx(x As Variant) As String
    Dim b() As Byte, i As Long, s As String
    If IsNull(x) Then Hx = "Null": Exit Function
    b = x
    s = "L" & (UBound(b) - LBound(b) + 1) & ":"
    For i = LBound(b) To UBound(b)
        If i - LBound(b) >= 8 Then s = s & "..": Exit For
        s = s & Right$("0" & Hex$(b(i)), 2)
    Next
    Hx = s
End Function

Private Function Txt(x As Variant) As String
    Dim i As Long, s As String, c As Long
    If IsNull(x) Then Txt = "Null": Exit Function
    s = "S" & Len(x) & "'"
    For i = 1 To Len(x)
        If i > 24 Then s = s & "..": Exit For
        c = AscW(Mid$(x, i, 1)) And &HFFFF&
        If c < 32 Or c > 126 Then s = s & "<" & Hex$(c) & ">" Else s = s & Mid$(x, i, 1)
    Next
    Txt = s & "'"
End Function

Private Function Prop(o As Object, ByVal nm As String) As String
    Dim x As Variant
    On Error Resume Next
    x = o.Properties(nm).Value
    If Err.Number <> 0 Then Prop = "-": Err.Clear Else Prop = V(x)
End Function

Private Sub AddLine(db As Database, ByVal item As Long, ByVal qty As Integer, ByVal t As String)
    db.Execute "INSERT INTO Lines (ItemID, Qty, Txt) VALUES (" & item & ", " & qty & ", '" & t & "')"
End Sub

Private Sub Build(ByVal path As String, ByVal ver As Long, ByVal uni As Boolean)
    Dim db As Database, td As TableDef, f As Field, ix As Index, rel As Relation, qd As QueryDef, rs As Recordset
    Dim b() As Byte, i As Integer, k As Long
    On Error Resume Next
    If Dir(path) <> "" Then Kill path
    NoCopy path
    Set db = DBEngine.CreateDatabase(path, dbLangGeneral, ver)
    Chk "create"
    Set td = db.CreateTableDef("Items")
    Set f = td.CreateField("ID", dbLong): f.Attributes = dbAutoIncrField: td.Fields.Append f
    Set f = td.CreateField("Name", dbText, 40): f.Required = True: td.Fields.Append f
    Set f = td.CreateField("Notes", dbMemo): f.AllowZeroLength = True: td.Fields.Append f
    Set f = td.CreateField("Flag", dbBoolean): f.DefaultValue = "True": td.Fields.Append f
    td.Fields.Append td.CreateField("B", dbByte)
    td.Fields.Append td.CreateField("I", dbInteger)
    Set f = td.CreateField("L", dbLong): f.ValidationRule = ">=0": f.ValidationText = "not negative": td.Fields.Append f
    td.Fields.Append td.CreateField("Cur", dbCurrency)
    td.Fields.Append td.CreateField("Sng", dbSingle)
    td.Fields.Append td.CreateField("Dbl", dbDouble)
    td.Fields.Append td.CreateField("Dt", dbDate)
    td.Fields.Append td.CreateField("Bin", dbBinary, 6)
    td.Fields.Append td.CreateField("Ole", dbLongBinary)
    td.Fields.Append td.CreateField("G", dbGUID)
    Set ix = td.CreateIndex("PrimaryKey"): ix.Primary = True: ix.Fields.Append ix.CreateField("ID"): td.Indexes.Append ix
    Set ix = td.CreateIndex("byName"): ix.Unique = True
    Set f = ix.CreateField("Name"): f.Attributes = dbDescending: ix.Fields.Append f
    td.Indexes.Append ix
    Set ix = td.CreateIndex("byCurDt"): ix.IgnoreNulls = True
    ix.Fields.Append ix.CreateField("Cur"): ix.Fields.Append ix.CreateField("Dt")
    td.Indexes.Append ix
    td.ValidationRule = "[B]<250"
    td.ValidationText = "B too big"
    db.TableDefs.Append td
    Chk "Items"
    db.Execute "CREATE TABLE Lines (LineID COUNTER CONSTRAINT LinesKey PRIMARY KEY, ItemID LONG, Qty INTEGER, Txt TEXT(20))"
    Chk "Lines"
    Set rel = db.CreateRelation("ItemLines", "Items", "Lines", dbRelationDeleteCascade)
    Set f = rel.CreateField("ID"): f.ForeignName = "ItemID": rel.Fields.Append f
    db.Relations.Append rel
    Chk "relation"
    Set qd = db.CreateQueryDef("qBig", "SELECT Name, Cur FROM Items WHERE Cur > 10 ORDER BY Name;")
    Set qd = db.CreateQueryDef("qParam", "PARAMETERS pMin Long; SELECT Items.Name, Lines.Qty FROM Items INNER JOIN Lines ON Items.ID = Lines.ItemID WHERE Lines.Qty >= [pMin] ORDER BY Lines.Qty DESC, Items.Name;")
    Set qd = db.CreateQueryDef("qUpd", "UPDATE Lines SET Qty = Qty + 1 WHERE Qty < 5;")
    Set qd = db.CreateQueryDef("qCount", "SELECT ItemID, Count(*) AS N, Sum(Qty) AS Total FROM Lines GROUP BY ItemID HAVING Count(*) > 0;")
    Chk "queries"
    Set td = db.TableDefs("Items")
    td.Properties.Append td.CreateProperty("Description", dbText, "the items")
    td.Fields("Name").Properties.Append td.Fields("Name").CreateProperty("Caption", dbText, "Item name")
    td.Fields("Cur").Properties.Append td.Fields("Cur").CreateProperty("DecimalPlaces", dbByte, 2)
    Chk "properties"
    Set rs = db.OpenRecordset("Items", dbOpenDynaset)
    For i = 1 To 6
        rs.AddNew
        rs!Name = "item " & i
        If i = 2 Then
            If uni Then rs!Name = "item " & ChrW(&H3B1) & ChrW(&H3B2) & ChrW(&H3A9) & " " & ChrW(&H20AC) Else rs!Name = "item caf" & Chr(233)
        End If
        If i < 6 Then
            rs!Notes = String(i * 1500 - 1400, Chr(64 + i))
            rs!Flag = (i Mod 2 = 0)
            rs!B = i * 40
            rs!I = i * 100 - 300
            rs!L = i * 100000
            rs!Cur = i * 7.25
            rs!Sng = i / 3
            rs!Dbl = -i * 1.5E+20
            rs!Dt = DateSerial(1890 + i * 20, i, i) + TimeSerial(i, i * 7, i * 11)
            ReDim b(0 To i)
            For k = 0 To i: b(k) = (k * 37 + i) And 255: Next
            rs!Bin = b
            ReDim b(0 To i * 3000 - 1)
            For k = 0 To UBound(b): b(k) = (k * 7 + i) And 255: Next
            rs!Ole = b
            If i <= 3 Then rs!G = "{6B29FC40-CA47-1067-B31D-00DD0106620" & i & "}"
        End If
        rs.Update
        Chk "row " & i
    Next
    rs.Close
    For i = 1 To 5
        For k = 1 To i
            AddLine db, i, k * 2, "line " & i & "." & k
        Next
    Next
    Chk "lines"
    db.Execute "DELETE FROM Items WHERE Name = 'item 4'"
    db.Execute "UPDATE Items SET Cur = Cur * 2 WHERE ID <= 2"
    db.QueryDefs("qUpd").Execute
    Chk "changes"
    db.Close
End Sub

Private Sub Change(ByVal path As String)
    Dim db As Database, td As TableDef, qd As QueryDef, rs As Recordset, f As Field
    On Error Resume Next
    Set db = DBEngine.OpenDatabase(path)
    Chk "open"
    db.Execute "ALTER TABLE Items ADD COLUMN Extra TEXT(10)"
    db.Execute "UPDATE Items SET Extra = 'e' & ID"
    Set rs = db.OpenRecordset("Items", dbOpenDynaset)
    rs.AddNew
    rs!Name = "item 7"
    rs!Cur = 99
    rs.Update
    rs.Close
    db.Execute "DROP INDEX byCurDt ON Items"
    db.Execute "CREATE TABLE T2 (K TEXT(5) CONSTRAINT T2Key PRIMARY KEY, D DATETIME, M MEMO)"
    db.Execute "INSERT INTO T2 VALUES ('a', #2001-02-03 04:05:06#, 'memo a')"
    db.Execute "INSERT INTO T2 (K) VALUES ('b')"
    db.Execute "CREATE INDEX byD ON T2 (D)"
    db.QueryDefs.Delete "qCount"
    db.QueryDefs("qBig").SQL = "SELECT Name FROM Items WHERE Cur > 20;"
    Set qd = db.CreateQueryDef("qNew", "SELECT DISTINCT Qty FROM Lines ORDER BY Qty;")
    Set td = db.TableDefs("Items")
    td.Fields("Name").Properties.Delete "Caption"
    td.Properties("Description").Value = "the items (changed)"
    Chk "changes"
    ' validation rules: a field's and the table's, with and without their text
    Set td = db.CreateTableDef("T3")
    Set f = td.CreateField("N", dbLong): f.ValidationRule = "Between 1 And 9": td.Fields.Append f
    Set f = td.CreateField("S", dbText, 10): f.ValidationRule = "Like 'a*'": f.ValidationText = "must start with a": td.Fields.Append f
    td.ValidationRule = "[N]<>5"
    db.TableDefs.Append td
    Chk "T3"
    db.Execute "INSERT INTO T3 (N, S) VALUES (3, 'abc')"
    ChkAlways "T3 ok"
    db.Execute "INSERT INTO T3 (N, S) VALUES (0, 'abc')"
    ChkAlways "T3 field N"
    db.Execute "INSERT INTO T3 (N, S) VALUES (2, 'xyz')"
    ChkAlways "T3 field S"
    db.Execute "INSERT INTO T3 (N, S) VALUES (5, 'abc')"
    ChkAlways "T3 table"
    db.Execute "INSERT INTO T3 (N) VALUES (Null)"
    ChkAlways "T3 null"
    db.Execute "INSERT INTO Items (Name, L) VALUES ('bad', -5)"
    ChkAlways "Items field L"
    db.Execute "INSERT INTO Items (Name, B) VALUES ('bad2', 255)"
    ChkAlways "Items table"
    Set rs = db.OpenRecordset("T3", dbOpenDynaset)
    rs.AddNew
    rs!N = 12
    rs.Update
    ChkAlways "T3 recordset"
    rs.CancelUpdate
    rs.Close
    db.Close
End Sub

Private Sub Dump(ByVal path As String, ByVal title As String, ByVal full As Boolean)
    Dim db As Database, td As TableDef, f As Field, ix As Index, rel As Relation, qd As QueryDef, rs As Recordset
    Dim s As String, i As Integer
    On Error Resume Next
    NoCopy path
    Out "== " & title
    Set db = DBEngine.OpenDatabase(path)
    Chk "open"
    Out "version " & db.Version
    For Each td In db.TableDefs
        If (td.Attributes And dbSystemObject) = 0 Then
            Out "table " & td.Name & " attr=" & td.Attributes & " count=" & td.RecordCount & " vr=" & V(td.ValidationRule) & " vt=" & V(td.ValidationText) & " desc=" & Prop(td, "Description")
            For Each f In td.Fields
                Out "  field " & f.Name & " type=" & f.Type & " size=" & f.Size & " attr=" & f.Attributes & " req=" & f.Required & " zl=" & f.AllowZeroLength & " pos=" & f.OrdinalPosition & " def=" & V(f.DefaultValue) & " vr=" & V(f.ValidationRule) & " vt=" & V(f.ValidationText) & " cap=" & Prop(f, "Caption") & " dec=" & Prop(f, "DecimalPlaces")
            Next
            For Each ix In td.Indexes
                s = ""
                For Each f In ix.Fields
                    s = s & IIf((f.Attributes And dbDescending) <> 0, " -", " +") & f.Name
                Next
                Out "  index " & ix.Name & " p=" & ix.Primary & " u=" & ix.Unique & " r=" & ix.Required & " in=" & ix.IgnoreNulls & " f=" & ix.Foreign & " :" & s
            Next
        End If
    Next
    Chk "tabledefs"
    For Each rel In db.Relations
        s = ""
        For Each f In rel.Fields
            s = s & " " & f.Name & "=" & f.ForeignName
        Next
        Out "relation " & rel.Name & " " & rel.Table & " -> " & rel.ForeignTable & " attr=" & rel.Attributes & " :" & s
    Next
    Chk "relations"
    For Each qd In db.QueryDefs
        If Left$(qd.Name, 1) <> "~" Then
            Out "query " & qd.Name & " type=" & qd.Type & " timeout=" & qd.ODBCTimeout
            Out "  sql " & Replace(Replace(qd.SQL, vbCrLf, " | "), vbCr, " | ")
        End If
    Next
    Chk "querydefs"
    If full Then
        Set rs = db.OpenRecordset("SELECT * FROM Items ORDER BY ID", dbOpenSnapshot)
        Do Until rs.EOF
            s = " "
            For i = 0 To rs.Fields.Count - 1
                Select Case rs.Fields(i).Type
                Case dbMemo, dbText: s = s & " " & Txt(rs.Fields(i).Value)
                Case dbBinary, dbLongBinary: s = s & " " & Hx(rs.Fields(i).Value)
                Case Else: s = s & " " & V(rs.Fields(i).Value)
                End Select
            Next
            Out s
            rs.MoveNext
        Loop
        rs.Close
        DumpRs db.OpenRecordset("SELECT * FROM Lines ORDER BY LineID", dbOpenSnapshot), "Lines"
        Set qd = db.QueryDefs("qParam")
        qd.Parameters("pMin") = 6
        DumpRs qd.OpenRecordset(dbOpenSnapshot), "qParam 6"
        For Each td In db.TableDefs
            If td.Name = "T2" Then DumpRs db.OpenRecordset("SELECT * FROM T2 ORDER BY K", dbOpenSnapshot), "T2"
        Next
    Else
        Set rs = db.OpenRecordset("SELECT Count(*) AS N, Sum(Cur) AS S FROM Items", dbOpenSnapshot)
        Out "  items " & V(rs!N) & " " & V(rs!S)
        rs.Close
    End If
    Chk "rows"
    db.Close
End Sub

Private Sub One(ByVal name As String, ByVal ver As Long, ByVal other As Long, ByVal uni As Boolean)
    Dim p As String, c As String, o As String
    On Error Resume Next
    p = App.Path & "\" & name & ".mdb"
    c = App.Path & "\" & name & "c.mdb"
    o = App.Path & "\" & name & "o.mdb"
    Build p, ver, uni
    Dump p, name & " made", True
    Change p
    Dump p, name & " changed", True
    If Dir(c) <> "" Then Kill c
    If Dir(o) <> "" Then Kill o
    DBEngine.CompactDatabase p, c
    Chk "compact"
    Dump c, name & " compacted", False
    If other <> 0 Then
        ' (Jet 4 to Jet 3: VB6 refuses this database, 3301)
        DBEngine.CompactDatabase p, o, , other
        Chk "compact other"
        Dump o, name & " other version", False
    End If
    DBEngine.CompactDatabase p, c
    ChkAlways "compact onto a file"
End Sub

Sub Main()
    OpenOut "d_jet"
    One "jet4", dbVersion40, 0, True
    One "jet3", dbVersion30, dbVersion40, False
    CloseOut
End Sub
