Attribute VB_Name = "ModJetRel"
Option Explicit
' Phase 14: the attributes of relations kept in the .mdb -- one-to-one (dbRelationUnique: a unique foreign index,
' duplicates refused), one-to-many, not enforced (dbRelationDontEnforce) -- in the session and read again from the
' file (the copy LinBASIC keeps, file.mdb.lbdb, deleted first).

Private Sub NoCopy(ByVal p As String)
    On Error Resume Next
    Kill p & ".lbdb"
    Err.Clear
End Sub

Private Sub Show(db As Database)
    Dim rl As Relation, td As TableDef, ix As Index, i As Integer, s As String
    On Error Resume Next
    For Each rl In db.Relations
        If Left$(rl.Name, 4) <> "MSys" Then
            s = ""
            For i = 0 To rl.Fields.Count - 1
                s = s & " " & rl.Fields(i).Name & ">" & rl.Fields(i).ForeignName
            Next
            Out "rel " & rl.Name & " " & rl.Table & " -> " & rl.ForeignTable & " attr=" & rl.Attributes & s
        End If
    Next
    Chk "relations"
    For Each td In db.TableDefs
        If (td.Attributes And dbSystemObject) = 0 Then
            For Each ix In td.Indexes
                s = ""
                For i = 0 To ix.Fields.Count - 1
                    s = s & " " & ix.Fields(i).Name
                Next
                Out "index " & td.Name & "." & ix.Name & " primary=" & ix.Primary & " unique=" & ix.Unique & " foreign=" & ix.Foreign & " required=" & ix.Required & " fields:" & s
            Next
        End If
    Next
    Chk "indexes"
End Sub

Private Sub Rows(db As Database, ByVal t As String)
    Dim rs As Recordset, s As String
    On Error Resume Next
    Set rs = db.OpenRecordset("SELECT * FROM [" & t & "] ORDER BY ID")
    Do Until rs.EOF
        s = s & " " & V(rs!ID.Value) & "/" & V(rs!K.Value)
        rs.MoveNext
    Loop
    rs.Close
    Out "  " & t & ":" & s
    Chk "rows " & t
End Sub

Private Sub Try(db As Database, ByVal sql As String)
    On Error Resume Next
    db.Execute sql, dbFailOnError
    Out "  " & sql & " -> " & Err.Number
    Err.Clear
End Sub

Sub Main()
    Dim path As String, db As Database, rl As Relation, f As Field
    OpenOut "d_jetrel"
    On Error Resume Next
    path = App.Path & "\rel.mdb"
    Kill path
    NoCopy path
    Err.Clear
    Set db = DBEngine.CreateDatabase(path, dbLangGeneral, dbVersion40)
    db.Execute "CREATE TABLE P (K LONG CONSTRAINT PK PRIMARY KEY, S TEXT(10))"
    db.Execute "CREATE TABLE C1 (ID LONG CONSTRAINT PK PRIMARY KEY, K LONG)"
    db.Execute "CREATE TABLE C2 (ID LONG CONSTRAINT PK PRIMARY KEY, K LONG)"
    db.Execute "CREATE TABLE C3 (ID LONG CONSTRAINT PK PRIMARY KEY, K LONG)"
    db.Execute "INSERT INTO P VALUES (1, 'one')"
    db.Execute "INSERT INTO P VALUES (2, 'two')"
    Chk "tables"
    Set rl = db.CreateRelation("R1", "P", "C1", dbRelationUnique)
    Set f = rl.CreateField("K"): f.ForeignName = "K": rl.Fields.Append f
    db.Relations.Append rl
    Chk "R1"
    Set rl = db.CreateRelation("R2", "P", "C2", 0)
    Set f = rl.CreateField("K"): f.ForeignName = "K": rl.Fields.Append f
    db.Relations.Append rl
    Chk "R2"
    Set rl = db.CreateRelation("R3", "P", "C3", dbRelationDontEnforce)
    Set f = rl.CreateField("K"): f.ForeignName = "K": rl.Fields.Append f
    db.Relations.Append rl
    Chk "R3"
    Out "== made"
    Show db
    Try db, "INSERT INTO C1 VALUES (1, 1)"
    Try db, "INSERT INTO C1 VALUES (2, 1)"
    Try db, "INSERT INTO C1 VALUES (3, 2)"
    Try db, "INSERT INTO C1 VALUES (4, Null)"
    Try db, "INSERT INTO C1 VALUES (5, Null)"
    Try db, "UPDATE C1 SET K = 1 WHERE ID = 3"
    Try db, "UPDATE C1 SET K = 1 WHERE ID = 1"
    Try db, "INSERT INTO C2 VALUES (1, 1)"
    Try db, "INSERT INTO C2 VALUES (2, 1)"
    Try db, "INSERT INTO C2 VALUES (3, 7)"
    Try db, "INSERT INTO C3 VALUES (1, 1)"
    Try db, "INSERT INTO C3 VALUES (2, 9)"
    Rows db, "C1"
    Rows db, "C2"
    Rows db, "C3"
    db.Close
    NoCopy path
    Out "== read again"
    Set db = DBEngine.OpenDatabase(path)
    Chk "open"
    Show db
    Try db, "INSERT INTO C1 VALUES (6, 2)"
    Try db, "INSERT INTO C1 VALUES (7, 3)"
    Try db, "DELETE FROM P WHERE K = 2"
    Try db, "INSERT INTO C3 VALUES (3, 8)"
    db.Relations.Delete "R1"
    Chk "delete R1"
    Try db, "INSERT INTO C1 VALUES (8, 2)"
    Rows db, "C1"
    Rows db, "C3"
    db.Close
    NoCopy path
    Out "== after deleting R1"
    Set db = DBEngine.OpenDatabase(path)
    Show db
    db.CreateQueryDef "QU1", "SELECT K FROM P" & vbCrLf & "UNION SELECT K FROM C2"
    db.CreateQueryDef "QU2", "SELECT K FROM P UNION ALL SELECT K FROM C2 ORDER BY K"
    Chk "unions"
    db.Close
    NoCopy path
    Out "== unions read again"
    Set db = DBEngine.OpenDatabase(path)
    Out "QU1: " & Replace(db.QueryDefs("QU1").SQL, vbCrLf, "|")
    Out "QU2: " & Replace(db.QueryDefs("QU2").SQL, vbCrLf, "|")
    Chk "read unions"
    db.Close
    CloseOut
End Sub
