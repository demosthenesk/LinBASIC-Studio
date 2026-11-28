Attribute VB_Name = "ModJetRule"
Option Explicit
' Phase 14: when Jet checks validation rules -- rows already in the file that break a rule (a rule set after
' them), a change of another field of such a row (SQL and Recordset.Edit), a rule with Now(), Null values;
' in the session and read again from the file (the copy LinBASIC keeps, file.mdb.lbdb, deleted first).

Private Sub NoCopy(ByVal p As String)
    On Error Resume Next
    Kill p & ".lbdb"
    Err.Clear
End Sub

Private Sub Try(db As Database, ByVal sql As String)
    On Error Resume Next
    db.Execute sql, dbFailOnError
    Out "  " & sql & " -> " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

Private Function Z(ByVal s As String) As String
    Z = "'" & Replace(s, Chr$(0), "\0") & "'"
End Function

Private Sub Props(db As Database)
    Dim td As TableDef
    Set td = db.TableDefs("R")
    Out "props: S zl=" & td.Fields("S").AllowZeroLength & " A vr=" & Z(td.Fields("A").ValidationRule) & " A vt=" & Z(td.Fields("A").ValidationText) & _
        " D vr=" & Z(td.Fields("D").ValidationRule) & " B def=" & Z(td.Fields("B").DefaultValue) & " R vr=" & Z(td.ValidationRule) & " R vt=" & Z(td.ValidationText)
End Sub

Private Sub Rows(db As Database)
    Dim rs As Recordset, s As String
    On Error Resume Next
    Set rs = db.OpenRecordset("SELECT * FROM R ORDER BY ID")
    Do Until rs.EOF
        s = s & " " & V(rs!ID.Value) & "/" & V(rs!A.Value) & "/" & V(rs!B.Value) & "/" & IIf(IsNull(rs!D.Value), "Null", IIf(rs!D.Value > Now, "future", "past")) & "/" & V(rs!S.Value)
        rs.MoveNext
    Loop
    rs.Close
    Out "  R:" & s
    Chk "rows"
End Sub

Private Sub EditB(db As Database, ByVal id As Long, ByVal b As Long)
    Dim rs As Recordset
    On Error Resume Next
    Set rs = db.OpenRecordset("SELECT * FROM R WHERE ID = " & id, dbOpenDynaset)
    rs.Edit
    rs!B = b
    rs.Update
    Out "  Edit B of " & id & " = " & b & " -> " & Err.Number & " " & Err.Description
    Err.Clear
    rs.CancelUpdate
    Err.Clear
    rs.Close
End Sub

Private Sub Session(db As Database)
    Try db, "UPDATE R SET B = 11 WHERE ID = 1"
    EditB db, 1, 13
    EditB db, 2, 12
    Try db, "UPDATE R SET A = A WHERE ID = 1"
    Try db, "UPDATE R SET S = S WHERE ID = 1"
    Try db, "UPDATE R SET S = '' WHERE ID = 2"
    Try db, "INSERT INTO R (ID, A, B, D) VALUES (10, -1, 1, Now() + 10)"
    Try db, "INSERT INTO R (ID, A, B, D) VALUES (11, 1, 1, Now() - 10)"
    Try db, "INSERT INTO R (ID, A, B, D) VALUES (12, 1, 1, Now() + 10)"
    Try db, "INSERT INTO R (ID, A, B, D) VALUES (13, Null, Null, Null)"
    Try db, "INSERT INTO R (ID, A, B, D) VALUES (14, 5, 150, Null)"
    Try db, "UPDATE R SET A = -7 WHERE ID = 12"
    Rows db
End Sub

Sub Main()
    Dim path As String, db As Database, td As TableDef
    OpenOut "d_jetrule"
    On Error Resume Next
    path = App.Path & "\rule.mdb"
    Kill path
    NoCopy path
    Err.Clear
    Set db = DBEngine.CreateDatabase(path, dbLangGeneral, dbVersion40)
    db.Execute "CREATE TABLE R (ID LONG CONSTRAINT PK PRIMARY KEY, A LONG, B LONG, D DATETIME, S TEXT(10))"
    db.Execute "INSERT INTO R VALUES (1, -5, 1, #1/1/1997#, '')"
    db.Execute "INSERT INTO R VALUES (2, -6, 200, #1/1/1998#, 'x')"
    Chk "rows"
    Set td = db.TableDefs("R")
    td.Fields("A").ValidationRule = ">0"
    Chk "rule A"
    td.Fields("A").ValidationText = "A must be positive"
    td.Fields("S").AllowZeroLength = False
    Chk "zero length S"
    td.Fields("D").ValidationRule = ">Now()"
    Chk "rule D"
    td.ValidationRule = "[B]<100"
    Chk "rule table"
    td.ValidationText = "B too big"
    Chk "text table"
    td.Fields("B").DefaultValue = "5"
    Chk "default B"
    Props db
    Out "== session"
    Rows db
    Session db
    db.Close
    NoCopy path
    Out "== read again"
    Set db = DBEngine.OpenDatabase(path)
    Chk "open"
    Props db
    Rows db
    Try db, "DELETE FROM R WHERE ID >= 10"
    Session db
    db.Close
    CloseOut
End Sub
