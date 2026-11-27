Attribute VB_Name = "ModDao7"
Option Explicit
' DBEngine.RegisterDatabase (phase 12.4: Register of the ODBC Logon of the Visual Data Manager): a data source of the
' user written with its driver and its attributes, then opened by its name alone; registered again with another
' database; a driver that is not there; no name. The SQLite ODBC driver of Christian Werner.

Private gDb As String, gDb2 As String

Private Function P(ByVal s As String) As String
    P = Replace(Replace(Replace(s, gDb2, "<db2>"), gDb, "<db>"), App.Path, "<dir>")
End Function

Private Sub Errs(ByVal what As String)
    Dim e As Error, i As Integer
    If Err.Number = 0 Then Exit Sub
    Out "  ! " & what & ": " & Err.Number & " " & P(Err.Description) & " [" & Err.Source & "]"
    For i = 0 To DBEngine.Errors.Count - 1
        Set e = DBEngine.Errors(i)
        Out "    err " & i & ": " & e.Number & " " & P(e.Description) & " [" & e.Source & "]"
    Next
    Err.Clear
End Sub

Private Sub Use(ByVal table As String)
    Dim db As Database, rs As Recordset
    Set db = DBEngine.Workspaces(0).OpenDatabase("", dbDriverNoPrompt, False, "ODBC;DSN=LBDAO7;")
    If Err.Number <> 0 Then Errs "open": Exit Sub
    Out "open: database=" & Mid$(db.Name, InStrRev(Replace(db.Name, "/", "\"), "\") + 1)
    db.Execute "CREATE TABLE " & table & " (ID INTEGER PRIMARY KEY, Name VARCHAR(20))", dbSQLPassThrough
    Errs "create"
    db.Execute "INSERT INTO " & table & " (ID, Name) VALUES (7, 'seven')", dbSQLPassThrough
    Errs "insert"
    Set rs = db.OpenRecordset("SELECT ID, Name FROM " & table, dbOpenSnapshot, dbSQLPassThrough)
    If Err.Number <> 0 Then
        Errs "select"
    Else
        Do While Not rs.EOF
            Out "row " & rs!ID & " " & rs!Name
            rs.MoveNext
        Loop
        rs.Close
    End If
    db.Close
End Sub

Sub Main()
    OpenOut "d_dao7"
    On Error Resume Next
    gDb = App.Path & "\dao7.db"
    gDb2 = App.Path & "\dao7b.db"
    If Dir(gDb) <> "" Then Kill gDb
    If Dir(gDb2) <> "" Then Kill gDb2

    Out "== register"
    DBEngine.RegisterDatabase "LBDAO7", "SQLite3 ODBC Driver", True, "Database=" & gDb & vbCr & "Description=d_dao7"
    Errs "register"
    Use "T1"
    Out "file=" & (Dir(gDb) <> "") & " file2=" & (Dir(gDb2) <> "")

    Out "== again, another database"
    DBEngine.RegisterDatabase "LBDAO7", "SQLite3 ODBC Driver", True, "Database=" & gDb2
    Errs "register again"
    Use "T2"
    Out "file2=" & (Dir(gDb2) <> "")

    Out "== a driver that is not there"
    DBEngine.RegisterDatabase "LBDAO7X", "No Such Driver 7", True, ""
    Errs "register unknown"

    Out "== no name"
    DBEngine.RegisterDatabase "", "SQLite3 ODBC Driver", True, ""
    Errs "register empty"
    Out "== end"
    CloseOut
End Sub
