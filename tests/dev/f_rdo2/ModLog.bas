Attribute VB_Name = "ModLog"
Option Explicit

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_rdo2.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

Public Function DbPath() As String
    DbPath = App.Path & "\f_rdo2.db"
End Function

Public Function CS() As String
    CS = "Driver={SQLite3 ODBC Driver};Database=" & DbPath
End Function

' the database of the test, made again each time with RDO
Sub Main()
    Dim cn As rdoConnection
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    If Dir(DbPath) <> "" Then Kill DbPath
    Set cn = rdoEngine.rdoEnvironments(0).OpenConnection("", rdDriverNoPrompt, False, CS)
    cn.Execute "CREATE TABLE People (ID INTEGER PRIMARY KEY, Name VARCHAR(20) NOT NULL, Age SMALLINT, Active BIT)"
    cn.Execute "INSERT INTO People VALUES (1, 'Anna', 31, 1)"
    cn.Execute "INSERT INTO People VALUES (2, 'Bill', 25, 0)"
    cn.Execute "INSERT INTO People VALUES (3, 'Cleo', 42, 1)"
    cn.Close
    Form1.Show
End Sub

' the rows as the database has them now (another connection)
Public Function DbRows() As String
    Dim cn As rdoConnection, rs As rdoResultset, s As String
    Set cn = rdoEngine.rdoEnvironments(0).OpenConnection("", rdDriverNoPrompt, True, CS)
    Set rs = cn.OpenResultset("SELECT ID, Name, Age, Active FROM People ORDER BY ID", rdOpenStatic, rdConcurReadOnly)
    Do Until rs.EOF
        s = s & " " & rs!ID & ":" & rs!Name & ":" & rs!Age & ":" & rs!Active
        rs.MoveNext
    Loop
    rs.Close
    cn.Close
    DbRows = s
End Function
