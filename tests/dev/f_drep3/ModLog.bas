Attribute VB_Name = "ModLog"
Option Explicit

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_drep3.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

Public Function DbPath() As String
    DbPath = App.Path & "\f_drep3.db"
End Function

Public Function CS() As String
    CS = "Driver={SQLite3 ODBC Driver};Database=" & DbPath
End Function

' the database of the test (SQLite ODBC, as the RDO tests), then the form
Sub Main()
    Dim cn As ADODB.Connection
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    If Dir(DbPath) <> "" Then Kill DbPath
    WriteLog "== start"
    Set cn = New ADODB.Connection
    cn.Open CS
    cn.Execute "CREATE TABLE People (ID INTEGER PRIMARY KEY, Name VARCHAR(20), Age SMALLINT, Active BIT, Notes VARCHAR(40))"
    cn.Execute "INSERT INTO People VALUES (1, 'Anna', 31, 1, 'first')"
    cn.Execute "INSERT INTO People VALUES (2, 'Bill', 25, 0, 'second')"
    cn.Execute "INSERT INTO People VALUES (3, 'Cleo', 42, 1, NULL)"
    cn.Execute "INSERT INTO People VALUES (4, 'Dino', 37, 0, 'fourth')"
    cn.Execute "INSERT INTO People VALUES (5, 'Eve', 28, 1, 'fifth')"
    cn.Execute "INSERT INTO People VALUES (6, 'Finn', 50, 0, 'sixth')"
    cn.Execute "INSERT INTO People VALUES (7, 'Gus', 19, 1, 'seventh')"
    cn.Execute "INSERT INTO People VALUES (8, 'Hana', 64, 0, 'eighth')"
    cn.Close
    Form1.Show
End Sub

' a bookmark as the Name of its record (a clone of rs)
Public Function RowName(rs As ADODB.Recordset, bm As Variant) As String
    Dim c As ADODB.Recordset
    On Error GoTo bad
    If IsNull(bm) Or IsEmpty(bm) Then RowName = "(" & TypeName(bm) & ")": Exit Function
    Set c = rs.Clone
    c.Bookmark = bm
    RowName = c!Name
    Exit Function
bad:
    RowName = "? " & Err.Number
End Function

Public Function Arr(v As Variant) As String
    Dim i As Long, s As String
    On Error GoTo bad
    If Not IsArray(v) Then Arr = TypeName(v): Exit Function
    s = TypeName(v) & "(" & LBound(v) & " To " & UBound(v) & ")"
    For i = LBound(v) To UBound(v)
        s = s & " " & v(i)
    Next
    Arr = s
    Exit Function
bad:
    Arr = s & " ! " & Err.Number
End Function

' the rows of the database now (another connection)
Public Function DbRows() As String
    Dim cn As ADODB.Connection, r As ADODB.Recordset, s As String
    Set cn = New ADODB.Connection
    cn.Open CS
    Set r = cn.Execute("SELECT ID, Name, Age, Active FROM People ORDER BY ID")
    Do Until r.EOF
        s = s & " " & r!ID & ":" & r!Name & ":" & r!Age & ":" & r!Active
        r.MoveNext
    Loop
    r.Close
    cn.Close
    DbRows = s
End Function
