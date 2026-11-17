Attribute VB_Name = "ModLog"
Option Explicit

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_dfmt4.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

' a value with its type: Long:31, String:"x", Null, Empty, Object:Nothing
Public Function V(x As Variant) As String
    If IsObject(x) Then
        If x Is Nothing Then V = "Nothing" Else V = "Object:" & TypeName(x)
    ElseIf IsNull(x) Then
        V = "Null"
    ElseIf IsEmpty(x) Then
        V = "Empty"
    ElseIf VarType(x) = vbString Then
        V = "String:""" & x & """"
    Else
        V = TypeName(x) & ":" & x
    End If
End Function

' the database of the test, made with DAO without a reference; Jet 3 (dbVersion30) so that the Data
' control of VB6 (DAO 3.51) opens it too
Private Sub MakeDb()
    Dim e As Object, db As Object, p As String
    p = App.Path & "\dfmt4.mdb"
    If Dir(p) <> "" Then Kill p
    Set e = CreateObject("DAO.DBEngine.36")
    Set db = e.CreateDatabase(p, ";LANGID=0x0409;CP=1252;COUNTRY=0", 32)
    db.Execute "CREATE TABLE People (ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(20) NOT NULL, Age SHORT, Active BIT, Born DATETIME, Salary CURRENCY, Notes MEMO)"
    db.Execute "INSERT INTO People (Name, Age, Active, Born, Salary, Notes) VALUES ('Anna', 31, True, #1994-03-05#, 1234.5, 'first')"
    db.Execute "INSERT INTO People (Name, Age, Active) VALUES ('Bill', 25, False)"
    db.Execute "INSERT INTO People (Name, Active, Born, Salary, Notes) VALUES ('Cleo', True, #2000-12-31 13:45:00#, 99.99, 'third')"
    db.Close
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    On Error Resume Next
    ChDrive App.Path
    Err.Clear
    ChDir App.Path
    If Err.Number <> 0 Then WriteLog "chdir ! " & Err.Number & " " & Err.Description
    On Error GoTo 0
    MakeDb
    If LCase$(CurDir$) <> LCase$(App.Path) Then
        ' the Adodc of the .frm opens "dfmt4.mdb" in the current folder: App.Path is a network path (VM), the
        ' current folder stays the folder of VB6: the database goes there too
        On Error Resume Next
        FileCopy App.Path & "\dfmt4.mdb", CurDir$ & "\dfmt4.mdb"
        If Err.Number <> 0 Then WriteLog "copy ! " & Err.Number & " " & Err.Description
        On Error GoTo 0
    End If
    WriteLog "db made"
    Form1.Show
End Sub
