Attribute VB_Name = "ModLog"
Option Explicit

Private Function LogPath() As String
    LogPath = App.Path & "\expected\f_data.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

' the database of the test, made with DAO without a reference (late bound, as the Data control uses it)
Private Sub MakeDb()
    Dim e As Object, db As Object, p As String
    p = App.Path & "\people.mdb"
    If Dir(p) <> "" Then Kill p
    Set e = CreateObject("DAO.DBEngine.36")
    Set db = e.CreateDatabase(p, ";LANGID=0x0409;CP=1252;COUNTRY=0", 32)   ' dbVersion30: the Data control of VB6 is DAO 3.51
    db.Execute "CREATE TABLE People (ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(20) NOT NULL, Age SHORT, Active BIT, Notes MEMO)"
    db.Execute "INSERT INTO People (Name, Age, Active, Notes) VALUES ('Anna', 31, True, 'first')"
    db.Execute "INSERT INTO People (Name, Age, Active) VALUES ('Bill', 25, False)"
    db.Execute "INSERT INTO People (Name, Age, Active, Notes) VALUES ('Cleo', 42, True, 'third')"
    db.Close
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    MakeDb
    WriteLog "db made"
    Form1.Show
End Sub
