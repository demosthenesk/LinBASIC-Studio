Attribute VB_Name = "ModLog"
Option Explicit

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_ado.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

' the database of the test, made with DAO without a reference
Public Function MakeDb() As String
    Dim e As Object, db As Object, p As String
    p = App.Path & "\people.mdb"
    If Dir(p) <> "" Then Kill p
    Set e = CreateObject("DAO.DBEngine.36")
    Set db = e.CreateDatabase(p, ";LANGID=0x0409;CP=1252;COUNTRY=0", 64)
    db.Execute "CREATE TABLE People (ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(20) NOT NULL, Age SHORT, Active BIT, Notes MEMO)"
    db.Execute "INSERT INTO People (Name, Age, Active, Notes) VALUES ('Anna', 31, True, 'first')"
    db.Execute "INSERT INTO People (Name, Age, Active) VALUES ('Bill', 25, False)"
    db.Execute "INSERT INTO People (Name, Age, Active, Notes) VALUES ('Cleo', 42, True, 'third')"
    db.Execute "CREATE TABLE Towns (TownID LONG CONSTRAINT PrimaryKey PRIMARY KEY, Town TEXT(30))"
    db.Execute "INSERT INTO Towns VALUES (1, 'Athens')"
    db.Execute "INSERT INTO Towns VALUES (2, 'Patras')"
    db.Execute "INSERT INTO Towns VALUES (3, 'Volos')"
    db.Close
    MakeDb = p
End Function

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    MakeDb
    WriteLog "db made"
    Form1.Show
End Sub
