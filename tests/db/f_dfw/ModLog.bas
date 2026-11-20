Attribute VB_Name = "ModLog"
Option Explicit
' f_dfw: what the forms of a data form wizard need from the runtime (LinBASIC Studio, phase 9.2)

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_dfw.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

Public Function DbPath() As String
    DbPath = App.Path & "\dfw.mdb"
End Function

' the database of the test, made with DAO without a reference
Public Sub MakeDb()
    Dim e As Object, db As Object
    If Dir(DbPath) <> "" Then Kill DbPath
    Set e = CreateObject("DAO.DBEngine.36")
    Set db = e.CreateDatabase(DbPath, ";LANGID=0x0409;CP=1252;COUNTRY=0", 64)
    db.Execute "CREATE TABLE Writers (WID LONG CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(30), Born SHORT)"
    db.Execute "INSERT INTO Writers VALUES (1, 'Ames, Ann', 1950)"
    db.Execute "INSERT INTO Writers VALUES (2, 'Boyd, Ben', 1961)"
    db.Execute "INSERT INTO Writers (WID, Name) VALUES (3, 'Cole, Cy')"
    db.Execute "CREATE TABLE Pubs (PID LONG CONSTRAINT PrimaryKey PRIMARY KEY, PName TEXT(30))"
    db.Execute "INSERT INTO Pubs VALUES (1, 'Alpha')"
    db.Execute "INSERT INTO Pubs VALUES (2, 'Beta')"
    db.Execute "INSERT INTO Pubs VALUES (3, 'Gamma')"
    db.Execute "INSERT INTO Pubs VALUES (4, 'Delta')"
    db.Execute "CREATE TABLE Books (BID LONG CONSTRAINT PrimaryKey PRIMARY KEY, Title TEXT(30), PID LONG)"
    db.Execute "INSERT INTO Books VALUES (10, 'Rivers', 2)"
    db.Execute "INSERT INTO Books VALUES (11, 'Stones', 2)"
    db.Execute "INSERT INTO Books VALUES (12, 'Clouds', 3)"
    db.Execute "INSERT INTO Books VALUES (13, 'Roads', 4)"
    db.Execute "INSERT INTO Books VALUES (14, 'Hills', 4)"
    db.Execute "INSERT INTO Books VALUES (15, 'Lakes', 4)"
    db.Close
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    MakeDb
    WriteLog "db made"
    Form1.Show
End Sub
