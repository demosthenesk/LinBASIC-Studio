Attribute VB_Name = "ModLog"
Option Explicit

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_dbgunb.txt"
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
    p = App.Path & "\dbgunb.mdb"
    If Dir(p) <> "" Then Kill p
    Set e = CreateObject("DAO.DBEngine.36")
    Set db = e.CreateDatabase(p, ";LANGID=0x0409;CP=1252;COUNTRY=0", 32)   ' dbVersion30: the Data control of VB6 is DAO 3.51
    db.Execute "CREATE TABLE Groups (GID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, GName TEXT(20) NOT NULL)"
    db.Execute "INSERT INTO Groups (GName) VALUES ('Alpha')"
    db.Execute "INSERT INTO Groups (GName) VALUES ('Beta')"
    db.Execute "INSERT INTO Groups (GName) VALUES ('Gamma')"
    db.Execute "CREATE TABLE People (ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(20) NOT NULL, Age SHORT, GroupID LONG)"
    db.Execute "INSERT INTO People (Name, Age, GroupID) VALUES ('Anna', 31, 1)"
    db.Execute "INSERT INTO People (Name, Age, GroupID) VALUES ('Bill', 25, 2)"
    db.Execute "INSERT INTO People (Name, Age, GroupID) VALUES ('Cleo', 42, 3)"
    db.Execute "INSERT INTO People (Name, Age, GroupID) VALUES ('Dora', 19, 2)"
    db.Close
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    MakeDb
    WriteLog "db made"
    Form1.Show
End Sub
