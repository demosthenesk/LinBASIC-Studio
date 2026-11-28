Attribute VB_Name = "ModJetPw"
Option Explicit
' Phase 14: the password of a Jet database (NewPassword, ";pwd=" in OpenDatabase), Jet 4 and Jet 3; the file is read
' again (the copy LinBASIC keeps next to it, file.mdb.lbdb, deleted first).

Private Sub NoCopy(ByVal p As String)
    On Error Resume Next
    Kill p & ".lbdb"
    Err.Clear
End Sub

Private Sub TryOpen(ByVal p As String, ByVal con As String)
    Dim db As Database, rs As Recordset
    On Error Resume Next
    NoCopy p
    If con = "-" Then Set db = DBEngine.OpenDatabase(p) Else Set db = DBEngine.OpenDatabase(p, False, False, con)
    ChkAlways "open " & con
    If Not db Is Nothing Then
        Set rs = db.OpenRecordset("SELECT Count(*) AS N FROM T")
        Out "  rows " & V(rs!N)
        rs.Close
        db.Close
    End If
End Sub

Private Sub One(ByVal name As String, ByVal ver As Long)
    Dim p As String, db As Database
    On Error Resume Next
    p = App.Path & "\" & name & ".mdb"
    If Dir(p) <> "" Then Kill p
    NoCopy p
    Out "== " & name
    Set db = DBEngine.CreateDatabase(p, dbLangGeneral, ver)
    db.Execute "CREATE TABLE T (K LONG, S TEXT(10))"
    db.Execute "INSERT INTO T VALUES (1, 'one')"
    db.NewPassword "", "Secret1"
    ChkAlways "new password"
    db.Close
    TryOpen p, "-"
    TryOpen p, ";pwd=wrong"
    TryOpen p, ";pwd=secret1"
    TryOpen p, ";pwd=Secret1"
    TryOpen p, ";PWD=Secret1;"
    NoCopy p
    Set db = DBEngine.OpenDatabase(p, False, False, ";pwd=Secret1")
    db.NewPassword "Secret1", "x"
    ChkAlways "new password, shared"
    db.Close
    NoCopy p
    Set db = DBEngine.OpenDatabase(p, True, False, ";pwd=Secret1")
    ChkAlways "open exclusive"
    db.NewPassword "wrong", "x"
    ChkAlways "new password, wrong old one"
    db.NewPassword "Secret1", "123456789012345678901"
    ChkAlways "new password, 21 characters"
    db.NewPassword "Secret1", "Ab_12345678901234567"
    ChkAlways "new password, 20 characters"
    db.Execute "INSERT INTO T VALUES (2, 'two')"
    db.Close
    TryOpen p, ";pwd=Ab_12345678901234567"
    TryOpen p, ";pwd=Ab_1234567890123456"
    NoCopy p
    Set db = DBEngine.OpenDatabase(p, True, False, ";pwd=Ab_12345678901234567")
    db.NewPassword "Ab_12345678901234567", ""
    ChkAlways "no password"
    db.Close
    TryOpen p, "-"
    If Dir(App.Path & "\" & name & "c.mdb") <> "" Then Kill App.Path & "\" & name & "c.mdb"
    DBEngine.CompactDatabase p, App.Path & "\" & name & "c.mdb"
    ChkAlways "compact"
End Sub

Sub Main()
    OpenOut "d_jetpw"
    One "pw4", dbVersion40
    One "pw3", dbVersion30
    CloseOut
End Sub
