Attribute VB_Name = "ModLog"
Option Explicit

Public Function LogPath() As String
    LogPath = App.Path & "\expected\d_bind.txt"
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
    Run
End Sub


Private Sub Run()
    Dim cn As New ADODB.Connection, rs As New ADODB.Recordset, bc As New BindingCollection
    Dim p As New CPerson, src As New CSource, b As Binding, bc2 As New BindingCollection, p2 As New CPerson
    On Error Resume Next
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\people.mdb"
    rs.Open "SELECT * FROM People ORDER BY ID", cn, adOpenKeyset, adLockOptimistic
    WriteLog "bc " & TypeName(bc) & " count=" & bc.Count & " mode=" & bc.UpdateMode & " member=[" & bc.DataMember & "]"
    Set bc.DataSource = rs
    WriteLog "datasource " & TypeName(bc.DataSource)
    Set b = bc.Add(p, "Name", "Name")
    WriteLog "binding " & TypeName(b) & " key=[" & b.Key & "] field=" & b.DataField & " prop=" & b.PropertyName & " changed=" & b.DataChanged & " obj=" & TypeName(b.Object)
    bc.Add p, "Age", "Age", , "age"
    WriteLog "count " & bc.Count & " item age=" & bc("age").DataField & " item 0=" & bc(0).PropertyName
    WriteLog "p " & p.Name & "/" & p.Age
    rs.MoveNext
    WriteLog "moved p " & p.Name & "/" & p.Age
    p.Name = "Billy"
    WriteLog "changed binding=" & bc(0).DataChanged & " field=" & rs!Name
    rs.MoveNext
    WriteLog "moved p " & p.Name & "/" & p.Age
    rs.MovePrevious
    WriteLog "back p " & p.Name & "/" & p.Age & " field=" & rs!Name
    p.Age = 30
    bc.UpdateMode = vbUpdateWhenPropertyChanges
    p.Name = "Bob"
    WriteLog "mode 1 field=" & rs!Name & " age=" & rs!Age & " editmode=" & rs.EditMode
    rs.MoveFirst
    bc.UpdateControls
    WriteLog "updatecontrols p " & p.Name & "/" & p.Age
    bc.Remove "age"
    WriteLog "removed count " & bc.Count
    bc.Remove 5
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    bc.Add p, "Nope", "Name"
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    bc.Add p, "Age", "NoField"
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "== class source"
    Set bc2.DataSource = src
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    bc2.Add p2, "Name", "Text"
    WriteLog "p2 " & p2.Name
    src.MoveNext
    WriteLog "src moved p2 " & p2.Name
    bc2.DataMember = "other"
    WriteLog "member other p2 " & p2.Name
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "end"
End Sub
