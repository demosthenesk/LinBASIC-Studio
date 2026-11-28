Attribute VB_Name = "ModJetLink"
Option Explicit
' Phase 14: linked tables (TableDef.Connect ";DATABASE=...", SourceTableName) kept in the .mdb: made, used, read
' again from the file (the copies LinBASIC keeps, file.mdb.lbdb, deleted first), changed, deleted.

Private Sub NoCopy(ByVal p As String)
    On Error Resume Next
    Kill p & ".lbdb"
    Err.Clear
End Sub

Private Function P(ByVal s As String) As String
    P = Replace(s, App.Path, "<dir>")
End Function

Private Sub Dump(ByVal path As String, ByVal title As String)
    Dim db As Database, td As TableDef, rs As Recordset, s As String
    On Error Resume Next
    NoCopy path
    Out "== " & title
    Set db = DBEngine.OpenDatabase(path)
    Chk "open"
    For Each td In db.TableDefs
        If (td.Attributes And dbSystemObject) = 0 Then
            Out "table " & td.Name & " attr=" & td.Attributes & " connect=" & P(td.Connect) & " source=" & td.SourceTableName & " fields=" & td.Fields.Count
            Set rs = db.OpenRecordset("SELECT * FROM [" & td.Name & "] ORDER BY 1")
            Chk "rows " & td.Name
            s = ""
            Do Until rs.EOF
                s = s & " " & V(rs.Fields(0).Value) & "/" & V(rs.Fields(1).Value)
                rs.MoveNext
            Loop
            rs.Close
            Out "  rows:" & s
        End If
    Next
    db.Close
End Sub

Sub Main()
    Dim src As String, mpath As String, db As Database, td As TableDef
    OpenOut "d_jetlink"
    On Error Resume Next
    src = App.Path & "\lsrc.mdb"
    mpath = App.Path & "\lmain.mdb"
    Kill src: Kill mpath
    NoCopy src: NoCopy mpath
    Err.Clear
    Set db = DBEngine.CreateDatabase(src, dbLangGeneral, dbVersion40)
    db.Execute "CREATE TABLE T (K LONG CONSTRAINT PK PRIMARY KEY, S TEXT(20))"
    db.Execute "INSERT INTO T VALUES (1, 'one')"
    db.Execute "INSERT INTO T VALUES (2, 'two')"
    db.Execute "CREATE TABLE U (A TEXT(5), B LONG)"
    db.Execute "INSERT INTO U VALUES ('x', 10)"
    db.Close
    Chk "src"
    Set db = DBEngine.CreateDatabase(mpath, dbLangGeneral, dbVersion40)
    db.Execute "CREATE TABLE Own (N LONG, M TEXT(10))"
    db.Execute "INSERT INTO Own VALUES (5, 'own')"
    Set td = db.CreateTableDef("LT")
    td.Connect = ";DATABASE=" & src
    td.SourceTableName = "T"
    db.TableDefs.Append td
    Chk "link"
    db.Execute "INSERT INTO LT VALUES (3, 'three')"
    Chk "insert through the link"
    db.Close
    Dump mpath, "made"
    Dump src, "source"
    Set db = DBEngine.OpenDatabase(mpath)
    Set td = db.CreateTableDef("LU")
    td.Connect = ";DATABASE=" & src
    td.SourceTableName = "U"
    db.TableDefs.Append td
    Chk "link 2"
    db.TableDefs.Delete "LT"
    Chk "delete link"
    db.Close
    Dump mpath, "changed"
    CloseOut
End Sub
