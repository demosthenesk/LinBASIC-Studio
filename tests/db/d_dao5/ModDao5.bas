Attribute VB_Name = "ModDao5"
Option Explicit
' External databases in Jet SQL and linked tables (phase 12.4: the Import/Export and the Attachments of the Visual
' Data Manager): IN 'file' and [;database=file].table in SELECT, SELECT INTO, INSERT INTO, UPDATE, DELETE and a
' join; a TableDef linked to a table of another database (Connect, SourceTableName), its fields, indexes,
' attributes, a recordset over it, changes through it, RefreshLink; the errors.

Private gMain As String, gOther As String

Private Function P(ByVal s As String) As String
    P = Replace(Replace(Replace(s, gOther, "<other>"), gMain, "<main>"), App.Path, "<dir>")
End Function

Private Sub Rows(db As Database, ByVal sql As String)
    Dim rs As Recordset
    On Error Resume Next
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Err.Number <> 0 Then
        Out "  ! open " & P(sql) & ": " & Err.Number & " " & P(Err.Description) & " [" & Err.Source & "]"
        Err.Clear
        Exit Sub
    End If
    DumpRs rs, P(sql)
    rs.Close
End Sub

Private Sub Exe(db As Database, ByVal sql As String)
    On Error Resume Next
    db.Execute sql, dbFailOnError
    If Err.Number <> 0 Then
        Out "  ! exec " & P(sql) & ": " & Err.Number & " " & P(Err.Description) & " [" & Err.Source & "]"
        Err.Clear
    Else
        Out "exec " & P(sql) & " -> " & db.RecordsAffected
    End If
End Sub

Sub Main()
    Dim db As Database, other As Database, td As TableDef, f As Field, ix As Index, rs As Recordset, s As String
    OpenOut "d_dao5"
    On Error Resume Next
    gMain = App.Path & "\dao5.mdb"
    gOther = App.Path & "\dao5other.mdb"
    If Dir(gMain) <> "" Then Kill gMain
    If Dir(gOther) <> "" Then Kill gOther
    Set db = DBEngine.CreateDatabase(gMain, dbLangGeneral, dbVersion40)
    Set other = DBEngine.CreateDatabase(gOther, dbLangGeneral, dbVersion40)
    Chk "create"
    db.Execute "CREATE TABLE People (ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(30), Age INTEGER, Born DATETIME)"
    db.Execute "INSERT INTO People (Name, Age, Born) VALUES ('Ann', 31, #1993-02-03#)"
    db.Execute "INSERT INTO People (Name, Age, Born) VALUES ('Bob', 45, #1979-11-20#)"
    db.Execute "INSERT INTO People (Name, Age, Born) VALUES ('Cid', 12, #2012-06-01#)"
    other.Execute "CREATE TABLE Pets (PetID COUNTER CONSTRAINT PetKey PRIMARY KEY, Owner TEXT(30), Kind TEXT(10), Legs BYTE)"
    other.Execute "CREATE INDEX idxOwner ON Pets (Owner)"
    other.Execute "INSERT INTO Pets (Owner, Kind, Legs) VALUES ('Ann', 'Cat', 4)"
    other.Execute "INSERT INTO Pets (Owner, Kind, Legs) VALUES ('Bob', 'Parrot', 2)"
    other.Execute "INSERT INTO Pets (Owner, Kind, Legs) VALUES ('Ann', 'Dog', 4)"
    other.Close
    Chk "tables"

    Out "== IN 'file'"
    Rows db, "SELECT * FROM Pets IN '" & gOther & "' ORDER BY PetID"
    Rows db, "SELECT Kind, Legs FROM Pets IN """ & gOther & """ WHERE Legs = 4 ORDER BY Kind"
    Rows db, "SELECT Owner, Count(*) AS N FROM Pets IN '" & gOther & "' GROUP BY Owner"
    Rows db, "SELECT * FROM Pets IN '' [;database=" & gOther & "] ORDER BY PetID DESC"
    Out "== [;database=file].table"
    Rows db, "SELECT * FROM [;database=" & gOther & "].Pets WHERE Owner = 'Ann'"
    Rows db, "SELECT p.Name, x.Kind FROM People AS p INNER JOIN [;database=" & gOther & "].[Pets] AS x ON p.Name = x.Owner ORDER BY x.PetID"
    Rows db, "SELECT People.Name, Pets.Kind FROM People, Pets IN '" & gOther & "' WHERE People.Name = Pets.Owner ORDER BY Pets.Kind"
    Out "== SELECT INTO"
    Exe db, "SELECT * INTO Copy1 IN '" & gOther & "' FROM People"
    Exe db, "SELECT Name, Age INTO [;database=" & gOther & "].Copy2 FROM People WHERE Age > 20"
    Exe db, "SELECT * INTO Copy1 IN '" & gOther & "' FROM People"
    Exe db, "SELECT * INTO PetsHere FROM Pets IN '" & gOther & "' WHERE Kind <> 'Dog'"
    Rows db, "SELECT * FROM PetsHere"
    Set other = DBEngine.OpenDatabase(gOther)
    For Each td In other.TableDefs
        If (td.Attributes And dbSystemObject) = 0 Then
            s = td.Name & ":"
            For Each f In td.Fields
                s = s & " " & f.Name & "/" & f.Type & "/" & f.Size
            Next
            Out s
        End If
    Next
    other.Close
    Out "== INSERT INTO, UPDATE, DELETE"
    Exe db, "INSERT INTO Copy1 IN '" & gOther & "' (Name, Age) VALUES ('Dee', 50)"
    Exe db, "INSERT INTO [;database=" & gOther & "].Copy2 SELECT Name, Age FROM People WHERE Age < 20"
    Exe db, "UPDATE [;database=" & gOther & "].Copy2 SET Age = Age + 1 WHERE Name = 'Cid'"
    Exe db, "UPDATE Copy1 IN '" & gOther & "' SET Age = 0 WHERE Name = 'Dee'"
    Exe db, "DELETE FROM Copy1 IN '" & gOther & "' WHERE Name = 'Bob'"
    Exe db, "DELETE * FROM [;database=" & gOther & "].Copy2 WHERE Age > 40"
    Rows db, "SELECT * FROM Copy1 IN '" & gOther & "' ORDER BY Name"
    Rows db, "SELECT * FROM [;database=" & gOther & "].Copy2 ORDER BY Name"
    Out "== errors"
    Rows db, "SELECT * FROM Pets IN '" & App.Path & "\nosuch.mdb'"
    Rows db, "SELECT * FROM NoTable IN '" & gOther & "'"
    Exe db, "INSERT INTO NoTable IN '" & gOther & "' (A) VALUES (1)"
    Rows db, "SELECT * FROM [;database=" & App.Path & "\nosuch.mdb].Pets"

    Out "== linked table"
    Set td = db.CreateTableDef("LinkPets")
    td.Connect = ";DATABASE=" & gOther
    td.SourceTableName = "Pets"
    db.TableDefs.Append td
    Chk "append link"
    Set td = db.TableDefs("LinkPets")
    Out "attrs=" & Hex(td.Attributes) & " connect=" & P(td.Connect) & " source=" & td.SourceTableName & " updatable=" & td.Updatable
    s = "fields:"
    For Each f In td.Fields
        s = s & " " & f.Name & "/" & f.Type & "/" & f.Size & "/" & f.SourceTable & "/" & f.SourceField
    Next
    Out s
    s = "indexes:"
    For Each ix In td.Indexes
        s = s & " " & ix.Name & "(" & ix.Fields & ")" & ix.Primary
    Next
    Out s
    Chk "link info"
    Out "count=" & td.RecordCount
    Chk "link count"
    Rows db, "SELECT * FROM LinkPets ORDER BY PetID"
    Rows db, "SELECT p.Name, l.Kind FROM People p INNER JOIN LinkPets l ON p.Name = l.Owner ORDER BY l.Kind"
    Set rs = db.OpenRecordset("LinkPets")
    Out "default type=" & rs.Type
    Chk "link open"
    rs.Close
    Set rs = db.OpenRecordset("LinkPets", dbOpenTable)
    ChkAlways "link open table"
    Set rs = db.OpenRecordset("LinkPets", dbOpenDynaset)
    rs.AddNew
    rs!Owner = "Cid"
    rs!Kind = "Fish"
    rs!Legs = 0
    rs.Update
    Chk "link addnew"
    rs.MoveFirst
    rs.Edit
    rs!Legs = 3
    rs.Update
    Chk "link edit"
    rs.Close
    Exe db, "UPDATE LinkPets SET Kind = UCase(Kind) WHERE Legs = 4"
    Exe db, "DELETE FROM LinkPets WHERE Kind = 'PARROT'"
    Rows db, "SELECT * FROM Pets IN '" & gOther & "' ORDER BY PetID"
    td.RefreshLink
    Chk "refreshlink"
    Out "== a link that breaks"
    Set td = db.CreateTableDef("LinkBad")
    td.Connect = ";DATABASE=" & App.Path & "\nosuch.mdb"
    td.SourceTableName = "Pets"
    db.TableDefs.Append td
    Out "  ! append bad link: " & Err.Number & " " & P(Err.Description) & " [" & Err.Source & "]": Err.Clear
    Set td = db.CreateTableDef("LinkBad2")
    td.Connect = ";DATABASE=" & gOther
    td.SourceTableName = "NoTable"
    db.TableDefs.Append td
    ChkAlways "append bad table"
    db.TableDefs.Delete "LinkPets"
    Chk "delete link"
    Rows db, "SELECT Count(*) AS N FROM Pets IN '" & gOther & "'"
    db.Close
    Set db = DBEngine.OpenDatabase(gMain)
    Out "reopened tables:"
    For Each td In db.TableDefs
        If (td.Attributes And dbSystemObject) = 0 Then Out "  " & td.Name & " " & Hex(td.Attributes)
    Next
    db.Close
    Kill gMain
    Kill gOther
    CloseOut
End Sub
