Attribute VB_Name = "ModDao6"
Option Explicit
' DAO over ODBC (phase 12.4: Open DataBase > ODBC... of the Visual Data Manager): a database of an ODBC data source
' opened by Jet (OpenDatabase with "ODBC;..."), its TableDefs, Fields and Indexes, recordsets (dynaset, snapshot,
' pass-through), changes of rows, Execute (Jet SQL and pass-through), a pass-through QueryDef, transactions, the
' errors (DBEngine.Errors), a table of the data source linked to an Access database. The SQLite ODBC driver of
' Christian Werner (the same version as on Linux).

Private gDb As String, gMdb As String, gCon As String

Private Function P(ByVal s As String) As String
    P = Replace(Replace(Replace(s, gDb, "<db>"), gMdb, "<mdb>"), App.Path, "<dir>")
End Function

Private Sub Errs(ByVal what As String)
    Dim e As Error, i As Integer
    If Err.Number = 0 Then Exit Sub
    Out "  ! " & what & ": " & Err.Number & " " & P(Err.Description) & " [" & Err.Source & "]"
    For i = 0 To DBEngine.Errors.Count - 1
        Set e = DBEngine.Errors(i)
        Out "    err " & i & ": " & e.Number & " " & P(e.Description) & " [" & e.Source & "]"
    Next
    Err.Clear
End Sub

Private Sub Rows(db As Database, ByVal sql As String, Optional ByVal typ As Integer = dbOpenSnapshot, Optional ByVal opt As Long = 0)
    Dim rs As Recordset
    On Error Resume Next
    If opt Then Set rs = db.OpenRecordset(sql, typ, opt) Else Set rs = db.OpenRecordset(sql, typ)
    If Err.Number <> 0 Then Errs "open " & P(sql): Exit Sub
    DumpRs rs, P(sql) & " type=" & rs.Type & " upd=" & rs.Updatable
    rs.Close
End Sub

Private Sub Exe(db As Database, ByVal sql As String, Optional ByVal opt As Long = 0)
    On Error Resume Next
    db.Execute sql, opt
    If Err.Number <> 0 Then Errs "exec " & P(sql) Else Out "exec " & P(sql) & " -> " & db.RecordsAffected
End Sub

' a table: of the data source itself (remote: Jet gives no reliable types, sizes or positions of its fields with
' this driver, only their names) or linked to an Access database
Private Sub ShowTable(td As TableDef, ByVal remote As Boolean)
    Dim f As Field, ix As Index, s As String
    On Error Resume Next
    Out "table " & td.Name & " attrs=" & Hex(td.Attributes) & " src=" & td.SourceTableName & " upd=" & td.Updatable & " connect=" & Left$(P(td.Connect), 30)
    For Each f In td.Fields
        If remote Then
            Out "  field " & f.Name
        Else
            Out "  field " & f.Name & " type=" & f.Type & " size=" & f.Size & " attrs=" & f.Attributes & " req=" & f.Required & " ord=" & f.OrdinalPosition
        End If
    Next
    Errs "fields " & td.Name
    For Each ix In td.Indexes
        Out "  index (" & ix.Fields & ")"
        Errs "index fields"
        Out "    name=" & ix.Name
        Errs "index name"
        Out "    primary=" & ix.Primary
        Errs "index primary"
        Out "    unique=" & ix.Unique
        Errs "index unique"
    Next
    Errs "table " & td.Name
End Sub

Sub Main()
    Dim ws As Workspace, db As Database, mdb As Database, td As TableDef, rs As Recordset, qd As QueryDef, s As String, i As Integer
    Dim ix As Index
    OpenOut "d_dao6"
    On Error Resume Next
    gDb = App.Path & "\dao6.db"
    gMdb = App.Path & "\dao6.mdb"
    If Dir(gDb) <> "" Then Kill gDb
    If Dir(gMdb) <> "" Then Kill gMdb
    gCon = "ODBC;DSN=LBDAO6;Database=" & gDb
    Set ws = DBEngine.Workspaces(0)

    Out "== open"
    Set db = ws.OpenDatabase("", dbDriverNoPrompt, False, gCon)
    Errs "open"
    Out "name=" & P(db.Name) & " connect=" & Left$(P(db.Connect), 30) & " updatable=" & db.Updatable & " tabledefs=" & db.TableDefs.Count
    Exe db, "CREATE TABLE People (ID INTEGER PRIMARY KEY AUTOINCREMENT, Name VARCHAR(20) NOT NULL, Age SMALLINT, Born DATE, Salary NUMERIC(10,2), Active BIT, Notes TEXT, UNIQUE(Name))", dbSQLPassThrough
    Exe db, "CREATE TABLE Pets (PetID INTEGER PRIMARY KEY, Owner VARCHAR(20), Kind VARCHAR(10))", dbSQLPassThrough
    Exe db, "CREATE INDEX idxOwner ON Pets (Owner)", dbSQLPassThrough
    Exe db, "INSERT INTO People (Name, Age, Born, Salary, Active, Notes) VALUES ('Anna', 31, '1993-04-05', 1234.5, 1, 'first')", dbSQLPassThrough
    Exe db, "INSERT INTO People (Name, Age, Active) VALUES ('Bill', 25, 0)", dbSQLPassThrough
    Exe db, "INSERT INTO Pets (PetID, Owner, Kind) VALUES (1, 'Anna', 'Cat')", dbSQLPassThrough
    db.TableDefs.Refresh
    i = 0
    For Each td In db.TableDefs
        If (td.Attributes And dbSystemObject) = 0 And LCase$(Left$(td.Name, 7)) <> "sqlite_" Then ShowTable td, True: i = i + 1
    Next
    Out "tables=" & i
    Errs "tabledefs"

    Out "== Jet SQL"
    Exe db, "INSERT INTO People (Name, Age, Active) VALUES ('Cleo', 42, True)"
    Exe db, "UPDATE People SET Age = Age + 1 WHERE Name Like 'B*'"
    Exe db, "DELETE FROM People WHERE Name = 'Nobody'"
    Rows db, "SELECT * FROM People ORDER BY ID"
    Rows db, "SELECT Name, Age FROM People WHERE Age > 30 ORDER BY Name", dbOpenDynaset
    Rows db, "SELECT People.Name, Pets.Kind FROM People INNER JOIN Pets ON People.Name = Pets.Owner"
    Rows db, "SELECT Count(*) AS N, Max(Age) AS M FROM People"
    Out "== pass-through"
    Rows db, "SELECT Name, Age * 2 AS Twice FROM People ORDER BY Name", dbOpenSnapshot, dbSQLPassThrough
    Rows db, "SELECT Name FROM People", dbOpenDynaset, dbSQLPassThrough
    Exe db, "UPDATE People SET Notes = 'x' WHERE Notes IS NULL", dbSQLPassThrough

    Out "== table recordsets"
    Set rs = db.OpenRecordset("People")
    Errs "open People"
    Out "default type=" & rs.Type & " upd=" & rs.Updatable
    rs.MoveLast
    Out "count=" & rs.RecordCount
    rs.Close
    Set rs = db.OpenRecordset("People", dbOpenTable)
    Errs "open table"
    Set rs = db.OpenRecordset("People", dbOpenDynaset)
    rs.AddNew
    rs!Name = "Dino"
    rs!Age = 37
    rs!Active = False
    rs.Update
    Errs "addnew"
    rs.FindFirst "Name = 'Cleo'"
    Errs "find"
    Out "nomatch=" & rs.NoMatch
    Errs "nomatch"
    Out "name=" & rs.Fields("Name").Value
    Errs "found"
    rs.Edit
    rs!Age = 50
    rs.Update
    Errs "edit"
    rs.FindFirst "Name = 'Anna'"
    rs.Edit
    rs!Age = 60
    rs.Update
    Errs "edit anna"
    rs.CancelUpdate
    rs.FindFirst "Name = 'Bill'"
    rs.Delete
    Errs "delete"
    rs.Close
    Rows db, "SELECT ID, Name, Age, Active FROM People ORDER BY ID"

    Out "== QueryDefs"
    Set qd = db.CreateQueryDef("")
    qd.SQL = "SELECT Count(*) FROM People"
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Errs "temp qd"
    Out "temp qd=" & V(rs(0).Value)
    rs.Close
    Set qd = db.CreateQueryDef("")
    qd.Connect = gCon
    qd.SQL = "SELECT Name FROM People ORDER BY Name DESC"
    qd.ReturnsRecords = True
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    Errs "pass-through qd"
    Out "pt qd first=" & V(rs(0).Value)
    rs.Close
    Set qd = db.CreateQueryDef("Saved", "SELECT * FROM People")
    Errs "saved qd"

    Out "== transactions"
    ws.BeginTrans
    Exe db, "DELETE FROM People"
    ws.Rollback
    Rows db, "SELECT Count(*) AS N FROM People"
    ws.BeginTrans
    Exe db, "UPDATE People SET Age = 1 WHERE Name = 'Cleo'"
    ws.CommitTrans
    Rows db, "SELECT Name, Age FROM People WHERE Name = 'Cleo'"

    Out "== errors"
    Exe db, "INSERT INTO People (Name) VALUES ('Anna')"
    Exe db, "INSERT INTO NoTable (A) VALUES (1)"
    Exe db, "SELEC * FROM People", dbSQLPassThrough
    Rows db, "SELECT * FROM NoTable"
    Set mdb = ws.OpenDatabase("", dbDriverNoPrompt, False, "ODBC;Driver={No Such Driver};Database=x")
    Errs "bad driver"
    db.Close

    Out "== linked ODBC table"
    Set mdb = ws.CreateDatabase(gMdb, dbLangGeneral)
    Set td = mdb.CreateTableDef("LinkPeople")
    td.Connect = gCon
    td.SourceTableName = "People"
    mdb.TableDefs.Append td
    Errs "append link"
    ShowTable mdb.TableDefs("LinkPeople"), False
    Rows mdb, "SELECT Name, Age FROM LinkPeople ORDER BY Name"
    Set rs = mdb.OpenRecordset("LinkPeople")
    Errs "open link"
    Out "link type=" & rs.Type & " upd=" & rs.Updatable
    rs.AddNew
    rs!Name = "Emma"
    rs!Age = 22
    rs.Update
    Errs "link addnew"
    rs.Close
    Exe mdb, "UPDATE LinkPeople SET Age = Age + 100 WHERE Name = 'Emma'"
    Rows mdb, "SELECT Name, Age FROM LinkPeople WHERE Name = 'Emma'"

    Out "== DDL of DAO and SELECT INTO over ODBC"
    Set db = ws.OpenDatabase("", dbDriverNoPrompt, False, gCon)
    Errs "reopen"
    Set td = db.CreateTableDef("Things")
    td.Fields.Append td.CreateField("TID", dbLong)
    td.Fields.Append td.CreateField("Label", dbText, 30)
    Set ix = td.CreateIndex("idxLabel")
    ix.Fields = "Label"
    td.Indexes.Append ix
    db.TableDefs.Append td
    Errs "append table"
    db.TableDefs.Refresh
    ShowTable db.TableDefs("Things"), True
    Exe db, "INSERT INTO Things (TID, Label) VALUES (1, 'one')"
    Rows db, "SELECT * FROM Things"
    Set td = db.TableDefs("Things")
    td.Fields.Append td.CreateField("Extra", dbInteger)
    Errs "append field"
    db.TableDefs.Refresh
    ShowTable db.TableDefs("Things"), True
    db.TableDefs.Delete "Things"
    Errs "delete table"
    db.TableDefs.Refresh
    i = 0
    For Each td In db.TableDefs
        If td.Name = "Things" Then i = 1
    Next
    Out "things there=" & i
    mdb.Execute "CREATE TABLE Local (A LONG, B TEXT(10), C DATETIME)"
    mdb.Execute "INSERT INTO Local VALUES (1, 'x', #2001-02-03#)"
    mdb.Execute "INSERT INTO Local VALUES (2, 'y', Null)"
    Errs "local"
    Exe mdb, "SELECT * INTO [" & gCon & "].Copied FROM Local"
    Rows db, "SELECT A, B FROM Copied ORDER BY A", dbOpenSnapshot, dbSQLPassThrough
    Exe mdb, "SELECT * INTO Back FROM [" & gCon & "].Copied WHERE A > 1"
    Rows mdb, "SELECT * FROM Back"
    db.Close
    mdb.Close
    Kill gMdb
    Kill gDb
    CloseOut
End Sub
