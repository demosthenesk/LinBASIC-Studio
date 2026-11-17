Attribute VB_Name = "ModDao1"
Option Explicit

Private Sub ShowTable(db As Database, ByVal name As String)
    Dim td As TableDef, fl As Field, ix As Index, i As Integer, s As String
    Set td = db.TableDefs(name)
    Out "table " & td.Name & " attr=" & td.Attributes & " fields=" & td.Fields.Count & " indexes=" & td.Indexes.Count & " updatable=" & td.Updatable
    For Each fl In td.Fields
        Out "  field " & fl.Name & " type=" & fl.Type & " size=" & fl.Size & " attr=" & fl.Attributes & _
            " req=" & fl.Required & " zl=" & fl.AllowZeroLength & " pos=" & fl.OrdinalPosition & _
            " def=" & V(fl.DefaultValue)
    Next
    For Each ix In td.Indexes
        s = ""
        For i = 0 To ix.Fields.Count - 1
            s = s & " " & ix.Fields(i).Name & "/" & ix.Fields(i).Attributes
        Next
        Out "  index " & ix.Name & " primary=" & ix.Primary & " unique=" & ix.Unique & " req=" & ix.Required & _
            " ignorenulls=" & ix.IgnoreNulls & " foreign=" & ix.Foreign & " :" & s
    Next
End Sub

Sub Main()
    Dim ws As Workspace, db As Database, td As TableDef, fl As Field, ix As Index, rel As Relation
    Dim rs As Recordset, rs2 As Recordset, p As String, i As Integer, s As String, bm As Variant, n As Long
    OpenOut "d_dao1"
    On Error Resume Next

    Out "== engine"
    Out TypeName(DBEngine) & " " & DBEngine.Version & " ws=" & DBEngine.Workspaces.Count
    Set ws = DBEngine.Workspaces(0)
    Out TypeName(ws) & " " & ws.Name & " user=" & ws.UserName & " dbs=" & ws.Databases.Count & " type=" & ws.Type
    Chk "engine"

    Out "== create"
    p = App.Path & "\dao1.mdb"
    If Dir(p) <> "" Then Kill p
    Set db = DBEngine.CreateDatabase(p, dbLangGeneral)
    Chk "create"
    Out TypeName(db) & " " & FileOf(db.Name) & " updatable=" & db.Updatable & " dbs=" & ws.Databases.Count
    Out "tabledefs " & db.TableDefs.Count & " querydefs " & db.QueryDefs.Count & " relations " & db.Relations.Count
    For Each td In db.TableDefs
        Out "  " & td.Name & " attr=" & td.Attributes
    Next
    Chk "tables"

    Out "== ddl"
    db.Execute "CREATE TABLE Authors (Au_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(30) NOT NULL, " & _
        "Born SHORT, Salary CURRENCY, Hired DATETIME, Active BIT, Notes MEMO, Rate DOUBLE, Score SINGLE, Age BYTE, Big LONG)"
    Chk "create table"
    db.Execute "CREATE INDEX ByName ON Authors (Name)"
    Chk "create index"
    ShowTable db, "Authors"
    Chk "show"
    db.TableDefs.Refresh
    ShowTable db, "Authors"
    Chk "show after refresh"

    Out "== tabledef objects"
    Set td = db.CreateTableDef("Books")
    Set fl = td.CreateField("BookID", dbLong)
    fl.Attributes = fl.Attributes Or dbAutoIncrField
    td.Fields.Append fl
    td.Fields.Append td.CreateField("Title", dbText, 40)
    td.Fields.Append td.CreateField("AuID", dbLong)
    td.Fields.Append td.CreateField("Price", dbCurrency)
    td.Fields.Append td.CreateField("Pub", dbDate)
    Set fl = td.CreateField("Pages", dbInteger)
    fl.DefaultValue = "100"
    td.Fields.Append fl
    Set fl = td.CreateField("Code", dbText, 5)
    fl.Required = True
    td.Fields.Append fl
    Out "unappended fields " & td.Fields.Count & " " & td.Fields(1).Name & " size=" & td.Fields(1).Size
    Set ix = td.CreateIndex("PrimaryKey")
    ix.Primary = True
    ix.Fields.Append ix.CreateField("BookID")
    td.Indexes.Append ix
    Set ix = td.CreateIndex("ByTitle")
    Set fl = ix.CreateField("Title")
    fl.Attributes = dbDescending
    ix.Fields.Append fl
    td.Indexes.Append ix
    Set ix = td.CreateIndex("ByCode")
    ix.Unique = True
    ix.Fields.Append ix.CreateField("Code")
    td.Indexes.Append ix
    db.TableDefs.Append td
    Chk "append books"
    Set rel = db.CreateRelation("AuthorBooks", "Authors", "Books")
    Set fl = rel.CreateField("Au_ID")
    fl.ForeignName = "AuID"
    rel.Fields.Append fl
    db.Relations.Append rel
    Chk "relation"
    Out "relations " & db.Relations.Count & " " & db.Relations(0).Name & " " & db.Relations(0).Table & "->" & _
        db.Relations(0).ForeignTable & " attr=" & db.Relations(0).Attributes & " " & _
        db.Relations(0).Fields(0).Name & "=" & db.Relations(0).Fields(0).ForeignName
    db.TableDefs.Refresh
    ShowTable db, "Books"
    Out "tabledefs " & db.TableDefs.Count
    Chk "books"

    Out "== insert"
    db.Execute "INSERT INTO Authors (Name, Born, Salary, Hired, Active, Rate, Score, Age, Big) " & _
        "VALUES ('Smith', 1950, 1234.5, #3/15/1990#, True, 1.5, 2.25, 40, 100000)", dbFailOnError
    Out "affected " & db.RecordsAffected
    db.Execute "INSERT INTO Authors (Name, Born, Hired, Active) VALUES (""O'Brien"", 1962, #12/1/1985 10:30:00#, False)"
    db.Execute "INSERT INTO Authors (Name, Born, Notes) VALUES ('jones', 1971, 'memo text')"
    db.Execute "INSERT INTO Authors (Name) VALUES ('Adams')"
    Out "affected " & db.RecordsAffected
    Chk "insert"
    db.Execute "INSERT INTO Authors (Born) VALUES (1900)", dbFailOnError
    ChkAlways "insert without name"
    db.Execute "INSERT INTO Nope (A) VALUES (1)", dbFailOnError
    ChkAlways "insert into nope"
    db.Execute "INSERT INTO Authors (Name, Born) VALUES ('Too long name for the thirty char field', 1)", dbFailOnError
    ChkAlways "too long"
    db.Execute "UPDATE Authors SET Born = Born + 1 WHERE Born > 1960", dbFailOnError
    Out "update affected " & db.RecordsAffected
    Chk "update"

    Out "== recordsets"
    Set rs = db.OpenRecordset("Authors")
    Out "default type " & rs.Type & " count " & rs.RecordCount & " updatable " & rs.Updatable & " name " & rs.Name & " editmode " & rs.EditMode
    Chk "open table"
    DumpRs rs, "Authors (table)"
    rs.Close
    Set rs = db.OpenRecordset("SELECT * FROM Authors ORDER BY Name", dbOpenDynaset)
    Out "dynaset type " & rs.Type & " count " & rs.RecordCount & " bof " & rs.BOF & " eof " & rs.EOF & " pos " & rs.AbsolutePosition
    rs.MoveLast
    Out "after movelast count " & rs.RecordCount & " pos " & rs.AbsolutePosition & " name " & rs!Name
    rs.MoveFirst
    Out "first " & V(rs!Au_ID) & " " & V(rs!Name) & " " & V(rs("Born")) & " " & V(rs.Fields(3)) & " " & V(rs!Hired) & _
        " " & V(rs!Active) & " " & V(rs!Notes) & " " & V(rs!Rate) & " " & V(rs!Score) & " " & V(rs!Age) & " " & V(rs!Big)
    rs.MovePrevious
    Out "before first bof " & rs.BOF & " eof " & rs.EOF & " pos " & rs.AbsolutePosition
    s = rs!Name
    ChkAlways "field at bof"
    rs.MovePrevious
    ChkAlways "moveprevious at bof"
    rs.Move 2
    Out "move 2 from bof: " & rs!Name & " pos " & rs.AbsolutePosition
    rs.Move -1
    Out "move -1: " & rs!Name
    rs.AbsolutePosition = 3
    Out "absolute 3: " & rs!Name
    rs.MoveLast
    rs.MoveNext
    Out "after last eof " & rs.EOF & " pos " & rs.AbsolutePosition
    rs.MoveNext
    ChkAlways "movenext at eof"
    Out "field types:"
    For i = 0 To rs.Fields.Count - 1
        With rs.Fields(i)
            Out "  " & .Name & " type=" & .Type & " size=" & .Size & " src=" & .SourceTable & "." & .SourceField & " upd=" & .DataUpdatable
        End With
    Next
    Chk "fields"
    s = rs!Nope
    ChkAlways "no field"
    s = rs.Fields(99).Name
    ChkAlways "field 99"
    rs.Close

    Out "== snapshot"
    Set rs = db.OpenRecordset("Authors", dbOpenSnapshot)
    Out "snapshot type " & rs.Type & " updatable " & rs.Updatable & " count " & rs.RecordCount
    rs.Edit
    ChkAlways "edit snapshot"
    rs.Close
    Set rs = db.OpenRecordset("Nope")
    ChkAlways "open nope"
    Set rs = db.OpenRecordset("SELEC * FROM Authors")
    ChkAlways "bad sql"
    Set rs = db.OpenRecordset("SELECT * FROM Authors WHERE Nope = 1")
    ChkAlways "unknown column"
    Set rs = db.OpenRecordset("SELECT * FROM Authors WHERE (Born > 1")
    ChkAlways "bad where"

    Out "== addnew / edit"
    Set rs = db.OpenRecordset("SELECT * FROM Authors ORDER BY Au_ID", dbOpenDynaset)
    rs.MoveLast
    Out "count " & rs.RecordCount & " last " & rs!Name
    rs.AddNew
    Out "editmode " & rs.EditMode & " new id " & V(rs!Au_ID) & " name " & V(rs!Name) & " pos " & rs.AbsolutePosition
    rs!Name = "Baker"
    rs!Born = 1980
    rs!Hired = #7/4/2001#
    rs!Salary = 99.99
    rs.Update
    Chk "addnew"
    Out "after update current " & rs!Name & " editmode " & rs.EditMode & " count " & rs.RecordCount
    rs.Bookmark = rs.LastModified
    Out "lastmodified " & rs!Name & " " & V(rs!Au_ID) & " " & V(rs!Salary)
    rs.Update
    ChkAlways "update without edit"
    rs.Edit
    rs!Born = "abc"
    ChkAlways "type mismatch"
    rs!Name = Null
    rs.Update
    ChkAlways "required null"
    rs.CancelUpdate
    rs.Edit
    rs!Name = "Baker Jr"
    rs!Notes = ""
    rs.Update
    Chk "edit"
    Out "edited " & rs!Name & " notes " & V(rs!Notes)
    rs.AddNew
    rs!Name = "x"
    rs.CancelUpdate
    Out "cancel editmode " & rs.EditMode & " current " & rs!Name
    Chk "cancel"
    bm = rs.Bookmark
    Out "bookmark " & V(bm)
    rs.MoveFirst
    rs.Bookmark = bm
    Out "back to " & rs!Name
    Set rs2 = rs.Clone
    Out "clone current? " & V(rs2.AbsolutePosition)
    ChkAlways "clone pos"
    rs2.Bookmark = bm
    Out "clone at bookmark " & rs2!Name
    rs.MoveFirst
    rs.Delete
    Out "deleted, editmode " & rs.EditMode
    s = rs!Name
    ChkAlways "field after delete"
    rs.MoveNext
    Out "next after delete " & rs!Name & " count " & rs.RecordCount
    rs.Requery
    rs.MoveLast
    Out "requery count " & rs.RecordCount
    Chk "requery"
    rs.Close

    Out "== find"
    Set rs = db.OpenRecordset("SELECT * FROM Authors ORDER BY Name", dbOpenDynaset)
    rs.FindFirst "Name Like 'b*'"
    Out "findfirst b* " & rs.NoMatch & " " & rs!Name
    rs.FindNext "Name Like 'b*'"
    Out "findnext b* " & rs.NoMatch & " " & rs!Name
    rs.FindNext "Name Like 'b*'"
    Out "findnext again " & rs.NoMatch & " " & rs!Name
    rs.FindLast "Born > 1900"
    Out "findlast born>1900 " & rs.NoMatch & " " & rs!Name
    rs.FindPrevious "Born > 1900"
    Out "findprevious " & rs.NoMatch & " " & rs!Name
    rs.FindFirst "Hired = #3/15/1990#"
    Out "find date " & rs.NoMatch & " " & rs!Name
    rs.FindFirst "Name = 'SMITH'"
    Out "find upper " & rs.NoMatch & " " & rs!Name
    rs.FindFirst "Name = 'Nobody'"
    Out "find none " & rs.NoMatch & " pos " & rs.AbsolutePosition
    rs.FindFirst "Nope = 1"
    ChkAlways "find bad"
    rs.FindFirst "IsNull(Born)"
    Out "find isnull " & rs.NoMatch
    Chk "find"
    rs.Close

    Out "== seek"
    Set rs = db.OpenRecordset("Authors", dbOpenTable)
    rs.Seek "=", 2
    ChkAlways "seek without index"
    rs.Index = "PrimaryKey"
    rs.Seek "=", 2
    Out "seek =2 " & rs.NoMatch & " " & rs!Name
    rs.Seek ">=", 4
    Out "seek >=4 " & rs.NoMatch & " " & rs!Name
    rs.Seek ">", 100
    Out "seek >100 " & rs.NoMatch
    rs.Seek "<", 3
    Out "seek <3 " & rs.NoMatch & " " & rs!Name
    rs.Index = "ByName"
    rs.MoveFirst
    Out "byname first " & rs!Name
    rs.Seek "=", "JONES"
    Out "seek jones " & rs.NoMatch & " " & rs!Name
    rs.Seek ">=", "C"
    Out "seek >=C " & rs.NoMatch & " " & rs!Name
    Chk "seek"
    rs.Close

    Out "== sort / filter"
    Set rs = db.OpenRecordset("SELECT Name, Born FROM Authors", dbOpenDynaset)
    rs.Sort = "Born DESC"
    Set rs2 = rs.OpenRecordset
    DumpRs rs2, "sorted born desc"
    rs.Filter = "Born < 1970"
    Set rs2 = rs.OpenRecordset
    DumpRs rs2, "filter born < 1970"
    Chk "sort"

    Out "== constraints"
    db.Execute "INSERT INTO Books (Title, AuID, Code) VALUES ('Book A', 2, 'A1')", dbFailOnError
    db.Execute "INSERT INTO Books (Title, AuID, Code, Price, Pub) VALUES ('Book B', 3, 'B1', 12.5, #1/31/1999#)", dbFailOnError
    Chk "books insert"
    db.Execute "INSERT INTO Books (Title, AuID, Code) VALUES ('Book C', 77, 'C1')", dbFailOnError
    ChkAlways "no author"
    db.Execute "INSERT INTO Books (Title, AuID, Code) VALUES ('Book D', 2, 'A1')", dbFailOnError
    ChkAlways "duplicate code"
    db.Execute "INSERT INTO Books (Title, AuID) VALUES ('Book E', 2)", dbFailOnError
    ChkAlways "required code"
    db.Execute "DELETE FROM Authors WHERE Au_ID = 2", dbFailOnError
    ChkAlways "delete referenced"
    Set rs = db.OpenRecordset("Books", dbOpenDynaset)
    rs.AddNew
    rs!Title = "Book F"
    rs!AuID = 3
    rs!Code = ""
    rs.Update
    ChkAlways "zero length"
    rs.CancelUpdate
    rs.AddNew
    rs!Title = "Book G"
    rs!AuID = 3
    rs!Code = "A1"
    rs.Update
    ChkAlways "duplicate by recordset"
    rs.CancelUpdate
    rs.AddNew
    rs!Title = "Book H"
    rs!AuID = 3
    rs!Code = "H1"
    Out "default pages " & V(rs!Pages)
    rs.Update
    Chk "book h"
    DumpRs rs, "books"
    rs.Close

    Out "== transactions"
    ws.BeginTrans
    db.Execute "DELETE FROM Books"
    Out "in trans affected " & db.RecordsAffected
    ws.Rollback
    Out "after rollback " & db.OpenRecordset("SELECT Count(*) AS N FROM Books")!N
    BeginTrans
    db.Execute "UPDATE Books SET Price = 1"
    CommitTrans
    Out "after commit " & V(db.OpenRecordset("SELECT Sum(Price) FROM Books").Fields(0).Value)
    ws.CommitTrans
    ChkAlways "commit without begin"

    Out "== close / reopen"
    db.Close
    Out "dbs after close " & ws.Databases.Count
    Out "closed name " & db.Name
    ChkAlways "use closed"
    Set db = OpenDatabase(p)
    Out "reopened " & FileOf(db.Name) & " tables " & db.TableDefs.Count & " dbs " & ws.Databases.Count & " " & (ws.Databases(0).Name = db.Name)
    ShowTable db, "Books"
    Set rs = db.OpenRecordset("Authors")
    Out "authors " & rs.RecordCount
    rs.Close
    db.TableDefs.Delete "Nope"
    ChkAlways "delete nope"
    db.Execute "CREATE TABLE Temp (A LONG, B TEXT)"
    Out "tables " & db.TableDefs.Count
    db.TableDefs.Delete "Temp"
    db.TableDefs.Refresh
    Out "tables after delete " & db.TableDefs.Count
    s = db.TableDefs("Temp").Name
    ChkAlways "deleted table"
    db.Close
    Set db = OpenDatabase(App.Path & "\nofile.mdb")
    Out "  ! open missing: " & Err.Number & " " & Err.Source
    Err.Clear
    Out "errors " & DBEngine.Errors.Count & " " & DBEngine.Errors(0).Number & " " & DBEngine.Errors(0).Source
    Out "end"
    CloseOut
End Sub
