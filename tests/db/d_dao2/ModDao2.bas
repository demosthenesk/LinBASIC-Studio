Attribute VB_Name = "ModDao2"
Option Explicit

Private Sub Q(db As Database, ByVal sql As String)
    Dim rs As Recordset
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Err.Number <> 0 Then
        ChkAlways sql
        Exit Sub
    End If
    DumpRs rs, sql
    rs.Close
    Chk sql
End Sub

' the types of the columns only (Decimal values: LinBASIC has no Decimal Variant)
Private Sub QT(db As Database, ByVal sql As String)
    Dim rs As Recordset, i As Integer, s As String
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Err.Number <> 0 Then
        ChkAlways sql
        Exit Sub
    End If
    s = "-- " & sql & vbCrLf & "  types:"
    For i = 0 To rs.Fields.Count - 1
        s = s & " " & rs.Fields(i).Name & "/" & rs.Fields(i).Type
    Next
    Out s
    rs.Close
End Sub

Sub Main()
    Dim db As Database, rs As Recordset, qd As QueryDef, pr As Parameter, p As String, i As Integer, s As String
    OpenOut "d_dao2"
    On Error Resume Next
    p = App.Path & "\dao2.mdb"
    If Dir(p) <> "" Then Kill p
    Set db = CreateDatabase(p, dbLangGeneral)
    db.Execute "CREATE TABLE T (ID LONG, Name TEXT(20), Grp TEXT(5), Amount CURRENCY, D DATETIME, F DOUBLE, B BIT, M MEMO)"
    db.Execute "INSERT INTO T VALUES (1, 'apple', 'a', 10.5, #1/15/2000#, 1.25, True, 'x')"
    db.Execute "INSERT INTO T VALUES (2, 'Banana', 'a', 20, #2/29/2000 13:45:00#, -2.5, False, Null)"
    db.Execute "INSERT INTO T VALUES (3, 'cherry', 'b', 7.25, #12/31/1999#, 0, True, '')"
    db.Execute "INSERT INTO T VALUES (4, 'Date', 'b', Null, Null, 100, False, 'long memo')"
    db.Execute "INSERT INTO T VALUES (5, 'elder*berry', 'c', 3, #6/1/2001 08:00:00#, 3.75, True, Null)"
    db.Execute "CREATE TABLE U (ID LONG, TID LONG, Tag TEXT(10))"
    db.Execute "INSERT INTO U VALUES (1, 1, 'red')"
    db.Execute "INSERT INTO U VALUES (2, 1, 'green')"
    db.Execute "INSERT INTO U VALUES (3, 3, 'red')"
    db.Execute "INSERT INTO U VALUES (4, 9, 'blue')"
    Chk "setup"

    Out "== select"
    Q db, "SELECT * FROM T"
    Q db, "SELECT Name FROM T ORDER BY Name"
    Q db, "SELECT Name FROM T ORDER BY Name DESC"
    Q db, "SELECT [Name], T.Grp FROM [T] WHERE ID IN (2,4)"
    Q db, "SELECT Name FROM T WHERE Name = 'APPLE'"
    Q db, "SELECT Name FROM T WHERE Name < 'c'"
    Q db, "SELECT Name FROM T WHERE Name Like 'b*' OR Name Like '?ate' OR Name Like '*[*]*'"
    Q db, "SELECT Name FROM T WHERE Name Like '[a-c]*' AND Name Not Like '*n*'"
    Q db, "SELECT Name FROM T WHERE Name Like ""*e*"""
    Q db, "SELECT Name FROM T WHERE D > #1/1/2000# AND D < #12/31/2000#"
    Q db, "SELECT Name FROM T WHERE D Between #1/1/2000# And #3/1/2000#"
    Q db, "SELECT Name FROM T WHERE Amount Is Null OR M Is Null"
    Q db, "SELECT Name FROM T WHERE B = True"
    Q db, "SELECT Name FROM T WHERE B"
    Q db, "SELECT Name FROM T WHERE Not B"
    Q db, "SELECT TOP 2 Name FROM T ORDER BY ID DESC"
    Q db, "SELECT DISTINCT Grp FROM T"
    Q db, "SELECT DISTINCTROW Grp FROM T"

    Out "== expressions"
    Q db, "SELECT ID * 2 AS Dbl, ID / 2 AS Half, ID \ 2 AS IDiv, ID Mod 3 AS M3, -ID AS Neg, ID ^ 2 AS Sq FROM T WHERE ID < 4"
    Q db, "SELECT Name & '!' AS Bang, Name + '?' AS Plus, Name & Null AS CatNull, Name + Null AS PlusNull, ID & '' AS IdText FROM T WHERE ID < 3"
    Q db, "SELECT 1 AS One, 'x' AS X, True AS T, #1/2/2003# AS Dt, Null AS N, ID > 2 AS Cmp, Not B AS NB FROM T WHERE ID < 4"
    QT db, "SELECT 1.5 AS OneHalf, 3000000000 AS Big, ID + 0.5 AS IdF, 1E3 AS Sci FROM T"
    Q db, "SELECT Amount * 2 AS A2, Amount + F AS AF, F * 2 AS F2, Amount / 2 AS AH, ID * F AS IF2 FROM T WHERE ID < 3"
    Q db, "SELECT UCase(Name), LCase(Name), Left(Name, 2), Right(Name, 2), Mid(Name, 2, 3), Len(Name) FROM T WHERE ID < 3"
    Q db, "SELECT Trim('  a  ') AS Tr, LTrim('  a') AS Lt, RTrim('a  ') AS Rt, InStr(Name, 'an') AS Ins, Space(2) & 'x' AS Sp, String(3, 'z') AS St FROM T WHERE ID = 2"
    Q db, "SELECT Year(D), Month(D), Day(D), Hour(D), Minute(D), Weekday(D) FROM T WHERE ID = 2"
    Q db, "SELECT Format(D, 'yyyy-mm-dd') AS F1, Format(F, '0.00') AS F2, Format(ID, '000') AS F3 FROM T WHERE ID < 3"
    Q db, "SELECT DateAdd('d', 1, D) AS Plus1, DateDiff('d', #1/1/2000#, D) AS Diff, DateSerial(2000, 1, 2) AS DS, DatePart('q', D) AS Qt FROM T WHERE ID < 3"
    Q db, "SELECT IIf(ID > 2, 'big', 'small') AS Size, IIf(ID > 2, ID, F) AS Mixed, IsNull(M) AS IsN FROM T"
    Q db, "SELECT Nz(Amount, -1) AS NzA FROM T"
    Q db, "SELECT Val('12abc') AS V1, Str(ID) AS S1, CStr(ID) AS S2, CInt('42') AS I1, CLng(2.5) AS L1, CDbl('1.5') AS D1, CCur(F) AS C1, Int(-2.5) AS I2, Fix(-2.5) AS F1, Abs(F) AS A1, Sgn(F) AS S3 FROM T WHERE ID = 2"
    Q db, "SELECT Asc(Name) AS A, Chr(65) AS C, Sqr(16) AS S, Round(F, 1) AS R FROM T WHERE ID = 1"

    Out "== aggregates"
    Q db, "SELECT Count(*) AS N, Count(Amount) AS NA, Sum(Amount) AS S, Avg(Amount) AS A, Min(Name) AS Mn, Max(D) AS Mx FROM T"
    Q db, "SELECT Grp, Count(*) AS N, Sum(ID) AS SI, Sum(F) AS SF, Avg(ID) AS AI, First(Name) AS Fi, Last(Name) AS La FROM T GROUP BY Grp ORDER BY Grp"
    Q db, "SELECT Grp, Count(*) AS N FROM T GROUP BY Grp HAVING Count(*) > 1"
    Q db, "SELECT StDev(F) AS SD, Var(F) AS VR FROM T"

    Out "== joins"
    Q db, "SELECT T.Name, U.Tag FROM T INNER JOIN U ON T.ID = U.TID ORDER BY T.Name, U.Tag"
    Q db, "SELECT T.Name, U.Tag FROM T LEFT JOIN U ON T.ID = U.TID ORDER BY T.ID, U.Tag"
    Q db, "SELECT T.Name, U.Tag FROM T RIGHT JOIN U ON T.ID = U.TID ORDER BY U.ID"
    Q db, "SELECT * FROM T, U WHERE T.ID = U.TID AND U.Tag = 'red'"
    Q db, "SELECT x.Name FROM T AS x WHERE x.ID IN (SELECT TID FROM U)"
    Q db, "SELECT Name FROM T WHERE EXISTS (SELECT * FROM U WHERE U.TID = T.ID AND Tag = 'green')"
    Q db, "SELECT Name FROM T UNION SELECT Tag FROM U"

    Out "== action queries"
    db.Execute "SELECT * INTO T2 FROM T WHERE ID < 3", dbFailOnError
    Out "select into " & db.RecordsAffected
    Q db, "SELECT * FROM T2"
    db.Execute "DELETE * FROM T2 WHERE ID = 1", dbFailOnError
    Out "delete * " & db.RecordsAffected
    db.Execute "UPDATE T2 SET Name = Name & '2', F = F * 2", dbFailOnError
    Out "update " & db.RecordsAffected
    db.Execute "INSERT INTO T2 SELECT * FROM T WHERE ID > 3", dbFailOnError
    Out "insert select " & db.RecordsAffected
    db.Execute "ALTER TABLE T2 ADD COLUMN Extra TEXT(10)", dbFailOnError
    db.Execute "ALTER TABLE T2 DROP COLUMN M", dbFailOnError
    Chk "alter"
    Q db, "SELECT * FROM T2"
    db.Execute "DROP TABLE T2", dbFailOnError
    Chk "drop"
    db.Execute "UPDATE T SET Name = 'zz' WHERE ID = 99", dbFailOnError
    Out "update none " & db.RecordsAffected
    db.Execute "SELECT * FROM T", dbFailOnError
    ChkAlways "execute select"

    Out "== querydefs"
    Set qd = db.CreateQueryDef("qBig", "SELECT Name, ID FROM T WHERE ID > 2 ORDER BY ID")
    Out "qd " & qd.Name & " type=" & qd.Type & " returns=" & qd.ReturnsRecords & " params=" & qd.Parameters.Count & " updatable=" & qd.Updatable
    Out "sql [" & qd.SQL & "]"
    Q db, "qBig"
    Q db, "SELECT * FROM qBig WHERE ID < 5"
    Set qd = db.CreateQueryDef("qParam", "PARAMETERS [Min ID] Long, pName Text ( 255 ); SELECT Name FROM T WHERE ID >= [Min ID] OR Name = pName")
    Out "qd " & qd.Name & " type=" & qd.Type & " params=" & qd.Parameters.Count
    For Each pr In qd.Parameters
        Out "  param " & pr.Name & " type=" & pr.Type & " value=" & V(pr.Value) & " dir=" & pr.Direction
    Next
    Out "sql [" & qd.SQL & "]"
    Set rs = qd.OpenRecordset
    ChkAlways "open without params"
    qd.Parameters("Min ID") = 4
    qd!pName = "apple"
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    DumpRs rs, "qParam 4 apple"
    Chk "qparam"
    Set rs = db.OpenRecordset("qParam")
    ChkAlways "open param query by name"
    Set qd = db.CreateQueryDef("qUpd", "UPDATE T SET F = F + 1 WHERE ID = 1")
    Out "qd " & qd.Name & " type=" & qd.Type & " returns=" & qd.ReturnsRecords
    qd.Execute
    Out "qd affected " & qd.RecordsAffected
    Set qd = db.CreateQueryDef("", "SELECT Count(*) FROM T")
    Out "temp qd name [" & qd.Name & "] count " & db.QueryDefs.Count
    Set rs = qd.OpenRecordset
    Out "temp result " & V(rs(0))
    Set qd = db.CreateQueryDef("qBig", "SELECT 1")
    ChkAlways "duplicate querydef"
    db.QueryDefs("qBig").SQL = "SELECT Name FROM T WHERE ID = 1"
    Q db, "qBig"
    For Each qd In db.QueryDefs
        Out "  querydef " & qd.Name & " type=" & qd.Type
    Next
    db.QueryDefs.Delete "qUpd"
    Out "querydefs " & db.QueryDefs.Count
    Chk "querydefs"
    db.Close

    Out "== biblio"
    Set db = OpenDatabase(App.Path & "\BIBLIO.MDB", False, True)
    Chk "open biblio"
    Out "biblio version " & db.Version & " updatable " & db.Updatable
    For i = 0 To db.TableDefs.Count - 1
        If (db.TableDefs(i).Attributes And dbSystemObject) = 0 Then
            Out "  table " & db.TableDefs(i).Name & " " & db.TableDefs(i).RecordCount & " fields " & db.TableDefs(i).Fields.Count
        End If
    Next
    For Each qd In db.QueryDefs
        Out "  query " & qd.Name & " type=" & qd.Type & " [" & qd.SQL & "]"
    Next
    Set rs = db.OpenRecordset("Titles")
    Out "titles type " & rs.Type & " count " & rs.RecordCount
    rs.Index = "PrimaryKey"
    rs.Seek "=", "0-0233058-8-6"
    Out "seek " & rs.NoMatch & " " & V(rs!Title)
    rs.Close
    Set rs = db.OpenRecordset("SELECT TOP 5 Authors.Author, Titles.Title, Titles.[Year Published] FROM (Authors INNER JOIN [Title Author] ON Authors.Au_ID = [Title Author].Au_ID) INNER JOIN Titles ON [Title Author].ISBN = Titles.ISBN WHERE Authors.Author Like 'Date*' ORDER BY Titles.Title")
    DumpRs rs, "biblio join"
    Set rs = db.OpenRecordset("All Titles")
    rs.MoveLast
    Out "all titles " & rs.RecordCount & " type " & rs.Type
    Set rs = db.OpenRecordset("SELECT [Year Published], Count(*) AS N FROM Titles GROUP BY [Year Published] HAVING Count(*) > 600 ORDER BY [Year Published]")
    DumpRs rs, "biblio years"
    rs.Edit
    ChkAlways "edit readonly"
    db.Close
    Out "end"
    CloseOut
End Sub
