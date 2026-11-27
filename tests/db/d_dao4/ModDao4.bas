Attribute VB_Name = "ModDao4"
Option Explicit
' The Properties collections of DAO (phase 12.4: the Database Window of the Visual Data Manager lists them): the
' names, types, Inherited and values of the properties of the DBEngine, the Workspace, the Database, a TableDef, its
' Fields and Indexes, a QueryDef and its Fields, a Relation, a Recordset and its Fields; a property added, changed and
' deleted (CreateProperty, Append, Delete).

Private gPath As String

Private Function PV(ByVal x As Variant) As String
    Dim s As String
    If VarType(x) = vbDate Then PV = "D(date)": Exit Function
    s = V(x)
    If Len(gPath) > 0 Then s = Replace(s, gPath, "<path>")
    PV = s
End Function

Private Sub DumpProps(ByVal what As String, ps As Properties)
    Dim p As Property, i As Integer, x As Variant, nm As String, t As Integer, inh As Boolean
    On Error Resume Next
    Out "== " & what & " count=" & ps.Count
    Chk "count"
    For i = 0 To ps.Count - 1
        Set p = ps(i)
        nm = p.Name
        t = p.Type
        inh = p.Inherited
        Err.Clear
        x = Empty
        x = p.Value
        If Err.Number <> 0 Then
            Out "  " & nm & " type=" & t & " inh=" & inh & " err=" & Err.Number
            Err.Clear
        Else
            Out "  " & nm & " type=" & t & " inh=" & inh & " = " & PV(x)
        End If
    Next
End Sub

Sub Main()
    Dim e As DBEngine, db As Database, td As TableDef, f As Field, ix As Index, qd As QueryDef, rs As Recordset
    Dim rel As Relation, p As Property, s As String
    OpenOut "d_dao4"
    On Error Resume Next
    gPath = App.Path & "\dao4.mdb"
    If Dir(gPath) <> "" Then Kill gPath
    Set db = DBEngine.CreateDatabase(gPath, dbLangGeneral, dbVersion40)
    Chk "create"
    db.Execute "CREATE TABLE People (ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(30) NOT NULL, " & _
               "Born DATETIME, Salary CURRENCY, Notes MEMO, Active BIT, Score DOUBLE, Grade BYTE)"
    db.Execute "CREATE INDEX idxName ON People (Name DESC)"
    db.Execute "CREATE TABLE Pets (PetID LONG CONSTRAINT PetKey PRIMARY KEY, Owner LONG, Kind TEXT(10))"
    db.Execute "INSERT INTO People (Name, Born, Salary, Notes, Active, Score, Grade) VALUES " & _
               "('Ann', #1980-02-03#, 1200.5, 'x', True, 2.5, 3)"
    Chk "tables"
    Set rel = db.CreateRelation("PeoplePets", "People", "Pets", dbRelationUpdateCascade)
    Set f = rel.CreateField("ID")
    f.ForeignName = "Owner"
    rel.Fields.Append f
    db.Relations.Append rel
    Chk "relation"
    Set qd = db.CreateQueryDef("Adults", "SELECT Name, Born FROM People WHERE Born < #2000-01-01#")
    Chk "querydef"

    DumpProps "DBEngine", DBEngine.Properties
    DumpProps "Workspace", DBEngine.Workspaces(0).Properties
    DumpProps "Database", db.Properties
    Set td = db.TableDefs("People")
    DumpProps "TableDef People", td.Properties
    For Each f In td.Fields
        DumpProps "Field People." & f.Name, f.Properties
    Next
    For Each ix In td.Indexes
        DumpProps "Index People." & ix.Name, ix.Properties
    Next
    Set qd = db.QueryDefs("Adults")
    DumpProps "QueryDef Adults", qd.Properties
    DumpProps "QueryDef Field Adults.Name", qd.Fields("Name").Properties
    DumpProps "Relation PeoplePets", db.Relations("PeoplePets").Properties
    Set rs = db.OpenRecordset("People", dbOpenDynaset)
    DumpProps "Recordset People", rs.Properties
    DumpProps "Recordset Field People.Name", rs.Fields("Name").Properties
    rs.Close

    Out "== user properties"
    Set td = db.TableDefs("People")
    Set p = td.CreateProperty("Description", dbText, "The people")
    td.Properties.Append p
    Chk "append"
    Out "td description=" & PV(td.Properties("Description").Value) & " count=" & td.Properties.Count
    td.Properties("Description").Value = "Changed"
    Chk "change"
    Out "td description=" & PV(td.Properties("Description").Value)
    Set f = td.Fields("Name")
    Set p = f.CreateProperty("Caption", dbText, "Full Name")
    f.Properties.Append p
    Chk "field append"
    Out "field caption=" & PV(td.Fields("Name").Properties("Caption").Value)
    td.Properties.Append td.CreateProperty("Description", dbText, "twice")
    ChkAlways "append twice"
    s = td.Properties("NoSuch").Value
    ChkAlways "no such property"
    td.Properties.Delete "Description"
    Chk "delete"
    Out "count after delete=" & td.Properties.Count
    td.Properties.Delete "Name"
    ChkAlways "delete built-in"
    db.Close
    Set db = DBEngine.OpenDatabase(gPath)
    Out "reopened caption=" & PV(db.TableDefs("People").Fields("Name").Properties("Caption").Value)
    Chk "reopened"
    db.Close
    Kill gPath
    CloseOut
End Sub
