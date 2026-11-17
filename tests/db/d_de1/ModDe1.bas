Attribute VB_Name = "ModDe1"
Option Explicit

Private Function MakeDb() As String
    Dim e As Object, db As Object, p As String
    p = App.Path & "\de1.mdb"
    If Dir(p) <> "" Then Kill p
    Set e = CreateObject("DAO.DBEngine.36")
    Set db = e.CreateDatabase(p, ";LANGID=0x0409;CP=1252;COUNTRY=0", 64)
    db.Execute "CREATE TABLE Customers (CustID TEXT(5) CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(30), Country TEXT(15))"
    db.Execute "INSERT INTO Customers VALUES ('ALFKI', 'Alfreds', 'Germany')"
    db.Execute "INSERT INTO Customers VALUES ('BONAP', 'Bon app', 'France')"
    db.Execute "INSERT INTO Customers VALUES ('BLAUS', 'Blauer See', 'Germany')"
    db.Execute "INSERT INTO Customers VALUES ('NOORD', 'Noord', 'Netherlands')"
    db.Execute "CREATE TABLE Orders (OrderID LONG CONSTRAINT PrimaryKey PRIMARY KEY, CustID TEXT(5), Amount CURRENCY, ODate DATETIME)"
    db.Execute "INSERT INTO Orders VALUES (10, 'ALFKI', 100.5, #2026-01-05#)"
    db.Execute "INSERT INTO Orders VALUES (11, 'ALFKI', 20, #2026-02-01#)"
    db.Execute "INSERT INTO Orders VALUES (12, 'BONAP', 7.25, #2026-02-03#)"
    db.Execute "INSERT INTO Orders VALUES (13, 'BLAUS', 300, #2026-03-01#)"
    db.Execute "INSERT INTO Orders VALUES (14, 'ALFKI', 1, #2026-03-09#)"
    db.Execute "CREATE TABLE Items (ItemID LONG CONSTRAINT PrimaryKey PRIMARY KEY, OrderID LONG, Product TEXT(20), Qty SHORT, Price CURRENCY)"
    db.Execute "INSERT INTO Items VALUES (1, 10, 'Tea', 2, 10)"
    db.Execute "INSERT INTO Items VALUES (2, 10, 'Jam', 1, 80.5)"
    db.Execute "INSERT INTO Items VALUES (3, 11, 'Tea', 4, 5)"
    db.Execute "INSERT INTO Items VALUES (4, 13, 'Ale', 10, 30)"
    db.Execute "INSERT INTO Items VALUES (5, 12, 'Bread', 1, 7.25)"
    db.Close
    MakeDb = p
End Function

Private Sub ShowRs(rs As ADODB.Recordset, ByVal ind As String, ByVal depth As Integer)
    Dim f As ADODB.Field, s As String, n As Integer
    For Each f In rs.Fields
        Out ind & "field " & f.Name & " type=" & f.Type & " size=" & f.DefinedSize & " attr=" & f.Attributes
    Next
    If rs.BOF And rs.EOF Then Out ind & "(no rows)": Exit Sub
    rs.MoveFirst
    Do Until rs.EOF
        s = ind & "row"
        For Each f In rs.Fields
            If f.Type = 136 Then s = s & " " & f.Name & "=[" & f.Value.RecordCount & "]" Else s = s & " " & V(f.Value)
        Next
        Out s
        If depth > 0 Then
            For Each f In rs.Fields
                If f.Type = 136 Then ShowRs f.Value, ind & "    ", depth - 1
            Next
        End If
        rs.MoveNext
        n = n + 1
        If n > 20 Then Exit Do
    Loop
End Sub

Private Sub ShowCmd(ByVal name As String)
    Dim c As ADODB.Command
    Set c = deTest.Commands(name)
    Chk "command " & name
    Out "command " & name & " type=" & c.CommandType & " params=" & c.Parameters.Count & " text=" & c.CommandText
    Chk "command text " & name
End Sub

Sub Main()
    Dim i As Integer, rs As ADODB.Recordset
    OpenOut "d_de1"
    On Error Resume Next
    MakeDb
    Chk "makedb"
    Out "== the object"
    Out "typename=" & TypeName(deTest)
    Out "connections=" & deTest.Connections.Count & " commands=" & deTest.Commands.Count & " recordsets=" & deTest.Recordsets.Count
    Chk "counts"
    Out "conn1 " & TypeName(deTest.Conn1) & " state=" & deTest.Conn1.State & " cs=" & deTest.Conn1.ConnectionString
    Chk "conn1"
    Out "conn1 cursorloc=" & deTest.Conn1.CursorLocation & " provider=" & deTest.Conn1.Provider
    For i = 1 To deTest.Commands.Count
        Out "commands(" & i & ")=" & TypeName(deTest.Commands(i)) & " " & deTest.Commands(i).Name
    Next
    Chk "commands loop"
    For i = 1 To deTest.Recordsets.Count
        Out "recordsets(" & i & ") state=" & deTest.Recordsets(i).State
    Next
    Chk "recordsets loop"
    ShowCmd "Customers_Grouping"
    ShowCmd "Customers"
    ShowCmd "Orders"
    ShowCmd "Items"
    ShowCmd "Germans"
    Out "rsGermans " & TypeName(deTest.rsGermans) & " state=" & deTest.rsGermans.State
    Chk "rsGermans before"

    Out "== a plain command"
    deTest.Germans
    Chk "Germans"
    Set rs = deTest.rsGermans
    Out "state=" & rs.State & " cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " cursorloc=" & rs.CursorLocation & " count=" & rs.RecordCount & " source=" & rs.Source
    Out "conn1 state=" & deTest.Conn1.State & " provider=" & deTest.Conn1.Provider & " cursorloc=" & deTest.Conn1.CursorLocation
    ShowRs rs, "  ", 0
    deTest.Germans
    Chk "Germans again (open)"
    rs.Close
    deTest.Germans
    Chk "Germans after close"
    Out "count=" & deTest.rsGermans.RecordCount
    deTest.rsGermans.Close

    Out "== the hierarchy"
    deTest.Customers_Grouping
    Chk "Customers_Grouping"
    Set rs = deTest.rsCustomers_Grouping
    Chk "rsCustomers_Grouping"
    Out "state=" & rs.State & " cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " count=" & rs.RecordCount
    Out "source=" & rs.Source
    ShowRs rs, "  ", 3
    rs.Close

    Out "== the commands of the hierarchy, through Commands"
    Out "commands(Customers) state=" & deTest.Commands("Customers").State
    Chk "commands customers"
    Set rs = deTest.Recordsets("Customers")
    Chk "recordsets(Customers)"
    Out "recordsets(Customers) " & TypeName(rs) & " state=" & rs.State
    Chk "recordsets customers state"

    Out "== update through the hierarchy"
    deTest.Customers_Grouping
    Set rs = deTest.rsCustomers_Grouping
    Out "locktype=" & rs.LockType
    rs.MoveFirst
    rs.MoveNext
    Dim ch As ADODB.Recordset
    Set ch = rs("Customers").Value
    Chk "chapter"
    Out "chapter locktype=" & ch.LockType & " first=" & V(ch!Name.Value)
    ch!Name = "Alfreds F"
    ch.Update
    Chk "update"
    rs.Close
    deTest.Germans
    Out "server " & V(deTest.rsGermans!Name.Value)
    deTest.rsGermans.Close

    Out "== close"
    deTest.Conn1.Close
    Chk "close conn"
    Out "conn1 state=" & deTest.Conn1.State
    deTest.Germans
    Chk "Germans after conn close"
    Out "rsGermans state=" & deTest.rsGermans.State & " conn1 state=" & deTest.Conn1.State
    Out "end"
    CloseOut
End Sub
