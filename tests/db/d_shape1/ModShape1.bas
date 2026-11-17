Attribute VB_Name = "ModShape1"
Option Explicit

Private Function MakeDb() As String
    Dim e As Object, db As Object, p As String
    p = App.Path & "\shape1.mdb"
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

Private Sub ShowFields(rs As ADODB.Recordset, ByVal ind As String)
    Dim f As ADODB.Field
    For Each f In rs.Fields
        Out ind & "field " & f.Name & " type=" & f.Type & " size=" & f.DefinedSize & " attr=" & f.Attributes
    Next
End Sub

Private Sub Rows(rs As ADODB.Recordset, ByVal ind As String, ByVal depth As Integer)
    Dim f As ADODB.Field, s As String, n As Integer
    If rs.BOF And rs.EOF Then Out ind & "(no rows)": Exit Sub
    rs.MoveFirst
    Do Until rs.EOF
        s = ind
        For Each f In rs.Fields
            If f.Type = 136 Then
                s = s & " " & f.Name & "=[" & f.Value.RecordCount & "]"
            Else
                s = s & " " & V(f.Value)
            End If
        Next
        Out s
        If depth > 0 Then
            For Each f In rs.Fields
                If f.Type = 136 Then Rows f.Value, ind & "    ", depth - 1
            Next
        End If
        rs.MoveNext
        n = n + 1
        If n > 30 Then Exit Do
    Loop
End Sub

Private Function Op(cn As ADODB.Connection, ByVal sql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
    Set Op = rs
End Function

Sub Main()
    Dim cn As ADODB.Connection, rs As ADODB.Recordset, ch As ADODB.Recordset, ch2 As ADODB.Recordset, p As String, s As String
    OpenOut "d_shape1"
    On Error Resume Next
    p = MakeDb()
    Chk "makedb"
    Set cn = New ADODB.Connection
    cn.Open "Provider=MSDataShape;Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & p
    Chk "open"
    Out "provider=" & cn.Provider & " state=" & cn.State & " cursorloc=" & cn.CursorLocation

    Out "== append relate"
    Set rs = Op(cn, "SHAPE {SELECT * FROM Customers ORDER BY CustID} APPEND ({SELECT * FROM Orders ORDER BY OrderID} AS rsOrders RELATE CustID TO CustID)")
    Chk "shape1"
    Out "cursortype=" & rs.CursorType & " locktype=" & rs.LockType & " cursorloc=" & rs.CursorLocation & " count=" & rs.RecordCount
    Out "source=" & rs.Source
    ShowFields rs, "  "
    Set ch = rs("rsOrders").Value
    Chk "chapter value"
    Out "child " & TypeName(ch) & " count=" & ch.RecordCount & " cursortype=" & ch.CursorType & " locktype=" & ch.LockType & " source=" & ch.Source
    ShowFields ch, "    "
    Rows rs, "  ", 1
    rs.MoveFirst
    Set ch = rs("rsOrders").Value
    Out "sync first count=" & ch.RecordCount
    rs.MoveNext
    Out "sync next count=" & ch.RecordCount & " same=" & (ch Is rs("rsOrders").Value)
    Set ch2 = rs("rsOrders").Value
    Out "after move new object count=" & ch2.RecordCount
    rs.Close

    Out "== nested with aggregates"
    s = "SHAPE {SELECT CustID, Name FROM Customers ORDER BY CustID} APPEND " & _
        "((SHAPE {SELECT OrderID, CustID, Amount FROM Orders ORDER BY OrderID} APPEND ({SELECT OrderID, Product, Qty, Price FROM Items ORDER BY ItemID} AS rsItems RELATE OrderID TO OrderID), " & _
        "SUM(rsItems.Qty) AS TotQty, COUNT(rsItems.Product) AS NItems) AS rsOrders RELATE CustID TO CustID), " & _
        "SUM(rsOrders.Amount) AS Total, COUNT(rsOrders.OrderID) AS NOrders, AVG(rsOrders.Amount) AS AvgAmt, MAX(rsOrders.Amount) AS MaxAmt, MIN(rsOrders.Amount) AS MinAmt"
    Set rs = Op(cn, s)
    Chk "shape2"
    ShowFields rs, "  "
    ShowFields rs("rsOrders").Value, "    "
    Rows rs, "  ", 2
    rs.Close

    Out "== compute by"
    Set rs = Op(cn, "SHAPE {SELECT * FROM Customers ORDER BY CustID} AS rsCust COMPUTE rsCust, COUNT(rsCust.CustID) AS N BY Country")
    Chk "compute"
    Out "count=" & rs.RecordCount
    ShowFields rs, "  "
    Rows rs, "  ", 1
    rs.Close
    Set rs = Op(cn, "SHAPE {SELECT * FROM Orders} AS rsO COMPUTE SUM(rsO.Amount) AS Grand, rsO")
    Chk "compute all"
    ShowFields rs, "  "
    Rows rs, "  ", 0
    rs.Close

    Out "== calc"
    Set rs = Op(cn, "SHAPE {SELECT OrderID, Amount FROM Orders ORDER BY OrderID} APPEND CALC(Amount * 2) AS Dbl, CALC(IIf(Amount > 50, ""big"", ""small"")) AS Sz")
    Chk "calc"
    ShowFields rs, "  "
    Rows rs, "  ", 0
    rs.Close

    Out "== relate to parameter"
    Set rs = Op(cn, "SHAPE {SELECT * FROM Customers ORDER BY CustID} APPEND ({SELECT * FROM Orders WHERE CustID = ? ORDER BY OrderID} AS rsP RELATE CustID TO PARAMETER 0)")
    Chk "param"
    Rows rs, "  ", 1
    rs.Close

    Out "== names"
    Set rs = Op(cn, "SHAPE {SELECT * FROM Customers ORDER BY CustID} APPEND ({SELECT * FROM Orders} RELATE CustID TO CustID)")
    Chk "no alias"
    Out "  last field=" & rs.Fields(rs.Fields.Count - 1).Name
    rs.Close
    Set rs = Op(cn, "SHAPE {SELECT * FROM Customers} AS C APPEND ({SELECT * FROM Orders} AS O RELATE CustID TO CustID) AS Kids")
    Chk "alias after"
    Out "  last field=" & rs.Fields(rs.Fields.Count - 1).Name
    rs.Close

    Out "== update through the parent"
    Set rs = New ADODB.Recordset
    rs.Open "SHAPE {SELECT * FROM Customers ORDER BY CustID} APPEND ({SELECT * FROM Orders ORDER BY OrderID} AS rsOrders RELATE CustID TO CustID)", cn, adOpenStatic, adLockOptimistic
    Chk "open optimistic"
    Out "locktype=" & rs.LockType
    rs!Name = "Alfreds Futterkiste"
    rs.Update
    Chk "parent update"
    Set ch = rs("rsOrders").Value
    ch!Amount = 101
    ch.Update
    Chk "child update"
    rs.Close
    Set rs = cn.Execute("SELECT Name FROM Customers WHERE CustID = 'ALFKI'")
    Out "  server name=" & V(rs!Name.Value)
    rs.Close
    Set rs = cn.Execute("SELECT Amount FROM Orders WHERE OrderID = 10")
    Out "  server amount=" & V(rs!Amount.Value)
    Chk "plain execute"
    rs.Close

    Out "== errors"
    Set rs = Op(cn, "SHAPE {SELECT * FROM Customers} APPEND")
    ChkAlways "syntax"
    Set rs = Op(cn, "SHAPE {SELECT * FROM Customers} APPEND ({SELECT * FROM Orders} AS x RELATE Nope TO CustID)")
    ChkAlways "bad field"
    Set rs = Op(cn, "SHAPE {SELECT * FROM Nope}")
    ChkAlways "bad table"
    Set rs = Op(cn, "SELECT Name FROM Customers ORDER BY Name")
    Chk "plain sql"
    Out "plain count=" & rs.RecordCount & " first=" & V(rs!Name.Value)
    rs.Close
    cn.Close
    Out "end"
    CloseOut
End Sub
