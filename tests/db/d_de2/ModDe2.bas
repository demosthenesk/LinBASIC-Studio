Attribute VB_Name = "ModDe2"
Option Explicit

Private Function MakeDb() As String
    Dim e As Object, db As Object, p As String
    p = App.Path & "\de2.mdb"
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
    Dim c As ADODB.Command, p As ADODB.Parameter
    Set c = deDe2.Commands(name)
    Chk "command " & name
    Out "command " & name & " type=" & c.CommandType & " params=" & c.Parameters.Count & " text=" & c.CommandText
    Chk "command text " & name
    For Each p In c.Parameters
        Out "  param " & p.Name & " type=" & p.Type & " dir=" & p.Direction & " size=" & p.Size & " prec=" & p.Precision & " scale=" & p.NumericScale & " attr=" & p.Attributes & " value=" & V(p.Value)
    Next
    Chk "params " & name
End Sub

' a file the report wrote, line by line
Private Sub DumpFile(ByVal p As String)
    Dim f As Integer, s As String, i As Long, c As String, t As String
    If Dir(p) = "" Then Out "  (no file " & p & ")": Exit Sub
    f = FreeFile
    Open p For Binary As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    t = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = vbLf Then
            Out "  |" & t
            t = ""
        ElseIf c = vbCr Then
            t = t & "<cr>"
        ElseIf Asc(c) < 32 Then
            t = t & "<" & Asc(c) & ">"
        Else
            t = t & c
        End If
    Next
    If t <> "" Then Out "  |" & t
End Sub

Private Sub Wait()
    Dim t As Single
    t = Timer
    Do While rptDe2.AsyncCount > 0
        DoEvents
        If Timer - t > 20 Then Out "  (timeout)": Exit Do
    Loop
End Sub

Sub Main()
    Dim i As Integer, j As Integer, s As String, p As String, rs As ADODB.Recordset, o As Object
    OpenOut "d_de2"
    On Error Resume Next
    MakeDb
    Chk "makedb"
    Out "== the commands"
    Out "commands=" & deDe2.Commands.Count & " recordsets=" & deDe2.Recordsets.Count
    ShowCmd "ByCountry"
    ShowCmd "Range"
    ShowCmd "Cust"

    Out "== a command with a parameter"
    deDe2.ByCountry "Germany"
    Chk "ByCountry Germany"
    Set rs = deDe2.rsByCountry
    Out "state=" & rs.State & " count=" & rs.RecordCount & " source=" & rs.Source
    ShowRs rs, "  ", 0
    Out "param after=" & V(deDe2.Commands("ByCountry").Parameters(0).Value)
    rs.Close
    deDe2.ByCountry "France"
    Chk "ByCountry France"
    ShowRs deDe2.rsByCountry, "  ", 0
    deDe2.rsByCountry.Close
    Out "-- the parameter set first, then Execute"
    deDe2.Commands("ByCountry").Parameters("Param1").Value = "Netherlands"
    Chk "set param"
    Set rs = deDe2.Commands("ByCountry").Execute
    Chk "execute"
    ShowRs rs, "  ", 0
    rs.Close

    Out "== two parameters"
    deDe2.Range 10, 200
    Chk "Range 10 200"
    ShowRs deDe2.rsRange, "  ", 0
    deDe2.rsRange.Close
    deDe2.Range "250", "1000"
    Chk "Range strings"
    ShowRs deDe2.rsRange, "  ", 0
    deDe2.rsRange.Close
    deDe2.Range 1, "x"
    Chk "Range x"
    Out "state=" & deDe2.rsRange.State

    Out "== a child with a parameter"
    deDe2.Cust
    Chk "Cust"
    Set rs = deDe2.rsCust
    Out "state=" & rs.State & " count=" & rs.RecordCount
    Out "source=" & rs.Source
    ShowRs rs, "  ", 1
    rs.Close

    Out "== the report"
    Out "typename=" & TypeName(rptDe2) & " datamember=" & rptDe2.DataMember & " sections=" & rptDe2.Sections.Count
    Chk "report"
    For i = 1 To rptDe2.Sections.Count
        With rptDe2.Sections(i)
            s = "  section " & i & " " & .Name & " height=" & .Height & " controls=" & .Controls.Count & ":"
            For j = 1 To .Controls.Count
                s = s & " " & TypeName(.Controls(j)) & ":" & .Controls(j).Name
            Next
            Out s
        End With
    Next
    Chk "sections"
    Set o = rptDe2.Sections("Cust_Header").Controls("shpBox")
    Out "shpBox shape=" & o.Shape & " borderstyle=" & o.BorderStyle & " backstyle=" & o.BackStyle & " left=" & o.Left & " width=" & o.Width & " height=" & o.Height
    Chk "shpBox"
    Set o = rptDe2.Sections("Cust_Header").Controls("imgLogo")
    Out "imgLogo sizemode=" & o.SizeMode & " align=" & o.PictureAlignment & " picture=" & TypeName(o.Picture) & " left=" & o.Left & " top=" & o.Top
    Chk "imgLogo"
    Out "shpDot shape=" & rptDe2.Sections("CustOrders_Detail").Controls("shpDot").Shape
    Out "shpOval shape=" & rptDe2.Sections("Cust_Footer").Controls("shpOval").Shape
    Chk "shapes"
    p = App.Path & "\out.txt"
    If Dir(p) <> "" Then Kill p
    rptDe2.ExportReport rptKeyText, p, True, False
    Chk "export text"
    Wait
    DumpFile p
    Unload rptDe2
    Chk "unload"
    Out "end"
    CloseOut
End Sub
