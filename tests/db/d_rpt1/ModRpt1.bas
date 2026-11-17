Attribute VB_Name = "ModRpt1"
Option Explicit

Private Function MakeDb() As String
    Dim e As Object, db As Object, p As String
    p = App.Path & "\rpt1.mdb"
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
    db.Execute "INSERT INTO Orders VALUES (11, 'ALFKI', 1020, #2026-02-01#)"
    db.Execute "INSERT INTO Orders VALUES (12, 'BONAP', 7.25, #2026-02-03#)"
    db.Execute "INSERT INTO Orders VALUES (13, 'BLAUS', 300, #2026-03-01#)"
    db.Execute "INSERT INTO Orders VALUES (14, 'ALFKI', 1, #2026-03-09#)"
    db.Close
    MakeDb = p
End Function

' a file the report wrote, line by line, with what cannot be seen shown
Private Sub DumpFile(ByVal p As String)
    Dim f As Integer, s As String, i As Long, c As String, t As String, n As Long
    If Dir(p) = "" Then Out "  (no file " & p & ")": Exit Sub
    f = FreeFile
    Open p For Binary As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    Out "  bytes=" & Len(s)
    t = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = vbLf Then
            Out "  |" & t
            t = ""
            n = n + 1
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
    Do While rptTest.AsyncCount > 0
        DoEvents
        If Timer - t > 20 Then Out "  (timeout)": Exit Do
    Loop
End Sub

Sub Main()
    Dim i As Integer, j As Integer, s As String, p As String
    OpenOut "d_rpt1"
    On Error Resume Next
    MakeDb
    Chk "makedb"
    Out "== the object"
    Out "typename=" & TypeName(rptTest) & " caption=" & rptTest.Caption & " title=" & rptTest.Title
    Out "margins=" & rptTest.LeftMargin & "," & rptTest.TopMargin & "," & rptTest.RightMargin & "," & rptTest.BottomMargin & " width=" & rptTest.ReportWidth
    Out "datamember=" & rptTest.DataMember & " datasource=" & TypeName(rptTest.DataSource)
    Chk "props"
    Out "sections=" & rptTest.Sections.Count
    For i = 1 To rptTest.Sections.Count
        With rptTest.Sections(i)
            s = "  section " & i & " " & .Name & " height=" & .Height & " controls=" & .Controls.Count & ":"
            For j = 1 To .Controls.Count
                s = s & " " & TypeName(.Controls(j)) & ":" & .Controls(j).Name
            Next
            Out s
        End With
    Next
    Chk "sections"
    Out "detail by name=" & rptTest.Sections("Orders_Detail").Name
    Out "label=" & rptTest.Sections("PageFooter").Controls("lblPage").Caption
    Chk "by name"
    Out "exportformats=" & rptTest.ExportFormats.Count
    For i = 1 To rptTest.ExportFormats.Count
        With rptTest.ExportFormats(i)
            Out "  format " & i & " key=" & .Key & " type=" & .FormatType & " string=" & .FileFormatString & " filter=" & .FileFilter
            Out "  template " & Replace(Replace(.Template, vbCr, "<cr>"), vbLf, "<lf>")
        End With
    Next
    Chk "formats"

    Out "== export text"
    p = App.Path & "\out.txt"
    If Dir(p) <> "" Then Kill p
    rptTest.ExportReport rptKeyText, p, True, False
    Chk "export text"
    Wait
    DumpFile p

    Out "== export html"
    p = App.Path & "\out.htm"
    If Dir(p) <> "" Then Kill p
    rptTest.ExportReport rptKeyHTML, p, True, False
    Chk "export html"
    Wait
    DumpFile p

    Out "== export again without overwrite"
    rptTest.ExportReport rptKeyText, App.Path & "\out.txt", False, False
    Chk "no overwrite"
    Wait

    Out "== a format of the program"
    rptTest.ExportFormats.Add "mine", rptFmtText, "Mine (*.lst)", "*.lst", ""
    Chk "add format"
    Out "exportformats=" & rptTest.ExportFormats.Count & " template=" & Replace(Replace(rptTest.ExportFormats("mine").Template, vbCr, "<cr>"), vbLf, "<lf>")
    rptTest.ExportFormats.Add "mine", rptFmtText, "Mine (*.lst)", "*.lst", ""
    Chk "add again"
    p = App.Path & "\out.lst"
    If Dir(p) <> "" Then Kill p
    rptTest.ExportReport "mine", p, True, False, rptRangeFromTo, 1, 1
    Chk "export mine"
    Wait
    DumpFile p
    Out "== titles"
    rptTest.Title = "My title"
    rptTest.Sections("ReportHeader").Controls("lblTitle").Caption = "Title: %i"
    Chk "set title"
    p = App.Path & "\out2.htm"
    If Dir(p) <> "" Then Kill p
    rptTest.ExportReport rptKeyHTML, p, True, False
    Wait
    DumpFile p
    Unload rptTest
    Chk "unload"
    Out "end"
    CloseOut
End Sub
