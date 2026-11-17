Attribute VB_Name = "ModRpt2"
Option Explicit

Private Sub MakeDb()
    Dim e As Object, db As Object, p As String, i As Integer, w As String
    p = App.Path & "\rpt2.mdb"
    If Dir(p) <> "" Then Kill p
    Set e = CreateObject("DAO.DBEngine.36")
    Set db = e.CreateDatabase(p, ";LANGID=0x0409;CP=1252;COUNTRY=0", 64)
    db.Execute "CREATE TABLE Nums (N LONG CONSTRAINT PrimaryKey PRIMARY KEY, Word TEXT(40))"
    For i = 1 To 75
        w = String$(1 + (i Mod 9), Chr$(65 + (i Mod 26))) & " w" & i
        db.Execute "INSERT INTO Nums VALUES (" & i & ", '" & w & "')"
    Next
    db.Close
End Sub

Private Sub DumpFile(ByVal p As String)
    Dim f As Integer, s As String, i As Long, c As String, t As String
    If Dir(p) = "" Then Out "  (no file " & p & ")": Exit Sub
    f = FreeFile
    Open p For Binary As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    Out "  bytes=" & Len(s)
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
    Do While rptTwo.AsyncCount > 0
        DoEvents
        If Timer - t > 20 Then Out "  (timeout)": Exit Do
    Loop
End Sub

Sub Main()
    Dim p As String
    OpenOut "d_rpt2"
    On Error Resume Next
    MakeDb
    Chk "makedb"
    Out "screen tppx=" & Screen.TwipsPerPixelX
    Out "== text"
    p = App.Path & "\two.txt"
    If Dir(p) <> "" Then Kill p
    rptTwo.ExportReport rptKeyText, p, True, False
    Chk "export text"
    Wait
    DumpFile p
    Out "== html"
    p = App.Path & "\two.htm"
    If Dir(p) <> "" Then Kill p
    rptTwo.ExportReport rptKeyHTML, p, True, False
    Chk "export html"
    Wait
    DumpFile p
    Out "== pages 2 to 2"
    p = App.Path & "\two2.txt"
    If Dir(p) <> "" Then Kill p
    rptTwo.ExportReport rptKeyText, p, True, False, rptRangeFromTo, 2, 2
    Chk "export range"
    Wait
    DumpFile p
    Unload rptTwo
    Out "end"
    CloseOut
End Sub
