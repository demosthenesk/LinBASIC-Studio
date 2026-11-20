Attribute VB_Name = "ModLog"
Option Explicit
Public fMain As MDIForm1

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\f_aw5.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Public Function FormNames() As String
    Dim f As Object, s As String
    For Each f In Forms
        s = s & " " & f.Name
    Next
    FormNames = Trim$(s)
End Function

Public Sub Main()
    Dim r As Double
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_aw5.txt"
    Err.Clear
    r = Shell("lbaw5_no_such_program", vbHide)
    WriteLog "shell missing " & Err.Number & " " & r
    Err.Clear
    r = Shell("lbaw5_no_such_program /x", vbNormalFocus)
    WriteLog "shell missing args " & Err.Number & " " & r
    Err.Clear
    On Error GoTo 0
    ' the Application Wizard: Set fMainForm = New frmMain; Load fMainForm; ... fMainForm.Show
    Set fMain = New MDIForm1
    Load fMain
    WriteLog "mdi loaded " & Forms.Count & " " & FormNames()
    fMain.Show
    frmChild.Show
    WriteLog "child shown " & Forms.Count & " " & FormNames() & " " & frmChild.MDIChild
    WriteLog "after " & Forms.Count & " " & FormNames()
    Form1.Show
End Sub
