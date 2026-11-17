Attribute VB_Name = "ModLog"
Option Explicit
Public gStep As Integer

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\f_basic.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Sub Main()
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_basic.txt"
    On Error GoTo 0
    WriteLog "main start"
    Form1.Show
    WriteLog "main end " & Forms.Count
End Sub
