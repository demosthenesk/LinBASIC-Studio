Attribute VB_Name = "ModMain"
Option Explicit
' A form that unloads itself in its Form_Load: "Object was unloaded" (364) to the code that loaded it
' (Load, Show, a property), not to the startup (VB6 oracle DEV-9/do2)

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\f_unload92.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Sub Main()
    Dim fm As Form
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_unload92.txt"
    Err.Clear
    Load Form1
    WriteLog "Load: " & Err.Number & " " & Err.Description
    Err.Clear
    Form1.Show
    WriteLog "Show: " & Err.Number
    Err.Clear
    Form1.Caption = "x"
    WriteLog "Caption: " & Err.Number
    Err.Clear
    Load Form2
    WriteLog "Form2: " & Err.Number & " " & Forms.Count
    For Each fm In Forms
        Unload fm
    Next
    WriteLog "count " & Forms.Count
End Sub
