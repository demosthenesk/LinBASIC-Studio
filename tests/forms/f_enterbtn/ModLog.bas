Attribute VB_Name = "ModLog"
Option Explicit

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\f_enterbtn.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
