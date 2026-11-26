Attribute VB_Name = "ModLog"
Option Explicit

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\f_help.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Public Sub StartLog()
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_help.txt"
End Sub

Public Function ShortName(ByVal s As String) As String
    ShortName = Mid$(s, InStrRev(s, "\") + 1)
End Function
