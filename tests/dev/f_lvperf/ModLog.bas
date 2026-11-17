Attribute VB_Name = "ModLog"
Option Explicit
' ListView speed (LinBASIC Studio phase 7, robot: an IP subnet calculator that adds 131,072 items): the results
' go to expected\f_lvperf.txt (compared with VB6), the times to expected\timing-f_lvperf.txt (only to read)

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_lvperf.txt"
End Function

Public Function TimePath() As String
    TimePath = App.Path & "\expected\timing-f_lvperf.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

Public Sub WriteTime(ByVal what As String, ByVal secs As Single)
    Dim f As Integer
    f = FreeFile
    Open TimePath For Append As #f
    Print #f, what & ": " & Format$(secs, "0.00") & " s"
    Close #f
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    If Dir(TimePath) <> "" Then Kill TimePath
    WriteLog "== start"
    Form1.Show
End Sub
