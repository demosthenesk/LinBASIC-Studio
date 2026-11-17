Attribute VB_Name = "ModLog"
Option Explicit

Public gRunning As Boolean      ' False while the IDE runs the user controls at design time

Private Function LogPath() As String
    LogPath = App.Path & "\expected\f_uctl2.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    If Not gRunning Then Exit Sub
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

Sub Main()
    gRunning = True
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    Form1.Show
End Sub
