Attribute VB_Name = "ModLog"
Option Explicit

Public gRunning As Boolean      ' False while the IDE runs the user control at design time
Public gSerial As Long

Private Function LogPath() As String
    LogPath = App.Path & "\expected\f_uctl.txt"
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
    WriteLog "main"
    Form1.Show
    WriteLog "main shown"
End Sub
