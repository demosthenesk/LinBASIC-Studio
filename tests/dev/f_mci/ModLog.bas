Attribute VB_Name = "ModLog"
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

' the number of MCI devices the program has open
Public Function OpenDevices() As String
    Dim b As String
    b = String$(64, 0)
    mciSendString "sysinfo all quantity open", b, 64, 0
    OpenDevices = Left$(b, InStr(b & Chr$(0), Chr$(0)) - 1)
End Function

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_mci.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, Replace(s, App.Path, "<dir>", , , vbTextCompare)
    Close #f
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    Form1.Show
End Sub
