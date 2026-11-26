Attribute VB_Name = "ModLog"
Option Explicit

' f_inetasync (10.4): UserControl.AsyncRead / CancelAsyncRead, the AsyncReadProgress and AsyncReadComplete
' events, the AsyncProperty object, UserControl.Hyperlink; against tests/dev/inet_server.py (HTTP 47120).
' server.txt next to the program: the address of the server, else this computer.
Public gHost As String

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_inetasync.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, Mask(s)
    Close #f
End Sub

' every progress event (their number changes from run to run): not compared
Public Sub WriteRaw(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\f_inetasync-raw.txt" For Append As #f
    Print #f, Mask(s)
    Close #f
End Sub

Public Function Mask(ByVal s As String) As String
    s = Replace(s, App.Path, "<dir>", , , vbTextCompare)
    s = Replace(s, Replace(App.Path, "\", "/"), "<dir>", , , vbTextCompare)
    If gHost <> "" Then s = Replace(s, gHost, "<server>")
    Mask = Replace(s, "\", "/")                ' the separators of the paths: as on Linux

End Function

Sub Main()
    Dim f As Integer
    gHost = "127.0.0.1"
    If Dir(App.Path & "\server.txt") <> "" Then
        f = FreeFile
        Open App.Path & "\server.txt" For Input As #f
        Line Input #f, gHost
        Close #f
        gHost = Trim$(gHost)
    End If
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    If Dir(App.Path & "\expected\f_inetasync-raw.txt") <> "" Then Kill App.Path & "\expected\f_inetasync-raw.txt"
    Form1.Show
End Sub
