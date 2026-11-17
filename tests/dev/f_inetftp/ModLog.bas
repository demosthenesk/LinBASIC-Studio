Attribute VB_Name = "ModLog"
Option Explicit

' the address of the test server (tests/dev/inet_server.py): server.txt next to the program, else this computer
Public gHost As String

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_inetftp.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, Mask(s)
    Close #f
End Sub

' the raw order of the StateChanged events (it changes from run to run): not compared
Public Sub WriteRaw(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\f_inetftp-raw.txt" For Append As #f
    Print #f, Mask(s)
    Close #f
End Sub

Public Function Mask(ByVal s As String) As String
    s = Replace(s, App.Path, "<dir>", , , vbTextCompare)
    If gHost <> "" Then s = Replace(s, gHost, "<server>")
    Mask = s
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
    If Dir(App.Path & "\expected\f_inetftp-raw.txt") <> "" Then Kill App.Path & "\expected\f_inetftp-raw.txt"
    Form1.Show
End Sub
