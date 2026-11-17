Attribute VB_Name = "CliMain"
Option Explicit
' SingleUse, MultiUse and GlobalSingleUse objects of an ActiveX EXE (LinBASIC Studio, phase 4.3);
' the server is made (File > Make SUServer.exe) before the client runs: late bound, no reference

Private Function LogPath() As String
    LogPath = App.Path & "\..\expected\c_single.txt"
End Function

Private Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

Sub Main()
    Dim s1 As Object, s2 As Object, m1 As Object, m2 As Object, m3 As Object, g As Object
    If Dir(App.Path & "\..\expected", vbDirectory) = "" Then MkDir App.Path & "\..\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    Set m1 = CreateObject("SUServer.CMulti")
    Set m2 = CreateObject("SUServer.CMulti")
    m1.Add 5
    WriteLog m1.Info
    WriteLog m2.Info
    Set s1 = CreateObject("SUServer.CSingle")
    Set s2 = CreateObject("SUServer.CSingle")
    s1.Add 7
    WriteLog s1.Info
    WriteLog s2.Info
    s2.Add 100
    WriteLog s1.Info
    WriteLog s2.Info
    WriteLog m1.Info
    Set m3 = s1.MakeMulti()
    m3.Add 1
    WriteLog m3.Info
    WriteLog s1.Info
    WriteLog m2.Info
    Set g = CreateObject("SUServer.CGlob")
    WriteLog g.GInfo
    WriteLog g.GInfo
    WriteLog TypeName(s1) & " " & TypeName(m1) & " " & TypeName(g)
    Set s1 = Nothing
    Set s2 = Nothing
    Set s1 = CreateObject("SUServer.CSingle")
    WriteLog s1.Info
    WriteLog m1.Info
    WriteLog "end"
End Sub
