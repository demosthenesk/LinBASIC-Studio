Attribute VB_Name = "ModLog"
Option Explicit
' ByVal of a property of an intrinsic control to a Declare parameter As Any (LinBASIC Studio phase 7, robot:
' WizzForm Sender, Sound Association): VB6 binds Text1.Text early, a String goes (not a Variant: error 49)

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_byvalctl.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    WriteLog "== start"
    Form1.Show
End Sub
