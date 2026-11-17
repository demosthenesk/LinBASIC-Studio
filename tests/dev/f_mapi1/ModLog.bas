Attribute VB_Name = "ModLog"
Option Explicit
' MAPI 1 (LinBASIC Studio phase 7.9): the log of the test

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_mapi1.txt"
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

' a value with its type
Public Function ShowV(v As Variant) As String
    On Error GoTo bad
    If IsObject(v) Then ShowV = "(" & TypeName(v) & ")": Exit Function
    If IsArray(v) Then ShowV = "(" & TypeName(v) & ")": Exit Function
    ShowV = TypeName(v) & ":" & v
    Exit Function
bad:
    ShowV = TypeName(v) & ":?" & Err.Number
End Function
