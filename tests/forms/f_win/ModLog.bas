Attribute VB_Name = "ModLog"
Option Explicit
Public gChildCount As Integer

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\f_win.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Public Sub StartLog()
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_win.txt"
End Sub

Public Sub LogBorders(f As Object, ByVal name As String)
    WriteLog name & " " & (f.Width - f.ScaleWidth) & " " & (f.Height - f.ScaleHeight) & " " & f.BorderStyle & " " & f.ControlBox & " " & f.MinButton & " " & f.MaxButton & " " & f.ShowInTaskbar & " " & f.WindowState & " " & f.Moveable
End Sub
