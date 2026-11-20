Attribute VB_Name = "ModMain"
Option Explicit

Sub Main()
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Output As #f
    Print #f, "console ok"
    Close #f
End Sub
