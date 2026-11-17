Attribute VB_Name = "ModMain"
Option Explicit
Sub Main()
    Log "ran"
End Sub
Private Sub NeverCalled()
    undefinedThing = 1
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
