Attribute VB_Name = "ModOut"
Option Explicit

Public gF As Integer

Public Sub OpenOut(ByVal test As String)
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    gF = FreeFile
    Open App.Path & "\expected\" & test & ".txt" For Output As #gF
End Sub

Public Sub CloseOut()
    Close #gF
End Sub

Public Sub Out(ByVal s As String)
    Print #gF, s
End Sub
