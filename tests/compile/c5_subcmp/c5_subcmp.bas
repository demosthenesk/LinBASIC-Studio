Attribute VB_Name = "ModMain"
Private Sub Chk(ByVal b As Boolean)
    Log "chk " & b
End Sub
Sub Main()
    Chk (1 + 1) = 2
    Chk (1 + 1) = 3
    Debug.Assert (5 Mod 4&) = 1
    Log "ran"
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
