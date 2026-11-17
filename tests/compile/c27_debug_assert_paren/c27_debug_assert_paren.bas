Attribute VB_Name = "ModMain"
Sub Main()
    Debug.Assert (5 Mod 4&) = 1
    Debug.Print (1 + 1) = 2
    Log "ran"
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
