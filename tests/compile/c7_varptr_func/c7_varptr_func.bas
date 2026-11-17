Attribute VB_Name = "ModMain"
Private Function S() As String
    S = "abc"
End Function
Sub Main()
    Dim p As Long
    p = VarPtr(S())
    Log "ran " & (p <> 0)
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
