Attribute VB_Name = "ModMain"
Sub Main()
    Dim o As New CP
    Log "ran " & o.Value
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
