Attribute VB_Name = "ModMain"
Sub Main()
    Dim x As Variant, o As Object, n As Long
    On Error Resume Next
    For Each x In o
        n = n + 1
        Log "body " & n & " err=" & Err.Number
        If n > 2 Then Exit For
    Next
    Log "ran " & n & " err=" & Err.Number
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
