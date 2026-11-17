Attribute VB_Name = "ModMain"
Sub Main()
    Dim a() As Variant
    On Error Resume Next
    a = Split("x,y", ",")
    Log "ran err=" & Err.Number & " " & UBound(a)
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
