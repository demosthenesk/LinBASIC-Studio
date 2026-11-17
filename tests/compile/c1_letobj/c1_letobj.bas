Attribute VB_Name = "ModMain"
Private Function F() As Collection
    F = New Collection
End Function
Sub Main()
    Dim c As Collection
    On Error Resume Next
    Set c = F()
    Log "ran err=" & Err.Number & " nothing=" & (c Is Nothing)
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
