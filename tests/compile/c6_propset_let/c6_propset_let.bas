Attribute VB_Name = "ModMain"
Sub Main()
    Dim o As New CP
    On Error Resume Next
    o.Fnt = "Courier New"
    Log "ran err=" & Err.Number & " name=" & o.Fnt.Name
    Err.Clear
    o.Inner
    Log "inner err=" & Err.Number & " name=" & o.Fnt.Name
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
