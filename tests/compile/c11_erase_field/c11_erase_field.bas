Attribute VB_Name = "ModMain"
Private Type TBox
    Name As String
    Items As Variant
End Type
Sub Main()
    Dim b(1 To 2) As TBox
    b(1).Items = Array(1, 2, 3)
    On Error Resume Next
    Erase b(1).Items
    Log "ran err=" & Err.Number & " " & IsArray(b(1).Items) & " " & IsEmpty(b(1).Items) & " " & TypeName(b(1).Items)
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
