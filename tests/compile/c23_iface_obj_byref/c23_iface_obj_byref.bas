Attribute VB_Name = "ModMain"
Sub Main()
    Dim i As IFace, c As New Collection
    c.Add 1
    Set i = New CImpl
    Set i.O = c
    Log "o " & i.O.Count
    Log "ran"
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
