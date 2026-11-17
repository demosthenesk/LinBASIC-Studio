Attribute VB_Name = "ModMain"
Sub Main()
    Dim i As IFace, c As New Collection
    Set i = New CImpl
    i.V = 5
    Log "v " & i.V
    Set i.V = c
    Log "obj " & TypeName(i.V)
    Log "ran"
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
