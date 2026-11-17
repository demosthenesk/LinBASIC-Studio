Attribute VB_Name = "ModMain"
Sub Main()
    Dim i As IFace
    Set i = New CImpl
    i.S = "abc"
    Log "s " & i.S
    Log "ran"
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
