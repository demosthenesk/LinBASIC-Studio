Attribute VB_Name = "ModMain"
Private Const A As Long = B * 2
Private Const B As Long = 3
Sub Main()
    Log "ran " & A & " " & B
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
