Attribute VB_Name = "ModMain"
Private Const K As Long = CLng(vbBlue)
Private Const K2 As Integer = CInt(2.5) + CInt("3")
Sub Main()
    Log "ran " & K & " " & K2
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
