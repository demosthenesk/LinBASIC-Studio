Attribute VB_Name = "ModMain"
Private Function G(Optional ByVal c As Long = CLng(vbBlue)) As Long
    G = c
End Function
Sub Main()
    Log "ran " & G()
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
