Attribute VB_Name = "ModMain"
Private Function G(Optional ByVal w As Long = wAfter) As Long
    G = w
End Function
Private Enum eWhen
    wBefore = 1
    wAfter = 2
End Enum
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
