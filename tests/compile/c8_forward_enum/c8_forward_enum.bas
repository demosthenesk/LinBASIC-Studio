Attribute VB_Name = "ModMain"
Private Const kLater As Long = wBoth * 10
Private Enum eWhen
    wBefore = 1
    wAfter = 2
    wBoth = wBefore Or wAfter
End Enum
Private Function G(Optional ByVal w As eWhen = eWhen.wAfter) As Long
    G = w
End Function
Sub Main()
    Log "ran " & G() & " " & wBoth & " " & kLater
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
