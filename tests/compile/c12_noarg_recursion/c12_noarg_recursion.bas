Attribute VB_Name = "ModMain"
Private mDepth As Long
Private Function Spin() As Long
    mDepth = mDepth + 1
    If mDepth < 4 Then Spin
    Spin = mDepth
End Function
Sub Main()
    Dim r As Long
    r = Spin()
    Log "ran " & r & " " & mDepth
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
