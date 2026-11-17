Attribute VB_Name = "Module1"
Option Explicit
' A member that the Err object does not have (Err.numer, a typo in the XChart.ctl of ActiveChart): VB6 compiles it
' (Compile On Demand too) and it is a run-time error, trapped by On Error Resume Next (robot udt2: 6 projects)

Private Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\o2_errmember.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(App.Path & "\expected\o2_errmember.txt") <> "" Then Kill App.Path & "\expected\o2_errmember.txt"
    WriteLog "start"
    Good
    On Error Resume Next
    Typo
    WriteLog "after typo " & Err.Number
    On Error GoTo 0
    WriteLog "end"
End Sub

Private Sub Good()
    On Error Resume Next
    Err.Raise 5
    WriteLog "good " & Err.Number
    Err.Clear
End Sub

Private Sub Typo()
    On Error Resume Next
    WriteLog "typo start"
    Dim v As Variant
    If Err.numer Then Err.Clear
    WriteLog "typo end " & Err.Number
    v = Err.numer
    WriteLog "value " & Err.Number & " " & Err.Description & " [" & v & "]"
End Sub
