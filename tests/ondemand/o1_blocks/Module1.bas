Attribute VB_Name = "Module1"
Option Explicit
' Compile On Demand (a run of the IDE, F5): a procedure with a block error (an If without End If, End Sub inside
' an If, a Dim between Select Case and its first Case, a For without Next, a Sub without End Sub) is compiled only
' when it is called, so the program runs when it does not call them (robot 2026-11-12: 18 projects of PSC).

Private f As Integer

Private Sub WriteLog(ByVal s As String)
    Print #f, s
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    f = FreeFile
    Open App.Path & "\expected\o1_blocks.txt" For Output As #f
    WriteLog "start"
    Good1
    WriteLog "after good1: " & Twice(21)
    Good2
    WriteLog "end"
    Close #f
End Sub

Private Sub NoEndIf(ByVal x As Long)
    If x > 0 Then
        WriteLog "positive"
    WriteLog "never"
End Sub

Private Sub Good1()
    WriteLog "good1"
End Sub

Private Sub BadSelect(ByVal k As Long)
    Select Case k
    Dim q As Long
    Case 1
        q = 1
    End Select
End Sub

Private Function NoNext() As Long
    Dim i As Long
    For i = 1 To 3
        NoNext = NoNext + i
End Function

Private Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function

Private Sub NoEndSub()
    WriteLog "never either"

Private Sub Good2()
    WriteLog "good2"
End Sub
