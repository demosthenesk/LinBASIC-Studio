Attribute VB_Name = "Module1"
Option Explicit
Public Total As Long
Private Names(2) As String

Sub Main()
    Dim i As Integer, s As String
    Dim v As Variant
    s = "abc"
    For i = 1 To 3
        Total = Add(Total, i)
    Next
    v = Total * 2
    Debug.Print s; v
    Stop
    Debug.Print "after stop"; 1 / (i - 4)
End Sub

Function Add(ByVal a As Long, ByVal b As Long) As Long
    Dim r As Long
    r = a + b
    Add = r
End Function
