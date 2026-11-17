Attribute VB_Name = "Module2"
Option Explicit
Public Count As Long

Sub Main()
    Dim i As Long, x As Long, c As Class1
    x = 1
    x = x + 10
    x = x + 100
    Debug.Print "x"; x
    For i = 1 To 5
        Count = Count + 1
    Next
    On Error Resume Next
    x = 1 / 0
    On Error GoTo 0
    Debug.Print "count"; Count
    Set c = New Class1
    c.Oops
    Debug.Print "end"
End Sub
