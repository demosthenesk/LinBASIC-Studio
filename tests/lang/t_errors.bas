Attribute VB_Name = "ModErrors"
Option Explicit
Private Sub Fails()
    Dim x As Integer
    x = 10 / 0
End Sub

Private Function SafeDiv(a As Double, b As Double) As Variant
    On Error GoTo bad
    SafeDiv = a / b
    Exit Function
bad:
    SafeDiv = "error " & Err.Number
End Function

Public Sub T_Errors(f As Integer)
    Dim i As Integer
    On Error Resume Next
    i = 1 / 0
    Print #f, "div0"; Err.Number; Err.Description
    Err.Clear
    Print #f, "cleared"; Err.Number
    i = 40000
    Print #f, "overflow"; Err.Number
    Err.Clear
    i = CInt("abc")
    Print #f, "mismatch"; Err.Number
    Err.Clear
    Fails
    Print #f, "from sub"; Err.Number
    Err.Clear
    Err.Raise 1234, "MySource", "My message"
    Print #f, "raise"; Err.Number; " "; Err.Source; " "; Err.Description
    Err.Clear
    Err.Raise vbObjectError + 5
    Print #f, "object error"; Err.Number - vbObjectError
    On Error GoTo 0
    Print #f, SafeDiv(1, 4); " "; SafeDiv(1, 0); " "; SafeDiv(0, 0)
    On Error GoTo handler
    i = 0
    Error 5
    Print #f, "resumed"; i
    Exit Sub
handler:
    i = i + 1
    Print #f, "handler"; Err.Number; Erl
    If i < 3 Then Resume
    Resume Next
End Sub
