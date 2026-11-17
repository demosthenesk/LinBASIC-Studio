Attribute VB_Name = "ModResNext"
Option Explicit
' Phase 6.8: On Error Resume Next and an error in the condition of If, ElseIf, a one-line If, Do While/Until,
' Loop While/Until and While (robot: Jumble and SkinDemo, Loop Until EOF(n) of a file that did not open).
' Every loop stops itself after 3 rounds, whatever the answer.

Private Function Bad() As Long
    Err.Raise 5
End Function

Public Sub T_ResNext(f As Integer)
    Dim n As Long, z As Long
    On Error Resume Next
    z = 0
    If 1 / z > 0 Then
        Print #f, "if: then"; Err.Number
    Else
        Print #f, "if: else"; Err.Number
    End If
    Print #f, "if: after"; Err.Number
    Err.Clear
    If z = 1 Then
        Print #f, "elseif: if"
    ElseIf Bad() = 1 Then
        Print #f, "elseif: elseif"; Err.Number
    Else
        Print #f, "elseif: else"; Err.Number
    End If
    Err.Clear
    If Bad() Then Print #f, "line if: then"; Err.Number Else Print #f, "line if: else"; Err.Number
    Err.Clear

    n = 0
    Do While Bad() = 0
        n = n + 1
        If n >= 3 Then Exit Do
    Loop
    Print #f, "do while:"; n; Err.Number
    Err.Clear
    n = 0
    Do Until Bad() = 0
        n = n + 1
        If n >= 3 Then Exit Do
    Loop
    Print #f, "do until:"; n; Err.Number
    Err.Clear
    n = 0
    Do
        n = n + 1
        If n >= 3 Then Exit Do
    Loop Until Bad() = 1
    Print #f, "loop until:"; n; Err.Number
    Err.Clear
    n = 0
    Do
        n = n + 1
        If n >= 3 Then Exit Do
    Loop While Bad() = 0
    Print #f, "loop while:"; n; Err.Number
    Err.Clear
    n = 0
    While Bad() = 0
        n = n + 1
        If n >= 3 Then GoTo outw
    Wend
outw:
    Print #f, "while:"; n; Err.Number
    Err.Clear
    n = 0
    Do
        n = n + 1
        If n >= 3 Then Exit Do
    Loop Until EOF(99)
    Print #f, "loop until eof:"; n; Err.Number
    Err.Clear
End Sub
