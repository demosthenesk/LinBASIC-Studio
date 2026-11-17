Attribute VB_Name = "ModControl"
Option Explicit
Public Sub T_Control(f As Integer)
    Dim i As Integer, j As Integer, n As Long, s As String
    For i = 1 To 3
        If i = 1 Then
            Print #f, "one";
        ElseIf i = 2 Then
            Print #f, "two";
        Else
            Print #f, "other";
        End If
    Next i
    Print #f,
    For i = 10 To 1 Step -3: Print #f, i;: Next
    Print #f, "after"; i
    Dim x As Double
    For x = 0 To 1 Step 0.25: Print #f, x;: Next x
    Print #f,
    For i = 5 To 1: Print #f, "never": Next
    Print #f, "empty loop"; i
    n = 0
    Do While n < 5: n = n + 1: Loop
    Print #f, "do while"; n
    Do: n = n - 2: Loop Until n < 0
    Print #f, "loop until"; n
    Do Until n >= 3
        n = n + 1
        If n = 2 Then Exit Do
    Loop
    Print #f, "exit do"; n
    While n < 10: n = n + 3: Wend
    Print #f, "wend"; n
    For i = 1 To 3
        For j = 1 To 3
            If j = 2 Then Exit For
            Print #f, i * 10 + j;
        Next j
    Next i
    Print #f,
    s = "beta"
    Select Case s
    Case "alpha": Print #f, "a"
    Case "beta", "gamma": Print #f, "b or g"
    Case Else: Print #f, "else"
    End Select
    For i = -1 To 12 Step 4
        Select Case i
        Case Is < 0: Print #f, "neg";
        Case 0 To 4: Print #f, "small";
        Case 5, 7, 9: Print #f, "odd";
        Case Else: Print #f, "big";
        End Select
    Next
    Print #f,
    i = 0
again:
    i = i + 1
    If i < 3 Then GoTo again
    Print #f, "goto"; i
    If i = 3 Then Print #f, "single line": Print #f, "two statements" Else Print #f, "no"
    If i <> 3 Then Print #f, "no" Else Print #f, "else part"
End Sub
