Attribute VB_Name = "ModFlow"
Option Explicit

Private Function SumAll(ParamArray nums() As Variant) As Double
    Dim i As Long
    For i = LBound(nums) To UBound(nums)
        SumAll = SumAll + nums(i)
    Next
End Function

Private Function CountArgs(ByVal prefix As String, ParamArray rest() As Variant) As String
    Dim v As Variant, s As String
    s = prefix & UBound(rest) & ":"
    For Each v In rest
        s = s & TypeName(v) & ","
    Next
    CountArgs = s
End Function

Private Function Classify(ByVal n As Integer) As String
    On n GoTo one, two, three
    Classify = "other"
    Exit Function
one:
    Classify = "one": Exit Function
two:
    Classify = "two": Exit Function
three:
    Classify = "three"
End Function

Public Sub T_Flow(f As Integer)
    Dim s As String, i As Integer
    s = ""
    GoSub addA
    GoSub addB
    GoSub addA
    Print #f, s
    For i = 0 To 4
        Print #f, Classify(i); " ";
    Next
    Print #f,
    s = ""
    For i = 1 To 3
        On i GoSub g1, g2, g3
    Next
    Print #f, s
    Print #f, SumAll(1, 2, 3.5); SumAll(); CountArgs("p", 1, "x", 2.5, True); " "; CountArgs("q")
    Print #f, "abc" Like "a*"; "abc" Like "a?c"; "a1c" Like "a#c"; "abc" Like "a[b-d]c"; "aec" Like "a[!b-d]c"
    Print #f, "ABC" Like "abc"; "a*c" Like "a[*]c"; "" Like "*"; "abc" Like "ab"; "x" Like "[]"; "Ab" Like "[A-Z][a-z]"
    On Error Resume Next
    Return
    Print #f, "err"; Err.Number
    Err.Clear
    i = -1
    On i GoTo g1
    Print #f, "err"; Err.Number
    Exit Sub
addA:
    s = s & "A"
    Return
addB:
    s = s & "B"
    GoSub addA
    Return
g1:
    s = s & "1": Return
g2:
    s = s & "2": Return
g3:
    s = s & "3": Return
End Sub
