Attribute VB_Name = "ModNumFmt"
Option Explicit

Private Function TryD(ByVal s As String) As String
    On Error GoTo bad
    TryD = CStr(CDbl(s))
    Exit Function
bad:
    TryD = "E" & Err.Number
End Function

Private Function TryC(ByVal s As String) As String
    On Error GoTo bad
    TryC = CStr(CCur(s))
    Exit Function
bad:
    TryC = "E" & Err.Number
End Function

Public Sub T_NumFmt(f As Integer)
    Dim d As Double, s As Single, i As Integer
    ' small numbers: where does the E-form start? (Double, 15 digits)
    d = 1
    For i = 1 To 22
        d = d / 10
        Print #f, "d"; i; CStr(d); "|"; Str(d); "|"; CStr(d * 1.5); "|"; CStr(d * 1.23456789); "|"; CStr(d * 1.23456789012345)
    Next
    ' Single, 7 digits
    s = 1
    For i = 1 To 12
        s = s / 10
        Print #f, "s"; i; CStr(s); "|"; Str(s); "|"; CStr(s * 1.5!); "|"; CStr(s * 1.234567!)
    Next
    ' large numbers
    d = 1
    For i = 1 To 17
        d = d * 10
        Print #f, "D"; i; CStr(d); "|"; CStr(d * 1.5); "|"; Str(d * 1.23456789)
    Next
    s = 1
    For i = 1 To 9
        s = s * 10
        Print #f, "S"; i; CStr(s); "|"; CStr(s * 1.5!); "|"; Str(s * 1.234567!)
    Next
    ' negative numbers and Print
    Print #f, CStr(-0.5); "|"; Str(-0.5); "|"; CStr(-1.5E-10); "|"; Str(-1.5E-10)
    Print #f, -0.5; 0.1!; CSng(1) / 3; -1.5E-10; 1.5E-10
    ' Write # (locale independent?)
    Write #f, 0.5, -0.5, 1.5, 0.5@, -0.25@, 1E-20, 2.5E-10, CSng(0.25), 1234567.5
    ' Currency
    Print #f, 0.5@; -0.5@; 1.5@; CStr(0.5@); "|"; Str(0.5@); "|"; Str(-0.25@); "|"; CStr(123456789.1234@)
    ' concatenation
    Print #f, "x" & 0.5 & "|" & 1.5! & "|" & 0.5@ & "|" & -0.5 & "|" & 1E-20
    ' text -> number with the regional settings
    Print #f, TryD("1,5"); "|"; TryD("1.5"); "|"; TryD("1.000,5"); "|"; TryD(".5"); "|"; TryD(",5"); "|"; TryD("1e3"); "|"; TryD(" 2 "); "|"; TryD("1,2,3"); "|"; TryD("1.2.3")
    Print #f, TryC("1,5"); "|"; TryC("2.5"); "|"; TryC("-0,25"); "|"; TryD("") ; "|"; TryD("abc"); "|"; TryD("&H10")
    Print #f, IsNumeric("1,5"); IsNumeric("1.5"); IsNumeric("1.2.3"); IsNumeric(""); IsNumeric(" 7 ")
    Print #f, Val("1,5"); Val("1.5"); Val("2,5e1")
    ' Round (banker's rounding on the binary value)
    Print #f, Round(2.345, 2); Round(2.5); Round(3.5); Round(-2.5); Round(0.125, 2); Round(1.005, 2); Round(2.675, 2); Round(0.5); Round(1.5)
    Print #f, CInt(2.5); CInt(3.5); CLng(-2.5); CInt(0.5); CInt(1.5)
End Sub
