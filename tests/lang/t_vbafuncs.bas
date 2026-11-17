Attribute VB_Name = "ModVbaFuncs"
Option Explicit

Private Function Doubles(ByVal n As Long) As Long()
    Dim r() As Long, i As Long
    ReDim r(1 To n)
    For i = 1 To n
        r(i) = i * 2
    Next
    Doubles = r
End Function

Public Sub T_VbaFuncs(f As Integer)
    Dim parts() As String, v As Variant, e As Variant, w As Variant
    parts = Split("a,b,,c", ",")
    Print #f, LBound(parts); UBound(parts); "["; Join(parts, "|"); "]"
    parts = Split("one two  three")
    Print #f, UBound(parts); "["; parts(3); "]"
    parts = Split("a-b-c-d", "-", 2)
    Print #f, UBound(parts); parts(1)
    parts = Split("", ",")
    Print #f, UBound(parts)
    parts = Split("aXbxc", "x", -1, vbTextCompare)
    Print #f, UBound(parts); Join(parts)
    v = Array(1, "two", 3.5, True)
    Print #f, IsArray(v); TypeName(v); VarType(v); LBound(v); UBound(v); v(1); v(2)
    v(1) = "TWO"
    Print #f, v(1); Join(Array("x", "y"), "+")
    For Each e In v
        Print #f, TypeName(e); " ";
    Next
    Print #f,
    w = v
    w(0) = 100
    Print #f, v(0); w(0)
    v = Split("p,q", ",")
    Print #f, TypeName(v); UBound(v); v(1)
    Dim fl() As String
    fl = Filter(Split("apple,banana,cherry,grape", ","), "an")
    Print #f, UBound(fl); Join(fl, ",")
    fl = Filter(Split("apple,banana,cherry", ","), "an", False)
    Print #f, Join(fl, ",")
    Dim dl() As Long
    dl = Doubles(3)
    Print #f, LBound(dl); UBound(dl); dl(3)
    Dim a2() As Long, a3() As Long
    ReDim a2(2)
    a2(0) = 5
    a3 = a2
    a3(0) = 9
    Print #f, a2(0); a3(0); UBound(a3); IsArray(a2); IsArray(5)

    Dim d As Date
    d = DateSerial(2024, 1, 31)
    Print #f, DateAdd("m", 1, d); "|"; DateAdd("d", 30, d); "|"; DateAdd("yyyy", -1, d); "|"; DateAdd("h", 25, d); "|"; DateAdd("ww", 2, d); "|"; DateAdd("q", 1, d)
    Print #f, DateDiff("d", d, DateSerial(2024, 3, 1)); DateDiff("m", d, DateSerial(2024, 3, 1)); DateDiff("yyyy", DateSerial(2023, 12, 31), d); DateDiff("ww", d, DateSerial(2024, 3, 1)); DateDiff("h", d, d + 1.5); DateDiff("n", d, d + 0.5)
    Print #f, DatePart("yyyy", d); DatePart("q", d); DatePart("m", d); DatePart("y", d); DatePart("d", d); DatePart("w", d); DatePart("ww", d); DatePart("h", d + 0.75)
    Print #f, DateValue("15/3/2010"); "|"; TimeValue("13:04"); "|"; Weekday(d, vbMonday)
    Print #f, InStr(1, "Hello", "L", vbTextCompare); InStr(1, "Hello", "L", vbBinaryCompare); InStr(3, "abcabc", "b"); InStrRev("abcabc", "b"); StrComp("a", "B", vbTextCompare); StrComp("a", "B")
    Print #f, StrConv("hello wORLD", vbProperCase); "|"; StrConv("ABC", vbLowerCase); "|"; StrConv("abc", vbUpperCase)
    Print #f, Choose(2, "a", "b", "c"); "|"; IsNull(Choose(5, "a")); "|"; Switch(False, 1, True, 2); "|"; Replace("aXbXc", "x", "-", , , vbTextCompare); "|"; Replace("aaaa", "a", "b", 2, 2)
    Print #f, Hex$(-1&); " "; Oct$(64); CLng(&HFFFF&); &HFFFF
End Sub
