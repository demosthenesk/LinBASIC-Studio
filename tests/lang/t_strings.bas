Attribute VB_Name = "ModStrings"
Option Explicit
Public Sub T_Strings(f As Integer)
    Dim s As String, t As String
    s = "Hello, World"
    Print #f, Len(s); Left$(s, 5); "|"; Right$(s, 5); "|"; Mid$(s, 8); "|"; Mid$(s, 8, 3); "|"
    Print #f, InStr(s, "o"); InStr(6, s, "o"); InStr(s, "z"); InStr(1, s, "WORLD"); InStr(1, s, "WORLD", vbTextCompare)
    Print #f, InStrRev(s, "o"); InStrRev(s, "l", 5)
    Print #f, UCase$(s); " "; LCase$(s)
    Print #f, "["; Trim$("  ab  "); "]["; LTrim$("  ab  "); "]["; RTrim$("  ab  "); "]"
    Print #f, "["; Space$(3); "]["; String$(4, "*"); "]["; String$(3, 65); "]"
    Print #f, Asc("A"); Asc("abc"); Chr$(66); Chr(97); AscW("A"); ChrW(67)
    Print #f, StrReverse("abc"); " "; Replace("aaa", "a", "bb"); " "; Replace("abcabc", "b", "", 1, 1)
    Print #f, StrComp("a", "b"); StrComp("b", "a"); StrComp("a", "A"); StrComp("a", "A", vbTextCompare)
    t = "abcdef"
    Mid$(t, 2, 3) = "XYZW"
    Print #f, t
    Mid(t, 5) = "12345"
    Print #f, t
    Print #f, "a" & 1 & True & 2.5; "x" + "y"
    Print #f, "concat number"; "5" + 3; "5" & 3
    Print #f, Len(""); Left$("abc", 0); "|"; Left$("abc", 10); "|"; Mid$("abc", 5); "|"
    s = vbNullString
    Print #f, Len(s); Len(vbCrLf); Asc(vbTab)
    Print #f, "Abc" = "abc"; "abc" < "abd"; "ab" < "abc"; "Z" < "a"
End Sub
