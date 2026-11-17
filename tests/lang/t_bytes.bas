Attribute VB_Name = "ModBytes"
Option Explicit
' Byte strings (Phase 6.3): a String with any number of bytes, against VB6
Private mF As Integer

Private Function Hx(ByVal s As String) As String
    Dim b() As Byte, i As Long, h As String
    If LenB(s) = 0 Then Hx = "-": Exit Function
    b = s
    For i = LBound(b) To UBound(b): h = h & Right$("0" & Hex$(b(i)), 2) & " ": Next
    Hx = RTrim$(h)
End Function

Private Sub Rep(ByVal nm As String, v As Variant, ByVal e As Long)
    Dim t As String, s As String
    On Error Resume Next
    t = "?"
    If VarType(v) = vbString Then
        s = v
        t = "String LenB " & LenB(s) & " Len " & Len(s) & " [" & Hx(s) & "]"
    Else
        t = TypeName(v) & " " & CStr(v)
    End If
    Print #mF, nm & " = " & t & IIf(e <> 0, " err" & e, "")
End Sub

Public Sub T_Bytes(f As Integer)
    Dim rv As Variant
    mF = f
    On Error Resume Next
    Err.Clear: rv = "(none)"
    rv = ChrB(65)
    Rep "ChrB(65)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) & ChrB(66)
    Rep "ChrB(65) & ChrB(66)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) & ChrB(66) & ChrB(67)
    Rep "ChrB(65) & ChrB(66) & ChrB(67)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = "A" & ChrB(66)
    Rep "'A' & ChrB(66)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) & "B"
    Rep "ChrB(65) & 'B'", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(200)
    Rep "ChrB(200)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(0)
    Rep "ChrB(0)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(255)
    Rep "ChrB(255)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(256)
    Rep "ChrB(256)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(-1)
    Rep "ChrB(-1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB$(66)
    Rep "ChrB$(66)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) & ChrB(66) & "C"
    Rep "ChrB(65) & ChrB(66) & 'C'", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) & ChrB(66) & ChrB(67) & ChrB(68)
    Rep "ChrB(65) & ChrB(66) & ChrB(67) & ChrB(68)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = "" & ChrB(65)
    Rep "'' & ChrB(65)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) & ""
    Rep "ChrB(65) & ''", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = AscB("A")
    Rep "AscB('A')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = AscB(ChrW(937))
    Rep "AscB(ChrW(937))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = AscB(ChrB(200))
    Rep "AscB(ChrB(200))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = AscB("")
    Rep "AscB('')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = AscB(ChrB(65) & ChrB(66))
    Rep "AscB(ChrB(65) & ChrB(66))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = AscB(Chr(200))
    Rep "AscB(Chr(200))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = AscW(ChrB(65))
    Rep "AscW(ChrB(65))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Asc(ChrB(65))
    Rep "Asc(ChrB(65))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = AscW(ChrB(65) & ChrB(66))
    Rep "AscW(ChrB(65) & ChrB(66))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Asc(ChrB(65) & ChrB(66))
    Rep "Asc(ChrB(65) & ChrB(66))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LenB("ABC")
    Rep "LenB('ABC')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LenB(ChrB(65))
    Rep "LenB(ChrB(65))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Len(ChrB(65))
    Rep "Len(ChrB(65))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Len(ChrB(65) & ChrB(66))
    Rep "Len(ChrB(65) & ChrB(66))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Len(ChrB(65) & ChrB(66) & ChrB(67))
    Rep "Len(ChrB(65) & ChrB(66) & ChrB(67))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LenB("")
    Rep "LenB('')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LenB(StrConv("ABC", vbFromUnicode))
    Rep "LenB(StrConv('ABC', vbFromUnicode))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LenB(CVar(ChrB(65)))
    Rep "LenB(CVar(ChrB(65)))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LenB(CVar(12345))
    Rep "LenB(CVar(12345))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LeftB("ABC", 1)
    Rep "LeftB('ABC', 1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LeftB("ABC", 2)
    Rep "LeftB('ABC', 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LeftB("ABC", 3)
    Rep "LeftB('ABC', 3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LeftB("ABC", 10)
    Rep "LeftB('ABC', 10)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LeftB("ABC", 0)
    Rep "LeftB('ABC', 0)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LeftB("ABC", -1)
    Rep "LeftB('ABC', -1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = RightB("ABC", 1)
    Rep "RightB('ABC', 1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = RightB("ABC", 2)
    Rep "RightB('ABC', 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = RightB("ABC", 3)
    Rep "RightB('ABC', 3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = RightB("ABC", 10)
    Rep "RightB('ABC', 10)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = RightB("ABC", 0)
    Rep "RightB('ABC', 0)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB("ABC", 2)
    Rep "MidB('ABC', 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB("ABC", 2, 2)
    Rep "MidB('ABC', 2, 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB("ABC", 3, 2)
    Rep "MidB('ABC', 3, 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB("ABC", 6)
    Rep "MidB('ABC', 6)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB("ABC", 7)
    Rep "MidB('ABC', 7)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB("ABC", 0)
    Rep "MidB('ABC', 0)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB("ABC", 1, 0)
    Rep "MidB('ABC', 1, 0)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB("ABC", 2, 100)
    Rep "MidB('ABC', 2, 100)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB$("ABC", 3)
    Rep "MidB$('ABC', 3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = LeftB$("ABC", 3)
    Rep "LeftB$('ABC', 3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = RightB$("ABC", 3)
    Rep "RightB$('ABC', 3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MidB(ChrB(65) & ChrB(66) & ChrB(67), 2, 1)
    Rep "MidB(ChrB(65) & ChrB(66) & ChrB(67), 2, 1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = RightB(ChrB(65) & ChrB(66) & ChrB(67), 1)
    Rep "RightB(ChrB(65) & ChrB(66) & ChrB(67), 1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB("ABC", "B")
    Rep "InStrB('ABC', 'B')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB("ABC", ChrB(66))
    Rep "InStrB('ABC', ChrB(66))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB("ABC", ChrB(0))
    Rep "InStrB('ABC', ChrB(0))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(2, "ABAB", "A")
    Rep "InStrB(2, 'ABAB', 'A')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(1, ChrB(65) & ChrB(66), ChrB(66))
    Rep "InStrB(1, ChrB(65) & ChrB(66), ChrB(66))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB("ABC", "")
    Rep "InStrB('ABC', '')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB("", "A")
    Rep "InStrB('', 'A')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(4, "ABC", "C")
    Rep "InStrB(4, 'ABC', 'C')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(7, "ABC", "")
    Rep "InStrB(7, 'ABC', '')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(8, "ABC", "")
    Rep "InStrB(8, 'ABC', '')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(0, "ABC", "A")
    Rep "InStrB(0, 'ABC', 'A')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(1, "abc", "B", vbTextCompare)
    Rep "InStrB(1, 'abc', 'B', vbTextCompare)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(1, "abc", "B", vbBinaryCompare)
    Rep "InStrB(1, 'abc', 'B', vbBinaryCompare)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(ChrB(1) & ChrB(2) & ChrB(3), ChrB(2) & ChrB(3))
    Rep "InStrB(ChrB(1) & ChrB(2) & ChrB(3), ChrB(2) & ChrB(3))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = InStrB(2, ChrW(&H4142) & "C", ChrB(&H41))
    Rep "InStrB(2, ChrW(&H4142) & 'C', ChrB(&H41))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Left(ChrB(65) & ChrB(66) & ChrB(67), 1)
    Rep "Left(ChrB(65) & ChrB(66) & ChrB(67), 1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Right(ChrB(65) & ChrB(66) & ChrB(67), 1)
    Rep "Right(ChrB(65) & ChrB(66) & ChrB(67), 1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Mid(ChrB(65) & ChrB(66) & ChrB(67), 1)
    Rep "Mid(ChrB(65) & ChrB(66) & ChrB(67), 1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Mid(ChrB(65) & ChrB(66) & ChrB(67), 2)
    Rep "Mid(ChrB(65) & ChrB(66) & ChrB(67), 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = UCase(ChrB(97) & ChrB(0) & ChrB(98))
    Rep "UCase(ChrB(97) & ChrB(0) & ChrB(98))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = UCase(ChrB(97) & ChrB(98) & ChrB(99))
    Rep "UCase(ChrB(97) & ChrB(98) & ChrB(99))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Trim(ChrB(32) & ChrB(0) & ChrB(65))
    Rep "Trim(ChrB(32) & ChrB(0) & ChrB(65))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrReverse(ChrB(65) & ChrB(66) & ChrB(67))
    Rep "StrReverse(ChrB(65) & ChrB(66) & ChrB(67))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Replace(ChrB(65) & ChrB(0) & ChrB(66), "A", "x")
    Rep "Replace(ChrB(65) & ChrB(0) & ChrB(66), 'A', 'x')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Space(1) & ChrB(65)
    Rep "Space(1) & ChrB(65)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = String(2, ChrB(65))
    Rep "String(2, ChrB(65))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CStr(ChrB(65))
    Rep "CStr(ChrB(65))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CVar(ChrB(65) & ChrB(66) & ChrB(67))
    Rep "CVar(ChrB(65) & ChrB(66) & ChrB(67))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) = ""
    Rep "ChrB(65) = ''", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) & ChrB(0) = "A"
    Rep "ChrB(65) & ChrB(0) = 'A'", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) = ChrB(65)
    Rep "ChrB(65) = ChrB(65)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) < ChrB(66)
    Rep "ChrB(65) < ChrB(66)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(66) > "A"
    Rep "ChrB(66) > 'A'", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrComp(ChrB(65), ChrB(66), vbBinaryCompare)
    Rep "StrComp(ChrB(65), ChrB(66), vbBinaryCompare)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrComp(ChrB(65), "", vbBinaryCompare)
    Rep "StrComp(ChrB(65), '', vbBinaryCompare)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrComp(ChrB(65), ChrB(65) & ChrB(0), vbBinaryCompare)
    Rep "StrComp(ChrB(65), ChrB(65) & ChrB(0), vbBinaryCompare)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrComp(ChrB(65), ChrB(65) & ChrB(0), vbTextCompare)
    Rep "StrComp(ChrB(65), ChrB(65) & ChrB(0), vbTextCompare)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = "A" = ChrB(65) & ChrB(0)
    Rep "'A' = ChrB(65) & ChrB(0)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) & ChrB(66) & ChrB(67) = ChrB(65) & ChrB(66) & ChrB(67)
    Rep "ChrB(65) & ChrB(66) & ChrB(67) = ChrB(65) & ChrB(66) & ChrB(67)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) & ChrB(66) & ChrB(67) = ChrB(65) & ChrB(66) & ChrB(68)
    Rep "ChrB(65) & ChrB(66) & ChrB(67) = ChrB(65) & ChrB(66) & ChrB(68)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = ChrB(65) Like "*"
    Rep "ChrB(65) Like '*'", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Len(ChrB(65)) = 0
    Rep "Len(ChrB(65)) = 0", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv("AB" & ChrW(915), vbFromUnicode)
    Rep "StrConv('AB' & ChrW(915), vbFromUnicode)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv(StrConv("AB" & ChrW(915), vbFromUnicode), vbUnicode)
    Rep "StrConv(StrConv('AB' & ChrW(915), vbFromUnicode), vbUnicode)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv("AB", vbFromUnicode)
    Rep "StrConv('AB', vbFromUnicode)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv(ChrB(65), vbUnicode)
    Rep "StrConv(ChrB(65), vbUnicode)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv(ChrB(65) & ChrB(66) & ChrB(195), vbUnicode)
    Rep "StrConv(ChrB(65) & ChrB(66) & ChrB(195), vbUnicode)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv("", vbFromUnicode)
    Rep "StrConv('', vbFromUnicode)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv(ChrB(0), vbUnicode)
    Rep "StrConv(ChrB(0), vbUnicode)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv("A" & ChrW(8364), vbFromUnicode)
    Rep "StrConv('A' & ChrW(8364), vbFromUnicode)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv(ChrB(128), vbUnicode)
    Rep "StrConv(ChrB(128), vbUnicode)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv(ChrB(65), vbUpperCase)
    Rep "StrConv(ChrB(65), vbUpperCase)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrConv("abc", vbFromUnicode + vbUpperCase)
    Rep "StrConv('abc', vbFromUnicode + vbUpperCase)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Val(ChrB(49) & ChrB(0) & ChrB(50))
    Rep "Val(ChrB(49) & ChrB(0) & ChrB(50))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CStr(ChrB(49)) & "x"
    Rep "CStr(ChrB(49)) & 'x'", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = IsNumeric(ChrB(49))
    Rep "IsNumeric(ChrB(49))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Hex(AscB(ChrB(171)))
    Rep "Hex(AscB(ChrB(171)))", rv, Err.Number
    ' Byte arrays and strings
    Dim b() As Byte, b2() As Byte, s As String, i As Long, k As Long, t As String
    Err.Clear
    ReDim b(2): b(0) = 65: b(1) = 66: b(2) = 67
    s = b
    Rep "s = b(0 To 2)", s, Err.Number
    Rep "StrConv(b, vbUnicode)", StrConv(b, vbUnicode), Err.Number
    b2 = StrConv("ABC", vbFromUnicode)
    Rep "UBound(StrConv(ABC, vbFromUnicode))", UBound(b2), Err.Number
    b2 = "AB"
    Rep "UBound(b = AB)", UBound(b2), Err.Number
    b2 = ChrB(1) & ChrB(2) & ChrB(3)
    Rep "UBound(b = 3 bytes)", UBound(b2), Err.Number
    b2 = ""
    Rep "UBound(b = empty)", UBound(b2), Err.Number
    Err.Clear: ReDim b(0): b(0) = 200: s = b
    Rep "s = b(0) = 200", s, Err.Number
    ' MidB statement
    Err.Clear: s = "ABC": MidB(s, 2, 1) = "Z"
    Rep "MidB(ABC, 2, 1) = Z", s, Err.Number
    Err.Clear: s = "ABC": MidB(s, 3) = "ZY"
    Rep "MidB(ABC, 3) = ZY", s, Err.Number
    Err.Clear: s = "ABC": MidB(s, 6, 4) = "ZYX"
    Rep "MidB(ABC, 6, 4) = ZYX", s, Err.Number
    Err.Clear: s = ChrB(65) & ChrB(66) & ChrB(67): MidB(s, 3, 1) = "Z"
    Rep "MidB(3 bytes, 3, 1) = Z", s, Err.Number
    Err.Clear: s = "ABC": Mid(s, 2, 1) = ChrB(90)
    Rep "Mid(ABC, 2, 1) = ChrB(90)", s, Err.Number
    ' building a string byte by byte
    Err.Clear: s = ""
    For i = 1 To 5: s = s & ChrB(64 + i): Next
    Rep "5 x ChrB", s, Err.Number
    t = ""
    For i = 1 To LenB(s): t = t & Chr$(AscB(MidB$(s, i, 1))): Next
    Rep "read back", t, Err.Number
    ' a byte string in a fixed-length string and in a Type
    Dim fx As String * 3
    Err.Clear: fx = ChrB(65) & ChrB(66) & ChrB(67)
    Rep "String * 3 = 3 bytes", fx, Err.Number
    ' files
    Dim g As Integer, hx As String, bb As Byte
    g = FreeFile
    Open App.Path & "\bytes.bin" For Binary As #g
    s = ChrB(65) & ChrB(66) & ChrB(67)
    Put #g, , s
    s = ChrB(68) & ChrB(69) & ChrB(70)
    Put #g, , s
    Rep "LOF after Put", LOF(g), Err.Number
    hx = ""
    For k = 1 To LOF(g): Get #g, k, bb: hx = hx & Right$("0" & Hex$(bb), 2): Next
    Rep "file bytes", hx, Err.Number
    s = String(3, 0): Get #g, 1, s
    Rep "Get String(3)", s, Err.Number
    Close #g
    Kill App.Path & "\bytes.bin"
    g = FreeFile
    Open App.Path & "\bytes.txt" For Output As #g
    Print #g, ChrB(65) & ChrB(66) & ChrB(67);
    Print #g, "|"; StrConv("AB" & ChrW(915), vbFromUnicode); "|"
    Write #g, ChrB(65) & ChrB(66) & ChrB(67)
    Close #g
    g = FreeFile
    Open App.Path & "\bytes.txt" For Binary As #g
    hx = ""
    For k = 1 To LOF(g): Get #g, k, bb: hx = hx & Right$("0" & Hex$(bb), 2): Next
    Close #g
    Kill App.Path & "\bytes.txt"
    Rep "Print/Write # bytes", hx, Err.Number
End Sub
