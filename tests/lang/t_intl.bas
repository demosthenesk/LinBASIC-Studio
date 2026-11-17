Attribute VB_Name = "ModIntl"
Option Explicit
' Phase 6.7: international text - StrConv with a LocaleID (the ANSI code pages of Windows, single and double
' byte), vbWide / vbNarrow / vbKatakana / vbHiragana, case and comparison of Greek, Turkish and other letters.
' Everything is printed as numbers (the output file is ANSI).

Private Function U(ParamArray c() As Variant) As String
    Dim i As Long, s As String
    For i = 0 To UBound(c)
        s = s & ChrW$(c(i))
    Next
    U = s
End Function

Private Function Codes(ByVal s As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(s)
        r = r & " " & Hex$(AscW(Mid$(s, i, 1)) And &HFFFF&)
    Next
    Codes = Len(s) & ":" & r
End Function

Private Function BytesOf(ByVal s As String) As String
    Dim b() As Byte, i As Long, r As String
    b = s
    For i = 0 To UBound(b)
        r = r & " " & Hex$(b(i))
    Next
    BytesOf = LenB(s) & ":" & r
End Function

Private Function Raw(ParamArray v() As Variant) As String
    ' a byte string from the numbers
    Dim b() As Byte, i As Long
    ReDim b(UBound(v))
    For i = 0 To UBound(v)
        b(i) = v(i)
    Next
    Raw = b
End Function

Private Sub Conv(f As Integer, ByVal what As String, ByVal s As String, ByVal how As Long, ByVal lcid As Long)
    Dim r As String
    On Error Resume Next
    r = StrConv(s, how, lcid)
    If Err.Number Then
        Print #f, what; " err"; Err.Number
        Err.Clear
    ElseIf how And vbFromUnicode Then
        Print #f, what; " "; BytesOf(r)
    Else
        Print #f, what; " "; Codes(r)
    End If
End Sub

Private Sub ConvNoLcid(f As Integer, ByVal what As String, ByVal s As String, ByVal how As Long)
    Dim r As String
    On Error Resume Next
    r = StrConv(s, how)
    If Err.Number Then
        Print #f, what; " err"; Err.Number
        Err.Clear
    ElseIf how And vbFromUnicode Then
        Print #f, what; " "; BytesOf(r)
    Else
        Print #f, what; " "; Codes(r)
    End If
End Sub

' ---- 1. Unicode -> the code page of a locale, and back
Private Sub PartCodePages(f As Integer)
    Dim lcids As Variant, names As Variant, texts As Variant, tn As Variant, i As Long, j As Long, s As String
    lcids = Array(1033, 1032, 1045, 1049, 1055, 1037, 1025, 1063, 1066, 1054, 1041, 2052, 1042, 1028)
    names = Array("1252", "1253", "1250", "1251", "1254", "1255", "1256", "1257", "1258", "874", "932", "936", "949", "950")
    texts = Array(U(65, 122, 48), U(233, 252, 241, 223, 8364), U(913, 940, 962, 974, 902), U(1041, 1078, 1103, 1105), _
                  U(287, 304, 305, 351), U(1488, 1513, 1514), U(1575, 1604, 1740), U(12354, 12450, 12539, 65393), _
                  U(26085, 26412, 35486), U(20013, 25991), U(54620, 44544), U(39636, 33274), U(3585, 3647), U(259, 7841), _
                  U(65, 0, 66), U(65533, 57344))
    tn = Array("ascii", "latin", "greek", "cyrillic", "turkish", "hebrew", "arabic", "kana", "kanji", "hanzi", "hangul", _
               "trad", "thai", "viet", "nul", "special")
    For i = 0 To UBound(lcids)
        For j = 0 To UBound(texts)
            Conv f, "cp" & names(i) & " " & tn(j), texts(j), vbFromUnicode, lcids(i)
        Next
    Next
    ' bytes -> Unicode: valid and broken double-byte sequences
    Conv f, "932 a", Raw(&H82, &HA0, &H83, &H41, &HB1, &H41), vbUnicode, 1041
    Conv f, "932 lead alone", Raw(&H41, &H82), vbUnicode, 1041
    Conv f, "932 bad trail", Raw(&H82, &H20, &H41), vbUnicode, 1041
    Conv f, "932 user area", Raw(&HF0, &H40), vbUnicode, 1041
    Conv f, "932 undefined", Raw(&H85, &H40, &HA0, &HFD), vbUnicode, 1041
    Conv f, "936 a", Raw(&HD6, &HD0, &HCE, &HC4, &H80), vbUnicode, 2052
    Conv f, "949 a", Raw(&HC7, &HD1, &HB1, &HDB, &HA4, &HA1), vbUnicode, 1042
    Conv f, "950 a", Raw(&HA4, &HA4, &HA4, &HE5, &HA2, &HB0), vbUnicode, 1028
    Conv f, "1253 all", Raw(&H80, &H81, &HAA, &HD2, &HFF, &HE1), vbUnicode, 1032
    Conv f, "1255 a", Raw(&HE0, &HF9, &HC0, &HCA, &HD9), vbUnicode, 1037
    Conv f, "874 a", Raw(&HA1, &HDB, &HFC, &H80), vbUnicode, 1054
    Conv f, "1258 a", Raw(&HC3, &HD2, &HDE, &HEC), vbUnicode, 1066
    Conv f, "1251 a", Raw(&H98, &HC0, &HFF, &HB9), vbUnicode, 1049
    ' the locale of the system (no LocaleID), and LocaleID 0 / a wrong one
    ConvNoLcid f, "system greek", U(913, 940, 12354), vbFromUnicode
    Conv f, "lcid 0", U(913, 12354), vbFromUnicode, 0
    Conv f, "lcid 1", U(913, 12354), vbFromUnicode, 1
    Conv f, "lcid 99999", U(913, 12354), vbFromUnicode, 99999
    Conv f, "lcid 2048", U(913, 12354), vbFromUnicode, 2048
    Conv f, "lcid 1024", U(913, 12354), vbFromUnicode, 1024
    ' LenB and the odd byte count of a double-byte conversion
    s = StrConv(U(12354, 65), vbFromUnicode, 1041)
    Print #f, "LenB"; LenB(s); Len(s)
End Sub

' ---- 2. vbWide / vbNarrow / vbKatakana / vbHiragana
Private Sub PartWidth(f As Integer)
    Dim s As String, k As Variant, lc As Variant, how As Variant, hn As Variant, i As Long, j As Long
    s = U(65, 97, 49, 32, 33, 126, 12354, 12450, 12459, 12460, 65393, 65398, 65438, 12290, 12539, 12540, 65377, 913, 1041)
    how = Array(vbWide, vbNarrow, vbKatakana, vbHiragana, vbWide + vbKatakana, vbNarrow + vbHiragana, vbWide + vbNarrow, _
                vbKatakana + vbHiragana, vbUpperCase + vbWide, vbProperCase + vbNarrow)
    hn = Array("wide", "narrow", "katakana", "hiragana", "wide+kata", "narrow+hira", "wide+narrow", "kata+hira", "upper+wide", _
               "proper+narrow")
    For i = 0 To UBound(how)
        ConvNoLcid f, hn(i) & " system", s, how(i)
        For Each lc In Array(1041, 2052, 1042, 1028, 1033)
            Conv f, hn(i) & " " & lc, s, how(i), lc
        Next
    Next
    Conv f, "wide fromunicode", U(65, 12354), vbWide + vbFromUnicode, 1041
    Conv f, "narrow unicode", Raw(&H82, &H60), vbNarrow + vbUnicode, 1041
End Sub

' ---- 3. case: UCase / LCase / StrConv case with a locale
Private Sub PartCase(f As Integer)
    Dim s As String, t As String
    s = U(940, 941, 942, 943, 972, 973, 974, 912, 944, 962, 963, 931, 902, 970, 223, 255, 105, 73, 305, 304, 491, 1105, 64256, 453)
    Print #f, "UCase "; Codes(UCase$(s))
    Print #f, "LCase "; Codes(LCase$(s))
    ConvNoLcid f, "upper system", s, vbUpperCase
    ConvNoLcid f, "lower system", s, vbLowerCase
    Conv f, "upper 1055", s, vbUpperCase, 1055
    Conv f, "lower 1055", s, vbLowerCase, 1055
    Conv f, "upper 1033", s, vbUpperCase, 1033
    t = U(940, 955, 966, 945, 32, 946, 942, 964, 945, 32, 963, 943, 947, 956, 945, 962, 32, 105, 115, 116, 97, 110, 98, 117, 108)
    ConvNoLcid f, "proper system", t, vbProperCase
    Conv f, "proper 1055", t, vbProperCase, 1055
    ConvNoLcid f, "proper marks", U(97, 39, 98, 45, 99, 46, 100, 9, 101, 48, 102, 95, 103, 183, 104), vbProperCase
End Sub

' ---- 4. comparing: Binary and Text, Greek and the others
Private Sub PartCompare(f As Integer)
    Dim pairs As Variant, i As Long, a As String, b As String
    pairs = Array(U(945), U(913), U(940), U(945), U(902), U(940), U(962), U(963), U(963), U(931), U(223), U(115, 115), _
                  U(230), U(97, 101), U(197), U(229), U(65), U(913), U(945), U(122), U(945), U(1072), U(97), U(224), _
                  U(105), U(304), U(305), U(73), U(12354), U(12450), U(65393), U(12450), U(65), U(65313), U(50), U(178), _
                  U(97, 98), U(97, 45, 98), U(39, 97), U(97), U(948, 940, 962), U(916, 913, 931), U(912), U(953), U(970), U(953))
    For i = 0 To UBound(pairs) Step 2
        a = pairs(i): b = pairs(i + 1)
        Print #f, "cmp"; i \ 2; StrComp(a, b, vbBinaryCompare); StrComp(a, b, vbTextCompare); _
                  InStr(1, a, b, vbTextCompare); a < b
    Next
    Print #f, "instr greek"; InStr(1, U(913, 914, 915), U(946), vbTextCompare); InStr(1, U(940, 946), U(902), vbTextCompare)
    Print #f, "replace greek "; Codes(Replace(U(913, 945, 940, 902), U(945), "x", , , vbTextCompare))
    Print #f, "filter "; UBound(Filter(Array(U(913), U(945), U(946)), U(945), True, vbTextCompare))
End Sub

' ---- 5. Asc, Chr and the code page of the system (Greek)
Private Sub PartAscChr(f As Integer)
    Dim i As Long, r As String
    On Error Resume Next
    For i = 128 To 255 Step 9
        r = r & " " & Hex$(AscW(Chr$(i)))
    Next
    Print #f, "Chr"; r
    Print #f, "Asc"; Asc(U(913)); Asc(U(8364)); Asc(U(12354)); Asc(U(1041)); Asc(U(902))
    Print #f, "AscB"; AscB(U(913)); AscB(U(12354))
    r = Chr$(&H82A0)
    Print #f, "Chr(&H82A0)"; Err.Number; Len(r)
    Err.Clear
    r = Chr$(-1)
    Print #f, "Chr(-1)"; Err.Number
    Err.Clear
    Print #f, "ChrB "; Codes(ChrB(&HE1) & ChrB(0))
    Print #f, "Len of 932 bytes"; Len(StrConv(Raw(&H82, &HA0, &H41), vbUnicode, 1041)); Len(StrConv(Raw(&H82, &HA0, &H41), vbUnicode))
    Print #f, "Val"; Val(U(65297, 65298)); Val(U(1633, 1634)); IsNumeric(U(65297))
    Print #f, "UCase ansi "; Codes(UCase$(Chr$(225) & Chr$(220) & Chr$(242)))
End Sub

Public Sub T_Intl(f As Integer)
    PartCodePages f
    PartWidth f
    PartCase f
    PartCompare f
    PartAscChr f
End Sub
