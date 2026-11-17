Attribute VB_Name = "ModDecimal"
Option Explicit
' The Decimal subtype of the Variant (Phase 6.2): every case against VB6
Private mF As Integer

Private Function D(ByVal s As String) As Variant
    D = CDec(Replace(s, ".", Mid$(CStr(1.5), 2, 1)))
End Function

Private Function MaxD() As Variant
    MaxD = CDec("79228162514264337593543950335")
End Function

Private Function Twice(ByVal x As Double) As Double
    Twice = x * 2
End Function

Private Function Half(ByVal x As Variant) As Variant
    Half = x / 2
End Function

Private Sub Rep(ByVal nm As String, v As Variant, ByVal e As Long)
    Dim t As String
    On Error Resume Next
    t = "?"
    t = TypeName(v) & " " & CStr(v)
    Print #mF, nm & " = " & t & IIf(e <> 0, " err" & e, "")
End Sub

Public Sub T_Decimal(f As Integer)
    Dim rv As Variant
    mF = f
    On Error Resume Next
    Err.Clear: rv = "(none)"
    rv = CDec("123")
    Rep "CDec('123')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("123.456")
    Rep "D('123.456')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("-0.5")
    Rep "D('-0.5')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("79228162514264337593543950335")
    Rep "D('79228162514264337593543950335')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("79228162514264337593543950336")
    Rep "D('79228162514264337593543950336')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("-79228162514264337593543950335")
    Rep "D('-79228162514264337593543950335')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.0000000000000000000000000001")
    Rep "D('0.0000000000000000000000000001')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.00000000000000000000000000001")
    Rep "D('0.00000000000000000000000000001')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.00000000000000000000000000005")
    Rep "D('0.00000000000000000000000000005')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.00000000000000000000000000015")
    Rep "D('0.00000000000000000000000000015')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.23456789012345678901234567890")
    Rep "D('1.23456789012345678901234567890')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("12345678901234567890.123456789")
    Rep "D('12345678901234567890.123456789')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("7.9228162514264337593543950335")
    Rep "D('7.9228162514264337593543950335')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.10")
    Rep "D('1.10')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.100000")
    Rep "D('1.100000')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.0")
    Rep "D('0.0')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("-0")
    Rep "D('-0')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1E+5")
    Rep "D('1E+5')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.5E-3")
    Rep "D('1.5E-3')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("  42  ")
    Rep "D('  42  ')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec("abc")
    Rep "CDec('abc')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec("")
    Rep "CDec('')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("&H10")
    Rep "D('&H10')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1,000.5")
    Rep "D('1,000.5')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1.1#)
    Rep "CDec(1.1#)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(0.1!)
    Rep "CDec(0.1!)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1# / 3#)
    Rep "CDec(1# / 3#)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(2# / 3#)
    Rep "CDec(2# / 3#)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(123456789012345.6789#)
    Rep "CDec(123456789012345.6789#)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1E+20)
    Rep "CDec(1E+20)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1E+28)
    Rep "CDec(1E+28)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1E+29)
    Rep "CDec(1E+29)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1E-20)
    Rep "CDec(1E-20)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1E-28)
    Rep "CDec(1E-28)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1E-29)
    Rep "CDec(1E-29)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(3.14159265358979)
    Rep "CDec(3.14159265358979)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(32767)
    Rep "CDec(32767)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(-32768)
    Rep "CDec(-32768)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(&H7FFFFFFF)
    Rep "CDec(&H7FFFFFFF)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(922337203685477.5807@)
    Rep "CDec(922337203685477.5807@)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(0.0001@)
    Rep "CDec(0.0001@)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(12.5@)
    Rep "CDec(12.5@)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(True)
    Rep "CDec(True)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(False)
    Rep "CDec(False)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(Empty)
    Rep "CDec(Empty)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(Null)
    Rep "CDec(Null)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(#1/1/2000#)
    Rep "CDec(#1/1/2000#)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(#1/1/2000 12:00:00 PM#)
    Rep "CDec(#1/1/2000 12:00:00 PM#)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(CByte(200))
    Rep "CDec(CByte(200))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1.5!)
    Rep "CDec(1.5!)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(16777217!)
    Rep "CDec(16777217!)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(0.3!)
    Rep "CDec(0.3!)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1))
    Rep "VarType(CDec(1))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = TypeName(CDec(1))
    Rep "TypeName(CDec(1))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = IsNumeric(CDec(1))
    Rep "IsNumeric(CDec(1))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) + 1)
    Rep "VarType(CDec(1) + 1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) + 1.5)
    Rep "VarType(CDec(1) + 1.5)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) * 1@)
    Rep "VarType(CDec(1) * 1@)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) / 2)
    Rep "VarType(CDec(1) / 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(7) \ 2)
    Rep "VarType(CDec(7) \ 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(7) Mod 2)
    Rep "VarType(CDec(7) Mod 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(2) ^ 2)
    Rep "VarType(CDec(2) ^ 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(-CDec(2))
    Rep "VarType(-CDec(2))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) + "1")
    Rep "VarType(CDec(1) + '1')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) + True)
    Rep "VarType(CDec(1) + True)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) & "")
    Rep "VarType(CDec(1) & '')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(Int(CDec(1)))
    Rep "VarType(Int(CDec(1)))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(Abs(CDec(1)))
    Rep "VarType(Abs(CDec(1)))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) + 1!)
    Rep "VarType(CDec(1) + 1!)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) + CDate(1))
    Rep "VarType(CDec(1) + CDate(1))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) + Empty)
    Rep "VarType(CDec(1) + Empty)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(CDec(1) + Null)
    Rep "VarType(CDec(1) + Null)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.1") + D("2.2")
    Rep "D('1.1') + D('2.2')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.1") - D("2.2")
    Rep "D('1.1') - D('2.2')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.1") * D("2.2")
    Rep "D('1.1') * D('2.2')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.1") / D("2.2")
    Rep "D('1.1') / D('2.2')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) / CDec(3)
    Rep "CDec(1) / CDec(3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(2) / CDec(3)
    Rep "CDec(2) / CDec(3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) / CDec(7)
    Rep "CDec(1) / CDec(7)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(10) / CDec(4)
    Rep "CDec(10) / CDec(4)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) / CDec(8)
    Rep "CDec(1) / CDec(8)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(100) / CDec(3)
    Rep "CDec(100) / CDec(3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1000000) / CDec(3)
    Rep "CDec(1000000) / CDec(3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("12345678901234567890") / CDec(3)
    Rep "D('12345678901234567890') / CDec(3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) / CDec(0)
    Rep "CDec(1) / CDec(0)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(0) / CDec(0)
    Rep "CDec(0) / CDec(0)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(0) / CDec(5)
    Rep "CDec(0) / CDec(5)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() + 1
    Rep "MaxD() + 1", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() - 1
    Rep "MaxD() - 1", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() * 2
    Rep "MaxD() * 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = -MaxD() - 1
    Rep "-MaxD() - 1", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() / D("0.5")
    Rep "MaxD() / D('0.5')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() / 10
    Rep "MaxD() / 10", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() * D("0.1")
    Rep "MaxD() * D('0.1')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() + D("0.4")
    Rep "MaxD() + D('0.4')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() + D("0.5")
    Rep "MaxD() + D('0.5')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() + D("0.6")
    Rep "MaxD() + D('0.6')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.1") * D("0.1")
    Rep "D('0.1') * D('0.1')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.0000000000000000000000000001") * D("1.0000000000000000000000000001")
    Rep "D('1.0000000000000000000000000001') * D('1.0000000000000000000000000001')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.0000000000000000000000000001") * D("0.1")
    Rep "D('0.0000000000000000000000000001') * D('0.1')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.0000000000000000000000000001") / 2
    Rep "D('0.0000000000000000000000000001') / 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.0000000000000000000000000003") / 2
    Rep "D('0.0000000000000000000000000003') / 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("12345678901234567890") * D("12345678901234567890")
    Rep "D('12345678901234567890') * D('12345678901234567890')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1234567890123.4567890123456") * D("1234567.8901234567")
    Rep "D('1234567890123.4567890123456') * D('1234567.8901234567')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) / D("0.0000000000000000000000000001")
    Rep "CDec(1) / D('0.0000000000000000000000000001')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("3.3333333333333333333333333333") * 3
    Rep "D('3.3333333333333333333333333333') * 3", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.10") + D("1.10")
    Rep "D('1.10') + D('1.10')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.10") * 2
    Rep "D('1.10') * 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("2.50") * D("2.0")
    Rep "D('2.50') * D('2.0')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(2.5) * 2
    Rep "CDec(2.5) * 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.000") - D("1")
    Rep "D('1.000') - D('1')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("5.0") / D("2.0")
    Rep "D('5.0') / D('2.0')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1E-28") + 1
    Rep "D('1E-28') + 1", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.5") + D("0.0000000000000000000000000001")
    Rep "D('0.5') + D('0.0000000000000000000000000001')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("123456789012345678901234567.8") + D("0.05")
    Rep "D('123456789012345678901234567.8') + D('0.05')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("123456789012345678901234567.8") + D("0.15")
    Rep "D('123456789012345678901234567.8') + D('0.15')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("7922816251426433759354395033.5") + D("0.1")
    Rep "D('7922816251426433759354395033.5') + D('0.1')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + 1
    Rep "CDec(1) + 1", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + 100000
    Rep "CDec(1) + 100000", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + 1.5
    Rep "CDec(1) + 1.5", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + 1.5!
    Rep "CDec(1) + 1.5!", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.1") + 0.2
    Rep "D('0.1') + 0.2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + 1@
    Rep "CDec(1) + 1@", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + "2"
    Rep "CDec(1) + '2'", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + True
    Rep "CDec(1) + True", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) & "x"
    Rep "CDec(1) & 'x'", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + Null
    Rep "CDec(1) + Null", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + Empty
    Rep "CDec(1) + Empty", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) + CDate(1)
    Rep "CDec(1) + CDate(1)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) * 2.5@
    Rep "CDec(1) * 2.5@", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.1") * 3
    Rep "D('0.1') * 3", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = 3 * D("0.1")
    Rep "3 * D('0.1')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = 1@ / CDec(3)
    Rep "1@ / CDec(3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = 1# / CDec(3)
    Rep "1# / CDec(3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) / 3#
    Rep "CDec(1) / 3#", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) / 3!
    Rep "CDec(1) / 3!", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = "1" + CDec(2)
    Rep "'1' + CDec(2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(7) \ CDec(2)
    Rep "CDec(7) \ CDec(2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(-7) \ 2
    Rep "CDec(-7) \ 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("7.5") \ 2
    Rep "D('7.5') \ 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(7) Mod CDec(2)
    Rep "CDec(7) Mod CDec(2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("7.5") Mod 2
    Rep "D('7.5') Mod 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("-7.5") Mod 2
    Rep "D('-7.5') Mod 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(2) ^ 2
    Rep "CDec(2) ^ 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(2) ^ D("0.5")
    Rep "CDec(2) ^ D('0.5')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(5) And CDec(3)
    Rep "CDec(5) And CDec(3)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Not CDec(5)
    Rep "Not CDec(5)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() \ 2
    Rep "MaxD() \ 2", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = -CDec(5)
    Rep "-CDec(5)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = -D("0")
    Rep "-D('0')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("-0") * 1
    Rep "D('-0') * 1", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Abs(D("-5.5"))
    Rep "Abs(D('-5.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Int(D("-5.5"))
    Rep "Int(D('-5.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Fix(D("-5.5"))
    Rep "Fix(D('-5.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Int(D("5.5"))
    Rep "Int(D('5.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Sgn(D("-5.5"))
    Rep "Sgn(D('-5.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Sgn(CDec(0))
    Rep "Sgn(CDec(0))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Round(D("2.5"))
    Rep "Round(D('2.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Round(D("3.5"))
    Rep "Round(D('3.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Round(D("2.345"), 2)
    Rep "Round(D('2.345'), 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Round(D("2.355"), 2)
    Rep "Round(D('2.355'), 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Round(D("-2.5"))
    Rep "Round(D('-2.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Round(CDec(1) / CDec(3), 5)
    Rep "Round(CDec(1) / CDec(3), 5)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Round(MaxD(), 2)
    Rep "Round(MaxD(), 2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Round(D("1.23456789012345678901234567"), 27)
    Rep "Round(D('1.23456789012345678901234567'), 27)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CInt(D("2.5"))
    Rep "CInt(D('2.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CInt(D("3.5"))
    Rep "CInt(D('3.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CInt(D("32767.5"))
    Rep "CInt(D('32767.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CLng(D("2147483647.4"))
    Rep "CLng(D('2147483647.4'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CLng(D("2147483647.5"))
    Rep "CLng(D('2147483647.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDbl(CDec(1) / CDec(3))
    Rep "CDbl(CDec(1) / CDec(3))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CSng(CDec(1) / CDec(3))
    Rep "CSng(CDec(1) / CDec(3))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CCur(CDec(1) / CDec(3))
    Rep "CCur(CDec(1) / CDec(3))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CCur(D("0.00005"))
    Rep "CCur(D('0.00005'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CCur(D("0.00015"))
    Rep "CCur(D('0.00015'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CCur(MaxD())
    Rep "CCur(MaxD())", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CByte(D("255.4"))
    Rep "CByte(D('255.4'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CByte(D("255.5"))
    Rep "CByte(D('255.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CBool(CDec(0))
    Rep "CBool(CDec(0))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CBool(D("0.1"))
    Rep "CBool(D('0.1'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CStr(CDec(1) / CDec(3))
    Rep "CStr(CDec(1) / CDec(3))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CStr(D("1.10"))
    Rep "CStr(D('1.10'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CStr(D("-0.001"))
    Rep "CStr(D('-0.001'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CStr(MaxD())
    Rep "CStr(MaxD())", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CStr(D("1E-28"))
    Rep "CStr(D('1E-28'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Str(D("1.5"))
    Rep "Str(D('1.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Str(D("-1.5"))
    Rep "Str(D('-1.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Str(CDec(1) / CDec(3))
    Rep "Str(CDec(1) / CDec(3))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Val(CStr(D("1.5")))
    Rep "Val(CStr(D('1.5')))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = VarType(Val("1.5"))
    Rep "VarType(Val('1.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDate(CDec(36526))
    Rep "CDate(CDec(36526))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDate(D("36526.5"))
    Rep "CDate(D('36526.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Hex(CDec(255))
    Rep "Hex(CDec(255))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Hex(D("4294967296"))
    Rep "Hex(D('4294967296'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Oct(CDec(8))
    Rep "Oct(CDec(8))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Hex(D("255.5"))
    Rep "Hex(D('255.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Format(CDec(1) / CDec(3), "0.00000")
    Rep "Format(CDec(1) / CDec(3), '0.00000')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Format(D("1234567.891"), "#,##0.00")
    Rep "Format(D('1234567.891'), '#,##0.00')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Format(CDec(1) / CDec(3))
    Rep "Format(CDec(1) / CDec(3))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Format(MaxD())
    Rep "Format(MaxD())", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Format(D("0.5"), "0%")
    Rep "Format(D('0.5'), '0%')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Format(D("1234.5"), "Currency")
    Rep "Format(D('1234.5'), 'Currency')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Format(MaxD(), "0.00")
    Rep "Format(MaxD(), '0.00')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Format(D("12345678901234567890.12345"), "0.0000")
    Rep "Format(D('12345678901234567890.12345'), '0.0000')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = FormatNumber(CDec(1) / CDec(3), 4)
    Rep "FormatNumber(CDec(1) / CDec(3), 4)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = Len(CStr(CDec(1) / CDec(3)))
    Rep "Len(CStr(CDec(1) / CDec(3)))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = IsNumeric(D("1.5"))
    Rep "IsNumeric(D('1.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(CDec(1) / CDec(3))
    Rep "CDec(CDec(1) / CDec(3))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CVar(D("1.5"))
    Rep "CVar(D('1.5'))", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) / CDec(3) = 1 / 3
    Rep "CDec(1) / CDec(3) = 1 / 3", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) / CDec(3) < 1 / 3
    Rep "CDec(1) / CDec(3) < 1 / 3", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) = 1
    Rep "CDec(1) = 1", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.1") = 0.1
    Rep "D('0.1') = 0.1", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.1") = 0.1!
    Rep "D('0.1') = 0.1!", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1.10") = D("1.1")
    Rep "D('1.10') = D('1.1')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) < CDec(2)
    Rep "CDec(1) < CDec(2)", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(2) > 1.5!
    Rep "CDec(2) > 1.5!", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("0.1") + D("0.2") = D("0.3")
    Rep "D('0.1') + D('0.2') = D('0.3')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) = "1"
    Rep "CDec(1) = '1'", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(1) = True
    Rep "CDec(1) = True", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(-1) = True
    Rep "CDec(-1) = True", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() > 1E+28
    Rep "MaxD() > 1E+28", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = MaxD() = MaxD()
    Rep "MaxD() = MaxD()", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = StrComp(CStr(CDec(1)), "1")
    Rep "StrComp(CStr(CDec(1)), '1')", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = D("1E-28") > 0
    Rep "D('1E-28') > 0", rv, Err.Number
    Err.Clear: rv = "(none)"
    rv = CDec(5) = 5@
    Rep "CDec(5) = 5@", rv, Err.Number
    ' Variant variables, For, Select Case, arrays, passing
    Dim v As Variant, n As Long, a(2) As Variant, dd As Double, cc As Currency
    Err.Clear: n = 0
    For v = CDec(0) To CDec(1) Step D("0.25"): n = n + 1: Next
    Rep "For Decimal 0 To 1 Step 0.25", n & " " & CStr(v) & " " & TypeName(v), Err.Number
    Err.Clear: n = 0
    For v = CDec(1) To CDec(0) Step D("-0.3"): n = n + 1: Next
    Rep "For Decimal 1 To 0 Step -0.3", n & " " & CStr(v) & " " & TypeName(v), Err.Number
    Err.Clear: v = D("2.5")
    Select Case v
    Case 2.5: Rep "Select Case 2.5", "yes", Err.Number
    Case Else: Rep "Select Case 2.5", "no", Err.Number
    End Select
    Err.Clear: a(0) = CDec(1): a(1) = a(0) / 3: a(2) = a(1) * 3
    Rep "array a(2)", a(2), Err.Number
    Err.Clear: dd = CDec(1) / CDec(3)
    Rep "Double = Decimal", dd, Err.Number
    Err.Clear: cc = CDec(1) / CDec(3)
    Rep "Currency = Decimal", cc, Err.Number
    Err.Clear: v = CDec(1): v = v + 0.5
    Rep "v = v + 0.5", v, Err.Number
    Err.Clear: v = CDec(1): v = v * 1000000000: v = v * 1000000000: v = v * 1000000000: v = v * 10
    Rep "v * 1e28", v, Err.Number
    Err.Clear: v = CDec(1): v = v * 1000000000: v = v * 1000000000: v = v * 1000000000: v = v * 1000
    Rep "v * 1e30", v, Err.Number
    Err.Clear: Rep "Twice(D(1.25))", Twice(D("1.25")), Err.Number
    Err.Clear: Rep "Half(D(1.25))", Half(D("1.25")), Err.Number
    ' Print # and Write # of a Decimal
    Print #mF, "Print #:"; CDec(1) / CDec(3); D("-2.5"); MaxD()
    Write #mF, CDec(1) / CDec(3), D("-2.5")
    ' a Decimal in a binary file
    Dim g As Integer, w As Variant, b As Byte, k As Long, hx As String
    g = FreeFile
    Open App.Path & "\dec.bin" For Binary As #g
    Put #g, , D("-123.456")
    Put #g, , MaxD()
    Rep "LOF dec.bin", LOF(g), Err.Number
    hx = ""
    For k = 1 To LOF(g): Get #g, k, b: hx = hx & Right$("0" & Hex$(b), 2): Next
    Rep "bytes", hx, Err.Number
    Get #g, 1, w
    Rep "Get #1", w, Err.Number
    Rep "Get #1 type", TypeName(w), Err.Number
    Get #g, , w
    Rep "Get #2", w, Err.Number
    Close #g
    Kill App.Path & "\dec.bin"
End Sub
