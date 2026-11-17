Attribute VB_Name = "ModConvert"
Option Explicit
Public Sub T_Convert(f As Integer)
    Print #f, CStr(1.5); "|"; CStr(0.5); "|"; CStr(-0.25); "|"; CStr(100); "|"; CStr(True); "|"; CStr(1E+15); "|"; CStr(1E+16)
    Print #f, CStr(123456789012345#); "|"; CStr(1234567890123456#); "|"; CStr(0.0001); "|"; CStr(0.00001); "|"; CStr(1.5E-07)
    Print #f, Str(1.5); "|"; Str(-1.5); "|"; Str(0.5); "|"; Str(42); "|"; Str(-42)
    Print #f, CStr(CSng(1) / 3); "|"; CStr(1 / 3); "|"; CStr(2 / 3); "|"; CStr(CSng(2) / 3)
    Print #f, Val("12abc"); Val("  1 2 3"); Val("&H10"); Val("&O10"); Val("1.5e2"); Val("abc"); Val("-.5")
    Print #f, CDbl("3.25"); CInt("12"); CLng("  70000  "); CBool("True"); CBool(0); CBool(-3)
    Print #f, Hex$(255); " "; Hex(-1); " "; Hex(65535); " "; Hex(-1&); " "; Oct$(8); " "; Oct(-1)
    Print #f, CCur(1.23456); CCur("2.5"); CByte(255.4)
    Print #f, 1E+300 * 10 > 0
End Sub
