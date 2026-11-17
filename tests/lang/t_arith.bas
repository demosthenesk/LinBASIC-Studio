Attribute VB_Name = "ModArith"
Option Explicit
Public Sub T_Arith(f As Integer)
    Dim i As Integer, j As Integer, l As Long, s As Single, d As Double, c As Currency, b As Byte
    i = 7: j = 2
    Print #f, "int ops"; i + j; i - j; i * j; i / j; i \ j; i Mod j; i ^ j
    Print #f, "neg mod"; -7 \ 2; -7 Mod 2; 7 Mod -2; -7 Mod -2
    Print #f, "precedence"; 2 + 3 * 4; (2 + 3) * 4; -2 ^ 2; 2 ^ -1; 10 - 4 - 3; 2 ^ 3 ^ 2
    Print #f, "types "; TypeName(i + j); " "; TypeName(i / j); " "; TypeName(i \ j); " "; TypeName(i * 1&); " "; TypeName(i ^ j)
    s = 1.5
    Print #f, "single"; s / 2; TypeName(s / i); TypeName(s * l); TypeName(s + 1)
    l = 100000
    Print #f, "long"; l * 3; l \ 7; l Mod 7; TypeName(l + i)
    d = 1 / 3
    Print #f, "double"; d; d * 3; 1E+20 * 3; 2.5E-10
    c = 1.2345
    Print #f, "currency"; c; c * 3; c / 3; TypeName(c * 2); 10@ / 3@
    b = 100
    Print #f, "byte"; b + 50; TypeName(b + 1); TypeName(b + b)
    Print #f, "cint"; CInt(2.5); CInt(3.5); CInt(-2.5); CInt(2.4999); CLng(1234567.5); CLng(-0.5)
    Print #f, "fix int"; Fix(-2.7); Int(-2.7); Fix(2.7); Int(2.7); Round(2.5); Round(3.5); Round(2.345, 2)
    Print #f, "abs sgn"; Abs(-5); Abs(-2.5); Sgn(-3); Sgn(0); Sgn(4.2)
    Print #f, "math"; Sqr(16); Int(Exp(1) * 1000); Int(Log(10) * 1000); Int(Atn(1) * 4000)
    Print #f, "assign rounding";
    i = 2.5: Print #f, i;
    i = 3.5: Print #f, i;
    l = -4.5: Print #f, l
    Print #f, "bool math"; True + 1; True * 2; False - 1; Not 5; Not True; 5 And 3; 5 Or 3; 5 Xor 3
    Print #f, "compare"; 1 < 2; 2 = 2; 3 <> 3; "a" < "b"; "B" < "a"
End Sub
