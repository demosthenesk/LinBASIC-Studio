Attribute VB_Name = "ModPrint"
Option Explicit
Public Sub T_Print(f As Integer)
    Print #f, 1; 2; 3
    Print #f, -1; -2; "x"; 3
    Print #f, 1, 2, "three", 4.5
    Print #f, "a"; Tab(10); "b"; Tab(5); "c"
    Print #f, "x"; Spc(4); "y"
    Print #f, True; False; " "; 1.5; 1.5!; 2.25@
    Print #f, "no newline";
    Print #f, " continues"
    Print #f,
    Print #f, "after empty"
    Print #f, 123456789; 1234567890123#; 0.1; 1E-20; 1.5E+300
    Print #f, CSng(123456.7); CSng(1234567.8); CSng(0.000123)
    Print #f, 100000 * 100000#; 12345678901234567#
    Print #f, "end"
End Sub
