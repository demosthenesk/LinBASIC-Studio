Attribute VB_Name = "ModVariant"
Option Explicit
Public Sub T_Variant(f As Integer)
    Dim v As Variant, w As Variant, e As Variant
    Print #f, "empty"; IsEmpty(e); VarType(e); TypeName(e); "["; e; "]"; e + 1; "["; e & "x"; "]"
    v = 32767
    Print #f, TypeName(v); v + 1; TypeName(v + 1)
    v = 2147483647
    Print #f, TypeName(v); TypeName(v + 1)
    v = "10": w = 5
    Print #f, v + w; TypeName(v + w); v & w; v * w
    v = "abc": w = "def"
    Print #f, v + w; v < w
    v = Null
    Print #f, IsNull(v); IsNull(v + 1); IsNull(v & "x"); "["; v & "x"; "]"; VarType(v)
    v = 1.5: w = 2
    Print #f, v * w; TypeName(v * w); v / w; v \ w; v Mod w
    v = True
    Print #f, v; TypeName(v); v + 1; TypeName(v + 1)
    v = CByte(200)
    Print #f, TypeName(v); v + v; TypeName(v + v)
    v = 1: w = "1"
    Print #f, v = w; TypeName(CVar(3.5)); VarType(CVar("x")); VarType(CVar(1&))
    v = 5
    If v Then Print #f, "truthy"
    v = 0
    If Not v Then Print #f, "not zero"
    Print #f, IIf(3 > 2, "yes", "no"); IIf(False, 1, 2.5)
End Sub
