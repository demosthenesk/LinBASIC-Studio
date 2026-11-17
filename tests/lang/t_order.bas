Attribute VB_Name = "ModOrder"
Option Explicit
' VB evaluates the parts of an expression and the arguments of a call from left to right
Private mLog As String
Private mN As Long

Private Function Lg(ByVal s As String) As String
    mLog = mLog & s
    Lg = s
End Function

Private Function N(ByVal v As Integer) As Integer
    mN = mN * 10 + v
    N = v
End Function

Private Function Bump() As Long
    mN = mN + 1
    Bump = mN
End Function

Private Function Pair(ByVal a As Integer, ByVal b As Integer) As String
    Pair = a & "," & b
End Function

Private Function Three(ByVal a As String, b As String, ByVal c As String) As String
    Three = a & b & c
End Function

Public Sub T_Order(f As Integer)
    Dim s As String, x As Long, v As Variant, c As Collection
    mLog = ""
    s = Lg("a") & Lg("b") & Lg("c")
    Print #f, "concat "; s; " "; mLog
    mLog = ""
    s = Lg("a") & (Lg("b") & Lg("c"))
    Print #f, "paren "; s; " "; mLog
    mN = 0
    x = N(1) + N(2) * N(3)
    Print #f, "arith "; x; mN
    mN = 0
    x = N(4) - N(5)
    Print #f, "sub "; x; mN
    mN = 0
    s = Pair(N(1), N(2))
    Print #f, "args "; s; mN
    mN = 0
    mLog = ""
    s = Three(Lg("p"), Lg("q"), Lg("r"))
    Print #f, "three "; s; " "; mLog
    mN = 0
    v = Bump() & "-" & Bump() & "-" & Bump()
    Print #f, "bump "; v
    mN = 0
    Print #f, "cmp "; (Bump() < Bump()); mN
    mN = 0
    Print #f, "print "; Bump(); Bump(); Bump()
    mLog = ""
    If Lg("x") = Lg("y") Or Lg("z") = "z" Then Print #f, "if "; mLog
    mN = 0
    Set c = New Collection
    c.Add Bump(), CStr(Bump())
    Print #f, "late "; c(1); c.Count; mN
    mN = 0
    v = Array(Bump(), Bump(), Bump())
    Print #f, "array "; v(0); v(1); v(2)
    mLog = ""
    Print #f, "mid "; Mid$(Lg("hello"), Len(Lg("ab")), Len(Lg("abc")))
    Print #f, "midlog "; mLog
End Sub
