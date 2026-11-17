Attribute VB_Name = "OrderTest"
Option Explicit
' In which order does VB6 evaluate the operands of an operator and the arguments of a call?
' Every helper appends its number to mLog; each line shows the order the calls ran in.
Private mLog As String

Private Function NI(ByVal v As Integer) As Integer
    mLog = mLog & v
    NI = v
End Function
Private Function NL(ByVal v As Long) As Long
    mLog = mLog & v
    NL = v
End Function
Private Function ND(ByVal v As Double) As Double
    mLog = mLog & v
    ND = v
End Function
Private Function NC(ByVal v As Currency) As Currency
    mLog = mLog & v
    NC = v
End Function
Private Function NV(ByVal v As Variant) As Variant
    mLog = mLog & v
    NV = v
End Function
Private Function NS(ByVal v As String) As String
    mLog = mLog & v
    NS = v
End Function
Private Function NB(ByVal v As Integer) As Boolean
    mLog = mLog & v
    NB = (v Mod 2) = 1
End Function
Private Function Three(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Three = a + b + c
End Function
Private Function ThreeRef(a As Long, b As Long, c As Long) As Long
    ThreeRef = a + b + c
End Function

Private Function InIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0                           ' Debug.Print runs only in the IDE
    InIDE = Err.Number <> 0
    Err.Clear
End Function

Private Sub Show(f As Integer, ByVal what As String)
    Print #f, what; " "; mLog
    mLog = ""
End Sub

Sub Main()
    Dim f As Integer, x As Long, d As Double, v As Variant, s As String, b As Boolean, c As Currency
    Dim a(1 To 9) As Long, i As Long, k As Long
    f = FreeFile
    Open App.Path & "\" & App.EXEName & IIf(InIDE(), "-IDE", "") & ".txt" For Output As #f
    mLog = ""
    ' Integer
    x = NI(1) + NI(2): Show f, "int +"
    x = NI(1) - NI(2): Show f, "int -"
    x = NI(1) * NI(2): Show f, "int *"
    d = NI(1) / NI(2): Show f, "int /"
    x = NI(1) \ NI(2): Show f, "int \"
    x = NI(1) Mod NI(2): Show f, "int Mod"
    d = NI(1) ^ NI(2): Show f, "int ^"
    x = NI(1) + NI(2) * NI(3): Show f, "int 1+2*3"
    x = NI(1) * NI(2) + NI(3): Show f, "int 1*2+3"
    x = NI(1) + NI(2) + NI(3): Show f, "int 1+2+3"
    x = (NI(1) + NI(2)) + (NI(3) + NI(4)): Show f, "int (1+2)+(3+4)"
    x = NI(1) + 5 + NI(2): Show f, "int 1+5+2"
    k = 5: x = k + NI(1) + NI(2): Show f, "int k+1+2"
    x = -NI(1) + NI(2): Show f, "int -1+2"
    ' Long, Double, Currency, Variant
    x = NL(1) + NL(2): Show f, "long +"
    x = NL(1) - NL(2): Show f, "long -"
    x = NL(1) * NL(2): Show f, "long *"
    x = NL(1) + NL(2) * NL(3): Show f, "long 1+2*3"
    d = ND(1) + ND(2): Show f, "double +"
    d = ND(1) - ND(2): Show f, "double -"
    d = ND(1) * ND(2): Show f, "double *"
    d = ND(1) / ND(2): Show f, "double /"
    d = ND(1) + ND(2) * ND(3): Show f, "double 1+2*3"
    c = NC(1) + NC(2): Show f, "currency +"
    v = NV(1) + NV(2): Show f, "variant +"
    v = NV(1) - NV(2): Show f, "variant -"
    v = NV(1) + NV(2) * NV(3): Show f, "variant 1+2*3"
    v = NV(1) & NV(2): Show f, "variant &"
    ' mixed types
    d = NI(1) + ND(2): Show f, "int+double"
    d = ND(1) + NI(2): Show f, "double+int"
    x = NI(1) + NL(2): Show f, "int+long"
    v = NI(1) + NV(2): Show f, "int+variant"
    ' strings
    s = NS("1") & NS("2"): Show f, "string &"
    s = NS("1") & NS("2") & NS("3"): Show f, "string 1&2&3"
    s = NS("1") + NS("2"): Show f, "string +"
    s = NS("1") & (NS("2") & NS("3")): Show f, "string 1&(2&3)"
    ' comparisons and logic
    b = NI(1) < NI(2): Show f, "int <"
    b = NI(1) = NI(2): Show f, "int ="
    b = ND(1) > ND(2): Show f, "double >"
    b = NS("1") = NS("2"): Show f, "string ="
    b = NV(1) = NV(2): Show f, "variant ="
    x = NI(1) And NI(2): Show f, "int And"
    x = NI(1) Or NI(2): Show f, "int Or"
    b = NB(1) And NB(2): Show f, "bool And"
    b = NB(1) Or NB(2): Show f, "bool Or"
    If NB(1) And NB(2) Then x = 0
    Show f, "If And"
    If NI(1) + NI(2) = NI(3) Then x = 0
    Show f, "If 1+2=3"
    ' calls
    x = Three(NL(1), NL(2), NL(3)): Show f, "call ByVal"
    x = Three(NL(1) + NL(2), NL(3), NL(4)): Show f, "call 1+2,3,4"
    i = 1: k = 2: x = ThreeRef(NL(1), i, NL(3)): Show f, "call ByRef"
    x = NL(1) + Three(NL(2), NL(3), NL(4)): Show f, "1+call(2,3,4)"
    x = Three(NL(1), NL(2), NL(3)) + NL(4): Show f, "call(1,2,3)+4"
    v = IIf(NB(1), NV(2), NV(3)): Show f, "IIf"
    s = Mid$(NS("abc"), NL(2), NL(1)): Show f, "Mid$"
    x = InStr(NL(1), NS("x"), NS("y")): Show f, "InStr"
    ' assignment: the index or the value first?
    a(NL(1)) = NL(2): Show f, "a(1) = 2"
    a(NL(1)) = NL(2) + NL(3): Show f, "a(1) = 2+3"
    Print #f, "print"; NL(1); NL(2); NL(3)
    Show f, "print items"
    Close #f
    MsgBox "Done: " & App.EXEName & IIf(InIDE(), "-IDE", "") & ".txt"
End Sub
