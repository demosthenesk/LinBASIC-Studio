Attribute VB_Name = "ModProcs"
Option Explicit
Private Sub AddOne(n As Integer)
    n = n + 1
End Sub

Private Sub AddOneVal(ByVal n As Integer)
    n = n + 1
End Sub

Private Function Greet(Optional ByVal who As String = "World", Optional ByVal times As Integer = 1) As String
    Dim i As Integer
    For i = 1 To times
        Greet = Greet & "Hello " & who & "! "
    Next
End Function

Private Function Has(Optional v As Variant) As Boolean
    Has = Not IsMissing(v)
End Function

Private Function Fib(ByVal n As Long) As Long
    If n < 2 Then Fib = n Else Fib = Fib(n - 1) + Fib(n - 2)
End Function

Private Function Counter() As Long
    Static c As Long
    c = c + 1
    Counter = c
End Function

Private Sub Swap(a As String, b As String)
    Dim t As String
    t = a: a = b: b = t
End Sub

Private Function Twice(x As Variant) As Variant
    x = x * 2
    Twice = x
End Function

Public Sub T_Procs(f As Integer)
    Dim i As Integer, a As String, b As String, v As Variant
    i = 1
    AddOne i
    Print #f, "byref"; i
    AddOne (i)
    Print #f, "paren"; i
    AddOneVal i
    Print #f, "byval"; i
    Call AddOne(i)
    Print #f, "call"; i
    Print #f, Greet(); Greet("VB", 2); Greet(times:=2)
    Print #f, Has(); Has(1)
    Print #f, "fib"; Fib(20)
    Print #f, "static"; Counter(); Counter(); Counter()
    a = "left": b = "right"
    Swap a, b
    Print #f, a; " "; b
    v = 21
    Print #f, Twice(v); v
End Sub
