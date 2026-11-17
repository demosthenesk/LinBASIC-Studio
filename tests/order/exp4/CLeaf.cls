VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLeaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function IM(ByVal n As Integer) As Integer
    Order4.mLog = Order4.mLog & n
    IM = Choose(n, 40, 3, 2, 1)
End Function
Public Function LM(ByVal n As Integer) As Long
    Order4.mLog = Order4.mLog & n
    LM = Choose(n, 40, 3, 2, 1)
End Function
Public Function DM(ByVal n As Integer) As Double
    Order4.mLog = Order4.mLog & n
    DM = Choose(n, 40, 3, 2, 1)
End Function
Public Function CM(ByVal n As Integer) As Currency
    Order4.mLog = Order4.mLog & n
    CM = Choose(n, 40, 3, 2, 1)
End Function
Public Function VM(ByVal n As Integer) As Variant
    Order4.mLog = Order4.mLog & n
    VM = Choose(n, 40, 3, 2, 1)
End Function
Public Function SM(ByVal n As Integer) As String
    Order4.mLog = Order4.mLog & n
    SM = Choose(n, "40", "3", "2", "1")
End Function
Public Function BM(ByVal n As Integer) As Boolean
    Order4.mLog = Order4.mLog & n
    BM = Choose(n, True, False, True, False)
End Function
