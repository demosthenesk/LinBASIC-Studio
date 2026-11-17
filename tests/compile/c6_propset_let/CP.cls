VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m As StdFont
Private Sub Class_Initialize()
    Set m = New StdFont
    m.Name = "Arial"
End Sub
Public Property Get Fnt() As StdFont
    Set Fnt = m
End Property
Public Property Set Fnt(ByVal v As StdFont)
    Set m = v
End Property
Public Sub Inner()
    Fnt = "Times New Roman"
End Sub
