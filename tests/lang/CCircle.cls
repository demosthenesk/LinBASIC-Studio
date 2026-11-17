VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IShape

Public Radius As Double

Private Function IShape_Area() As Double
    IShape_Area = 3 * Radius * Radius
End Function

Private Property Get IShape_Name() As String
    IShape_Name = "circle"
End Property

Private Sub IShape_Grow(ByVal k As Double)
    Radius = Radius * k
End Sub
