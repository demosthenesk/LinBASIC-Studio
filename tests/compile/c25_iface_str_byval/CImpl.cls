VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IFace
Private mV As Variant, mO As Collection, mS As String
Private Property Get IFace_S() As String
    IFace_S = mS
End Property
Private Property Let IFace_S(ByVal RHS As String)
    Log "let s"
    mS = RHS
End Property
