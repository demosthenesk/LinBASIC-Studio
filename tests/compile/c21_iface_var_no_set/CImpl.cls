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
Private Property Get IFace_V() As Variant
    If IsObject(mV) Then Set IFace_V = mV Else IFace_V = mV
End Property
Private Property Let IFace_V(RHS As Variant)
    Log "let"
    mV = RHS
End Property
