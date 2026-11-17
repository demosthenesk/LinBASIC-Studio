VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVarImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IVarIface
Private mLevel As Long, mLabel As String, mPayload As Variant, mOwner As Collection
Public Log As String

Private Property Get IVarIface_Level() As Long
    Log = Log & "GetLevel;"
    IVarIface_Level = mLevel
End Property
Private Property Let IVarIface_Level(ByVal RHS As Long)
    Log = Log & "LetLevel" & RHS & ";"
    mLevel = RHS * 10
End Property
Private Property Get IVarIface_Label() As String
    IVarIface_Label = "[" & mLabel & "]"
End Property
Private Property Let IVarIface_Label(ByVal RHS As String)
    mLabel = RHS
End Property
Private Property Get IVarIface_Payload() As Variant
    If IsObject(mPayload) Then Set IVarIface_Payload = mPayload Else IVarIface_Payload = mPayload
End Property
Private Property Let IVarIface_Payload(RHS As Variant)
    Log = Log & "LetPayload;"
    mPayload = RHS
End Property
Private Property Set IVarIface_Payload(RHS As Variant)
    Log = Log & "SetPayload;"
    Set mPayload = RHS
End Property
Private Property Get IVarIface_Owner() As Collection
    Set IVarIface_Owner = mOwner
End Property
Private Property Set IVarIface_Owner(ByVal RHS As Collection)
    Log = Log & "SetOwner;"
    Set mOwner = RHS
End Property
Private Function IVarIface_Describe() As String
    IVarIface_Describe = "CVarImpl:" & mLevel & ":" & mLabel
End Function
