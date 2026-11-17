VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private mName As String, mAge As Variant

Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(ByVal v As String)
    WriteLog "  let name " & v
    mName = v
    PropertyChanged "Name"
End Property
Public Property Get Age() As Variant
    Age = mAge
End Property
Public Property Let Age(ByVal v As Variant)
    WriteLog "  let age " & v
    mAge = v
    PropertyChanged "Age"
End Property
