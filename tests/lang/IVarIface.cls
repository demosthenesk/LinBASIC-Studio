VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVarIface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' an interface with Public variables: the implementing class has Property Get / Let / Set for them
Public Level As Long
Public Label As String
Public Payload As Variant
Public Owner As Collection

Public Function Describe() As String
    Describe = "IVarIface:" & Level & ":" & Label
End Function
