VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' GlobalMultiUse: its members are used without an object
Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function

Public Function TakeTrace() As String
    TakeTrace = gTrace
    gTrace = ""
End Function

Public Property Get LibName() As String
    LibName = App.Title
End Property
