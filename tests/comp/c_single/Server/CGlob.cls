VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGlob"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' GlobalSingleUse: a global object, in a copy of the server of its own
Public Function GInfo() As String
    gObjects = gObjects + 1
    GInfo = "glob objects=" & gObjects & " mains=" & gMains & " total=" & gTotal
End Function
