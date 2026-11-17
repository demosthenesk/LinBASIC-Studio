VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' MultiUse: the objects of all the clients come from one copy of the server
Private mN As Long

Private Sub Class_Initialize()
    gObjects = gObjects + 1
    mN = gObjects
End Sub

Public Function Info() As String
    Info = "multi n=" & mN & " objects=" & gObjects & " mains=" & gMains & " total=" & gTotal
End Function

Public Sub Add(ByVal v As Long)
    gTotal = gTotal + v
End Sub
