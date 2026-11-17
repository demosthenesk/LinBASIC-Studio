VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Tick(ByVal n As Long, Cancel As Boolean)
Public Event Done(ByVal total As Long)
Public Event Unused()

Public Function Run(ByVal count As Long) As Long
    Dim i As Long, c As Boolean, n As Long
    For i = 1 To count
        RaiseEvent Tick(i, c)
        If c Then Exit For
        n = n + 1
    Next
    RaiseEvent Done(n)
    Run = n
End Function
