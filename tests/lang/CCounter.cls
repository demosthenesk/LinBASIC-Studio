VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCount As Long
Private mName As String
Public Tag As Variant

Private Sub Class_Initialize()
    mCount = 0
    gLog = gLog & "[init]"
End Sub

Private Sub Class_Terminate()
    gLog = gLog & "[term " & mName & "]"
End Sub

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal v As String)
    mName = v
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Sub Increment(Optional ByVal by As Long = 1)
    mCount = mCount + by
End Sub

Public Function Twice() As Long
    Twice = mCount * 2
End Function

Public Function Self() As CCounter
    Set Self = Me
End Function
