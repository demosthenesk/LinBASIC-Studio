VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mSrc As CSource
Attribute mSrc.VB_VarHelpID = -1
Public Tag As String

Public Sub Attach(s As CSource)
    Set mSrc = s
End Sub

Public Sub Detach()
    Set mSrc = Nothing
End Sub

Private Sub mSrc_Tick(ByVal n As Long, Cancel As Boolean)
    gEv = gEv & Tag & "t" & n & ";"
    If n = 3 Then Cancel = True
End Sub

Private Sub mSrc_Done(ByVal total As Long)
    gEv = gEv & Tag & "done" & total & ";"
End Sub
