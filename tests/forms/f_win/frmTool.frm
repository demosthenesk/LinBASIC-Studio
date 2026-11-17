VERSION 5.00
Begin VB.Form frmTool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tool"
   ClientHeight    =   2000
   ClientLeft      =   1500
   ClientTop       =   1500
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2000
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "frmTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPaints As Integer
Private mLogPaint As Boolean

Private Sub Form_Load()
    WriteLog "tool load " & Me.AutoRedraw
End Sub

Private Sub Form_Resize()
    WriteLog "tool resize " & Me.ScaleWidth & " " & Me.ScaleHeight
End Sub

Private Sub Form_Activate()
    WriteLog "tool activate"
End Sub

Private Sub Form_Deactivate()
    WriteLog "tool deactivate"
End Sub

Private Sub Form_Paint()
    mPaints = mPaints + 1
    If mPaints = 1 Then WriteLog "tool first paint"
    If mLogPaint Then WriteLog "tool paint"
End Sub

Public Sub RunTests()
    Dim o As Object
    Set o = Me
    LogBorders Me, "tool"
    WriteLog "tool pos " & Me.Left & " " & Me.Top
    On Error Resume Next
    o.BorderStyle = 2
    WriteLog "err " & Err.Number
    Err.Clear
    o.MinButton = True
    WriteLog "err " & Err.Number
    Err.Clear
    On Error GoTo 0
    Me.Width = 4500
    WriteLog "width " & Me.Width & " " & Me.ScaleWidth
    Me.Move 600, 600, 5000, 3500
    WriteLog "moved " & Me.Left & " " & Me.Top & " " & Me.Width & " " & Me.Height & " " & Me.ScaleWidth & " " & Me.ScaleHeight
    mLogPaint = True
    Me.Refresh
    mLogPaint = False
    WriteLog "after refresh"
    Me.WindowState = vbMinimized
    WriteLog "state " & Me.WindowState
    Me.WindowState = vbNormal
    WriteLog "state " & Me.WindowState & " " & Me.ScaleWidth & " " & Me.ScaleHeight
    Me.ScaleMode = vbPixels
    WriteLog "pixels " & Me.ScaleWidth & " " & Me.ScaleHeight & " " & Me.Width
End Sub
