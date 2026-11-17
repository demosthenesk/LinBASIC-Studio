VERSION 5.00
Begin VB.Form frmSize 
   Caption         =   "Size"
   ClientHeight    =   2000
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2000
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "frmSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub RunTests()
    LogBorders Me, "size"
    Me.WindowState = vbMaximized
    WriteLog "max " & Me.WindowState & " " & (Me.ScaleWidth > 3000) & " " & (Me.ScaleHeight > 2000)
    Me.WindowState = vbNormal
    WriteLog "normal " & Me.WindowState & " " & Me.ScaleWidth & " " & Me.ScaleHeight
    Me.Height = Me.Height + 300
    WriteLog "taller " & Me.ScaleHeight
End Sub

Private Sub Form_Resize()
    WriteLog "size resize " & Me.WindowState
End Sub
