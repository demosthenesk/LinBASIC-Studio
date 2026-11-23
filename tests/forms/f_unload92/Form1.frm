VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1200
   ClientLeft      =   1500
   ClientTop       =   1500
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   2400
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    WriteLog "Form1 Load"
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "Form1 Unload"
End Sub
