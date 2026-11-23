VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1200
   ClientLeft      =   1500
   ClientTop       =   1500
   ClientWidth     =   2400
   LinkTopic       =   "Form2"
   ScaleHeight     =   1200
   ScaleWidth      =   2400
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    WriteLog "Form2 Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "Form2 Unload"
End Sub
