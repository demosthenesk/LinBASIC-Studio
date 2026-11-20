VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "AW5 MDI"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    WriteLog "mdi load " & Forms.Count
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    WriteLog "mdi unload"
End Sub
