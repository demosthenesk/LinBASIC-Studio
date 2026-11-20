VERSION 5.00
Begin VB.Form frmChild 
   Caption         =   "AW5 Document"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1800
   ScaleWidth      =   3000
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    WriteLog "child load " & Forms.Count
End Sub
