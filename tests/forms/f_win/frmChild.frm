VERSION 5.00
Begin VB.Form frmChild 
   Caption         =   "frmChild"
   ClientHeight    =   2000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2000
   ScaleWidth      =   3000
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mId As Integer

Private Sub Form_Load()
    gChildCount = gChildCount + 1
    mId = gChildCount
    WriteLog "child load " & mId & " [" & Me.Tag & "] " & Me.Visible
End Sub

Private Sub Form_Activate()
    WriteLog "child activate " & mId
End Sub

Private Sub Form_Deactivate()
    WriteLog "child deactivate " & mId
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WriteLog "child queryunload " & mId & " " & UnloadMode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "child unload " & mId
End Sub

Private Sub Form_Terminate()
    WriteLog "child terminate " & mId
End Sub
