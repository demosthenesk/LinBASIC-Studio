VERSION 5.00
Begin VB.UserControl Ghost 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   480
End
Attribute VB_Name = "Ghost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' invisible at run time, like a Timer

Private Sub UserControl_Show()
    WriteLog "  ghost show"
End Sub

Private Sub UserControl_Hide()
    WriteLog "  ghost hide"
End Sub

Public Function Hello() As String
    Hello = "boo"
End Function
