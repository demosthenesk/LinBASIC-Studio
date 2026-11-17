VERSION 5.00
Begin VB.UserControl Box 
   AccessKeys      =   "b"
   BackColor       =   &H0080FFFF&
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ControlContainer=   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   2400
End
Attribute VB_Name = "Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' a container: the controls the form puts in it

Private Sub UserControl_Initialize()
    WriteLog "  box initialize"
End Sub

Private Sub UserControl_InitProperties()
    WriteLog "  box initproperties"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    WriteLog "  box readproperties " & Extender.Name
End Sub

Private Sub UserControl_EnterFocus()
    WriteLog "  box enterfocus"
End Sub

Private Sub UserControl_ExitFocus()
    WriteLog "  box exitfocus"
End Sub

Private Sub UserControl_GotFocus()
    WriteLog "  box gotfocus"
End Sub

Private Sub UserControl_LostFocus()
    WriteLog "  box lostfocus"
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    WriteLog "  box accesskeypress " & KeyAscii
End Sub

Private Sub UserControl_Terminate()
    WriteLog "  box terminate"
End Sub

Public Function Inside() As String
    Dim c As Object, s As String
    For Each c In ContainedControls
        s = s & c.Name & " "
    Next
    Inside = ContainedControls.Count & " " & s
End Function
