VERSION 5.00
Begin VB.UserControl Edit 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ScaleHeight     =   495
   ScaleWidth      =   1935
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' a text box in a user control: its focus events and the user control's

Private mId As String

Private Sub UserControl_Initialize()
    WriteLog "  edit initialize"
End Sub

Private Sub UserControl_InitProperties()
    mId = Extender.Name
    WriteLog "  edit initproperties " & mId
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mId = Extender.Name
    txt.Text = PropBag.ReadProperty("Text", "")
    WriteLog "  edit readproperties " & mId & " [" & txt.Text & "]"
End Sub

Private Sub UserControl_Resize()
    WriteLog "  edit resize " & mId & " " & ScaleWidth & "x" & ScaleHeight
End Sub

Private Sub UserControl_Show()
    WriteLog "  edit show " & mId
End Sub

Private Sub UserControl_Hide()
    WriteLog "  edit hide " & mId
End Sub

Private Sub UserControl_EnterFocus()
    WriteLog "  edit enterfocus " & mId
End Sub

Private Sub UserControl_ExitFocus()
    WriteLog "  edit exitfocus " & mId
End Sub

Private Sub UserControl_GotFocus()
    WriteLog "  edit gotfocus " & mId
End Sub

Private Sub UserControl_LostFocus()
    WriteLog "  edit lostfocus " & mId
End Sub

Private Sub txt_GotFocus()
    WriteLog "  edit txt gotfocus " & mId
End Sub

Private Sub txt_LostFocus()
    WriteLog "  edit txt lostfocus " & mId
End Sub

Private Sub UserControl_Terminate()
    WriteLog "  edit terminate " & mId
End Sub

Public Property Get Text() As String
    Text = txt.Text
End Property

Public Property Let Text(ByVal s As String)
    txt.Text = s
    PropertyChanged "Text"
End Property
