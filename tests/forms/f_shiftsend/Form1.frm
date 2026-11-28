VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Shift held and SendKeys"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4200
      Top             =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Discord (MudCalc): Shift+Enter in a TextBox whose KeyPress sends {TAB} moves back, as Shift+Tab: the keys of
' SendKeys go with the Shift the user holds (here keybd_event holds it)

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private mStep As Integer

Private Sub Form_Load()
    On Error Resume Next
    Kill App.Path & "\expected\f_shiftsend.txt"
    Timer1.Enabled = True
End Sub

Private Function AC() As String
    On Error Resume Next
    AC = "?"
    AC = Me.ActiveControl.Name
End Function

Private Sub Text1_GotFocus()
    WriteLog "text1 gotfocus"
End Sub
Private Sub Text2_GotFocus()
    WriteLog "text2 gotfocus"
End Sub
Private Sub Text3_GotFocus()
    WriteLog "text3 gotfocus"
End Sub
Private Sub Command1_GotFocus()
    WriteLog "command1 gotfocus"
End Sub
Private Sub Command1_Click()
    WriteLog "command1 click"
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "text2 keydown " & KeyCode & " shift " & Shift
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    WriteLog "text2 keypress " & KeyAscii
    If KeyAscii = 13 Then                ' Enter = Tab
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "text3 keydown " & KeyCode & " shift " & Shift
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    WriteLog "text3 keypress " & KeyAscii
End Sub

Private Sub Timer1_Timer()
    mStep = mStep + 1
    WriteLog "step " & mStep & " active " & AC()
    Select Case mStep
    Case 1
        Text2.SetFocus
    Case 2
        keybd_event &H10, 0, 0, 0
    Case 3
        SendKeys "{ENTER}"
    Case 5
        keybd_event &H10, 0, 2, 0
    Case 6
        Text3.SetFocus
    Case 7
        keybd_event &H10, 0, 0, 0
    Case 8
        SendKeys "ab"
    Case 10
        keybd_event &H10, 0, 2, 0
        WriteLog "text3 [" & Text3.Text & "]"
    Case 11
        Text2.SetFocus
    Case 12
        SendKeys "{ENTER}"
    Case 14
        Timer1.Enabled = False
        Unload Me
    End Select
End Sub
