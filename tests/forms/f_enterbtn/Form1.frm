VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Enter on a button"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Other"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
' Phase 14 issue (Discord, MudCalc): Enter in a TextBox whose KeyPress sends {TAB} moves the focus to a
' CommandButton; Enter on the focused button then fires its Click (the button's KeyPress also sends {TAB})

Private mStep As Integer

Private Sub Form_Load()
    On Error Resume Next
    Kill App.Path & "\expected\f_enterbtn.txt"
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
Private Sub Text2_LostFocus()
    WriteLog "text2 lostfocus"
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "text2 keydown " & KeyCode
End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    WriteLog "text2 keyup " & KeyCode
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    WriteLog "text2 keypress " & KeyAscii
    If KeyAscii = 13 Then                ' Enter = Tab
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub Command1_GotFocus()
    WriteLog "command1 gotfocus"
End Sub
Private Sub Command1_LostFocus()
    WriteLog "command1 lostfocus"
End Sub
Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "command1 keydown " & KeyCode
End Sub
Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
    WriteLog "command1 keyup " & KeyCode
End Sub
Private Sub Command1_KeyPress(KeyAscii As Integer)
    WriteLog "command1 keypress " & KeyAscii
    If KeyAscii = 13 Then                ' Enter = Tab
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub Command1_Click()
    WriteLog "command1 click [" & Text2.Text & "]"
End Sub
Private Sub Text3_GotFocus()
    WriteLog "text3 gotfocus"
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    WriteLog "text3 keypress " & KeyAscii
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "text3 keydown " & KeyCode
End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
    WriteLog "text3 keyup " & KeyCode
End Sub
Private Sub Text3_LostFocus()
    WriteLog "text3 lostfocus"
End Sub
Private Sub Command3_Click()
    WriteLog "command3 click (default)"
End Sub
Private Sub Command3_GotFocus()
    WriteLog "command3 gotfocus"
End Sub
Private Sub Command4_Click()
    WriteLog "command4 click (cancel)"
End Sub
Private Sub Command4_GotFocus()
    WriteLog "command4 gotfocus"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "form keydown " & KeyCode
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    WriteLog "form keypress " & KeyAscii
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    WriteLog "form keyup " & KeyCode
End Sub
Private Sub Command2_GotFocus()
    WriteLog "command2 gotfocus"
End Sub
Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "command2 keydown " & KeyCode
End Sub
Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
    WriteLog "command2 keyup " & KeyCode
End Sub
Private Sub Command2_KeyPress(KeyAscii As Integer)
    WriteLog "command2 keypress " & KeyAscii
End Sub
Private Sub Command2_Click()
    WriteLog "command2 click"
End Sub

Private Sub Timer1_Timer()
    mStep = mStep + 1
    WriteLog "step " & mStep & " active " & AC()
    Select Case mStep
    Case 1
        Text2.SetFocus
    Case 2
        SendKeys "100{ENTER}"
    Case 4
        SendKeys "{ENTER}"
    Case 6
        Command2.SetFocus
    Case 7
        SendKeys "{ENTER}"
    Case 9
        Command2.SetFocus
    Case 10
        SendKeys "~"
    Case 12
        Command3.Default = True
        Command4.Cancel = True
        Text3.SetFocus
    Case 13
        SendKeys "{ENTER}"
    Case 15
        SendKeys "{ESC}"
    Case 17
        Command1.SetFocus
    Case 18
        SendKeys "{ENTER}"
    Case 20
        Me.KeyPreview = True
        Text3.SetFocus
    Case 21
        SendKeys "{ENTER}"
    Case 23
        SendKeys "x{ESC}"
    Case 25
        Text2.SetFocus
    Case 26
        SendKeys "{ENTER}"
    Case 28
        Command2.SetFocus
    Case 29
        SendKeys "{ENTER}"
    Case 31
        Timer1.Enabled = False
        Unload Me
    End Select
End Sub
