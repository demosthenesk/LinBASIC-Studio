VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Misc"
   ClientHeight    =   3600
   ClientLeft      =   1500
   ClientTop       =   1500
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   0  'Manual
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4200
      Top             =   3000
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   900
      Left            =   240
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   4
      Top             =   2160
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Target"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      CausesValidation=   0   'False
      Caption         =   "NoVal"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mnuPopA 
         Caption         =   "Item A"
      End
      Begin VB.Menu mnuPopB 
         Caption         =   "Item B"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const LEFTDOWN = &H2, LEFTUP = &H4

Private WithEvents mText As TextBox
Private mStep As Integer

Private Sub Pump()
    Dim t As Single
    t = Timer
    Do
        DoEvents
    Loop While Timer - t < 0.3 And Timer >= t
End Sub

Private Sub MoveTo(ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long)
    Dim p As POINTAPI
    p.X = X
    p.Y = Y
    ClientToScreen hWnd, p
    SetCursorPos p.X, p.Y
    Pump
End Sub

Private Sub Form_Load()
    StartLog
    WriteLog "load " & Text1.CausesValidation & " " & Command1.CausesValidation & " " & Picture1.DragMode & " " & Command2.DragMode & " " & Controls.Count & " " & Me.MousePointer
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    WriteLog "activate " & Me.ActiveControl.Name
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    WriteLog "text1 validate [" & Text1.Text & "]"
    If Text1.Text = "" Then Cancel = True
End Sub

Private Sub Text1_GotFocus()
    WriteLog "text1 gotfocus"
End Sub

Private Sub Text1_LostFocus()
    WriteLog "text1 lostfocus"
End Sub

Private Sub Text2_GotFocus()
    WriteLog "text2 gotfocus"
End Sub

Private Sub Text2_LostFocus()
    WriteLog "text2 lostfocus"
End Sub

Private Sub Command1_GotFocus()
    WriteLog "command1 gotfocus"
End Sub

Private Sub mnuPop_Click()
    WriteLog "pop click"
End Sub

Private Sub mnuPopA_Click()
    WriteLog "pop A click"
End Sub

Private Sub mnuPopB_Click()
    WriteLog "pop B click"
End Sub

Private Sub mText_Change()
    WriteLog "mtext change [" & mText.Text & "]"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "picture1 mousedown"
End Sub

Private Sub Picture1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    WriteLog "picture1 dragover " & Source.Name & " " & State
End Sub

Private Sub Command2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    WriteLog "command2 dragover " & Source.Name & " " & X & " " & Y & " " & State
End Sub

Private Sub Command2_DragDrop(Source As Control, X As Single, Y As Single)
    WriteLog "command2 dragdrop " & Source.Name & " " & X & " " & Y
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    WriteLog "form dragover " & Source.Name & " " & State
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    WriteLog "form dragdrop " & Source.Name & " " & X & " " & Y
End Sub

Private Sub Timer1_Timer()
    Dim c As Control
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        Text2.SetFocus
        WriteLog "after setfocus " & Me.ActiveControl.Name
    Case 2
        WriteLog "now " & Me.ActiveControl.Name
        Command1.SetFocus
        WriteLog "after setfocus " & Me.ActiveControl.Name
    Case 3
        Text1.SetFocus
        WriteLog "after setfocus " & Me.ActiveControl.Name
    Case 4
        Text1.Text = "ok"
        SendKeys "{TAB}", True
        WriteLog "after tab " & Me.ActiveControl.Name
    Case 5
        Text1.SetFocus
        WriteLog "after setfocus " & Me.ActiveControl.Name
    Case 6
        Text1.Text = ""
        On Error Resume Next
        Me.ValidateControls
        WriteLog "validatecontrols " & Err.Number & " " & Me.ActiveControl.Name
        Err.Clear
        On Error GoTo 0
        Text1.Text = "x"
    Case 7
        SendKeys "{DOWN}{ENTER}"
        PopupMenu mnuPop
        WriteLog "after popup"
        SendKeys "{ESC}"
        PopupMenu mnuPop
        WriteLog "after popup 2"
    Case 8
        Set mText = Controls.Add("VB.TextBox", "txtNew")
        WriteLog "added " & TypeName(mText) & " " & mText.Name & " " & mText.Visible & " " & mText.Left & " " & mText.Top & " " & mText.Width & " " & mText.Height & " [" & mText.Text & "] " & Controls.Count & " " & mText.TabIndex
        mText.Text = "abc"
        Set c = Controls.Add("VB.CommandButton", "cmdNew", Picture1)
        WriteLog "added2 " & c.Name & " " & c.Container.Name & " [" & c.Caption & "] " & Controls.Count
        Set mText = Nothing
        Set c = Nothing
        Controls.Remove "txtNew"
        Controls.Remove "cmdNew"
        WriteLog "removed " & Controls.Count
        On Error Resume Next
        Controls.Remove "Text1"
        WriteLog "err " & Err.Number
        Err.Clear
        On Error GoTo 0
    Case 9
        Command1.MousePointer = vbHourglass
        WriteLog "pointer " & Command1.MousePointer & " " & Me.MousePointer
        Command1.MousePointer = vbDefault
        Set Me.Icon = LoadPicture(App.Path & "\test.ico")
        WriteLog "icon " & Me.Icon.Type & " " & Me.Icon.Width & " " & Me.Icon.Height
    Case 10
        MoveTo Picture1.hWnd, 40, 30
        mouse_event LEFTDOWN, 0, 0, 0, 0
        Pump
        MoveTo Command2.hWnd, 40, 16
        mouse_event LEFTUP, 0, 0, 0, 0
        Pump
        WriteLog "drag done"
    Case 11
        MoveTo Me.hWnd, 300, 200
        Picture1.Drag vbBeginDrag
        WriteLog "drag begun"
        MoveTo Command2.hWnd, 20, 10
        Picture1.Drag vbEndDrag
        WriteLog "drag ended"
        Pump
    Case 12
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub
