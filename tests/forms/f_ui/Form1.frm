VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "UI tests"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4200
      Top             =   3000
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   2640
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1800
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "Form1.frx":0018
      Left            =   2640
      List            =   "Form1.frx":0000
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "first"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mStep As Integer

Private Sub Form_Load()
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_ui.txt"
    On Error GoTo 0
    WriteLog "load " & Me.KeyPreview & " " & Me.FontName & " " & Me.FontSize & " " & Me.FontBold
    WriteLog "text2 font " & Text2.FontName & " " & Text2.FontSize & " " & Text2.FontBold & " " & Text2.Font.Weight & " " & Text2.Font.Charset
    WriteLog "colors " & Hex$(Me.BackColor) & " " & Hex$(Text1.BackColor) & " " & Hex$(Text1.ForeColor) & " " & Hex$(Text2.BackColor) & " " & Hex$(Text2.ForeColor) & " " & Hex$(Command1.BackColor)
    WriteLog "list " & List1.ListCount & " " & List1.List(0) & " " & List1.List(2) & " " & List1.ItemData(1) & " " & List1.ItemData(2)
    WriteLog "menu " & mnuFile.Caption & " " & mnuOpen.Checked & " " & mnuExit.Enabled & " " & mnuSep.Caption & " " & mnuFile.Visible & " " & TypeName(mnuOpen)
    WriteLog "picture " & Picture1.ScaleMode & " " & Picture1.ScaleWidth & " " & Picture1.ScaleHeight & " " & Picture1.AutoRedraw & " " & Hex$(Picture1.BackColor)
    Picture1.PSet (10, 20), vbRed
    WriteLog "pset " & Picture1.Point(10, 20) & " " & Picture1.Point(11, 20) & " " & Picture1.CurrentX & " " & Picture1.CurrentY
    Picture1.Line (30, 30)-(39, 39), vbBlue, BF
    WriteLog "line " & Picture1.Point(35, 35) & " " & Picture1.Point(40, 40) & " " & Picture1.CurrentX & " " & Picture1.CurrentY
    Picture1.Line (50, 50)-(59, 50), RGB(0, 128, 0)
    WriteLog "line2 " & Picture1.Point(55, 50) & " " & Picture1.Point(59, 50) & " " & Picture1.Point(55, 51)
    Picture1.Cls
    WriteLog "cls " & Picture1.Point(10, 20) & " " & Picture1.CurrentX & " " & Picture1.CurrentY
    Picture1.ScaleMode = vbTwips
    WriteLog "twips " & Picture1.ScaleWidth & " " & Picture1.ScaleHeight
    Picture1.Scale (0, 0)-(10, 10)
    WriteLog "user scale " & Picture1.ScaleMode & " " & Picture1.ScaleWidth & " " & Picture1.ScaleHeight & " " & Picture1.ScaleLeft
    WriteLog "rgb " & RGB(1, 2, 3) & " " & QBColor(12) & " " & vbRed & " " & vbButtonFace
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    WriteLog "activate " & (Me.ActiveControl Is Text1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "form keydown " & KeyCode & " " & Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    WriteLog "form keypress " & KeyAscii
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    WriteLog "form keyup " & KeyCode & " " & Shift
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

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "text2 keydown " & KeyCode & " " & Shift
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    WriteLog "text2 keypress " & KeyAscii
    If KeyAscii = Asc("b") Then KeyAscii = Asc("B")
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    WriteLog "text2 keyup " & KeyCode & " " & Shift
End Sub

Private Sub Text2_Change()
    WriteLog "text2 change [" & Text2.Text & "]"
End Sub

Private Sub Command1_GotFocus()
    WriteLog "command1 gotfocus"
End Sub

Private Sub Command1_Click()
    WriteLog "command1 click"
End Sub

Private Sub mnuOpen_Click()
    WriteLog "menu open click"
End Sub

Private Sub Timer1_Timer()
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        Text2.SetFocus
        WriteLog "after setfocus " & (Me.ActiveControl Is Text2)
    Case 2
        SendKeys "ab", True
        WriteLog "text2 [" & Text2.Text & "] " & Text2.SelStart
        SendKeys "{BACKSPACE}{ENTER}", True
        WriteLog "text2 [" & Text2.Text & "]"
    Case 3
        Command1.SetFocus
        SendKeys " ", True
        WriteLog "after space"
    Case 4
        mnuOpen.Checked = Not mnuOpen.Checked
        mnuExit.Enabled = False
        WriteLog "menu " & mnuOpen.Checked & " " & mnuExit.Enabled
        Me.BackColor = vbRed
        Text1.FontBold = True
        Text1.FontSize = 12
        WriteLog "changed " & Hex$(Me.BackColor) & " " & Text1.Font.Bold & " " & Text1.FontSize
        Text1.Font.Size = 10
        Text1.Font.Name = "Arial"
        WriteLog "font " & Text1.FontSize & " " & Text1.FontName
        mnuOpen_Click
    Case 5
        Timer1.Enabled = False
        Unload Me
    End Select
End Sub
