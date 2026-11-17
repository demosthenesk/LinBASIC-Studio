VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse"
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
      Height          =   1200
      Left            =   2400
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   240
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1500
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
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const LEFTDOWN = &H2, LEFTUP = &H4, RIGHTDOWN = &H8, RIGHTUP = &H10

Private mStep As Integer
Private mMoved As Boolean

Private Sub Pump()
    Dim t As Single
    t = Timer
    Do
        DoEvents
    Loop While Timer - t < 0.3 And Timer >= t
End Sub

Private Sub ClickAt(ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal downFlag As Long, ByVal upFlag As Long, ByVal times As Integer)
    Dim p As POINTAPI, i As Integer
    p.X = X
    p.Y = Y
    ClientToScreen hWnd, p
    SetCursorPos p.X, p.Y
    Pump
    For i = 1 To times
        mouse_event downFlag, 0, 0, 0, 0
        mouse_event upFlag, 0, 0, 0, 0
    Next
    Pump
End Sub

Private Sub Form_Load()
    StartLog
    WriteLog "load " & (Me.hWnd <> 0) & " " & (Command1.hWnd <> 0) & " " & (Picture1.hWnd <> Command1.hWnd)
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "form mousedown " & Button & " " & Shift & " " & X & " " & Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "form mouseup " & Button & " " & Shift & " " & X & " " & Y
End Sub

Private Sub Form_Click()
    WriteLog "form click"
End Sub

Private Sub Form_DblClick()
    WriteLog "form dblclick"
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "command1 mousedown " & Button & " " & Shift & " " & X & " " & Y
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "command1 mouseup " & Button & " " & Shift & " " & X & " " & Y
End Sub

Private Sub Command1_Click()
    WriteLog "command1 click"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "picture1 mousedown " & Button & " " & Shift & " " & X & " " & Y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "picture1 mouseup " & Button & " " & Shift & " " & X & " " & Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mMoved Then WriteLog "picture1 mousemove " & Button & " " & X & " " & Y
    mMoved = True
End Sub

Private Sub Picture1_Click()
    WriteLog "picture1 click"
End Sub

Private Sub Picture1_DblClick()
    WriteLog "picture1 dblclick"
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "label1 mousedown " & Button & " " & Shift & " " & X & " " & Y
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "label1 mouseup " & Button & " " & Shift & " " & X & " " & Y
End Sub

Private Sub Label1_Click()
    WriteLog "label1 click"
End Sub

Private Sub Label1_DblClick()
    WriteLog "label1 dblclick"
End Sub

Private Sub Timer1_Timer()
    Dim p As POINTAPI, q As POINTAPI
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        ClickAt Command1.hWnd, 40, 16, LEFTDOWN, LEFTUP, 1
    Case 2
        mMoved = False
        ClickAt Picture1.hWnd, 10, 5, LEFTDOWN, LEFTUP, 2
    Case 3
        ClickAt Me.hWnd, 200, 180, RIGHTDOWN, RIGHTUP, 1
    Case 4
        ClickAt Me.hWnd, 26, 85, LEFTDOWN, LEFTUP, 2
    Case 5
        q.X = 26
        q.Y = 85
        ClientToScreen Me.hWnd, q
        GetCursorPos p
        WriteLog "cursor " & (p.X - q.X) & " " & (p.Y - q.Y)
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub
