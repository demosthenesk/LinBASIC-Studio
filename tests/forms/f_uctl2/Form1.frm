VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "User controls 2"
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
   Begin UC2Tests.Ghost Ghost1 
      Left            =   3600
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin UC2Tests.Box Box1 
      Height          =   1200
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   2117
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "inside"
         Top             =   120
         Width           =   1215
      End
   End
   Begin UC2Tests.Edit Edit1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Text            =   "abc"
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "first"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' user controls 2 (LinBASIC Studio, phase 4.3): focus (EnterFocus, ExitFocus, the Extender's GotFocus and
' LostFocus), a user control as a container, InvisibleAtRuntime, AccessKeys, Controls.Add of a user control

Private mStep As Integer

Private Sub Form_Load()
    WriteLog "form load"
    WriteLog "ghost " & Ghost1.Hello & " " & Ghost1.Name
    WriteLog "container " & Text2.Container.Name & " " & TypeName(Text2.Container) & " inside " & Box1.Inside
    WriteLog "edit " & Edit1.Text & " " & TypeName(Edit1)
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    WriteLog "form activate " & ActiveControl.Name
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

Private Sub Edit1_GotFocus()
    WriteLog "edit1 gotfocus"
End Sub

Private Sub Edit1_LostFocus()
    WriteLog "edit1 lostfocus"
End Sub

Private Sub Box1_GotFocus()
    WriteLog "box1 gotfocus"
End Sub

Private Sub Box1_LostFocus()
    WriteLog "box1 lostfocus"
End Sub

Private Sub Timer1_Timer()
    Dim c As Control
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep & " active " & ActiveControl.Name
    Select Case mStep
    Case 1
        Edit1.SetFocus
    Case 2
        Text2.SetFocus
    Case 3
        Box1.SetFocus
    Case 4
        Text1.SetFocus
    Case 5
        SendKeys "%b", True
    Case 6
        Set c = Controls.Add("UC2Tests.Edit", "EditX")
        WriteLog "added " & TypeName(c) & " " & c.Visible & " " & Controls.Count
        c.Move 2700, 120, 1935, 495
        c.Text = "new"
        c.Visible = True
    Case 7
        WriteLog "editx " & Controls("EditX").Text
        Controls.Remove "EditX"
        WriteLog "removed " & Controls.Count
        Text1.SetFocus
    Case 8
        Box1.Visible = False
        WriteLog "box hidden " & Text2.Visible
        Box1.Visible = True
        Box1.Move 240, 1440
        WriteLog "box moved " & Text2.Left & " " & Text2.Top
    Case 9
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "form unload"
End Sub
