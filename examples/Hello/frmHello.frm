VERSION 5.00
Begin VB.Form frmHello 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hello, LinBASIC Studio"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClock 
      Interval        =   1000
      Left            =   4560
      Top             =   2520
   End
   Begin VB.Frame fraGreeting 
      Caption         =   "Greeting"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2655
      Begin VB.CheckBox chkCapitals 
         Caption         =   "In CAPITALS"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optGreeting 
         Caption         =   "Good morning"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optGreeting 
         Caption         =   "Hello"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdSay 
      Caption         =   "&Say it"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "World"
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblClock 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lblName 
      Caption         =   "Your &name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   400
      Width           =   1095
   End
End
Attribute VB_Name = "frmHello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Hello - the first program: a form, its controls and their events
Option Explicit

Private Sub Form_Load()
    tmrClock_Timer
End Sub

Private Sub cmdSay_Click()
    Dim msg As String
    If Trim$(txtName.Text) = "" Then
        MsgBox "Please type your name first.", vbExclamation, App.Title
        txtName.SetFocus
        Exit Sub
    End If
    msg = optGreeting(IIf(optGreeting(0).Value, 0, 1)).Caption & ", " & txtName.Text & "!"
    If chkCapitals.Value = vbChecked Then msg = UCase$(msg)
    MsgBox msg, vbInformation, App.Title
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub tmrClock_Timer()
    lblClock.Caption = Format$(Now, "hh:nn:ss")
End Sub
