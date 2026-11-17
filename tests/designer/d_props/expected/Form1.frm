VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Δοκιμή ιδιοτήτων"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Enabled         =   0   'False
      Height          =   480
      Left            =   5160
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox txtName 
      Height          =   360
      Index           =   0
      Left            =   5160
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   4320
      Width           =   1800
   End
   Begin VB.FileListBox File1 
      Height          =   1140
      Hidden          =   -1  'True
      Left            =   3240
      Pattern         =   "*.txt"
      ReadOnly        =   0   'False
      TabIndex        =   14
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group"
      ForeColor       =   &H8000000D&
      Height          =   1440
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   3000
      Begin VB.OptionButton Option2 
         Caption         =   "Inside"
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2400
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   1440
      Left            =   120
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   118
      TabIndex        =   11
      Top             =   2640
      Width           =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5040
      Top             =   1200
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1200
      Left            =   4560
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   10
      Left            =   2040
      Max             =   100
      Min             =   1
      SmallChange     =   2
      TabIndex        =   9
      Top             =   1200
      Value           =   50
      Width           =   2400
   End
   Begin VB.ListBox List1 
      Columns         =   2
      Height          =   1200
      IntegralHeight  =   0   'False
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   1200
      Width           =   1800
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5880
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   1800
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   360
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Value           =   -1  'True
      Width           =   1800
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   360
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Value           =   2  'Grayed
      Width           =   1800
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Value           =   1  'Checked
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   480
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   480
      Left            =   3960
      MousePointer    =   11  'Hourglass
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "abc"
      ToolTipText     =   "tip"
      Top             =   120
      Width           =   1800
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      DrawMode        =   7  'Invert
      X1              =   3600
      X2              =   5040
      Y1              =   2640
      Y2              =   3360
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   1440
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "He said ""hi"""
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1800
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "&Recent"
         Begin VB.Menu mnuItem 
            Caption         =   "Item 0"
            Index           =   0
         End
         Begin VB.Menu mnuItem 
            Caption         =   "Item 1"
            Index           =   1
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   %{BKSP}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Checked         =   -1  'True
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuHidden 
         Caption         =   "&Hidden"
         HelpContextID   =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRun 
         Caption         =   "Run"
         Shortcut        =   +^{F5}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
