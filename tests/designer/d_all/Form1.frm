VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   480
      Index           =   1
      Left            =   5040
      TabIndex        =   17
      Top             =   5670
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   480
      Index           =   0
      Left            =   3600
      TabIndex        =   16
      Top             =   5640
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1440
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   3000
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1800
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   3000
      TabIndex        =   12
      Top             =   2880
      Width           =   1200
   End
   Begin VB.DirListBox Dir1 
      Height          =   1200
      Left            =   1560
      TabIndex        =   11
      Top             =   2880
      Width           =   1200
   End
   Begin VB.DriveListBox Drive1 
      Height          =   360
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Left            =   6480
      Top             =   1440
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1200
      Left            =   6000
      TabIndex        =   9
      Top             =   1440
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   4440
      TabIndex        =   8
      Top             =   1440
      Width           =   1200
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   3000
      TabIndex        =   7
      Top             =   1440
      Width           =   1200
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   480
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   480
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   480
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   480
      Left            =   3000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   1200
   End
   Begin VB.PictureBox Picture1 
      Height          =   1200
      Left            =   120
      ScaleHeight     =   1140
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   120
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Line Line1 
      X1              =   6000
      X2              =   7200
      Y1              =   2880
      Y2              =   3360
   End
   Begin VB.Shape Shape1 
      Height          =   1200
      Left            =   4440
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   480
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
