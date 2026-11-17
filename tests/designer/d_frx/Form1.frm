VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":038A
   ScaleHeight     =   6000
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   480
      Left            =   5760
      TabIndex        =   10
      Tag             =   $"Form1.frx":09CC
      ToolTipText     =   "Short tip"
      Top             =   2640
      Width           =   1440
   End
   Begin VB.TextBox Text3 
      Height          =   720
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   $"Form1.frx":0AD4
      Top             =   2640
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   720
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0C04
      Top             =   2640
      Width           =   2400
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0C19
      Left            =   3960
      List            =   "Form1.frx":0C2A
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1800
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   2040
      List            =   "Form1.frx":0C34
      TabIndex        =   4
      Top             =   1200
      Width           =   1800
   End
   Begin VB.ListBox List1 
      DragIcon        =   "Form1.frx":0C49
      Height          =   1035
      ItemData        =   "Form1.frx":0FD3
      Left            =   120
      List            =   "Form1.frx":0FE3
      TabIndex        =   3
      Top             =   1200
      Width           =   1800
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   4920
      MouseIcon       =   "Form1.frx":0FFB
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pic"
      Height          =   720
      Left            =   3360
      Picture         =   "Form1.frx":1385
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1440
   End
   Begin VB.PictureBox Picture1 
      Height          =   960
      Left            =   120
      Picture         =   "Form1.frx":170F
      ScaleHeight     =   900
      ScaleWidth      =   1380
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":1D51
      Height          =   1200
      Left            =   2640
      TabIndex        =   9
      Top             =   3600
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Form1.frx":1F85
      Height          =   600
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   2400
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   2520
      Picture         =   "Form1.frx":1F93
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   1680
      Picture         =   "Form1.frx":249F
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
