VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   720
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form2.frx":0000
      Top             =   1440
      Width           =   2400
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "Form2.frx":0007
      Left            =   1680
      List            =   "Form2.frx":0017
      TabIndex        =   1
      Top             =   120
      Width           =   1800
   End
   Begin VB.PictureBox Picture1 
      Height          =   960
      Left            =   120
      Picture         =   "Form2.frx":002C
      ScaleHeight     =   900
      ScaleWidth      =   1380
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
