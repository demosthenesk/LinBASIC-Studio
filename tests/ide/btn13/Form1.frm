VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "A wide button text"
      Height          =   720
      Left            =   240
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   3000
   End
   Begin VB.CommandButton Command2 
      Height          =   960
      Left            =   1800
      Picture         =   "Form1.frx":0C42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   1080
      Left            =   240
      Picture         =   "Form1.frx":1884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
