VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help Form2"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   HelpContextID   =   500
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "WhatsThisMode"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      WhatsThisHelpID =   700
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "What's This 600"
      Top             =   240
      WhatsThisHelpID =   600
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    WriteLog "form2 " & Me.WhatsThisHelp & " " & Me.WhatsThisButton & " " & Text1.WhatsThisHelpID & " " & Command1.WhatsThisHelpID & " " & Me.HelpContextID
End Sub
