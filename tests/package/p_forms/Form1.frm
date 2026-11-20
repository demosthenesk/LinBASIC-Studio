VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2000
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2000
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Output As #f
    Print #f, "forms ok " & Command1.Caption
    Close #f
    Unload Me
End Sub
