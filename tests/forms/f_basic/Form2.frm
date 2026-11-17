VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1500
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Second form"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
    WriteLog "Form2 Initialize"
End Sub

Private Sub Form_Load()
    WriteLog "Form2 Load " & Label1.Caption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WriteLog "Form2 QueryUnload " & UnloadMode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "Form2 Unload"
End Sub

Private Sub Form_Terminate()
    WriteLog "Form2 Terminate"
End Sub
