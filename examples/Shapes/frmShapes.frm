VERSION 5.00
Begin VB.Form frmShapes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shapes"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMove 
      Interval        =   30
      Left            =   6480
      Top             =   4560
   End
   Begin VB.HScrollBar hsbSpeed 
      Height          =   255
      LargeChange     =   2
      Left            =   1080
      Max             =   10
      Min             =   1
      TabIndex        =   2
      Top             =   4680
      Value           =   4
      Width           =   2655
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   4620
      Width           =   1215
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00402000&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   735
   End
End
Attribute VB_Name = "frmShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Shapes - graphics on a PictureBox: Line, Circle, PSet, Print, FillStyle and a Timer animation
Option Explicit

Private Type Ball
    X As Single
    Y As Single
    DX As Single
    DY As Single
    R As Single
    Color As Long
End Type

Private Balls(1 To 6) As Ball

Private Sub Form_Load()
    Dim i As Integer
    Randomize 7
    For i = 1 To 6
        With Balls(i)
            .R = 12 + Rnd * 18
            .X = .R + Rnd * (picCanvas.ScaleWidth - 2 * .R)
            .Y = .R + Rnd * (picCanvas.ScaleHeight - 2 * .R)
            .DX = IIf(Rnd > 0.5, 1, -1) * (0.5 + Rnd)
            .DY = IIf(Rnd > 0.5, 1, -1) * (0.5 + Rnd)
            .Color = QBColor(9 + i)
        End With
    Next
    Draw
End Sub

Private Sub Draw()
    Dim i As Integer, x As Single
    picCanvas.Cls
    picCanvas.FillStyle = vbFSTransparent
    For x = 0 To picCanvas.ScaleWidth Step 20       ' a grid
        picCanvas.Line (x, 0)-(x, picCanvas.ScaleHeight), RGB(0, 64, 128)
    Next
    For x = 0 To picCanvas.ScaleHeight Step 20
        picCanvas.Line (0, x)-(picCanvas.ScaleWidth, x), RGB(0, 64, 128)
    Next
    picCanvas.Line (20, 20)-(140, 80), RGB(255, 200, 0), B   ' a box, a filled box and an ellipse
    picCanvas.Line (160, 20)-(280, 80), RGB(0, 160, 80), BF
    picCanvas.Circle (360, 50), 60, RGB(255, 255, 255), , , 0.5
    picCanvas.FillStyle = vbFSSolid
    For i = 1 To 6
        With Balls(i)
            picCanvas.FillColor = .Color
            picCanvas.Circle (.X, .Y), .R, vbWhite
        End With
    Next
    picCanvas.CurrentX = 8: picCanvas.CurrentY = picCanvas.ScaleHeight - 20
    picCanvas.ForeColor = vbWhite
    picCanvas.Print "LinBASIC Studio - " & Format$(Now, "hh:nn:ss")
End Sub

Private Sub tmrMove_Timer()
    Dim i As Integer, s As Single
    s = hsbSpeed.Value
    For i = 1 To 6
        With Balls(i)
            .X = .X + .DX * s
            .Y = .Y + .DY * s
            If .X < .R Or .X > picCanvas.ScaleWidth - .R Then .DX = -.DX
            If .Y < .R Or .Y > picCanvas.ScaleHeight - .R Then .DY = -.DY
        End With
    Next
    Draw
End Sub

Private Sub cmdPause_Click()
    tmrMove.Enabled = Not tmrMove.Enabled
    cmdPause.Caption = IIf(tmrMove.Enabled, "&Pause", "&Go")
End Sub
