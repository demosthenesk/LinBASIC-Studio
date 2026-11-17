VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Phase 6.8"
   ClientHeight    =   3600
   ClientLeft      =   1500
   ClientTop       =   1500
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   0  'Manual
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   2400
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Label"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   2000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   240
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Phase 6.8: For Each over the form, Form.Count, the default properties of Menu, Shape and Line,
' Set PictureBox = picture (its default property), members a control does not have

Private Sub Form_Load()
    Dim c As Control, s As String, v As Variant, n As Integer
    StartLog
    WriteLog "count " & Count & " " & Me.Count & " " & Controls.Count
    On Error Resume Next
    For Each c In Me
        n = n + 1
        s = n & " " & TypeName(c) & " " & c.Name
        Err.Clear
        v = c
        If Err.Number Then s = s & " def=err" & Err.Number Else s = s & " def=" & CStr(v)
        Err.Clear
        c.BackColor = vbRed
        s = s & " back=" & Err.Number
        Err.Clear
        s = s & " left=" & c.Left
        s = s & " lefterr=" & Err.Number
        Err.Clear
        WriteLog s
    Next
    WriteLog "loop " & n & " " & Err.Number
    On Error GoTo 0
    mnuOpen = False
    WriteLog "menu " & mnuOpen.Enabled & " " & mnuOpen.Caption & " " & mnuOpen
    mnuOpen = True
    Shape1 = 2
    WriteLog "shape " & Shape1.Shape & " " & Shape1
    Line1 = False
    WriteLog "line " & Line1.Visible & " " & Line1
    Set Picture1 = LoadPicture()
    WriteLog "pic " & Picture1.Picture.Handle & " " & TypeName(Picture1.Picture)
    Unload Me
End Sub
