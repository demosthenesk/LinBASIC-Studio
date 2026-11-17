VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form1
   Caption         =   "CoolBar 7"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   400
      Left            =   6600
      Top             =   3000
   End
   Begin ComCtl3.CoolBar CoolBar1
      Align           =   1  'Align Top
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   1931
      BandCount       =   3
      _CBWidth        =   7200
      _CBHeight       =   1095
      _Version        =   "6.7.9839"
      Child1          =   "Text1"
      MinHeight1      =   330
      Width1          =   3000
      NewRow1         =   0   'False
      Child2          =   "Text2"
      MinHeight2      =   330
      Width2          =   2505
      NewRow2         =   0   'False
      Caption3        =   "Three"
      Child3          =   "Text3"
      MinHeight3      =   360
      NewRow3         =   -1  'True
      Begin VB.TextBox Text3
         Height          =   360
         Left            =   720
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox Text2
         Height          =   330
         Left            =   3240
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox Text1
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Label Label1
      Caption         =   "Do not touch the mouse while this test runs."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' CoolBar 7 (LinBASIC Studio phase 7.4): moving a band to another row with the mouse, each case from the same bands
' (Bands.Clear, then Add): straight down / up, with the cursor moving left or right, in one step or in four, into a
' row of one band or of two, below the CoolBar (a new row)

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

Private mStep As Integer

Private Sub Pump()
    Dim i As Integer
    For i = 1 To 20
        DoEvents
    Next
End Sub

Private Sub Drag(ByVal c As Control, ByVal dx As Long, ByVal dy As Long, ByVal steps As Integer)
    Dim i As Integer, r As RECT, X As Long, Y As Long
    GetWindowRect c.hwnd, r
    X = r.Left - 6
    Y = (r.Top + r.Bottom) \ 2
    SetCursorPos X, Y
    Pump
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    Pump
    For i = 1 To steps
        SetCursorPos X + dx * i \ steps, Y + dy * i \ steps
        Pump
    Next
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Pump
End Sub

' the bands again: widths in pixels (-1: no Width), "/" begins a new row: "200 167 / 0" = Text1 Text2 | Text3
Private Sub Bands(ByVal spec As String)
    Dim parts() As String, i As Integer, nr As Boolean, n As Integer, b As Object
    CoolBar1.Bands.Clear
    parts = Split(spec, " ")
    For i = 0 To UBound(parts)
        If parts(i) = "/" Then
            nr = True
        Else
            n = n + 1
            Set b = CoolBar1.Bands.Add(, , , , nr)
            b.MinHeight = 330
            If Val(parts(i)) > 0 Then b.Width = Val(parts(i)) * 15
            Set b.Child = Choose(n, Text1, Text2, Text3)
            nr = False
        End If
    Next
End Sub

Private Function G(ByVal c As Control) As String
    G = c.Name & "(" & c.Left \ 15 & "," & c.Top \ 15 & "," & c.Width \ 15 & ")"
End Function

Private Sub State(ByVal what As String)
    Dim i As Integer, s As String
    s = what & ": h=" & CoolBar1.Height \ 15 & " rows=" & CoolBar1.RowCount & " " & G(Text1) & " " & G(Text2) & " " & G(Text3) & " |"
    For i = 1 To CoolBar1.Bands.Count
        With CoolBar1.Bands(i)
            s = s & " " & .Child.Name & ":pos=" & .Position & ",w=" & .Width \ 15 & ",nr=" & .NewRow
        End With
    Next
    WriteLog s
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

' each case: the bands, the band to move (1 2 3), dx, dy (pixels), steps
Private Sub Case1(ByVal name As String, ByVal spec As String, ByVal which As Integer, ByVal dx As Long, ByVal dy As Long, ByVal steps As Integer)
    Bands spec
    Pump
    State name & " before"
    Drag Choose(which, Text1, Text2, Text3), dx, dy, steps
End Sub

Private Sub Timer1_Timer()
    Static last As String
    Timer1.Enabled = False
    On Error Resume Next
    If last <> "" Then State last & " after"
    mStep = mStep + 1
    Select Case mStep
    Case 1: last = "s1 B down 1 step": Case1 last, "200 167 / 0", 2, 0, 36, 1
    Case 2: last = "s2 B down 4 steps": Case1 last, "200 167 / 0", 2, 0, 36, 4
    Case 3: last = "s3 B down, left 100": Case1 last, "200 167 / 0", 2, -100, 36, 4
    Case 4: last = "s4 B down, right 100": Case1 last, "200 167 / 0", 2, 100, 36, 4
    Case 5: last = "s5 B down, left 100, 1 step": Case1 last, "200 167 / 0", 2, -100, 36, 1
    Case 6: last = "s6 C up 1 step": Case1 last, "200 167 / 300", 3, 0, -24, 1
    Case 7: last = "s7 C up 4 steps": Case1 last, "200 167 / 300", 3, 0, -24, 4
    Case 8: last = "s8 C up, right 100": Case1 last, "200 167 / 300", 3, 100, -24, 4
    Case 9: last = "s9 C up, right 250": Case1 last, "200 167 / 300", 3, 250, -24, 4
    Case 10: last = "s10 C up, right 400": Case1 last, "200 167 / 300", 3, 400, -24, 4
    Case 11: last = "s11 C up, right 250, 1 step": Case1 last, "200 167 / 300", 3, 250, -24, 1
    Case 12: last = "s12 C up, no widths": Case1 last, "-1 -1 / -1", 3, 0, -24, 4
    Case 13: last = "s13 B up from row 2 of 3": Case1 last, "200 / 0 / 0", 2, 0, -24, 4
    Case 14: last = "s14 B down to row 3 of 3": Case1 last, "200 / 0 / 0", 2, 0, 24, 4
    Case 15: last = "s15 C down, new row": Case1 last, "100 100 100", 3, 0, 40, 4
    Case 16: last = "s16 B down, new row": Case1 last, "100 100 100", 2, 0, 40, 4
    Case 17: last = "s17 A down, into row 2": Case1 last, "200 167 / 0", 1, 0, 24, 4
    Case 18: last = "s18 A down, right 300": Case1 last, "200 167 / 0", 1, 300, 24, 4
    Case 19: last = "s19 C up above the bar": Case1 last, "200 / 167 / 0", 3, 0, -70, 4
    Case 20: last = "s20 B left 50, then down": Case1 last, "200 167 / 0", 2, -50, 0, 2
    Case 21
        State "s20 continued"
        Drag Text2, 0, 30, 4
        last = "s20 then down"
    Case 22
        Unload Me
        Exit Sub
    End Select
    If Err.Number <> 0 Then WriteLog "  step " & mStep & " ! " & Err.Number & " " & Err.Description: Err.Clear
    Timer1.Enabled = True
End Sub
