VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form1
   Caption         =   "CoolBar 5"
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
' CoolBar 5 (LinBASIC Studio phase 7.4): the user moves the bands with the mouse (mouse_event): along the row by the
' gripper, to a new row and back, a band of the second row into the first, with FixedOrder, a double click on a gripper

Private Type POINTAPI
    X As Long
    Y As Long
End Type
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

' the gripper of the band of a child: back pixels left of the child (6: a band with no caption; 41: "Three"), in the middle
Private Sub GripOf(ByVal c As Control, X As Long, Y As Long, ByVal back As Long)
    Dim r As RECT
    GetWindowRect c.hwnd, r
    X = r.Left - back
    Y = (r.Top + r.Bottom) \ 2
End Sub

Private Sub Drag(ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long)
    Dim i As Integer
    SetCursorPos X, Y
    Pump
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    Pump
    For i = 1 To 4
        SetCursorPos X + dx * i \ 4, Y + dy * i \ 4
        Pump
    Next
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Pump
End Sub

Private Sub DblClick(ByVal X As Long, ByVal Y As Long)
    SetCursorPos X, Y
    Pump
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Pump
End Sub

Private Function G(ByVal c As Control) As String
    G = c.Name & "(" & c.Left & "," & c.Top & "," & c.Width & "," & c.Height & ")"
End Function

Private Sub State(ByVal what As String)
    Dim i As Integer, s As String
    s = what & ": h=" & CoolBar1.Height & " rows=" & CoolBar1.RowCount & " " & G(Text1) & " " & G(Text2) & " " & G(Text3)
    WriteLog s
    For i = 1 To CoolBar1.Bands.Count
        With CoolBar1.Bands(i)
            WriteLog "  b" & i & " pos=" & .Position & " w=" & .Width & " h=" & .Height & " nr=" & .NewRow & " child=" & .Child.Name
        End With
    Next
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    WriteLog "  ev heightchanged " & NewHeight
End Sub
Private Sub CoolBar1_Resize()
    WriteLog "  ev resize h=" & CoolBar1.Height
End Sub
Private Sub CoolBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "  ev mousedown " & Button & " " & X & " " & Y
End Sub
Private Sub CoolBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLog "  ev mouseup " & Button & " " & X & " " & Y
End Sub
Private Sub CoolBar1_Click()
    WriteLog "  ev click"
End Sub
Private Sub CoolBar1_DblClick()
    WriteLog "  ev dblclick"
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim X As Long, Y As Long
    Timer1.Enabled = False
    mStep = mStep + 1
    On Error Resume Next
    Select Case mStep
    Case 1
        State "start"
        GripOf Text2, X, Y, 6
        WriteLog "drag band 2 left 40"
        Drag X, Y, -40, 0
    Case 2
        State "band 2 left"
        GripOf Text2, X, Y, 6
        WriteLog "drag band 2 right 60"
        Drag X, Y, 60, 0
    Case 3
        State "band 2 right"
        GripOf Text2, X, Y, 6
        WriteLog "drag band 2 down 40"
        Drag X, Y, 0, 40
    Case 4
        State "band 2 down"
        GripOf Text2, X, Y, 6
        WriteLog "drag band 2 up 25"
        Drag X, Y, 0, -25
    Case 5
        State "band 2 up"
        GripOf Text3, X, Y, 41
        WriteLog "drag band 3 up 25, right 100"
        Drag X, Y, 100, -25
    Case 6
        State "band 3 up"
        GripOf Text1, X, Y, 6
        WriteLog "drag band 1 down 40"
        Drag X, Y, 0, 40
    Case 7
        State "band 1 down"
        CoolBar1.FixedOrder = True
        State "fixed order"
        GripOf Text2, X, Y, 6
        WriteLog "drag band 2 (fixed order) down 40"
        Drag X, Y, 0, 40
    Case 8
        State "fixed order, band 2 down"
        GripOf Text2, X, Y, 6
        WriteLog "drag band 2 (fixed order) up 25"
        Drag X, Y, 0, -25
    Case 9
        State "fixed order, band 2 up"
        CoolBar1.FixedOrder = False
        GripOf Text2, X, Y, 6
        WriteLog "double click on the gripper of band 2"
        DblClick X, Y
    Case 10
        State "double click"
        GripOf Text2, X, Y, 6
        WriteLog "double click again"
        DblClick X, Y
    Case 11
        State "double click again"
        Unload Me
        Exit Sub
    End Select
    If Err.Number <> 0 Then WriteLog "  step " & mStep & " ! " & Err.Number & " " & Err.Description
    Timer1.Enabled = True
End Sub
