VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "User controls"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4200
      Top             =   3000
   End
   Begin UCTests.Counter Counter2 
      Height          =   600
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1058
   End
   Begin UCTests.Counter Counter1 
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Tag             =   "first"
      Top             =   240
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1058
      Value           =   5
      Step            =   2
      Caption         =   "A"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' user controls (LinBASIC Studio, phase 4.3): the events of a UserControl, its properties from the .frm
' (PropertyBag), Extender and Ambient, RaiseEvent to the form, a control array of user controls

Private mStep As Integer

Private Sub Form_Initialize()
    WriteLog "form initialize"
End Sub

Private Sub Form_Load()
    Dim c As Control, s As String
    WriteLog "form load"
    WriteLog "typename " & TypeName(Counter1) & " " & (TypeOf Counter1 Is Counter) & " " & TypeName(Counter2(0))
    WriteLog "counter1 " & Counter1.Value & " default " & Counter1 & " step " & Counter1.Step & " caption " & Counter1.Caption
    WriteLog "counter2 " & Counter2(0).Value & " step " & Counter2(0).Step & " caption " & Counter2(0).Caption
    WriteLog "extender " & Counter1.Name & " tag " & Counter1.Tag & " " & Counter1.Left & "," & Counter1.Top & " " & _
        Counter1.Width & "x" & Counter1.Height & " tabindex " & Counter1.TabIndex & " index " & Counter2(0).Index
    For Each c In Controls
        s = s & c.Name & ":" & TypeName(c) & " "
    Next
    WriteLog "controls " & Controls.Count & " " & s
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    WriteLog "form activate"
End Sub

Private Sub Form_Resize()
    WriteLog "form resize"
End Sub

Private Sub Counter1_Changed(ByVal NewValue As Long, Cancel As Boolean)
    WriteLog "counter1 changed " & NewValue
    If NewValue > 100 Then Cancel = True
End Sub

Private Sub Counter2_Changed(Index As Integer, ByVal NewValue As Long, Cancel As Boolean)
    WriteLog "counter2(" & Index & ") changed " & NewValue
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        Counter1.Value = 7
        WriteLog "value " & Counter1.Value
        Counter1 = 9
        WriteLog "default " & Counter1.Value
        Counter1.Value = 500
        WriteLog "after veto " & Counter1.Value
        Counter1.PressUp
        WriteLog "pressup " & Counter1.Value
        Counter1.Step = 10
        Counter1.PressUp
        WriteLog "pressup " & Counter1.Value
    Case 2
        WriteLog "info " & Counter1.Info
        Counter1.Tag = "t1"
        Counter1.Move 120, 120, 2400, 900
        WriteLog "info " & Counter1.Info
        Counter1.Width = 2000
        Counter1.Left = 300
        WriteLog "moved " & Counter1.Left & " " & Counter1.Width
    Case 3
        Counter1.Visible = False
        WriteLog "visible " & Counter1.Visible
        Counter1.Visible = True
        WriteLog "visible " & Counter1.Visible
        Counter1.Enabled = False
        WriteLog "enabled " & Counter1.Enabled
        Counter1.Enabled = True
    Case 4
        Load Counter2(1)
        WriteLog "loaded " & Counter2(1).Value & " " & Counter2(1).Caption & " " & Counter2(1).Visible & " " & Counter2(1).Top
        Counter2(1).Top = 1900
        Counter2(1).Visible = True
        Counter2(1).Value = 3
        Counter2(0).Value = 4
        WriteLog "values " & Counter2(0).Value & " " & Counter2(1).Value & " count " & Controls.Count
        Unload Counter2(1)
        WriteLog "unloaded count " & Controls.Count
    Case 5
        Dim o As Object
        Set o = Counter1
        WriteLog "object " & TypeName(o) & " " & o.Value & " " & o.Name
        On Error Resume Next
        o.NoSuchThing = 1
        WriteLog "nosuch " & Err.Number
        Err.Clear
        Counter1.Value = "abc"
        WriteLog "mismatch " & Err.Number
        Err.Clear
        On Error GoTo 0
    Case 6
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "form unload"
End Sub

Private Sub Form_Terminate()
    WriteLog "form terminate"
End Sub
