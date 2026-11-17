VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "RTL"
   ClientHeight    =   4200
   ClientLeft      =   1500
   ClientTop       =   1500
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   4800
   StartUpPosition =   0  'Manual
   Begin VB.TextBox Text1 
      Text            =   "abc"
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2000
   End
   Begin VB.Label Label1 
      Caption         =   "Label"
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   540
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Button"
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   2000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check"
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1380
      Width           =   2000
   End
   Begin VB.ListBox List1 
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Text            =   "Combo"
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Text            =   "plain"
      Top             =   2760
      Width           =   2000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Phase 6.7: RightToLeft on a Windows that is not BiDi (Greek): what the property says and what changing it does.

Private Sub Report(ByVal what As String, c As Object)
    On Error Resume Next
    Dim v As Variant, a As Variant
    v = c.RightToLeft
    If Err.Number Then v = "err " & Err.Number
    Err.Clear
    a = c.Alignment
    If Err.Number Then a = "-"
    Err.Clear
    WriteLog what & ": RightToLeft " & v & ", Alignment " & a
End Sub

Private Sub SetIt(ByVal what As String, c As Object, ByVal value As Boolean)
    On Error Resume Next
    c.RightToLeft = value
    WriteLog what & " = " & value & ": err " & Err.Number & ", now " & c.RightToLeft
    Err.Clear
End Sub

Private Sub Form_Load()
    StartLog
    List1.AddItem "one"
    Report "Form", Me
    Report "Text1", Text1
    Report "Label1", Label1
    Report "Command1", Command1
    Report "Check1", Check1
    Report "List1", List1
    Report "Combo1", Combo1
    Report "Text2", Text2
    SetIt "Text1", Text1, False
    SetIt "Text2", Text2, True
    SetIt "Label1", Label1, False
    SetIt "Form", Me, False
    SetIt "Check1", Check1, False
    Dim t As TextBox
    Set t = Controls.Add("VB.TextBox", "Text3")
    Report "Text3 (added)", t
    WriteLog "done"
    Unload Me
End Sub
