VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
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
      Interval        =   100
      Left            =   4200
      Top             =   3000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   2640
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCancelOnce As Boolean

Private Sub Form_Initialize()
    WriteLog "Form1 Initialize"
End Sub

Private Sub Form_Load()
    WriteLog "Form1 Load " & Me.Name & " " & Me.Caption & " " & Forms.Count
    WriteLog "controls " & Me.Controls.Count & " " & TypeName(Command1) & " " & TypeName(Text1) & " " & TypeName(Option1) & " " & TypeName(Option1(0)) & " " & TypeName(Timer1)
    WriteLog "command1 " & Command1.Caption & " " & Command1.Left & " " & Command1.Top & " " & Command1.Width & " " & Command1.Height & " " & Command1.Enabled & " " & Command1.Visible & " " & Command1.TabIndex
    WriteLog "text1 [" & Text1.Text & "] " & Text1 & " " & Len(Text1.Text) & " " & Text1.SelStart
    WriteLog "label1 " & Label1.Caption & " " & Label1
    WriteLog "option1 " & Option1.Count & " " & Option1.LBound & " " & Option1.UBound & " " & Option1(0).Value & " " & Option1(1).Index
    WriteLog "timer1 " & Timer1.Interval & " " & Timer1.Enabled
    Text1.Text = "hello"
    WriteLog "after text " & Text1.Text
    Text1 = "world"
    Check1.Value = vbChecked
    Option1(1).Value = True
    WriteLog "options " & Option1(0).Value & " " & Option1(1).Value
    List1.AddItem "a"
    List1.AddItem "b"
    List1.AddItem "c", 0
    WriteLog "list " & List1.ListCount & " " & List1.List(0) & " " & List1.ListIndex
    List1.ListIndex = 1
    WriteLog "list text " & List1.Text
    Combo1.AddItem "x"
    Combo1.AddItem "y"
    Combo1.ListIndex = 1
    WriteLog "combo " & Combo1.Text & " " & Combo1.ListCount
    Command1.Value = True
    Load Option1(2)
    WriteLog "loaded option " & Option1.Count & " " & Option1(2).Visible & " " & Option1(2).Top
    Command1.Caption = "Go"
    Me.Caption = "Main"
    WriteLog "end load " & Command1.Caption & " " & Me.Caption
    Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
    WriteLog "Form1 Resize"
End Sub

Private Sub Form_Activate()
    WriteLog "Form1 Activate"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WriteLog "Form1 QueryUnload " & UnloadMode
    If mCancelOnce Then
        Cancel = True
        mCancelOnce = False
        WriteLog "cancelled"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "Form1 Unload"
End Sub

Private Sub Form_Terminate()
    WriteLog "Form1 Terminate"
End Sub

Private Sub Command1_Click()
    WriteLog "Command1 Click"
End Sub

Private Sub Text1_Change()
    WriteLog "Text1 Change " & Text1.Text
End Sub

Private Sub Check1_Click()
    WriteLog "Check1 Click " & Check1.Value
End Sub

Private Sub Option1_Click(Index As Integer)
    WriteLog "Option1 Click " & Index
End Sub

Private Sub List1_Click()
    WriteLog "List1 Click " & List1.ListIndex
End Sub

Private Sub Combo1_Click()
    WriteLog "Combo1 Click " & Combo1.ListIndex
End Sub

Private Sub Timer1_Timer()
    gStep = gStep + 1
    WriteLog "Timer " & gStep
    Select Case gStep
    Case 1
        Load Form2
        WriteLog "forms " & Forms.Count & " " & Form2.Visible
        Form2.Show
        WriteLog "shown " & Form2.Visible
    Case 2
        Form2.Hide
        WriteLog "hidden " & Form2.Visible & " " & Forms.Count
        Unload Form2
        WriteLog "unloaded " & Forms.Count
        WriteLog "caption again " & Form2.Caption
        WriteLog "forms " & Forms.Count
        Unload Form2
        Set Form2 = Nothing
        WriteLog "form2 nothing"
    Case 3
        Unload Option1(2)
        WriteLog "option count " & Option1.Count
        mCancelOnce = True
        Unload Me
        WriteLog "still loaded " & Forms.Count
    Case 4
        Timer1.Enabled = False
        Unload Me
        WriteLog "after unload me " & Forms.Count
    End Select
End Sub
