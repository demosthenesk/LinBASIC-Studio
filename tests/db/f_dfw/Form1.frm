VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Writers"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4080
      Top             =   120
   End
   Begin VB.TextBox txtFields
      DataField       =   "WID"
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtFields
      DataField       =   "Name"
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtFields
      DataField       =   "Born"
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblStatus
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' text boxes of a control array bound to a data source class through its data member

Private WithEvents Src As CWriters
Attribute Src.VB_VarHelpID = -1
Private mStep As Integer

Private Function Cur() As String
    Dim t As TextBox, s As String
    For Each t In txtFields
        s = s & "|" & t.Text
    Next
    Cur = "[" & Mid$(s, 2) & "] status=[" & lblStatus.Caption & "]"
End Function

Private Sub Form_Load()
    Dim t As TextBox, n As Integer
    WriteLog "form1 load " & Cur()
    Set Src = New CWriters
    WriteLog "class made " & Cur()
    For Each t In Me.txtFields
        n = n + 1
        t.DataMember = "Primary"
        WriteLog "  member set " & n & " [" & t.DataMember & "]"
        Set t.DataSource = Src
        WriteLog "  bound " & n & " " & t.DataField & " " & TypeName(t.DataSource) & " " & Cur()
    Next
    WriteLog "form1 load end " & Cur()
    Timer1.Enabled = True
End Sub

Private Sub Src_MoveComplete()
    lblStatus.Caption = "Record: " & CStr(Src.AbsolutePosition)
    WriteLog "  form1: movecomplete " & lblStatus.Caption
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "form1 step " & mStep & " " & Cur()
    Select Case mStep
    Case 1
        Src.MoveNext
        WriteLog "after movenext " & Cur()
        Src.MoveLast
        WriteLog "after movelast " & Cur()
        Timer1.Enabled = True
    Case 2
        Src.MoveFirst
        WriteLog "after movefirst " & Cur()
        Form2.Show
        Unload Me
    End Select
End Sub
