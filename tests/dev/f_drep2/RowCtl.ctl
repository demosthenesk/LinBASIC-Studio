VERSION 5.00
Begin VB.UserControl RowCtl
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   ScaleHeight     =   600
   ScaleWidth      =   4200
   Begin VB.CheckBox chkActive
      Caption         =   "Active"
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox txtAge
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   60
      Width           =   735
   End
   Begin VB.TextBox txtName
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1755
   End
End
Attribute VB_Name = "RowCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' DataRepeater (LinBASIC Studio phase 7.5): the control the DataRepeater repeats, one row of People. PName, PAge and
' PActive are bindable (Procedure Attributes: Property is data bound, Show in DataBindings collection); Info is not.
' It writes when it is made and ended into the log of the test that runs (expected\f_drep*.txt).

Private mInfo As String

Private Sub L(ByVal s As String)
    Dim f As Integer, n As String
    On Error Resume Next
    n = Dir(App.Path & "\expected\f_drep*.txt")
    If n = "" Then Exit Sub
    f = FreeFile
    Open App.Path & "\expected\" & n For Append As #f
    Print #f, s
    Close #f
End Sub

Public Property Get PName() As String
Attribute PName.VB_MemberFlags = "101c"
    PName = txtName.Text
End Property

Public Property Let PName(ByVal v As String)
    txtName.Text = v
    PropertyChanged "PName"
End Property

Public Property Get PAge() As String
Attribute PAge.VB_MemberFlags = "101c"
    PAge = txtAge.Text
End Property

Public Property Let PAge(ByVal v As String)
    txtAge.Text = v
    PropertyChanged "PAge"
End Property

Public Property Get PActive() As Boolean
Attribute PActive.VB_MemberFlags = "101c"
    PActive = (chkActive.Value = vbChecked)
End Property

Public Property Let PActive(ByVal v As Boolean)
    chkActive.Value = IIf(v, vbChecked, vbUnchecked)
    PropertyChanged "PActive"
End Property

Public Property Get Info() As String
    Info = mInfo
End Property

Public Property Let Info(ByVal v As String)
    mInfo = v
End Property

Private Sub txtName_Change()
    PropertyChanged "PName"
End Sub

Private Sub txtAge_Change()
    PropertyChanged "PAge"
End Sub

Private Sub chkActive_Click()
    PropertyChanged "PActive"
End Sub

Private Sub UserControl_Initialize()
    L "  ctl init"
End Sub

Private Sub UserControl_Terminate()
    L "  ctl term"
End Sub
