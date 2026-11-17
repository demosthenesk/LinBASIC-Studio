VERSION 5.00
Begin VB.UserControl Counter 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   ScaleHeight     =   600
   ScaleWidth      =   1800
   Begin VB.CommandButton cmdUp 
      Caption         =   "+"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblVal 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' a user control: a value with a + button; Value is its default property

Public Event Changed(ByVal NewValue As Long, Cancel As Boolean)

Const DEF_CAPTION = "Cnt"

Private mValue As Long
Private mStep As Long
Private mCaption As String
Private mId As Long

Private Sub L(ByVal s As String)
    WriteLog "  uc" & mId & " " & s
End Sub

Private Sub UserControl_Initialize()
    gSerial = gSerial + 1
    mId = gSerial
    L "initialize " & ScaleWidth & "x" & ScaleHeight
End Sub

Private Sub UserControl_InitProperties()
    mValue = 0
    mStep = 1
    mCaption = DEF_CAPTION
    lblVal.Caption = mValue
    L "initproperties " & Extender.Name & " " & Ambient.DisplayName
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mValue = PropBag.ReadProperty("Value", 0)
    mStep = PropBag.ReadProperty("Step", 1)
    mCaption = PropBag.ReadProperty("Caption", DEF_CAPTION)
    lblVal.Caption = mValue
    L "readproperties " & Extender.Name & " value=" & mValue & " step=" & mStep & " caption=" & mCaption & _
      " usermode=" & Ambient.UserMode & " display=" & Ambient.DisplayName
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    L "writeproperties"
    PropBag.WriteProperty "Value", mValue, 0
    PropBag.WriteProperty "Step", mStep, 1
    PropBag.WriteProperty "Caption", mCaption, DEF_CAPTION
End Sub

Private Sub UserControl_Resize()
    L "resize " & ScaleWidth & "x" & ScaleHeight
End Sub

Private Sub UserControl_Show()
    L "show"
End Sub

Private Sub UserControl_Hide()
    L "hide"
End Sub

Private Sub UserControl_Terminate()
    L "terminate"
End Sub

Private Sub cmdUp_Click()
    L "cmdup click"
    Value = mValue + mStep
End Sub

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
    Value = mValue
End Property

Public Property Let Value(ByVal v As Long)
    Dim Cancel As Boolean
    RaiseEvent Changed(v, Cancel)
    If Cancel Then
        L "vetoed " & v
        Exit Property
    End If
    mValue = v
    lblVal.Caption = v
    PropertyChanged "Value"
End Property

Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal s As String)
    mCaption = s
    PropertyChanged "Caption"
End Property

Public Property Get Step() As Long
    Step = mStep
End Property

Public Property Let Step(ByVal n As Long)
    mStep = n
    PropertyChanged "Step"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal b As Boolean)
    UserControl.Enabled = b
    PropertyChanged "Enabled"
End Property

Public Sub PressUp()
    cmdUp.Value = True
End Sub

Public Function Info() As String
    Info = "back=" & Hex(Ambient.BackColor) & " fore=" & Hex(Ambient.ForeColor) & " font=" & Ambient.Font.Name & _
        " usermode=" & Ambient.UserMode & " ext=" & Extender.Name & " " & Extender.Left & "," & Extender.Top & " " & _
        Extender.Width & "x" & Extender.Height & " tag=[" & Extender.Tag & "] parent=" & Extender.Parent.Name & _
        " visible=" & Extender.Visible & " scale=" & ScaleWidth & "x" & ScaleHeight & _
        " lbl=" & lblVal.Caption & " ctls=" & Controls.Count
End Function
