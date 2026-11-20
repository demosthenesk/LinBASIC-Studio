VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "AW5 Toolbar"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            ImageKey        =   "Left"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            ImageKey        =   "Right"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            ImageKey        =   "Bold"
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   4740
      TabIndex        =   1
      Top             =   1200
      Width           =   4800
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   4740
      TabIndex        =   2
      Top             =   600
      Width           =   4800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3480
      Top             =   2400
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0352
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06A4
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09F6
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D48
            Key             =   "Bold"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 9.2a-5: what the Application Wizard needs of the runtime - the ImageList after the Toolbar in the .frm (ImageKey),
' ButtonWidth / ButtonHeight (HIMETRIC in the .frm), Value of the buttons of a group, the order of the controls with
' Align, the Caption of a menu separator; in ModLog: Shell of a program that does not exist, New MDIForm
Private mStep As Integer

Private Function Btn(ByVal i As Integer) As String
    Dim b As Object
    Set b = Toolbar1.Buttons(i)
    Btn = "[" & b.Key & "] style " & b.Style & " value " & b.Value & " image " & TypeName(b.Image) & " [" & b.Image & "]"
End Function

Private Function Grp() As String
    Grp = Toolbar1.Buttons("Left").Value & Toolbar1.Buttons("Center").Value & Toolbar1.Buttons("Right").Value
End Function

Private Sub Form_Load()
    WriteLog "form1 load " & Forms.Count
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        WriteLog "imagelist " & TypeName(Toolbar1.ImageList) & " " & Toolbar1.ImageList.ListImages.Count & " " & Toolbar1.ImageList.ListImages(1).Key & " " & Toolbar1.ImageList.ListImages("Bold").Index
        WriteLog "buttons " & Toolbar1.Buttons.Count
        For i = 1 To Toolbar1.Buttons.Count
            WriteLog "  b" & i & " " & Btn(i)
        Next
        WriteLog "button size " & CLng(Toolbar1.ButtonWidth) & " " & CLng(Toolbar1.ButtonHeight)
        WriteLog "align " & Toolbar1.Top & " " & (Picture2.Top = Toolbar1.Top + Toolbar1.Height) & " " & (Picture1.Top = Picture2.Top + Picture2.Height) & " " & Picture1.Align & " " & Picture2.Align & " " & (Toolbar1.Width = Me.ScaleWidth)
        WriteLog "menu [" & mnuSep.Caption & "] [" & mnuExit.Caption & "] " & mnuSep.Enabled
    Case 2
        WriteLog "group " & Grp()
        Toolbar1.Buttons("Center").Value = tbrPressed
        WriteLog "group center " & Grp()
        Toolbar1.Buttons("Left").Value = tbrPressed
        WriteLog "group left " & Grp()
        Toolbar1.Buttons("Left").Value = tbrUnpressed
        WriteLog "group none " & Grp()
        Toolbar1.Buttons("Bold").Value = tbrPressed
        WriteLog "check " & Toolbar1.Buttons("Bold").Value
        Toolbar1.Buttons("Bold").Value = tbrUnpressed
        WriteLog "check off " & Toolbar1.Buttons("Bold").Value
        Toolbar1.Buttons("New").Image = "Bold"
        WriteLog "image key " & Btn(1)
        Toolbar1.Buttons("New").Image = 2
        WriteLog "image index " & Btn(1)
    Case 3
        WriteLog "end " & Forms.Count & " " & FormNames()
        Unload frmChild
        Unload fMain
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub
