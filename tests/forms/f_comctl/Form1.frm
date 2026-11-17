VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Common controls"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5400
      Top             =   3360
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   741
      ButtonWidth     =   1032
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "open"
            Object.ToolTipText     =   "Open a file"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bold"
            Key             =   "chk"
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   50
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   4
      Max             =   20
      SelStart        =   5
      TickFrequency   =   2
      Value           =   5
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   1455
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2566
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "First"
            Key             =   "one"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Second"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4125
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4419
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   "msg"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "two"
            TextSave        =   "two"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mStep As Integer

Private Sub Form_Load()
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_comctl.txt"
    On Error GoTo 0
    WriteLog "progress " & ProgressBar1.Min & " " & ProgressBar1.Max & " " & ProgressBar1.Value & " " & TypeName(ProgressBar1)
    WriteLog "slider " & Slider1.Min & " " & Slider1.Max & " " & Slider1.Value & " " & Slider1.LargeChange & " " & Slider1.SmallChange & " " & Slider1.TickFrequency
    WriteLog "panels " & StatusBar1.Panels.Count & " " & StatusBar1.Panels(1).Text & " " & StatusBar1.Panels("msg").Index & " " & StatusBar1.Panels(2).Alignment & " " & StatusBar1.Panels(1).AutoSize
    WriteLog "buttons " & Toolbar1.Buttons.Count & " " & Toolbar1.Buttons(1).Caption & " " & Toolbar1.Buttons("chk").Style & " " & Toolbar1.Buttons(1).ToolTipText
    WriteLog "tabs " & TabStrip1.Tabs.Count & " " & TabStrip1.Tabs(1).Caption & " " & TabStrip1.Tabs("one").Index & " [" & TabStrip1.Tabs(2).Key & "]"
    WriteLog "images " & ImageList1.ListImages.Count & " " & ImageList1.ImageWidth
    Timer1.Enabled = True
End Sub

Private Sub Slider1_Change()
    WriteLog "slider change " & Slider1.Value
End Sub

Private Sub TabStrip1_Click()
    WriteLog "tab click " & TabStrip1.SelectedItem.Index
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    WriteLog "button click " & Button.Key
End Sub

Private Sub Report(ByVal what As String)
    WriteLog what & " err " & Err.Number
    Err.Clear
End Sub

Private Sub Timer1_Timer()
    Dim p As MSComctlLib.Panel, b As MSComctlLib.Button, t As MSComctlLib.Tab, li As MSComctlLib.ListImage
    Dim s As String, v As Variant
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        On Error Resume Next
        ProgressBar1.Value = 25
        WriteLog "progress value " & ProgressBar1.Value
        ProgressBar1.Value = 100
        Report "progress 100"
        ProgressBar1.Max = 200
        ProgressBar1.Value = 100
        WriteLog "progress max " & ProgressBar1.Max & " " & ProgressBar1.Value

        Slider1.Value = 7
        Slider1.Value = 30
        Report "slider 30"
        WriteLog "slider value " & Slider1.Value & " ticks " & Slider1.GetNumTicks

        Set p = StatusBar1.Panels.Add(, "k3", "three")
        WriteLog "panel add " & p.Index & " " & p.Key & " " & p.Text & " " & StatusBar1.Panels.Count
        StatusBar1.Panels.Add , "k3", "dup"
        Report "panel dup"
        StatusBar1.Panels.Add 9
        Report "panel index 9"
        StatusBar1.Panels.Add , "5", "num"
        Report "panel numeric key"
        StatusBar1.Panels(9).Text = "x"
        Report "panel 9"
        StatusBar1.Panels("none").Text = "x"
        Report "panel none"
        s = ""
        For Each v In StatusBar1.Panels
            s = s & v.Text & ","
        Next
        WriteLog "panels each " & s
        Set p = StatusBar1.Panels(1)
        StatusBar1.Panels.Remove 1
        WriteLog "panel removed " & StatusBar1.Panels.Count & " " & StatusBar1.Panels(1).Text
        s = p.Text
        Report "removed panel text"
        StatusBar1.Panels(1).Key = "k3"
        Report "panel key dup"
        StatusBar1.Style = sbrSimple
        StatusBar1.SimpleText = "simple"
        WriteLog "simple " & StatusBar1.Style & " " & StatusBar1.SimpleText

        Set b = Toolbar1.Buttons.Add(, "new", "New", tbrCheck)
        WriteLog "button add " & b.Index & " " & b.Caption & " " & b.Style & " " & Toolbar1.Buttons.Count
        WriteLog "button value " & Toolbar1.Buttons("chk").Value
        Toolbar1.Buttons("chk").Value = tbrPressed
        WriteLog "button value " & Toolbar1.Buttons("chk").Value
        Toolbar1.Buttons.Add , , , tbrSeparator
        WriteLog "buttons " & Toolbar1.Buttons.Count & " " & Toolbar1.Buttons(4).Style
        Toolbar1.Buttons.Remove "new"
        WriteLog "buttons after remove " & Toolbar1.Buttons.Count & " " & Toolbar1.Buttons(3).Style
        Toolbar1.Buttons(1).Image = 1
        Report "button image without list"

        Set t = TabStrip1.Tabs.Add(, "t3", "Third")
        WriteLog "tab add " & t.Index & " " & t.Key & " " & TabStrip1.Tabs.Count
        WriteLog "selected " & TabStrip1.SelectedItem.Index & " " & TabStrip1.Tabs(1).Selected
        Set TabStrip1.SelectedItem = TabStrip1.Tabs(2)
        WriteLog "selected " & TabStrip1.SelectedItem.Index & " " & TabStrip1.Tabs(2).Selected
        TabStrip1.Tabs(3).Selected = True
        WriteLog "selected " & TabStrip1.SelectedItem.Caption
        TabStrip1.Tabs.Remove 1
        WriteLog "tabs " & TabStrip1.Tabs.Count & " " & TabStrip1.Tabs(1).Caption

        Set li = ImageList1.ListImages.Add(1, "a", LoadPicture(App.Path & "\test.bmp"))
        WriteLog "image add " & li.Index & " " & li.Key & " " & ImageList1.ListImages.Count & " " & ImageList1.ImageWidth & " " & ImageList1.ImageHeight
        ImageList1.ListImages.Add , "a", LoadPicture(App.Path & "\test.bmp")
        Report "image dup"
        ImageList1.ListImages.Add , "b"
        Report "image no picture"
        ImageList1.ImageWidth = 32
        Report "imagewidth with images"
        Set Toolbar1.ImageList = ImageList1
        Toolbar1.Buttons(1).Image = "a"
        WriteLog "button image " & Toolbar1.Buttons(1).Image
        Toolbar1.Buttons(1).Image = "zz"
        Report "button image zz"
        ImageList1.ListImages.Remove "a"
        WriteLog "images " & ImageList1.ListImages.Count
    Case 2
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub
