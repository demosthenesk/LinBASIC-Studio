VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Help Form1"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5400
   HelpContextID   =   100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4200
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame 300"
      Height          =   975
      HelpContextID   =   300
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2400
      Begin VB.CheckBox Check1 
         Caption         =   "Check in frame"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text without ID"
      Top             =   600
      Width           =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Button 200"
      Height          =   375
      HelpContextID   =   200
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         HelpContextID   =   250
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' f_help (10.1): App.HelpFile, HelpContextID, WhatsThisHelp / WhatsThisButton / WhatsThisHelpID,
' Err.HelpFile / Err.HelpContext, the help properties of the CommonDialog. VB6 writes them to
' expected\f_help.txt. The help itself (F1, What's This, ShowHelp) only in LinBASIC's test run
' (VBRT_HELP_LOG: what the help opens, compared with expected\f_help-help.txt).

Private Sub Form_Load()
    StartLog
    WriteLog "load helpfile [" & App.HelpFile & "]"
    App.HelpFile = App.Path & "\h1.chm"
    WriteLog "set helpfile " & ShortName(App.HelpFile)
    WriteLog "ids " & Me.HelpContextID & " " & Command1.HelpContextID & " " & Text1.HelpContextID & " " & Frame1.HelpContextID & " " & Check1.HelpContextID
    WriteLog "whatsthis " & Me.WhatsThisHelp & " " & Me.WhatsThisButton & " " & Command1.WhatsThisHelpID
    WriteLog "menu " & mnuOpen.HelpContextID & " " & mnuFile.HelpContextID
    TestErrors
    TestDialog
    TestSet
    Load Form2
End Sub

Private Sub TestErrors()
    Dim v As Long, z As Long
    On Error Resume Next
    v = 1 \ z
    WriteLog "err11 " & Err.Number & " [" & Err.HelpFile & "] " & Err.HelpContext
    Err.Clear
    Err.Raise 1000, "FHelp", "Test error", App.Path & "\h1.chm", 400
    WriteLog "err1000 " & Err.Number & " " & Err.Source & " " & Err.Description & " [" & ShortName(Err.HelpFile) & "] " & Err.HelpContext
    Err.Clear
    Err.Raise 1002
    WriteLog "err1002 [" & Err.HelpFile & "] " & Err.HelpContext & " " & Err.Description
    Err.Clear
    Err.Raise 1003, , , App.Path & "\h1.chm"
    WriteLog "err1003 [" & ShortName(Err.HelpFile) & "] " & Err.HelpContext
    Err.Clear
    WriteLog "clear [" & Err.HelpFile & "] " & Err.HelpContext
    Err.HelpFile = "x.chm"
    Err.HelpContext = 7
    WriteLog "set [" & Err.HelpFile & "] " & Err.HelpContext
    Err.Clear
    WriteLog "clear2 [" & Err.HelpFile & "] " & Err.HelpContext
End Sub

Private Sub TestDialog()
    With CommonDialog1
        WriteLog "dlg0 " & .HelpCommand & " " & .HelpContext & " [" & .HelpFile & "] [" & .HelpKey & "]"
        .HelpCommand = cdlHelpKey
        .HelpContext = 300
        .HelpFile = "h1.chm"
        .HelpKey = "Frame"
        WriteLog "dlg1 " & .HelpCommand & " " & .HelpContext & " [" & .HelpFile & "] [" & .HelpKey & "]"
    End With
End Sub

Private Sub TestSet()
    Dim f As Object
    Set f = Me
    On Error Resume Next
    f.HelpContextID = 150
    WriteLog "form id " & Err.Number & " " & Me.HelpContextID
    Err.Clear
    f.WhatsThisHelp = True
    WriteLog "form wth " & Err.Number & " " & Me.WhatsThisHelp
    Err.Clear
    f.WhatsThisButton = True
    WriteLog "form wtb " & Err.Number & " " & Me.WhatsThisButton
    Err.Clear
    Command1.HelpContextID = 210
    Text1.WhatsThisHelpID = 610
    mnuOpen.HelpContextID = 260
    WriteLog "set ids " & Command1.HelpContextID & " " & Text1.WhatsThisHelpID & " " & mnuOpen.HelpContextID
    Command1.HelpContextID = 200
    Text1.WhatsThisHelpID = 0
    mnuOpen.HelpContextID = 250
    Me.HelpContextID = 100
    f.WhatsThisHelp = False
    Err.Clear
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If Environ$("VBRT_HELP_LOG") <> "" Then RunHelp
    WriteLog "end"
    Unload Form2
    Unload Me
End Sub

' LinBASIC only: what the help opens goes to VBRT_HELP_LOG
Private Sub RunHelp()
    Command1.SetFocus
    SendKeys "{F1}", True
    Text1.SetFocus
    SendKeys "{F1}", True
    Check1.SetFocus
    SendKeys "{F1}", True
    With CommonDialog1
        .HelpFile = App.Path & "\h1abs.chm"
        .HelpCommand = cdlHelpContext
        .HelpContext = 300
        .ShowHelp
        .HelpFile = App.HelpFile
        .HelpCommand = cdlHelpContents
        .ShowHelp
        .HelpCommand = cdlHelpKey
        .HelpKey = "Frame"
        .ShowHelp
        .HelpCommand = cdlHelpContextPopup
        .HelpContext = 700
        .ShowHelp
    End With
    Form2.Show
    Form2.Text1.SetFocus
    SendKeys "{F1}", True
    Form2.Command1.ShowWhatsThis
    Form2.Text1.WhatsThisHelpID = 0
    Form2.Text1.ShowWhatsThis
    Form2.WhatsThisMode
    Form2.Hide
    Me.HelpContextID = 0
    Text1.SetFocus
    SendKeys "{F1}", True
    App.HelpFile = App.Path & "\nofile.chm"
    Command1.SetFocus
    SendKeys "{F1}", True
End Sub
