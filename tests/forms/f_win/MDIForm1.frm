VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Tests"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private f1 As frmChild
Private f2 As frmChild
Private mStep As Integer
Private mResized As Boolean

Private Function ActCap() As String
    If Me.ActiveForm Is Nothing Then ActCap = "(none)" Else ActCap = Me.ActiveForm.Caption
End Function

Private Sub MDIForm_Load()
    StartLog
    WriteLog "mdi load " & TypeName(Me) & " " & Forms.Count & " " & Me.AutoShowChildren
    Set f1 = New frmChild
    f1.Tag = "one"
    f1.Caption = "Doc 1"
    WriteLog "f1 " & f1.Tag & " " & f1.Visible & " " & f1.MDIChild
    f1.Show
    Set f2 = New frmChild
    f2.Caption = "Doc 2"
    f2.Show
    WriteLog "children " & Forms.Count & " " & (f1 Is f2) & " " & TypeName(f1) & " " & f1.Caption & " " & f2.Caption & " " & ActCap()
    Timer1.Enabled = True
End Sub

Private Sub MDIForm_Activate()
    WriteLog "mdi activate " & ActCap()
End Sub

Private Sub MDIForm_Deactivate()
    WriteLog "mdi deactivate"
End Sub

Private Sub MDIForm_Resize()
    If Not mResized Then WriteLog "mdi resize"
    mResized = True
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WriteLog "mdi queryunload " & UnloadMode
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    WriteLog "mdi unload " & Forms.Count
    Set f2 = Nothing
    WriteLog "mdi unload end"
End Sub

Private Sub Timer1_Timer()
    Dim s As String, t As Object
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        WriteLog "active " & ActCap() & " " & (Screen.ActiveForm Is Me)
        f1.SetFocus
        WriteLog "after setfocus " & ActCap()
        WriteLog "count " & Forms.Count
        s = frmChild.Caption
        WriteLog "default " & s & " " & Forms.Count & " " & frmChild.Visible & " " & (frmChild Is f1)
    Case 2
        Unload f1
        WriteLog "after unload " & Forms.Count
        WriteLog "count " & Forms.Count
        Set f1 = Nothing
        WriteLog "released " & Forms.Count
    Case 3
        frmTool.Show
        WriteLog "tool shown"
    Case 4
        frmTool.RunTests
        Unload frmTool
        WriteLog "tool unloaded " & Forms.Count
    Case 5
        Load frmFixed
        LogBorders frmFixed, "fixed"
        Unload frmFixed
        Load frmNone
        LogBorders frmNone, "none"
        Unload frmNone
        Load frmToolWin
        LogBorders frmToolWin, "toolwin"
        Unload frmToolWin
        frmSize.Show
    Case 6
        frmSize.RunTests
        Unload frmSize
    Case 7
        WriteLog "screen " & Screen.TwipsPerPixelX & " " & Screen.TwipsPerPixelY & " " & (Screen.Width > 0) & " " & (Screen.Height > 0) & " " & (Screen.FontCount > 0) & " " & Screen.MousePointer & " " & Screen.ActiveForm.Name
        Screen.MousePointer = vbHourglass
        WriteLog "pointer " & Screen.MousePointer
        Screen.MousePointer = vbDefault
        WriteLog "app " & App.Title & "|" & App.EXEName & "|" & App.Major & "." & App.Minor & "." & App.Revision & "|" & App.CompanyName & "|" & App.ProductName & "|" & App.FileDescription & "|" & App.Comments & "|" & App.PrevInstance & "|" & App.TaskVisible & "|" & App.StartMode
        Clipboard.Clear
        Clipboard.SetText "hello clip"
        WriteLog "clip " & Clipboard.GetText & "|" & Clipboard.GetFormat(vbCFText) & "|" & Clipboard.GetFormat(vbCFBitmap)
        Clipboard.SetData LoadPicture(App.Path & "\test.bmp")
        Set t = Clipboard.GetData(vbCFBitmap)
        WriteLog "clip2 " & Clipboard.GetFormat(vbCFText) & "|" & Clipboard.GetFormat(vbCFBitmap) & "|" & t.Width & "|" & t.Type
        Clipboard.Clear
        WriteLog "clip3 [" & Clipboard.GetText & "]|" & Clipboard.GetFormat(vbCFText)
    Case 8
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub
