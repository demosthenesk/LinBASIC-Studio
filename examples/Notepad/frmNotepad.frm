VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNotepad 
   Caption         =   "Untitled - Notepad"
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   6120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditTime 
         Caption         =   "Time/&Date"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFormatFont 
         Caption         =   "&Font..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Notepad"
      End
   End
End
Attribute VB_Name = "frmNotepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Notepad - menus, the CommonDialog control, a multi-line TextBox, the Clipboard and text files
Option Explicit

Private mFile As String         ' the file being edited ("" for a new one)
Private mChanged As Boolean     ' changed since it was opened or saved

Private Sub Form_Load()
    txtText.Text = "Welcome to LinBASIC Studio!" & vbCrLf & vbCrLf & _
                   "This Notepad is a VB6 project: a form with menus, a TextBox and a CommonDialog." & vbCrLf & _
                   "Open, edit and save text files, or press F5 for the time and date." & vbCrLf
    mChanged = False
    ShowTitle
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub
    txtText.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not OkToDiscard() Then Cancel = True
End Sub

Private Sub txtText_Change()
    If Not mChanged Then
        mChanged = True
        ShowTitle
    End If
End Sub

Private Sub ShowTitle()
    Dim name As String
    If mFile = "" Then name = "Untitled" Else name = Mid$(mFile, InStrRev(Replace(mFile, "\", "/"), "/") + 1)
    Caption = IIf(mChanged, "*", "") & name & " - Notepad"
End Sub

Private Function OkToDiscard() As Boolean
    OkToDiscard = True
    If Not mChanged Then Exit Function
    Select Case MsgBox("Save the changes?", vbYesNoCancel + vbQuestion, "Notepad")
    Case vbYes: OkToDiscard = SaveFile(mFile)
    Case vbCancel: OkToDiscard = False
    End Select
End Function

Private Function SaveFile(ByVal path As String) As Boolean
    Dim f As Integer
    On Error GoTo Failed
    If path = "" Then
        dlgFile.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        dlgFile.FileName = ""
        dlgFile.CancelError = True
        dlgFile.ShowSave
        path = dlgFile.FileName
    End If
    f = FreeFile
    Open path For Output As #f
    Print #f, txtText.Text;
    Close #f
    mFile = path
    mChanged = False
    ShowTitle
    SaveFile = True
    Exit Function
Failed:
    If Err.Number <> cdlCancel Then MsgBox Err.Description, vbCritical, "Notepad"
End Function

Private Sub mnuFileNew_Click()
    If Not OkToDiscard() Then Exit Sub
    txtText.Text = ""
    mFile = ""
    mChanged = False
    ShowTitle
End Sub

Private Sub mnuFileOpen_Click()
    Dim f As Integer
    If Not OkToDiscard() Then Exit Sub
    On Error GoTo Failed
    dlgFile.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    dlgFile.CancelError = True
    dlgFile.ShowOpen
    f = FreeFile
    Open dlgFile.FileName For Input As #f
    txtText.Text = Input$(LOF(f), #f)
    Close #f
    mFile = dlgFile.FileName
    mChanged = False
    ShowTitle
    Exit Sub
Failed:
    If Err.Number <> cdlCancel Then MsgBox Err.Description, vbCritical, "Notepad"
End Sub

Private Sub mnuFileSave_Click()
    SaveFile mFile
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveFile ""
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuEditCut_Click()
    Clipboard.Clear
    Clipboard.SetText txtText.SelText
    txtText.SelText = ""
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtText.SelText
End Sub

Private Sub mnuEditPaste_Click()
    txtText.SelText = Clipboard.GetText
End Sub

Private Sub mnuEditSelectAll_Click()
    txtText.SelStart = 0
    txtText.SelLength = Len(txtText.Text)
End Sub

Private Sub mnuEditTime_Click()
    txtText.SelText = Format$(Now, "hh:nn dd/mm/yyyy")
End Sub

Private Sub mnuFormatFont_Click()
    On Error GoTo Cancelled
    dlgFile.Flags = cdlCFScreenFonts
    dlgFile.FontName = txtText.FontName
    dlgFile.FontSize = txtText.FontSize
    dlgFile.CancelError = True
    dlgFile.ShowFont
    txtText.FontName = dlgFile.FontName
    txtText.FontSize = dlgFile.FontSize
    txtText.FontBold = dlgFile.FontBold
    txtText.FontItalic = dlgFile.FontItalic
Cancelled:
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "Notepad " & App.Major & "." & App.Minor & vbCrLf & "An example of LinBASIC Studio.", vbInformation, "About Notepad"
End Sub
