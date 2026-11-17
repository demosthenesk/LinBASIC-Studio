VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Controls 2"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4800
      Top             =   3600
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   3240
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1695
      Left            =   4800
      Max             =   0
      Min             =   10
      TabIndex        =   1
      Top             =   240
      Value           =   5
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   20
      Left            =   240
      Max             =   100
      SmallChange     =   5
      TabIndex        =   0
      Top             =   1560
      Value           =   10
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   1320
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   5
      Top             =   240
      Width           =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   2640
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   2760
      Shape           =   2  'Oval
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mStep As Integer

Private Function LastPart(ByVal s As String) As String
    Dim i As Integer, j As Integer
    If Len(s) > 1 And (Right$(s, 1) = "\" Or Right$(s, 1) = "/") Then s = Left$(s, Len(s) - 1)
    i = InStrRev(s, "\")
    j = InStrRev(s, "/")
    If j > i Then i = j
    LastPart = Mid$(s, i + 1)
End Function

Private Sub MakeFile(ByVal name As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\ctl2data\" & name For Output As #f
    Print #f, name
    Close #f
End Sub

Private Sub Form_Load()
    Dim c As Control, s As String
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_ctl2.txt"
    Kill App.Path & "\ctl2data\A.txt"
    Kill App.Path & "\ctl2data\b.txt"
    Kill App.Path & "\ctl2data\c.dat"
    RmDir App.Path & "\ctl2data\beta"
    RmDir App.Path & "\ctl2data\Alpha"
    RmDir App.Path & "\ctl2data"
    On Error GoTo 0
    MkDir App.Path & "\ctl2data"
    MkDir App.Path & "\ctl2data\beta"
    MkDir App.Path & "\ctl2data\Alpha"
    MakeFile "b.txt"
    MakeFile "A.txt"
    MakeFile "c.dat"

    For Each c In Me.Controls
        s = s & " " & TypeName(c)
    Next
    WriteLog "controls" & s
    WriteLog "shape " & Shape1.Shape & " " & Hex$(Shape1.BorderColor) & " " & Shape1.FillStyle & " " & Hex$(Shape1.FillColor) & " " & Shape1.BackStyle & " " & Shape1.BorderWidth & " " & Shape1.BorderStyle & " " & Shape1.Visible
    WriteLog "line " & Line1.X1 & " " & Line1.Y1 & " " & Line1.X2 & " " & Line1.Y2 & " " & Line1.BorderWidth & " " & Hex$(Line1.BorderColor) & " " & Line1.BorderStyle
    Shape1.Shape = 3
    Line1.X2 = 2000
    WriteLog "changed " & Shape1.Shape & " " & Line1.X2
    WriteLog "image " & Image1.Width & " " & Image1.Height & " " & Image1.Stretch & " " & Image1.Picture.Type & " " & Image1.Picture.Width

    WriteLog "hscroll " & HScroll1.Value & " " & HScroll1.Min & " " & HScroll1.Max & " " & HScroll1.SmallChange & " " & HScroll1.LargeChange
    WriteLog "vscroll " & VScroll1.Value & " " & VScroll1.Min & " " & VScroll1.Max & " " & VScroll1.SmallChange & " " & VScroll1.LargeChange
    HScroll1.Value = 50
    HScroll1.Value = 50
    VScroll1.Value = 2
    On Error Resume Next
    HScroll1.Value = 200
    WriteLog "err " & Err.Number
    Err.Clear
    VScroll1.Value = 11
    WriteLog "err " & Err.Number
    Err.Clear
    On Error GoTo 0
    WriteLog "values " & HScroll1.Value & " " & VScroll1.Value

    Set Image1.Picture = LoadPicture(App.Path & "\test.bmp")
    WriteLog "loaded " & Image1.Width & " " & Image1.Height & " " & Image1.Picture.Type & " " & Image1.Picture.Width & " " & Image1.Picture.Height & " " & CLng(ScaleX(Image1.Picture.Width, vbHimetric, vbPixels))
    Set Picture1.Picture = LoadPicture(App.Path & "\test.bmp")
    WriteLog "picture " & Picture1.Width & " " & Picture1.Height & " " & Picture1.ScaleWidth & " " & Picture1.Point(2, 2) & " " & Picture1.Point(30, 2)
    Picture1.PSet (5, 5), vbGreen
    Picture1.Cls
    WriteLog "after cls " & Picture1.Point(2, 2) & " " & Picture1.Point(5, 5)
    On Error Resume Next
    Set Image1.Picture = LoadPicture(App.Path & "\nofile.bmp")
    WriteLog "err " & Err.Number & " " & Image1.Picture.Width
    Err.Clear
    On Error GoTo 0
    Set Image1.Picture = LoadPicture()
    WriteLog "empty " & Image1.Picture.Type & " " & Image1.Picture.Width & " " & Image1.Width

    WriteLog "drive " & (Drive1.ListCount > 0) & " " & (Len(Drive1.Drive) > 0)
    Dir1.Path = App.Path & "\ctl2data"
    WriteLog "dir " & Dir1.ListCount & " " & Dir1.ListIndex & " " & LastPart(Dir1.List(-1)) & " " & LastPart(Dir1.List(0)) & " " & LastPart(Dir1.List(1)) & " " & (Dir1.List(0) = Dir1.Path & "\" & LastPart(Dir1.List(0)) Or Dir1.List(0) = Dir1.Path & "/" & LastPart(Dir1.List(0)))
    File1.Path = Dir1.Path
    WriteLog "files " & File1.ListCount & " " & File1.List(0) & " " & File1.List(1) & " " & File1.List(2) & " [" & File1.FileName & "] " & File1.Pattern
    WriteLog "attrs " & File1.Normal & " " & File1.Archive & " " & File1.Hidden & " " & File1.ReadOnly & " " & File1.System
    File1.Pattern = "*.txt"
    WriteLog "txt " & File1.ListCount & " " & File1.List(0) & " " & File1.List(1)
    File1.ListIndex = 1
    WriteLog "selected " & File1.FileName & " " & LastPart(File1.Path)
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    WriteLog "activate " & Me.ActiveControl.Name
End Sub

Private Sub Form_Deactivate()
    WriteLog "deactivate"
End Sub

Private Sub HScroll1_Change()
    WriteLog "hscroll change " & HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    WriteLog "hscroll scroll " & HScroll1.Value
End Sub

Private Sub HScroll1_GotFocus()
    WriteLog "hscroll gotfocus"
End Sub

Private Sub HScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    WriteLog "hscroll keydown " & KeyCode
End Sub

Private Sub HScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
    WriteLog "hscroll keyup " & KeyCode
End Sub

Private Sub VScroll1_Change()
    WriteLog "vscroll change " & VScroll1.Value
End Sub

Private Sub VScroll1_GotFocus()
    WriteLog "vscroll gotfocus"
End Sub

Private Sub Dir1_Change()
    WriteLog "dir1 change " & LastPart(Dir1.Path)
End Sub

Private Sub File1_PathChange()
    WriteLog "file1 pathchange " & LastPart(File1.Path)
End Sub

Private Sub File1_PatternChange()
    WriteLog "file1 patternchange " & File1.Pattern
End Sub

Private Sub File1_Click()
    WriteLog "file1 click " & File1.ListIndex
End Sub

Private Sub Timer1_Timer()
    Dim r As Integer, s As String
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        SendKeys "{RIGHT}", True
        SendKeys "{PGDN}", True
        SendKeys "{END}", True
        SendKeys "{HOME}", True
        WriteLog "hscroll " & HScroll1.Value
    Case 2
        VScroll1.SetFocus
        SendKeys "{DOWN}", True
        SendKeys "{PGUP}", True
        WriteLog "vscroll " & VScroll1.Value
    Case 3
        SendKeys "{ENTER}"
        r = MsgBox("first", vbOKCancel, "T")
        WriteLog "msgbox " & r
        SendKeys "{TAB}{ENTER}"
        r = MsgBox("second", vbYesNo)
        WriteLog "msgbox " & r
        SendKeys "{ESC}"
        r = MsgBox("third", vbYesNoCancel + vbQuestion)
        WriteLog "msgbox " & r
        SendKeys "hello{ENTER}"
        s = InputBox("name?", "T", "def")
        WriteLog "inputbox [" & s & "]"
        SendKeys "{ESC}"
        s = InputBox("x", , "def")
        WriteLog "inputbox [" & s & "]"
        SendKeys "{ENTER}"
        s = InputBox("y", , "keep")
        WriteLog "inputbox [" & s & "]"
    Case 4
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub
