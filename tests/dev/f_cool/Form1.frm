VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form1
   Caption         =   "CoolBar"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6600
      Top             =   3000
   End
   Begin ComCtl3.CoolBar CoolBar1
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   1323
      _CBWidth        =   7200
      _CBHeight       =   750
      _Version        =   "6.7.9782"
      Child1          =   "Picture1"
      MinHeight1      =   330
      Width1          =   3000
      NewRow1         =   0   'False
      Child2          =   "Combo1"
      MinHeight2      =   315
      Width2          =   2505
      NewRow2         =   0   'False
      MinHeight3      =   360
      NewRow3         =   -1  'True
      Begin VB.TextBox Text1
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   420
         Width           =   1215
      End
      Begin VB.ComboBox Combo1
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   30
         Width           =   2100
      End
      Begin VB.PictureBox Picture1
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   165
         ScaleHeight     =   330
         ScaleWidth      =   2700
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   2700
      End
   End
   Begin VB.Label Label1
      Caption         =   "below"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' CoolBar (COMCT332.OCX, LinBASIC Studio phase 7.4): its bands from the .frm (Child, MinHeight, Width, NewRow),
' where the band puts its child, Bands (Add, Remove, Item, Count, Clear), the properties of a band and of the
' CoolBar, RowCount and Height, the event HeightChanged

Private mStep As Integer

Private Function Geo(ByVal c As Object) As String
    On Error Resume Next
    Geo = c.Name & "(" & c.Left & "," & c.Top & "," & c.Width & "," & c.Height & ")"
    If Err.Number <> 0 Then Geo = "?" & Err.Number
End Function

Private Function B(ByVal i As Integer) As String
    Dim s As String, x As Object
    On Error Resume Next
    With CoolBar1.Bands(i)
        s = "b" & i & " idx=" & .Index & " pos=" & .Position & " key=[" & .Key & "] cap=[" & .Caption & "] w=" & .Width & _
            " minw=" & .MinWidth & " minh=" & .MinHeight & " h=" & .Height & " nr=" & .NewRow & " vis=" & .Visible & _
            " st=" & .Style & " img=" & .Image & " av=" & .AllowVertical & " ucc=" & .UseCoolbarColors
        Set x = .Child
        If x Is Nothing Then s = s & " child=-" Else s = s & " child=" & Geo(x)
    End With
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description
    B = s
End Function

Private Sub State(ByVal what As String)
    Dim i As Integer
    WriteLog what & ": cb h=" & CoolBar1.Height & " w=" & CoolBar1.Width & " rows=" & CoolBar1.RowCount & _
        " n=" & CoolBar1.Bands.Count & " label=" & Label1.Top & " form=" & ScaleHeight
    For i = 1 To CoolBar1.Bands.Count
        WriteLog "  " & B(i)
    Next
    WriteLog "  free " & Geo(Text1)
End Sub

Private Sub ErrLog(ByVal what As String)
    If Err.Number <> 0 Then WriteLog "  " & what & " ! " & Err.Number & " " & Err.Description: Err.Clear
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    WriteLog "  ev heightchanged " & NewHeight & " h=" & CoolBar1.Height
End Sub

Private Sub CoolBar1_Resize()
    WriteLog "  ev resize h=" & CoolBar1.Height
End Sub

Private Sub Form_Load()
    State "load"
    WriteLog "cb props fixed=" & CoolBar1.FixedOrder & " variant=" & CoolBar1.VariantHeight & " borders=" & _
        CoolBar1.BandBorders & " emboss=" & CoolBar1.EmbossPicture & " orient=" & CoolBar1.Orientation & _
        " back=" & Hex$(CoolBar1.BackColor) & " fore=" & Hex$(CoolBar1.ForeColor) & " pic=" & TypeName(CoolBar1.Picture) & _
        " il=" & TypeName(CoolBar1.ImageList)
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    Static done As Boolean
    If done Then Exit Sub
    done = True
    State "activate"
End Sub

Private Sub Timer1_Timer()
    Dim x As Object
    Timer1.Enabled = False
    mStep = mStep + 1
    On Error Resume Next
    Select Case mStep
    Case 1      ' a band on a new row, then back
        CoolBar1.Bands(2).NewRow = True
        State "band2 newrow"
        CoolBar1.Bands(2).NewRow = False
        State "band2 same row"
    Case 2      ' a band hidden and shown
        CoolBar1.Bands(3).Visible = False
        State "band3 hidden"
        CoolBar1.Bands(3).Visible = True
        State "band3 shown"
    Case 3      ' the properties of a band
        With CoolBar1.Bands(1)
            .Caption = "One"
            .Key = "first"
            .Width = 2000
            .MinWidth = 1500
            .MinHeight = 500
        End With
        ErrLog "band1"
        State "band1 changed"
        WriteLog "item first " & CoolBar1.Bands("first").Index
        ErrLog "item first"
    Case 4      ' Add, a child of its own, Add at 1
        Set x = CoolBar1.Bands.Add(, "k4", "Four", , True)
        ErrLog "add"
        WriteLog "added " & TypeName(x) & " " & x.Index & " " & x.Key
        State "added"
        Set CoolBar1.Bands("k4").Child = Text1
        ErrLog "child"
        State "child text1"
        Set x = CoolBar1.Bands.Add(1, "k0", "Zero")
        ErrLog "add at 1"
        State "added at 1"
    Case 5      ' a band of fixed size; Remove; the errors of Bands
        CoolBar1.Bands("k4").Style = cc3BandFixedSize
        ErrLog "style"
        State "k4 fixed"
        CoolBar1.Bands.Remove "k0"
        ErrLog "remove k0"
        CoolBar1.Bands.Remove "k4"
        ErrLog "remove k4"
        State "removed"
        WriteLog "item 9 " & TypeName(CoolBar1.Bands(9))
        ErrLog "item 9"
        WriteLog "item nope " & TypeName(CoolBar1.Bands("nope"))
        ErrLog "item nope"
        CoolBar1.Bands.Add 9
        ErrLog "add 9"
        CoolBar1.Bands.Add , "first"
        ErrLog "add first again"
        CoolBar1.Bands.Remove 99
        ErrLog "remove 99"
        CoolBar1.Bands(1).Width = -5
        ErrLog "width -5"
        CoolBar1.Bands(1).Style = 7
        ErrLog "style 7"
        State "after errors"
    Case 6      ' the properties of the CoolBar
        CoolBar1.VariantHeight = False
        ErrLog "variant"
        State "variant off"
        CoolBar1.BandBorders = False
        CoolBar1.FixedOrder = True
        CoolBar1.EmbossPicture = True
        ErrLog "props"
        WriteLog "props fixed=" & CoolBar1.FixedOrder & " borders=" & CoolBar1.BandBorders & " emboss=" & CoolBar1.EmbossPicture
        State "borders off"
        CoolBar1.Bands(1).Caption = ""
        CoolBar1.Bands(1).MinWidth = 0
        CoolBar1.Bands(1).Width = 3000
        State "borders off, no caption"
        CoolBar1.Bands(2).Caption = "Two"
        State "borders off, caption two"
        CoolBar1.BandBorders = True
        State "borders on, caption two"
    Case 7      ' the form narrower: the last band of a row
        Me.Width = Me.Width - 1500
        State "narrower"
        Me.Width = Me.Width + 1500
    Case 8      ' Clear
        CoolBar1.Bands.Clear
        ErrLog "clear"
        State "cleared"
        Set x = CoolBar1.Bands.Add()
        ErrLog "add empty"
        State "one band"
    Case 9
        Unload Me
        Exit Sub
    End Select
    ErrLog "step " & mStep
    Timer1.Enabled = True
End Sub
