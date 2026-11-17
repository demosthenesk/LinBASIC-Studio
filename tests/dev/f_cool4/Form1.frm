VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1
   Caption         =   "CoolBar 4"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList IL16
      Left            =   6600
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList IL32
      Left            =   6000
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar CoolBar1
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   1058
      BandCount       =   2
      _CBWidth        =   6000
      _CBHeight       =   600
      _Version        =   "6.7.9839"
      Caption1        =   "Ab"
      Child1          =   "Text1"
      MinHeight1      =   300
      Width1          =   2400
      NewRow1         =   0   'False
      Child2          =   "Text2"
      MinHeight2      =   300
      NewRow2         =   0   'False
      Begin VB.TextBox Text2
         Height          =   300
         Left            =   3000
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   30
         Width           =   1215
      End
      Begin VB.TextBox Text1
         Height          =   300
         Left            =   480
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   30
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' CoolBar 4 (LinBASIC Studio phase 7.4): where the child of a band goes with an image of the ImageList (16 and 32
' pixels, with and without a caption, with and without a gripper), a band alone (a gripper or not), and the
' vertical CoolBar (Orientation, AllowVertical)

Private Function G(ByVal c As Control) As String
    G = c.Name & "(" & c.Left & "," & c.Top & "," & c.Width & "," & c.Height & "," & c.Visible & ")"
End Function

Private Sub State(ByVal what As String)
    Dim i As Integer, s As String
    s = what & ": cb w=" & CoolBar1.Width & " h=" & CoolBar1.Height & " rows=" & CoolBar1.RowCount & " " & G(Text1) & " " & G(Text2)
    For i = 1 To CoolBar1.Bands.Count
        With CoolBar1.Bands(i)
            s = s & " | b" & i & " h=" & .Height & " w=" & .Width & " img=" & .Image & " pos=" & .Position
        End With
    Next
    WriteLog s
End Sub

Private Sub ErrLog(ByVal what As String)
    If Err.Number <> 0 Then WriteLog "  " & what & " ! " & Err.Number & " " & Err.Description: Err.Clear
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    WriteLog "  ev heightchanged " & NewHeight & " w=" & CoolBar1.Width & " h=" & CoolBar1.Height
End Sub

Private Sub CoolBar1_Resize()
    WriteLog "  ev resize w=" & CoolBar1.Width & " h=" & CoolBar1.Height
End Sub

Private Sub Form_Load()
    On Error Resume Next
    IL16.ListImages.Add , "a", LoadPicture(App.Path & "\i16.bmp")
    IL16.ListImages.Add , "b", LoadPicture(App.Path & "\i16.bmp")
    IL32.ListImages.Add , "a", LoadPicture(App.Path & "\i32.bmp")
    ErrLog "images"
    CoolBar1.Bands(1).Image = 1
    ErrLog "image without imagelist"
    WriteLog "load img=" & CoolBar1.Bands(1).Image
End Sub

Private Sub Form_Activate()
    Static done As Boolean
    If done Then Exit Sub
    done = True
    On Error Resume Next
    State "start"
    Set CoolBar1.ImageList = IL16
    ErrLog "imagelist"
    State "imagelist 16, b1 image"
    CoolBar1.Bands(1).Image = 1
    ErrLog "b1 image"
    State "b1 image 1"
    CoolBar1.Bands(1).Image = "b"
    ErrLog "b1 image key"
    State "b1 image b"
    CoolBar1.Bands(1).Image = 9
    ErrLog "b1 image 9"
    CoolBar1.Bands(1).Image = "zz"
    ErrLog "b1 image zz"
    State "b1 image bad"
    CoolBar1.Bands(1).Image = 1
    CoolBar1.Bands(1).Caption = ""
    State "b1 image, no caption"
    CoolBar1.Bands(2).Image = 1
    State "b2 image"
    Set CoolBar1.ImageList = IL32
    ErrLog "imagelist 32"
    State "imagelist 32"
    CoolBar1.Bands(1).Caption = "Ab"
    State "b1 caption again"
    CoolBar1.FixedOrder = True
    State "fixed order"
    CoolBar1.FixedOrder = False
    CoolBar1.Bands(2).Visible = False
    State "one band"
    CoolBar1.Bands(1).Image = 0
    State "one band, no image"
    CoolBar1.Bands(2).Visible = True
    CoolBar1.Bands(2).Image = 0
    Set CoolBar1.ImageList = Nothing
    ErrLog "imagelist nothing"
    State "no imagelist"
    CoolBar1.Orientation = cc3OrientationVertical
    ErrLog "vertical"
    State "vertical"
    CoolBar1.Bands(2).NewRow = True
    State "vertical, b2 new row"
    CoolBar1.Bands(1).AllowVertical = False
    State "vertical, b1 not vertical"
    CoolBar1.Bands(1).AllowVertical = True
    CoolBar1.Bands(2).NewRow = False
    CoolBar1.Orientation = cc3OrientationHorizontal
    State "horizontal"
    Unload Me
End Sub
