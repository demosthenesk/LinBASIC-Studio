VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form1
   Caption         =   "CoolBar 2"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin ComCtl3.CoolBar CoolBar1
      Height          =   1200
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   2117
      BandCount       =   2
      BandBorders     =   0   'False
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      BackColor       =   12632256
      ForeColor       =   255
      _CBWidth        =   6000
      _CBHeight       =   1200
      _Version        =   "6.7.9782"
      Caption1        =   "Cap"
      Key1            =   "kk"
      Tag1            =   "tg"
      Child1          =   "Text1"
      MinWidth1       =   600
      MinHeight1      =   300
      Width1          =   2400
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Style2          =   1
      UseCoolbarColors2=   0   'False
      BackColor2      =   65535
      ForeColor2      =   16711680
      EmbossPicture2  =   -1  'True
      FixedBackground2=   0   'False
      UseCoolbarPicture2=   0   'False
      MinHeight2      =   360
      NewRow2         =   -1  'True
      Visible2        =   -1  'True
      Begin VB.TextBox Text1
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   30
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' CoolBar 3 (LinBASIC Studio phase 7.4; the first .frm of f_cool2, before the VB6 IDE saved it): the names of the properties of the bands in the .frm (Caption1, Key1,
' Tag1, Style2, BackColor2 ...), a CoolBar that is not aligned, a child that is not in the CoolBar

Private Sub Form_Load()
    Dim i As Integer, s As String
    On Error Resume Next
    s = "cb l=" & CoolBar1.Left & " t=" & CoolBar1.Top & " w=" & CoolBar1.Width & " h=" & CoolBar1.Height & " fixed=" & _
        CoolBar1.FixedOrder & " borders=" & CoolBar1.BandBorders & " variant=" & CoolBar1.VariantHeight & " emboss=" & _
        CoolBar1.EmbossPicture & " back=" & Hex$(CoolBar1.BackColor) & " fore=" & Hex$(CoolBar1.ForeColor)
    WriteLog s
    For i = 1 To CoolBar1.Bands.Count
        With CoolBar1.Bands(i)
            s = "b" & i & " cap=[" & .Caption & "] key=[" & .Key & "] tag=[" & .Tag & "] st=" & .Style & " ucc=" & _
                .UseCoolbarColors & " back=" & Hex$(.BackColor) & " fore=" & Hex$(.ForeColor) & " emboss=" & .EmbossPicture & _
                " fixedbg=" & .FixedBackground & " ucp=" & .UseCoolbarPicture & " av=" & .AllowVertical & " minw=" & _
                .MinWidth & " minh=" & .MinHeight & " w=" & .Width & " nr=" & .NewRow & " vis=" & .Visible & " pic=" & TypeName(.Picture)
        End With
        If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
        WriteLog s
    Next
    Set CoolBar1.Bands(2).Child = Command1
    WriteLog "child outside ! " & Err.Number & " " & Err.Description & " | " & TypeName(CoolBar1.Bands(2).Child)
    Err.Clear
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    WriteLog "activate h=" & CoolBar1.Height & " rows=" & CoolBar1.RowCount & " text1=" & Text1.Left & "," & Text1.Top & "," & _
        Text1.Width & "," & Text1.Height & " cmd=" & Command1.Left & "," & Command1.Top
    Unload Me
End Sub
