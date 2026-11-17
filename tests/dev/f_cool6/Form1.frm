VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "CoolBar 6"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   635
      ImageList       =   "ImageList1"
      _CBWidth        =   7800
      _CBHeight       =   360
      _Version        =   "6.7.9839"
      Image1          =   "1"
      Child1          =   "Text1"
      MinHeight1      =   300
      Width1          =   3000
      NewRow1         =   0   'False
      BandBackColor2  =   -2147483638
      Child2          =   "Text2"
      MinHeight2      =   300
      Width2          =   2400
      BandPicture2    =   "Form1.frx":0352
      UseCoolbarPicture2=   0   'False
      NewRow2         =   0   'False
      BandEmbossHighlight2=   16776960
      BandEmbossShadow2=   8388863
      BandEmbossPicture2=   -1  'True
      MinHeight3      =   300
      NewRow3         =   0   'False
      Visible3        =   0   'False
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   3195
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   30
         Width           =   2205
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   225
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   30
         Width           =   2745
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' CoolBar 6 (LinBASIC Studio phase 7.4): the properties of the bands that the VB6 IDE writes for an image, a picture,
' a hidden band and the emboss colours (set in the property pages, then saved), read back at run time

Private Function G(ByVal c As Control) As String
    G = c.Name & "(" & c.Left & "," & c.Top & "," & c.Width & "," & c.Height & "," & c.Visible & ")"
End Function

Private Sub Form_Activate()
    Dim i As Integer, s As String
    On Error Resume Next
    WriteLog "cb h=" & CoolBar1.Height & " rows=" & CoolBar1.RowCount & " il=" & TypeName(CoolBar1.ImageList) & _
        " images=" & ImageList1.ListImages.Count & " pic=" & TypeName(CoolBar1.Picture) & " emboss=" & CoolBar1.EmbossPicture & _
        " hi=" & Hex$(CoolBar1.EmbossHighlight) & " sh=" & Hex$(CoolBar1.EmbossShadow) & " " & G(Text1) & " " & G(Text2)
    For i = 1 To CoolBar1.Bands.Count
        With CoolBar1.Bands(i)
            s = "b" & i & " img=" & .Image & " vis=" & .Visible & " pic=" & TypeName(.Picture)
            If Not .Picture Is Nothing Then s = s & " picw=" & .Picture.Width
            s = s & " emboss=" & .EmbossPicture & " hi=" & Hex$(.EmbossHighlight) & " sh=" & Hex$(.EmbossShadow) & _
                " ucp=" & .UseCoolbarPicture & " fixedbg=" & .FixedBackground & " h=" & .Height & " w=" & .Width
        End With
        If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
        WriteLog s
    Next
    Unload Me
End Sub
