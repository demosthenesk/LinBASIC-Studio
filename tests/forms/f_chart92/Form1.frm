VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "Charts"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   1800
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   1800
      Left            =   2640
      OleObjectBlob   =   "Form1.frx":2365
      TabIndex        =   1
      Top             =   120
      Width           =   2400
   End
   Begin MSChart20Lib.MSChart MSChart3 
      Height          =   1800
      Left            =   5160
      OleObjectBlob   =   "Form1.frx":46BB
      TabIndex        =   2
      Top             =   120
      Width           =   2400
   End
   Begin MSChart20Lib.MSChart MSChart4 
      Height          =   1800
      Left            =   120
      OleObjectBlob   =   "Form1.frx":624E
      TabIndex        =   3
      Top             =   2040
      Width           =   2400
   End
   Begin MSChart20Lib.MSChart MSChart5 
      Height          =   1800
      Left            =   2640
      OleObjectBlob   =   "Form1.frx":817C
      TabIndex        =   4
      Top             =   2040
      Width           =   2400
   End
   Begin MSChart20Lib.MSChart MSChart6 
      Height          =   1800
      Left            =   5160
      OleObjectBlob   =   "Form1.frx":A0B7
      TabIndex        =   5
      Top             =   2040
      Width           =   2400
   End
   Begin MSChart20Lib.MSChart MSChart7 
      Height          =   1800
      Left            =   120
      OleObjectBlob   =   "Form1.frx":B828
      TabIndex        =   6
      Top             =   3960
      Width           =   2400
   End
   Begin MSChart20Lib.MSChart MSChart8 
      Height          =   1800
      Left            =   2640
      OleObjectBlob   =   "Form1.frx":D74B
      TabIndex        =   7
      Top             =   3960
      Width           =   2400
   End
   Begin MSChart20Lib.MSChart MSChart9 
      Height          =   1800
      Left            =   5160
      OleObjectBlob   =   "Form1.frx":F670
      TabIndex        =   8
      Top             =   3960
      Width           =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' f_chart92 (phase 9.2): the settings of an MSChart drawn on a form are in its OleObjectBlob (.frx); the runtime
' reads them (lbvc/project/chartblob.py): VB6 wrote the charts 1 - 8 (DEV-9 mc1, mc3, mc4), 9 is a new one of ours

Private Sub Form_Load()
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_chart92.txt"
    On Error GoTo 0
    Dump MSChart1, "1"
    Dump MSChart2, "2"
    Dump MSChart3, "3"
    Dump MSChart4, "4"
    Dump MSChart5, "5"
    Dump MSChart6, "6"
    Dump MSChart7, "7"
    Dump MSChart8, "8"
    Dump MSChart9, "9"
    WriteLog "done"
    Unload Me
End Sub

Private Sub Dump(c As Object, ByVal nm As String)
    Dim s As String, v As Variant, r As Integer, k As Integer, n As Integer
    On Error Resume Next
    v = Empty: v = c.chartType: Rep nm, "type", v
    v = Empty: v = c.ShowLegend: Rep nm, "legend", v
    v = Empty: v = c.Stacking: Rep nm, "stacking", v
    v = Empty: v = c.TitleText: Rep nm, "title", "[" & v & "]"
    v = Empty: v = c.FootnoteText: Rep nm, "footnote", "[" & v & "]"
    v = Empty: v = c.RowCount & "x" & c.ColumnCount: Rep nm, "size", v
    v = Empty: v = c.RowLabelCount & "/" & c.ColumnLabelCount: Rep nm, "label levels", v
    s = ""
    For r = 1 To c.RowCount
        c.Row = r
        s = s & c.RowLabel & ":"
        For k = 1 To c.ColumnCount
            c.Column = k
            s = s & " " & CLng(CDbl(c.Data) * 10)
        Next
        s = s & ";"
    Next
    Rep nm, "data", s
    s = ""
    For k = 1 To c.ColumnCount
        c.Column = k
        s = s & c.ColumnLabel & ","
    Next
    Rep nm, "columns", s
    v = Empty: v = c.Title.VtFont.Name: Rep nm, "title font", v
    v = Empty: v = CLng(c.Title.VtFont.Size * 100): Rep nm, "title size", v
    v = Empty: v = c.Legend.Location.LocationType: Rep nm, "legend loc", v
    v = Empty: v = CLng(c.Plot.View3d.Rotation) & " " & CLng(c.Plot.View3d.Elevation): Rep nm, "view3d", v
    v = Empty: v = c.Plot.Axis(1).ValueScale.Auto: Rep nm, "y auto", v
    v = Empty: v = CLng(c.Plot.Axis(1).ValueScale.Maximum): Rep nm, "y max", v
    n = c.ColumnCount
    If n > 2 Then n = 2
    s = ""
    For k = 1 To n
        With c.Plot.SeriesCollection(k).DataPoints(-1).Brush.FillColor
            s = s & " " & .Red & "," & .Green & "," & .Blue
        End With
    Next
    Rep nm, "colors", s
End Sub

Private Sub Rep(ByVal nm As String, ByVal what As String, ByVal v As Variant)
    If Err.Number <> 0 Then
        WriteLog nm & " " & what & " err " & Err.Number
        Err.Clear
    Else
        WriteLog nm & " " & what & " " & v
    End If
End Sub
