VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmChart 
   Caption         =   "Sales chart"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid grdSales 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   13
      Cols            =   4
      AllowUserResizing=   1
   End
   Begin MSChart20Lib.MSChart chtSales 
      Height          =   5295
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin VB.Label lblType 
      Caption         =   "Chart type:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   1035
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Chart - MSFlexGrid and MSChart: a table of numbers and its chart, a 2D array as ChartData
Option Explicit

Private Sub Form_Load()
    Dim data(0 To 12, 0 To 3) As Variant, m As Integer, c As Integer
    data(0, 1) = "Books": data(0, 2) = "Music": data(0, 3) = "Games"   ' row 0: the names of the series
    For m = 1 To 12
        data(m, 0) = Format$(DateSerial(2026, m, 1), "mmm")           ' column 0: the labels of the rows
        data(m, 1) = Round(40 + 25 * Sin(m / 2) + m * 3, 1)
        data(m, 2) = Round(30 + 15 * Cos(m / 3) + m, 1)
        data(m, 3) = Round(20 + m * m / 3, 1)
    Next
    With grdSales                                ' the table
        .Rows = 13: .Cols = 4
        .ColWidth(0) = 600
        For m = 0 To 12
            For c = 0 To 3
                If m > 0 And c > 0 Then
                    .TextMatrix(m, c) = Format$(data(m, c), "0.0")
                ElseIf Not IsEmpty(data(m, c)) Then
                    .TextMatrix(m, c) = data(m, c)
                End If
            Next
        Next
        For c = 1 To 3
            .ColWidth(c) = 800
            .ColAlignment(c) = flexAlignRightCenter
        Next
    End With
    chtSales.ChartData = data                    ' its chart: the labels come with the numbers
    chtSales.TitleText = "Sales 2026"
    chtSales.ShowLegend = True
    cboType.AddItem "Bars"
    cboType.AddItem "Lines"
    cboType.AddItem "Area"
    cboType.AddItem "Steps"
    cboType.ListIndex = 0
End Sub

Private Sub cboType_Click()
    Select Case cboType.ListIndex
    Case 0: chtSales.chartType = VtChChartType2dBar
    Case 1: chtSales.chartType = VtChChartType2dLine
    Case 2: chtSales.chartType = VtChChartType2dArea
    Case 3: chtSales.chartType = VtChChartType2dStep
    End Select
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Or ScaleWidth < 5000 Or ScaleHeight < 2000 Then Exit Sub
    grdSales.Height = ScaleHeight - grdSales.Top - 120
    chtSales.Move chtSales.Left, chtSales.Top, ScaleWidth - chtSales.Left - 120, ScaleHeight - 240
End Sub
