VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Est"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1500
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   2646
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "C0"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   "C1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         Size            =   1500
         RecordSelectors =   0   'False
         AllowSizing     =   0   'False
         AllowRowSizing  =   0   'False
         AllowFocus      =   0   'False
         ScrollGroup     =   2
         ScrollBars      =   2
         MarqueeStyle    =   3
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1000
            Alignment       =   1
            DividerStyle    =   0
            Button          =   -1  'True
            WrapText        =   -1  'True
            Object.Visible         =   0   'False
            Locked          =   -1  'True
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1500
      Left            =   120
      TabIndex        =   1
      Top             =   1700
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   2646
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "C0"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   "C1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowSizing     =   0   'False
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   0
            ColumnAllowSizing=   0   'False
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f As Integer
Private Sub P(ByVal s As String)
    Print #f, s
End Sub
Private Function V(o As Object, ByVal n As String) As String
    On Error Resume Next
    Dim r As Variant
    r = CallByName(o, n, VbGet)
    If Err.Number <> 0 Then V = n & "=!" & Err.Number Else V = n & "=" & CStr(r)
End Function
Private Sub Grid(g As Object)
    Dim i As Integer, s As String, c As Object, sp As Object, k
    P g.Name & ": " & V(g, "AllowRowSizing") & " " & V(g, "RecordSelectors") & " " & V(g, "MarqueeStyle") & " " & V(g, "ScrollBars") & " " & V(g, "Splits.Count")
    On Error Resume Next
    P "  splits=" & g.Splits.Count & " err=" & Err.Number: Err.Clear
    Set sp = g.Splits(0)
    If Err.Number <> 0 Then P "  split0 !" & Err.Number: Err.Clear
    s = "  split0:"
    For Each k In Array("AllowSizing", "Locked", "MarqueeStyle", "ScrollBars", "ScrollGroup", "Size", "SizeMode", "AllowFocus", "AllowRowSizing", "RecordSelectors")
        s = s & " " & V(sp, CStr(k))
    Next
    P s
    For i = 0 To g.Columns.Count - 1
        Set c = g.Columns(i)
        s = "  col" & i & ":"
        For Each k In Array("AllowSizing", "WrapText", "Button", "DividerStyle", "Locked", "Visible", "Alignment", "Width")
            s = s & " " & V(c, CStr(k))
        Next
        P s
    Next
    s = "  split0 cols:"
    For i = 0 To sp.Columns.Count - 1
        s = s & " " & V(sp.Columns(i), "AllowSizing") & "/" & V(sp.Columns(i), "Locked")
    Next
    P s
End Sub
Private Sub Form_Load()
    f = FreeFile
    Open App.Path & "\expected\f_dgsplit.txt" For Output As #f
    Grid DataGrid1
    Grid DataGrid2
    On Error Resume Next
    DataGrid1.Columns(1).AllowSizing = False
    P "set col1 AllowSizing False: err=" & Err.Number & " " & V(DataGrid1.Columns(1), "AllowSizing"): Err.Clear
    DataGrid1.Splits(0).AllowSizing = True
    P "set split AllowSizing True: err=" & Err.Number & " " & V(DataGrid1.Splits(0), "AllowSizing") & " col0 " & V(DataGrid1.Columns(0), "AllowSizing") & " col1 " & V(DataGrid1.Columns(1), "AllowSizing"): Err.Clear
    DataGrid2.Columns(1).WrapText = True: DataGrid2.Columns(1).Button = True: DataGrid2.Columns(1).DividerStyle = 3
    P "set col1 wrap/button/div: err=" & Err.Number & " " & V(DataGrid2.Columns(1), "WrapText") & " " & V(DataGrid2.Columns(1), "Button") & " " & V(DataGrid2.Columns(1), "DividerStyle"): Err.Clear
    P "end"
    Close #f
    End
End Sub
