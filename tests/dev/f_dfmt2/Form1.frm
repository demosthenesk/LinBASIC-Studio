VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1
   Caption         =   "StdDataFormat 2"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8400
      Top             =   120
   End
   Begin VB.TextBox txtSalG
      DataField       =   "Salary"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtActive
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "Yes"
         FalseValue      =   "No"
         NullValue       =   "?"
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   7
      EndProperty
      DataField       =   "Active"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtAge2
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "000"
         HaveTrueFalseNull=   1
         TrueValue       =   ""
         FalseValue      =   ""
         NullValue       =   "n/a"
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataField       =   "Age"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtBornG
      DataField       =   "Born"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CheckBox chkAct
      Caption         =   "Active"
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "T"
         FalseValue      =   "F"
         NullValue       =   "N"
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   7
      EndProperty
      DataField       =   "Active"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1
      Bindings        =   "Form1.frx":0000
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Who"
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
         DataField       =   "Age"
         Caption         =   "Age"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Salary"
         Caption         =   "Pay"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Active"
         Caption         =   "On"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Yes"
            FalseValue      =   "No"
            NullValue       =   "?"
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Born"
         Caption         =   "Born"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "yyyy-mm-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   800.126
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1
      Height          =   330
      Left            =   120
      Top             =   4920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "People"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' StdDataFormat 2 (LinBASIC Studio phase 7.7): a DataGrid bound at design time (Bindings) with columns of its own
' and their DataFormat; the values of the events Format / UnFormat with a General / Boolean format; NullValue of
' a custom format; a CheckBox with a Boolean format

Private mStep As Integer
Private WithEvents fmtG As StdFormat.StdDataFormat
Attribute fmtG.VB_VarHelpID = -1
Private WithEvents fmtB As StdFormat.StdDataFormat
Attribute fmtB.VB_VarHelpID = -1
Private WithEvents fmtCol As StdFormat.StdDataFormat
Attribute fmtCol.VB_VarHelpID = -1
Private nCol As Long

Private Function Cur() As String
    Cur = "[" & txtSalG.Text & "|" & txtActive.Text & "|" & txtAge2.Text & "|" & txtBornG.Text & "|" & chkAct.Value & "]"
End Function

Private Function Fld() As String
    Dim r As ADODB.Recordset
    Set r = Adodc1.Recordset
    Fld = "fields ID=" & r!ID & " Name=" & V(r!Name.Value) & " Age=" & V(r!Age.Value) & " Active=" & V(r!Active.Value) & _
        " Born=" & V(r!Born.Value) & " Salary=" & V(r!Salary.Value)
End Function

Private Function Grid() As String
    Dim s As String, i As Integer, t As String
    On Error Resume Next
    s = "grid cols=" & DataGrid1.Columns.Count & " row=" & DataGrid1.Row
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    For i = 0 To DataGrid1.Columns.Count - 1
        With DataGrid1.Columns(i)
            s = s & vbCrLf & "   " & i & " " & .Caption & "(" & .DataField & ")"
            t = "?": t = .Text
            If Err.Number <> 0 Then t = "! " & Err.Number: Err.Clear
            s = s & " text=" & t
            t = "?": t = V(.Value)
            If Err.Number <> 0 Then t = "! " & Err.Number: Err.Clear
            s = s & " value=" & t
            s = s & " nf=""" & .NumberFormat & """ df=" & TypeName(.DataFormat) & ":" & .DataFormat.Type & ":""" & .DataFormat.Format & """"
            s = s & " true=" & V(.DataFormat.TrueValue) & " null=" & V(.DataFormat.NullValue)
            If Err.Number <> 0 Then s = s & " ! " & Err.Number: Err.Clear
            s = s & " align=" & .Alignment & " width=" & .Width & " visible=" & .Visible & " locked=" & .Locked
            If Err.Number <> 0 Then s = s & " ! " & Err.Number: Err.Clear
        End With
    Next
    Grid = s
End Function

Private Sub ErrLog(ByVal what As String)
    If Err.Number <> 0 Then WriteLog "  " & what & " ! " & Err.Number & " " & Err.Description: Err.Clear
End Sub

Private Function Tgt(ByVal dv As StdFormat.StdDataValue) As String
    On Error Resume Next
    Tgt = TypeName(dv.TargetObject)
    Tgt = Tgt & ":" & dv.TargetObject.Name
    If Err.Number <> 0 Then Tgt = Tgt & " ! " & Err.Number
End Function

Private Sub fmtG_Format(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev g format " & V(DataValue.Value) & " target=" & Tgt(DataValue)
End Sub

Private Sub fmtG_UnFormat(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev g unformat " & V(DataValue.Value) & " target=" & Tgt(DataValue)
    If DataValue.Value = "42" Then DataValue.Value = 4242
End Sub

Private Sub fmtB_Format(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev b format " & V(DataValue.Value) & " target=" & Tgt(DataValue)
End Sub

Private Sub fmtB_UnFormat(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev b unformat " & V(DataValue.Value) & " target=" & Tgt(DataValue)
End Sub

Private Sub fmtCol_Format(ByVal DataValue As StdFormat.StdDataValue)
    nCol = nCol + 1
    If nCol <= 4 Then WriteLog "  ev col format " & V(DataValue.Value) & " target=" & Tgt(DataValue)
End Sub

Private Sub fmtCol_UnFormat(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev col unformat " & V(DataValue.Value) & " target=" & Tgt(DataValue)
End Sub

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    WriteLog "  adodc error " & ErrorNumber & " " & Description & " scode=" & Scode & " src=" & Source
    fCancelDisplay = True
End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
    WriteLog "  grid error " & DataError & " " & DataGrid1.ErrorText
    Response = 0
End Sub

Private Sub Form_Load()
    Dim f As New StdFormat.StdDataFormat
    WriteLog "form load new=" & TypeName(f)
    On Error Resume Next
    Set fmtG = txtSalG.DataFormat
    Set fmtB = txtActive.DataFormat
    Set fmtCol = DataGrid1.Columns(1).DataFormat
    ErrLog "set formats"
    WriteLog "col same=" & (fmtCol Is DataGrid1.Columns(1).DataFormat)
    WriteLog "before " & Grid()
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dfmt2.mdb"
    Adodc1.RecordSource = "SELECT * FROM People ORDER BY ID"
    Adodc1.Refresh
    ErrLog "refresh"
    WriteLog "refreshed " & Cur()
    WriteLog Grid()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep & " " & Cur()
    On Error Resume Next
    Select Case mStep
    Case 1
        Adodc1.Recordset.MoveNext
        WriteLog "bill " & Cur()
        WriteLog Grid()
        Adodc1.Recordset.MoveNext
        WriteLog "cleo " & Cur()
        WriteLog Grid()
    Case 2      ' UnFormat of General and Boolean formats
        Adodc1.Recordset.MoveFirst
        txtSalG.Text = "12,5"
        txtActive.Text = "no"
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "back " & Cur()
        WriteLog Fld()
        txtSalG.Text = "42"
        txtActive.Text = "?"
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "back " & Cur()
        WriteLog Fld()
        txtAge2.Text = "n/a"
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "back " & Cur()
        WriteLog Fld()
    Case 3      ' the CheckBox with a Boolean format
        chkAct.Value = 0
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "chk " & Cur()
        WriteLog Fld()
    Case 4      ' values into the grid
        DataGrid1.Columns(1).Text = "12"
        ErrLog "age text"
        DataGrid1.Columns(3).Text = "Yes"
        ErrLog "active text"
        DataGrid1.Columns(2).Text = "4,5"
        ErrLog "pay text"
        WriteLog "edited " & Grid()
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog Fld()
        WriteLog Grid()
    Case 5      ' the columns again: DataSource, ReBind, ClearFields, Refresh
        Set DataGrid1.DataSource = Adodc1
        ErrLog "datasource"
        WriteLog "set source " & Grid()
        DataGrid1.ReBind
        ErrLog "rebind"
        WriteLog "rebind " & Grid()
        Adodc1.Refresh
        ErrLog "refresh"
        WriteLog "refresh " & Grid()
        DataGrid1.ClearFields
        ErrLog "clearfields"
        WriteLog "clearfields " & Grid()
        DataGrid1.ReBind
        ErrLog "rebind"
        WriteLog "rebind " & Grid()
    Case 6
        Unload Me
        Exit Sub
    End Select
    ErrLog "step"
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "form unload " & Cur()
End Sub
