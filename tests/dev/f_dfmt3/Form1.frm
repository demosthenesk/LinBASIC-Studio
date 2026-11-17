VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1
   Caption         =   "StdDataFormat 3"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8400
      Top             =   120
   End
   Begin VB.TextBox txtSal
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataField       =   "Salary"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1
      Bindings        =   "Form1.frx":0000
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3413
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2
      Bindings        =   "Form1.frx":0015
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3413
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
      ColumnCount     =   3
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
         DataField       =   "Nope"
         Caption         =   "Nope"
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
      BeginProperty Column02 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid3
      Bindings        =   "Form1.frx":002A
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3413
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Salary"
         Caption         =   "Pay"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column01 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1
      Height          =   330
      Left            =   120
      Top             =   6720
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dfmt3.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dfmt3.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM People ORDER BY ID"
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
' StdDataFormat 3 (LinBASIC Studio phase 7.7): DataGrids bound at design time to an ADO Data control that opens by
' itself (ConnectionString and RecordSource in the .frm): do they keep the columns of the design, with a field
' that does not exist, with fewer columns; Refresh, HoldFields, ClearFields, ReBind; the values of the cells

Private mStep As Integer

Private Function G(ByVal gr As Object) As String
    Dim s As String, i As Integer, t As String
    On Error Resume Next
    s = gr.Name & " cols=" & gr.Columns.Count & " row=" & gr.Row
    If Err.Number <> 0 Then s = s & " ! " & Err.Number: Err.Clear
    For i = 0 To gr.Columns.Count - 1
        With gr.Columns(i)
            s = s & " | " & .Caption & "(" & .DataField & ")"
            t = "?": t = .Text
            If Err.Number <> 0 Then t = "!" & Err.Number: Err.Clear
            s = s & "=" & t
            t = "?": t = V(.Value)
            If Err.Number <> 0 Then t = "!" & Err.Number: Err.Clear
            s = s & "/" & t & " [" & .DataFormat.Type & ":" & .DataFormat.Format & "] w=" & .Width
            If Err.Number <> 0 Then s = s & " !" & Err.Number: Err.Clear
        End With
    Next
    G = s
End Function

Private Sub LogAll(ByVal what As String)
    WriteLog what
    WriteLog "  " & G(DataGrid1)
    WriteLog "  " & G(DataGrid2)
    WriteLog "  " & G(DataGrid3)
End Sub

Private Sub ErrLog(ByVal what As String)
    If Err.Number <> 0 Then WriteLog "  " & what & " ! " & Err.Number & " " & Err.Description: Err.Clear
End Sub

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    WriteLog "  adodc error " & ErrorNumber & " " & Description
    fCancelDisplay = True
End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
    WriteLog "  grid1 error " & DataError & " " & DataGrid1.ErrorText
    Response = 0
End Sub

Private Sub DataGrid2_Error(ByVal DataError As Integer, Response As Integer)
    WriteLog "  grid2 error " & DataError & " " & DataGrid2.ErrorText
    Response = 0
End Sub

Private Sub DataGrid3_Error(ByVal DataError As Integer, Response As Integer)
    WriteLog "  grid3 error " & DataError & " " & DataGrid3.ErrorText
    Response = 0
End Sub

Private Sub Form_Load()
    WriteLog "form load rs=" & TypeName(Adodc1.Recordset) & " sal=[" & txtSal.Text & "]"
    LogAll "load"
End Sub

Private Sub Form_Activate()
    Static done As Boolean
    If done Then Exit Sub
    done = True
    WriteLog "activate rs=" & TypeName(Adodc1.Recordset) & " sal=[" & txtSal.Text & "]"
    LogAll "activate"
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep & " sal=[" & txtSal.Text & "]"
    On Error Resume Next
    Select Case mStep
    Case 1
        Adodc1.Recordset.MoveNext
        LogAll "bill"
        Adodc1.Recordset.MoveNext
        LogAll "cleo"
    Case 2
        Adodc1.Recordset.MoveFirst
        DataGrid1.Columns(1).Text = "7"
        ErrLog "age 7"
        DataGrid1.Columns(2).Text = "3,5"
        ErrLog "pay 3,5"
        DataGrid1.Columns(3).Text = "No"
        ErrLog "on No"
        LogAll "edited"
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "fields Age=" & V(Adodc1.Recordset!Age.Value) & " Salary=" & V(Adodc1.Recordset!Salary.Value) & " Active=" & V(Adodc1.Recordset!Active.Value)
        LogAll "back"
    Case 3
        DataGrid3.HoldFields
        ErrLog "holdfields"
        Adodc1.Refresh
        ErrLog "refresh"
        LogAll "refresh"
        DataGrid1.ReBind
        ErrLog "rebind"
        LogAll "rebind 1"
    Case 4
        DataGrid1.ClearFields
        ErrLog "clearfields"
        LogAll "clearfields 1"
        DataGrid1.ReBind
        ErrLog "rebind"
        LogAll "rebind 1"
        Adodc1.RecordSource = "SELECT ID, Name FROM People ORDER BY ID"
        Adodc1.Refresh
        ErrLog "refresh"
        LogAll "two fields"
    Case 5
        Unload Me
        Exit Sub
    End Select
    ErrLog "step"
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "form unload"
End Sub
