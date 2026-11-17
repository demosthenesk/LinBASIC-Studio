VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1
   Caption         =   "StdDataFormat"
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
   Begin VB.TextBox txtName
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtAge
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataField       =   "Age"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
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
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtChk
      BeginProperty DataFormat 
         Type            =   4
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   8
      EndProperty
      DataField       =   "Active"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtBorn
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-mm-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataField       =   "Born"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtSal1
      BeginProperty DataFormat 
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
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtSal2
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1032
         SubFormatType   =   1
      EndProperty
      DataField       =   "Salary"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtSalG
      DataField       =   "Salary"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblNotes
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   ""
         FalseValue      =   ""
         NullValue       =   "(null)"
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      DataField       =   "Notes"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblAge
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "\A\g\e 0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   10
      EndProperty
      DataField       =   "Age"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtDao
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataField       =   "Age"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtBind
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   480
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
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
         Caption         =   "Name"
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
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
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
   Begin VB.Data Data1
      Caption         =   "People (DAO)"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' StdDataFormat (MSSTDFMT.DLL, LinBASIC Studio phase 7.7): the DataFormat of bound controls (from the .frm and
' set in code), Type / Format / TrueValue / FalseValue / NullValue, the events Format / UnFormat / Changed,
' StdDataFormats, the DataFormat of a Binding (BindingCollection), of a DataGrid column, with a DAO Data control

Private mStep As Integer
Private WithEvents fmtAge As StdFormat.StdDataFormat
Attribute fmtAge.VB_VarHelpID = -1
Private WithEvents fmtName As StdFormat.StdDataFormat
Attribute fmtName.VB_VarHelpID = -1
Private bc As Object
Private rs2 As ADODB.Recordset

Private Function Cur() As String
    Cur = "[" & txtName.Text & "|" & txtAge.Text & "|" & txtActive.Text & "|" & txtChk.Text & "|" & txtBorn.Text & "|" & _
        txtSal1.Text & "|" & txtSal2.Text & "|" & txtSalG.Text & "|" & lblNotes.Caption & "|" & lblAge.Caption & "]"
End Function

Private Function Fld() As String
    Dim r As ADODB.Recordset
    Set r = Adodc1.Recordset
    Fld = "fields ID=" & r!ID & " Name=" & V(r!Name.Value) & " Age=" & V(r!Age.Value) & " Active=" & V(r!Active.Value) & _
        " Born=" & V(r!Born.Value) & " Salary=" & V(r!Salary.Value) & " Notes=" & V(r!Notes.Value)
End Function

Private Function DF(ByVal c As Object) As String
    Dim f As Object, s As String
    On Error Resume Next
    Set f = c.DataFormat
    If Err.Number <> 0 Then DF = c.Name & " ! " & Err.Number & " " & Err.Description: Exit Function
    s = c.Name & " " & TypeName(f)
    If Not f Is Nothing Then s = s & " " & DFo(f)
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description
    DF = s
End Function

Private Function DFo(ByVal f As Object) As String
    On Error Resume Next
    DFo = "type=" & f.Type & " format=""" & f.Format & """ true=" & V(f.TrueValue) & " false=" & V(f.FalseValue) & _
        " null=" & V(f.NullValue) & " day=" & f.FirstDayOfWeek & " week=" & f.FirstWeekOfYear
    If Err.Number <> 0 Then DFo = DFo & " ! " & Err.Number & " " & Err.Description
End Function

Private Function Grid() As String
    Dim s As String, i As Integer
    On Error Resume Next
    s = "grid cols=" & DataGrid1.Columns.Count
    For i = 0 To DataGrid1.Columns.Count - 1
        With DataGrid1.Columns(i)
            s = s & " " & .Caption & "(" & .DataField & ")=" & .Text & "/" & .Value & " nf=""" & .NumberFormat & """ df=" & TypeName(.DataFormat)
            If Not .DataFormat Is Nothing Then s = s & ":""" & .DataFormat.Format & """"
        End With
        If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    Next
    s = s & " row=" & DataGrid1.Row
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description
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

Private Sub fmtAge_Format(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev age format " & V(DataValue.Value) & " target=" & Tgt(DataValue)
End Sub

Private Sub fmtAge_UnFormat(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev age unformat " & V(DataValue.Value) & " target=" & Tgt(DataValue)
End Sub

Private Sub fmtAge_Changed()
    WriteLog "  ev age changed"
End Sub

Private Sub fmtName_Format(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev name format " & V(DataValue.Value) & " target=" & Tgt(DataValue)
    If Not IsNull(DataValue.Value) Then DataValue.Value = "<" & DataValue.Value & ">"
End Sub

Private Sub fmtName_UnFormat(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev name unformat " & V(DataValue.Value) & " target=" & Tgt(DataValue)
    If IsNull(DataValue.Value) Then
        Exit Sub
    ElseIf Left$(DataValue.Value, 1) = "<" Then
        DataValue.Value = Mid$(DataValue.Value, 2, Len(DataValue.Value) - 2)
    Else
        DataValue.Value = UCase$(DataValue.Value)
    End If
End Sub

Private Sub fmtName_Changed()
    WriteLog "  ev name changed " & DFo(fmtName)
End Sub

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    WriteLog "  adodc error " & ErrorNumber & " " & Description & " scode=" & Scode & " src=" & Source
    fCancelDisplay = True
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
    WriteLog "  data1 error " & DataErr & " " & Error$(DataErr)
    Response = 0
End Sub

Private Sub Form_Load()
    Dim c As Control
    WriteLog "form load"
    For Each c In Controls
        If TypeOf c Is TextBox Or TypeOf c Is Label Then WriteLog "df " & DF(c)
    Next
    On Error Resume Next
    Set fmtAge = txtAge.DataFormat
    ErrLog "set fmtAge"
    WriteLog "fmtAge " & TypeName(fmtAge) & " same=" & (fmtAge Is txtAge.DataFormat)
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dfmt.mdb"
    Adodc1.RecordSource = "SELECT * FROM People ORDER BY ID"
    Adodc1.Refresh
    ErrLog "refresh"
    WriteLog "refreshed " & Cur()
    Set DataGrid1.DataSource = Adodc1
    ErrLog "grid source"
    WriteLog Grid()
    Data1.DatabaseName = App.Path & "\dfmt.mdb"
    Data1.RecordSource = "SELECT * FROM People ORDER BY ID"
    Data1.Refresh
    ErrLog "data1 refresh"
    WriteLog "dao [" & txtDao.Text & "] df " & DF(txtDao)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim f1 As StdFormat.StdDataFormat, f2 As StdFormat.StdDataFormat, fs As Object, x As Variant
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep & " " & Cur()
    On Error Resume Next
    Select Case mStep
    Case 1      ' the values of the formats: Null, a record without Age
        Adodc1.Recordset.MoveNext
        WriteLog "bill " & Cur()
        WriteLog Grid()
        Adodc1.Recordset.MoveNext
        WriteLog "cleo " & Cur()
        WriteLog Grid()
    Case 2      ' the values back into the fields (UnFormat)
        Adodc1.Recordset.MoveFirst
        txtAge.Text = "7"
        txtActive.Text = "No"
        txtBorn.Text = "2001-02-03"
        txtSal1.Text = "12.5"
        WriteLog "edited " & Cur() & " changed=" & txtAge.DataChanged & txtActive.DataChanged & txtBorn.DataChanged & txtSal1.DataChanged
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        ErrLog "moveprevious"
        WriteLog "back " & Cur()
        WriteLog Fld()
    Case 3
        txtSal2.Text = "3,25"
        txtActive.Text = "Yes"
        txtChk.Text = "0"
        txtName.Text = "Anna2"
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "back " & Cur()
        WriteLog Fld()
    Case 4      ' values that do not fit
        txtActive.Text = "Maybe"
        txtAge.Text = "abc"
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        WriteLog "moved " & Cur()
        Adodc1.Recordset.MoveFirst
        ErrLog "movefirst"
        WriteLog "first " & Cur()
        WriteLog Fld()
    Case 5      ' a format of the program, with events
        Set fmtName = New StdFormat.StdDataFormat
        WriteLog "new " & DFo(fmtName)
        Set txtName.DataFormat = fmtName
        ErrLog "set dataformat"
        WriteLog "set " & Cur() & " same=" & (txtName.DataFormat Is fmtName)
        Adodc1.Recordset.MoveNext
        WriteLog "next " & Cur()
    Case 6
        txtName.Text = "zed"
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "back " & Cur()
        WriteLog Fld()
        txtName.Text = "<kim>"
        Adodc1.Recordset.MoveNext
        Adodc1.Recordset.MovePrevious
        WriteLog Fld()
    Case 7      ' the properties: Changed, values that do not fit
        fmtName.Type = fmtCustom
        ErrLog "type custom"
        fmtName.Format = "@@@@@@!"
        ErrLog "format"
        WriteLog "custom " & Cur()
        fmtName.TrueValue = 5
        fmtName.FalseValue = "F"
        fmtName.NullValue = Null
        ErrLog "values"
        WriteLog "values " & DFo(fmtName)
        fmtName.Type = 9
        ErrLog "type 9"
        fmtName.Type = -1
        ErrLog "type -1"
        fmtName.FirstDayOfWeek = 8
        ErrLog "day 8"
        fmtName.FirstWeekOfYear = 4
        ErrLog "week 4"
        fmtName.FirstDayOfWeek = fmtMonday
        fmtName.FirstWeekOfYear = fmtFirstFullWeek
        ErrLog "day/week"
        WriteLog "props " & DFo(fmtName)
        fmtName.Type = fmtGeneral
        fmtName.Format = ""
        Adodc1.Recordset.MoveNext
        Adodc1.Recordset.MovePrevious
        WriteLog "general " & Cur()
    Case 8      ' a format of the .frm changed in code; Nothing
        fmtAge.Format = "0000"
        ErrLog "age format"
        WriteLog "age 0000 " & Cur()
        Adodc1.Recordset.MoveNext
        Adodc1.Recordset.MovePrevious
        WriteLog "moved " & Cur()
        Set txtAge.DataFormat = Nothing
        ErrLog "nothing"
        WriteLog "nothing " & DF(txtAge) & " " & Cur()
        Adodc1.Recordset.MoveNext
        Adodc1.Recordset.MovePrevious
        WriteLog "moved " & Cur()
        Set f1 = New StdFormat.StdDataFormat
        f1.Type = fmtBoolean
        f1.TrueValue = "T"
        f1.FalseValue = "F"
        f1.NullValue = "N"
        Set lblAge.DataFormat = f1
        Set txtChk.DataFormat = f1
        ErrLog "boolean"
        Adodc1.Recordset.MoveNext
        Adodc1.Recordset.MovePrevious
        WriteLog "boolean " & Cur()
        Adodc1.Recordset.MoveLast
        WriteLog "last " & Cur()
    Case 9      ' StdDataFormats
        Set fs = New StdFormat.StdDataFormats
        WriteLog "fs " & TypeName(fs) & " count=" & fs.Count
        Set f1 = New StdFormat.StdDataFormat
        f1.Format = "0.0"
        Set f2 = New StdFormat.StdDataFormat
        f2.Format = "#"
        fs.Add f1, "Age"
        ErrLog "add age"
        fs.Add f2, "Salary"
        ErrLog "add salary"
        fs.Add f2
        ErrLog "add no index"
        fs.Add f1, "Age"
        ErrLog "add age again"
        WriteLog "count=" & fs.Count
        WriteLog "item age " & fs("Age").Format & " salary " & fs.Item("salary").Format
        ErrLog "items"
        WriteLog "item 0 " & fs(0).Format
        ErrLog "item 0"
        WriteLog "item 1 " & fs(1).Format
        ErrLog "item 1"
        WriteLog "item 2 " & fs(2).Format
        ErrLog "item 2"
        WriteLog "item nope " & TypeName(fs("Nope"))
        ErrLog "item nope"
        For Each x In fs
            WriteLog "  each " & TypeName(x) & " " & x.Format
        Next
        ErrLog "for each"
        fs.Remove "Age"
        ErrLog "remove"
        WriteLog "removed count=" & fs.Count
        fs.Remove "Nope"
        ErrLog "remove nope"
        fs.Clear
        WriteLog "cleared count=" & fs.Count
    Case 10     ' a Binding with a DataFormat
        Set rs2 = New ADODB.Recordset
        rs2.Open "SELECT * FROM People ORDER BY ID", Adodc1.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
        Set bc = New BindingCollection
        Set bc.DataSource = rs2
        Set f1 = New StdFormat.StdDataFormat
        f1.Format = "0.000"
        bc.Add txtBind, "Text", "Salary", f1, "sal"
        ErrLog "bc add"
        WriteLog "bind [" & txtBind.Text & "] df=" & TypeName(bc("sal").DataFormat) & " " & DFo(bc("sal").DataFormat)
        rs2.MoveLast
        WriteLog "last [" & txtBind.Text & "]"
        txtBind.Text = "7.5"
        rs2.MoveFirst
        ErrLog "movefirst"
        rs2.MoveLast
        WriteLog "back [" & txtBind.Text & "] field=" & V(rs2!Salary.Value)
        Set f2 = New StdFormat.StdDataFormat
        f2.Format = "#"
        Set bc("sal").DataFormat = f2
        ErrLog "binding dataformat"
        WriteLog "set [" & txtBind.Text & "]"
        rs2.MoveFirst
        WriteLog "first [" & txtBind.Text & "]"
    Case 11     ' the Data control (DAO)
        Data1.Recordset.MoveNext
        WriteLog "dao next [" & txtDao.Text & "]"
        txtDao.Text = "5"
        Data1.Recordset.MovePrevious
        ErrLog "dao moveprevious"
        WriteLog "dao back [" & txtDao.Text & "] field=" & V(Data1.Recordset!Age.Value)
    Case 12     ' a format for a column of the grid in code
        Adodc1.Recordset.MoveFirst
        WriteLog Grid()
        Set f1 = New StdFormat.StdDataFormat
        f1.Format = "0000"
        Set DataGrid1.Columns(1).DataFormat = f1
        ErrLog "column dataformat"
        WriteLog Grid()
        DataGrid1.Columns(2).NumberFormat = "0.0"
        ErrLog "numberformat"
        WriteLog Grid()
        Adodc1.Recordset.MoveNext
        WriteLog Grid()
        WriteLog "cell " & DataGrid1.Columns(1).CellText(DataGrid1.Bookmark) & " / " & DataGrid1.Columns(2).CellValue(DataGrid1.Bookmark)
        ErrLog "cell"
    Case 13
        Unload Me
        Exit Sub
    End Select
    ErrLog "step"
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "form unload " & Cur()
End Sub
