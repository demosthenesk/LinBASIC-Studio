VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1
   Caption         =   "StdDataFormat 5"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   5400
      Top             =   120
   End
   Begin MSDataListLib.DataList DataList1
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   1826
      _Version        =   393216
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   1
         Format          =   "000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo1
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox Text1
      DataField       =   "Age"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1
      Height          =   330
      Left            =   120
      Top             =   3000
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
' StdDataFormat 5 (LinBASIC Studio phase 7.7): the DataFormat of a DataList (in the .frm) and of a DataCombo (set in
' code, with events), both bound to the field Age and listing the table Ages (BoundColumn A, ListField Label)

Private mStep As Integer
Private WithEvents fmtC As StdFormat.StdDataFormat
Attribute fmtC.VB_VarHelpID = -1
Private rsAges As ADODB.Recordset

Private Function Cur() As String
    On Error Resume Next
    Cur = "[list bound=" & DataList1.BoundText & " text=" & DataList1.Text & " matched=" & DataList1.MatchedWithList & _
        " | combo bound=" & DataCombo1.BoundText & " text=" & DataCombo1.Text & " matched=" & DataCombo1.MatchedWithList & _
        " | text1=" & Text1.Text & "]"
    If Err.Number <> 0 Then Cur = Cur & " ! " & Err.Number & " " & Err.Description
End Function

Private Function Fld() As String
    Fld = "field Age=" & V(Adodc1.Recordset!Age.Value)
End Function

Private Function DF(ByVal c As Object) As String
    Dim f As Object
    On Error Resume Next
    Set f = c.DataFormat
    DF = c.Name & " " & TypeName(f)
    If Not f Is Nothing Then DF = DF & " type=" & f.Type & " format=""" & f.Format & """"
    If Err.Number <> 0 Then DF = DF & " ! " & Err.Number & " " & Err.Description
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

Private Sub fmtC_Format(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev combo format " & V(DataValue.Value) & " target=" & Tgt(DataValue)
End Sub

Private Sub fmtC_UnFormat(ByVal DataValue As StdFormat.StdDataValue)
    WriteLog "  ev combo unformat " & V(DataValue.Value) & " target=" & Tgt(DataValue)
End Sub

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    WriteLog "  adodc error " & ErrorNumber & " " & Description
    fCancelDisplay = True
End Sub

Private Sub Form_Load()
    WriteLog "form load " & DF(DataList1) & " | " & DF(DataCombo1)
    On Error Resume Next
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dfmt5.mdb"
    Adodc1.RecordSource = "SELECT * FROM People ORDER BY ID"
    Adodc1.Refresh
    ErrLog "refresh"
    Set rsAges = New ADODB.Recordset
    rsAges.Open "SELECT * FROM Ages ORDER BY A", Adodc1.Recordset.ActiveConnection, adOpenStatic, adLockReadOnly
    Set DataList1.RowSource = rsAges
    DataList1.ListField = "Label"
    DataList1.BoundColumn = "A"
    Set DataCombo1.RowSource = rsAges
    DataCombo1.ListField = "Label"
    DataCombo1.BoundColumn = "A"
    Set fmtC = New StdFormat.StdDataFormat
    fmtC.Format = "000"
    Set DataCombo1.DataFormat = fmtC
    ErrLog "combo dataformat"
    WriteLog "lists " & Cur()
    Set DataList1.DataSource = Adodc1
    DataList1.DataField = "Age"
    Set DataCombo1.DataSource = Adodc1
    DataCombo1.DataField = "Age"
    ErrLog "bind"
    WriteLog "bound " & Cur() & " " & DF(DataCombo1)
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
        Adodc1.Recordset.MoveNext
        WriteLog "cleo " & Cur()
        Adodc1.Recordset.MoveFirst
        WriteLog "anna " & Cur()
    Case 2      ' the combo: an item of the list
        DataCombo1.Text = "forty-two"
        WriteLog "combo forty-two " & Cur() & " changed=" & DataCombo1.DataChanged
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "back " & Cur()
        WriteLog Fld()
    Case 3      ' the list: a bound value
        DataList1.BoundText = "25"
        WriteLog "list 25 " & Cur() & " changed=" & DataList1.DataChanged
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "back " & Cur()
        WriteLog Fld()
    Case 4      ' the combo: a bound value written as the format writes it
        DataCombo1.BoundText = "031"
        WriteLog "combo 031 " & Cur() & " changed=" & DataCombo1.DataChanged
        Adodc1.Recordset.MoveNext
        ErrLog "movenext"
        Adodc1.Recordset.MovePrevious
        WriteLog "back " & Cur()
        WriteLog Fld()
    Case 5
        Unload Me
        Exit Sub
    End Select
    ErrLog "step"
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "form unload " & Cur()
End Sub
