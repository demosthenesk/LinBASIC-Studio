VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1
   Caption         =   "ADO Data control"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6600
      Top             =   120
   End
   Begin MSDataListLib.DataCombo DataCombo1
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataList DataList1
      Height          =   1035
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   1826
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
   End
   Begin VB.CheckBox Check1
      Caption         =   "Active"
      DataField       =   "Active"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2
      DataField       =   "Age"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   2415
   End
   Begin VB.TextBox Text1
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1
      Height          =   330
      Left            =   120
      Top             =   4560
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
   Begin VB.Label Label1
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' the ADO Data control (LinBASIC Studio, phase 4.4): its recordset and events, the controls bound to it,
' DataGrid, DataList and DataCombo

Private mStep As Integer
Private rsTowns As ADODB.Recordset

Private Function Cur() As String
    Cur = "[" & Text1.Text & "|" & Text2.Text & "|" & Check1.Value & "|" & Label1.Caption & "]"
End Function

Private Function Grid() As String
    Dim s As String, i As Integer
    On Error Resume Next
    s = "grid cols=" & DataGrid1.Columns.Count
    For i = 0 To DataGrid1.Columns.Count - 1
        s = s & " " & DataGrid1.Columns(i).Caption & "=" & DataGrid1.Columns(i).Text
    Next
    s = s & " row=" & DataGrid1.Row & " approx=" & DataGrid1.ApproxCount
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description
    Grid = s
End Function

Private Sub Form_Load()
    WriteLog "form load " & Adodc1.Caption & " cursor=" & Adodc1.CursorType & " lock=" & Adodc1.LockType & " loc=" & Adodc1.CursorLocation & " cmd=" & Adodc1.CommandType & " eof=" & Adodc1.EOFAction
    On Error Resume Next
    WriteLog "rs in load " & TypeName(Adodc1.Recordset)
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    On Error GoTo 0
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\people.mdb"
    Adodc1.RecordSource = "SELECT * FROM People ORDER BY ID"
    Adodc1.Refresh
    WriteLog "refreshed " & Cur() & " rs=" & TypeName(Adodc1.Recordset) & " count=" & Adodc1.Recordset.RecordCount
    Set DataGrid1.DataSource = Adodc1
    WriteLog Grid()
    Set rsTowns = New ADODB.Recordset
    rsTowns.Open "SELECT * FROM Towns ORDER BY TownID", Adodc1.Recordset.ActiveConnection, adOpenStatic, adLockReadOnly
    Set DataList1.RowSource = rsTowns
    DataList1.ListField = "Town"
    DataList1.BoundColumn = "TownID"
    Set DataCombo1.RowSource = rsTowns
    DataCombo1.ListField = "Town"
    DataCombo1.BoundColumn = "TownID"
    WriteLog "list text=[" & DataList1.Text & "] bound=" & DataList1.BoundText & " combo=[" & DataCombo1.Text & "]"
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    WriteLog "form activate " & Cur()
End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    WriteLog "  willmove " & adReason & " " & Cur() & " changed=" & Text1.DataChanged & "/" & Text2.DataChanged
End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    WriteLog "  movecomplete " & adReason & " " & Cur() & " eof=" & pRecordset.EOF & " bof=" & pRecordset.BOF
End Sub

Private Sub Adodc1_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    WriteLog "  willchangerecord " & adReason & " " & Cur()
End Sub

Private Sub Adodc1_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    WriteLog "  recordchangecomplete " & adReason & " " & Cur()
End Sub

Private Sub Adodc1_EndOfRecordset(fMoreData As Boolean, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    WriteLog "  endofrecordset " & Cur()
End Sub

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    WriteLog "  adodc error " & ErrorNumber & " " & Description & " scode=" & Scode & " src=" & Source
    fCancelDisplay = True
End Sub

Private Sub DataList1_Click()
    WriteLog "  list click [" & DataList1.Text & "] " & DataList1.BoundText
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep & " " & Cur()
    On Error Resume Next
    Select Case mStep
    Case 1
        WriteLog "rs " & Adodc1.Recordset.CursorType & "/" & Adodc1.Recordset.LockType & "/" & Adodc1.Recordset.CursorLocation & " pos=" & Adodc1.Recordset.AbsolutePosition
        WriteLog "changed " & Text1.DataChanged & " field " & Text1.DataField & " source " & TypeName(Text1.DataSource)
        Adodc1.Recordset.MoveNext
        WriteLog "after movenext " & Cur() & " " & Grid()
    Case 2
        Text1.Text = "Billy"
        WriteLog "edited changed=" & Text1.DataChanged & " field=" & Adodc1.Recordset!Name & " editmode=" & Adodc1.Recordset.EditMode
        Adodc1.Recordset.MoveNext
        WriteLog "after movenext " & Cur()
        Adodc1.Recordset.MovePrevious
        WriteLog "back " & Cur() & " field=" & Adodc1.Recordset!Name
    Case 3
        Adodc1.Recordset.MoveLast
        Adodc1.Recordset.MoveNext
        WriteLog "eof " & Adodc1.Recordset.EOF & " " & Cur()
        Adodc1.Recordset.MoveFirst
        WriteLog "first " & Cur() & " " & Grid()
    Case 4
        Adodc1.Recordset.AddNew
        WriteLog "addnew " & Cur() & " editmode=" & Adodc1.Recordset.EditMode
        Text1.Text = "Dora"
        Text2.Text = "19"
        Adodc1.Recordset.Update
        WriteLog "update " & Cur() & " count=" & Adodc1.Recordset.RecordCount & " id=" & Adodc1.Recordset!ID
    Case 5
        DataList1.BoundText = "2"
        WriteLog "list bound 2 text=[" & DataList1.Text & "] matched=" & DataList1.MatchedWithList
        DataCombo1.BoundText = "3"
        WriteLog "combo bound 3 text=[" & DataCombo1.Text & "] matched=" & DataCombo1.MatchedWithList
        DataCombo1.Text = "Patras"
        WriteLog "combo text Patras bound=" & DataCombo1.BoundText & " matched=" & DataCombo1.MatchedWithList
        DataCombo1.Text = "Sparta"
        WriteLog "combo text Sparta bound=[" & DataCombo1.BoundText & "] matched=" & DataCombo1.MatchedWithList
    Case 6
        Adodc1.RecordSource = "SELECT * FROM People WHERE Age > 20 ORDER BY Age DESC"
        Adodc1.Refresh
        WriteLog "refreshed " & Cur() & " count=" & Adodc1.Recordset.RecordCount & " " & Grid()
    Case 7
        Adodc1.Recordset.Delete
        WriteLog "deleted " & Cur()
        Adodc1.Recordset.MoveNext
        WriteLog "next " & Cur() & " " & Grid()
    Case 8
        Text1.Text = ""
        Adodc1.Recordset.MoveFirst
        WriteLog "empty name " & Cur()
        If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    Case 9
        Adodc1.RecordSource = "SELECT * FROM Nope"
        Adodc1.Refresh
        WriteLog "missing table " & Err.Number
        Err.Clear
        WriteLog "after " & Cur()
    Case 10
        Unload Me
        Exit Sub
    End Select
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLog "form unload " & Cur()
End Sub
