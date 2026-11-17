VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Data control"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4200
      Top             =   3000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Active"
      DataField       =   "Active"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "Age"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "People"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label Label1 
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   2520
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
' the Data control (LinBASIC Studio, phase 4.4): its recordset, the bound controls, the order of
' Validate and Reposition, DataChanged and the saving of a changed record, UpdateControls, UpdateRecord,
' Refresh with another RecordSource, its errors

Private mStep As Integer

Private Function Cur() As String
    Cur = "[" & Text1.Text & "|" & Text2.Text & "|" & Check1.Value & "|" & Label1.Caption & "]"
End Function

Private Sub Form_Load()
    WriteLog "form load " & Data1.Caption & " type=" & Data1.RecordsetType & " eof=" & Data1.EOFAction & " bof=" & Data1.BOFAction & " ro=" & Data1.ReadOnly & " conn=" & Data1.Connect
    On Error Resume Next
    WriteLog "rs in load " & TypeName(Data1.Recordset)
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    On Error GoTo 0
    Data1.DatabaseName = App.Path & "\people.mdb"
    Data1.RecordSource = "People"
    WriteLog "set " & Cur()
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    WriteLog "form activate " & Cur()
End Sub

Private Sub Data1_Reposition()
    WriteLog "  reposition " & Cur() & " editmode=" & Data1.Recordset.EditMode
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
    WriteLog "  validate action=" & Action & " save=" & Save & " changed=" & Text1.DataChanged & "/" & Text2.DataChanged & "/" & Check1.DataChanged & " " & Cur()
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
    WriteLog "  error " & DataErr & " response=" & Response
    Response = 0
End Sub

Private Sub Timer1_Timer()
    Dim rs As Object
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep & " " & Cur()
    On Error Resume Next
    Select Case mStep
    Case 1
        Set rs = Data1.Recordset
        WriteLog "rs " & TypeName(rs) & " type=" & rs.Type & " count=" & rs.RecordCount & " name=" & rs.Name & " db=" & TypeName(Data1.Database) & " " & (Data1.Database.Name = Data1.DatabaseName)
        WriteLog "changed " & Text1.DataChanged & " field " & Text1.DataField & " source " & TypeName(Text1.DataSource)
        Data1.Recordset.MoveNext
        WriteLog "after movenext " & Cur()
    Case 2
        Text1.Text = "Billy"
        WriteLog "edited changed=" & Text1.DataChanged & " field=" & Data1.Recordset!Name
        Data1.Recordset.MoveNext
        WriteLog "after movenext " & Cur()
        Data1.Recordset.MovePrevious
        WriteLog "back " & Cur()
    Case 3
        Text2.Text = "99"
        Check1.Value = 1
        Data1.UpdateControls
        WriteLog "updatecontrols " & Cur() & " changed=" & Text2.DataChanged
        Text2.Text = "26"
        Data1.UpdateRecord
        WriteLog "updaterecord " & Cur() & " field=" & Data1.Recordset!Age & " changed=" & Text2.DataChanged
    Case 4
        Data1.Recordset.MoveLast
        Data1.Recordset.MoveNext
        WriteLog "eof " & Data1.Recordset.EOF & " " & Cur()
        Data1.Recordset.MoveFirst
    Case 5
        Data1.Recordset.AddNew
        WriteLog "addnew " & Cur() & " editmode=" & Data1.Recordset.EditMode
        Text1.Text = "Dora"
        Text2.Text = "19"
        Data1.Recordset.Update
        WriteLog "update " & Cur() & " count=" & Data1.Recordset.RecordCount
        If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
        Data1.Recordset.MoveLast
        WriteLog "last " & Cur()
    Case 6
        Data1.RecordSource = "SELECT * FROM People WHERE Age > 20 ORDER BY Age DESC"
        WriteLog "recordsource set " & Cur()
        Data1.Refresh
        WriteLog "refreshed " & Cur() & " count=" & Data1.Recordset.RecordCount
        Data1.Recordset.FindFirst "Name = 'Anna'"
        WriteLog "found " & Data1.Recordset.NoMatch & " " & Cur()
    Case 7
        Data1.Recordset.Delete
        WriteLog "deleted " & Cur()
        Data1.Recordset.MoveNext
        WriteLog "next " & Cur()
    Case 8
        Text1.Text = ""
        Data1.Recordset.MoveFirst
        WriteLog "empty name " & Cur()
        If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
    Case 9
        Data1.DatabaseName = App.Path & "\nofile.mdb"
        Data1.Refresh
        WriteLog "missing db " & Err.Number
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
