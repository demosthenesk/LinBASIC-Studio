VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Begin VB.Form Form1
   Caption         =   "RDO 2"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4200
      Top             =   1800
   End
   Begin MSRDC.MSRDC MSRDC1
      Height          =   330
      Left            =   120
      Top             =   1560
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      _Version        =   393216
   End
   Begin VB.TextBox Text1
      DataField       =   "Name"
      DataSource      =   "MSRDC1"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text2
      DataField       =   "Age"
      DataSource      =   "MSRDC1"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   2415
   End
   Begin VB.CheckBox Check1
      Caption         =   "Active"
      DataField       =   "Active"
      DataSource      =   "MSRDC1"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' RDO 2 (LinBASIC Studio phase 7.8): the RemoteData control (MSRDC20.OCX) and the controls bound to it: when it opens,
' its properties, moving, changing a bound control, UpdateRow / UpdateControls, EOFAction, Validate / Reposition / Error

Private mStep As Integer

Private Function St() As String
    Dim s As String
    On Error Resume Next
    s = "text1=[" & Text1.Text & "] text2=[" & Text2.Text & "] check1=" & Check1.Value
    s = s & " rs=" & TypeName(MSRDC1.Resultset)
    s = s & " rows=" & MSRDC1.Resultset.RowCount & " abs=" & MSRDC1.Resultset.AbsolutePosition & " edit=" & MSRDC1.EditMode & _
        " bof=" & MSRDC1.Resultset.BOF & " eof=" & MSRDC1.Resultset.EOF & " type=" & MSRDC1.Resultset.Type & _
        " lock=" & MSRDC1.Resultset.LockType
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    St = s
End Function

Private Sub Log(ByVal what As String)
    WriteLog what & ": " & St()
End Sub

Private Sub ErrLog(ByVal what As String)
    If Err.Number <> 0 Then WriteLog "  " & what & " ! " & Err.Number & " " & Err.Description: Err.Clear
End Sub

Private Sub MSRDC1_Validate(Action As Integer, Reserved As Integer)
    WriteLog "  ev validate action=" & Action & " reserved=" & Reserved & " text1=[" & Text1.Text & "] changed=" & Text1.DataChanged
End Sub

Private Sub MSRDC1_Reposition()
    WriteLog "  ev reposition text1=[" & Text1.Text & "]"
End Sub

Private Sub MSRDC1_Error(ByVal Number As Long, Description As String, ByVal Scode As Long, ByVal Source As String, _
        ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLog "  ev error " & Number & " [" & Source & "] " & Description & " scode=" & Scode & " cancel=" & CancelDisplay
    CancelDisplay = True
End Sub

Private Sub MSRDC1_QueryCompleted()
    WriteLog "  ev querycompleted"
End Sub

Private Sub Form_Load()
    Dim s As String
    On Error Resume Next
    With MSRDC1
        s = "props dsn=[" & .DataSourceName & "] sql=[" & .SQL & "] connect=[" & .Connect & "] prompt=" & .Prompt & _
            " type=" & .ResultsetType & " lock=" & .LockType & " cursor=" & .CursorDriver & " keyset=" & .KeysetSize & _
            " rowset=" & .RowsetSize & " maxrows=" & .MaxRows & " options=" & .Options & " login=" & .LoginTimeout & _
            " qtime=" & .QueryTimeout & " ro=" & .ReadOnly & " eof=" & .EOFAction & " bof=" & .BOFAction & " enabled=" & .Enabled & _
            " caption=[" & .Caption & "] app=" & .Appearance & " back=" & Hex$(.BackColor) & " fore=" & Hex$(.ForeColor) & _
            " batch=" & .BatchSize & " crit=" & .UpdateCriteria & " op=" & .UpdateOperation & " errthr=" & .ErrorThreshold & _
            " log=[" & .LogMessages & "] user=[" & .UserName & "] pwd=[" & .Password & "] ver=" & .Version & _
            " env=" & TypeName(.Environment) & " cn=" & TypeName(.Connection) & " trans=" & .Transactions
    End With
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog s
    Log "load before"
    MSRDC1.Connect = CS
    MSRDC1.SQL = "SELECT * FROM People ORDER BY ID"
    MSRDC1.Prompt = rdDriverNoPrompt
    MSRDC1.Refresh
    ErrLog "refresh"
    Log "load after refresh"
    WriteLog "  cn=" & TypeName(MSRDC1.Connection) & " env=" & TypeName(MSRDC1.Environment)
    ErrLog "objects"
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    Static done As Boolean
    If done Then Exit Sub
    done = True
    Log "activate"
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    On Error Resume Next
    Select Case mStep
    Case 1
        Log "timer"
        MSRDC1.Resultset.MoveNext
        Log "movenext"
        MSRDC1.Resultset.MoveLast
        Log "movelast"
        MSRDC1.Resultset.MoveNext
        Log "past the end"
        MSRDC1.Resultset.MoveFirst
        Log "movefirst"
    Case 2
        Text1.Text = "Anne"
        Text2.Text = "32"
        Log "changed"
        MSRDC1.Resultset.MoveNext
        ErrLog "move after change"
        Log "moved"
        WriteLog "  db:" & DbRows()
        MSRDC1.Resultset.MovePrevious
        Text1.Text = "Zed"
        MSRDC1.UpdateControls
        Log "updatecontrols"
        Text1.Text = "Anna"
        MSRDC1.UpdateRow
        ErrLog "updaterow"
        Log "updaterow"
        WriteLog "  db:" & DbRows()
    Case 3
        Text2.Text = "abc"
        MSRDC1.Resultset.MoveNext
        ErrLog "type mismatch"
        Log "after bad age"
        MSRDC1.UpdateControls
        ErrLog "updatecontrols"
        Log "updatecontrols"
        MSRDC1.Refresh
        ErrLog "refresh"
        Log "refresh"
    Case 4
        MSRDC1.EOFAction = rdAddNew
        MSRDC1.Resultset.MoveLast
        MSRDC1.Resultset.MoveNext
        ErrLog "eof addnew"
        Log "eofaction addnew"
        Text1.Text = "Dino"
        Text2.Text = "37"
        MSRDC1.UpdateRow
        ErrLog "updaterow new"
        Log "after new"
        WriteLog "  db:" & DbRows()
        MSRDC1.EOFAction = rdMoveLast
        MSRDC1.Refresh
        ErrLog "refresh"
        Log "refresh"
    Case 5
        MSRDC1.ReadOnly = True
        MSRDC1.Refresh
        ErrLog "readonly refresh"
        Log "readonly"
        Text1.Text = "Nope"
        MSRDC1.Resultset.MoveNext
        ErrLog "readonly change"
        Log "readonly moved"
        WriteLog "  db:" & DbRows()
        MSRDC1.UpdateControls
        MSRDC1.ReadOnly = False
        MSRDC1.Refresh
        ErrLog "refresh"
        Log "refresh"
    Case 6
        MSRDC1.SQL = "SELECT * FROM People WHERE Age > 30 ORDER BY ID"
        MSRDC1.Refresh
        ErrLog "new sql"
        Log "new sql"
        MSRDC1.SQL = "SELECT * FROM Nowhere"
        MSRDC1.Refresh
        ErrLog "bad sql"
        Log "bad sql"
        MSRDC1.SQL = "SELECT * FROM People ORDER BY ID"
        MSRDC1.Refresh
        ErrLog "back"
        Log "back"
    Case 7
        MSRDC1.Resultset.MoveNext
        MSRDC1.Resultset.Delete
        ErrLog "delete"
        Log "deleted"
        MSRDC1.Resultset.MoveNext
        Log "after delete next"
        WriteLog "  db:" & DbRows()
        MSRDC1.Caption = "People"
        MSRDC1.Enabled = False
        Log "disabled"
    Case 8
        WriteLog "end"
        End
    End Select
    ErrLog "step " & mStep
    Timer1.Enabled = True
End Sub
