VERSION 5.00
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Begin VB.Form Form1
   Caption         =   "DataRepeater 1"
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
   Begin MSDataRepeaterLib.DataRepeater DataRepeater1
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4789
      _StreamID       =   -1412567295
      _Version        =   393216
      Caption         =   "People"
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10}
         _StreamID       =   -1412567295
         _Version        =   65536
         Name            =   "RepCtl.RowCtl"
      EndProperty
      RepeaterBindings=   2
      BeginProperty RepeaterBinding(0) {7D21A594-FC9B-11D0-A320-00AA00688B10}
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "PName"
         DataField       =   "Name"
         Key             =   "kName"
      EndProperty
      BeginProperty RepeaterBinding(1) {7D21A594-FC9B-11D0-A320-00AA00688B10}
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "PAge"
         DataField       =   "Age"
         Key             =   "kAge"
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' DataRepeater 1 (LinBASIC Studio phase 7.5): RepCtl.RowCtl repeated for the records of an ADO Recordset; the
' RepeaterBindings of the .frm, its properties and collections, the events, moves of the recordset and of the control,
' a change made through the RepeatedControl, Add / Remove of bindings and their errors

Private cn As ADODB.Connection
Private rs As ADODB.Recordset
Private mStep As Integer

Private Sub Chk(ByVal what As String)
    If Err.Number <> 0 Then WriteLog "  " & what & " ! " & Err.Number & " [" & Err.Source & "] " & Err.Description: Err.Clear
End Sub

Private Sub Info(ByVal what As String)
    Dim s As String
    On Error Resume Next
    With DataRepeater1
        s = what & ": rows=" & .VisibleRows
        s = s & " active=" & .ActiveRow
        s = s & " count=" & .RecordCount
        s = s & " cur=" & RowName(rs, .CurrentRecord)
        s = s & " rs=" & rs!Name
        s = s & " ctl=" & .RepeatedControl.PName & "/" & .RepeatedControl.PAge & "/" & .RepeatedControl.PActive
    End With
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog s
End Sub

Private Sub Bindings(ByVal what As String)
    Dim b As Object
    On Error Resume Next
    WriteLog what & ": bindings=" & DataRepeater1.RepeaterBindings.Count
    For Each b In DataRepeater1.RepeaterBindings
        WriteLog "  binding " & b.Index & " " & b.PropertyName & "=" & b.DataField & " key=[" & b.Key & "] changed=" & _
            b.DataChanged & " fmt=" & TypeName(b.DataFormat) & " tag=" & TypeName(b.Tag)
    Next
    Chk what
End Sub

Private Sub DataRepeater1_ActiveRowChanged()
    WriteLog "  ev activerowchanged active=" & DataRepeater1.ActiveRow
End Sub

Private Sub DataRepeater1_VisibleRecordsChanged()
    WriteLog "  ev visiblerecordschanged"
End Sub

Private Sub DataRepeater1_CurrentRecordChanged()
    On Error Resume Next
    WriteLog "  ev currentrecordchanged cur=" & RowName(rs, DataRepeater1.CurrentRecord)
End Sub

Private Sub DataRepeater1_DataUpdate(Cancel As Boolean, ByVal Record As Variant, ByVal Field As String)
    WriteLog "  ev dataupdate cancel=" & Cancel & " record=" & RowName(rs, Record) & " field=[" & Field & "]"
End Sub

Private Sub DataRepeater1_RepeatedControlLoaded()
    On Error Resume Next
    WriteLog "  ev repeatedcontrolloaded ctl=" & TypeName(DataRepeater1.RepeatedControl)
End Sub

Private Sub DataRepeater1_RepeatedControlUnload()
    WriteLog "  ev repeatedcontrolunload"
End Sub

Private Sub Form_Load()
    Dim s As String
    On Error Resume Next
    With DataRepeater1
        s = "load: name=[" & .RepeatedControlName & "] caption=[" & .Caption & "] style=" & .CaptionStyle & _
            " rowind=" & .RowIndicator & " scroll=" & .ScrollBars & " divider=" & .RowDividerStyle & " integral=" & _
            .IntegralHeight & " border=" & .BorderStyle & " app=" & .Appearance & " back=" & Hex$(.BackColor) & _
            " fore=" & Hex$(.ForeColor) & " enabled=" & .Enabled & " member=[" & .DataMember & "]"
    End With
    Chk "props"
    WriteLog s
    Bindings "load"
    WriteLog "load props=" & Arr(DataRepeater1.PropertyNames)
    Chk "propertynames"
    WriteLog "load fields=" & Arr(DataRepeater1.DataFields)
    Chk "datafields"
    WriteLog "load ctl=" & TypeName(DataRepeater1.RepeatedControl)
    Chk "repeatedcontrol"
    Info "load before data"
    Set cn = New ADODB.Connection
    cn.Open CS
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT ID, Name, Age, Active, Notes FROM People ORDER BY ID", cn, adOpenStatic, adLockOptimistic
    Chk "open"
    Set DataRepeater1.DataSource = rs
    Chk "datasource"
    Info "load after data"
    WriteLog "load fields=" & Arr(DataRepeater1.DataFields) & " props=" & Arr(DataRepeater1.PropertyNames)
    Chk "arrays"
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    Static done As Boolean
    If done Then Exit Sub
    done = True
    Info "activate"
End Sub

Private Sub Timer1_Timer()
    Dim i As Long, c As ADODB.Recordset, b As Object
    Timer1.Enabled = False
    mStep = mStep + 1
    On Error Resume Next
    Select Case mStep
    Case 1
        Info "timer"
        rs.MoveNext
        Info "movenext"
        rs.MoveLast
        Info "movelast"
        rs.MoveFirst
        Info "movefirst"
    Case 2
        For i = 1 To DataRepeater1.VisibleRows
            WriteLog "  visible " & i & " = " & RowName(rs, DataRepeater1.VisibleRecords(i))
        Next
        Chk "visible"
        WriteLog "  visible 0 = " & RowName(rs, DataRepeater1.VisibleRecords(0))
        Chk "visible 0"
        WriteLog "  visible past = " & RowName(rs, DataRepeater1.VisibleRecords(DataRepeater1.VisibleRows + 1))
        Chk "visible past"
        DataRepeater1.ActiveRow = 3
        Chk "activerow 3"
        Info "activerow 3"
        Set c = rs.Clone
        c.Move 5
        DataRepeater1.CurrentRecord = c.Bookmark
        Chk "currentrecord"
        Info "currentrecord 6th"
        DataRepeater1.CurrentRecord = "x"
        Chk "currentrecord x"
        DataRepeater1.ActiveRow = 99
        Chk "activerow 99"
        rs.MoveFirst
        Info "movefirst"
    Case 3
        DataRepeater1.RepeatedControl.PName = "Zed"
        Chk "set pname"
        Info "set pname"
        WriteLog "  rs name=" & rs!Name & " editmode=" & rs.EditMode
        rs.MoveNext
        Chk "move after change"
        Info "moved"
        WriteLog "  db:" & DbRows()
        rs.MovePrevious
        Info "back"
    Case 4
        Set b = DataRepeater1.RepeaterBindings.Add("PActive", "Active", , "kAct")
        Chk "add pactive"
        WriteLog "  added " & TypeName(b) & " index=" & b.Index
        Bindings "after add"
        Info "after add"
        DataRepeater1.RepeaterBindings.Add "PName", "Notes"
        Chk "add pname again"
        DataRepeater1.RepeaterBindings.Add "Nope", "Name"
        Chk "add nope"
        DataRepeater1.RepeaterBindings.Add "Info", "Notes"
        Chk "add info"
        DataRepeater1.RepeaterBindings.Add "PActive", "NoField"
        Chk "add nofield"
        WriteLog "  item kAge=" & DataRepeater1.RepeaterBindings("kAge").DataField & " item 0=" & _
            DataRepeater1.RepeaterBindings(0).PropertyName
        Chk "item"
        WriteLog "  item 9=" & DataRepeater1.RepeaterBindings(9).PropertyName
        Chk "item 9"
        WriteLog "  item zz=" & DataRepeater1.RepeaterBindings("zz").PropertyName
        Chk "item zz"
        DataRepeater1.RepeaterBindings.Remove "kAge"
        Chk "remove"
        Bindings "after remove"
        rs.MoveNext
        Info "after remove moved"
    Case 5
        Set DataRepeater1.DataSource = Nothing
        Chk "datasource nothing"
        Info "no data"
        WriteLog "end"
        End
    End Select
    Chk "step " & mStep
    Timer1.Enabled = True
End Sub
