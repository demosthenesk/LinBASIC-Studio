VERSION 5.00
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Begin VB.Form Form1
   Caption         =   "DataRepeater 2"
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
      CaptionStyle    =   0
      Caption         =   "DataRepeater1"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' DataRepeater 2 (LinBASIC Studio phase 7.5): RepeatedControlName and RepeaterBindings set in code, a ProgID that is
' not there, changes of the recordset (Update, AddNew, Requery, Delete, Filter) seen by the DataRepeater, Refresh,
' ActiveRow and VisibleRecords set in the middle, at the top and at the bottom of the records, Item by index, an invalid
' DataField

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

Private Sub VisList(ByVal what As String)
    Dim i As Long, s As String
    On Error Resume Next
    s = "  " & what & " visible:"
    For i = 1 To DataRepeater1.VisibleRows
        s = s & " " & RowName(rs, DataRepeater1.VisibleRecords(i))
    Next
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog s
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
    ' a record being added (AddNew): its bookmark is one of the Cursor Engine's own
    If rs.EditMode = adEditAdd Then
        WriteLog "  ev dataupdate cancel=" & Cancel & " record=(new) field=[" & Field & "]"
    Else
        WriteLog "  ev dataupdate cancel=" & Cancel & " record=" & RowName(rs, Record) & " field=[" & Field & "]"
    End If
End Sub

Private Sub DataRepeater1_RepeatedControlLoaded()
    On Error Resume Next
    WriteLog "  ev repeatedcontrolloaded ctl=" & TypeName(DataRepeater1.RepeatedControl)
End Sub

Private Sub DataRepeater1_RepeatedControlUnload()
    WriteLog "  ev repeatedcontrolunload"
End Sub

Private Sub Form_Load()
    Dim b As Object
    On Error Resume Next
    WriteLog "load: name=[" & DataRepeater1.RepeatedControlName & "] ctl=" & TypeName(DataRepeater1.RepeatedControl) & _
        " bindings=" & DataRepeater1.RepeaterBindings.Count & " props=" & Arr(DataRepeater1.PropertyNames)
    Chk "load"
    DataRepeater1.RepeatedControlName = "NoSuch.Thing"
    Chk "no such"
    WriteLog "after no such: name=[" & DataRepeater1.RepeatedControlName & "] ctl=" & TypeName(DataRepeater1.RepeatedControl)
    Chk "after no such"
    DataRepeater1.RepeatedControlName = "RepCtl.RowCtl"
    Chk "name"
    WriteLog "after name: name=[" & DataRepeater1.RepeatedControlName & "] ctl=" & TypeName(DataRepeater1.RepeatedControl) & _
        " props=" & Arr(DataRepeater1.PropertyNames)
    Chk "after name"
    Set b = DataRepeater1.RepeaterBindings.Add("PName", "Name")
    Chk "add pname"
    WriteLog "  binding key=[" & b.Key & "] index=" & b.Index
    DataRepeater1.RepeaterBindings.Add "PActive", "Active", , "act"
    Chk "add pactive"
    Info "before data"
    Set cn = New ADODB.Connection
    cn.Open CS
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT ID, Name, Age, Active FROM People ORDER BY ID", cn, adOpenStatic, adLockOptimistic
    Set DataRepeater1.DataSource = rs
    Chk "datasource"
    Info "after data"
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    Static done As Boolean
    If done Then Exit Sub
    done = True
    Info "activate"
End Sub

Private Sub Timer1_Timer()
    Dim c As ADODB.Recordset
    Timer1.Enabled = False
    mStep = mStep + 1
    On Error Resume Next
    Select Case mStep
    Case 1
        rs!Name = "Anne"
        rs.Update
        Chk "update"
        Info "updated"
        rs.MoveLast
        rs.AddNew
        rs!ID = 9
        rs!Name = "Ivan"
        rs!Age = 33
        rs!Active = True
        rs.Update
        Chk "addnew"
        Info "added"
        rs.MoveFirst
        Info "movefirst"
    Case 2
        rs.Requery
        Chk "requery"
        Info "requery"
        DataRepeater1.Refresh
        Chk "refresh"
        Info "refresh"
    Case 3
        WriteLog "  item 1=" & DataRepeater1.RepeaterBindings(1).PropertyName & " item act=" & _
            DataRepeater1.RepeaterBindings("act").DataField
        Chk "item"
        DataRepeater1.RepeaterBindings.Add "PAge", "NoField"
        Chk "add nofield"
        DataRepeater1.RepeaterBindings.Add "Info", "Name"
        Chk "add info"
        WriteLog "  bindings=" & DataRepeater1.RepeaterBindings.Count
        Set c = rs.Clone
        c.Move 4
        DataRepeater1.CurrentRecord = c.Bookmark
        Chk "cur 5th"
        Info "cur 5th"
        VisList "cur 5th"
        DataRepeater1.ActiveRow = DataRepeater1.VisibleRows
        Chk "activerow last"
        Info "activerow last"
        VisList "activerow last"
        DataRepeater1.ActiveRow = 1
        Chk "activerow 1"
        Info "activerow 1"
        VisList "activerow 1"
        rs.MoveFirst
        DataRepeater1.ActiveRow = 2
        Chk "activerow 2 at top"
        Info "activerow 2 at top"
        VisList "activerow 2 at top"
        rs.MoveLast
        DataRepeater1.ActiveRow = 1
        Chk "activerow 1 at bottom"
        Info "activerow 1 at bottom"
        VisList "activerow 1 at bottom"
        c.MoveFirst
        c.Move 2
        DataRepeater1.VisibleRecords(1) = c.Bookmark
        Chk "visiblerecords 1"
        Info "visiblerecords 1"
        VisList "visiblerecords 1"
        DataRepeater1.VisibleRecords(0) = c.Bookmark
        Chk "visiblerecords 0"
        rs.MoveFirst
        Info "movefirst"
        VisList "movefirst"
    Case 4
        ' the last step runs to its End in one go: no event comes after it
        DataRepeater1.RepeaterBindings.Clear
        Chk "clear"
        Info "no bindings"
        DataRepeater1.RepeaterBindings.Add "PAge", "Age"
        Chk "add page"
        Info "page bound"
        rs.MoveFirst
        rs.MoveNext
        rs.Delete
        Chk "delete"
        Info "deleted"
        rs.MoveNext
        Info "after delete next"
        WriteLog "  db:" & DbRows()
        rs.Filter = "Age > 30"
        Chk "filter"
        Info "filter"
        rs.Filter = adFilterNone
        Chk "no filter"
        Info "no filter"
        DataRepeater1.RepeatedControlName = ""
        Chk "name empty"
        Info "no control"
        WriteLog "end"
        End
    End Select
    Chk "step " & mStep
    Timer1.Enabled = True
End Sub
