VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Form1
   Caption         =   "DBGrid DBList DBCombo"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1
      Bindings        =   "Form1.frx":0000
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin MSDBCtls.DBList DBList1
      Bindings        =   "Form1.frx":09E7
      DataField       =   "GroupID"
      DataSource      =   "Data1"
      Height          =   1035
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1826
      _Version        =   393216
      ListField       =   "GName"
      BoundColumn     =   "GID"
   End
   Begin MSDBCtls.DBCombo DBCombo1
      Bindings        =   "Form1.frx":09FB
      DataField       =   "GroupID"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "GName"
      BoundColumn     =   "GID"
      Text            =   "DBCombo1"
   End
   Begin VB.TextBox Text1
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Timer Timer1
      Interval        =   500
      Left            =   6600
      Top             =   4200
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
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Data Data2
      Caption         =   "Groups"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' DBGrid (DBGRID32.OCX) bound to a Data control, and DBList / DBCombo (DBLIST32.OCX) with a RowSource and a
' DataSource (Data1 People, Data2 Groups): their properties, how they follow the Data control and move it,
' the columns of the grid, the errors, the events.
Private gEv As String

Private Sub L(ByVal s As String)
    WriteLog s
End Sub

Private Function Ev() As String
    If gEv <> "" Then Ev = " ev[" & Trim$(gEv) & "]"
    gEv = ""
End Function

' a value as text: Null, Empty, an array (a bookmark), or its type and value
Private Function V(x As Variant) As String
    If IsNull(x) Then
        V = "Null"
    ElseIf IsEmpty(x) Then
        V = "Empty"
    ElseIf IsArray(x) Then
        V = TypeName(x)
    ElseIf IsObject(x) Then
        V = "obj " & TypeName(x)
    Else
        V = TypeName(x) & " '" & x & "'"
    End If
End Function

Private Function BmEq(a As Variant, b As Variant) As Boolean
    Dim i As Long
    If IsNull(a) Or IsNull(b) Or IsEmpty(a) Or IsEmpty(b) Then Exit Function
    If IsArray(a) And IsArray(b) Then
        If UBound(a) - LBound(a) <> UBound(b) - LBound(b) Then Exit Function
        For i = 0 To UBound(a) - LBound(a)
            If a(LBound(a) + i) <> b(LBound(b) + i) Then Exit Function
        Next
        BmEq = True
    Else
        BmEq = (a = b)
    End If
End Function

' the record of Data1 and the state of the grid
Private Function Cur() As String
    Dim s As String
    On Error Resume Next
    s = "rec=" & Text1.Text & " pos=" & Data1.Recordset.AbsolutePosition
    s = s & " | grid Row=" & DBGrid1.Row
    If Err Then s = s & "!" & Err.Number: Err.Clear
    s = s & " Col=" & DBGrid1.Col
    If Err Then s = s & "!" & Err.Number: Err.Clear
    s = s & " Text=" & V(DBGrid1.Text)
    If Err Then s = s & "!" & Err.Number: Err.Clear
    s = s & " bm=rs:" & BmEq(DBGrid1.Bookmark, Data1.Recordset.Bookmark)
    If Err Then s = s & "!" & Err.Number: Err.Clear
    Cur = s
End Function

Private Function Lst() As String
    Dim s As String
    On Error Resume Next
    s = "list Text=" & V(DBList1.Text) & " Bound=" & V(DBList1.BoundText) & " Matched=" & DBList1.MatchedWithList & _
        " Sel=" & V(DBList1.SelectedItem) & " Changed=" & DBList1.DataChanged
    If Err Then s = s & " !" & Err.Number & " " & Err.Description: Err.Clear
    s = s & " | combo Text=" & V(DBCombo1.Text) & " Bound=" & V(DBCombo1.BoundText) & " Matched=" & DBCombo1.MatchedWithList & _
        " Sel=" & V(DBCombo1.SelectedItem) & " Changed=" & DBCombo1.DataChanged
    If Err Then s = s & " !" & Err.Number & " " & Err.Description: Err.Clear
    Lst = s
End Function

Private Sub Cols()
    Dim i As Integer, c As Object, s As String
    On Error Resume Next
    s = "  Columns.Count=" & DBGrid1.Columns.Count
    If Err Then s = s & "!" & Err.Number: Err.Clear
    L s
    For i = 0 To DBGrid1.Columns.Count - 1
        Set c = DBGrid1.Columns(i)
        s = "    col " & i & ": " & TypeName(c) & " Caption=" & V(c.Caption) & " DataField=" & V(c.DataField) & _
            " ColIndex=" & c.ColIndex & " Visible=" & c.Visible & " Locked=" & c.Locked & " Alignment=" & c.Alignment & _
            " Value=" & V(c.Value) & " Text=" & V(c.Text)
        If Err Then s = s & " !" & Err.Number & " " & Err.Description: Err.Clear
        L s
    Next
End Sub

Private Sub Try(ByVal what As String)
    Dim s As String, x As Variant
    gEv = ""
    On Error Resume Next
    Select Case what
    Case "MoveNext": Data1.Recordset.MoveNext
    Case "MoveLast": Data1.Recordset.MoveLast
    Case "MoveFirst": Data1.Recordset.MoveFirst
    Case "Row=2": DBGrid1.Row = 2
    Case "Row=0": DBGrid1.Row = 0
    Case "Row=9": DBGrid1.Row = 9
    Case "Col=2": DBGrid1.Col = 2
    Case "Col=7": DBGrid1.Col = 7
    Case "Refresh": DBGrid1.Refresh
    Case "ReBind": DBGrid1.ReBind
    Case "ClearFields": DBGrid1.ClearFields
    Case "HoldFields": DBGrid1.HoldFields
    Case "Text=X": DBGrid1.Text = "X"
    Case "Data1.Refresh": Data1.Refresh
    End Select
    If Err Then s = " ! " & Err.Number & " " & Err.Description: Err.Clear
    DoEvents
    L "  " & what & s & " -> " & Cur() & Ev()
End Sub

Private Sub RunAll()
    Dim bm As Variant, s As String, i As Integer, x As Variant
    On Error Resume Next

    L "== the grid"
    With DBGrid1
        L "  TypeName " & TypeName(DBGrid1) & " Caption=" & V(.Caption) & " AllowAddNew=" & .AllowAddNew & " AllowDelete=" & _
          .AllowDelete & " AllowUpdate=" & .AllowUpdate & " ColumnHeaders=" & .ColumnHeaders & " RecordSelectors=" & _
          .RecordSelectors & " HeadLines=" & .HeadLines & " DataMode=" & .DataMode & " MarqueeStyle=" & .MarqueeStyle
        If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
        L "  ScrollBars=" & .ScrollBars & " Enabled=" & .Enabled & " ApproxCount=" & .ApproxCount & " LeftCol=" & .LeftCol & _
          " FirstRow=" & TypeName(.FirstRow) & " VisibleRows>=4=" & (.VisibleRows >= 4) & " VisibleCols>0=" & (.VisibleCols > 0) & _
          " DataChanged=" & .DataChanged & " EditActive=" & .EditActive & " AddNewMode=" & .AddNewMode
        If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
        L "  Splits.Count=" & .Splits.Count & " SelBookmarks.Count=" & .SelBookmarks.Count & " default=" & V(DBGrid1)
        If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    End With
    L "  " & Cur() & Ev()
    Cols
    s = "  Columns(""Name"").Value=" & V(DBGrid1.Columns("Name").Value)
    If Err Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    L s
    s = "  Columns(9): "
    Set x = DBGrid1.Columns(9)
    If Err Then s = s & "! " & Err.Number & " " & Err.Description: Err.Clear
    L s
    s = "  Columns(""Nope""): "
    Set x = DBGrid1.Columns("Nope")
    If Err Then s = s & "! " & Err.Number & " " & Err.Description: Err.Clear
    L s

    L "== moving"
    Try "MoveNext"
    Try "MoveNext"
    Try "Row=0"
    Try "Row=2"
    Try "Row=9"
    Try "Col=2"
    Try "Col=7"
    Try "MoveLast"
    bm = Data1.Recordset.Bookmark
    Try "MoveFirst"
    gEv = ""
    DBGrid1.Bookmark = bm
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    DoEvents
    L "  Bookmark = last -> " & Cur() & Ev()
    s = "  CellText(last)=" & V(DBGrid1.Columns(1).CellText(bm)) & " CellValue(last)=" & V(DBGrid1.Columns(2).CellValue(bm))
    If Err Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    L s
    Data1.Recordset.MoveFirst
    DoEvents
    s = "  GetBookmark(1)=2nd:"
    Data1.Recordset.MoveNext
    x = Data1.Recordset.Bookmark
    Data1.Recordset.MovePrevious
    DoEvents
    s = s & BmEq(DBGrid1.GetBookmark(1), x) & " RowBookmark(1)=2nd:" & BmEq(DBGrid1.RowBookmark(1), x)
    If Err Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    L s
    gEv = ""
    L "  " & Cur()

    L "== changes in the recordset"
    Data1.Recordset.Edit
    Data1.Recordset!Name = "Anne"
    Data1.Recordset.Update
    DoEvents
    L "  Edit Anne: Columns(1).Text=" & V(DBGrid1.Columns(1).Text) & " " & Cur() & Ev()
    Try "Refresh"
    Data1.Recordset.AddNew
    Data1.Recordset!Name = "Emil"
    Data1.Recordset!Age = 50
    Data1.Recordset!GroupID = 1
    Data1.Recordset.Update
    DoEvents
    L "  AddNew Emil: ApproxCount=" & DBGrid1.ApproxCount & " " & Cur() & Ev()
    Try "Refresh"
    L "  ApproxCount=" & DBGrid1.ApproxCount
    Try "Data1.Refresh"
    L "  ApproxCount=" & DBGrid1.ApproxCount
    Try "Text=X"
    L "  Text after: " & V(DBGrid1.Text) & " rec=" & Text1.Text & " EditActive=" & DBGrid1.EditActive & " DataChanged=" & DBGrid1.DataChanged

    L "== the columns"
    DBGrid1.Columns(0).Caption = "Number"
    DBGrid1.Columns(2).Visible = False
    DBGrid1.Columns(1).Width = 2000
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    L "  Caption Number, col 2 hidden: VisibleCols>0=" & (DBGrid1.VisibleCols > 0) & " col1 Width=" & DBGrid1.Columns(1).Width
    Cols
    Try "ReBind"
    Cols
    Try "HoldFields"
    Try "ReBind"
    Cols
    Try "ClearFields"
    Cols
    Try "ReBind"
    Cols
    DBGrid1.Caption = "People"
    L "  Caption=" & V(DBGrid1.Caption)
    s = "  Columns.Add(1): "
    Set x = DBGrid1.Columns.Add(1)
    If Err Then s = s & "! " & Err.Number & " " & Err.Description: Err.Clear Else s = s & TypeName(x) & " count=" & DBGrid1.Columns.Count & " DataField=" & V(x.DataField)
    L s
    s = "  Columns.Remove 1: "
    DBGrid1.Columns.Remove 1
    If Err Then s = s & "! " & Err.Number & " " & Err.Description: Err.Clear Else s = s & "count=" & DBGrid1.Columns.Count
    L s

    L "== DBList, DBCombo"
    Data1.Recordset.MoveFirst
    DoEvents
    gEv = ""
    L "  ListField=" & V(DBList1.ListField) & " BoundColumn=" & V(DBList1.BoundColumn) & " DataField=" & V(DBList1.DataField) & _
      " Style=" & DBCombo1.Style & " MatchEntry=" & DBList1.MatchEntry & " Locked=" & DBList1.Locked & " TypeName " & _
      TypeName(DBList1) & " " & TypeName(DBCombo1)
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    L "  VisibleCount>0=" & (DBList1.VisibleCount > 0) & " VisibleItems(0)=" & V(DBList1.VisibleItems(0)) & " default=" & V(DBList1)
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    L "  first: " & Lst() & Ev()
    Data1.Recordset.MoveNext
    DoEvents
    L "  MoveNext: " & Lst() & Ev()
    DBList1.BoundText = "3"
    DoEvents
    L "  list BoundText = 3: " & Lst() & Ev()
    Data1.Recordset.MoveNext
    DoEvents
    Data1.Recordset.MovePrevious
    DoEvents
    L "  moved away and back: GroupID=" & Data1.Recordset!GroupID & " " & Lst() & Ev()
    DBList1.Text = "Alpha"
    DoEvents
    L "  list Text = Alpha: " & Lst() & Ev()
    DBList1.Text = "Zeta"
    DoEvents
    L "  list Text = Zeta: " & Lst() & Ev()
    DBCombo1.Text = "Gam"
    DoEvents
    L "  combo Text = Gam: " & Lst() & Ev()
    DBCombo1.BoundText = "1"
    DoEvents
    L "  combo BoundText = 1: " & Lst() & Ev()
    DBCombo1.SelStart = 1
    DBCombo1.SelLength = 2
    L "  combo Sel: " & DBCombo1.SelStart & "," & DBCombo1.SelLength & " " & V(DBCombo1.SelText)
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    Data1.UpdateControls
    DoEvents
    L "  UpdateControls: " & Lst() & Ev()
    Data2.Recordset.AddNew
    Data2.Recordset!GName = "Delta"
    Data2.Recordset.Update
    DoEvents
    L "  Groups AddNew Delta: list " & V(DBList1.VisibleItems(3)) & Ev()
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    DBList1.ReFill
    DBCombo1.ReFill
    DoEvents
    L "  ReFill: " & V(DBList1.VisibleItems(3)) & " " & Lst() & Ev()
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    L "end"
End Sub

Private Sub Data1_Reposition()
    gEv = gEv & " Rep"
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
    gEv = gEv & " Val(" & Action & "," & Save & ")"
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    gEv = gEv & " RCC(" & IIf(IsNull(LastRow), "Null", "bm") & "," & LastCol & ")"
End Sub

Private Sub DBGrid1_Click()
    gEv = gEv & " GClick"
End Sub

Private Sub DBGrid1_BeforeUpdate(Cancel As Integer)
    gEv = gEv & " BU"
End Sub

Private Sub DBGrid1_AfterUpdate()
    gEv = gEv & " AU"
End Sub

Private Sub DBGrid1_Change()
    gEv = gEv & " GChange"
End Sub

Private Sub DBGrid1_Error(ByVal DataError As Integer, Response As Integer)
    gEv = gEv & " GErr(" & DataError & ")"
    Response = 0
End Sub

Private Sub DBList1_Click()
    gEv = gEv & " LClick"
End Sub

Private Sub DBCombo1_Change()
    gEv = gEv & " CChange"
End Sub

Private Sub Form_Load()
    Data1.DatabaseName = App.Path & "\dbgrid.mdb"
    Data1.RecordSource = "SELECT * FROM People ORDER BY ID"
    Data2.DatabaseName = App.Path & "\dbgrid.mdb"
    Data2.RecordSource = "SELECT * FROM Groups ORDER BY GID"
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    gEv = ""
    RunAll
    Unload Me
End Sub
