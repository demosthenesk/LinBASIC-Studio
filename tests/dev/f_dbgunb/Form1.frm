VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Form1
   Caption         =   "DBGrid unbound, a .frx of LinBASIC Studio"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin MSDBGrid.DBGrid DBGrid2
      Bindings        =   "Form1.frx":09DB
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "Form1.frx":09EF
      TabIndex        =   1
      Top             =   2160
      Width           =   4935
   End
   Begin MSDBCtls.DBList DBList1
      Bindings        =   "Form1.frx":13C2
      Height          =   1035
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1826
      _Version        =   393216
      ListField       =   "Name"
      BoundColumn     =   "ID"
   End
   Begin MSDBCtls.DBCombo DBCombo1
      Bindings        =   "Form1.frx":13D6
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Name"
      BoundColumn     =   "ID"
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
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' DBGrid1 unbound (DataMode = 1): its rows come from arrays through the Unbound events. DBGrid2 bound to
' Data1 with columns made at design time. The .frx (the OleObjectBlobs, the Bindings of DBGrid2, DBList1,
' DBCombo1) was written by LinBASIC Studio (lbvc/project/oleblob.py): VB6 reads it back here.
Private Keys(1 To 20) As Long, Texts(1 To 20) As String, nRows As Long
Private gEv As String, nRead As Long

Private Sub L(ByVal s As String)
    WriteLog s
End Sub

Private Function Ev() As String
    If gEv <> "" Then Ev = " ev[" & Trim$(gEv) & "]"
    gEv = ""
End Function

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

Private Sub Props(g As Object, ByVal nm As String)
    Dim s As String, i As Integer
    On Error Resume Next
    s = "  " & nm & " Caption=" & V(g.Caption) & " DataMode=" & g.DataMode & " AllowAddNew=" & g.AllowAddNew & _
        " AllowDelete=" & g.AllowDelete & " AllowUpdate=" & g.AllowUpdate & " MarqueeStyle=" & g.MarqueeStyle & _
        " RecordSelectors=" & g.RecordSelectors & " HeadLines=" & g.HeadLines
    If Err Then s = s & " !" & Err.Number: Err.Clear
    L s
    s = "    Font=" & g.Font.Name & " " & g.Font.Size & " " & g.Font.Bold & " HeadFont=" & g.HeadFont.Name & " " & _
        g.HeadFont.Size & " " & g.HeadFont.Italic & " Width=" & g.Width & " Height=" & g.Height & " Columns.Count=" & g.Columns.Count
    If Err Then s = s & " !" & Err.Number: Err.Clear
    L s
    For i = 0 To g.Columns.Count - 1
        s = "    col " & i & " Caption=" & V(g.Columns(i).Caption) & " DataField=" & V(g.Columns(i).DataField) & _
            " Width=" & g.Columns(i).Width & " Alignment=" & g.Columns(i).Alignment
        If Err Then s = s & " !" & Err.Number: Err.Clear
        L s
    Next
End Sub

' the unbound grid: where it is, its current cell
Private Function Cur() As String
    Dim s As String
    On Error Resume Next
    s = "Row=" & DBGrid1.Row & " Col=" & DBGrid1.Col & " Text=" & V(DBGrid1.Text) & " Bookmark=" & V(DBGrid1.Bookmark)
    If Err Then s = s & " !" & Err.Number: Err.Clear
    s = s & " FirstRow=" & V(DBGrid1.FirstRow) & " DataChanged=" & DBGrid1.DataChanged
    If Err Then s = s & " !" & Err.Number: Err.Clear
    Cur = s & IIf(nRead > 0, " (read)", "")
    nRead = 0
End Function

Private Sub Try(ByVal what As String)
    Dim s As String
    gEv = ""
    nRead = 0
    On Error Resume Next
    Select Case what
    Case "Row=2": DBGrid1.Row = 2
    Case "Row=0": DBGrid1.Row = 0
    Case "Col=1": DBGrid1.Col = 1
    Case "Text=Zed": DBGrid1.Text = "Zed"
    Case "Refresh": DBGrid1.Refresh
    Case "ReBind": DBGrid1.ReBind
    Case "Scroll 0,2": DBGrid1.Scroll 0, 2
    Case "Bookmark=5": DBGrid1.Bookmark = 5
    Case "ApproxCount=6": DBGrid1.ApproxCount = 6
    Case "Columns(1).Value=Wye": DBGrid1.Columns(1).Value = "Wye"
    Case "Bookmark=1": DBGrid1.Bookmark = 1
    End Select
    If Err Then s = " ! " & Err.Number & " " & Err.Description: Err.Clear
    DoEvents
    L "  " & what & s & " -> " & Cur() & Ev()
End Sub

Private Sub RunAll()
    Dim s As String, i As Integer, x As Variant
    On Error Resume Next
    L "== the design-time properties (the .frx of LinBASIC Studio)"
    Props DBGrid1, "DBGrid1"
    Props DBGrid2, "DBGrid2"
    s = "  DBList1 RowSource="
    s = s & TypeName(DBList1.RowSource)
    If Err Then s = s & "! " & Err.Number & " " & Err.Description: Err.Clear
    s = s & " ListField=" & V(DBList1.ListField) & " VisibleItems(0)=" & V(DBList1.VisibleItems(0))
    If Err Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    s = s & " | DBCombo1 Style=" & DBCombo1.Style & " VisibleCount=" & DBCombo1.VisibleCount & " VisibleItems(0)="
    s = s & V(DBCombo1.VisibleItems(0))
    If Err Then s = s & "! " & Err.Number & " " & Err.Description: Err.Clear
    L s
    DBList1.Text = "Cleo"
    DBCombo1.BoundText = "4"
    L "  DBList1 Text=Cleo: " & V(DBList1.Text) & " " & V(DBList1.BoundText) & " " & DBList1.MatchedWithList & _
      " | DBCombo1 BoundText=4: " & V(DBCombo1.Text) & " " & DBCombo1.MatchedWithList
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear

    L "== DBGrid2 bound, its columns"
    L "  ApproxCount=" & DBGrid2.ApproxCount & " Text=" & V(DBGrid2.Text) & " Columns(1).Value=" & V(DBGrid2.Columns(1).Value) & _
      " Columns(""Person"").Text=" & V(DBGrid2.Columns("Person").Text)
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    Data1.Recordset.MoveNext
    DoEvents
    L "  MoveNext: Row=" & DBGrid2.Row & " Text=" & V(DBGrid2.Text) & " Columns(1).Value=" & V(DBGrid2.Columns(1).Value)
    DBGrid2.ReBind
    L "  ReBind: Columns.Count=" & DBGrid2.Columns.Count & " Caption(0)=" & V(DBGrid2.Columns(0).Caption) & _
      " DataField(1)=" & V(DBGrid2.Columns(1).DataField) & " Text=" & V(DBGrid2.Text)
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear

    L "== DBGrid1 unbound"
    gEv = ""
    L "  first: " & Cur() & " ApproxCount=" & DBGrid1.ApproxCount & " VisibleRows>=3=" & (DBGrid1.VisibleRows >= 3)
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    s = "  Columns(0).Value=" & V(DBGrid1.Columns(0).Value) & " Columns(1).Text=" & V(DBGrid1.Columns(1).Text)
    If Err Then s = s & " ! " & Err.Number: Err.Clear
    s = s & " CellText(3)=" & V(DBGrid1.Columns(1).CellText(3)) & " CellValue(3)=" & V(DBGrid1.Columns(0).CellValue(3))
    If Err Then s = s & " ! " & Err.Number: Err.Clear
    s = s & " GetBookmark(1)=" & V(DBGrid1.GetBookmark(1)) & " RowBookmark(2)=" & V(DBGrid1.RowBookmark(2))
    If Err Then s = s & " ! " & Err.Number: Err.Clear
    L s
    Try "Row=2"
    Try "Col=1"
    Try "Text=Zed"
    Try "Row=0"
    L "  Texts(3)=" & Texts(3)
    Try "Refresh"
    Try "Columns(1).Value=Wye"
    Try "Bookmark=5"
    L "  Texts(1)=" & Texts(1)
    Try "ReBind"
    Try "Scroll 0,2"
    Try "Bookmark=1"
    Try "ApproxCount=6"
    L "  ApproxCount=" & DBGrid1.ApproxCount & " AddNewMode=" & DBGrid1.AddNewMode & " Columns(1).Text=" & V(DBGrid1.Columns(1).Text)
    If Err Then L "  ! " & Err.Number & " " & Err.Description: Err.Clear
    L "end"
End Sub

Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim i As Long, r As Long, st As Long, n As Long
    nRead = nRead + 1
    If IsNull(StartLocation) Then
        If ReadPriorRows Then st = nRows + 1 Else st = 0
    Else
        st = StartLocation
    End If
    For i = 0 To RowBuf.RowCount - 1
        If ReadPriorRows Then r = st - 1 - i Else r = st + 1 + i
        If r < 1 Or r > nRows Then Exit For
        RowBuf.Value(i, 0) = Keys(r)
        RowBuf.Value(i, 1) = Texts(r)
        RowBuf.Bookmark(i) = r
        n = n + 1
    Next
    RowBuf.RowCount = n
End Sub

Private Sub DBGrid1_UnboundGetRelativeBookmark(StartLocation As Variant, ByVal OffSet As Long, NewLocation As Variant, ApproximatePosition As Long)
    Dim r As Long
    If IsNull(StartLocation) Then
        If OffSet > 0 Then r = 0 Else r = nRows + 1
    Else
        r = StartLocation
    End If
    r = r + OffSet
    If r >= 1 And r <= nRows Then
        NewLocation = r
        ApproximatePosition = r
    Else
        NewLocation = Null
    End If
End Sub

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    Dim c As Integer, s As String, r As Long
    r = WriteLocation
    s = " Write(" & V(WriteLocation)
    For c = 0 To RowBuf.ColumnCount - 1
        s = s & "|" & V(RowBuf.Value(0, c))
        If Not IsNull(RowBuf.Value(0, c)) Then
            If c = 0 Then Keys(r) = RowBuf.Value(0, c) Else Texts(r) = RowBuf.Value(0, c)
        End If
    Next
    gEv = gEv & s & ")"
End Sub

Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
    nRows = nRows + 1
    Keys(nRows) = nRows * 10
    If Not IsNull(RowBuf.Value(0, 1)) Then Texts(nRows) = RowBuf.Value(0, 1)
    NewRowBookmark = nRows
    gEv = gEv & " Add(" & nRows & ")"
End Sub

Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
    gEv = gEv & " Delete(" & V(Bookmark) & ")"
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    gEv = gEv & " RCC(" & V(LastRow) & "," & LastCol & ")"
End Sub

Private Sub DBGrid1_BeforeUpdate(Cancel As Integer)
    gEv = gEv & " BU"
End Sub

Private Sub DBGrid1_AfterUpdate()
    gEv = gEv & " AU"
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    gEv = gEv & " BCU(" & ColIndex & "," & V(OldValue) & ")"
End Sub

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    gEv = gEv & " ACU(" & ColIndex & ")"
End Sub

Private Sub DBGrid1_Error(ByVal DataError As Integer, Response As Integer)
    gEv = gEv & " Err(" & DataError & ")"
    Response = 0
End Sub

Private Sub Form_Load()
    Dim i As Long
    nRows = 6
    For i = 1 To nRows
        Keys(i) = i * 10
        Texts(i) = "R" & i
    Next
    Data1.DatabaseName = App.Path & "\dbgunb.mdb"
    Data1.RecordSource = "SELECT * FROM People ORDER BY ID"
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    gEv = ""
    RunAll
    Unload Me
End Sub
