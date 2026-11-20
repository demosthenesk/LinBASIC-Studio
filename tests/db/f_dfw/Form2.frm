VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2
   Caption         =   "Pubs"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3600
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   5400
      Top             =   120
   End
   Begin MSDataGridLib.DataGrid DataGrid1
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
   End
   Begin VB.TextBox txtTitle
      DataField       =   "Title"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtPub
      DataField       =   "PName"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' master / detail with MSDataShape: a DataGrid and a text box bound to the chapter of the master

Private cn As ADODB.Connection
Private rsP As ADODB.Recordset
Private rsC As ADODB.Recordset
Private rsHeld As ADODB.Recordset
Private mStep As Integer

Private Function Grid() As String
    Dim s As String, i As Integer
    On Error Resume Next
    s = "grid cols=" & DataGrid1.Columns.Count
    For i = 0 To DataGrid1.Columns.Count - 1
        s = s & " " & DataGrid1.Columns(i).Caption & "=" & DataGrid1.Columns(i).Text
    Next
    s = s & " approx=" & DataGrid1.ApproxCount
    If Err.Number <> 0 Then s = s & " ! " & Err.Number
    Grid = s
End Function

Private Function Kids(rs As ADODB.Recordset) As String
    On Error Resume Next
    Kids = "n=" & rs.RecordCount & " eof=" & rs.EOF
    If Err.Number <> 0 Then Kids = "! " & Err.Number
End Function

Private Function Cur() As String
    Cur = "[" & txtPub.Text & "|" & txtTitle.Text & "] chapter " & Kids(rsC) & " held " & Kids(rsHeld) & " " & Grid()
End Function

Private Sub Form_Load()
    WriteLog "form2 load"
    Set cn = New ADODB.Connection
    cn.Open "Provider=MSDataShape;Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbPath
    Set rsP = New ADODB.Recordset
    rsP.CursorLocation = adUseClient
    rsP.Open "SHAPE {SELECT PID, PName FROM Pubs ORDER BY PID} APPEND ({SELECT BID, Title, PID FROM Books ORDER BY BID} AS Books RELATE PID TO PID)", cn, adOpenStatic, adLockOptimistic
    WriteLog "chapter field type=" & rsP("Books").Type & " value " & TypeName(rsP("Books").Value) & " underlying " & TypeName(rsP("Books").UnderlyingValue)
    Set rsHeld = rsP("Books").Value
    Set txtPub.DataSource = rsP
    Set rsC = rsP("Books").UnderlyingValue
    WriteLog "same " & (rsC Is rsHeld)
    Set DataGrid1.DataSource = rsC
    Set txtTitle.DataSource = rsC
    WriteLog "form2 load end " & Cur()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "form2 step " & mStep & " " & Cur()
    If mStep <= 3 Then
        rsP.MoveNext
        WriteLog "master " & rsP!PID & " " & Cur()
        Timer1.Enabled = True
    ElseIf mStep = 4 Then
        rsP.MoveFirst
        WriteLog "master first " & rsP!PID & " " & Cur()
        On Error Resume Next
        rsC.MoveLast
        WriteLog "detail last on an empty chapter: " & Err.Number
        Err.Clear
        On Error GoTo 0
        rsP.MoveNext
        rsC.MoveLast
        WriteLog "detail last " & rsP!PID & " " & Cur()
        Timer1.Enabled = True
    Else
        WriteLog "done"
        Unload Me
    End If
End Sub
