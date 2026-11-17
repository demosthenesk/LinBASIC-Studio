VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "msrdc20.Ocx"
Begin VB.Form Form1 
   Caption         =   "RDO 3"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin MSRDC.MSRDC MSRDC1 
      Height          =   330
      Left            =   120
      Top             =   1560
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      _Version        =   393216
      Options         =   0
      CursorDriver    =   1
      BOFAction       =   1
      EOFAction       =   2
      RecordsetType   =   3
      LockType        =   4
      QueryType       =   0
      Prompt          =   1
      Appearance      =   1
      QueryTimeout    =   20
      RowsetSize      =   10
      LoginTimeout    =   10
      KeysetSize      =   50
      MaxRows         =   100
      ErrorThreshold  =   -1
      BatchSize       =   20
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   -1  'True
      DataSourceName  =   ""
      RecordSource    =   "SELECT * FROM People ORDER BY ID"
      UserName        =   "bob"
      Password        =   ""
      Connect         =   "Driver={SQLite3 ODBC Driver};Database=f_rdo3.db"
      LogMessages     =   ""
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' RDO 3 (LinBASIC Studio phase 7.8): the properties of the RemoteData control as the VB6 IDE writes them in the .frm
' (set in the Properties window, then saved), read back when the form loads: it opens on its own

Private Sub MSRDC1_Error(ByVal Number As Long, Description As String, ByVal Scode As Long, ByVal Source As String, _
        ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLog "  ev error " & Number & " [" & Source & "] " & Description
    CancelDisplay = True
End Sub

Private Sub Form_Load()
    Dim s As String
    On Error Resume Next
    With MSRDC1
        s = "load dsn=[" & .DataSourceName & "] sql=[" & .SQL & "] prompt=" & .Prompt & _
            " type=" & .ResultsetType & " lock=" & .LockType & " cursor=" & .CursorDriver & " keyset=" & .KeysetSize & _
            " rowset=" & .RowsetSize & " maxrows=" & .MaxRows & " options=" & .Options & " login=" & .LoginTimeout & _
            " qtime=" & .QueryTimeout & " ro=" & .ReadOnly & " eof=" & .EOFAction & " bof=" & .BOFAction & " enabled=" & .Enabled & _
            " caption=[" & .Caption & "] app=" & .Appearance & " back=" & Hex$(.BackColor) & " fore=" & Hex$(.ForeColor) & _
            " batch=" & .BatchSize & " crit=" & .UpdateCriteria & " op=" & .UpdateOperation & " errthr=" & .ErrorThreshold & _
            " user=[" & .UserName & "] font=" & .Font.Name & " " & .Font.Size
    End With
    If Err.Number <> 0 Then s = s & " ! " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog s
    WriteLog "load text1=[" & Text1.Text & "] rs=" & TypeName(MSRDC1.Resultset)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    WriteLog "activate text1=[" & Text1.Text & "] rs=" & TypeName(MSRDC1.Resultset) & " rows=" & MSRDC1.Resultset.RowCount & _
        " connect=" & (InStr(MSRDC1.Connect, "SQLite3") > 0)
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description
    Unload Me
End Sub
