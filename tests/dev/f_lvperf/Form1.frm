VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1
   Caption         =   "ListView speed"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7800
      Top             =   4080
   End
   Begin MSComctlLib.ListView ListView1
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628}
         Text            =   "Network Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628}
         SubItemIndex    =   1
         Text            =   "Host Address Range"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628}
         SubItemIndex    =   2
         Text            =   "Broadcast Address"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ListView speed: N items with two subitems each, added as the robot's IP subnet calculator does
' (ListItems.Add p, , text, then ListItems(p).ListSubItems.Add twice), Clear, items alone, items with SubItems()

Private Const N = 20000

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Report(ByVal what As String)
    With ListView1.ListItems
        If .Count = 0 Then
            WriteLog what & ": Count=0"
        ElseIf .Item(.Count).ListSubItems.Count > 0 Then
            WriteLog what & ": Count=" & .Count & " first=" & .Item(1).Text & " last=" & .Item(.Count).Text & _
                " sub=" & .Item(.Count).ListSubItems.Count & " " & .Item(.Count).SubItems(1) & " " & .Item(.Count).SubItems(2)
        Else
            WriteLog what & ": Count=" & .Count & " first=" & .Item(1).Text & " last=" & .Item(.Count).Text
        End If
    End With
End Sub

Private Sub Timer1_Timer()
    Dim p As Long, t As Single, it As ListItem
    Timer1.Enabled = False

    t = Timer
    For p = 1 To N
        ListView1.ListItems.Add p, , "10.0." & p
        ListView1.ListItems(p).ListSubItems.Add , , "10.1." & p & "  -  10.2." & p
        ListView1.ListItems(p).ListSubItems.Add , , "10.3." & p
    Next
    WriteTime "Add p + ListItems(p).ListSubItems.Add x2", Timer - t
    Report "subitems"

    t = Timer
    ListView1.ListItems.Clear
    WriteTime "Clear", Timer - t
    Report "clear"

    t = Timer
    For p = 1 To N
        ListView1.ListItems.Add , , "item " & p
    Next
    WriteTime "Add alone", Timer - t
    Report "alone"
    ListView1.ListItems.Clear

    t = Timer
    For p = 1 To N
        Set it = ListView1.ListItems.Add(, , "x" & p)
        it.SubItems(1) = "y" & p
        it.SubItems(2) = "z" & p
    Next
    WriteTime "Add + SubItems(1), SubItems(2)", Timer - t
    Report "subitems()"

    WriteLog "end"
    End
End Sub
