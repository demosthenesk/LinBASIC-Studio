VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Common controls 2"
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
      Interval        =   200
      Left            =   6600
      Top             =   4800
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3413
      _Version        =   393217
      Style           =   7
      LineStyle       =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   1440
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "size"
         Text            =   "Size"
         Object.Width           =   900
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   3
      Cols            =   3
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      _Version        =   393216
      Tabs            =   3
      TabsPerRow      =   3
      TabHeight       =   520
      TabCaption(0)   =   "Alpha"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Beta"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Gamma"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command1 
         Caption         =   "On Beta"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "On Alpha"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   2
      Max             =   5
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mStep As Integer

Private Sub Form_Load()
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_comctl2.txt"
    On Error GoTo 0
    WriteLog "tree " & TreeView1.Nodes.Count & " " & TreeView1.Style & " " & TreeView1.LineStyle & " " & TypeName(TreeView1)
    WriteLog "list " & ListView1.View & " " & ListView1.ColumnHeaders.Count & " " & ListView1.ColumnHeaders(2).Text & " " & ListView1.ColumnHeaders("size").Index
    WriteLog "grid " & MSFlexGrid1.Rows & " " & MSFlexGrid1.Cols & " " & MSFlexGrid1.FixedRows & " " & MSFlexGrid1.FixedCols & " " & MSFlexGrid1.Row & " " & MSFlexGrid1.Col
    WriteLog "sstab " & SSTab1.Tabs & " " & SSTab1.Tab & " " & SSTab1.TabCaption(1) & " " & Command1.Left & " " & Label1.Left
    WriteLog "updown " & UpDown1.Min & " " & UpDown1.Max & " " & UpDown1.Value & " " & UpDown1.Increment
    Timer1.Enabled = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    WriteLog "sstab click " & PreviousTab & " -> " & SSTab1.Tab
End Sub

Private Sub UpDown1_Change()
    WriteLog "updown change " & UpDown1.Value
End Sub

Private Sub Report(ByVal what As String)
    WriteLog what & " err " & Err.Number
    Err.Clear
End Sub

Private Sub Timer1_Timer()
    Dim n As MSComctlLib.Node, li As MSComctlLib.ListItem, s As String, v As Variant, i As Integer
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        On Error Resume Next
        With TreeView1.Nodes
            .Add , , "r1", "Root1"
            .Add , , "r2", "Root2"
            .Add "r1", tvwChild, "c1", "Child1"
            .Add "r1", tvwChild, "c2", "Child2"
            .Add "c1", tvwNext, "c15", "Child15"
            .Add "c1", tvwPrevious, "c05", "Child05"
            .Add "c1", tvwChild, "g1", "Grand1"
            .Add "r1", tvwFirst, "r0", "Root0"
            .Add "r2", tvwLast, "r3", "Root3"
        End With
        WriteLog "nodes " & TreeView1.Nodes.Count & " " & TreeView1.Nodes(1).Text & " " & TreeView1.Nodes("g1").Index
        s = ""
        Set n = TreeView1.Nodes("r1").Child
        Do While Not n Is Nothing
            s = s & n.Text & ","
            Set n = n.Next
        Loop
        WriteLog "children of r1 " & TreeView1.Nodes("r1").Children & " " & s
        Set n = TreeView1.Nodes("r0")
        WriteLog "roots " & n.FirstSibling.Text & " " & n.LastSibling.Text & " " & n.Next.Text
        WriteLog "path " & TreeView1.Nodes("g1").FullPath & " " & TreeView1.Nodes("g1").Parent.Text & " " & TreeView1.Nodes("g1").Root.Text
        WriteLog "expanded " & TreeView1.Nodes("r1").Expanded & " visible " & TreeView1.Nodes("g1").Visible
        TreeView1.Nodes("g1").EnsureVisible
        WriteLog "expanded " & TreeView1.Nodes("r1").Expanded & " " & TreeView1.Nodes("c1").Expanded & " visible " & TreeView1.Nodes("g1").Visible
        TreeView1.Nodes.Add "zz", tvwChild, "x", "x"
        Report "relative missing"
        TreeView1.Nodes.Add , , "r1", "dup"
        Report "node dup"
        Set TreeView1.Nodes("c2").Parent = TreeView1.Nodes("r2")
        WriteLog "moved " & TreeView1.Nodes("c2").Parent.Text & " " & TreeView1.Nodes("r1").Children & " " & TreeView1.Nodes("r2").Children
        Set TreeView1.Nodes("r1").Parent = TreeView1.Nodes("g1")
        Report "parent loop"
        TreeView1.Nodes.Remove "c1"
        WriteLog "removed " & TreeView1.Nodes.Count & " " & TreeView1.Nodes("r1").Children & " " & TreeView1.Nodes(5).Text
        Set TreeView1.SelectedItem = TreeView1.Nodes("r2")
        WriteLog "selected " & TreeView1.SelectedItem.Key & " " & TreeView1.Nodes("r2").Selected

        ListView1.ColumnHeaders.Add , "date", "Date", 1200
        WriteLog "columns " & ListView1.ColumnHeaders.Count & " " & ListView1.ColumnHeaders(3).Key & " " & ListView1.ColumnHeaders(3).SubItemIndex
        Set li = ListView1.ListItems.Add(, "a", "beta")
        li.SubItems(1) = "20"
        li.SubItems(2) = "x"
        Set li = ListView1.ListItems.Add(, "b", "alpha")
        li.SubItems(1) = "3"
        Set li = ListView1.ListItems.Add(1, "c", "gamma")
        li.SubItems(1) = "100"
        WriteLog "items " & ListView1.ListItems.Count & " " & ListView1.ListItems(1).Text & " " & ListView1.ListItems("a").Index & " " & ListView1.ListItems("a").SubItems(1)
        li.SubItems(5) = "no"
        Report "subitem 5"
        WriteLog "find " & ListView1.FindItem("alp", lvwText, , lvwPartial).Key
        ListView1.SortKey = 1
        ListView1.Sorted = True
        s = ""
        For Each v In ListView1.ListItems
            s = s & v.Text & ","
        Next
        WriteLog "sorted " & s & " " & ListView1.ListItems("c").Index
        ListView1.SortOrder = lvwDescending
        s = ""
        For i = 1 To ListView1.ListItems.Count
            s = s & ListView1.ListItems(i).Key & ","
        Next
        WriteLog "desc " & s
        ListView1.Sorted = False
        Set ListView1.SelectedItem = ListView1.ListItems("b")
        WriteLog "selected " & ListView1.SelectedItem.Key & " " & ListView1.ListItems("b").Selected & " " & ListView1.ListItems("a").Selected
        ListView1.ListItems.Remove "a"
        WriteLog "after remove " & ListView1.ListItems.Count & " " & ListView1.ListItems("c").Index
        ListView1.ListItems.Add , "b", "dup"
        Report "item dup"

        With MSFlexGrid1
            .TextMatrix(0, 1) = "H1"
            .TextMatrix(1, 1) = "a"
            .TextMatrix(2, 2) = "b"
            .AddItem "x" & vbTab & "10" & vbTab & "z"
            .AddItem "y" & vbTab & "2", 1
            WriteLog "grid rows " & .Rows & " " & .TextMatrix(1, 0) & " " & .TextMatrix(4, 1) & " " & .TextMatrix(3, 2)
            .Row = 2: .Col = 1
            .Text = "set"
            WriteLog "grid text " & .TextMatrix(2, 1) & " " & .Row & " " & .Col
            .RowSel = 3: .ColSel = 2
            WriteLog "clip [" & Replace(Replace(.Clip, vbTab, "|"), vbCr, "/") & "]"
            .Col = 1: .ColSel = 1
            .Sort = flexSortNumericAscending
            s = ""
            For i = 1 To .Rows - 1
                s = s & .TextMatrix(i, 1) & ","
            Next
            WriteLog "grid sorted " & s
            .RemoveItem 1
            WriteLog "grid removed " & .Rows & " " & .TextMatrix(1, 1)
            .Row = 10
            Report "grid row 10"
            v = .TextMatrix(9, 9)
            Report "grid textmatrix"
            .FormatString = "<Name|>Age|^Mid"
            WriteLog "format " & .TextMatrix(0, 0) & " " & .TextMatrix(0, 2) & " " & .ColAlignment(1) & " " & .Cols
        End With

        SSTab1.Tab = 1
        WriteLog "sstab " & SSTab1.Tab & " " & Command1.Left & " " & Label1.Left
        SSTab1.Tab = 5
        Report "sstab 5"
        SSTab1.TabCaption(2) = "G"
        WriteLog "caption " & SSTab1.TabCaption(2)

        UpDown1.Value = 4
        UpDown1.Value = 9
        Report "updown 9"
        WriteLog "updown " & UpDown1.Value
    Case 2
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub
