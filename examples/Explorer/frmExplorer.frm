VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExplorer 
   Caption         =   "Explorer"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   4725
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11033
            Key             =   "folder"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "count"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   4575
      Left            =   2880
      TabIndex        =   2
      Top             =   60
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.DirListBox dirFolders 
      Height          =   4140
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   2745
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2745
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Explorer - the drive and folder lists of VB, the ListView and StatusBar of the Common Controls,
' and the files of a folder with Dir$, FileLen and FileDateTime
Option Explicit

Private Sub Form_Load()
    ShowFiles
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub
    If ScaleHeight < 1500 Or ScaleWidth < 4000 Then Exit Sub
    dirFolders.Height = ScaleHeight - sbrStatus.Height - dirFolders.Top - 60
    lvwFiles.Move lvwFiles.Left, lvwFiles.Top, ScaleWidth - lvwFiles.Left - 60, ScaleHeight - sbrStatus.Height - 120
End Sub

Private Sub drvDrives_Change()
    On Error GoTo NotReady
    dirFolders.Path = drvDrives.Drive
    Exit Sub
NotReady:
    MsgBox "The drive is not ready.", vbExclamation, "Explorer"
    drvDrives.Drive = dirFolders.Path
End Sub

Private Sub dirFolders_Change()
    ShowFiles
End Sub

Private Sub ShowFiles()
    Dim folder As String, name As String, item As ListItem, total As Double, n As Long
    folder = dirFolders.Path
    If Right$(folder, 1) <> "\" And Right$(folder, 1) <> "/" Then folder = folder & "\"
    lvwFiles.ListItems.Clear
    On Error Resume Next
    name = Dir$(folder & "*.*")
    Do While name <> ""
        Set item = lvwFiles.ListItems.Add(, , name)
        item.SubItems(1) = Format$(FileLen(folder & name) / 1024, "#,##0.0") & " KB"
        item.SubItems(2) = Format$(FileDateTime(folder & name), "yyyy-mm-dd hh:nn")
        total = total + FileLen(folder & name)
        n = n + 1
        name = Dir$
    Loop
    sbrStatus.Panels("folder").Text = dirFolders.Path
    sbrStatus.Panels("count").Text = n & " files, " & Format$(total / 1024, "#,##0") & " KB"
End Sub

Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwFiles.SortKey = ColumnHeader.Index - 1
    lvwFiles.SortOrder = IIf(lvwFiles.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    lvwFiles.Sorted = True
End Sub
