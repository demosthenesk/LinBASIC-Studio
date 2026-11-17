VERSION 5.00
Begin VB.Form frmContacts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.Data datContacts 
      Caption         =   "Contacts"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Width           =   5655
   End
   Begin VB.TextBox txtNotes 
      DataField       =   "Notes"
      DataSource      =   "datContacts"
      Height          =   855
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "Email"
      DataSource      =   "datContacts"
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "Phone"
      DataSource      =   "datContacts"
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      DataField       =   "FullName"
      DataSource      =   "datContacts"
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblNotes 
      Caption         =   "No&tes:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label lblEmail 
      Caption         =   "&E-mail:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1500
      Width           =   975
   End
   Begin VB.Label lblPhone 
      Caption         =   "&Phone:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label lblName 
      Caption         =   "N&ame:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Address Book - DAO: a database made by the program (CreateDatabase, a TableDef, an index),
' a Data control and TextBoxes bound to its fields
Option Explicit

Private Sub Form_Load()
    Dim file As String
    file = App.Path & "\Contacts.mdb"
    If Dir$(file) = "" Then MakeDatabase file
    datContacts.DatabaseName = file
    datContacts.RecordSource = "SELECT * FROM Contacts ORDER BY FullName"
    datContacts.Refresh
    ShowCount
End Sub

Private Sub MakeDatabase(ByVal file As String)
    Dim db As DAO.Database, td As DAO.TableDef, idx As DAO.Index
    Set db = DBEngine.Workspaces(0).CreateDatabase(file, dbLangGeneral)
    Set td = db.CreateTableDef("Contacts")
    td.Fields.Append td.CreateField("ID", dbLong)
    td.Fields("ID").Attributes = dbAutoIncrField
    td.Fields.Append td.CreateField("FullName", dbText, 60)
    td.Fields.Append td.CreateField("Phone", dbText, 30)
    td.Fields.Append td.CreateField("Email", dbText, 60)
    td.Fields.Append td.CreateField("Notes", dbMemo)
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Fields.Append idx.CreateField("ID")
    td.Indexes.Append idx
    db.TableDefs.Append td
    db.Execute "INSERT INTO Contacts (FullName, Phone, Email, Notes) VALUES ('Ada Lovelace', '+44 20 7946 0001', 'ada@example.com', 'The first programmer.')"
    db.Execute "INSERT INTO Contacts (FullName, Phone, Email, Notes) VALUES ('Alan Kay', '+1 650 555 0102', 'alan@example.com', 'Smalltalk.')"
    db.Execute "INSERT INTO Contacts (FullName, Phone, Email, Notes) VALUES ('Grace Hopper', '+1 202 555 0103', 'grace@example.com', 'COBOL, and the first bug.')"
    db.Close
End Sub

Private Sub ShowCount()
    With datContacts.Recordset
        If .RecordCount = 0 Then
            lblCount.Caption = "empty"
        Else
            .MoveLast: .MoveFirst
            lblCount.Caption = .RecordCount & " contacts"
        End If
    End With
End Sub

Private Sub cmdAdd_Click()
    datContacts.Recordset.AddNew
    txtName.SetFocus
End Sub

Private Sub cmdDelete_Click()
    With datContacts.Recordset
        If .EOF Or .BOF Then Exit Sub
        If MsgBox("Delete " & txtName.Text & "?", vbQuestion + vbYesNo, "Address Book") = vbNo Then Exit Sub
        .Delete
        .MoveNext
        If .EOF And .RecordCount > 0 Then .MoveLast
    End With
    ShowCount
End Sub

Private Sub datContacts_Reposition()
    datContacts.Caption = "Contact " & (datContacts.Recordset.AbsolutePosition + 1)
End Sub
