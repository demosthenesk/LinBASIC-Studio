VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      DataField       =   "Subject"
      DataSource      =   "Data1"
      Height          =   1425
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   2400
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      DataField       =   "PubID"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2400
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      BOFAction       =   1  'BOF
      Caption         =   "Titles"
      Connect         =   "Access 2000;"
      DatabaseName    =   "BIBLIO.MDB"
      DefaultCursorType=   1  'ODBCCursor
      DefaultType     =   1  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   -1  'True
      Height          =   360
      Left            =   0
      Options         =   64
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Titles"
      Top             =   3240
      Width           =   6000
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "ISBN"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
