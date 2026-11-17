VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} deRpt02 
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   _ExtentX        =   9102
   _ExtentY        =   8943
   FolderFlags     =   7
   TypeLibGuid     =   "{942A8A74-8144-4043-9EC6-7BC57232DC33}"
   TypeInfoGuid    =   "{130DDD16-0DF1-4671-89B9-A5AA15DC06D9}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Conn1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\NoSuchFolder\rpt2.mdb"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   1
   BeginProperty Recordset1 
      CommandName     =   "Nums"
      CommDispId      =   1002
      RsDispId        =   1003
      CommandText     =   "SELECT N, Word FROM Nums ORDER BY N"
      ActiveConnectionName=   "Conn1"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   255
         Type            =   3
         Name            =   "N"
         Caption         =   "N"
      EndProperty
      BeginProperty Field2 
         Precision       =   255
         Size            =   40
         Scale           =   255
         Type            =   202
         Name            =   "Word"
         Caption         =   "Word"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "deRpt02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DataEnvironment_Initialize()
    Conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rpt2.mdb"
End Sub
