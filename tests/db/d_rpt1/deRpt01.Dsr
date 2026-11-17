VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} deRpt01 
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   _ExtentX        =   9102
   _ExtentY        =   8943
   FolderFlags     =   7
   TypeLibGuid     =   "{07F5AC10-9437-406D-B075-1CA60E607D52}"
   TypeInfoGuid    =   "{CFA99A0A-0A47-46A9-B92F-4793ED8A3C3D}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Conn1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\NoSuchFolder\rpt1.mdb"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   2
   BeginProperty Recordset1 
      CommandName     =   "Customers"
      CommDispId      =   1002
      RsDispId        =   1003
      CommandText     =   "SELECT * FROM Customers ORDER BY CustID"
      ActiveConnectionName=   "Conn1"
      CommandType     =   1
      Grouping        =   -1  'True
      GroupingName    =   "Customers_Grouping"
      Expanded        =   -1  'True
      SummaryExpanded =   -1  'True
      DetailExpanded  =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   255
         Size            =   5
         Scale           =   255
         Type            =   202
         Name            =   "CustID"
         Caption         =   "CustID"
      EndProperty
      BeginProperty Field2 
         Precision       =   255
         Size            =   30
         Scale           =   255
         Type            =   202
         Name            =   "Name"
         Caption         =   "Name"
      EndProperty
      BeginProperty Field3 
         Precision       =   255
         Size            =   15
         Scale           =   255
         Type            =   202
         Name            =   "Country"
         Caption         =   "Country"
      EndProperty
      GroupingName    =   "Customers_Grouping"
      NumGroups       =   1
      BeginProperty Grouping1 
         Precision       =   255
         Size            =   15
         Scale           =   255
         Type            =   202
         Name            =   "Country"
         Caption         =   "Country"
      EndProperty
      ParamCount      =   0
      CommandParent   =   ""
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "Orders"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "SELECT * FROM Orders ORDER BY OrderID"
      ActiveConnectionName=   "Conn1"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Customers"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   255
         Type            =   3
         Name            =   "OrderID"
         Caption         =   "OrderID"
      EndProperty
      BeginProperty Field2 
         Precision       =   255
         Size            =   5
         Scale           =   255
         Type            =   202
         Name            =   "CustID"
         Caption         =   "CustID"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   255
         Type            =   6
         Name            =   "Amount"
         Caption         =   "Amount"
      EndProperty
      BeginProperty Field4 
         Precision       =   255
         Size            =   8
         Scale           =   255
         Type            =   7
         Name            =   "ODate"
         Caption         =   "ODate"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      CommandParent   =   "Customers"
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "CustID"
         ChildField      =   "CustID"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "deRpt01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DataEnvironment_Initialize()
    Conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rpt1.mdb"
End Sub
