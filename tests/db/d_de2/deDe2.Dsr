VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} deDe2 
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   _ExtentX        =   9102
   _ExtentY        =   8943
   FolderFlags     =   7
   TypeLibGuid     =   "{7A0E5C21-3B49-4C4E-9F11-2D6B8E3A0C51}"
   TypeInfoGuid    =   "{5C2B7E90-1D84-4A6F-8B3E-6F0A9C2D4E73}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Conn1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\NoSuchFolder\de2.mdb"
      Expanded        =   -1  'True
      QuoteChar       =   0
      SeparatorChar   =   0
   EndProperty
   NumRecordsets   =   4
   BeginProperty Recordset1 
      CommandName     =   "ByCountry"
      CommDispId      =   1002
      RsDispId        =   1003
      CommandText     =   "SELECT CustID, Name FROM Customers WHERE Country = ? ORDER BY CustID"
      ActiveConnectionName=   "Conn1"
      CommandType     =   1
      ParentCommandName=   ""
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
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
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Param1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "Range"
      CommDispId      =   1004
      RsDispId        =   1005
      CommandText     =   "SELECT OrderID, Amount FROM Orders WHERE Amount >= ? AND Amount <= ? ORDER BY OrderID"
      ActiveConnectionName=   "Conn1"
      CommandType     =   1
      ParentCommandName=   ""
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   255
         Type            =   3
         Name            =   "OrderID"
         Caption         =   "OrderID"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   255
         Type            =   6
         Name            =   "Amount"
         Caption         =   "Amount"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "MinA"
         UserName        =   "MinA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "MaxA"
         UserName        =   "MaxA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "Cust"
      CommDispId      =   1006
      RsDispId        =   1007
      CommandText     =   "SELECT CustID, Name FROM Customers ORDER BY CustID"
      ActiveConnectionName=   "Conn1"
      CommandType     =   1
      ParentCommandName=   ""
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
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
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "CustOrders"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "SELECT OrderID, CustID, Amount FROM Orders WHERE CustID = ? ORDER BY OrderID"
      ActiveConnectionName=   "Conn1"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Cust"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
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
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Param1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "CustID"
         ChildField      =   "Param1"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "deDe2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DataEnvironment_Initialize()
    Conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\de2.mdb"
End Sub

Private Sub Conn1_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Out "event Conn1_WillExecute source=" & Source
End Sub
