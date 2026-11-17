VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} deTest 
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   _ExtentX        =   9102
   _ExtentY        =   8943
   FolderFlags     =   7
   TypeLibGuid     =   "{4581853E-4C19-40E7-B309-B5142C3B32C9}"
   TypeInfoGuid    =   "{0035ADF0-6A95-4FFA-A594-0179C9F05C5C}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Conn1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\NoSuchFolder\de1.mdb"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   4
   BeginProperty Recordset1 
      CommandName     =   "Customers"
      CommDispId      =   1002
      RsDispId        =   1003
      CommandText     =   "Customers"
      ActiveConnectionName=   "Conn1"
      CommandType     =   2
      dbObjectType    =   1
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
      CommandText     =   "Orders"
      ActiveConnectionName=   "Conn1"
      CommandType     =   2
      dbObjectType    =   1
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
      AggregateCount  =   2
      BeginProperty Aggregate1 
         Name            =   "TotQty"
         AggOn           =   "Items"
         AggField        =   "Qty"
         AggType         =   1
         AggFunction     =   7
         Type            =   5
         Name            =   "TotQty"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
      BeginProperty Aggregate2 
         Name            =   "NItems"
         AggOn           =   "Items"
         AggField        =   "Product"
         AggType         =   1
         AggFunction     =   3
         Type            =   3
         Name            =   "NItems"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "Items"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "SELECT OrderID, Product, Qty, Price FROM Items ORDER BY ItemID"
      ActiveConnectionName=   "Conn1"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Orders"
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
         Size            =   20
         Scale           =   255
         Type            =   202
         Name            =   "Product"
         Caption         =   "Product"
      EndProperty
      BeginProperty Field3 
         Precision       =   5
         Size            =   2
         Scale           =   255
         Type            =   2
         Name            =   "Qty"
         Caption         =   "Qty"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   255
         Type            =   6
         Name            =   "Price"
         Caption         =   "Price"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      CommandParent   =   "Orders"
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "OrderID"
         ChildField      =   "OrderID"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "Germans"
      CommDispId      =   1004
      RsDispId        =   1005
      CommandText     =   "SELECT * FROM Customers WHERE Country = 'Germany' ORDER BY CustID"
      ActiveConnectionName=   "Conn1"
      CommandType     =   1
      Expanded        =   -1  'True
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
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "deTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DataEnvironment_Initialize()
    Out "event DataEnvironment_Initialize conn state=" & Conn1.State
    Conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\de1.mdb"
End Sub

Private Sub DataEnvironment_Terminate()
    Out "event DataEnvironment_Terminate"
End Sub

Private Sub Conn1_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Out "event Conn1_WillConnect status=" & adStatus
End Sub

Private Sub Conn1_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Out "event Conn1_ConnectComplete status=" & adStatus
End Sub

Private Sub Conn1_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Out "event Conn1_WillExecute cursortype=" & CursorType & " locktype=" & LockType & " options=" & Options & " source=" & Source
End Sub

Private Sub Conn1_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Out "event Conn1_ExecuteComplete status=" & adStatus
End Sub

Private Sub rsGermans_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "event rsGermans_MoveComplete reason=" & adReason
End Sub
