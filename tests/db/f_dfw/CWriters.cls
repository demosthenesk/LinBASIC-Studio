VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWriters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' a data source class with one data member, as the class binding of a data form wizard
Private WithEvents rsW As ADODB.Recordset
Attribute rsW.VB_VarHelpID = -1
Public Event MoveComplete()

Private Sub Class_Initialize()
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbPath
    Set rsW = New ADODB.Recordset
    rsW.Open "SELECT WID, Name, Born FROM Writers ORDER BY WID", cn, adOpenStatic, adLockOptimistic
    WriteLog "  class: opened, members=" & DataMembers.Count
    DataMembers.Add "Primary"
    WriteLog "  class: members=" & DataMembers.Count
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    WriteLog "  class: getdatamember [" & DataMember & "]"
    If DataMember = "Primary" Then Set Data = rsW
End Sub

Private Sub rsW_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    WriteLog "  class: rs movecomplete " & adReason
    RaiseEvent MoveComplete
End Sub

Public Property Get AbsolutePosition() As Long
    AbsolutePosition = rsW.AbsolutePosition
End Property

Public Sub MoveNext()
    rsW.MoveNext
End Sub

Public Sub MoveLast()
    rsW.MoveLast
End Sub

Public Sub MoveFirst()
    rsW.MoveFirst
End Sub
