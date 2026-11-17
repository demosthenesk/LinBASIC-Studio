VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mRs As ADODB.Recordset
Attribute mRs.VB_VarHelpID = -1

Public Sub Run(cn As ADODB.Connection)
    On Error Resume Next
    Set mRs = New ADODB.Recordset
    mRs.Open "SELECT ID, Name, Age FROM People ORDER BY ID", cn, adOpenKeyset, adLockOptimistic
    Out "opened"
    mRs.MoveNext
    Out "next"
    mRs.MoveLast
    Out "last"
    mRs!Age = 99
    Out "assigned"
    mRs.Update
    Out "updated"
    mRs.AddNew
    Out "added"
    mRs!Name = "Gus"
    mRs.Update
    Out "updated new"
    mRs.Delete
    Out "deleted"
    mRs.MoveFirst
    mRs.MovePrevious
    Out "before first"
    mRs.Requery
    Out "requeried"
    mRs.Close
    Out "closed"
    Chk "events"
    Set mRs = Nothing
End Sub

Private Sub mRs_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "  willmove reason=" & adReason & " status=" & adStatus
End Sub

Private Sub mRs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "  movecomplete reason=" & adReason & " status=" & adStatus & " eof=" & pRecordset.EOF & " bof=" & pRecordset.BOF
End Sub

Private Sub mRs_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "  willchangefield n=" & cFields & " status=" & adStatus
End Sub

Private Sub mRs_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "  fieldchangecomplete n=" & cFields & " status=" & adStatus
End Sub

Private Sub mRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "  willchangerecord reason=" & adReason & " n=" & cRecords & " status=" & adStatus
End Sub

Private Sub mRs_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "  recordchangecomplete reason=" & adReason & " n=" & cRecords & " status=" & adStatus
End Sub

Private Sub mRs_WillChangeRecordset(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "  willchangerecordset reason=" & adReason & " status=" & adStatus
End Sub

Private Sub mRs_RecordsetChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "  recordsetchangecomplete reason=" & adReason & " status=" & adStatus
End Sub

Private Sub mRs_EndOfRecordset(fMoreData As Boolean, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Out "  endofrecordset more=" & fMoreData & " status=" & adStatus
End Sub
