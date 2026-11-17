VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private rs As ADODB.Recordset, rs2 As ADODB.Recordset

Private Sub Class_Initialize()
    Set rs = New ADODB.Recordset
    rs.Fields.Append "Text", adVarChar, 20
    rs.Open
    rs.AddNew "Text", "one": rs.Update
    rs.AddNew "Text", "two": rs.Update
    rs.MoveFirst
    Set rs2 = New ADODB.Recordset
    rs2.Fields.Append "Text", adVarChar, 20
    rs2.Open
    rs2.AddNew "Text", "other one": rs2.Update
    rs2.MoveFirst
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    WriteLog "  getdatamember [" & DataMember & "]"
    If DataMember = "other" Then Set Data = rs2 Else Set Data = rs
End Sub

Public Sub MoveNext()
    rs.MoveNext
End Sub
