Attribute VB_Name = "ModDao3"
Option Explicit

Private Sub ShowTd(td As TableDef)
    Dim fl As Field, ix As Index, i As Integer, s As String
    Out "table " & td.Name & " attr=" & td.Attributes & " fields=" & td.Fields.Count & " indexes=" & td.Indexes.Count & " rows=" & td.RecordCount
    For Each fl In td.Fields
        Out "  field " & fl.Name & " type=" & fl.Type & " size=" & fl.Size & " attr=" & fl.Attributes & _
            " req=" & fl.Required & " zl=" & fl.AllowZeroLength & " pos=" & fl.OrdinalPosition & " def=" & V(fl.DefaultValue)
    Next
    For Each ix In td.Indexes
        s = ""
        For i = 0 To ix.Fields.Count - 1
            s = s & " " & ix.Fields(i).Name & "/" & ix.Fields(i).Attributes
        Next
        Out "  index " & ix.Name & " primary=" & ix.Primary & " unique=" & ix.Unique & " req=" & ix.Required & _
            " ignorenulls=" & ix.IgnoreNulls & " foreign=" & ix.Foreign & " :" & s
    Next
End Sub

Private Sub ShowDb(ByVal file As String)
    Dim db As Database, td As TableDef, qd As QueryDef, rl As Relation, rs As Recordset, i As Integer, s As String
    On Error Resume Next
    Set db = OpenDatabase(App.Path & "\" & file, False, True)
    Chk "open " & file
    Out "== " & file & " version " & db.Version
    For Each td In db.TableDefs
        If (td.Attributes And dbSystemObject) = 0 Then ShowTd td
    Next
    Chk "tables"
    For Each rl In db.Relations
        s = ""
        For i = 0 To rl.Fields.Count - 1
            s = s & " " & rl.Fields(i).Name & "=" & rl.Fields(i).ForeignName
        Next
        Out "relation " & rl.Name & " " & rl.Table & "->" & rl.ForeignTable & " attr=" & rl.Attributes & " :" & s
    Next
    Chk "relations"
    For Each qd In db.QueryDefs
        Out "query " & qd.Name & " type=" & qd.Type & " params=" & qd.Parameters.Count & " returns=" & qd.ReturnsRecords
        Out "[" & qd.SQL & "]"
        If qd.Parameters.Count = 0 And (qd.Type = 0 Or qd.Type = 128) Then
            Set rs = qd.OpenRecordset(dbOpenSnapshot)
            If Err.Number = 0 Then
                If Not rs.EOF Then rs.MoveLast
                Out "  rows " & rs.RecordCount & " fields " & rs.Fields.Count
                rs.Close
            End If
            Chk "  open"
        End If
    Next
    Chk "queries"
    db.Close
End Sub

Sub Main()
    OpenOut "d_dao3"
    On Error Resume Next
    ShowDb "NWIND.MDB"
    ShowDb "BIBLIO.MDB"
    Out "end"
    CloseOut
End Sub
