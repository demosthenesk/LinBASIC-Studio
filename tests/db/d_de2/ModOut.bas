Attribute VB_Name = "ModOut"
Option Explicit

Public gF As Integer

Public Sub OpenOut(ByVal test As String)
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    gF = FreeFile
    Open App.Path & "\expected\" & test & ".txt" For Output As #gF
End Sub

Public Sub CloseOut()
    Close #gF
End Sub

Public Sub Out(ByVal s As String)
    Print #gF, s
End Sub

' a value as text: its type and the value, without the regional settings
Public Function V(x As Variant) As String
    Dim t As String
    If IsObject(x) Then
        If x Is Nothing Then V = "Nothing" Else V = "O" & TypeName(x)
        Exit Function
    End If
    t = TypeName(x)
    Select Case VarType(x)
    Case vbNull: V = "Null"
    Case vbEmpty: V = "Empty"
    Case vbString: V = "S'" & x & "'"
    Case vbBoolean: V = "B" & IIf(x, "T", "F")
    Case vbDate: V = "D" & Format$(x, "yyyy-mm-dd hh:nn:ss")
    Case vbInteger, vbLong, vbByte: V = Left$(t, 1) & Trim$(Str$(x))
    Case vbSingle, vbDouble: V = Left$(t, 2) & Trim$(Str$(x))
    Case vbCurrency: V = "C" & Trim$(Str$(x))
    Case vbDecimal: V = "M" & Trim$(Str$(x))
    Case Else
        If (VarType(x) And vbArray) <> 0 Then
            V = "A" & t & "(" & LBound(x) & " To " & UBound(x) & ")"
        Else
            V = "?" & t & "/" & VarType(x)
        End If
    End Select
End Function

' the error, if any, then clear it
Public Sub Chk(ByVal what As String)
    If Err.Number <> 0 Then
        Out "  ! " & what & ": " & Err.Number & " " & Err.Description & " [" & Err.Source & "]"
        Err.Clear
    End If
End Sub

Public Sub ChkAlways(ByVal what As String)
    Out "  ! " & what & ": " & Err.Number & " " & Err.Description & " [" & Err.Source & "]"
    Err.Clear
End Sub

' the rows from the current one to the end
Public Sub DumpRest(rs As ADODB.Recordset, ByVal title As String)
    Dim i As Integer, s As String, n As Long
    Out "-- " & title
    Do Until rs.EOF
        s = " "
        For i = 0 To rs.Fields.Count - 1
            s = s & " " & V(rs.Fields(i).Value)
        Next
        Out s
        n = n + 1
        rs.MoveNext
        If n > 50 Then Exit Do
    Loop
    Out "  rows " & n
End Sub

Public Sub DumpRs(rs As ADODB.Recordset, ByVal title As String)
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    DumpRest rs, title
End Sub

Public Function Pos(rs As ADODB.Recordset) As String
    If rs.EOF Then
        Pos = "<eof>"
    ElseIf rs.BOF Then
        Pos = "<bof>"
    Else
        Pos = "@" & rs.AbsolutePosition & " " & rs.Fields("Name").Value
    End If
End Function
