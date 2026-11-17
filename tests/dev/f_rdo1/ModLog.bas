Attribute VB_Name = "ModLog"
Option Explicit

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_rdo1.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

' Err, then the rdoErrors of the engine
Public Sub Chk(ByVal what As String)
    Dim e As rdoError, n As Long, d As String
    If Err.Number = 0 Then Exit Sub
    n = Err.Number: d = Err.Description
    WriteLog "  " & what & " ! " & n & " [" & Err.Source & "] " & d
    Err.Clear
    For Each e In rdoEngine.rdoErrors
        WriteLog "    rdoError " & e.Number & " state=" & e.SQLState & " ret=" & e.SQLRetcode & " [" & e.Source & "] " & e.Description
    Next
End Sub

Public Function V(x As Variant) As String
    If IsNull(x) Then
        V = "Null"
    ElseIf IsEmpty(x) Then
        V = "Empty"
    Else
        V = TypeName(x) & ":" & CStr(x)
    End If
End Function
