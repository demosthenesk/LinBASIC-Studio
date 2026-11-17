Attribute VB_Name = "ModMci2"
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private gP As String

Private Function M(ByVal cmd As String) As String
    Dim b As String, r As Long
    b = String$(256, 0)
    r = mciSendString(cmd, b, 256, 0)
    If InStr(b, Chr$(0)) > 0 Then b = Left$(b, InStr(b, Chr$(0)) - 1)
    If r = 0 And Left$(cmd, 5) = "open " Then
        Out Replace(cmd, gP, "<dir>", , , vbTextCompare) & " -> ok"
    ElseIf r = 0 Then
        Out Replace(cmd, gP, "<dir>", , , vbTextCompare) & " -> '" & b & "'"
    Else
        Out Replace(cmd, gP, "<dir>", , , vbTextCompare) & " -> error " & r
    End If
    M = b
End Function

Private Sub Probe(ByVal f As String)
    Out "== " & f
    M "open """ & gP & "\" & f & """ type sequencer alias s"
    M "status s time format"
    M "status s length"
    M "status s tempo"
    M "set s time format ms"
    M "status s length"
    M "seek s to 500"
    M "status s position"
    M "status s tempo"
    M "seek s to 1000"
    M "status s position"
    M "seek s to 1500"
    M "status s position"
    M "status s tempo"
    M "seek s to 2500"
    M "status s position"
    M "seek s to end"
    M "status s position"
    M "status s tempo"
    M "set s time format song pointer"
    M "seek s to 4"
    M "status s position"
    M "set s time format ms"
    M "status s position"
    M "close s"
End Sub

Sub Main()
    OpenOut "d_mci2"
    gP = App.Path
    Probe "tune2.mid"
    Probe "tune3.mid"
    Probe "tune4.mid"
    Probe "tune5.mid"
    Probe "tune6.mid"
    Out "end"
    CloseOut
End Sub
