Attribute VB_Name = "ModMci"
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, _
    ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private gP As String

' the path of the test folder in the commands: the folder is replaced by <dir>
Private Function Shown(ByVal s As String) As String
    Shown = Replace(s, gP, "<dir>", , , vbTextCompare)
End Function

Private Function ErrText(ByVal e As Long) As String
    Dim b As String, r As Long
    b = String$(256, 0)
    r = mciGetErrorString(e, b, 256)
    If InStr(b, Chr$(0)) > 0 Then b = Left$(b, InStr(b, Chr$(0)) - 1)
    ErrText = r & " '" & b & "'"
End Function

' send a command; write the command, the result and the returned string
Private Function M(ByVal cmd As String) As String
    Dim b As String, r As Long
    b = String$(256, 0)
    r = mciSendString(cmd, b, 256, 0)
    If InStr(b, Chr$(0)) > 0 Then b = Left$(b, InStr(b, Chr$(0)) - 1)
    If r = 0 Then
        Out Shown(cmd) & " -> '" & Shown(b) & "'"
    Else
        Out Shown(cmd) & " -> error " & r & " (" & (r And &HFFFF&) & ")"
    End If
    M = b
End Function

Private Sub Status(ByVal dev As String, ParamArray items() As Variant)
    Dim i As Integer
    For i = 0 To UBound(items)
        M "status " & dev & " " & items(i)
    Next
End Sub

Sub Main()
    Dim i As Long, s As String, b As String, r As Long
    OpenOut "d_mci"
    gP = App.Path

    Out "== waveaudio"
    M "open """ & gP & "\tone.wav"" type waveaudio alias w"
    Status "w", "mode", "length", "time format", "position", "start position", "ready", "media present", _
        "number of tracks", "current track", "channels", "samplespersec", "bitspersample", "bytespersec", _
        "alignment", "format tag", "level", "input", "output"
    M "set w time format samples"
    Status "w", "length", "time format"
    M "set w time format bytes"
    Status "w", "length", "time format"
    M "set w time format ms"
    Status "w", "time format"
    M "set w time format milliseconds"
    M "seek w to 500"
    Status "w", "position", "mode"
    M "seek w to end"
    Status "w", "position", "mode"
    M "seek w to start"
    Status "w", "position"
    M "play w from 0 to 300 wait"
    Status "w", "position", "mode"
    M "play w from 1000 wait"
    Status "w", "position", "mode"
    M "set w time format samples"
    M "play w from 0 to 2205 wait"
    Status "w", "position"
    M "set w time format ms"
    M "capability w can play"
    M "capability w can record"
    M "capability w can save"
    M "capability w can eject"
    M "capability w device type"
    M "capability w has audio"
    M "capability w has video"
    M "capability w uses files"
    M "capability w compound device"
    M "capability w can insert"
    M "info w file"
    M "pause w"
    Status "w", "mode"
    M "resume w"
    M "stop w"
    Status "w", "mode"
    M "play w from 2000"
    M "seek w to 99999"
    M "status w bogus"
    M "status w"
    M "set w channels 2"
    M "seek w"
    M "play w to 100 from 50 wait"
    Status "w", "position"
    M "sysinfo all quantity open"
    M "sysinfo all name 1 open"
    M "sysinfo w installname"
    M "open """ & gP & "\tone.wav"" type waveaudio alias w"
    M "close w"
    Status "w", "mode"
    M "close w"

    Out "== waveaudio 8 bit stereo"
    M "open """ & gP & "\beep8.wav"" alias b"
    Status "b", "length", "channels", "samplespersec", "bitspersample", "bytespersec", "alignment"
    M "set b time format samples"
    Status "b", "length"
    M "set b time format bytes"
    Status "b", "length"
    M "set b time format ms"
    M "seek b to 250"
    M "set b time format bytes"
    Status "b", "position"
    M "close b"

    Out "== open errors"
    M "open """ & gP & "\nofile.wav"" type waveaudio alias x"
    M "open """ & gP & "\tone.wav"" type nodevice alias x"
    M "open """ & gP & "\tone.wav"" alias w1"
    M "open """ & gP & "\tone.wav"" alias w1"
    M "close w1"
    M "open new type waveaudio alias r"
    Status "r", "length", "mode", "media present"
    M "capability r can record"
    M "close r"
    M "bogus w"
    M "play"
    M "status nodev mode"
    M ""

    Out "== sequencer"
    M "open """ & gP & "\tune.mid"" type sequencer alias s"
    Status "s", "length", "time format", "tempo", "number of tracks", "division", "position", "mode", "media present"
    M "set s time format song pointer"
    Status "s", "length", "time format"
    M "set s time format ms"
    M "seek s to 1000"
    Status "s", "position"
    M "seek s to end"
    Status "s", "position", "tempo"
    M "capability s can record"
    M "capability s device type"
    M "capability s can save"
    M "close s"

    Out "== avivideo"
    M "open """ & gP & "\clip.avi"" type avivideo alias v"
    Status "v", "length", "time format", "frame rate", "mode", "position", "speed", "nominal frame rate"
    M "where v source"
    M "where v destination"
    M "set v time format ms"
    Status "v", "length"
    M "seek v to 500"
    M "set v time format frames"
    Status "v", "position"
    M "capability v has video"
    M "capability v device type"
    M "capability v can stretch"
    M "close v"

    Out "== mpegvideo"
    M "open """ & gP & "\tone.wav"" type mpegvideo alias m"
    Status "m", "length", "time format", "mode"
    M "capability m device type"
    M "close m"

    Out "== close all"
    M "open """ & gP & "\tone.wav"" alias a1"
    M "open """ & gP & "\beep8.wav"" alias a2"
    M "sysinfo all quantity open"
    M "close all"
    M "sysinfo all quantity open"

    Out "== error strings"
    For i = 256 To 350
        s = ErrText(i)
        If Left$(s, 1) <> "0" Then Out i & " " & s
    Next
    Out ErrText(0)
    Out ErrText(5)
    Out "end"
    CloseOut
End Sub
