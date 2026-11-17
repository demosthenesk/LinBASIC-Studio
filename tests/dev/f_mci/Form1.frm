VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "MMControl"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1500
      Left            =   120
      ScaleHeight     =   1440
      ScaleWidth      =   1935
      TabIndex        =   2
      Top             =   1200
      Width           =   2000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   120
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MMControl2 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      PlayEnabled     =   -1  'True
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      AutoEnable      =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gStep As Integer, gWait As Integer, gDone As Integer, gUpd As Long, gProbe As Integer
Private gCodes As String, gWhat As String

Private Function B(ByVal x As Boolean) As String
    B = IIf(x, "T", "F")
End Function

Private Function Btns(m As MMControl) As String
    Btns = "en=" & B(m.PrevEnabled) & B(m.NextEnabled) & B(m.PlayEnabled) & B(m.PauseEnabled) & B(m.BackEnabled) & _
        B(m.StepEnabled) & B(m.StopEnabled) & B(m.RecordEnabled) & B(m.EjectEnabled) & _
        " vis=" & B(m.PrevVisible) & B(m.NextVisible) & B(m.PlayVisible) & B(m.PauseVisible) & B(m.BackVisible) & _
        B(m.StepVisible) & B(m.StopVisible) & B(m.RecordVisible) & B(m.EjectVisible)
End Function

Private Function St(m As MMControl) As String
    St = "mode=" & m.Mode & " pos=" & m.Position & " len=" & m.Length & " err=" & m.Error & " msg='" & m.ErrorMessage & "'"
End Function

Private Sub Cmd(m As MMControl, ByVal c As String)
    On Error Resume Next
    m.Command = c
    If Err.Number <> 0 Then
        WriteLog "  " & c & ": ! " & Err.Number & " " & Err.Description & " | error=" & m.Error & " msg='" & m.ErrorMessage & "'"
        Err.Clear
    Else
        WriteLog "  " & c & ": ok " & St(m)
    End If
End Sub

Private Sub Props(m As MMControl)
    On Error Resume Next
    WriteLog "  type='" & m.DeviceType & "' file='" & m.FileName & "' auto=" & m.AutoEnable & " notify=" & m.Notify & _
        " wait=" & m.Wait & " share=" & m.Shareable & " orient=" & m.Orientation & " tf=" & m.TimeFormat & _
        " upd=" & m.UpdateInterval & " silent=" & m.Silent & " recmode=" & m.RecordMode & " border=" & m.BorderStyle & _
        " enabled=" & m.Enabled & " cmd='" & m.Command & "'"
    WriteLog "  " & Btns(m)
    WriteLog "  mode=" & m.Mode & " error=" & m.Error & " msg='" & m.ErrorMessage & "' id=" & m.DeviceID & _
        " from=" & m.From & " to=" & m.To & " track=" & m.Track & " frames=" & m.Frames & " notifyvalue=" & m.NotifyValue & _
        " hwnddisplay=" & m.hWndDisplay & " usewin=" & m.UsesWindows
    If Err.Number <> 0 Then WriteLog "  ! props " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "  len=" & m.Length: If Err.Number <> 0 Then WriteLog "  ! len " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "  pos=" & m.Position: If Err.Number <> 0 Then WriteLog "  ! pos " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "  start=" & m.Start: If Err.Number <> 0 Then WriteLog "  ! start " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "  tracks=" & m.Tracks: If Err.Number <> 0 Then WriteLog "  ! tracks " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "  tracklen=" & m.TrackLength: If Err.Number <> 0 Then WriteLog "  ! tracklen " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "  trackpos=" & m.TrackPosition: If Err.Number <> 0 Then WriteLog "  ! trackpos " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "  can play=" & m.CanPlay & " record=" & m.CanRecord & " eject=" & m.CanEject & " step=" & m.CanStep
    If Err.Number <> 0 Then WriteLog "  ! can " & Err.Number & " " & Err.Description: Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    WriteLog "load"
    WriteLog "MMControl1 " & TypeName(MMControl1)
    Props MMControl1
    WriteLog "MMControl2"
    Props MMControl2
    On Error Resume Next
    MMControl1.TimeFormat = 11
    WriteLog "tf 11: " & Err.Number & " " & Err.Description: Err.Clear
    MMControl1.Orientation = 2
    WriteLog "orient 2: " & Err.Number & " " & Err.Description: Err.Clear
    MMControl1.UpdateInterval = -1
    WriteLog "upd -1: " & Err.Number & " " & Err.Description & " upd=" & MMControl1.UpdateInterval: Err.Clear
    MMControl1.Mode = 525
    WriteLog "mode set: " & Err.Number & " " & Err.Description: Err.Clear
    MMControl1.Length = 5
    WriteLog "len set: " & Err.Number & " " & Err.Description: Err.Clear
    MMControl1.Error = 5
    WriteLog "error set: " & Err.Number & " " & Err.Description: Err.Clear
    On Error GoTo EH
    WriteLog "== open without a file"
    Cmd MMControl1, "Open"
    WriteLog "== bad command"
    Cmd MMControl1, "Bogus"
    Cmd MMControl1, "Play"
    WriteLog "== missing file"
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.FileName = App.Path & "\nofile.wav"
    Cmd MMControl1, "Open"
    WriteLog "  open devices=" & OpenDevices()
    WriteLog "== open tone.wav"
    MMControl1.FileName = App.Path & "\tone.wav"
    Cmd MMControl1, "Open"
    WriteLog "  open devices=" & OpenDevices() & " wait=" & MMControl1.Wait & " from=" & MMControl1.From & " to=" & MMControl1.To
    Props MMControl1
    MMControl1.TimeFormat = 8
    WriteLog "  bytes: len=" & MMControl1.Length & " tf=" & MMControl1.TimeFormat
    MMControl1.TimeFormat = 9
    WriteLog "  samples: len=" & MMControl1.Length
    MMControl1.TimeFormat = 0
    WriteLog "== play with wait"
    MMControl1.From = 0
    MMControl1.To = 200
    MMControl1.Wait = True
    Cmd MMControl1, "Play"
    WriteLog "  after: wait=" & MMControl1.Wait & " notify=" & MMControl1.Notify & " from=" & MMControl1.From & " to=" & MMControl1.To
    WriteLog "  " & Btns(MMControl1)
    WriteLog "== seek / prev / next / step / back"
    MMControl1.To = 700
    Cmd MMControl1, "Seek"
    On Error Resume Next
    MMControl1.Position = 0
    WriteLog "  position set: " & Err.Number & " " & Err.Description & " | " & St(MMControl1)
    Err.Clear
    On Error GoTo EH
    WriteLog "  open devices=" & OpenDevices()
    Cmd MMControl1, "Next"
    Cmd MMControl1, "Prev"
    Cmd MMControl1, "Step"
    Cmd MMControl1, "Back"
    Cmd MMControl1, "Eject"
    Cmd MMControl1, "Sound"
    Cmd MMControl1, "Pause"
    Cmd MMControl1, "Stop"
    WriteLog "== wait: one command or all?"
    MMControl1.Command = "Stop"
    MMControl1.Wait = True
    MMControl1.From = 0
    MMControl1.To = 300
    MMControl1.Command = "Play"
    WriteLog "  play 1: mode=" & MMControl1.Mode & " wait=" & MMControl1.Wait
    MMControl1.From = 0
    MMControl1.To = 300
    MMControl1.Command = "Play"
    WriteLog "  play 2: mode=" & MMControl1.Mode & " wait=" & MMControl1.Wait
    MMControl1.Wait = False
    MMControl1.Command = "Stop"
    MMControl1.Wait = False
    gStep = 1
    Timer1.Enabled = True
    Exit Sub
EH:
    WriteLog "  !! " & Err.Number & " " & Err.Description
    Resume Next
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
    gDone = gDone + 1
    gCodes = gCodes & IIf(gCodes = "", "", " ") & NotifyCode
End Sub

Private Sub MMControl1_StatusUpdate()
    gUpd = gUpd + 1
End Sub

Private Sub MMControl1_PlayCompleted(Errorcode As Long)
    WriteLog "  playcompleted " & Errorcode
End Sub

Private Sub Timer1_Timer()
    On Error GoTo EH
    gWait = gWait + 1
    Select Case gStep
    Case 1
        ' each probe: start it, then wait 2 seconds without reading the control, then the Done codes
        gWait = 0: gCodes = ""
        Select Case gProbe
        Case 0
            WriteLog "== probes of Notify / Done (the codes in the order they came)"
            MMControl1.UpdateInterval = 100
            gUpd = 0
            MMControl1.Command = "Prev"
            MMControl1.Notify = True
            MMControl1.Command = "Play"
            gWhat = "A1 notify, play to the end"
        Case 1
            MMControl1.Command = "Prev"
            MMControl1.From = 0
            MMControl1.To = 400
            MMControl1.Notify = True
            MMControl1.Command = "Play"
            gWhat = "A2 notify, play from 0 to 400"
        Case 2
            MMControl1.Command = "Prev"
            MMControl1.Command = "Play"
            gWhat = "A3 play, no notify set again"
        Case 3
            MMControl1.Command = "Prev"
            MMControl1.Notify = True
            MMControl1.Command = "Play"
            MMControl1.Notify = True
            MMControl1.Command = "Play"
            gWhat = "A4 notify play, notify play"
        Case 4
            MMControl1.Command = "Prev"
            MMControl1.Notify = True
            MMControl1.Command = "Play"
            MMControl1.Command = "Stop"
            gWhat = "A5 notify play, stop"
        Case 5
            MMControl1.Command = "Prev"
            MMControl1.Notify = True
            MMControl1.Command = "Play"
            MMControl1.Notify = True
            MMControl1.Command = "Stop"
            gWhat = "A6 notify play, notify stop"
        Case 6
            MMControl1.Command = "Prev"
            MMControl1.Notify = True
            gWhat = "A7 notify, then reading Mode and Position: " & MMControl1.Mode & " " & MMControl1.Position
        Case 7
            MMControl1.Command = "Prev"
            MMControl1.Notify = True
            MMControl1.Command = "Seek"
            gWhat = "A8 notify seek"
        Case 8
            MMControl1.Command = "Prev"
            MMControl1.Notify = True
            MMControl1.Command = "Step"
            gWhat = "A9 notify step (waveaudio)"
        Case Else
            gStep = 3
            Exit Sub
        End Select
        gStep = 2
    Case 2
        If gWait < 20 Then Exit Sub
        WriteLog "  " & gWhat & ": [" & gCodes & "] mode=" & MMControl1.Mode & " notify=" & MMControl1.Notify & " value=" & MMControl1.NotifyValue
        gProbe = gProbe + 1
        gStep = 1
    Case 3
        MMControl1.UpdateInterval = 0
        WriteLog "  updates>0=" & (gUpd > 0)
        WriteLog "== close"
        Cmd MMControl1, "Close"
        WriteLog "  open devices=" & OpenDevices()
        Props MMControl1
        WriteLog "== avivideo in a picture box"
        MMControl1.DeviceType = "AVIVideo"
        MMControl1.FileName = App.Path & "\clip.avi"
        Cmd MMControl1, "Open"
        MMControl1.hWndDisplay = Picture1.hWnd
        WriteLog "  tf=" & MMControl1.TimeFormat & " len=" & MMControl1.Length & " frames=" & MMControl1.Frames & " usewin=" & MMControl1.UsesWindows & " canstep=" & MMControl1.CanStep & " hwnd ok=" & (MMControl1.hWndDisplay = Picture1.hWnd)
        Cmd MMControl1, "Step"
        Cmd MMControl1, "Step"
        Cmd MMControl1, "Back"
        MMControl1.Frames = 3
        Cmd MMControl1, "Step"
        MMControl1.Wait = True
        MMControl1.From = 2
        MMControl1.To = 6
        Cmd MMControl1, "Play"
        Cmd MMControl1, "Close"
        WriteLog "== MMControl2 (AutoEnable off)"
        MMControl2.FileName = App.Path & "\tone.wav"
        Cmd MMControl2, "Open"
        WriteLog "  " & Btns(MMControl2)
        MMControl2.Command = "Close"
        WriteLog "end"
        WriteLog "  open devices=" & OpenDevices()
        Timer1.Enabled = False
        Unload Me
    End Select
    Exit Sub
EH:
    WriteLog "  !! " & Err.Number & " " & Err.Description & " step=" & gStep
    Resume Next
End Sub
