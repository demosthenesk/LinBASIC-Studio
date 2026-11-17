VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "MSComm"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   120
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      RThreshold      =   1
      InputLen        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gStep As Integer, gWait As Integer, gEv1 As String, gEv2 As String

Private Sub Props(c As MSComm)
    On Error Resume Next
    WriteLog "  port=" & c.CommPort & " settings=" & c.Settings & " open=" & c.PortOpen & " in=" & c.InBufferSize & _
        " out=" & c.OutBufferSize & " inputlen=" & c.InputLen & " mode=" & c.InputMode & " rth=" & c.RThreshold & _
        " sth=" & c.SThreshold & " hs=" & c.Handshaking & " dtr=" & c.DTREnable & " rts=" & c.RTSEnable & _
        " null=" & c.NullDiscard & " eof=" & c.EOFEnable & " parrep='" & c.ParityReplace & "'"
    If Err.Number <> 0 Then WriteLog "  ! props " & Err.Number & " " & Err.Description: Err.Clear
    WriteLog "  cdto=" & c.CDTimeout & " ctsto=" & c.CTSTimeout & " dsrto=" & c.DSRTimeout & " interval=" & c.Interval & _
        " break=" & c.Break
    If Err.Number <> 0 Then WriteLog "  ! props2 " & Err.Number & " " & Err.Description: Err.Clear
End Sub

Private Sub Try(c As MSComm, ByVal what As String)
    Dim v As Variant
    On Error Resume Next
    Select Case what
    Case "InBufferCount": v = c.InBufferCount
    Case "OutBufferCount": v = c.OutBufferCount
    Case "Input": v = c.Input
    Case "CDHolding": v = c.CDHolding
    Case "CTSHolding": v = c.CTSHolding
    Case "DSRHolding": v = c.DSRHolding
    Case "Output": c.Output = "x"
    Case "Break": c.Break = True
    Case "InBufferCount=0": c.InBufferCount = 0
    Case "OutBufferCount=0": c.OutBufferCount = 0
    Case "Close": c.PortOpen = False
    End Select
    If Err.Number <> 0 Then
        WriteLog "  " & what & ": ! " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        WriteLog "  " & what & ": ok " & TypeName(v) & " " & v
    End If
End Sub

Private Sub SetP(c As MSComm, ByVal what As String, ByVal v As Variant)
    On Error Resume Next
    Select Case what
    Case "Settings": c.Settings = v
    Case "CommPort": c.CommPort = v
    Case "InBufferSize": c.InBufferSize = v
    Case "OutBufferSize": c.OutBufferSize = v
    Case "InputLen": c.InputLen = v
    Case "InputMode": c.InputMode = v
    Case "RThreshold": c.RThreshold = v
    Case "SThreshold": c.SThreshold = v
    Case "Handshaking": c.Handshaking = v
    Case "ParityReplace": c.ParityReplace = v
    Case "CDTimeout": c.CDTimeout = v
    Case "Interval": c.Interval = v
    Case "CommEvent": c.CommEvent = v
    Case "CommID": c.CommID = v
    End Select
    If Err.Number <> 0 Then
        WriteLog "  " & what & "=" & v & ": ! " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        Select Case what
        Case "Settings": WriteLog "  " & what & "=" & v & ": ok " & c.Settings
        Case "CommPort": WriteLog "  " & what & "=" & v & ": ok " & c.CommPort
        Case "InBufferSize": WriteLog "  " & what & "=" & v & ": ok " & c.InBufferSize
        Case "InputLen": WriteLog "  " & what & "=" & v & ": ok " & c.InputLen
        Case "ParityReplace": WriteLog "  " & what & "=" & v & ": ok '" & c.ParityReplace & "'"
        Case Else: WriteLog "  " & what & "=" & v & ": ok"
        End Select
    End If
End Sub

Private Function Hx(v As Variant) As String
    Dim i As Long, s As String
    If IsArray(v) Then
        s = TypeName(v) & "(" & LBound(v) & " To " & UBound(v) & ")"
        For i = LBound(v) To UBound(v)
            s = s & " " & v(i)
        Next
    Else
        s = TypeName(v) & " '"
        For i = 1 To Len(v)
            If Asc(Mid$(v, i, 1)) < 32 Then s = s & "<" & Asc(Mid$(v, i, 1)) & ">" Else s = s & Mid$(v, i, 1)
        Next
        s = s & "'"
    End If
    Hx = s
End Function

Private Sub Form_Load()
    Dim i As Integer, v As Variant
    WriteLog "load " & TypeName(MSComm1)
    WriteLog "MSComm1"
    Props MSComm1
    WriteLog "MSComm2"
    Props MSComm2
    WriteLog "== closed port"
    For Each v In Array("InBufferCount", "OutBufferCount", "Input", "CDHolding", "CTSHolding", "DSRHolding", "Output", "Break", "InBufferCount=0", "OutBufferCount=0", "Close")
        Try MSComm1, CStr(v)
    Next
    WriteLog "== settings"
    For Each v In Array("19200,e,7,2", "9600,N,8,1", "110,m,5,1.5", "300,s,6,1", "128000,o,8,1", "256000,n,8,1", "1200,n,4,1", "1200,n,9,1", "1200,x,8,1", _
                        "1200,n,8,3", "1200,n,8,2.5", "abc", "", "9600", "9600,n", "9600,n,8", " 9600 , n , 8 , 1", "9600,n,8,1,", "1234,n,8,1", "0,n,8,1", "-9600,n,8,1", "9600,n,8,1")
        SetP MSComm1, "Settings", v
    Next
    WriteLog "== other properties"
    For Each v In Array(0, 1, 16, 17, 99, 255, 256, -1)
        SetP MSComm1, "CommPort", v
    Next
    For Each v In Array(0, 1, 1024, 32767, -1)
        SetP MSComm1, "InBufferSize", v
    Next
    SetP MSComm1, "OutBufferSize", 0
    SetP MSComm1, "OutBufferSize", 512
    SetP MSComm1, "InputLen", -1
    SetP MSComm1, "InputLen", 5
    SetP MSComm1, "InputLen", 0
    SetP MSComm1, "InputMode", 2
    SetP MSComm1, "InputMode", -1
    SetP MSComm1, "RThreshold", -1
    SetP MSComm1, "SThreshold", -1
    SetP MSComm1, "Handshaking", 4
    SetP MSComm1, "Handshaking", -1
    SetP MSComm1, "ParityReplace", "ab"
    SetP MSComm1, "ParityReplace", ""
    SetP MSComm1, "ParityReplace", "?"
    SetP MSComm1, "CDTimeout", 100
    SetP MSComm1, "Interval", 100
    SetP MSComm1, "CommEvent", 2
    SetP MSComm1, "CommID", 5
    WriteLog "== a port that is not there"
    MSComm1.CommPort = 9
    On Error Resume Next
    MSComm1.PortOpen = True
    WriteLog "  open 9: " & Err.Number & " " & Err.Description & " open=" & MSComm1.PortOpen: Err.Clear
    On Error GoTo 0
    WriteLog "== COM1 and COM2"
    MSComm1.CommPort = 1
    MSComm1.Settings = "9600,n,8,1"
    MSComm1.InBufferSize = 1024
    MSComm1.OutBufferSize = 512
    MSComm1.CDTimeout = 0
    MSComm1.Interval = 1000
    On Error Resume Next
    MSComm1.PortOpen = True
    If Err.Number <> 0 Then WriteLog "  open 1: " & Err.Number & " " & Err.Description: Err.Clear
    MSComm2.PortOpen = True
    If Err.Number <> 0 Then WriteLog "  open 2: " & Err.Number & " " & Err.Description: Err.Clear
    On Error GoTo 0
    If Not (MSComm1.PortOpen And MSComm2.PortOpen) Then
        WriteLog "  no pair of ports: open1=" & MSComm1.PortOpen & " open2=" & MSComm2.PortOpen
        If MSComm1.PortOpen Then MSComm1.PortOpen = False
        If MSComm2.PortOpen Then MSComm2.PortOpen = False
        gStep = 99
        Timer1.Enabled = True
        Exit Sub
    End If
    WriteLog "  opened"
    Props MSComm1
    On Error Resume Next
    MSComm1.PortOpen = True
    WriteLog "  open again: " & Err.Number & " " & Err.Description: Err.Clear
    MSComm1.CommPort = 3
    WriteLog "  commport while open: " & Err.Number & " " & Err.Description & " port=" & MSComm1.CommPort: Err.Clear
    MSComm1.InBufferSize = 2048
    WriteLog "  inbuffersize while open: " & Err.Number & " " & Err.Description & " size=" & MSComm1.InBufferSize: Err.Clear
    MSComm1.Settings = "19200,n,8,1"
    WriteLog "  settings while open: " & Err.Number & " " & Err.Description & " " & MSComm1.Settings: Err.Clear
    MSComm1.Settings = "9600,n,8,1"
    MSComm1.Settings = "bad"
    WriteLog "  bad settings while open: " & Err.Number & " " & Err.Description & " " & MSComm1.Settings: Err.Clear
    WriteLog "  id>0=" & (MSComm1.CommID > 0) & " in=" & MSComm1.InBufferCount & " out=" & MSComm1.OutBufferCount
    MSComm1.Output = 5
    WriteLog "  output a number: " & Err.Number & " " & Err.Description: Err.Clear
    On Error GoTo 0
    MSComm1.InBufferCount = 0
    MSComm2.InBufferCount = 0
    gEv1 = "": gEv2 = ""
    MSComm1.Output = "HELLO"
    gStep = 1: gWait = 0
    Timer1.Enabled = True
End Sub

Private Sub MSComm1_OnComm()
    gEv1 = gEv1 & " " & MSComm1.CommEvent
End Sub

Private Sub MSComm2_OnComm()
    If Right$(gEv2, 2) <> " " & MSComm2.CommEvent Then gEv2 = gEv2 & " " & MSComm2.CommEvent
End Sub

Private Function Arrived(ByVal n As Integer) As Boolean
    gWait = gWait + 1
    Arrived = MSComm2.InBufferCount >= n Or gWait > 50
End Function

Private Sub Timer1_Timer()
    Dim b() As Byte, v As Variant
    On Error GoTo EH
    Select Case gStep
    Case 99
        WriteLog "end"
        Timer1.Enabled = False
        Unload Me
    Case 1
        If Not Arrived(5) Then Exit Sub
        WriteLog "== text"
        WriteLog "  count=" & MSComm2.InBufferCount & " events2=" & gEv2 & " events1=" & gEv1 & " out=" & MSComm1.OutBufferCount
        MSComm2.InputLen = 2
        WriteLog "  input " & Hx(MSComm2.Input) & " left=" & MSComm2.InBufferCount
        MSComm2.InputLen = 0
        WriteLog "  input " & Hx(MSComm2.Input) & " left=" & MSComm2.InBufferCount
        WriteLog "  input " & Hx(MSComm2.Input) & " left=" & MSComm2.InBufferCount
        WriteLog "  event=" & MSComm2.CommEvent
        gEv2 = ""
        MSComm2.NullDiscard = True
        MSComm1.Output = "A" & Chr$(0) & "B"
        gStep = 2: gWait = 0
    Case 2
        If Not Arrived(2) Then Exit Sub
        WriteLog "== null discard"
        WriteLog "  count=" & MSComm2.InBufferCount & " input " & Hx(MSComm2.Input)
        MSComm2.NullDiscard = False
        MSComm2.EOFEnable = True
        gEv2 = ""
        MSComm1.Output = "C" & Chr$(26) & "D"
        gStep = 3: gWait = 0
    Case 3
        If Not Arrived(3) Then Exit Sub
        WriteLog "== eof"
        WriteLog "  count=" & MSComm2.InBufferCount & " events2=" & gEv2 & " input " & Hx(MSComm2.Input)
        MSComm2.EOFEnable = False
        WriteLog "== binary"
        MSComm2.InputMode = comInputModeBinary
        ReDim b(0 To 4)
        b(0) = 0: b(1) = 1: b(2) = 128: b(3) = 255: b(4) = 65
        MSComm1.Output = b
        gStep = 4: gWait = 0
    Case 4
        If Not Arrived(5) Then Exit Sub
        WriteLog "  count=" & MSComm2.InBufferCount
        MSComm2.InputLen = 3
        v = MSComm2.Input
        WriteLog "  input " & Hx(v)
        MSComm2.InputLen = 0
        v = MSComm2.Input
        WriteLog "  input " & Hx(v)
        v = MSComm2.Input
        WriteLog "  empty " & Hx(v) & " isarray=" & IsArray(v) & " isempty=" & IsEmpty(v)
        MSComm2.InputMode = comInputModeText
        WriteLog "== output a string in binary mode, and 8 bit text"
        MSComm1.Output = Chr$(200) & Chr$(10) & "z"
        gStep = 5: gWait = 0
    Case 5
        If Not Arrived(3) Then Exit Sub
        v = MSComm2.Input
        WriteLog "  input len=" & Len(v) & " asc=" & Asc(Mid$(v, 1, 1)) & "," & Asc(Mid$(v, 2, 1)) & "," & Asc(Mid$(v, 3, 1))
        WriteLog "== sthreshold"
        gEv1 = ""
        MSComm1.SThreshold = 1
        MSComm1.Output = "12345"
        gStep = 6: gWait = 0
    Case 6
        If Not Arrived(5) Then Exit Sub
        WriteLog "  events1=" & gEv1
        v = MSComm2.Input
        MSComm1.SThreshold = 0
        WriteLog "== rthreshold 0: no events"
        gEv2 = ""
        MSComm2.RThreshold = 0
        MSComm1.Output = "xyz"
        gStep = 7: gWait = 0
    Case 7
        If Not Arrived(3) Then Exit Sub
        WriteLog "  events2=[" & gEv2 & "] count=" & MSComm2.InBufferCount
        MSComm2.InBufferCount = 0
        WriteLog "  cleared count=" & MSComm2.InBufferCount
        WriteLog "== close"
        MSComm1.PortOpen = False
        MSComm2.PortOpen = False
        WriteLog "  open=" & MSComm1.PortOpen & " event=" & MSComm1.CommEvent
        On Error Resume Next
        WriteLog "  count after close=" & MSComm2.InBufferCount
        If Err.Number <> 0 Then WriteLog "  ! " & Err.Number & " " & Err.Description: Err.Clear
        WriteLog "end"
        Timer1.Enabled = False
        Unload Me
    End Select
    Exit Sub
EH:
    WriteLog "  !! " & Err.Number & " " & Err.Description & " step=" & gStep
    Resume Next
End Sub
