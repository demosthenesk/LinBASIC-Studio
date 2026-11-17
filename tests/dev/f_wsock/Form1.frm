VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Winsock"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3600
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wsSrv 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsCli 
      Left            =   720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsX 
      Left            =   1320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsU1 
      Left            =   1920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock wsU2 
      Left            =   2520
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   47112
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Winsock (MSWINSCK.OCX) over the loopback: a TCP server (a control array) and a client, UDP, the errors.
' The events of each side are kept apart (gS, gC, gU) and written at fixed points, so that the order in
' which Windows delivers the messages of two sockets does not change the output.
Private Const PORT As Long = 47110
Private gS As String, gC As String, gU As String, gRecv As String, gWant As Long, gUn As Integer, gNext As Integer
Private gAccepted As Boolean, gErr As Long, gX As String, gPeek As Boolean, gMax As Long, gSkip As Boolean

Private Sub L(ByVal s As String)
    WriteLog s
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
            If Asc(Mid$(v, i, 1)) < 32 Or Asc(Mid$(v, i, 1)) > 126 Then s = s & "<" & Asc(Mid$(v, i, 1)) & ">" Else s = s & Mid$(v, i, 1)
        Next
        s = s & "'"
    End If
    Hx = s
End Function

Private Function PV(c As Winsock, ByVal what As String) As String
    Dim v As Variant
    On Error Resume Next
    Select Case what
    Case "protocol": v = c.Protocol
    Case "remotehost": v = "'" & c.RemoteHost & "'"
    Case "remoteport": v = c.RemotePort
    Case "localport": v = c.LocalPort
    End Select
    ' a port the system chose: its number changes from run to run
    Select Case what
    Case "remoteport", "localport": If Err.Number = 0 And v <> 0 And (v < PORT Or v > PORT + 3) Then v = "eph"
    Case "state": v = c.State
    Case "bytes": v = c.BytesReceived
    Case "remotehostip": v = "'" & c.RemoteHostIP & "'"
    Case "sockethandle": v = IIf(c.SocketHandle = -1, -1, "socket")
    End Select
    If Err.Number <> 0 Then PV = " " & what & "=!" & Err.Number Else PV = " " & what & "=" & v
End Function

Private Sub Props(c As Winsock, ByVal nm As String)
    L "  " & nm & ":" & PV(c, "protocol") & PV(c, "remotehost") & PV(c, "remoteport") & PV(c, "localport") & PV(c, "state") & _
      PV(c, "bytes") & PV(c, "remotehostip") & PV(c, "sockethandle")
End Sub

Private Function Cond(ByVal what As String) As Boolean
    On Error Resume Next
    Select Case what
    Case "connected": Cond = wsCli.State = sckConnected And gAccepted
    Case "recv": Cond = Len(gRecv) >= gWant
    Case "clibytes": Cond = wsCli.BytesReceived >= gWant
    Case "sendcomplete": Cond = InStr(gC, "SendComplete") > 0
    Case "srvclose": Cond = InStr(gS, "Close") > 0
    Case "clierror": Cond = wsCli.State = sckError
    Case "xerror": Cond = wsX.State = sckError Or InStr(gX, "Error") > 0
    Case "udp": Cond = gUn >= gWant
    End Select
End Function

Private Function WaitFor(ByVal what As String) As Boolean
    Dim t As Single
    t = Timer
    Do
        DoEvents
        If Cond(what) Then WaitFor = True: Exit Function
    Loop While Timer - t < 8 And Timer >= t
    L "  (timeout: " & what & ")"
End Function

Private Sub Try(ByVal what As String)
    Dim v As Variant, o As Object
    On Error Resume Next
    Select Case what
    Case "X.SendData": wsX.SendData "abc"
    Case "X.GetData": wsX.GetData v
    Case "X.PeekData": wsX.PeekData v
    Case "X.Accept 12345": wsX.Accept 12345
    Case "X.Close": wsX.Close
    Case "X.Connect (no host)": wsX.Connect
    Case "X.Protocol=2": wsX.Protocol = 2
    Case "X.Protocol=-1": wsX.Protocol = -1
    Case "X.RemotePort=-1": wsX.RemotePort = -1
    Case "X.RemotePort=70000": wsX.RemotePort = 70000
    Case "X.LocalPort=-5": wsX.LocalPort = -5
    Case "X.LocalPort=70000": wsX.LocalPort = 70000
    Case "X.RemoteHostIP=": Set o = wsX: o.RemoteHostIP = "1.2.3.4"
    Case "X.State=": Set o = wsX: o.State = 1
    Case "X.BytesReceived=": Set o = wsX: o.BytesReceived = 1
    Case "X.Listen (UDP)": wsX.Protocol = sckUDPProtocol: wsX.Listen
    Case "X.Connect (UDP)": wsX.Protocol = sckUDPProtocol: wsX.Connect "127.0.0.1", PORT
    Case "X.Listen (port in use)": wsX.Protocol = sckTCPProtocol: wsX.LocalPort = PORT: wsX.Listen
    Case "S0.Listen again": wsSrv(0).Listen
    Case "S0.LocalPort= (listening)": wsSrv(0).LocalPort = PORT + 5
    Case "S0.Protocol= (listening)": wsSrv(0).Protocol = sckUDPProtocol
    Case "S0.Connect (listening)": wsSrv(0).Connect "127.0.0.1", PORT
    Case "S0.SendData (listening)": wsSrv(0).SendData "x"
    Case "C.Connect (connected)": wsCli.Connect "127.0.0.1", PORT
    Case "C.SendData Empty": wsCli.SendData Empty
    Case "C.SendData Null": wsCli.SendData Null
    Case "C.SendData Nothing": wsCli.SendData o
    Case "C.GetData vbVariant": wsCli.GetData v, vbVariant
    Case "C.GetData vbObject": wsCli.GetData v, vbObject
    Case "C.Connect (error state)": wsCli.Connect "127.0.0.1", PORT
    Case "S1.SendData (closing)": wsSrv(1).SendData "x"
    Case "S1.GetData (closing)": wsSrv(1).GetData v
    Case "U1.Bind (open)": wsU1.Bind PORT + 3
    Case "U1.Listen": wsU1.Listen
    Case "U1.Accept 1": wsU1.Accept 1
    Case "U2.SendData": wsU2.SendData "datagram one"
    Case "U2.SendData (again)": wsU2.SendData "datagram one"
    Case "U2.RemoteHost=, SendData": wsU2.RemoteHost = "127.0.0.1": wsU2.SendData "datagram one"
    Case "U2.Close, Bind, SendData": wsU2.Close: wsU2.Bind: wsU2.SendData "datagram one"
    End Select
    If Err.Number <> 0 Then
        L "  " & what & ": ! " & Err.Number & " " & Err.Description & " [" & Err.Source & "]"
    Else
        L "  " & what & ": ok " & TypeName(v)
    End If
    gErr = Err.Number
    Err.Clear
End Sub

Private Sub RunAll()
    Dim b(3) As Byte, bb() As Byte, v As Variant, s As String, i As Integer, lg As Long, d As Double
    Dim bo As Boolean, cy As Currency, f As Single, by As Byte, dt As Date

    L "== new controls"
    L "  TypeName " & TypeName(wsX) & " | host name " & (Len(wsX.LocalHostName) > 0) & " | LocalIP " & (wsX.LocalIP Like "*.*.*.*") & _
      " | SocketHandle " & wsX.SocketHandle
    Props wsX, "wsX"
    Props wsU1, "wsU1"
    Props wsU2, "wsU2"
    L "== errors"
    For Each v In Array("X.SendData", "X.GetData", "X.PeekData", "X.Accept 12345", "X.Close", "X.Protocol=2", "X.Protocol=-1", _
                        "X.RemotePort=-1", "X.RemotePort=70000", "X.LocalPort=-5", "X.LocalPort=70000", "X.RemoteHostIP=", "X.State=", _
                        "X.BytesReceived=", "X.Listen (UDP)", "X.Close")
        Try CStr(v)
    Next
    Try "X.Connect (UDP)"
    Props wsX, "wsX after Connect (UDP)"
    wsX.Close
    wsX.Protocol = sckTCPProtocol
    wsX.RemoteHost = ""
    Try "X.Connect (no host)"
    L "  X state " & wsX.State
    WaitFor "xerror"
    L "  X:" & gX & " | state " & wsX.State
    wsX.Close

    L "== TCP: listen, connect, accept"
    wsSrv(0).LocalPort = PORT
    wsSrv(0).Listen
    L "  listen: state " & wsSrv(0).State & " localport " & wsSrv(0).LocalPort
    For Each v In Array("S0.Listen again", "S0.LocalPort= (listening)", "S0.Protocol= (listening)", "S0.Connect (listening)", _
                        "S0.SendData (listening)", "X.Listen (port in use)")
        Try CStr(v)
    Next
    wsX.Close
    wsCli.RemoteHost = "127.0.0.1"
    wsCli.RemotePort = PORT
    wsCli.Connect
    L "  connect: state " & wsCli.State
    WaitFor "connected"
    L "  client:" & gC
    L "  server:" & gS
    gC = "": gS = ""
    L "  client: state " & wsCli.State & " remotehostip " & wsCli.RemoteHostIP & " remoteport " & wsCli.RemotePort & _
      " localport>0 " & (wsCli.LocalPort > 0)
    L "  wsSrv(1): state " & wsSrv(1).State & " remotehostip " & wsSrv(1).RemoteHostIP & " remoteport=client " & _
      (wsSrv(1).RemotePort = wsCli.LocalPort) & " localport " & wsSrv(1).LocalPort
    L "  wsSrv(0): state " & wsSrv(0).State & " remotehostip " & wsSrv(0).RemoteHostIP
    Try "C.Connect (connected)"

    L "== client -> server: a String"
    gWant = 9: gRecv = ""
    wsCli.SendData "Hello" & Chr$(0) & Chr$(255) & "AB"
    L "  senddata: state " & wsCli.State
    WaitFor "recv"
    L "  server got " & Hx(gRecv) & " |" & gS
    gS = ""
    WaitFor "sendcomplete"
    L "  client:" & gC
    gC = ""

    L "== server -> client: a Byte array"
    b(0) = 1: b(1) = 2: b(2) = 200: b(3) = 0
    gWant = 4
    wsSrv(1).SendData b
    WaitFor "clibytes"
    L "  client:" & gC & " | BytesReceived " & wsCli.BytesReceived
    gC = ""
    wsCli.PeekData v, vbArray + vbByte
    L "  peek " & Hx(v) & " | left " & wsCli.BytesReceived
    wsCli.GetData bb, vbArray + vbByte, 3
    L "  get 3: " & LBound(bb) & " To " & UBound(bb) & ":" & bb(0) & " " & bb(1) & " " & bb(2) & " | left " & wsCli.BytesReceived
    v = Empty
    wsCli.GetData v
    L "  the rest " & Hx(v) & " | left " & wsCli.BytesReceived
    wsCli.GetData v
    L "  nothing left " & Hx(v)

    L "== server -> client: numbers"
    gWant = 2 + 4 + 8 + 2 + 8 + 4 + 1 + 8
    wsSrv(1).SendData CInt(258)
    wsSrv(1).SendData CLng(16909060)
    wsSrv(1).SendData CDbl(1.5)
    wsSrv(1).SendData True
    wsSrv(1).SendData CCur(12.34)
    wsSrv(1).SendData CSng(-2.25)
    wsSrv(1).SendData CByte(77)
    wsSrv(1).SendData #1/2/2003 4:05:06 AM#
    WaitFor "clibytes"
    L "  BytesReceived " & wsCli.BytesReceived
    gC = ""
    wsCli.GetData i
    wsCli.GetData lg, vbLong
    wsCli.GetData d, vbDouble
    wsCli.GetData bo, vbBoolean
    wsCli.GetData cy, vbCurrency
    wsCli.GetData f, vbSingle
    wsCli.GetData by, vbByte
    wsCli.GetData dt, vbDate
    L "  " & i & " " & lg & " " & d & " " & bo & " " & cy & " " & f & " " & by & " " & Format$(dt, "yyyy-mm-dd hh:nn:ss") & _
      " | left " & wsCli.BytesReceived

    L "== server -> client: a String, GetData with maxLen"
    gWant = 6
    wsSrv(1).SendData "abcdef"
    WaitFor "clibytes"
    gC = ""
    wsCli.GetData s, , 4
    L "  maxLen 4: '" & s & "' | left " & wsCli.BytesReceived
    wsCli.GetData s
    L "  the rest: '" & s & "'"
    gWant = 3
    wsSrv(1).SendData "xyz"
    WaitFor "clibytes"
    gC = ""
    wsCli.PeekData s
    L "  peek: '" & s & "' | left " & wsCli.BytesReceived
    v = Empty
    wsCli.GetData v, vbString
    L "  get vbString: " & Hx(v) & " | left " & wsCli.BytesReceived
    For Each v In Array("C.SendData Empty", "C.SendData Null", "C.SendData Nothing", "C.GetData vbVariant", "C.GetData vbObject")
        Try CStr(v)
    Next
    gS = "": gRecv = ""

    L "== client closes"
    wsCli.Close
    L "  client: state " & wsCli.State
    WaitFor "srvclose"
    L "  server:" & gS & " | state " & wsSrv(1).State
    gS = ""
    Try "S1.SendData (closing)"
    Try "S1.GetData (closing)"
    wsSrv(1).Close
    L "  wsSrv(1).Close: state " & wsSrv(1).State
    wsSrv(0).Close
    L "  wsSrv(0).Close: state " & wsSrv(0).State

    L "== connection refused"
    gC = ""
    wsCli.Connect "127.0.0.1", PORT + 1
    L "  connect: state " & wsCli.State
    WaitFor "clierror"
    L "  client:" & gC & " | state " & wsCli.State
    gC = ""
    Try "C.Connect (error state)"
    wsCli.Close
    L "  close: state " & wsCli.State

    L "== a host name"
    gAccepted = False
    wsSrv(0).Listen
    wsCli.Connect "localhost", PORT
    L "  connect localhost: state " & wsCli.State
    WaitFor "connected"
    L "  client:" & gC
    L "  server:" & gS
    gC = "": gS = ""
    L "  client: state " & wsCli.State & " remotehost " & wsCli.RemoteHost & " remotehostip " & wsCli.RemoteHostIP
    wsCli.Close
    wsSrv(0).Close
    wsSrv(2).Close
    DoEvents

    L "== UDP"
    Props wsU1, "wsU1"
    Props wsU2, "wsU2"
    wsU1.Bind PORT + 2
    L "  bind: state " & wsU1.State & " localport " & wsU1.LocalPort
    Props wsU1, "wsU1"
    Try "U1.Bind (open)"
    Try "U1.Listen"
    Try "U1.Accept 1"
    gWant = 1: gUn = 0: gU = ""
    ' VB6: 10047 when the RemoteHost of a UDP control comes only from the .frm
    Try "U2.SendData"
    Props wsU2, "wsU2"
    Try "U2.SendData (again)"
    Try "U2.RemoteHost=, SendData"
    If gErr <> 0 Then L "  UDP: nothing could be sent": GoTo UdpEnd
    Props wsU2, "wsU2"
    WaitFor "udp"
    L "  U1:" & gU
    gU = ""
    Props wsU1, "wsU1 after the event"
    ' PeekData in the event
    gWant = 1: gUn = 0: gPeek = True
    wsU2.SendData "first"
    WaitFor "udp"
    gPeek = False
    L "  U1:" & gU
    gU = ""
    ' GetData with maxLen 3 of 7 bytes
    gWant = 1: gUn = 0: gMax = 3
    wsU2.SendData "second!"
    WaitFor "udp"
    gMax = 0
    L "  U1:" & gU
    gU = ""
    ' not read in the event
    gWant = 1: gUn = 0: gSkip = True
    wsU2.SendData "third"
    WaitFor "udp"
    gSkip = False
    L "  U1:" & gU
    gU = ""
    Props wsU1, "wsU1 after the event"
    s = "?"
    Err.Clear
    On Error Resume Next
    wsU1.GetData s
    L "  GetData after the event: err " & Err.Number & " '" & s & "'"
    On Error GoTo 0
    ' two datagrams, then the answer
    gWant = 2: gUn = 0
    wsU2.SendData "one"
    wsU2.SendData "two"
    WaitFor "udp"
    L "  U1:" & gU
    gU = ""
    gWant = 1: gUn = 0
    wsU1.RemoteHost = "127.0.0.1"
    wsU1.RemotePort = wsU2.LocalPort
    wsU1.SendData "reply"
    WaitFor "udp"
    L "  U2:" & gU
    gU = ""
UdpEnd:
    wsU1.Close
    wsU2.Close
    L "  closed: " & wsU1.State & " " & wsU2.State
    L "end"
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    RunAll
    Unload Me
End Sub

Private Sub wsCli_Connect()
    gC = gC & " Connect state=" & wsCli.State
End Sub

Private Sub wsCli_Close()
    gC = gC & " Close"
End Sub

Private Sub wsCli_DataArrival(ByVal bytesTotal As Long)
    gC = gC & " DataArrival(" & bytesTotal & ")"
End Sub

Private Sub wsCli_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    gC = gC & " SendProgress(" & bytesSent & "," & bytesRemaining & ")"
End Sub

Private Sub wsCli_SendComplete()
    gC = gC & " SendComplete"
End Sub

Private Sub wsCli_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, _
                        ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    gC = gC & " Error(" & Number & ", " & Description & ", " & Hex$(Scode) & ", " & Source & ", '" & HelpFile & "', " & _
         HelpContext & ", " & CancelDisplay & ") state=" & wsCli.State
    CancelDisplay = True
End Sub

Private Sub wsSrv_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    gS = gS & " ConnectionRequest(" & Index & ") state=" & wsSrv(Index).State & " ip=" & wsSrv(Index).RemoteHostIP
    gNext = gNext + 1
    Load wsSrv(gNext)
    wsSrv(gNext).Accept requestID
    gS = gS & " accepted by " & gNext & " state=" & wsSrv(gNext).State
    gAccepted = True
End Sub

Private Sub wsSrv_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim s As String
    gS = gS & " DataArrival(" & Index & "," & bytesTotal & ") bytes=" & wsSrv(Index).BytesReceived
    wsSrv(Index).GetData s, vbString
    gRecv = gRecv & s
End Sub

Private Sub wsSrv_Close(Index As Integer)
    gS = gS & " Close(" & Index & ") state=" & wsSrv(Index).State
End Sub

Private Sub wsSrv_SendComplete(Index As Integer)
End Sub

Private Sub wsU1_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    On Error Resume Next
    gU = gU & " U1.DataArrival(" & bytesTotal & ") bytes=" & wsU1.BytesReceived & " from " & wsU1.RemoteHostIP & ":" & _
         IIf(wsU1.RemotePort = wsU2.LocalPort, "U2", wsU1.RemotePort)
    If gSkip Then gUn = gUn + 1: Exit Sub
    If gPeek Then
        wsU1.PeekData s
        gU = gU & " peek '" & s & "' bytes=" & wsU1.BytesReceived
    End If
    If gMax > 0 Then wsU1.GetData s, vbString, gMax Else wsU1.GetData s
    If Err.Number <> 0 Then gU = gU & " ! " & Err.Number & " " & Err.Description: Err.Clear
    gU = gU & " got '" & s & "' bytes=" & wsU1.BytesReceived
    gUn = gUn + 1
End Sub

Private Sub wsU2_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    On Error Resume Next
    wsU2.GetData s
    gU = gU & " U2.DataArrival(" & bytesTotal & ") got '" & s & "' from " & wsU2.RemoteHostIP & ":" & _
         IIf(wsU2.RemotePort = wsU1.LocalPort, "U1", wsU2.RemotePort)
    If Err.Number <> 0 Then gU = gU & " ! " & Err.Number & " " & Err.Description: Err.Clear
    gUn = gUn + 1
End Sub

Private Sub wsU1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, _
                       ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    gU = gU & " U1.Error(" & Number & ")"
    CancelDisplay = True
End Sub

Private Sub wsX_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, _
                      ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    gX = gX & " Error(" & Number & ", " & Description & ") state=" & wsX.State
    CancelDisplay = True
End Sub

Private Sub wsX_Connect()
    gX = gX & " Connect"
End Sub
