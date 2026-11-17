VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1
   Caption         =   "Inet HTTP"
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
   Begin InetCtlsObjects.Inet Inet1
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet2
      Left            =   960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The Inet control (MSINET.OCX) with HTTP, against tests/dev/inet_server.py (port 47120): the URL and its
' parts, OpenURL, Execute with GET / HEAD / POST / PUT, GetHeader, GetChunk, Cancel, the errors.
' The StateChanged events are written once each, in the order they first came (expected\f_inet-raw.txt
' has them all).
Private gEv As String, gBase As String

Private Sub L(ByVal s As String)
    WriteLog s
End Sub

Private Function Hx(v As Variant, Optional ByVal maxc As Long = 400) As String
    Dim i As Long, s As String, c As Long, n As Long
    If IsArray(v) Then
        n = 0
        For i = LBound(v) To UBound(v)
            n = (n + v(i) * ((i Mod 7) + 1)) Mod 65521
        Next
        s = TypeName(v) & "(" & LBound(v) & " To " & UBound(v) & ") sum " & n
        If UBound(v) >= LBound(v) Then s = s & " first " & v(LBound(v)) & " last " & v(UBound(v))
    ElseIf IsNull(v) Or IsEmpty(v) Then
        s = TypeName(v)
    Else
        s = TypeName(v) & "(" & Len(v) & ") '"
        For i = 1 To Len(v)
            If i > maxc Then s = s & "...": Exit For
            c = AscW(Mid$(v, i, 1))
            If c < 32 Or c > 126 Then s = s & "<" & c & ">" Else s = s & Mid$(v, i, 1)
        Next
        s = s & "'"
    End If
    Hx = s
End Function

' the states in the order they first came
' the states that came, in numeric order (the order in which they come changes from run to run); without
' 1 and 2 (resolving the name): WinINet gives them or not, as its cache of names has it
Private Function Sq(ByVal ev As String) As String
    Dim a() As String, i As Long, s As String, seen(20) As Boolean
    a = Split(Trim$(ev), " ")
    For i = 0 To UBound(a)
        If a(i) <> "" Then
            If Val(a(i)) >= 0 And Val(a(i)) <= 20 Then seen(Val(a(i))) = True
        End If
    Next
    For i = 3 To 20
        If seen(i) Then s = s & " " & i
    Next
    Sq = Trim$(s)
End Function

Private Sub Props(c As Inet, ByVal nm As String)
    On Error Resume Next
    Dim s As String
    s = "  " & nm & ": URL='" & c.URL & "'"
    s = s & " Protocol=" & c.Protocol
    s = s & " RemoteHost='" & c.RemoteHost & "'"
    s = s & " RemotePort=" & c.RemotePort
    s = s & " Document='" & c.Document & "'"
    s = s & " UserName='" & c.UserName & "'"
    s = s & " Password='" & c.Password & "'"
    If Err.Number <> 0 Then s = s & " !" & Err.Number: Err.Clear
    L s
    s = "    AccessType=" & c.AccessType & " Proxy='" & c.Proxy & "' RequestTimeout=" & c.RequestTimeout & _
        " ResponseCode=" & c.ResponseCode & " ResponseInfo=" & Hx(c.ResponseInfo, 200) & " StillExecuting=" & c.StillExecuting & _
        " hInternet=" & IIf(c.hInternet = 0, "0", "handle")
    If Err.Number <> 0 Then s = s & " !" & Err.Number & " " & Err.Description: Err.Clear
    L s
End Sub

Private Sub SetU(ByVal u As String)
    On Error Resume Next
    Inet1.URL = u
    If Err.Number <> 0 Then L "  URL=" & u & ": ! " & Err.Number & " " & Err.Description: Err.Clear
    Props Inet1, "URL=" & u
End Sub

Private Sub Try(ByVal what As String)
    Dim v As Variant, o As Object
    On Error Resume Next
    Set o = Inet1
    Select Case what
    Case "Protocol=9": Inet1.Protocol = 9
    Case "Protocol=-1": Inet1.Protocol = -1
    Case "RemotePort=-1": Inet1.RemotePort = -1
    Case "RequestTimeout=-1": Inet1.RequestTimeout = -1
    Case "AccessType=3": Inet1.AccessType = 3
    Case "AccessType=2 (no proxy)": Inet1.AccessType = icNamedProxy
    Case "StillExecuting=": o.StillExecuting = True
    Case "ResponseCode=": o.ResponseCode = 5
    Case "ResponseInfo=": o.ResponseInfo = "x"
    Case "hInternet=": o.hInternet = 5
    Case "GetChunk (no request)": v = Inet1.GetChunk(10)
    Case "GetHeader (no request)": v = Inet1.GetHeader()
    Case "Cancel (no request)": Inet1.Cancel
    Case "Execute (bad operation)": Inet2.Execute gBase & "/hello.txt", "FOO"
    Case "Execute (CD on HTTP)": Inet2.Execute gBase & "/hello.txt", "CD sub"
    Case "Execute (no URL)": Inet2.URL = "": Inet2.Execute , "GET"
    Case "OpenURL (no URL)": Inet1.URL = "": v = Inet1.OpenURL()
    Case "OpenURL (bad DataType)": v = Inet1.OpenURL(gBase & "/hello.txt", 5)
    End Select
    If Err.Number <> 0 Then
        L "  " & what & ": ! " & Err.Number & " " & Err.Description & " [" & Err.Source & "]"
    Else
        L "  " & what & ": ok " & Hx(v, 60)
    End If
    Err.Clear
End Sub

Private Sub Open1(ByVal u As String, Optional ByVal bytes As Boolean)
    Dim v As Variant
    Drain
    On Error Resume Next
    If bytes Then v = Inet1.OpenURL(u, icByteArray) Else v = Inet1.OpenURL(u)
    If Err.Number <> 0 Then
        L "  OpenURL " & u & ": ! " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        L "  OpenURL " & u & ": " & Hx(v)
    End If
    WriteRaw "OpenURL " & u & ":" & gEv
    L "    events " & Sq(gEv) & " | ResponseCode " & Inet1.ResponseCode & " ResponseInfo " & Hx(Inet1.ResponseInfo, 200) & _
      " StillExecuting " & Inet1.StillExecuting
End Sub

' the events of the request before can still come: they are let in first
Private Sub Drain()
    Dim t As Single
    t = Timer
    Do
        DoEvents
    Loop While Timer - t < 0.3 And Timer >= t
    gEv = ""
End Sub

Private Function WaitDone() As Boolean
    Dim t As Single
    t = Timer
    Do
        DoEvents
        If Not Inet1.StillExecuting Then WaitDone = True: Exit Function
    Loop While Timer - t < 15 And Timer >= t
    Inet1.Cancel
    L "    (timeout: Cancel, StillExecuting " & Inet1.StillExecuting & ")"
End Function

' Execute, then the state, the headers and the data in chunks of the given size
Private Sub Ex(ByVal u As String, ByVal op As String, ByVal data As String, ByVal hdrs As String, _
               Optional ByVal chunk As Long = 10, Optional ByVal bytes As Boolean)
    Dim v As Variant, all As String, n As Long, t As Long, b() As Byte, sum As Long, i As Long, size0 As Long, sizes As String
    size0 = chunk
    Drain
    On Error Resume Next
    If hdrs <> "" Then
        Inet1.Execute u, op, data, hdrs
    ElseIf data <> "" Then
        Inet1.Execute u, op, data
    ElseIf op <> "" Then
        Inet1.Execute u, op
    Else
        Inet1.Execute u
    End If
    If Err.Number <> 0 Then
        L "  Execute " & u & " " & op & ": ! " & Err.Number & " " & Err.Description
        Err.Clear
        Exit Sub
    End If
    L "  Execute " & u & " " & op & ": at return StillExecuting " & Inet1.StillExecuting
    WaitDone
    WriteRaw "Execute " & u & " " & op & ":" & gEv
    L "    events " & Sq(gEv) & " | ResponseCode " & Inet1.ResponseCode & " ResponseInfo " & Hx(Inet1.ResponseInfo, 200)
    v = Inet1.GetHeader()
    If Err.Number <> 0 Then L "    GetHeader(): ! " & Err.Number & " " & Err.Description: Err.Clear Else L "    GetHeader(): " & Hx(v)
    v = Inet1.GetHeader("Content-Length")
    If Err.Number <> 0 Then L "    Content-Length: ! " & Err.Number: Err.Clear Else L "    Content-Length: " & Hx(v)
    v = Inet1.GetHeader("content-type")
    If Err.Number <> 0 Then L "    content-type: ! " & Err.Number: Err.Clear Else L "    content-type: " & Hx(v)
    v = Inet1.GetHeader("X-Missing")
    If Err.Number <> 0 Then L "    X-Missing: ! " & Err.Number: Err.Clear Else L "    X-Missing: " & Hx(v)
    Do
        chunk = size0
        If bytes Then v = Inet1.GetChunk(chunk, icByteArray) Else v = Inet1.GetChunk(chunk)
        If n < 3 Then sizes = sizes & " " & chunk
        If Err.Number <> 0 Then L "    GetChunk: ! " & Err.Number & " " & Err.Description: Err.Clear: Exit Do
        If bytes Then
            If Not IsArray(v) Then Exit Do
            If UBound(v) < LBound(v) Then Exit Do
            For i = LBound(v) To UBound(v)
                sum = (sum + v(i) * (((t + i - LBound(v)) Mod 7) + 1)) Mod 65521
            Next
            t = t + UBound(v) - LBound(v) + 1
        Else
            If Len(v) = 0 Then Exit Do
            all = all & v
            t = t + Len(v)
        End If
        n = n + 1
        If n > 10000 Then Exit Do
    Loop
    If bytes Then
        L "    GetChunk(" & size0 & ", icByteArray): " & n & " chunks, " & t & " bytes, sum " & sum & ", then " & Hx(v) & _
          " | Size after:" & sizes & " | last " & chunk
    Else
        L "    GetChunk(" & size0 & "): " & n & " chunks: " & Hx(all) & ", then " & Hx(v) & " | Size after:" & sizes & _
          " | last " & chunk
    End If
End Sub

Private Sub RunAll()
    Dim v As Variant
    gBase = "http://" & gHost & ":47120"
    L "== new controls"
    L "  TypeName " & TypeName(Inet1)
    Props Inet1, "Inet1"
    Props Inet2, "Inet2"

    L "== the URL and its parts"
    SetU gBase & "/hello.txt"
    SetU "ftp://vb:six@" & gHost & ":47121/sub/inner.txt"
    SetU "https://www.example.com/a/b.htm?x=1#frag"
    SetU "www.example.com/page.htm"
    SetU "mailto:someone@example.com"
    SetU "gopher://gopher.example.com/"
    SetU ""
    Inet1.URL = "http://h1/doc.htm"
    Inet1.RemoteHost = "h2"
    Props Inet1, "RemoteHost=h2"
    Inet1.RemotePort = 8080
    Props Inet1, "RemotePort=8080"
    Inet1.Document = "/other.htm"
    Props Inet1, "Document=/other.htm"
    Inet1.Protocol = icFTP
    Props Inet1, "Protocol=icFTP"
    Inet1.UserName = "u"
    Inet1.Password = "p"
    Props Inet1, "UserName, Password"
    Inet1.Protocol = icHTTPS
    Props Inet1, "Protocol=icHTTPS"
    For Each v In Array("Protocol=9", "Protocol=-1", "RemotePort=-1", "RequestTimeout=-1", "AccessType=3", _
                        "StillExecuting=", "ResponseCode=", "ResponseInfo=", "hInternet=", _
                        "GetChunk (no request)", "GetHeader (no request)", "Cancel (no request)")
        Try CStr(v)
    Next
    Inet1.URL = ""
    Inet1.UserName = ""
    Inet1.Password = ""
    Props Inet1, "cleared"

    L "== OpenURL"
    Open1 gBase & "/hello.txt"
    Props Inet1, "after OpenURL"
    Open1 gBase & "/bin", True
    Open1 gBase & "/bin"
    Open1 gBase & "/big", True
    Open1 gBase & "/empty"
    Open1 gBase & "/greek"
    Open1 gBase & "/echo"
    Open1 gBase & "/missing"
    Open1 gBase & "/error"
    Open1 gBase & "/redirect"
    Open1 "http://" & gHost & ":47129/refused"
    Open1 "http://no-such-host.invalid/"
    Open1 "xyz://" & gHost & "/"
    Inet1.URL = gBase & "/hello.txt"
    Open1 ""
    For Each v In Array("OpenURL (no URL)", "OpenURL (bad DataType)")
        Try CStr(v)
    Next

    L "== Execute"
    Ex gBase & "/hello.txt", "GET", "", ""
    Props Inet1, "after Execute"
    Ex gBase & "/hello.txt", "", "", "", 4
    Ex gBase & "/hello.txt", "HEAD", "", ""
    Ex gBase & "/bin", "GET", "", "", 100, True
    Ex gBase & "/big", "GET", "", "", 4096, True
    Ex gBase & "/echo", "POST", "name=vb&ver=6", "Content-Type: application/x-www-form-urlencoded", 1000
    Ex gBase & "/echo", "POST", "a=1", "X-One: 1" & vbCrLf & "X-Two: 2" & vbCrLf, 1000
    Ex gBase & "/echo", "POST", "", "", 1000
    Ex gBase & "/echo", "PUT", "put data", "", 1000
    Ex gBase & "/echo", "GET", "", "X-Mine: yes", 1000
    Ex gBase & "/missing", "GET", "", ""
    Ex gBase & "/error", "GET", "", ""
    Ex gBase & "/redirect", "GET", "", ""
    Ex "http://" & gHost & ":47129/refused", "GET", "", ""
    Ex "http://no-such-host.invalid/", "GET", "", ""
    Inet1.URL = gBase & "/empty"
    Ex "", "GET", "", ""

    L "== a slow request: Execute again, Cancel"
    Drain
    On Error Resume Next
    Inet1.Execute gBase & "/slow", "GET"
    L "  Execute /slow: ! " & Err.Number & " | StillExecuting " & Inet1.StillExecuting
    Err.Clear
    Inet1.Execute gBase & "/hello.txt", "GET"
    L "  Execute again: ! " & Err.Number & " " & Err.Description
    Err.Clear
    v = Inet1.OpenURL(gBase & "/hello.txt")
    L "  OpenURL meanwhile: ! " & Err.Number & " " & Err.Description & " | " & Hx(v)
    Err.Clear
    Inet1.Cancel
    L "  Cancel: ! " & Err.Number & " | StillExecuting " & Inet1.StillExecuting
    Err.Clear
    On Error GoTo 0
    WaitDone
    WriteRaw "slow + Cancel:" & gEv
    L "    events " & Sq(gEv) & " | ResponseCode " & Inet1.ResponseCode & " ResponseInfo " & Hx(Inet1.ResponseInfo, 200)
    Ex gBase & "/hello.txt", "GET", "", ""

    ' last: WinINet keeps the user and the password of a server for the rest of the process
    L "== authentication"
    Open1 gBase & "/auth"
    Ex gBase & "/auth", "GET", "", ""
    Inet1.UserName = "vb"
    Inet1.Password = "six"
    Props Inet1, "UserName, Password"
    Open1 gBase & "/auth"
    Ex gBase & "/auth", "GET", "", ""
    Inet1.UserName = ""
    Inet1.Password = ""
    Open1 "http://vb:six@" & gHost & ":47120/auth"
    Open1 gBase & "/auth"
    Ex gBase & "/echo", "GET", "", "", 1000
    Inet1.URL = "http://vb:wrong@" & gHost & ":47120/auth"
    Ex "", "GET", "", ""

    ' VB6's bugs with "CD" on HTTP and unknown verbs are not here: LinBASIC Studio does not have them
    ' (tests/dev/f_inetfix)
    L "end"
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    gEv = gEv & " " & State
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    RunAll
    Unload Me
End Sub
