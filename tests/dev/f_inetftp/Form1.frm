VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1
   Caption         =   "Inet FTP"
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The Inet control (MSINET.OCX) with FTP, against tests/dev/inet_server.py (port 47121, user vb, password six):
' Execute with the FTP commands, GetChunk of their answers, OpenURL of files and folders, the errors.
' The local files of PUT and GET are in the folder of the program (the current folder).
Private gEv As String, gBase As String

Private Sub L(ByVal s As String)
    WriteLog s
End Sub

Private Function Hx(v As Variant, Optional ByVal maxc As Long = 600) As String
    Dim i As Long, s As String, c As Long, n As Long
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            n = (n + v(i) * ((i Mod 7) + 1)) Mod 65521
        Next
        s = TypeName(v) & "(" & LBound(v) & " To " & UBound(v) & ") sum " & n
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
    Loop While Timer - t < 20 And Timer >= t
    Inet1.Cancel
    L "    (timeout: Cancel, StillExecuting " & Inet1.StillExecuting & ")"
End Function

Private Function ReadFile(ByVal p As String) As String
    Dim f As Integer, s As String
    On Error Resume Next
    If Dir(p) = "" Then ReadFile = "(no file)": Exit Function
    f = FreeFile
    Open p For Binary As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    ReadFile = Hx(s)
End Function

' Execute an FTP command (on the connection the control has, or a new one for u), then what it gives
Private Sub Fx(ByVal u As String, ByVal op As String)
    Dim v As Variant, all As String, n As Long, ok As Boolean
    Drain
    On Error Resume Next
    If u <> "" Then Inet1.Execute u, op Else Inet1.Execute , op
    If Err.Number <> 0 Then
        L "  " & op & ": ! " & Err.Number & " " & Err.Description
        Err.Clear
        Exit Sub
    End If
    ok = WaitDone()
    WriteRaw op & ":" & gEv
    L "  " & op & ": events " & Sq(gEv) & " | ResponseCode " & Inet1.ResponseCode & " ResponseInfo " & Hx(Inet1.ResponseInfo, 300)
    If Not ok Then Exit Sub             ' VB6: GetChunk after a request that was cancelled waits for ever
    Do
        v = Inet1.GetChunk(1024)
        If Err.Number <> 0 Then L "    GetChunk: ! " & Err.Number & " " & Err.Description: Err.Clear: Exit Do
        If Len(v) = 0 Then Exit Do
        all = all & v
        n = n + 1
        If n > 1000 Then Exit Do
    Loop
    If n > 0 Then L "    GetChunk: " & Hx(all)
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
    L "    events " & Sq(gEv) & " | ResponseCode " & Inet1.ResponseCode & " ResponseInfo " & Hx(Inet1.ResponseInfo, 300)
End Sub

Private Sub RunAll()
    Dim f As Integer
    gBase = "ftp://vb:six@" & gHost & ":47121"
    On Error Resume Next
    ChDrive App.Path
    ChDir App.Path
    Kill "local_get.txt"
    Kill "local_bin.dat"
    On Error GoTo 0
    f = FreeFile
    Open "local_put.txt" For Output As #f
    Print #f, "uploaded by the test"
    Close #f

    L "== the session"
    Inet1.RequestTimeout = 15
    Inet1.URL = gBase
    L "  URL " & Inet1.URL & " | Protocol " & Inet1.Protocol & " RemoteHost " & Inet1.RemoteHost & " RemotePort " & _
      Inet1.RemotePort & " UserName " & Inet1.UserName & " Password " & Inet1.Password
    Fx gBase, "PWD"
    Fx "", "DIR"
    Fx "", "LS"
    Fx "", "DIR sub"
    Fx "", "DIR hello.txt"
    Fx "", "CD sub"
    Fx "", "PWD"
    Fx "", "DIR"
    Fx "", "CDUP"
    Fx "", "PWD"
    Fx "", "SIZE hello.txt"
    Fx "", "SIZE big.dat"
    Fx "", "MKDIR newdir"
    Fx "", "DIR"
    Fx "", "RMDIR newdir"
    Fx "", "PUT local_put.txt up.txt"
    Fx "", "DIR up.txt"
    Fx "", "RENAME up.txt up2.txt"
    Fx "", "GET up2.txt local_get.txt"
    L "    local_get.txt: " & ReadFile("local_get.txt")
    Fx "", "RECV bin.dat local_bin.dat"
    L "    local_bin.dat: " & ReadFile("local_bin.dat")
    Fx "", "GET hello.txt local_get.txt"
    L "    local_get.txt again: " & ReadFile("local_get.txt")
    Fx "", "SEND local_put.txt up3.txt"
    Fx "", "DELETE up2.txt"
    Fx "", "DELETE up3.txt"
    Fx "", "DIR"
    L "== errors"
    Fx "", "CD nope"
    Fx "", "PWD"
    Fx "", "GET nope.txt local_nope.txt"
    L "    local_nope.txt: " & ReadFile("local_nope.txt")
    Fx "", "SIZE nope.txt"
    Fx "", "DELETE nope.txt"
    Fx "", "RMDIR nope"
    Fx "", "PUT no_such_local.txt x.txt"
    Fx "", "FOO bar"
    Fx "", "GET"
    Fx "", "CLOSE"
    Fx "", "PWD"
    Fx "", "QUIT"
    Fx "ftp://vb:wrong@" & gHost & ":47121", "DIR"
    Fx "", "CLOSE"
    Fx "ftp://" & gHost & ":47121", "PWD"
    Fx "", "CLOSE"
    Fx "ftp://" & gHost & ":47129", "DIR"
    Fx "", "CLOSE"

    ' OpenURL of ftp:// is not here: VB6 waits for the time of RequestTimeout and fails, or ends (oracle)
    On Error Resume Next
    Kill "local_put.txt"
    Kill "local_get.txt"
    Kill "local_bin.dat"
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
