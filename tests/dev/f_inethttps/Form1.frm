VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Inet HTTPS"
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

' The Inet control with https:// (Linux: libcurl, Windows: WinINet) against tests/dev/inet_server.py: 47122 has a
' certificate the test trusts (VBRT_CA_FILE), 47123 one it does not. No VB6 output: the expected output is the one
' of lbvc, checked by hand (the same on Linux and Windows).
Private gEv As String, gS As String

Private Sub L(ByVal s As String)
    WriteLog s
End Sub

Private Function Hx(v As Variant, Optional ByVal maxc As Long = 200) As String
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

Private Function FirstLine(ByVal s As String) As String
    If InStr(s, vbCrLf) > 0 Then s = Left$(s, InStr(s, vbCrLf) - 1)
    FirstLine = s
End Function

Private Function LastLine(ByVal s As String) As String
    If Right$(s, 2) = vbCrLf Then s = Left$(s, Len(s) - 2)
    If InStrRev(s, vbCrLf) > 0 Then s = Mid$(s, InStrRev(s, vbCrLf) + 2)
    LastLine = s
End Function

Private Sub Drain()
    Dim t As Single
    t = Timer
    Do
        DoEvents
    Loop While Timer - t < 0.3 And Timer >= t
    gEv = ""
End Sub

Private Sub WaitDone()
    Dim t As Single
    t = Timer
    Do
        DoEvents
    Loop While Inet1.StillExecuting And Timer - t < 20 And Timer >= t
End Sub

Private Sub Open1(ByVal u As String, Optional ByVal bytes As Boolean)
    Dim v As Variant
    Drain
    On Error Resume Next
    If bytes Then v = Inet1.OpenURL(u, icByteArray) Else v = Inet1.OpenURL(u)
    L "  OpenURL " & u & ": ! " & Err.Number & " " & Hx(v)
    Err.Clear
    L "    events " & Sq(gEv) & " | ResponseCode " & Inet1.ResponseCode
End Sub

Private Sub Ex(ByVal u As String, ByVal op As String, ByVal data As String, ByVal hdrs As String)
    Dim v As Variant, all As String
    Drain
    On Error Resume Next
    If hdrs <> "" Then Inet1.Execute u, op, data, hdrs Else Inet1.Execute u, op
    If Err.Number <> 0 Then L "  Execute " & u & " " & op & ": ! " & Err.Number & " " & Err.Description: Err.Clear: Exit Sub
    WaitDone
    L "  Execute " & u & " " & op & ": events " & Sq(gEv) & " | ResponseCode " & Inet1.ResponseCode & " ResponseInfo " & _
      Hx(Inet1.ResponseInfo)
    ' the echo of a POST: its length is the one of the headers the library sends (WinINet of Windows / of Wine, libcurl)
    L "    status: " & FirstLine(Inet1.GetHeader()) & " | Content-Length " & IIf(op = "POST", "(echo)", Inet1.GetHeader("Content-Length"))
    Do
        v = Inet1.GetChunk(1000)
        If Len(v) = 0 Then Exit Do
        all = all & v
    Loop
    If op = "POST" Then
        L "    data: " & Hx(FirstLine(all)) & " ... " & Hx(LastLine(all))
    Else
        L "    data: " & Hx(all)
    End If
End Sub

Private Sub RunAll()
    gS = "https://" & gHost & ":47122"
    L "== the URL"
    Inet1.URL = gS & "/hello.txt"
    L "  Protocol " & Inet1.Protocol & " RemotePort " & Inet1.RemotePort & " Document " & Inet1.Document & " URL " & Inet1.URL
    Inet1.URL = "https://www.example.com/"
    L "  Protocol " & Inet1.Protocol & " RemotePort " & Inet1.RemotePort & " URL " & Inet1.URL

    L "== OpenURL"
    Open1 gS & "/hello.txt"
    Open1 gS & "/bin", True
    Open1 gS & "/redirect"
    Open1 gS & "/auth"
    Open1 "https://vb:six@" & gHost & ":47122/auth"
    Open1 gS & "/auth"
    Open1 "https://" & gHost & ":47123/hello.txt"
    Open1 "https://" & gHost & ":47124/hello.txt"
    Open1 "https://" & gHost & ":47125/hello.txt"

    L "== Execute"
    Ex gS & "/hello.txt", "GET", "", ""
    Ex gS & "/hello.txt", "HEAD", "", ""
    Ex gS & "/echo", "POST", "name=vb", "Content-Type: application/x-www-form-urlencoded"
    Ex gS & "/missing", "GET", "", ""
    Ex "https://" & gHost & ":47123/hello.txt", "GET", "", ""
    Ex "https://" & gHost & ":47124/hello.txt", "GET", "", ""
    Ex "https://" & gHost & ":47125/hello.txt", "GET", "", ""
    Ex "https://" & gHost & ":47129/refused", "GET", "", ""
    Ex "https://no-such-host.invalid/", "GET", "", ""
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
