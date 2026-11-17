VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Inet fixes"
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

' What LinBASIC Studio does on purpose not as VB6 (the bugs of MSINET.OCX, runtime/forms/ocx_inet.inc):
' no VB6 output; the expected output is the one of lbvc, checked by hand.
Private gBase As String

Private Sub L(ByVal s As String)
    WriteLog s
End Sub

Private Function Hx(v As Variant, Optional ByVal maxc As Long = 200) As String
    Dim i As Long, s As String, c As Long
    If IsArray(v) Then
        s = TypeName(v) & "(" & LBound(v) & " To " & UBound(v) & ")"
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

Private Sub WaitDone()
    Dim t As Single
    t = Timer
    Do
        DoEvents
    Loop While Inet1.StillExecuting And Timer - t < 15 And Timer >= t
End Sub

Private Function FirstLine(ByVal s As String) As String
    If InStr(s, vbCrLf) > 0 Then s = Left$(s, InStr(s, vbCrLf) - 1)
    FirstLine = s
End Function

Private Sub RunAll()
    Dim v As Variant, t As Single
    gBase = "http://" & gHost & ":47120"
    On Error Resume Next

    L "== an FTP command on HTTP: an error, the control is not busy after it"
    Inet1.Execute gBase & "/hello.txt", "CD sub"
    L "  CD sub: ! " & Err.Number & " " & Err.Description & " | StillExecuting " & Inet1.StillExecuting & _
      " ResponseCode " & Inet1.ResponseCode
    Err.Clear
    Inet1.Execute gBase & "/hello.txt", "GET"
    L "  then GET: ! " & Err.Number
    Err.Clear
    WaitDone
    L "  " & Hx(Inet1.GetChunk(100))

    L "== other HTTP verbs are sent"
    Inet1.Execute gBase & "/echo", "DELETE"
    WaitDone
    L "  DELETE: ! " & Err.Number & " | " & FirstLine(Inet1.GetHeader()) & " | " & FirstLine(Inet1.GetChunk(1000))
    Err.Clear
    Inet1.Execute gBase & "/hello.txt", "OPTIONS"
    WaitDone
    L "  OPTIONS: ! " & Err.Number & " | " & FirstLine(Inet1.GetHeader()) & " | StillExecuting " & Inet1.StillExecuting
    Err.Clear
    Inet1.Execute gBase & "/hello.txt", "BAD/VERB"
    L "  BAD/VERB: ! " & Err.Number & " " & Err.Description & " | StillExecuting " & Inet1.StillExecuting
    Err.Clear

    L "== GetChunk after Cancel returns at once"
    Inet1.Execute gBase & "/slow", "GET"
    Inet1.Cancel
    t = Timer
    v = Inet1.GetChunk(10)
    L "  GetChunk: " & Hx(v) & " | at once " & (Timer - t < 1) & " | StillExecuting " & Inet1.StillExecuting

    L "== OpenURL of ftp://"
    Inet1.RequestTimeout = 15
    v = Inet1.OpenURL("ftp://vb:six@" & gHost & ":47121/hello.txt")
    L "  hello.txt: ! " & Err.Number & " " & Hx(v)
    Err.Clear
    v = Inet1.OpenURL("ftp://vb:six@" & gHost & ":47121/bin.dat", icByteArray)
    L "  bin.dat: ! " & Err.Number & " " & Hx(v)
    Err.Clear
    v = Inet1.OpenURL("ftp://vb:six@" & gHost & ":47121/sub/")
    L "  sub/: ! " & Err.Number & " " & Hx(v)
    Err.Clear
    v = Inet1.OpenURL("ftp://vb:six@" & gHost & ":47121/nope.txt")
    L "  nope.txt: ! " & Err.Number & " " & Hx(v) & " | ResponseCode " & Inet1.ResponseCode & " " & Hx(Inet1.ResponseInfo)
    Err.Clear
    L "end"
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    RunAll
    Unload Me
End Sub
