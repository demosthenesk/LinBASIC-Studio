VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "AsyncRead"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   1200
   End
   Begin InetAsync.Loader Loader1 
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   1058
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The steps run one after the other from the Timer: each starts reads and waits for their AsyncReadComplete
' events (or for its time, then "timeout").
Private gStep As Long, gWait As Long, gDeadline As Single, gCancelSlow As Boolean

Private Function Url(ByVal p As String) As String
    Url = "http://" & gHost & ":47120" & p
End Function

Private Sub L(ByVal s As String)
    WriteLog s
End Sub

Private Sub Form_Load()
    L "start: events " & Loader1.Events
    gStep = 0
    Timer1.Enabled = True
End Sub

' wait for n AsyncReadComplete events, at most secs seconds
Private Sub WaitFor(ByVal n As Long, ByVal secs As Single)
    gWait = n
    gDeadline = Timer + secs
End Sub

Private Sub Loader1_Done(ByVal PropName As String)
    gWait = gWait - 1
End Sub

Private Sub Loader1_Progress(ByVal PropName As String, ByVal BytesRead As Long)
    If PropName = "Slow" And gCancelSlow Then
        gCancelSlow = False
        L "  (a progress event of Slow: CancelAsyncRead)"
        Loader1.Cancel "Slow"
    End If
End Sub

Private Sub Timer1_Timer()
    If gWait > 0 Then
        If Timer < gDeadline Then Exit Sub
        L "  timeout: " & gWait & " complete event(s) did not come"
        gWait = 0
    End If
    Timer1.Enabled = False
    gStep = gStep + 1
    RunStep gStep
    If gStep < 99 Then Timer1.Enabled = True
End Sub

Private Sub RunStep(ByVal n As Long)
    Select Case n
    Case 1
        L "1 local picture"
        WaitFor 1, 10
        Loader1.Start App.Path & "\pic.bmp", vbAsyncTypePicture, "Pic1"
    Case 2
        L "2 local file as a file"
        WaitFor 1, 10
        Loader1.Start App.Path & "\data.dat", vbAsyncTypeFile, "File1"
    Case 3
        L "3 local file as bytes"
        WaitFor 1, 10
        Loader1.Start App.Path & "\data.dat", vbAsyncTypeByteArray, "Bytes1"
    Case 4
        L "4 file:// URL as bytes"
        WaitFor 1, 10
        Loader1.Start "file:///" & Replace(App.Path, "\", "/") & "/data.dat", vbAsyncTypeByteArray, "Bytes2"
    Case 5
        L "5 a local file that is not there"
        WaitFor 1, 2
        Loader1.Start App.Path & "\missing.dat", vbAsyncTypeFile, "Miss1"
    Case 6
        L "6 no PropertyName"
        WaitFor 1, 10
        Loader1.Start App.Path & "\data.dat", vbAsyncTypeByteArray
    Case 7
        L "7 AsyncType 7"
        WaitFor 1, 2
        Loader1.Start App.Path & "\data.dat", 7, "Bad1"
    Case 8
        L "8 empty Target"
        WaitFor 1, 2
        Loader1.Start "", vbAsyncTypeFile, "Empty1"
    Case 9
        L "9 HTTP text as a file"
        WaitFor 1, 15
        Loader1.Start Url("/hello.txt"), vbAsyncTypeFile, "HFile"
    Case 10
        L "10 HTTP text as bytes, ForceUpdate"
        WaitFor 1, 15
        Loader1.Start Url("/hello.txt"), vbAsyncTypeByteArray, "HBytes", vbAsyncReadForceUpdate
    Case 11
        L "11 HTTP picture"
        WaitFor 1, 15
        Loader1.Start Url("/pic.bmp"), vbAsyncTypePicture, "HPic", vbAsyncReadForceUpdate
    Case 12
        L "12 HTTP 100000 bytes"
        WaitFor 1, 20
        Loader1.Start Url("/big"), vbAsyncTypeByteArray, "HBig", vbAsyncReadForceUpdate
    Case 13
        L "13 HTTP 404"
        WaitFor 1, 15
        Loader1.Start Url("/missing"), vbAsyncTypeByteArray, "H404", vbAsyncReadForceUpdate
    Case 14
        L "14 HTTP 500"
        WaitFor 1, 15
        Loader1.Start Url("/error"), vbAsyncTypeFile, "H500", vbAsyncReadForceUpdate
    Case 15
        L "15 HTTP refused (port 47199)"
        WaitFor 1, 30
        Loader1.Start "http://" & gHost & ":47199/x.txt", vbAsyncTypeByteArray, "HRefused"
    Case 16
        L "16 HTTP redirect"
        WaitFor 1, 15
        Loader1.Start Url("/redirect"), vbAsyncTypeByteArray, "HRedir", vbAsyncReadForceUpdate
    Case 17
        L "17 SynchronousDownload"
        WaitFor 1, 15
        Loader1.Start Url("/hello.txt?sync"), vbAsyncTypeByteArray, "HSync", vbAsyncReadSynchronousDownload Or vbAsyncReadForceUpdate
        L "  after the call"
    Case 18
        L "18 the same PropertyName twice"
        WaitFor 2, 15
        Loader1.Start Url("/slow?a"), vbAsyncTypeByteArray, "Dup", vbAsyncReadForceUpdate
        Loader1.Start Url("/slow?b"), vbAsyncTypeByteArray, "Dup", vbAsyncReadForceUpdate
    Case 19
        L "19 two reads at once"
        WaitFor 2, 15
        Loader1.Start Url("/slow?c"), vbAsyncTypeByteArray, "TwoA", vbAsyncReadForceUpdate
        Loader1.Start App.Path & "\data.dat", vbAsyncTypeByteArray, "TwoB"
    Case 20
        L "20 CancelAsyncRead in a progress event"
        WaitFor 1, 8
        gCancelSlow = True
        Loader1.Start Url("/slowbig"), vbAsyncTypeByteArray, "Slow", vbAsyncReadForceUpdate
    Case 21
        L "21 CancelAsyncRead right after AsyncRead"
        WaitFor 1, 6
        Loader1.Start Url("/slow?d"), vbAsyncTypeByteArray, "Slow2", vbAsyncReadForceUpdate
        Loader1.Cancel "Slow2"
    Case 22
        L "22 CancelAsyncRead of a name that is not read, without a name"
        WaitFor 0, 0
        Loader1.Cancel "Nothing"
        Loader1.Cancel
    Case 23
        L "23 CancelAsyncRead without a name, a read without a name"
        WaitFor 1, 6
        Loader1.Start Url("/slow?e"), vbAsyncTypeByteArray
        Loader1.Cancel
    Case 24
        L "24 Hyperlink"
        WaitFor 0, 0
        Loader1.Links True, Url("/hello.txt")
    Case 25
        L "end"
        gStep = 99
        Unload Me
    End Select
End Sub
