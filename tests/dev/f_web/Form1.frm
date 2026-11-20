VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "f_web"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   5400
      Top             =   3900
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      ExtentX         =   10186
      ExtentY         =   6588
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' f_web - the WebBrowser control (phase 9.2): the events that LinBASIC Studio gives as VB6 with Internet Explorer,
' their URL, Flags, Cancel, LocationURL / LocationName, Busy, ReadyState, Type, the Back / Forward commands, the
' error of GoForward without a page. The paths are written relative to App.Path (VB6 ran it from \\VBOXSVR\...).
Dim nStep As Integer, bBack As Boolean, bFwd As Boolean

Private Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\expected\f_web.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Private Function Norm(ByVal s As String) As String
    s = Replace(s, App.Path, "<app>", , , vbTextCompare)
    Norm = Replace(s, "\", "/")
End Function

Private Function Loc() As String
    Dim u As String
    u = WebBrowser1.LocationURL
    If InStrRev(u, "/") > 0 And Left$(u, 6) <> "about:" Then u = Mid$(u, InStrRev(u, "/") + 1)
    Loc = u
End Function

Private Function Ty() As String
    On Error Resume Next
    Ty = WebBrowser1.Type
    If Err.Number <> 0 Then Ty = "error " & Err.Number
End Function

Private Sub Form_Load()
    On Error Resume Next
    Kill App.Path & "\expected\f_web.txt"
    On Error GoTo 0
    Log "Load: rs=" & WebBrowser1.ReadyState & " busy=" & WebBrowser1.Busy & " loc=[" & Loc() & "] type=" & Ty()
    WebBrowser1.Navigate App.Path & "\page1.html"
    Log "after Navigate: busy=" & WebBrowser1.Busy & " rs=" & WebBrowser1.ReadyState
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    On Error GoTo bad
    nStep = nStep + 1
    Log "step " & nStep & ": back=" & bBack & " fwd=" & bFwd & " loc=" & Loc()
    Select Case nStep
    Case 1: WebBrowser1.Navigate App.Path & "\page2.html"
    Case 2: WebBrowser1.GoBack
    Case 3: WebBrowser1.GoForward
    Case 4: WebBrowser1.GoForward
    Case 5: WebBrowser1.Navigate App.Path & "\cancel.html"
    Case 6: WebBrowser1.Navigate "about:blank"
    Case 7
        Log "blank: loc=" & Loc() & " name=" & WebBrowser1.LocationName & " rs=" & WebBrowser1.ReadyState & " busy=" & WebBrowser1.Busy
        Timer1.Enabled = False
        Log "end"
        Unload Me
    End Select
    Exit Sub
bad:
    Log "error " & Err.Number & ": " & Err.Description
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Log "BeforeNavigate2 " & Norm(URL) & " flags=" & Flags & " target=[" & TargetFrameName & "]"
    If InStr(URL, "cancel") > 0 Then Cancel = True: Log "  cancelled"
End Sub

Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
    If Command = CSC_NAVIGATEBACK Then bBack = Enable
    If Command = CSC_NAVIGATEFORWARD Then bFwd = Enable
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Log "NavigateComplete2 " & Norm(URL) & " loc=" & Loc() & " busy=" & WebBrowser1.Busy
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Log "DocumentComplete " & Norm(URL) & " loc=" & Loc() & " name=" & WebBrowser1.LocationName & " busy=" & _
        WebBrowser1.Busy & " rs=" & WebBrowser1.ReadyState & " type=" & Ty()
End Sub

Private Sub WebBrowser1_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    Log "NavigateError " & Norm(URL) & " " & StatusCode
End Sub
