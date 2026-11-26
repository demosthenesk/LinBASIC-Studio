VERSION 5.00
Begin VB.UserControl Loader 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ScaleHeight     =   600
   ScaleWidth      =   2400
   Begin VB.Label lblState 
      Caption         =   "Loader"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' AsyncRead and its events; what the AsyncProperty objects say goes to the log

Public Event Progress(ByVal PropName As String, ByVal BytesRead As Long)
Public Event Done(ByVal PropName As String)

Private mEvents As Long
Private mInfo As Collection         ' "c" & name: the status codes of its progress events (without repeats)
Private mLast As Collection         ' "l" & name: the last status code; "n" & name: the number of events

Private Function Getv(ByVal k As String) As String
    On Error Resume Next
    Getv = mInfo(k)
End Function
Private Sub Putv(ByVal k As String, ByVal v As String)
    On Error Resume Next
    mInfo.Remove k
    mInfo.Add v, k
End Sub

Private Sub UserControl_Initialize()
    Set mInfo = New Collection
    Set mLast = New Collection
End Sub

Public Property Get Events() As Long
    Events = mEvents
End Property

' AsyncRead with the arguments given; the log says what came back and how many events came during the call
Public Sub Start(ByVal Target As String, ByVal AType As Long, Optional PropName As Variant, Optional Options As Variant)
    Dim before As Long, nm As String
    before = mEvents
    If IsMissing(PropName) Then nm = "(none)" Else nm = PropName
    Putv "c" & nm, ""
    Putv "n" & nm, "0"
    On Error Resume Next
    If IsMissing(PropName) Then
        AsyncRead Target, AType
    ElseIf IsMissing(Options) Then
        AsyncRead Target, AType, PropName
    Else
        AsyncRead Target, AType, PropName, Options
    End If
    WriteLog "  AsyncRead " & nm & ": err " & Err.Number & IIf(Err.Number <> 0, " " & Err.Description, "") & ", events during the call " & (mEvents - before)
    Err.Clear
End Sub

Public Sub Cancel(Optional PropName As Variant)
    Dim before As Long
    before = mEvents
    On Error Resume Next
    If IsMissing(PropName) Then CancelAsyncRead Else CancelAsyncRead PropName
    WriteLog "  CancelAsyncRead " & IIf(IsMissing(PropName), "(none)", PropName) & ": err " & Err.Number & IIf(Err.Number <> 0, " " & Err.Description, "")
    WriteRaw "cancel: events during the call " & (mEvents - before)
    Err.Clear
End Sub

Public Sub Links(ByVal nav As Boolean, ByVal url As String)
    Dim h As Object
    On Error Resume Next
    Set h = UserControl.Hyperlink
    WriteLog "  Hyperlink: err " & Err.Number & " " & TypeName(h)
    Err.Clear
    UserControl.Hyperlink.GoBack
    WriteLog "  GoBack: err " & Err.Number & " " & Err.Description
    Err.Clear
    UserControl.Hyperlink.GoForward
    WriteLog "  GoForward: err " & Err.Number & " " & Err.Description
    Err.Clear
    If nav Then
        UserControl.Hyperlink.NavigateTo url
        WriteLog "  NavigateTo: err " & Err.Number & " " & Err.Description
        Err.Clear
        UserControl.Hyperlink.NavigateTo App.Path & "\data.dat"
        WriteLog "  NavigateTo file: err " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    UserControl.Hyperlink.NavigateTo ""
    WriteLog "  NavigateTo empty: err " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

Private Function Hx(v As Variant) As String
    Dim i As Long, s As String, c As Long, n As Long
    If IsArray(v) Then
        n = 0
        For i = LBound(v) To UBound(v)
            n = (n + v(i) * ((i Mod 7) + 1)) Mod 65521
        Next
        s = TypeName(v) & "(" & LBound(v) & " To " & UBound(v) & ") sum " & n
        If UBound(v) >= LBound(v) Then s = s & " first " & v(LBound(v)) & " last " & v(UBound(v))
    Else
        s = TypeName(v) & "(" & Len(v) & ") '"
        For i = 1 To Len(v)
            If i > 60 Then s = s & "...": Exit For
            c = AscW(Mid$(v, i, 1))
            If c < 32 Or c > 126 Then s = s & "<" & c & ">" Else s = s & Mid$(v, i, 1)
        Next
        s = s & "'"
    End If
    Hx = s
End Function

' a file that the read gives: is it the Target itself, how long, what is in it
Private Function FileText(ByVal p As String, ByVal tgt As String) As String
    Dim f As Integer, b() As Byte, n As Long, s As String
    On Error Resume Next
    If LCase$(p) = LCase$(tgt) Then s = "the Target itself" Else s = "another file"
    n = FileLen(p)
    If Err.Number <> 0 Then FileText = s & ", FileLen err " & Err.Number: Exit Function
    s = s & ", " & n & " bytes"
    If n > 0 Then
        f = FreeFile
        Open p For Binary Access Read As #f
        ReDim b(0 To n - 1)
        Get #f, , b
        Close #f
        s = s & ", " & Hx(b)
    End If
    FileText = s
End Function

Private Function Pv(ap As AsyncProperty, ByVal what As String) As String
    On Error Resume Next
    Dim v As Variant
    Select Case what
    Case "AsyncType": v = ap.AsyncType
    Case "PropertyName": v = "'" & ap.PropertyName & "'"
    Case "Target": v = "'" & ap.Target & "'"
    Case "BytesRead": v = ap.BytesRead
    Case "BytesMax": v = ap.BytesMax
    Case "Status": v = "'" & ap.Status & "'"
    Case "StatusCode": v = ap.StatusCode
    End Select
    If Err.Number <> 0 Then Pv = what & " err " & Err.Number Else Pv = what & " " & v
End Function

' the status codes without 1 and 2 (FindingResource, Connecting: they come or not, as the name cache of
' WinINet has it), unless there is nothing else
Private Function Codes(ByVal s As String) As String
    Dim a() As String, i As Long, r As String
    a = Split(s, " ")
    For i = 0 To UBound(a)
        If a(i) <> "1" And a(i) <> "2" And a(i) <> "" Then r = r & " " & a(i)
    Next
    If r = "" Then Codes = s Else Codes = Trim$(r)
End Function

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    Dim nm As String, c As Long
    mEvents = mEvents + 1
    nm = AsyncProp.PropertyName
    If nm = "" Then nm = "(none)"
    c = AsyncProp.StatusCode
    WriteRaw "progress " & nm & " " & c & " '" & AsyncProp.Status & "' " & AsyncProp.BytesRead & "/" & AsyncProp.BytesMax
    If Getv("l" & nm) <> CStr(c) Then Putv "c" & nm, Trim$(Getv("c" & nm) & " " & c)
    Putv "l" & nm, CStr(c)
    Putv "n" & nm, CStr(Val(Getv("n" & nm)) + 1)
    RaiseEvent Progress(nm, AsyncProp.BytesRead)
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    Dim nm As String, v As Variant, p As StdPicture, s As String
    mEvents = mEvents + 1
    nm = AsyncProp.PropertyName
    If nm = "" Then nm = "(none)"
    WriteLog "  complete " & nm & ": " & Pv(AsyncProp, "PropertyName") & ", " & Pv(AsyncProp, "AsyncType") & ", " & Pv(AsyncProp, "Target")
    WriteLog "    " & Pv(AsyncProp, "StatusCode") & ", " & Pv(AsyncProp, "Status") & ", " & Pv(AsyncProp, "BytesRead") & ", " & Pv(AsyncProp, "BytesMax")
    WriteLog "    progress events: " & IIf(Val(Getv("n" & nm)) > 0, "yes", "none") & ", codes " & Codes(Getv("c" & nm))
    WriteRaw "complete " & nm & " progress events " & Getv("n" & nm)
    On Error Resume Next
    If AsyncProp.AsyncType = vbAsyncTypePicture Then
        Set p = AsyncProp.Value
        If Err.Number <> 0 Then
            s = "Value err " & Err.Number & " " & Err.Description
        ElseIf p Is Nothing Then
            s = "Value Nothing"
        Else
            s = "Value " & TypeName(p) & " Type " & p.Type & " " & p.Width & " x " & p.Height
            Set UserControl.Picture = p
        End If
    Else
        v = AsyncProp.Value
        If Err.Number <> 0 Then
            s = "Value err " & Err.Number & " " & Err.Description
        ElseIf AsyncProp.AsyncType = vbAsyncTypeFile Then
            s = "Value " & TypeName(v) & ": " & FileText(v, AsyncProp.Target)
        Else
            s = "Value " & Hx(v)
        End If
    End If
    WriteLog "    " & s
    Err.Clear
    lblState.Caption = nm & " done"
    RaiseEvent Done(nm)
End Sub
