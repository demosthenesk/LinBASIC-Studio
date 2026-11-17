VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1
   Caption         =   "MAPI 3"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2880
      Top             =   1200
   End
   Begin MSMAPI.MAPISession MAPISession1
      Left            =   240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1
      Left            =   1200
      Top             =   240
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
' MAPI 3 (LinBASIC Studio phase 7.9): the messages of the provider (lbmapi, its calls in the log): Save (a new one,
' the same again), two messages of another program (MIME, quoted-printable, base64, unread), Fetch (the unread ones,
' all, sorted, of a type), a message read (its recipients, its attachments as files, the flags, the date), MsgIndex,
' Reply / ReplyAll / Forward / Copy into the compose buffer, a copy saved over its message, Delete of a message, of a
' recipient and of an attachment of a message read, the errors

Private mSeen As String
Private mStep As Integer

Private Const MSG_READ = "MsgIndex MsgID MsgSubject MsgNoteText MsgOrigDisplayName MsgOrigAddress MsgDateReceived " & _
    "MsgType MsgConversationID MsgRead MsgSent MsgReceiptRequested MsgCount RecipCount AttachmentCount"
Private Const SESSION_PROPS = "Action DownLoadMail LogonUI NewSession Password UserName SessionID"
Private Const MSG_PROPS = "AddressCaption AddressEditFieldCount AddressLabel AddressModifiable AddressResolveUI " & _
    "AttachmentCount AttachmentIndex AttachmentName AttachmentPathName AttachmentPosition AttachmentType " & _
    "FetchMsgType FetchSorted FetchUnreadOnly MsgConversationID MsgCount MsgDateReceived MsgID MsgIndex " & _
    "MsgNoteText MsgOrigAddress MsgOrigDisplayName MsgRead MsgReceiptRequested MsgSent MsgSubject MsgType " & _
    "RecipAddress RecipCount RecipDisplayName RecipIndex RecipType SessionID Action"

' an error: its number, and the first time its source and text (taken at once: a function with an error handler
' that runs after it clears Err)
Private Function E(ByVal en As Long, ByVal es As String, ByVal ed As String) As String
    E = "!" & en
    If InStr(mSeen, "|" & en & "|") = 0 Then
        mSeen = mSeen & "|" & en & "|"
        WriteLog "    (" & en & " [" & es & "] " & ed & ")"
    End If
End Function

Private Sub Props(ByVal what As String, o As Object, ByVal names As String)
    Dim n As Variant, v As Variant, s As String, en As Long, es As String, ed As String
    s = what & ":"
    For Each n In Split(names, " ")
        v = Empty
        On Error Resume Next
        v = CallByName(o, CStr(n), VbGet)
        en = Err.Number: es = Err.Source: ed = Err.Description
        On Error GoTo 0
        If en <> 0 Then s = s & " " & n & "=" & E(en, es, ed) Else s = s & " " & n & "=" & ShowV(v)
    Next
    WriteLog s
End Sub

Private Sub SetP(o As Object, ByVal n As String, ByVal v As Variant)
    Dim en As Long, es As String, ed As String
    On Error Resume Next
    CallByName o, n, VbLet, v
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo 0
    If en <> 0 Then WriteLog "  " & n & " = " & ShowV(v) & " " & E(en, es, ed) Else WriteLog "  " & n & " = " & ShowV(v)
End Sub

Private Sub Act(o As Object, ByVal n As String, ParamArray a() As Variant)
    Dim en As Long, es As String, ed As String
    On Error Resume Next
    Select Case UBound(a)
    Case -1: CallByName o, n, VbMethod
    Case 0: CallByName o, n, VbMethod, a(0)
    End Select
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo 0
    If en <> 0 Then WriteLog "  " & n & " " & E(en, es, ed) Else WriteLog "  " & n & " ok"
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

' the message at MsgIndex: its properties, recipients and attachments
Private Sub ShowMsg(ByVal what As String)
    Dim i As Long, n As Long
    Props what, MAPIMessages1, MSG_READ
    On Error Resume Next
    n = MAPIMessages1.RecipCount
    For i = 0 To n - 1
        MAPIMessages1.RecipIndex = i
        Props "  recip " & i, MAPIMessages1, "RecipType RecipDisplayName RecipAddress"
    Next
    n = MAPIMessages1.AttachmentCount
    For i = 0 To n - 1
        MAPIMessages1.AttachmentIndex = i
        Props "  attachment " & i, MAPIMessages1, "AttachmentName AttachmentPathName AttachmentPosition AttachmentType"
        WriteLog "    text: " & FileText(MAPIMessages1.AttachmentPathName)
    Next
    If Err.Number <> 0 Then WriteLog "  ! " & Err.Number: Err.Clear
End Sub

Private Function FileText(ByVal path As String) As String
    Dim f As Integer, s As String
    On Error GoTo bad
    f = FreeFile
    Open path For Binary As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    FileText = Replace(Replace(s, vbCr, "\r"), vbLf, "\n")
    Exit Function
bad:
    FileText = "! " & Err.Number
End Function

Private Sub ShowAll(ByVal what As String)
    Dim i As Long
    WriteLog what & ": MsgCount=" & MAPIMessages1.MsgCount
    For i = 0 To MAPIMessages1.MsgCount - 1
        SetP MAPIMessages1, "MsgIndex", i
        ShowMsg "message " & i
    Next
End Sub

Private Sub NewMsg(ByVal subj As String, ByVal note As String)
    Act MAPIMessages1, "Compose"
    SetP MAPIMessages1, "MsgSubject", subj
    SetP MAPIMessages1, "MsgNoteText", note
    SetP MAPIMessages1, "RecipIndex", 0
    SetP MAPIMessages1, "RecipDisplayName", "Bob Smith"
    SetP MAPIMessages1, "RecipAddress", "SMTP:bob@example.com"
    SetP MAPIMessages1, "RecipType", 1
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    Select Case mStep
    Case 1
        WriteLog "-- signon; a message saved, saved again (two other messages are in the store, unread)"
        SetP MAPISession1, "LogonUI", False
        Act MAPISession1, "SignOn"
        SetP MAPIMessages1, "SessionID", MAPISession1.SessionID
        NewMsg "First", "Body one"
        SetP MAPIMessages1, "AttachmentIndex", 0
        SetP MAPIMessages1, "AttachmentPathName", App.Path & "\att1.txt"
        SetP MAPIMessages1, "AttachmentPosition", 2
        Act MAPIMessages1, "Save"
        Props "saved", MAPIMessages1, "MsgIndex MsgID MsgCount MsgSubject"
        SetP MAPIMessages1, "MsgSubject", "First (2)"
        Act MAPIMessages1, "Save"
        Props "saved again", MAPIMessages1, "MsgIndex MsgID MsgCount MsgSubject"
        Calls
    Case 2
        WriteLog "-- fetch: the unread ones, then all"
        SetP MAPIMessages1, "FetchUnreadOnly", True
        Act MAPIMessages1, "Fetch"
        Calls
        ShowAll "unread"
        Calls
        SetP MAPIMessages1, "FetchUnreadOnly", False
        Act MAPIMessages1, "Fetch"
        Calls
        ShowAll "all"
        Calls
    Case 3
        WriteLog "-- fetch: sorted, of a type, none"
        SetP MAPIMessages1, "FetchSorted", True
        Act MAPIMessages1, "Fetch"
        Props "sorted", MAPIMessages1, "MsgCount MsgIndex MsgID MsgSubject"
        SetP MAPIMessages1, "MsgIndex", 2
        Props "sorted 2", MAPIMessages1, "MsgIndex MsgID MsgSubject MsgRead"
        SetP MAPIMessages1, "MsgIndex", 3
        SetP MAPIMessages1, "FetchSorted", False
        SetP MAPIMessages1, "FetchMsgType", "IPM.Custom"
        Act MAPIMessages1, "Fetch"
        Props "custom", MAPIMessages1, "MsgCount MsgIndex MsgID MsgSubject MsgType"
        SetP MAPIMessages1, "FetchMsgType", "IPM.Nothing"
        Act MAPIMessages1, "Fetch"
        Props "nothing", MAPIMessages1, "MsgCount MsgIndex MsgID MsgSubject"
        SetP MAPIMessages1, "FetchMsgType", ""
        Calls
    Case 4
        WriteLog "-- reply, reply all, forward, copy; a message read cannot change"
        Act MAPIMessages1, "Fetch"
        SetP MAPIMessages1, "MsgIndex", 0
        Act MAPIMessages1, "Reply"
        ShowMsg "reply"
        SetP MAPIMessages1, "MsgIndex", 0
        Act MAPIMessages1, "ReplyAll"
        ShowMsg "reply all"
        SetP MAPIMessages1, "MsgIndex", 1
        Act MAPIMessages1, "Forward"
        ShowMsg "forward"
        SetP MAPIMessages1, "MsgIndex", 1
        Act MAPIMessages1, "Copy"
        ShowMsg "copy"
        SetP MAPIMessages1, "MsgIndex", 0
        SetP MAPIMessages1, "MsgSubject", "Changed"
        SetP MAPIMessages1, "RecipIndex", 0
        SetP MAPIMessages1, "RecipDisplayName", "Changed"
        Calls
    Case 5
        WriteLog "-- a copy changed and saved over its message; delete a recipient, an attachment, messages"
        SetP MAPIMessages1, "MsgIndex", 1
        Act MAPIMessages1, "Copy"
        SetP MAPIMessages1, "MsgSubject", "Custom one, changed"
        Act MAPIMessages1, "Save"
        Act MAPIMessages1, "Fetch"
        ShowAll "after save"
        SetP MAPIMessages1, "MsgIndex", 1
        Act MAPIMessages1, "Delete", 1
        Act MAPIMessages1, "Delete", 2
        Props "after delete recip", MAPIMessages1, "MsgIndex RecipCount AttachmentCount"
        Act MAPIMessages1, "Delete", 0
        Props "after delete", MAPIMessages1, "MsgCount MsgIndex MsgID MsgSubject"
        Act MAPIMessages1, "Delete"
        Props "after delete 2", MAPIMessages1, "MsgCount MsgIndex MsgID MsgSubject"
        Act MAPIMessages1, "Fetch"
        Props "fetched", MAPIMessages1, "MsgCount MsgIndex MsgID MsgSubject"
        SetP MAPIMessages1, "MsgIndex", 9
        Calls
    Case 6
        WriteLog "-- signoff"
        Act MAPISession1, "SignOff"
        Props "session", MAPISession1, SESSION_PROPS
        Calls
        WriteLog "end"
        End
    End Select
    Timer1.Enabled = True
End Sub
