VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1
   Caption         =   "MAPI 1"
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
' MAPI 1 (LinBASIC Studio phase 7.9): MAPISession and MAPIMessages without a mail program of Simple MAPI on the
' computer: the properties and their defaults, SignOn / SignOff, the methods without a session and with one that is
' not, the compose buffer (recipients, attachments), the indexes, Action, the errors

Private mSeen As String
Private mStep As Integer

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
    Props "session", MAPISession1, SESSION_PROPS
    Props "messages", MAPIMessages1, MSG_PROPS
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    Select Case mStep
    Case 1
        WriteLog "-- signon (defaults)"
        Act MAPISession1, "SignOn"
        Props "session", MAPISession1, SESSION_PROPS
        Act MAPISession1, "SignOff"
        SetP MAPISession1, "Action", 1
        SetP MAPISession1, "Action", 2
        SetP MAPISession1, "Action", 3
        SetP MAPISession1, "Action", 0
        WriteLog "-- signon (no UI, a user)"
        SetP MAPISession1, "LogonUI", False
        SetP MAPISession1, "DownLoadMail", False
        SetP MAPISession1, "NewSession", True
        SetP MAPISession1, "UserName", "tester"
        SetP MAPISession1, "Password", "secret"
        Act MAPISession1, "SignOn"
        Props "session", MAPISession1, SESSION_PROPS
        SetP MAPISession1, "SessionID", 5
    Case 2
        WriteLog "-- messages without a session"
        Act MAPIMessages1, "Compose"
        Act MAPIMessages1, "Fetch"
        Act MAPIMessages1, "Send"
        Act MAPIMessages1, "Save"
        Act MAPIMessages1, "Copy"
        Act MAPIMessages1, "Delete"
        Act MAPIMessages1, "Forward"
        Act MAPIMessages1, "Reply"
        Act MAPIMessages1, "ReplyAll"
        Act MAPIMessages1, "ResolveName"
        SetP MAPIMessages1, "MsgIndex", -1
        SetP MAPIMessages1, "MsgIndex", 0
        SetP MAPIMessages1, "RecipIndex", 0
        SetP MAPIMessages1, "AttachmentIndex", 0
        SetP MAPIMessages1, "MsgSubject", "Hello"
        SetP MAPIMessages1, "RecipDisplayName", "Bob"
        Props "messages", MAPIMessages1, MSG_PROPS
    Case 3
        WriteLog "-- messages with a session that is not one"
        SetP MAPIMessages1, "SessionID", 1234
        Act MAPIMessages1, "Compose"
        Props "messages", MAPIMessages1, MSG_PROPS
        SetP MAPIMessages1, "MsgSubject", "Hello"
        SetP MAPIMessages1, "MsgNoteText", "Body"
        SetP MAPIMessages1, "MsgType", "IPM.Note"
        SetP MAPIMessages1, "MsgReceiptRequested", True
        SetP MAPIMessages1, "RecipIndex", 0
        SetP MAPIMessages1, "RecipDisplayName", "Bob"
        SetP MAPIMessages1, "RecipAddress", "SMTP:bob@example.com"
        SetP MAPIMessages1, "RecipType", 1
        SetP MAPIMessages1, "RecipIndex", 1
        SetP MAPIMessages1, "RecipDisplayName", "Carol"
        SetP MAPIMessages1, "RecipType", 2
        SetP MAPIMessages1, "RecipIndex", 3
        SetP MAPIMessages1, "RecipIndex", -1
        SetP MAPIMessages1, "RecipType", 7
        SetP MAPIMessages1, "AttachmentIndex", 0
        SetP MAPIMessages1, "AttachmentPathName", "C:\Temp\ModLog.bas"
        SetP MAPIMessages1, "AttachmentName", "log.bas"
        SetP MAPIMessages1, "AttachmentPosition", 0
        SetP MAPIMessages1, "AttachmentType", 0
        SetP MAPIMessages1, "AttachmentType", 5
        SetP MAPIMessages1, "AttachmentIndex", 2
        Props "messages", MAPIMessages1, MSG_PROPS
        SetP MAPIMessages1, "RecipIndex", 0
        Props "recip 0", MAPIMessages1, "RecipIndex RecipDisplayName RecipAddress RecipType RecipCount"
        Act MAPIMessages1, "Delete", 1
        Props "after delete recip", MAPIMessages1, "RecipIndex RecipDisplayName RecipType RecipCount"
        Act MAPIMessages1, "Delete", 2
        Props "after delete attachment", MAPIMessages1, "AttachmentIndex AttachmentName AttachmentCount"
        Act MAPIMessages1, "Delete", 3
        Act MAPIMessages1, "Send"
        Act MAPIMessages1, "ResolveName"
        Act MAPIMessages1, "Fetch"
        Act MAPIMessages1, "Save"
        Act MAPIMessages1, "Delete", 0
        Props "messages", MAPIMessages1, MSG_PROPS
    Case 4
        WriteLog "-- the other properties"
        SetP MAPIMessages1, "AddressEditFieldCount", 2
        SetP MAPIMessages1, "AddressEditFieldCount", 5
        SetP MAPIMessages1, "AddressCaption", "Pick"
        SetP MAPIMessages1, "AddressLabel", "To:"
        SetP MAPIMessages1, "AddressModifiable", True
        SetP MAPIMessages1, "AddressResolveUI", True
        SetP MAPIMessages1, "FetchMsgType", "IPM.Note"
        SetP MAPIMessages1, "FetchSorted", True
        SetP MAPIMessages1, "FetchUnreadOnly", True
        SetP MAPIMessages1, "MsgCount", 3
        SetP MAPIMessages1, "MsgID", "x"
        SetP MAPIMessages1, "MsgOrigAddress", "a@b"
        SetP MAPIMessages1, "MsgRead", True
        SetP MAPIMessages1, "MsgDateReceived", "2026/01/01"
        SetP MAPIMessages1, "RecipCount", 4
        SetP MAPIMessages1, "AttachmentCount", 4
        SetP MAPIMessages1, "Action", 6
        SetP MAPIMessages1, "Action", 99
        SetP MAPIMessages1, "SessionID", 0
        Act MAPIMessages1, "Compose"
        Props "messages", MAPIMessages1, MSG_PROPS
        Props "session", MAPISession1, SESSION_PROPS
    Case 5
        WriteLog "end"
        End
    End Select
    Timer1.Enabled = True
End Sub
