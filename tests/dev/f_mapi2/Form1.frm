VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1
   Caption         =   "MAPI 2"
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
' MAPI 2 (LinBASIC Studio phase 7.9): MAPISession and MAPIMessages with LinBASIC's provider of Simple MAPI (lbmapi,
' its calls in the log): SignOn / SignOff (a second SignOn, a wrong password, DownLoadMail, NewSession), Send (the
' recipients and attachments of the compose buffer as a MapiMessage, names resolved by the provider, a recipient
' unknown or ambiguous, an attachment not there, no recipients, the dialog sent or canceled), ResolveName, Show (the
' address book, the details of a recipient), the errors

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
    Timer1.Enabled = True
End Sub

Private Sub Compose2()
    Act MAPIMessages1, "Compose"
    SetP MAPIMessages1, "MsgSubject", "Hello"
    SetP MAPIMessages1, "MsgNoteText", "Line one" & vbCrLf & "Line two"
    SetP MAPIMessages1, "RecipIndex", 0
    SetP MAPIMessages1, "RecipDisplayName", "Bob Smith"
    SetP MAPIMessages1, "RecipType", 1
    SetP MAPIMessages1, "RecipIndex", 1
    SetP MAPIMessages1, "RecipDisplayName", "Dave"
    SetP MAPIMessages1, "RecipAddress", "SMTP:dave@example.com"
    SetP MAPIMessages1, "RecipType", 2
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    Select Case mStep
    Case 1
        WriteLog "-- signon: the password of the provider not given"
        Act MAPISession1, "SignOn"
        Props "session", MAPISession1, SESSION_PROPS
        Calls
        WriteLog "-- signon while signed on, signoff"
        Act MAPISession1, "SignOn"
        Act MAPISession1, "SignOff"
        Props "session", MAPISession1, SESSION_PROPS
        Calls
        WriteLog "-- signon: no UI, a wrong password, then the right one"
        SetP MAPISession1, "LogonUI", False
        SetP MAPISession1, "UserName", "tester"
        SetP MAPISession1, "Password", "wrong"
        Act MAPISession1, "SignOn"
        Props "session", MAPISession1, SESSION_PROPS
        SetP MAPISession1, "Password", "secret"
        Act MAPISession1, "SignOn"
        Props "session", MAPISession1, SESSION_PROPS
        Calls
        WriteLog "-- signoff twice, NewSession, DownLoadMail"
        Act MAPISession1, "SignOff"
        Props "session", MAPISession1, SESSION_PROPS
        Act MAPISession1, "SignOff"
        SetP MAPISession1, "NewSession", True
        SetP MAPISession1, "DownLoadMail", False
        Act MAPISession1, "SignOn"
        Props "session", MAPISession1, SESSION_PROPS
        Calls
    Case 2
        WriteLog "-- send: a name to resolve, an address, two attachments, a receipt"
        SetP MAPIMessages1, "SessionID", MAPISession1.SessionID
        Compose2
        SetP MAPIMessages1, "MsgReceiptRequested", True
        SetP MAPIMessages1, "AttachmentIndex", 0
        SetP MAPIMessages1, "AttachmentPathName", App.Path & "\att1.txt"
        SetP MAPIMessages1, "AttachmentPosition", 3
        SetP MAPIMessages1, "AttachmentIndex", 1
        SetP MAPIMessages1, "AttachmentPathName", App.Path & "\att2.dat"
        SetP MAPIMessages1, "AttachmentName", "data.bin"
        SetP MAPIMessages1, "AttachmentPosition", 0
        Act MAPIMessages1, "Send"
        Props "after send", MAPIMessages1, MSG_PROPS
        Calls
        WriteLog "-- send: an unknown recipient, an ambiguous one, an attachment not there, no recipients"
        Compose2
        SetP MAPIMessages1, "RecipIndex", 0
        SetP MAPIMessages1, "RecipDisplayName", "Nobody"
        Act MAPIMessages1, "Send"
        SetP MAPIMessages1, "RecipDisplayName", "Bob"
        Act MAPIMessages1, "Send"
        Props "after ambiguous", MAPIMessages1, "MsgIndex RecipCount RecipIndex RecipDisplayName RecipAddress RecipType"
        Compose2
        SetP MAPIMessages1, "AttachmentIndex", 0
        SetP MAPIMessages1, "AttachmentPathName", App.Path & "\nothere.txt"
        Act MAPIMessages1, "Send"
        Act MAPIMessages1, "Compose"
        SetP MAPIMessages1, "MsgSubject", "Nobody to"
        Act MAPIMessages1, "Send"
        Props "no recipients", MAPIMessages1, "MsgIndex RecipCount RecipIndex RecipDisplayName RecipAddress RecipType"
        Calls
    Case 3
        WriteLog "-- resolve names"
        Compose2
        SetP MAPIMessages1, "RecipIndex", 0
        SetP MAPIMessages1, "RecipDisplayName", "Carol"
        Act MAPIMessages1, "ResolveName"
        Props "carol", MAPIMessages1, "RecipCount RecipIndex RecipDisplayName RecipAddress RecipType"
        SetP MAPIMessages1, "RecipIndex", 1
        Act MAPIMessages1, "ResolveName"
        Props "dave", MAPIMessages1, "RecipCount RecipIndex RecipDisplayName RecipAddress RecipType"
        SetP MAPIMessages1, "RecipDisplayName", "Bob"
        Act MAPIMessages1, "ResolveName"
        Props "bob", MAPIMessages1, "RecipCount RecipIndex RecipDisplayName RecipAddress RecipType"
        SetP MAPIMessages1, "AddressResolveUI", True
        Act MAPIMessages1, "ResolveName"
        Props "bob with UI", MAPIMessages1, "RecipCount RecipIndex RecipDisplayName RecipAddress RecipType"
        SetP MAPIMessages1, "AddressResolveUI", False
        SetP MAPIMessages1, "RecipDisplayName", "x@example.org"
        SetP MAPIMessages1, "RecipAddress", ""
        Act MAPIMessages1, "ResolveName"
        Props "an address", MAPIMessages1, "RecipCount RecipIndex RecipDisplayName RecipAddress RecipType"
        Calls
    Case 4
        WriteLog "-- send with the dialog: sent, canceled"
        Compose2
        Act MAPIMessages1, "Send", True
        Props "after dialog", MAPIMessages1, "MsgIndex RecipCount MsgSubject"
        Setting "name=Tester" & vbCrLf & "dialog=cancel" & vbCrLf
        Compose2
        Act MAPIMessages1, "Send", True
        Calls
        WriteLog "-- the address book, the details of a recipient (canceled, then OK)"
        Compose2
        SetP MAPIMessages1, "AddressCaption", "Pick them"
        SetP MAPIMessages1, "AddressEditFieldCount", 3
        SetP MAPIMessages1, "AddressLabel", "Send to"
        SetP MAPIMessages1, "AddressModifiable", True
        Act MAPIMessages1, "Show"
        Setting "name=Tester" & vbCrLf
        Act MAPIMessages1, "Show"
        Props "after show", MAPIMessages1, "RecipCount RecipIndex RecipDisplayName RecipAddress RecipType"
        SetP MAPIMessages1, "RecipIndex", 0
        Act MAPIMessages1, "Show", True
        SetP MAPIMessages1, "AddressEditFieldCount", 0
        Act MAPIMessages1, "Show"
        SetP MAPIMessages1, "AddressEditFieldCount", 4
        Act MAPIMessages1, "Show"
        Props "after show 4", MAPIMessages1, "RecipCount RecipIndex RecipDisplayName RecipAddress RecipType"
        Calls
    Case 5
        WriteLog "-- signoff, send without a session"
        Act MAPISession1, "SignOff"
        Compose2
        Act MAPIMessages1, "Send"
        SetP MAPIMessages1, "SessionID", 0
        Act MAPIMessages1, "Send"
        Props "session", MAPISession1, SESSION_PROPS
        Calls
        WriteLog "end"
        End
    End Select
    Timer1.Enabled = True
End Sub
