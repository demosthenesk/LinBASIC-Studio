VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1
   Caption         =   "MAPI 4"
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
' MAPI 4 (LinBASIC Studio phase 7.9): when the MAPIMessages control asks the provider for a message (the calls
' after each step): Fetch, MsgIndex, MsgID, MsgRead before the text, the attachments before the text, the text, Reply
' after the text; the compose buffer across Fetch; Send / Save / Delete on a message read; the Action properties of
' MAPISession and MAPIMessages with a provider

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
    Timer1.Enabled = True
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

' a step and what the provider was asked for it
Private Sub P(ByVal what As String, ByVal names As String)
    Props what, MAPIMessages1, names
    Calls
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mStep = mStep + 1
    Select Case mStep
    Case 1
        WriteLog "-- signon by Action; fetch, then one property at a time (two unread messages with a file each)"
        SetP MAPISession1, "LogonUI", False
        SetP MAPISession1, "Action", 1
        Props "session", MAPISession1, SESSION_PROPS
        SetP MAPIMessages1, "SessionID", MAPISession1.SessionID
        Calls
        Act MAPIMessages1, "Fetch"
        Calls
        P "count", "MsgCount"
        SetP MAPIMessages1, "MsgIndex", 0
        Calls
        P "id", "MsgID"
        P "read", "MsgRead"
        P "subject", "MsgSubject MsgOrigDisplayName RecipCount"
        P "attachments", "AttachmentCount"
        SetP MAPIMessages1, "AttachmentIndex", 0
        P "attachment", "AttachmentName AttachmentPathName AttachmentPosition"
        WriteLog "    text: " & FileText(MAPIMessages1.AttachmentPathName)
        P "note", "MsgNoteText MsgRead AttachmentCount"
    Case 2
        WriteLog "-- the second message: the text first, then the attachments, Reply after the text"
        SetP MAPIMessages1, "MsgIndex", 1
        Calls
        P "note", "MsgNoteText"
        P "attachments", "AttachmentCount MsgRead"
        Act MAPIMessages1, "Reply"
        P "reply", "MsgIndex MsgSubject RecipCount AttachmentCount"
        SetP MAPIMessages1, "MsgIndex", 1
        P "message 1 again", "MsgSubject AttachmentCount"
        SetP MAPIMessages1, "MsgIndex", 0
        Act MAPIMessages1, "Copy"
        P "copy of 0", "MsgIndex MsgSubject MsgOrigDisplayName MsgOrigAddress AttachmentCount"
    Case 3
        WriteLog "-- the compose buffer across Fetch; Send, Save, Delete on a message read and on the buffer"
        Act MAPIMessages1, "Compose"
        SetP MAPIMessages1, "MsgSubject", "Kept"
        Act MAPIMessages1, "Fetch"
        P "after fetch", "MsgIndex MsgCount MsgSubject"
        SetP MAPIMessages1, "MsgIndex", -1
        P "buffer", "MsgIndex MsgSubject RecipCount"
        SetP MAPIMessages1, "MsgIndex", 0
        Act MAPIMessages1, "Send"
        Act MAPIMessages1, "Save"
        Act MAPIMessages1, "Forward"
        P "forward", "MsgIndex MsgSubject"
        SetP MAPIMessages1, "MsgIndex", -1
        Act MAPIMessages1, "Delete"
        P "buffer after delete", "MsgIndex MsgCount MsgSubject"
    Case 4
        WriteLog "-- the Action of MAPIMessages"
        SetP MAPIMessages1, "Action", 6
        P "action 6", "MsgIndex MsgSubject RecipCount"
        SetP MAPIMessages1, "MsgSubject", "By action"
        SetP MAPIMessages1, "RecipIndex", 0
        SetP MAPIMessages1, "RecipDisplayName", "Bob Smith"
        SetP MAPIMessages1, "RecipAddress", "SMTP:bob@example.com"
        SetP MAPIMessages1, "RecipIndex", 1
        SetP MAPIMessages1, "RecipDisplayName", "Carol Jones"
        SetP MAPIMessages1, "RecipAddress", "SMTP:carol@example.com"
        SetP MAPIMessages1, "AttachmentIndex", 0
        SetP MAPIMessages1, "AttachmentPathName", App.Path & "\att1.txt"
        SetP MAPIMessages1, "RecipIndex", 1
        SetP MAPIMessages1, "Action", 14
        P "action 14", "RecipCount RecipIndex RecipDisplayName AttachmentCount"
        SetP MAPIMessages1, "Action", 15
        P "action 15", "RecipCount AttachmentCount AttachmentIndex"
        SetP MAPIMessages1, "Action", 4
        P "action 4", "MsgIndex MsgID MsgCount"
        SetP MAPIMessages1, "Action", 1
        P "action 1", "MsgCount"
        SetP MAPIMessages1, "MsgIndex", 0
        SetP MAPIMessages1, "Action", 7
        P "action 7", "MsgIndex MsgSubject RecipCount"
        SetP MAPIMessages1, "MsgIndex", 0
        SetP MAPIMessages1, "Action", 8
        P "action 8", "MsgIndex MsgSubject RecipCount"
        SetP MAPIMessages1, "MsgIndex", 0
        SetP MAPIMessages1, "Action", 9
        P "action 9", "MsgIndex MsgSubject RecipCount"
        SetP MAPIMessages1, "MsgIndex", 0
        SetP MAPIMessages1, "Action", 5
        P "action 5", "MsgIndex MsgSubject MsgID"
        SetP MAPIMessages1, "MsgIndex", 2
        P "message 2", "MsgID MsgSubject"
        SetP MAPIMessages1, "Action", 10
        P "action 10", "MsgCount MsgIndex MsgID"
        SetP MAPIMessages1, "Action", 13
        SetP MAPIMessages1, "Action", 16
        SetP MAPIMessages1, "Action", 0
        Calls
    Case 5
        WriteLog "-- signoff by Action"
        SetP MAPISession1, "Action", 2
        Props "session", MAPISession1, SESSION_PROPS
        Calls
        WriteLog "end"
        End
    End Select
    Timer1.Enabled = True
End Sub
