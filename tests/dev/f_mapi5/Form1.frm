VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.Ocx"
Begin VB.Form Form1 
   Caption         =   "MAPI 5"
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
      DownloadMail    =   0   'False
      LogonUI         =   0   'False
      NewSession      =   -1  'True
      Password        =   "secret"
      UserName        =   "tester"
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1200
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   3
      AddressModifiable=   -1  'True
      AddressResolveUI=   -1  'True
      FetchSorted     =   -1  'True
      FetchUnreadOnly =   -1  'True
      AddressCaption  =   "Pick"
      AddressLabel    =   "To:"
      FetchMsgType    =   "IPM.Note"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' MAPI 5 (LinBASIC Studio phase 7.9): how the IDE writes MAPISession and MAPIMessages in the .frm (their properties
' changed in the Properties window, then saved), and the properties the program gets from it

Private Const SESSION_PROPS = "DownLoadMail LogonUI NewSession Password UserName SessionID"
Private Const MSG_PROPS = "AddressCaption AddressEditFieldCount AddressLabel AddressModifiable AddressResolveUI " & _
    "FetchMsgType FetchSorted FetchUnreadOnly SessionID MsgIndex"

Private Sub Props(ByVal what As String, o As Object, ByVal names As String)
    Dim n As Variant, v As Variant, s As String
    s = what & ":"
    For Each n In Split(names, " ")
        On Error Resume Next
        v = Empty
        v = CallByName(o, CStr(n), VbGet)
        If Err.Number <> 0 Then s = s & " " & n & "=!" & Err.Number: Err.Clear Else s = s & " " & n & "=" & ShowV(v)
        On Error GoTo 0
    Next
    WriteLog s
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Props "session", MAPISession1, SESSION_PROPS
    Props "messages", MAPIMessages1, MSG_PROPS
    WriteLog "end"
    End
End Sub
