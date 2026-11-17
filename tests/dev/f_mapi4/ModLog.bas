Attribute VB_Name = "ModLog"
Option Explicit
' MAPI 4 (LinBASIC Studio phase 7.9): the log of the test, the store of the provider of Simple MAPI (lbmapi: its calls)

Public Function LogPath() As String
    LogPath = App.Path & "\expected\f_mapi4.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

' the folder of the messages of the provider (LinBASIC's lbmapi: $LBMAPI_DIR, in the VM C:\LBMAPI\store)
Public Function Store() As String
    Store = Environ("LBMAPI_DIR")
    If Store = "" Then Store = "C:\LBMAPI\store"
End Function

Private Sub KillAll(ByVal pattern As String)
    On Error Resume Next
    Kill pattern
End Sub

Private Sub PutFile(ByVal path As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, text;
    Close #f
End Sub

' an empty store: no messages, the log of the calls on, the settings and the address book of the test
Public Sub NewStore()
    Dim s As String
    s = Store
    On Error Resume Next
    MkDir s
    On Error GoTo 0
    KillAll s & "\inbox\cur\*.*"
    KillAll s & "\inbox\new\*.*"
    KillAll s & "\sent\cur\*.*"
    KillAll s & "\tmp\*.*"
    PutFile s & "\calls.txt", ""
    PutFile s & "\provider.ini", "name=Tester" & vbCrLf & "address=SMTP:tester@example.com" & vbCrLf & _
        "now=2026/01/02 03:04" & vbCrLf
    PutFile s & "\addressbook.txt", "Bob Smith" & vbTab & "SMTP:bob@example.com" & vbCrLf & _
        "Bobby Brown" & vbTab & "SMTP:bobby@example.com" & vbCrLf & "Carol Jones" & vbTab & "SMTP:carol@example.com" & vbCrLf
End Sub

Public Sub Setting(ByVal text As String)
    PutFile Store & "\provider.ini", text
End Sub

' the paths of the test and of the store as names
Public Function Clean(ByVal s As String) As String
    Clean = Replace(Replace(s, App.Path, "<app>"), Store, "<store>")
End Function

' what the provider was asked since the last time
Public Sub Calls()
    Dim f As Integer, s As String
    f = FreeFile
    Open Store & "\calls.txt" For Input As #f
    Do Until EOF(f)
        Line Input #f, s
        WriteLog "    | " & Clean(s)
    Loop
    Close #f
    PutFile Store & "\calls.txt", ""
End Sub

Sub Main()
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    WriteLog "== start"
    NewStore
    PutFile App.Path & "\att1.txt", "first attachment" & vbCrLf
    On Error Resume Next
    MkDir Store & "\inbox"
    MkDir Store & "\inbox\new"
    MkDir Store & "\inbox\cur"
    On Error GoTo 0
    PutFile Store & "\inbox\new\000100", _
        "Date: Mon, 5 Jan 2026 10:20:30 +0200" & vbCrLf & _
        "From: Eve Adams <eve@example.com>" & vbCrLf & _
        "To: Tester <tester@example.com>" & vbCrLf & _
        "Subject: With a file" & vbCrLf & _
        "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: multipart/mixed; boundary=""XYZ""" & vbCrLf & vbCrLf & _
        "--XYZ" & vbCrLf & _
        "Content-Type: text/plain" & vbCrLf & vbCrLf & _
        "Text of one" & vbCrLf & _
        "--XYZ" & vbCrLf & _
        "Content-Type: text/plain; name=""note.txt""" & vbCrLf & _
        "Content-Transfer-Encoding: base64" & vbCrLf & _
        "Content-Disposition: attachment; filename=""note.txt""" & vbCrLf & vbCrLf & _
        "SGVsbG8gZnJvbSBFdmU=" & vbCrLf & _
        "--XYZ--" & vbCrLf
    PutFile Store & "\inbox\new\000101", _
        "Date: Tue, 6 Jan 2026 09:00:00 +0200" & vbCrLf & _
        "From: ""Bob Smith"" <bob@example.com>" & vbCrLf & _
        "To: Tester <tester@example.com>, ""Carol Jones"" <carol@example.com>" & vbCrLf & _
        "Subject: Second" & vbCrLf & _
        "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: multipart/mixed; boundary=""B2""" & vbCrLf & vbCrLf & _
        "--B2" & vbCrLf & _
        "Content-Type: text/plain" & vbCrLf & vbCrLf & _
        "Text of two" & vbCrLf & _
        "--B2" & vbCrLf & _
        "Content-Type: application/octet-stream; name=""c.txt""" & vbCrLf & _
        "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf & _
        "Q0ND" & vbCrLf & _
        "--B2--" & vbCrLf
    Form1.Show
End Sub

Public Function ShowV(v As Variant) As String
    On Error GoTo bad
    If IsObject(v) Then ShowV = "(" & TypeName(v) & ")": Exit Function
    If IsArray(v) Then ShowV = "(" & TypeName(v) & ")": Exit Function
    ShowV = TypeName(v) & ":" & Clean(CStr(v))
    Exit Function
bad:
    ShowV = TypeName(v) & ":?" & Err.Number
End Function
