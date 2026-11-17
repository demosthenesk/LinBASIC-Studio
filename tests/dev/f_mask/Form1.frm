VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1
   Caption         =   "MaskEdBox"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Interval        =   300
      Left            =   4200
      Top             =   1800
   End
   Begin VB.TextBox Text1
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin MSMask.MaskEdBox MaskEdBox1
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "(###) ###-####"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The Masked Edit control (MSMASK32.OCX): its properties with and without a Mask, what Text, ClipText and
' FormattedText give, the errors, the Format property, and typing (SendKeys): ValidationError, Change,
' the literals that the caret jumps over, AutoTab. The form must stay the active window while it runs.
Private gEv As String

Private Sub L(ByVal s As String)
    WriteLog s
End Sub

' a string with its quotes, the characters out of 32..126 as <n>
Private Function Q(ByVal s As String) As String
    Dim i As Long, c As Long, r As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 32 Or c > 126 Then r = r & "<" & c & ">" Else r = r & Mid$(s, i, 1)
    Next
    Q = "'" & r & "'"
End Function

' the state of MaskEdBox1
Private Function St() As String
    Dim s As String
    On Error Resume Next
    With MaskEdBox1
        s = "Text=" & Q(.Text)
        If Err Then s = "Text=!" & Err.Number: Err.Clear
        s = s & " Clip=" & Q(.ClipText)
        If Err Then s = s & " Clip=!" & Err.Number: Err.Clear
        s = s & " Fmt=" & Q(.FormattedText)
        If Err Then s = s & " Fmt=!" & Err.Number: Err.Clear
        s = s & " Max=" & .MaxLength
        If Err Then s = s & " Max=!" & Err.Number: Err.Clear
        s = s & " Sel=" & .SelStart & "," & .SelLength
        If Err Then s = s & " Sel=!" & Err.Number: Err.Clear
    End With
    St = s
End Function

Private Function Ev() As String
    If gEv <> "" Then Ev = " ev[" & Trim$(gEv) & "]"
    gEv = ""
End Function

' a property of MaskEdBox1 set, the error, the state after it
Private Sub SetP(ByVal what As String, ByVal v As Variant)
    Dim e As String
    gEv = ""
    On Error Resume Next
    Select Case what
    Case "Mask": MaskEdBox1.Mask = v
    Case "Text": MaskEdBox1.Text = v
    Case "ClipText": MaskEdBox1.ClipText = v
    Case "FormattedText": MaskEdBox1.FormattedText = v
    Case "PromptChar": MaskEdBox1.PromptChar = v
    Case "PromptInclude": MaskEdBox1.PromptInclude = v
    Case "AllowPrompt": MaskEdBox1.AllowPrompt = v
    Case "ClipMode": MaskEdBox1.ClipMode = v
    Case "MaxLength": MaskEdBox1.MaxLength = v
    Case "Format": MaskEdBox1.Format = v
    Case "SelStart": MaskEdBox1.SelStart = v
    Case "SelLength": MaskEdBox1.SelLength = v
    Case "SelText": MaskEdBox1.SelText = v
    Case "AutoTab": MaskEdBox1.AutoTab = v
    Case "Default": MaskEdBox1 = v
    End Select
    If Err Then e = " ! " & Err.Number & " " & Err.Description: Err.Clear
    L "  " & what & " = " & Q(CStr(v)) & e & " -> " & St() & Ev()
End Sub

' a clean start: no mask, no text
Private Sub Clean(ByVal m As String)
    On Error Resume Next
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = ""
    MaskEdBox1.PromptInclude = True
    MaskEdBox1.AllowPrompt = False
    MaskEdBox1.ClipMode = 0
    MaskEdBox1.PromptChar = "_"
    MaskEdBox1.Format = ""
    MaskEdBox1.MaxLength = 64
    MaskEdBox1.Mask = m
    If Err Then L "  (clean " & Q(m) & ": ! " & Err.Number & ")": Err.Clear
    gEv = ""
End Sub

' keys typed into the control that has the focus
Private Sub Ty(ByVal k As String)
    gEv = ""
    SendKeys k, True
    DoEvents
    L "  keys " & k & " -> " & St() & Ev()
End Sub

Private Sub Focus1()
    MaskEdBox1.SetFocus
    DoEvents
    gEv = ""
End Sub

Private Sub RunAll()
    Dim m As Variant, v As Variant
    On Error Resume Next

    L "== defaults"
    With MaskEdBox1
        L "  TypeName " & TypeName(MaskEdBox1) & " Mask=" & Q(.Mask) & " PromptChar=" & Q(.PromptChar) & _
          " PromptInclude=" & .PromptInclude & " AllowPrompt=" & .AllowPrompt & " AutoTab=" & .AutoTab & _
          " ClipMode=" & .ClipMode & " Format=" & Q(.Format) & " HideSelection=" & .HideSelection
        L "  BorderStyle=" & .BorderStyle & " Appearance=" & .Appearance & " Enabled=" & .Enabled & _
          " default=" & Q(MaskEdBox1) & " " & St()
    End With
    With MaskEdBox2
        L "  MaskEdBox2: Mask=" & Q(.Mask) & " Text=" & Q(.Text) & " Clip=" & Q(.ClipText) & " Fmt=" & _
          Q(.FormattedText) & " Max=" & .MaxLength
    End With
    If Err Then L "  ! " & Err.Number: Err.Clear

    L "== the masks"
    For Each m In Array("##/##/####", "(###) ###-####", "AAA-999", "aaa-999", "&&&&", "CCC", "??-??", ">????", _
                        "<????", ">??<??", "\##\#", "##:##", "#,###.##", "99.99", "#", "a\\b", "A#?&C9a", "x")
        Clean ""
        SetP "Mask", m
    Next
    Clean ""
    SetP "Mask", ""

    L "== Text with a mask"
    Clean "##/##/####"
    For Each v In Array("12/34/5678", "12345678", "1/2/3", "__/__/____", "ab/cd/efgh", "12/34/5678x", "1_/__/____", _
                        "12-34-5678", "", "12/34/567", "12/34/5678")
        SetP "Text", v
    Next
    Clean "(###) ###-####"
    For Each v In Array("(555) 123-4567", "5551234567", "(555) 123-456", "(555)_123-4567")
        SetP "Text", v
    Next
    Clean "AAA-999"
    For Each v In Array("abc-123", "ab1-12_", "ABC-12", "a c-1 3", "___-___")
        SetP "Text", v
    Next
    Clean ">????"
    For Each v In Array("abcd", "AbCd", "ab1d", "ab__")
        SetP "Text", v
    Next
    Clean "<??>??"
    SetP "Text", "ABcd"
    Clean "#,###.##"
    For Each v In Array("1,234.56", "1.234,56", "1234.56", "_,___.__")
        SetP "Text", v
    Next
    Clean "&&&&"
    For Each v In Array("a b!", "ab", "ab" & Chr$(9) & "c", "____")
        SetP "Text", v
    Next
    Clean "CCC"
    For Each v In Array("a b", "1!2", "   ")
        SetP "Text", v
    Next
    Clean "\##\#"
    For Each v In Array("#1#", "#a#", "#_#")
        SetP "Text", v
    Next
    Clean "##/##/####"
    SetP "Default", "11/22/3333"

    L "== ClipText and FormattedText set"
    Clean "##/##/####"
    SetP "ClipText", "12345678"
    SetP "FormattedText", "12/34/5678"

    L "== PromptChar"
    Clean "##/##/####"
    SetP "Text", "12/3_/____"
    SetP "PromptChar", "*"
    SetP "PromptChar", " "
    SetP "PromptChar", ""
    SetP "PromptChar", "ab"
    SetP "PromptChar", "1"
    SetP "Text", "12/3*/****"
    SetP "Text", "12/3 /    "
    Clean ""
    SetP "PromptChar", "#"

    L "== PromptInclude"
    Clean "##/##/####"
    SetP "Text", "12/3_/____"
    SetP "PromptInclude", False
    SetP "Text", "12/34/    "
    SetP "Text", "12/34/5___"
    SetP "Text", "12/34/5"
    SetP "PromptInclude", True

    L "== AllowPrompt"
    Clean "##/##/####"
    SetP "AllowPrompt", True
    SetP "Text", "12/3_/____"

    L "== SelStart SelLength SelText ClipMode"
    Clean "##/##/####"
    SetP "Text", "12/34/5678"
    SetP "SelStart", 1
    SetP "SelLength", 4
    L "    SelText=" & Q(MaskEdBox1.SelText)
    SetP "ClipMode", 1
    L "    SelText=" & Q(MaskEdBox1.SelText)
    SetP "ClipMode", 0
    SetP "SelText", "99"
    SetP "SelStart", 0
    SetP "SelLength", 2
    SetP "SelText", "7"
    SetP "SelStart", 3
    SetP "SelLength", 0
    SetP "SelText", "5"
    SetP "SelText", "x"
    SetP "SelStart", 20
    SetP "SelStart", -1
    SetP "SelLength", 50

    L "== MaxLength"
    Clean ""
    SetP "MaxLength", 5
    SetP "Text", "abcdefg"
    SetP "Text", "abc"
    SetP "MaxLength", 2
    SetP "MaxLength", 0
    SetP "MaxLength", -1
    SetP "MaxLength", 65
    Clean "##/##/####"
    SetP "MaxLength", 3
    SetP "MaxLength", 10
    SetP "MaxLength", 12

    L "== Mask over a text"
    Clean ""
    SetP "Text", "12345"
    SetP "Mask", "###"
    Clean ""
    SetP "Text", "1234"
    SetP "Mask", "##/##"
    Clean ""
    SetP "Text", "abc"
    SetP "Mask", "###"
    Clean "##/##/####"
    SetP "Text", "12/34/5678"
    SetP "Mask", "##-##-####"
    SetP "Mask", ""

    L "== Change"
    Clean ""
    SetP "Text", "abc"
    SetP "Text", "abc"
    SetP "Mask", "###"
    SetP "Text", "123"
    SetP "Text", "123"
    SetP "PromptChar", "*"
    SetP "Mask", ""

    L "== Format"
    Clean ""
    SetP "Format", "$#,##0.00"
    SetP "Text", "1234.5"
    Text1.SetFocus
    DoEvents
    L "    focus elsewhere: " & St() & Ev()
    Focus1
    L "    focus back: " & St() & Ev()
    SetP "Format", "0.00%"
    SetP "Text", "0.256"
    Text1.SetFocus
    DoEvents
    L "    focus elsewhere: " & St() & Ev()
    SetP "Text", "abc"
    L "    after: " & St() & Ev()
    Clean "##/##/####"
    SetP "Format", "0.00"
    SetP "Text", "12/34/5678"

    L "== typing, the mask ##/##/####"
    Clean "##/##/####"
    Focus1
    L "  focus: " & St() & Ev()
    Ty "12"
    Ty "a"
    Ty "34"
    Ty "/"
    Ty "5678"
    Ty "9"
    Ty "{BACKSPACE}"
    Ty "{HOME}"
    Ty "{DELETE}"
    Ty "{RIGHT}{RIGHT}{RIGHT}"
    Ty "{BACKSPACE}"
    Ty "{END}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}"
    Ty "_"
    Ty " "
    Ty "+{HOME}"
    Ty "0"

    L "== typing, AllowPrompt"
    Clean "##/##"
    SetP "AllowPrompt", True
    Focus1
    Ty "1_"
    Ty "2"

    L "== typing, letters"
    Clean ">??-<??"
    Focus1
    Ty "abcd"
    Ty "{HOME}AB"
    Ty "1"
    Clean "AAA"
    Focus1
    Ty "a1!"
    Clean "&&&"
    Focus1
    Ty "a !"
    Clean "#,###.##"
    Focus1
    Ty "1234"
    Ty "."
    Ty ","
    Ty "56"

    L "== typing, no mask"
    Clean ""
    SetP "MaxLength", 3
    Focus1
    Ty "abcd"
    Ty "{BACKSPACE}x"

    L "== typing, PromptInclude False"
    Clean "##/##"
    SetP "PromptInclude", False
    Focus1
    Ty "1"
    Ty "{RIGHT}{RIGHT}{RIGHT}3"

    L "== AutoTab"
    Clean "##"
    SetP "AutoTab", True
    Focus1
    Ty "1"
    Ty "2"
    L "    ActiveControl " & ActiveControl.Name
    Clean "##"
    SetP "AutoTab", False
    Focus1
    Ty "12"
    Ty "3"
    L "    ActiveControl " & ActiveControl.Name

    L "== MaskEdBox2 typing"
    MaskEdBox2.SetFocus
    DoEvents
    gEv = ""
    SendKeys "5551234567", True
    DoEvents
    L "  MaskEdBox2: Text=" & Q(MaskEdBox2.Text) & " Clip=" & Q(MaskEdBox2.ClipText) & " Sel=" & MaskEdBox2.SelStart & Ev()
    If Err Then L "  ! " & Err.Number: Err.Clear
    L "end"
End Sub

Private Sub MaskEdBox1_Change()
    gEv = gEv & " C"
End Sub

Private Sub MaskEdBox1_ValidationError(InvalidText As String, StartPosition As Integer)
    gEv = gEv & " V(" & Q(InvalidText) & "," & StartPosition & ")"
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
    gEv = gEv & " K" & KeyAscii
End Sub

Private Sub MaskEdBox1_GotFocus()
    gEv = gEv & " G1"
End Sub

Private Sub MaskEdBox1_LostFocus()
    gEv = gEv & " L1"
End Sub

Private Sub MaskEdBox2_Change()
    gEv = gEv & " C2"
End Sub

Private Sub MaskEdBox2_ValidationError(InvalidText As String, StartPosition As Integer)
    gEv = gEv & " V2(" & Q(InvalidText) & "," & StartPosition & ")"
End Sub

Private Sub Text1_GotFocus()
    gEv = gEv & " GT"
End Sub

Private Sub Text1_LostFocus()
    gEv = gEv & " LT"
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    RunAll
    Unload Me
End Sub
