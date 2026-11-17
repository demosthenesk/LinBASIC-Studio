Attribute VB_Name = "ModTrapErr"
Option Explicit
' Phase 6.6: the trappable errors of VB6 - the text of every error number (Error$, Err.Raise), and which
' error each operation raises (language, strings, numbers, arrays, files, objects, collections).

Private g() As Integer
Private dirT As String

Private Sub Show(f As Integer, ByVal what As String)
    If Err.Number = 0 Then
        Print #f, what; ": no error"
    Else
        Print #f, what; ":"; Err.Number; Err.Description
    End If
    Err.Clear
End Sub

' ---- 1. the texts: Error$(n) for every number with a text of its own, and Err.Raise n
Private Sub PartTexts(f As Integer)
    Dim n As Long, s As String, other As String, cnt As Long
    other = Error$(1001)
    Print #f, "default: "; other
    For n = 1 To 65535
        s = Error$(n)
        If s <> other Then
            cnt = cnt + 1
            Print #f, n; s
        End If
    Next
    Print #f, "texts:"; cnt
    On Error Resume Next
    s = Error$(0)
    Print #f, "Error$(0): ["; s; "]"; Err.Number
    Err.Clear
    s = Error$(-1)
    Show f, "Error$(-1)"
    s = Error$(65536)
    Show f, "Error$(65536)"
    s = Error$()
    Print #f, "Error$(): ["; s; "]"
    ' Err.Raise: the description and the help context of a VB error
    Dim k As Variant
    For Each k In Array(3, 5, 6, 7, 9, 11, 13, 52, 53, 62, 91, 94, 380, 424, 438, 457, 481, 1001, 31004, 32000)
        Err.Raise k
        Print #f, "raise"; k; ":"; Err.Number; Err.Description; "|"; Err.HelpContext; "|"; Err.Source = App.Title
        Err.Clear
    Next
    Err.Raise 0
    Show f, "raise 0"
    Err.Raise 65536
    Show f, "raise 65536"
    Err.Raise -5
    Show f, "raise -5"
    Err.Raise vbObjectError + 1
    Print #f, "raise vbObjectError+1:"; Err.Number - vbObjectError; Err.Description
    Err.Clear
    Err.Raise vbObjectError + 513, "Src", "Mine"
    Print #f, "raise vbObjectError+513:"; Err.Number - vbObjectError; Err.Description; " "; Err.Source
    Err.Clear
    Err.Raise &H80070005
    Print #f, "raise &H80070005:"; Hex$(Err.Number); " "; Err.Description
    Err.Clear
    Error 11
    Print #f, "Error 11:"; Err.Number; Err.Description; "|"; Err.HelpContext
    Err.Clear
    Err.Raise 5, , "custom"
    Print #f, "raise 5 custom:"; Err.Number; Err.Description; "|"; Err.HelpContext
    Err.Clear
    Err.Number = 6
    Print #f, "Err.Number = 6:"; Err.Number; "["; Err.Description; "]"
    Err.Clear
    Err.Description = "only text"
    Print #f, "Err.Description only:"; Err.Number; "["; Err.Description; "]"
    Err.Clear
End Sub

' ---- 2. language, strings, numbers, dates
Private Sub ReturnAlone()
    Return
End Sub

Private Sub ResumeAlone()
    Resume Next
End Sub

Private Sub TakeElem(e As Integer)
    ReDim g(5)
End Sub

Private Sub PartLang(f As Integer)
    Dim i As Integer, l As Long, b As Byte, d As Double, s As String, v As Variant, c As Currency, dt As Date
    On Error Resume Next
    ReturnAlone
    Show f, "Return without GoSub"
    ResumeAlone
    Show f, "Resume without error"
    ReDim g(3)
    TakeElem g(1)
    Show f, "ReDim of a locked array"
    i = 32767
    i = i + 1
    Show f, "Integer + 1"
    l = 2147483647
    l = l + 1
    Show f, "Long + 1"
    b = 255
    b = b + 1
    Show f, "Byte + 1"
    i = CInt(40000)
    Show f, "CInt(40000)"
    b = CByte(-1)
    Show f, "CByte(-1)"
    l = CLng(3000000000#)
    Show f, "CLng(3e9)"
    c = CCur(1E+20)
    Show f, "CCur(1e20)"
    d = 1E+308
    d = d * 10
    Show f, "Double * 10"
    i = 1 \ 0
    Show f, "1 \ 0"
    i = 5 Mod 0
    Show f, "5 Mod 0"
    d = 1 / 0
    Show f, "1 / 0"
    d = 0 / 0
    Show f, "0 / 0"
    d = 0 ^ -1
    Show f, "0 ^ -1"
    d = (-8) ^ 0.5
    Show f, "(-8) ^ 0.5"
    d = Sqr(-1)
    Show f, "Sqr(-1)"
    d = Log(0)
    Show f, "Log(0)"
    d = Exp(1000)
    Show f, "Exp(1000)"
    i = CInt("abc")
    Show f, "CInt(""abc"")"
    v = "a"
    v = v + 1
    Show f, """a"" + 1"
    dt = CDate("not a date")
    Show f, "CDate(text)"
    dt = DateSerial(10000, 1, 1)
    Show f, "DateSerial(10000)"
    dt = DateAdd("x", 1, Now)
    Show f, "DateAdd(""x"")"
    s = MonthName(13)
    Show f, "MonthName(13)"
    s = WeekdayName(8)
    Show f, "WeekdayName(8)"
    s = Mid$("abc", 0)
    Show f, "Mid$(s, 0)"
    s = Left$("abc", -1)
    Show f, "Left$(s, -1)"
    s = Chr$(256)
    Show f, "Chr$(256)"
    s = Chr$(-1)
    Show f, "Chr$(-1)"
    s = ChrW$(65536)
    Show f, "ChrW$(65536)"
    i = Asc("")
    Show f, "Asc("""")"
    s = String$(-1, "a")
    Show f, "String$(-1)"
    s = Space$(-1)
    Show f, "Space$(-1)"
    i = InStr(0, "abc", "b")
    Show f, "InStr(0)"
    s = Replace("abc", "", "x")
    Show f, "Replace(find """")"
    s = Replace("abc", "b", "x", 0)
    Show f, "Replace(start 0)"
    v = Split("a,b", ",", -2)
    Show f, "Split(count -2)"
    d = Round(1.5, -1)
    Show f, "Round(x, -1)"
    s = StrConv("abc", 3333)
    Show f, "StrConv(3333)"
    s = Format$(1, String$(300, "0"))
    Show f, "Format long"
    i = "a" Like "[a"
    Show f, "Like ""[a"""
    s = Null
    Show f, "String = Null"
    i = CInt(Null)
    Show f, "CInt(Null)"
    v = Null
    i = v
    Show f, "Integer = Null Variant"
    s = Hex$(Null)
    Show f, "Hex$(Null)"
    v = Empty
    v = v(0)
    Show f, "Empty(0)"
    s = Error$(1.5E+10)
    Show f, "Error$(1.5e10)"
    Dim a(5) As Integer, dyn() As Integer
    i = a(6)
    Show f, "a(6)"
    i = dyn(0)
    Show f, "dyn(0)"
    i = UBound(dyn)
    Show f, "UBound(dyn)"
    ReDim dyn(-1 To -5)
    Show f, "ReDim dyn(-1 To -5)"
    Dim k As Integer
    k = 0
    GoTo inside
    For k = 1 To 2
inside:
        i = i + 1
    Next
    Show f, "GoTo into For"
    s = Space$(&H7FFFFFF0)
    Show f, "Space$(huge)"
    Print #f, "len"; Len(s)
End Sub

' ---- 3. files and folders
Private Sub PartFiles(f As Integer)
    Dim h As Integer, h2 As Integer, s As String, l As Long, fn As String, fn2 As String
    On Error Resume Next
    dirT = App.Path & "\t66"
    Kill dirT & "\*.*"
    RmDir dirT & "\sub"
    RmDir dirT
    Err.Clear
    MkDir dirT
    fn = dirT & "\a.txt"
    fn2 = dirT & "\b.txt"
    h = FreeFile
    Open fn For Output As #h
    Print #h, "line1"
    Close #h
    Open dirT & "\none.txt" For Input As #h
    Show f, "Open missing for Input"
    Open dirT & "\nodir\x.txt" For Output As #h
    Show f, "Open in missing folder"
    Open "" For Input As #h
    Show f, "Open """""
    Open dirT For Output As #h
    Show f, "Open a folder for Output"
    Open fn For Input As #0
    Show f, "Open as #0"
    Open fn For Input As #512
    Show f, "Open as #512"
    Print #77, "x"
    Show f, "Print # not open"
    Close #77
    Show f, "Close # not open"
    s = Input$(1, #77)
    Show f, "Input$ not open"
    l = LOF(77)
    Show f, "LOF not open"
    l = EOF(77)
    Show f, "EOF not open"
    Open fn For Input As #h
    Open fn For Input As #h
    Show f, "Open the same number twice"
    h2 = FreeFile
    Open fn For Output As #h2
    Show f, "Open for Output while open for Input"
    Close #h2
    Print #h, "x"
    Show f, "Print # on Input"
    Line Input #h, s
    Line Input #h, s
    Show f, "Line Input past end"
    Input #h, s
    Show f, "Input # past end"
    s = Input$(5, #h)
    Show f, "Input$ past end"
    Kill fn
    Show f, "Kill an open file"
    Name fn As fn2
    Show f, "Name an open file"
    Close #h
    Open fn For Output As #h
    Line Input #h, s
    Show f, "Line Input on Output"
    Get #h, 1, s
    Show f, "Get on Output"
    Close #h
    Open fn For Random As #h Len = 10
    s = String$(20, "x")
    Put #h, 1, s
    Show f, "Put a long String, Len 10"
    Put #h, 0, l
    Show f, "Put record 0"
    Get #h, -1, l
    Show f, "Get record -1"
    Seek #h, 0
    Show f, "Seek 0"
    Close #h
    Open fn For Binary As #h
    Line Input #h, s
    Show f, "Line Input on Binary"
    Close #h
    Open fn For Binary Access Read Write Lock Read Write As #h
    h2 = FreeFile
    Open fn For Binary As #h2
    Show f, "Open a locked file"
    Close #h2
    Close #h
    Open fn2 For Output As #h
    Close #h
    Name fn As fn2
    Show f, "Name onto an existing file"
    Name dirT & "\none.txt" As dirT & "\c.txt"
    Show f, "Name a missing file"
    Kill dirT & "\none.txt"
    Show f, "Kill a missing file"
    Kill dirT & "\*.zzz"
    Show f, "Kill no match"
    l = FileLen(dirT & "\none.txt")
    Show f, "FileLen missing"
    l = GetAttr(dirT & "\none.txt")
    Show f, "GetAttr missing"
    SetAttr dirT & "\none.txt", vbNormal
    Show f, "SetAttr missing"
    s = FileDateTime(dirT & "\none.txt")
    Show f, "FileDateTime missing"
    FileCopy dirT & "\none.txt", dirT & "\c.txt"
    Show f, "FileCopy missing"
    MkDir dirT
    Show f, "MkDir existing"
    MkDir dirT & "\nodir\sub"
    Show f, "MkDir in missing folder"
    RmDir dirT & "\nodir"
    Show f, "RmDir missing"
    RmDir dirT
    Show f, "RmDir not empty"
    ChDir dirT & "\nodir"
    Show f, "ChDir missing"
    s = Dir(dirT & "\nodir\*.*")
    Show f, "Dir in missing folder"
    s = CurDir("?")
    Show f, "CurDir(""?"")"
    Kill dirT & "\*.*"
    RmDir dirT
    Err.Clear
End Sub

' ---- 4. objects and collections
Private Sub PartObjects(f As Integer)
    Dim o As Object, c As Collection, v As Variant, i As Long, s As String
    Dim d As Object, fso As Object
    On Error Resume Next
    i = o.Count
    Show f, "Nothing.Count"
    v = 5
    i = v.Count
    Show f, "5.Count"
    Set c = New Collection
    c.Add 1, "k"
    c.Add 2, "k"
    Show f, "Collection key twice"
    v = c(5)
    Show f, "Collection(5)"
    v = c(0)
    Show f, "Collection(0)"
    v = c("nokey")
    Show f, "Collection(""nokey"")"
    c.Remove "nokey"
    Show f, "Collection.Remove missing"
    c.Add 3, , 9
    Show f, "Collection.Add before 9"
    c.Add 4, 5
    Show f, "Collection.Add key 5 (a number)"
    Set o = c
    o.NoSuchMethod
    Show f, "late NoSuchMethod"
    i = o.Count(1)
    Show f, "late Count(1)"
    o.Add
    Show f, "late Add()"
    Set o = CreateObject("No.SuchClass")
    Show f, "CreateObject missing"
    Set o = GetObject(, "No.SuchClass")
    Show f, "GetObject missing"
    For Each v In o
    Next
    Show f, "For Each over Nothing"
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "a", 1
    d.Add "a", 2
    Show f, "Dictionary key twice"
    d.Remove "zz"
    Show f, "Dictionary.Remove missing"
    v = d.Items
    v = v(5)
    Show f, "Dictionary.Items(5)"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set o = fso.GetFile(App.Path & "\none.txt")
    Show f, "fso.GetFile missing"
    Set o = fso.GetFolder(App.Path & "\nodir")
    Show f, "fso.GetFolder missing"
    Set o = fso.OpenTextFile(App.Path & "\none.txt")
    Show f, "fso.OpenTextFile missing"
    fso.CreateFolder App.Path
    Show f, "fso.CreateFolder existing"
    fso.DeleteFile App.Path & "\none.txt"
    Show f, "fso.DeleteFile missing"
    s = CallByName(c, "NoSuch", VbGet)
    Show f, "CallByName missing"
    Set o = New Collection
    Set o = Nothing
    o.Add 1
    Show f, "Nothing.Add"
    s = TypeName(o)
    Print #f, "TypeName(Nothing): "; s
End Sub

' ---- 5. Out of stack space
Private Function Deep(ByVal n As Long) As Long
    Dim pad(20) As Long
    pad(0) = n
    Deep = Deep(n + 1) + pad(0)
End Function

Private Sub PartStack(f As Integer)
    Dim r As Long
    On Error Resume Next
    r = Deep(1)
    Show f, "endless recursion"
End Sub

Public Sub T_TrapErr(f As Integer)
    PartTexts f
    PartLang f
    PartFiles f
    PartObjects f
    PartStack f
End Sub
