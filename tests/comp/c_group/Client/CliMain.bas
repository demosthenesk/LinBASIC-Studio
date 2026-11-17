Attribute VB_Name = "CliMain"
Option Explicit
' components test (LinBASIC Studio, phase 4.3): an ActiveX DLL (CompLib) and its client in one group

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Function LogPath() As String
    LogPath = App.Path & "\..\expected\c_group.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

Public Function MyProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WriteLog "callback " & hWnd & " " & Msg & " " & wParam & " " & lParam
    MyProc = hWnd + Msg * 10 + wParam * 100 + lParam * 1000
End Function

Private Function PtrOf(ByVal p As Long) As Long
    PtrOf = p
End Function

Sub Main()
    If Dir(App.Path & "\..\expected", vbDirectory) = "" Then MkDir App.Path & "\..\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    WriteLog "start"

    ' 1. objects of the DLL
    Dim c As CompLib.CCalc
    WriteLog "trace0 " & TakeTrace()
    Set c = New CompLib.CCalc
    WriteLog "trace1 " & TakeTrace()
    WriteLog "typename " & TypeName(c) & " " & (TypeOf c Is CompLib.CCalc) & " " & (TypeOf c Is CCalc)
    c.Value = 5
    c.Add 7
    WriteLog "value " & c.Value & " default " & c
    Dim o As Object
    Set o = CreateObject("CompLib.CCalc")
    o.Value = 40
    WriteLog "createobject " & TypeName(o) & " " & o.Value & " instances " & c.Instances
    Set o = Nothing
    WriteLog "trace2 " & TakeTrace() & " instances " & c.Instances
    Dim c3 As New CCalc
    WriteLog "asnew " & c3.Value & " " & TakeTrace()
    Set c3 = Nothing
    WriteLog "trace3 " & TakeTrace()
    Dim v As Variant
    Set v = New CompLib.CCalc
    v.Add 3
    WriteLog "variant " & v.Value & " " & TypeName(v)
    Set v = Nothing
    WriteLog "trace4 " & TakeTrace()

    ' 2. GlobalMultiUse
    WriteLog "global " & Twice(21) & " " & CompLib.Twice(5) & " " & LibName

    ' 3. PublicNotCreatable and Private classes
    Dim n As CompLib.CNode
    Set n = c.MakeNode("leaf")
    WriteLog "node " & n.Name & " " & TypeName(n)
    WriteLog "hidden " & c.HiddenWork()
    On Error Resume Next
    Set o = CreateObject("CompLib.CNode")
    WriteLog "create notcreatable " & Err.Number
    Err.Clear
    Set o = CreateObject("CompLib.CHidden")
    WriteLog "create private " & Err.Number
    Err.Clear
    Set o = CreateObject("CompLib.NoSuch")
    WriteLog "create nosuch " & Err.Number
    Err.Clear
    On Error GoTo 0

    ' 4. events
    Dim s As New CSink
    Set s.Calc = c
    WriteLog "fire " & c.Fire(10)
    s.Veto = True
    WriteLog "fire " & c.Fire(100)
    Set s.Calc = Nothing
    WriteLog "fire " & c.Fire(1)

    ' 5. Implements: an interface of the DLL, in the client and in the DLL
    Dim g As New CMyGreeter
    WriteLog "greeter " & c.UseGreeter(g, "Bob")
    Dim ig As CompLib.IGreeter
    Set ig = c.MakeGreeter()
    WriteLog "dll greeter " & ig.Greet("Ann") & " " & TypeName(ig) & " " & (TypeOf ig Is IGreeter)
    WriteLog "dll greeter via dll " & c.UseGreeter(ig, "Tom")
    Set ig = g
    WriteLog "client greeter " & ig.Greet("Joe") & " " & TypeName(ig)

    ' 6. errors raised in the DLL
    On Error Resume Next
    c.Fail 1
    WriteLog "fail1 " & (Err.Number - vbObjectError) & " " & Err.Source & " " & Err.Description
    Err.Clear
    c.Fail 2
    WriteLog "fail2 " & Err.Number & " " & Err.Source & " " & Err.Description
    Err.Clear
    c.Fail 3
    WriteLog "fail3 " & Err.Number & " " & Err.Source & " " & Err.Description
    Err.Clear
    c.Fail 4
    WriteLog "fail4 " & (Err.Number - vbObjectError) & " " & Err.Source & " " & Err.Description
    Err.Clear
    On Error GoTo 0

    ' 7. arguments
    Dim a As Long, b As Long
    a = 1: b = 2
    c.Swap a, b
    WriteLog "swap " & a & " " & b
    WriteLog "sum " & c.Sum(1, 2, 3, 4) & " " & c.Sum()
    WriteLog "optional " & c.Describe() & " " & c.Describe("x")

    ' 8. For Each over an object of the DLL
    Dim bag As New CompLib.CBag, it As Variant, t As String
    bag.Add "a": bag.Add 2: bag.Add "c"
    For Each it In bag
        t = t & it & ";"
    Next
    WriteLog "foreach " & bag.Count & " " & t

    ' 9. AddressOf
    Dim r As Long
    r = CallWindowProc(AddressOf MyProc, 1, 2, 3, 4)
    WriteLog "addressof " & r
    r = CallWindowProc(PtrOf(AddressOf MyProc), 5, 6, 7, 8)
    WriteLog "addressof ptr " & r & " " & (PtrOf(AddressOf MyProc) <> 0) & " " & (PtrOf(AddressOf MyProc) = PtrOf(AddressOf MyProc))

    ' 10. Err = number (the default property of Err)
    On Error Resume Next
    Err.Raise 5
    Err = 0
    WriteLog "err0 " & Err.Number & " [" & Err.Description & "]"
    Err = 11
    WriteLog "err11 " & Err.Number & " [" & Err.Description & "]"
    Err.Clear
    Err.Number = 6
    WriteLog "errnum6 " & Err.Number & " [" & Err.Description & "]"
    Err.Clear
    On Error GoTo 0

    ' 11. the end: the objects go
    Set n = Nothing
    Set bag = Nothing
    WriteLog "trace5 " & TakeTrace() & " instances " & c.Instances
    Set c = Nothing
    WriteLog "trace6 " & TakeTrace()
    WriteLog "end"
End Sub
