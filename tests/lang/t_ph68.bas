Attribute VB_Name = "ModPh68"
Option Explicit
' Phase 6.8: what the robot found in the PSC projects that is doubtful in VB6 - a Function's own name used
' as a statement (a recursive call), qualified constants of the type libraries, App and VB.Global as objects,
' For Each over what is not a collection, Erase of a Variant field, Input # into the elements of a Variant
' array, an interface with Public variables.

Private Type TBox
    Name As String
    Items As Variant
End Type

Private mCalls As Long

Private Function Fact(ByVal n As Long) As Long
    mCalls = mCalls + 1
    If n <= 1 Then Fact = 1: Exit Function
    Fact n - 1                      ' the function's name as a statement: a call (the result is lost)
    Call Fact(n - 2)
    Fact = n * Fact(n - 1)
End Function

Private Sub ShowApp(f As Integer, a As Object)
    Print #f, "app"; TypeName(a); "|"; (a.Title = App.Title); "|"; a.Major; a.Minor; a.Revision; "|"; (a.Path = App.Path); (a.EXEName = App.EXEName)
End Sub

Private Function TryEach(f As Integer, o As Object) As Long
    Dim x As Variant, n As Long
    On Error Resume Next
    For Each x In o
        n = n + 1
        Print #f, "  in body"; n; Err.Number
        If n > 3 Then Exit For
    Next
    Print #f, "  after loop"; n; Err.Number
    TryEach = n
End Function

Public Sub T_Ph68(f As Integer)
    Dim r As Long
    ' 1. recursion through the function's own name as a statement
    mCalls = 0
    r = Fact(5)
    Print #f, "fact"; r; mCalls

    ' 2. qualified constants
    Print #f, "q1"; VBRUN.MouseButtonConstants.vbLeftButton; MouseButtonConstants.vbRightButton; VBRUN.vbMiddleButton
    Print #f, "q2"; VBA.VbStrConv.vbUpperCase; VbStrConv.vbProperCase; VBA.vbTextCompare
    Print #f, "q3"; SystemColorConstants.vbWindowBackground; ColorConstants.vbRed; VBRUN.ColorConstants.vbBlue
    Print #f, "q4"; (VBA.vbNullString = ""); Len(VBA.vbCrLf); VBA.IIf(True, "y", "n"); VbDayOfWeek.vbMonday
    Print #f, "q5"; FormShowConstants.vbModal; VBRUN.LoadResConstants.vbResIcon; ShiftConstants.vbCtrlMask

    ' 3. App and VB.Global as objects
    ShowApp f, App
    Dim g As Object, gs As String
    On Error Resume Next
    Set g = VB.Global
    Print #f, "global set"; Err.Number
    Err.Clear
    gs = "?": gs = TypeName(g)
    Print #f, "global tn "; gs; Err.Number
    Err.Clear
    gs = "?": gs = TypeName(VB.Global)
    Print #f, "global tn2 "; gs; Err.Number
    Err.Clear
    gs = "?": gs = TypeName(VB.Global.App)
    Print #f, "global app "; gs; Err.Number
    Err.Clear
    gs = "?": gs = CStr(VB.Global.App.Path = App.Path)
    Print #f, "global path "; gs; Err.Number
    Err.Clear
    On Error GoTo 0

    ' 4. For Each over what is not a collection, with On Error Resume Next
    Dim c As New Collection
    c.Add 1: c.Add 2
    Print #f, "each coll"
    r = TryEach(f, c)
    Print #f, r
    Dim fnt As New StdFont
    Print #f, "each font"
    r = TryEach(f, fnt)
    Print #f, r
    Dim nothingObj As Object
    Print #f, "each nothing"
    r = TryEach(f, nothingObj)
    Print #f, r

    ' 5. Erase of a Variant field that holds an array
    Dim b(1 To 2) As TBox
    b(1).Items = Array(1, 2, 3)
    On Error Resume Next
    Dim v As Variant
    v = Array("a", "b")
    Erase v
    Print #f, "erase v"; Err.Number; IsArray(v); IsEmpty(v); TypeName(v)
    Err.Clear
    On Error GoTo 0

    ' 6. Input # into the elements of a Variant that holds an array
    Dim fn As Integer, d As Variant, p As String
    p = App.Path & "\ph68.tmp"
    fn = FreeFile
    Open p For Output As #fn
    Write #fn, 1, "two", 3.5
    Close #fn
    d = Array(0, 0, 0)
    fn = FreeFile
    Open p For Input As #fn
    Input #fn, d(0), d(1), d(2)
    Close #fn
    Kill p
    Print #f, "input"; d(0); d(1); d(2); TypeName(d(0)); TypeName(d(1)); TypeName(d(2))

    ' 7. an interface with Public variables
    Dim impl As New CVarImpl, iv As IVarIface
    Set iv = impl
    iv.Level = 3
    iv.Label = "x"
    iv.Payload = 5
    Set iv.Payload = c
    Set iv.Owner = c
    Print #f, "iface"; iv.Level; iv.Label; TypeName(iv.Payload); iv.Owner.Count; iv.Describe
    Print #f, "iface log "; impl.Log
    Dim plain As New IVarIface
    plain.Level = 7
    plain.Label = "p"
    Print #f, "plain"; plain.Level; plain.Label; plain.Describe
End Sub
