Attribute VB_Name = "CliMain"
Option Explicit
' components test (LinBASIC Studio, phase 7): a Public Type of an ActiveX DLL in its client:
' variables, Variant (VT_RECORD), Collection, early and late binding, ByRef, events

Private Function LogPath() As String
    LogPath = App.Path & "\..\expected\c_udt.txt"
End Function

Public Sub WriteLog(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open LogPath For Append As #f
    Print #f, s
    Close #f
End Sub

Private Function Show(it As TItem) As String
    Show = it.Name & " " & it.Qty & " " & it.Price & " [" & it.Code & "] " & it.Tags(1) & " " & it.Tags(2)
End Function

Sub Main()
    If Dir(App.Path & "\..\expected", vbDirectory) = "" Then MkDir App.Path & "\..\expected"
    If Dir(LogPath) <> "" Then Kill LogPath
    WriteLog "start"

    ' 1. a variable of the Type, from a method of the DLL
    Dim s As UdtLib.CStore, a As TItem, b As UdtLib.TItem
    Set s = New UdtLib.CStore
    a = s.MakeItem("pen", 3, 1.5)
    WriteLog "make " & Show(a)
    b = a
    b.Qty = 4
    WriteLog "copy " & a.Qty & " " & b.Qty

    ' 2. Variant
    Dim v As Variant, v2 As Variant
    v = a
    WriteLog "variant " & VarType(v) & " " & TypeName(v) & " " & IsObject(v) & " " & IsEmpty(v)
    a.Qty = 99
    b = v
    WriteLog "from variant " & Show(b)
    v2 = v
    b = v2
    WriteLog "variant copy " & b.Name & " " & b.Qty
    WriteLog "typename call " & TypeName(s.MakeItem("x", 1, 1))
    WriteLog "kindof " & s.KindOf(v) & " / " & s.KindOf(a)
    a.Qty = 3

    ' 3. events and the Collection of the DLL
    Dim k As New CSink
    Set k.Store = s
    s.Add a
    k.Factor = 10
    s.Add s.MakeItem("ink", 2, 4)
    WriteLog "trace " & s.TakeTrace() & " count " & s.Count
    b = s.GetItem(2)
    WriteLog "get " & Show(b)
    WriteLog "total " & s.Total()

    ' 4. ByRef
    s.Bump a
    WriteLog "bump " & a.Qty & " " & a.Tags(2)
    s.BumpVal 5, a
    WriteLog "bumpval " & a.Qty

    ' 5. late binding
    Dim o As Object, c As TItem
    Set o = s
    k.Factor = 0
    o.Add a
    WriteLog "late trace " & o.TakeTrace() & " count " & o.Count
    c = o.GetItem(3)
    WriteLog "late get " & Show(c)
    o.Bump c
    WriteLog "late bump " & c.Qty & " " & c.Tags(2)
    v = o.MakeItem("cap", 7, 0.5)
    WriteLog "late variant " & VarType(v) & " " & TypeName(v)
    c = v
    WriteLog "late make " & Show(c)
    WriteLog "late total " & o.Total()

    ' 6. a Collection of the client
    Dim col As New Collection, x As Variant, n As Long
    col.Add a, "a"
    col.Add c
    b = col("a")
    WriteLog "col " & b.Name & " " & b.Qty & " " & TypeName(col(2))
    For Each x In col
        b = x
        n = n + b.Qty
    Next
    WriteLog "col foreach " & n

    ' 7. errors
    Dim w As Variant
    On Error Resume Next
    w = 5
    b = w
    WriteLog "err long " & Err.Number
    Err.Clear
    w = Empty
    b = w
    WriteLog "err empty " & Err.Number
    Err.Clear
    w = s.MakeOther(1, 2)
    WriteLog "other " & VarType(w) & " " & TypeName(w) & " " & Err.Number
    Err.Clear
    b = w
    WriteLog "err other " & Err.Number
    Err.Clear
    n = 0
    n = v.Qty
    WriteLog "late field " & Err.Number & " " & n
    Err.Clear
    Dim t1 As String, t2 As String
    t1 = v.Name
    t2 = v.Tags(2)
    WriteLog "late fields " & t1 & " " & t2 & " " & Err.Number
    Err.Clear
    t1 = "-"
    t1 = TypeName(v.Price)
    WriteLog "late typename price " & t1 & " " & Err.Number
    Err.Clear
    t1 = "-"
    t1 = TypeName(v.Tags)
    WriteLog "late typename tags " & t1 & " " & Err.Number
    Err.Clear
    v.Qty = 55
    WriteLog "late let " & Err.Number
    Err.Clear
    b = v
    WriteLog "late let value " & b.Qty
    v.Tags(1) = "new"
    WriteLog "late let elem " & Err.Number
    Err.Clear
    b = v
    WriteLog "late let elem value " & b.Tags(1)
    n = -1
    n = v.Nope
    WriteLog "late nope " & Err.Number & " " & n
    Err.Clear
    On Error GoTo 0

    ' 8. named arguments (early bound)
    s.BumpVal Item:=c, n:=2
    WriteLog "named bumpval " & c.Qty
    b = s.MakeItem(Price:=2.5, Qty:=6, Name:="nm")
    WriteLog "named make " & Show(b)
    WriteLog "end"
End Sub
