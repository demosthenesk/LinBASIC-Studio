Attribute VB_Name = "ModEnum"
Option Explicit

Public Sub T_Enum(f As Integer)
    Dim it As CItems, v As Variant, n As Long, u As IUnknown, c As Collection, o As Object
    Set it = New CItems
    it.Add "alpha", "a"
    it.Add 2
    it.Add 3.5, "c"
    Print #f, "count"; it.Count
    Print #f, "default "; it(1); " "; it("c"); " "; it.Item(2)
    For Each v In it
        n = n + 1
        Print #f, "each"; n; v
    Next
    Set o = it
    n = 0
    For Each v In o
        n = n + 1
    Next
    Print #f, "late each"; n
    Set c = New Collection
    c.Add "x": c.Add "y"
    Set u = c.[_NewEnum]
    Print #f, "iunknown "; TypeName(u) <> ""; u Is Nothing
    For Each v In c
        Print #f, "coll "; v
    Next
    Call Randomize(1)
    Print #f, "rnd "; Rnd() < 1
    Dim e As CItems
    Set e = New CItems
    n = 0
    For Each v In e
        n = n + 1
    Next
    Print #f, "empty"; n
End Sub
