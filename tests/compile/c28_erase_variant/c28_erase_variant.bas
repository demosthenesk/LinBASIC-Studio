Attribute VB_Name = "ModMain"
Private Type TBox
    Items As Variant
End Type
Sub Main()
    Dim v As Variant, b As TBox, e As Long, u As Long, s As String
    On Error Resume Next
    v = Array(1, 2, 3)
    Erase v
    e = Err.Number: Err.Clear
    u = -99: u = UBound(v)
    s = "v " & e & " " & IsArray(v) & " " & IsEmpty(v) & " " & TypeName(v) & " ub=" & u & " err=" & Err.Number
    Log s
    Err.Clear
    s = "?": s = TypeName(v(0))
    Log "  elem0 " & s & " err=" & Err.Number
    Err.Clear
    b.Items = Array(1, 2, 3)
    Erase b.Items
    e = Err.Number: Err.Clear
    u = -99: u = UBound(b.Items)
    s = "v " & e & " " & IsArray(b.Items) & " " & IsEmpty(b.Items) & " " & TypeName(b.Items) & " ub=" & u & " err=" & Err.Number
    Log s
    Err.Clear
    s = "?": s = TypeName(b.Items(0))
    Log "  elem0 " & s & " err=" & Err.Number
    Err.Clear
    Log "ran"
End Sub
Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
