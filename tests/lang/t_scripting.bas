Attribute VB_Name = "ModScripting"
Option Explicit

Private Function SortWords(ByVal s As String) As String
    Dim w As Variant, i As Long, j As Long, t As String
    w = Split(Trim$(s), " ")
    For i = 0 To UBound(w) - 1
        For j = i + 1 To UBound(w)
            If w(j) < w(i) Then t = w(i): w(i) = w(j): w(j) = t
        Next
    Next
    SortWords = Join(w, " ")
End Function

Public Sub T_Scripting(f As Integer)
    Dim fso As New FileSystemObject, ts As TextStream, fl As File, fo As Folder, x As Variant
    Dim base As String, p As String, n As Long, s As String
    Print #f, TypeName(fso)
    Print #f, fso.GetFileName("c:\a\b.txt"); "|"; fso.GetBaseName("c:\a\b.txt"); "|"; fso.GetExtensionName("c:\a\b.txt")
    Print #f, fso.GetParentFolderName("c:\a\b.txt"); "|"; fso.GetFileName("c:\a\dir\"); "|"; fso.GetDriveName("c:\a")
    Print #f, fso.BuildPath("c:\a", "b"); "|"; fso.BuildPath("c:\a\", "b"); "|"; fso.GetExtensionName("noext")

    base = fso.BuildPath(App.Path, "fsotest")
    If fso.FolderExists(base) Then fso.DeleteFolder base
    Set fo = fso.CreateFolder(base)
    Print #f, "folder "; fo.Name; " "; fso.FolderExists(base); " "; TypeName(fo)
    On Error Resume Next
    fso.CreateFolder base
    Print #f, "again"; Err.Number
    Err.Clear
    On Error GoTo 0

    p = fso.BuildPath(base, "one.txt")
    Set ts = fso.CreateTextFile(p, True)
    ts.WriteLine "Line one"
    ts.Write "Line "
    ts.Write "two"
    ts.WriteLine
    ts.WriteBlankLines 2
    ts.WriteLine "Last"
    Print #f, "write at"; ts.Line; ts.Column
    ts.Close
    Set fl = fso.GetFile(p)
    Print #f, TypeName(fl); " "; fl.Name; fl.Size; fso.FileExists(p)

    Set ts = fso.OpenTextFile(p, ForReading)
    Do While Not ts.AtEndOfStream
        n = ts.Line
        s = ts.ReadLine
        Print #f, n; "["; s; "]"
    Loop
    On Error Resume Next
    s = ts.ReadLine
    Print #f, "past end"; Err.Number
    Err.Clear
    On Error GoTo 0
    ts.Close

    Set ts = fso.OpenTextFile(p, ForAppending)
    ts.Write "+tail"
    ts.Close
    Set ts = fl.OpenAsTextStream(ForReading)
    Print #f, "read3 ["; ts.Read(3); "] col"; ts.Column
    ts.Skip 2
    ts.SkipLine
    x = ts.ReadAll
    Print #f, "all"; Len(x); Right(x, 10)
    ts.Close

    fso.CopyFile p, fso.BuildPath(base, "two.txt")
    fso.CopyFile p, fso.BuildPath(base, "three.dat")
    fso.MoveFile fso.BuildPath(base, "three.dat"), fso.BuildPath(base, "four.dat")
    s = ""
    For Each x In fo.Files
        s = s & x.Name & " "
    Next
    Print #f, "files "; SortWords(s)    ' the order of the folder depends on the file system
    Print #f, "count"; fo.Files.Count
    fso.DeleteFile fso.BuildPath(base, "*.txt")
    Print #f, "after delete"; fo.Files.Count; fso.FileExists(p)
    fso.CreateFolder fso.BuildPath(base, "sub")
    For Each x In fo.SubFolders
        Print #f, "sub "; x.Name; " "; x.ParentFolder.Name
    Next

    On Error Resume Next
    Set fl = fso.GetFile(fso.BuildPath(base, "none.txt"))
    Print #f, "getfile"; Err.Number
    Err.Clear
    Set fo = fso.GetFolder(fso.BuildPath(base, "none"))
    Print #f, "getfolder"; Err.Number
    Err.Clear
    Set ts = fso.OpenTextFile(fso.BuildPath(base, "none.txt"))
    Print #f, "open"; Err.Number
    Err.Clear
    On Error GoTo 0
    fso.DeleteFolder base
    Print #f, "deleted "; fso.FolderExists(base)

    Dim d As Scripting.Dictionary, o As Object, k As Variant
    Set d = New Scripting.Dictionary
    d.Add "a", 1
    d.Add "b", "two"
    d.Add 3, 3.5
    d("c") = "set"
    Print #f, TypeName(d); d.Count; d("b"); " "; d.Item(3); " "; d.Exists("a"); d.Exists("A"); d.Exists(4)
    For Each k In d
        Print #f, "key "; k; " = "; d(k)
    Next
    x = d.Keys
    Print #f, "keys"; LBound(x); UBound(x); x(0); x(3)
    x = d.Items
    Print #f, "items "; x(1)
    d.Key("c") = "cc"
    d.Remove "a"
    Print #f, "after"; d.Count; d.Exists("cc"); d.Exists("a")
    s = d("new")
    Print #f, "read adds"; d.Count; IsEmpty(d("new"))
    On Error Resume Next
    d.Add "b", 0
    Print #f, "dup"; Err.Number
    Err.Clear
    d.Remove "zz"
    Print #f, "remove"; Err.Number
    Err.Clear
    d.CompareMode = TextCompare
    Print #f, "mode"; Err.Number
    Err.Clear
    On Error GoTo 0
    d.RemoveAll
    d.CompareMode = TextCompare
    d.Add "Key", 1
    Print #f, "text "; d.Exists("KEY"); d.Count
    x = d.Keys
    d.RemoveAll
    x = d.Keys
    Print #f, "empty"; UBound(x)
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "x", 10
    Print #f, TypeName(o); o("x"); o.Count
    Set o = CreateObject("Scripting.FileSystemObject")
    Print #f, TypeName(o); o.FileExists("none")
    On Error Resume Next
    Set o = CreateObject("No.Such.Thing")
    Print #f, "create"; Err.Number
    On Error GoTo 0
End Sub
