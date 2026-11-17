Attribute VB_Name = "ModBinFiles"
Option Explicit

Private Type TRec
    Id As Long
    Name As String * 8
    Score As Double
    Flag As Boolean
End Type

Private Type TVar
    S As String
    N As Integer
End Type

Private Function HexBytes(ByVal path As String) As String
    Dim g As Integer, b As Byte, s As String, i As Long
    g = FreeFile
    Open path For Binary Access Read As #g
    For i = 1 To LOF(g)
        Get #g, , b
        s = s & Right$("0" & Hex$(b), 2)
    Next
    Close #g
    HexBytes = s
End Function

Public Sub T_BinFiles(f As Integer)
    Dim fs As String * 5, fz As String * 3
    Print #f, Len(fz); Asc(fz)
    fs = "ab"
    Print #f, "["; fs; "]"; Len(fs)
    fs = "abcdefgh"
    Print #f, "["; fs; "]"
    Mid$(fs, 2, 1) = "X"
    Print #f, "["; fs; "]"
    fs = ""
    Dim vs As Variant
    vs = fs
    Print #f, "["; fs; "]"; Len(Trim$(fs)); Len(vs)

    Dim r As TRec, r2 As TRec, tv As TVar, tv2 As TVar
    r.Id = 7: r.Name = "Bob": r.Score = 1.5: r.Flag = True
    tv.S = "xyz": tv.N = 3
    Print #f, Len(r); Len(tv); "["; r.Name; "]"

    Dim p As String, g As Integer, i As Long
    Dim iv As Integer, lv As Long, sv As String, fv As Single, bv As Boolean, yv As Byte
    Dim vv As Variant, cv As Currency, dv As Date, arr(1 To 3) As Integer
    p = App.Path & "\bin_test.dat"
    On Error Resume Next
    Kill p
    On Error GoTo 0
    g = FreeFile
    Open p For Binary As #g
    iv = 258: Put #g, , iv
    lv = 70000: Put #g, , lv
    sv = "Hi": Put #g, , sv
    fv = 1.5: Put #g, , fv
    bv = True: Put #g, , bv
    yv = 200: Put #g, , yv
    vv = 5: Put #g, , vv
    vv = "ab": Put #g, , vv
    cv = 1.25: Put #g, , cv
    dv = DateSerial(2000, 1, 1): Put #g, , dv
    Put #g, , r
    Put #g, , tv
    arr(1) = 1: arr(2) = 2: arr(3) = 3
    Put #g, , arr
    Print #f, LOF(g); Loc(g); Seek(g)
    Close #g
    Print #f, HexBytes(p)

    Dim i2 As Integer, l2 As Long, s2 As String, f2 As Single, b2 As Boolean, y2 As Byte
    Dim v2 As Variant, s3 As String * 3
    Open p For Binary As #g
    Get #g, , i2
    Get #g, , l2
    s2 = Space$(2)
    Get #g, , s2
    Get #g, , f2
    Get #g, , b2
    Get #g, , y2
    Print #f, i2; l2; s2; f2; b2; y2; Loc(g); Seek(g)
    Get #g, , v2
    Print #f, TypeName(v2); v2;
    Get #g, , v2
    Print #f, TypeName(v2); v2
    Get #g, 3, l2
    Get #g, 7, s3
    Print #f, l2; "["; s3; "]"; Seek(g)
    Seek #g, 42
    Get #g, , r2
    Get #g, , tv2
    Print #f, r2.Id; "["; r2.Name; "]"; r2.Score; r2.Flag; "["; tv2.S; "]"; tv2.N; Seek(g)
    Print #f, EOF(g);
    Get #g, 75, i2
    Print #f, i2; EOF(g);
    i2 = 99
    Get #g, 77, i2
    Print #f, i2; EOF(g); Loc(g)
    Close #g

    Kill p
    Open p For Random As #g Len = Len(r)
    For i = 1 To 3
        r.Id = i: r.Name = "N" & i: r.Score = i / 2: r.Flag = (i Mod 2 = 0)
        Put #g, i, r
    Next
    Print #f, LOF(g); Loc(g); Seek(g)
    Get #g, 2, r2
    Print #f, r2.Id; "["; r2.Name; "]"; r2.Score; r2.Flag; Loc(g); Seek(g)
    Put #g, 5, r
    Print #f, LOF(g); Loc(g)
    Get #g, 4, r2
    Print #f, r2.Id; Asc(r2.Name); r2.Score; r2.Flag
    Get #g, 9, r2
    Print #f, r2.Id; EOF(g); Loc(g)
    Close #g

    Kill p
    Open p For Random As #g Len = 16
    sv = "hello": Put #g, 1, sv
    lv = 12345: Put #g, 2, lv
    vv = "ab": Put #g, 3, vv
    Print #f, LOF(g)
    Get #g, 1, s2
    Get #g, 2, l2
    Get #g, 3, v2
    Print #f, "["; s2; "]"; l2; v2
    On Error Resume Next
    sv = String$(20, "x")
    Put #g, 1, sv
    Print #f, "err"; Err.Number; Err.Description
    Err.Clear
    Close #g
    Kill p
End Sub
