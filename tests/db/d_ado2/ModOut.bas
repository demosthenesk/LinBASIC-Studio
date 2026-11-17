Attribute VB_Name = "ModOut"
Option Explicit

Public gF As Integer

Public Sub OpenOut(ByVal test As String)
    If Dir(App.Path & "\expected", vbDirectory) = "" Then MkDir App.Path & "\expected"
    gF = FreeFile
    Open App.Path & "\expected\" & test & ".txt" For Output As #gF
End Sub

Public Sub CloseOut()
    Close #gF
End Sub

Public Sub Out(ByVal s As String)
    Print #gF, s
End Sub

' a value as text: its type and the value, without the regional settings
Public Function V(x As Variant) As String
    Dim t As String
    If IsObject(x) Then
        If x Is Nothing Then V = "Nothing" Else V = "O" & TypeName(x)
        Exit Function
    End If
    t = TypeName(x)
    Select Case VarType(x)
    Case vbNull: V = "Null"
    Case vbEmpty: V = "Empty"
    Case vbString: V = "S'" & x & "'"
    Case vbBoolean: V = "B" & IIf(x, "T", "F")
    Case vbDate: V = "D" & Format$(x, "yyyy-mm-dd hh:nn:ss")
    Case vbInteger, vbLong, vbByte: V = Left$(t, 1) & Trim$(Str$(x))
    Case vbSingle, vbDouble: V = Left$(t, 2) & Trim$(Str$(x))
    Case vbCurrency: V = "C" & Trim$(Str$(x))
    Case vbDecimal: V = "M" & Trim$(Str$(x))
    Case Else
        If (VarType(x) And vbArray) <> 0 Then
            V = "A" & t & "(" & LBound(x) & " To " & UBound(x) & ")"
        Else
            V = "?" & t & "/" & VarType(x)
        End If
    End Select
End Function

' the error, if any, then clear it
Public Sub Chk(ByVal what As String)
    If Err.Number <> 0 Then
        Out "  ! " & what & ": " & Err.Number & " " & Err.Description & " [" & Err.Source & "]"
        Err.Clear
    End If
End Sub

Public Sub ChkAlways(ByVal what As String)
    Out "  ! " & what & ": " & Err.Number & " " & Err.Description & " [" & Err.Source & "]"
    Err.Clear
End Sub

' the rows from the current one to the end
Public Sub DumpRest(rs As ADODB.Recordset, ByVal title As String)
    Dim i As Integer, s As String, n As Long
    Out "-- " & title
    Do Until rs.EOF
        s = " "
        For i = 0 To rs.Fields.Count - 1
            s = s & " " & V(rs.Fields(i).Value)
        Next
        Out s
        n = n + 1
        rs.MoveNext
        If n > 50 Then Exit Do
    Loop
    Out "  rows " & n
End Sub

Public Sub DumpRs(rs As ADODB.Recordset, ByVal title As String)
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    DumpRest rs, title
End Sub

Public Function Pos(rs As ADODB.Recordset) As String
    If rs.EOF Then
        Pos = "<eof>"
    ElseIf rs.BOF Then
        Pos = "<bof>"
    Else
        Pos = "@" & rs.AbsolutePosition & " " & rs.Fields("Name").Value
    End If
End Function

Public Function MakeDb(ByVal file As String) As String
    Dim e As Object, db As Object, p As String
    p = App.Path & "\" & file
    If Dir(p) <> "" Then Kill p
    Set e = CreateObject("DAO.DBEngine.36")
    Set db = e.CreateDatabase(p, ";LANGID=0x0409;CP=1252;COUNTRY=0", 64)
    db.Execute "CREATE TABLE People (ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(20) NOT NULL, Age SHORT, " & _
        "Born DATETIME, Active BIT, Salary CURRENCY, Rate DOUBLE, Notes MEMO, Score SINGLE, Big LONG, Small BYTE)"
    db.Execute "CREATE INDEX ByName ON People (Name)"
    db.Execute "INSERT INTO People (Name, Age, Born, Active, Salary, Rate, Notes, Score, Big, Small) VALUES ('Anna', 31, #1993-04-05#, True, 1234.5, 0.25, 'first', 1.5, 100000, 7)"
    db.Execute "INSERT INTO People (Name, Age, Born, Active, Salary, Rate) VALUES ('Bill', 25, #1999-12-31 23:30:00#, False, 99.99, 1)"
    db.Execute "INSERT INTO People (Name, Age, Active, Notes) VALUES ('Cleo', 42, True, 'third')"
    db.Execute "INSERT INTO People (Name, Age, Active) VALUES ('Dino', 37, False)"
    db.Execute "INSERT INTO People (Name, Age, Active) VALUES ('Bert', 19, True)"
    db.Execute "CREATE TABLE Pets (PetID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Owner LONG, Pet TEXT(30))"
    db.Execute "INSERT INTO Pets (Owner, Pet) VALUES (1, 'cat')"
    db.Execute "INSERT INTO Pets (Owner, Pet) VALUES (1, 'dog')"
    db.Execute "INSERT INTO Pets (Owner, Pet) VALUES (3, 'fish')"
    db.CreateQueryDef "OlderThan", "PARAMETERS [MinAge] Short; SELECT Name, Age FROM People WHERE Age > [MinAge] ORDER BY Age"
    db.CreateQueryDef "Adults", "SELECT Name, Age FROM People WHERE Age >= 21 ORDER BY Name"
    db.Close
    MakeDb = p
End Function
