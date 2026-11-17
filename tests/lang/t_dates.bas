Attribute VB_Name = "ModDates"
Option Explicit
Public Sub T_Dates(f As Integer)
    Dim d As Date, v As Variant
    d = DateSerial(2003, 1, 2)
    Print #f, d; "|"; CStr(d); "|"; Year(d); Month(d); Day(d); Weekday(d)
    d = DateSerial(2003, 1, 2) + TimeSerial(13, 4, 5)
    Print #f, d; "|"; CStr(d)
    d = TimeSerial(9, 5, 0)
    Print #f, d; "|"; CStr(d); Hour(d); Minute(d); Second(d)
    d = #12/31/1999 11:59:59 PM#
    Print #f, CStr(d); "|"; CDbl(d)
    d = #1/1/1900#
    Print #f, CStr(d); CDbl(d)
    d = 0
    Print #f, "["; d; "]"; CStr(d)
    d = 1.5
    Print #f, CStr(d)
    v = d
    Print #f, TypeName(v); v + 1; TypeName(v + 1)
    d = DateSerial(2024, 2, 30)
    Print #f, CStr(d)
    d = DateSerial(2024, 14, 1)
    Print #f, CStr(d)
    Print #f, CStr(CDate("2003-01-02")); "|"; CStr(CDate("15/3/2010")); "|"; CStr(CDate("13:04"))
    Write #f, DateSerial(2003, 1, 2), TimeSerial(13, 4, 5), DateSerial(2003, 1, 2) + TimeSerial(13, 4, 5)
    Print #f, CStr(#1/2/2003#) & "x"; "|"; #1/2/2003# + 1; "|"; #1/2/2003# - #1/1/2003#
End Sub
