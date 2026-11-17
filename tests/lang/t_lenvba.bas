Attribute VB_Name = "ModLenVba"
Option Explicit
' Phase 7 (robot: LedClock, UXC): Len of a function of VBA that gives a Variant (Hour, Year, TimeSerial ...): the
' length of its text, as of a Variant (lbvc types some of them Integer / Date and wanted a variable).

Public Sub T_LenVba(f As Integer)
    Dim d As Date, n As Long, i As Integer
    d = DateSerial(2026, 1, 5) + TimeSerial(9, 7, 3)
    Print #f, "hour:"; Len(Hour(d)); Hour(d)
    Print #f, "minute:"; Len(Minute(d)); "second:"; Len(Second(d))
    Print #f, "year:"; Len(Year(d)); "month:"; Len(Month(d)); "day:"; Len(Day(d)); "weekday:"; Len(Weekday(d))
    d = DateSerial(2026, 11, 25) + TimeSerial(18, 21, 41)
    Print #f, "two digits:"; Len(Hour(d)); Len(Minute(d)); Len(Second(d)); Len(Month(d)); Len(Day(d))
    Print #f, "timeserial:"; Len(TimeSerial(18, 21, 1)) - Len(CStr(TimeSerial(18, 21, 1)))
    Print #f, "dateserial:"; Len(DateSerial(2026, 1, 5)) - Len(CStr(DateSerial(2026, 1, 5)))
    Print #f, "weekdayname:"; Len(WeekdayName(5, , vbSaturday)) - Len(CStr(WeekdayName(5, , vbSaturday)))
    Print #f, "in an expression:"; Len(Hour(d)) + Len(Year(d)) * 10
    n = 12345
    i = 7
    Print #f, "variables:"; Len(n); Len(i); Len(d)
End Sub
