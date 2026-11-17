Attribute VB_Name = "ModFormat"
Option Explicit

Public Sub T_Format(f As Integer)
    Dim d As Date, x As Double
    x = 1234.5678
    Print #f, Format(x); "|"; Format(x, "0.00"); "|"; Format(x, "#,##0.00"); "|"; Format(x, "000000"); "|"; Format(x, "#.#")
    Print #f, Format(0.5, "0%"); "|"; Format(0.125, "0.0%"); "|"; Format(-x, "#,##0.00;(#,##0.00)"); "|"; Format(0, "0.00;-0.00;zero"); "|"; Format(1234567, "#,##0")
    Print #f, Format(x, "0.00E+00"); "|"; Format(0.00012, "0.0E-0"); "|"; Format(5, "\$0"); "|"; Format(5, """n=""0"); "|"; Format(3.14159, "##.###"); "|"; Format(0.5, "#.##")
    Print #f, Format(x, "General Number"); "|"; Format(x, "Fixed"); "|"; Format(x, "Standard"); "|"; Format(x, "Percent"); "|"; Format(x, "Scientific")
    Print #f, Format(x, "Currency"); "|"; Format(-x, "Currency"); "|"; Format(1, "Yes/No"); "|"; Format(0, "True/False"); "|"; Format(1, "On/Off")
    d = DateSerial(2003, 1, 2) + TimeSerial(13, 4, 5)
    Print #f, Format(d, "yyyy-mm-dd hh:nn:ss"); "|"; Format(d, "d/m/yy"); "|"; Format(d, "dddd, mmmm d, yyyy"); "|"; Format(d, "ddd mmm"); "|"; Format(d, "h:nn AM/PM"); "|"; Format(d, "hh:nn am/pm")
    Print #f, Format(d, "General Date"); "|"; Format(d, "Long Date"); "|"; Format(d, "Medium Date"); "|"; Format(d, "Short Date")
    Print #f, Format(d, "Long Time"); "|"; Format(d, "Medium Time"); "|"; Format(d, "Short Time"); "|"; Format(d, "q w ww y")
    Print #f, Format(d, "c"); "|"; Format(d, "ddddd"); "|"; Format(d, "ttttt"); "|"; Format(d, "dddddd"); "|"; Format(d, "mmmm"); "|"; Format(d, "dd mmmm")
    Print #f, Format("abc", ">"); "|"; Format("ABC", "<"); "|"; Format("ab", "@@@@"); "|"; Format("ab", "!@@@@"); "|"; Format("12345", "@@-@@@")
    Print #f, FormatNumber(x, 2); "|"; FormatNumber(-0.5, 1, vbTrue); "|"; FormatNumber(-0.5, 1, vbFalse, vbTrue); "|"; FormatNumber(1234567.891); "|"; FormatNumber(1234.5, 0, , , vbFalse)
    Print #f, FormatCurrency(x); "|"; FormatPercent(0.256); "|"; FormatPercent(0.256, 2); "|"; FormatDateTime(d); "|"; FormatDateTime(d, vbLongDate); "|"; FormatDateTime(d, vbShortTime)
    Print #f, MonthName(1); "|"; MonthName(2, True); "|"; WeekdayName(1); "|"; WeekdayName(2, True); "|"; WeekdayName(1, False, vbMonday)
    Print #f, "["; Format(Empty); "]["; Format(""); "]["; Format(True); "]["; Format(12, ""); "]["; Format(-3); "]"
    Print #f, Format(0, "Yes/No"); "|"; Format(1, "True/False"); "|"; Format(0, "On/Off"); "|"; Format(True, "Yes/No"); "|"; WeekdayName(7)
End Sub
