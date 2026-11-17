VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Common controls 3"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6600
      Top             =   4200
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2355
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      _Version        =   393216
      FullWidth       =   65
      FullHeight      =   49
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mStep As Integer
Private mChanges As Integer
Private MSChart1 As Object       ' made by Controls.Add: VB6 keeps the settings of a chart in the .frx

Private Sub Form_Load()
    On Error Resume Next
    MkDir App.Path & "\expected"
    Kill App.Path & "\expected\f_comctl3.txt"
    On Error GoTo 0
    Set MSChart1 = Controls.Add("MSChart20Lib.MSChart", "MSChart1")
    MSChart1.Move 3480, 120, 3615, 2895
    MSChart1.Visible = True
    WriteLog "rtb " & TypeName(RichTextBox1) & " [" & RichTextBox1.Text & "] " & RichTextBox1.MultiLine & " " & RichTextBox1.ScrollBars
    WriteLog "chart " & MSChart1.chartType & " " & MSChart1.RowCount & " " & MSChart1.ColumnCount & " " & MSChart1.ShowLegend
    WriteLog "anim " & TypeName(Animation1) & " " & Animation1.AutoPlay & " " & Animation1.Center & " " & Animation1.BackStyle
    Timer1.Enabled = True
End Sub

Private Sub RichTextBox1_Change()
    mChanges = mChanges + 1
End Sub

Private Sub Report(ByVal what As String)
    WriteLog what & " err " & Err.Number
    Err.Clear
End Sub

Private Sub Timer1_Timer()
    Dim s As String, v As Variant, i As Integer, r As Integer, c As Integer, ch2 As Object
    Dim arr(1 To 4, 1 To 3) As Variant
    Timer1.Enabled = False
    mStep = mStep + 1
    WriteLog "step " & mStep
    Select Case mStep
    Case 1
        On Error Resume Next
        ' ---- RichTextBox
        mChanges = 0
        RichTextBox1.Text = "Hello world"
        WriteLog "text [" & RichTextBox1.Text & "] " & Len(RichTextBox1.Text) & " changes " & mChanges & " sel " & RichTextBox1.SelStart
        RichTextBox1.Text = "Hello world"
        WriteLog "same text changes " & mChanges
        RichTextBox1.Text = "line1" & vbCrLf & "line2" & vbCrLf & "line3"
        WriteLog "lines " & Len(RichTextBox1.Text) & " " & RichTextBox1.GetLineFromChar(0) & " " & RichTextBox1.GetLineFromChar(8) & " " & RichTextBox1.GetLineFromChar(15)
        RichTextBox1.SelStart = 7: RichTextBox1.SelLength = 5
        WriteLog "sel [" & RichTextBox1.SelText & "] " & RichTextBox1.SelStart & " " & RichTextBox1.SelLength
        RichTextBox1.SelBold = True
        WriteLog "bold " & RichTextBox1.SelBold & " changes " & mChanges
        RichTextBox1.SelStart = 0: RichTextBox1.SelLength = 5
        WriteLog "bold first " & RichTextBox1.SelBold
        RichTextBox1.SelStart = 0: RichTextBox1.SelLength = 12
        WriteLog "bold mixed " & IsNull(RichTextBox1.SelBold) & " [" & Replace(RichTextBox1.SelText, vbCrLf, "|") & "]"
        RichTextBox1.SelStart = 7: RichTextBox1.SelLength = 5
        RichTextBox1.SelItalic = True: RichTextBox1.SelUnderline = True: RichTextBox1.SelStrikeThru = True
        WriteLog "styles " & RichTextBox1.SelItalic & " " & RichTextBox1.SelUnderline & " " & RichTextBox1.SelStrikeThru
        RichTextBox1.SelColor = RGB(255, 0, 0)
        WriteLog "color " & RichTextBox1.SelColor
        RichTextBox1.SelFontName = "Arial": RichTextBox1.SelFontSize = 12
        WriteLog "font " & RichTextBox1.SelFontName & " " & RichTextBox1.SelFontSize
        RichTextBox1.SelFontSize = 0
        Report "fontsize 0"
        RichTextBox1.SelAlignment = rtfCenter
        WriteLog "align " & RichTextBox1.SelAlignment
        RichTextBox1.SelAlignment = 5
        Report "align 5"
        RichTextBox1.SelStart = 0: RichTextBox1.SelLength = 0
        WriteLog "first " & RichTextBox1.SelAlignment & " " & RichTextBox1.SelFontName & " " & RichTextBox1.SelFontSize & " " & RichTextBox1.SelColor
        RichTextBox1.SelStart = 0: RichTextBox1.SelLength = Len(RichTextBox1.Text)
        WriteLog "all " & IsNull(RichTextBox1.SelAlignment) & " " & IsNull(RichTextBox1.SelFontName) & " " & IsNull(RichTextBox1.SelColor) & " " & IsNull(RichTextBox1.SelItalic)
        RichTextBox1.SelStart = 7: RichTextBox1.SelLength = 0
        RichTextBox1.SelIndent = 300: RichTextBox1.SelHangingIndent = 150: RichTextBox1.SelRightIndent = 200
        WriteLog "indent " & RichTextBox1.SelIndent & " " & RichTextBox1.SelHangingIndent & " " & RichTextBox1.SelRightIndent
        RichTextBox1.SelStart = 14: RichTextBox1.SelLength = 0
        RichTextBox1.SelBullet = True
        WriteLog "bullet " & RichTextBox1.SelBullet
        RichTextBox1.SelStart = 0
        WriteLog "bullet first " & RichTextBox1.SelBullet
        RichTextBox1.SelStart = 14: RichTextBox1.SelLength = 5
        RichTextBox1.SelCharOffset = 60
        WriteLog "offset " & RichTextBox1.SelCharOffset
        RichTextBox1.SelProtected = True
        WriteLog "protected " & RichTextBox1.SelProtected
        RichTextBox1.SelProtected = False
        RichTextBox1.SelTabCount = 2
        RichTextBox1.SelTabs(0) = 500: RichTextBox1.SelTabs(1) = 1000
        WriteLog "tabs " & RichTextBox1.SelTabCount & " " & RichTextBox1.SelTabs(0) & " " & RichTextBox1.SelTabs(1)
        ' Find, Span, UpTo
        RichTextBox1.Text = "The cat and the hat and the bat"
        WriteLog "find " & RichTextBox1.Find("the") & " " & RichTextBox1.SelStart & " " & RichTextBox1.SelLength
        WriteLog "find2 " & RichTextBox1.Find("the", 1) & " " & RichTextBox1.Find("The", 1, , rtfMatchCase) & " " & RichTextBox1.Find("at", 0, 10, rtfWholeWord) & " " & RichTextBox1.Find("zzz")
        RichTextBox1.SelStart = 10: RichTextBox1.SelLength = 0
        WriteLog "find3 " & RichTextBox1.Find("the") & " " & RichTextBox1.SelStart
        RichTextBox1.SelStart = 2: RichTextBox1.SelLength = 0
        WriteLog "find4 " & RichTextBox1.Find("the", , , rtfNoHighlight) & " " & RichTextBox1.SelStart & " " & RichTextBox1.SelLength
        RichTextBox1.SelStart = 4: RichTextBox1.SelLength = 0
        RichTextBox1.Span "abcdefghijklmnopqrstuvwxyz", True, False
        WriteLog "span " & RichTextBox1.SelStart & " " & RichTextBox1.SelLength & " [" & RichTextBox1.SelText & "]"
        RichTextBox1.SelLength = 0
        RichTextBox1.UpTo "h", True, False
        WriteLog "upto " & RichTextBox1.SelStart & " " & RichTextBox1.SelLength
        RichTextBox1.Span " ", True, True
        WriteLog "span neg " & RichTextBox1.SelStart & " " & RichTextBox1.SelLength
        RichTextBox1.SelStart = 10: RichTextBox1.SelLength = 0
        RichTextBox1.UpTo "c", False, False
        WriteLog "upto back " & RichTextBox1.SelStart & " " & RichTextBox1.SelLength
        RichTextBox1.SelStart = 10: RichTextBox1.SelLength = 0
        RichTextBox1.Span "nda ", False, False
        WriteLog "span back " & RichTextBox1.SelStart & " " & RichTextBox1.SelLength
        ' RTF
        RichTextBox1.Text = "abc"
        s = RichTextBox1.TextRTF
        WriteLog "rtf " & Left(s, 5) & " " & (InStr(s, "abc") > 0) & " " & (Asc(Right(s, 1)))
        WriteLog "rtf [" & Replace(Replace(s, vbCrLf, "|"), Chr(0), "#") & "]"
        RichTextBox1.TextRTF = "{\rtf1\ansi{\fonttbl{\f0 Arial;}}\f0\fs20 one \b two\b0  three\par four\par}"
        WriteLog "rtf text [" & Replace(RichTextBox1.Text, vbCrLf, "|") & "] " & Len(RichTextBox1.Text)
        RichTextBox1.SelStart = 4: RichTextBox1.SelLength = 3
        WriteLog "rtf bold " & RichTextBox1.SelBold & " " & RichTextBox1.SelFontName & " " & RichTextBox1.SelFontSize
        RichTextBox1.TextRTF = "plain text"
        WriteLog "rtf plain [" & RichTextBox1.Text & "]"
        RichTextBox1.Text = "x"
        RichTextBox1.SelStart = 1
        RichTextBox1.SelRTF = "{\rtf1\ansi \i y\i0}"
        WriteLog "selrtf [" & RichTextBox1.Text & "] " & RichTextBox1.SelStart
        RichTextBox1.SelStart = 1: RichTextBox1.SelLength = 1
        WriteLog "selrtf italic " & RichTextBox1.SelItalic
        RichTextBox1.SelStart = 0: RichTextBox1.SelLength = 1
        WriteLog "selrtf get [" & Replace(Replace(RichTextBox1.SelRTF, vbCrLf, "|"), Chr(0), "#") & "]"
        ' files
        RichTextBox1.Text = "saved text"
        RichTextBox1.SaveFile App.Path & "\rtb.rtf", rtfRTF
        Report "save rtf"
        RichTextBox1.SaveFile App.Path & "\rtb.txt", rtfText
        Report "save txt"
        RichTextBox1.Text = ""
        RichTextBox1.LoadFile App.Path & "\rtb.rtf", rtfRTF
        WriteLog "load rtf [" & RichTextBox1.Text & "] " & FileLen(App.Path & "\rtb.txt")
        RichTextBox1.LoadFile App.Path & "\rtb.txt", rtfText
        WriteLog "load txt [" & RichTextBox1.Text & "]"
        RichTextBox1.Text = "before"
        RichTextBox1.LoadFile App.Path & "\rtb.txt", rtfRTF
        Report "load txt as rtf"
        WriteLog "after [" & RichTextBox1.Text & "]"
        RichTextBox1.LoadFile App.Path & "\nofile.rtf"
        Report "load missing"
        Kill App.Path & "\rtb.rtf"
        Kill App.Path & "\rtb.txt"
        RichTextBox1.MultiLine = False
        Report "multiline set"
        WriteLog "props " & RichTextBox1.MultiLine & " " & RichTextBox1.ScrollBars & " " & RichTextBox1.Locked & " " & RichTextBox1.MaxLength & " " & RichTextBox1.HideSelection
        RichTextBox1.SelStart = -1
        Report "selstart -1"
        ' ---- MSChart
        MSChart1.Row = 1: MSChart1.Column = 1
        WriteLog "chart point " & IsNumeric(MSChart1.Data) & " " & MSChart1.RowLabel & " " & MSChart1.ColumnLabel
        arr(1, 2) = "Sales": arr(1, 3) = "Cost"
        arr(2, 1) = "Jan": arr(2, 2) = 10: arr(2, 3) = 4
        arr(3, 1) = "Feb": arr(3, 2) = 20: arr(3, 3) = 8
        arr(4, 1) = "Mar": arr(4, 2) = 30: arr(4, 3) = 12
        MSChart1.ChartData = arr
        WriteLog "chartdata " & MSChart1.RowCount & " " & MSChart1.ColumnCount & " " & MSChart1.Row & " " & MSChart1.Column
        MSChart1.Row = 2: MSChart1.Column = 2
        WriteLog "point " & MSChart1.Data & " " & MSChart1.RowLabel & " " & MSChart1.ColumnLabel
        v = MSChart1.ChartData
        WriteLog "get " & LBound(v, 1) & " " & UBound(v, 1) & " " & LBound(v, 2) & " " & UBound(v, 2) & " [" & v(1, 1) & "] " & v(1, 2) & " " & v(2, 1) & " " & v(2, 2) & " " & TypeName(v(2, 2))
        MSChart1.Row = 9
        Report "row 9"
        MSChart1.Column = 0
        Report "column 0"
        MSChart1.AutoIncrement = True
        MSChart1.Row = 1: MSChart1.Column = 1
        MSChart1.Data = 15
        WriteLog "autoinc " & MSChart1.Row & " " & MSChart1.Column
        MSChart1.Row = 1: MSChart1.Column = 2
        MSChart1.Data = 5
        WriteLog "autoinc2 " & MSChart1.Row & " " & MSChart1.Column
        MSChart1.AutoIncrement = False
        MSChart1.chartType = VtChChartType2dLine
        MSChart1.TitleText = "Title"
        WriteLog "title " & MSChart1.Title.Text & " " & MSChart1.chartType
        MSChart1.Title.Text = "T2"
        WriteLog "title2 " & MSChart1.TitleText
        MSChart1.ShowLegend = True
        WriteLog "legend " & MSChart1.Legend.Location.Visible
        MSChart1.Plot.SeriesCollection(1).LegendText = "S1"
        WriteLog "series " & MSChart1.Plot.SeriesCollection.Count & " " & MSChart1.Plot.SeriesCollection(1).LegendText
        MSChart1.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Set 255, 0, 0
        WriteLog "fill " & MSChart1.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Red
        MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
        MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 50
        WriteLog "scale " & MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto & " " & MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum
        MSChart1.Plot.SeriesCollection(9).LegendText = "x"
        Report "series 9"
        With MSChart1.DataGrid
            .SetSize 1, 1, 6, 2
            WriteLog "grid " & .RowCount & " " & .ColumnCount & " " & MSChart1.RowCount & " " & .RowLabel(6, 1) & " " & .ColumnLabel(2, 1)
            .SetData 6, 2, 99, False
            .RowLabel(6, 1) = "Last"
            .InsertRows 1, 2
            WriteLog "grid ins " & .RowCount & " " & .RowLabel(8, 1) & " [" & .RowLabel(1, 1) & "]"
            .DeleteRows 1, 2
            .DeleteColumns 1, 1
            WriteLog "grid del " & .RowCount & " " & .ColumnCount
        End With
        MSChart1.Row = 6: MSChart1.Column = 1
        WriteLog "grid data " & MSChart1.Data
        WriteLog "typename " & TypeName(MSChart1.Plot) & " " & TypeName(MSChart1.DataGrid) & " " & TypeName(MSChart1.Plot.Axis(1))
        MSChart1.chartType = 99
        Report "charttype 99"
        ' ---- Animation
        Animation1.Play
        Report "play closed"
        Animation1.Open App.Path & "\nofile.avi"
        Report "open missing"
        Animation1.Open App.Path & "\test.avi"
        Report "open"
        Animation1.Play 1, 0, 2
        Report "play"
        Animation1.Play 1, 0, 9
        Report "play 9"
        Animation1.Stop
        Report "stop"
        Animation1.AutoPlay = True
        Animation1.Stop
        Report "stop autoplay"
        Animation1.AutoPlay = False
        Animation1.BackStyle = 1
        Animation1.Center = False
        WriteLog "anim " & Animation1.AutoPlay & " " & Animation1.Center & " " & Animation1.BackStyle
        Animation1.BackStyle = 3
        Report "backstyle 3"
        Animation1.Close
        Report "close"
        Animation1.Close
        Report "close again"
        WriteLog "changes " & mChanges
        ' Change, one operation at a time
        RichTextBox1.Text = "a" & vbCrLf & "b": mChanges = 0
        RichTextBox1.SelStart = 0: RichTextBox1.SelLength = 1
        RichTextBox1.SelBold = True: WriteLog "ch bold " & mChanges
        RichTextBox1.SelBold = True: WriteLog "ch bold again " & mChanges
        RichTextBox1.SelLength = 0: RichTextBox1.SelItalic = True: WriteLog "ch italic caret " & mChanges
        RichTextBox1.SelFontSize = 0: WriteLog "ch size0 " & mChanges & " " & RichTextBox1.SelFontSize
        RichTextBox1.SelAlignment = 2: WriteLog "ch align " & mChanges
        RichTextBox1.SelAlignment = 5: WriteLog "ch align5 " & mChanges
        RichTextBox1.SelIndent = 100: WriteLog "ch indent " & mChanges
        RichTextBox1.SelBullet = True: WriteLog "ch bullet " & mChanges & " " & RichTextBox1.SelIndent & " " & RichTextBox1.SelHangingIndent
        RichTextBox1.SelTabCount = 1: WriteLog "ch tabcount " & mChanges
        RichTextBox1.SelTabs(0) = 300: WriteLog "ch tabs " & mChanges
        RichTextBox1.SelProtected = True: WriteLog "ch protected " & mChanges
        RichTextBox1.SelStart = 3: RichTextBox1.SelLength = 1
        RichTextBox1.SelColor = vbRed: RichTextBox1.SelCharOffset = 40: WriteLog "ch color offset " & mChanges
        WriteLog "ch rtf [" & Replace(RichTextBox1.TextRTF, vbCrLf, "|") & "] " & mChanges
        RichTextBox1.SelText = "B": WriteLog "ch seltext " & mChanges
        RichTextBox1.SelStart = 0: RichTextBox1.SelLength = 0
        RichTextBox1.TextRTF = RichTextBox1.TextRTF: WriteLog "ch textrtf " & mChanges & " [" & Replace(RichTextBox1.Text, vbCrLf, "|") & "]"
        RichTextBox1.SelRTF = "": WriteLog "ch selrtf empty " & mChanges
        RichTextBox1.SaveFile App.Path & "\rtb.rtf": WriteLog "ch save " & mChanges
        RichTextBox1.LoadFile App.Path & "\rtb.rtf": WriteLog "ch load " & mChanges
        RichTextBox1.LoadFile App.Path & "\nofile.rtf": WriteLog "ch load missing " & mChanges
        Kill App.Path & "\rtb.rtf"
        RichTextBox1.Text = "": WriteLog "ch empty " & mChanges
        RichTextBox1.Text = "": WriteLog "ch empty again " & mChanges
        ' MSChart: the data of a new chart
        Set ch2 = Controls.Add("MSChart20Lib.MSChart", "MSChart2")
        i = 0
        For r = 1 To ch2.RowCount
            For c = 1 To ch2.ColumnCount
                ch2.Row = r: ch2.Column = c
                If Val(ch2.Data) = Int(Val(ch2.Data)) And Val(ch2.Data) >= 0 And Val(ch2.Data) < 100 Then i = i + 1
            Next
        Next
        WriteLog "chart defaults " & i & " of " & ch2.RowCount * ch2.ColumnCount & " " & ch2.chartType & " " & ch2.RowLabelCount & " " & ch2.ColumnLabelCount
        Controls.Remove "MSChart2"
        Err.Clear
        Animation1.Center = True: Report "center"
        WriteLog "center " & Animation1.Center
    Case 2
        Unload Me
        Exit Sub
    End Select
    Timer1.Enabled = True
End Sub
