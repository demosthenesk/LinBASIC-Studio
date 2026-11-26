VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   ScaleHeight     =   3000
   ScaleWidth      =   4800
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   3000
      Top             =   1000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   600
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   7
      BuddyControl    =   "Text1"
      BuddyDispid     =   196609
      OrigLeft        =   720
      OrigTop         =   120
      OrigRight       =   960
      OrigBottom      =   405
      Max             =   20
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "abc"
      Top             =   600
      Width           =   600
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   3
      BuddyControl    =   "Text2"
      BuddyDispid     =   196611
      OrigLeft        =   720
      OrigTop         =   600
      OrigRight       =   960
      OrigBottom      =   885
      Max             =   20
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Integer, n As Integer, ch As Integer
Private Sub UpDown1_Change()
    ch = ch + 1
End Sub
Private Sub Form_Initialize()
    f = FreeFile
    Open App.Path & "\expected\f_updown.txt" For Output As #f
End Sub
Private Sub Form_Load()
    Print #f, "Load: Text1=" & Text1.Text & " Text2=" & Text2.Text & " V1=" & UpDown1.Value & " V2=" & UpDown2.Value
End Sub
Private Sub Form_Activate()
    Print #f, "Activate: Text1=" & Text1.Text & " Text2=" & Text2.Text
End Sub
Private Sub Timer1_Timer()
    n = n + 1
    If n = 1 Then
        On Error Resume Next
        Print #f, "Timer: Text1=" & Text1.Text & " Text2=" & Text2.Text
        UpDown1.Value = 9
        Print #f, "Set 9: Text1=" & Text1.Text
        Text1.Text = "12"
        Print #f, "Text 12: V1=" & UpDown1.Value
        UpDown1.Max = 5
        Print #f, "Max 5: V1=" & UpDown1.Value & " Text1=" & Text1.Text & " changes=" & ch
        UpDown1.Min = 2: UpDown1.Max = 20: UpDown1.Value = 2
        Print #f, "V 2: changes=" & ch
        UpDown1.Min = 4
        Print #f, "Min 4: V1=" & UpDown1.Value & " Text1=" & Text1.Text & " changes=" & ch
        UpDown1.Max = 1
        Print #f, "Max 1 (< Min): V1=" & UpDown1.Value & " Text1=" & Text1.Text & " changes=" & ch & " Min=" & UpDown1.Min & " Max=" & UpDown1.Max
        Err.Clear: UpDown1.Value = 0
        Print #f, "V 0: err=" & Err.Number & " V1=" & UpDown1.Value & " Text1=" & Text1.Text & " changes=" & ch
        Err.Clear: UpDown1.Value = 10
        Print #f, "V 10: err=" & Err.Number & " V1=" & UpDown1.Value & " Text1=" & Text1.Text & " changes=" & ch
        Err.Clear: UpDown1.Value = 2
        Print #f, "V 2: err=" & Err.Number & " V1=" & UpDown1.Value & " Text1=" & Text1.Text & " changes=" & ch
        Err.Clear: UpDown1.Min = 3
        Print #f, "Min 3 err=" & Err.Number
        Err.Clear: UpDown1.Max = 50
        Print #f, "Max 50 err=" & Err.Number & " V1=" & UpDown1.Value & " changes=" & ch
        Err.Clear: UpDown1.Max = 1: UpDown1.Min = 3
        Print #f, "Min 3: V1=" & UpDown1.Value & " Min=" & UpDown1.Min & " Max=" & UpDown1.Max & " changes=" & ch
    Else
        Close #f
        End
    End If
End Sub
