VERSION 5.00
Begin VB.Form Form1
   Caption         =   "ByVal control"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2880
      Top             =   1200
   End
   Begin VB.TextBox Text1
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Hello"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1
      Caption         =   "World!"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal n As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (s As Any) As Long

Private Function Bytes(b() As Byte, ByVal n As Long) As String
    Dim i As Long, s As String
    For i = 0 To n - 1
        s = s & " " & b(i)
    Next
    Bytes = s
End Function

Private Function Nothing1() As Long
    Nothing1 = 0
End Function

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim b(15) As Byte, v As Variant, n As Long
    Timer1.Enabled = False
    CopyMemory b(0), ByVal Text1.Text, Len(Text1.Text)
    WriteLog "text:" & Bytes(b, 5)
    Erase b
    CopyMemory b(0), ByVal Label1.Caption, Len(Label1.Caption) + Nothing1()
    WriteLog "caption:" & Bytes(b, 6)
    Erase b
    CopyMemory b(0), ByVal Me.Text1.Text, 3
    WriteLog "me.text:" & Bytes(b, 4)
    Erase b
    CopyMemory b(0), ByVal Form1.Label1.Caption, 2
    WriteLog "form1.caption:" & Bytes(b, 3)
    WriteLog "lstrlen:" & lstrlen(ByVal Text1.Text) & " " & lstrlen(ByVal Label1.Caption)
    v = "abc"
    On Error Resume Next
    n = lstrlen(ByVal v)
    WriteLog "variant: " & Err.Number & " " & n
    On Error GoTo 0
    WriteLog "end"
    End
End Sub
