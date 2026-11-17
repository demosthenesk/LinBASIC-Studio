# LinBASIC Studio - writes the hand-made VB6 sample files used by test_projectfiles.py
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
# The samples are written the way VB6 writes them: CRLF line ends, ANSI code page
# (cp1253, Greek Windows).  Real VB6 projects from the MSDN samples are tested
# separately once they are available locally (they are never committed).
import os, struct
S = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples")
os.makedirs(S, exist_ok=True)

def w(name, text):
    with open(os.path.join(S, name), "wb") as f:
        f.write(text.replace("\n", "\r\n").encode("cp1253"))

w("Project1.vbp", r'''Type=Exe
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\WINDOWS\system32\stdole2.tlb#OLE Automation
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
Form=Form1.frm
Module=Module1; Module1.bas
Class=CCounter; CCounter.cls
Startup="Form1"
HelpFile=""
Title="Project1"
ExeName32="Demo.exe"
Command32=""
Name="Project1"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1

[MS Transaction Server]
AutoRefresh=1
''')

w("Form1.frm", '''VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Demo ""quoted"""
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2143
      View            =   3
      LabelWrap       =   -1  'True
      _Version        =   393217
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Επιλογές"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton optChoice 
         Caption         =   "Second"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optChoice 
         Caption         =   "First"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "Form1.frx":0028
      Left            =   2520
      List            =   "Form1.frx":0034
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":0049
      Top             =   960
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   1320
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Debug.Print "Γεια σου"
End Sub

Private Sub optChoice_Click(Index As Integer)
    Text1.Text = optChoice(Index).Caption
End Sub
''')

w("CCounter.cls", '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Value As Long
Public Event Changed(ByVal NewValue As Long)

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property
''')

w("Module1.bas", '''Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()
End Sub
''')

w("Group1.vbg", r'''VBGROUP 5.0
StartupProject=Project1.vbp
Project=Lib\Lib.vbp
''')

# Form1.frx: Icon @0000, ItemData @0028, List @0034, Text @0049 ($-less multi-line Text)
ico = b"\x00\x00\x01\x00" + b"\x11" * 20
frx = struct.pack("<I", len(ico) + 8) + b"lt\x00\x00" + struct.pack("<I", len(ico)) + ico
frx += b"\0" * (0x28 - len(frx))
frx += struct.pack("<HH", 2, 2) + b"".join(struct.pack("<H", len(s)) + s for s in (b"10", b"20"))
frx += b"\0" * (0x34 - len(frx))
lst = struct.pack("<HH", 2, 5) + b"".join(struct.pack("<H", len(s)) + s for s in (b"Alpha", b"Beta"))
frx += lst + b"\0" * (0x49 - 0x34 - len(lst))
txt = "Line 1\r\nΓραμμή 2".encode("cp1253")
frx += struct.pack("<B", len(txt)) + txt          # the Text of a TextBox: [u8 len], as VB6 writes it
with open(os.path.join(S, "Form1.frx"), "wb") as f:
    f.write(frx)
