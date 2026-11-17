VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SysInfo.Ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   720
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Show(ByVal w As String)
    Debug.Print w; " ac="; SysInfo1.ACStatus; " st="; SysInfo1.BatteryStatus; " pct="; SysInfo1.BatteryLifePercent; " life="; SysInfo1.BatteryLifeTime; " full="; SysInfo1.BatteryFullTime
End Sub

Private Sub Form_Load()
    Show "load"
    Debug.Print "os "; SysInfo1.OSPlatform; SysInfo1.OSVersion; SysInfo1.OSBuild; " sb="; SysInfo1.ScrollBarSize
End Sub

Private Sub SysInfo1_PowerStatusChanged()
    Show "event PowerStatusChanged"
End Sub

Private Sub SysInfo1_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
    Debug.Print "event DeviceArrival"; DeviceType; DeviceID; " "; DeviceName
End Sub

Private Sub SysInfo1_SettingChanged(ByVal Item As Integer)
    Debug.Print "event SettingChanged"; Item
End Sub

Private Sub SysInfo1_PowerQuerySuspend(Cancel As Boolean)
    Debug.Print "event PowerQuerySuspend cancel="; Cancel
End Sub

Private Sub SysInfo1_PowerSuspend()
    Dim t As Single
    t = Timer
    Do While Timer - t < 0.3
        DoEvents
    Loop
    Debug.Print "event PowerSuspend (0.3s of work)"
End Sub

Private Sub SysInfo1_PowerResume()
    Debug.Print "event PowerResume"
End Sub

Private Sub Timer1_Timer()
    Show "end"
    Unload Me
End Sub
