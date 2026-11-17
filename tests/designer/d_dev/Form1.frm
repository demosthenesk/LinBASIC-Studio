VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSComm2 
      CommPort        =   2
      EOFEnable       =   -1  'True
      Handshaking     =   2  'comRTS
      InBufferSize    =   2048
      InputLen        =   10
      InputMode       =   1  'comInputModeBinary
      Left            =   5280
      NullDiscard     =   -1  'True
      ParityReplace   =   "#"
      RThreshold      =   1
      RTSEnable       =   -1  'True
      Settings        =   "19200,e,7,1"
      SThreshold      =   2
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4560
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MCI.MMControl MMControl2 
      AutoEnable      =   0   'False
      EjectVisible    =   0   'False
      Frames          =   5
      Height          =   495
      Left            =   120
      PlayEnabled     =   -1  'True
      RecordVisible   =   0   'False
      Shareable       =   -1  'True
      StopVisible     =   0   'False
      TabIndex        =   1
      Top             =   960
      UpdateInterval  =   500
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   "WaveAudio"
      FileName        =   "C:\Windows\Media\tada.wav"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
