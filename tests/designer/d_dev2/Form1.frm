VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.Ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
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
   Begin MSCommLib.MSComm MSComm4 
      Left            =   6720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   300
      ParitySetting   =   4
      DataBits        =   4
      StopBits        =   2
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   6000
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   3
      ParityReplace   =   120
      BaudRate        =   115200
      ParitySetting   =   3
      DataBits        =   5
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   5280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   16
      InputLen        =   1
      BaudRate        =   1200
      ParitySetting   =   1
      DataBits        =   6
      StopBits        =   1
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4560
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      OutBufferSize   =   1024
      ParityReplace   =   0
      StopBits        =   2
      SThreshold      =   1
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
      Orientation     =   1
      RecordMode      =   0
      BorderStyle     =   0
      Silent          =   -1  'True
      Enabled         =   0   'False
      MousePointer    =   2
      Frames          =   2
      UpdateInterval  =   0
      PrevEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      EjectEnabled    =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
