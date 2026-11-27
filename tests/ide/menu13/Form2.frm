VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.Menu mnuMain 
      Caption         =   "mainmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuRun 
         Caption         =   "Run this App"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuH1 
         Caption         =   "Item"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
