Attribute VB_Name = "SrvMod"
Option Explicit
' the module of the ActiveX EXE: its globals belong to one running copy of the server (a process)

Public gMains As Long       ' Sub Main of this copy
Public gObjects As Long     ' objects made in this copy
Public gTotal As Long       ' a sum kept by this copy

Sub Main()
    gMains = gMains + 1
End Sub
