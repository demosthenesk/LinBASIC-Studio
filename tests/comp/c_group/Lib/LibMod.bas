Attribute VB_Name = "LibMod"
Option Explicit
' the module of the ActiveX DLL: its globals are shared by all its objects

Public gTrace As String
Public gCount As Long
Public gSerial As Long

Sub Main()
    AddTrace "lib main"
End Sub

Public Sub AddTrace(ByVal s As String)
    gTrace = gTrace & s & "|"
End Sub
