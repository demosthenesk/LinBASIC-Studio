VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMyGreeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' the client implements the interface of the DLL
Implements CompLib.IGreeter

Private Function IGreeter_Greet(ByVal who As String) As String
    IGreeter_Greet = "Hi " & who & " from the client"
End Function
