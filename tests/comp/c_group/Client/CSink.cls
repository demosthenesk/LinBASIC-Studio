VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' WithEvents over a class of the DLL
Public WithEvents Calc As CompLib.CCalc
Attribute Calc.VB_VarHelpID = -1
Public Veto As Boolean

Private Sub Calc_Changed(ByVal NewValue As Long, Cancel As Boolean)
    WriteLog "event Changed " & NewValue & " veto=" & Veto
    Cancel = Veto
End Sub
