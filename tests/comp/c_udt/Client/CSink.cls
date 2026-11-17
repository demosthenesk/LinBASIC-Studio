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

' an event of the DLL with a Public Type argument (ByRef: what the handler changes, the DLL sees)
Public WithEvents Store As UdtLib.CStore
Attribute Store.VB_VarHelpID = -1
Public Factor As Long

Private Sub Store_Added(Item As UdtLib.TItem)
    WriteLog "event Added " & Item.Name & " " & Item.Qty & " " & Item.Code & " " & Item.Tags(1)
    If Factor <> 0 Then Item.Qty = Item.Qty * Factor
End Sub
