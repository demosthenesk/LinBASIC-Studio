VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Public Types of a public class: the program that uses the DLL knows them, and they go into a Variant
Public Type TItem
    Name As String
    Qty As Long
    Price As Double
    Code As String
    Tags(1 To 2) As String
End Type

Public Type TOther
    A As Long
    B As Long
End Type

Public Event Added(Item As TItem)

Private mItems As Collection
Private mTrace As String

Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

Public Function MakeItem(ByVal Name As String, ByVal Qty As Long, ByVal Price As Double) As TItem
    MakeItem.Name = Name
    MakeItem.Qty = Qty
    MakeItem.Price = Price
    MakeItem.Code = UCase$(Left$(Name, 2))
    MakeItem.Tags(1) = "t1-" & Name
    MakeItem.Tags(2) = "t2-" & Name
End Function

Public Function MakeOther(ByVal A As Long, ByVal B As Long) As TOther
    MakeOther.A = A
    MakeOther.B = B
End Function

Public Sub Add(Item As TItem)
    RaiseEvent Added(Item)
    mItems.Add Item
    mTrace = mTrace & "add " & Item.Name & " " & Item.Qty & "|"
End Sub

Public Function GetItem(ByVal Index As Long) As TItem
    GetItem = mItems(Index)
End Function

Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Sub Bump(Item As TItem)
    Item.Qty = Item.Qty + 1
    Item.Tags(2) = Item.Tags(2) & "+"
End Sub

Public Sub BumpVal(ByVal n As Long, Item As TItem)
    Item.Qty = Item.Qty + n
End Sub

Public Function Total() As Double
    Dim v As Variant, t As TItem
    For Each v In mItems
        t = v
        Total = Total + t.Qty * t.Price
    Next
End Function

Public Function KindOf(v As Variant) As String
    KindOf = VarType(v) & " " & TypeName(v)
End Function

Public Function TakeTrace() As String
    TakeTrace = mTrace
    mTrace = ""
End Function
