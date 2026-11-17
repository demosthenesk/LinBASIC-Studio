VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' a collection class: Item is the default member, NewEnum makes For Each work

Private mCol As Collection

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Public Sub Add(ByVal Value As Variant, Optional ByVal Key As Variant)
    If IsMissing(Key) Then mCol.Add Value Else mCol.Add Value, Key
End Sub

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get Item(ByVal Index As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Item = mCol(Index)
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Function
