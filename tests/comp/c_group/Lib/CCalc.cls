VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' MultiUse: creatable with New and CreateObject; Value is the default member
Public Event Changed(ByVal NewValue As Long, Cancel As Boolean)

Private mValue As Long
Private mId As Long

Private Sub Class_Initialize()
    gCount = gCount + 1
    gSerial = gSerial + 1
    mId = gSerial
    AddTrace "init " & mId
End Sub

Private Sub Class_Terminate()
    gCount = gCount - 1
    AddTrace "term " & mId
End Sub

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
    Value = mValue
End Property

Public Property Let Value(ByVal v As Long)
    mValue = v
End Property

Public Sub Add(ByVal n As Long)
    mValue = mValue + n
End Sub

Public Property Get Instances() As Long
    Instances = gCount
End Property

Public Function Fire(ByVal n As Long) As String
    Dim Cancel As Boolean
    RaiseEvent Changed(mValue + n, Cancel)
    If Cancel Then
        Fire = "cancelled " & mValue
    Else
        mValue = mValue + n
        Fire = "applied " & mValue
    End If
End Function

Public Function MakeNode(ByVal nm As String) As CNode
    Set MakeNode = New CNode
    MakeNode.Name = nm
End Function

Public Function HiddenWork() As String
    Dim h As New CHidden
    HiddenWork = h.Work(3) & " " & TypeName(h)
End Function

Public Function UseGreeter(ByVal g As IGreeter, ByVal who As String) As String
    UseGreeter = "[" & g.Greet(who) & "]"
End Function

Public Function MakeGreeter() As IGreeter
    Set MakeGreeter = New CStdGreeter
End Function

Public Sub Fail(ByVal kind As Long)
    Dim x As Long, z As Long
    Select Case kind
    Case 1: Err.Raise vbObjectError + 513, , "custom fail"
    Case 2: x = 1 / z
    Case 3: Err.Raise 5
    Case 4: Err.Raise vbObjectError + 514, "MySource", "with source"
    End Select
End Sub

Public Sub Swap(a As Long, b As Long)
    Dim t As Long
    t = a: a = b: b = t
End Sub

Public Function Sum(ParamArray v() As Variant) As Long
    Dim i As Long
    For i = LBound(v) To UBound(v)
        Sum = Sum + v(i)
    Next
End Function

Public Function Describe(Optional ByVal tag As String = "none") As String
    Describe = "tag=" & tag
End Function
