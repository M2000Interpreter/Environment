VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mIndexes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim ar()
Public Linkto As mIndexes
Private Sub Class_Initialize()
    ReDim ar(0)
End Sub
Friend Sub Dump(n As Long, RHS As Variant)
    SwapVariant ar(n), RHS
End Sub
Sub GetArr(that)
    that = ar()
    Erase ar()
End Sub
Property Get Value(n As Long) As Variant
Attribute Value.VB_UserMemId = 0
    If n >= 0 And n < UBound(ar) Then
        Value = ar(n)
    End If
End Property
Property Let Value(n As Long, RHS As Variant)
    If n >= 0 And n < UBound(ar) Then
        ar(n) = RHS
    ElseIf n >= UBound(ar) Then
        ReDim Preserve ar(n + 1)
        ar(n) = RHS
    End If
End Property
Public Sub IndexOpt(n As Long)
    If n >= 0 And n < UBound(ar) Then
        OptVariant ar(n)
    ElseIf n >= UBound(ar) Then
        ReDim Preserve ar(n + 1)
        OptVariant ar(n)
    End If
End Sub
Property Get count()
    count = UBound(ar)
End Property
