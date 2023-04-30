VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iBoxArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Property Get CodeName() As String

End Property

Public Property Let CodeName(vNewValue As String)

End Property

Public Property Get item(curitem As Long) As Variant

End Property

Public Property Let item(curitem As Long, vNewValue As Variant)

End Property

Public Property Set item(curitem As Long, vNewValue As Variant)

End Property
Public Property Get Arr() As Boolean

End Property

Public Property Let Arr(ByVal vNewValue As Boolean)

End Property

Public Property Get GroupRef() As Object

End Property

Public Property Set GroupRef(vNewValue As Object)

End Property

Public Property Get refgroup() As Object

End Property

Public Property Set refgroup(ByVal vNewValue As Object)

End Property
Property Get IsEmpty() As Variant

End Property
Public Property Get IhaveClass() As Boolean

End Property

Public Property Let IhaveClass(ByVal vNewValue As Boolean)


End Property
Public Property Get itemObject(curitem As Long) As Variant

End Property
Public Property Get ItemType(curitem) As String

End Property
Public Property Get arrname() As String

End Property

Public Property Let arrname(ByVal vNewValue As String)

End Property

Public Property Get Final() As Boolean

End Property

Public Property Let Final(ByVal vNewValue As Boolean)

End Property

Public Property Get common() As Boolean

End Property

Public Property Let common(ByVal vNewValue As Boolean)

End Property
Public Sub PushDim(nDim As Long, Optional limBase As Variant)

End Sub

Public Sub PushEnd(Optional thisarr)

End Sub
Public Sub PushProp(A As Variant)

End Sub
Function IsObjAt(curitem As Long, peekvalue) As Boolean

End Function

