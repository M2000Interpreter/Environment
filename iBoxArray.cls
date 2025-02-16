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
' Public myarrbase As Long

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
Public Property Get arr() As Boolean

End Property

Public Property Let arr(ByVal vNewValue As Boolean)

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
Public Sub PushProp(a As Variant)

End Sub
Function IsObjAt(curitem As Long, peekvalue) As Boolean

End Function


Public Property Get myarrbase() As Long

End Property

Public Property Let myarrbase(ByVal vNewValue As Long)

End Property

Public Function SerialItem(item As Variant, cursor As Long, Command As Long) As Boolean
'
End Function

Public Property Get Count() As Long

End Property

Public Function MyIsNumeric(v As Variant) As Boolean

End Function

Public Function MyIsObject(v As Variant) As Boolean

End Function

Public Function IsEnum2(curitem As Long, p As Variant) As Boolean

End Function

Public Function ItemIsObject(curitem) As Boolean

End Function

Public Function IsStringItem(curitem As Long) As Variant

End Function

Public Function itemnumeric(curitem As Long) As Variant

End Function

Public Sub SortDesTuple(Optional Low As Long = -1, Optional high As Long = -1)

End Sub

Public Sub SortTuple(Optional Low As Long = -1, Optional high As Long = -1)

End Sub

Public Function EmptyArraySameType() As Object

End Function

Public Sub CopyArraySlice(K As iBoxArray, fromS As Long, toS As Long)

End Sub

Public Sub CopyArraySliceFast(K As iBoxArray, fromS As Long, toS As Long)

End Sub

Public Sub CopyArray(K As iBoxArray)

End Sub

Public Function IsObj() As Boolean

End Function

Public Property Get Index() As Long

End Property

Public Property Let Index(ByVal vNewValue As Long)

End Property

Public Property Get Value() As Variant

End Property

Public Sub AppendArray(K As iBoxArray)

End Sub

Public Sub CopyArrayRevFast(K As iBoxArray)

End Sub

Public Sub CopyArrayRev(K As iBoxArray)

End Sub

Public Property Get MyTypeToBe() As Integer

End Property


Public Function GetCopy() As Variant()

End Function

Public Sub processAppend(m As Object)

End Sub

Public Function GetUDTValue(ByVal p As Long, Name$, Optional ndx)

End Function

Public Property Get ItemTypeNum(curitem) As Integer

End Property
Public Sub SwapItem(curitem As Long, thisvariant)

End Sub
Public Function UpperMonoLimit()

End Function
Public Function PlaceValue2UDTArray(p As Long, Name$, v, Index As Long) As Boolean

End Function
Public Function PlaceValue2UDT(p As Long, Name$, v) As Boolean

End Function
Public Sub GetPorp(aProp As PropReference, curitem As Long)

End Sub
Public Property Let ItemStr(curitem As Long, item As String)

End Property
Public Function IsEnum(p As Variant) As Boolean

End Function
Public Sub Neg(v As Long)

End Sub
Public Property Get ArrPtr() As Long

End Property
Public Property Get itemPtr(curitem As Long) As Long

End Property

Public Function bareteamgroup() As Object

End Function
Public Sub StartResize()

End Sub
