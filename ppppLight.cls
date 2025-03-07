VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ppppLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements iBoxArray
Option Explicit
Public refgroup As Object, GroupRef As Object, Arr As Boolean, IhaveClass As Boolean, CodeName As String
Public Final As Boolean, common As Boolean, myarrbase As Long
Private ndnum As Long, dnum As Long
Private actualData As Variant
Public Property Set item(curitem As Long, item As Variant)
    Set iBoxArray_Item(curitem) = item
End Property
Public Function itemnumeric(curitem As Long) As Variant
If Not Arr Then
    If curitem = -2 Then
    
    ElseIf curitem < 0 Then
        Dim usehandler As mHandler
        Set usehandler = GroupRef
        If usehandler.objref.Done Then
            If usehandler.objref.IsObj Then
                itemnumeric = 0
            Else
                itemnumeric = usehandler.objref.Value
            End If
        End If
    End If
End If
End Function
Public Property Let arrname(aName As String)
    CodeName = aName
End Property
Public Property Get arrname() As String
    arrname = CodeName
End Property
Public Sub PushDim(nDim As Long, Optional limBase As Variant)
ndnum = nDim
End Sub
Public Sub PushEnd(Optional thisarr As Variant)
dnum = ndnum
End Sub
Public Sub PushProp(A As Variant)
Set GroupRef = A
Arr = False
End Sub
Public Function IsObjAt(curitem As Long, peekvalue As Variant) As Boolean
    IsObjAt = iBoxArray_IsObjAt(curitem, peekvalue)
End Function
Public Function IsStringItem(curitem As Long) As Variant
Dim usehandler As mHandler
If curitem < 0 And curitem <> -2 Then
On Error Resume Next
Set usehandler = GroupRef
If Err.Number Then InternalError: Exit Function
On Error GoTo 0
If usehandler.T1 <> 1 Then Err.Raise 1010, , "Need an Inventory Object"
If usehandler.objref.Done Then
    If usehandler.objref.IsObj Then
        
    Else
        IsStringItem = usehandler.objref.IsString
    End If
End If
End If
End Function
Private Sub Class_Initialize()
'ndnum = 0
Dim ss$, j As Long
For j = 1 To 4
ss$ = ss$ & Chr(65 + Int((23 * Rnd) + 1))
Next j
CodeName = ChrW(&H1FFF) + ss$

End Sub

Private Sub iBoxArray_AppendArray(K As iBoxArray)
'
End Sub

Private Property Let iBoxArray_Arr(ByVal RHS As Boolean)
Arr = RHS
End Property

Private Property Get iBoxArray_Arr() As Boolean
iBoxArray_Arr = Arr
End Property

Private Property Let iBoxArray_arrname(ByVal RHS As String)
CodeName = RHS
End Property

Private Property Get iBoxArray_arrname() As String
iBoxArray_arrname = CodeName
End Property

Private Property Get iBoxArray_ArrPtr() As Long
Static alfa(0) As Variant
iBoxArray_ArrPtr = VarPtr(alfa(0))
End Property

Private Function iBoxArray_bareteamgroup() As Object
'
End Function

Private Property Let iBoxArray_CodeName(RHS As String)
CodeName = RHS
End Property

Private Property Get iBoxArray_CodeName() As String
iBoxArray_CodeName = CodeName
End Property

Private Sub iBoxArray_CopyArray(K As iBoxArray)
'
End Sub

Private Sub iBoxArray_CopyArrayRev(K As iBoxArray)
'
End Sub

Private Sub iBoxArray_CopyArrayRevFast(K As iBoxArray)
'
End Sub

Private Sub iBoxArray_CopyArraySlice(K As iBoxArray, fromS As Long, toS As Long)
'
End Sub

Private Sub iBoxArray_CopyArraySliceFast(K As iBoxArray, fromS As Long, toS As Long)
'
End Sub

Private Property Get iBoxArray_count() As Long
iBoxArray_count = 1
End Property

Private Function iBoxArray_EmptyArraySameType() As Object
'
Set iBoxArray_EmptyArraySameType = New tuple
WrongType
End Function

Private Property Let iBoxArray_Final(ByVal RHS As Boolean)
Final = RHS
End Property

Private Property Get iBoxArray_Final() As Boolean
iBoxArray_Final = Final
End Property

Private Function iBoxArray_GetCopy() As Variant()
Dim p()
iBoxArray_GetCopy = p
End Function

Private Sub iBoxArray_GetPorp(aProp As PropReference, curitem As Long)

End Sub

Private Function iBoxArray_GetUDTValue(ByVal p As Long, Name As String, Optional ndx As Variant) As Variant

End Function

Private Property Set iBoxArray_GroupRef(RHS As Object)
    Set GroupRef = RHS
End Property

Private Property Get iBoxArray_GroupRef() As Object
    Set iBoxArray_GroupRef = GroupRef
End Property

Private Property Let iBoxArray_IhaveClass(ByVal RHS As Boolean)
IhaveClass = RHS
End Property

Private Property Get iBoxArray_IhaveClass() As Boolean
iBoxArray_IhaveClass = IhaveClass
End Property

Private Property Get iBoxArray_Index() As Long

End Property

Private Property Let iBoxArray_Index(ByVal RHS As Long)

End Property

Private Property Get iBoxArray_IsEmpty() As Variant
iBoxArray_IsEmpty = dnum = 0
End Property

Private Function iBoxArray_IsEnum(p As Variant) As Boolean

End Function

Private Function iBoxArray_IsEnum2(curitem As Long, p As Variant) As Boolean

End Function

Private Function iBoxArray_isObj() As Boolean
If Not TypeOf GroupRef Is mHandler Then GoTo e1245
Dim usehandler As mHandler
Set usehandler = GroupRef
If usehandler.T1 <> 1 Then
e1245:
    Err.Raise 5010, , "Need an Inventory Object"
    Exit Function
End If
iBoxArray_isObj = usehandler.objref.IsObj
End Function

Private Function iBoxArray_IsObjAt(curitem As Long, peekvalue As Variant) As Boolean
If curitem = -2 Then
    Set peekvalue = GroupRef
    iBoxArray_IsObjAt = True
ElseIf curitem < 0 Then
    If Not TypeOf GroupRef Is mHandler Then GoTo e1245
Dim usehandler As mHandler
Set usehandler = GroupRef
If usehandler.T1 <> 1 Then
e1245:
    Err.Raise 5010, , "Need an Inventory Object"
End If
    If usehandler.objref.Done Then
        If usehandler.objref.IsObj Then
            Set peekvalue = usehandler.objref.ValueObj
            iBoxArray_IsObjAt = True
        Else
            peekvalue = usehandler.objref.Value
        End If
    Else
        Set peekvalue = usehandler
    End If
Else
    If Not Arr Then
    Set peekvalue = GroupRef
    Else
    Set peekvalue = actualData
    End If
    iBoxArray_IsObjAt = True
End If



End Function

Private Function iBoxArray_IsStringItem(curitem As Long) As Variant
iBoxArray_IsStringItem = IsStringItem(curitem)
End Function

Private Property Set iBoxArray_Item(curitem As Long, item As Variant)
'
On Error Resume Next
If curitem < 0 Then
If Not GroupRef Is Nothing Then
    If TypeOf GroupRef Is mHandler Then
        Dim usehandler As mHandler
        Set usehandler = GroupRef
        If usehandler.T1 = 1 Then
            usehandler.objref.Index = Abs(curitem + 100)
            usehandler.objref.Done = True
            Set usehandler.objref.ValueObj = item
            Exit Property
           
        End If
    ElseIf curitem = -2 Then
    If Typename$(item) = Typename$(GroupRef) Then
    Set GroupRef = item
    
    Exit Property
    End If
    End If
End If
MyEr "Wrong setting", "Λάθος όρισμα"
Else
Set actualData = item
End If

End Property

Private Property Let iBoxArray_Item(curitem As Long, item As Variant)
Dim usehandler As mHandler
If curitem < 0 Then
If curitem <> -2 Then
Set usehandler = GroupRef
If usehandler.objref.Done Then
    If MyIsObject(item) Then
    
    Else
        usehandler.objref.Value = item
    End If
    usehandler.objref.Done = False
End If

End If
End If
End Property

Private Property Get iBoxArray_Item(curitem As Long) As Variant
If curitem = -2 Then
    Set iBoxArray_Item = GroupRef
ElseIf curitem < 0 Then
    If Not TypeOf GroupRef Is mHandler Then GoTo e1245
Dim usehandler As mHandler
Set usehandler = GroupRef
If usehandler.T1 <> 1 Then
e1245:
    Err.Raise 5010, , "Need an Inventory Object"
End If
    If usehandler.objref.Done Then
        If usehandler.objref.IsObj Then
            Set iBoxArray_Item = usehandler.objref.ValueObj
        Else
            iBoxArray_Item = usehandler.objref.Value
        End If
    Else
        Set iBoxArray_Item = usehandler
    End If
Else
    If Not Arr Then
    Set iBoxArray_Item = GroupRef
    Else
    If IsObject(actualData) Then
    Set iBoxArray_Item = actualData
    Else
    iBoxArray_Item = actualData
    End If
    End If
End If
End Property

Private Function iBoxArray_ItemIsObject(curitem As Variant) As Boolean

End Function

Private Function iBoxArray_itemnumeric(curitem As Long) As Variant
SwapVariant iBoxArray_itemnumeric, itemnumeric(curitem)
End Function

Private Property Get iBoxArray_itemObject(curitem As Long) As Variant
If Arr Then
    
    If MyIsObject(actualData) Then
        Set iBoxArray_itemObject = actualData
    End If
    
Else
    Set iBoxArray_itemObject = GroupRef
End If
End Property

Private Property Get iBoxArray_itemPtr(curitem As Long) As Long
Static A
iBoxArray_itemPtr = VarPtr(A)
End Property

Private Property Let iBoxArray_ItemStr(curitem As Long, RHS As String)
'
End Property

Private Property Get iBoxArray_ItemType(curitem As Variant) As String
On Error GoTo there1
If Arr Then
   
        iBoxArray_ItemType = VarTypeName(actualData)
   
Else
If curitem = -2 Then
iBoxArray_ItemType = Typename(GroupRef)
ElseIf curitem = 0 Then
If Not TypeOf GroupRef Is mHandler Then GoTo e1245
GoTo a123
ElseIf curitem < 0 Then
If Not TypeOf GroupRef Is mHandler Then GoTo e1245
a123:
Dim usehandler As mHandler
Set usehandler = GroupRef
If usehandler.T1 <> 1 Then
e1245:
   iBoxArray_ItemType = Typename(GroupRef)
   Exit Function
End If
    If usehandler.objref.Done Then
        If usehandler.objref.IsObj Then
            iBoxArray_ItemType = Typename(usehandler.objref.ValueObj)
        Else
            Dim p
            p = usehandler.objref.Value
            iBoxArray_ItemType = VarTypeName(p)
        End If
    Else
        iBoxArray_ItemType = Typename(usehandler)
    End If

End If

End If
there1:

End Property

Private Property Get iBoxArray_ItemTypeNum(curitem As Variant) As Integer

End Property

Private Function iBoxArray_MyIsNumeric(v As Variant) As Boolean

End Function

Private Function iBoxArray_MyIsObject(v As Variant) As Boolean
'
End Function

Private Property Get iBoxArray_MyTypeToBe() As Integer
'
End Property

Private Sub iBoxArray_Neg(v As Long)
'
End Sub

Private Function iBoxArray_PlaceValue2UDT(p As Long, Name As String, v As Variant) As Boolean
'
End Function

Private Function iBoxArray_PlaceValue2UDTArray(p As Long, Name As String, v As Variant, Index As Long) As Boolean
'
End Function

Private Sub iBoxArray_processAppend(m As Object)
'
End Sub

Private Sub iBoxArray_PushDim(nDim As Long, Optional limBase As Variant)
' this has to be 1
ndnum = nDim
End Sub

Private Sub iBoxArray_PushEnd(Optional thisarr As Variant)
dnum = ndnum
End Sub

Private Sub iBoxArray_PushProp(A As Variant)
PushProp A
End Sub

Private Property Set iBoxArray_refgroup(ByVal RHS As Object)
'
Set refgroup = RHS
End Property

Private Property Get iBoxArray_refgroup() As Object
Set iBoxArray_refgroup = refgroup
End Property
Private Property Let iBoxArray_common(ByVal RHS As Boolean)
common = RHS
End Property

Private Property Get iBoxArray_common() As Boolean
iBoxArray_common = common
End Property
Private Property Let iBoxArray_myarrbase(ByVal RHS As Long)
myarrbase = RHS
End Property

Private Property Get iBoxArray_myarrbase() As Long
iBoxArray_myarrbase = myarrbase
End Property

Private Function iBoxArray_SerialItem(item As Variant, cursor As Long, Command As Long) As Boolean

End Function
Private Sub iBoxArray_SortDesTuple(Optional Low As Long = -1&, Optional high As Long = -1&)
'
End Sub

Private Sub iBoxArray_SortTuple(Optional Low As Long = -1&, Optional high As Long = -1&)
'
End Sub

Private Sub iBoxArray_StartResize()
'
End Sub

Private Sub iBoxArray_SwapItem(curitem As Long, thisvariant As Variant)
'
End Sub

Private Function iBoxArray_UpperMonoLimit() As Variant
'
End Function

Private Property Get iBoxArray_value() As Variant

End Property
Friend Property Get Unique() As Boolean
    Dim obj As IUnknown, GetRefCount As Long
    Set obj = Me
    CopyMemory GetRefCount, ByVal (ObjPtr(obj)) + 4, 4
    If GetRefCount > 4 Then Exit Property
    Unique = True
End Property
