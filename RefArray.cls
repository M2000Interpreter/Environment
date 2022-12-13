VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RefArray"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function HashData Lib "shlwapi" (ByVal straddr As Long, ByVal ByteSize As Long, ByVal res As Long, ByVal ressize As Long) As Long
Private Declare Function pArrPtr Lib "msvbvm60" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, retval As Long)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)

Private p As Variant
Private ihavearray As Boolean, isok As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private PadDim(0 To 1) As Long
Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal saPtr As Long, saBound As Long) As Long
Function NewRef(Optional ByRef RHS) As RefArray
Dim z As New RefArray
If Not IsMissing(RHS) Then z.writevalue = RHS
Set NewRef = z
End Function
Property Get Copy() As RefArray
Dim z As New RefArray, pp
If isok Then
    Dim i As Long, j As Long, m
    For i = UBound(p) To 0 Step -1
        If IsObject(p(i)) Then
            If TypeOf p(i) Is RefArray Then
                Set m = p(i).Copy()
                z(i) = m
            End If
        Else
            pp = p(i)
            j = Elements(pp)
            If Typename$(pp) = "Variant()" Then
            If j > 0 Then
                For j = j - 1 To 0 Step -1
                
                    If IsObject(p(j)) Then
                        Set m = p(i)
                        On Error Resume Next
                        Set m = p(i).Copy()
                        Err.Clear
                        z(i) = m
                        Set m = Nothing
                    End If
                Next j
            End If
            End If
            z.Value(i) = pp
        End If
    Next i
Else
    z.writevalue = p
End If
Set Copy = z
End Property
Property Let writevalue(ByRef RHS)
If IsObject(RHS) Then
    Set p = RHS
Else
    ihavearray = VarType(RHS) = 8204
    p = RHS
    If ihavearray Then
    isok = Elements(p) > 0
    End If
    
End If
End Property
Property Let Value(Optional where, Optional what, that)
If IsMissing(where) Then
If IsObject(that) Then
    Set p = that
Else
    If IsObject(p) Then Set p = Nothing: p = Empty
    p = that
End If

ElseIf ihavearray Then
    If isok Then
        If UBound(p) < where Then
            ReDim Preserve p(where * 2)
        End If
    Else
        ReDim p(where)
    End If
entry1:
    If IsMissing(what) Then
        If VarType(that) = vbObject Then
            Set p(where) = that
        Else
            p(where) = that
        End If
    ElseIf what >= 0 Then
        If Elements(p(where)) + 1 < what Then
            Dim ThisPtr As Long
            PadDim(0) = what + 1
            GetMem4 UnsignedAdd(VarPtr(p(where)), 8), ThisPtr
            If SafeArrayRedim(ThisPtr, PadDim(0)) <> 0& Then Err.Raise 5
            PutMem4 UnsignedAdd(VarPtr(p(where)), 8), ThisPtr
            
        End If
        If VarType(that) = vbObject Then
            Set p(where)(what) = that
        Else
        On Error Resume Next
            p(where)(what) = that
        End If
    End If
ElseIf Not IsMissing(where) Then
    ihavearray = True
    p = Array()
    ReDim p(where)
    isok = True
    GoTo entry1
End If
End Property
Property Get Value(Optional where, Optional what)
Attribute Value.VB_UserMemId = 0
If ihavearray Then
    If isok Then
        If UBound(p) < where Then
            ReDim Preserve p(where)
        End If
    Else
        ReDim Preserve p(where)
    End If
    If IsMissing(what) Then
        If VarType(p(where)) = vbObject Then
            Set Value = p(where)
        Else
            Value = p(where)
        End If
    Else
    Dim pp As Long
    pp = Elements(p(where))
        If pp > 0 Then
            If what >= 0 And what < pp Then
                If VarType(p(where)(what)) = vbObject Then
                    Set Value = p(where)(what)
                Else
                    Value = p(where)(what)
                End If
            End If
        End If
    End If
Else
If IsObject(p) Then
    Set Value = p
Else
    Value = p
End If
End If
End Property

Property Get Count(Optional m)
Count = 0&
If isok Then
    If IsMissing(m) Then
        Count = Elements(p)
    ElseIf m >= 0 And m <= UBound(p) Then
        On Error Resume Next
        Count = UBound(p(m)) + 1
        Err.Clear
    End If
End If
End Property
Sub UBoundTo(where)
    If ihavearray Then
        If isok Then
            ReDim Preserve p(where)
        End If
    End If
End Sub
Property Get ArrPtr(Optional where) As Long
If ihavearray Then
    If isok Then
        If IsMissing(where) Then
            GetMem4 UnsignedAdd(VarPtr(p), 8), ArrPtr
            GetMem4 UnsignedAdd(ArrPtr, 12), ArrPtr
        Else
            Dim pp As Long
            pp = Elements(p(where))
            If pp > 0 Then
            GetMem4 UnsignedAdd(VarPtr(p(where)), 8), ArrPtr
            GetMem4 UnsignedAdd(ArrPtr, 12), ArrPtr
            End If
        End If
    End If
End If
End Property
Property Get SafeArrPtr(Optional where) As Long
If ihavearray Then
    If isok Then
        If IsMissing(where) Then
            GetMem4 UnsignedAdd(VarPtr(p), 8), SafeArrPtr
        Else
            Dim pp As Long
            pp = Elements(p(where))
            If pp > 0 Then
                GetMem4 UnsignedAdd(VarPtr(p(where)), 8), SafeArrPtr
            End If
        End If
    End If
End If
End Property
Function Elements(vArray As Variant) As Long
    Dim lPtr            As Long
    Const VT_BYREF      As Long = &H4000
    If IsArray(vArray) Then
    lPtr = Peek(UnsignedAdd(VarPtr(vArray), 8))
    If (Peek(VarPtr(vArray)) And VT_BYREF) <> 0 Then
        lPtr = Peek(lPtr)
    End If
    If lPtr <> 0 Then
        Elements = Peek(UnsignedAdd(lPtr, 16))
    End If
    End If
End Function
Public Function Peek(ByVal lPtr As Long) As Long
    GetMem4 lPtr, Peek
End Function

Public Function PeekInt(ByVal lPtr As Long) As Integer
    GetMem2 lPtr, PeekInt
End Function
Public Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
Public Sub DefArrayAt(where, Arrtype, Optional size As Long = 10)
If size < 1 Then size = 1
Dim aPtr As Long, bb As Byte
    
Select Case Arrtype
Case vbByte
    Dim Dummy7() As Byte
    ReDim Dummy7(size)
    Value(where) = Dummy7()
Case vbInteger
    Dim Dummy8() As Integer
    ReDim Dummy8(size)
    Value(where) = Dummy8()
Case vbLong
    Dim Dummy() As Long
    ReDim Dummy(size)
    Value(where) = Dummy()
Case vbDouble
    Dim Dummy5() As Double
    ReDim Dummy5(size)
    Value(where) = Dummy5()
Case vbSingle
    Dim Dummy6() As Single
    ReDim Dummy6(size)
    Value(where) = Dummy6()
Case vbString
    Dim Dummy1() As String
    ReDim Dummy1(size)
    Value(where) = Dummy1()
Case vbCurrency
    Dim Dummy3() As Currency
    ReDim Dummy3(size)
    Value(where) = Dummy3()
Case vbVariant
    Dim Dummy4()
    ReDim Dummy4(size)
    Value(where) = Dummy4()
End Select

End Sub
Property Get ItemType(item) As Long
    If ihavearray And isok Then
        If Not IsMissing(item) Then
            ItemType = VarType(p(item))
        Else
            ItemType = VarType(p)
        End If
    End If
End Property
Property Get Hash(item As Long, Optional what As Long = -1, Optional ByVal size As Long = 4) As Byte()
    Dim mPtr As Long, ssize As Long, aPtr As Long
    If size < 1 Then Exit Property
    If ihavearray And isok Then
    
    Select Case VarType(Value(item))
    Case 8200
        If what < 0 Or what > UBound(p(item)) Then
            Err.Raise 5
        Else
            GetMem4 UnsignedAdd(VarPtr(p(item)), 8), mPtr
            GetMem4 UnsignedAdd(mPtr, 12), mPtr
            GetMem4 mPtr + 4 * what, mPtr
            If mPtr = 0 Then
                ReDim Hash(size - 1)
                mPtr = pArrPtr(Hash)
                GetMem4 mPtr, mPtr  ' is by reference (we get *safearray1D)
                GetMem4 UnsignedAdd(mPtr, 12), mPtr 'we get the *pvData)
                PutMem1 mPtr, 1  ' just set 1
            Else
                ReDim Hash(size - 1)
                aPtr = pArrPtr(Hash)
                GetMem4 aPtr, aPtr  ' is by reference (we get *safearray1D)
                GetMem4 UnsignedAdd(aPtr, 12), aPtr 'we get the *pvData)
                GetMem4 mPtr - 4, ssize
                HashData mPtr, ssize, aPtr, size
            End If
        End If
    End Select
    End If
End Property
Property Get HashkEY(Key, Optional ByVal size As Long = 4) As Byte()
    Dim aPtr As Long, sPtr As Long, ssize As Long
    If size < 1 Then Exit Property
        ReDim HashkEY(size - 1)
        aPtr = pArrPtr(HashkEY)
        GetMem4 aPtr, aPtr  ' is by reference (we get *safearray1D)
        GetMem4 UnsignedAdd(aPtr, 12), aPtr 'we get the *pvData)
        GetMem4 VarPtr(Key) + 8, sPtr
        GetMem4 sPtr - 4, ssize
        If ssize = 0 Then
            PutMem1 aPtr, 1
        Else
            HashData sPtr, ssize, aPtr, size
        End If
End Property
Public Function tuple(ParamArray v())
    Dim p() As Variant
    p = v
    If Elements(p) > 0 Then
        tuple = p
    Else
        Dim z() As Variant
        tuple = z()
    End If
End Function

