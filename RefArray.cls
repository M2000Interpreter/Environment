VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RefArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const VbLongLong = 20
Private Enum VARENUM
   VT_I8 = &H14
End Enum
Private Type Dec_Hdr
    DecType     As Integer
    DecScale    As Byte
    DecSign     As Byte
End Type
Private Type SAFEARRAYBOUND
    cElements                   As Long
    lLbound                     As Long
End Type
Private Declare Function HashData Lib "shlwapi" (ByVal straddr As Long, ByVal ByteSize As Long, ByVal res As Long, ByVal ressize As Long) As Long
'Private Declare Function CompareString Lib "kernel32" Alias "CompareStringW" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long) As Long
'Private Declare Function CompareStringOrdinal Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long, ByVal bIgnoreCase As Long) As Long
'Private Declare Function SafeArrayAccessData Lib "OleAut32.dll" (ByVal pSA As Long, ppVdata As Long) As Long
'Private Declare Function SafeArrayUnaccessData Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function pArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, retval As Long)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private mError As Long
Private p As Variant
Private ihavearray As Boolean, isok As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal saPtr As Long, ByVal saBound As Long) As Long
Function NewRef(Optional ByRef RHS) As refArray
Dim z As New refArray
If Not IsMissing(RHS) Then z.writevalue = RHS
Set NewRef = z
End Function
Property Get Copy() As refArray
Dim z As New refArray, pp
If isok Then
    Dim i As Long, j As Long, m
    For i = UBound(p) To 0 Step -1
        If IsObject(p(i)) Then
            If TypeOf p(i) Is refArray Then
                Set m = p(i).Copy()
                z(i) = m
            End If
        Else
            pp = p(i)
            j = Elements(pp)
            If VarType(pp) = vbVariant + vbArray Then
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
            z(i) = pp
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
            Dim ThisPtr As Long, PadDim As SAFEARRAYBOUND
            PadDim.cElements = what + 1
            ThisPtr = Peek(UnsignedAdd(VarPtr(p(where)), 8))
            If SafeArrayRedim(ByVal ThisPtr, VarPtr(PadDim)) <> 0& Then Err.Raise 5
        End If
        On Error Resume Next
        Select Case VarType(p(where)) - vbArray
        Case vbObject
            Set p(where)(what) = that
        Case VbLongLong ' 20
                If Not VarType(that) = 20 Then
                    that = cInt64(that)
                End If
                If Not VarType(that) = 20 Then mError = 6: Exit Property
                CopyMemory ByVal ((ArrPtr(where) Xor &H80000000) + 8 * CLng(what)) Xor &H80000000, ByVal ((VarPtr(that) Xor &H80000000) + 8) Xor &H80000000, 8
        Case Else
            p(where)(what) = that
        End Select
        If Err Then mError = Err.Number: Err.Clear
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
                
                Select Case VarType(p(where)) - vbArray
                Case vbObject
                    Set Value = p(where)(what)
                Case 20
                    CopyMemory ByVal ((VarPtr(Value) Xor &H80000000) + 8) Xor &H80000000, ByVal ((ArrPtr(where) Xor &H80000000) + 8 * CLng(what)) Xor &H80000000, 8
                    PutMem2 ByVal VarPtr(Value), 20
                Case vbVariant
                    If VarType(p(where)(what)) = VbLongLong Then
                        CopyMemory ByVal VarPtr(Value), ByVal ((ArrPtr(where) Xor &H80000000) + 16 * CLng(what)) Xor &H80000000, 16
                    Else
                        Value = p(where)(what)
                    End If
                Case Else
                    
                    Value = p(where)(what)
                End Select
                
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
Property Get AssignError()
    AssignError = mError
End Property
Property Get count(Optional m)
count = 0&
If isok Then
    If IsMissing(m) Then
        count = Elements(p)
    ElseIf m >= 0 And m <= UBound(p) Then
        On Error Resume Next
        count = UBound(p(m)) + 1
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
Dim aPtr As Long, BB As Byte
Dim Dummy0() As Object
Dim Dummy1() As String
Dim Dummy2() As Long
Dim Dummy3() As Currency
Dim Dummy4() As Variant
Dim Dummy5() As Double
Dim Dummy6() As Single
Dim Dummy7() As Byte
Dim Dummy8() As Integer
Select Case Arrtype
Case vbObject
    ReDim Dummy0(size)
    Value(where) = Dummy0()
Case vbString
    ReDim Dummy1(size)
    Value(where) = Dummy1()
Case vbLong
    ReDim Dummy2(size)
    Value(where) = Dummy2()
Case vbCurrency
    ReDim Dummy3(size)
    Value(where) = Dummy3()
Case vbDecimal
    ReDim Dummy4(size)
    Value(where) = Dummy4()
    PutMem2 VarPtr(p(where)), vbDecimal + vbArray
Case vbDouble
    ReDim Dummy5(size)
    Value(where) = Dummy5()
Case vbSingle
    ReDim Dummy6(size)
    Value(where) = Dummy6()
Case vbByte
    ReDim Dummy7(size)
    Value(where) = Dummy7()
Case vbInteger
    ReDim Dummy8(size)
    Value(where) = Dummy8()
Case VbLongLong
    ReDim Dummy3(size)
    Value(where) = Dummy3()
    PutMem2 VarPtr(p(where)), vbArray + VbLongLong
Case vbVariant
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
Property Get Hash(item As Long, Optional what As Long = -1, Optional ByVal size As Long = 4) As Long
    Dim mPtr As Long, ssize As Long
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
                Hash = 1
            Else
                GetMem4 mPtr - 4, ssize
                HashData mPtr, ssize, VarPtr(Hash), size
            End If
        End If
    End Select
    End If
End Property
Friend Function CompareStrEq(where As Long, what As Long, thisstr As String) As Boolean
        Const CSTR_EQUAL            As Long = 2&
        Dim mPtr As Long, ssize As Long, lPtr As Long
        GetMem4 ((VarPtr(p(where)) Xor &H80000000) + 8) Xor &H80000000, mPtr
        ''SafeArrayAccessData mPtr, iPtr
        GetMem4 ((mPtr Xor &H80000000) + 12) Xor &H80000000, mPtr
        GetMem4 mPtr + 4 * what, mPtr
            
            If mPtr = 0 Then
                GoTo ex1
            Else
                GetMem4 mPtr - 4, ssize
                ssize = ssize \ 2
                If ssize <> Len(thisstr) Then Exit Function 'GoTo ex1
                If ssize = 0 Then CompareStrEq = True: Exit Function 'GoTo ex1
                lPtr = StrPtr(thisstr)
                While ssize > 0
                    If PeekInt(lPtr) <> PeekInt(mPtr) Then Exit Function
                    ssize = ssize - 1
                    lPtr = ((lPtr Xor &H80000000) + 2) Xor &H80000000
                    mPtr = ((mPtr Xor &H80000000) + 2) Xor &H80000000
                Wend
                CompareStrEq = True
            End If
ex1:
''SafeArrayUnaccessData mPtr

End Function
Public Function cInt64(v As Variant) As Variant
    Dim DecHdr As Dec_Hdr, m As Long
    On Error GoTo er111
    cInt64 = CDec(v)
     If VarType(v) = vbString Then
        If InStr(1, v, "&h", vbTextCompare) = 1 Then
            Do
                m = Len(v)
                v = Replace(v, "&H0", "&H", , , vbTextCompare)
            Loop Until Len(v) = m Or m < 5
            If m = 10 Then
                If cInt64 < 0 Then
                    cInt64 = CDec("&H100000000") + cInt64
                End If
            End If
            
        End If
    End If
    CopyMemory DecHdr, ByVal VarPtr(cInt64), LenB(DecHdr)
    If DecHdr.DecScale Then
    cInt64 = Fix(cInt64)
    End If
    GetMem4 VarPtr(cInt64) + 12, m
    If VarType(v) = vbString Then If m < 0 And Len(v) <= 10 Then GoTo er111
    PutMem4 VarPtr(cInt64), VT_I8
    If (DecHdr.DecSign <> 0) And (cInt64 > 0) Then cInt64 = -cInt64
    If (VarType(cInt64) <> VT_I8) Then Err.Raise 6
    Exit Function
er111:
     Err.Raise 6
End Function
Property Get vtType(Optional where) As Integer
If ihavearray Then
    If isok Then
        If IsMissing(where) Then
            GetMem2 VarPtr(p), vtType
        Else
            Dim pp As Long
            pp = Elements(p(where))
            If pp > 0 Then
                GetMem2 VarPtr(p(where)), vtType
                vtType = vtType - vbArray
            End If
        End If
    End If
End If
End Property
