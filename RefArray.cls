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
Private OneDimension As Boolean
Public flat As Boolean
Function NewRef(Optional ByRef RHS) As refArray
Dim z As New refArray
If Not IsMissing(RHS) Then z.writevalue = RHS
z.flat = flat
Set NewRef = z
End Function
Property Get Copy() As refArray
Dim z As New refArray, pp
If isok Then
        Dim i As Long, j As Long, m, ra As refArray
        If flat Then
            z.writevalue = p
            z.flat = flat
            If VarType(p) = vbArray + vbObject Then
                For i = UBound(p) To 0 Step -1
                    If Not p(i) Is Nothing Then
                        If TypeOf p(i) Is refArray Then
                            Set m = p(i).Copy()
                            z(i) = m
                        Else
                            Set m = p(i)
                            On Error GoTo 123
                            Set m = p(i).Copy()
                            z(i) = m
123
                            Err.Clear
                            Set m = Nothing
                        End If
                    End If
                Next i
            End If
            Set Copy = z
            Exit Property
        End If
Dim ppp As Object
    For i = UBound(p) To 0 Step -1
        If IsObject(p(i)) Then
        Set ppp = p(i)
            If ppp Is Nothing Then
                z(i) = CVar(Nothing)
            ElseIf TypeOf p(i) Is refArray Then
                Set ra = p(i)
                If ra.MarkTwoDimension Then
                    Set ra = ra.Copy
                    ra.MarkTwoDimension = True
                     z(i) = CVar(ra)
                Else
                   z(i) = CVar(ra.Copy)
                End If
                
            End If
        Else
            pp = p(i)
            j = Elements(pp)
            If VarType(pp) = vbVariant + vbArray Then
            If j > 0 Then
                For j = j - 1 To 0 Step -1
                
                    If IsObject(pp(j)) Then
                        
                        Set m = pp(j)
                        On Error Resume Next
                        Set m = pp(j).Copy()
                        Err.Clear
                        z(j) = m
                        Set m = Nothing
                    End If
                Next j
            End If
            ElseIf VarType(pp) = vbArray + vbObject Then
            If j > 0 Then
                For j = j - 1 To 0 Step -1
                
                   
                        Set m = pp(j)
                        If Not m Is Nothing Then
                        On Error Resume Next
                        Set m = m.Copy()
                        Err.Clear
                        Set ppp = m
                         Set z(j) = ppp
                        End If
                        Set m = Nothing
                   
                Next j
            End If
            End If
            z(i) = pp
        End If
    Next i
Else
    z.writevalue = p
End If
z.MarkTwoDimension = Not OneDimension
Set Copy = z
End Property
Property Let writevalue(ByRef RHS)
If IsObject(RHS) Then
    Set p = RHS
Else
    ihavearray = (VarType(RHS) And vbArray) <> 0
   ' flat = VarType(RHS) = vbArray + vbVariant
    p = RHS
    If ihavearray Then
    isok = Elements(p) > 0
    End If
    
End If
End Property
Property Let Value(Optional where, Optional what, that)
Dim mt As Integer, wt As Integer
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
        
        '    ReDim Preserve p(where * 2)
            ReDim Preserve p(where)
        
        End If
    Else
        ReDim p(where)
    End If
entry1:
    If IsMissing(what) Or flat Then
        If flat And Not IsMissing(what) Then
            If where = 0 Then
                If VarType(what) = vbObject Then
                    Set p(what) = that
                ElseIf VarType(p(what)) = VbLongLong Then
                    If Not VarType(that) = VbLongLong Then
                        that = cInt64(that)
                    End If
                    If Not VarType(that) = VbLongLong Then mError = 6: Exit Property
                    CopyMemory ByVal ((ArrPtr() Xor &H80000000) + 8 * CLng(what)) Xor &H80000000, ByVal ((VarPtr(that) Xor &H80000000) + 8) Xor &H80000000, 8
                Else
                    p(what) = that
                End If
            Else
                If VarType(that) = vbObject Then
                    Set p(where) = that
                ElseIf VarType(p(where)) = VbLongLong Then
                    If Not VarType(that) = VbLongLong Then
                        that = cInt64(that)
                    End If
                    If Not VarType(that) = VbLongLong Then mError = 6: Exit Property
                    CopyMemory ByVal ((ArrPtr() Xor &H80000000) + 8 * CLng(where)) Xor &H80000000, ByVal ((VarPtr(that) Xor &H80000000) + 8) Xor &H80000000, 8
                Else
                    p(where) = that
                End If
            End If
        Else
            On Error Resume Next
            mt = vtType()
            wt = VarType(that)
            If wt = vbObject Then
                If mt = 8201 Or mt = 8204 Then
                    Set p(where) = that
                Else
                    mError = 91
                End If
            ElseIf mt = 8201 Then ' vbObject + vbArray
                mError = 91
            ElseIf mt = 8212 Then
                If Not VarType(that) = VbLongLong Then
                    that = cInt64(that)
                End If
                If Not VarType(that) = VbLongLong Then mError = 6: Exit Property
                CopyMemory ByVal ((ArrPtr() Xor &H80000000) + 8 * CLng(where)) Xor &H80000000, ByVal ((VarPtr(that) Xor &H80000000) + 8) Xor &H80000000, 8
            ElseIf mt = 8200 Then
                If wt = vbString Then
                        p(where) = that
                Else
                    mError = 5030
                End If
            ElseIf wt = vbString Then
                If mt = vbVariant Or mt = vbString Then
                    p(where) = that
                Else
                    mError = 5020
                End If
            ElseIf myVarType(that, vbObject) Then
                Set p(where) = that
            Else
                p(where) = that
            End If
            If Err Then mError = Err.Number: Err.Clear
        End If
    ElseIf what >= 0 Then
        If Elements(p(where)) < what + 1 Then
            Dim ThisPtr As Long, PadDim As SAFEARRAYBOUND
            PadDim.cElements = what + 1
            ThisPtr = Peek(UnsignedAdd(VarPtr(p(where)), 8))
            If SafeArrayRedim(ByVal ThisPtr, VarPtr(PadDim)) <> 0& Then Err.Raise 5
        End If
        On Error Resume Next
            mt = vtType(where)
            wt = VarType(that)
        Select Case mt
        Case vbObject
            If wt = vbObject Then
                Set p(where)(what) = that
            ElseIf myVarType(that, vbObject) Then
                    Set p(where)(what) = that
            Else
                mError = 91
            End If
        Case VbLongLong ' 20
                If Not VarType(that) = VbLongLong Then
                    that = cInt64(that)
                End If
                If Not VarType(that) = VbLongLong Then mError = 6: Exit Property
                CopyMemory ByVal ((ArrPtr(where) Xor &H80000000) + 8 * CLng(what)) Xor &H80000000, ByVal ((VarPtr(that) Xor &H80000000) + 8) Xor &H80000000, 8
        Case vbVariant
            If wt = vbObject Then
                Set p(where)(what) = that
            ElseIf myVarType(that, vbObject) Then
                Set p(where)(what) = that
            Else
                p(where)(what) = that
            End If
        Case vbString
            If wt = vbString Then
                p(where)(what) = that
            Else
                mError = 5030
            End If
        Case Else
            If wt = vbString Then
                mError = 5020
            Else
                p(where)(what) = that
            End If
        End Select
        If Err Then mError = Err.Number: Err.Clear
    End If
ElseIf Not IsMissing(where) Then
    ihavearray = True
    If where < 0 Then
        p = that
        isok = True
    Else
        p = Array()
        ReDim p(where)
         isok = True
        GoTo entry1
    End If
End If
End Property
Property Get Value(Optional where, Optional what)
Attribute Value.VB_UserMemId = 0
If ihavearray Then
    If isok Then
    If IsMissing(where) Then
        Exit Property
    Else
        If UBound(p) < where Then
            ReDim Preserve p(where)
        End If
    End If
    Else
        ReDim Preserve p(where)
    End If
    If IsMissing(what) Or flat Then
        If flat And Not IsMissing(what) Then
            If where = 0 Then
                If VarType(p(what)) = vbObject Then
                    Set Value = p(what)
                Else
                    Value = p(what)
                End If
            Else
                Select Case VarType(p(where))
                
                Case vbObject
                    Set Value = p(where)
                Case VbLongLong
                    CopyMemory ByVal ((VarPtr(Value) Xor &H80000000) + 8) Xor &H80000000, ByVal ((ArrPtr() Xor &H80000000) + 8 * CLng(where)) Xor &H80000000, 8
                    PutMem2 ByVal VarPtr(Value), VbLongLong
                Case Else
                    Value = p(where)
                End Select
            End If
        Else
                Select Case VarType(p(where))
                Case vbObject
                    Set Value = p(where)
                Case VbLongLong
                    CopyMemory ByVal ((VarPtr(Value) Xor &H80000000) + 8) Xor &H80000000, ByVal ((ArrPtr() Xor &H80000000) + 8 * CLng(where)) Xor &H80000000, 8
                    PutMem2 ByVal VarPtr(Value), VbLongLong
                Case Else
                    Value = p(where)
                End Select
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
                    ElseIf myVarType(p(where)(what), vbObject) Then
                     Set Value = p(where)(what)
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
    mError = 0
End Property
Property Get count(Optional m)
count = 0&
If isok Then
    If IsMissing(m) Then
        count = Elements(p)
    
    ElseIf myVarType(p, vbObject) Then
        If p Is Nothing Then
        ElseIf TypeOf p Is refArray Then
            If m >= 0 And m <= p.count Then
                On Error Resume Next
                count = p.count(m)
                Err.Clear
            End If
        End If
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
If size < 0 Then size = 0
Dim aPtr As Long, bb As Byte
Dim Dummy0() As Object
Dim Dummy1() As String
Dim Dummy2() As Long
Dim Dummy3() As Currency
Dim Dummy4() As Variant
Dim Dummy5() As Double
Dim Dummy6() As Single
Dim Dummy7() As Byte
Dim Dummy8() As Integer
Dim Dummy9() As Boolean
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
    If where >= 0 Then
    PutMem2 VarPtr(p(where)), vbDecimal + vbArray
    Else
    PutMem2 VarPtr(p), vbDecimal + vbArray
    End If
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
    If where >= 0 Then
    PutMem2 VarPtr(p(where)), vbArray + VbLongLong
    Else
    PutMem2 VarPtr(p), vbArray + VbLongLong
    End If
Case vbVariant
    ReDim Dummy4(size)
    Value(where) = Dummy4()
Case vbBoolean
    ReDim Dummy9(size)
    Value(where) = Dummy9()
End Select
If where > 0 Then OneDimension = False
End Sub
Property Get ItemType(item) As Long
    If ihavearray And isok Then
        If Not IsMissing(item) Then
            Dim vt As Integer
                GetMem2 VarPtr(p(item)), vt
                ItemType = vt
           
        Else
            ItemType = VarType(p)
        End If
    End If
End Property
Public Function IsInnerRefArray(item, R As refArray) As Boolean
 If ihavearray And isok Then
     
            Dim vt As Integer, pp As Object
                If Me.MarkTwoDimension Then
                If Elements(p) > item Then
                    GetMem2 VarPtr(p(item)), vt
                    If vt = 9 Then
                        Set pp = p(item)
                    If Not pp Is Nothing Then
                    If TypeOf pp Is refArray Then
                        Set R = pp
                        IsInnerRefArray = True
                        End If
                        End If
                    End If
                End If
                ElseIf Elements(p(0)) > item Then
                    GetMem2 VarPtr(p(0)(item)), vt
                    If vt = 9 Then
                        Set pp = p(0)(item)
                    If Not pp Is Nothing Then
                    If TypeOf pp Is refArray Then
                        Set R = pp
                        IsInnerRefArray = True
                        End If
                        End If
                    End If
            End If
        End If

End Function

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
Property Get vtType(Optional where) As Integer
If ihavearray Then
    If isok Then
        If IsMissing(where) Or flat Then
            If Not IsMissing(where) Then
                If VarType(p) = vbArray + vbVariant Then
                    GetMem2 VarPtr(p(where)), vtType
                Else
                    vtType = VarType(p(where))
                End If
            Else
                GetMem2 VarPtr(p), vtType
            End If
        Else
            Dim pp As Long
            If Elements(p) >= where Then
            pp = Elements(p(where))
            If pp > 0 Then
                GetMem2 VarPtr(p(where)), vtType
                vtType = vtType - vbArray
            Else
                GetMem2 VarPtr(p(where)), vtType ' not an array ??
            End If
            Else
            vtType = -1
            End If
        End If
    End If
End If
End Property
Property Get MarkTwoDimension()
    If flat Then
        MarkTwoDimension = False
    Else
        MarkTwoDimension = Not CVar(OneDimension)
    End If
End Property
Property Let MarkTwoDimension(RHS)
    If Not flat Then
        OneDimension = Not CBool(RHS)
    End If
End Property
Private Sub Class_Initialize()
OneDimension = True
End Sub
Public Sub ResetToType(thistype As Integer, Optional many)
OneDimension = True
flat = thistype <> 12
If Not IsMissing(many) Then
DefArrayAt -1, thistype, Abs(CLng(many))
Else
DefArrayAt -1, thistype, 0
End If
End Sub
Public Sub RedimForFlat(size As Long)
    If ihavearray Then
    If isok Then
    If flat Then
        ReDim Preserve p(size)
    End If
    End If
    End If
End Sub
