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
Private Const vbLongLong = 20

Private Type SAFEARRAYBOUND
    cElements                   As Long
    lLbound                     As Long
End Type
Private Declare Function HashData Lib "shlwapi" (ByVal straddr As Long, ByVal ByteSize As Long, ByVal res As Long, ByVal ressize As Long) As Long
Private mError As Long
Private p As Variant
Private ihavearray As Boolean, isok As Boolean
Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal saPtr As Long, ByVal saBound As Long) As Long
Private OneDimension As Boolean
Public flat As Boolean
Friend Sub CreateRef(ByVal a As Long)
    If flat Then
        VarByRef a, p
    Else
        VarByRef a, p(0)
    End If
End Sub
Private Sub VarByRef(ByVal a As Long, ByRef b As Variant)
Dim t(0 To 3) As Long
   CopyMemory t(0), ByVal VarPtr(b), 16
   t(0) = t(0) Or &H4000
   t(2) = VarPtr(b) + 8
   CopyMemory ByVal a, t(0), 16
End Sub
Function NewRef(Optional ByRef RHS) As refArray
Dim z As New refArray
If Not IsMissing(RHS) Then z.writevalue = RHS
z.flat = flat
Set NewRef = z
End Function
Property Get Copy() As refArray
Dim z As New refArray, pp
If isok Then
        Dim i As Long, j As Long, m, rA As refArray
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
                Set rA = p(i)
                If rA.MarkTwoDimension Then
                    Set rA = rA.Copy
                    rA.MarkTwoDimension = True
                     z(i) = CVar(rA)
                Else
                   z(i) = CVar(rA.Copy)
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
Property Let Value(Optional where, Optional what, Optional WHAT1, Optional what2, that)
Dim mt As Integer, wt As Integer, ss$

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
                ElseIf VarType(p(what)) = vbLongLong Then
                    If Not VarType(that) = vbLongLong Then
                        that = cInt64(that)
                    End If
                    If Not VarType(that) = vbLongLong Then mError = 6: Exit Property
                    MemCopy ByVal ((ArrPtr() Xor &H80000000) + 8 * CLng(what)) Xor &H80000000, ((VarPtr(that) Xor &H80000000) + 8) Xor &H80000000, 8
                Else
                    p(what) = that
                End If
            Else
                If VarType(that) = vbObject Then
                    Set p(where) = that
                ElseIf VarType(p(where)) = vbLongLong Then
                    If Not VarType(that) = vbLongLong Then
                        that = cInt64(that)
                    End If
                    If Not VarType(that) = vbLongLong Then mError = 6: Exit Property
                    MemCopy ((ArrPtr() Xor &H80000000) + 8 * CLng(where)) Xor &H80000000, ((VarPtr(that) Xor &H80000000) + 8) Xor &H80000000, 8
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
                If Not VarType(that) = vbLongLong Then
                    that = cInt64(that)
                End If
                If Not VarType(that) = vbLongLong Then mError = 6: Exit Property
                MemCopy ((ArrPtr() Xor &H80000000) + 8 * CLng(where)) Xor &H80000000, ((VarPtr(that) Xor &H80000000) + 8) Xor &H80000000, 8
            ElseIf mt = 8200 Then
                If wt = vbString Then
                    p(where) = that
                Else
                    mError = 5030
                End If
            ElseIf mt = 8199 Then
                p(where) = CDate(that)
                If Err Then mError = 5040: Err.Clear
                
            ElseIf wt = vbString Then
                If mt = vbVariant Or mt = vbString Or mt = 8204 Then
                    p(where) = that
                Else
                    mError = 5020
                End If
            ElseIf wt = vbDate Then
                p(where) = CDate(that)
            ElseIf myVarType(that, vbObject) Then
                Set p(where) = that
                
            ElseIf (MemInt(VarPtr(that)) And vbArray) <> 0 Then
            
                SwapVariant p(where), that
            Else
                p(where) = that
            End If
            If Err Then mError = Err.Number: Err.Clear
        End If
    ElseIf what >= 0 Then
        If Elements(p(where)) < what + 1 Then
            Dim ThisPtr As Long, PadDim As SAFEARRAYBOUND
            PadDim.cElements = what + 1
            ThisPtr = MemLong(UnsignedAdd(VarPtr(p(where)), 8))
            If SafeArrayRedim(ByVal ThisPtr, VarPtr(PadDim)) <> 0& Then Err.Raise 5
        End If
        On Error Resume Next
            mt = vtType(where)
            wt = VarType(that)
        Select Case mt
        Case vbDataObject  ' 13 iUnKnown
            If wt = vbDataObject Then
                Set p(where)(what) = that
            ElseIf myVarType(that, vbDataObject) Then
                    Set p(where)(what) = that
            Else
                mError = 91
            End If
        Case vbObject
            If wt = vbObject Then
                Set p(where)(what) = that
            ElseIf myVarType(that, vbObject) Then
                    Set p(where)(what) = that
            Else
                mError = 91
            End If
        Case vbLongLong ' 20
                If Not VarType(that) = vbLongLong Then
                    that = cInt64(that)
                End If
                If Not VarType(that) = vbLongLong Then mError = 6: Exit Property
                MemCopy ByVal ((ArrPtr(where) Xor &H80000000) + 8 * CLng(what)) Xor &H80000000, ByVal ((VarPtr(that) Xor &H80000000) + 8) Xor &H80000000, 8
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
            ElseIf wt = vbBoolean Then
                p(where)(what) = Format$(that, DefBooleanString)
            ElseIf wt = 20 Then
                p(where)(what) = CStr(that)
            Else
                p(where)(what) = fixthis(that)
            End If
        Case vbDate
            p(where)(what) = CDate(that)
            If Err Then mError = 5040: Err.Clear
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
        p = what
        isok = True
    Else
        p = Array()
        ReDim p(where)
         isok = True
        GoTo entry1
    End If
End If
End Property
Property Get Value(Optional where, Optional what, Optional WHAT1, Optional what2)
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
                ElseIf VarType(p(what)) = vbDataObject Then
                    Set Value = p(what)
                Else
                    Value = p(what)
                End If
            Else
                Select Case VarType(p(where))
                
                Case vbObject, vbDataObject
                    Set Value = p(where)
                Case vbLongLong
                    MemCopy ((VarPtr(Value) Xor &H80000000) + 8) Xor &H80000000, ((ArrPtr() Xor &H80000000) + 8 * CLng(where)) Xor &H80000000, 8
                    MemInt(VarPtr(Value)) = vbLongLong
                Case Else
                    Value = p(where)
                End Select
            End If
        Else
                Select Case VarType(p(where))
                Case vbObject, vbDataObject ' vbDataObject is vt_IUnknown
                    Set Value = p(where)
                Case vbLongLong
                    MemCopy ((VarPtr(Value) Xor &H80000000) + 8) Xor &H80000000, ((ArrPtr() Xor &H80000000) + 8 * CLng(where)) Xor &H80000000, 8
                    MemInt(VarPtr(Value)) = vbLongLong
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
                Case vbObject, vbDataObject
                    Set Value = p(where)(what)
                Case 20
                    MemCopy ((VarPtr(Value) Xor &H80000000) + 8) Xor &H80000000, ((ArrPtr(where) Xor &H80000000) + 8 * CLng(what)) Xor &H80000000, 8
                    MemInt(VarPtr(Value)) = 20
                Case vbVariant
                    If VarType(p(where)(what)) = vbLongLong Then
                        MemCopy VarPtr(Value), ByVal ((ArrPtr(where) Xor &H80000000) + 16 * CLng(what)) Xor &H80000000, 16
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
Dim pp As Long
If ihavearray Then
    If isok Then
        If IsMissing(where) Then
            ArrPtr = MemLong(UnsignedAdd(VarPtr(p), 8))
            ArrPtr = MemLong(UnsignedAdd(ArrPtr, 12))
        Else
            If Elements(p) <= where Then
            
            ArrPtr = MemLong(UnsignedAdd(VarPtr(p(0)), 8))
            ArrPtr = MemLong(UnsignedAdd(ArrPtr, 12))
            
            ElseIf Elements(p(where)) > 0 Then
            
            ArrPtr = MemLong(UnsignedAdd(VarPtr(p(where)), 8))
            ArrPtr = MemLong(UnsignedAdd(ArrPtr, 12))
            
            End If
       
        End If
    End If
End If
End Property
Property Get SafeArrPtr(Optional where) As Long
If ihavearray Then
    If isok Then
        If IsMissing(where) Then
            SafeArrPtr = MemLong(UnsignedAdd(VarPtr(p), 8))
        Else
            Dim pp As Long
            pp = Elements(p(where))
            If pp > 0 Then
                SafeArrPtr = MemLong(UnsignedAdd(VarPtr(p(where)), 8))
            End If
        End If
    End If
End If
End Property
Function Elements(vArray As Variant) As Long
    Dim lPtr            As Long
    Const VT_BYREF      As Long = &H4000
    If IsArray(vArray) Then
    lPtr = MemLong(UnsignedAdd(VarPtr(vArray), 8))
    If (MemLong(VarPtr(vArray)) And VT_BYREF) <> 0 Then
        lPtr = MemLong(lPtr)
    End If
    If lPtr <> 0 Then
        Elements = MemLong(UnsignedAdd(lPtr, 16))
    End If
    End If
End Function
Public Function Elements2(vArray As Variant) As Long
    Dim lPtr            As Long
    Const VT_BYREF      As Long = &H4000
    If IsArray(vArray) Then
    lPtr = MemLong(UnsignedAdd(VarPtr(vArray), 8))
    If (MemLong(VarPtr(vArray)) And VT_BYREF) <> 0 Then
        lPtr = MemLong(lPtr)
    End If
    If lPtr <> 0 Then
        Dim i As Long, p As Long
        p = 1
        For i = 1 To MemInt(lPtr)
            p = p * MemLong(UnsignedAdd(lPtr, 8 + i * 8))
        Next i
        Elements2 = p
    End If
    End If
End Function
Public Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
Public Sub DefArrayAt(where, Arrtype, Optional Size As Long = 10, Optional Size1 As Long, Optional Size2 As Long)
If Size < 0 Then Size = 0
If ihavearray Then
    If isok Then
        If UBound(p) < where Then
            ReDim Preserve p(where)
        End If
    Else
        ReDim p(where)
    End If
Else
    ihavearray = True
        p = Array()
        If where = -1 Then
        ReDim p(0)
        Else
        ReDim p(where)
        End If
         isok = True

End If
Dim pp

If where = -1 Then
    pp = LibMemory.EmptyArray(1, Arrtype)
    ReDim pp(Size)
    SwapVariant p, pp
Else
    If Size1 <= 0 Then
        pp = LibMemory.EmptyArray(1, Arrtype)
        ReDim pp(Size)
    ElseIf Size2 <= 0 Then
        pp = LibMemory.EmptyArray(2, Arrtype)
        ReDim pp(Size, Size1)
    Else
        pp = LibMemory.EmptyArray(3, Arrtype)
        ReDim pp(Size, Size1, Size2)
    End If
    SwapVariant p(where), pp
End If
ihavearray = True
isok = True
If where > 0 Then OneDimension = False
End Sub
Property Get ItemType(item) As Long
On Error GoTo alfa
    If ihavearray And isok Then
        If Not IsMissing(item) Then
        
            ItemType = MemInt(VarPtr(p(item)))
        Else
            ItemType = VarType(p)
        End If
    End If
    Exit Property
alfa:
    ItemType = -1
End Property
Public Function IsInnerRefArray(item, R As refArray) As Boolean
 If ihavearray And isok Then
            Dim pp As Object
                If Me.MarkTwoDimension Then
                If Elements(p) > item Then
                    If MemInt(VarPtr(p(item))) = 9 Then
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
                    If MemInt(VarPtr(p(0)(item))) = 9 Then
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
Property Get Hash(item As Long, Optional what As Long = -1, Optional ByVal Size As Long = 4) As Long
    Dim mPtr As Long, ssize As Long
    If Size < 1 Then Exit Property
    If ihavearray And isok Then
    
    Select Case VarType(Value(item))
    Case 8200
        If what < 0 Or what > UBound(p(item)) Then
            Err.Raise 5
        Else
            mPtr = MemLong(UnsignedAdd(VarPtr(p(item)), 8))
            mPtr = MemLong(UnsignedAdd(mPtr, 12))
            mPtr = MemLong(mPtr + 4 * what)
            If mPtr = 0 Then
                Hash = 1
            Else
                ssize = MemLong(mPtr - 4)
                HashData mPtr, ssize, VarPtr(Hash), Size
            End If
        End If
    End Select
    End If
End Property
Friend Function CompareStrEq(where As Long, what As Long, thisstr As String) As Boolean
        Const CSTR_EQUAL            As Long = 2&
        Dim mPtr As Long, ssize As Long, lPtr As Long
        mPtr = MemLong(((VarPtr(p(where)) Xor &H80000000) + 8) Xor &H80000000)
        ''SafeArrayAccessData mPtr, iPtr
        mPtr = MemLong(((mPtr Xor &H80000000) + 12) Xor &H80000000)
        mPtr = MemLong(mPtr + 4 * what)
            
            If mPtr = 0 Then
                GoTo ex1
            Else
                ssize = MemLong(mPtr - 4)
                ssize = ssize \ 2
                If ssize <> Len(thisstr) Then Exit Function 'GoTo ex1
                If ssize = 0 Then CompareStrEq = True: Exit Function 'GoTo ex1
                lPtr = StrPtr(thisstr)
                While ssize > 0
                    If MemInt(lPtr) <> MemInt(mPtr) Then Exit Function
                    ssize = ssize - 1
                    lPtr = ((lPtr Xor &H80000000) + 2) Xor &H80000000
                    mPtr = ((mPtr Xor &H80000000) + 2) Xor &H80000000
                Wend
                CompareStrEq = True
            End If
ex1:
''SafeArrayUnaccessData mPtr

End Function
Private Function VarTypeNameList(v, Optional addstr As String) As String
Dim n As Integer
    n = MemInt(VarPtr(v))
    Select Case n
        Case 0
            VarTypeNameList = "Empty"
        Case 1
            VarTypeNameList = "Null"
        Case 2
            VarTypeNameList = "Integer"
        Case 3
            VarTypeNameList = "Long"
        Case 4
            VarTypeNameList = "Single"
        Case 5
            VarTypeNameList = "Double"
        Case 6
            VarTypeNameList = "Currency"
        Case 7
            VarTypeNameList = "Date"
        Case 8
            VarTypeNameList = "String"
        Case 10
            VarTypeNameList = "Error"
        Case 9, 13
            VarTypeNameList = Typename(v)
        Case 11
            VarTypeNameList = "Boolean"
        Case 12, 1
            VarTypeNameList = "Variant"
        Case 14
            VarTypeNameList = "Decimal"
        Case 17
            VarTypeNameList = "Byte"
        Case 20
            VarTypeNameList = "Long Long"
        Case 36
            VarTypeNameList = Typename(v)
        Case Is > 8192
            VarTypeNameList = ArrayDesc(v, addstr)
        Case Else
            VarTypeNameList = "type" & VarType(v)
    End Select
End Function
Public Function ArrayDesc(vArray As Variant, str As String) As String
    Dim lPtr            As Long
    Const VT_BYREF      As Long = &H4000
    If IsArray(vArray) Then
        lPtr = MemLong(UnsignedAdd(VarPtr(vArray), 8))
        If (MemLong(VarPtr(vArray)) And VT_BYREF) <> 0 Then
            lPtr = MemLong(lPtr)
        End If
        If lPtr <> 0 Then
            Dim i As Long, p As String, aType As String
            Select Case MemInt(VarPtr(vArray)) And &HFFF
            Case 2
                aType = "Integer"
            Case 3
                aType = "Long"
            Case 4
                aType = "Single"
            Case 5
                aType = "Double"
            Case 6
                aType = "Currency"
            Case 7
                aType = "Date"
            Case 8
                aType = "String"
            Case 9, 13  '   VT_DISPATCH = 9, VT_UNKNOWN = 13
                aType = "Object"
            Case 11
                aType = "Boolean"
            Case 12, 1
                aType = "Variant"
            Case 14
                aType = "Decimal"
            Case 17 ' VT_UI1 = 17
                aType = "Byte"
            Case 20
                aType = "Long Long"
            Case 27 '  VT_SAFEARRAY = 27
                aType = "Array"
            Case 29 ' VT_USERDEFINED = 29
                aType = "UDT"
            Case 36
                aType = Typename(vArray)
            End Select
            For i = 1 To MemInt(lPtr)
                p = p & "[" & MemLong(UnsignedAdd(lPtr, 8 + i * 8)) & "]"
            Next i
            ArrayDesc = aType & str & p
            Exit Function
        End If
        ArrayDesc = "EmptyArray"
        Exit Function
    End If
    ArrayDesc = "type" & VarType(vArray)
End Function
Property Get vtTypeString(Optional where) As String
If ihavearray Then
    If isok Then
        If IsMissing(where) Or flat Then
            If Not IsMissing(where) Then
                vtTypeString = VarTypeNameList(p(where))
            Else
                If MarkTwoDimension Then
                    vtTypeString = VarTypeNameList(p(0), Mid$(VarTypeNameList(p), 8))
                Else
                    vtTypeString = VarTypeNameList(p)
                End If
            End If
        Else
            vtTypeString = VarTypeNameList(p(where))
        End If
    End If
End If
End Property
Property Get vtType(Optional where) As Integer
If ihavearray Then
    If isok Then
        If IsMissing(where) Or flat Then
            If Not IsMissing(where) Then
                If VarType(p) = vbArray + vbVariant Then
                    vtType = MemInt(VarPtr(p(where)))
                Else
                    vtType = VarType(p(where))
                End If
            Else
                vtType = MemInt(VarPtr(p))
                
            End If
        Else
            Dim pp As Long
            If Elements(p) >= where Then
            On Error Resume Next
            pp = Elements(p(where))
            If Err Then

            vtType = -1
            Exit Property
            End If
            If pp > 0 Then
                vtType = MemInt(VarPtr(p(where))) - vbArray
            Else
                vtType = MemInt(VarPtr(p(where))) ' not an array ??
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
Public Sub RedimForFlat(Size As Long)
    If ihavearray Then
    If isok Then
    If flat Then
        ReDim Preserve p(Size)
    End If
    End If
    End If
End Sub
Public Function myVarType(z, j As Integer) As Boolean
    myVarType = MemInt(VarPtr(z)) = j
End Function
Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
   Static t(0 To 3) As Long  ' 4 Longs * 4 bytes each = 16 bytes
   MemCopy VarPtr(t(0)), ByVal VarPtr(a), 16
   MemCopy VarPtr(a), VarPtr(b), 16
   MemCopy VarPtr(b), VarPtr(t(0)), 16
End Sub
Public Function cInt64(p)
    Static maxlonglong, limitlonglong, OneLongLong, OneBigLongLong
    Dim a, i As Integer
    If MemInt(VarPtr(maxlonglong)) = 0 Then
        maxlonglong = CDec("18446744073709551616")
        limitlonglong = CDec("9223372036854775808")
        MemInt(VarPtr(OneLongLong)) = 20
        MemByte(VarPtr(OneLongLong) + 8) = 1
        MemInt(VarPtr(OneBigLongLong)) = 20
        MemByte(VarPtr(OneBigLongLong) + 12) = 1
    End If
    i = MemInt(VarPtr(p))
    Select Case i
    Case vbDecimal
        a = Fix(p)
        a = a - Int(a / maxlonglong) * maxlonglong
        If a < -limitlonglong - 1 Then
            While a <= -limitlonglong - 1: a = a + maxlonglong: Wend
        End If
        While a >= limitlonglong: a = a - maxlonglong: Wend
        cInt64 = -OneLongLong And a
    Case 20
        cInt64 = p
    Case vbLong, vbInteger
        cInt64 = -OneLongLong And p
    Case Else
        On Error GoTo er1
        
        a = Fix(CDec(p))
        If a > limitlonglong Or a <= -limitlonglong Then
        a = a - Int(a / (maxlonglong)) * (maxlonglong)
        If a <= -limitlonglong - 1 Then
            While a <= -limitlonglong - 1: a = a + maxlonglong: Wend
        End If
        While a >= limitlonglong: a = a - maxlonglong: Wend
        End If
        cInt64 = -OneLongLong And a
        If i = vbString Then
            If Left$(p, 1) = "&" And a < 0 Then
            Select Case Len(p)
            Case 10
                If InStr("89ABCDEF", UCase(Mid$(p, 3, 1))) > 0 Then
                cInt64 = OneBigLongLong + cInt64
                End If
            Case 18
                If Mid$(p, 3, 8) = "00000000" Then
                    cInt64 = OneBigLongLong + cInt64
                End If
            Case 11 To 17
            If cInt64 < 0 Then cInt64 = OneBigLongLong + cInt64
            End Select
            End If
         End If
    End Select
    Exit Function
er1:
    cInt64 = OneLongLong - OneLongLong
End Function
Private Function fixthis(p As Variant) As String
        fixthis = LTrim$(str(p))
        If Left$(fixthis, 1) = "." Then
        fixthis = "0" + fixthis
        ElseIf Left$(fixthis, 2) = "-." Then
        fixthis = "-0" + Mid$(fixthis, 2)
        End If
        If InStr(fixthis, ".") > 0 Then
        If NoUseDec Then fixthis = Replace(fixthis, ".", NowDec$)
        End If
End Function
