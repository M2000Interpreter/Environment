VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const vbLongLong = 20
Private maxlonglong, limitlonglong
Private Declare Function HashData Lib "shlwapi" (ByVal straddr As Long, ByVal ByteSize As Long, ByVal res As Long, ByVal ressize As Long) As Long
Private HasHandlers As Boolean
Private mem As New refArray
Private hashsize As Integer
Private p As Long, nextVal As Long, datasize As Long, hashmod As Long, lastDataSize As Long
Private Const keyArray = 1, DataArray = 2, HashArray = 0, FirstKey = 3, Pleft = 4, HashSizeStore = 5, nextValStore = 6
Sub Prepare(ByVal Prefdatasize As Long, Optional HandlersOnly)  ' Optional
If mem.Elements(p) = 0 Then
datasize = Prefdatasize
lastDataSize = datasize
HasHandlers = Not IsMissing(HandlersOnly)
hashmod = CLng(datasize * 1.5)
    If datasize < 150 Then
        hashsize = 1
    ElseIf datasize < 21500 Then
        hashsize = 2
    Else
        hashsize = 4
    End If
End If
End Sub
Function MyIsNumeric(v As Variant) As Boolean
Dim n As Byte
n = MemByte(VarPtr(v))
If n < 2 Then Exit Function
If n <= 5 Then MyIsNumeric = True
End Function
Function Add(val, bkey, Optional thistype) As Long
If HasHandlers Then Err.Raise "Only Handlers Support"
Dim akey
If MyIsNumeric(bkey) Then bkey = CLng(bkey)
If VarType(bkey) <> vbLong And VarType(bkey) <> vbString Then Err.Raise 5
    If mem.Count = 0 Then
        If datasize = 0 Then datasize = 10: lastDataSize = datasize: hashmod = CLng(datasize * 1.5)
        If hashsize = 0 Then hashsize = 1
        Select Case hashsize
        Case 1
            mem.DefArrayAt Pleft, vbByte, datasize
            mem.DefArrayAt FirstKey, vbByte, datasize
            mem.DefArrayAt HashArray, vbByte, hashmod
        Case 2
            mem.DefArrayAt Pleft, vbInteger, datasize
            mem.DefArrayAt FirstKey, vbInteger, datasize
            mem.DefArrayAt HashArray, vbInteger, hashmod
        Case 4
            mem.DefArrayAt Pleft, vbLong, datasize
            mem.DefArrayAt FirstKey, vbLong, datasize
            mem.DefArrayAt HashArray, vbLong, hashmod
        End Select
        mem.DefArrayAt keyArray, vbString, datasize
        mem(HashSizeStore) = hashsize
        If IsMissing(thistype) Then
        Select Case VarType(val)
        Case vbLong, vbInteger, vbSingle, vbCurrency, vbString, vbByte, vbDate, vbLongLong, vbDecimal, vbObject, vbDataObject
            mem.DefArrayAt DataArray, VarType(val), datasize
        Case Else
            mem.DefArrayAt DataArray, vbVariant, datasize
        End Select
        Else
            mem.DefArrayAt DataArray, thistype, datasize
        End If
        nextVal = -1
    ElseIf datasize = 0 Then
        datasize = mem.Count(Pleft)
        lastDataSize = datasize
        hashmod = mem.Count(HashArray)
        hashsize = mem(HashSizeStore)
        nextVal = mem(nextValStore)
    End If
    If nextVal >= datasize Then
        datasize = datasize * 2
        lastDataSize = datasize
        Select Case hashsize
        Case 1
            If datasize < 150 Then
                ReHashByte
            Else
                ReHash2
            End If
        Case 2
            If datasize < 21500 Then
                ReHashInteger
            Else
                ReHash4
            End If
        Case Else
            ReHashLong
        End Select
    End If
    nextVal = nextVal + 1
    mem(nextValStore) = nextVal
    
    If VarType(bkey) = vbLong Then
        akey = space$(2)
        MemLong(StrPtr(akey)) = bkey
    Else
        SwapVariant akey, bkey
    End If
    mem(keyArray, nextVal) = akey
    mem(DataArray, nextVal) = val
    Dim pv As Long, pv1 As Integer, pv2 As Byte
    pv = mem.Hash(keyArray, nextVal, hashsize)
    Select Case hashsize
    Case 1
        pv2 = MemByte(VarPtr(pv)) And CByte(127)
        mem(FirstKey, nextVal) = pv2
        pv = pv2 Mod hashmod
    Case 2
        pv1 = MemInt(VarPtr(pv)) And &H7FFF
        mem(FirstKey, nextVal) = pv1
        pv = pv1 Mod hashmod
    Case Else
        pv = pv And &H7FFFFFFF
        mem(FirstKey, nextVal) = pv
        pv = pv Mod hashmod
    End Select
    If pv = 0 Then pv = 1
    Dim K
    K = mem(HashArray, pv)
    If Not K = nextVal + 1 Then
        mem(HashArray, pv) = nextVal + 1
        mem(Pleft, nextVal) = K
        Add = nextVal
    End If
End Function
Function AddHandler(akey As String) As Long
    If mem.Count = 0 Then
        HasHandlers = True
        If datasize = 0 Then datasize = 10: lastDataSize = datasize: hashmod = CLng(datasize * 1.5)
        If hashsize = 0 Then hashsize = 1
        Select Case hashsize
        Case 1
            mem.DefArrayAt Pleft, vbByte, datasize
            mem.DefArrayAt FirstKey, vbByte, datasize
            mem.DefArrayAt HashArray, vbByte, hashmod
        Case 2
            mem.DefArrayAt Pleft, vbInteger, datasize
            mem.DefArrayAt FirstKey, vbInteger, datasize
            mem.DefArrayAt HashArray, vbInteger, hashmod
        Case 4
            mem.DefArrayAt Pleft, vbLong, datasize
            mem.DefArrayAt FirstKey, vbLong, datasize
            mem.DefArrayAt HashArray, vbLong, hashmod
        End Select
        mem.DefArrayAt keyArray, vbString, datasize
        mem(HashSizeStore) = hashsize
        nextVal = -1
    ElseIf datasize = 0 Then
        HasHandlers = True
        mem(DataArray) = Empty
        lastDataSize = datasize
        datasize = mem.Count(Pleft)
        hashmod = mem.Count(HashArray)
        hashsize = mem(HashSizeStore)
        nextVal = mem(nextValStore)
    Else
        If Not HasHandlers Then Err.Raise "Use Prepare first to use Handlers"
    End If
    
    If nextVal >= datasize Then
        datasize = datasize * 2
        lastDataSize = datasize
        Select Case hashsize
        Case 1
            If datasize < 150 Then
                ReHashByte
            Else
                ReHash2
            End If
        Case 2
            If datasize < 21500 Then
                ReHashInteger
            Else
                ReHash4
            End If
        Case Else
            ReHashLong
        End Select
    End If
    nextVal = nextVal + 1
    mem(nextValStore) = nextVal
    mem(keyArray, nextVal) = akey
    Dim pv As Long, pv1 As Integer, pv2 As Byte
    pv = mem.Hash(keyArray, nextVal, hashsize)
    Select Case hashsize
    Case 1
        pv2 = MemByte(VarPtr(pv)) And CByte(127)
        mem(FirstKey, nextVal) = pv2
        pv = pv2 Mod hashmod
    Case 2
        pv1 = MemInt(VarPtr(pv)) And &H7FFF
        mem(FirstKey, nextVal) = pv1
        pv = pv1 Mod hashmod
    Case Else
        pv = pv And &H7FFFFFFF
        mem(FirstKey, nextVal) = pv
        pv = pv Mod hashmod
    End Select
    'If pV = 0 Then pV = 1
    Dim K
    K = mem(HashArray, pv)
    If Not K = nextVal + 1 Then
        mem(HashArray, pv) = nextVal + 1
        mem(Pleft, nextVal) = K
        AddHandler = nextVal
    End If
End Function

Function Find(bkey, Optional v, Optional where, Optional queue) As Boolean
    If HasHandlers Then Err.Raise "Only Handlers Support"
    Dim findfirst As Boolean
    findfirst = True
    If datasize = 0 Then
        datasize = mem.Count(Pleft)
        lastDataSize = datasize
        hashmod = mem.Count(HashArray)
        hashsize = mem(HashSizeStore)
        nextVal = mem(nextValStore)
    End If
    If Not IsMissing(queue) Then
    If Not IsNumeric(queue) Then Exit Function
    
    If queue > 0 And queue < nextVal Then findfirst = False
    End If
    
    If VarType(bkey) <> vbLong And VarType(bkey) <> vbString Then Err.Raise 5
    
    Dim akey As String
   
    If VarType(bkey) = vbLong Then
        akey = space$(2)
        
        MemLong(StrPtr(akey)) = bkey
    Else
        akey = bkey
    End If
    
    Dim K As Long
    If findfirst Then
        
        
        Dim pv As Long, pv1 As Integer, pv2 As Byte
        pv = HashkEY(akey, hashsize)
        If pv = 0 Then Exit Function
        Select Case hashsize
        Case 1
            pv2 = MemByte(VarPtr(pv)) And CByte(127)
            pv = pv2 Mod hashmod
        Case 2
            pv1 = MemInt(VarPtr(pv)) And &H7FFF
            pv = pv1 Mod hashmod
        Case Else
            pv = pv And &H7FFFFFFF
            pv = pv Mod hashmod
        End Select
        If pv = 0 Then pv = 1
        K = mem(HashArray, pv)
        If K = 0 Then Exit Function
        K = K - 1
       
        If mem.CompareStrEq(keyArray, K, akey) Then
        
            If Not IsMissing(v) Then
                If mem.vtType(DataArray, K) = 9 Then
                    Set v = mem.Value(DataArray, K)
                Else
                    v = mem(DataArray, K)
                End If
            End If
            If Not IsMissing(where) Then where = K
            If Not IsMissing(queue) Then queue = mem(Pleft, K)
            Find = True
        Else
            
            K = mem(Pleft, K) - 1
jump1:
            Do While K >= 0
                If mem.CompareStrEq(keyArray, K, akey) Then
                    If Not IsMissing(v) Then
                        If mem.vtType(DataArray, K) = 9 Then
                            Set v = mem.Value(DataArray, K)
                        Else
                            v = mem(DataArray, K)
                        End If
                    End If
                    If Not IsMissing(where) Then where = K
                    If Not IsMissing(queue) Then queue = mem(Pleft, K)
                    Find = True
                    Exit Function
                End If
                queue = Empty
                K = mem(Pleft, K) - 1
            Loop
        End If
    Else
    K = CLng(queue) - 1
    GoTo jump1
    End If
End Function
Function FindHandler(akey As String, hashcode As Long, Handler As Long) As Boolean
        If datasize = 0 Then
            datasize = mem.Count(Pleft)
            lastDataSize = datasize
            hashmod = mem.Count(HashArray)
            hashsize = mem(HashSizeStore)
            nextVal = mem(nextValStore)
        End If

        Dim K As Long
        Dim pv As Long, pv1 As Integer, pv2 As Byte
        pv = HashkEY(akey, hashsize)
        Select Case hashsize
        Case 1
            pv2 = MemByte(VarPtr(pv)) And CByte(127)
            pv = pv2 Mod hashmod
        Case 2
            pv1 = MemInt(VarPtr(pv)) And &H7FFF
            pv = pv1 Mod hashmod
        Case Else
            pv = pv And &H7FFFFFFF
            pv = pv Mod hashmod
        End Select
        If pv = 0 Then pv = 1
        K = mem(HashArray, pv)
        If K = 0 Then Exit Function
        K = K - 1
        
        If mem.CompareStrEq(keyArray, K, akey) Then
            Handler = K
            hashcode = mem(FirstKey, K)
            FindHandler = True
            Exit Function
        End If
End Function

Property Get Keys()
    Keys = mem(keyArray)
End Property
Property Get Values()
    Values = mem(DataArray)
End Property
Property Set Value(ByVal where As Long, RHS)
    If HasHandlers Then Exit Property
    If where < 0 Or where > mem.Count(DataArray) Then Exit Property
    mem(DataArray, where) = RHS
End Property
Property Let Value(ByVal where As Long, RHS)
    If HasHandlers Then Exit Property
    If where < 0 Or where > mem.Count(DataArray) Then Exit Property
    mem(DataArray, where) = RHS
End Property
Property Get Value(ByVal where As Long)
Attribute Value.VB_UserMemId = 0
    Dim ret
     
    If Not HasHandlers Then
        If where < 0 Or where > mem.Count(DataArray) Then Exit Property
        If mem.vtType(DataArray, where) = 9 Then
            Set ret = mem.Value(DataArray, where)
            SwapVariant ret, Value
        Else
            Value = mem(DataArray, where)
        End If
    Else
        If where < 0 Or where > nextVal Then Value = False Else Value = True
    End If
End Property
Property Let Key(ByVal where As Long, bkey)
    If HasHandlers Then Err.Raise "Only Handlers Support"
      Dim akey As String
   
    If MyIsNumeric(bkey) Then
        akey = space$(2)
        MemLong(StrPtr(akey)) = bkey
    Else
        akey = bkey
    End If
     Dim pv As Long, pv1 As Integer, pv2 As Byte
    mem(keyArray, where) = akey
    pv = mem.Hash(keyArray, where, hashsize)
    
    Select Case hashsize
    Case 1
        pv2 = MemByte(VarPtr(pv)) And CByte(127)
        mem(FirstKey, where) = pv2
    Case 2
        pv1 = MemInt(VarPtr(pv)) And &H7FFF
        mem(FirstKey, where) = pv1
    Case Else
        pv = pv And &H7FFFFFFF
        mem(FirstKey, where) = pv
    End Select
    
    Rehash
End Property
Property Let KeyNoReHash(ByVal where As Long, bkey)
    If HasHandlers Then Err.Raise "Only Handlers Support"
      Dim akey As String
   
    If MyIsNumeric(bkey) Then
        akey = space$(2)
        MemLong(StrPtr(akey)) = bkey
    Else
        akey = bkey
    End If
     Dim pv As Long, pv1 As Integer, pv2 As Byte
    mem(keyArray, where) = akey
    pv = mem.Hash(keyArray, where, hashsize)
    
    Select Case hashsize
    Case 1
        pv2 = MemByte(VarPtr(pv)) And CByte(127)
        mem(FirstKey, where) = pv2
    Case 2
        pv1 = MemInt(VarPtr(pv)) And &H7FFF
        mem(FirstKey, where) = pv1
    Case Else
        pv = pv And &H7FFFFFFF
        mem(FirstKey, where) = pv
    End Select
End Property
Property Get Key(ByVal where As Long)
    If where < 0 Or where > mem.Count(DataArray) Then Exit Property
    Key = mem(keyArray, where)
End Property
Friend Property Set PlaceRefArray(that As refArray)
    Set mem = that.Copy()
    datasize = 0
    lastDataSize = datasize
End Property
Property Get Copy() As HashList
Set Copy = New HashList
If HasHandlers Then Copy.Prepare 0, HasHandlers
Set Copy.PlaceRefArray = mem
End Property
Property Get HashkEY(Key$, Optional ByVal Size As Long = 4) As Long
        If Size < 1 Then Exit Property
        If LenB(Key$) = 0 Then
            HashkEY = 1
        Else
            HashData StrPtr(Key$), LenB(Key$), VarPtr(HashkEY), Size
        End If
End Property
Property Get Count()
    If datasize = 0 Then
        datasize = mem.Count(Pleft)
        lastDataSize = datasize
        hashmod = mem.Count(HashArray)
        hashsize = mem(HashSizeStore)
        nextVal = mem(nextValStore) = nextVal
    End If
    Count = nextVal + 1

End Property
Sub Clear()
    Set mem = New refArray
    datasize = 0
    lastDataSize = datasize
End Sub
Function KeyNumeric(akey) As Long
If VarType(akey) <> vbString Then Err.Raise 5
If Len(akey) = 2 Then
    KeyNumeric = MemLong(StrPtr(akey))
End If
End Function
Property Get KeyNum(ByVal where As Long) As Long
    If where < 0 Or where > mem.Count(DataArray) Then Exit Property
    KeyNum = KeyNumeric(mem(keyArray, where))
End Property
Property Get KeyNum2(where As Long) As Long
    If where < 0 Or where > mem.Count(DataArray) Then Exit Property
    KeyNum2 = KeyNumeric(mem(keyArray, where)) And &HFFFF&
End Property
Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
   Static t(0 To 3) As Long  ' 4 Longs * 4 bytes each = 16 bytes
   MemCopy VarPtr(t(0)), ByVal VarPtr(a), 16
   MemCopy VarPtr(a), VarPtr(b), 16
   MemCopy VarPtr(b), VarPtr(t(0)), 16
End Sub
Private Sub ReHashByte()
    mem(HashArray) = Empty
    If lastDataSize <> datasize Then ' lastDataSize = datasize when we change a key (the name of key).
        mem(FirstKey, datasize) = CByte(0)
        mem(Pleft, datasize) = CByte(0)
        mem(keyArray, datasize) = ""
    End If
    If Not HasHandlers Then
        mem(DataArray, datasize) = mem(DataArray, 0)
    End If
    hashmod = CLng(datasize * 1.5)
    mem.DefArrayAt HashArray, vbByte, hashmod
    Dim pv As Long, K As Byte, I As Byte
    For I = 0 To Count - 1
    pv = mem(FirstKey, I) Mod hashmod
    K = mem(HashArray, pv)
    'If Not k = i + 1 Then
        mem(HashArray, pv) = I + 1
        mem(Pleft, I) = K
    'End If
Next I
End Sub
Private Sub ReHashInteger()
    mem(HashArray) = Empty
    If lastDataSize <> datasize Then
        mem(FirstKey, datasize) = 0
        mem(Pleft, datasize) = 0
        mem(keyArray, datasize) = ""
    End If
    If Not HasHandlers Then
        mem(DataArray, datasize) = mem(DataArray, 0)
    End If
    hashmod = CLng(datasize * 1.5)
    mem.DefArrayAt HashArray, vbInteger, hashmod
    Dim pv As Long, K As Integer, I As Integer
    For I = 0 To Count - 1
    pv = mem(FirstKey, I) Mod hashmod
    K = mem(HashArray, pv)
    'If Not k = i + 1 Then
    mem(HashArray, pv) = I + 1
    mem(Pleft, I) = K
    'End If
Next I
End Sub
Private Sub ReHashLong()
    mem(HashArray) = Empty
    If lastDataSize <> datasize Then
    mem(FirstKey, datasize) = 0&
    mem(Pleft, datasize) = 0&
    mem(keyArray, datasize) = ""
    End If
    hashmod = CLng(datasize * 1.5)
    mem.DefArrayAt HashArray, vbLong, hashmod
    If Not HasHandlers Then
        mem(DataArray, datasize) = mem(DataArray, 0)
    End If
    Dim pv As Long, K As Long, I As Long
    For I = 0 To Count - 1
    pv = mem(FirstKey, I) Mod hashmod
    K = mem(HashArray, pv)
    'If Not k = i + 1 Then  ' these was for security..
    mem(HashArray, pv) = I + 1
    mem(Pleft, I) = K
    'End If
Next I
End Sub
Private Sub ReHash2()
    Dim pv As Long, K As Integer, I As Integer
    Dim pv1 As Integer
    hashsize = 2
    mem(HashSizeStore) = hashsize
    mem(HashArray) = Empty
    mem(FirstKey) = Empty
    mem(Pleft) = Empty
    hashmod = CLng(datasize * 1.5)
    mem.DefArrayAt Pleft, vbInteger, datasize
    mem.DefArrayAt FirstKey, vbInteger, datasize
    mem.DefArrayAt HashArray, vbInteger, hashmod
    mem(keyArray, datasize) = ""
    If Not HasHandlers Then
        mem(DataArray, datasize) = mem(DataArray, 0)
    End If
    For I = 0 To Count - 1
        p = mem.Hash(keyArray, CLng(I), hashsize)
        pv1 = MemInt(VarPtr(p)) And &H7FFF
        
        mem(FirstKey, I) = pv1
        pv = pv1 Mod hashmod
        K = mem(HashArray, pv)
        'If Not k = i + 1 Then
        mem(HashArray, pv) = I + 1
        mem(Pleft, I) = K
        'End If
    Next I
End Sub
Private Sub ReHash4()
    Dim pv As Long, K As Long, I As Long
    Dim pv2 As Long
    hashsize = 4
    mem(HashSizeStore) = hashsize
    mem(HashArray) = Empty
    mem(FirstKey) = Empty
    mem(Pleft) = Empty
    hashmod = CLng(datasize * 1.5)
    mem.DefArrayAt Pleft, vbLong, datasize
    mem.DefArrayAt FirstKey, vbLong, datasize
    mem.DefArrayAt HashArray, vbLong, hashmod
    mem(keyArray, datasize) = ""
    If Not HasHandlers Then
        mem(DataArray, datasize) = mem(DataArray, 0)
    End If
    For I = 0 To Count - 1
        pv2 = mem.Hash(keyArray, I, hashsize)
        pv2 = pv2 And &H7FFFFFFF
        mem(FirstKey, I) = pv2
        pv = pv2 Mod hashmod
        K = mem(HashArray, pv)
        'If Not k = i + 1 Then
        mem(HashArray, pv) = I + 1
        mem(Pleft, I) = K
        'End If
    Next I
End Sub
Sub Rehash()
If Count > 0 Then
    Select Case hashsize
    Case 1
        ReHashByte
    Case 2
        ReHashInteger
    Case Else
        ReHashLong
    End Select
End If
End Sub
Function Initialized(p) As Boolean
    Initialized = mem.Elements(p) > 0
End Function
Sub SwapStrings(a$, b$)
    Dim I As Long, j As Long
    I = MemLong(VarPtr(a$))
    j = MemLong(VarPtr(b$))
    MemLong(VarPtr(a$)) = j
    MemLong(VarPtr(b$)) = I
End Sub
Public Function create(ParamArray v()) As HashList
    Dim hl As New HashList, I As Long
    Dim pp() As Variant
    pp = v
    If mem.Elements(pp) > 0 Then
        For I = 0 To mem.Elements(pp) - 1
            hl.Add CVar(pp(I)), I + 1, vbVariant
        Next I
    End If
    Set create = hl
End Function
Public Function cInt64(p)
    Static maxlonglong, limitlonglong, OneLongLong, OneBigLongLong
    Dim a, I As Integer
    If MemInt(VarPtr(maxlonglong)) = 0 Then
        maxlonglong = CDec("18446744073709551616")
        limitlonglong = CDec("9223372036854775808")
        MemInt(VarPtr(OneLongLong)) = 20
        MemByte(VarPtr(OneLongLong) + 8) = 1
        MemInt(VarPtr(OneBigLongLong)) = 20
        MemByte(VarPtr(OneBigLongLong) + 12) = 1
    End If
    I = MemInt(VarPtr(p))
    Select Case I
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
        If I = vbString Then
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

