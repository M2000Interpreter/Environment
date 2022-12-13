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
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, retval As Byte)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, retval As Long)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private mem As New RefArray
Private hashsize As Integer
Private p() As Byte, nextVal As Long, datasize As Long, hashmod As Long
Private Const keyArray = 1, DataArray = 2, HashArray = 0, FirstKey = 3, Pleft = 4, HashSizeStore = 5, nextValStore = 6
Sub Prepare(Prefdatasize As Long)  ' Optional
If mem.Elements(p) = 0 Then
datasize = Prefdatasize
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
GetMem1 VarPtr(v), n
If n < 2 Then Exit Function
If n < 5 Then MyIsNumeric = True
End Function
Sub Add(val, bkey, Optional thistype)
Dim aKey
If MyIsNumeric(bkey) Then bkey = CLng(bkey)
If VarType(bkey) <> vbLong And VarType(bkey) <> vbString Then Err.Raise 5
    If mem.Count = 0 Then
        If datasize = 0 Then datasize = 10: hashmod = CLng(datasize * 1.5)
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
        Case vbLong, vbInteger, vbSingle, vbCurrency, vbString, vbByte
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
        hashmod = mem.Count(HashArray)
        hashsize = mem(HashSizeStore)
        nextVal = mem(nextValStore)
    End If
    If nextVal >= datasize Then
        datasize = datasize * 2
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
        aKey = Space$(2)
        PutMem4 StrPtr(aKey), bkey
    Else
        SwapVariant aKey, bkey
    End If
    mem(keyArray, nextVal) = aKey
    mem(DataArray, nextVal) = val
    p() = mem.Hash(keyArray, nextVal, hashsize)
    Dim pV As Long, pv1 As Integer, pv2 As Byte
    Select Case hashsize
    Case 1
        GetMem1 VarPtr(p(0)), pv2
        pv2 = pv2 And CByte(127)
        mem(FirstKey, nextVal) = pv2
        pV = pv2 Mod hashmod
    Case 2
        GetMem2 VarPtr(p(0)), pv1
        pv1 = pv1 And &H7FFF
        mem(FirstKey, nextVal) = pv1
        pV = pv1 Mod hashmod
    Case Else
        GetMem4 VarPtr(p(0)), pV
        pV = pV And &H7FFFFFFF
        mem(FirstKey, nextVal) = pV
        pV = pV Mod hashmod
    End Select
    'If pV = 0 Then pV = 1
    Dim k
    k = mem(HashArray, pV)
    If Not k = nextVal + 1 Then
        mem(HashArray, pV) = nextVal + 1
        mem(Pleft, nextVal) = k
     Else
     Stop
    End If
End Sub
Function Find(bkey, Optional v, Optional where, Optional queue) As Boolean
    Dim findfirst As Boolean
    findfirst = True
    If datasize = 0 Then
        datasize = mem.Count(Pleft)
        hashmod = mem.Count(HashArray)
        hashsize = mem(HashSizeStore)
        nextVal = mem(nextValStore)
    End If
    If Not IsMissing(queue) Then
    If Not IsNumeric(queue) Then Exit Function
    
    If queue > 0 And queue < nextVal Then findfirst = False
    End If
    
    If VarType(bkey) <> vbLong And VarType(bkey) <> vbString Then Err.Raise 5
    
    Dim aKey
   
    If VarType(bkey) = vbLong Then
        aKey = Space$(2)
        
        PutMem4 StrPtr(aKey), bkey
    Else
        aKey = bkey
    End If
    
    Dim k As Long
    If findfirst Then
        p = mem.HashkEY(aKey, hashsize)
        If mem.Elements(p) = 0 Then Exit Function
        Dim pV As Long, pv1 As Integer, pv2 As Byte
        Select Case hashsize
        Case 1
            GetMem1 VarPtr(p(0)), pv2
            pv2 = pv2 And CByte(127)
            pV = pv2 Mod hashmod
        Case 2
            GetMem2 VarPtr(p(0)), pv1
            pv1 = pv1 And &H7FFF
            pV = pv1 Mod hashmod
        Case Else
            GetMem4 VarPtr(p(0)), pV
            pV = pV And &H7FFFFFFF
            pV = pV Mod hashmod
        End Select
        If pV = 0 Then pV = 1
        k = mem(HashArray, pV)
        If k = 0 Then Exit Function
        k = k - 1
        If aKey = mem(keyArray, k) Then
            If Not IsMissing(v) Then v = mem(DataArray, k)
            If Not IsMissing(where) Then where = k
            If Not IsMissing(queue) Then queue = mem(Pleft, k)
            Find = True
        Else
            
            k = mem(Pleft, k) - 1
jump1:
            Do While k >= 0
                If aKey = mem(keyArray, k) Then
                    If Not IsMissing(v) Then v = mem(DataArray, k)
                    If Not IsMissing(where) Then where = k
                    If Not IsMissing(queue) Then queue = mem(Pleft, k)
                    Find = True
                    Exit Function
                End If
                queue = Empty
                k = mem(Pleft, k) - 1
            Loop
        End If
    Else
    k = CLng(queue) - 1
    GoTo jump1
    End If
End Function
Property Get Keys()
    Keys = mem(keyArray)
End Property
Property Get Values()
    Values = mem(DataArray)
End Property
Property Let Value(ByVal where As Long, RHS)
    If where < 0 Or where > mem.Count(DataArray) Then Exit Property
    mem(DataArray, where) = RHS
End Property
Property Get Value(ByVal where As Long)
Attribute Value.VB_UserMemId = 0
    If where < 0 Or where > mem.Count(DataArray) Then Exit Property
    Value = mem(DataArray, where)
End Property
Property Get Key(ByVal where As Long)
    If where < 0 Or where > mem.Count(DataArray) Then Exit Property
    Key = mem(keyArray, where)
End Property
Friend Property Set PlaceRefArray(that As RefArray)
    Set mem = that.Copy()
    datasize = 0
End Property
Property Get Copy() As HashList
Set Copy = New HashList
Set Copy.PlaceRefArray = mem
End Property
Property Get Count()
If mem.Elements(p) = 0 Then
    Count = 0
Else
    If datasize = 0 Then
        datasize = mem.Count(Pleft)
        hashmod = mem.Count(HashArray)
        hashsize = mem(HashSizeStore)
        nextVal = mem(nextValStore) = nextVal
    End If
    Count = nextVal + 1
End If
End Property
Sub Clear()
    Set mem = New RefArray
    datasize = 0
End Sub
Function KeyNumeric(aKey) As Long
If VarType(aKey) <> vbString Then Err.Raise 5
If Len(aKey) = 2 Then
    GetMem4 StrPtr(aKey), KeyNumeric
End If
End Function
Private Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
   Static t(0 To 3) As Long  ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(a), 16
   CopyMemory ByVal VarPtr(a), ByVal VarPtr(b), 16
   CopyMemory ByVal VarPtr(b), t(0), 16
End Sub
Private Sub ReHashByte()
    mem(HashArray) = Empty
    mem.DefArrayAt HashArray, vbByte, hashmod
    mem(FirstKey, datasize) = CByte(0)
    mem(Pleft, datasize) = CByte(0)
    mem(keyArray, datasize) = ""
    mem(DataArray, datasize) = mem(DataArray, 0)
    hashmod = CLng(datasize * 1.5)
    Dim pV As Long, k As Byte, i As Byte
    For i = 0 To Count - 1
    pV = mem(FirstKey, i) Mod hashmod
    k = mem(HashArray, pV)
    'If Not k = i + 1 Then
        mem(HashArray, pV) = i + 1
        mem(Pleft, i) = k
    'End If
Next i
End Sub
Private Sub ReHashInteger()
    mem(HashArray) = Empty
    mem.DefArrayAt HashArray, vbInteger, hashmod
    mem(FirstKey, datasize) = 0
    mem(Pleft, datasize) = 0
    mem(keyArray, datasize) = ""
    mem(DataArray, datasize) = mem(DataArray, 0)
    hashmod = CLng(datasize * 1.5)
    Dim pV As Long, k As Integer, i As Integer
    For i = 0 To Count - 1
    pV = mem(FirstKey, i) Mod hashmod
    k = mem(HashArray, pV)
    'If Not k = i + 1 Then
    mem(HashArray, pV) = i + 1
    mem(Pleft, i) = k
    'End If
Next i
End Sub
Private Sub ReHashLong()
    mem(HashArray) = Empty
    mem.DefArrayAt HashArray, vbLong, hashmod
    mem(FirstKey, datasize) = 0&
    mem(Pleft, datasize) = 0&
    hashmod = CLng(datasize * 1.5)
    mem(keyArray, datasize) = ""
    mem(DataArray, datasize) = mem(DataArray, 0)
    Dim pV As Long, k As Long, i As Long
    For i = 0 To Count - 1
    pV = mem(FirstKey, i) Mod hashmod
    k = mem(HashArray, pV)
    'If Not k = i + 1 Then
    mem(HashArray, pV) = i + 1
    mem(Pleft, i) = k
    'End If
Next i
End Sub
Private Sub ReHash2()
    Dim pV As Long, k As Integer, i As Integer
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
    mem(DataArray, datasize) = mem(DataArray, 0)
    For i = 0 To Count - 1
        p() = mem.Hash(keyArray, CLng(i), hashsize)
        GetMem2 VarPtr(p(0)), pv1
        pv1 = pv1 And &H7FFF
        mem(FirstKey, i) = pv1
        pV = pv1 Mod hashmod
        k = mem(HashArray, pV)
        'If Not k = i + 1 Then
        mem(HashArray, pV) = i + 1
        mem(Pleft, i) = k
        'End If
    Next i
End Sub
Private Sub ReHash4()
    Dim pV As Long, k As Long, i As Long
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
    mem(DataArray, datasize) = mem(DataArray, 0)
    For i = 0 To Count - 1
        p() = mem.Hash(keyArray, i, hashsize)
        GetMem4 VarPtr(p(0)), pv2
        pv2 = pv2 And &H7FFFFFFF
        mem(FirstKey, i) = pv2
        pV = pv2 Mod hashmod
        k = mem(HashArray, pV)
        'If Not k = i + 1 Then
        mem(HashArray, pV) = i + 1
        mem(Pleft, i) = k
        'End If
    Next i
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
    Dim i As Long, j As Long
    GetMem4 VarPtr(a$), i
    GetMem4 VarPtr(b$), j
    PutMem4 VarPtr(a$), j
    PutMem4 VarPtr(b$), i
End Sub
Public Function create(ParamArray v()) As HashList
    Dim hl As New HashList, i As Long
    Dim pp() As Variant
    pp = v
    If mem.Elements(pp) > 0 Then
        For i = 0 To mem.Elements(pp) - 1
            hl.Add CVar(pp(i)), i + 1, vbVariant
        Next i
    End If
    Set create = hl
End Function

