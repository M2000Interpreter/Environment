VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StructCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' No same key. --- this is changed
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
' We get this from tlb (see reference)
Private Type item  ' if you add change swapb (check with lenb() function)
    Key As String
    struct As Object
    Offset As Long
    items As Long
    sValue As Long
    firsthash As Long
    lastpos As Long
    Pleft As Long
End Type
Private PriveSpace() As item
Dim MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Private lastfind As String
Private lasttype As Integer
Private lastkey As String
Private Hash() As Long
Public Done As Boolean
Public Index As Long
Public structLen As Long
Private ParentIndex As Long
Public IamBusy As Boolean
Public useclid As Long
Const UseUcaseKeys = True
Private entrance As Long
Public Tag As String
Private Sub Class_Initialize()
    MaxSpace = 5
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
End Sub
Friend Sub MakeSpace(ThatSpace As Long)
    MaxSpace = ThatSpace
    ReDim PriveSpace(MaxSpace) As item
    toplim = -1
End Sub
Friend Sub AddMoreSpace(ThatSpace As Long)
    MaxSpace = MaxSpace + ThatSpace
    ReDim Preserve PriveSpace(MaxSpace) As item
End Sub
Public Sub AddKey2(RHS, Optional aValue As Variant)
Stop
End Sub
Public Sub AddKey(RHS, Optional aValue As Variant)
Stop
End Sub
Public Property Get StructOffset() As Long
If Index > -1 Then
    StructOffset = PriveSpace(Index).Offset
End If
End Property
Public Property Get StructMany() As Long
If Index > -1 Then
    StructMany = PriveSpace(Index).items
End If
End Property
Public Property Get StructObj() As Object
If Index > -1 Then
    Set StructObj = PriveSpace(Index).struct
End If
End Property
Friend Sub AddKeyStruct(RHS, Offset As Long, manyitems As Long, obj As Object)
    Index = -1
    lastkey = Normalize(RHS)
    If Len(lastkey) = 0 Then RHS = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0): lastkey = RHS
    Done = False
    lasttype = VarType(RHS)
    Value = Offset
    With PriveSpace(Index)
    Set .struct = obj
    .items = manyitems
    End With
End Sub


Public Sub ToKey(RHS)
    Done = Find(RHS)
    If Done Then lastkey = Normalize(RHS)
End Sub

Property Let Value(RHS As Variant)
Done = False
If Index = -1 Then
    ItemCreator lastkey, RHS
Else
    PriveSpace(Index).Offset = CLng(RHS)
End If
Done = True
End Property
Property Set ValueObj(RHS As Variant)
Done = False
End Property
Property Get sValue() As Long
Done = False
If Index = -1 Then
Else
    Done = True
    sValue = PriveSpace(Index).sValue
End If
End Property
Property Let sValue(RHS As Long)
If Index = -1 Then
Else
    Done = True
    PriveSpace(Index).sValue = RHS
End If
End Property
Property Get KeyTypeValue() As Integer
Done = False
KeyTypeValue = 8
End Property
Property Let KeyTypeValue(RHS As Integer)
If Index = -1 Then
Else
    Done = True
    If RHS <> 8 Then
        Err.Raise 13
    End If
End If
End Property

Property Get Value() As Variant
Done = False
If Index = -1 Then
Else
    Done = True
    Value = PriveSpace(Index).Offset
End If
End Property
Public Function IsString() As Boolean
End Function
Private Function MyVal() As Variant
On Error Resume Next
MyVal = PriveSpace(Index).Key
End Function
Property Get ValueObj() As Variant
    Done = False
    Set ValueObj = Nothing
End Property
Friend Sub ExpandHash()
    hashlen = MaxSpace * 2 + 3
    Rehash
End Sub
Private Sub Rehash()
    Dim i As Long
    ReDim Hash(hashlen) As Long
    For i = 0 To toplim
           place HashFunc2(i), i
        
    Next i
End Sub
Private Function CheckHash() As Boolean
    Dim i As Long
    For i = 0 To toplim
        If PriveSpace(i).Pleft > 0 Then
        If PriveSpace(PriveSpace(i).Pleft - 1).Key = vbNullString Then Exit Function
        If i = (PriveSpace(i).Pleft - 1) Then Exit Function
        End If
    Next i
    CheckHash = True
End Function

Public Function Normalize(Key) As String
    If VarType(Key) = vbString Then
            Normalize = myUcase(Key)
            If Len(Normalize) = 0 Then Err.Raise 13
    Else
        Err.Raise 13
    End If
End Function
Private Function Malloc() As Long
    If toplim + 1 >= MaxSpace Then
        MaxSpace = MaxSpace * 2
        ReDim Preserve PriveSpace(MaxSpace) As item
        If MaxSpace > hashlen * 3 / 4 Then ExpandHash
    End If
    toplim = toplim + 1
    Malloc = toplim
End Function
Property Get Count()
    Count = toplim + 1
End Property
Property Get IsEmpty()
    IsEmpty = toplim = -1
End Property
Property Get LastKnown() As String
    LastKnown = lastfind
End Property
Function IsObj() As Boolean
End Function
Function IsObjAt(curitem As Long, peekvalue) As Boolean
End Function
Function IsEnum(p As Variant) As Boolean
End Function

Function Find(RHS) As Boolean
Dim K As Long, Key As String, k1 As Long
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
ParentIndex = -1
Done = False
K = Hash(HashFunc(Key)) - 1
k1 = -2
If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then Find = True: lastfind = Key: Index = K: Done = True: Exit Function
     ParentIndex = K
     k1 = K
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0 Or k1 = K
If k1 = K Then
Rehash
K = Hash(HashFunc(Key)) - 1
k1 = -2
ParentIndex = -1
If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then Find = True: lastfind = Key: Index = K: Done = True: Exit Function
     ParentIndex = K
     k1 = K
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0 Or K = k1
End If
End If
End If

End Function
Function FindOne(RHS, num) As Boolean
Dim K As Long, Key As String, k1 As Long
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
ParentIndex = -1
Done = False
K = Hash(HashFunc(Key)) - 1
k1 = -2
If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then
        num = num + 1
        If num = 0 Then FindOne = True: lastfind = Key: Index = K: Done = True: Exit Function
    End If
     ParentIndex = K
     k1 = K
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0 Or k1 = K
If k1 = K Then
Rehash
K = Hash(HashFunc(Key)) - 1
k1 = -2
ParentIndex = -1
If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then
      num = num + 1
        If num = 0 Then FindOne = True: lastfind = Key: Index = K: Done = True: Exit Function
    End If
     ParentIndex = K
     k1 = K
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0 Or K = k1
End If
End If
End If

End Function
Private Function Find2(RHS, Parndx As Long, nDx As Long) As Boolean
Dim K As Long, Key As String, k1 As Long
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
Parndx = -1
K = Hash(HashFunc(Key)) - 1
k1 = -2
If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then Find2 = True: nDx = K: Exit Function
     Parndx = K
     k1 = K
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0 Or k1 = K
If k1 = K Then
Parndx = -1
K = Hash(HashFunc(Key)) - 1
k1 = -2
If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then Find2 = True: nDx = K: Exit Function
     Parndx = K
     k1 = K
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0 Or k1 = K
End If
End If
End If
End Function

Function ExistKey0(RHS) As Boolean
Dim K As Long, Key As String, k1 As Long
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
K = Hash(HashFunc(Key)) - 1
k1 = -2
If K >= 0 Then
    Do
        If PriveSpace(K).Key = Key Then ExistKey0 = True: Index = K: Exit Function
        k1 = K
        K = PriveSpace(K).Pleft - 1
    Loop Until K < 0 Or k1 = K
    If K = k1 Then
        Rehash
        K = Hash(HashFunc(Key)) - 1
        k1 = -2
        If K >= 0 Then
        Do
            If PriveSpace(K).Key = Key Then ExistKey0 = True: Index = K: Exit Function
            k1 = K
            K = PriveSpace(K).Pleft - 1
        Loop Until K < 0 Or k1 = K
        End If
    End If
End If
End Function
Function ExistKey(RHS) As Boolean
Dim K As Long, k1 As Long, Key As String
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
K = Hash(HashFunc(Key)) - 1
k1 = -2
If K >= 0 Then
    Do
        If PriveSpace(K).Key = Key Then ExistKey = True: Index = K: Exit Function
        k1 = K
        K = PriveSpace(K).Pleft - 1
        
    Loop Until K < 0 Or K = k1
    If K = k1 Then
        Rehash
        K = Hash(HashFunc(Key)) - 1
        If K >= 0 Then
        Do
            If PriveSpace(K).Key = Key Then ExistKey = True: Index = K: Exit Function
            k1 = K
            K = PriveSpace(K).Pleft - 1
        Loop Until K < 0 Or K = k1
        End If
    End If
End If
End Function
Function ExistKey3(RHS) As Boolean
Dim K As Long, k1 As Long, Key As String
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
k1 = -2
K = Hash(HashFunc(Key)) - 1
If K >= 0 Then
    Do
        If PriveSpace(K).Key = Key Then ExistKey3 = True: Exit Function
        k1 = K
        K = PriveSpace(K).Pleft - 1
    Loop Until K < 0 Or K = k1
    If K = k1 Then
        Rehash
        K = Hash(HashFunc(Key)) - 1
        If K >= 0 Then
            Do
                If PriveSpace(K).Key = Key Then ExistKey3 = True: Exit Function
                k1 = K
                K = PriveSpace(K).Pleft - 1
            Loop Until K < 0 Or K = k1
        End If
    End If
End If
End Function

Private Function ExistKey2(Key As String, nkey As Long) As Boolean
Dim K As Long, k1 As Long
If Len(Key) = 0 Then ExistKey2 = True: Exit Function
K = Hash(HashFunc1(nkey)) - 1
k1 = -2
If K >= 0 Then
    Do
        If PriveSpace(K).Key = Key Then ExistKey2 = True: Exit Function
        k1 = K
        K = PriveSpace(K).Pleft - 1
    Loop Until K < 0 Or K = k1
    If K = k1 Then
        Rehash
        K = Hash(HashFunc(Key)) - 1
        If K >= 0 Then
        Do
            If PriveSpace(K).Key = Key Then ExistKey2 = True: Exit Function
            k1 = K
            K = PriveSpace(K).Pleft - 1
            
        Loop Until K < 0 Or K = k1
        End If
    End If

End If
End Function

Private Sub ItemCreator(Key As String, nValue As Variant)
Dim a As Long, kk As Long
Done = False
kk = HD(Key)
If ExistKey2(Key, kk) Then Exit Sub  ' done is false
a = Malloc()
With PriveSpace(a)
    .Key = Key
    .Offset = CLng(nValue)
    .firsthash = kk
End With
place HashFunc1(kk), a
Index = a
Done = True
End Sub
Friend Sub ItemCreator2(Key As String, nValue As Long, dvalue As Long)
Dim a As Long, kk As Long
Done = False
kk = HD(Key)
a = Malloc()
With PriveSpace(a)
    .Key = Key
    .Offset = CLng(nValue)
    .sValue = dvalue
    .firsthash = kk
End With
place HashFunc1(kk), a
Index = -1
End Sub
Property Get Percent()
Percent = 100 * Count / hashlen
End Property
Private Sub place(ByVal b, ByVal a)
    Dim K As Long
    K = Hash(b)
    If Not K = a + 1 Then
         Hash(b) = a + 1
         PriveSpace(a).Pleft = K
    End If
    PriveSpace(a).lastpos = b
End Sub
Private Function HashFunc2(where As Long)
HashFunc2 = PriveSpace(where).firsthash Mod hashlen
End Function
Private Function HashFunc1(nkey As Long)
HashFunc1 = nkey Mod hashlen
End Function
Private Function HashFunc(a$)
HashFunc = HD(a$) Mod hashlen
End Function
Private Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(a), 16
   CopyMemory ByVal VarPtr(a), ByVal VarPtr(b), 16
   CopyMemory ByVal VarPtr(b), t(0), 16
End Sub
' only for tests
Private Function GetRefCount(ByRef vvv) As Long
    If Not IsObject(vvv) Then Exit Function
    Dim obj As IUnknown
    Set obj = vvv
    If obj Is Nothing Then Exit Function
    CopyMemory GetRefCount, ByVal (ObjPtr(obj)) + 4, 4
    GetRefCount = GetRefCount - 2
    Set obj = Nothing
End Function
Public Property Get ReferCount() As Long
ReferCount = GetRefCount(Me)
End Property
Public Property Get KeyToString2() As Variant
If Index > -1 Then
    KeyToString2 = PriveSpace(Index).Key + str$(PriveSpace(Index).lastpos < toplim) + str$(PriveSpace(Index).Pleft)
End If
End Property
Public Property Get KeyToString() As Variant
If Index > -1 Then
    KeyToString = PriveSpace(Index).Key
End If
End Property
Public Property Get ReferCountValue() As Long
End Property
Public Sub ToStart()
Index = 0
Done = Count > 0
End Sub
Public Sub ToEnd()
Index = Count - 1
Done = Count > 0
End Sub
Public Sub NextIndex()
Done = False
If Index = -1 Then
Else
    If Index < toplim Then Done = True: Index = Index + 1
End If
End Sub
Public Function ToNextIndex()
Done = False
If Index = -1 Then
Else
    If Index < toplim Then Done = True: Index = Index + 1
End If
ToNextIndex = Done
End Function
Public Sub PrevIndex()
Done = False
If Index = -1 Then
Else
    If Index > 0 Then Done = True: Index = Index - 1
End If
End Sub
Public Function ToPrevIndex()
    Done = False
    If Index = -1 Then
    Else
        If Index > 0 Then Done = True: Index = Index - 1
    End If
    ToPrevIndex = Done
End Function

Private Sub Class_Terminate()
    Erase PriveSpace()
End Sub

Public Property Get KeyToString3(there As Long) As String
    KeyToString3 = PriveSpace(there).Key
End Property
Public Function ExportKeys() As Object
    Dim a As New tuple, i As Long
    a.PushDim Count
    a.PushEnd
    a.arr = True
    For i = 0 To Count - 1
        a.item(i) = PriveSpace(i).Key
    Next
    Set ExportKeys = a
End Function

Friend Sub AddSeven(a As Long, Key As String, iValue As Variant, sValue As Long, firsthash As Variant, par1 As Variant, par2 As Variant)
    With PriveSpace(a)
    .Key = Key
    If IsObject(iValue) Then Set .struct = iValue
    .sValue = sValue
    .firsthash = firsthash
    .Offset = par1
    .items = par2
    End With
    toplim = toplim + 1
End Sub
Friend Sub CopyCollection(this As StructCollection)
    Set this = New StructCollection
    Dim c As Long, D As Long
    D = 0
    c = Count
    this.MakeSpace MaxSpace
    While c > D
        With PriveSpace(D)
            this.AddSeven D, .Key, .struct, .sValue, .firsthash, .Offset, .items
        End With
        D = D + 1
    Wend
    this.ExpandHash
    ' do not forget to fix Tag outside
    this.structLen = structLen
End Sub
Friend Sub MergeCollection(this As StructCollection, Offset As Long, lasthead$)
'   we place our data to This collection
    Dim D As Long, c As Long, b As Long
    D = this.Count
    c = Count
    this.AddMoreSpace c
    While b < c
        With PriveSpace(b)
            this.AddSeven D + b, lasthead$ + .Key, .struct, .sValue, .firsthash, Offset + .Offset, .items
        End With
        b = b + 1
    Wend
    this.ExpandHash
End Sub

