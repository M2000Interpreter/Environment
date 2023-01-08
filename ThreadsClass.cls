VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mSortstyle As VbCompareMethod
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type item
    Key As Long
    iValue As Variant
    sValue As Long
    lastpos As Long
    Pleft As Long  ' a list
End Type
Private PriveSpace() As item
Dim MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private lastfind As Long
Private lastkey As Long
Private Hash() As Long
Public Done As Boolean
Public index As Long
Private ParentIndex As Long
Private Sub Class_Initialize()
    MaxSpace = 5
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
End Sub
Public Sub AddKey(RHS, Optional aValue As Variant)
If RHS = 0 Then Exit Sub
    index = -1
    lastkey = CLng(RHS)
    Done = False
    If Not IsMissing(aValue) Then
    
   
        If IsObject(aValue) Then
        Value = -1
        Set ValueObj = aValue
        Else
        Value = aValue
        End If
    Else
        Value = aValue
    End If
End Sub



Public Sub ToKey(RHS)
    Done = Find(RHS)
    If Done Then lastkey = CLng(RHS)
End Sub
Public Sub RemoveWithNoFind()
Dim new_item As item
Dim k As Long, k1 As Long, vvv As Variant, top1 As Long
If toplim < 0 Then Exit Sub
top1 = toplim
If Done Then
       If ParentIndex > -1 Then
            PriveSpace(ParentIndex).Pleft = PriveSpace(index).Pleft

        Else
        
        Hash(PriveSpace(index).lastpos) = PriveSpace(index).Pleft

        End If
        If top1 = index Then
            PriveSpace(top1) = new_item
        Else
            SwapVariant vvv, PriveSpace(top1).iValue
            PriveSpace(index).iValue = -1
            PriveSpace(index) = PriveSpace(top1)
            PriveSpace(top1) = new_item
            SwapVariant vvv, PriveSpace(index).iValue
        End If
        k = Hash(PriveSpace(index).lastpos) - 1
        If k = top1 Then
            ' we have to give the new position to Hash()
            Hash(PriveSpace(index).lastpos) = index + 1
        Else
            Do While k >= 0 And k <> top1 And k <> k1
                   k1 = k
                   k = PriveSpace(k).Pleft - 1
            Loop
            If k = k1 Then
               Rehash
            ElseIf k = top1 Then
               PriveSpace(k1).Pleft = index + 1
            End If
        End If
  
here:
       toplim = toplim - 1
      
        
       
        
  ReduceHash toplim
End If
End Sub
Public Sub Remove(RHS)
Dim k As Long, k1 As Long, vvv As Variant, top1 As Long, nDx As Long, Parndx As Long

If toplim < 0 Then Exit Sub
top1 = toplim
If Find2(RHS, Parndx, nDx) Then
        Done = True
        index = nDx
        ' here we make a clean to pointers to PriveSpace, so we exclude the removing item
        If Parndx > -1 Then
            ' if we have a parent then just push to parent the Pleft pointer
            ' Pleft is always +1 (so Pleft 0 give -1 null for this context)
            PriveSpace(Parndx).Pleft = PriveSpace(nDx).Pleft

        Else
        ' so we have no parent, just out Pleft to Hash
        Hash(PriveSpace(nDx).lastpos) = PriveSpace(nDx).Pleft

        End If
        ' we don't have to see Hash() any more for RHS
        ' but we have to see Hash() for top1
        Dim new_item As item
        ' We have to get change top1 to ndx, maybe we remove top1
        ' we can check this but it is faster to not check this (more often top1 isn't ndx)
        ' Using SwapVariant we can swap any kind of Variant, including objects (swaping pointers only)
        SwapVariant vvv, PriveSpace(top1).iValue
        PriveSpace(nDx).iValue = -1 ' is good to pass a number - so we break any pointer to object
        PriveSpace(nDx) = PriveSpace(top1) ' now we perform a copy (no objects included)
        PriveSpace(top1) = new_item
        ' so now we put back iValue (maybe is an object)
        SwapVariant vvv, PriveSpace(nDx).iValue
        ' here Lastpos is from previous PriveSpace(top1) which is PriveSpace(ndx)
        k = Hash(PriveSpace(nDx).lastpos) - 1
        If k = top1 Then
            ' we have to give the new position to Hash()
            Hash(PriveSpace(nDx).lastpos) = nDx + 1
        Else
            Do While k >= 0 And k <> top1 And k <> k1
                   k1 = k
                   k = PriveSpace(k).Pleft - 1
            Loop
            If k = k1 Then
               Rehash
            ElseIf k = top1 Then
               PriveSpace(k1).Pleft = nDx + 1
            End If
        End If
  
here:
       toplim = toplim - 1
       ReduceHash toplim
Else
Done = False
End If
End Sub


Property Let Value(RHS As Variant)
Done = False
If index = -1 Then
ItemCreator lastkey, RHS

Else
    PriveSpace(index).iValue = RHS
End If
Done = True
End Property
Property Set ValueObj(RHS As Variant)
Done = False
If index = -1 Then
    ItemCreator lastkey, RHS
Else
    Set PriveSpace(index).iValue = RHS
End If
Done = True
End Property
Property Get sValue() As Long
Done = False
If index = -1 Then
Else
    Done = True
    sValue = PriveSpace(index).sValue
End If
End Property
Property Let sValue(RHS As Long)
If index = -1 Then
Else
    Done = True
    PriveSpace(index).sValue = RHS
End If
End Property
Property Get Value() As Variant
Done = False
If index = -1 Then
ElseIf Not IsObject(PriveSpace(index).iValue) Then
    Done = True
    Value = PriveSpace(index).iValue
    
    If VarTypeName(Value) = "Error" Then
    Value = KeyToString
    End If
End If
End Property

Property Get ValueObj() As Variant
Done = False
If index = -1 Then
ElseIf IsObject(PriveSpace(index).iValue) Then
    Set ValueObj = PriveSpace(index).iValue
    Done = True
End If

End Property
Private Sub ExpandHash()
    hashlen = hashlen * 2 + 3
    Rehash
End Sub
Private Sub Rehash()
    Dim i As Long
    ReDim Hash(hashlen) As Long
    For i = 0 To toplim
        place HashFunc2(i), i
    Next i
End Sub

Private Sub ReduceHash(newTop As Long)

   If newTop <= 5 Then
            If newTop < 0 Then
                newTop = -1
                toplim = -1
                MaxSpace = 5
                ReDim PriveSpace(MaxSpace) As item
                
                ReDim Hash(MaxSpace * 2 + 3)
                hashlen = MaxSpace * 2 + 3
            ElseIf MaxSpace > 40 Then
                MaxSpace = 5
                ReDim Preserve PriveSpace(MaxSpace) As item
                
                hashlen = MaxSpace * 2 + 3
                
                Rehash
            End If
    Else
            On Error Resume Next
            ReDim Preserve PriveSpace(MaxSpace) As item
            If Err.Number > 0 Then Exit Sub
            If MaxSpace - newTop + 1 > 2 * newTop + 2 Then
            MaxSpace = 2 * (newTop + 1) + 1
            ReDim Preserve PriveSpace(MaxSpace) As item
            hashlen = MaxSpace * 2 + 3
            Rehash
    End If
   End If
End Sub



Private Function Malloc() As Long
    If toplim + 1 >= MaxSpace Then
        MaxSpace = MaxSpace * 2
        ReDim Preserve PriveSpace(MaxSpace) As item
        If MaxSpace > hashlen Then ExpandHash
    End If
    toplim = toplim + 1
    Malloc = toplim
End Function
Property Get count()
    count = toplim + 1
End Property

Function IsObj() As Boolean
If index = -1 Then
Else
IsObj = IsObject(PriveSpace(index).iValue)
End If
End Function
Function Find(RHS) As Boolean
If RHS = 0 Then Exit Function
If toplim = -1 Then Exit Function
Dim k As Long, Key As Long, k1 As Long
Key = CLng(RHS)
ParentIndex = -1
Done = False

k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
k1 = -2
Do
    If PriveSpace(k).Key = Key Then Find = True: lastfind = Key: index = k: Done = True: Exit Function
    ParentIndex = k
    k1 = k
    k = PriveSpace(k).Pleft - 1
Loop Until k < 0 Or k = k1
If k = k1 Then
Rehash
k = Hash(HashFunc(Key)) - 1
If k >= 0 Then
ParentIndex = -1
k1 = -2
Do
    If PriveSpace(k).Key = Key Then Find = True: lastfind = Key: index = k: Done = True: Exit Function
    ParentIndex = k
    k1 = k
    k = PriveSpace(k).Pleft - 1
Loop Until k < 0 Or k = k1
End If
End If
End If
End Function
Private Function Find2(RHS, Parndx As Long, nDx As Long) As Boolean
If toplim = -1 Then Exit Function
Dim k As Long, Key As Long, k1 As Long

Key = CLng(RHS)
Parndx = -1
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then Find2 = True: nDx = k: Exit Function
     Parndx = k
     k1 = k
     k = PriveSpace(k).Pleft - 1
     
Loop Until k < 0 Or k1 = k
If k1 = k Then
Rehash
Parndx = -1
k1 = -2
k = Hash(HashFunc(Key)) - 1
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then Find2 = True: nDx = k: Exit Function
     Parndx = k
     k1 = k
     k = PriveSpace(k).Pleft - 1
     
Loop Until k < 0 Or k1 = k

End If
End If
End If
End Function

Function ExistKey(RHS) As Boolean
If toplim = -1 Then Exit Function
Dim k As Long, Key As Long, k1 As Long
Key = CLng(RHS)
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then ExistKey = True: index = k: Exit Function
    If k = PriveSpace(k).Pleft - 1 Then Exit Do
    k1 = k
    k = PriveSpace(k).Pleft - 1
    
Loop Until k < 0 Or k = k1
If k = k1 Then
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then ExistKey = True: index = k: Exit Function
    If k = PriveSpace(k).Pleft - 1 Then Exit Do
    k1 = k
    k = PriveSpace(k).Pleft - 1
    
Loop Until k < 0 Or k = k1
End If
End If
End If
End Function

Private Sub ItemCreator(kk As Long, nValue As Variant)
Dim a As Long
Done = False
If ExistKey(kk) Then Exit Sub  ' done is false
a = Malloc()

With PriveSpace(a)
    .Key = kk
    If IsObject(nValue) Then
    Set .iValue = nValue
    Else
    .iValue = nValue
    End If
End With
place HashFunc1(kk), a
index = a
Done = True
End Sub

Property Get Percent()
Percent = 100 * count / hashlen
End Property
Private Sub place(ByVal b, ByVal a)
    Dim k As Long
    k = Hash(b)
    If Not k = a + 1 Then
         Hash(b) = a + 1
         PriveSpace(a).Pleft = k
    End If
    PriveSpace(a).lastpos = b
End Sub

Private Function HashFunc2(where As Long)
HashFunc2 = MyMod(uintnew(PriveSpace(where).Key) + hashlen + 1, hashlen)
End Function
Private Function HashFunc1(nkey As Long)
HashFunc1 = MyMod(uintnew(nkey) + hashlen + 1, hashlen)
End Function
Private Function HashFunc(a As Long)
HashFunc = MyMod(uintnew(a) + hashlen + 1, hashlen)
End Function
Private Function UINT(ByVal a As Integer) As Long
 Dim b As Integer
 b = a And &HFFFF
 If b < 0 Then
 UINT = CLng(&H10000 + b)
 Else
 UINT = CLng(b)
 End If
 
 End Function
Private Function MyModOld(r1 As Double, po) As Double
MyModOld = r1 - Fix(r1 / po) * po
End Function
Private Function MyMod(r1, po) As Variant
MyMod = r1 - Fix(r1 / po) * po
End Function
Private Function uintnew(ByVal a As Currency) As Currency
If a > 2147483647@ Then a = 2147483647@
If a < -2147483648@ Then a = -2147483648@
If a < 0 Then
uintnew = 4294967296@ + a
Else
uintnew = a
End If
End Function
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
If index > -1 Then

    KeyToString2 = PriveSpace(index).Key + Str$(PriveSpace(index).lastpos < toplim) + Str$(PriveSpace(index).Pleft) ' + Str$(Int(Percent * 1000) / 100) + Str$(hashlen) + Str$(count)
End If
End Property
Public Property Get KeyLong() As Long
If index > -1 Then
 KeyLong = PriveSpace(index).Key
End If
End Property
Public Property Get KeyToString() As Variant
If index > -1 Then
    KeyToString = PriveSpace(index).Key
End If
End Property
Public Property Get ReferCountValue() As Long
If index = -1 Then
ElseIf IsObject(PriveSpace(index).iValue) Then
    ReferCountValue = GetRefCount(PriveSpace(index).iValue)
End If
End Property
Public Sub ToStart()
index = 0
Done = count > 0
End Sub
Public Sub ToEnd()
index = count - 1
Done = count > 0
End Sub

Public Sub NextIndex()
Done = False
    If index = -1 Then
 
    Else
        If index < toplim Then Done = True: index = index + 1
    End If

End Sub
Public Sub PrevIndex()
Done = False
    If index = -1 Then
    
    Else
        If index > 0 Then Done = True: index = index - 1
    End If

End Sub
Private Sub Class_Terminate()
    Erase PriveSpace()
End Sub

