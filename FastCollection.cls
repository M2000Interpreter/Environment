VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' George's FastCollection
' Only Add, and Find.
' Key as string may be a hex number (used for enumerates)
' zero$ has chr$(1)
Const zero$ = ""
Private Declare Function CompareString Lib "kernel32" Alias "CompareStringW" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long) As Long
Const NowDec$ = ","
' No same key. --- this is changed
Private AllowSameKey As Boolean
' We can remove a key - if AllowSameKey=false

' We can add key (only unique keys)
' return variant (can be anything)
'**************
'* Propertiea:
'*      Count
'*'* from 1 to Count we can access all the members of FastCollection
'*     .Index=1  used for input/output
'* We can get Key or Value
'*      .KeyToString
'*      .Value
'* We can change Valus
'*    set .ValueObj = objectX
'*     .ValueVar
'**************************
'* Methods
'* AddKey "10" [, variant]
'* ToKey "10"
'* SwapValues "10", "30"
'*
'* .Done ..operation is ok
'* .Remove 10
'* .AddKey without parameter add 1 to lastkey
'* .LastKey
'* .LastIndex
'* .Frwd  (for index, exit if not .done)
'* .Back  (for index, exit if not .done)
Private mSortstyle As VbCompareMethod
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type item
    Key As String
    KeyType As Integer
    iValue As Variant
    sValue As Long
    firsthash As Long
    lastpos As Long
    Pleft As Long  ' a list
End Type
Private PriveSpace() As item
Dim MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Public Stable As Long
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, retval As Byte)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private lastfind As String
Private lasttype As Integer
Private lastkey As String
Private Hash() As Long
Public Done As Boolean
Public Index As Long
Public structLen As Long
Private ParentIndex As Long
Public NumericSort As Boolean
Public IamBusy As Boolean
Public useclid As Long
Private UseUcaseKeys As Boolean, Sorted As Boolean
Private uScol() As Byte, useSCol As Boolean
Private entrance As Long
Public Tag As String
Dim zeroitem As item, Piv As item, swapb(100) As Byte



Private Sub Class_Initialize()
    MaxSpace = 5
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
    mSortstyle = vbTextCompare
End Sub
Public Sub ResetGui()
    MaxSpace = 50
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
    mSortstyle = vbTextCompare
End Sub
Public Sub SetBinaryCompare()
mSortstyle = vbBinaryCompare
End Sub
Public Sub SetTextCompare()
mSortstyle = vbTextCompare
End Sub
Public Sub SetDataBaseCompare()
mSortstyle = vbDatabaseCompare
End Sub


Public Sub AddKey2(RHS, Optional aValue As Variant)
IamBusy = True
    Index = -1
    
    lastkey = Normalize(RHS)
    If Len(lastkey) = 0 Then RHS = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0): lastkey = RHS
    Done = False
    lasttype = VarType(RHS)
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
    
IamBusy = False
End Sub
Public Sub AddKey(RHS, Optional aValue As Variant)
    Index = -1
    
     lastkey = Normalize(RHS)
 
    If Len(lastkey) = 0 Then RHS = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0): lastkey = RHS
    Done = False
    lasttype = VarType(RHS)
    If Not IsMissing(aValue) Then
    
   
        If IsObject(aValue) Then
        Value = -1
        Set ValueObj = aValue
        Else
        Value = aValue
        End If
    Else
        Value = aValue
        If Index >= 0 Then
        PriveSpace(Index).KeyType = lasttype
        Else
        PriveSpace(Index).KeyType = 0 ' invalid
        End If
    End If
End Sub
Public Sub Sort()
If toplim = -1 Then Exit Sub
IamBusy = True
If AllowSameKey Then
If Stable Then
InsertionSort PriveSpace(), LBound(PriveSpace()), toplim
Else
If NumericSort Then
QuickSortItemsNumbers PriveSpace(), LBound(PriveSpace()), toplim
Else
QuickSortItemsWithSameKey PriveSpace(), LBound(PriveSpace()), toplim
End If
End If
Else
If NumericSort Then
QuickSortItemsNumbers PriveSpace(), LBound(PriveSpace()), toplim
Else
QuickSortItems PriveSpace(), LBound(PriveSpace()), toplim
End If
End If
Rehash
Sorted = True
IamBusy = False
End Sub
Public Sub SortDes()

If toplim = -1 Then Exit Sub
IamBusy = True
If AllowSameKey Then
If Stable Then
InsertionSort PriveSpace(), LBound(PriveSpace()), toplim, 1
Else
If NumericSort Then
QuickSortItemsRevNumbers PriveSpace(), LBound(PriveSpace()), toplim
Else
QuickSortItemsRevWithSameKey PriveSpace(), LBound(PriveSpace()), toplim

End If
End If
Else
If NumericSort Then
QuickSortItemsRevNumbers PriveSpace(), LBound(PriveSpace()), toplim
Else
QuickSortItemsRev PriveSpace(), LBound(PriveSpace()), toplim
End If
End If
Rehash
Sorted = True
IamBusy = False
End Sub


Public Sub ToKey(RHS)
    Done = Find(RHS)
    If Done Then lastkey = Normalize(RHS)
End Sub
Public Sub RemoveWithNoFind()
'If AllowSameKey Then Exit Sub
Dim new_item As item
Dim K As Long, k1 As Long, vvv As Variant, top1 As Long
If toplim < 0 Then Exit Sub
top1 = toplim
If Done Then
Sorted = False
       If ParentIndex > -1 Then
            PriveSpace(ParentIndex).Pleft = PriveSpace(Index).Pleft

        Else
        
        Hash(PriveSpace(Index).lastpos) = PriveSpace(Index).Pleft

        End If
        If top1 = Index Then
            PriveSpace(top1) = new_item
        Else
            SwapVariant vvv, PriveSpace(top1).iValue
            PriveSpace(Index).iValue = -1
            PriveSpace(Index) = PriveSpace(top1)
            PriveSpace(top1) = new_item
            SwapVariant vvv, PriveSpace(Index).iValue
        End If
        K = Hash(PriveSpace(Index).lastpos) - 1
        If K = top1 Then
            ' we have to give the new position to Hash()
            Hash(PriveSpace(Index).lastpos) = Index + 1
        Else
            Do While K >= 0 And K <> top1 And K <> k1
                   k1 = K
                   K = PriveSpace(K).Pleft - 1
            Loop
            If K = k1 Then
               Rehash
            ElseIf K = top1 Then
               PriveSpace(k1).Pleft = Index + 1
            End If
        End If
  
here:
       toplim = toplim - 1
      
        
       
        
  ReduceHash toplim
End If
End Sub
Public Sub drop(RHS)
'If Not AllowSameKey Then Exit Sub
RHS = RHS - 1
If (toplim - RHS) >= 0 And (RHS >= 0) Then
Dim I As Long
For I = toplim To RHS Step -1
       With PriveSpace(I)
        If .Pleft > 0 Then
            Hash(.lastpos) = .Pleft
            .Pleft = 0
        Else
            Hash(.lastpos) = 0
        End If
        .Key = vbNullString
        .iValue = CDbl(0)
        End With
toplim = RHS - 1
  ReduceHash toplim
Next I
Else
   MaxSpace = 5
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
    mSortstyle = vbTextCompare
End If

End Sub
Public Sub Remove(RHS)
If AllowSameKey Then Exit Sub
Dim K As Long, k1 As Long, vvv As Variant, top1 As Long, nDx As Long, Parndx As Long
If toplim < 0 Then Exit Sub
top1 = toplim
If Find2(RHS, Parndx, nDx) Then
        Sorted = False
        Done = True
        Index = nDx
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
        If top1 = nDx Then
        PriveSpace(top1) = new_item
        Else
        SwapVariant vvv, PriveSpace(top1).iValue
        PriveSpace(nDx).iValue = -1 ' is good to pass a number - so we break any pointer to object
        PriveSpace(nDx) = PriveSpace(top1) ' now we perform a copy (no objects included)
        PriveSpace(top1) = new_item
        ' so now we put back iValue (maybe is an object)
        SwapVariant vvv, PriveSpace(nDx).iValue
        End If
        ' here Lastpos is from previous PriveSpace(top1) which is PriveSpace(ndx)
        K = Hash(PriveSpace(nDx).lastpos) - 1
        If K = top1 Then
            ' we have to give the new position to Hash()
            Hash(PriveSpace(nDx).lastpos) = nDx + 1
        Else
         
         Do While K >= 0 And K <> top1 And K <> k1
                k1 = K
                K = PriveSpace(K).Pleft - 1
         Loop
         If K = k1 Then
            Rehash
         ElseIf K = top1 Then
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
If Index = -1 Then
ItemCreator lastkey, RHS

Else
    PriveSpace(Index).iValue = RHS
End If
Done = True
End Property
Property Set ValueObj(RHS As Variant)
Done = False
If Index = -1 Then
    ItemCreator lastkey, RHS
Else
    Set PriveSpace(Index).iValue = RHS
End If
Done = True
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
If Index = -1 Then
Else
    Done = True
    KeyTypeValue = PriveSpace(Index).KeyType
End If
End Property
Property Let KeyTypeValue(RHS As Integer)
If Index = -1 Then
Else
    Done = True
    PriveSpace(Index).KeyType = RHS
End If
End Property

Property Get Value() As Variant
Done = False
If Index = -1 Then
ElseIf Not IsObject(PriveSpace(Index).iValue) Then
    Done = True
    Value = PriveSpace(Index).iValue
    
    If VarTypeName(Value) = "Error" Then
    
    Value = MyVal()
    If VarType(Value) = vbString Then Value = PriveSpace(Index).Key
    
    End If
End If
End Property
Public Function IsString() As Boolean
Dim n As Byte
Done = False
If Index = -1 Then
Else
    Done = True
     GetMem1 VarPtr(PriveSpace(Index).iValue), n
     If n = 10 Then n = PriveSpace(Index).KeyType
End If
IsString = n = 8
End Function
Private Function MyVal() As Variant
On Error Resume Next
MyVal = PriveSpace(Index).Key
Select Case PriveSpace(Index).KeyType
Case vbString
If MyVal Like "[-0-9.]*" Then
    MyVal = CDbl(Replace$(MyVal, ".", NowDec$))
Else
    MyVal = vbNullString
End If
Case vbDouble
MyVal = CDbl(Replace$(MyVal, ".", NowDec$))
Case vbCurrency
MyVal = CCur(Replace$(MyVal, ".", NowDec$))
Case vbDecimal
MyVal = CDec(Replace$(MyVal, ".", NowDec$))
Case 20
MyVal = cInt64(MyVal)
Case Else
MyVal = CDbl(Replace$(MyVal, ".", NowDec$))
End Select

End Function
Function ValueType(ch As Long, p As Variant, s As String) As Boolean
Done = False
Dim val1 As Variant
If Index = -1 Then
ValueType = False
ElseIf Not IsObject(PriveSpace(Index).iValue) Then
ValueType = True
    Done = True
    val1 = PriveSpace(Index).iValue
    If CheckInt64(val1) Then
           If ch = 2 Then
            s = val1
        Else
            p = 0
            On Error Resume Next
            p = val1
        End If
    ElseIf Typename(val1) = "Error" Then
        If ch = 2 Then
            s = KeyToString
        Else
        p = 0
            On Error Resume Next
            p = KeyToNumber
        End If
    Else
        If ch = 2 Then
            s = val1
        Else
            p = 0
            On Error Resume Next
            p = val1
        End If
    End If
End If
End Function
Property Get ValueObj() As Variant
Done = False
If Index = -1 Then
ElseIf IsObject(PriveSpace(Index).iValue) Then
    Set ValueObj = PriveSpace(Index).iValue
    Done = True
End If

End Property
Private Sub ExpandHash()
    hashlen = MaxSpace * 2 + 3
    Rehash
End Sub
Private Sub Rehash()
    Dim I As Long
    ReDim Hash(hashlen) As Long
    For I = 0 To toplim
           place HashFunc2(I), I
        
    Next I
End Sub
Private Function CheckHash() As Boolean
    Dim I As Long
    For I = 0 To toplim
        If PriveSpace(I).Pleft > 0 Then
        If PriveSpace(PriveSpace(I).Pleft - 1).Key = vbNullString Then Exit Function
        If I = (PriveSpace(I).Pleft - 1) Then Exit Function
        End If
    Next I
    CheckHash = True
End Function


Private Sub ReduceHash(newTop As Long)
entrance = entrance + 1
If entrance > 0 Then Exit Sub
    If newTop <= 5 Then
            If newTop < 0 Then
                newTop = -1
                toplim = -1
                MaxSpace = 5
                ReDim PriveSpace(MaxSpace) As item
                hashlen = MaxSpace * 2 + 3
                ReDim Hash(hashlen)
                
            ElseIf MaxSpace > 40 Then
                MaxSpace = 5
                ReDim Preserve PriveSpace(MaxSpace) As item
                
                hashlen = MaxSpace * 2 + 3
                
                Rehash
            End If
    Else
            If MaxSpace - newTop + 1 > 2 * newTop + 2 Then
            MaxSpace = 2 * (newTop + 1) + 1
            hashlen = MaxSpace * 2 + 3
            ReDim Preserve PriveSpace(MaxSpace) As item
            Rehash
    End If
   End If
   entrance = entrance - 1
End Sub
Public Function Normalize(Key) As String

        If VarType(Key) = vbString Then
        If UseUcaseKeys Then
            Normalize = myUcase(Key)
        Else
            Normalize = Key
        End If
        ElseIf VarType(Key) = vbBoolean Then
        Normalize = str(CLng(Key))
        Else
        Normalize = LTrim$(str$(Key))
        End If
End Function
Function ChangeKey(key1, Key2)
If AllowSameKey Then Exit Function
    Dim Key As String, NewKey As String
    Dim K As Long
    If key1 = Key2 Then ChangeKey = True
    Key = Normalize(key1)
    If Len(Key) = 0 Then Exit Function
    If ExistKey(Key2) Then Exit Function
    
    NewKey = Normalize(Key2)
    If Len(NewKey) = 0 Then Exit Function
    
    K = Hash(HashFunc(Key)) - 1
    If K >= 0 Then
    Do
            If PriveSpace(K).Key = Key Then
                PriveSpace(K).Key = NewKey
                PriveSpace(K).firsthash = HD(NewKey)
                ReDim Hash(MaxSpace * 2 + 3)
                hashlen = MaxSpace * 2 + 3
                Rehash
                ChangeKey = True
                Sorted = False
                Exit Function
            End If
            K = PriveSpace(K).Pleft - 1
    Loop Until K < 0
    End If
End Function
Sub swap(key1, Key2)
Dim K As Long, C As item
If Find(key1) Then
K = Index
If Find(Key2) Then

    C = PriveSpace(K): PriveSpace(K) = PriveSpace(Index): PriveSpace(Index) = C
    Index = K
    Rehash
    Sorted = False
End If
End If
End Sub
Sub SwapValues(key1, Key2)
Dim K As Long
If Find(key1) Then
K = Index
If Find(Key2) Then
    SwapVariant PriveSpace(K).iValue, PriveSpace(Index).iValue
End If
End If
End Sub
Sub SwapNext(key1)
Dim K As Long
On Error Resume Next
If Done Then
    K = Index
    If Find(key1) Then
        SwapVariant PriveSpace(K).iValue, PriveSpace(Index).iValue
    End If
End If
End Sub
Sub SwapByIndex(RHS As Long)
On Error Resume Next
If Done Then
    SwapVariant PriveSpace(RHS).iValue, PriveSpace(Index).iValue
    Index = RHS
End If
End Sub


Private Function Malloc() As Long
    If toplim + 1 >= MaxSpace Then
        MaxSpace = MaxSpace * 2
        ReDim Preserve PriveSpace(MaxSpace) As item
        If MaxSpace > hashlen * 3 / 4 Then ExpandHash
    End If
    toplim = toplim + 1
    Malloc = toplim
End Function
Property Get count()
    count = toplim + 1
End Property
Public Property Let UcaseKeys(RHS)
If toplim = -1 Then
UseUcaseKeys = CBool(RHS)
End If
End Property

Property Get IsEmpty()
    IsEmpty = toplim = -1
End Property
Property Get LastKnown() As String
    LastKnown = lastfind
End Property
Function IsObj() As Boolean
If Index = -1 Then
Else
IsObj = IsObject(PriveSpace(Index).iValue)
End If
End Function
Function IsObjAt(curitem As Long, peekvalue) As Boolean
If Index = -1 Then
Else
If IsObject(PriveSpace(Index).iValue) Then
    Set peekvalue = PriveSpace(Index).iValue
    IsObjAt = True
End If
End If
End Function
Function IsEnum(p As Variant) As Boolean
Dim usehandler As mHandler
If Index = -1 Then
ElseIf IsObject(PriveSpace(Index).iValue) Then
If VarTypeName(PriveSpace(Index).iValue) = "mHandler" Then
    Set usehandler = PriveSpace(Index).iValue
    IsEnum = usehandler.T1 = 4
    If IsEnum Then
        If myVarType(usehandler.index_cursor, vbString) Then
            p = usehandler.index_cursor
        Else
            p = usehandler.index_cursor * usehandler.sign
        End If
    End If
End If
End If
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
If AllowSameKey Then Exit Function
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
Dim A As Long, kk As Long
Done = False
Sorted = False
kk = HD(Key)
If Not AllowSameKey Then If ExistKey2(Key, kk) Then Exit Sub  ' done is false
A = Malloc()
With PriveSpace(A)
    .Key = Key
    If IsObject(nValue) Then
    Set .iValue = nValue
    Else
    .iValue = nValue
    End If
    .firsthash = kk
End With

place HashFunc1(kk), A

Index = A
Done = True
End Sub
Friend Sub ItemCreator2(Key As String, nValue As Long, dvalue As Long)
Dim A As Long, kk As Long
Done = False
Sorted = False
kk = HD(Key)
' without check
'If Not AllowSameKey Then If ExistKey2(Key, kk) Then Exit Sub  ' done is false
A = Malloc()
With PriveSpace(A)
    .Key = Key
    .iValue = nValue
    .sValue = dvalue
    .firsthash = kk
End With
place HashFunc1(kk), A
Index = -1
End Sub
Property Get Percent()
Percent = 100 * count / hashlen
End Property
Private Sub place(ByVal b, ByVal A)
    Dim K As Long
    K = Hash(b)
    If Not K = A + 1 Then
         Hash(b) = A + 1
         PriveSpace(A).Pleft = K
    End If
    PriveSpace(A).lastpos = b
End Sub
Private Function HashFunc2(where As Long)
HashFunc2 = PriveSpace(where).firsthash Mod hashlen
End Function
Private Function HashFunc1(nkey As Long)
HashFunc1 = nkey Mod hashlen
End Function
Private Function HashFunc(A$)
HashFunc = HD(A$) Mod hashlen
End Function


Private Sub SwapVariant(ByRef A As Variant, ByRef b As Variant)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(A), 16
   CopyMemory ByVal VarPtr(A), ByVal VarPtr(b), 16
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

    KeyToString2 = PriveSpace(Index).Key + str$(PriveSpace(Index).lastpos < toplim) + str$(PriveSpace(Index).Pleft) ' + Str$(Int(Percent * 1000) / 100) + Str$(hashlen) + Str$(count)
End If
End Property
Public Property Get KeyToString() As Variant
If Index > -1 Then

    KeyToString = PriveSpace(Index).Key
End If
End Property
Public Property Get HassKeyNumberPart() As Variant
If Index > -1 Then
HassKeyNumberPart = Left$(PriveSpace(Index).Key, 1) Like "[-0-9.]*"
End If
End Property
Public Property Get KeyToNumber() As Variant
If Index > -1 Then
    KeyToNumber = MyVal()
    If VarType(KeyToNumber) = vbString Then KeyToNumber = CDbl(0)
End If
End Property
Public Property Get ReferCountValue() As Long
If Index = -1 Then
ElseIf IsObject(PriveSpace(Index).iValue) Then
    ReferCountValue = GetRefCount(PriveSpace(Index).iValue)
End If
End Property
Public Sub ToStart()
Index = 0
Done = count > 0
End Sub
Public Sub ToEnd()
Index = count - 1
Done = count > 0
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

Private Sub QuickSortItems(arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
  If UB - LB = 1 Then
     M1 = LB

     If CompareStr2(arr(M1).Key, arr(UB).Key, mSortstyle) = 1 Then SwapItem VarPtr(arr(M1)), VarPtr(arr(UB))
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSort1 arr(), LB, UB
  Exit Sub
  Else
    M1 = (LB + UB) \ 2
   LoadItem VarPtr(Piv), VarPtr(arr(M1))
   M1 = LB
   Do While CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) = -1: M1 = M1 + 1: Loop
    
  End If
  
  M2 = UB
  Do
    Do While CompareStr2(arr(M2).Key, Piv.Key, mSortstyle) = 1: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      SwapItem VarPtr(arr(M1)), VarPtr(arr(M2))
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) = -1: M1 = M1 + 1: Loop
  Loop
  LoadItem VarPtr(Piv), VarPtr(zeroitem)
  If LB < M2 Then QuickSortItems arr, LB, M2
  If M1 < UB Then QuickSortItems arr, M1, UB
End Sub
Private Sub QuickSortItemsRevWithSameKey(arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
  If UB - LB = 1 Then
     M1 = LB

     If CompareStr2(arr(M1).Key, arr(UB).Key, mSortstyle) = -1 Then SwapItem VarPtr(arr(M1)), VarPtr(arr(UB))
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSort1 arr(), LB, UB
  Exit Sub
  
  Else
    M1 = (LB + UB) \ 2
    LoadItem VarPtr(Piv), VarPtr(arr(M1))
    If CompareStr2(arr(LB).Key, Piv.Key, mSortstyle) = 0 Then
            M2 = UB - 1
            M1 = LB
            LoadItem VarPtr(Piv), VarPtr(arr(LB))
            Do
                M1 = M1 + 1
                If M1 > M2 Then
                    If CompareStr2(arr(UB).Key, Piv.Key, mSortstyle) = 1 Then SwapItem VarPtr(arr(LB)), VarPtr(arr(UB))
                    Exit Sub
                End If
            Loop Until CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) <> 0
            LoadItem VarPtr(Piv), VarPtr(arr(M1))
            If M1 > LB Then If CompareStr2(arr(LB).Key, Piv.Key, mSortstyle) = -1 Then SwapItem VarPtr(arr(LB)), VarPtr(arr(M1)): LoadItem VarPtr(Piv), VarPtr(arr(M1))
    Else
        M1 = LB
        Do While CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) = 1: M1 = M1 + 1: Loop
        End If
    End If
  
  M2 = UB
  Do
    Do While CompareStr2(arr(M2).Key, Piv.Key, mSortstyle) = -1: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      SwapItem VarPtr(arr(M1)), VarPtr(arr(M2))
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) = 1: M1 = M1 + 1: Loop
  Loop
  LoadItem VarPtr(Piv), VarPtr(zeroitem)
  If LB < M2 Then QuickSortItemsRevWithSameKey arr, LB, M2
  If M1 < UB Then QuickSortItemsRevWithSameKey arr, M1, UB
End Sub
Private Sub QuickSortItemsWithSameKey(arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
  If UB - LB = 1 Then
     M1 = LB

     If CompareStr2(arr(M1).Key, arr(UB).Key, mSortstyle) = 1 Then SwapItem VarPtr(arr(M1)), VarPtr(arr(UB))
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSort1 arr(), LB, UB
  Exit Sub
  
  Else
        M1 = (LB + UB) \ 2
        LoadItem VarPtr(Piv), VarPtr(arr(M1))
        If CompareStr2(arr(LB).Key, Piv.Key, mSortstyle) = 0 Then
                M2 = UB - 1
                M1 = LB
                LoadItem VarPtr(Piv), VarPtr(arr(LB))
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If CompareStr2(arr(UB).Key, Piv.Key, mSortstyle) = -1 Then SwapItem VarPtr(arr(LB)), VarPtr(arr(UB))
                        Exit Sub
                    End If
                Loop Until CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) <> 0
                LoadItem VarPtr(Piv), VarPtr(arr(M1))
                If M1 > LB Then If CompareStr2(arr(LB).Key, Piv.Key, mSortstyle) = 1 Then SwapItem VarPtr(arr(LB)), VarPtr(arr(M1)): LoadItem VarPtr(Piv), VarPtr(arr(M1))
        Else
                M1 = LB
                Do While CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) = -1: M1 = M1 + 1: Loop
        End If
  End If
  
  M2 = UB
  Do
    Do While CompareStr2(arr(M2).Key, Piv.Key, mSortstyle) = 1: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      SwapItem VarPtr(arr(M1)), VarPtr(arr(M2))
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) = -1: M1 = M1 + 1: Loop
  Loop
  LoadItem VarPtr(Piv), VarPtr(zeroitem)
  If LB < M2 Then QuickSortItemsWithSameKey arr, LB, M2
  If M1 < UB Then QuickSortItemsWithSameKey arr, M1, UB
End Sub
Private Sub QuickSortItemsNumbers(arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
On Error GoTo abc1
     If UB - LB = 1 Then
     M1 = LB
       If compareStr4(arr(M1).Key, arr(UB).Key) = 1 Then SwapItem VarPtr(arr(M1)), VarPtr(arr(UB))
      Exit Sub
    ElseIf UB - LB < 7 Then
        InsertionSort arr(), LB, UB
        Exit Sub
     Else
       M1 = (LB + UB) \ 2 '+ 1
       LoadItem VarPtr(Piv), VarPtr(arr(M1))
             If compareStr4(arr(LB).Key, Piv.Key) = 0 Then
                M2 = UB - 1
                M1 = LB
                LoadItem VarPtr(Piv), VarPtr(arr(LB))
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If compareStr4(arr(UB).Key, Piv.Key) = -1 Then SwapItem VarPtr(arr(LB)), VarPtr(arr(UB))
                        Exit Sub
                    End If
                Loop Until compareStr4(arr(M1).Key, Piv.Key) <> 0
                LoadItem VarPtr(Piv), VarPtr(arr(M1))
                If M1 > LB Then If compareStr4(arr(LB).Key, Piv.Key) = 1 Then SwapItem VarPtr(arr(LB)), VarPtr(arr(M1)): LoadItem VarPtr(Piv), VarPtr(arr(M1))
            Else
                M1 = LB
                Do While compareStr4(arr(M1).Key, Piv.Key) = -1: M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While compareStr4(arr(M2).Key, Piv.Key) = 1: M2 = M2 - 1: Loop
      If M1 <= M2 Then
        SwapItem VarPtr(arr(M1)), VarPtr(arr(M2))
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While compareStr4(arr(M1).Key, Piv.Key) = -1: M1 = M1 + 1: Loop
    Loop
    LoadItem VarPtr(Piv), VarPtr(zeroitem)
    If LB < M2 Then QuickSortItemsNumbers arr, LB, M2
    If M1 < UB Then QuickSortItemsNumbers arr, M1, UB
    Exit Sub
abc1:
    

End Sub

Private Sub QuickSortItemsRevNumbers(arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
On Error GoTo abc1
     If UB - LB = 1 Then
     M1 = LB
       If compareStr4(arr(M1).Key, arr(UB).Key) = -1 Then SwapItem VarPtr(arr(M1)), VarPtr(arr(UB))
      Exit Sub
    ElseIf UB - LB < 7 Then
        InsertionSort arr(), LB, UB, 1
        Exit Sub
     Else
       M1 = (LB + UB) \ 2 '+ 1
       LoadItem VarPtr(Piv), VarPtr(arr(M1))
             If compareStr4(arr(LB).Key, Piv.Key) = 0 Then
                M2 = UB - 1
                M1 = LB
                LoadItem VarPtr(Piv), VarPtr(arr(LB))
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If compareStr4(arr(UB).Key, Piv.Key) = 1 Then SwapItem VarPtr(arr(LB)), VarPtr(arr(UB))
                        Exit Sub
                    End If
                Loop Until compareStr4(arr(M1).Key, Piv.Key) <> 0
                LoadItem VarPtr(Piv), VarPtr(arr(M1))
                If M1 > LB Then If compareStr4(arr(LB).Key, Piv.Key) = -1 Then SwapItem VarPtr(arr(LB)), VarPtr(arr(M1)): LoadItem VarPtr(Piv), VarPtr(arr(M1))
            Else
                M1 = LB
                Do While compareStr4(arr(M1).Key, Piv.Key) = 1: M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While compareStr4(arr(M2).Key, Piv.Key) = -1: M2 = M2 - 1: Loop
      If M1 <= M2 Then
        SwapItem VarPtr(arr(M1)), VarPtr(arr(M2))
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While compareStr4(arr(M1).Key, Piv.Key) = 1: M1 = M1 + 1: Loop
    Loop
    LoadItem VarPtr(Piv), VarPtr(zeroitem)
    If LB < M2 Then QuickSortItemsRevNumbers arr, LB, M2
    If M1 < UB Then QuickSortItemsRevNumbers arr, M1, UB
    Exit Sub
abc1:
    

End Sub
Private Sub QuickSortItemsRev(arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
  If UB - LB = 1 Then
     M1 = LB

     If CompareStr2(arr(M1).Key, arr(UB).Key, mSortstyle) = -1 Then SwapItem VarPtr(arr(M1)), VarPtr(arr(UB))
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSort1 arr(), LB, UB, 1
  Exit Sub
  Else
  ' we have unique names
    M1 = (LB + UB) \ 2
    LoadItem VarPtr(Piv), VarPtr(arr(M1))
    M1 = LB
   Do While CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) = 1: M1 = M1 + 1: Loop
  End If
  
  M2 = UB
  Do
    Do While CompareStr2(arr(M2).Key, Piv.Key, mSortstyle) = -1: M2 = M2 - 1: Loop

    If M1 <= M2 Then
      SwapItem VarPtr(arr(M1)), VarPtr(arr(M2))
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While CompareStr2(arr(M1).Key, Piv.Key, mSortstyle) = 1: M1 = M1 + 1: Loop

  Loop
  LoadItem VarPtr(Piv), VarPtr(zeroitem)
  If LB < M2 Then QuickSortItemsRev arr, LB, M2
  If M1 < UB Then QuickSortItemsRev arr, M1, UB

End Sub
Public Function AllowAnyKey()
AllowSameKey = True
Stable = True
End Function

Function compareStr4(A$, ByVal b$) As Long
If useSCol Then
compareStr4 = compareStr444(A$, b$)
Else
compareStr4 = compareStr44((A$), (b$))
End If
End Function
Private Function compareStr444(A$, b$) As Long
Dim aa() As String, bb() As String, part As Long, ret As Long
aa = Split(A$, zero$)
bb = Split(b$, zero$)
again:
If uScol(part) < 2 Then
ret = CompareStr2(aa(part), bb(part), mSortstyle)
Else
ret = compareStr44(aa(part), bb(part))
End If
If ret = 0 Then
part = part + 1: If UBound(aa) < part Then Exit Function
GoTo again
End If

If (uScol(part) And 1) = 1 Then compareStr444 = -ret Else compareStr444 = ret

End Function
Private Function compareStr44(A$, b$) As Long
Dim I As Long, j As Long, a1$, b1$, P1 As Variant, P2 As Variant, n$, K As Long, k1 As Long
If CompareString(useclid, &H1000, StrPtr(A$), Len(A$), StrPtr(b$), Len(b$)) = 2 Then Exit Function
n$ = "[-0-9.]"
K = Sgn(Len(A$) - Len(b$))
k1 = K
again:
j = IIf(Len(A$) >= Len(b$), Len(b$), Len(A$))

For I = 1 To j

    If Mid$(A$, I, 1) Like n$ Then
    a1$ = Mid$(A$, I)
    b1$ = Mid$(b$, I)
        IsNumberD2 a1$, P1
        If IsNumberD2(b1$, P2) Then
            compareStr44 = Sgn(MyRound(P1, 8) - MyRound(P2, 8))
            If compareStr44 = 0 Then
                If Len(a1$) * Len(b1$) <> 0 Then
                    If Len(n$) > 6 Then n$ = "[-0-9]"
                    K = Sgn(Len(a1$) - Len(b1$))
                     
                    A$ = a1$
                    b$ = b1$
                   
                    GoTo again
                End If
                If Len(a1$) + Len(b1$) = 0 Then
                    compareStr44 = K
                Else
                    compareStr44 = Sgn(Len(a1$) - Len(b1$))
                End If
                Exit Function
            End If
            Exit Function
        Else
        A$ = Mid$(A$, I)
    b$ = Mid$(b$, I)
    IsNumberD2 A$, P1
       
          If Len(A$) > 0 Then
       
            Select Case CompareString(useclid, &H1000, StrPtr(A$), Len(A$), StrPtr(b$), Len(b$))
                Case 1, 3
                compareStr44 = -1
                Case 2
                If K > 0 Then
                compareStr44 = -1
                Else
                compareStr44 = 1
                End If
            End Select
        Else
        compareStr44 = -1
        End If
        End If
        Exit Function
    ElseIf Mid$(b$, I, 1) Like n$ Then
    A$ = Mid$(A$, I)
    b$ = Mid$(b$, I)
    IsNumberD2 b$, P1
    If Len(b$) > 0 Then
       
            Select Case CompareString(useclid, &H1000, StrPtr(A$), Len(A$), StrPtr(b$), Len(b$))
                Case 3
                compareStr44 = 1
                Case 1, 2
                If K < 0 Then
                compareStr44 = 1
                Else
                compareStr44 = -1
                End If
                
                
            End Select
        Else
        compareStr44 = 1
        End If
        Exit Function
    End If
    compareStr44 = -2 + CompareString(useclid, 0, StrPtr(A$) + I * 2 - 2, 1, StrPtr(b$) + I * 2 - 2, 1)
    If compareStr44 <> 0 Then
    Exit Function
    End If
Next I
compareStr44 = Sgn(Len(A$) - Len(b$))
If compareStr44 = 0 Then compareStr44 = k1
End Function

Function CompareStr2(A$, b$, what As Long) As Long
On Error GoTo comperr
Select Case what
Case 1 To 2
CompareStr2 = StrComp(A$, b$, what)
Case Else
If useclid = 0 Then
CompareStr2 = StrComp(A$, b$)
Else
' we cant use something that give "equal" keys
CompareStr2 = -2 + CompareString(useclid, &H1000, StrPtr(A$), Len(A$), StrPtr(b$), Len(b$))
End If
End Select
Exit Function
comperr:
CompareStr2 = -2 + CompareString(0, 0, StrPtr(A$), Len(A$), StrPtr(b$), Len(b$))
End Function

Public Property Get IsQueue() As Variant
IsQueue = AllowSameKey
End Property
Private Sub InsertionSort(arr() As item, ByVal iMin As Long, ByVal iMax As Long, Optional what As Long = -1)
    Dim I As Long
    Dim j As Long
    I = iMin + 1
    iMax = iMax + 1
    While I < iMax
        j = I
        Do While j > iMin
        If compareStr4(arr(j).Key, arr(j - 1).Key) = what Then
            SwapItem VarPtr(arr(j)), VarPtr(arr(j - 1))
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        I = I + 1
    Wend
End Sub
Private Sub InsertionSort1(arr() As item, ByVal iMin As Long, ByVal iMax As Long, Optional what As Long = -1)
    Dim I As Long
    Dim j As Long
    I = iMin + 1
    iMax = iMax + 1
    While I < iMax
        j = I
        Do While j > iMin
        If CompareStr2(arr(j).Key, arr(j - 1).Key, mSortstyle) = what Then
            SwapItem VarPtr(arr(j)), VarPtr(arr(j - 1))
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        I = I + 1
    Wend
End Sub
Public Property Get KeyToString3(there As Long) As String

    KeyToString3 = PriveSpace(there).Key

End Property

Private Function IsNumberD2(A$, D As Variant) As Boolean
' for inline stacitems
If VarType(D) = vbEmpty Then D = 0#
Dim a1 As Long
If A$ <> "" Then
For a1 = 1 To Len(A$) + 1
Select Case Mid$(A$, a1, 1)
Case " ", ChrW(160), vbTab
If a1 > 1 Then Exit For
Case Is = Chr(2)
If a1 = 1 Then Exit Function
Exit For
End Select
Next a1
If a1 > Len(A$) Then a1 = Len(A$) + 1
    If IsNumberOnly(A$, 1, D, a1) Then
        A$ = Mid$(A$, a1)
        IsNumberD2 = True
    Else
    IsNumberD2 = False
    End If
Else
    IsNumberD2 = False
End If

End Function
Private Function IsNumberOnly(A$, fr As Long, r As Variant, lr As Long) As Boolean
Dim sg As Long, sng As Long, n$, ig$, DE$, sg1 As Long, ex$   ', e$
' ti kanei to e$
If A$ = vbNullString Then IsNumberOnly = False: Exit Function
sg = 1
sng = fr - 1
    Do While sng < Len(A$)
    sng = sng + 1
    Select Case Mid$(A$, sng, 1)
    Case " ", "+"
    Case "-"
    sg = -sg
    Case Else
    Exit Do
    End Select
    Loop
n$ = Mid$(A$, sng)

If val("0" + Mid$(A$, sng, 1)) = 0 And Left(Mid$(A$, sng, 1), sng) <> "0" And Left(Mid$(A$, sng, 1), sng) <> "." Then
IsNumberOnly = False

Else
'compute ig$
    If Mid$(A$, sng, 1) = "." Then
    ' no long part
    ig$ = "0"
    DE$ = "."

    Else
    Do While sng <= Len(A$)
        
        Select Case Mid$(A$, sng, 1)
        Case "0" To "9"
        ig$ = ig$ + Mid$(A$, sng, 1)
        Case "."
        DE$ = "."
        Exit Do
        Case Else
        Exit Do
        End Select
       sng = sng + 1
    Loop
    End If
    ' compute decimal part
    If Len(DE$) = 0 Then
    If Len(A$) > sng& Then
    If InStr("EeÅå", Mid$(A$, sng&, 1)) > 0 Then
    If InStr("1234567890+-", Mid$(A$, sng& + 1, 1)) > 0 Then
       ' fr = fr + 1
        DE$ = "."
        
        GoTo cont123
    End If
    End If
    End If
    Else
      sng = sng + 1
cont123:
        Do While sng <= Len(A$)
       
        Select Case Mid$(A$, sng, 1)
        Case " "
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ + Mid$(A$, sng, 1)
        Else
        DE$ = DE$ + Mid$(A$, sng, 1)
        End If
        Case "E", "e", "Å", "å"
               If ex$ = vbNullString Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If

        
        
        Case "+", "-"
        If sg1 And Len(ex$) = 1 Then
         ex$ = ex$ + Mid$(A$, sng, 1)
        Else
        Exit Do
        End If
        Case Else
        Exit Do
        End Select
         sng = sng + 1
        Loop
        If ex$ = "E" Or ex$ = "E-" Or ex$ = "E+" Then
        sng = sng - Len(ex$)
        End If
    End If
    If ig$ = vbNullString Then
        IsNumberOnly = False
        lr = 1
    Else
        If sg < 0 Then ig$ = "-" + ig$
        If Len(ig$ + DE$) > 13 And LenB(ex$) = 0 Then
            On Error Resume Next
            If Len(DE$) > 0 Then
                Mid$(DE$, 1, 1) = DefaultDec$
                r = CDec(ig$ + DE$)
            Else
                r = CDec(ig$)
            End If
            If Err.Number = 6 Then
                r = CDbl(ig$ + DE$)
            End If
         Else
            r = val(ig$ + DE$ + ex$)
             If Err.Number > 0 Then
             Err.Clear
             IsNumberOnly = False
             End If

            End If
      'A$ = Mid$(A$, sng)
    lr = sng - fr + 1
    IsNumberOnly = True
    End If
End If
End Function
Public Sub FeedSCol(ParamArray A())
If UBound(A) >= 0 Then
If Typename(A(0)) = "Variant()" Then
A() = A(0)

End If
End If
ReDim uScol(UBound(A) + 1)
Dim I As Long
For I = 0 To UBound(A)
Select Case A(I)
Case 1, 2, 3

uScol(I) = A(I)
Case Else
uScol(I) = 0
End Select
Next I
useSCol = True
NumericSort = True
End Sub
Public Sub FeedSCol2(A())
ReDim uScol(UBound(A) + 1)
Dim I As Long
For I = 0 To UBound(A)
Select Case A(I)
Case 1, 2, 3
uScol(I) = A(I)
Case Else
uScol(I) = 0
End Select
Next I
useSCol = True
NumericSort = True
End Sub
Public Sub DisposeCol()
useSCol = False
NumericSort = False
End Sub

Friend Sub FeedSCol3(A() As Byte)
uScol() = A()
useSCol = True
NumericSort = True
End Sub
Friend Function ExportScol3() As Byte()
ExportScol3 = uScol()
End Function

Public Function MakeCompKey(ParamArray A()) As String
Dim ret$, I As Long
ret$ = Normalize(A(0))
For I = 1 To UBound(A)
ret$ = ret$ + Chr$(1) + Normalize(A(I))
Next I
MakeCompKey = ret$
End Function
Private Sub LoadItem(A As Long, b As Long)
If A = b Then Exit Sub
CopyMemory ByVal A, ByVal b, LenB(Piv)
End Sub
Private Sub SwapItem(ByVal A As Long, ByVal b As Long)
If A = b Then Exit Sub
CopyMemory ByVal VarPtr(swapb(0)), ByVal A, LenB(Piv)
CopyMemory ByVal A, ByVal b, LenB(Piv)
CopyMemory ByVal b, ByVal VarPtr(swapb(0)), LenB(Piv)
End Sub
Public Function SearchKeyPart(A$) As Boolean
' need to have a sorted list
If Sorted Then
IamBusy = True
If useSCol Then
Else


End If
IamBusy = False
Else
' inventory not sorted
End If
End Function
