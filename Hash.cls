VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 Private Type item
    Key As String
    nDx As Long
    KeyType As Integer  ' 0 as is , 1, reference/don't delete
    originaltype As Byte
    globalvar As Byte
    firsthash As Long
    lastpos As Long
    Pleft As Long  ' a list
 End Type
'
Private PriveSpace() As item
Public MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Public Index As Long
Private Hash() As Long
Private entrance As Long
Public Sub BigSize(n As Long)
    MaxSpace = n
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
End Sub
Private Sub Class_Initialize()
    MaxSpace = 10
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1

End Sub
Friend Sub ReadVar(where, Name$, num As Long)
    Name$ = PriveSpace(where).Key
    num = PriveSpace(where).nDx
End Sub
Public Property Get lastNDX() As Long
If Index < 0 Then
lastNDX = 0
Else
lastNDX = PriveSpace(Index).nDx
End If
End Property
Public Property Get lastkey() As String
If Index < 0 Then
lastkey = vbCrLf
Else
lastkey = PriveSpace(Index).Key
End If
End Property

Private Sub ExpandHash()

    Dim I As Long
    hashlen = hashlen * 2 + 3
    
    ReDim Hash(hashlen) As Long
    For I = 0 To toplim
        place HashFunc2(I), I
    Next I
    
End Sub


Friend Sub ReduceHash(newTop As Long, v())
Dim aa As Variant, ditem As item, ditemcl As item
    ditemcl.nDx = -1
    Dim I As Long
     If toplim < newTop Then Exit Sub
     entrance = entrance + 1
     For I = toplim To newTop Step -1
        ditem = PriveSpace(I)
        PriveSpace(I) = ditemcl
       With ditem
        If .Pleft > 0 Then
            Hash(.lastpos) = .Pleft
            .Key = vbNullString
            .Pleft = 0
        Else
            Hash(.lastpos) = 0
            .Key = vbNullString
        End If
        If .KeyType = 0 And .nDx >= 0 Then
            If IsObject(v(.nDx)) Then
             Set aa = v(.nDx)
             Set v(.nDx) = Nothing
             Set aa = v(.nDx)
            End If
            
            v(.nDx) = Empty
            End If
        .KeyType = 0
        .globalvar = CByte(0)
        .originaltype = CByte(0)
        .nDx = -1
            
        End With

    Next I
If entrance = 1 Then
    If MaxSpace >= 6004 Then
        If MaxSpace - newTop + 1 > 2 * newTop + 2 Then
            If newTop > 3000 Then
                MaxSpace = 2 * (newTop + 1) + 1
            Else
                MaxSpace = 6003
            End If
            ReDim Preserve PriveSpace(MaxSpace) As item
            
            ReDim Hash(MaxSpace * 2 + 3)
            hashlen = MaxSpace * 2 + 3
            '  reHash...
            toplim = newTop - 1
            For I = 0 To toplim
                place HashFunc2(I), I
            Next I
        Else
            toplim = newTop - 1
        End If
    Else
        toplim = newTop - 1
    End If
End If
entrance = entrance - 1
End Sub

Private Function Malloc() As Long

    If toplim + 1 >= MaxSpace Then
           '' expand hash
           MaxSpace = MaxSpace * 2
           ReDim Preserve PriveSpace(MaxSpace) As item
           If MaxSpace > hashlen Then ExpandHash
    End If
        toplim = toplim + 1
        Malloc = toplim
End Function
Friend Sub Expand(NewMaxSpace As Long)
           MaxSpace = NewMaxSpace
           ReDim Preserve PriveSpace(MaxSpace) As item
           If MaxSpace > hashlen Then ExpandHash
End Sub
Property Get count()
count = toplim + 1
End Property
Friend Function Find2(Key As String, num As Long, Optional checktype As Boolean, Optional isglobal As Boolean) As Boolean
Dim K As Long
num = 0

If Len(Key) = 0 Then Exit Function
K = Hash(HashFunc(Key)) - 1

If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then
    num = PriveSpace(K).nDx: Find2 = True: Index = K:  checktype = PriveSpace(K).originaltype: isglobal = PriveSpace(K).globalvar: Exit Function
    End If
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0
End If
End Function
Friend Function Find3(Key As String, num As Long, isglobal As Boolean) As Boolean
Dim K As Long
num = 0

If Len(Key) = 0 Then Exit Function
K = Hash(HashFunc(Key)) - 1

If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then
    num = PriveSpace(K).nDx: Find3 = True: Index = K: isglobal = PriveSpace(K).globalvar: Exit Function
    End If
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0
End If
End Function
Friend Function FindOlder(Key As String, num As Long, olderfromthat As Long) As Boolean
Dim K As Long
num = 0

If Len(Key) = 0 Then Exit Function
K = Hash(HashFunc(Key)) - 1

If K >= 0 Then
Do
    If PriveSpace(K).Key = Key And K <= olderfromthat Then
    num = PriveSpace(K).nDx: FindOlder = True: Index = K: Exit Function
    End If
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0
End If
End Function
Friend Function Find(Key As String, num As Long) As Boolean
Dim K As Long
num = 0

If Len(Key) = 0 Then Exit Function
K = Hash(HashFunc(Key)) - 1

If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then
    num = PriveSpace(K).nDx: Find = True: Index = K: Exit Function
    End If
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0
End If
End Function
Friend Function ExistKey(Key As String) As Boolean
Dim K As Long

If Len(Key) = 0 Then Exit Function
K = Hash(HashFunc(Key)) - 1

If K >= 0 Then
Do
    If PriveSpace(K).Key = Key Then ExistKey = True: Exit Function
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0
End If
End Function

Friend Function findRebound(Key As String, newnum As Long, v()) As Boolean
Dim K As Long
If Len(Key) = 0 Then Exit Function
K = Hash(HashFunc(Key)) - 1

If K >= 0 Then
Do
With PriveSpace(K)
    If .Key = Key Then
    If .KeyType = 0 Then
        If IsObject(v(.nDx)) Then Set v(.nDx) = Nothing
        If newnum <> .nDx Then v(.nDx) = Empty
        .KeyType = 1
    End If
    .nDx = newnum
     findRebound = True: Exit Function
    End If
    End With
     K = PriveSpace(K).Pleft - 1
Loop Until K < 0
End If
End Function
Friend Sub ItemCreator(Key As String, num As Long, Optional isRef As Boolean = False, Optional isglobal As Boolean = False, Optional deftype As Boolean = False)
Dim A As Long
A = Malloc()
On Error GoTo exxx
place HashFunc(Key), A
With PriveSpace(A)
    .nDx = num
    .KeyType = CByte(Abs(isRef))
    .originaltype = CByte(deftype)
    .globalvar = CByte(isglobal)
    .firsthash = HD(Key)
     SwapStrings .Key, Key
End With
Index = A
exxx:
End Sub
Friend Sub ItemCreator2(id As idHash, Key As String, num As Long, Optional isRef As Boolean = False, Optional isglobal As Boolean = False, Optional deftype As Boolean = False)
Dim A As Long
A = Malloc()
On Error GoTo exxx
place id.lastHashFunc0 Mod hashlen, A
With PriveSpace(A)
    .nDx = num
    .KeyType = CByte(Abs(isRef))
    .originaltype = CByte(deftype)
    .globalvar = CByte(isglobal)
    .firsthash = HD(Key)
     .Key = Key
End With
Index = A
exxx:
End Sub
Property Get Percent()
Percent = count / hashlen
End Property
Private Sub place(ByVal b, ByVal A)
    Dim K As Long
    K = Hash(b)
    If Not Hash(b) = A + 1 Then
         Hash(b) = A + 1
         PriveSpace(A).Pleft = K
    End If
    PriveSpace(A).lastpos = b
End Sub

Private Function HashFunc2(where As Long)
HashFunc2 = PriveSpace(where).firsthash Mod hashlen
End Function

Private Function HashFunc(A$)
HashFunc = HD(A$) Mod hashlen
End Function
Sub flat()
Index = -1
End Sub

Friend Property Get vType(where As Long) As Boolean
    vType = CBool(PriveSpace(where).originaltype)
End Property
Friend Property Let vType(where As Long, RHS As Boolean)
    If RHS Then
        PriveSpace(where).originaltype = 255
    Else
        PriveSpace(where).originaltype = 0
    End If
End Property

