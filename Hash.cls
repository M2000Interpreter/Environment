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
Public index As Long
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
If index < 0 Then
lastNDX = 0
Else
lastNDX = PriveSpace(index).nDx
End If
End Property
Public Property Get lastkey() As String
If index < 0 Then
lastkey = vbCrLf
Else
lastkey = PriveSpace(index).Key
End If
End Property

Private Sub ExpandHash()

    Dim i As Long
    hashlen = hashlen * 2 + 3
    
    ReDim Hash(hashlen) As Long
    For i = 0 To toplim
        place HashFunc2(i), i
    Next i
    
End Sub


Friend Sub ReduceHash(newTop As Long, v())
Dim aa As Variant, ditem As item, ditemcl As item
    ditemcl.nDx = -1
    Dim i As Long
     If toplim < newTop Then Exit Sub
     entrance = entrance + 1
     For i = toplim To newTop Step -1
        ditem = PriveSpace(i)
        PriveSpace(i) = ditemcl
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

    Next i
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
            For i = 0 To toplim
                place HashFunc2(i), i
            Next i
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
Property Get Count()
Count = toplim + 1
End Property
Friend Function Find2(Key As String, num As Long, Optional checktype As Boolean, Optional isglobal As Boolean) As Boolean
Dim k As Long
num = 0

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    num = PriveSpace(k).nDx: Find2 = True: index = k:  checktype = PriveSpace(k).originaltype: isglobal = PriveSpace(k).globalvar: Exit Function
    End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Function Find3(Key As String, num As Long, isglobal As Boolean) As Boolean
Dim k As Long
num = 0

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    num = PriveSpace(k).nDx: Find3 = True: index = k: isglobal = PriveSpace(k).globalvar: Exit Function
    End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Function Find(Key As String, num As Long) As Boolean
Dim k As Long
num = 0

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    num = PriveSpace(k).nDx: Find = True: index = k: Exit Function
    End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Function ExistKey(Key As String) As Boolean
Dim k As Long

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then ExistKey = True: Exit Function
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function

Friend Function findRebound(Key As String, newnum As Long, v()) As Boolean
Dim k As Long
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
With PriveSpace(k)
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
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Sub ItemCreator(Key As String, num As Long, Optional isRef As Boolean = False, Optional isglobal As Boolean = False, Optional deftype As Boolean = False)
Dim a As Long
a = Malloc()
On Error GoTo exxx
place HashFunc(Key), a
With PriveSpace(a)
    .nDx = num
    .KeyType = CByte(Abs(isRef))
    .originaltype = CByte(deftype)
    .globalvar = CByte(isglobal)
    .firsthash = HD(Key)
     SwapStrings .Key, Key
End With
index = a
exxx:
End Sub
Friend Sub ItemCreator2(Id As idHash, Key As String, num As Long, Optional isRef As Boolean = False, Optional isglobal As Boolean = False, Optional deftype As Boolean = False)
Dim a As Long
a = Malloc()
On Error GoTo exxx
place Id.lastHashFunc0 Mod hashlen, a
With PriveSpace(a)
    .nDx = num
    .KeyType = CByte(Abs(isRef))
    .originaltype = CByte(deftype)
    .globalvar = CByte(isglobal)
    .firsthash = HD(Key)
     .Key = Key
End With
index = a
exxx:
End Sub
Property Get Percent()
Percent = Count / hashlen
End Property
Private Sub place(ByVal b, ByVal a)
    Dim k As Long
    k = Hash(b)
    If Not Hash(b) = a + 1 Then
         Hash(b) = a + 1
         PriveSpace(a).Pleft = k
    End If
    PriveSpace(a).lastpos = b
End Sub

Private Function HashFunc2(where As Long)
HashFunc2 = PriveSpace(where).firsthash Mod hashlen
End Function

Private Function HashFunc(a$)
HashFunc = HD(a$) Mod hashlen
End Function
Sub flat()
index = -1
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

