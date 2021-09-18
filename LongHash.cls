VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LongHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 Private Type item
 Key As Long
 Container As String
 infostr As String
 firsthash As Long
 lastpos As Long
 Pleft As Long  ' a list
 End Type
'
Private PriveSpace() As item
Dim MaxSpace As Long
Dim hashlen As Long
Dim lastkey As Long, lastfind As Long
Dim toplim As Long
Private ParentIndex As Long
Public Done As Boolean
Public index As Long
Private Hash() As Long

Private Sub Class_Initialize()
    MaxSpace = 30
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
End Sub
Public Sub ForceFlush(NewSpace As Long)
    MaxSpace = NewSpace
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
End Sub

Private Sub ExpandHash()
    Dim i As Long
    hashlen = hashlen * 2 + 3
    
    ReDim Hash(hashlen) As Long
    For i = 0 To toplim
    place HashFunc2(i), i
    Next i
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
Property Get count()
count = toplim + 1
End Property
Private Function Find(Key As Long) As Boolean
Dim k As Long
ParentIndex = -1
Done = False
k = Hash(HashFunc(Key)) - 1
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    Find = True: lastfind = Key: index = k: Exit Function
    End If
     ParentIndex = k
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Function ExistKey(Key As Long) As Boolean
Dim k As Long
Done = False
k = Hash(HashFunc(Key)) - 1
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then ExistKey = True: index = k: Done = True: Exit Function
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function

Public Property Get Key() As Long
If index > -1 Then
    Key = PriveSpace(index).Key
End If
End Property
Public Property Get Child() As Boolean
If index > -1 Then
    Child = Hash(PriveSpace(index).lastpos) - 1 <> index
End If
End Property
Public Property Get HasCollision() As Boolean
If index > -1 Then
    HasCollision = (Hash(PriveSpace(index).lastpos) - 1 <> index) Or (PriveSpace(index).Pleft > 0)
End If
End Property
Public Sub AddKey(RHS As Long, Optional aValue As String = vbNullString, Optional aInfo As String = vbNullString)
    index = -1
    lastkey = RHS
    Done = False
    ItemCreator lastkey, aValue, aInfo
    
End Sub

Private Sub ItemCreator(Key As Long, storethis As String, andthis As String)
Dim a As Long, first As Long
Done = False
first = HD(Key)
a = Malloc()
With PriveSpace(a)
    .Key = Key
    .Container = storethis
    .infostr = andthis
    .firsthash = first
place HashFunc1(first), a
End With
End Sub
Public Sub ItemCreator2(Key As Long, storethis As String)
Dim first As Long
first = HD(Key)
index = Malloc()
With PriveSpace(index)
    .Key = Key
    .Container = storethis
    .firsthash = first
place HashFunc1(first), index
End With
End Sub

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
Public Function HD(ByVal v As Long) As Long
       If v = 0 Then v = 1 Else HD = v
End Function
Private Function HashFunc1(readyhash)
HashFunc1 = readyhash Mod hashlen
End Function
Private Function HashFunc(a As Long)
HashFunc = HD(a) Mod hashlen
End Function

Property Get Value() As String
If index = -1 Then
Else
    Value = PriveSpace(index).Container
End If
End Property
Property Let ValueStr(RHS As String)
If index = -1 Then
Else
    PriveSpace(index).Container = RHS
End If
End Property

Property Get StrPointer() As Long
If index = -1 Then
Else
    StrPointer = StrPtr(PriveSpace(index).Container)
End If
End Property

Property Get Info() As String
If index = -1 Then
Else
    Info = PriveSpace(index).infostr
End If
End Property
Property Get Percent()
Percent = 100 * count / hashlen
End Property


