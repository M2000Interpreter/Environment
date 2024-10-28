VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "idHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 Private Type item
    Key As String
    nDx As Long
    firsthash As Long
    lastpos As Long
    Pleft As Long  ' a list
 End Type
'
Private PriveSpace() As item
Dim MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, RetVal As Integer)
Private lastfind As String
Private Hash() As Long
Public Index As Long
Public lastHashFunc0 As Long
Private limitGlobal As Long
Private clean As New mStiva2
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


Private Sub ExpandHash()

    Dim i As Long
    hashlen = hashlen * 2 + 3
    
    ReDim Hash(hashlen) As Long
    For i = 0 To toplim

    place HashFunc2(i), i
    Next i
End Sub
Friend Sub ReduceHash(newTop As Long)
    Dim i As Long
If newTop = 0 Or newTop > toplim Then Exit Sub
    For i = toplim To newTop Step -1
       
       With PriveSpace(i)
        
    
        If .Pleft > 0 Then
            Hash(.lastpos) = .Pleft
            .Key = vbNullString
            .Pleft = 0
        Else
            Hash(.lastpos) = 0
            .Key = vbNullString
        End If
        .nDx = 0
            
        End With

    Next i
    If newTop <= 50 Then
            If MaxSpace > 50 Then
                        MaxSpace = 50
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
            If MaxSpace - newTop + 1 > 2 * newTop + 2 Then
            
                MaxSpace = 2 * (newTop + 1) + 1
                
                ReDim Preserve PriveSpace(MaxSpace) As item
                hashlen = MaxSpace * 2 + 3
                ReDim Hash(hashlen)
               
             '  reHash...
             toplim = newTop - 1
            For i = 0 To toplim
                place HashFunc2(i), i
            Next i
        
            Else
 toplim = newTop - 1
    End If
   End If
End Sub
Function RenameKey(Key As String, NewKey As String)
Dim k As Long

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    
    PriveSpace(k).Key = NewKey
    PriveSpace(k).firsthash = HD(NewKey)
                ReDim Hash(MaxSpace * 2 + 3)
               hashlen = MaxSpace * 2 + 3
             '  reHash...
            For k = 0 To toplim
                place HashFunc2(k), k
            Next k

    
            RenameKey = True
            Exit Function
               
 End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0

 End If
            
End Function
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
Property Get LastKnown() As String
LastKnown = lastfind
End Property

Friend Function Find(Key As String, num As Long, Optional badindex As Long = 0) As Boolean
Dim k As Long

num = 0
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    num = PriveSpace(k).nDx
    If num > 0 Then
        Find = True: lastfind = Key: Index = k: Exit Function
    ElseIf badindex Then
    
       If badindex <= k Then Find = True: lastfind = Key: Index = k: Exit Function
       If k >= limitGlobal Then If num < -1 Then Find = True: lastfind = Key: Index = k: Exit Function
    End If
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
Friend Sub ItemCreator2(Key As String, num As Long)
Dim A As Long
A = Malloc()
place HashFuncStored(Key), A
With PriveSpace(A)
    .Key = Key
    .nDx = num
    .firsthash = HD(Key)
End With


End Sub

Friend Sub ItemCreator(Key As String, num As Long)
Dim A As Long
A = Malloc()
place HashFunc(Key), A
With PriveSpace(A)
    .Key = Key
    .nDx = num
    .firsthash = HD(Key)
End With

End Sub
Property Get Percent()
Percent = count / hashlen
End Property
Private Sub place(ByVal b, ByVal A)
    Dim k As Long
    k = Hash(b)
    If Not Hash(b) = A + 1 Then
         Hash(b) = A + 1
         PriveSpace(A).Pleft = k
    End If
    PriveSpace(A).lastpos = b
End Sub

Private Function HashFunc2(where As Long)
HashFunc2 = PriveSpace(where).firsthash Mod hashlen
End Function

Private Function HashFunc(A$)
lastHashFunc0 = HD(A$)
HashFunc = lastHashFunc0 Mod hashlen
End Function

Private Function HashFuncStored(A$)
HashFuncStored = lastHashFunc0 Mod hashlen
End Function

Public Sub pushtopGlobal()
clean.PushLong limitGlobal
clean.PushLong toplim + 1
limitGlobal = toplim + 1
End Sub
Public Sub poptopGlobal()
ReduceHash CLng(clean.PopVal)
limitGlobal = CLng(clean.PopVal)
End Sub

Public Sub pushtop()
clean.PushLong toplim + 1
End Sub
Public Sub poptop()
 ReduceHash CLng(clean.PopVal)
End Sub
