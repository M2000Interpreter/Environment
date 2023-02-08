VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' George Karras, Kalithea, Greece
' version 3.1
' GNU Lesser General Public License version 3

Option Explicit
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function HashData Lib "shlwapi" (ByVal straddr As Long, ByVal ByteSize As Long, ByVal res As Long, ByVal ressize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type item
    Key As String
    iValue As Variant
    firsthash As Long
    lastpos As Long
    Pleft As Long  ' a list
End Type
Private PriveSpace() As item
Dim MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Private lastkey As String
Private Hash() As Long
Public Done As Boolean
Public Index As Long
Private ParentIndex As Long
Private entrance As Long
Private curdot$
Public Function HD(A$) As Long
Dim ret As Long
ret = HashData(StrPtr(A$), LenB(A$), VarPtr(HD), 4)
HD = HD And &H7FFFFFFF
If HD = 0 Then HD = 1
End Function

Private Sub Class_Initialize()
    MaxSpace = 5
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
    curdot$ = GetDeflocaleString(14)
End Sub
Private Sub AddKey(RHS, Optional aValue As Variant)
    Index = -1
    lastkey = Normalize(RHS)
    If Len(lastkey) = 0 Then RHS = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0): lastkey = RHS
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
Sub DeleteKey(RHS)
    If Find(RHS) Then
        RemoveOne
    End If
End Sub
Friend Function Find(RHS) As Boolean
Dim k As Long, Key As String, k1 As Long
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
ParentIndex = -1
Done = False
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then Find = True: Index = k: Done = True: Exit Function
     ParentIndex = k
     k1 = k
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0 Or k1 = k
If k1 = k Then
Rehash
k = Hash(HashFunc(Key)) - 1
k1 = -2
ParentIndex = -1
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then Find = True: Index = k: Done = True: Exit Function
     ParentIndex = k
     k1 = k
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0 Or k = k1
End If
End If
End If

End Function

Friend Sub RemoveOne()
Dim new_item As item
Dim k As Long, k1 As Long, vvv As Variant, top1 As Long
If toplim < 0 Then Exit Sub
top1 = toplim
If Done Then
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
        k = Hash(PriveSpace(Index).lastpos) - 1
        If k = top1 Then
            ' we have to give the new position to Hash()
            Hash(PriveSpace(Index).lastpos) = Index + 1
        Else
            Do While k >= 0 And k <> top1 And k <> k1
                   k1 = k
                   k = PriveSpace(k).Pleft - 1
            Loop
            If k = k1 Then
               Rehash
            ElseIf k = top1 Then
               PriveSpace(k1).Pleft = Index + 1
            End If
        End If
  
here:
       toplim = toplim - 1
      
        
       
        
  ReduceHash toplim
End If
End Sub
Public Sub drop(RHS)
RHS = RHS - 1
If (toplim - RHS) >= 0 And (RHS >= 0) Then
Dim i As Long
For i = toplim To RHS Step -1
       With PriveSpace(i)
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
Next i
Else
   MaxSpace = 5
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
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

Property Get Value() As Variant
Done = False
If Not Index = -1 Then
    With PriveSpace(Index)
        If Not IsObject(.iValue) Then
            Done = True
                Value = .iValue
        End If
    End With
End If
End Property
Property Get NullValue() As Variant
End Property

Property Get ValueObj() As Variant
Done = False
If Not Index = -1 Then
    With PriveSpace(Index)
        If IsObject(.iValue) Then
            Set ValueObj = .iValue
            Done = True
        End If
    End With
End If

End Property
Private Sub ExpandHash()
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
            Normalize = Key
        ElseIf VarType(Key) = vbBoolean Then
        Normalize = Str(CLng(Key))
        Else
        Normalize = LTrim$(Str$(Key))
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
Property Get count()
    count = toplim + 1
End Property

Property Get IsEmpty()
    IsEmpty = toplim = -1
End Property
Function IsObj() As Boolean
If Index = -1 Then
Else
IsObj = IsObject(PriveSpace(Index).iValue)
End If
End Function


Function ExistKey(RHS) As Boolean
Dim k As Long, k1 As Long, Key As String
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
    Do
        If PriveSpace(k).Key = Key Then ExistKey = True: Index = k: Exit Function
        k1 = k
        k = PriveSpace(k).Pleft - 1
        
    Loop Until k < 0 Or k = k1
    If k = k1 Then
        Rehash
        k = Hash(HashFunc(Key)) - 1
        If k >= 0 Then
        Do
            If PriveSpace(k).Key = Key Then ExistKey = True: Index = k: Exit Function
            k1 = k
            k = PriveSpace(k).Pleft - 1
        Loop Until k < 0 Or k = k1
        End If
    End If
End If
End Function
Private Function ExistKey2(Key As String, nkey As Long) As Boolean
Dim k As Long, k1 As Long
If Len(Key) = 0 Then ExistKey2 = True: Exit Function
k = Hash(HashFunc1(nkey)) - 1
k1 = -2
If k >= 0 Then
    Do
        If PriveSpace(k).Key = Key Then ExistKey2 = True: Exit Function
        k1 = k
        k = PriveSpace(k).Pleft - 1
    Loop Until k < 0 Or k = k1
    If k = k1 Then
        Rehash
        k = Hash(HashFunc(Key)) - 1
        If k >= 0 Then
        Do
            If PriveSpace(k).Key = Key Then ExistKey2 = True: Exit Function
            k1 = k
            k = PriveSpace(k).Pleft - 1
            
        Loop Until k < 0 Or k = k1
        End If
    End If

End If
End Function

Private Sub ItemCreator(Key As String, nValue As Variant)
Dim A As Long, kk As Long
Done = False
kk = HD(Key)
If ExistKey2(Key, kk) Then Exit Sub  ' done is false
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
Property Get Percent()
Percent = 100 * count / hashlen
End Property
Private Sub place(ByVal b, ByVal A)
    Dim k As Long
    k = Hash(b)
    If Not k = A + 1 Then
         Hash(b) = A + 1
         PriveSpace(A).Pleft = k
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


Public Property Get KeyToString() As String
If Index > -1 Then

    KeyToString = PriveSpace(Index).Key
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
Public Property Get KeyToStringPos(there As Long) As String
    KeyToStringPos = PriveSpace(there).Key
End Property

Public Function StringToEscapeStr(RHS As Variant) As Variant
Dim i As Long, cursor As Long, ch As String
cursor = 0
Dim del As String
Dim H9F As String
For i = 1 To Len(RHS)
                ch = Mid$(RHS, i, 1)
                cursor = cursor + 1
                Select Case AscW(ch)
                    Case 92:   ch = "\\"
                    Case 34:   ch = "\"""
                    Case 10:   ch = "\n"
                    Case 13:   ch = "\r"
                    Case 9:    ch = "\t"
                    Case 8:    ch = "\b"
                    Case 12: ch = "\f"
                    Case 0 To 31, 127 To &H9F
                        ch = "\u" & Right$("000" & Hex$(AscW(ch)), 4)
                    Case Is > 255
                       ch = "\u" & Right$("000" & Hex$(AscW(ch)), 4)
                End Select
                If cursor + Len(ch) > Len(StringToEscapeStr) Then StringToEscapeStr = StringToEscapeStr + space$(500)
                Mid$(StringToEscapeStr, cursor, Len(ch)) = ch
                cursor = cursor + Len(ch) - 1
Next
If cursor > 0 Then StringToEscapeStr = Left$(StringToEscapeStr, cursor)

End Function
Public Function EscapeStrToString(ByVal RHS As String) As String
Dim i As Long, cursor As Long, ch As String
     For cursor = 1 To Len(RHS)
        ch = Mid$(RHS, cursor, 1)
        i = i + 1
        Select Case ch
            Case """": GoTo ok1
            Case "\":
                cursor = cursor + 1
                ch = Mid$(RHS, cursor, 1)
                Select Case LCase$(ch) 'We'll make this forgiving though lowercase is proper.
        '            Case "\", "/": ch = ch
                    Case """":      ch = """"
                    Case "a":       ch = Chr$(7)
                    Case "n":      ch = vbLf
                    Case "r":      ch = vbCr
                    Case "t":      ch = vbTab
                    Case "b":      ch = vbBack
                    Case "f":      ch = vbFormFeed
                    Case "u":      ch = ParseHexChar(RHS, cursor, Len(RHS))
                End Select
        End Select
                If i + Len(ch) > Len(EscapeStrToString) Then EscapeStrToString = EscapeStrToString + space$(500)
                Mid$(EscapeStrToString, i, Len(ch)) = ch
                i = i + Len(ch) - 1
    Next
ok1:
    If i > 0 Then EscapeStrToString = Left$(EscapeStrToString, i)
End Function

Private Function ParseHexChar( _
    ByRef Text As String, _
    ByRef cursor As Long, _
    ByVal LenOfText As Long) As String
    
    Const ASCW_OF_ZERO As Long = &H30&
    Dim Length As Long
    Dim ch As String
    Dim DigitValue As Long
    Dim Value As Long

    For cursor = cursor + 1 To LenOfText
        ch = Mid$(Text, cursor, 1)
        Select Case ch
            Case "0" To "9", "A" To "F", "a" To "f"
                Length = Length + 1
                If Length > 4 Then Exit For
                If ch > "9" Then
                    DigitValue = (AscW(ch) And &HF&) + 9
                Else
                    DigitValue = AscW(ch) - ASCW_OF_ZERO
                End If
                Value = Value * &H10& + DigitValue
            Case Else
                Exit For
        End Select
    Next
    If Length = 0 Then Err.Raise 5 'No hex digits at all.
    cursor = cursor - 1
    ParseHexChar = ChrW$(Value)
End Function

Function anObject(Key As String, RHS) As JsonObject
    Dim m As New JsonObject
    m.Assign Key, RHS
    Set anObject = m
End Function
Function anArray(Index As Long, RHS) As JsonArray
    Dim m As New JsonArray
    If IsObject(RHS) Then
        Set m.ValueObj(Index) = RHS
    Else
        m.Value(Index) = RHS
    End If
    Set anArray = m
    
End Function
Property Get ValueIsNull(Key As String) As Boolean
    If ExistKey(Key) Then
        ValueIsNull = VarType(PriveSpace(Index).iValue) = vbEmpty
    Else
        ValueIsNull = True
    End If
End Property
Property Get ValueIsObj(Key As String) As Boolean
    If ExistKey(Key) Then
        If IsObject(PriveSpace(Index).iValue) Then
            ValueIsObj = TypeOf PriveSpace(Index).iValue Is JsonObject
        End If
    End If
End Property
Property Get ValueIsArray(Key As String) As Boolean
    If ExistKey(Key) Then
        If IsObject(PriveSpace(Index).iValue) Then
            ValueIsArray = TypeOf PriveSpace(Index).iValue Is JsonArray
        End If
    End If
End Property
Property Get ValueIsBoolean(Key As String) As Boolean
    If ExistKey(Key) Then
        ValueIsBoolean = VarType(PriveSpace(Index).iValue) = vbBoolean
    End If
End Property
Function ValType(Key As String) As Long
If Not IsEmpty Then
    If ExistKey(Key) Then
        If Not VarType(PriveSpace(Index).iValue) = vbEmpty Then
            If IsObject(PriveSpace(Index).iValue) Then
                If TypeOf PriveSpace(Index).iValue Is JsonArray Then
                    ValType = 4
                Else
                    ValType = 5
                End If
            Else
                Select Case VarType(PriveSpace(Index).iValue)
                Case 0, 1, 10
                Case vbString
                    ValType = 1
                Case vbBoolean
                    ValType = 2
                Case 8209
                    ValType = 33
                Case Else
                    ValType = 3
                End Select
            End If
        End If
    End If
End If
End Function
Private Function ValType2() As Long
If Not VarType(PriveSpace(Index).iValue) = vbEmpty Then
    If IsObject(PriveSpace(Index).iValue) Then
        If TypeOf PriveSpace(Index).iValue Is JsonArray Then
            ValType2 = 4
        Else
            ValType2 = 5
        End If
    Else
        Select Case VarType(PriveSpace(Index).iValue)
        Case 1, 10
        Case vbString
            ValType2 = 1
        Case vbBoolean
            ValType2 = 2
        Case 8209
            ValType2 = 33
        Case Else
            ValType2 = 3
        End Select
    End If
End If
End Function

Sub Assign(Key As String, RHS)
    If ExistKey(Key) Then
        If IsObject(RHS) Then
            Set PriveSpace(Index).iValue = RHS
        Else
            PriveSpace(Index).iValue = RHS
        End If
    Else
        AddKey Key, RHS
    End If
End Sub
Sub AssignPath(path$, RHS, Optional sep As String = ".")
Dim part$(), W As Long, n As Long, s As Long
part$ = Split2(path$, sep)
Dim m As Object, ma As JsonArray, mo As JsonObject
s = UBound(part$)
n = 0
Set m = Me
again:
If n > UBound(part$) Then Exit Sub
If TypeOf m Is JsonArray Then
    Set ma = m
    W = val(part$(n))
    If n = s Then
        ma.Assign W, RHS
    ElseIf Left$(part$(n), 1) = " " Then
            If ma.count <= W Or Not ma.ValType(W) > 3 Then
                    If Left$(part$(n + 1), 1) = " " Then
                        ma.Assign W, New JsonArray
                    Else
                        ma.Assign W, New JsonObject
                    End If
            End If
            Set m = ma(W)
            n = n + 1
            GoTo again
    Else
           Err.Raise 8005, "Not a numeric index:" + part$(n)
    End If
Else
    Set mo = m
    If n = s Then
        mo.Assign part$(n), RHS
    ElseIf n = 0 Then
        If Not Left$(part$(0), 1) = " " Then
            If Not ExistKey(part$(0)) Then
                If Left$(part$(n + 1), 1) = " " Then
                mo.Assign part$(n), New JsonArray
                Else
                mo.Assign part$(n), New JsonObject
                End If
            ElseIf Not IsObj() Then
                If Left$(part$(n + 1), 1) = " " Then
                    mo.Assign part$(n), New JsonArray
                Else
                    mo.Assign part$(n), New JsonObject
                End If
             End If
             Set m = mo(part$(0))
            n = n + 1
            GoTo again
        Else
            Err.Raise 8005, "Not a key:" + part$(0)
        End If
    ElseIf Not Left$(part$(n), 1) = " " Then
        If mo.ExistKey(part$(n)) Then
            If Not mo.ValType(part$(n)) > 3 Then
              If Left$(part$(n + 1), 1) = " " Then
                mo.Assign part$(n), New JsonArray
                Else
                mo.Assign part$(n), New JsonObject
                End If
            End If
        Else
                If Left$(part$(n + 1), 1) = " " Then
                mo.Assign part$(n), New JsonArray
                Else
                mo.Assign part$(n), New JsonObject
                End If
        End If
        Set m = mo(part$(n))
        n = n + 1
        GoTo again
    Else
        Err.Raise 8005, "Not an object key:" + part$(n)
    End If
End If
End Sub
Function ToString(Key As String, Optional dot As String = ".")
Dim W As Long, hlp$, ja As JsonArray, jo As JsonObject
W = ValType(Key)
Select Case W
Case 0
    ToString = "null"
Case 1
    ToString = """" + StringToEscapeStr(PriveSpace(Index).iValue) + """"
Case 2
    If PriveSpace(Index).iValue Then ToString = "true" Else ToString = "false"
Case 3
    ToString = Replace(PriveSpace(Index).iValue, curdot$, dot)
Case 4
    Set ja = PriveSpace(Index).iValue
    ToString = ja.Json
Case 5
    Set jo = PriveSpace(Index).iValue
    ToString = jo.Json
Case 33
    ToString = Replace(PriveSpace(Index).iValue, curdot$, dot)
End Select
End Function
Private Function ToString2(n As Long, ww As Long)
Dim W As Long, hlp$, ja As JsonArray, jo As JsonObject
W = ValType2()
Select Case W
Case 0
    ToString2 = "null"
Case 1
    ToString2 = """" + StringToEscapeStr(PriveSpace(Index).iValue) + """"
Case 2
    If PriveSpace(Index).iValue Then ToString2 = "true" Else ToString2 = "false"
Case 3
    ToString2 = Trim$(Str$(PriveSpace(Index).iValue))
Case 4
    Set ja = PriveSpace(Index).iValue
    ToString2 = ja.Json(n, ww)
Case 5
    Set jo = PriveSpace(Index).iValue
    ToString2 = jo.Json(n, ww)
Case 33
    ToString2 = Replace(PriveSpace(Index).iValue, curdot$, ".")
End Select
End Function
Property Get Json(Optional sp As Long = 0, Optional W As Long = -1) As String
    Dim hlp$, acc$, nl$
    sp = Abs(sp)
    If sp Then nl$ = vbCrLf + space$(sp): hlp$ = space$(sp): If W = -1 Then W = sp
    For Index = 0 To count - 1
        acc$ = acc$ + hlp$ + nl$ + """" + StringToEscapeStr(PriveSpace(Index).Key) + """" + " : " + ToString2(sp - W * (sp > 0), W)
        If Index = 0 Then
            If sp Then
                hlp$ = ","
            Else
                hlp$ = ", "
            End If
        End If
        
    Next
    
    If sp = 0 Then
    Json = "{" + acc$ + "}"
    Else
    Json = "{" + acc$ + vbCrLf + space(sp - W) + "}"
    End If
    Index = -1
End Property
Property Get item(Key As String)
Attribute item.VB_UserMemId = 0
If ExistKey(Key) Then
    Select Case ValType2()
    Case 0
    Case 4, 5
        Set item = PriveSpace(Index).iValue
    Case 33
Dim s As String
    s = PriveSpace(Index).iValue
    If InStr(s, curdot$) > 0 Then
        s = Replace(PriveSpace(Index).iValue, curdot$, ".")
    End If
    If InStr(s, ".") > 25 Then
        item = val(s)
    Else
        On Error Resume Next
        item = CDec(Replace(s, ".", curdot$))
        If Err Then
            Err.Clear
            item = val(s)
        End If
    End If
    Case Else
        item = PriveSpace(Index).iValue
    End Select
End If
End Property
Property Let item(Key As String, RHS)
    Assign Key, RHS

End Property
Property Set item(Key As String, RHS)
    Assign Key, RHS
End Property

Property Get ItemPath(path$, Optional sep As String = ".")
Dim part$(), W As Long, n As Long, ma As JsonArray, mo As JsonObject
part$ = Split2(path$, sep)
Dim m As Object

n = 0
Set m = Me
again:
If n > UBound(part$) Then Exit Property
If TypeOf m Is JsonArray Then
    Set ma = m
    W = val(part$(n))
    Select Case ma.ValType(W)
    Case 0
    Case 4, 5
        Set m = ma(W)
        n = n + 1
        If n <= UBound(part$) Then GoTo again
        Set ItemPath = m
    Case Else
        ItemPath = ma(W)
    End Select
Else
    Set mo = m
    Select Case mo.ValType(part$(n))
    Case 0
    Case 4, 5
        Set m = mo(part$(n))
        n = n + 1
        If n <= UBound(part$) Then GoTo again
        Set ItemPath = m
    Case Else
        ItemPath = mo(part$(n))
    End Select
End If
End Property
Private Function Split2(A$, Optional sep As String = ".") As String()
Dim s() As String, i As Long, look As Boolean, many As Long, lastsep As Long
Dim strip As Boolean
' pass one
ReDim s(0)
If Len(A$) = 0 Then Split2 = s: Exit Function
many = 0
look = True
For i = 1 To Len(A$)
    Select Case Mid$(A$, i, 1)
    Case "["
        If lastsep = i - 1 Then look = False
    Case "]"
        look = True
    Case sep
        If look Then many = many + 1: lastsep = i
    End Select
Next i
ReDim s(many)
lastsep = 0
many = 0
look = True
For i = 1 To Len(A$)
    Select Case Mid$(A$, i, 1)
    Case "["
        If lastsep = i - 1 Then look = False: strip = True
    Case "]"
        look = True
    Case sep
        If look Then
            If strip Then
                s(many) = Trim$(Mid$(A$, lastsep + 2, i - lastsep - 3))
            Else
                s(many) = Trim$(Mid$(A$, lastsep + 1, i - lastsep - 1))
                If IsNumeric(s(many)) Then s(many) = " " + s(many)
            End If
            strip = False
            lastsep = i: many = many + 1
        End If
    End Select
Next i

If strip Then
    s(many) = Trim$(Mid$(A$, lastsep + 2, i - lastsep - 3))
Else
    s(many) = Trim$(Mid$(A$, lastsep + 1, i - lastsep - 1))
    If IsNumeric(s(many)) Then s(many) = " " + s(many)
End If
Split2 = s
End Function
Function Parser(A$) As Object
Dim i As Long, jumpstring As Long, ch As String, level As Long, v
Dim b() As Byte
ReDim b(0)
Dim ma As JsonArray, mo As JsonObject
Dim markin As Long
Dim p()
ReDim p(100), mm(100)
level = -1
For i = 1 To Len(A$)
ch = Mid$(A$, i, 1)
If ch = "{" Then
    level = level + 1
    If level > UBound(p) Then ReDim Preserve p(level * 2): ReDim Preserve mm(level * 2)
    p(level) = ""
    Set mm(level) = New JsonObject
    Set mo = mm(level)
ElseIf ch = "}" Then
    If ma Is mm(level) Then Exit Function
    level = level - 1
    If level < 0 Then Set Parser = mo: Exit Function
    If TypeOf mm(level) Is JsonArray Then
    Set ma = mm(level)
        ma.Assign 0 + p(level), mm(level + 1)
    Else
        Set mo = mm(level)
        mo.Assign "" + p(level), mm(level + 1)
    End If
ElseIf ch = "[" Then
    level = level + 1
    If level > UBound(p) Then ReDim Preserve p(level * 2): ReDim Preserve mm(level * 2)
    p(level) = 0
    Set mm(level) = New JsonArray
    Set ma = mm(level)
ElseIf ch = "]" Then
    If mo Is mm(level) Then Exit Function
    level = level - 1
    If level < 0 Then Set Parser = ma:  Exit Function
    If TypeOf mm(level) Is JsonArray Then
    Set ma = mm(level)
        ma.Assign 0 + p(level), mm(level + 1)
    Else
        Set mo = mm(level)
        mo.Assign "" + p(level), mm(level + 1)
    End If
ElseIf ch = """" Then
    markin = i + 1
    i = i + 1
    ch = Mid$(A$, i, 1)
    While i < Len(A$) And ch <> """"
        If ch = "\" Then i = i + 1
        i = i + 1
        ch = Mid$(A$, i, 1)
    Wend
    If ch = """" Then
        If ma Is mm(level) Then
            ma.Assign 0 + p(level), ma.EscapeStrToString(Mid$(A$, markin, i - markin))
        ElseIf p(level) = "" Then
            p(level) = mo.EscapeStrToString(Mid$(A$, markin, i - markin))
            If p(level) = "" Then Exit Function
            While i < Len(A$) And ch <> ":"
                i = i + 1
                ch = Mid$(A$, i, 1)
            Wend
            If ch <> ":" Then Exit Function
        Else
            mo.Assign "" + p(level), mo.EscapeStrToString(Mid$(A$, markin, i - markin))
        End If
    Else
        Exit Function
    End If
ElseIf ch = "," Then
    If ma Is mm(level) Then p(level) = p(level) + 1 Else p(level) = ""
ElseIf ch = "t" Then
    If Mid$(A$, i, 4) = "true" Then
        i = i + 3
        If ma Is mm(level) Then
            ma.Assign 0 + p(level), True
        Else
            mo.Assign "" + p(level), True
        End If
    Else
        Exit Function
    End If
ElseIf ch = "f" Then
    If Mid$(A$, i, 5) = "false" Then
        i = i + 4
        If ma Is mm(level) Then
            ma.Assign 0 + p(level), False
        Else
            mo.Assign "" + p(level), False
        End If
    Else
        Exit Function
    End If
ElseIf ch = "n" Then
    If Mid$(A$, i, 4) = "null" Then
        i = i + 3
        If ma Is mm(level) Then
            ma.Assign 0 + p(level), ma.NullValue
        Else
            mo.Assign "" + p(level), mo.NullValue
        End If
    Else
        Exit Function
    End If
ElseIf ch Like "[+-0123456789.]" Then
    markin = i
    i = i + 1
    Do While i <= Len(A$)
        ch = Mid$(A$, i, 1)
        Select Case ch
        Case Is < "!", "]", "}", ","
        If Not IsNumeric(Mid$(A$, markin, i - markin)) Then Exit Function
        i = i - 1
        Exit Do
        End Select
        i = i + 1
    Loop
    If i > Len(A$) Then Exit Function
    ch$ = Mid$(A$, markin, i - markin + 1)
    If ma Is mm(level) Then
        If Len(ch) > 28 Then
            If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            b() = ch
            v = b()
            ma.Assign 0 + p(level), v
        ElseIf InStr(1, ch, "e", vbTextCompare) = 0 And Len(ch) > 14 Then
        If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            ma.Assign 0 + p(level), CDec(ch)
        Else
            ma.Assign 0 + p(level), val(ch)
        End If
    Else
        If Len(ch) > 28 Then
            If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            b() = ch
            v = b()
            mo.Assign "" + p(level), v
        ElseIf InStr(1, ch, "e", vbTextCompare) = 0 And Len(ch) > 14 Then
            If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            mo.Assign "" + p(level), CDec(ch)
        Else
            mo.Assign "" + p(level), val(ch)
        End If
    End If
End If
Next
End Function
Public Function GetDeflocaleString(ByVal this As Long) As String
On Error GoTo 1234
    Dim Buffer As String, ret&, R&
    Buffer = String$(514, 0)
    ret = GetLocaleInfoW(0, this, StrPtr(Buffer), Len(Buffer))
    GetDeflocaleString = Left$(Buffer, ret - 1)
1234:
End Function
Property Get DotChar() As String
    DotChar = curdot$
End Property
Function BigNumber(ByVal ch As String) As Variant
Dim b() As Byte
   If Len(ch) > 28 Then
            If InStr(ch$, ".") > 0 And curdot$ <> "." Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            b() = ch
            BigNumber = b()
        ElseIf InStr(1, ch, "e", vbTextCompare) = 0 And Len(ch) > 14 Then
            If InStr(ch$, ".") > 0 And curdot$ <> "." Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            BigNumber = CDec(ch)
        Else
            If curdot$ <> "." Then If InStr(ch$, curdot$) > 0 Then Mid$(ch$, InStr(ch$, curdot$)) = "."
            BigNumber = val(ch)
        End If
End Function


