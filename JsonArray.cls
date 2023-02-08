VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' George Karras, Kalithea, Greece
' Version 3.2
' GNU Lesser General Public License version 3
Option Explicit
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Dim ar()
Private curdot$
Private Sub Class_Initialize()
    ReDim ar(0)
End Sub

Friend Property Get Value(n As Long) As Variant
    If n >= 0 And n < UBound(ar) Then
        Value = ar(n)
    End If
End Property
Friend Property Let Value(n As Long, RHS As Variant)
    If n >= 0 And n < UBound(ar) Then
        ar(n) = RHS
    ElseIf n >= UBound(ar) Then
        ReDim Preserve ar(n + 1)
        ar(n) = RHS
    End If
End Property
Friend Property Set ValueObj(n As Long, RHS As Variant)
    If Not IsObject(RHS) Then Err.Raise 8001, , "Not an object"
    If Not TypeOf RHS Is JsonArray Then
        If Not TypeOf RHS Is JsonObject Then
                Err.Raise 8002, , "Not an a proper json object"
        End If
    End If
    If n >= 0 And n < UBound(ar) Then
        Set ar(n) = RHS
    ElseIf n >= UBound(ar) Then
        ReDim Preserve ar(n + 1)
        Set ar(n) = RHS
    End If
End Property

Friend Property Get ValueObj(n As Long) As Variant
    If n >= 0 And n < UBound(ar) Then
        Set ValueObj = ar(n)
    End If
End Property
Property Get ValueIsBoolean(n As Long) As Boolean
    If n >= 0 And n < UBound(ar) Then
        ValueIsBoolean = VarType(ar(n)) = vbBoolean
    End If
End Property
Property Get ValueIsObj(n As Long) As Boolean
    If n >= 0 And n < UBound(ar) Then
        If IsObject(ar(n)) Then
            ValueIsObj = TypeOf ar(n) Is JsonObject
        End If
    End If
End Property
Property Get ValueIsArray(n As Long) As Boolean
    If n >= 0 And n < UBound(ar) Then
        If IsObject(ar(n)) Then
            ValueIsArray = TypeOf ar(n) Is JsonArray
        End If
    End If
End Property
Property Get ValueIsNull(n As Long) As Boolean
    If n >= 0 And n < UBound(ar) Then
        ValueIsNull = VarType(ar(n)) = vbEmpty
    Else
        ValueIsNull = True
    End If
End Property
Property Get NullValue() As Variant
End Property
Property Get count()
    count = UBound(ar)
End Property

Property Get IsEmpty()
    IsEmpty = UBound(ar) = -1
End Property
Private Sub ClearVariant(ByRef VarVar)
Dim t(0 To 3) As Long
   CopyMemory ByVal VarPtr(VarVar), t(0), 16
End Sub
Private Sub SwapVariant(ByRef A As Variant, ByRef b As Variant)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(A), 16
   CopyMemory ByVal VarPtr(A), ByVal VarPtr(b), 16
   CopyMemory ByVal VarPtr(b), t(0), 16
End Sub
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

Sub InsertNull(Index As Long, Optional many As Long = 1)
If many <= 0 Then Exit Sub
Dim W As Long, i As Long
W = count
ReDim Preserve ar(W + many)
For i = W - 1 To Index Step -1
    SwapVariant ar(i + many), ar(i)
    ClearVariant ar(i)
Next

End Sub
Sub DeleteItems(Index As Long, Optional many As Long = 1)
If many <= 0 Then Exit Sub
Dim W As Long, i As Long
If W - many = 0 Then
ReDim ar(0)
Else
W = count

For i = Index + many To W - 1
    SwapVariant ar(i - many), ar(i)
    ClearVariant ar(i)
Next
ReDim Preserve ar(W - many)
End If
End Sub
Sub Assign(Index As Long, RHS)
    If IsObject(RHS) Then
        Set ValueObj(Index) = RHS
    Else
        Value(Index) = RHS
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
    ElseIf n = 0 Then
        If Left$(part$(0), 1) = " " Then
        
            If Not ma.ValType(W) > 3 Then
                If Left$(part$(1), 1) = " " Then
                    ma.Assign W, New JsonArray
                Else
                    ma.Assign W, New JsonObject
                End If
            End If
            Set m = ar(W)
            n = n + 1
            GoTo again
        Else
            Err.Raise 8005, "Not a numeric index:" + part$(0)
        End If
    ElseIf Left$(part$(n), 1) = " " Then
            If Not ma.ValType(W) > 3 Then
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
Function anArray(Index As Long, RHS) As JsonArray
    Dim m As New JsonArray
    If IsObject(RHS) Then
        Set m.ValueObj(Index) = RHS
    Else
        
        m.Value(Index) = RHS
    End If
    Set anArray = m
    
End Function
Function anObject(Key As String, RHS) As JsonObject
    Dim m As New JsonObject
    m.Assign Key, RHS
    Set anObject = m
End Function
Function ValType(Index As Long) As Long
If Not IsEmpty Then
    If Not ValueIsNull(Index) Then  ' 0 for null
        If IsObject(ar(Index)) Then
            If TypeOf ar(Index) Is JsonArray Then
                ValType = 4
            Else
                ValType = 5
            End If
        Else
            Select Case VarType(ar(Index))
            Case 1, 10
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
End Function
Function ToString(Index As Long, Optional dot As String = ".")
Dim W As Long, hlp$, ja As JsonArray, jo As JsonObject
W = ValType(Index)
Select Case W
Case 0
    ToString = "null"
Case 1
    ToString = """" + StringToEscapeStr(ar(Index)) + """"
Case 2
    If ar(Index) Then ToString = "true" Else ToString = "false"
Case 3
    ToString = Replace(ar(Index), curdot$, dot)
Case 4
    Set ja = ar(Index)
    ToString = ja.Json
Case 5
    Set jo = ar(Index)
    ToString = jo.Json
Case 33
        ToString = Replace(ar(Index), curdot$, dot)
End Select
End Function
Private Function ToString2(Index As Long, n As Long, ww As Long)
Dim W As Long, hlp$, ja As JsonArray, jo As JsonObject
W = ValType(Index)
Select Case W
Case 0
    ToString2 = "null"
Case 1
    ToString2 = """" + StringToEscapeStr(ar(Index)) + """"
Case 2
    If ar(Index) Then ToString2 = "true" Else ToString2 = "false"
Case 3
    ToString2 = LTrim$(Str$(ar(Index)))
Case 4
    Set ja = ar(Index)
    ToString2 = ja.Json(n, ww)
Case 5
    Set jo = ar(Index)
    ToString2 = jo.Json(n, ww)
Case 33
    ToString2 = Replace(ar(Index), curdot$, ".")
End Select
End Function
Property Get Json(Optional sp As Long = 0, Optional W As Long = -1) As String
    Dim i As Long, hlp$, acc$, nl$
    sp = Abs(sp)
    If sp Then nl$ = vbCrLf + space$(sp): hlp$ = space$(sp): If W = -1 Then W = sp
    For i = 0 To count - 1
        acc$ = acc$ + hlp$ + nl$ + ToString2(i, sp - W * (sp > 0), W)
        If i = 0 Then
             If sp Then
                hlp$ = ","
            Else
                hlp$ = ", "
            End If
        End If
    Next i
        If sp = 0 Then
    Json = "[" + acc$ + "]"
    Else
    Json = "[" + acc$ + vbCrLf + space(sp - W) + "]"
    End If
End Property
Property Get item(Index As Long)
Attribute item.VB_UserMemId = 0
Select Case ValType(Index)
Case 0
Case 4, 5
    Set item = ar(Index)
Case 33
Dim s As String
    s = ar(Index)
    If InStr(s, curdot$) > 0 Then
        s = Replace(ar(Index), curdot$, ".")
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
    item = ar(Index)
End Select
End Property
Property Let item(Index As Long, RHS)
    Assign Index, RHS
End Property
Property Set item(Index As Long, RHS)
    Assign Index, RHS
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
            mo.Assign "" + p(level), ma.EscapeStrToString(Mid$(A$, markin, i - markin))
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
        ElseIf InStr(1, ch, "e", vbTextCompare) = 0 And ch > 14 Then
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
Function BigNumberOld(ByVal ch As String) As Variant
Dim b() As Byte
   If Len(ch) > 28 Then
            If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            b() = ch
            BigNumberOld = b()
        ElseIf InStr(1, ch, "e", vbTextCompare) = 0 And Len(ch) > 14 Then
            If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            BigNumberOld = CDec(ch)
        Else
            If InStr(ch$, curdot$) > 0 Then Mid$(ch$, InStr(ch$, curdot$)) = "."
            BigNumberOld = val(ch)
        End If
End Function

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


