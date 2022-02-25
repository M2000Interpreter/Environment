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
Attribute VB_Exposed = False
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
Private Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(a), 16
   CopyMemory ByVal VarPtr(a), ByVal VarPtr(b), 16
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

Sub InsertNull(index As Long, Optional many As Long = 1)
If many <= 0 Then Exit Sub
Dim w As Long, i As Long
w = count
ReDim Preserve ar(w + many)
For i = w - 1 To index Step -1
    SwapVariant ar(i + many), ar(i)
    ClearVariant ar(i)
Next

End Sub
Sub DeleteItems(index As Long, Optional many As Long = 1)
If many <= 0 Then Exit Sub
Dim w As Long, i As Long
If w - many = 0 Then
ReDim ar(0)
Else
w = count

For i = index + many To w - 1
    SwapVariant ar(i - many), ar(i)
    ClearVariant ar(i)
Next
ReDim Preserve ar(w - many)
End If
End Sub
Sub Assign(index As Long, RHS)
    If IsObject(RHS) Then
        Set ValueObj(index) = RHS
    Else
        Value(index) = RHS
    End If
End Sub
Sub AssignPath(Path$, RHS, Optional sep As String = ".")
Dim part$(), w As Long, n As Long, s As Long

part$ = Split2(Path$, sep)
Dim m As Object, ma As JsonArray, mo As JsonObject
s = UBound(part$)
n = 0
Set m = Me
again:
If n > UBound(part$) Then Exit Sub
If TypeOf m Is JsonArray Then
    Set ma = m
    w = val(part$(n))
    If n = s Then
        ma.Assign w, RHS
    ElseIf n = 0 Then
        If Left$(part$(0), 1) = " " Then
        
            If Not ma.ValType(w) > 3 Then
                If Left$(part$(1), 1) = " " Then
                    ma.Assign w, New JsonArray
                Else
                    ma.Assign w, New JsonObject
                End If
            End If
            Set m = ar(w)
            n = n + 1
            GoTo again
        Else
            Err.Raise 8005, "Not a numeric index:" + part$(0)
        End If
    ElseIf Left$(part$(n), 1) = " " Then
            If Not ma.ValType(w) > 3 Then
                If Left$(part$(n + 1), 1) = " " Then
                ma.Assign w, New JsonArray
                Else
                ma.Assign w, New JsonObject
                End If
            End If
            Set m = ma(w)
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
Function anArray(index As Long, RHS) As JsonArray
    Dim m As New JsonArray
    If IsObject(RHS) Then
        Set m.ValueObj(index) = RHS
    Else
        
        m.Value(index) = RHS
    End If
    Set anArray = m
    
End Function
Function anObject(Key As String, RHS) As JsonObject
    Dim m As New JsonObject
    m.Assign Key, RHS
    Set anObject = m
End Function
Function ValType(index As Long) As Long
If Not IsEmpty Then
    If Not ValueIsNull(index) Then  ' 0 for null
        If IsObject(ar(index)) Then
            If TypeOf ar(index) Is JsonArray Then
                ValType = 4
            Else
                ValType = 5
            End If
        Else
            Select Case VarType(ar(index))
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
Function ToString(index As Long, Optional dot As String = ".")
Dim w As Long, hlp$, ja As JsonArray, jo As JsonObject
w = ValType(index)
Select Case w
Case 0
    ToString = "null"
Case 1
    ToString = """" + StringToEscapeStr(ar(index)) + """"
Case 2
    If ar(index) Then ToString = "true" Else ToString = "false"
Case 3
    ToString = Replace(ar(index), curdot$, dot)
Case 4
    Set ja = ar(index)
    ToString = ja.Json
Case 5
    Set jo = ar(index)
    ToString = jo.Json
Case 33
        ToString = Replace(ar(index), curdot$, dot)
End Select
End Function
Private Function ToString2(index As Long, n As Long, ww As Long)
Dim w As Long, hlp$, ja As JsonArray, jo As JsonObject
w = ValType(index)
Select Case w
Case 0
    ToString2 = "null"
Case 1
    ToString2 = """" + StringToEscapeStr(ar(index)) + """"
Case 2
    If ar(index) Then ToString2 = "true" Else ToString2 = "false"
Case 3
    ToString2 = LTrim$(Str$(ar(index)))
Case 4
    Set ja = ar(index)
    ToString2 = ja.Json(n, ww)
Case 5
    Set jo = ar(index)
    ToString2 = jo.Json(n, ww)
Case 33
    ToString2 = Replace(ar(index), curdot$, ".")
End Select
End Function
Property Get Json(Optional sp As Long = 0, Optional w As Long = -1) As String
    Dim i As Long, hlp$, acc$, nl$
    sp = Abs(sp)
    If sp Then nl$ = vbCrLf + space$(sp): hlp$ = space$(sp): If w = -1 Then w = sp
    For i = 0 To count - 1
        acc$ = acc$ + hlp$ + nl$ + ToString2(i, sp - w * (sp > 0), w)
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
    Json = "[" + acc$ + vbCrLf + space(sp - w) + "]"
    End If
End Property
Property Get item(index As Long)
Attribute item.VB_UserMemId = 0
Select Case ValType(index)
Case 0
Case 4, 5
    Set item = ar(index)
Case 33
Dim s As String
    s = ar(index)
    If InStr(s, curdot$) > 0 Then
        s = Replace(ar(index), curdot$, ".")
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
    item = ar(index)
End Select
End Property
Property Let item(index As Long, RHS)
    Assign index, RHS
End Property
Property Set item(index As Long, RHS)
    Assign index, RHS
End Property

Property Get ItemPath(Path$, Optional sep As String = ".")
Dim part$(), w As Long, n As Long, ma As JsonArray, mo As JsonObject
part$ = Split2(Path$, sep)
Dim m As Object

n = 0
Set m = Me
again:
If n > UBound(part$) Then Exit Property
If TypeOf m Is JsonArray Then
    Set ma = m
    w = val(part$(n))
    Select Case ma.ValType(w)
    Case 0
    Case 4, 5
        Set m = ma(w)
        n = n + 1
        If n <= UBound(part$) Then GoTo again
        Set ItemPath = m
    Case Else
        ItemPath = ma(w)
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
Private Function Split2(a$, Optional sep As String = ".") As String()
Dim s() As String, i As Long, look As Boolean, many As Long, lastsep As Long
Dim strip As Boolean
' pass one
ReDim s(0)
If Len(a$) = 0 Then Split2 = s: Exit Function
many = 0
look = True
For i = 1 To Len(a$)
    Select Case Mid$(a$, i, 1)
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
For i = 1 To Len(a$)
    Select Case Mid$(a$, i, 1)
    Case "["
        If lastsep = i - 1 Then look = False: strip = True
    Case "]"
        look = True
    Case sep
        If look Then
            If strip Then
                s(many) = Trim$(Mid$(a$, lastsep + 2, i - lastsep - 3))
            Else
                s(many) = Trim$(Mid$(a$, lastsep + 1, i - lastsep - 1))
                If IsNumeric(s(many)) Then s(many) = " " + s(many)
            End If
            strip = False
            lastsep = i: many = many + 1
        End If
    End Select
Next i
If strip Then
    s(many) = Trim$(Mid$(a$, lastsep + 2, i - lastsep - 3))
Else
    s(many) = Trim$(Mid$(a$, lastsep + 1, i - lastsep - 1))
    If IsNumeric(s(many)) Then s(many) = " " + s(many)
End If
Split2 = s
End Function
Function Parser(a$) As Object
Dim i As Long, jumpstring As Long, ch As String, Level As Long, v
Dim b() As Byte
ReDim b(0)
Dim ma As JsonArray, mo As JsonObject
Dim markin As Long
Dim p()
ReDim p(100), mm(100)
Level = -1
For i = 1 To Len(a$)
ch = Mid$(a$, i, 1)
If ch = "{" Then
    Level = Level + 1
    If Level > UBound(p) Then ReDim Preserve p(Level * 2): ReDim Preserve mm(Level * 2)
    p(Level) = ""
    Set mm(Level) = New JsonObject
    Set mo = mm(Level)
ElseIf ch = "}" Then
    If ma Is mm(Level) Then Exit Function
    Level = Level - 1
    If Level < 0 Then Set Parser = mo: Exit Function
    If TypeOf mm(Level) Is JsonArray Then
    Set ma = mm(Level)
        ma.Assign 0 + p(Level), mm(Level + 1)
    Else
        Set mo = mm(Level)
        mo.Assign "" + p(Level), mm(Level + 1)
    End If
ElseIf ch = "[" Then
    Level = Level + 1
    If Level > UBound(p) Then ReDim Preserve p(Level * 2): ReDim Preserve mm(Level * 2)
    p(Level) = 0
    Set mm(Level) = New JsonArray
    Set ma = mm(Level)
ElseIf ch = "]" Then
    If mo Is mm(Level) Then Exit Function
    Level = Level - 1
    If Level < 0 Then Set Parser = ma:  Exit Function
    If TypeOf mm(Level) Is JsonArray Then
    Set ma = mm(Level)
        ma.Assign 0 + p(Level), mm(Level + 1)
    Else
        Set mo = mm(Level)
        mo.Assign "" + p(Level), mm(Level + 1)
    End If
ElseIf ch = """" Then
    markin = i + 1
    i = i + 1
    ch = Mid$(a$, i, 1)
    While i < Len(a$) And ch <> """"
        If ch = "\" Then i = i + 1
        i = i + 1
        ch = Mid$(a$, i, 1)
    Wend
    If ch = """" Then
        If ma Is mm(Level) Then
            ma.Assign 0 + p(Level), ma.EscapeStrToString(Mid$(a$, markin, i - markin))
        ElseIf p(Level) = "" Then
            p(Level) = mo.EscapeStrToString(Mid$(a$, markin, i - markin))
            If p(Level) = "" Then Exit Function
            While i < Len(a$) And ch <> ":"
                i = i + 1
                ch = Mid$(a$, i, 1)
            Wend
            If ch <> ":" Then Exit Function
        Else
            mo.Assign "" + p(Level), ma.EscapeStrToString(Mid$(a$, markin, i - markin))
        End If
    Else
        Exit Function
    End If
ElseIf ch = "," Then
    If ma Is mm(Level) Then p(Level) = p(Level) + 1 Else p(Level) = ""
ElseIf ch = "t" Then
    If Mid$(a$, i, 4) = "true" Then
        i = i + 3
        If ma Is mm(Level) Then
            ma.Assign 0 + p(Level), True
        Else
            mo.Assign "" + p(Level), True
        End If
    Else
        Exit Function
    End If
ElseIf ch = "f" Then
    If Mid$(a$, i, 5) = "false" Then
        i = i + 4
        If ma Is mm(Level) Then
            ma.Assign 0 + p(Level), False
        Else
            mo.Assign "" + p(Level), False
        End If
    Else
        Exit Function
    End If
ElseIf ch = "n" Then
    If Mid$(a$, i, 4) = "null" Then
        i = i + 3
        If ma Is mm(Level) Then
            ma.Assign 0 + p(Level), ma.NullValue
        Else
            mo.Assign "" + p(Level), mo.NullValue
        End If
    Else
        Exit Function
    End If
ElseIf ch Like "[+-0123456789.]" Then
    markin = i
    i = i + 1
    Do While i <= Len(a$)
        ch = Mid$(a$, i, 1)
        Select Case ch
        Case Is < "!", "]", "}", ","
        If Not IsNumeric(Mid$(a$, markin, i - markin)) Then Exit Function
        i = i - 1
        Exit Do
        End Select
        i = i + 1
    Loop
    If i > Len(a$) Then Exit Function
    ch$ = Mid$(a$, markin, i - markin + 1)
    If ma Is mm(Level) Then
        If Len(ch) > 28 Then
            If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            b() = ch
            v = b()
            ma.Assign 0 + p(Level), v
        ElseIf InStr(1, ch, "e", vbTextCompare) = 0 And ch > 14 Then
            If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            ma.Assign 0 + p(Level), CDec(ch)
        Else
            ma.Assign 0 + p(Level), val(ch)
        End If
    Else
        If Len(ch) > 28 Then
            If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            b() = ch
            v = b()
            mo.Assign "" + p(Level), v
        ElseIf InStr(1, ch, "e", vbTextCompare) = 0 And Len(ch) > 14 Then
            If InStr(ch$, ".") > 0 Then Mid$(ch$, InStr(ch$, ".")) = curdot$
            mo.Assign "" + p(Level), CDec(ch)
        Else
            mo.Assign "" + p(Level), val(ch)
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


