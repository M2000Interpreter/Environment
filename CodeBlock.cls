VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodeBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' each basetask link to a codeblock and use own cursor
Option Explicit
Private WithEvents LL As Lexar, myabort As Boolean
Attribute LL.VB_VarHelpID = -1
Public CodeDoc As String
Private Type ast
    codepos As Long  ' in document CodeDoc
    codetype As Byte
    codemore As Byte
    codehelp As Integer
    codevar As Long
End Type
Dim Code() As ast
Private mcode As Long
Private varcost() As Variant
Private vcost As Long
Private mb As Long
Private LastCode As Long, LastVar As Long
Const strline = ":==<=<>=>+=-=/=*=*()^++--*#&/!=~,[]{}%@\;.?<=>**|" + """"""
Private ms As mStiva2, docvar As Document
Private Identifiers As FastCollection
Private mycode$, SkipCode As Boolean, CurLine As Long
Private Function GetLang(a$) As Long
GetLang = Abs(AscW(a$) < 256)
End Function


Public Function ProperStatement1110(there As Long, ByVal Lang As Long) As String
With Code(there)
    If .codetype = 11 Then
    ProperStatement1110 = Identifiers.KeyToString3(.codevar)
    If Lang = Abs(AscW(ProperStatement1110) < 255) Then there = there + 1: Exit Function
    ProperStatement1110 = vbNullString
    ElseIf .codetype = 10 Then
    ProperStatement1110 = Identifiers.KeyToString3(.codevar)
    If Lang = Abs(AscW(ProperStatement1110) < 255) Then there = there + 1: Exit Function
    ProperStatement1110 = vbNullString
    End If
End With

End Function

Public Function Statement1110(there As Long, Lang As Long) As String
With Code(there)
    If .codetype = 10 Or .codetype > 160 Then Statement1110 = Identifiers.KeyToString3(.codevar): Lang = Abs(AscW(Statement1110) < 255): there = there + 1
End With

End Function
Function SkipId(there, a$) As Boolean
With Code(there)
    If .codetype > 170 Then If a$ = Identifiers.KeyToString3(.codevar) Then there = there + 1: SkipId = True
End With

End Function
Property Get Statement11(there As Long) As String
With Code(there)
    If .codetype > 170 Then Statement11 = Identifiers.KeyToString3(.codevar): there = there + 1
End With
End Property
Property Get Statement10(there As Long) As String
With Code(there)
    If .codetype > 160 Then Statement10 = Identifiers.KeyToString3(.codevar): there = there + 1
End With
End Property

Property Get Comma(there As Long) As Boolean
With Code(there)
    If .codetype = 8 Then If .codevar = -34 Then there = there + 1: Comma = True: Exit Function
   ' If .codetype = 14 Then If .codevar = -34 Then there = there + 1: Comma = True
End With
End Property
Property Get Semi(there As Long) As Boolean
With Code(there)
    If .codetype = 8 Then If .codevar = -42 Then there = there + 1: Semi = True
End With
End Property
Property Get EndLine(there As Long) As Boolean
With Code(there)
    If .codetype = 0 Then there = there + 1: EndLine = True
End With
End Property
Property Get oper(there As Long) As Boolean
With Code(there)
    If .codetype = 14 Then there = there + 1: oper = True
End With
End Property
Property Get VarNum(there As Long) As Variant
VarNum = varcost(Code(there).codevar)
there = there + 1
End Property
Function IsNum(there As Long) As Boolean
With Code(there)
    IsNum = .codetype = 100
End With
End Function
Function MaybeIsSymbolAny(there As Long, a$) As Boolean
Dim j As Long
With Code(there)
    If .codevar < -2 Then
        j = -.codevar - 2
            Select Case j
                Case Is > 200
                    ' nothing
                Case Is > 100
                    ' nothing
                Case Is < 1
                Case Else
                  '  If a$ = "<=>" Then
                   ' MaybeIsSymbolAny = True
                   ' Else
                    MaybeIsSymbolAny = InStr(a$, Mid$(strline, j, 1)) > 0
                   ' End If
                End Select
    End If
End With
End Function
Function FastSymbol(there As Long, a$) As Boolean
Dim j As Long
With Code(there)
    If .codetype = 14 Then
        j = -.codevar - 2
            Select Case j
                Case Is > 200
                    FastSymbol = Mid$(strline, j - 200, 3) = a$
                Case Is > 100
                    FastSymbol = Mid$(strline, j - 100, 2) = a$
                Case Is < 1
                ' nothing
                Case Else
                    FastSymbol = Mid$(strline, j, 1) = a$
                End Select
    ElseIf .codetype = 19 Then
        FastSymbol = a$ = ")"
    End If
End With
If FastSymbol Then there = there + 1
End Function
Friend Function IsStrExp(basestack As basetask, there As Long, a$, Optional check As Boolean = True) As Boolean
With Code(there)
    Select Case .codetype
    Case 15, 16
    a$ = varcost(Code(there).codevar)
    there = there + 1
    IsStrExp = True
    End Select
End With
End Function
Friend Function IsExpA(bstack As basetask, there As Long, rr As Variant, parenthesis As Long, Optional ByVal noand As Boolean = True, Optional Comp As Boolean = True, Optional ByPass As Boolean = False) As Boolean
'bstack As basetask, aa$, rr As Variant, parenthesis As Long, Optional ByVal noand As Boolean = True, Optional comp As Boolean = True, Optional ByPass As Boolean = False
Dim sg As Integer, oldthere As Long
Dim R As Variant, ac As Variant, po As Variant
oldthere = there

po = CDbl(1)
ac = 0#
Do
    If FastSymbol(there, "-") Then
        po = -po
    ElseIf Not FastSymbol(there, "+") Then
        Exit Do
    End If
Loop
If IsNum(there) Then
    R = varcost(Code(there).codevar)
    If Sgn(po) = -1 Then R = -R
    there = there + 1
    IsExpA = True
End If
If Not IsExpA Then
If Not there = oldthere Then MissNumExpr
there = oldthere
Else
R = rr
End If
End Function
Function aheadstatusFast(there As Long)
Dim st As Long
st = there
Select Case Code(there).codetype
Case 11
End Select

End Function
Friend Function logical(basestack As basetask, pc As Long, D As Variant, Optional par As Long = 0, Optional flatobject As Boolean = False, Optional ByPass As Boolean = False) As Boolean
Dim b$, s2$, S3$ ' , OSTAC$
Dim ah As String, s$
If ByPass Then
ah = aheadstatusFast(basestack.tmpstr)  ' wrong
 
Else
    ah = aheadstatusFast(pc)
End If
If InStr(ah, "l") = 0 Then
If InStr(ah, "N") > 0 Then

If Not IsNumber(basestack, pc, D, flatobject) Then
    Set basestack.lastobj = Nothing
Else
    logical = True
End If
Else
If par > 0 Then
If Not GetArr(basestack, pc, D, s2$, 0) Then
    Set basestack.lastobj = Nothing
Else
    If Len(s$) > 0 Then
    If AscW(s$) = 8 Then
    par = par - 1
    If ah = "S" Then Exit Function
    IsStr1 basestack, s$, b$
    s2$ = s$

   GoTo conthere
   End If
    End If
    logical = True
End If

If logical Then
par = par - 1
End If
Exit Function
Else
logical = False
End If
End If
s$ = NLtrim$(s$)
Exit Function
End If
On Error Resume Next

If Err.Number > 0 Then Exit Function
s2$ = s$

If Left$(ah, 1) <> "N" Then
 IsStrExp basestack, pc, b$
conthere:
logical = False
If Not mTextCompare Then
If FastSymbol(pc, "=") Then
    logical = False
    If IsStrExp(basestack, pc, s2$) Then
    logical = True
    D = b$ = s2$
    Exit Function
    Else
    If LastErNum = -2 Then logical = True
    Exit Function
    End If
ElseIf FastSymbol(pc, "<") Then
    logical = False
    If Left$(s$, 1) = "=" Then
        Mid$(s$, 1, 1) = " "
        If Left$(s$, 2) = " >" Then
            Mid$(s$, 2, 1) = " "
            If IsStrExp(basestack, pc, s2$) Then
                logical = True
                Select Case b$
                Case Is < s2$
                    D = -1
                Case Is = s2$
                    D = 0
                Case Else
                    D = 1
                End Select
                Exit Function
            Else
                If LastErNum = -2 Then logical = True
            Exit Function
            End If
        Else
            If IsStrExp(basestack, pc, s2$) Then
                logical = True
                D = b$ <= s2$
                Exit Function
            Else
                If LastErNum = -2 Then logical = True
                Exit Function
            End If
        End If
    ElseIf Left$(s$, 1) = ">" Then
        Mid$(s$, 1, 1) = " "
        If IsStrExp(basestack, pc, s2$) Then
            logical = True
            D = b$ <> s2$
            Exit Function
        Else
            If LastErNum = -2 Then logical = True
            Exit Function
        End If
    ElseIf IsStrExp(basestack, pc, s2$) Then
    logical = True
    D = b$ < s2$
    Exit Function
            Else
    If LastErNum = -2 Then logical = True
    Exit Function
    End If
ElseIf FastSymbol(pc, ">") Then
    logical = False
    If Left$(s$, 1) = "=" Then
        Mid$(s$, 1, 1) = " "
        If IsStrExp(basestack, pc, s2$) Then
            logical = True
            D = b$ >= s2$
            Exit Function
        Else
            If LastErNum = -2 Then logical = True
            Exit Function
        End If
    ElseIf IsStrExp(basestack, pc, s2$) Then
    logical = True
    D = b$ > s2$
    Exit Function
            Else
    If LastErNum = -2 Then logical = True
    Exit Function
    End If
ElseIf FastSymbol(pc, "~") Then
    logical = False
    If IsStrExp(basestack, pc, s2$) Then
    logical = True
    D = b$ Like s2$
    Exit Function
            Else
    If LastErNum = -2 Then logical = True
    Exit Function
    End If
End If
Else
If FastSymbol(pc, "=") Then
    logical = False
    If IsStrExp(basestack, pc, s2$) Then
    logical = True
   ' d = CompareStr2(b$, s2$) = 0
    Exit Function
    Else
    If LastErNum = -2 Then logical = True
    Exit Function
    End If

    ElseIf FastSymbol(pc, "<") Then
    logical = False
    If Left$(s$, 1) = "=" Then
        Mid$(s$, 1, 1) = " "
        If Left$(s$, 2) = " >" Then
            Mid$(s$, 2, 1) = " "
            If IsStrExp(basestack, pc, s2$) Then
                logical = True
                'd = CompareStr2(b$, s2$)
                Exit Function
            Else
                If LastErNum = -2 Then logical = True
                Exit Function
            End If
        ElseIf IsStrExp(basestack, pc, s2$) Then
            logical = True
          '  d = CompareStr2(b$, s2$) < 1
            Exit Function
        Else
            If LastErNum = -2 Then logical = True
            Exit Function
        End If
    ElseIf Left$(s$, 1) = ">" Then
        Mid$(s$, 1, 1) = " "
        If IsStrExp(basestack, pc, s2$) Then
            logical = True
           ' d = CompareStr2(b$, s2$) <> 0
            Exit Function
        Else
            If LastErNum = -2 Then logical = True
            Exit Function
        End If
    ElseIf IsStrExp(basestack, pc, s2$) Then
        logical = True
       ' d = CompareStr2(b$, s2$) = -1
        Exit Function
    Else
        If LastErNum = -2 Then logical = True
        Exit Function
    End If
ElseIf FastSymbol(pc, ">") Then
    logical = False
    If Left$(s$, 1) = "=" Then
        Mid$(s$, 1, 1) = " "
        If IsStrExp(basestack, pc, s2$) Then
            logical = True
          ' d = CompareStr2(b$, s2$) > -1
            Exit Function
        Else
            If LastErNum = -2 Then logical = True
            Exit Function
        End If
    ElseIf IsStrExp(basestack, pc, s2$) Then
        logical = True
       ' d = CompareStr2(b$, s2$) = 1
        Exit Function
    Else
        If LastErNum = -2 Then logical = True
    Exit Function
    End If
ElseIf FastSymbol(pc, "~") Then
    logical = False
    If IsStrExp(basestack, pc, s2$) Then
    logical = True
    D = b$ Like s2$
    Exit Function
            Else
    If LastErNum = -2 Then logical = True
    Exit Function
    End If
End If

End If
If LastErNum <> -2 Then s$ = s2$

Else

s$ = s2$
cont145:
If IsNumber(basestack, pc, D, flatobject) Then
logical = True
Else

End If
End If
End Function


Friend Function GetArr(bstack As basetask, pc As Long, p As Variant, s$, skip As Long) As Boolean
Dim pppp As mArray, x1 As Long
x1 = 1
Set pppp = New mArray: pppp.myarrbase = 0: pppp.PushDim (1): pppp.PushEnd: pppp.Arr = True
If skip = 1 Then GoTo firstexp Else If skip = 2 Then GoTo firststr
If IsExp(bstack, pc, p) Then
firstexp:
            x1 = x1 + 1
            pppp.SerialItem 0, x1, 10
            If bstack.lastobj Is Nothing Then
                pppp.item(x1 - 2) = p
            Else
                Set pppp.item(x1 - 2) = bstack.lastobj
                Set bstack.lastobj = Nothing
            End If
ElseIf IsStrExp(bstack, pc, s$) Then
firststr:
            x1 = x1 + 1
            pppp.SerialItem 0#, x1, 10
            If bstack.lastobj Is Nothing Then
                pppp.item(x1 - 2) = s$
            Else
                Set pppp.item(x1 - 2) = bstack.lastobj
                Set bstack.lastobj = Nothing
            End If
End If
Do While MaybeIsSymbolAny(pc, ",)")
If Code(pc).codetype = 8 Then pc = pc + 1
If MaybeIsSymbolAny(pc, ")") Then

pppp.SerialItem 0, x1 - 1, 10
Exit Do
ElseIf IsExp(bstack, pc, p) Then
x1 = x1 + 1
            pppp.SerialItem 0, x1, 10
            If bstack.lastobj Is Nothing Then
                pppp.item(x1 - 2) = p
            Else
                Set pppp.item(x1 - 2) = bstack.lastobj
                Set bstack.lastobj = Nothing
            End If
ElseIf IsStrExp(bstack, pc, s$) Then
x1 = x1 + 1
            pppp.SerialItem 0, x1, 10
            If bstack.lastobj Is Nothing Then
                pppp.item(x1 - 2) = s$
            Else
                Set pppp.item(x1 - 2) = bstack.lastobj
                Set bstack.lastobj = Nothing
            End If
End If
Loop
p = 0
Set bstack.lastobj = New mHandler
bstack.lastobj.t1 = 3
Set bstack.lastobj.objref = pppp
End Function
Friend Function IsExp(basestack As basetask, pc As Long, R As Variant, Optional ByVal noand1 As Boolean = True, Optional flatobject As Boolean = False, Optional Comp As Boolean = True) As Boolean
Dim par As Long, parin As Long
If LastErNum = -2 Then LastErNum = 0
Select Case Code(pc).codetype
Case 0, 8, 9, 141
Exit Function
End Select
parin = par
IsExp = IsExpA(basestack, pc, R, par, noand1, Comp)
again:
If par > 0 Then
If MaybeIsSymbolAny(pc, ",") Then
parin = par
IsExp = GetArr(basestack, pc, R, "", 1)
If LastErNum = -2 Then IsExp = False Else

If IsExp Then
par = par - 1: If par > 0 Then GoTo again
If basestack.lastobj Is Nothing Then
IsExp = IsExpA(basestack, pc, R, par, noand1, Comp, True)
End If
End If
End If
End If
If Not IsExp Then
If basestack.usestack Then
IsExp = IsNumber(basestack, pc, R)
End If
End If
If LastErNum = -2 Then IsExp = False
While par > 0
IsExp = FastSymbol(pc, ")")
par = par - 1
Wend
If flatobject Then Set basestack.lastobj = Nothing
End Function
Friend Function IsNumber(bstack As basetask, there As Long, R As Variant, Optional flatobject As Boolean = False) As Boolean
Dim sg As Integer, oldthere As Long
oldthere = there
sg = 1
Do While MaybeIsSymbolAny(there, "-+")
If FastSymbol(there, "-") Then sg = -sg
FastSymbol there, "+"
Loop
If IsNum(there) Then
    R = varcost(Code(there).codevar)
    If sg = -1 Then R = -R
    there = there + 1
    IsNumber = True
End If
If Not IsNumber Then
If Not there = oldthere Then MissNumExpr
there = oldthere
End If
'0, 9, 141




End Function


Property Get CodeStr() As String
    CodeStr = mycode$
End Property
Property Get top() As Long
top = LastCode
End Property
Public Function codeline(fromhere As Long)
Dim there As Long, NewLine As Long
For there = fromhere To LastCode
    If Code(there).codetype = 0 Then Exit For
Next there

NewLine = there + 1
there = Code(there).codepos - Code(fromhere).codepos + 1
If Code(fromhere).codepos > 0 Then
codeline = Mid$(mycode$, Code(fromhere).codepos, there)
Else
codeline = vbNullString
End If
fromhere = NewLine

End Function
Public Sub Construct(a$)
    Set ms = New mStiva2
    Set LL = New Lexar
    Set Identifiers = New FastCollection
    LL.Construct2 a$
    SwapStrings a$, mycode$
End Sub
Private Sub Class_Initialize()
LastCode = -1
LastVar = -1
mcode = 50
ReDim Code(mcode) As ast
vcost = 30
ReDim varcost(vcost) As Variant
End Sub
Private Sub GetHandler()
LastVar = LastVar + 1
If LastVar >= vcost Then
    vcost = LastVar + 300
    ReDim Preserve varcost(vcost) As Variant
End If
End Sub

Private Sub LL_abort(ok As Boolean)
ok = myabort
End Sub

Private Sub LL_addone(n As Long, t As Long)
If t = 3 Or SkipCode Then Exit Sub
Dim p As Long, v As String
LastCode = LastCode + 1
Code(LastCode).codevar = -1
If LastCode >= mcode Then
    mcode = LastCode + 500
    ReDim Preserve Code(mcode) As ast
End If
If t = 5 Then t = 14
If t = 14 Or t = 18 Then
If ms.LookTopVal = "&" Then
        ms.drop 1
        p = CLng(ms.PopVal)
        v = myUcase(CStr(ms.PopVal), True)
        If Not Identifiers.ExistKey(v) Then
            Identifiers.AddKey v
        End If
        Code(p).codevar = Identifiers.Index
End If

End If
Select Case t
Case 140
With Code(LastCode)
    .codepos = n + 1
    .codetype = t
    .codevar = -1
        ms.PushVal CVar(LastCode)
        ms.PushStr "}"
End With

Case 141
With Code(LastCode)
    .codepos = n
    .codetype = t
   If ms.LookTopVal = "}" Then
    ms.drop 1
    p = CLng(ms.PopVal)
    Code(p).codevar = LastCode + 1
    .codevar = p + 1
    End If
End With

Case Else
 If t = 9 And LastCode > 0 Then
   If Code(LastCode - 1).codetype = 9 Then
    LastCode = LastCode - 1
    Exit Sub
   End If
   
   End If

 Do While t = 0 And LastCode > 0
   If Code(LastCode - 1).codetype = 9 Or Code(LastCode - 1).codetype = 0 Then
    LastCode = LastCode - 1
    Else
again:
    If IsThen(CStr(ms.LookTopVal)) Then
    If ms.LookSecondVal + 1 < LastCode Then
        If Code(CLng(ms.LookSecondVal)).codemore > 0 Then
                ms.drop 1
                p = CLng(ms.PopVal)
                Code(p).codehelp = CInt(LastCode - p + 1)
                GoTo again
        End If
        
    Else
        Code(CLng(ms.LookSecondVal)).codemore = 0
    End If
   ElseIf iselse(CStr(ms.LookTopVal)) Then
           If ms.LookSecondVal + 1 < LastCode Then
        If Code(CLng(ms.LookSecondVal)).codemore > 0 Then
                    ms.drop 1
                    p = CLng(ms.PopVal)
                    Code(p).codehelp = CInt(LastCode - p + 1)
                    GoTo again
        End If
        
    Else
        Code(CLng(ms.LookSecondVal)).codemore = 0
    End If
   End If
    Exit Do
   End If
   Loop
If t = 0 Or t = 9 Then
If LastCode > 0 Then
If Code(LastCode - 1).codetype = 19 Then
If ms.LookTopVal = "&" Then
        ms.drop 1
        p = CLng(ms.PopVal)
        v = CStr(ms.PopVal)
        If Not Identifiers.ExistKey(v) Then
            Identifiers.AddKey v
        End If
        Code(p).codevar = Identifiers.Index
End If
End If
End If
End If
With Code(LastCode)
 
    
    .codepos = n
    .codetype = t
    .codevar = -1
  
End With
End Select
End Sub


Private Sub LL_addstr(v As String, UseCase As Boolean)
If SkipCode Then Exit Sub
Dim z As Long, s$, zz As Integer
With Code(LastCode)
Select Case .codetype
Case 15, 16
' do nothing
Case Else
If Len(v) = 0 Then
    .codetype = 0
    Exit Sub
End If
End Select

'If LastCode = 845 Then Stop

Select Case .codetype
Case 212
GetHandler
If UCase(Left(v, 2)) = "0X" Or UCase(Left(v, 2)) = "0Χ" Then
    Mid$(v, 1, 2) = "&H"
    If Right$(v, 1) = "%" Then
    varcost(LastVar) = CVar(CInt(val(v)))
    Else
    varcost(LastVar) = CVar(CLng(val(v)))
    End If
End If
.codevar = LastVar
.codetype = 100

Case 218
GetHandler
If UCase(Left(v, 2)) = "0X" Or UCase(Left(v, 2)) = "0Χ" Then
    v = Mid$(v, 3)
    varcost(LastVar) = HexToUnsigned(v)
Else
    varcost(LastVar) = CVar(-(CDbl(UNPACKLNG(Right$(v, 2)) * 65536#) + CDbl(UNPACKLNG(Mid$(v, 4, 2)) * 256#) + CDbl(UNPACKLNG(Mid$(v, 2, 2)))))
End If
.codevar = LastVar
.codetype = 100
Case 11, 10, 112, 111, Is > 160

If .codetype = 165 Then
If LastCode = 0 Then
    ms.PushStr UCase(v)
    ms.PushStr "#"
    ms.PushVal CVar(LastCode)
    ms.PushStr ")"
    Exit Sub
Else
zz = Code(LastCode - 1).codetype
If zz = 9 Or zz = 0 Then
    ms.PushStr UCase(v)
    ms.PushStr "#"
    ms.PushVal CVar(LastCode)
    ms.PushStr ")"
    Exit Sub
End If
End If
End If
If Right$(v, 1) = "(" Then
    ms.PushVal CVar(LastCode)
    ms.PushStr ")"
End If
If .codetype = 171 Then
    If UseCase Then v = myUcase(v, True)
    
    If IsIf(v) Then
        If isEnd(CStr(ms.LookTopVal)) Then
            If Not GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                SyntaxError
                myabort = True
                Exit Sub
            End If
            ms.drop 2
            If iselse(CStr(ms.LookTopVal)) Then
                If GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                    ms.drop 1
                    z = CLng(ms.PopVal)
                    Code(z).codehelp = CInt(LastCode - z + 1)
                End If
            ElseIf IsThen(CStr(ms.LookTopVal)) Then
                If GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                    ms.drop 1
                    z = CLng(ms.PopVal)
                    Code(z).codehelp = CInt(LastCode - z + 1)
                End If
            ElseIf IsElseIf(CStr(ms.LookTopVal)) Then
                If GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                    ms.drop 1
                    z = CLng(ms.PopVal)
                    Code(z).codehelp = CInt(LastCode - z + 1)
                End If
            End If
        Else
        ms.PushVal CVar(LastCode)
        ms.PushStr v
        End If
    ElseIf isEnd(v) Then
        ms.PushVal CVar(LastCode)
        ms.PushStr v
    ElseIf IsThen(v) Then
        If IsIf(CStr(ms.LookTopVal)) Then
            If GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                ms.drop 2
                ms.PushVal CVar(LastCode)
                ms.PushStr v
                .codetype = 10
                .codemore = 1
            Else
                GoTo syer
            End If
        ElseIf IsElseIf(ms.LookTopVal) Then
            If GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                ms.drop 1
                z = CLng(ms.PopVal)
                Code(z).codehelp = CInt(LastCode - z)   ' NOT +1
                ms.PushVal CVar(LastCode)
                ms.PushStr v
                .codetype = 10
                .codemore = 1
            Else
syer:
                SyntaxError
                myabort = True
            End If
        Else
            GoTo syer
        End If
    ElseIf IsElseIf(v) Then
        If IsThen(CStr(ms.LookTopVal)) Then
iselse:
            If GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                ms.drop 1
                z = CLng(ms.PopVal)
                Code(z).codehelp = CInt(LastCode - z)   ' NOT +1
                ms.PushVal CVar(LastCode)
                 ms.PushStr v
                 .codetype = 10
                 .codemore = 1
            Else
            SyntaxError
                myabort = True
            End If
        ElseIf iselse(CStr(ms.LookTopVal)) Then
        If Code(CLng(ms.LookSecondVal)).codemore <> 1 Then
        MyEr "No Else.If after last Else", "Όχι Αλλιώς.Αν μετά το τελευταίο Αλλιώς"
        myabort = True
        Else
        GoTo iselse
        End If
        End If
    ElseIf iselse(v) Then
        If IsIf(CStr(ms.LookTopVal)) Then
            If GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                ms.drop 1
                z = CLng(ms.PopVal)
                Code(z).codehelp = CInt(LastCode - z)   ' NOT +1
                ms.PushVal CVar(LastCode)
                ms.PushStr v
                .codetype = 10
                .codemore = 2
            Else
                GoTo syer
            End If
        ElseIf IsThen(CStr(ms.LookTopVal)) Then
            If GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                ms.drop 1
                z = CLng(ms.PopVal)
                Code(z).codehelp = CInt(LastCode - z + 1)
                ms.PushVal CVar(LastCode)
                ms.PushStr v
                .codetype = 10
                .codemore = 2
            Else
                GoTo syer
            End If
        ElseIf IsElseIf(CStr(ms.LookTopVal)) Then
            If GetLang(CStr(ms.LookTopVal)) = GetLang(v) Then
                ms.drop 1
                z = CLng(ms.PopVal)
                Code(z).codehelp = CInt(LastCode - z)   ' NOT +1
                ms.PushVal CVar(LastCode)
                ms.PushStr v
                .codetype = 10
                .codemore = 1
            Else
                GoTo syer
            End If
        Else
            MyEr "No If for ELSE", "Δεν υπάρχει Αν για το Αλλιώς"
            myabort = True
'            GoTo syer
        End If
    End If
Else
    If UseCase Then
        v = myUcase(v, True)
        If Code(LastCode).codetype = 171 Then
            If v = "REM" Or v = "ΣΗΜ" Then
                LastCode = LastCode - 1
                SkipCode = True
                Exit Sub
            End If
        End If
    End If
End If

If Not Identifiers.ExistKey(v) Then
    Identifiers.AddKey v
End If
.codevar = Identifiers.Index
Case 16
If .codevar = -1 Then
    GetHandler
    .codevar = LastVar
    varcost(LastVar) = CVar(v)
Else
    varcost(LastVar) = varcost(LastVar) + CVar(v)
End If

Case 24, 19, 18, 17, 14, 9, 8, 5
    If .codetype = 17 Then
    ms.PushVal CVar(LastCode)
    ms.PushStr ")"
    ElseIf .codetype = 19 Then
    If ms.count > 0 Then
        If ms.LookTopVal <> ")" Then
        MyEr "expected )", "περίμενα )"
        myabort = True
        Else
        ms.drop 1
        z = CLng(ms.PopVal)
        Code(z).codehelp = CInt(LastCode - z + 1)
        
        If ms.LookTopVal = "#" Then
        ms.drop 1
        ms.PushVal CVar(z)
        ms.PushStr "&"
        
        End If
        End If
        End If
    ElseIf .codetype = 5 Then
    .codetype = 14
    ElseIf .codetype = 8 Then
    If ms.LookTopVal = ")" Then
        z = ms.LookSecondVal
        If Code(z).codetype = 17 Then Code(z).codetype = 21   ' 21 for tuple
    End If
    ElseIf .codetype = 9 Then
    If .codevar <> -1 Then Exit Sub
    End If
    z = InStr(strline, v)
    If z > 0 Then
    Select Case Len(v)
    Case 1
        .codevar = -z - 2
    Case 2
        .codevar = -z - 102
    Case 3
        .codevar = -z - 202
    End Select

    Else
    '    Stop   ' DON'T FORGET TO DELETE THIS
     
    GetHandler
    .codevar = LastVar
    varcost(LastVar) = CVar(v)
    End If
Case 13
'Stop
' HTML COLOR

GetHandler

.codevar = LastVar
.codetype = 100
Case Else
    If Len(v) = 0 Then
    .codevar = -2
    Else
    
    GetHandler
    .codevar = LastVar
    varcost(LastVar) = CVar(v)
    End If
End Select

End With
End Sub

Private Sub LL_addval(v As Variant)
If SkipCode Then Exit Sub
With Code(LastCode)
Select Case .codetype
Case 100, 13
GetHandler
varcost(LastVar) = v
.codevar = LastVar
End Select
End With
End Sub
Private Sub Add(s As String)
If docvar.IsEmpty Then
docvar.textDoc = s
Else
docvar.InsertDoc docvar.LastParagraph, docvar.TextParagraphLen(docvar.LastParagraph) + 1, s
End If

End Sub
Function RetStr(X As Integer, base As Long) As String
If X > 0 Then RetStr = " skip to" + str$(X + base)
End Function
Sub ExportStr(b As Object)

Dim bstack As basetask
Set bstack = b
Dim i As Long, j As Long, numbers As Long, strconst As Long
Set docvar = New Document
If myabort Then
If LastCode > 0 Then
If pagio = "GREEK" Then
Add LastErNameGR & " στη θέση χαρακτήρα " & Code(LastCode).codepos
Else
Add LastErName & " at char pos " & Code(LastCode).codepos
End If
MyEr "", ""
RepPlain bstack, bstack.Owner, (docvar.textDoc)
End If
Exit Sub
End If
For i = 0 To LastCode
    With Code(i)
    
    Add str$(i) + str$(.codepos) + str$(.codetype)
    Select Case .codetype
    Case 9
        If .codevar < -1 Then
        Add " :" + vbCrLf
        Else
        Add " Identifier: " + Identifiers.KeyToString3(.codevar) + RetStr(.codehelp, i) & " id (" & .codevar & ")" & vbCrLf
        End If
    Case 21
        Add " TUPLE (" + RetStr(.codehelp, i) + vbCrLf
    Case 10, 11, 112, 111, Is > 160
        If Identifiers.count < .codevar Then
            Add " Syntax Error:" + vbCrLf
            Add " " + codeline(i) + vbCrLf
            GoTo ex12
        Else
            If .codevar = -1 Then
                Add " Syntax Error:" + vbCrLf
                Add " " + codeline(i) + vbCrLf
                GoTo ex12
            Else
                Add " Identifier: " + Identifiers.KeyToString3(.codevar) + RetStr(.codehelp, i) & " id (" & .codevar & ")" & vbCrLf
            End If
        End If
    Case 140
        Add " START BRACKET (end bracket: " & (.codevar) & ")" & vbCrLf
    Case 141
        Add " END BRACKET (start  bracket: " & (.codevar) & ")" & vbCrLf
    Case 20
    Add " Label " + varcost(.codevar) + vbCrLf
    Case Else
        Select Case .codevar
            Case Is < -2
                j = -.codevar - 2
                Select Case j
                Case Is > 200
                    Add " " + Mid$(strline, j - 200, 3) & " symbol id (" & j & ")" & vbCrLf
                Case Is > 100
                    Add " " + Mid$(strline, j - 100, 2) & " symbol id (" & j & ")" & vbCrLf
                    
                Case Else
                'If Mid$(strline, j, 1) = ")" And .codetype <> 19 Then Stop
                    Add " " + Mid$(strline, j, 1) + RetStr(.codehelp, i) & " symbol id (" & j & ")" & vbCrLf
                End Select
            Case -2
                Add " " + """""" + vbCrLf
            Case Is > -1

                    'add " " + Replace(Replace(CStr(varcost(.codevar)), vbCrLf, "\n"), vbTab, "\t") + vbCrLf
                If .codetype = 16 Then
                    Add " {" & (varcost(.codevar)) & "}" + RetStr(.codehelp, i) & " strliteral (" & .codevar & ")" & vbCrLf
                
                        strconst = strconst + 1
                ElseIf .codetype = 15 Then
                    Add " " + """" & (varcost(.codevar)) & """" & RetStr(.codehelp, i) & " strliteral (" & .codevar & ")" & vbCrLf
                        strconst = strconst + 1
                ElseIf .codetype = 14 Then
                    Add " >>>>>>>>>>>>>>>>>>>  " & (varcost(.codevar)) + RetStr(.codehelp, i) + vbCrLf
                        
                Else
                    On Error Resume Next
                    Add " " & (varcost(.codevar)) & " " & Typename$(varcost(.codevar)) & " numliteral (" & .codevar & ")" & vbCrLf
                    If Err.Number > 0 Then
                        Add " Syntax Error:" & vbCrLf
                        Add " " + codeline(i) + vbCrLf
                        GoTo ex12
                    End If
                    numbers = numbers + 1
               
                End If
            Case Else
                Add vbCrLf
        End Select
    End Select
    End With
Next i
Add "Number Literals:" + str$(numbers) + vbCrLf
Add "String Literals:" + str$(strconst) + vbCrLf
Add "Identifiers:" + str$(Identifiers.count) + vbCrLf
If LastVar - numbers - strconst + 1 > 0 Then
Add "Labels :" + str$(LastVar - numbers - strconst + 1) + vbCrLf
End If
ex12:
' unhide next two lines to get results to clipboard
If ms.Total > 0 Then
Set docvar = New Document
i = CLng(ms.LookSecondVal)
If pagio = "GREEK" Then
Add "Δεν έκλεισε σωστά στη θέση χαρακτήρα " & Code(i).codepos
Else
Add "Not Properly closed at char pos " & Code(i).codepos
End If
Else
Clipboard.Clear
SetTextData CF_UNICODETEXT, (docvar.textDoc)      'set as unicode text
End If
        Clipboard.Clear
       SetTextData CF_UNICODETEXT, docvar.textDoc
RepPlain bstack, bstack.Owner, (docvar.textDoc)

        'For i = 0 To LastVar
        'RepPlain bstack, bstack.Owner, (varcost(i))
        'Next i
Set docvar = Nothing
End Sub

Private Function IsIf(v As String) As Boolean
    If v = "IF" Or v = "AN" Then IsIf = True
End Function
Private Function iselse(v As String) As Boolean
    If v = "ELSE" Or v = "ΑΛΛΙΩΣ" Then iselse = True
End Function
Private Function IsElseIf(v As String) As Boolean
    If v = "ELSE.IF" Or v = "ΑΛΛΙΩΣ.ΑΝ" Then IsElseIf = True
End Function
Private Function IsThen(v As String) As Boolean
   If v = "THEN" Or v = "ΤΟΤΕ" Then IsThen = True
End Function
Function isEnd(a$) As Boolean
a$ = myUcase(a$, True)
Select Case a$
Case "END", "ΤΕΛΟΣ"
isEnd = True
End Select
End Function

Private Sub LL_NewLine(ByVal i As Long)
SkipCode = False
CurLine = i
End Sub
