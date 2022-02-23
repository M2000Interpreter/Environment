VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallBack2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'class for call back
Private objref As basetask
Private mineHere$
Public Event ErrorAtt(ByRef recover As Variant)
Private mObjRef As Object, mName$
Private reghash As New FastCollection
Dim locked As Boolean
Friend Sub NoPublic(bstack As basetask, where$)
        Set objref = New basetask
        Set objref.Sorosref = bstack.soros
        Set objref.Owner = bstack.Owner
        mineHere$ = where
End Sub
Friend Property Get CopyOfMe() As Object
Dim a As New CallBack2
a.NoPublic objref, mineHere$
Set CopyOfMe = a
End Property
Property Let ShowGui(RHS As Boolean)
If Forms.count = 0 Then Exit Property
If RHS Then
Form1.Visible = True
Form1.WindowState = 0
If Form1.top >= VirtualScreenHeight Then Form1.move ScrInfo(Console).Left, ScrInfo(Console).top
Else
If Form1.WindowState = 0 Then
If Not (QRY Or GFQRY Or Forms.count > 5) Then Form1.Visible = False
End If
End If
End Property
Public Sub EventRef(RHS As Object, Optional Name As String)
Set mObjRef = RHS
mName$ = Name$
End Sub
Public Sub MethodType(ByVal mName$, mType$)
mName$ = myUcase(mName$)
reghash.AddKey mName$, mType$
End Sub
Public Sub CallGlobal(modulename$, ParamArray a())
If status <> 0 Then Exit Sub
Dim ohere$
ohere$ = here$
Dim i As Long, ms As New mStiva, ar As mArray, j As Long
Dim manydim As Long, maxitems As Long, probe As Long, v() As Variant, probelow As Long
Dim oldsoros As mStiva
Set oldsoros = objref.soros
Set objref.Sorosref = ms
For i = 0 To UBound(a())
Select Case VarType(a(i))
Case Is >= vbArray
' only with error control
If Typename(a(i)) = "Variant()" Then
Set ar = New mArray
v() = a(i)
On Error Resume Next
For j = 1 To 60
    probe = UBound(v, j)
    If Err Then Exit For
Next j
manydim = j - 1
On Error Resume Next
Dim items As Long
For j = 1 To manydim
    
    probe = UBound(v, j)
    If Err Then Exit For
    probelow = LBound(v, j)
    ar.PushDim probe - probelow + 1
Next j
ar.PushEnd
'ar.RevOrder = True
ar.CopySerialize v()
ms.DataObj ar
Set ar = Nothing
Erase v()
Else
ms.DataValLong 0
End If
Case vbEmpty
ms.DataValLong 0
Case vbVariant, vbString
ms.DataStr CStr(a(i))
Case Else
ms.DataVal CDbl(a(i))
End Select
Next i
PushStage objref, 0
Set ms = Nothing
here$ = vbNullString
If subHash.ExistKey(myUcase(modulename$, True)) Then
    probe = interpret(objref, modulename$)
Else
   
    Dim what
        what = False
    If mObjRef Is Nothing Then
    RaiseEvent ErrorAtt(what)
        If what = True Then Reset
    Else
        On Error Resume Next
        If mName$ <> "" Then
            mObjRef.Run mName$ + "_ErrorAtt", what
            If Err Then Err.Clear: mObjRef.Run mName$ + ":ErrorAtt", what
            
        Else
            mObjRef.Run "ErrorAtt", what
        End If
        Err.Clear
        If what = True Then
            Reset
        Else
            Err.Raise 12345, , "Not Found"
        End If
    End If
    
End If
PopStage objref
Set objref.Sorosref = oldsoros
Set oldsoros = Nothing
here$ = ohere$
End Sub
Private Function GetSubFullName2(nm$, fullname$, retnum As Long) As Boolean
Dim dummy As Long
If Len(here$) > 0 Then
    If Len(here$) > Len(nm$) Then
        If Mid$(here$, Len(here$) - Len(nm$)) = "." + nm$ Then
            GetSubFullName2 = subHash.Find2(nm$, retnum, dummy)
          fullname = nm$: Exit Function
        End If
    End If
End If

If here$ <> "" Then
     GetSubFullName2 = subHash.Find2(here$ & "." & nm$, retnum, dummy)
  If GetSubFullName2 Then fullname$ = here$ & "." & nm$: Exit Function
End If

    GetSubFullName2 = subHash.Find2(nm$, retnum, dummy)
 If GetSubFullName2 Then fullname$ = nm$


End Function
Public Sub CallModule(modulename$, ParamArray a())
Static once As Boolean, where As Long, what, full$
If status <> 0 Then Exit Sub
If Not GetSubFullName2(myUcase(modulename$, True), full$, where) Then
    what = False
    
    If mObjRef Is Nothing Then
    RaiseEvent ErrorAtt(what)
       If what = True Then Reset
    Else
        On Error Resume Next
        If mName$ <> "" Then
            mObjRef.Run mName$ + "_ErrorAtt", what
            If Err Then Err.Clear: mObjRef.Run mName$ + ":ErrorAtt", what
            
        Else
            mObjRef.Run "ErrorAtt", what
        End If
        Err.Clear
        If what = True Then
            Reset
        Else
        MyEr "Not Found " + modulename$, "Δεν Βρέθηκε " + modulename$
        End If
    End If
Exit Sub
End If

If once Then Exit Sub
Dim ohere$
once = True
ohere$ = here$
here$ = full$


Dim i As Long, ms As New mStiva, ar As mArray, j As Long
Dim manydim As Long, maxitems As Long, probe As Long, v() As Variant, probelow As Long
Dim IhaveSignature As Boolean, sign$, sign2$, dummy$, look As Long
If reghash.Find(myUcase(modulename$, True)) Then
sign$ = reghash.Value
sign2$ = sign$
IhaveSignature = True
End If
' no byrefernece yet
Dim oldsoros As mStiva
Set oldsoros = objref.soros
Set objref.Sorosref = ms


For i = 0 To UBound(a())

Select Case VarType(a(i))
Case Is >= vbArray
' only with error control
If Typename(a(i)) = "Variant()" Then
Set ar = New mArray
v() = a(i)
On Error Resume Next
For j = 1 To 60
    probe = UBound(v, j)
    If Err Then Exit For
Next j
manydim = j - 1
On Error Resume Next
For j = manydim To 1 Step -1
    
    probe = UBound(v, j)
    If Err Then Exit For
    probelow = LBound(v, j)
    ar.PushDim probe - probelow + 1
Next j
ar.PushEnd
ar.RevOrder = True

ar.CopySerialize v()
''
If IhaveSignature Then
look = IsLabelOnly(sign$, dummy$)
If look = 2 Then ' byref
look = IsLabelOnly(sign$, dummy$)
look = globalvarGroup("b" & (var2used) & "." & dummy$, True) ' perhaps this is a fault, why is true??
ElseIf look > 4 Then
FastSymbol sign, ")", True
ms.DataObj ar
Else
SyntaxError
GoTo there
End If
Else
ms.DataObj ar
End If
Set ar = Nothing
Erase v()
Else
ms.DataValLong 0
End If


Case vbEmpty
ms.DataValLong 0
Case vbVariant, vbString
ms.DataStr CStr(a(i))
Case Else

ms.DataVal CDbl(a(i))

End Select
Next i
PushStage objref, 0
objref.soros.MergeTop ms

probe = FastCallModule(objref, where)

If probe <> 1 Then
    If mObjRef Is Nothing Then
    RaiseEvent ErrorAtt(what)
       If what = True Then Reset
    Else
        On Error Resume Next
        If mName$ <> "" Then
            mObjRef.Run mName$ + "_ErrorAtt", what
            If Err Then Err.Clear: mObjRef.Run mName$ + ":ErrorAtt", what
            
        Else
            mObjRef.Run "ErrorAtt", what
        End If
        Err.Clear
        If what = True Then
            Reset
        Else
            Err.Raise 12345, , "Not Found"
        End If
    End If

End If

   


there:
PopStage objref

Set objref.Sorosref = oldsoros
Set oldsoros = Nothing
here$ = ohere$
once = False
End Sub

Public Sub ExecuteStatement(ByVal code$)
Dim what$, Pos As Long, func As Boolean, param$, drop$
If status <> 0 Then Exit Sub
If IsLabelOnly(code$, what$) > 4 Then
    func = True
ElseIf what$ = vbNullString Then
    Exit Sub
ElseIf LTrim$(code$) = vbNullString Then
' one call
Call Run(what$)
Exit Sub
End If
Do
 Pos = 1
  drop$ = aheadstatus(code$, True, Pos)
  If drop$ = vbNullString And Pos > Len(code$) Then Exit Sub
  If param$ <> "" Then
  param$ = param$ + ", " + Left$(code$, Pos - 1)
  Else
  param$ = Left$(code$, Pos - 1)
  End If
  code$ = Mid$(code$, Pos)
  Loop Until Not FastSymbol(code$, ",")
  If func Then
  Call Run(what$ + " " + param$ + ")")
  Else
  Call Run(what$ + " " + param$)
  End If
End Sub
Public Function Eval(code$) As Variant
Dim Pos As Long, p As Double, s As String, drop$
 Pos = 1
  drop$ = aheadstatus(code$, True, Pos)
  If drop$ = vbNullString And Pos > Len(code$) Then Exit Function
  If InStr(drop$, "N") Then
  If IsExp(objref, Left$(code$, Pos), p) Then
  If Not objref.lastobj Is Nothing Then
  Set Eval = objref.lastobj
  
  Else
    Eval = p
   End If
  End If
  Else
  If IsStrExp(objref, Left$(code$, Pos), s$) Then
  If Not objref.lastobj Is Nothing Then
      Set Eval = objref.lastobj
  Else
    Eval = s$
  End If
  End If
  End If
Set objref.lastobj = Nothing
End Function
Public Sub AddExecCode(code$)
If locked Then Exit Sub
Call Run(code$, False)
End Sub
Public Function Run(ByVal code$, Optional clean As Boolean = True) As Boolean
Dim ohere$, stay As Variant
If status <> 0 Then Exit Function
If elevatestatus2 Then Exit Function
elevatestatus2 = 1
If clean Then PushStage objref, 0
ohere$ = here$
here$ = mineHere$
NoAction = True
NOEXECUTION = False
objref.toprinter = False
MOUT = False
ResetBreak
Run = True
        LastErNum = 0: LastErNum1 = 0
        LastErName = vbNullString
        LastErNameGR = vbNullString
If Execute(objref, code$, False) <> 1 Then
Run = False

If Not App.StartMode = vbSModeStandalone Then
                 MOUT = True
                 ProcTitle objref, Chr$(34) + Chr$(34) + ",0", 0
End If
            
            
ClearLoadedForms
If NERR Then
If mObjRef Is Nothing Then
    RaiseEvent ErrorAtt(stay)
        If stay = True Then Reset Else Err.Raise 12345, , "Not Found"
    Else
        On Error Resume Next
        If mName$ <> "" Then
            mObjRef.Run mName$ + "_ErrorAtt", stay
            If Err Then Err.Clear: mObjRef.Run mName$ + ":ErrorAtt", stay
            
        Else
            mObjRef.Run "ErrorAtt", stay
        End If
        Err.Clear
        If stay = True Then
            Reset
        Else
             Err.Raise 12345, , "Not Found"
        End If
    End If
If Not stay Then GoTo finalize
End If
    objref.toprinter = False
    If MOUT Then
            NOEXECUTION = False
            ResetBreak
            MOUT = interpret(objref, "new: flush: clear")
            MOUT = False
        Else
        
        If NOEXECUTION Then
                closeAll
        End If

        End If
        
MOUT = True
NoAction = False
If ExTarget Then Run = False
End If
finalize:
If NERR Then
' problem
PopStage objref

End If
here$ = ohere$

If clean Then PopStage objref
elevatestatus2 = 0
Dim exityes
If NERR Then NERR = False: Err.Raise 50500, , "Fatsl Error"
End Function
Property Get ErrorGr() As String
ErrorGr = LastErNameGR
End Property
Property Get ErrorEn() As String
ErrorEn = LastErName
End Property
Property Get status() As Long
If elevatestatus2 Then
status = -elevatestatus2
Else
status = LastErNum
End If
End Property
Public Sub Reset()

    NERR = False
        LastErNum = 0: LastErNum1 = 0
        LastErName = vbNullString
        LastErNameGR = vbNullString
 

        
End Sub

Private Sub Class_Initialize()
        Set objref = New basetask
        Set objref.Owner = Form1.DIS
        mineHere$ = vbNullString 'HERE$
End Sub
Public Function SpellUnicode(a$)
' use spellunicode to get numbers
' and make a ListenUnicode...with numbers for input text
Dim b$, i As Long
For i = 1 To Len(a$) - 1
b$ = b$ & CStr(AscW(Mid$(a$, i, 1))) & ","
Next i
SpellUnicode = b$ & CStr(AscW(Right$(a$, 1)))
SpellUnicode = b$ & CStr(AscW(Right$(a$, 1)))
End Function
Public Function ListenUnicode(ParamArray aa() As Variant) As String
Dim all$, i As Long
For i = 0 To UBound(aa)
    all$ = all$ & ChrW(aa(i))
Next i
ListenUnicode = all$
End Function
Private Sub Class_Terminate()
Set mObjRef = Nothing
mName$ = vbNullString
Set objref.Owner = Nothing
Set objref.Sorosref = Nothing
Set objref = Nothing
End Sub
' This is a way to create a fhandler outside M2000 and pass it to
' M2000 system, creating a M2000 File Handler.
' The M2000 File Handler now can be used from Print, Write, Input, Seek, Seek(), Get, Put, Close
' The exclusive or not is upon creation of the original file handler
' By Default we use open for Input (fway=1), so adjust it accordingly
' The fstep say 1 (means one character, this can change to something else only for Field type,fway=4)
' The unic=1 means we use UTF16LE (One Word width, 2 bytes)
' Use unic=0 for ANSI (is bound to Locale number)

Public Function FileHandler(Fhandler, Optional fway = 1, Optional Fstep = 1, Optional uinc = 1) As Variant
FileHandler = BigFileHandler(CLng(Fhandler), CLng(fway), CLng(Fstep), CLng(uinc))
End Function


