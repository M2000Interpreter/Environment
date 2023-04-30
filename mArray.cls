VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iBoxArray
Dim mmname As String
Dim dnum As Long, ndnum As Long
Dim limit(10) As Long  ' max 9 dimensions
Dim dtable(10) As Long
Dim nlimit(10) As Long  ' max 9 dimensions
Dim limBase(10) As Long
Dim nLimbase(10) As Long
Private Const mHdlr = "mHandler"
Private Const mgrp = "Group"
Private Const mlmbd = "lambda"
Private Const mArr = "mArray"
Private actualData
Dim maxitems As Long
Dim maxitems1 As Long
Dim nmaxitems As Long
Dim nmaxitems1 As Long
Public comevents As Boolean
Public myarrbase As Long
Public common As Boolean
Public IhaveClass As Boolean
Public IHaveGui As Boolean
' get a floating Group, maybe with groups inside....as floating group
' used also for other mHandler with t1 = 1 (Inventories)
Public GroupRef As Object
' only for check
'Private mGroupRef As Object
Public CodeName As String
Private oldmaxitems As Long
Public Arr As Boolean, Stable As Boolean
Public RevOrder As Boolean
Public MyTypeToBe As Integer
Private Type SAFEArrayPart
   cDims                        As Integer
   fFeatures                    As Integer
   cbElements                   As Long
   cLocks                       As Long
   pvData                       As Long
   End Type
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Integer)
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
        Private Declare Sub CopyPart Lib "kernel32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SafeArrayGetDim Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function SafeArrayLock Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function SafeArrayUnlock Lib "OleAut32.dll" (ByVal pSA As Long) As Long
  
Private Declare Function SafeArrayAccessData Lib "OleAut32.dll" (ByVal pSA As Long, ppVdata As Long) As Long
Private Declare Function SafeArrayUnaccessData Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function SafeArrayGetElemsize Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Sub SafeArrayCopyData Lib "OleAut32.dll" (ByVal psaSource As Any, ByVal psaTarget As Any)
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type
' no second sort allowed per array
Private sortmult As Long, sortoff As Long, upart As Long, part As Long, Col(), Piv, piv1()
Public Index As Long, Final As Boolean, refgroup As Object, grouperror As Boolean
Public refArray As Variant, StrArray As Boolean, IsByValue As Boolean
Public Sub Compute0(p As Variant, ss$)
If Final Then CantAssignValue: Exit Sub
Dim v As Long, sp As Variant, ii As Variant
Dim n As Byte
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
    GetMem1 VarPtr(actualData(v)), n
    If n = 20 Then
        CopyMemory ByVal VarPtr(ii) + 8, ByVal VarPtr(actualData(v)) + 8, 8
        PutMem2 VarPtr(ii), 20
        Select Case ss$
        Case "=", "g"
            item(v) = ii
            GoTo conthere
        Case "+"
            item(v) = ii + p
        Case "*"
            item(v) = p * ii
        Case "-"
            item(v) = ii - p
        Case "/"
            If p = 0 Then Exit Do
            item(v) = ii / p
        Case "~", "!!"
            item(v) = -1 - (ii <> 0)
            GoTo conthere
        Case "++"
            item(v) = 1 + ii
            GoTo conthere
        Case "--"
            item(v) = ii - 1
            GoTo conthere
        Case "-!"
            item(v) = -ii
            GoTo conthere
        Case "+="
            item(v) = ii + p
        Case "*="
            item(v) = p * ii
        Case "-="
            item(v) = ii - p
        Case "/="
            If p = 0 Then DevZero: Exit Do
            item(v) = ii / p
        Case "DIV", "ÄÉÁ"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            item(v) = Fix(ii / p)
        Case "DIV#", "ÄÉÁ#"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            If p < 0 Then
                item(v) = Int((ii - Abs(ii - Abs(p) * Int(ii / Abs(p)))) / p)
            Else
                item(v) = Int(ii / p)
            End If
        Case "MOD", "ÕÐÏË", "ÕÐÏËÏÉÐÏ"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            sp = ii - Fix(ii / p) * p
            If Abs(sp) >= Abs(p) Then sp = sp - sp
            actualData(v) = sp
        Case "MOD#", "ÕÐÏË#", "ÕÐÏËÏÉÐÏ#"
             If p = 0 Then
                DevZero
                Exit Do
            End If
            sp = Abs(ii - Abs(p) * Int(ii / Abs(p)))
            If Abs(sp) >= Abs(p) Then sp = sp - sp
            actualData(v) = sp
        Case Else
            WrongOperator
            Exit Do
        End Select
        ii = Empty
    ElseIf (n < 8) Or (n = 11) Or (n = 17) Or (n = 14) Then
        Select Case ss$
        Case "=", "g"
            actualData(v) = p
            GoTo conthere
        Case "+"
            actualData(v) = p + actualData(v)
        Case "*"
            actualData(v) = p * actualData(v)
        Case "-"
            actualData(v) = actualData(v) - p
        Case "/"
            If p = 0 Then Exit Do
            actualData(v) = actualData(v) / p
        Case "~", "!!"
            actualData(v) = -1 - (actualData(v) <> 0)
        Case "++"
            actualData(v) = 1 + actualData(v)
        Case "--"
            actualData(v) = actualData(v) - 1
        Case "-!"
            actualData(v) = -actualData(v)
            GoTo conthere
        Case "+="
            actualData(v) = actualData(v) + p
        Case "*="
            actualData(v) = p * actualData(v)
        Case "-="
            actualData(v) = actualData(v) - p
        Case "/="
            If p = 0 Then DevZero: Exit Do
            actualData(v) = actualData(v) / p
        Case "DIV", "ÄÉÁ"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            actualData(v) = Fix(actualData(v) / p)
        Case "DIV#", "ÄÉÁ#"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            If p < 0 Then
                actualData(v) = Int((actualData(v) - Abs(actualData(v) - Abs(p) * Int(actualData(v) / Abs(p)))) / p)
            Else
                actualData(v) = Int(actualData(v) / p)
            End If
        Case "MOD", "ÕÐÏË", "ÕÐÏËÏÉÐÏ"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            sp = actualData(v) - Fix(actualData(v) / p) * p
            If Abs(sp) >= Abs(p) Then sp = sp - sp
            actualData(v) = sp
        Case "MOD#", "ÕÐÏË#", "ÕÐÏËÏÉÐÏ#"
             If p = 0 Then
                DevZero
                Exit Do
            End If
            sp = Abs(actualData(v) - Abs(p) * Int(actualData(v) / Abs(p)))
            If Abs(sp) >= Abs(p) Then sp = sp - sp
            actualData(v) = sp
        Case Else
            WrongOperator
            Exit Do
        End Select
    End If
    If RoundDouble Then If n = vbDouble Then actualData(v) = MyRound(actualData(v), 13)
conthere:
    v = v + 1
    Err.Clear
Loop

End Sub
Public Sub Compute(p As Double, ss$)
If Final Then CantAssignValue: Exit Sub
Dim v As Long, ii As Variant
Dim n As Byte
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
    GetMem1 VarPtr(actualData(v)), n
    If n = 20 Then
        CopyMemory ByVal VarPtr(ii) + 8, ByVal VarPtr(actualData(v)) + 8, 8
        PutMem2 VarPtr(ii), 20
        Select Case ss$
        Case "=", "g"
            item(v) = ii
            GoTo conthere
        Case "+"
            item(v) = ii + p
        Case "*"
            item(v) = p * ii
        Case "-"
            item(v) = ii - p
        Case "/"
            If p = 0 Then Exit Do
            actualData(v) = actualData(v) / p
        Case "~", "!!"
            item(v) = -1 - (ii <> 0)
            GoTo conthere
        Case "++"
            item(v) = 1 + ii
            GoTo conthere
        Case "--"
            item(v) = ii - 1
            GoTo conthere
        Case "-!"
            item(v) = -ii
            GoTo conthere
        End Select

    ElseIf (n < 8) Or (n = 11) Or (n = 17) Or (n = 14) Then
        Select Case ss$
        Case "=", "g"
            actualData(v) = p
            GoTo conthere
        Case "+"
            actualData(v) = p + actualData(v)
        Case "*"
            actualData(v) = p * actualData(v)
        Case "-"
        actualData(v) = actualData(v) - p
        If RoundDouble Then If VarType(actualData(v)) = vbDouble Then actualData(v) = MyRound(actualData(v), 13)
        Case "/"
            If p = 0 Then Exit Do
            actualData(v) = actualData(v) / p
        Case "~", "!!"
            actualData(v) = -1 - (actualData(v) <> 0)
            GoTo conthere
        Case "++"
            actualData(v) = 1 + actualData(v)
            GoTo conthere
        Case "--"
            actualData(v) = actualData(v) - 1
            GoTo conthere
        Case "-!"
            actualData(v) = -actualData(v)
            GoTo conthere
        End Select
    End If
    If RoundDouble Then If VarType(actualData(v)) = vbDouble Then actualData(v) = MyRound(actualData(v), 13)
conthere:
    v = v + 1
    Err.Clear
Loop

End Sub
Public Function CompareItem(v As Long, p As Variant, ss$, ret As Boolean) As Boolean
ret = False
On Error GoTo there
Dim ii As Variant
Dim n As Byte
GetMem1 VarPtr(actualData(v)), n
If n = 20 Then
    CopyMemory ByVal VarPtr(ii) + 8, ByVal VarPtr(actualData(v)) + 8, 8
    PutMem2 VarPtr(ii), 20
    Select Case ss$
    Case ">"
    ret = ii > p
    Case ">="
    ret = ii >= p
    Case "<"
    ret = ii < p
    Case "<="
    ret = ii <= p
    Case "<>"
    ret = ii <> p
    Case "="
    ret = ii = p
    End Select
    ii = Empty
ElseIf (n < 8) Or (n = 11) Or (n = 17) Or (n = 14) Then
    Select Case ss$
    Case ">"
    ret = actualData(v) > p
    Case ">="
    ret = actualData(v) >= p
    Case "<"
    ret = actualData(v) < p
    Case "<="
    ret = actualData(v) <= p
    Case "<>"
    ret = actualData(v) <> p
    Case "="
    ret = actualData(v) = p
    End Select

Else
Select Case UCase(ss$)
Case ">"
ret = actualData(v) > p
Case ">="
ret = actualData(v) >= p
Case "<"
ret = actualData(v) < p
Case "<="
ret = actualData(v) <= p
Case "<>"
ret = actualData(v) <> p
Case "="
ret = actualData(v) = p
Case "LIKE"
ret = actualData(v) Like Replace$(CStr(p), "%", "*")

End Select
End If
CompareItem = True
there:
End Function
Public Sub Compute2(p As Variant, ss$)
If Final Then CantAssignValue: Exit Sub
Dim v As Long, sp As Variant, ii As Variant
Dim n As Byte
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
    GetMem1 VarPtr(actualData(v)), n
    If n = 20 Then
        CopyMemory ByVal VarPtr(ii) + 8, ByVal VarPtr(actualData(v)) + 8, 8
        PutMem2 VarPtr(ii), 20
        Select Case ss$
            Case "="
                item(v) = ii
                GoTo conthere
            Case "+="
                item(v) = ii + p
            Case "*="
                item(v) = p * ii
            Case "-="
                item(v) = ii - p
            Case "/="
                If p = 0 Then DevZero: Exit Do
                item(v) = ii / p
            Case "DIV", "ÄÉÁ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                item(v) = Fix(ii / p)
            Case "DIV#", "ÄÉÁ#"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                If p < 0 Then
                    item(v) = Int((ii - Abs(ii - Abs(p) * Int(ii / Abs(p)))) / p)
                Else
                    item(v) = Int(ii / p)
                End If
            Case "MOD", "ÕÐÏË", "ÕÐÏËÏÉÐÏ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = ii - Fix(ii / p) * p
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                item(v) = sp
             Case "MOD#", "ÕÐÏË#", "ÕÐÏËÏÉÐÏ#"
                 If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = Abs(ii - Abs(p) * Int(ii / Abs(p)))
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                item(v) = sp
            Case Else
                WrongOperator
                Exit Do
            End Select
            ii = Empty
    ElseIf (n < 8) Or (n = 11) Or (n = 17) Or (n = 14) Then
        Select Case ss$
            Case "="
                actualData(v) = p
                GoTo conthere
            Case "+="
                actualData(v) = actualData(v) + p
            Case "*="
                actualData(v) = p * actualData(v)
            Case "-="
                actualData(v) = actualData(v) - p
            Case "/="
                If p = 0 Then DevZero: Exit Do
                 actualData(v) = actualData(v) / p
            Case "DIV", "ÄÉÁ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                actualData(v) = Fix(actualData(v) / p)
            Case "DIV#", "ÄÉÁ#"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                If p < 0 Then
                    actualData(v) = Int((actualData(v) - Abs(actualData(v) - Abs(p) * Int(actualData(v) / Abs(p)))) / p)
                Else
                    actualData(v) = Int(actualData(v) / p)
                End If
            Case "MOD", "ÕÐÏË", "ÕÐÏËÏÉÐÏ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = actualData(v) - Fix(actualData(v) / p) * p
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                actualData(v) = sp
             Case "MOD#", "ÕÐÏË#", "ÕÐÏËÏÉÐÏ#"
                 If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = Abs(actualData(v) - Abs(p) * Int(actualData(v) / Abs(p)))
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                actualData(v) = sp
            Case Else
                WrongOperator
                Exit Do
            End Select
        End If
    If RoundDouble Then If n = vbDouble Then actualData(v) = MyRound(actualData(v), 13)
conthere:
    v = v + 1
    Err.Clear
Loop
End Sub
Public Sub Compute3(ss$)
Dim v As Long, ii As Variant
Dim n As Byte
    If Final Then CantAssignValue: Exit Sub
    On Error Resume Next
    v = 0
    Do While v <= Me.UpperMonoLimit
        GetMem1 VarPtr(actualData(v)), n
        If n = 20 Then
            CopyMemory ByVal VarPtr(ii) + 8, ByVal VarPtr(actualData(v)) + 8, 8
            PutMem2 VarPtr(ii), 20
            Select Case ss$
            Case "~"
                ii = -1 - (ii <> 0)
            Case "++"
                ii = 1 + ii
            Case "--"
                ii = ii - 1
            Case "-!"
                ii = -ii
            End Select
            item(v) = ii
            ii = Empty
        ElseIf (n < 8) Or (n = 11) Or (n = 17) Or (n = 14) Then
            Select Case ss$
            Case "~"
                actualData(v) = -1 - (actualData(v) <> 0)
            Case "++"
                actualData(v) = 1 + actualData(v)
            Case "--"
                actualData(v) = actualData(v) - 1
            Case "-!"
                actualData(v) = -actualData(v)
            End Select
        End If
    v = v + 1
    Err.Clear
Loop
End Sub
Public Function PlaceValue2UDT(ByVal p As Long, Name$, v) As Boolean
On Error Resume Next
Module10.PlaceValue2UDT actualData(p), Name$, v
If Err Then
    MyEr Err.Description, Err.Description
    Err.Clear
    Exit Function
End If
PlaceValue2UDT = True
End Function
Public Function GetUDTValue(ByVal p As Long, Name$, v) As Boolean
On Error Resume Next
GetUDTValue = Module10.GetUDTValue(actualData(p), Name$, v)
If Err Then
    MyEr Err.Description, Err.Description
    Err.Clear
End If
End Function
Property Get Value() As Variant
If Index >= 0 And Index <= maxitems1 Then

If IsObject(actualData(Index)) Then
    If Typename(actualData(Index)) = "PropReference" Then
        Value = actualData(Index)
    Else
        Set Value = actualData(Index)
    End If
Else
            If myVarType(actualData, vbObject) Then
               
                    Value = actualData(Index)
                ElseIf VarType(actualData(Index)) = 20 Then
                
                    CopyMemory ByVal VarPtr(Value) + 8, ByVal VarPtr(actualData(Index)) + 8, 8
                    PutMem2 VarPtr(Value), 20
                Else
                    Value = actualData(Index)
                 
                End If

End If
End If
End Property
Public Sub PushProp(A As Variant)
Set GroupRef = A
Arr = False
End Sub
Friend Sub ExportArrayNow()
If myVarType(actualData, vbObject) Then
' ????
Else
CopyVariantArray refArray, actualData
IsByValue = False
End If
End Sub
Friend Function ExportArrayCopy() As Variant
If myVarType(actualData, vbObject) Then
    Set ExportArrayCopy = actualData.Copy
Else
    ExportArrayCopy = actualData
End If
End Function

Friend Sub ExportStringArrayNow()
StrArray = True
End Sub

Friend Function ExportStrArray() As String()
If myVarType(actualData, vbObject) Then Exit Function
If VarType(actualData) <> vbArray + vbVariant Then Exit Function
Dim i As Long
Dim n As Byte, s$, m_base As Long, r1 As Long, where As Long

ReDim strAr(maxitems) As String
For i = 0 To maxitems - 1
GetMem1 VarPtr(actualData(i)), n
If n = 8 Then
strAr(i) = actualData(i)
Else
End If
Next i
ExportStrArray = strAr
End Function
Private Function makeonedim(items As Long) As Variant()
ReDim makeonedim(0 To items - 1)
End Function

Friend Sub FeedFromSrings(Carrier As Variant)
Dim A() As String, i As Long
If myVarType(actualData, vbObject) Then Exit Sub
If VarType(actualData) <> vbArray + vbVariant Then Exit Sub
A() = Carrier
Carrier = Array()
Dim zz As Variant, rCount As Long, dimActual As Long, dummy As Long, tst As Long
Dim ii As Long
CopyMemory ByVal VarPtr(refArray), zz, 16
On Error GoTo TT
rCount = UBound(A()) - LBound(A) + 1
dimActual = 1
TT:
If dimActual = 0 Then
If dnum = 0 Then Exit Sub
Class_Terminate
Class_Initialize
Exit Sub
End If
StartResize
If dnum >= dimActual Then
PushDim rCount, limBase(dimActual - 1)
Else
PushDim rCount
End If
Dim sum As Long
sum = rCount
again:
dummy = dimActual
On Error GoTo ok1
tst = UBound(A(), dimActual + 1)
rCount = UBound(A(), dimActual + 1) - LBound(A, dimActual + 1) + 1
dimActual = dimActual + 1
sum = sum * rCount
ok1:
If dimActual <> dummy Then
If dnum >= dimActual Then
PushDim rCount, limBase(dimActual - 1)
Else
PushDim rCount
End If
GoTo again
End If
If dimActual <> 1 Then
MyEr "wrong array dimensions", "ëÜèïò äéáóôÜóåéò ðßíáêá"
Else
actualData = makeonedim(sum)
ii = LBound(A)
For i = LBound(actualData, 1) To UBound(actualData, 1)
actualData(i) = A(ii)
ii = ii + 1
Next i
PushEnd
End If
End Sub
Friend Sub FixArray()
Dim zz As Variant, rCount As Long, dimActual As Long, dummy As Long, tst As Long
If myVarType(actualData, vbObject) Then Exit Sub
If VarType(actualData) <> vbArray + vbVariant Then Exit Sub
Dim hlp
CopyMemory ByVal VarPtr(refArray), zz, 16
If IsByValue Then Exit Sub
If StrArray Then

Else
On Error GoTo TT
rCount = Elements(actualData)   'UBound(actualdata, 1) - LBound(actualdata, 1) + 1
dimActual = 1
TT:
If dimActual = 0 Then
If dnum = 0 Then Exit Sub
Class_Terminate
Class_Initialize
Exit Sub
End If
StartResize
If dnum >= dimActual Then
PushDim rCount, limBase(dimActual - 1)
Else
PushDim rCount
End If
Dim sum As Long
sum = rCount
again:
dummy = dimActual
On Error GoTo ok1
tst = UBound(actualData, dimActual + 1)
rCount = UBound(actualData, dimActual + 1) - LBound(actualData, dimActual + 1) + 1
dimActual = dimActual + 1
sum = sum * rCount
ok1:
If dimActual <> dummy Then
If dnum >= dimActual Then
PushDim rCount, limBase(dimActual - 1)
Else
PushDim rCount
End If
GoTo again
End If
If oldmaxitems <> sum Then
hlp = actualData
actualData = makeonedim(sum)
CopySerialize hlp
End If

PushEnd
'Dim z As Long, i As Long
'If count <> rCount Then
'i = count
'maxitems1 = UBound(actualdata()) - LBound(actualdata)
'maxitems = maxitems1 + 1
'z = 1
'If dnum > 1 Then

  'z = i \ limit(dnum - 1)

'   If maxitems Mod z <> 0 Then
'  maxitems = maxitems + z - maxitems Mod z
'  maxitems1 = maxitems - 1
'  ReDim Preserve actualdata(maxitems1)
'  End If
'  limit(dnum - 1) = maxitems \ z
'Else
'If maxitems > 0 Then dnum = 1
' dtable(0) = 1
'limit(0) = maxitems
'End If
End If
'

End Sub
Friend Sub LoadTuple(ar As Variant)
Dim z As Long, i As Long
actualData = ar
If count <> UBound(actualData) - LBound(actualData) + 1 Then
i = count
maxitems1 = UBound(actualData) - LBound(actualData)
maxitems = maxitems1 + 1
z = 1

If dnum > 1 Then

  z = i \ limit(dnum - 1)

   If maxitems Mod z <> 0 Then
  maxitems = maxitems + z - maxitems Mod z
  maxitems1 = maxitems - 1
  ReDim Preserve actualData(maxitems1)
  End If
  limit(dnum - 1) = maxitems \ z

Else
dnum = 1
dtable(0) = 1
limit(0) = maxitems
End If
End If
'
End Sub
Friend Sub CopySerialize(mmm As Variant)
' may we get an exception here
If VarType(actualData) = vbArray + vbVariant Then
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, r1 As Long, r2 As Long
GetMem4 UnsignedAdd(VarPtr(actualData), 8), m_base_0
If m_base_0 <> 0 Then
GetMem4 LibMemory.ArrPtr(mmm), m_base
'GetMem4 VarPtrArray(actualdata()), m_base_0
r1 = SafeArrayLock(m_base)
r2 = SafeArrayLock(m_base_0)
r1 = SafeArrayAccessData(ByVal m_base, first12)
r2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
r2 = SafeArrayUnaccessData(ByVal m_base_0)
r1 = SafeArrayUnaccessData(ByVal m_base)
r2 = SafeArrayUnlock(m_base_0)
r1 = SafeArrayUnlock(m_base)
End If
End If
End Sub
Friend Sub CopyBackSerialize(mmm As Variant)
' may we get an exception here
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, r1 As Long, r2 As Long
If Not RevOrder Then Err.Raise vbObjectError + 12345, "CopyBack", "Failed. is not vb6"
Dim backup
backup = actualData
GetMem4 LibMemory.ArrPtr(mmm), m_base
GetMem4 LibMemory.ArrPtr(backup), m_base_0
r1 = SafeArrayLock(m_base)
r2 = SafeArrayLock(m_base_0)
r1 = SafeArrayAccessData(ByVal m_base, first12)
r2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
r2 = SafeArrayUnaccessData(ByVal m_base_0)
r1 = SafeArrayUnaccessData(ByVal m_base)
r2 = SafeArrayUnlock(m_base_0)
r1 = SafeArrayUnlock(m_base)
End Sub
Public Sub Car(m As Object)
Dim k As New mArray, usehandler As mHandler, s$
Dim i As Long, obj As Object, obj2 As Object, spare As lambda
If count > 0 Then
    k.PushDim 1: k.PushEnd: k.SerialItem actualData(0), 0, 1
    If ItemIsObject(0) Then
        s$ = ItemType(0)
        Select Case s$
        Case mgrp
            Set obj = item(0)
            CopyGroupObj obj, obj2
            Set k.item(0) = obj2
            Set obj2 = Nothing
        Case mlmbd
            Dim actualdata1()
            With item(0)
            .CopyTo spare, actualdata1()
            End With
            Set k.item(0) = spare
            Set spare = Nothing
        Case Else
            If s$ = mHdlr Then
                Set usehandler = item(0)
                If usehandler.t1 = 3 Then
                    Set m = usehandler
                    Set k = Nothing
                    Exit Sub
                End If
             ElseIf s$ = mArr Then
                Set usehandler = New mHandler
                usehandler.t1 = 3
                Set usehandler.objref = item(0)
                Set m = usehandler
                Set k = Nothing
                Exit Sub
            End If
            Set k.item(0) = item(0)
        End Select
    Else
        k.item(0) = item(0)
    End If
End If
k.Arr = True
Set m = k
Set k = Nothing
End Sub
Public Sub Cdr(m As Object)
Dim k As New mArray
Dim i As Long, obj As Object, obj2 As Object, spare As lambda
If dnum > 0 Then
    k.StartResize: k.PushDim Me.count - 1: k.PushEnd
Else
    k.Arr = True
    Set k.GroupRef = Me.GroupRef
    k.IHaveGui = Me.IHaveGui
    k.IhaveClass = Me.IhaveClass
    Set m = k
    Exit Sub
End If
k.myarrbase = myarrbase
k.RevOrder = False
For i = 1 To Me.UpperMonoLimit
    If IsObject(Me.item(i)) Then
        Select Case Typename(Me.item(i))
        Case mgrp
            Set obj = Me.item(i)
            CopyGroupObj obj, obj2
            Set k.item(i - 1) = obj2
            Set obj2 = Nothing
        Case mlmbd
            Dim actualdata1()
            With Me.item(i)
                .CopyTo spare, actualdata1()
            End With
            Set k.item(i - 1) = spare
            Set spare = Nothing
        Case Else
            Set k.item(i - 1) = Me.item(i)
        End Select
    Else
        k.item(i - 1) = Me.item(i)
    End If
Next i
' why not in CAR ???
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IhaveClass = Me.IhaveClass
Set m = k
End Sub
Public Sub processAppend(m As Object)
Dim pppp As New mArray, pppp1 As mArray
Set pppp1 = m
If dnum > 0 Then
    Me.CopyArray pppp
    If Not pppp1.IsEmpty Then pppp1.AppendArray pppp
ElseIf Not pppp1.IsEmpty Then
    Set pppp = pppp1
End If
pppp.Arr = True
Set m = pppp
Set pppp = Nothing
Set pppp1 = Nothing
End Sub

Public Sub AppendArray(k As mArray)
'append this to k
If k.count = 0 Then
    Me.CopyArray k
    Exit Sub
End If
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
Dim fromhere As Long
fromhere = k.UpperMonoLimit + 1
k.SerialItem v, Me.count + k.count, 10
k.myarrbase = myarrbase
k.RevOrder = RevOrder
For i = 0 To Me.UpperMonoLimit
    If IsObject(Me.item(i)) Then
        Select Case Typename(Me.item(i))
        Case mgrp
            Set obj = Me.item(i)
            CopyGroupObj obj, obj2
            Set k.item(i + fromhere) = obj2
            Set obj2 = Nothing
        Case Else
            Set k.item(i + fromhere) = Me.item(i)
        End Select
    Else
        k.item(i + fromhere) = Me.item(i)
    End If
Next i
'' new
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IhaveClass = Me.IhaveClass
End Sub



Public Sub CopyArray(k As mArray)
'copy this to k

If k Is Me Then Exit Sub
If k Is Nothing Then Set k = New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
k.MyTypeToBe = MyTypeToBe
' no use, we do the redim in PushEnd
'k.SerialItem v, Me.UpperMonoLimit, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For i = 0 To dnum - 1
    k.PushDim limit(i), limBase(i)
Next i

If myVarType(actualData, vbObject) And MyTypeToBe <> vbVariant Then
    k.PushEnd CVar(actualData.Copy)
    k.RevOrder = RevOrder
    Set k.GroupRef = Me.GroupRef
Else
    k.PushEnd
    k.RevOrder = RevOrder
    Set k.GroupRef = Me.GroupRef
    For i = 0 To Me.UpperMonoLimit
        If myVarType(Me.item(i), vbObject) Then
            Select Case Typename(Me.item(i))
            Case mgrp
                Set obj = Me.item(i)
                CopyGroupObj obj, obj2
                Set k.item(i) = obj2
                Set obj2 = Nothing
            Case mlmbd
                Dim actualdata1()
                With Me.item(i)
                .CopyTo spare, actualdata1()
                End With
                Set k.item(i) = spare
                Set spare = Nothing
            Case "RefArray"
                Set k.item(i) = Me.item(i).Copy()
            Case Else
                Set k.item(i) = Me.item(i)
            End Select
        Else
            k.item(i) = Me.item(i)
        End If
    Next i
'' new
End If
k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IhaveClass = Me.IhaveClass
k.Final = Me.Final
End Sub
Public Sub CopyArrayRev(k As mArray)
'copy this to k
If k Is Me Then Exit Sub
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
Dim mm As Long, M1 As Long
k.MyTypeToBe = MyTypeToBe
'k.SerialItem v, Me.UpperMonoLimit, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For i = 0 To dnum - 1
    k.PushDim limit(i), limBase(i)
Next i
k.PushEnd
k.RevOrder = RevOrder
Set k.GroupRef = Me.GroupRef
mm = Me.UpperMonoLimit
If myVarType(actualData, vbObject) And MyTypeToBe <> vbVariant Then
    For i = 0 To Me.UpperMonoLimit
        M1 = mm - i
        k.item(i) = Me.item(M1)
    Next i
Else
For i = 0 To Me.UpperMonoLimit
    M1 = mm - i
    If myVarType(Me.item(M1), vbObject) Then
        Select Case Typename(Me.item(M1))
        Case mgrp
            Set obj = Me.item(i)
            CopyGroupObj obj, obj2
            Set k.item(i) = obj2
            Set obj2 = Nothing
        Case mlmbd
            Dim actualdata1()
            With Me.item(M1)
                .CopyTo spare, actualdata1()
            End With
            Set k.item(i) = spare
            Set spare = Nothing
        Case "RefArray"
            Set k.item(i) = Me.item(i).Copy()
        Case Else
            Set k.item(i) = Me.item(M1)
        End Select
    Else
        k.item(i) = Me.item(M1)
    End If
Next i
End If
'' new
k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IhaveClass = Me.IhaveClass
k.Final = Me.Final
End Sub
Public Sub CopyArrayRevFast(k As mArray)
'copy this to k
If k Is Me Then Exit Sub
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
Dim mm As Long, M1 As Long
k.MyTypeToBe = MyTypeToBe
'k.SerialItem v, Me.UpperMonoLimit, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For i = 0 To dnum - 1
    k.PushDim limit(i), limBase(i)
Next i
k.PushEnd
k.RevOrder = RevOrder
Set k.GroupRef = Me.GroupRef
mm = Me.UpperMonoLimit
If myVarType(actualData, vbObject) And MyTypeToBe <> vbVariant Then
For i = 0 To Me.UpperMonoLimit
    M1 = mm - i
    k.item(i) = actualData(M1)
Next i
Else
For i = 0 To Me.UpperMonoLimit
    M1 = mm - i
    If myVarType(Me.item(M1), vbObject) Then
        Set k.item(i) = Me.item(M1)
    Else
        k.item(i) = Me.item(M1)
    End If
Next i
End If
'' new
k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IhaveClass = Me.IhaveClass
k.Final = Me.Final
End Sub
Public Sub CopyArraySlice(k As mArray, fromS As Long, toS As Long)
'copy this to k
' make one dimemsion only
If k Is Me Then Exit Sub
If k Is Nothing Then Set k = New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
k.MyTypeToBe = MyTypeToBe
'k.SerialItem v, toS - fromS + 1, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For i = 0 To 0
    k.PushDim toS - fromS + 1, 0
Next i
k.PushEnd
k.RevOrder = RevOrder
Set k.GroupRef = Me.GroupRef
For i = fromS To toS
    If IsObject(Me.item(i)) Then
        Select Case Typename(Me.item(i))
        Case mgrp
            Set obj = Me.item(i)
            CopyGroupObj obj, obj2
            Set k.item(i - fromS) = obj2
            Set obj2 = Nothing
        Case mlmbd
            Dim actualdata1()
            With Me.item(i)
            .CopyTo spare, actualdata1()
            End With
            Set k.item(i - fromS) = spare
            Set spare = Nothing
        Case "RefArray"
            Set k.item(i - fromS) = Me.item(i).Copy()
        Case Else
            Set k.item(i - fromS) = Me.item(i)
        End Select
    Else
        k.item(i - fromS) = Me.item(i)
    End If
Next i
'' new
k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IhaveClass = Me.IhaveClass
k.Final = Me.Final
End Sub
Public Sub CopyArraySliceFast(k As mArray, fromS As Long, toS As Long)
'copy this to k
' make one dimemsion only
If k Is Me Then Exit Sub
If k Is Nothing Then Set k = New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
k.MyTypeToBe = MyTypeToBe
'k.SerialItem v, toS - fromS + 1, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For i = 0 To 0
    k.PushDim toS - fromS + 1, 0
Next i
k.PushEnd
k.RevOrder = RevOrder
Set k.GroupRef = Me.GroupRef
For i = fromS To toS
If IsObject(Me.item(i)) Then
    Set k.item(i - fromS) = Me.item(i)
Else
    k.item(i - fromS) = Me.item(i)
End If
Next i

k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IhaveClass = Me.IhaveClass
k.Final = Me.Final
End Sub

Public Function UpperMonoLimit()
UpperMonoLimit = maxitems1
End Function
Public Property Get count()
    If dnum = 0 Then count = 0 Else count = maxitems1 + 1 ' sa
End Property
Public Property Let arrname(aName As String)
    mmname = aName
End Property
Public Property Get arrname() As String
    arrname = mmname
End Property
Public Sub PushDim(nDim As Long, Optional limBase As Variant)
If ndnum < 10 Then
    nlimit(ndnum) = nDim
    If IsMissing(limBase) Then
        nLimbase(ndnum) = myarrbase
    Else
        nLimbase(ndnum) = CLng(limBase)
    End If
    nmaxitems = nmaxitems * nDim
    ndnum = ndnum + 1
End If
nmaxitems1 = nmaxitems - 1
End Sub
Public Sub PushEnd(Optional thisarr)
Dim z As Variant
CopyMemory ByVal VarPtr(refArray), z, 16
On Error GoTo there11
If ndnum = 0 Then
    dnum = 0
    maxitems = 1
    maxitems1 = 0
    If myVarType(actualData, vbObject) Then
        actualData.RedimForFlat maxitems1
    Else
        ReDim actualData(maxitems1)
    End If
Else
    Dim i As Long
    For i = 0 To 9
        limit(i) = nlimit(i): limBase(i) = nLimbase(i)
    Next i

    dnum = ndnum
    Dim isnew As Boolean
    isnew = maxitems = 1
    maxitems = nmaxitems
    maxitems1 = nmaxitems1
    If nmaxitems1 < 0 Then
        If isnew And Not MyTypeToBe = vbVariant Then
            Dim ar As refArray
            Set ar = New refArray
            ar.ResetToType MyTypeToBe, 1
            Set actualData = ar
        ElseIf myVarType(actualData, vbObject) Then
            Set ar = actualData
            ar.RedimForFlat 1
        Else
            ReDim actualData(1)
        End If
        
    Else
        If Not IsMissing(thisarr) Then
                Set actualData = thisarr
        ElseIf isnew And Not MyTypeToBe = vbVariant Then
          '  Dim aR As refArray
            Set ar = New refArray
            ar.ResetToType MyTypeToBe, maxitems1
            Set actualData = ar
        Else
        If myVarType(actualData, vbObject) Then
            Set ar = actualData
            ar.RedimForFlat maxitems1
        Else
 
                ReDim Preserve actualData(maxitems1)
       
        End If
        End If
    End If
    Dim mX As Long
    mX = maxitems
    i = 0
    If limit(i) > 0 Then
        If RevOrder Then
            Dim revlimit(), revlimbase()
            ReDim revlimit(UBound(limit())), revlimbase(UBound(limit()))
            For i = 0 To dnum - 1
                revlimit(i) = limit(dnum - i - 1)
                revlimbase(i) = limBase(dnum - i - 1)
            Next i
            For i = 0 To dnum - 1
                limit(i) = revlimit(i)
                limBase(i) = revlimbase(i)
            Next i
        End If
        For i = 0 To dnum - 1
            dtable(i) = mX \ limit(i)
            mX = dtable(i)
        Next i
    Else
        GoTo there11
    End If
End If
Exit Sub
there11:
dnum = 0
maxitems = 1
maxitems1 = 0
    If myVarType(actualData, vbObject) Then
        Set ar = actualData
        ar.RedimForFlat maxitems
    Else
        ReDim actualData(maxitems)
    End If
End Sub
Public Sub StartResize()
oldmaxitems = maxitems1
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
End Sub
Friend Sub GetDnum(cursor As Long, item, lim)
If dnum = 0 Then item = 0: lim = 0: Exit Sub
If RevOrder Then
item = limit(dnum - cursor - 1)
lim = limBase(dnum - cursor - 1)
Else
item = limit(cursor)
lim = limBase(cursor)
End If
End Sub
Friend Function bDnum() As Variant
bDnum = dnum
End Function

Public Function SerialItem(item As Variant, cursor As Long, Command As Long) As Boolean
Dim aa As Variant
If Command = 1 Then
If dnum = 0 Then Exit Function
If cursor < 0 Or cursor >= maxitems Then Exit Function
If VarType(item) = vbObject Then
Set actualData(cursor) = item
Else
actualData(cursor) = item
End If
ElseIf Command = 2 Then
If dnum = 0 Then Exit Function
If cursor < 0 Or cursor >= maxitems Then Exit Function
item = actualData(cursor)
ElseIf Command = 3 Then
    Fill item
ElseIf Command = 4 Then
    cursor = maxitems
ElseIf Command = 5 Then
cursor = dnum
item = dnum
ElseIf Command = 6 Then
If cursor > 9 Or cursor < 0 Then
item = 0
Else
If RevOrder Then
item = limit(dnum - cursor - 1)
Else
item = limit(cursor)
End If
End If
ElseIf Command = 7 Then  'erase all data
Fill aa
ElseIf Command = 8 Then
If dnum = 1 Then
If cursor = -1 Then
dnum = 0 ' no actions
maxitems = 1
maxitems1 = 0
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
CopyMemory ByVal VarPtr(refArray), aa, 16
If myVarType(actualData, vbObject) Then
    actualData.RedimForFlat 0
Else
    ReDim actualData(0)
End If
CodeName = GetAnyName$
Arr = True
limit(0) = 0
Else
    If myVarType(actualData, vbObject) Then
        actualData.RedimForFlat cursor
    Else
        ReDim Preserve actualData(cursor) 'As Variant
    End If
    maxitems = cursor
End If
    limit(0) = maxitems

End If
ElseIf Command = 9 Then
If dnum = 1 Then
    ReDim Preserve actualData(cursor) 'As Variant
    maxitems = cursor - 1
    limit(0) = maxitems
End If
maxitems1 = maxitems - 1
ElseIf Command = 10 Then
   
    ReDim Preserve actualData(cursor) ' As Variant
    maxitems = cursor
    limit(0) = maxitems
dnum = 1
maxitems1 = maxitems - 1
ElseIf Command = 11 Then   ' new from 9.4  rev 4
If cursor > 9 Or cursor < 0 Then
item = 0
Else
If RevOrder Then
item = limBase(dnum - cursor - 1)
Else
item = limBase(cursor)
End If
End If
ElseIf Command = 12 Then   ' new from 9.4  rev 4
    For Command = 0 To 9
        limBase(Command) = cursor
    Next Command
    myarrbase = cursor
End If
SerialItem = True

End Function

Public Function PushOffset(curitem As Long, curdim As Long, nDim As Long) As Boolean
If curdim >= dnum Then
' error...
PushOffset = False
Else


    If RevOrder Then
         Dim pr As Long
        pr = dnum - curdim - 1
    If nDim >= limit(pr) - limBase(pr) Then Exit Function
    curitem = curitem + dtable(pr) * (nDim + limBase(pr))
    Else
    If nDim >= limit(curdim) - limBase(curdim) Then Exit Function
    curitem = curitem + dtable(curdim) * (nDim + limBase(curdim))
    End If

    
PushOffset = True
End If
End Function
Public Property Get itemPtr(curitem As Long) As Long
itemPtr = VarPtr(actualData(curitem))
End Property
Friend Sub CopyVariantArray(ByRef A As Variant, ByRef b As Variant)
    CopyMemory ByVal VarPtr(A), ByVal VarPtr(b), 16
End Sub
Friend Sub SwapItem(curitem As Long, thisvariant)
Static p
If myVarType(actualData, vbObject) Then
p = actualData(curitem)
actualData(curitem) = thisvariant
thisvariant = p
Else
SwapVariant actualData(curitem), thisvariant
End If
End Sub
Public Function IsStringItem(curitem As Long) As Variant
Dim usehandler As mHandler, n As Byte
If Arr Then
If myVarType(actualData, vbObject) Then
IsStringItem = VarType(actualData(curitem)) = vbString

Else
GetMem1 VarPtr(actualData(curitem)), n
If n = 9 Then
    If Not actualData(curitem) Is Nothing Then
        If TypeOf actualData(curitem) Is mHandler Then
            Set usehandler = actualData(curitem)
            If usehandler.t1 = 4 Then
                If myVarType(usehandler.index_cursor, vbString) Then
                    IsStringItem = True
                    Exit Function
                End If
            End If
        End If
    End If
End If
IsStringItem = n = 8
End If

ElseIf curitem < 0 And curitem <> -2 Then
On Error Resume Next
Set usehandler = GroupRef
If Err.Number Then InternalError: Exit Function
On Error GoTo 0
If usehandler.t1 <> 1 Then Err.Raise 1010, , "Need an Inventory Object"
If usehandler.objref.Done Then
    If usehandler.objref.IsObj Then
        
    Else
        IsStringItem = usehandler.objref.IsString
    End If
End If
End If
End Function
Public Property Get ItemIsObject(curitem) As Boolean
On Error GoTo there1
If Arr Then
    If maxitems1 >= curitem Then
        If MyIsObject(actualData(curitem)) Then
            ItemIsObject = Not actualData(curitem) Is Nothing
        End If
    End If
Else
    ItemIsObject = Not GroupRef Is Nothing
End If
there1:

End Property

Public Property Get ItemType(curitem) As String
On Error GoTo there1
If Arr Then
    If maxitems1 >= curitem Then
        ItemType = VarTypeName(actualData(curitem))
    End If
Else
If curitem = -2 Then
ItemType = Typename(GroupRef)
ElseIf curitem < 0 Then
If Not TypeOf GroupRef Is mHandler Then GoTo e1245
Dim usehandler As mHandler
Set usehandler = GroupRef
If usehandler.t1 <> 1 Then
e1245:
   ItemType = Typename(GroupRef)
End If
    If usehandler.objref.Done Then
        If usehandler.objref.IsObj Then
            ItemType = Typename(usehandler.objref.ValueObj)
        Else
            Dim p
            p = usehandler.objref.Value
            ItemType = VarTypeName(p)
        End If
    Else
        ItemType = Typename(usehandler)
    End If

End If

End If
there1:

End Property
Public Property Get ItemTypeNum(curitem) As String
On Error GoTo there1
If Arr Then
    If maxitems1 >= curitem Then
        ItemTypeNum = MemInt(VarPtr(actualData(curitem)))
    End If
Else
If curitem = -2 Then
ItemTypeNum = 9
ElseIf curitem < 0 Then
If Not TypeOf GroupRef Is mHandler Then GoTo e1245
Dim usehandler As mHandler
Set usehandler = GroupRef
If usehandler.t1 <> 1 Then
e1245:
   ItemTypeNum = 9
End If
    If usehandler.objref.Done Then
        If usehandler.objref.IsObj Then
            ItemTypeNum = 9
        Else
            Dim p
            p = usehandler.objref.Value
            ItemTypeNum = MemInt(VarPtr(p))
        End If
    Else
        ItemTypeNum = 9
    End If

End If

End If
there1:

End Property
Friend Sub GetPorp(aProp As PropReference, curitem As Long)
If Arr Then
    Set aProp = actualData(curitem)
Else
    Set aProp = GroupRef
End If
End Sub
Public Property Get item(curitem As Long) As Variant
On Error GoTo there1
If Arr Then
    If maxitems1 >= curitem Then
        If ItemIsObject(curitem) Then
            If Typename(actualData(curitem)) = "PropReference" Then
                item = actualData(curitem).Value
            Else
                Set item = actualData(curitem)
            End If
        Else
                If myVarType(actualData, vbObject) Then
                    item = actualData(curitem)
                ElseIf VarType(actualData(curitem)) = 20 Then
                
                    CopyMemory ByVal VarPtr(item) + 8, ByVal VarPtr(actualData(curitem)) + 8, 8
                    PutMem2 VarPtr(item), 20
                Else
                    item = actualData(curitem)
                End If
        End If
        Exit Property
    End If
Else
If curitem = -2 Then
Set item = GroupRef
ElseIf curitem < 0 Then
If Not TypeOf GroupRef Is mHandler Then GoTo e1245
Dim usehandler As mHandler
Set usehandler = GroupRef
If usehandler.t1 <> 1 Then
e1245:
    Err.Raise 5010, , "Need an Inventory Object"
End If
    If usehandler.objref.Done Then
        If usehandler.objref.IsObj Then
            Set item = usehandler.objref.ValueObj
        Else
            item = usehandler.objref.Value
        End If
    Else
        Set item = usehandler
    End If

End If
'End If
End If
there1:

End Property
Friend Property Get itemObject(curitem As Long) As Variant
On Error GoTo there1
If Arr Then
    Set itemObject = actualData(curitem)
Else
    If curitem = -2 Then
        Set itemObject = GroupRef
    ElseIf curitem < 0 Then
        If Not TypeOf GroupRef Is mHandler Then GoTo e1245
        Dim usehandler As mHandler
        Set usehandler = GroupRef
        If usehandler.t1 <> 1 Then
e1245:
            Err.Raise 5010, , "Need an Inventory Object"
        End If
        If usehandler.objref.Done Then
            If usehandler.objref.IsObj Then
                Set itemObject = usehandler.objref.ValueObj
            End If
        Else
            Set itemObject = usehandler
        End If
    End If
End If
there1:

End Property

Public Property Get itemnumeric(curitem As Long) As Variant
On Error GoTo there1
Dim usehandler As mHandler
If Arr Then
    If maxitems1 >= curitem Then
        If IsObject(actualData(curitem)) Then
            If Typename(actualData(curitem)) = "PropReference" Then
                itemnumeric = actualData(curitem).Value
            ElseIf Typename(actualData(curitem)) = mHdlr Then
                Set usehandler = actualData(curitem)
                If usehandler.t1 = 4 Then
                If myVarType(usehandler.index_cursor, vbString) Then
                    itemnumeric = usehandler.index_cursor
                Else
                    If usehandler.sign = -1 Then
                        itemnumeric = CDbl(-usehandler.index_cursor)
                    Else
                        itemnumeric = CDbl(usehandler.index_cursor)
                    End If
                    End If
                Else
                    itemnumeric = 0
                End If
            Else
                itemnumeric = 0
            End If
        Else
                If myVarType(actualData, vbObject) Then
                    itemnumeric = actualData(curitem)
                ElseIf VarType(actualData(curitem)) = 20 Then
                
                    CopyMemory ByVal VarPtr(itemnumeric) + 8, ByVal VarPtr(actualData(curitem)) + 8, 8
                    PutMem2 VarPtr(itemnumeric), 20
                Else
                    itemnumeric = actualData(curitem)
                End If
                
            'End If
        End If
        Exit Property
    End If
Else
    If curitem = -2 Then
    
    ElseIf curitem < 0 Then
        Set usehandler = GroupRef
        If usehandler.objref.Done Then
            If usehandler.objref.IsObj Then
                itemnumeric = 0
            Else
                itemnumeric = usehandler.objref.Value
            End If
        End If
    End If
End If
there1:

End Property

Public Property Let item(curitem As Long, item As Variant)
Dim usehandler As mHandler, ar As refArray
On Error Resume Next
If curitem < 0 Then
If curitem <> -2 Then
Set usehandler = GroupRef
If usehandler.objref.Done Then
    If MyIsObject(item) Then
    
    Else
        usehandler.objref.Value = item
    End If
    usehandler.objref.Done = False
End If

End If
ElseIf Final Then
If Not MyIsObject(actualData(curitem)) Then
If VarType(actualData(curitem)) = vbEmpty Then
actualData(curitem) = item
Else
CantAssignValue
End If
Else
CantAssignValue
End If
ElseIf myVarType(actualData, vbObject) Then
Set ar = actualData
'If myVarType(item, vbObject) Then
    ar(curitem) = item
    Select Case ar.AssignError
    Case 0
    Case 6
        OverflowValue ar.vtType(curitem)
    Case 5020
        missNumber
    Case 5030
        MissString
    Case 5040
        MyEr "Missing date or string as date", "Ëåßðåé çìåñïìçíßá Þ áëöáñéèìçôéêü ùò çìåñïìçíßá"
    Case Else ' 91
        MyEr "no obect to assign", "Äåí âñÞêá áíôéêåßìåíï íá êáôá÷ùñÞóù"
    End Select
ElseIf Not MyIsObject(actualData(curitem)) Then
If myVarType(item, 20) Then
actualData(curitem) = Empty
SwapVariant actualData(curitem), item
Else
actualData(curitem) = item
End If
Else
If Not MyIsObject(item) Then
    If TypeOf actualData(curitem) Is PropReference Then
        actualData(curitem).Value = item
    ElseIf TypeOf actualData(curitem) Is VarItem Then
        actualData(curitem) = item
    ElseIf TypeOf actualData(curitem) Is mHandler Then
        Set usehandler = actualData(curitem)
        If usehandler.t1 = 4 Then
        ' just throw the enum variable and place the numeber.
            actualData(curitem) = item
        End If
    Else
        Set actualData(curitem) = Nothing
        actualData(curitem) = 0&
    End If
Else
    actualData(curitem).Value = item
End If
End If

End Property
Public Property Let ItemStr(curitem As Long, item As String)
Dim usehandler As mHandler
On Error Resume Next
If curitem < 0 Then
If Not Arr Then
If curitem <> -2 Then
Set usehandler = GroupRef
If usehandler.t1 <> 1 Then Err.Raise 1010, , "Need an Inventory Object"
If usehandler.objref.Done Then
    If usehandler.objref.IsString Then
        usehandler.objref.Value = item
    End If
End If
End If
End If
ElseIf Final Then
If Not MyIsObject(actualData(curitem)) Then
If VarType(actualData(curitem)) = vbEmpty Then
MoveStringToVariant item, actualData(curitem)
Else
CantAssignValue
End If
Else
CantAssignValue
End If
ElseIf Not IsObject(actualData(curitem)) Then  ' why not myisobject ???
Select Case VarType(actualData(curitem))
Case vbVariant
MoveStringToVariant item, actualData(curitem)
Case Else
actualData(curitem) = item

End Select

Else
If TypeOf actualData(curitem) Is PropReference Then
    actualData(curitem).ValueStr = item
Else
    Set actualData(curitem) = Nothing
    actualData(curitem) = 0&
End If
End If

End Property
Public Property Set item(curitem As Long, item As Variant)
If Final Then CantAssignValue: Exit Property
On Error Resume Next
If curitem < 0 Then
If Not GroupRef Is Nothing Then
    If TypeOf GroupRef Is mHandler Then
        Dim usehandler As mHandler
        Set usehandler = GroupRef
        If usehandler.t1 = 1 Then
            usehandler.objref.Index = Abs(curitem + 100)
            usehandler.objref.Done = True
            Set usehandler.objref.ValueObj = item
            Exit Property
           
        End If
    ElseIf curitem = -2 Then
    If Typename$(item) = Typename$(GroupRef) Then
    Set GroupRef = item
    
    Exit Property
    End If
    End If
End If
MyEr "Wrong setting", "ËÜèïò üñéóìá"
ElseIf curitem <= maxitems1 Then
If myVarType(actualData, vbObject) Then
actualData(curitem) = item
Else
Set actualData(curitem) = item
End If
Else

If grouperror Then
MyEr "Something is wrong with this group", "ÊÜôé ðÜåé óôñáâÜ ìå áõôÞ ôçí ïìÜäá"
Else
MyEr "Something is wrong with this array", "ÊÜôé ðÜåé óôñáâÜ ìå áõôüí ôïí ðßíáêá"
End If
End If
End Property

Private Sub Fill(item As Variant)
Dim aa As Object
If dnum > 0 Then
Dim i As Long
If IhaveClass Then
For i = 0 To maxitems - 1
CopyGroupObj GroupRef, aa, True
aa.ToDelete = False
Set actualData(i) = aa
Next i
Else
If CheckInt64(item) Then
For i = 0 To maxitems - 1
actualData(i) = item
Next i
ElseIf VarTypeName(item) = "Empty" Then
Dim z As Variant
CopyMemory ByVal VarPtr(refArray), z, 16
If myVarType(actualData, vbObject) Then
    actualData.RedimForFlat maxitems
Else
    ReDim actualData(maxitems)
End If


''maxitems = 0

Else
For i = 0 To maxitems - 1
actualData(i) = item
Next i
End If
End If
End If
End Sub
Public Sub Fillobj(bstack As basetask)
Dim obj As Variant, i As Long
Set obj = bstack.lastobj
    Set bstack.lastobj = Nothing
For i = 0 To maxitems - 1
Set actualData(i) = obj
Next i
Set obj = Nothing
End Sub
Public Sub FillLambda(bstack As basetask)
Dim obj As Variant, i As Long
Set obj = bstack.lastobj
    Set bstack.lastobj = Nothing
For i = 0 To maxitems - 1
 CopyLambda obj, bstack
Set actualData(i) = bstack.lastobj
Set bstack.lastobj = Nothing
Next i
Set obj = Nothing
End Sub
Sub MakeitObject2(var As Variant)
Dim aa As Object
Set aa = New Group
Set var = aa
End Sub
Sub CopyGroupObj(myobject1 As Object, myobject2 As Object, Optional LinkToMe As Boolean = False)
Dim s$, vvl As Variant, x1 As Long
Dim subgroup As Object, sub2 As Object, myArray As mArray, mySecondArray As mArray
Dim mygroup2 As New Group
Dim myobject As Group
Set myobject = myobject1
If myobject.IamApointer Then
    Set myobject2 = myobject
    Exit Sub
End If
With myobject
    If .IamFloatGroup Then
        .PeekItem 1, vvl
        mygroup2.BeginFloat vvl + 2
        For x1 = 0 To vvl * 2 + 2
            .PeekItem x1, vvl
            s$ = VarTypeName(vvl)
            If s$ = mgrp Then
                Set subgroup = vvl
                CopyGroupObj subgroup, sub2
                Set subgroup = Nothing
                Set vvl = sub2
                Set sub2 = Nothing
            ElseIf s$ = mArr Then
                If Not vvl.common Then
                    Set myArray = vvl
                    Set mySecondArray = New mArray
                    myArray.CopyArray mySecondArray
                    Set myArray = Nothing
                    Set vvl = mySecondArray
                    Set mySecondArray = Nothing
                End If
            End If
            mygroup2.PokeItem x1, vvl
        Next x1
        mygroup2.HasStrValue = myobject.HasStrValue
        mygroup2.HasValue = myobject.HasValue
        mygroup2.HasSet = myobject.HasSet
        mygroup2.HasParameters = myobject.HasParameters
        mygroup2.HasParametersSet = myobject.HasParametersSet
        Set mygroup2.SuperClassList = myobject.SuperClassList
        mygroup2.HasRemove = myobject.HasRemove
        'If Not GroupRef Is Nothing Then
        If LinkToMe Then
            Set mygroup2.LinkRef = myobject
            mygroup2.link.ToDelete = True
        Else
            .PeekItem x1, vvl
            mygroup2.PokeItem x1, vvl
        End If
        Set mygroup2.Events = .Events
        mygroup2.highpriorityoper = .highpriorityoper
        mygroup2.HasUnary = .HasUnary
        mygroup2.ToDelete = False
        Set mygroup2.mytypes = .mytypes
    End If
End With
Set myobject2 = mygroup2
End Sub
Property Get IsEmpty() As Variant
IsEmpty = dnum = 0
End Property
Function IsObj() As Boolean
If Index >= 0 And Index <= maxitems1 Then
If IsObject(actualData(Index)) Then
    IsObj = True
End If
End If
End Function
Function IsObjAt(curitem As Long, peekvalue) As Boolean
On Error GoTo there1
peekvalue = Empty
If Arr Then
    If maxitems1 >= curitem Then
        If ItemIsObject(curitem) Then
            If Typename(actualData(curitem)) = "PropReference" Then
                peekvalue = actualData(curitem).Value
            Else
                Set peekvalue = actualData(curitem)
                IsObjAt = True
            End If
        Else
                If myVarType(actualData, vbObject) Then
                    peekvalue = actualData(curitem)
                ElseIf VarType(actualData(curitem)) = 20 Then
                
                    CopyMemory ByVal VarPtr(peekvalue) + 8, ByVal VarPtr(actualData(curitem)) + 8, 8
                    PutMem2 VarPtr(peekvalue), 20
                Else
                    peekvalue = actualData(curitem)
                End If
        End If
        Exit Function
    End If
Else
If curitem = -2 Then
    Set peekvalue = GroupRef
    IsObjAt = True
ElseIf curitem < 0 Then
If Not TypeOf GroupRef Is mHandler Then GoTo e1245
Dim usehandler As mHandler
Set usehandler = GroupRef
If usehandler.t1 <> 1 Then
e1245:
    Err.Raise 5010, , "Need an Inventory Object"
End If
    If usehandler.objref.Done Then
        If usehandler.objref.IsObj Then
            Set peekvalue = usehandler.objref.ValueObj
            IsObjAt = True
        Else
            peekvalue = usehandler.objref.Value
        End If
    Else
        Set peekvalue = usehandler
        IsObjAt = True
    End If

End If
End If
there1:
End Function
Function IsEnum(p As Variant) As Boolean
Dim usehandler As mHandler
If Index >= 0 And Index <= maxitems1 Then
If IsObject(actualData(Index)) Then
If Typename(actualData(Index)) = mHdlr Then
    Set usehandler = actualData(Index)
    IsEnum = usehandler.t1 = 4
     If IsEnum Then
        If myVarType(usehandler.index_cursor, vbString) Then
            p = usehandler.index_cursor
        Else
            p = usehandler.index_cursor * usehandler.sign
        End If
     End If
End If
End If
End If
End Function
Private Sub Class_Initialize()
dnum = 0 ' no actions
MyTypeToBe = vbVariant
maxitems = 1
maxitems1 = 0
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
actualData = Array(Empty)
CodeName = GetAnyName$
Arr = True
myarrbase = -ArrBase
refArray = 0&
End Sub

Private Sub Class_Terminate()
Dim z As Variant
CopyMemory ByVal VarPtr(refArray), z, 16
Set GroupRef = Nothing
If Elements(actualData) > 0 Then
If myVarType(actualData, vbObject) Then
ElseIf VarType(actualData) = vbArray + vbVariant Then
Erase actualData
End If
End If
End Sub
Private Function GetAnyName$()
Dim ss$, j As Long
For j = 1 To 4
ss$ = ss$ & Chr(65 + Int((23 * Rnd) + 1))
Next j
GetAnyName$ = ChrW(&H1FFF) + ss$
End Function
Function MyIsNumeric(v As Variant) As Boolean
Dim n As Byte
GetMem1 VarPtr(v), n
If n < 2 Then Exit Function
If n < 8 Then MyIsNumeric = True: Exit Function
MyIsNumeric = (n = 11) Or (n = 17) Or (n = 14) Or (n = 20)
End Function
Function MyIsObject(v As Variant) As Boolean
Dim n As Byte
GetMem1 VarPtr(v), n
MyIsObject = n = 9
End Function
Function GetCopy() As Variant()
GetCopy = actualData
End Function
Private Sub InsertionSort1(ByVal iMin As Long, ByVal iMax As Long)
    Dim i As Long
    Dim j As Long
    i = iMin + 1
    iMax = iMax + 1
    While i < iMax
        j = i
        Do While j > iMin
        If item(j) < item(j - 1) Then
            SwapVariant actualData(j), actualData(j - 1)
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        i = i + 1
    Wend
End Sub
Private Sub InsertionSort1NSW(ByVal iMin As Long, ByVal iMax As Long)
    Dim i As Long, p1
    Dim j As Long
    i = iMin + 1
    iMax = iMax + 1
    While i < iMax
        j = i
        Do While j > iMin
        If item(j) < item(j - 1) Then
            p1 = Empty
             SwapVariant p1, actualData(j - 1)
             item(j - 1) = actualData(j)
             item(j) = p1
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        i = i + 1
    Wend
    p1 = Empty
End Sub
Private Sub InsertionSort(ByVal iMin As Long, ByVal iMax As Long)
    Dim i As Long
    Dim j As Long
    i = iMin + 1
    iMax = iMax + 1
    While i < iMax
        j = i
        Do While j > iMin
        If item(j) > item(j - 1) Then
            SwapVariant actualData(j), actualData(j - 1)
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        i = i + 1
    Wend
End Sub
Private Sub InsertionSortNSW(ByVal iMin As Long, ByVal iMax As Long)
    Dim i As Long, p1
    Dim j As Long
    i = iMin + 1
    iMax = iMax + 1
    While i < iMax
        j = i
        Do While j > iMin
        If item(j) > item(j - 1) Then
             p1 = Empty
             SwapVariant p1, actualData(j - 1)
             item(j - 1) = actualData(j)
             item(j) = p1
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        i = i + 1
    Wend
    p1 = Empty
End Sub

Private Sub QuickSortItemsWithSameKey(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long

  If UB - LB = 1 Then
     M1 = LB

     If item(M1) > item(UB) Then SwapVariant actualData(M1), actualData(UB)
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSort1 LB, UB
  Exit Sub
  
  Else
    M1 = (LB + UB) \ 2
   Piv = item(M1)
   
           M1 = LB
   Do While item(M1) < Piv: M1 = M1 + 1: Loop
    End If
  
  M2 = UB
  Do
    Do While item(M2) > Piv: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      If M1 <> M2 Then SwapVariant actualData(M1), actualData(M2)
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While item(M1) < Piv: M1 = M1 + 1: Loop
  Loop
  If LB < M2 Then QuickSortItemsWithSameKey LB, M2
  If M1 < UB Then QuickSortItemsWithSameKey M1, UB
End Sub
Private Sub QuickSortItemsWithSameKeyO(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long, p1

  If UB - LB = 1 Then
     M1 = LB

     If item(M1) > item(UB) Then
        'SwapVariant actualData(M1), actualData(UB)
             p1 = Empty
             SwapVariant p1, actualData(M1)
             item(M1) = actualData(UB)
             item(UB) = p1
     End If
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSort1NSW LB, UB
  Exit Sub
  
  Else
    M1 = (LB + UB) \ 2
   Piv = item(M1)
   
           M1 = LB
   Do While item(M1) < Piv: M1 = M1 + 1: Loop
    End If
  
  M2 = UB
  Do
    Do While item(M2) > Piv: M2 = M2 - 1: Loop
    If M1 <= M2 Then
        If M1 <> M2 Then
      ' SwapVariant actualData(M1), actualData(M2)
            p1 = Empty
            SwapVariant p1, actualData(M1)
            item(M1) = actualData(M2)
            item(M2) = p1
        End If
        M1 = M1 + 1
        M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While item(M1) < Piv: M1 = M1 + 1: Loop
  Loop
  If LB < M2 Then QuickSortItemsWithSameKeyO LB, M2
  If M1 < UB Then QuickSortItemsWithSameKeyO M1, UB
End Sub

Private Sub QuickSortItemsRevWithSameKey(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
  If UB - LB = 1 Then
     M1 = LB

     If item(M1) < item(UB) Then SwapVariant actualData(M1), actualData(UB)
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSort LB, UB
  Exit Sub
  
  Else
    M1 = (LB + UB) \ 2
   Piv = item(M1)
   
           M1 = LB
   Do While item(M1) > Piv: M1 = M1 + 1: Loop
  End If
  
  M2 = UB
  Do
    Do While item(M2) < Piv: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      If M1 <> M2 Then SwapVariant actualData(M1), actualData(M2)
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While item(M1) > Piv: M1 = M1 + 1: Loop
  Loop
  If LB < M2 Then QuickSortItemsRevWithSameKey LB, M2
  If M1 < UB Then QuickSortItemsRevWithSameKey M1, UB
End Sub
Private Sub QuickSortItemsRevWithSameKeyO(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long, p1
  If UB - LB = 1 Then
     M1 = LB

     If item(M1) < item(UB) Then
     'SwapVariant actualData(M1), actualData(UB)
            p1 = Empty
            SwapVariant p1, actualData(M1)
            item(M1) = actualData(UB)
            item(UB) = p1
     End If
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSortNSW LB, UB
  Exit Sub
  
  Else
    M1 = (LB + UB) \ 2
   Piv = item(M1)
   
           M1 = LB
   Do While item(M1) > Piv: M1 = M1 + 1: Loop
  End If
  
  M2 = UB
  Do
    Do While item(M2) < Piv: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      If M1 <> M2 Then
      'SwapVariant actualData(M1), actualData(M2)
            p1 = Empty
            SwapVariant p1, actualData(M1)
            item(M1) = actualData(M2)
            item(M2) = p1

      End If
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While item(M1) > Piv: M1 = M1 + 1: Loop
  Loop
  If LB < M2 Then QuickSortItemsRevWithSameKeyO LB, M2
  If M1 < UB Then QuickSortItemsRevWithSameKeyO M1, UB
End Sub
Public Sub SortTuple(Optional Low As Long = -1, Optional high As Long = -1)
Dim i As Long, n As Byte
If dnum < 1 Then Exit Sub
If count > 0 Then
If high = -1 Then high = count - 1
If high > count - 1 Then high = count - 1
If Low = -1 Then Low = 0
If Low >= high Then Exit Sub

For i = Low To high
GetMem1 VarPtr(actualData(i)), n
If (n > 8 Or n < 2) And Not (n = 14 Or n = 20) Then Exit Sub
Next i
If MyIsObject(actualData) Then
    QuickSortItemsWithSameKeyO Low, high
Else
    QuickSortItemsWithSameKey Low, high
End If

End If
End Sub
Public Sub SortDesTuple(Optional Low As Long = -1, Optional high As Long = -1)
Dim i As Long, n As Byte
If dnum < 1 Then Exit Sub
If count > 0 Then
If high = -1 Then high = count - 1
If high > count - 1 Then high = count - 1
If Low = -1 Then Low = 0
If Low >= high Then Exit Sub

For i = Low To high
GetMem1 VarPtr(actualData(i)), n
If (n > 8 Or n < 2) And Not (n = 14 Or n = 20) Then Exit Sub
Next i
If MyIsObject(actualData) Then
QuickSortItemsRevWithSameKeyO Low, high
Else
QuickSortItemsRevWithSameKey Low, high
End If
End If
End Sub
Private Sub InsertionSort10obj(ByVal iMin As Long, ByVal iMax As Long)
    Dim i As Long
    Dim j As Long
    Dim j1 As Long, j0 As Long, k As Long, p1

    i = iMin + 1
    iMax = iMax + 1
    j0 = i * sortmult
    While i < iMax
        j = i
        j1 = j0 + sortoff
        Do While j > iMin
        ''Debug.Print "insertion compare"
again:
        On Error GoTo ex1
        If item(j1) = item(j1 - sortmult) Then
        part = part + 2
        If part <= upart Then
        
        
        j1 = j1 - sortoff + Col(part - 1)
        sortoff = Col(part - 1)
        GoTo again
        End If
ex1:
        Exit Do
        Else
        
        If Col(part) = 0 Then
        If item(j1) < item(j1 - sortmult) Then
           ' Debug.Print "insertion swap"
            For k = j1 - sortoff To j1 - sortoff + sortmult - 1
'            SwapVariant actualData(k), actualData(k - sortmult)
            p1 = Empty
            SwapVariant p1, actualData(k)
            item(k) = actualData(k - sortmult)
            item(k - sortmult) = p1

            
            
            
            Next k
            j = j - 1
            If part <> 3 Then
                j1 = j1 - sortoff - sortmult + Col(2)
                part = 3
                sortoff = Col(2)
            Else
                j1 = j1 - sortmult
            End If
        Else
            Exit Do
        End If
        Else
        If item(j1) > item(j1 - sortmult) Then
           ' Debug.Print "insertion swap"
            For k = j1 - sortoff To j1 - sortoff + sortmult - 1
            'SwapVariant actualData(k), actualData(k - sortmult)
            p1 = Empty
            SwapVariant p1, actualData(k)
            item(k) = actualData(k - sortmult)
            item(k - sortmult) = p1
            Next k
            j = j - 1
            If part <> 3 Then
                j1 = j1 - sortoff - sortmult + Col(2)
                part = 3
                sortoff = Col(2)
            Else
                j1 = j1 - sortmult
            End If
            
        Else
            Exit Do
        End If
        End If
        End If
        Loop
        If part <> 3 Then
        part = 3
        sortoff = Col(2)
        End If
        i = i + 1
        j0 = j0 + sortmult
    Wend
End Sub
Private Sub InsertionSort10(ByVal iMin As Long, ByVal iMax As Long)
    Dim i As Long
    Dim j As Long
    Dim j1 As Long, j0 As Long, k As Long

    i = iMin + 1
    iMax = iMax + 1
    j0 = i * sortmult
    While i < iMax
        j = i
        j1 = j0 + sortoff
        Do While j > iMin
        ''Debug.Print "insertion compare"
again:
        On Error GoTo ex1
        If item(j1) = item(j1 - sortmult) Then
        part = part + 2
        If part <= upart Then
        
        
        j1 = j1 - sortoff + Col(part - 1)
        sortoff = Col(part - 1)
        GoTo again
        End If
ex1:
        Exit Do
        Else
        
        If Col(part) = 0 Then
        If item(j1) < item(j1 - sortmult) Then
           ' Debug.Print "insertion swap"
            For k = j1 - sortoff To j1 - sortoff + sortmult - 1
            SwapVariant actualData(k), actualData(k - sortmult)
            Next k
            j = j - 1
            If part <> 3 Then
                j1 = j1 - sortoff - sortmult + Col(2)
                part = 3
                sortoff = Col(2)
            Else
                j1 = j1 - sortmult
            End If
        Else
            Exit Do
        End If
        Else
        If item(j1) > item(j1 - sortmult) Then
           ' Debug.Print "insertion swap"
            For k = j1 - sortoff To j1 - sortoff + sortmult - 1
            SwapVariant actualData(k), actualData(k - sortmult)
            Next k
            j = j - 1
            If part <> 3 Then
                j1 = j1 - sortoff - sortmult + Col(2)
                part = 3
                sortoff = Col(2)
            Else
                j1 = j1 - sortmult
            End If
            
        Else
            Exit Do
        End If
        End If
        End If
        Loop
        If part <> 3 Then
        part = 3
        sortoff = Col(2)
        End If
        i = i + 1
        j0 = j0 + sortmult
    Wend
End Sub
Private Sub Swap10(ByVal a1, ByVal A2)
Static k As Long
If a1 = A2 Then Exit Sub
'Debug.Print "swap ", a1, A2
a1 = a1 * sortmult
A2 = A2 * sortmult
For k = 0 To sortmult - 1
SwapVariant actualData(a1 + k), actualData(A2 + k)
Next k
End Sub
Private Sub Swap10obj(ByVal a1, ByVal A2)
Static k As Long
Dim p1
If a1 = A2 Then Exit Sub
'Debug.Print "swap ", a1, A2
a1 = a1 * sortmult
A2 = A2 * sortmult
For k = 0 To sortmult - 1
p1 = Empty
SwapVariant p1, actualData(a1 + k)
item(a1 + k) = actualData(A2 + k)
item(A2 + k) = p1
'SwapVariant actualData(a1 + k), actualData(A2 + k)
Next k
End Sub
Private Sub FeedPiv(ByVal a1)
Static k As Long
On Error Resume Next
Dim ii As Variant
Dim n As Byte
'Debug.Print "feed piv "; a1
a1 = a1 * sortmult
For k = 0 To sortmult - 1
    GetMem1 VarPtr(actualData(a1 + k)), n
    If n = 20 Then
       CopyMemory ByVal VarPtr(piv1(k)), ByVal VarPtr(actualData(a1 + k)), 16
    Else
        piv1(k) = actualData(a1 + k)
    End If
Next k
End Sub

Private Function Compare10(ByVal a1) As Long
Static k As Long, ret As Long
On Error Resume Next
Dim ii As Variant, ii2 As Variant, ii3 As Variant
Dim n As Byte, n1 As Byte, n2 As Byte
a1 = a1 * sortmult
For k = 2 To upart - 1 Step 2
n2 = 0
GetMem1 VarPtr(actualData(a1 + Col(k))), n
If n = 20 Then
    CopyMemory ByVal VarPtr(ii) + 8, ByVal VarPtr(actualData(a1 + Col(k))) + 8, 8
    PutMem2 VarPtr(ii), 20
    GetMem1 VarPtr(piv1(Col(k))), n1
    If n1 = 20 Then
        CopyMemory ByVal VarPtr(ii2) + 8, ByVal VarPtr(piv1(Col(k))) + 8, 8
        PutMem2 VarPtr(ii2), 20
        ret = ii = ii2
    Else
        ret = ii = piv1(Col(k))
    End If
Else
    ret = actualData(a1 + Col(k)) = piv1(Col(k))
End If

If ret = 0 Then
    If n = 20 And n1 = 20 Then
        GetMem1 VarPtr(piv1(Col(k + 1))), n2
        If n2 = 20 Then
            CopyMemory ByVal VarPtr(ii3) + 8, ByVal VarPtr(piv1(Col(k + 1))) + 8, 8
            PutMem2 VarPtr(ii3), 20
            If ii < ii2 Then
                If ii3 Then
                    Compare10 = -1
                Else
                    Compare10 = 1
                End If
                Exit Function
            Else
            
                If ii3 Then
                    Compare10 = 1
                Else
                    Compare10 = -1
                End If
                Exit Function
            End If
        Else
            If ii < ii2 Then
                If Col(k + 1) Then
                    Compare10 = -1
                Else
                    Compare10 = 1
                End If
                Exit Function
            Else
            
                If Col(k + 1) Then
                    Compare10 = 1
                Else
                    Compare10 = -1
                End If
                Exit Function
            End If
        End If
    ElseIf n = 20 Then
        GetMem1 VarPtr(piv1(Col(k + 1))), n2
        If n2 = 20 Then
            CopyMemory ByVal VarPtr(ii3) + 8, ByVal VarPtr(piv1(Col(k + 1))) + 8, 8
            PutMem2 VarPtr(ii3), 20
            If ii < piv1(Col(k)) Then
                If ii3 Then
                    Compare10 = -1
                Else
                    Compare10 = 1
                End If
                Exit Function
            Else
                If ii3 Then
                    Compare10 = 1
                Else
                    Compare10 = -1
                End If
                Exit Function
            End If
        Else
            If ii < piv1(Col(k)) Then
                If Col(k + 1) Then
                    Compare10 = -1
                Else
                    Compare10 = 1
                End If
                Exit Function
            Else
                If Col(k + 1) Then
                    Compare10 = 1
                Else
                    Compare10 = -1
                End If
                Exit Function
            End If
        End If
    Else
        If actualData(a1 + Col(k)) < piv1(Col(k)) Then
            If Col(k + 1) Then
                Compare10 = -1
            Else
                Compare10 = 1
            End If
            Exit Function
        Else
            If Col(k + 1) Then
                Compare10 = 1
            Else
                Compare10 = -1
            End If
            Exit Function
        End If
    End If
End If
Next k
' Compare10=0
End Function
Private Sub QuickSortItemsWithSameKey10obj(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long

  If UB - LB < 7 Then
  InsertionSort10obj LB, UB
  Exit Sub
  
  End If
    M1 = (LB + UB) \ 2
    FeedPiv M1
    
    M1 = LB
    Do While Compare10(M1) = 1: M1 = M1 + 1: Loop
  
  M2 = UB
  Do
    Do While Compare10(M2) = -1: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      Swap10obj M1, M2
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While Compare10(M1) = 1: M1 = M1 + 1: Loop
  Loop
  If LB < M2 Then QuickSortItemsWithSameKey10obj LB, M2
  If M1 < UB Then QuickSortItemsWithSameKey10obj M1, UB
End Sub
Private Sub QuickSortItemsWithSameKey10(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long

  If UB - LB < 7 Then
  'Debug.Print LB, UB
  InsertionSort10 LB, UB
  Exit Sub
  
  End If
    M1 = (LB + UB) \ 2
    FeedPiv M1
    
    M1 = LB
    Do While Compare10(M1) = 1: M1 = M1 + 1: Loop
  
  M2 = UB
  Do
    Do While Compare10(M2) = -1: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      Swap10 M1, M2
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While Compare10(M1) = 1: M1 = M1 + 1: Loop
  Loop
  If LB < M2 Then QuickSortItemsWithSameKey10 LB, M2
  If M1 < UB Then QuickSortItemsWithSameKey10 M1, UB
End Sub

Public Sub SortColumns(ParamArray col0())
Dim usehandler As mHandler
If dnum <> 2 Then Exit Sub
Dim mbase2 As Long

Dim i As Long, pppp As mArray
GetDnum 1, i, mbase2
If UBound(col0()) = 0 Then
If CheckInt64(col0(0)) Then
' typename has problem with 20
ElseIf VarTypeName(col0(0)) = "Variant()" Then
    col0() = col0(0)
ElseIf VarTypeName(col0(0)) = mHdlr Then
    Set usehandler = col0(0)
    If usehandler.t1 = 3 Then
        Set pppp = usehandler.objref
        Set usehandler = Nothing
        pppp.ExportArrayNow
        col0() = pppp.refArray
    Else
        Exit Sub
    End If
End If

End If
Col() = col0()
If Typename(Col) = "Variant()" Then
Dim Col2()
Col2() = Col(0)
ReDim Col(UBound(Col2))
Col() = Col2()

End If
sortmult = dtable(0)
If UBound(Col()) < 1 Then

ReDim Col(4)
Col(0) = 0
Col(1) = count \ sortmult - 1
Col(2) = 0
Col(3) = 0
ElseIf UBound(Col()) < 3 Then

ReDim Preserve Col(4)
Col(2) = -mbase2
Col(3) = 0
GoTo there1

Else
'For i = LBound(col0) To UBound(col0)
'Col(i) = CLng(col0(i))
'Next i
there1:
If Col(1) = -1 Then Col(1) = count \ sortmult - 1

If Col(1) > count \ sortmult - 1 Then Col(1) = count \ sortmult - 1
If Col(0) < 0 Then Col(0) = 0
If Col(0) >= Col(1) Then Exit Sub

For i = LBound(Col) + 2 To UBound(Col) - 1 Step 2
Col(i) = Col(i) + mbase2
If Col(i) < 0 Or Col(i) >= sortmult Then Exit Sub
'If col(i + 1) <> 0 Then col(i + 1) = 1
Next i

End If
sortoff = Col(2)
upart = UBound(Col())
part = 3

ReDim piv1(sortmult)
If Stable Then
    If MyIsObject(actualData) Then
        InsertionSort10obj Col(0), Col(1)
    Else
        InsertionSort10 Col(0), Col(1)
    End If
Else
If MyIsObject(actualData) Then
    QuickSortItemsWithSameKey10obj Col(0), Col(1)
Else
    QuickSortItemsWithSameKey10 Col(0), Col(1)
End If
End If
ReDim piv1(1)
End Sub
Property Get ArrPtr() As Long  ' 32 bit address
Dim ar As refArray
If myVarType(actualData, vbObject) Then
    If actualData.count > 0 Then
        Set ar = actualData
        ArrPtr = ar.ArrPtr
    End If
ElseIf Elements(actualData) > 0 Then
    ArrPtr = VarPtr(actualData(0))
End If
End Property
Function Elements(vArray As Variant) As Long
    Dim lPtr            As Long
    Const VT_BYREF      As Long = &H4000
    If IsArray(vArray) Then
    lPtr = Peek(UnsignedAdd(VarPtr(vArray), 8))
    If (Peek(VarPtr(vArray)) And VT_BYREF) <> 0 Then
        lPtr = Peek(lPtr)
    End If
    If lPtr <> 0 Then
        Elements = Peek(UnsignedAdd(lPtr, 16))
    End If
    End If
End Function
Public Function Peek(ByVal lPtr As Long) As Long
    GetMem4 lPtr, Peek
End Function

Public Function PeekInt(ByVal lPtr As Long) As Integer
    GetMem2 lPtr, PeekInt
End Function
Public Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Property Let iBoxArray_Arr(ByVal RHS As Boolean)
Arr = RHS
End Property

Private Property Get iBoxArray_Arr() As Boolean
iBoxArray_Arr = Arr
End Property

Private Property Let iBoxArray_arrname(ByVal RHS As String)
arrname = RHS
End Property

Private Property Get iBoxArray_arrname() As String
iBoxArray_arrname = mmname
End Property

Private Property Let iBoxArray_CodeName(RHS As String)
' nothing
End Property

Private Property Get iBoxArray_CodeName() As String
iBoxArray_CodeName = CodeName
End Property

Private Property Let iBoxArray_common(ByVal RHS As Boolean)
common = RHS
End Property

Private Property Get iBoxArray_common() As Boolean
iBoxArray_common = common
End Property

Private Property Let iBoxArray_Final(ByVal RHS As Boolean)
Final = RHS
End Property

Private Property Get iBoxArray_Final() As Boolean
iBoxArray_Final = Final
End Property

Private Property Set iBoxArray_GroupRef(RHS As Object)
Set GroupRef = RHS
End Property

Private Property Get iBoxArray_GroupRef() As Object
Set iBoxArray_GroupRef = GroupRef
End Property

Private Property Let iBoxArray_IhaveClass(ByVal RHS As Boolean)
    IhaveClass = RHS
End Property

Private Property Get iBoxArray_IhaveClass() As Boolean
    iBoxArray_IhaveClass = IhaveClass
End Property

Private Property Get iBoxArray_IsEmpty() As Variant
    iBoxArray_IsEmpty = IsEmpty
End Property

Private Function iBoxArray_IsObjAt(curitem As Long, peekvalue As Variant) As Boolean
iBoxArray_IsObjAt = IsObjAt(curitem, peekvalue)
End Function

Private Property Set iBoxArray_Item(curitem As Long, RHS As Variant)
Set item(curitem) = RHS
End Property

Private Property Let iBoxArray_Item(curitem As Long, RHS As Variant)
' not used
End Property

Private Property Get iBoxArray_Item(curitem As Long) As Variant
If Me.ItemIsObject(curitem) Then
Set iBoxArray_Item = Me.item(curitem)
Else
iBoxArray_Item = Me.item(curitem)
End If
End Property

Private Property Get iBoxArray_itemObject(curitem As Long) As Variant
Set iBoxArray_itemObject = itemObject(curitem)
End Property

Private Property Get iBoxArray_ItemType(curitem As Variant) As String
    iBoxArray_ItemType = ItemType(curitem)
End Property

Private Sub iBoxArray_PushDim(nDim As Long, Optional limBase As Variant)
If Not IsMissing(limBase) Then
    PushDim nDim, limBase
Else
    PushDim nDim
End If
End Sub

Private Sub iBoxArray_PushEnd(Optional thisarr As Variant)
    If Not IsMissing(thisarr) Then
        PushEnd thisarr
    Else
        PushEnd
    End If
End Sub

Private Sub iBoxArray_PushProp(A As Variant)
PushProp A
End Sub

Private Property Set iBoxArray_refgroup(ByVal RHS As Object)
Set refgroup = RHS
End Property


Private Property Get iBoxArray_refgroup() As Object
Set iBoxArray_refgroup = refgroup
End Property
