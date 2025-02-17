VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' tuple is ONE dimension array
Implements iBoxArray
Dim mmname As String
Dim dnum As Long, ndnum As Long
'Dim limit(10) As Long  ' max 9 dimensions
'Dim dtable(10) As Long
'Dim nlimit(10) As Long  ' max 9 dimensions
'Dim limBase(10) As Long
'Dim nLimbase(10) As Long
Private Const mHdlr = "mHandler"
Private Const mgrp = "Group"
Private Const mlmbd = "lambda"
Private Const mArr = "mArray"
Private Const mTuple = "tuple"
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
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, retval As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, retval As Byte)
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
Private Type ValueCarrier
    val As Variant
    obj As BigInteger
    typ As Integer
End Type
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
        Case "DIV", "ƒ…¡"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            item(v) = Fix(ii / p)
        Case "DIV#", "ƒ…¡#"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            If p < 0 Then
                item(v) = Int((ii - Abs(ii - Abs(p) * Int(ii / Abs(p)))) / p)
            Else
                item(v) = Int(ii / p)
            End If
        Case "MOD", "’–œÀ", "’–œÀœ…–œ"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            sp = ii - Fix(ii / p) * p
            If Abs(sp) >= Abs(p) Then sp = sp - sp
            actualData(v) = sp
        Case "MOD#", "’–œÀ#", "’–œÀœ…–œ#"
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
        Case "DIV", "ƒ…¡"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            actualData(v) = Fix(actualData(v) / p)
        Case "DIV#", "ƒ…¡#"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            If p < 0 Then
                actualData(v) = Int((actualData(v) - Abs(actualData(v) - Abs(p) * Int(actualData(v) / Abs(p)))) / p)
            Else
                actualData(v) = Int(actualData(v) / p)
            End If
        Case "MOD", "’–œÀ", "’–œÀœ…–œ"
            If p = 0 Then
                DevZero
                Exit Do
            End If
            sp = actualData(v) - Fix(actualData(v) / p) * p
            If Abs(sp) >= Abs(p) Then sp = sp - sp
            actualData(v) = sp
        Case "MOD#", "’–œÀ#", "’–œÀœ…–œ#"
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
Dim n As Byte, BI As BigInteger, BI2 As BigInteger, rd As Boolean, ctype As Boolean

rd = RoundDouble
On Error Resume Next
ctype = myVarType(actualData, vbObject)

If MemInt(VarPtr(p)) = 9 Then
    If TypeOf p Is BigInteger Then
    Set BI2 = p
    p = CDec(BI2.ToString)
    If Err.Number > 0 Then
        p = Infinity()
    End If
    Else
        p = 0
    End If
End If

v = 0
Do While v <= Me.UpperMonoLimit
    GetMem1 VarPtr(actualData(v)), n
    If n = 9 Then
        If TypeOf actualData(v) Is BigInteger Then
            Set BI = actualData(v)
            If BI2 Is Nothing Then
                Set BI2 = Module13.CreateBigInteger(CStr(Int(p)))
            End If
            bigintOperations CVar(BI2), BI, ss$
            If ctype Then
                actualData(v) = CVar(BI)
            Else
                Set actualData(v) = BI
            End If
            Set BI = Nothing
        End If
        GoTo conthere
    ElseIf n = 20 Then
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
            Case "DIV", "ƒ…¡"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                item(v) = Fix(ii / p)
            Case "DIV#", "ƒ…¡#"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                If p < 0 Then
                    item(v) = Int((ii - Abs(ii - Abs(p) * Int(ii / Abs(p)))) / p)
                Else
                    item(v) = Int(ii / p)
                End If
            Case "MOD", "’–œÀ", "’–œÀœ…–œ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = ii - Fix(ii / p) * p
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                item(v) = sp
             Case "MOD#", "’–œÀ#", "’–œÀœ…–œ#"
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
            Case "DIV", "ƒ…¡"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                actualData(v) = Fix(actualData(v) / p)
            Case "DIV#", "ƒ…¡#"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                If p < 0 Then
                    actualData(v) = Int((actualData(v) - Abs(actualData(v) - Abs(p) * Int(actualData(v) / Abs(p)))) / p)
                Else
                    actualData(v) = Int(actualData(v) / p)
                End If
            Case "MOD", "’–œÀ", "’–œÀœ…–œ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = actualData(v) - Fix(actualData(v) / p) * p
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                actualData(v) = sp
             Case "MOD#", "’–œÀ#", "’–œÀœ…–œ#"
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
        ElseIf n = vbString Then
            If MemInt(VarPtr(p)) <> vbString Then
                p = fixthis(p)
            End If
            Select Case ss$
                Case "="
                    actualData(v) = p
                    GoTo conthere
                Case "+="
                    actualData(v) = actualData(v) + p
            End Select
        End If
    If rd Then If n = vbDouble Then actualData(v) = MyRound(actualData(v), 13)
conthere:
    v = v + 1
    Err.Clear
Loop
End Sub
Sub bigintOperations1(ByRef BI As BigInteger, ss$)
If ss$ = "++" Then
    Set BI = BI.Add(Module13.CreateBigInteger("1"))
ElseIf ss$ = "--" Then
    Set BI = BI.Add(Module13.CreateBigInteger("-1"))
ElseIf ss$ = "-!" Then
    BI.negate
End If
End Sub
Sub bigintOperations(BI2 As BigInteger, ByRef BI As BigInteger, ss$)
If ss$ = "=" Then
    Set BI = BI2
ElseIf ss$ = "*=" Then
    Set BI = BI.multiply(BI2)
ElseIf ss$ = "/=" Then
    Set BI = BI.divide(BI2)
ElseIf ss$ = "+=" Then
    Set BI = BI.Add(BI2)
ElseIf ss$ = "-=" Then
    Set BI = BI.subtract(BI2)
Else
Select Case ss$
Case "MOD", "’–œÀ", "’–œÀœ…–œ"
    Set BI = BI.Modulus(BI2)
Case "DIV", "ƒ…¡"
        Set BI = BI.divide(BI2)
End Select
End If
End Sub
Public Sub Compute3(ss$)
Dim v As Long, ii As Variant
Dim n As Byte, ctype As Boolean
Dim BI As BigInteger
    
    If Final Then CantAssignValue: Exit Sub
    On Error Resume Next
    ctype = myVarType(actualData, vbObject)
    v = 0
    Do While v <= Me.UpperMonoLimit
        GetMem1 VarPtr(actualData(v)), n
        If n = 9 Then
            If TypeOf actualData(v) Is BigInteger Then
                Set BI = actualData(v)
                bigintOperations1 BI, ss$
                If ctype Then
                    actualData(v) = CVar(BI)
                Else
                    Set actualData(v) = BI
                End If
                Set BI = Nothing
            End If
        ElseIf n = 20 Then
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
Public Sub Neg(v As Long)
Select Case VarType(itemnumeric(v))
Case vbBoolean
    item(v) = Not CBool(itemnumeric(v))
Case vbInteger
    item(v) = CInt(Not CBool(itemnumeric(v)))
Case vbLong
    item(v) = CLng(Not CBool(itemnumeric(v)))
Case vbCurrency
    
    item(v) = CCur(Not CBool(itemnumeric(v)))
Case vbDecimal
    item(v) = CDec(Not CBool(itemnumeric(v)))
Case 20
    If CBool(itemnumeric(v)) Then
    item(v) = -OneLongLong()
    Else
    item(v) = -OneLongLong() And CVar(0)
    End If
Case Else
    item(v) = CDbl(Not CBool(itemnumeric(v)))
End Select
End Sub
Public Function PlaceValue2UDTArray(ByVal p As Long, Name$, v, Index As Long) As Boolean
On Error Resume Next
Module10.PlaceValue2UDTArray actualData(p), Name$, v, Index

If Err Then
    MyEr Err.Description, Err.Description
    Err.Clear
    Exit Function
End If
PlaceValue2UDTArray = True
End Function
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
Public Function GetUDTValue(ByVal p As Long, Name$, Optional ndx)
On Error Resume Next
If IsMissing(ndx) Then
GetUDTValue = Module10.GetUDTValue(actualData(p), Name$)
Else
GetUDTValue = Module10.GetUDTValueArray(actualData(p), Name$, CLng(ndx))
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
'
End Sub
Friend Sub FixArray()
'
End Sub
Friend Sub LoadTuple(Ar As Variant)
Dim z As Long, i As Long
actualData = Ar
If Count <> UBound(actualData) - LBound(actualData) + 1 Then
i = Count
maxitems1 = UBound(actualData) - LBound(actualData)
maxitems = maxitems1 + 1
z = 1
dnum = 1
End If
'
End Sub
Friend Sub CopySerialize(mmm As Variant)
' may we get an exception here
If VarType(actualData) = vbArray + vbVariant Then
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, r1 As Long, R2 As Long
GetMem4 UnsignedAdd(VarPtr(actualData), 8), m_base_0
If m_base_0 <> 0 Then
GetMem4 LibMemory.ArrPtr(mmm), m_base
'GetMem4 VarPtrArray(actualdata()), m_base_0
r1 = SafeArrayLock(m_base)
R2 = SafeArrayLock(m_base_0)
r1 = SafeArrayAccessData(ByVal m_base, first12)
R2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
R2 = SafeArrayUnaccessData(ByVal m_base_0)
r1 = SafeArrayUnaccessData(ByVal m_base)
R2 = SafeArrayUnlock(m_base_0)
r1 = SafeArrayUnlock(m_base)
End If
End If
End Sub
Friend Sub CopyBackSerialize(mmm As Variant)
' may we get an exception here
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, r1 As Long, R2 As Long
If Not RevOrder Then Err.Raise vbObjectError + 12345, "CopyBack", "Failed. is not vb6"
Dim backup
backup = actualData
GetMem4 LibMemory.ArrPtr(mmm), m_base
GetMem4 LibMemory.ArrPtr(backup), m_base_0
r1 = SafeArrayLock(m_base)
R2 = SafeArrayLock(m_base_0)
r1 = SafeArrayAccessData(ByVal m_base, first12)
R2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
R2 = SafeArrayUnaccessData(ByVal m_base_0)
r1 = SafeArrayUnaccessData(ByVal m_base)
R2 = SafeArrayUnlock(m_base_0)
r1 = SafeArrayUnlock(m_base)
End Sub
Public Sub Car(m As Object)
Dim K As New tuple, usehandler As mHandler, s$
Dim i As Long, obj As Object, obj2 As Object, spare As lambda
If Count > 0 Then
    K.PushDim 1: K.PushEnd: K.SerialItem actualData(0), 0, 1
    If ItemIsObject(0) Then
        s$ = ItemType(0)
        Select Case s$
        Case mgrp
            Set obj = item(0)
            CopyGroupObj obj, obj2
            Set K.item(0) = obj2
            Set obj2 = Nothing
        Case mlmbd
            Dim actualdata1()
            With item(0)
            .CopyTo spare, actualdata1()
            End With
            Set K.item(0) = spare
            Set spare = Nothing
        Case mHdlr
                Set usehandler = item(0)
                If usehandler.T1 = 3 Then
                    Set m = usehandler
                    Set K = Nothing
                    Exit Sub
                End If
        Case mArr, mTuple
                Set usehandler = New mHandler
                usehandler.T1 = 3
                Set usehandler.objref = item(0)
                Set m = usehandler
                Set K = Nothing
                Exit Sub
        Case Else
            Set K.item(0) = item(0)
        End Select
    Else
        K.item(0) = item(0)
    End If
End If
K.Arr = True
Set m = K
Set K = Nothing
End Sub
Public Sub Cdr(m As Object)
Dim K As New tuple
Dim i As Long, obj As Object, obj2 As Object, spare As lambda
If dnum > 0 Then
    K.StartResize: K.PushDim Me.Count - 1: K.PushEnd
Else
    K.Arr = True
    Set K.GroupRef = Me.GroupRef
    K.IHaveGui = Me.IHaveGui
    K.IhaveClass = Me.IhaveClass
    Set m = K
    Exit Sub
End If
K.myarrbase = myarrbase
K.RevOrder = False
For i = 1 To Me.UpperMonoLimit
    If IsObject(Me.item(i)) Then
        Select Case Typename(Me.item(i))
        Case mgrp
            Set obj = Me.item(i)
            CopyGroupObj obj, obj2
            Set K.item(i - 1) = obj2
            Set obj2 = Nothing
        Case mlmbd
            Dim actualdata1()
            With Me.item(i)
                .CopyTo spare, actualdata1()
            End With
            Set K.item(i - 1) = spare
            Set spare = Nothing
        Case Else
            Set K.item(i - 1) = Me.item(i)
        End Select
    Else
        K.item(i - 1) = Me.item(i)
    End If
Next i
' why not in CAR ???
Set K.GroupRef = Me.GroupRef
K.IHaveGui = Me.IHaveGui
K.IhaveClass = Me.IhaveClass
Set m = K
End Sub
Public Sub processAppend(m As Object)
Dim pppp As New tuple, pppp1 As tuple, pppp0  As mArray
If TypeOf m Is mArray Then
    Set pppp0 = m
    pppp0.CopyArray2tuple pppp1
Else
    Set pppp1 = m
End If
If dnum > 0 Then
    Me.CopyArray pppp
    If Not pppp1.IsEmpty Then pppp1.AppendArray pppp
ElseIf Not pppp1.IsEmpty Then
    Set pppp = pppp1
End If
pppp.Arr = True
pppp.Final = False
Set m = pppp
Set pppp = Nothing
Set pppp1 = Nothing
End Sub

Public Sub AppendArray(K As tuple)
'append this to k
If K.Count = 0 Then
    Me.CopyArray K
    Exit Sub
End If
Dim v As Variant, i As Long, obj As Object, obj2 As Object
Dim fromhere As Long
fromhere = K.UpperMonoLimit + 1
K.SerialItem v, Me.Count + K.Count, 10
K.myarrbase = myarrbase
K.RevOrder = RevOrder
For i = 0 To Me.UpperMonoLimit
    If IsObject(Me.item(i)) Then
        Select Case Typename(Me.item(i))
        Case mgrp
            Set obj = Me.item(i)
            CopyGroupObj obj, obj2
            Set K.item(i + fromhere) = obj2
            Set obj2 = Nothing
        Case Else
            Set K.item(i + fromhere) = Me.item(i)
        End Select
    Else
        K.item(i + fromhere) = Me.item(i)
    End If
Next i
'' new
Set K.GroupRef = Me.GroupRef
K.IHaveGui = Me.IHaveGui
K.IhaveClass = Me.IhaveClass
End Sub



Public Sub CopyTuple2Array(K As mArray)
'copy this to k
If K Is Nothing Then Set K = New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
K.MyTypeToBe = MyTypeToBe
' no use, we do the redim in PushEnd
'k.SerialItem v, Me.UpperMonoLimit, 8
K.myarrbase = myarrbase
K.RevOrder = False
K.StartResize
K.PushDim Count, 0
If myVarType(actualData, vbObject) And MyTypeToBe <> vbVariant Then
    K.PushEnd CVar(actualData.Copy)
    K.RevOrder = RevOrder
    Set K.GroupRef = Me.GroupRef
Else
    K.PushEnd
    K.RevOrder = RevOrder
    Set K.GroupRef = Me.GroupRef
    For i = 0 To Me.UpperMonoLimit
        If myVarType(Me.item(i), vbObject) Then
            Select Case Typename(Me.item(i))
            Case mgrp
                Set obj = Me.item(i)
                CopyGroupObj obj, obj2
                Set K.item(i) = obj2
                Set obj2 = Nothing
            Case mlmbd
                Dim actualdata1()
                With Me.item(i)
                .CopyTo spare, actualdata1()
                End With
                Set K.item(i) = spare
                Set spare = Nothing
            Case "RefArray"
                Set K.item(i) = Me.item(i).Copy()
            Case Else
                Set K.item(i) = Me.item(i)
            End Select
        Else
            K.item(i) = Me.item(i)
        End If
    Next i
'' new
End If
K.Arr = Me.Arr
K.IHaveGui = Me.IHaveGui
K.IhaveClass = Me.IhaveClass
K.Final = Me.Final
End Sub


Public Sub CopyArray(K As tuple)
'copy this to k

If K Is Me Then Exit Sub
If K Is Nothing Then Set K = New tuple
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
K.MyTypeToBe = MyTypeToBe
' no use, we do the redim in PushEnd
'k.SerialItem v, Me.UpperMonoLimit, 8
K.myarrbase = myarrbase
K.RevOrder = False
K.StartResize
K.PushDim Count, 0

If myVarType(actualData, vbObject) And MyTypeToBe <> vbVariant Then
    K.PushEnd CVar(actualData.Copy)
    K.RevOrder = RevOrder
    Set K.GroupRef = Me.GroupRef
Else
    K.PushEnd
    K.RevOrder = RevOrder
    'Set K.GroupRef = Me.GroupRef
    For i = 0 To Me.UpperMonoLimit
        If myVarType(Me.item(i), vbObject) Then
            Select Case Typename(Me.item(i))
            Case mgrp
                Set obj = Me.item(i)
                CopyGroupObj obj, obj2
                Set K.item(i) = obj2
                Set obj2 = Nothing
            Case mlmbd
                Dim actualdata1()
                With Me.item(i)
                .CopyTo spare, actualdata1()
                End With
                Set K.item(i) = spare
                Set spare = Nothing
            Case "RefArray"
                Set K.item(i) = Me.item(i).Copy()
            Case Else
                Set K.item(i) = Me.item(i)
            End Select
        Else
            K.item(i) = Me.item(i)
        End If
    Next i
'' new
End If
K.Arr = Me.Arr
K.IHaveGui = Me.IHaveGui
K.IhaveClass = False 'Me.IhaveClass
K.Final = Me.Final
End Sub
Public Sub CopyArrayRev(K As tuple)
'copy this to k
If K Is Me Then Exit Sub
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
Dim mm As Long, M1 As Long
If K Is Nothing Then Set K = New tuple
K.MyTypeToBe = MyTypeToBe
'k.SerialItem v, Me.UpperMonoLimit, 8
K.myarrbase = myarrbase
K.RevOrder = False
K.StartResize
K.PushDim Count, 0
K.PushEnd
K.RevOrder = RevOrder
'Set K.GroupRef = Me.GroupRef
mm = Me.UpperMonoLimit
If myVarType(actualData, vbObject) And MyTypeToBe <> vbVariant Then
    For i = 0 To Me.UpperMonoLimit
        M1 = mm - i
        K.item(i) = Me.item(M1)
    Next i
Else
For i = 0 To Me.UpperMonoLimit
    M1 = mm - i
    If myVarType(Me.item(M1), vbObject) Then
        Select Case Typename(Me.item(M1))
        Case mgrp
            Set obj = Me.item(i)
            CopyGroupObj obj, obj2
            Set K.item(i) = obj2
            Set obj2 = Nothing
        Case mlmbd
            Dim actualdata1()
            With Me.item(M1)
                .CopyTo spare, actualdata1()
            End With
            Set K.item(i) = spare
            Set spare = Nothing
        Case "RefArray"
            Set K.item(i) = Me.item(i).Copy()
        Case Else
            Set K.item(i) = Me.item(M1)
        End Select
    Else
        K.item(i) = Me.item(M1)
    End If
Next i
End If
'' new
K.Arr = Me.Arr
K.IHaveGui = Me.IHaveGui
K.IhaveClass = False 'Me.IhaveClass
K.Final = Me.Final
End Sub
Public Sub CopyArrayRevFast(K As tuple)
'copy this to k
If K Is Me Then Exit Sub
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
Dim mm As Long, M1 As Long
K.MyTypeToBe = MyTypeToBe
'k.SerialItem v, Me.UpperMonoLimit, 8
K.myarrbase = myarrbase
K.RevOrder = False
K.StartResize
K.PushDim Count, 0
K.PushEnd
K.RevOrder = RevOrder
Set K.GroupRef = Me.GroupRef
mm = Me.UpperMonoLimit
If myVarType(actualData, vbObject) And MyTypeToBe <> vbVariant Then
For i = 0 To Me.UpperMonoLimit
    M1 = mm - i
    K.item(i) = actualData(M1)
Next i
Else
For i = 0 To Me.UpperMonoLimit
    M1 = mm - i
    If myVarType(Me.item(M1), vbObject) Then
        Set K.item(i) = Me.item(M1)
    Else
        K.item(i) = Me.item(M1)
    End If
Next i
End If
'' new
K.Arr = Me.Arr
K.IHaveGui = Me.IHaveGui
K.IhaveClass = Me.IhaveClass
K.Final = Me.Final
End Sub
Public Function EmptyArraySameType() As Object
Dim K As tuple
Set K = New tuple
K.MyTypeToBe = MyTypeToBe
K.myarrbase = myarrbase
K.RevOrder = False
K.StartResize
K.PushEnd
K.RevOrder = RevOrder
Set K.GroupRef = Nothing
K.Arr = Me.Arr
K.IHaveGui = Me.IHaveGui
K.IhaveClass = False
K.Final = Me.Final
Set EmptyArraySameType = K
End Function
Public Sub CopyArraySlice(K As tuple, fromS As Long, toS As Long)
'copy this to k
' make one dimemsion only
If K Is Me Then Exit Sub
If K Is Nothing Then Set K = New tuple
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
K.MyTypeToBe = MyTypeToBe
'k.SerialItem v, toS - fromS + 1, 8
K.myarrbase = myarrbase
K.RevOrder = False
K.StartResize
For i = 0 To 0
    K.PushDim toS - fromS + 1, 0
Next i
K.PushEnd
K.RevOrder = RevOrder
'Set K.GroupRef = Me.GroupRef
For i = fromS To toS
    If IsObject(Me.item(i)) Then
        Select Case Typename(Me.item(i))
        Case mgrp
            Set obj = Me.item(i)
            CopyGroupObj obj, obj2
            Set K.item(i - fromS) = obj2
            Set obj2 = Nothing
        Case mlmbd
            Dim actualdata1()
            With Me.item(i)
            .CopyTo spare, actualdata1()
            End With
            Set K.item(i - fromS) = spare
            Set spare = Nothing
        Case "RefArray"
            Set K.item(i - fromS) = Me.item(i).Copy()
        Case Else
            Set K.item(i - fromS) = Me.item(i)
        End Select
    Else
        K.item(i - fromS) = Me.item(i)
    End If
Next i
'' new
K.Arr = Me.Arr
K.IHaveGui = Me.IHaveGui
K.IhaveClass = False 'Me.IhaveClass
K.Final = Me.Final
End Sub
Public Sub CopyArraySliceFast(K As tuple, fromS As Long, toS As Long)
'copy this to k
' make one dimemsion only
If K Is Me Then Exit Sub
If K Is Nothing Then Set K = New tuple
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
K.MyTypeToBe = MyTypeToBe
'k.SerialItem v, toS - fromS + 1, 8
K.myarrbase = myarrbase
K.RevOrder = False
K.StartResize
For i = 0 To 0
    K.PushDim toS - fromS + 1, 0
Next i
K.PushEnd
K.RevOrder = RevOrder
Set K.GroupRef = Me.GroupRef
For i = fromS To toS
If IsObject(Me.item(i)) Then
    Set K.item(i - fromS) = Me.item(i)
Else
    K.item(i - fromS) = Me.item(i)
End If
Next i

K.Arr = Me.Arr
K.IHaveGui = Me.IHaveGui
K.IhaveClass = Me.IhaveClass
K.Final = Me.Final
End Sub

Public Function UpperMonoLimit()
UpperMonoLimit = maxitems1
End Function
Public Property Get Count()
    If dnum = 0 Then Count = 0 Else Count = maxitems1 + 1 ' sa
End Property
Public Property Let arrname(aName As String)
    mmname = aName
End Property
Public Property Get arrname() As String
    arrname = mmname
End Property
Public Sub PushDim(nDim As Long, Optional limBase As Variant)
ndnum = 1
nmaxitems = nmaxitems * nDim
If Not IsMissing(limBase) Then
   myarrbase = CLng(limBase)
End If
nmaxitems1 = nmaxitems - 1
End Sub
Public Sub PushEnd(Optional thisarr)
Dim z As Variant
Dim mX As Long
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
    dnum = 1 'ndnum
    Dim isnew As Boolean
    isnew = maxitems = 1
    maxitems = nmaxitems
    maxitems1 = nmaxitems1
    If nmaxitems1 < 0 Then
        If isnew And Not MyTypeToBe = vbVariant Then
            Dim Ar As refArray
            Set Ar = New refArray
            Ar.ResetToType MyTypeToBe, 1
            Set actualData = Ar
        ElseIf myVarType(actualData, vbObject) Then
            Set Ar = actualData
            Ar.RedimForFlat 1
        Else
            ReDim actualData(1)
        End If
        
    Else
        If Not IsMissing(thisarr) Then
                Set actualData = thisarr
        ElseIf isnew And Not MyTypeToBe = vbVariant And Not MyTypeToBe = 200 Then
          '  Dim aR As refArray
            Set Ar = New refArray
            Ar.ResetToType MyTypeToBe, maxitems1
            Set actualData = Ar
        Else
        If myVarType(actualData, vbObject) Then
            Set Ar = actualData
            mX = Ar.Count
            Ar.RedimForFlat maxitems1
            
            If MyTypeToBe = 201 Then
                For i = mX To maxitems
                    actualData(i) = CVar(ZeroBig)
                Next
            End If
        Else
            mX = UBound(actualData)
            
            ReDim Preserve actualData(maxitems1)
            If MyTypeToBe = 200 Then
                For i = mX To maxitems1
                    actualData(i) = nMath2.cxZero
                Next
            End If
        End If
        End If
    End If
    
    mX = maxitems
    i = 0
    If mX = 0 Then GoTo there11
End If
Exit Sub
there11:
dnum = 0
maxitems = 1
maxitems1 = 0
    If myVarType(actualData, vbObject) Then
        Set Ar = actualData
        Ar.RedimForFlat maxitems
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
If cursor = 1 Then
item = maxitems
lim = 0
Else
item = 0: lim = 0
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
If cursor <> 1 Then
item = 0
Else
If dnum > 0 Then
item = maxitems
Else
item = 0
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
maxitems = 0
Else
    If myVarType(actualData, vbObject) Then
        actualData.RedimForFlat cursor
    Else
        ReDim Preserve actualData(cursor) 'As Variant
    End If
    maxitems = cursor
End If
   

End If
ElseIf Command = 9 Then
If dnum = 1 Then
    ReDim Preserve actualData(cursor) 'As Variant
    maxitems = cursor - 1
End If
maxitems1 = maxitems - 1
ElseIf Command = 10 Then
   If myVarType(actualData, vbObject) Then
    actualData.RedimForFlat cursor
    
   Else
    ReDim Preserve actualData(cursor) ' As Variant
    End If
    maxitems = cursor
    
    dnum = 1
    maxitems1 = maxitems - 1

ElseIf Command = 11 Then   ' new from 9.4  rev 4
    item = myarrbase
ElseIf Command = 12 Then   ' new from 9.4  rev 4
    myarrbase = cursor
End If
SerialItem = True

End Function
Friend Sub SupportRefArray(Ar As refArray)
Set Ar = Ar.Copy()
If Ar.Count = 1 Then
    MyTypeToBe = Ar.vtType(0)
    PushDim Ar.Count(0)
    Ar.Value = Ar.Value(0)
    Ar.flat = True
    
    PushEnd Ar
   
Else
    MyEr "RefArray has more dimensions than one", "œ RefArray ›˜ÂÈ ÂÒÈÛÛ¸ÙÂÒÂÚ ‰È·ÛÙ‹ÛÂÈÚ ·¸ ÏÈ·"
    
End If
End Sub
Public Function PushOffset(curitem As Long, curdim As Long, nDim As Long) As Boolean
If curdim >= dnum Then
' error...
PushOffset = False
Else


    
  '  If nDim >= limit(curdim) - limBase(curdim) Then Exit Function
  '  curitem = curitem + dtable(curdim) * (nDim + limBase(curdim))
    
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
actualData(curitem) = CVar(thisvariant)
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
            If usehandler.T1 = 4 Then
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
If usehandler.T1 <> 1 Then Err.Raise 1010, , "Need an Inventory Object"
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
Dim Ar As refArray
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
If usehandler.T1 <> 1 Then
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
If Left$(ItemType, 2) = "cx" Then ItemType = Mid$(ItemType, 3)
End Property
Public Property Get ItemTypeNum(curitem) As Integer
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
If usehandler.T1 <> 1 Then
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
If usehandler.T1 <> 1 Then
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
Private Function TestItemBigLess(A As ValueCarrier, b As ValueCarrier) As Boolean
Dim BI As BigInteger
On Error GoTo there
If A.typ = b.typ Then
    If A.typ = 0 Then
        If A.val < b.val Then TestItemBigLess = True
    Else
        TestItemBigLess = A.obj.compare(b.obj) = -1
    End If
ElseIf A.typ = 0 Then
    If VarType(A.val) = vbString Then
        TestItemBigLess = False
    Else
        Set BI = Module13.CreateBigInteger(Format$(Int(A.val), "0"))
        TestItemBigLess = BI.compare(b.obj) = -1
    End If
Else
    If VarType(b.val) = vbString Then
        TestItemBigLess = True
    Else
        Set BI = Module13.CreateBigInteger(Format$(Int(b.val), "0"))
        TestItemBigLess = A.obj.compare(BI) = -1
    End If
End If
there:

End Function
Private Function TestItemBigMore(A As ValueCarrier, b As ValueCarrier) As Boolean
On Error GoTo there
Dim BI As BigInteger
If A.typ = b.typ Then
    If A.typ = 0 Then
        If A.val > b.val Then TestItemBigMore = True
    Else
        TestItemBigMore = A.obj.compare(b.obj) = 1
    End If
ElseIf A.typ = 0 Then
    If VarType(A.val) = vbString Then
        TestItemBigMore = True
    Else
        Set BI = Module13.CreateBigInteger(Format$(Int(A.val), "0"))
        TestItemBigMore = BI.compare(b.obj) = 1
    End If
Else
    If VarType(b.val) = vbString Then
        TestItemBigMore = False
    Else
        Set BI = Module13.CreateBigInteger(Format$(Int(b.val), "0"))
        TestItemBigMore = A.obj.compare(BI) = 1
    End If
End If
there:
End Function

Private Property Get itemBig(curitem As Long) As ValueCarrier
On Error GoTo there
Dim p
If Arr Then
    If maxitems1 >= curitem Then
        If ItemIsObject(curitem) Then
            If actualData(curitem) Is Nothing Then
                itemBig.val = ""
            ElseIf TypeOf actualData(curitem) Is PropReference Then
                SwapVariant itemBig.val, actualData(curitem).Value
                If MemInt(VarPtr(itemBig.val)) = vbObject Then
                   itemBig.val = ""
                End If
            ElseIf TypeOf actualData(curitem) Is BigInteger Then
                Set itemBig.obj = actualData(curitem)
                itemBig.typ = 1
            End If
        Else
            If myVarType(actualData, vbObject) Then
                p = actualData(curitem)
                Select Case VarType(p)
                Case vbObject
                If Not p Is Nothing Then
                    If TypeOf p Is BigInteger Then
                        Set itemBig.obj = p
                        itemBig.typ = 1
                    ElseIf TypeOf p Is PropReference Then
                        SwapVariant p, actualData(curitem).Value
                        If MemInt(VarPtr(itemBig.val)) = vbObject Then
                           itemBig.val = ""
                        End If
                    Else
                        itemBig.val = ""
                    End If
                Else
                    itemBig.val = ""
                End If
                Case 20
                    CopyMemory ByVal VarPtr(itemBig.val) + 8, ByVal VarPtr(p) + 8, 8
                    PutMem2 VarPtr(itemBig.val), 20
                Case 36
                    itemBig.val = ""
                Case Else
                    itemBig.val = p
                End Select
            Else
                Select Case VarType(actualData(curitem))
                Case 20
                    CopyMemory ByVal VarPtr(itemBig.val) + 8, ByVal VarPtr(actualData(curitem)) + 8, 8
                    PutMem2 VarPtr(itemBig.val), 20
                Case 36
                    itemBig.val = ""
                Case Else
                    itemBig.val = actualData(curitem)
                End Select
            End If
        End If
    End If
End If
there:
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
        If usehandler.T1 <> 1 Then
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
                If usehandler.T1 = 4 Then
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
Dim usehandler As mHandler, Ar As refArray
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
    Set Ar = actualData
    Ar(curitem) = item
    Select Case Ar.AssignError
    Case 0
    Case 6
        OverflowValue Ar.vtType(curitem)
    Case 5020
        missNumber
    Case 5030
        MissString
    Case 5040
        MyEr "Missing date or string as date", "ÀÂﬂÂÈ ÁÏÂÒÔÏÁÌﬂ· ﬁ ·Îˆ·ÒÈËÏÁÙÈÍ¸ ˘Ú ÁÏÂÒÔÏÁÌﬂ·"
    Case 5050
        MyEr "not a biginteger or a number", "ƒÂÌ ‚ÒﬁÍ· ÏÂ„‹ÎÔ ·Í›Ò·ÈÔ ﬁ ‹ÎÎÔ ·ÒÈËÏ¸"
    Case Else ' 91
        MyEr "no obect to assign", "ƒÂÌ ‚ÒﬁÍ· ·ÌÙÈÍÂﬂÏÂÌÔ Ì· Í·Ù·˜˘ÒﬁÛ˘"
    End Select
ElseIf Not MyIsObject(actualData(curitem)) Then

If myVarType(item, 20) Then
actualData(curitem) = Empty
SwapVariant actualData(curitem), item
ElseIf MyTypeToBe = 200 Then
If TypeOf item Is Complex Then
    SwapVariant actualData(curitem), item
Else
    actualData(curitem) = nMath2.cxNew(CDbl(item), 0)
    If Err.Number > 0 Then
        WrongType
        Exit Property
    End If
End If
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
        If usehandler.T1 = 4 Then
        ' just throw the enum variable and place the numeber.
            actualData(curitem) = item
        End If
    ElseIf TypeOf actualData(curitem) Is BigInteger Then
        actualData(curitem) = item
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
If usehandler.T1 <> 1 Then Err.Raise 1010, , "Need an Inventory Object"
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
If MyTypeToBe = 200 Then WrongType: Exit Property
On Error Resume Next
If curitem < 0 Then
If Not GroupRef Is Nothing Then
    If TypeOf GroupRef Is mHandler Then
        Dim usehandler As mHandler
        Set usehandler = GroupRef
        If usehandler.T1 = 1 Then
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
MyEr "Wrong setting", "À‹ËÔÚ ¸ÒÈÛÏ·"
ElseIf curitem <= maxitems1 Then
If myVarType(actualData, vbObject) Then
actualData(curitem) = item
Else
Set actualData(curitem) = item
End If
Else

If grouperror Then
MyEr "Something is wrong with this group", " ‹ÙÈ ‹ÂÈ ÛÙÒ·‚‹ ÏÂ ·ıÙﬁ ÙÁÌ ÔÏ‹‰·"
Else
MyEr "Something is wrong with this array", " ‹ÙÈ ‹ÂÈ ÛÙÒ·‚‹ ÏÂ ·ıÙ¸Ì ÙÔÌ ﬂÌ·Í·"
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
ElseIf myVarType(item, vbEmpty) Then
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
Dim myTuple As tuple, mySecondTuple As tuple
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
            Select Case s$
            Case mgrp
                Set subgroup = vvl
                CopyGroupObj subgroup, sub2
                Set subgroup = Nothing
                Set vvl = sub2
                Set sub2 = Nothing
            Case mArr
                If Not vvl.common Then
                    Set myArray = vvl
                    Set mySecondArray = New mArray
                    myArray.CopyArray mySecondArray
                    Set myArray = Nothing
                    Set vvl = mySecondArray
                    Set mySecondArray = Nothing
                End If
            Case mTuple
                If Not vvl.common Then
                    Set myTuple = vvl
                    Set mySecondTuple = New tuple
                    myTuple.CopyArray mySecondTuple
                    Set myTuple = Nothing
                    Set vvl = mySecondTuple
                    Set mySecondTuple = Nothing
                End If
            End Select
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
If usehandler.T1 <> 1 Then
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
    IsEnum = usehandler.T1 = 4
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
Function IsEnum2(curitem As Long, p As Variant) As Boolean
Dim usehandler As mHandler
If curitem >= 0 And curitem <= maxitems1 Then
If IsObject(actualData(curitem)) Then
If Typename(actualData(curitem)) = mHdlr Then
    Set usehandler = actualData(curitem)
    IsEnum2 = usehandler.T1 = 4
     If IsEnum2 Then
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
MyIsNumeric = (n = 11) Or (n = 17) Or (n = 14) Or (n = 20) Or (n = 36)
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
        If TestItemBigLess(itemBig(j), itemBig(j - 1)) Then
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
    Dim i As Long, P1
    Dim j As Long
    i = iMin + 1
    iMax = iMax + 1
    While i < iMax
        j = i
        Do While j > iMin
        If TestItemBigLess(itemBig(j), itemBig(j - 1)) Then
            P1 = Empty
            SwapVariant P1, actualData(j - 1)
            item(j - 1) = CVar(actualData(j))
            item(j) = P1
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        i = i + 1
    Wend
End Sub
Private Sub InsertionSort(ByVal iMin As Long, ByVal iMax As Long)
    Dim i As Long
    Dim j As Long
    i = iMin + 1
    iMax = iMax + 1
    While i < iMax
        j = i
        Do While j > iMin
        If TestItemBigMore(itemBig(j), itemBig(j - 1)) Then
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
    Dim i As Long, P1
    Dim j As Long
    i = iMin + 1
    iMax = iMax + 1
    While i < iMax
        j = i
        Do While j > iMin
        If TestItemBigMore(itemBig(j), itemBig(j - 1)) Then
             P1 = Empty
             SwapVariant P1, actualData(j - 1)
             item(j - 1) = CVar(actualData(j))
             item(j) = P1
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        i = i + 1
    Wend
End Sub

Private Sub QuickSortItemsWithSameKey(ByVal LB As Long, ByVal UB As Long)
    Dim M1 As Long, M2 As Long, Piv As ValueCarrier
    If UB - LB < 7 Then
        If UB - LB = 1 Then
            M1 = LB
            If TestItemBigMore(itemBig(M1), itemBig(UB)) Then SwapVariant actualData(M1), actualData(UB)
            Exit Sub
        Else
            InsertionSort1 LB, UB
            Exit Sub
        End If
    Else
        M1 = (LB + UB) \ 2
        Piv = itemBig(M1)
        M1 = LB
        Do While TestItemBigLess(itemBig(M1), Piv): M1 = M1 + 1: Loop
    End If
    M2 = UB
    Do
        Do While TestItemBigMore(itemBig(M2), Piv): M2 = M2 - 1: Loop
        If M1 <= M2 Then
            If M1 <> M2 Then SwapVariant actualData(M1), actualData(M2)
            M1 = M1 + 1
            M2 = M2 - 1
        End If
        If M1 > M2 Then Exit Do
        Do While TestItemBigLess(itemBig(M1), Piv): M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortItemsWithSameKey LB, M2
    If M1 < UB Then QuickSortItemsWithSameKey M1, UB
End Sub
Private Sub QuickSortItemsWithSameKeyO(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long, P1, Piv As ValueCarrier
    If UB - LB < 7 Then
        If UB - LB = 1 Then
            M1 = LB
            If TestItemBigMore(itemBig(M1), itemBig(UB)) Then
                 P1 = Empty
                 SwapVariant P1, actualData(M1)
                 item(M1) = CVar(actualData(UB))
                 
                 item(UB) = P1
            End If
            Exit Sub
        Else
            InsertionSort1NSW LB, UB
            Exit Sub
        End If
    Else
        M1 = (LB + UB) \ 2
        Piv = itemBig(M1)
     
        M1 = LB
        Do While TestItemBigLess(itemBig(M1), Piv): M1 = M1 + 1: Loop
    End If
    
    M2 = UB
    Do
        Do While TestItemBigMore(itemBig(M2), Piv): M2 = M2 - 1: Loop
        If M1 <= M2 Then
            If M1 <> M2 Then
                P1 = Empty
                SwapVariant P1, actualData(M1)
                item(M1) = CVar(actualData(M2))
                item(M2) = P1
            End If
            M1 = M1 + 1
            M2 = M2 - 1
        End If
        If M1 > M2 Then Exit Do
        Do While TestItemBigLess(itemBig(M1), Piv): M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortItemsWithSameKeyO LB, M2
    If M1 < UB Then QuickSortItemsWithSameKeyO M1, UB
End Sub

Private Sub QuickSortItemsRevWithSameKey(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long, Piv As ValueCarrier
    If UB - LB < 7 Then
        If UB - LB = 1 Then
              M1 = LB
              If TestItemBigMore(itemBig(M1), itemBig(UB)) Then SwapVariant actualData(M1), actualData(UB)
              Exit Sub
        Else
              InsertionSort LB, UB
              Exit Sub
        End If
    Else
        M1 = (LB + UB) \ 2
        Piv = itemBig(M1)
        M1 = LB
        Do While TestItemBigMore(itemBig(M1), Piv): M1 = M1 + 1: Loop
  End If
  M2 = UB
  Do
        Do While TestItemBigLess(itemBig(M2), Piv): M2 = M2 - 1: Loop
        If M1 <= M2 Then
            If M1 <> M2 Then SwapVariant actualData(M1), actualData(M2)
            M1 = M1 + 1
            M2 = M2 - 1
        End If
        If M1 > M2 Then Exit Do
        Do While TestItemBigMore(itemBig(M1), Piv): M1 = M1 + 1: Loop
  Loop
  If LB < M2 Then QuickSortItemsRevWithSameKey LB, M2
  If M1 < UB Then QuickSortItemsRevWithSameKey M1, UB
End Sub
Private Sub QuickSortItemsRevWithSameKeyO(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long, P1, Piv As ValueCarrier
    If UB - LB < 7 Then
        If UB - LB = 1 Then
            M1 = LB
            If TestItemBigLess(itemBig(M1), itemBig(UB)) Then
                P1 = Empty
                SwapVariant P1, actualData(M1)
                item(M1) = CVar(actualData(UB)) ' using cvar we handle objects too
                item(UB) = P1
            End If
            Exit Sub
        Else
            InsertionSortNSW LB, UB
            Exit Sub
        End If
    Else
        M1 = (LB + UB) \ 2
        Piv = itemBig(M1)
        M1 = LB
        Do While TestItemBigMore(itemBig(M1), Piv): M1 = M1 + 1: Loop
    End If
    M2 = UB
    Do
        Do While TestItemBigLess(itemBig(M2), Piv): M2 = M2 - 1: Loop
        If M1 <= M2 Then
            If M1 <> M2 Then
                P1 = Empty
                SwapVariant P1, actualData(M1)
                item(M1) = CVar(actualData(M2))
                item(M2) = P1
            End If
            M1 = M1 + 1
            M2 = M2 - 1
        End If
        If M1 > M2 Then Exit Do
        Do While TestItemBigMore(itemBig(M1), Piv): M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortItemsRevWithSameKeyO LB, M2
    If M1 < UB Then QuickSortItemsRevWithSameKeyO M1, UB
End Sub
Public Sub SortTuple(Optional Low As Long = -1, Optional high As Long = -1)
Dim i As Long, n As Byte
If dnum < 1 Then Exit Sub
If Count > 0 Then
If high = -1 Then high = Count - 1
If high > Count - 1 Then high = Count - 1
If Low = -1 Then Low = 0
If Low >= high Then Exit Sub

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
If Count > 0 Then
If high = -1 Then high = Count - 1
If high > Count - 1 Then high = Count - 1
If Low = -1 Then Low = 0
If Low >= high Then Exit Sub

If MyIsObject(actualData) Then
QuickSortItemsRevWithSameKeyO Low, high
Else
QuickSortItemsRevWithSameKey Low, high
End If
End If
End Sub

Public Sub SortColumns(ParamArray col0())
    MyEr "not for tuple", "¸˜È „È· ·ıÙ¸Ï·ÙÔ ﬂÌ·Í·"
End Sub
Property Get ArrPtr() As Long  ' 32 bit address
Dim Ar As refArray
If myVarType(actualData, vbObject) Then
    If actualData.Count > 0 Then
        Set Ar = actualData
        ArrPtr = Ar.ArrPtr
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

Private Sub iBoxArray_AppendArray(K As iBoxArray)
Dim z As tuple
If Not K Is Nothing Then If TypeOf K Is tuple Then Set z = K
AppendArray z
Set K = z
End Sub

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

Private Property Get iBoxArray_ArrPtr() As Long
iBoxArray_ArrPtr = ArrPtr()
End Property

Private Function iBoxArray_bareteamgroup() As Object
iBoxArray_bareteamgroup = Null
End Function

Private Property Let iBoxArray_CodeName(RHS As String)
CodeName = RHS
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

Private Sub iBoxArray_CopyArray(K As iBoxArray)
Dim z As tuple
If Not K Is Nothing Then If TypeOf K Is tuple Then Set z = K
CopyArray z
Set K = z
End Sub

Private Sub iBoxArray_CopyArrayRev(K As iBoxArray)
Dim z As tuple
If Not K Is Nothing Then If TypeOf K Is tuple Then Set z = K
CopyArrayRev z
Set K = z
End Sub

Private Sub iBoxArray_CopyArrayRevFast(K As iBoxArray)
Dim z As tuple
If Not K Is Nothing Then If TypeOf K Is tuple Then Set z = K
CopyArrayRevFast z
Set K = z
End Sub

Private Sub iBoxArray_CopyArraySlice(K As iBoxArray, fromS As Long, toS As Long)
Dim z As tuple
If Not K Is Nothing Then If TypeOf K Is tuple Then Set z = K
CopyArraySlice z, fromS, toS
Set K = z
End Sub

Private Sub iBoxArray_CopyArraySliceFast(K As iBoxArray, fromS As Long, toS As Long)
Dim z As tuple
If Not K Is Nothing Then If TypeOf K Is tuple Then Set z = K
iBoxArray_CopyArraySliceFast z, fromS, toS
Set K = z
End Sub

Private Property Get iBoxArray_count() As Long
    iBoxArray_count = Count
End Property

Private Function iBoxArray_EmptyArraySameType() As Object
    Set iBoxArray_EmptyArraySameType = EmptyArraySameType
End Function

Private Property Let iBoxArray_Final(ByVal RHS As Boolean)
    Final = RHS
End Property

Private Property Get iBoxArray_Final() As Boolean
    iBoxArray_Final = Final
End Property

Private Function iBoxArray_GetCopy() As Variant()
    iBoxArray_GetCopy = GetCopy
End Function

Private Sub iBoxArray_GetPorp(aProp As PropReference, curitem As Long)
    GetPorp aProp, curitem
End Sub

Private Function iBoxArray_GetUDTValue(ByVal p As Long, Name As String, Optional ndx As Variant) As Variant
If IsMissing(ndx) Then
    iBoxArray_GetUDTValue = GetUDTValue(p, Name)
Else
    iBoxArray_GetUDTValue = GetUDTValue(p, Name, ndx)
End If
End Function

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

Private Property Let iBoxArray_Index(ByVal RHS As Long)
    Index = RHS
End Property

Private Property Get iBoxArray_Index() As Long
    iBoxArray_Index = Index
End Property

Private Property Get iBoxArray_IsEmpty() As Variant
    iBoxArray_IsEmpty = IsEmpty
End Property

Private Function iBoxArray_IsEnum2(curitem As Long, p As Variant) As Boolean
    iBoxArray_IsEnum2 = IsEnum2(curitem, p)
End Function

Private Function iBoxArray_isObj() As Boolean
    iBoxArray_isObj = IsObj
End Function

Private Function iBoxArray_IsObjAt(curitem As Long, peekvalue As Variant) As Boolean
iBoxArray_IsObjAt = IsObjAt(curitem, peekvalue)
End Function

Private Function iBoxArray_IsStringItem(curitem As Long) As Variant
    iBoxArray_IsStringItem = IsStringItem(curitem)
End Function

Private Property Set iBoxArray_Item(curitem As Long, RHS As Variant)
    Set item(curitem) = RHS
End Property

Private Property Let iBoxArray_Item(curitem As Long, RHS As Variant)
    item(curitem) = RHS
End Property

Private Property Get iBoxArray_Item(curitem As Long) As Variant
SwapVariant iBoxArray_Item, Me.item(curitem)
End Property
Private Function iBoxArray_PlaceValue2UDTArray(p As Long, Name As String, v As Variant, Index As Long) As Boolean
iBoxArray_PlaceValue2UDTArray = PlaceValue2UDTArray(p, Name, v, Index)
End Function
Private Function iBoxArray_PlaceValue2UDT(p As Long, Name As String, v As Variant) As Boolean
iBoxArray_PlaceValue2UDT = PlaceValue2UDT(p, Name, v)
End Function
Private Function iBoxArray_ItemIsObject(curitem As Variant) As Boolean
    iBoxArray_ItemIsObject = ItemIsObject(curitem)
End Function
Private Property Let iBoxArray_ItemStr(curitem As Long, RHS As String)
    ItemStr(curitem) = RHS
End Property
Private Function iBoxArray_itemnumeric(curitem As Long) As Variant
    iBoxArray_itemnumeric = itemnumeric(curitem)
End Function
Private Sub iBoxArray_Neg(v As Long)
    Neg v
End Sub
Private Property Get iBoxArray_itemObject(curitem As Long) As Variant
 IsObjAt curitem, iBoxArray_itemObject
End Property

Private Property Get iBoxArray_itemPtr(curitem As Long) As Long
iBoxArray_itemPtr = itemPtr(curitem)
End Property

Private Property Get iBoxArray_ItemType(curitem As Variant) As String
    iBoxArray_ItemType = ItemType(curitem)
End Property

Private Property Get iBoxArray_ItemTypeNum(curitem As Variant) As Integer
iBoxArray_ItemTypeNum = ItemTypeNum(curitem)
End Property

Private Function iBoxArray_MyIsNumeric(v As Variant) As Boolean
iBoxArray_MyIsNumeric = MyIsNumeric(v)
End Function

Private Function iBoxArray_MyIsObject(v As Variant) As Boolean
    iBoxArray_MyIsObject = MyIsObject(v)
End Function

Private Property Get iBoxArray_MyTypeToBe() As Integer
    iBoxArray_MyTypeToBe = MyTypeToBe
End Property

Private Sub iBoxArray_processAppend(m As Object)
    processAppend m
End Sub

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

Private Property Let iBoxArray_myarrbase(ByVal RHS As Long)
    myarrbase = RHS
End Property

Private Property Get iBoxArray_myarrbase() As Long
    iBoxArray_myarrbase = myarrbase
End Property

Private Function iBoxArray_SerialItem(item As Variant, cursor As Long, Command As Long) As Boolean
    iBoxArray_SerialItem = SerialItem(item, cursor, Command)
End Function

Private Sub iBoxArray_SortDesTuple(Optional Low As Long = -1&, Optional high As Long = -1&)
    SortDesTuple Low, high
End Sub

Private Sub iBoxArray_SortTuple(Optional Low As Long = -1&, Optional high As Long = -1&)
    SortTuple Low, high
End Sub

Private Sub iBoxArray_StartResize()
StartResize
End Sub

Private Sub iBoxArray_SwapItem(curitem As Long, thisvariant As Variant)
SwapItem curitem, thisvariant
End Sub

Private Function iBoxArray_UpperMonoLimit() As Variant
iBoxArray_UpperMonoLimit = UpperMonoLimit
End Function

Private Property Get iBoxArray_value() As Variant
    If IsObj Then
        Set iBoxArray_value = Value
    Else
        iBoxArray_value = Value
    End If
End Property
Private Function iBoxArray_IsEnum(p As Variant) As Boolean
    iBoxArray_IsEnum = IsEnum(p)
End Function
