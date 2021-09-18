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
Dim mmname As String
Dim dnum As Long, ndnum As Long
Dim limit(10) As Long  ' max 9 dimensions
Dim dtable(10) As Long
Dim nlimit(10) As Long  ' max 9 dimensions
Dim limBase(10) As Long
Dim nLimbase(10) As Long

Private actualdata() As Variant
Dim maxitems As Long
Dim maxitems1 As Long
Dim nmaxitems As Long
Dim nmaxitems1 As Long
Public comevents As Boolean
Public myarrbase As Long
Public common As Boolean
Public IHaveClass As Boolean
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
Private Type SAFEArrayPart
   cDims                        As Integer
   fFeatures                    As Integer
   cbElements                   As Long
   cLocks                       As Long
   pvData                       As Long
   End Type
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal addr As Long, retval As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Long)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, retval As Integer)
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal addr As Long, retval As Byte)
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
    lLBound As Long
End Type
' no second sort allowed per array
Private sortmult As Long, sortoff As Long, upart As Long, part As Long, Col(), Piv, piv1()
Public Index As Long, Final As Boolean, refgroup As Object, grouperror As Boolean
Public refArray As Variant
Public Sub Compute0(p As Variant, ss$)
If Final Then CantAssignValue: Exit Sub
Dim v As Long, sp As Variant
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
    If MyIsNumeric(actualdata(v)) Then
        Select Case ss$
            Case "=", "g"
            actualdata(v) = p
            GoTo conthere
            Case "+"
            actualdata(v) = p + actualdata(v)
            Case "*"
             actualdata(v) = p * actualdata(v)
            Case "-"
            actualdata(v) = actualdata(v) - p
            Case "/"
            If p = 0 Then Exit Do
             actualdata(v) = actualdata(v) / p
            
              Case "~", "!!"
             actualdata(v) = -1 - (actualdata(v) <> 0)
             Case "++"
            actualdata(v) = 1 + actualdata(v)
            Case "--"
            actualdata(v) = actualdata(v) - 1
              Case "-!"
            actualdata(v) = -actualdata(v)
            GoTo conthere
            Case "+="
            actualdata(v) = p + actualdata(v)
        
            Case "*="
             actualdata(v) = p * actualdata(v)
        
            Case "-="
            actualdata(v) = actualdata(v) - p
            
            Case "/="
                If p = 0 Then DevZero: Exit Do
                actualdata(v) = actualdata(v) / p
                                
            Case "DIV", "ÄÉÁ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                actualdata(v) = Fix(actualdata(v) / p)
            Case "DIV#", "ÄÉÁ#"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                If p < 0 Then
                    actualdata(v) = Int((actualdata(v) - Abs(actualdata(v) - Abs(p) * Int(actualdata(v) / Abs(p)))) / p)
                Else
                    actualdata(v) = Int(actualdata(v) / p)
                End If
                 If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
            Case "MOD", "ÕÐÏË", "ÕÐÏËÏÉÐÏ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = actualdata(v) - Fix(actualdata(v) / p) * p
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                actualdata(v) = sp
             Case "MOD#", "ÕÐÏË#", "ÕÐÏËÏÉÐÏ#"
                 If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = Abs(actualdata(v) - Abs(p) * Int(actualdata(v) / Abs(p)))
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                actualdata(v) = sp
            Case Else
                WrongOperator
                Exit Do
            End Select
        End If
        If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
conthere:
        v = v + 1
        Err.Clear
Loop

End Sub
Public Sub Compute(p As Double, ss$)
If Final Then CantAssignValue: Exit Sub
Dim v As Long
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
    If MyIsNumeric(actualdata(v)) Then
        Select Case ss$
        Case "=", "g"
            actualdata(v) = p
            GoTo conthere
        Case "+"
            actualdata(v) = p + actualdata(v)
        Case "*"
            actualdata(v) = p * actualdata(v)
        Case "-"
        actualdata(v) = actualdata(v) - p
        If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
        Case "/"
            If p = 0 Then Exit Do
            actualdata(v) = actualdata(v) / p
        Case "~", "!!"
            actualdata(v) = -1 - (actualdata(v) <> 0)
            GoTo conthere
        Case "++"
            actualdata(v) = 1 + actualdata(v)
            GoTo conthere
        Case "--"
            actualdata(v) = actualdata(v) - 1
            GoTo conthere
        Case "-!"
            actualdata(v) = -actualdata(v)
            GoTo conthere
        End Select
    End If
    If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
conthere:
    v = v + 1
    Err.Clear
Loop

End Sub
Public Function CompareItem(v As Long, p As Variant, ss$, ret As Boolean) As Boolean
ret = False
On Error GoTo there
If MyIsNumeric(actualdata(v)) Then
Select Case ss$
Case ">"
ret = actualdata(v) > p
Case ">="
ret = actualdata(v) >= p
Case "<"
ret = actualdata(v) < p
Case "<="
ret = actualdata(v) <= p
Case "<>"
ret = actualdata(v) <> p
Case "="
ret = actualdata(v) = p
End Select

Else
Select Case UCase(ss$)
Case ">"
ret = actualdata(v) > p
Case ">="
ret = actualdata(v) >= p
Case "<"
ret = actualdata(v) < p
Case "<="
ret = actualdata(v) <= p
Case "<>"
ret = actualdata(v) <> p
Case "="
ret = actualdata(v) = p
Case "LIKE"
ret = actualdata(v) Like Replace$(CStr(p), "%", "*")

End Select
End If
CompareItem = True
there:
End Function
Public Sub Compute2(p As Variant, ss$)
If Final Then CantAssignValue: Exit Sub
Dim v As Long, sp As Variant
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
    If MyIsNumeric(actualdata(v)) Then
        Select Case ss$
            Case "="
                actualdata(v) = p
                GoTo conthere
            Case "+="
                actualdata(v) = actualdata(v) + p
            Case "*="
                actualdata(v) = p * actualdata(v)
            Case "-="
                actualdata(v) = actualdata(v) - p
            Case "/="
                If p = 0 Then DevZero: Exit Do
                 actualdata(v) = actualdata(v) / p
            Case "DIV", "ÄÉÁ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                actualdata(v) = Fix(actualdata(v) / p)
            Case "DIV#", "ÄÉÁ#"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                If p < 0 Then
                    actualdata(v) = Int((actualdata(v) - Abs(actualdata(v) - Abs(p) * Int(actualdata(v) / Abs(p)))) / p)
                Else
                    actualdata(v) = Int(actualdata(v) / p)
                End If
                 If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
            Case "MOD", "ÕÐÏË", "ÕÐÏËÏÉÐÏ"
                If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = actualdata(v) - Fix(actualdata(v) / p) * p
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                actualdata(v) = sp
             Case "MOD#", "ÕÐÏË#", "ÕÐÏËÏÉÐÏ#"
                 If p = 0 Then
                    DevZero
                    Exit Do
                End If
                sp = Abs(actualdata(v) - Abs(p) * Int(actualdata(v) / Abs(p)))
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                actualdata(v) = sp
            Case Else
                WrongOperator
                Exit Do
         End Select
    End If
conthere:
    If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
    v = v + 1
    Err.Clear
Loop
End Sub
Public Sub Compute3(ss$)
Dim v As Long
    If Final Then CantAssignValue: Exit Sub
    On Error Resume Next
    v = 0
    Do While v <= Me.UpperMonoLimit
        If MyIsNumeric(actualdata(v)) Then
        Select Case ss$
        Case "~"
            actualdata(v) = -1 - (actualdata(v) <> 0)
        Case "++"
            actualdata(v) = 1 + actualdata(v)
        Case "--"
            actualdata(v) = actualdata(v) - 1
        Case "-!"
            actualdata(v) = -actualdata(v)
        End Select
    End If
    v = v + 1
    Err.Clear
Loop

End Sub
Property Get Value() As Variant
If Index >= 0 And Index <= maxitems1 Then

If IsObject(actualdata(Index)) Then
    If Typename(actualdata(Index)) = "PropReference" Then
        Value = actualdata(Index)
    Else
        Set Value = actualdata(Index)
    End If
Else
Value = actualdata(Index)
End If
End If
End Property
Public Sub PushProp(a As Variant)
Set GroupRef = a
Arr = False
End Sub
Friend Sub ExportArrayNow()
CopyVariantArray refArray, actualdata()
End Sub
Friend Sub FixArray()
Dim zz As Variant
CopyMemory ByVal VarPtr(refArray), zz, 16
Dim z As Long, i As Long
If count <> UBound(actualdata()) - LBound(actualdata) + 1 Then
i = count
maxitems1 = UBound(actualdata()) - LBound(actualdata)
maxitems = maxitems1 + 1
z = 1
If dnum > 1 Then

  z = i \ limit(dnum - 1)

   If maxitems Mod z <> 0 Then
  maxitems = maxitems + z - maxitems Mod z
  maxitems1 = maxitems - 1
  ReDim Preserve actualdata(maxitems1)
  End If
  limit(dnum - 1) = maxitems \ z
Else
limit(0) = maxitems
End If
End If
'
End Sub
Friend Sub LoadTuple(ar As Variant)
Dim z As Long, i As Long
actualdata = ar
If count <> UBound(actualdata()) - LBound(actualdata) + 1 Then
i = count
maxitems1 = UBound(actualdata()) - LBound(actualdata)
maxitems = maxitems1 + 1
z = 1

If dnum > 1 Then

  z = i \ limit(dnum - 1)

   If maxitems Mod z <> 0 Then
  maxitems = maxitems + z - maxitems Mod z
  maxitems1 = maxitems - 1
  ReDim Preserve actualdata(maxitems1)
  End If
  limit(dnum - 1) = maxitems \ z
Else
dnum = 1
limit(0) = maxitems
End If
End If
'
End Sub

Friend Sub CopySerialize(mmm() As Variant)
' may we get an exception here
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, R1 As Long, r2 As Long

GetMem4 VarPtrArray(mmm()), m_base
GetMem4 VarPtrArray(actualdata()), m_base_0
R1 = SafeArrayLock(m_base)
r2 = SafeArrayLock(m_base_0)
R1 = SafeArrayAccessData(ByVal m_base, first12)
r2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
r2 = SafeArrayUnaccessData(ByVal m_base_0)
R1 = SafeArrayUnaccessData(ByVal m_base)
r2 = SafeArrayUnlock(m_base_0)
R1 = SafeArrayUnlock(m_base)

End Sub
Friend Sub CopyBackSerialize(mmm() As Variant)
' may we get an exception here
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, R1 As Long, r2 As Long
If Not RevOrder Then Err.Raise vbObjectError + 12345, "CopyBack", "Failed. is not vb6"
Dim backup()
backup() = actualdata()
GetMem4 VarPtrArray(mmm()), m_base
GetMem4 VarPtrArray(backup()), m_base_0
R1 = SafeArrayLock(m_base)
r2 = SafeArrayLock(m_base_0)
R1 = SafeArrayAccessData(ByVal m_base, first12)
r2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
r2 = SafeArrayUnaccessData(ByVal m_base_0)
R1 = SafeArrayUnaccessData(ByVal m_base)
r2 = SafeArrayUnlock(m_base_0)
R1 = SafeArrayUnlock(m_base)
End Sub
Public Sub Car(m As Object)
Dim k As New mArray, usehandler As mHandler
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
If count > 0 Then
    k.PushDim 1
    k.PushEnd
    k.SerialItem actualdata(0), 0, 1
    If IsObject(Me.item(0)) Then
    Select Case Typename(Me.item(0))
    Case "Group"
    Set obj = Me.item(0)
    CopyGroupObj obj, obj2
    Set k.item(0) = obj2
    Set obj2 = Nothing
    Case "lambda"
    Dim actualdata1()
    With Me.item(0)
    .CopyTo spare, actualdata1()
    End With
    Set k.item(0) = spare
    Set spare = Nothing
    Case Else
        If TypeOf Me.item(0) Is mHandler Then
            Set usehandler = Me.item(0)
            If usehandler.t1 = 3 Then
                Set m = usehandler
                Set k = Nothing
                Exit Sub
            End If
         ElseIf TypeOf Me.item(0) Is mArray Then
            Set usehandler = New mHandler
            usehandler.t1 = 3
            Set usehandler.objref = Me.item(0)
            Set m = usehandler
            Set k = Nothing
            Exit Sub
        End If
        Set k.item(0) = Me.item(0)
    End Select
    Else
        k.item(0) = Me.item(0)
    End If
End If
k.Arr = True
Set m = k
Set k = Nothing
End Sub
Public Sub Cdr(m As Object)
Dim k As New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
If dnum > 0 Then
k.StartResize: k.PushDim Me.count - 1: k.PushEnd
Else
k.Arr = True
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
Set m = k
Exit Sub
End If
k.myarrbase = myarrbase
k.RevOrder = False
For i = 1 To Me.UpperMonoLimit
If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i - 1) = obj2
Set obj2 = Nothing
Case "lambda"
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
'' new
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
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
If Right$(k.arrname, 2) = "%(" Then
For i = 0 To Me.UpperMonoLimit
If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i + fromhere) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata()
With Me.item(i)
.CopyTo spare, actualdata()
End With
Set k.item(i + fromhere) = spare
Set spare = Nothing
Case Else
Set k.item(i + fromhere) = Me.item(i)
End Select
Else
k.item(i + fromhere) = Int(Me.item(i))
End If
Next i
Else
''''''''''****************************
For i = 0 To Me.UpperMonoLimit

If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
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
End If
'' new
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
End Sub



Public Sub CopyArray(k As mArray)
'copy this to k

If k Is Me Then Exit Sub
If k Is Nothing Then Set k = New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
k.SerialItem v, Me.UpperMonoLimit, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For i = 0 To dnum - 1
k.PushDim limit(i), limBase(i)
Next i

k.PushEnd
k.RevOrder = RevOrder
Set k.GroupRef = Me.GroupRef
If Right$(k.arrname, 2) = "%(" Then
For i = 0 To Me.UpperMonoLimit

If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata()
With Me.item(i)
.CopyTo spare, actualdata()
End With
Set k.item(i) = spare
Set spare = Nothing

Case Else
Set k.item(i) = Me.item(i)
End Select
Else
k.item(i) = Int(Me.item(i))
End If
Next i
Else
''''''''''****************************
For i = 0 To Me.UpperMonoLimit

If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata1()
With Me.item(i)
.CopyTo spare, actualdata1()
End With
Set k.item(i) = spare
Set spare = Nothing
Case Else
Set k.item(i) = Me.item(i)
End Select
Else
k.item(i) = Me.item(i)
End If
Next i
End If
'' new

k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
k.Final = Me.Final
End Sub
Public Sub CopyArrayRev(k As mArray)
'copy this to k
If k Is Me Then Exit Sub
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
Dim mm As Long, M1 As Long
k.SerialItem v, Me.UpperMonoLimit, 8
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
If Right$(k.arrname, 2) = "%(" Then
For i = 0 To Me.UpperMonoLimit
M1 = mm - i
If IsObject(Me.item(M1)) Then
Select Case Typename(Me.item(M1))
Case "Group"
Set obj = Me.item(M1)
CopyGroupObj obj, obj2
Set k.item(i) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata()
With Me.item(M1)
.CopyTo spare, actualdata()
End With
Set k.item(i) = spare
Set spare = Nothing

Case Else
Set k.item(i) = Me.item(M1)
End Select
Else
k.item(i) = Int(Me.item(M1))
End If
Next i
Else
''''''''''****************************
For i = 0 To Me.UpperMonoLimit
M1 = mm - i
If IsObject(Me.item(M1)) Then
Select Case Typename(Me.item(M1))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata1()
With Me.item(M1)
.CopyTo spare, actualdata1()
End With
Set k.item(i) = spare
Set spare = Nothing
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
k.IHaveClass = Me.IHaveClass
k.Final = Me.Final
End Sub
Public Sub CopyArrayRevFast(k As mArray)
'copy this to k
If k Is Me Then Exit Sub
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
Dim mm As Long, M1 As Long
k.SerialItem v, Me.UpperMonoLimit, 8
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
For i = 0 To Me.UpperMonoLimit
M1 = mm - i
If IsObject(Me.item(M1)) Then
Set k.item(i) = Me.item(M1)
Else
k.item(i) = Me.item(M1)
End If
Next i
'' new
k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
k.Final = Me.Final
End Sub
Public Sub CopyArraySlice(k As mArray, fromS As Long, toS As Long)
'copy this to k
' make one dimemsion only
If k Is Me Then Exit Sub
If k Is Nothing Then Set k = New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
k.SerialItem v, toS - fromS + 1, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For i = 0 To 0
    k.PushDim toS - fromS + 1, 0
Next i
k.PushEnd
k.RevOrder = RevOrder
Set k.GroupRef = Me.GroupRef
If Right$(k.arrname, 2) = "%(" Then
For i = fromS To toS
If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i - fromS) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata()
With Me.item(i)
.CopyTo spare, actualdata()
End With
Set k.item(i - fromS) = spare
Set spare = Nothing

Case Else
Set k.item(i - fromS) = Me.item(i)
End Select
Else
k.item(i - fromS) = Int(Me.item(i))
End If
Next i
Else
''''''''''****************************
For i = fromS To toS

If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i - fromS) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata1()
With Me.item(i)
.CopyTo spare, actualdata1()
End With
Set k.item(i - fromS) = spare
Set spare = Nothing
Case Else
Set k.item(i - fromS) = Me.item(i)
End Select
Else
k.item(i - fromS) = Me.item(i)
End If
Next i
End If
'' new

k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
k.Final = Me.Final
End Sub
Public Sub CopyArraySliceFast(k As mArray, fromS As Long, toS As Long)
'copy this to k
' make one dimemsion only
If k Is Me Then Exit Sub
If k Is Nothing Then Set k = New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
k.SerialItem v, toS - fromS + 1, 8
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
k.IHaveClass = Me.IHaveClass
k.Final = Me.Final
End Sub

Public Function UpperMonoLimit()
UpperMonoLimit = maxitems1
End Function
Public Property Get count()
If dnum = 0 Then
count = 0
Else
count = maxitems1 + 1 ' sa
End If
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
Public Sub PushEnd()
Dim z As Variant
CopyMemory ByVal VarPtr(refArray), z, 16
On Error GoTo there11
If ndnum = 0 Then
dnum = 0
maxitems = 1
maxitems1 = 0
ReDim actualdata(maxitems1) As Variant
Else
Dim i As Long
For i = 0 To 9
limit(i) = nlimit(i): limBase(i) = nLimbase(i)
Next i

    dnum = ndnum
    maxitems = nmaxitems
    maxitems1 = nmaxitems1

If nmaxitems1 < 0 Then
    ReDim actualdata(1) As Variant
Else
    ReDim Preserve actualdata(maxitems1) As Variant
End If
Dim mx As Long
mx = maxitems
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
    dtable(i) = mx / limit(i)
    mx = dtable(i)
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
ReDim actualdata(maxitems)
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
Set actualdata(cursor) = item
Else
actualdata(cursor) = item
End If
ElseIf Command = 2 Then
If dnum = 0 Then Exit Function
If cursor < 0 Or cursor >= maxitems Then Exit Function
item = actualdata(cursor)
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
ReDim actualdata(0)
CodeName = GetAnyName$
Arr = True
limit(0) = 0
Else
    ReDim Preserve actualdata(cursor) As Variant
    
    maxitems = cursor
End If
    limit(0) = maxitems

End If
ElseIf Command = 9 Then
If dnum = 1 Then
    ReDim Preserve actualdata(cursor) As Variant
    maxitems = cursor - 1
    limit(0) = maxitems
End If
maxitems1 = maxitems - 1
ElseIf Command = 10 Then
   
    ReDim Preserve actualdata(cursor) As Variant
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
itemPtr = VarPtr(actualdata(curitem))
End Property
Friend Sub CopyVariantArray(ByRef a As Variant, ByRef b As Variant)
    CopyMemory ByVal VarPtr(a), ByVal VarPtr(b), 16
End Sub
Friend Sub SwapItem(curitem As Long, thisvariant)
SwapVariant actualdata(curitem), thisvariant
End Sub
Friend Function IsStringItem(curitem As Long) As Variant
Dim usehandler As mHandler
If Arr Then
IsStringItem = Typename(actualdata(curitem)) = "String"
ElseIf curitem < 0 And curitem <> -2 Then
On Error Resume Next
Set usehandler = GroupRef
If Err.Number Then InternalError: Exit Function
On Error GoTo 0
If usehandler.t1 <> 1 Then Err.Raise 1010, , "Need an Inventory Object"
If usehandler.objref.Done Then
    If usehandler.objref.IsObj Then
        
    Else
        IsStringItem = Typename(usehandler.objref.Value) = "String"
    End If
End If
End If
End Function
Public Property Get ItemIsObject(curitem) As Boolean
On Error GoTo there1
If Arr Then
    If maxitems1 >= curitem Then
        If MyIsObject(actualdata(curitem)) Then
            ItemIsObject = Not actualdata(curitem) Is Nothing
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
        ItemType = Typename(actualdata(curitem))
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
            ItemType = Typename(usehandler.objref.Value)
        End If
    Else
        ItemType = Typename(usehandler)
    End If

End If

End If
there1:

End Property
Friend Sub GetPorp(aProp As PropReference, curitem As Long)
If Arr Then
    Set aProp = actualdata(curitem)
Else
    Set aProp = GroupRef
End If
End Sub
Public Property Get item(curitem As Long) As Variant
On Error GoTo there1
If Arr Then
    If maxitems1 >= curitem Then
        If IsObject(actualdata(curitem)) Then
            If Typename(actualdata(curitem)) = "PropReference" Then
                item = actualdata(curitem).Value
            Else
                Set item = actualdata(curitem)
            End If
        Else
            item = actualdata(curitem)
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
    Set itemObject = actualdata(curitem)
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
        If IsObject(actualdata(curitem)) Then
            If Typename(actualdata(curitem)) = "PropReference" Then
                itemnumeric = actualdata(curitem).Value
            ElseIf Typename(actualdata(curitem)) = "mHandler" Then
                Set usehandler = actualdata(curitem)
                If usehandler.t1 = 4 Then
                    If usehandler.sign = -1 Then
                        itemnumeric = CDbl(-usehandler.index_cursor)
                    Else
                        itemnumeric = CDbl(usehandler.index_cursor)
                    End If
                Else
                    itemnumeric = 0
                End If
            Else
                itemnumeric = 0
            End If
        Else
            If VarType(actualdata(curitem)) = 8 Then
                itemnumeric = val(actualdata(curitem))
            Else
                itemnumeric = actualdata(curitem)
            End If
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
Dim usehandler As mHandler
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
If Not MyIsObject(actualdata(curitem)) Then
If VarType(actualdata(curitem)) = vbEmpty Then
actualdata(curitem) = item
Else
CantAssignValue
End If
Else
CantAssignValue
End If
ElseIf Not IsObject(actualdata(curitem)) Then
actualdata(curitem) = item
Else
If Not MyIsObject(item) Then
    If TypeOf actualdata(curitem) Is PropReference Then
        actualdata(curitem).Value = item
    ElseIf TypeOf actualdata(curitem) Is VarItem Then
        actualdata(curitem) = item
    ElseIf TypeOf actualdata(curitem) Is mHandler Then
        Set usehandler = actualdata(curitem)
        If usehandler.t1 = 4 Then
        ' just throw the enum variable and place the numeber.
            actualdata(curitem) = item
        End If
    Else
        Set actualdata(curitem) = Nothing
        actualdata(curitem) = CLng(0)
    End If
Else
    actualdata(curitem).Value = item
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
    If usehandler.objref.IsObj Then
        
    Else
        usehandler.objref.Value = item
    End If
End If
End If
End If
ElseIf Final Then
If Not MyIsObject(actualdata(curitem)) Then
If VarType(actualdata(curitem)) = vbEmpty Then
MoveStringToVariant item, actualdata(curitem)
Else
CantAssignValue
End If
Else
CantAssignValue
End If
ElseIf Not IsObject(actualdata(curitem)) Then
MoveStringToVariant item, actualdata(curitem)
Else
If TypeOf actualdata(curitem) Is PropReference Then
    actualdata(curitem).ValueStr = item
Else
    Set actualdata(curitem) = Nothing
    actualdata(curitem) = CLng(0)
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
Set actualdata(curitem) = item
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
If IHaveClass Then
For i = 0 To maxitems - 1
CopyGroupObj GroupRef, aa, True
aa.ToDelete = False
Set actualdata(i) = aa
Next i
Else
If Typename(item) = "Empty" Then
Dim z As Variant
CopyMemory ByVal VarPtr(refArray), z, 16
ReDim actualdata(maxitems)
''maxitems = 0

Else
For i = 0 To maxitems - 1
actualdata(i) = item
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
Set actualdata(i) = obj
Next i
Set obj = Nothing
End Sub
Public Sub FillLambda(bstack As basetask)
Dim obj As Variant, i As Long
Set obj = bstack.lastobj
    Set bstack.lastobj = Nothing
For i = 0 To maxitems - 1
 CopyLambda obj, bstack
Set actualdata(i) = bstack.lastobj
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
 Dim v As Long, s$, frm$, vvl As Variant, x1 As Long, ss$
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
                    If Typename(vvl) = "Group" Then
                    Set subgroup = vvl
                    CopyGroupObj subgroup, sub2
                    Set subgroup = Nothing
                    Set vvl = sub2
                    Set sub2 = Nothing
                    ElseIf Typename(vvl) = "mArray" Then
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
If IsObject(actualdata(Index)) Then
    IsObj = True
End If
End If
End Function
Function IsEnum(p As Variant) As Boolean
Dim usehandler As mHandler
If Index >= 0 And Index <= maxitems1 Then
If IsObject(actualdata(Index)) Then
If Typename(actualdata(Index)) = "mHandler" Then
    Set usehandler = actualdata(Index)
    IsEnum = usehandler.t1 = 4
    If IsEnum Then p = usehandler.index_cursor * usehandler.sign
End If
End If
End If
End Function
Private Sub Class_Initialize()
dnum = 0 ' no actions
maxitems = 1
maxitems1 = 0
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
ReDim actualdata(0)
CodeName = GetAnyName$
Arr = True
myarrbase = -ArrBase
refArray = 0&
End Sub

Private Sub Class_Terminate()
Dim z As Variant
CopyMemory ByVal VarPtr(refArray), z, 16
Set GroupRef = Nothing

Erase actualdata()
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
MyIsNumeric = (n = 11) Or (n = 17) Or (n = 14)
End Function
Function MyIsObject(v As Variant) As Boolean
Dim n As Byte
GetMem1 VarPtr(v), n
MyIsObject = n = 9
End Function
Function GetCopy() As Variant()
GetCopy = actualdata
End Function
Private Sub InsertionSort1(ByVal iMin As Long, ByVal iMax As Long)
    Dim i As Long
    Dim j As Long
    i = iMin + 1
    iMax = iMax + 1
    While i < iMax
        j = i
        Do While j > iMin
        If actualdata(j) < actualdata(j - 1) Then
            SwapVariant actualdata(j), actualdata(j - 1)
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
        If actualdata(j) > actualdata(j - 1) Then
            SwapVariant actualdata(j), actualdata(j - 1)
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        i = i + 1
    Wend
End Sub
Private Sub QuickSortItemsWithSameKey(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long

  If UB - LB = 1 Then
     M1 = LB

     If actualdata(M1) > actualdata(UB) Then SwapVariant actualdata(M1), actualdata(UB)
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSort1 LB, UB
  Exit Sub
  
  Else
    M1 = (LB + UB) \ 2
   Piv = actualdata(M1)
   
           M1 = LB
   Do While actualdata(M1) > Piv: M1 = M1 + 1: Loop
    End If
  
  M2 = UB
  Do
    Do While actualdata(M2) < Piv: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      If M1 <> M2 Then SwapVariant actualdata(M1), actualdata(M2)
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While actualdata(M1) > Piv: M1 = M1 + 1: Loop
  Loop
  If LB < M2 Then QuickSortItemsWithSameKey LB, M2
  If M1 < UB Then QuickSortItemsWithSameKey M1, UB
End Sub
Private Sub QuickSortItemsRevWithSameKey(ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
  If UB - LB = 1 Then
     M1 = LB

     If actualdata(M1) < actualdata(UB) Then SwapVariant actualdata(M1), actualdata(UB)
     Exit Sub
  ElseIf UB - LB < 7 Then
  InsertionSort LB, UB
  Exit Sub
  
  Else
    M1 = (LB + UB) \ 2
   Piv = actualdata(M1)
   
           M1 = LB
   Do While actualdata(M1) < Piv: M1 = M1 + 1: Loop
  End If
  
  M2 = UB
  Do
    Do While actualdata(M2) > Piv: M2 = M2 - 1: Loop
    If M1 <= M2 Then
      If M1 <> M2 Then SwapVariant actualdata(M1), actualdata(M2)
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While actualdata(M1) < Piv: M1 = M1 + 1: Loop
  Loop
  If LB < M2 Then QuickSortItemsRevWithSameKey LB, M2
  If M1 < UB Then QuickSortItemsRevWithSameKey M1, UB
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
GetMem1 VarPtr(actualdata(i)), n
If (n > 8 Or n < 2) And Not n = 14 Then Exit Sub
Next i
QuickSortItemsWithSameKey Low, high

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
GetMem1 VarPtr(actualdata(i)), n
If (n > 8 Or n < 2) And Not n = 14 Then Exit Sub
Next i
QuickSortItemsRevWithSameKey Low, high
End If
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
        If actualdata(j1) = actualdata(j1 - sortmult) Then
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
        If actualdata(j1) < actualdata(j1 - sortmult) Then
           ' Debug.Print "insertion swap"
            For k = j1 - sortoff To j1 - sortoff + sortmult - 1
            SwapVariant actualdata(k), actualdata(k - sortmult)
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
        If actualdata(j1) > actualdata(j1 - sortmult) Then
           ' Debug.Print "insertion swap"
            For k = j1 - sortoff To j1 - sortoff + sortmult - 1
            SwapVariant actualdata(k), actualdata(k - sortmult)
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
SwapVariant actualdata(a1 + k), actualdata(A2 + k)
Next k
End Sub
Private Sub FeedPiv(ByVal a1)
Static k As Long
On Error Resume Next
'Debug.Print "feed piv "; a1
a1 = a1 * sortmult
For k = 0 To sortmult - 1
    piv1(k) = actualdata(a1 + k)
Next k
End Sub

Private Function Compare10(ByVal a1) As Long
Static k As Long, ret As Long
On Error Resume Next
'Debug.Print "compare ", a1
a1 = a1 * sortmult
For k = 2 To upart - 1 Step 2
ret = actualdata(a1 + Col(k)) = piv1(Col(k))
If ret = 0 Then
    If actualdata(a1 + Col(k)) < piv1(Col(k)) Then
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
Next k
' Compare10=0
End Function
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
If Typename(col0(0)) = "Variant()" Then
    col0() = col0(0)
ElseIf Typename(col0(0)) = "mHandler" Then
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
For i = LBound(col0) To UBound(col0)
Col(i) = CLng(col0(i))
Next i
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
InsertionSort10 Col(0), Col(1)
Else
QuickSortItemsWithSameKey10 Col(0), Col(1)
End If
ReDim piv1(1)
End Sub

