VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mStiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mstack As New Collection
Private mMany As Long
Dim BeginFrom As Long
' interface for iteration
Public Index As Long
' new idea to not include a VarItem
' we place three items, the slot, the type (or subtype), the data.
' so each item is a multiple of three

Public Function IsEmpty() As Boolean
IsEmpty = mstack.Count = 0
End Function
Public Property Get StackItem(where As Long) As Variant
Dim aa As VarItem
On Error Resume Next
Set aa = mstack.item(where)
If aa.slot = 1 Then Set StackItem = aa.mItem Else StackItem = aa.mItem
End Property
Public Sub MakeTopItemBack(where As Long)
'we pick an item from top to any other place
Dim aa As VarItem
On Error Resume Next
If where > 1 Then
Set aa = mstack.item(1)
mstack.Remove 1
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , , where - 1
End If
End If
End Sub
Public Sub MakeTopItem(where As Long)
'we pick an item and to the top
Dim aa As VarItem
On Error Resume Next
If where > 1 Then
Set aa = mstack.item(where)
mstack.Remove where
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
End If
End Sub
Public Sub Copy2TopItem(where As Long)
'we pick an item and to the top
Dim aa As VarItem, bb As VarItem, mm As Object, mn As mArray, work As Object
On Error Resume Next
Set bb = mstack.item(where)
If bb.ObjectRef Is Nothing Then
    If bb.slot = 4 Then bb.slot = 5
    Set aa = bb
Else
    Set work = bb.ObjectRef
    Select Case Typename(work)
        Case "Group"
            Set aa = NewVarItem()
            aa.Typename = "*"
            CopyGroupObj work, mm
            Set aa.ObjectRef = mm
        Case "lambda"
            Set aa = NewVarItem()
            aa.Typename = "*"
            ' use copylambdaAny from module 1
            CopyLambdaAny work, mm
            Set aa.ObjectRef = mm
        Case "mArray"
            Set aa = NewVarItem()
            aa.Typename = "*"
            Set mn = New mArray
            work.CopyArray mn
            Set aa.ObjectRef = mn
        Case Else
            Set aa = mstack.item(where)
        End Select
    Set work = Nothing
End If
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Function CopyMe() As mStiva
Set CopyMe = New mStiva
CopyMe.Copy2TopNItems2FromStiva Me.Total, Me
End Function
Public Function CopyMe2(ByVal many As Long) As mStiva
Set CopyMe2 = New mStiva
If many < 0 Then
many = -many
CopyMe2.Copy2BottomNItems2FromStiva many, Me
Else
CopyMe2.Copy2TopNItems2FromStiva many, Me
End If
End Function
Public Function MergeFromOther(that As Object) As mStiva
Dim mm As mStiva
Set mm = that
Set MergeFromOther = CopyMe()
MergeFromOther.MergeBottom mm
End Function
Public Sub Copy2TopNItems2FromStiva(many As Long, Source As mStiva)
'we pick an item and to the top
Dim aa As VarItem, bb As VarItem, mm As Object, mn As mArray, work As Object
On Error Resume Next
For many = many To 1 Step -1
    If Source.Total >= many Then
        Set bb = Source.StackPickRef2(many)
        If bb.ObjectRef Is Nothing Then
            If bb.slot = 4 Then bb.slot = 5
            Set aa = bb
        Else
            Set work = bb.ObjectRef
            Select Case Typename(work)
            Case "Group"
                Set aa = NewVarItem()
                aa.Typename = "*"
                CopyGroupObj work, mm
                Set aa.ObjectRef = mm
            Case "lambda"
                Set aa = NewVarItem()
                aa.Typename = "*"
                ' use copylambdaAny from module 1
                CopyLambdaAny work, mm
                Set aa.ObjectRef = mm
            Case "mArray"
                Set aa = NewVarItem()
                aa.Typename = "*"
                Set mn = New mArray
                work.CopyArray mn
                Set aa.ObjectRef = mn
            Case Else
            
            
            Set aa = bb
            
            End Select
            Set work = Nothing
        End If
        If mstack.Count = 0 Then
            mstack.Add aa
        Else
            mstack.Add aa, , 1
        End If
        Inc
    Else
        Exit For
    End If
Next many
End Sub
Public Sub Copy2BottomNItems2FromStiva(many1 As Long, Source As mStiva)
' reverse order
Dim aa As VarItem, bb As VarItem, mm As Object, mn As mArray, work As Object, many As Long
On Error Resume Next
For many = Source.Total - many1 + 1 To Source.Total

If Source.Total >= many Then
    Set bb = Source.StackPickRef2(many)
    If bb.ObjectRef Is Nothing Then
        Set aa = bb
    Else
        Set work = bb.ObjectRef
        Select Case Typename(work)
        Case "Group"
            Set aa = NewVarItem()
            aa.Typename = "*"
            CopyGroupObj work, mm
            Set aa.ObjectRef = mm
        Case "lambda"
            Set aa = NewVarItem()
            aa.Typename = "*"
            ' use copylambdaAny from module 1
            CopyLambdaAny work, mm
            Set aa.ObjectRef = mm
        Case "mArray"
            Set aa = NewVarItem()
            aa.Typename = "*"
            Set mn = New mArray
            work.CopyArray mn
            Set aa.ObjectRef = mn
        Case Else
        Set aa = bb
        End Select
        Set work = Nothing
        End If
        If mstack.Count = 0 Then
            mstack.Add aa
        Else
            mstack.Add aa, , 1
        End If
        Inc
    Else
        Exit For
    End If
Next many
End Sub
Public Property Get StackPick(where As Long) As VarItem
Dim a As VarItem
dec
Set a = mstack.item(where)
mstack.Remove where

Set StackPick = a
End Property
Public Property Get StackPickRef(where As Long) As VarItem
Set StackPickRef = mstack.item(where)
End Property
Public Property Get StackPickRef2(where As Long) As VarItem
Set StackPickRef2 = mstack.item(where)
If StackPickRef2.slot = 4 Then StackPickRef2.slot = 5
End Property

Public Property Get StackItemType(where As Long) As String
Dim aa As VarItem
On Error GoTo mess

Set aa = mstack.item(where)
StackItemType = aa.Typename
Exit Property
mess:
MyEr "offset out of limits", "������� ����� �����"
End Property
Public Property Get StackItemTypeiBoxArray(where As Long) As Boolean
Dim aa As VarItem, bb As Object
On Error Resume Next
Set aa = mstack.item(where)
If aa.Typename = "*" Then
Set bb = aa.ObjectRef
If TypeOf bb Is mHandler Then
    StackItemTypeiBoxArray = CheckAnyArray(bb)
ElseIf TypeOf bb Is iBoxArray Then
    StackItemTypeiBoxArray = True
End If
End If
End Property
Public Property Get StackItemTypeObjectType(where As Long) As String
Dim aa As VarItem, bb As Object, mh As mHandler
On Error Resume Next

Set aa = mstack.item(where)
Select Case aa.Typename
Case "N", "S", "?", "L"
StackItemTypeObjectType = vbNullString
Case "*"
Set bb = aa.ObjectRef
If TypeOf bb Is mHandler Then
Set mh = bb
Set bb = mh.objref
CheckDeepAny bb
If Not TypeOf bb Is mHandler Then
StackItemTypeObjectType = "*[" + Typename(mh.objref) + "]"
Else
Set mh = bb
If mh.t1 = 3 Or mh.t1 = 0 Then
StackItemTypeObjectType = "*[" + Typename(mh.objref) + "]"
ElseIf mh.t1 = 1 Then
StackItemTypeObjectType = "*[Inventory]"
ElseIf mh.t1 = 2 Then
StackItemTypeObjectType = "*[Buffer]"
ElseIf mh.t1 = 4 Then
StackItemTypeObjectType = "*[Enum]"
End If
End If
Else
StackItemTypeObjectType = VarTypeName(bb)
End If
Set bb = Nothing
Case Else
Set bb = aa.ObjectRef
StackItemTypeObjectType = VarTypeName(bb)
Set bb = Nothing
End Select
End Property
Public Property Get StackItemTypeIsObject(where As Long) As Boolean
Dim aa As VarItem
On Error Resume Next

Set aa = mstack.item(where)
Select Case aa.Typename
Case "N", "S", "L", "?"
Case Else
StackItemTypeIsObject = True
End Select
End Property
Public Property Get Total() As Long
If mMany < 0 Then mMany = 0
Total = mMany
End Property
Public Property Get Count() As Long
Count = mMany
End Property
Private Sub Inc()
mMany = mMany + 1
End Sub
Private Sub dec()

mMany = mMany - 1
End Sub
Public Sub drop(ByVal items As Long)
Dim num As Long
Dim aa As VarItem
On Error Resume Next
mMany = mMany - items
If mMany < 0 Then mMany = 0: Set mstack = New Collection: Exit Sub
For num = 1 To items
    Set aa = mstack.item(1)
    mstack.Remove 1
    aa.Recycle
    
    
    
Next num

End Sub
Public Sub Flush()
Dim num As Long, aa As VarItem
For num = 1 To mstack.Count
dec
  Set aa = mstack(1)
   mstack.Remove 1
  aa.Recycle
  
 
Next num
mMany = 0
End Sub
Public Sub PushUndefine()
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = "?"
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub

Public Sub PushVal(a As Variant)
Dim aa As VarItem
Set aa = NewVarItem()
If MemInt(VarPtr(a)) = vbString Then
aa.Typename = "S"
aa.ItemDouble = a
Else
aa.Typename = "N"
aa.ItemDouble = a
End If
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Sub PushOptional()
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = ">"
aa.ItemDouble = 0
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Sub PushValLong(a As Variant)
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = "L"
aa.ItemVariant = a
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub

Public Sub PushObj(RHS As Variant)
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = "*"  'something
Set aa.ObjectRef = RHS
Set RHS = Nothing
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Sub PushObjVarItem(RHS As Object)
If mstack.Count = 0 Then
mstack.Add RHS
Else
mstack.Add RHS, , 1
End If
Inc
End Sub
Public Property Get PopObj() As Variant
On Error Resume Next
Dim aa As VarItem
dec
Set aa = mstack.item(1)
Set PopObj = aa.ObjectRef
mstack.Remove 1
aa.Recycle

  '
End Property
Public Property Get objref(where As Long) As Object
On Error Resume Next
Set objref = mstack.item(where).ObjectRef
End Property
Public Sub PushStr(ByRef a As String)
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = "S"
aa.ItemStr = a
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Sub PushStrVariant(ByRef a As Variant)
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = "S"
aa.ItemStrVar = a
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1
End If
Inc
End Sub
Public Function LookNull() As Boolean
On Error Resume Next
If mstack.Count = 0 Then
LookNull = False
Else
LookNull = mstack.item(1).NoValue

End If
End Function
Public Function PopType() As String
On Error Resume Next
If mstack.Count = 0 Then
PopType = vbNullString
Else

PopType = mstack.item(1).Typename
End If
End Function
Public Property Get LookTopVal() As Variant
On Error Resume Next
If mMany > 0 Then

LookTopVal = mstack.item(1).ItemDouble
End If
End Property
Public Property Get LookNVal(n As Long) As Variant
On Error Resume Next
LookNVal = mstack.item(n).ItemVariant
End Property

Public Property Get PopVal() As Variant
On Error Resume Next
Dim aa As VarItem
dec
Set aa = mstack.item(1)
mstack.Remove 1
PopVal = aa.ItemDouble
aa.Recycle

End Property
Public Property Get PopVariant() As Variant
On Error Resume Next
Dim aa As VarItem
dec
Set aa = mstack.item(1)
mstack.Remove 1
PopVariant = aa.ItemVariant
aa.Recycle


End Property
Public Property Get PopAnyNoObject() As Variant
On Error Resume Next
Dim aa As VarItem
dec
Set aa = mstack.item(1)
PopAnyNoObject = aa.ItemNoObject
mstack.Remove 1
aa.Recycle


End Property
Public Property Get PopValLong() As Variant
On Error Resume Next
Dim aa As VarItem
dec
Set aa = mstack.item(1)
mstack.Remove 1
PopValLong = CLng(aa.ItemDouble)
aa.Recycle


End Property
Public Sub PopStr2(s$)
On Error Resume Next
Dim aa As VarItem
dec
Set aa = mstack.item(1)
mstack.Remove 1
aa.PopString s$
'PopStr = aa.ItemStr
aa.Recycle
End Sub
Public Property Get PopStr() As String
On Error Resume Next
Dim aa As VarItem
dec
Set aa = mstack.item(1)
mstack.Remove 1
aa.PopString PopStr
'PopStr = aa.ItemStr
aa.Recycle



End Property
Public Sub DataValLong(a As Variant)
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = "L"
aa.ItemVariant = a
mstack.Add aa
Inc
End Sub
Public Sub DataVal(a As Variant)
Dim aa As VarItem
Set aa = NewVarItem()
If MemInt(VarPtr(a)) = vbString Then
aa.Typename = "S"
aa.ItemDouble = a
Else
aa.Typename = "N"
aa.ItemVariant = a
End If
mstack.Add aa
Inc
End Sub
Public Sub DataOptional()
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = ">"
aa.ItemDouble = 0
mstack.Add aa
Inc
End Sub
Public Sub DataStrVariant(ByRef a As Variant)
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = "S"
aa.ItemStrVar = a
mstack.Add aa
Inc
End Sub

Public Sub DataStr(ByRef a As String)
Dim aa As VarItem
Set aa = NewVarItem()
aa.Typename = "S"
aa.ItemStr = a$
mstack.Add aa
Inc
End Sub
Public Sub DataObj(RHS As Variant)
Dim aa As VarItem
If RHS Is Nothing Then
    Set aa = NewVarItem()
    aa.Typename = "*"  'something
    Set aa.ObjectRef = RHS
ElseIf TypeOf RHS Is VarItem Then
    Set aa = RHS
    Set RHS = Nothing
Else
    Set aa = NewVarItem()
    aa.Typename = "*"  'something
    Set aa.ObjectRef = RHS
    Set RHS = Nothing
End If
mstack.Add aa
Inc
End Sub
Public Sub DataObjVaritem(RHS As Object)
mstack.Add RHS
Inc
End Sub
Public Sub Begin(from As Long)
BeginFrom = from
End Sub
Public Sub AddItem(ByRef a As String)
Dim aa As VarItem
Set aa = NewVarItem()
    aa.Typename = "S"
    aa.ItemStr = a$
    If mstack.Count = 0 Then
        BeginFrom = 1
        mstack.Add aa
    ElseIf BeginFrom > mstack.Count Then
        mstack.Add aa
    Else
        If BeginFrom = 1 Then
            mstack.Add aa, , 1
        Else
         mstack.Add aa, , , BeginFrom - 1
        End If
    End If
        BeginFrom = BeginFrom + 1

Inc
End Sub
Public Sub MergeTop(a As mStiva)
If Me Is a Then Exit Sub
Dim i As Long, aa As VarItem

If a.Total > 0 Then
Do
i = a.Total
Set aa = a.StackPick(i)
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , 1  'push so first of a is first on mstack
End If
Inc
Loop Until i <= 1
End If
End Sub

Public Sub MergeBottomCopyArray(a As iBoxArray)
If a.Count = 0 Then Exit Sub
Dim x1 As Long, vvl As Variant, aa As Object
Dim subgroup As Object, sub2 As Object, myArray As mArray, mySecondArray As mArray
Dim myTuple As tuple, mySecondTuple As tuple
 For x1 = 0 To a.Count - 1
                                        
                    If a.MyIsObject(a.item(x1)) Then
                    Set vvl = a.item(x1)
                    If Typename(vvl) = "Group" Then
                        Set subgroup = vvl
                        CopyGroupObj subgroup, sub2
                        Set subgroup = Nothing
                        Set vvl = sub2
                        Set sub2 = Nothing
                    ElseIf Typename(vvl) = "mArray" Then
                        Set myArray = vvl
                        Set mySecondArray = New mArray
                        myArray.CopyArray mySecondArray
                        Set myArray = Nothing
                        Set vvl = mySecondArray
                        Set mySecondArray = Nothing
                    ElseIf Typename(vvl) = "tuple" Then
                        Set myTuple = vvl
                        Set mySecondTuple = New tuple
                        myTuple.CopyArray mySecondTuple
                        Set myTuple = Nothing
                        Set vvl = mySecondTuple
                        Set mySecondTuple = Nothing
                    ElseIf Typename(vvl) = "VarItem" Then
                            DataOptional
                            GoTo here
                    End If
                    ' DO SOMETHING
                        Set aa = vvl
                        DataObj aa
                        Set vvl = Nothing
                        Set aa = Nothing
                    Else
                        If a.MyIsNumeric(a.item(x1)) Then
                        
                            DataVal a.item(x1)
                        Else
                            DataStr a.item(x1)
                        End If
                     End If
here:
                Next x1


End Sub
Public Sub MergeTopCopyArray(a As iBoxArray)
If a.Count = 0 Then Exit Sub
Dim x1 As Long, vvl As Variant, aa As Object
Dim subgroup As Object, sub2 As Object, myArray As mArray, mySecondArray As mArray
 For x1 = 0 To a.Count - 1
                                        
                    If a.MyIsObject(a.item(x1)) Then
                    Set vvl = a.item(x1)
                    If Typename(vvl) = "Group" Then
                        Set subgroup = vvl
                        CopyGroupObj subgroup, sub2
                        Set subgroup = Nothing
                        Set vvl = sub2
                        Set sub2 = Nothing
                    ElseIf Typename(vvl) = "mArray" Then
                        Set myArray = vvl
                        Set mySecondArray = New mArray
                        myArray.CopyArray mySecondArray
                        Set myArray = Nothing
                        Set vvl = mySecondArray
                        Set mySecondArray = Nothing
                    End If
                    ' DO SOMETHING
                        Set aa = vvl
                        PushObj aa
                        Set vvl = Nothing
                        Set aa = Nothing
                    Else
                        If a.MyIsNumeric(a.item(x1)) Then
                            PushVal a.item(x1)
                        Else
                            PushStr a.item(x1)
                        End If
                     End If
    
                Next x1


End Sub
Function ExportArray(ByVal items As Long) As iBoxArray

If items > Count Then items = Count
Set ExportArray = New tuple
If items = 0 Then Exit Function
Dim i As Long, aa As VarItem
With ExportArray
    .myarrbase = 0
    .PushDim (items)
    .PushEnd
    .arr = True


For i = 1 To items
Set aa = mstack.item(i)
Select Case aa.Typename
Case "S"
.item(i - 1) = aa.ItemStr
Case "N", "?", "L"
.item(i - 1) = aa.ItemDouble
Case Else
Set .item(i - 1) = aa.ObjectRef
End Select
Next i
End With
Set aa = Nothing
drop items

End Function

Public Sub MergeBottom(a As mStiva)
If Me Is a Then Exit Sub
Dim i As Long, aa As VarItem
If a.Total > 0 Then
Do
Set aa = a.StackPick(1)
mstack.Add aa
Inc
Loop Until a.Total < 1
End If
End Sub



Private Sub Class_Terminate()
Flush
End Sub
Sub CopyGroupObj(myobject1 As Object, myobject2 As Object)
Dim s$, frm$, vvl As Variant, x1 As Long
Dim subgroup As Object, sub2 As Object, myArray As mArray, mySecondArray As mArray
Dim mygroup2 As New Group
Dim myobject As Group
Set myobject = myobject1
If myobject.IamApointer Then Set myobject2 = myobject1: Exit Sub
With myobject
    If .IamFloatGroup Then
        .PeekItem 1, vvl
        mygroup2.BeginFloat vvl + 2
         For x1 = 0 To vvl * 2 + 3
            .PeekItem x1, vvl
            s$ = VarTypeName(vvl)
            If s$ = "Group" Then
                Set subgroup = vvl
                CopyGroupObj subgroup, sub2
                Set subgroup = Nothing
                Set vvl = sub2
                Set sub2 = Nothing
            ElseIf s$ = "mArray" Then
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
    End If
    mygroup2.HasStrValue = .HasStrValue
    mygroup2.HasValue = .HasValue
    mygroup2.HasSet = .HasSet
    mygroup2.HasParameters = .HasParameters
    mygroup2.HasParametersSet = .HasParametersSet
    mygroup2.HasRemove = .HasRemove
    Set mygroup2.SuperClassList = .SuperClassList
    Set mygroup2.Events = .Events
    mygroup2.highpriorityoper = .highpriorityoper
    mygroup2.HasUnary = .HasUnary
    mygroup2.ToDelete = .ToDelete
    Set mygroup2.mytypes = .mytypes
End With
Set myobject2 = mygroup2
End Sub
Function Copy() As mStiva
Dim i As Long
Set Copy = New mStiva
For i = Total To 1 Step -1
Copy.PushStr Me.StackItem(i)
Next i
End Function
Property Get Value() As Variant
Dim aa As VarItem
If Index >= 0 And Index < mMany Then
    Set aa = mstack.item(Index + 1)
    If aa.slot = 1 Then
        Set Value = aa.mItem
    Else
        Value = aa.mItem
    End If
End If
End Property
Function IsObj() As Boolean
Dim aa As VarItem
If Index >= 0 And Index < mMany Then
    Set aa = mstack.item(Index + 1)
    If aa.slot = 1 Then
        IsObj = True
    End If
End If
End Function
Function IsObjAt(curitem As Long, peekvalue) As Boolean
Dim aa As VarItem
If Index >= 0 And Index < mMany Then
    Set aa = mstack.item(Index + 1)
    If aa.slot = 1 Then
        Set peekvalue = aa.ObjectRef
        IsObjAt = True
    End If
End If

End Function
Function IsEnum(p As Variant) As Boolean
Dim aa As VarItem, usehandler As mHandler
If Index >= 0 And Index < mMany Then
    Set aa = mstack.item(Index + 1)
    If aa.slot = 1 Then
        If Typename(aa.mItem) = "mHandler" Then
            Set usehandler = aa.mItem
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
Sub SplitStackTop(fromstiva As mHandler, item As Long)
Dim aa As VarItem, i As Long
Dim thisstiva As mStiva
Set thisstiva = fromstiva.objref
If item > thisstiva.Count Then item = thisstiva.Count
For i = 1 To item
    Set aa = thisstiva.StackPick(1)
        mstack.Add aa
        Inc
Next i
End Sub
Sub SplitStackBottom(fromstiva As mHandler, item As Long)
Dim aa As VarItem, i As Long
Dim thisstiva As mStiva
Set thisstiva = fromstiva.objref
If item < 1 Then item = 1
For i = item To thisstiva.Count
    Set aa = thisstiva.StackPick(item)
        mstack.Add aa
        Inc
Next i
End Sub


