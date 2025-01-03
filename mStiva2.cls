VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mStiva2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, retval As Byte)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, retval As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
' a stack only for strings for group objects
Private mstack As New Collection
Public Property Get Total() As Long
Total = mstack.count
End Property
Public Property Get count() As Long
count = mstack.count
End Property
Function Copy() As mStiva2
Dim I As Long
Set Copy = New mStiva2
For I = Total To 1 Step -1
Copy.PushVariant mstack.item(I)
Next I
End Function
Public Property Get IsDecimal() As Boolean
Dim what, t As Byte
If mstack.count > 0 Then
what = mstack.item(1)
GetMem1 VarPtr(what), t
IsDecimal = t = 14
End If
End Property
Public Property Get IsCurrency() As Boolean
Dim what, t As Byte
If mstack.count > 0 Then
what = mstack.item(1)
GetMem1 VarPtr(what), t
IsCurrency = t = 14
End If
End Property
Public Property Get LookTopVal() As Variant
If mstack.count > 0 Then
LookTopVal = mstack.item(1)
End If
End Property
Public Function LookTopValIsLong(p As Long) As Boolean
Dim A, t As Byte
If mstack.count > 0 Then
A = mstack.item(1)
GetMem1 VarPtr(A), t
If t = 3 Then
LookTopValIsLong = True
mstack.Remove 1
p = CLng(A)
End If
End If

End Function
Public Property Get LookSecondVal() As Variant
If mstack.count > 1 Then
LookSecondVal = mstack.item(2)
End If
End Property
Public Property Get PopVal() As Variant
If mstack.count > 0 Then
PopVal = mstack.item(1)
   mstack.Remove 1
End If
End Property
Public Sub PushVal(A As Variant)
If mstack.count = 0 Then
mstack.Add A
Else
mstack.Add A, , 1
End If
End Sub
Public Sub PushVariant(A As Variant)
If mstack.count = 0 Then
mstack.Add A
Else
mstack.Add A, , 1
End If
End Sub
Public Sub PushLong(A As Long)
If mstack.count = 0 Then
mstack.Add CVar(A)
Else
mstack.Add CVar(A), , 1
End If
End Sub
Public Sub MakeTopItem(where As Long)
'we pick an item and to the top
Dim A As Variant
On Error Resume Next
If where > 1 Then
A = mstack.item(where)
mstack.Remove where
If mstack.count = 0 Then
mstack.Add A
Else
mstack.Add A, , 1
End If
End If
End Sub
Public Sub Flush()
Dim num As Long
For num = 1 To mstack.count
   mstack.Remove 1
Next num
End Sub
Public Sub PushStr(ByRef A As String)
If mstack.count = 0 Then
mstack.Add CVar(A)
Else
mstack.Add CVar(A), , 1
End If
End Sub
Public Property Get StackItem2(where As Long, maxvalue As Long) As Variant
If where > mstack.count Then
StackItem2 = maxvalue + 1
Else
StackItem2 = mstack.item(where)
End If
End Property
Public Property Get StackItem(where As Long) As Variant
StackItem = mstack.item(where)
End Property
Public Property Get PopObj() As Object
On Error Resume Next
Set PopObj = mstack.item(1)
mstack.Remove 1

End Property
Public Sub DataObj(RHS As Object)
Dim aa
Set aa = RHS
mstack.Add aa
End Sub
Public Sub DataVal(A As Variant)
mstack.Add A
End Sub
Public Sub DataStr(ByRef A As String)
Dim v As Variant
v = vbNullString
SwapString2Variant A, v
mstack.Add v
End Sub
Sub GetStackPick(where, ByRef v As Variant)
SwapVariant v, mstack.item(where)
mstack.Remove where
End Sub

Public Property Get StackPick(where As Long) As Variant
StackPick = mstack.item(where)
mstack.Remove where
End Property
Public Sub MergeBottom(A As mStiva2)
If Me Is A Then Exit Sub
Dim I As Long, aa As Variant
If A.Total > 0 Then
Do
aa = A.StackPick(1)
mstack.Add aa
Loop Until A.Total < 1
End If
End Sub
Public Sub DataStrUn(A As String)
Dim aa As Variant, I As Long, b As String
b = Split(A)(0)
aa = A
If mstack.count = 0 Then
mstack.Add aa
Else
For I = 1 To Total
If Split(mstack.item(I))(0) = b Then

mstack.Remove I
If mstack.count = 0 Then
mstack.Add aa
ElseIf I = 1 Then
mstack.Add aa, , 1
Else
mstack.Add aa, , , I - 1
End If
Exit Sub
End If
Next I
mstack.Add aa
End If
End Sub
Public Sub drop(ByVal items As Long)
Dim num As Long
On Error Resume Next
If mstack.count < items Then: Set mstack = New Collection: Exit Sub
For num = 1 To items
    mstack.Remove 1
Next num

End Sub
Public Sub PushSecond(ByVal Long2 As Long)
Dim Long1
Dim Pack2Long As Variant
    Pack2Long = 0@
    Long1 = CLng(mstack.item(1))
    mstack.Remove 1
    PutMem4 VarPtr(Pack2Long) + 4, Long1
    PutMem4 VarPtr(Pack2Long) + 8, Long2
    If mstack.count = 0 Then
        mstack.Add Pack2Long
    Else
        mstack.Add Pack2Long, , 1
    End If
End Sub
Public Sub PushSecondThird(ByVal Long2 As Long, ByVal Long3 As Long)
Dim Long1
Dim Pack2Long As Variant
    
    Long1 = CLng(mstack.item(1))
    mstack.Remove 1
    PutMem1 VarPtr(Pack2Long), 14
    PutMem4 VarPtr(Pack2Long) + 4, Long1
    PutMem4 VarPtr(Pack2Long) + 8, Long2
    PutMem4 VarPtr(Pack2Long) + 12, Long3
    If mstack.count = 0 Then
        mstack.Add Pack2Long
    Else
        mstack.Add Pack2Long, , 1
    End If
End Sub
Public Sub Push2long(ByVal Long1 As Long, ByVal Long2 As Long)
    Dim Pack2Long As Variant
    Pack2Long = 0@
    PutMem4 VarPtr(Pack2Long) + 4, Long1
    PutMem4 VarPtr(Pack2Long) + 8, Long2
    If mstack.count = 0 Then
        mstack.Add Pack2Long
    Else
        mstack.Add Pack2Long, , 1
    End If
End Sub
Public Sub Push3long(ByVal Long1 As Long, ByVal Long2 As Long, ByVal Long3 As Long)
    Dim Pack3Long As Variant
    PutMem1 VarPtr(Pack3Long), 14
    PutMem4 VarPtr(Pack3Long) + 4, Long1
    PutMem4 VarPtr(Pack3Long) + 8, Long2
    PutMem4 VarPtr(Pack3Long) + 12, Long3
    If mstack.count = 0 Then
        mstack.Add Pack3Long
    Else
        mstack.Add Pack3Long, , 1
    End If
End Sub

Public Function Pop2Long(ByRef Long2 As Long) As Long
If mstack.count > 0 Then
    Dim PackLoad As Variant
    PackLoad = mstack.item(1)
    mstack.Remove 1
    GetMem4 VarPtr(PackLoad) + 4, Pop2Long
    GetMem4 VarPtr(PackLoad) + 8, Long2
End If
End Function
Public Function Pop3Long(Long2 As Long, Long3 As Long) As Long
If mstack.count > 0 Then
    Dim PackLoad As Variant
    PackLoad = mstack.item(1)
    mstack.Remove 1
    GetMem4 VarPtr(PackLoad) + 4, Pop3Long
    GetMem4 VarPtr(PackLoad) + 8, Long2
    GetMem4 VarPtr(PackLoad) + 12, Long3
End If
End Function
Public Function isPop3Long(Long2 As Long, Long3 As Long) As Long
If mstack.count > 0 Then
    Dim PackLoad As Variant, t As Byte
    PackLoad = mstack.item(1)
    GetMem1 VarPtr(PackLoad), t
    If t = 14 Then
    mstack.Remove 1
    GetMem4 VarPtr(PackLoad) + 4, isPop3Long
    GetMem4 VarPtr(PackLoad) + 8, Long2
    GetMem4 VarPtr(PackLoad) + 12, Long3
    Else
    isPop3Long = 0
    End If
End If
End Function
Public Function IsPop2Long(ByRef Long2 As Long) As Long
If mstack.count > 0 Then
    Dim PackLoad As Variant, t As Byte
    PackLoad = mstack.item(1)
    GetMem1 VarPtr(PackLoad), t
    If t = 6 Then
        mstack.Remove 1
        GetMem4 VarPtr(PackLoad) + 4, IsPop2Long
        GetMem4 VarPtr(PackLoad) + 8, Long2
    Else
        IsPop2Long = 0
    End If
End If
End Function
Public Sub PushObj(RHS As Variant)
If mstack.count = 0 Then
mstack.Add RHS
Else
mstack.Add RHS, , 1
End If
End Sub
Public Function IsEmpty() As Boolean
IsEmpty = mstack.count = 0
End Function
Public Property Get PopStr() As String
    On Error Resume Next
    SwapString2Variant PopStr, mstack.item(1)
    mstack.Remove 1
End Property

