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
' a stack only for strings for group objects
Private mstack As New Collection
Public Property Get Total() As Long
Total = mstack.Count
End Property
Public Property Get Count() As Long
Count = mstack.Count
End Property
Function Copy() As mStiva2
Dim i As Long
Set Copy = New mStiva2
For i = Total To 1 Step -1
Copy.PushVariant mstack.item(i)
Next i
End Function
Public Property Get LookTopVal() As Variant
If mstack.Count > 0 Then
LookTopVal = mstack.item(1)
End If
End Property
Public Property Get LookSecondVal() As Variant
If mstack.Count > 1 Then
LookSecondVal = mstack.item(2)
End If
End Property
Public Property Get PopVal() As Variant
If mstack.Count > 0 Then
PopVal = mstack.item(1)
   mstack.Remove 1
End If
End Property
Public Sub PushVal(a As Variant)
If mstack.Count = 0 Then
mstack.Add a
Else
mstack.Add a, , 1
End If
End Sub
Public Sub PushVariant(a As Variant)
If mstack.Count = 0 Then
mstack.Add a
Else
mstack.Add a, , 1
End If
End Sub
Public Sub PushLong(a As Long)
If mstack.Count = 0 Then
mstack.Add CVar(a)
Else
mstack.Add CVar(a), , 1
End If
End Sub
Public Sub MakeTopItem(where As Long)
'we pick an item and to the top
Dim a As Variant
On Error Resume Next
If where > 1 Then
a = mstack.item(where)
mstack.Remove where
If mstack.Count = 0 Then
mstack.Add a
Else
mstack.Add a, , 1
End If
End If
End Sub
Public Sub Flush()
Dim num As Long
For num = 1 To mstack.Count
   mstack.Remove 1
Next num
End Sub
Public Sub PushStr(ByRef a As String)
If mstack.Count = 0 Then
mstack.Add CVar(a)
Else
mstack.Add CVar(a), , 1
End If
End Sub
Public Property Get StackItem2(where As Long, maxvalue As Long) As Variant
If where > mstack.Count Then
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
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , , mstack.Count
End If
End Sub
Public Sub DataVal(a As Variant)
If mstack.Count = 0 Then
mstack.Add a
Else
mstack.Add a, , , mstack.Count
End If
End Sub
Public Sub DataStr(ByRef a As String)
Dim v As Variant
v = vbNullString
SwapString2Variant a, v
If mstack.Count = 0 Then
mstack.Add v
Else
mstack.Add v, , , mstack.Count
End If
End Sub
Sub GetStackPick(where, ByRef v As Variant)
SwapVariant v, mstack.item(where)
mstack.Remove where
End Sub

Public Property Get StackPick(where As Long) As Variant
StackPick = mstack.item(where)
mstack.Remove where
End Property
Public Sub MergeBottom(a As mStiva2)
If Me Is a Then Exit Sub
Dim i As Long, aa As Variant
If a.Total > 0 Then
Do
aa = a.StackPick(1)
If mstack.Count = 0 Then
mstack.Add aa
Else
mstack.Add aa, , , mstack.Count    'push so first of a is first on mstack
End If
Loop Until a.Total < 1
End If
End Sub
Public Sub DataStrUn(a As String)
Dim aa As Variant, i As Long, b As String
b = Split(a)(0)
aa = a
If mstack.Count = 0 Then
mstack.Add aa
Else
For i = 1 To Total
If Split(mstack.item(i))(0) = b Then

mstack.Remove i
If mstack.Count = 0 Then
mstack.Add aa
ElseIf i = 1 Then
mstack.Add aa, , 1
Else
mstack.Add aa, , , i - 1
End If
Exit Sub
End If
Next i
mstack.Add aa, , , mstack.Count
End If
End Sub
Public Sub drop(ByVal items As Long)
Dim num As Long
On Error Resume Next
If mstack.Count < items Then: Set mstack = New Collection: Exit Sub
For num = 1 To items
    mstack.Remove 1
Next num

End Sub
