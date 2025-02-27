VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Generic Handler
Option Explicit
Private mObjRef As Object
Private mt1 As Variant, mindirect As Long
Public ReadOnly As Boolean
Public index_start As Variant ' 1 for start
Public index_End As Variant  ' -1 for end
Public index_cursor As Variant
Public sign As Variant
Public UseIterator As Boolean
Public oEnum As Variant
Private mIamEnum As Boolean
Friend Property Get objref() As Object
    Set objref = mObjRef
End Property

Friend Property Set objref(RHS As Object)
    Set mObjRef = RHS
End Property

Private Sub Class_Initialize()
indirect = -1
index_start = 0
index_End = -1
UseIterator = False
sign = 1
End Sub

''Public Par1 As Variant
Private Sub Class_Terminate()
If Not objref Is Nothing Then
'If Typename(objref) = "GuiM2000" Then Stop
Set objref = Nothing
End If
End Sub
Friend Sub CopyTo(a As mHandler)
If IamEnum Then Set a = Me: Exit Sub
Set a = New mHandler
a.T1 = T1
If T1 = 1 Then a.ReadOnly = ReadOnly
a.indirect = indirect
a.index_start = index_start
a.index_End = index_End
a.index_cursor = index_cursor
a.UseIterator = UseIterator
a.sign = sign
Set a.objref = mObjRef
End Sub
Friend Function Iterate()
        Iterate = GetNext(ObjPtr(oEnum), index_cursor)
        If IsObject(index_cursor) Then Set mObjRef = index_cursor: index_cursor = vbEmpty
        index_End = 1
        If Not Iterate Then index_End = -1
End Function
Friend Sub ConstructEnumerator(penum As Variant)
IamEnum = True
Set oEnum = penum
If GetNext(ObjPtr(penum), index_cursor) Then
If IsObject(index_cursor) Then Set mObjRef = index_cursor: index_cursor = vbEmpty
    index_End = 0
End If
End Sub
Friend Function CreateFromPicture(b As Long, w1 As Long, w2 As Long, w3 As Long) As Object
If T1 = 2 Then
Dim mm As MemBlock
Set mm = mObjRef
Set CreateFromPicture = mm.CreateFromPicture(b, w1, w2, w3)
End If
End Function

Property Get T1() As Variant
T1 = mt1
End Property

Friend Property Let T1(ByVal RHS As Variant)
mt1 = RHS
End Property

Property Get indirect() As Long
indirect = mindirect
End Property

Friend Property Let indirect(ByVal RHS As Long)
mindirect = RHS
End Property
Property Get IamEnum() As Boolean
IamEnum = mIamEnum
End Property

Friend Property Let IamEnum(ByVal RHS As Boolean)
mIamEnum = RHS
End Property
