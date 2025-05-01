VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim VarRef As String
Dim FuncRef As String
Public oldFuncRef As String
Dim LocalRef As String  ' used only for classes
Dim mName As String
Dim mFloatGroupName As String
Dim FloatGroup As Boolean
Dim bank
Dim mTeam As Boolean
Dim LinkToGroup As Group
Dim mTag As String
Public Patch As String
Public IamGlobal As Boolean, IamRef As Boolean
Public HasValue As Boolean, HasSet As Boolean, HasStrValue As Boolean, HasParameters As Boolean
Public HasParametersSet As Boolean
Public SuperClassList As Object
Public IamSuperClass As Boolean
Private eventlist As FastCollection  ' we have to copy this
Public highpriorityoper As String
Public HasUnary As Boolean
Public IamApointer As Boolean, lasthere$
Public original As Long
Public isRef As Boolean ' One time use
Public LastOpen As String
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public PointerPtr As Long
Public ToDelete As Boolean, HasRemove As Boolean, mytypes As Collection
Public IamCleared As Boolean
Public Fields As Variant
Sub Add2Field(st$)
    Dim there As Long
    On Error Resume Next
    there = UBound(Fields)
    If Err Then
        Err.Clear
        there = 0
        Fields = LibMemory.EmptyArray(1, vbString)
    End If
    ReDim Preserve Fields(there + 1)
    Fields(there) = st$
    st$ = vbNullString
End Sub
Function ReadField(there As Long) As String
    ReadField = Fields(there)
End Function
Sub MergeField(st$)
Dim b As String, i As Long
b = Split(st$)(0)
If UBound(Fields) - 1 = 0 Then
    i = UBound(Fields)
    ReDim Preserve Fields(i + 1)
    Fields(i) = st$
Else
    For i = 1 To UBound(Fields) - 1
        If Split(Fields(i))(0) = b Then
            Fields(i) = st$
            Exit Sub
        End If
    Next i
    i = UBound(Fields)
    ReDim Preserve Fields(i + 1)
    Fields(i) = st$
End If
End Sub
Property Get FieldsCount() As Long
    On Error Resume Next
  
    FieldsCount = UBound(Fields) - 1
    
End Property
Friend Property Let PutIs(RHS As String)
If mytypes Is Nothing Then
    Set mytypes = New Collection
    mytypes.Add RHS, RHS
Else
On Error Resume Next
    mytypes.Add RHS, RHS
End If
End Property
Function TypeGroup(RHS) As Boolean
If Not mytypes Is Nothing Then
    On Error Resume Next
    TypeGroup = Len(mytypes(RHS))
    Err.Clear
End If
End Function
Private Function TypeGroupNow(RHS) As Boolean
    On Error Resume Next
  '  TypeGroupNow = Len(mytypes(RHS))
    TypeGroupNow = True
    Err.Clear
End Function
Sub mergeTypes(RHS As Collection)
Dim extypes As Collection, K
Set extypes = RHS
If extypes Is Nothing Then Exit Sub
If mytypes Is RHS Then
    Set mytypes = New Collection
End If
For Each K In extypes
    PutIs = K
Next
End Sub
Function MatchTypes(RHS As Collection) As Boolean
Dim extypes As Collection, K
Set extypes = RHS
If extypes Is Nothing Then Exit Function
If mytypes Is Nothing Then Exit Function
If mytypes Is RHS Then
    MatchTypes = True
    Exit Function
End If
If mytypes.Count * extypes.Count = 0 Then Exit Function
For Each K In extypes
    If TypeGroupNow(K) Then MatchTypes = True: Exit Function
Next

End Function
Public Property Get Events() As Object
Set Events = eventlist
End Property
Public Property Set Events(RHS As Object)
Set eventlist = RHS
End Property
Public Sub ResetEvents()
If eventlist Is Nothing Then Exit Sub
Dim i As Long
With eventlist
For i = 0 To .Count - 1
    .Index = i
    .Done = True
    .Value = vbNullString
    .sValue = 0
Next i
.Done = False
End With

End Sub
Public Sub FlushEvents()
If eventlist Is Nothing Then Exit Sub
Dim eventlist2 As New FastCollection
Dim i As Long
With eventlist
For i = 0 To .Count - 1
    .Index = i
    .Done = True
    eventlist2.AddKey .KeyToString, ""
Next i
.Done = False
End With
Set eventlist = eventlist2
End Sub
Public Sub RegisterEvent(ByVal a$)
a$ = myUcase(a$)
If eventlist Is Nothing Then Set eventlist = New FastCollection
If eventlist.ExistKey(a$) Then Exit Sub
eventlist.AddKey a$, vbNullString
End Sub
Public Function EventFuncPos(a$, retName$, Pos As Long) As Boolean
If eventlist Is Nothing Then EventFuncPos = 0: Exit Function
With eventlist
If eventlist.ExistKey(a$) Then
    retName$ = .Value
    If retName$ <> vbNullString Then Pos = .sValue: EventFuncPos = True
End If
End With
End Function

Public Function PrepareSoros(ByRef v(), ByVal topname$) As mStiva2
Dim i As Long, AllSoros As New mStiva2, s As String, vv As mStiva2
For i = 1 To FieldsCount
s = Fields(i)
'Debug.Print fields(i), topname$, GroupName

If Left$(s, 1) = "*" Then
AllSoros.DataStr "*" + topname$ + GroupName + Mid$(s, 2)
 If Not IsObject(v(val(Split(s)(1)))) Then Set v(val(Split(s)(1))) = New Group
' If here$ <> "" And Not IamGlobal Then
 Set vv = v(val(Split(s)(1))).PrepareSoros(v(), Patch + ".")
' Else
 Set vv = v(val(Split(s)(1))).PrepareSoros(v(), topname$ + GroupName)
 'End If
 
AllSoros.MergeBottom vv
Else
If val(Split(s)(1)) = 0 Then
AllSoros.DataStr s
Else
AllSoros.DataStr topname$ + GroupName + s
End If
End If
Next
Set PrepareSoros = AllSoros

End Function
Public Function PrepareSoros1(ByRef v(), ByVal topname$) As mStiva2
Dim i As Long, AllSoros As New mStiva2, vv As mStiva2, s As String
For i = 1 To FieldsCount
s = Fields(i)
If Left$(s, 1) = "*" Then
AllSoros.DataStr "*" + topname$ + "." + Split(Mid$(s, 2))(0)

 Set vv = v(val(Split(s)(1))).PrepareSoros1(v(), GroupName)
 
AllSoros.MergeBottom vv
Else
AllSoros.DataStr GroupName + Split(s)(0)
End If
Next
Dim a() As String
If FuncRef$ <> "" Then
a() = Split(Replace(FuncRef$, Chr$(3), Chr$(2)), Chr$(2))
If UBound(a()) > 0 Then
For i = 1 To UBound(a())
AllSoros.DataStr topname$ + "." + Split(a(i))(0)
Next i
End If
End If
Set PrepareSoros1 = AllSoros

End Function
Public Function PrepareSorosToCompare(ByRef v()) As mStiva2
Dim i As Long, AllSoros As New mStiva2, vv As mStiva2, s As String
For i = 1 To FieldsCount
    s = Fields(i)
    If Left$(s, 1) = "*" Then
        AllSoros.DataStr "*" + Split(Mid$(s, 2))(0)
        If Not IsObject(v(val(Split(s)(1)))) Then Set v(val(Split(s)(1))) = New Group
        Set vv = v(val(Split(s)(1))).PrepareSoros1(v(), "")
        AllSoros.MergeBottom vv
    Else
        AllSoros.DataStr (Split(s)(0))
    End If
Next
Dim a() As String
If FuncRef$ <> "" Then
a() = Split(Replace(FuncRef$, Chr$(3), Chr$(2)), Chr$(2))
If UBound(a()) > 0 Then
For i = 1 To UBound(a())
AllSoros.DataStr (Split(a(i))(0))
Next i
End If
End If
Set PrepareSorosToCompare = AllSoros
End Function
Public Function types() As String
Dim i As Long, s$
i = FieldsCount
On Error Resume Next
While i > 0
If InStr(Fields(i), "( ") > 0 Then
s$ = "A" + s$
Else
s$ = "V" + s$
End If

i = i - 1

Wend
types = s$
End Function
Public Sub PeekItem(cursor As Long, item)
'' used in the float mode
Dim JJ As Variant

If IsObject(bank(cursor)) Then
Set item = bank(cursor)
Else

If mTeam Then
        If cursor = bank(1) * 2 + 3 Then
            ' we need common definitions for functions and modules all elements
                If Not LinkToGroup.IamApointer Then
                
                LinkToGroup.PeekItem 1, JJ
               ' If JJ = 0 Then Exit Sub
                LinkToGroup.PeekItem JJ * 2 + 3, item
                        Exit Sub
                        End If
            ElseIf cursor = bank(1) * 2 + 2 Then
                    If Not LinkToGroup.IamApointer Then
                    LinkToGroup.PeekItem 1, JJ
                  '  If JJ = 0 Then Exit Sub
                    LinkToGroup.PeekItem JJ * 2 + 2, item
                          Exit Sub
                          End If
            End If
End If
If MemInt(VarPtr(bank(cursor))) = 20 Then
item = MemCur(VarPtr(bank(cursor)) + 8)
MemInt(VarPtr(item)) = 20
Else
item = bank(cursor)
End If
End If

End Sub
Public Sub PokeItem(cursor As Long, item)
' used in the float mode
'Debug.Print Typename$(item)
If IsObject(item) Then

Set bank(cursor) = item
Else
' always last two items are for local vars and functions
If mTeam And cursor >= bank(1) * 2 + 2 Then

'' do nothing
Else
 bank(cursor) = item
 End If
End If
End Sub

Public Sub BeginFloat(Total As Long)
FloatGroup = True
 bank = LibMemory.EmptyArray(1, vbVariant)
 If Total > 0 Then ReDim bank(Total * 2 - 1)
'bank.grouperror = True
'bank.PushDim Total * 2
'bank.PushEnd
End Sub

Public Property Get IamFloatGroup()
IamFloatGroup = FloatGroup
End Property
Public Property Get IamTeam()
IamTeam = mTeam
End Property
Public Sub EndFloat()
'If FloatGroup Then bank.SerialItem (0), 0, 7
FloatGroup = False
bank = Empty
End Sub
Public Property Set LinkRefPointer(a As Object)
If FloatGroup Then
If Typename(a) <> "Group" Then
Exit Property
End If
mTeam = False
Set LinkToGroup = a
If a Is Nothing Then Exit Property
'
bank(bank(1) * 2 + 3) = vbNullString
'mTeam = True
Else
' only when is in an array as Float Group
Set LinkToGroup = a
End If
End Property
Public Property Set LinkRef(a As Object)
If FloatGroup Then
If Typename(a) <> "Group" Then
Exit Property
End If
mTeam = False
Set LinkToGroup = a
If a Is Nothing Then Exit Property
'erase functions,,,
 bank(bank(1) * 2 + 3) = vbNullString
mTeam = True
Else
' only when is in an array as Float Group
Set LinkToGroup = a
End If
End Property
Friend Property Get LinkGroup() As Group
Set LinkGroup = LinkToGroup
End Property
Public Property Get link() As Object
Set link = LinkToGroup
End Property

Public Property Let Varlist(vNewValue As String)
VarRef$ = vNewValue
End Property
Public Property Get Varlist() As String
Varlist = VarRef$
End Property


Public Property Get FuncList() As String
If IamApointer Then
FuncList = Me.link.FuncList
Else
FuncList = FuncRef$
End If
End Property
Public Property Let FuncList(vNewValue As String)
FuncRef$ = vNewValue
End Property


Public Property Get localList() As String
localList = LocalRef$
End Property

Public Property Let localList(vNewValue As String)
LocalRef$ = vNewValue
End Property

Public Property Get FloatGroupName() As String
 FloatGroupName = mFloatGroupName
End Property

Public Property Let FloatGroupName(vNewValue As String)
mFloatGroupName = vNewValue
End Property

Public Property Get GroupName() As String
GroupName = mName
End Property

Public Property Let GroupName(vNewValue As String)
mName = vNewValue
End Property


Private Sub Class_Initialize()
Dim f() As String
ReDim f(1) As String
Fields = f()
End Sub

Private Sub Class_Terminate()
Dim iLastErNum As Long, iLastErNum1 As Long
Set eventlist = Nothing
If IamApointer Then
    If Not LinkToGroup Is Nothing Then
        If LinkToGroup.IamFloatGroup Then
            If LinkToGroup.HasRemove Then
                If Not LinkToGroup.ToDelete Then
                     If LinkToGroup.refcount1 = 0 Then
                        LinkToGroup.ToDelete = True
                        iLastErNum = LastErNum
                        iLastErNum1 = LastErNum1
                        LastErNum = 0
                        LastErNum1 = 0
                        ProcessOperRemove link
                        LastErNum = iLastErNum
                        LastErNum1 = iLastErNum1
                    End If
                End If
            End If
        End If
    End If
ElseIf FloatGroup And Not ToDelete And HasRemove Then
    iLastErNum = LastErNum
    iLastErNum1 = LastErNum1
    LastErNum = 0
    LastErNum1 = 0
    ProcessOperRemove Me
    LastErNum = iLastErNum
    LastErNum1 = iLastErNum1
End If
End Sub

Public Property Get edittag() As String
If mTeam Then
edittag = LinkToGroup.edittag
Else
edittag = mTag
End If
End Property

Public Property Let edittag(vNewValue As String)
If Not mTeam Then mTag = vNewValue
End Property

Public Sub ResetGroup()
IamCleared = True
Set SuperClassList = Nothing
If FloatGroup Then EndFloat
Set LinkToGroup = Nothing
FloatGroup = False
ReDim Fields(0)
Set mytypes = Nothing
HasValue = False
HasSet = False
HasStrValue = False
HasParameters = False
IamSuperClass = False
oldFuncRef = vbNullString
ResetFunctionList FuncRef
LocalRef = vbNullString
VarRef = vbNullString
ToDelete = False
HasRemove = True
highpriorityoper = vbNullString
HasUnary = False

End Sub
Public Property Get refcount1() As Long
Dim M As Long
M = GetRefCount(Me)
If M > 3 Then refcount1 = M - 3 Else refcount1 = 0
End Property
Private Function GetRefCount(ByRef vvv) As Long
    If Not IsObject(vvv) Then Exit Function
    Dim obj As IUnknown
    Set obj = vvv
    If obj Is Nothing Then Exit Function
    CopyMemory GetRefCount, ByVal (ObjPtr(obj)) + 4, 4
    GetRefCount = GetRefCount - 2
    Set obj = Nothing
End Function
