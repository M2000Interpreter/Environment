VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
                         ByVal lpszCLSID As Long, _
                         ByRef Clsid As GUID) As Long
Public EnumName As String
Private iGUID As GUID
Public LastGUIDstr As String
Private Type item
    Key As String
    KeyUCase As String
    iValue As Variant
    Pad As String
End Type
Public Done As Boolean
Public Index As Long
Private PriveSpace() As item, mInterface As Boolean
Dim MaxSpace As Long
Dim toplim As Long
Private Sub Class_Initialize()
    MaxSpace = 20
    ReDim PriveSpace(MaxSpace) As item
    toplim = -1
End Sub
Property Get Value() As Variant
Done = False
If Index = -1 Then
Else
    Done = True
    Value = PriveSpace(Index).iValue
    
End If
End Property
Property Get PadValue() As String
Done = False
If Index = -1 Then
Else
    Done = True
    PadValue = PriveSpace(Index).Pad
End If
End Property

Property Let Guidonce(s As String)
If toplim = -1 Then
    mInterface = CLSIDFromString(StrPtr(s), iGUID) = 0
    If mInterface Then
        LastGUIDstr = s
    End If
End If
End Property
Property Get IsInterface() As Boolean
    IsInterface = mInterface
End Property
Property Get GuidAddr() As Long
    GuidAddr = VarPtr(iGUID)
End Property

Private Function Malloc() As Long
    If toplim + 1 >= MaxSpace Then
        MaxSpace = MaxSpace * 2
        ReDim Preserve PriveSpace(MaxSpace) As item
    End If
    toplim = toplim + 1
    Malloc = toplim
End Function
Property Get Count()
    Count = toplim + 1
End Property
Property Get ZeroValue()
If toplim >= 0 Then ZeroValue = PriveSpace(0).iValue
End Property

Property Get IsEmpty()
    IsEmpty = toplim = -1
End Property

Public Sub addone(Key As String, nValue As Variant)
Dim a As Long
a = Malloc()
With PriveSpace(a)
    .Key = Key
    .KeyUCase = myUcase(Key, True)
    .iValue = nValue
End With
End Sub
Public Sub addone1(Key As String, nValue As Variant, pd$)
Dim a As Long
a = Malloc()
With PriveSpace(a)
    .Key = Key
    .KeyUCase = myUcase(Key, True)
    .iValue = nValue
    .Pad = pd$
End With
End Sub
Function ExistFromOther(RHS) As Boolean
Dim i As Long
For i = 0 To toplim
If RHS = PriveSpace(i).iValue Then ExistFromOther = True: Exit Function
Next i
End Function
Function ExistFromOther2(useHandler As mHandler) As Boolean
Dim other As Enumeration
Set other = useHandler.objref
other.Index = useHandler.index_start
Dim s$, i As Long
s$ = other.KeyToString
For i = 0 To toplim
If s$ = PriveSpace(i).Key Then ExistFromOther2 = True: useHandler.index_start = i: Exit For
Next i

End Function
Function SearchSimple(what As String, ByRef ok As Boolean) As Variant
Dim i As Long
For i = 0 To toplim
If what = PriveSpace(i).KeyUCase Then ok = True: SearchSimple = PriveSpace(i).iValue: Index = i: Exit Function
Next i
End Function
Function SearchValue(v As Variant, ByRef ok As Boolean) As Variant
Dim i As Long, mm As mHandler
ok = False
again:
For i = 0 To toplim
If v = PriveSpace(i).iValue Then
ok = True
Set mm = New mHandler
mm.t1 = 4
Set mm.objref = Me
mm.index_cursor = v
mm.index_start = i
mm.sign = 1
Set SearchValue = mm
Exit Function
End If
Next i
If Not ok Then
If MemInt(VarPtr(v)) = vbString Then
v = val(v)
GoTo again
Else

v = -v
For i = 0 To toplim
If v = PriveSpace(i).iValue Then
ok = True
Set mm = New mHandler
mm.t1 = 4
Set mm.objref = Me
mm.index_cursor = v
mm.index_start = i
mm.sign = -1
Set SearchValue = mm
Exit Function
End If
Next i
End If
End If
Set SearchValue = New mHandler
End Function
Public Property Get KeyToString() As Variant
    If Index <> -1 Then KeyToString = PriveSpace(Index).Key
End Property

Private Sub Class_Terminate()
    Erase PriveSpace()
End Sub


