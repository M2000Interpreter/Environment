VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stdCallFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mLib As String
Dim mFunction As String
Dim mConctruct As String
Dim RetType2Stack As Long   ' 0 is long, 1 is string without zeros
Dim mCallType As Long
Dim t() As Long
Dim paramnum As Long
Public FreeHere As Boolean
Public CallAddr As Long
Public IsInterfaceCall As Boolean
Private riid As IDispatch.IID
Private funcAddrOffset As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
                         ByVal lpszCLSID As Long, _
                         ByRef Clsid As IDispatch.IID) As Long
Private Declare Function DispCallFunc Lib "oleaut32" ( _
                         ByVal pvInstance As Any, _
                         ByVal oVft As Long, _
                         ByVal cc As Integer, _
                         ByVal vtReturn As Integer, _
                         ByVal cActuals As Long, _
                         ByRef prgVt As Any, _
                         ByRef prgpVarg As Any, _
                         ByRef pvargResult As Variant) As Long
Property Get Funcoffset() As Long
Funcoffset = funcAddrOffset
End Property
Public Function FixInterface(s As String, thisfunc As Long) As Boolean
    Dim hRet As Long
    IsInterfaceCall = False
    hRet = CLSIDFromString(StrPtr(s), riid)
    If hRet Then Exit Function
    FixInterface = True
    IsInterfaceCall = True
    funcAddrOffset = thisfunc * 4  ' 32 bit
    CallAddr = 0
End Function

Public Function GetInterface(myobject As stdole.IUnknown, ByRef retobject As stdole.IUnknown) As Boolean
    Const CC_STDCALL          As Long = 4
    If IsInterfaceCall Then
    Dim params(2)   As Variant
    Dim types(2)    As Integer
    Dim list(2)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    Dim pSomething
    params(0) = VarPtr(riid)
    params(1) = VarPtr(retobject)
    For pIndex = 0 To 1
        list(pIndex) = VarPtr(params(pIndex)):  types(pIndex) = VarType(params(pIndex))
    Next
        resultCall = DispCallFunc(ObjPtr(myobject), &H0&, CC_STDCALL, vbLong, 2, types(0), list(0), pReturn)
        If resultCall < 0 Then Exit Function
        If pReturn = &H80004002 Then
            MyEr "interface not supported", "Η διεπαφή δεν υποστηρίζεται"
            Exit Function
        ElseIf pReturn = &H80070057 Then
            MyEr "wrong argument", "λάθος τύπος ορίσματος"
            Exit Function
        End If
    End If
    GetInterface = True
End Function

Public Sub CallThis(ByVal vNewValue As String, ParamList As String, Lang As Long)
Dim l As Long, check$, M As Long
If Not Me.IsInterfaceCall Then
    l = InStr(vNewValue, ".")
    If l <> 0 Then
    mLib = Left$(vNewValue, l - 1)
    mFunction = Mid$(vNewValue, l + 1)
    Else
    mLib = vNewValue
    mFunction = vbNullString ' what about that??
    End If
    
    mConctruct = ParamList
    If InStr(vNewValue, ".dll") > 0 Then
    mLib = mLib + ".dll"
    End If
    If LCase(mFunction) Like "dll.*" Then
    mFunction = Mid$(mFunction, 5)
    ElseIf LCase(mFunction) Like "dll*" Then
    mFunction = Mid$(mFunction, 4)
    End If
End If
Do While ParamList <> ""

    If IsLabelSymbolNew(ParamList, "ΜΑΚΡΥΣ", "LONG", Lang) Then
        IsLabelSymbolNew ParamList, "ΜΑΚΡΥΣ", "LONG", Lang
        If FastSymbol(ParamList, "&") Then l = 16 Else l = 0
        If IsPureLabel(ParamList, check$) = 1 Then
            M = 2
        Else
            ParamList = "error"
            Exit Sub
        End If
    ElseIf IsLabelSymbolNew(ParamList, "ΔΙΠΛΟΣ", "DOUBLE", Lang) Then
        If FastSymbol(ParamList, "&") Then l = 16 Else l = 0
        If IsPureLabel(ParamList, check$) = 1 Then
            M = 64
        Else
            ParamList = "error"
            Exit Sub
        End If
    Else
        If FastSymbol(ParamList, "&") Then l = 16 Else l = 0
            M = IsPureLabel(ParamList, check$)
            If M = 0 Then
                If FastSymbol(ParamList, "...", , 3) Then
                PushType -100  ' Paramarray
                Exit Do
            End If
        End If
    End If
    PushType M + l   'above 16 is byref
    If IsArrayItem(paramnum) Then
        If Not FastSymbol(ParamList, ")") Then Exit Do
    End If
    If Not FastSymbol(ParamList, ",") Then Exit Do
    If check$ = vbNullString Then Exit Do
    If paramnum = 63 Then Exit Do
Loop
ParamList = Trim$(ParamList)   ' if we have somethig then that is an error
End Sub
Public Property Get LIB() As String
LIB = mLib
End Property

Public Property Get func() As String
func = mFunction
End Property

Public Function ReadType(there As Long) As Long
If there < 0 Then Exit Function
If t(there) < 0 Then
ReadType = t(there)
Else
ReadType = (t(there) And &HFFEF)
End If
End Function

Public Function IsArrayItem(there As Long) As Boolean
Dim a As Long
a = (t(there) And &HFFEF)
' 5 and 7
IsArrayItem = (a < 8 And a > 4)
End Function
Public Function checktype(there As Long, ByVal thistype As Long) As Boolean
Select Case thistype
Case 1, 4, 5, 7
thistype = 1  ' AS IS
Case Else
thistype = 2 ' BY POINTER
End Select
Select Case (t(there) And &HFFEF)
Case 1, 4, 5, 7
checktype = thistype = 1 ' AS IS
Case Else
checktype = thistype = 2 ' BY POINTER
End Select
End Function
Public Function IsByRef(there As Long) As Boolean
IsByRef = (t(there) And &H10) <> 0
End Function


Private Sub PushType(ParamType As Long)
Dim hlp As Long
If paramnum + 1 >= UBound(t()) - LBound(t()) Then
hlp = UBound(t()) * 2
ReDim Preserve t(hlp)
End If
paramnum = paramnum + 1
t(paramnum) = ParamType


End Sub

Private Sub Class_Initialize()
ReDim t(16) As Long
paramnum = -1
End Sub

Public Property Get Count() As Long
Count = paramnum + 1
End Property

Public Property Get RetType() As Long
RetType = RetType2Stack
End Property

Public Property Let RetType(ByVal vNewValue As Long)
RetType2Stack = vNewValue
End Property

Public Property Get CallType() As Long
CallType = mCallType
End Property

Public Property Let CallType(ByVal vNewValue As Long)
mCallType = vNewValue
End Property

Private Sub Class_Terminate()
If FreeHere Then
If mFunction <> "" Then
RemoveDll mFunction, True
End If
End If
End Sub
Public Function GetAddress() As Long
If IsInterfaceCall Then
    MyEr "Not for Interfaces", "Όχι για διεπαφές"
    Exit Function
End If
If Me.CallAddr = 0 Then
If Left$(func, 1) = "#" Then
    CallAddr = Fcall.GetFuncPtrOrd(LIB, func)
Else
    CallAddr = Fcall.GetFuncPtr(LIB, func)
    
End If
End If

    GetAddress = CallAddr

End Function
