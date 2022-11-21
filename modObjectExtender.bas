Attribute VB_Name = "modObjectExtender"
Option Explicit
Public btASM As Long
' modObjectExtender
'
' event support for Late-Bound objects
' low level COM Projekt - by [rm_code] 2005
Private Declare Function ObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef objDest As Object, ByVal pObject As Long) As Long

Public Type EventSink
    pVTable     As Long     ' VTable pointer
    pClass      As Long     ' ComShinkEvent pointer
    cRef        As Long     ' reference counter
    IID         As GUID     ' interface IID
    hMem        As Long     ' memory address
End Type
' for DEP
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Const MEM_DECOMMIT = &H4000
Private Const MEM_RELEASE = &H8000
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const PAGE_EXECUTE_READWRITE = &H40

Private Declare Function CallWindowProcA Lib "user32" ( _
    ByVal adr As Long, ByVal p1 As Long, ByVal p2 As Long, _
    ByVal p3 As Long, ByVal p4 As Long) As Long
Public Declare Function VariantCopyIndPtr Lib "oleaut32" Alias "VariantCopyInd" ( _
    ByVal pvargDest As Long, ByVal pvargSrc As Long) As Long

Public Declare Function SysAllocStringPtr Lib "oleaut32" ( _
    ByVal pStr As Long) As Long

Public Declare Function SysReAllocString Lib "oleaut32" ( _
    ByVal StrSrc As Long, ByVal StrNew As Long) As Long

Public Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" ( _
    PtrDest() As Any) As Long

Public Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal dwLen As Long)

Public Declare Sub FillMem Lib "kernel32" Alias "RtlFillMemory" ( _
    pDst As Any, ByVal dlen As Long, ByVal Fill As Byte)

Public Declare Function IsEqualGUID Lib "ole32" ( _
    rguid1 As GUID, rguid2 As GUID) As Long

Public Declare Function CLSIDFromString Lib "ole32" ( _
    ByVal lpsz As Long, GUID As Any) As Long

Public Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal uFlags As Long, ByVal dwBytes As Long) As Long

Public Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Public Declare Function LCID_def1 Lib "kernel32" Alias "GetSystemDefaultLCID" ( _
    ) As Long

Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_NOTIMPL As Long = &H80004001

'             GMEM | GMEM_ZEROINIT
Private Const GPTR As Long = &H40&

Public Const IIDSTR_IUnknown As String = _
    "{00000000-0000-0000-C000-000000000046}"

Public Const IIDSTR_IDispatch As String = _
    "{00020400-0000-0000-C000-000000000046}"

Public Const IIDSTR_IConnectionPoint As String = _
    "{B196B286-BAB4-101A-B69C-00AA00341D07}"

Public Const IIDSTR_IEnumConnectionPoints As String = _
    "{B196B285-BAB4-101A-B69C-00AA00341D07}"

Public Const IIDSTR_IConnectionPointContainer As String = _
    "{B196B284-BAB4-101A-B69C-00AA00341D07}"

Public IID_IUnknown     As GUID
Public IID_IDispatch    As GUID

Private Const MAXCODE   As Long = &HEC00&
Private ObjExt_vtbl(6) As Long
''get lcid_def1() once
Public LCID_DEF As Long

Private Type IUnknown100
    QueryInterface              As Long
    AddRef                      As Long
    Release                     As Long
End Type
Private Type IEnum100
    iunk                        As IUnknown100
    Next                        As Long
    skip                        As Long
    Reset                       As Long
    Clone                       As Long
End Type
Public Declare Function vbaCastObj Lib "msvbvm60" _
                         Alias "__vbaCastObj" ( _
                         ByRef cObj As Any, _
                         ByRef pIID As Any) As Long
Private Declare Function PutMem4 Lib "msvbvm60" ( _
                         ByRef pDst As Any, _
                         ByVal NewValue As Long) As Long
Public Function GetNext(pECP As Long, usethis As Variant) As Boolean
Dim hRet As Long
Dim cReturned   As Long
Dim pVTblECP    As Long
Dim oECP        As IEnum100
Dim pCP         As Long
    CpyMem pVTblECP, ByVal pECP, 4
    CpyMem oECP, ByVal pVTblECP, Len(oECP)

hRet = CallPointer(oECP.Next, pECP, 1, VarPtr(usethis), VarPtr(cReturned))

If hRet Then Exit Function
If cReturned = 0 Then Exit Function
GetNext = True

End Function

Public Sub InitObjExtender()
    Static blnInit  As Boolean
    If blnInit Then Exit Sub

    CLSIDFromString StrPtr(IIDSTR_IUnknown), IID_IUnknown
    CLSIDFromString StrPtr(IIDSTR_IDispatch), IID_IDispatch

    ObjExt_vtbl(0) = addr(AddressOf ObjExt_QueryInterface)
    ObjExt_vtbl(1) = addr(AddressOf ObjExt_AddRef)
    ObjExt_vtbl(2) = addr(AddressOf ObjExt_Release)
    ObjExt_vtbl(3) = addr(AddressOf ObjExt_GetTypeInfoCount)
    ObjExt_vtbl(4) = addr(AddressOf ObjExt_GetTypeInfo)
    ObjExt_vtbl(5) = addr(AddressOf ObjExt_GetIDsOfNames)
    ObjExt_vtbl(6) = addr(AddressOf ObjExt_Invoke)

    blnInit = True
End Sub

' IUnknown::QueryInterface
Private Function ObjExt_QueryInterface(this As EventSink, riid As GUID, pObj As Long) As Long

    ' IUnknown
    If IsEqualGUID(riid, IID_IUnknown) Then
        pObj = VarPtr(this)
        ObjExt_AddRef this

    ' IDispatch
    ElseIf IsEqualGUID(riid, IID_IDispatch) Then
        pObj = VarPtr(this)
        ObjExt_AddRef this

    ' event interface
    ElseIf IsEqualGUID(riid, this.IID) Then
        pObj = VarPtr(this)
        ObjExt_AddRef this

    ' not an implemented interface
    Else
        pObj = 0
        ObjExt_QueryInterface = E_NOINTERFACE

    End If
End Function

' IUnknown::AddRef
Private Function ObjExt_AddRef(this As EventSink) As Long
    this.cRef = this.cRef + 1
    ObjExt_AddRef = this.cRef
End Function

' IUnknown::Release
Private Function ObjExt_Release(this As EventSink) As Long
    this.cRef = this.cRef - 1
    ObjExt_Release = this.cRef

    ' if reference count is 0, free the object
    If this.cRef = 0 Then GlobalFree this.hMem
    'If this.cRef = 0 Then CoTaskMemFree this.hMem
End Function

' IDispatch::GetTypeInfoCount
Private Function ObjExt_GetTypeInfoCount(this As EventSink, pctinfo As Long) As Long
    pctinfo = 0
    ObjExt_GetTypeInfoCount = E_NOTIMPL
End Function

' IDispatch::GetTypeInfo
Private Function ObjExt_GetTypeInfo(this As EventSink, ByVal iTInfo As Long, ByVal lcid As Long, ppTInfo As Long) As Long
    ppTInfo = 0
    ObjExt_GetTypeInfo = E_NOTIMPL
End Function

' IDispatch::GetIDsOfNames
Private Function ObjExt_GetIDsOfNames(this As EventSink, riid As GUID, rgszNames As Long, ByVal cNames As Long, ByVal lcid As Long, rgDispId As Long) As Long
    ObjExt_GetIDsOfNames = E_NOTIMPL
End Function

' IDispatch::Invoke
Private Function ObjExt_Invoke(this As EventSink, _
         ByVal dispIdMember As Long, _
         riid As GUID, _
         ByVal lcid As Long, _
         ByVal wFlags As Integer, _
         ByVal pDispParams As Long, _
         ByVal pVarResult As Long, _
         ByVal pExcepInfo As Long, _
         puArgErr As Long) As Long

    ' get the object extender class
    ' which owns this event sink
        Dim objext  As ComShinkEvent
        Set objext = ResolveObjPtr(this.pClass)
        objext.fireevent dispIdMember, pDispParams
    ' forward the event
    
End Function

Public Function CreateEventSink(IID As GUID, objext As ComShinkEvent) As Object
    Dim sink    As EventSink

    ' our event sink object :)
    With sink
        .cRef = 1
        .IID = IID
        .pClass = ObjPtr(objext)
        .pVTable = VarPtr(ObjExt_vtbl(0))
    End With
    ' allocate some memory for our object
    sink.hMem = GlobalAlloc(GPTR, Len(sink))
    'sink.hMem = CoTaskMemAlloc(Len(sink))
    If sink.hMem = 0 Then Exit Function
    CpyMem ByVal sink.hMem, sink, Len(sink)

    ' return the object
    CpyMem CreateEventSink, sink.hMem, 4&
End Function
Private Function addr(p As Long) As Long
    addr = p
End Function

' Pointer->Object
Public Function ResolveObjPtr(ByVal Ptr As Long) As IUnknown
ObjSetAddRef ResolveObjPtr, Ptr
End Function
' ObjSetAddRef ObjectFromPtr, Ptr

Public Function Getenumerator(uOnk As IUnknown) As Boolean
Getenumerator = True
End Function


Public Function CallPointer(ByVal fnc As Long, ParamArray params()) As Long
Static once As Boolean
If once Then Exit Function
once = True

If btASM = 0 Then
 btASM = VirtualAlloc(ByVal 0&, MAXCODE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
 End If
 If btASM = 0 Then
 MyEr "DEP or memory problem", "Πρόβλημα με την μνήμη"
 CallPointer = -1
 Exit Function
 End If
    VirtualLock btASM, MAXCODE
   ' Dim btASM(MAXCODE - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    pASM = btASM

    FillMem ByVal pASM, MAXCODE, &HCC

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    If UBound(params) = 0 Then
        If IsArray(params(0)) Then
            For i = UBound(params(0)) To 0 Step -1
                AddPush pASM, CLng(params(0)(i))    ' PUSH dword
            Next
        Else
           AddPush pASM, CLng(params(0))       ' PUSH dword
        End If
    Else
        For i = UBound(params) To 0 Step -1
            AddPush pASM, CLng(params(i))           ' PUSH dword
        Next
    End If

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET
    Dim bt As Byte
Dim ii As Long

'For ii = btASM To pASM - 1
' CpyMem bt, ByVal ii, 1
' Debug.Print Right$("00" + Hex$(bt), 2);
' Next ii
' Debug.Print
 
  CallPointer = CallWindowProcA(btASM, _
                                  0, 0, 0, 0)
            
             VirtualUnlock btASM, MAXCODE


    once = False
End Function
Public Sub ReleaseMem()
If btASM <> 0 Then
        VirtualFree btASM, MAXCODE, MEM_DECOMMIT
        VirtualFree btASM, 0, MEM_RELEASE
End If
End Sub
Private Sub AddPush(pASM As Long, lng As Long)
    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall(pASM As Long, addr As Long)
    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong(pASM As Long, lng As Long)
    CpyMem ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte(pASM As Long, bt As Byte)
    CpyMem ByVal pASM, bt, 1
    pASM = pASM + 1
End Sub
Public Function CallCdecl( _
    ByVal lpfn As Long, _
    ParamArray Args() As Variant _
) As Long

    Dim btASM(&HEC00& - 1)  As Byte
    Dim pASM                As Long
    Dim btArgSize           As Byte
    Dim i                   As Integer

    pASM = VarPtr(btASM(0))

    If UBound(Args) = 0 Then
        If IsArray(Args(0)) Then
            For i = UBound(Args(0)) To 0 Step -1
                AddPush pASM, CLng(Args(0)(i))    ' PUSH dword
                btArgSize = btArgSize + 4
            Next
        Else
            For i = UBound(Args) To 0 Step -1
                AddPush pASM, CLng(Args(i))       ' PUSH dword
                btArgSize = btArgSize + 4
            Next
        End If
    Else
        For i = UBound(Args) To 0 Step -1
            AddPush pASM, CLng(Args(i))           ' PUSH dword
            btArgSize = btArgSize + 4
        Next
    End If

    AddByte pASM, &HB8
    AddLong pASM, lpfn
    AddByte pASM, &HFF
    AddByte pASM, &HD0
    AddByte pASM, &H83
    AddByte pASM, &HC4
    AddByte pASM, btArgSize
    AddByte pASM, &HC2
    AddByte pASM, &H10
    AddByte pASM, &H0

    CallCdecl = CallWindowProcA(VarPtr(btASM(0)), _
                               0, 0, 0, 0)
End Function

