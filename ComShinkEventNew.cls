VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComShinkEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' from ComShinkEvent changed to fit in M2000
'
' event support for Late-Bound objects
' low level COM Projekt - by [rm_code] 2005

'
' HOW IT WORKS
'
' An object, wich supports events,
' should implement IConnectionPointContainer.
' This interface let's you enumerate all
' event interfaces.
' Once you have the event interface you want,
' you can get IConnectionPoint for this interface,
' advise your event sink to it, and start
' recieving events on it.
' If the event interface has IDispatch implemented,
' all raised events will go through
' IDispatch::Invoke.
'
' And that's what we use:
' We just take the first best event interface we
' get from IEnumConnectionPoints, create a
' event sink which implements IDispatch,
' and capture all events through Invoke().
'

Private Type IUnknown
    QueryInterface              As Long
    AddRef                      As Long
    Release                     As Long
End Type

Private Type IDispatch
    iunk                        As IUnknown
    GetTypeInfoCount            As Long
    GetTypeInfo                 As Long
    GetIDsOfNames               As Long
    Invoke                      As Long
End Type

Private Type IConnectionPointContainer
    iunk                        As IUnknown
    EnumConnectionPoints        As Long
    FindConnectionPoint         As Long
End Type

Private Type IConnectionPoint
    iunk                        As IUnknown
    GetConnectionInterface      As Long
    GetConnectionPointContainer As Long
    Advise                      As Long
    Unadvise                    As Long
    EnumConnections             As Long
End Type

Private Type IEnumConnectionPoints
    iunk                        As IUnknown
    Next                        As Long
    skip                        As Long
    Reset                       As Long
    Clone                       As Long
End Type

Private Type ITypeInfo
    iunk                        As IUnknown
    GetTypeAttr                 As Long
    GetTypeComp                 As Long
    GetFuncDesc                 As Long
    GetVarDesc                  As Long
    GetNames                    As Long
    GetRefTypeOfImplType        As Long
    GetImplTypeFlags            As Long
    GetIDsOfNames               As Long
    Invoke                      As Long
    GetDocumentation            As Long
    GetDllEntry                 As Long
    GetRefTypeInfo              As Long
    AddressOfMember             As Long
    CreateInstance              As Long
    GetMops                     As Long
    GetContainingTypeLib        As Long
    ReleaseTypeAttr             As Long
    ReleaseFuncDesc             As Long
    ReleaseVarDesc              As Long
End Type

Private Type ITypeLib
    iunk                        As IUnknown
    GetTypeInfoCount            As Long
    GetTypeInfo                 As Long
    GetTypeInfoType             As Long
    GetTypeInfoOfGuid           As Long
    GetLibAttr                  As Long
    GetTypeComp                 As Long
    GetDocumentation            As Long
    IsName                      As Long
    FindName                    As Long
    ReleaseTLibAttr             As Long
End Type

Private Type DISPPARAMS
    rgPointerToVariantArray     As Long
    rgPointerToLONGNamedArgs    As Long
    cArgs                       As Long
    cNamedArgs                  As Long
End Type

Private Type SAFEARRAYBOUND
    cElements                   As Long
    lLBound                     As Long
End Type

Private Type SAFEARRAY_1D
   cDims                        As Integer
   fFeatures                    As Integer
   cbElements                   As Long
   cLocks                       As Long
   pvData                       As Long
   Bounds(0 To 0)               As SAFEARRAYBOUND
End Type

Private Enum Varenum
    VT_EMPTY = 0
    VT_NULL = 1
    VT_I2 = 2
    VT_I4 = 3
    VT_R4 = 4
    VT_R8 = 5
    VT_CY = 6
    VT_DATE = 7
    VT_BSTR = 8
    VT_DISPATCH = 9
    VT_ERROR = 10
    VT_BOOL = 11
    VT_VARIANT = 12
    VT_UNKNOWN = 13
    VT_DECIMAL = 14
    VT_I1 = 16
    VT_UI1 = 17
    VT_UI2 = 18
    VT_UI4 = 19
    VT_I8 = 20
    VT_UI8 = 21
    VT_INT = 22
    VT_UINT = 23
    VT_VOID = 24
    VT_HRESULT = 25
    VT_PTR = 26
    VT_SAFEARRAY = 27
    VT_CARRAY = 28
    VT_USERDEFINED = 29
    VT_LPSTR = 30
    VT_LPWSTR = 31
    VT_FILETIME = 64
    VT_BLOB = 65
    VT_STREAM = 66
    VT_STORAGE = 67
    VT_STREAMED_OBJECT = 68
    VT_STORED_OBJECT = 69
    VT_BLOB_OBJECT = 70
    VT_CF = 71
    VT_CLSID = 72
    VT_STREAMED_PROPSET = 73
    VT_STORED_PROPSET = 74
    VT_BLOB_PROPSET = 75
    VT_VERBOSE_ENUM = 76
    VT_BSTR_BLOB = &HFFF
    VT_VECTOR = &H1000
    VT_ARRAY = &H2000
    VT_BYREF = &H4000
    VT_RESERVED = &H8000
    VT_ILLEGAL = &HFFFF
    VT_ILLEGALMASKED = &HFFF
    vt_typemask = &HFFF
End Enum

Public Event EventRaised(ByVal strName As String, params() As Variant)

Private oCPC        As IConnectionPointContainer
Private pCPC        As Long
Private pVTblCPC    As Long

Private oECP        As IEnumConnectionPoints
Private pECP        As Long
Private pVTblECP    As Long

Private oCP         As IConnectionPoint
Private pCP         As Long
Private pVTblCP     As Long

Private oUnk        As IUnknown
Private pUnk        As Long
Private pVTblUnk    As Long

Private oTarget     As Object
Private pTarget     As Long
Private pVTblTarget As Long

Private oSink       As Object
Private dwCookie    As Long

Private iid_event   As GUID

Private blnAttached As Boolean
' m2000 connect
Public modulename$  ' module name plus name of object
Public modulenameonly$
Public VarIndex As Long, ItemIndex As Long
' by Edanmo
Private exclude As New LongHash, inUse As New LongHash
Private Sub pvSetParamArray(ByVal Ptr As Long, aParams() As Variant, ref() As Long)
    Dim tDPAR       As DISPPARAMS
    Dim iVarType    As Integer, iVar As Integer
    Dim lIdx        As Long
    Dim lPtr        As Long
    Dim lVarPtr     As Long
    Dim IP          As Long, oldPtr As Long
    Dim Zero
    CpyMem tDPAR, ByVal Ptr, Len(tDPAR)
    lVarPtr = tDPAR.rgPointerToVariantArray

    For lIdx = 0 To tDPAR.cArgs - 1
            IP = tDPAR.cArgs - lIdx
            CpyMem lPtr, ByVal lVarPtr + (lIdx * 16&), 2
            
            CpyMem iVarType, ByVal VarPtr(aParams(IP)), 2&
            Debug.Print "aparam Type:", iVarType
            CpyMem lPtr, ByVal lVarPtr + (lIdx * 16&) + 8&, 4&
            CpyMem iVar, ByVal lVarPtr + (lIdx * 16&), 2
            iVarType = iVarType And &HFFF
            If ref(IP) Then
            Debug.Print "old final ptr:", lPtr, " from ", lVarPtr + (lIdx * 16&)
            lPtr = ref(IP)
            If iVarType = 12 Then
              CpyMem ByVal lPtr, ByVal VarPtr(aParams(IP)), 16
            ElseIf iVarType = 14 Then
                CpyMem ByVal lPtr + 2, ByVal VarPtr(aParams(IP)) + 2, 14
            ElseIf iVarType = vbDouble Then
            Debug.Print lPtr, "<-", VarPtr(aParams(IP)) + 8
            CpyMem ByVal lPtr, ByVal VarPtr(aParams(IP)) + 8, 8
            Else
            Debug.Print "StrPtr", StrPtr(aParams(IP))
            Debug.Print lPtr, "<-", VarPtr(aParams(IP)) + 8, "pointer only"
            CpyMem oldPtr, ByVal lPtr, 4
            Debug.Print "oldPtr ", oldPtr
            CpyMem ByVal lPtr, ByVal VarPtr(aParams(IP)) + 8&, 8&
            ''CpyMem lPtr, ByVal lVarPtr + (lIdx * 16&) + 8&, 8&
            ''Debug.Print "NewPtr ", lPtr
            ''CpyMem lPtr, ByVal lPtr, 4&
            ''Debug.Print "NewPtr deref ", lPtr
            End If
            iVarType = iVarType + &H4000
            'CpyMem ByVal VarPtr(iVar), ByVal lVarPtr + (lIdx * 16&), 2
            'Debug.Print "final before Type:", iVarType
            
            CpyMem ByVal lVarPtr + (lIdx * 16&), VarPtr(iVarType), 2
            'Debug.Print "final after Type:", iVarType
            CpyMem aParams(IP), Zero, 16

            End If
    Next
End Sub
Private Sub pvSetParamArray1(ByVal Ptr As Long, aParams() As Variant, ref() As Long)
    Dim tDPAR       As DISPPARAMS
    Dim iVarType    As Long
    Dim lIdx        As Long
    Dim lPtr        As Long
    Dim lVarPtr     As Long
    Dim IP          As Long
    Dim Zero
    CpyMem tDPAR, ByVal Ptr, Len(tDPAR)
    lVarPtr = tDPAR.rgPointerToVariantArray

    For lIdx = 0 To tDPAR.cArgs - 1
            IP = tDPAR.cArgs - lIdx
            CpyMem iVarType, ByVal VarPtr(aParams(IP)), 2&
            CpyMem lPtr, ByVal lVarPtr + (lIdx * 16&) + 8&, 4&
            iVarType = iVarType And &HBFFF&
            If ref(IP) Then
            lPtr = ref(IP)
            If iVarType = 12 Or lPtr < 0 Then
                If lPtr < 0 Then
                    CpyMem ByVal -lPtr, ByVal VarPtr(aParams(IP)), 16
                Else
                    CpyMem ByVal lPtr, ByVal VarPtr(aParams(IP)), 16
                End If
            ElseIf iVarType = 14 Then
                CpyMem ByVal lPtr + 2, ByVal VarPtr(aParams(IP)) + 2, 14
            Else
                Select Case iVarType
                Case vbDouble, vbCurrency, vbDate
                    CpyMem ByVal lPtr, ByVal VarPtr(aParams(IP)) + 8, 8
                Case vbBoolean, vbInteger
                    CpyMem ByVal lPtr, ByVal VarPtr(aParams(IP)) + 8, 2
                Case vbByte
                    CpyMem ByVal lPtr, ByVal VarPtr(aParams(IP)) + 8, 1
                Case vbEmpty, vbNull
                Case Else
                    CpyMem ByVal lPtr, ByVal VarPtr(aParams(IP)) + 8, 4
                End Select
            End If
            If lPtr >= 0 Then
            iVarType = iVarType + &H4000
                   
            CpyMem ByVal lVarPtr + (lIdx * 16&), VarPtr(iVarType), 2
            End If
            CpyMem aParams(IP), Zero, 16

            End If
    Next
End Sub

Private Function pvGetParamArray(ByVal Ptr As Long, aParams() As Variant, isRefAr() As Long) As Long
    Dim tDPAR           As DISPPARAMS
    Dim lIdx            As Long
    Dim lVarPtr         As Long
    CpyMem tDPAR, ByVal Ptr, Len(tDPAR)
    lVarPtr = tDPAR.rgPointerToVariantArray
    
    If tDPAR.cArgs = 0 Then Exit Function
    Dim isRef As Boolean
    ReDim aParams(1 To tDPAR.cArgs)
    ReDim isRefAr(1 To tDPAR.cArgs) As Long
    For lIdx = 1 To tDPAR.cArgs
        Process VarPtr(aParams(lIdx)), lVarPtr + ((tDPAR.cArgs - lIdx) * 16&), isRefAr(lIdx)
     Next
    pvGetParamArray = tDPAR.cArgs
End Function
Private Function GetMemberName(obj As Object, ByVal dispid As Long, IID As GUID, first$) As String
    Dim oTypeLib    As ITypeLib
    Dim pTypeLib    As Long
    Dim pVTblTpLib  As Long

    Dim oTypeInfo   As ITypeInfo
    Dim pTypeInfo   As Long
    Dim pVTblTpInfo As Long

    Dim oDispatch   As IDispatch
    Dim hRet        As Long
    Dim dwIndex     As Long
    Dim pcNames     As Long
    Dim pVTbl       As Long

    Dim strName     As String
    ' get IDispatche from the object
    pVTbl = ObjPtr(obj)
    CpyMem pVTbl, ByVal pVTbl, 4
    CpyMem oDispatch, ByVal pVTbl, Len(oDispatch)

    ' get ITypeInfo
    hRet = CallPointer(oDispatch.GetTypeInfo, ObjPtr(obj), 0, LCID_DEF, VarPtr(pTypeInfo))
    If hRet Then GoTo finalize

    ' ITypeInfo VTable
    CpyMem pVTblTpInfo, ByVal pTypeInfo, 4
    CpyMem oTypeInfo, ByVal pVTblTpInfo, Len(oTypeInfo)

    ' let's first try to get the name
    ' of the member by using the current TypeInfo
    hRet = CallPointer(oTypeInfo.GetNames, pTypeInfo, dispid, VarPtr(strName), 1, VarPtr(pcNames))
    If Len(strName) > 0 Then first$ = strName$: strName = vbNullString Else first$ = "val"
    hRet = CallPointer(oTypeInfo.GetContainingTypeLib, pTypeInfo, VarPtr(pTypeLib), VarPtr(dwIndex))
    If hRet Then GoTo finalize
    If first$ = "val" Then first$ = "value"
    ' ITypeLib VTable
    CpyMem pVTblTpLib, ByVal pTypeLib, 4
    CpyMem oTypeLib, ByVal pVTblTpLib, Len(oTypeLib)

    ' GetTypeInfoOfGUID
    hRet = CallPointer(oTypeLib.GetTypeInfoOfGuid, pTypeLib, VarPtr(iid_event), VarPtr(pTypeInfo))
    If hRet Then GoTo finalize


    ' ITypeInfo VTable
    CpyMem pVTblTpInfo, ByVal pTypeInfo, 4
    CpyMem oTypeInfo, ByVal pVTblTpInfo, Len(oTypeInfo)

    ' GetNames
    hRet = CallPointer(oTypeInfo.GetNames, pTypeInfo, dispid, VarPtr(strName), 1, VarPtr(pcNames))

finalize:
    If Len(strName) = 0 Then
        ' no string... :(
        ' instead return the dispip
        If first$ <> "" Then
            GetMemberName = first$
             first$ = vbNullString
        Else
            GetMemberName = LTrim$(Str$(dispid))
            first$ = LTrim$(Str$(dispid))
        End If
    Else
        If Len(first$) = 0 Then first$ = LTrim$(Str$(dispid))
        GetMemberName = strName

        End If
End Function

Public Sub fireevent(ByVal dispid As Long, ByVal params As Long)
' change to fire events, when also othet events are fired.
    Static once As Boolean
    If once Then Exit Sub
    If exclude.ExistKey(dispid) Then Exit Sub
    once = True
    Dim strEvent    As String
    Dim vParams()   As Variant
    Dim NumVar As Long, what$, delme As Boolean
    ' get the name of the event
    If inUse.ExistKey(dispid) Then
        strEvent = inUse.Value
        what$ = inUse.Info
    Else
        strEvent = GetMemberName(oTarget, dispid, iid_event, what$)
        inUse.AddKey dispid, strEvent, what$
    End If
 
    Dim refc() As Long
    NumVar = pvGetParamArray(params, vParams, refc)
    On Error Resume Next
   
    If Not CallEventFromCOM(Me, strEvent, what$, NumVar, vParams, refc, delme, ItemIndex) Then
       MyEr "Problem With Event: " + strEvent, "�������� �� �� �������: " + strEvent
    End If
    If delme Then exclude.AddKey dispid, ""
    pvSetParamArray params, vParams, refc
    once = False
End Sub

Public Sub Detach()
    Dim hRet    As Long

    If Not blnAttached Then Exit Sub

    ' clean up
    hRet = CallPointer(oCP.Unadvise, pCP, dwCookie)
    hRet = CallPointer(oCP.iunk.Release, pCP)
    hRet = CallPointer(oECP.iunk.Release, pECP)
    hRet = CallPointer(oCPC.iunk.Release, pCPC)
    Set oTarget = Nothing

    blnAttached = False
End Sub

Public Function Attach(obj As Object) As Boolean
    Dim cReturned   As Long
    Dim hRet        As Long
    Dim iid_ICP     As GUID

    ' already connected to an object?
    If blnAttached Then Detach

    ' IUnknown VTable
    Set oTarget = obj
    pTarget = ObjPtr(obj)
    CpyMem pVTblTarget, ByVal pTarget, 4   ' make Vtbl pointer from object pointer
    CpyMem oUnk, ByVal pVTblTarget, Len(oUnk)

    ' IID string -> GUID struct
    hRet = CLSIDFromString(StrPtr(IIDSTR_IConnectionPointContainer), iid_ICP)
    If hRet Then Exit Function

    ' get IConnectionPointContainer
    hRet = CallPointer(oUnk.QueryInterface, pTarget, VarPtr(iid_ICP), VarPtr(pCPC))
    If hRet Then Exit Function

    ' IConnectionPointContainer VTable
    CpyMem pVTblCPC, ByVal pCPC, 4
    CpyMem oCPC, ByVal pVTblCPC, Len(oCPC)

    ' get IEnumConnectionPoints
    hRet = CallPointer(oCPC.EnumConnectionPoints, pCPC, VarPtr(pECP))
    If hRet Then Exit Function

    ' IEnumConnectionPoints VTable
    CpyMem pVTblECP, ByVal pECP, 4
    CpyMem oECP, ByVal pVTblECP, Len(oECP)

    ' take the first best Connection Point
    hRet = CallPointer(oECP.Next, pECP, 1, VarPtr(pCP), VarPtr(cReturned))
    If hRet Then Exit Function

    ' IConnectionPoint VTable
    CpyMem pVTblCP, ByVal pCP, 4
    CpyMem oCP, ByVal pVTblCP, Len(oCP)

    ' IID of the event interface
    hRet = CallPointer(oCP.GetConnectionInterface, pCP, VarPtr(iid_event))
    If hRet Then Exit Function

    ' create a new event sink
    Set oSink = CreateEventSink(iid_event, Me)

    ' advise the event sink
    hRet = CallPointer(oCP.Advise, pCP, ObjPtr(oSink), VarPtr(dwCookie))
    If hRet Then Exit Function

    ' wohoo, done!
    blnAttached = True
    Attach = True
End Function

Private Sub Class_Initialize()
    InitObjExtender
    ItemIndex = -1
End Sub

Private Sub Class_Terminate()
Detach
End Sub
Public Function Attached() As Boolean
    Attached = blnAttached
End Function
Sub Process(VarProc As Long, a As Long, Ptr1 As Long)
    Dim t(0 To 8) As Integer, Zero, Ptr As Long, ItsVariantRef As Boolean
    CpyMem ByVal VarPtr(t(0)), ByVal a, 16
again:
    CpyMem Ptr, ByVal VarPtr(t(4)), 4

    If (t(0) And &H4000) > 0 Then
        Ptr1 = Ptr
        t(0) = CInt(CLng(t(0)) And &HBFFF&)
        If t(0) <> 14 And t(0) <> 12 Then
            t(1) = 0: t(2) = 0: t(3) = 0
            Select Case t(0)
            Case vbDouble, vbCurrency, vbDate
                CpyMem ByVal VarPtr(t(4)), ByVal Ptr, 8
            Case vbBoolean, vbInteger
                CpyMem ByVal VarPtr(t(4)), ByVal Ptr, 2
            Case vbByte
                CpyMem ByVal VarPtr(t(4)), ByVal Ptr, 1
            Case vbEmpty, vbNull
            ' do nothing
            Case Else
                CpyMem ByVal VarPtr(t(4)), ByVal Ptr, 4
            End Select
        ElseIf t(0) = 12 Then
            CpyMem ByVal VarPtr(t(0)), ByVal Ptr, 16
            ItsVariantRef = True
            If (t(0) And &H4000 > 0) Then GoTo again
        Else
            CpyMem ByVal VarPtr(t(1)), ByVal Ptr + 2, 14
        End If
    ElseIf t(0) = 12 Then
        Ptr1 = Ptr
        CpyMem ByVal VarPtr(t(0)), ByVal Ptr, 16
    Else
        If Not ItsVariantRef Then
            Ptr1 = 0
            
            CpyMem ByVal a, VarPtr(Zero), 16
        Else
            Ptr1 = -Ptr1
            CpyMem ByVal a, VarPtr(Zero), 16
        End If
    End If
    CpyMem ByVal VarProc, ByVal VarPtr(t(0)), 16
    
End Sub


