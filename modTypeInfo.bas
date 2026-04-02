Attribute VB_Name = "modTypeInfo"
Option Explicit
Private Declare Function lstrlenA Lib "KERNEL32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "KERNEL32" (ByVal lpString As Long) As Long
Private Declare Sub RtlMoveMemory Lib "KERNEL32" (dst As Any, src As Any, ByVal BLen As Long)

' modTypeInfo - enumerate object members and get member infos
'
' low level COM project by [rm] 2005
' parameter description
Private Type TPARAMDESC
    pPARAMDESCEX            As Long     ' valid if PARAMFLAG_FHASDEFAULT
    wParamFlags             As Integer  ' parameter flags (in,out,...)
End Type

' extended parameter description
Private Type TPARAMDESCEX
    cBytes                  As Long     ' size of structure
    varDefaultValue         As Variant  ' default value of parameter
End Type
Private Type TTYPEDESC
    pTypeDesc               As Long     ' vt = VT_PTR: points to another TYPEDESC
                                        ' vt = VT_CARRAY: points to another TYPEDESC
                                        ' vt = VT_USERDEFINED: pTypeDesc is a HREFTYPE instead of a pointer
    vt                      As Integer  ' vartype
End Type
Private Type TELEMDESC
    tdesc                   As TTYPEDESC    ' type description
    pdesc                   As TPARAMDESC   ' parameter description
End Type

Private Type TYPEATTR
    GUID(15)                As Byte
    tLCID                   As Long
    dwReserved              As Long
    memidConstructor        As Long
    memidDestructor         As Long
    pstrSchema              As Long
    cbSizeInstance          As Long
    typekind                As Long
    cFuncs                  As Integer
    cVars                   As Integer
    cImplTypes              As Integer
    cbSizeVft               As Integer
    cbAlignment             As Integer
    wTypeFlags              As Integer
    wMajorVerNum            As Integer
    wMinorVerNum            As Integer
    tdescAlias              As Long
    idldescType             As Long
End Type

Private Type FUNCDESC
    memid                   As Long
    lprgscode               As Long
    lprgelemdescParam       As Long
    funcking                As Long
    invkind                 As Long
    callconv                As Long
    cParams                 As Integer
    cParamsOpt              As Integer
    oVft                    As Integer
    cScodes                 As Integer
    elemdesc                As TELEMDESC ' Contains the return type of the function
    wFuncFlags              As Integer  ' function flags
End Type

' array description
Private Type TARRAYDESC
    tdescElem               As TTYPEDESC    ' type description
    cDims                   As Integer      ' number of dimensions
End Type
Private Type SAFEARRAYBOUND
    cElements               As Long
    lLBound                 As Long
End Type

Public Enum VARKIND
    VAR_PERSISTANCE = 0             '
    VAR_STATIC                      '
    VAR_CONST                       '
    VAR_DISPATCH                    '
End Enum

Private Type VARDESC
    memid                   As Long     ' member ID
    lpstrSchema             As Long     '
    uInstVal                As Long     ' vkind = VAR_PERINSTANCE: offset of this variable within the instance
                                        ' vkind = VAR_CONST: value of it as a variant
    elemdescVar             As TELEMDESC ' variable type
    wVarFlags               As Integer  ' variable flags
    vkind                   As Long     ' variable kind
End Type

' parameter flags
Public Enum PARAMFLAGS
    PARAMFLAG_NONE = &H0            ' ...
    PARAMFLAG_FIN = &H1             ' in
    PARAMFLAG_FOUT = &H2            ' out
    PARAMFLAG_FLCID = &H4           ' lcid
    PARAMFLAG_FRETVAL = &H8         ' return value
    PARAMFLAG_FOPT = &H10           ' optional
    PARAMFLAG_FHASDEFAULT = &H20    ' default value
    PARAMFLAG_FHASCUSTDATA = &H40   ' custom data
End Enum

Public Type fncinf
    name                    As String
    addr                    As Long
    params                  As Integer
End Type

Public Type enmeinf
    name                    As String
    invkind                 As invokekind
    params                  As Integer
    
End Type
Public Const DISP_E_PARAMNOTFOUND = &H80020004
Public Enum Varenum
    VT_EMPTY = 0&                   '
    VT_NULL = 1&                    ' 0
    VT_I2 = 2&                      ' signed 2 bytes integer
    VT_I4 = 3&                      ' signed 4 bytes integer
    VT_R4 = 4&                      ' 4 bytes float
    VT_R8 = 5&                      ' 8 bytes float
    VT_CY = 6&                      ' currency
    VT_DATE = 7&                    ' date
    VT_BSTR = 8&                    ' BStr
    VT_DISPATCH = 9&                ' IDispatch
    VT_ERROR = 10&                  ' error value
    VT_BOOL = 11&                   ' boolean
    VT_VARIANT = 12&                ' variant
    VT_UNKNOWN = 13&                ' IUnknown
    VT_DECIMAL = 14&                ' decimal
    VT_I1 = 16&                     ' signed byte
    VT_UI1 = 17&                    ' unsigned byte
    VT_UI2 = 18&                    ' unsigned 2 bytes integer
    VT_UI4 = 19&                    ' unsigned 4 bytes integer
    VT_I8 = 20&                     ' signed 8 bytes integer
    VT_UI8 = 21&                    ' unsigned 8 bytes integer
    VT_INT = 22&                    ' integer
    VT_UINT = 23&                   ' unsigned integer
    VT_VOID = 24&                   ' 0
    VT_HRESULT = 25&                ' HRESULT
    VT_PTR = 26&                    ' pointer
    VT_SAFEARRAY = 27&              ' safearray
    VT_CARRAY = 28&                 ' carray
    VT_USERDEFINED = 29&            ' userdefined
    VT_LPSTR = 30&                  ' LPStr
    VT_LPWSTR = 31&                 ' LPWStr
    VT_RECORD = 36&                 ' Record
    VT_FILETIME = 64&               ' File Time
    VT_BLOB = 65&                   ' Blob
    VT_STREAM = 66&                 ' Stream
    VT_STORAGE = 67&                ' Storage
    VT_STREAMED_OBJECT = 68&        ' Streamed Obj
    VT_STORED_OBJECT = 69&          ' Stored Obj
    VT_BLOB_OBJECT = 70&            ' Blob Obj
    VT_CF = 71&                     ' CF
    VT_CLSID = 72&                  ' Class ID
    VT_BSTR_BLOB = &HFFF&           ' BStr Blob
    VT_VECTOR = &H1000&             ' Vector
    VT_ARRAY = &H2000&              ' Array
    VT_BYREF = &H4000&              ' ByRef
    VT_RESERVED = &H8000&           ' Reserved
    VT_ILLEGAL = &HFFFF&            ' illegal
End Enum

Private Declare Sub CpyMem Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal dwLen As Long)


Private Declare Sub SysFreeString Lib "oleaut32" ( _
    ByVal bstr As Long)

Public Enum invokekind
    INVOKE_FUNC = &H1
    INVOKE_PROPERTY_GET = &H2
    INVOKE_PROPERTY_PUT = &H4
    INVOKE_PROPERTY_PUTREF = &H8
End Enum

Public Function GetObjMembers(mList As FastCollection, obj As Object) As Long
    Dim vtblObj         As Long, ret            As Long
    Dim vtblTpInf       As Long, vtblTpInfV(21) As Long
    Dim ppTInfo         As Long, rgBstrNames    As Long
    Dim ppFuncDesc      As Long, fncdsc         As FUNCDESC
    Dim pAttr           As Long, attr           As TYPEATTR
    Dim NameLen         As Long, strName        As String
    Dim pGetTpInf       As Long
    Dim cFnc            As Integer
    Dim oInfo           As Object
    Dim iunk            As IUnknown
    Set mList = New FastCollection
    Dim strNames()      As String
    Dim cFncs           As Long, i As Long
    ' VTable of passed object
    vtblObj = ObjPtr(obj)
    CpyMem vtblObj, ByVal vtblObj, 4
    CpyMem pGetTpInf, ByVal vtblObj + 4 * 4, 4

    ' IDispatch->GetTypeInfo
    ret = CallPointer(pGetTpInf, ObjPtr(obj), 0, LCID_DEF, VarPtr(ppTInfo))
    If ret Then Exit Function

    CpyMem oInfo, ppTInfo, 4
    Set iunk = oInfo
    CpyMem oInfo, 0&, 4

    ' VTable of ITypeInfo
    CpyMem vtblTpInf, ByVal ppTInfo, 4
    ' IUnknown(3) + ITypeInfo(19) = 22
    CpyMem vtblTpInfV(0), ByVal vtblTpInf, 22 * 4

    ' ITypeInfo->GetTypeAttr
  
    ret = CallPointer(vtblTpInfV(3), ppTInfo, VarPtr(pAttr))
    If ret Then Exit Function

    ' get TypeAttributes struct
      
      
    CpyMem attr, ByVal pAttr, Len(attr)

    ' ITypeInfo->ReleaseTypeAttr
    
    ret = CallPointer(vtblTpInfV(19), ppTInfo, VarPtr(pAttr))
    If ret Then
    
    Debug.Print "Couldn't release TypeAttr"
    
    End If

    ' go through all members
    For cFnc = 0 To attr.cFuncs - 1

        ' ITypeInfo->GetFuncDesc
        ret = CallPointer(vtblTpInfV(5), ppTInfo, cFnc, VarPtr(ppFuncDesc))
        If ret Then GoTo NextItem

        ' read function descriptor struct
        CpyMem fncdsc, ByVal ppFuncDesc, Len(fncdsc)

        ' ITypeInfo->ReleaseFuncDesc
        ret = CallPointer(vtblTpInfV(20), ppTInfo, ppFuncDesc)

        ' ITypeInfo->GetNames for current member
        ret = CallPointer(vtblTpInfV(12), ppTInfo, fncdsc.memid, VarPtr(rgBstrNames), 0, 0, 0)
        If ret Then GoTo NextItem

        ' read its name (Unicode)
        CpyMem NameLen, ByVal rgBstrNames - 4, 4
        strName = Space$(NameLen / 2)
        CpyMem ByVal StrPtr(strName), ByVal rgBstrNames, NameLen
        SysFreeString rgBstrNames
        mList.AddKey UCase(strName), ""
        Select Case fncdsc.invkind
            Case INVOKE_FUNC:
                strName = "Function " + strName
            Case INVOKE_PROPERTY_GET:
                strName = "Property Get " + strName
            Case INVOKE_PROPERTY_PUT:
                strName = "Property Let " + strName
            Case INVOKE_PROPERTY_PUTREF:
                strName = "Property Set " + strName
        End Select
        mList.ToEnd  ' move to last
        
        If fncdsc.cParams > 0 Then
        cFncs = 0
        ReDim strNames(fncdsc.cParams) As String
' GetNames offset 7 (long)
        ret = CallPointer(vtblTpInfV(7), ppTInfo, fncdsc.memid, VarPtr(strNames(0)), 1 + fncdsc.cParams, VarPtr(cFncs))
        If Not ret Then
         strName = strName + "("
            For i = 1 To fncdsc.cParams
            strName = strName + strNames(i)
            If i < fncdsc.cParams Then strName = strName + ", "
            Next i
        strName = strName + ")"
        End If
        End If
        mList.Value = strName
        
        
        
       
NextItem:
    Next
    









    GetObjMembers = True

    Set iunk = Nothing
End Function

Public Function GetFncInfo(obj As Object, fnc As String) As fncinf
    Dim vtblObj         As Long, ret            As Long
    Dim vtblTpInf       As Long, vtblTpInfV(21) As Long
    Dim ppTInfo         As Long, rgBstrNames    As Long
    Dim ppFuncDesc      As Long, fncdsc         As FUNCDESC
    Dim pAttr           As Long, attr           As TYPEATTR
    Dim NameLen         As Long, strName        As String
    Dim pGetTpInf       As Long
    Dim cFnc            As Integer
    Dim oInfo           As Object
    Dim iunk            As IUnknown
    ' this can be used to call a function using Callpointer mFncinf.addr,  ObjPtr(cSay), parameter
    ' but need typelib
    
    
    ' VTable of passed object
    vtblObj = ObjPtr(obj)
    CpyMem vtblObj, ByVal vtblObj, 4
    CpyMem pGetTpInf, ByVal vtblObj + 4 * 4, 4

    ' IDispatch->GetTypeInfo
    ret = CallPointer(pGetTpInf, ObjPtr(obj), 0, LCID_DEF, VarPtr(ppTInfo))
    If ret Then Exit Function

    CpyMem oInfo, ppTInfo, 4
    Set iunk = oInfo
    CpyMem oInfo, 0&, 4

    ' VTable of ITypeInfo
    CpyMem vtblTpInf, ByVal ppTInfo, 4
    ' IUnknown(3) + ITypeInfo(19) = 22
    CpyMem vtblTpInfV(0), ByVal vtblTpInf, 22 * 4

    ' ITypeInfo->GetTypeAttr
    ret = CallPointer(vtblTpInfV(3), ppTInfo, VarPtr(pAttr))
    If ret Then Exit Function

    ' get TypeAttributes struct
    CpyMem attr, ByVal pAttr, Len(attr)

    ' ITypeInfo->ReleaseTypeAttr
    ret = CallPointer(vtblTpInfV(19), ppTInfo, VarPtr(pAttr))

    ' go through all members
    For cFnc = 0 To attr.cFuncs - 1

        ' ITypeInfo->GetFuncDesc
        ret = CallPointer(vtblTpInfV(5), ppTInfo, cFnc, VarPtr(ppFuncDesc))
        If ret Then GoTo NextItem

        ' read function descriptor struct
        CpyMem fncdsc, ByVal ppFuncDesc, Len(fncdsc)

        ' ITypeInfo->ReleaseFuncDesc
        ret = CallPointer(vtblTpInfV(20), ppTInfo, ppFuncDesc)

        ' ITypeInfo->GetNames for current member
        Dim mmname$
        mmname$ = Space$(200)
          ret = CallPointer(vtblTpInfV(12), ppTInfo, fncdsc.memid, StrPtr(mmname$), 0, 0, 0)
        If ret Then GoTo NextItem
strName = GetBStrFromPtr(VarPtr(mmname$))

    
        If StrComp(strName, fnc, vbTextCompare) = 0 Then
            With GetFncInfo
                .name = strName
                .params = fncdsc.cParams
                CpyMem .addr, ByVal vtblObj + fncdsc.oVft, 4
                Exit For
            End With
        End If

NextItem:
    Next

    Set iunk = Nothing
End Function
Private Function GetBStrFromPtr(lpSrc As Long, Optional ByVal ANSI As Boolean) As String
Dim SLen As Long
  If lpSrc = 0 Then Exit Function
  If ANSI Then SLen = lstrlenA(lpSrc) Else SLen = lstrlenW(lpSrc)
  If SLen Then GetBStrFromPtr = Space$(SLen) Else Exit Function
      
  Select Case ANSI
    Case True: RtlMoveMemory ByVal GetBStrFromPtr, ByVal lpSrc, SLen
    Case Else: RtlMoveMemory ByVal StrPtr(GetBStrFromPtr), ByVal lpSrc, SLen * 2
  End Select
End Function
Public Function VTableEntry(obj As Object, ByVal entry As Integer) As Long
    Dim pVTbl       As Long
    Dim lngEntry    As Long

    pVTbl = ObjPtr(obj)
    CpyMem pVTbl, ByVal pVTbl, 4

    CpyMem lngEntry, ByVal pVTbl + &H1C + entry * 4 - 4, 4
    VTableEntry = lngEntry
End Function

