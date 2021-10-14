Attribute VB_Name = "Module2"
'
' For working with ActiveX.dll libraries without registration.
' Krivous Anatolii Anatolevich (The trick), 2015.
' Cut down by Elroy, 2016.
'
Option Explicit
'
Private Type GUID
    data1       As Long
    data2       As Integer
    data3       As Integer
    data4(7)    As Byte
End Type
'
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszCLSID As Long, ByRef clsid As GUID) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef src As Any, ByRef dst As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Any, ByVal oVft As Long, ByVal cc As Integer, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Any, ByRef prgpvarg As Any, ByRef pvargResult As Variant) As Long
Private Declare Function LoadTypeLibEx Lib "oleaut32" (ByVal szFile As Long, ByVal regkind As Long, ByRef pptlib As IUnknown) As Long
Private Declare Function memcpy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
'
Private Const IID_IClassFactory   As String = "{00000001-0000-0000-C000-000000000046}"
Private Const IID_IUnknown        As String = "{00000000-0000-0000-C000-000000000046}"
Private Const CC_STDCALL          As Long = 4
Private Const REGKIND_NONE        As Long = 2
Private Const TKIND_COCLASS       As Long = 5
'
Private iidClsFctr      As GUID
Private iidUnk          As GUID
Private isInit          As Boolean
'

Public Function NewObjectFromActivexDll(ByRef pathToDll As String, _
                                        ByRef className As String) As IUnknown
    ' Create object by Name.
    ' Uses DLL as the TLB.
    '
    Set NewObjectFromActivexDll = CreateObjectEx2(pathToDll, pathToDll, className)
End Function

Public Sub UnloadActivexDll(ByRef path As String)
    ' Unload DLL if not used.
    '
    Dim hLib    As Long
    Dim lpAddr  As Long
    Dim ret     As Long
    Dim spot    As Long
    '
    spot = 1
    If isInit Then
        spot = 2
        hLib = GetModuleHandle(StrPtr(path))
        If hLib <> 0 Then
            spot = 3
            lpAddr = GetProcAddress(hLib, "DllCanUnloadNow")
            If lpAddr <> 0 Then
                spot = 4
                ret = DllCanUnloadNow(lpAddr)
                If ret = 0 Then
                    FreeLibrary hLib
                    Exit Sub
                End If
            End If
        End If
    End If
    If Not bCompiled Then MsgBox "Didn't unload " & Chr$(34) & path & Chr$(34) & " but got to spot " & Format(spot) & "."
End Sub

'******************************************************************************
'
' Private from here down.
'
'******************************************************************************
'
Private Function CreateObjectEx2(ByRef pathToDll As String, _
                                 ByRef pathToTLB As String, _
                                 ByRef className As String) As IUnknown
    ' Create object by Name.
    ' The DLL can be used as the TLB with VB6 ActiveX.DLL files.
    '
    Dim typeLib As IUnknown
    Dim typeInf As IUnknown
    Dim ret     As Long
    Dim pAttr   As Long
    Dim tKind   As Long
    Dim clsid   As GUID
    '
    ret = LoadTypeLibEx(StrPtr(pathToTLB), REGKIND_NONE, typeLib)
    If ret Then
        Err.Raise ret
        Exit Function
    End If
    ret = ITypeLib_FindName(typeLib, className, 0, typeInf, 0, 1)
    If typeInf Is Nothing Then
        Err.Raise &H80040111, , "Class not found in type library"
        Exit Function
    End If
    ITypeInfo_GetTypeAttr typeInf, pAttr
    GetMem4 ByVal pAttr + &H28, tKind
    If tKind = TKIND_COCLASS Then
        memcpy clsid, ByVal pAttr, Len(clsid)
    Else
        Err.Raise &H80040111, , "Class not found in type library"
        Exit Function
    End If
    ITypeInfo_ReleaseTypeAttr typeInf, pAttr
    Set CreateObjectEx2 = CreateObjectEx1(pathToDll, clsid)
End Function

Private Function CreateObjectEx1(ByRef path As String, _
                                ByRef clsid As GUID) As IUnknown
    ' Create object by CLSID and path.
    '
    Dim hLib    As Long
    Dim lpAddr  As Long
    Dim isLoad  As Boolean
    Dim ret     As Long
    Dim out     As IUnknown
    '
    hLib = GetModuleHandle(StrPtr(path))
    If hLib = 0 Then
        hLib = LoadLibrary(StrPtr(path))
        If hLib = 0 Then
            Err.Raise 53, , Error(53) & " " & Chr$(34) & path & Chr$(34)
            Exit Function
        End If
        isLoad = True
    End If
    lpAddr = GetProcAddress(hLib, "DllGetClassObject")
    If lpAddr = 0 Then
        If isLoad Then FreeLibrary hLib
        Err.Raise 453, , "Can't find dll entry point DllGetClasesObject in " & Chr$(34) & path & Chr$(34)
        Exit Function
    End If
    If Not isInit Then
        CLSIDFromString StrPtr(IID_IClassFactory), iidClsFctr
        CLSIDFromString StrPtr(IID_IUnknown), iidUnk
        isInit = True
    End If
    ret = DllGetClassObject(lpAddr, clsid, iidClsFctr, out)
    If ret = 0 Then
        ret = IClassFactory_CreateInstance(out, 0, iidUnk, CreateObjectEx1)
    Else
        If isLoad Then FreeLibrary hLib
        Err.Raise ret
        Exit Function
    End If
    Set out = Nothing
    If ret Then
        If isLoad Then FreeLibrary hLib
        Err.Raise ret
    End If
End Function

Private Function DllGetClassObject(ByVal funcAddr As Long, _
                                   ByRef clsid As GUID, _
                                   ByRef iid As GUID, _
                                   ByRef out As IUnknown) As Long
    ' Call "DllGetClassObject" function using a pointer.
    '
    Dim params(2)   As Variant
    Dim types(2)    As Integer
    Dim list(2)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    '
    params(0) = VarPtr(clsid)
    params(1) = VarPtr(iid)
    params(2) = VarPtr(out)
    '
    For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    resultCall = DispCallFunc(0&, funcAddr, CC_STDCALL, vbLong, 3, types(0), list(0), pReturn)
    If resultCall Then Err.Raise 5: Exit Function
    DllGetClassObject = pReturn
End Function

Private Function DllCanUnloadNow(ByVal funcAddr As Long) As Long
    ' Call "DllCanUnloadNow" function using a pointer.
    '
    Dim resultCall  As Long
    Dim pReturn     As Variant
    '
    resultCall = DispCallFunc(0&, funcAddr, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, pReturn)
    If resultCall Then Err.Raise 5: Exit Function
    DllCanUnloadNow = pReturn
End Function

Private Function IClassFactory_CreateInstance(ByVal obj As IUnknown, _
                                              ByVal punkOuter As Long, _
                                              ByRef riid As GUID, _
                                              ByRef out As IUnknown) As Long
    ' Call "IClassFactory:CreateInstance" method.
    '
    Dim params(2)   As Variant
    Dim types(2)    As Integer
    Dim list(2)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    '
    params(0) = punkOuter
    params(1) = VarPtr(riid)
    params(2) = VarPtr(out)
    '
    For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    resultCall = DispCallFunc(obj, &HC, CC_STDCALL, vbLong, 3, types(0), list(0), pReturn)
    If resultCall Then Err.Raise resultCall: Exit Function
    IClassFactory_CreateInstance = pReturn
End Function

Private Function ITypeLib_FindName(ByVal obj As IUnknown, _
                                   ByRef szNameBuf As String, _
                                   ByVal lHashVal As Long, _
                                   ByRef ppTInfo As IUnknown, _
                                   ByRef rgMemId As Long, _
                                   ByRef pcFound As Integer) As Long
    ' Call "ITypeLib:FindName" method.
    '
    Dim params(4)   As Variant
    Dim types(4)    As Integer
    Dim list(4)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    '
    params(0) = StrPtr(szNameBuf)
    params(1) = lHashVal
    params(2) = VarPtr(ppTInfo)
    params(3) = VarPtr(rgMemId)
    params(4) = VarPtr(pcFound)
    '
    For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    resultCall = DispCallFunc(obj, &H2C, CC_STDCALL, vbLong, 5, types(0), list(0), pReturn)
    If resultCall Then Err.Raise resultCall: Exit Function
    ITypeLib_FindName = pReturn
End Function

Private Sub ITypeInfo_GetTypeAttr(ByVal obj As IUnknown, _
                                  ByRef ppTypeAttr As Long)
    ' Call "ITypeInfo:GetTypeAttr" method.
    '
    Dim resultCall  As Long
    Dim pReturn     As Variant
    '
    pReturn = VarPtr(ppTypeAttr)
    resultCall = DispCallFunc(obj, &HC, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(pReturn), 0)
    If resultCall Then Err.Raise resultCall: Exit Sub
End Sub

Private Sub ITypeInfo_ReleaseTypeAttr(ByVal obj As IUnknown, _
                                      ByVal ppTypeAttr As Long)
    ' Call "ITypeInfo:ReleaseTypeAttr" method.
    '
    Dim resultCall  As Long
    '
    resultCall = DispCallFunc(obj, &H4C, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(CVar(ppTypeAttr)), 0)
    If resultCall Then Err.Raise resultCall: Exit Sub
End Sub

Private Function bCompiled() As Boolean
    On Error GoTo Errored
    Debug.Print 1 / 0
    bCompiled = True
Errored:
End Function

