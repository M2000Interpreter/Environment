Attribute VB_Name = "Fcall"
' This is a module from Olaf Schmidt changed for M2000 needs
Option Explicit
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal callconv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef RETVAR As Variant) As Long
Private Declare Function GetProcByName Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal nOrdinal As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Dst As Any, Src As Any, ByVal BLen As Long)
Private Declare Function SysStringLen Lib "oleaut32" (ByVal bstr As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Private Enum CALLINGCONVENTION_ENUM
  cc_fastcall
  CC_CDECL
  CC_PASCAL
  CC_MACPASCAL
  CC_STDCALL
  CC_FPFASTCALL
  CC_SYSCALL
  CC_MPWCDECL
  CC_MPWPASCAL
End Enum

Private LibHdls As New FastCollection, vType(0 To 63) As Integer, vPtr(0 To 63) As Long
Sub HandleStringInBuffer(a As MemBlock)

End Sub
Public Sub SetLibHdls()
    LibHdls.UcaseKeys = True
End Sub

Public Function stdCallW(sDLL As String, sFunc As String, ByVal RetType As Variant, p() As Variant, j As Long)
Dim v(), HRes As Long, i As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To j - 1 ''UBound(V)
    If VarType(p(i)) = vbString Then
    v(i) = CLng(StrPtr(p(i)))
    vPtr(i) = VarPtr(v(i))
    vType(i) = vbString
    Else
    vType(i) = VarType(v(i))
    vPtr(i) = VarPtr(v(i))
    End If
    
  Next i
  If Left$(sFunc, 1) = "#" Then
  HRes = DispCallFunc(0, GetFuncPtrOrd(sDLL, sFunc), CC_STDCALL, CInt(RetType), j, vType(0), vPtr(0), stdCallW)
  Else
  HRes = DispCallFunc(0, GetFuncPtr(sDLL, sFunc), CC_STDCALL, CInt(RetType), j, vType(0), vPtr(0), stdCallW)
  End If
  If HRes Then Err.Raise HRes
' p() = v()
 If VarType(stdCallW) = vbNull Then
        stdCallW = vbEmpty
 End If
End Function
Public Function Fast_stdCallW(ByVal addr As Long, ByVal RetType As Variant, p() As Variant, j As Long)
Dim v(), HRes As Long, i As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To j - 1 ''UBound(V)
    If VarType(p(i)) = vbString Then
    v(i) = CLng(StrPtr(p(i)))
    vPtr(i) = VarPtr(v(i))
    vType(i) = vbString
    Else
    vType(i) = VarType(v(i))
    vPtr(i) = VarPtr(v(i))
    End If
    
  Next i

  HRes = DispCallFunc(0, addr, CC_STDCALL, CInt(RetType), j, vType(0), vPtr(0), Fast_stdCallW)

  If HRes Then Err.Raise HRes
' p() = v()
 If VarType(Fast_stdCallW) = vbNull Then
    Fast_stdCallW = vbEmpty
 End If
End Function
Public Function Fast_obj_stdCallW(obj As stdole.IUnknown, ByVal addroffset As Long, ByVal RetType As Variant, p() As Variant, j As Long)
Dim v(), HRes As Long, i As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To j - 1 ''UBound(V)
    If VarType(p(i)) = vbString Then
    v(i) = CLng(StrPtr(p(i)))
    vPtr(i) = VarPtr(v(i))
    vType(i) = vbString
    Else
    vType(i) = VarType(v(i))
    vPtr(i) = VarPtr(v(i))
    End If
    
  Next i

  HRes = DispCallFunc(ObjPtr(obj), addroffset, CC_STDCALL, CInt(RetType), j, vType(0), vPtr(0), Fast_obj_stdCallW)

  If HRes Then Err.Raise HRes
' p() = v()
 If VarType(Fast_obj_stdCallW) = vbNull Then
    Fast_obj_stdCallW = vbEmpty
 End If
End Function
Public Function cdeclCallW(sDLL As String, sFunc As String, ByVal RetType As Variant, p() As Variant, j As Long)
Dim i As Long, pFunc As Long, v(), HRes As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To j - 1
    If VarType(p(i)) = vbString Then v(i) = StrPtr(p(i))
    vType(i) = VarType(v(i))
    vPtr(i) = VarPtr(v(i))
  Next i
   If Left$(sFunc, 1) = "#" Then
     HRes = DispCallFunc(0, GetFuncPtrOrd(sDLL, sFunc), CC_CDECL, CInt(RetType), j, vType(0), vPtr(0), cdeclCallW)
   Else
  HRes = DispCallFunc(0, GetFuncPtr(sDLL, sFunc), CC_CDECL, CInt(RetType), j, vType(0), vPtr(0), cdeclCallW)
  End If
  If HRes Then Err.Raise HRes
  If VarType(cdeclCallW) = vbNull Then
    cdeclCallW = vbEmpty
  End If
End Function

Public Function Fast_cdeclCallW(ByVal addr, ByVal RetType As Variant, p() As Variant, j As Long)
Dim i As Long, pFunc As Long, v(), HRes As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To j - 1
    If VarType(p(i)) = vbString Then v(i) = StrPtr(p(i))
    vType(i) = VarType(v(i))
    vPtr(i) = VarPtr(v(i))
  Next i

  HRes = DispCallFunc(0, addr, CC_CDECL, CInt(RetType), j, vType(0), vPtr(0), Fast_cdeclCallW)

  If HRes Then Err.Raise HRes
  If VarType(Fast_cdeclCallW) = vbNull Then
  Fast_cdeclCallW = vbEmpty
  End If
End Function


Public Function stdCallA(sDLL As String, sFunc As String, ByVal RetType As Variant, ParamArray p() As Variant)
Dim i As Long, pFunc As Long, v(), HRes As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To UBound(v)
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbFromUnicode): v(i) = StrPtr(p(i))
    vType(i) = VarType(v(i))
    vPtr(i) = VarPtr(v(i))
  Next i
  
  HRes = DispCallFunc(0, GetFuncPtr(sDLL, sFunc), CC_STDCALL, RetType, i, vType(0), vPtr(0), stdCallA)
  
  For i = 0 To UBound(p) 'back-conversion of the ANSI-String-Results
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbUnicode)
  Next i
  If HRes Then Err.Raise HRes
End Function

Public Function cdeclCallA(sDLL As String, sFunc As String, ByVal RetType As VbVarType, ParamArray p() As Variant)
Dim i As Long, pFunc As Long, v(), HRes As Long
 
  v = p 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
  For i = 0 To UBound(v)
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbFromUnicode): v(i) = StrPtr(p(i))
    vType(i) = VarType(v(i))
    vPtr(i) = VarPtr(v(i))
  Next i
  
  HRes = DispCallFunc(0, GetFuncPtr(sDLL, sFunc), CC_CDECL, RetType, i, vType(0), vPtr(0), cdeclCallA)
  
  For i = 0 To UBound(p) 'back-conversion of the ANSI-String-Results
    If VarType(p(i)) = vbString Then p(i) = StrConv(p(i), vbUnicode)
  Next i
  If HRes Then Err.Raise HRes
End Function

Public Function vtblCall(pUnk As Long, ByVal vtblIdx As Long, ParamArray p() As Variant)
Dim i As Long, v(), HRes As Long
  If pUnk = 0 Then Exit Function

  v = p 'make a copy of the params, to prevent problems with VT_ByRef-Members in the ParamArray
  For i = 0 To UBound(v)
    vType(i) = VarType(v(i))
    vPtr(i) = VarPtr(v(i))
  Next i
  
  HRes = DispCallFunc(pUnk, vtblIdx * 4, CC_STDCALL, vbLong, i, vType(0), vPtr(0), vtblCall)
  If HRes Then Err.Raise HRes
End Function

Public Function GetFuncPtr(sLib As String, sFunc As String) As Long

Dim hLib As Long

    If LibHdls.Find(sLib) Then
        hLib = LibHdls.Value
    Else
     
      hLib = LoadLibrary(StrPtr(sLib))
      If hLib = 0 Then Err.Raise vbObjectError, , "Dll not found (or loadable): " & sLib
      LibHdls.AddKey sLib, hLib
    End If
  'End If
  GetFuncPtr = GetProcByName(hLib, sFunc)
  If GetFuncPtr = 0 Then MyEr "EntryPoint not found: " + sFunc + " in: " + sLib, "EntryPoint not found: " + sFunc + " στο: " + sLib
End Function
Public Sub RemoveDll(ByVal sLib As String, Optional noErr As Boolean)
Dim v As Long, s As String
s = ExtractPath(sLib)
If s = "" Then s = mcd + sLib
If LibHdls.Find(sLib) Then
    If FreeLibrary(LibHdls.Value) <> 0 Then
        LibHdls.RemoveWithNoFind
    Else
    v = GetLastError()
    If Not noErr Then MyEr "η βιβλιοθήκη δεν μπορεί να αφαιρεθεί, κωδικός λάθους:(" & v & ")", "dll not removes, error code:(" & v & ")"
    
    End If
ElseIf LibHdls.Find(s) Then
If FreeLibrary(LibHdls.Value) <> 0 Then
        LibHdls.RemoveWithNoFind
    Else
    v = GetLastError()
    If Not noErr Then MyEr "η βιβλιοθήκη δεν μπορεί να αφαιρεθεί, κωδικός λάθους:(" & v & ")", "dll not removes, error code:(" & v & ")"
    
    End If
Else
MyEr "δεν υπάρχει η βιβλιοθήκη", "dll not found"
End If
End Sub

Public Function GetFuncPtrOrd(sLib As String, sFunc As String) As Long
Dim hLib As Long
Dim lfunc As Long

lfunc = val(Mid$(sFunc, 2))

    If LibHdls.Find(sLib) Then
        hLib = LibHdls.Value
    Else
      hLib = LoadLibrary(StrPtr(sLib))
      If hLib = 0 Then Err.Raise vbObjectError, , "Dll not found (or loadable): " & sLib
      LibHdls.AddKey sLib, hLib
    End If
   ' End If
  GetFuncPtrOrd = GetProcByOrdinal(hLib, lfunc)
  If GetFuncPtrOrd = 0 Then MyEr "EntryPoint not found: " + sFunc + " in: " + sLib, "EntryPoint not found: " + sFunc + " στο: " + sLib
End Function
Public Function GetBStrFromBstrPtr(lpSrc As Long) As String
Dim slen As Long
  If lpSrc = 0 Then Exit Function
  slen = SysStringLen(lpSrc)
  If slen Then GetBStrFromBstrPtr = space$(slen) Else Exit Function

  RtlMoveMemory ByVal StrPtr(GetBStrFromBstrPtr), ByVal lpSrc, slen * 2

End Function
Public Function GetBStrFromPtr(lpSrc As Long, Optional ByVal ANSI As Boolean) As String
Dim slen As Long
  If lpSrc = 0 Then Exit Function
  If ANSI Then slen = lstrlenA(lpSrc) Else slen = lstrlenW(lpSrc)
  If slen Then GetBStrFromPtr = space$(slen) Else Exit Function
      
  Select Case ANSI
    Case True: RtlMoveMemory ByVal GetBStrFromPtr, ByVal lpSrc, slen
    Case Else: RtlMoveMemory ByVal StrPtr(GetBStrFromPtr), ByVal lpSrc, slen * 2
  End Select
End Function
Public Sub CleanupLibHandles() 'not really needed - but callable (usually at process-shutdown) to clear things up
Dim LibHdl
LibHdls.ToStart
While LibHdls.Done
    FreeLibrary LibHdls.Value
    LibHdls.NextIndex
Wend
'  For Each LibHdl In LibHdls: FreeLibrary LibHdl: Next
  Set LibHdls = Nothing
End Sub
Function IsWine() As Boolean
Static www As Boolean, wwb As Boolean
If www Then
Else
Err.Clear
Dim hLib As Long, ntdll As String
On Error Resume Next
ntdll = "ntdll"
hLib = LoadLibrary(StrPtr(ntdll))
wwb = GetProcByName(hLib, "wine_get_version") <> 0
If hLib <> 0 Then FreeLibrary hLib
If Err.Number > 0 Then wwb = False
www = True
End If
IsWine = wwb
End Function

