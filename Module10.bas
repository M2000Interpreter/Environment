Attribute VB_Name = "Module10"
Option Explicit
' some utilities for String$(  string as converter)
' some code from dilettante - vbforums
' http://www.vbforums.com/showthread.php?342995-VB6-URL-Path-String-Manipulation-Functions
' removed all InStrB. strings expexted to be as UTF16LE (as strings in VB6)
' if we have an ascii string then we have to convert it before use it
' also if we have a UTF8 string we have to convert it before use it here
Private Type PeekArrayType
    Ptr         As Long
    Reserved    As Currency
End Type
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal callconv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef RETVAR As Variant) As Long
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
Private Declare Function PeekArray Lib "kernel32" Alias "RtlMoveMemory" (Arr() As Any, Optional ByVal Length As Long = 4) As PeekArrayType
Private Declare Function SafeArrayGetDim Lib "OleAut32.dll" (ByVal Ptr As Long) As Long
Private Declare Function vbaVarLateMemSt Lib "msvbvm60" _
                         Alias "__vbaVarLateMemSt" ( _
                         ByRef vDst As Variant, _
                         ByRef sName As Any, _
                         ByVal vValue As Variant) As Long
Private Declare Function vbaVarLateMemCallLdRf CDecl Lib "msvbvm60" _
                         Alias "__vbaVarLateMemCallLdRf" ( _
                         ByVal vDst As Long, _
                         ByVal vSrc As Long, _
                         ByVal sName As Long, _
                         ByVal cArgs As Long, _
                         ByVal vArgs As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, retval As Long)
Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = 0
Private Const INTERNET_MAX_URL_LENGTH As Long = 2083
Private Const URL_ESCAPE_PERCENT As Long = &H1000&
Private Const URL_PART_SCHEME As Long = 1
Private Const URL_PART_HOSTNAME As Long = 2
Private Const URL_PART_USERNAME As Long = 3
Private Const URL_PART_PASSWORD As Long = 4
Private Const URL_PART_PORT As Long = 5
Private Const URL_PART_QUERY As Long = 6
Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeW" ( _
    ByVal pszURL As Long, _
    ByVal pszEscaped As Long, _
    ByRef pcchEscaped As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeW" ( _
    ByVal pszURL As Long, _
    ByVal pszUnescaped As Long, _
    ByRef pcchUnescaped As Long, _
    ByVal dwFlags As Long) As Long
Private Const CONST_HOSTNAME = 2
Private Declare Function UrlGetPart Lib "shlwapi" Alias "UrlGetPartW" ( _
    ByVal pszIn As Long, _
    ByVal pszOut As Long, _
    pcchOut As Long, _
    ByVal dwPart As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function UrlCanonicalizeApi Lib "shlwapi" Alias "UrlCanonicalizeW" ( _
    ByVal pszURL As Long, _
    ByVal pszCanonicalized As Long, _
    pcchCanonicalized As Long, _
    ByVal dwFlags As Long) As Long
Const NO_ERROR = 0&
Const MOVEFILE_REPLACE_EXISTING = &H1
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_BEGIN = 0
Const FILE_CURRENT = 1
Const FILE_END = 2
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const FILE_APPEND_DATA = &H4
Const FILE_READ_ATTRIBUTES As Long = &H80&
Const CREATE_NEW = 1
Const CREATE_ALWAYS = 2
Const OPEN_EXISTING = 3
Const OPEN_ALLWAYS = 4
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const ERROR_SHARING_VIOLATION = 32&
Const INVALID_HANDLE_VALUE = (-1&)
Const INVALID_SET_FILE_POINTER = (-1&)
Const FILE_ATTRIBUTE_NORMAL = &H80
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private InUseHandlers As New FastCollection
Private FreeUseHandlers As New FastCollection
Enum MyOpenFileType
    ForInput = 1  ' Read  - Begin
    ForOutput = 2 ' write - End
    ForAppend = 3 ' Read+Write - End
    ForField = 4 ' Read+Write - Begin
End Enum
Enum MyOpenFileExclusive
    ExclusiveON = 0  ' Lock at open from 0 to
    ExclusiveOff = 1 ' not Lock
End Enum
Public FileError As Long
Public Enum Ftypes
    FnoUse = 0
    Finput = 1
    Foutput = 2
    Fappend = 3
    Frandom = 4
End Enum
Sub FileReadString(FileH As Long, r$, bytes As Long)
    Dim Buf1() As Byte
    If bytes <= 0 Then r$ = vbNullString: Exit Sub
    If FileH = 0 Then Exit Sub
    ReDim Buf1(0 To bytes - 1)
    r$ = Buf1()
    API_ReadFile FileH, bytes, Buf1()
    CopyMemory ByVal StrPtr(r$), Buf1(0), bytes
End Sub
Sub FileReadBytes(FileH As Long, Buf1() As Byte, bytes As Long)
    If bytes <= 0 Then Exit Sub
    If FileH = 0 Then Exit Sub
    API_ReadFile FileH, bytes, Buf1
End Sub
Sub FileWriteString(FileH As Long, r$)
    Dim Buf1() As Byte, bytes As Long
    If LenB(r$) = 0 Then Exit Sub
    If FileH = 0 Then Exit Sub
    bytes = LenB(r$)
    ReDim Buf1(bytes - 1)
    CopyMemory Buf1(0), ByVal StrPtr(r$), bytes
    API_WriteFile FileH, bytes, Buf1()
End Sub
Sub FileWriteBytes(FileH As Long, ByRef Buf1() As Byte)
    Dim bytes As Long
    If PeekArray(Buf1()).Ptr = 0 Then Exit Sub
    If FileH = 0 Then Exit Sub
    bytes = UBound(Buf1()) - LBound(Buf1()) + 1
    API_WriteFile FileH, bytes, Buf1()
End Sub
Private Function GetType(bstack As basetask, b$, p, v As Long, W$, Lang As Long, VarStat As Boolean, temphere$, noVarStat As Boolean) As Integer
Dim ss$, skip As Boolean, checktype As Boolean
    If noVarStat Then
        If GetVar(bstack, W$, v, , , True, , checktype) Then
           skip = True
           On Error Resume Next
        End If
    End If
    If IsLabelSymbolNew(b$, "ΑΡΙΘΜΟΣ", "DECIMAL", Lang) Then
            If FastSymbol(b$, "=") Then
                If Not IsNumberD2(b$, p) Then missNumber: Exit Function
            ElseIf skip Then
                p = var(v)
            End If
            p = CDec(p)
            If Err.Number Then p = CDec(0)
    ElseIf IsLabelSymbolNew(b$, "ΔΙΠΛΟΣ", "DOUBLE", Lang) Then
            If FastSymbol(b$, "=") Then
            If Not IsNumberD2(b$, p) Then missNumber: Exit Function
            
            ElseIf skip Then
                p = var(v)
            End If
            p = CDbl(p)
            If Err.Number Then p = CDbl(0)
    ElseIf IsLabelSymbolNew(b$, "ΑΠΛΟΣ", "SINGLE", Lang) Then
            If FastSymbol(b$, "=") Then
            If Not IsNumberD2(b$, p) Then missNumber: Exit Function
            
            ElseIf skip Then
                p = var(v)
            End If
            p = CSng(p)
            If Err.Number Then p = CSng(0)
    ElseIf IsLabelSymbolNew(b$, "ΛΟΓΙΚΟΣ", "BOOLEAN", Lang) Then
            If FastSymbol(b$, "=") Then
            If Not IsNumberD2(b$, p) Then missNumber: Exit Function
            
            ElseIf skip Then
                p = var(v)
            End If
            p = CBool(p)
            If Err.Number Then p = False
    ElseIf IsLabelSymbolNew(b$, "ΜΑΚΡΥΣ", "LONG", Lang) Then
        If IsLabelSymbolNew(b$, "ΜΑΚΡΥΣ", "LONG", Lang) Then
            If FastSymbol(b$, "=") Then
            If Not IsNumberD2(b$, p) Then missNumber: Exit Function
            
            ElseIf skip Then
                p = var(v)
            End If
            p = cInt64(p)
            If Err.Number Then p = cInt64(CDec(0))
        Else
            If FastSymbol(b$, "=") Then
            If Not IsNumberD2(b$, p) Then missNumber: Exit Function
            
            ElseIf skip Then
                p = var(v)
            End If
            p = CLng(p)
            If Err.Number Then p = CLng(0)
        End If
    ElseIf IsLabelSymbolNew(b$, "ΑΚΕΡΑΙΟΣ", "INTEGER", Lang) Then
        If FastSymbol(b$, "=") Then
        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
        
        ElseIf skip Then
                p = var(v)
        End If
        p = CInt(p)
        If Err.Number Then p = CInt(0)
    ElseIf IsLabelSymbolNew(b$, "ΛΟΓΙΣΤΙΚΟΣ", "CURRENCY", Lang) Then
        If FastSymbol(b$, "=") Then
        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
        
        ElseIf skip Then
                p = var(v)
        End If
        p = CCur(p)
    ElseIf IsLabelSymbolNew(b$, "ΓΡΑΜΜΑ", "STRING", Lang) Then
        If FastSymbol(b$, "=") Then
        If Not ISSTRINGA(b$, ss$) Then MissString: Exit Function
        
        ElseIf skip Then
               If MemInt(VarType(v)) = vbString Then
                    ss$ = var(v)
                Else
                    ss$ = ""
                End If
        End If
        p = vbNullString
        SwapString2Variant ss$, p
    ElseIf IsLabelSymbolNew(b$, "ΑΤΥΠΟΣ", "VARIANT", Lang) Then
        If FastSymbol(b$, "=") Then
            If Not IsNumberD2(b$, p) Then
                If ISSTRINGA(b$, ss$) Then
                    p = vbNullString
                    SwapString2Variant ss$, p
                Else
                    missNumber
                    Exit Function
                End If
            End If
        End If
        v = globalvar(W$, p, , VarStat, temphere$, useType:=False)
        If extreme Then GetType = 2 Else GetType = 1
        Exit Function
    ElseIf IsLabelSymbolNew(b$, "ΨΗΦΙΟ", "BYTE", Lang) Then
        If FastSymbol(b$, "=") Then
        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
        
        ElseIf skip Then
                p = var(v)
        End If
        p = CByte(p)
        If Err.Number Then p = CByte(0)
    ElseIf IsLabelSymbolNew(b$, "ΗΜΕΡΟΜΗΝΙΑ", "DATE", Lang) Then
        If FastSymbol(b$, "=") Then
            If Not IsNumberD2(b$, p) Then
                If ISSTRINGA(b$, ss$) Then
                    p = vbNullString
                    SwapString2Variant ss$, p
                Else
                    missNumber
                    Exit Function
                End If
            Else
                
                
            End If
        ElseIf skip Then
            p = var(v)
        End If
        p = CDate(p)
        If Err.Number Then p = CDate(0)
    ElseIf IsLabelSymbolNew(b$, "ΜΥΓΑΔΙΚΟΣ", "COMPLEX", Lang) Then
            If FastSymbol(b$, "=") Then
                p = nMath2.cxZero
                If FastSymbol(b$, "(") Then
                Dim p1
                If Not IsNumberD2(b$, p1) Then
                    GoTo ER0001
                End If
                p.r = CDbl(p1)
                If Not FastSymbol(b$, ",") Then GoTo ER0001
                If Not IsNumberD2(b$, p1) Then
                    GoTo ER0001
                End If
                p.I = CDbl(p1)
                b$ = NLtrim(b$)
                If Not UCase$(Left$(b$, 2)) = "I)" Then GoTo ER0001
                Mid$(b$, 1, 2) = "  "
            Else
                GoTo ER0001
            End If
            Else
                p = nMath2.cxZero
            End If
    ElseIf IsLabelSymbolNew(b$, "ΜΕΓΑΛΟΣΑΚΕΡΑΙΟΣ", "BIGINTEGER", Lang) Then
        Set p = New BigInteger
        If FastSymbol(b$, "=") Then
        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
            
        End If
        If MemInt(VarPtr(p)) = vbString Then
            Set p = Module13.CreateBigInteger(CStr(p))
        End If
    ElseIf Not IsEnumAs(bstack, b$, p) Then
            ExpectedEnumType
            Exit Function
    End If

    If noVarStat Then
    If skip Then GoTo there1
        If Not GetVar(bstack, W$, v, , , True, , checktype) Then
            v = globalvar(W$, p, , VarStat, temphere$)
        Else
there1:
            If Not checktype Then
ER0001:
                WrongType
                Exit Function
            End If
            If MyIsObject(p) Then
                If Not p Is Nothing Then
                    If TypeOf p Is BigInteger Then
                        Set var(v) = p
                        GoTo abcd
                    End If
                End If
                bstack.soros.PushObj p
                If Not MyRead(5, bstack, W$, Lang, "") Then
                    Exit Function
                End If
            Else
                var(v) = p
            End If
        End If
    Else
        v = globalvar(W$, p, , VarStat, temphere$)
    End If
abcd:
    Err.Clear
    If extreme Then GetType = 2 Else GetType = 1
End Function

' We have to provide a new file created before
' in any case
' 0 is error.
' non 0 is the filehandler from M2000 (not the real handler)
Function MyOpenFile(ByVal f$, ftype As MyOpenFileType, fexc As MyOpenFileExclusive, fstp As Long, unif As Long) As Long
If Left$(f$, 2) <> "\\" Then
f$ = "\\?\" + f$
End If
Dim FileH As Long
FileError = 511
On Error GoTo there:
Select Case ftype
Case ForInput
    FileH = CreateFile(StrPtr(f$), GENERIC_READ, (FILE_SHARE_READ Or FILE_SHARE_WRITE) * fexc, ByVal 0&, OPEN_EXISTING, 0&, 0&)
Case ForOutput
    FileH = CreateFile(StrPtr(f$), GENERIC_WRITE, (FILE_SHARE_READ Or FILE_SHARE_WRITE) * fexc, ByVal 0&, CREATE_ALWAYS, 0&, 0&)
Case Else
    FileH = CreateFile(StrPtr(f$), GENERIC_READ Or GENERIC_WRITE, (FILE_SHARE_READ Or FILE_SHARE_WRITE) * fexc, ByVal 0&, OPEN_EXISTING, 0&, 0&)
End Select
If FileH = INVALID_HANDLE_VALUE Then
FileError = GetLastError()
 If FileError = ERROR_SHARING_VIOLATION Then
    MyEr "Can't open file, Not Sharing allowed" & FileError, "Δεν μπορώ να ανοίξω το αρχείο, δεν επιτρέπεται μοίρασμα"
    Exit Function
 End If
there:
    MyEr "Can't open file, error :" & FileError, "Δεν μπορώ να ανοίξω το αρχείο :" & FileError
    Exit Function
End If
' So now we get the filehandler as 1
MyOpenFile = BigFileHandler(CVar(FileH), ftype, fstp, unif)
Select Case ftype
Case ForOutput, ForAppend
    SetFilePointer FileH, 0&, 0&, FILE_END
Case Else
    SetFilePointer FileH, 0&, 0&, FILE_BEGIN
End Select
End Function
Private Function proc101(b$) As String
        proc101 = vbNullString
        If FastSymbol(b$, "+=", , 2) Then
        proc101 = "+"
        ElseIf FastSymbol(b$, "/=", , 2) Then
        proc101 = "/"
        ElseIf FastSymbol(b$, "-=", , 2) Then
        proc101 = "-"
        ElseIf FastSymbol(b$, "*=", , 2) Then
        proc101 = "*"
        ElseIf IsOperator0(b$, "++", 2) Then
        proc101 = "++"
        ElseIf IsOperator0(b$, "--", 2) Then
        proc101 = "--"
        ElseIf IsOperator0(b$, "-!", 2) Then
        proc101 = "-!"
        ElseIf IsOperator0(b$, "~") Then
        proc101 = "!!"
        ElseIf FastSymbol(b$, "<=", , 2) Then
        proc101 = "g"
        End If
End Function

Function procObject(bstack As basetask, W$, p, v As Long, useType As Boolean, VarStat As Boolean, isglobal As Boolean, NewStat As Boolean) As Boolean
Dim myobject As Object, usehandler As mHandler, usehandler1 As mHandler
        Set myobject = bstack.lastobj
        If TypeOf bstack.lastobj Is Group Then ' oh is a group
            Set bstack.lastobj = Nothing
            If myobject.IamApointer Then
                Set var(v) = myobject
            Else
                If useType Then
                    myobject.ToDelete = True
                    UnFloatGroup bstack, W$, v, myobject, VarStat Or isglobal, , myVarType(var(v), vbEmpty)        ' global??
                    If Len(bstack.UseGroupname) <> 0 Then
                        var(v).IamRef = True
                        If Not (VarStat Or isglobal) Then
                            globalvar W$, CVar(v), True, True
                        End If
                    End If
                Else
                    'Set p = myobject
                    MakeGroupPointer bstack, CVar(myobject) ' p
                    Set var(v) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                    Set bstack.lastpointer = Nothing
                End If
            End If
            Set myobject = Nothing
        ElseIf CheckIsmArray(myobject) Then
            Set usehandler = New mHandler
            Set var(v) = usehandler
            usehandler.T1 = 3
            Set usehandler.objref = myobject
            If TypeOf bstack.lastobj Is mHandler Then
                Set usehandler1 = bstack.lastobj
                With usehandler1
                    If .UseIterator Then
                        usehandler.UseIterator = True
                        usehandler.index_start = .index_start
                        usehandler.index_End = .index_End
                        usehandler.index_cursor = .index_cursor
                    End If
                End With
            End If
            Set usehandler = Nothing
            Set usehandler1 = Nothing
        ElseIf TypeOf myobject Is mHandler Then
            Set usehandler = myobject
            If usehandler.indirect > -1 Then
                If MyIsObject(var(usehandler.indirect)) Then
                    ' we pass an indirect handler (Static in module)
                    ' as a non static in var(v), so we can return it, but why???
                    Set var(v) = var(usehandler.indirect)
                Else
                    BadObjectDecl
                    Exit Function
                End If
            Else
                Set var(v) = usehandler
                If usehandler.T1 = 4 Then
                    If MemInt(VarPtr(usehandler.index_cursor)) = vbString Then
                    
                    Else
                        If usehandler.sign * usehandler.index_cursor <> p Then usehandler.sign = -usehandler.sign
                    End If
                End If
            End If
            If TypeOf bstack.lastobj Is mHandler Then
                Set usehandler1 = bstack.lastobj
                If VarTypeName(var(v)) = "mHandler" Then
                    Set usehandler = var(v)
                    With usehandler1
                        If .UseIterator Then
                            usehandler.UseIterator = True
                            usehandler.index_start = .index_start
                            usehandler.index_End = .index_End
                            usehandler.index_cursor = .index_cursor
                        End If
                    End With
                    Set usehandler = Nothing
                End If
                Set usehandler1 = Nothing
            End If
            Set bstack.lastobj = Nothing
        ElseIf TypeOf myobject Is lambda Then
            If funid.Find(W$ + "(", (0)) Then
                funid.ItemCreator W$ + "(", -2
            End If
            If here$ = vbNullString Or VarStat Or NewStat Then
                GlobalSub W$ + "()", "", , , v
            Else
                GlobalSub here$ + "." + bstack.GroupName + W$ + "()", "", , , v
            End If
            Set var(v) = myobject
            Set bstack.lastobj = Nothing
        ElseIf TypeOf myobject Is mEvent Then
            Set var(v) = myobject
            CopyEvent var(v), bstack
            Set var(v) = bstack.lastobj
        ElseIf TypeOf myobject Is BigInteger Then
            Set bstack.lastobj = Nothing
            Set var(v) = CopyBigInteger(myobject)
        ElseIf VarType(var(v)) = vbEmpty Then
            Set var(v) = myobject
        Else
            Set myobject = Nothing
            Set bstack.lastobj = Nothing
            If VarType(var(v)) = vbLong Then
                NoObjectpAssignTolong
            ElseIf VarType(var(v)) = vbInteger Then
                NoObjectpAssignToInteger
            Else
                NoObjectAssign
                If MyIsNumeric(var(v)) Then
                If var(v) = vbEmpty Then var(v) = 0#
                End If
            End If
            Exit Function 'GoTo err000
        End If
        Set bstack.lastpointer = Nothing
        Set bstack.lastobj = Nothing
        Set myobject = Nothing
        procObject = True
End Function

Private Function readvarv(v, ss$, p) As Boolean
    Dim sp
    readvarv = True
    Select Case ss$
    Case "DIV", "ΔΙΑ"
        v = Fix(v / p)
    Case "DIV#", "ΔΙΑ#"
        If p < 0 Then
            v = Int((v - Abs(v - Abs(p) * Int(v / Abs(p)))) / p)
        Else
            v = Int(v / p)
        End If
    Case "MOD", "ΥΠΟΛ", "ΥΠΟΛΟΙΠΟ"
        sp = v - Fix(v / p) * p
        If Abs(sp) >= Abs(p) Then sp = sp - sp
        v = sp
    Case "MOD#", "ΥΠΟΛ#", "ΥΠΟΛΟΙΠΟ#"
        sp = Abs(v - Abs(p) * Int(v / Abs(p)))
        If Abs(sp) >= Abs(p) Then sp = sp - sp
        v = sp
    Case Else
        readvarv = False
    End Select
End Function

Private Function readvarvLong(v, ss$, p) As Boolean
    Dim sp
    readvarvLong = True
    Select Case ss$
    Case "DIV", "ΔΙΑ"
     var(v) = CLng(Fix(var(v) / p))
    Case "DIV#", "ΔΙΑ#"
        If p < 0 Then
            var(v) = CLng((var(v) - Abs(var(v) - Abs(p) * Int(var(v) / Abs(p)))) / p)
        Else
            var(v) = CLng(var(v) / p)
        End If
    Case "MOD", "ΥΠΟΛ", "ΥΠΟΛΟΙΠΟ"
        sp = var(v) - Fix(var(v) / p) * p
        If Abs(sp) >= Abs(p) Then sp = sp - sp
        var(v) = CLng(sp)
    Case "MOD#", "ΥΠΟΛ#", "ΥΠΟΛΟΙΠΟ#"
        sp = Abs(var(v) - Abs(p) * Int(var(v) / Abs(p)))
        If Abs(sp) >= Abs(p) Then sp = sp - sp
        var(v) = CLng(sp)
    Case Else
        readvarvLong = False
    End Select
End Function


Public Property Get uni(RHS) As Long
Dim h
If InUseHandlers.ExistKey(RHS) Then
    h = InUseHandlers.Value
    uni = h(3)
End If
End Property

Public Property Get Fstep(RHS) As Long
Dim h
If InUseHandlers.ExistKey(RHS) Then
    h = InUseHandlers.Value
    Fstep = h(2)
End If
End Property
Public Property Get Fkind(RHS) As Ftypes
Dim h
If InUseHandlers.ExistKey(RHS) Then
    h = InUseHandlers.Value
    Fkind = h(1)
End If
End Property
Public Property Let FileSeek(RHS, vvv)
Dim h, where, ret As Currency, LowLong As Long, HighLong As Long
Dim FileError As Long
ret = CCur(Int(vvv)) - 1
If InUseHandlers.ExistKey(RHS) Then
    h = InUseHandlers.Value
    where = h(0)
    Size2Long ret, LowLong, HighLong
    LowLong = SetFilePointer(where, LowLong, HighLong, FILE_BEGIN)
    FileError = GetLastError()
    If LowLong = INVALID_SET_FILE_POINTER And FileError <> 0 Then
        MyEr "Can't write the seek value", "Δεν μπορώ να γράψω τη τιμή μετάθεσης"
    End If
Else
    MyEr "No such file handler", "Δεν υπάρχει τέτοιο χειριστής αρχείου"
End If
End Property
Public Property Get FileSeek(RHS) As Variant
Dim h, where, ret As Currency, LowLong As Long, HighLong As Long
FileSeek = ret
If InUseHandlers.ExistKey(RHS) Then
    h = InUseHandlers.Value
    where = h(0)
    LowLong = 0
    HighLong = 0
    LowLong = SetFilePointer(where, LowLong, HighLong, FILE_CURRENT)
    If LowLong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value", "Δεν μπορώ να διαβάσω τη τιμή μετάθεσης"
    Else
        Long2Size LowLong, HighLong, ret
        FileSeek = ret + 1
    End If
Else
    MyEr "No such file handler", "Δεν υπάρχει τέτοιο χειριστής αρχείου"
End If
End Property
Public Property Let FileSeekFH(where As Long, ByVal ret As Currency)
Dim LowLong As Long, HighLong As Long
Dim FileError As Long
    Size2Long ret - 1@, LowLong, HighLong
    LowLong = SetFilePointer(where, LowLong, HighLong, FILE_BEGIN)
    FileError = GetLastError()
    If LowLong = INVALID_SET_FILE_POINTER And FileError <> 0 Then
        MyEr "Can't write the seek value", "Δεν μπορώ να γράψω τη τιμή μετάθεσης"
    End If
End Property
Public Property Get FileSeekFH(where As Long) As Currency
Dim LowLong As Long, HighLong As Long
    LowLong = SetFilePointer(where, LowLong, HighLong, FILE_CURRENT)
    If LowLong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value", "Δεν μπορώ να διαβάσω τη τιμή μετάθεσης"
    Else
        Long2Size LowLong, HighLong, FileSeekFH
        FileSeekFH = FileSeekFH + 1
    End If
End Property
Public Property Get FileEOFFH(where As Long) As Boolean
Dim ret As Currency, LowLong As Long, HighLong As Long
Dim fsize As Currency
    LowLong = SetFilePointer(where, LowLong, HighLong, FILE_CURRENT)
    If LowLong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value on Eof() function", "Δεν μπορώ να διαβάσω τη τιμή μετάθεσης στη συνάρτηση Μετάθεση()"
    Else
        Long2Size LowLong, HighLong, ret
        LowLong = GetFileSize(where, HighLong)
        Long2Size LowLong, HighLong, fsize
        FileEOFFH = ret >= fsize
    End If
End Property

Public Property Get FileEOF(RHS) As Boolean
Dim h, where, ret As Currency, LowLong As Long, HighLong As Long
Dim fsize As Currency
If InUseHandlers.ExistKey(RHS) Then
    h = InUseHandlers.Value
    where = h(0)
    LowLong = SetFilePointer(where, LowLong, HighLong, FILE_BEGIN)
    If LowLong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value on Eof() function", "Δεν μπορώ να διαβάσω τη τιμή μετάθεσης στη συνάρτηση Μετάθεση()"
    Else
        Long2Size LowLong, HighLong, ret
        LowLong = GetFileSize(where, HighLong)
        Long2Size LowLong, HighLong, fsize
        FileEOF = ret >= fsize
    End If
Else
    MyEr "No such file handler", "Δεν υπάρχει τέτοιο χειριστής αρχείου"
End If

End Property
Public Function BigFileHandler(FH As Long, ftype As Long, fst As Long, unif As Long) As Long
Static MaxNum As Long
Dim where As Long
If FreeUseHandlers.count = 0 Then
    MaxNum = MaxNum + 1
    InUseHandlers.AddKey CVar(MaxNum), Array(FH, ftype, fst, unif)
    InUseHandlers.sValue = FH
    BigFileHandler = MaxNum
Else
    FreeUseHandlers.ToEnd
    where = CLng(FreeUseHandlers.KeyToNumber)
    FreeUseHandlers.RemoveWithNoFind
    FreeUseHandlers.Done = False
    InUseHandlers.AddKey CVar(where), Array(FH, ftype, fst, unif)
    InUseHandlers.sValue = FH
    BigFileHandler = where
End If
End Function
'internal use
Public Function ReadFileHandler(h&) As Variant
If InUseHandlers.Find(CVar(h&)) Then
    ReadFileHandler = InUseHandlers.sValue
Else
    MyEr "No such file handler", "Δεν υπάρχει τέτοιο χειριστής αρχείου"
End If
End Function
' internal use  You have to close file first
Public Sub CloseHandler(RHS)
Dim h&, Ar() As Variant
On Error Resume Next
If InUseHandlers.ExistKey(RHS) Then
    h& = CLng(InUseHandlers.sValue)
    API_CloseFile h&
    h& = InUseHandlers.KeyToNumber
    InUseHandlers.RemoveWithNoFind
    FreeUseHandlers.AddKey CVar(h&)
Else
    ' no error... (I am thinking about it)
End If

End Sub
Public Sub CloseAllHandlers()
Dim h&
On Error Resume Next
Do While InUseHandlers.count > 0
    InUseHandlers.ToEnd
    h& = CLng(InUseHandlers.sValue)
    API_CloseFile h&
    h& = InUseHandlers.KeyToNumber
    InUseHandlers.RemoveWithNoFind
    FreeUseHandlers.AddKey CVar(h&)
Loop
End Sub
Function myFileLen(ByVal FileName As String) As Currency
If Left$(FileName, 2) <> "\\" Then
FileName = "\\?\" + FileName
End If
Dim FileH As Long
Dim ret As Long, ok As Long
On Error Resume Next
FileH = CreateFile(StrPtr(FileName), _
                FILE_READ_ATTRIBUTES, _
                 0, _
                ByVal 0&, OPEN_EXISTING, 0&, 0&)

If Err.Number > 0 Or FileH = -1 Then
    Err.Clear
there:
    MyEr "Can't read the file length", "Δεν μπορώ να διαβάσω το μήκος του αρχείου"
    myFileLen = -1@
Else
    ok = API_FileSize(FileH, myFileLen)
    API_CloseFile FileH
    If ok Then GoTo there  ' no zero means error
End If
On Error GoTo 0
End Function
Public Sub API_OpenFile(ByVal FileName As String, ByRef FileNumber As Long, ByRef FileSize As Currency, SetPointerTo As Long)
Dim FileH As Long
Dim ret As Long, ok As Long
On Error Resume Next
FileH = CreateFile(StrPtr(FileName), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_ALLWAYS, 0, 0)
If Err.Number > 0 Then
    Err.Clear
    FileNumber = -1
Else
there:
    FileNumber = FileH
    ret = SetFilePointer(FileH, 0, 0, FILE_BEGIN)
    API_FileSize FileH, FileSize
End If
On Error GoTo 0
End Sub

Function API_FileSize(ByVal FileNumber As Long, ByRef FileSize As Currency) As Long
    Dim FileSizeL As Long
    Dim FileSizeH As Long, ok As Long
    FileSizeH = 0
    FileSizeL = GetFileSize(FileNumber, FileSizeH)
    If FileSizeL = -1& Then
        ok = GetLastError
        If ok Then API_FileSize = ok: Exit Function
    End If
    
    Long2Size FileSizeL, FileSizeH, FileSize
    API_FileSize = 0
End Function

Public Sub API_ReadFile(ByVal FileNumber As Long, ByRef BlockSize As Long, ByRef Data() As Byte)
Dim PosL As Long
Dim PosH As Long
Dim SizeRead As Long
Dim ret As Long
ret = SetFilePointer(FileNumber, PosL, PosH, FILE_CURRENT)
ret = ReadFile(FileNumber, Data(0), BlockSize, SizeRead, 0&)
BlockSize = SizeRead
End Sub
Public Sub API_ReadBLOCK(ByVal FileNumber As Long, ByVal BlockSize As Long, ByVal Addr As Long)
Dim PosL As Long
Dim PosH As Long
Dim SizeRead As Long
Dim ret As Long
ret = SetFilePointer(FileNumber, PosL, PosH, FILE_CURRENT)
ret = ReadFile(FileNumber, ByVal Addr, BlockSize, SizeRead, 0&)
End Sub

Public Sub API_CloseFile(ByVal FileNumber As Long)
Dim ret As Long
FlushFileBuffers FileNumber
ret = CloseHandle(FileNumber)
End Sub

Public Function API_WriteFile(ByVal FileNumber As Long, ByRef BlockSize As Long, ByRef Data() As Byte) As Boolean
Dim PosL As Long
Dim PosH As Long
Dim SizeWrit As Long
Dim ret As Long
ret = SetFilePointer(FileNumber, PosL, PosH, FILE_CURRENT)
ret = WriteFile(FileNumber, Data(0), BlockSize, SizeWrit, 0&)
API_WriteFile = (BlockSize = SizeWrit)
End Function

Private Sub Size2Long(ByVal FileSize As Currency, ByRef LongLow As Long, ByRef LongHigh As Long)
    Static ret As Currency
    ret = FileSize / 10000@
    GetMem4 VarPtr(ret), LongLow
    GetMem4 VarPtr(ret) + 4, LongHigh
End Sub

Private Sub Long2Size(ByVal LongLow As Long, ByVal LongHigh As Long, ByRef FileSize As Currency)
    FileSize = 0
    PutMem4 VarPtr(FileSize), LongLow
    PutMem4 VarPtr(FileSize) + 4, LongHigh
    FileSize = FileSize * 10000@
End Sub



Public Function ApiCanonicalize(ByVal Url As String, Optional dwFlags As Long = 0) As String
    Url = Left$(Url, INTERNET_MAX_URL_LENGTH)
   Dim dwSize As Long, res As String
   
   If Len(Url) > 0 Then
   
      ApiCanonicalize = space$(INTERNET_MAX_URL_LENGTH)
      dwSize = Len(ApiCanonicalize)
     
      If UrlCanonicalizeApi(StrPtr(Url), _
                    StrPtr(ApiCanonicalize), _
                    dwSize, _
                    dwFlags) = 0 Then
   
         ApiCanonicalize = Left$(ApiCanonicalize, dwSize)
         Else
         ApiCanonicalize = ""
         
      End If
   End If
 
End Function
Public Function GetUrlParts(ByVal sUrl As String, _
                             Optional ByVal dwPart As Long = 1, _
                             Optional ByVal dwFlags As Long = 0) As String

   Dim sPart As String
   Dim dwSize As Long
   
   If Len(sUrl) > 0 Then
   
      sPart = space$(INTERNET_MAX_URL_LENGTH)
      dwSize = Len(sPart)
     
      If UrlGetPart(StrPtr(sUrl), _
                    StrPtr(sPart), _
                    dwSize, _
                    dwPart, _
                    dwFlags) = 0 Then
   
         GetUrlParts = Left$(sPart, dwSize)
         
      End If
   End If

End Function
Public Function GetUrlQuery(ByVal Address As String) As String
    GetUrlQuery = GetUrlParts(UrlCanonicalize2(URLDecode(Address, True)), URL_PART_QUERY)

End Function
Public Function GetUrlPort(ByVal Address As String) As String
        GetUrlPort = GetUrlParts(UrlCanonicalize2(URLDecode(Address, True)), URL_PART_PORT)

End Function
Public Function URLDecode( _
    ByVal Url As String, _
    Optional ByVal PlusSpace As Boolean = True, Optional Flags As Long = 0) As String
    Url = Left$(Url, INTERNET_MAX_URL_LENGTH)
    Dim cchUnescaped As Long
    Dim hResult As Long
    
    If PlusSpace Then Url = Replace$(Url, "+", " ")
    cchUnescaped = Len(Url)
    URLDecode = String$(cchUnescaped, 0)
    hResult = UrlUnescape(StrPtr(Url), StrPtr(URLDecode), cchUnescaped, Flags)
    If hResult = E_POINTER Then
        URLDecode = String$(cchUnescaped, 0)
        hResult = UrlUnescape(StrPtr(Url), StrPtr(URLDecode), cchUnescaped, Flags)
    End If
    
    If hResult <> S_OK Then
        MyEr "can't decode this url", "δεν μπορώ να αποκωδικοποιήσω την διεύθυνση"
        Exit Function
    End If
    
    URLDecode = Left$(URLDecode, cchUnescaped)
End Function

Public Function URLEncode( _
    ByVal Url As String, _
    Optional ByVal SpacePlus As Boolean = True) As String
    Url = Left$(Url, INTERNET_MAX_URL_LENGTH)
    Dim cchEscaped As Long
    Dim hResult As Long
    If SpacePlus Then
      
        Url = Replace$(Url, " ", "+")
    End If
    cchEscaped = Len(Url) * 1.5
    URLEncode = String$(cchEscaped, 0)
    hResult = UrlEscape(StrPtr(Url), StrPtr(URLEncode), cchEscaped, URL_ESCAPE_PERCENT + &H40000)
    If hResult = E_POINTER Then
        URLEncode = String$(cchEscaped, 0)
        hResult = UrlEscape(StrPtr(Url), StrPtr(URLEncode), cchEscaped, URL_ESCAPE_PERCENT + &H40000)
    End If
    If hResult <> S_OK Then
      Exit Function
    End If
    
    URLEncode = Left$(URLEncode, cchEscaped)
 
End Function


Public Function GetParentAddress(ByVal Address As String, Optional includeroot As Boolean = False) As String
    Dim lngCharCount    As Long
    Dim lngBCount       As Long
    Dim strOutput       As String
     ' new from me
    Dim exclude As String
    
    If includeroot Then
    Address = URLDecode(Address)
    Else
    Address = RemoveRootName(URLDecode(Address, True), False)
    End If
    exclude = GetUrlParts(Address, URL_PART_QUERY)
    If Len(exclude) > 0 Then
    Address = Left$(Address, InStr(Address, exclude) - 2)
    
    End If
    GetParentAddress = ExtractPath(Address, False)
End Function
Private Function GetDomainName2(ByVal Address As String) As String
    Dim strOutput       As String
    Dim strTemp         As String
    Dim lngLoopCount    As Long
    Dim lngBCount       As Long
    Dim lngCharCount    As Long
    Dim I As Long
    ' new from me
   ' Address = URLDecode(Address, True)
    '
     strOutput$ = Replace(Address, "\", "/")
    lngCharCount = Len(strOutput)
    I = InStr(1, strOutput, "/")
    If I Then
        If I - InStr(1, strOutput, ":") > 1 Then
        Exit Function
        Else
        Do Until ((strTemp = "/") Or (lngLoopCount = lngCharCount))
            lngLoopCount = lngLoopCount + 1
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            lngBCount = lngBCount + 1
        Loop
        End If
    Else
    Exit Function
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
 
    If (InStr(1, strOutput, "/")) Then
        Do Until strTemp <> "/"
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    If Left$(strOutput, 1) = "[" Then
        lngBCount = InStr(strOutput, "]")
        If lngBCount > 0 Then strOutput = Left$(strOutput, lngBCount) Else strOutput = vbNullString
        GetDomainName2 = strOutput
        Exit Function
    ElseIf Not strOutput = vbNullString Then
    If InStr(1, strOutput, "/", vbTextCompare) = 0 Then
    I = InStr(1, strOutput, "@", vbTextCompare)
    If I > 0 Then GoTo 500
    End If
    End If
    
    On Error Resume Next
    strOutput = Left$(strOutput, InStr(1, strOutput, "/", vbTextCompare) - 1)
    If Err.Number > 0 Then strOutput = vbNullString
500    GetDomainName2 = strOutput
End Function
Public Function GetDomainName(ByVal Address As String, Optional userinfo As Boolean = False) As String
    Dim strOutput       As String
    Dim strTemp         As String
    Dim lngLoopCount    As Long
    Dim lngBCount       As Long
    Dim lngCharCount    As Long
    Dim I As Long
    ' new from me
    Address = URLDecode(Address, True)
    '
    strOutput$ = Replace(Address, "\", "/")
    lngCharCount = Len(strOutput)
    I = InStr(1, strOutput, "/")
    If I Then
        If I - InStr(1, strOutput, ":") > 1 Then
        Exit Function
        Else
        Do Until ((strTemp = "/") Or (lngLoopCount = lngCharCount))
            lngLoopCount = lngLoopCount + 1
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            lngBCount = lngBCount + 1
        Loop
        End If
    Else
    Exit Function
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
 
    If (InStr(1, strOutput, "/")) Then
        Do Until strTemp <> "/"
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    If Left$(strOutput, 1) = "[" Then
        lngBCount = InStr(strOutput, "]")
        If lngBCount > 0 Then strOutput = Left$(strOutput, lngBCount) Else strOutput = vbNullString
        GetDomainName = strOutput
        Exit Function
    ElseIf Not strOutput = vbNullString Then
    If InStr(1, strOutput, "/", vbTextCompare) = 0 Then
    If Not userinfo Then
    I = InStr(1, strOutput, "@", vbTextCompare)
    If I > 0 Then strOutput = Mid$(strOutput, I + 1)
    If Not strOutput = vbNullString Then If InStr(1, strOutput, ".", vbTextCompare) > 0 Then GoTo 500
    Else
    GoTo 500
    End If
    End If
    End If
    On Error Resume Next
    strOutput = Left$(strOutput, InStr(1, strOutput, "/", vbTextCompare) - 1)
    If Err.Number > 0 Then strOutput = vbNullString
500
    GetDomainName = strOutput
End Function
Public Function GetUrlPath(ByVal Address As String) As String
    Dim exclude As String, domain As String, scheme As String, W As Long
    
    
    Address = URLDecode(Address, False)
    scheme = GetUrlParts(Address)
    domain = GetDomainName(Address, True)
    exclude = GetUrlParts(Address, URL_PART_QUERY)
    If Len(domain) > 0 Then
        Address = UrlCanonicalize(Address)
    Else 'remove scheme only
        If Left$(Address, Len(scheme)) = scheme Then Address = Mid$(Address, Len(scheme) + 2)
    End If
    If domain <> vbNullString Then
    Address = Mid$(Address, Len(domain) + 1)
    ElseIf Not Address = vbNullString Then
    If InStr(Address, "//") = 0 Then
    If Left$(Address, Len(scheme)) = scheme Then Address = Mid$(Address, Len(scheme) + 2)
    End If
    End If
  
    If Not Address = vbNullString Then
    W = InStr(Address, "#")
    If W > 0 Then Address = Left$(Address, W - 1)
    End If
      If Len(exclude) > 0 Then
        Address = Left$(Address, Len(Address) - Len(exclude) - 1)
    End If
    GetUrlPath = Address
End Function
Private Function UrlCanonicalize2(ByVal pstrAddress As String) As String
    Dim strOutput       As String
    Dim strTemp         As String
    Dim lngLoopCount    As Long
    Dim lngBCount       As Long
    Dim lngCharCount    As Long

    strOutput$ = Replace(pstrAddress, "\", "/")
    lngCharCount = Len(strOutput)
 
    If (InStr(1, strOutput, "/")) Then
        Do Until ((strTemp = "/") Or (lngLoopCount = lngCharCount))
            lngLoopCount = lngLoopCount + 1
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
 
    If (InStr(1, strOutput, "/")) Then
        Do Until (strTemp <> "/")
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
    lngCharCount = Len(strOutput)
        If (InStr(1, strOutput, "/")) Then
        Do Until (strTemp <> "/")
            strTemp = Mid$(strOutput, lngCharCount - lngBCount, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    End If
 
   UrlCanonicalize2 = Left$(strOutput, Len(strOutput) - lngBCount)
    
   
End Function
Public Function UrlCanonicalize(ByVal pstrAddress As String) As String
    Dim strOutput       As String
    Dim strTemp         As String
    Dim lngLoopCount    As Long
    Dim lngBCount       As Long
    Dim lngCharCount    As Long
    ' new from me
    pstrAddress = URLDecode(pstrAddress, False)
    '
    strOutput$ = Replace(pstrAddress, "\", "/")
    lngCharCount = Len(strOutput)
 
    If (InStr(1, strOutput, "/")) Then
        Do Until ((strTemp = "/") Or (lngLoopCount = lngCharCount))
            lngLoopCount = lngLoopCount + 1
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
 
    If (InStr(1, strOutput, "/")) Then
        Do Until (strTemp <> "/")
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
    lngCharCount = Len(strOutput)
        If (InStr(1, strOutput, "/")) Then
        Do Until (strTemp <> "/")
            strTemp = Mid$(strOutput, lngCharCount - lngBCount, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Left$(strOutput, Len(strOutput) - lngBCount)
    ' strOutput = Replace(strOutput, "%20", " ") ' not used more
 
    UrlCanonicalize = strOutput
End Function
' we can use PurifyName()
'Public Function RemoveIllegals(ByVal pstrCheckString As String) As String
     
'End Function
Public Function GetHost(Url$) As String
        Dim W As Long
        GetHost = GetDomainName(Url$, True)
        If GetHost <> vbNullString Then
        If Left$(GetHost, 1) <> "[" Then
            W = InStr(GetHost, "@")
            If W > 0 Then GetHost = Mid$(GetHost, W + 1)
            If GetHost <> vbNullString Then
                 W = InStr(GetHost, ":")
                If W > 0 Then GetHost = Left$(GetHost, W - 1)
            End If
        Else
            W = InStr(GetHost, "]")
            GetHost = Mid$(GetHost, 2, W - 2)
        End If
        End If
End Function
Public Function RemoveRootName(ByVal pstrPath As String, _
                               ByVal pblnGetLowestLevelName As Boolean) _
                              As String
 
    Dim strOutput       As String
    Dim lngLoopCount    As Long
    Dim lngBCount       As Long
    Dim lngCharCount    As Long
    Dim strTemp         As String
     ' new from me
    pstrPath = URLDecode(pstrPath, True)
    '
    strOutput = Replace(pstrPath, "\", "/")
    lngCharCount = Len(strOutput)
 
    If (InStr(1, strOutput, "/")) Then
        Do Until ((strTemp = "/") Or (lngLoopCount = lngCharCount))
            lngLoopCount = lngLoopCount + 1
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
    If (InStr(1, strOutput, "/")) Then
        Do Until (strTemp <> "/")
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
    lngCharCount = Len(strOutput)
 
    If (InStr(1, strOutput, "/")) Then
        Do Until (strTemp <> "/")
            strTemp = Mid$(strOutput, lngCharCount - lngBCount, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = Left$(strOutput, Len(strOutput) - lngBCount)
    strOutput = Right$(strOutput, Len(strOutput) - InStr(1, strOutput, "/", vbTextCompare))
 
    If (pblnGetLowestLevelName) Then _
        strOutput = Right$(strOutput, Len(strOutput) - InStrRev(strOutput, "/"))
 
    'strOutput = Replace(strOutput, "%20", " ")
 
    RemoveRootName = strOutput
End Function
' ExpEnvirStr(string) as string  exist
Public Function URLEncodeEsc(cc As String, Optional space_as_plus As Boolean = False, Optional typedata As Long = 0) As String
   cc = StrConv(utf8encode(cc), vbUnicode)
    Dim slen As Long, M$: slen = Len(cc)
    Dim I As Long
    
    If slen > 0 Then
        ReDim res(slen) As String
        Dim ccode As Byte
        Dim cp1, cp2, cp3 As Integer
        Dim space As String
    
        If space_as_plus Then space = "+" Else space = "%20"
    If typedata = 0 Then
            For I = 1 To slen
            ccode = Asc(Mid$(cc, I, 1))
            Select Case ccode
                Case 97 To 122, 65 To 90, 48 To 57
                     res(I) = Chr(ccode)
                Case 32
                    res(I) = space
                Case 0 To 15
                    res(I) = "%0" & Hex(ccode)
                Case Else
                    res(I) = "%" & Hex(ccode)
            End Select
        Next I
    ElseIf typedata = 1 Then
        ' RFC3986
        For I = 1 To slen
            ccode = Asc(Mid$(cc, I, 1))
            Select Case ccode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                     res(I) = Chr(ccode)
                Case 32
                    res(I) = space
                Case 0 To 15
                    res(I) = "%0" & Hex(ccode)
                Case Else
                    res(I) = "%" & Hex(ccode)
            End Select
        Next I
    ElseIf typedata = 2 Then
        ' HMTL5
        For I = 1 To slen
            ccode = Asc(Mid$(cc, I, 1))
            Select Case ccode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 42
                     res(I) = Chr(ccode)
                Case 32
                    res(I) = space
                Case 0 To 15
                    res(I) = "%0" & Hex(ccode)
                Case Else
                    res(I) = "%" & Hex(ccode)
            End Select
        Next I
    End If
        URLEncodeEsc = Join(res, "")
    End If
End Function
Function DecodeEscape(C$, plus_as_space As Boolean) As String
If plus_as_space Then C$ = Replace(C$, "+", " ")
Dim A() As String, I As Long
A() = Split(C$, "%")
For I = 1 To UBound(A())
A(I) = Chr(val("&h" + Left$(A(I), 2))) + Mid$(A(I), 3)
Next I
DecodeEscape = utf8decode(StrConv(Join(A(), ""), vbFromUnicode))

End Function
Sub ClearState1()
If Not NOEDIT Then
NOEDIT = True
Else
If QRY Then QRY = False
End If
Sleep 300
Set Basestack1 = Nothing
abt = False
Set comhash = New sbHash
Set numid = New idHash
Set funid = New idHash
Set strid = New idHash
Set strfunid = New idHash
NERR = False
TaskMaster.Dispose
CloseAllConnections
CleanupLibHandles
' restore DB.Provider for User
JetPrefixUser = JetPrefixHelp
JetPostfixUser = JetPostfixHelp
' SET ARRAY BASE TO ZERO
ArrBase = 0
ReDim sbf(0), var(0)
Set globalstack = Nothing

End Sub
Public Sub NoTypeFound()
MyEr "No type found", "δεν βρήκα τύπο"
End Sub

Public Sub UseArrow()
MyEr "Use -> to get a pointer to group", "Χρησιμοποίησε το -> για να πάρεις δείκτη σε ομάδα"
End Sub
Public Sub NoPointerinVar(W$)
MyEr "No pointer in " + W$, "Δεν υπάρχει δείκτης στην " + W$
End Sub
Public Sub OnlyForGroupPointers()
MyEr "Only for Group Pointers", "Μόνο για δείκτες ομάδων"
End Sub
Public Sub CanyAssignPointer2Group()
MyEr "Can't assign pointer to named group", "Δεν μπορώ να βάλω δείκτη σε επώνυμη ομάδα"
End Sub
Public Sub MissingPointer()
MyEr "No Pointer Found", "Δεν βρήκα δείκτη"
End Sub
Public Sub ManyDots()
    MyEr "too many dots", "πολλές τελείες"
End Sub
Public Sub ExpectedPointer()
MyEr "Expected pointer", "Περίμενα δείκτη"
End Sub
Public Sub WrongFatArrow()
MyEr "Wrong use of => operator", "Κακή χρήση του τελεστή =>"
End Sub
Public Sub NeedString()
MyEr "Need a string", "Χρειάζομαι ένα αλφαριθμητικό"
End Sub
Public Sub FoundNoStringItem()
MyEr "Not a string array item", "Δεν έχει αλφαριθμητικό ο πίνακας"
End Sub
Public Sub MissingIndexMore()
MyEr "Missing one [index] more", "Δεν βρήκα έναν δείκτη ακόμα [δείκτης]"
End Sub
Public Sub ExpRefArray(I As Long)
MyEr "expected RefArray at index " & I, "περίμενα RefArray στον δείκτη " & I
End Sub
Private Sub PartExecVar(ss$, v, p, sp)
                Select Case ss$
                Case "DIV", "ΔΙΑ"
                    If VarType(var(v)) = 20 Then
                    If Not VarType(p) = 20 Then p = cInt64(p)
                    var(v) = var(v) \ p
                    Else
                    var(v) = Fix(var(v) / p)
                    End If
                Case "DIV#", "ΔΙΑ#"
                If VarType(var(v)) = 20 Then
                If Not VarType(p) = 20 Then p = cInt64(p)
                If p < 0 Then
                        var(v) = ((var(v) - Abs(var(v) - Abs(p) * (var(v) \ Abs(p)))) \ p)
                    Else
                        var(v) = var(v) \ p
                    End If
                Else
                    If p < 0 Then
                        var(v) = Int((var(v) - Abs(var(v) - Abs(p) * Int(var(v) / Abs(p)))) / p)
                    Else
                        var(v) = Int(var(v) / p)
                    End If
                End If
                Case "MOD", "ΥΠΟΛ", "ΥΠΟΛΟΙΠΟ"
                    If VarType(var(v)) = 20 Then
                    If Not VarType(p) = 20 Then p = cInt64(p)
                        sp = var(v) - (var(v) \ p) * p
                    Else
                        sp = var(v) - Fix(var(v) / p) * p
                    End If
                    If Abs(sp) >= Abs(p) Then sp = sp - sp
                    var(v) = sp
                 Case "MOD#", "ΥΠΟΛ#", "ΥΠΟΛΟΙΠΟ#"
                    If VarType(var(v)) = 20 Then
                    If Not VarType(p) = 20 Then p = cInt64(p)
                        sp = Abs(var(v) - Abs(p) * (var(v) \ Abs(p)))
                    Else
                        sp = Abs(var(v) - Abs(p) * Int(var(v) / Abs(p)))
                    End If
                    If Abs(sp) >= Abs(p) Then sp = sp - sp
                    var(v) = sp
                Case Else
                    WrongOperator
                End Select
End Sub
Public Function ExecuteVar(Exec1 As Long, ByVal jumpto As Long, bstack As basetask, W$, b$, v As Long, Lang As Long, VarStat As Boolean, NewStat As Boolean, nchr As Integer, ss$, sss As Long, temphere$, noVarStat As Boolean) As Long
Dim I As Long, p As Variant, myobject As Object, ok As Boolean, sw$, sp As Variant, useType As Boolean
Dim pppp As mArray, lasttype As Integer, pppp1 As mArray, isglobal As Boolean, usehandler As mHandler, usehandler1 As mHandler, idx As mIndexes, myProp As PropReference
Dim newid As Boolean, Ar As refArray, ww As Integer
Dim ppppAny As iBoxArray
 Const mProp = "PropReference"
 Const mHdlr = "mHandler"
 Const mGroup = "Group"
 Const myArray = "mArray"
 Const a123 = "={-"
Const b123 = vbCr + "'\/"
Const b1234 = vbCr + "'\/:"
Const b12345 = vbCr + "'\/:}"
Const RemChar = "'\/"
On jumpto GoTo Case1, Case2, Case3, Case4, case5, Case6, Case7, Case8, Case8new
Exit Function
Case1:
    Select Case CheckThis(bstack, W$, b$, v, Lang)
    Case 0
    useType = True
    Case 1
    useType = True
    GoTo assignvalue
    Case 2
    useType = True
    GoTo somethingelse
    Case 3
    useType = True
    GoTo assignpointer
    Case -1
        GoTo err000
    End Select
    I = MyTrimL(b$)

    If VarStat Then
     ' MAKE A GLOBAL SO ONLY = ALLOWED
     
        If FastOperator2(b$, "=", I) Then
            GoTo jumpiflocal
        Else
            p = 0#
            If IsLabelSymbolNew(b$, "ΩΣ", "AS", Lang) Then
                On GetType(bstack, b$, p, v, W$, Lang, VarStat, temphere$, noVarStat) GoTo NewCheck, NewCheck2
                Exit Function
            ElseIf FastSymbol(b$, "->", , 2) Then
                v = globalvar(W$, p, , VarStat, temphere$)
                GoTo assignpointer
             Else
             If GetSub(W$ + "()", v) Then
checkplease1:
                If Not sbf(v).IamAClass Then
                 WrongType
                 ExecuteVar = 0
                 Exit Function
                End If
                If Not AddGroupFromClass(bstack, b$, W$, VarStat, False, temphere$) Then
                ExecuteVar = 0
                 Exit Function
                End If
            ElseIf GetSub(W$ + "$()", v) Then
            GoTo checkplease1
            Else
                v = globalvar(W$, p, , VarStat, temphere$)
                End If
                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
            End If
        End If
    ElseIf NewStat Or noVarStat Then
        ' MAKE A NEW ONE SO ONLY = ALLOWED
        If FastOperator2(b$, "=", I) Then
            GoTo jumpiflocal
        Else
            p = 0#
            If IsLabelSymbolNew(b$, "ΩΣ", "AS", Lang) Then
                On GetType(bstack, b$, p, v, W$, Lang, VarStat, temphere$, noVarStat) GoTo NewCheck, NewCheck2
                Exit Function
            Else
checkhereClass:
            If GetSub(W$ + "()", v) Then
checkplease2:
                If Not sbf(v).IamAClass Then
                    GoTo noisnotAclass
                 End If
cont12987:
                If Not AddGroupFromClass(bstack, b$, W$, False, NewStat, temphere$) Then
                Exec1 = 0: ExecuteVar = 11
                 Exit Function
                End If
            ElseIf GetSub(W$ + "$()", v) Then
            GoTo checkplease2
            Else
noisnotAclass:
                If comhash.Find2(W$, (0), v) Then
                If v = 44 Then
                    GoTo cont12987
                End If
                End If
                v = globalvar(W$, p, , VarStat, temphere$)
            End If
                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
            
            End If
        End If
    ElseIf nchr > 31 Then
        If Left$(b$, 1) = "_" Then
            If nchr <> 61 Then
                SyntaxError
                Exit Function
            End If
            If GetVar(bstack, W$, v, True, , , , useType) Then
                W$ = varhash.lastkey
                Mid$(b$, 1, 1) = " "
                GoTo assignvalue
            ElseIf GetlocalVar(W$, v) Then
                useType = varhash.vType(varhash.Index)
                If TypeOf var(v) Is Group Then
                    If Not var(v).IamRef Then
                        W$ = varhash.lastkey
                    End If
                Else
                    W$ = varhash.lastkey
                End If
                Mid$(b$, 1, 1) = " "
                GoTo assignvalue
            Else
                Mid$(b$, 1, 1) = "="
                If AscW(Left$(W$, 1)) = &H1FFF Then
                    If here$ = vbNullString Then
                        If varhash.Find(W$, v) Then
                            GoTo fromthis
                        End If
                    Else
                        If varhash.Find(here$ + "." + W$, v) Then
                            GoTo fromthis
                        End If
                    End If
                Else
                    UnknownVariable W$
                End If
                GoTo err000
            End If
        ElseIf MaybeIsSymbol(b$, "/*-+=~^|<") Then
            If Mid$(b$, I, 2) = "//" Then
                If GetSub(W$, v) Then
                    ExecuteVar = 6 ' GoTo autogosub
                Else
                    Exec1 = 0
                End If
                Exit Function
            End If
            If Mid$(b$, I, 2) = "<=" Then
        ' LOOK GLOBAL
                If GetVar(bstack, W$, v, True, , , , useType, isglobal) Then
                    W$ = varhash.lastkey
                    Mid$(b$, I, 2) = "  "
                    GoTo assignvalue
                ElseIf GetlocalVar(W$, v) Then
                    useType = varhash.vType(varhash.Index)
                    If TypeOf var(v) Is Group Then
                        If Not var(v).IamRef Then
                            W$ = varhash.lastkey
                        End If
                    Else
                        W$ = varhash.lastkey
                    End If
                    Mid$(b$, I, 2) = "  "
                    GoTo assignvalue
                Else
                    Mid$(b$, I, 1) = " "
                    I = I + 1
                    If AscW(Left$(W$, 1)) = &H1FFF Then
                        If here$ = vbNullString Then
                            If varhash.Find(W$, v) Then
                                GoTo fromthis
                            End If
                        Else
                            If varhash.Find(here$ + "." + W$, v) Then
                                GoTo fromthis
                            End If
                        End If
                    Else
                        UnknownVariable W$
                    End If
                    GoTo err000
                End If
                ' do something here
            ElseIf varhash.Find2(here$ + "." + myUcase(W$), v, useType) Then
                    '   If TypeOf var(v) Is Group Then w$ = varhash.lastkey
                    ' CHECK VAR
fromthis:           If FastOperator(b$, "=", I) Then
assignvalue:
                        If MyIsNumeric(var(v)) Then
assignvalue2:
                            If IsExp(bstack, b$, p) Then
assignvalue3:
                                If bstack.lastobj Is Nothing Then
                                    If useType And Not newid Then
                                        If AssignTypeNumeric(p, VarType(var(v))) Then
                                            var(v) = p
                                        Else
                                            GoTo err000
                                        End If
                                    Else
                                        var(v) = p
                                    End If
                                Else
checkobject:
                                    If MemInt(VarPtr(bstack.lastobj)) = 13 Then
                                        Set var(v) = bstack.lastobj
                                        Set bstack.lastobj = Nothing
                                    Else
                                        If Not procObject(bstack, W$, p, v, useType, VarStat, isglobal, NewStat) Then GoTo err000
                                    End If

                                End If
                            ElseIf LastErNum1 < 0 Then
                                Exec1 = 0: ExecuteVar = 11
                                Exit Function
                            ElseIf IsStrExp(bstack, b$, ss$, (Len(bstack.tmpstr) = 0) And newid) Then
                                If bstack.lastobj Is Nothing Then
                                    If newid Or Not useType Or VarStat Or NewStat Or noVarStat Then
                                        var(v) = ss$
                                    ElseIf useType And MemInt(VarPtr(var(v))) = vbString Then
                                        var(v) = ss$
                                    ElseIf useType And MemInt(VarPtr(var(v))) = vbUserDefinedType Then
                                        MissType
                                        GoTo err000
                                    ElseIf ss$ = vbNullString Then
                                        var(v) = 0#
                                    Else
                                        If IsNumberCheck(ss$, p) Then
                                            If useType Then
                                                If AssignTypeNumeric(p, MemInt(VarPtr(var(v)))) Then
                                                    var(v) = p
                                                Else
                                                    GoTo err000
                                                End If
                                            Else
                                                If MemInt(VarPtr(var(v))) = vbLong Then
                                                    On Error Resume Next
                                                    var(v) = CLng(p)
                                                    If Err.Number > 0 Then OverflowValue: GoTo err000
                                                    On Error GoTo 0
                                                ElseIf MemInt(VarPtr(var(v))) = vbInteger Then
                                                    On Error Resume Next
                                                    var(v) = CInt(p)
                                                    If Err.Number > 0 Then OverflowValue vbInteger: GoTo err000
                                                    On Error GoTo 0
                                                Else
                                                    var(v) = p
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    GoTo checkobject
                                End If
                            Else
                                If var(v) = vbEmpty Then var(v) = 0#
                                NoValueForVar W$
                                GoTo err000
                            End If
                            If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                        ElseIf Not MyIsObject(var(v)) Then
                            ww = MemInt(VarPtr(var(v)))
                            If useType And Not newid Then
                                If ww = vbUserDefinedType Then
                                    If IsExp(bstack, b$, p, flatobject:=True, nostring:=True) Then
                                        If MemInt(VarPtr(p)) = vbUserDefinedType Then
                                            If Typename(p) = Typename(var(v)) Then
                                                SwapVariant var(v), p
                                            Else
                                                GoTo notypevarV
                                            End If
                                        Else
                                            GoTo notypevarV
                                        End If
                                    Else
                                        GoTo notypevarV
                                    End If
                                ElseIf ww = vbString Then
                                    If IsExp(bstack, b$, p, , True) Then
                                        Assign2 ss$, p
                                        GoTo assignvalue3
                                    End If
                                    GoTo assignvaluestr1
                                Else
                                    GoTo assignvalue2
                                End If
                            Else
                                GoTo assignvalue2
                            End If
                    Else
                        If Left$(b$, 2) <> " >" Then
                            If useType = False Then
                                var(v) = Empty
                                GoTo assignvalue2
                            End If
                        Else
                            useType = True
                        End If
assigngroup:
                        If var(v) Is Nothing Then
                            If IsExp(bstack, b$, p) Then
                                If Not bstack.lastobj Is Nothing Then
                                    Set p = bstack.lastobj
                                    If TypeOf p Is Group Then
                                        If Not p.IamApointer Then MakeGroupPointer bstack, p
                                        Set var(v) = bstack.lastobj
                                        Set bstack.lastobj = Nothing
                                        Set bstack.lastpointer = Nothing
                                        If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                                    End If
                                End If
                            End If
                            AssigntoNothing  ' Use Declare
                            GoTo err000
                        ElseIf TypeOf var(v) Is Group Then
                            If IsExp(bstack, b$, p) Then
hasstr1:
                                If var(v).HasSet Then
                                    Set myobject = bstack.soros
                                    Set bstack.Sorosref = New mStiva
                                    If bstack.lastobj Is Nothing Then
                                        bstack.soros.PushVal p
                                    ElseIf TypeOf bstack.lastobj Is mHandler Then
                                        Set usehandler = bstack.lastobj
                                    If usehandler.T1 = 4 Then
                                        bstack.soros.PushVal p
                                    Else
                                        bstack.soros.DataObj bstack.lastobj
                                    End If
                                Else
                                    If TypeOf bstack.lastobj Is VarItem Then
                                        bstack.soros.DataOptional
                                    Else
                                        bstack.soros.DataObj bstack.lastobj
                                    End If
                                    Set bstack.lastobj = Nothing
                                End If
                                NeoCall2 bstack, W$ + "." + ChrW(&H1FFF) + ":=()", ok
                                Set bstack.Sorosref = myobject
                                Set myobject = Nothing
                            ElseIf bstack.lastobj Is Nothing Then
                                NeedAGroupInRightExpression
                                GoTo err000
                            ElseIf TypeOf bstack.lastobj Is Group Then
                                Set myobject = bstack.lastobj
                                Set bstack.lastobj = Nothing
                                ss$ = bstack.GroupName
                                If var(v).HasValue Or var(v).HasSet Then
                                    PropCantChange
                                    GoTo err000
                                Else
                                    If Len(var(v).GroupName) > Len(W$) Then
                                        sw$ = here$
                                        here$ = vbNullString
                                        UnFloatGroupReWriteVars bstack, var(v).Patch, v, myobject
                                        here = sw$
                                        myobject.ToDelete = True
                                    Else
                                        bstack.GroupName = Left$(W$, Len(W$) - Len(var(v).GroupName) + 1)
                                        If Len(var(v).GroupName) > 0 Then
                                            W$ = Left$(var(v).GroupName, Len(var(v).GroupName) - 1)
                                            sw$ = here$
                                            here$ = vbNullString
                                            UnFloatGroupReWriteVars bstack, W$, v, myobject
                                            here = sw$
                                            myobject.ToDelete = True
                                        ElseIf var(v).IamApointer And myobject.IamApointer Then
                                            Set var(v) = myobject
                                        Else
                                            Set myobject = Nothing
                                            bstack.GroupName = ss$
                                            If var(v).IamApointer Then
                                                UseArrow
                                            Else
                                                GroupWrongUse
                                            End If
                                            GoTo err000
                                        End If
                                    End If
                                End If
                                Set myobject = Nothing
                                bstack.GroupName = ss$
                                Set bstack.lastpointer = Nothing
                            Else
                                GoTo WrongObj
                            End If
                            If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                        ElseIf IsStrExp(bstack, b$, ss$, False) Then
                            p = vbNullString
                            SwapString2Variant ss$, p
                            GoTo hasstr1
                        Else
                        
noexpression:
                            If Left$(b$, 1) = ">" Then
noexpression1:
                                If var(v).IamApointer Then
                                    If var(v).link.IamFloatGroup Then
                                        ExecuteVar = 10
                                        Mid$(b$, 1, 1) = ChrW(3)
                                    Else
                                        ExecuteVar = 9
                                        Mid$(b$, 1, 1) = Chr$(0)
                                End If
                                Set bstack.lastpointer = var(v)
                                Exit Function
                            Else
                                NoPointerinVar (W$)
                            End If
                        End If
                        Set myobject = Nothing
                        Set bstack.lastobj = Nothing
                        MissNumExpr
                        GoTo err000
                    End If
                    Exit Function
                ElseIf TypeOf var(v) Is PropReference Then
                    If IsExp(bstack, b$, p) Then
                        If FastSymbol(b$, "@") Then
                            If IsExp(bstack, b$, sp, flatobject:=True) Then
                                If MemInt(VarPtr(sp)) = vbString Then
                                    SwapString2Variant ss$, sp
                                    var(v).Index = ss$: ss$ = vbNullString
                                Else
                                    var(v).Index = sp: sp = 0
                                End If
                            ElseIf IsStrExp(bstack, b$, ss$, False) Then
                                var(v).Index = ss$: ss$ = vbNullString
                            End If
                            var(v).UseIndex = True
                        End If
                        var(v).Value = p
                    Else
                        GoTo noexpression
                    End If
                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                    ElseIf TypeOf var(v) Is lambda Then
                        ' exist and take something else
                        If IsExp(bstack, b$, p) Then
                            If bstack.lastobj Is Nothing Then
                                Expected "lambda", "λάμδα"
                            ElseIf TypeOf bstack.lastobj Is lambda Then
                                Set var(v) = bstack.lastobj
                                Set bstack.lastobj = Nothing
                                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                            Else
                                Expected "lambda", "λάμδα"
                            End If
                            GoTo err000
    
                        Else
                            MissNumExpr
                            GoTo err000
                        End If
                    ElseIf TypeOf var(v) Is mHandler Then  ' CHECK IF IT IS A HANDLER
                        Set usehandler = var(v)
                        If IsExp(bstack, b$, p) Then
                            If usehandler.ReadOnly Then
                                ReadOnly
                                GoTo err000
                            End If
jumpbackhere:
                            Set usehandler = var(v)
                            If bstack.lastobj Is Nothing Then
                                If usehandler.T1 = 4 Then
checkfromstring:
                                    Set myobject = usehandler.objref.SearchValue(p, ok)
                                    If ok Then
                                        Set var(v) = myobject
                                    Else
                                        ExpectedEnumType
                                        GoTo err000
                                    End If
                                Else
                                    NoObjectFound
                                    GoTo err000
                                End If
                            ElseIf Typename(bstack.lastobj) = mHdlr Then
                            
                                Set usehandler1 = New mHandler
                                Set usehandler = bstack.lastobj
                                usehandler.CopyTo usehandler1
                                If usehandler.indirect > 0 Then
                                    Set myobject = usehandler1
                                    CheckDeepAny myobject
                                    usehandler.indirect = -1
                                    Set usehandler.objref = myobject
                                    Set var(v) = usehandler
                                    Set usehandler1 = New mHandler
                                    bstack.lastobj.CopyTo usehandler1
                                 ElseIf usehandler1.T1 = 4 Then
                                    Set usehandler = var(v)
                                    If Not usehandler.objref Is usehandler1.objref Then
                                        If usehandler.objref.EnumName = usehandler1.objref.EnumName Then
                                            If usehandler.objref.ExistFromOther2(usehandler1) Then
                                                Set usehandler1.objref = usehandler.objref
                                            ElseIf usehandler.objref.ExistFromOther(usehandler1.index_cursor) Then
                                                Set usehandler1.objref = usehandler.objref
                                                usehandler1.index_start = usehandler.objref.Index
                                            Else
                                                GoTo contwrong1
                                            End If
                                        ElseIf usehandler.objref.ExistFromOther2(usehandler1) Then
                                                Set usehandler1.objref = usehandler.objref
                                        Else
                                        'usehandler1.objref.Index = usehandler1.index_start
                                   
                                        
contwrong1:
                                            WrongType
                                            Set bstack.lastobj = Nothing
                                            GoTo err000
                                        End If
                                    End If
                                End If
                                Set var(v) = usehandler1
                                
                            ElseIf Typename(bstack.lastobj) = myArray Then
                                Set usehandler1 = New mHandler
                                usehandler1.T1 = 3
                                Set usehandler1.objref = bstack.lastobj
                                Set var(v) = usehandler1
                                
                            Else
                                
                                Set usehandler1 = var(v)
                                usehandler1.T1 = 0
                                Set usehandler1.objref = bstack.lastobj
                                
                            End If
                            Set usehandler1 = Nothing
                            Set myobject = Nothing
                        Else
                            If usehandler.T1 = 4 Then
                                If IsStrExp(bstack, b$, ss$) Then
                                    p = vbNullString
                                    SwapString2Variant ss$, p
                                    GoTo checkfromstring
                                End If
                            End If
                            MissNumExpr
                            GoTo err000
                            
                        End If
                        Set bstack.lastobj = Nothing
                        Set myobject = Nothing
                      ElseIf TypeOf var(v) Is Constant Then
                        If myVarType(var(v).Value, vbEmpty) Then
                            If IsExp(bstack, b$, p) Then
                                If bstack.lastobj Is Nothing Then
                                    var(v).DefineOnce p
                                Else
                                    CantAssignValue
                                    MissNumExpr
                                    GoTo err000
                                End If
                            Else
                                MissNumExpr
                                GoTo err000
                            End If
                        Else
                            If InStr(ss$, ".") = 0 Or var(v).flag Then
                                CantAssignValue
                            Else
                                NoOperatorForThatObject "="
                            End If
                            GoTo err000
                        End If
                      ElseIf TypeOf var(v) Is mEvent Then
                      If IsExp(bstack, b$, p) Then
                      If Typename$(bstack.lastobj) = "mEvent" Then
                            Set var(v) = bstack.lastobj
                            CopyEvent var(v), bstack
                            Set var(v) = bstack.lastobj
                            Set bstack.lastobj = Nothing
                            End If
                        Else
misnum:                         MissNumExpr
                            GoTo err000
                        End If
          
                    ElseIf MyIsObject(var(v)) Then
                    If IsExp(bstack, b$, p) Then
                    Set myobject = bstack.lastobj
                    If Not myobject Is Nothing Then
                        Set p = myobject
                        Set bstack.lastobj = Nothing
                        If VarTypeName(p) = VarTypeName(var(v)) Then
                        If TypeOf p Is BigInteger Then
                            If Not var(v) Is p Then
                                Set var(v) = CopyBigInteger(p, var(v))
                            End If
                        Else
                            Set var(v) = p
                        End If
                        Set myobject = Nothing
                    ElseIf TypeOf var(v) Is refArray Then
                    Set Ar = var(v)
                       Set myobject = p
                       If Not CheckIsmArray(myobject) Then
                       GoTo WrongObj
                       End If
                       Set p = myobject
                       Set myobject = Nothing
                        If TypeOf p Is mArray Then
                            Set pppp = p
                            If pppp.count = 0 Then
                                Ar.ResetToType Ar.vtType(0), 0
                                Ar.UnFlat
                            Else
                            p = CVar(pppp.ExportArrayCopy)
                            
                            If myVarType(p, vbEmpty) Then
                            Set Ar = pppp.ExportArrayCopy
                            
                            Else
                            
                            Ar.writevalue = p
                            Ar.flat = True
                            Ar.RedimForFlat pppp.count - 1
                        
                            
                            End If
                            Ar.UnFlat
                            Set var(v) = Ar
                            End If
                        
                        Else
                        GoTo WrongObj
                        End If
                    Else
                        GoTo WrongObj
                    End If
                Else
                If TypeOf var(v) Is BigInteger Then
                    On Error GoTo C12313
                    If MyIsNumeric(p) Then
                        Set var(v) = Module13.CreateBigInteger(CStr(Int(p)))
                    Else
                        Set var(v) = Module13.CreateBigInteger(CStr(p))
                    End If
                Else
C12313:
                    GoTo WrongObj
                    End If
                End If
                Else
                    GoTo misnum
                End If
    
                Else
                    GoTo somethingelse
                End If
            End If
        Else
            ' or do something else
            
somethingelse:
            I = MyTrimL(b$)
            If InStr("/*-+=~^&|<>", Mid$(b$, I, 1)) > 0 Then
                If InStr("/*-+=~^&|<>!", Mid$(b$, I + 1, 1)) > 0 Then
                    ss$ = Mid$(b$, I, 2)
                    If ss$ = "=&" Then
                    ss$ = "= "
                    Mid$(b$, I, 1) = " "
                    Else
                    Mid$(b$, I, 2) = "  "
                    End If
                ElseIf AscW(b$) = 124 Then
                   
                    Mid$(b$, I, 1) = " "
                    ww = FastPureLabel(b$, ss$, , , , , False)
                    If ww = 1 Or ww = 5 Then
                        ss$ = "@@"
                    Else
                        WrongOperator
                    End If
                Else
                    ss$ = Mid$(b$, I, 1)
                    Mid$(b$, I, 1) = " "
                    
                End If
            Else
                ExecuteVar = 6: Exit Function
            End If
            
            If MyIsNumeric2(var(v), lasttype) Then
            On Error GoTo LONGERR
            If lasttype = vbInteger Then
                Select Case ss$
                Case "="
                    v = globalvar(W$, CInt(p), , VarStat, temphere$)
                    GoTo assignvalue2
                Case "+="
                    If IsExp(bstack, b$, p) Then
                        var(v) = CInt(Int(p) + var(v))
                    Else
                        GoTo noexpression
                    End If
                Case "-="
                    If IsExp(bstack, b$, p) Then
                        var(v) = CInt(-Int(p) + var(v))
                    Else
                        GoTo noexpression
                    End If
                Case "*="
                    If IsExp(bstack, b$, p) Then
                        var(v) = CInt(Int(p) * var(v))
                    Else
                        GoTo noexpression
                    End If
                Case "/="
                    If IsExp(bstack, b$, p) Then
                        If Int(p) = 0 Then
                            DevZero
                            GoTo err000
                        End If
                        var(v) = CInt(var(v) \ Int(p))
                    Else
                        GoTo noexpression
                    End If
                Case "-!"
                    var(v) = CInt(-var(v))
                Case "++"
                    var(v) = CInt(1 + var(v))
                Case "--"
                    var(v) = CInt(var(v) - 1)
                Case "~"
                    var(v) = CInt(Not CBool(var(v)))
                Case "@@"
                    FastPureLabel b$, ss$, , True
                    If Mid$(b$, 1, 1) = "#" Then ss$ = ss$ + "#": Mid$(b$, 1, 1) = " "
                    If IsExp(bstack, b$, p) Then
                        If Int(p) = 0 Then
                            DevZero
                            GoTo err000
                        End If
                        If Not readvarv(var(v), ss$, p) Then
                            WrongOperator
                            GoTo err000
                        End If
                        var(v) = CInt(var(v))
                    Else
                        GoTo noexpression
                    End If
                   
                Case Else
                ExecuteVar = 6: Exit Function
            End Select
            GoTo checksyntax
            ElseIf VarType(var(v)) = vbLong Then
                Select Case ss$
                    Case "="
                        v = globalvar(W$, CLng(p), , VarStat, temphere$)
                        GoTo assignvalue2
                    Case "+="
                        If IsExp(bstack, b$, p) Then
                            var(v) = CLng(Int(p) + var(v))
                        Else
                            GoTo noexpression
                        End If
                    Case "-="
                        If IsExp(bstack, b$, p) Then
                            var(v) = CLng(-Int(p) + var(v))
                        Else
                            GoTo noexpression
                        End If
                    Case "*="
                        If IsExp(bstack, b$, p) Then
                            var(v) = CLng(Int(p) * var(v))
                        Else
                            GoTo noexpression
                        End If
                    Case "/="
                        If IsExp(bstack, b$, p) Then
                            If Int(p) = 0 Then
                                DevZero
                                GoTo err000
                            End If
                            var(v) = CLng(var(v) \ Int(p))
                        Else
                            GoTo noexpression
                        End If
                    Case "-!"
                        var(v) = CLng(-var(v))
                    Case "++"
                        var(v) = CLng(1 + var(v))
                    Case "--"
                        var(v) = CLng(var(v) - 1)
                    Case "~"
                        var(v) = CLng(Not CBool(var(v)))
                    Case "@@"
                    FastPureLabel b$, ss$, , True
                    If Mid$(b$, 1, 1) = "#" Then ss$ = ss$ + "#": Mid$(b$, 1, 1) = " "
                    If IsExp(bstack, b$, p) Then
                        If Int(p) = 0 Then
                            DevZero
                            GoTo err000
                        End If
                        If Not readvarvLong(v, ss$, p) Then
                            WrongOperator
                        End If
                    Else
                        GoTo noexpression
                    End If
                Case Else
                ExecuteVar = 6: Exit Function
            End Select
checksyntax:
                        If NocharsInLine(b$) Then ExecuteVar = 8: Exit Function
                        If MaybeIsSymbol(b$, b12345) Then
                         If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                        End If
                        SyntaxError
                        GoTo err000
                On Error GoTo 0
                Else
                
                On Error Resume Next
                
              '  lasttype = VarType(var(v))
                Select Case ss$
                    Case "="
                        v = globalvar(W$, p, , VarStat, temphere$)
                        GoTo assignvalue2
                    Case "+="
                        If IsExp(bstack, b$, p) Then
                           
                            var(v) = p + var(v)
                            
                            If Err.Number = 6 Then
                            Err.Clear
                            var(v) = CDbl(p) + CDbl(var(v))
                            End If
                            If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "-="
                        If IsExp(bstack, b$, p) Then
                            var(v) = -p + var(v)
                            If Err.Number = 6 Then
                            Err.Clear
                            var(v) = CDbl(-p) + CDbl(var(v))
                            End If
                            If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "*="
                        If IsExp(bstack, b$, p) Then
                         '   If lasttype = 20 Then If Not VarType(p) = 20 Then p = cInt64(p)
                            sp = var(v)
                            sp = p * var(v)
                            If Err.Number = 6 Then
                            Err.Clear
                            var(v) = CDbl(p) * CDbl(var(v))
                            Else
                            var(v) = sp
                            End If
                            If RoundDouble Then If lasttype = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "/="
                        If IsExp(bstack, b$, p) Then
                            If p = 0# Then
                                DevZero
                                GoTo err000
                            End If
                            If VarType(var(v)) = 20 Then
                                If Not VarType(p) = 20 Then p = cInt64(p)
                                var(v) = var(v) \ p
                            Else
                                var(v) = var(v) / p
                            End If
                            
                            If Err.Number = 6 Then
                            Err.Clear
                            var(v) = CDbl(var(v)) / CDbl(p)
                            End If
                            If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "-!"
                        var(v) = -var(v)
                    Case "++"
                        
                        'sp = sp + 1
                     
                        var(v) = var(v) + 1
                     
                    Case "--"
                        var(v) = var(v) - 1
                    Case "~"
                    Select Case VarType(var(v))
                    Case vbBoolean
                        var(v) = Not CBool(var(v))
                    Case vbCurrency
                        var(v) = CCur(Not CBool(var(v)))
                    Case vbDecimal
                        var(v) = CDec(Not CBool(var(v)))
                    Case Else
                        var(v) = CDbl(Not CBool(var(v)))
                        End Select
                    Case "->"
                        GoTo assignpointer
                    Case "@@"
                        FastPureLabel b$, ss$, , True
                        If Mid$(b$, 1, 1) = "#" Then ss$ = ss$ + "#": Mid$(b$, 1, 1) = " "
                        If IsExp(bstack, b$, p) Then
                            If Int(p) = 0 Then
                                DevZero
                                GoTo err000
                            End If
                            PartExecVar ss$, v, p, sp
                        Else
                            GoTo noexpression
                        End If
                       
                    Case Else
                    
                    If Err.Number = 6 Then
                        Overflow
                        Err.Clear
                    ElseIf Len(ss$) > 0 Then
                    If GetSub(W$, v) Then
              
                    Mid$(b$, 1, Len(ss$)) = ss$
                    ExecuteVar = 6 ' GoTo autogosub
                    Exit Function
                    Else
                        WrongOperator
                        Exec1 = 0
                        End If
                    Else
                        SyntaxError
                    End If
                    
                    GoTo err000

                End Select
                If Err.Number = 6 Then
                    Err.Clear
                    GoTo LONGERR
                ElseIf Not VarType(var(v)) = lasttype Then
                    If useType Then
                        'If lasttype = vbSingle Or lasttype = vbDouble Or lasttype = vbCurrency Or lasttype = 20 Then
                            If Not AssignTypeNumeric2(var(v), CLng(lasttype)) Then GoTo LONGERR
                        
                        'Else
                         '   GoTo LONGERR
                        'End If
                    End If
                End If
                
                On Error GoTo 0
                GoTo checksyntax
                End If
                ElseIf Not MyIsObject(var(v)) Then
                
                If MemInt(VarPtr(var(v))) = vbString Then
                GoTo stroper001
                Else
                If MemInt(VarPtr(var(v))) = vbUserDefinedType Then
                If ss$ = "@@" Then
                ww = FastPureLabel(b$, ss$)
                If ww > 0 Then
                If ww = 1 Then
Z1123698:
                    If FastSymbol(b$, "=") Then
                        If IsExp(bstack, b$, p) Then
                            Err.Clear
                            On Error Resume Next
                            If ww = 5 Then
                                PlaceValue2UDTArray var(v), ss$, p, I
                            Else
                                PlaceValue2UDT var(v), ss$, p
                            End If
                            If Err Then
                                MyEr Err.Description, Err.Description
                                GoTo err000
                            End If
                        ElseIf IsStrExp(bstack, b$, W$, False) Then
                            p = ""
                            SwapString2Variant W$, p
                            Err.Clear
                            On Error Resume Next
                            If ww = 5 Then
                            PlaceValue2UDTArray var(v), ss$, p, I
                            Else
                            PlaceValue2UDT var(v), ss$, p
                            End If
                            If Err Then
                                MyEr Err.Description, Err.Description
                                GoTo err000
                            End If
                        End If
                    End If
                ElseIf ww = 5 Then

                If IsExp(bstack, b$, p) Then
                    I = CLng(p)
                    If FastSymbol(b$, ")") Then GoTo Z1123698
                    SyntaxError
                    GoTo err000
                End If
                Else
                SyntaxError
                GoTo err000
                End If
                Else
                SyntaxError
                GoTo err000
                End If
                Else
                WrongOperator
                GoTo err000
                End If
                Else
                MissNumExpr
                GoTo err000
                End If
                End If
                ElseIf var(v) Is Nothing Then
                If ss$ = "->" Then
                 GoTo assignpointer
                End If
                ElseIf TypeOf var(v) Is Group Then
                If ss$ = "->" Then
                 GoTo assignpointer
                End If
                If var(v).IamApointer Then
                If var(v).link.IamFloatGroup Then
                    MyPush bstack, b$
                    Set bstack.lastobj = var(v).link
                    ProcessOper bstack, myobject, ss$, (0), 1
                    If Not bstack.lastobj Is Nothing Then
                    Set var(v).LinkRef = bstack.lastobj
                    Set bstack.lastobj = Nothing
                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
  
                    Else
                    GoTo here1234
                    End If
                    Else
                    W$ = var(v).lasthere + "." + var(v).GroupName
                End If
                End If
                Set myobject = bstack.soros
comeoper:
                Set bstack.Sorosref = New mStiva
                    If IsExp(bstack, b$, p) Then
                        If bstack.lastobj Is Nothing Then
                            bstack.soros.PushVal p
                        Else
                    If TypeOf bstack.lastobj Is VarItem Then
                    bstack.soros.DataOptional
                    Else
                    bstack.soros.DataObj bstack.lastobj
                    End If
                            Set bstack.lastobj = Nothing
                        End If
                    End If

                    NeoCall2 bstack, W$ + "." + ChrW(&H1FFF) + ss$ + "()", ok
                     Set bstack.Sorosref = myobject
                 Set myobject = Nothing
                    If Not ok Then
here1234:
                        If LastErNum = 0 Then MissOperator ss$
                        GoTo err000
                    End If
                Else
                    Set myobject = var(v)
                    
                    If CheckIsmArray(myobject) Then
                    If ss$ = "@@" Then
                            If FastPureLabel(b$, ss$, , True) Then
                                If Mid$(b$, 1, 1) = "#" Then ss$ = ss$ + "#": Mid$(b$, 1, 1) = " "
                              
                           Else
                                WrongOperator
                           End If
                        End If
                    

                        If IsExp(bstack, b$, p) Then
                            If Not bstack.lastobj Is Nothing Then
                                If TypeOf bstack.lastobj Is mArray Then
                                    Set usehandler = New mHandler
                                    usehandler.T1 = 3
                                    Set usehandler.objref = bstack.lastobj
                                    Set var(v) = usehandler
                                Else
                                Set pppp1 = myobject
                                    Set myobject = bstack.lastobj
                                    If CheckIsmArray(myobject) Then
                                        Set usehandler = New mHandler
                                        usehandler.T1 = 3
                                        Set usehandler.objref = myobject
                                        Set var(v) = usehandler
                                    ElseIf TypeOf myobject Is mHandler And ss$ <> vbNullString Then
                                        Set usehandler = myobject
                                        If usehandler.T1 = 4 Then
                                            pppp1.Compute2 p, ss$
                                        End If
                                    Else
NotArray1:
                                        NotArray
                                        GoTo err000
                                    End If
                                End If
                            Else
                            
                                myobject.Compute2 p, ss$
                            End If
                            Set usehandler = Nothing
                            Set myobject = Nothing
                            Set bstack.lastobj = Nothing
                        Else
                        
                            myobject.Compute3 ss$
                            Set myobject = Nothing
                            Set bstack.lastobj = Nothing
                        End If
                    ElseIf TypeOf myobject Is mHandler Then
                        Set usehandler = myobject
                        If usehandler.T1 = 4 Then
                            If usehandler.ReadOnly Then
                                    ReadOnly
                                    GoTo err000
                            ElseIf ss$ = "++" Then
                                If usehandler.index_start < usehandler.objref.count - 1 Then
                                    usehandler.index_start = usehandler.index_start + 1
                                    usehandler.objref.Index = usehandler.index_start
                                    usehandler.index_cursor = usehandler.objref.Value
                                End If
                            ElseIf ss$ = "--" Then
                                If usehandler.index_start > 0 Then
                                    usehandler.index_start = usehandler.index_start - 1
                                    usehandler.objref.Index = usehandler.index_start
                                    usehandler.index_cursor = usehandler.objref.Value
                                End If
                            ElseIf ss$ = "-!" Then
                                usehandler.sign = -usehandler.sign
                            Else
                                NoOperatorForThatObject ss$
                                GoTo err000
                            End If
                            Set usehandler = Nothing
                        ElseIf usehandler.T1 = 2 Then
contstruct11:
contstruct11err:
                            'MyEr "not defined yet", "δεν έχει οριστεί ακόμα"
                            If ww = 1 Then Mid$(b$, 1, 1) = "|"
                        Set usehandler = var(v)
                        
                        If Not TakeOffset(bstack, usehandler, b$, sp, p, ww - 8) Then
                        
                            GoTo err000
                        End If
                        Else
                            NoOperatorForThatObject ss$
                            GoTo err000
                        End If
                    Else
                        NoOperatorForThatObject ss$
                        GoTo err000
                    End If
                End If
            End If
                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                 

            

        ElseIf Not bstack.StaticCollection Is Nothing Then
            If bstack.ExistVar(W$, ok) Then
                If FastOperator(b$, "=", I) Then
                
                    If IsExp(bstack, b$, p) Then
checkobject1:
                        Set myobject = bstack.lastobj
                        If CheckIsmArray(myobject) Then
                            Set bstack.lastobj = myobject
                            bstack.SetVarobJ W$, bstack.lastobj
                        ElseIf CheckLastHandler(myobject) Then
                            Set usehandler = myobject
                            If usehandler.T1 = 2 Then
                                bstack.SetVarobJ W$, myobject
                            ElseIf usehandler.T1 = 1 Then
                                Set usehandler = New mHandler
                                usehandler.T1 = 1
                                Set usehandler.objref = myobject
                                Set myobject = usehandler
                                Set usehandler = Nothing
                                bstack.SetVarobJ W$, myobject
                            ElseIf usehandler.T1 = 3 Then
                                bstack.SetVarobJ W$, myobject
                            ElseIf usehandler.T1 = 4 Then
                                bstack.SetVarobJ W$, myobject
                            Else
                               GoTo aproblem1
                            End If
                        
                        ElseIf ok Then
                        bstack.ReadVar W$, sp
                        If TypeOf sp Is mHandler Then
                        Set usehandler = sp
                        If usehandler.T1 = 4 Then
                         Set sp = usehandler.objref.SearchValue(p, ok)
                         
                            If Not ok Then GoTo aproblem1
                           
                            bstack.SetVarobJ W$, sp
                                    Else
                                    GoTo aproblem1
                                    End If
        
                        Else
                        GoTo aproblem1
                        End If
                      
                        
                        Else
                            bstack.SetVar W$, p
                        End If
                        Set myobject = Nothing
                        Set bstack.lastobj = Nothing
                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                 

                    ElseIf IsStrExp(bstack, b$, ss$, False) Then ' Len(bstack.tmpstr) = 0
                            If ss$ = vbNullString Then
                            p = 0#
                            Else
                            p = ss$
                            End If
                            GoTo checkobject1
                    Else
                        If ok Then
                        bstack.ReadVar W$, sp
                        
                        If TypeOf sp Is Group Then
                        If Left$(b$, 1) = ">" Then
                        Set bstack.lastpointer = sp
                        Mid$(b$, 1, 1) = Chr$(3)
                         ExecuteVar = 10
                         Exit Function
                        Else
                        GoTo aproblem1
                        End If
                        Else
                        GoTo aproblem1
                        End If
                        Else
                        GoTo aproblem1
                        End If
                    End If
                Else
                    If InStr("/*-+~|", Mid$(b$, I, 1)) > 0 Then
                        If InStr("=+-!", Mid$(b$, I + 1, 1)) > 0 Then
                            ss$ = Mid$(b$, I, 2)
                            Mid$(b$, I, 2) = "  "
                        ElseIf Mid$(b$, I, 1) = "|" Then
                        Mid$(b$, I, 1) = " "
                        If FastPureLabel(b$, ss$, , True) = 1 Then
                            If Mid$(b$, 1, 1) = "#" Then ss$ = ss$ + "#": Mid$(b$, 1, 1) = " "
                        Else
                            WrongOperator
                        End If
                        
                        Else
                            ss$ = Mid$(b$, I, 1)
                            Mid$(b$, I, 1) = " "
                        End If
                    End If
                    If Right$(ss$, 1) = "=" Or Len(ss$) > 2 Then
                    If IsExp(bstack, b$, p) Then
                 If Not bstack.AlterVar(W$, p, ss$, False) Then GoTo err000
                 Else
                 GoTo aproblem1
                 End If
                 Else
                 If Not bstack.AlterVar(W$, p, ss$, False) Then GoTo err000
                 End If
                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                

                End If
                
                
            End If
            If FastOperator(b$, "=", I) Then ' MAKE A NEW ONE IF FOUND =
                If FastOperator(b$, ">", I + 1) Then
                    If GetVar(bstack, W$, v, True) Then
                        GoTo jumphere1
                    Else
                        SyntaxError
                        Set bstack.lastobj = Nothing
                        GoTo err000
                    End If
                Else
                v = globalvar(W$, p, , VarStat, temphere$)
                GoTo assignvalue
                End If
            ElseIf FastOperator(b$, "->", I, 2) Then
                GoTo jumpforpointer
            ElseIf GetVar(bstack, W$, v, True) Then
                    GoTo somethingelse
            End If
        ElseIf FastOperator(b$, "=", I) Then ' MAKE A NEW ONE IF FOUND =
              newid = True
jumpiflocal:
            If FastOperator(b$, ">", I) Then
            If GetVar(bstack, W$, v, True, , , , useType) Then
jumphere1:
            If Not var(v) Is Nothing Then
            If TypeOf var(v) Is Group Then
                GoTo noexpression1
                End If
            End If
            End If
            OnlyForGroupPointers
            GoTo err000
            ElseIf AscW(W$) = &H1FFF Then
            If GetVar(bstack, W$, v, True, , , , useType) Then newid = False: GoTo assignvalue
            If GetlocalVar(W$, v) Then useType = varhash.vType(varhash.Index): newid = False: GoTo assignvalue
            Else
            If noVarStat Then
                If GetlocalVar(W$, v) Then useType = varhash.vType(varhash.Index): newid = False: GoTo assignvalue
            End If
            v = globalvar(W$, p, , VarStat, temphere$)

            GoTo assignvalue
            End If
        ElseIf FastOperator(b$, "->", I, 2) Then ' MAKE A NEW ONE IF FOUND =
jumpforpointer:
            If AscW(W$) = &H1FFF Then
            If GetVar(bstack, W$, v, True) Then GoTo assignpointer
            If GetlocalVar(W$, v) Then GoTo assignpointer
            Else
            If GetVar(bstack, W$, v, True, , , , , ok) Then
            If ok Then
            v = globalvar(W$, p, , VarStat, temphere$)
            End If
            Else
            v = globalvar(W$, p, , VarStat, temphere$)
            End If
            GoTo assignpointer
            End If
        ElseIf GetVar(bstack, W$, v, True) Then
        ' CHECK FOR GLOBAL
            GoTo somethingelse
        End If
    ElseIf noVarStat Then
        GoTo checkhereClass
    End If
End If
'***********************
Exit Function
assignpointer:
                    If GetPointer(bstack, b$) Then
                        If MyIsObject(var(v)) Then
                            If var(v) Is Nothing Then
                                GoTo jumpgrouphere
                            ElseIf var(v).IamApointer Then
jumpgrouphere:
                                Set var(v) = bstack.lastpointer
                                ElseIf var(v).FieldsCount > 0 Or var(v).FuncList <> vbNullString Then
                                    CanyAssignPointer2Group
                                    Set bstack.lastpointer = Nothing
                                    Set bstack.lastobj = Nothing  '???
                                    GoTo err000
                                Else
                                    Set var(v) = bstack.lastpointer
                                End If
                            Else
                                Set var(v) = bstack.lastpointer
                            End If
                            Set bstack.lastpointer = Nothing
                            Set bstack.lastobj = Nothing  '???
                        Else
                            MissingPointer
                            Set bstack.lastobj = Nothing
                            GoTo err000
                        End If
                        If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                
'***********************
'' Case 2
'' no case 2 here
Case3:
If AscW(W$) = 46 Then
               If Not expanddot(bstack, W$) Then
               ManyDots
                GoTo err000
               End If
Else
Select Case CheckThis(bstack, W$, b$, v, Lang)
Case 0
useType = True
Case 1
useType = True
GoTo assignvaluestr1
Case -1
GoTo err000
End Select
End If

        ss$ = vbNullString
        If Left$(b$, 1) = "_" Then
        If nchr <> 61 Then
            SyntaxError
            Exit Function
        End If
        ss$ = "g"
        Mid$(b$, 1, 1) = " "
        GoTo again12345
        ElseIf FastSymbol(b$, ".") Then
        If GetVar(bstack, W$, v) Then  'GetlocalVar(w$, V) Then
             If MaybeIsSymbol(b$, "-+*/<~") Then
            If Right$(var(v), 1) = ")" Then
              b$ = var(v) + b$
              Else
             bstack.tmpstr = var(v) + Left$(b$, 1)
               BackPort b$
             End If
             ElseIf lookOne(b$, "=") Then
             If Right$(var(v), 1) = ")" Then
              b$ = var(v) + b$
              sss = Len(b$)
             Else
              bstack.tmpstr = var(v) + "=_"
              BackPort b$
              End If
             Else

             IsLabelDot temphere$, b$, W$
              If lookOne(b$, "=") Then
             W$ = var(v) + "." + W$
              bstack.tmpstr = W$ + "=_"
              BackPort b$
              ElseIf MaybeIsSymbol(b$, "-+*/<~") Then
               bstack.tmpstr = var(v) + "." + W$ + Left$(b$, 1)
               BackPort b$
              ElseIf Len(W$) = 0 Then
               bstack.tmpstr = var(v) + " " + Left$(b$, 1)
               BackPort b$
              Else
               bstack.tmpstr = var(v) + "." + W$ + " " + Left$(b$, 1)
               BackPort b$
               
              End If
             End If
              ExecuteVar = 5: Exit Function
        Else
                UnKnownWeak W
        End If
        End If
        I = MyTrimL(b$)
        If I > Len(b$) Then
        
        ElseIf InStr("/*-+=~^&|<>", Mid$(b$, I, 1)) > 0 Then
        
                    If InStr("/*-+=~^&|<>!", Mid$(b$, I + 1, 1)) > 0 Then
                        ss$ = Mid$(b$, I, 2)
                        If ss$ = "=&" Then
                        ss$ = "="
                        Mid$(b$, I, 1) = " "
                        Else
                        Mid$(b$, I, 2) = "  "
                        End If
                        If ss$ = "<=" Then ss$ = "g"
                        
                    Else
                        ss$ = Mid$(b$, I, 1)
                        Mid$(b$, I, 1) = " "
                    End If
                End If
                If ss$ <> "" Then
                    If ss$ = "=" Then
                        If VarStat Then
                            If IsStrExp(bstack, b$, ss$) Then
                                GoTo cont184575
                            Else
                                NoValueForVar W$
                                GoTo err000
                            End If
                        Else
                            If NewStat Then
                                    If IsStrExp(bstack, b$, ss$) Then globalvar W$, ss$, , VarStat, temphere$
                            Else
                                If AscW(W$) = &H1FFF Then
                                If GetVar(bstack, W$, v, True) Then GoTo assignvaluestr1
                                If GetlocalVar(W$, v) Then GoTo assignvaluestr1
                            ElseIf GetlocalVar(W$, v) Then
assignvaluestr1:
                  ''                 If TypeOf var(v) Is Group Then w$ = varhash.lastkey  ' don't know yet

                                If IsStrExp(bstack, b$, ss$) Then
str99399:
                                    If VarTypeName(var(v)) = mProp Then
                                        If FastSymbol(b$, "@") Then
                                            If IsExp(bstack, b$, sp) Then
                                                var(v).Index = sp: sp = 0
                                            ElseIf IsStrExp(bstack, b$, sw$, Len(bstack.tmpstr) = 0) Then
                                                var(v).Index = sw$: sw$ = vbNullString
                                            End If
                                            var(v).UseIndex = True
                                        End If
                                        var(v).Value = ss$
                                    ElseIf Typename$(bstack.lastobj) = "lambda" Then
                                        If MemInt(VarPtr(var(v))) = vbObject Then
                                        If VarTypeName(var(v)) = "Constant" Then GoTo itsAconstant
                                        End If
                                        If VarTypeName(var(v)) = "lambda" Then
                                            Set var(v) = bstack.lastobj
                                        Else
                                            If here$ = vbNullString Or VarStat Or NewStat Then
                                                GlobalSub W$ + "()", "", , , v
                                            Else
                                                GlobalSub here$ + "." + bstack.GroupName + W$ + "()", "", , , v
                                            End If
                                            Set var(v) = bstack.lastobj
                                        End If
                                        Set bstack.lastobj = Nothing
                                    ElseIf Typename$(var(v)) = mGroup Then
                                        If var(v).HasSet Then
                                            Set myobject = bstack.soros
                                            Set bstack.Sorosref = New mStiva
                                            If bstack.lastobj Is Nothing Then
                                                bstack.soros.PushStr ss$
                                            Else
                                                If TypeOf bstack.lastobj Is VarItem Then
                                                    bstack.soros.DataOptional
                                                Else
                                                    bstack.soros.DataObj bstack.lastobj
                                                End If
                                                Set bstack.lastobj = Nothing
                                            End If
                                            NeoCall2 bstack, Left$(W$, Len(W$) - 1) + "." + ChrW(&H1FFF) + ":=()", ok
                                            Set bstack.Sorosref = myobject
                                            Set myobject = Nothing
                                        Else
                                            If bstack.lastobj Is Nothing Then
                                                NeedAGroupInRightExpression
                                                GoTo err000
                                            ElseIf TypeOf bstack.lastobj Is Group Then
                                                Set myobject = bstack.lastobj
                                                Set bstack.lastobj = Nothing
                                                ss$ = bstack.GroupName
                                                If var(v).HasValue Or var(v).HasSet Then
                                                    PropCantChange
                                                    GoTo err000
                                                Else
                                                    W$ = Left$(W$, Len(W$) - 1)
                                                    If Len(var(v).GroupName) > Len(W$) Then
                                                        UnFloatGroupReWriteVars bstack, W$, v, myobject
                                                    Else
                                                        bstack.GroupName = Left$(W$, Len(W$) - Len(var(v).GroupName) + 1)
                                                        If Len(var(v).GroupName) > 0 Then
                                                            W$ = Left$(var(v).GroupName, Len(var(v).GroupName) - 1)
                                                            UnFloatGroupReWriteVars bstack, W$, v, myobject
                                                        Else
                                                            GroupWrongUse
                                                            GoTo err000
                                                        End If
                                                    End If
                                                End If
                                                Set myobject = Nothing
                                                bstack.GroupName = ss$
                                            Else
                                                GroupCantSetValue
                                            End If
                                        End If
                                    Else
                                        If CheckVarOnlyNo(var(v), ss$) Then
                                            If VarTypeName(var(v)) = "Constant" Then
itsAconstant:
                                                If myVarType(var(v).Value, vbEmpty) Then
                                                    If bstack.lastobj Is Nothing Then
                                                        var(v).DefineOnce ss$
                                                        If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                                                    Else
                                                        NoObjectAssign
                                                        MissNumExpr
                                                        GoTo err000
                                                    End If
                                                Else
                                                    CantAssignValue
                                                End If
                                            Else
                                                ExpectedObj VarTypeName(var(v))
                                            End If
                                            GoTo err000
                                        End If
                                    End If
                                ElseIf IsExp(bstack, b$, p, , True) Then
                                    Assign ss$, p
                                    GoTo str99399
                                End If
                            ElseIf Not bstack.StaticCollection Is Nothing Then
                                If bstack.ExistVar(W$) Then
                                    If IsStrExp(bstack, b$, ss$, Len(bstack.tmpstr) = 0) Then bstack.SetVar W$, ss$ Else GoTo aproblem1
                                ElseIf IsStrExp(bstack, b$, ss$, Len(bstack.tmpstr) = 0) Then
                                    GoTo cont184575
                                End If
                            ElseIf IsStrExp(bstack, b$, ss$, False) Then
cont184575:
                                If bstack.lastobj Is Nothing Then
                                    globalvarStr W$, ss$, , VarStat, temphere$
                                Else
                                    If Typename$(bstack.lastobj) = "lambda" Then
                                        If NewStat Then
                                            NoNewLambda
                                            Exit Function
                                        Else
                                            I = 0
                                            If strfunid.Find(W$ + "(", (I)) Then
                                            strfunid.ItemCreator W$ + "(", -2
                                        End If
                                        If VarStat Then
                                            I = globalvar(W$, p, , VarStat, temphere$)
                                        Else
                                        If Not GetVar(bstack, W$, I, True) Then I = globalvar(W$, p, , , temphere$)
                                    End If
                                    If VarTypeName(var(I)) = "Constant" Then
                                        CantAssignValue
                                        GoTo err000
                                    End If
                                    If here$ = vbNullString Or VarStat Then
                                        GlobalSub W$ + "()", "", , , I
                                    Else
                                        GlobalSub here$ + "." + bstack.GroupName + W$ + "()", "", , , I
                                    End If
                                End If
                                Set myobject = bstack.lastobj
                                Set bstack.lastobj = Nothing
                                If I <> 0 Then
                                    Set var(I) = myobject
                                    Set myobject = Nothing
                                End If
                            ElseIf Typename$(bstack.lastobj) = mGroup Then
                                If Not ProcGroup(200 + (VarStat Or isglobal), bstack, W$, Lang) Then
                                    GoTo err000
                            End If
                        Else
                            NoValueForVar W$
                            GoTo err000
                        End If
                    End If
                Else
                    NoValueForVar W$
                    GoTo err000
                End If
            End If
        End If
    Else    ' g
again12345:
        If GetVar(bstack, W$, v, ss$ = "g") Then
stroper001:
            sw$ = ss$
            p = W$
            W$ = varhash.lastkey
            If IsExp(bstack, b$, p) Then
                Assign ss$, p
                If sw = "+=" Then Set bstack.lastobj = Nothing
                GoTo strcont111
            End If
            If IsStrExp(bstack, b$, ss$, False) Then
strcont111:
                 If VarTypeName(var(v)) = mProp Then
                        If FastSymbol(b$, "@") Then
                            If IsExp(bstack, b$, sp) Then
                            var(v).Index = sp: sp = 0
                            ElseIf IsStrExp(bstack, b$, sw$, Len(bstack.tmpstr) = 0) Then
                            var(v).Index = sw$: sw$ = vbNullString
                            End If
                             var(v).UseIndex = True
                        End If
                        var(v).Value = ss$
                  ElseIf VarTypeName(var(v)) = "Constant" Then
                  If myVarType(var(v).Value, vbEmpty) Then
                  var(v).DefineOnce ss$
                  Else
                      CantAssignValue
                  End If
                
                   ElseIf Not bstack.lastobj Is Nothing Then
                            If TypeOf bstack.lastobj Is lambda Then
                             Set var(v) = bstack.lastobj
                             GlobalSub W$ + "()", "", , , v
                             Set bstack.lastobj = Nothing
                   
                    Else
                        NoValueForVar W$
                    End If
                    ElseIf Typename$(var(v)) = mGroup Then
                           If sw$ = "g" Then
                           sw$ = ":="
                           If Not var(v).HasSet Then GroupCantSetValue: GoTo err000
                           End If
                           Set myobject = bstack.soros
                            Set bstack.Sorosref = New mStiva
                           If bstack.lastobj Is Nothing Then
                                bstack.soros.PushStr ss$
                            Else
                                If TypeOf bstack.lastobj Is VarItem Then
                                bstack.soros.DataOptional
                                Else
                                bstack.soros.DataObj bstack.lastobj
                                End If
                                Set bstack.lastobj = Nothing
                            End If
a325674:
                            
                            NeoCall2 bstack, Left$(W$, Len(W$) - 1) + "." + ChrW(&H1FFF) + sw$ + "()", ok
                             Set bstack.Sorosref = myobject
                            Set myobject = Nothing
                            If Not ok Then GoTo here1234
                                
                Else
                If LenB(sw$) = 0 Or sw$ = "g" Or sw$ = "+=" Then
               CheckVar var(v), ss$, sw$ = "+="
               Else
                            NoValueForVar W$
                            GoTo err000
               End If
               
               End If
               Else
               ' check
                If Typename$(var(v)) = mGroup Then
                    Set myobject = bstack.soros
                    Set bstack.Sorosref = New mStiva
                    GoTo a325674
                ElseIf MemInt(VarPtr(var(v))) = vbString Then
                    MissStringExpr
                    NoValueForVar CStr(p)
                    GoTo err000
                End If
               
            End If
        Else
            If ss$ = "g" Then ss$ = vbNullString:   GoTo again12345
            Nosuchvariable W$
          End If
    End If
Else
        If VarStat Or NewStat Or noVarStat Then
            
            globalvar W$, ss$, , VarStat, temphere$
  
          
                    sss = Len(b$)
                ExecuteVar = 4: Exit Function
                
        End If

                 NoValueForVar W$
                    GoTo err000
             
End If
 ExecuteVar = 7
Exit Function
Case4:
If AscW(W$) = 46 Then
               If Not expanddot(bstack, W$) Then
               ManyDots
                GoTo err000
               End If
Else
Select Case CheckThis(bstack, W$, b$, v, Lang)
Case 1
useType = True
GoTo assignvalue100
Case -1
GoTo err000
End Select

End If
        If Left$(b$, 1) = "_" Then
        If nchr <> 61 Then
            SyntaxError
            Exit Function
        End If
        ss$ = "g"
        Mid$(b$, 1, 1) = " "
        GoTo again1234567
        ElseIf MaybeIsSymbol(b$, "=-+*/<~") Then
    If FastSymbol(b$, "=") Then
    
    If VarStat Then
    
    If IsExp(bstack, b$, p) Then
                                       
                globalvar W$, p, , VarStat, temphere$
                Else
                      If LastErNum <> -2 Then
                     NoValueForVar W$
                    GoTo err000
                     End If
            
            End If
    Else
        If AscW(W$) = &H1FFF Then
        If GetVar(bstack, W$, v, True) Then GoTo assignvalue100
        If varhash.Find2(here$ + "." + myUcase(W$), v, useType) Then GoTo assignvalue100
        ElseIf varhash.Find2(here$ + "." + myUcase(W$), v, useType) Then
assignvalue100:
                If IsExp(bstack, b$, p) Then
                If VarTypeName(var(v)) = mProp Then
                  If FastSymbol(b$, "@") Then
                        If IsExp(bstack, b$, sp) Then
                            var(v).Index = sp: sp = 0
                        ElseIf IsStrExp(bstack, b$, ss$, Len(bstack.tmpstr) = 0) Then
                        var(v).Index = ss$: ss$ = vbNullString
                        End If
                         var(v).UseIndex = True
                        End If
                        var(v).Value = MyRound(p)
                ElseIf Not bstack.lastobj Is Nothing Then
                        If TypeOf bstack.lastobj Is lambda Then
                            If VarTypeName(var(v)) = "lambda" Then
                                 Set var(v) = bstack.lastobj
                            Else
                                GlobalSub W$ + "()", "", , , v
                                Set var(v) = bstack.lastobj
                            End If
                            Set bstack.lastobj = Nothing
                        Else
                            
                            ExpectedObj VarTypeName(var(v))
                            GoTo err000
                            End If
               ElseIf MyIsObject(var(v)) Then
                        If TypeOf var(v) Is Constant Then
                              If myVarType(var(v).Value, vbEmpty) Then
                                    If bstack.lastobj Is Nothing Then
                                          var(v).DefineOnce MyRound(p)
                                          If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                                    Else
                                          NoObjectAssign
                                          MissNumExpr
                                          GoTo err000
                                    End If
                       
                                Else
                            CantAssignValue
                            End If
                        Else
                           ExpectedObj VarTypeName(var(v))
                           
                        End If
                        GoTo err000
                Else
                        p = MyRound(p)
                        If useType Then
                                   If AssignTypeNumeric(p, VarType(var(v))) Then
                                       var(v) = p
                                   Else
                                       GoTo err000
                                   End If
                       Else
                     var(v) = p
                     End If
                End If
 
                If Err.Number = 6 Then Exec1 = 0: ExecuteVar = 1: Exit Function
                On Error GoTo 0
                End If
            ElseIf Not bstack.StaticCollection Is Nothing Then
            If bstack.ExistVar(W$) Then
            If IsExp(bstack, b$, p) Then bstack.SetVar W$, MyRound(p) Else GoTo aproblem1
            
            ElseIf IsExp(bstack, b$, p) Then
            GoTo abc2345
            End If
            
            ElseIf IsExp(bstack, b$, p) Then
abc2345:
            If Not bstack.lastobj Is Nothing Then
            If TypeOf bstack.lastobj Is lambda Then
                            v = globalvar(W$, p, , VarStat, temphere$)
                            If NewStat Then  '' ???
                                            NoNewLambda
                                            Exit Function
                                        Else
                                               If here$ = vbNullString Or VarStat Then
                                                GlobalSub W$ + "()", "", , , v
                                            Else
                                                GlobalSub here$ + "." + bstack.GroupName + W$ + "()", "", , , v
                                            End If
                                        End If
                                Set var(v) = bstack.lastobj
                                Set bstack.lastobj = Nothing
                                Else
                                    NoValueForVar W$
                                    GoTo err000
                                End If
           Else
            p = MyRound(p)
            
            globalvar W$, p, , VarStat, temphere$
    End If
          
                Else
                      If LastErNum <> -2 Then
aproblem1:
                     NoValueForVar W$
                    GoTo err000
                     End If
            
            End If
    End If
    Else
        ss$ = proc101(b$)  ' procedure too long...problem
again1234567:
        If GetVar(bstack, W$, v, ss$ = "g") Then
        'NOT YET FOR PropReference
      If MyIsObject(var(v)) Then
            If VarTypeName(var(v)) = mProp Then
                    GoTo err000
            End If
            If TypeOf var(v) Is Constant Then
                NoOperatorForThatObject ss$
                GoTo err000
            End If
            
      End If
            If Len(ss$) = 1 Then
                        If IsExp(bstack, b$, p) Then
                        AssignTypeNumeric sp, VarType(var(v))
                                On Error Resume Next
                                Select Case ss$
                                Case "=", "g"
                                    var(v) = MyRound(p)
                                Case "+"
                                    var(v) = MyRound(p) + var(v)
                                Case "*"
                                    var(v) = MyRound(MyRound(p) * var(v))
                                Case "-"
                                    var(v) = var(v) - MyRound(p)
                                Case "/"
                                    If MyRound(p) = 0 Then Exec1 = 0: ExecuteVar = 1: Exit Function
                                    var(v) = MyRound(var(v) / MyRound(p))
                                    
         
                   
                                End Select
                                If Err.Number = 6 Then Exec1 = 0: ExecuteVar = 1: Exit Function
                                On Error GoTo 0
                                AssignTypeNumeric var(v), VarType(sp)
                                GoTo checksyntax
                        Else
                                Exec1 = 0: ExecuteVar = 1: Exit Function
                        End If
            Else
                If ss$ = "++" Then
                    var(v) = 1 + var(v)
                ElseIf ss$ = "--" Then
                    var(v) = var(v) - 1
                ElseIf ss$ = "-!" Then
                    var(v) = -var(v)
                Else
                Select Case VarType(var(v))
                    Case vbBoolean
                        var(v) = Not CBool(var(v))
                    Case vbCurrency
                        var(v) = CCur(Not CBool(var(v)))
                    Case vbDecimal
                        var(v) = CDec(Not CBool(var(v)))
                    Case Else
                        var(v) = CDbl(Not CBool(var(v)))
                        End Select
                End If
  
            End If
            GoTo checksyntax
     Else
        If v = -1 Then
         If Len(ss$) = 1 Then If Not IsExp(bstack, b$, p) Then GoTo err000
                    
             If Not bstack.AlterVar(W$, p, ss$, True) Then GoTo err000
                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck

    Else
        If ss$ = "g" Then ss$ = "=":   GoTo again1234567
       NoValueForVar W$
       GoTo err000
       End If
      End If
   
    End If
Else
    If VarStat Or NewStat Or noVarStat Or noVarStat Then
              p = 0#
              If IsLabelSymbolNew(b$, "ΩΣ", "AS", Lang) Then
  
               If IsLabelSymbolNew(b$, "ΑΡΙΘΜΟΣ", "DECIMAL", Lang) Then
                    If FastSymbol(b$, "=") Then If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    p = CDec(p)
            ElseIf IsLabelSymbolNew(b$, "ΔΙΠΛΟΣ", "DOUBLE", Lang) Then
                            If FastSymbol(b$, "=") Then If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                p = CDbl(p)
            ElseIf IsLabelSymbolNew(b$, "ΑΠΛΟΣ", "SINGLE", Lang) Then
                        If FastSymbol(b$, "=") Then If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                p = CSng(p)
            ElseIf IsLabelSymbolNew(b$, "ΛΟΓΙΚΟΣ", "BOOLEAN", Lang) Then
                    If FastSymbol(b$, "=") Then If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                p = CBool(p)
            ElseIf IsLabelSymbolNew(b$, "ΜΑΚΡΥΣ", "LONG", Lang) Then
                If IsLabelSymbolNew(b$, "ΜΑΚΡΥΣ", "LONG", Lang) Then
                    If FastSymbol(b$, "=") Then If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    p = cInt64(p)
                Else
                    If FastSymbol(b$, "=") Then If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    p = CLng(p)
                End If
            ElseIf IsLabelSymbolNew(b$, "ΑΚΕΡΑΙΟΣ", "INTEGER", Lang) Then
                    If FastSymbol(b$, "=") Then If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                p = CInt(p)
            ElseIf IsLabelSymbolNew(b$, "ΜΙΓΑΔΙΚΟΣ", "COMPLEX", Lang) Then
                    If FastSymbol(b$, "=") Then
                    If Not FastSymbol(b$, "(") Then missNumber: Exit Function
                    If Not IsNumberD2(b$, sp) Then missNumber: Exit Function
                    If Not FastSymbol(b$, ",") Then missNumber: Exit Function
                    If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    b$ = NLtrim$(b$)
                    If Len(b$) >= 2 Then
                    If Not UCase(Left$(b$, 2)) = "I)" Then Mid$(b$, 1, 2) = "  ": missNumber: Exit Function
                    Mid$(b$, 1, 2) = "  "
                    End If
                    Else
                        missNumber
                        Exit Function
                    End If
                    p = nMath2.cxNew(CDbl(sp), CDbl(p))
            ElseIf IsLabelSymbolNew(b$, "ΛΟΓΙΣΤΙΚΟΣ", "CURRENCY", Lang) Then
                If FastSymbol(b$, "=") Then If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                p = CCur(p)
            ElseIf IsLabelSymbolNew(b$, "ΑΤΥΠΟΣ", "VARIANT", Lang) Then
                If FastSymbol(b$, "=") Then
                    If Not IsNumberD2(b$, p) Then
                        If ISSTRINGA(b$, ss$) Then
                            p = ss$
                        Else
                            missNumber
                            Exit Function
                        End If
                    End If
                End If
            ElseIf IsLabelSymbolNew(b$, "ΨΗΦΙΟ", "BYTE", Lang) Then
                If FastSymbol(b$, "=") Then If Not IsNumberD2(b$, p) Then missNumber: Exit Function
            ElseIf IsLabelSymbolNew(b$, "ΗΜΕΡΟΜΗΝΙΑ", "DATE", Lang) Then
                If FastSymbol(b$, "=") Then
                    If IsNumberD2(b$, p) Then
                        p = CDate(p)
                    ElseIf ISSTRINGA(b$, ss$) Then
                        p = CDate(ss$)
                    Else
                        missNumber
                        Exit Function
                    End If
                Else
                    p = CDate(0#)
            End If
            Else
                If Not IsEnumAs(bstack, b$, p) Then
                    ExpectedEnumType
                    NoTypeFound
                    Exit Function
                End If
            End If
     
    End If
    
    p = Int(p)
        globalvar W$, p, , VarStat, temphere$
        sss = Len(b$): ExecuteVar = 4: Exit Function
    Else
        NoValueForVar W$
        GoTo err000
    End If
End If
Exit Function
    

case5:
    If AscW(W$) = 46 Then
        If Not expanddot(bstack, W$) Then
            ManyDots
            GoTo err000
        End If
    End If
    If funid.Find(W$, I) Then
        If I > 0 Then funid.ItemCreator W$, -I
    End If
    If VarStat Or NewStat Or noVarStat Then
        If noVarStat Then
            If neoGetArray(bstack, W$, ppppAny, , , , True) Then
            If Not TypeOf ppppAny Is mArray Then
                WrongType
                Exit Function
            End If
            GlobalArrResize ppppAny, bstack, W$, b$, v
            If IsLabelSymbolNew(b$, "ΩΣ", "AS", Lang) Then
            ww = IsLabelOnly(b$, sw$)
            End If
            If FastSymbol(b$, "=") Then
                If IsExp(bstack, b$, p) Then
                    Set pppp = ppppAny
                    pppp.SerialItem p, 0, 3
                ElseIf IsStrExp(bstack, b$, sw$) Then
                    Set pppp = ppppAny
                    p = sw$
                    pppp.SerialItem p, 0, 3
                End If
                Set ppppAny = Nothing
                Set pppp = Nothing
            End If
            
            Else
            
            MakeArray bstack, W$, 5, b$, pppp, True
            End If
        
        Else
        MakeArray bstack, W$, 5, b$, pppp, NewStat, VarStat
        End If
        sss = Len(b$): ExecuteVar = 4: Exit Function
    End If
    aheadstatusSkipParam b$, I
    I = I + 1
    If MaybeIsSymbol3lot(b$, b12345, I) Or I > Len(b$) Then
        If Mid$(b$, I, 2) = ":=" Then GoTo arr1111
        If Mid$(b$, I, 2) = "/=" Then GoTo arr1111
        bstack.tmpstr = ss$
        ExecuteVar = 2  ' GoTo autogosub
        Exit Function
    End If
arr1111:
    If neoGetArray(bstack, W$, ppppAny, , , , True) Then
againarray:
        If ppppAny Is Nothing Then
            GoTo err000
        End If
        If Not ppppAny.Arr Then
            If Not NeoGetArrayItem(ppppAny, bstack, W$, v, b$, , , , True, idx) Then GoTo errorarr
        ElseIf FastSymbol(b$, ")") Then
            Set pppp = ppppAny
            'need to found an expression
            If FastSymbol(b$, "=") Then
                If IsExp(bstack, b$, p) Then
                    If Not bstack.lastobj Is Nothing Then
                        If TypeOf bstack.lastobj Is mHandler Then
                            Set usehandler = bstack.lastobj
                            If usehandler.indirect >= 0 Then
                                ' no copy..just a reference
                                Set bstack.lastobj = var(usehandler.indirect)
                            Else
                                Set bstack.lastobj = usehandler.objref
                            End If
                            Set usehandler = Nothing
                            If TypeOf bstack.lastobj Is mArray Then
                                Set pppp1 = bstack.lastobj
                                pppp1.CopyArray pppp
                                pppp.Final = False
                            Else
                                GoTo NotArray1
                            End If
                        Else
                            Set pppp1 = bstack.lastobj
                            pppp1.CopyArray pppp
                            pppp.Final = False
                        End If
                        Set bstack.lastobj = Nothing
                        If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                    Else
                        Set pppp1 = New mArray: pppp1.PushDim (1): pppp1.PushEnd
                        pppp1.SerialItem 0, 2, 9
                        pppp1.Arr = True
                        If bstack.lastobj Is Nothing Then
                            pppp1.item(0) = p
                        Else
                            Set pppp1.item(0) = bstack.lastobj
                            Set bstack.lastobj = Nothing
                        End If
                        pppp1.CopyArray pppp
                        If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                    End If
                Else
                    SyntaxError
                End If
                GoTo err000
            End If
    ElseIf Not NeoGetArrayItem(ppppAny, bstack, W$, v, b$) Then
errorarr:
If LastErNum = -2 Then

Execute bstack, b$, True
GoTo err000
Else
Exec1 = 0
ExecuteVar = 1
Exit Function
End If
End If
'On Error Resume Next

If MaybeIsSymbol(b$, ":+-*/~|") Or v = -2 Then
here66678:
'****************************************
'*********************************************
    With ppppAny
        If ppppAny.Final Then CantAssignValue: GoTo err000
        If Not .Arr Then
            If v = -2 Then GoTo con123
        ElseIf MyIsObject(.item(v)) Then
            If v = -2 Then
con123:
                Set myobject = .item(v)
                If myobject.HasParametersSet Then
                    If myobject.HasSet Then
                        W$ = Left$(W$, Len(W$) - 1)
                        Set myobject = bstack.soros
                        Set bstack.Sorosref = New mStiva
                        PushParamGeneral bstack, b$
                        If Not FastSymbol(b$, ")", True) Then
                            Set bstack.Sorosref = myobject
                            GoTo err000
                        End If
                        If FastSymbol(b$, "=") Then
                            If IsExp(bstack, b$, p) Then
                                If bstack.lastobj Is Nothing Then
                                    bstack.soros.DataVal p
                                Else
                                    If TypeOf bstack.lastobj Is VarItem Then
                                        bstack.soros.DataOptional
                                    Else
                                        bstack.soros.DataObj bstack.lastobj
                                    End If
                                    Set bstack.lastobj = Nothing
                                End If
                                NeoCall2 bstack, W$ + "." + ChrW(&H1FFF) + ":=()", ok
                                Set bstack.Sorosref = myobject
                                Set myobject = Nothing
                                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                            Else
                                Set bstack.Sorosref = myobject
                                Set myobject = Nothing
                                GoTo noexpression
                            End If
                        End If
                        Set bstack.Sorosref = myobject
                        SyntaxError
                        GoTo err000
                    Else
                    End If
                Else
a1297654:
                    I = MyTrimL(b$)
                    If lookTwoSame(b$, "/") Then
                     ' error>>>
                        Exec1 = 0: If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                    ElseIf MaybeIsSymbol(Mid$(b$, I + 1, 1), "/*-+=~^&|<>?") Then
                        ss$ = Mid$(b$, I, 2)
                        Mid$(b$, 1, I + 1) = space(I + 1)
                    Else
                        ss$ = Mid$(b$, I, 1)
                        Mid$(b$, 1, I) = space(I)
                    End If
                    Set myobject = Nothing
                    If ss$ = "->" Then
                        If GetPointer(bstack, b$) Then
                            Set .item(v) = bstack.lastpointer
                            Set bstack.lastpointer = Nothing
                            Set bstack.lastobj = Nothing
                            If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                        End If
                    ElseIf ss$ = ":=" Then
                        GoTo contassignhere
                    Else
                        If .item(v).IamApointer Then
                            If .item(v).link.IamFloatGroup Then
                                MyPush bstack, b$
                                Set bstack.lastobj = .item(v).link
                            Else
                                W$ = .item(v).lasthere + "." + .item(v).GroupName
                                Set bstack.lastobj = Nothing
                                Set bstack.lastpointer = Nothing
                                GoTo comeoper
                            End If
                        Else
                            MyPush bstack, b$
                            Set bstack.lastobj = .item(v)
                        End If
                        ProcessOper bstack, myobject, ss$, (0), 1
                        If Not bstack.lastobj Is Nothing Then
                            If TypeOf bstack.lastobj Is Group Then
                                If .item(v).IamApointer Then
                                    Set .item(v).LinkRef = bstack.lastobj
                                Else
                                    Set .item(v) = bstack.lastobj
                                End If
                                Set bstack.lastobj = Nothing
                                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                            End If
                        End If
                    End If
                End If
            Else
                Set myobject = .item(v)
                If TypeOf myobject Is Group Then GoTo a1297654
                Set myobject = Nothing
            End If
        End If
    End With
    If Not TypeOf ppppAny Is mArray Then
        GoTo NotArray1
    End If
    Set pppp = ppppAny
    With pppp
        If FastSymbol(b$, ":=", , 2) Then
            If Not .Arr Then GoTo NotArray1
    ' new on rev 20
contassignhere:
            If GetData(bstack, b$, myobject) Then
                FeedArray pppp, v, myobject
                ExecuteVar = 7
            Else
                GoTo err000
            End If
            Exit Function
        ElseIf .IsStringItem(v) Then
            If FastSymbol(b$, "+=", , 2) Then
            If IsExp(bstack, b$, p) Then
                 If MemInt(VarPtr(p)) = vbString Then
                     SwapString2Variant sw$, p
                     p = Empty
                 Else
                     sw$ = vbNullString
                 End If
            ElseIf Not IsStrExp(bstack, b$, sw$, False) Then
                GoTo err000
            End If
                
                .item(v) = .item(v) + sw$
            Else
                WrongOperator
                GoTo err000
            End If
        Else
            AssignTypeNumeric sp, VarType(.item(v))
            If IsOperator0(b$, "++", 2) Then
                .item(v) = .itemnumeric(v) + 1
                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
            ElseIf IsOperator0(b$, "--", 2) Then
                .item(v) = .itemnumeric(v) - 1
                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
            ElseIf FastSymbol(b$, "+=", , 2) Then
                ww = Len(b$)
                If Not IsExp(bstack, b$, p) Then
                    If ww = Len(b$) Then
                        If IsStrExp(bstack, b$, ss$, False) Then
                             If Not .IsStringItem(v) Then
                                If .ItemTypeNum(v) = vbEmpty Then
                                    .ItemStr(v) = ss$
                                Else
                                    p = .itemnumeric(v)
                                    Assign sw$, p
                                    .ItemStr(v) = sw$ + ss$

                                End If
                             Else
                             .ItemStr(v) = .item(v) + ss$
                             End If
                             If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                        Else
                            GoTo err000
                        End If
                    Else
                        GoTo err000
                    End If
                Else
                .item(v) = .itemnumeric(v) + p
                End If
                
            ElseIf FastSymbol(b$, "-=", , 2) Then
                If Not IsExp(bstack, b$, p, flatobject:=True, nostring:=True) Then GoTo err000
                .item(v) = .itemnumeric(v) - p
            ElseIf FastSymbol(b$, "*=", , 2) Then
                If Not IsExp(bstack, b$, p, flatobject:=True, nostring:=True) Then GoTo err000
                .item(v) = .itemnumeric(v) * p
            ElseIf FastSymbol(b$, "/=", , 2) Then
                If Not IsExp(bstack, b$, p, flatobject:=True, nostring:=True) Then GoTo err000
                If p = 0# Then
                    DevZero
                Else
                    .item(v) = .itemnumeric(v) / p
                End If
            ElseIf IsOperator0(b$, "-!", 2) Then
                .item(v) = -.itemnumeric(v)
                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
            ElseIf IsOperator0(b$, "~") Then
                .Neg v
                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
            ElseIf IsOperator0(b$, "|") Then
                ' UDT
                If .ItemTypeNum(v) = vbUserDefinedType Then
                    ww = FastPureLabel(b$, ss$)
                    If ww = 1 Then
                        If FastSymbol(b$, "=") Then
                            If IsExp(bstack, b$, p) Then
                                If .PlaceValue2UDT(v, ss$, p) Then
                                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                                End If
                            ElseIf IsStrExp(bstack, b$, W$, False) Then
                                p = ""
                                SwapString2Variant W$, p
                                .PlaceValue2UDT v, ss$, p
                                 If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                            Else
                                GoTo noexpression
                            End If
                        Else
                            SyntaxError
                        End If
                        GoTo err000
                    ElseIf ww = 5 Then
                        If IsExp(bstack, b$, sp) Then
                        If FastSymbol(b$, ")") Then
                            If FastSymbol(b$, "=") Then
                            If IsExp(bstack, b$, p) Then
zzz123:
                                If .PlaceValue2UDTArray(v, ss$, p, CLng(sp)) Then
                                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                                End If
                            ElseIf IsStrExp(bstack, b$, W$, False) Then
                                p = ""
                                SwapString2Variant W$, p
                                GoTo zzz123
                            Else
                                GoTo noexpression
                            End If
                        Else
                            SyntaxError
                        End If
                        GoTo err000
                        
                        End If
                        End If
                    Else
                        GoTo noexpression
                    End If
                ElseIf FastPureLabel(b$, ss$, , True) Then
                    If Mid$(b$, 1, 1) = "#" Then ss$ = ss$ + "#": Mid$(b$, 1, 1) = " "
                    If IsExp(bstack, b$, p, flatobject:=True, nostring:=True) Then
                        If Int(p) = 0 Then
                            DevZero
                            GoTo err000
                        End If
                        SwapVariant sp, .item(v)
                        If Not readvarv(sp, ss$, p) Then
                            WrongOperator
                            GoTo err000
                        End If
                        .item(v) = sp
                    Else
                        GoTo noexpression
                    End If
                Else
                    WrongOperator
                    GoTo err000
                End If
            ElseIf FastSymbol(b$, "->", , 2) Then
                If Not GetPointer(bstack, b$) Then GoTo err000
                If Typename(bstack.lastobj) = mGroup Then
                    If Typename(.item(v)) <> mGroup Then
                        If bstack.lastpointer Is Nothing Then
                            Set .item(v) = bstack.lastobj
                        Else
                            Set .item(v) = bstack.lastpointer
                        End If
                    Else
                        If .item(v).IamApointer Then
                            If bstack.lastpointer Is Nothing Then
                                ExpectedPointer
                                GoTo err000
                            Else
                                Set .item(v) = bstack.lastpointer
                            End If
                        End If
                    End If
                    Set bstack.lastobj = Nothing
                    Set bstack.lastpointer = Nothing
                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                End If
               GoTo err000
            Else
               GoTo err000
            End If
        End If
        If Not MemInt(VarPtr(.item(v))) = MemInt(VarPtr(sp)) Then
            p = .itemnumeric(v)
            AssignTypeNumeric p, MemInt(VarPtr(sp))
            .item(v) = MyRound(p, 28)
        End If
    End With
    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
End If
    If IsOperatorNoRemove(b$, ".") Then
        If ppppAny.ItemType(v) = mGroup Then
            If ppppAny.item(v).IamApointer Then
                If ppppAny.item(v).link.IamFloatGroup Then
                    Mid$(b$, 1, 1) = ChrW(7)
                    Exec1 = SpeedGroup(bstack, ppppAny, "", W$, b$, v)
                    Exit Function
                Else
                    Set bstack.lastpointer = pppp.item(v)
                    Mid$(b$, 1, 1) = Chr(0)
                    ExecuteVar = 9
                    Exit Function
                End If
            Else
                Mid$(b$, 1, 1) = ChrW(7)
                Exec1 = SpeedGroup(bstack, ppppAny, "", W$, b$, v)
            End If
            If Exec1 = 0 Then GoTo err000
            If extreme Then GoTo NewCheck2 Else GoTo NewCheck
        End If
    ElseIf IsOperator(b$, "(") Then
        If ppppAny.ItemType(v) = myArray Then
            Set ppppAny = ppppAny.item(v)
            GoTo againarray
        ElseIf ppppAny.ItemType(v) = mGroup Then
again12568:
            Set myobject = bstack.soros
            Set bstack.Sorosref = New mStiva
            PushParamStraight bstack, b$
            If Not FastSymbol(b$, ")", True) Then
                    Set bstack.Sorosref = myobject
                    GoTo err000
            End If
            If Not FastSymbol(b$, "=", True) Then
                sss = 0
                ExecuteVar = 3: Exit Function
            End If
            If Not IsExp(bstack, b$, p) Then
                If LastErNum = -2 Then
                    Execute bstack, b$, True
                Else
                    MissNumExpr
                End If
                GoTo err000
            End If
            If bstack.lastobj Is Nothing Then
                bstack.soros.DataVal p
            Else
                bstack.soros.DataObj bstack.lastobj
                Set bstack.lastobj = Nothing
            End If
            Exec1 = SpeedGroup(bstack, pppp, "@READ2", "", b$, v)
            If Exec1 = 0 Then GoTo err000
            If extreme Then GoTo NewCheck2 Else GoTo NewCheck
        End If
    ElseIf Not FastSymbol(b$, "=") Then
        MissingSymbol "="
        sss = 0
        ExecuteVar = 3: Exit Function
    End If
    If Left$(b$, 1) = ">" Then
    
        If MyIsObject(ppppAny.item(v)) Then
            If TypeOf ppppAny.itemObject(v) Is Group Then
                If ppppAny.item(v).IamApointer Then
                    Set bstack.lastpointer = ppppAny.item(v)
                    If bstack.lastpointer.link.IamFloatGroup Then
                    Mid$(b$, 1, 1) = " "
                    If MaybeIsSymbol3(b$, "=", I) Then
                        Mid$(b$, I - 1, 2) = "  "
                        If IsExp(bstack, b$, p) Then GoTo here12500
                        GoTo err000
                    Else
                        If FastSymbol(b$, "(") Then
                        GoTo again12568
                    End If
                End If
                Mid$(b$, 1, 1) = ChrW(7)
                ExecuteVar = 10
            Else
                Mid$(b$, 1, 1) = " "
                If MaybeIsSymbol3(b$, "=", I) Then
                    Mid$(b$, I - 1, 2) = "  "
                    W$ = ppppAny.item(v).lasthere + "." + ppppAny.item(v).GroupName
                    If GetVar(bstack, W$, v, True) Then GoTo assigngroup
                Else
                    If FastSymbol(b$, "(") Then
                        W$ = ppppAny.item(v).lasthere + "." + ppppAny.item(v).GroupName + "("
                        If neoGetArray(bstack, W$, ppppAny, , True, , True) Then
                            GoTo againarray
                        End If
                    Else
                        Mid$(b$, 1, 1) = ChrW(0)
                    End If
                End If
                ExecuteVar = 9
            End If
            Exit Function
            End If
        End If
     End If
     WrongFatArrow
     GoTo err000
ElseIf Not IsExp(bstack, b$, p) Then
    If IsStrExp(bstack, b$, sw$) Then
        p = vbNullString
        SwapString2Variant sw$, p
    Else
        If LastErNum = -2 Then
            Execute bstack, b$, True
        Else
            MissNumExpr
        End If
        GoTo err000
    End If
End If
again12569:
     If Not bstack.lastobj Is Nothing Then
         If Typename(bstack.lastobj) = mGroup Then
            If bstack.lastobj.IamApointer Then
                Set ppppAny.item(v) = bstack.lastobj
                Set bstack.lastobj = Nothing
                Set bstack.lastpointer = Nothing
                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
            End If
        End If
        Set myobject = ppppAny.GroupRef
        If ppppAny.IhaveClass Then
            Set ppppAny.item(v) = bstack.lastobj
            Set ppppAny.item(v).LinkRef = myobject
           ' Set pppp = ppppAny
            With ppppAny.item(v)
                .HasStrValue = myobject.HasStrValue
                .HasValue = myobject.HasValue
                .HasSet = myobject.HasSet
                .HasParameters = myobject.HasParameters
                .HasParametersSet = myobject.HasParametersSet
                Set .SuperClassList = myobject.SuperClassList
                If .Events Is Nothing Then
                    Set .Events = myobject.Events
                End If
                .highpriorityoper = myobject.highpriorityoper
                .HasUnary = myobject.HasUnary
                .HasRemove = myobject.HasRemove
                Set bstack.lastobj = Nothing
                .ToDelete = False
            End With
        Else
            If Typename(bstack.lastobj) = mHdlr Then
                Set ppppAny.item(v) = bstack.lastobj
            Else
                If Not bstack.lastobj Is Nothing Then
                    If TypeOf bstack.lastobj Is iBoxArray Then
                        If bstack.lastobj.Arr Then
                            Set ppppAny.item(v) = CopyArray(bstack.lastobj)
                        Else
                            Set ppppAny.item(v) = bstack.lastobj
                        End If
                    Else
                        If TypeOf bstack.lastobj Is Group Then
                            
                            Set ppppAny.item(v) = Nothing
                            bstack.lastobj.ToDelete = False
                            Set ppppAny.item(v) = bstack.lastobj
                            If Not myobject Is Nothing Then
                                If ppppAny.item(v).LinkRef Is Nothing Then
                                    Set pppp.item(v).LinkRef = myobject
                                End If
                            End If
                            Set bstack.lastobj = Nothing
                            ' CopyBigInteger
                        ElseIf TypeOf bstack.lastobj Is BigInteger Then
                            Set p = bstack.lastobj
                            Set bstack.lastobj = Nothing
                            Set ppppAny.item(v) = CopyBigInteger(p)
                        Else
                            Set ppppAny.item(v) = bstack.lastobj
                        End If
                    End If
                Else
                    Set ppppAny.item(v) = bstack.lastobj
                    If TypeOf bstack.lastobj Is Group Then Set ppppAny.item(v).LinkRef = myobject
                End If
            End If
        End If
        
        Set bstack.lastobj = Nothing
     Else
        If ppppAny.Arr Then
            If ppppAny.ItemType(v) = mGroup Then
here12500:
                If ppppAny.item(v).IamApointer Then
                    If ppppAny.item(v).link.HasSet Then GoTo here65654
                End If
                If ppppAny.item(v).HasSet Then
here65654:
                    bstack.soros.PushVal p
                    Exec1 = SpeedGroup(bstack, ppppAny, "@READ", W$, b$, v)
                Else
                    If p = 0# Then
                        pppp.item(v) = 0&  ' release pointer
                    Else
                        GroupCantSetValue
                    End If
                End If
            Else
                If TypeOf ppppAny Is mArray Then
                    Set pppp = ppppAny
                    pppp.item(v) = p
                Else
                    ppppAny.item(v) = p
                End If
                If LastErNum1 Then GoTo err000
            End If
        ElseIf Typename(ppppAny.GroupRef) = mProp Then
            Set myProp = ppppAny.GroupRef
            myProp.PushIndexes idx
            myProp.Value = p
            Set myProp = Nothing
        ElseIf ppppAny.ItemType(v) = mGroup Then
            If ppppAny.item(v).HasSet Then
                bstack.soros.PushVal p
                Exec1 = SpeedGroup(bstack, ppppAny, "@READ", W$, b$, v)
            Else
                GroupCantSetValue
            End If
        ElseIf ppppAny.ItemType(v) = mProp Then
            Set myProp = ppppAny.itemObject(v)
            myProp.PushIndexes idx
            myProp.Value = p
            Set myProp = Nothing
        ElseIf Not pppp.Arr Then
            NoAssignThere
        End If
    End If
    If TypeOf ppppAny Is mArray Then
    Set pppp = ppppAny
    Do While FastSymbol(b$, ",")
        If pppp.UpperMonoLimit > v Then
            v = v + 1
            If Not IsExp(bstack, b$, p) Then GoTo err000
            If Not bstack.lastobj Is Nothing Then
            Set myobject = pppp.GroupRef
            If pppp.IhaveClass Then
                Set pppp.item(v) = bstack.lastobj
                    With pppp.item(v)
                    .HasStrValue = myobject.HasStrValue
                    .HasValue = myobject.HasValue
                    .HasSet = myobject.HasSet
                    .HasParameters = myobject.HasParameters
                    .HasParametersSet = myobject.HasParametersSet
                    Set .SuperClassList = myobject.SuperClassList
                    Set .Events = myobject.Events
                    .highpriorityoper = myobject.highpriorityoper
                    .HasUnary = myobject.HasUnary
                    .HasRemove = myobject.HasRemove
                End With
            Else
                Set pppp.item(v) = bstack.lastobj
            End If
            If Not myobject Is Nothing Then
                If pppp.ItemType(v) = mGroup Then Set pppp.item(v).LinkRef = myobject
            End If
            Set bstack.lastobj = Nothing
        Else
            pppp.item(v) = p
        End If
    Else
        Exit Do
    End If
Loop
End If
Else
    If LastErNum <> 0 Then GoTo err000
    bstack.tmpstr = ss$
    ExecuteVar = 2  ' GoTo autogosub
    Exit Function
End If
Exit Function
Case6:
If AscW(W$) = 46 Then
               If Not expanddot(bstack, W$) Then
               ManyDots
                GoTo err000
               End If
End If
If VarStat Or NewStat Or noVarStat Then
 If strfunid.Find(W$, I) Then
    If I > 0 Then strfunid.ItemCreator W$, -I
      End If
MakeArray bstack, W$, 6, b$, pppp, NewStat, VarStat
 'If Not lookone(b$, ",") Then b$ = " :" + b$
        sss = Len(b$): ExecuteVar = 4: Exit Function
End If
If neoGetArray(bstack, W$, ppppAny, , , , True) Then
    If Not ppppAny.Arr Then
If Not NeoGetArrayItem(ppppAny, bstack, W$, v, b$, , , , , idx) Then GoTo err000
GoTo there12567
ElseIf FastSymbol(b$, ")") Then
    'need to found an expression - HEREHERE
        If FastSymbol(b$, "=") Then
            If IsStrExp(bstack, b$, W$) Then
                If Not bstack.lastobj Is Nothing Then
                    If TypeOf bstack.lastobj Is mHandler Then
                        Set usehandler = bstack.lastobj
                        
                        If TypeOf usehandler.objref Is mArray Then
                        If TypeOf ppppAny Is mArray Then
                            Set pppp = ppppAny
                            Set pppp1 = usehandler.objref
                            pppp.Final = False
                            pppp1.CopyArray pppp
                            Set pppp1 = Nothing
                            Else
                            GoTo NotArray1
                            End If
                        Else
                        GoTo NotArray1
                        End If
                    ElseIf TypeOf bstack.lastobj Is mArray Then
                        If TypeOf ppppAny Is mArray Then
                            Set pppp = ppppAny
                            Set pppp1 = bstack.lastobj
                            pppp1.CopyArray pppp
                            pppp.Final = False
                            Set pppp1 = Nothing
                        Else
                            GoTo NotArray1
                        End If
                    Else
                        GoTo NotArray1
                    End If
                    Set bstack.lastobj = Nothing
                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                End If
            ElseIf IsExp(bstack, b$, p) Then
        If Not bstack.lastobj Is Nothing Then
        If TypeOf bstack.lastobj Is mHandler Then
            Set usehandler = bstack.lastobj
         If usehandler.indirect >= 0 Then
         Set bstack.lastobj = var(usehandler.indirect)
         Else
         Set bstack.lastobj = usehandler.objref
         End If
         Set usehandler = Nothing
        If TypeOf bstack.lastobj Is mArray Then
            If TypeOf ppppAny Is mArray Then
                Set pppp = ppppAny
                bstack.lastobj.CopyArray pppp
                pppp.Final = False
            Else
                GoTo NotArray1
            End If
         Else
         GoTo NotArray1
         End If
        Else
            If TypeOf ppppAny Is mArray Then
                Set pppp = ppppAny
                bstack.lastobj.CopyArray pppp
            Else
                GoTo NotArray1
            End If
        End If
            Set bstack.lastobj = Nothing
            If extreme Then GoTo NewCheck2 Else GoTo NewCheck
        Else
            Set pppp1 = New mArray: pppp1.PushDim (1): pppp1.PushEnd
            pppp1.SerialItem 0, 2, 9
            pppp1.Arr = True
            If bstack.lastobj Is Nothing Then
                pppp1.item(0) = vbNullString
            Else
                Set pppp1.item(0) = bstack.lastobj
                Set bstack.lastobj = Nothing
            End If
            pppp1.CopyArray pppp
            If extreme Then GoTo NewCheck2 Else GoTo NewCheck
        End If
    Else
                SyntaxError
            End If
            GoTo err000
        End If
  
        
        End If
If v = -2 Then GoTo checkpar
againstrarr:
If Not NeoGetArrayItem(ppppAny, bstack, W$, v, b$) Then GoTo err000
'On Error Resume Next
' WHY BEFORE WAS : If pppp.itemtype(v) = myArray And Not pppp.Arr Then
there12567:
If ppppAny.Arr Then
If ppppAny.ItemType(v) = myArray Then
If FastSymbol(b$, "(") Then
Set ppppAny = ppppAny.item(v)
GoTo againstrarr
End If
End If
End If
If v = -2 Then
checkpar:
W$ = Left$(W$, Len(W$) - 1)
Set myobject = bstack.soros
Set bstack.Sorosref = New mStiva
PushParamGeneral bstack, b$
'PushParamStraight bstack, b$
    If Not FastSymbol(b$, ")", True) Then
            Set bstack.Sorosref = myobject
            GoTo err000
    End If
    If FastSymbol(b$, "=") Then
        If IsStrExp(bstack, b$, ss$) Then
       ' bstack.soros.MakeTopItemBack bstack.soros.Total - v + 1
                                If bstack.lastobj Is Nothing Then
                                    bstack.soros.DataStr ss$
                                Else
                                    If TypeOf bstack.lastobj Is VarItem Then
                                    bstack.soros.DataOptional
                                    Else
                                    bstack.soros.DataObj bstack.lastobj
                                    End If
                                    Set bstack.lastobj = Nothing
                                End If
                                
                                NeoCall2 bstack, Left$(W$, Len(W$) - 1) + "." + ChrW(&H1FFF) + ":=()", ok
                                Set bstack.Sorosref = myobject
                                Set myobject = Nothing
                                If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                                
                                
                                End If
                                
                                
                                End If
                                Set bstack.Sorosref = myobject
                                Set myobject = Nothing
                                SyntaxError
                                GoTo err000

ElseIf Not FastSymbol(b$, "=") Then
    If Not TypeOf ppppAny Is mArray Then
                        GoTo WrongObj
    End If
    Set pppp = ppppAny
    If pppp.Arr Then
        If FastSymbol(b$, ":=", , 2) Then GoTo contassignhere
    End If
    
  If IsOperator(b$, "+=", 2) Then
        If pppp.IsStringItem(v) Then
            If Not IsStrExp(bstack, b$, ss$, False) Then GoTo err000
            If bstack.lastobj Is Nothing Then
                pppp.ItemStr(v) = pppp.item(v) + ss$
            Else
                NeedString
                GoTo err000
            End If
        Else
            FoundNoStringItem
            GoTo err000
        End If
    ElseIf IsOperator(b$, "(") Then
        If pppp.ItemType(v) = myArray Then
            Set pppp = pppp.item(v)
            Set ppppAny = pppp
            GoTo againstrarr
        Else ' only group here
   
            Set myobject = bstack.soros
            Set bstack.Sorosref = New mStiva
            PushParamStraight bstack, b$
            If Not FastSymbol(b$, ")", True) Then
                    Set bstack.Sorosref = myobject
                    GoTo err000
            End If
            If Not FastSymbol(b$, "=", True) Then
                sss = 0
                ExecuteVar = 3: Exit Function
            End If
            If Not IsStrExp(bstack, b$, ss$) Then
                If LastErNum = -2 Then
                Execute bstack, b$, True
                Else
                MissNumExpr
                End If
                GoTo err000
            End If
   '
            If bstack.lastobj Is Nothing Then
                bstack.soros.DataStr ss$
            Else
                bstack.soros.DataObj bstack.lastobj
                Set bstack.lastobj = Nothing
            End If
            Set ppppAny = pppp
            Exec1 = SpeedGroup(bstack, ppppAny, "@READ2", "", b$, v)
            If extreme Then GoTo NewCheck2 Else GoTo NewCheck
        End If
   ElseIf FastSymbol(b$, "->", , 2) Then
    If Not GetPointer(bstack, b$) Then GoTo err000
    With pppp
    If Typename(bstack.lastobj) = mGroup Then
        If Typename(.item(v)) <> mGroup Then
        
        If bstack.lastpointer Is Nothing Then
                Set .item(v) = bstack.lastobj
        Else
                Set .item(v) = bstack.lastpointer
        End If
        
        Else
        If .item(v).IamApointer Then
            If bstack.lastpointer Is Nothing Then
                ExpectedPointer
                   GoTo err000
            Else
                Set .item(v) = bstack.lastpointer
            End If
        End If
        End If
        
        Set bstack.lastobj = Nothing
        Set bstack.lastpointer = Nothing
          If extreme Then GoTo NewCheck2 Else GoTo NewCheck
    End If
    End With
   GoTo err000
   ElseIf FastSymbol(b$, "+=", , 2) Then
   If IsStrExp(bstack, b$, ss$) Then
    
          CheckVar pppp.item(v), ss$, True
    
        
  Exit Function
  Else
  GoTo err000
  End If
    Else
        GoTo err000
    End If
Else
    If IsExp(bstack, b$, p, nostring:=False) Then
        Assign ss$, p
        GoTo jmp1112
    ElseIf Not IsStrExp(bstack, b$, ss$, False) Then
        GoTo err000
    End If
jmp1112:
    If Not MyIsObject(ppppAny.item(v)) Then
    
    If TypeOf ppppAny Is mArray Then
        If ppppAny.Arr Then
        Set pppp = ppppAny
        If pppp.count = 0 Then
            pppp.GroupRef.Value = ss$
        ElseIf bstack.lastobj Is Nothing Then
            pppp.ItemStr(v) = ss$
        Else
            If Typename(bstack.lastobj) = myArray Then
                If bstack.lastobj.Arr Then
                    Set pppp.item(v) = CopyArray(bstack.lastobj)
                Else
                    Set pppp.item(v) = bstack.lastobj.GroupRef
                End If
            Else
                Set pppp.item(v) = bstack.lastobj
            End If
            Set bstack.lastobj = Nothing
        End If
    
        Else
        If v < 0 And v <> -2 Then
        NoAssignThere
        Else
        Set myProp = pppp.GroupRef
        myProp.PushIndexes idx
        myProp.Value = ss$
        End If
        End If
        Else
        SyntaxError
        End If
    ElseIf ppppAny.ItemType(v) = mGroup Then
        If ppppAny.item(v).HasSet Then
        bstack.soros.PushStr ss$
            Exec1 = SpeedGroup(bstack, ppppAny, "@READ", W$, b$, v)
        Else
        GroupCantSetValue
        End If
    ElseIf ppppAny.ItemType(v) = mProp Then
        Set myProp = ppppAny.itemObject(v)
        With myProp
            .PushIndexes idx
            .ValueStr = ss$
        End With
        Set myProp = Nothing
    Else
        CheckVar ppppAny.item(v), ss$
    End If
    If TypeOf ppppAny Is mArray Then
        Set pppp = ppppAny
        Do While FastSymbol(b$, ",")
        If pppp.UpperMonoLimit > v Then
        v = v + 1
        If Not IsStrExp(bstack, b$, ss$) Then MissStringExpr: GoTo err000
        
        If Not MyIsObject(pppp.item(v)) Then
          pppp.item(v) = ss$
          Else
                CheckVar pppp.item(v), ss$
        End If
        Else
        Exit Do
        End If
        Loop
        End If
End If
Else
GoTo err000
End If
Exit Function
Case8:
p = CheckThis(bstack, W$, b$, v, Lang)
If AscW(W$) = 46 Then
               If Not expanddot(bstack, W$) Then
                ManyDots
                
               End If
End If
FastSymbol1 b$, "["
Case8new:
st11234:
If Not IsExp(bstack, b$, p, , flatobject:=True, nostring:=True) Then
    MissNumExpr
    GoTo err000
End If
If Not FastSymbol(b$, "]") Then
    SyntaxError
    GoTo err000
End If
st112233:

If Left$(b$, 1) = "[" Then
    I = Abs(CLng(p))
    Mid$(b$, 1, 1) = " "
    If IsExp(bstack, b$, p, , flatobject:=True, nostring:=True) Then
        ok = True
        If Not FastSymbol(b$, "]") Then
            SyntaxError
        Else
            If Left$(b$, 1) = "[" Then
                If GetVar(bstack, W$, v, True) Then
                If Not Typename(var(v)) = "RefArray" Then
                        WrongType
                      GoTo err000
                End If
                Set Ar = var(v)
st12123434:
                If Not Ar.MarkTwoDimension Then
st9994:
                    If Typename$(Ar.Value(0, I)) = "RefArray" Then
                        Set Ar = Ar.Value(0, I)
st3535435:
                        If Ar.MarkTwoDimension Then
                            'FastSymbol1 b$, "["
                            Mid$(b$, 1, 1) = " "
                            I = p
st38383:
                            If IsExp(bstack, b$, p, , flatobject:=True, nostring:=True) Then
                                If Not FastSymbol(b$, "]") Then
                                    SyntaxError
                                Else
                                    GoTo entry100101
                                End If
                            End If
                        ElseIf Typename$(Ar(0, p)) = "RefArray" Then
                            
                            Set Ar = Ar(0, p)
st29939:
                            Mid$(b$, 1, 1) = " "
                            If IsExp(bstack, b$, p, , flatobject:=True, nostring:=True) Then
                                If Not FastSymbol(b$, "]") Then
                                    SyntaxError
                                Else
                                   If Ar.MarkTwoDimension Then
                                    GoTo st3535435
                                   Else
                                   If Left$(b$, 1) = "[" Then
                                    I = p
                                    GoTo st9994
                                    Else
                                        I = 0
                                        GoTo entry100101
                                    End If
                                   End If
                                End If
                            End If
                        Else
                        I = 0
                        GoTo st29939
                        End If
                    Else
                        I = 0
                        GoTo entry100101
                    End If
                End If
                    WrongType
                    GoTo err000
                Else
                    UnknownVariable W$
                End If
            End If
            GoTo entry100101
        End If
    Else
        MissNumExpr
    End If
Else
    ok = False
    I = 0

    
entry100101:
       If jumpto = 9 Then
            ww = -100
            GoTo entry00101
       Else
        Select Case Left$(b$, 1)
        Case "="
            Mid(b$, 1, 1) = " "
            If Left$(b$, 2) = " >" Then
                GoTo forwidearrow
            End If
            ww = 8
        Case "~"
            ww = 0
            Mid$(b$, 1, 1) = " "
            GoTo entry00101
        Case "|"
        ww = 9
        Mid$(b$, 1, 1) = " "
        GoTo entry00101
       
        Case Else
            Select Case Left$(b$, 2)
            Case "++": ww = 1: Mid$(b$, 1, 2) = "  ": GoTo entry00101
            Case "--": ww = 2: Mid$(b$, 1, 2) = "  ": GoTo entry00101
            Case "-!": ww = 3: Mid$(b$, 1, 2) = "  ": GoTo entry00101
            Case "+=": ww = 4
            Case "-=": ww = 5
            Case "*=": ww = 6
            Case "/=": ww = 7
            Case "<="
                Mid$(b$, 1, 2) = "  "
                If Not IsExp(bstack, b$, sp, , flatobject:=True) Then
                    If IsStrExp(bstack, b$, sw$, False) Then
                        sp = sw$
                    Else
                        GoTo cont00100203
                    End If
                End If
                If GetVar(bstack, W$, v, True) Then
                    ww = 8
                    GoTo entry00121
                Else
                    UnknownVariable W$
                End If
            Case Else
                WrongOperator
                Exit Function
            End Select
            Mid$(b$, 1, 2) = "  "
        End Select
    End If
    If ww > 3 Then
        If False Then
forwidearrow:
            If varhash.Find2(here$ + "." + myUcase(W$), v) Then
                
            ElseIf GetVar(bstack, W$, v, True) Then
                
            Else
                GoTo cont00100203
            End If
            If Typename(var(v)) = "RefArray" Then
                Set Ar = var(v)
entry00022:
                If Ar.MarkTwoDimension And Not ok Then
                    GoTo entry00123
                End If
                p = Abs(Int(p))
                sp = I
                If Ar.vtType(0) = vbObject And ok Then
                    If Ar.count > 1 Then
                    
                    ElseIf Ar(0, sp) Is Nothing Then
                        GoTo nRefArray
                    ElseIf TypeOf Ar(0, sp) Is refArray Then
                        Set Ar = Ar(0, sp)
                        sp = 0
                        ok = False
                        GoTo entry00022
                    Else
                        WrongObject
                        GoTo nRefArray
                    End If
                ElseIf Ar.count < I Then
                        OutOfLimit
                        GoTo nRefArray
                ElseIf Ar.count(sp) = 0 Then
                        OutOfLimit
                        GoTo nRefArray
                ElseIf Ar.count(sp) <= p Then
                        OutOfLimit
                        GoTo nRefArray
                End If
                
                If myVarType(Ar(sp, p), vbObject) Then
                
                Set p = Ar(sp, p)
                If p Is Nothing Then
                    NoOperatorForThatObject "=>"
                ElseIf TypeOf p Is Group Then
                    If p.link Is Nothing Then
                        NoOperatorForThatObject "=>"
                        
                    ElseIf p.link.IamFloatGroup Then
                        ExecuteVar = 10
        
                        Mid$(b$, 1, 2) = ChrW(7) + ChrW(3)
                        Set bstack.lastpointer = p
                        Exit Function
                    Else
                        ExecuteVar = 9
                        Mid$(b$, 1, 1) = Chr$(0) + Chr$(0) ' cause we have two chars
                        Set bstack.lastpointer = p
                        Exit Function
                    End If
                Else
                    NoOperatorForThatObject "=>"
                End If
                Else
                    WrongType
                End If
    
                
            End If
            GoTo cont00100203
        Else
            If IsExp(bstack, b$, sp, nostring:=False) Then
            
entry00101:
                If Not Ar Is Nothing Then GoTo entry00122
                If varhash.Find2(here$ + "." + myUcase(W$), v, useType) Then
entry00121:
                    If Typename(var(v)) = "RefArray" Then
                        Set Ar = var(v)
entry00122:
                        If Ar.MarkTwoDimension And Not ok Then
                        If Ar.vtType(0) = vbVariant Then
                        If Not bstack.lastobj Is Nothing Then
                        If TypeOf bstack.lastobj Is refArray Then
                        If Not bstack.lastobj Is Ar Then
                            If ww = 8 Then
                            
                             Ar(p) = CVar(bstack.lastobj)
                                Set bstack.lastobj = Nothing
                                        While FastSymbol(b$, ",")
                                        If Not IsExp(bstack, b$, sp) Then
                                            WrongType
                                            GoTo err000
                                        End If
                                        p = p + 1
                                        Ar(p) = CVar(bstack.lastobj)
                                        Set bstack.lastobj = Nothing
                                        Wend
                            GoTo NewCheck
                            End If
                        End If
                        End If
                        End If
                        End If
                        
entry00123:
                           MissingIndexMore
                        Else
                        
                            p = Abs(Int(p))
                            
                            If Ar.IsInnerRefArray(I, Ar) Then
                                I = 0
                                ok = False
                                GoTo entry00122
                            End If
                            If (Ar.vtType(0) = vbObject) And ok Then  ' Or ar.vtType(0) = vbVariant
                                If Ar.count > 1 Then
                                                                
                                If Ar.count(CVar(I)) = 0 Then
                                    Set sp = bstack.lastobj
                                    Set bstack.lastobj = Nothing
                                    GoTo count0
                                End If
                                
                                GoTo takeitnow
                                
                                ElseIf Ar(0, CVar(I)) Is Nothing Then
                                ' error for [ ]
                                If Not bstack.lastobj Is Nothing Then
                                If Ar.count(CVar(I)) = 0 Then
                                    Set sp = bstack.lastobj
                                    Set bstack.lastobj = Nothing
                                    GoTo count0
                                End If
                                GoTo takeitnow
                                Else
nRefArray:
                                    ExpRefArray I
                                    GoTo cont00100203
                                End If
                                ElseIf TypeOf Ar(0, CVar(I)) Is refArray Then
                                    Set Ar = Ar(0, CVar(I))
                                    I = 0
                                    ok = False
                                    GoTo entry00122
                                Else
                                    WrongObject
                                    GoTo nRefArray
                                End If
                            ElseIf Ar.count(CVar(I)) = 0 Then
count0:
                                
                                Ar.DefArrayAt I, Ar.vtType(0), CLng(p)
                                ' if ar.vtType(0)=vbstring  ......... check for string
                                Select Case ww
                                Case 0
                                Ar(CVar(I), p) = True
                                Case 8, 4, 18, 14
                                Ar(CVar(I), p) = sp
                                GoTo st9993993
                                
                                Case 5
                                    If myVarType(sp, vbString) Then
                                        Ar(CVar(I), p) = sp
                                    Else
                                        Ar(CVar(I), p) = -sp
                                    End If
                                End Select
                            Else

                                If Not bstack.lastobj Is Nothing Then
takeitnow:
                                    If ww <> 8 Then
                                        If ww = -100 Then
                                        If bstack.IsObjectRef(myobject) Then
                                            Ar(CVar(I), p) = CVar(myobject)
                                            GoTo NewCheck2
                                        End If
                                        Else
                                        WrongOperator
                                        End If
                                    Else
                                        If bstack.lastobj Is Nothing Then
                                        If sp <> 0 Then
                                        MissType
                                        GoTo err000
                                        End If
                                        ElseIf TypeOf bstack.lastobj Is Group Then
                                                Set sp = bstack.lastobj
                                                If Not sp.IamApointer Then
                                                Set bstack.lastobj = Nothing
                                                Set bstack.lastpointer = Nothing
                                                MakeGroupPointer bstack, sp
                                                sp = 0
                                                Else
                                                Set sp = Nothing
                                                End If
                                        ElseIf TypeOf bstack.lastobj Is BigInteger Then
                                            Set sp = bstack.lastobj
                                            Set bstack.lastobj = Nothing
                                            Ar(CVar(I), p) = CopyBigInteger(sp)
                                            Set sp = Nothing
                                            While FastSymbol(b$, ",")
                                            If Not IsExp(bstack, b$, sp) Then
                                                WrongType
                                                GoTo err000
                                            End If
                                            p = p + 1
                                            If Not TypeOf bstack.lastobj Is BigInteger Then
                                                WrongType
                                                GoTo err000
                                            End If
                                            Set sp = bstack.lastobj
                                            Set bstack.lastobj = Nothing
                                            Ar(CVar(I), p) = CopyBigInteger(sp)
                                            Set sp = Nothing
                                            Wend
                                            
                                            GoTo cont11100329
                                        End If
                                    ' check this
                                        Ar(CVar(I), p) = CVar(bstack.lastobj)
                                        Set bstack.lastobj = Nothing
                                        While FastSymbol(b$, ",")
                                        If Not IsExp(bstack, b$, sp) Then
                                            WrongType
                                            GoTo err000
                                        End If
                                        p = p + 1
                                        Ar(CVar(I), p) = CVar(bstack.lastobj)
                                        Set bstack.lastobj = Nothing
                                        Wend
                                        
                                    End If
cont11100329:
                                Else
                                    Select Case ww
                                    Case -100
                                    
                                    If Ar.vtType(CVar(I), p) = 36 Then
                                    If bstack.IsAny(sp) Then
                                    If FastSymbol(b$, "|") Then
                                    
                                        ww = IsLabel(bstack, b$, ss$)
                                        If ww = 1 Then
                                        Ar.PlaceValue2UDT I, p, ss$, sp
                                        GoTo thatsall
                                        ElseIf ww = 5 Then
                                        ww = p
                                        If IsExp(bstack, b$, p, , flatobject:=True) Then
                                        If FastSymbol(b$, ")") Then
                                            Ar.PlaceValue2UDT I, ww, ss$, sp, p
                                            GoTo thatsall
                                            End If
                                        End If
                                        End If
                                        
                                    End If
                                    End If
                                    SyntaxError
                                        GoTo err000
                                    End If
                                    If bstack.IsNumber(sp) Then
                                    Ar(CVar(I), p) = sp
                                    ElseIf bstack.IsString(sw$) Then
                                    Ar(CVar(I), p) = CVar(sw$)
                                    ElseIf bstack.IsObjectRef(myobject) Then
                                    Ar(CVar(I), p) = CVar(myobject)
                                    End If

                                    Case 0:
                                    Ar(CVar(I), p) = Ar(CVar(I), p) = 0
                                    Case 1: Ar(CVar(I), p) = Ar(CVar(I), p) + 1
                                    Case 2: Ar(CVar(I), p) = Ar(CVar(I), p) - 1
                                    Case 3: Ar(CVar(I), p) = -Ar(CVar(I), p)
                                    Case 4: Ar(CVar(I), p) = Ar(CVar(I), p) + sp
                                    Case 5: Ar(CVar(I), p) = Ar(CVar(I), p) - sp
                                    Case 6: Ar(CVar(I), p) = Ar(CVar(I), p) * sp
                                    Case 7: Ar(CVar(I), p) = Ar(CVar(I), p) / sp
                                    Case 8
                                    
                                    Ar(CVar(I), p) = sp
st9993993:
                                    
                                    While FastSymbol(b$, ",")
                                    If Not IsExp(bstack, b$, sp, , flatobject:=True) Then
                                        If IsStrExp(bstack, b$, sw$, False) Then
                                            sp = ""
                                            SwapString2Variant sw$, sp
                                        Else
                                            SyntaxError
                                            GoTo err000
                                        End If
                                    End If
                                    p = p + 1
                                    Ar(CVar(I), p) = sp
                                    Wend
                                    Case 9:
                                        ww = IsLabel(bstack, b$, ss$)
                                        If ww = 1 Then
                                        
                                        If FastSymbol(b$, "=") Then
                                        
                                        If Not IsExp(bstack, b$, sp) Then
                                        If IsStrExp(bstack, b$, sw$, False) Then
                                            sp = ""
                                            SwapString2Variant sw$, sp
                                        Else
                                            SyntaxError
                                            GoTo err000
                                        End If
                                        End If
                                        Ar.PlaceValue2UDT I, p, ss$, sp
                                        Else
                                        
                                        v = Ar(CVar(I), p)
                                        If IsExp(bstack, b$, sp) Then
                                            If Not readvarv(v, ss$, sp) Then
                                                ExecuteVar = 0
                                                Exit Function
                                            End If
                                            Ar(CVar(I), p) = v
                                        End If
                                   
                                        '
                                        End If
                                        
                                        ElseIf ww = 5 Then
                                        
                                        If IsExp(bstack, b$, sp) Then
                                            ww = CLng(sp)
                                            If FastSymbol(b$, ")") Then
                                                If FastSymbol(b$, "=") Then
                                                    If Not IsExp(bstack, b$, sp) Then
                                                        If IsStrExp(bstack, b$, sw$, False) Then
                                                            sp = ""
                                                            SwapString2Variant sw$, sp
                                                        Else
                                                            SyntaxError
                                                            GoTo err000
                                                        End If
                                                    End If
                                                Ar.PlaceValue2UDT I, p, ss$, sp, CVar(ww)
                                                End If
                                            End If
                                        End If
                                    End If
                                    Case 14: Ar(CVar(I), p) = Ar(CVar(I), p) + sp
                                    Case 18
                                    Ar(CVar(I), p) = sp
                                    While FastSymbol(b$, ",")
                                    If Not IsExp(bstack, b$, sp, , flatobject:=True) Then
                                        If IsStrExp(bstack, b$, sw$, False) Then
                                            sp = sw$
                                        Else
                                            SyntaxError
                                            GoTo err000
                                        End If
                                    End If
                                    p = p + 1
                                    Ar(CVar(I), p) = sp
                                    Wend
                                    End Select
                                End If
                            End If
                            Select Case Ar.AssignError
                            Case 6
                               
                                OverflowValue VarType(Ar(I, p))
                                GoTo err000
                            Case 0
                            Case Else
                                If ww = 14 Or ww = 18 Then
                                    NeedString
                                Else
                                    MissType
                                End If
                                GoTo err000
                            End Select
thatsall:
                            Set bstack.lastobj = Nothing
                            Set bstack.lastpointer = Nothing
                            GoTo NewCheck
                        End If
                    ElseIf Typename(var(v)) = mHdlr Then
                        ' ww=8 p= index   a[p]=
                        ' ww=9 p= index   a[p]|=
                        ' contstruct11 for ww=9
                        If ww = 9 Then Mid$(b$, 1, 1) = "|"
                        Set usehandler = var(v)
                        
                        If TakeOffset(bstack, usehandler, b$, sp, p, ww - 8) Then
                       GoTo NewCheck
                        End If
                       GoTo err000
                    Else
                      GoTo WrongObj
                    End If
                ElseIf ww <> 8 Or here$ = vbNullString Then
                    If GetVar(bstack, W$, v, True) Then
                        GoTo entry00121
                    Else
                        UnknownVariable W$
                    End If
                End If
            ElseIf IsStrExp(bstack, b$, sw$, False) Then
                If ww = 8 Or ww = 4 Then
                    sp = sw$
                    ww = ww + 10
                    GoTo entry00101
                Else
                    WrongOperator
                End If
            End If
        End If
    End If
End If
cont00100203:
GoTo err000
Exit Function
Case7:
If AscW(W$) = 46 Then
               If Not expanddot(bstack, W$) Then
               ManyDots
                
               End If
End If
If VarStat Or NewStat Or noVarStat Then
MakeArray bstack, W$, 7, b$, pppp, NewStat, VarStat
 'If Not MaybeIsSymbol(b$, ",") Then b$ = " :" + b$
        sss = Len(b$): ExecuteVar = 4: Exit Function
End If
If neoGetArray(bstack, W$, pppp) Then
    If FastSymbol(b$, ")") Then
    'need to found an expression
        If FastSymbol(b$, "=") Then
            If IsExp(bstack, b$, p) Then
                If Not bstack.lastobj Is Nothing Then
                    bstack.lastobj.CopyArray pppp
                    Set bstack.lastobj = Nothing
                    If extreme Then GoTo NewCheck2 Else GoTo NewCheck
                End If
            Else
                SyntaxError
            End If
            GoTo err000
        End If
        End If
againintarr:
If Not NeoGetArrayItem(pppp, bstack, W$, v, b$) Then GoTo err000
'On Error Resume Next
If pppp.ItemType(v) = myArray And pppp.Arr Then
If FastSymbol(b$, "(") Then
Set pppp = pppp.item(v)
GoTo againintarr
End If
End If
If lookTwoSame(b$, "/") Then
GoTo err000
ElseIf MaybeIsSymbol(b$, "+-*/~|") Then
On Error Resume Next
With pppp
If IsOperator0(b$, "++", 2) Then
.item(v) = .itemnumeric(v) + 1
ElseIf IsOperator0(b$, "--", 2) Then
.item(v) = .itemnumeric(v) - 1
ElseIf FastSymbol(b$, "+=", , 2) Then
If Not IsExp(bstack, b$, p) Then GoTo err000
.item(v) = .itemnumeric(v) + MyRound(p)
ElseIf FastSymbol(b$, "-=", , 2) Then
If Not IsExp(bstack, b$, p) Then GoTo err000
.item(v) = .itemnumeric(v) - MyRound(p)
ElseIf FastSymbol(b$, "*=", , 2) Then
If Not IsExp(bstack, b$, p) Then GoTo err000
.item(v) = MyRound(.itemnumeric(v) * MyRound(p))
ElseIf FastSymbol(b$, "/=", , 2) Then
If Not IsExp(bstack, b$, p) Then GoTo err000
If MyRound(p) = 0 Then
 DevZero
 Else
 .item(v) = MyRound(.itemnumeric(v) / MyRound(p))
End If
ElseIf IsOperator0(b$, "-!", 2) Then
.item(v) = -.itemnumeric(v)
ElseIf IsOperator0(b$, "~") Then
        Select Case VarType(.itemnumeric(v))
            Case vbBoolean
                .item(v) = Not CBool(.itemnumeric(v))
            Case vbInteger
                .item(v) = CInt(Not CBool(.itemnumeric(v)))
            Case vbLong
                .item(v) = CLng(Not CBool(.itemnumeric(v)))
            Case vbCurrency
                .item(v) = CCur(Not CBool(.itemnumeric(v)))
            Case vbDecimal
                .item(v) = CDec(Not CBool(.itemnumeric(v)))
            Case Else
                .item(v) = CDbl(Not CBool(.itemnumeric(v)))
        End Select
ElseIf IsOperator0(b$, "|") Then
    If FastPureLabel(b$, ss$, , True) Then
        If Mid$(b$, 1, 1) = "#" Then ss$ = ss$ + "#": Mid$(b$, 1, 1) = " "
        If IsExp(bstack, b$, p) Then
            If Int(p) = 0 Then
                DevZero
                GoTo err000
            End If
            SwapVariant sp, .item(v)
            If Not readvarv(sp, ss$, p) Then
                WrongOperator
                GoTo err000
            End If
            .item(v) = CInt(sp)
        Else
            GoTo noexpression
        End If
    Else
      WrongOperator
    End If
Else
      WrongOperator
End If
End With
On Error GoTo 0
If extreme Then GoTo NewCheck2 Else GoTo NewCheck
End If
If Not FastSymbol(b$, "=") Then
  If FastSymbol(b$, ":=", , 2) Then
  
    GoTo contassignhere
End If
GoTo err000
End If
    If Not IsExp(bstack, b$, p) Then MissNumExpr: GoTo err000
    If Not bstack.lastobj Is Nothing Then
        If TypeOf bstack.lastobj Is mArray Then
            If bstack.lastobj.Arr Then
                Set pppp.item(v) = CopyArray(bstack.lastobj)
            Else
                Set pppp.item(v) = bstack.lastobj
                If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
            End If
        Else
            Set pppp.item(v) = bstack.lastobj
            If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
        End If
Else
p = MyRound(p)
If Err.Number > 0 Then GoTo err000
pppp.item(v) = p
Do While FastSymbol(b$, ",")
If pppp.UpperMonoLimit > v Then
v = v + 1
If Not IsExp(bstack, b$, p) Then MissNumExpr: GoTo err000
If Not bstack.lastobj Is Nothing Then
    MissNumExpr
    Set bstack.lastobj = Nothing
    GoTo err000
End If
pppp.item(v) = MyRound(p)
Else
Exit Do
End If
Loop
End If
Else
GoTo err000
End If
        Exit Function
LONGERR:
    If Err.Number = 6 Then
            OverflowValue lasttype
            GoTo err000
    ElseIf Err.Number = 450 Then
            WrongOperator
            GoTo err000
    ElseIf Err.Number = 0 Then
            OverflowValue lasttype
            GoTo err000
    End If
Case2:
Exit Function
notypevarV:
noType Typename(var(v))
GoTo err000
WrongObj:
WrongObject
err000:
            Exec1 = 0: ExecuteVar = 8: Exit Function
NewCheck:
    If CheckFree(b$) Then
NewCheck2:
    ExecuteVar = 7
    Else
    SyntaxError
    End If
End Function
Function GetGlobalVarOlder(nm$, I As Long, older As Long) As Boolean
If older <= 0 Then
If varhash.Find(myUcase(nm$), I) Then
GetGlobalVarOlder = True
End If
Else
If varhash.FindOlder(myUcase(nm$), I, older) Then
GetGlobalVarOlder = True
End If
End If
End Function
Function MyRead(jump As Long, bstack As basetask, rest$, Lang As Long, Optional ByVal what$, Optional usex1 As Long, Optional exist As Boolean = False) As Boolean
Dim ps As mStiva, bs As basetask, f As Boolean, ohere$, par As Boolean, flag As Boolean, flag2 As Boolean, ok As Boolean
Dim s$, ss$, pa$, x1 As Long, y1 As Long, I As Long, myobject As Object, it As Long, useoptionals As Boolean, optlocal As Boolean
Dim M As mStiva, checktype As Boolean, allowglobals As Boolean, isAglobal As Boolean, look As Boolean, ByPass As Boolean
Dim usehandler As mHandler, ff As Long, usehandler1 As mHandler, Ar As refArray, jumpAs As Boolean, udt As Boolean, mark As Long
Const mProp = "PropReference"
Const mHdlr = "mHandler"
Const mGroup = "Group"
Const myArray = "mArray"
MyRead = True
Dim p As Variant, X As Double
Dim pppp As mArray
ohere$ = here$
Dim Col As Long
Dim ihavetype As Boolean
look = jump = 1 Or jump = 7
On jump GoTo read, refer, commit, readnew, readlocal, readlet, readfromsub, link, readUDT
Exit Function

commit:
If Len(bstack.UseGroupname) > 0 Then
    f = True
    Col = 1
    Set bs = bstack
    GoTo contFromRebound
Else
    BadReBound
    MyRead = False
End If
Exit Function

link:
flag2 = True

refer:
Col = 1
GoTo read123

readlocal:
flag = True
GoTo read123

readfromsub:
If FastSymbol(rest$, "?") Then useoptionals = True

readnew:
flag2 = True
GoTo read123
readUDT:
udt = True
readlet:
allowglobals = True
Set bs = bstack
x1 = usex1
If x1 = 1 Then
If FastSymbol(rest$, "|") Then
If bs.IsAny(p) Then
    If GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True, ok) Then
        If MemInt(VarPtr(var(I))) = 36 Then
            If placevalue2(bstack, var(I), rest$, p) Then
                MyRead = True
                Exit Function
            End If

        End If
        WrongType
        SyntaxError
        
        Exit Function
    End If
End If
End If
End If
If x1 > 3 Then x1 = Abs(IsLabel(bstack, rest$, what$))
'***********************
Select Case x1
Case 1
    If bs.IsObjectRef(myobject) Then
        MyRead = True
        If GetVar3(bstack, what$, I, , , flag, s$, checktype, isAglobal, True, ok) Then
            If Typename$(myobject) = VarTypeName(var(I)) Then
                If Typename$(var(I)) = mGroup Then
                    ss$ = bstack.GroupName
                    If s$ <> "" Then what$ = s$
                    If Len(var(I).GroupName) > Len(what$) Then
                        If var(I).IamRef Then
                            s$ = here$
                            here$ = vbNullString
                            UnFloatGroupReWriteVars bstack, what$, I, myobject
                            here = s$
                        Else
                            UnFloatGroupReWriteVars bstack, what$, I, myobject
                        End If
                        myobject.ToDelete = True
                    Else
                        bstack.GroupName = Left$(what$, Len(what$) - Len(var(I).GroupName) + 1)
                        If Len(var(I).GroupName) > 0 Then
                            what$ = Left$(var(I).GroupName, Len(var(I).GroupName) - 1)
                            s$ = here$
                            here$ = vbNullString
                            UnFloatGroupReWriteVars bstack, what$, I, myobject
                            here = s$
                            myobject.ToDelete = True
                        ElseIf var(I).IamApointer And myobject.IamApointer Then
                            Set var(I) = myobject
                        Else
                            Set myobject = Nothing
                            bstack.GroupName = ss$
                            GroupWrongUse
                            MyRead = False
                            Exit Function
                        End If
                    End If
                    Set myobject = Nothing
                    bstack.GroupName = ss$
                ElseIf Typename$(var(I)) = mHdlr Then
                    Set usehandler = myobject
                    Set usehandler1 = var(I)
                    If usehandler1.ReadOnly Then
                        MyRead = False
                        ReadOnly
                        Exit Function
                    ElseIf usehandler1.T1 = usehandler.T1 Then
                        If usehandler1.T1 = 4 Then
                            If usehandler1.objref Is usehandler.objref Then
                                Set var(I) = myobject
                            ElseIf usehandler1.objref.EnumName = usehandler.objref.EnumName Then
                                If usehandler1.objref.ExistFromOther2(usehandler) Then
                                    Set usehandler.objref = usehandler1.objref
                                ElseIf usehandler1.objref.ExistFromOther(usehandler.index_cursor) Then
                                    Set usehandler.objref = usehandler1.objref
                                    usehandler.index_start = usehandler1.objref.Index
                                    Set var(I) = usehandler
                                Else
                                    GoTo er103
                                End If
                            ElseIf usehandler1.objref.ExistFromOther2(usehandler) Then
                                Set usehandler.objref = usehandler1.objref
                            Else
                                GoTo er103
                            End If
                        Else
                            Set var(I) = myobject
                        End If
                    ElseIf usehandler1.T1 <> 4 And myobject.T1 = 3 Then
                        Set var(I) = myobject
                    Else
                        GoTo er103
                    End If
                    Set usehandler = Nothing
                    Set usehandler1 = Nothing
                Else
                    Set var(I) = myobject
                End If
            ElseIf x1 = 1 And CheckIsmArray(myobject) Then
                Set usehandler = New mHandler
                Set var(I) = usehandler
                usehandler.T1 = 3
                Set usehandler.objref = myobject
                Set myobject = Nothing
                Set usehandler = Nothing
            Else
                If TypeOf myobject Is mHandler Then
                    Set usehandler = myobject
                    If usehandler.T1 = 4 Then
                        p = usehandler.index_cursor
                        Set myobject = Nothing
                        Set usehandler = Nothing
                        GoTo itisinumber
                    End If
                End If
                GoTo er103
            End If
        Else
            I = globalvar(what$, 0)
            If Typename$(myobject) = mGroup Then
                If myobject.IamApointer Then
                    Set var(I) = myobject
                Else
                    UnFloatGroup bstack, bstack.GroupName + what$, I, myobject, here$ = vbNullString Or Len(bstack.UseGroupname) > 0, , True
                    myobject.ToDelete = True
                End If
            ElseIf Typename$(myobject) = "mEvent" Then
                Set var(I) = myobject
            ElseIf Typename$(myobject) = "lambda" Then
                Set var(I) = myobject
                If ohere$ = vbNullString Then
                    GlobalSub what$ + "()", "", , , I
                Else
                    GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , I
                End If
            ElseIf Typename$(myobject) = mHdlr Then
                Set usehandler = myobject
                If usehandler.indirect > -1 Then
                    Set var(I) = MakeitObjectGeneric(usehandler.indirect)
                Else
                    Set var(I) = usehandler
                End If
            ElseIf Typename$(myobject) = myArray Then
                Set usehandler = New mHandler
                Set var(I) = usehandler
                usehandler.T1 = 3
                Set usehandler.objref = myobject
                Set usehandler = Nothing
            Else
                Set var(I) = myobject
            End If
        End If
    ElseIf bs.IsNumber(p) Then
itisinumber:
        If GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True, ok) Then
            If MyIsObject(var(I)) Then
                If TypeOf var(I) Is Group Then
                    If var(I).HasSet Then
                        Set M = bstack.soros
                        Set bstack.Sorosref = New mStiva
                        bstack.soros.PushVal p
                        NeoCall2 bstack, what$ + "." + ChrW(&H1FFF) + ":=()", ok
                        Set bstack.Sorosref = M
                        Set M = Nothing
                    Else
                        GoTo there182741
                    End If
                Else
there182741:
                    If TypeOf var(I) Is Constant Then
                        If var(I).flag Then
                            CantAssignValue
                            Exit Function
                        Else
                            
                            '' find it
                            
                        End If
                        
                    ElseIf TypeOf var(I) Is mHandler Then
                        Set usehandler = var(I)
                        If usehandler.T1 <> 4 Then GoTo er104
                        Set myobject = usehandler.objref.SearchValue(p, ok)
                        If ok Then
                            Set var(I) = myobject
                        Else
                            GoTo er112
                        End If
                    Else
                        GoTo er104
                    End If
                End If
            Else
                If udt Then
                    PlaceValue2UDT var(I), rest$, p
                ElseIf checktype Then
                    If ihavetype Then
                        If VarType(var(I)) <> VarType(p) Then
                            GoTo er109
                        ElseIf AssignTypeNumeric(p, VarType(var(I))) Then
                            var(I) = p
                        Else
                            GoTo er105
                        End If
                    ElseIf AssignTypeNumeric(p, VarType(var(I))) Then
                        var(I) = p
                    Else
                        GoTo er105
                    End If
                Else
                    var(I) = p
                End If
            End If
        ElseIf I = -1 Then
                bstack.SetVar what$, p
        Else
                If Not exist Then globalvar what$, p Else Nosuchvariable what$
        End If
    Else
        If bs.IsString(s$) Then
        GoTo jump001
           ' If GetVar3(bstack, what$, i, , , flag, s$, checktype, isAglobal, True, ok) Then
                
            'End If
        End If
    End If
Case 3
    If bs.IsString(s$) Then
        MyRead = True
jump001:
        If GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True) Then

            If MyIsObject(var(I)) Then
                If TypeOf var(I) Is Group Then
                    Set M = bstack.soros
                    Set bstack.Sorosref = New mStiva
                    bstack.soros.PushStr s$
                    NeoCall2 bstack, Left$(what$, Len(what$) - 1) + "." + ChrW(&H1FFF) + ":=()", ok
                    Set bstack.Sorosref = M
                    Set M = Nothing
                ElseIf TypeOf var(I) Is Constant Then
                    CantAssignValue
                    MyRead = False
                    Exit Function
                Else
                    CheckVar var(I), s$
                End If
            Else
                var(I) = s$
            End If
        ElseIf I = -1 Then
            bstack.SetVar what$, s$
        Else
            If Not exist Then globalvar what$, s$ Else Nosuchvariable what$
        End If
    Else
        bstack.soros.drop 1
        MissStackStr
        MyRead = False
    End If
   
Case 4
    If bs.IsNumber(p) Then
        MyRead = True
        If GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True) Then
            var(I) = MyRound(p)
        ElseIf I = -1 Then
            bstack.SetVar what$, p
        Else
        If Not exist Then globalvar what$, MyRound(p) Else Nosuchvariable what$
        End If
    Else
        bstack.soros.drop 1
        MissStackNumber
        MyRead = False
    End If
Case 5, 7
    MyRead = False
    If FastSymbol(rest$, ")") Then
        MyRead = globalArrByPointer(bs, bstack, what$, flag2, allowglobals): If Not MyRead Then SyntaxError: Exit Function
    Else
        If neoGetArray(bstack, what$, pppp) And Not flag2 Then
            If Not NeoGetArrayItem(pppp, bs, what$, it, rest$) Then Exit Function
        Else
            Exit Function
        End If
        If IsOperator(rest$, ".") Then
            If Not pppp.ItemType(it) = mGroup Then
                MyEr "Expected group", "Περίμενα ομάδα"
                MyRead = False: Exit Function
            Else
                I = 1
                aheadstatus rest$, False, I
                ss$ = Left$(rest$, I - 1)
                MyRead = SpeedGroup(bstack, pppp, "@READ", ".", ss$, it) <> 0
                Set pppp = Nothing
                rest$ = Mid$(rest$, I)
            End If
        ElseIf IsOperator(rest$, "|") Then
        
        
        MyRead = placevalue(bs, pppp, it, rest$)
          
        Else
            If bs.IsObjectRef(myobject) Then
                If Typename$(myobject) = mGroup Then
                    If myobject.IamFloatGroup Then
                        Set pppp.item(it) = myobject
                        Set myobject = Nothing
                    Else
                        BadGroupHandle
                        MyRead = False
                        Set myobject = Nothing
                        Exit Function
                    End If
                    ElseIf Typename$(myobject) = "lambda" Then
                        Set pppp.item(it) = myobject
                        Set myobject = Nothing
                    ElseIf Typename$(myobject) = myArray Then
                                  If myobject.Arr Then
                        Set pppp.item(it) = CopyArray(myobject)
                    Else
                        Set pppp.item(it) = myobject
                    End If
                    Set myobject = Nothing
                ElseIf Typename$(myobject) = mHdlr Then
                    Set usehandler = myobject
                    If usehandler.indirect > -0 Then
                        Set pppp.item(it) = usehandler
                    Else
                        p = usehandler.T1
                        If CheckDeepAny(myobject) Then
                            If TypeOf myobject Is mHandler Then
                                Set pppp.item(it) = myobject
                            Else
                                Set usehandler = New mHandler
                                Set pppp.item(it) = usehandler
                                usehandler.T1 = p
                                Set usehandler.objref = myobject
                                Set usehandler = Nothing
                            End If
                            Set myobject = Nothing
                        End If
                    End If
                ElseIf Typename$(myobject) = mProp Then
                    Set pppp.item(it) = myobject
                    Set myobject = Nothing
                End If
            ElseIf Not bs.IsNumber(p) Then
                If bs.IsString(s$) Then
                    pppp.item(it) = s$
                Else
                    bstack.soros.drop 1
                    MissStackNumber
                    MyRead = False
                    Exit Function
                End If
            ElseIf x1 = 7 Then
                pppp.item(it) = Round(p)
            Else
                pppp.item(it) = p
            End If
        End If
        MyRead = True
    End If
 Case 6
    MyRead = False
    If FastSymbol(rest$, ")") Then
        MyRead = globalArrByPointer(bs, bstack, what$, flag2): If Not MyRead Then SyntaxError: Exit Function
    Else
        If neoGetArray(bstack, what$, pppp) And Not flag2 Then
            If Not NeoGetArrayItem(pppp, bs, what$, it, rest$) Then Exit Function
        Else
            Exit Function
        End If
        If Not bs.IsString(s$) Then
            If bs.IsObjectRef(myobject) Then
                If Typename$(myobject) = "lambda" Then
                    Set pppp.item(it) = myobject
                    Set myobject = Nothing
                ElseIf Typename$(myobject) = mGroup Then
                    Set pppp.item(it) = myobject
                    Set myobject = Nothing
                ElseIf Typename$(myobject) = myArray Then
                    If myobject.Arr Then
                        Set pppp.item(it) = CopyArray(myobject)
                    Else
                        Set pppp.item(it) = myobject
                    End If
                    Set myobject = Nothing
                ElseIf Typename$(myobject) = mHdlr Then
                    Set usehandler = myobject
                    If usehandler.indirect > -0 Then
                        Set pppp.item(it) = myobject
                    Else
                        p = usehandler.T1
                        If CheckDeepAny(myobject) Then
                            If TypeOf myobject Is mHandler Then
                                Set pppp.item(it) = myobject
                            Else
                                Set usehandler = New mHandler
                                Set pppp.item(it) = usehandler
                                usehandler.T1 = p
                                Set usehandler.objref = myobject
                                Set usehandler = Nothing
                            End If
                            Set myobject = Nothing
                        End If
                    End If
                ElseIf Typename$(myobject) = mProp Then
                    Set pppp.item(it) = myobject
                    Set myobject = Nothing
                Else
                    MissStackStr
                    Exit Function
                End If
            Else
                bstack.soros.drop 1
                MissStackStr
                Exit Function
            End If
        Else
            If Not MyIsObject(pppp.item(it)) Then
                pppp.item(it) = s$
            ElseIf pppp.ItemType(it) = mGroup Then
            ' do something
            Else
                Set pppp.item(it) = New Document
                CheckVar pppp.item(it), s$
            End If
        End If
        MyRead = True
    End If
'*****************************************************
    Case 8
    GoTo jump8
    End Select
    p = 0#
    Exit Function
read:
If FastSymbol(rest$, "?") Then useoptionals = True

flag2 = Fast2LabelNoNum(rest$, "ΝΕΟ", 3, "NEW", 3, 3)
If Not flag2 Then flag = Fast2LabelNoNum(rest$, "ΤΟΠΙΚΑ", 6, "LOCAL", 5, 6)
read123:
mark = varhash.count - 1
Set bs = bstack
contFromRebound:
par = Fast2LabelNoNum(rest$, "ΑΠΟ", 3, "FROM", 4, 4)
If par And f Then
SyntaxError
MyRead = False
Exit Function
End If
If par Then
' make it general...
x1 = Abs(IsLabelBig(bstack, rest$, ss$, , , , par))

If x1 > 0 And x1 <> 1 Then rest$ = ss$ + " " + rest$
If x1 = 1 Then

 If getvar2(bstack, ss$, I, , , flag) Then
 If MyIsObject(var(I)) Then
            ' need to make new stack frame with pointers to
        If Typename(var(I)) <> mGroup Then MyRead = True: Exit Function
        Set ps = New mStiva
        If ohere$ <> "" And Not var(I).IamGlobal Then
        Set myobject = var(I).PrepareSoros(var(), ohere$ + ".")
        Else
        Set myobject = var(I).PrepareSoros(var(), "")
        End If
                
        With myobject
               For x1 = 1 To .Total
                  s$ = .StackItem(x1) & " "
                  If Left$(s$, 1) = "*" Then '' we have a group
                  s$ = Split(Mid$(s$, 2))(0)
                  Else
                  s$ = Split(s$)(0)
                  End If

 ''we place references

                If Right$(s$, 2) = "()" Then
                        ps.DataStr Left$(s$, Len(s$) - 2)
                ElseIf Right$(s$, 1) = "(" Then
                        ps.DataStr Left$(s$, Len(s$) - 1)
                Else
                        ps.DataStr s$
                End If
 ''bstack.Soros.DataStr .StackItem(X1)
            Next x1
        End With
Set myobject = Nothing
 Set bs = New basetask
    Set bs.Sorosref = ps
    If FastSymbol(rest$, ";") Then
    bstack.soros.MergeTop ps
    
    MyRead = True
    Exit Function
    End If
    If Not FastSymbol(rest$, ",") Then
    MissPar
    MyRead = False
 Exit Function
    End If
    
 Else
 MissingGroup
 MyRead = False
 Exit Function
 End If
 Else
 MissingGroup
 MyRead = False
 Exit Function
 End If
 Col = 1 ' this is a switch... look down
 
ElseIf IsStrExp(bstack, rest$, ss$) Then
Set ps = New mStiva
Do While ss$ <> ""
If ISSTRINGA(ss$, pa$) Then
ps.DataStr pa$
ElseIf IsNumberD(ss$, X) Then
ps.DataVal X
Else
Exit Do
End If
Loop
Set bs = New basetask
Set bs.Sorosref = ps
    If Not FastSymbol(rest$, ",") Then
    MissPar
    MyRead = False
 Exit Function
    End If
End If
End If
' from here is not MyRead = True
Do
again1:
MyRead = False
ihavetype = False
If look Then If FastSymbol(rest$, "?") Then useoptionals = True
If FastSymbol(rest$, ",") Then bs.soros.drop 1: GoTo again1
If FastSymbol(rest$, "&") Or Col = 1 Then
' so now for GROUP variables we use only by reference
Select Case Abs(IsLabel(bstack, rest$, what$))
Case 1
If bs.IsString(s$) Then
     
          
    
    If GetGlobalVarOlder(s$, I, mark) Then
    
conthereifglobal:
        If flag2 Then
            If Not f Then
                If Not flag Then
                    If ohere$ <> "" Then
                        GoTo contpush12
                    Else
                        NoSecReF
                        Exit Do
                    End If
                End If
            End If
        ElseIf GetVar3(bstack, what$, it, , , flag, , , , True) And Not f Then
               If flag Then
                   If ohere$ <> "" Then GoTo contpush12
               Else
                          If GetlocalVar(what$, y1) = False And ohere$ <> "" Then
                                   GoTo contpush12
                              Else
                                  NoSecReF
                                 Exit Do
                            End If
                End If
                
   
   Else
contpush12:
       what$ = myUcase(what$)
backfromstr:
' no jumpAs here
                If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
                   If Not MyIsObject(var(I)) Then
                        p = var(I)
                        If Not varhash.vType(varhash.Index) Then
                            If Not Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then MyRead = False: MissType: Exit Function
                            GoTo jumpref02
                        End If
checkconstant:
                        Select Case VarType(p)
                        Case vbDecimal
                            If Not Fast2Varl(rest$, "ΑΡΙΘΜΟΣ", 7, "DECIMAL", 7, 7, ff) Then MyRead = False: MissType: Exit Function
                        Case vbDouble
                            If Not Fast2Varl(rest$, "ΔΙΠΛΟΣ", 6, "DOUBLE", 6, 6, ff) Then MyRead = False: MissType: Exit Function
                        Case vbSingle
                            If Not Fast2Varl(rest$, "ΑΠΛΟΣ", 5, "SINGLE", 6, 6, ff) Then MyRead = False: MissType: Exit Function
                        Case vbBoolean
                            If Not Fast2Varl(rest$, "ΛΟΓΙΚΟΣ", 7, "BOOLEAN", 7, 7, ff) Then MyRead = False: MissType: Exit Function
                        Case vbLong
                            If Not Fast2Varl(rest$, "ΜΑΚΡΥΣ", 6, "LONG", 4, 6, ff) Then MyRead = False: MissType: Exit Function
                        Case vbInteger
                            If Not Fast2Varl(rest$, "ΑΚΕΡΑΙΟΣ", 8, "INTEGER", 7, 8, ff) Then MyRead = False: MissType: Exit Function
                        Case vbCurrency
                            If Not Fast2Varl(rest$, "ΛΟΓΙΣΤΙΚΟΣ", 10, "CURRENCY", 8, 10, ff) Then MyRead = False: MissType: Exit Function
                        Case 20
                            If Not Fast2Varl(rest$, "ΜΑΚΡΥΣ", 6, "LONG", 4, 6, ff) Then MyRead = False: MissType: Exit Function
                            If Not Fast2Varl(rest$, "ΜΑΚΡΥΣ", 6, "LONG", 4, 6, ff) Then MyRead = False: MissType: Exit Function
                        Case vbString
                            If Not Fast2Varl(rest$, "ΓΡΑΜΜΑ", 6, "STRING", 6, 6, ff) Then
                                MyRead = False: MissType: Exit Function
                            End If
                        Case vbByte
                            If Not Fast2Varl(rest$, "ΨΗΦΙΟ", 5, "BYTE", 4, 5, ff) Then MyRead = False: MissType: Exit Function
                        Case vbDate
                            If Not Fast2Varl(rest$, "ΗΜΕΡΟΜΗΝΙΑ", 10, "DATE", 4, 10, ff) Then MyRead = False: MissType: Exit Function
                        Case 36
                            If TypeOf p Is cxComplex Then
                            If Not Fast2Varl(rest$, "ΜΙΓΑΔΙΚΟΣ", 9, "COMPLEX", 7, 9, ff) Then MyRead = False: MissType: Exit Function
                            Else
                             IsLabel bstack, rest$, ss$
                             If LCase(Typename(p)) <> LCase(ss$) Then
                                MyRead = False: MissType: Exit Function
                             End If
                            End If
                        Case Else
                            p = IsLabel(bstack, rest$, (what$))  ' just throw any name
                        End Select
                    Else
                        If TypeOf var(I) Is Group Then
                            If Fast2Varl(rest$, "ΟΜΑΔΑ", 5, "GROUP", 5, 5, ff) Then
                            If var(I).IamApointer Then
                                If var(I).link.IamFloatGroup Then GoTo errgr
                                    If Len(var(I).lasthere) = 0 Then
                                        If Not GetVar(bstack, var(I).GroupName, I, True) Then GoTo errgr
                                    Else
                                        If Not GetVar(bstack, var(I).lasthere + "." + var(I).GroupName, I, True) Then GoTo errgr

                                    End If
                                 End If
                            ElseIf Fast2Varl(rest$, "ΔΕΙΚΤΗΣ", 7, "POINTER", 7, 7, ff) Then
                                     If Not var(I).IamApointer Then MyRead = False: MissType: Exit Function
                            Else
                            If FastSymbol(rest$, "*") Then
                                If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
                                If Not var(I).IamApointer Then GoTo errgr
                                If var(I).link.IamFloatGroup Then
                                    If Not var(I).link.TypeGroup(s$) Then GoTo errgr
                                Else
                                    If Len(var(I).lasthere) = 0 Then
                                    If GetVar(bstack, var(I).GroupName, it, True) Then
                                        If Not var(it).TypeGroup(s$) Then GoTo errgr
                                    Else
                                        GoTo noref01
                                    End If
                                    Else
                                        If GetVar(bstack, var(I).lasthere + "." + var(I).GroupName, it, True) Then
                                            If Not var(it).TypeGroup(s$) Then GoTo errgr
                                        Else
                                            GoTo noref01
                                        End If
                                    End If
                                    it = 0
                                End If

                            ElseIf FastPureLabel(rest$, s$, , True) = 1 Then
                                    If var(I).IamApointer Then

                                    If var(I).link.IamFloatGroup Then
                                        MyRead = False: MissType: Exit Function
                                    End If
                                        If Len(var(I).lasthere) = 0 Then
                                            If Not GetVar(bstack, var(I).GroupName, I, True) Then
                                                GoTo errgr
                                            End If
                                        ElseIf Not GetVar(bstack, var(I).lasthere + "." + var(I).GroupName, I, True) Then
                                                GoTo errgr
                                        End If
                                    End If
                                    If Not var(I).TypeGroup(s$) Then GoTo errgr
                            Else
                                    MyRead = False: MissType: Exit Function
                            End If
                            End If
                        ElseIf TypeOf var(I) Is mHandler Then
                            Set usehandler = var(I)
                            If TypeOf usehandler.objref Is mArray Then
                                If Not Fast2Varl(rest$, "ΠΙΝΑΚΑΣ", 7, "ARRAY", 5, 7, ff) Then MyRead = False: MissType: Exit Function
                            ElseIf TypeOf usehandler.objref Is FastCollection Then
                                If Not Fast2Varl(rest$, "ΚΑΤΑΣΤΑΣΗ", 9, "INVENTORY", 9, 9, ff) Then
                                     If Not Fast2Varl(rest$, "ΛΙΣΤΑ", 5, "LIST", 4, 5, ff) Then
                                        If Not Fast2Varl(rest$, "ΟΥΡΑ", 4, "QUEUE", 5, 5, ff) Then
                                            MyRead = False: MissType: Exit Function
                                        ElseIf Not usehandler.objref.IsQueue Then
                                            MyRead = False: MissType: Exit Function
                                            Exit Function
                                        End If
                                    ElseIf usehandler.objref.IsQueue Then
                                        MyRead = False: MissType: Exit Function
                                        Exit Function
                                    End If
                                End If
                            ElseIf TypeOf usehandler.objref Is mStiva Then
                                If Not Fast2Varl(rest$, "ΣΩΡΟΣ", 5, "STACK", 5, 5, ff) Then MyRead = False: MissType: Exit Function
                            ElseIf TypeOf usehandler.objref Is MemBlock Then
                                If Not Fast2Varl(rest$, "ΔΙΑΡΘΡΩΣΗ", 9, "BUFFER", 6, 9, ff) Then
                                    If FastPureLabel(rest$, s$, , True, , , , True) = 1 Then
                                        If usehandler.objref.UseStruct Then
                                            If usehandler.objref.structref.Tag = s$ Then
                                                GoTo jumpref01
                                            End If
                                        End If
                                    End If
                                    MyRead = False: MissType: Exit Function
                                End If
                            ElseIf usehandler.T1 = 4 Then
                                If FastPureLabel(rest$, s$, , True, , , False) = 1 Then
                                
                                    If Not s$ = myUcase(usehandler.objref.EnumName, True) Then
                                        If GetSub(s$ + "()", x1) Then
                                            If sbf(x1).IamAClass Then
                                                GoTo er113
                                            End If
                                        ElseIf GetSub(bstack.GroupName + s$ + "()", x1) Then
                                            If sbf(x1).IamAClass Then
                                                GoTo er113
                                            End If
                                        End If
                                        p = usehandler.index_cursor
                                        Set usehandler = Nothing
                                    
                                        GoTo checkconstant
                                    Else
                                        FastPureLabel rest$, s$
                                        
                                    End If
                                Else
                                    MyRead = False: MissType: Exit Function
                                End If
                              
                   
                            Else
                                p = IsLabel(bstack, rest$, (what$)) ' just throw any name
                            End If
                        ElseIf TypeOf var(I) Is lambda Then
                            If Not Fast2Varl(rest$, "ΛΑΜΔΑ", 5, "LAMBDA", 6, 6, ff) Then MyRead = False: MissType: Exit Function
                        
                        ElseIf TypeOf var(I) Is mEvent Then
                            If Not Fast2Varl(rest$, "ΓΕΓΟΝΟΣ", 7, "EVENT", 5, 5, ff) Then MyRead = False: MissType: Exit Function
                        ElseIf TypeOf var(I) Is Constant Then
                            
                            p = var(I)
                            If var(I).vType Then
                                If Not Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then MyRead = False: MissType: Exit Function
                                GoTo jumpref02
                            End If
                            GoTo checkconstant
                         ElseIf TypeOf var(I) Is refArray Then
                         
                        If FastSymbol(rest$, "*") Then
                            If IsLabel(bstack, rest$, ss$) = 0 Then
                                GoTo er110
                            End If
                            If LCase(ss$) = "long" Then
                                If IsLabel(bstack, rest$, ss$) = 1 Then
                                If LCase(ss$) = "long" Then
                                    ss$ = "long long"
                                    
                                Else
                                    SyntaxError
                                    MyRead = False
                                    Exit Function
                                End If
                                Else
                                ss$ = "long"
                                End If
                            End If
                            
                                Set Ar = var(I)
                                If LCase(VarTypeName(Ar(0, 0))) <> LCase(ss$) Then
                                    If Ar.vtType(0) = vbVariant And LCase(ss$) = "variant" Then
                                    
                                    Else
                                    GoTo er103
                                    End If
                                End If
                                Set Ar = Nothing
                                GoTo jumpref01
                                
                            
                        End If
                        ElseIf TypeOf var(I) Is BigInteger Then
                            If Not Fast2Varl(rest$, "ΜΕΓΑΛΟΣΑΚΕΡΑΙΟΣ", 15, "BIGINTEGER", 10, 15, ff) Then MyRead = False: MissType: Exit Function
                        Else
                        
                            p = IsLabel(bstack, rest$, (what$)) ' just throw any name
                        End If
                    End If
                End If
jumpref01:
            If Not LinkGroup(bstack, what$, var(I)) Then
jumpref02:
            If f Then

                If Not ReboundVar(bstack, what$, I) Then globalvar what$, I, True
            Else
    
                globalvar what$, I, True, useType:=varhash.vType(varhash.Index)
                If VarTypeName(var(I)) = "lambda" Then
islambda:
                
                    If ohere$ = vbNullString Then
                        GlobalSub what$ + "()", "", , , I
                    Else
                        GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , I
                    End If
                ElseIf VarTypeName(var(I)) = "Constant" Then
                    If var(I).flag Then GoTo islambda
                End If
                
            End If
        Else
            it = globalvar(what$, it)
            MakeitObject2 var(it)
            Dim aG As Group, bG As Group
            Set bG = var(it)
            If var(I).IamApointer Then
            If var(I).link.IamFloatGroup Then
               Set bG.LinkRef = var(I).link
                bG.IamApointer = True
                bG.isRef = True
            Else
            
            Set aG = var(I).link
            
            With aG
            
                bG.edittag = .edittag
                bG.FuncList = .FuncList
                bG.GroupName = myUcase(what$) + "."
                'Set bG.Sorosref = .soros.Copy
                If UBound(.Fields) > 1 Then bG.Fields = .Fields
                bG.HasValue = .HasValue
                bG.HasSet = .HasSet
                bG.HasStrValue = .HasStrValue
                bG.HasParameters = .HasParameters
                bG.HasParametersSet = .HasParametersSet
                bG.HasRemove = .HasRemove
                        Set bG.Events = .Events
            
                bG.highpriorityoper = .highpriorityoper
                bG.HasUnary = .HasUnary
                If Len(here$) > 0 Then
                bG.Patch = here$ + "." + what$
                Else
                bG.Patch = what$
                End If
                Set bG.mytypes = .mytypes
            End With
            End If
            
            Else
            Set aG = var(I)
            With aG
            
                bG.edittag = .edittag
                bG.FuncList = .FuncList
                bG.GroupName = myUcase(what$) + "."
             '   Set bG.Sorosref = .soros.Copy
                If UBound(.Fields) > 1 Then bG.Fields = .Fields
                bG.HasValue = .HasValue
                bG.HasSet = .HasSet
                bG.HasStrValue = .HasStrValue
                bG.HasParameters = .HasParameters
                bG.HasParametersSet = .HasParametersSet
                bG.HasRemove = .HasRemove
                        Set bG.Events = .Events
            
                bG.highpriorityoper = .highpriorityoper
                bG.HasUnary = .HasUnary
                If Len(here$) > 0 Then
                bG.Patch = here$ + "." + what$
                Else
                bG.Patch = what$
                End If
                Set bG.mytypes = .mytypes
            End With
             bG.IamRef = Len(bstack.UseGroupname) > 0
             End If
            If var(I).HasStrValue Then
                globalvar what$ + "$", it, True
            End If
            Set bG = Nothing
            Set aG = Nothing
        End If
        MyRead = True
    End If
     Else
        If Left$(s$, 1) = "#" Then  ' for copy in
            s$ = Mid$(s$, 2)
            If IsNumberNew(bstack, (s$), p, False) Then
                ss$ = "_" + str$(var2used)
                If bstack.SubLevel > 0 Then
                    MyEr "not for Read statement", "όχι για την Διάβασε"
                    MyRead = False
                    Exit Function
                End If
                I = globalvar(ss$, 0#, , True)
                If bstack.IamChild Then FeedCopyInOut bstack.Parent, s$, I, ""
                UnFloatGroup bstack, ss$, I, bstack.lastobj, True, , True
                bstack.lastobj.ToDelete = True
                GoTo conthereifglobal
            End If
        Else
            If GetVar3(bstack, s$, I, True, , , , , , True) Then GoTo conthereifglobal
        End If
noref01:
         NoReference
        MyRead = False
        Exit Function
     End If
    Else
If bs.IsObjectRef(myobject) Then

If TypeOf myobject Is Group Then
If myobject.IamApointer Then
I = AllocVar()
Set var(I) = myobject
Set myobject = Nothing
GoTo backfromstr
End If
GoTo er103
Else
If Not GetVar3(bstack, what$, I, , , flag, , , , True) Then I = globalvar(what$, 0)
If jumpAs Then jumpAs = False: GoTo existAs06
ff = 1
If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs06:
If Not FastPureLabel(rest$, ss$) = 1 Then
    GoTo er110
Else
CheckItemType bstack, CVar(myobject), vbNullString, s$, , ok
If ok Then
    If ss$ <> s$ Then GoTo er103

ElseIf UCase(ss$) <> UCase(s$) Then GoTo er103
End If
End If
End If
End If
CreateFormOtherObject var(I), myobject
 MyRead = True
Set myobject = Nothing
Else
    GoTo noref01
    End If
    End If

Case 3, 4
    If bs.IsString(s$) Then
        If GetGlobalVarOlder(s$, I, mark) Then
            If flag2 Then
            If Not f Then
                If Not flag Then
                    If ohere$ <> "" Then
                            'GoTo contpush12
               If MyIsObject(var(I)) Then
                    If Typename(var(I)) = "lambda" Then
                        If ohere$ = vbNullString Then
                            GlobalSub what$ + "()", "", , , I
                        Else
                            GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , I
                        End If
                        globalvar what$, I, True
                    ElseIf Typename(var(I)) = mGroup Then
                        what$ = Left$(what$, Len(what$) - 1)
                        GoTo backfromstr
                    Else
                        globalvar what$, I, True
                    End If
                Else
                     globalvar what$, I, True
                 End If
               MyRead = True
               Else
                        NoSecReF
                        Exit Do
                    End If
                End If
            End If
        ElseIf GetVar3(bstack, what$, it, , , flag, , , , True) And Not f Then
            If Not flag Then
                          If GetlocalVar(what$, y1) = False And ohere$ <> "" Then

                                GoTo contherestr
                              Else
                                  NoSecReF
                                 Exit Do
                            End If
                End If
                NoSecReF
                Exit Do
            Else
                If f Then
                    If Not ReboundVar(bstack, what$, I) Then globalvar what$, I, True
                Else
contherestr:

            If MyIsObject(var(I)) Then
                    If Typename(var(I)) = "lambda" Then
                         If ohere$ = vbNullString Then
                             GlobalSub what$ + "()", "", , , I
                         Else
                             GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , I
                         End If
                         globalvar what$, I, True
                    ElseIf Typename(var(I)) = mGroup Then
                         what$ = Left$(what$, Len(what$) - 1)
                         GoTo backfromstr
                    Else
                         globalvar what$, I, True
                     End If
                Else
                globalvar what$, I, True
                End If
                End If
                MyRead = True
            End If
        Else
            GoTo noref01
        End If
    Else
        GoTo noref01
    End If
Case 5, 6, 7
    If bs.IsString(s$) Then  ' get the pointer!!!!!
        If lookOne(s$, "{") Then
            If Not FastSymbol(rest$, ")") Then
                GoTo er107
            Else
                s$ = Left$(what$, Len(what$) - 1) + " " + s$
conter00:
                If f Then
                    ss$ = here$
                    here$ = vbNullString
                    If Not MyFunction(0, bstack, s$, 1, , flag2) Then
                        GoTo er106
                        Exit Do
                    Else
                        sbf(bstack.IndexSub).sbgroup = s$
                        I = Len(s$)
                        If I > 0 Then
                            If varhash.Find(Left$(s$, I - 1), I) Then sbf(bstack.IndexSub).tpointer = I
                        Else
                            sbf(bstack.IndexSub).tpointer = 0
                        End If
                    End If
                    here$ = ss$
                Else
                    If Not MyFunction(0, bstack, s$, 1, , flag2) Then
                        GoTo er106
                        Exit Do
                    Else
                        sbf(bstack.IndexSub).sbgroup = s$
                        I = Len(s$)
                        If I > 0 Then
                            If varhash.Find(Left$(s$, I - 1), I) Then sbf(bstack.IndexSub).tpointer = I
                        Else
                            sbf(bstack.IndexSub).tpointer = 0
                        End If
                    End If
                End If
                MyRead = True
            End If
        Else
            I = CopyArrayItemsNoFormated(bstack, s$)
            If I <> 0 Then
                If Not FastSymbol(rest$, ")") Then GoTo er107
                If f And I > 0 Then '' look about f - work for refer but no refer can be done...why???
                    If Not ReboundArr(bstack, what$, I) Then GoTo arrconthere
                Else
arrconthere:
                what$ = myUcase(what$)
                If ohere$ = vbNullString Then
                    If varhash.ExistKey(what$) Then
                      If flag2 And Not f And Not flag Then
                        If I < 0 Then I = -I
                           varhash.ItemCreator what$, I
                           ' what$ now is empty string
                    Else
                        GoTo er108
                        End If
                    Else
                        varhash.ItemCreator what$, I
                    End If
                Else
                    If varhash.ExistKey(ohere$ + "." + what$) Then
                    If flag2 And Not f And Not flag Then
                        I = Abs(I)
                           varhash.ItemCreator ohere$ + "." + what$, I, True, True
                    Else
                        GoTo er108
                    End If
                    Else
                      I = Abs(I)
                    varhash.ItemCreator ohere$ + "." + what$, I, True, True
                    End If
                End If

            End If
            MyRead = True
        Else
        ' get function
        If GetSub(s$, I) Then
        If Len(sbf(I).sbgroup) > 0 Then
        If sbf(I).Extern > 0 Then
        s$ = Left$(what$, Len(what$) - 1) + " {CALL EXTERN" + str$(sbf(I).Extern) + "'" + ChrW(&H1FFD) + "}" + sbf(I).sbgroup
        Else
        s$ = Left$(what$, Len(what$) - 1) + " {" + sbf(I).sb + "}" + sbf(I).sbgroup
        End If
        Else
        If sbf(I).Extern > 0 Then
        s$ = Left$(what$, Len(what$) - 1) + " {CALL EXTERN" + str$(sbf(I).Extern) + "'" + ChrW(&H1FFD) + "}"
        Else
        s$ = Left$(what$, Len(what$) - 1) + " {" + sbf(I).sb + "}"
        End If
        End If
        GoTo conter00
        End If
        End If
    End If
End If
Case Else
    Exit Do
End Select
Else
' here not for LET any more
read2:
x1 = Abs(IsLabel(bstack, rest$, what$))
fromEnumDeref:

Select Case x1
Case 1
    If bs.IsObjectRef(myobject) Then
        MyRead = True
        If flag2 Then
        GoTo comehere
        ElseIf flag Then
            p = 0#
           I = globalvar(what$, p)
           GoTo contread123
        End If
       If GetVar3(bstack, what$, I, , , flag, s$, checktype, isAglobal, True, ok) Then
        If isAglobal And Not allowglobals Then
        GoTo comehere
        End If
contread123:
                If Typename$(myobject) = VarTypeName(var(I)) Then
                    If Typename$(var(I)) = mGroup Then
                                 ss$ = bstack.GroupName
                                 If Len(s$) > 0 Then what$ = s$
                                 If Len(var(I).GroupName) > Len(what$) Then
                                    If jumpAs Then jumpAs = False: GoTo existAs01
                                    ff = 1
                                    If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs01:
                                   If FastSymbol(rest$, "*") Then
                                        If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
                                            If s$ = "POINTER" Then
                                                If Not myobject.IamApointer Then GoTo errgr
                                            ElseIf s$ = "ΔΕΙΚΤΗΣ" Then
                                                If Not myobject.IamApointer Then GoTo errgr
                                            
                                            ElseIf myobject.IamApointer Then
                                                If Not myobject.link.TypeGroup(s$) Then
                                                    GoTo errgr
                                                End If
                                            Else
                                                GoTo errgr
                                            End If
                                    Else
                                        If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
                                            If s$ = "POINTER" Then
                                                If Not myobject.IamApointer Then GoTo errgr
                                            ElseIf s$ = "ΔΕΙΚΤΗΣ" Then
                                                If Not myobject.IamApointer Then GoTo errgr
                                            ElseIf s$ = "GROUP" Then
                                                ' get pointer too
                                            ElseIf s$ = "ΟΜΑΔΑ" Then
                                                ' get pointer too
                                            ElseIf myobject.IamApointer Then
                                                If Not myobject.link.TypeGroup(s$) Then
                                                    GoTo errgr
                                                End If
                                            ElseIf Not myobject.TypeGroup(s$) Then
                                                GoTo errgr

                                            End If
                                   End If
                                    End If
                                    If var(I).IamRef Then
                                        SwapStrings s$, here$
                                        here$ = vbNullString
                                        UnFloatGroupReWriteVars bstack, what$, I, myobject
                                        SwapStrings here, s$
                                        s$ = vbNullString
                                    Else
                                        UnFloatGroupReWriteVars bstack, what$, I, myobject
                                        myobject.ToDelete = True
                                    End If
                                Else
                                    If Len(var(I).Patch) > 0 Then what$ = var(I).Patch
                                    bstack.GroupName = Left$(what$, Len(what$) - Len(var(I).GroupName) + 1)
                                    If Len(var(I).GroupName) > 0 Then
                                        If jumpAs Then jumpAs = False: GoTo existAs02
                                        ff = 1
                                        If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs02:
                                            If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
                                            If Not myobject.TypeGroup(s$) Then GoTo errgr
                                        End If
                                        what$ = Left$(var(I).GroupName, Len(var(I).GroupName) - 1)
                                        SwapStrings s$, here$
                                        here$ = vbNullString
                                        UnFloatGroupReWriteVars bstack, what$, I, myobject, , , ByPass
                                        myobject.ToDelete = True
                                        ByPass = False
                                        SwapStrings here, s$
                                        s$ = vbNullString
                                    ElseIf var(I).IamApointer And myobject.IamApointer Then
                                    If jumpAs Then jumpAs = False: GoTo existAs03
                                    ff = 1
                                    If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs03:
    
                                    FastSymbol rest$, "*"
      
                                    If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
      
                                    If myobject.link.IamFloatGroup Then
                                    
                                    If Not myobject.link.TypeGroup(s$) Then
                                        If s$ = "POINTER" Then
                                        ElseIf s$ = "ΔΕΙΚΤΗΣ" Then
                                        Else
                                        GoTo errgr
                                        End If
                                    End If
                                    Else
                                        If Len(myobject.lasthere) = 0 Then
                                            If GetVar(bstack, myobject.GroupName, it, True) Then
                                                If Not var(I).TypeGroup(s$) Then GoTo errgr
                                            End If
                                        ElseIf GetVar(bstack, myobject.lasthere + "." + myobject.GroupName, it, True) Then
                                                If Not var(I).TypeGroup(s$) Then GoTo errgr
                                        End If
                                    End If
                                    
                                    End If
                                    Set var(I) = myobject
                                    
                                    Else
                                        Set myobject = Nothing
                                        bstack.GroupName = ss$
                                        GroupWrongUse
                                        MyRead = False
                                        Exit Function
                                    End If
                                End If
                                Set myobject = Nothing
                                bstack.GroupName = ss$
    
                   
                    ElseIf Typename$(var(I)) = mHdlr Then
                        Set usehandler = myobject
                        Set usehandler1 = var(I)
                        If usehandler1.ReadOnly Then
                            MyRead = False
                           ReadOnly
                           Exit Function
                           
                        ElseIf usehandler1.T1 = usehandler.T1 Then
                            If usehandler1.T1 = 4 Then
                                If usehandler1.objref Is usehandler.objref Then
                                    Set var(I) = myobject
                                ElseIf usehandler1.objref.EnumName = usehandler.objref.EnumName Then
                                    If usehandler1.objref.ExistFromOther2(usehandler) Then
                                        Set usehandler.objref = usehandler1.objref
                                    ElseIf usehandler1.objref.ExistFromOther(usehandler.index_cursor) Then
                                        Set usehandler.objref = usehandler1.objref
                                        usehandler.index_start = usehandler1.objref.Index
                                        Set var(I) = usehandler
                                    Else
                                        GoTo er103
                                    End If
                                ElseIf usehandler1.objref.ExistFromOther2(usehandler) Then
                                    Set usehandler.objref = usehandler1.objref
                                Else
                                    GoTo er103
                                End If
                            Else
                            
                                Set var(I) = myobject
                           End If
                        ElseIf usehandler1.T1 <> 4 And usehandler.T1 = 3 Then
                           Set var(I) = myobject
                        Else
                          GoTo er103
                        End If
                    Else
                    Set var(I) = myobject
                    End If
                    If jumpAs Then jumpAs = False: GoTo existAs04
                    ff = 1
                    
                    If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs04:
                    If Not FastPureLabel(rest$, s$) = 1 Then
                        GoTo er110
                    End If
                    ' no second time
                    If Not myobject Is Nothing Then
                    If TypeOf myobject Is mHandler Then
                        
                        Set usehandler = myobject
                        s$ = myUcase(s$)
                        If usehandler.T1 = 1 Then
                        ff = 0
                        
                        If Not Fast2Varl(s$, "ΚΑΤΑΣΤΑΣΗ", 9, "INVENTORY", 9, 9, ff) Then
                        If usehandler.objref.IsQueue Then
                            If Not Fast2Varl(s$, "ΟΥΡΑ", 4, "QUEUE", 5, 5, ff) Then GoTo er103
                        Else
                            If Not Fast2Varl(s$, "ΛΙΣΤΑ", 5, "LIST", 4, 5, ff) Then GoTo er103
                        End If
                        End If
                        ElseIf usehandler.T1 = 3 Then
                        ff = 0
                        If Fast2Varl(s$, "ΠΙΝΑΚΑΣ", 7, "ARRAY", 5, 7, ff) Then
                            If Not CheckIsmArray(myobject) Then GoTo er103
                        ElseIf Fast2Varl(s$, "ΣΩΡΟΣ", 5, "STACK", 5, 5, ff) Then
                            If Not CheckIsmStiva(myobject) Then GoTo er103
                        Else
                            GoTo er103
                        End If
                        ElseIf usehandler.T1 = 4 Then
                        Else
                        GoTo er103
                        End If
                    End If
                    End If
                    
                    End If
                ElseIf x1 = 1 And CheckIsmArray(myobject) Then
               
                    ''bstack.lastobj.CopyArray pppp
                    Set usehandler = New mHandler
                    Set var(I) = usehandler
                    usehandler.T1 = 3
                    Set usehandler.objref = myobject
                    Set usehandler = Nothing
                    If jumpAs Then jumpAs = False: GoTo existAs07
                    ff = 1
                    If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs07:
                        If Not Fast2Varl(rest$, "ΠΙΝΑΚΑΣ", 7, "ARRAY", 5, 7, ff) Then GoTo er103
                    End If
                    Set myobject = Nothing
                ElseIf MyIsNumeric(var(I)) Then
                If Not myobject Is Nothing Then
                If TypeOf myobject Is Group Then
                If myobject.IamApointer Then
                Set var(I) = myobject
                GoTo cont10101
                End If
                Else
                If TypeOf myobject Is mHandler Then
                    Set usehandler = myobject
                    If usehandler.T1 = 4 Then
                        p = usehandler.index_cursor
                        If MemInt(VarPtr(p)) <> vbString Then
                            p = p * usehandler.sign
                            bs.soros.PushVal p
                        Else
                            bs.soros.PushStrVariant p
                        End If
                        Set usehandler1 = Nothing
                        jumpAs = False
                        GoTo fromEnumDeref
                        
                    End If
                
                End If
                End If
                End If
                GoTo er103
                ElseIf IsObject(var(I)) Then
                If TypeOf var(I) Is Group Then
                Set bstack.lastobj = myobject
                If TypeOf myobject Is mHandler Then
                    Set usehandler = myobject
                    If usehandler.T1 = 4 Then
                        p = usehandler.index_cursor
                        If MemInt(VarPtr(p)) <> vbString Then p = p * usehandler.sign
                    End If
                End If
                GoTo checkenum
                Else
                GoTo er103
                End If
                Else
                     MyRead = False
                     If VarTypeName(var(I)) = "Nothing" Then
                     MissingObjRef
                     Else
                    GoTo er103
                    End If
                    Exit Function
                    
                End If
cont10101:
        Else
        If I = -1 Then
        ff = 0
        If TypeOf myobject Is mHandler Then
        Set usehandler = myobject
        If jumpAs Then jumpAs = False: GoTo existAs08
        ff = 1
        If Fast2Varl(rest$, "ΩΣ", 2, "AS", 2, 2, ff) Then
existAs08:
                If usehandler.T1 = 1 Then
                    If Not Fast2Varl(rest$, "ΚΑΤΑΣΤΑΣΗ", 9, "INVENTORY", 9, 9, ff) Then
                         If Not Fast2Varl(rest$, "ΛΙΣΤΑ", 5, "LIST", 4, 5, ff) Then
                            If Not Fast2Varl(rest$, "ΟΥΡΑ", 4, "QUEUE", 5, 5, ff) Then
                                WrongObject
                                MyRead = False
                                Exit Function
                            ElseIf Not usehandler.objref.IsQueue Then
                                WrongObject
                                MyRead = False
                                Exit Function
                            End If
                        ElseIf usehandler.objref.IsQueue Then
                            WrongObject
                            MyRead = False
                            Exit Function
                        End If
                    End If
                ElseIf usehandler.T1 = 2 Then
                    If Not Fast2Varl(rest$, "ΔΙΑΡΘΡΩΣΗ", 9, "BUFFER", 6, 9, ff) Then
                            WrongObject
                            MyRead = False
                            Exit Function
                    End If
                ElseIf usehandler.T1 = 3 Then
                    If Not Fast2Varl(rest$, "ΠΙΝΑΚΑΣ", 7, "ARRAY", 5, 7, ff) Then
                    If Not Fast2Varl(rest$, "ΣΩΡΟΣ", 5, "STACK", 5, 5, ff) Then
                    
                            WrongObject
                            MyRead = False
                            Exit Function
                    End If
                    End If
                    
                ElseIf usehandler.T1 = 4 Then
                    If Not FastType(rest$, usehandler.objref.EnumName) Then
                    
                        p = usehandler.index_cursor
                        Set myobject = Nothing
                        Set usehandler = Nothing
                        GoTo conthereEnum
                    End If
                End If
        End If
        Set usehandler = Nothing
        bstack.SetVarobJ what$, myobject
        GoTo loopcont123
        ElseIf TypeOf myobject Is Group Then
        
        If jumpAs Then jumpAs = False: GoTo existAs10
        ff = 1
        If Fast2Varl(rest$, "ΩΣ", 2, "AS", 2, 2, ff) Then
existAs10:
            If Fast2Varl(rest$, "ΔΕΙΚΤΗΣ", 7, "POINTER", 7, 7, ff) Then GoTo checkpointer
            WrongObject
            MyRead = False
            Exit Function
        End If
checkpointer:
        If myobject.IamApointer Then
        If myobject.link.IamFloatGroup Then
            bstack.SetVarobJ what$, myobject
           GoTo loopcont123
        End If
        End If
        End If
        If jumpAs Then GoTo er110
        If FastPureLabel(rest$, ss$, , True) = 1 Then
        If check2(ss$, "ΩΣ", "AS") Then GoTo er110
        End If
        WrongObject
        MyRead = False
        Exit Function
        End If
        
comehere:
        I = globalvar(what$, 0)
        it = varhash.Index
      
        If Typename$(myobject) = mGroup Then
        If jumpAs Then jumpAs = False: GoTo existAs12
        ff = 1
            If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs12:
                If Fast2Varl(rest$, "ΔΕΙΚΤΗΣ", 7, "POINTER", 7, 7, ff) Then
                    If Not myobject.IamApointer Then
errgr:
                          WrongObject
                          MyRead = False
                          Exit Function
                    End If
                    GoTo contpointer
                ElseIf Not Fast2Varl(rest$, "ΟΜΑΔΑ", 5, "GROUP", 5, 5, ff) Then
                       If FastSymbol(rest$, "*") Then
                               If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
                                If Not myobject.IamApointer Then GoTo errgr
                                If myobject.link.IamFloatGroup Then
                                    If Not myobject.link.TypeGroup(s$) Then GoTo errgr
                                Else
                                    If Len(myobject.lasthere) = 0 Then
                                    If GetVar(bstack, myobject.GroupName, it, True) Then
                                        If Not var(it).TypeGroup(s$) Then GoTo errgr
                                    End If
                                    ElseIf GetVar(bstack, myobject.lasthere + "." + myobject.GroupName, it, True) Then
                                            If Not var(it).TypeGroup(s$) Then GoTo errgr
                                    End If
                                    it = 0
                                End If
                            GoTo contpointer
                        ElseIf myobject.IamApointer Then
                            If FastPureLabel(rest$, s$, , True) = 1 Then
                                If myobject.link.IamFloatGroup Then
                                    If Not myobject.link.TypeGroup(s$) Then GoTo errgr
                                    GoTo oop0
                                ElseIf Len(myobject.lasthere) = 0 Then
                                    If GetVar(bstack, myobject.GroupName, I, True) Then
                                        If Not var(I).TypeGroup(s$) Then GoTo errgr
                                        CopyGroup2 var(I), bstack
                                        GoTo oop1
                                    Else
                                        GoTo errgr
                                    End If
                                ElseIf GetVar(bstack, myobject.lasthere + "." + myobject.GroupName, I, True) Then
                                    If Not var(I).TypeGroup(s$) Then GoTo errgr
                                    CopyGroup2 var(I), bstack
                                    GoTo oop1
                                Else
                                    GoTo errgr
                                End If
                                
                            Else
                                SyntaxError
                                MyRead = False
                                Exit Function
                            End If
                        ElseIf FastPureLabel(rest$, s$, , True) <> 1 Then
                              GoTo errgr
                        Else
                              If Not myobject.TypeGroup(s$) Then GoTo errgr
                        End If
                    
                    End If
                    If myobject.IamApointer Then
                        If myobject.link.IamFloatGroup Then
oop0:
                            If myobject.link Is NullGroup Then
                                Set myobject = New Group
                                myobject.BeginFloat 0
                                myobject.EndFloat
                                UnFloatGroup bstack, bstack.GroupName + what$, I, myobject, here$ = vbNullString Or Len(bstack.UseGroupname) > 0
                            Else
                                UnFloatGroup bstack, bstack.GroupName + what$, I, myobject.link, here$ = vbNullString Or Len(bstack.UseGroupname) > 0
                            End If
                        Else
                            CopyPointerRef bstack, myobject
oop1:
                            UnFloatGroup bstack, bstack.GroupName + what$, I, bstack.lastobj, here$ = vbNullString Or Len(bstack.UseGroupname) > 0
                            bstack.lastobj.ToDelete = True
                        End If
                    Else
                        UnFloatGroup bstack, bstack.GroupName + what$, I, myobject, here$ = vbNullString Or Len(bstack.UseGroupname) > 0, , True
                        myobject.ToDelete = True
                    End If
                    
                Else
contpointer:
                    If myobject.IamApointer Then
                        Set var(I) = myobject
                    Else
                        UnFloatGroup bstack, bstack.GroupName + what$, I, myobject, Not (here$ = vbNullString Xor Len(bstack.UseGroupname) > 0), , True
                        myobject.ToDelete = True
                    End If
               End If
               ' var(i).IamRef = Len(bstack.UseGroupname) > 0
            ElseIf Typename$(myobject) = "mEvent" Then
                If jumpAs Then jumpAs = False: GoTo existAs121
                ff = 1
                If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs121:
                If Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then
jump0001233:
                        
                        varhash.vType(it) = False
                        If FastSymbol(rest$, "=") Then
                            optlocal = Not useoptionals: useoptionals = True
                            If Not IsNumberD2(rest$, (p), False) Then
                            If Not ISSTRINGA(rest$, s$) Then
                                SyntaxError
                                Exit Function
                            End If
                            End If
                        End If
                ElseIf Not Fast2Varl(rest$, "ΓΕΓΟΝΟΣ", 7, "EVENT", 5, 5, ff) Then
                        WrongObject
                        Exit Function
                    End If
                End If
                Set var(I) = myobject
            ElseIf Typename$(myobject) = "lambda" Then
                    If jumpAs Then jumpAs = False: GoTo existAs13
                    ff = 1
                    If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs13:
                        If Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then
                           If FastSymbol(rest$, "=") Then
                            optlocal = Not useoptionals: useoptionals = True
                            If Not IsNumberD2(rest$, (p), False) Then
                            If Not ISSTRINGA(rest$, s$) Then
                                SyntaxError
                                Exit Function
                            End If
                            End If
                        End If
                    ElseIf Not Fast2Varl(rest$, "ΛΑΜΔΑ", 5, "LAMBDA", 6, 6, ff) Then
                        WrongObject
                        Exit Function
                    End If
                End If
                Set var(I) = myobject
                If ohere$ = vbNullString Then
                    GlobalSub what$ + "()", "", , , I
                Else
                    GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , I
                End If
            ElseIf Typename$(myobject) = mHdlr Then
                Set usehandler1 = myobject
                If MyIsObject(var(I)) Then
                    If TypeOf var(I) Is mHandler Then
                        Set usehandler = var(I)
                        If usehandler.ReadOnly Then
                            ReadOnly
                            MyRead = False
                            Exit Function
                        End If
                    End If
                End If
                If usehandler1.indirect > -1 Then
                If jumpAs Then jumpAs = False: GoTo existAs14
                ff = 1
                    If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs14:
                        If IsLabel(bstack, rest$, ss$) = 0 Then
                            GoTo er110
                        End If
                        If LCase(Typename(var(usehandler1.indirect))) <> LCase(ss$) Then
                            WrongObject
                            MyRead = False
                            Exit Function
                        End If
                    End If
                    Set var(I) = MakeitObjectGeneric(usehandler1.indirect)
                Else
                 If jumpAs Then jumpAs = False: GoTo existAs15
                 ff = 1
                    If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs15:
                        If usehandler1.T1 = 1 Then
                            If Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then
                                GoTo jump0001233
                            ElseIf Not Fast2Varl(rest$, "ΚΑΤΑΣΤΑΣΗ", 9, "INVENTORY", 9, 9, ff) Then
                                If Not Fast2Varl(rest$, "ΛΙΣΤΑ", 5, "LIST", 4, 5, ff) Then
                                    If Not Fast2Varl(rest$, "ΟΥΡΑ", 4, "QUEUE", 5, 5, ff) Then
                                        WrongObject
                                        MyRead = False
                                        Exit Function
                                    ElseIf Not usehandler1.objref.IsQueue Then
                                        WrongObject
                                        MyRead = False
                                        Exit Function
                                    End If
                                ElseIf usehandler1.objref.IsQueue Then
                                    WrongObject
                                    MyRead = False
                                    Exit Function
                                End If
                            End If
                        ElseIf usehandler1.T1 = 2 Then
                            If Not Fast2Varl(rest$, "ΔΙΑΡΘΡΩΣΗ", 9, "BUFFER", 6, 9, ff) Then
                                If Not Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then
                                If FastPureLabel(rest$, s$, , True, , , , True) = 1 Then
                                
                                If usehandler1.objref.UseStruct Then
                                If usehandler1.objref.structref.Tag = s$ Then
                                
                                GoTo t14
                                
                                End If
                                End If
                                End If
                                    WrongObject
                                    MyRead = False
                                    Exit Function
                                Else
                                
                                    GoTo jump0001233
                                End If
                            End If
                        ElseIf usehandler1.T1 = 3 Then
                            If Fast2Varl(rest$, "ΠΙΝΑΚΑΣ", 7, "ARRAY", 5, 7, ff) Then
                                If Not CheckIsmArray(myobject) Then GoTo er103
                                Set usehandler = New mHandler
                                Set usehandler.objref = myobject
                                usehandler.T1 = 3
                                Set myobject = usehandler
                                Set usehandler = Nothing
                                Set usehandler1 = Nothing
                            ElseIf Fast2Varl(rest$, "ΣΩΡΟΣ", 5, "STACK", 5, 5, ff) Then
                                If Not CheckIsmStiva(myobject) Then GoTo er103
                                Set usehandler = New mHandler
                                Set usehandler.objref = myobject
                                usehandler.T1 = 3
                                Set myobject = usehandler
                                Set usehandler = Nothing
                                Set usehandler1 = Nothing
                            ElseIf Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then
                                If Not CheckIsmArray(myobject) Then
                                    If Not CheckIsmStiva(myobject) Then
                                        GoTo er103
                                    End If
                                End If
                                Set usehandler = New mHandler
                                Set usehandler.objref = myobject
                                usehandler.T1 = 3
                                Set myobject = usehandler
                                Set usehandler = Nothing
                                Set usehandler1 = Nothing
                                GoTo jump0001233
                            ElseIf Typename(usehandler1.objref) = mHdlr Then
                                Set usehandler1 = usehandler1.objref
                                If usehandler1.T1 = 4 Then
                                    If FastType(rest$, usehandler1.objref.EnumName) Then
                                        Set usehandler = New mHandler
                                        usehandler.T1 = 4
                                        Set myobject = usehandler1.objref
                                        usehandler.index_cursor = myobject.Value
                                        Set usehandler.objref = myobject
                                        usehandler.index_start = myobject.Index
                                        usehandler.sign = 1
                                        Set myobject = usehandler
                                        GoTo t14
                                    End If
                                End If
                            Else
                                WrongObject
                                MyRead = False
                                Exit Function
                            End If
                        ElseIf usehandler1.T1 = 4 Then
                            If FastPureLabel(rest$, s$, , True, , , False) = 1 Then
                                If Not s$ = myUcase(usehandler1.objref.EnumName, True) Then
                                    
                                    If GetSub(s$ + "()", I) Then
                                        If sbf(I).IamAClass Then
                                            GoTo er113
                                        End If
                                    ElseIf GetSub(bstack.GroupName + s$ + "()", I) Then
                                        If sbf(I).IamAClass Then
                                            GoTo er113
                                        End If
                                    End If
                                    ' no error yet
                                    If MyIsObject(usehandler1.index_cursor) Then
                                        bs.soros.PushObj usehandler1.index_cursor
                                    ElseIf myVarType(usehandler1.index_cursor, vbString) Then
                                        bs.soros.PushStrVariant usehandler1.index_cursor
                                        
                                    Else
                                        bs.soros.PushVal usehandler1.index_cursor * usehandler1.sign
                                    End If
                                    Set usehandler1 = Nothing
                                    jumpAs = True
                                    GoTo fromEnumDeref
                                    
                                Else
                                    FastPureLabel rest$, s$
                                End If
                            Else
                                GoTo er112
                            End If
                        If FastSymbol(rest$, "=") Then
                            ' drop type
                            If FastPureLabel(rest$, s$, , , True) <> 1 Then
                                MyRead = False
                                SyntaxError
                                Exit Function
                            End If
                        End If
                    End If
                ElseIf usehandler1.T1 = 4 Then
                    If FastSymbol(rest$, "=") Then
                    If IsEnumLabelOnly(bstack, rest$) Then
                        Set usehandler = bstack.lastobj
                        Set bstack.lastobj = Nothing
                        If usehandler1.objref.EnumName <> usehandler.objref.EnumName Then
                        MyEr "expected enum type " + usehandler.objref.EnumName, "περίμενα τύπο απαριθμητή " + usehandler.objref.EnumName
                        MyRead = False
                        Exit Function
                        End If
                    End If
                    End If
                End If
t14:
                Set var(I) = myobject
                End If
            ElseIf Typename$(myobject) = myArray Then
                   If jumpAs Then jumpAs = False: GoTo existAs16
                   ff = 1
                   If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs16:
                   If Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then
                    ElseIf Not Fast2Varl(rest$, "ΠΙΝΑΚΑΣ", 7, "ARRAY", 5, 7, ff) Then
                        WrongObject
                        Exit Function
                    End If
            End If
                Set usehandler = New mHandler
                Set var(I) = usehandler
                usehandler.T1 = 3
                Set usehandler.objref = myobject
                Set usehandler = Nothing
            Else
            If jumpAs Then jumpAs = False: GoTo existAs17
            ff = 1
             If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs17:
                    If Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then
                        GoTo jump0001233
                    End If
                    If Not Fast2Varl(rest$, "ΔΕΙΚΤΗΣ", 7, "POINTER", 7, 7, ff) Then
                    
                    
                         If Fast2Varl(rest$, "ΜΕΓΑΛΟΣΑΚΕΡΑΙΟΣ", 15, "BIGINTEGER", 10, 15, ff) Then
                         ss$ = "BigInteger"
                         If LCase(Typename(myobject)) = LCase(ss$) Then
                            Set myobject = CopyBigInteger(myobject)
                         End If
                          If FastSymbol(rest$, "=") Then
                            optlocal = Not useoptionals: useoptionals = True
                            If Not IsNumberD2(rest$, (p), False) Then
                            If Not ISSTRINGA(rest$, s$) Then
                                SyntaxError
                                Exit Function
                            End If
                            End If
                        End If
                        ElseIf FastSymbol(rest$, "*") Then
                            If IsLabel(bstack, rest$, ss$) = 0 Then
                                GoTo er110
                            End If
                            If LCase(ss$) = "long" Then
                                If IsLabel(bstack, rest$, ss$) = 1 Then
                                If LCase(ss$) = "long" Then
                                    ss$ = "long long"
                                    
                                Else
                                    SyntaxError
                                    MyRead = False
                                    Exit Function
                                End If
                                Else
                                ss$ = "long"
                                End If
                            End If
                            If Typename(myobject) = "RefArray" Then
                                Set Ar = myobject
                                If LCase(VarTypeName(Ar(0, 0))) <> LCase(ss$) Then
                                    If Ar.vtType(0) = vbVariant And LCase(ss$) = "variant" Then
                                    
                                    Else
                                    GoTo er103
                                    End If
                                End If
                                Set Ar = Nothing
                                GoTo contsethere
                                
                            End If
                        End If
                    ElseIf IsLabel(bstack, rest$, ss$) = 0 Then
                        GoTo er110
                    End If
                
                If LCase(Typename(myobject)) <> LCase(ss$) Then
                        WrongObject
                        MyRead = False
                        Exit Function
                End If
            End If
contsethere:
                Set var(I) = myobject
            
            End If
            
            
            
            
            Set myobject = Nothing
        End If
    ElseIf bs.IsNumber(p) Then
contStr1:
    ihavetype = True
contStr2:
    If Not lookOne(rest$, ",") Then
    ' FF used again
        If jumpAs Then jumpAs = False: GoTo conthereEnum
        ff = 1
        If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
conthereEnum:
            ihavetype = True
            If Not FastPureLabel(rest$, s$, , , True) = 1 Then
            SyntaxError
            Exit Function
            End If
            On Error GoTo er1234
            ss$ = myUcase(s$, AscW(s$) > 255)
            Select Case ss$
            Case "ΑΡΙΘΜΟΣ", "DECIMAL"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                it = vbDecimal
                p = CDec(p)
            Case "ΔΙΠΛΟΣ", "DOUBLE"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                it = vbDouble
                p = CDbl(p)
            Case "ΑΠΛΟΣ", "SINGLE"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                it = vbSingle
                p = CSng(p)
            Case "ΛΟΓΙΚΟΣ", "BOOLEAN"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                it = vbBoolean
                p = CBool(p)
            Case "ΜΑΚΡΥΣ", "LONG"
                If Fast2Varl(rest$, "ΜΑΚΡΥΣ", 6, "LONG", 4, 6, ff) Then
                    If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                    it = 20
                    p = cInt64(p)
                Else
                    If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                    it = vbLong
                    p = CLng(p)
                End If
            Case "ΑΚΕΡΑΙΟΣ", "INTEGER"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                it = vbInteger
                p = CInt(p)
            Case "ΛΟΓΙΣΤΙΚΟΣ", "CURRENCY"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                p = CCur(p)
            Case "ΓΡΑΜΜΑ", "STRING"
                If FastSymbol(rest$, "=") Then
                optlocal = Not useoptionals
                useoptionals = True
                If Not ISSTRINGA(rest$, (s$)) Then
                        MissString
                        Exit Function
                End If
                End If
                If VarType(p) <> vbString Then
                        MissString
                        Exit Function
                End If
            Case "ΜΙΓΑΔΙΚΟΣ", "COMPLEX"
                If FastSymbol(rest$, "=") Then
                    optlocal = Not useoptionals
                    useoptionals = True
                    If Not FastSymbol(rest$, "(") Then missNumber: Exit Function
                    If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                    'X = CDbl(p)
                    If Not FastSymbol(rest$, ",") Then missNumber: Exit Function
                    If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                    rest$ = NLtrim$(rest$)
                    If Len(rest$) >= 2 Then
                    If Not UCase(Left$(rest$, 2)) = "I)" Then Mid$(rest$, 1, 2) = "  ": missNumber: Exit Function
                        Mid$(rest$, 1, 2) = "  "
                    Else
                        missNumber
                        Exit Function
                    End If

                    End If
                    'p = nMath2.cxNew(X, CDbl(p))
            Case "ΑΤΥΠΟΣ", "VARIANT"
                If FastSymbol(rest$, "=") Then
                optlocal = Not useoptionals: useoptionals = True
                If Not IsNumberD2(rest$, (p), False) Then
                If Not ISSTRINGA(rest$, s$) Then
                    SyntaxError
                    Exit Function
                End If
                End If
                End If
                ihavetype = False
            Case "ΜΕΓΑΛΟΣΑΚΕΡΑΙΟΣ", "BIGINTEGER"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                
                Set p = Module13.CreateBigInteger(CStr(CDec(p)))
            Case "ΨΗΦΙΟ", "BYTE"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                p = CByte(p)
            Case "ΗΜΕΡΟΜΗΝΙΑ", "DATE"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                p = CDate(p)
            Case Else
                If MemInt(VarPtr(p)) = 36 Then
                    If Not LCase(Typename(p)) = LCase(s$) Then
                    GoTo messnotype
                    End If
                Else
                    ss$ = s$
                    it = True
                  If MyIsNumeric(p) Then X = p: it = False
                   
                  If IsEnumAs(bstack, ss$, p, ok, rest$) Then
                    If Not it Then
                        Set usehandler = p
                        p = X
                        Set usehandler = usehandler.objref.SearchValue(p, ok)
                        Set myobject = usehandler
                        If ok Then
                            Set p = myobject
                        Else
                            ExpectedEnumType
                            MyRead = False
                            Exit Function
                        End If
                    
                    End If
                    
                  Else
                  
messnotype:
                  MyEr "No type [" + s$ + "] found", "δεν βρήκα τύπο [" + s$ + "]"
                  Exit Function
                  End If
                  End If
            End Select
        ElseIf FastSymbol(rest$, "=") Then
            If Not IsNumberD2(rest$, (p)) Then
            If IsEnumLabelOnly(bstack, rest$) Then
             
            
            Set usehandler = bstack.lastobj
             Set bstack.lastobj = Nothing
            Set p = usehandler.objref.SearchValue(p, ok)
            If ok Then
                  GoTo contenumok
            End If
             'usehandler.objref.SearchValue
             ExpectedEnumType
             Exit Function
            Else
                missNumber
                Exit Function
            End If
            End If
        ElseIf FastSymbol(rest$, "|") Then
            If GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True, ok) Then
                If MemInt(VarPtr(var(I))) = 36 Then
                    If placevalue2(bstack, var(I), rest$, p) Then
                    MyRead = True
                       GoTo loopcont123
                    End If
                End If
                WrongType
                SyntaxError
            
                Exit Function
            End If
        End If
    End If
contenumok:
        MyRead = True
        If flag2 Then
            globalvar what$, p
        ElseIf GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True, ok) Then
            ihavetype = False
            If isAglobal And Not allowglobals Then
                 globalvar what$, p
            ElseIf MyIsObject(var(I)) Then
                If var(I) Is Nothing Then
                    MissingObjRef
                ElseIf TypeOf var(I) Is Group Then
checkenum:
                    If var(I).HasSet Then
                       
                        Set M = bstack.soros
                        Set bstack.Sorosref = New mStiva
                        bstack.soros.PushVal p
                        NeoCall2 bstack, what$ + "." + ChrW(&H1FFF) + ":=()", ok
                        Set bstack.Sorosref = M
                        Set M = Nothing
                    Else
                        GoTo there18274
                    End If
                Else
there18274:
                    If TypeOf var(I) Is Constant Then
                        CantAssignValue
                    ElseIf TypeOf var(I) Is mHandler Then
                        Set usehandler = var(I)
                        If usehandler.T1 <> 4 Then GoTo er104
                        If MyIsObject(p) Then
                            If Not TypeOf p Is mHandler Then
                                    ' errorhere
                                End If
                                Set usehandler1 = p
                                If usehandler.objref Is usehandler1.objref Then
                                    Set myobject = usehandler1
                                Else
                                    p = Empty
                                    If Not usehandler1.T1 = 4 Then
                                    'error here
                                    End If
                                    p = usehandler1.index_cursor
                                    Set myobject = usehandler.objref.SearchValue(p, ok)
                                End If
                            Else
                                Set myobject = usehandler.objref.SearchValue(p, ok)
                            End If
                            Set usehandler1 = Nothing
                                If ok Then
                                    Set var(I) = myobject
                                    GoTo cont112233
                                Else
                                    GoTo er112
                                End If
                            Else
                                GoTo er104
                            End If
                            Exit Function
                            End If
                        Else
                            If checktype Then
                                If ihavetype Then
                                    If VarType(var(I)) <> VarType(p) Then
                                        GoTo er109
                                    ElseIf AssignTypeNumeric(p, VarType(var(I))) Then
                                        var(I) = p
                                    Else
                                        GoTo er105
                                    End If
                                ElseIf AssignTypeNumeric(p, VarType(var(I))) Then
                                    var(I) = p
                                Else
                                    GoTo er105
                                End If
                            Else
                                var(I) = p
                            End If
                        End If
                    ElseIf I = -1 Then
                        If ok Then
                           bstack.SetVarobJvalue what$, p
                        Else
                            bstack.SetVar what$, p
                        End If
                    Else
                        globalvar what$, p, useType:=ihavetype
                    End If
cont112233:
                    p = 0#
                ElseIf bs.IsOptional Then
                    MyRead = True
                    If Not lookOne(rest$, ",") Then
                        checktype = True
                        If jumpAs Then jumpAs = False: GoTo existAs18
                        ff = 1
                        If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs18:
                            ihavetype = True
                            If Not FastPureLabel(rest$, s$, , , True) = 1 Then
                                If FastSymbol(rest$, "*") Then
                                    If FastPureLabel(rest$, ss$, , True) = 1 Then
                                        ' pointer
                                        GoTo cont234356
                                    End If
                                End If
                                SyntaxError
                                Exit Function
                            End If
                            ss$ = myUcase(s$, AscW(s$) > 255)
                        Select Case ss$
                        Case "ΑΡΙΘΜΟΣ", "DECIMAL"
                                p = CDec(0)
                        Case "ΔΙΠΛΟΣ", "DOUBLE"
                                p = 0#
                        Case "ΑΠΛΟΣ", "SINGLE"
                                p = 0!
                        Case "ΛΟΓΙΚΟΣ", "BOOLEAN"
                                p = False
                        Case "ΜΑΚΡΥΣ", "LONG"
                            If FastPureLabel(rest$, s$, , , True) = 1 Then
                                If ss$ = myUcase(s$, AscW(s$) > 255) Then
                                    p = cInt64(0)
                                Else
                                    UknownType ss$ + " " + s$
                                    MyRead = False
                                    Exit Function
                                End If
                            Else
                                p = 0&
                            End If
                                
                        Case "ΑΚΕΡΑΙΟΣ", "INTEGER"
                                p = 0
                        Case "ΛΟΓΙΣΤΙΚΟΣ", "CURRENCY"
                            p = 0@
                        Case "ΓΡΑΜΜΑ", "STRING"
                            p = vbNullString
                        Case "ΑΤΥΠΟΣ", "VARIANT"
                            p = Empty
                            ihavetype = False
                            checktype = False
                        Case "ΜΙΓΑΔΙΚΟΣ", "COMPLEX"
                            p = nMath2.cxZero
                        Case "ΨΗΦΙΟ", "BYTE"
                            p = CByte(0)
                        Case "ΗΜΕΡΟΜΗΝΙΑ", "DATE"
                            p = CDate(0#)
                        Case Else
cont234356:
                        
                            If Not flag2 And GetVar3(bstack, what$, I, , , flag, , , isAglobal, True, ok) Then
                                If isAglobal Then
                                    GoTo er110
                                Else
                                    If MyIsNumeric(var(I)) Then
                                        If IsEnumAs(bstack, s$, p, ok, rest$) Then
                                            If ok Then Set var(I) = p
                                        End If
                                    ElseIf TypeOf var(I) Is Group Then
                                        If Not ss$ = "ΟΜΑΔΑ" Then
                                            If Not ss$ = "GROUP" Then
                                                If var(I).IamApointer Then
                                                    If Not ss$ = "ΔΕΙΚΤΗΣ" Then
                                                        If Not ss$ = "POINTER" Then
                                                            If Not var(I).link.TypeGroup(ss$) Then
                                                                MyRead = False: MissType: Exit Function
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    MyRead = False: MissType: Exit Function
                                                End If
                                            End If
                                        End If
                                    ElseIf TypeOf var(I) Is mHandler Then
                                        Set usehandler = var(I)
                                        If TypeOf usehandler.objref Is mArray Then
                                            If Not ss$ = "ΠΙΝΑΚΑΣ" Then
                                                If Not ss$ = "ARRAY" Then
                                                    MyRead = False: MissType: Exit Function
                                                End If
                                            End If
                                        ElseIf TypeOf usehandler.objref Is FastCollection Then
                                            If Not ss$ = "ΚΑΤΑΣΤΑΣΗ" Then
                                                If Not ss$ = "INVENTORY" Then
                                                    If Not ss$ = "ΛΙΣΤΑ" Then
                                                        If Not ss$ = "LIST" Then
                                                            If Not ss$ = "ΟΥΡΑ" Then
                                                                If Not ss$ = "QUEUE" Then
                                                                    MyRead = False: MissType: Exit Function
                                                                End If
                                                            End If
                                                            If Not usehandler.objref.IsQueue Then
                                                                MyRead = False: MissType: Exit Function
                                                            End If
                                                        Else
                                                            GoTo islist
                                                        End If
islist:
                                                        If usehandler.objref.IsQueue Then
                                                            MyRead = False: MissType: Exit Function
                                                            Exit Function
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        ElseIf TypeOf usehandler.objref Is MemBlock Then
                                            If Not ss$ = "ΔΙΑΡΘΡΩΣΗ" Then
                                                If Not ss$ = "BUFFER" Then
                                                    MyRead = False: MissType: Exit Function
                                                End If
                                            End If
                                        ElseIf TypeOf usehandler.objref Is Enumeration Then
                                            If usehandler.objref.EnumName = s$ Then
                                                If FastSymbol(rest$, "=") Then
                                                    If FastPureLabel(rest$, ss$, , , True) <> 1 Then
                                                        MyRead = False: SyntaxError: Exit Function
                                                    End If
                                                Else
                                                   MyRead = False: SyntaxError: Exit Function
                                                End If
                                            Else
                                                MyRead = False: MissType: Exit Function
                                            End If
                                            Set usehandler = var(I)
                                        Else
                                            p = IsLabel(bstack, rest$, (what$)) ' just throw any name
                                        End If
                                        Set usehandler = Nothing
                                    ElseIf TypeOf var(I) Is lambda Then
                                            If Not ss$ = "ΛΑΜΔΑ" Then
                                                If Not ss$ = "LAMBDA" Then
                                                    MyRead = False: MissType: Exit Function
                                                End If
                                            End If
                                    ElseIf TypeOf var(I) Is mEvent Then
                                            If Not ss$ = "ΓΕΓΟΝΟΣ" Then
                                                If Not ss$ = "EVENT" Then
                                                    MyRead = False: MissType: Exit Function
                                                End If
                                            End If
                                    ElseIf TypeOf var(I) Is Constant Then
                                        p = var(I)
                                        GoTo checkconstant
                                    ElseIf IsEnumAs(bstack, (s$), p, ok, rest$) Then
                                        If ok Then
                                            CheckItemType bstack, var(I), vbNullString, ss$, , ok
                                            If s$ = ss$ Then
                                                Set var(I) = p
                                            End If
                                        End If
                                       ' just throw any name
                                    End If
                                    GoTo loopcont123
                                End If
                                ' never pass from here
                            Else
                                If I = -1 Then
                                    If Len(ss$) = 0 Then
                                        FastPureLabel rest$, s$, , , True
                                        
                                        ss$ = myUcase(s$, True)
                                    End If
                                    bstack.ReadVar what$, p
                                    Set myobject = p
                                    Set usehandler = myobject
                                    ff = 0
                                    If Not check2(ss$, "ΔΕΙΚΤΗΣ", "POINTER") Then
                                        If usehandler.T1 = 1 Then
                                            If Not check2(ss$, "ΚΑΤΑΣΤΑΣΗ", "INVENTORY") Then
                                                If Not check2(ss$, "ΛΙΣΤΑ", "LIST") Then
                                                    If Not check2(ss$, "ΟΥΡΑ", "QUEUE") Then
                                                        WrongObject
                                                        MyRead = False
                                                        Exit Function
                                                    ElseIf Not usehandler.objref.IsQueue Then
                                                        WrongObject
                                                        MyRead = False
                                                        Exit Function
                                                    End If
                                                ElseIf usehandler.objref.IsQueue Then
                                                    WrongObject
                                                    MyRead = False
                                                    Exit Function
                                                End If
                                            End If
                                        ElseIf usehandler.T1 = 2 Then
                                            If Not check2(ss$, "ΔΙΑΡΘΡΩΣΗ", "BUFFER") Then
                                                WrongObject
                                                MyRead = False
                                                Exit Function
                                            End If
                                        ElseIf usehandler.T1 = 3 Then
                                            If Not check2(ss$, "ΠΙΝΑΚΑΣ", "ARRAY") Then
                                                If Not check2(s$, "ΣΩΡΟΣ", "STACK") Then
                                                    WrongObject
                                                    MyRead = False
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                    ElseIf usehandler.T1 = 4 Then
                                        If Not FastType(s$, usehandler.objref.EnumName) Then
                                            p = usehandler.index_cursor
                                            Set usehandler = Nothing
                                            Set myobject = Nothing
                                            GoTo conthereEnum
                                        End If
                                    End If
                                    GoTo loopcont123
                                ElseIf check2(ss$, "ΟΜΑΔΑ", "GROUP") Then
                                useoptionals = False
                                ElseIf check2(ss$, "ΠΙΝΑΚΑΣ", "ARRAY") Then
                                useoptionals = False
                                ElseIf check2(ss$, "ΚΑΤΑΣΤΑΣΗ", "INVENTORY") Then
                                useoptionals = False
                                ElseIf check2(ss$, "ΔΙΑΡΘΡΩΣΗ", "BUFFER") Then
                                useoptionals = False
                                ElseIf check2(ss$, "ΛΑΜΔΑ", "LAMBDA") Then
                                useoptionals = False
                                ElseIf check2(ss$, "ΓΕΓΟΝΟΣ", "EVENT") Then
                                    useoptionals = False
                                ElseIf check2(ss$, "ΛΙΣΤΑ", "LIST") Then
                                    useoptionals = False
                                ElseIf check2(ss$, "ΟΥΡΑ", "QUEUE") Then
                                    useoptionals = False
                                ElseIf check2(ss$, "ΜΕΓΑΛΟΣΑΚΕΡΑΙΟΣ", "BIGINTEGER") Then
                                    useoptionals = False
                                    
                                    If FastSymbol(rest$, "=") Then
                                        Set p = New BigInteger
                                        If IsNumberD2(rest$, p, True, True) Then
                                             
                                             Set p = Module13.CreateBigInteger(CStr(p))

                                        Else
                                            missNumber
                                            Exit Function
                                        End If
                                    Else
                                        Set p = New BigInteger
                                    End If
                                    optlocal = Not useoptionals: useoptionals = True
                                    GoTo A038340
                                    
                                ElseIf IsEnumAs(bstack, s$, p, ok, rest$) Then
                                    If ok Then
                                        optlocal = Not useoptionals: useoptionals = True
                                    Else
                                        GoTo err10
                                    End If
                                Else
                                    GoTo er110
                                End If
                            End If
                            GoTo cont1459
                        End Select

                        If FastSymbol(rest$, "=") Then
                            it = VarType(p)
                            If it = 36 Then
                                If TypeOf p Is cxComplex Then
                                    
                                If Not FastSymbol(rest$, "(") Then missNumber: Exit Function
                                If Not IsNumberD2(rest$, p) Then missNumber: Exit Function
                                X = CDbl(p)
                                If Not FastSymbol(rest$, ",") Then missNumber: Exit Function
                                If Not IsNumberD2(rest$, p) Then missNumber: Exit Function
                                If Len(rest$) >= 2 Then
                                    If Not UCase(Left$(rest$, 2)) = "I)" Then Mid$(rest$, 1, 2) = "  ": missNumber: Exit Function
                                    Mid$(rest$, 1, 2) = "  "
                                End If

                                p = nMath2.cxNew(X, CDbl(p))

                                Else
                                    SyntaxError
                                    Exit Function
                                End If
                            ElseIf Not IsNumber(bstack, rest$, p, True) Then
                            If Not ihavetype Then
                                            If ISSTRINGA(rest$, s$) Then
                                                p = s$
                                            Else
                                                MissStackStr
                                                Exit Function
                                            End If
                            ElseIf VarType(p) = vbString Then
                                        If ISSTRINGA(rest$, s$) Then
                                                p = s$
                                            Else
                                                MissStackStr
                                                Exit Function
                                            End If
                                ElseIf VarType(p) = vbDate Then
                                        If ISSTRINGA(rest$, s$) Then
                                                p = CDate(s$)
                                            Else
                                                MissStackStr
                                                Exit Function
                                            End If
                            Else
                                missNumber
                                Exit Function
                            End If
                            Else
                            If VarType(p) <> it Then
                            If ihavetype Then
                            On Error GoTo er1234
                            Select Case it
                            Case vbBoolean
                            p = CBool(p)
                            Case vbInteger
                            p = CInt(p)
                            Case vbLong
                            p = CLng(p)
                            Case 20
                            p = cInt64(p)
                            Case vbDouble
                            p = CDbl(p)
                            Case vbSingle
                            p = CSng(p)
                            Case vbCurrency
                            p = CCur(p)
                            Case vbDecimal
                            p = CDec(p)
                            Case vbDate
                            p = CDate(p)
                            End Select
                            
                            End If
                            End If
                            End If
                            optlocal = Not useoptionals: useoptionals = True
                        End If
A038340:
                        If Len(rest$) > 0 Then
                            If InStr("!@#%~&", Left$(rest$, 1)) > 0 Then
                                Mid$(rest$, 1, 1) = " "
                            End If
                        End If
                    ElseIf FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p) Then
                            If ISSTRINGA(rest$, ss$) Then
                                p = ss$
                                GoTo optOk
                            ElseIf IsEnumLabelOnly(bstack, rest$) Then
                                Set p = bstack.lastobj
                                Set bstack.lastobj = Nothing
                                GoTo optOk
                            ' here enum
                            End If
                            missNumber
                            Exit Function
                        End If
optOk:
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                Else
                    p = 0#
                End If
                    
cont1459:
                If flag2 Then
                    globalvar what$, p
                    If Not useoptionals Then GoTo err100
                ElseIf GetVar3(bstack, what$, I, , , flag, , , isAglobal, True, ok) Then
                    If isAglobal Then
                        If Not useoptionals Then GoTo err10
                        globalvar what$, p
                    ElseIf ihavetype Then
                        If VarType(p) <> VarType(var(I)) Then
                            If Not AssignTypeNumeric(var(I), VarType(p)) Then
                            MyRead = False
                            Exit Function
                        End If
                    End If
                End If
                ' just skip read for this value
            ElseIf I = -1 Then
                If Not ok Then
                    bstack.SetVar what$, p
                End If
                If Not useoptionals Then GoTo err100
            Else
                If VarType(p) = vbEmpty Then p = 0#
                globalvar what$, p, useType:=checktype
                 checktype = False
                 If Not useoptionals Then GoTo err10
            End If
            p = 0#
        Else
            If bs.IsString(s$) Then
                p = s$
               If jumpAs Then jumpAs = False: GoTo existAs05
                ff = 1
               If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs05:
                    If Not Fast2Varl(rest$, "ΑΤΥΠΟΣ", 6, "VARIANT", 7, 7, ff) Then
                        If Not Fast2Varl(rest$, "ΓΡΑΜΜΑ", 6, "STRING", 6, 6, ff) Then
                            If Not FastPureLabel(rest$, ss$, , , True) = 1 Then
                                SyntaxError
                                Exit Function
                            ElseIf Not IsEnumAs(bstack, ss$, p, ok, rest$) Then
                                    GoTo er110
                            ElseIf ok Then
                                Set usehandler = p
                                Set usehandler = usehandler.objref.SearchValue(CVar(s$), ok)
                                If ok Then
                                    Set p = usehandler
                                    Set usehandler = Nothing
                                Else
                                    Expected "enumeration", "απαρίθμηση"
                                    Exit Function
                                End If
                                GoTo contenumok
                            End If
                        Else
                            If FastSymbol(rest$, "=") Then
                                If Not ISSTRINGA(rest$, s$) Then
                                    SyntaxError
                                    Exit Function
                                End If
                            End If
                        End If
                    ElseIf FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), False) Then
                            If Not ISSTRINGA(rest$, s$) Then
                                SyntaxError
                                Exit Function
                            End If
                        End If
                      GoTo contStr2
                    Else
                    GoTo contStr2
                    End If
                End If
                GoTo contStr1
            End If
            bstack.soros.drop 1
            MissStackNumber
            MyRead = False
            Exit Do
        End If
Case 3
    If bs.IsString(s$) Then
        If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not ISSTRINGA(rest$, (s$)) Then GoTo er111
contstrhere:
        MyRead = True
        If flag2 Then
            globalvar what$, s$
        ElseIf FastSymbol(rest$, "|") Then
            If GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True, ok) Then
                If MemInt(VarPtr(var(I))) = 36 Then
                    p = ""
                    SwapString2Variant s$, p
                    If placevalue2(bstack, var(I), rest$, p) Then
                       GoTo loopcont123
                    End If
                End If
                WrongType
                SyntaxError
            
                Exit Function
            End If
        ElseIf GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True, ok) Then
            If isAglobal And Not allowglobals Then
                globalvar what$, s$
            ElseIf MyIsObject(var(I)) Then
                If TypeOf var(I) Is Group Then
                        Set M = bstack.soros
                        Set bstack.Sorosref = New mStiva
                        bstack.soros.PushStr s$
                        NeoCall2 bstack, Left$(what$, Len(what$) - 1) + "." + ChrW(&H1FFF) + ":=()", ok
                        Set bstack.Sorosref = M
                        Set M = Nothing
                ElseIf TypeOf var(I) Is Constant Then
                CantAssignValue
                MyRead = False
                Exit Function
                Else
                    CheckVar var(I), s$
                End If
            Else
                var(I) = s$
            End If
        ElseIf I = -1 Then
            bstack.SetVar what$, s$
        Else
            globalvar what$, s$
        End If
    ElseIf bs.IsOptional Then
    s$ = vbNullString
       If FastSymbol(rest$, "=") Then
    If Not ISSTRINGA(rest$, s$) Then GoTo er111
                optlocal = Not useoptionals: useoptionals = True
    End If
     
       MyRead = True
       
        If flag2 Then
            If Not useoptionals Then GoTo err100
            globalvar what$, s$
        ElseIf GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True) Then
            If isAglobal And Not allowglobals Then
               globalvar what$, s$
                End If
        ElseIf I = -1 Then
      
        Else
            If Not useoptionals Then GoTo err10
            globalvar what$, s$
        End If
    ElseIf bs.IsObjectRef(myobject) Then
        If Typename$(myobject) = "lambda" Then
            If flag2 Then
               I = globalvar(what$, Empty)
            ElseIf GetVar3(bstack, what$, I, , , flag) Then
                CheckVar var(I), s$
            Else
                I = globalvar(what$, Empty)
            End If
            Set var(I) = myobject
            If ohere$ = vbNullString Then
                GlobalSub what$ + "()", "", , , I
            Else
                GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , I
            End If
            MyRead = True
        ElseIf Typename$(myobject) = mGroup Then
        Set bstack.lastobj = myobject
         Set myobject = Nothing
         If jumpAs Then jumpAs = False: GoTo existAs19
         ff = 1
         If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs19:
            If Not Fast2Varl(rest$, "ΟΜΑΔΑ", 5, "GROUP", 5, 5, ff) Then SyntaxError: Exit Function
        
            If Not ProcGroup(100, bstack, what$, Lang) Then
                MyRead = False
                Set bstack.lastobj = Nothing
                Exit Function
            End If
                Set myobject = Nothing
                MyRead = True
           
        Else
                 If Not ProcGroup(100, bstack, what$, Lang) Then
                MyRead = False
                Set bstack.lastobj = Nothing
                Exit Function
            End If
                Set myobject = Nothing
                MyRead = True
        End If
         Else
                MyRead = False
            End If
    Else
         If FastSymbol(rest$, "=") Then
         optlocal = Not useoptionals: useoptionals = True: If Not ISSTRINGA(rest$, s$) Then GoTo er111
         GoTo contstrhere
         
         Else
    bstack.soros.drop 1
        MissStackStr
        MyRead = False
        Exit Do
        End If
    End If
Case 4
    If bs.IsNumber(p) Then
        If Not lookOne(rest$, ",") Then
        If jumpAs Then jumpAs = False: GoTo existAs20
        ff = 1
        If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs20:
            If Fast2Varl(rest$, "ΑΡΙΘΜΟΣ", 7, "DECIMAL", 7, 7, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                    p = CDec(p)
            ElseIf Fast2Varl(rest$, "ΔΙΠΛΟΣ", 6, "DOUBLE", 6, 6, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If

                p = CDbl(p)
            ElseIf Fast2Varl(rest$, "ΑΠΛΟΣ", 5, "SINGLE", 6, 6, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CSng(p)
            ElseIf Fast2Varl(rest$, "ΛΟΓΙΚΟΣ", 7, "BOOLEAN", 7, 7, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CBool(p)
            ElseIf Fast2Varl(rest$, "ΜΑΚΡΥΣ", 6, "LONG", 4, 6, ff) Then
                    If Fast2Varl(rest$, "ΜΑΚΡΥΣ", 6, "LONG", 4, 6, ff) Then
                        If FastSymbol(rest$, "=") Then
                            If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                            optlocal = Not useoptionals: useoptionals = True
                        End If
                        p = cInt64(p)
                    Else
                        If FastSymbol(rest$, "=") Then
                            If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                            optlocal = Not useoptionals: useoptionals = True
                        End If
                        p = CLng(p)
                    End If
            ElseIf Fast2Varl(rest$, "ΑΚΕΡΑΙΟΣ", 8, "INTEGER", 7, 8, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CInt(p)
            ElseIf Fast2Varl(rest$, "ΛΟΓΙΣΤΙΚΟΣ", 10, "CURRENCY", 8, 10, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CCur(p)
            Else
            GoTo er110
            End If
            
            ElseIf FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If

    End If
    
        MyRead = True
        If flag2 Then
            globalvar what$, MyRound(p)
        ElseIf GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True, ok) Then
            If isAglobal And Not allowglobals Then
                globalvar what$, MyRound(p)
            Else
                var(I) = MyRound(p)
            End If
        ElseIf I = -1 Then
            bstack.SetVar what$, p
        Else
            globalvar what$, MyRound(p)
        End If
    ElseIf bs.IsOptional Then

        MyRead = True
        If Not lookOne(rest$, ",") Then
        If jumpAs Then jumpAs = False: GoTo existAs21
        ff = 1
        If Fast2VarNoTrim(rest$, "ΩΣ", 2, "AS", 2, 3, ff) Then
existAs21:
            If Fast2Varl(rest$, "ΑΡΙΘΜΟΣ", 7, "DECIMAL", 7, 7, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                    p = CDec(p)
            ElseIf Fast2Varl(rest$, "ΔΙΠΛΟΣ", 6, "DOUBLE", 6, 6, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CDbl(p)
            ElseIf Fast2Varl(rest$, "ΑΠΛΟΣ", 5, "SINGLE", 6, 6, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CSng(p)
            ElseIf Fast2Varl(rest$, "ΛΟΓΙΚΟΣ", 7, "BOOLEAN", 7, 7, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CBool(p)
            ElseIf Fast2Varl(rest$, "ΜΑΚΡΥΣ", 6, "LONG", 4, 6, ff) Then
                    If Fast2Varl(rest$, "ΜΑΚΡΥΣ", 6, "LONG", 4, 6, ff) Then
                        If FastSymbol(rest$, "=") Then
                            If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                            optlocal = Not useoptionals: useoptionals = True
                        End If
                        p = cInt64(p)
                    Else
                        If FastSymbol(rest$, "=") Then
                            If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                            optlocal = Not useoptionals: useoptionals = True
                        End If
                        p = CLng(p)
                    End If
            ElseIf Fast2Varl(rest$, "ΑΚΕΡΑΙΟΣ", 8, "INTEGER", 7, 8, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CInt(p)
                
            ElseIf Fast2Varl(rest$, "ΛΟΓΙΣΤΙΚΟΣ", 10, "CURRENCY", 8, 10, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CCur(p)
            Else
            GoTo er110
            Exit Function
            End If
        ElseIf FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
          Else
            p = 0#
        End If
        
            p = MyRound(p)
        If flag2 Then
         If Not useoptionals Then GoTo err100
            globalvar what$, p
        ElseIf GetVar3(bstack, what$, I, , , flag, , checktype, isAglobal, True) Then
            If isAglobal And Not allowglobals Then
               globalvar what$, p
                End If
        ElseIf I = -1 Then
            
        Else
        If Not useoptionals Then GoTo err100
        If VarType(p) = vbEmpty Then p = 0#
        
            globalvar what$, p
        End If
    Else
        bstack.soros.drop 1
        MissStackNumber
        MyRead = False
        Exit Do
    End If
Case 5, 7
    MyRead = False
    If FastSymbol(rest$, ")") Then
        MyRead = globalArrByPointer(bs, bstack, what$, flag2, allowglobals): If Not MyRead Then SyntaxError: Exit Do
    Else
        If neoGetArray(bstack, what$, pppp) And Not flag2 Then
            If Not NeoGetArrayItem(pppp, bs, what$, it, rest$) Then Exit Do
        Else
            Exit Do
        End If
    If IsOperator(rest$, ".") Then
        If Not pppp.ItemType(it) = mGroup Then
            MyEr "Expected group", "Περίμενα ομάδα"
            MyRead = False: Exit Function
        Else
             I = 1
            aheadstatus rest$, False, I
            ss$ = Left$(rest$, I - 1)
            MyRead = SpeedGroup(bstack, pppp, "@READ", ".", ss$, it) <> 0
            Set pppp = Nothing
            rest$ = Mid$(rest$, I)
            GoTo loopcont123
        End If
    ElseIf IsOperator(rest$, "|") Then
    If placevalue(bstack, pppp, it, rest$) Then
        MyRead = True
        GoTo loopcont123
    End If
    End If
    If bs.IsObjectRef(myobject) Then
        If Typename$(myobject) = mGroup Then
            If myobject.IamFloatGroup Then
                Set pppp.item(it) = myobject
                Set myobject = Nothing
                MyRead = True
            Else
                BadGroupHandle
                MyRead = False
                Set myobject = Nothing
                Exit Function
            End If
            GoTo loopcont123
        ElseIf Typename$(myobject) = "lambda" Then
            Set pppp.item(it) = myobject
            Set myobject = Nothing
            MyRead = True
            GoTo loopcont123
        ElseIf Typename$(myobject) = myArray Then
            If myobject.Arr Then
                Set pppp.item(it) = CopyArray(myobject)
            Else
                Set pppp.item(it) = myobject
            End If
            Set myobject = Nothing
            MyRead = True
            GoTo loopcont123
        ElseIf Typename$(myobject) = mHdlr Then
            If myobject.indirect > -0 Then
                Set pppp.item(it) = myobject
            Else
                p = myobject.T1
                If CheckDeepAny(myobject) Then
                    If TypeOf myobject Is mHandler Then
                        Set pppp.item(it) = myobject
                    Else
                        Set usehandler = New mHandler
                        Set pppp.item(it) = usehandler
                        usehandler.T1 = p
                        Set usehandler.objref = myobject
                        Set usehandler = Nothing
                    End If
                    Set myobject = Nothing
                End If
            End If
            MyRead = True
            GoTo loopcont123
        ElseIf Typename$(myobject) = mProp Then
            Set pppp.item(it) = myobject
            Set myobject = Nothing
            MyRead = True
            GoTo loopcont123
        ElseIf Typename$(myobject) = "BigInteger" Then
            Set pppp.item(it) = CopyBigInteger(myobject)
            Set myobject = Nothing
            MyRead = True
            GoTo loopcont123
        End If
    ElseIf bs.IsOptionalForArray(useoptionals) Then
        ' do nothing
        MyRead = True
    Else
    
        If Not bs.IsNumber(p) Then
            If pppp.IsStringItem(it) Then
                If bs.IsString(s$) Then
                pppp.item(it) = s$
                Else
                bstack.soros.drop 1
                MissStackStr
                MyRead = False
                Exit Do
                End If
            ElseIf pppp.MyTypeToBe = vbVariant Then
            If bs.IsString(s$) Then
                pppp.item(it) = s$
                Else
                bstack.soros.drop 1
                MissStackStr
                MyRead = False
                Exit Do
                End If
            Else
            bstack.soros.drop 1
                MissStackNumber
                MyRead = False
                Exit Do
                End If
            ElseIf x1 = 7 Then
                pppp.item(it) = Round(p)
            Else
                pppp.item(it) = p
            End If
        End If
        MyRead = True
    End If
Case 6
    MyRead = False
    If FastSymbol(rest$, ")") Then
        MyRead = globalArrByPointer(bs, bstack, what$, flag2): If Not MyRead Then SyntaxError: Exit Do
    Else
        If neoGetArray(bstack, what$, pppp) And Not flag2 Then
            If Not NeoGetArrayItem(pppp, bs, what$, it, rest$) Then Exit Do
        Else
            Exit Do
        End If
        If Not bs.IsString(s$) Then
            If bs.IsObjectRef(myobject) Then
                If Typename$(myobject) = "lambda" Then
                    Set pppp.item(it) = myobject
                    Set myobject = Nothing
                ElseIf Typename$(myobject) = mGroup Then
                    Set pppp.item(it) = myobject
                    Set myobject = Nothing

            ElseIf Typename$(myobject) = myArray Then
                    If myobject.Arr Then
                        Set pppp.item(it) = CopyArray(myobject)
                    Else
                        Set pppp.item(it) = myobject
                    End If
                    Set myobject = Nothing
                ElseIf Typename$(myobject) = mHdlr Then
                    If myobject.indirect > -0 Then
                    Set pppp.item(it) = myobject
                    Else
                    p = myobject.T1
                    If CheckDeepAny(myobject) Then
                    If TypeOf myobject Is mHandler Then
                    Set pppp.item(it) = myobject
                    Else
                        Set usehandler = New mHandler
                        Set pppp.item(it) = usehandler
                        usehandler.T1 = p
                        Set usehandler.objref = myobject
                        Set usehandler = Nothing
                    End If
                    Set myobject = Nothing
                    End If
                    
                    End If
              ElseIf Typename$(myobject) = mProp Then
                    Set pppp.item(it) = myobject
                    Set myobject = Nothing
                Else
                    MissStackStr
                    Exit Do
                End If
          ElseIf bs.IsOptionalForArray(useoptionals) Then
            MyRead = True
            Else
            bstack.soros.drop 1
                MissStackStr
                Exit Do
            End If

        Else
            If Not MyIsObject(pppp.item(it)) Then
                pppp.item(it) = s$
            ElseIf pppp.ItemType(it) = mGroup Then
            ' do something
            Else
                Set pppp.item(it) = New Document
                CheckVar pppp.item(it), s$
            End If
        End If
        MyRead = True
    End If
Case 8
jump8:
it = -1
x1 = ExecuteVar(1, 9, bstack, what$, rest$, it, 0&, False, False, 91, "", 0, "", False)
MyRead = True

End Select
End If
loopcont123:
If MaybeIsSymbol(rest$, "@&!#~") Then SyntaxError: MyRead = False
If optlocal Then useoptionals = False

Loop Until Not FastSymbol(rest$, ",")
Exit Function
err10:
            MyEr "Variable " + what$ + " can't initialized", "Η μεταβλητή " + what$ + " δεν αρχικοποιθηκε"
            MyRead = False
            Exit Function
err100:
            MyEr "Parameter is not optional", "Η παράμετρος είναι απαραίτητη"
            MyRead = False
            Exit Function
er103:
            MyEr "Wrong object type", "Λάθος τύπος αντικειμένου"
            MyRead = False
            Exit Function
er104:
            MyEr "Can't assign value to object", "Δεν μπορώ να δώσω τιμή σε αντικείμενο"
            MyRead = False
            Exit Function
er105:
            MyEr "Cant' Assign number", "Δεν μπορώ να θέσω τιμή κατά το διάβασμα τιμών"
            MyRead = False
            Exit Function
er106:
            MyEr "No function definition founded", "Δεν βρέθηκε ορισμός συνάρτησης"
            MyRead = False
            Exit Function
er107:
            MyEr "Syntax error, use )", "Συντακτικό λάθος βάλε )"
            MyRead = False
            Exit Function
er108:
            MyEr "Try other array name", "Δοκίμασε άλλο όνομα πίνακα"
            MyRead = False
            Exit Function
er109:
            MyEr "Cant' change type of variable", "Δεν μπορώ να αλλάξω τύπο μεταβλητής"
            MyRead = False
            Exit Function
er110:
            MyEr "need a right type after as", "χρειάζομαι ένα σωστό τύπο μετά την ως"
            MyRead = False
            Exit Function
er111:
            MyEr "Missing String literal", "Δεν βρήκα σταθερή αλφαριθμητική"
            MyRead = False
er112:
            MyEr "Wrong Enumeration type", "Λάθος τύπος απαριθμητή"
            MyRead = False
            Exit Function
er113:
            MyEr "Expected Group of type " + s$, "Περίμενα Ομάδα τύπου " + s$
            MyRead = False
            Exit Function
er1234:
            Err.Clear
            OverflowValue CInt(it)
            MyRead = False
End Function

Public Sub PlaceValue2UDT(p, Name$, v)
    If Len(Name$) < 0 Then Exit Sub
    vbaVarLateMemSt p, ByVal StrPtr(LCase(Name$)), v
End Sub
Public Sub PlaceValue2UDTArray(p, ByVal Name$, v, Index As Long)
    Dim vv, zero
    If Right$(Name$, 1) = "(" Then Name$ = Left$(Name$, Len(Name$) - 1)
    If Len(Name$) < 0 Then Exit Sub
    SwapVariant vv, GetUDTValue(p, Name$)
    If Index < LBound(vv) Or Index > UBound(vv) Then
        errOutOfLimit
        Exit Sub
    End If
    vv(Index) = v
    CopyMemory ByVal VarPtr(vv), ByVal VarPtr(zero), 16
  
End Sub

Private Sub AssingByRef( _
            ByRef o As Variant, v)
    o = v
End Sub
Function UDTValue(p, Name$)
        On Error Resume Next
        SwapVariant UDTValue, GetUDTValue(p, Name$)
End Function
Function GetUDTValue(p, Name$)
    If Len(Name$) < 1 Then Exit Function
    If MemInt(VarPtr(p)) <> vbUserDefinedType Then Exit Function
    
    Dim r As Long, ret
    Static fptr As Long
    Static t(8) As Integer
    For r = 0 To 7: t(r) = 0: Next
    Static btASM(50)  As Byte
    If fptr = 0 Then fptr = GetFuncPtr("msvbvm60.dll", "__vbaVarLateMemCallLdRf")
    'R = vbaVarLateMemCallLdRf(VarPtr(t(0)), VarPtr(p), StrPtr(LCase(Name$)), (0&), (0&))
    r = CallCdecl(btASM, fptr, VarPtr(t(0)), VarPtr(p), StrPtr(LCase(Name$)), (0&), (0&))

again:
    If (t(0) And &H4000) > 0 Then
        t(0) = t(0) - &H4000
        Select Case t(0)
        Case vbDouble, vbCurrency, 20, vbDate
            CopyMemory ByVal VarPtr(GetUDTValue) + 8, ByVal MemLong(VarPtr(t(4))), 8
        Case vbDecimal
            CopyMemory ByVal VarPtr(GetUDTValue), ByVal MemLong(VarPtr(t(4))), 16
            GoTo endfix
        Case vbVariant
            CopyMemory ByVal VarPtr(t(0)), ByVal MemLong(VarPtr(t(4))), 16
            GoTo again
        Case Else
            CopyMemory ByVal VarPtr(GetUDTValue) + 8, ByVal MemLong(VarPtr(t(4))), 4
        End Select
        MemInt(VarPtr(GetUDTValue)) = t(0)
    Else
        CopyMemory ByVal VarPtr(GetUDTValue), ByVal VarPtr(t(1)), 16
    End If
endfix:
End Function
Function GetUDTValueArray(p, Name$, index1 As Long)
    If Len(Name$) < 1 Then Exit Function
    If MemInt(VarPtr(p)) <> vbUserDefinedType Then Exit Function
    Name$ = Left$(Name$, Len(Name$) - 1)
    Dim r As Long, tst
    Static fptr As Long
    Static btASM(50)  As Byte
    If fptr = 0 Then fptr = GetFuncPtr("msvbvm60.dll", "__vbaVarLateMemCallLdRf")
    r = 1
    r = CallCdecl(btASM, fptr, VarPtr(tst), VarPtr(p), StrPtr(LCase(Name$)), (0&), (0&))
    If index1 < LBound(tst) Or index1 > UBound(tst) Then
        errOutOfLimit
        Exit Function
    End If
    GetUDTValueArray = CVar(tst(index1))
End Function
Private Function placevalue(bstack As basetask, pppp As mArray, there As Long, rest$) As Boolean
Dim x1 As Long, sp, ss$, Index
If bstack.IsAny(sp) Then
    x1 = IsLabelOnly(rest$, ss$)
    If x1 = 1 Then
            placevalue = pppp.PlaceValue2UDT(there, ss$, sp)
    ElseIf x1 = 5 Then
        If IsExp(bstack, rest$, Index, , True) Then
        If FastSymbol(rest$, ")") Then
            placevalue = pppp.PlaceValue2UDTArray(there, ss$, sp, CLng(Index))
            End If
        End If
    End If
End If
End Function
Private Function placevalue2(bstack As basetask, v, rest$, sp) As Boolean
Dim x1 As Long, ss$, Index
    x1 = IsLabelOnly(rest$, ss$)
    If x1 = 1 Then
            PlaceValue2UDT v, ss$, sp
    ElseIf x1 = 5 Then
        If IsExp(bstack, rest$, Index, , True) Then
            If FastSymbol(rest$, ")") Then
                PlaceValue2UDTArray v, ss$, sp, CLng(Index)
            End If
        End If
    End If
    placevalue2 = Not (Err.Number > 0 Or LastErNum > 0)
End Function
Private Sub Assign(ss$, p)
        Select Case MemInt(VarPtr(p))
        Case vbString
            SwapString2Variant ss$, p
        Case vbBoolean
            ss$ = Format$(p, DefBooleanString)
        Case 20
            ss$ = CStr(p)
        Case vbDate
            ss$ = p
        Case 36
            ss$ = "*" + Typename(p)
        Case Else
            ss$ = fixthis(p)
        End Select
End Sub
Private Sub Assign2(ss$, p)
        Select Case MemInt(VarPtr(p))
        Case vbString
        Case vbBoolean
            p = Format$(p, DefBooleanString)
        Case 20
            p = CStr(p)
        Case 36
            p = "*" + Typename(p)
        Case Else
            ss$ = fixthis(p)
            p = vbNullString
            SwapString2Variant ss$, p
        End Select
End Sub

Public Function fixthis(p As Variant) As String
    If TypeOf p Is cxComplex Then
        If p.I = 0 Then
            fixthis = fixthis(CVar(p.r))
        ElseIf p.r = 0 Then
            fixthis = "(" & fixthis(CVar(p.I)) & "i)"
        Else
            If p.I < 0 Then fixthis = "" Else fixthis = "+"
            If Abs(p.I) = 1 Then
                If p.I < 0 Then
                    fixthis = "(" & fixthis(CVar(p.r)) & "-i)"
                Else
                    fixthis = "(" & fixthis(CVar(p.r)) & "+i)"
                End If
            Else
                fixthis = "(" & fixthis(CVar(p.r)) & fixthis & fixthis(CVar(p.I)) & "i)"
            End If
        End If
    ElseIf MemInt(VarPtr(p)) = vbDate Then
            fixthis = FormatDateWithLocale(GetlocaleString2(&H1F, Clid), CDate(p), Clid)
    Else
        fixthis = LTrim$(str(p))
        If Left$(fixthis, 1) = "." Then
        fixthis = "0" + fixthis
        ElseIf Left$(fixthis, 2) = "-." Then
        fixthis = "-0" + Mid$(fixthis, 2)
        End If
        If InStr(fixthis, ".") > 0 Then
        If NoUseDec Then fixthis = Replace(fixthis, ".", NowDec$)
        End If
    End If
End Function

Public Function IntSqrdEC(sA) As Variant
    sA = Int(sA)
    If sA = 0 Or sA < 0 Then
        MyEr "Zero or negative paramter for integer Square Root", "Μηδενική ή Αρνητική παράμετρος για ακέραια τετραγωνική ρίζα"
        Exit Function
    ElseIf sA > CDec("9903520314283042199192993792") Then
        OverflowValue vbDecimal
        Exit Function
    End If
    Dim q, r, t, z
    z = CDec(sA)
    r = CDec("0")
    q = CDec("1")
    Do
    q = q * 4
    Loop Until q > sA
    Do
        If q <= 1 Then Exit Do
        q = Int(q / 4&)
        t = z - r - q
        r = Int(r / 2)
        If t >= -1 Then
            z = t
            r = r + q
        End If
    Loop
    IntSqrdEC = r
End Function
Function CopyBigInteger(p, Optional A) As BigInteger
    Dim check As BigInteger
    Set check = p
    If check.Unique Then
        Set CopyBigInteger = check
    Exit Function
    End If
    Set CopyBigInteger = New BigInteger
    If Not IsMissing(A) Then
        CopyBigInteger.Load2 p, A
    Else
        CopyBigInteger.Load2 p
    End If
End Function

