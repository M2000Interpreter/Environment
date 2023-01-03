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

Private Declare Function PeekArray Lib "kernel32" Alias "RtlMoveMemory" (Arr() As Any, Optional ByVal Length As Long = 4) As PeekArrayType
Private Declare Function SafeArrayGetDim Lib "OleAut32.dll" (ByVal Ptr As Long) As Long
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
    ByVal pszUrl As Long, _
    ByVal pszEscaped As Long, _
    ByRef pcchEscaped As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeW" ( _
    ByVal pszUrl As Long, _
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
    ByVal pszUrl As Long, _
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
Sub FileReadString(FileH As Long, R$, bytes As Long)
    Dim Buf1() As Byte
    If bytes <= 0 Then R$ = vbNullString: Exit Sub
    If FileH = 0 Then Exit Sub
    ReDim Buf1(0 To bytes - 1)
    R$ = Buf1()
    API_ReadFile FileH, bytes, Buf1()
    CopyMemory ByVal StrPtr(R$), Buf1(0), bytes
End Sub
Sub FileReadBytes(FileH As Long, Buf1() As Byte, bytes As Long)
    If bytes <= 0 Then Exit Sub
    If FileH = 0 Then Exit Sub
    API_ReadFile FileH, bytes, Buf1
End Sub
Sub FileWriteString(FileH As Long, R$)
    Dim Buf1() As Byte, bytes As Long
    If LenB(R$) = 0 Then Exit Sub
    If FileH = 0 Then Exit Sub
    bytes = LenB(R$)
    ReDim Buf1(bytes - 1)
    CopyMemory Buf1(0), ByVal StrPtr(R$), bytes
    API_WriteFile FileH, bytes, Buf1()
End Sub
Sub FileWriteBytes(FileH As Long, ByRef Buf1() As Byte)
    Dim bytes As Long
    If PeekArray(Buf1()).Ptr = 0 Then Exit Sub
    If FileH = 0 Then Exit Sub
    bytes = UBound(Buf1()) - LBound(Buf1()) + 1
    API_WriteFile FileH, bytes, Buf1()
End Sub
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
    MyEr "Can't open file, Not Sharing allowed" & FileError, "ƒÂÌ ÏÔÒ˛ Ì· ·ÌÔﬂÓ˘ ÙÔ ·Ò˜ÂﬂÔ, ‰ÂÌ ÂÈÙÒ›ÂÙ·È ÏÔﬂÒ·ÛÏ·"
    Exit Function
 End If
there:
    MyEr "Can't open file, error :" & FileError, "ƒÂÌ ÏÔÒ˛ Ì· ·ÌÔﬂÓ˘ ÙÔ ·Ò˜ÂﬂÔ :" & FileError
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
Public Property Get uni(RHS) As Long
Dim H
If InUseHandlers.ExistKey(RHS) Then
    H = InUseHandlers.Value
    uni = H(3)
End If
End Property

Public Property Get Fstep(RHS) As Long
Dim H
If InUseHandlers.ExistKey(RHS) Then
    H = InUseHandlers.Value
    Fstep = H(2)
End If
End Property
Public Property Get Fkind(RHS) As Ftypes
Dim H
If InUseHandlers.ExistKey(RHS) Then
    H = InUseHandlers.Value
    Fkind = H(1)
End If
End Property
Public Property Let FileSeek(RHS, vvv)
Dim H, where, ret As Currency, lowlong As Long, highlong As Long
Dim FileError As Long
ret = CCur(Int(vvv)) - 1
If InUseHandlers.ExistKey(RHS) Then
    H = InUseHandlers.Value
    where = H(0)
    Size2Long ret, lowlong, highlong
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_BEGIN)
    FileError = GetLastError()
    If lowlong = INVALID_SET_FILE_POINTER And FileError <> 0 Then
        MyEr "Can't write the seek value", "ƒÂÌ ÏÔÒ˛ Ì· „Ò‹¯˘ ÙÁ ÙÈÏﬁ ÏÂÙ‹ËÂÛÁÚ"
    End If
Else
    MyEr "No such file handler", "ƒÂÌ ı‹Ò˜ÂÈ Ù›ÙÔÈÔ ˜ÂÈÒÈÛÙﬁÚ ·Ò˜ÂﬂÔı"
End If
End Property
Public Property Get FileSeek(RHS) As Variant
Dim H, where, ret As Currency, lowlong As Long, highlong As Long
FileSeek = ret
If InUseHandlers.ExistKey(RHS) Then
    H = InUseHandlers.Value
    where = H(0)
    lowlong = 0
    highlong = 0
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_CURRENT)
    If lowlong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value", "ƒÂÌ ÏÔÒ˛ Ì· ‰È·‚‹Û˘ ÙÁ ÙÈÏﬁ ÏÂÙ‹ËÂÛÁÚ"
    Else
        Long2Size lowlong, highlong, ret
        FileSeek = ret + 1
    End If
Else
    MyEr "No such file handler", "ƒÂÌ ı‹Ò˜ÂÈ Ù›ÙÔÈÔ ˜ÂÈÒÈÛÙﬁÚ ·Ò˜ÂﬂÔı"
End If
End Property
Public Property Let FileSeekFH(where As Long, ByVal ret As Currency)
Dim lowlong As Long, highlong As Long
Dim FileError As Long
    Size2Long ret - 1@, lowlong, highlong
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_BEGIN)
    FileError = GetLastError()
    If lowlong = INVALID_SET_FILE_POINTER And FileError <> 0 Then
        MyEr "Can't write the seek value", "ƒÂÌ ÏÔÒ˛ Ì· „Ò‹¯˘ ÙÁ ÙÈÏﬁ ÏÂÙ‹ËÂÛÁÚ"
    End If
End Property
Public Property Get FileSeekFH(where As Long) As Currency
Dim lowlong As Long, highlong As Long
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_CURRENT)
    If lowlong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value", "ƒÂÌ ÏÔÒ˛ Ì· ‰È·‚‹Û˘ ÙÁ ÙÈÏﬁ ÏÂÙ‹ËÂÛÁÚ"
    Else
        Long2Size lowlong, highlong, FileSeekFH
        FileSeekFH = FileSeekFH + 1
    End If
End Property
Public Property Get FileEOFFH(where As Long) As Boolean
Dim ret As Currency, lowlong As Long, highlong As Long
Dim fsize As Currency
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_CURRENT)
    If lowlong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value on Eof() function", "ƒÂÌ ÏÔÒ˛ Ì· ‰È·‚‹Û˘ ÙÁ ÙÈÏﬁ ÏÂÙ‹ËÂÛÁÚ ÛÙÁ ÛıÌ‹ÒÙÁÛÁ ÃÂÙ‹ËÂÛÁ()"
    Else
        Long2Size lowlong, highlong, ret
        lowlong = GetFileSize(where, highlong)
        Long2Size lowlong, highlong, fsize
        FileEOFFH = ret >= fsize
    End If
End Property

Public Property Get FileEOF(RHS) As Boolean
Dim H, where, ret As Currency, lowlong As Long, highlong As Long
Dim fsize As Currency
If InUseHandlers.ExistKey(RHS) Then
    H = InUseHandlers.Value
    where = H(0)
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_BEGIN)
    If lowlong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value on Eof() function", "ƒÂÌ ÏÔÒ˛ Ì· ‰È·‚‹Û˘ ÙÁ ÙÈÏﬁ ÏÂÙ‹ËÂÛÁÚ ÛÙÁ ÛıÌ‹ÒÙÁÛÁ ÃÂÙ‹ËÂÛÁ()"
    Else
        Long2Size lowlong, highlong, ret
        lowlong = GetFileSize(where, highlong)
        Long2Size lowlong, highlong, fsize
        FileEOF = ret >= fsize
    End If
Else
    MyEr "No such file handler", "ƒÂÌ ı‹Ò˜ÂÈ Ù›ÙÔÈÔ ˜ÂÈÒÈÛÙﬁÚ ·Ò˜ÂﬂÔı"
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
Public Function ReadFileHandler(H&) As Variant
If InUseHandlers.Find(CVar(H&)) Then
    ReadFileHandler = InUseHandlers.sValue
Else
    MyEr "No such file handler", "ƒÂÌ ı‹Ò˜ÂÈ Ù›ÙÔÈÔ ˜ÂÈÒÈÛÙﬁÚ ·Ò˜ÂﬂÔı"
End If
End Function
' internal use  You have to close file first
Public Sub CloseHandler(RHS)
Dim H&, ar() As Variant
On Error Resume Next
If InUseHandlers.ExistKey(RHS) Then
    H& = CLng(InUseHandlers.sValue)
    API_CloseFile H&
    H& = InUseHandlers.KeyToNumber
    InUseHandlers.RemoveWithNoFind
    FreeUseHandlers.AddKey CVar(H&)
Else
    ' no error... (I am thinking about it)
End If

End Sub
Public Sub CloseAllHandlers()
Dim H&
On Error Resume Next
Do While InUseHandlers.count > 0
    InUseHandlers.ToEnd
    H& = CLng(InUseHandlers.sValue)
    API_CloseFile H&
    H& = InUseHandlers.KeyToNumber
    InUseHandlers.RemoveWithNoFind
    FreeUseHandlers.AddKey CVar(H&)
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
    MyEr "Can't read the file length", "ƒÂÌ ÏÔÒ˛ Ì· ‰È·‚‹Û˘ ÙÔ ÏﬁÍÔÚ ÙÔı ·Ò˜ÂﬂÔı"
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



Public Function ApiCanonicalize(ByVal url As String, Optional dwFlags As Long = 0) As String
    url = Left$(url, INTERNET_MAX_URL_LENGTH)
   Dim dwSize As Long, res As String
   
   If Len(url) > 0 Then
   
      ApiCanonicalize = space$(INTERNET_MAX_URL_LENGTH)
      dwSize = Len(ApiCanonicalize)
     
      If UrlCanonicalizeApi(StrPtr(url), _
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
    ByVal url As String, _
    Optional ByVal PlusSpace As Boolean = True, Optional Flags As Long = 0) As String
    url = Left$(url, INTERNET_MAX_URL_LENGTH)
    Dim cchUnescaped As Long
    Dim hResult As Long
    
    If PlusSpace Then url = Replace$(url, "+", " ")
    cchUnescaped = Len(url)
    URLDecode = String$(cchUnescaped, 0)
    hResult = UrlUnescape(StrPtr(url), StrPtr(URLDecode), cchUnescaped, Flags)
    If hResult = E_POINTER Then
        URLDecode = String$(cchUnescaped, 0)
        hResult = UrlUnescape(StrPtr(url), StrPtr(URLDecode), cchUnescaped, Flags)
    End If
    
    If hResult <> S_OK Then
        MyEr "can't decode this url", "‰ÂÌ ÏÔÒ˛ Ì· ·ÔÍ˘‰ÈÍÔÔÈﬁÛ˘ ÙÁÌ ‰ÈÂ˝ËıÌÛÁ"
        Exit Function
    End If
    
    URLDecode = Left$(URLDecode, cchUnescaped)
End Function

Public Function URLEncode( _
    ByVal url As String, _
    Optional ByVal SpacePlus As Boolean = True) As String
    url = Left$(url, INTERNET_MAX_URL_LENGTH)
    Dim cchEscaped As Long
    Dim hResult As Long
    If SpacePlus Then
      
        url = Replace$(url, " ", "+")
    End If
    cchEscaped = Len(url) * 1.5
    URLEncode = String$(cchEscaped, 0)
    hResult = UrlEscape(StrPtr(url), StrPtr(URLEncode), cchEscaped, URL_ESCAPE_PERCENT + &H40000)
    If hResult = E_POINTER Then
        URLEncode = String$(cchEscaped, 0)
        hResult = UrlEscape(StrPtr(url), StrPtr(URLEncode), cchEscaped, URL_ESCAPE_PERCENT + &H40000)
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
    Dim i As Long
    ' new from me
   ' Address = URLDecode(Address, True)
    '
     strOutput$ = Replace(Address, "\", "/")
    lngCharCount = Len(strOutput)
    i = InStr(1, strOutput, "/")
    If i Then
        If i - InStr(1, strOutput, ":") > 1 Then
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
    i = InStr(1, strOutput, "@", vbTextCompare)
    If i > 0 Then GoTo 500
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
    Dim i As Long
    ' new from me
    Address = URLDecode(Address, True)
    '
    strOutput$ = Replace(Address, "\", "/")
    lngCharCount = Len(strOutput)
    i = InStr(1, strOutput, "/")
    If i Then
        If i - InStr(1, strOutput, ":") > 1 Then
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
    i = InStr(1, strOutput, "@", vbTextCompare)
    If i > 0 Then strOutput = Mid$(strOutput, i + 1)
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
Public Function GetHost(url$) As String
        Dim W As Long
        GetHost = GetDomainName(url$, True)
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
    Dim slen As Long, m$: slen = Len(cc)
    Dim i As Long
    
    If slen > 0 Then
        ReDim res(slen) As String
        Dim ccode As Byte
        Dim cp1, cp2, cp3 As Integer
        Dim space As String
    
        If space_as_plus Then space = "+" Else space = "%20"
    If typedata = 0 Then
            For i = 1 To slen
            ccode = Asc(Mid$(cc, i, 1))
            Select Case ccode
                Case 97 To 122, 65 To 90, 48 To 57
                     res(i) = Chr(ccode)
                Case 32
                    res(i) = space
                Case 0 To 15
                    res(i) = "%0" & Hex(ccode)
                Case Else
                    res(i) = "%" & Hex(ccode)
            End Select
        Next i
    ElseIf typedata = 1 Then
        ' RFC3986
        For i = 1 To slen
            ccode = Asc(Mid$(cc, i, 1))
            Select Case ccode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                     res(i) = Chr(ccode)
                Case 32
                    res(i) = space
                Case 0 To 15
                    res(i) = "%0" & Hex(ccode)
                Case Else
                    res(i) = "%" & Hex(ccode)
            End Select
        Next i
    ElseIf typedata = 2 Then
        ' HMTL5
        For i = 1 To slen
            ccode = Asc(Mid$(cc, i, 1))
            Select Case ccode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 42
                     res(i) = Chr(ccode)
                Case 32
                    res(i) = space
                Case 0 To 15
                    res(i) = "%0" & Hex(ccode)
                Case Else
                    res(i) = "%" & Hex(ccode)
            End Select
        Next i
    End If
        URLEncodeEsc = Join(res, "")
    End If
End Function
Function DecodeEscape(c$, plus_as_space As Boolean) As String
If plus_as_space Then c$ = Replace(c$, "+", " ")
Dim a() As String, i As Long
a() = Split(c$, "%")
For i = 1 To UBound(a())
a(i) = Chr(val("&h" + Left$(a(i), 2))) + Mid$(a(i), 3)
Next i
DecodeEscape = utf8decode(StrConv(Join(a(), ""), vbFromUnicode))

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
Function MyRead(jump As Long, bstack As basetask, rest$, Lang As Long, Optional ByVal what$, Optional usex1 As Long, Optional exist As Boolean = False) As Boolean
Dim ps As mStiva, bs As basetask, f As Boolean, ohere$, par As Boolean, flag As Boolean, flag2 As Boolean, ok As Boolean
Dim s$, ss$, pa$, x1 As Long, y1 As Long, i As Long, myobject As Object, it As Long, useoptionals As Boolean, optlocal As Boolean
Dim m As mStiva, checktype As Boolean, allowglobals As Boolean, isAglobal As Boolean, look As Boolean, ByPass As Boolean
Dim usehandler As mHandler, ff As Long, usehandler1 As mHandler
Const mProp = "PropReference"
Const mHdlr = "mHandler"
Const mGroup = "Group"
Const myArray = "mArray"
MyRead = True
Dim p As Variant, x As Double
Dim pppp As mArray
ohere$ = here$
Dim Col As Long
Dim ihavetype As Boolean
look = jump = 1 Or jump = 7
On jump GoTo read, refer, commit, readnew, readlocal, readlet, readfromsub, link
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

readlet:
allowglobals = True
Set bs = bstack
x1 = usex1
If x1 > 3 Then x1 = Abs(IsLabel(bstack, rest$, what$))
'***********************
Select Case x1
Case 1
    If bs.IsObjectRef(myobject) Then
        MyRead = True
        If GetVar3(bstack, what$, i, , , flag, s$, checktype, isAglobal, True, ok) Then
            If Typename$(myobject) = VarTypeName(var(i)) Then
                If Typename$(var(i)) = mGroup Then
                    ss$ = bstack.GroupName
                    If s$ <> "" Then what$ = s$
                    If Len(var(i).GroupName) > Len(what$) Then
                        If var(i).IamRef Then
                            s$ = here$
                            here$ = vbNullString
                            UnFloatGroupReWriteVars bstack, what$, i, myobject
                            here = s$
                        Else
                            UnFloatGroupReWriteVars bstack, what$, i, myobject
                        End If
                        myobject.ToDelete = True
                    Else
                        bstack.GroupName = Left$(what$, Len(what$) - Len(var(i).GroupName) + 1)
                        If Len(var(i).GroupName) > 0 Then
                            what$ = Left$(var(i).GroupName, Len(var(i).GroupName) - 1)
                            s$ = here$
                            here$ = vbNullString
                            UnFloatGroupReWriteVars bstack, what$, i, myobject
                            here = s$
                            myobject.ToDelete = True
                        ElseIf var(i).IamApointer And myobject.IamApointer Then
                            Set var(i) = myobject
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
                ElseIf Typename$(var(i)) = mHdlr Then
                    Set usehandler = myobject
                    Set usehandler1 = var(i)
                    If usehandler1.ReadOnly Then
                        MyRead = False
                        ReadOnly
                        Exit Function
                    ElseIf usehandler1.t1 = usehandler.t1 Then
                        If usehandler1.t1 = 4 Then
                            If usehandler1.objref Is usehandler.objref Then
                                Set var(i) = myobject
                            ElseIf usehandler1.objref.EnumName = usehandler.objref.EnumName Then
                                If usehandler1.objref.ExistFromOther(usehandler.index_cursor) Then
                                    Set usehandler.objref = usehandler1.objref
                                    Set var(i) = usehandler
                                Else
                                    GoTo er103
                                End If
                            Else
                                GoTo er103
                            End If
                        Else
                            Set var(i) = myobject
                        End If
                    ElseIf usehandler1.t1 <> 4 And myobject.t1 = 3 Then
                        Set var(i) = myobject
                    Else
                        GoTo er103
                    End If
                    Set usehandler = Nothing
                    Set usehandler1 = Nothing
                Else
                    Set var(i) = myobject
                End If
            ElseIf x1 = 1 And CheckIsmArray(myobject) Then
                Set usehandler = New mHandler
                Set var(i) = usehandler
                usehandler.t1 = 3
                Set usehandler.objref = myobject
                Set myobject = Nothing
                Set usehandler = Nothing
            Else
                If TypeOf myobject Is mHandler Then
                    Set usehandler = myobject
                    If usehandler.t1 = 4 Then
                        p = usehandler.index_cursor
                        Set myobject = Nothing
                        Set usehandler = Nothing
                        GoTo itisinumber
                    End If
                End If
                GoTo er103
            End If
        Else
            i = globalvar(what$, 0)
            If Typename$(myobject) = mGroup Then
                If myobject.IamApointer Then
                    Set var(i) = myobject
                Else
                    UnFloatGroup bstack, bstack.GroupName + what$, i, myobject, here$ = vbNullString Or Len(bstack.UseGroupname) > 0, , True
                    myobject.ToDelete = True
                End If
            ElseIf Typename$(myobject) = "mEvent" Then
                Set var(i) = myobject
            ElseIf Typename$(myobject) = "lambda" Then
                Set var(i) = myobject
                If ohere$ = vbNullString Then
                    GlobalSub what$ + "()", "", , , i
                Else
                    GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , i
                End If
            ElseIf Typename$(myobject) = mHdlr Then
                Set usehandler = myobject
                If usehandler.indirect > -1 Then
                    Set var(i) = MakeitObjectGeneric(usehandler.indirect)
                Else
                    Set var(i) = usehandler
                End If
            ElseIf Typename$(myobject) = myArray Then
                Set usehandler = New mHandler
                Set var(i) = usehandler
                usehandler.t1 = 3
                Set usehandler.objref = myobject
                Set usehandler = Nothing
            Else
                Set var(i) = myobject
            End If
        End If
    ElseIf bs.IsNumber(p) Then
itisinumber:
        If GetVar3(bstack, what$, i, , , flag, , checktype, isAglobal, True, ok) Then
            If MyIsObject(var(i)) Then
                If TypeOf var(i) Is Group Then
                    If var(i).HasSet Then
                        Set m = bstack.soros
                        Set bstack.Sorosref = New mStiva
                        bstack.soros.PushVal p
                        NeoCall2 bstack, what$ + "." + ChrW(&H1FFF) + ":=()", ok
                        Set bstack.Sorosref = m
                        Set m = Nothing
                    Else
                        GoTo there182741
                    End If
                Else
there182741:
                    If TypeOf var(i) Is Constant Then
                        If var(i).flag Then
                            CantAssignValue
                            Exit Function
                        Else
                            
                            Stop
                            
                        End If
                        
                    ElseIf TypeOf var(i) Is mHandler Then
                        Set usehandler = var(i)
                        If usehandler.t1 <> 4 Then GoTo er104
                        Set myobject = usehandler.objref.SearchValue(p, ok)
                        If ok Then
                            Set var(i) = myobject
                        Else
                            GoTo er112
                        End If
                    Else
                        GoTo er104
                    End If
                End If
            Else
                If checktype Then
                    If ihavetype Then
                        If VarType(var(i)) <> VarType(p) Then
                            GoTo er109
                        ElseIf AssignTypeNumeric(p, VarType(var(i))) Then
                            var(i) = p
                        Else
                            GoTo er105
                        End If
                    ElseIf AssignTypeNumeric(p, VarType(var(i))) Then
                        var(i) = p
                    Else
                        GoTo er105
                    End If
                Else
                    var(i) = p
                End If
            End If
        ElseIf i = -1 Then
                bstack.SetVar what$, p
        Else
                If Not exist Then globalvar what$, p Else Nosuchvariable what$
        End If
    End If
Case 3
    If bs.IsString(s$) Then
        MyRead = True
        If GetVar3(bstack, what$, i, , , flag, , checktype, isAglobal, True) Then
            If MyIsObject(var(i)) Then
                If TypeOf var(i) Is Group Then
                    Set m = bstack.soros
                    Set bstack.Sorosref = New mStiva
                    bstack.soros.PushStr s$
                    NeoCall2 bstack, Left$(what$, Len(what$) - 1) + "." + ChrW(&H1FFF) + ":=()", ok
                    Set bstack.Sorosref = m
                    Set m = Nothing
                ElseIf TypeOf var(i) Is Constant Then
                    CantAssignValue
                    MyRead = False
                    Exit Function
                Else
                    CheckVar var(i), s$
                End If
            Else
                var(i) = s$
            End If
        ElseIf i = -1 Then
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
        If GetVar3(bstack, what$, i, , , flag, , checktype, isAglobal, True) Then
            var(i) = MyRound(p)
        ElseIf i = -1 Then
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
                MyEr "Expected group", "–ÂÒﬂÏÂÌ· ÔÏ‹‰·"
                MyRead = False: Exit Function
            Else
                i = 1
                aheadstatus rest$, False, i
                ss$ = Left$(rest$, i - 1)
                MyRead = SpeedGroup(bstack, pppp, "@READ", ".", ss$, it) <> 0
                Set pppp = Nothing
                rest$ = Mid$(rest$, i)
            End If
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
                        p = usehandler.t1
                        If CheckDeepAny(myobject) Then
                            If TypeOf myobject Is mHandler Then
                                Set pppp.item(it) = myobject
                            Else
                                Set usehandler = New mHandler
                                Set pppp.item(it) = usehandler
                                usehandler.t1 = p
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
                        p = usehandler.t1
                        If CheckDeepAny(myobject) Then
                            If TypeOf myobject Is mHandler Then
                                Set pppp.item(it) = myobject
                            Else
                                Set usehandler = New mHandler
                                Set pppp.item(it) = usehandler
                                usehandler.t1 = p
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
    End Select
    p = 0#
    Exit Function
read:
If FastSymbol(rest$, "?") Then useoptionals = True

flag2 = Fast2LabelNoNum(rest$, "Õ≈œ", 3, "NEW", 3, 3)
If Not flag2 Then flag = Fast2LabelNoNum(rest$, "‘œ–… ¡", 6, "LOCAL", 5, 6)
read123:
Set bs = bstack
contFromRebound:
par = Fast2LabelNoNum(rest$, "¡–œ", 3, "FROM", 4, 4)
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

 If getvar2(bstack, ss$, i, , , flag) Then
 If MyIsObject(var(i)) Then
            ' need to make new stack frame with pointers to
        If Typename(var(i)) <> mGroup Then MyRead = True: Exit Function
        Set ps = New mStiva
        If ohere$ <> "" And Not var(i).IamGlobal Then
        Set myobject = var(i).PrepareSoros(var(), ohere$ + ".")
        Else
        Set myobject = var(i).PrepareSoros(var(), "")
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
ElseIf IsNumberD(ss$, x) Then
ps.DataVal x
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
     
          
    
    If GetGlobalVar(s$, i) Then
    
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
               If Not flag Then
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
                If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                   If Not MyIsObject(var(i)) Then
                        p = var(i)
                        If Not varhash.vType(varhash.index) Then
                            If Not Fast2Varl(rest$, "¡‘’–œ”", 6, "VARIANT", 7, 7, ff) Then MyRead = False: MissType: Exit Function
                            GoTo jumpref02
                        End If
checkconstant:
                        Select Case VarType(p)
                        Case vbDecimal
                            If Not Fast2Varl(rest$, "¡—…»Ãœ”", 7, "DECIMAL", 7, 7, ff) Then MyRead = False: MissType: Exit Function
                        Case vbDouble
                            If Not Fast2Varl(rest$, "ƒ…–Àœ”", 6, "DOUBLE", 6, 6, ff) Then MyRead = False: MissType: Exit Function
                        Case vbSingle
                            If Not Fast2Varl(rest$, "¡–Àœ”", 5, "SINGLE", 6, 6, ff) Then MyRead = False: MissType: Exit Function
                        Case vbBoolean
                            If Not Fast2Varl(rest$, "Àœ√… œ”", 7, "BOOLEAN", 7, 7, ff) Then MyRead = False: MissType: Exit Function
                        Case vbLong
                            If Not Fast2Varl(rest$, "Ã¡ —’”", 6, "LONG", 4, 6, ff) Then MyRead = False: MissType: Exit Function
                        Case vbInteger
                            If Not Fast2Varl(rest$, "¡ ≈—¡…œ”", 8, "INTEGER", 7, 8, ff) Then MyRead = False: MissType: Exit Function
                        Case vbCurrency
                            If Not Fast2Varl(rest$, "Àœ√…”‘… œ”", 10, "CURRENCY", 8, 10, ff) Then MyRead = False: MissType: Exit Function
                        Case 20
                            If Not Fast2Varl(rest$, "Ã¡ —’”", 6, "LONG", 4, 6, ff) Then MyRead = False: MissType: Exit Function
                            If Not Fast2Varl(rest$, "Ã¡ —’”", 6, "LONG", 4, 6, ff) Then MyRead = False: MissType: Exit Function
                        Case vbString
                            If Not Fast2Varl(rest$, "√—¡ÃÃ¡", 6, "STRING", 6, 6, ff) Then
                                MyRead = False: MissType: Exit Function
                            End If
                        Case Else
                            p = IsLabel(bstack, rest$, (what$))  ' just throw any name
                        End Select
                    Else
                        If TypeOf var(i) Is Group Then
                            If Fast2Varl(rest$, "œÃ¡ƒ¡", 5, "GROUP", 5, 5, ff) Then
                            If var(i).IamApointer Then
                                If var(i).link.IamFloatGroup Then GoTo errgr
                                    If Len(var(i).lasthere) = 0 Then
                                        If Not GetVar(bstack, var(i).GroupName, i, True) Then GoTo errgr
                                    Else
                                        If Not GetVar(bstack, var(i).lasthere + "." + var(i).GroupName, i, True) Then GoTo errgr

                                    End If
                                 End If
                            ElseIf Fast2Varl(rest$, "ƒ≈… ‘«”", 7, "POINTER", 7, 7, ff) Then
                                     If Not var(i).IamApointer Then MyRead = False: MissType: Exit Function
                            Else
                            If FastSymbol(rest$, "*") Then
                                If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
                                If Not var(i).IamApointer Then GoTo errgr
                                If var(i).link.IamFloatGroup Then
                                    If Not var(i).link.TypeGroup(s$) Then GoTo errgr
                                Else
                                    If Len(var(i).lasthere) = 0 Then
                                    If GetVar(bstack, var(i).GroupName, it, True) Then
                                        If Not var(it).TypeGroup(s$) Then GoTo errgr
                                    Else
                                        GoTo noref01
                                    End If
                                    Else
                                        If GetVar(bstack, var(i).lasthere + "." + var(i).GroupName, it, True) Then
                                            If Not var(it).TypeGroup(s$) Then GoTo errgr
                                        Else
                                            GoTo noref01
                                        End If
                                    End If
                                    it = 0
                                End If

                            ElseIf FastPureLabel(rest$, s$, , True) = 1 Then
                                    If var(i).IamApointer Then

                                    If var(i).link.IamFloatGroup Then
                                        MyRead = False: MissType: Exit Function
                                    End If
                                        If Len(var(i).lasthere) = 0 Then
                                            If Not GetVar(bstack, var(i).GroupName, i, True) Then
                                                GoTo errgr
                                            End If
                                        ElseIf Not GetVar(bstack, var(i).lasthere + "." + var(i).GroupName, i, True) Then
                                                GoTo errgr
                                        End If
                                    End If
                                    If Not var(i).TypeGroup(s$) Then GoTo errgr
                            Else
                                    MyRead = False: MissType: Exit Function
                            End If
                            End If
                        ElseIf TypeOf var(i) Is mHandler Then
                            Set usehandler = var(i)
                            If TypeOf usehandler.objref Is mArray Then
                                If Not Fast2Varl(rest$, "–…Õ¡ ¡”", 7, "ARRAY", 5, 7, ff) Then MyRead = False: MissType: Exit Function
                            ElseIf TypeOf usehandler.objref Is FastCollection Then
                                If Not Fast2Varl(rest$, " ¡‘¡”‘¡”«", 9, "INVENTORY", 9, 9, ff) Then
                                     If Not Fast2Varl(rest$, "À…”‘¡", 5, "LIST", 4, 5, ff) Then
                                        If Not Fast2Varl(rest$, "œ’—¡", 4, "QUEUE", 5, 5, ff) Then
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
                                If Not Fast2Varl(rest$, "”Ÿ—œ”", 5, "STACK", 5, 5, ff) Then MyRead = False: MissType: Exit Function
                            ElseIf TypeOf usehandler.objref Is MemBlock Then
                                If Not Fast2Varl(rest$, "ƒ…¡—»—Ÿ”«", 9, "BUFFER", 6, 9, ff) Then MyRead = False: MissType: Exit Function
                            ElseIf usehandler.t1 = 4 Then
                              If Not FastType(rest$, usehandler.objref.EnumName) Then MyRead = False: MissType: Exit Function
                   
                            Else
                                p = IsLabel(bstack, rest$, (what$)) ' just throw any name
                            End If
                        ElseIf TypeOf var(i) Is lambda Then
                            If Not Fast2Varl(rest$, "À¡Ãƒ¡", 5, "LAMBDA", 6, 6, ff) Then MyRead = False: MissType: Exit Function
                        
                        ElseIf TypeOf var(i) Is mEvent Then
                            If Not Fast2Varl(rest$, "√≈√œÕœ”", 7, "EVENT", 5, 5, ff) Then MyRead = False: MissType: Exit Function
                        ElseIf TypeOf var(i) Is Constant Then
                            
                            p = var(i)
                            If var(i).vType Then
                                If Not Fast2Varl(rest$, "¡‘’–œ”", 6, "VARIANT", 7, 7, ff) Then MyRead = False: MissType: Exit Function
                                GoTo jumpref02
                            End If
                            GoTo checkconstant
                        Else
                            p = IsLabel(bstack, rest$, (what$)) ' just throw any name
                        End If
                    End If
                End If
jumpref01:
            If Not LinkGroup(bstack, what$, var(i)) Then
jumpref02:
            If f Then

                If Not ReboundVar(bstack, what$, i) Then globalvar what$, i, True
            Else
    
                globalvar what$, i, True, UseType:=varhash.vType(varhash.index)
                If VarTypeName(var(i)) = "lambda" Then
islambda:
                
                    If ohere$ = vbNullString Then
                        GlobalSub what$ + "()", "", , , i
                    Else
                        GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , i
                    End If
                ElseIf VarTypeName(var(i)) = "Constant" Then
                    If var(i).flag Then GoTo islambda
                End If
                
            End If
        Else
            it = globalvar(what$, it)
            MakeitObject2 var(it)
            If var(i).IamApointer Then
            If var(i).link.IamFloatGroup Then
               Set var(it).LinkRef = var(i).link
                var(it).IamApointer = True
                var(it).isRef = True
            Else
            With var(i).link
            
                var(it).edittag = .edittag
                var(it).FuncList = .FuncList
                var(it).GroupName = myUcase(what$) + "."
                Set var(it).Sorosref = .soros.Copy
                var(it).HasValue = .HasValue
                var(it).HasSet = .HasSet
                var(it).HasStrValue = .HasStrValue
                var(it).HasParameters = .HasParameters
                var(it).HasParametersSet = .HasParametersSet
                var(it).HasRemove = .HasRemove
                        Set var(it).Events = .Events
            
                var(it).highpriorityoper = .highpriorityoper
                var(it).HasUnary = .HasUnary
                If Len(here$) > 0 Then
                var(it).Patch = here$ + "." + what$
                Else
                var(it).Patch = what$
                End If
                Set var(it).mytypes = .mytypes
            End With
            End If
            
            Else
            With var(i)
            
                var(it).edittag = .edittag
                var(it).FuncList = .FuncList
                var(it).GroupName = myUcase(what$) + "."
                Set var(it).Sorosref = .soros.Copy
                var(it).HasValue = .HasValue
                var(it).HasSet = .HasSet
                var(it).HasStrValue = .HasStrValue
                var(it).HasParameters = .HasParameters
                var(it).HasParametersSet = .HasParametersSet
                var(it).HasRemove = .HasRemove
                        Set var(it).Events = .Events
            
                var(it).highpriorityoper = .highpriorityoper
                var(it).HasUnary = .HasUnary
                If Len(here$) > 0 Then
                var(it).Patch = here$ + "." + what$
                Else
                var(it).Patch = what$
                End If
                Set var(it).mytypes = .mytypes
            End With
             var(it).IamRef = Len(bstack.UseGroupname) > 0
             End If
            If var(i).HasStrValue Then
                globalvar what$ + "$", it, True
            End If
            
        End If
        MyRead = True
    End If
     Else
        If Left$(s$, 1) = "#" Then  ' for copy in
            s$ = Mid$(s$, 2)
            If IsNumberNew(bstack, (s$), p, False) Then
                ss$ = "_" + Str$(var2used)
                If bstack.SubLevel > 0 Then
                    MyEr "not for Read statement", "¸˜È „È· ÙÁÌ ƒÈ‹‚·ÛÂ"
                    MyRead = False
                    Exit Function
                End If
                i = globalvar(ss$, 0#, , True)
                If bstack.IamChild Then FeedCopyInOut bstack.Parent, s$, i, ""
                UnFloatGroup bstack, ss$, i, bstack.lastobj, True, , True
                bstack.lastobj.ToDelete = True
                GoTo conthereifglobal
            End If
        Else
            If GetVar3(bstack, s$, i, True, , , , , , True) Then GoTo conthereifglobal
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
i = AllocVar()
Set var(i) = myobject
Set myobject = Nothing
GoTo backfromstr
End If
GoTo er103
Else
If Not GetVar3(bstack, what$, i, , , flag, , , , True) Then i = globalvar(what$, 0)
If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
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
CreateFormOtherObject var(i), myobject
 MyRead = True
Set myobject = Nothing
Else
    GoTo noref01
    End If
    End If

Case 3, 4
    If bs.IsString(s$) Then
        If GetGlobalVar(s$, i) Then
            If flag2 Then
            If Not f Then
                If Not flag Then
                    If ohere$ <> "" Then
                            'GoTo contpush12
               If MyIsObject(var(i)) Then
                    If Typename(var(i)) = "lambda" Then
                        If ohere$ = vbNullString Then
                            GlobalSub what$ + "()", "", , , i
                        Else
                            GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , i
                        End If
                        globalvar what$, i, True
                    ElseIf Typename(var(i)) = mGroup Then
                        what$ = Left$(what$, Len(what$) - 1)
                        GoTo backfromstr
                    Else
                        globalvar what$, i, True
                    End If
                Else
                     globalvar what$, i, True
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
                    If Not ReboundVar(bstack, what$, i) Then globalvar what$, i, True
                Else
contherestr:

            If MyIsObject(var(i)) Then
                    If Typename(var(i)) = "lambda" Then
                         If ohere$ = vbNullString Then
                             GlobalSub what$ + "()", "", , , i
                         Else
                             GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , i
                         End If
                         globalvar what$, i, True
                    ElseIf Typename(var(i)) = mGroup Then
                         what$ = Left$(what$, Len(what$) - 1)
                         GoTo backfromstr
                    Else
                         globalvar what$, i, True
                     End If
                Else
                globalvar what$, i, True
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
                        i = Len(s$)
                        If i > 0 Then
                            If varhash.Find(Left$(s$, i - 1), i) Then sbf(bstack.IndexSub).tpointer = i
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
                        i = Len(s$)
                        If i > 0 Then
                            If varhash.Find(Left$(s$, i - 1), i) Then sbf(bstack.IndexSub).tpointer = i
                        Else
                            sbf(bstack.IndexSub).tpointer = 0
                        End If
                    End If
                End If
                MyRead = True
            End If
        Else
            i = CopyArrayItemsNoFormated(bstack, s$)
            If i <> 0 Then
                If Not FastSymbol(rest$, ")") Then GoTo er107
                If f And i > 0 Then '' look about f - work for refer but no refer can be done...why???
                    If Not ReboundArr(bstack, what$, i) Then GoTo arrconthere
                Else
arrconthere:
                what$ = myUcase(what$)
                If ohere$ = vbNullString Then
                    If varhash.ExistKey(what$) Then
                      If flag2 And Not f And Not flag Then
                        If i < 0 Then i = -i
                           varhash.ItemCreator what$, i
                           ' what$ now is empty string
                    Else
                        GoTo er108
                        End If
                    Else
                        varhash.ItemCreator what$, i
                    End If
                Else
                    If varhash.ExistKey(ohere$ + "." + what$) Then
                    If flag2 And Not f And Not flag Then
                        i = Abs(i)
                           varhash.ItemCreator ohere$ + "." + what$, i, True, True
                    Else
                        GoTo er108
                    End If
                    Else
                      i = Abs(i)
                    varhash.ItemCreator ohere$ + "." + what$, i, True, True
                    End If
                End If

            End If
            MyRead = True
        Else
        ' get function
        If GetSub(s$, i) Then
        If Len(sbf(i).sbgroup) > 0 Then
        If sbf(i).Extern > 0 Then
        s$ = Left$(what$, Len(what$) - 1) + " {CALL EXTERN" + Str$(sbf(i).Extern) + "'" + ChrW(&H1FFD) + "}" + sbf(i).sbgroup
        Else
        s$ = Left$(what$, Len(what$) - 1) + " {" + sbf(i).sb + "}" + sbf(i).sbgroup
        End If
        Else
        If sbf(i).Extern > 0 Then
        s$ = Left$(what$, Len(what$) - 1) + " {CALL EXTERN" + Str$(sbf(i).Extern) + "'" + ChrW(&H1FFD) + "}"
        Else
        s$ = Left$(what$, Len(what$) - 1) + " {" + sbf(i).sb + "}"
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
'If x1 <> 0 Then
'    If what$ <> myUcase(what$) Then Stop
'End If
Select Case x1
Case 1
    If bs.IsObjectRef(myobject) Then
        MyRead = True
        If flag2 Then
        GoTo comehere
        ElseIf flag Then
            p = 0#
           i = globalvar(what$, p)
           GoTo contread123
        End If
       If GetVar3(bstack, what$, i, , , flag, s$, checktype, isAglobal, True, ok) Then
        If isAglobal And Not allowglobals Then
        GoTo comehere
        End If
contread123:
                If Typename$(myobject) = VarTypeName(var(i)) Then
                    If Typename$(var(i)) = mGroup Then
                                 ss$ = bstack.GroupName
                                 If Len(s$) > 0 Then what$ = s$
                                 If Len(var(i).GroupName) > Len(what$) Then
                                    ff = 1
                                    If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                                   If FastSymbol(rest$, "*") Then
                                        If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
                                            If s$ = "POINTER" Then
                                                If Not myobject.IamApointer Then GoTo errgr
                                            ElseIf s$ = "ƒ≈… ‘«”" Then
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
                                            ElseIf s$ = "ƒ≈… ‘«”" Then
                                                If Not myobject.IamApointer Then GoTo errgr
                                            ElseIf s$ = "GROUP" Then
                                                ' get pointer too
                                            ElseIf s$ = "œÃ¡ƒ¡" Then
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
                                    If var(i).IamRef Then
                                        SwapStrings s$, here$
                                        here$ = vbNullString
                                        UnFloatGroupReWriteVars bstack, what$, i, myobject
                                        SwapStrings here, s$
                                        s$ = vbNullString
                                    Else
                                        UnFloatGroupReWriteVars bstack, what$, i, myobject
                                        myobject.ToDelete = True
                                    End If
                                Else
                                    If Len(var(i).Patch) > 0 Then what$ = var(i).Patch
                                    bstack.GroupName = Left$(what$, Len(what$) - Len(var(i).GroupName) + 1)
                                    If Len(var(i).GroupName) > 0 Then
                                        ff = 1
                                        If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                                            If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
                                            If Not myobject.TypeGroup(s$) Then GoTo errgr
                                        End If
                                        what$ = Left$(var(i).GroupName, Len(var(i).GroupName) - 1)
                                        SwapStrings s$, here$
                                        here$ = vbNullString
                                        UnFloatGroupReWriteVars bstack, what$, i, myobject, , , ByPass
                                        myobject.ToDelete = True
                                        ByPass = False
                                        SwapStrings here, s$
                                        s$ = vbNullString
                                    ElseIf var(i).IamApointer And myobject.IamApointer Then
                                    ff = 1
                                    If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
    
                                    FastSymbol rest$, "*"
      
                                    If FastPureLabel(rest$, s$, , True) <> 1 Then SyntaxError: MyRead = False: Exit Function
      
                                    If myobject.link.IamFloatGroup Then
                                    
                                    If Not myobject.link.TypeGroup(s$) Then
                                        If s$ = "POINTER" Then
                                        ElseIf s$ = "ƒ≈… ‘«”" Then
                                        Else
                                        GoTo errgr
                                        End If
                                    End If
                                    Else
                                        If Len(myobject.lasthere) = 0 Then
                                            If GetVar(bstack, myobject.GroupName, it, True) Then
                                                If Not var(i).TypeGroup(s$) Then GoTo errgr
                                            End If
                                        ElseIf GetVar(bstack, myobject.lasthere + "." + myobject.GroupName, it, True) Then
                                                If Not var(i).TypeGroup(s$) Then GoTo errgr
                                        End If
                                    End If
                                    
                                    End If
                                    Set var(i) = myobject
                                    
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
    
                   
                    ElseIf Typename$(var(i)) = mHdlr Then
                        Set usehandler = myobject
                        Set usehandler1 = var(i)
                        If usehandler1.ReadOnly Then
                            MyRead = False
                           ReadOnly
                           Exit Function
                           
                        ElseIf usehandler1.t1 = usehandler.t1 Then
                            If usehandler1.t1 = 4 Then
                                If usehandler1.objref Is usehandler.objref Then
                                    Set var(i) = myobject
                                ElseIf usehandler1.objref.EnumName = usehandler.objref.EnumName Then
                                    If usehandler1.objref.ExistFromOther(usehandler.index_cursor) Then
                                        Set usehandler.objref = usehandler1.objref
                                        Set var(i) = usehandler
                                    Else
                                        GoTo er103
                                    End If
                                Else
                                    GoTo er103
                                End If
                            Else
                            
                                Set var(i) = myobject
                           End If
                        ElseIf usehandler1.t1 <> 4 And usehandler.t1 = 3 Then
                           Set var(i) = myobject
                        Else
                          GoTo er103
                        End If
                    Else
                    Set var(i) = myobject
                    End If
                    If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                    
                    If Not FastPureLabel(rest$, s$) = 1 Then
                        GoTo er110
                    End If
                    ' no second time
                    If Not myobject Is Nothing Then
                    If TypeOf myobject Is mHandler Then
                        
                        Set usehandler = myobject
                        s$ = myUcase(s$)
                        If usehandler.t1 = 1 Then
                        ff = 0
                        
                        If Not Fast2Varl(s$, " ¡‘¡”‘¡”«", 9, "INVENTORY", 9, 9, ff) Then
                        If usehandler.objref.IsQueue Then
                            If Not Fast2Varl(s$, "œ’—¡", 4, "QUEUE", 5, 5, ff) Then GoTo er103
                        Else
                            If Not Fast2Varl(s$, "À…”‘¡", 5, "LIST", 4, 5, ff) Then GoTo er103
                        End If
                        End If
                        ElseIf usehandler.t1 = 3 Then
                        ff = 0
                        If Fast2Varl(s$, "–…Õ¡ ¡”", 7, "ARRAY", 5, 7, ff) Then
                            If Not CheckIsmArray(myobject) Then GoTo er103
                        ElseIf Fast2Varl(s$, "”Ÿ—œ”", 5, "STACK", 5, 5, ff) Then
                            If Not CheckIsmStiva(myobject) Then GoTo er103
                        Else
                            GoTo er103
                        End If
                        ElseIf usehandler.t1 = 4 Then
                        Else
                        GoTo er103
                        End If
                    End If
                    End If
                    
                    End If
                ElseIf x1 = 1 And CheckIsmArray(myobject) Then
               
                    ''bstack.lastobj.CopyArray pppp
                    Set usehandler = New mHandler
                    Set var(i) = usehandler
                    usehandler.t1 = 3
                    Set usehandler.objref = myobject
                    Set usehandler = Nothing
                    If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                        If Not Fast2Varl(rest$, "–…Õ¡ ¡”", 7, "ARRAY", 5, 7, ff) Then GoTo er103
                    End If
                    Set myobject = Nothing
                ElseIf MyIsNumeric(var(i)) Then
                If Not myobject Is Nothing Then
                If TypeOf myobject Is Group Then
                If myobject.IamApointer Then
                Set var(i) = myobject
                GoTo cont10101
                End If
                End If
                End If
                GoTo er103
                Else
                     MyRead = False
                     If VarTypeName(var(i)) = "Nothing" Then
                     MissingObjRef
                     Else
                    GoTo er103
                    End If
                    Exit Function
                    
                End If
cont10101:
        Else
        If i = -1 Then
        ff = 0
        If TypeOf myobject Is mHandler Then
        Set usehandler = myobject
        If Fast2Varl(rest$, "Ÿ”", 2, "AS", 2, 2, ff) Then
                If usehandler.t1 = 1 Then
                    If Not Fast2Varl(rest$, " ¡‘¡”‘¡”«", 9, "INVENTORY", 9, 9, ff) Then
                         If Not Fast2Varl(rest$, "À…”‘¡", 5, "LIST", 4, 5, ff) Then
                            If Not Fast2Varl(rest$, "œ’—¡", 4, "QUEUE", 5, 5, ff) Then
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
                ElseIf usehandler.t1 = 2 Then
                    If Not Fast2Varl(rest$, "ƒ…¡—»—Ÿ”«", 9, "BUFFER", 6, 9, ff) Then
                            WrongObject
                            MyRead = False
                            Exit Function
                    End If
                ElseIf usehandler.t1 = 3 Then
                    If Not Fast2Varl(rest$, "–…Õ¡ ¡”", 7, "ARRAY", 5, 7, ff) Then
                    If Not Fast2Varl(rest$, "”Ÿ—œ”", 5, "STACK", 5, 5, ff) Then
                    
                            WrongObject
                            MyRead = False
                            Exit Function
                    End If
                    End If
                    
                ElseIf usehandler.t1 = 4 Then
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
        If Fast2Varl(rest$, "Ÿ”", 2, "AS", 2, 2, ff) Then
            If Fast2Varl(rest$, "ƒ≈… ‘«”", 7, "POINTER", 7, 7, ff) Then GoTo checkpointer
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
        If FastPureLabel(rest$, ss$, , True) = 1 Then
        If check2(ss$, "Ÿ”", "AS") Then GoTo er110
        End If
        WrongObject
        MyRead = False
        Exit Function
        End If
        
comehere:
        i = globalvar(what$, 0)
        it = varhash.index
      
        If Typename$(myobject) = mGroup Then
            If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                If Fast2Varl(rest$, "ƒ≈… ‘«”", 7, "POINTER", 7, 7, ff) Then
                    If Not myobject.IamApointer Then
errgr:
                          WrongObject
                          MyRead = False
                          Exit Function
                    End If
                    GoTo contpointer
                ElseIf Not Fast2Varl(rest$, "œÃ¡ƒ¡", 5, "GROUP", 5, 5, ff) Then
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
                                    If GetVar(bstack, myobject.GroupName, i, True) Then
                                        If Not var(i).TypeGroup(s$) Then GoTo errgr
                                        CopyGroup2 var(i), bstack
                                        GoTo oop1
                                    Else
                                        GoTo errgr
                                    End If
                                ElseIf GetVar(bstack, myobject.lasthere + "." + myobject.GroupName, i, True) Then
                                    If Not var(i).TypeGroup(s$) Then GoTo errgr
                                    CopyGroup2 var(i), bstack
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
                                UnFloatGroup bstack, bstack.GroupName + what$, i, myobject, here$ = vbNullString Or Len(bstack.UseGroupname) > 0
                            Else
                                UnFloatGroup bstack, bstack.GroupName + what$, i, myobject.link, here$ = vbNullString Or Len(bstack.UseGroupname) > 0
                            End If
                        Else
                            CopyPointerRef bstack, myobject
oop1:
                            UnFloatGroup bstack, bstack.GroupName + what$, i, bstack.lastobj, here$ = vbNullString Or Len(bstack.UseGroupname) > 0
                            bstack.lastobj.ToDelete = True
                        End If
                    Else
                        UnFloatGroup bstack, bstack.GroupName + what$, i, myobject, here$ = vbNullString Or Len(bstack.UseGroupname) > 0, , True
                        myobject.ToDelete = True
                    End If
                    
                Else
contpointer:
                    If myobject.IamApointer Then
                        Set var(i) = myobject
                    Else
                        UnFloatGroup bstack, bstack.GroupName + what$, i, myobject, Not (here$ = vbNullString Xor Len(bstack.UseGroupname) > 0), , True
                        myobject.ToDelete = True
                    End If
               End If
               ' var(i).IamRef = Len(bstack.UseGroupname) > 0
            ElseIf Typename$(myobject) = "mEvent" Then
                If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                If Fast2Varl(rest$, "¡‘’–œ”", 6, "VARIANT", 7, 7, ff) Then
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
                ElseIf Not Fast2Varl(rest$, "√≈√œÕœ”", 7, "EVENT", 5, 5, ff) Then
                        WrongObject
                        Exit Function
                    End If
                End If
                Set var(i) = myobject
            ElseIf Typename$(myobject) = "lambda" Then
                    If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                        If Fast2Varl(rest$, "¡‘’–œ”", 6, "VARIANT", 7, 7, ff) Then
                           If FastSymbol(rest$, "=") Then
                            optlocal = Not useoptionals: useoptionals = True
                            If Not IsNumberD2(rest$, (p), False) Then
                            If Not ISSTRINGA(rest$, s$) Then
                                SyntaxError
                                Exit Function
                            End If
                            End If
                        End If
                    ElseIf Not Fast2Varl(rest$, "À¡Ãƒ¡", 5, "LAMBDA", 6, 6, ff) Then
                        WrongObject
                        Exit Function
                    End If
                End If
                Set var(i) = myobject
                If ohere$ = vbNullString Then
                    GlobalSub what$ + "()", "", , , i
                Else
                    GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , i
                End If
            ElseIf Typename$(myobject) = mHdlr Then
                Set usehandler1 = myobject
                If MyIsObject(var(i)) Then
                If TypeOf var(i) Is mHandler Then
                    Set usehandler = var(i)
                        If usehandler.ReadOnly Then
                            ReadOnly
                            MyRead = False
                            Exit Function
                        End If
                    End If
                End If
                If usehandler1.indirect > -1 Then
                    If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                        If IsLabel(bstack, rest$, ss$) = 0 Then
                            GoTo er110
                        End If
                        If LCase(Typename(var(usehandler1.indirect))) <> LCase(ss$) Then
                            WrongObject
                            MyRead = False
                        Exit Function
                End If
             End If
                Set var(i) = MakeitObjectGeneric(usehandler1.indirect)
                Else
                If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
      
                If usehandler1.t1 = 1 Then
                    If Fast2Varl(rest$, "¡‘’–œ”", 6, "VARIANT", 7, 7, ff) Then
                    GoTo jump0001233
                    ElseIf Not Fast2Varl(rest$, " ¡‘¡”‘¡”«", 9, "INVENTORY", 9, 9, ff) Then
                         If Not Fast2Varl(rest$, "À…”‘¡", 5, "LIST", 4, 5, ff) Then
                            If Not Fast2Varl(rest$, "œ’—¡", 4, "QUEUE", 5, 5, ff) Then
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
                ElseIf usehandler1.t1 = 2 Then
                    If Not Fast2Varl(rest$, "ƒ…¡—»—Ÿ”«", 9, "BUFFER", 6, 9, ff) Then
                        If Not Fast2Varl(rest$, "¡‘’–œ”", 6, "VARIANT", 7, 7, ff) Then
                                WrongObject
                                MyRead = False
                                Exit Function
                        Else
                            GoTo jump0001233
                        End If
                    End If
                ElseIf usehandler1.t1 = 3 Then
                    
                    If Fast2Varl(rest$, "–…Õ¡ ¡”", 7, "ARRAY", 5, 7, ff) Then
                        If Not CheckIsmArray(myobject) Then GoTo er103
                        Set usehandler = New mHandler
                        Set usehandler.objref = myobject
                        usehandler.t1 = 3
                        Set myobject = usehandler
                        Set usehandler = Nothing
                        Set usehandler1 = Nothing
                    ElseIf Fast2Varl(rest$, "”Ÿ—œ”", 5, "STACK", 5, 5, ff) Then
                        If Not CheckIsmStiva(myobject) Then GoTo er103
                        Set usehandler = New mHandler
                        Set usehandler.objref = myobject
                        usehandler.t1 = 3
                        Set myobject = usehandler
                        Set usehandler = Nothing
                        Set usehandler1 = Nothing
                    ElseIf Fast2Varl(rest$, "¡‘’–œ”", 6, "VARIANT", 7, 7, ff) Then
                        If Not CheckIsmArray(myobject) Then
                        If Not CheckIsmStiva(myobject) Then
                            GoTo er103
                        End If
                        End If
                        Set usehandler = New mHandler
                        Set usehandler.objref = myobject
                        usehandler.t1 = 3
                        Set myobject = usehandler
                        Set usehandler = Nothing
                        Set usehandler1 = Nothing
                    GoTo jump0001233
                    ElseIf Typename(usehandler1.objref) = mHdlr Then
                        Set usehandler1 = usehandler1.objref
                        If usehandler1.t1 = 4 Then
                            If FastType(rest$, usehandler1.objref.EnumName) Then
                                Set usehandler = New mHandler
                                usehandler.t1 = 4
                                Set myobject = usehandler1.objref
                                usehandler.index_cursor = myobject.Value
                                Set usehandler.objref = myobject
                                usehandler.index_start = myobject.index
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
                    
                ElseIf usehandler1.t1 = 4 Then
                    If FastPureLabel(rest$, s$, , True) = 1 Then
                    If Not s$ = myUcase(usehandler1.objref.EnumName, True) Then
                    If GetSub(s$ + "()", i) Then
                    If sbf(i).IamAClass Then
                        GoTo er113
                    End If
                    ElseIf GetSub(bstack.GroupName + s$ + "()", i) Then
                    If sbf(i).IamAClass Then
                        GoTo er113
                    End If
                    End If
                        GoTo er112
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
                End If
t14:
                Set var(i) = myobject
                End If
            ElseIf Typename$(myobject) = myArray Then
                   
                   If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                    If Not Fast2Varl(rest$, "–…Õ¡ ¡”", 7, "ARRAY", 5, 7, ff) Then
                        WrongObject
                        Exit Function
                    End If
            End If
                Set usehandler = New mHandler
                Set var(i) = usehandler
                usehandler.t1 = 3
                Set usehandler.objref = myobject
                Set usehandler = Nothing
            Else
             If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                    If Fast2Varl(rest$, "¡‘’–œ”", 6, "VARIANT", 7, 7, ff) Then
                        GoTo jump0001233
                    End If
                    If Not Fast2Varl(rest$, "ƒ≈… ‘«”", 7, "POINTER", 7, 7, ff) Then
                    ElseIf IsLabel(bstack, rest$, ss$) = 0 Then
                        GoTo er110
                    End If
                
                If LCase(Typename(myobject)) <> LCase(ss$) Then
                        WrongObject
                        MyRead = False
                        Exit Function
                End If
            End If
                Set var(i) = myobject
            
            End If
            
            
            
            
            Set myobject = Nothing
        End If
    ElseIf bs.IsNumber(p) Then
contStr1:
    ihavetype = True
    If Not lookOne(rest$, ",") Then
    ' FF used again
    
        If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
conthereEnum:
            ihavetype = True
            If Not FastPureLabel(rest$, s$, , , True) = 1 Then
            SyntaxError
            Exit Function
            End If
            ss$ = myUcase(s$, AscW(s$) > 255)
            Select Case ss$
            Case "¡—…»Ãœ”", "DECIMAL"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                p = CDec(p)
            Case "ƒ…–Àœ”", "DOUBLE"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                p = CDbl(p)
            Case "¡–Àœ”", "SINGLE"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                p = CSng(p)
            Case "Àœ√… œ”", "BOOLEAN"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                p = CBool(p)
            Case "Ã¡ —’”", "LONG"
                If Fast2Varl(rest$, "Ã¡ —’”", 6, "LONG", 4, 6, ff) Then
                    If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                    p = cInt64(p)
                Else
                    If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                    p = CLng(p)
                End If
            Case "¡ ≈—¡…œ”", "INTEGER"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                p = CInt(p)
            Case "Àœ√…”‘… œ”", "CURRENCY"
                If FastSymbol(rest$, "=") Then optlocal = Not useoptionals: useoptionals = True: If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
                p = CCur(p)
            Case "√—¡ÃÃ¡", "STRING"
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
            Case "¡‘’–œ”", "VARIANT"
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
            Case Else
                ss$ = s$
                it = True
                  If MyIsNumeric(p) Then x = p: it = False
                  If IsEnumAs(bstack, ss$, p, ok, rest$) Then
                    If Not it Then
                    
                        Set usehandler = p
                        p = x
                        Set usehandler = usehandler.objref.SearchValue(p, ok)
                        Set myobject = usehandler
                        If ok Then
                            Set p = myobject
                            
                            ' GoTo loopcont123
                        Else
                            ExpectedEnumType
                            MyRead = False
                            Exit Function
                        End If
                    
                    End If
                  Else
messnotype:
                  MyEr "No type [" + s$ + "] found", "‰ÂÌ ‚ÒﬁÍ· Ù˝Ô [" + s$ + "]"
                  Exit Function
                  End If
            End Select
        ElseIf FastSymbol(rest$, "=") Then
            If Not IsNumberD2(rest$, (p)) Then missNumber: Exit Function
        End If
    End If
        MyRead = True
        If flag2 Then
            globalvar what$, p
        ElseIf GetVar3(bstack, what$, i, , , flag, , checktype, isAglobal, True, ok) Then
            ihavetype = False
            If isAglobal And Not allowglobals Then
                 globalvar what$, p
            ElseIf MyIsObject(var(i)) Then
                If var(i) Is Nothing Then
                    MissingObjRef
                ElseIf TypeOf var(i) Is Group Then
                    If var(i).HasSet Then
                       
                        Set m = bstack.soros
                        Set bstack.Sorosref = New mStiva
                        bstack.soros.PushVal p
                        NeoCall2 bstack, what$ + "." + ChrW(&H1FFF) + ":=()", ok
                        Set bstack.Sorosref = m
                        Set m = Nothing
                    Else
                        GoTo there18274
                    End If
                Else
there18274:
                    If TypeOf var(i) Is Constant Then
                        CantAssignValue
                    ElseIf TypeOf var(i) Is mHandler Then
                        Set usehandler = var(i)
                        If usehandler.t1 <> 4 Then GoTo er104
                        If MyIsObject(p) Then
                            If Not TypeOf p Is mHandler Then
                                    ' errorhere
                                End If
                                Set usehandler1 = p
                                If usehandler.objref Is usehandler1.objref Then
                                    Set myobject = usehandler1
                                Else
                                    p = Empty
                                    If Not usehandler1.t1 = 4 Then
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
                                    Set var(i) = myobject
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
                                    If VarType(var(i)) <> VarType(p) Then
                                        GoTo er109
                                    ElseIf AssignTypeNumeric(p, VarType(var(i))) Then
                                        var(i) = p
                                    Else
                                        GoTo er105
                                    End If
                                ElseIf AssignTypeNumeric(p, VarType(var(i))) Then
                                    var(i) = p
                                Else
                                    GoTo er105
                                End If
                            Else
                                var(i) = p
                            End If
                        End If
                    ElseIf i = -1 Then
                        If ok Then
                           bstack.SetVarobJvalue what$, p
                        Else
                            bstack.SetVar what$, p
                        End If
                    Else
                        globalvar what$, p, UseType:=ihavetype
                    End If
cont112233:
                    p = 0#
                ElseIf bs.IsOptional Then
                    MyRead = True
                    If Not lookOne(rest$, ",") Then
                        checktype = True
                        If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
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
                        Case "¡—…»Ãœ”", "DECIMAL"
                                p = CDec(0)
                        Case "ƒ…–Àœ”", "DOUBLE"
                                p = 0#
                        Case "¡–Àœ”", "SINGLE"
                                p = 0!
                        Case "Àœ√… œ”", "BOOLEAN"
                                p = False
                        Case "Ã¡ —’”", "LONG"
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
                                
                        Case "¡ ≈—¡…œ”", "INTEGER"
                                p = 0
                        Case "Àœ√…”‘… œ”", "CURRENCY"
                            p = 0@
                        Case "√—¡ÃÃ¡", "STRING"
                            p = vbNullString
                        Case "¡‘’–œ”", "VARIANT"
                            p = Empty
                            ihavetype = False
                            checktype = False
                        Case Else
cont234356:
                        
                            If Not flag2 And GetVar3(bstack, what$, i, , , flag, , , isAglobal, True, ok) Then
                                If isAglobal Then
                                    GoTo er110
                                Else
                                    If MyIsNumeric(var(i)) Then
                                        If IsEnumAs(bstack, s$, p, ok, rest$) Then
                                            If ok Then Set var(i) = p
                                        End If
                                    ElseIf TypeOf var(i) Is Group Then
                                        If Not ss$ = "œÃ¡ƒ¡" Then
                                            If Not ss$ = "GROUP" Then
                                                If var(i).IamApointer Then
                                                    If Not ss$ = "ƒ≈… ‘«”" Then
                                                        If Not ss$ = "POINTER" Then
                                                            If Not var(i).link.TypeGroup(ss$) Then
                                                                MyRead = False: MissType: Exit Function
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    MyRead = False: MissType: Exit Function
                                                End If
                                            End If
                                        End If
                                    ElseIf TypeOf var(i) Is mHandler Then
                                        Set usehandler = var(i)
                                        If TypeOf usehandler.objref Is mArray Then
                                            If Not ss$ = "–…Õ¡ ¡”" Then
                                                If Not ss$ = "ARRAY" Then
                                                    MyRead = False: MissType: Exit Function
                                                End If
                                            End If
                                        ElseIf TypeOf usehandler.objref Is FastCollection Then
                                            If Not ss$ = " ¡‘¡”‘¡”«" Then
                                                If Not ss$ = "INVENTORY" Then
                                                    If Not ss$ = "À…”‘¡" Then
                                                        If Not ss$ = "LIST" Then
                                                            If Not ss$ = "œ’—¡" Then
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
                                            If Not ss$ = "ƒ…¡—»—Ÿ”«" Then
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
                                            Set usehandler = var(i)
                                        Else
                                            p = IsLabel(bstack, rest$, (what$)) ' just throw any name
                                        End If
                                        Set usehandler = Nothing
                                    ElseIf TypeOf var(i) Is lambda Then
                                            If Not ss$ = "À¡Ãƒ¡" Then
                                                If Not ss$ = "LAMBDA" Then
                                                    MyRead = False: MissType: Exit Function
                                                End If
                                            End If
                                    ElseIf TypeOf var(i) Is mEvent Then
                                            If Not ss$ = "√≈√œÕœ”" Then
                                                If Not ss$ = "EVENT" Then
                                                    MyRead = False: MissType: Exit Function
                                                End If
                                            End If
                                    ElseIf TypeOf var(i) Is Constant Then
                                        p = var(i)
                                        GoTo checkconstant
                                    ElseIf IsEnumAs(bstack, (s$), p, ok, rest$) Then
                                        If ok Then
                                            CheckItemType bstack, var(i), vbNullString, ss$, , ok
                                            If s$ = ss$ Then
                                                Set var(i) = p
                                            End If
                                        End If
                                       ' just throw any name
                                    End If
                                    GoTo loopcont123
                                End If
                                ' never pass from here
                            Else
                                If i = -1 Then
                                    If Len(ss$) = 0 Then
                                        FastPureLabel rest$, s$, , , True
                                        
                                        ss$ = myUcase(s$, True)
                                    End If
                                    bstack.ReadVar what$, p
                                    Set myobject = p
                                    Set usehandler = myobject
                                    ff = 0
                                    If Not check2(ss$, "ƒ≈… ‘«”", "POINTER") Then
                                        If usehandler.t1 = 1 Then
                                            If Not check2(ss$, " ¡‘¡”‘¡”«", "INVENTORY") Then
                                                If Not check2(ss$, "À…”‘¡", "LIST") Then
                                                    If Not check2(ss$, "œ’—¡", "QUEUE") Then
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
                                        ElseIf usehandler.t1 = 2 Then
                                            If Not check2(ss$, "ƒ…¡—»—Ÿ”«", "BUFFER") Then
                                                WrongObject
                                                MyRead = False
                                                Exit Function
                                            End If
                                        ElseIf usehandler.t1 = 3 Then
                                            If Not check2(ss$, "–…Õ¡ ¡”", "ARRAY") Then
                                                If Not check2(s$, "”Ÿ—œ”", "STACK") Then
                                                    WrongObject
                                                    MyRead = False
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                    ElseIf usehandler.t1 = 4 Then
                                        If Not FastType(s$, usehandler.objref.EnumName) Then
                                            p = usehandler.index_cursor
                                            Set usehandler = Nothing
                                            Set myobject = Nothing
                                            GoTo conthereEnum
                                        End If
                                    End If
                                    GoTo loopcont123
                                ElseIf check2(ss$, "œÃ¡ƒ¡", "GROUP") Then
                                useoptionals = False
                                ElseIf check2(ss$, "–…Õ¡ ¡”", "ARRAY") Then
                                useoptionals = False
                                ElseIf check2(ss$, " ¡‘¡”‘¡”«", "INVENTORY") Then
                                useoptionals = False
                                ElseIf check2(ss$, "ƒ…¡—»—Ÿ”«", "BUFFER") Then
                                useoptionals = False
                                ElseIf check2(ss$, "À¡Ãƒ¡", "LAMBDA") Then
                                useoptionals = False
                                ElseIf check2(ss$, "√≈√œÕœ”", "EVENT") Then
                                    useoptionals = False
                                ElseIf check2(ss$, "À…”‘¡", "LIST") Then
                                    useoptionals = False
                                ElseIf check2(ss$, "œ’—¡", "QUEUE") Then
                                    useoptionals = False
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
                            If Not IsNumberD2(rest$, p, False) Then
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
                            Else
                                missNumber
                                Exit Function
                            End If
                            End If
                            optlocal = Not useoptionals: useoptionals = True
                        End If
                        If Len(rest$) > 0 Then
                            If InStr("!@#%~&", Left$(rest$, 1)) > 0 Then
                                Mid$(rest$, 1, 1) = " "
                            End If
                        End If
                    ElseIf FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                Else
                    p = 0#
                End If
                    
cont1459:
                If flag2 Then
                    globalvar what$, p
                    If Not useoptionals Then GoTo err100
                ElseIf GetVar3(bstack, what$, i, , , flag, , , isAglobal, True, ok) Then
                    If isAglobal Then
                        If Not useoptionals Then GoTo err10
                        globalvar what$, p
                    ElseIf ihavetype Then
                        If VarType(p) <> VarType(var(i)) Then
                            If Not AssignTypeNumeric(var(i), VarType(p)) Then
                            MyRead = False
                            Exit Function
                        End If
                    End If
                End If
                ' just skip read for this value
            ElseIf i = -1 Then
                If Not ok Then
                    bstack.SetVar what$, p
                End If
                If Not useoptionals Then GoTo err100
            Else
                If VarType(p) = vbEmpty Then p = 0#
                globalvar what$, p, UseType:=checktype
                 checktype = False
                 If Not useoptionals Then GoTo err10
            End If
            p = 0#
        Else
            If bs.IsString(s$) Then
                p = s$
               If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                    If Not Fast2Varl(rest$, "¡‘’–œ”", 6, "VARIANT", 7, 7, ff) Then
                        If Not Fast2Varl(rest$, "√—¡ÃÃ¡", 6, "STRING", 6, 6, ff) Then
                            GoTo er110
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
        ElseIf GetVar3(bstack, what$, i, , , flag, , checktype, isAglobal, True, ok) Then
        If isAglobal And Not allowglobals Then
            globalvar what$, s$
        ElseIf MyIsObject(var(i)) Then
                If TypeOf var(i) Is Group Then
                        Set m = bstack.soros
                        Set bstack.Sorosref = New mStiva
                        bstack.soros.PushStr s$
                        NeoCall2 bstack, Left$(what$, Len(what$) - 1) + "." + ChrW(&H1FFF) + ":=()", ok
                        Set bstack.Sorosref = m
                        Set m = Nothing
                ElseIf TypeOf var(i) Is Constant Then
                CantAssignValue
                MyRead = False
                Exit Function
                Else
                    CheckVar var(i), s$
                End If
            Else
                var(i) = s$
            End If
        ElseIf i = -1 Then
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
        ElseIf GetVar3(bstack, what$, i, , , flag, , checktype, isAglobal, True) Then
            If isAglobal And Not allowglobals Then
               globalvar what$, s$
                End If
        ElseIf i = -1 Then
      
        Else
            If Not useoptionals Then GoTo err10
            globalvar what$, s$
        End If
    ElseIf bs.IsObjectRef(myobject) Then
        If Typename$(myobject) = "lambda" Then
            If flag2 Then
               i = globalvar(what$, s$)
            ElseIf GetVar3(bstack, what$, i, , , flag) Then
                CheckVar var(i), s$
            Else
                i = globalvar(what$, s$)
            End If
            Set var(i) = myobject
            If ohere$ = vbNullString Then
                GlobalSub what$ + "()", "", , , i
            Else
                GlobalSub ohere$ + "." + bstack.GroupName + what$ + "()", "", , , i
            End If
            MyRead = True
        ElseIf Typename$(myobject) = mGroup Then
        Set bstack.lastobj = myobject
         Set myobject = Nothing
         If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
                    If Not Fast2Varl(rest$, "œÃ¡ƒ¡", 5, "GROUP", 5, 5, ff) Then SyntaxError: Exit Function
        
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
        
        If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
            If Fast2Varl(rest$, "¡—…»Ãœ”", 7, "DECIMAL", 7, 7, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                    p = CDec(p)
            ElseIf Fast2Varl(rest$, "ƒ…–Àœ”", 6, "DOUBLE", 6, 6, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If

                p = CDbl(p)
            ElseIf Fast2Varl(rest$, "¡–Àœ”", 5, "SINGLE", 6, 6, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CSng(p)
            ElseIf Fast2Varl(rest$, "Àœ√… œ”", 7, "BOOLEAN", 7, 7, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CBool(p)
            ElseIf Fast2Varl(rest$, "Ã¡ —’”", 6, "LONG", 4, 6, ff) Then
                    If Fast2Varl(rest$, "Ã¡ —’”", 6, "LONG", 4, 6, ff) Then
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
            ElseIf Fast2Varl(rest$, "¡ ≈—¡…œ”", 8, "INTEGER", 7, 8, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, (p), True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CInt(p)
            ElseIf Fast2Varl(rest$, "Àœ√…”‘… œ”", 10, "CURRENCY", 8, 10, ff) Then
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
        ElseIf GetVar3(bstack, what$, i, , , flag, , checktype, isAglobal, True, ok) Then
            If isAglobal And Not allowglobals Then
                globalvar what$, MyRound(p)
            Else
                var(i) = MyRound(p)
            End If
        ElseIf i = -1 Then
            bstack.SetVar what$, p
        Else
            globalvar what$, MyRound(p)
        End If
    ElseIf bs.IsOptional Then

        MyRead = True
        If Not lookOne(rest$, ",") Then
        If Fast2VarNoTrim(rest$, "Ÿ”", 2, "AS", 2, 3, ff) Then
            If Fast2Varl(rest$, "¡—…»Ãœ”", 7, "DECIMAL", 7, 7, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                    p = CDec(p)
            ElseIf Fast2Varl(rest$, "ƒ…–Àœ”", 6, "DOUBLE", 6, 6, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CDbl(p)
            ElseIf Fast2Varl(rest$, "¡–Àœ”", 5, "SINGLE", 6, 6, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CSng(p)
            ElseIf Fast2Varl(rest$, "Àœ√… œ”", 7, "BOOLEAN", 7, 7, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CBool(p)
            ElseIf Fast2Varl(rest$, "Ã¡ —’”", 6, "LONG", 4, 6, ff) Then
                    If Fast2Varl(rest$, "Ã¡ —’”", 6, "LONG", 4, 6, ff) Then
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
            ElseIf Fast2Varl(rest$, "¡ ≈—¡…œ”", 8, "INTEGER", 7, 8, ff) Then
                    If FastSymbol(rest$, "=") Then
                        If Not IsNumberD2(rest$, p, True) Then missNumber: Exit Function
                        optlocal = Not useoptionals: useoptionals = True
                    End If
                p = CInt(p)
                
            ElseIf Fast2Varl(rest$, "Àœ√…”‘… œ”", 10, "CURRENCY", 8, 10, ff) Then
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
        ElseIf GetVar3(bstack, what$, i, , , flag, , checktype, isAglobal, True) Then
            If isAglobal And Not allowglobals Then
               globalvar what$, p
                End If
        ElseIf i = -1 Then
            
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
            MyEr "Expected group", "–ÂÒﬂÏÂÌ· ÔÏ‹‰·"
            MyRead = False: Exit Function
        Else
             i = 1
            aheadstatus rest$, False, i
            ss$ = Left$(rest$, i - 1)
            MyRead = SpeedGroup(bstack, pppp, "@READ", ".", ss$, it) <> 0
            Set pppp = Nothing
            rest$ = Mid$(rest$, i)
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
                p = myobject.t1
                If CheckDeepAny(myobject) Then
                    If TypeOf myobject Is mHandler Then
                        Set pppp.item(it) = myobject
                    Else
                        Set usehandler = New mHandler
                        Set pppp.item(it) = usehandler
                        usehandler.t1 = p
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
        End If
    ElseIf bs.IsOptionalForArray(useoptionals) Then
        ' do nothing
        MyRead = True
    Else
        If Not bs.IsNumber(p) Then
            bstack.soros.drop 1
                MissStackNumber
                MyRead = False
                Exit Do
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
                    p = myobject.t1
                    If CheckDeepAny(myobject) Then
                    If TypeOf myobject Is mHandler Then
                    Set pppp.item(it) = myobject
                    Else
                        Set usehandler = New mHandler
                        Set pppp.item(it) = usehandler
                        usehandler.t1 = p
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

End Select
End If
loopcont123:
If MaybeIsSymbol(rest$, "@&!#~") Then SyntaxError: MyRead = False
If optlocal Then useoptionals = False

Loop Until Not FastSymbol(rest$, ",")
Exit Function
err10:
            MyEr "Variable " + what$ + " can't initialized", "« ÏÂÙ·‚ÎÁÙﬁ " + what$ + " ‰ÂÌ ·Ò˜ÈÍÔÔÈËÁÍÂ"
            MyRead = False
            Exit Function
err100:
            MyEr "Parameter is not optional", "« ·Ò‹ÏÂÙÒÔÚ ÂﬂÌ·È ··Ò·ﬂÙÁÙÁ"
            MyRead = False
            Exit Function
er103:
            MyEr "Wrong object type", "À‹ËÔÚ Ù˝ÔÚ ·ÌÙÈÍÂÈÏ›ÌÔı"
            MyRead = False
            Exit Function
er104:
            MyEr "Can't assign value to object", "ƒÂÌ ÏÔÒ˛ Ì· ‰˛Û˘ ÙÈÏﬁ ÛÂ ·ÌÙÈÍÂﬂÏÂÌÔ"
            MyRead = False
            Exit Function
er105:
            MyEr "Cant' Assign number", "ƒÂÌ ÏÔÒ˛ Ì· Ë›Û˘ ÙÈÏﬁ Í·Ù‹ ÙÔ ‰È‹‚·ÛÏ· ÙÈÏ˛Ì"
            MyRead = False
            Exit Function
er106:
            MyEr "No function definition founded", "ƒÂÌ ‚Ò›ËÁÍÂ ÔÒÈÛÏ¸Ú ÛıÌ‹ÒÙÁÛÁÚ"
            MyRead = False
            Exit Function
er107:
            MyEr "Syntax error, use )", "”ıÌÙ·ÍÙÈÍ¸ Î‹ËÔÚ ‚‹ÎÂ )"
            MyRead = False
            Exit Function
er108:
            MyEr "Try other array name", "ƒÔÍﬂÏ·ÛÂ ‹ÎÎÔ ¸ÌÔÏ· ﬂÌ·Í·"
            MyRead = False
            Exit Function
er109:
            MyEr "Cant' change type of variable", "ƒÂÌ ÏÔÒ˛ Ì· ·ÎÎ‹Ó˘ Ù˝Ô ÏÂÙ·‚ÎÁÙﬁÚ"
            MyRead = False
            Exit Function
er110:
            MyEr "need a type after as", "˜ÒÂÈ‹ÊÔÏ·È ›Ì· Ù˝Ô ÏÂÙ‹ ÙÁÌ ˘Ú"
            MyRead = False
            Exit Function
er111:
            MyEr "Missing String literal", "ƒÂÌ ‚ÒﬁÍ· ÛÙ·ËÂÒﬁ ·Îˆ·ÒÈËÏÁÙÈÍﬁ"
            MyRead = False
er112:
            MyEr "Wrong Enumeration type", "À‹ËÔÚ Ù˝ÔÚ ··ÒÈËÏÁÙﬁ"
            MyRead = False
            Exit Function
er113:
            MyEr "Expected Group of type " + s$, "–ÂÒﬂÏÂÌ· œÏ‹‰· Ù˝Ôı " + s$
            MyRead = False
            Exit Function
End Function

