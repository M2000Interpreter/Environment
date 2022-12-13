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
    reserved    As Currency
End Type

Private Declare Function PeekArray Lib "kernel32" Alias "RtlMoveMemory" (Arr() As Any, Optional ByVal Length As Long = 4) As PeekArrayType
Private Declare Function SafeArrayGetDim Lib "OleAut32.dll" (ByVal Ptr As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Long)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal addr As Long, retval As Long)
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
Dim h, where, ret As Currency, lowlong As Long, highlong As Long
Dim FileError As Long
ret = CCur(Int(vvv)) - 1
If InUseHandlers.ExistKey(RHS) Then
    h = InUseHandlers.Value
    where = h(0)
    Size2Long ret, lowlong, highlong
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_BEGIN)
    FileError = GetLastError()
    If lowlong = INVALID_SET_FILE_POINTER And FileError <> 0 Then
        MyEr "Can't write the seek value", "Δεν μπορώ να γράψω τη τιμή μετάθεσης"
    End If
Else
    MyEr "No such file handler", "Δεν υπάρχει τέτοιο χειριστής αρχείου"
End If
End Property
Public Property Get FileSeek(RHS) As Variant
Dim h, where, ret As Currency, lowlong As Long, highlong As Long
FileSeek = ret
If InUseHandlers.ExistKey(RHS) Then
    h = InUseHandlers.Value
    where = h(0)
    lowlong = 0
    highlong = 0
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_CURRENT)
    If lowlong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value", "Δεν μπορώ να διαβάσω τη τιμή μετάθεσης"
    Else
        Long2Size lowlong, highlong, ret
        FileSeek = ret + 1
    End If
Else
    MyEr "No such file handler", "Δεν υπάρχει τέτοιο χειριστής αρχείου"
End If
End Property
Public Property Let FileSeekFH(where As Long, ByVal ret As Currency)
Dim lowlong As Long, highlong As Long
Dim FileError As Long
    Size2Long ret - 1@, lowlong, highlong
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_BEGIN)
    FileError = GetLastError()
    If lowlong = INVALID_SET_FILE_POINTER And FileError <> 0 Then
        MyEr "Can't write the seek value", "Δεν μπορώ να γράψω τη τιμή μετάθεσης"
    End If
End Property
Public Property Get FileSeekFH(where As Long) As Currency
Dim lowlong As Long, highlong As Long
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_CURRENT)
    If lowlong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value", "Δεν μπορώ να διαβάσω τη τιμή μετάθεσης"
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
        MyEr "Can't read the seek value on Eof() function", "Δεν μπορώ να διαβάσω τη τιμή μετάθεσης στη συνάρτηση Μετάθεση()"
    Else
        Long2Size lowlong, highlong, ret
        lowlong = GetFileSize(where, highlong)
        Long2Size lowlong, highlong, fsize
        FileEOFFH = ret >= fsize
    End If
End Property

Public Property Get FileEOF(RHS) As Boolean
Dim h, where, ret As Currency, lowlong As Long, highlong As Long
Dim fsize As Currency
If InUseHandlers.ExistKey(RHS) Then
    h = InUseHandlers.Value
    where = h(0)
    lowlong = SetFilePointer(where, lowlong, highlong, FILE_BEGIN)
    If lowlong = INVALID_SET_FILE_POINTER Then
        MyEr "Can't read the seek value on Eof() function", "Δεν μπορώ να διαβάσω τη τιμή μετάθεσης στη συνάρτηση Μετάθεση()"
    Else
        Long2Size lowlong, highlong, ret
        lowlong = GetFileSize(where, highlong)
        Long2Size lowlong, highlong, fsize
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
Dim h&, ar() As Variant
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
Public Sub API_ReadBLOCK(ByVal FileNumber As Long, ByVal BlockSize As Long, ByVal addr As Long)
Dim PosL As Long
Dim PosH As Long
Dim SizeRead As Long
Dim ret As Long
ret = SetFilePointer(FileNumber, PosL, PosH, FILE_CURRENT)
ret = ReadFile(FileNumber, ByVal addr, BlockSize, SizeRead, 0&)
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
        MyEr "can't decode this url", "δεν μπορώ να αποκωδικοποιήσω την διεύθυνση"
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
    Dim exclude As String, domain As String, scheme As String, w As Long
    
    
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
    w = InStr(Address, "#")
    If w > 0 Then Address = Left$(Address, w - 1)
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
        Dim w As Long
        GetHost = GetDomainName(url$, True)
        If GetHost <> vbNullString Then
        If Left$(GetHost, 1) <> "[" Then
            w = InStr(GetHost, "@")
            If w > 0 Then GetHost = Mid$(GetHost, w + 1)
            If GetHost <> vbNullString Then
                 w = InStr(GetHost, ":")
                If w > 0 Then GetHost = Left$(GetHost, w - 1)
            End If
        Else
            w = InStr(GetHost, "]")
            GetHost = Mid$(GetHost, 2, w - 2)
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

