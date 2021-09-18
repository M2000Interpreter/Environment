Attribute VB_Name = "Module4"
Option Explicit
Private Const clOneMask = 16515072          '000000 111111 111111 111111
Private Const clTwoMask = 258048            '111111 000000 111111 111111
Private Const clThreeMask = 4032            '111111 111111 000000 111111
Private Const clFourMask = 63               '111111 111111 111111 000000

Private Const clHighMask = 16711680         '11111111 00000000 00000000
Private Const clMidMask = 65280             '00000000 11111111 00000000
Private Const clLowMask = 255               '00000000 00000000 11111111

Private Const cl2Exp18 = 262144             '2 to the 18th power
Private Const cl2Exp12 = 4096               '2 to the 12th
Private Const cl2Exp6 = 64                  '2 to the 6th
Private Const cl2Exp8 = 256                 '2 to the 8th
Private Const cl2Exp16 = 65536              '2 to the 16th

Public mSysHandlerWasSet
Public clickMe As Long, clickMe2 As Long
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_NO_BUFFERING = &H20000000
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000

Public Const PIPE_ACCESS_DUPLEX = &H3
Public Const PIPE_ACCESS_INBOUND = &H1
Public Const PIPE_READMODE_MESSAGE = &H2
Public Const PIPE_TYPE_MESSAGE = &H4
Public Const PIPE_WAIT = &H0
Public Const PIPE_NOWAIT = &H1
Public Const WRITE_DAC = &H40000
Public Const PIPE_READMODE_BYTE = &H0
Public Const PIPE_TYPE_BYTE = &H0
Public Const ERROR_NO_DATA = 232&
Public Const NMPWAIT_USE_DEFAULT_WAIT = &H0
Public Const ERROR_PIPE_LISTENING = 536&
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINESCROLL = &HB6
Private Const EM_GETLINECOUNT = 186
Public Const INVALID_HANDLE_VALUE = -1
Declare Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeW" (ByVal lpName As Long, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As Any) As Long
Declare Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpOverlapped As Long) As Long
Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long) As Long
 Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, _
      ByVal nNumberOfBytesToRead As Long, _
      lpNumberOfBytesRead As Long, _
      lpOverlapped As Any) As Long
Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeW" (ByVal lpNamedPipeName As Long, ByVal nTimeOut As Long) As Long
      
'Declare Function CallNamedPipe Lib "KERNEL32" Alias "CallNamedPipeW" (ByVal lpNamedPipeName As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, ByVal nTimeOut As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
'' MoveFile
Declare Function MoveFile Lib "kernel32" Alias "MoveFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3
Public Const CREATE_NEW = 1
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright �1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
   (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
    
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
    Private Const LWA_Defaut         As Long = &H2
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Const GW_HWNDNEXT = 2
Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_CUT As Long = &H300
Private Const WM_COPY As Long = &H301
Private Const WM_PAST As Long = &H302
Private Const EM_CANUNDO = &HC6
Private Const WM_USER = &H400
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_UNDO = WM_USER + 23
Public defWndProc As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
                
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long
Private Declare Function SendMessageAsLong Lib "user32" _
       Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
       ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const SND_APPLICATION = &H80 ' look for application specific association
Private Const SND_ALIAS = &H10000 ' name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000 ' name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1 ' play asynchronously
Private Const SND_FILENAME = &H20000 ' name is a file name
Private Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4 ' lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Private Const SND_NOSTOP = &H10 ' don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000 ' don't wait if the driver is busy
Private Const SND_PURGE = &H40 ' purge non-static events for task
Private Const SND_RESOURCE = &H40004 ' name is a resource name or atom
Private Const SND_SYNC = &H0 ' play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameW" (ByVal lpszLongPath As Long, _
ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Type SHFILEOPSTRUCTW
    hWnd As Long
    wFunc As Long
    pFrom As Long 'String
    pTo As Long 'String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As Long  'String
End Type

Public Enum FOACTION
    FO_MOVE = &H1
    FO_COPY = &H2
    FO_DELETE = &H3
    FO_RENAME = &H4
End Enum

Public Enum FOFACTION
    FOF_ALLOWUNDO = &H40
    FOF_CONFIRMMOUSE = &H2
    FOF_FILESONLY = &H80
    FOF_MULTIDESTFILES = &H1
    FOF_NO_CONNECTED_ELEMENTS = &H2000
    FOF_NOCONFIRMATION = &H10
    FOF_NOCONFIRMMKDIR = &H200
    FOF_NOCOPYSECURITYATTRIBS = &H800
    FOF_NOERRORUI = &H400
    FOF_NORECURSION = &H1000
    FOF_RENAMEONCOLLISION = &H8
    FOF_SILENT = &H4
    FOF_SIMPLEPROGRESS = &H100
    FOF_WANTMAPPINGHANDLE = &H20
    FOF_WANTNUKEWARNING = &H4000
End Enum
Private Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName As String * 64
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName As String * 64
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type
Private Declare Function getTimeFormat Lib "kernel32" Alias "GetTimeFormatW" ( _
    ByVal Locale As Long, _
    ByVal dwFlags As Long, _
    ByRef lpTime As SYSTEMTIME, _
    ByVal lpFormat As Long, _
    ByVal lpTimeStr As Long, _
    ByVal cchTime As Long _
) As Long
Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatW" ( _
    ByVal Locale As Long, _
    ByVal dwFlags As Long, _
    ByRef lpDate As SYSTEMTIME, _
    ByVal lpFormat As Long, _
    ByVal lpDateStr As Long, _
    ByVal cchDate As Long _
) As Long
Private Declare Function VarDateFromStr Lib "OleAut32.dll" ( _
    ByVal psDateIn As Long, _
    ByVal lcid As Long, _
    ByVal uwFlags As Long, _
    ByRef dtOut As Date) As Long
Private Const S_OK = 0
Private Const DISP_E_BADVARTYPE = &H80020008
Private Const DISP_E_OVERFLOW = &H8002000A
Private Const DISP_E_TYPEMISMATCH = &H80020005
Private Const E_INVALIDARG = &H80070057
Private Const E_OUTOFMEMORY = &H8007000E
Private Declare Function VariantTimeToSystemTime Lib "OleAut32.dll" ( _
    ByVal vtime As Date, _
    ByRef lpSystemTime As SYSTEMTIME _
) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32.dll" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
  
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpszPath As Long, ByVal lpSA As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private EncbTrans(63) As Byte, EnclPowers8(255) As Long, EnclPowers16(255) As Long
Dim lPowers18(63) As Long, bTrans(255) As Byte, lPowers6(63) As Long, lPowers12(63) As Long
Public Function DateFromString(ByVal sDateIn As String, ByVal lcid As Long) As Date

    Dim hResult As Long
    Dim dtOut As Date

    ' Do not want user's own settings to override the standard formatting settings
    ' if they are using the same locale that we are converting from.
    '
    Const LOCALE_NOUSEROVERRIDE = &H80000000

    ' Do the conversion
    hResult = VarDateFromStr(StrPtr(sDateIn), lcid, LOCALE_NOUSEROVERRIDE, dtOut)

    Select Case hResult

        Case S_OK:
            DateFromString = dtOut
        Case Else
            MyEr "Can't convert to date", "��� ����� �� �� ��������� �� ����������"
    End Select

End Function
Public Function FormatTimeWithLocale(ByRef the_sFormat As String, the_datTime As Date, ByVal the_nLocale As Long) As String

    Dim uSystemTime As SYSTEMTIME
    Dim nBufferSize As Long


    If VariantTimeToSystemTime(the_datTime, uSystemTime) = 1 Then
        nBufferSize = getTimeFormat(the_nLocale, 0&, uSystemTime, StrPtr(the_sFormat), 0&, 0&)
        If nBufferSize > 0 Then
            FormatTimeWithLocale = space$(nBufferSize - 1)
            getTimeFormat the_nLocale, 0&, uSystemTime, StrPtr(the_sFormat), StrPtr(FormatTimeWithLocale), nBufferSize
        End If

    End If

End Function

Public Function FormatDateWithLocale(ByRef the_sFormat As String, the_datDate As Date, ByVal the_nLocale As Long) As String

    Dim uSystemTime As SYSTEMTIME
    Dim nBufferSize As Long

    'https://stackoverflow.com/questions/11530790/force-date-to-us-format-regardless-of-locale-settings
    If VariantTimeToSystemTime(the_datDate, uSystemTime) = 1 Then
        nBufferSize = GetDateFormat(the_nLocale, 0&, uSystemTime, StrPtr(the_sFormat), 0&, 0&)
        If nBufferSize > 0 Then
            FormatDateWithLocale = space$(nBufferSize - 1)
            GetDateFormat the_nLocale, 0&, uSystemTime, StrPtr(the_sFormat), StrPtr(FormatDateWithLocale), nBufferSize
        End If

    End If

End Function
Public Function GetTimeZoneInfo() As String

Dim TZI As TIME_ZONE_INFORMATION ' receives information on the time zone
Dim retval As Long ' return value
Dim c As Long ' counter variable needed to display time zone name

    retval = GetTimeZoneInformation(TZI) ' read information on the computer's selected time zone
     If zones.ExistKey(Replace(StrConv(TZI.StandardName, vbFromUnicode), Chr(0), "")) Then
     ' do nothing. now zone$ has set the index field to standard name.
     End If
    If retval = 2 Then
    GetTimeZoneInfo = Replace(StrConv(TZI.DaylightName, vbFromUnicode), Chr(0), "")
    Else
    GetTimeZoneInfo = Replace(StrConv(TZI.StandardName, vbFromUnicode), Chr(0), "")
    End If
    
End Function
Public Function GetUTCDate() As Date
    Dim tzTime As TIME_ZONE_INFORMATION
    Dim lngUTCTime As Long
     
    lngUTCTime = GetTimeZoneInformation(tzTime)
    GetUTCDate = DateAdd("n", tzTime.Bias + tzTime.DaylightBias, Now)
End Function
Public Function GetUTCTime() As Date
    Dim tzTime As TIME_ZONE_INFORMATION
    Dim lngUTCTime As Long
     
    lngUTCTime = GetTimeZoneInformation(tzTime)
    GetUTCTime = DateAdd("n", tzTime.Bias + tzTime.DaylightBias, Now)
End Function
Public Sub SetUp64()
    Dim lTemp As Long
    For lTemp = 0 To 63                                 'Fill the translation table.
        Select Case lTemp
            Case 0 To 25
                EncbTrans(lTemp) = 65 + lTemp              'A - Z
            Case 26 To 51
                EncbTrans(lTemp) = 71 + lTemp              'a - z
            Case 52 To 61
                EncbTrans(lTemp) = lTemp - 4               '1 - 0
            Case 62
                EncbTrans(lTemp) = 43                      'Chr(43) = "+"
            Case 63
                EncbTrans(lTemp) = 47                      'Chr(47) = "/"
        End Select
         lPowers6(lTemp) = lTemp * cl2Exp6
        lPowers12(lTemp) = lTemp * cl2Exp12
        lPowers18(lTemp) = lTemp * cl2Exp18
    Next lTemp
    For lTemp = 0 To 255                                'Fill the translation table.
        Select Case lTemp
            Case 65 To 90
                bTrans(lTemp) = lTemp - 65              'A - Z
            Case 97 To 122
                bTrans(lTemp) = lTemp - 71              'a - z
            Case 48 To 57
                bTrans(lTemp) = lTemp + 4               '1 - 0
            Case 43
                bTrans(lTemp) = 62                      'Chr(43) = "+"
            Case 47
                bTrans(lTemp) = 63                      'Chr(47) = "/"
        End Select
        EnclPowers8(lTemp) = lTemp * cl2Exp8
        EnclPowers16(lTemp) = lTemp * cl2Exp16
    Next lTemp

End Sub

   
Public Function PathMakeDirs(ByVal Pathd As String) As Boolean
        Pathd = PurifyPath(Pathd)
      PathMakeDirs = 0 <> CreateDirectory(StrPtr(Pathd), 0&)
   End Function
 
Function PurifyPath(Spath$) As String
Dim a$(), i
If Spath$ = vbNullString Then Exit Function
a$() = Split(Spath, "\")
If isdir(a$(LBound(a$()))) Then i = i + 1
For i = LBound(a$()) + i To UBound(a$())
a$(i) = PurifyName(a$(i))
Next i
If LBound(a()) = UBound(a()) Then
PurifyPath = a$(UBound(a$()))
Else
PurifyPath = ExtractPath(Join(a$, "\") & "\", False)
End If
End Function
Public Function PurifyName(sStr As String) As String
Dim noValidcharList
noValidcharList = "*?\<>:/|" + Chr$(34)
Dim a$, i As Long, ddt As Boolean, j As Long
If Len(sStr) > 0 Then
a$ = space$(Len(sStr))
j = 1
For i = 1 To Len(sStr)
If InStr(noValidcharList, Mid$(sStr, i, 1)) = 0 Then
Mid$(a$, j, 1) = Mid$(sStr, i, 1)
Else
Mid$(a$, j, 1) = "-"
End If
j = j + 1
Next i
a$ = Left$(a$, j - 1)
End If
PurifyName = a$
End Function

Public Sub FixPath(s$)
Dim frm$
If s$ <> "" Then
If Left$(s$, 1) = "." And Mid$(s$, 2, 1) <> "." Then
s$ = mcd + Mid$(s$, 2)
End If
frm$ = ExtractPath(s$)
If frm$ = vbNullString Or Left$(s$, 2) = ".." Then
    s$ = mcd + s$
Else
    If Left$(frm$, 2) = "\\" Or Mid$(frm$, 2, 1) = ":" Then
    'root
    Else
    s$ = userfiles$ + s$
    End If
End If
End If
End Sub
Public Function RenameFile(ByVal sSourceFile As String, ByVal sDesFile As String) As Boolean
Dim F$, fd$
If Not CanKillFile(sSourceFile) Then Exit Function
If ExtractType(sSourceFile) = vbNullString Then sSourceFile = sSourceFile + ".gsb"
If ExtractType(sDesFile) = vbNullString Then
If ExtractNameOnly(sDesFile, True) = ExtractNameOnly(sSourceFile, True) Then
sDesFile = ExtractNameOnly(sDesFile, True) + ".bck"
Else
sDesFile = ExtractNameOnly(sDesFile, True) + ".gsb"
End If
End If
sSourceFile = CFname(sSourceFile)
If sSourceFile = vbNullString Or CFname(sDesFile) <> "" Then
BadFilename
Exit Function
Else
sDesFile = ExtractPath(sSourceFile) + ExtractName(sDesFile, True)
End If
If Left$(sSourceFile, 2) <> "\\" Then
F$ = "\\?\" + sSourceFile
Else
F$ = sSourceFile
End If
If Left$(sDesFile, 2) <> "\\" Then
fd$ = "\\?\" + sDesFile
Else
fd$ = sDesFile
End If
RenameFile = 0 <> MoveFile(StrPtr(F$), StrPtr(fd$))

End Function

Public Function RenameFile2(ByVal sSourceFile As String, ByVal sDesFile As String) As Boolean
Dim F$, fd$
sDesFile = ExtractPath(sSourceFile) + ExtractName(sDesFile, True)
If Left$(sSourceFile, 2) <> "\\" Then
F$ = "\\?\" + sSourceFile
Else
F$ = sSourceFile
End If
If Left$(sDesFile, 2) <> "\\" Then
fd$ = "\\?\" + sDesFile
Else
fd$ = sDesFile
End If
RenameFile2 = 0 <> CopyFile(StrPtr(F$), StrPtr(fd$), 1)
KillFile F$
End Function
Public Function CanKillFile(FileName$) As Boolean
FixPath FileName$
If Not IsSupervisor Then
    If Left$(FileName$, 1) = "." Then
        CanKillFile = True
    Else
      If strTemp <> "" Then
            If Not mylcasefILE(strTemp) = mylcasefILE(Left$(FileName$, Len(strTemp))) Then
            CanKillFile = mylcasefILE(userfiles) = mylcasefILE(Left$(FileName$, Len(userfiles)))
            Else
            CanKillFile = True
            End If
        Else
            CanKillFile = mylcasefILE(userfiles) = mylcasefILE(Left$(FileName$, Len(userfiles)))
        End If
    End If
Else
    CanKillFile = True
End If

End Function
Public Function MakeACopy(ByVal sSourceFile As String, ByVal sDesFile As String) As Boolean
If Not CanKillFile(sSourceFile) Then Exit Function
Dim F$, fd$
If Left$(sSourceFile, 2) <> "\\" Then
F$ = "\\?\" + sSourceFile
Else
F$ = sSourceFile
End If
If Left$(sDesFile, 2) <> "\\" Then
fd$ = "\\?\" + sDesFile
Else
fd$ = sDesFile
End If

MakeACopy = 0 <> CopyFile(StrPtr(F$), StrPtr(fd$), 0)
End Function

Public Function NeoUnicodeFile(FileName$) As Boolean
Dim hFile, counter
Dim F$, F1$
Sleep 10
If Not CanKillFile(FileName$) Then Exit Function
If Left$(FileName$, 2) <> "\\" Then
F$ = "\\?\" + FileName$
Else
F$ = FileName$
End If
On Error Resume Next
F1$ = Dir(F$)  '' THIS IS THEWORKAROUND FOR THE PROBLEMATIC CREATIFILE (I GOT SOME HANGS)

hFile = CreateFile(StrPtr(F$), GENERIC_WRITE, ByVal 0, ByVal 0, 2, FILE_ATTRIBUTE_NORMAL, ByVal 0)
FlushFileBuffers hFile
Sleep 10

CloseHandle hFile

NeoUnicodeFile = (CFname(GetDosPath(F$)) <> "")

Sleep 10

'need "\\?\" before

'now we can use the getdospath from normal Open File


End Function

Public Function GetDosPath(LongPath As String) As String

Dim s As String
Dim i As Long
Dim PathLength As Long

        i = Len(LongPath) * 2 + 2

        s = String(1024, 0)

        PathLength = GetShortPathName(StrPtr(LongPath), StrPtr(s), i)

        GetDosPath = Left$(s, PathLength)

End Function

Sub PlaySoundNew(F As String)

If F = vbNullString Then
PlaySound 0&, 0&, SND_PURGE
Else
If ExtractType(F) = vbNullString Then F = F & ".WAV"
F = CFname(F)
PlaySound StrPtr(F), ByVal 0&, SND_FILENAME Or SND_ASYNC
End If
End Sub  ' SND_MEMORY
Sub PlaySoundNew2(F As Long)
PlaySound (F), ByVal 0&, SND_MEMORY Or SND_ASYNC
End Sub

       
Public Sub SetTrans(oForm As Form, Optional bytAlpha As Byte = 255, Optional lColor As Long = 0, Optional TRMODE As Boolean = False)
    Dim lStyle As Long
    lStyle = GetWindowLong(oForm.hWnd, GWL_EXSTYLE)
    If Not (lStyle And WS_EX_LAYERED) = WS_EX_LAYERED Then _
        SetWindowLong oForm.hWnd, GWL_EXSTYLE, lStyle Or WS_EX_LAYERED
       
    SetLayeredWindowAttributes oForm.hWnd, lColor, bytAlpha, IIf(TRMODE, LWA_COLORKEY Or LWA_Defaut, LWA_Defaut)
    UpdateWindow oForm.hWnd
End Sub


Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function


Function Trans2pipe(pipe$, what$) As Boolean
Dim s$, IL As Long, ok As Long, hPipe As Long, ok2 As Long
s$ = validpipename(pipe$)
Dim b() As Byte
b() = what$
ReDim Preserve b(Len(what$) * 2 + 20) As Byte
Trans2pipe = True
hPipe = CreateFile(StrPtr(s$), GENERIC_WRITE, ByVal 0, ByVal 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)
If hPipe <> INVALID_HANDLE_VALUE Then
ok2 = WriteFile(hPipe, b(0), Len(what$) * 2, IL, ByVal 0)
ok = WaitNamedPipe(StrPtr(s$), 1000)
ok = GetLastError = 0
Trans2pipe = ok2 > 0 Or ok
Else
Trans2pipe = False
End If
CloseHandle hPipe
Sleep 1
End Function
Function validpipename(ByVal a$) As String
Dim b$
a$ = myUcase(a$)
b$ = Left$(a$, InStr(1, a$, "\pipe\", vbTextCompare))
If b$ = vbNullString Then
validpipename = "\\" & strMachineName & "\pipe\" & a$
Else
validpipename = a$
End If
End Function

Function Included(afile$, ByVal simple$) As String
Dim a As Document
On Error GoTo inc1
Dim what As Long
Dim st&, pa&, po&
st& = 1
Dim word$(), it As Long, max As Long, Line$, ok As Boolean, Min As Long
simple$ = simple$ + "|"
word$() = Split(simple$, "|")


max = UBound(word$()) - 1
If max > 0 Then
For st& = 0 To max
If word$(st&) = vbNullString Then
MyEr "Need to give a string of type ""word1|word2....""", "���������� �� ������ ��� ������������� ""����|��������..."""
Exit Function
End If
Next
End If
st& = 1
If Len(simple$) <= 1 Then
Included = ExtractName(afile$, True)
Else
    Sleep 1
    Set a = New Document
    
    a.lcid = cLid
    a.ReadUnicodeOrANSI afile$, , what
    
    If InStr(simple$, vbCr) > 0 Then
    'work with any char but using computer locale
    If InStr(1, a.textDoc, word$(0), vbTextCompare) > 0 Then
                Included = ExtractName(afile$, True)
    End If
    Else
    ' work in paragraphs..
  
again:

    st& = a.FindStr(word$(0), st&, pa&, po&)
    
    If st& > 0 Then
    
     If max > 0 Then
     po& = po& + Len(word$(0))
     Line$ = a.TextParagraph(a.ParagraphFromOrder(pa&))
     For it = 1 To max
     ok = InStr(po&, Line$, word$(it), vbTextCompare) > 0
     If Not ok Then Exit For
     Next it
     st& = st& + Len(word$(0))
    If ok Then Included = ExtractName(afile$, True) Else GoTo again
    
     Else
            Included = ExtractName(afile$, True)
           End If
    End If
    End If
    Set a = Nothing

End If
inc1:
End Function

' from VbForoums after corrections by me.
' http://www.vbforums.com/showthread.php?379072-VB-Fast-Base64-Encoding-and-Decoding
Public Function Decode64(sString As String, ok As Boolean) As String
    ok = True
    If Len(sString) = 0 Then Exit Function
    Dim bOut() As Byte, bIn() As Byte, lQuad As Long, iPad As Integer, lChar As Long, lPos As Long, sOut As String
    Dim lTemp As Long

    bIn = StrConv(sString, vbFromUnicode)              'Load the input byte array.
    ReDim bOut((((UBound(bIn) + 1) \ 4) * 3) - 1 + 4)     'Prepare the output buffer.
    lChar = 0
    Dim lChar2 As Long, lChar3 As Long, lChar4 As Long, lchar1 As Long
    Dim ubnd As Long
    ubnd = UBound(bIn)
    lChar3 = lChar - 1
    Do
        ' take 4
        
        lChar = lChar3 + 1
        If lChar >= ubnd Then lChar3 = lChar: GoTo finish
        ok = False
        Do
        Select Case bIn(lChar)
            Case 65 To 90, 97 To 122, 48 To 57, 43, 47
            Exit Do
            Case 61
            lChar3 = lChar: GoTo finish
            Exit Do
            Case 10, 13, 32
            lChar = lChar + 1
            If lChar > ubnd Then lChar3 = lChar: GoTo finish
            Case Else
            
            lChar = lChar + 1
            lChar3 = lChar: GoTo finish
            GoTo finish
        End Select
        Loop
        lchar1 = lChar + 1
        If lchar1 > ubnd Then lChar3 = lchar1: GoTo finish
        Do
        Select Case bIn(lchar1)
            Case 65 To 90, 97 To 122, 48 To 57, 43, 47
            Exit Do
            Case 61
            lChar3 = lchar1: GoTo finish
            Case 10, 13, 32
            lchar1 = lchar1 + 1
            If lchar1 > ubnd Then lChar3 = lchar1: GoTo finish
            Case Else
            lchar1 = lchar1 + 1
            lChar3 = lchar1: GoTo finish
        End Select
        Loop
        lChar2 = lchar1 + 1
        If lChar2 > ubnd Then lChar3 = lChar2: GoTo finish
        Do
        Select Case bIn(lChar2)
            Case 65 To 90, 97 To 122, 48 To 57, 43, 47
            Exit Do
            Case 61
           lChar3 = lChar2: GoTo finish
            Exit Do
            Case 10, 13, 32
            lChar2 = lChar2 + 1
            If lChar2 > ubnd Then lChar3 = lChar2: GoTo finish
            Case Else
            lChar2 = lChar2 + 1
            'If lChar2 > ubnd Then
            lChar3 = lChar2: GoTo finish
            'GoTo finish
        End Select
        Loop
        lChar3 = lChar2 + 1
        If lChar3 > ubnd Then GoTo finish
        Do
        Select Case bIn(lChar3)
            Case 65 To 90, 97 To 122, 48 To 57, 43, 47
            Exit Do
            Case 61
            GoTo finish
            Case 10, 13, 32
            lChar3 = lChar3 + 1
            If lChar3 > ubnd Then GoTo finish
            Case Else
            lChar3 = lChar3 + 1
            If lChar3 > ubnd Then GoTo finish
            GoTo finish
        End Select
        Loop
        ok = True
        lQuad = lPowers18(bTrans(bIn(lChar))) + lPowers12(bTrans(bIn(lchar1))) + _
                lPowers6(bTrans(bIn(lChar2))) + bTrans(bIn(lChar3))           'Rebuild the bits.
        lTemp = lQuad And clHighMask                    'Mask for the first byte
        bOut(lPos) = lTemp \ cl2Exp16                   'Shift it down
        lTemp = lQuad And clMidMask                     'Mask for the second byte
        bOut(lPos + 1) = lTemp \ cl2Exp8                'Shift it down
        bOut(lPos + 2) = lQuad And clLowMask            'Mask for the third byte
        lPos = lPos + 3
    Loop
finish:
    If Not ok Then
    iPad = 0
    
    Dim offset As Long
        Do While lChar3 + offset <= ubnd
        Select Case bIn(lChar3 + offset)
            Case 10, 13, 32
            If iPad > 0 Then offset = 0: Exit Do
            offset = offset + 1
            
            Case 61
            iPad = iPad + 1
            offset = offset + 1
            Case Else
             If lChar2 = 0 Then GoTo error1
                
              If iPad = 0 Then
                  iPad = Int(-(lChar3 = lChar2) - (lChar3 = lchar1) - (lChar3 = lChar)) * 3
              Else
                If lChar3 <> lChar2 + 1 And iPad = 1 Then
                   iPad = Int(-(lChar3 = lChar2) - (lChar3 = lchar1) - (lChar3 = lChar)) * 3
                
                End If
              End If
            
             offset = 0
            
            Exit Do
         
        End Select
        Loop
        If iPad > 3 Then GoTo error1
        If iPad = 0 Then
            If (lPos + 2) Mod 3 + 1 <> 3 Then
                GoTo error1
            End If
        Else
            If (lPos + 3 - iPad) Mod 3 + iPad <> 3 Then
                GoTo error1
            End If
        End If
 
        ok = True
        If lChar3 > ubnd Then GoTo cont1
        If lChar = lChar3 Then
        lChar = lChar + 1: lChar3 = lChar3 + 1
        lchar1 = lChar3
        lChar2 = lChar3
        End If
        If lchar1 = lChar3 Then
        lchar1 = lchar1 + 1: lChar3 = lChar3 + 1
        lChar2 = lChar3
        End If
        If lChar2 = lChar3 Then
        lChar2 = lChar2 + 1: lChar3 = lChar3 + 1
        End If
        If lChar3 <= ubnd Then
        lQuad = lPowers18(bTrans(bIn(lChar))) + lPowers12(bTrans(bIn(lchar1))) + _
        lPowers6(bTrans(bIn(lChar2))) + bTrans(bIn(lChar3))           'Rebuild the bits.
        lTemp = lQuad And clHighMask                    'Mask for the first byte
        bOut(lPos) = lTemp \ cl2Exp16                   'Shift it down
        lTemp = lQuad And clMidMask                     'Mask for the second byte
        bOut(lPos + 1) = lTemp \ cl2Exp8                'Shift it down
        bOut(lPos + 2) = lQuad And clLowMask
        lPos = lPos + 3 - iPad
        End If
        
        GoTo cont1
error1:
            Exit Function
    End If
cont1:
If lPos Mod 2 = 1 Then
    sOut = StrConv(String$(lPos, Chr(0)), vbFromUnicode)
Else
    sOut = String$((lPos + 1) \ 2, Chr(0))
    End If
    CopyMemory ByVal StrPtr(sOut), bOut(0), LenB(sOut)
    Decode64 = sOut
End Function


Public Function Encode64(sString As String, Optional compact As Boolean = False, Optional ByVal leftmargin As Long = 0) As String

    Dim bOut() As Byte, bIn() As Byte
    Dim lChar As Long, lTrip As Long, iPad As Integer, lLen As Long, lTemp As Long, lPos As Long, lOutSize As Long
    
        If LenB(sString) = 0 Then Exit Function
    iPad = (3 - LenB(sString) Mod 3) Mod 3                          'See if the length is divisible by 3
    ReDim bIn(0 To LenB(sString) - 1 + iPad)
    CopyMemory bIn(0), ByVal StrPtr(sString), LenB(sString)
    lLen = ((UBound(bIn) + 1) \ 3) * 4                  'Length of resulting string.
    ' set to 60 wchar for each line break
    lTemp = lLen \ 60                                   'Added space for vbCrLfs.
    lOutSize = ((lTemp * 2) + leftmargin * (lTemp + 1) + lLen) - 1        'Calculate the size of the output buffer.
    ReDim bOut(lOutSize)                                'Make the output buffer.
    
    lLen = 0                                            'Reusing this one, so reset it.
    Dim insertspace As Boolean
    If leftmargin > 0 Then insertspace = True
    For lChar = LBound(bIn) To UBound(bIn) - 2 Step 3
        If insertspace Then
        If leftmargin > 0 Then
            For lPos = lPos To lPos + leftmargin - 1
                bOut(lPos) = 32
            Next lPos
        End If
        insertspace = False
        End If
        lTrip = EnclPowers16(bIn(lChar)) + EnclPowers8(bIn(lChar + 1)) + bIn(lChar + 2)    'Combine the 3 bytes
        lTemp = lTrip And clOneMask                     'Mask for the first 6 bits
        bOut(lPos) = EncbTrans(lTemp \ cl2Exp18)           'Shift it down to the low 6 bits and get the value
        lTemp = lTrip And clTwoMask                     'Mask for the second set.
        bOut(lPos + 1) = EncbTrans(lTemp \ cl2Exp12)       'Shift it down and translate.
        lTemp = lTrip And clThreeMask                   'Mask for the third set.
        bOut(lPos + 2) = EncbTrans(lTemp \ cl2Exp6)        'Shift it down and translate.
        bOut(lPos + 3) = EncbTrans(lTrip And clFourMask)   'Mask for the low set.
        If lLen = 60 And Not compact Then                               'Ready for a newline
            bOut(lPos + 4) = 13                         'Chr(13) = vbCr
            bOut(lPos + 5) = 10                         'Chr(10) = vbLf
            lLen = 0                                    'Reset the counter
            lPos = lPos + 6
            insertspace = True
        Else
            lLen = lLen + 4
            lPos = lPos + 4
        End If
    Next lChar
    ReDim Preserve bOut(lPos - 1)
    lOutSize = lPos - 1
    If bOut(lOutSize) = 10 Then lOutSize = lOutSize - 2 'Shift the padding chars down if it ends with CrLf.
    
    
    If iPad = 1 Then                                    'Add the padding chars if any.
        bOut(lOutSize) = 61                             'Chr(61) = "="
    ElseIf iPad = 2 Then
        bOut(lOutSize) = 61
        bOut(lOutSize - 1) = 61
    End If
    
    Encode64 = StrConv(bOut, vbUnicode)                 'Convert back to a string and return it.
    
End Function

'
Public Function FileToEncode64(F$, leftmargin As Long) As String
Dim i As Long, fl As Long

fl = FileLen(GetDosPath(F$))
If fl = 0 Then Exit Function
    i = FreeFile
    Open GetDosPath(F$) For Binary Access Read As i
    Dim bOut() As Byte, bIn() As Byte
    Dim lChar As Long, lTrip As Long, iPad As Integer, lLen As Long, lTemp As Long, lPos As Long, lOutSize As Long
    iPad = (3 - FileLen(GetDosPath(F$)) Mod 3) Mod 3                          'See if the length is divisible by 3
    ReDim bIn(0 To fl - 1)
    Get #i, , bIn()
    Close i
    ReDim Preserve bIn(0 To fl - 1 + iPad)
    lLen = ((UBound(bIn) + 1) \ 3) * 4                  'Length of resulting string.
    ' set to 60 wchar for each line break
    lTemp = lLen \ 60                                   'Added space for vbCrLfs.
    lOutSize = ((lTemp * 2) + leftmargin * (lTemp + 1) + lLen) - 1        'Calculate the size of the output buffer.
    ReDim bOut(lOutSize)                                'Make the output buffer.
    
    lLen = 0                                            'Reusing this one, so reset it.
    Dim insertspace As Boolean
    If leftmargin > 0 Then insertspace = True
    For lChar = 0 To fl - 2 + iPad Step 3
        If insertspace Then
        If leftmargin > 0 Then
            For lPos = lPos To lPos + leftmargin - 1
                bOut(lPos) = 32
            Next lPos
          '  lLen = lLen + LeftMargin
        End If
        insertspace = False
        End If
        lTrip = EnclPowers16(bIn(lChar)) + EnclPowers8(bIn(lChar + 1)) + bIn(lChar + 2)    'Combine the 3 bytes
        lTemp = lTrip And clOneMask                     'Mask for the first 6 bits
        bOut(lPos) = EncbTrans(lTemp \ cl2Exp18)           'Shift it down to the low 6 bits and get the value
        lTemp = lTrip And clTwoMask                     'Mask for the second set.
        bOut(lPos + 1) = EncbTrans(lTemp \ cl2Exp12)       'Shift it down and translate.
        lTemp = lTrip And clThreeMask                   'Mask for the third set.
        bOut(lPos + 2) = EncbTrans(lTemp \ cl2Exp6)        'Shift it down and translate.
        bOut(lPos + 3) = EncbTrans(lTrip And clFourMask)   'Mask for the low set.
        If lLen = 60 Then                               'Ready for a newline
            bOut(lPos + 4) = 13                         'Chr(13) = vbCr
            bOut(lPos + 5) = 10                         'Chr(10) = vbLf
            lLen = 0                                    'Reset the counter
            lPos = lPos + 6
            insertspace = True
        Else
            lLen = lLen + 4
            lPos = lPos + 4
        End If
    Next lChar
    ReDim Preserve bOut(lPos - 1)
    lOutSize = lPos - 1
    If bOut(lOutSize) = 10 Then lOutSize = lOutSize - 2 'Shift the padding chars down if it ends with CrLf.
    
    
    If iPad = 1 Then                                    'Add the padding chars if any.
        bOut(lOutSize) = 61                             'Chr(61) = "="
    ElseIf iPad = 2 Then
        bOut(lOutSize) = 61
        bOut(lOutSize - 1) = 61
    End If
    
    FileToEncode64 = StrConv(bOut, vbUnicode)



End Function



Public Sub ShowForm1()
On Error Resume Next
If Form1.Visible Then
If Screen.ActiveForm Is Form1 Then
Form1.Show , Form5
Else
Form1.Show , Form5
Form1.SetFocus
End If
End If
If UseMe Is Nothing Then Exit Sub
UseMe.StopTimer2
End Sub

' TASK MASTER TIMING ROUTINE
Public Sub RunMe()
Dim Cancel As Integer
Static once
If once Then Exit Sub
once = True
UseMe.StopTimer
If InitOk = 0 Then
myToken = InitGDIPlus()
InitOk = 1
End If
backhere:
UseMe.CliRun
If UseMe Is Nothing Then Exit Sub

UseMe.Shutdown Cancel

If Cancel Then GoTo backhere
If Not UseMe.IhaveExtForm Then
UseMe.ShowGui = False
End If
Set UseMe = Nothing
If InitOk > 0 Then
InitOk = 0
FreeGDIPlus myToken
End If
once = False
End Sub



