Attribute VB_Name = "PicHandler"
Option Explicit
Private Declare Function HashData Lib "shlwapi" (ByVal straddr As Long, ByVal ByteSize As Long, ByVal res As Long, ByVal ressize As Long) As Long
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Any)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
Const VK_CAPITAL = &H14
Private Type KeyboardBytes
     kbByte(0 To 255) As Byte
End Type

Private Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
 
Dim kbArray As KeyboardBytes
Public fonttest As PictureBox
Private Declare Function GetTextMetrics Lib "gdi32" _
Alias "GetTextMetricsA" (ByVal hDC As Long, _
lpMetrics As TEXTMETRIC) As Long
Private Type TEXTMETRIC
tmHeight As Long
tmAscent As Long
tmDescent As Long
tmInternalLeading As Long
tmExternalLeading As Long
tmAveCharWidth As Long
tmMaxCharWidth As Long
tmWeight As Long
tmOverhang As Long
tmDigitizedAspectX As Long
tmDigitizedAspectY As Long
tmFirstChar As Byte
tmLastChar As Byte
tmDefaultChar As Byte
tmBreakChar As Byte
tmItalic As Byte
tmUnderlined As Byte
tmStruckOut As Byte
tmPitchAndFamily As Byte
tmCharSet As Byte
End Type
Dim TM As TEXTMETRIC
Public UseMe As Callback, byPassCallback As Boolean
Public osnum As Long

Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Integer)
Private Declare Function GetEnhMetaFileBits Lib "gdi32" (ByVal hmf As Long, ByVal nsize As Long, lpvData As Any) As Long
Private Declare Function CopyEnhMetaFile Lib "gdi32.dll" Alias "CopyEnhMetaFileW" (ByVal hemfSrc As Long, lpszFile As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long

Public MediaPlayer1 As New MovieModule
Public MediaBack1 As New MovieModule
Public form5iamloaded As Boolean
Public loadfileiamloaded As Boolean
Public sumhDC As Long  ' check it
Public Rixecode As String
Public MYSCRnum2stop As Long
Public octava As Integer, NOTA As Integer, ENTASI As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const Face$ = "C C#D D#E F F#G G#A A#B  "
Public CLICK_COUNT As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Enum Enum_OperatingPlatform
  Platform_Windows_32 = 0
  Platform_Windows_95_98_ME = 1
  Platform_Windows_NT_2K_XP = 2
End Enum
Public Enum Enum_OperatingSystem
   System_Windows_32 = 0
  System_Windows_95 = 1
  System_Windows_98 = 2
  System_Windows_ME = 3
  System_Windows_NT = 4
  System_Windows_2K = 5
  System_Windows_XP = 6
  System_Windows_Vista = 6
  System_Windows_7 = 7
  System_Windows_8 = 8
  System_Windows_81 = 9
  System_Windows_10 = 10
  System_Windows_New = 100
End Enum
Public PobjNum As Long

'*************************
Public Type TagSize
    cx As Long
    cy As Long
End Type
Declare Function GetAspectRatioFilterEx Lib "gdi32" (ByVal hDC As Long, lpAspectRatio As TagSize) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32.dll" (ByRef lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
Private Type XFORM  ' used for stretching/skewing a region
    eM11 As Single  ' note: some versions of this UDT have
    eM12 As Single  ' the elements as double -- wrong!
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type
Public Const RGN_OR = 2
'**********************************
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const Pi = 3.14159265359
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function timeGetTime Lib "kernel32.dll" Alias "GetTickCount" () As Long

Type Bitmap
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Declare Function RegisterClipboardFormat Lib "user32" Alias _
   "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private m_cfHTMLClipFormat As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" _
    (ByVal wFormat As Long) As Long
 Public Const CF_UNICODETEXT = 13
   Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" ( _
      ByVal pSecurityDescriptor As Long, _
      ByVal dwRevision As Long) As Long

   Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" ( _
      ByVal pSecurityDescriptor As Long, _
      ByVal bDaclPresent As Long, _
      ByVal pDacl As Long, _
      ByVal bDaclDefaulted As Long) As Long
 Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long


Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_DISCARDED = &H4000
Private Const GMEM_FIXED = &H0
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GMEM_LOCKCOUNT = &HFF
Private Const GMEM_MODIFY = &H80
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const GMEM_LOWER = GMEM_NOT_BANKED
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public frame As Boolean
Public PhotoBmp As Long
Public W As Long
Public BACKSPRITE As String
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Public Declare Function joyGetDevCapsA Lib "winmm.dll" (ByVal uJoyID As Long, pjc As JOYCAPS, ByVal cjc As Long) As Long

Public Type JOYCAPS
    wMid As Integer
    wPid As Integer
    szPname As String * 32
    wXmin As Long
    wXmax As Long
    wYmin As Long
    wYmax As Long
    wZmin As Long
    wZmax As Long
    wNumButtons As Long
    wPeriodMin As Long
    wPeriodMax As Long
    wRmin As Long
    wRmax As Long
    wUmin As Long
    wUmax As Long
    wVmin As Long
    wVmax As Long
    wCaps As Long
    wMaxAxes As Long
    wNumAxes As Long
    wMaxButtons As Long
    szRegKey As String * 32
    szOEMVxD As String * 260
End Type

Public Type JOYINFOEX
    dwSize As Long
    dwFlags As Long
    dwXpos As Long
    dwYpos As Long
    dwZpos As Long
    dwRpos As Long
    dwUpos As Long
    dwVpos As Long
    dwButtons As Long
    dwButtonNumber As Long
    dwPOV As Long
    dwReserved1 As Long
    dwReserved2 As Long
End Type
Public Type MYJOYSTATtype
enabled As Boolean
lngButton As Long
joyPaD As direction
AnalogX As Long
AnalogY As Long
Wait2Read As Boolean
End Type
Public MYJOYEX As JOYINFOEX
Public MYJOYSTAT(0 To 15) As MYJOYSTATtype

Public MYJOYCAPS As JOYCAPS

Public Enum direction
    DirectionNone = 0
    DirectionLeft = 1
    DirectionRight = 2
    DirectionUp = 3
    DirectionDown = 4
    DirectionLeftUp = 5
    DirectionLeftDown = 6
    DirectionRightUp = 7
    DirectionRightDown = 8
End Enum
Const LOCALE_IDEFAULTANSICODEPAGE As Long = &H1004
Const TCI_SRCCODEPAGE = 2
Private Type FONTSIGNATURE
    fsUsb(4) As Long
    fsCsb(2) As Long
End Type
Private Type CHARSETINFO
    ciCharset As Long
    ciACP As Long
    fs As FONTSIGNATURE
End Type
Private Declare Function TranslateCharsetInfo Lib "gdi32" ( _
    lpSrc As Long, _
    lpcs As CHARSETINFO, _
    ByVal dwFlags As Long _
) As Long
Public reopen4 As Boolean, reopen2 As Boolean
Public HelpFile As New Document, UseMDBHELP As Boolean, Form4Loaded As Boolean

Public Const HKL_PREV = 0
Public Const HKL_NEXT = 1

Public Const KLF_ACTIVATE = &H1
Public Const KLF_SUBSTITUTE_OK = &H2
Public Const KLF_UNLOADPREVIOUS = &H4
Public Const KLF_REORDER = &H8

''' Size of KeyboardLayoutName (number of characters), including nul terminator
Public Const KL_NAMELENGTH = 9

Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
Declare Function UnloadKeyboardLayout Lib "user32" (ByVal HKL As Long) As Long
Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal Flags As Long) As Long
Public Function HighLong(ByVal p) As Long
    If MemInt(VarPtr(p)) <> 20 Then p = cInt64(p)
    HighLong = MemLong(VarPtr(p) + 12)
End Function
Public Function LowLong(ByVal p) As Long
    If Not myVarType(p, 20) Then p = cInt64(p)
    LowLong = MemLong(VarPtr(p) + 8)
End Function
Function Hex64$(a, Optional bytes = 8)
    Dim p, p1, z
    z = cInt64(a)
    p = MemLong(VarPtr(z) + 8)
    p1 = MemLong(VarPtr(z) + 12)
    Hex64$ = Right$(Right$("0000000" + Hex$(p1), 8) + Right$("0000000" + Hex$(p), 8), bytes * 2)
End Function
Public Function OneLongLong() As Variant
    Static p
    If p = Empty Then
        PutMem2 VarPtr(p), 20
        PutMem1 VarPtr(p) + 8, 1
    End If
    OneLongLong = p
End Function
Public Function OneBigLongLong() As Variant
    Static p
    If p = Empty Then
        PutMem2 VarPtr(p), 20
        PutMem1 VarPtr(p) + 12, 1
    End If
    OneBigLongLong = p
End Function
Public Function MaskLowLongLong() As Variant
    Static p
    If p = Empty Then
    MemInt(VarPtr(p)) = 20
    MemLong(VarPtr(p) + 8) = -1&
    End If
    MaskLowLongLong = p
End Function
Public Function Signed(a) As Long
    Dim p
    p = Fix(CDec(a))
    Signed = MemLong(VarPtr(p) + 8)
End Function
Public Function UnsignedSub(a As Long, b As Long)
    Static ua, UB
    If ua = Empty Then
        MemInt(VarPtr(ua)) = 20
        MemInt(VarPtr(UB)) = 20
    End If
    MemLong(VarPtr(ua) + 8) = a
    MemLong(VarPtr(UB) + 8) = b
    ua = ua - UB
    UnsignedSub = MemLong(VarPtr(ua) + 8)
End Function

Public Function cInt64(p)
    Static maxlonglong, limitlonglong, OneLongLong, OneBigLongLong
    Dim a, i As Integer
    If MemInt(VarPtr(maxlonglong)) = 0 Then
        maxlonglong = CDec("18446744073709551616")
        limitlonglong = CDec("9223372036854775808")
        MemInt(VarPtr(OneLongLong)) = 20
        MemByte(VarPtr(OneLongLong) + 8) = 1
        MemInt(VarPtr(OneBigLongLong)) = 20
        MemByte(VarPtr(OneBigLongLong) + 12) = 1
    End If
    i = MemInt(VarPtr(p))
    Select Case i
    Case vbDecimal
        a = Fix(p)
        a = a - Int(a / maxlonglong) * maxlonglong
        If a < -limitlonglong - 1 Then
            While a <= -limitlonglong - 1: a = a + maxlonglong: Wend
        End If
        While a >= limitlonglong: a = a - maxlonglong: Wend
        cInt64 = -OneLongLong And a
    Case 20
        cInt64 = p
    Case vbLong, vbInteger
        cInt64 = -OneLongLong And p
    Case Else
        On Error GoTo er1
        
        a = Fix(CDec(p))
        If a > limitlonglong Or a <= -limitlonglong Then
        a = a - Int(a / (maxlonglong)) * (maxlonglong)
        If a <= -limitlonglong - 1 Then
            While a <= -limitlonglong - 1: a = a + maxlonglong: Wend
        End If
        While a >= limitlonglong: a = a - maxlonglong: Wend
        End If
        cInt64 = -OneLongLong And a
        If i = vbString Then
            If Left$(p, 1) = "&" And a < 0 Then
            Select Case Len(p)
            Case 10
                If InStr("89ABCDEF", UCase(Mid$(p, 3, 1))) > 0 Then
                cInt64 = OneBigLongLong + cInt64
                End If
            Case 18
                If Mid$(p, 3, 8) = "00000000" Then
                    cInt64 = OneBigLongLong + cInt64
                End If
            Case 11 To 17
            If cInt64 < 0 Then cInt64 = OneBigLongLong + cInt64
            End Select
            End If
         End If
    End Select
    Exit Function
er1:
    cInt64 = OneLongLong - OneLongLong
End Function
Public Sub PlaceIcon(a As StdPicture)
On Error Resume Next
If UseMe Is Nothing Then Exit Sub
UseMe.GetIcon a
End Sub
Public Sub PlaceCaption(ByVal a$)
Dim m As Callback, f As Form
On Error Resume Next
Set f = Screen.ActiveForm
If UseMe Is Nothing Then Exit Sub
    If Not UseMe.IamVisible Then
        If Len(a$) = 0 Then a$ = "M2000" Else Set f = Nothing
        Form1.CaptionW = a$
        If UseMe.IhaveExtForm Then
        UseMe.SetExtCaption a$
        Else
        
            Form3.Timer1.Interval = 30
            Form3.Timer1.enabled = True
            Form3.CaptionWsilent = a$
            Form3.CaptionW = a$
            Form1.CaptionW = vbNullString
            Form3.WindowState = 0
        End If
Else
    If a$ = vbNullString Then
        If UseMe.IhaveExtForm Then UseMe.SetExtCaption a$
        Form1.CaptionW = vbNullString
        Form1.Visible = False
    Else
        If UseMe.IhaveExtForm Then
            UseMe.SetExtCaption a$
        End If
        Form1.CaptionW = a$
       
        
        If f Is Form1 Then
            If Form1.Visible Then
            Form1.SetFocus
            Else
            If UseMe.WindowState > 0 Then
                
                Form1.Visible = True
                Form1.SetFocus
            End If
            End If
         End If
       Set f = Nothing
       
    End If
    ttl = False
End If
ttl = False
Exit Sub
If Not f Is Nothing Then
    If f Is Form1 Then
        If Form1.Visible Then
            Form1.SetFocus
         ElseIf Not UseMe Is Nothing Then
            If Form1.TrueVisible Then
                If UseMe.IhaveExtForm Then
                    UseMe.ExtWindowState = 1
                End If
            End If
        End If
    Else
        If f.Visible Then f.SetFocus
    End If
End If
Err.Clear
End Sub
Public Function StartJoypadk(Optional ByVal jn As Long = 0) As Boolean
    If joyGetDevCapsA(jn, MYJOYCAPS, 404) <> 0 Then 'Get Joypadk info
    MYJOYSTAT(jn).enabled = False
        StartJoypadk = False
    Else
        MYJOYEX.dwSize = 64
        MYJOYEX.dwFlags = 255
        Call joyGetPosEx(jn, MYJOYEX)
        MYJOYSTAT(jn).Wait2Read = False
         MYJOYSTAT(jn).enabled = True
        StartJoypadk = True
    End If
End Function
Public Sub ClearJoyAll()

Dim jn As Long
For jn = 0 To 15
MYJOYSTAT(jn).Wait2Read = False
Next jn
End Sub
Public Sub FlushJoyAll()

Dim jn As Long
For jn = 0 To 15
MYJOYSTAT(jn).enabled = False
Next jn
End Sub

Public Sub PollJoypadk()

    Dim jn As Long, wh As Long
    ' Get the Joypadk information
    For jn = 0 To 15
    If MYJOYSTAT(jn).enabled Then
    If Not MYJOYSTAT(jn).Wait2Read Then
      MYJOYEX.dwSize = 64
    MYJOYEX.dwFlags = 255
    Call joyGetPosEx(jn, MYJOYEX)
    wh = MYJOYEX.dwButtons
    
     With MYJOYSTAT(jn)
     .Wait2Read = False
        If wh <> 0 Then .lngButton = (Log(wh) / Log(2)) + 1 Else .lngButton = 0
            .AnalogX = MYJOYEX.dwXpos
            .AnalogY = MYJOYEX.dwYpos
            If (MYJOYEX.dwXpos = 0 And MYJOYEX.dwYpos = 0) Then
            .joyPaD = DirectionLeftUp
        ElseIf (MYJOYEX.dwXpos = 0 And MYJOYEX.dwYpos = 65535) Then
            .joyPaD = DirectionLeftDown
        ElseIf (MYJOYEX.dwXpos = 65535 And MYJOYEX.dwYpos = 0) Then
            .joyPaD = DirectionRightUp
        ElseIf (MYJOYEX.dwXpos = 65535 And MYJOYEX.dwYpos = 65535) Then
            .joyPaD = DirectionRightDown
        ElseIf (MYJOYEX.dwXpos = 0) Then
            .joyPaD = DirectionLeft
        ElseIf (MYJOYEX.dwXpos = 65535) Then
            .joyPaD = DirectionRight
        ElseIf (MYJOYEX.dwYpos = 0) Then
            .joyPaD = DirectionUp
        ElseIf (MYJOYEX.dwYpos = 65535) Then
            .joyPaD = DirectionDown
        Else
            .joyPaD = DirectionNone
        End If
          .Wait2Read = True
        End With
    End If
    End If
    Next jn
End Sub

Public Function OperatingPlatform() As Enum_OperatingPlatform
    Dim lpVersionInformation As OSVERSIONINFO
    lpVersionInformation.dwOSVersionInfoSize = Len(lpVersionInformation)
    Call GetVersionExA(lpVersionInformation)
    OperatingPlatform = lpVersionInformation.dwPlatformId
End Function
Public Function OperatingSystem() As Enum_OperatingSystem

Dim lpVersionInformation As OSVERSIONINFO
If osnum = 0 Then
    
    lpVersionInformation.dwOSVersionInfoSize = Len(lpVersionInformation)
    Call GetVersionExA(lpVersionInformation)


  If (lpVersionInformation.dwPlatformId = Platform_Windows_32) Then

        osnum = System_Windows_32
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 0) Then
        osnum = System_Windows_95
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 10) Then
        osnum = System_Windows_98
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 90) Then
        osnum = System_Windows_ME
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion < 5) Then
        osnum = System_Windows_NT
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion = 0) Then
        osnum = System_Windows_2K
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion >= 1) Then
        osnum = System_Windows_XP
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 6) And (lpVersionInformation.dwMinorVersion = 0) Then
        osnum = System_Windows_Vista
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 6) And (lpVersionInformation.dwMinorVersion = 1) Then
        osnum = System_Windows_7
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 6) And (lpVersionInformation.dwMinorVersion = 2) Then
        osnum = System_Windows_8
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 6) And (lpVersionInformation.dwMinorVersion = 3) Then
        osnum = System_Windows_81
      ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 10) And (lpVersionInformation.dwMinorVersion = 0) Then
        osnum = System_Windows_10
    ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion >= 10) And (lpVersionInformation.dwMinorVersion >= 0) Then
        osnum = System_Windows_New
        Else
               osnum = System_Windows_32
    End If
    End If
  OperatingSystem = osnum
End Function
Public Function os() As String
  os = OsInfo.OSName
End Function
Public Function Platform() As String
    Platform = OsInfo.Platform
End Function
Function check_mem() As Long

    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim MemStat As MEMORYSTATUS
    'retrieve the memory status
    GlobalMemoryStatus MemStat
    check_mem = MemStat.dwAvailPhys \ 1024 \ 1024
 End Function
'
'
' Implemantation of string bitmaps
' Width - Heigth - DATA
Public Function cDib(a, mdib As cDIBSection) As Boolean
On Error GoTo e1111
cDib = False
If Len(a) >= 12 Then
' read magicNo, witdh, height
If Left$(a, 4) = "cDIB" Then
Dim W As Long, H As Long
W = val("&H" & Mid$(a, 5, 4))
H = val("&H" & Mid$(a, 9, 4))
If Len(a) * 2 < ((W * 3 + 3) \ 4) * 4 * H - 24 Then Exit Function
mdib.ClearUp

If mdib.create(W, H) Then
If Len(a) * 2 < mdib.BytesPerScanLine * H + 24 Then Exit Function
CopyMemory ByVal mdib.DIBSectionBitsPtr, ByVal StrPtr(a) + 24, mdib.BytesPerScanLine * H
cDib = True
End If
End If
End If
e1111:
End Function
Function CDib2Pic(a) As StdPicture
Dim aa As New cDIBSection, emptypic As New StdPicture
If cDib(a, aa) Then
    Set CDib2Pic = aa.Picture()
Else
    Set CDib2Pic = emptypic
End If
End Function
Public Function SetDIBPixel(ssdib As Variant, ByVal X As Long, ByVal Y As Long, aColor As Long) As Double
Dim W As Long, H As Long, bpl As Long, rgb(2) As Byte
W = val("&H" & Mid$(ssdib, 5, 4))
H = val("&H" & Mid$(ssdib, 9, 4))
X = W - Abs(X) - 1
Y = Abs(H - Y - 1) Mod H

If Len(ssdib) * 2 < ((W * 3 + 3) \ 4) * 4 * H - 24 Then Exit Function
If W * H <> 0 Then
bpl = (LenB(ssdib) - 24) \ H
W = (W - X - 1) Mod W
H = Y * bpl + W * 3 + 24
CopyMemory rgb(0), ByVal StrPtr(ssdib) + H, 3
W = rgb(0): rgb(0) = rgb(2): rgb(2) = W
bpl = 0
CopyMemory ByVal VarPtr(bpl), rgb(0), 3
SetDIBPixel = -1# * bpl
CopyMemory rgb(0), ByVal VarPtr(aColor), 3
W = rgb(0): rgb(0) = rgb(2): rgb(2) = W
CopyMemory ByVal StrPtr(ssdib) + H, rgb(0), 3
End If
End Function
Public Function GetDIBPixel(ssdib As Variant, ByVal X As Long, ByVal Y As Long) As Double
Dim W As Long, H As Long, bpl As Long, rgb(2) As Byte
'a = ssdib$
W = val("&H" & Mid$(ssdib, 5, 4))
H = val("&H" & Mid$(ssdib, 9, 4))
X = W - Abs(X) - 1
Y = Abs(H - Y - 1) Mod H
If Len(ssdib) * 2 < ((W * 3 + 3) \ 4) * 4 * H - 24 Then Exit Function
If W * H <> 0 Then
bpl = (LenB(ssdib) - 24) \ H   ' Len(ssdib$) 2 bytes per char
W = (W - X - 1) Mod W

H = Y * bpl + W * 3 + 24
CopyMemory rgb(0), ByVal StrPtr(ssdib) + H, 3
W = rgb(0): rgb(0) = rgb(2): rgb(2) = W
bpl = 0
CopyMemory ByVal VarPtr(bpl), rgb(0), 3
GetDIBPixel = -1# * bpl

'GetDIBPixel = -(rgb(0) * 256# * 256# + rgb(1) * 256# + rgb(2))
End If
End Function
Public Function cDIBwidth1(a) As Long
Dim W As Long, H As Long
cDIBwidth1 = -1

W = val("&H" & Mid$(a, 5, 4))
H = val("&H" & Mid$(a, 9, 4))
If Len(a) * 2 < ((W * 3 + 3) \ 4) * 4 * H - 24 Then Exit Function
cDIBwidth1 = W
End Function
Public Function cDIBwidth(a) As Long
Dim W As Long, H As Long
cDIBwidth = -1
If Len(a) >= 12 Then
If Left$(a, 4) = "cDIB" Then
W = val("&H" & Mid$(a, 5, 4))
H = val("&H" & Mid$(a, 9, 4))
If Len(a) * 2 < ((W * 3 + 3) \ 4) * 4 * H - 24 Then Exit Function
cDIBwidth = W
End If
End If
End Function
Public Function cDIBheight1(a) As Long
Dim W As Long, H As Long
cDIBheight1 = -1
W = val("&H" & Mid$(a, 5, 4))
H = val("&H" & Mid$(a, 9, 4))
If Len(a) * 2 < ((W * 3 + 3) \ 4) * 4 * H - 24 Then Exit Function
cDIBheight1 = H
End Function
Public Function cDIBheight(a) As Long
Dim W As Long, H As Long
cDIBheight = -1
If Len(a) >= 12 Then
If Left$(a, 4) = "cDIB" Then
W = val("&H" & Mid$(a, 5, 4))
H = val("&H" & Mid$(a, 9, 4))
If Len(a) * 2 < ((W * 3 + 3) \ 4) * 4 * H - 24 Then Exit Function
cDIBheight = H
End If
End If
End Function

Public Function ARRAYtoStr(ffff() As Byte) As String
Dim a As String, j As Long
For j = 1 To UBound(ffff())
a = a + ChrW(ffff(j))
Next j
ARRAYtoStr = a
End Function
Public Sub LoadArray(ffff() As Byte, a As String)
Dim j As Long
ReDim ffff(1 To Len(a)) As Byte
For j = 1 To UBound(ffff())
ffff(j) = CByte(AscW(Mid$(a, j, 1)) And &HFF)
Next j

End Sub
Public Function GetTag$()
Dim ss$, j As Long
''
For j = 1 To 16
ss$ = ss$ & Chr(65 + Int((23 * Rnd) + 1))
Next j
GetTag$ = ss$
End Function

Public Function DIBtoSTR(mdib As cDIBSection) As String
Dim a As String
If mdib.Width > 0 Then
a = String$(mdib.BytesPerScanLine * mdib.Height \ 2 + 12, Chr(0))
Mid$(a, 1, 12) = "cDIB" + Right$("0000" + Hex$(mdib.Width), 4) + Right$("0000" + Hex$(mdib.Height), 4)
CopyMemory ByVal StrPtr(a) + 24, ByVal mdib.DIBSectionBitsPtr, mdib.BytesPerScanLine * mdib.Height
DIBtoSTR = a
End If
End Function
Public Function DpiScrX() As Long
Dim lhWNd As Long, lHDC As Long
    lhWNd = GetDesktopWindow()
    lHDC = GetDC(lhWNd)
    DpiScrX = GetDeviceCaps(lHDC, LOGPIXELSX)
    ReleaseDC lhWNd, lHDC
End Function

Public Function bitsPerPixel() As Long
Dim lhWNd As Long, lHDC As Long
    lhWNd = GetDesktopWindow()
    lHDC = GetDC(lhWNd)
    bitsPerPixel = GetDeviceCaps(lHDC, BITSPIXEL)
    ReleaseDC lhWNd, lHDC
End Function
Public Function RotateMaskDib(cDibbuffer0 As cDIBSection, Optional ByVal Angle! = 0, Optional ByVal zoomfactor As Single = 100, _
    Optional bckColor As Long = &HFFFFFF, Optional Alpha As Long = 100)
    Dim ang As Long
    ang = CLng(Angle!)
Angle! = -(CLng(Angle!) Mod 360) * 1.745329E-02!
If cDibbuffer0.hDib = 0 Then Exit Function
If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100#
Dim myw As Long, myh As Long, piw As Long, pih As Long, pix As Long, piy As Long
Dim a As Single, b As Single, k As Single, R As Single
Dim br As Byte, bG As Byte, bbb As Byte ', ba$
Dim BR1 As Byte, BG1 As Byte, bbb1 As Byte, ppBa As Long
BR1 = 255 * ((100 - Alpha) / 100#)
BG1 = 255 * ((100 - Alpha) / 100#)
bbb1 = 255 * ((100 - Alpha) / 100#)
ppBa = VarPtr(bckColor)
GetMem1 ppBa, bbb
GetMem1 ppBa + 1, bG
GetMem1 ppBa + 2, br

'ba$ = Hex$(bckColor)
'ba$ = Right$("00000" & ba$, 6)
'BR = val("&h" & Mid$(ba$, 1, 2))
'BG = val("&h" & Mid$(ba$, 3, 2))
'bbb = val("&h" & Mid$(ba$, 5, 2))
Dim pw As Long, ph As Long
    piw = cDibbuffer0.Width
    pih = cDibbuffer0.Height
    R = Atn(CSng(piw) / CSng(pih)) + Pi / 2#
     k = Fix(Abs((piw / Cos(R) / 2) * zoomfactor) + 0.5)

Dim cDIBbuffer1 As Object
 Dim olddpix As Long, olddpiy As Long
 olddpix = cDibbuffer0.dpix
 olddpiy = cDibbuffer0.dpiy
 myw = 2 * k
myh = 2 * k

    pw = cDibbuffer0.Width
    ph = cDibbuffer0.Height
 cDibbuffer0.ClearUp
Call cDibbuffer0.create(myw, myh)
cDibbuffer0.GetDpi olddpix, olddpiy
cDibbuffer0.Cls bckColor

there:
Dim bDib2() As Byte, bDib1() As Byte
Dim X As Long, Y As Long
Dim lc As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
On Error Resume Next

 '   cDIBbuffer0.WhiteBits
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibbuffer0.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibbuffer0.BytesPerScanLine()
        .pvData = cDibbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4

    Dim nx As Long, ny As Long
    Dim image_x As Long, image_y As Long, temp_image_x As Long, temp_image_y As Long
    Dim x_step As Long, y_step As Long, x_step2 As Long, y_step2 As Long
    Dim screen_x As Long, screen_y As Long, mmx As Long, mmy As Long


    
 
       Dim pw1 As Long, ph1 As Long
          Dim sX As Single, sY As Single
    Dim xf As Single, yf As Single
    Dim xf1 As Single, yf1 As Single
    Dim pws As Single, phs As Single
    pw1 = pw
    ph1 = ph
    pws = pw
    phs = ph
    R = Atn(CSng(myw) / CSng(myh))
    k = -myw / (2# * Sin(R))
    

       x_step2 = CLng(Fix(Cos(Angle! + Pi / 2) * pw))
    y_step2 = CLng(Fix(Sin(Angle! + Pi / 2) * ph))

    x_step = CLng(Fix(Cos(Angle!) * pw))
    y_step = CLng(Fix(Sin(Angle!) * ph))
  image_x = CLng(Fix(pw / 2 - Fix(k * Sin(Angle! - R)))) * pw
   image_y = CLng(Fix(ph / 2 + Fix(k * Cos(Angle! - R)))) * ph
Dim pw1out As Long, ph1out As Long, pwOut As Long, phOut As Long, much As Single
''Dim cw1 As Long, ch1 As Long, outf As Single, fadex As Long, fadey As Long, outf1 As Single, outf2 As Single
pw1 = pw1 - 1
ph1 = ph1 - 1
pw1out = pw1 - 1
ph1out = ph1 - 1

Dim nomalo As Boolean
nomalo = Not (ang Mod 90 = 0)
    For screen_y = 0 To myh - 1
        temp_image_x = image_x
        temp_image_y = image_y
        For screen_x = 0 To (myw - 1) * 3 Step 3
  
                  sX = temp_image_x / pws
                sY = temp_image_y / phs
                mmx = Int(sX)
                mmy = Int(sY)

           
                    If mmx >= 1 And mmx <= pw1out And mmy >= 1 And mmy <= ph1out Then
          xf = (sX - CSng(mmx))
             xf1 = (1! - xf)
                      yf = (sY - CSng(mmy))
                      yf1 = 1! - yf
                  
                   
                      bDib1(screen_x, screen_y) = BR1
                        
                        bDib1(screen_x + 1, screen_y) = BR1
                       bDib1(screen_x + 2, screen_y) = BR1
                        If nomalo Then
                      If mmx <= 1 Then
                      
                      bDib1(screen_x, screen_y) = br * xf1
                        
                        bDib1(screen_x + 1, screen_y) = br * xf1  ' * yf / 2
                       bDib1(screen_x + 2, screen_y) = br * xf1 '* yf / 2
                       ElseIf mmx >= pw1out Then
                        bDib1(screen_x, screen_y) = br * xf
                        
                        bDib1(screen_x + 1, screen_y) = br * xf

                        bDib1(screen_x + 2, screen_y) = br * xf
                       End If
                       If mmy >= ph1out Then
                         bDib1(screen_x, screen_y) = br * yf
                        
                        bDib1(screen_x + 1, screen_y) = br * yf
                       bDib1(screen_x + 2, screen_y) = br * yf
                       ElseIf mmy <= 1 Then
                          bDib1(screen_x, screen_y) = br * yf1
                        
                        bDib1(screen_x + 1, screen_y) = br * yf1
                       bDib1(screen_x + 2, screen_y) = br * yf1
                      End If
               
                 End If
                    End If
            temp_image_x = temp_image_x + x_step
            temp_image_y = temp_image_y + y_step
       Next screen_x
        image_x = image_x + x_step2
        image_y = image_y + y_step2
    Next screen_y
    
   
  
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
     
End Function

Public Function Merge3Dib(backdib As cDIBSection, maskdib As cDIBSection, frontdib As cDIBSection, Optional Reverse As Boolean = False)

Dim X As Long, Y As Long

Dim xmax As Long, yMax As Long
    yMax = backdib.Height - 1
    xmax = backdib.Width - 1
Dim bDib() As Byte
Dim tSA As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = backdib.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = backdib.BytesPerScanLine()
        .pvData = backdib.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    
Dim bDib1() As Byte
Dim tSA1 As SAFEARRAY2D
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = maskdib.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = maskdib.BytesPerScanLine()
        .pvData = maskdib.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4
    
Dim bDib2() As Byte
Dim tSA2 As SAFEARRAY2D
    With tSA2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = frontdib.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = frontdib.BytesPerScanLine()
        .pvData = frontdib.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib2()), VarPtr(tSA2), 4
        '-----------------------------------------------
        If Reverse Then
        
    For X = 0 To (xmax * 3) Step 3
        For Y = yMax To 0 Step -1
            bDib(X, Y) = (CLng(bDib(X, Y)) * bDib1(X, Y) + CLng(bDib2(X, Y)) * (255 - bDib1(X, Y))) \ 256
            bDib(X + 1, Y) = (CLng(bDib(X + 1, Y)) * bDib1(X + 1, Y) + CLng(bDib2(X + 1, Y)) * (255 - bDib1(X + 1, Y))) \ 256
            bDib(X + 2, Y) = (CLng(bDib(X + 2, Y)) * bDib1(X + 2, Y) + CLng(bDib2(X + 2, Y)) * (255 - bDib1(X + 2, Y))) \ 256
        Next Y
        Next X
        Else
     For X = 0 To (xmax * 3) Step 3
        For Y = yMax To 0 Step -1
            bDib(X, Y) = (CLng(bDib2(X, Y)) * bDib1(X, Y) + CLng(bDib(X, Y)) * (255 - bDib1(X, Y))) \ 256
            bDib(X + 1, Y) = (CLng(bDib2(X + 1, Y)) * bDib1(X + 1, Y) + CLng(bDib(X + 1, Y)) * (255 - bDib1(X + 1, Y))) \ 256
            bDib(X + 2, Y) = (CLng(bDib2(X + 2, Y)) * bDib1(X + 2, Y) + CLng(bDib(X + 2, Y)) * (255 - bDib1(X + 2, Y))) \ 256
        Next Y
        Next X
        End If

   '-----------------------------------------------
     CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
        CopyMemory ByVal VarPtrArray(bDib2), 0&, 4
 End Function

Public Sub CanvasSize(cDibbuffer0 As cDIBSection, ByVal wcm As Double, ByVal hcm As Double, Optional ByVal rep As Boolean = False, Optional Max As Integer = 0, Optional yshift As Long = 0, Optional bColor As Long = &HFFFFFF, Optional usepixel As Boolean = False, Optional ByVal Percent As Single = 85, Optional ByVal linewidth As Long = 4)
' top left align only
Dim piw As Long, pih As Long, stx As Long, sty As Long, stOffx As Long, stOffy As Long, stBorderX As Long, stBorderY As Long, strx As Long, stry As Long, i As Long, j As Long

Dim cDIBbuffer1 As New cDIBSection
If Not usepixel Then
piw = CLng(wcm * cDibbuffer0.dpix / 2.54)
pih = CLng(hcm * cDibbuffer0.dpiy / 2.54)
Else
piw = wcm
pih = hcm
End If
If cDIBbuffer1.create(piw, pih) Then
    cDIBbuffer1.Cls bColor
    cDIBbuffer1.GetDpiDIB cDibbuffer0
    
     stx = 0: sty = 0
     If rep Then
      cDIBbuffer1.needHDC
     stOffx = cDIBbuffer1.Width Mod cDibbuffer0.Width
     stOffy = cDIBbuffer1.Height Mod cDibbuffer0.Height
     strx = cDIBbuffer1.Width \ cDibbuffer0.Width
     stry = cDIBbuffer1.Height \ cDibbuffer0.Height
     stBorderX = stOffx \ (strx + 1)
     stBorderY = stOffy \ (stry + 1)
                If Max = 0 Then Max = strx * stry
       sty = stBorderY
                For j = 1 To stry
                stx = stBorderX
                             For i = 1 To strx
                           
                            If Max = 0 Then Exit For
                            cDibbuffer0.PaintPicture cDIBbuffer1.HDC1, stx, sty + yshift
                            Max = Max - 1
                               stx = stx + cDibbuffer0.Width + stBorderX
                           
                            Next i
                 If Max = 0 Then Exit For
                   sty = sty + cDibbuffer0.Height + stBorderY
                Next j
                cDIBbuffer1.FreeHDC
     ElseIf usepixel Then
     
     cDibbuffer0.ThumbnailPaintdib cDIBbuffer1, Percent, , , , , , , linewidth
     
     Else
      cDIBbuffer1.needHDC
            cDibbuffer0.PaintPicture cDIBbuffer1.HDC1, stx, sty + yshift
            cDIBbuffer1.FreeHDC
     End If
    
     
     Set cDibbuffer0 = cDIBbuffer1
    End If
End Sub

Public Sub RotateDibNew(cDibbuffer0 As cDIBSection, Optional ByVal Angle! = 0, _
        Optional ByVal zoomfactor As Single = 1, Optional bckColor As Long = &HFFFFFF)
Const Pi = 3.14159!
Dim myw As Long, myh As Long, piw As Long, pih As Long, pix As Long, piy As Long
Dim k As Single, R As Single
Dim br As Byte, bG As Byte, bbb As Byte, ppBa As Long
Dim bDib() As Byte, bDib1() As Byte
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
Dim image_x As Single, image_y As Single, temp_image_x As Single, temp_image_y As Single
Dim x_step As Single, y_step As Single, x_step2 As Single, y_step2 As Single
Dim screen_x As Long, screen_y As Long, mmx As Long, mmy As Long
Dim pw As Long, ph As Long
Dim sX As Single, sY As Single
Dim xf As Single, yf As Single
Dim xf1 As Single, yf1 As Single
Dim pws As Single, phs As Single
Dim pw1 As Long, ph1 As Long

Dim b As Single
b = CSng(Angle! Mod 90 = 0)
Angle! = -MyMod(Angle!, 360!) * 1.745329E-02!
On Error Resume Next
If cDibbuffer0.hDib = 0 Then Exit Sub
If zoomfactor <= 0.01! Then zoomfactor = 0.01!


ppBa = VarPtr(bckColor)
GetMem1 ppBa, bbb
GetMem1 ppBa + 1, bG
GetMem1 ppBa + 2, br
piw = cDibbuffer0.Width
pih = cDibbuffer0.Height
R = Atn(piw / pih) + Pi / 2!
k = Abs((piw / Cos(R) / 2!) * zoomfactor)
Dim cDIBbuffer1 As Object
Set cDIBbuffer1 = New cDIBSection
If piw <= 1 Then piw = 2
If pih <= 1 Then pih = 2
Call cDIBbuffer1.create((piw) * zoomfactor, (pih) * zoomfactor)
cDIBbuffer1.GetDpiDIB cDibbuffer0
cDibbuffer0.needHDC
cDIBbuffer1.LoadPictureStretchBlt cDibbuffer0.HDC1, , , , , pix, piy, piw, pih
cDibbuffer0.FreeHDC
myw = Fix(2 * k)
myh = Fix(2 * k)
cDibbuffer0.ClearUp
If cDibbuffer0.create(CLng(myw), CLng(myh)) Then
    On Error Resume Next
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    cDibbuffer0.WhiteBits
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibbuffer0.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibbuffer0.BytesPerScanLine()
        .pvData = cDibbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4


    pw = cDIBbuffer1.Width
    ph = cDIBbuffer1.Height

    R = Atn(CSng(myw) / CSng(myh))
    k = -CSng(myw) / (2! * Sin(R))
  
    Const pidicv2 = 1.570795!
    pw1 = pw + 1
    ph1 = ph + 1
   
    pws = pw1 * zoomfactor
    phs = ph1 * zoomfactor
    image_x = ((pws - zoomfactor - b) / 2 - (k * Sin(Angle! - R))) * pw
    image_y = ((phs - zoomfactor - b) / 2 + (k * Cos(Angle! - R))) * ph
    image_x = image_x - MyMod(image_x, CSng(dv15))
    image_y = image_y - MyMod(image_y, CSng(dv15))
   

    x_step2 = Cos(Angle! + pidicv2) * pw
    y_step2 = Sin(Angle! + pidicv2) * ph
    
    x_step = Cos(Angle!) * pw
    y_step = Sin(Angle!) * ph
    
    pws = pws + 1
    phs = phs + 1
    pw1 = pw - 1
    ph1 = ph - 1
    
    For screen_y = 0 To myh - 1
        temp_image_x = image_x
        temp_image_y = image_y
        For screen_x = 0 To (myw - 1) * 3 Step 3
            sX = temp_image_x / pws
            sY = temp_image_y / phs
            mmx = Int(sX)
            mmy = Int(sY)
            If mmx >= 0 And mmx < pw1 And mmy >= 0 And mmy < ph1 Then
                 xf = Abs((sX - CSng(mmx)))
                 xf1 = 1! - xf
                 yf = Abs((sY - CSng(mmy)))
                 yf1 = 1! - yf
                 mmx = mmx * 3
                 bDib1(screen_x, screen_y) = yf1 * (xf1 * bDib(mmx, mmy) + xf * bDib(mmx + 3, mmy)) + yf * (xf1 * bDib(mmx, mmy + 1) + xf * bDib(mmx + 3, mmy + 1))
                 bDib1(screen_x + 1, screen_y) = yf1 * (xf1 * bDib(mmx + 1, mmy) + xf * bDib(mmx + 4, mmy)) + yf * (xf1 * bDib(mmx + 1, mmy + 1) + xf * bDib(mmx + 4, mmy + 1))
                 bDib1(screen_x + 2, screen_y) = yf1 * (xf1 * bDib(mmx + 2, mmy) + xf * bDib(mmx + 5, mmy)) + yf * (xf1 * bDib(mmx + 2, mmy + 1) + xf * bDib(mmx + 5, mmy + 1))
             Else
                 bDib1(screen_x, screen_y) = br
                 bDib1(screen_x + 1, screen_y) = bG
                 bDib1(screen_x + 2, screen_y) = bbb
             End If
            temp_image_x = temp_image_x + x_step
            temp_image_y = temp_image_y + y_step
       Next screen_x
        image_x = image_x + x_step2
        image_y = image_y + y_step2
    Next screen_y
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
End If

Set cDIBbuffer1 = Nothing
End Sub
Public Function GetBackSpriteHDC(bstack As basetask, thisHDC As Long, piw As Long, pih As Long, Optional ByVal Angle! = 0, Optional ByVal zoomfactor As Single = 100)
    ' piw, pih pixels
    
Angle! = -MyMod(Angle!, 360!) * 1.74532925199433E-02
If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100#
Dim myw As Long, myh As Long
  
myw = Round((Abs(piw * Cos(Angle!)) + Abs(pih * Sin(Angle!))) * zoomfactor, 0) + 4
myh = Round((Abs(piw * Sin(Angle!)) + Abs(pih * Cos(Angle!))) * zoomfactor, 0) + 4
Dim prive As basket
prive = players(GetCode(bstack.Owner))
Dim cDibbuffer0 As New cDIBSection
If cDibbuffer0.create(myw, myh) Then
On Error GoTo there
        With bstack.Owner
         If bstack.toprinter Then
         cDibbuffer0.LoadPictureBlt thisHDC, Int(.ScaleX(prive.XGRAPH, 0, 3) - myw \ 2), Int(.ScaleX(prive.YGRAPH, 0, 3) - myh \ 2)
         Else
        cDibbuffer0.LoadPictureBlt thisHDC, Int(.ScaleX(prive.XGRAPH, 1, 3) - myw \ 2), Int(.ScaleX(prive.YGRAPH, 1, 3) - myh \ 2)
            End If
            BACKSPRITE = DIBtoSTR(cDibbuffer0)
        End With
End If
there:
End Function

'
Public Function GetBackSprite(bstack As basetask, piw As Long, pih As Long, Optional ByVal Angle! = 0, Optional ByVal zoomfactor As Single = 100)
    ' piw, pih pixels
    
Angle! = -MyMod(Angle!, 360!) * 1.74532925199433E-02
If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100#
Dim myw As Long, myh As Long
  
myw = Round((Abs(piw * Cos(Angle!)) + Abs(pih * Sin(Angle!))) * zoomfactor, 0) + 4
myh = Round((Abs(piw * Sin(Angle!)) + Abs(pih * Cos(Angle!))) * zoomfactor, 0) + 4
Dim prive As basket
prive = players(GetCode(bstack.Owner))
Dim cDibbuffer0 As New cDIBSection
If cDibbuffer0.create(myw, myh) Then
On Error GoTo there
        With bstack.Owner
         If bstack.toprinter Then
         cDibbuffer0.LoadPictureBlt bstack.Owner.hDC, Int(.ScaleX(prive.XGRAPH, 0, 3) - myw \ 2), Int(.ScaleX(prive.YGRAPH, 0, 3) - myh \ 2)
         Else
        cDibbuffer0.LoadPictureBlt bstack.Owner.hDC, Int(.ScaleX(prive.XGRAPH, 1, 3) - myw \ 2), Int(.ScaleX(prive.YGRAPH, 1, 3) - myh \ 2)
            End If
            BACKSPRITE = DIBtoSTR(cDibbuffer0)
        End With
End If
there:
End Function

Private Function MyMod(r1 As Single, po As Single) As Single
MyMod = r1 - Fix(r1 / po) * po
End Function
'
Public Function RotateDib(bstack As basetask, cDibbuffer0 As cDIBSection, Optional ByVal Angle! = 0, Optional ByVal zoomfactor As Single = 100, _
    Optional bckColor As Long = -1, Optional nogetback As Boolean = False, Optional Alpha As Long = 100, Optional amask$ = vbNullString)
    Const Pi = 3.14159!
     Dim b As Single
   b = CSng(CLng(Angle!) Mod 90 = 0)
Angle! = -MyMod(Angle!, 360!) * 1.745329E-02!
Const pidicv2 = 1.570795!
If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100#
Dim cDIBbuffer1 As cDIBSection, cDIBbuffer2 As cDIBSection
If zoomfactor < 1! Then
    If amask$ <> "" Then
            If Left$(amask$, 4) = "cDIB" Then
                Set cDIBbuffer1 = New cDIBSection
                If Not cDib(amask$, cDIBbuffer1) Then
                    Set cDIBbuffer1 = Nothing
                    GoTo exitHere
                End If
                Set cDIBbuffer2 = New cDIBSection
                cDIBbuffer2.CreateFromPicture cDIBbuffer1.Picture2(zoomfactor)
            End If
    End If
    Set cDIBbuffer1 = New cDIBSection
    cDIBbuffer1.CreateFromPicture cDibbuffer0.Picture2(zoomfactor)
    Set cDibbuffer0 = cDIBbuffer1
    Set cDIBbuffer1 = Nothing
    zoomfactor = 1
ElseIf amask$ <> "" Then
        If Left$(amask$, 4) = "cDIB" Then
        Set cDIBbuffer2 = New cDIBSection
        If Not cDib(amask$, cDIBbuffer2) Then
            Set cDIBbuffer2 = Nothing
            GoTo exitHere
        End If
        End If
End If









Dim myw As Long, myh As Long, piw As Long, pih As Long, pix As Long, piy As Long
Dim k As Single, R As Single, ppBa As Long
Dim br As Byte, bG As Byte, bbb As Byte, ba$
Dim BR1 As Byte, BG1 As Byte, bbb1 As Byte
ppBa = VarPtr(bckColor)
GetMem1 ppBa, bbb
GetMem1 ppBa + 1, bG
GetMem1 ppBa + 2, br

    piw = cDibbuffer0.Width
    pih = cDibbuffer0.Height
 Set cDIBbuffer1 = cDibbuffer0 'New cDIBSection
 Set cDibbuffer0 = New cDIBSection
myw = Round((Abs(piw * Cos(Angle!)) + Abs(pih * Sin(Angle!))) * zoomfactor, 0)
myh = Round((Abs(piw * Sin(Angle!)) + Abs(pih * Cos(Angle!))) * zoomfactor, 0)
cDibbuffer0.ClearUp
Dim prive As basket
prive = players(GetCode(bstack.Owner))
If cDibbuffer0.create(myw, myh) Then
On Error GoTo there

   
With bstack.Owner
    If bstack.toprinter Then
        cDibbuffer0.LoadPictureBlt bstack.Owner.hDC, Int(.ScaleX(prive.XGRAPH, 0, 3) - myw \ 2), Int(.ScaleX(prive.YGRAPH, 0, 3) - myh \ 2)
    Else
        cDibbuffer0.LoadPictureBlt bstack.Owner.hDC, Int(.ScaleX(prive.XGRAPH, 1, 3) - myw \ 2), Int(.ScaleX(prive.YGRAPH, 1, 3) - myh \ 2)
    End If
    If Not nogetback Then BACKSPRITE = DIBtoSTR(cDibbuffer0)
End With
   
there:
On Error Resume Next
Dim bDib() As Byte, bDib1() As Byte, bDib2() As Byte
''Dim x As Long, y As Long
''Dim lc As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
Dim tSA2 As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibbuffer0.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibbuffer0.BytesPerScanLine()
        .pvData = cDibbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4


    ''Dim nx As Long, ny As Long
    Dim image_x As Single, image_y As Single, temp_image_x As Single, temp_image_y As Single
    Dim x_step As Single, y_step As Single, x_step2 As Single, y_step2 As Single
    Dim screen_x As Long, screen_y As Long, mmx As Long, mmy As Long, mmy1 As Long

    Dim pws As Single, phs As Single
    Dim dest As Long, pw As Single, ph As Single
       pw = cDIBbuffer1.Width
      ph = cDIBbuffer1.Height
      
  
     pws = pw * zoomfactor
    phs = ph * zoomfactor
    
    R = Atn(CSng(myw) / CSng(myh))
    k = -CSng(myw) / (2! * Sin(R))
    
    x_step = Cos(Angle!) * pw
    y_step = Sin(Angle!) * ph

    x_step2 = Cos(Angle! + pidicv2) * pw
    y_step2 = Sin(Angle! + pidicv2) * ph
  image_x = ((pws - b) / 2 - (k * Sin(Angle! - R))) * pw
   image_y = ((phs - b) / 2 + (k * Cos(Angle! - R))) * ph
      image_x = image_x - MyMod(image_x, CSng(dv15))
   image_y = image_y - MyMod(image_y, CSng(dv15))
  pws = pws + 1
  phs = phs + 1
    If Not cDIBbuffer2 Is Nothing Then
                   With tSA2
                   .cbElements = 1
                   .cDims = 2
                   .Bounds(0).lLbound = 0
                   .Bounds(0).cElements = cDIBbuffer2.Height
                   .Bounds(1).lLbound = 0
                   .Bounds(1).cElements = cDIBbuffer2.BytesPerScanLine()
                   .pvData = cDIBbuffer2.DIBSectionBitsPtr
                   End With
                   CopyMemory ByVal VarPtrArray(bDib2()), VarPtr(tSA2), 4
 

                For screen_y = 0 To myh - 1
                 temp_image_x = image_x
                 temp_image_y = image_y
                 For screen_x = 0 To (myw - 1) * 3 Step 3
                
                         mmx = Int(temp_image_x / pws)
                         mmy = Int(temp_image_y / phs)
                
                
     
                
                              If mmx >= 0 And mmx < pw And mmy >= 0 And mmy < ph Then  'new
                                 mmx = mmx * 3
                                                           If bDib(mmx, mmy) <> br Or bDib(mmx + 1, mmy) <> bG Or bDib(mmx + 2, mmy) <> bbb Then
                                 bDib1(screen_x, screen_y) = (bDib(mmx, mmy) * CLng(255 - bDib2(mmx, mmy)) + bDib1(screen_x, screen_y) * CLng(bDib2(mmx, mmy))) \ 255
                                 bDib1(screen_x + 1, screen_y) = (bDib(mmx + 1, mmy) * CLng(255 - bDib2(mmx + 1, mmy)) + bDib1(screen_x + 1, screen_y) * CLng(bDib2(mmx + 1, mmy))) \ 255
                                 bDib1(screen_x + 2, screen_y) = (bDib(mmx + 2, mmy) * CLng(255 - bDib2(mmx + 2, mmy)) + bDib1(screen_x + 2, screen_y) * CLng(bDib2(mmx + 2, mmy))) \ 255
                 End If
                             End If

                     temp_image_x = temp_image_x + x_step
                     temp_image_y = temp_image_y + y_step
                Next screen_x
                 image_x = image_x + x_step2
                 image_y = image_y + y_step2
                Next screen_y
                 CopyMemory ByVal VarPtrArray(bDib2), 0&, 4
    Set cDIBbuffer2 = Nothing
    '*********************************************************
    Else

    For screen_y = 0 To myh - 1
        temp_image_x = image_x
        temp_image_y = image_y
        For screen_x = 0 To (myw - 1) * 3 Step 3
  
                    mmx = Int(temp_image_x / pws)
                    mmy = Int(temp_image_y / phs)
                     If mmx >= 0 And mmx < pw And mmy >= 0 And mmy < ph Then ' new
                        mmx = mmx * 3
                        If bDib(mmx, mmy) <> br Or bDib(mmx + 1, mmy) <> bG Or bDib(mmx + 2, mmy) <> bbb Then
                                      If Alpha = 0 Then
                                      ElseIf Alpha = 100 Then
                                        bDib1(screen_x, screen_y) = bDib(mmx, mmy)
                                      bDib1(screen_x + 1, screen_y) = bDib(mmx + 1, mmy)
                                      bDib1(screen_x + 2, screen_y) = bDib(mmx + 2, mmy)
                                    
                                      Else
                                      
                                      bDib1(screen_x, screen_y) = (bDib(mmx, mmy) * Alpha + bDib1(screen_x, screen_y) * (100 - Alpha)) \ 100
                                      bDib1(screen_x + 1, screen_y) = (bDib(mmx + 1, mmy) * Alpha + bDib1(screen_x + 1, screen_y) * (100 - Alpha)) \ 100
                                      bDib1(screen_x + 2, screen_y) = (bDib(mmx + 2, mmy) * Alpha + bDib1(screen_x + 2, screen_y) * (100 - Alpha)) \ 100
                                      End If
                        Else
                        

                        End If
                    End If
          
            temp_image_x = temp_image_x + x_step
            temp_image_y = temp_image_y + y_step
       Next screen_x
        image_x = image_x + x_step2
        image_y = image_y + y_step2
    Next screen_y
    End If
exitHere:
    
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4



    
    End If
Set cDIBbuffer1 = Nothing

End Function








Public Function RotateDib1(bstack As basetask, cDibbuffer0 As cDIBSection, Optional ByVal Angle! = 0, Optional ByVal zoomfactor As Single = 100, _
   Optional bckColor As Long = -1, Optional BACKx As Long, Optional BACKy As Long)
   Const Pi = 3.14159!
   Dim b As Single
   
   b = CSng(Angle! Mod 90 = 0)
Angle! = -MyMod(Angle!, 360!) * 1.745329E-02!
If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100!
Dim myw As Single, myh As Single, piw As Long, pih As Long, pix As Long, piy As Long
Dim k As Single, R As Single
Const pidicv2 = 1.570795!
'If zoomfactor = 1 And angle! = 0 Then Exit Function
    piw = cDibbuffer0.Width
    pih = cDibbuffer0.Height
 Dim cDIBbuffer1 As Object, cDIBbuffer2 As Object
 Set cDIBbuffer1 = New cDIBSection

Call cDIBbuffer1.create(piw, pih)
cDIBbuffer1.GetDpiDIB cDibbuffer0
cDibbuffer0.needHDC
cDIBbuffer1.LoadPictureBlt cDibbuffer0.HDC1
cDibbuffer0.FreeHDC
  
 
myw = Round((Abs(piw * Cos(Angle!)) + Abs(pih * Sin(Angle!))) * zoomfactor, 0)
myh = Round((Abs(piw * Sin(Angle!)) + Abs(pih * Cos(Angle!))) * zoomfactor, 0)

cDibbuffer0.ClearUp
If cDibbuffer0.create(myw, myh) Then
On Error GoTo there
If bckColor >= 0 Then
cDibbuffer0.Cls bckColor
Else
        With bstack.Owner
        If bstack.toprinter Then
        cDibbuffer0.LoadPictureBlt .hDC, Int(.ScaleX(BACKx, 0, 3)), Int(.ScaleX(BACKy, 0, 3))
        Else
                      cDibbuffer0.LoadPictureBlt .hDC, .ScaleX(BACKx, 1, 3), .ScaleX(BACKy, 1, 3)
           End If
        End With
        End If
there:
On Error Resume Next
Dim bDib() As Byte, bDib1() As Byte, bDib2() As Byte
''Dim x As Long, y As Long
''Dim lc As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
Dim tSA2 As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibbuffer0.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibbuffer0.BytesPerScanLine()
        .pvData = cDibbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4

    Dim image_x As Single, image_y As Single, temp_image_x As Single, temp_image_y As Single
    Dim x_step As Single, y_step As Single, x_step2 As Single, y_step2 As Single
    Dim screen_x As Long, screen_y As Long, mmx As Long, mmy As Long, mmy1 As Long
   Dim sX As Single, sY As Single
    Dim xf As Single, yf As Single
    Dim xf1 As Single, yf1 As Single
    Dim pws As Single, phs As Single

    Dim pw As Long, ph As Long
     pw = piw
    ph = pih
    R = Atn(CSng(myw) / CSng(myh))
    k = -myw / (2! * Sin(R))
  
    Dim pw1 As Long, ph1 As Long
    
     pw1 = pw + 1
    ph1 = ph + 1
  
   
   
         pws = pw1 * zoomfactor
    phs = ph1 * zoomfactor
  image_x = ((pws - zoomfactor - b) / 2 - (k * Sin(Angle! - R))) * pw
   image_y = ((phs - zoomfactor - b) / 2 + (k * Cos(Angle! - R))) * ph
   image_x = image_x - MyMod(image_x, CSng(dv15))
   image_y = image_y - MyMod(image_y, CSng(dv15))
   x_step2 = Cos(Angle! + pidicv2) * pw
    y_step2 = Sin(Angle! + pidicv2) * ph

    x_step = Cos(Angle!) * pw
    y_step = Sin(Angle!) * ph
  pws = pws + 1
  phs = phs + 1
    pw1 = pw - 1
    ph1 = ph - 1
    For screen_y = 0 To myh - 1
        temp_image_x = image_x
        temp_image_y = image_y
        For screen_x = 0 To (myw - 1) * 3 Step 3
                sX = temp_image_x / pws
                sY = temp_image_y / phs
                mmx = Int(sX)
                mmy = Int(sY)

                           
                 
           If mmx >= 0 And mmx < pw1 And mmy >= 0 And mmy < ph1 Then
         xf = Abs((sX - CSng(mmx)))
             xf1 = 1! - xf
                      yf = Abs((sY - CSng(mmy)))
                      yf1 = 1! - yf
           
           
         mmx = mmx * 3
          
              
              
                    
                        
                        bDib1(screen_x, screen_y) = yf1 * (xf1 * bDib(mmx, mmy) + xf * bDib(mmx + 3, mmy)) + yf * (xf1 * bDib(mmx, mmy + 1) + xf * bDib(mmx + 3, mmy + 1))
                        bDib1(screen_x + 1, screen_y) = yf1 * (xf1 * bDib(mmx + 1, mmy) + xf * bDib(mmx + 4, mmy)) + yf * (xf1 * bDib(mmx + 1, mmy + 1) + xf * bDib(mmx + 4, mmy + 1))
                        bDib1(screen_x + 2, screen_y) = yf1 * (xf1 * bDib(mmx + 2, mmy) + xf * bDib(mmx + 5, mmy)) + yf * (xf1 * bDib(mmx + 2, mmy + 1) + xf * bDib(mmx + 5, mmy + 1))
          
                    End If
            temp_image_x = temp_image_x + x_step
            temp_image_y = temp_image_y + y_step
       Next screen_x
       
        image_x = image_x + x_step2
        image_y = image_y + y_step2
    Next screen_y
exitHere:
    
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4



    End If
    
Set cDIBbuffer1 = Nothing

End Function



Sub Conv24(cDibbuffer0 As Object)
 Dim cDIBbuffer1 As Object
 Set cDIBbuffer1 = New cDIBSection
Call cDIBbuffer1.create(cDibbuffer0.Width, cDibbuffer0.Height)
cDIBbuffer1.LoadPictureBlt cDibbuffer0.hDC
Set cDibbuffer0 = cDIBbuffer1
Set cDIBbuffer1 = Nothing
End Sub
Public Function CmpHeight_pixels(s As Single) As Single
CmpHeight_pixels = s * 20# / DYP
End Function
Public Function CmpHeight(s As Single) As Single
CmpHeight = s * 20#
End Function
Public Function FindSpriteByTag(sp As Long) As Long
Dim i As Long
For i = 0 To PobjNum
If val("0" & Form1.dSprite(i).Tag) = sp Then
FindSpriteByTag = i
Exit For
End If
Next i
End Function
Sub RsetRegion(ob As Control)
With ob

Call SetWindowRgn(.hWnd, (0), False)
End With
End Sub
Public Function RotateRegion(hRgn As Long, Angle As Single, ByVal piw As Long, ByVal pih As Long, ByVal Size As Single) As Long
Dim k As Single, R As Single, aa As Single
aa = (CLng(Angle! * 100) Mod 36000) / 100

Angle! = -Angle * 1.74532925199433E-02
   R = Atn(piw / CSng(pih)) + Pi / 2!
    k = piw / Cos(R)
    Dim myw As Long, myh As Long
 myw = Round((Abs(piw * Cos(Angle!)) + Abs(pih * Sin(Angle!))) * Size, 0)
myh = Round((Abs(piw * Sin(Angle!)) + Abs(pih * Cos(Angle!))) * Size, 0)
hRgn = ScaleRegion(hRgn, Size)


    Dim uXF As XFORM
    Dim D2R As Single, rData() As Byte, rSize As Long
    uXF.eM11 = Cos(Angle!)
    uXF.eM12 = Sin(Angle!)
    uXF.eM21 = -Sin(Angle!)
    uXF.eM22 = Cos(Angle!)
k = Abs(k)

uXF.eDx = Round(k * Cos(Angle! - R) / 2! + k / 2!, 0)
uXF.eDy = Round(k * Sin(Angle! - R) / 2! + k / 2!, 0)


    rSize = GetRegionData(hRgn, rSize, ByVal 0&)
    
    ReDim rData(0 To rSize - 1)
    Call GetRegionData(hRgn, rSize, ByVal VarPtr(rData(0)))
    
RotateRegion = ExtCreateRegion(ByVal VarPtr(uXF), rSize, ByVal VarPtr(rData(0)))

DeleteObject hRgn
    
End Function


Public Function ScaleRegion(hRgn As Long, Size As Single) As Long
  Dim uXF As XFORM
    Dim D2R As Single, rData() As Byte, rSize As Long

    uXF.eM11 = Size
    uXF.eM12 = 0
    uXF.eM21 = 0
    uXF.eM22 = Size

    uXF.eDx = 0
    uXF.eDy = 0
    rSize = GetRegionData(hRgn, rSize, ByVal 0&)
    If rSize > 1 Then
    ReDim rData(0 To rSize - 1)
    Call GetRegionData(hRgn, rSize, ByVal VarPtr(rData(0)))
    ScaleRegion = ExtCreateRegion(ByVal VarPtr(uXF), rSize, ByVal VarPtr(rData(0)))
    End If
     DeleteObject hRgn
End Function
Function GetNewSpriteObj(Priority As Long, s$, tr As Long, rr As Long, Optional ByVal SZ As Single = 1, Optional ByVal rot As Single = 0, Optional bb$ = vbNullString) As Long
Dim photo As Object, myRgn As Long, oldobj As Long
Dim photo2 As Object
 oldobj = FindSpriteByTag(Priority)
 If oldobj Then
' this priority...is used
' so change only image
SpriteGetOtherImage oldobj, s$, tr, rr, SZ, rot, bb$
GetNewSpriteObj = oldobj

Exit Function
Else
      Set photo = New cDIBSection
        Set photo2 = New cDIBSection
        If cDib(s$, photo) Then
            If rr >= 0 Then
                If bb$ <> "" Then
                    If cDib(bb$, photo2) Then
                        myRgn = fRegionFromBitmap2(photo2)
                    Else
                        myRgn = fRegionFromBitmap2(photo, tr, CInt(rr))
                    End If
                Else
                    myRgn = fRegionFromBitmap2(photo, tr, CInt(rr))
                End If
                If myRgn = 0 Then
                    myRgn = CreateRectRgn(0, 0, photo.Width, photo.Height)
                End If
            Else
                myRgn = CreateRectRgn(0, 0, photo.Width, photo.Height)
            End If
 ''''''''''''''''If SZ <> 1 Then myRgn = ScaleRegion(myRgn, SZ)
            myRgn = RotateRegion(myRgn, (rot), photo.Width * SZ, photo.Height * SZ, SZ)



 RotateDibNew photo, (rot), 1, tr

addSprite
Load Form1.dSprite(PobjNum)
With Form1.dSprite(PobjNum)
.Height = photo.Height * DYP * SZ
.Width = photo.Width * DXP * SZ
.Picture = photo.Picture(SZ)


players(PobjNum).X = .Width / 2
players(PobjNum).Y = .Height / 2
Call SetWindowRgn(.hWnd, myRgn, 0)

.Tag = Priority
On Error Resume Next
.ZOrder 0
.Font.Name = Form1.DIS.Font.Name
.Font.charset = Form1.DIS.Font.charset
.Font.Size = SZ
.Font.Strikethrough = False
.Font.Underline = False
.Font.Italic = Form1.DIS.Font.Italic
.Font.bold = Form1.DIS.Font.bold
.Font.Name = Form1.DIS.Font.Name
.Font.charset = Form1.DIS.Font.charset
.Font.Size = SZ

End With
'''DeleteObject myRgn  ' from windows...

GetNewSpriteObj = PobjNum
With players(PobjNum)
.MAXXGRAPH = .X * 2
.MAXYGRAPH = .Y * 2
.hRgn = True
End With
SetText Form1.dSprite(PobjNum)
End If

End If
Dim i As Long, k As Integer

For i = Priority To 32
k = FindSpriteByTag(i)
If k <> 0 Then Form1.dSprite(k).ZOrder 0
Next i


End Function
Function CollidePlayers(Priority As Long, Percent As Long) As Long
Dim i As Long, k As Integer, suma As Long
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, it As Long
k = FindSpriteByTag(Priority)
If k = 0 Then Exit Function
With Form1.dSprite(k)
it = val("0" & .Tag)
x1 = .Left + .Width * (100 - Percent) / 200 - players(it).HotSpotX
y1 = .top + .Height * (100 - Percent) / 200 - players(it).HotSpotY
x2 = .Left + .Width * (1 - (100 - Percent) / 200) - players(it).HotSpotX
y2 = .top + .Height * (1 - (100 - Percent) / 200) - players(it).HotSpotY
End With
For i = Priority - 1 To 1 Step -1
k = FindSpriteByTag(i)
If k <> 0 Then
    With Form1.dSprite(k)
        If (x2 < .Left + .Width / 4) Or (x1 >= .Left + .Width * 3 / 4) Or (y2 <= .top + .Height / 4) Or (y1 > .top + .Height * 3 / 4) Then
        Else
        suma = suma + 2 ^ (i - 1)
        End If
    End With
End If
Next i
CollidePlayers = suma
End Function
Function SpriteVisible(Priority As Long) As Boolean
Dim k As Long
    k = FindSpriteByTag(Priority)
    If k = 0 Then Exit Function
    SpriteVisible = Form1.dSprite(k).Visible
End Function
Function CollideArea(Priority As Long, Percent As Long, basestack As basetask, rest$) As Boolean
' nx2 isn't width but absolute line at nx2
' means not inside
Dim nx1 As Long, ny1 As Long, nx2 As Long, ny2 As Long, p As Double
If IsExp(basestack, rest$, p) Then
nx1 = CLng(p): If Not FastSymbol(rest$, ",") Then Exit Function
If IsExp(basestack, rest$, p) Then
ny1 = CLng(p): If Not FastSymbol(rest$, ",") Then Exit Function
If IsExp(basestack, rest$, p) Then
nx2 = CLng(p): If Not FastSymbol(rest$, ",") Then Exit Function
If IsExp(basestack, rest$, p) Then
ny2 = CLng(p)
End If
End If
End If
End If


Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, k As Long
k = FindSpriteByTag(Priority)
If k = 0 Then Exit Function
x1 = Form1.dSprite(k).Left + Form1.dSprite(k).Width * (100 - Percent) / 200
y1 = Form1.dSprite(k).top + Form1.dSprite(k).Height * (100 - Percent) / 200
x2 = x1 + Form1.dSprite(k).Width * (1 - 2 * (100 - Percent) / 200)
y2 = y1 + Form1.dSprite(k).Height * (1 - 2 * (100 - Percent) / 200)
If x2 < nx1 Or x1 >= nx2 Or y2 <= ny1 Or y1 > ny2 Then
CollideArea = False
Else
CollideArea = True
End If
End Function
Function GetNewLayerObj(Priority As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
Dim photo As cDIBSection, myRgn As Long, oldobj As Long

Set photo = New cDIBSection
If photo.create(lWidth / DXP, lHeight / DYP) Then
photo.WhiteBits
addSprite
Load Form1.dSprite(PobjNum)
With Form1.dSprite(PobjNum)
.Height = lHeight
.Width = lWidth
.Picture = photo.Picture(1)
.Picture = LoadPicture("")
' NO REGION
.Tag = Priority
On Error Resume Next
.ZOrder 0
End With
GetNewLayerObj = PobjNum
Dim i As Long, k As Integer
For i = Priority To 32
k = FindSpriteByTag(i)
If k <> 0 Then Form1.dSprite(k).ZOrder 0
Next i
End If
End Function

Function PosSpriteX(aPrior As Long) As Long ' before take from priority the original sprite
'
Dim k As Long
k = FindSpriteByTag(aPrior)
If k < 1 Or k > PobjNum Then Exit Function
PosSpriteX = Form1.dSprite(k).Left
End Function
Function PosSpriteY(aPrior As Long) As Long ' before take from priority the original sprite
Dim k As Long
k = FindSpriteByTag(aPrior)
If k < 1 Or k > PobjNum Then Exit Function
 PosSpriteY = Form1.dSprite(k).top
End Function

Sub PosSprite(aPrior As Long, ByVal X As Long, ByVal Y As Long) ' ' before take from priority the original sprite
Dim k As Long
k = FindSpriteByTag(aPrior)
If k < 1 Or k > PobjNum Then Exit Sub
 

Form1.dSprite(k).move X, Y

End Sub
Sub SrpiteHideShow(ByVal aPrior As Long, ByVal wh As Boolean) ' this is a priority
On Error Resume Next
Dim k As Long
k = FindSpriteByTag(aPrior)
If k < 1 Or k > PobjNum Then Exit Sub
Form1.dSprite(k).Visible = wh
If wh Then
If Form1.Visible Then
MyDoEvents1 Form1.dSprite(k)
End If
End If
End Sub
Sub SpriteControl(ByVal aPrior As Long, ByVal bPrior As Long) ' these are priorities
Dim k As Long, m As Long, i As Long
k = FindSpriteByTag(aPrior)

If k = 0 Then Exit Sub  ' there is no such a player

    m = FindSpriteByTag(bPrior)
        If m = 0 Then Exit Sub
        Form1.dSprite(k).Tag = bPrior
        Form1.dSprite(m).Tag = aPrior
        
    If aPrior < bPrior Then
    For i = aPrior To 32
        k = FindSpriteByTag(i)
        If k <> 0 Then Form1.dSprite(k).ZOrder 0
    Next i
    Else
    For i = bPrior To 32
        m = FindSpriteByTag(i)
        If m <> 0 Then Form1.dSprite(m).ZOrder 0
    Next i
End If
End Sub
Sub SpriteControlOver(ByVal aPrior As Long, ByVal bPrior As Long) ' these are priorities
Dim k As Long, m As Long, i As Long, LL As Long
k = FindSpriteByTag(aPrior)

If k = 0 Then Exit Sub  ' there is no such a player

    LL = FindSpriteByTag(bPrior)
        If LL = 0 Then Exit Sub
        LL = bPrior + 1
     For i = aPrior + 1 To bPrior
        m = FindSpriteByTag(i)
        If m <> 0 Then
            Form1.dSprite(m).ZOrder 0
            bPrior = Form1.dSprite(m).Tag
            Form1.dSprite(m).Tag = aPrior
            aPrior = bPrior
        End If
    Next i
    Form1.dSprite(k).ZOrder 0
    bPrior = Form1.dSprite(k).Tag
    Form1.dSprite(k).Tag = aPrior
    For i = LL To 32
        m = FindSpriteByTag(i)
        If m <> 0 Then
            Form1.dSprite(m).ZOrder 0
        End If
    Next i

End Sub
Sub SpriteControlUnder(ByVal aPrior As Long, ByVal bPrior As Long) ' these are priorities
Dim k As Long, m As Long, i As Long, LL As Long
k = FindSpriteByTag(aPrior)

If k = 0 Then Exit Sub  ' there is no such a player

    LL = FindSpriteByTag(bPrior)
        If LL = 0 Then Exit Sub
    LL = bPrior - 1
     For i = k - 1 To LL Step -1
        m = FindSpriteByTag(i)
        If m <> 0 Then
            Form1.dSprite(m).ZOrder 1
            bPrior = Form1.dSprite(m).Tag
            Form1.dSprite(m).Tag = aPrior
            aPrior = bPrior
        End If
    Next i
    Form1.dSprite(k).ZOrder 1
    bPrior = Form1.dSprite(k).Tag
    Form1.dSprite(k).Tag = aPrior
    For i = LL To 1 Step -1
        m = FindSpriteByTag(i)
        If m <> 0 Then
            Form1.dSprite(m).ZOrder 1
        End If
    Next i
    

End Sub
Private Sub SpriteGetOtherImage(s As Long, b$, tran As Long, rrr As Long, SZ As Single, rot As Single, Optional bb$ = vbNullString) ' before take from priority the original sprite
Dim photo As Object, myRgn As Long
Dim photo2 As Object
If s < 1 Or s > PobjNum Then Exit Sub

      Set photo = New cDIBSection
       Set photo2 = New cDIBSection
           If cDib(b$, photo) Then
 
 If rrr >= 0 Then
 If bb$ <> "" Then
   If cDib(bb$, photo2) Then
 myRgn = fRegionFromBitmap2(photo2)
 Else
 myRgn = fRegionFromBitmap2(photo, tran, CInt(rrr))
 End If

 
 Else
 myRgn = fRegionFromBitmap2(photo, tran, CInt(rrr))
 End If
  If myRgn = 0 Then
 myRgn = CreateRectRgn(2, 2, photo.Width - 2, photo.Height - 2)
 End If
 
 Else

myRgn = CreateRectRgn(2, 2, photo.Width - 2, photo.Height - 2)


 End If



''If SZ <> 1 Then myRgn = ScaleRegion(myRgn, SZ)

myRgn = RotateRegion(myRgn, (rot), photo.Width * SZ, photo.Height * SZ, SZ)

 RotateDibNew photo, (rot), 1, tran
 

 
Dim oldtag As Long


With Form1.dSprite(s)
.Height = photo.Height * DYP * SZ
.Width = photo.Width * DXP * SZ
.Picture = photo.Picture(SZ)
.Left = .Left + players(s).X - .Width / 2
players(s).X = .Width / 2
.top = .top + players(s).Y - .Height / 2
players(s).Y = .Height / 2
Call SetWindowRgn(.hWnd, myRgn, True)
''''''''''''''''''''''''UpdateWindow .hwnd
 ''DeleteObject myRgn

End With
With players(s)

.MAXXGRAPH = .X * 2
.MAXYGRAPH = .Y * 2
End With
SetText Form1.dSprite(s)
End If
End Sub

Sub addSprite()
PobjNum = PobjNum + 1
'
End Sub
Sub ClrSprites()
On Error Resume Next
Dim i As Long, Zero As basket, zerocounter As Counters
If PobjNum > 0 Then
For i = PobjNum To 1 Step -1
        players(i) = Zero
        Prefresh(i) = zerocounter
        PobjNum = i
If Form1.dSprite.count > PobjNum Then Unload Form1.dSprite(PobjNum)
Next i
PobjNum = 0

End If
' pObject

End Sub
Public Function fRegionFromBitmap2(picSource As cDIBSection, Optional lBackColor As Long = &HFFFFFF, Optional range As Integer = 0) As Long
Dim myRgn() As RECT
Dim lReturn   As Long
Dim lRgnTmp   As Long
Dim lSkinRgn  As Long
Dim lStart    As Long
Dim lRow      As Long
Dim lCol      As Long
'............
Dim bDib() As Byte
Dim tSA As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = picSource.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = picSource.BytesPerScanLine()
        .pvData = picSource.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
'.........................
Dim br As Integer, bG As Integer, bbb As Integer, ppBa As Long  ', ba$, copy1 as long
ppBa = VarPtr(lBackColor)
GetMem1 ppBa, bbb
GetMem1 ppBa + 1, bG
GetMem1 ppBa + 2, br

'ba$ = Hex$(lBackColor)
'ba$ = Right$("00000" & ba$, 6)
'BR = val("&h" & Mid$(ba$, 1, 2))
'BG = val("&h" & Mid$(ba$, 3, 2))
'bbb = val("&h" & Mid$(ba$, 5, 2))

'..................................
Dim mmx As Long, mmy As Long, cc As Long

Dim GLHEIGHT, GLWIDTH As Long
    GLHEIGHT = picSource.Height
    GLWIDTH = picSource.Width
    ReDim myRgn(picSource.Height * 4) As RECT
    Dim rectCount As Long, oldrect
    rectCount = -1
  mmy = -1 ''GLHEIGHT
    For lRow = GLHEIGHT - 1 To 0 Step -1
        lCol = 0
        mmx = 0
      mmy = mmy + 1
        Do While lCol < GLWIDTH
            ' Skip all pixels in a row with the same
            ' color as the background color.
            '
            Do While lCol < GLWIDTH
             
            If Abs(bDib(mmx, mmy) - br) > range Or Abs(bDib(mmx + 1, mmy) - bG) > range Or Abs(bDib(mmx + 2, mmy) - bbb) > range Then Exit Do
               lCol = lCol + 1
                mmx = mmx + 3
            Loop

            If lCol < GLWIDTH Then
                '
                ' Get the start and end of the block of pixels in the
                ' row that are not the same color as the background.
                '
                lStart = lCol
               
                Do While lCol < GLWIDTH
                 If Not (Abs(bDib(mmx, mmy) - br) > range Or Abs(bDib(mmx + 1, mmy) - bG) > range Or Abs(bDib(mmx + 2, mmy) - bbb) > range) Then Exit Do

                mmx = mmx + 3
                    lCol = lCol + 1
                   
                Loop
                
                If lCol > GLWIDTH Then lCol = GLWIDTH
                If rectCount + 2 >= UBound(myRgn()) Then
                ReDim Preserve myRgn(UBound(myRgn()) * 2)
                End If
                
               oldrect = rectCount
              rectCount = rectCount + 1
              SetRect myRgn(rectCount + 2&), lStart, lRow, lCol, lRow + 1

             ''lCol = GLWIDTH
               
            End If
        Loop
    Next

    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
   
    fRegionFromBitmap2 = c_CreatePartialRegion(myRgn(), 2&, rectCount + 1&, 0&, picSource.Width)

End Function


Public Function fRegionFromBitmap(picSource As cDIBSection, Optional lBackColor As Long = &HFFFFFF, Optional range As Integer = 0) As Long
Dim lReturn   As Long
Dim lRgnTmp   As Long
Dim lSkinRgn  As Long
Dim lStart    As Long
Dim lRow      As Long
Dim lCol      As Long
'............
Dim bDib() As Byte
Dim tSA As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = picSource.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = picSource.BytesPerScanLine()
        .pvData = picSource.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
'.........................
Dim br As Integer, bG As Integer, bbb As Integer, ppBa As Long, ba$
ppBa = VarPtr(lBackColor)
GetMem1 ppBa, bbb
GetMem1 ppBa + 1, bG
GetMem1 ppBa + 2, br
ba$ = Hex$(lBackColor)
ba$ = Right$("00000" & ba$, 6)
br = val("&h" & Mid$(ba$, 1, 2))
bG = val("&h" & Mid$(ba$, 3, 2))
bbb = val("&h" & Mid$(ba$, 5, 2))

'..................................
Dim mmx As Long, mmy As Long, cc As Long

Dim GLHEIGHT, GLWIDTH As Long
    GLHEIGHT = picSource.Height
    GLWIDTH = picSource.Width
lSkinRgn = CreateRectRgn(0, 0, 0, 0)
  mmy = GLHEIGHT

    For lRow = 0 To GLHEIGHT - 1
        lCol = 0
        mmx = 0
      mmy = mmy - 1
        Do While lCol < GLWIDTH
            ' Skip all pixels in a row with the same
            ' color as the background color.
            '
            Do While lCol < GLWIDTH
            If Abs(bDib(mmx, mmy) - br) > range Or Abs(bDib(mmx + 1, mmy) - bG) > range Or Abs(bDib(mmx + 2, mmy) - bbb) > range Then Exit Do
                lCol = lCol + 1
                mmx = mmx + 3
            Loop

            If lCol < GLWIDTH Then
                '
                ' Get the start and end of the block of pixels in the
                ' row that are not the same color as the background.
                '
                lStart = lCol
                Do While lCol < GLWIDTH
                If Not (Abs(bDib(mmx, mmy) - br) > range Or Abs(bDib(mmx + 1, mmy) - bG) > range Or Abs(bDib(mmx + 2, mmy) - bbb) > range) Then Exit Do

                mmx = mmx + 3
                    lCol = lCol + 1
                Loop
                If lCol > GLWIDTH Then lCol = GLWIDTH
                '
              
                lRgnTmp = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
                lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                Call DeleteObject(lRgnTmp)
            End If
        Loop
    Next

    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
fRegionFromBitmap = lSkinRgn

End Function

Public Function GetFrequency(Oct As Integer, no As Integer)
Dim lngNote As Long
lngNote = ((Oct - 1) * 12 + no) - 37
GetFrequency = 440 * (2 ^ (lngNote / 12))
End Function
Public Function GetNote(Oct As Integer, no As Integer) As Long
GetNote = Oct * 12 + no
End Function
Public Sub PlayTune(ss$)
Dim octave As Integer, i As Long, v$
Dim note As Integer
Dim silence As Boolean
octave = 4
ss$ = ss$ & " "
For i = 1 To Len(ss$) - 1
v$ = Mid$(ss$, i, 2)
note = InStr(Face$, UCase(v$))
If note = 24 Then

If silence Then
Sleep beeperBEAT / 2
Else
silence = True
Sleep beeperBEAT + beeperBEAT / 2
End If
Else
If note = 0 Then note = InStr(Face$, UCase(Left$(v$, 1)) & " ") Else i = i + 1
If note <> 0 Then
' look for number

If Mid$(ss$, i + 1, 1) <> "" Then If InStr("1234567", Mid$(ss$, i + 1, 1)) > 0 Then octave = val(Mid$(ss$, i + 1, 1)): i = i + 1
' no volume control here
silence = False
Beeper GetFrequency(octave, (note + 1) / 2), beeperBEAT
End If
End If
Next i
End Sub
Public Function PlayTuneMIDI(ss$, octave2play As Integer, note2play As Integer, subbeat As Long, volume2play As Long) As Boolean

Dim i As Long, v$, nomore As Boolean, yesvol As Boolean, probe2play As Integer
ss$ = ss$ & " "
note2play = 0
i = 1
If Trim$(ss$) = vbNullString Then note2play = 0: Exit Function
If Asc(ss$) <> 32 Then
v$ = Mid$(ss$, i, 2)
probe2play = InStr(Face$, UCase(v$))
Else
probe2play = 24
End If

If probe2play = 24 Then

    note2play = 24
    i = i + 1
If Mid$(ss$, i, 1) = "@" Then
            i = i + 1
           If InStr("12345", Mid$(ss$, i, 1)) > 0 Then
           subbeat = val(Mid$(ss$, i, 1))
        i = i + 1
      End If
     End If
     If Mid$(ss$, i, 1) = "V" Then
            i = i + 1
            v$ = vbNullString
        Do While InStr("1234567890", Mid$(ss$, i, 1)) > 0 And (Mid$(ss$, i, 1) <> "")
        v$ = v$ & Mid$(ss$, i, 1)
        i = i + 1
        Loop
        volume2play = val("0" & v$)
     End If
    PlayTuneMIDI = True
    GoTo th
Else
If probe2play = 0 Then
probe2play = InStr(Face$, UCase(Left$(v$, 1)) & " ")
Else
i = i + 1
End If


If probe2play <> 0 Then
i = i + 1
' look for number
If Mid$(ss$, i, 1) <> "" Then
      If InStr("1234567", Mid$(ss$, i, 1)) > 0 Then
        octave2play = val(Mid$(ss$, i, 1))
         i = i + 1
        End If
        If Mid$(ss$, i, 1) = "@" Then
            i = i + 1
           If InStr("12345", Mid$(ss$, i, 1)) > 0 Then
           subbeat = val(Mid$(ss$, i, 1))
        i = i + 1
      End If
       
    End If
         If Mid$(ss$, i, 1) = "V" Then
            i = i + 1
            v$ = vbNullString
        Do While InStr("1234567890", Mid$(ss$, i, 1)) > 0 And (Mid$(ss$, i, 1) <> "")
        v$ = v$ & Mid$(ss$, i, 1)
        i = i + 1
        Loop
        volume2play = val("0" & v$)
     End If
End If

' so we have it here
note2play = probe2play

PlayTuneMIDI = True
End If
End If
th:
ss$ = Mid$(ss$, 1, Len(ss$) - 1) ' drop space
If i = 1 Then note2play = 0: PlayTuneMIDI = False: Exit Function

ss$ = Mid$(ss$, i)

End Function
Public Sub sThread(ByVal ThID As Long, ByVal Thinterval As Double, ByVal ThCode As String, ByVal where$)
Dim task As TaskInterface
 Set task = New counter
          Set task.Owner = Form1.DIS
          ' not use holdtime yet
          task.Parameters ThID, Thinterval, ThCode, Pow2minusOne(32), where$ ', holdtime
          TaskMaster.AddTask task, tmHigh

End Sub
Public Sub sThreadInternal(bs As basetask, ByVal ThID As Long, ByVal Thinterval As Double, ByVal ThCode As String, holdtime As Double, threadhere$, Nostretch)
On Error Resume Next
Dim task As TaskInterface, bsdady As basetask
Set bsdady = bs.Parent
' above 20000 the thid
 Set task = New myProcess
          Set task.Owner = bsdady.Owner
          Set task.Process = bs

          Set bsdady.LinkThread(ThID) = bs.Process
          Set bs = Nothing
          task.Parameters ThID, Thinterval, ThCode, holdtime, threadhere$, Nostretch
          TaskMaster.rest
          TaskMaster.AddTask task
           DoEvents
          Set task = Nothing
          
          Set bsdady = Nothing
          TaskMaster.RestEnd

End Sub
Public Function SetTextData( _
        ByVal lFormatId As Long, _
         sText As String _
    ) As Boolean
    ' use strptr and lenb
    Dim hMem As Long, lPtr As Long
    Dim lSize As Long
        lSize = LenB(sText)
    hMem = GlobalAlloc(0, lSize + 2)
If (hMem > 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, ByVal StrPtr(sText), lSize + 1
        GlobalUnlock hMem
       If (OpenClipboard(0) <> 0) Then
    
      SetClipboardData lFormatId, hMem
      CloseClipboard
       End If
          
    End If
    

End Function
Public Function HTML(sText As String, _
   Optional sContextStart As String = "<HTML><BODY>", _
   Optional sContextEnd As String = "</BODY></HTML>") As Byte()
   Dim m_sDescription As String
    m_sDescription = "Version:1.0" + vbCrLf + _
                  "StartHTML:aaaaaaaaaa" + vbCrLf + _
                  "EndHTML:bbbbbbbbbb" + vbCrLf + _
                  "StartFragment:cccccccccc" + vbCrLf + _
                  "EndFragment:dddddddddd" + vbCrLf
    Dim a() As Byte, b() As Byte, c() As Byte
   '' sText = "<FONT FACE=Arial SIZE=1 COLOR=BLUE>" + sText + "</FONT>"
   
    a() = Utf16toUtf8(sContextStart & "<!--StartFragment -->")
    b() = Utf16toUtf8(sText)
    c() = Utf16toUtf8("<!--EndFragment -->" & sContextEnd)
   Dim sData As String, mdata As Long, eData As Long, fData As Long

   
    eData = UBound(a()) - LBound(a()) + 1
   mdata = UBound(b()) - LBound(b()) + 1
   fData = UBound(c()) - LBound(c()) + 1
   m_sDescription = Replace(m_sDescription, "aaaaaaaaaa", Format(Len(m_sDescription), "0000000000"))
   m_sDescription = Replace(m_sDescription, "bbbbbbbbbb", Format(Len(m_sDescription) + eData + mdata + fData, "0000000000"))
   m_sDescription = Replace(m_sDescription, "cccccccccc", Format(Len(m_sDescription) + eData, "0000000000"))
   m_sDescription = Replace(m_sDescription, "dddddddddd", Format(Len(m_sDescription) + eData + mdata, "0000000000"))
  Dim all() As Byte, m() As Byte
  ReDim all(Len(m_sDescription) + eData + mdata + fData)
  
  m() = Utf16toUtf8(m_sDescription)
  CopyMemory all(0), m(0), Len(m_sDescription)
  CopyMemory all(Len(m_sDescription)), a(0), eData
  CopyMemory all(Len(m_sDescription) + eData), b(0), mdata
  CopyMemory all(Len(m_sDescription) + eData + mdata), c(0), fData
  HTML = all()
  
End Function

Public Function SimpleHtmlData(ByVal sText As String)
Dim lFormatId As Long, bb() As Byte
lFormatId = RegisterCF
If lFormatId <> 0 Then
If sText = vbNullString Then Exit Function
bb() = HTML(sText)
If CBool(OpenClipboard(0)) Then
   
      Dim hMemHandle As Long, lpData As Long
      'If IsWine Then
      'hMemHandle = GlobalAlloc(0, UBound(bb()) - LBound(bb()) + 10)
      'Else
      hMemHandle = GlobalAlloc(0, UBound(bb()) - LBound(bb()) + 10)
      'End If
      If CBool(hMemHandle) Then
               
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            
            CopyMemory ByVal lpData, bb(0), UBound(bb()) - LBound(bb())
            GlobalUnlock hMemHandle
            EmptyClipboard
            SetClipboardData lFormatId, hMemHandle
                        
         End If
      
      End If
   
      Call CloseClipboard
   End If



End If
End Function
Function RegisterCF() As Long


   'Register the HTML clipboard format
   If (m_cfHTMLClipFormat = 0) Then
      m_cfHTMLClipFormat = RegisterClipboardFormat("HTML Format")
   End If
   RegisterCF = m_cfHTMLClipFormat
   
End Function
Public Function SetTextDataLong( _
        ByVal lFormatId As Long, _
         dLong As Long _
    ) As Boolean
    ' use strptr and lenb
    Dim hMem As Long, lPtr As Long
    Dim checkme As Long
    Dim lSize As Long
        lSize = 4
    hMem = GlobalAlloc(0, lSize)
If (hMem > 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, dLong, lSize
        CopyMemory checkme, ByVal lPtr, lSize
        GlobalUnlock hMem
       If (OpenClipboard(0) <> 0) Then
       SetClipboardData lFormatId, hMem
      CloseClipboard
       End If
          
    End If
    

End Function
Public Function SetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean
Dim lSize As Long
Dim lPtr As Long
Dim hMem As Long

    lSize = UBound(bData) - LBound(bData) + 1
    hMem = GlobalAlloc(GMEM_DDESHARE + GMEM_MOVEABLE, lSize)
    If (hMem <> 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, bData(LBound(bData)), lSize
        GlobalUnlock hMem
        OpenClipboard Form1.hWnd
        EmptyClipboard
        If (SetClipboardData(lFormatId, hMem) <> 0) Then
          SetBinaryData = True
        End If
       CloseClipboard
    End If
End Function

Public Function GetClipboardMemoryHandle( _
        ByVal lFormatId As Long _
    ) As Long

    
    ' If the format id is there:
    If (IsClipboardFormatAvailable(lFormatId) <> 0) Then
        ' Get the global memory handIsClipboardFormatAvailable(lFormatId)le to the clipboard data:
       
        GetClipboardMemoryHandle = GetClipboardData(lFormatId)
        
    End If
End Function
Private Function GetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean
' Returns a byte array containing binary data on the clipboard for
' format lFormatID:
Dim hMem As Long, lSize As Long, lPtr As Long
    
    ' Ensure the return array is clear:
    Erase bData
    
    hMem = GetClipboardMemoryHandle(lFormatId)
    ' If success:
    If (hMem <> 0) Then
        ' Get the size of this memory block:
        lSize = GlobalSize(hMem)
        ' Get a pointer to the memory:
        lPtr = GlobalLock(hMem)
        If (lSize > 0) Then
            ' Resize the byte array to hold the data:
            ReDim bData(0 To lSize - 2) As Byte
            ' Copy from the pointer into the array:
            CopyMemory bData(0), ByVal lPtr, lSize - 1
        End If
        ' Unlock the memory block:
        GlobalUnlock hMem
        ' Success:
        GetBinaryData = (lSize > 0)
        ' Don't free the memory - it belongs to the clipboard.
    End If
End Function

Public Function GetTextData(ByVal lFormatId As Long) As String
Dim bData() As Byte, sr As String, sr1 As String
sr1 = Clipboard.GetText(1)
If (OpenClipboard(0) <> 0) Then

        
        If (GetBinaryData(lFormatId, bData())) Then
        sr = bData
        If IsWine Then
                sr1 = Left$(sr, Len(sr1))
                GetTextData = Left$(sr1, Len(sr1))
        Else
                    GetTextData = Left$(sr, Len(sr1))
        End If
        End If
End If
CloseClipboard
End Function
Public Function GetImageEmf() As mHandler
Dim hMem As Long, hEmf As Long, bytes As Long
Dim hCOPY As Long
Dim mypic As New cDIBSection
Dim aPic As MemBlock
Dim Carrier As New mHandler
Const CF_DIB = 8
Const CF_ENHMETAFILE = 14
Dim okb As Boolean
If IsClipboardFormatAvailable(CF_ENHMETAFILE) Then
     If (OpenClipboard(0) <> 0) Then
         hEmf = GetClipboardData(CF_ENHMETAFILE)
         If hEmf <> 0 Then
             ' hCOPY = CopyEnhMetaFile(hemf, ByVal 0)
              
               bytes = GetEnhMetaFileBits(hEmf, bytes, ByVal 0)
                If bytes Then
                    Set aPic = New MemBlock
                    aPic.Construct 1, bytes
                    
                    Call GetEnhMetaFileBits(hEmf, bytes, ByVal aPic.GetBytePtr(0))
                    aPic.SubType = 2 ' emf
                End If
               ' DeleteEnhMetaFile hCOPY
        End If
    End If
    CloseClipboard
    End If
    If aPic Is Nothing Then
        mypic.ClearUp
        mypic.create 128, 128
        mypic.WhiteBits
        mypic.GetDpi 96, 96
        mypic.SaveDibToMeMBlock aPic
        aPic.SubType = 1 ' bitmap
    End If
    Set Carrier.objref = aPic
    Carrier.t1 = 2
    Set GetImageEmf = Carrier


End Function
Public Function GetImageDIB() As mHandler
Dim hMem As Long
Dim hDib As Long
Dim mypic As New cDIBSection
Dim aPic As MemBlock
Dim Carrier As New mHandler
Const CF_DIB = 8
Dim okb As Boolean
     If (OpenClipboard(0) <> 0) Then
        
        hMem = GetClipboardData(CF_DIB)
        If hMem <> 0 Then
           hDib = GlobalLock(hMem)
           mypic.ClearUp
           okb = mypic.CreateFromDIB(hDib)
           If Not okb Then
                hDib = GlobalUnlock(hMem)
                CloseClipboard
                If Clipboard.GetFormat(2) Then
                    mypic.CreateFromPicture Clipboard.GetData(2)
                    okb = mypic.Height
                End If
           End If
           If okb Then
                If mypic.dpix = 0 Then mypic.GetDpi 96, 96
                    If mypic.Height > 0 And mypic.hDib <> 0 Then
                         mypic.SaveDibToMeMBlock aPic
                         aPic.SubType = 1 ' bitmap
                    End If
                End If
        If hMem <> 0 Then Call GlobalUnlock(hMem)
    End If
    CloseClipboard
    If aPic Is Nothing Then
        mypic.ClearUp
        mypic.create 128, 128
        mypic.WhiteBits
        mypic.GetDpi 96, 96
        mypic.SaveDibToMeMBlock aPic
        aPic.SubType = 1 ' bitmap
    End If
    Set Carrier.objref = aPic
    Carrier.t1 = 2
    Set GetImageDIB = Carrier
End If

End Function
Public Function GetImage() As String
Dim hMem As Long, hDib As Long
Dim mypic As New cDIBSection
Const CF_DIB = 8
Dim okb As Boolean
    
If (OpenClipboard(0) <> 0) Then
    hMem = GetClipboardData(CF_DIB)
    If hMem <> 0 Then
        hDib = GlobalLock(hMem)
        mypic.ClearUp
        okb = mypic.CreateFromDIB(hDib)
        If Not okb Then
            hDib = GlobalUnlock(hMem)
            CloseClipboard
            If Clipboard.GetFormat(2) Then
                mypic.CreateFromPicture Clipboard.GetData(2)
                okb = mypic.Height
            End If
        End If
        If okb Then
            If mypic.bitsPerPixel <> 24 Then Conv24 mypic
            If mypic.dpix = 0 Then mypic.GetDpi 96, 96
            If mypic.Height > 0 And mypic.hDib <> 0 Then
                GetImage = DIBtoSTR(mypic)
            End If
        End If
        Call GlobalUnlock(hMem)
    End If
    CloseClipboard
End If

End Function

Public Function MsgBoxN(a$, Optional v As Variant = 5, Optional b$) As Long
AskInput = False
If ASKINUSE Then

Exit Function
End If
    AskTitle$ = b$
    Dim resp As Double
       v = v And &HF
     DialogSetupLang DialogLang
    If DialogLang = 1 Then
        If v = vbRetryCancel Then
        AskOk$ = "RETRY"
        ElseIf v = vbYesNo Then
        AskOk$ = "YES"
        AskCancel$ = "*NO"
        ElseIf v = vbOKCancel Then
        AskOk$ = "OK"
        AskCancel$ = "*" + AskCancel$
        Else
        AskOk$ = "*OK"
        AskCancel$ = vbNullString
        End If
        
        AskText$ = a$ + "..?" + vbCrLf
    Else
             If v = vbRetryCancel Then
        AskOk$ = "епамакгьг"
        ElseIf v = vbYesNo Then
        AskOk$ = "маи"
        AskCancel$ = "*ови"
        ElseIf v = vbOKCancel Then
         AskOk$ = "емтанеи"
         AskCancel$ = "*" + AskCancel$
        Else
        AskCancel$ = vbNullString
        AskOk$ = "емтанеи"
        End If
        AskText$ = a$ + "..;" + vbCrLf
    End If

    resp = Form1.NeoASK(Basestack1)
 
    If v = vbYesNo Then
        If resp = 1 Then MsgBoxN = vbYes Else MsgBoxN = vbNo
    ElseIf v = vbOKCancel Then
        If resp = 1 Then MsgBoxN = vbOK Else MsgBoxN = vbCancel
    ElseIf v = vbRetryCancel Then
        If resp = 1 Then MsgBoxN = vbRetry Else MsgBoxN = vbCancel
    Else
    MsgBoxN = 1
    End If
End Function
Public Function InputBoxN(a$, b$, vv$, thisresp As Double) As String
Dim resp As Double
If ASKINUSE Then

Exit Function
End If
     DialogSetupLang DialogLang

    AskText$ = a$
    AskTitle$ = b$
    AskInput = True
    AskStrInput$ = Trim$(vv$)
    

    resp = Form1.NeoASK(Basestack1)
        If resp = 1 Then InputBoxN = Basestack1.soros.PopStr
          AskInput = False
          thisresp = resp
End Function
Public Function ask(a$, Optional retry As Boolean = False) As Double
'If Form3.Visible Then
'If Form3.WindowState = 1 Then
'Form3.Timer1.enabled = False
'Form3.Timer1.Interval = 32760
'Form3.WindowState = 0
If retry Then
    If Form1.Visible Then
    ask = MsgBoxN(a$, vbRetryCancel + vbQuestion + vbSystemModal, MesTitle$)
    Else
    ask = MsgBoxN(a$, vbRetryCancel + vbQuestion + vbSystemModal, MesTitle$)
    End If

Else
    If Form1.Visible Then
    ask = MsgBoxN(a$, vbOKCancel + vbQuestion + vbSystemModal, MesTitle$)
    Else
    ask = MsgBoxN(a$, vbOKCancel + vbQuestion + vbSystemModal, MesTitle$)
    End If
End If
'Form3.WindowState = 1
'Form3.Timer1.enabled = False
'Form3.Timer1.Interval = 100
'Exit Function
'End If
'End If
'ask = MsgBoxN(a$, vbOKCancel + vbQuestion + vbSystemModal, MesTitle$)
End Function
Public Function SpellUnicode(a$)
' use spellunicode to get numbers
' and make a ListenUnicode...with numbers for input text
Dim b$, i As Long
For i = 1 To Len(a$) - 1
b$ = b$ & CStr(AscW(Mid$(a$, i, 1))) & ","
Next i
SpellUnicode = b$ & CStr(AscW(Right$(a$, 1)))
SpellUnicode = b$ & CStr(AscW(Right$(a$, 1)))
End Function
Public Function ListenUnicode(ParamArray aa() As Variant) As String
Dim all$, i As Long
For i = 0 To UBound(aa)
    all$ = all$ & ChrW(aa(i))
Next i
ListenUnicode = all$
End Function
Function Convert2(a$, localeid As Long) As String  ' to feed textboxes
Dim b$, i&
If a$ <> "" Then
For i& = 1 To Len(a$)
b$ = b$ + Left$(StrConv(ChrW$(AscW(Left$(StrConv(Mid$(a$, i, 1) + Chr$(0), 128, localeid), 1))), 64, 1033), 1)

Next i&
Convert2 = b$
End If
End Function
Function Convert3(a$, localeid As Long) As String  ' to feed textboxes
Dim b$, i&
If a$ <> "" Then
If localeid = 0 Then localeid = Clid
For i& = 1 To Len(a$)
b$ = b$ + Left$(StrConv(ChrW$(AscW(Left$(StrConv(Mid$(a$, i, 1) + Chr$(0), 128, 1033), 1))), 64, localeid), 1)

Next i&
Convert3 = b$
End If
End Function
Function Convert2Ansi(a$, localeid As Long) As String
Dim b$, i&
If a$ <> "" Then
For i& = 1 To Len(a$)
    b$ = b$ + Left$(StrConv(ChrW$(AscW(Left$(StrConv(Mid$(a$, i, 1) + Chr$(0), 128, localeid), 1))), 64, LCID_DEF), 1)
Next i&
Convert2Ansi = b$
End If
End Function
Function GetCodePage(Optional localeid As Long = 1032) As Long
  Dim Buffer As String, ret&
   Buffer = String$(100, 0)

        ret = GetLocaleInfoW(localeid, LOCALE_IDEFAULTANSICODEPAGE, StrPtr(Buffer), 10)
If ret > 0 Then
GetCodePage = val(Mid$(Buffer, 1, 41))
End If
End Function
Function GetCharSet(CodePage As Long)
'
 Dim cp As CHARSETINFO
     If TranslateCharsetInfo(ByVal CodePage, cp, TCI_SRCCODEPAGE) Then
        GetCharSet = cp.ciCharset
    End If
End Function
Sub SwapStringVariant(ByRef b As Variant, ByRef a As Variant)
   Static t As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory ByVal VarPtr(t), ByVal VarPtr(a) + 8, 4
   CopyMemory ByVal VarPtr(a) + 8, ByVal VarPtr(b) + 8, 4
   CopyMemory ByVal VarPtr(b) + 8, ByVal VarPtr(t), 4
End Sub
Sub SwapString2Variant(ByRef s$, ByRef a As Variant)
   Static t As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory ByVal VarPtr(t), ByVal VarPtr(a) + 8, 4
   CopyMemory ByVal VarPtr(a) + 8, ByVal VarPtr(s$), 4
   CopyMemory ByVal VarPtr(s$), ByVal VarPtr(t), 4
End Sub
Sub SwapString2VariantPointer(ByRef s$, ByVal a As Long)
   Static t As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory ByVal VarPtr(t), ByVal a + 8, 4
   CopyMemory ByVal a + 8, ByVal VarPtr(s$), 4
   CopyMemory ByVal VarPtr(s$), ByVal VarPtr(t), 4
End Sub
Sub MoveStringToVariant(ByRef s$, ByRef a As Variant)
   Static t As Long ' 4 Longs * 4 bytes each = 16 bytes
   a = vbNullString
   CopyMemory ByVal VarPtr(t), ByVal VarPtr(a) + 8, 4
   CopyMemory ByVal VarPtr(a) + 8, ByVal VarPtr(s$), 4
   CopyMemory ByVal VarPtr(s$), ByVal VarPtr(t), 4
End Sub
Function IsOptional(ByRef v) As Boolean
Dim t(0 To 2) As Long
CopyMemory t(0), ByVal VarPtr(v), 12
IsOptional = t(0) = 10 And t(2) = -2147352572
End Function
Sub OptVariant(ByRef VarOpt)
Dim t(0 To 3) As Long
t(0) = 10 ' VT_ERROR
t(2) = -2147352572
   CopyMemory ByVal VarPtr(VarOpt), t(0), 16
End Sub

Sub NullVariant(ByRef VarNull)
Dim t(0 To 3) As Long
t(0) = 1 ' VT_NULL
   CopyMemory ByVal VarPtr(VarNull), t(0), 16
End Sub
Function myIsNull(VarNull As Variant) As Boolean
Dim p As Integer
GetMem2 VarPtr(VarNull), p
myIsNull = p = 1
End Function

' VarByRef VarPtr(var2(items)), var(i)
Sub VarByRefVariant(ByVal a As Long, ByRef b As Variant)
    Dim t(0 To 3) As Long
  
   t(0) = vbVariant Or &H4000
   t(2) = VarPtr(b)
   CopyMemory ByVal a, t(0), 16
End Sub
Sub VarByRef(ByVal a As Long, ByRef b As Variant)
Dim t(0 To 3) As Long
   CopyMemory t(0), ByVal VarPtr(b), 16
   t(0) = t(0) Or &H4000
   t(2) = VarPtr(b) + 8
   CopyMemory ByVal a, t(0), 16
End Sub
Sub ArrByRef(ByVal b As Long)
Dim t(1) As Long
   CopyMemory t(0), ByVal VarPtr(b), 4
   t(0) = t(0) Or &H4000
   CopyMemory ByVal VarPtr(b), t(0), 4
End Sub
Sub VarByRefDecimal(ByVal a As Long, ByRef b As Variant)
Dim t(0 To 3) As Long
   CopyMemory t(0), ByVal VarPtr(b), 2
   t(0) = t(0) Or &H4000
   t(2) = VarPtr(b)
   CopyMemory ByVal a, t(0), 16
End Sub
Sub VarByRefCleanRef2(ByVal a As Long)
Dim t(0 To 3) As Long
   CopyMemory t(0), ByVal a, 2
   t(0) = t(0) And &HFFFFBFFF
   CopyMemory ByVal a, t(0), 2
End Sub
Sub VarByRefCleanRef(ByRef a As Long)
Dim t(0 To 3) As Long
   CopyMemory t(0), ByVal VarPtr(a), 2
   t(0) = t(0) And &HFFFFBFFF
   CopyMemory ByVal a, t(0), 2
End Sub
Sub VarByRefClean(ByVal a As Long)
Static z As Variant
CopyMemory ByVal a, z, 16
End Sub
Function VariantIsRef(ByVal a As Long) As Boolean
Static z As Integer
   CopyMemory z, ByVal a, 2
   VariantIsRef = (z And &H4000) = &H4000
End Function
Sub SwapVariantRef(ByRef a As Variant, ByVal b As Long)
    Static z As Variant
   Static t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory ByVal VarPtr(t(0)), ByVal b, 16
  ' t(0) = 17 ' t(0) Or &H4000
   
   CopyMemory ByVal VarPtr(a), ByVal VarPtr(t(0)), 16
   CopyMemory ByVal b, ByVal VarPtr(z), 16
End Sub
Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
   Static t(0 To 3) As Long  ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(a), 16
   CopyMemory ByVal VarPtr(a), ByVal VarPtr(b), 16
   CopyMemory ByVal VarPtr(b), t(0), 16
End Sub
Sub SwapVariant2(ByRef a As Variant, ByRef b As mArray, i As Long)
   Static t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(a), 16
   CopyMemory ByVal VarPtr(a), ByVal b.itemPtr(i), 16
   CopyMemory ByVal b.itemPtr(i), t(0), 16
End Sub
Sub SwapVariant3(ByRef a As mArray, k As Long, ByRef b As mArray, i As Long)
   Static t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal a.itemPtr(k), 16
   CopyMemory ByVal a.itemPtr(k), ByVal b.itemPtr(i), 16
   CopyMemory ByVal b.itemPtr(i), t(0), 16
End Sub
Sub EmptyVariantArrayItem(ByRef b As mArray, i As Long)
   Dim a As Variant
   Static t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(a), 16
   CopyMemory ByVal VarPtr(a), ByVal b.itemPtr(i), 16
   CopyMemory ByVal b.itemPtr(i), t(0), 16
End Sub
Private Function c_CreatePartialRegion(rgnRects() As RECT, ByVal lIndex As Long, ByVal uIndex As Long, ByVal leftOffset As Long, ByVal cx As Long, Optional ByVal xFrmPtr As Long) As Long
'' from Lavolpe, a very fast ROUTINE
    ' Creates a region from a Rect() array and optionally stretches the region

    On Error Resume Next
    ' Note: Ideally, contiguous rows vertically of equal height & width should
    ' be combined into one larger row. However, thru trial & error I found
    ' that Windows does this for us and taking the extra time to do it ourselves
    ' is too cumbersome & slows down the results.
    
    ' the first 32 bytes of a region contain the header describing the region.
    ' Well, 32 bytes equates to 2 rectangles (16 bytes each), so I'll
    ' cheat a little & use rectangles to store the header
    With rgnRects(lIndex - 2&) ' bytes 0-15
        .Left = 32                      ' length of region header in bytes
        .top = 1                        ' required cannot be anything else
        .Right = uIndex - lIndex + 1&   ' number of rectangles for the region
        .Bottom = .Right * 16&          ' byte size used by the rectangles;
    End With                            ' ^^ can be zero & Windows will calculate
    
    With rgnRects(lIndex - 1&) ' bytes 16-31 bounding rectangle identification
        .Left = leftOffset                  ' left
        .top = rgnRects(lIndex).top         ' top
        .Right = leftOffset + cx            ' right
        .Bottom = rgnRects(uIndex).Bottom   ' bottom
    End With
    ' call function to create region from our byte (RECT) array
    c_CreatePartialRegion = ExtCreateRegion(ByVal xFrmPtr, (rgnRects(lIndex - 2&).Right + 2&) * 16&, rgnRects(lIndex - 2&))
    If Err Then Err.Clear

End Function

Function FoundLocaleId(a$) As Long
If Convert3(Convert2(a$, 1032), 1032) = a$ Then
    FoundLocaleId = 1032
ElseIf Convert3(Convert2(a$, 1033), 1033) = a$ Then
    FoundLocaleId = 1033
ElseIf Convert3(Convert2(a$, Clid), Clid) = a$ Then
 FoundLocaleId = Clid
End If
End Function
Function FoundSpecificLocaleId(a$, this As Long) As Long
If Convert3(Convert2(a$, this), this) = a$ Then FoundSpecificLocaleId = True
End Function
Function ismine1(ByVal a$) As Boolean  '  START A BLOCK
ismine1 = True
a$ = myUcase(a$, True)
Select Case a$
Case "PART", "LIB", "PROTOTYPE"
Case "леяос", "пяытотупо"
Case Else
ismine1 = False
End Select
End Function
Function ismine2(ByVal a$) As Boolean  ' CAN START A BLOCK OR DO SOMETHING
ismine2 = True
a$ = myUcase(a$, True)
Select Case a$
Case "ABOUT", "AFTER", "BACK", "BACKGROUND", "CLASS", "COLOR", "DECLARE", "DRAWING", "ELSE", "ENUM", "ENUMERATION", "EVENT", "EVERY", "GLOBAL", "FOR", "FKEY", "FUNCTION", "GROUP", "INVENTORY", "LAYER", "LOCAL", "MAIN.TASK", "MODULE", "OPERATOR", "PATH", "PEN", "PROPERTY", "PRINTER", "PRINTING", "REMOVE", "SET", "STACK", "START", "STRUCTURE", "TASK.MAIN", "THEN", "THREAD", "TRY", "WIDTH", "VALUE", "WHILE"
Case "аявг", "аккиыс", "аниа", "аниа(", "апая", "апаяихлгсг", "цецомос", "цемийо", "цемийг", "цемийес", "циа", "дес", "диацяажг", "долг", "ейтупытгс", "ейтупысг", "емы", "епипедо", "хесе", "хесе(", "идиотгтес", "ивмос", "идиотгта", "йахе", "йатастасг", "йкасг", "йкеиди", "йуяио.еяцо", "лета", "мгла", "олада", "ояисе", "павос", "пема", "пеяи", "пеяихыяио", "сумаятгсг", "сыяос", "текестгс", "тлгла", "топийа", "топийг", "топийес", "тоте", "вяыла"
Case "CONST", "стахеяг", "стахеяес", "сведио", "SUPERCLASS", "упеяйкасг", "DO", "REPEAT", "епамекабе", "епамакабе"
Case "->"
Case Else
ismine2 = False
End Select
End Function
Function ismine22(ByVal a$) As Boolean  ' CAN START A BLOCK AFTER AN EXPRESSION, WE CAN PASS STRING BLOCK IN EXPRESSION
ismine22 = True
a$ = myUcase(a$, True)
Select Case a$
Case "FOR", "WHILE", "циа", "емы"
Case Else
ismine22 = False
End Select
End Function
Function ismine33(ByVal a$) As Boolean  '
ismine33 = True
a$ = myUcase(a$, True)
Select Case a$
Case "CASE", "ле"
Case Else
ismine33 = False
End Select
End Function

Function ismine5(ByVal a$) As Boolean  '  make
ismine5 = True
a$ = myUcase(a$, True)
Select Case a$
Case "GLOBAL", "цемийо", "цемийг", "цемийес"
Case Else
ismine5 = False
End Select
End Function

Function ismine3(ByVal a$) As Boolean  ' CAN START A NEW COMMAND, PROBLEM WITH ELSE
ismine3 = True
a$ = myUcase(a$, True)
Select Case a$
Case "ELSE", "THEN"
Case "аккиыс", "тоте"
Case Else
ismine3 = False
End Select
End Function

Function ismine(ByVal a$) As Boolean
ismine = True
a$ = myUcase(a$, True)
Select Case a$
Case "@(", "$(", "~(", "?", "->", "[]"
Case "ABOUT", "ABOUT$", "ABS(", "ADDRESSOF", "ADD.LICENSE$(", "AFTER", "ALWAYS", "AND", "ANGLE", "APPDIR$", "APPEND", "APPEND.DOC", "APPLICATION"
Case "ARRAY", "ARRAY$(", "ARRAY(", "AS", "ASC(", "ASCENDING", "ASK$(", "ASK(", "ASSERT", "ATN("
Case "BACK", "BACKGROUND", "BACKWARD(", "BANK(", "BASE", "BEEP", "BINARY", "BINARY.ADD(", "BINARY.AND(", "BINARY.NEG(", "BINARY.NOT("
Case "BINARY.OR(", "BINARY.ROTATE(", "BINARY.SHIFT(", "BINARY.XOR(", "BITMAPS", "BMP$(", "BOLD"
Case "BOOLEAN", "BORDER", "BREAK", "BROWSER", "BROWSER$", "BUFFER", "BUFFER(", "BYTE", "CALL", "CASE", "CASCADE", "CAT", "CAR("
Case "CDATE(", "CDR(", "CEIL(", "CENTER", "CHANGE", "CHARSET", "CHOOSE.COLOR", "CHOOSE.FONT", "CHOOSE.OBJECT", "CHOOSE.ORGAN"
Case "CHR$(", "CHRCODE$(", "CHRCODE(", "CIRCLE", "CLASS", "CLEAR", "CLIPBOARD", "CLIPBOARD$", "CLIPBOARD.DRAWING", "CLIPBOARD.IMAGE", "CLIPBOARD.IMAGE$"
Case "CLOSE", "CLS", "CODE", "CODEPAGE", "COLLIDE(", "COLOR", "COLOR(", "COLORS"
Case "COLOUR(", "COM", "COMMAND", "COMMAND$", "COMMIT", "COMMON", "COMPARE(", "COMPRESS", "COMPUTER", "COMPUTER$", "CONCURRENT", "CONST", "CONS("
Case "CONTINUE", "CONTROL$", "COPY", "COS(", "CTIME(", "CURRENCY", "CURSOR", "CURVE"
Case "DATA", "DATE", "DATE$(", "DATE(", "DATEFIELD", "DB.PROVIDER", "DB.USER", "DECIMAL", "DECLARE", "DEF", "DELETE"
Case "DESCENDING", "DESKTOP", "DIM", "DIMENSION(", "DIR", "DIR$", "DIV", "DO"
Case "DOC.LEN(", "DOC.PAR(", "DOC.UNIQUE.WORDS(", "DOC.WORDS(", "DOCUMENT", "DOS", "DOUBLE", "DOWN", "DRAW", "DRAWING"
Case "DRAWINGS", "DRIVE$(", "DRIVE.SERIAL(", "DROP", "DRW$(", "DURATION", "EACH("
Case "EDIT", "EDIT.DOC", "ELSE", "ELSE.IF", "EMPTY", "END", "END(", "ENUM", "ENUMERATION", "ENVELOPE$(", "EOF("
Case "ERASE", "ERROR", "ERROR$", "ESCAPE", "EVAL(", "EVAL$(", "EVENT", "EVENTS", "EVERY", "EXCLUSIVE", "EXECUTE", "EXIST(", "EXIST.DIR("
Case "EXIT", "EXPORT", "EXTERN", "FALSE", "FAST", "FIELD", "FIELD$(", "FILE$("
Case "FILE.APP$(", "FILE.NAME$(", "FILE.NAME.ONLY$(", "FILE.PATH$(", "FILE.STAMP(", "FILE.TITLE$(", "FILE.TYPE$(", "FILELEN(", "FILES"
Case "FILL", "FILTER(", "FILTER$(", "FINAL", "FIND", "FKEY", "FLOODFILL", "FLOOR(", "FLUSH", "FOLD(", "FOLD$(", "FONT", "FONTNAME$", "FOR"
Case "FORM", "FORMAT$(", "FORMLABEL", "FORWARD(", "FRAC(", "FRAME", "FREQUENCY(", "FROM", "FUNCTION", "FUNCTION$(", "FUNCTION("
Case "GARBAGE", "GET", "GETOBJECT(", "GLOBAL", "GOSUB", "GOTO", "GRABFRAME$", "GRADIENT", "GREEK", "GROUP", "GROUP(", "GROUP$("
Case "GROUP.COUNT(", "HALT", "HAVE(", "HEIGHT", "HELP", "HEX", "HEX$(", "HIDE", "HIDE$(", "HIGH", "HIFI", "HIGHWORD("
Case "HILOWWORD(", "HIWORD(", "HOLD", "HSL(", "HTML", "HWND", "ICON", "IF", "IF(", "IF$(", "IMAGE", "IMAGE(", "IMAGE.X("
Case "IMAGE.X.PIXELS(", "IMAGE.Y(", "IMAGE.Y.PIXELS(", "IN", "INFINITY", "INKEY$", "INKEY(", "INLINE", "INPUT", "INPUT$("
Case "INSERT", "INSTR(", "INT(", "INTEGER", "INTERFACE", "INTERVAL", "INTERNET", "INTERNET$", "INVENTORY", "IS", "ISLET", "ISNUM", "ISWINE", "ITALIC"
Case "JOYPAD", "JOYPAD(", "JOYPAD.ANALOG.X(", "JOYPAD.ANALOG.Y(", "JOYPAD.DIRECTION(", "JPG$(", "KEEP", "KEY$", "KEYBOARD"
Case "KEYPRESS(", "LAMBDA", "LAMBDA(", "LAMBDA$", "LAMBDA$(", "LAN$", "LANDSCAPE", "LATIN", "LAYER", "LAZY$(", "LCASE$(", "LEFT$(", "LEFTPART$(", "LEGEND", "LEN"
Case "LEN(", "LEN.DISP(", "LET", "LETTER$", "LIB", "LICENSE", "LINE", "LINESPACE", "LINK", "LIST", "LN("
Case "LOAD", "LOAD.DOC", "LOCAL", "LOCALE", "LOCALE$(", "LOCALE(", "LOG(", "LONG", "LOOP"
Case "LOWORD(", "LOWWORD(", "LTRIM$(", "MAIN.TASK", "MAP(", "MARK", "MASTER", "MATCH(", "MAX(", "MAX.DATA$("
Case "MAX.DATA(", "MDB(", "MEDIA", "MEDIA.COUNTER", "MEMBER$(", "MEMBER.TYPE$(", "MEMO", "MEMORY", "MENU"
Case "MENU$(", "MENU.VISIBLE", "MENUITEMS", "MERGE.DOC", "METHOD", "MID$(", "MIN(", "MIN.DATA$(", "MIN.DATA("
Case "MOD", "MODE", "MODULE", "MODULE$", "MODULE(", "MODULES", "MODULE.NAME$", "MONITOR", "MONITORS", "MONITOR.STACK", "MONITOR.STACK.SIZE", "MOTION", "MOTION.W", "MOTION.WX"
Case "MOTION.WY", "MOTION.X", "MOTION.XW", "MOTION.Y", "MOTION.YW", "MOUSE", "MOUSE.ICON", "MOUSE.KEY", "MOUSE.X"
Case "MOUSE.Y", "MOUSEA.X", "MOUSEA.Y", "MOVE", "MOVIE", "MOVIE.COUNTER", "MOVIE.DEVICE$", "MOVIE.ERROR$", "MOVIE.HEIGHT", "MOVIE.STATUS$", "MOVIE.WIDTH"
Case "MOVIES", "MUSIC", "MUSIC.COUNTER", "NAME", "NEW", "NEXT"
Case "NORMAL", "NOT", "NOTHAVE(", "NOTHING", "NOW", "NUMBER", "NULL", "OBJECT", "OFF", "OLE", "ON"
Case "OPEN", "OPEN.FILE", "OPEN.IMAGE", "OPERATOR", "OPTIMIZATION", "OR", "ORDER", "ORDER(", "OSBIT", "OS$", "OUT", "OUTPUT"
Case "OVER", "OVERWRITE", "PAGE", "PARAGRAPH$(", "PARAGRAPH(", "PARAGRAPH.INDEX(", "PARAM(", "PARAM$(", "PARAMETERS$", "PART", "PARENT", "PASSWORD"
Case "PATH", "PATH$(", "PAUSE", "PEN", "PI", "PIECE$(", "PIPE", "PIPENAME$(", "PLATFORM$", "PLAY"
Case "PLAYER", "PLAYSCORE", "POINT", "POINTER", "POINTER(", "POINT(", "POLYGON", "PORTRAIT", "POS", "POS(", "POS.X", "POS.Y", "PRINT"
Case "PRINTER", "PRINTERNAME$", "PRINTING", "PRIVATE", "PROFILER", "PROPERTY", "PROPERTY(", "PROPERTY$(", "PROPERTIES", "PROPERTIES$", "PROTOTYPE", "PSET", "PUBLIC", "PUSH", "PUT", "QUEUE", "QUOTE$("
Case "RANDOM", "RANDOM(", "READ", "READY(", "RECORDS(", "RECURSION.LIMIT", "REFER", "REFRESH", "RELEASE", "REM"
Case "REMOVE", "REPEAT", "REPLACE$(", "REPORT", "REPORTLINES", "RESTART", "RETRIEVE", "RETURN", "REV(", "REVISION"
Case "RIGHT", "RIGHT$(", "RIGHTPART$(", "RINSTR(", "RND", "ROUND(", "ROW", "RTRIM$(", "SAVE", "SAVE.AS", "SAVE.DOC", "SCALE.X"
Case "SCALE.Y", "SCAN", "SCORE", "SCREEN.PIXELS", "SCREEN.X", "SCREEN.Y", "SCRIPT", "SCROLL", "SEARCH"
Case "SEEK", "SEEK(", "SELECT", "SEQUENTIAL", "SET", "SETTINGS", "SGN(", "SHIFT", "SHIFTBACK", "SHORTDIR$("
Case "SHOW", "SHOW$(", "SIN(", "SINGLE", "SINT(", "SIZE", "SIZE.X(", "SIZE.Y(", "SLICE(", "SLOW", "SMOOTH"
Case "SND$(", "SORT", "SORT(", "SOUND", "SOUNDREC", "SOUNDREC.LEVEL", "SOUNDS", "SPEECH", "SPEECH$(", "SPLIT", "SPRITE"
Case "SPRITE$", "SQRT(", "STACK", "STACK(", "STACK$(", "STACK.SIZE", "STACKITEM$(", "STACKITEM(", "STACKTYPE$(", "START", "START(", "STATIC"
Case "STEP", "STEREO", "STOCK", "STOP", "STR$(", "STREAM", "STRING", "STRING$(", "STRREV$(", "STRUCTURE", "SUB", "SUBDIR", "SUM(", "SUPERCLASS"
Case "SWAP", "SWEEP", "SWITCHES", "TAB", "TAB(", "TABLE", "TAN(", "TARGET"
Case "TARGETS", "TASK.MAIN", "TEMPNAME$", "TEMPORARY$", "TEST", "TEST(", "TEXT", "THEN", "THIS"
Case "THREAD", "THREAD.PLAN", "THREADS", "THREADS$", "TICK", "TIME$(", "TIME(", "TIMECOUNT", "TITLE", "TITLE$("
Case "TO", "TODAY", "TONE", "TOP", "TRIM$(", "TRUE", "TRY", "TUNE", "TWIPSX"
Case "TWIPSY", "TYPE", "TYPE$(", "UCASE$(", "UINT(", "UNARY", "UNDER", "UNICODE", "UNION.DATA$(", "UNIQUE", "UNTIL"
Case "UP", "UPDATABLE", "UPDATE", "USE", "USER", "USERS", "USER.NAME$", "USGN("
Case "VAL(", "VAL$(", "VALID(", "VALUE", "VALUE(", "VALUE$", "VARIANT", "VARPTR(", "VERSION", "VIEW", "VOID", "VOLUME"
Case "WAIT", "WCHAR", "WEAK", "WEAK$(", "WHILE", "WHEN", "WIDE", "WIDTH", "WIN", "WINDOW"
Case "WITH", "WITHEVENTS", "WORDS", "WRITABLE(", "WRITE", "WRITER", "X.TWIPS", "XOR", "Y.TWIPS", "адеиас", "аниысг", "апеияо", "апой.ая$(", "апой.де$("
Case "адеиасе", "ахя(", "ай(", "айеяаио.дуадийо(", "айеяаио", "айеяаиос", "акт", "акгхес", "акгхгс", "аккацг", "аккацг$("
Case "аккане", "аккиыс", "аккиыс.ам", "ам", "ам(", "амап(", "амап$(", "ам$(", "амт(", "амтийеилемо", "ама", "амафгтгсг", "амахеыягсг", "амайтгсг", "амакоцио"
Case "амакоцио$", "амакусг.охомгс", "амакусг.у", "амакусг.ь", "амакусг.в", "амакутгс", "амаломг", "амамеысг", "амажояа", "амаье"
Case "амехесе", "амоицла.аявеиоу", "амоицла.еийомас", "амоине", "амтецяаье", "амтицяаье", "амтийеилемо(", "амы", "аниа", "аниа(", "аниа$", "аниа(", "апая", "апаяихлгсг", "апедысе"
Case "апко", "апкос", "апо", "апохгйеусг.ыс", "апой$(", "апойопг", "апойкеистийа", "апок(", "аяца", "аяихло", "аяихлос", "аяихлос.паяацяажоу("
Case "аяис$(", "аяистеяолеяос$(", "аявеиа", "аявеио", "аявеио$(", "аявеиоу.лгйос(", "аявеиоу.сталпа(", "аявг", "аявг(", "атупо", "атупос", "аукос"
Case "аукос$(", "аукоу", "ауноуса", "ауто", "ажаияесг", "ажгсе", "баке", "баке.адеиа$(", "басг"
Case "басг(", "басг.паяовос", "басг.вягстгс", "баье", "бектистопоигсг", "бгла", "богхеиа", "цецомос", "цецомота", "целисе", "целисла"
Case "цемийес", "цемийг", "цемийо", "циа", "цомийо", "цяалла", "цяалла$", "цяаллатосеияа", "цяаллатосеияа$", "цяаллесамажояас", "цяаллг"
Case "цяажг$(", "цяаье", "цягцояа", "цымиа", "дапед(", "деийтг.лояжг", "деийтгс", "деийтгс(", "деийтгс.йол", "деийтгс.ь", "деийтгс.у", "деийтгс.в"
Case "деийтгса.ь", "деийтгса.у", "деийтгса.в", "деине", "дей(", "дейаен", "дейаен$(", "дем", "демевеи(", "дениос", "дени$(", "денилеяос$(", "дес", "дглосио"
Case "диа", "диабасе", "диацяажг", "диадийтуо", "диадийтуо$", "диадовийо", "диайопг", "диайоптес", "диалесоу", "диаяхяысг", "диаяхяысг(", "диаяйеиа", "диастасг("
Case "диастиво", "диажамеиа", "диажамеиа$", "диажамо", "диажуцг", "диайопг", "диейоье", "диепажг", "диеухл(", "диеухумсгапо", "дийтуо$", "диояхысе"
Case "дипка", "дипко", "дипкос", "дойилг", "дойилг(", "долг", "дяолеас", "дуадийг.пеяистяожг(", "дуадийо", "дуадийо(", "дуадийо.айеяаио("
Case "дуадийо.амти(", "дуадийо.амтистяожо(", "дуадийо.апо(", "дуадийо.г(", "дуадийо.йаи(", "дуадийо.окисхгсг(", "дуадийо.ови(", "дуадийо.пяосхесг(", "дуадийо.пяо(", "дуолиса(", "дысе"
Case "еццяажес(", "еццяажо", "еццяажоу.кенеис(", "еццяажоу.лгйос(", "еццяажоу.ломадийес.кенеис(", "еццяажоу.пая(", "еццяаьило(", "ецйуяо(", "еий$("
Case "еийома", "еийома(", "еийома.ь(", "еийома.ь.сглеиа(", "еийома.у(", "еийома.у.сглеиа(", "еийома.в(", "еийома.в.сглеиа(", "еийомес", "еийомидио", "еимая", "еимаи", "еимця"
Case "еисацыцг", "еисацыцг$(", "еисацыцгс", "ейдосг", "ейтекесг", "ейтупысг", "ейтупысгс", "ейтупытгс", "ейтупытгс$", "ейжя(", "ейжя$("
Case "ейжяасг(", "ейжяасг$(", "екецвос", "екецвос.сыяоу", "екецвос.лецехос.сыяоу", "еккгмийа", "емаомола$", "емхесг", "емйол$", "емйол(", "емтасг", "емтокг$"
Case "емы", "емысе", "CONS(", "емысг.сеияас$(", "емысг(", "енацыцг", "енодос", "енытеяийг", "епам$(", "епамакабе", "епамекабе"
Case "епамы", "епекене", "епекене.амтийеилемо", "епекене.цяаллатосеияа", "епекене.ояцамо", "епекене.вяыла", "епицяажг", "епийаияо", "епикене", "епикене.амтийеилемо", "епикене.цяаллатосеияа"
Case "епикене.ояцамо", "епикене.вяыла", "епикоцес", "епикоцес$(", "епикоцес.жамеяес", "епикоцг", "епикоцг$(", "епикоцгс", "епипедо"
Case "епистяожг", "епижамеиа", "еполемо", "еполема(", "етийета.жоялас", "етоило(", "еуяеиа", "еуяесг", "еуяиа", "ежап(", "ежаялоцг.аявеиоу$(", "ежаялоцг.йат$", "ежаялоцг"
Case "евеи(", "еыс", "г", "гл(", "глеяа$(", "глеяа(", "глеяолгмиа", "гво$(", "гвоцяажгсг", "гвоцяажгсгс.епипедо"
Case "гвои", "гвос", "хесе", "хесе(", "хесг", "хесг(", "хесг.ь", "хесг.у", "хесг.в", "хесгдениа(", "идиотгта", "идиотгта(", "идиотгта$("
Case "идиотгтес", "идиотгтес$", "идиытийо", "исвмг", "исвмг$(", "ивмос", "йахаяг", "йахаяо", "йахе", "йахетг", "йахе(", "йаи", "йакесе", "йалпукг"
Case "йаме", "йамомийа", "йат", "йат$", "йатакоцои", "йатакоцос", "йатастасг", "йатастасг.таимиас$", "йатавыягсг", "йаты"
Case "йатылисо(", "йеилемо", "йемг", "йемо", "йемтяо", "йеж$(", "йимгсг", "йимгсг.п", "йимгсг.пу"
Case "йимгсг.пв", "йимгсг.у", "йимгсг.уп", "йимгсг.ь", "йимгсг.ьп", "йимгсг.в", "йимгсг.вп", "йкасг", "йкеиди", "йкеисе", "йкилан.у", "йкилан.ь"
Case "йкилан.в", "йол$", "йоимос", "йоимои", "йомсока", "йяата", "йяатгсе", "йяужо$(", "йяуье", "йуйкийа", "йуйкос"
Case "йукисг", "йуяио", "йуяио.еяцо", "йыд(", "йыдийа", "йыдийосекида", "кабг", "кабг(", "кабг.амакоцийо.у(", "кабг.амакоцийо.ь("
Case "кабг.амакоцийо.в(", "кабг.йатеухумсг(", "кахос", "кахос$", "кахос.таимиас$", "калда", "калда(", "калда$", "калда$(", "катимийа", "кенеис", "киста", "коц("
Case "коцийо", "коцийос", "коцистийо", "коцистийос", "коцос", "коцос$(", "кс$", "кж(", "лайяу", "лайяус", "ле", "лец(", "лец$(", "лецако("
Case "лецако.сеияас$(", "лецако.сеияас(", "лецецомота", "лецехос", "лецехос.сыяоу", "лецехос.у(", "лецехос.в(", "леходос", "лекос$(", "лекоус.тупос$("
Case "лекыдиа", "леяос", "леяос(", "леяос$(", "лес$(", "лета", "летахесг", "летахесг(", "левяи", "лгдемийос", "лгйос", "лгйос(", "лгйос.елж("
Case "лий(", "лий$(", "лийяо(", "лийяо.сеияас$(", "лийяо.сеияас(", "лийяос.йатакоцос$(", "лмглг", "ломадиаио", "ломадиаиос", "ломадийо", "лояжг$(", "лоусийг", "лоусийг.летягтгс", "лпип"
Case "лпяоста(", "маи", "меа", "мео", "меои", "меос", "мгла", "мглата", "мглата$"
Case "нейима", "одгциа", "одгцос$(", "охомг", "охомес", "олака", "ойм$(", "олада", "олада(", "олада$(", "олада.сумоко(", "омола", "омола.аявеиоу$("
Case "омола.аявеиоу.ломо$(", "омола.тлглатос$", "омола.вягстг$", "ояифомтиа", "ояио.амадяолгс", "ояисе", "ояож(", "осо", "оуяа", "ови", "паифеижымг", "паийтгс", "паийтгс(", "паине", "пай(", "пай$(", "памта"
Case "памы", "памылисо(", "паяацяажос$(", "паяацяажос(", "паяал(", "паяал$(", "паяахесг$(", "паяахуяо", "паяалетяои$", "паяе", "паяейаяе$"
Case "паяелбокг", "патглемо(", "павос", "педиа", "педио", "педио$(", "пеф$(", "пема", "пеяи"
Case "пеяи$", "пеяихыяио", "пета", "пи", "пимайас", "пимайас$(", "пимайас(", "пимайес", "писы("
Case "пкациа", "пкаисио", "пкатос", "пкатос.сглеиоу", "пкатос.таимиас", "пкатжояла$", "пкгйтяокоцио", "покуцымо"
Case "пяос", "пяосаялоцгс", "пяосхесе.еццяажо", "пяосхгйг", "пяытотупо", "пяыто(", "пяосыяимо$", "пяовеияо", "пяовеияо$", "пяовеияо.еийома", "пяовеияо.еийома$", "пяовеияо.сведио", "яеула", "яифа("
Case "яоутима", "яоутимас", "яухлисеис", "яыта$(", "яыта(", "саяысе", "сбгсе", "се"
Case "сеияа", "сеияиайос.дисйоу(", "секида", "семаяио", "сгл", "сгл(", "сглади", "сглеио", "сглеио(", "сглеяа", "сйоупидиа", "стахеяг", "стахеяес", "статийг", "статийес"
Case "стг", "стгкг", "стгкг(", "стгм", "сто", "стой", "стовои", "стовос", "стяоцц(", "суццяажеас"
Case "суццяажг", "суцйяиме(", "суцйяоусг(", "суцвымеусе.еццяажо", "сулпиесг", "сулпкгяысг", "сум(", "сумаятгсг", "сумаятгсгс", "сумаятгсг$("
Case "сумаятгсг(", "сумевисе", "сумхгла", "сус", "сусйеуг", "сусйеуг.пяобокгс$", "сустгла", "сувмотгта(", "свд$(", "сведио", "сведиа"
Case "сведио.мглатым", "сыяос", "сыяос(", "сыяос$(", "сыяоутупос$(", "сысе", "сысе.еццяажо", "таимиа", "таимиа.летягтгс", "таимиес"
Case "танг", "танг(", "танимолгсг", "танимолгсг(", "таутисг(", "таутовяомо", "текестг", "текестгс", "текийг", "текийо", "текийос", "текос", "текос(", "тек(", "тий", "титкос.аявеиоу$(", "тилг"
Case "тилг(", "тилг$(", "тилгсыяоу$(", "тилгсыяоу(", "типота", "титкос", "титкос$(", "тлгла", "тлгла(", "тлгла$", "тлглата", "томос"
Case "тон.еж(", "топийа", "топийес", "топийг", "топийо", "топийо$(", "топийо(", "топос$(", "топос.аявеиоу$("
Case "тоте", "тяап(", "тупос", "тупос$(", "тупос.аявеиоу$(", "тупысе", "туваиос", "туваиос(", "тыяа", "у.сглеиа"
Case "упаявеи(", "упаявеи.йатакоцос(", "упеяйкасг", "уплея(", "упо", "упойатакоцос", "упок", "упокоцистг", "упокоцистгс$", "упокоипо"
Case "уполмгла", "упыяа(", "уьгкос", "уьос", "уьос.сглеиоу", "уьос.таимиас", "жайекос$(", "жамеяо$(", "жаядиа", "жеяе"
Case "жеяеписы", "жхимоуса", "жиктяо(", "жиктяо$(", "жомто", "жояла", "жояла$", "жоятос", "жоятысе", "жоятысе.еццяажо"
Case "жяасг", "жымг", "жыто$(", "в.сглеиа", "вая$(", "ваяайтгяес", "ваяане", "ваяйыд$("
Case "ваяйыд(", "вйж(", "вягсг", "вягстг", "вягстгс", "вягстым", "вяомос$(", "вяомос(", "вяыла", "вяыла(", "вяылата", "вяылатисе"
Case "выяисла", "ьеудес", "ьеудгс", "ьгжио", "ь.сглеиа", "ыс"
Case Else
ismine = False
End Select
End Function
Private Function IsNumberQuery(a$, fr As Long, R As Variant, lr As Long, skipdecimals As Boolean) As Boolean
Dim sg As Long, sng As Long, n$, ig$, DE$, sg1 As Long, ex$, rr As Double
' ti kanei to e$
If a$ = vbNullString Then IsNumberQuery = False: Exit Function
sg = 1
sng = fr - 1
    Do While sng < Len(a$)
    sng = sng + 1
    Select Case Mid$(a$, sng, 1)
    Case " ", "+" ', ChrW(160)
    Case "-"
    sg = -sg
    Case Else
    Exit Do
    End Select
    Loop
n$ = Mid$(a$, sng)

If val("0" & Mid$(a$, sng, 1)) = 0 And Left(Mid$(a$, sng, 1), sng) <> "0" And Left(Mid$(a$, sng, 1), sng) <> "." Then
IsNumberQuery = False

Else
'compute ig$
    If Mid$(a$, sng, 1) = "." And Not skipdecimals Then
    ' no long part
    ig$ = "0"
    DE$ = "."

    Else
    Do While sng <= Len(a$)
        
        Select Case Mid$(a$, sng, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(a$, sng, 1)
        Case "."
        If skipdecimals Then IsNumberQuery = False: Exit Function
        DE$ = "."
        Exit Do
        Case Else
        Exit Do
        End Select
       sng = sng + 1
    Loop
    End If
    If Len(DE$) = 0 Then
        If Len(a$) >= sng& Then
            If InStr("EeеЕ", Mid$(a$, sng&, 1)) > 0 Then
            fr = fr + 1
                DE$ = "."
                GoTo cont1234
            End If
        End If
    Else
      sng = sng + 1
cont1234:
        Do While sng <= Len(a$)
       
        Select Case Mid$(a$, sng, 1)
        Case " " ', ChrW(160)
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        DE$ = DE$ & Mid$(a$, sng, 1)
        End If
        Case "E", "e", "е", "Е" ' ************check it
             If ex$ = vbNullString Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
       
        
        Case "+", "-"
        If sg1 And Len(ex$) = 1 Then
         ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        Exit Do
        End If
        Case Else
        Exit Do
        End Select
         sng = sng + 1
        Loop
        If sg1 Then
            If Len(ex$) < 3 Then
                If ex$ = "E" Then
                    ex$ = " "
                ElseIf ex$ = "E-" Or ex$ = "E+" Then
                    ex$ = "  "
                End If
            End If
        End If
    End If
    If ig$ = vbNullString Then
    IsNumberQuery = False
    lr = 1
    Else
    If sg < 0 Then ig$ = "-" & ig$
    Err.Clear
    On Error Resume Next
    n$ = ig$ & DE$ & ex$
    sng = Len(ig$ & DE$ & ex$)
    rr = val(ig$ & DE$ & ex$)
    If Err.Number > 0 Then
         lr = 0
    Else
        R = rr
       lr = sng - fr + 2
       IsNumberQuery = True
    End If
    
    End If
End If
End Function



Static Function ValidNum(a$, Final As Boolean, Optional cutdecimals As Boolean = False, Optional checktype As Long = 0) As Boolean
Dim R As Long
Dim r1 As Long
r1 = 1
           If Not NoUseDec Then
                                If OverideDec Then
                                    a$ = Replace(a$, NowDec$, ".")
                                 End If
                            Else
                                a$ = Replace(a$, QueryDecString, ".")
                            End If
              
Dim v As Double, b$
If Final Then
r1 = IsNumberOnly(a$, r1, v, R, cutdecimals)

r1 = (r1 And Len(a$) <= R) Or (a$ = vbNullString)
If r1 Then
Select Case checktype
Case vbLong
On Error Resume Next
    v = CLng(v)
    If Err.Number > 0 Then Err.Clear: r1 = False

Case vbSingle
On Error Resume Next
     v = CSng(v)
    If Err.Number > 0 Then Err.Clear: r1 = False
Case vbDecimal
On Error Resume Next
    v = CDec(v)
    If Err.Number > 0 Then Err.Clear: r1 = False
Case vbCurrency
On Error Resume Next
    v = CCur(v)
    If Err.Number > 0 Then Err.Clear: r1 = False
End Select


End If
Else
If (a$ = "-") Or a$ = vbNullString Then
r1 = True
Else
 r1 = IsNumberQuery(a$, r1, v, R, cutdecimals)
    If a$ <> "" Then
         If R < 2 Then
                r1 = Not (R <= Len(a$))
                a$ = vbNullString
        Else
                r1 = r1 And Not R <= Len(a$)
                a$ = Mid$(a$, 1, R - 1)
        End If
        If cutdecimals Then
        If InStr(a$, "e") > 0 Or InStr(a$, "E") > 0 Then
        a$ = Replace(a$, "e", ""): a$ = Replace(a$, "E", "")
        r1 = False
        End If
        End If
 End If
 End If
 End If
  If Not NoUseDec Then
                                If OverideDec Then
                                    a$ = Replace(a$, ".", NowDec$)
                                 End If
                            Else
                                a$ = Replace(a$, ".", QueryDecString)
                            End If
ValidNum = r1
End Function

Function ValidNumberOnly(a$, R As Variant, skipdec As Boolean) As Boolean
If VarType(R) = vbString Then
    R = CVar(0)
Else
    R = R - R
End If
ValidNumberOnly = IsNumberOnly(a$, (1), R, (0), skipdec)
End Function
Function ValidNumberOnlyClean(a$, R As Variant, skipdec As Boolean) As Long
On Error Resume Next
If VarType(R) = vbString Then
    R = CVar(0)
Else
    R = R - R
End If
Dim fr As Long, lr As Long
fr = 1
If IsNumberOnly(a$, fr, R, lr, skipdec) Then
ValidNumberOnlyClean = lr
Else
ValidNumberOnlyClean = -1
End If

End Function
Private Function IsNumberOnly(a$, fr As Long, R As Variant, lr As Long, skipdecimals As Boolean) As Boolean
Dim sg As Long, sng As Long, n$, ig$, DE$, sg1 As Long, ex$   ', e$
' ti kanei to e$
If a$ = vbNullString Then IsNumberOnly = False: Exit Function
sg = 1
sng = fr - 1
    Do While sng < Len(a$)
    sng = sng + 1
    Select Case Mid$(a$, sng, 1)
    Case " ", "+"
    Case "-"
    sg = -sg
    Case Else
    Exit Do
    End Select
    Loop
n$ = Mid$(a$, sng)

If val("0" & Mid$(a$, sng, 1)) = 0 And Left(Mid$(a$, sng, 1), sng) <> "0" And Left(Mid$(a$, sng, 1), sng) <> "." Then
IsNumberOnly = False

Else
'compute ig$
    If Mid$(a$, sng, 1) = "." And Not skipdecimals Then
    ' no long part
    ig$ = "0"
    DE$ = "."

    Else
    Do While sng <= Len(a$)
        
        Select Case Mid$(a$, sng, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(a$, sng, 1)
        Case "."
        If skipdecimals Then Exit Do
        DE$ = "."
        Exit Do
        Case Else
        Exit Do
        End Select
       sng = sng + 1
    Loop
    End If
    ' compute decimal part
    If Len(DE$) = 0 Then
    If Not skipdecimals Then
    If Len(a$) > sng& Then
    If InStr("EeеЕ", Mid$(a$, sng&, 1)) > 0 Then
    If InStr("1234567890+-", Mid$(a$, sng& + 1, 1)) > 0 Then
        DE$ = "."
        GoTo cont123
    End If
    End If
    End If
    End If
    Else
      
      sng = sng + 1
cont123:
        Do While sng <= Len(a$)
       
        Select Case Mid$(a$, sng, 1)
        Case " "
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        DE$ = DE$ & Mid$(a$, sng, 1)
        End If
        Case "E", "e", "е", "Е"  ' ************check it
        If skipdecimals Then Exit Do
             If ex$ = vbNullString Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If

        
        
        Case "+", "-"
        If sg1 And Len(ex$) = 1 Then
             ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        Exit Do
        End If
        Case Else
        Exit Do
        End Select
         sng = sng + 1
        Loop
        
        If ex$ = "E" Or ex$ = "E-" Or ex$ = "E+" Then
        sng = sng - Len(ex$)
        End If
    End If
    If ig$ = vbNullString Then
        IsNumberOnly = False
        lr = 1
    Else
        If sg < 0 Then ig$ = "-" & ig$
        If Len(ig$ + DE$) > 13 And LenB(ex$) = 0 Then
            On Error Resume Next
            If Len(DE$) > 0 Then
                Mid$(DE$, 1, 1) = DefaultDec$
                R = CDec(ig$ & DE$)
            Else
                R = CDec(ig$)
            End If
            If Err.Number = 6 Then
                R = CDbl(ig$ & DE$)
            End If
         Else
            R = val(ig$ & DE$ & ex$)
             If Err.Number > 0 Then
             Err.Clear
             IsNumberOnly = False
             End If

            End If
      'A$ = Mid$(A$, sng)
    lr = sng - fr + 1
    IsNumberOnly = True
    End If
End If
End Function
Public Function ScrX() As Long
ScrX = GetSystemMetrics(SM_CXSCREEN) * dv15
End Function
Public Function ScrY() As Long
ScrY = GetSystemMetrics(SM_CYSCREEN) * dv15
End Function

Public Function MyTrimL3Len(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL3Len = 0: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 7
    Case Else
     
   Exit For
  End Select
  Next i
 MyTrimL3Len = (i - p2) \ 2
End Function
Public Function MyTrimL2(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL2 = 1: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 7
    Case Else
     MyTrimL2 = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimL2 = l + 2
End Function

Public Function MyTrimR(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimR = 1: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p4 To p2 Step -2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     MyTrimR = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimR = l + 2
End Function


Public Function MyTrimL2NoTab(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL2NoTab = 0: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     MyTrimL2NoTab = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimL2NoTab = 0
End Function

Public Function MyTrimRfrom(s$, st As Long, ByVal en As Long) As Long
Dim i&
Dim p2 As Long, p1 As Integer, p4 As Long
  If st > Len(s$) Then MyTrimRfrom = en: Exit Function
  If en > Len(s$) Then MyTrimRfrom = en: Exit Function
  If en <= st Then MyTrimRfrom = en: Exit Function
  If st < 1 Then MyTrimRfrom = en: Exit Function
  p2 = StrPtr(s) + (st - 1) * 2: en = en - 1
  p4 = p2 + (en - st) * 2
  For i = p4 To p2 Step -2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 9
    Case Else
     ' MyTrimRfrom = en + 1
     MyTrimRfrom = (i - p2) \ 2 + st + 1
   Exit Function
  End Select
  Next i
 MyTrimRfrom = st
End Function
Public Function MyTrimCR(s$) As String
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long, p22 As Long
l = Len(s): If l = 0 Then Exit Function

  p2 = StrPtr(s): l = l - 1
  p22 = p2
  p4 = p2 + l * 2
  For i = p4 To p2 Step -2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 10, 13
    Case Else
     Exit For
  End Select
  Next i
  p4 = i
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 10, 13
    Case Else
     
   Exit For
  End Select
  Next i
  p2 = i
  If p2 > p4 Then MyTrimCR = vbNullString Else MyTrimCR = Mid$(s$, (p2 - p22) \ 2 + 1, (p4 - p2) \ 2 + 1)
 
End Function

Public Function MyTrim(s$) As String
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long, p22 As Long
l = Len(s): If l = 0 Then Exit Function

  p2 = StrPtr(s): l = l - 1
  p22 = p2
  p4 = p2 + l * 2
  For i = p4 To p2 Step -2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     Exit For
  End Select
  Next i
  p4 = i
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     
   Exit For
  End Select
  Next i
  p2 = i
  If p2 > p4 Then MyTrim = vbNullString Else MyTrim = Mid$(s$, (p2 - p22) \ 2 + 1, (p4 - p2) \ 2 + 1)
 
End Function
Public Function MyTrimLW(s$) As String
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long, p22 As Long
l = Len(s): If l = 0 Then Exit Function

  p2 = StrPtr(s): l = l - 1
  p22 = p2
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     
   Exit For
  End Select
  Next i
  p2 = i
  If p2 > p4 Then MyTrimLW = vbNullString Else MyTrimLW = Mid$(s$, (p2 - p22) \ 2 + 1, (p4 - p2) \ 2 + 1)
 
End Function
Public Function MyTrimRW(s$) As String
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long, p22 As Long
l = Len(s): If l = 0 Then Exit Function

  p2 = StrPtr(s): l = l - 1
  p22 = p2
  p4 = p2 + l * 2
  For i = p4 To p2 Step -2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     Exit For
  End Select
  Next i
  p4 = i
   If p2 > p4 Then MyTrimRW = vbNullString Else MyTrimRW = Mid$(s$, (p2 - p22) \ 2 + 1, (p4 - p2) \ 2 + 1)
 
End Function

Public Function MyTrimRB(s$) As String
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long, p22 As Long
l = LenB(s): If l = 0 Then Exit Function

  p2 = StrPtr(s): l = l - 1
  p22 = p2
  p4 = p2 + l
  For i = p4 To p2 Step -1
  GetMem1 i, p1
  Select Case p1
    Case 32
    Case Else
   Exit For
  End Select
  Next i
  p4 = i
  If p2 > p4 Then MyTrimRB = vbNullString Else MyTrimRB = MidB$(s$, (p2 - p22) + 1, (p4 - p2) + 1)
 
End Function
Public Function MyTrimLB(s$) As String
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long, p22 As Long
l = LenB(s): If l = 0 Then Exit Function

  p2 = StrPtr(s): l = l - 1
  p22 = p2
  p4 = p2 + l
  For i = p2 To p4 Step 1
  GetMem1 i, p1
  Select Case p1
    Case 32
    Case Else
  
   Exit For
  End Select
    Next i
    p2 = i
  If p2 > p4 Then MyTrimLB = vbNullString Else MyTrimLB = MidB$(s$, (p2 - p22) + 1, (p4 - p2) + 1)
 
End Function
Public Function MyTrimB(s$) As String
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long, p22 As Long
l = LenB(s): If l = 0 Then Exit Function

  p2 = StrPtr(s): l = l - 1
  p22 = p2
  p4 = p2 + l
  For i = p4 To p2 Step -1
  GetMem1 i, p1
  Select Case p1
    Case 32
    Case Else
  
   Exit For
  End Select
  Next i
  p4 = i
  For i = p2 To p4 Step 1
  GetMem1 i, p1
  Select Case p1
    Case 32
    Case Else

   Exit For
  End Select
  Next i
  p2 = i
  If p2 > p4 Then MyTrimB = vbNullString Else MyTrimB = MidB$(s$, (p2 - p22) + 1, (p4 - p2) + 1)
 
End Function
Function IsLabelAnew(where$, a$, R$, Lang As Long) As Long
' for left side...no &

Dim rr&, one As Boolean, c$, gr As Boolean
R$ = vbNullString
' NEW FOR REV 156  - WE WANT TO RUN WITH GREEK COMMANDS IN ANY COMPUTER
Dim i&, l As Long, p3 As Integer
Dim p2 As Long, p1 As Integer, p4 As Long
l = Len(a$): If l = 0 Then IsLabelAnew = 0: Lang = 1: Exit Function

p2 = StrPtr(a$): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 13
    
    If i < p4 Then
    GetMem2 i + 2, p3
    If p3 = 10 Then
    IsLabelAnew = 1234
    If i + 6 > p4 Then
    a$ = vbNullString
    Else
    i = i + 4
    Do While i < p4

    GetMem2 i, p1
    If p1 = 32 Or p1 = 160 Then
    i = i + 2
    Else
    GetMem2 i + 2, p3
    If p1 <> 13 And p3 <> 10 Then Exit Do
    i = i + 4
    End If
    Loop
    a$ = Mid$(a$, (i + 2 - p2) \ 2)
    End If
    Else
    If i > p2 Then a$ = Mid$(a$, (i - 2 - p2) \ 2)
    End If
    Else
    If i > p2 Then a$ = Mid$(a$, (i - 2 - p2) \ 2)
    End If
    
    Lang = 1
    Exit Function
    Case 32, 160, 9
    Case Else

   Exit For
  End Select
  Next i
    If i > p4 Then a$ = vbNullString: IsLabelAnew = 0: Exit Function
  For i = i To p4 Step 2
  GetMem2 i, p1
  If p1 < 256 Then
  Select Case p1
        Case 64  '"@"
            If i < p4 And R$ <> "" Then
                GetMem2 i + 2, p1
                where$ = R$
                R$ = vbNullString
            Else
              IsLabelAnew = 0: a$ = Mid$(a$, (i - p2) \ 2): Exit Function
            End If
        Case 63 '"?"
        If R$ = vbNullString Then
            R$ = "?"
            i = i + 4
        Else
            i = i + 2
        End If
        a$ = Mid$(a$, (i - p2) \ 2)
        IsLabelAnew = 1
        Lang = -1
              
        Exit Function

        Case 46 '"."
            If one Then
                Exit For
            ElseIf R$ <> "" And i < p4 Then
                GetMem2 i + 2, p1
                If ChrW(p1) = "." Or ChrW(p1) = " " Then
                If ChrW(p1) = "." And i + 2 < p4 Then
                    GetMem2 i + 4, p1
                    If ChrW(p1) = " " Then i = i + 4: Exit For
                Else
                    i = i + 2
                   Exit For
                End If
            End If
                GetMem2 i, p1
                R$ = R$ & ChrW(p1)
                rr& = 1
            End If
      Case 38 ' "&"
            If R$ = vbNullString Then
            rr& = 2
            'a$ = Mid$(a$, 2)
            End If
            Exit For
    Case 92, 94, 123 To 126, 160 '"\","^", "{" To "~"
          Exit For
        
        Case 48 To 57, 95 '"0" To "9", "_"
              If one Then

            Exit For
            ElseIf R$ <> "" Then
            R$ = R$ & ChrW(p1)
            '' A$ = Mid$(A$, 2)
            rr& = 1 'is an identifier or floating point variable
            Else
            Exit For
            End If
        Case Is < 0, Is > 64 ' >=A and negative
            If one Then
            Exit For
            Else
            R$ = R$ & ChrW(p1)
            rr& = 1 'is an identifier or floating point variable
            End If
        Case 36 ' "$"
            If one Then Exit For
            If R$ <> "" Then
            one = True
            rr& = 3 ' is string variable
            R$ = R$ & ChrW(p1)
            Else
            Exit For
            End If
        Case 37 ' "%"
            If one Then Exit For
            If R$ <> "" Then
            one = True
            rr& = 4 ' is long variable
            R$ = R$ & ChrW(p1)
            Else
            Exit For
            End If
            
        Case 40 ' "("
            If R$ <> "" Then
            If i + 4 <= p4 Then
                GetMem2 i + 2, p1
                GetMem2 i + 2, p3
                If ChrW(p1) + ChrW(p3) = ")@" Then
                    R$ = R$ & "()."
                    i = i + 4
                Else
                    GoTo i1233
                End If
                            Else
i1233:
                                       Select Case rr&
                                       Case 1
                                       rr& = 5 ' float array or function
                                       Case 3
                                       rr& = 6 'string array or function
                                       Case 4
                                       rr& = 7 ' long array
                                       Case Else
                                       Exit For
                                       End Select
                     GetMem2 i, p1
                                        R$ = R$ & ChrW(p1)
                                        i = i + 2
                                      ' A$ = Mid$(A$, 2)
                                   Exit For
                            
                          End If
               Else
                        Exit For
            
            End If
        Case Else
        Exit For
  End Select

        Else
         If one Then
              Exit For
              Else
              gr = True
              R$ = R$ & ChrW(p1)
              rr& = 1 'is an identifier or floating point variable
              End If
    End If


    Next i
  If i > p4 Then a$ = vbNullString Else If (i + 2 - p2) \ 2 > 1 Then a$ = Mid$(a$, (i + 2 - p2) \ 2)
       R$ = myUcase(R$, gr)
       Lang = 1 + CLng(gr)

    IsLabelAnew = rr&


End Function
Public Function NLtrim$(a$)
If Len(a$) > 0 Then NLtrim$ = Mid$(a$, MyTrimL(a$))
End Function
Public Function NLTrim2$(a$)
If Len(a$) > 0 Then NLTrim2$ = Mid$(a$, MyTrimL2(a$))
End Function
Public Function StringId(aHash As idHash, bHash As idHash, Optional ahashbackup As idHash, Optional bhashbackup As idHash) As Boolean
Dim myid(), i As Long
Dim myfun()
myid() = Array("ABOUT$", 1, "пеяи$", 1, "CONTROL$", 2, "THREADS$", 3, "мглата$", 33, "LAN$", 4, "дийтуо$", 4, "GRABFRAME$", 5, "паяейаяе$", 5 _
, "емаомола$", 6, "TEMPNAME$", 6, "TEMPORARY$", 7, "пяосыяимо$", 7, "USER.NAME$", 8, "омола.вягстг$", 8 _
, "COMPUTER$", 9, "упокоцистгс$", 9, "CLIPBOARD$", 10, "пяовеияо$", 10, "CLIPBOARD.IMAGE$", 11, "пяовеияо.еийома$", 11 _
, "паяалетяои$", 12, "PARAMETERS$", 12, "OS$", 13, "кс$", 13, "емтокг$", 14, "COMMAND$", 14, "кахос$", 15, "ERROR$", 34, "MODULE$", 16, "тлгла$", 16 _
, "PRINTERNAME$", 17, "ейтупытгс$", 17, "PROPERTIES$", 18, "идиотгтес$", 18, "MOVIE.STATUS$", 19, "йатастасг.таимиас$", 19 _
, "MOVIE.DEVICE$", 20, "сусйеуг.пяобокгс$", 20, "MOVIE.ERROR$", 21, "кахос.таимиас$", 21, "PLATFORM$", 22, "пкатжояла$", 22 _
, "FONTNAME$", 23, "цяаллатосеияа$", 23, "BROWSER$", 24, "амакоцио$", 24, "SPRITE$", 25, "диажамеиа$", 25 _
, "APPDIR$", 26, "ежаялоцг.йат$", 26, "DIR$", 27, "йат$", 27, "KEY$", 28, "йол$", 28, "INKEY$", 29, "емйол$", 29, "LETTER$", 30, "цяалла$", 30, "LAMBDA$", 31, "калда$", 35, "жояла$", 32 _
, "омола.тлглатос$", 36, "MODULE.NAME$", 36, "INTERNET$", 37, "диадийтуо$", 37)
If Not ahashbackup Is Nothing Then
    For i = 0 To UBound(myid()) Step 2
        ahashbackup.ItemCreator CStr(myid(i)), CLng(myid(i + 1))
    Next i
End If
For i = 0 To UBound(myid()) Step 2
    aHash.ItemCreator CStr(myid(i)), CLng(myid(i + 1))
Next i
myfun() = Array("FORMAT$(", 1, "лояжг$(", 1, "EVAL$(", 2, "ейжя$(", 2, "ейжяасг$(", 2, "STACKTYPE$(", 3, "сыяоутупос$(", 3 _
, "STACKITEM$(", 4, "тилгсыяоу$(", 4, "исвмг$(", 5, "WEAK$(", 5, "коцос$(", 6, "SPEECH$(", 6, "ASK$(", 7, "яыта$(", 7 _
, "LOCALE$(", 8, "топийо$(", 8, "SHORTDIR$(", 9, "лийяос.йатакоцос$(", 9, "FILTER$(", 10, "жиктяо$(", 10, "коцос$(", 11, "SPEECH$(", 11 _
, "аявеио$(", 12, "FILE$(", 12, "PARAM$(", 13, "паяал$(", 13, "LAZY$(", 14, "ойм$(", 14, "INPUT$(", 15, "еисацыцг$(", 15 _
, "MEMBER.TYPE$(", 16, "лекоус.тупос$(", 16, "MEMBER$(", 17, "лекос$(", 17, "PIPENAME$(", 18, "аукос$(", 18, "FILE.TYPE$(", 20, "тупос.аявеиоу$(", 20, "FILE.NAME.ONLY$(", 21, "омола.аявеиоу.ломо$(", 21, "FILE.NAME$(", 22, "омола.аявеиоу$(", 22 _
, "FILE.PATH$(", 23, "топос.аявеиоу$(", 23, "одгцос$(", 24, "DRIVE$(", 19, "титкос.аявеиоу$(", 25, "FILE.TITLE$(", 25 _
, "ежаялоцг.аявеиоу$(", 26, "FILE.APP$(", 26, "HIDE$(", 27, "йяужо$(", 27, "LEFTPART$(", 28, "аяистеяолеяос$(", 28 _
, "RIGHTPART$(", 29, "денилеяос$(", 29, "ARRAY$(", 30, "пимайас$(", 30, "TYPE$(", 31, "тупос$(", 31, "PARAGRAPH$(", 32, "паяацяажос$(", 32 _
, "UNION.DATA$(", 33, "емысг.сеияас$(", 33, "MAX.DATA$(", 34, "лецако.сеияас$(", 34, "MIN.DATA$(", 35, "лийяо.сеияас$(", 35 _
, "FUNCTION$(", 36, "сумаятгсг$(", 36, "HEX$(", 37, "дейаен$(", 37, "SHOW$(", 38, "жамеяо$(", 38, "MENU$(", 39, "епикоцг$(", 39, "епикоцес$(", 39 _
, "REPLACE$(", 40, "аккацг$(", 40, "PATH$(", 41, "топос$(", 41, "UCASE$(", 42, "йеж$(", 42, "LCASE$(", 43, "пеф$(", 43, "STRING$(", 44, "епам$(", 44, "MID$(", 45, "лес$(", 45 _
, "LEFT$(", 46, "аяис$(", 46, "RIGHT$(", 47, "дени$(", 47, "SND$(", 48, "гво$(", 48, "BMP$(", 49, "еий$(", 49, "JPG$(", 50, "жыто$(", 50 _
, "TRIM$(", 51, "апой$(", 51, "QUOTE$(", 52, "паяахесг$(", 52, "сыяос$(", 53, "STACK$(", 53, "ADD.LICENSE$(", 54, "баке.адеиа$(", 54 _
, "ENVELOPE$(", 55, "жайекос$(", 55, "FIELD$(", 56, "педио$(", 56, "DRW$(", 57, "свд$(", 57, "TIME$(", 58, "вяомос$(", 58, "DATE$(", 59, "глеяа$(", 59 _
, "STR$(", 60, "цяажг$(", 60, "CHRCODE$(", 61, "ваяйыд$(", 61, "CHR$(", 62, "вая$(", 62, "GROUP$(", 63, "олада$(", 63, "PROPERTY$(", 64, "идиотгта$(", 64, "TITLE$(", 65, "титкос$(", 65, "IF$(", 66, "ам$(", 66, "леяос$(", 67, "PIECE$(", 67, "STRREV$(", 68, "амап$(", 68 _
, "RTRIM$(", 69, "апой.де$(", 69, "LTRIM$(", 70, "апой.ая$(", 70, "STACKITEM(", 1, "тилгсыяоу(", 1, "ARRAY(", 2, "пимайас(", 2, "CONS(", 3, "емысг(", 3, "CAR(", 4, "пяыто(", 4, "CDR(", 5, "еполема(", 5, "VAL(", 6, "тилг(", 6, "аниа(", 6, "EVAL(", 7 _
, "ейжя(", 7, "ейжяасг(", 7)


If Not bhashbackup Is Nothing Then
For i = 0 To UBound(myfun()) Step 2
    bhashbackup.ItemCreator CStr(myfun(i)), CLng(myfun(i + 1))
Next i
End If
For i = 0 To UBound(myfun()) Step 2
    bHash.ItemCreator CStr(myfun(i)), CLng(myfun(i + 1))
Next i
StringId = True

End Function
Public Function NumberId(aHash As idHash, bHash As idHash, Optional ahashbackup As idHash, Optional bhashbackup As idHash) As Boolean
Dim myid(), i As Long
Dim myfun()
myid() = Array("THIS", 1, "ауто", 1, "RND", 2, "туваиос", 2, "PEN", 3, "пема", 3, "HWND", 4, "паяахуяо", 4, "LOCALE", 5, "топийо", 5, "CODEPAGE", 6, "йыдийосекида", 6 _
, "SPEECH", 7, "коцос", 7, "ERROR", 8, "кахос", 8, "SCREEN.Y", 9, "амакусг.у", 9, "амакусг.ь", 9, "SCREEN.X", 10, "амакусг.в", 10, "TWIPSY", 11, "уьос.сглеиоу", 11 _
, "TWIPSX", 12, "пкатос.сглеиоу", 12, "REPORTLINES", 13, "цяаллесамажояас", 13, "LINESPACE", 14, "диастиво", 14, "MODE", 15, "тупос", 15 _
, "MEMORY", 16, "лмглг", 16, "CHARSET", 17, "ваяайтгяес", 17, "ITALIC", 18, "пкациа", 18, "BOLD", 19, "жаядиа", 19, "COLORS", 20, "вяылата", 20 _
, "ауноуса", 21, "ASCENDING", 21, "жхимоуса", 22, "DESCENDING", 22, "BOOLEAN", 23, "коцийос", 23, "BYTE", 24, "ьгжио", 24 _
, "INTEGER", 25, "айеяаиос", 25, "LONG", 26, "лайяус", 26, "CURRENCY", 27, "коцистийос", 27, "SINGLE", 28, "апкос", 28, "DOUBLE", 29, "дипкос", 29 _
, "DATEFIELD", 30, "глеяолгмиа", 30, "BINARY", 31, "дуадийо", 31, "TEXT", 32, "йеилемо", 32, "OLE", 33, "MEMO", 34, "уполмгла", 34, "REVISION", 35, "амахеыягсг", 35, "BROWSER", 36, "амакоцио", 36, "VERSION", 37, "ейдосг", 37, "MOTION.X", 38, "йимгсг.в", 38, "MOTION.Y", 39, "йимгсг.у", 39, "йимгсг.ь", 39, "MOTION.XW", 40, "йимгсг.вп", 40, "MOTION.WX", 40, "йимгсг.пв", 40, "MOTION.YW", 41, "йимгсг.уп", 41, "йимгсг.ьп", 41, "MOTION.WY", 41, "йимгсг.пу", 41, "йимгсг.пь", 41 _
, "FIELD", 42, "педио", 42, "MOUSE.KEY", 43, "деийтгс.йол", 43, "MOUSE", 44, "деийтгс", 44, "MOUSE.X", 45, "деийтгс.в", 45 _
, "MOUSE.Y", 46, "деийтгс.у", 46, "деийтгс.ь", 46, "MOUSEA.X", 47, "деийтгса.в", 47, "MOUSEA.Y", 48, "деийтгса.у", 48, "деийтгса.ь", 48, "TRUE", 49, "акгхес", 49, "акгхгс", 49 _
, "FALSE", 50, "ьеудес", 50, "ьеудгс", 50, "STACK.SIZE", 51, "лецехос.сыяоу", 51, "ISNUM", 52, "еимая", 52, "PI", 53, "пи", 53 _
, "NOT", 54, "ови", 54, "дем", 54, "ISLET", 55, "еимця", 55, "WIDTH", 56, "пкатос", 56, "POINT", 57, "сглеио", 57, "POS.X", 58, "хесг.в", 58, "POS.Y", 59, "хесг.ь", 59, "хесг.у", 59 _
, "SCALE.X", 60, "йкилан.в", 60, "в.сглеиа", 60, "X.TWIPS", 60, "SCALE.Y", 61, "йкилан.ь", 61, "йкилан.у", 61, "ь.сглеиа", 61, "у.сглеиа", 61, "Y.TWIPS", 61, "EMPTY", 62, "йемо", 62 _
, "MOVIE.COUNTER", 63, "MEDIA.COUNTER", 63, "MUSIC.COUNTER", 63, "таимиа.летягтгс", 63, "лоусийг.летягтгс", 63 _
, "PLAYSCORE", 64, "паифеижымг", 64, "MOVIE", 65, "MEDIA", 65, "MUSIC", 65, "таимиа", 65, "лоусийг", 65, "DURATION", 66, "диаяйеиа", 66 _
, "VOLUME", 67, "емтасг", 67, "TAB", 68, "стгкг", 68, "HEIGHT", 69, "уьос", 69, "POS", 70, "хесг", 70, "ROW", 71, "цяаллг", 71, "TIMECOUNT", 72, "жоятос", 72 _
, "TICK", 73, "тий", 73, "TODAY", 74, "сглеяа", 74, "NOW", 75, "тыяа", 75, "MENU.VISIBLE", 76, "епикоцес.жамеяес", 76, "MENUITEMS", 77, "епикоцес", 77 _
, "MENU", 78, "епикоцг", 78, "NUMBER", 79, "аяихлос", 79, "тилг", 79, "LAMBDA", 80, "калда", 81, "GROUP", 83, "олада", 83, "ARRAY", 84, "пимайас", 84, "[]", 85 _
, "сыяос", 86, "STACK", 86, "ISWINE", 87, "SHOW", 88, "охомг", 88, "OSBIT", 89, "WINDOW", 90, "сусйеуг", 90, "MONITOR.STACK", 91, "екецвос.сыяоу", 91, "MONITOR.STACK.SIZE", 92, "екецвос.лецехос.сыяоу", 92, "?", 93, "диаяхяысг", 94, "BUFFER", 94, "йатастасг", 95, "INVENTORY", 95, "LIST", 96, "киста", 96, "QUEUE", 97, "оуяа", 97, "INFINITY", 82, "апеияо", 82, "еккгмийа", 98, "GREEK", 98 _
, "INTERNET", 99, "диадийтуо", 99, "CLIPBOARD.IMAGE", 100, "пяовеияо.еийома", 100, "CLIPBOARD.DRAWING", 101, "пяовеияо.сведио", 101, "MONITORS", 102, "охомес", 102, "DOS", 103, "йомсока", 103, "SOUNDREC.LEVEL", 104, "гвоцяажгсгс.епипедо", 104, "ADDRESSOF", 105, "диеухумсгапо", 105, "пкатос.таимиас", 106, "MOVIE.WIDTH", 106, "уьос.таимиас", 107, "MOVIE.HEIGHT", 107)
If Not ahashbackup Is Nothing Then
For i = 0 To UBound(myid()) Step 2
    ahashbackup.ItemCreator CStr(myid(i)), CLng(myid(i + 1))
Next i
End If
For i = 0 To UBound(myid()) Step 2
    aHash.ItemCreator CStr(myid(i)), CLng(myid(i + 1))
Next i
myfun() = Array("PARAM(", 1, "паяал(", 1, "STACKITEM(", 2, "тилгсыяоу(", 2, "SGN(", 3, "сгл(", 3, "FRAC(", 4, "дей(", 4, "MATCH(", 5, "таутисг(", 5 _
, "LOCALE(", 6, "топийо(", 6, "FILELEN(", 7, "аявеиоу.лгйос(", 7, "TAB(", 8, "стгкг(", 8, "KEYPRESS(", 9, "патглемо(", 9, "INKEY(", 10, "емйол(", 10 _
, "тлгла(", 11, "MODULE(", 11, "басг(", 12, "MDB(", 12, "ASK(", 13, "яыта(", 13, "суцйяоусг(", 14, "COLLIDE(", 14, "лецехос.у(", 15, "SIZE.Y(", 15, "лецехос.в(", 16, "SIZE.X(", 16 _
, "WRITABLE(", 17, "еццяаьило(", 17, "COLOR(", 18, "COLOUR(", 18, "вяыла(", 18, "DIMENSION(", 19, "диастасг(", 19, "ARRAY(", 20, "пимайас(", 20 _
, "FUNCTION(", 21, "сумаятгсг(", 21, "DRIVE.SERIAL(", 22, "сеияиайос.дисйоу(", 22, "FILE.STAMP(", 23, "аявеиоу.сталпа(", 23, "EXIST.DIR(", 25, "упаявеи.йатакоцос(", 25 _
, "EXIST(", 26, "упаявеи(", 26, "JOYPAD(", 27, "кабг(", 27, "JOYPAD.DIRECTION(", 28, "кабг.йатеухумсг(", 28, "JOYPAD.ANALOG.X(", 29, "кабг.амакоцийо.в(", 29 _
, "JOYPAD.ANALOG.Y(", 30, "кабг.амакоцийо.у(", 30, "кабг.амакоцийо.ь(", 30, "IMAGE.X(", 31, "еийома.в(", 31, "IMAGE.Y(", 32, "еийома.у(", 32, "еийома.ь(", 32, "IMAGE.X.PIXELS(", 33, "еийома.в.сглеиа(", 33 _
, "IMAGE.Y.PIXELS(", 34, "еийома.у.сглеиа(", 34, "еийома.ь.сглеиа(", 34, "VALID(", 35, "ецйуяо(", 35, "EVAL(", 36, "ейжя(", 36, "ейжяасг(", 36, "POINT(", 37, "сглеио(", 37 _
, "CTIME(", 38, "упыяа(", 38, "CDATE(", 39, "уплея(", 39, "TIME(", 40, "вяомос(", 40, "DATE(", 41, "глеяа(", 41, "VAL(", 42, "тилг(", 42, "аниа(", 42, "RINSTR(", 107, "хесгдениа(", 43 _
, "INSTR(", 106, "хесг(", 44, "RECORDS(", 45, "еццяажес(", 45, "GROUP.COUNT(", 46, "олада.сумоко(", 46, "PARAGRAPH(", 47, "паяацяажос(", 47, "PARAGRAPH.INDEX(", 48, "аяихлос.паяацяажоу(", 48 _
, "BACKWARD(", 49, "писы(", 49, "FORWARD(", 50, "лпяоста(", 50, "DOC.PAR(", 51, "еццяажоу.пая(", 51, "MAX.DATA(", 52, "лецако.сеияас(", 52, "MIN.DATA(", 53, "лийяо.сеияас(", 53 _
, "MAX(", 54, "лецако(", 54, "MIN(", 55, "лийяо(", 55, "COMPARE(", 56, "суцйяиме(", 56, "DOC.UNIQUE.WORDS(", 57, "еццяажоу.ломадийес.кенеис(", 57, "DOC.WORDS(", 58, "еццяажоу.кенеис(", 58 _
, "DOC.LEN(", 59, "еццяажоу.лгйос(", 59, "LEN.DISP(", 60, "лгйос.елж(", 60, "LEN(", 61, "лгйос(", 61, "SQRT(", 62, "яифа(", 62, "FREQUENCY(", 63, "сувмотгта(", 63 _
, "LOG(", 64, "коц(", 64, "LN(", 65, "кж(", 65, "ATN(", 66, "тон.еж(", 66, "TAN(", 67, "ежап(", 67, "COS(", 68, "сум(", 68, "SIN(", 69, "гл(", 69, "ABS(", 70, "апок(", 70, "LOWORD(", 71, "LOWWORD(", 71, "йатылисо(", 71, "йахе(", 72, "EACH(", 73 _
, "HIWORD(", 74, "HIGHWORD(", 74, "памылисо(", 74, "BINARY.NEG(", 75, "дуадийо.амти(", 75, "дуадийо.амтистяожо(", 75, "BINARY.OR(", 76, "дуадийо.г(", 76 _
, "BINARY.AND(", 77, "дуадийо.йаи(", 77, "BINARY.XOR(", 78, "дуадийо.апо(", 78, "HILOWWORD(", 79, "дуолиса(", 79, "BINARY.SHIFT(", 80, "дуадийо.окисхгсг(", 80 _
, "BINARY.ROTATE(", 81, "дуадийг.пеяистяожг(", 81, "SINT(", 82, "айеяаио.дуадийо(", 82, "USGN(", 83, "дуадийо(", 83, "UINT(", 84, "дуадийо.айеяаио(", 84, "ROUND(", 85, "стяоцц(", 85 _
, "INT(", 86, "ай(", 86, "SEEK(", 87, "летахесг(", 87, "EOF(", 88, "текос(", 88, "RANDOM(", 89, "туваиос(", 89, "CHRCODE(", 90, "ваяйыд(", 90, "ASC(", 91, "йыд(", 91 _
, "GROUP(", 92, "олада(", 92, "TEST(", 93, "дойилг(", 93, "CONS(", 94, "емысг(", 94, "CAR(", 95, "пяыто(", 95, "CDR(", 96, "еполема(", 96, "сыяос(", 24, "STACK(", 24, "READY(", 97, "етоило(", 97, "PROPERTY(", 98, "идиотгта(", 98, "IF(", 99, "ам(", 99, "ORDER(", 100, "танг(", 100, "BANK(", 101, "тяап(", 101, "CEIL(", 102, "ояож(", 102, "FLOOR(", 86, "дапед(", 86, "еийома(", _
103, "IMAGE(", 103, "BUFFER(", 104, "диаяхяысг(", 104, "BINARY.NOT(", 105, "дуадийо.ови(", 105, "POINTER(", 108, "деийтгс(", 108, "BINARY.ADD(", 109, "дуадийо.пяосхесг(", 109, "дуадийо.пяо(", 109, "HSL(", 110, "вйж(", 110, "PLAYER(", 111, "паийтгс(", 111, "GETOBJECT(", 112, "ANTIKEIMENO(", 112, "VARPTR(", 113, "диеухл(", 113)
If Not bhashbackup Is Nothing Then
For i = 0 To UBound(myfun()) Step 2
    bhashbackup.ItemCreator CStr(myfun(i)), CLng(myfun(i + 1))
Next i
End If
For i = 0 To UBound(myfun()) Step 2
    bHash.ItemCreator CStr(myfun(i)), CLng(myfun(i + 1))
Next i
NumberId = True
End Function

Public Function allcommands(aHash As sbHash) As Boolean
Dim mycommands(), i As Long
mycommands() = Array("ABOUT", "AFTER", "APPEND", "APPEND.DOC", "ASSERT", "BACK", "BACKGROUND", "BASE", "BEEP", "BINARY", "BITMAPS", "BOOLEAN", "BOLD", "BREAK", "BROWSER", "BUFFER", "BYTE", "CALL", "CASE", "CAT", "CHANGE", "CHARSET", "CHOOSE.COLOR", "CHOOSE.FONT", "CHOOSE.OBJECT", "CHOOSE.ORGAN", "CIRCLE", "CLASS", "CLEAR", "CLIPBOARD", "CLOSE", "CLS", "CODEPAGE", "COLOR", "COMMIT", "COMPRESS", "CONST", "CONTINUE", "COPY", "CURRENCY", "CURSOR", "CURVE", "DATA", "DATE", "DB.PROVIDER", "DB.USER", "DECIMAL" _
, "DECLARE", "DEF", "DELETE", "DESKTOP", "DIM", "DIR", "DIV", "DO", "DOCUMENT", "DOS", "DOUBLE", "DRAW", "DRAWING", "DRAWINGS", "DROP", "DURATION", "EDIT", "EDIT.DOC", "ELSE", "ELSE.IF", "EMPTY", "END", "ENUM", "ENUMERATION", "ERASE", "ERROR", "ESCAPE", "EVENT", "EVERY", "EXECUTE", "EXIT", "EXPORT", "FAST", "FIELD", "FILES", "FILL", "FIND", "FKEY", "FLOODFILL", "FLUSH", "FONT", "FOR", "FORM", "FORMLABEL", "FRAME", "FUNCTION", "GET", "GLOBAL" _
, "GOSUB", "GOTO", "GRADIENT", "GREEK", "GROUP", "HALT", "HEIGHT", "HELP", "HEX", "HIDE", "HOLD", "HTML", "ICON", "IF", "IMAGE", "INLINE", "INPUT", "INSERT", "INTEGER", "INVENTORY", "ITALIC", "JOYPAD", "KEYBOARD", "LATIN", "LAYER", "LEGEND", "LET", "LINE", "LINESPACE", "LINK", "LIST", "LOAD", "LOAD.DOC", "LOCAL", "LOCALE", "LONG", "LOOP", "MAIN.TASK", "MARK", "MEDIA", "MENU", "MERGE.DOC", "METHOD", "MODE", "MODULE" _
, "MODULES", "MONITOR", "MOTION", "MOTION.W", "MOUSE.ICON", "MOVE", "MOVIE", "MOVIES", "MUSIC", "NAME", "NEW", "NEXT", "NORMAL", "OBJECT", "ON", "OPEN", "OPEN.FILE", "OPEN.IMAGE", "OPTIMIZATION", "ORDER", "OVER", "OVERWRITE", "PAGE", "PART", "PATH", "PEN", "PIPE", "PLAY", "PLAYER", "PLAYER(", "POLYGON", "PRINT", "PRINTER", "PRINTING", "PROFILER", "PROPERTIES", "PROTOTYPE", "PSET", "PUSH", "PUT", "READ", "RECURSION.LIMIT" _
, "REFER", "REFRESH", "RELEASE", "REM", "REMOVE", "REPEAT", "REPORT", "RESTART", "RETRIEVE", "RETURN", "SAVE", "SAVE.AS", "SAVE.DOC", "SCAN", "SCORE", "SCREEN.PIXELS", "SCRIPT", "SCROLL", "SEARCH", "SEEK", "SELECT", "SET", "SETTINGS", "SHIFT", "SHIFTBACK", "SHOW", "SINGLE", "SLOW", "SMOOTH", "SORT", "SOUND", "SOUNDREC", "SOUNDS", "SPEECH", "SPLIT", "SPRITE", "STACK", "START", "STATIC", "STEP", "STOCK", "STOP", "STRING", "STRUCTURE" _
, "SUB", "SUBDIR", "SUPERCLASS", "SWAP", "SWEEP", "SWITCHES", "TAB", "TABLE", "TARGET", "TARGETS", "TASK.MAIN", "TEST", "TEXT", "THEN", "THREAD", "THREAD.PLAN", "THREADS", "TITLE", "TONE", "TRY", "TUNE", "UPDATE", "USE", "USER", "VARIANT", "VERSION", "VIEW", "VOLUME", "WAIT", "WHILE", "WIDTH", "WIN", "WINDOW", "WITH", "WORDS", "WRITE", "WRITER", "адеиасе", "айеяаиос", "акт", "аккацг", "аккане", "аккиыс", "аккиыс.ам", "ам", "амафгтгсг" _
, "амахеыягсг", "амайтгсг", "амакоцио", "амакусг.охомгс", "амакутгс", "амаломг", "амамеысг", "амажояа", "амаье", "амехесе", "амоицла.аявеиоу", "амоицла.еийомас", "амоине", "амтецяаье", "амтицяаье", "амтийеилемо", "аниысг", "апая", "апаяихлгсг", "апедысе", "апкос", "апо", "апохгйеусг.ыс", "апойопг", "аяца", "аяихлос", "аявеиа", "аявеио", "аявг", "атупос", "аукос", "ауноуса", "ажаияесг", "ажгсе", "баке", "басг", "басг.паяовос", "басг.вягстгс", "баье", "бектистопоигсг" _
, "бгла", "богхеиа", "цецомос", "целисе", "цемийес", "цемийг", "цемийо", "циа", "цяалла", "цяаллатосеияа", "цяаллг", "цяаье", "цягцояа", "деийтг.лояжг", "деине", "дейаен", "дес", "диабасе", "диацяажг", "диайопг", "диайоптес", "диалесоу", "диаяхяысг", "диаяйеиа", "диастиво", "диажамеиа", "диажамо", "диажуцг", "диайопг", "диейоье", "диояхысе", "дипка", "дипкос", "дойилг", "долг", "дяолеас", "дуадийо", "дысе", "еццяажо", "еийома", "еийомес", "еийомидио" _
, "еисацыцг", "ейдосг", "ейтекесг", "ейтупысг", "ейтупытгс", "екецвос", "еккгмийа", "емхесг", "емтасг", "емы", "емысе", "енацыцг", "енодос", "епамакабе", "епамекабе", "епекене", "епекене.амтийеилемо", "епекене.цяаллатосеияа", "епекене.ояцамо", "епекене.вяыла", "епицяажг", "епийаияо", "епикене", "епикене.амтийеилемо", "епикене.цяаллатосеияа", "епикене.ояцамо", "епикене.вяыла", "епикоцес", "епикоцг", "епикоцгс" _
, "епипедо", "епистяожг", "епижамеиа", "еполемо", "етийета.жоялас", "еуяесг", "глеяолгмиа", "гвоцяажгсг", "гвои", "гвос", "хесе", "хесг", "идиотгтес", "исвмг", "ивмос", "йахаяг", "йахаяо", "йахе", "йакесе", "йалпукг", "йаме", "йамомийа", "йат", "йатакоцои", "йатакоцос", "йатастасг", "йатавыягсг", "йеилемо", "йемг", "йимгсг", "йимгсг.п", "йкасг", "йкеиди", "йкеисе", "йомсока", "йяата", "йяатгсе", "йяуье" _
, "йуйкийа", "йуйкос", "йукисг", "йуяио.еяцо", "кабг", "кахос", "катимийа", "кенеис", "киста", "коцийос", "коцистийос", "коцос", "лайяус", "ле", "леходос", "лекыдиа", "леяос", "лета", "летахесг", "лоусийг", "лпип", "мео", "мгла", "мглата", "нейима", "охомг", "олада", "олака", "омола", "ояио.амадяолгс", "ояисе", "паийтгс", "паине", "памы", "паяахуяо", "паяе", "паяелбокг", "павос", "педио", "пема", "пеяи" _
, "пеяихыяио", "пета", "пимайас", "пимайес", "пкациа", "пкаисио", "пкгйтяокоцио", "покуцымо", "пяос", "пяосхесе.еццяажо", "пяосхгйг", "пяытотупо", "пяовеияо", "яоутима", "яухлисеис", "с", "саяысе", "сбгсе", "сеияа", "секида", "семаяио", "сгл", "сглади", "стахеяг", "стахеяес", "статийг", "статийес", "стг", "стгм", "сто", "стой", "стовои", "стовос", "суццяажеас", "суццяажг", "суцвымеусе.еццяажо", "сулпиесг" _
, "сумаятгсг", "сумевисе", "сус", "сустгла", "сведио", "сведиа", "сведио.мглатым", "сыяос", "сысе", "сысе.еццяажо", "таимиа", "таимиес", "танг", "танимолгсг", "текос", "титкос", "тлгла", "тлглата", "томос", "топийа", "топийес", "топийг", "топийо", "тоте", "тупос", "тупысе", "упеяйкасг", "упойатакоцос", "жаядиа", "жеяе", "жеяеписы", "жомто", "жояла", "жоятысе" _
, "жоятысе.еццяажо", "жымг", "ваяайтгяес", "ваяане", "вягсг", "вягстг", "вягстгс", "вяыла", "вяылатисе", "ьгжио", "?")
For i = 0 To UBound(mycommands())

Select Case mycommands(i)
Case "SORT", "танимолгсг"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoSort)
Case "DEF", "йаме"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDef)
Case "NORMAL", "йамомийа"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoNormal)
Case "DOUBLE", "дипка", "дипкос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDouble)
Case "INTEGER", "айеяаиос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoInteger)
Case "LONG", "лайяус"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoLong)
Case "SINGLE", "апкос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoSingle)
Case "BOOLEAN", "коцийос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoBoolean)
Case "VARIANT", "атупос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoVariant)
Case "DECIMAL", "аяихлос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDecimal)
Case "CURRENCY", "коцистийос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoCurrency)
Case "BYTE", "ьгжио"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoByte)
Case "DATE", "глеяолгмиа"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDate)
Case "STRING", "цяалла"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoString)
Case "OBJECT", "амтийеилемо"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoObject)
Case "CURSOR", "дяолеас"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoTextCursor)
Case "MOUSE.ICON", "деийтг.лояжг"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoMouseIcon)
Case "FLOODFILL", "целисе"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoFloodFill)
Case "FILL", "баье"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoFill)
Case "IMAGE", "еийома"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoImage)
Case "RELEASE", "ажгсе"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoRelease)
Case "HOLD", "йяатгсе"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoHold)
Case "SUPERCLASS", "упеяйкасг"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoSuperClass)
Case "CLASS", "йкасг"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoClass)
Case "DIM", "пимайас", "пимайес"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDIM)
Case "PATH", "COLOR", "вяыла", "ивмос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPathDraw)
Case "TITLE", "титкос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoTitle)
Case "NEW", "мео"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoNew)
Case "MODULE", "тлгла"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoModule)
Case "GROUP", "олада"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoGroup)
Case "DRAWINGS", "сведиа"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDrawings)
Case "BITMAPS", "еийомес"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoBitmaps)
Case "MOVIES", "таимиес"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoMovies)
Case "SOUNDS", "гвои"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoSounds)
Case "FUNCTION", "сумаятгсг"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoFunction)
Case "STEP", "бгла"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoStep)
Case "COPY", "амтецяаье", "амтицяаье"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoCopy)
Case "охомг", "CLS"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoCls)
Case "пема", "PEN"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPen)
Case "WAIT", "амаломг"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoWait)
Case "EVENT", "цецомос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoEvent)
Case "SET", "хесе"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoSet)
Case "INPUT", "еисацыцг"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoInput)
Case "CLEAR", "йахаяо"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoClear)
Case "DECLARE", "ояисе"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDeclare)
Case "METHOD", "леходос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoMethod)
Case "WITH", "ле"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoWith)
Case "DATA", "сеияа"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoData)
Case "PUSH", "баке"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPush)
Case "SWAP", "аккане"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoSwap)
Case "COMMIT", "амехесе"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoComm)
Case "REFER", "апедысе"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoRef)
Case "READ", "диабасе"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoRead)
Case "LET", "стг", "стгм", "сто"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoLet)
Case "PRINT", "тупысе", "?"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPrint)
Case "CALL", "йакесе"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 38
Case "CHOOSE.OBJECT", "епекене.амтийеилемо", "епикене.амтийеилемо"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoChooseObj)
Case "CHOOSE.FONT", "епекене.цяаллатосеияа", "епикене.цяаллатосеияа"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoChooseFont)
Case "REM", "сгл"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoRem)
Case "LINESPACE", "диастиво"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoLinespace)
Case "BOLD", "жаядиа"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoBold)
Case "MODE", "тупос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoMode)
Case "GRADIENT", "жомто"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoGradient)
Case "FILES", "аявеиа"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoFiles)
Case "CAT", "йатакоцои", "йат"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoCat)
Case "MOVE", "хесг"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoMove)
Case "ваяане", "DRAW"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDraw)
Case "WIDTH", "павос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoWidth)
Case "покуцымо", "POLYGON"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPoly)
Case "CIRCLE", "йуйкос"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoCircle)
Case "йалпукг", "CURVE"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoCurve)
Case "TEXT", "йеилемо"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoText)
Case "HTML"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoHtml)
Case "STRUCTURE", "долг"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoStructure)
Case "басг", "BASE"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoBase)
Case "аявеио", "TABLE"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoTable)
Case "ейтекесг", "EXECUTE"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoExecute)
Case "амайтгсг", "RETRIEVE"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoRetr)
Case "амафгтгсг", "SEARCH"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoSearch)
Case "пяосхгйг", "APPEND"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoAppend)
Case "ажаияесг", "DELETE"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDelete)
Case "танг", "ORDER"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoOrder)
Case "сулпиесг", "COMPRESS"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoCompact)
Case "LAYER", "епипедо"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoLayer)
Case "PRINTER", "ейтупытгс"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPrinter)
Case "PAGE", "секида"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPage)
Case "PLAYER", "паийтгс"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPlayer)
Case "SPRITE", "диажамо", "диажамеиа"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoSprite)
Case "MODULES", "тлглата"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoModules)
Case "CLIPBOARD", "пяовеияо"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoClipBoard)
Case "паине", "PLAY"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPlayScore)
Case "SCORE", "жымг"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoScore)
Case "REPORT", "амажояа"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoReport)
Case "BACK", "BACKGROUND", "пеяихыяио"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoBack)
Case "OVER", "памы"
    aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoOver)
Case "PROTOTYPE", "пяытотупо"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoProto)
Case "SHIFTBACK", "жеяеписы"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoShiftBack)
Case "SHIFT", "жеяе"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoShift)
Case "LOAD", "жоятысе"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoLoad)
Case "DROP", "пета"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDrop)
Case "апаяихлгсг", "апая", "ENUMERATION", "ENUM"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoEnum)
Case "DESKTOP", "епижамеиа"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoDesktop)
Case "вяылатисе", "PSET"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoPset)
Case "ASSERT", "аниысг"
aHash.ItemCreator CStr(mycommands(i)), ProcPtr(AddressOf NeoAssert)
Case "IF", "ам"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 1
Case "ELSE", "аккиыс"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 2
Case "ELSE.IF", "аккиыс.ам"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 3
Case "SELECT", "епекене", "епикене"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 4
Case "TRY", "дес"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 5
Case "FOR", "циа"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 6
Case "NEXT", "еполемо"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 7
Case "REFRESH", "амамеысг"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 8
Case "емы", "WHILE"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 9
Case "DO", "REPEAT", "епамакабе", "епамекабе"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 10
Case "GOTO", "пяос"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 11
Case "SUB", "яоутима"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 12
Case "GOSUB", "диалесоу"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 13
Case "апо", "ON"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 14
Case "LOOP", "йуйкийа"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 15
Case "BREAK", "диейоье"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 16
Case "CONTINUE", "сумевисе"
     aHash.ItemCreator2 CStr(mycommands(i)), 0, 17
Case "RESTART", "нейима"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 18
Case "RETURN", "епистяожг"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 19
Case "END", "текос"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 20
Case "енодос", "EXIT"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 21
Case "INLINE", "емхесг"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 22
Case "UPDATE", "епийаияо"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 23
Case "мгла", "THREAD"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 24
Case "AFTER", "лета"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 25
Case "леяос", "PART"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 26
Case "статийг", "статийес", "STATIC"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 27
Case "йахе", "EVERY"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 28
Case "йуяио.еяцо", "MAIN.TASK", "TASK.MAIN"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 29
Case "SCAN", "саяысе"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 30
Case "TARGET", "стовос"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 31
Case "LOCAL", "топийа", "топийг", "топийес"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 34  ' Local A, B, C=10, K
Case "GLOBAL", "цемийо", "цемийг", "цемийес"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 35   'Global A, B=6, X
Case "CONST", "стахеяг", "стахеяес"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 36
Case "BINARY", "дуадийо"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 39
Case "HALT", "акт"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 40
Case "STOP", "диайопг"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 41
Case "DRAWING", "сведио"
    aHash.ItemCreator2 CStr(mycommands(i)), 0, 42
Case Else
    aHash.ItemCreator CStr(mycommands(i)), 0
End Select
Next i
allcommands = True
End Function
Private Function ProcPtr(ByVal nAddress As Long) As Long
    ProcPtr = nAddress
End Function
Public Sub StoreFont(aName$, aSize As Single, ByVal aCharset As Long)
Err.Clear
On Error Resume Next
If aSize < 1 Then aSize = 1
fonttest.Font.Size = aSize
If Err.Number > 0 Then aSize = 12: fonttest.Font.Size = aSize
    fonttest.FontName = aName$
    fonttest.Font.bold = True
    fonttest.Font.Italic = True
    fonttest.Font.charset = aCharset
        fonttest.FontName = aName$
    fonttest.Font.bold = True
    fonttest.Font.Italic = True
    fonttest.Font.charset = aCharset
    fonttest.Font.Size = aSize
    aSize = fonttest.Font.Size '' return
End Sub
Public Function InternalLeadingSpace() As Long
On Error Resume Next
    GetTextMetrics fonttest.hDC, TM
  With TM
InternalLeadingSpace = (.tmInternalLeading = 0) Or Not (.tmInternalLeading > 0)
End With
End Function
Public Function AverCharSpace(DDD As Object, Optional breakchar) As Long
On Error Resume Next
Dim tmm As TEXTMETRIC
    GetTextMetrics DDD.hDC, tmm
  With tmm
AverCharSpace = .tmAveCharWidth
breakchar = .tmBreakChar
End With
End Function
Sub TimeZones(zHash As FastCollection)
Dim zones(), i As Long
zones() = Array("Dateline Standard Time", -12, "UTC-11", -11, "Aleutian Standard Time", -10, "Hawaiian Standard Time", -10, "Marquesas Standard Time", -8.5, "Alaskan Standard Time", -9, "UTC-09", -9, "Pacific Standard Time (Mexico)", -8, "UTC-08", -8, "Pacific Standard Time", -8, "US Mountain Standard Time", -7, "Mountain Standard Time (Mexico)", -7, _
"Mountain Standard Time", -7, "Central America Standard Time", -6, "Central Standard Time", -6, "Easter Island Standard Time", -6, "Central Standard Time (Mexico)", -6, "Canada Central Standard Time", -6, "SA Pacific Standard Time", -5, "Eastern Standard Time (Mexico)", -5, "Eastern Standard Time", -5, "Haiti Standard Time", -5, "Cuba Standard Time", -5, "US Eastern Standard Time", -5, _
"Turks And Caicos Standard Time", -5, "Paraguay Standard Time", -4, "Atlantic Standard Time", -4, "Venezuela Standard Time", -4, "Central Brazilian Standard Time", -4, "SA Western Standard Time", -4, "Pacific SA Standard Time", -4, "Newfoundland Standard Time", -2.5, "Tocantins Standard Time", -3, "E. South America Standard Time", -3, "SA Eastern Standard Time", -3, "Argentina Standard Time", -3, _
"Greenland Standard Time", -3, "Montevideo Standard Time", -3, "Magallanes Standard Time", -3, "Saint Pierre Standard Time", -3, "Bahia Standard Time", -3, "UTC-02", -2, "Mid-Atlantic Standard Time", -2, "Azores Standard Time", -1, "Cape Verde Standard Time", -1, "UTC", 0, "Morocco Standard Time", 0, "GMT Standard Time", 0, _
"Greenwich Standard Time", 0, "W. Europe Standard Time", 1, "Central Europe Standard Time", 1, "Romance Standard Time", 1, "Sao Tome Standard Time", 1, "Central European Standard Time", 1, "W. Central Africa Standard Time", 1, "Jordan Standard Time", 2, "GTB Standard Time", 2, "Middle East Standard Time", 2, "Egypt Standard Time", 2, "E. Europe Standard Time", 2, _
"Syria Standard Time", 2, "West Bank Standard Time", 2, "South Africa Standard Time", 2, "FLE Standard Time", 2, "Israel Standard Time", 2, "Kaliningrad Standard Time", 2, "Sudan Standard Time", 2, "Libya Standard Time", 2, "Namibia Standard Time", 2, "Arabic Standard Time", 3, "Turkey Standard Time", 3, "Arab Standard Time", 3, _
"Belarus Standard Time", 3, "Russian Standard Time", 3, "E. Africa Standard Time", 3, "Iran Standard Time", 3.5, "Arabian Standard Time", 4, "Astrakhan Standard Time", 4, "Azerbaijan Standard Time", 4, "Russia Time Zone 3", 4, "Mauritius Standard Time", 4, "Saratov Standard Time", 4, "Georgian Standard Time", 4, "Caucasus Standard Time", 4, _
"Afghanistan Standard Time", 4.5, "West Asia Standard Time", 5, "Ekaterinburg Standard Time", 5, "Pakistan Standard Time", 5, "India Standard Time", 5.5, "Sri Lanka Standard Time", 5.5, "Nepal Standard Time", 5.75, "Central Asia Standard Time", 6, "Bangladesh Standard Time", 6, "Omsk Standard Time", 6, "Myanmar Standard Time", 6.5, "SE Asia Standard Time", 7, _
"Altai Standard Time", 7, "W. Mongolia Standard Time", 7, "North Asia Standard Time", 7, "N. Central Asia Standard Time", 7, "Tomsk Standard Time", 7, "China Standard Time", 8, "North Asia East Standard Time", 8, "Singapore Standard Time", 8, "W. Australia Standard Time", 8, "Taipei Standard Time", 8, "Ulaanbaatar Standard Time", 8, "Aus Central W. Standard Time", 8.75, _
"Transbaikal Standard Time", 9, "Tokyo Standard Time", 9, "North Korea Standard Time", 9, "Korea Standard Time", 9, "Yakutsk Standard Time", 9, "Cen. Australia Standard Time", 9.5, "AUS Central Standard Time", 9.5, "E. Australia Standard Time", 10, "AUS Eastern Standard Time", 10, "West Pacific Standard Time", 10, "Tasmania Standard Time", 10, "Vladivostok Standard Time", 10, _
"Lord Howe Standard Time", 10.5, "Bougainville Standard Time", 11, "Russia Time Zone 10", 11, "Magadan Standard Time", 11, "Norfolk Standard Time", 11, "Sakhalin Standard Time", 11, "Central Pacific Standard Time", 11, "Russia Time Zone 11", 12, "New Zealand Standard Time", 12, "UTC+12", 12, "Fiji Standard Time", 12, "Kamchatka Standard Time", 12, _
"Chatham Islands Standard Time", 12.75, "UTC+13", 13, "Tonga Standard Time", 13, "Samoa Standard Time", 13, "Line Islands Standard Time", 14)
For i = 0 To UBound(zones) - 1 Step 2
zHash.AddKey zones(i), zones(i + 1)
If InStr(zones(i), "Standard") > 1 Then
zHash.AddKey Replace(zones(i), "Standard", "Daylight"), zones(i + 1) + 1
End If
Next i
zHash.Sort
End Sub
Public Function HD(a$) As Long
Dim ret As Long
ret = HashData(StrPtr(a$), LenB(a$), VarPtr(HD), 4)
HD = HD And &H7FFFFFFF
If HD = 0 Then HD = 1
End Function
Public Function HD1(aa As Currency) As Long
Dim ret As Long
ret = HashData(VarPtr(aa), 8, VarPtr(HD1), 4)
HD1 = HD1 And &H7FFFFFFF
If HD1 = 0 Then HD1 = 1
End Function
Sub Main()
'' not used
'' If App.StartMode = vbSModeStandalone Then NeoSubMain
Dim m As New Callback

m.Run "start"
If m.status = 0 Then
m.Cli Form1.commandW, ">"
m.Reset
End If
m.ShowGui = False
Debug.Print "ok"
ResetTokenFinal
End Sub
Function ProcWriter(basestack As basetask, rest$, Lang As Long) As Boolean
Dim prive As Long
prive = GetCode(basestack.Owner)
If Lang = 1 Then
PlainBaSket basestack.Owner, players(prive), "George Karras (C), Kallithea Attikis, Greece 1999-2022"
Else
PlainBaSket basestack.Owner, players(prive), ListenUnicode(915, 953, 974, 961, 947, 959, 962, 32, 922, 945, 961, 961, 940, 962, 32, 40, 67, 41, 44, 32, 922, 945, 955, 955, 953, 952, 941, 945, 32, 913, 964, 964, 953, 954, 942, 962, 44, 32, 917, 955, 955, 940, 948, 945, 32, 49, 57, 57, 57, 45, 50, 48, 50, 48)
End If
crNew basestack, players(prive)
ProcWriter = True
End Function

Sub SendAKey(ByVal keycode As Integer, ByVal shift As Boolean, ByVal ctrl As Boolean, ByVal alt As Boolean)
Dim extended As Byte, Map As Integer, smap As Integer, cmap As Integer, amap As Integer, cap As Long, old As Long
Const key_release As Byte = 2
If keycode > 500 Then extended = 1: keycode = keycode - 500
If extended = 0 Then
If keycode > 64 And keycode < 91 Then
    If Not CapsLockOn() Then shift = Not shift
End If
End If

Map = MapVirtualKey(keycode, 0)
smap = MapVirtualKey(&H10, 0)
cmap = MapVirtualKey(&H11, 0)
amap = MapVirtualKey(&H12, 0)


keycode = keycode Mod 255
' press key
If shift Then keybd_event &H10, smap, 0, 0
If ctrl Then keybd_event &H11, cmap, 0, 0
If alt Then keybd_event &H12, amap, 0, 0
keybd_event keycode, Map, extended, 0

' release key
keybd_event keycode, Map, KEYEVENTF_KEYUP + extended, 0
If shift Then keybd_event &H10, smap, KEYEVENTF_KEYUP, 0
If ctrl Then keybd_event &H11, cmap, KEYEVENTF_KEYUP, 0
If alt Then keybd_event &H12, amap, KEYEVENTF_KEYUP, 0

End Sub
Public Function CapsLockOn() As Boolean
    GetKeyboardState kbArray
    CapsLockOn = (kbArray.kbByte(VK_CAPITAL) And 1) = 1
End Function
Public Function WinKey() As Boolean
    Const VK_RWIN As Long = &H5C
    Const VK_LWIN As Long = &H5B
 '   GetKeyboardState kbArray
'    WinKey = ((kbArray.kbByte(VK_RWIN) And 1) = 1) Or ((kbArray.kbByte(VK_LWIN) And 1) = 1)
    
    WinKey = KeyPressedLong(VK_RWIN) Or KeyPressedLong(VK_LWIN)
    
    
    
End Function
Public Sub ChangeTo408()
    Dim old
    old = ActivateKeyboardLayout((408), &H100&)
    If old = 0 Then
    LoadKeyboardLayout "00000408", KLF_ACTIVATE
    End If
End Sub
Public Sub ChangeTo409()
    Dim old
    old = ActivateKeyboardLayout((409), &H100&)
    If old = 0 Then
    LoadKeyboardLayout "00000409", KLF_ACTIVATE
    End If
End Sub
Public Function Keyboards(what) As String
   Keyboards = GetlocaleString2(94&, what)
End Function

Sub GetQrCode(basestack As basetask, a$, ab$)
Dim ErrLevel As Long, X As Variant, QRcolor As Long, sq As Boolean
sq = True
On Error Resume Next
If FastSymbol(a$, ",") Then
    If IsExp(basestack, a$, X, , True) Then
        QRcolor = mycolor(X)
    End If
Else
    GoTo step2
End If

If FastSymbol(a$, ",") Then
    If IsExp(basestack, a$, X, , True) Then
        If X = 1 Then
            ErrLevel = 1
        ElseIf X = 2 Then
            ErrLevel = 2
        ElseIf X = 3 Then
            ErrLevel = 3
        End If
    End If
    If FastSymbol(a$, ",") Then
    If IsExp(basestack, a$, X, , True) Then
        If X = 0 Then sq = False
    End If
End If

End If
step2:
Set basestack.lastobj = Nothing
Dim aPic As MemBlock, usehandler As mHandler
Dim bytes As Long
Dim emfP As StdPicture
Set emfP = QRCodegenBarcode(ab$, QRcolor, , sq, ErrLevel)
If emfP Is Nothing Then Exit Sub
Set aPic = New MemBlock
bytes = GetEnhMetaFileBits(emfP.Handle, bytes, ByVal 0)
If bytes Then
    aPic.Construct 1, bytes
    Call GetEnhMetaFileBits(emfP.Handle, bytes, ByVal aPic.GetBytePtr(0))
    aPic.SubType = 2 ' emf
    Set usehandler = New mHandler
    usehandler.t1 = 2
    Set usehandler.objref = aPic
    Set basestack.lastobj = usehandler
End If
End Sub
Sub testA()
Dim a
a = cInt64("&HAAFFAAFF00000001")
a = cInt64("&H8000000080000001")
Debug.Print Hex$(LowLong(a))
Debug.Print Hex$(HighLong(a))
Debug.Print Hex64$(LowLong(a), 4)
Debug.Print Hex64$(HighLong(a), 4)
Debug.Print Hex64$(CDec(a))
End Sub
