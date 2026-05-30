VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ShellPipe (from version 7)
' Converted from a control to a class

Private Const WIN32NULL As Long = 0&
Private Const WIN32FALSE As Long = 0&
Private Const WIN32TRUE As Long = 1&
Private Const WAIT_OBJECT_0 As Long = 0&
Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const INFINITE As Long = -1&
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const SW_HIDE As Long = 0&
Private Const STD_INPUT_HANDLE As Long = -10&
Private Const STD_OUTPUT_HANDLE As Long = -11&
Private Const STD_ERROR_HANDLE As Long = -12&
Private Const HANDLE_FLAG_INHERIT As Long = &H1&
Private Const CTRL_C_EVENT As Long = 0&
Private Const CTRL_BREAK_EVENT As Long = 1&
Private Const ERROR_ACCESS_DENIED As Long = 5&
Private Const ERROR_INVALID_HANDLE As Long = 6&
Private Const ERROR_HANDLE_EOF As Long = 38&
Private Const ERROR_BROKEN_PIPE As Long = 109&

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESSINFO
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Function CreatePipe Lib "kernel32" _
    (ByRef phReadPipe As Long, _
     ByRef phWritePipe As Long, _
     ByRef lpPipeAttributes As Any, _
     ByVal nSize As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" _
    (ByVal ApplicationName As String, _
     ByVal CommandLine As String, _
     ByVal lpProcessAttributes As Long, _
     ByVal lpThreadAttributes As Long, _
     ByVal bInheritHandles As Long, _
     ByVal dwCreationFlags As Long, _
     ByVal lpEnvironment As Long, _
     ByVal CurrentDirectory As String, _
     ByRef StartupInformation As STARTUPINFO, _
     ByRef ProcessInformation As PROCESSINFO) As Long

Private Declare Function GenerateConsoleCtrlEvent Lib "kernel32" _
    (ByVal dwCtrlEvent As Long, _
     ByVal dwProcessGroupId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hHandle As Long, _
     ByRef lpExitCode As Long) As Long

Private Declare Function PeekNamedPipe Lib "kernel32" _
    (ByVal hNamedPipe As Long, _
     ByVal lpBuf As Long, _
     ByVal nBufSize As Long, _
     ByRef lpBytesRead As Long, _
     ByRef lpTotalBytesAvail As Long, _
     ByVal lpBytesLeftThisMessage As Long) As Long

Private Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, _
     ByVal lpBuf As Long, _
     ByVal nNumberOfBytesToRead As Long, _
     ByRef lpNumberOfBytesRead As Long, _
     ByVal lpOverlapped As Any) As Long

Private Declare Function SetHandleInformation Lib "kernel32" _
    (ByVal hObject As Long, _
     ByVal dwMask As Long, _
     ByVal dwFlags As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hHandle As Long, _
     ByVal uExitCode As Long) As Long

Private Declare Function WaitForInputIdle Lib "user32" ( _
    ByVal hProcess As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, _
     ByVal dwMilliseconds As Long) As Long

Private Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, _
     ByVal lpBuf As String, _
     ByVal cToWrite As Long, _
     ByRef cWritten As Long, _
     ByVal lpOverlapped As Any) As Long

Private piProc As PROCESSINFO
Private saPipe As SECURITY_ATTRIBUTES
Private hChildInPipeRd As Long
Private hChildInPipeWr As Long
Private hChildOutPipeRd As Long
Private hChildOutPipeWr As Long
Private hChildErrPipeRd As Long
Private hChildErrPipeWr As Long
Private blnFinishedChild As Boolean
Private blnProcessActive As Boolean
Private PipeOpenIn As Boolean
Private PipeOpenOut As Boolean
Private PipeOpenErr As Boolean
Private BufferIn As SPBuffer
Private BufferOut As SPBuffer
Private BufferErr As SPBuffer

Private mErrAsOut As Boolean
Private mWaitForIdle As Long

Public Enum SPEOF_TYPES
    SPEOF_NORMAL = 0
    SPEOF_BROKEN_PIPE = ERROR_BROKEN_PIPE
End Enum

Public Enum SP_RESULTS
    SP_SUCCESS = 0
    SP_CREATEPIPEFAILED = &H80042B00
    SP_CREATEPROCFAILED = &H80042B01
End Enum
Private mCp, mLCid As Long
Public Event DataArrival(ByVal CharsTotal As Long)
Public Event ErrDataArrival(ByVal CharsTotal As Long)
Public Event EOF(ByVal EOFType As SPEOF_TYPES)
Public Event ErrEOF(ByVal EOFType As SPEOF_TYPES)
Public Event Error(ByVal Number As Long, _
                   ByVal Source As String, _
                   CancelDisplay As Boolean)
Public Event ChildFinished()
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( _
    ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String, _
    ByVal cchData As Long _
) As Long

Private Const LOCALE_IDEFAULTANSICODEPAGE As Long = &H1004

Public Function CodePageFromLCID(ByVal lcid As Long) As Long
    Dim sBuffer As String
    Dim n As Long

    sBuffer = String$(16, vbNullChar)

    n = GetLocaleInfo(lcid, LOCALE_IDEFAULTANSICODEPAGE, sBuffer, Len(sBuffer))

    If n > 0 Then
        CodePageFromLCID = CLng(Left$(sBuffer, n - 1))
    Else
        CodePageFromLCID = 0
    End If
End Function
Public Property Get Active() As Boolean
    If blnProcessActive Then 'Last we knew, it was active.
        If WaitForSingleObject(piProc.hProcess, 0&) <> WAIT_OBJECT_0 Then
            Active = True
        Else
            blnProcessActive = False
            Active = False
        End If
    Else
        Active = False
    End If
End Property

Public Sub ClosePipe()
    CloseHandle hChildInPipeWr
    PipeOpenIn = False
End Sub

Public Function FinishChild(Optional ByVal WaitMs As Long = 500, _
                            Optional ByVal KillReturnCode = 666) As Long
    If blnFinishedChild Then
        FinishChild = -1 'Already "Finished Child."
    Else
        If blnProcessActive Then
            If WaitForSingleObject(piProc.hProcess, WaitMs) <> WAIT_OBJECT_0 Then
                TerminateProcess piProc.hProcess, KillReturnCode
                WaitForSingleObject piProc.hProcess, INFINITE
            End If
            blnProcessActive = False
        End If
        
        BufferErr.Clear
        BufferOut.Clear
        BufferIn.Clear
        GetExitCodeProcess piProc.hProcess, FinishChild
        If PipeOpenOut Then ClosePipeOut
        If PipeOpenErr Then ClosePipeErr
        If PipeOpenIn Then ClosePipe
        CloseHandle piProc.hThread
        CloseHandle piProc.hProcess
        blnFinishedChild = True
    End If
End Function

Public Property Get ErrAsOut() As Boolean
    ErrAsOut = mErrAsOut
End Property

Public Property Let ErrAsOut(ByVal RHS As Boolean)
    mErrAsOut = RHS
End Property

Public Function GetData(Optional ByVal MaxLen As Long = -1) As String
    GetData = BufferOut.GetData(MaxLen)
End Function

Public Function GetLine() As String
    GetLine = BufferOut.GetLine()
End Function

Public Property Get HasLine() As Boolean
    HasLine = BufferOut.HasLine
End Property

Public Function ErrGetData(Optional ByVal MaxLen As Long = -1) As String
    ErrGetData = BufferErr.GetData(MaxLen)
End Function

Public Function ErrGetLine() As String
    ErrGetLine = BufferErr.GetLine()
End Function

Public Property Get ErrHasLine() As Boolean
    ErrHasLine = BufferErr.HasLine
End Property

Public Sub Interrupt(Optional ByVal Break As Boolean = False)
    Dim lngEvent As Long
    Dim ErrNum As Long
    Dim Cancel As Boolean
    
    lngEvent = IIf(Break, CTRL_BREAK_EVENT, CTRL_C_EVENT)
    If GenerateConsoleCtrlEvent(lngEvent, 0&) = 0 Then
        ErrNum = Err.LastDllError
        RaiseEvent Error(ErrNum, "ShellPipe.Interrupt.ConsoleCtrlEvent", Cancel)
        If Not Cancel Then
            Err.Raise ErrNum, Typename(Me), "Interrupt ConsoleCtrlEvent error"
        End If
    End If
End Sub

Public Property Get Length() As Long
    Length = BufferOut.Length
End Property

Public Property Get ErrLength() As Long
    ErrLength = BufferErr.Length
End Property

Public Function Run( _
    ByVal CommandLine As String, _
    Optional ByVal CurrentDir As String = vbNullString) _
    As SP_RESULTS
    
    Dim siStart As STARTUPINFO
    
    With saPipe
        .nLength = Len(saPipe)
        .lpSecurityDescriptor = WIN32NULL
        .bInheritHandle = WIN32TRUE
    End With
    
    If CreatePipe(hChildOutPipeRd, hChildOutPipeWr, saPipe, 0&) = WIN32FALSE Then
        Run = SP_CREATEPIPEFAILED
        Exit Function
    End If
    SetHandleInformation hChildOutPipeRd, HANDLE_FLAG_INHERIT, 0&
    
    If CreatePipe(hChildErrPipeRd, hChildErrPipeWr, saPipe, 0&) = WIN32FALSE Then
        CloseHandle hChildOutPipeRd
        CloseHandle hChildOutPipeWr
        Run = SP_CREATEPIPEFAILED
        Exit Function
    End If
    SetHandleInformation hChildErrPipeRd, HANDLE_FLAG_INHERIT, 0&
    
    If CreatePipe(hChildInPipeRd, hChildInPipeWr, saPipe, 0&) = WIN32FALSE Then
        CloseHandle hChildOutPipeRd
        CloseHandle hChildOutPipeWr
        CloseHandle hChildErrPipeRd
        CloseHandle hChildErrPipeWr
        Run = SP_CREATEPIPEFAILED
        Exit Function
    End If
    SetHandleInformation hChildInPipeWr, HANDLE_FLAG_INHERIT, 0&
    
    With siStart
        .cb = Len(siStart)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .hStdOutput = hChildOutPipeWr
        .hStdError = hChildErrPipeWr
        .hStdInput = hChildInPipeRd
        'Leave other fields 0/Null.
    End With
    
    'Clear all fields, global UDT and we may have been used more than once.
    With piProc
        .hProcess = 0
        .hThread = 0
        .dwProcessID = 0
        .dwThreadID = 0
    End With
    
    'In case "" was passed to us:
    If Len(CurrentDir) = 0 Then CurrentDir = vbNullString
        
    If CreateProcessA(vbNullString, CommandLine, WIN32NULL, WIN32NULL, WIN32TRUE, _
                      NORMAL_PRIORITY_CLASS, WIN32NULL, CurrentDir, _
                      siStart, piProc) = WIN32FALSE Then
        blnProcessActive = False
        Run = SP_CREATEPROCFAILED
    Else
        CloseHandle hChildOutPipeWr
        CloseHandle hChildErrPipeWr
        CloseHandle hChildInPipeRd
        blnProcessActive = True
        blnFinishedChild = False
        PipeOpenIn = True
        PipeOpenOut = True
        PipeOpenErr = True
        If WaitForIdle > 0 Then WaitForInputIdle piProc.hProcess, WaitForIdle
        Run = SP_SUCCESS
    End If
End Function

Public Sub SendData(ByVal Data As String)



    BufferIn.Append Data
    WriteData
End Sub

Public Sub SendLine(ByVal Line As String, Optional ByVal UseLFs As Boolean = True)
    If UseLFs Then
        SendData Line & vbCrLf
    Else
        SendData Line & vbCr
    End If
End Sub

Public Property Get WaitForIdle() As Long
    WaitForIdle = mWaitForIdle
End Property

Public Property Let WaitForIdle(ByVal RHS As Long)
    If RHS < 0 Then Err.Raise &H80042B03, Typename(Me), "WaitForIdle must be >= 0"
    mWaitForIdle = RHS
End Property

Public Sub ProcessLoop()
    If Active Then
        ReadData
     '   MyDoEvents 'Let events raised in ReadData be handled.
        WriteData
    Else
        'Last gasp.
        ReadData
     '   MyDoEvents 'Let events raised in ReadData be handled.
        RaiseEvent ChildFinished
    End If
End Sub

Private Sub ClosePipeOut()
    CloseHandle hChildOutPipeRd
    PipeOpenOut = False
End Sub

Private Sub ClosePipeErr()
    CloseHandle hChildErrPipeRd
    PipeOpenErr = False
End Sub

Private Sub ReadData()
    Dim Buffer As String
    Dim AvailChars As Long
    Dim CharsRead As Long
    Dim ErrNum As Long
    Dim Cancel As Boolean
    
    If PipeOpenOut Then
        If PeekNamedPipe(hChildOutPipeRd, WIN32NULL, 0&, WIN32NULL, AvailChars, WIN32NULL) <> WIN32FALSE Then
            If AvailChars > 0 Then
                Buffer = String$(AvailChars + 100, 0)
                
                'If ReadFile(hChildOutPipeRd, StrPtr(Buffer), AvailChars, CharsRead, WIN32NULL) <> WIN32FALSE Then
                CharsRead = AvailChars
                If FileReadString2(hChildOutPipeRd, Buffer, CharsRead) Then
                    If CharsRead > 0 Then
                        BufferOut.Append Buffer
                        'BufferOut.Append Left$(Buffer, CharsRead)
                        RaiseEvent DataArrival(BufferOut.Length)
                    Else
                        RaiseEvent EOF(SPEOF_NORMAL)
                        ClosePipeOut
                    End If
                Else
                    ErrNum = Err.LastDllError
                    If ErrNum = ERROR_BROKEN_PIPE Then
                        RaiseEvent EOF(SPEOF_BROKEN_PIPE)
                        ClosePipeOut
                    Else
                        RaiseEvent Error(ErrNum, "ShellPipe.ReadData.ReadFile", Cancel)
                        If Not Cancel Then
                            Err.Raise ErrNum, Typename(Me), "ReadData ReadFile error"
                        End If
                    End If
                End If
            End If
        Else
            ErrNum = Err.LastDllError
            Select Case ErrNum
                Case ERROR_BROKEN_PIPE
                    RaiseEvent EOF(SPEOF_BROKEN_PIPE)
                    ClosePipeOut
                    
                Case ERROR_ACCESS_DENIED, ERROR_INVALID_HANDLE
                    'Ignore as "no input."
                    
                Case Else
                    RaiseEvent Error(ErrNum, "ShellPipe.ReadData.PeekNamedPipe", Cancel)
                    If Not Cancel Then
                        Err.Raise ErrNum, Typename(Me), "ReadData PeeknamedPipe error"
                    End If
            End Select
        End If
    End If
    
    If PipeOpenErr Then
        If PeekNamedPipe(hChildErrPipeRd, WIN32NULL, 0&, WIN32NULL, AvailChars, WIN32NULL) <> WIN32FALSE Then
            If AvailChars > 0 Then
                Buffer = String$(AvailChars, 0)
                'If ReadFile(hChildErrPipeRd, ByVal Buffer, AvailChars, CharsRead, WIN32NULL) <> WIN32FALSE Then
                
                CharsRead = AvailChars
                
                If FileReadString2(hChildErrPipeRd, Buffer, CharsRead) Then
                    
                    If CharsRead > 0 Then
                        If mErrAsOut Then
                            BufferOut.Append Buffer
                '            BufferOut.Append Left$(Buffer, CharsRead)
                            RaiseEvent DataArrival(BufferOut.Length)
                        Else
                            BufferErr.Append Buffer
                            'BufferErr.Append Left$(Buffer, CharsRead)
                            RaiseEvent ErrDataArrival(BufferErr.Length)
                        End If
                    Else
                        RaiseEvent ErrEOF(SPEOF_NORMAL)
                        ClosePipeErr
                    End If
                Else
                    ErrNum = Err.LastDllError
                    If ErrNum = ERROR_BROKEN_PIPE Then
                        If Not mErrAsOut Then RaiseEvent ErrEOF(SPEOF_BROKEN_PIPE)
                        ClosePipeErr
                    Else
                        RaiseEvent Error(ErrNum, "ShellPipe.ReadData.ReadFile", Cancel)
                        If Not Cancel Then
                            Err.Raise ErrNum, Typename(Me), "ReadData ReadFile error"
                        End If
                    End If
                End If
            End If
        Else
            ErrNum = Err.LastDllError
            Select Case ErrNum
                Case ERROR_BROKEN_PIPE
                    If Not mErrAsOut Then RaiseEvent ErrEOF(SPEOF_BROKEN_PIPE)
                    ClosePipeErr
                    
                Case ERROR_ACCESS_DENIED, ERROR_INVALID_HANDLE
                    'Ignore as "no input."
                    
                Case Else
                    RaiseEvent Error(ErrNum, "ShellPipe.ReadData.PeekNamedPipe", Cancel)
                    If Not Cancel Then
                        Err.Raise ErrNum, Typename(Me), "ReadData PeeknamedPipe error"
                    End If
            End Select
        End If
    End If
End Sub

Private Sub WriteData()
    Dim Buffer As String
    Dim CharsWritten As Long
    Dim ErrNum As Long
    Dim Cancel As Boolean
    
    If PipeOpenIn Then
        If BufferIn.Length > 0 Then
            BufferIn.PeekBuffer Buffer
            If WriteChars2File(hChildInPipeWr, Buffer, Len(Buffer), CharsWritten, 0&) <> WIN32FALSE Then
             
            'If WriteFile(hChildInPipeWr, ByVal Buffer, Len(Buffer), CharsWritten, 0&) <> WIN32FALSE Then
                BufferIn.DeleteData CharsWritten
            Else
                ErrNum = Err.LastDllError
                RaiseEvent Error(ErrNum, "ShellPipe.WriteData.WriteFile", Cancel)
                If Not Cancel Then
                    Err.Raise ErrNum, Typename(Me), "WriteData WriteFile error"
                End If
            End If
        End If
    Else
        BufferIn.Clear
    End If
End Sub
Private Function WriteChars2File(ByVal hFile As Long, _
     lpBuf As String, _
     ByVal cToWrite As Long, _
     ByRef cWritten As Long, _
     ByVal lpOverlapped As Long) As Long
     Dim Buf1() As Byte, bytes As Long, stb As Long
     
     If LenB(lpBuf) = 0 Or cToWrite = 0 Then WriteChars2File = -1: Exit Function

     Select Case mCp
     Case -1
        Buf1 = StrConv(lpBuf, vbFromUnicode, 1033)
        bytes = UBound(Buf1) + 1
        stb = bytes
        WriteChars2File = API_Try_WriteFile(hFile, bytes, Buf1())
        cWritten = cToWrite
     Case 0
        Buf1 = Conv2CodePage(lpBuf, DefCodePage)
        bytes = UBound(Buf1) + 1
        WriteChars2File = API_Try_WriteFile(hFile, bytes, Buf1())
        cWritten = bytes
     Case 1
        Buf1 = Conv2CodePage(lpBuf, CLng(mLCid))
        bytes = UBound(Buf1) + 1
        stb = bytes
        WriteChars2File = API_Try_WriteFile(hFile, bytes, Buf1())
        If bytes = 0 Then
           cWritten = 0
        ElseIf stb > bytes Then
           cWritten = CharUtf16(Buf1(), bytes, CLng(mLCid))
        Else
           cWritten = cToWrite
        End If
     Case 65001
     
        
        Buf1 = Conv2CodePage(lpBuf, 65001) ' Utf16toUtf8(lpBuf)
        bytes = UBound(Buf1) + 1
        
        stb = bytes
        WriteChars2File = API_Try_WriteFile(hFile, bytes, Buf1())
        If bytes = 0 Then
           cWritten = 0
        ElseIf stb > bytes Then
           cWritten = CharUtf16(Buf1(), bytes, 65001)
        Else
           cWritten = cToWrite
        End If
     
     Case Else
        Buf1 = Conv2CodePage(lpBuf, CLng(mCp))
        bytes = UBound(Buf1) + 1
        stb = bytes
        WriteChars2File = API_Try_WriteFile(hFile, bytes, Buf1())
        If bytes = 0 Then
           cWritten = 0
        ElseIf stb > bytes Then
           cWritten = CharUtf16(Buf1(), bytes, CLng(mCp))
        Else
           cWritten = cToWrite
        End If
     
     End Select
     
End Function
Private Function FileReadString2(FileH As Long, r$, bytes As Long) As Boolean
    Dim Buf1() As Byte, dummy As Long, d$
    If bytes <= 0 Then r$ = vbNullString: Exit Function
    If FileH = 0 Then Exit Function
    ReDim Buf1(0 To bytes - 1)
    'r$ = Buf1()
    dummy = API_Try_ReadFile(FileH, bytes, Buf1())
    If bytes > 0 Then
    FileReadString2 = True
    ReDim Preserve Buf1(0 To bytes - 1)
    r$ = Buf1()
    CopyMemory ByVal StrPtr(r$), Buf1(0), bytes
    Select Case mCp
        Case -1
        Case 0
            r$ = StrConv(r$, vbUnicode)
        Case 1
            r$ = OEMdecode(r$, mLCid)
        Case 65001
            r$ = utf8decode(r$)
        Case Else
        On Error Resume Next
            d$ = StrConv(r$, vbUnicode, mLCid)
            If Err Then
                MyEr "can't use the selected cp", "δεν μπορώ χρησιμοποιήσω την επιλεγμένη κωδικοποίηση"
                r$ = StrConv(r$, vbUnicode)
            Else
                SwapStrings d$, r$
            End If
        End Select
    
    Else
        r$ = vbNullString
    End If
End Function

Private Sub Class_Initialize()
    blnFinishedChild = True
    mCp = DefCodePage
    mLCid = LCID_DEF
    Set BufferOut = New SPBuffer
    Set BufferErr = New SPBuffer
    Set BufferIn = New SPBuffer
    ErrAsOut = True
    PollInterval = 50
    WaitForIdle = 200
End Sub


Private Sub Class_Terminate()
FinishChild
End Sub

Public Property Get ChCp() As Variant
ChCp = mCp
End Property

Public Property Let ChCp(ByVal vNewValue As Variant)
If vNewValue = -1 Then
mCp = -1
mLCid = LCIDFromCodePage(CLng(mCp))
    
Else
mCp = vNewValue
If mCp < 10000 And mCp <> 0 Then
    mLCid = LCIDFromCodePage(CLng(mCp))
    If mLCid = 0 Then
    mLCid = mCp
    mCp = 1
    End If
End If
End If
End Property
Public Property Let ChCpFromLCID(ByVal vNewValue As Variant)
Dim oCp As Long
oCp = mCp
On Error Resume Next
If vNewValue = 0 Then
vNewValue = lcid
End If
mCp = CodePageFromLCID(CLng(vNewValue))
If Err Then
    Err.Clear
    mLCid = CLng(vNewValue)
    mCp = oCp
End If
End Property
