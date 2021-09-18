VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SerialPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type DCB
  DCBlength As Long
  BaudRate As Long
  fBitFields As Long
  wReserved As Integer
  XonLim As Integer
  XoffLim As Integer
  ByteSize As Byte
  parity As Byte
  StopBits As Byte
  XonChar As Byte
  XoffChar As Byte
  ErrorChar As Byte
  EofChar As Byte
  EvtChar As Byte
  wReserved1 As Integer
End Type

' The structure of the fBitFields field.
' FieldName             Bit #     Description
' -----------------     -----     ------------------------------
' fBinary                 1       Windows does not support nonbinary mode transfers, so this member must be =1.
' fParity                 2       If =1, parity checking is performed and errors are reported
' fOutxCtsFlow            3       If =1 and CTS is turned off, output is suspended until CTS is sent again.
' fOutxDsrFlow            4       If =1 and DSR is turned off, output is suspended until DSR is sent again.
' fDtrControl             5,6     DTR flow control (2 bits)
' fDsrSensitivity         7       The driver ignores any bytes received, unless the DSR modem input line is high.
' fTXContinueOnXoff       8       XOFF continues Tx
' fOutX                   9       If =1, TX stops when the XoffChar character is received and starts again when the XonChar character is received.
' fInX                   10       Indicates whether XON/XOFF flow control is used during reception.
' fErrorChar             11       Indicates whether bytes received with parity errors are replaced with the character specified by the ErrorChar.
' fNull                  12       If =1, null bytes are discarded when received.
' fRtsControl            13,14    RTS flow control (2 bits)
' fAbortOnError          15       If =1, the driver terminates all I/O operations with an error status if an error occurs.
' fDummy2                16       reserved

'---------fBitFields-------------
Const F_BINARY = 1
Const F_PARITY = 2
Const F_OUTX_CTS_FLOW = 4
Const F_OUTX_DSR_FLOW = 8

'DTR Control Flow Values.
Const F_DTR_CONTROL_ENABLE = &H10
Const F_DTR_CONTROL_HANDSHAKE = &H20

Const F_DSR_SENSITIVITY = &H40
Const F_TX_CONTINUE_ON_XOFF = &H80
Const F_OUT_X = &H100
Const F_IN_X = &H200
Const F_ERROR_CHAR = &H400
Const F_NULL = &H800

'RTS Control Flow Values
Const F_RTS_CONTROL_ENABLE = &H1000
Const F_RTS_CONTROL_HANDSHAKE = &H2000
Const F_RTS_CONTROL_TOGGLE = &H3000

Const F_ABORT_ON_ERROR = &H4000

'---------Parity flags--------
Const EVENPARITY = 2
Const MARKPARITY = 3
Const NOPARITY = 0
Const ODDPARITY = 1
Const SPACEPARITY = 4

'---------StopBits-----------
Const ONESTOPBIT = 0
Const ONE5STOPBITS = 1
Const TWOSTOPBITS = 2

'-----------------------------------------------------------------------------------------------
Private Type COMMTIMEOUTS
  ReadIntervalTimeout As Long
  ReadTotalTimeoutMultiplier As Long
  ReadTotalTimeoutConstant As Long
  WriteTotalTimeoutMultiplier As Long
  WriteTotalTimeoutConstant As Long
End Type
'-----------------------------------------------------------------------------------------------

'Constants for the dwDesiredAccess parameter of the CreateFile() function
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000

'Constants for the dwShareMode parameter of the CreateFile() function
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2

'Constants for the dwCreationDisposition parameter of the CreateFile() function
Const CREATE_NEW = 1
Const CREATE_ALWAYS = 2
Const OPEN_EXISTING = 3

'Constants for the dwFlagsAndAttributes parameter of the CreateFile() function
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_FLAG_OVERLAPPED = &H40000000

'-----------------------------------------------------------------------------------------------
'Error codes reported by the CreateFile().
'More error codes with descriptions are available at MSDN
Const ERROR_FILE_NOT_FOUND = 2
Const ERROR_ACCESS_DENIED = 5
Const ERROR_INVALID_HANDLE = 6
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
' added by George from https://github.com/grricks/vb6SerialAPI/blob/master/clsVB6serialAPI.cls
Private Declare Function EscapeCommFunction Lib "kernel32" _
    (ByVal nCid As Long, ByVal nFunc As Long) As Long
Private Declare Function PurgeComm Lib "kernel32" _
    (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetCommModemStatus Lib "kernel32" _
    (ByVal hFile As Long, lpModemStat As Long) As Long
' end addition

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetCommState Lib "kernel32" (ByVal hFile As Long, lpDCB As DCB) As Long
Private Declare Function GetCommState Lib "kernel32" (ByVal hFile As Long, lpDCB As DCB) As Long

Private Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, _
        lpCommTimeouts As COMMTIMEOUTS) As Long

Private Declare Function GetCommTimeouts Lib "kernel32" (ByVal hFile As Long, _
        lpCommTimeouts As COMMTIMEOUTS) As Long

Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, _
         ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) _
         As Long

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, _
         ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, _
         ByVal lpOverlapped As Long) As Long
Private h As Long, port$, BaudRate, mportnumber As Long
Private openserial As Boolean
Private timeouts As COMMTIMEOUTS
Dim BufferOut As MemBlock, bOutOffset As Long
Dim BufferIn As MemBlock, bInOffset As Long, bInBase As Long
Dim mlock As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Event DataArrival(ByVal CharsTotal As Long)
Public Event ReadyToSend()
Public Event Finished()


Public Sub OpenPort(ByVal portnumber As Long, Optional ByVal rate As Long = 9600, Optional ByVal parity As Long = 0, Optional ByVal StopBits As Long = 0, Optional ByVal Flags As Long = 0)
    Dim rc As Long
    
    If portnumber >= 1 And portnumber <= 9 Then
        port$ = "\\.\COM" + Trim$(Str$(portnumber))
    ElseIf portnumber > 9 And portnumber <= 24 Then
        port$ = "\\\\.\\COM" + Trim$(Str$(portnumber))
    Else
        Exit Sub
    End If
    
    h = CreateFile(StrPtr(port$), GENERIC_READ + GENERIC_WRITE, ByVal 0, ByVal 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)

    If h = -1 Then
        rc = Err.LastDllError
        Select Case rc 'Two typical error codes when trying to open a serial port:
         Case ERROR_ACCESS_DENIED  ' - The serial port opened by another application
           MyEr "The serial port COM" + Str$(portnumber) + " is used by another program", "Η σειριακή πόρτα COM" + Str$(portnumber) + " χρησιμοποιείται από άλλο πρόγραμμα"
         Case ERROR_FILE_NOT_FOUND ' - The serial port does not exist, check the port name specified in the CreateFile()
           MyEr "The serial port COM" + Str$(portnumber) + " does not exist", "Η σειριακή πόρτα COM" + Str$(portnumber) + " δεν υπάρχει"
         Case Else
           MyEr "CreateFile failed, the error code is " & Str(rc), "Η σύνδεση απέτυχε με κωδικό λάθους: " & Str(rc)
        End Select
        Exit Sub
    End If
    openserial = True
    Dim d As DCB 'The DCB structure and the SetCommState() function allow to set the baud rate and the byte size of the serial port.
    rc = GetCommState(h, d)
    d.ByteSize = 8
    d.BaudRate = rate
    d.fBitFields = F_BINARY Or Flags 'Windows does not support non-binary data transfers so the flag must always be set in the DCB structure.
    
    'Another example how to set some flags in the DCB.
    'd.fBitFields = F_BINARY Or F_PARITY Or F_RTS_CONTROL_ENABLE
    
    d.StopBits = ONESTOPBIT
    d.parity = NOPARITY
    rc = SetCommState(h, d)
    If rc = 0 Then
      Err1 Err.LastDllError, "SetCommState"
      GoTo close_and_exit
    End If
    SetReadTimeOuts
    mportnumber = portnumber
     'Because we don't want communication timeouts to hang the VB code,
    Exit Sub
close_and_exit:
    CloseSerial
End Sub
Sub SetReadTimeOuts(Optional ByVal IntervalTimeout As Long = 3, Optional ByVal TotalTimeoutConstant As Long = 20, Optional ByVal TotalTimeoutMultiplier As Long = 0)
    Dim rc As Long
        If openserial Then
        rc = GetCommTimeouts(h, timeouts)  'we need to specify the maximum time Windows will wait for incoming data
        timeouts.ReadIntervalTimeout = IntervalTimeout  'The max. time in milliseconds between arrival of any two bytes
        timeouts.ReadTotalTimeoutConstant = TotalTimeoutConstant 'The max. time the ReadFile() function will wait for data.
        timeouts.ReadTotalTimeoutMultiplier = TotalTimeoutMultiplier
        rc = SetCommTimeouts(h, timeouts)
        If rc = 0 Then Err1 Err.LastDllError, "SetCommTimeOuts": GoTo close_and_exit
    End If
    Exit Sub
close_and_exit:
    CloseSerial
End Sub
Sub ProcessLoop(Optional chunklen As Long = 10)
    Dim rd As Long, rc As Long, wr As Long
    Dim bRead(1 To 10) As Byte, s$
    If mlock Then Err.Raise "Locked buffers in serial port " & mportnumber
    mlock = True
    If openserial Then
        If bOutOffset + chunklen >= BufferOut.SizeByte Then
            BufferOut.ResizeItems BufferOut.SizeByte + chunklen * 10
        End If
      '  GoTo 1000
        rc = ReadFile(h, ByVal BufferOut.GetBytePtr(bOutOffset), chunklen, rd, 0)    'rd indicates how many bytes received from the port.
        If rc = 0 Then
          Err1 Err.LastDllError, "ReadFile"
          GoTo close_and_exit
        End If
        If rd > 0 Then
            bOutOffset = bOutOffset + rd
            mlock = False
            RaiseEvent DataArrival(bOutOffset)
            mlock = True
        End If
        If Not (bInOffset > 0 And bInOffset - bInBase > 0) Then
            mlock = False
            RaiseEvent ReadyToSend
            mlock = True
        End If
        If bInOffset > 0 And bInOffset - bInBase > 0 Then
            rc = WriteFile(h, ByVal BufferIn.GetBytePtr(bInBase), bInOffset - bInBase, wr, 0)      'wr indicates how many bytes went to the port.
            If rc = 0 Then
                Err1 Err.LastDllError, "WriteFile"
                GoTo close_and_exit
            Else
                bInBase = bInBase + wr
            End If
        End If
    End If
    mlock = False
    Exit Sub
close_and_exit:
    CloseSerial
    mlock = False
End Sub

Public Sub CloseSerial()
    Dim rc As Long
    If openserial Then rc = CloseHandle(h): openserial = False
End Sub
Property Get MayReadWrite() As Boolean
    MayReadWrite = Not mlock
End Property
Property Get BufferOutBytes() As Long
    BufferOutBytes = BufferOut.SizeByte
End Property
Sub WriteBuffer(RHS As Object)
Dim a As MemBlock, aOffset As Long, aHandler As mHandler
If Typename(RHS) <> "mHandler" Then GoTo 100
Set aHandler = RHS
If aHandler.t1 <> 2 Then GoTo 100
If mlock Then Err.Raise "Locked input buffer in serial port " & mportnumber
mlock = True
Set a = aHandler.objref
If bInOffset - bInBase = 0 Then
    Set BufferIn = a
    bInOffset = a.SizeByte
    bInBase = 0
ElseIf bInBase = 0 Then
    BufferIn.ResizeItems bInOffset + a.SizeByte
    CopyMemory BufferIn.GetPtr(bInOffset), a.GetPtr(0), a.SizeByte
    bInOffset = BufferIn.SizeByte
Else
    CopyMemory BufferIn.GetPtr(0), a.GetPtr(bInBase), bInOffset - bInBase
    bInOffset = bInOffset - bInBase
    BufferIn.ResizeItems bInOffset + a.SizeByte
    CopyMemory BufferIn.GetPtr(bInOffset), a.GetPtr(0), a.SizeByte
    bInOffset = BufferIn.SizeByte
    Set BufferIn = a
    bInBase = 0
End If
mlock = False
Exit Sub
100 Err.Raise "Wrong input buffer in serial port " & mportnumber
End Sub
Sub WriteString(RHS As String)
    Dim a As New MemBlock, b As mHandler, c As Object
    a.Construct 1, LenB(RHS)
    CopyMemory ByVal a.GetPtr(0), ByVal StrPtr(RHS), LenB(RHS)
    Set b = New mHandler
    b.t1 = 2
    Set b.objref = a
    Set c = b
    WriteBuffer c
End Sub
Function ReadBuffer() As Object
    If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
    If BufferOut.SizeByte = 0 Then
        Err.Raise "Empty buffer in serial port " & mportnumber
    Else
        Dim a As MemBlock
        Set a = BufferOut
        Set BufferOut = New MemBlock
        BufferOut.Construct 1, 2048
        a.ResizeItems bOutOffset
        bOutOffset = 0
        Dim buf As New mHandler
        buf.t1 = 2
        Set buf.objref = a
        Set ReadBuffer = buf
    End If
End Function
Sub PortNotOpened()
    MyEr "Serial port is closed", "Η Σειριακή είναι κλειστή"
End Sub
Sub LINE_BREAK(b As Boolean)
    Dim rc As Long
    If openserial Then
        If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
        Const SETBREAK = 8&
        Const CLRBREAK = 9&
        If b Then
            rc = EscapeCommFunction(h, SETBREAK)
        Else
            rc = EscapeCommFunction(h, CLRBREAK)
        End If
        If rc = 0 Then Err2 Err.LastDllError
    Else
        PortNotOpened
    End If
End Sub
Sub LINE_DTR(b As Boolean)
    Dim rc As Long
    If openserial Then
        If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
        Const SETDTR = 5&
        Const CLRDTR = 6&
        If b Then
            rc = EscapeCommFunction(h, SETDTR)
        Else
            rc = EscapeCommFunction(h, CLRDTR)
        End If
        If rc = 0 Then Err2 Err.LastDllError
    Else
        PortNotOpened
    End If
End Sub
Sub LINE_RTS(b As Boolean)
    Dim rc As Long
    If openserial Then
        If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
        Const SETRTS = 3&
        Const CLRRTS = 4&
        If b Then
            rc = EscapeCommFunction(h, SETRTS)
        Else
            rc = EscapeCommFunction(h, CLRRTS)
        End If
        If rc = 0 Then Err2 Err.LastDllError
    Else
        PortNotOpened
    End If
End Sub
Function Inputs() As Long
    Dim rc As Long, R As Long
    rc = GetCommModemStatus(h, R)
    If rc = 0 Then Err1 Err.LastDllError, "GetCommModemStatus": R = -1
    Inputs = R
End Function
Function CD(Optional ByVal R As Long = -1) As Variant
    Dim rc As Long
    Const LINE_CD = &H80&
    If openserial Then
        If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
        If R = -1 Then rc = GetCommModemStatus(h, R): If rc = 0 Then Err1 Err.LastDllError, "GetCommModemStatus": R = -1
        CD = (R And LINE_CD) <> 0
    Else
        PortNotOpened
    End If
End Function
Function CTS(Optional ByVal R As Long = -1) As Variant
    Dim rc As Long
    Const LINE_CTS = &H10&
    If openserial Then
        If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
           If R = -1 Then rc = GetCommModemStatus(h, R): If rc = 0 Then Err1 Err.LastDllError, "GetCommModemStatus": R = -1
        CTS = (R And LINE_CTS) <> 0
    Else
        PortNotOpened
    End If
End Function
Function DSR(Optional ByVal R As Long = -1) As Variant
    Dim rc As Long
    Const LINE_DSR = &H20&
    If openserial Then
        If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
        If R = -1 Then rc = GetCommModemStatus(h, R): If rc = 0 Then Err1 Err.LastDllError, "GetCommModemStatus": R = -1
        DSR = (R And LINE_DSR) <> 0
    Else
        PortNotOpened
    End If
End Function
Function RING(Optional ByVal R As Long = -1) As Variant
    Dim rc As Long
    Const LINE_RING = &H40&
    If openserial Then
        If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
        If R = -1 Then rc = GetCommModemStatus(h, R): If rc = 0 Then Err1 Err.LastDllError, "GetCommModemStatus": R = -1
        RING = (R And LINE_RING) <> 0
    Else
        PortNotOpened
    End If
End Function
Function RLSD(Optional ByVal R As Long = -1) As Variant
    Dim rc As Long
    Const LINE_RLSD = &H80&
    If openserial Then
        If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
        If R = -1 Then rc = GetCommModemStatus(h, R): If rc = 0 Then Err1 Err.LastDllError, "GetCommModemStatus": R = -1
        RLSD = (R And LINE_RLSD) <> 0
    Else
        PortNotOpened
    End If
End Function
Sub FlushPort()
    Dim rd As Long
    If mlock Then Err.Raise "Locked buffer in serial port " & mportnumber
    mlock = True
    If Not openserial Then
        rd = 1
    Else
        rd = PurgeComm(h, &HF)
    End If
    If rd = 0 Then
        MyEr "Can't flush the port's buffers", "Δεν μπορώ να καθαρίσω τις μνήμες της πόρτας)"
    Else
        Set BufferOut = New MemBlock
        Set BufferIn = New MemBlock
        BufferOut.Construct 1, 2048
        BufferIn.Construct 1, 2048
        bInOffset = 0
        bInBase = 0
        bOutOffset = 0
    End If
    mlock = False
End Sub
Private Sub Class_Initialize()
Set BufferOut = New MemBlock
Set BufferIn = New MemBlock
BufferOut.Construct 1, 2048
BufferIn.Construct 1, 2048
End Sub

Private Sub Class_Terminate()
CloseSerial
End Sub
Private Sub Err2(ByVal rc As Long)
    Err1 rc, "EscapeCommFunction"
End Sub
Private Sub Err1(ByVal rc As Long, mess As String)
    MyEr mess & " failed, the error code is " & Str(rc), "η " & mess & " απέτυχε με κωδικό λάθους: " & Str(rc)
End Sub

