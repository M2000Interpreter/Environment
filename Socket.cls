VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mySocket As cAsyncSocket
Attribute mySocket.VB_VarHelpID = -1
Event OnResolve(IpAddress As String)
Event OnAccept()
Event OnClose()
Event OnConnect()
Event OnReceive()
Event OnSend()
Event OnError(ByVal ErrorCode As Long, ByVal EventMask As Long)
Event OnMessagePending(Handled As Boolean)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private m_hSocket               As Long
Private Const STR_CHR1              As String = "" '--- CHAR(1)
Const ucsScpUtf8 = 65001
Private SyncEnabled As Boolean, ReceiveOk As Boolean, SendOk As Boolean, OnConnect As Boolean
Private CloseOk As Boolean, AcceptOk As Boolean, ErrorOk As Boolean
Public NoError As Boolean
Public Sub CloseConnection()
If m_hSocket <> 0 Then
mySocket.Close_
m_hSocket = 0
Set mySocket = Nothing
End If
End Sub

Private Sub Class_Terminate()
CloseConnection
End Sub

Private Sub mySocket_OnAccept()
    If SyncEnabled Then
        AcceptOk = True
    Else
        RaiseEvent OnAccept
    End If
End Sub

Private Sub mySocket_OnClose()
    If SyncEnabled Then
        CloseOk = True
        NoError = False
    Else
        RaiseEvent OnClose
    End If
End Sub

Private Sub mySocket_OnConnect()
    If SyncEnabled Then
        OnConnect = True
    Else
        RaiseEvent OnConnect
    End If
End Sub

Private Sub mySocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As Long)
    If SyncEnabled Then
    If ErrorCode = 10035 Then
    SendOk = False
    Exit Sub
    End If
    If ErrorCode = 10053 Then Exit Sub
    ErrorOk = True
    If NoError Then Exit Sub
    MyEr "Socket Error: " & mySocket.GetErrorDescription(ErrorCode), "Σταθμός Λάθος: " & mySocket.GetErrorDescription(ErrorCode)
    Else
    RaiseEvent OnError(ErrorCode, EventMask)
    End If
End Sub

Private Sub mySocket_OnMessagePending(Handled As Boolean)
    If SyncEnabled Then
    
    Else
        RaiseEvent OnMessagePending(Handled)
    End If
End Sub

Private Sub mySocket_OnReceive()
    If SyncEnabled Then
    ReceiveOk = True
    Else
    RaiseEvent OnReceive
    End If
End Sub

Private Sub mySocket_OnResolve(IpAddress As String)
    If SyncEnabled Then
    
    Else
        RaiseEvent OnResolve(IpAddress)
    End If
End Sub

Private Sub mySocket_OnSend()
    If SyncEnabled Then
    SendOk = True
    Else
        RaiseEvent OnSend
    End If
End Sub
Property Get SocketHandle() As Variant
    SocketHandle = CVar(mySocket.SocketHandle)
End Property

Property Get LastError() As Variant
    LastError = CVar(mySocket.LastError)
End Property

Property Get SockOpt(ByVal OptionName As Variant, Optional ByVal Level As Integer) As Variant
    If m_hSocket = 0 Then
        socket_not_init
    Else
        If IsMissing(Level) Then Level = &HFFFF&
        SockOpt = CVar(mySocket.SockOpt(CLng(OptionName), CLng(Level)))
    End If
End Property

Property Let SockOpt(ByVal OptionName As Variant, Optional ByVal Level As Integer, ByVal Value As Variant)
    If m_hSocket = 0 Then
        socket_not_init
    Else
        If IsMissing(Level) Then Level = &HFFFF&
        mySocket.SockOpt(CLng(OptionName), CLng(Level)) = CLng(Value)
    End If
End Property

Property Get AvailableBytes() As Variant
    If m_hSocket = 0 Then
        socket_not_init
    Else
        AvailableBytes = CVar(mySocket.AvailableBytes)
    End If
End Property

Property Get HasPendingEvent() As Variant
    If m_hSocket = 0 Then
        socket_not_init
    Else
        HasPendingEvent = CVar(mySocket.HasPendingEvent)
    End If
End Property

Property Get HasPendingResolve() As Variant
    If m_hSocket = 0 Then
        socket_not_init
    Else
        HasPendingResolve = mySocket.HasPendingResolve
    End If
End Property
Friend Property Get OwnAsyncSocket() As cAsyncSocket
    Set OwnAsyncSocket = mySocket
End Property
Friend Property Set OwnAsyncSocket(RHS As cAsyncSocket)
    m_hSocket = mySocket.SocketHandle
End Property
' Methods
Private Sub socket_not_init()
MyEr "Socket not created", "Ο σταθμός δεν έχει δημιουργηθεί"
End Sub
Private Sub connection_not_sync()
MyEr "Connection isn't synchronus", "Η σύνδεση δεν είναι σύγχρονη"
End Sub

Public Function Create( _
            Optional ByVal SocketPort As Variant, _
            Optional ByVal SocketType As Variant, _
            Optional ByVal EventMask As Variant, _
            Optional SocketAddress As Variant) As Variant
        If IsMissing(SocketPort) Then SocketPort = 0
        If IsMissing(SocketType) Then SocketType = 0
        If IsMissing(EventMask) Then EventMask = 0
        Dim iSocketAddress As String
        If IsMissing(SocketAddress) Then iSocketAddress = vbNullString Else iSocketAddress = CVar(SocketAddress)
        If m_hSocket = 0 Then Set mySocket = New cAsyncSocket: m_hSocket = mySocket.SocketHandle
        SyncEnabled = False
        Create = mySocket.Create(CLng(SocketPort), CLng(SocketType), CLng(EventMask), iSocketAddress)
        SocketAddress = iSocketAddress
End Function
Public Function GetPeerName(PeerAddress As Variant, PeerPort As Variant) As Variant
        If m_hSocket = 0 Then socket_not_init: Exit Function
        Dim iPeerAddress As String, iPeerPort As Long
        GetPeerName = mySocket.GetPeerName(iPeerAddress, iPeerPort)
        PeerAddress = iPeerAddress
        PeerPort = iPeerPort
End Function
Public Function GetSockName(SocketAddress As Variant, SocketPort As Variant) As Variant
        If m_hSocket = 0 Then socket_not_init: Exit Function
        Dim iSocketAddress As String, iSocketPort As Long
        GetSockName = mySocket.GetSockName(iSocketAddress, iSocketPort)
        SocketAddress = iSocketAddress
        SocketPort = iSocketPort
End Function
Public Function GetLocalHost(HostName As Variant, HostAddress As Variant) As Variant
        If m_hSocket = 0 Then socket_not_init: Exit Function
        Dim iHostName As String, iHostAddress As String
        GetLocalHost = mySocket.GetLocalHost(iHostName, iHostAddress)
        HostName = iHostName
        HostAddress = iHostAddress
End Function
Public Function Accept(ConnectedSocket As Object, Optional SocketAddress As Variant, Optional SocketPort As Variant) As Variant
'' ConnectedSocket
    Dim iSocketAddress As String, iSocketPort As Long
    Dim iConnectedSocket As Socket, cSocket As cAsyncSocket
    If Typename(ConnectedSocket) <> "Socket" Then
    MyEr "Not a Socket Object", "Δεν βρήκα αντικείμενο ΣΤΑΘΜΟΣ"
    Exit Function
    End If
        Set iConnectedSocket = ConnectedSocket
        If iConnectedSocket.OwnAsyncSocket Is Nothing Then
            Set cSocket = New cAsyncSocket
            Set iConnectedSocket.OwnAsyncSocket = cSocket
            SyncEnabled = False
            Accept = mySocket.Accept(cSocket, iSocketAddress, iSocketPort)
        End If
        SocketAddress = iSocketAddress
        SocketPort = iSocketPort
    
End Function
Public Function Bind(Optional SocketAddress As Variant, Optional ByVal SocketPort As Variant) As Variant
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim iSocketAddress As String, iSocketPort As Long
    If Not IsMissing(SocketPort) Then iSocketPort = SocketPort
        Bind = mySocket.Bind(iSocketAddress, iSocketPort)
        SocketAddress = iSocketAddress
    
End Function


Public Function Connect(HostAddress As Variant, ByVal HostPort As Variant) As Variant
    Dim iHostAddress As String, iHostPort As Long
    iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If m_hSocket = 0 Then Set mySocket = New cAsyncSocket: m_hSocket = mySocket.SocketHandle
    SyncEnabled = False
    Connect = mySocket.Connect(iHostAddress, iHostPort)
    HostAddress = iHostAddress
End Function
Public Function IOCtl(ByVal Command As Variant, Argument As Variant) As Boolean
    Dim iCommmand As Long, iArgument As Long
    If m_hSocket = 0 Then socket_not_init: Exit Function
    iCommmand = Command
    iArgument = iArgument
    If m_hSocket = 0 Then Set mySocket = New cAsyncSocket: m_hSocket = mySocket.SocketHandle
    IOCtl = mySocket.IOCtl(iCommmand, iArgument)
    Argument = iArgument
End Function
Public Function Listen(Optional ByVal ConnectionBacklog As Variant) As Variant
    Dim iConnectionBacklog As Long
    If m_hSocket = 0 Then socket_not_init: Exit Function
    If IsMissing(ConnectionBacklog) Then iConnectionBacklog = 5 Else iConnectionBacklog = ConnectionBacklog
    SyncEnabled = False
    Listen = mySocket.Listen(iConnectionBacklog)
End Function
Public Function ReceiveText( _
            Optional HostAddress, _
            Optional HostPort, _
            Optional ByVal CodePage) As Variant
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim iHostAddress As String, iHostPort As Long, iCodePage As Long
    If IsMissing(HostAddress) Then iHostAddress = STR_CHR1 Else iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If IsMissing(CodePage) Then iCodePage = ucsScpUtf8 Else iCodePage = CodePage
    SyncEnabled = False
    ReceiveText = mySocket.ReceiveText(iHostAddress, iHostPort, iCodePage)
    HostAddress = iHostAddress
    HostPort = iHostPort
End Function
' for buffers change to pass buffer object
' BufPtr from 0 as a multiple of basic item length
Public Function Receive( _
            buf, _
            ByVal BufPtr, _
            ByVal BufLen, _
            Optional HostAddress, _
            Optional HostPort, _
            Optional ByVal Flags) As Variant
    Receive = 0&
    If Not MyIsObject(buf) Then
100
        MyEr "need a Buffer", "χρειάζομαι μια διάρθρωση": Exit Function
    ElseIf Not TypeOf buf Is MemBlock Then
        GoTo 100
    End If
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim b As MemBlock
    Set b = buf
    Dim iBufPtr As Long, iBufLen As Long, iHostAddress As String, iHostPort As Long, iFlags As Long
    iBufPtr = b.GetPtr(BufPtr)
    If iBufPtr = 0 Then
        MyEr "offest is not in Buffer", "η αρχική διεύθυνση είναι εκτός διάρθρωσης": Exit Function
    End If
    If Not b.ValidArea2(CDbl(BufPtr), CDbl(BufLen)) Then
        MyEr "need a bigger Buffer", "χρειάζομαι μια μεγαλύτερη διάρθρωση": Exit Function
    End If
    iBufLen = BufLen
    If IsMissing(HostAddress) Then iHostAddress = STR_CHR1 Else iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If Not IsMissing(Flags) Then iFlags = Flags
    SyncEnabled = False
    Receive = mySocket.Receive(iBufPtr, iBufLen, iHostAddress, iHostPort, iFlags)
    HostAddress = iHostAddress
    HostPort = iHostPort

End Function
Public Function SendText( _
            Text, _
            Optional HostAddress, _
            Optional ByVal HostPort, _
            Optional ByVal CodePage As Long) As Variant
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim iText As String, iHostAddress As String, iHostPort As Long, iCodePage As Long
    iText = Text
    iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If IsMissing(CodePage) Then iCodePage = ucsScpUtf8 Else iCodePage = CodePage
    SyncEnabled = False
    SendText = mySocket.SendText(iText, iHostAddress, iHostPort, iCodePage)
    Text = iText
    HostAddress = iHostAddress
End Function
'change to send from buffer
Public Function Send( _
            buf, _
            ByVal BufPtr, _
            ByVal BufLen, _
            Optional HostAddress, _
            Optional ByVal HostPort, _
            Optional ByVal Flags) As Variant
    Send = 0&
    If Not MyIsObject(buf) Then
100
        MyEr "need a Buffer", "χρειάζομαι μια διάρθρωση": Exit Function
    ElseIf Not TypeOf buf Is MemBlock Then
        GoTo 100
    End If
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim b As MemBlock
    Set b = buf
     Dim iBufPtr As Long, iBufLen As Long, iHostAddress As String, iHostPort As Long, iFlags As Long
    iBufPtr = b.GetPtr(BufPtr)
    If iBufPtr = 0 Then
        MyEr "offest is not in Buffer", "η αρχική διεύθυνση είναι εκτός διάρθρωσης": Exit Function
    End If
    If Not b.ValidArea2(CDbl(BufPtr), CDbl(BufLen)) Then
        MyEr "need a bigger Buffer", "χρειάζομαι μια μεγαλύτερη διάρθρωση": Exit Function
    End If
    iBufLen = BufLen
    If IsMissing(HostAddress) Then iHostAddress = STR_CHR1 Else iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If Not IsMissing(Flags) Then iFlags = Flags
    SyncEnabled = False
    Send = mySocket.Send(iBufPtr, iBufLen, iHostAddress, iHostPort, iFlags)
    HostAddress = iHostAddress
    
End Function
Public Function ShutDown(Optional ByVal How) As Variant
    Dim iHow As Long
    If m_hSocket = 0 Then socket_not_init: Exit Function
    If IsMissing(How) = 1 Then iHow = 1 Else iHow = How
    SyncEnabled = False
    ShutDown = mySocket.ShutDown(iHow)
End Function
Public Function GetErrorDescription(ByVal ErrorCode) As Variant
        If m_hSocket = 0 Then socket_not_init: Exit Function
        Dim iErrorCode As Long
        iErrorCode = ErrorCode
        SyncEnabled = False
        GetErrorDescription = mySocket.GetErrorDescription(iErrorCode)
End Function
'' Public Sub PostEvent(ByVal EventMask As UcsAsyncSocketEventMaskEnum)
'' sync part not yet

Public Function SyncConnect(HostAddress As Variant, ByVal HostPort As Variant, Optional ByVal Timeout As Variant) As Variant
Dim tm As clsProfiler
SyncEnabled = True
CloseOk = False
AcceptOk = False
ErrorOk = False
ReceiveOk = False
OnConnect = False
NoError = True
Dim iTimecount
Dim iHostAddress As String, iHostPort As Long, iTimeout As Long
    iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If m_hSocket = 0 Then Set mySocket = New cAsyncSocket: m_hSocket = mySocket.SocketHandle
    Dim ok As Boolean, escnow
    If IsMissing(Timeout) Then
    iTimeout = Znl(mySocket.SockOpt(&H1006), 5000)
    Else
    iTimeout = CLng(Timeout)
    End If
    If iTimeout > 0 Then
        Set tm = New clsProfiler
        tm.MARKONE
    End If
    ok = mySocket.Connect(iHostAddress, iHostPort)
    If ok Then
        Do
            ProcTask2 Basestack1
            If escok Then escnow = KeyPressed(&H1B)
        Loop Until OnConnect Or ErrorOk Or CloseOk Or escnow Or tm.MARKTWO > iTimeout
        SyncConnect = OnConnect
    Else
    SyncConnect = False
    End If
    
    HostAddress = iHostAddress


End Function
Friend Function SyncReadArray(Buffer() As Byte) As Boolean
    Const FUNC_NAME     As String = "ReadArray"
    Dim lAvailable      As Long
    Dim lBytes          As Long
    Dim lReceived       As Long
    
    On Error GoTo EH

        lAvailable = mySocket.AvailableBytes
        If lAvailable <= 0 Then
            lAvailable = mySocket.SockOpt(ucsSsoReceiveBuffer)
        End If
        If lAvailable <= 0 Then
            lAvailable = 4096
        End If
        ReDim Buffer(0 To lBytes + lAvailable - 1) As Byte
        Do
            If Not SyncReceive(VarPtr(Buffer(lBytes)), lAvailable, lReceived) Then
                GoTo QH
            End If
        Loop While lReceived = 0
        If lBytes + lReceived - 1 <> UBound(Buffer) Then
            ReDim Preserve Buffer(0 To lBytes + lReceived - 1) As Byte
        End If

    '--- succees
    SyncReadArray = True
QH:
EH:
End Function
Public Function SyncReadText(Optional ByVal CodePage As Long = ucsScpUtf8) As String
    If m_hSocket = 0 Then socket_not_init: Exit Function
    If Not SyncEnabled Then connection_not_sync: Exit Function
    Dim baBuffer()      As Byte
    Dim s$
    If SyncReadArray(baBuffer) Then
        s$ = mySocket.FromTextArray(baBuffer, CodePage)
        SyncReadText = s$
    End If
End Function
Public Function SyncReceiveText( _
            Optional ByVal NeedLen, _
            Optional ByVal Timeout, _
            Optional HostAddress, _
            Optional HostPort, _
            Optional ByVal CodePage) As Variant
    If m_hSocket = 0 Then socket_not_init: Exit Function
    If Not SyncEnabled Then connection_not_sync: Exit Function
    Dim tm As clsProfiler
    Dim iNeedLen As Long, iTimeout As Long
    If IsMissing(NeedLen) Then iNeedLen = 0 Else iNeedLen = NeedLen
    If IsMissing(iTimeout) Then
    iTimeout = Znl(mySocket.SockOpt(&H1006), 5000)
    Else
    iTimeout = CLng(Timeout)
    End If
    
    Dim iHostAddress As String, iHostPort As Long, iCodePage As Long
    If IsMissing(HostAddress) Then iHostAddress = STR_CHR1 Else iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If IsMissing(CodePage) Then iCodePage = ucsScpUtf8 Else iCodePage = CodePage
    ReceiveOk = False
    ErrorOk = False
    CloseOk = False
    SyncEnabled = True
    Dim escnow As Long, part$, Count As Long
    If iTimeout > 0 Then
        Set tm = New clsProfiler
        tm.MARKONE
    End If
    If NeedLen > 0 Then SyncReceiveText = space$(NeedLen): Count = 1
again:
        If Not ReceiveOk Then
         Do
            ProcTask2 Basestack1
            If escok Then escnow = KeyPressed(&H1B)
        Loop Until ReceiveOk Or ErrorOk Or CloseOk Or escnow Or tm.MARKTWO > iTimeout
        End If
        If ReceiveOk And Not (ErrorOk Or CloseOk Or escnow) Then
        ReceiveOk = False
        If NeedLen > 0 Then
             part$ = mySocket.ReceiveText(iHostAddress, iHostPort, iCodePage)
             If Len(part$) > 0 Then
                Mid$(SyncReceiveText, Count, Len(part$)) = part$
                Count = Count + Len(part$)
             End If
             If Count <= NeedLen Then GoTo again
        Else
        SyncReceiveText = mySocket.ReceiveText(iHostAddress, iHostPort, iCodePage)
        End If
        Else
        ReceiveOk = False
        SyncReceiveText = ""
        End If
    HostAddress = iHostAddress
    HostPort = iHostPort
End Function
Private Function SyncReceive( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Received As Long, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
            Dim tm As New clsProfiler, escnow As Boolean
            
            ErrorOk = False
            CloseOk = False
            If Timeout = 0 Then
                Timeout = Znl(mySocket.SockOpt(&H1006), 5000)
            End If
            tm.MARKONE
            Debug.Print "Receive now", ReceiveOk
again:
            If Not ReceiveOk Then
            Do
                ProcTask2 Basestack1
                If escok Then escnow = KeyPressed(&H1B)
            Loop Until ReceiveOk Or ErrorOk Or CloseOk Or escnow Or tm.MARKTWO > Timeout
            End If
             If ReceiveOk And Not (ErrorOk Or CloseOk Or escnow) Then  '
                ReceiveOk = False
                Received = mySocket.Receive(BufPtr, BufLen, HostAddress, HostPort, Flags)
                If mySocket.HasPendingEvent And Received = -1 Then ReceiveOk = False: GoTo again
                SyncReceive = True
            Else
                Debug.Print "Not Received", ErrorOk, CloseOk
                SyncReceive = False
            End If
           
End Function
Public Function SyncSendText( _
            Text, _
            Optional ByVal Timeout, _
            Optional HostAddress, _
            Optional ByVal HostPort, _
            Optional ByVal CodePage) As Variant
    Const FUNC_NAME     As String = "SendText"
    If m_hSocket = 0 Then socket_not_init: Exit Function
    If Not SyncEnabled Then connection_not_sync: Exit Function
    On Error GoTo EH
    Dim iText As String, iTimeout As Long, iHostAddress As String, iHostPort As Long, iCodePage As UcsAsyncSocketCodePageEnum
    iText = Text
    If Len(iText) = 0 Then
    Exit Function
    End If
    If Not IsMissing(Timeout) Then iTimeout = Timeout
    If IsMissing(CodePage) Then iCodePage = ucsScpUtf8 Else iCodePage = CLng(CodePage)
    SyncSendText = SyncSendArray(mySocket.ToTextArray(iText, iCodePage), iTimeout, iHostAddress, iHostPort)
    HostAddress = iHostAddress
    HostPort = iHostPort
    Exit Function
EH:
End Function
Private Function SyncSendArray( _
            Buffer() As Byte, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SyncSendArray"
    
    On Error GoTo EH
    If UBound(Buffer) < 0 Then
        SyncSendArray = True
    Else
        SyncSendArray = SyncSend(VarPtr(Buffer(0)), UBound(Buffer) + 1, Timeout, HostAddress, HostPort, Flags)
    End If
    Exit Function
EH:
End Function
Private Function SyncSend( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
            Dim tm As New clsProfiler, escnow As Boolean, lResult As Long
            ErrorOk = False
            CloseOk = False
            If Timeout = 0 Then
                Timeout = Znl(mySocket.SockOpt(&H1005), 5000)
            End If
            tm.MARKONE
            Debug.Print "Send now", SendOk
again:
            If Not SendOk Then
            Debug.Print "Enter Loop for waiting SendOk"
            Do
                ProcTask2 Basestack1
                If escok Then escnow = KeyPressed(&H1B)
            Loop Until SendOk Or ErrorOk Or CloseOk Or escnow Or tm.MARKTWO > Timeout
            End If
            If SendOk And Not (ErrorOk Or CloseOk Or escnow) Then
                Debug.Print "send start"
                lResult = mySocket.Send(BufPtr, BufLen, HostAddress, HostPort, Flags)
                If mySocket.HasPendingEvent And lResult = -1 Then SendOk = False: GoTo again
                If lResult = -1 Then GoTo failed
                
                If lResult < BufLen Then
                    BufPtr = UnsignedAdd(BufPtr, lResult)
                    BufLen = BufLen - lResult
                    ' tm.MARKONE ' reset timeout
                    GoTo again
                End If
                Debug.Print "Send " & BufLen & " bytes"
                SyncSend = True
            Else
failed:
                Debug.Print "Send Failed to send " & BufLen & " bytes"
                SyncSend = False
            End If
End Function
Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
Private Function Znl(ByVal lValue As Long, Optional IfEmptyLong As Variant = Null, Optional ByVal EmptyLong As Long = 0) As Variant
    Znl = IIf(lValue = EmptyLong, IfEmptyLong, lValue)
End Function


