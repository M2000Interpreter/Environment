VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHttpsRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const MODULE_NAME As String = "clsHttpsRequest"
Private WithEvents m_oSocket    As cTlsSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_sServerName           As String
Private m_oRootCa               As cTlsSocket
Attribute m_oRootCa.VB_VarHelpID = -1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type UcsParsedUrl
    Protocol        As String
    Host            As String
    Port            As Long
    Path            As String
    QueryString     As String
    Anchor          As String
    User            As String
    Pass            As String
End Type
Private uRemote As UcsParsedUrl, Zero As UcsParsedUrl
Private mHttpsRequestValue As String
Private vHeaders        As Variant
Private lHeaderLength   As Long
Private lContentLength  As Long
Private sEncoding       As String
Private sConnection     As String
Private sError As String
Event Start(ByVal remobtehost$, ByVal RemotePort As Long)
Event Cancel(ByRef CancelMe)
Event DownloadError(ByVal Number As Long, ByVal Description As String)
Event Complete()

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "HttpsRequest"
    Const HDR_CONTENT_LENGTH As String = "content-length:"
    Const HDR_TRANSFER_ENCODING As String = "transfer-encoding:"
    Const HDR_CONNECTION As String = "connection:"



End Sub
Public Function HttpsRequest(https, Optional ByVal aWait As Boolean = False) As Boolean
    Const FUNC_NAME     As String = "HttpsRequest"
    Const HDR_CONTENT_LENGTH As String = "content-length:"
    Const HDR_TRANSFER_ENCODING As String = "transfer-encoding:"
    Const HDR_CONNECTION As String = "connection:"
    Dim baRecv()        As Byte
    Dim sRequest        As String
    Dim bResult         As Boolean
    HeaderLength = 0
    lContentLength = 0
    sEncoding = 0

    sConnection = vbNullString
    Dim vElem           As Variant
    EmptyVariant VarPtr(vHeaders)
    uRemote = Zero
    sError = vbNullString
    If Not ParseUrl(Trim$(CStr(https)), uRemote, DefProtocol:="https") Then
        sError = "Parsing Url Error"
        Exit Function
    End If

    If Not m_oSocket Is Nothing Then
        If m_oSocket.IsClosed Then
            Set m_oSocket = Nothing
            m_oSocket.EnableM2000threads = aWait
        End If
    End If
    If m_sServerName <> uRemote.Host & ":" & uRemote.Port Or m_oSocket Is Nothing Then
        Set m_oSocket = New cTlsSocket
        m_oSocket.EnableM2000threads = aWait
        If LCase(uRemote.Protocol) = "http" Then
            If Not m_oSocket.SyncConnect(uRemote.Host, uRemote.Port, _
                    UseTls:=False, _
                    AlpnProtocols:="http/1.1") Then
                sError = m_oSocket.LastError.Description
                GoTo QH
            End If

        ElseIf Not m_oSocket.SyncConnect(uRemote.Host, uRemote.Port, _
                LocalFeatures:=IIf(pvIsKnownBadCertificate(uRemote.Host), ucsTlsIgnoreServerCertificateErrors, 0), _
                RootCa:=m_oRootCa, AlpnProtocols:="http/1.1") Then
            sError = m_oSocket.LastError.Description
            GoTo QH
        End If
        m_sServerName = uRemote.Host & ":" & uRemote.Port
    End If
    '--- send TLS application data and wait for reply
    sRequest = "GET " & uRemote.Path & uRemote.QueryString & " HTTP/1.1" & vbCrLf & _
               "Connection: keep-alive" & vbCrLf & _
               "Host: " & uRemote.Host & vbCrLf & vbCrLf
               
    If Not m_oSocket.SyncSendArray(StrConv(sRequest, vbFromUnicode)) Then
        sError = m_oSocket.LastError.Description
        GoTo QH
    End If
    lContentLength = -1
    Do
        bResult = m_oSocket.SyncReceiveArray(baRecv, Timeout:=5000)
        If UBound(baRecv) < 0 Then
            If m_oSocket.LastError <> 0 Then
                sError = m_oSocket.LastError.Description
                GoTo QH
            End If
            If m_oSocket.IsClosed Then
                Set m_oSocket = Nothing
                Exit Do
            End If
        Else
            mHttpsRequestValue = mHttpsRequestValue & StrConv(baRecv, vbUnicode)
           ' DebugLog MODULE_NAME, FUNC_NAME, "Len(mHttpsRequestValue)=" & Len(mHttpsRequestValue)
        End If
        If IsEmpty(vHeaders) Then
            lHeaderLength = InStr(1, mHttpsRequestValue, vbCrLf & vbCrLf) - 1
            If lHeaderLength > 0 Then
                vHeaders = Split(Left$(mHttpsRequestValue, lHeaderLength), vbCrLf)
                lHeaderLength = lHeaderLength + 4
                For Each vElem In vHeaders
                    If Left$(LCase(vElem), Len(HDR_CONTENT_LENGTH)) = HDR_CONTENT_LENGTH Then
                        lContentLength = val(Mid$(vElem, Len(HDR_CONTENT_LENGTH) + 1))
                    ElseIf Left$(LCase(vElem), Len(HDR_TRANSFER_ENCODING)) = HDR_TRANSFER_ENCODING Then
                        sEncoding = LCase$(Trim$(Mid$(vElem, Len(HDR_TRANSFER_ENCODING) + 1)))
                    ElseIf Left$(LCase(vElem), Len(HDR_CONNECTION)) = HDR_CONNECTION Then
                        sConnection = LCase$(Trim$(Mid$(vElem, Len(HDR_CONNECTION) + 1)))
                    End If
                Next
            End If
        End If
        If lContentLength >= 0 Then
            If Len(mHttpsRequestValue) >= lHeaderLength + lContentLength Then
                If Len(mHttpsRequestValue) <> lHeaderLength + lContentLength Then
                  '  DebugLog MODULE_NAME, FUNC_NAME, "Received " & Len(mHttpsRequestValue) & " instead of " & lHeaderLength + lContentLength, vbLogEventTypeWarning
                End If
                Exit Do
            End If
        ElseIf sEncoding = "chunked" Then
            If Right$(mHttpsRequestValue, 5) = "0" & vbCrLf & vbCrLf Then
                Exit Do
            End If
        End If
        If Not bResult Then
            sError = m_oSocket.LastError.Description
            GoTo QH
        End If
    Loop
    If Not m_oSocket Is Nothing And sConnection = "close" Then
        m_oSocket.Shutdown
        Set m_oSocket = Nothing
    End If

    HttpsRequest = True

    mHttpsRequestValue = Replace(Replace(mHttpsRequestValue, vbCr, vbNullString), vbLf, vbCrLf)
    
    If lContentLength = -1 And Len(mHttpsRequestValue) > 0 Then
    lHeaderLength = InStr(mHttpsRequestValue, vbCrLf + vbCrLf)
    
    End If
QH:
    If LenB(sError) <> 0 Then
        Set m_oSocket = Nothing
    Else
        If lContentLength > 0 Then
        
        Else
        
        End If
        
    End If
End Function
Property Let ServerName(RHS)
    m_sServerName = RHS
End Property

Private Function ParseUrl(sUrl As String, uParsed As UcsParsedUrl, Optional DefProtocol As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(?:(.*)://)?(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(/[^?#]*)?(\?[^#]*)?(#.*)?$"
        With .Execute(sUrl)
            If .Count > 0 Then
                With .item(0).SubMatches
                    uParsed.Protocol = .item(0)
                    uParsed.User = .item(1)
                    If LenB(uParsed.User) = 0 Then
                        uParsed.User = .item(2)
                    Else
                        uParsed.Pass = .item(2)
                    End If
                    uParsed.Host = .item(3)
                    uParsed.Port = val(Mid$(.item(4), 2))
                    If uParsed.Port = 0 Then
                        Select Case LCase$(IIf(LenB(uParsed.Protocol) = 0, DefProtocol, uParsed.Protocol))
                        Case "https"
                            uParsed.Port = 443
                        Case "socks5"
                            uParsed.Port = 1080
                        Case Else
                            uParsed.Port = 80
                        End Select
                    End If
                    uParsed.Path = .item(5)
                    If LenB(uParsed.Path) = 0 Then
                        uParsed.Path = "/"
                    End If
                    uParsed.QueryString = .item(6)
                    uParsed.Anchor = .item(7)
                End With
                ParseUrl = True
            End If
        End With
    End With
End Function

Private Function pvIsKnownBadCertificate(sHost As String) As Boolean
    Const STR_HOSTS     As String = "mikestoolbox.org|localhost"
    Dim vElem           As Variant
    
    For Each vElem In Split(STR_HOSTS, "|")
        If Right$(LCase$(sHost), Len(vElem)) = vElem Then
            pvIsKnownBadCertificate = True
            Exit For
        End If
    Next
End Function
Public Property Get HeadersCount() As Variant
If Not IsEmpty(vHeaders) Then
        Dim c As Collection
        Set c = vHeaders
        HeadersCount = c.Count
End If
End Property
Public Property Get Header(index) As Variant
    Header = vbNullString
    If Not IsEmpty(vHeaders) Then
        On Error Resume Next
        Dim c As Collection
        Set c = vHeaders
        Header = c.item(index)
    End If
End Property
Public Property Get Headers() As Variant
    If IsEmpty(vHeaders) Then
        Set Headers = New Collection
    Else
        Set Headers = vHeaders
    End If
End Property
' ConcatCollection
Public Property Get HeadersString() As Variant
    If IsEmpty(vHeaders) Then
        HeadersString = vbNullString
    Else
        Dim c As Collection
        Set c = vHeaders
        HeadersString = ConcatCollection(c)
    End If
End Property
Public Property Get Body() As Variant
    Dim i As Long
    i = InStr(mHttpsRequestValue, vbCrLf + vbCrLf)
    If i > 0 Then
        Body = Mid$(mHttpsRequestValue, InStr(mHttpsRequestValue, vbCrLf + vbCrLf) + 4)
    Else
        Body = mHttpsRequestValue
    End If
End Property
Public Property Get BodyFistLine() As Variant
On Error Resume Next
    Dim i As Long
    i = InStr(mHttpsRequestValue, vbCrLf + vbCrLf)
    If i > 0 Then
        BodyFistLine = Mid$(mHttpsRequestValue, i + 4)
        If Left$(BodyFistLine, 3) = "c" + vbCrLf Then
            BodyFistLine = Mid$(BodyFistLine, 4)
        End If
        i = InStr(BodyFistLine, vbCrLf)
        If i > 0 Then
            BodyFistLine = Mid$(BodyFistLine, 1, i - 1)
        End If
    Else
        BodyFistLine = mHttpsRequestValue
    End If
End Property
Sub EmptyVariant(ByVal b As Long)
   Dim a As Variant
   Static t(3) As Long
   CopyMemory t(0), ByVal VarPtr(a), 16
   CopyMemory ByVal VarPtr(a), ByVal b, 16
   CopyMemory ByVal b, t(0), 16
End Sub

Public Function ConcatCollection(oCol As Collection, Optional Separator As String = vbCrLf) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            If lSize <= Len(ConcatCollection) Then
                Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            End If
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function


