VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' VbAsyncSocket Project (c) 2018-2019 by wqweto@gmail.com
'
' Simple and thin WinSock API wrappers for VB6
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================

' Change the pvParseUrl() function
' addition of START event
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cHttpDownload"

'=========================================================================
' Events
'=========================================================================
Event Start(ByVal remotehost$, ByVal remoteport As Long)
Event DownloadProgress(ByVal BytesRead As Double, ByVal BytesTotal As Double)
Event DownloadError(ByVal Number As Long, ByVal Description As String)
Event DownloadComplete(ByVal LocalFileName As String)

'=========================================================================
' API
'=========================================================================

'--- for SHCreateStreamOnFile
Private Const STGM_WRITE                    As Long = 1
Private Const STGM_CREATE                   As Long = &H1000

'Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const HDR_CONTENT_LENGTH    As String = "content-length:"
Private Const HDR_LOCATION          As String = "location:"
Private Const LNG_FLUSH_LIMIT       As Long = 1024& * 1024
'--- errors
Private Const ERR_INVALID_URL       As String = "Invalid URL"
Private Const ERR_HTTP_ONLY         As String = "Http protocol supported only"
Private Const ERR_INVALID_REDIRECT  As String = "Invalid redirect to %1"
Private Const ERR_INVALID_RESPONSE  As String = "Server returned %1"

Private WithEvents m_oSocket    As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_uRemote               As UcsParsedUrl
Private m_sLocalFileName        As String
Private m_file                  As Boolean
Private m_dStartDate            As Date
'Private m_lCallbackPtr          As Long
Private m_wait_close            As Boolean
Private m_reading               As Long
Private m_eState                As UcsStateEnum
Private m_baRecvBuffer()        As Byte
Private m_dblBytesRead          As Double
Private m_dblContentLength      As Double

Private Enum UcsStateEnum
    ucsIdle
    ucsWaitHeaders
    ucsWaitBody
End Enum

Private Type UcsParsedUrl
    Protocol        As String
    Host            As String
    Port            As Long
    Path            As String
    User            As String
    Pass            As String
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    ' debug.print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get url() As String
    url = m_uRemote.Protocol & "://" & m_uRemote.Host & m_uRemote.Path
End Property

Property Get LocalFileName() As String
    LocalFileName = m_sLocalFileName
End Property

Property Get StartDate() As Date
    StartDate = m_dStartDate
End Property

'Friend Property Get CallbackWeakRef() As Object
'    Call vbaObjSetAddref(CallbackWeakRef, m_lCallbackPtr)
'End Property

'Friend Property Set CallbackWeakRef(oValue As Object)
'    m_lCallbackPtr = ObjPtr(oValue)
'End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub DownloadFile(url As String, LocalFileName As String)
    Const FUNC_NAME     As String = "DownloadFile"
    Dim hResult     As Long
    
    On Error GoTo EH
    If Not pvParseUrl(url, m_uRemote) Then
     '  On Error GoTo 0
        Err.Raise vbObjectError, , ERR_INVALID_URL
    End If
    If LCase$(m_uRemote.Protocol) <> "http" Then
        On Error GoTo 0
        Err.Raise vbObjectError, , ERR_HTTP_ONLY
    End If
    m_wait_close = False
        OpenStream (LocalFileName)
        If Not m_file Then
           Exit Sub
        End If
    m_dStartDate = Now
    m_eState = ucsIdle
    m_dblBytesRead = 0
    m_baRecvBuffer = vbNullString
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.create(SocketType:=ucsSckStream) Then
        On Error GoTo 0
        Err.Raise vbObjectError, , m_oSocket.GetErrorDescription(m_oSocket.LastError)
    End If
    If Not m_oSocket.Connect(m_uRemote.Host, m_uRemote.Port) Then
        On Error GoTo 0
        Err.Raise vbObjectError, , m_oSocket.GetErrorDescription(m_oSocket.LastError)
    End If
    ' debug.print "Start:", m_uRemote.Host
    RaiseEvent Start(m_uRemote.Host, m_uRemote.Port)
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Public Sub CancelDownload()
    Const FUNC_NAME     As String = "CancelDownload"
    
    On Error GoTo EH
    Set m_oSocket = Nothing
    'm_lCallbackPtr = 0
    m_eState = ucsIdle
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'= private ===============================================================

Private Function pvRecvHeaders(baBuffer() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvRecvHeaders"
    Const STR_DELIM     As String = vbCrLf & vbCrLf
    Dim lPos            As Long
    Dim sHeaders        As String
    Dim vSplit          As Variant
    Dim vElem           As Variant
    Dim uRedirect       As UcsParsedUrl
    
    On Error GoTo EH
    sHeaders = m_oSocket.FromTextArray(m_baRecvBuffer, ucsScpAcp) & m_oSocket.FromTextArray(baBuffer, ucsScpAcp)
    lPos = InStr(sHeaders, STR_DELIM)
    If lPos > 0 Then
        vSplit = Split(Left$(sHeaders, lPos), vbCrLf)
        Select Case Mid$(vSplit(0), 10, 3)
        Case "301", "302"
            For Each vElem In vSplit
                If Left$(LCase$(vElem), Len(HDR_LOCATION)) = HDR_LOCATION Then
                    vElem = Trim$(Mid$(vElem, Len(HDR_LOCATION) + 1))
                    If Not pvParseUrl(CStr(vElem), uRedirect) Then
                        pvSetError vbObjectError, Replace(ERR_INVALID_REDIRECT, "%1", vElem)
                        GoTo QH
                    End If
                    uRedirect.User = m_uRemote.User
                    uRedirect.Pass = m_uRemote.Pass
                    Exit For
                End If
            Next
            If LCase$(uRedirect.Protocol) <> "http" Then
                pvSetError vbObjectError, ERR_HTTP_ONLY
                GoTo QH
            End If
            m_uRemote = uRedirect
            m_oSocket.Close_
            If Not m_oSocket.create(SocketType:=ucsSckStream) Then
                pvSetError m_oSocket.LastError
                GoTo QH
            End If
            If Not m_oSocket.Connect(m_uRemote.Host, m_uRemote.Port) Then
                pvSetError m_oSocket.LastError
                GoTo QH
            End If
        Case "200"
            m_dblContentLength = -1
            For Each vElem In vSplit
                If Left$(LCase$(vElem), Len(HDR_CONTENT_LENGTH)) = HDR_CONTENT_LENGTH Then
                    m_dblContentLength = val(Mid$(vElem, Len(HDR_CONTENT_LENGTH) + 1))
                End If
            Next
            m_eState = ucsWaitBody
            pvRecvBody m_oSocket.ToTextArray(Mid$(sHeaders, lPos + Len(STR_DELIM)), ucsScpAcp)
        Case Else
            pvSetError vbObjectError, Replace(ERR_INVALID_RESPONSE, "%1", Mid$(vSplit(0), 10))
            GoTo QH
        End Select
    End If
    '--- success
    pvRecvHeaders = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvRecvBody(baBuffer() As Byte, Optional ByVal Flush As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvRecvBody"
    Dim lIdx            As Long
    Dim hResult         As Long
    
    On Error GoTo EH
    If UBound(baBuffer) >= 0 Then
        m_dblBytesRead = m_dblBytesRead + UBound(baBuffer) + 1
        
        WriteStream baBuffer()
       
        'If Not (m_wait_close And m_reading = 1) Then
        ' debug.print "DownloadProgress", m_sLocalFileName, m_dblBytesRead
    
        RaiseEvent DownloadProgress(m_dblBytesRead, m_dblContentLength)
       ' End If
        'If m_lCallbackPtr <> 0 Then
        '    Call CallbackWeakRef.DownloadProgress(Me, m_dblBytesRead, m_dblContentLength)
        'End If
    End If
    '--- success
    pvRecvBody = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Sub pvSetError(ByVal lNumber As Long, Optional sDescription As String)
    Const FUNC_NAME     As String = "pvSetError"
    
    On Error GoTo EH
    If LenB(sDescription) = 0 Then
        If Not m_oSocket Is Nothing Then
            sDescription = m_oSocket.GetErrorDescription(lNumber)
        Else
            sDescription = "Error " & lNumber
        End If
    End If
    CancelDownload
    RaiseEvent DownloadError(lNumber, sDescription)
    'If m_lCallbackPtr <> 0 Then
    '    Call CallbackWeakRef.DownloadError(Me, lNumber, sDescription)
    'End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvParseUrl(sUrl As String, uParsed As UcsParsedUrl) As Boolean
    Const URL_PART_SCHEME As Long = 1
    Const URL_PART_HOSTNAME As Long = 2
    Const URL_PART_USERNAME As Long = 3
    Const URL_PART_PASSWORD As Long = 4
    Const URL_PART_PORT As Long = 5
    With uParsed
    .Host = GetUrlParts(sUrl, URL_PART_HOSTNAME)
    .Pass = GetUrlParts(sUrl, URL_PART_PASSWORD)
    .Path = GetUrlPath(sUrl)
     If Len(.Path) = 0 Then .Path = "/"
    .Port = CLng("0" + GetUrlParts(sUrl, URL_PART_PORT))
    .Protocol = GetUrlParts(sUrl, URL_PART_SCHEME)
    .User = GetUrlParts(sUrl, URL_PART_USERNAME)
    If .Port = 0 Then
        Select Case LCase$(uParsed.Protocol)
        Case "https"
            uParsed.Port = 443
        Case "socks5"
            uParsed.Port = 1080
        Case Else
            uParsed.Port = 80
        End Select
    End If
    pvParseUrl = True
    End With
End Function


Private Sub Class_Terminate()
'
If Not m_oSocket Is Nothing Then
    m_oSocket.Close_
    CancelDownload
End If
End Sub

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oSocket_OnConnect()
    Const FUNC_NAME     As String = "m_oSocket_OnConnect"
    
    On Error GoTo EH
    m_eState = ucsWaitHeaders
    If Not m_oSocket.SendText("GET " & m_uRemote.Path & " HTTP/1.0" & vbCrLf & _
            "Host: " & m_uRemote.Host & vbCrLf & _
            "Accept: */*" & vbCrLf & vbCrLf, CodePage:=ucsScpAcp) Then
        pvSetError m_oSocket.LastError
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim baBuffer()      As Byte

    On Error GoTo EH
    If m_eState = ucsWaitHeaders Then
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            pvSetError m_oSocket.LastError
            GoTo QH
        End If
        If UBound(baBuffer) < 0 Then
            GoTo QH
        End If
        If Not pvRecvHeaders(baBuffer) Then
            GoTo QH
        End If
    End If

    If m_eState = ucsWaitBody Then
        If m_reading > 0 Then
            ' debug.print "reading in progess....", m_sLocalFileName, FileLen(GetDosPath(m_sLocalFileName)), m_reading
            m_reading = m_reading + 1
            Exit Sub
        End If
        
           
            m_reading = m_reading + 1
again:
             baBuffer = vbNullString
            If Not m_oSocket.ReceiveArray(baBuffer) Then
                pvSetError m_oSocket.LastError
                
                
             End If
            
      '
            If UBound(baBuffer) > 0 Then
again2:
            If pvRecvBody(baBuffer) Then
            ' debug.print "reading now....", m_sLocalFileName, FileLen(GetDosPath(m_sLocalFileName)), m_reading, m_wait_close
            Else
             RaiseEvent DownloadError(m_oSocket.LastError, m_oSocket.GetErrorDescription(m_oSocket.LastError))
            CancelDownload
            End If
            Else
            ' debug.print "missing data", m_sLocalFileName
            If Not m_wait_close Then If m_oSocket.ReceiveArray(baBuffer) Then If UBound(baBuffer) > 0 Then GoTo again2
            'Exit Sub
        End If
            
            If m_wait_close And m_reading < 2 Then
            RaiseEvent DownloadComplete(m_sLocalFileName)
            CancelDownload
            Exit Sub
            End If
       m_reading = m_reading - 1
       If m_reading > 0 Then GoTo again
    End If
QH:
    Exit Sub
EH:
m_reading = m_reading - 1
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnClose()
    Const FUNC_NAME     As String = "m_oSocket_OnClose"
    'Dim baBuffer()      As Byte
    If m_reading Then
        ' debug.print "here////////////////////////", m_sLocalFileName
        
        m_wait_close = True
        Exit Sub
    End If
    On Error GoTo EH
    If m_eState = ucsWaitBody Then
    m_wait_close = True
    m_reading = -1
    m_oSocket_OnReceive
        ' need to close file before evenet OnClose
        'RaiseEvent DownloadComplete(m_sLocalFileName)
        'If m_lCallbackPtr <> 0 Then
         '   Call CallbackWeakRef.DownloadComplete(Me, m_sLocalFileName)
        'End If
    End If
    CancelDownload
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    ' debug.print ErrorCode, m_sLocalFileName
    On Error GoTo EH
    pvSetError ErrorCode
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub
Private Sub OpenStream(s$)
If Not NeoUnicodeFile(s$) Then NoCreateFile: Exit Sub
Dim i As Long
i = FreeFile
m_sLocalFileName = s$
If CFname(GetDosPath(s$)) <> "" Then
    Open GetDosPath(s$) For Binary Access Write Lock Read Write As i
    Close i
    m_file = True
End If

End Sub
Private Function WriteStream(a() As Byte)
If m_file Then
Dim i As Long
i = FreeFile
If CFname(GetDosPath(m_sLocalFileName)) <> "" Then
    Open GetDosPath(m_sLocalFileName) For Binary Access Write Lock Read Write As i
    Seek i, LOF(i) + 1
    Put #i, , a()
    Close i
    m_file = True
End If
End If
End Function
