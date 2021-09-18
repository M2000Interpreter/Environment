Attribute VB_Name = "GetAdaptsInfo"
Option Explicit
'
'=============
'GetAdaptsInfo
'=============
'
'This module provides:
'
'   Public Type AdapterInfo
'   Public Function GetAdaptersInfo(
'       ByRef AdaptersInfo() As AdapterInfo) As Long
'
'The purpose is to return a list of network adapters,
'their first (possibly only) IP address, and their
'associated Default Gateway IP addresses.
'

Private Const MAX_ADAPTER_NAME_LENGTH = 260
Private Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Private Const ERROR_SUCCESS = 0
Private Const ERROR_NOT_SUPPORTED = 50
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_BUFFER_OVERFLOW = 111
Private Const ERROR_NO_DATA = 232

Public Type AdapterInfo
    Name As String
    AdapterIndex As Long
    IP As String
    Description As String
    GatewayIP As String
End Type

Private Type IP_ADDR_STRING
    Next As Long
    IpAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Private Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName As String * MAX_ADAPTER_NAME_LENGTH
    Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
    AddressLength As Long
    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    Index As Long
    Type As Long
    DhcpEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As IP_ADDR_STRING
    GatewayList As IP_ADDR_STRING
    DhcpServer As IP_ADDR_STRING
    HaveWins As Byte
    PrimaryWinsServer As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained As Long
    LeaseExpires As Long
End Type

Private Declare Function GetAdaptersInfoAPI Lib "IPHlpApi" Alias "GetAdaptersInfo" ( _
    APIAdapterInfo As Any, _
    pOutBufLen As Long) As Long
    
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
    
Private Function RNullTrim(ByVal Value As String) As String
    Dim lngNull As Long
    
    lngNull = InStr(Value, vbNullChar)
    RNullTrim = Left$(Value, lngNull - 1)
End Function

Public Function GetAdaptersInfo(ByRef AdaptersInfo() As AdapterInfo) As Long
    'Returns count of adapters.
    '
    'Call GetAdaptersInfoAPI() and update the list of adapters and
    'their Default Gateways.  If no adapters are found AdaptersInfo
    'is not altered.
    Dim pOutBufLen As Long
    Dim APIAdapterInfoBuffer() As Byte
    Dim pAdapt As Long
    Dim APIAdapterInfo As IP_ADAPTER_INFO
    Dim lngAdapterCount As Long
    
    'Note:  GetAdaptersInfoAPI() returns a linked list of adapter entries.
    
    'Find required buffer size.
    pOutBufLen = 0
    Select Case GetAdaptersInfoAPI(ByVal 0&, pOutBufLen)
        Case ERROR_SUCCESS
            Err.Raise &H8004B700, "GetAdaptersInfo", _
                      "GetAdaptersInfo Early Success: internal error"

        Case ERROR_NOT_SUPPORTED
            Err.Raise &H8004B710, "GetAdaptersInfo", _
                      "GetAdaptersInfo is not supported by this OS"

        Case ERROR_INVALID_PARAMETER
            Err.Raise &H8004B720, "GetAdaptersInfo", _
                      "GetAdaptersInfo Bad Parameters: internal error"

        Case ERROR_BUFFER_OVERFLOW
            ReDim APIAdapterInfoBuffer(pOutBufLen - 1)
        
            'Get adapter information by calling with adequate buffer.
            Select Case GetAdaptersInfoAPI(APIAdapterInfoBuffer(0), pOutBufLen)
                Case ERROR_SUCCESS
                    pAdapt = VarPtr(APIAdapterInfoBuffer(0))
                    
                    Do While pAdapt 'Not 0.
                        CopyMemory APIAdapterInfo, ByVal pAdapt, Len(APIAdapterInfo)
                        ReDim Preserve AdaptersInfo(lngAdapterCount)
                        With AdaptersInfo(lngAdapterCount)
                            .Name = RNullTrim(APIAdapterInfo.AdapterName)
                            .AdapterIndex = APIAdapterInfo.Index
                            .Description = RNullTrim(APIAdapterInfo.Description)
                            
                            'Take only 1st entry from each of next two lists, though
                            'on a server OS their can be several.
                            '
                            'In theory these may be null.  If so we store an empty
                            'String value.
                            .IP = RNullTrim(APIAdapterInfo.IpAddressList.IpAddress)
                            .GatewayIP = RNullTrim(APIAdapterInfo.GatewayList.IpAddress)
                        End With
                        pAdapt = APIAdapterInfo.Next
                        lngAdapterCount = lngAdapterCount + 1
                    Loop
                    
                    GetAdaptersInfo = lngAdapterCount
                
                Case ERROR_NOT_SUPPORTED
                    Err.Raise &H8004B730, "GetAdaptersInfo", _
                              "GetAdaptersInfo Late Failure: is not supported by this OS"

                Case ERROR_INVALID_PARAMETER
                    Err.Raise &H8004B740, "GetAdaptersInfo", _
                              "GetAdaptersInfo Late Failure, Bad Parameters: internal error"

                Case ERROR_BUFFER_OVERFLOW
                    Err.Raise &H8004B750, "GetAdaptersInfo", _
                              "GetAdaptersInfo Late Failure: buffer overflow"
                
                Case Else
                    Err.Raise &H8004B760, "GetAdaptersInfo", _
                              "GetAdaptersInfo Late Failure: system error " & CStr(Err.LastDllError)
            End Select
            
        Case ERROR_NO_DATA
            Exit Function
            
        Case Else
            Err.Raise &H8004B770, "GetAdaptersInfo", _
                      "GetAdaptersInfo system error " & CStr(Err.LastDllError)
    End Select
End Function
