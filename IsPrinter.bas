Attribute VB_Name = "AnyPrinter"
' module isprinter there!
' Get information about all of the local printers using structure 1.  Note how
' the elements of the array are loaded into an array of data structures manually.  Also
' note how the following special declares must be used to allow numeric string pointers
' to be used in place of strings:
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersW" (ByVal flags As Long, ByVal name As Long, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Const PRINTER_ENUM_LOCAL = &H2
Private Type PRINTER_INFO_1
        flags As Long
        pDescription As String
        pname As String
        pComment As String
End Type
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const MAX_PATH As Long = 260
Type BrowseInfo
    hOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszINSTRUCTIONS As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


Function IsPrinter() As Boolean
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    
    Dim longbuffer() As Long  ' resizable array receives information from the function
    Dim printinfo() As PRINTER_INFO_1  ' values inside longbuffer() will be put into here
    Dim numbytes As Long  ' size in bytes of longbuffer()
    Dim numneeded As Long  ' receives number of bytes necessary if longbuffer() is too small
    Dim numprinters As Long  ' receives number of printers found
    Dim c As Integer, retval As Long  ' counter variable & return value
      ' Get information about the local printers
      n$ = String$(8192 / 2, Chr(0))
    numbytes = 8192 ' should be sufficiently big, but it may not be
    ReDim longbuffer(0 To numbytes / 4) As Long  ' resize array -- note how 1 Long = 4 bytes
    retval = EnumPrinters(PRINTER_ENUM_LOCAL, StrPtr(n$), 1, longbuffer(0), numbytes, numneeded, numprinters)
    If retval = 0 Then  ' try enlarging longbuffer() to receive all necessary information
        numbytes = numneeded
        ReDim longbuffer(0 To numbytes / 4) As Long  ' make it large enough
        retval = EnumPrinters(PRINTER_ENUM_LOCAL, StrPtr(n$), 1, longbuffer(0), numbytes, numneeded, numprinters)
        If retval = 0 Then ' failed again!
         GoTo there1
        Exit Function
    End If
    End If
    
there1:
    On Error GoTo there
If Printers.count > 0 Then
IsPrinter = True
Else
there:
Err.Clear
    IsPrinter = numprinters <> 0
    End If
    
End Function

