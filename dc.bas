Attribute VB_Name = "Module3"
Option Explicit
Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Const RC_PALETTE As Long = &H100
Const SIZEPALETTE As Long = 104
Const RASTERCAPS As Long = 38
Dim sapi As Object
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type
Private Type GUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type
Private Type PicBmp
    Size As Long
    Type As Long
    hbmp As Long
    hPal As Long
    reserved As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As Long, ByVal hPrinter As Long) As Long
Private Declare Function ResetPrinter Lib "winspool.drv" Alias "ResetPrinterA" (ByVal hPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Type PRINTER_DEFAULTS
        pDataType As String
        pDevMode As DEVMODE
        DesiredAccess As Long
End Type
' New Win95 Page Setup dialogs are up to you
Private Type POINTL
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Private Const MAX_PATH As Long = 1024
Private Const MAX_PATH_UNICODE As Long = 260 * 2 - 1

Private Declare Function GetLongPathName Lib "kernel32" _
   Alias "GetLongPathNameW" _
  (ByVal lpszShortPath As Long, _
   ByVal lpszLongPath As Long, _
   ByVal cchBuffer As Long) As Long
 
Public Function GetLongName(strTest As String) As String
   Dim sLongPath As String
   Dim buff As String
   Dim cbbuff As Long
   Dim Result As Long
 
   buff = space$(MAX_PATH_UNICODE)
   cbbuff = Len(buff)
 
   Result = GetLongPathName(StrPtr(strTest), StrPtr(buff), cbbuff)
 
   If Result > 0 Then
      sLongPath = Left$(buff, Result)
   End If
 
   GetLongName = sLongPath
 
End Function
 


Function PathStrip2root(Path$) As String
Dim i As Long
If Len(Path$) >= 2 Then
If Mid$(Path$, 2, 1) = ":" Then
PathStrip2root = Left$(Path$, 2) & "\"
Else
i = InStrRev(Path$, Left$(Path$, 1))
If i > 1 Then
PathStrip2root = "\" & ExtractPath(Mid$(Path$, 2, i))
Else
PathStrip2root = Left$(Path$, 1)
End If

End If
End If
End Function

Sub Pprop()
    
    If ThereIsAPrinter = False Then Exit Sub
        
    Dim x As Printer
For Each x In Printers
If x.DeviceName = pname And x.Port = Port Then Exit For
Next x
Dim gp As Long, Td As PRINTER_DEFAULTS
Call OpenPrinter(x.DeviceName, gp, Td)
If form5iamloaded Then
Call PrinterProperties(Form5.hWnd, gp)
Else
Call PrinterProperties(Form1.hWnd, gp)
End If
Call ResetPrinter(gp, Td)
Call ClosePrinter(gp)
End Sub
Function CreateBitmapPicture(ByVal hbmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long, pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID

    'Fill GUID info
    With IID_IDispatch
        .data1 = &H20400
        .data4(0) = &HC0
        .data4(7) = &H46
    End With

    'Fill picture info
    With pic
        .Size = Len(pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hbmp = hbmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    'Create the picture
    R = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function
Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hbmp As Long, hBmpPrev As Long, R As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE

    'Create a compatible device context
    hDCMemory = CreateCompatibleDC(hDCSrc)
    'Create a compatible bitmap
    hbmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    'Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hbmp)

    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    'Does our picture use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        LogPal.palVersion = &H300
        'Number of palette entries
        LogPal.palNumEntries = 256
        'Retrieve the system palette entries
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        'Create the palette
        hPal = CreatePalette(LogPal)
        'Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        'Realize the palette
        R = RealizePalette(hDCMemory)
    End If

    'Copy the source image to our compatible device context
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    'Restore the old bitmap
    hbmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Select the palette
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    'Delete our memory DC
    R = DeleteDC(hDCMemory)

    Set hDCToPicture = CreateBitmapPicture(hbmp, hPal)
End Function

Function DriveType(ByVal Path$) As String
    Select Case GetDriveType(Path$)
        Case 2
            DriveType = "Μεταθέσιμο"
        Case 3
            DriveType = "Σταθερό"
        Case Is = 4
            DriveType = "Απόμακρο"
        Case Is = 5
            DriveType = "Cd-Rom"
        Case Is = 6
            DriveType = "Προσωρινό στην μνήμη"
        Case Else
            DriveType = "Απροσδιόριστο"
    End Select
End Function

Function DriveTypee(ByVal Path$) As String
    Select Case GetDriveType(Path$)
        Case 2
            DriveTypee = "Removable"
        Case 3
            DriveTypee = "Drive Fixed"
        Case Is = 4
            DriveTypee = "Remote"
        Case Is = 5
            DriveTypee = "Cd-Rom"
        Case Is = 6
            DriveTypee = "Ram disk"
        Case Else
            DriveTypee = "Unrecognized"
    End Select
End Function
Function DriveSerial(ByVal Path$) As Long
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim Serial As Long, Vname As String, FSName As String
    'Create buffers
    If Len(Path$) = 1 Then Path$ = Path$ & ":\"
    If Len(Path$) = 2 Then Path$ = Path$ & "\"
    Vname = String$(255, Chr$(0))
    FSName = String$(255, Chr$(0))
    'Get the volume information
    GetVolumeInformation Path$, Vname, 255, Serial, 0, 0, FSName, 255
    'Strip the extra chr$(0)'s
    'VName = Left$(VName, InStr(1, VName, Chr$(0)) - 1)
    'FSName = Left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
 DriveSerial = Serial
End Function

Function WeCanWrite(ByVal Path$) As Boolean
Dim SecondTry As Boolean, pp$
On Error GoTo wecant
pp$ = ExtractPath(Path$, , True)
pp$ = GetDosPath(pp$)
If pp$ = vbNullString Then
MyEr "Not writable device " & Path$, "Δεν μπορώ να γράψω στη συσκευή " & Path$
Exit Function
End If
pp$ = PathStrip2root(Path$)


    Select Case GetDriveType(pp$)

        Case 2, 3, 4, 6
          WeCanWrite = Not GetAttr(pp$) = vbReadOnly
        Case 5
           WeCanWrite = False
    End Select
   Exit Function
wecant:
                   If Err.Number > 0 Then
                Err.Clear
                 MyEr "Not writable device " & Path$, "Δεν μπορώ να γράψω στη συσκευή " & Path$
            WeCanWrite = False
                Exit Function
                End If

End Function
Public Function VoiceName(ByVal D As Double) As String
On Error Resume Next
Dim o As Object
If Typename(sapi) = "Nothing" Then Set sapi = CreateObject("sapi.spvoice")
If Typename(sapi) = "Nothing" Then VoiceName = vbNullString: Exit Function
D = Int(D)
If sapi.getvoices().Count >= D And D > 0 Then
For Each o In sapi.getvoices
D = D - 1
If D = 0 Then VoiceName = o.GetDescription: Exit For
Next o
End If
End Function
Public Function NumVoices() As Long
On Error Resume Next
If Typename(sapi) = "Nothing" Then Set sapi = CreateObject("sapi.spvoice")
If Typename(sapi) = "Nothing" Then NumVoices = -1: Exit Function
If sapi.getvoices().Count > 0 Then
NumVoices = sapi.getvoices().Count
End If
End Function
Public Sub SPEeCH(ByVal a$, Optional BOY As Boolean = False, Optional ByVal vNumber As Long = -1)
Static lastvoice As Long
If vNumber = -1 Then vNumber = lastvoice
On Error Resume Next
If vNumber = 0 Then vNumber = 1
If Typename(sapi) = "Nothing" Then Set sapi = CreateObject("sapi.spvoice")
If Typename(sapi) = "Nothing" Then Beep: Exit Sub
If sapi.getvoices().Count > 0 Then
If sapi.getvoices().Count < vNumber Or sapi.getvoices().Count < 1 Then vNumber = 1
 With sapi
         Set .Voice = .getvoices.item(vNumber - 1)
       If BOY Then
         .volume = IIf(vol = 0, 0, 50 + vol \ 2)
        
         .rate = 2
       ' boy
         .Speak "<pitch absmiddle='25'>" & a$
         Else
         
         'man
       .rate = 1
       .volume = IIf(vol = 0, 0, 50 + vol \ 2)
         .Speak "<pitch absmiddle='-5'>" & a$
         End If
       End With
       lastvoice = vNumber
End If
End Sub
Public Sub wwPlain2(bstack As basetask, mybasket As basket, ByVal what As String, ByVal wi As Long, ByVal Hi As Long, Optional scrollme As Boolean = False, Optional nosettext As Boolean = False, Optional frmt As Long = 0, Optional ByVal skip As Long = 0, Optional res As Long, Optional isAcolumn As Boolean = False, Optional collectit As Boolean = False, Optional nonewline As Boolean)
Dim DDD As Object, mDoc As Object, para() As String, i As Long
Dim n As Long, st As Long, st1 As Long, st0 As Long, w As Integer
Dim PX As Long, PY As Long, nowait As Boolean
Dim nopage As Boolean
Dim buf$, b$, npy As Long, lCount As Long, SCRnum2stop As Long
Dim nopr As Boolean, nohi As Long, w2 As Long, lastPara As Long
Dim dv2x15 As Long
dv2x15 = dv15 * 2
If collectit Then Set mDoc = New Document
Set DDD = bstack.Owner
If what = vbNullString Then
ReDim para(1)
para(1) = vbNullString
Else
para() = Split(what, vbCrLf)
End If
With mybasket
' from old code here
    tParam.iTabLength = .ReportTab
    PX = .curpos
    PY = .currow
    If Not nosettext Then
        If PX >= .mx Then
            nowait = True
            PX = 0
        End If
    End If
    If PX > .mx Then nowait = True
    If wi = 0 Then
        If nowait Then wi = .Xt * (.mx - PX) Else wi = .mx * .Xt
    Else
        If wi <= .mx Then wi = wi * .Xt
    End If
    
    wi = wi - CLng(dv2x15)
    If wi <= 0 Then Exit Sub
    If Hi < 0 Then
        Hi = -Hi - 2
        nohi = Hi
        nopr = True
    End If

    If Not nopr Then
        If Not nosettext Then
        If PY = .my And .double Then
            crNew bstack, mybasket
            PY = .currow
        End If
             LCTbasket DDD, mybasket, PY, PX
        End If
        DDD.currentX = DDD.currentX + dv2x15
        If Not scrollme Then
            If Hi >= 0 Then If (.my - PY) * .Yt < Hi Then Hi = (.my - PY) * .Yt
        Else
            If Hi > 1 Then
                If .pageframe <> 0 Then
                    lCount = holdcontrol(DDD, mybasket)
                    .pageframe = 0
                End If
                SCRnum2stop = holdcontrol(DDD, mybasket)
            End If
        End If
    End If
    npy = PY

    w2 = wi
    If bstack.IamThread Then nopage = True
' -----end old code---------

lastPara = UBound(para)
If Len(para(lastPara)) = 0 Then lastPara = lastPara - 1
    For i = LBound(para) To lastPara
    
        buf$ = vbNullString
nextline:
        If NOEXECUTION Then Exit For
        n = MyTextWidth(DDD, para(i))
        If n > wi Then
            st = Len(para(i))
            st1 = st + 1
            st0 = 1
            While st > st0 + 1
                st1 = (st + st0) \ 2
                w = AscW(Mid$(para(i), 1, st1))
                If w > -10241 And w < -9216 Then
                    If wi >= MyTextWidth(DDD, Mid$(para(i), 1, st1 + 1)) Then
                        st0 = st1
                    Else
                        st = st1
                    End If
                Else
                    If wi >= MyTextWidth(DDD, Mid$(para(i), 1, st1)) Then
                        st0 = st1
                    Else
                        st = st1
                    End If
                End If
            Wend
            st = rinstr(para(i), "_", Len(para(i)) - st0)
            st1 = rinstr(para(i), " ", Len(para(i)) - st0)
            If st > st1 Then
               st1 = st
            Else
                If MyTrimL3Len(Mid$(para(i), 1, st1)) = st1 Then st1 = st0
      
            End If
            If st1 >= Len(para(i)) Then
                buf$ = para(i)
            Else
                buf$ = Left$(para(i), st1)
                     If st1 = st0 Then
                     If Right$(buf$, 1) = "_" Then Mid$(buf$, st1, 1) = "-"
                        End If
                para(i) = LTrim(Mid$(para(i), st1 + 1))
            End If
             skip = skip - 1
             If skip < 0 Then
                 If Len(para(i)) = 0 Then GoTo last
                 If frmt > 0 Then
                     If Not nopr Then fullPlainWhere DDD, mybasket, RTrim$(buf$), w2, frmt, nowait, nonewline ' rtrim
                 Else
                     If Not nopr Then fullPlain DDD, mybasket, RTrim$(buf$), w2, nowait, nonewline
                 End If
                 If collectit Then mDoc.AppendParagraphOneLine RTrim$(buf$)
            
             End If
        Else
            skip = skip - 1
            If Len(buf$) > 0 Then para(i) = Mid$(para(i), MyTrimL3Len(para(i)) + 1)
            buf$ = para(i)
            para(i) = vbNullString
last:
        If skip >= 0 Then GoTo continue
        If Hi = 0 And frmt = 0 Then
        If Not scrollme Then
        If Not nopr Then
    
            MyPrintNew DDD, mybasket.uMineLineSpace, buf$, , nowait     ';   '************************************************************************************
    
            res = DDD.currentX
                    If Trim$(buf$) = vbNullString Then
                    DDD.currentX = ((DDD.currentX + .Xt \ 2) \ .Xt) * .Xt
                    Else
                    DDD.currentX = ((DDD.currentX + .Xt \ 1.2) \ .Xt) * .Xt
                    End If
            End If
          If Not (isAcolumn Or nosettext) Then If Not nonewline Then GoTo JUMPHERE
            Exit Sub
        End If
        If Not nopr Then
            fullPlainWhere DDD, mybasket, buf$, w2, frmt, nowait, nonewline
        End If
        Else
        If frmt > 0 Then
                If Not nopr Then fullPlainWhere DDD, mybasket, buf$, w2, frmt, nowait, nonewline
        ElseIf frmt = 0 Then
        If Not nopr Then fullPlainWhere DDD, mybasket, buf$, w2, 3, nowait, nonewline
        Else
                If Not nopr Then fullPlain DDD, mybasket, buf$, w2, nowait, nonewline  'DDD.Width ' w2
        End If
        End If
   
        If collectit Then
        mDoc.AppendParagraphOneLine Trim$(buf$)
        End If
             End If
                If isAcolumn Then Exit Sub
                'If UBound(para) = i Then Exit For
        If skip < 0 Or scrollme Then
JUMPHERE:
            lCount = lCount + 1
            npy = npy + 1
            
            If npy >= .my And scrollme Then
                         If Not nopr Then
                             If SCRnum2stop > 0 Then
                                 If lCount >= SCRnum2stop Then
                                   If Not bstack.toprinter Then
                                    If Not nowait Then
                                 
                                 If Not nopage Then
                                  DDD.Refresh
                                     Do
                
                                        ' mywait bstack, 10
                                                      If TaskMaster Is Nothing Then
                                                Sleep 10
                                                ElseIf Not TaskMaster.Processing And TaskMaster.QueueCount = 0 Then
                                                Sleep 10
                                                End If
                                                MyDoEventsNoRefresh
                                     Loop Until INKEY$ <> "" Or mouse <> 0 Or NOEXECUTION
                                     End If
                                     End If
                                     End If
                                     SCRnum2stop = .pageframe
                                     lCount = 1
                                 
                                 End If
                             End If
                           If Not bstack.toprinter Then
                                DDD.Refresh
                                ScrollUpNew DDD, mybasket
                            Else
                              getnextpage
                              npy = 1
                          End If
                End If
                npy = npy - 1
                      ''
         ElseIf npy >= .my Then
         
        If Not nopr Then crNew bstack, mybasket
               npy = npy - 1
              
        End If
      End If
If Not nopr Then LCTbasket DDD, mybasket, npy, PX: DDD.currentX = DDD.currentX + dv2x15
If skip < 0 Then Hi = Hi - 1
If Hi < 0 Then Exit For
continue:
     If Len(para(i)) > 0 Then GoTo nextline
    Next i
End With
finish:
If scrollme Then
HoldReset lCount, mybasket
End If
res = nohi - Hi
wi = DDD.currentX
If collectit Then bstack.soros.PushStr mDoc.textDoc
End Sub
Public Sub wwPlain(bstack As basetask, mybasket As basket, ByVal what As String, ByVal wi As Long, ByVal Hi As Long, Optional scrollme As Boolean = False, Optional nosettext As Boolean = False, Optional frmt As Long = 0, Optional ByVal skip As Long = 0, Optional res As Long, Optional isAcolumn As Boolean = False, Optional collectit As Boolean = False, Optional nonewline As Boolean)

Dim DDD As Object, mDoc As Object, para() As String, i As Long
Dim n As Long, st As Long, st1 As Long, st0 As Long, w As Integer
Dim PX As Long, PY As Long, nowait As Boolean
Dim nopage As Boolean
Dim buf$, b$, npy As Long, lCount As Long, SCRnum2stop As Long
Dim nopr As Boolean, nohi As Long, w2 As Long, lastPara As Long
Dim dv2x15 As Long, Extra As Long, cuts As Long, tabw As Long, olda As Long, lasttab As Long, INTD As Long
dv2x15 = dv15 * 2
Dim meta As Boolean
meta = TypeOf bstack.Owner Is MetaDc
If collectit Then Set mDoc = New Document
Set DDD = bstack.Owner
If what = vbNullString Then
ReDim para(1)
para(1) = vbNullString
Else
para() = Split(what, vbCrLf)
End If
With mybasket
' from old code here
    tParam.iTabLength = .ReportTab
    tabw = .ReportTab * AverCharSpace(DDD)
    PX = .curpos
    PY = .currow
    If Not nosettext Then
        If PX >= .mx Then
            nowait = True
            PX = 0
        End If
    End If
    If PX > .mx Then nowait = True
    If wi = 0 Then
        If nowait Then wi = .Xt * (.mx - PX) Else wi = .mx * .Xt
    Else
        If wi <= .mx Then wi = wi * .Xt
    End If
    
    wi = wi - CLng(dv2x15)
    If wi <= 0 Then Exit Sub
    If Hi < 0 Then
        Hi = -Hi - 2
        nohi = Hi
        nopr = True
    End If

    If Not nopr Then
        If Not nosettext Then
        If PY = .my And .double Then
            crNew bstack, mybasket
            PY = .currow
        End If
             LCTbasket DDD, mybasket, PY, PX
        End If
        DDD.currentX = DDD.currentX + dv2x15
        If Not scrollme Then
            If Hi >= 0 Then If (.my - PY) * .Yt < Hi Then Hi = (.my - PY) * .Yt
        Else
            If Hi > 1 Then
                If .pageframe <> 0 Then
                    lCount = holdcontrol(DDD, mybasket)
                    .pageframe = 0
                End If
                SCRnum2stop = holdcontrol(DDD, mybasket)
            End If
        End If
    End If
    npy = PY

    w2 = wi
    If bstack.IamThread Then nopage = True
' -----end old code---------

lastPara = UBound(para)
If Len(para(lastPara)) = 0 Then lastPara = lastPara - 1
    For i = LBound(para) To lastPara
    
        buf$ = vbNullString
nextline:
        If NOEXECUTION Then Exit For
        
        n = LowWord(GetTabbedTextExtent(DDD.hdc, StrPtr(para(i)), Len(para(i)), 1, tabw)) * DXP
'        n = MyTextWidth(ddd, para(i))
        If n > wi Then
            st = Len(para(i))
            st1 = st + 1
            st0 = 1
            While st > st0 + 1
                st1 = (st + st0) \ 2
                w = AscW(Mid$(para(i), 1, st1))
                If w > -10241 And w < -9216 Then
                    
                    If wi >= LowWord(GetTabbedTextExtent(DDD.hdc, StrPtr(para(i)), st1 + 1, 1, tabw)) * DXP Then
                        st0 = st1
                    Else
                        st = st1
                    End If
                Else
                    If wi >= LowWord(GetTabbedTextExtent(DDD.hdc, StrPtr(para(i)), st1, 1, tabw)) * DXP Then
                        st0 = st1
                    Else
                        st = st1
                    End If
                End If
            Wend
            st = rinstr(para(i), "_", Len(para(i)) - st0)
            st1 = rinstr(para(i), " ", Len(para(i)) - st0)
            If st > st1 Then
               st1 = st
            Else
                If MyTrimL3Len(Mid$(para(i), 1, st1)) = st1 Then st1 = st0
      
            End If
            If st1 >= Len(para(i)) Then
                buf$ = para(i)
            Else
                buf$ = Left$(para(i), st1)
                     If st1 = st0 Then
                     If Right$(buf$, 1) = "_" Then Mid$(buf$, st1, 1) = "-"
                        End If
                para(i) = LTrim(Mid$(para(i), st1 + 1))
            End If
             skip = skip - 1
             If skip < 0 Then
                 If Len(para(i)) = 0 Then GoTo last
                 If Not nopr Then
                 INTD = 0

                 Extra = 0

                 Select Case frmt
                 Case 0
                 INTD = TextWidth(DDD, space$(MyTrimL3Len(buf$)))
                            If INTD > 0 Then
                 buf$ = Mid$(buf$, MyTrimL3Len(buf$) + 1)
                 DDD.currentX = DDD.currentX + INTD
                 End If
                 buf$ = RTrim(buf$)
                 lasttab = rinstr(buf$, vbTab)
                 If lasttab > 0 Then
                 Extra = LowWord(TabbedTextOut(DDD.hdc, DDD.currentX \ DXP, DDD.currentY \ DXP, StrPtr(buf$), lasttab, 1, tabw, DDD.currentX \ DXP))
                 buf$ = Mid$(buf$, lasttab + 1)
                 DDD.currentX = DDD.currentX + Extra * DXP
                              
                 End If
                 olda = SetTextAlign(DDD.hdc, 0) 'TA_RTLREADING)
                 cuts = Len(buf$) - Len(Replace$(buf$, " ", ""))
                 
                 Extra = wi \ DXP - INTD \ DXP - Extra - LowWord(GetTabbedTextExtent(DDD.hdc, StrPtr(buf$), Len(buf$), 1, tabw))
                 
                 SetTextJustification DDD.hdc, Extra, cuts
                 
                '
                 TextOut DDD.hdc, DDD.currentX \ DXP, DDD.currentY \ DXP, StrPtr(buf$), Len(buf$)
                 
                 SetTextJustification DDD.hdc, 0, 0
                 olda = SetTextAlign(DDD.hdc, olda)
                 Case 1
                 buf$ = RTrim(buf$)
                 Extra = wi \ DXP - LowWord(GetTabbedTextExtent(DDD.hdc, StrPtr(buf$), Len(buf$), 1, tabw))
                 
                 Extra = LowWord(TabbedTextOut(DDD.hdc, DDD.currentX \ DXP + Extra, DDD.currentY \ DXP, StrPtr(buf$), Len(buf$), 1, tabw, DDD.currentX \ DXP + Extra))
                 Case 2
                 buf$ = Trim(buf$)
                 Extra = (wi \ DXP - LowWord(GetTabbedTextExtent(DDD.hdc, StrPtr(buf$), Len(buf$), 1, tabw))) \ 2
                 
                 Extra = LowWord(TabbedTextOut(DDD.hdc, DDD.currentX \ DXP + Extra, DDD.currentY \ DXP, StrPtr(buf$), Len(buf$), 1, tabw, DDD.currentX \ DXP + Extra))
                 
                 Case Else
                 Extra = LowWord(TabbedTextOut(DDD.hdc, DDD.currentX \ DXP, DDD.currentY \ DXP, StrPtr(buf$), Len(buf$), 1, tabw, DDD.currentX \ DXP))
                 
                 End Select

                 
                 End If
                 
                  '  If Not nopr Then fullPlain ddd, mybasket, RTrim$(Buf$), w2, nowait, nonewline
                 End If

        Else
            skip = skip - 1
            If Len(buf$) > 0 Then para(i) = Mid$(para(i), MyTrimL3Len(para(i)) + 1)
            buf$ = para(i)
            para(i) = vbNullString
last:
        If skip >= 0 Then GoTo continue
        If Hi = 0 And frmt = 0 Then
        If Not scrollme Then
        If Not nopr Then
    
            MyPrintNew DDD, mybasket.uMineLineSpace, buf$, , nowait     ';   '************************************************************************************
    
            res = DDD.currentX
                    If Trim$(buf$) = vbNullString Then
                    DDD.currentX = ((DDD.currentX + .Xt \ 2) \ .Xt) * .Xt
                    Else
                    DDD.currentX = ((DDD.currentX + .Xt \ 1.2) \ .Xt) * .Xt
                    End If
            End If
            If Not (isAcolumn Or nosettext) Then If Not nonewline Then GoTo JUMPHERE
            Exit Sub
        End If
        If Not nopr Then
            fullPlainWhere DDD, mybasket, buf$, w2, frmt, nowait, nonewline
        End If
        Else
        frmt = Abs(frmt)
        If frmt > 0 Then frmt = frmt Mod 4
        If frmt >= 0 Then
                If Not nopr Then
                INTD = 0
                Extra = 0

                
                Select Case frmt
                Case 1
                buf$ = RTrim(buf$)
                 Extra = wi \ DXP - LowWord(GetTabbedTextExtent(DDD.hdc, StrPtr(buf$), Len(buf$), 1, tabw))
                 
                 Extra = LowWord(TabbedTextOut(DDD.hdc, DDD.currentX \ DXP + Extra, DDD.currentY \ DXP, StrPtr(buf$), Len(buf$), 1, tabw, DDD.currentX \ DXP + Extra))
                
                Case 2
                buf$ = Trim(buf$)
                Extra = (wi \ DXP - LowWord(GetTabbedTextExtent(DDD.hdc, StrPtr(buf$), Len(buf$), 1, tabw))) \ 2
                 
                Extra = LowWord(TabbedTextOut(DDD.hdc, DDD.currentX \ DXP + Extra, DDD.currentY \ DXP, StrPtr(buf$), Len(buf$), 1, tabw, DDD.currentX \ DXP + Extra))

                Case 3, 0
                INTD = TextWidth(DDD, space$(MyTrimL3Len(buf$)))
                If INTD > 0 Then
                    buf$ = Mid$(buf$, MyTrimL3Len(buf$) + 1)
                    DDD.currentX = DDD.currentX + INTD
                End If
                buf$ = RTrim(buf$)
                
                Extra = LowWord(TabbedTextOut(DDD.hdc, DDD.currentX \ DXP, DDD.currentY \ DXP, StrPtr(buf$), Len(buf$), 1, tabw, DDD.currentX \ DXP))
                
                End Select
                End If
                
         End If
        End If
   
        If collectit Then
        mDoc.AppendParagraphOneLine Trim$(buf$)
        End If
             End If
                If isAcolumn Then Exit Sub
                'If UBound(para) = i Then Exit For
        If skip < 0 Or scrollme Then
JUMPHERE:
            lCount = lCount + 1
            npy = npy + 1
            If npy >= .my And scrollme Then
                If Not meta Then
                    If Not nopr Then
                        If SCRnum2stop > 0 Then
                            If lCount >= SCRnum2stop Then
                                If Not bstack.toprinter Then
                                    If Not nowait Then
                                        If Not nopage Then
                                            DDD.Refresh
                                            Do
                                                'mywait bstack, 10
                                                If TaskMaster Is Nothing Then
                                                Sleep 10
                                                ElseIf Not TaskMaster.Processing And TaskMaster.QueueCount = 0 Then
                                                Sleep 10
                                                End If
                                                MyDoEventsNoRefresh
                                                
                                            Loop Until INKEY$ <> "" Or mouse <> 0 Or NOEXECUTION
                                        End If
                                     End If
                                End If
                                SCRnum2stop = .pageframe
                                lCount = 1
                            End If
                        End If
                        If Not bstack.toprinter Then
                            DDD.Refresh
                            ScrollUpNew DDD, mybasket
                        Else
                            getnextpage
                            npy = 1
                        End If
                    End If
                    npy = npy - 1
                ElseIf npy >= .my Then
                    If Not nopr Then crNew bstack, mybasket
                    npy = npy - 1
                End If
                End If
            End If
If Not nopr Then LCTbasket DDD, mybasket, npy, PX: DDD.currentX = DDD.currentX + dv2x15
If skip < 0 Then Hi = Hi - 1
If Hi < 0 Then Exit For
continue:
     If Len(para(i)) > 0 Then GoTo nextline
    Next i
End With
finish:
If scrollme Then
HoldReset lCount, mybasket
End If
res = nohi - Hi
wi = DDD.currentX
If collectit Then bstack.soros.PushStr mDoc.textDoc
End Sub
Public Sub EnableMidi()
Dim curDevice As Long, rc As Long

 If hmidi = 0 Then
rc = GetFuncPtr("winmm.dll", "midiOutOpen")
If rc <> 0 Then
    rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
    If (rc <> 0) Then
       MyEr "Couldn't open midi device - Error #" & rc, "Δεν μπορώ να ανοίξω κανάλι Midi - λάθος #"" & rc"
    End If
    End If
    End If
End Sub
Public Sub instrument(insID As Long, Channel As Long)
EnableMidi
Dim midimsg As Long
    midimsg = (insID * 256) + &HC0 + Channel
    midiOutShortMsg hmidi, midimsg
End Sub
Public Sub DisableMidi()
  If hmidi <> 0 Then
  midiOutClose (hmidi)
  hmidi = 0
  End If
End Sub




