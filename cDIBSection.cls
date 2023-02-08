VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public emfSizeFactor
Private Const DT_NOPREFIX = &H800
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" _
        (ByVal hDC As Long, _
        ByVal hBrush As Long, _
        ByVal lpDrawStateProc As Long, _
        ByVal lParam As Long, _
        ByVal wParam As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal fuFlags As Long) As Long

      Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" _
          (ByVal hDC As Long, lpInitData As Any) As Long

      Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
          (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
          ByVal lpOutput As Long, lpInitData As Any) As Long

      Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) _
          As Long
      Private Type DOCINFO
          cbSize As Long
          lpszDocName As String
          lpszOutput As String
      End Type
      Private Declare Function StartDoc Lib "gdi32" Alias "StartDocA" _
          (ByVal hDC As Long, lpdi As DOCINFO) As Long

      Private Declare Function StartPage Lib "gdi32" (ByVal hDC As Long) _
          As Long

      Private Declare Function EndDoc Lib "gdi32" (ByVal hDC As Long) _
          As Long

      Private Declare Function EndPage Lib "gdi32" (ByVal hDC As Long) _
          As Long
Private last_zf As Single
Private Const Pi = 3.14159265359
Public Event InitProgress(ByVal lMax As Long)
Public Event Progress(ByVal lPosition As Long)
Public Event Complete(ByVal lTimeMs As Long)
Const PS_NULL = 5
Const PS_SOLID = 0
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByVal lppt As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As Long, ByVal lpOutput As String, pInitData As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GdiFlush Lib "gdi32" () As Long
Private SD As String

Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

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

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(256) As RGBQUAD
End Type
Private Type SIZEL
cx As Long
cy As Long
End Type

Private Type ENHMETAHEADER
iType As Long
nSize As Long
rclBounds As RECT
rclFrame As RECT
dSignature As Long
nVersion As Long
nBytes As Long
nRecords As Long
nHandles As Integer
sReserved As Integer
nDescription As Long
offDescription As Long
nPalEntries As Long
szlDevice As SIZEL
szlMillimeters As SIZEL
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hDC As Long, _
    pBitmapInfo As Any, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal Handle As Long, _
    ByVal dw As Long) As Long

Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Const COLORONCOLOR = 3
Private Const HALFTONE = 4

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Const DIB_PAL_COLORS = 1
Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function timeGetTime Lib "kernel32.dll" Alias "GetTickCount" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

' Clipboard functions:
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long

Private Const CF_BITMAP = 2
Private Const CF_DIB = 8
Private m_fontname As String
Private m_fontsize As Single
Private m_currentX As Long
Private m_currenty As Long
Private m_PageStyle As Long
Private m_transparent As Long
Private m_fonttransparent As Long
Private m_BackColor As Long
' Handle to the current DIBSection:
Private m_hDIb As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
Private m_PrinterhBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC As Long
Private p_hdc As Long

' Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_tBI As BITMAPINFO
'george karras the great
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Const RASTERCAPS As Long = 38
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113
Private Const HORZRES As Long = 8
Private Const VERTRES As Long = 10
Private Const PHYSICALHEIGHT As Long = 111
Private Const PHYSICALWIDTH As Long = 110
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
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
Private Enum ICAP_COMPRESSION
TWCP_NONE = 0
TWCP_PACKBITS = 1
TWCP_GROUP31D = 2
TWCP_GROUP31DEOL = 3
TWCP_GROUP32D = 4
TWCP_GROUP4 = 5
TWCP_JPEG = 6
TWCP_LZW = 7
TWCP_JBIG = 8
End Enum
Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
    End Type
Private Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private m_logx As Single
Private m_logy As Single
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFaceName As String * 33
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const OPAQUE = 2
Private PaintType As Boolean
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long
Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hEmf As Long, lpRect As Any) As Long
Private Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hmf As Long, ByVal nSize As Long, lpvData As Any) As Long
Private Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As Long
Private Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As Any) As Long
Private mType As Long
Private wmfemf As StdPicture
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_DEFAULTSIZE = &H8
Const DI_NORMAL = DI_MASK Or DI_IMAGE Or DI_DEFAULTSIZE
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private OpenPrinter As Boolean
Public Property Get PrinterHdc() As Long
PrinterHdc = p_hdc
p_hdc = 0
OpenPrinter = False
End Property

Public Sub CopyPrinter(A As Long)
p_hdc = A
OpenPrinter = True
End Sub

Public Property Get BitmapType() As Long
BitmapType = mType
End Property
Public Sub ResetBitmapTypeToBITMAP()
Set wmfemf = Nothing
mType = 4
End Sub

Public Sub Cls(Optional ByVal clearcolor = &HFFFFFF)
square 0, 0, Width - 1, Height - 1, clearcolor
End Sub
Private Function getRGB(bckColor As Variant) As Long
Dim br As Byte, bg As Byte, bbb As Byte, ba$
If VarType(bckColor) <> vbString Then ba$ = Hex$(bckColor) Else ba$ = bckColor
ba$ = Right$("00000" & ba$, 6)
getRGB = val("&h" + Mid$(ba$, 1, 2) + Mid$(ba$, 3, 2) + Mid$(ba$, 5, 2) + "&")
End Function
Private Function getRGBquad(rgb As RGBQUAD) As Long
getRGBquad = cUlng(rgb.rgbRed + rgb.rgbGreen * 256@ + rgb.rgbBlue * 256 * 256@)

End Function
Private Function cUlng(ByVal A As Currency) As Long ' πέρνει έναν umsign integer και τον κάνει νορμάλ χωρίς αλλαγή των bits
On Error GoTo cu1
A = Abs(Int(A))
If A > 2147483647@ Then
cUlng = A - 4294967296@
Else
cUlng = CLng(A)
End If
Exit Function
cu1:
cUlng = 0
End Function

Public Property Get BackColor() As Variant
BackColor = m_BackColor
End Property

Public Property Let BackColor(bck As Variant)
m_BackColor = getRGB(bck)
End Property

Public Property Get FontName() As String
FontName = m_fontname
End Property

Public Property Let FontName(FontNameS As String)
m_fontname = FontNameS
End Property
Public Property Get FontSize() As Variant
FontSize = m_fontsize
End Property

Public Property Let FontSize(fsize As Variant)
m_fontsize = fsize
End Property

Public Property Get Transparent() As Variant
Transparent = m_transparent
End Property

Public Property Let Transparent(trans As Variant)
m_transparent = trans
End Property
Public Property Get PageStyle() As Variant
PageStyle = m_PageStyle
End Property

Public Property Let PageStyle(pStyle As Variant)
m_PageStyle = pStyle
End Property
Public Property Get FontTransparent() As Boolean

FontTransparent = m_fonttransparent = 1
End Property

Public Property Let FontTransparent(trans As Boolean)
If trans Then
m_fonttransparent = 1
Else
m_fonttransparent = 2
End If
End Property
Public Property Get dpix() As Long
dpix = m_tBI.bmiHeader.biXPelsPerMeter * 0.0254
End Property
Public Property Get dpiy() As Long
dpiy = m_tBI.bmiHeader.biYPelsPerMeter * 0.0254
End Property
Public Property Get DPMX() As Long
DPMX = m_tBI.bmiHeader.biXPelsPerMeter
End Property
Public Property Get DPMY() As Long
DPMY = m_tBI.bmiHeader.biYPelsPerMeter
End Property
Public Property Get cmX() As Single
If m_tBI.bmiHeader.biXPelsPerMeter = 0 Then
GetDpi
End If
cmX = Me.Width / m_tBI.bmiHeader.biXPelsPerMeter * 100
End Property
Public Property Get cmy() As Single
If m_tBI.bmiHeader.biYPelsPerMeter = 0 Then
GetDpi
End If
cmy = Me.Height / m_tBI.bmiHeader.biYPelsPerMeter * 100
End Property
Public Sub GetDpiDIB(other As cDIBSection)
m_tBI.bmiHeader.biXPelsPerMeter = other.DPMX
m_tBI.bmiHeader.biYPelsPerMeter = other.DPMY
End Sub
Public Sub GetDpi(Optional ByVal Xdpi As Long = -1, Optional ByVal Ydpi As Long = -1)
If Xdpi <= 0 Then Xdpi = 1440 / Screen.TwipsPerPixelX
If Ydpi <= 0 Then Ydpi = 1440 / Screen.TwipsPerPixelY

With m_tBI.bmiHeader
       .biXPelsPerMeter = CLng(0.5 + (Xdpi) / 0.0254)
       .biYPelsPerMeter = CLng(0.5 + (Ydpi) / 0.0254)
 
End With
End Sub

Public Property Get bitsPerPixel() As Single
bitsPerPixel = m_tBI.bmiHeader.biBitCount
End Property
Public Property Get packet_bytes() As Long
packet_bytes = Me.header_bytes + Me.BytesPerScanLine * Me.Height
 End Property
Public Property Get header_bytes() As Long
Dim BITS As Long, uu As Long, cltused As Long
BITS = m_tBI.bmiHeader.biBitCount
       If BITS < 16 Then

       cltused = m_tBI.bmiHeader.biClrUsed
            If cltused = 0 Then
            uu = 40 + 4 * 2 ^ BITS
            Else
            uu = 40 + cltused * 4
            End If
    Else
    uu = 40
    End If
     header_bytes = uu
 End Property
 

Public Property Get LastZoom() As Single
LastZoom = last_zf
End Property
Public Function CopyToClipboard( _
        Optional ByVal bAsDIB As Boolean = True _
    ) As Boolean
Dim lhDCDesktop As Long
Dim lHDC As Long
Dim lhBmpOld As Long
Dim hObj As Long
Dim lFmt As Long
Dim b() As Byte
Dim tbi As BITMAPINFO
Dim lPtr As Long
Dim hDibCopy As Long

    lhDCDesktop = GetDC(GetDesktopWindow())
    If (lhDCDesktop <> 0) Then
        lHDC = CreateCompatibleDC(lhDCDesktop)
        If (lHDC <> 0) Then
            If (bAsDIB) Then
            lFmt = CF_DIB
             If (OpenClipboard(0) <> 0) Then
                        EmptyClipboard
                SentToClipBoard
                CloseClipboard
              End If
               
            Else

                hObj = CreateCompatibleBitmap(lhDCDesktop, Width, Height)
                If (hObj <> 0) Then
                    lhBmpOld = SelectObject(lHDC, hObj)
                    PaintPicture lHDC
                    SelectObject lHDC, lhBmpOld
                    lFmt = CF_BITMAP
                     If (OpenClipboard(0) <> 0) Then
                        EmptyClipboard
                        If (SetClipboardData(lFmt, hObj) <> 0) Then
                            CopyToClipboard = True
                        End If
                        CloseClipboard
                    End If
                End If
            End If
            DeleteDC lHDC
        End If
        ReleaseDC GetDesktopWindow(), lhDCDesktop
    End If
End Function
Private Function CreateDIB2( _
        ByVal lHDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean
    On Error GoTo XX
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
        .biXPelsPerMeter = CLng(0.5 + (1440 / Screen.TwipsPerPixelX) / 0.0254)
        .biYPelsPerMeter = CLng(0.5 + (1440 / Screen.TwipsPerPixelY) / 0.0254)
       .biClrUsed = 0
      .biClrImportant = 0
    End With
'
'GdiFlush
If m_lPtr <> 0 Then
ClearUp
Else
CreateDIB2 = False
End If
    hDib = CreateDIBSection( _
            lHDC, _
             m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)

    CreateDIB2 = (hDib <> 0)
XX:
End Function
Public Function CreateDIB( _
        ByVal lHDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean
    On Error GoTo XX
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
        .biXPelsPerMeter = CLng(0.5 + (1440 / Screen.TwipsPerPixelX) / 0.0254)
        .biYPelsPerMeter = CLng(0.5 + (1440 / Screen.TwipsPerPixelY) / 0.0254)
       .biClrUsed = 0
      .biClrImportant = 0
    End With
'
'GdiFlush
If m_lPtr <> 0 Then
ClearUp
Else
CreateDIB = False
End If
    hDib = CreateDIBSection( _
            lHDC, _
             m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)

    CreateDIB = (hDib <> 0)
XX:
End Function
Public Function CreateFromDIB(ByVal mib As Long) As Boolean
Dim BITS As Long, uu As Long
Dim lhDCDesktop As Long
needHDC
        If (m_hDC <> 0) Then
       CopyMemory m_tBI, ByVal mib, Me.header_bytes
 
 
          m_hDIb = CreateDIBSection( _
            m_hDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)

            If m_hDIb <> 0 Then
            CopyMemory ByVal m_lPtr, ByVal mib + Me.header_bytes, Me.BytesPerScanLine * Me.Height
           
 
            CreateFromDIB = True

        End If

        End If
              FreeHDC


    
End Function
Public Function CopyPicture(ByRef pict As Object)
' pict in pixels
On Error Resume Next
Dim osm As Long
osm = pict.ScaleMode
pict.ScaleMode = vbPixels
Me.ClearUp
    If create(pict.ScaleWidth, pict.ScaleHeight) Then
    
      LoadPictureBlt pict.hDC
      If Err.Number > 0 Then
      
      Me.RandomiseBits
      End If
    End If
      pict.ScaleMode = osm
End Function
Public Function CopyPicturePrinter(ByRef pict As Object)
' pict in pixels
On Error Resume Next
Dim osm As Long
osm = pict.ScaleMode
pict.ScaleMode = vbPixels
LoadPictureBlt pict.hDC
pict.ScaleMode = osm
End Function

Public Function create( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
    needHDC
      create = CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)
    FreeHDC

        'GdiFlush
End Function
Friend Function create32( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
    needHDC
      create32 = CreateDIB2(m_hDC, lWidth, lHeight, m_hDIb)
      GetDpi 96, 96
    FreeHDC
        'GdiFlush
End Function
Public Property Get BytesPerScanLine() As Long
    ' Scans must align on dword boundaries:
    Dim A As Long, b As Long

    A = m_tBI.bmiHeader.biWidth * m_tBI.bmiHeader.biBitCount
    b = A Mod 32
    If b > 0 Then b = 32 - b
    
    BytesPerScanLine = (A + b) \ 8
    
   
End Property


Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt( _
        ByVal lHDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    GdiFlush
 needHDC
 BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lHDC, lSrcLeft, lSrcTop, eRop
FreeHDC

End Sub

Public Sub LoadPictureStretchBlt( _
        ByVal lHDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    If (lSrcWidth < 0) Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If (lSrcHeight < 0) Then lSrcHeight = m_tBI.bmiHeader.biHeight
 ' ***********here********************
    Dim PrevStretchMode As Long
 needHDC
PrevStretchMode = SetStretchBltMode(m_hDC, COLORONCOLOR)
If PrevStretchMode Then
StretchBlt m_hDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, lHDC, lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight, eRop
SetStretchBltMode m_hDC, PrevStretchMode
End If
FreeHDC
End Sub
Public Sub LoadPictureStretchBltH( _
        ByVal lHDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    If (lSrcWidth < 0) Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If (lSrcHeight < 0) Then lSrcHeight = m_tBI.bmiHeader.biHeight
 ' ***********here********************
 needHDC
 Dim PrevStretchMode As Long
PrevStretchMode = SetStretchBltMode(m_hDC, HALFTONE)
If PrevStretchMode Then
If SetBrushOrgEx(m_hDC, 0, 0, (0)) Then
StretchBlt m_hDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, lHDC, lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight, eRop
SetBrushOrgEx m_hDC, 0, 0, (0)
End If
SetStretchBltMode m_hDC, PrevStretchMode
End If
FreeHDC
End Sub

Public Sub PaintPicture( _
        ByVal lHDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    GdiFlush
     needHDC
    BitBlt lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
FreeHDC
End Sub

Public Sub StretchPicture( _
        ByVal lHDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy, _
        Optional ratio As Double = 1#)
    If m_logx = 0 Then m_logx = 1
    If m_logy = 0 Then m_logy = 1
     If (lDestWidth = -1) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight = -1) Then lDestHeight = m_tBI.bmiHeader.biHeight
    If (lSrcWidth = -1) Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If (lSrcHeight = -1) Then lSrcHeight = m_tBI.bmiHeader.biHeight
Dim PrevStretchMode As Long
    Dim D As Long
needHDC
  PrevStretchMode = SetStretchBltMode(lHDC, COLORONCOLOR) 'coloroncolor
  If PrevStretchMode Then
  If SetBrushOrgEx(lHDC, 0, 0, (0)) Then
  If ratio = 1 Then
  D = StretchDIBits(lHDC, lDestLeft / m_logx, lDestTop / m_logy, lDestWidth / m_logx, lDestHeight / m_logy, lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight, ByVal m_lPtr, m_tBI, DIB_RGB_COLORS, eRop)

  ElseIf ratio > 1 Then
  D = StretchDIBits(lHDC, lDestLeft / m_logx, lDestTop / m_logy, lDestWidth / m_logx, lDestHeight / m_logy, lSrcLeft + (lSrcWidth - lSrcWidth / ratio) / 2, lSrcTop, lSrcWidth / ratio, lSrcHeight, ByVal m_lPtr, m_tBI, DIB_RGB_COLORS, eRop)
  
  Else
    D = StretchDIBits(lHDC, lDestLeft / m_logx, lDestTop / m_logy, lDestWidth / m_logx, lDestHeight / m_logy, lSrcLeft, lSrcTop + (lSrcHeight - lSrcHeight * ratio) / 2, lSrcWidth, lSrcHeight * ratio, ByVal m_lPtr, m_tBI, DIB_RGB_COLORS, eRop)
  End If
    SetBrushOrgEx lHDC, 0, 0, (0)
    End If
    SetStretchBltMode m_hDC, PrevStretchMode
    End If
FreeHDC
End Sub
Public Sub StretchPictureH( _
        ByVal lHDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy, _
        Optional ratio As Double = 1#)
    If m_logx = 0 Then m_logx = 1
    If m_logy = 0 Then m_logy = 1
     If (lDestWidth = -1) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight = -1) Then lDestHeight = m_tBI.bmiHeader.biHeight
    If (lSrcWidth = -1) Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If (lSrcHeight = -1) Then lSrcHeight = m_tBI.bmiHeader.biHeight
Dim PrevStretchMode As Long
    Dim D As Long
needHDC
  PrevStretchMode = SetStretchBltMode(lHDC, HALFTONE)
  If PrevStretchMode Then
If SetBrushOrgEx(m_hDC, 0, 0, (0)) Then
  If ratio = 1 Then
  D = StretchDIBits(lHDC, lDestLeft / m_logx, lDestTop / m_logy, lDestWidth / m_logx, lDestHeight / m_logy, lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight, ByVal m_lPtr, m_tBI, DIB_RGB_COLORS, eRop)

  ElseIf ratio > 1 Then
  D = StretchDIBits(lHDC, lDestLeft / m_logx, lDestTop / m_logy, lDestWidth / m_logx, lDestHeight / m_logy, lSrcLeft + (lSrcWidth - lSrcWidth / ratio) / 2, lSrcTop, lSrcWidth / ratio, lSrcHeight, ByVal m_lPtr, m_tBI, DIB_RGB_COLORS, eRop)
  
  Else
    D = StretchDIBits(lHDC, lDestLeft / m_logx, lDestTop / m_logy, lDestWidth / m_logx, lDestHeight / m_logy, lSrcLeft, lSrcTop + (lSrcHeight - lSrcHeight * ratio) / 2, lSrcWidth, lSrcHeight * ratio, ByVal m_lPtr, m_tBI, DIB_RGB_COLORS, eRop)
    End If
    SetBrushOrgEx lHDC, 0, 0, (0)
    End If
    SetStretchBltMode m_hDC, PrevStretchMode
    End If
FreeHDC
End Sub


Public Property Get hDib() As Long
    hDib = m_hDIb
End Property
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property
Public Sub RandomiseBits( _
        Optional ByVal bGray As Boolean = False _
    )
Dim bDib() As Byte
Dim X As Long, Y As Long, z As Long
Dim lc As Long
Dim tSA As SAFEARRAY2D
Dim xend As Long
Dim mStep As Long, cStep
mStep = bitsPerPixel \ 8
If mStep > 3 Then cStep = 3 Else cStep = mStep
    ' Get the bits in the from DIB section:
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

    ' random:

    
    xend = (Width - 1) * mStep
    If (bGray) Then
        For Y = 0 To m_tBI.bmiHeader.biHeight - 1
            For X = 0 To xend Step mStep
                lc = Rnd * 255
                For z = 0 To cStep
                bDib(X + z, Y) = lc
                Next z
            Next X
        Next Y
    Else
        For X = 0 To xend Step mStep
            For Y = 0 To m_tBI.bmiHeader.biHeight - 1
            
                bDib(X, Y) = 0
                For z = 1 To cStep
                bDib(X + 1, Y) = Rnd * 255
                Next z
            Next Y
        Next X
    End If
    
    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
End Sub

Public Sub WhiteBits()
square 0, 0, Width - 1, Height - 1, "ffffff"
End Sub
Public Sub ClearBits(m_color As Long)
Dim old_brush As Long, my_brush As Long
Dim mr As RECT
my_brush = CreateSolidBrush(m_color)
mr.Right = Width
mr.Bottom = Height
needHDC
FillRect m_hDC, mr, my_brush
FreeHDC
DeleteObject my_brush
End Sub
Public Sub ClearUp()
GdiFlush
If m_hDIb <> 0 Then
 If m_lPtr > 0 Then GlobalFree m_lPtr
 If DeleteObject(m_hDIb) <> 1 Then MyEr "Problem with GDI objects", "Πρόβλημα με το σύστημα γραφικών"

        m_tBI.bmiHeader.biHeight = 0
         ' end new
    m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0

End If
Set wmfemf = Nothing
End Sub
Public Sub needHDC()
' return hdc
If m_hDC <> 0 Then Exit Sub
Dim lhDCDesktop As Long
   lhDCDesktop = CreateDC(SD, 0, 0, ByVal 0)
  m_hDC = CreateCompatibleDC(lhDCDesktop)
If m_hDIb <> 0 And m_hDC <> 0 Then m_hBmpOld = SelectObject(m_hDC, m_hDIb)

    DeleteDC lhDCDesktop
    sumhDC = sumhDC + 1
End Sub
Public Sub FreeHDC()
GdiFlush
If m_hDC <> 0 Then
If m_hDIb <> 0 And m_hBmpOld <> 0 Then SelectObject m_hDC, m_hBmpOld
sumhDC = sumhDC - 1
DeleteDC m_hDC
m_hDC = 0
End If
End Sub

Private Sub MakePrinterhDC(dest As Object, PAGENAME$)
Dim di As DOCINFO
        
          Dim ret As Long
          Dim LastError As Long
          Dim Dummy3 As Object
            If UBound(MyDM) = 1 Then
            ShowProperties Dummy3, dest.DeviceName, MyDM
            End If
          p_hdc = CreateDC(dest.DriverName, dest.DeviceName, 0, MyDM(1))
          di.cbSize = Len(di)   ' 20
          di.lpszDocName = PAGENAME$
          ret = StartDoc(p_hdc, di)
          
 End Sub


Public Function Resample( _
        ByVal lNewHeight As Long, _
        ByVal lNewWidth As Long _
    ) As cDIBSection
Dim cDib As New cDIBSection
    If cDib.create(lNewWidth - 1, lNewHeight - 1) Then
        If (lNewWidth <> m_tBI.bmiHeader.biWidth) Or (lNewHeight <> m_tBI.bmiHeader.biHeight) Then
            ' Change in size, do resample:
                             needHDC
           cDib.LoadPictureStretchBltH m_hDC, , , lNewWidth, lNewHeight, , , Width, Height
        FreeHDC
      
        Else
            ' No size change so just return a copy:
             
             needHDC
            cDib.LoadPictureBlt m_hDC
        FreeHDC
        End If
        
    End If
    Set Resample = cDib
End Function

Private Sub swaplong(A As Long, b As Long)
Dim c As Long
c = A: A = b: b = c
End Sub

Public Sub ThumbnailPaint(DestPic As Object, Optional ByVal zf As Single = -1, Optional ByVal FitToPic As Boolean = True, Optional CROP As Boolean = False, Optional over As Boolean = False, Optional ByVal fr As Long = 0, Optional ByVal offsetX As Long = 0, Optional ByVal offsetY As Long, Optional ByVal Border As Long = 0, Optional ByVal rtxy As Double = 1#)
   Dim oldSc As Long, Dummy1 As Boolean, Dummy2 As Boolean, sw As Long, sh As Long
   Dim swProbe As Long, shProbe As Long, osh As Double
   Dim Myleft As Long, mytop As Long, mywidth As Long, myheight As Long
   
     Dim A As Single
         oldSc = DestPic.ScaleMode
   DestPic.ScaleMode = vbPixels

   
 
  If over Then
   swProbe = DestPic.ScaleX(DestPic.Width, 1, 3)
   shProbe = DestPic.ScaleY(DestPic.Height, 1, 3)
   Else
   swProbe = DestPic.ScaleWidth
   shProbe = DestPic.ScaleHeight
   End If
    sw = DestPic.ScaleWidth
   sh = DestPic.ScaleHeight
 If sh * sw = 0 Then DestPic.ScaleMode = oldSc: Exit Sub
   
   Dim osw As Long
   
If zf <= 0 Then

      
        If Not FitToPic Then
            StretchPicture DestPic.hDC, (sw - Width) / 2 + offsetX, (sh - Height) / 2 + offsetY, , , , , , , , rtxy
        Else
      If sw < sh Then
    osw = offsetX: offsetX = -offsetY:
    offsetY = osw
    
     rtxy = 1 / rtxy
   End If
            If CROP Then
                   
                    sh = sh - 2 * fr
                    sw = sw - 2 * fr
                    If shProbe < swProbe Then
                    osh = shProbe / sh
                        shProbe = CLng(sw * shProbe / swProbe) - Border * 2 / osh
                        swProbe = sw - Border * 2 * osh
                   Else
                   osh = swProbe / sw
            
                        swProbe = CLng(sh * swProbe / shProbe) - 2 * Border / osh
                        shProbe = sh - 2 * Border * osh
                   End If
                   ' Τελευταία αλλαγή - έφυγαν τα CLNG(border/osh) από τα x,y για να τυπώνει σωστά τα offset
                 If (swProbe / Width) > (shProbe / Height) Then
                    zf = Width / swProbe
                    If swProbe > shProbe Then
                        StretchPicture DestPic.hDC, CLng(fr) + offsetX + CLng(Border * osh), CLng((sh - sw * shProbe / swProbe) / 2 + fr) + offsetY + Border, sw - Border * 2, CLng(sw * shProbe / swProbe) - Border * 2, 0, CLng((Height - shProbe * zf)) / 2, Width, CLng(shProbe * zf), , rtxy
                    Else
                        StretchPicture DestPic.hDC, CLng(fr) + CLng((sw - sh * swProbe / shProbe) / 2) + offsetX + CLng(Border / osh), CLng(fr) + offsetY + Border, CLng(sh * swProbe / shProbe) - 2 * Border, sh - 2 * Border, 0, CLng((Height - shProbe * zf)) / 2, Width, CLng(shProbe * zf), , rtxy
                    End If
                Else
                    zf = Height / shProbe
                    If swProbe > shProbe Then
                     StretchPicture DestPic.hDC, CLng(fr) + offsetX + CLng(Border * osh), CLng((sh - sw * shProbe / swProbe) / 2 + fr) + offsetY + Border, sw - Border * 2, CLng(sw * shProbe / swProbe) - Border * 2, (Width - swProbe * zf) / 2, 0, CLng(swProbe * zf), Height, , rtxy
                    Else
                     StretchPicture DestPic.hDC, CLng(fr) + CLng((sw - sh * swProbe / shProbe) / 2) + offsetX + Border, CLng(fr) + offsetY + CLng(Border * osh), CLng(sh * swProbe / shProbe) - 2 * Border, sh - 2 * Border, (Width - swProbe * zf) / 2, 0, CLng(swProbe * zf), Height, , rtxy
                 End If
                  End If
             Else
                If (swProbe / Width) < (shProbe / Height) Then
                    zf = sw / Width
                Else
                    zf = sh / Height
                End If
                StretchPicture DestPic.hDC, (sw - Width * zf) / 2 + offsetX, (sh - Height * zf) / 2 + offsetY, CLng(Width * zf), CLng(Height * zf), , , , , , rtxy
            End If
        End If
    Else
        zf = zf / 100#
        If Not FitToPic Then
        'offsetx = (DestPic.ScaleWidth - width * zf) * (offsetx / 100#)
        'offsety = (DestPic.ScaleHeight - height * zf) * (offsety / 100#)
        
            StretchPicture DestPic.hDC, offsetX + Abs((DestPic.ScaleWidth - Width * zf) / 2), offsetY + Abs((DestPic.ScaleHeight - Height * zf) / 2), Width * zf, Height * zf, , , , , , rtxy
                    Else
            If DestPic.ScaleWidth / Width < DestPic.ScaleHeight / Height Then
                zf = zf * DestPic.ScaleWidth / Width
                
            Else
                zf = zf * DestPic.ScaleHeight / Height
            End If
            On Error Resume Next
     '       offsetx = offsetx * zf
        '    offsety = offsety * zf
                StretchPicture DestPic.hDC, offsetX + (DestPic.ScaleWidth - Width * zf) / 2, offsetY + (DestPic.ScaleHeight - Height * zf) / 2, Width * zf, Height * zf, , , , , , rtxy
        End If
    End If
  ' get out this...produce flickering in vb6
  '  UpdateWindow DestPic.hwnd 'DestPic.Refresh
  
    DestPic.ScaleMode = oldSc
    last_zf = zf
    ' gdiflush
End Sub
Public Sub EndPrint()
       If OpenPrinter Then
            Dim ret As Long
            GdiFlush
           ' MsgBox "end print " & m_lPtr
            ret = EndDoc(p_hdc)
           ' MsgBox "end enddoc " & ret
            ret = DeleteDC(p_hdc)
            'MsgBox "end deletedc " & ret
            p_hdc = 0
        OpenPrinter = False
    End If
End Sub
Public Sub ResetPageDM()
 If p_hdc <> 0 Then
 ResetDC p_hdc, MyDM(1)
 
 End If
End Sub
Property Get PrinterOn()
    PrinterOn = OpenPrinter
End Property
Public Sub ThumbnailPaintPrinter(copiesNUM As Long, Optional ByVal zf As Single = -1, Optional ByVal FitToPic As Boolean = True, Optional CROP As Boolean = False, Optional over As Boolean = False, Optional ByVal fr As Long = 0, Optional ByVal offsetX As Long = 0, Optional ByVal offsetY As Long, Optional ByVal Border As Long = 0, Optional ByVal rtxy As Double = 1#, Optional pname As String = vbNullString)
   Dim oldSc As Long, Dummy1 As Boolean, Dummy2 As Boolean, sw As Long, sh As Long
   Dim swProbe As Long, shProbe As Long, osh As Double
   Dim Myleft As Long, mytop As Long, mywidth As Long, myheight As Long
   Dim lox As Long, loy As Long
   
     Dim A As Single
   '  Exit Sub
If Not OpenPrinter Then
   CopiesCount copiesNUM, MyDM()
   MakePrinterhDC Printer, pname
   OpenPrinter = p_hdc <> 0
End If

   If p_hdc <> 0 Then
   m_PrinterhBmpOld = SelectObject(p_hdc, m_hDIb)
   StartPage p_hdc
   Else
    Beep
    Exit Sub
   End If

   If over Then
   swProbe = GetDeviceCaps(p_hdc, PHYSICALWIDTH) 'DestPic.ScaleX(DestPic.width, 1, 3)
   shProbe = GetDeviceCaps(p_hdc, PHYSICALHEIGHT) ' DestPic.ScaleY(DestPic.height, 1, 3)
   Else
   swProbe = GetDeviceCaps(p_hdc, HORZRES) 'sw
   shProbe = GetDeviceCaps(p_hdc, VERTRES) ' sh
   End If
    sw = GetDeviceCaps(p_hdc, HORZRES)
   sh = GetDeviceCaps(p_hdc, VERTRES)
      lox = GetDeviceCaps(p_hdc, LOGPIXELSX)
   loy = GetDeviceCaps(p_hdc, LOGPIXELSY)
If lox > loy Then
m_logy = 1
m_logx = loy / lox
ElseIf loy > lox Then
m_logy = lox / loy
m_logx = 1
End If
swProbe = swProbe * m_logx
shProbe = shProbe * m_logy
sw = sw * m_logx
sh = sh * m_logy
   Dim osw As Long
 If rtxy = 0 Then rtxy = 1
   If sw < sh Then
    'osw = offsetx: offsetx = -offsety:
    'offsety = -osw
    rtxy = 1 / rtxy
   End If
    ' offsety = 0
  
  ' End If
   
    If zf <= 0 Then
        If Not FitToPic Then
            StretchPicture p_hdc, (sw - Width) / 2 + offsetX, ((sh - Height) / 2 + offsetY), , , , , , , , rtxy
        Else
            If CROP Then
                   
                    sh = sh - 2 * fr
                    sw = sw - 2 * fr
                    If shProbe < swProbe Then
                    osh = shProbe / sh
                        shProbe = CLng(sw * shProbe / swProbe) - Border * 2 / osh
                        swProbe = sw - Border * 2 * osh
                   Else
                   osh = swProbe / sw
            
                        swProbe = CLng(sh * swProbe / shProbe) - 2 * Border / osh
                        shProbe = sh - 2 * Border * osh
                   End If
                   ' Τελευταία αλλαγή - έφυγαν τα CLNG(border/osh) από τα x,y για να τυπώνει σωστά τα offset
                 If (swProbe / Width) > (shProbe / Height) Then
                    zf = Width / swProbe
                    If swProbe > shProbe Then
                        StretchPicture p_hdc, CLng(fr) + offsetX + CLng(Border * osh), CLng((sh - sw * shProbe / swProbe) / 2 + fr) + offsetY + Border, sw - Border * 2, CLng(sw * shProbe / swProbe) - Border * 2, 0, CLng((Height - shProbe * zf)) / 2, Width, CLng(shProbe * zf), , rtxy
                    Else
                        StretchPicture p_hdc, CLng(fr) + CLng((sw - sh * swProbe / shProbe) / 2) + offsetX + CLng(Border / osh), CLng(fr) + offsetY + Border, CLng(sh * swProbe / shProbe) - 2 * Border, sh - 2 * Border, 0, CLng((Height - shProbe * zf)) / 2, Width, CLng(shProbe * zf), , rtxy
                    End If
                Else
                    zf = Height / shProbe
                    If swProbe > shProbe Then
                     StretchPicture p_hdc, CLng(fr) + offsetX + CLng(Border * osh), CLng((sh - sw * shProbe / swProbe) / 2 + fr) + offsetY + Border, sw - Border * 2, CLng(sw * shProbe / swProbe) - Border * 2, (Width - swProbe * zf) / 2, 0, CLng(swProbe * zf), Height, , rtxy
                    Else
                     StretchPicture p_hdc, CLng(fr) + CLng((sw - sh * swProbe / shProbe) / 2) + offsetX + Border, CLng(fr) + offsetY + CLng(Border * osh), CLng(sh * swProbe / shProbe) - 2 * Border, sh - 2 * Border, (Width - swProbe * zf) / 2, 0, CLng(swProbe * zf), Height, , rtxy
                 End If
                  End If
             Else
                If (swProbe / Width) < (shProbe / Height) Then
                    zf = sw / Width
                Else
                    zf = sh / Height
                End If
                StretchPicture p_hdc, (sw - Width * zf) / 2 + offsetX, (sh - Height * zf) / 2 + offsetY, CLng(Width * zf), CLng(Height * zf), , , , , , rtxy
            End If
        End If
    Else
        zf = zf / 100#
        If Not FitToPic Then
            StretchPicture p_hdc, Abs((sw - Width * zf) / 2), Abs((sh - Height * zf) / 2), Width * zf, Height * zf, , , , , , rtxy
        Else
            If sw / Width < sh / Height Then
                zf = zf * sw / Width
            Else
                zf = zf * sh / Height
            End If
                StretchPicture p_hdc, (sw - Width * zf) / 2, (sh - Height * zf) / 2, Width * zf, Height * zf, , , , , , rtxy
        End If
    End If
    FreeHDC
    If p_hdc <> 0 Then
    EndPage p_hdc
    SelectObject p_hdc, m_PrinterhBmpOld
    End If
    m_PrinterhBmpOld = 0
    CopiesCount (1), MyDM()
    last_zf = zf
    m_logx = 1: m_logy = 1
    
End Sub

Public Sub ThumbnailPaintDummy(DestPic As Object, Optional ByVal zf As Single = -1, Optional ByVal FitToPic As Boolean = True, Optional CROP As Boolean = False, Optional over As Boolean = False, Optional ByVal fr As Long = 0, Optional ByVal offsetX As Long = 0, Optional ByVal offsetY As Long, Optional ByVal Border As Long = 0, Optional ByVal rtxy As Double = 1#)
   Dim oldSc As Long, Dummy1 As Boolean, Dummy2 As Boolean, sw As Long, sh As Long
   Dim swProbe As Long, shProbe As Long, osh As Double
   Dim Myleft As Long, mytop As Long, mywidth As Long, myheight As Long
   
     Dim A As Single
         oldSc = DestPic.ScaleMode
   DestPic.ScaleMode = vbPixels

   
 
  If over Then
   swProbe = DestPic.ScaleX(DestPic.Width, 1, 3)
   shProbe = DestPic.ScaleY(DestPic.Height, 1, 3)
   Else
   swProbe = DestPic.ScaleWidth
   shProbe = DestPic.ScaleHeight
   End If
    sw = DestPic.ScaleWidth
   sh = DestPic.ScaleHeight
 If sh * sw = 0 Then DestPic.ScaleMode = oldSc: Exit Sub
   
   Dim osw As Long
   
If zf <= 0 Then

      
        If Not FitToPic Then
        '    StretchPicture DestPic.hdc, (sw - width) / 2 + offsetx, (sh - height) / 2 + offsety, , , , , , , , rtxy
        Else
      If sw < sh Then
    osw = offsetX: offsetX = -offsetY:
    offsetY = osw
    
     rtxy = 1 / rtxy
   End If
            If CROP Then
                   
                    sh = sh - 2 * fr
                    sw = sw - 2 * fr
                    If shProbe < swProbe Then
                    osh = shProbe / sh
                        shProbe = CLng(sw * shProbe / swProbe) - Border * 2 / osh
                        swProbe = sw - Border * 2 * osh
                   Else
                   osh = swProbe / sw
            
                        swProbe = CLng(sh * swProbe / shProbe) - 2 * Border / osh
                        shProbe = sh - 2 * Border * osh
                   End If
                   ' Τελευταία αλλαγή - έφυγαν τα CLNG(border/osh) από τα x,y για να τυπώνει σωστά τα offset
                 If (swProbe / Width) > (shProbe / Height) Then
                    zf = Width / swProbe
                    If swProbe > shProbe Then
                      ' StretchPicture DestPic.hdc, CLNG(fr) + offsetx + CLNG(border * osh), CLNG((sh - sw * shProbe / swProbe) / 2 + fr) + offsety + border, sw - border * 2, CLNG(sw * shProbe / swProbe) - border * 2, 0, CLNG((height - shProbe * zf)) / 2, width, CLNG(shProbe * zf), , rtxy
                    Else
                      '  StretchPicture DestPic.hdc, CLNG(fr) + CLNG((sw - sh * swProbe / shProbe) / 2) + offsetx + CLNG(border / osh), CLNG(fr) + offsety + border, CLNG(sh * swProbe / shProbe) - 2 * border, sh - 2 * border, 0, CLNG((height - shProbe * zf)) / 2, width, CLNG(shProbe * zf), , rtxy
                    End If
                Else
                    zf = Height / shProbe
                    If swProbe > shProbe Then
                '     StretchPicture DestPic.hdc, CLNG(fr) + offsetx + CLNG(border * osh), CLNG((sh - sw * shProbe / swProbe) / 2 + fr) + offsety + border, sw - border * 2, CLNG(sw * shProbe / swProbe) - border * 2, (width - swProbe * zf) / 2, 0, CLNG(swProbe * zf), height, , rtxy
                    Else
                 '    StretchPicture DestPic.hdc, CLNG(fr) + CLNG((sw - sh * swProbe / shProbe) / 2) + offsetx + border, CLNG(fr) + offsety + CLNG(border * osh), CLNG(sh * swProbe / shProbe) - 2 * border, sh - 2 * border, (width - swProbe * zf) / 2, 0, CLNG(swProbe * zf), height, , rtxy
                 End If
                  End If
             Else
                If (swProbe / Width) < (shProbe / Height) Then
                    zf = sw / Width
                Else
                    zf = sh / Height
                End If
             '   StretchPicture DestPic.hdc, (sw - width * zf) / 2 + offsetx, (sh - height * zf) / 2 + offsety, CLNG(width * zf), CLNG(height * zf), , , , , , rtxy
            End If
        End If
    Else
        zf = zf / 100#
        If Not FitToPic Then

           ' StretchPicture DestPic.hdc, offsetx + Abs((DestPic.ScaleWidth - width * zf) / 2), offsety + Abs((DestPic.ScaleHeight - height * zf) / 2), width * zf, height * zf, , , , , , rtxy
                    Else
            If DestPic.ScaleWidth / Width < DestPic.ScaleHeight / Height Then
                zf = zf * DestPic.ScaleWidth / Width
                
            Else
                zf = zf * DestPic.ScaleHeight / Height
            End If
            On Error Resume Next

         '       StretchPicture DestPic.hdc, offsetx + (DestPic.ScaleWidth - width * zf) / 2, offsety + (DestPic.ScaleHeight - height * zf) / 2, width * zf, height * zf, , , , , , rtxy
        End If
    End If
    
   ' UpdateWindow DestPic.hwnd 'DestPic.Refresh
  
    DestPic.ScaleMode = oldSc
    last_zf = zf
    ' gdiflush
End Sub
Public Sub ThumbnailPaintDummyOLD(DestPic As Object, Optional ByVal zf As Single = -1, Optional ByVal FitToPic As Boolean = True, Optional CROP As Boolean = False, Optional over As Boolean = False, Optional ByVal fr As Long = 0, Optional ByVal offsetX As Long = 0, Optional ByVal offsetY As Long, Optional ByVal Border As Long = 0)
   Dim oldSc As Long, Dummy1 As Boolean, Dummy2 As Boolean, sw As Long, sh As Long
   Dim swProbe As Long, shProbe As Long, osh As Double
   Dim Myleft As Long, mytop As Long, mywidth As Long, myheight As Long
     Dim A As Single
      
   oldSc = DestPic.ScaleMode
   DestPic.ScaleMode = vbPixels
     If over Then
   swProbe = DestPic.ScaleX(DestPic.Width, 1, 3)
   shProbe = DestPic.ScaleY(DestPic.Height, 1, 3)
   Else
   swProbe = DestPic.ScaleWidth
   shProbe = DestPic.ScaleHeight
   End If
    sw = DestPic.ScaleWidth
   sh = DestPic.ScaleHeight
   
   Dim osw As Long
   If sw > sh Then osw = offsetX: offsetX = offsetY: offsetY = osw
 
   
    If zf <= 0 Then
        If Not FitToPic Then
        Else
            If CROP Then
                   
                    sh = sh - 2 * fr
                    sw = sw - 2 * fr
                    If shProbe < swProbe Then
                    osh = shProbe / sh
                        shProbe = CLng(sw * shProbe / swProbe) - Border * 2 / osh
                        swProbe = sw - Border * 2 * osh
                   Else
                   osh = swProbe / sw
                        swProbe = CLng(sh * swProbe / shProbe) - 2 * Border / osh
                        shProbe = sh - 2 * Border * osh
                   End If
                 If (swProbe / Width) > (shProbe / Height) Then
                    zf = Width / swProbe
                Else
                    zf = Height / shProbe

                  End If
             Else
                If (swProbe / Width) < (shProbe / Height) Then
                    zf = sw / Width
                Else
                    zf = sh / Height
                End If
            End If
        End If
    Else
        zf = zf / 100#
        If Not FitToPic Then
        Else
            If DestPic.ScaleWidth / Width < DestPic.ScaleHeight / Height Then
                zf = zf * DestPic.ScaleWidth / Width
            Else
                zf = zf * DestPic.ScaleHeight / Height
            End If
        End If
    End If
    DestPic.ScaleMode = oldSc
    last_zf = zf
    ' gdiflush
End Sub
Public Sub ThumbnailPartPaint(DestPic As Object, _
                    Optional ByVal ThumbnailSizeX As Long = 32, _
                    Optional ByVal ThumbnailSizeY As Long = 32, _
                    Optional ByVal xPos As Long, _
                    Optional ByVal yPos As Long, _
                    Optional Border As Boolean = False, _
                    Optional ThumbnailDist As Long = 16, _
                    Optional ByVal lText As String = vbNullString, _
                    Optional offsetX As Long = 0, _
                    Optional offsetY As Long = 0)
                    Dim ThumbnailDistY As Long, zf As Single, ofsize As Single
   Dim oldSc As Long, ms As Long

     Dim dr1 As RECT
   oldSc = DestPic.ScaleMode

   DestPic.ScaleMode = vbPixels
    If lText <> "" Then
    ThumbnailDistY = DestPic.TextHeight("Ay") + 2
   ms = (ThumbnailSizeX + 8) / DestPic.TextWidth("a") - 1
   If ms > 2 Then
   If Len(lText) > ms Then lText = Left$(lText, ms - 2) & ".."
    Else
        lText = vbNullString
     ThumbnailDistY = 0
End If


    End If
            If (ThumbnailSizeX + 2 * ThumbnailDist * Border) / Width < (ThumbnailSizeY + 2 * ThumbnailDist * Border - ThumbnailDistY) / Height Then
                zf = (ThumbnailSizeX + 2 * ThumbnailDist * Border - 2) / Width
            Else
                zf = (ThumbnailSizeY + 2 * ThumbnailDist * Border - ThumbnailDistY - 2) / Height
            End If
            If Border Then
            DestPic.Line ((ThumbnailSizeX) * xPos + offsetX, (ThumbnailSizeY) * yPos + offsetY)-((ThumbnailSizeX) * (xPos + 1) + offsetX, (ThumbnailSizeY) * (yPos + 1) + offsetY), 0, B
            End If
Dim myx As Long, myy As Long, myw As Long, myh As Long
myx = ThumbnailSizeX '+ ThumbnailDist
myy = ThumbnailSizeY - ThumbnailDistY '+ ThumbnailDist
myw = Width * zf
myh = Height * zf
myx = (myx - myw) / 2 + offsetX '
myy = (myy - myh) / 2 + offsetY  '
On Error Resume Next
StretchPicture DestPic.hDC, (ThumbnailSizeX) * xPos + myx, (ThumbnailSizeY) * yPos + myy, myw, myh
If Err.Number > 0 Then
Else
     If ThumbnailDistY <> 0 Then
             DestPic.currentX = (ThumbnailSizeX) * xPos + ((ThumbnailSizeX) - DestPic.TextWidth(lText)) / 2 + offsetX
            DestPic.currentY = (ThumbnailSizeY) * (yPos + 1) + offsetY - ThumbnailDistY
            dr1.Left = DestPic.currentX
            dr1.Top = DestPic.currentY
            dr1.Right = DestPic.ScaleWidth
            dr1.Bottom = -DestPic.ScaleHeight
            DrawText DestPic.hDC, StrPtr(lText), Len(lText), dr1, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP
        
    End If
    End If
     DestPic.ScaleMode = oldSc
     last_zf = zf
     ' gdiflush
End Sub


Private Sub Class_Initialize()
PaintType = False
 SD = "DISPLAY" & Chr(0)
 emfSizeFactor = 4000
 m_fontname = "Verdana"
 m_fontsize = 8
 m_logx = 1
 m_logy = 1
 m_transparent = 1
End Sub

Private Sub Class_Terminate()
 ClearUp

End Sub

Public Sub SaveDib(f As String)
Dim bDib() As Byte
Dim tSA As SAFEARRAY2D
Dim A$, mf As Integer, ff As BITMAPFILEHEADER

On Error Resume Next
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

    mf = FreeFile

    With ff
        .bfType = &H4D42
        .bfSize = Len(ff) + Me.packet_bytes
        .bfOffBits = Len(ff) + Me.header_bytes
    End With
Dim maxcolors As Long
 maxcolors = m_tBI.bmiHeader.biBitCount
If maxcolors < 16 Then
If m_tBI.bmiHeader.biClrUsed = 0 Then
maxcolors = 2 * maxcolors
Else
maxcolors = m_tBI.bmiHeader.biClrUsed
End If
Else
maxcolors = 0
End If
Dim i As Long
If NeoUnicodeFile(f) Then
Open GetDosPath(f) For Binary As mf
Put #mf, , ff
Put #mf, , m_tBI.bmiHeader
If maxcolors > 0 Then
For i = 0 To maxcolors - 1
Put #mf, , m_tBI.bmiColors(i)
Next i
End If
Put #mf, , bDib()
Close #mf
End If
CopyMemory ByVal VarPtrArray(bDib), 0&, 4
End Sub

Public Function SaveDibToMeMBlock(mem As Object) As Boolean  ' 24 bits only
If m_tBI.bmiHeader.biBitCount <> 24 Then Exit Function
Dim Membl As New MemBlock
Dim ff As BITMAPFILEHEADER
On Error GoTo ex1
With ff
    .bfType = &H4D42
    .bfSize = Len(ff) + Me.packet_bytes
    .bfOffBits = Len(ff) + Me.header_bytes
End With
Membl.Construct 1, Len(ff) + Me.header_bytes + Me.BytesPerScanLine * Me.Height, 0 '-1
CopyMemory ByVal Membl.GetPtr(0), ff, 2
CopyMemory ByVal Membl.GetPtr(0) + 2, ff.bfSize, 12  ' vb align issue
CopyMemory ByVal Membl.GetPtr(0) + Len(ff), m_tBI, Len(m_tBI)
CopyMemory ByVal Membl.GetPtr(0) + Len(ff) + Me.header_bytes, ByVal m_lPtr, Me.BytesPerScanLine * Me.Height
Set mem = Membl
Membl.SubType = 1
SaveDibToMeMBlock = True
ex1:
End Function

Public Sub loadDIB(f As String)
Dim fl As Long, fmem() As Byte, FH As Long, ff As BITMAPFILEHEADER
If GetDosPath(f) = vbNullString Then MissFile: Exit Sub
On Error GoTo er11
fl = FileLen(GetDosPath(f))
If fl > 0 Then
ReDim fmem(fl) As Byte
FH = FreeFile

Open GetDosPath(f) For Binary As FH
Get #FH, , fmem()
Close FH

CreateFromDIB VarPtr(fmem(Len(ff)))
End If
Exit Sub
er11:
Err.Clear
On Error Resume Next
If FH <> 0 Then Close FH
MyEr "Can't read image", "Δεν μπορώ να διαβάσω το αρχείο"
End Sub

Private Sub SentToClipBoard()
Dim hMem As Long, Mpoke As Long
If Me.Height > 0 Then
    hMem = GlobalAlloc(0, Me.packet_bytes)
    If hMem > 0 Then
            Mpoke = GlobalLock(hMem)
        CopyMemory ByVal Mpoke, m_tBI, Me.header_bytes
        CopyMemory ByVal Mpoke + Me.header_bytes, ByVal m_lPtr, Me.BytesPerScanLine * Me.Height
        If (OpenClipboard(0) <> 0) Then
            EmptyClipboard
            
            If (SetClipboardData(CF_DIB, hMem) <> 0) Then
               ' CopyToClipboard = True
            End If
            CloseClipboard
        End If
    Call GlobalUnlock(hMem)
    End If
End If

End Sub
Public Sub Text(p$, X As Long, Y As Long, Optional s As Single = -1, Optional DE# = 0#, Optional Style As Long = 272, Optional mWidth As Long = -1, Optional expnd As Single = -1)
Dim old_brush As Long, old_pen As Long, my_brush As Long, my_pen As Long, ra As Long
Dim f As LOGFONT, hPrevFont As Long, hFont As Long
Dim fff As Long, dr1 As RECT
needHDC
If s < 0 Then s = m_fontsize
Dim ratioXY As Double
'On Error Resume Next
Dim icx As Long, icy As Long
s = s * dpiy / 96
ratioXY = 1
Dim degree As Double
degree = DE# * 1.74532925199433E-02
icx = (s * -20) / 15
  icy = CLng(Cos(degree + Pi) * icx)
    icx = CLng(Sin(degree + Pi) * icx)
'If degree <> 0 Then
degree = (degree) * 180# / Pi
  f.lfEscapement = CLng(10 * degree)
  f.lfFaceName = m_fontname + Chr$(0)
  f.lfCharSet = 161 '.hpSet
  f.lfQuality = 2 'PROOF_QUALITY
  f.lfHeight = (s * -20) / 15
If expnd > 0 Then
f.lfWidth = Abs(f.lfHeight) / ratioXY * expnd
Else
'''' ΑΛΛΑΓΗ ΓΙΑ ΒΑΘΜΟΛΟΓΙΟ
   '''''' f.lfWidth = Abs(f.lfHeight) / ratioXY / 2
    End If
  f.lfItalic = 0
  f.lfUnderline = 0
 ' .If DestPic.Font.Bold Or expnd Then
  'f.lfWeight = 900
  'Else
  f.lfWeight = 400

  'End If
  hFont = CreateFontIndirect(f)
  hPrevFont = SelectObject(m_hDC, hFont)

With dr1
.Top = Y
.Left = X
If mWidth = -1 Then
.Right = X + Width
Else
.Right = mWidth

End If
.Bottom = Y + Height
End With
fff = Style '(256 + 16) Or DT_NOPREFIX

'my_pen = CreatePen(PS_NULL, 1, ra)
'old_pen = SelectObject(m_hDC, my_pen)
SetBkMode m_hDC, m_fonttransparent
'my_brush = CreateSolidBrush(m_backcolor)
'old_brush = SelectObject(m_hDC, my_brush)
SetTextColor m_hDC, m_BackColor

DrawText m_hDC, StrPtr(p$), Len(p$), dr1, fff
'old_brush = SelectObject(m_hDC, old_brush)
'DeleteObject my_brush
'old_pen = SelectObject(m_hDC, old_pen)
'DeleteObject my_pen
hFont = SelectObject(m_hDC, hPrevFont)
  DeleteObject hFont

FreeHDC
End Sub
Public Sub TextW(p$, X As Long, Y As Long, x1 As Long, y1 As Long, Optional s As Single = -1, Optional DE# = 0#, Optional Style As Long = 272)
Dim old_brush As Long, old_pen As Long, my_brush As Long, my_pen As Long, ra As Long
Dim f As LOGFONT, hPrevFont As Long, hFont As Long
Dim fff As Long, dr1 As RECT, expnd As Long
needHDC
If s < 0 Then s = m_fontsize
Dim ratioXY As Double
'On Error Resume Next
Dim icx As Long, icy As Long
s = s * dpiy / 96
ratioXY = 1
Dim degree As Double
degree = DE# * 1.74532925199433E-02
icx = (s * -20) / 15
  icy = CLng(Cos(degree + Pi) * icx)
    icx = CLng(Sin(degree + Pi) * icx)
'If degree <> 0 Then
degree = (degree) * 180# / Pi
  f.lfEscapement = CLng(10 * degree)
  f.lfFaceName = m_fontname + Chr$(0)
  f.lfCharSet = 161 '.hpSet
  f.lfQuality = 2 'PROOF_QUALITY
  f.lfHeight = (s * -20) / 15
If expnd > 0 Then
f.lfWidth = Abs(f.lfHeight) / ratioXY * expnd
Else
'''' ΑΛΛΑΓΗ ΓΙΑ ΒΑΘΜΟΛΟΓΙΟ
   '''''' f.lfWidth = Abs(f.lfHeight) / ratioXY / 2
    End If
  f.lfItalic = 0
  f.lfUnderline = 0
 ' .If DestPic.Font.Bold Or expnd Then
  'f.lfWeight = 900
  'Else
  f.lfWeight = 400

  'End If
  hFont = CreateFontIndirect(f)
  hPrevFont = SelectObject(m_hDC, hFont)

With dr1
.Top = Y
.Left = X
.Right = x1
.Bottom = y1
End With
fff = Style '(256 + 16) Or DT_NOPREFIX
SetBkMode m_hDC, m_fonttransparent
'my_brush = CreateSolidBrush(m_backcolor)
'old_brush = SelectObject(m_hDC, my_brush)
SetTextColor m_hDC, m_BackColor
DrawText m_hDC, p$, Len(p$), dr1, fff
'old_brush = SelectObject(m_hDC, old_brush)
'DeleteObject my_brush
'old_pen = SelectObject(m_hDC, old_pen)
'DeleteObject my_pen
hFont = SelectObject(m_hDC, hPrevFont)
  DeleteObject hFont

FreeHDC


End Sub

Public Sub kiklos(Radius As Long, X As Long, Y As Long)  ' circle
'
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
x1 = X - Radius
y1 = Y - Radius
x2 = X + Radius
y2 = Y + Radius

Dim old_brush As Long, old_pen As Long, my_brush As Long, my_pen As Long, ra As Long

' set brush
 needHDC
my_pen = CreatePen(PS_NULL, 1, ra)
old_pen = SelectObject(m_hDC, my_pen)
my_brush = CreateSolidBrush(m_BackColor)
old_brush = SelectObject(m_hDC, my_brush)
Ellipse m_hDC, x1, y1, x2, y2
old_brush = SelectObject(m_hDC, old_brush)
DeleteObject my_brush
old_pen = SelectObject(m_hDC, old_pen)
DeleteObject my_pen
FreeHDC
' CreateSolidBrush color as rgbquad
' but not pen
' CreatePen with  pen style  as PS_NULL
' perform ellipse

End Sub
Public Sub Grami(x1 As Long, y1 As Long, x2 As Long, y2 As Long, w1 As Long, mycolor As Variant)
Dim old_brush As Long, my_brush As Long, m_color As Long
Dim old_pen As Long, my_pen As Long, p As POINTAPI
' line
m_color = getRGB(mycolor)
needHDC

MoveToEx m_hDC, x1, y1, p
my_pen = CreatePen(PS_SOLID, w1, m_color)
old_pen = SelectObject(m_hDC, my_pen)
''''''''''''''''''''''''''''''''''''''''''
old_pen = SelectObject(m_hDC, old_pen)
DeleteObject my_pen
FreeHDC
End Sub
Public Sub square(x1 As Long, y1 As Long, x2 As Long, y2 As Long, mycolor As Variant)
Dim old_brush As Long, my_brush As Long, m_color As Long
m_color = getRGB(mycolor)
Dim mr As RECT
' set brush
my_brush = CreateSolidBrush(m_color)
mr.Left = x1
mr.Right = x2 + 1 ' to +1 to αφαιρεί!
mr.Top = y1
mr.Bottom = y2 + 1
needHDC
FillRect m_hDC, mr, my_brush

FreeHDC
DeleteObject my_brush
End Sub


Public Property Get HDC1() As Long
 ' ************here********************
    HDC1 = m_hDC
End Property

Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long, pic As PicBmp, ipic As IPicture, IID_IDispatch As GUID

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
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    'Create the picture
    R = OleCreatePictureIndirect(pic, IID_IDispatch, 1, ipic)

    'Return the new picture
    Set CreateBitmapPicture = ipic
End Function
Function Picture2(Optional myscale As Single = -1) As Picture
Dim hdcSrc As Long, LeftSrc As Long, TopSrc As Long, widthSrc As Long, heightSrc As Long
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, R As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE
    If myscale = -1 And Not wmfemf Is Nothing Then
    Set Picture2 = wmfemf
    Exit Function
    End If
    If myscale < 0 Then myscale = 1
    needHDC
hdcSrc = m_hDC
LeftSrc = 0
TopSrc = 0
widthSrc = Width * myscale
heightSrc = Height * myscale
    'Create a compatible device context
    hDCMemory = CreateCompatibleDC(hdcSrc)
    'Create a compatible bitmap
    hBmp = CreateCompatibleBitmap(hdcSrc, widthSrc, heightSrc)
    'Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(hdcSrc, RASTERCAPS) ' Raster
    'Does our Picture2 use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hdcSrc, SIZEPALETTE) ' Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        LogPal.palVersion = &H300
        'Number of palette entries
        LogPal.palNumEntries = 256
        'Retrieve the system palette entries
        R = GetSystemPaletteEntries(hdcSrc, 0, 256, LogPal.palPalEntry(0))
        'Create the palette
        hPal = CreatePalette(LogPal)
        'Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        'Realize the palette
        R = RealizePalette(hDCMemory)
    End If

    'Copy the source image to our compatible device context
    If myscale = 1 Then
    R = BitBlt(hDCMemory, 0, 0, Width, Height, hdcSrc, 0, 0, vbSrcCopy)
Else
  Call SetStretchBltMode(hDCMemory, COLORONCOLOR)
    R = StretchDIBits(hDCMemory, 0, 0, widthSrc, heightSrc, 0, 0, Width, Height, ByVal m_lPtr, m_tBI, DIB_RGB_COLORS, vbSrcCopy)
End If
    'Restore the old bitmap
    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Select the palette
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    'Delete our memory DC
    R = DeleteDC(hDCMemory)
FreeHDC
    Set Picture2 = CreateBitmapPicture(hBmp, hPal)
End Function
Function Picture(Optional myscale As Single = -1) As Picture
Dim hdcSrc As Long, LeftSrc As Long, TopSrc As Long, widthSrc As Long, heightSrc As Long
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, R As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE
    If myscale = -1 And Not wmfemf Is Nothing Then
    Set Picture = wmfemf
    Exit Function
    End If
    If myscale < 0 Then myscale = 1
    needHDC
hdcSrc = m_hDC
LeftSrc = 0
TopSrc = 0
widthSrc = Width * myscale
heightSrc = Height * myscale
    'Create a compatible device context
    hDCMemory = CreateCompatibleDC(hdcSrc)
    'Create a compatible bitmap
    hBmp = CreateCompatibleBitmap(hdcSrc, widthSrc, heightSrc)
    'Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(hdcSrc, RASTERCAPS) ' Raster
    'Does our picture use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hdcSrc, SIZEPALETTE) ' Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        LogPal.palVersion = &H300
        'Number of palette entries
        LogPal.palNumEntries = 256
        'Retrieve the system palette entries
        R = GetSystemPaletteEntries(hdcSrc, 0, 256, LogPal.palPalEntry(0))
        'Create the palette
        hPal = CreatePalette(LogPal)
        'Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        'Realize the palette
        R = RealizePalette(hDCMemory)
    End If

    'Copy the source image to our compatible device context
    If myscale = 1 Then
    R = BitBlt(hDCMemory, 0, 0, Width, Height, hdcSrc, 0, 0, vbSrcCopy)
Else
  Call SetStretchBltMode(hDCMemory, HALFTONE)
    R = StretchDIBits(hDCMemory, 0, 0, widthSrc, heightSrc, 0, 0, Width, Height, ByVal m_lPtr, m_tBI, DIB_RGB_COLORS, vbSrcCopy)
End If
    'Restore the old bitmap
    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Select the palette
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    'Delete our memory DC
    R = DeleteDC(hDCMemory)
FreeHDC
    Set Picture = CreateBitmapPicture(hBmp, hPal)
End Function

Public Function CreateFromPictureHandle( _
        ByRef picthandle As Long _
    ) As Boolean
Dim lHDC As Long
Dim lhDCDesktop As Long
Dim lhBmpOld As Long
Dim tBMP As Bitmap
    ClearUp
    GetObjectAPI picthandle, Len(tBMP), tBMP
    If (create(tBMP.bmWidth, tBMP.bmHeight)) Then
   lhDCDesktop = CreateDC(SD, 0, 0, ByVal 0)
  lHDC = CreateCompatibleDC(lhDCDesktop)
            DeleteDC lhDCDesktop
            If (lHDC <> 0) Then
                lhBmpOld = SelectObject(lHDC, picthandle)
                LoadPictureBlt lHDC
                SelectObject lHDC, lhBmpOld
                CreateFromPictureHandle = True
            End If
     DeleteDC lHDC
        End If
    
End Function
Public Function CreateFromPicture(ByRef picthis As StdPicture, Optional usecolor As Long = &HFFFFFF, Optional actualwidth As Long, Optional actualheight As Long)
Dim lHDC As Long
Dim lhDCDesktop As Long
Dim lhBmpOld As Long
Dim tBMP As Bitmap
Dim kkx As RECT, autosize As Single
Dim newHandle As Long
Dim newPic As Object
ClearUp
If picthis Is Nothing Then
mType = -1
Set wmfemf = Nothing
Exit Function
End If

mType = picthis.Type

If picthis.Type = vbPicTypeIcon Then
Set wmfemf = picthis
'
Const emff = 24.9428571428571


Dim my_brush As Long
create picthis.Width * 0.567 / DXP, picthis.Height * 0.567 / DYP
Me.ClearBits m_BackColor
needHDC

my_brush = CreateSolidBrush(m_BackColor)
'DrawIconEx HDC1, 0, 0, picthis, 48, 48, 0, my_brush, &H3 ' DI_MASK

If IsWine Then
' THIS DRAW  INVERt PIXELS AS TRANSPARENT
DrawIconEx HDC1, 0, 0, picthis, 48, 48, 0, my_brush, DI_IMAGE
DrawState HDC1, 0, 0, picthis.Handle, 48, 48, 0, 0, 0, DI_IMAGE & DI_MASK
Else
' THIS CONVERT INVERT PIXELS TO BLACK
DrawState HDC1, 0, 0, picthis.Handle, 0, 0, 0, 0, 0, &H83
DrawState HDC1, my_brush, 0, picthis.Handle, 0, 0, 0, 0, 0, &H3
DrawState HDC1, my_brush, 0, picthis.Handle, 0, 0, 0, 0, 0, &H3
End If

DeleteObject my_brush
FreeHDC
'End If
ElseIf picthis.Type = vbPicTypeEMetafile Then
Set wmfemf = picthis

If Not (actualwidth = 0) Then
    kkx.Right = actualwidth
Else
    kkx.Right = Int(picthis.Width / 2540& * (1440 / dv15) * emfSizeFactor)
End If
' Int(picthis.width / 2540 * 1440 / dv15) '   Round(emfSizeFactor * picthis.width / 1.75797922568461 / dv15) + 1
If Not (actualheight = 0) Then
kkx.Bottom = actualheight
Else
kkx.Bottom = Int(picthis.Height / 2540& * (1440 / dv15) * emfSizeFactor)
End If
' Int(picthis.height / 2540 * 1440 / dv15) ' Round(emfSizeFactor * picthis.height / 1.75797922568461 / dv15) + 1


 If create(kkx.Right, kkx.Bottom) Then
 
 ClearBits usecolor

  needHDC
  kkx.Right = kkx.Right * 1.00166666666667
  kkx.Bottom = kkx.Bottom * 1.00333333333333
 PlayEnhMetaFile HDC1, picthis.Handle, kkx
 FreeHDC
 Else
 GoTo canmake
 End If
 ElseIf picthis.Type = vbPicTypeMetafile Then
Set wmfemf = picthis
kkx.Right = Form1.ScaleX(picthis.Width, vbHimetric, vbPixels) * emfSizeFactor - 1 ' Int(picthis.Width / 2540& * (1440 / dv15) * emfSizeFactor)
kkx.Bottom = Form1.ScaleY(picthis.Height, vbHimetric, vbPixels) * emfSizeFactor - 1 ' Int(picthis.Height / 2540& * (1440 / dv15) * emfSizeFactor)


 If create(kkx.Right, kkx.Bottom) Then
 ClearBits usecolor
  newHandle = ConvertWmftoEmF(picthis.Handle)
  needHDC
 PlayEnhMetaFile HDC1, newHandle, kkx
 FreeHDC
 DeleteEnhMetaFile newHandle
 Else
 GoTo canmake
 End If
 

Else

    GetObjectAPI picthis.Handle, Len(tBMP), tBMP
    If (create(tBMP.bmWidth, tBMP.bmHeight)) Then
    WhiteBits
   lhDCDesktop = CreateDC(SD, 0, 0, ByVal 0)
  lHDC = CreateCompatibleDC(lhDCDesktop)
            DeleteDC lhDCDesktop
            If (lHDC <> 0) Then
                lhBmpOld = SelectObject(lHDC, picthis.Handle)
                LoadPictureBlt lHDC
                SelectObject lHDC, lhBmpOld
            End If
     DeleteDC lHDC
          Else
        GoTo canmake
        End If
   
    End If
    Exit Function
canmake:
     create 2000, 500
 WhiteBits
 Me.FontSize = 128
 If autosize = 0 Then Exit Function
 Text "big....size " & (3 * Int(picthis.Width / autosize) * Int(picthis.Height / autosize) / 1024# / 1024#) & " Mbyte", 0, 0

End Function

Sub PageView(where As Object, Optional How As Long = -1, Optional vso As Object = Nothing, Optional hso As Object = Nothing, Optional Percent As Single = 1, Optional gid As Long = 0)
' object need gonia 375x375 frame if we use vso and hso.
' from vs and hs remove tabstop
' use scroll with pageview
' use timer fron vs.value and hs.value
' timer with interval 40 miliseconds

Dim redX As Long, redY As Long, ss As Single, aa As Single

Dim VredX As Long, VredY As Long
Static once As Boolean

If once Then Exit Sub
If where.ScaleHeight < DXP Or where.ScaleWidth < DYP Then Exit Sub
If How = -1 Then How = m_PageStyle
m_PageStyle = How
once = True
'On Error GoTo there1
If Not vso Is Nothing Then
 redX = vso.Width
End If
If Not hso Is Nothing Then
 redY = hso.Height
End If
If How = 0 Then
If Not hso Is Nothing Then
If hso.Visible Then hso.Visible = False
End If
If Not vso Is Nothing Then
If vso.Visible Then vso.Visible = False
End If
where.Cls
ThumbnailPaint where
ElseIf How = 1 Then

ThumbnailPaintDummy where
ss = LastZoom
aa = ss
where.Cls
ss = 100 / ss / (dpiy / 96) * Percent

ThumbnailPaintDummy where, ss ', , , , , , -100
aa = LastZoom
If (where.ScaleHeight / DYP) / (Height * aa) > 1 Then
If Not vso Is Nothing Then
vso.Visible = False: redY = 0
If Not hso Is Nothing Then
If (where.ScaleHeight - hso.Height) > 0 Then hso.move 0, where.ScaleHeight - hso.Height, where.ScaleWidth
End If
End If
Else
If Not vso Is Nothing Then
If vso.Visible = False Then
vso.Visible = True
vso.enabled = False: vso.Value = 0: vso.enabled = True
End If
vso.Visible = True
VredX = redX
End If

End If

If (where.ScaleWidth / DXP) / (Width * aa) > 1 Then
If Not hso Is Nothing Then
hso.Visible = False: redX = 0
If Not vso Is Nothing Then
If (where.ScaleWidth - vso.Width) > 0 Then vso.move where.ScaleWidth - vso.Width, 0, vso.Width, where.ScaleHeight
End If
End If
Else
If Not hso Is Nothing Then
If hso.Visible = False Then
hso.Visible = True
hso.enabled = False: hso.Value = 0: hso.enabled = True
End If
hso.Visible = True
VredY = redY
End If

End If
If Not vso Is Nothing Then
If vso.Visible Then
vso.Max = Height * aa - (where.ScaleHeight - VredY) \ DYP
vso.smallchange = 1
If (where.ScaleHeight - VredY) \ DYP > 0 Then vso.largechange = (where.ScaleHeight - VredY) \ DYP
vso.Refresh
redY = -(vso.Value)

End If
End If
If Not hso Is Nothing Then
If hso.Visible Then
hso.Max = Width * aa - (where.ScaleWidth - VredX) \ DXP
hso.smallchange = 1
If (where.ScaleWidth - VredX) \ DXP > 0 Then hso.largechange = (where.ScaleWidth - VredX) \ DXP
hso.Refresh
redX = -(hso.Value)
End If
End If
If Not vso Is Nothing Then If (where.ScaleWidth - vso.Width) > 0 And (where.ScaleHeight - VredY) > 0 Then vso.move where.ScaleWidth - vso.Width, 0, vso.Width, where.ScaleHeight - VredY
If Not hso Is Nothing Then If (where.ScaleHeight - hso.Height) > 0 And (where.ScaleWidth - VredX) > 0 Then hso.move 0, where.ScaleHeight - hso.Height, where.ScaleWidth - VredX

'*****************
ThumbnailPaint where, ss, , , , , -(((where.ScaleWidth) / DXP) - Width * aa) / 2 + redX, -(((where.ScaleHeight) / DYP) - Height * aa) / 2 + redY
' center
'ThumbnailPaint where, SS, , , , , , -(((where.ScaleHeight - redY) / 15) - Int(height * aa)) / 2
ElseIf How = 3 Then
where.Cls
ThumbnailPaintDummy where
ss = LastZoom
If ((where.ScaleWidth - redX) / DXP) / (Width * ss) < 1.0001 Then
If Not vso Is Nothing Then vso.Visible = False: redX = 0
Else
If Not vso Is Nothing Then
If vso.Visible = False Then
vso.Visible = True: vso.enabled = False: vso.Value = 0: vso.enabled = True
End If
vso.Visible = True: redX = vso.Width
End If
End If
ss = ((where.ScaleWidth - redX) / DXP) / (Width * ss)
ThumbnailPaintDummy where, ss * 100

aa = LastZoom
If Not hso Is Nothing Then
If hso.Visible Then hso.Visible = False
End If
If Not vso Is Nothing Then
If vso.Visible Then
'vso.Enabled = False
vso.Max = Height * aa - where.ScaleHeight \ DYP
vso.smallchange = 1
vso.largechange = where.ScaleHeight \ DYP
vso.Refresh
redY = -(vso.Value)
'vso.Enabled = True
Else
vso.Value = 0
redY = 0
End If
End If
If (where.ScaleWidth - vso.Width) > 0 Then vso.move where.ScaleWidth - vso.Width, 0, vso.Width, where.ScaleHeight

ThumbnailPaint where, ss * 100, , , , , -redX / DXP / 2, -(((where.ScaleHeight) / DYP) - Int(Height * aa)) / 2 + redY
Else
If (where.ScaleHeight / DYP) / (Height) > 1 Then
If Not vso Is Nothing Then
vso.Visible = False: redY = 0
If Not hso Is Nothing Then
If (where.ScaleHeight - hso.Height) > 0 Then hso.move 0, where.ScaleHeight - hso.Height, where.ScaleWidth
End If
End If
Else
If Not vso Is Nothing Then
If vso.Visible = False Then
vso.Visible = True: vso.enabled = False: vso.Value = 0: vso.enabled = True
End If
vso.Visible = True
VredX = redX
End If
End If

If (where.ScaleWidth / DXP) / (Width) > 1 Then
If Not hso Is Nothing Then
hso.Visible = False: redX = 0
If Not vso Is Nothing Then
If (where.ScaleWidth - vso.Width) > 0 Then vso.move where.ScaleWidth - vso.Width, 0, vso.Width, where.ScaleHeight
End If
End If
Else
If Not hso Is Nothing Then
If hso.Visible = False Then
hso.Visible = True: hso.enabled = False: hso.Value = 0: hso.enabled = True
End If
hso.Visible = True
VredY = redY
End If
End If
If Not vso Is Nothing Then
If vso.Visible Then
vso.Max = Height - (where.ScaleHeight - VredY) \ DYP
vso.smallchange = 1
If (where.ScaleHeight - VredY) \ DYP > 0 Then vso.largechange = (where.ScaleHeight - VredY) \ DYP
vso.Refresh
redY = -(vso.Value)

End If
End If
If Not hso Is Nothing Then
If hso.Visible Then
hso.Max = Width - (where.ScaleWidth - VredX) \ DXP
hso.smallchange = 1
If (where.ScaleWidth - VredX) \ DXP > 0 Then hso.largechange = (where.ScaleWidth - VredX) \ DXP
hso.Refresh
redX = -(hso.Value)
End If
End If
If Not vso Is Nothing Then If (where.ScaleWidth - vso.Width) > 0 And (where.ScaleHeight - VredY) > 0 Then vso.move where.ScaleWidth - vso.Width, 0, vso.Width, where.ScaleHeight - VredY
If Not hso Is Nothing Then If (where.ScaleHeight - hso.Height) > 0 And (where.ScaleWidth - VredX) > 0 Then hso.move 0, where.ScaleHeight - hso.Height, where.ScaleWidth - VredX

where.Cls
aa = 1
ThumbnailPaint where, -1, False, , , , -(((where.ScaleWidth) / DXP) - Width) / 2 + redX, -(((where.ScaleHeight) / DYP) - Height) / 2 + redY

End If
there1:
If Not hso Is Nothing And Not vso Is Nothing Then
If TypeOf where Is PictureBox Then
where.Parent.gonia(gid).Visible = vso.Visible And hso.Visible
Else
where.gonia(gid).Visible = vso.Visible And hso.Visible
End If
End If
where.Refresh
once = False
End Sub

Public Sub ThumbnailPaintdib(DestPic As cDIBSection, Optional ByVal zf As Single = -1, Optional ByVal FitToPic As Boolean = True, Optional CROP As Boolean = False, Optional over As Boolean = False, Optional ByVal fr As Long = 0, Optional ByVal offsetX As Long = 0, Optional ByVal offsetY As Long, Optional ByVal Border As Long = 0, Optional ByVal rtxy As Double = 1#)
   Dim oldSc As Long, Dummy1 As Boolean, Dummy2 As Boolean, sw As Long, sh As Long
   Dim swProbe As Long, shProbe As Long, osh As Double
   Dim Myleft As Long, mytop As Long, mywidth As Long, myheight As Long
   If Width = 0 Then GoTo tp1
   On Error GoTo tp1
     Dim A As Single
     
' this is only for redraw purposes.........bug of windows

     swProbe = DestPic.Width
   shProbe = DestPic.Height
 

    sw = DestPic.Width
   sh = DestPic.Height
 If sh * sw = 0 Then Exit Sub
   
   Dim osw As Long
   
If zf <= 0 Then

      
        If Not FitToPic Then
        DestPic.needHDC
            StretchPictureH DestPic.HDC1, (sw - Width) / 2 + offsetX, (sh - Height) / 2 + offsetY, , , , , , , , rtxy
            DestPic.FreeHDC
        Else
      If sw < sh Then
    osw = offsetX: offsetX = -offsetY:
    offsetY = osw
    
     rtxy = 1 / rtxy
   End If
            If CROP Then
                   
                    sh = sh - 2 * fr
                    sw = sw - 2 * fr
                    If shProbe < swProbe Then
                    osh = shProbe / sh
                        shProbe = CLng(sw * shProbe / swProbe) - Border * 2 / osh
                        swProbe = sw - Border * 2 * osh
                   Else
                   osh = swProbe / sw
            
                        swProbe = CLng(sh * swProbe / shProbe) - 2 * Border / osh
                        shProbe = sh - 2 * Border * osh
                   End If
                   ' Τελευταία αλλαγή - έφυγαν τα CLNG(border/osh) από τα x,y για να τυπώνει σωστά τα offset
                 If (swProbe / Width) > (shProbe / Height) Then
                    zf = Width / swProbe
                    If swProbe > shProbe Then
                    DestPic.needHDC
                        StretchPictureH DestPic.HDC1, CLng(fr) + offsetX + CLng(Border * osh), CLng((sh - sw * shProbe / swProbe) / 2 + fr) + offsetY + Border, sw - Border * 2, CLng(sw * shProbe / swProbe) - Border * 2, 0, CLng((Height - shProbe * zf)) / 2, Width, CLng(shProbe * zf), , rtxy
                    DestPic.FreeHDC
                    Else
                    DestPic.needHDC
                        StretchPictureH DestPic.HDC1, CLng(fr) + CLng((sw - sh * swProbe / shProbe) / 2) + offsetX + CLng(Border / osh), CLng(fr) + offsetY + Border, CLng(sh * swProbe / shProbe) - 2 * Border, sh - 2 * Border, 0, CLng((Height - shProbe * zf)) / 2, Width, CLng(shProbe * zf), , rtxy
                        DestPic.FreeHDC
                    End If
                Else
                    zf = Height / shProbe
                    If swProbe > shProbe Then
                    DestPic.needHDC
                     StretchPictureH DestPic.HDC1, CLng(fr) + offsetX + CLng(Border * osh), CLng((sh - sw * shProbe / swProbe) / 2 + fr) + offsetY + Border, sw - Border * 2, CLng(sw * shProbe / swProbe) - Border * 2, (Width - swProbe * zf) / 2, 0, CLng(swProbe * zf), Height, , rtxy
                    DestPic.FreeHDC
                    Else
                     DestPic.needHDC
                     StretchPictureH DestPic.HDC1, CLng(fr) + CLng((sw - sh * swProbe / shProbe) / 2) + offsetX + Border, CLng(fr) + offsetY + CLng(Border * osh), CLng(sh * swProbe / shProbe) - 2 * Border, sh - 2 * Border, (Width - swProbe * zf) / 2, 0, CLng(swProbe * zf), Height, , rtxy
                     DestPic.FreeHDC
                 End If
                  End If
             Else
                If (swProbe / Width) < (shProbe / Height) Then
                    zf = sw / Width
                Else
                    zf = sh / Height
                End If
                DestPic.needHDC
                StretchPictureH DestPic.HDC1, (sw - Width * zf) / 2 + offsetX, (sh - Height * zf) / 2 + offsetY, CLng(Width * zf), CLng(Height * zf), , , , , , rtxy
                DestPic.FreeHDC
            End If
        End If
    Else
        zf = zf / 100#
        If Not FitToPic Then
        'offsetx = (DestPic.WIDTH - width * zf) * (offsetx / 100#)
        'offsety = (DestPic.HEIGHT - height * zf) * (offsety / 100#)
        DestPic.needHDC
            StretchPictureH DestPic.HDC1, offsetX + Abs((DestPic.Width - Width * zf) / 2), offsetY + Abs((DestPic.Height - Height * zf) / 2), Width * zf, Height * zf, , , , , , rtxy
            DestPic.FreeHDC
                    Else
            
            If DestPic.Width / Width < DestPic.Height / Height Then
                zf = zf * DestPic.Width / Width
                
            Else
                zf = zf * DestPic.Height / Height
            End If
            On Error Resume Next
     '       offsetx = offsetx * zf
        '    offsety = offsety * zf
  If Border > 0 Then
If zf = 1 Then
'DestPic.Grami offsetx + border \ 2, offsety + border \ 2, offsetx + width - border \ 2, offsety + height * zf - border \ 2, border, backcolor

Else
DestPic.square offsetX + (DestPic.Width - Width * zf) / 2 - Border, offsetY + (DestPic.Height - Height * zf) / 2 - Border, offsetX + (DestPic.Width - Width * zf) / 2 + Width * zf + Border, offsetY + (DestPic.Height - Height * zf) / 2 + Height * zf + Border, BackColor
End If
End If
    
        DestPic.needHDC
                StretchPictureH DestPic.HDC1, offsetX + (DestPic.Width - Width * zf) / 2, offsetY + (DestPic.Height - Height * zf) / 2, Width * zf, Height * zf, , , , , , rtxy
               DestPic.FreeHDC
  If Border > 0 Then
If zf = 1 Then
Border = Border + 1
DestPic.square offsetX, offsetY, offsetX + Width - 1, offsetY + Border, BackColor '7
DestPic.square offsetX + Width - Border - 1, offsetY, offsetX + Width - 1, offsetY + Height - 1, BackColor
DestPic.square offsetX, offsetY + Height, offsetX + Width - 1, offsetY + Height - 2 - Border, BackColor  '8
DestPic.square offsetX, offsetY, offsetX + Border, offsetY + Height, BackColor

Else


End If
End If
               
        End If
    End If
    
tp1:
  If Err.Number > 0 Then zf = 1

    last_zf = zf
    ' gdiflush
End Sub

Function ConvertWmftoEmF(hmf As Long) As Long
Dim hEmf As Long
Dim bytes As Long
Dim Buffer() As Byte
bytes = GetMetaFileBitsEx(hmf, bytes, ByVal 0&)
If bytes Then
ReDim Buffer(1 To bytes) As Byte
Call GetMetaFileBitsEx(hmf, bytes, Buffer(1))
Else

Exit Function
End If

' Convert to enhanced format.
needHDC
ConvertWmftoEmF = SetWinMetaFileBits(bytes, Buffer(1), HDC1, ByVal 0&)
FreeHDC
End Function

Public Sub CopyDIBfromClipboard()
Dim hMem As Long, hDib As Long
Const CF_DIB = 8
Dim isbitmap As Boolean, okb As Boolean
                   
                        
  If (OpenClipboard(0) <> 0) Then
    hMem = GetClipboardData(CF_DIB)
            If hMem <> 0 Then
              hDib = GlobalLock(hMem)
              ClearUp
              okb = CreateFromDIB(hDib)
                     If Not okb Then
                        hDib = GlobalUnlock(hMem)
                        CloseClipboard
                        If Clipboard.GetFormat(2) Then
                            CreateFromPicture Clipboard.GetData(2)
                            okb = Height
                        End If
                    End If
                    If okb Then ' ----------open
                            If bitsPerPixel <> 24 Then
                               Dim aa As cDIBSection  ' JUST A REFERENCE
                               Set aa = Me
                               Conv24 aa
                            End If
                              If dpix = 0 Or dpiy = 0 Then GetDpi 96, 96
                              
                    End If
                    
            Call GlobalUnlock(hMem)
           
        End If
        CloseClipboard
    End If
End Sub


