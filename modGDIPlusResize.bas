Attribute VB_Name = "modGDIPlusResize"
Option Explicit

Private Type GUID
   data1    As Long
   data2    As Integer
   data3    As Integer
   data4(7) As Byte
End Type

Private Type PICTDESC
   Size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type
Public Type PICTDESC_META
  cbSizeOfStruct As Long
  PicType As Long
  hMeta As Long
  xExt As Long
  yExt As Long
End Type
Public Enum RotateFlipType
    RotateNoneFlipNone = 0
    Rotate90FlipNone = 1
    Rotate180FlipNone = 2
    Rotate270FlipNone = 3

    RotateNoneFlipX = 4
    Rotate90FlipX = 5
    Rotate180FlipX = 6
    Rotate270FlipX = 7

    RotateNoneFlipY = Rotate180FlipX
    Rotate90FlipY = Rotate270FlipX
    Rotate180FlipY = RotateNoneFlipX
    Rotate270FlipY = Rotate90FlipX

    RotateNoneFlipXY = Rotate180FlipNone
    Rotate90FlipXY = Rotate270FlipNone
    Rotate180FlipXY = RotateNoneFlipNone
    Rotate270FlipXY = Rotate90FlipNone
End Enum
Public Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As RotateFlipType) As Long
Public Type PICTDESC_EMETA
  cbSizeOfStruct As Long
  PicType As Long
  hEmf As Long
End Type
Private Type IID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(0 To 7)  As Byte
End Type
Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PWMFRect16
    Left   As Integer
    top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hmf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type
Public Enum CompositingMode
   CompositingModeSourceOver
   CompositingModeSourceCopy
End Enum
Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal graphics As Long, ByVal CompositingMd As CompositingMode) As Long
Private Declare Function OleLoadPicture Lib "olepro32" _
                              (pStream As Any, _
                              ByVal lSize As Long, _
                              ByVal fRunmode As Long, _
                              riid As Any, _
                              ppvObj As Any) As Long
Const CF_DIB = 8


Private Declare Function SetEnhMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpData As Any) As Long
Private Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hmf As Long, ByVal cbBuffer As Long, lpemh As Any) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long


Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
' GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Sub OleCreatePictureIndirect2 Lib "OleAut32.dll" Alias "OleCreatePictureIndirect" _
    (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, _
    lplpvObj As Object)
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long) As Long

' GDI+ functions
Public Enum GP_MetafileType
    GP_MT_Invalid = 0
    GP_MT_Wmf = 1
    GP_MT_WmfPlaceable = 2
    GP_MT_Emf = 3              'Old-style EMF consisting only of GDI commands
    GP_MT_EmfPlus = 4          'New-style EMF+ consisting only of GDI+ commands
    GP_MT_EmfDual = 5          'New-style EMF+ with GDI fallbacks for legacy rendering
End Enum
Private Type GDI_SizeL
    cx As Long
    cy As Long
End Type
Private Type GDI_MetaHeader
    mtType As Integer
    mtHeaderSize As Integer
    mtVersion As Integer
    mtSize As Long
    mtNoObjects As Integer
    mtMaxRecord As Long
    mtNoParameters As Integer
End Type

Private Type GDIP_EnhMetaHeader3
    iType As Long
    nSize As Long
    rclBounds As RECT1
    rclFrame As RECT1
    dSignature As Long
    nVersion As Long
    nBytes As Long
    nRecords As Long
    nHandles As Integer
    sReserved As Integer
    nDescription As Long
    offDescription As Long
    nPalEntries As Long
    szlDevice As GDI_SizeL
    szlMillimeters As GDI_SizeL
End Type

Private Type GP_MetafileHeader_UNION
'    muWmfHeader As GDI_MetaHeader
    muEmfHeader As GDIP_EnhMetaHeader3
End Type
Private Type GP_MetafileHeader
    mfType As GP_MetafileType
    mfSize As Long
    mfVersion As Long
    mfEmfPlusFlags As Long
    mfDpiX As Single
    mfDpiY As Single
    mfBoundsX As Long
    mfBoundsY As Long
    mfBoundsWidth As Long
    mfBoundsHeight As Long
    mfOrigHeader As GP_MetafileHeader_UNION
    mfEmfPlusHeaderSize As Long
    mfLogicalDpiX As Long
    mfLogicalDpiY As Long
End Type


Private Declare Function GdipGetMetafileHeaderFromMetafile Lib "gdiplus" (ByVal hMetafile As Long, ByRef dstHeader As GP_MetafileHeader) As Long

Private Declare Function GdipCreateRegionHrgn Lib "GdiPlus.dll" (ByVal hRgn As Long, Region As Long) As Long
Private Declare Function GdipDeleteRegion Lib "GdiPlus.dll" (ByVal Region As Long) As Long
Private Declare Function GdipGetImageType Lib "gdiplus" (ByVal Image As Long, ImageType As Long) As Long
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single) As Long
Private Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hWnd As Long, graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imgAttr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imgAttr As Long, ByVal clrAdjust As Long, ByVal clrAdjustEnabled As Long, ByRef clrMatrix As Any, ByRef grayMatrix As Any, ByVal clrMatrixFlags As Long) As Long
Private Declare Function GdipSetImageAttributesColorKeys Lib "GdiPlus.dll" (ByVal mImageattr As Long, ByVal mType As Long, ByVal mEnableFlag As Long, ByVal mColorLow As Long, ByVal mColorHigh As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (token As Long, gdipInput As GDIPlusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long

Private Declare Function GdipConvertToEmfPlus Lib "gdiplus" (ByVal hGraphics As Long, ByVal srcMetafile As Long, ByRef conversionSuccess As Long, ByVal typeOfEMF As Long, ByVal ptrToMetafileDescription As Long, ByRef dstMetafilePtr As Long) As Long

Private Declare Function GdipGetHemfFromMetafile Lib "gdiplus" (ByVal metafile As Long, hEmf As Long) As Long

Private Declare Function GdipDrawImageRectRectI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal Img As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "GdiPlus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "GdiPlus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, metafile As Long) As Long

Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal token As Long)
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipSetPenEndCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mCap As Long) As Long
Private Declare Function GdipSetPenStartCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mCap As Long) As Long
Private Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal Count As Long) As Long
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal pen As Long, ByVal lnJoin As Long) As Long
Private Declare Function GdipFillPolygon2I Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long) As Long
Private Declare Function GdipCreateHatchBrush Lib "GdiPlus.dll" (ByVal mHatchStyle As Long, ByVal mForecol As Long, ByVal mBackcol As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mDashStyle As Long) As Long
Private Declare Function GdipDrawCurveI Lib "gdiplus" (ByVal graphics As Long, ByVal mPen As Long, Points As Any, ByVal Count As Long) As Long
Private Declare Function GdipDrawCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal mPen As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As Long
Private Declare Function GdipFillClosedCurveI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long) As Long
Private Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long, ByVal tension As Single, ByVal FillMd As Long) As Long
Private Declare Function GdipDrawBeziersI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As Any, ByVal Count As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As Long, path As Long) As Long
Private Declare Function GdipAddPathBeziersI Lib "gdiplus" (ByVal path As Long, Points As Any, ByVal Count As Long) As Long
Private Declare Function GdipFillRegion Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal Region As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal path As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal graphics As Long, ByVal mPen As Long, ByVal path As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal mPen As Long, ByVal X As Long, ByVal Y As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawPieI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillPie Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFlush Lib "gdiplus" (ByVal graphics As Long, ByVal intention As Long) As Long
Private Declare Function GdipResetPageTransform Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal Brush As Long, ByVal argb As Long) As Long
Private Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal Brush As Long, argb As Long) As Long
Private Declare Function GdipCreateLineBrushI Lib "GdiPlus.dll" (point1 As Any, point2 As Any, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMode As Long, hBrush As Long) As Long

Private Declare Function GdipSetLineColors Lib "gdiplus" (ByVal Brush As Long, ByVal color1 As Long, ByVal color2 As Long) As Long
Private Declare Function GdipGetLineColors Lib "gdiplus" (ByVal Brush As Long, lColors As Long) As Long
Private Declare Function CreateStreamOnHGlobalolaf Lib "ole32" Alias "CreateStreamOnHGlobal" (ByVal hGlob&, ByVal fDeleteOnRelease As Long, ppstm As stdole.IUnknown) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" _
    (ByVal hGlobal As Any, ByVal fDeleteOnRelease As Long, _
    ByRef ppstm As Any) As Long
    ' ----==== GDI+ Enums ====----
Private Enum status 'GDI+ Status
    ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal pStream As Long, ClsIdEnc As Any, EncParms As Any) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" _
    (ByVal Stream As Any, ByRef Image As Long) As status
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
    (ByVal Bitmap As Long, ByRef hbmReturn As Long, _
    ByVal Background As Long) As status
' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2
Public InitOk As Boolean, myToken As Long
Private MetafileHeader1 As GP_MetafileHeader
Private Declare Function IStream_Size Lib "shlwapi" (ByVal pStream As Long, SizeLongLong As Any) As Long
Private Declare Function IStream_Reset Lib "shlwapi" (ByVal pStream As Long) As Long
Private Declare Function IStream_Read Lib "shlwapi" (ByVal pStream As Long, pBytes As Any, ByVal cb As Long) As Long

Public Function SaveImageToPngByteArray(gimage As Long) As Byte()
  SaveImageToPngByteArray = "" 'return an initialized, byt empty (0 to -1) ByteArray in case of an error
  Dim Strm As stdole.IUnknown, SizeLL(0 To 1) As Long, b() As Byte
  CreateStreamOnHGlobalolaf 0, 1, Strm
  If Strm Is Nothing Then Exit Function
  Dim mPngID(0 To 15) As Byte
  CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), mPngID(0) 'PngWriter-ClsID
  GdipSaveImageToStream gimage, ObjPtr(Strm), mPngID(0), ByVal 0&
  IStream_Size ObjPtr(Strm), SizeLL(0)
  If SizeLL(0) > 0 Then ReDim b(0 To SizeLL(0) - 1) Else Exit Function
  
  IStream_Reset ObjPtr(Strm)
  IStream_Read ObjPtr(Strm), b(0), SizeLL(0)
  SaveImageToPngByteArray = b
End Function
' Initialises GDI Plus
Private Sub SetTokenNow()
Exit Sub
If InitOk = 0 Then
myToken = InitGDIPlus()
InitOk = 1
'Else
'InitOk = InitOk + 1
End If
End Sub
Private Sub ResetTokenNow()
Exit Sub
If InitOk > 0 Then
InitOk = InitOk - 1
If InitOk = 0 Then FreeGDIPlus myToken
End If
End Sub
Public Sub ResetTokenFinal()
Exit Sub
If InitOk > 0 Then
InitOk = 0
FreeGDIPlus myToken
End If
End Sub
Public Function InitGDIPlus() As Long
    Dim token    As Long
    On Error GoTo Err1
    Dim gdipInit As GDIPlusStartupInput
    
    gdipInit.GdiPlusVersion = 1
    GdiplusStartup token, gdipInit, ByVal 0&
    InitGDIPlus = token
    Exit Function
Err1:
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(token As Long)
    GdiplusShutdown token
End Sub

' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(picFile As String, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite) As IPicture
    Dim hDC     As Long
    Dim hBitmap As Long
    Dim Img     As Long

    ' Load the image
    If GdipLoadImageFromFile(StrPtr(picFile), Img) <> 0 Then
        
        'Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        MyEr "GDI+ - can't load picture", "GDI+ - δεν μπορώ να φορώτσω την εικόνα"
        Exit Function
    End If
    
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GetImageDimension Img, Width, Height
        
    End If
    
    ' Initialise the hDC
    InitDC hDC, hBitmap, BackColor, Width, Height

    ' Resize the picture
    gdipResize Img, hDC, Width, Height
    GdipDisposeImage Img
    
    ' Get the bitmap back
    GetBitmap hDC, hBitmap
    
    ' Create the picture
    Set LoadPictureGDIPlus = gCreatePicture(hBitmap)
    
End Function
' Initialises the hDC to draw
Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
    Dim hBrush As Long
        
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
End Sub

Public Sub DrawLineGdi(hDC As Long, PenColor As Long, ByVal penwidth As Long, DashStyle As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
Dim mPen As Long, graphics As Long
'If DashStyle = 5 Then Exit Sub

GdipCreateFromHDC hDC, graphics
GdipSetSmoothingMode graphics, 4

If penwidth <= 1 Then penwidth = 1
If GdiPlusExec(GdipCreatePen1(PenColor, penwidth, UnitPixel, mPen)) = ok Then
    GdipSetPenEndCap mPen, 2
    GdipSetPenStartCap mPen, 2
    GdipSetPenDashStyle mPen, DashStyle
    GdipDrawLineI graphics, mPen, x1, y1, x2, y2
    GdipDeletePen mPen
End If
GdipDeleteGraphics graphics

End Sub
Public Sub DrawArcPieGdi(hDC As Long, PenColor As Long, BackColor As Long, ByVal FillStyle As Long, ByVal penwidth As Long, DashStyle As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, ByVal startAngle As Single, ByVal endAngle As Single)
Dim mPen As Long, graphics As Long, mBrush As Long, swap
'If DashStyle = 5 Then Exit Sub

endAngle = Round(MyMod(endAngle / 1.745329E-02!, 360), 4)
startAngle = Round(MyMod(startAngle / 1.745329E-02!, 360), 4)
If endAngle < 0 Then
    endAngle = 360! + endAngle
End If
If startAngle < 0 Then
    startAngle = 360! + startAngle
End If
If startAngle < endAngle Then
    swap = 360! - endAngle + startAngle
Else
    swap = startAngle - endAngle
End If
startAngle = endAngle
endAngle = swap




GdipCreateFromHDC hDC, graphics
GdipSetSmoothingMode graphics, 4
FillStyle = FillStyle - 2
'If DashStyle = 5 Then PenColor = -1
If penwidth <= 1 Then penwidth = 1
If FillStyle = -1 Then
If DashStyle <> 5 Then
    If GdiPlusExec(GdipCreatePen1(PenColor, penwidth, UnitPixel, mPen)) = ok Then
        GdipSetPenEndCap mPen, 2
        GdipSetPenStartCap mPen, 2
        GdipSetPenDashStyle mPen, DashStyle
        GdipDrawArcI graphics, mPen, x1, y1, x2, y2, startAngle, endAngle
        GdipDeletePen mPen
    End If
End If
Else
If FillStyle = -2 Then
    If GdiPlusExec(GdipCreateSolidFill(BackColor, mBrush)) = ok Then
        If DashStyle <> 5 Then
            If GdiPlusExec(GdipCreatePen1(PenColor, penwidth, UnitPixel, mPen)) = ok Then
                GdipSetPenEndCap mPen, 2
                GdipSetPenStartCap mPen, 2
                GdipSetPenLineJoin mPen, 2
                GdipSetPenDashStyle mPen, DashStyle
                
                GdipFillPie graphics, mBrush, x1, y1, x2, y2, startAngle, endAngle
                GdipDrawPieI graphics, mPen, x1, y1, x2, y2, startAngle, endAngle
                GdipDeletePen mPen
            End If
        Else
                'GdipFillEllipseI graphics, mBrush, x1, y1, x2, y2
                GdipFillPie graphics, mBrush, x1, y1, x2, y2, startAngle, endAngle
        End If
        GdipDeleteBrush mBrush
    End If
Else
If GdiPlusExec(GdipCreateHatchBrush(FillStyle, GDIP_ARGB1(255, BackColor), GDIP_ARGB1(0, BackColor), mBrush)) = ok Then
    If PenColor >= 0 Then
        If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            GdipFillPie graphics, mBrush, x1, y1, x2, y2, startAngle, endAngle
            GdipDrawPieI graphics, mPen, x1, y1, x2, y2, startAngle, endAngle
            GdipDeletePen mPen
        End If
    Else
            GdipDrawPieI graphics, mBrush, x1, y1, x2, y2, startAngle, endAngle
    End If
    GdipDeleteBrush mBrush
End If

End If
End If
GdipDeleteGraphics graphics

End Sub
Public Sub DrawEllipseGdi(hDC As Long, PenColor As Long, BackColor As Long, ByVal FillStyle As Long, ByVal penwidth As Long, DashStyle As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
Dim mPen As Long, graphics As Long, mBrush As Long

GdipCreateFromHDC hDC, graphics
If GDILines Then GdipSetSmoothingMode graphics, 4
FillStyle = FillStyle - 2
'If DashStyle = 5 Then PenColor = -1

If penwidth <= 1 Then penwidth = 1
If FillStyle = -1 Then
If DashStyle <> 5 Then
    If GdiPlusExec(GdipCreatePen1(PenColor, penwidth, UnitPixel, mPen)) = ok Then
        GdipSetPenEndCap mPen, 2
        GdipSetPenStartCap mPen, 2
        GdipSetPenDashStyle mPen, DashStyle
        GdipDrawEllipseI graphics, mPen, x1, y1, x2, y2
        GdipDeletePen mPen
    End If
End If
Else
If FillStyle = -2 Then
    If GdiPlusExec(GdipCreateSolidFill(BackColor, mBrush)) = ok Then
        If DashStyle <> 5 Then
            If GdiPlusExec(GdipCreatePen1(PenColor, penwidth, UnitPixel, mPen)) = ok Then
                GdipSetPenEndCap mPen, 2
                GdipSetPenStartCap mPen, 2
                GdipSetPenLineJoin mPen, 2
                GdipSetPenDashStyle mPen, DashStyle
                If BackColor <> 0 Then GdipFillEllipseI graphics, mBrush, x1, y1, x2, y2
                GdipDrawEllipseI graphics, mPen, x1, y1, x2, y2
                GdipDeletePen mPen
            End If
        Else
                GdipFillEllipseI graphics, mBrush, x1, y1, x2, y2
        End If
        GdipDeleteBrush mBrush
    End If
Else
'If GdiPlusExec(GdipCreateHatchBrush(fillstyle, (backcolor And &HFFFFFF) Or (PenColor And &HFF000000), PenColor And &HFF000000, mBrush)) = ok Then
If GdiPlusExec(GdipCreateHatchBrush(FillStyle, BackColor Or &HFF000000, GDIP_ARGB1(0, 0), mBrush)) = ok Then
    If DashStyle <> 5 Then
        If GdiPlusExec(GdipCreatePen1(PenColor Or &HFF000000, penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            If BackColor <> 0 Then GdipFillEllipseI graphics, mBrush, x1, y1, x2, y2
            GdipDrawEllipseI graphics, mPen, x1, y1, x2, y2
            GdipDeletePen mPen
        
        End If
    Else
            GdipFillEllipseI graphics, mBrush, x1, y1, x2, y2
    End If
    GdipDeleteBrush mBrush
End If

End If
End If
GdipDeleteGraphics graphics

End Sub
'
Public Sub DrawBezierGdi(hDC As Long, ByVal PenColor As Long, BackColor As Long, ByVal FillStyle As Long, ByVal penwidth As Long, DashStyle As Long, Points() As POINTAPI, Count As Long)
Dim mPen As Long, graphics As Long, mBrush As Long, mPath As Long

GdipCreateFromHDC hDC, graphics
If GDILines Then GdipSetSmoothingMode graphics, 4
FillStyle = FillStyle - 2
'If DashStyle = 5 Then PenColor = -1
If penwidth <= 1 Then penwidth = 1
If FillStyle = -1 Then
If DashStyle <> 5 Then
    If GdiPlusExec(GdipCreatePen1(PenColor, penwidth, UnitPixel, mPen)) = ok Then
        GdipSetPenEndCap mPen, 2
        GdipSetPenStartCap mPen, 2
        GdipSetPenLineJoin mPen, 2
        GdipSetPenDashStyle mPen, DashStyle
        GdipDrawBeziersI graphics, mPen, ByVal VarPtr(Points(0)), Count
        GdipDeletePen mPen
    End If
End If
Else
If FillStyle = -2 Then
If GdiPlusExec(GdipCreateSolidFill(BackColor, mBrush)) = ok Then
    If DashStyle <> 5 Then
        If GdiPlusExec(GdipCreatePen1(PenColor, penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            If GdiPlusExec(GdipCreatePath(1, mPath)) = ok Then
                GdipAddPathBeziersI mPath, ByVal VarPtr(Points(0)), Count
                If BackColor <> 0 Then GdipFillPath graphics, mBrush, mPath
                GdipDrawPath graphics, mPen, mPath
                GdipDeletePath mPath
            End If
            GdipDeletePen mPen
        End If
    Else
            If GdiPlusExec(GdipCreatePath(1, mPath)) = ok Then
                GdipAddPathBeziersI mPath, ByVal VarPtr(Points(0)), Count
                GdipFillPath graphics, mBrush, mPath
                GdipDeletePath mPath
            End If
    End If
    GdipDeleteBrush mBrush
End If
Else
If GdiPlusExec(GdipCreateHatchBrush(FillStyle, BackColor Or &HFF000000, GDIP_ARGB1(0, 0), mBrush)) = ok Then
    If DashStyle <> 5 Then
        If GdiPlusExec(GdipCreatePen1(PenColor Or &HFF000000, penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            If GdiPlusExec(GdipCreatePath(1, mPath)) = ok Then
                GdipAddPathBeziersI mPath, ByVal VarPtr(Points(0)), Count
                If FillStyle <> 1 Then GdipFillPath graphics, mBrush, mPath
                GdipDrawPath graphics, mPen, mPath
                GdipDeletePath mPath
            End If
            GdipDeletePen mPen
        End If
    ElseIf FillStyle <> 1 Then
            If GdiPlusExec(GdipCreatePath(1, mPath)) = ok Then
                GdipAddPathBeziersI mPath, ByVal VarPtr(Points(0)), Count
                GdipFillPath graphics, mBrush, mPath
                GdipDeletePath mPath
            End If
    End If
    GdipDeleteBrush mBrush
End If

End If
End If
GdipDeleteGraphics graphics

End Sub
Public Sub DrawPolygonGdi(hDC As Long, ByVal PenColor As Long, BackColor As Long, ByVal FillStyle As Long, ByVal penwidth As Long, DashStyle As Long, Points() As POINTAPI, Count As Long)
Dim mPen As Long, graphics As Long, mBrush As Long

GdipCreateFromHDC hDC, graphics
If GDILines Then GdipSetSmoothingMode graphics, 4
FillStyle = FillStyle - 2
'If DashStyle = 5 Then PenColor = -1
If penwidth <= 1 Then penwidth = 1
If FillStyle = -1 Then
If DashStyle <> 5 Then
    If GdiPlusExec(GdipCreatePen1(PenColor, penwidth, UnitPixel, mPen)) = ok Then
        GdipSetPenEndCap mPen, 2
        GdipSetPenStartCap mPen, 2
        GdipSetPenLineJoin mPen, 2
        GdipSetPenDashStyle mPen, DashStyle
        GdipDrawLinesI graphics, mPen, ByVal VarPtr(Points(0)), Count
        GdipDeletePen mPen
    End If
End If
Else
If FillStyle = -2 Then
If GdiPlusExec(GdipCreateSolidFill(BackColor, mBrush)) = ok Then
    If DashStyle <> 5 Then
        If GdiPlusExec(GdipCreatePen1(PenColor, penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            If BackColor <> 0 Then GdipFillPolygon2I graphics, mBrush, ByVal VarPtr(Points(0)), Count  ' graphics, mPen, x1, y1, x2, y2
            GdipDrawLinesI graphics, mPen, ByVal VarPtr(Points(0)), Count
            GdipDeletePen mPen
        End If
    Else
            If BackColor <> 0 Then GdipFillPolygon2I graphics, mBrush, ByVal VarPtr(Points(0)), Count   ' graphics, mPen, x1, y1, x2, y2
    End If
    GdipDeleteBrush mBrush
End If
Else
If GdiPlusExec(GdipCreateHatchBrush(FillStyle, BackColor Or &HFF000000, GDIP_ARGB1(0, 0), mBrush)) = ok Then
    If DashStyle <> 5 Then
        If GdiPlusExec(GdipCreatePen1(PenColor Or &HFF000000, penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            If FillStyle <> 1 Then GdipFillPolygon2I graphics, mBrush, ByVal VarPtr(Points(0)), Count ' graphics, mPen, x1, y1, x2, y2
            GdipDrawLinesI graphics, mPen, ByVal VarPtr(Points(0)), Count
            GdipDeletePen mPen
        End If
    ElseIf FillStyle <> 1 Then
            GdipFillPolygon2I graphics, mBrush, ByVal VarPtr(Points(0)), Count  ' graphics, mPen, x1, y1, x2, y2
    End If
    GdipDeleteBrush mBrush
End If

End If
End If
GdipDeleteGraphics graphics

End Sub
' Resize the picture using GDI plus
Private Sub gdipResize(Img As Long, hDC As Long, Width As Long, Height As Long)
    Dim graphics   As Long      ' Graphics Object Pointer
    Dim orWidth    As Long      ' Original Image Width
    Dim orHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
        GdipCreateFromHDC hDC, graphics
        GdipSetInterpolationMode graphics, 0

        GdipDrawImageRectI graphics, Img, 0, 0, Width, Height

    GdipDeleteGraphics graphics
End Sub
Private Sub gdipDrawToXYsimple(Img As Long, hDC As Long, DestX As Long, DestY As Long, Width As Long, Height As Long)
    Dim graphics   As Long      ' Graphics Object Pointer
    GdipCreateFromHDC hDC, graphics
   ' GdipSetInterpolationMode graphics, InterpolationModeHighQualityBicubic
    GdipSetPixelOffsetMode graphics, 0 ' 2  '4
    GdipDrawImageRectI graphics, Img, DestX, DestY, Width, Height
    GdipDeleteGraphics graphics
End Sub
Private Sub gdipResizeToXYsimple(Img As Long, hDC As Long, DestX As Long, DestY As Long, DestWidth As Long, DestHeight As Long, Optional bleft As Long = 0, Optional btop As Long = 0, Optional ByVal bWidth As Long = -1, Optional ByVal bHeight As Long = -1)
    Dim graphics   As Long      ' Graphics Object Pointer
    GdipCreateFromHDC hDC, graphics
   ' GdipSetInterpolationMode graphics, InterpolationModeHighQualityBicubic
    GdipSetPixelOffsetMode graphics, 0 ' 2  '4

    If bWidth = -1 Then bWidth = DestWidth
    If bHeight = -1 Then bHeight = DestHeight
     GdipDrawImageRectRectI graphics, Img, DestX, DestY, DestWidth, DestHeight, bleft, btop, bWidth, bHeight, UnitPixel, 0
    GdipDeleteGraphics graphics
End Sub
Private Function MyMod(r1 As Single, po As Single) As Single
MyMod = r1 - Fix(r1 / po) * po
End Function
Private Sub gdipResizeRotate(Scr As Object, Img As Long, Angle!, x1 As Long, y1 As Long, Width As Long, Height As Long, Optional mleft As Long = 0, Optional mtop As Long = 0, Optional Bitmap As Boolean = False)
    Dim clrMatrix(0 To 4, 0 To 4) As Single, Img2 As Long
    Dim hDC As Long, DestX As Long, DestY As Long, aType As PictureTypeConstants
    
     ' WRARNING  WIDTH AND HEIGHT TWIPS AS INPUT
    Dim graphics   As Long      ' Graphics Object Pointer
    
    Dim orWidth    As Long      ' Original Image Width
    Dim orHeight   As Long      ' Original Image Height
    Dim m_Attr As Long
    
       Const Pi = 3.14159!
    Angle! = -MyMod(Angle!, 360!)
    If Angle! < 0 Then Angle! = Angle! + 360!

     Const ColorAdjustTypeBitmap As Long = &H1&
    
   
    Dim prive As Long
    Dim bWidth As Long, bHeight As Long, a As Long, b As Long, Size As Single, bleft As Long, btop As Long
    Dim bWidth1 As Long, bHeight1 As Long, bLeft1 As Long, bTop1 As Long, SizeY As Single
    
    Dim shiftX As Long, shiftY As Long
    
    
    GetImageDimension Img, orWidth, orHeight, bleft, btop, bWidth, bHeight, shiftX, shiftY
    If bHeight = 0 Then
        GdipGetImageWidth Img, bWidth
        GdipGetImageHeight Img, bHeight
        If orWidth = 0 Or orHeight = 0 Then
        orWidth = bWidth
        orHeight = bHeight
        End If
        If orWidth = 0 Or orHeight = 0 Then Exit Sub
        If Width = -1 And Height <> -1 Then
           Height = Height / dv15
           Size = Height / orHeight
           Width = orWidth * Size
        ElseIf Width <> -1 And Height = -1 Then
           Width = Width / dv15
           Size = Width / orWidth
           Height = orHeight * Size
        ElseIf Width = -1 Then
           Width = orWidth
           Height = orHeight
        Else
           Width = Width / dv15
           Height = Height / dv15
           
        End If
    Bitmap = True
     Else
        If orWidth = 0 Or orHeight = 0 Then
        orWidth = bWidth
        orHeight = bHeight
        End If
        If orWidth = 0 Or orHeight = 0 Then Exit Sub
           Dim swap As Long
           If bleft < 0 Or btop < 0 Then
           swap = shiftX: shiftX = shiftY: shiftY = swap
           swap = Width: Width = Height: Height = swap
          swap = bWidth: bWidth = bHeight: bHeight = swap
          End If
        If Width = -1 And Height <> -1 Then
            Height = Height / dv15
            Size = Height / orHeight
            SizeY = Size
            Width = orWidth * Size
            bHeight1 = bHeight * Size
            bWidth1 = bWidth * Size
            bTop1 = btop * Size
            bLeft1 = bleft * Size
        ElseIf Width <> -1 And Height = -1 Then
            Width = Width / dv15
            Size = Width / orWidth
            SizeY = Size
            Height = orHeight * Size
            bHeight1 = bHeight * Size
            bWidth1 = bWidth * Size
            bTop1 = btop * Size
            bLeft1 = bleft * Size
        ElseIf Width = -1 Then
        Size = 1
        SizeY = Size
            Width = orWidth
            Height = orHeight
            bHeight1 = bHeight
            bWidth1 = bWidth
            bTop1 = btop
            bLeft1 = bleft
        Else
        Size = 1
            Width = Width / dv15
            Height = Height / dv15
            Size = Width / orWidth
            SizeY = Height / orHeight
            bHeight1 = bHeight * Size
            
            bWidth1 = bWidth * SizeY
            bTop1 = btop * Size
            bLeft1 = bleft * SizeY
        End If

            Bitmap = False
     End If
    If orWidth = 0 Or orHeight = 0 Then Exit Sub
    
    If Width = 0 Or Height = 0 Then Exit Sub
    On Error Resume Next
      Dim ax As Long, ay As Long, ax1 As Long, ay1 As Long
       If Not Bitmap Then
  With players(GetCode(Scr))
  If Angle >= 179.8 And Angle < 180.2 Then
           Angle! = 180.21
 
End If
      If swap <> 0 Then
      End If
      If .MAXXGRAPH > .MAXYGRAPH Then ax = .MAXXGRAPH Else ax = .MAXYGRAPH
        ax = (.MAXXGRAPH - .XGRAPH) / dv15 / 2: If ax < 0 Then ax = 0
        ay = (.MAXYGRAPH - .YGRAPH) / dv15 / 2: If ay < 0 Then ay = 0
     
        If (.MAXXGRAPH / dv15 - bWidth) > 0 Then ax = ax + (.MAXXGRAPH / dv15 - bWidth) / 2
        If (.MAXYGRAPH / dv15 - bHeight) > 0 Then ay = ay + (.MAXYGRAPH / dv15 - bHeight) / 2
       If ax < 0 Then ax = 0
       If ay < 0 Then ay = 0
       If ax > ay Then ay = ax Else ax = ay
    End With
    ax1 = ax * Size
    ay1 = ay * SizeY
    
        End If
    GdipCreateFromHDC Scr.hDC, graphics
    GdipTranslateWorldTransform graphics, Scr.ScaleX(x1, 1, 3), Scr.ScaleY(y1, 1, 3), 1  '
    

    GdipRotateWorldTransform graphics, Angle!, 1
  
    With players(GetCode(Scr))
    
    GdipTranslateWorldTransform graphics, Scr.ScaleX(.XGRAPH, 1, 3), Scr.ScaleY(.YGRAPH, 1, 3), 1

    
    GdipSetInterpolationMode graphics, InterpolationModeHighQualityBicubic
    GdipSetPixelOffsetMode graphics, 0  '-4 * (bitmap = True)
    If Bitmap Then
    '  GetImageDimension img, orWidth, orHeight
    GdipDrawImageRectRectI graphics, Img, -Width \ 2, -Height \ 2, Width, Height, mleft, mtop, orWidth, orHeight, UnitPixel, m_Attr
    Else
    
 'GdipDrawImageRectRectI graphics, Img, -bWidth1 \ 2 - 50 * Size, -bHeight1 \ 2 - 50 * Size, bWidth1 + 100 * Size, bHeight1 + 100 * Size, bleft - 50, btop - 50, bWidth + 100, bHeight + 100, UnitPixel, m_Attr
'Debug.Print 100, -bWidth1 \ 2 - 50 * Size, -bHeight1 \ 2 - 50 * Size, bWidth1 + 100 * Size, bHeight1 + 100 * Size, bleft - 50, btop - 50, bWidth + 100, bHeight + 100
'Debug.Print 0, Int(-bWidth1 / 2 - ax1), Int(-bHeight1 / 2 - ay1), bWidth1 + ax1 * 2, bHeight1 + ay1 * 2, bleft - ax, btop - ay, bWidth + ax * 2, bHeight + ay * 2

         GdipDrawImageRectRectI graphics, Img, -bWidth1 / 2 - ax1, -bHeight1 / 2 - ay1, bWidth1 + ax1 * 2, bHeight1 + ay1 * 2, bleft - ax, btop - ay, bWidth + ax * 2, bHeight + ay * 2, UnitPixel, m_Attr
      
     End If
     End With
     If m_Attr Then GdipDisposeImageAttributes m_Attr
    
    GdipDeleteGraphics graphics
End Sub


'
Private Sub gdipResizeToXY(bstack As basetask, Img As Long, Angle!, zoomfactor As Single, Alpha!, Optional BackColor As Long = -1, Optional mleft As Long = 0, Optional mtop As Long = 0)
    Dim clrMatrix(0 To 4, 0 To 4) As Single, Img2 As Long
    Dim hDC As Long, DestX As Long, DestY As Long, aType As PictureTypeConstants
    
    
    Dim graphics   As Long      ' Graphics Object Pointer
    Dim Width As Long
    Dim Height As Long
    
    Dim orWidth    As Long      ' Original Image Width
    Dim orHeight   As Long      ' Original Image Height
    Dim m_Attr As Long
    
       Const Pi = 3.14159!
    Angle! = -MyMod(Angle!, 360!)
    If Angle! < 0 Then Angle! = Angle! + 360!
  If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100#

     Const ColorAdjustTypeBitmap As Long = &H1&
    
   
    Dim prive As Long, Scr As Object
    Set Scr = bstack.Owner
    If Alpha! <> 0! Or BackColor <> -1 Then Call GdipCreateImageAttributes(m_Attr)
    

     If Alpha! <> 0! Then
            If clrMatrix(4, 4) = 0! Then
                clrMatrix(0, 0) = 1!: clrMatrix(1, 1) = 1!: clrMatrix(2, 2) = 1!
                clrMatrix(3, 3) = CSng((100! - Alpha!) / 100!) ' global blending; value between 0 & 1
                clrMatrix(4, 4) = 1! ' required; cannot be anything else
            End If
            If GdipSetImageAttributesColorMatrix(m_Attr, ColorAdjustTypeBitmap, 1&, clrMatrix(0, 0), clrMatrix(0, 0), 0&) Then
                    GdipDisposeImageAttributes m_Attr
                    m_Attr = 0&
            End If
    End If
    If m_Attr And BackColor >= 0 Then
     GdipSetImageAttributesColorKeys m_Attr, 1&, 1&, GDIP_ARGB1(0, BackColor), GDIP_ARGB1(255, BackColor)
    End If
    GetImageDimension Img, orWidth, orHeight
    
    Height = orHeight * zoomfactor
    Width = orWidth * zoomfactor
    
   On Error Resume Next
   
    GdipCreateFromHDC bstack.Owner.hDC, graphics
    GdipSetPixelOffsetMode graphics, 0
    GdipRotateWorldTransform graphics, Angle!, 1
    With players(GetCode(Scr))
     GdipTranslateWorldTransform graphics, Scr.ScaleX(.XGRAPH, 1, 3), Scr.ScaleY(.YGRAPH, 1, 3), 1
    End With
    GdipDrawImageRectRectI graphics, Img, -Width \ 2, -Height \ 2, Width, Height, mleft, mtop, orWidth, orHeight, UnitPixel, m_Attr
      'GdipDrawImageRectRectI graphics, img, -Width / 2, -Height / 2, Width, Height, mleft, 1 + mtop, orWidth, orHeight, UnitPixel, m_Attr
      If m_Attr Then GdipDisposeImageAttributes m_Attr
    
    GdipDeleteGraphics graphics
End Sub


' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
Private Sub GetBitmap(hDC As Long, hBitmap As Long)
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
End Sub

' Creates a Picture Object from a handle to a bitmap
Public Function gCreatePicture(hBitmap As Long, Optional aType As Long = PICTYPE_BITMAP) As StdPicture
    Dim IID_IDispatch As GUID
    Dim pic           As PICTDESC
    Dim ipic          As IPicture
    
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.data1 = &H20400
    IID_IDispatch.data4(0) = &HC0
    IID_IDispatch.data4(7) = &H46
        
    ' Fill Pic with necessary parts
    pic.Size = Len(pic)        ' Length of structure
    pic.Type = aType  ' Type of Picture (bitmap)
    pic.hBmp = hBitmap         ' Handle to bitmap

    ' Create the picture
    OleCreatePictureIndirect pic, IID_IDispatch, True, ipic
    Set gCreatePicture = ipic
End Function



' Fills in the wmfPlacable header
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, Width As Long, Height As Long)
    WmfHeader.BoundingBox.Right = Width
    WmfHeader.BoundingBox.Bottom = Height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
End Sub
Public Function ReadSizeImageFromBuffer(ResData() As Byte, Width As Long, Height As Long) As Boolean
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim hDC As Long
    Dim Img As Long
    Dim hBitmap As Long
    
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), False, Stream)
    If Not (Stream Is Nothing) Then
        If GdiPlusExec(GdipLoadImageFromStream(Stream, Img)) = ok Then
        Dim a As Long, b As Long
            GetImageDimension Img, Width, Height, , , a, b
            If a <> 0 Then Width = a - 2: Height = b - 3
            GdipDisposeImage Img
            ReadSizeImageFromBuffer = True
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.Clear
    Resume PROC_EXIT

End Function
Public Function LoadImageFromBuffer2(ResData() As Byte, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional FlipOrRotate As Long = 0) As IPicture
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim hDC As Long
    Dim Img As Long
    Dim hBitmap As Long
    Dim iType As Long
    ' Ressource in ByteArray speichern
    
    ' Stream erzeugen
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), _
    False, Stream)
    
    ' ist ein Stream vorhanden
    If Not (Stream Is Nothing) Then
        
        ' GDI+ Bitmapobjekt vom Stream erstellen
        If GdiPlusExec(GdipLoadImageFromStream( _
        Stream, Img)) = ok Then
                    If FlipOrRotate <> 0 Then
            If GdipGetImageType(Img, iType) = ok Then
                    If iType = 1 Then
                    If GdipImageRotateFlip(Img, (FlipOrRotate)) = ok Then FlipOrRotate = 0
                    End If
            End If
        End If
            
                  If Width = -1 Or Height = -1 Then
                   ' GetImageDimension img, Width, Height
                     GdipGetImageWidth Img, Width
                     GdipGetImageHeight Img, Height
                
                End If
                 ' Initialise the hDC
                  InitDC hDC, hBitmap, BackColor, Width, Height

                ' Resize the picture
                gdipResize Img, hDC, Width, Height
                GdipDisposeImage Img
    
    ' Get the bitmap back
    GetBitmap hDC, hBitmap

    ' Create the picture
    Set LoadImageFromBuffer2 = gCreatePicture(hBitmap)
            
             'GdipDisposeImage Img
            
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.Clear
    Resume PROC_EXIT

End Function
Public Function LoadImageFromBuffer3(ResData() As Byte, Width As Long, Height As Long, Optional ByVal BackColor As Long = vbWhite, Optional magia As Boolean = False) As IPicture
    
    On Error GoTo PROC_ERR
    Dim iType As Long
    Dim Stream As IUnknown
    Dim hDC As Long
    Dim Img As Long
    Dim hBitmap As Long
     Dim orWidth As Long, orHeight As Long, mtop As Long, mleft As Long
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), _
    False, Stream)
    Const GP_IT_Metafile = 2
Dim mWidth As Long, mHeight As Long
    If Not (Stream Is Nothing) Then
        If GdiPlusExec(GdipLoadImageFromStream( _
        Stream, Img)) = ok Then
    If GdipGetImageType(Img, iType) = ok Then
     If iType = GP_IT_Metafile Then
            
                GdipGetMetafileHeaderFromMetafile Img, MetafileHeader1
                If MetafileHeader1.mfOrigHeader.muEmfHeader.dSignature = 1179469088 Then
                mleft = MetafileHeader1.mfOrigHeader.muEmfHeader.rclBounds.Left
                mtop = MetafileHeader1.mfOrigHeader.muEmfHeader.rclBounds.top
                mWidth = MetafileHeader1.mfOrigHeader.muEmfHeader.rclBounds.Right - MetafileHeader1.mfOrigHeader.muEmfHeader.rclBounds.Left + 1
                mHeight = MetafileHeader1.mfOrigHeader.muEmfHeader.rclBounds.Bottom - MetafileHeader1.mfOrigHeader.muEmfHeader.rclBounds.top + 1
                Else
                
                mWidth = MetafileHeader1.mfBoundsWidth
                mHeight = MetafileHeader1.mfBoundsHeight
                
                End If
                If Width = -1 And Height = -1 Then
                    Width = mWidth
                    Height = mHeight
                ElseIf Width = -1 Then
                    Width = Height * mWidth / mHeight
                ElseIf Height = -1 Then
                    Height = Width * mHeight / mWidth
                End If
                
                orHeight = mHeight
                orWidth = mWidth

            If magia Then
                orHeight = mHeight + 1
                orWidth = mWidth + 1
                Width = Width + 1
                Height = Height + 1
                End If
                
            Else
                GdipGetImageWidth Img, orWidth
                GdipGetImageHeight Img, orHeight
            End If
    Else
      GetImageDimension Img, orWidth, orHeight
    End If
            
   If Width = -1 Or Height = -1 Then
            GetImageDimension Img, Width, Height
            End If

   
          InitDC hDC, hBitmap, BackColor, Width, Height

    
              Dim graphics   As Long      ' Graphics Object Pointer
    GdipCreateFromHDC hDC, graphics
    GdipSetInterpolationMode graphics, 7
    GdipSetPixelOffsetMode graphics, 4
    GdipDrawImageRectRectI graphics, Img, 0, 0, Width, Height, mleft, mtop, orWidth, orHeight, UnitPixel, 0, 0, 0
    GdipDeleteGraphics graphics


        
        GetBitmap hDC, hBitmap

    Set LoadImageFromBuffer3 = gCreatePicture(hBitmap)

         GdipDisposeImage Img

            
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.Clear
    Resume PROC_EXIT

End Function
Public Function DrawImageFromBuffer(ResData() As Byte, hDC As Long, Optional X As Long = 0&, Optional Y As Long = 0&, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim Img As Long
    Dim hBitmap As Long
    
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), False, Stream)
    
    If Not (Stream Is Nothing) Then
        If GdiPlusExec(GdipLoadImageFromStream(Stream, Img)) = ok Then
            
        Dim orWidth    As Long      ' Original Image Width
        Dim orHeight   As Long
        Dim bWidth    As Long
        Dim bHeight   As Long
        Dim bleft As Long, btop As Long
        

      GetImageDimension Img, orWidth, orHeight, bleft, btop, bWidth, bHeight
          If bWidth <> 0 Then
             orWidth = bWidth
             orHeight = bHeight
        End If
    
            If Width = -1 Or Height = -1 Then
                If Width = -1 Then
                    Width = orWidth
                    If Height = -1 Then
                        Height = orHeight
                    Else
                        Width = Height * orWidth / orHeight
                    End If
                ElseIf Height = -1 Then
                    Height = Width * orHeight / orWidth
                End If
            End If
        End If
        If bWidth = 0 Then
            gdipResizeToXYsimple Img, hDC, X, Y, Width, Height, bleft, btop, orWidth, orHeight
        Else
            gdipResizeToXYsimple Img, hDC, X, Y, Width, Height, bleft, btop, orWidth, orHeight
        End If
        
        GdipDisposeImage Img
    End If

    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.Clear

    Resume PROC_EXIT

End Function

Public Function GetEmfBoubdsPixels(EmfPtr() As Byte) As Long()
    Dim hEmf As Long, b() As Long, SizeByte As Long, bytes As Long
    SizeByte = UBound(EmfPtr()) - LBound(EmfPtr()) + 1
    hEmf = SetEnhMetaFileBits(SizeByte, EmfPtr(0))
    If hEmf <> 0 Then
    bytes = GetEnhMetaFileHeader(hEmf, bytes, ByVal 0)
    If bytes > 0 Then
        ReDim b(bytes \ 4)
        GetEnhMetaFileHeader hEmf, bytes, b(0)
    End If
    GetEmfBoubdsPixels = b
    DeleteEnhMetaFile hEmf
    End If
End Function
Private Sub GetImageDimension(Img As Long, Width As Long, Height As Long, Optional btop As Long, Optional bleft As Long, Optional bWidth As Long, Optional bHeight As Long, Optional shX As Long, Optional shY As Long)
    Dim iType As Long, mtop As Long, mleft As Long
    Const GP_IT_Metafile = 2
    If GdipGetImageType(Img, iType) = ok Then
            If iType = GP_IT_Metafile Then

            GdipGetMetafileHeaderFromMetafile Img, MetafileHeader1
           
            If MetafileHeader1.mfOrigHeader.muEmfHeader.dSignature = 1179469088 Then
        
           bWidth = MetafileHeader1.mfBoundsWidth
             bHeight = MetafileHeader1.mfBoundsHeight
             bleft = MetafileHeader1.mfBoundsX
             
             btop = MetafileHeader1.mfBoundsY
             
                mleft = MetafileHeader1.mfOrigHeader.muEmfHeader.rclBounds.Left
                mtop = MetafileHeader1.mfOrigHeader.muEmfHeader.rclBounds.top
                Height = bHeight
                Width = bWidth
                mtop = btop
                mleft = bleft
                
                shX = btop - mtop
                shY = bleft - mleft

                Else
                
                bWidth = MetafileHeader1.mfBoundsWidth
                bHeight = MetafileHeader1.mfBoundsHeight
                bleft = MetafileHeader1.mfBoundsX
                btop = MetafileHeader1.mfBoundsY
                Height = bHeight
                Width = bWidth
                shX = btop
                shY = bleft
                'GoTo there1
                End If
            Else
there1:
                GdipGetImageWidth Img, Width
                GdipGetImageHeight Img, Height
                
            End If
    End If
End Sub
Public Function DrawSpriteFromBuffer(bstack As basetask, ResData() As Byte, sprt As Boolean, Angle!, zoomfactor!, blend!, Optional BackColor As Long = -1, Optional IsEmf As Boolean = False) As Boolean
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim Img As Long
    Dim hBitmap As Long, hDC As Long
    
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), False, Stream)
    Dim Width As Long, Height As Long
    If Not (Stream Is Nothing) Then
        If GdiPlusExec(GdipLoadImageFromStream(Stream, Img)) = ok Then
            GetImageDimension Img, Width, Height
            
            If sprt Then GetBackSprite bstack, Width * 2, Height, Angle!, zoomfactor
            If IsEmf Then
            Dim K
            K = GetEmfBoubdsPixels(ResData)
            Const v = 26.4583333333333    '26.3245
            gdipResizeToXY bstack, Img, Angle!, zoomfactor!, blend!, BackColor, K(6) / v, K(7) / v
            Else
            gdipResizeToXY bstack, Img, Angle!, zoomfactor!, blend!, BackColor
            End If
            GdipDisposeImage Img
            
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.Clear

    Resume PROC_EXIT

End Function
Public Function DrawEmfFromBuffer(Scr As Object, ResData() As Byte, X As Long, Y As Long, Angle!, Width As Long, Height As Long) As Boolean
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim Img As Long, iType As Long
    Dim hBitmap As Long, hDC As Long
    
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), False, Stream)
    If Not (Stream Is Nothing) Then
        If GdiPlusExec(GdipLoadImageFromStream(Stream, Img)) = ok Then
            gdipResizeRotate Scr, Img, Angle!, X, Y, (Width), (Height)
            GdipDisposeImage Img
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.Clear

    Resume PROC_EXIT

End Function
Public Function LoadImageFromBuffer( _
 ResData() As Byte, Optional aFlipRotateType) As StdPicture
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim lBitmap As Long
    Dim hBitmap As Long
    Dim iType As Long
    ' Ressource in ByteArray speichern
    
    ' Stream erzeugen
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), _
    False, Stream)
    
    ' ist ein Stream vorhanden
    If Not (Stream Is Nothing) Then
        
        ' GDI+ Bitmapobjekt vom Stream erstellen
        If GdiPlusExec(GdipLoadImageFromStream( _
        Stream, lBitmap)) = ok Then
        If GdipGetImageType(lBitmap, iType) = 0 Then
        If iType = 2 Then
                Set LoadImageFromBuffer = gCreatePicture(lBitmap, iType)
        Else
            ' Handle des Bitmapobjektes ermitteln
            If GdiPlusExec(GdipCreateHBITMAPFromBitmap( _
            lBitmap, hBitmap, 0)) = ok Then
                
                ' StdPicture Objekt erstellen
                If Not IsMissing(aFlipRotateType) Then GdipImageRotateFlip lBitmap, CLng(aFlipRotateType)
                
                
                Set LoadImageFromBuffer = _
                HandleToPicture(hBitmap, iType)
                
            End If
            End If
            End If
            ' Bitmapobjekt lφschen
            Call GdiPlusExec(GdipDisposeImage(lBitmap))
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.Clear

    Resume PROC_EXIT

End Function
Private Function GdiErrorString(ByVal lError As status) As String
    Dim s As String
    
    Select Case lError
    Case GenericError:              s = "Generic Error."
    Case InvalidParameter:          s = "Invalid Parameter."
    Case OutOfMemory:               s = "Out Of Memory."
    Case ObjectBusy:                s = "Object Busy."
    Case InsufficientBuffer:        s = "Insufficient Buffer."
    Case NotImplemented:            s = "Not Implemented."
    Case Win32Error:                s = "Win32 Error."
    Case WrongState:                s = "Wrong State."
    Case Aborted:                   s = "Aborted."
    Case FileNotFound:              s = "File Not Found."
    Case ValueOverflow:             s = "Value Overflow."
    Case AccessDenied:              s = "Access Denied."
    Case UnknownImageFormat:        s = "Unknown Image Format."
    Case FontFamilyNotFound:        s = "FontFamily Not Found."
    Case FontStyleNotFound:         s = "FontStyle Not Found."
    Case NotTrueTypeFont:           s = "Not TrueType Font."
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version."
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized."
    Case PropertyNotFound:          s = "Property Not Found."
    Case PropertyNotSupported:      s = "Property Not Supported."
    Case Else:                      s = "Unknown GDI+ Error."
    End Select
    
    GdiErrorString = s
End Function
Private Function GdiPlusExec(ByVal lReturn As status) As status
    Dim lCurErr As status
    If lReturn = status.ok Then
        lCurErr = status.ok
    Else
        lCurErr = lReturn
        Dim er$
    er$ = "GDI+: " & GdiErrorString(lReturn) & " GDI+ Error:" & lReturn
    MyEr er$, er$
    Err.Clear
    End If
    GdiPlusExec = lCurErr
End Function
Public Function HandleToPictureFromBits(hMem As Long, cbmem As Long) As StdPicture
    

    Dim IID_IDispatch As IID
    Dim oPicture As IPicture
    Dim istm As stdole.IUnknown
    

    
    ' Initialisiert das IPicture Interface ID
    With IID_IDispatch
        .data1 = &H20400
        .data4(0) = &HC0
        .data4(7) = &H46
    End With
    Dim Img As Long
    If (CreateStreamOnHGlobal(hMem, 1, istm) = 0) Then

    OleLoadPicture ByVal ObjPtr(istm), cbmem, 0, IID_IDispatch, oPicture
     
    Set HandleToPictureFromBits = oPicture
 
    End If
End Function
Public Function HandleToPicture(ByVal hGDIHandle As Long, _
    ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hPal As Long = 0) As StdPicture
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As IID
    Dim oPicture As IPicture
    
    ' Initialisiert die PICTDESC Structur
    With tPictDesc
         .Size = Len(tPictDesc)
        .Type = ObjectType
        .hBmp = hGDIHandle
        .hPal = hPal
    End With
    
    ' Initialisiert das IPicture Interface ID
    With IID_IPicture
        .data1 = &H7BF80981
        .data2 = &HBF32
        .data3 = &H101A
        .data4(0) = &H8B
        .data4(1) = &HBB
        .data4(3) = &HAA
        .data4(5) = &H30
        .data4(6) = &HC
        .data4(7) = &HAB
    End With
    
    ' Erzeugen des Objekts
    OleCreatePictureIndirect2 tPictDesc, _
    IID_IPicture, True, oPicture
    
    ' Rόckgabe des Pictureobjekts
    Set HandleToPicture = oPicture
    
End Function
Function GDIP_ARGB(Alpha As Long, Red As Long, Green As Long, Blue As Long) As Long
Dim b As Byte
GetMem1 VarPtr(Alpha), b
PutMem1 VarPtr(GDIP_ARGB) + 3, b
GetMem1 VarPtr(Red), b
PutMem1 VarPtr(GDIP_ARGB) + 2, b
GetMem1 VarPtr(Green), b
PutMem1 VarPtr(GDIP_ARGB) + 1, b
GetMem1 VarPtr(Blue), b
PutMem1 VarPtr(GDIP_ARGB), b
End Function
Function GDIP_ARGB1(Alpha As Long, color As Long) As Long
Dim b As Byte
GetMem1 VarPtr(Alpha), b
PutMem1 VarPtr(GDIP_ARGB1) + 3, b
GetMem1 VarPtr(color) + 2, b
PutMem1 VarPtr(GDIP_ARGB1), b
GetMem1 VarPtr(color) + 1, b
PutMem1 VarPtr(GDIP_ARGB1) + 1, b
GetMem1 VarPtr(color), b
PutMem1 VarPtr(GDIP_ARGB1) + 2, b

End Function
Sub M2000Pen(ByVal Alpha As Long, color As Long)
Dim b As Byte, b1 As Byte
GetMem1 VarPtr(Alpha), b
PutMem1 VarPtr(color) + 3, b
GetMem1 VarPtr(color) + 2, b
GetMem1 VarPtr(color), b1
PutMem1 VarPtr(color) + 2, b1
PutMem1 VarPtr(color), b
End Sub
Sub GdiPlusGradient(hDC As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, ByVal color1 As Long, ByVal color2 As Long, Optional UseVertical As Long)
Dim point(4) As POINTAPI, graphics As Long, hBrush As Long

point(0).X = x1
point(0).Y = y1
point(1).X = x2
point(1).Y = y1
point(2).X = x2
point(2).Y = y2
point(3).X = x1
point(3).Y = y2
point(4).X = x1
point(4).Y = y2
UseVertical = Abs(UseVertical <> 0) * 2
GdipCreateFromHDC hDC, graphics
GdipCreateLineBrushI point(UseVertical), point(1), color2, color1, 1, hBrush
GdipFillPolygon2I graphics, hBrush, point(0), 5
GdipDeleteBrush hBrush
GdipDeleteGraphics graphics
End Sub
Sub GdiPlusGradientRegion(whois As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, ByVal color1 As Long, ByVal color2 As Long, Optional UseVertical As Long)
Dim point(4) As POINTAPI, graphics As Long, hBrush As Long, hRegion As Long, myRgn As Long
myRgn = CreateRectRgn(0&, 0&, 0&, 0&)
If GetWindowRgn(Form1.dSprite(whois).hWnd, myRgn) <> 0 Then
point(0).X = x1
point(0).Y = y1
point(1).X = x2
point(1).Y = y1
point(2).X = x2
point(2).Y = y2
point(3).X = x1
point(3).Y = y2
point(4).X = x1
point(4).Y = y2
UseVertical = Abs(UseVertical <> 0) * 2
GdipCreateFromHDC Form1.dSprite(whois).hDC, graphics
GdipCreateRegionHrgn myRgn, hRegion
GdipCreateLineBrushI point(UseVertical), point(1), color2, color1, 1, hBrush
GdipFillRegion graphics, hBrush, hRegion
GdipDeleteBrush hBrush
GdipDeleteRegion hRegion
GdipDeleteGraphics graphics
End If
'DeleteObject myRgn
End Sub
