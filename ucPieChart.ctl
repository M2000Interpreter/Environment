VERSION 5.00
Begin VB.UserControl ucPieChart 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
   Begin VB.Timer tmrMOUSEOVER 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ucPieChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Module Name: ucChartArea
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 23/06/2020
'Version: 1.0.0
'-----------------------------------------------
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTL) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
'Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
'Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
'Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
'Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
'Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
'Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GdipCreatePen2 Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipFlattenPath Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mMatrix As Long, ByVal mFlatness As Single) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipGetPointCount Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mCount As Long) As Long
Private Declare Function GdipRotatePathGradientTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mAngle As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTranslatePathGradientTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mDx As Single, ByVal mDy As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTranslateTextureTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mDx As Single, ByVal mDy As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipAddPathEllipseI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColor As Long, ByRef mCount As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipSetPenEndCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mEndCap As LineCap) As Long
Private Declare Function GdipSetPenStartCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mStartCap As LineCap) As Long
Private Declare Function GdipDrawEllipse Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipSaveGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByRef mState As Long) As Long
Private Declare Function GdipRestoreGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mState As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RectL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As WrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTF, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As StringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As StringAlignment) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipDrawArc Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathCurveI Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipSetCompositingQuality Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mCompositingQuality As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathCurve2I Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTL, ByVal mCount As Long, ByVal mTension As Single) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipAddPathLine2I Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipDrawLinesI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipDrawCurveI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipFillClosedCurveI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreateLineBrushI Lib "GdiPlus.dll" (ByRef mPoint1 As POINTL, ByRef mPoint2 As POINTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mWrapMode As WrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As CombineMode) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTF, ByVal mStringFormat As Long, ByRef mBoundingBox As RECTF, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
Private Declare Function GdipDrawPie Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipFillPie Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long


Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal graphics As Long, result As Long) As Long
Private Declare Function GdipAddPathPie Lib "gdiplus" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal path As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipGetPathLastPoint Lib "gdiplus" (ByVal path As Long, lastPoint As POINTF) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Type PICTDESC
    lSize               As Long
    lType               As Long
    hBmp                As Long
    hPal                As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



Private Enum CombineMode
    CombineModeReplace = &H0
    CombineModeIntersect = &H1
    CombineModeUnion = &H2
    CombineModeXor = &H3
    CombineModeExclude = &H4
    CombineModeComplement = &H5
End Enum


Private Type POINTL
    X As Long
    Y As Long
End Type

Private Type POINTF
    X As Single
    Y As Single
End Type

Private Type SIZEF
    Width As Single
    Height As Single
End Type

Private Enum LineCap
    LineCapFlat = &H0
    LineCapSquare = &H1
    LineCapRound = &H2
    LineCapTriangle = &H3
    LineCapNoAnchor = &H10
    LineCapSquareAnchor = &H11
    LineCapRoundAnchor = &H12
    LineCapDiamondAnchor = &H13
    LineCapArrowAnchor = &H14
    LineCapCustom = &HFF
    LineCapAnchorMask = &HF0
End Enum

Private Enum MatrixOrder
    MatrixOrderPrepend = &H0
    MatrixOrderAppend = &H1
End Enum

Private Type RECTF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type RectL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Enum CaptionAlignmentH
    cLeft
    cCenter
    cRight
End Enum

Private Enum CaptionAlignmentV
    cTop
    cMiddle
    cBottom
End Enum

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
  
Private Enum StringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum

Private Enum StringFormatFlags
    StringFormatFlagsNone = &H0
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

Private Enum WrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipY = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

Private Const GWL_WNDPROC               As Long = -4
Private Const GW_OWNER                  As Long = 4
Private Const WS_CHILD                  As Long = &H40000000
Private Const UnitPixel                 As Long = &H2&
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const GDIP_OK                   As Long = &H0

Public Event ItemClick(Index As Long)
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event PrePaint(hdc As Long)
'Public Event PostPaint(ByVal hdc As Long)
'Public Event KeyPress(KeyAscii As Integer)
'Public Event KeyUp(KeyCode As Integer, Shift As Integer)
'Public Event KeyDown(KeyCode As Integer, Shift As Integer)
'Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

Public Enum ChartStyle
    CS_PIE
    CS_DONUT
End Enum

Public Enum ChartOrientation
    CO_Vertical
    CO_Horizontal
End Enum

Public Enum ucPC_LegendAlign
    LA_LEFT
    LA_TOP
    LA_RIGHT
    LA_BOTTOM
End Enum

Public Enum LabelsPositions
    LP_Inside
    LP_Outside
    LP_TwoColumns
End Enum

Dim c_lhWnd As Long
Dim nScale As Single

Dim m_Title As String
Dim m_TitleFont As StdFont
Dim m_TitleForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_BackColorOpacity As Long
Dim m_ForeColor As OLE_COLOR
Dim m_FillOpacity As Long
Dim m_Border As Boolean
Dim m_LinesCurve As Boolean
Dim m_LinesWidth As Long
Dim m_VerticalLines As Boolean
Dim m_FillGradient As Boolean

Dim m_HorizontalLines As Boolean
Dim m_ChartStyle As ChartStyle
Dim m_ChartOrientation As ChartOrientation
Dim m_LegendAlign As ucPC_LegendAlign
Dim m_LegendVisible As Boolean
Dim m_WordWrap As Boolean
Dim m_DonutWidth As Single
Dim m_SeparatorLine As Boolean
Dim m_SeparatorLineColor As OLE_COLOR
Dim m_SeparatorLineWidth As Single
Dim m_LabelsVisible As Boolean
Dim m_LabelsPositions As LabelsPositions
Dim m_LabelsFormats As String
Dim m_BorderColor As OLE_COLOR
Dim m_BorderRound As Long
Dim m_Rotation  As Long
Dim m_CenterCircle As POINTF

Private Type tItem
    ItemName As String
    Value As Single
    text As String
    TextWidth As Long
    TextHeight As Long
    ItemColor As Long
    Special As Boolean
    hPath As Long
    LegendRect As RectL
End Type


Dim m_Item() As tItem
Dim ItemsCount As Long
Dim HotItem As Long
Dim m_PT As POINTL
Dim m_Left As Long
Dim m_Top As Long
Dim GdipToken As Long

Public Property Get Image(Optional ByVal Width As Long, Optional ByVal Height As Long) As IPicture
    Dim lDC As Long
    Dim TempDC As Long
    Dim hBmp As Long, OldBmp As Long
    Dim hBrush As Long
    Dim RECT As RECT
    Dim lColor As Long
    Dim uDesc As PICTDESC
    Dim aInput(3) As Long
    
    If Width = 0 Then Width = UserControl.ScaleWidth
    If Height = 0 Then Height = UserControl.ScaleHeight
    
    lDC = GetDC(0&)
    TempDC = CreateCompatibleDC(lDC)
    hBmp = CreateCompatibleBitmap(lDC, Width, Height)
    OldBmp = SelectObject(TempDC, hBmp)
    RECT.Right = Width
    RECT.Bottom = Height
    
    lColor = m_BackColor
    If (lColor And &H80000000) Then lColor = GetSysColor(lColor And &HFF&)
    
    hBrush = CreateSolidBrush(lColor)
    FillRect TempDC, RECT, hBrush
    DeleteObject hBrush
    
    Draw TempDC, Width, Height
    
    With uDesc
        .lSize = Len(uDesc)
        .lType = vbPicTypeBitmap
        .hBmp = hBmp
    End With
    
    aInput(0) = &H7BF80980
    aInput(1) = &H101ABF32
    aInput(2) = &HAA00BB8B
    aInput(3) = &HAB0C3000


    Call OleCreatePictureIndirect(uDesc, aInput(0), 1, Image)
    
    ReleaseDC 0&, lDC
    SelectObject TempDC, OldBmp
    DeleteDC TempDC
    
End Property

Public Sub GetCenterPie(X As Single, Y As Single)
    X = m_CenterCircle.X
    Y = m_CenterCircle.Y
End Sub

Public Property Get count() As Long
    count = ItemsCount
End Property

Public Property Let Special(Index As Long, Value As Boolean)
    m_Item(Index).Special = Value
    Me.Refresh
End Property

Public Property Get Special(Index As Long) As Boolean
    Special = m_Item(Index).Special
End Property

Public Property Let ItemColor(Index As Long, Value As OLE_COLOR)
    m_Item(Index).ItemColor = Value
    Me.Refresh
End Property

Public Property Get ItemColor(Index As Long) As OLE_COLOR
    ItemColor = m_Item(Index).ItemColor
End Property

Public Sub Clear()
    Dim i As Long
    For i = 0 To ItemsCount - 1
         GdipDeletePath m_Item(i).hPath
    Next
    ItemsCount = 0
    ReDim m_Item(0)
    Me.Refresh
End Sub
Public Function AddItem(ByVal ItemName As String, Value As Single, ItemColor As Long, Optional Special As Boolean)
    ReDim Preserve m_Item(ItemsCount)
    With m_Item(ItemsCount)
       .ItemName = ItemName
       .ItemColor = mycolor(ItemColor)
       .Value = Value
       .Special = Special
    End With
    ItemsCount = ItemsCount + 1
End Function

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal New_Value As String)
    m_Title = New_Value
    PropertyChanged "Title"
    Refresh
End Property

Public Property Get TitleFont() As StdFont
    Set TitleFont = m_TitleFont
End Property

Public Property Let TitleFont(ByVal New_Value As StdFont)
    Set m_TitleFont = New_Value
    PropertyChanged "TitleFont"
    Refresh
End Property

Public Property Set TitleFont(New_Font As StdFont)
    With m_TitleFont
        .Name = New_Font.Name
        .Size = New_Font.Size
        .bold = New_Font.bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .charset = New_Font.charset
    End With
    PropertyChanged "TitleFont"
    Refresh
End Property

Public Property Get TitleForeColor() As OLE_COLOR
    TitleForeColor = M2000color(m_TitleForeColor)
End Property

Public Property Let TitleForeColor(ByVal New_Value As OLE_COLOR)
    m_TitleForeColor = mycolor(New_Value)
    PropertyChanged "TitleForeColor"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = M2000color(m_BackColor)
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    m_BackColor = mycolor(New_Value)
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get BackColorOpacity() As Long
    BackColorOpacity = m_BackColorOpacity
End Property

Public Property Let BackColorOpacity(ByVal New_Value As Long)
    m_BackColorOpacity = New_Value
    PropertyChanged "BackColorOpacity"
    Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = M2000color(m_ForeColor)
End Property

Public Property Let ForeColor(ByVal New_Value As OLE_COLOR)
    m_ForeColor = mycolor(New_Value)
    PropertyChanged "ForeColor"
    Refresh
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Let Font(ByVal New_Value As StdFont)
    Set UserControl.Font = New_Value
    PropertyChanged "Font"
    Refresh
End Property

Public Property Set Font(New_Font As StdFont)
    With UserControl.Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .bold = New_Font.bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .charset = New_Font.charset
    End With
    PropertyChanged "Font"
    Refresh
End Property

Public Property Get FillOpacity() As Long
    FillOpacity = m_FillOpacity
End Property

Public Property Let FillOpacity(ByVal New_Value As Long)
    m_FillOpacity = New_Value
    PropertyChanged "FillOpacity"
    Refresh
End Property

Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Value As Boolean)
    m_Border = New_Value
    PropertyChanged "Border"
    Refresh
End Property

Public Property Get BorderRound() As Long
    BorderRound = m_BorderRound
End Property

Public Property Let BorderRound(ByVal New_Value As Long)
    m_BorderRound = New_Value
    PropertyChanged "BorderRound"
    Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = M2000color(m_BorderColor)
End Property

Public Property Let BorderColor(ByVal New_Value As OLE_COLOR)
    m_BorderColor = mycolor(New_Value)
    PropertyChanged "BorderColor"
    Refresh
End Property

Public Property Get LabelsFormats() As String
    LabelsFormats = m_LabelsFormats
End Property

Public Property Let LabelsFormats(ByVal New_Value As String)
    m_LabelsFormats = New_Value
    PropertyChanged "LabelsFormats"
    Refresh
End Property

Public Property Get FillGradient() As Boolean
    FillGradient = m_FillGradient
End Property

Public Property Let FillGradient(ByVal New_Value As Boolean)
    m_FillGradient = New_Value
    PropertyChanged "FillGradient"
    Refresh
End Property

Public Property Get ChartStyle() As ChartStyle
    ChartStyle = m_ChartStyle
End Property

Public Property Let ChartStyle(ByVal New_Value As ChartStyle)
    m_ChartStyle = New_Value
    PropertyChanged "ChartStyle"
    Refresh
End Property

Public Property Get LegendAlign() As ucPC_LegendAlign
    LegendAlign = m_LegendAlign
End Property

Public Property Let LegendAlign(ByVal New_Value As ucPC_LegendAlign)
    m_LegendAlign = New_Value
    PropertyChanged "LegendAlign"
    Refresh
End Property

Public Property Get LegendVisible() As Boolean
    LegendVisible = m_LegendVisible
End Property

Public Property Let LegendVisible(ByVal New_Value As Boolean)
    m_LegendVisible = New_Value
    PropertyChanged "LegendVisible"
    Refresh
End Property

Public Property Get DonutWidth() As Single
    DonutWidth = m_DonutWidth
End Property

Public Property Let DonutWidth(ByVal New_Value As Single)
    m_DonutWidth = New_Value
    PropertyChanged "DonutWidth"
    Refresh
End Property

Public Property Get SeparatorLine() As Boolean
    SeparatorLine = m_SeparatorLine
End Property

Public Property Let SeparatorLine(ByVal New_Value As Boolean)
    m_SeparatorLine = New_Value
    PropertyChanged "SeparatorLine"
    Refresh
End Property

Public Property Get SeparatorLineWidth() As Single
    SeparatorLineWidth = m_SeparatorLineWidth
End Property

Public Property Let SeparatorLineWidth(ByVal New_Value As Single)
    m_SeparatorLineWidth = New_Value
    PropertyChanged "SeparatorLineWidth"
    Refresh
End Property

Public Property Get SeparatorLineColor() As OLE_COLOR
    SeparatorLineColor = M2000color(m_SeparatorLineColor)
End Property

Public Property Let SeparatorLineColor(ByVal New_Value As OLE_COLOR)
    m_SeparatorLineColor = mycolor(New_Value)
    PropertyChanged "SeparatorLineColor"
    Refresh
End Property

Public Property Get LabelsPositions() As LabelsPositions
    LabelsPositions = m_LabelsPositions
End Property

Public Property Let LabelsPositions(ByVal New_Value As LabelsPositions)
    m_LabelsPositions = New_Value
    PropertyChanged "LabelsPositions"
    Refresh
End Property

Public Property Get LabelsVisible() As Boolean
    LabelsVisible = m_LabelsVisible
End Property

Public Property Let LabelsVisible(ByVal New_Value As Boolean)
    m_LabelsVisible = New_Value
    PropertyChanged "LabelsVisible"
    Refresh
End Property

Public Property Get Rotation() As Long
    Rotation = m_Rotation
End Property

Public Property Let Rotation(ByVal New_Value As Long)
    m_Rotation = New_Value Mod 360
    If m_Rotation < 0 Then m_Rotation = 360 + m_Rotation
    PropertyChanged "Rotation"
    Refresh
End Property




Private Sub tmrMOUSEOVER_Timer()
    Dim pt As POINTL
    Dim RECT As RectL
  
    GetCursorPos pt
    ScreenToClient c_lhWnd, pt
 
    With RECT
        .Left = m_PT.X - (m_Left - ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode))
        .Top = m_PT.Y - (m_Top - ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode))
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight
    End With

    If PtInRectL(RECT, pt.X, pt.Y) = 0 Then
        'mHotBar = -1
        HotItem = -1
        tmrMOUSEOVER.Interval = 0
        UserControl.Refresh
        'RaiseEvent MouseLeave
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    c_lhWnd = UserControl.ContainerHwnd

    With PropBag
        m_Title = .ReadProperty("Title", Ambient.DisplayName)
        m_BackColor = .ReadProperty("BackColor", vbWhite)
        m_BackColorOpacity = .ReadProperty("BackColorOpacity", 100)
        m_ForeColor = .ReadProperty("ForeColor", vbBlack)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        m_FillOpacity = .ReadProperty("FillOpacity", 100)
        m_Border = .ReadProperty("Border", False)
        m_BorderColor = .ReadProperty("BorderColor", &HF4F4F4)
        m_LinesCurve = .ReadProperty("LinesCurve", False)
        m_LinesWidth = .ReadProperty("LinesWidth", 1)
        m_VerticalLines = .ReadProperty("VerticalLines", False)
        m_FillGradient = .ReadProperty("FillGradient", False)
        m_LabelsVisible = .ReadProperty("LabelsVisible", False)
        m_HorizontalLines = .ReadProperty("HorizontalLines", True)
        m_ChartStyle = .ReadProperty("ChartStyle", CS_PIE)
        m_ChartOrientation = .ReadProperty("ChartOrientation", CO_Vertical)
        m_LegendAlign = .ReadProperty("LegendAlign", LA_RIGHT)
        m_LegendVisible = .ReadProperty("LegendVisible", True)
        Set m_TitleFont = .ReadProperty("TitleFont", Ambient.Font)
        m_TitleForeColor = .ReadProperty("TitleForeColor", Ambient.ForeColor)
        m_DonutWidth = .ReadProperty("DonutWidth", 50!)
        m_SeparatorLine = .ReadProperty("SeparatorLine", True)
        m_SeparatorLineWidth = .ReadProperty("SeparatorLineWidth", 2!)
        m_SeparatorLineColor = .ReadProperty("SeparatorLineColor", vbWhite)
        m_LabelsPositions = .ReadProperty("LabelsPositions", LP_Inside)
        m_LabelsVisible = .ReadProperty("LabelsVisible", m_LabelsVisible)
        m_LabelsFormats = .ReadProperty("LabelsFormats", "{P}%")
        m_Rotation = .ReadProperty("Rotation", 0)
        m_BorderRound = .ReadProperty("BorderRound", 0)
    End With
        
    If Not Ambient.UserMode Then Call Example
End Sub

Private Sub UserControl_Terminate()
    Dim i As Long
    For i = 0 To ItemsCount - 1
        GdipDeletePath m_Item(i).hPath
    Next
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Title", m_Title, Ambient.DisplayName
        .WriteProperty "BackColor", m_BackColor, vbWhite
        .WriteProperty "BackColorOpacity", m_BackColorOpacity, 100
        .WriteProperty "ForeColor", m_ForeColor, vbBlack
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "FillOpacity", m_FillOpacity, 100
        .WriteProperty "Border", m_Border, False
        .WriteProperty "BorderColor", m_BorderColor, &HF4F4F4
        .WriteProperty "LinesCurve", m_LinesCurve, False
        .WriteProperty "LinesWidth", m_LinesWidth, 1
        .WriteProperty "VerticalLines", m_VerticalLines, False
        .WriteProperty "FillGradient", m_FillGradient, False
        .WriteProperty "LabelsVisible", m_LabelsVisible, False
        .WriteProperty "HorizontalLines", m_HorizontalLines, True
        .WriteProperty "ChartStyle", m_ChartStyle, CS_PIE
        .WriteProperty "ChartOrientation", m_ChartOrientation, CO_Vertical
        .WriteProperty "LegendAlign", m_LegendAlign, LA_RIGHT
        .WriteProperty "LegendVisible", m_LegendVisible, True
        .WriteProperty "TitleFont", m_TitleFont, Ambient.Font
        .WriteProperty "TitleForeColor", m_TitleForeColor, Ambient.ForeColor
        .WriteProperty "DonutWidth", m_DonutWidth, 50!
        .WriteProperty "SeparatorLine", m_SeparatorLine, True
        .WriteProperty "SeparatorLineWidth", m_SeparatorLineWidth, 2!
        .WriteProperty "LabelsPositions", m_LabelsPositions, LP_Inside
        .WriteProperty "LabelsFormats", m_LabelsFormats, "{P}%"
        .WriteProperty "SeparatorLineColor", m_SeparatorLineColor, vbWhite
        .WriteProperty "Rotation", m_Rotation, 0
        .WriteProperty "BorderRound", m_BorderRound, 0
    End With
End Sub

Private Sub UserControl_InitProperties()
    m_Title = Ambient.DisplayName
    m_BackColor = vbWhite
    m_BackColorOpacity = 100
    m_ForeColor = vbBlack
    Set UserControl.Font = Ambient.Font
    m_FillOpacity = 100
    m_Border = False
    m_BorderColor = &HF4F4F4
    m_LinesWidth = 1
    m_VerticalLines = False
    m_FillGradient = False
    m_HorizontalLines = True
    m_ChartStyle = CS_PIE
    m_ChartOrientation = CO_Vertical
    m_LegendAlign = LA_RIGHT
    m_LegendVisible = True
    m_TitleFont.Name = UserControl.Font.Name
    m_TitleFont.Size = UserControl.Font.Size + 8
    m_DonutWidth = 50!
    m_SeparatorLine = True
    m_SeparatorLineWidth = 2!
    m_SeparatorLineColor = vbWhite
    m_LabelsVisible = True
    m_LabelsPositions = LP_Inside
    m_LabelsFormats = "{P}%"
    m_BorderRound = 0
    
    c_lhWnd = UserControl.ContainerHwnd

    
    If Not Ambient.UserMode Then Call Example
End Sub


Private Function GetTextSize(ByVal hGraphics As Long, ByVal text As String, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal bWordWrap As Boolean, ByRef SZ As SIZEF) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RECTF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long, hdc As Long
    Dim BB As RECTF, CF As Long, LF As Long
    
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
    End If
        
    If oFont.bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        

    hdc = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    ReleaseDC 0&, hdc

    layoutRect.Width = Width * nScale: layoutRect.Height = Height * nScale

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    
    GdipMeasureString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, BB, CF, LF
    
    SZ.Width = BB.Width
    SZ.Height = BB.Height
    
    GdipDeleteFont hFont
    GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily


End Function


Private Function DrawText(ByVal hGraphics As Long, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal ForeColor As Long, Optional HAlign As CaptionAlignmentH, Optional VAlign As CaptionAlignmentV, Optional bWordWrap As Boolean) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RECTF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim hdc As Long

  
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) <> GDIP_OK Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) <> GDIP_OK Then Exit Function
        'If GdipGetGenericFontFamilySerif(hFontFamily) Then Exit Function
    End If

    If GdipCreateStringFormat(0, 0, hFormat) = GDIP_OK Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        'GdipSetStringFormatFlags hFormat, HotkeyPrefixShow
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
    End If
        
    If oFont.bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        

    hdc = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    ReleaseDC 0&, hdc

    layoutRect.Left = X: layoutRect.Top = Y
    layoutRect.Width = Width: layoutRect.Height = Height

    If GdipCreateSolidFill(ForeColor, hBrush) = GDIP_OK Then
        If GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont) = GDIP_OK Then
            GdipDrawString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, hBrush
            GdipDeleteFont hFont
        End If
        GdipDeleteBrush hBrush
    End If
    
    If hFormat Then GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily


End Function

Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If UserControl.Enabled Then
        HitResult = vbHitResultHit
        If Ambient.UserMode Then
            Dim pt As POINTL

            If tmrMOUSEOVER.Interval = 0 Then
                GetCursorPos pt
                ScreenToClient c_lhWnd, pt
                m_PT.X = pt.X - X
                m_PT.Y = pt.Y - Y
    
                m_Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
                m_Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)
 
              
                tmrMOUSEOVER.Interval = 1
                'RaiseEvent MouseEnter
            End If

        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    HotItem = -1
    nScale = GetWindowsDPI
    Set m_TitleFont = New StdFont
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
    Dim i As Long
    Dim lResult As Long
    
    If Button <> vbLeftButton Then Exit Sub
    
    For i = 0 To ItemsCount - 1

        lResult = 0
        Call GdipIsVisiblePathPoint(m_Item(i).hPath, X, Y, 0&, lResult)
    
        If lResult Then
            If i = HotItem Then
                 RaiseEvent ItemClick(i)
            End If
            Exit Sub
        End If
    
    Next
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim lResult As Long
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button <> 0 Then Exit Sub
    For i = 0 To ItemsCount - 1

        If PtInRectL(m_Item(i).LegendRect, X, Y) Then
            If i <> HotItem Then
                HotItem = i
                Me.Refresh
            End If
            Exit Sub
        End If
    
        lResult = 0
        Call GdipIsVisiblePathPoint(m_Item(i).hPath, X, Y, 0&, lResult)
    
        If lResult Then
            If i <> HotItem Then
                HotItem = i
                Me.Refresh
            End If
            Exit Sub
        End If
    
    Next
    
    If HotItem <> -1 Then
        HotItem = -1
        Me.Refresh
    End If

End Sub

Private Function PtInRectL(RECT As RectL, ByVal X As Single, ByVal Y As Single) As Boolean
    With RECT
        If X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height Then
            PtInRectL = True
        End If
    End With
End Function
'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyDown(KeyCode, Shift)
'End Sub
'
'Private Sub UserControl_KeyPress(KeyAscii As Integer)
'     RaiseEvent KeyPress(KeyAscii)
'End Sub
'
'Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyUp(KeyCode, Shift)
'End Sub
'
'Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
'End Sub
'
'Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
'End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Paint()
    Draw UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub


Private Function SafeRange(Value, Min, Max)
    
    If Value < Min Then
        SafeRange = Min
    ElseIf Value > Max Then
        SafeRange = Max
    Else
        SafeRange = Value
    End If
End Function


Public Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
    'By LaVople
    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to RGBA format
    ' Passing VB system color constants is allowed, i.e., vbButtonFace
    ' Pass Opacity as a value from 0 to 255

    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
    
End Function



'*1
Private Sub Draw(hdc As Long, ScaleWidth As Long, ScaleHeight As Long)
    Dim hGraphics As Long
    Dim hBrush As Long, hPen As Long
    Dim i As Single, j As Long
    Dim mHeight As Single
    Dim mWidth As Single
    Dim mPenWidth As Single
    Dim MarginLeft As Single
    Dim MarginRight As Single
    Dim TopHeader As Single
    Dim Footer As Single
    Dim TextWidth As Single
    Dim TextHeight As Single
    Dim XX As Single, YY As Single
    Dim lForeColor As Long
    Dim RectL As RectL
    Dim LabelsRect As RectL
    Dim PT16 As Single
    Dim ColRow As Integer
    Dim TitleSize As SIZEF
    Dim sDisplay As String
    Dim SafePercent As Single
    Dim Min As Single, LastAngle As Single, Angle As Single, Total  As Single
    Dim DonutSize As Single
    Dim R1 As Single, R2 As Single, R3 As Single
    Dim cx As Single, cy   As Single
    Dim Left As Single, Top As Single
    Dim Percent As Single
    Const PItoRAD = 3.141592 / 180
    Dim lTop As Single
    Dim sLabelText As String
    Dim bAngMaj180 As Boolean
    Dim LblWidth As Single
    Dim LblHeight As Single
    Dim mFormat As String
    Dim A As Single
    Dim Displacement As Single
    Dim lColor As Long
    
    If GdipCreateFromHDC(hdc, hGraphics) Then Exit Sub
  
    Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)

    PT16 = 16 * nScale
    mPenWidth = 1 * nScale
    DonutSize = m_DonutWidth * nScale

    MarginLeft = PT16
    TopHeader = PT16
    MarginRight = PT16
    Footer = PT16

    If m_LegendVisible Then
        For i = 0 To ItemsCount - 1
            m_Item(i).TextHeight = UserControl.TextHeight(m_Item(i).ItemName) * 1.5
            m_Item(i).TextWidth = UserControl.TextWidth(m_Item(i).ItemName) * 1.5 + m_Item(i).TextHeight
        Next
    End If

    If Len(m_Title) Then
        Call GetTextSize(hGraphics, m_Title, ScaleWidth, 0, m_TitleFont, True, TitleSize)
        TopHeader = TopHeader + TitleSize.Height
    End If

    mWidth = ScaleWidth - MarginLeft - MarginRight
    mHeight = ScaleHeight - TopHeader - Footer

    'Calculate the Legend Area
    If m_LegendVisible Then
        ColRow = 1
        Select Case m_LegendAlign
            Case LA_RIGHT, LA_LEFT
                With LabelsRect
                    TextWidth = 0
                    TextHeight = 0
                    For i = 0 To ItemsCount - 1
                        If TextHeight + m_Item(i).TextHeight > mHeight Then
                            .Width = .Width + TextWidth
                            ColRow = ColRow + 1
                            TextWidth = 0
                            TextHeight = 0
                        End If
    
                        TextHeight = TextHeight + m_Item(i).TextHeight
                        .Height = .Height + m_Item(i).TextHeight
    
                        If TextWidth < m_Item(i).TextWidth Then
                            TextWidth = m_Item(i).TextWidth '+ PT16
                        End If
                    Next
                    .Width = .Width + TextWidth
                    If m_LegendAlign = LA_LEFT Then
                        MarginLeft = MarginLeft + .Width
                    Else
                        MarginRight = MarginRight + .Width
                    End If
                    mWidth = mWidth - .Width
                End With
    
            Case LA_BOTTOM, LA_TOP
                With LabelsRect
             
                    .Height = m_Item(0).TextHeight + PT16 / 2
                    TextWidth = 0
                    For i = 0 To ItemsCount - 1
                        If TextWidth + m_Item(i).TextWidth > mWidth Then
                            .Height = .Height + m_Item(i).TextHeight
                            ColRow = ColRow + 1
                            TextWidth = 0
                        End If
                        TextWidth = TextWidth + m_Item(i).TextWidth
                        .Width = .Width + m_Item(i).TextWidth
                    Next
                    If m_LegendAlign = LA_TOP Then
                        TopHeader = TopHeader + .Height
                    End If
                    mHeight = mHeight - .Height
                End With
        End Select
    End If
    
    
    Dim RECTF As RECTF
    With RECTF
        .Width = ScaleWidth - 1 * nScale
        .Height = ScaleHeight - 1 * nScale
    End With
    
    RoundRect hGraphics, RECTF, RGBtoARGB(m_BackColor, m_BackColorOpacity), RGBtoARGB(m_BorderColor, 100), m_BorderRound * nScale, m_Border


'    'Background
'    If m_BackColorOpacity > 0 Then
'        GdipCreateSolidFill RGBtoARGB(m_BackColor, m_BackColorOpacity), hBrush
'        GdipFillRectangleI hGraphics, hBrush, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
'        GdipDeleteBrush hBrush
'    End If
'
'    'Border
'    If m_Border Then
'        Call GdipCreatePen1(RGBtoARGB(m_BorderColor, 50), mPenWidth, &H2, hPen)
'        GdipDrawRectangleI hGraphics, hPen, mPenWidth / 2, mPenWidth / 2, UserControl.ScaleWidth - mPenWidth, UserControl.ScaleHeight - mPenWidth
'        GdipDeletePen hPen
'    End If
'

    
    'Sum of itemes
    For i = 0 To ItemsCount - 1
        Total = Total + m_Item(i).Value
    Next
    
    'calculate max size of labels
    For i = 0 To ItemsCount - 1
        With m_Item(i)
            Percent = Round(100 * .Value / Total, 1)
            If i < ItemsCount - 1 Then
                SafePercent = SafePercent + Percent
            Else
                Percent = Round(100 - SafePercent, 1)
            End If
            .text = Replace(m_LabelsFormats, "{A}", .ItemName)
            .text = Replace(.text, "{P}", Percent)
            .text = Replace(.text, "{V}", Round(.Value, 1))
            .text = Replace(.text, "{LF}", vbLf)
            
            TextWidth = UserControl.TextWidth(.text) * 1.3
            TextHeight = UserControl.TextHeight(.text) * 1.3
            If TextWidth > LblWidth Then LblWidth = TextWidth
            If TextHeight > LblHeight Then LblHeight = TextHeight
        End With
    Next
    
    'size of pie
    If m_LabelsPositions = LP_Outside Or m_LabelsPositions = LP_TwoColumns Then
        Min = IIf(mWidth - LblWidth * 2 < mHeight - LblHeight * 2, mWidth - LblWidth * 2, mHeight - LblHeight * 2)
    Else
        Min = IIf(mWidth < mHeight, mWidth, mHeight)
    End If
    

    If Min / 3 < DonutSize Then DonutSize = Min / 3
    XX = MarginLeft + mWidth / 2 - Min / 2
    YY = TopHeader + mHeight / 2 - Min / 2
    m_CenterCircle.X = MarginLeft + mWidth / 2
    m_CenterCircle.Y = TopHeader + mHeight / 2
    R1 = Min / 2
    
'    If m_SeparatorLine Then
'        GdipCreateSolidFill RGBtoARGB(m_SeparatorLineColor, m_BackColorOpacity), hBrush
'        GdipFillEllipseI hGraphics, hBrush, XX - m_SeparatorLineWidth, YY - m_SeparatorLineWidth, Min + m_SeparatorLineWidth * 2, Min + m_SeparatorLineWidth * 2
'        GdipDeleteBrush hBrush
'    End If
    
    LastAngle = m_Rotation - 90
    For i = 0 To ItemsCount - 1
        Angle = 360 * m_Item(i).Value / Total


 '*1
        If m_Item(i).Special Then
            R2 = PT16 / 1.5
            Left = XX + (R2 * Cos((LastAngle + Angle / 2) * PItoRAD))
            Top = YY + (R2 * Sin((LastAngle + Angle / 2) * PItoRAD))
        Else
            Left = XX
            Top = YY
        End If
        
        If m_Item(i).hPath <> 0 Then GdipDeletePath m_Item(i).hPath
        GdipCreatePath 0, m_Item(i).hPath
        
        If m_ChartStyle = CS_DONUT Then
            GdipAddPathArc m_Item(i).hPath, Left, Top, Min, Min, LastAngle, Angle
            GdipAddPathArc m_Item(i).hPath, Left + DonutSize, Top + DonutSize, Min - DonutSize * 2, Min - DonutSize * 2, LastAngle + Angle, -Angle
        Else
            GdipAddPathPie m_Item(i).hPath, Left, Top, Min, Min, LastAngle, Angle
        End If

        If HotItem = i Then
            lColor = RGBtoARGB(ShiftColor(m_Item(i).ItemColor, vbWhite, 150), m_FillOpacity)
        Else
            lColor = RGBtoARGB(m_Item(i).ItemColor, m_FillOpacity)
        End If
        If m_FillGradient Then
            With RectL
                .Left = MarginLeft - R2
                .Top = TopHeader - R2
                .Width = mWidth + R2 * 2
                .Height = mHeight + R2 * 2
            End With
            GdipCreateLineBrushFromRectWithAngleI RectL, lColor, RGBtoARGB(vbWhite, 100), 180 + LastAngle + Angle / 2, 0, WrapModeTile, hBrush
        Else
            GdipCreateSolidFill lColor, hBrush
        End If
        GdipFillPath hGraphics, hBrush, m_Item(i).hPath
        GdipDeleteBrush hBrush
        
        R1 = Min / 2
        R2 = m_Item(i).TextWidth / 2
        R3 = m_Item(i).TextHeight / 2
    
        cx = XX + Min / 2 + TextWidth
        cy = YY + Min / 2 + TextHeight
        
        Left = cx + ((R1 - R2) * Cos((LastAngle + Angle / 2) * PItoRAD)) - R2
        Top = cy + ((R1 - R3) * Sin((LastAngle + Angle / 2) * PItoRAD)) - R3
        'DrawText hGraphics, m_Item(i).ItemName, Left, Top, R2 * 2, R3 * 2, UserControl.Font, lForeColor, cCenter, cMiddle
        LastAngle = LastAngle + Angle '+ 2
    Next
        
        
'*2

   

    LastAngle = m_Rotation - 90
    bAngMaj180 = False
    For i = 0 To ItemsCount - 1
        Angle = 360 * m_Item(i).Value / Total

        If m_SeparatorLine Then
            GdipCreatePen1 RGBtoARGB(m_SeparatorLineColor, 100), m_SeparatorLineWidth * nScale, &H2, hPen
            GdipSetPenEndCap hPen, &H2
            
            R1 = (Min + mPenWidth / 2) / 2
            R2 = (Min - mPenWidth / 2) / 2 - DonutSize
    
            cx = XX + Min / 2
            cy = YY + Min / 2
            
            Left = cx + (R1 * Cos((LastAngle) * PItoRAD))
            Top = cy + (R1 * Sin((LastAngle) * PItoRAD))
            
            If m_ChartStyle = CS_DONUT Then
                cx = cx + (R2 * Cos((LastAngle) * PItoRAD))
                cy = cy + (R2 * Sin((LastAngle) * PItoRAD))
            Else
                'GdipDrawEllipseI hGraphics, hPen, XX, YY, Min, Min
            End If
            
            GdipDrawLineI hGraphics, hPen, Left, Top, cx, cy
            
            GdipDeletePen hPen
        End If

        TextWidth = LblWidth
        TextHeight = LblHeight
        
        If m_LabelsPositions = LP_Inside Then
            If DonutSize > TextWidth Then TextWidth = DonutSize
            If DonutSize > TextHeight Then TextHeight = DonutSize
        End If
        
        R2 = TextWidth / 2
        R3 = TextHeight / 2
        Displacement = IIf(m_Item(i).Special, PT16 / 1.5, 0)
        
        cx = XX + Min / 2
        cy = YY + Min / 2
        
        A = LastAngle + Angle / 2
        
        If m_LabelsPositions = LP_Inside Then
            Left = cx + ((R1 - R2 + Displacement) * Cos(A * PItoRAD)) - R2
            Top = cy + ((R1 - R3 + Displacement) * Sin(A * PItoRAD)) - R3
        Else
            Left = cx + ((R1 + R2 + Displacement) * Cos(A * PItoRAD)) - R2
            Top = cy + ((R1 + R3 + Displacement) * Sin(A * PItoRAD)) - R3
        End If
        If m_LabelsVisible Then
            If m_LabelsPositions = LP_TwoColumns Then
                Dim LineOut As Integer
                LineOut = UserControl.TextHeight("Aj") / 2
                GdipCreateSolidFill RGBtoARGB(m_Item(i).ItemColor, 50), hBrush
                GdipCreatePen1 RGBtoARGB(m_Item(i).ItemColor, 100), 1 * nScale, &H2, hPen
    
                If (LastAngle + Angle / 2 + 90) Mod 359 < 180 Then
                    If bAngMaj180 Then
                        bAngMaj180 = False
                        lTop = Top
                    End If
                    
                    If lTop <= 0 Then lTop = Top
    
                    If Top < lTop Then
                        lTop = lTop
                    Else
                        lTop = Top
                    End If
    
                    Left = XX + Min + PT16
               
                    GdipFillRectangleI hGraphics, hBrush, Left, lTop, TextWidth, TextHeight
                    DrawText hGraphics, m_Item(i).text, Left, lTop, TextWidth, TextHeight, UserControl.Font, RGBtoARGB(m_ForeColor, 100), cCenter, cMiddle
                    lTop = lTop + TextHeight
                    
                    Left = cx + (R1 * Cos(A * PItoRAD))
                    Top = cy + (R1 * Sin(A * PItoRAD))
                    cx = cx + ((R1 + LineOut) * Cos(A * PItoRAD))
                    cy = cy + ((R1 + LineOut) * Sin(A * PItoRAD))
                    
                    GdipDrawLineI hGraphics, hPen, Left, Top, cx, cy
                    Left = XX + Min + PT16
                    Top = lTop - TextHeight / 2
                    GdipDrawLineI hGraphics, hPen, cx, cy, Left, Top
                Else
                    If bAngMaj180 = False Then
                        bAngMaj180 = True
                        lTop = TopHeader + mHeight
                    End If
                    
                    If lTop <= 0 Then lTop = Top
                    
                    If Top > lTop Then
                        lTop = lTop
                    Else
                        lTop = Top
                    End If
                    
                    Left = XX - TextWidth - PT16
                    GdipFillRectangleI hGraphics, hBrush, Left, lTop, TextWidth, TextHeight
                    DrawText hGraphics, m_Item(i).text, Left, lTop, TextWidth, TextHeight, UserControl.Font, RGBtoARGB(m_ForeColor, 100), cCenter, cMiddle
                    Left = cx + (R1 * Cos(A * PItoRAD))
                    Top = cy + (R1 * Sin(A * PItoRAD))
                    cx = cx + ((R1 + LineOut) * Cos(A * PItoRAD))
                    cy = cy + ((R1 + LineOut) * Sin(A * PItoRAD))
                    GdipDrawLineI hGraphics, hPen, Left, Top, cx, cy
                    Left = XX - PT16
                    Top = lTop + TextHeight / 2
                    GdipDrawLineI hGraphics, hPen, cx, cy, Left, Top
                    lTop = lTop - TextHeight
                End If
                GdipDeleteBrush hBrush
                GdipDeletePen hPen
                
            ElseIf m_LabelsPositions = LP_Inside Then
                'lForeColor = IIf(IsDarkColor(m_Item(i).ItemColor), &H808080, vbWhite)
                'DrawText hGraphics, m_Item(i).Text, Left + 1, Top + 1, TextWidth, TextHeight, UserControl.Font, RGBtoARGB(lForeColor, 100), cCenter, cMiddle
                If HotItem = i Then
                    lColor = ShiftColor(m_Item(i).ItemColor, vbWhite, 150)
                Else
                     lColor = m_Item(i).ItemColor
                End If
                lForeColor = IIf(IsDarkColor(lColor), vbWhite, vbBlack)
                DrawText hGraphics, m_Item(i).text, Left, Top, TextWidth, TextHeight, UserControl.Font, RGBtoARGB(lForeColor, 100), cCenter, cMiddle
            Else
                DrawText hGraphics, m_Item(i).text, Left, Top, TextWidth, TextHeight, UserControl.Font, RGBtoARGB(m_ForeColor, 100), cCenter, cMiddle
            End If
        End If
        LastAngle = LastAngle + Angle '+ 2
    Next
               

    
    If m_LegendVisible Then
        For i = 0 To ItemsCount - 1
            lForeColor = RGBtoARGB(m_ForeColor, 100)
            Select Case m_LegendAlign
                Case LA_RIGHT, LA_LEFT
                    With LabelsRect
                        TextWidth = 0
                        
                        If .Left = 0 Then
                            TextHeight = 0
                            If m_LegendAlign = LA_LEFT Then
                                .Left = PT16
                            Else
                                .Left = MarginLeft + mWidth + PT16
                            End If
                            If ColRow = 1 Then
                                .Top = TopHeader + mHeight / 2 - .Height / 2
                            Else
                                .Top = TopHeader
                            End If
                        End If
                        
                        If TextWidth < m_Item(i).TextWidth Then
                            TextWidth = m_Item(i).TextWidth '+ PT16
                        End If
        
                        If TextHeight + m_Item(i).TextHeight > mHeight Then
                             If i > 0 Then .Left = .Left + TextWidth
                            .Top = TopHeader
                             TextHeight = 0
                        End If
                        m_Item(i).LegendRect.Left = .Left
                        m_Item(i).LegendRect.Top = .Top
                        m_Item(i).LegendRect.Width = m_Item(i).TextWidth
                        m_Item(i).LegendRect.Height = m_Item(i).TextHeight
                        
                        With m_Item(i).LegendRect
                            GdipCreateSolidFill RGBtoARGB(m_Item(i).ItemColor, 100), hBrush
                            GdipFillEllipseI hGraphics, hBrush, .Left, .Top + m_Item(i).TextHeight / 4, m_Item(i).TextHeight / 2, m_Item(i).TextHeight / 2
                            GdipDeleteBrush hBrush
                        End With
                        DrawText hGraphics, m_Item(i).ItemName, .Left + m_Item(i).TextHeight / 1.5, .Top, m_Item(i).TextWidth, m_Item(i).TextHeight, UserControl.Font, lForeColor, cLeft, cMiddle
                        TextHeight = TextHeight + m_Item(i).TextHeight
                        .Top = .Top + m_Item(i).TextHeight
                        
                    End With
                    
                Case LA_BOTTOM, LA_TOP
                    With LabelsRect
                        If .Left = 0 Then
                            If ColRow = 1 Then
                                .Left = MarginLeft + mWidth / 2 - .Width / 2
                            Else
                                .Left = MarginLeft
                            End If
                            If m_LegendAlign = LA_TOP Then
                                .Top = PT16 + TitleSize.Height
                            Else
                                .Top = TopHeader + mHeight + TitleSize.Height - PT16 / 2
                            End If
                        End If
        
                        If .Left + m_Item(i).TextWidth - MarginLeft > mWidth Then
                            .Left = MarginLeft
                            .Top = .Top + m_Item(i).TextHeight
                        End If
        
                        GdipCreateSolidFill RGBtoARGB(m_Item(i).ItemColor, 100), hBrush
                        GdipFillEllipseI hGraphics, hBrush, .Left, .Top + m_Item(i).TextHeight / 4, m_Item(i).TextHeight / 2, m_Item(i).TextHeight / 2
                        GdipDeleteBrush hBrush
                        m_Item(i).LegendRect.Left = .Left
                        m_Item(i).LegendRect.Top = .Top
                        m_Item(i).LegendRect.Width = m_Item(i).TextWidth
                        m_Item(i).LegendRect.Height = m_Item(i).TextHeight
                        
                        DrawText hGraphics, m_Item(i).ItemName, .Left + m_Item(i).TextHeight / 1.5, .Top, m_Item(i).TextWidth, m_Item(i).TextHeight, UserControl.Font, lForeColor, cLeft, cMiddle
                        .Left = .Left + m_Item(i).TextWidth '+ M_ITEM(i).TextHeight / 1.5
                    End With
            End Select
        

        Next
    End If
  

    'Title
    If Len(m_Title) Then
        DrawText hGraphics, m_Title, 0, PT16 / 2, ScaleWidth, TopHeader, m_TitleFont, RGBtoARGB(m_TitleForeColor, 100), cCenter, cTop, True
    End If
    
    ShowToolTips hGraphics

    Call GdipDeleteGraphics(hGraphics)
    

End Sub


Private Sub ShowToolTips(hGraphics As Long)
    Dim i As Long, j As Long
    Dim sDisplay As String
    Dim bBold As Boolean
    Dim RECTF As RECTF
    Dim LW As Single
    Dim lForeColor As Long
    Dim sText As String
    Dim TM As Single
    Dim pt As POINTF
    Dim SZ As SIZEF
    
    If HotItem > -1 Then

        lForeColor = RGBtoARGB(m_ForeColor, 100)
        LW = m_LinesWidth * nScale
        TM = UserControl.TextHeight("Aj") / 4
    
        sText = m_Item(HotItem).ItemName & ": " & m_Item(HotItem).text
        GetTextSize hGraphics, sText, 0, 0, UserControl.Font, False, SZ

        
        With RECTF
            GdipGetPathLastPoint m_Item(HotItem).hPath, pt
            .Left = pt.X
            .Top = pt.Y
            .Width = SZ.Width + TM * 2
            .Height = SZ.Height + TM * 2
            
            If .Left < 0 Then .Left = LW
            If .Top < 0 Then .Top = LW
            If .Left + .Width >= UserControl.ScaleWidth - LW Then .Left = UserControl.ScaleWidth - .Width - LW
            If .Top + .Height >= UserControl.ScaleHeight - LW Then .Top = UserControl.ScaleHeight - .Height - LW
        End With
                    
        RoundRect hGraphics, RECTF, RGBtoARGB(m_BackColor, 90), RGBtoARGB(m_Item(HotItem).ItemColor, 80), TM


        With RECTF
            .Left = .Left + TM
            .Top = .Top
            DrawText hGraphics, m_Item(HotItem).ItemName & ": ", .Left, .Top, .Width, .Height, UserControl.Font, lForeColor, cLeft, cMiddle
            GetTextSize hGraphics, m_Item(HotItem).ItemName & ": ", 0, 0, UserControl.Font, False, SZ
            bBold = UserControl.Font.bold
            UserControl.Font.bold = True
            DrawText hGraphics, m_Item(HotItem).text, .Left + SZ.Width, .Top, .Width, .Height, UserControl.Font, lForeColor, cLeft, cMiddle
            UserControl.Font.bold = bBold
        End With

    End If
End Sub

Private Sub RoundRect(ByVal hGraphics As Long, RECT As RECTF, ByVal BackColor As Long, ByVal BorderColor As Long, ByVal Round As Single, Optional bBorder As Boolean = True)
    Dim hPen As Long, hBrush As Long
    Dim mPath As Long
    
    GdipCreateSolidFill BackColor, hBrush
    If bBorder Then GdipCreatePen1 BorderColor, 1 * nScale, &H2, hPen

    If Round = 0 Then
        With RECT
            GdipFillRectangleI hGraphics, hBrush, .Left, .Top, .Width, .Height
            If hPen Then GdipDrawRectangleI hGraphics, hPen, .Left, .Top, .Width, .Height
        End With
    Else
        If GdipCreatePath(&H0, mPath) = 0 Then
            Round = Round * 2
            With RECT
                GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
                GdipAddPathArcI mPath, .Left + .Width - Round, .Top, Round, Round, 270, 90
                GdipAddPathArcI mPath, .Left + .Width - Round, .Top + .Height - Round, Round, Round, 0, 90
                GdipAddPathArcI mPath, .Left, .Top + .Height - Round, Round, Round, 90, 90
                GdipClosePathFigure mPath
            End With
            GdipFillPath hGraphics, hBrush, mPath
            If hPen Then GdipDrawPath hGraphics, hPen, mPath
            Call GdipDeletePath(mPath)
        End If
    End If
        
    Call GdipDeleteBrush(hBrush)
    If hPen Then Call GdipDeletePen(hPen)
    
End Sub


Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
  
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function

Private Function IsDarkColor(ByVal color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor color, 0, VarPtr(color)
    CopyMemory BGRA(0), color, 4&
  
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382

End Function

Private Sub Example()
    Me.AddItem "Juan", 70, vbRed
    Me.AddItem "Adan", 20, vbGreen
    Me.AddItem "Pedro", 10, vbBlue
End Sub

