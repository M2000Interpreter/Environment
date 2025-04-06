VERSION 5.00
Begin VB.UserControl ucChartBar 
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
Attribute VB_Name = "ucChartBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Module Name: ucChartBar
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 08/08/2020
'Version: 1.0.0
'-----------------------------------------------
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTL) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GdipCreatePen2 Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipFlattenPath Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mMatrix As Long, ByVal mFlatness As Single) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipGetPointCount Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mCount As Long) As Long
Private Declare Function GdipRotatePathGradientTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mAngle As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTranslatePathGradientTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mDx As Single, ByVal mDy As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTranslateTextureTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mDx As Single, ByVal mDy As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef graphics As Long) As Long
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
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function OleTranslateColor Lib "OleAut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
  
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
    top As Single
    Width As Single
    Height As Single
End Type

Private Type RectL
    Left As Long
    top As Long
    Width As Long
    Height As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Public Enum ucCB_TextAlignmentH
    cLeft
    cCenter
    cRight
End Enum

Public Enum brLabelsPositions
    LP_TOP
    LP_CENTER
    LP_BOTTOM
    LP_ABOBE
End Enum

Private Enum TextAlignmentV
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

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
'Public Event PrePaint(hdc As Long)
'Public Event PostPaint(ByVal hdc As Long)
'Public Event KeyPress(KeyAscii As Integer)
'Public Event KeyUp(KeyCode As Integer, Shift As Integer)
'Public Event KeyDown(KeyCode As Integer, Shift As Integer)
'Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

Public Enum ChartStyle
    CS_GroupedColumn
    CS_StackedBars
    CS_StackedBarsPercent
End Enum

Public Enum ChartBarOrientation
    CO_Vertical
    CO_Horizontal
End Enum

Public Enum ucCB_LegendAlign
    LA_LEFT
    LA_TOP
    LA_RIGHT
    LA_BOTTOM
End Enum

Dim nScale As Single

Dim m_Title As String
Dim m_BackColor As OLE_COLOR
Dim m_BackColorOpacity As Long
Dim m_ForeColor As OLE_COLOR
Dim m_LinesColor As OLE_COLOR
Dim m_FillOpacity As Long
Dim m_Border As Boolean
Dim m_LinesCurve As Boolean
Dim m_LinesWidth As Long
Dim m_VerticalLines As Boolean
Dim m_FillGradient As Boolean
Dim m_LabelsVisible As Boolean
Dim m_HorizontalLines As Boolean
Dim m_ChartStyle As ChartStyle
Dim m_ChartBarOrientation As ChartBarOrientation
Dim m_LegendAlign As ucCB_LegendAlign
Dim m_LegendVisible As Boolean
Dim m_AxisXVisible As Boolean
Dim m_AxisYVisible As Boolean
Dim m_WordWrap As Boolean
Dim m_TitleFont As StdFont
Dim m_TitleForeColor As OLE_COLOR
Dim m_AxisMax As Single
Dim m_AxisMin As Single
Dim m_LabelsPositions As brLabelsPositions
Dim m_AxisAngle As Single
Dim m_AxisAlign As ucCB_TextAlignmentH
Dim m_LabelsFormats As String
Dim m_BorderRound As Long
Dim m_BorderColor As OLE_COLOR

Private Type tSerie
    SerieName As String
    TextWidth As Long
    TextHeight As Long
    SeireColor As Long
    Values As Collection
    Rects() As RectL
    LegendRect As RectL
    CustomColors As Collection
End Type

Dim cAxisItem As Collection
Dim m_Serie() As tSerie
Dim SerieCount As Long
Dim mHotSerie As Long
Dim mHotBar As Long
Dim m_PT As POINTL
Dim m_Left As Long
Dim m_Top As Long

'*-
Public Sub Clear()
    mHotBar = -1
    mHotSerie = -1
    Erase m_Serie
    SerieCount = 0
    Me.Refresh
End Sub


Private Sub tmrMOUSEOVER_Timer()
    Dim pt As POINTL
    Dim RECT As RectL
  
    GetCursorPos pt
    ScreenToClient UserControl.ContainerHwnd, pt
 
    With RECT
        .Left = m_PT.X - (m_Left - ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode))
        .top = m_PT.Y - (m_Top - ScaleY(Extender.top, vbContainerSize, UserControl.ScaleMode))
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight
    End With

    If PtInRectL(RECT, pt.X, pt.Y) = 0 Then
        mHotBar = -1
        mHotSerie = -1
        tmrMOUSEOVER.Interval = 0
        UserControl.Refresh
        'RaiseEvent MouseLeave
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrExit
    Dim i As Long, j As Long
    For i = 0 To SerieCount - 1
        With m_Serie(i)
            If PtInRectL(.LegendRect, X, Y) Then
                If i <> mHotSerie Then
                    mHotSerie = i
                    mHotBar = -1
                    Me.Refresh
                End If
                Exit Sub
            End If
            For j = 1 To .Values.Count
                If PtInRectL(.Rects(j - 1), X, Y) Then
                    If i <> mHotSerie Then
                        mHotSerie = i
                        mHotBar = j - 1
                        Me.Refresh
                    End If
                    Exit Sub
                End If
            Next
        End With
    Next
    
    If mHotSerie <> -1 Then
        mHotSerie = -1
        Me.Refresh
    End If
    
ErrExit:
End Sub

Private Function PtInRectL(RECT As RectL, ByVal X As Single, ByVal Y As Single) As Boolean
    With RECT
        If X >= .Left And X <= .Left + .Width And Y >= .top And Y <= .top + .Height Then
            PtInRectL = True
        End If
    End With
End Function

Public Function UpdateSerie(ByVal Index As Integer, ByVal SerieName As String, ByVal SerieColor As Long, Values As Collection)
    Dim TempCol As Collection
    Dim i As Long, j As Long
    Dim dif As Single
    Dim bCancel As Boolean
    Dim bVisible As Boolean
    bCancel = True
    bVisible = Me.LabelsVisible
    Me.LabelsVisible = False
    With m_Serie(Index)
        .SerieName = SerieName
        .SeireColor = SerieColor
        Set TempCol = .Values
        
        For i = 1 To 10
            Set .Values = New Collection
            For j = 1 To Values.Count
                dif = Values(j) - TempCol(j)
                .Values.Add Round(TempCol(j) + i * dif / 10)
            Next
            Me.Refresh
            DoEvents
            Wait 1
            '
        Next
        Set .Values = Values
        Me.LabelsVisible = bVisible
    End With
End Function


Private Function Wait(Interval As Integer)
    Dim t As Single
    t = Timer + Interval / 100
    Do While t > Timer
        'DoEvents
    Loop
End Function

Public Function AddLineSeries(ByVal SerieName As String, Values As Collection, ByVal SerieColor As Long, Optional cCustomColors As Collection)
    ReDim Preserve m_Serie(SerieCount)
    With m_Serie(SerieCount)
       .SerieName = SerieName
       .SeireColor = mycolor(SerieColor)
       Set .Values = Values
       Set .CustomColors = cCustomColors
    End With
    SerieCount = SerieCount + 1
End Function

Public Function AddAxisItems(AxisItems As Collection, Optional ByVal WordWrap As Boolean, Optional AxisAngle As Single, Optional AxisAlign As ucCB_TextAlignmentH = cCenter)
    Set cAxisItem = AxisItems
    m_WordWrap = WordWrap
    m_AxisAngle = AxisAngle
    m_AxisAlign = AxisAlign
End Function

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal New_Value As String)
    m_Title = New_Value
    PropertyChanged "Title"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
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
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_Value As OLE_COLOR)
    m_ForeColor = mycolor(New_Value)
    PropertyChanged "ForeColor"
    Refresh
End Property

Public Property Get LinesColor() As OLE_COLOR
    LinesColor = m_LinesColor
End Property

Public Property Let LinesColor(ByVal New_Value As OLE_COLOR)
    m_LinesColor = mycolor(New_Value)
    PropertyChanged "LinesColor"
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

Public Property Get LinesCurve() As Boolean
    LinesCurve = m_LinesCurve
End Property

Public Property Let LinesCurve(ByVal New_Value As Boolean)
    m_LinesCurve = New_Value
    PropertyChanged "LinesCurve"
    Refresh
End Property

Public Property Get LinesWidth() As Long
    LinesWidth = m_LinesWidth
End Property

Public Property Let LinesWidth(ByVal New_Value As Long)
    m_LinesWidth = New_Value
    PropertyChanged "LinesWidth"
    Refresh
End Property

Public Property Get VerticalLines() As Boolean
    VerticalLines = m_VerticalLines
End Property

Public Property Let VerticalLines(ByVal New_Value As Boolean)
    m_VerticalLines = New_Value
    PropertyChanged "VerticalLines"
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

Public Property Get LabelsVisible() As Boolean
    LabelsVisible = m_LabelsVisible
End Property

Public Property Let LabelsVisible(ByVal New_Value As Boolean)
    m_LabelsVisible = New_Value
    PropertyChanged "LabelsVisible"
    Refresh
End Property

Public Property Get HorizontalLines() As Boolean
    HorizontalLines = m_HorizontalLines
End Property

Public Property Let HorizontalLines(ByVal New_Value As Boolean)
    m_HorizontalLines = New_Value
    PropertyChanged "HorizontalLines"
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

Public Property Get ChartBarOrientation() As ChartBarOrientation
    ChartBarOrientation = m_ChartBarOrientation
End Property

Public Property Let ChartBarOrientation(ByVal New_Value As ChartBarOrientation)
    m_ChartBarOrientation = New_Value
    PropertyChanged "ChartBarOrientation"
    Refresh
End Property

Public Property Get LegendAlign() As ucCB_LegendAlign
    LegendAlign = m_LegendAlign
End Property

Public Property Let LegendAlign(ByVal New_Value As ucCB_LegendAlign)
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

Public Property Get AxisXVisible() As Boolean
    AxisXVisible = m_AxisXVisible
End Property

Public Property Let AxisXVisible(ByVal New_Value As Boolean)
    m_AxisXVisible = New_Value
    PropertyChanged "AxisXVisible"
    Refresh
End Property

Public Property Get AxisYVisible() As Boolean
    AxisYVisible = m_AxisYVisible
End Property

Public Property Let AxisYVisible(ByVal New_Value As Boolean)
    m_AxisYVisible = New_Value
    PropertyChanged "AxisYVisible"
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
    TitleForeColor = m_TitleForeColor
End Property

Public Property Let TitleForeColor(ByVal New_Value As OLE_COLOR)
    m_TitleForeColor = mycolor(New_Value)
    PropertyChanged "TitleForeColor"
    Refresh
End Property

Public Property Get LabelsPositions() As brLabelsPositions
    LabelsPositions = m_LabelsPositions
End Property

Public Property Let LabelsPositions(ByVal New_Value As brLabelsPositions)
    m_LabelsPositions = New_Value
    PropertyChanged "LabelsPositions"
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

Public Property Get BorderRound() As Long
    BorderRound = m_BorderRound
End Property

Public Property Let BorderRound(ByVal New_Value As Long)
    m_BorderRound = New_Value
    PropertyChanged "BorderRound"
    Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_Value As OLE_COLOR)
    m_BorderColor = mycolor(New_Value)
    PropertyChanged "BorderColor"
    Refresh
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Title = .ReadProperty("Title", Ambient.DisplayName)
        m_BackColor = .ReadProperty("BackColor", vbWhite)
        m_BackColorOpacity = .ReadProperty("BackColorOpacity", 100)
        m_ForeColor = .ReadProperty("ForeColor", vbBlack)
        m_LinesColor = .ReadProperty("LinesColor", &HF4F4F4)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        m_FillOpacity = .ReadProperty("FillOpacity", 50)
        m_Border = .ReadProperty("Border", False)
        m_LinesCurve = .ReadProperty("LinesCurve", False)
        m_LinesWidth = .ReadProperty("LinesWidth", 1)
        m_VerticalLines = .ReadProperty("VerticalLines", False)
        m_FillGradient = .ReadProperty("FillGradient", False)
        m_LabelsVisible = .ReadProperty("LabelsVisible", False)
        m_HorizontalLines = .ReadProperty("HorizontalLines", True)
        m_ChartStyle = .ReadProperty("ChartStyle", CS_GroupedColumn)
        m_ChartBarOrientation = .ReadProperty("ChartBarOrientation", CO_Vertical)
        m_LegendAlign = .ReadProperty("LegendAlign", LA_RIGHT)
        m_LegendVisible = .ReadProperty("LegendVisible", True)
        m_AxisXVisible = .ReadProperty("AxisXVisible", True)
        m_AxisYVisible = .ReadProperty("AxisYVisible", True)
        Set m_TitleFont = .ReadProperty("TitleFont", Ambient.Font)
        m_TitleForeColor = .ReadProperty("TitleForeColor", Ambient.ForeColor)
        m_AxisMax = .ReadProperty("AxisMax", 0)
        m_AxisMin = .ReadProperty("AxisMin", 0)
        m_LabelsPositions = .ReadProperty("LabelsPositions", LP_CENTER)
        m_LabelsFormats = .ReadProperty("LabelsFormats", "{V}")
        m_BorderRound = .ReadProperty("BorderRound", 0)
        m_BorderColor = .ReadProperty("BorderColor", &HF4F4F4)
    End With
        
  '  If Not Ambient.UserMode Then Call Example
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Title", m_Title, Ambient.DisplayName
        .WriteProperty "BackColor", m_BackColor, vbWhite
        .WriteProperty "BackColorOpacity", m_BackColorOpacity, 100
        .WriteProperty "ForeColor", m_ForeColor, vbBlack
        .WriteProperty "LinesColor", m_LinesColor, &HF4F4F4
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "FillOpacity", m_FillOpacity, 50
        .WriteProperty "Border", m_Border, False
        .WriteProperty "LinesCurve", m_LinesCurve, False
        .WriteProperty "LinesWidth", m_LinesWidth, 1
        .WriteProperty "VerticalLines", m_VerticalLines, False
        .WriteProperty "FillGradient", m_FillGradient, False
        .WriteProperty "LabelsVisible", m_LabelsVisible, False
        .WriteProperty "HorizontalLines", m_HorizontalLines, True
        .WriteProperty "ChartStyle", m_ChartStyle, CS_GroupedColumn
        .WriteProperty "ChartBarOrientation", m_ChartBarOrientation, CO_Vertical
        .WriteProperty "LegendAlign", m_LegendAlign, LA_RIGHT
        .WriteProperty "LegendVisible", m_LegendVisible, True
        .WriteProperty "AxisXVisible", m_AxisXVisible, True
        .WriteProperty "AxisYVisible", m_AxisYVisible, True
        .WriteProperty "TitleFont", m_TitleFont, Ambient.Font
        .WriteProperty "TitleForeColor", m_TitleForeColor, Ambient.ForeColor
        .WriteProperty "AxisMax", m_AxisMax, 0
        .WriteProperty "AxisMin", m_AxisMin, 0
        .WriteProperty "LabelsPositions", m_LabelsPositions, LP_CENTER
        .WriteProperty "LabelsFormats", m_LabelsFormats, "{V}"
        .WriteProperty "BorderRound", m_BorderRound, 0
        .WriteProperty "BorderColor", m_BorderColor, &HF4F4F4
    End With
End Sub

Private Sub UserControl_InitProperties()
    m_Title = Ambient.DisplayName
    m_BackColor = vbWhite
    m_BackColorOpacity = 100
    m_ForeColor = vbBlack
    m_LinesColor = &HF4F4F4
    Set UserControl.Font = Ambient.Font
    m_FillOpacity = 50
    m_Border = False
    m_LinesWidth = 1
    m_VerticalLines = False
    m_FillGradient = False
    m_HorizontalLines = True
    m_ChartStyle = CS_GroupedColumn
    m_ChartBarOrientation = CO_Vertical
    m_LegendAlign = LA_RIGHT
    m_LegendVisible = True
    m_AxisXVisible = True
    m_AxisYVisible = True
    m_TitleFont.Name = UserControl.Font.Name
    m_TitleFont.Size = UserControl.Font.Size + 8
    m_AxisMax = 0
    m_AxisMin = 0
    m_LabelsPositions = LP_CENTER
    m_LabelsFormats = "{V}"
    m_BorderRound = 0
    m_BorderColor = &HF4F4F4
End Sub

Private Function GetTextSize(ByVal hGraphics As Long, ByVal Text As String, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal bWordWrap As Boolean, ByRef SZ As SIZEF) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RECTF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont

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
        

    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(UserControl.hDC, LOGPIXELSY), 72)

    layoutRect.Width = Width * nScale: layoutRect.Height = Height * nScale

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
      
    Dim bb As RECTF, CF As Long, LF As Long

    GdipMeasureString hGraphics, StrPtr(Text), -1, hFont, layoutRect, hFormat, bb, CF, LF

    SZ.Width = bb.Width
    SZ.Height = bb.Height
    
    GdipDeleteFont hFont
    GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
End Function

Private Function DrawText(ByVal hGraphics As Long, ByVal Text As String, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal oFont As StdFont, ByVal ForeColor As Long, Optional HAlign As ucCB_TextAlignmentH, Optional VAlign As TextAlignmentV, Optional bWordWrap As Boolean, Optional Angle As Single) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RECTF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim hDC As Long
    Dim W As Single, H As Single
    W = Width
    H = Height
    If Angle <> 0 Then
        Select Case Angle
            Case Is <= 90
                W = Width + Angle * (Height - Width) / 90
            Case Is < 180
                W = Width + (180 - Angle) * (Height - Width) / 90
            Case Is < 270
                W = Width + (Angle Mod 90) * (Height - Width) / 90
            Case Else
                W = Width + (360 - Angle) * (Height - Width) / 90
         End Select
         
        X = X - ((W - Width) / 2)
     
        Width = W
        
        GdipTranslateWorldTransform hGraphics, X + Width / 2, Y + Height / 2, 0
        GdipRotateWorldTransform hGraphics, Angle, 0
        GdipTranslateWorldTransform hGraphics, -(X + Width / 2), -(Y + Height / 2), 0
    End If
        

  
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) <> GDIP_OK Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) <> GDIP_OK Then Exit Function
    End If

    If GdipCreateStringFormat(0, 0, hFormat) = GDIP_OK Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
    End If
        
    If oFont.bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        

    hDC = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(UserControl.hDC, LOGPIXELSY), 72)
    ReleaseDC 0&, hDC

    layoutRect.Left = X: layoutRect.top = Y
    layoutRect.Width = Width: layoutRect.Height = Height

    If GdipCreateSolidFill(ForeColor, hBrush) = GDIP_OK Then
        If GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont) = GDIP_OK Then
            GdipDrawString hGraphics, StrPtr(Text), -1, hFont, layoutRect, hFormat, hBrush
            GdipDeleteFont hFont
        End If
        GdipDeleteBrush hBrush
    End If
    
    If hFormat Then GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
    If Angle <> 0 Then GdipResetWorldTransform hGraphics

End Function

Public Function GetWindowsDPI() As Double
    Dim hDC As Long, LPX  As Double
    hDC = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    ReleaseDC 0, hDC

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If UserControl.enabled Then
        HitResult = vbHitResultHit
        If Ambient.UserMode Then
            Dim pt As POINTL

            If tmrMOUSEOVER.Interval = 0 Then
                GetCursorPos pt
                ScreenToClient UserControl.ContainerHwnd, pt
                m_PT.X = pt.X - X
                m_PT.Y = pt.Y - Y
    
                m_Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
                m_Top = ScaleY(Extender.top, vbContainerSize, UserControl.ScaleMode)
 
              
                tmrMOUSEOVER.Interval = 1
                'RaiseEvent MouseEnter
            End If

        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    nScale = GetWindowsDPI
    Set cAxisItem = New Collection
    Set m_TitleFont = New StdFont
    mHotSerie = -1
End Sub

Public Property Get AxisMax() As Single
    AxisMax = m_AxisMax
End Property

Public Property Let AxisMax(ByVal New_Value As Single)
    m_AxisMax = New_Value
    PropertyChanged "AxisMax"
    Refresh
End Property

Public Property Get AxisMin() As Single
    AxisMin = m_AxisMin
End Property

Public Property Let AxisMin(ByVal New_Value As Single)
    m_AxisMin = New_Value
    PropertyChanged "AxisMin"
    Refresh
End Property


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
    If m_ChartBarOrientation = CO_Vertical Then
        DrawVertical
    Else
        DrawHorizontal
    End If
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub

Private Function ManageGDIToken(ByVal projectHwnd As Long) As Long ' by LaVolpe
    If projectHwnd = 0& Then Exit Function
    
    Dim hwndGDIsafe     As Long                 'API window to monitor IDE shutdown
    
    Do
        hwndGDIsafe = GetParent(projectHwnd)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    ' ok, got the highest level parent, now find highest level owner
    Do
        hwndGDIsafe = GetWindow(projectHwnd, GW_OWNER)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    
    hwndGDIsafe = FindWindowEx(projectHwnd, 0&, "Static", "GDI+Safe Patch")
    
    If hwndGDIsafe Then
        ManageGDIToken = hwndGDIsafe    ' we already have a manager running for this VB instance
        Exit Function                   ' can abort
    End If
    
    Dim gdiSI           As GDIPlusStartupInput  'GDI+ startup info
    Dim gToken          As Long                 'GDI+ instance token
    
    On Error Resume Next
    gdiSI.GdiPlusVersion = 1                    ' attempt to start GDI+
    GdiplusStartup gToken, gdiSI
    If gToken = 0& Then                         ' failed to start
        If Err Then Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Dim z_ScMem         As Long                 'Thunk base address
    Dim z_Code()        As Long                 'Thunk machine-code initialised here
    Dim nAddr           As Long                 'hwndGDIsafe prev window procedure

    Const WNDPROC_OFF   As Long = &H30          'Offset where window proc starts from z_ScMem
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const MEM_LEN       As Long = &HD4          'Byte length of thunk
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        ' we make the api window a child so we can use FindWindowEx to locate it easily
        hwndGDIsafe = CreateWindowExA(0&, "Static", "GDI+Safe Patch", WS_CHILD, 0&, 0&, 0&, 0&, projectHwnd, 0&, App.hInstance, ByVal 0&)
        If hwndGDIsafe <> 0 Then
        
            ReDim z_Code(0 To MEM_LEN \ 4 - 1)
        
            z_Code(12) = &HD231C031: z_Code(13) = &HBBE58960: z_Code(14) = &H12345678: z_Code(15) = &H3FFF631: z_Code(16) = &H74247539: z_Code(17) = &H3075FF5B: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875
            z_Code(20) = &H2C73FF24: z_Code(21) = &H890853FF: z_Code(22) = &HBFF1C45: z_Code(23) = &H2287D81: z_Code(24) = &H75000000: z_Code(25) = &H443C707: z_Code(26) = &H2&: z_Code(27) = &H2C753339: z_Code(28) = &H2047B81: z_Code(29) = &H75000000
            z_Code(30) = &H2C73FF23: z_Code(31) = &HFFFFFC68: z_Code(32) = &H2475FFFF: z_Code(33) = &H681C53FF: z_Code(34) = &H12345678: z_Code(35) = &H3268&: z_Code(36) = &HFF565600: z_Code(37) = &H43892053: z_Code(38) = &H90909020: z_Code(39) = &H10C261
            z_Code(40) = &H562073FF: z_Code(41) = &HFF2453FF: z_Code(42) = &H53FF1473: z_Code(43) = &H2873FF18: z_Code(44) = &H581053FF: z_Code(45) = &H89285D89: z_Code(46) = &H45C72C75: z_Code(47) = &H800030: z_Code(48) = &H20458B00: z_Code(49) = &H89145D89
            z_Code(50) = &H81612445: z_Code(51) = &H4C4&: z_Code(52) = &HC63FF00

            z_Code(1) = 0                                                   ' shutDown mode; used internally by ASM
            z_Code(2) = zFnAddr("user32", "CallWindowProcA")                ' function pointer CallWindowProc
            z_Code(3) = zFnAddr("kernel32", "VirtualFree")                  ' function pointer VirtualFree
            z_Code(4) = zFnAddr("kernel32", "FreeLibrary")                  ' function pointer FreeLibrary
            z_Code(5) = gToken                                              ' Gdi+ token
            z_Code(10) = LoadLibrary("gdiplus")                             ' library pointer (add reference)
            z_Code(6) = GetProcAddress(z_Code(10), "GdiplusShutdown")       ' function pointer GdiplusShutdown
            z_Code(7) = zFnAddr("user32", "SetWindowLongA")                 ' function pointer SetWindowLong
            z_Code(8) = zFnAddr("user32", "SetTimer")                       ' function pointer SetTimer
            z_Code(9) = zFnAddr("user32", "KillTimer")                      ' function pointer KillTimer
        
            z_Code(14) = z_ScMem                                            ' ASM ebx start point
            z_Code(34) = z_ScMem + WNDPROC_OFF                              ' subclass window procedure location
        
            RtlMoveMemory z_ScMem, VarPtr(z_Code(0)), MEM_LEN               'Copy the thunk code/data to the allocated memory
        
            nAddr = SetWindowLong(hwndGDIsafe, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Subclass our API window
            RtlMoveMemory z_ScMem + 44, VarPtr(nAddr), 4& ' Add prev window procedure to the thunk
            gToken = 0& ' zeroize so final check below does not release it
            
            ManageGDIToken = hwndGDIsafe    ' return handle of our GDI+ manager
        Else
            VirtualFree z_ScMem, 0, MEM_RELEASE     ' failure - release memory
            z_ScMem = 0&
        End If
    Else
        VirtualFree z_ScMem, 0, MEM_RELEASE           ' failure - release memory
        z_ScMem = 0&
    End If
    
    If gToken Then GdiplusShutdown gToken       ' release token if error occurred
    
End Function

'Funcion para combinar dos colores
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

Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)  'Get the specified procedure address
End Function

Private Function SafeRange(Value, Min, Max)
    If Value < Min Then
        SafeRange = Min
    ElseIf Value > Max Then
        SafeRange = Max
    Else
        SafeRange = Value
    End If
End Function

Public Function RGBtoARGB(ByVal RGBcolor As Long, ByVal Opacity As Long) As Long
    'By LaVople
    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to RGBA format
    ' Passing VB system color constants is allowed, i.e., vbButtonFace
    ' Pass Opacity as a value from 0 to 255

    If (RGBcolor And &H80000000) Then RGBcolor = GetSysColor(RGBcolor And &HFF&)
    RGBtoARGB = (RGBcolor And &HFF00&) Or (RGBcolor And &HFF0000) \ &H10000 Or (RGBcolor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
    
End Function

Private Function GetMax() As Single
    Dim i As Long, j As Long, M As Single
    Dim sum As Single
    If SerieCount = 0 Then Exit Function
    If (m_ChartStyle = CS_StackedBars) Or (m_ChartStyle = CS_StackedBarsPercent) Then
        For i = 1 To m_Serie(0).Values.Count
            sum = 0
            For j = 0 To SerieCount - 1
                sum = sum + m_Serie(j).Values(i)
            Next
            
            If M < sum Then M = sum
        Next
    Else
        For i = 0 To SerieCount - 1
           For j = 1 To m_Serie(i).Values.Count
               If M < m_Serie(i).Values(j) Then
                   M = m_Serie(i).Values(j)
               End If
           Next
        Next
    End If
    GetMax = M

End Function

Private Function GetMin() As Single
    Dim i As Long, j As Long, M As Single
    Dim sum As Single
    If SerieCount = 0 Then Exit Function
    If (m_ChartStyle = CS_StackedBars) Or (m_ChartStyle = CS_StackedBarsPercent) Then
        For i = 1 To m_Serie(0).Values.Count
            sum = 0
            For j = 0 To SerieCount - 1
                If m_Serie(j).Values(i) < 0 Then
                    sum = sum + m_Serie(j).Values(i)
                End If
            Next

            If M > sum Then M = sum
        Next
    Else
        For i = 0 To SerieCount - 1
           For j = 1 To m_Serie(i).Values.Count
               If m_Serie(i).Values(j) < M Then
                   M = m_Serie(i).Values(j)
               End If
           Next
        Next
    End If
    GetMin = M

End Function


Private Function SumSerieValues(Index As Long, Optional bPositives As Boolean) As Single
    Dim i As Long
    Dim M As Single
    For i = 0 To SerieCount - 1
        M = M + Abs(m_Serie(i).Values(Index))
    Next
    SumSerieValues = M
End Function
'*1
Private Sub DrawVertical()
    Dim hGraphics As Long, hPath As Long
    Dim hBrush As Long, hPen As Long
    Dim mRect As RectL
    Dim Min As Single, Max As Single
    Dim iStep As Single
    Dim nVal As Single
    Dim NumDecim As Single
    Dim forLines As Single, toLines As Single
    Dim i As Single, j As Long
    Dim mHeight As Single
    Dim mWidth As Single
    Dim PtDistance As Single
    Dim AxisDistance As Single
    'Dim PT2() As RECTL
    Dim mPenWidth As Single
    Dim MarginLeft As Single
    Dim MarginRight As Single
    Dim TopHeader As Single
    Dim Footer As Single
    Dim TextWidth As Single
    Dim TextHeight As Single
    Dim XX As Single, YY As Single
    Dim yRange As Single
    Dim lForeColor As Long
    Dim LW As Long
    Dim RectL As RectL
    Dim BarWidth As Single
    Dim lColor As Long
    Dim LabelsRect As RectL
    Dim AxisX As SIZEF
    Dim AxisY As SIZEF
    Dim PT16 As Single
    Dim PT24 As Single
    Dim ColRow As Integer
    Dim LineSpace As Single
    Dim TitleSize As SIZEF
    Dim sDisplay As String
    Dim ZeroPoint As Long
    Dim LastPositive() As Long, LastNegative() As Long
    Dim Value As Single
    Dim BarSpace As Single
    Dim RangeHeight As Single
    
    
    
    If GdipCreateFromHDC(hDC, hGraphics) Then Exit Sub
  
    Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
    Call GdipSetCompositingQuality(hGraphics, &H3) 'CompositingQualityGammaCorrected
    
    'PT16 = 16 * nScale
    PT16 = (UserControl.ScaleWidth + UserControl.ScaleHeight) * 2.5 / 100
    
    PT24 = 24 * nScale
    mPenWidth = 1 * nScale
    LW = m_LinesWidth * nScale
    lForeColor = RGBtoARGB(m_ForeColor, 100)
    
    If SerieCount > 1 Then BarSpace = LW * 4
    
    Max = IIf(m_AxisMax > 0, m_AxisMax, GetMax())
    Min = IIf(m_AxisMin < 0, m_AxisMin, GetMin())
            
    If m_AxisXVisible Then
        For i = 1 To cAxisItem.Count
            TextWidth = UserControl.TextWidth(cAxisItem(i)) * 1.3
            TextHeight = UserControl.TextHeight(cAxisItem(i)) * 1.3
            If TextWidth > AxisX.Width Then AxisX.Width = TextWidth
            If TextHeight > AxisX.Height Then AxisX.Height = TextHeight
        Next

        If m_AxisAngle <> 0 Then
            With AxisX
                Select Case m_AxisAngle
                    Case Is <= 90
                        .Height = .Height + m_AxisAngle * (.Width - .Height) / 90
                    Case Is < 180
                        .Height = .Height + (180 - m_AxisAngle) * (.Width - .Height) / 90
                    Case Is < 270
                        .Height = .Height + (m_AxisAngle Mod 90) * (.Width - .Height) / 90
                    Case Else
                        .Height = .Height + (360 - m_AxisAngle) * (.Width - .Height) / 90
                 End Select
             End With
        End If
    End If
    
    If m_AxisYVisible Then
        Value = IIf(Len(CStr(Max)) > Len(CStr(Min)), Max, Min)
        sDisplay = Replace(m_LabelsFormats, "{V}", Value)
        sDisplay = Replace(sDisplay, "{LF}", vbLf)
        If Len(sDisplay) = 1 Then sDisplay = "0.9"
        AxisY.Width = UserControl.TextWidth(CStr(sDisplay)) * 1.5
        AxisY.Height = UserControl.TextHeight(CStr(sDisplay)) * 1.5
    End If

    
    If m_LegendVisible Then
        For i = 0 To SerieCount - 1
            m_Serie(i).TextHeight = UserControl.TextHeight(m_Serie(i).SerieName) * 1.5
            m_Serie(i).TextWidth = UserControl.TextWidth(m_Serie(i).SerieName) * 1.5 + m_Serie(i).TextHeight
        Next
    End If

    If Len(m_Title) Then
        Call GetTextSize(hGraphics, m_Title, UserControl.ScaleWidth, 0, m_TitleFont, True, TitleSize)
    End If
    
    MarginRight = PT16
    TopHeader = PT16 + TitleSize.Height
    MarginLeft = PT16 + AxisY.Width
    Footer = PT16 + AxisX.Height
    
    mWidth = UserControl.ScaleWidth - MarginLeft - MarginRight
    mHeight = UserControl.ScaleHeight - TopHeader - Footer
    
    If m_LegendVisible Then
        ColRow = 1
        Select Case m_LegendAlign
            Case LA_RIGHT, LA_LEFT
                With LabelsRect
                    TextWidth = 0
                    TextHeight = 0
                    For i = 0 To SerieCount - 1
                        If TextHeight + m_Serie(i).TextHeight > mHeight Then
                            .Width = .Width + TextWidth
                            ColRow = ColRow + 1
                            TextWidth = 0
                            TextHeight = 0
                        End If
    
                        TextHeight = TextHeight + m_Serie(i).TextHeight
                        .Height = .Height + m_Serie(i).TextHeight
    
                        If TextWidth < m_Serie(i).TextWidth Then
                            TextWidth = m_Serie(i).TextWidth '+ PT16
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
             
                    .Height = m_Serie(0).TextHeight + PT16 / 2
                    TextWidth = 0
                    For i = 0 To SerieCount - 1
                        If TextWidth + m_Serie(i).TextWidth > mWidth Then
                            .Height = .Height + m_Serie(i).TextHeight
                            ColRow = ColRow + 1
                            TextWidth = 0
                        End If
                        TextWidth = TextWidth + m_Serie(i).TextWidth
                        .Width = .Width + m_Serie(i).TextWidth
                    Next
                    If m_LegendAlign = LA_TOP Then
                        TopHeader = TopHeader + .Height
                    End If
                    mHeight = mHeight - .Height
                End With
        End Select
    End If
    
    If cAxisItem.Count Then
        AxisDistance = (mWidth - mPenWidth) / cAxisItem.Count
    End If
    
    If SerieCount > 0 Then
        PtDistance = (mWidth - mPenWidth) / m_Serie(0).Values.Count
    End If
    
    If (m_ChartStyle = CS_StackedBars) Or (m_ChartStyle = CS_StackedBarsPercent) Then
        BarWidth = (PtDistance / 2)
    Else
        BarWidth = (PtDistance / (SerieCount + 0.5))
    End If
    
    LineSpace = BarWidth * 20 / 100
    NumDecim = 1
    
    If m_AxisMin Then forLines = m_AxisMin
    If m_AxisMax Then toLines = m_AxisMax
  
    If m_ChartStyle = CS_StackedBarsPercent Then
        iStep = 10
        If Max > 0 Then toLines = 100
        If Min < 0 Then forLines = -100
        Do
            RangeHeight = (mHeight / ((toLines + Abs(forLines)) / (iStep * NumDecim)))
            
            If RangeHeight < AxisY.Height Then
                Select Case iStep
                    Case Is = 10: iStep = 20
                    Case Is = 20: iStep = 50
                    Case Is = 50: iStep = 100: Exit Do
                End Select
            Else
                Exit Do
            End If
        Loop
    Else
        nVal = Max + Abs(Min)

        Do While nVal > 9.5
            nVal = nVal / 9.99
            NumDecim = NumDecim * 10
        Loop

        Select Case nVal
            Case 0 To 1.999999
                iStep = 0.2
            Case 2 To 4.799999
                iStep = 0.5
            Case 4.8 To 9.599999
                iStep = 1
        End Select
        
        Dim nDec As Single
        nDec = 1
        Do
            If nDec * iStep * NumDecim > IIf(Max > Abs(Min), Max, Abs(Min)) * 3 Then Exit Do
            
            If Max > 0 Then
                If m_AxisMax = 0 Then
                    toLines = CInt((Max / NumDecim + iStep) / iStep) * (iStep * NumDecim)
                End If
            End If
    
            If Min < 0 Then
                If m_AxisMin = 0 Then
                    forLines = CInt((Min / (iStep * NumDecim)) - 1) * (iStep * NumDecim)
                End If
            End If
            
            RangeHeight = (mHeight / ((toLines + Abs(forLines)) / (iStep * NumDecim)))
            
            If RangeHeight < AxisY.Height Then

                
                Select Case iStep
                    Case Is = 0.2 * nDec: iStep = 0.5 * nDec
                    Case Is = 0.5 * nDec: iStep = 1 * nDec
                    Case Is = 1 * nDec: nDec = nDec * 10: iStep = 0.2 * nDec
                End Select
            Else
                Exit Do
            End If
        Loop
    End If
    
    Dim RECTF As RECTF
    With RECTF
        .Width = UserControl.ScaleWidth - 1 * nScale
        .Height = UserControl.ScaleHeight - 1 * nScale
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

    'HORIZONTAL LINES AND vertical axis
    Call GdipCreatePen1(RGBtoARGB(m_LinesColor, 100), mPenWidth, &H2, hPen)
    
    YY = TopHeader + mHeight
    yRange = forLines
    
    If toLines = 0 And forLines = 0 Then toLines = 1
    RangeHeight = (mHeight / ((toLines + Abs(forLines)) / (iStep * NumDecim)))
    ZeroPoint = TopHeader + mHeight - RangeHeight * (Abs(forLines) / (iStep * NumDecim))
    
    For i = forLines / (iStep * NumDecim) To toLines / (iStep * NumDecim)
        If m_HorizontalLines Then
            GdipDrawLine hGraphics, hPen, MarginLeft, YY, UserControl.ScaleWidth - MarginRight - mPenWidth, YY
        End If
        
        If m_AxisYVisible Then
            sDisplay = Replace(m_LabelsFormats, "{V}", yRange)
            sDisplay = Replace(sDisplay, "{LF}", vbLf)
            DrawText hGraphics, sDisplay, 0, YY - RangeHeight / 2, MarginLeft - 8 * nScale, RangeHeight, UserControl.Font, lForeColor, cRight, cMiddle
        End If
        YY = YY - RangeHeight
        yRange = yRange + CCur(iStep * NumDecim)
    Next
   
    If m_VerticalLines And SerieCount > 0 Then
        For i = 0 To m_Serie(0).Values.Count '- 1
            XX = MarginLeft + PtDistance * i
            GdipDrawLine hGraphics, hPen, XX, TopHeader, XX, TopHeader + mHeight + 4 * nScale
        Next
    End If
   
    GdipDeletePen hPen
    
    If ((m_ChartStyle = CS_StackedBars) Or (m_ChartStyle = CS_StackedBarsPercent)) And SerieCount > 0 Then
        ReDim LastPositive(m_Serie(0).Values.Count - 1)
        ReDim LastNegative(m_Serie(0).Values.Count - 1)
        For i = 0 To m_Serie(0).Values.Count - 1
            LastPositive(i) = ZeroPoint
            LastNegative(i) = ZeroPoint
        Next
    End If
    
    For i = 0 To SerieCount - 1
        'Calculo
        ReDim m_Serie(i).Rects(m_Serie(i).Values.Count - 1)
        If (m_ChartStyle = CS_StackedBars) Or (m_ChartStyle = CS_StackedBarsPercent) Then
        
            If m_ChartStyle = CS_StackedBarsPercent Then
                For j = 0 To m_Serie(i).Values.Count - 1
                    Max = SumSerieValues(j + 1, True)
                    Value = m_Serie(i).Values(j + 1)
                    
                    With m_Serie(i).Rects(j)
                        .Left = MarginLeft + PtDistance * j + BarWidth / 2
                        
                        If Value >= 0 Then
                            .Height = (Value * (ZeroPoint - TopHeader) / Max)
                            .top = LastPositive(j) - .Height
                            LastPositive(j) = .top
                        Else
                            .top = LastNegative(j)
                            .Height = (Abs(Value) * (TopHeader + mHeight - ZeroPoint) / Max)
                            LastNegative(j) = .top + .Height
                        End If
                        
                        .Width = BarWidth
                    End With
                
                Next
            Else
                
                For j = 0 To m_Serie(i).Values.Count - 1
                    Value = m_Serie(i).Values(j + 1)
                    
                    With m_Serie(i).Rects(j)
                        .Left = MarginLeft + PtDistance * j + BarWidth / 2
                        If Value >= 0 Then
                            .Height = (Value * (Max * (ZeroPoint - TopHeader) / toLines) / Max)
                            .top = LastPositive(j) - .Height
                            LastPositive(j) = .top
                        Else
                            .top = LastNegative(j)
                            .Height = (Value * (Min * (TopHeader + mHeight - ZeroPoint) / forLines) / Min)
                            LastNegative(j) = .top + .Height
                        End If
                        .Width = BarWidth
                    End With
                Next
            End If
        Else
            For j = 0 To m_Serie(i).Values.Count - 1
                Value = m_Serie(i).Values(j + 1)
                With m_Serie(i).Rects(j)
                    .Left = MarginLeft + PtDistance * j + BarWidth / 4 + BarWidth * i + BarSpace / 2
                    If Value >= 0 Then
                        .top = ZeroPoint - (Value * (ZeroPoint - TopHeader) / toLines)
                        .Height = ZeroPoint - .top
                    Else
                        .top = ZeroPoint
                        .Height = (Value * (TopHeader + mHeight - ZeroPoint) / forLines)
                    End If
                    .Width = BarWidth - BarSpace
                End With
            Next
        End If

        
        With RectL
            .top = TopHeader
            .Width = UserControl.ScaleWidth
            .Height = UserControl.ScaleHeight
        End With
        
        For j = 0 To UBound(m_Serie(i).Rects)
        
            If Not m_Serie(i).CustomColors Is Nothing Then
                lColor = m_Serie(i).CustomColors.item(j + 1)
            Else
                lColor = m_Serie(i).SeireColor
            End If
            
            If i = mHotSerie And (mHotBar = j Or mHotBar = -1) Then
                GdipCreatePen1 RGBtoARGB(lColor, 100), LW * 2, &H2, hPen
                lColor = ShiftColor(lColor, vbWhite, 90)
            Else
                GdipCreatePen1 RGBtoARGB(lColor, 100), LW, &H2, hPen
            End If

            If m_FillGradient Then
                GdipCreateLineBrushFromRectWithAngleI RectL, RGBtoARGB(lColor, m_FillOpacity), RGBtoARGB(vbWhite, IIf(m_FillOpacity < 100, 0, 100)), 90, 0, WrapModeTile, hBrush
            Else
                GdipCreateSolidFill RGBtoARGB(lColor, m_FillOpacity), hBrush
            End If
                            
            With m_Serie(i).Rects(j)
                GdipFillRectangleI hGraphics, hBrush, .Left, .top, .Width, .Height
                GdipDrawRectangleI hGraphics, hPen, .Left, .top, .Width, .Height
            End With
            
            GdipDeleteBrush hBrush
            GdipDeletePen hPen
        Next
      

        If m_LegendVisible Then
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
                                .top = TopHeader + mHeight / 2 - .Height / 2
                            Else
                                .top = TopHeader
                            End If
                        End If
                        
                        If TextWidth < m_Serie(i).TextWidth Then
                            TextWidth = m_Serie(i).TextWidth '+ PT16
                        End If
        
                        If TextHeight + m_Serie(i).TextHeight > mHeight Then
                             If i > 0 Then .Left = .Left + TextWidth
                            .top = TopHeader
                             TextHeight = 0
                        End If
                        m_Serie(i).LegendRect.Left = .Left
                        m_Serie(i).LegendRect.top = .top
                        m_Serie(i).LegendRect.Width = m_Serie(i).TextWidth
                        m_Serie(i).LegendRect.Height = m_Serie(i).TextHeight
                        
                        GdipCreateSolidFill RGBtoARGB(m_Serie(i).SeireColor, 100), hBrush
                        GdipFillRectangleI hGraphics, hBrush, .Left, .top + m_Serie(i).TextHeight / 4, m_Serie(i).TextHeight / 2, m_Serie(i).TextHeight / 2
                        GdipDeleteBrush hBrush
                        
                        DrawText hGraphics, m_Serie(i).SerieName, .Left + m_Serie(i).TextHeight / 1.5, .top, m_Serie(i).TextWidth, m_Serie(i).TextHeight, UserControl.Font, lForeColor, cLeft, cMiddle
                        TextHeight = TextHeight + m_Serie(i).TextHeight
                        .top = .top + m_Serie(i).TextHeight
                        
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
                                .top = PT16 + TitleSize.Height
                            Else
                                .top = TopHeader + mHeight + TitleSize.Height + PT16 / 2
                            End If
                        End If
        
                        If .Left + m_Serie(i).TextWidth - MarginLeft > mWidth Then
                            .Left = MarginLeft
                            .top = .top + m_Serie(i).TextHeight
                        End If
        
                        GdipCreateSolidFill RGBtoARGB(m_Serie(i).SeireColor, 100), hBrush
                        GdipFillRectangleI hGraphics, hBrush, .Left, .top + m_Serie(i).TextHeight / 4, m_Serie(i).TextHeight / 2, m_Serie(i).TextHeight / 2
                        GdipDeleteBrush hBrush
                        m_Serie(i).LegendRect.Left = .Left
                        m_Serie(i).LegendRect.top = .top
                        m_Serie(i).LegendRect.Width = m_Serie(i).TextWidth
                        m_Serie(i).LegendRect.Height = m_Serie(i).TextHeight
                        
                        DrawText hGraphics, m_Serie(i).SerieName, .Left + m_Serie(i).TextHeight / 1.5, .top, m_Serie(i).TextWidth, m_Serie(i).TextHeight, UserControl.Font, lForeColor, cLeft, cMiddle
                        .Left = .Left + m_Serie(i).TextWidth '+ m_Serie(i).TextHeight / 1.5
                    End With
            End Select
        End If

    Next

    
    If m_LabelsVisible Then
         For i = 0 To SerieCount - 1
            For j = 0 To m_Serie(i).Values.Count - 1
                mRect = m_Serie(i).Rects(j)
                With mRect
                    sDisplay = Replace(m_LabelsFormats, "{V}", m_Serie(i).Values(j + 1))
                    sDisplay = Replace(sDisplay, "{LF}", vbLf)
                    TextHeight = UserControl.TextHeight(sDisplay) * 1.3
                    TextWidth = UserControl.TextWidth(sDisplay) * 1.5
                    If (TextHeight > .Height Or m_LabelsPositions = LP_ABOBE) And m_ChartStyle = CS_GroupedColumn Then
                        .top = .top - TextHeight
                        .Height = TextHeight
                        lColor = RGBtoARGB(m_Serie(i).SeireColor, 100)
                    Else
                        If Not m_Serie(i).CustomColors Is Nothing Then
                            lColor = m_Serie(i).CustomColors(j)
                        Else
                            lColor = m_ForeColor
                        End If
                        If IsDarkColor(lColor) Then
                            lColor = RGBtoARGB(vbWhite, 100)
                        Else
                            lColor = RGBtoARGB(vbBlack, 100)
                        End If
                    End If
                    
                    If TextWidth > .Width Then
                        .Left = .Left + .Width / 2 - TextWidth / 2
                        .Width = TextWidth
                    End If
                    
                    
                    DrawText hGraphics, sDisplay, .Left, .top, .Width, .Height, UserControl.Font, lColor, cCenter, m_LabelsPositions
                End With
            Next
        Next
    End If
        
    'a line to overlap the base of the rectangle

    Call GdipCreatePen1(RGBtoARGB(m_LinesColor, 100), LW, &H2, hPen)
    GdipDrawLine hGraphics, hPen, MarginLeft, ZeroPoint, UserControl.ScaleWidth - MarginRight - mPenWidth, ZeroPoint
    GdipDeletePen hPen
    
    '*-
    'Horizontal Axis
    If m_AxisXVisible Then
        For i = 1 To cAxisItem.Count
            XX = MarginLeft + AxisDistance * (i - 1)
            DrawText hGraphics, cAxisItem(i), XX, TopHeader + mHeight, AxisDistance, Footer, UserControl.Font, lForeColor, m_AxisAlign, cMiddle, m_WordWrap, m_AxisAngle
        Next
    End If

    'Title
    If Len(m_Title) Then
        DrawText hGraphics, m_Title, 0, PT16 / 2, UserControl.ScaleWidth, TopHeader, m_TitleFont, RGBtoARGB(m_TitleForeColor, 100), cCenter, cTop, True
    End If


    ShowToolTips hGraphics, BarWidth

    Call GdipDeleteGraphics(hGraphics)
    

End Sub

Private Sub ShowToolTips(hGraphics As Long, BarWidth As Single)
    Dim i As Long, j As Long
    Dim sDisplay As String
    Dim bBold As Boolean
    Dim RECTF As RECTF
    Dim LW As Single
    Dim lForeColor As Long
    Dim TM As Single
    Dim SZ As SIZEF
    
    If mHotBar > -1 Then
        TM = UserControl.TextHeight("Aj") / 4
        lForeColor = RGBtoARGB(m_ForeColor, 100)
        LW = m_LinesWidth * nScale
        
        For i = 0 To SerieCount - 1
            For j = 1 To m_Serie(i).Values.Count
                
                Dim sText As String
                If mHotSerie = i And mHotBar = j - 1 Then

                    If cAxisItem.Count = m_Serie(i).Values.Count Then
                        sText = cAxisItem(j) & vbCrLf
                    End If
                    sDisplay = Replace(m_LabelsFormats, "{V}", m_Serie(i).Values(j))
                    sDisplay = Replace(sDisplay, "{LF}", vbLf)
                    sText = sText & m_Serie(i).SerieName & ": " & sDisplay

                    GetTextSize hGraphics, sText, 0, 0, UserControl.Font, False, SZ
                    
                    With RECTF

                        .Left = m_Serie(i).Rects(j - 1).Left + BarWidth / 2 - SZ.Width / 2
                        .top = m_Serie(i).Rects(j - 1).top - SZ.Height - 10 * nScale
                        .Width = SZ.Width + TM * 2
                        .Height = SZ.Height + TM * 2
                        
                        If .Left < 0 Then .Left = LW
                        If .top < 0 Then .top = LW
                        If .Left + .Width >= UserControl.ScaleWidth - LW Then .Left = UserControl.ScaleWidth - .Width - LW
                        If .top + .Height >= UserControl.ScaleHeight - LW Then .top = UserControl.ScaleHeight - .Height - LW
                    End With
                    
                    RoundRect hGraphics, RECTF, RGBtoARGB(m_BackColor, 90), RGBtoARGB(m_Serie(i).SeireColor, 80), TM


                    With RECTF
                        .Left = .Left + TM
                        .top = .top + TM
                        If cAxisItem.Count = m_Serie(i).Values.Count Then
                            DrawText hGraphics, cAxisItem(j), .Left, .top, .Width, 0, UserControl.Font, lForeColor, cLeft, cTop
                            GetTextSize hGraphics, cAxisItem(j), 0, 0, UserControl.Font, False, SZ
                            .top = .top + SZ.Height
                        End If
                        
                        DrawText hGraphics, m_Serie(i).SerieName & ": ", .Left, .top, .Width, 0, UserControl.Font, lForeColor, cLeft, cTop
                        GetTextSize hGraphics, m_Serie(i).SerieName & ": ", 0, 0, UserControl.Font, False, SZ
                        .Left = .Left + SZ.Width
                        bBold = UserControl.Font.bold
                        UserControl.Font.bold = True
                        DrawText hGraphics, sDisplay, .Left, .top, .Width, 0, UserControl.Font, lForeColor, cLeft, cTop
                        UserControl.Font.bold = bBold
                    End With
                               
                End If
            Next
        Next
    End If
End Sub
'*3
Private Sub DrawHorizontal()
    Dim hGraphics As Long, hPath As Long
    Dim hBrush As Long, hPen As Long
    Dim mRect As RectL
    Dim Min As Single, Max As Single
    Dim iStep As Single
    Dim nVal As Single
    Dim NumDecim As Single
    Dim forLines As Single, toLines As Single
    Dim i As Single, j As Long
    Dim mHeight As Single
    Dim mWidth As Single
    Dim PtDistance As Single
    Dim AxisDistance As Single
    Dim PT2() As RectL
    Dim mPenWidth As Single
    Dim MarginLeft As Single
    Dim MarginRight As Single
    Dim TopHeader As Single
    Dim Footer As Single
    Dim TextWidth As Single
    Dim TextHeight As Single
    Dim XX As Single, YY As Single
    Dim yRange As Single
    Dim lForeColor As Long
    Dim LW As Long
    Dim RectL As RectL
    Dim BarWidth As Single
    Dim lColor As Long
    Dim LabelsRect As RectL
    Dim AxisX As SIZEF
    Dim AxisY As SIZEF
    Dim PT16 As Single
    Dim PT24 As Single
    Dim ColRow As Integer
    Dim LineSpace As Single
    Dim TitleSize As SIZEF
    Dim sDisplay As String
    Dim ZeroPoint As Long
    Dim LastPositive() As Long, LastNegative() As Long
    Dim Value As Single
    Dim BarSpace As Single
    Dim RangeHeight As Single
    
    If GdipCreateFromHDC(hDC, hGraphics) Then Exit Sub
  
    Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
    Call GdipSetCompositingQuality(hGraphics, &H3) 'CompositingQualityGammaCorrected
        
    'PT16 = 16 * nScale
    PT16 = (UserControl.ScaleWidth + UserControl.ScaleHeight) * 2.5 / 100
    
    PT24 = 24 * nScale
    mPenWidth = 1 * nScale
    LW = m_LinesWidth * nScale
    lForeColor = RGBtoARGB(m_ForeColor, 100)
    If SerieCount > 1 Then BarSpace = LW * 4
    
    Max = IIf(m_AxisMax > 0, m_AxisMax, GetMax())
    Min = IIf(m_AxisMin < 0, m_AxisMin, GetMin())
    
    If m_AxisYVisible Then
        For i = 1 To cAxisItem.Count
            TextWidth = UserControl.TextWidth(cAxisItem(i)) * 1.3
            TextHeight = UserControl.TextHeight(cAxisItem(i)) * 1.5
            If TextWidth > AxisY.Width Then AxisY.Width = TextWidth
            If TextHeight > AxisY.Height Then AxisY.Height = TextHeight
        Next
        
        If m_AxisAngle <> 0 Then
            With AxisY
                Select Case m_AxisAngle
                    Case Is <= 90
                        .Width = .Width + m_AxisAngle * (.Height - .Width) / 90
                    Case Is < 180
                        .Width = .Width + (180 - m_AxisAngle) * (.Height - .Width) / 90
                    Case Is < 270
                        .Width = .Width + (m_AxisAngle Mod 90) * (.Height - .Width) / 90
                    Case Else
                        .Width = .Width + (360 - m_AxisAngle) * (.Height - .Width) / 90
                 End Select
             End With
        End If
    End If
    
    If m_AxisXVisible Then
        Value = IIf(Len(CStr(Max)) > Len(CStr(Min)), Max, Min)
        sDisplay = Replace(m_LabelsFormats, "{V}", Value)
        sDisplay = Replace(sDisplay, "{LF}", vbLf)
        AxisX.Width = UserControl.TextWidth(CStr(sDisplay)) * 1.5
        AxisX.Height = UserControl.TextHeight(CStr(sDisplay)) * 1.5
    End If
    
    If m_LegendVisible Then
        For i = 0 To SerieCount - 1
            m_Serie(i).TextHeight = UserControl.TextHeight(m_Serie(i).SerieName) * 1.5
            m_Serie(i).TextWidth = UserControl.TextWidth(m_Serie(i).SerieName) * 1.5 + m_Serie(i).TextHeight
        Next
    End If
    
    If Len(m_Title) Then
        Call GetTextSize(hGraphics, m_Title, UserControl.ScaleWidth, 0, m_TitleFont, True, TitleSize)
    End If
    
    MarginRight = PT16
    TopHeader = PT16 + TitleSize.Height
    MarginLeft = PT16 + AxisY.Width
    Footer = PT16 + AxisX.Height
    
    mWidth = UserControl.ScaleWidth - MarginLeft - MarginRight
    mHeight = UserControl.ScaleHeight - TopHeader - Footer
    
    If m_LegendVisible Then
        ColRow = 1
        Select Case m_LegendAlign
            Case LA_RIGHT, LA_LEFT
                With LabelsRect
                    TextWidth = 0
                    TextHeight = 0
                    For i = 0 To SerieCount - 1
                        If TextHeight + m_Serie(i).TextHeight > mHeight Then
                            .Width = .Width + TextWidth
                            ColRow = ColRow + 1
                            TextWidth = 0
                            TextHeight = 0
                        End If
    
                        TextHeight = TextHeight + m_Serie(i).TextHeight
                        .Height = .Height + m_Serie(i).TextHeight
    
                        If TextWidth < m_Serie(i).TextWidth Then
                            TextWidth = m_Serie(i).TextWidth '+ PT16
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
                    If SerieCount > 0 Then .Height = m_Serie(0).TextHeight + PT16 / 2
                    TextWidth = 0
                    For i = 0 To SerieCount - 1
                        If TextWidth + m_Serie(i).TextWidth > mWidth Then
                            .Height = .Height + m_Serie(i).TextHeight
                            ColRow = ColRow + 1
                            TextWidth = 0
                        End If
                        TextWidth = TextWidth + m_Serie(i).TextWidth
                        .Width = .Width + m_Serie(i).TextWidth
                    Next
                    If m_LegendAlign = LA_TOP Then
                        TopHeader = TopHeader + .Height
                    End If
                    mHeight = mHeight - .Height
                End With
        End Select
    End If

    If cAxisItem.Count Then
        AxisDistance = (mHeight - mPenWidth) / cAxisItem.Count
    End If
    
    If SerieCount > 0 Then
        PtDistance = (mHeight - mPenWidth) / m_Serie(0).Values.Count
    End If
    
    If (m_ChartStyle = CS_StackedBars) Or (m_ChartStyle = CS_StackedBarsPercent) Then
        BarWidth = (PtDistance / 2)
    Else
        BarWidth = (PtDistance / (SerieCount + 0.5))
    End If
    
    LineSpace = BarWidth * 20 / 100
    NumDecim = 1
    
    If m_AxisMin Then forLines = m_AxisMin
    If m_AxisMax Then toLines = m_AxisMax
  
    If m_ChartStyle = CS_StackedBarsPercent Then
        iStep = 10
        
        If Max > 0 Then toLines = 100
        If Min < 0 Then forLines = -100
        Do
            RangeHeight = (mWidth / ((toLines + Abs(forLines)) / (iStep * NumDecim)))
            If RangeHeight < AxisX.Width Then
                Select Case iStep
                    Case Is = 10: iStep = 20
                    Case Is = 20: iStep = 50
                    Case Is = 50: iStep = 100: Exit Do
                End Select
            Else
                Exit Do
            End If
        Loop
    Else
        nVal = Max + Abs(Min)

        Do While nVal > 9.5
            nVal = nVal / 9.99
            NumDecim = NumDecim * 10
        Loop

        Select Case nVal
            Case 0 To 1.999999
                iStep = 0.2
            Case 2 To 4.799999
                iStep = 0.5
            Case 4.8 To 9.599999
                iStep = 1
        End Select
        
        Dim nDec As Single
        nDec = 1
        Do
            If nDec * iStep * NumDecim > IIf(Max > Abs(Min), Max, Abs(Min)) Then Exit Do
             
            If Max > 0 Then
                If m_AxisMax = 0 Then
                    toLines = CInt((Max / NumDecim + iStep) / iStep) * (iStep * NumDecim)
                End If
            End If
    
            If Min < 0 Then
                If m_AxisMin = 0 Then
                    forLines = CInt((Min / (iStep * NumDecim)) - 1) * (iStep * NumDecim)
                End If
            End If
            
            RangeHeight = (mWidth / ((toLines + Abs(forLines)) / (iStep * NumDecim)))
           
            If RangeHeight < AxisX.Width Then

               
                Select Case iStep
                    Case Is = 0.2 * nDec: iStep = 0.5 * nDec
                    Case Is = 0.5 * nDec: iStep = 1 * nDec
                    Case Is = 1 * nDec: nDec = nDec * 10: iStep = 0.2 * nDec
                End Select
            Else
                Exit Do
            End If
        Loop
 
    End If

    Dim RECTF As RECTF
    With RECTF
        .Width = UserControl.ScaleWidth - 1 * nScale
        .Height = UserControl.ScaleHeight - 1 * nScale
    End With
    RoundRect hGraphics, RECTF, RGBtoARGB(m_BackColor, m_BackColorOpacity), RGBtoARGB(m_BorderColor, 100), m_BorderRound * nScale, m_Border

'     'Background
'     If m_BackColorOpacity > 0 Then
'         GdipCreateSolidFill RGBtoARGB(m_BackColor, m_BackColorOpacity), hBrush
'         GdipFillRectangleI hGraphics, hBrush, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
'         GdipDeleteBrush hBrush
'     End If
'
'     'Border
'     If m_Border Then
'         Call GdipCreatePen1(RGBtoARGB(m_LinesColor, 50), mPenWidth, &H2, hPen)
'         GdipDrawRectangleI hGraphics, hPen, mPenWidth / 2, mPenWidth / 2, UserControl.ScaleWidth - mPenWidth, UserControl.ScaleHeight - mPenWidth
'         GdipDeletePen hPen
'     End If

    'vertical LINES AND vertical axis
    Call GdipCreatePen1(RGBtoARGB(m_LinesColor, 100), mPenWidth, &H2, hPen)
     
    YY = TopHeader + mHeight
    XX = MarginLeft
    yRange = forLines
    If toLines = 0 And forLines = 0 Then toLines = 1
    
    RangeHeight = (mWidth / ((toLines + Abs(forLines)) / (iStep * NumDecim)))

    ZeroPoint = MarginLeft + RangeHeight * (Abs(forLines) / (iStep * NumDecim))
    
    For i = forLines / (iStep * NumDecim) To toLines / (iStep * NumDecim)
        If m_VerticalLines Then
            GdipDrawLine hGraphics, hPen, XX, TopHeader, XX, TopHeader + mHeight - mPenWidth
        End If
        
        If m_AxisXVisible Then
            sDisplay = Replace(m_LabelsFormats, "{V}", yRange)
            sDisplay = Replace(sDisplay, "{LF}", vbLf)
            DrawText hGraphics, sDisplay, XX - RangeHeight / 2, YY + 8 * nScale, RangeHeight, Footer, UserControl.Font, lForeColor, cCenter, cTop
            'DrawText hGraphics, sDisplay, 0, Yy - RangeHeight / 2, MarginLeft - 8 * nScale, RangeHeight, UserControl.Font, lForeColor, cRight, cMiddle

        End If

        XX = XX + RangeHeight
        yRange = yRange + CCur(iStep * NumDecim)
    Next
     

     If m_HorizontalLines And SerieCount > 0 Then
         For i = 0 To m_Serie(0).Values.Count
             YY = TopHeader + PtDistance * i
             GdipDrawLine hGraphics, hPen, MarginLeft, YY, MarginLeft + mWidth, YY
         Next
     End If
     
     GdipDeletePen hPen

    If ((m_ChartStyle = CS_StackedBars) Or (m_ChartStyle = CS_StackedBarsPercent)) And SerieCount > 0 Then
        ReDim LastPositive(m_Serie(0).Values.Count - 1)
        ReDim LastNegative(m_Serie(0).Values.Count - 1)
        For i = 0 To m_Serie(0).Values.Count - 1
            LastPositive(i) = ZeroPoint
            LastNegative(i) = ZeroPoint
        Next
    End If
    
    For i = 0 To SerieCount - 1
        'Calculo
        ReDim m_Serie(i).Rects(m_Serie(i).Values.Count - 1)
        If (m_ChartStyle = CS_StackedBars) Or (m_ChartStyle = CS_StackedBarsPercent) Then
        
            If m_ChartStyle = CS_StackedBarsPercent Then
                For j = 0 To m_Serie(i).Values.Count - 1
                    Max = SumSerieValues(j + 1, True)
                    Value = m_Serie(i).Values(j + 1)
                    
                    With m_Serie(i).Rects(j)
                        .top = TopHeader + PtDistance * j + BarWidth / 2
                        
                        If Value >= 0 Then
                            .Left = LastPositive(j)
                            .Width = (Value * (MarginLeft + mWidth - ZeroPoint) / Max)
                            
                            LastPositive(j) = .Left + .Width
                        Else
                            .Width = (Abs(Value) * (ZeroPoint - MarginLeft) / Max)
                            .Left = LastNegative(j) - .Width
                            LastNegative(j) = .Left
                        End If
                        
                        .Height = BarWidth
                    End With
                
                Next
            Else
                
                For j = 0 To m_Serie(i).Values.Count - 1
                    Value = m_Serie(i).Values(j + 1)
                    
                    With m_Serie(i).Rects(j)
                        .top = TopHeader + PtDistance * j + BarWidth / 2
                        
                        If Value >= 0 Then
                            .Left = LastPositive(j)
                            .Width = (Value * (Max * (MarginLeft + mWidth - ZeroPoint) / toLines) / Max)
                            LastPositive(j) = .Left + .Width
                            
                        Else
                            .Width = (Value * (Min * (ZeroPoint - MarginLeft) / forLines) / Min)
                            .Left = LastNegative(j) - .Width
                            LastNegative(j) = .Left
                            
                        End If

                        .Height = BarWidth
                    End With
                Next
            End If
        Else
            For j = 0 To m_Serie(i).Values.Count - 1
                Value = m_Serie(i).Values(j + 1)
            
                With m_Serie(i).Rects(j)
                    .top = TopHeader + PtDistance * j + BarWidth / 4 + BarSpace / 2 + BarWidth * i
                    If Value >= 0 Then
                        .Left = ZeroPoint
                        .Width = (Value * (MarginLeft + mWidth - ZeroPoint) / toLines)
                    Else
                        .Left = ZeroPoint - (Value * (ZeroPoint - MarginLeft) / forLines)
                        .Width = ZeroPoint - .Left
                    End If
                    .Height = BarWidth - BarSpace
                End With
            Next
        End If
        
        With RectL
            .top = TopHeader
            .Width = UserControl.ScaleWidth - MarginRight
            .Height = UserControl.ScaleHeight
        End With
         
        For j = 0 To UBound(m_Serie(i).Rects)
        
            If Not m_Serie(i).CustomColors Is Nothing Then
                lColor = m_Serie(i).CustomColors.item(j + 1)
            Else
                lColor = m_Serie(i).SeireColor
            End If
            
            If i = mHotSerie And (mHotBar = j Or mHotBar = -1) Then
                GdipCreatePen1 RGBtoARGB(lColor, 100), LW * 2, &H2, hPen
                lColor = ShiftColor(lColor, vbWhite, 90)
            Else
                GdipCreatePen1 RGBtoARGB(lColor, 100), LW, &H2, hPen
            End If

            If m_FillGradient Then
                GdipCreateLineBrushFromRectWithAngleI RectL, RGBtoARGB(lColor, m_FillOpacity), RGBtoARGB(vbWhite, IIf(m_FillOpacity < 100, 0, 100)), 180, 0, WrapModeTile, hBrush
            Else
                GdipCreateSolidFill RGBtoARGB(lColor, m_FillOpacity), hBrush
            End If
                            
            With m_Serie(i).Rects(j)
                GdipFillRectangleI hGraphics, hBrush, .Left, .top, .Width, .Height
                GdipDrawRectangleI hGraphics, hPen, .Left, .top, .Width, .Height
            End With
            
            GdipDeleteBrush hBrush
            GdipDeletePen hPen
        Next
     
        If m_LegendVisible Then
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
                                .top = TopHeader + mHeight / 2 - .Height / 2
                            Else
                                .top = TopHeader
                            End If
                        End If
                        
                        If TextWidth < m_Serie(i).TextWidth Then
                            TextWidth = m_Serie(i).TextWidth '+ PT16
                        End If
        
                        If TextHeight + m_Serie(i).TextHeight > mHeight Then
                             If i > 0 Then .Left = .Left + TextWidth
                            .top = TopHeader
                             TextHeight = 0
                        End If
                        m_Serie(i).LegendRect.Left = .Left
                        m_Serie(i).LegendRect.top = .top
                        m_Serie(i).LegendRect.Width = m_Serie(i).TextWidth
                        m_Serie(i).LegendRect.Height = m_Serie(i).TextHeight
                        
                        GdipCreateSolidFill RGBtoARGB(m_Serie(i).SeireColor, 100), hBrush
                        GdipFillRectangleI hGraphics, hBrush, .Left, .top + m_Serie(i).TextHeight / 4, m_Serie(i).TextHeight / 2, m_Serie(i).TextHeight / 2
                        GdipDeleteBrush hBrush
                        
                        DrawText hGraphics, m_Serie(i).SerieName, .Left + m_Serie(i).TextHeight / 1.5, .top, m_Serie(i).TextWidth, m_Serie(i).TextHeight, UserControl.Font, lForeColor, cLeft, cMiddle
                        TextHeight = TextHeight + m_Serie(i).TextHeight
                        .top = .top + m_Serie(i).TextHeight
                        
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
                                .top = PT16 + TitleSize.Height
                            Else
                                .top = TopHeader + mHeight + TitleSize.Height + PT16 / 2
                            End If
                        End If
        
                        If .Left + m_Serie(i).TextWidth - MarginLeft > mWidth Then
                            .Left = MarginLeft
                            .top = .top + m_Serie(i).TextHeight
                        End If
        
                        GdipCreateSolidFill RGBtoARGB(m_Serie(i).SeireColor, 100), hBrush
                        GdipFillRectangleI hGraphics, hBrush, .Left, .top + m_Serie(i).TextHeight / 4, m_Serie(i).TextHeight / 2, m_Serie(i).TextHeight / 2
                        GdipDeleteBrush hBrush
                        m_Serie(i).LegendRect.Left = .Left
                        m_Serie(i).LegendRect.top = .top
                        m_Serie(i).LegendRect.Width = m_Serie(i).TextWidth
                        m_Serie(i).LegendRect.Height = m_Serie(i).TextHeight
                        
                        DrawText hGraphics, m_Serie(i).SerieName, .Left + m_Serie(i).TextHeight / 1.5, .top, m_Serie(i).TextWidth, m_Serie(i).TextHeight, UserControl.Font, lForeColor, cLeft, cMiddle
                        .Left = .Left + m_Serie(i).TextWidth '+ m_Serie(i).TextHeight / 1.5
                    End With
            End Select
        End If

    Next
            
    If m_LabelsVisible Then
         For i = 0 To SerieCount - 1
            For j = 0 To m_Serie(i).Values.Count - 1
                mRect = m_Serie(i).Rects(j)
                With mRect
                    sDisplay = Replace(m_LabelsFormats, "{V}", m_Serie(i).Values(j + 1))
                    sDisplay = Replace(sDisplay, "{LF}", vbLf)
                    TextHeight = UserControl.TextHeight(sDisplay) * 1.3
                    TextWidth = UserControl.TextWidth(sDisplay) * 1.5
                    If (TextWidth > .Width Or m_LabelsPositions = LP_ABOBE) And m_ChartStyle = CS_GroupedColumn Then
                        .Left = .Left + .Width + PT16 / 10
                        .Width = TextWidth
                        lColor = RGBtoARGB(m_Serie(i).SeireColor, 100)
                    Else
                        If Not m_Serie(i).CustomColors Is Nothing Then
                            lColor = m_Serie(i).CustomColors(j)
                        Else
                            lColor = m_ForeColor
                        End If
                        If IsDarkColor(lColor) Then
                            lColor = RGBtoARGB(vbWhite, 100)
                        Else
                            lColor = RGBtoARGB(vbBlack, 100)
                        End If
                    End If
                    
                    If TextHeight > .Height Then
                        .top = .top + .Height / 2 - TextHeight / 2
                        .Height = TextHeight
                    End If
                    
                    
                    DrawText hGraphics, sDisplay, .Left, .top, .Width, .Height, UserControl.Font, lColor, m_LabelsPositions, cMiddle
                End With
            Next
        Next
    End If



     'a line to overlap the base of the rectangle
     Call GdipCreatePen1(RGBtoARGB(m_LinesColor, 100), LW, &H2, hPen)
     GdipDrawLine hGraphics, hPen, ZeroPoint, TopHeader, ZeroPoint, TopHeader + mHeight - mPenWidth
     GdipDeletePen hPen
     
     'vertical Axis
     If m_AxisYVisible Then
         For i = 1 To cAxisItem.Count
            YY = TopHeader + AxisDistance * (i - 1)
            If m_LegendAlign = LA_LEFT Then
                XX = LabelsRect.Left + LabelsRect.Width
            Else
                XX = PT16
            End If

            DrawText hGraphics, cAxisItem(i), XX, YY, MarginLeft - XX - PT16 / 10, AxisDistance, UserControl.Font, lForeColor, m_AxisAlign, cMiddle, m_WordWrap, m_AxisAngle
         Next
     End If

     'Title
     If Len(m_Title) Then
         DrawText hGraphics, m_Title, 0, PT16 / 2, UserControl.ScaleWidth, TopHeader, m_TitleFont, RGBtoARGB(m_TitleForeColor, 100), cCenter, cTop, True
     End If
     
     ShowToolTips hGraphics, BarWidth
     
     Call GdipDeleteGraphics(hGraphics)

End Sub



Private Function IsDarkColor(ByVal color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor color, 0, VarPtr(color)
    CopyMemory BGRA(0), color, 4&
  
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382

End Function

Private Sub RoundRect(ByVal hGraphics As Long, RECT As RECTF, ByVal BackColor As Long, ByVal BorderColor As Long, ByVal Round As Single, Optional bBorder As Boolean = True)
    Dim hPen As Long, hBrush As Long
    Dim mPath As Long
    
    GdipCreateSolidFill BackColor, hBrush
    If bBorder Then GdipCreatePen1 BorderColor, 1 * nScale, &H2, hPen

    If Round = 0 Then
        With RECT
            GdipFillRectangleI hGraphics, hBrush, .Left, .top, .Width, .Height
            If hPen Then GdipDrawRectangleI hGraphics, hPen, .Left, .top, .Width, .Height
        End With
    Else
        If GdipCreatePath(&H0, mPath) = 0 Then
            Round = Round * 2
            With RECT
                GdipAddPathArcI mPath, .Left, .top, Round, Round, 180, 90
                GdipAddPathArcI mPath, .Left + .Width - Round, .top, Round, Round, 270, 90
                GdipAddPathArcI mPath, .Left + .Width - Round, .top + .Height - Round, Round, Round, 0, 90
                GdipAddPathArcI mPath, .Left, .top + .Height - Round, Round, Round, 90, 90
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
Property Get enabled() As Boolean
    enabled = UserControl.enabled
End Property

Property Let enabled(ByVal bValue As Boolean)
    If UserControl.enabled <> bValue Then
        UserControl.enabled = bValue
    End If
    PropertyChanged
End Property


