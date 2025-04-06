VERSION 5.00
Begin VB.UserControl ShapeEx 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ClipBehavior    =   0  'None
   HasDC           =   0   'False
   ScaleHeight     =   192
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ToolboxBitmap   =   "ShapeEx.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Timer tmrPainting 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   144
      Top             =   108
   End
End
Attribute VB_Name = "ShapeEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--- for MST subclassing (1)
#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)

Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcAddressByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Private m_pSubclass         As IUnknown
'--- End for MST subclassing (1)

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type T_MSG
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As T_MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Const PM_REMOVE = &H1

Private Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Private Declare Function SetWorldTransform Lib "gdi32" (ByVal hDC As Long, lpXform As XFORM) As Long
Private Declare Function ModifyWorldTransform Lib "gdi32" (ByVal hDC As Long, lpXform As XFORM, ByVal iMode As Long) As Long
Private Const MWT_IDENTITY = 1
Private Const MWT_LEFTMULTIPLY = 2
'Private Const MWT_RIGHTMULTIPLY = 3

Private Const GM_ADVANCED = 2
'Private Const GM_COMPATIBLE = 1

Private Const Pi = 3.14159265358979

Private Const WM_USER As Long = &H400
Private Const WM_INVALIDATE As Long = WM_USER + 11 ' custom message

Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'Private Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function OleTranslateColor Lib "OleAut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

 
Private Type POINTL
    X As Long
    Y As Long
End Type

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipFillEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, ByVal Width As Single, ByVal unit As Long, pen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GdipFillPieI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal Count As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByRef pPoints As Any, ByVal Count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipDrawClosedCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As Long
Private Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long, ByVal tension As Single, ByVal FillMode As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipAddPathEllipseI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColor As Long, ByRef mCount As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipCreatePathGradientI Lib "gdiplus" (Points As POINTL, ByVal Count As Long, ByVal WrapMd As Long, polyGradient As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipDrawPieI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
                              
Private Enum FillModeConstants
    FillModeAlternate = &H0
    FillModeWinding = &H1
End Enum

Private Const UnitPixel = 2
Private Const QualityModeLow As Long = 1
Private Const SmoothingModeAntiAlias As Long = &H4

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Event Click()
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Event MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -604

Public Enum SEShapeConstants
    seShapeRectangle = vbShapeRectangle ' 0
    seShapeSquare = vbShapeSquare ' 1
    seShapeOval = vbShapeOval ' 2
    seShapeCircle = vbShapeCircle ' 3
    seShapeRoundedRectangle = vbShapeRoundedRectangle ' 4
    seShapeRoundedSquare = vbShapeRoundedSquare ' 5
    seShapeTriangleEquilateral = 6
    seShapeTriangleIsosceles = 7
    seShapeTriangleScalene = 8
    seShapeTriangleRight = 9
    seShapeRhombus = 10
    seShapeKite = 11
    seShapeDiamond = 12
    seShapeTrapezoid = 13
    seShapeParalellogram = 14
    seShapeSemicircle = 15
    seShapeRegularPolygon = 16
    seShapeStar = 17
    seShapeJaggedStar = 18
    seShapeHeart = 19
    seShapeArrow = 20
    seShapeCrescent = 21
    seShapeDrop = 22
    seShapeEgg = 23
    seShapeLocation = 24
    seShapeSpeaker = 25
    seShapeCloud = 26
    seShapeTalk = 27
    seShapeShield = 28
    seShapePie = 29
End Enum

Public Enum SEBackStyleConstants
    seTransparent = 0
    seOpaque = 1
End Enum

Public Enum SEFillStyle2Constants
    seFSSolid = vbFSSolid
    seFSTransparent = vbFSTransparent
End Enum

Public Enum SEQualityConstants
    seQualityLow = 0
    seQualityHigh = 1
End Enum

Public Enum SEStyle3DConstants
    seStyle3DNone = 0
    seStyle3DLight = 1
    seStyle3DShadow = 2
    seStyle3DBoth = 3
End Enum

Public Enum SEStyle3DEffectConstants
    seStyle3EffectAuto = 0
    seStyle3EffectDiffuse = 1
    seStyle3EffectGem = 2
End Enum

Public Enum SEFlippedConstants
    seFlippedNo = 0
    seFlippedHorizontally = 1
    seFlippedVertically = 2
    seFlippedBoth = 3
End Enum

Public Enum SESubclassingConstants
    seSCNo = 0 ' never ' In most cases subclassing is not necessary (and without subclassing is safer for running in the IDE), but in some special cases when the figure needs to be painted outside the control bounds, it may experience glitches.
    seSCYes = 1 ' always
    seSCNotInIDE = 2 ' compiled will use subclassing
    seSCNotInIDEDesignTime = 3 ' IDE run time and compiled will use subclassing
End Enum


' Property defaults
Private Const mdef_BackColor = vbWindowBackground
Private Const mdef_BackStyle = seTransparent
Private Const mdef_BorderColor = vbWindowText
Private Const mdef_Shape = seShapeRectangle
Private Const mdef_FillColor = vbBlack
Private Const mdef_FillStyle = vbFSTransparent
Private Const mdef_BorderStyle = vbBSSolid
Private Const mdef_BorderWidth = 1
Private Const mdef_Clickable = True
Private Const mdef_Quality = seQualityHigh
Private Const mdef_RotationDegrees = 0
Private Const mdef_Opacity = 100
Private Const mdef_Shift = 0
Private Const mdef_Vertices = 5
Private Const mdef_CurvingFactor = 0
Private Const mdef_Flipped = seFlippedNo
Private Const mdef_MousePointer = vbDefault
Private Const mdef_Style3D = seStyle3DNone
Private Const mdef_Style3DEffect = seStyle3EffectAuto
Private Const mdef_UseSubclassing = seSCNotInIDEDesignTime ' seSCYes

' Properties
Private mBackColor  As Long
Private mBackStyle As SEBackStyleConstants
Private mBorderColor As Long
Private mShape As SEShapeConstants
Private mFillColor As Long
Private mFillStyle  As Long
Private mBorderStyle  As BorderStyleConstants
Private mBorderWidth  As Integer
Private mClickable As Boolean
Private mQuality As SEQualityConstants
Private mRotationDegrees As Single
Private mOpacity As Single
Private mShift As Single
Private mVertices As Integer
Private mCurvingFactor As Integer
Private mFlipped As SEFlippedConstants
Private mMousePointer As Integer
Private mMouseIcon As StdPicture
Private mStyle3D As SEStyle3DConstants
Private mStyle3DEffect As SEStyle3DEffectConstants
Private mUseSubclassing As SESubclassingConstants

Private mAttached As Boolean
Private mShiftPutAutomatically As Single
Private mCurvingFactor2 As Single
Private mUserMode As Boolean
Private mSubclassed As Boolean
Private mDrawingOutsideUC As Boolean
Property Get enabled() As Boolean
    enabled = UserControl.enabled
End Property
Property Let enabled(ByVal bValue As Boolean)
    If UserControl.enabled <> bValue Then
        UserControl.enabled = bValue
    End If
    PropertyChanged
End Property
Private Sub tmrPainting_Timer()
    tmrPainting.enabled = False
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "UserMode" Then mUserMode = Ambient.UserMode
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If mUserMode Then
        If mClickable Then
            HitResult = vbHitResultHit
        End If
    Else
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_InitProperties()
    mBackColor = mdef_BackColor
    mBackStyle = mdef_BackStyle
    mBorderColor = mdef_BorderColor
    mShape = mdef_Shape
    mFillColor = mdef_FillColor
    mFillStyle = mdef_FillStyle
    mBorderStyle = mdef_BorderStyle
    mBorderWidth = mdef_BorderWidth
    mClickable = mdef_Clickable
    mQuality = mdef_Quality
    mRotationDegrees = mdef_RotationDegrees
    mOpacity = mdef_Opacity
    mShift = mdef_Shift
    mVertices = mdef_Vertices
    mCurvingFactor = mdef_CurvingFactor
    mFlipped = mdef_Flipped
    mMousePointer = mdef_MousePointer
    Set mMouseIcon = Nothing
    mStyle3D = mdef_Style3D
    mStyle3DEffect = mdef_Style3DEffect
    mUseSubclassing = mdef_UseSubclassing
    
    On Error Resume Next
    mUserMode = Ambient.UserMode
    On Error GoTo 0
    SetCurvingFactor2
    pvSubclass
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If mClickable Then
        If (KeyAscii = vbKeySpace) Then
            UserControl_Click
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    Dim hRgn As Long
    Dim rgnRect As RECT
    Dim hRgnExpand As Long
    Dim iExpandForPen As Long
    Dim iGMPrev As Long
    Dim mtx1 As XFORM, mtx2 As XFORM, c As Single, s As Single
    Dim iExpandOutsideForAngle As Long
    Dim iExpandOutsideForFigure As Long
    Dim iExpandOutsideForCurve As Long
    Dim iLng As Long
    Dim iLng2 As Long
    Dim iShift As Long
    Static sTheLastTimeWasExpanded As Boolean
    Dim iAuxExpand As Long
    
    If (mRotationDegrees > 0) Or (mFlipped <> seFlippedNo) Then
        iGMPrev = SetGraphicsMode(UserControl.hDC, GM_ADVANCED)
        ModifyWorldTransform UserControl.hDC, mtx1, MWT_IDENTITY
        If mRotationDegrees = 0 Then
            c = 1
            s = 0
        Else
            c = Cos(-mRotationDegrees / 360 * 2 * Pi)
            s = Sin(-mRotationDegrees / 360 * 2 * Pi)
        End If
        mtx1.eM11 = c: mtx1.eM12 = s: mtx1.eM21 = -s: mtx1.eM22 = c: mtx1.eDx = (UserControl.ScaleWidth - 1) / 2: mtx1.eDy = (UserControl.ScaleHeight - 1) / 2
        If mFlipped = seFlippedHorizontally Then
            mtx2.eM11 = -1: mtx2.eM22 = 1: mtx2.eDx = (UserControl.ScaleWidth - 1) / 2: mtx2.eDy = -(UserControl.ScaleHeight - 1) / 2
        ElseIf mFlipped = seFlippedVertically Then
            mtx2.eM11 = 1: mtx2.eM22 = -1: mtx2.eDx = -(UserControl.ScaleWidth - 1) / 2: mtx2.eDy = (UserControl.ScaleHeight - 1) / 2
        ElseIf mFlipped = seFlippedBoth Then
            mtx2.eM11 = -1: mtx2.eM22 = -1: mtx2.eDx = (UserControl.ScaleWidth - 1) / 2: mtx2.eDy = (UserControl.ScaleHeight - 1) / 2
        Else
            mtx2.eM11 = 1: mtx2.eM22 = 1: mtx2.eDx = -(UserControl.ScaleWidth - 1) / 2: mtx2.eDy = -(UserControl.ScaleHeight - 1) / 2
        End If
    End If
        
    iExpandForPen = mBorderWidth / 2
    If mCurvingFactor > 0 Then
        iAuxExpand = UserControl.ScaleWidth / UserControl.ScaleHeight * (1 + mCurvingFactor / 10)
        If iAuxExpand > iExpandForPen Then
            iExpandForPen = iAuxExpand
        End If
    End If
    If (mShape > seShapeRoundedSquare) Then
        If UserControl.ScaleWidth > UserControl.ScaleHeight Then
            iAuxExpand = UserControl.ScaleWidth / UserControl.ScaleHeight * mBorderWidth
        Else
            iAuxExpand = UserControl.ScaleHeight / UserControl.ScaleWidth * mBorderWidth
        End If
        If (mShape = seShapeStar) Or (mShape = seShapeJaggedStar) Then
            iExpandForPen = iExpandForPen * mVertices / 6
        End If
        If iAuxExpand > iExpandForPen Then
            iExpandForPen = iAuxExpand
        End If
    End If
    If ShapeHasShift(mShape) Then
        If UserControl.ScaleWidth > UserControl.ScaleHeight Then
            iShift = mShift * UserControl.ScaleWidth / 100
        Else
            iShift = mShift * UserControl.ScaleHeight / 100
        End If
        iLng = UserControl.ScaleWidth / 2 - iShift * 1.3
        If iLng < 0 Then
            iExpandOutsideForFigure = Abs(iLng)
        Else
            iLng = UserControl.ScaleWidth / 2 + iShift - UserControl.ScaleWidth
            If iLng > 0 Then
                iExpandOutsideForFigure = iLng
            Else
                iLng = UserControl.ScaleWidth / 2 + iShift
                If iLng < 0 Then
                    iExpandOutsideForFigure = Abs(iLng)
                End If
                If iExpandOutsideForFigure < Abs(iShift) Then
                    iExpandOutsideForFigure = Abs(iShift)
                End If
            End If
        End If
    End If
    If mCurvingFactor <> 0 Then
        iExpandOutsideForCurve = (UserControl.Width ^ 2 + UserControl.Height ^ 2) ^ 0.5 * 1.2
    End If
    If (mRotationDegrees <> 0) Then
        If (mShape <> seShapeCircle) And (mShape <> seShapeStar) And (mShape <> seShapeJaggedStar) Then
            iLng = Abs((UserControl.Width - UserControl.Height) / 2)
            iLng2 = (UserControl.Width ^ 2 + UserControl.Height ^ 2) ^ 0.5
            If iLng < iLng2 Then
                iExpandOutsideForAngle = iLng
            Else
                iExpandOutsideForAngle = iLng2
            End If
        End If
    End If
    
    mDrawingOutsideUC = False
    hRgn = CreateRectRgn(0, 0, 0, 0)
    If GetClipRgn(UserControl.hDC, hRgn) = 0& Then  ' hDc is one passed to Paint
        DeleteObject hRgn: hRgn = 0
    Else
        GetRgnBox hRgn, rgnRect             ' get its bounds & adjust our region accordingly (i.e.,expand 1 pixel)
        'Debug.Print "Rect: "; rgnRect.Left, rgnRect.Top, rgnRect.Right, rgnRect.Bottom
        
        If (iExpandForPen <> 0) Or (iExpandOutsideForAngle <> 0) Or (iExpandOutsideForFigure <> 0) Or sTheLastTimeWasExpanded Then
            hRgnExpand = CreateRectRgn(rgnRect.Left - iExpandForPen - iExpandOutsideForAngle - iExpandOutsideForFigure, rgnRect.top - iExpandForPen - iExpandOutsideForAngle - iExpandOutsideForFigure, rgnRect.Right + iExpandForPen + iExpandOutsideForAngle + iExpandOutsideForFigure, rgnRect.Bottom + iExpandForPen + iExpandOutsideForAngle + iExpandOutsideForFigure)
        
            SelectClipRgn UserControl.hDC, hRgnExpand
            DeleteObject hRgnExpand
            If Not tmrPainting.enabled Then
                If (UserControl.ContainerHwnd <> 0) And mSubclassed Then
                    PostMessage UserControl.ContainerHwnd, WM_INVALIDATE, 0&, 0&
                End If
            End If
            mDrawingOutsideUC = True
        End If
    End If
    sTheLastTimeWasExpanded = (iExpandForPen <> 0) Or (iExpandOutsideForAngle <> 0) Or (iExpandOutsideForFigure <> 0)
    tmrPainting.enabled = False
    tmrPainting.enabled = True
    
    If (mRotationDegrees > 0) Or (mFlipped <> seFlippedNo) Then
        SetWorldTransform UserControl.hDC, mtx1
        ModifyWorldTransform UserControl.hDC, mtx2, MWT_LEFTMULTIPLY
    End If
    
    Draw
    
    If hRgnExpand <> 0 Then SelectClipRgn UserControl.hDC, hRgn  ' restore original clip region
    If hRgn <> 0 Then DeleteObject hRgn
    
    If (mRotationDegrees > 0) Or (mFlipped <> seFlippedNo) Then
        SetGraphicsMode UserControl.hDC, iGMPrev
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mBackColor = PropBag.ReadProperty("BackColor", mdef_BackColor)
    mBackStyle = PropBag.ReadProperty("BackStyle", mdef_BackStyle)
    mBorderColor = PropBag.ReadProperty("BorderColor", mdef_BorderColor)
    mShape = PropBag.ReadProperty("Shape", mdef_Shape)
    mFillColor = PropBag.ReadProperty("FillColor", mdef_FillColor)
    mFillStyle = PropBag.ReadProperty("FillStyle", mdef_FillStyle)
    mBorderStyle = PropBag.ReadProperty("BorderStyle", mdef_BorderStyle)
    mBorderWidth = PropBag.ReadProperty("BorderWidth", mdef_BorderWidth)
    mClickable = PropBag.ReadProperty("Clickable", mdef_Clickable)
    mQuality = PropBag.ReadProperty("Quality", mdef_Quality)
    mRotationDegrees = PropBag.ReadProperty("RotationDegrees", mdef_RotationDegrees)
    mOpacity = PropBag.ReadProperty("Opacity", mdef_Opacity)
    mShift = PropBag.ReadProperty("Shift", mdef_Shift)
    mShiftPutAutomatically = PropBag.ReadProperty("ShiftPutAutomatically", 0)
    mVertices = PropBag.ReadProperty("Vertices", mdef_Vertices)
    mCurvingFactor = PropBag.ReadProperty("CurvingFactor", mdef_CurvingFactor)
    mFlipped = PropBag.ReadProperty("Flipped", mdef_Flipped)
    mMousePointer = PropBag.ReadProperty("MousePointer", mdef_MousePointer)
    Set mMouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    mStyle3D = PropBag.ReadProperty("Style3D", mdef_Style3D)
    mStyle3DEffect = PropBag.ReadProperty("Style3DEffect", mdef_Style3DEffect)
    mUseSubclassing = PropBag.ReadProperty("UseSubclassing", mdef_UseSubclassing)
    
    UserControl.MousePointer = mMousePointer
    Set UserControl.MouseIcon = mMouseIcon
    
    On Error Resume Next
    mUserMode = Ambient.UserMode
    On Error GoTo 0
    SetCurvingFactor2
    pvSubclass
End Sub

Private Sub UserControl_Terminate()
    pvUnsubclass
   On Error Resume Next
    If (mBorderWidth > 1) Or (mRotationDegrees > 0) Then InvalidateRectAsNull UserControl.ContainerHwnd, 0&, 1& ' paint the container when the control is deleted if the BorderWidth is greater than 1 or the control is rotated (if it painted outside its bounds)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", mBackColor, mdef_BackColor
    PropBag.WriteProperty "BackStyle", mBackStyle, mdef_BackStyle
    PropBag.WriteProperty "BorderColor", mBorderColor, mdef_BorderColor
    PropBag.WriteProperty "Shape", mShape, mdef_Shape
    PropBag.WriteProperty "FillColor", mFillColor, mdef_FillColor
    PropBag.WriteProperty "FillStyle", mFillStyle, mdef_FillStyle
    PropBag.WriteProperty "BorderStyle", mBorderStyle, mdef_BorderStyle
    PropBag.WriteProperty "BorderWidth", mBorderWidth, mdef_BorderWidth
    PropBag.WriteProperty "Clickable", mClickable, mdef_Clickable
    PropBag.WriteProperty "Quality", mQuality, mdef_Quality
    PropBag.WriteProperty "RotationDegrees", mRotationDegrees, mdef_RotationDegrees
    PropBag.WriteProperty "Opacity", mOpacity, mdef_Opacity
    PropBag.WriteProperty "Shift", mShift, mdef_Shift
    PropBag.WriteProperty "ShiftPutAutomatically", mShiftPutAutomatically, 0
    PropBag.WriteProperty "Vertices", mVertices, mdef_Vertices
    PropBag.WriteProperty "CurvingFactor", mCurvingFactor, mdef_CurvingFactor
    PropBag.WriteProperty "Flipped", mFlipped, mdef_Flipped
    PropBag.WriteProperty "MousePointer", mMousePointer, mdef_MousePointer
    PropBag.WriteProperty "MouseIcon", mMouseIcon, Nothing
    PropBag.WriteProperty "Style3D", mStyle3D, mdef_Style3D
    PropBag.WriteProperty "Style3DEffect", mStyle3DEffect, mdef_Style3DEffect
    PropBag.WriteProperty "UseSubclassing", mUseSubclassing, mdef_UseSubclassing
End Sub


Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BorderColor.VB_UserMemId = -503
    BorderColor = M2000color(mBorderColor)
End Property

Public Property Let BorderColor(ByVal nValue As OLE_COLOR)
    If nValue <> mBorderColor Then
        mBorderColor = mycolor(nValue)
        Me.Refresh
        PropertyChanged "BorderColor"
    End If
End Property


Public Property Get Shape() As SEShapeConstants
Attribute Shape.VB_Description = "Returns/sets a value indicating the appearance of a control."
Attribute Shape.VB_ProcData.VB_Invoke_Property = ";Apariencia"
    Shape = mShape
End Property

Public Property Let Shape(ByVal nValue As SEShapeConstants)
    If nValue <> mShape Then
        If (nValue < seShapeRectangle) Or (nValue > seShapePie) Then Err.Raise 380, Typename(Me): Exit Property
        If ShapeHasShift(mShape) Then
            If mShift = mShiftPutAutomatically Then
                mShift = 0
                mShiftPutAutomatically = 0
            End If
        End If
        mShape = nValue
        If ShapeHasShift(mShape) Then
            If mShift = 0 Then
                mShift = 20
                mShiftPutAutomatically = mShift
            End If
        End If
        Me.Refresh
        PropertyChanged "Shape"
    End If
End Property

Private Function ShapeHasShift(nShape As SEShapeConstants) As Boolean
    Select Case nShape
        Case seShapeTriangleScalene, seShapeKite, seShapeDiamond, seShapeTrapezoid, seShapeParalellogram, seShapeArrow, seShapeStar, seShapeJaggedStar, seShapeTalk, seShapeCrescent
            ShapeHasShift = True
    End Select
End Function

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BackColor.VB_UserMemId = -501
    BackColor = M2000color(mBackColor)
End Property

Public Property Let BackColor(ByVal nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        mBackColor = mycolor(nValue)
        Me.Refresh
        PropertyChanged "BackColor"
    End If
End Property


Public Property Get BackStyle() As SEBackStyleConstants
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BackStyle.VB_UserMemId = -502
    BackStyle = mBackStyle
End Property

Public Property Let BackStyle(ByVal nValue As SEBackStyleConstants)
    If nValue <> mBackStyle Then
        mBackStyle = nValue
        Me.Refresh
        PropertyChanged "BackStyle"
    End If
End Property


Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
Attribute FillColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute FillColor.VB_UserMemId = -510
    FillColor = M2000color(mFillColor)
End Property

Public Property Let FillColor(ByVal nValue As OLE_COLOR)
    If nValue <> mFillColor Then
        mFillColor = mycolor(nValue)
        Me.Refresh
        PropertyChanged "FillColor"
    End If
End Property


Public Property Get FillStyle() As SEFillStyle2Constants
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
Attribute FillStyle.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute FillStyle.VB_UserMemId = -511
    FillStyle = mFillStyle
End Property

Public Property Let FillStyle(ByVal nValue As SEFillStyle2Constants)
    If nValue <> mFillStyle Then
        If (nValue < seFSSolid) Or (nValue > seFSTransparent) Then Err.Raise 380, Typename(Me): Exit Property
        mFillStyle = nValue
        Me.Refresh
        PropertyChanged "FillStyle"
    End If
End Property


Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal nValue As BorderStyleConstants)
    If nValue <> mBorderStyle Then
        If (nValue < vbTransparent) Or (nValue > vbBSInsideSolid) Then Err.Raise 380, Typename(Me): Exit Property
        mBorderStyle = nValue
        Me.Refresh
        PropertyChanged "BorderStyle"
    End If
End Property


Public Property Get BorderWidth() As Long
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BorderWidth.VB_UserMemId = -505
    BorderWidth = mBorderWidth
End Property

Public Property Let BorderWidth(ByVal nValue As Long)
    If nValue < 1 Then
        nValue = 1
    End If
    If nValue <> mBorderWidth Then
        mBorderWidth = nValue
        Me.Refresh
        PropertyChanged "BorderWidth"
    End If
End Property


Public Property Get Clickable() As Boolean
Attribute Clickable.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
    Clickable = mClickable
End Property

Public Property Let Clickable(ByVal nValue As Boolean)
    If nValue <> mClickable Then
        mClickable = nValue
        Me.Refresh
        PropertyChanged "Clickable"
    End If
End Property


Public Property Get Quality() As SEQualityConstants
Attribute Quality.VB_ProcData.VB_Invoke_Property = ";Apariencia"
    Quality = mQuality
End Property

Public Property Let Quality(ByVal nValue As SEQualityConstants)
    If nValue <> mQuality Then
        mQuality = nValue
        Me.Refresh
        PropertyChanged "Quality"
    End If
End Property


Public Property Get RotationDegrees() As Single
    RotationDegrees = mRotationDegrees
End Property

Public Property Let RotationDegrees(ByVal nValue As Single)
    Dim iFraction As Single
    
    If nValue <> mRotationDegrees Then
        iFraction = nValue - Round(nValue)
        nValue = nValue Mod 360
        If nValue < 0 Then nValue = nValue + 360
        nValue = nValue + iFraction
        If nValue >= 360 Then
            nValue = nValue - 360
        ElseIf nValue < 0 Then
            nValue = nValue + 360
        End If
        If nValue <> mRotationDegrees Then
            mRotationDegrees = nValue
            Me.Refresh
            If mDrawingOutsideUC Then
                If (UserControl.ContainerHwnd <> 0) And mSubclassed Then
                    PostMessage UserControl.ContainerHwnd, WM_INVALIDATE, 0&, 0&
                End If
            End If
            PropertyChanged "RotationDegrees"
        End If
    End If
End Property


Public Property Get Opacity() As Single
    Opacity = mOpacity
End Property

Public Property Let Opacity(ByVal nValue As Single)
    If nValue <> mOpacity Then
        If nValue > 100 Then
            nValue = 100
        ElseIf nValue < 0 Then
            nValue = 0
        End If
        If nValue <> mOpacity Then
            mOpacity = nValue
            Me.Refresh
            If mDrawingOutsideUC Then
                If (UserControl.ContainerHwnd <> 0) And mSubclassed Then
                    PostMessage UserControl.ContainerHwnd, WM_INVALIDATE, 0&, 0&
                End If
            End If
            PropertyChanged "Opacity"
        End If
    End If
End Property


Public Property Get shift() As Single
    shift = mShift
End Property

Public Property Let shift(ByVal nValue As Single)
    If nValue <> mShift Then
        mShift = nValue
        Me.Refresh
        PropertyChanged "Shift"
    End If
End Property


Public Property Get Vertices() As Integer
    Vertices = mVertices
End Property

Public Property Let Vertices(ByVal nValue As Integer)
    If nValue <> mVertices Then
        mVertices = nValue
        If mVertices < 2 Then mVertices = 2
        If mVertices > 100 Then mVertices = 100
        Me.Refresh
        PropertyChanged "Vertices"
    End If
End Property


Public Property Get CurvingFactor() As Integer
    CurvingFactor = mCurvingFactor
End Property

Public Property Let CurvingFactor(ByVal nValue As Integer)
    If nValue <> mCurvingFactor Then
        mCurvingFactor = nValue
        If mCurvingFactor < -100 Then mCurvingFactor = -100
        If mCurvingFactor > 100 Then mCurvingFactor = 100
        SetCurvingFactor2
        Me.Refresh
        If mDrawingOutsideUC Then
            If (UserControl.ContainerHwnd <> 0) And mSubclassed Then
                PostMessage UserControl.ContainerHwnd, WM_INVALIDATE, 0&, 0&
            End If
        End If
        PropertyChanged "CurvingFactor"
    End If
End Property


Public Property Get Flipped() As SEFlippedConstants
    Flipped = mFlipped
End Property

Public Property Let Flipped(ByVal nValue As SEFlippedConstants)
    If nValue <> mFlipped Then
        mFlipped = nValue
        Me.Refresh
        PropertyChanged "Flipped"
    End If
End Property


Public Property Get MousePointer() As VBRUN.MousePointerConstants
    MousePointer = mMousePointer
End Property

Public Property Let MousePointer(ByVal nValue As VBRUN.MousePointerConstants)
    If nValue <> mMousePointer Then
        mMousePointer = nValue
        UserControl.MousePointer = mMousePointer
        PropertyChanged "MousePointer"
    End If
End Property


Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = mMouseIcon
End Property

Public Property Let MouseIcon(ByVal nValue As StdPicture)
    Set MouseIcon = nValue
End Property

Public Property Set MouseIcon(ByVal nValue As StdPicture)
    If Not nValue Is mMouseIcon Then
        Set mMouseIcon = nValue
        Set UserControl.MouseIcon = mMouseIcon
        PropertyChanged "MouseIcon"
    End If
End Property


Public Property Get Style3D() As SEStyle3DConstants
    Style3D = mStyle3D
End Property

Public Property Let Style3D(ByVal nValue As SEStyle3DConstants)
    If nValue <> mStyle3D Then
        If (nValue < seStyle3DNone) Or (nValue > seStyle3DBoth) Then Err.Raise 380, Typename(Me): Exit Property
        mStyle3D = nValue
        Me.Refresh
        PropertyChanged "Style3D"
    End If
End Property


Public Property Get Style3DEffect() As SEStyle3DEffectConstants
    Style3DEffect = mStyle3DEffect
End Property

Public Property Let Style3DEffect(ByVal nValue As SEStyle3DEffectConstants)
    If nValue <> mStyle3DEffect Then
        If (nValue < seStyle3EffectAuto) Or (nValue > seStyle3EffectGem) Then Err.Raise 380, Typename(Me): Exit Property
        mStyle3DEffect = nValue
        Me.Refresh
        PropertyChanged "Style3DEffect"
    End If
End Property


Public Property Get UseSubclassing() As SESubclassingConstants
    UseSubclassing = mUseSubclassing
End Property

Public Property Let UseSubclassing(ByVal nValue As SESubclassingConstants)
    Dim iMessage As T_MSG
    
    If nValue <> mUseSubclassing Then
        If Not mUserMode Then
            If nValue = seSCNo Then
                MsgBox "In most cases subclassing is not necessary (and without subclassing is safer for running in the IDE), but in some special cases when the figure needs to be painted outside the control bounds, it may experience glitches.", vbInformation
            End If
        End If
        mUseSubclassing = nValue
        If mUseSubclassing <> seSCNo Then
            If mSubclassed Then pvUnsubclass
            pvSubclass
            If mSubclassed Then
                If mDrawingOutsideUC Then
                    If (UserControl.ContainerHwnd <> 0) And mSubclassed Then
                        PostMessage UserControl.ContainerHwnd, WM_INVALIDATE, 0&, 0&
                    End If
                End If
            Else
                If UserControl.ContainerHwnd <> 0 Then
                    PeekMessage iMessage, UserControl.ContainerHwnd, WM_INVALIDATE, WM_INVALIDATE, PM_REMOVE  ' remove posted message, if any
                End If
            End If
        Else
            If mSubclassed Then pvUnsubclass
            If UserControl.ContainerHwnd <> 0 Then
                PeekMessage iMessage, UserControl.ContainerHwnd, WM_INVALIDATE, WM_INVALIDATE, PM_REMOVE  ' remove posted message, if any
            End If
        End If
        PropertyChanged "UseSubclassing"
    End If
End Property


Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property
    
Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
    UserControl.Refresh
End Sub

Private Sub Draw()
    Dim iDiameter As Long
    Dim iGraphics As Long
    Dim iFillColor As Long
    Dim iFilled As Boolean
    Dim iheight As Long
    Dim iRoundSize As Long
    Dim iPts() As POINTL
    Dim iEdge As Long
    Dim iUCWidth As Long
    Dim iUCHeight As Long
    Dim iLng As Long
    Dim c As Long
    Dim iPts2() As POINTL
    Dim iPts3() As POINTL
    Dim iShift As Long
    Dim iHalfBorderWidth As Long
    
    If GdipCreateFromHDC(UserControl.hDC, iGraphics) = 0 Then
        
        If mFillStyle = seFSSolid Then
            iFilled = True
            iFillColor = mFillColor
        ElseIf mBackStyle = seOpaque Then
            iFilled = True
            iFillColor = mBackColor
        Else
            iFilled = False
        End If
        
        iUCWidth = UserControl.ScaleWidth - 1
        iUCHeight = UserControl.ScaleHeight - 1
        
        If mShape = seShapeOval Then
            If iFilled Then
                FillEllipse iGraphics, iFillColor, 0, 0, iUCWidth, iUCHeight
            End If
            If mBorderStyle <> vbTransparent Then
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, 0, 0, iUCWidth, iUCHeight
            End If
        ElseIf mShape = seShapeCircle Then
            If iUCWidth < iUCHeight Then
                iDiameter = iUCWidth
            Else
                iDiameter = iUCHeight
            End If
            If iFilled Then
                FillEllipse iGraphics, iFillColor, iUCWidth / 2 - iDiameter / 2, iUCHeight / 2 - iDiameter / 2, iDiameter, iDiameter
            End If
            If mBorderStyle <> vbTransparent Then
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth / 2 - iDiameter / 2, iUCHeight / 2 - iDiameter / 2, iDiameter, iDiameter
            End If
        ElseIf mShape = seShapeSquare Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iheight = iUCWidth
            Else
                iheight = iUCHeight
            End If
            
            ReDim iPts(3)
            iPts(0).X = iUCWidth / 2 - iheight / 2
            iPts(0).Y = iUCHeight / 2 - iheight / 2
            iPts(1).X = iUCWidth / 2 - iheight / 2
            iPts(1).Y = iUCHeight / 2 - iheight / 2 + iheight
            iPts(2).X = iUCWidth / 2 - iheight / 2 + iheight
            iPts(2).Y = iUCHeight / 2 - iheight / 2 + iheight
            iPts(3).X = iUCWidth / 2 - iheight / 2 + iheight
            iPts(3).Y = iUCHeight / 2 - iheight / 2
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(1).Y = iPts(1).Y - iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeRoundedRectangle Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iRoundSize = UserControl.ScaleWidth * 0.125
            Else
                iRoundSize = UserControl.ScaleHeight * 0.125
            End If
            If iFilled Then
                FillRoundRect iGraphics, iFillColor, 0, 0, iUCWidth, iUCHeight, iRoundSize
            End If
            If mBorderStyle <> vbTransparent Then
                DrawRoundRect iGraphics, mBorderColor, mBorderWidth, 0, 0, iUCWidth, iUCHeight, iRoundSize
            End If
        ElseIf mShape = seShapeRoundedSquare Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iheight = UserControl.ScaleWidth
            Else
                iheight = UserControl.ScaleHeight
            End If
            iRoundSize = iheight * 0.125
            If iFilled Then
                FillRoundRect iGraphics, iFillColor, UserControl.ScaleWidth / 2 - iheight / 2, UserControl.ScaleHeight / 2 - iheight / 2, iheight - 1, iheight - 1, iRoundSize
            End If
            If mBorderStyle <> vbTransparent Then
                DrawRoundRect iGraphics, mBorderColor, mBorderWidth, UserControl.ScaleWidth / 2 - iheight / 2, UserControl.ScaleHeight / 2 - iheight / 2, iheight - 1, iheight - 1, iRoundSize
            End If
        ElseIf mShape = seShapeTriangleEquilateral Then
            ReDim iPts(2)
            
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iEdge = UserControl.ScaleWidth
            Else
                iEdge = UserControl.ScaleHeight
            End If
            
'            iEdge = iHeight * 2 / 3 ^ 0.5
            iheight = (3 ^ 0.5 * iEdge) / 2
            iPts(0).X = iUCWidth / 2
            iPts(0).Y = iUCHeight / 2 - iheight / 2
            iPts(1).X = iUCWidth / 2 - iEdge / 2
            iPts(1).Y = iUCHeight / 2 + iheight / 2
            iPts(2).X = iUCWidth / 2 + iEdge / 2
            iPts(2).Y = iUCHeight / 2 + iheight / 2
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth / 2
                iPts(2).X = iPts(2).X - iHalfBorderWidth / 2
            End If
                
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
            
        ElseIf mShape = seShapeTriangleIsosceles Then
            ReDim iPts(2)
            
            iPts(0).X = iUCWidth / 2
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth / 2
                iPts(2).X = iPts(2).X - iHalfBorderWidth / 2
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeTriangleScalene Then
            ReDim iPts(2)
            
            iPts(0).X = iUCWidth / 2 - (iUCWidth / 100 * mShift)
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth / 2
                iPts(2).X = iPts(2).X - iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeTriangleRight Then
            ReDim iPts(2)
            
            iPts(0).X = 0
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth / 2
                iPts(2).X = iPts(2).X - iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeRhombus Then
            ReDim iPts(3)
            
            iPts(0).X = iUCWidth / 2
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight / 2
            iPts(2).X = iUCWidth / 2
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = iUCHeight / 2
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
            End If
             
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeKite Then
            ReDim iPts(3)
            
            iLng = iUCHeight / 2 - (iUCHeight / 100 * mShift / 20 * 15)
            
            iPts(0).X = iUCWidth / 2
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iLng
            iPts(2).X = iUCWidth / 2
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = iLng
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
            End If
             
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeDiamond Then
            ReDim iPts(4)
            
            iLng = iUCHeight / 2 - (iUCHeight / 100 * mShift / 20 * 15)
            
            iPts(0).X = iUCWidth * 0.33
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iLng
            iPts(2).X = iUCWidth / 2
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = iLng
            iPts(4).X = iUCWidth * 0.66
            iPts(4).Y = 0
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(4).Y = iPts(4).Y + iHalfBorderWidth
            End If
             
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeTrapezoid Then
            ReDim iPts(3)
            
            iLng = (iUCWidth / 100 * mShift)
            If iLng > iUCWidth / 2 Then
                iLng = iUCWidth / 2
            End If
            iPts(0).X = iLng
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth - iLng
            iPts(3).Y = 0
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(1).Y = iPts(1).Y - iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeParalellogram Then
            ReDim iPts(3)
            
            iLng = (iUCWidth / 100 * mShift)
            If iLng > iUCWidth Then
                iLng = iUCWidth
            End If
            iPts(0).X = iLng
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth - iLng
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = 0
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(1).Y = iPts(1).Y - iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeSemicircle Then
            If iFilled Then
                FillSemicircle iGraphics, iFillColor, 0, 0, iUCWidth, iUCHeight
            End If
            If mBorderStyle <> vbTransparent Then
                DrawSemicircle iGraphics, mBorderColor, mBorderWidth, 0, 0, iUCWidth, iUCHeight
            End If
        ElseIf mShape = seShapeRegularPolygon Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iheight = UserControl.ScaleWidth
            Else
                iheight = UserControl.ScaleHeight
            End If
            
            ReDim iPts(mVertices - 1)
            
            If mBorderStyle = vbBSInsideSolid Then
                iheight = iheight - mBorderWidth / 2
                If iheight < mBorderWidth Then iheight = mBorderWidth
            End If
            
            For c = 0 To mVertices - 1
                iPts(c).X = (iheight / 2) * Cos(2 * Pi * (c + 1) / mVertices) + iUCWidth / 2
                iPts(c).Y = (iheight / 2) * Sin(2 * Pi * (c + 1) / mVertices) + iUCHeight / 2
            Next c
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf (mShape = seShapeStar) Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iheight = UserControl.ScaleWidth
            Else
                iheight = UserControl.ScaleHeight
            End If
            
            ReDim iPts(mVertices * 2 - 1)
            
            If mBorderStyle = vbBSInsideSolid Then
                iheight = iheight - mBorderWidth / 2
                If iheight < mBorderWidth Then iheight = mBorderWidth
            End If
            
            For c = 0 To mVertices * 2 - 1
                iPts(c).X = (iheight / 2) * Cos(2 * Pi * (c + 1) / (mVertices * 2)) + iUCWidth / 2
                iPts(c).Y = (iheight / 2) * Sin(2 * Pi * (c + 1) / (mVertices * 2)) + iUCHeight / 2
            Next c
            
            ReDim iPts2(mVertices - 1)
            iShift = (iheight / 100 * mShift / 3) + 10
            
            For c = 0 To mVertices - 1
                iPts2(c).X = (iheight / 2 - iShift) * Cos(2 * Pi * (c + 1) / mVertices) + iUCWidth / 2
                iPts2(c).Y = (iheight / 2 - iShift) * Sin(2 * Pi * (c + 1) / mVertices) + iUCHeight / 2
            Next c
            
            ReDim iPts3(mVertices * 2 - 1)
            For c = 0 To mVertices * 2 - 1
                If c Mod 2 = 0 Then
                    iPts3(c).X = iPts2(c / 2).X
                    iPts3(c).Y = iPts2(c / 2).Y
                Else
                    iPts3(c).X = iPts((c + 1) Mod (UBound(iPts) + 1)).X
                    iPts3(c).Y = iPts((c + 1) Mod (UBound(iPts) + 1)).Y
                End If
            Next c
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts3, FillModeWinding
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts3
            End If
        ElseIf (mShape = seShapeJaggedStar) Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iheight = UserControl.ScaleWidth
            Else
                iheight = UserControl.ScaleHeight
            End If
            
            ReDim iPts(mVertices * 2 - 1)
            
            If mBorderStyle = vbBSInsideSolid Then
                iheight = iheight - mBorderWidth / 2
                If iheight < mBorderWidth Then iheight = mBorderWidth
            End If
            
            For c = 0 To mVertices * 2 - 1
                iPts(c).X = (iheight / 2) * Cos(2 * Pi * (c + 1) / (mVertices * 2)) + iUCWidth / 2
                iPts(c).Y = (iheight / 2) * Sin(2 * Pi * (c + 1) / (mVertices * 2)) + iUCHeight / 2
            Next c
            
            ReDim iPts2(mVertices - 1)
            iShift = (iheight / 100 * mShift / 3) + 10
            
            For c = 0 To mVertices - 1
                iPts2(c).X = (iheight / 2 - iShift) * Cos(2 * Pi * (c + 1) / mVertices) + iUCWidth / 2
                iPts2(c).Y = (iheight / 2 - iShift) * Sin(2 * Pi * (c + 1) / mVertices) + iUCHeight / 2
            Next c
            
            ReDim iPts3(mVertices * 2 - 1)
            For c = 0 To mVertices * 2 - 1
                If c Mod 2 = 0 Then
                    iPts3(c).X = iPts2(c / 2).X
                    iPts3(c).Y = iPts2(c / 2).Y
                Else
                    iPts3(c).X = iPts(c).X
                    iPts3(c).Y = iPts(c).Y
                End If
            Next c
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts3, FillModeWinding
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts3
            End If
        ElseIf mShape = seShapeHeart Then
            ReDim iPts(13)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            iPts(0).X = iUCWidth * 0.5
            iPts(0).Y = iUCHeight * 0.19
            iPts(1).X = iUCWidth * 0.35
            iPts(1).Y = iUCHeight * 0.04
            iPts(2).X = iUCWidth * 0.15
            iPts(2).Y = iUCHeight * 0.03
            iPts(3).X = iUCWidth * 0.005
            iPts(3).Y = iUCHeight * 0.2
            iPts(4).X = iUCWidth * 0.02
            iPts(4).Y = iUCHeight * 0.45
            iPts(5).X = iUCWidth * 0.2 ''''
            iPts(5).Y = iUCHeight * 0.7 '''
            iPts(6).X = iUCWidth * 0.49
            iPts(6).Y = iUCHeight * 0.99
            iPts(7).X = iUCWidth * 0.51
            iPts(7).Y = iUCHeight * 0.99
            iPts(8).X = iUCWidth * 0.8 '''
            iPts(8).Y = iUCHeight * 0.7 '''
            iPts(9).X = iUCWidth * 0.98
            iPts(9).Y = iUCHeight * 0.45
            iPts(10).X = iUCWidth * 0.995
            iPts(10).Y = iUCHeight * 0.2
            iPts(11).X = iUCWidth * 0.85
            iPts(11).Y = iUCHeight * 0.03
            iPts(12).X = iUCWidth * 0.65
            iPts(12).Y = iUCHeight * 0.04
            iPts(13).X = iUCWidth * 0.5
            iPts(13).Y = iUCHeight * 0.19
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
                
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.45
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.45
            End If
        ElseIf mShape = seShapeArrow Then
            ReDim iPts(6)
            
            iLng = iUCWidth * (0.75 - mShift / 100 * 0.75 / 20 * 15)
            If iLng > iUCWidth * 0.95 Then iLng = iUCWidth * 0.95
            
            iPts(0).X = iUCWidth * 0.005
            iPts(0).Y = iUCHeight * 0.25
            iPts(1).X = iLng
            iPts(1).Y = iUCHeight * 0.25
            iPts(2).X = iLng
            iPts(2).Y = iUCHeight * 0.005
            iPts(3).X = iUCWidth * 0.995
            iPts(3).Y = iUCHeight / 2
            iPts(4).X = iLng
            iPts(4).Y = iUCHeight * 0.995
            iPts(5).X = iLng
            iPts(5).Y = iUCHeight * 0.75
            iPts(6).X = iUCWidth * 0.005
            iPts(6).Y = iUCHeight * 0.75
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(2).Y = iPts(2).Y + iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(4).Y = iPts(4).Y - iHalfBorderWidth
                iPts(6).X = iPts(6).X + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeCrescent Then
            
            ReDim iPts(11)
            iLng = iUCWidth * (0.2 + mShift / 50)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            ' top
            iPts(0).X = iUCWidth * 0.25 + iLng
            iPts(0).Y = iUCHeight * 0.005
            iPts(1).X = iUCWidth * 0.245 + iLng * 0.52
            iPts(1).Y = iUCHeight * 0.04
            ' left
            iPts(2).X = iUCWidth * 0.24
            iPts(2).Y = iUCHeight * 0.2
            iPts(3).X = iUCWidth * 0.1
            iPts(3).Y = iUCHeight * 0.5
            iPts(4).X = iUCWidth * 0.24
            iPts(4).Y = iUCHeight * 0.8
            ' bottom
            iPts(5).X = iUCWidth * 0.245 + iLng * 0.52
            iPts(5).Y = iUCHeight * 0.96
            iPts(6).X = iUCWidth * 0.25 + iLng
            iPts(6).Y = iUCHeight * 0.995
            ' right
            iPts(7).X = iUCWidth * 0.25 + iLng * 0.72
            iPts(7).Y = iUCHeight * 0.92
            iPts(8).X = iUCWidth * 0.25 + iLng * 0.44
            iPts(8).Y = iUCHeight * 0.77
            iPts(9).X = iUCWidth * 0.25 + iLng * 0.3
            iPts(9).Y = iUCHeight * 0.5
            iPts(10).X = iUCWidth * 0.25 + iLng * 0.44
            iPts(10).Y = iUCHeight * 0.23
            iPts(11).X = iUCWidth * 0.25 + iLng * 0.72
            iPts(11).Y = iUCHeight * 0.08
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
        
'            UserControl.DrawWidth = 10
'            On Error Resume Next
'            For c = 0 To UBound(iPts)
'                UserControl.PSet (iPts(c).X, iPts(c).Y), vbRed
'            Next
'            On Error GoTo 0
        
        ElseIf mShape = seShapeDrop Then
            ReDim iPts(11)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            iPts(0).X = iUCWidth * 0.49
            iPts(0).Y = iUCHeight * 0.005
            iPts(1).X = iUCWidth * 0.25
            iPts(1).Y = iUCHeight * 0.23
            
            iPts(2).X = iUCWidth * 0.05
            iPts(2).Y = iUCHeight * 0.5
            iPts(3).X = iUCWidth * 0.05
            iPts(3).Y = iUCHeight * 0.75
            
            iPts(4).X = iUCWidth * 0.2
            iPts(4).Y = iUCHeight * 0.9
            iPts(5).X = iUCWidth * 0.4
            iPts(5).Y = iUCHeight * 0.98
            iPts(6).X = iUCWidth * 0.6
            iPts(6).Y = iUCHeight * 0.98
            iPts(7).X = iUCWidth * 0.8
            iPts(7).Y = iUCHeight * 0.9
            
            iPts(8).X = iUCWidth * 0.95
            iPts(8).Y = iUCHeight * 0.75
            iPts(9).X = iUCWidth * 0.95
            iPts(9).Y = iUCHeight * 0.5
            
            iPts(10).X = iUCWidth * 0.75
            iPts(10).Y = iUCHeight * 0.23
            iPts(11).X = iUCWidth * 0.51
            iPts(11).Y = iUCHeight * 0.005
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
        ElseIf mShape = seShapeEgg Then
            ReDim iPts(11)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            iPts(0).X = iUCWidth * 0.4
            iPts(0).Y = iUCHeight * 0.1
            iPts(1).X = iUCWidth * 0.2
            iPts(1).Y = iUCHeight * 0.26
            
            iPts(2).X = iUCWidth * 0.05
            iPts(2).Y = iUCHeight * 0.53
            iPts(3).X = iUCWidth * 0.05
            iPts(3).Y = iUCHeight * 0.75
            
            iPts(4).X = iUCWidth * 0.18
            iPts(4).Y = iUCHeight * 0.92
            iPts(5).X = iUCWidth * 0.4
            iPts(5).Y = iUCHeight * 0.99
            iPts(6).X = iUCWidth * 0.6
            iPts(6).Y = iUCHeight * 0.99
            iPts(7).X = iUCWidth * 0.82
            iPts(7).Y = iUCHeight * 0.92
            
            iPts(8).X = iUCWidth * 0.95
            iPts(8).Y = iUCHeight * 0.75
            iPts(9).X = iUCWidth * 0.95
            iPts(9).Y = iUCHeight * 0.53
            
            iPts(10).X = iUCWidth * 0.8
            iPts(10).Y = iUCHeight * 0.26
            iPts(11).X = iUCWidth * 0.6
            iPts(11).Y = iUCHeight * 0.1
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
            
'            UserControl.DrawWidth = 10
'            On Error Resume Next
'            For c = 0 To UBound(iPts)
'                UserControl.PSet (iPts(c).X, iPts(c).Y), IIf(c = 4, vbGreen, IIf(c = 13, vbBlue, vbRed))
'            Next
'            On Error GoTo 0

        ElseIf mShape = seShapeLocation Then
            Dim iUCWidthOrig As Long
            Dim iUCHeightOrig As Long
            
            iUCWidthOrig = iUCWidth
            iUCHeightOrig = iUCHeight
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            If iFilled Then
                ReDim iPts(24)
                
                ' start going from bottom middle to left
                iPts(0).X = iUCWidth * 0.49
                iPts(0).Y = iUCHeight * 0.98
                iPts(1).X = iUCWidth * 0.28
                iPts(1).Y = iUCHeight * 0.77
                ' outer left
                iPts(2).X = iUCWidth * 0.05
                iPts(2).Y = iUCHeight * 0.5
                iPts(3).X = iUCWidth * 0.05
                iPts(3).Y = iUCHeight * 0.25
                ' outer top
                iPts(4).X = iUCWidth * 0.23
                iPts(4).Y = iUCHeight * 0.097
                iPts(5).X = iUCWidth * 0.4
                iPts(5).Y = iUCHeight * 0.05
                iPts(6).X = iUCWidth * 0.6
                iPts(6).Y = iUCHeight * 0.05
                iPts(7).X = iUCWidth * 0.77
                iPts(7).Y = iUCHeight * 0.097
                ' outer right
                iPts(8).X = iUCWidth * 0.95
                iPts(8).Y = iUCHeight * 0.25
                iPts(9).X = iUCWidth * 0.95
                iPts(9).Y = iUCHeight * 0.5
                ' going from right to bottom
                iPts(10).X = iUCWidth * 0.72
                iPts(10).Y = iUCHeight * 0.77
                ' at the bottom
                iPts(11).X = iUCWidth * 0.51
                iPts(11).Y = iUCHeight * 0.98
                iPts(12).X = iUCWidth * 0.5
                iPts(12).Y = iUCHeight * 0.97
                ' go inside, bottom of circle
                iPts(13).X = iUCWidth * 0.5
                iPts(13).Y = iUCHeight * 0.641
                iPts(14).X = iUCWidth * 0.47
                iPts(14).Y = iUCHeight * 0.591
                ' inner right of circle
                iPts(15).X = iUCWidth * 0.65
                iPts(15).Y = iUCHeight * 0.52
                iPts(16).X = iUCWidth * 0.73
                iPts(16).Y = iUCHeight * 0.38
                ' inner top of circle
                iPts(17).X = iUCWidth * 0.62
                iPts(17).Y = iUCHeight * 0.23
                iPts(18).X = iUCWidth * 0.38
                iPts(18).Y = iUCHeight * 0.23
                ' inner left of circle
                iPts(19).X = iUCWidth * 0.26
                iPts(19).Y = iUCHeight * 0.38
                iPts(20).X = iUCWidth * 0.34
                iPts(20).Y = iUCHeight * 0.52
                ' again in bottom of circle
                iPts(21).X = iUCWidth * 0.48
                iPts(21).Y = iUCHeight * 0.581
                iPts(22).X = iUCWidth * 0.48
                iPts(22).Y = iUCHeight * 0.601
                iPts(23).X = iUCWidth * 0.5
                iPts(23).Y = iUCHeight * 0.641
                ' go to outer bottom (to join the start)
                iPts(24).X = iUCWidth * 0.5
                iPts(24).Y = iUCHeight * 0.945
                
                If mBorderStyle = vbBSInsideSolid Then
                    iHalfBorderWidth = mBorderWidth / 2
                    For c = 0 To UBound(iPts)
                        iPts(c).X = iPts(c).X + iHalfBorderWidth
                        iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                    Next
                End If
                
                FillClosedCurve iGraphics, iFillColor, iPts, 0.55, FillModeWinding
                If mBorderStyle <> vbTransparent Then
                    DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidthOrig / 2 - iUCWidthOrig * 0.47 / 2, iUCHeightOrig * 0.202, iUCWidthOrig * 0.47, iUCHeightOrig * 0.372
                End If
                If mBorderStyle <> vbTransparent Then
                    ReDim iPts(11)
                    
                    ' start going from bottom middle to left
                    iPts(0).X = iUCWidth * 0.49
                    iPts(0).Y = iUCHeight * 0.98
                    iPts(1).X = iUCWidth * 0.28
                    iPts(1).Y = iUCHeight * 0.77
                    ' outer left
                    iPts(2).X = iUCWidth * 0.05
                    iPts(2).Y = iUCHeight * 0.5
                    iPts(3).X = iUCWidth * 0.05
                    iPts(3).Y = iUCHeight * 0.25
                    ' outer top
                    iPts(4).X = iUCWidth * 0.23
                    iPts(4).Y = iUCHeight * 0.097
                    iPts(5).X = iUCWidth * 0.4
                    iPts(5).Y = iUCHeight * 0.05
                    iPts(6).X = iUCWidth * 0.6
                    iPts(6).Y = iUCHeight * 0.05
                    iPts(7).X = iUCWidth * 0.77
                    iPts(7).Y = iUCHeight * 0.097
                    ' outer right
                    iPts(8).X = iUCWidth * 0.95
                    iPts(8).Y = iUCHeight * 0.25
                    iPts(9).X = iUCWidth * 0.95
                    iPts(9).Y = iUCHeight * 0.5
                    ' going from right to bottom
                    iPts(10).X = iUCWidth * 0.72
                    iPts(10).Y = iUCHeight * 0.77
                    iPts(11).X = iUCWidth * 0.51
                    iPts(11).Y = iUCHeight * 0.98
                    
                    If mBorderStyle = vbBSInsideSolid Then
                        iHalfBorderWidth = mBorderWidth / 2
                        For c = 0 To UBound(iPts)
                            iPts(c).X = iPts(c).X + iHalfBorderWidth
                            iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                        Next
                    End If
                    
                    DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
                End If
            ElseIf mBorderStyle <> vbTransparent Then
                ReDim iPts(11)
                
                ' start going from bottom middle to left
                iPts(0).X = iUCWidth * 0.49
                iPts(0).Y = iUCHeight * 0.98
                iPts(1).X = iUCWidth * 0.28
                iPts(1).Y = iUCHeight * 0.77
                ' outer left
                iPts(2).X = iUCWidth * 0.05
                iPts(2).Y = iUCHeight * 0.5
                iPts(3).X = iUCWidth * 0.05
                iPts(3).Y = iUCHeight * 0.25
                ' outer top
                iPts(4).X = iUCWidth * 0.23
                iPts(4).Y = iUCHeight * 0.097
                iPts(5).X = iUCWidth * 0.4
                iPts(5).Y = iUCHeight * 0.05
                iPts(6).X = iUCWidth * 0.6
                iPts(6).Y = iUCHeight * 0.05
                iPts(7).X = iUCWidth * 0.77
                iPts(7).Y = iUCHeight * 0.097
                ' outer right
                iPts(8).X = iUCWidth * 0.95
                iPts(8).Y = iUCHeight * 0.25
                iPts(9).X = iUCWidth * 0.95
                iPts(9).Y = iUCHeight * 0.5
                ' going from right to bottom
                iPts(10).X = iUCWidth * 0.72
                iPts(10).Y = iUCHeight * 0.77
                iPts(11).X = iUCWidth * 0.51
                iPts(11).Y = iUCHeight * 0.98
                
                If mBorderStyle = vbBSInsideSolid Then
                    iHalfBorderWidth = mBorderWidth / 2
                    For c = 0 To UBound(iPts)
                        iPts(c).X = iPts(c).X + iHalfBorderWidth
                        iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                    Next
                End If
                
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
                
                'DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth / 2 - iUCWidth * 0.47 / 2, iUCHeight * 0.205, iUCWidth * 0.47, iUCHeight * 0.365
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth / 2 - iUCWidth * 0.47 / 2, iUCHeight * 0.202, iUCWidth * 0.47, iUCHeight * 0.372
            End If
        ElseIf mShape = seShapeSpeaker Then
            ReDim iPts(5)
            
            iPts(0).X = 0
            iPts(0).Y = iUCHeight * 0.28
            iPts(1).X = iUCWidth * 0.37
            iPts(1).Y = iUCHeight * 0.28
            iPts(2).X = iUCWidth
            iPts(2).Y = 0
            iPts(3).X = iUCWidth
            iPts(3).Y = iUCHeight
            iPts(4).X = iUCWidth * 0.37
            iPts(4).Y = iUCHeight * 0.72
            iPts(5).X = 0
            iPts(5).Y = iUCHeight * 0.72
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y + iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y - iHalfBorderWidth
                iPts(5).X = iPts(5).X + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeCloud Then
            ReDim iPts(19)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            ' bottom, starting at the middle and going left
            iPts(0).X = iUCWidth * 0.49
            iPts(0).Y = iUCHeight * 0.995
            iPts(1).X = iUCWidth * 0.2
            iPts(1).Y = iUCHeight * 0.995
            ' left
            iPts(2).X = iUCWidth * 0.015
            iPts(2).Y = iUCHeight * 0.85
            iPts(3).X = iUCWidth * 0.015
            iPts(3).Y = iUCHeight * 0.6
            ' left middle
            iPts(4).X = iUCWidth * 0.11
            iPts(4).Y = iUCHeight * 0.45
            ' point pushing inside
            iPts(5).X = iUCWidth * 0.22
            iPts(5).Y = iUCHeight * 0.4
            ' going up
            iPts(6).X = iUCWidth * 0.25
            iPts(6).Y = iUCHeight * 0.2
            iPts(7).X = iUCWidth * 0.29
            iPts(7).Y = iUCHeight * 0.12
            ' top
            iPts(8).X = iUCWidth * 0.35
            iPts(8).Y = iUCHeight * 0.07
            iPts(9).X = iUCWidth * 0.5
            iPts(9).Y = iUCHeight * 0.1
            iPts(10).X = iUCWidth * 0.63
            iPts(10).Y = iUCHeight * 0.3
            ' going down, new part
            iPts(11).X = iUCWidth * 0.63
            iPts(11).Y = iUCHeight * 0.3
            iPts(12).X = iUCWidth * 0.72
            iPts(12).Y = iUCHeight * 0.27
            iPts(13).X = iUCWidth * 0.78
            iPts(13).Y = iUCHeight * 0.37
            iPts(14).X = iUCWidth * 0.8
            iPts(14).Y = iUCHeight * 0.56
            iPts(15).X = iUCWidth * 0.8
            iPts(15).Y = iUCHeight * 0.56
            ' to the right
            iPts(16).X = iUCWidth * 0.9
            iPts(16).Y = iUCHeight * 0.7
            iPts(17).X = iUCWidth * 0.9
            iPts(17).Y = iUCHeight * 0.9
            iPts(18).X = iUCWidth * 0.8
            iPts(18).Y = iUCHeight * 0.995
            iPts(19).X = iUCWidth * 0.51
            iPts(19).Y = iUCHeight * 0.995

            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If

            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
        ElseIf mShape = seShapeTalk Then
            iLng = mShift
            If iLng < 0 Then iLng = 0
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            iShift = iUCWidth / 100 * (mShift - 18) * 0.5
            If iShift > 300 Then iShift = 300
            If iShift < -300 Then iShift = -300
            If iLng > 0 Then
                ReDim iPts(16)
            Else
                ReDim iPts(9)
            End If
            
            ' left
            If iLng > 0 Then
                iPts(0).X = iUCWidth * 0.09
                iPts(0).Y = iUCHeight * 0.74
            Else
                iPts(0).X = iUCWidth * 0.15
                iPts(0).Y = iUCHeight * 0.78
            End If
            iPts(1).X = iUCWidth * 0.05
            iPts(1).Y = iUCHeight * 0.65
            iPts(2).X = iUCWidth * 0.05
            iPts(2).Y = iUCHeight * 0.27
            ' top
            iPts(3).X = iUCWidth * 0.15
            iPts(3).Y = iUCHeight * 0.1
            iPts(4).X = iUCWidth * 0.5
            iPts(4).Y = iUCHeight * 0.05
            iPts(5).X = iUCWidth * 0.85
            iPts(5).Y = iUCHeight * 0.1
            ' right
            iPts(6).X = iUCWidth * 0.99
            iPts(6).Y = iUCHeight * 0.26
            iPts(7).X = iUCWidth * 0.965
            iPts(7).Y = iUCHeight * 0.66
            ' bottom
            iPts(8).X = iUCWidth * 0.78
            iPts(8).Y = iUCHeight * 0.77
            iPts(9).X = iUCWidth * 0.4
            iPts(9).Y = iUCHeight * 0.78
            If iLng > 0 Then
                ' bottom left, the following is the start of the spike
                iPts(10).X = iUCWidth * 0.31
                iPts(10).Y = iUCHeight * 0.78 + iShift * 0.035
                iPts(11).X = iPts(10).X
                iPts(11).Y = iPts(10).Y
                iPts(12).X = iUCWidth * 0.25
                iPts(12).Y = iUCHeight * 0.81 + iShift * 0.04
                ' bottom left, the following is the point spike
                iPts(13).X = iUCWidth * 0.01 - iShift
                iPts(13).Y = iUCHeight * 0.99 + iShift * 0.5
                iPts(14).X = iUCWidth * 0.115
                iPts(14).Y = iUCHeight * 0.85
                iPts(15).X = iUCWidth * 0.14
                iPts(15).Y = iUCHeight * 0.81
                iPts(16).X = iUCWidth * 0.15
                iPts(16).Y = iUCHeight * 0.77
            End If
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            
'            UserControl.DrawWidth = 10
'            On Error Resume Next
'            For c = 0 To UBound(iPts)
'                UserControl.PSet (iPts(c).X, iPts(c).Y), IIf(c = 10, vbGreen, IIf(c = 13, vbBlue, vbRed))
'            Next
'            On Error GoTo 0

            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
            
            If mShift < 0 Then
                iShift = iShift * -1
                If iFilled Then
                    FillEllipse iGraphics, iFillColor, iUCWidth * 0.24 - iShift * 0.4, iUCHeight * 0.79 + iShift * 0.05, iUCWidth * 0.05 + iUCWidth * 0.05 * iShift / 150, iUCHeight * 0.1 + iUCHeight * 0.1 * iShift / 150
                    FillEllipse iGraphics, iFillColor, iUCWidth * 0.23 - iShift * 0.7, iUCHeight * 0.84 + iShift * 0.16, iUCWidth * 0.035 + iUCWidth * 0.035 * iShift / 150, iUCHeight * 0.07 + iUCHeight * 0.07 * iShift / 150
                    FillEllipse iGraphics, iFillColor, iUCWidth * 0.18 - iShift * 0.9, iUCHeight * 0.92 + iShift * 0.22, iUCWidth * 0.025 + iUCWidth * 0.025 * iShift / 150, iUCHeight * 0.05 + iUCHeight * 0.05 * iShift / 150
                End If
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth * 0.24 - iShift * 0.4, iUCHeight * 0.79 + iShift * 0.05, iUCWidth * 0.05 + iUCWidth * 0.05 * iShift / 150, iUCHeight * 0.1 + iUCHeight * 0.1 * iShift / 150
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth * 0.23 - iShift * 0.7, iUCHeight * 0.84 + iShift * 0.16, iUCWidth * 0.035 + iUCWidth * 0.035 * iShift / 150, iUCHeight * 0.07 + iUCHeight * 0.07 * iShift / 150
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth * 0.18 - iShift * 0.9, iUCHeight * 0.92 + iShift * 0.22, iUCWidth * 0.025 + iUCWidth * 0.025 * iShift / 150, iUCHeight * 0.05 + iUCHeight * 0.05 * iShift / 150
            End If
            
        ElseIf mShape = seShapeShield Then
            ReDim iPts(22)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            'point top
            iPts(0).X = iUCWidth * 0.51
            iPts(0).Y = iUCHeight * 0.005
            iPts(1).X = iUCWidth * 0.49
            iPts(1).Y = iUCHeight * 0.005
            ' side top-left
            iPts(2).X = iUCWidth * 0.4
            iPts(2).Y = iUCHeight * 0.07 '
            iPts(3).X = iUCWidth * 0.275
            iPts(3).Y = iUCHeight * 0.14 '
            iPts(4).X = iUCWidth * 0.14
            iPts(4).Y = iUCHeight * 0.202  '
            iPts(5).X = iUCWidth * 0.047
            iPts(5).Y = iUCHeight * 0.237
            ' point left
            iPts(6).X = iUCWidth * 0.005
            iPts(6).Y = iUCHeight * 0.252
            iPts(7).X = iUCWidth * 0.007
            iPts(7).Y = iUCHeight * 0.262
            iPts(8).X = iUCWidth * 0.01
            iPts(8).Y = iUCHeight * 0.28
            ' side bottom-left
            iPts(9).X = iUCWidth * 0.1
            iPts(9).Y = iUCHeight * 0.57
            iPts(10).X = iUCWidth * 0.27
            iPts(10).Y = iUCHeight * 0.83
            ' point bottom
            iPts(11).X = iUCWidth * 0.465
            iPts(11).Y = iUCHeight * 0.973
            iPts(12).X = iUCWidth * 0.5
            iPts(12).Y = iUCHeight * 0.995
            iPts(13).X = iUCWidth * 0.535
            iPts(13).Y = iUCHeight * 0.973
            ' side bottom right
            iPts(14).X = iUCWidth * 0.73
            iPts(14).Y = iUCHeight * 0.83
            iPts(15).X = iUCWidth * 0.9
            iPts(15).Y = iUCHeight * 0.57
            ' point right
            iPts(16).X = iUCWidth * 0.99
            iPts(16).Y = iUCHeight * 0.28
            iPts(17).X = iUCWidth * 0.993
            iPts(17).Y = iUCHeight * 0.262
            iPts(18).X = iUCWidth * 0.995
            iPts(18).Y = iUCHeight * 0.252
            ' side top right
            iPts(19).X = iUCWidth * 0.953
            iPts(19).Y = iUCHeight * 0.237
            iPts(20).X = iUCWidth * 0.86
            iPts(20).Y = iUCHeight * 0.202
            iPts(21).X = iUCWidth * 0.725
            iPts(21).Y = iUCHeight * 0.14 '
            iPts(22).X = iUCWidth * 0.6
            iPts(22).Y = iUCHeight * 0.07 '
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
        ElseIf mShape = seShapePie Then
            If iFilled Then
                FillPie iGraphics, iFillColor, 0, 0, iUCWidth, iUCHeight, mShift + 60
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPie iGraphics, mBorderColor, mBorderWidth, 0, 0, iUCWidth, iUCHeight, mShift + 60
            End If
        Else ' mShape = seShapeRectangle
            ReDim iPts(3)
            
            iPts(0).X = 0
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = 0
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(1).Y = iPts(1).Y - iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        End If
        
        Call GdipDeleteGraphics(iGraphics)
    End If
End Sub

Private Sub FillPolygon(ByVal nGraphics As Long, ByVal nColor As Long, Points() As POINTL, Optional nFillMode As FillModeConstants = FillModeAlternate)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints() As POINTL
    Dim c As Long
    
    If mStyle3D <> 0 Then
        If mStyle3DEffect = seStyle3EffectAuto Then
            If (mShape = seShapeParalellogram) Or (mShape = seShapeRectangle) Or (mShape = seShapeSquare) Or (mShape = seShapeTrapezoid) Or (mShape = seShapeTriangleScalene) Or (mShape = seShapeTriangleRight) Or (mShape = seShapeTriangleIsosceles) Or (mShape = seShapeTriangleEquilateral) Then
                iStyle3DEffect = seStyle3EffectDiffuse
            Else
                iStyle3DEffect = seStyle3EffectGem
            End If
        Else
            iStyle3DEffect = mStyle3DEffect
        End If
        
        If iStyle3DEffect = seStyle3EffectDiffuse Then
            GdipCreatePath 0&, iPath
            iRect = ScaleRect(GetPointsLRect(Points), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
            GdipAddPathEllipseI iPath, iRect.Left, iRect.top, iRect.Right - iRect.Left, iRect.Bottom - iRect.top
            iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
        Else
            ReDim iPoints(UBound(Points))
            For c = 0 To UBound(Points)
                iPoints(c) = Points(c)
            Next
            If mCurvingFactor <> 0 Then
                If (mShape = seShapeTriangleScalene) Or (mShape = seShapeTriangleRight) Then
                    ' add a point
                    ReDim Preserve iPoints(3)
                    iPoints(3).X = (iPoints(0).X ^ 2 + iPoints(2).X ^ 2) ^ 0.5 + UserControl.ScaleWidth / 300 * Abs(mCurvingFactor)
                    iPoints(3).Y = (iPoints(0).Y ^ 2 + iPoints(2).Y ^ 2) ^ 0.5 - UserControl.ScaleHeight / 300 * Abs(mCurvingFactor)
                End If
                iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
            End If
            iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
        End If
        If iRet = 0 Then
            GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
            GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
        End If
    Else
        iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        If mCurvingFactor = 0 Then
            GdipFillPolygonI nGraphics, hBrush, Points(0), UBound(Points) + 1, nFillMode
        Else
            GdipFillClosedCurve2I nGraphics, hBrush, Points(0), UBound(Points) + 1, mCurvingFactor2, nFillMode
        End If
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
    
End Sub

Private Sub DrawPolygon(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawnWidth As Long, Points() As POINTL)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawnWidth, UnitPixel, hPen) = 0 Then
        If ((mBorderStyle = vbBSSolid) Or (mBorderStyle = vbBSInsideSolid)) Or (mBorderWidth > 1) Then
            Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        Else
            Call GdipSetSmoothingMode(nGraphics, QualityModeLow)
        End If
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mCurvingFactor = 0 Then
            GdipDrawPolygonI nGraphics, hPen, Points(0), UBound(Points) + 1
        Else
            GdipDrawClosedCurve2I nGraphics, hPen, Points(0), UBound(Points) + 1, mCurvingFactor2
        End If
        Call GdipDeletePen(hPen)
    End If
    
End Sub

Private Sub FillClosedCurve(ByVal nGraphics As Long, ByVal nColor As Long, Points() As POINTL, ByVal nTension As Single, Optional nFillMode As FillModeConstants = FillModeAlternate)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iPoints() As POINTL
    Dim iPath As Long
    Dim iStyle3DEffect As Long
    Dim iRect As RECT
    
    If mStyle3D <> 0 Then
        If mStyle3DEffect = seStyle3EffectAuto Then
            If (mShape = seShapeCrescent) Or (mShape = seShapeLocation) Or (mShape = seShapeCloud) Then
                iStyle3DEffect = seStyle3EffectDiffuse
            Else
                iStyle3DEffect = seStyle3EffectGem
            End If
        Else
            iStyle3DEffect = mStyle3DEffect
        End If
        
        If iStyle3DEffect = seStyle3EffectDiffuse Then
            GdipCreatePath 0&, iPath
            iRect = ScaleRect(GetPointsLRect(Points), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
            GdipAddPathEllipseI iPath, iRect.Left, iRect.top, iRect.Right - iRect.Left, iRect.Bottom - iRect.top
            iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
        Else
            iPoints = ExpandPointsL(Points, 0.05)
            If mCurvingFactor <> 0 Then
                iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
            End If
            iRet = GdipCreatePathGradientI(iPoints(0), UBound(Points) + 1, 0&, hBrush)
        End If
        If iRet = 0 Then
            GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
            GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
        End If
    Else
        iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        GdipFillClosedCurve2I nGraphics, hBrush, Points(0), UBound(Points) + 1, nTension, nFillMode
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
    
End Sub

Private Sub DrawClosedCurve(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawnWidth As Long, Points() As POINTL, ByVal nTension As Single)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawnWidth, UnitPixel, hPen) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        GdipDrawClosedCurve2I nGraphics, hPen, Points(0), UBound(Points) + 1, nTension
        Call GdipDeletePen(hPen)
    End If
    
End Sub


Private Sub FillEllipse(ByVal nGraphics As Long, ByVal nColor As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints(3) As POINTL
    
    If mStyle3D <> 0 Then
        If mStyle3DEffect = seStyle3EffectAuto Then
            iStyle3DEffect = seStyle3EffectDiffuse
        Else
            iStyle3DEffect = mStyle3DEffect
        End If
        
        If iStyle3DEffect = seStyle3EffectDiffuse Then
            GdipCreatePath 0&, iPath
            GdipAddPathEllipseI iPath, X, Y, nWidth, nHeight
            iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
        Else
            iPoints(0).X = X
            iPoints(0).Y = Y
            iPoints(1).X = X + nWidth
            iPoints(1).Y = Y
            iPoints(2).X = X + nWidth
            iPoints(2).Y = Y + nHeight
            iPoints(3).X = X
            iPoints(3).Y = Y + nHeight
            iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
        End If
        If iRet = 0 Then
            GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
            GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
        End If
    Else
        iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        GdipFillEllipseI nGraphics, hBrush, X, Y, nWidth, nHeight
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
End Sub

Private Sub DrawEllipse(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawnWidth As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawnWidth, UnitPixel, hPen) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mBorderStyle = vbBSInsideSolid Then
            X = X + nDrawnWidth / 2
            Y = Y + nDrawnWidth / 2
            nWidth = nWidth - nDrawnWidth
            nHeight = nHeight - nDrawnWidth
        End If
        GdipDrawEllipseI nGraphics, hPen, X, Y, nWidth, nHeight
        Call GdipDeletePen(hPen)
    End If
    
End Sub

Private Sub FillRoundRect(ByVal nGraphics As Long, ByVal nColor As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nRoundSize As Long = 10)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints() As POINTL
    
    nRoundSize = nRoundSize * 2
    If mStyle3D <> 0 Then
        If mStyle3DEffect = seStyle3EffectAuto Then
            iStyle3DEffect = seStyle3EffectDiffuse
        Else
            iStyle3DEffect = mStyle3DEffect
        End If

        ReDim iPoints(3)
        iPoints(0).X = X
        iPoints(0).Y = Y
        iPoints(1).X = X + nWidth
        iPoints(1).Y = Y
        iPoints(2).X = X + nWidth
        iPoints(2).Y = Y + nHeight
        iPoints(3).X = X
        iPoints(3).Y = Y + nHeight
        If iStyle3DEffect = seStyle3EffectDiffuse Then
            GdipCreatePath 0&, iPath
            iRect = ScaleRect(GetPointsLRect(iPoints), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
            GdipAddPathEllipseI iPath, iRect.Left, iRect.top, iRect.Right - iRect.Left, iRect.Bottom - iRect.top
            iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
        Else
            If mCurvingFactor <> 0 Then
                iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
            End If
            iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
        End If
        If iRet = 0 Then
            GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
            GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
        End If
    Else
        iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
    End If

    If iRet = 0 Then
        GdipCreatePath &H0, iPath
        GdipAddPathArcI iPath, X, Y, nRoundSize, nRoundSize, 180, 90
        GdipAddPathArcI iPath, X + nWidth - nRoundSize, Y, nRoundSize, nRoundSize, 270, 90
        GdipAddPathArcI iPath, X + nWidth - nRoundSize, Y + nHeight - nRoundSize, nRoundSize, nRoundSize, 0, 90
        GdipAddPathArcI iPath, X, Y + nHeight - nRoundSize, nRoundSize, nRoundSize, 90, 90
        GdipClosePathFigure iPath
        GdipFillPath nGraphics, hBrush, iPath
        
        Call GdipDeletePath(iPath)
        Call GdipDeleteBrush(hBrush)
    End If
End Sub

Private Sub DrawRoundRect(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawnWidth As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nRoundSize As Long = 10)
    Dim iPath As Long
    Dim hPen As Long
    
    nRoundSize = nRoundSize * 2
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawnWidth, UnitPixel, hPen) = 0 Then
        If ((mBorderStyle = vbBSSolid) Or (mBorderStyle = vbBSInsideSolid)) Or (mBorderWidth > 1) Then
            Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        Else
            Call GdipSetSmoothingMode(nGraphics, QualityModeLow)
        End If
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mBorderStyle = vbBSInsideSolid Then
            X = X + nDrawnWidth / 2
            Y = Y + nDrawnWidth / 2
            nWidth = nWidth - nDrawnWidth
            nHeight = nHeight - nDrawnWidth
        End If

        GdipCreatePath &H0, iPath
        GdipAddPathArcI iPath, X, Y, nRoundSize, nRoundSize, 180, 90
        GdipAddPathArcI iPath, X + nWidth - nRoundSize, Y, nRoundSize, nRoundSize, 270, 90
        GdipAddPathArcI iPath, X + nWidth - nRoundSize, Y + nHeight - nRoundSize, nRoundSize, nRoundSize, 0, 90
        GdipAddPathArcI iPath, X, Y + nHeight - nRoundSize, nRoundSize, nRoundSize, 90, 90
        GdipClosePathFigure iPath
        GdipDrawPath nGraphics, hPen, iPath
        
        Call GdipDeletePath(iPath)
        Call GdipDeletePen(hPen)
    End If
End Sub

Private Sub FillSemicircle(ByVal nGraphics As Long, ByVal nColor As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints() As POINTL
    
    If mStyle3D <> 0 Then
        If mStyle3DEffect = seStyle3EffectAuto Then
            iStyle3DEffect = seStyle3EffectDiffuse
        Else
            iStyle3DEffect = mStyle3DEffect
        End If
        
        ReDim iPoints(3)
        iPoints(0).X = X
        iPoints(0).Y = Y
        iPoints(1).X = X + nWidth
        iPoints(1).Y = Y
        iPoints(2).X = X + nWidth
        iPoints(2).Y = Y + nHeight
        iPoints(3).X = X
        iPoints(3).Y = Y + nHeight
        If iStyle3DEffect = seStyle3EffectDiffuse Then
            GdipCreatePath 0&, iPath
            iRect = ScaleRect(GetPointsLRect(iPoints), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
            GdipAddPathEllipseI iPath, iRect.Left, iRect.top, iRect.Right - iRect.Left, iRect.Bottom - iRect.top
            iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
        Else
            If mCurvingFactor <> 0 Then
                iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
            End If
            iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
        End If
        If iRet = 0 Then
            GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
            GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
        End If
    Else
        iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        GdipFillPieI nGraphics, hBrush, X, Y, nWidth, nHeight * 2, 180, 180
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
End Sub

Private Sub DrawSemicircle(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawnWidth As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawnWidth, UnitPixel, hPen) = 0 Then
        If ((mBorderStyle = vbBSSolid) Or (mBorderStyle = vbBSInsideSolid)) Or (mBorderWidth > 1) Then
            Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        Else
            Call GdipSetSmoothingMode(nGraphics, QualityModeLow)
        End If
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mBorderStyle = vbBSInsideSolid Then
            X = X + nDrawnWidth / 2
            Y = Y + nDrawnWidth / 2
            nWidth = nWidth - nDrawnWidth
            nHeight = nHeight - nDrawnWidth
        End If
        GdipDrawArcI nGraphics, hPen, X, Y, nWidth, nHeight * 2 + nDrawnWidth, 180, 180
        GdipDrawLineI nGraphics, hPen, X + nDrawnWidth / 2 - 1, Y + nHeight, X + nWidth - nDrawnWidth / 2 + 1, Y + nHeight
        Call GdipDeletePen(hPen)
    End If
    
End Sub

Private Function ConvertColor(nColor As Long, nOpacity As Single) As Long
    Dim BGRA(0 To 3) As Byte
    Dim iColor As Long
    
    TranslateColor nColor, 0&, iColor
    
    BGRA(3) = CByte((nOpacity / 100) * 255)
    BGRA(0) = ((iColor \ &H10000) And &HFF)
    BGRA(1) = ((iColor \ &H100) And &HFF)
    BGRA(2) = (iColor And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

Private Sub InitGDI()
    Dim aInput(0 To 3)  As Long
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
End Sub


Private Property Get SmoothingMode() As Long
    If mQuality = seQualityHigh Then
        SmoothingMode = SmoothingModeAntiAlias
    Else
        SmoothingMode = QualityModeLow
    End If
End Property


'--- for MST subclassing (2)
'Autor: wqweto http://www.vbforums.com/showthread.php?872819
'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================
Private Sub pvSubclass()
    Dim iDo As Boolean
    
    If UserControl.ContainerHwnd <> 0 Then
        If mUseSubclassing = seSCYes Then
           iDo = True
        ElseIf mUseSubclassing = seSCNotInIDE Then
            iDo = Not InIDE
        ElseIf mUseSubclassing = seSCNotInIDEDesignTime Then
            If mUserMode Then
                iDo = True
            Else
                iDo = Not InIDE
            End If
        End If
        If iDo Then
            Set m_pSubclass = InitSubclassingThunk(UserControl.ContainerHwnd, InitAddressOfMethod().SubclassProc(0, 0, 0, 0, 0))
            mSubclassed = True
        End If
    End If
End Sub

Private Function InIDE() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        Err.Clear
        On Error Resume Next
        Debug.Print 1 / 0
        If Err.Number Then
            sValue = 1
        Else
            sValue = 2
        End If
        Err.Clear
    End If
    InIDE = (sValue = 1)
End Function

Private Sub pvUnsubclass()
    Set m_pSubclass = Nothing
    mSubclassed = False
End Sub

Private Function InitAddressOfMethod() As ShapeEx
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(Me), 5, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function InitSubclassingThunk(ByVal hWnd As Long, ByVal pfnCallback As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepwEDMAV1aLdCQUg8YIgz4AdC+L+oHH/BEzAIvCBQQRMwCri8IFQBEzAKuLwgVQETMAq4vCBXgRMwCruQkAAADzpYHC/BEzAFJqFP9SEFqL+IvCq7gBAAAAq4tEJAyri3QkFKWlg+8UagBX/3IM/3cI/1IYi0QkGIk4Xl+4MBIzAC1wEDMAwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEdRiLClL/cQz/cgj/URyLVCQEiwpS/1EUM8DCBACQVYvsi1UYiwqLQSyFwHQ1Uv/QWoP4AXdUg/gAdQmBfQwDAgAAdEaLClL/UTBahcB1O4sKUmrw/3Ek/1EoWqkAAAAIdShSM8BQUI1EJARQjUQkBFD/dRT/dRD/dQz/dQj/cgz/UhBZWFqFyXURiwr/dRT/dRD/dQz/dQj/USBdwhgADx8A" ' 29.3.2019 13:04:54
    Const THUNK_SIZE    As Long = 448
    Dim hThunk          As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(Me)
    aParams(1) = pfnCallback
    hThunk = GetProp(pvGetGlobalHwnd(), "InitSubclassingThunk")
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 410)      '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 412)      '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 413)      '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        Call SetProp(pvGetGlobalHwnd(), "InitSubclassingThunk", hThunk)
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function pvGetIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvGetIdeOwner = True
End Function

Private Function pvGetGlobalHwnd() As Long
    pvGetGlobalHwnd = FindWindowEx(0, 0, "STATIC", App.hInstance & ":" & App.Threadid & ":MST Global Data")
    If pvGetGlobalHwnd = 0 Then
        pvGetGlobalHwnd = CreateWindowEx(0, "STATIC", App.hInstance & ":" & App.Threadid & ":MST Global Data", _
            0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0)
    End If
End Function

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute SubclassProc.VB_MemberFlags = "40"
    #If hWnd And wParam And lParam And Handled Then '--- touch args
    #End If
    Select Case wMsg
        Case WM_INVALIDATE
            Dim iMessage As T_MSG
            
            PeekMessage iMessage, hWnd, WM_INVALIDATE, WM_INVALIDATE, PM_REMOVE  ' remove posted message, if any
            InvalidateRectAsNull hWnd, 0&, 1&
    End Select
    If Not mUserMode Then
        Handled = True
        SubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
    End If
End Function
'--- End for MST subclassing (2)

Private Sub SetCurvingFactor2()
    If mCurvingFactor < 0 Then
        mCurvingFactor2 = mCurvingFactor / 100 * 0.5
    Else
        mCurvingFactor2 = mCurvingFactor / 100 * 1
    End If
End Sub

' From Leandro Ascierto
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

Private Function ExpandPointsL(nPoints() As POINTL, nExpand As Single) As POINTL()
    Dim iCount As Long
    Dim c As Long
    Dim iRect As RECT
    Dim iCenterX As Long
    Dim iCenterY As Long
    Dim iRet() As POINTL
    
    iRect.top = 0
    iRect.Bottom = 0
    iRect.Left = UserControl.ScaleWidth * 2
    iRect.top = UserControl.ScaleHeight * 2
    iCount = UBound(nPoints) + 1
    For c = 0 To iCount - 1
        If nPoints(c).X < iRect.Left Then iRect.Left = nPoints(c).X
        If nPoints(c).X > iRect.Right Then iRect.Right = nPoints(c).X
        If nPoints(c).Y < iRect.top Then iRect.top = nPoints(c).Y
        If nPoints(c).Y > iRect.Bottom Then iRect.Bottom = nPoints(c).Y
    Next
    iCenterX = (iRect.Left + iRect.Right) / 2
    iCenterY = (iRect.top + iRect.Bottom) / 2
    ReDim iRet(iCount - 1)
    For c = 0 To iCount - 1
        If nPoints(c).X > iCenterX Then
            iRet(c).X = nPoints(c).X + Abs(iCenterX - nPoints(c).X) * nExpand
        ElseIf nPoints(c).X < iCenterX Then
            iRet(c).X = nPoints(c).X - Abs(iCenterX - nPoints(c).X) * nExpand
        Else
            iRet(c).X = nPoints(c).X
        End If
        If nPoints(c).Y > iCenterY Then
            iRet(c).Y = nPoints(c).Y + Abs(iCenterY - nPoints(c).Y) * nExpand
        ElseIf nPoints(c).Y < iCenterY Then
            iRet(c).Y = nPoints(c).Y - Abs(iCenterY - nPoints(c).Y) * nExpand
        Else
            iRet(c).Y = nPoints(c).Y
        End If
    Next
    ExpandPointsL = iRet
End Function

Private Function GetPointsLRect(nPoints() As POINTL) As RECT
    Dim iCount As Long
    Dim c As Long
    Dim iRect As RECT
    
    iRect.top = 0
    iRect.Bottom = 0
    iRect.Left = UserControl.ScaleWidth * 2
    iRect.top = UserControl.ScaleHeight * 2
    iCount = UBound(nPoints) + 1
    For c = 0 To iCount - 1
        If nPoints(c).X < iRect.Left Then iRect.Left = nPoints(c).X
        If nPoints(c).X > iRect.Right Then iRect.Right = nPoints(c).X
        If nPoints(c).Y < iRect.top Then iRect.top = nPoints(c).Y
        If nPoints(c).Y > iRect.Bottom Then iRect.Bottom = nPoints(c).Y
    Next
    GetPointsLRect = iRect
End Function

Private Function ScaleRect(nRect As RECT, ByVal nScale As Single) As RECT
    Dim iRect As RECT
    
    nScale = (nScale - 1) / 2
    
    iRect = nRect
    InflateRect iRect, (nRect.Right - nRect.Left) * nScale, (nRect.Bottom - nRect.top) * nScale
    ScaleRect = iRect
End Function

Private Sub DrawPie(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawnWidth As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nAngleMissingPart As Single)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawnWidth, UnitPixel, hPen) = 0 Then
        If ((mBorderStyle = vbBSSolid) Or (mBorderStyle = vbBSInsideSolid)) Or (mBorderWidth > 1) Then
            Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        Else
            Call GdipSetSmoothingMode(nGraphics, QualityModeLow)
        End If
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mBorderStyle = vbBSInsideSolid Then
            X = X + nDrawnWidth / 2
            Y = Y + nDrawnWidth / 2
            nWidth = nWidth - nDrawnWidth
            nHeight = nHeight - nDrawnWidth
        End If
        
        nAngleMissingPart = Abs(nAngleMissingPart)
        If nAngleMissingPart > 360 Then nAngleMissingPart = 360
        GdipDrawPieI nGraphics, hPen, X, Y, nWidth, nHeight, 0, 360 - nAngleMissingPart
        
        Call GdipDeletePen(hPen)
    End If
End Sub

Private Sub FillPie(ByVal nGraphics As Long, ByVal nColor As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nAngleMissingPart As Single)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints() As POINTL
    
    If mStyle3D <> 0 Then
        If mStyle3DEffect = seStyle3EffectAuto Then
            iStyle3DEffect = seStyle3EffectDiffuse
        Else
            iStyle3DEffect = mStyle3DEffect
        End If
        
        ReDim iPoints(3)
        iPoints(0).X = X
        iPoints(0).Y = Y
        iPoints(1).X = X + nWidth
        iPoints(1).Y = Y
        iPoints(2).X = X + nWidth
        iPoints(2).Y = Y + nHeight
        iPoints(3).X = X
        iPoints(3).Y = Y + nHeight
        If iStyle3DEffect = seStyle3EffectDiffuse Then
            GdipCreatePath 0&, iPath
            iRect = ScaleRect(GetPointsLRect(iPoints), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
            GdipAddPathEllipseI iPath, iRect.Left, iRect.top, iRect.Right - iRect.Left, iRect.Bottom - iRect.top
            iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
        Else
            If mCurvingFactor <> 0 Then
                iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
            End If
            iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
        End If
        If iRet = 0 Then
            GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
            GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
        End If
    Else
        iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        nAngleMissingPart = Abs(nAngleMissingPart)
        If nAngleMissingPart > 360 Then nAngleMissingPart = 360
        GdipFillPieI nGraphics, hBrush, X, Y, nWidth, nHeight, 0, 360 - nAngleMissingPart
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
End Sub
