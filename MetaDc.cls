VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MetaDc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type SIZEL
cX As Long
cY As Long
End Type
Private Type ENHMETAHEADER  ' 88
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
Private Type METAFILEPICT
    mm As Long
    xExt As Long
    yExt As Long
    hmf As Long
End Type
Private Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hmf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    reserved    As Long
    CheckSum    As Integer
End Type
Private mHeader As ENHMETAHEADER
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileW" (ByVal hdcRef As Long, ByVal lpFileName As Long, ByVal lpRect As Long, ByVal lpDescription As Long) As Long
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hmf As Long, ByVal cbBuffer As Long, lpemh As Any) As Long
Private Declare Function GetEnhMetaFileBits Lib "gdi32" (ByVal hmf As Long, ByVal nSize As Long, lpvData As Any) As Long
Private Declare Function CopyEnhMetaFile Lib "gdi32.dll" Alias "CopyEnhMetaFileW" (ByVal hemfSrc As Long, lpszFile As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hEmf As Long, lpRect As Any) As Long
Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Private Declare Function GetGraphicsMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As Any) As Long
Private Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hmf As Long, ByVal nSize As Long, lpvData As Any) As Long
Private Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As Long


Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long



Private Const MM_TEXT = 1
Private Const MM_LOENGLISH = 4

 Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim pm As POINTAPI

Dim boundrect As RECT
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Type PICTDESC
    Size                As Long
    Type                As Long
    hBmpOrIcon          As Long
    hPal                As Long
End Type
Private s$
Private m_fontname As String
Private m_fontsize As Single
Private m_currentX As Single
Private m_currenty As Single
Private m_PageStyle As Long
Private m_transparent As Long
Private m_fonttransparent As Long
Private m_backcolor As Long
Private m_hDC As Long, m_Status As Boolean, m_mousepointer As Integer
Private moufa As StdPicture
Private m_font As StdFont, m_hFntOld As Long, m_prive As Long
' maybe these change to single
Private m_top As Long, m_left As Long, m_width As Long, m_height As Long, m_pen As Long
Private m_DrawWidth As Long, m_fillstyle As Long, m_DrawStyle, hOldPen As Long, m_hpen As Long
Private m_fillcolor As Long, m_brushOld&, m_brush&, m_linebrush&
Private PrevStretchMode As Long
Dim WithEvents afont As StdFont
Attribute afont.VB_VarHelpID = -1
Private Type XFORM
        eM11 As Single
        eM12 As Single
        eM21 As Single
        eM22 As Single
        eDx As Single
        eDy As Single
End Type

 Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) _
          As Long
Private Declare Function GetWorldTransform Lib "gdi32" (ByVal hDC As Long, lpXform As XFORM) As Long
Private Declare Function SetWorldTransform Lib "gdi32" (ByVal hDC As Long, lpXform As XFORM) As Long

Private Const GM_ADVANCED = 2
Private Const GM_COMPATIBLE = 1
Private mDrawMode
Private gOldWorld As XFORM


Public Property Get Name() As String
    Name = "Emf"
End Property

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Public Property Get status() As Boolean
    status = m_Status
End Property
Public Property Get mouseicon() As StdPicture
    Set mouseicon = moufa
End Property
Public Property Let mouseicon(RHS As StdPicture)
    Set moufa = RHS
End Property
Private Function MyMod(R1, po) As Variant
MyMod = R1 - Fix(R1 / po) * po
End Function
Public Sub create(pen As Long, fakeWidth As Long, fakeHeight As Long)
    Static onetime As Boolean
    Dim dib As cDIBSection
    If onetime Then Exit Sub
    onetime = True
    Set dib = New cDIBSection
    If Not dib.create32(100, 100) Then Exit Sub
    m_width = fakeWidth
    m_height = fakeHeight
    m_pen = pen
    If m_width <> 0 And m_height <> 0 Then
    m_width = (m_width \ dv15) * dv15
    m_height = (m_height \ dv15) * dv15
    boundrect.Left = 0
    boundrect.Top = 0
    boundrect.Right = (m_width - dv15) * (2540 / 1440)
    boundrect.Bottom = (m_height - dv15) * (2540 / 1440)
    fakeWidth = boundrect.Right * 1440 / 2540
    fakeHeight = boundrect.Bottom * 1440 / 2540
    dib.needHDC
        m_hDC = CreateEnhMetaFile(dib.HDC1, 0, VarPtr(boundrect.Left), StrPtr(s$))
    dib.FreeHDC
    Else
    'm_width = m_width - 2 * dv15
    'm_height = m_height - 2 * dv15
    dib.needHDC
        m_hDC = CreateEnhMetaFile(dib.HDC1, 0, 0, StrPtr(s$))
    dib.FreeHDC
    End If
    m_DrawStyle = vbSolid  ' 0
    If m_hDC Then
        Call SetTextColor(m_hDC, pen)
        m_transparent = 1
        Call SetBkMode(m_hDC, m_transparent)  ' transparent background
        m_hpen = CreatePen(0, m_DrawWidth, m_pen)
        hOldPen = SelectObject(m_hDC, m_hpen)
        m_linebrush = CreateSolidBrush(m_pen)
        
        m_Status = True
        If Not m_font Is Nothing Then FixFontProp
     End If
     
End Sub
Public Function getEmfObj() As Object
Dim aPic As MemBlock
Dim bytes As Long
Dim emfP As StdPicture
Set emfP = Me.getEmf
If emfP Is Nothing Then Exit Function
Set aPic = New MemBlock
bytes = GetEnhMetaFileBits(emfP.Handle, bytes, ByVal 0)
If bytes Then
    aPic.Construct 1, bytes
    Call GetEnhMetaFileBits(emfP.Handle, bytes, ByVal aPic.GetBytePtr(0))
    aPic.SubType = 2 ' emf
End If
Set getEmfObj = aPic
End Function
Private Sub Cleanup()
Dim retval As Long
    If m_hFntOld <> 0 Then SelectObject m_hDC, m_hFntOld
    If hOldPen <> 0 Then SelectObject m_hDC, hOldPen: DeleteObject m_hpen
    If m_brushOld Then SelectObject m_hDC, m_brushOld: DeleteObject m_brush
    If m_linebrush <> 0 Then DeleteObject m_linebrush
End Sub
Public Function getEmf() As StdPicture

If m_Status Then
    Cleanup
    Set getEmf = pvMetaToPicture
    m_Status = False
End If
End Function
Private Function pvMetaToPicture() As IPicture
    Dim uDesc           As PICTDESC
    Dim aGUID(0 To 3)   As Long
    
    '--- fill struct
    With uDesc
        .Size = Len(uDesc)
        .Type = vbPicTypeEMetafile
        .hBmpOrIcon = CloseEnhMetaFile(m_hDC)
    End With

    m_hDC = 0
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    Call OleCreatePictureIndirect(uDesc, aGUID(0), True, pvMetaToPicture)

End Function
Public Property Get Picture() As StdPicture
If m_Status Then
    Cleanup
     Set Picture = pvMetaToPicture
    m_Status = False
End If
End Property

Private Sub afont_FontChanged(ByVal PropertyName As String)
    FixFontProp
End Sub

Private Sub Class_Initialize()
    s$ = "M2000" + Chr$(0) + "Export"
    m_DrawWidth = 1
End Sub

Public Property Get mousepointer() As Integer
    mousepointer = m_mousepointer
End Property

Public Property Let mousepointer(ByVal RHS As Integer)
    m_mousepointer = RHS
End Property
Public Property Get BackColor() As Long
BackColor = m_backcolor
End Property

Public Property Let BackColor(RHS As Long)
    m_backcolor = RHS
End Property

Public Property Get FontName() As String
    FontName = m_fontname
End Property

Public Property Let FontName(ANameS As String)
    m_fontname = ANameS
    FixFontProp
End Property
Public Property Get FontSize() As Variant
    FontSize = m_fontsize
End Property

Public Property Let FontSize(fsize As Variant)
If fsize < 1 Then fsize = 1
    m_fontsize = fsize
    m_font.Size = fsize
End Property

Public Property Get Transparent() As Variant
    Transparent = m_transparent
     If m_Status Then If m_hDC Then Call SetBkMode(m_hDC, m_transparent)
End Property

Public Property Let Transparent(trans As Variant)
    m_transparent = trans
End Property
Public Property Get Font() As Font
    Set Font = m_font
End Property
Private Sub FixFontProp()
With m_font
m_fontname = .Name
m_fontsize = .Size
If m_hFntOld <> 0 Then SelectObject m_hDC, m_hFntOld
Dim MyFont As IFont
Set MyFont = m_font
m_hFntOld = SelectObject(m_hDC, MyFont.hFont)
End With
End Sub

Public Property Set Font(New_Font As Font)
    Dim MyFont As IFont
    If m_Status Then
        If m_hFntOld <> 0 Then SelectObject m_hDC, m_hFntOld
    End If
    Set MyFont = New_Font
    MyFont.Clone m_font
    Set MyFont = m_font
    Set afont = m_font
    If m_Status Then
        FixFontProp
    End If
End Property

Private Sub Class_Terminate()
    If m_Status Then
        If m_hDC <> 0 Then
             Cleanup
             DeleteDC m_hDC
        End If
    End If
    Set m_font = Nothing
    Set afont = Nothing
    
End Sub

Public Property Get currentX() As Single
    currentX = m_currentX
End Property

Public Property Let currentX(ByVal RHS As Single)
    m_currentX = RHS
End Property

Public Property Get currentY() As Single
    currentY = m_currenty
End Property

Public Property Let currentY(ByVal RHS As Single)
m_currenty = RHS
End Property
Public Property Get prive() As Long
    prive = m_prive
End Property

Public Property Let prive(ByVal RHS As Long)
m_prive = RHS
End Property

Public Property Get Top() As Long
Top = m_top
End Property

Public Property Let Top(ByVal RHS As Long)
m_top = RHS
End Property

Public Property Get Left() As Long
Left = m_left
End Property

Public Property Let Left(ByVal RHS As Long)
m_left = RHS
End Property

Public Sub move(aLeft, Optional aTop, Optional aWidth, Optional aHeight)
    m_left = aLeft
If Not IsMissing(aTop) Then
    m_top = aTop
End If
If Not IsMissing(aWidth) Then
    m_width = aWidth
End If
If Not IsMissing(aHeight) Then
    m_height = aHeight
End If
End Sub

Public Property Get ForeColor() As Long
    ForeColor = m_pen
End Property
Public Sub makepen()
        Call SetTextColor(m_hDC, m_pen)
        If hOldPen <> 0 Then SelectObject m_hDC, hOldPen
        If m_hpen <> 0 Then DeleteObject m_hpen
        m_hpen = CreatePen(m_DrawStyle, m_DrawWidth, m_pen)
        hOldPen = SelectObject(m_hDC, m_hpen)
        If m_linebrush <> 0 Then DeleteObject m_linebrush
        m_linebrush = CreateSolidBrush(m_pen)
End Sub

Public Property Let ForeColor(ByVal RHS As Long)
    m_pen = RHS
    If m_Status Then If m_hDC Then makepen
End Property


Public Property Get fillstyle() As Long
fillstyle = m_fillstyle
End Property
Private Sub MakeBrush()
Const NULL_BRUSH = 5
Const DC_PEN = 19
Dim ret As Long
If m_fillstyle > 1 Then
    If m_brushOld Then SelectObject m_hDC, m_brushOld
    If m_brush Then DeleteObject m_brush
 
    Select Case m_fillstyle
    Case 6
        ret = 4
    Case 7
        ret = 5
    Case 4
        ret = 2
    Case 5
        ret = 2
    Case 3
        ret = 1
    Case 2
        ret = 0

    End Select
    m_brush = CreateHatchBrush(ret, m_fillcolor)
    m_brushOld = SelectObject(m_hDC, m_brush)
ElseIf m_fillstyle = 1 Then
    If m_brushOld Then SelectObject m_hDC, m_brushOld
    If m_brush Then DeleteObject m_brush
    m_brush = GetStockObject(NULL_BRUSH)
    m_brushOld = SelectObject(m_hDC, m_brush)
Else
    If m_brushOld Then SelectObject m_hDC, m_brushOld
    If m_brush Then DeleteObject m_brush
   ' m_brush = GetStockObject(DC_PEN)
   m_brush = CreateSolidBrush(m_fillcolor)
    
    m_brushOld = SelectObject(m_hDC, m_brush)
End If
End Sub

Public Property Let fillstyle(ByVal RHS As Long)
m_fillstyle = RHS
If m_Status Then If m_hDC Then MakeBrush
End Property

Public Property Get DrawWidth() As Long
DrawWidth = m_DrawWidth
End Property

Public Property Let DrawWidth(ByVal RHS As Long)
m_DrawWidth = RHS
Dim retval As Long
If m_Status Then
If m_hDC Then
If hOldPen <> 0 Then retval = SelectObject(m_hDC, hOldPen)
If m_hpen <> 0 Then retval = DeleteObject(m_hpen)
m_hpen = CreatePen(m_DrawStyle, m_DrawWidth, m_pen)
hOldPen = SelectObject(m_hDC, m_hpen)
End If
End If
End Property

Public Function ScaleX(ByVal RHS As Single, Optional ByVal dummy As Integer = 1, Optional ByVal dumm2 As Integer = 1) As Single
    ScaleX = Form1.ScaleX(RHS, dummy, dumm2)
End Function

Public Function ScaleY(ByVal RHS As Single, Optional ByVal dummy As Integer = 1, Optional ByVal dumm2 As Integer = 1) As Single
    ScaleY = Form1.ScaleY(RHS, dummy, dumm2)
End Function

Public Property Get DrawStyle() As Long
DrawStyle = m_DrawStyle
End Property

Public Property Let DrawStyle(ByVal RGS As Long)
m_DrawStyle = RGS
If m_hDC Then makepen
End Property

Public Property Get Visible() As Boolean
Visible = True
End Property

Public Property Let Visible(ByVal RHS As Boolean)
' do nothing
End Property

Public Sub Refresh()
' do nothing
End Sub

Public Property Get Width() As Long
    Width = m_width
End Property
Public Property Get Scalewidth() As Long
    Scalewidth = m_width
End Property
Public Property Let Width(ByVal RHS As Long)
    m_width = (RHS \ dv15) * dv15
End Property

Public Property Get fillcolor() As Long
fillcolor = m_fillcolor
End Property

Public Property Let fillcolor(ByVal RHS As Long)
m_fillcolor = RHS
If m_Status Then If m_hDC Then MakeBrush
End Property

Public Property Get Height() As Long
Height = m_height
End Property
Public Property Get Scaleheight() As Long
Scaleheight = m_height
End Property
Public Property Let Height(ByVal RHS As Long)
m_height = (RHS \ dv15) * dv15
End Property
' line bf
Public Sub Line2(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, RGBcolor As Long, Optional FillBack As Boolean = False, Optional DrawBlock As Boolean = False)
Static rp As RECT
If m_DrawWidth = 0 Then m_DrawWidth = 1
Const NULL_BRUSH = 5
Dim nowold As Long, oldpen As Long, linepen As Long, hollow As Long
If m_Status Then

Call SetBkMode(m_hDC, 0)


' draw here
If FillBack Then
linepen = CreateSolidBrush(RGBcolor)
rp.Left = x1 \ dv15
rp.Top = y1 \ dv15
rp.Right = Int(x2 / dv15 + 0.5)
rp.Bottom = Int(y2 / dv15 + 0.5) + 1
FillRect m_hDC, rp, linepen
DeleteObject linepen
If DrawBlock Then GoTo cont123
ElseIf DrawBlock Then
'linepen = CreateSolidBrush(RGBcolor)
rp.Left = x1 \ dv15
rp.Top = y1 \ dv15
rp.Right = Int(x2 / dv15 + 0.5)
rp.Bottom = Int(y2 / dv15 + 0.5) + 1
cont123:

linepen = CreatePen(m_DrawStyle, m_DrawWidth, RGBcolor)
'hollow = GetStockObject(NULL_BRUSH)
oldpen = SelectObject(m_hDC, linepen)
'nowold = SelectObject(m_hDC, hollow)
'Rectangle m_hDC, rp.Left, rp.top, rp.Right - 1, rp.Bottom - 1
'SelectObject m_hDC, nowold
MoveToEx m_hDC, rp.Left, rp.Top, ByVal 0
LineTo m_hDC, rp.Right, rp.Top
'MoveToEx m_hDC, x1 \ dv15, y1 \ dv15, ByVal 0
LineTo m_hDC, rp.Right, rp.Bottom
'MoveToEx m_hDC, x1 \ dv15, y1 \ dv15, ByVal 0
LineTo m_hDC, rp.Left, rp.Bottom
'MoveToEx m_hDC, x1 \ dv15, y1 \ dv15, ByVal 0
LineTo m_hDC, rp.Left, rp.Top




SelectObject m_hDC, oldpen
DeleteObject linepen
Else
linepen = CreatePen(m_DrawStyle, m_DrawWidth, RGBcolor)
oldpen = SelectObject(m_hDC, linepen)
nowold = SelectObject(m_hDC, m_linebrush)
MoveToEx m_hDC, x1 \ dv15, y1 \ dv15, ByVal 0
LineTo m_hDC, x2 \ dv15, y2 \ dv15
SelectObject m_hDC, nowold
SelectObject m_hDC, oldpen
DeleteObject linepen
End If
Call SetBkMode(m_hDC, m_transparent)
End If
End Sub

Public Property Get FontTransparent() As Boolean
FontTransparent = m_fonttransparent
End Property

Public Property Let FontTransparent(ByVal RHS As Boolean)
' do nothing
m_fonttransparent = RHS
End Property
Public Sub PlayEmfInside(picthis As StdPicture, x As Long, y As Long, Optional Width As Long = -1, Optional Height As Long = -1, Optional rot As Single, Optional Xs As Long, Optional Ys As Long)
    Dim kkx As RECT, oMap As Long
    Dim myXForm As XFORM
    
    If m_Status Then
    If m_hDC Then
    If Width = -1 Or Height = -1 Then
        If Width = -1 Then
            Width = picthis.Width / 1.75797922568461
            If Height = -1 Then
                Height = 1# * picthis.Height / 1.75797922568461
            Else
                Width = 1# * Height / picthis.Height * picthis.Width
            End If
        Else
                Height = 1# * picthis.Height * Width / picthis.Width
            End If
        End If
    End If
      
    
    
    GetWorldTransform m_hDC, gOldWorld
    SetGraphicsMode m_hDC, GM_ADVANCED
    oMap = GetMapMode(m_hDC)
    SetMapMode m_hDC, MM_TEXT
    If Width = 0 Or Height = 0 Then
    Else
    Dim kkx1 As Long, kky1 As Long
    kkx1 = Round((Width) / dv15)
    kky1 = Round((Height) / dv15)
    myXForm = buildXForm(x, y, rot, kkx1, kky1)
    kkx.Left = Xs - kkx1 / 2
    kkx.Top = Ys - kky1 / 2
    kkx.Right = Xs + kkx1 / 2
    kkx.Bottom = Ys + kky1 / 2
    
    
    SetWorldTransform m_hDC, myXForm
      kkx.Right = kkx.Right * 1.00166666666667
  kkx.Bottom = kkx.Bottom * 1.00333333333333
    PlayEnhMetaFile m_hDC, picthis.Handle, kkx
    End If
    
    SetWorldTransform m_hDC, gOldWorld
    SetGraphicsMode m_hDC, GM_COMPATIBLE
    SetMapMode m_hDC, oMap
    
    End If
    
    
End Sub
Public Sub PlayWmfInside(picthis As StdPicture, x As Long, y As Long, Optional Width As Long = -1, Optional Height As Long = -1, Optional rot As Single, Optional Xs As Long, Optional Ys As Long)
  Dim pHandle As Long
    
    Dim kkx As RECT, oMap As Long
    Dim myXForm As XFORM
    
    If m_Status Then
    If m_hDC Then
    If Width = -1 Or Height = -1 Then
        If Width = -1 Then
            Width = picthis.Width / 1.75797922568461
            If Height = -1 Then
                Height = 1# * picthis.Height / 1.75797922568461
            Else
                Width = 1# * Height / picthis.Height * picthis.Width
            End If
        Else
                Height = 1# * picthis.Height * Width / picthis.Width
            End If
        End If
    End If
      
     pHandle = getEmfFromWmf(picthis, Width, Height)
    
    GetWorldTransform m_hDC, gOldWorld
    SetGraphicsMode m_hDC, GM_ADVANCED
    oMap = GetMapMode(m_hDC)
    SetMapMode m_hDC, MM_TEXT
    If Width = 0 Or Height = 0 Then
    Else
    Dim kkx1 As Long, kky1 As Long
    kkx1 = Round((Width) / dv15)
    kky1 = Round((Height) / dv15)
    myXForm = buildXForm(x, y, rot, kkx1, kky1)
    kkx.Left = Xs - kkx1 / 2
    kkx.Top = Ys - kky1 / 2
    kkx.Right = Xs + kkx1 / 2
    kkx.Bottom = Ys + kky1 / 2
    
    
    SetWorldTransform m_hDC, myXForm
      kkx.Right = kkx.Right * 1.00166666666667
  kkx.Bottom = kkx.Bottom * 1.00333333333333
    PlayEnhMetaFile m_hDC, pHandle, kkx
    End If
    
    SetWorldTransform m_hDC, gOldWorld
    SetGraphicsMode m_hDC, GM_COMPATIBLE
    SetMapMode m_hDC, oMap
    
    End If
    
    DeleteMetaFile pHandle
End Sub
Private Sub PlayEmfInsideSimple(picthis As StdPicture, x As Long, y As Long, Optional Width As Long = -1, Optional Height As Long = -1)
    Dim kkx As RECT

      
    kkx.Left = x
    kkx.Top = y
    kkx.Right = Width + x
    kkx.Bottom = Height + y
   kkx.Right = kkx.Right * 1.00166666666667
  kkx.Bottom = kkx.Bottom * 1.00333333333333
    SetGraphicsMode m_hDC, GM_ADVANCED
    PlayEnhMetaFile m_hDC, picthis.Handle, kkx
SetGraphicsMode m_hDC, GM_COMPATIBLE
    
    
End Sub
Private Sub PlayWmfInsideSimple(picthis As StdPicture, x As Long, y As Long, Optional Width As Long = -1, Optional Height As Long = -1)

    Dim pHandle As Long
     pHandle = getEmfFromWmf(picthis, Width * dv15, Height * dv15)
    Dim kkx As RECT

      
    kkx.Left = x
    kkx.Top = y
    kkx.Right = Width + x
    kkx.Bottom = Height + y
   kkx.Right = kkx.Right * 1.00166666666667
  kkx.Bottom = kkx.Bottom * 1.00333333333333
    SetGraphicsMode m_hDC, GM_ADVANCED
    PlayEnhMetaFile m_hDC, pHandle, kkx
SetGraphicsMode m_hDC, GM_COMPATIBLE
    
    DeleteMetaFile pHandle
End Sub
Private Function buildXForm(ByVal x0 As Long, ByVal y0 As Long, ByVal q As Single, piw As Long, pih As Long) As XFORM
    Dim R As Single, k As Single
    q = -(CLng(q * 100) Mod 36000) * 1.74532925199433E-04

        R = Atn(piw / CSng(pih)) + Pi / 2!
        k = piw / Cos(R) / 2
        buildXForm.eM11 = Round(Cos(q), 4)
        buildXForm.eM12 = Round(Sin(q), 4)
        buildXForm.eM21 = Round(-Sin(q), 4)
        buildXForm.eM22 = Round(Cos(q), 4)
        k = Abs(k)
        buildXForm.eDx = x0 'Round((k - k * Cos(q - R)) / 2!, 0)
        buildXForm.eDy = y0 'Round((k - k * Sin(q - R)) / 2!, 0)
       

End Function
Private Function RadDeg(ByVal angle As Double) As Double
     Const pidivby180 = 1.74532925199433E-02
    RadDeg = angle * pidivby180
End Function
Public Sub PaintPicture(aPic As IPictureDisp, ByVal x1 As Single, ByVal y1 As Single, Optional ByVal Width As Single, Optional ByVal Height As Single, Optional ByVal x2 As Single, Optional ByVal y2 As Single, Optional ByVal Width2 As Single, Optional ByVal Height2 As Single, Optional ByVal Opcode As Long = 0)
If m_Status Then
    If m_hDC Then
        If IsMissing(Width) Then Width = -1
        If IsMissing(Height) Then Height = -1
    
    
    
        If Width = -1 And Height = -1 Then
                Width = Form1.ScaleX(aPic.Width, vbHimetric, vbPixels)
                Height = Form1.ScaleY(aPic.Height, vbHimetric, vbPixels)
        ElseIf Width = -1 Then
        Height = Height / dv15
            Width = 1# * Height * Form1.ScaleX(aPic.Width, vbHimetric, vbPixels) / Form1.ScaleY(aPic.Height, vbHimetric, vbPixels)
        ElseIf Height = -1 Then
        Width = Width / dv15
        Height = 1# * Width * Form1.ScaleY(aPic.Height, vbHimetric, vbPixels) / Form1.ScaleX(aPic.Width, vbHimetric, vbPixels)
        Else
        Width = Width / dv15
        Height = Height / dv15
        End If
        
        x1 = x1 / dv15
        y1 = y1 / dv15
       If aPic.Type = 4 Then
            PlayEmfInsideSimple aPic, CLng(x1), CLng(y1), CLng(Width), CLng(Height)
        ElseIf aPic.Type = 2 Then
            PlayWmfInsideSimple aPic, CLng(x1), CLng(y1), CLng(Width), CLng(Height)
       Else
            Dim ph As New cDIBSection
            ph.emfSizeFactor = 1
            ph.CreateFromPicture aPic, m_backcolor
            ph.StretchPictureH m_hDC, CLng(x1), CLng(y1), CLng(Width), CLng(Height)
       End If
    End If
End If

End Sub
  

Public Property Get DrawMode() As Long
DrawMode = mDrawMode
End Property

Public Property Let DrawMode(ByVal RHS As Long)

mDrawMode = RHS
If m_Status Then
    If m_hDC <> 0 Then
        SetROP2 m_hDC, RHS
    End If
End If

End Property
Private Function getEmfFromWmf(Wmf As StdPicture, w As Long, h As Long) As Long
 Dim zz As METAFILEPICT
 Dim bb As wmfPlaceableFileHeader
  zz.xExt = w '* dv15
 zz.yExt = h '* dv15
 
 Dim bytes As Long
Dim Buffer() As Byte
   bytes = GetMetaFileBitsEx(Wmf.Handle, bytes, ByVal 0&)
 If bytes Then
ReDim Buffer(1 To bytes) As Byte
Call GetMetaFileBitsEx(Wmf.Handle, bytes, Buffer(1))
Else

Exit Function
End If
    On Error Resume Next
    zz.mm = 1
    getEmfFromWmf = SetWinMetaFileBits(bytes, Buffer(1), Form1.hDC, zz)
End Function

Public Property Get Tag() As Variant
Tag = "0"
End Property

Public Property Let Tag(ByVal RHS As Variant)
' nothing
End Property
