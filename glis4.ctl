VERSION 5.00
Begin VB.UserControl gList 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   FillColor       =   &H80000002&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   MousePointer    =   1  'Arrow
   OLEDropMode     =   1  'Manual
   PropertyPages   =   "glis4.ctx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   7245
   Begin VB.Timer BlinkTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1590
      Top             =   855
   End
   Begin VB.Timer Timer2bar 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2745
      Top             =   2565
   End
   Begin VB.Timer Timer1bar 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1950
      Top             =   1710
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5475
      Top             =   3585
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5505
      Top             =   1035
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5940
      Top             =   2595
   End
End
Attribute VB_Name = "gList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


' use of Extender
Option Explicit
Dim waitforparent As Boolean
Dim havefocus As Boolean, UKEY$
Dim dummy As Long
Dim nopointerchange As Boolean
Dim PrevLocale As Long
Private Type Myshape
    Visible As Boolean
    hatchType As Long
    top As Long
    Left As Long
    Width As Long
    Height As Long
End Type

Private mynum$, dragslow As Long, lastshift As Integer, HandleOverride As Boolean
Private lastEditFlag As Boolean, SkipReadEditflag As Boolean
Public BypassKey As Boolean, AdjustColumns As Boolean
Public BlinkON As Boolean, prive As Long
Private mBlinkTime As Long, AdjustColumn As Integer, AdjustColumnSum As Long
Public UseTab As Boolean
Public InternalCursor As Boolean, SkipChars As Long
Public OverrideShow As Boolean
Public HideCaretOnexit As Boolean, blockheight As Boolean
Public overrideTextHeight As Long
Public AutoHide As Boolean, NoWheel As Boolean
Private missMouseClick As Boolean
Public bypassfirstClick As Boolean, Grid As Boolean, GridColor As Long
Private Shape1 As Myshape, Shape2 As Myshape, Shape3 As Myshape
Private Type RECT
        Left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type
Private Enum flagtype
    fselected = 1
    fchecked = 2
    fradiobutton = 4
    fline = 8
    joinpRevline = 16   ' if this is true means we have a join with previous row.
End Enum
Private Type itemlist
    Flags As Integer ' selected
    morerows As Integer  ' pixels??? not used now
    content As JsonObject
End Type
Private ehat$, JoinLines As Long
Private fast As Boolean
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hDC As Long, ByVal lpsz As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long, ByVal lpDrawTextParams As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (ByRef lpRect As RECT) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long


Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, RetVal As Integer)

Private Const PS_NULL = 5
Private Const PS_SOLID = 0
Public restrictLines As Long, UseHeaderOnly As Boolean
Private nowX As Single, nowY As Single
Private marvel As Boolean
Private Const DT_BOTTOM As Long = &H8&
Private Const DT_CALCRECT As Long = &H400&
Private Const DT_CENTER As Long = &H1&
Private Const DT_EDITCONTROL As Long = &H2000&
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_EXPANDTABS As Long = &H40&
Private Const DT_EXTERNALLEADING As Long = &H200&
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const DT_INTERNAL As Long = &H1000&
Private Const DT_LEFT As Long = &H0&
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_NOCLIP As Long = &H100&
Private Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
Private Const DT_NOPREFIX As Long = &H800&
Private Const DT_PATH_ELLIPSIS As Long = &H4000&
Private Const DT_PREFIXONLY As Long = &H200000
Private Const DT_RIGHT As Long = &H2&
Private Const DT_SINGLELINE As Long = &H20&
Private Const DT_TABSTOP As Long = &H80&
Private Const DT_TOP As Long = &H0&
Private Const DT_VCENTER As Long = &H4&
Private Const DT_WORDBREAK As Long = &H10&
Private Const DT_WORD_ELLIPSIS As Long = &H40000

Const m_def_Text = vbNullString
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = 0
Const m_def_Enabled = False
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_dcolor = &H333333
Const m_def_CapColor = &HAAFFBB
Const m_def_Showbar = True
Const m_def_sync = vbNullString

Dim m_sync As String
Dim m_BackColor As Long
Dim m_ForeColor As Long
'Dim m_Enabled As Boolean
Dim m_font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_CapColor As Long
Dim m_dcolor As Long

Dim m_showbar As Boolean
Dim mParts() As Long
Dim mTabs As Long, mLeftTab As Long, mCurTab As Long, TwipsCurTab As Long
Dim mList() As itemlist
Dim topitem As Long
Dim itemcount As Long
Dim Mselecteditem As Long, dragfocus As Boolean
Event RealCurReplace(a$)
Event DragOverCursor(ok As Boolean)
Event OnResize()
Event selected(item As Long)
Event SelectedMultiAdd(item As Long)
Event SelectedMultiSub(item As Long)
Event Selected2(item As Long)
Event softSelected(item As Long)
Event ValidMove(X As Single, y As Single)
Event Maybelanguage()
Event MouseUp(X As Single, y As Single)
Event SpecialColor(RGBcolor As Long)
Event RemoveOne(that As String)
Event PushMark2Undo(that As String)
Event PushUndoIfMarked()
Event addone(that As String)
Event getpair(a$, b$)
Event MayRefresh(ok As Boolean)
Event CheckGotFocus()
Event CheckLostFocus()
Event DragData(ThatData As String)
Event DragPasteData(ThatData As String)
Event DropOk(ok As Boolean)
Event DropFront(ok As Boolean)
Event ScrollMove(item As Long)
Event RefreshDesktop()
Event PrepareContainer()
Event NeedDoEvents()
Event OutPopUp(X As Single, y As Single, myButton As Integer)
Event SplitLine()
Event LineUp()
Event LineDown()
Event PureListOn()
Event PureListOff()
Event HaveMark(Yes As Boolean)
Event GroupUndo()
Event MarkCut(ThatData As String)
Event markin()
Event MarkOut()
Event MarkDestroyAny()
Event MarkDestroy()
Event MarkDelete(preservecursor As Boolean)
Event WordMarked(ThisWord As String)
Event ShowExternalCursor()
Event ChangeSelStart(thisselstart As Long)
Event ReadListItem(item As Long, content As String)
Event ChangeListItem(item As Long, content As String)
Event HeaderSelected(Button As Integer)
Event BlockCaret(item As Long, blockme As Boolean, skipme As Boolean)
Event ScrollSelected(item As Long, y As Long)
Event MenuChecked(item As Long)
Event PromptLine(ThatLine As Long)
Event PanLeftRight(direction As Boolean)
Event GetBackPicture(pic As Object)
Event KeyDown(KeyCode As Integer, shift As Integer)
Event KeyDownAfter(KeyCode As Integer, shift As Integer)
Event SyncKeyboard(item As Integer)
Event SyncKeyboardUnicode(a$)
Event PreviewKeyboardUnicode(ByRef a$)
Event Find(Key As String, where As Long, skip As Boolean)
Event ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Event ExposeRectCol(ByVal item As Long, ByVal Col As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Event ExposeListcount(cListCount As Long)
Event ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal X As Long, ByVal y As Long)
Event GetRealX1(mHdc As Long, ByVal ExtSelStart As Long, ByVal that$, retvalue As Long, ok As Boolean)
Event MouseMove(Button As Integer, shift As Integer, X As Single, y As Single)
Event SpinnerValue(ThatValue As Long)
Event RegisterGlist(this As gList)
Event UnregisterGlist()
Event DeployMenu()
Event CascadeSelect(item As Long) ' 1 based
Event BlinkNow(Face As Boolean)
Event CtrlPlusF1()
Event EnterOnly()
Event RefreshOnly()
Event CorrectCursorAfterDrag()
Event DragOverNow(a As Boolean)
Event DragOverDone(a As Boolean)
Event HeadLineChange(a$)
Event AddSelStart(val As Long, shift As Integer)
Event SubSelStart(val As Long, shift As Integer)
Event rtl(thisHDC As Long, item As Long, where As Long, mark10 As Long, mark20 As Long, Offset As Long)
'Event RTL2(s$, where As Long, mark10 As Long, mark20 As Long, Offset As Long)
Event DelExpandSS()
Event SetExpandSS(val As Long)
Event ExpandSelStart(val As Long)
Event Fkey(a As Integer)
Event CaretDeal(Deal As Long)
Event AccKey(m As Long)
Event ReadColumnItem(item As Long, Col As Long, content As String)
Event ReadColumnProp(item As Long, Col As Long, Prop As String, content As String)
Event ReadColumnPropNum(item As Long, Col As Long, Prop As String, content)
Event ChangeColumnListItem(item As Long, Col As Long, content As String)
Event WindowKey(VbKeyThis As Integer)

Private state As Boolean
Private secreset As Boolean
Private scrollme As Long
Private lY As Long, dr As Boolean
Private drc As Boolean
Private scrTwips As Single
Private cy As Long
Private cx As Long
Dim myt As Long, FaceBlink As Boolean
Dim mytPixels As Long
' for not moving rows/columns
Public TopRows As Long, LeftColumns As Long
Public NoMoveDrag As Boolean, NarrowSelect As Boolean
Public BarColor As Long
Public BarHatch As Long
Public BarHatchColor As Long
Public LeaveonChoose As Boolean
Public BypassLeaveonChoose As Boolean
Public LastSelected As Long
Public NoPanLeft As Boolean
Public NoPanRight As Boolean
Private LastVScroll As Long
Public FreeMouse As Boolean
Public NoCaretShow As Boolean
Public NoBarClick As Boolean
Public NoEscapeKey As Boolean
Public InfoDropBarClick As Boolean
Dim valuepoint As Long, minimumWidth As Long
Dim mValue As Long, mmax As Long, mmin As Long, mLargeChange As Long  ' min 1
Dim mSmallChange As Long  ' min 1
Dim mVertical As Boolean

Dim OurDraw As Boolean, GetOpenValue As Long
Dim lastX As Single, LastY As Single

Private mjumptothemousemode As Boolean
Private mpercent As Single
Private BarWidth As Long
Private NoFire As Boolean
Public addpixels As Long
Public StickBar As Boolean
Dim Hidebar As Boolean
Dim myEnabled As Boolean
Public WrapText As Boolean
Public CenterText As Boolean, RightText As Boolean
Public VerticalCenterText As Boolean
Private mHeadline As String
Private mHeadlineHeight As Long
Private mHeadlineHeightTwips As Long
Public MultiSelect As Boolean
Public LeftMarginPixels As Long
Dim Buffer As Long
Public FloatList As Boolean
Public MoveParent As Boolean
Public BlockItemcount As Boolean
Private useFloatList As Boolean
Public NoVerMove As Boolean, NoHorMove As Boolean
Public HeadLineHeightMinimum As Long
Private mPreserveNpixelsHeaderRight As Long
Public AutoPanPos As Boolean   ' used if we have no EditFlag
Public FloatLimitLeft As Long
Public FloatLimitTop As Long
Public mEditFlag As Boolean
Public mSortstyle As Integer
Public SingleLineSlide As Boolean
Private mSelstart As Long
Private caretCreated As Boolean
Public MultiLineEditBox As Boolean
Public NoScroll As Boolean
Public MarkNext As Long  ' 0 - markin, 1- Markout
Public Noflashingcaret As Boolean
Public NoFreeMoveUpDown As Boolean  ' if true then keyup and keydown scroll up down the list
Public PromptLineIdent As Long ' to be a console we need prompt line to have some chars untouch perhaps this ">"
Public FadeLastLinePart As Long ' if is zero then no use at all
Public LastLinePart As String
Public Spinner As Boolean ' if true and restrictline =1 - we have events for up down values
Public maxchar As Long ' for non multiline
Public WordCharLeft As String
Public WordCharRight As String
Public WordCharRightButIncluded As String
Public WordCharLeftButIncluded As String
Public DropEnabled As Boolean
Public DragEnabled As Boolean, NoArrowUp As Boolean, NoArrowDown As Boolean, Arrows2Tab As Boolean
Public NoHeaderBackground As Boolean
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&) ' not NT?
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Const DWL_ANYTHREAD& = 0
Const LOCALE_ILANGUAGE = 1
Private Declare Function PeekMessageW Lib "user32" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Const WM_KEYFIRST = &H100
Const WM_CHAR = &H102
Const WM_KEYLAST = &H108
Private Type POINTAPI
    X As Long
    y As Long
End Type
Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private timestamp As Double, timestamp1 As Double
Private doubleclick As Long, preservedoubleclick As Long
Private dbLx As Long, dbly As Long
Private PX As Long, PY As Long
Public SkipForm As Boolean
Public DropKey As Boolean
Public MenuGroup As String
Private mTabStop As Boolean
Private oldpointer As Integer
Private himc As Long
Private Const VK_PROCESSKEY = &HE5
Private Const GCS_COMPREADSTR = &H1
Private Const GCS_RESULTSTR = &H800
Private Const GCS_COMPSTR = &H8
Private leave As Long
Private Type DRAWTEXTPARAMS
     cbSize As Long
     iTabLength As Long
     iLeftMargin As Long
     iRightMargin As Long
     uiLengthDrawn As Long
End Type
Dim tParam As DRAWTEXTPARAMS
Public SuspDraw As Boolean
Private Type TEXTMETRICW
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Integer
        tmLastChar As Integer
        tmDefaultChar As Integer
        tmBreakChar As Integer
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type
Private TM As TEXTMETRICW
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, lpMetrics As TEXTMETRICW) As Long
Private LastNumX As Boolean
Public Function GetLastKeyPressed() As Long
Dim Message As Msg
If mynum$ <> "" Then
    GetLastKeyPressed = -1
ElseIf PeekMessageW(Message, 0, WM_CHAR, WM_KEYLAST, 0) Then
    GetLastKeyPressed = Message.wParam
Else
    GetLastKeyPressed = -1
End If
End Function
Public Property Let HeadlineHeight(ByVal RHS As Long)
If HeadLine <> "" Then
mHeadlineHeight = RHS
mHeadlineHeightTwips = CLng(RHS * scrTwips)

Else
mHeadlineHeight = 0
mHeadlineHeightTwips = 0

End If
End Property

Public Property Get HeadlineHeightTwips() As Long
'' for dynamic controls
If HeadLine <> "" Then
HeadlineHeightTwips = CLng(mHeadlineHeight * scrTwips)
Else
HeadlineHeightTwips = myt

End If
End Property
Public Property Get HeadlineHeight() As Long
If HeadLine <> "" Then
HeadlineHeight = mHeadlineHeight
Else
HeadlineHeight = 0

End If
End Property

Public Property Let HeadLine(ByVal RHS As String)
If mHeadline = vbNullString Then
' reset headlineheight
RaiseEvent HeadLineChange(RHS)
mHeadline = RHS
HeadlineHeight = UserControlTextHeightPixels()

Exit Property

End If
mHeadline = RHS
End Property

Public Property Get HeadLine() As String
HeadLine = mHeadline
End Property
Public Sub PrepareToShow(Optional Delay As Single = 10)
 BarWidth = UserControlTextWidth("W")
 CalcAndShowBar1
Timer1.enabled = False
If Delay < 1 Then Delay = 1
If fast Then
fast = False
Timer1.Interval = Delay
Else
Timer1.Interval = Delay * 5
End If
Timer1.enabled = True
End Sub
Public Sub PressSoft()
secreset = False
RaiseEvent Selected2(SELECTEDITEM - 1)
End Sub
Public Property Get ScrollFrom() As Long
    ScrollFrom = topitem
End Property
Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal RHS As Integer)
    m_BorderStyle = RHS
    
 If BackStyle = 0 Then UserControl.BorderStyle = -(m_BorderStyle <> 0) Else UserControl.BorderStyle = 0
    PropertyChanged "BorderStyle"
End Property
Public Property Get sync() As String
    sync = m_sync
End Property
Public Property Let sync(ByVal New_sync As String)
    If Ambient.UserMode Then Err.Raise 393
    m_sync = New_sync
    PropertyChanged "sync"
End Property
Public Property Get hWnd() As Long
hWnd = UserControl.hWnd
End Property
Public Property Let Text(ByVal new_text As String)
Clear True
If mParts(1) = 0 Then mParts(1) = Me.WidthPixels
If new_text <> "" Then
    If Right$(new_text, 2) <> vbCrLf And new_text <> "" Then
        new_text = new_text + vbCrLf
    End If
    Dim mpos As Long, b$
    Do
    b$ = GetStrUntilB(mpos, vbCrLf, new_text)
    additemFast b$  ' and blank lines
    Loop Until mpos > Len(new_text) Or mpos = 0
End If
If GetTopUserControl(Me).Ambient.UserMode = False Then
    Repaint
    SELECTEDITEM = 0
    CalcAndShowBar
    ShowMe
End If

PropertyChanged "Text"
End Property
Public Sub AddTextColumn(ByVal thiscolumn As Long, ByVal new_text As String)
If thiscolumn < 1 Then thiscolumn = 1
If thiscolumn > mTabs Then mTabs = thiscolumn: ReDim Preserve mParts(1 To mTabs)
If new_text <> "" Then
    If Right$(new_text, 2) <> vbCrLf And new_text <> "" Then
        new_text = new_text + vbCrLf
    End If
    Dim mpos As Long, b$, i As Long
    
    i = 0
    Do
        If i < itemcount Then
            With mList(i)
                If .content Is Nothing Then Set .content = New JsonObject
                .content.AssignPath "C." & thiscolumn, GetStrUntilB(mpos, vbCrLf, new_text)
                .Flags = 0
            End With
            i = i + 1
        End If
    Loop Until mpos > Len(new_text) Or mpos = 0 Or i > itemcount
End If
If GetTopUserControl(Me).Ambient.UserMode = False Then
    Repaint
    SELECTEDITEM = 0
    CalcAndShowBar
    ShowMe
End If
End Sub
Public Property Let ListText(ByVal new_text As String)
Clear True
If Right$(new_text, 2) <> vbCrLf And new_text <> "" Then
new_text = new_text + vbCrLf
End If
Dim mpos As Long, b$
Do
b$ = GetStrUntilB(mpos, vbCrLf, new_text)

If Left$(b$, 1) <> "_" Then
additemFast b$
Else
b$ = Mid$(b$, 2)
If b$ = vbNullString Then
AddSep
Else
additemFast b$
menuEnabled(itemcount - 1) = False
End If
End If
Loop Until mpos > Len(new_text) Or mpos = 0
Repaint
SELECTEDITEM = 0
CalcAndShowBar
ShowMe
End Property
Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
Dim i As Long, Pad$
Text = space$(500)
RaiseEvent PureListOn

Dim thiscur, l As Long
thiscur = 1

For i = 0 To listcount - 1
Pad$ = list(i) + vbCrLf
l = Len(Pad)
If Len(Text) < thiscur + l Then Text = Text + space$((thiscur + l) + 100)
Mid$(Text, thiscur, l) = Pad$
thiscur = thiscur + l
Next i
Text = Left$(Text, thiscur - 1)

RaiseEvent PureListOff
End Property
Public Property Get TextAtColumn(ByVal RHS As Long) As String
Dim i As Long, Pad$
If RHS < 1 Then RHS = 1
If RHS > mTabs Then RHS = mTabs
TextAtColumn = space$(500)
RaiseEvent PureListOn
Dim thiscur, l As Long
thiscur = 1
For i = 0 To listcount - 1
Pad$ = listAtColumn(i, RHS) + vbCrLf
l = Len(Pad)
If Len(TextAtColumn) < thiscur + l Then TextAtColumn = TextAtColumn + space$((thiscur + l) + 100)
Mid$(TextAtColumn, thiscur, l) = Pad$
thiscur = thiscur + l
Next i
TextAtColumn = Left$(TextAtColumn, thiscur - 1)
RaiseEvent PureListOff
End Property

Public Sub ScrollToTextEdit(ThatTopItem As Long, Optional this As Long = -2)
On Error GoTo scroend
topitem = ThatTopItem
If topitem < 0 Then topitem = 0
If this > -2 Then
SELECTEDITEM = this
End If
CalcAndShowBar1
Timer1.enabled = True
scroend:
End Sub
Public Sub ScrollTo(ThatTopItem As Long, Optional this As Long = -2)
On Error GoTo scroend

If ThatTopItem + lines >= listcount Then
        If ThatTopItem - lines < 0 Then topitem = 0 Else topitem = listcount - lines - 1
Else
topitem = ThatTopItem
End If
If topitem < 0 Then topitem = 0
If this > -2 Then
SELECTEDITEM = this
End If
CalcAndShowBar1
Timer1.enabled = True
scroend:
End Sub
Public Sub ScrollToSilent(ThatTopItem As Long, Optional this As Long = -2)
On Error GoTo scroend
topitem = ThatTopItem
If topitem < 0 Then topitem = 0
If this > -2 Then
SELECTEDITEM = this
End If
If BarVisible Then Redraw ShowBar
Timer1.enabled = True
scroend:
End Sub
Public Sub CalcAndShowBar()
CalcAndShowBar1
ShowMe2
End Sub
Private Sub CalcAndShowBar1()
Dim OldValue As Long, oldmax As Long
OldValue = Value
oldmax = Max
On Error GoTo calcend
state = True

   On Error Resume Next
            Err.Clear
    If Not Spinner Then
            If listcount - 1 - lines < 1 Then
            Max = 1
            Else
            Max = listcount - 1 - lines
            largechange = lines
            End If
            If Err.Number > 0 Then
                Value = listcount - 1
                Max = listcount - 1
            End If
                      Value = topitem
        End If

state = False
If listcount < lines + 2 Then
BarVisible = False
Else
Redraw Hidebar

End If
calcend:
End Sub
Public Property Get ListValue() As String
' this was text before
RaiseEvent PureListOn
If SELECTEDITEM <= 0 Then Else ListValue = list(ListIndex)
RaiseEvent PureListOff
End Property

Public Property Get listcount() As Long
Dim thatlistcount As Long
RaiseEvent ExposeListcount(thatlistcount)
If thatlistcount > 0 Then
listcount = thatlistcount
Else
  listcount = itemcount
  End If
End Property
Public Property Let ShowBar(ByVal RHS As Boolean)
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
mytPixels = CLng(myt / scrTwips)
myt = CLng(mytPixels * scrTwips)
Else
mytPixels = (UserControlTextHeightPixels() + addpixels)
myt = mytPixels * scrTwips

End If


    m_showbar = RHS
    BarWidth = UserControlTextWidth("W")
    
    state = True
    Value = 0
    state = False
 
    If listcount >= lines Then
BarVisible = (m_showbar Or StickBar Or AutoHide) Or Hidebar
Else
Redraw (m_showbar Or StickBar Or AutoHide) Or Hidebar
End If
   
'RepaintScrollBar
End Property
Public Property Get ShowBar() As Boolean
If Hidebar Then
ShowBar = True ' TEMPORARY USE
Else

    ShowBar = m_showbar Or StickBar Or AutoHide
    End If
End Property

Public Property Let BackColor(ByVal RHS As OLE_COLOR)

    m_BackColor = RHS
UserControl.BackColor = RHS
  PropertyChanged "BackColor"
    
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor 'UserControl.Backcolor
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal RHS As OLE_COLOR)
    m_ForeColor = RHS
    UserControl.ForeColor = Abs(RHS)
    PropertyChanged "ForeColor"
End Property
Public Property Get CapColor() As OLE_COLOR
   CapColor = m_CapColor
End Property

Public Property Let CapColor(ByVal RHS As OLE_COLOR)
    m_CapColor = RHS
    PropertyChanged "CapColor"
End Property
Public Property Get dcolor() As OLE_COLOR
    dcolor = m_dcolor
  
End Property

Public Property Let dcolor(ByVal RHS As OLE_COLOR)
    m_dcolor = RHS
    PropertyChanged "dcolor"
End Property
Public Property Get enabled() As Boolean
    enabled = myEnabled
End Property
Public Property Let enabled(ByVal RHS As Boolean)
    myEnabled = RHS
 
    PropertyChanged "Enabled"
    On Error Resume Next
    If Not waitforparent Then Exit Property
    Extender.TabStop = TabStopSoft And RHS
End Property
Public Property Let TabStop(ByVal RHS As Boolean)
    On Error Resume Next
    mTabStop = RHS
    If Not waitforparent Then Exit Property
    Extender.TabStop = RHS
End Property
Public Property Let TabStopSoft(ByVal RHS As Boolean)
    mTabStop = RHS
    TabStop = RHS
End Property

Public Property Get TabStopSoft() As Boolean
    TabStopSoft = mTabStop
End Property
Public Property Get AveCharWith() As Long
GetTextMetrics UserControl.hDC, TM
AveCharWith = TM.tmAveCharWidth
End Property
Public Property Get Font() As Font
    Set Font = m_font
End Property

Public Function CloneFont(Font As IFont) As StdFont
  Font.Clone CloneFont
End Function
Public Property Set Font(New_Font As Font)
    Set m_font = New_Font
    Set UserControl.Font = m_font
    GetTextMetrics UserControl.hDC, TM
    If restrictLines > 0 Then
        myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
        mytPixels = myt / scrTwips
        myt = mytPixels * scrTwips
    Else
        mytPixels = CLng((UserControlTextHeightPixels() + addpixels))
        myt = CLng(mytPixels * scrTwips)
    End If
    HeadlineHeight = UserControlTextHeightPixels()
    PropertyChanged "Font"
End Property
Public Sub CalcNewFont()
GetTextMetrics UserControl.hDC, TM
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
Else
mytPixels = CLng((UserControlTextHeightPixels() + addpixels))
myt = CLng(mytPixels * scrTwips)
End If
HeadlineHeight = UserControlTextHeightPixels()

If ListIndex >= 0 Then
CalcAndShowBar1
    ShowThis ListIndex + 1
Else
    ShowMe True
End If

End Sub

Public Property Get FontSize() As Single

  FontSize = m_font.Size
 
End Property

Public Property Let FontSize(New_FontSize As Single)
     If New_FontSize < 6 Then
  m_font.Size = 6
     Else
m_font.Size = New_FontSize
End If
GetTextMetrics UserControl.hDC, TM
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
mytPixels = CLng(myt / scrTwips)
myt = CLng(mytPixels * scrTwips)
Else

mytPixels = (UserControlTextHeightPixels() + addpixels)
myt = mytPixels * scrTwips

End If


End Property

Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle

End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
On Error Resume Next
    m_BackStyle = New_BackStyle
    If m_BackStyle = 0 Then UserControl.BorderStyle = -(m_BorderStyle <> 0) Else UserControl.BorderStyle = 0
    PropertyChanged "BackStyle"
 
End Property



Private Sub usercontrol_GotFocus1()
Dim YYT As Long
YYT = myt
DrawMode = vbCopyPen
If SELECTEDITEM > 0 Then
If SELECTEDITEM - topitem - 1 <= lines Then
If BackStyle = 1 Then

Line (scrollme + scrTwips, (SELECTEDITEM - topitem) * YYT)-(scrollme + UserControl.Width, (SELECTEDITEM - topitem - 1) * YYT), 0, B

Else
Line (scrollme, (SELECTEDITEM - topitem) * YYT)-(scrollme + UserControl.Width, (SELECTEDITEM - topitem - 1) * YYT), 0, B


End If
End If
End If
DrawMode = vbCopyPen
Timer1.Interval = 40
Timer1.enabled = True
End Sub

Public Sub LargeBar1KeyDown(KeyCode As Integer, shift As Integer)
Timer1.enabled = False
If ListIndex < 0 Then
Else
PressKey KeyCode, shift
End If
End Sub

Private Sub BlinkTimer_Timer()
If mBlinkTime = 0 Then BlinkON = False
If BlinkON Then
    BlinkTimer.Interval = mBlinkTime
    FaceBlink = Not FaceBlink
    RaiseEvent BlinkNow(FaceBlink)
    ShowPan
Else
    BlinkTimer.enabled = False
End If

End Sub

Private Sub Timer1bar_Timer()
processXY lastX, LastY
End Sub

Private Sub timer2bar_Timer()
If m_showbar Or Shape1.Visible Or Spinner Then Redraw
On Error Resume Next
If Me.Parent.Visible = False Then Timer2bar.enabled = False
End Sub
Public Sub GiveSoftFocus()
RaiseEvent CheckGotFocus
havefocus = True
SoftEnterFocus
If Not NoWheel And Not UseHeaderOnly Then RaiseEvent RegisterGlist(Me)
End Sub

Private Sub UserControl_GotFocus()
RaiseEvent CheckGotFocus
havefocus = True
dragfocus = False
SoftEnterFocus
If Not NoWheel And Not UseHeaderOnly Then RaiseEvent RegisterGlist(Me)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
On Error GoTo fin
If BypassKey Then KeyAscii = 0: Exit Sub
If DropKey Then KeyAscii = 0: Exit Sub
Dim bb As Boolean, kk$, pair$, b1 As Boolean
If ListIndex < 0 Then
    If KeyAscii = 13 Then RaiseEvent EnterOnly
    If ParentPreview Then
        If UKEY$ <> "" Then
            kk$ = UKEY$
            UKEY$ = vbNullString
        Else
            kk$ = GetKeY(KeyAscii)
        End If
        kk$ = GetKeY(KeyAscii)
        RaiseEvent PreviewKeyboardUnicode(kk$)
        If Len(kk$) = 0 Then KeyAscii = 0: Exit Sub
    End If
Else
    If Not state Then
        If KeyAscii = 13 And myEnabled And Not MultiLineEditBox Then
            KeyAscii = 0
            If SELECTEDITEM > 0 Then
                secreset = False
                RaiseEvent Selected2(SELECTEDITEM - 1)
            End If
            RaiseEvent PreviewKeyboardUnicode(Chr$(13))
        ElseIf KeyAscii = 27 Then  ' can be used if not enabled...to quit
            KeyAscii = 0
            If Not NoEscapeKey Then
                SELECTEDITEM = -1
                secreset = False
                RaiseEvent Selected2(-2)
            ElseIf Not LeaveonChoose Then
                RaiseEvent PreviewKeyboardUnicode(Chr$(27))
            End If
        Else
            If myEnabled Then
                If mTabs = 1 Then
                    If maxchar = 0 Or (maxchar > Len(list(SELECTEDITEM - 1)) Or MultiLineEditBox) Then
                        RaiseEvent SyncKeyboard(KeyAscii)
                        If KeyAscii = 9 And UseTab Then
                            If Len(UKEY$) = 0 Then
                                KeyAscii = 0
                            Else
                                RaiseEvent KeyDown(KeyAscii, lastshift)
                                If Not KeyAscii = 0 Then GetKeY2 KeyAscii, lastshift
                            End If
                            If KeyAscii = 0 Then Exit Sub
                        End If
                        If ((KeyAscii = 9 And UseTab) Or (KeyAscii > 31) And SELECTEDITEM > 0) Then
                            If EditFlagSpecial Then
                                bb = enabled
                                enabled = False
                                RaiseEvent HaveMark(b1)
                                RaiseEvent PushUndoIfMarked
                                RaiseEvent MarkDelete(False)
                                If b1 Then RaiseEvent GroupUndo
                                enabled = bb
                            End If
                            If EditFlagSpecial And ((KeyAscii > 32 And KeyAscii <> 127) Or (KeyAscii = 9 And UseTab)) Then
                                If UKEY$ <> "" Then
                                    kk$ = UKEY$
                                    UKEY$ = vbNullString
                                Else
                                    kk$ = GetKeY(KeyAscii)
                                End If
                                If ParentPreview Then RaiseEvent PreviewKeyboardUnicode(kk$)
                                RaiseEvent getpair(kk$, pair$)
                                If Len(kk$) = 0 Then Exit Sub
                                If SelStart = 0 Then mSelstart = 1
                                If pair$ <> "" Then
                                    If b1 Then
                                        RaiseEvent MarkCut(ehat$)
                                        If InStr(ehat$, Chr(13)) > 0 Then ehat$ = Left$(ehat$, InStr(ehat$, Chr(13)) - 1)
                                            kk$ = kk$ + ehat$ + pair$
                                        Else
                                            kk$ = kk$ + pair$
                                        End If
                                    End If
                                    If AscW(kk$) = 13 Then
                                    Exit Sub
                                End If
                                RaiseEvent RemoveOne(kk$)
                                If KeyAscii = 44 And Len(kk$) = 2 Then
                                    RaiseEvent SetExpandSS(mSelstart + 2)
                                    SelStartEventAlways = SelStart + 2
                                    RaiseEvent PureListOn
                                    pair$ = list(SELECTEDITEM - 1)
                                    list(SELECTEDITEM - 1) = Left$(pair$, SelStart - 3) + kk$ + Mid$(pair$, SelStart - 2)
                                    RaiseEvent PureListOff
                                Else
                                    RaiseEvent SetExpandSS(mSelstart + 1)
                                    SelStartEventAlways = SelStart + 1
                                    RaiseEvent DelExpandSS
                                    RaiseEvent PureListOn
                                    pair$ = list(SELECTEDITEM - 1)
                                    list(SELECTEDITEM - 1) = Left$(pair$, SelStart - 2) + kk$ + Mid$(pair$, SelStart - 1)
                                    RaiseEvent PureListOff
                                End If
                                RaiseEvent SetExpandSS(mSelstart)
                            Else
                                If UKEY$ <> "" Then
                                    kk$ = UKEY$
                                    UKEY$ = vbNullString
                                Else
                                    kk$ = GetKeY(KeyAscii)
                                End If
                                If ParentPreview Then RaiseEvent PreviewKeyboardUnicode(kk$)
                                RaiseEvent SyncKeyboardUnicode(kk$)
                            End If
                        Else
                            If ParentPreview Then RaiseEvent PreviewKeyboardUnicode(Chr$(KeyAscii))
                        End If
                    End If
                Else
                    If maxchar = 0 Or (maxchar > Len(listAtColumn(SELECTEDITEM - 1, mCurTab)) Or MultiLineEditBox) Then
                        RaiseEvent SyncKeyboard(KeyAscii)
                        If KeyAscii = 9 And UseTab Then
                            If Len(UKEY$) = 0 Then
                                KeyAscii = 0
                            Else
                                RaiseEvent KeyDown(KeyAscii, lastshift)
                                If Not KeyAscii = 0 Then GetKeY2 KeyAscii, lastshift
                            End If
                            If KeyAscii = 0 Then Exit Sub
                        End If
                        If ((KeyAscii = 9 And UseTab) Or (KeyAscii > 31) And SELECTEDITEM > 0) Then
                            If EditFlagSpecial Then
                                bb = enabled
                                enabled = False
                                RaiseEvent HaveMark(b1)
                                RaiseEvent PushUndoIfMarked
                                RaiseEvent MarkDelete(False)
                                If b1 Then RaiseEvent GroupUndo
                                enabled = bb
                            End If
                            If EditFlagSpecial And ((KeyAscii > 32 And KeyAscii <> 127) Or (KeyAscii = 9 And UseTab)) Then
                                If UKEY$ <> "" Then
                                    kk$ = UKEY$
                                    UKEY$ = vbNullString
                                Else
                                    kk$ = GetKeY(KeyAscii)
                                End If
                                If ParentPreview Then RaiseEvent PreviewKeyboardUnicode(kk$)
                                RaiseEvent getpair(kk$, pair$)
                                If Len(kk$) = 0 Then Exit Sub
                                If SelStart = 0 Then mSelstart = 1
                                If pair$ <> "" Then
                                    If b1 Then
                                        RaiseEvent MarkCut(ehat$)
                                        If InStr(ehat$, Chr(13)) > 0 Then ehat$ = Left$(ehat$, InStr(ehat$, Chr(13)) - 1)
                                            kk$ = kk$ + ehat$ + pair$
                                        Else
                                            kk$ = kk$ + pair$
                                        End If
                                    End If
                                    If AscW(kk$) = 13 Then
                                    Exit Sub
                                End If
                                RaiseEvent RemoveOne(kk$)
                                If KeyAscii = 44 And Len(kk$) = 2 Then
                                    RaiseEvent SetExpandSS(mSelstart + 2)
                                    SelStartEventAlways = SelStart + 2
                                    RaiseEvent PureListOn
                                    pair$ = listAtColumn(SELECTEDITEM - 1, mCurTab)
                                    listAtColumn(SELECTEDITEM - 1, mCurTab) = Left$(pair$, SelStart - 3) + kk$ + Mid$(pair$, SelStart - 2)
                                    RaiseEvent PureListOff
                                Else
                                    RaiseEvent SetExpandSS(mSelstart + 1)
                                    SelStartEventAlways = SelStart + 1
                                    RaiseEvent DelExpandSS
                                    RaiseEvent PureListOn
                                    pair$ = listAtColumn(SELECTEDITEM - 1, mCurTab)
                                    listAtColumn(SELECTEDITEM - 1, mCurTab) = Left$(pair$, SelStart - 2) + kk$ + Mid$(pair$, SelStart - 1)
                                    RaiseEvent PureListOff
                                End If
                                RaiseEvent SetExpandSS(mSelstart)
                            Else
                                If UKEY$ <> "" Then
                                    kk$ = UKEY$
                                    UKEY$ = vbNullString
                                Else
                                    kk$ = GetKeY(KeyAscii)
                                End If
                                If ParentPreview Then RaiseEvent PreviewKeyboardUnicode(kk$)
                                RaiseEvent SyncKeyboardUnicode(kk$)
                            End If
                        Else
                            If ParentPreview Then RaiseEvent PreviewKeyboardUnicode(Chr$(KeyAscii))
                        End If
                    End If
                End If
             End If
        End If
    End If
fin:
    KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If BlinkON Then
    If Not BlinkTimer.enabled Then
        BlinkTimer.Interval = mBlinkTime
        BlinkTimer.enabled = True
        FaceBlink = Not FaceBlink
        RaiseEvent BlinkNow(FaceBlink)
        ShowPan
    End If
ElseIf useFloatList And FloatList Then
    Timer1.enabled = False
    If TypeOf Extender.Container Is GuiM2000 Then
        RaiseEvent RefreshDesktop
    Else
        RaiseEvent PrepareContainer
    End If
    
Else
    Timer1.enabled = False
    Timer1.Interval = 30
    If Not enabled Then Exit Sub

    If listcount > 0 Or MultiLineEditBox And Not UseHeaderOnly Then
        If OverrideShow And Not HandleOverride Then
            ShowMe
        Else
            ShowMe2
        End If
        HandleOverride = False
    Else
        ShowMe headeronly:=UseHeaderOnly
    End If

    Refresh
End If
End Sub

Private Sub Timer2_Timer()
If drc Then
If topitem > 0 Then
topitem = topitem - 1
 SELECTEDITEM = topitem + 1

Timer1.Interval = 0
Timer1.Interval = 100
  Timer1.enabled = True
 End If
Else
If topitem + 1 < listcount - lines Then
topitem = topitem + 1
 If topitem + lines + 1 <= listcount Then SELECTEDITEM = topitem + lines + 1
Timer1.Interval = 0
Timer1.Interval = 100
  Timer1.enabled = True
  End If
End If
state = True
 On Error Resume Next
 Err.Clear

    If SELECTEDITEM >= listcount Then
 Value = listcount - 1
  state = False
  Exit Sub
        Else
    Value = topitem
    End If
    state = False
 If Timer2.enabled = False Then
If SELECTEDITEM - topitem > 0 And SELECTEDITEM - topitem - 1 <= lines And cx > 0 And cx < UserControl.ScaleWidth Then
 If SELECTEDITEM > 0 Then
         If Not BlockItemcount Then
             REALCUR SELECTEDITEM - 1, cx - scrollme, dummy, mSelstart
             mSelstart = mSelstart + 1
RaiseEvent ChangeSelStart(mSelstart)
             End If
 RaiseEvent selected(SELECTEDITEM)
 End If
 End If
 Else
 Timer3.enabled = True
 End If
End Sub





Private Sub Timer3_Timer()
Timer3.enabled = False
DOT3
End Sub
Private Sub DOT3()
If SELECTEDITEM > listcount Then
Timer3.enabled = False
Exit Sub
End If
If SELECTEDITEM > 0 Then
' why???
'ShowMe2
RaiseEvent ScrollSelected(SELECTEDITEM, cy * myt)

End If
End Sub


Public Sub SoftEnterFocus()

If bypassfirstClick Then
missMouseClick = True
FreeMouse = True
End If
state = Not enabled
Noflashingcaret = Not enabled
If EditFlagSpecial Then
If Not Spinner Then state = Not MultiLineEditBox
End If
RaiseEvent ShowExternalCursor
If Not Timer1.enabled Then PrepareToShow 5
End Sub

Private Sub SoftExitFocus()
If Not havefocus Then Exit Sub
Noflashingcaret = True
state = True ' no keyboard input

secreset = False
Timer2.enabled = False
FreeMouse = False
If (Not BypassLeaveonChoose) And LeaveonChoose Then
If Not MultiLineEditBox Then If EditFlagSpecial And caretCreated Then caretCreated = False: DestroyCaret
SELECTEDITEM = -1: RaiseEvent Selected2(-2)
End If
If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar

RaiseEvent ShowExternalCursor
state = False
End Sub



Private Sub UserControl_Initialize()
SkipChars = 1
mSortstyle = vbTextCompare
mCurTab = 1
mTabs = 1
mLeftTab = 1
ReDim mParts(1 To 1)
mParts(1) = 0
tParam.cbSize = LenB(tParam)
tParam.iTabLength = 4
mTabStop = True
Buffer = 100
Set m_font = UserControl.Font
ReDim mList(0 To Buffer)
Dim i As Long
scrTwips = 1440# / GetDeviceCaps(UserControl.hDC, LOGPIXELSX)
dragslow = 1
DrawWidth = 1
DrawStyle = 0
NoPanLeft = True
NoPanRight = True
Clear
maxchar = 50
WordCharLeft = " ,."
WordCharRight = " ,."
BarColor = &H63DFFE  '&HC3C3C3
Shape1.hatchType = 1
End Sub
Property Let TabWidthChar(RHS As Long)
    tParam.iTabLength = Abs(RHS)
End Property
Private Sub UserControl_InitProperties()
 BackColor = m_def_BackColor
   ForeColor = m_def_ForeColor
    CapColor = m_def_CapColor
 dcolor = m_def_dcolor
mValue = 0
mmin = 0
mVertical = False
mjumptothemousemode = False
minimumWidth = 60
mLargeChange = 1
mSmallChange = 1
mmax = 100
mpercent = 0.07
NoPanLeft = True
NoPanRight = True

End Sub
Public Sub RefreshNow()
    If NoFreeMoveUpDown Then
    If ListIndex < topitem Then topitem = ListIndex Else topitem = topitem - 1
    If ListIndex - topitem > lines Then topitem = ListIndex - lines
    If topitem < 0 Then topitem = 0
       ShowMe2
    Else

If ListIndex < topitem Then topitem = ListIndex
      PrepareToShow 5
    End If
End Sub
Public Sub PressKey(KeyCode As Integer, shift As Integer, Optional NoEvents As Boolean = False)
Dim lcnt As Long, osel As Long, lsep As Long
If shift <> 0 And KeyCode = 16 Then Exit Sub
Timer1.enabled = False
If BlinkON Then BlinkTimer.enabled = True
'Timer1.Interval = 1000
Dim LastListIndex As Long, bb As Boolean, val As Long
LastListIndex = ListIndex
If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
    If Not Spinner Then If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar
    If (KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) And (shift And 2) = 2 Then
        If MarkNext = 0 Then
            RaiseEvent KeyDown(KeyCode, shift)
            If KeyCode <> 0 Then GetKeY2 KeyCode, shift
        End If
        If KeyCode = 0 Then Exit Sub
    End If
    If MarkNext = 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
End If

If KeyCode = 93 Then
' you have to clear myButton, here keycode
RaiseEvent OutPopUp(nowX, nowY, KeyCode)
End If
Select Case KeyCode
Case vbKeyHome
If EditFlagSpecial And (shift And 2) = 0 Then
    RaiseEvent DelExpandSS
    If mSelstart = 1 Then
    If mTabs = 1 Then
        mSelstart = Len(list(ListIndex)) - Len(NLtrim(list(ListIndex))) + 1
        Else
        mSelstart = Len(listAtColumn(ListIndex, mCurTab)) - Len(NLtrim(listAtColumn(ListIndex, mCurTab))) + 1
        End If
    Else
        mSelstart = 1
    End If
Else
       SELECTEDITEM = 1
       lcnt = listcount
       osel = SELECTEDITEM
       secreset = False
       lsep = ListSep(ListIndex)
       Do While Not (Not lsep Or ListIndex >= lcnt - 1)
            secreset = False
            SELECTEDITEM = osel + 1
            osel = SELECTEDITEM
            lsep = ListSep(ListIndex)
        Loop
        If ListSep(ListIndex) Then ListIndex = LastListIndex Else ShowThis SELECTEDITEM
        RaiseEvent ChangeSelStart(SelStart)
        If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
End If
Case vbKeyEnd
If EditFlagSpecial And (shift And 2) = 0 Then
    If mTabs = 1 Then
        mSelstart = Len(list(ListIndex)) + 1
    Else
        mSelstart = Len(listAtColumn(ListIndex, mCurTab)) + 1
    End If
    RaiseEvent SetExpandSS(mSelstart)
Else
    SELECTEDITEM = listcount
    osel = SELECTEDITEM
    secreset = False
    Do While Not (Not ListSep(ListIndex) Or ListIndex = 0)
        secreset = False
        SELECTEDITEM = osel - 1
        osel = SELECTEDITEM
        If osel < 0 Then Exit Do
    Loop
    If osel < 0 Then
    SELECTEDITEM = listcount - 1
    End If
    secreset = False
    If ListSep(ListIndex) Then ListIndex = LastListIndex Else ShowThis SELECTEDITEM
    RaiseEvent SetExpandSS(mSelstart)
    RaiseEvent ChangeSelStart(SelStart)
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
End If
Case vbKeyPageUp
    If shift = 0 Then RaiseEvent MarkDestroyAny
    If (shift And 2) = 2 Then
        ShowThis 1
        If EditFlagSpecial Then
            RaiseEvent DelExpandSS
            If mSelstart = 1 Then
                If mTabs = 1 Then
                    mSelstart = Len(list(ListIndex)) - Len(NLtrim(list(ListIndex))) + 1
                Else
                    mSelstart = Len(listAtColumn(ListIndex, mCurTab)) - Len(NLtrim(listAtColumn(ListIndex, mCurTab))) + 1
                End If
            Else
                mSelstart = 1
           End If
        End If
    ElseIf SELECTEDITEM - lines < 0 Then
        If SELECTEDITEM - 1 > 0 Then
            SELECTEDITEM = SELECTEDITEM - 1
        Else
            PrepareToShow 5
            If shift <> 0 Then If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
            shift = 0: KeyCode = 0: Exit Sub
        End If
    Else
        If topitem < SELECTEDITEM - (lines + 1) \ 2 Then
            If topitem = 0 Then
                SELECTEDITEM = SELECTEDITEM - 1
            Else
                SELECTEDITEM = topitem
            End If
        Else
            SELECTEDITEM = SELECTEDITEM - (lines + 1) \ 2 - 1
        End If
    End If
    osel = SELECTEDITEM - 1
    secreset = False
    If osel > 0 Then
        Do While ListSep(osel)
            secreset = False
            If osel = 0 Then Exit Do
            osel = osel - 1
        Loop
    End If
    osel = osel + 1
    secreset = False
    If ListSep(osel - 1) Then
        Do While ListSep(osel - 1)
            secreset = False
            If osel < 2 Then Exit Do
            osel = osel - 1
        Loop
        If ListSep(osel - 1) Then
            ListIndex = LastListIndex
            ShowMe2
        Else
            SELECTEDITEM = LastListIndex + 1
            ShowThis osel
        End If
    Else
        SELECTEDITEM = LastListIndex + 1
        ShowThis osel
    End If
    RaiseEvent ChangeSelStart(SelStart)
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
    If shift <> 0 Then If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
    shift = 0: KeyCode = 0: Exit Sub
Case vbKeyUp
    If Spinner Then Exit Sub
    
    If SELECTEDITEM > 1 Then
        osel = SELECTEDITEM
        Do
            osel = osel - 1
            secreset = False
            SELECTEDITEM = osel
        Loop Until Not (ListSep(ListIndex) Or ListJoin(ListIndex)) Or ListIndex = 0
        secreset = False
        If Not MultiLineEditBox Then
            If ListSep(ListIndex) Or ListJoin(ListIndex) Then ListIndex = LastListIndex Else ShowThis osel
        Else
            ShowThis osel
        End If
    Else
        If Not MultiLineEditBox Then If ListSep(ListIndex) Or ListJoin(ListIndex) Then ListIndex = LastListIndex
    End If
    RaiseEvent ExpandSelStart((mSelstart))
    RaiseEvent ChangeSelStart((mSelstart))
    If Not NoEvents Then If SELECTEDITEM > 0 Then If Not NoArrowUp Then RaiseEvent selected(SELECTEDITEM)
    If shift <> 0 Then
        If ListIndex < topitem Then topitem = ListIndex
        PrepareToShow 5
        If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
    Else
        RaiseEvent MarkDestroyAny
        MarkNext = 0
        If NoFreeMoveUpDown Then
            If ListIndex < topitem Then topitem = ListIndex Else topitem = topitem - 1
            If ListIndex - topitem > lines Then topitem = ListIndex - lines
            If topitem < 0 Then topitem = 0
            ShowMe2
        Else
            KeyCode = 0
            If ListIndex < topitem Then topitem = ListIndex
            PrepareToShow 5
        End If
        Exit Sub
    End If
    shift = 0: KeyCode = 0: Exit Sub
Case vbKeyDown
    If Spinner Then Exit Sub
    lcnt = listcount
    If SELECTEDITEM < lcnt Then
        osel = SELECTEDITEM
        Do
            osel = osel + 1
            secreset = False
            SELECTEDITEM = osel
        Loop Until Not (ListSep(ListIndex) Or ListJoin(ListIndex)) Or ListIndex = lcnt - 1
        secreset = False
        If Not MultiLineEditBox Then
            If ListSep(ListIndex) Or ListJoin(ListIndex) Then ListIndex = LastListIndex Else ShowThis osel
        Else
            ShowThis osel
        End If
    Else
        If Not MultiLineEditBox Then If ListSep(ListIndex) Or ListJoin(ListIndex) Then ListIndex = LastListIndex
    End If
    RaiseEvent ExpandSelStart((mSelstart))
    RaiseEvent ChangeSelStart((mSelstart))
    If Not NoEvents Then If SELECTEDITEM > 0 Then If Not NoArrowDown Then RaiseEvent selected(SELECTEDITEM)
    If shift <> 0 Then
        If ListIndex - topitem > lines Then topitem = ListIndex - lines
        PrepareToShow 5
        If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
    Else
        RaiseEvent MarkDestroyAny
        MarkNext = 0
        If NoFreeMoveUpDown Then
            If topitem + lines + 2 > listcount Then
                If ListIndex - topitem > lines Then topitem = ListIndex - lines
            Else
                topitem = topitem + 1
            End If
            ShowMe2
        Else
            KeyCode = 0
            If ListIndex - topitem > lines Then topitem = ListIndex - lines
            PrepareToShow 5
        End If
        Exit Sub
    End If
     KeyCode = 0: Exit Sub
Case vbKeyPageDown
    lcnt = listcount
    If shift = 0 Then RaiseEvent MarkDestroyAny
    If (shift And 2) = 2 Then
        ShowThis lcnt
        If EditFlagSpecial Then
            mSelstart = Len(list(ListIndex)) + 1
            RaiseEvent SetExpandSS(mSelstart)
        End If
    ElseIf SELECTEDITEM + (lines + 1) \ 2 >= lcnt Then
        If listcount > SELECTEDITEM Then
            SELECTEDITEM = SELECTEDITEM + 1
        Else
            PrepareToShow 5
            If shift <> 0 Then If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
            shift = 0: KeyCode = 0: Exit Sub
        End If
    ElseIf (SELECTEDITEM - topitem) <= (lines + 1) \ 2 Then
        If topitem + (lines + 1) + 1 <= lcnt Then
            SELECTEDITEM = topitem + (lines + 1) + 1
        Else
            SELECTEDITEM = SELECTEDITEM + 1
        End If
    Else
        SELECTEDITEM = SELECTEDITEM + (lines + 1) \ 2 + 1
    End If
    osel = SELECTEDITEM - 1
    secreset = False
    If osel < lcnt - 1 And osel > 0 Then
    Do While ListSep(osel)
        secreset = False
        If osel + 1 = lcnt Then Exit Do
        osel = osel + 1
    Loop
    End If
    secreset = False
    If ListSep(osel) Then
        ListIndex = LastListIndex
        ShowMe2
    Else
        SELECTEDITEM = LastListIndex + 1 ' osel
        ShowThis osel + 1
    End If
    RaiseEvent ChangeSelStart(SelStart)
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
    If shift <> 0 Then If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
    shift = 0: KeyCode = 0: Exit Sub
Case vbKeySpace
    If SELECTEDITEM > 0 Then
        If EditFlagSpecial Then
        If mSelstart = 0 Then mSelstart = 1
            If mTabs = 1 Then
                If maxchar = 0 Or (maxchar > Len(list(SELECTEDITEM - 1)) Or MultiLineEditBox) Then
                    bb = enabled
                    enabled = False
                    RaiseEvent PushUndoIfMarked
                    RaiseEvent MarkDelete(False)
                    enabled = bb
                    SelStartEventAlways = SelStart
                    RaiseEvent DelExpandSS
                    RaiseEvent PureListOn
                    If shift = 5 Then
                        list(SELECTEDITEM - 1) = Left$(list(SELECTEDITEM - 1), SelStart - 1) & ChrW(&H2007) & Mid$(list(SELECTEDITEM - 1), SelStart)
                        RaiseEvent RemoveOne(ChrW(&H2007))
                    ElseIf shift = 3 Then
                        list(SELECTEDITEM - 1) = Left$(list(SELECTEDITEM - 1), SelStart - 1) & ChrW(&HA0) & Mid$(list(SELECTEDITEM - 1), SelStart)
                        RaiseEvent RemoveOne(ChrW(&HA0))
                    Else
                        list(SELECTEDITEM - 1) = Left$(list(SELECTEDITEM - 1), SelStart - 1) & " " & Mid$(list(SELECTEDITEM - 1), SelStart)
                        RaiseEvent RemoveOne(" ")
                    End If
                    RaiseEvent PureListOff
                    SelStartEventAlways = SelStart + 1
                    RaiseEvent SetExpandSS(mSelstart)
                    KeyCode = 0
                    If listcount > 0 Or MultiLineEditBox Then
                        If OverrideShow And Not HandleOverride Then
                            ShowMe
                        Else
                          ShowMe2
                        End If
                        HandleOverride = False
                    Else
                        ShowMe
                    End If
                End If
            Else
                If maxchar = 0 Or (maxchar > Len(list(SELECTEDITEM - 1)) Or MultiLineEditBox) Then
                    bb = enabled
                    enabled = False
                    RaiseEvent PushUndoIfMarked
                    RaiseEvent MarkDelete(False)
                    enabled = bb
                    SelStartEventAlways = SelStart
                    RaiseEvent DelExpandSS
                    RaiseEvent PureListOn
                    If shift = 5 Then
                        listAtColumn(SELECTEDITEM - 1, mCurTab) = Left$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart - 1) & ChrW(&H2007) & Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart)
                        RaiseEvent RemoveOne(ChrW(&H2007))
                    ElseIf shift = 3 Then
                        listAtColumn(SELECTEDITEM - 1, mCurTab) = Left$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart - 1) & ChrW(&HA0) & Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart)
                        RaiseEvent RemoveOne(ChrW(&HA0))
                    Else
                        listAtColumn(SELECTEDITEM - 1, mCurTab) = Left$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart - 1) & " " & Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart)
                        RaiseEvent RemoveOne(" ")
                    End If
                    RaiseEvent PureListOff
                    SelStartEventAlways = SelStart + 1
                    RaiseEvent SetExpandSS(mSelstart)
                    KeyCode = 0
                    'PrepareToShow 2
                    If listcount > 0 Or MultiLineEditBox Then
                        If OverrideShow And Not HandleOverride Then
                            ShowMe
                        Else
                          ShowMe2
                        End If
                        HandleOverride = False
                    Else
                        ShowMe
                    End If
                End If
            End If
            Exit Sub
    Else
        If (MultiSelect Or ListMenu(SELECTEDITEM - 1)) Then
            If ListRadio(SELECTEDITEM - 1) And ListSelected(SELECTEDITEM - 1) Then
                ' do nothing
            Else
                ListSelected(SELECTEDITEM - 1) = Not ListSelected(SELECTEDITEM - 1)
                ' from 1 to listcount
                If MultiSelect Then
                    If ListSelected(SELECTEDITEM - 1) Then
                        RaiseEvent SelectedMultiAdd(SELECTEDITEM)
                    Else
                        RaiseEvent SelectedMultiSub(SELECTEDITEM)
                    End If
                Else
                    RaiseEvent MenuChecked(SELECTEDITEM)
                End If
            End If
        End If
    End If
End If
Case vbKeyLeft
RaiseEvent DelExpandSS
If EditFlagSpecial Then
    val = 1
    RaiseEvent SubSelStart(val, shift)
    If MultiLineEditBox Then
        If SelStart > val Then
            mSelstart = SelStart - val
            If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
            RaiseEvent MayRefresh(bb)
            If bb Then ShowMe2
        ElseIf ListIndex > 0 Then
            ShowThis SELECTEDITEM - 1
            If mTabs = 1 Then
                mSelstart = Len(list(ListIndex)) + 1
            Else
                mSelstart = Len(listAtColumn(ListIndex, mCurTab)) + 1
            End If
            If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
            If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
        End If
    ElseIf SelStart > val Then
            If mTabs = 1 Then
                osel = Len(list(ListIndex))
            Else
                osel = Len(listAtColumn(ListIndex, mCurTab))
            End If
            If (SelStart - val) < osel Then mSelstart = SelStart - val Else mSelstart = osel
            If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
    End If
Else
    If Not NoEvents Then
    If SELECTEDITEM > 0 Then
            If mTabs > 1 Then
                If mCurTab > 1 Then
                    lcnt = mCurTab - 1
                    Do While mParts(lcnt) < 4
                        lcnt = lcnt - 1
                       If lcnt = 0 Then Exit Do
                    Loop
                    If lcnt <> 0 Then
                        mCurTab = lcnt
                        TwipsCurTab = TwipsCurTab - mParts(mCurTab) * scrTwips
                    End If
                End If
                RaiseEvent selected(SELECTEDITEM)
                RaiseEvent MayRefresh(bb)
                If bb Then ShowMe2
                
            End If
        If Not Arrows2Tab Then
            RaiseEvent selected(SELECTEDITEM)
        End If
        End If
    End If
End If
Case vbKeyRight
If EditFlagSpecial Then
    val = 1
    RaiseEvent AddSelStart(val, shift)
    If mTabs = 1 Then
        If MultiLineEditBox Then
            If SelStart <= Len(list(SELECTEDITEM - 1)) - val + 1 Then
                mSelstart = SelStart + val
                If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
                RaiseEvent MayRefresh(bb)
                If bb Then ShowMe2
            ElseIf ListIndex < listcount - 1 Then
                ListindexPrivateUse = ListIndex + 1
                mSelstart = 1
                If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
                If (SELECTEDITEM - topitem) > lines + 1 Then topitem = topitem + 1
                If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
            End If
        Else
            If SelStart <= Len(list(SELECTEDITEM - 1)) - val + 1 Then
                mSelstart = SelStart + val
                If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
            End If
        End If
    Else
        If MultiLineEditBox Then
            If SelStart <= Len(listAtColumn(SELECTEDITEM - 1, mCurTab)) - val + 1 Then
                mSelstart = SelStart + val
                If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
                RaiseEvent MayRefresh(bb)
                If bb Then ShowMe2
            ElseIf ListIndex < listcount - 1 Then
                ListindexPrivateUse = ListIndex + 1
                mSelstart = 1
                If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
                If (SELECTEDITEM - topitem) > lines + 1 Then topitem = topitem + 1
                If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
            End If
        Else
            If SelStart <= Len(listAtColumn(SELECTEDITEM - 1, mCurTab)) - val + 1 Then
                mSelstart = SelStart + val
                If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
            End If
        End If
    End If
Else
    If Not NoEvents Then
        If SELECTEDITEM > 0 Then
            If mTabs > 1 Then
                If mCurTab < mTabs Then
                    lcnt = mCurTab + 1
                    Do While mParts(lcnt) < 4
                        lcnt = lcnt + 1
                       If lcnt > mTabs Then Exit Do
                    Loop
                    If lcnt <= mTabs Then
                        TwipsCurTab = TwipsCurTab + mParts(mCurTab) * scrTwips
                        mCurTab = lcnt
                    End If
                End If
                RaiseEvent selected(SELECTEDITEM)
                RaiseEvent MayRefresh(bb)
                If bb Then ShowMe2
            End If
            If Not Arrows2Tab Then RaiseEvent selected(SELECTEDITEM)
        End If
    End If
End If
Case vbKeyDelete
If EditFlagSpecial Then
If mSelstart = 0 Then mSelstart = 1
If mTabs = 1 Then
    If SelStart > Len(list(SELECTEDITEM - 1)) Then
        mSelstart = Len(list(SELECTEDITEM - 1)) + 1
        If listcount > SELECTEDITEM Then
            If Not NoEvents Then
                RaiseEvent LineDown
                RaiseEvent addone(vbCr)
            End If
        End If
    Else
        RaiseEvent PureListOn
        val = 1
        RaiseEvent AddSelStart(val, shift)
        RaiseEvent addone(Mid$(list(SELECTEDITEM - 1), SelStart, val))
        list(SELECTEDITEM - 1) = Left$(list(SELECTEDITEM - 1), SelStart - 1) + Mid$(list(SELECTEDITEM - 1), SelStart + val)
        RaiseEvent SetExpandSS(mSelstart)
        RaiseEvent PureListOff
        ShowMe2
    End If
Else
    If SelStart > Len(listAtColumn(SELECTEDITEM - 1, mCurTab)) Then
        mSelstart = Len(listAtColumn(SELECTEDITEM - 1, mCurTab)) + 1
        If listcount > SELECTEDITEM Then
            If Not NoEvents Then
                RaiseEvent LineDown
                RaiseEvent addone(vbCr)
            End If
        End If
    Else
        RaiseEvent PureListOn
        val = 1
        RaiseEvent AddSelStart(val, shift)
        RaiseEvent addone(Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart, val))
        listAtColumn(SELECTEDITEM - 1, mCurTab) = Left$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart - 1) + Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart + val)
        RaiseEvent SetExpandSS(mSelstart)
        RaiseEvent PureListOff
        ShowMe2
    End If
End If
End If
Case vbKeyBack
If EditFlagSpecial Then
If mTabs = 1 Then
    If SelStart > 1 Then
        val = 1
        RaiseEvent PureListOn
        RaiseEvent SubSelStart(val, shift)
        osel = Len(list(ListIndex))
        If (SelStart - val) < osel Then SelStart = SelStart - val Else SelStart = osel
        
        'SelStart = SelStart - val  ' make it a delete because we want selstart to take place before list() take value     RaiseEvent PureListOn
        RaiseEvent addone(Mid$(list(SELECTEDITEM - 1), SelStart, val))
        list(SELECTEDITEM - 1) = Left$(list(SELECTEDITEM - 1), SelStart - 1) + Mid$(list(SELECTEDITEM - 1), SelStart + val)
        RaiseEvent SetExpandSS(mSelstart)
        RaiseEvent PureListOff
        ShowMe2  'refresh now
    Else
        If mSelstart = 0 Then mSelstart = 1
        RaiseEvent SetExpandSS(mSelstart)
        If Not NoEvents Then RaiseEvent LineUp
    End If
Else
    If SelStart > 1 Then
        val = 1
        RaiseEvent PureListOn
        RaiseEvent SubSelStart(val, shift)
        osel = Len(listAtColumn(ListIndex, mCurTab))
        
        If (SelStart - val) < osel Then SelStart = SelStart - val Else SelStart = osel
        'SelStart = SelStart - val  ' make it a delete because we want selstart to take place before list() take value     RaiseEvent PureListOn
        RaiseEvent addone(Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart, val))
        listAtColumn(SELECTEDITEM - 1, mCurTab) = Left$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart - 1) + Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), SelStart + val)
        RaiseEvent SetExpandSS(mSelstart)
        RaiseEvent PureListOff
        ShowMe2  'refresh now
    Else
        If mSelstart = 0 Then mSelstart = 1
        RaiseEvent SetExpandSS(mSelstart)
        If Not NoEvents Then RaiseEvent LineUp
    End If
End If
End If
Case vbKeyReturn
If MultiLineEditBox Then
    RaiseEvent SplitLine
    RaiseEvent SetExpandSS(mSelstart)
    RaiseEvent RemoveOne(vbCrLf)
Else
    If mTabs > 1 Then
    
                If (shift And 7) = 1 Then
                    If mCurTab > 1 Then
                        lcnt = mCurTab - 1
                        Do While mParts(lcnt) < 4
                            lcnt = lcnt - 1
                           If lcnt = 0 Then Exit Do
                        Loop
                        If lcnt <> 0 Then
                            mCurTab = lcnt
                            TwipsCurTab = TwipsCurTab - mParts(mCurTab) * scrTwips
                        End If
                    Else
                        'mCurTab = mTabs
                        lcnt = mTabs
                        Do While mParts(lcnt) < 4
                            lcnt = lcnt - 1
                           If lcnt = 0 Then Exit Do
                        Loop
                        If lcnt <> 0 Then
                            mCurTab = mTabs
                            For lcnt = 1 To mCurTab - 1
                                TwipsCurTab = TwipsCurTab + mParts(lcnt) * scrTwips
                            Next lcnt
                        End If
                    End If
                Else
                    If mCurTab < mTabs Then
                        lcnt = mCurTab + 1
                        Do While mParts(lcnt) < 4
                            lcnt = lcnt + 1
                            If lcnt > mTabs Then Exit Do
                        Loop
                        If lcnt <= mTabs Then
                            TwipsCurTab = TwipsCurTab + mParts(mCurTab) * scrTwips
                            mCurTab = lcnt
                        End If
                    Else
                        TwipsCurTab = 0
                        mCurTab = 1
                        lcnt = 1
                        Do While mParts(lcnt) < 4
                            lcnt = lcnt + 1
                           If lcnt > mTabs Then Exit Do
                        Loop
                        If lcnt <= mTabs Then
                            mCurTab = lcnt
                        
                        End If
                    End If
                End If
                
                mSelstart = 1
                
               RaiseEvent selected(SELECTEDITEM)
            End If
    RaiseEvent EnterOnly
End If
End Select
If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
End If

If MultiLineEditBox Then
    SelStartEventAlways = SelStart
    If shift Or Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Me.PrepareToShow 5
Else
    KeyCode = 0
    SelStartEventAlways = SelStart
    Me.PrepareToShow 5
End If
KeyCode = 0
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, shift As Integer)
On Error GoTo fin
If PrevLocale <> GetLocale Then RaiseEvent Maybelanguage
If BypassKey Then KeyCode = 0: shift = 0: Exit Sub
Dim i As Long, k As Integer

lastshift = shift

If KeyCode = 18 Then
'RaiseEvent Maybelanguage
ElseIf KeyCode = 112 And (shift And 2) = 2 Then
KeyCode = 0
shift = 0
RaiseEvent CtrlPlusF1
Exit Sub
ElseIf KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
    k = ((KeyCode - vbKeyF1 + 1) + 12 * (shift And 1)) + 24 * (1 + ((shift And 2) = 0)) - 1000 * ((shift And 4) = 4)
    RaiseEvent Fkey(k)
    If k = 0 Then KeyCode = 0: shift = 0
ElseIf KeyCode = 16 And shift <> 0 Then
  '  RaiseEvent Maybelanguage
ElseIf KeyCode = vbKeyV Then
Exit Sub
Else
If KeyCode = 27 And NoEscapeKey Then
KeyCode = 0
Else
RaiseEvent RefreshOnly
End If
End If
i = -1
If shift <> 4 And mynum$ <> "" Then
On Error Resume Next
If Left$(mynum$, 1) = "0" Then
i = val(mynum$)
Else
i = val(mynum$)
End If
mynum$ = vbNullString
If i > 32 Then
If i >= &H10000 And i <= &H10FFFF Then
i = i - &H10000
UKEY$ = ChrW(UINT(i \ &H400& + &HD800&)) + ChrW(UINT((i And &H3FF&) + &HDC00&))
Else
UKEY$ = ChrW(i)
End If
If LastNumX Then UserControl_KeyPress 44
RefreshNow
Exit Sub
End If
i = -1
Else
i = GetLastKeyPressed
End If

 If i <> -1 And i <> 94 Then
 If i = 13 Then
 UKEY$ = vbNullString
 Else
 UKEY$ = ChrW(i)
 End If
 Else
UKEY$ = vbNullString
 End If
fin:
End Sub

Private Sub UserControl_LostFocus()
useFloatList = False
doubleclick = 0
Fkey = 0
If Not NoWheel Then RaiseEvent UnregisterGlist
RaiseEvent CheckLostFocus
If myEnabled Then SoftExitFocus
havefocus = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, shift As Integer, X As Single, y As Single)
' cut area
Dim YYT As Long, oldbutton As Integer
On Error GoTo there
If FloatList And UseHeaderOnly Then
        missMouseClick = False
        FloatListMe useFloatList, X, y
        Exit Sub
End If
If DropKey Then Exit Sub

If missMouseClick Then Exit Sub
If MousePointer = vbSizeWE And (Button And 3) > 0 And Not useFloatList And Not UseHeaderOnly Then
    If X < UserControl.Width - scrTwips Then
        If (shift And 3) = 1 Or (Button And 2) = 2 Then
            YYT = mParts(AdjustColumn)
            mParts(AdjustColumn) = CLng(X - scrTwips) \ scrTwips - AdjustColumnSum
            If mTabs > AdjustColumn Then
                If mParts(AdjustColumn + 1) + (YYT - mParts(AdjustColumn)) > 0 Then
                    mParts(AdjustColumn + 1) = mParts(AdjustColumn + 1) + (YYT - mParts(AdjustColumn))
                Else
                    mParts(AdjustColumn + 1) = 1
                End If
            End If
            YYT = 0
        Else
        
        mParts(AdjustColumn) = (CLng(X - scrTwips) \ scrTwips - AdjustColumnSum)
        End If
        TwipsCurTab = 0
        Dim k As Integer
        For k = 1 To mCurTab - 1
        TwipsCurTab = TwipsCurTab + mParts(k) * scrTwips
        Next k
        If mParts(AdjustColumn) < 1 Then mParts(AdjustColumn) = 1: MousePointer = 1
    Else
     MousePointer = 1
    End If
    PrepareToShow 2
Exit Sub
End If
nowX = X
nowY = y
If (Button And 2) = 2 And (Not (MousePointer = vbSizeWE And Not useFloatList And Not UseHeaderOnly)) Then Exit Sub
If myt = 0 Then Exit Sub
FreeMouse = True

If mHeadlineHeightTwips = 0 Then
YYT = y \ myt
Else
    If y < mHeadlineHeightTwips Then
        If y < 0 Then
        YYT = -1
        Else
        YYT = 0
        End If
    Else
    YYT = (y - mHeadlineHeightTwips) \ myt + 1
    End If
End If
If YYT < 0 Then YYT = 0
If (YYT >= 0 And (YYT < listcount Or listcount = 0) And myEnabled) Then

oldbutton = Button

If mHeadline <> "" And Timer2.enabled = False Then
    If YYT = 0 Then ' we move in mHeadline
        ' -1 is mHeadline
        ' headline listen clicks if  list is disabled...
        RaiseEvent ExposeItemMouseMove(Button, -1, CLng(X / scrTwips), CLng(y / scrTwips))
        If (X < Width - mPreserveNpixelsHeaderRight) Or (mPreserveNpixelsHeaderRight = 0) Then RaiseEvent HeaderSelected(Button)
        If oldbutton <> Button Then
        Button = 0
        Exit Sub
        End If
    ElseIf myEnabled Then
       If Not Shape1.Visible Then
        RaiseEvent ExposeItemMouseMove(Button, topitem - TopRows + YYT - 1, CLng(X / scrTwips), CLng(y - (YYT - 1) * myt - mHeadlineHeightTwips) / scrTwips)
        End If
    End If
ElseIf myEnabled Then
    RaiseEvent ExposeItemMouseMove(Button, topitem - TopRows + YYT, CLng(X / scrTwips), CLng((y - YYT * myt) / scrTwips))
End If
If oldbutton <> Button Then Exit Sub
End If
YYT = YYT + (mHeadline <> "")
lastX = X
LastY = y

If (X > Width - BarWidth) And BarVisible And EnabledBar And Button = 1 Then
If Vertical Then
GetOpenValue = valuepoint - y + mHeadlineHeightTwips
Else
'GetOpenValue = valuepoint - x ' NOT USED HERE
End If
If processXY(lastX, LastY, False) And myEnabled Then
FreeMouse = False
End If
Timer3.enabled = False
Else
cx = X
If Not dr Then lY = X
dr = True

If cy = y \ myt Then Timer3.enabled = False: cy = y \ myt
End If
If MarkNext = 4 Then
    RaiseEvent MarkDestroyAny
End If
there:
End Sub

Private Sub UserControl_MouseMove(Button As Integer, shift As Integer, X As Single, y As Single)
On Error GoTo there
If FloatList And UseHeaderOnly And Button <> 0 Then
        missMouseClick = False
        FloatListMe useFloatList, X, y
        Exit Sub
End If
If DropKey Then Exit Sub
Dim osel As Long, tListcount As Long, YYT As Long, oldbutton As Integer, k As Integer, sum As Long
If missMouseClick Then Exit Sub
If Abs(PX - X) <= 60 And Abs(PY - y) <= 60 Then Exit Sub
PX = X
PY = y

RaiseEvent MouseMove(Button, shift, X, y)
If myt = 0 Or Not myEnabled Then Exit Sub
If (Button And 2) = 2 And (Not (MousePointer = vbSizeWE And Not useFloatList)) Then Exit Sub

tListcount = listcount

If timestamp = 0 Or (timestamp - Timer) > 1 Then timestamp = Timer
If (timestamp + 0.02) > Timer And shift = 0 Then Exit Sub
timestamp = Timer
If Not FreeMouse Then Exit Sub
If Button = 0 Then
    If Not nopointerchange Then
        If mTabs > 1 And AdjustColumns And Not useFloatList Then
            osel = CLng(X)
            For k = 1 To mTabs
                sum = sum + mParts(k)
            Next k
            For k = mTabs To 1 Step -1
                If mParts(k) > 0 Then
                    If Abs(osel - sum * scrTwips - scrTwips) <= scrTwips * 2 Then
                        MousePointer = vbSizeWE
                        AdjustColumnSum = sum - mParts(k)
                        AdjustColumn = k
                        GoTo ok_i_found
                    End If
                sum = sum - mParts(k)
                End If
            Next k
            If MousePointer < 2 Or (MousePointer = vbSizeWE And Not useFloatList) Then MousePointer = 1
ok_i_found:
            osel = 0
        ElseIf Not UseHeaderOnly Then
            If MousePointer < 2 Then MousePointer = 1
        End If
    End If
End If
If MousePointer = vbSizeWE And (Button And 3) > 0 And Not useFloatList And Not UseHeaderOnly Then
    If X < UserControl.Width - scrTwips Then
            If (shift And 3) = 1 Or (Button And 2) = 2 Then
            YYT = mParts(AdjustColumn)
            mParts(AdjustColumn) = CLng(X - scrTwips) \ scrTwips - AdjustColumnSum
            If mTabs > AdjustColumn Then
                If mParts(AdjustColumn + 1) + (YYT - mParts(AdjustColumn)) > 0 Then
                    mParts(AdjustColumn + 1) = mParts(AdjustColumn + 1) + (YYT - mParts(AdjustColumn))
                Else
                    mParts(AdjustColumn + 1) = 1
                End If
            End If
            YYT = 0
            Else
        mParts(AdjustColumn) = (CLng(X - scrTwips) \ scrTwips - AdjustColumnSum)
        End If
        TwipsCurTab = 0
        For k = 1 To mCurTab - 1
        TwipsCurTab = TwipsCurTab + mParts(k) * scrTwips
        Next k
        If mParts(AdjustColumn) < 1 Then mParts(AdjustColumn) = 1: MousePointer = 1
    Else
     MousePointer = 1
    End If
    PrepareToShow 2
Exit Sub
End If

If (X > Width - BarWidth) And tListcount > lines + 1 And Not BarVisible Then
    Hidebar = True: BarVisible = m_showbar Or AutoHide Or MultiLineEditBox
ElseIf (X < Width - BarWidth) And Button = 0 And BarVisible And (StickBar Or AutoHide) Then
    Hidebar = False
    BarVisible = False
End If
If OurDraw Then
    barMouseMove Button, shift, X, y
    Exit Sub
End If
cx = X
Timer3.enabled = False
If mHeadlineHeightTwips = 0 Then
    YYT = y \ myt
Else
    If y < mHeadlineHeightTwips Then
        If y < 0 Then
        YYT = -1
        Else
        YYT = 0
        End If
    Else
    YYT = (y - mHeadlineHeightTwips) \ myt + 1
    End If
End If
oldbutton = Button
If (Button And 3) > 0 And useFloatList And FloatList Then
        FloatListMe useFloatList, X, y: Button = 0
    Else
        If Not nopointerchange Then
            If MousePointer > 1 And Not (MousePointer = vbSizeWE And Not useFloatList) And Not UseHeaderOnly Then MousePointer = 1
        End If
End If
    If mHeadline <> "" Then
        If YYT = 0 Then ' we move in mHeadline
            ' -1 is mHeadline
            If (Button And 3) > 0 And FloatList And Not useFloatList Then FloatListMe useFloatList, X, y: Button = 0
            RaiseEvent ExposeItemMouseMove(Button, -1, CLng(X / scrTwips), CLng(y / scrTwips))
        ElseIf Not (MousePointer = vbSizeWE And Not useFloatList) Then
            If YYT <= TopRows Then
                RaiseEvent ExposeItemMouseMove(Button, YYT - 1, CLng(X / scrTwips), CLng(y - (YYT - 1) * myt) / scrTwips)
            Else
                RaiseEvent ExposeItemMouseMove(Button, topitem - TopRows + YYT - 1, CLng(X / scrTwips), CLng(y - (YYT - 1) * myt) / scrTwips)
            End If
        End If
    ElseIf Not (MousePointer = vbSizeWE And Not useFloatList) Then
        If YYT < TopRows Then
            RaiseEvent ExposeItemMouseMove(Button, YYT, CLng(X / scrTwips), CLng(y - YYT * myt) / scrTwips)
        Else
            RaiseEvent ExposeItemMouseMove(Button, topitem - TopRows + YYT, CLng(X / scrTwips), CLng(y - YYT * myt) / scrTwips)
        End If
    End If
    If oldbutton <> Button Then Exit Sub
    YYT = YYT + (mHeadline <> "")
    If (Button And 3) = 0 Then
        If YYT >= 0 And YYT <= lines Then
            If topitem + YYT < tListcount Then
                secreset = False
            End If
        End If
    ElseIf dr Then
        If MultiLineEditBox And (Button = 1) And secreset Then
            If MarkNext > 3 Then
            ElseIf MarkNext = 0 Then
                MarkNext = 1
                RaiseEvent markin
            End If
        End If
        If (SELECTEDITEM <> (topitem + YYT + 1)) And SELECTEDITEM >= 0 And Button <> 0 Then secreset = False
        ' special for M2000  (StickBar And x > Width / 2)
        If shift = 0 And ((Not scrollme > 0) And (X > Width / 2 Or (mTabs > 1 And (X - TwipsCurTab) > mParts(mCurTab) * scrTwips \ 2)) Or Not SingleLineSlide) And StickBar And MarkNext = 0 And tListcount > lines + 1 Then
            If Abs(LastY - y) < scrTwips * 2 Then LastY = y: Exit Sub
            Hidebar = True
            CalcAndShowBar1
            If LastY < y Then
                y = scrTwips * 2
            Else
                y = ScaleHeight - scrTwips
            End If
            If Abs(lastX - X) < scrTwips * 4 Or Not MultiLineEditBox Then
                lastX = X
                LastY = y
                If Vertical Then
                    GetOpenValue = valuepoint - y + mHeadlineHeightTwips
                Else
              '  GetOpenValue = valuepoint - x ' NO USED HERE
                End If
                If processXY(lastX, LastY, True) Then
                    FreeMouse = False
                End If
                Timer3.enabled = False
                Exit Sub
            Else
                If YYT >= 0 And YYT <= lines Then shift = 1: GoTo there1
            End If
        End If
        If mHeadline <> "" And y < mHeadlineHeightTwips Then
        ' we sent twips not pixels
        ' move...me..??
        
        ElseIf (y - mHeadlineHeightTwips) < myt / 2 And (topitem + YYT > 0) Then
        'scroll up
            drc = True
            Timer2.enabled = True
        ElseIf y > ScaleHeight - myt \ 2 And (tListcount <> 1) Then
            drc = False
            Timer2.enabled = True
        ElseIf YYT >= 0 And YYT <= lines + TopRows Then
there1:
            If MultiLineEditBox And (Button = 1) Then
                If MarkNext = 1 Then
                    shift = 1
                    RaiseEvent MarkOut
                ElseIf shift = 0 And MarkNext = 2 Then
                    MarkNext = 0  ' so markNext=2 we have a complete marked text
                    RaiseEvent MarkDestroy
                End If
            End If
            If Timer2.enabled Then
                Timer2.enabled = False
            End If
            If topitem + YYT - TopRows < tListcount Then
                If (cx > ScaleWidth / 4 And cx < ScaleWidth * 3 / 4) And scrollme = 0 Then X = lY
                If Not SELECTEDITEM = topitem - TopRows + YYT + 1 And Not (YYT < TopRows And SELECTEDITEM = YYT + 1) Then
                    osel = SELECTEDITEM
                    SELECTEDITEM = topitem + YYT + 1
                    If YYT < TopRows Then
                        SELECTEDITEM = YYT + 1
                    End If
            If ListJoin(SELECTEDITEM - 1) Then
             Do While SELECTEDITEM > 1 And ListJoin(SELECTEDITEM - 1)
             SELECTEDITEM = SELECTEDITEM - 1
             Loop
             End If
             If SELECTEDITEM - 1 < topitem Then
              topitem = SELECTEDITEM - 1
             End If
                    If Not BlockItemcount Then
                        REALCUR SELECTEDITEM - 1, cx - scrollme, dummy, mSelstart
                        mSelstart = mSelstart + 1
                        
                        RaiseEvent ChangeSelStart(mSelstart)
                    End If
                    If MultiLineEditBox And (Button = 1) Then
                        If shift = 1 And MarkNext = 0 Then
                            MarkNext = 1
                            RaiseEvent markin
                        ElseIf shift = 1 And MarkNext = 1 Then
                            RaiseEvent MarkOut
                        End If
                    End If
                    If StickBar Or AutoHide Then DOT3
                        If X - lY > 0 And Not NoPanRight Then
                            scrollme = (X - lY)
                        ElseIf X - lY < 0 And Not NoPanLeft Then
                            scrollme = (X - lY)
                        Else
                            If Not EditFlagSpecial Then scrollme = 0
                        End If
                        If Not EditFlagSpecial Then If scrollme > 0 Then scrollme = 0
                ElseIf cy <> YYT Then
                    cy = YYT
                    Timer3.enabled = True
                Else
                    If Not Timer1.enabled Then
                        If Not BlockItemcount Then
                            REALCUR SELECTEDITEM - 1, cx - scrollme, dummy, mSelstart
                            mSelstart = mSelstart + 1
                            
                            RaiseEvent ChangeSelStart(mSelstart)
                        End If
                        If MultiLineEditBox And (Button = 1) Then
                        If shift = 1 And MarkNext = 0 Then
                            MarkNext = 1
                            RaiseEvent markin
                        ElseIf shift = 1 And MarkNext = 1 Then
                            RaiseEvent MarkOut
                        End If
                    End If
                    If X - lY > 0 And Not NoPanRight Then
                        scrollme = (X - lY)
                    ElseIf X - lY < 0 And Not NoPanLeft Then
                        scrollme = (X - lY)
                    Else
                        If Not EditFlagSpecial Then scrollme = 0
                    End If
                    
                    If MarkNext = 4 Then
                        RaiseEvent MarkDestroyAny
                    End If
                    Timer1.Interval = 20
                    Timer1.enabled = True
                End If
                Timer3.enabled = False
            End If
        End If
    End If
End If
there:
End Sub

Public Sub CheckMark()
' if shift =0
    If MarkNext >= 1 Then
    If MarkNext < 4 Then
                MarkNext = 0  ' so markNext=2 we have a complete marked text
                RaiseEvent MarkDestroy
                ShowMe2
                Else
                MarkNext = MarkNext - 1
                End If
      End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, shift As Integer, X As Single, y As Single)
On Error GoTo there
If DropKey Then Exit Sub
If missMouseClick Then missMouseClick = False: Exit Sub
If MousePointer = vbSizeWE And Not useFloatList And Not UseHeaderOnly Then
    MousePointer = 1
    Exit Sub
End If
If Button = 1 Or UseHeaderOnly Then RaiseEvent MouseUp(X / scrTwips, y / scrTwips)
If (Button And 2) = 2 Then
X = nowX
y = nowY
End If
If useFloatList Then
    Timer1.enabled = False
    If MoveParent Then
        If TypeOf Extender.Container Is GuiM2000 Then
            RaiseEvent RefreshDesktop
        Else
            RaiseEvent PrepareContainer
        End If
    End If
    useFloatList = False
End If
If myt = 0 Then Exit Sub
Timer1bar.Interval = 100
Timer1bar.enabled = False
If OurDraw Then
OurDraw = False
Exit Sub
End If
Timer2.enabled = False
If Not (FreeMouse Or Not myEnabled) Then Exit Sub

With UserControl
 If (X < 0 Or y < 0 Or X > .Width Or y > .Height) And (LeaveonChoose And Not BypassLeaveonChoose) Then
If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar
 SELECTEDITEM = -1
 RaiseEvent Selected2(-2)
 Exit Sub
 End If
End With
cx = X
If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar
If Timer3.enabled Then
cy = y: DOT3
End If
Timer3.enabled = False
If Timer2.enabled Then
 Timer2.enabled = False
 End If
Dim YYT As Long
  If dr Then
                    lY = 0
            
                    If scrollme < -myt Then
                        RaiseEvent PanLeftRight(False)
                    ElseIf scrollme > myt Then
                        RaiseEvent PanLeftRight(True)
                    Else
                    dr = False
                    GoTo jump1
                    End If
                 If Not EditFlagSpecial Then scrollme = 0
                    Timer1.enabled = True
                    dr = False
                End If
jump1:
If mHeadlineHeightTwips = 0 Then
YYT = y \ myt
Else
    If y < mHeadlineHeightTwips Then
        If y < 0 Then
        YYT = -1
        Else
        YYT = 0
        End If
    Else
    YYT = (y - mHeadlineHeightTwips) \ myt + 1
    End If
End If


If YYT = -1 Then Button = 0
If mHeadline <> "" And YYT = 0 Then Button = 0
YYT = YYT + (mHeadline <> "")

If YYT >= 0 And YYT <= lines + TopRows Then

If topitem + YYT - TopRows < listcount Then

If (Button And 3) > 0 And myEnabled Then

    
    If secreset Then
        ' this is a double click
        secreset = False
         If Not ListSep(topitem - TopRows + YYT) Then
            If MarkNext = 0 And (EditFlagSpecial Or MultiLineEditBox) Then
                If MultiLineEditBox And Not EditFlagSpecial Then
                    REALCUR SELECTEDITEM - 1, cx - scrollme, dummy, mSelstart
                    mSelstart = mSelstart + 1
                End If
                MarkWord
            Else
                
        Timer1.enabled = False
        If (((SELECTEDITEM <> (topitem - TopRows + YYT + 1)) And Not secreset) Or EditFlagSpecial Or (mTabs > 1 And NarrowSelect)) And Not ListSep(topitem - TopRows + YYT) Then
             SELECTEDITEM = topitem - TopRows + YYT + 1 ' we have a new selected item
             If YYT < TopRows Then
                SELECTEDITEM = YYT + 1
             End If
             If ListJoin(SELECTEDITEM - 1) Then
             Do While SELECTEDITEM > 1 And ListJoin(SELECTEDITEM - 1)
             SELECTEDITEM = SELECTEDITEM - 1
             Loop
             End If
             If SELECTEDITEM - 1 < topitem Then
              topitem = SELECTEDITEM - 1
             End If
             ' compute selstart always
            If Not BlockItemcount Then
                REALCUR SELECTEDITEM - 1, cx - scrollme, dummy, mSelstart
                mSelstart = mSelstart + 1
                RaiseEvent SetExpandSS(mSelstart)
                RaiseEvent ChangeSelStart(mSelstart)
                
            End If
            RaiseEvent selected(SELECTEDITEM)  ' broadcast
         End If
            If SELECTEDITEM = topitem - TopRows + YYT + 1 Then
             If YYT < TopRows Then
                SELECTEDITEM = YYT + 1
             End If
            If ListJoin(SELECTEDITEM - 1) Then
             Do While SELECTEDITEM > 1 And ListJoin(SELECTEDITEM - 1)
             SELECTEDITEM = SELECTEDITEM - 1
             Loop
             End If
             If SELECTEDITEM - 1 < topitem Then
              topitem = SELECTEDITEM - 1
             End If
                If MultiSelect Or ListMenu(SELECTEDITEM - 1) Then
                        If ListRadio(SELECTEDITEM - 1) And ListSelected(SELECTEDITEM - 1) Then
                        ' do nothing
                        ElseIf ListRadio(SELECTEDITEM - 1) Then
                            ListSelected(SELECTEDITEM - 1) = Not ListSelected(SELECTEDITEM - 1)
                            If MultiSelect Then
                                If ListSelected(SELECTEDITEM - 1) Then
                                    RaiseEvent SelectedMultiAdd(SELECTEDITEM)
                                Else
                                    RaiseEvent SelectedMultiSub(SELECTEDITEM)
                                End If
                            Else
                                RaiseEvent MenuChecked(SELECTEDITEM)
                            End If
                        End If
                        End If
                    End If
                RaiseEvent Selected2(SELECTEDITEM - 1)
                
            End If
        End If
        
    Else

        Timer1.enabled = False
        If (((SELECTEDITEM <> (topitem - TopRows + YYT + 1)) And Not secreset) Or EditFlagSpecial Or (mTabs > 1)) And Not ListSep(topitem - TopRows + YYT) Then
             SELECTEDITEM = topitem - TopRows + YYT + 1 ' we have a new selected item
             If YYT < TopRows Then
                SELECTEDITEM = YYT + 1
             End If
             If ListJoin(SELECTEDITEM - 1) Then
             Do While SELECTEDITEM > 1 And ListJoin(SELECTEDITEM - 1)
             SELECTEDITEM = SELECTEDITEM - 1
             Loop
             End If
             If SELECTEDITEM - 1 < topitem Then
              topitem = SELECTEDITEM - 1
             End If
             
             ' compute selstart always
            If Not BlockItemcount Then
                REALCUR SELECTEDITEM - 1, cx - scrollme, dummy, mSelstart
                mSelstart = mSelstart + 1
                RaiseEvent SetExpandSS(mSelstart)
                RaiseEvent ChangeSelStart(mSelstart)
                
            End If
            RaiseEvent selected(SELECTEDITEM)  ' broadcast
         End If
            If SELECTEDITEM = topitem + YYT + 1 Then
             If YYT < TopRows Then
                SELECTEDITEM = YYT + 1
             End If
            If ListJoin(SELECTEDITEM - 1) Then
             Do While SELECTEDITEM > 1 And ListJoin(SELECTEDITEM - 1)
             SELECTEDITEM = SELECTEDITEM - 1
             Loop
             End If
             If SELECTEDITEM - 1 < topitem Then
              topitem = SELECTEDITEM - 1
             End If
                If MultiSelect Or ListMenu(SELECTEDITEM - 1) Then
                    If (X / scrTwips > 0) And (X / scrTwips < LeftMarginPixels) Then
                        If ListRadio(SELECTEDITEM - 1) And ListSelected(SELECTEDITEM - 1) Then
                        ' do nothing
                        Else
                            ListSelected(SELECTEDITEM - 1) = Not ListSelected(SELECTEDITEM - 1)
                            If MultiSelect Then
                                If ListSelected(SELECTEDITEM - 1) Then
                                    RaiseEvent SelectedMultiAdd(SELECTEDITEM)
                                Else
                                    RaiseEvent SelectedMultiSub(SELECTEDITEM)
                                End If
                            Else
                                RaiseEvent MenuChecked(SELECTEDITEM)
                            End If
                        End If
                            If Not enabled Then Exit Sub
                            PrepareToShow 5
                            Exit Sub
                        End If
                    End If
End If

End If
If secreset = False Then If shift = 0 Then CheckMark
If Not enabled Then Exit Sub
secreset = True
PrepareToShow 5
 If Button = 2 Then
RaiseEvent OutPopUp(X, y, Button)

End If
''
End If
'End If

End If
End If
there:
End Sub



Private Sub UserControl_OLECompleteDrag(Effect As Long)
dragfocus = False
If Effect = 0 Then
' CANCEL...
If marvel Then
RaiseEvent MarkDestroy
ShowMe2

End If
ElseIf Effect = vbDropEffectMove Then
If marvel Then
RaiseEvent PushUndoIfMarked
If Not NoMoveDrag Then RaiseEvent MarkDelete(False)
End If
End If
Effect = 0
RaiseEvent MarkDestroyAny
HideCaretOnexit = False
Timer2.enabled = False
If marvel Then RaiseEvent CorrectCursorAfterDrag
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, shift As Integer, X As Single, y As Single)
Dim something$, ok As Boolean
If DropKey Then Exit Sub
 
If (Effect And 3) > 0 Then
If Data.GetFormat(vbCFText) Or Data.GetFormat((13)) Then

If (Button And 1) = 0 Then
    If (shift And 2) = 2 Then
        Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectMove
            End If
        End If
End If
End If
RaiseEvent DropOk(ok)
If marvel Then

Else
RaiseEvent MarkDestroyAny
ok = True
End If
If ok Then
        If Data.GetFormat(13) Then
          
          something$ = Data.GetData(13)
          Else
        
            something$ = Data.GetData(vbCFText)
            End If
something$ = Replace(something$, ChrW(0), "")

If marvel Then
RaiseEvent DropFront(ok)
If ok Then
RaiseEvent selected(SELECTEDITEM)

    RaiseEvent DragPasteData(something$)
 
   If Effect = vbDropEffectMove Then
 RaiseEvent addone(something$)
 
   RaiseEvent MarkDelete(True)
    RaiseEvent RemoveOne("")
    Else

        RaiseEvent MarkDestroyAny
    End If
Else
If Effect = vbDropEffectMove Then
    RaiseEvent addone(something$)
    RaiseEvent PushMark2Undo(something$)
    RaiseEvent MarkDelete(True)
    
Else
    RaiseEvent MarkDestroyAny
End If
    RaiseEvent selected(SELECTEDITEM)
    RaiseEvent DragPasteData(something$)
    
End If
Else
RaiseEvent selected(SELECTEDITEM)
RaiseEvent DragPasteData(something$)

End If
marvel = False



Else
Effect = 0
End If

End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, shift As Integer, X As Single, y As Single, state As Integer)
On Error Resume Next
If DropKey Then Exit Sub
If Not DropEnabled Then Effect = 0: Exit Sub
Dim tListcount As Long, YYT As Long, oldpb As Boolean, oldbp As Boolean
If Not TaskMaster Is Nothing Then
        If TaskMaster.QueueCount > 0 And Not STbyST Then
            oldbp = bypasstrace
            bypasstrace = True
            TaskMaster.RestEnd1
            TaskMaster.TimerTick
            TaskMaster.RestEnd1
            bypasstrace = oldbp
        End If
    End If
tListcount = listcount
 If state = vbOver Then
 
If mHeadline <> "" And y < mHeadlineHeightTwips Then
' we sent twips not pixels
' move...me..??

ElseIf (y - mHeadlineHeightTwips) < myt / 2 And (topitem + YYT > 0) Then
                drc = True
                Timer2.enabled = True
        
        ElseIf y > ScaleHeight - myt \ 2 And (tListcount <> 1) Then
                drc = False
                Timer2.enabled = True
        Else
                Timer2.enabled = False
             '  If marvel Then
             
                If Not Timer1.enabled Then
                If Not havefocus Then RaiseEvent DragOverCursor(dragfocus)
                HideCaretOnexit = False: MovePos X, y
                If CBool(shift And 1) Then chooseshow
                
                End If

                              
                              
                               
              '  End If
            If Data.GetFormat(vbCFText) Or Data.GetFormat((13)) Then
                        If (shift And 2) = 2 Then
                            Effect = vbDropEffectCopy
                        Else

                            Effect = vbDropEffectMove
                        End If
                Else
                    Effect = vbDropEffectNone
            End If
            End If
ElseIf state = vbLeave Then
Dim ok As Boolean
dragfocus = False
missMouseClick = True
If Not marvel And Effect = 0 Then RaiseEvent DragOverDone(ok)
If Not ok Then
        Timer2.enabled = False
        
        Timer3.enabled = True
        Effect = vbDropEffectNone
        HideCaretOnexit = True
        If caretCreated Then caretCreated = False: DestroyCaret
        MovePos X, y
        If CBool(shift And 1) Then chooseshow
    End If
ElseIf state = vbEnter Then
ok = False
If Not marvel Then RaiseEvent DragOverNow(ok)
If Not ok Then
    If Not Timer1.enabled Then
        HideCaretOnexit = False
        MovePos X, y
        If CBool(shift And 1) Then chooseshow
    Else
        dragfocus = False
    End If
End If

                               
        End If
        
             If Data.GetFormat(vbCFText) Or Data.GetFormat((13)) Then
                    If (shift And 2) = 2 Then
                       Effect = vbDropEffectCopy
                       Else

                           Effect = vbDropEffectMove
                           End If
            Else
                Effect = vbDropEffectNone
        End If
      
End Sub



Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Dim oldbp As Boolean
On Error Resume Next
   If Not TaskMaster Is Nothing Then
        If TaskMaster.QueueCount > 0 And Not STbyST Then
            oldbp = bypasstrace
            bypasstrace = True
            TaskMaster.RestEnd1
            TaskMaster.TimerTick
            TaskMaster.RestEnd1
            bypasstrace = oldbp
        End If
    End If
   
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
If DropKey Then Exit Sub
If Not DragEnabled Then Exit Sub
Dim aa() As Byte, this$
RaiseEvent DragData(this$)
aa = this$ & ChrW$(0)
 Data.SetData aa(), 13
Data.SetData aa(), vbCFText
AllowedEffects = vbDropEffectCopy + vbDropEffectMove
End Sub
Public Sub MovePos(ByVal X As Single, ByVal y As Single)

Dim dummy As Long, YYT As Long, M_CURSOR As Long
dragslow = 0.02
If mHeadlineHeightTwips = 0 Then
YYT = y \ myt + 1
Else
    If y < mHeadlineHeightTwips Then
        If y < 0 Then
        Exit Sub
        Else
        YYT = 1
        End If
    Else
    YYT = (y - mHeadlineHeightTwips) \ myt + 1
    End If
End If
YYT = YYT - 1
If topitem + YYT < listcount Then
REALCUR topitem + YYT, X - scrollme, dummy, M_CURSOR
ListindexPrivateUse = topitem + YYT
If ListIndex = -1 Then
        If itemcount = 0 Then
        additemFast ""
        End If
        ListindexPrivateUse = 0

End If
SelStart = M_CURSOR + 1

Else
ListindexPrivateUse = listcount - 1
            If ListIndex = -1 Then
            If itemcount = 0 Then
            additemFast ""
            End If
            ListindexPrivateUse = 0
            
            End If
SelStart = Len(list(ListIndex)) + 1

End If
RaiseEvent selected(SELECTEDITEM)

RaiseEvent ChangeSelStart(SelStart)
dragslow = 1
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
m_sync = .ReadProperty("sync", m_def_sync)
NoFire = True
Value = .ReadProperty("Value", 0)
Max = .ReadProperty("Max", 100)
Min = .ReadProperty("Min", 0)
largechange = .ReadProperty("LargeChange", 1)
smallchange = .ReadProperty("SmallChange", 1)
Percent = .ReadProperty("Percent", 0.07)
Vertical = .ReadProperty("Vertical", False)
jumptothemousemode = .ReadProperty("JumptoTheMouseMode", False)
NoFire = False
Set Font = .ReadProperty("Font", Ambient.Font)


myEnabled = .ReadProperty("Enabled", m_def_Enabled)


BackStyle = .ReadProperty("BackStyle", m_def_BackStyle)
BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)

m_showbar = .ReadProperty("ShowBar", m_def_Showbar)
dcolor = .ReadProperty("dcolor", m_def_dcolor)
BackColor = .ReadProperty("BackColor", m_def_BackColor)
ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
CapColor = .ReadProperty("CapColor", m_def_CapColor)

   Text = .ReadProperty("Text", m_def_Text)

   End With
   If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
Else

myt = UserControlTextHeight() + addpixels * scrTwips
End If
HeadlineHeight = UserControlTextHeight() / scrTwips
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
waitforparent = True
End Sub
Public Sub Dynamic()
overrideTextHeight = 0
   If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) / restrictLines
Else

myt = UserControlTextHeight() + addpixels * scrTwips
End If
HeadlineHeight = UserControlTextHeight() / scrTwips
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
waitforparent = True
End Sub

Private Sub UserControl_Show()
If Not design() Then
'CalcAndShowBar
fast = True
SoftEnterFocus

End If
End Sub

Private Sub UserControl_Terminate()
If LastGlist Is Me Then Set LastGlist = Nothing
waitforparent = True
Set m_font = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

With PropBag
     .WriteProperty "sync", m_sync, m_def_sync
    .WriteProperty "Value", Value, 0
    .WriteProperty "Max", Max, 100
    .WriteProperty "Min", Min, 0
    .WriteProperty "LargeChange", largechange, 1
    .WriteProperty "SmallChange", smallchange, 1
    .WriteProperty "Percent", Percent, 0.07
    .WriteProperty "Vertical", Vertical, False
    .WriteProperty "JumptoTheMouseMode", jumptothemousemode, False
    .WriteProperty "Font", m_font, Ambient.Font
    .WriteProperty "Enabled", myEnabled, m_def_Enabled
    .WriteProperty "BackStyle", m_BackStyle, m_def_BackStyle
   .WriteProperty "BorderStyle", m_BorderStyle, m_def_BorderStyle
    .WriteProperty "ShowBar", m_showbar, m_def_Showbar
    .WriteProperty "dcolor", dcolor, m_def_dcolor
     .WriteProperty "Backcolor", BackColor, m_def_BackColor
       .WriteProperty "ForeColor", ForeColor, m_def_ForeColor
    .WriteProperty "CapColor", CapColor, m_def_CapColor

      .WriteProperty "Text", Text, ""
      End With

End Sub
Property Get ListIndex() As Long
If SELECTEDITEM < 0 Then
ListIndex = -1  ' CHANGED
Else
ListIndex = SELECTEDITEM - 1
End If
End Property
Property Let ListIndex(ByVal item As Long)
On Error GoTo there
If listcount <= lines + 1 Then
BarVisible = False
Else
On Error Resume Next
Redraw m_showbar And Extender.Visible
Err.Clear
On Error GoTo there
End If
Do While ListJoin(item) And item > 0
    item = item - 1
Loop
If item < listcount Then SELECTEDITEM = item + 1
SkipReadEditflag = False
If SELECTEDITEM > 0 Then

RaiseEvent softSelected(SELECTEDITEM)
Else
SELECTEDITEM = 0
End If
there:
End Property
Public Sub FloatListMe(state As Boolean, X As Single, y As Single)
Static preX As Single, preY As Single
Dim dummy As Long, vx As Single, vy As Single
On Error GoTo there
If Extender.Parent Is Nothing Then Exit Sub
If Not state Then
preX = X
preY = y
state = True
MousePointer = 0
doubleclick = 0
GetMonitorsAgain
Else
If NoVerMove Then preY = y
If NoHorMove Then preX = X
If Extender.Visible Then
If NoVerMove Then
    MousePointer = 9
ElseIf NoHorMove Then
    MousePointer = 7
Else
    MousePointer = 5
End If
RaiseEvent NeedDoEvents
If MoveParent Then
    If Not TypeOf Extender.Container Is VB.PictureBox Then
        If (Extender.Container.top + (y - preY) < MinMonitorTop) Then preY = y + Extender.Container.top - MinMonitorTop
        If (Extender.Container.Left + (X - preX) < MinMonitorLeft) Then preX = X + Extender.Container.Left - MinMonitorLeft
        If ((Extender.Container.top + y - preY) > FloatLimitTop) And FloatLimitTop > ((Extender.Parent.Left + X - preX) > FloatLimitLeft) And FloatLimitLeft > MinMonitorTop Then preY = Extender.Parent.top + y - FloatLimitTop
        If ((Extender.Container.Left + X - preX) > FloatLimitLeft) And FloatLimitLeft > MinMonitorLeft Then preX = Extender.Parent.Left + X - FloatLimitLeft
        Else
            If ((Extender.Container.top + y - preY) > FloatLimitTop) And FloatLimitTop > ((Extender.Container.Left + X - preX) > FloatLimitLeft) And FloatLimitLeft > 0 Then preY = Extender.Parent.top + y - FloatLimitTop
            If ((Extender.Container.Left + X - preX) > FloatLimitLeft) And FloatLimitLeft > 0 Then preX = Extender.Container.Left + X - FloatLimitLeft
            If (Extender.Container.top + (y - preY) < 0) Then preY = y + Extender.Container.top
            If (Extender.Container.Left + (X - preX) < 0) Then preX = X + Extender.Container.Left
    End If
        
        
        
    
        vx = Extender.Container.Left + (X - preX)
        vy = Extender.Container.top + (y - preY)
        RaiseEvent ValidMove(vx, vy)
        Extender.Container.move vx, vy
        Timer1.enabled = False
        Timer1.Interval = 2
        Timer1.enabled = True
        'If TypeOf Extender.Container Is GuiM2000 Then
        '    RaiseEvent RefreshDesktop
        'Else
        '    RaiseEvent PrepareContainer
        'End If
Else
Extender.ZOrder
If (Extender.top + (y - preY) < 0) Then preY = y + Extender.top
If (Extender.Left + (X - preX) < 0) Then preX = X + Extender.Left
If ((Extender.top + y - preY) > FloatLimitTop) And FloatLimitTop > 0 Then preY = Extender.top + y - FloatLimitTop

If ((Extender.Left + X - preX) > FloatLimitLeft) And FloatLimitLeft > 0 Then preX = Extender.Left + X - FloatLimitLeft
    vx = Extender.Left + (X - preX)
    vy = Extender.top + (y - preY)
    RaiseEvent ValidMove(vx, vy)
    
    Extender.move vx, vy
End If
End If
If Me.BackStyle = 1 Then ShowMe2
End If
there:
End Sub
Property Let listAtColumn(ByVal item As Long, ByVal thiscolumn As Long, ByVal b$)
On Error GoTo nnn1
If itemcount = 0 Or BlockItemcount Then
    If listAtColumn(item, thiscolumn) <> b$ Then
        RaiseEvent ChangeColumnListItem(item, thiscolumn, b$)
    End If
Exit Property
End If
If item >= 0 And item < itemcount Then
With mList(item)
If .content Is Nothing Then Set .content = New JsonObject
If Not .content.ItemPath("C." & thiscolumn) = b$ Then RaiseEvent ChangeColumnListItem(item, thiscolumn, b$)
.content.AssignPath "C." & thiscolumn, CVar(b$)
.Flags = .Flags Or (fline + fselected + joinpRevline)
.Flags = .Flags Xor (fline + fselected + joinpRevline)
End With
End If
nnn1:
End Property
Public Sub JoinLine(ByVal item As Long)
' JoinPrevLine
On Error GoTo nnn1
Dim i As Long
If itemcount = 0 Or BlockItemcount Then
    Exit Sub
End If
If item >= 1 And item < itemcount Then
With mList(item)
If Not .content Is Nothing Then Set .content = Nothing
    If (.Flags And joinpRevline) = 0 Then
        JoinLines = JoinLines + 1
        .Flags = ((.Flags Or 15) Xor 15) Or joinpRevline
        i = item - 1
        If .morerows > 0 Then
            mList(i).morerows = .morerows + 1
            .morerows = 0
        Else
            Do While i > 0 And (mList(i).Flags And joinpRevline) <> 0
                i = i - 1
            Loop
            mList(i).morerows = mList(i).morerows + 1
        End If
    End If
End With
End If
nnn1:
End Sub
Public Sub SplitLine(ByVal item As Long)
' JoinPrevLine
On Error GoTo nnn1
Dim i As Long
If itemcount = 0 Or BlockItemcount Then
    Exit Sub
End If
If item >= 1 And item < itemcount Then
With mList(item)
If (.Flags And joinpRevline) <> 0 Then
JoinLines = JoinLines - 1
.Flags = .Flags Or joinpRevline
.Flags = .Flags Xor joinpRevline
    i = item - 1
    Do While i > 0 And (mList(i).Flags And joinpRevline) <> 0
        i = i - 1
    Loop
    mList(i).morerows = item - i + 1
    i = item + 1
    Do While i < itemcount And (mList(i).Flags And joinpRevline) <> 0
        i = i + 1
    Loop
    .morerows = i - item - 1
End If
End With
End If
nnn1:
End Sub
Public Sub PropertyLet(ByVal Prop As String, ByVal RHS, Optional item, Optional thiscolumn)
Dim that$, rep As Variant
Prop$ = Replace(Prop$, ".", "")
If Len(Prop$) = 0 Then Exit Sub
Prop$ = UCase(Prop$)
If itemcount = 0 Or BlockItemcount Then Exit Sub
If IsMissing(item) Then
      Select Case Prop$
            Case "EDIT"
                 mEditFlag = CBool(RHS)
            Case "CENTERTEXT"
                CenterText = CBool(RHS)
            Case "RIGHTTEXT"
                 RightText = CBool(RHS)
            Case "WRAPTEXT"
                 WrapText = CBool(RHS)
            Case "VERTICALCENTERTEXT"
                 VerticalCenterText = CBool(RHS)
            Case "LASTLINEPART"
                LastLinePart = CStr(RHS)
            End Select
Else
    If item < 0 Then Exit Sub
    If item >= listcount Then Exit Sub
    If IsMissing(thiscolumn) Then
        With mList(item)
            If .content Is Nothing Then Set .content = New JsonObject
            .content.AssignPath "P.0.[" & Prop$ + "]", RHS
        End With
    Else
        If thiscolumn < 1 Or thiscolumn > mTabs Then thiscolumn = 1
        With mList(item)
            If .content Is Nothing Then Set .content = New JsonObject
            .content.AssignPath "P." & thiscolumn & ".[" & Prop$ + "]", RHS
        End With
    End If
End If
End Sub
Property Get PropAtColumnNum(ByVal item As Long, ByVal thiscolumn As Long, ByVal Prop As String) As Variant
Dim that As Variant
Prop$ = Replace(Prop$, ".", "")
If Len(Prop$) = 0 Then Exit Property
Prop$ = UCase(Prop$)
If thiscolumn < 1 Or thiscolumn > mTabs Then thiscolumn = 1
If itemcount = 0 Or BlockItemcount Then
    RaiseEvent ReadColumnPropNum(item, thiscolumn, Prop$, that)
    If VarType(that) = vbEmpty Then
        GoTo alfa
    Else
        PropAtColumnNum = that
    End If
    
    Exit Property
End If
If item < 0 Then Exit Property
If item >= listcount Then
'Err.Raise vbObjectError + 1050
Else
With mList(item)
If .content Is Nothing Then
    PropAtColumnNum = 0
Else
    that = .content.ItemPath("P." & thiscolumn & ".[" & Prop$ & "]")
    If VarType(that) = vbEmpty Then
        that = .content.ItemPath("P.0.[" & Prop$ & "]")
alfa:
        If VarType(that) = 0 Then
            Select Case Prop$
            Case "EDIT"
                that = mEditFlag
            Case "CENTERTEXT"
                that = CenterText
            Case "RIGHTTEXT"
                that = RightText
            Case "WRAPTEXT"
                that = WrapText
            Case "VERTICALCENTERTEXT"
                that = VerticalCenterText
            Case "LASTLINEPART"
                that = LastLinePart
            End Select
        End If
    End If
    PropAtColumnNum = that
End If
End With
End If
End Property
Property Get PropAtColumn(ByVal item As Long, ByVal thiscolumn As Long, ByVal Prop As String) As String
Dim that$, rep As Variant
Prop$ = Replace(Prop$, ".", "")
If Len(Prop$) = 0 Then Exit Property
Prop$ = UCase(Prop$)
If thiscolumn < 1 Or thiscolumn > mTabs Then thiscolumn = 1
If itemcount = 0 Or BlockItemcount Then
    RaiseEvent ReadColumnProp(item, thiscolumn, Prop$, that$)
    If that$ = "" Then
        GoTo alfa
    Else
        PropAtColumn = that$
    End If
    Exit Property
End If
If item < 0 Then Exit Property
If item >= listcount Then
'Err.Raise vbObjectError + 1050
Else
With mList(item)
If .content Is Nothing Then
    PropAtColumn = vbNullString
Else
    rep = .content.ItemPath("P." & thiscolumn & ".[" & Prop$ & "]")
    If VarType(rep) = vbEmpty Then
        rep = .content.ItemPath("P.0.[" & Prop$ & "]")
alfa:
        If VarType(rep) = 0 Then
            Select Case Prop$
            Case "EDIT"
                rep = mEditFlag
            Case "CENTERTEXT"
                rep = CenterText
            Case "RIGHTTEXT"
                rep = RightText
            Case "WRAPTEXT"
                rep = WrapText
            Case "VERTICALCENTERTEXT"
                rep = VerticalCenterText
            Case "LASTLINEPART"
                rep = LastLinePart
            End Select
        End If
    End If
    If VarType(rep) = vbBoolean Then
        If rep Then PropAtColumn = "1" Else PropAtColumn = "0"
    ElseIf Not VarType(rep) = vbEmpty Then
        If Not VarType(rep) = vbObject Then
        PropAtColumn = CStr(rep)
        End If
    End If
End If
End With
End If
End Property

Property Get listAtColumn(ByVal item As Long, ByVal thiscolumn As Long) As String
Dim that$
If thiscolumn < 1 Or thiscolumn > mTabs Then thiscolumn = 1
If itemcount = 0 Or BlockItemcount Then
    RaiseEvent ReadColumnItem(item, thiscolumn, that$)
    listAtColumn = that$
    Exit Property
End If
If item < 0 Then Exit Property
If item >= listcount Then
'Err.Raise vbObjectError + 1050
Else
With mList(item)
If .content Is Nothing Then
    listAtColumn = vbNullString
Else
    listAtColumn = .content.ItemPath("C." & thiscolumn)
End If
End With
End If
End Property
Property Get list(item As Long) As String
Dim that$
If itemcount = 0 Or BlockItemcount Then
    RaiseEvent ReadListItem(item, that$)
    list = that$
    Exit Property
End If
If item < 0 Then Exit Property
If item >= listcount Then
'Err.Raise vbObjectError + 1050
Else
With mList(item)
If .content Is Nothing Then
    list = vbNullString
Else
    list = .content.ItemPath("C.1")
End If
End With
End If
End Property

Property Let list(item As Long, ByVal b$)
On Error GoTo nnn1
If itemcount = 0 Or BlockItemcount Then
    If list(item) <> b$ Then
        RaiseEvent ChangeListItem(item, b$)
    End If
Exit Property
End If
If mTabs = 1 Then If mParts(1) = 0 Then mParts(1) = Me.WidthPixels

If item >= 0 Then
With mList(item)
If .content Is Nothing Then Set .content = New JsonObject
If Not .content.ItemPath("C.1") = b$ Then RaiseEvent ChangeListItem(item, b$)
.content.AssignPath "C.1", CVar(b$)
.Flags = .Flags Or (fline + fselected + joinpRevline)
.Flags = .Flags Xor (fline + fselected + joinpRevline)
End With
End If
nnn1:
End Property
Property Let menuEnabled(item As Long, ByVal RHS As Boolean)
If item >= 0 Then
With mList(item)
.Flags = .Flags Or fline
If RHS Then
.Flags = .Flags Xor fline
End If
''.Line = Not RHS   ' The line flag used as enabled flag, in reverse logic
End With
End If
End Property
Property Let ListSep(item As Long, ByVal RHS As Boolean)
If itemcount = 0 Or BlockItemcount Then Exit Property
If item >= 0 And item < itemcount Then
With mList(item)
'.Line = RHS
.Flags = .Flags Or fline
If Not RHS Then
.Flags = .Flags Xor fline
End If
End With
End If
End Property
Property Get ListSep(item As Long) As Boolean
Dim skip As Boolean, blockit As Boolean
RaiseEvent BlockCaret(item, blockit, skip)
If skip Then
ListSep = blockit
Exit Property
End If
If itemcount = 0 Or BlockItemcount Then Exit Property
If item >= 0 And item < itemcount Then
With mList(item)
ListSep = .Flags And fline
End With
End If
End Property
Property Get ListJoin(item As Long) As Boolean
If itemcount = 0 Or BlockItemcount Then Exit Property
If item >= 0 And item < itemcount Then
With mList(item)
ListJoin = .Flags And joinpRevline
End With
End If
End Property
Property Get JoinLinesCount() As Long
JoinLinesCount = JoinLines
End Property
Property Get ListMoreLines(item As Long) As Integer
If itemcount > 0 And Not BlockItemcount Then
    If item >= 0 And item < itemcount Then
        ListMoreLines = mList(item).morerows
    End If
End If
End Property
Property Let ListSelected(item As Long, ByVal b As Boolean)
Dim first As Long, last As Long
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 And item < itemcount Then

If mList(item).Flags And fradiobutton Then
        ' erase all
        first = item
        While first > 0 And (mList(first).Flags And fradiobutton)
        first = first - 1
        Wend
        If Not (mList(first).Flags And fradiobutton) Then first = first + 1
        last = item
        While last < listcount - 1 And (mList(last).Flags And fradiobutton)
        last = last + 1
        Wend
        If Not (mList(last).Flags And fradiobutton) Then last = last - 1
        For first = first To last
        With mList(first)
        'mList(first).selected = False
        .Flags = .Flags Or fselected
        .Flags = .Flags Xor fselected
        End With
        Next first
End If
With mList(item)
        .Flags = .Flags Or fselected
        If Not b Then .Flags = .Flags Xor fselected
        '.selected = b
End With
End If
End If
End Property
Property Let ListSelectedNoRadioCare(item As Long, ByVal b As Boolean)
Dim first As Long, last As Long
If itemcount > 0 And Not BlockItemcount Then
    If item >= 0 And item < itemcount Then
        With mList(item)
            .Flags = .Flags Or fselected
            If Not b Then
            .Flags = .Flags Xor fselected
            End If
        End With
    End If
End If
End Property
Property Get ListSelected(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
    If item >= 0 And item < itemcount Then
        With mList(item)
            ListSelected = .Flags And fselected
        End With
    End If
End If
End Property
Property Let ListRadio(item As Long, ByVal b As Boolean)
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 And item < itemcount Then
    With mList(item)
        .Flags = .Flags Or fradiobutton
        If Not b Then
            .Flags = .Flags Xor fradiobutton
        End If
    End With
End If
End If
End Property
Property Get ListRadio(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 And item < itemcount Then
    With mList(item)
        ListRadio = .Flags And fradiobutton
    End With
End If
End If
End Property
Property Get ListMenu(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
    If item >= 0 And item < itemcount Then
        With mList(item)
        'ListMenu = .radiobutton Or .checked
            ListMenu = .Flags And (fradiobutton + fchecked)
        End With
    End If
End If
End Property
Property Let ListChecked(item As Long, ByVal b As Boolean)
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mList(item)
'.checked = b
    .Flags = .Flags Or fchecked
    If Not b Then
        .Flags = .Flags Xor fchecked
    End If
End With
End If
End If
End Property
Property Get ListChecked(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
    If item >= 0 Then
        With mList(item)
            ListChecked = .Flags And fchecked
        End With
    End If
End If
End Property
Public Sub MoveTo(ByVal Key As String)
Dim i As Long
Key = Replace(Key, "]", "?")
Key = Replace(Key, "[", "[[]")
Key = Replace(Key, "?", "[]]")
RaiseEvent PureListOn
On Error GoTo 123
For i = 0 To listcount - 1
If list(i) Like Key Then Exit For
Next i
123
RaiseEvent PureListOff
If i < listcount Then
ListIndex = i
End If
End Sub
Public Function FindItemStartWidth(ByVal Key As String, NoCase As Boolean, ByVal Offset) As Long
Dim i As Long, j As Long
j = Len(Key)
i = -1
FindItemStartWidth = -1
If j = 0 Then Exit Function
If NoCase Then
For i = Offset To listcount - 1
If StrComp(Left$(list(i), j), Key, vbTextCompare) = 0 Then Exit For
Next i
Else
For i = Offset To listcount - 1
If StrComp(Left$(list(i), j), Key, vbBinaryCompare) = 0 Then Exit For
Next i
End If
If i < listcount Then
FindItemStartWidth = i
End If
End Function
Public Function Find(ByVal Key As String) As Long
Dim i As Long, skipme As Boolean
i = -1
Key = Replace(Key, "]", "?")
Key = Replace(Key, "[", "[[]")
Key = Replace(Key, "?", "[]]")
RaiseEvent Find(Key, i, skipme)
If skipme Then Find = i: Exit Function
Find = -1
RaiseEvent PureListOn
On Error GoTo 123
If mTabs > 0 Then
For i = 0 To listcount - 1
    If listAtColumn(i, mCurTab) Like Key Then Exit For
Next i

Else
For i = 0 To listcount - 1
If list(i) Like Key Then Exit For
Next i
End If
123
RaiseEvent PureListOff
If i < listcount Then
Find = i
End If
End Function
Public Sub ShowThis(ByVal item As Long, Optional noselect As Boolean)
On Error GoTo skipthis
If item <= 0 Then item = 1
If Extender.Parent Is Nothing Then Exit Sub
SkipReadEditflag = False
If listcount <= lines + 1 Then
    BarVisible = False
Else
    BarVisible = m_showbar And Extender.Visible

End If
If item > 0 And item <= listcount Then
    If MultiLineEditBox Then FindRealCursor item
    If item - topitem > 0 And item - topitem <= lines + 1 Then
            If item > 1 Then
                While ListJoin(item - 1) And item > 1
                item = item - 1
                Wend
            End If
        If restrictLines > 0 Then
            If listcount <= topitem + lines Then
                topitem = listcount - lines - 1
                If topitem < 0 Then topitem = 0
            End If
        End If
        SELECTEDITEM = item
        
        If SELECTEDITEM = listcount Then
            state = True
            Value = Max
            state = False
        End If
    Else
        If MultiLineEditBox And False Then
            If item < lines / 2 Then
                topitem = 0
            Else
                If item + lines / 2 > listcount Then
                    topitem = listcount - lines - 1
                Else
                    topitem = item - lines / 2
                End If
            End If
        Else
            If item - topitem > lines Then
                topitem = item - lines + 1
            Else
                topitem = item - 1
            End If
        End If
        If item > 1 Then
            While ListJoin(item - 1) And item > 1
            item = item - 1
            Wend
        End If
        CalcAndShowBar1
        SELECTEDITEM = item
        chooseshow
    End If
   If Not noselect Then If Not Timer1.enabled Then PrepareToShow 10
Exit Sub

End If
If MultiLineEditBox Then Exit Sub
If noselect Then
SELECTEDITEM = 0
  End If
ShowMe2
skipthis:
End Sub
Public Sub RepaintScrollBar()
If m_showbar Or StickBar Or AutoHide Or Shape1.Visible Or Spinner Then Redraw
If Not BarVisible Then Refresh
End Sub
Public Sub Clear(Optional ByVal interface As Boolean = False)
SELECTEDITEM = -1
LastSelected = -2
itemcount = 0
JoinLines = 0
If hWnd <> 0 Then HideCaret (hWnd)
state = True
mValue = 0  ' if here you have an error you forget to apply VALUE as default property
showshapes
LastVScroll = 1
'max = 0
state = False
topitem = 0
Buffer = 100
ReDim mList(0 To Buffer)
Dim i As Long
If interface Then
 '   barvisible = False
    ShowMe
End If
End Sub
Public Sub ClearClick()
SELECTEDITEM = -1
secreset = False
End Sub
Public Sub PrepareClick()
'bypassfirstClick = False
secreset = True
End Sub
Public Function DblClick() As Boolean
DblClick = secreset
secreset = False
End Function


Private Sub UserControl_Resize()
'If Not design() Then
RaiseEvent OnResize
CalcAndShowBar

'End If
End Sub
Public Sub additem(a$)
Dim i As Long

If itemcount = Buffer Then
Buffer = Buffer * 2
ReDim Preserve mList(0 To Buffer)
End If
itemcount = itemcount + 1
With mList(itemcount - 1)
Set .content = New JsonObject
If Trim(a$) <> "" Then .content.AssignPath "C.1", CVar(a$)
.Flags = .Flags Or (fline + fselected + joinpRevline)
.Flags = .Flags Xor (fline + fselected + joinpRevline)
End With
Timer1.enabled = False
Timer1.Interval = 100
Timer1.enabled = True
End Sub
Public Sub additemAtListIndex(a$)
Dim i As Long
If itemcount = Buffer Then
Buffer = Buffer * 2
ReDim Preserve mList(0 To Buffer)
End If
itemcount = itemcount + 1
For i = itemcount - 1 To ListIndex + 1 Step -1
mList(i) = mList(i - 1)
Next i
With mList(i)
Set .content = New JsonObject
On Error Resume Next
Set .content = .content.Parser(mList(i + 1).content.Json())
If .content.ExistKey("C") Then .content.RemoveOne
.content.AssignPath "C.1", CVar(a$)
.Flags = .Flags Or (fline + fselected + joinpRevline)
.Flags = .Flags Xor (fline + fselected + joinpRevline)
End With
Timer1.enabled = False
Timer1.Interval = 100
Timer1.enabled = True
End Sub
Public Sub AddRowsAtListIndex(ByVal Rows As Integer)
Dim i As Long
If itemcount <= 0 Then Exit Sub
While itemcount + Rows > Buffer
    Buffer = Buffer * 2
    ReDim Preserve mList(0 To Buffer)
Wend
itemcount = itemcount + Rows
For i = itemcount - 1 To ListIndex + Rows Step -1
    mList(i) = mList(i - Rows)
Next i
For i = ListIndex + 1 To ListIndex + Rows
With mList(i)
.Flags = joinpRevline
Set .content = Nothing
End With
Next i
mList(ListIndex).morerows = mList(ListIndex).morerows + Rows
JoinLines = JoinLines + Rows
Timer1.enabled = False
Timer1.Interval = 10
Timer1.enabled = True
End Sub
Public Sub DropRowsAtListIndex(ByVal Rows As Integer)
Dim i As Long, emp As itemlist
If itemcount <= 0 Then Exit Sub
If mList(ListIndex).morerows > 0 And mList(ListIndex).morerows < Rows Then
    Rows = mList(ListIndex).morerows
End If
If mList(ListIndex).morerows >= Rows Then
For i = ListIndex + 1 To itemcount - Rows - 1
    mList(i) = mList(i + Rows)
Next i
For i = i To itemcount - 1
    mList(i) = emp
Next i
mList(ListIndex).morerows = mList(ListIndex).morerows - Rows
JoinLines = JoinLines - Rows
Timer1.enabled = False
Timer1.Interval = 10
Timer1.enabled = True
End If
End Sub

Public Sub AddSep()
Dim i As Long

If itemcount = Buffer Then
Buffer = Buffer * 2
ReDim Preserve mList(0 To Buffer)
End If
itemcount = itemcount + 1
ListSep(itemcount - 1) = True
Timer1.enabled = False
Timer1.Interval = 100
Timer1.enabled = True
End Sub
Public Sub additemFast(a$)
Dim i As Long
If itemcount = Buffer Then
Buffer = Buffer * 2
ReDim Preserve mList(0 To Buffer)
End If
itemcount = itemcount + 1
With mList(itemcount - 1)
Set .content = New JsonObject
.content.AssignPath "C.1", CVar(a$)
.Flags = .Flags Or (fline + fselected + joinpRevline)
.Flags = .Flags Xor (fline + fselected + joinpRevline)
End With
End Sub
Public Sub Removeitem(ByVal ii As Long)
Dim i As Long
If ii < 0 Then Exit Sub
If ii = itemcount - 1 Then
Else
For i = ii + 1 To itemcount - 1
mList(i - 1) = mList(i)
Next i
End If
itemcount = itemcount - 1

If listcount < 0 Then
itemcount = 0
Clear
Exit Sub
End If
If itemcount < Buffer \ 2 And Buffer > 100 Then
Buffer = Buffer \ 2
ReDim Preserve mList(0 To Buffer)
End If
SELECTEDITEM = 0
If listcount <= lines + 1 Then
BarVisible = False
End If
Timer1.enabled = True


End Sub
Public Sub ShowMe(Optional visibleme As Boolean = False, Optional headeronly As Boolean)
Dim REALX As Long, RealX2 As Long, myt1, oldtopitem As Long, s1 As String
Dim mcenter As Boolean, mwraptext As Boolean, mVcenter As Boolean, mEditFlag1 As Boolean
On Error GoTo there
If SuspDraw Then Exit Sub
If visibleme Then
    BarWidth = UserControlTextWidth("W")
    CalcAndShowBar1
    Timer1.enabled = True: Exit Sub
End If
If listcount = 0 And HeadLine = vbNullString Then
    Repaint
    Exit Sub
End If
Dim i As Long, j As Long, G$, nr As RECT, fg As Long, hnr As RECT, skipme As Boolean, nfg As Long, tmprows As Long
If MultiSelect And LeftMarginPixels < mytPixels Then LeftMarginPixels = mytPixels
If Not headeronly Then Repaint
mEditFlag1 = mEditFlag
currentY = 0
nr.top = 0
nr.Left = 0
nr.Bottom = mytPixels + 1
hnr.Bottom = mytPixels + 1
nr.Right = Width / scrTwips
hnr.Right = Width / scrTwips
If mHeadline <> "" Then
    nr.Bottom = HeadlineHeight
    RaiseEvent ExposeRect(-1, VarPtr(nr), UserControl.hDC, skipme)
    nr.Bottom = HeadlineHeight
    CalcRectHeader UserControl.hDC, mHeadline, hnr, DT_CENTER
    If Not skipme Then
        If hnr.Bottom < HeadLineHeightMinimum Then
            hnr.Bottom = HeadLineHeightMinimum
        End If
        If Not blockheight Then
        If mHeadlineHeight <> hnr.Bottom Then
            HeadlineHeight = hnr.Bottom
            nr.Bottom = mHeadlineHeight
        End If
        End If
        If Not NoHeaderBackground Then FillBack UserControl.hDC, nr, CapColor
    End If
    hnr.top = (nr.Bottom - hnr.Bottom) \ 2
    hnr.Bottom = nr.Bottom - hnr.top
    hnr.Left = 0
    hnr.Right = nr.Right
    PrintLineControlHeader UserControl.hDC, mHeadline, hnr, DT_CENTER
    If headeronly Then
        Refresh
        Exit Sub
    End If
     nr.top = nr.Bottom
    nr.Bottom = nr.top + mytPixels + 1
End If
If AutoPanPos Then
    If SelStart = 0 Then SelStart = 1
    scrollme = 0
again123:
    REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
    RealX2 = scrollme + REALX
    If Not NoScroll Then If RealX2 > Width * 0.8 * dragslow Then scrollme = scrollme - Width * 0.2 * dragslow: GoTo again123
        If RealX2 - Width * 0.2 * dragslow < 0 Then
            If Not NoScroll Then
                scrollme = scrollme + Width * 0.2 * dragslow
                If scrollme > 0 Then scrollme = 0 Else GoTo again123
            End If
        End If
    End If
    If SingleLineSlide Then
        nr.Left = LeftMarginPixels
    Else
        nr.Left = scrollme / scrTwips + LeftMarginPixels
    End If
    If listcount = 0 Then
        BarVisible = False
        Exit Sub
    End If
    If SELECTEDITEM > TopRows Then
        oldtopitem = topitem
        topitem = 0
        j = SELECTEDITEM - lines / 2 - 1
        If j < 0 Then j = 0
        If listcount <= lines + 1 Then
           topitem = 0
        Else
            If j + lines >= listcount Then
                If listcount - lines >= 0 Then
                    topitem = listcount - lines - 1
                End If
            Else
                If dragslow < 1 Or Not MultiLineEditBox Then
                If Not MultiLineEditBox Then
                    If SELECTEDITEM - oldtopitem > 0 And SELECTEDITEM - oldtopitem <= lines + 1 Then
                        topitem = oldtopitem
                    ElseIf SELECTEDITEM - oldtopitem > lines Then
                        topitem = SELECTEDITEM - lines + 1
                    Else
                        topitem = SELECTEDITEM - 1
                    End If
                    j = topitem
                Else
                    topitem = oldtopitem
                    j = oldtopitem
                End If
            Else
                topitem = j
            End If
        End If
        state = True
        On Error Resume Next
        Err.Clear
        If Not Spinner Then
            If listcount - 1 - lines < 1 Then
                Max = 1
            Else
                Max = listcount - 1 - lines
            End If
            If Err.Number > 0 Then
                Value = listcount - 1  ' ??
                Max = listcount - 1
            Else
                Value = j
            End If
        End If
        state = False
    End If
Else
    state = True
    On Error Resume Next
    Err.Clear
    If Not Spinner Then
        Max = listcount - 1
        If Err.Number > 0 Then
            Value = listcount - 1
            Max = listcount - 1
        End If
    End If
    state = False
End If
j = topitem + lines
If j >= listcount Then j = listcount - 1
If topitem < TopRows Then topitem = TopRows: j = j + TopRows

If listcount > 0 Then
    currentX = scrollme
    DrawStyle = vbSolid
    fg = Me.ForeColor
    If havefocus Or dragfocus Then
        caretCreated = False
        DestroyCaret
    End If
    Dim onemore As Long
    If restrictLines = 0 Then
        If nr.top + (topitem + TopRows - j + 1) * mytPixels < HeightPixels Then
            onemore = 1
        End If
    End If
    Dim k As Integer, onr As Long, sum As Long, hRgn As Long, ii As Long
    k = 1
    For ii = 0 To TopRows - 1
        i = ii
        If i = SELECTEDITEM - 1 Then
            tmprows = topitem - TopRows
        End If
        onr = nr.Right
        
        RaiseEvent ExposeRect(i, VarPtr(nr), UserControl.hDC, skipme)
        
        If Not skipme Then
                Do While Me.ListJoin(ii + 1)
                    ii = ii + 1
                    nr.Bottom = nr.Bottom + mytPixels
                Loop
            If i = SELECTEDITEM - 1 And Not NoCaretShow And Not ListSep(i) Then
            
                If mTabs = 1 Then
                    nr.Left = scrollme / scrTwips + LeftMarginPixels
                Else
                    nr.Left = LeftMarginPixels
                End If
                nfg = fg
                
                
                RaiseEvent SpecialColor(nfg)
                If mTabs > 1 Then
                    For k = 1 To mTabs
                    
                     If mParts(k) > 0 Then
                        mEditFlag1 = CBool(PropAtColumnNum(i, k, "Edit"))
                        mcenter = CBool(PropAtColumnNum(i, k, "CenterText"))
                        mwraptext = CBool(PropAtColumnNum(i, k, "WrapText"))
                        mVcenter = CBool(PropAtColumnNum(i, k, "VerticalCenterText"))
                        If (NarrowSelect And mCurTab = k) Or Not NarrowSelect Then
                            nfg = fg
                            If nfg <> fg Then Me.ForeColor = nfg
                            If mEditFlag1 Then
                                If nfg = fg Then Me.ForeColor = fg
                            ElseIf nfg = fg Then
                                If Me.BackColor = 0 Then
                                    Me.ForeColor = &HFFFFFF
                                Else
                                    Me.ForeColor = 0
                                End If
                            End If
                        End If
                        nr.Right = nr.Left + mParts(k) - 1
                        nr.Right = nr.Right + 1
                        nr.Bottom = nr.Bottom + 1
                        hRgn = CreateRectRgnIndirect(nr)
                        nr.Right = nr.Right - 1
                        nr.Bottom = nr.Bottom - 1
                        SelectClipRgn UserControl.hDC, hRgn
                        RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                        If Not skipme Then
                            If mwraptext Then
                                hnr = nr
                                s1 = listAtColumn(i, k)
                                LineAddTopOffsetPixels s1, hnr
                                PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                            Else
                                hnr = nr
                                If mEditFlag1 And k = mCurTab Then hnr.Bottom = nr.top + mytPixels + 1
                                PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), hnr, mwraptext, mcenter, mVcenter
                            End If
                        End If
                        SelectClipRgn UserControl.hDC, &H0
                        DeleteObject hRgn
                        nr.Left = nr.Right + 1
                        Me.ForeColor = fg
                        If nr.Left >= onr Then Exit For
                     End If
                    Next k
                Else
                    If nfg <> fg Then Me.ForeColor = nfg
                    If mEditFlag1 Then
                        If nfg = fg Then Me.ForeColor = fg
                    ElseIf nfg = fg Then
                        If Me.BackColor = 0 Then
                            Me.ForeColor = &HFFFFFF
                        Else
                            Me.ForeColor = 0
                        End If
                    End If
                    If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                        MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i, True
                    End If
                    PrintLineControlSingle UserControl.hDC, list(i), nr
                End If
                Me.ForeColor = fg
            Else
                If ListSep(i) And list(i) = vbNullString And Not ListJoin(i) Then
                   hnr.Left = 0
                   hnr.Right = nr.Right
                   hnr.top = nr.top + mytPixels \ 2
                   hnr.Bottom = hnr.top + 1
                   FillBack UserControl.hDC, hnr, ForeColor
                Else
                    If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                        MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i
                    End If
                    If ListSep(i) Then
                        ForeColor = dcolor
                    End If
                    If mTabs > 1 Then
                        For k = 1 To mTabs
                            If mParts(k) > 0 Then
                            mcenter = CBool(PropAtColumnNum(i, k, "CenterText"))
                            mwraptext = CBool(PropAtColumnNum(i, k, "WrapText"))
                            mVcenter = CBool(PropAtColumnNum(i, k, "VerticalCenterText"))
                            nr.Right = nr.Left + mParts(k) - 1
                            nr.Right = nr.Right + 1
                            nr.Bottom = nr.Bottom + 1
                            hRgn = CreateRectRgnIndirect(nr)
                            nr.Right = nr.Right - 1
                            nr.Bottom = nr.Bottom - 1
                            SelectClipRgn UserControl.hDC, hRgn
                            RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                            If Not skipme Then
                                If mwraptext Then
                                    hnr = nr
                                    s1 = listAtColumn(i, k)
                                    LineAddTopOffsetPixels s1, hnr
                                    PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                                Else
                                    PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), nr, mwraptext, mcenter, mVcenter
                                End If
                            End If
                            SelectClipRgn UserControl.hDC, &H0
                            DeleteObject hRgn
                            nr.Left = nr.Right + 1
                            If nr.Left >= onr Then Exit For
                            End If
                        Next k
                        
                        nr.Right = onr
                    Else
                        PrintLineControlSingle UserControl.hDC, list(i), nr
                    End If
                        
                    End If
            End If
            If SingleLineSlide Then
                nr.Left = LeftMarginPixels
            Else
                nr.Left = scrollme / scrTwips + LeftMarginPixels
            End If
        End If
        nr.top = nr.top + mytPixels * (ii - i + 1)
        nr.Bottom = nr.top + mytPixels + 1
        ForeColor = fg
    
    Next ii
    k = 1
    For ii = topitem To j + onemore
        i = ii
        onr = nr.Right
        RaiseEvent ExposeRect(i, VarPtr(nr), UserControl.hDC, skipme)
        If Not skipme Then
            Do While Me.ListJoin(ii + 1)
                ii = ii + 1
                nr.Bottom = nr.Bottom + mytPixels
            Loop
            If i = SELECTEDITEM - 1 And Not NoCaretShow And Not ListSep(i) Then
                If mTabs = 1 Then
                    nr.Left = scrollme / scrTwips + LeftMarginPixels
                Else
                    nr.Left = LeftMarginPixels
                End If
                nfg = fg
                
                
                RaiseEvent SpecialColor(nfg)
                If mTabs > 1 Then
                    For k = 1 To mTabs
                     If mParts(k) > 0 Then
                        mEditFlag1 = CBool(PropAtColumnNum(i, k, "Edit"))
                        mcenter = CBool(PropAtColumnNum(i, k, "CenterText"))
                        mwraptext = CBool(PropAtColumnNum(i, k, "WrapText"))
                        mVcenter = CBool(PropAtColumnNum(i, k, "VerticalCenterText"))
                        If (NarrowSelect And mCurTab = k) Or Not NarrowSelect Then
                            nfg = fg
                            If nfg <> fg Then Me.ForeColor = nfg
                            If mEditFlag1 Then
                                If nfg = fg Then Me.ForeColor = fg
                            ElseIf nfg = fg Then
                                If Me.BackColor = 0 Then
                                    Me.ForeColor = &HFFFFFF
                                Else
                                    Me.ForeColor = 0
                                End If
                            End If
                        End If
                        nr.Right = nr.Left + mParts(k)
                        nr.Bottom = nr.Bottom + 1
                        hRgn = CreateRectRgnIndirect(nr)
                        nr.Right = nr.Right - 1
                        nr.Bottom = nr.Bottom - 1
                        SelectClipRgn UserControl.hDC, hRgn
                        RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                        If Not skipme Then
                            If mwraptext Then
                                hnr = nr
                                s1 = listAtColumn(i, k)
                                LineAddTopOffsetPixels s1, hnr
                                PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                            Else
                                hnr = nr
                                If mEditFlag1 And k = mCurTab Then hnr.Bottom = nr.top + mytPixels + 1
                                PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), hnr, mwraptext, mcenter, mVcenter
                            End If
                        End If
                        SelectClipRgn UserControl.hDC, &H0
                        DeleteObject hRgn
                        nr.Left = nr.Right + 1
                        Me.ForeColor = fg
                        If nr.Left >= onr Then Exit For
                     End If
                    Next k
                Else
                    If nfg <> fg Then Me.ForeColor = nfg
                    If mEditFlag1 Then
                        If nfg = fg Then Me.ForeColor = fg
                    ElseIf nfg = fg Then
                        If Me.BackColor = 0 Then
                            Me.ForeColor = &HFFFFFF
                        Else
                            Me.ForeColor = 0
                        End If
                    End If
                    If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                        MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i, True
                    End If
                    PrintLineControlSingle UserControl.hDC, list(i), nr
                End If
                Me.ForeColor = fg
            Else
                If ListSep(i) And list(i) = vbNullString And Not ListJoin(i) Then
                   hnr.Left = 0
                   hnr.Right = nr.Right
                   hnr.top = nr.top + mytPixels \ 2
                   hnr.Bottom = hnr.top + 1
                   FillBack UserControl.hDC, hnr, ForeColor
                Else
                    If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                        MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i
                    End If
                    If ListSep(i) Then
                        ForeColor = dcolor
                    End If
                    If mTabs > 1 Then
                        For k = 1 To mTabs
                            If mParts(k) > 0 Then
                                mcenter = CBool(PropAtColumnNum(i, k, "CenterText"))
                                mwraptext = CBool(PropAtColumnNum(i, k, "WrapText"))
                                mVcenter = CBool(PropAtColumnNum(i, k, "VerticalCenterText"))
                                nr.Right = nr.Left + mParts(k)
                                nr.Bottom = nr.Bottom + 1
                                hRgn = CreateRectRgnIndirect(nr)
                                nr.Right = nr.Right - 1
                                nr.Bottom = nr.Bottom - 1
                                SelectClipRgn UserControl.hDC, hRgn
                                RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                                If Not skipme Then
                                    If mwraptext Then
                                        hnr = nr
                                        s1 = listAtColumn(i, k)
                                        LineAddTopOffsetPixels s1, hnr
                                        PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                                    Else
                                        PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), nr, mwraptext, mcenter, mVcenter
                                    End If
                                End If
                                SelectClipRgn UserControl.hDC, &H0
                                DeleteObject hRgn
                                nr.Left = nr.Right + 1
                                If nr.Left >= onr Then Exit For
                            End If
                        Next k
                        nr.Right = onr
                    Else
                        PrintLineControlSingle UserControl.hDC, list(i), nr
                    End If
                End If
            End If
            If SingleLineSlide Then
                nr.Left = LeftMarginPixels
            Else
                nr.Left = scrollme / scrTwips + LeftMarginPixels
            End If
        End If
        nr.top = nr.top + mytPixels * (ii - i + 1)
        nr.Bottom = nr.top + mytPixels + 1
        ForeColor = fg
    Next ii
    DrawMode = vbInvert
    myt1 = myt - scrTwips
    
    If SELECTEDITEM > 0 Then
        i = SELECTEDITEM - 1
        ii = i
        Do While ListJoin(ii + 1)
            ii = ii + 1
        Loop
        mEditFlag1 = CBool(PropAtColumnNum(i, mCurTab, "Edit"))
        If Not NoCaretShow And Not mEditFlag1 And i <> ii Then
            If ii - topitem - 1 <= lines + onemore And (ii > topitem Or ii <= TopRows) And Not ListSep(ii - 1) Then
                i = ii - i
                GoTo there1
            End If
        End If
        i = ii - i
        If SELECTEDITEM - topitem - 1 <= lines + onemore And (SELECTEDITEM > topitem Or SELECTEDITEM <= TopRows) And Not ListSep(SELECTEDITEM - 1) Then
            mEditFlag1 = CBool(PropAtColumnNum(SELECTEDITEM - 1, mCurTab, "Edit"))
            mcenter = CBool(PropAtColumnNum(SELECTEDITEM - 1, mCurTab, "CenterText"))
            mwraptext = CBool(PropAtColumnNum(SELECTEDITEM - 1, mCurTab, "WrapText"))
            mVcenter = CBool(PropAtColumnNum(SELECTEDITEM - 1, mCurTab, "VerticalCenterText"))
            If Not NoCaretShow Then
                If mEditFlag1 Then
                    If SelStart = 0 Then SelStart = 1
                    DrawStyle = vbSolid
                    RaiseEvent PureListOff
                    If mcenter Then
                        If mTabs = 1 Then
                            REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + (UserControl.ScaleWidth - UserControlTextWidth(list$(SELECTEDITEM - 1))) / 2 + LeftMarginPixels * scrTwips
                            RealX2 = scrollme / 2 + REALX
                        Else
                            REALX = UserControlTextWidth(Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), 1, SelStart - 1)) + (mParts(mCurTab) * scrTwips - UserControlTextWidth(listAtColumn(SELECTEDITEM - 1, mCurTab))) / 2 + TwipsCurTab + LeftMarginPixels * scrTwips
                            RealX2 = scrollme / 2 + REALX
                        End If
                    Else
                        If mTabs = 1 Then
                            skipme = False
                            RaiseEvent GetRealX1(UserControl.hDC, SelStart, list(SELECTEDITEM - 1), REALX, skipme)
                            If skipme Then
                                skipme = False
                                REALX = (REALX + LeftMarginPixels) * scrTwips
                            Else
                                REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
                            End If
                            RealX2 = scrollme + REALX
                        Else
                            skipme = False
                            RaiseEvent GetRealX1(UserControl.hDC, SelStart, listAtColumn(SELECTEDITEM - 1, mCurTab), REALX, skipme)
                            If skipme Then
                                skipme = False
                                REALX = (REALX + LeftMarginPixels) * scrTwips + TwipsCurTab
                            Else
                                REALX = UserControlTextWidth(Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), 1, SelStart - 1)) + LeftMarginPixels * scrTwips + TwipsCurTab
                            End If
                            RealX2 = scrollme + REALX
                        End If
                    End If
                    RaiseEvent PureListOn
                    If InternalCursor Or dragfocus Then
                        ' not used ???
                        If Noflashingcaret Or Not havefocus Then
                            Line (scrollme + REALX, (SELECTEDITEM + TopRows - topitem - 1 + tmprows) * myt + myt1 + mHeadlineHeightTwips)-(RealX2, (SELECTEDITEM + TopRows - topitem - 1 - tmprows) * myt + mHeadlineHeightTwips), ForeColor
                        Else
                            If mTabs > 1 Then
                                If RealX2 < TwipsCurTab Then
                                ElseIf RealX2 > TwipsCurTab + mParts(mCurTab) * scrTwips Then
                                Else
                                    ShowMyCaretInTwips RealX2, (SELECTEDITEM + TopRows - topitem - 1 + tmprows) * myt + mHeadlineHeightTwips
                                End If
                            Else
                                ShowMyCaretInTwips RealX2, (SELECTEDITEM + TopRows - topitem - 1 + tmprows) * myt + mHeadlineHeightTwips
                            End If
                        End If
                    End If
                    If Not NoScroll Then If RealX2 > Width * 0.8 * dragslow Then scrollme = scrollme - Width * 0.2 * dragslow: PrepareToShow 10
                    If RealX2 - Width * 0.2 * dragslow < 0 Then
                        If Not NoScroll Then
                            scrollme = scrollme + Width * 0.2 * dragslow
                            If scrollme > 0 Then scrollme = 0 Else PrepareToShow 10
                        End If
                    End If
                Else
there1:
                    DrawStyle = vbInvisible
                    
                    If mTabs = 1 Or Not NarrowSelect Then
                        If BackStyle = 1 Then
                            Line (scrTwips, (i + SELECTEDITEM - topitem + TopRows + tmprows) * myt + mHeadlineHeightTwips)-(0 + UserControl.Width, (SELECTEDITEM - topitem - 1 + TopRows + tmprows) * myt + mHeadlineHeightTwips - scrTwips / 2), 0, B
                        Else
                            Line (0, (i + SELECTEDITEM - topitem + TopRows + tmprows) * myt + mHeadlineHeightTwips)-(0 + UserControl.Width, (SELECTEDITEM - topitem + TopRows - 1 + tmprows) * myt + mHeadlineHeightTwips), 0, B
                        End If
                    Else
                         If BackStyle = 1 Then
                            Line (scrTwips + TwipsCurTab, (i + SELECTEDITEM - topitem + TopRows + tmprows) * myt + mHeadlineHeightTwips)-(TwipsCurTab + mParts(mCurTab) * scrTwips - 2 * scrTwips, (SELECTEDITEM - topitem - 1 + TopRows + tmprows) * myt + mHeadlineHeightTwips), 0, B
                        Else
                            Line (TwipsCurTab, (i + SELECTEDITEM - topitem + TopRows + tmprows) * myt + mHeadlineHeightTwips)-(TwipsCurTab + mParts(mCurTab) * scrTwips - scrTwips, (SELECTEDITEM - topitem + TopRows - 1 + tmprows) * myt + mHeadlineHeightTwips), 0, B
                        End If
                    End If
                End If
            End If
        Else
        
            HideCaret (hWnd)
        End If
    End If
    currentY = 0
    currentX = 0
    DrawMode = vbCopyPen
End If
DrawStyle = vbSolid
LastVScroll = Value
If Grid Then
    If mHeadlineHeightTwips > 0 Then
        Line (0, mHeadlineHeightTwips)-(UserControl.Width, mHeadlineHeightTwips), GridColor
    End If
    For i = 1 To lines + 1 + TopRows
        If i <= TopRows Then ii = i Else ii = i + topitem - TopRows
        If Not ListJoin(ii) Then
            Line (0, i * myt + mHeadlineHeightTwips)-(UserControl.Width, i * myt + mHeadlineHeightTwips), GridColor
        End If
    Next
           
    
    If mTabs > 1 Then
    sum = -scrTwips
    Line (sum, mHeadlineHeightTwips)-(sum, UserControl.Height), GridColor
    For i = 1 To mTabs
        sum = sum + mParts(i) * scrTwips
        Line (sum, mHeadlineHeightTwips)-(sum, UserControl.Height), GridColor
    Next
    End If
    End If
         If BorderStyle = 1 And BackStyle = 1 Then
                hnr.Left = 0
                hnr.Right = Me.WidthPixels
                hnr.top = 0
                hnr.Bottom = Me.HeightPixels
                onemore = CreateSolidBrush(Me.BackColor)
                FrameRect UserControl.hDC, hnr, onemore
                DeleteObject onemore
            End If
RepaintScrollBar
RaiseEvent ScrollMove(topitem)
there:
End Sub

Public Sub ShowMe2()
On Error GoTo there
If SuspDraw Then Exit Sub
Dim nr As RECT, j As Long, i As Long, skipme As Boolean, fg As Long, hnr As RECT, nfg As Long, nfg1 As Long
Dim REALX As Long, RealX2 As Long, myt1, s1 As String
Dim mcenter As Boolean, mwraptext As Boolean, mVcenter As Boolean, mEditFlag1 As Boolean
BarWidth = UserControlTextWidth("W")
If listcount = 0 And HeadLine = vbNullString Then
    Repaint
    HideCaret (hWnd)
    Exit Sub
End If
If MultiSelect And LeftMarginPixels < mytPixels Then LeftMarginPixels = mytPixels
Repaint
nr.top = 0
nr.Left = 0 '
hnr.Left = 0  ' no scrolling
nr.Bottom = mytPixels + 1
hnr.Bottom = mytPixels + 1
nr.Right = Width / scrTwips
hnr.Right = Width / scrTwips
If mHeadline <> "" Then
    nr.Bottom = HeadlineHeight
    RaiseEvent ExposeRect(-1, VarPtr(nr), UserControl.hDC, skipme)
    nr.Bottom = HeadlineHeight
    CalcRectHeader UserControl.hDC, mHeadline, hnr, DT_CENTER
    If Not skipme Then
        If Not blockheight Then
            If mHeadlineHeight <> hnr.Bottom Then
                HeadlineHeight = hnr.Bottom
                nr.Bottom = mHeadlineHeight
            End If
        End If
        If Not NoHeaderBackground Then FillBack UserControl.hDC, nr, CapColor
        
    End If
    hnr.top = (nr.Bottom - hnr.Bottom) \ 2
    hnr.Bottom = nr.Bottom - hnr.top
    hnr.Left = 0
    hnr.Right = nr.Right
    PrintLineControlHeader UserControl.hDC, mHeadline, hnr, DT_CENTER
    nr.top = nr.Bottom
    nr.Bottom = nr.top + mytPixels + 1
End If
If AutoPanPos Then
    If SelStart = 0 Then SelStart = 1
    scrollme = 0
again123:
    REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
    RealX2 = scrollme + REALX
    If Not NoScroll Then If RealX2 > Width * 0.8 * dragslow Then scrollme = scrollme - Width * 0.2 * dragslow: GoTo again123
    If RealX2 - Width * 0.2 * dragslow < 0 Then
        If Not NoScroll Then
            scrollme = scrollme + Width * 0.2 * dragslow
            If scrollme > 0 Then scrollme = 0 Else GoTo again123
        End If
    End If
End If
If SingleLineSlide Then
    nr.Left = LeftMarginPixels
Else
    nr.Left = scrollme / scrTwips + LeftMarginPixels
End If
If topitem < TopRows Then topitem = TopRows
j = topitem + lines
Dim tmprows As Long
If j >= listcount Then j = listcount - 1
If listcount = 0 Then
    BarVisible = False
    Exit Sub
Else
    DrawStyle = vbSolid
    If havefocus Or dragfocus Then
        caretCreated = False
        DestroyCaret
    End If
    fg = Me.ForeColor
    nfg = fg
    nfg1 = fg
    RaiseEvent SpecialColor(nfg1)
    Dim onemore As Long, ii As Long
    If restrictLines = 0 Then
        If nr.top + (topitem + TopRows - j + 1) * mytPixels < HeightPixels Then
            onemore = 1
        End If
    End If
    Dim k As Integer, onr As Long, sum As Long, hRgn As Long, ont As Long
    For ii = 0 To TopRows - 1
        i = ii
        currentX = scrollme
        currentY = 0
        onr = nr.Right
        ont = nr.top
        RaiseEvent ExposeRect(i, VarPtr(nr), UserControl.hDC, skipme)
        If i = SELECTEDITEM - 1 Then
            tmprows = topitem - TopRows
        End If
        If Not skipme Then
                Do While Me.ListJoin(ii + 1)
                    ii = ii + 1
                    nr.Bottom = nr.Bottom + mytPixels
                Loop
            If i = SELECTEDITEM - 1 And Not NoCaretShow And Not ListSep(i) Then
            
                If mTabs = 1 Then
                        nr.Left = scrollme / scrTwips + LeftMarginPixels
                Else
                    nr.Left = LeftMarginPixels
                End If
                If mTabs > 1 Then
                    nr.Left = LeftMarginPixels
                    For k = 1 To mCurTab - 1
                        nr.Left = nr.Left + mParts(k)
                    Next k
                    mEditFlag1 = CBool(PropAtColumnNum(i, mCurTab, "EDIT"))
                    mcenter = CBool(PropAtColumnNum(i, mCurTab, "CenterText"))
                    mwraptext = CBool(PropAtColumnNum(i, mCurTab, "WrapText"))
                    mVcenter = CBool(PropAtColumnNum(i, mCurTab, "VerticalCenterText"))
                    k = mCurTab
                    nr.Right = nr.Left + mParts(k) - 1
                    hnr = nr
                    nr.Left = nr.Left + scrollme / scrTwips
                    hnr.Right = nr.Right + 1
                    hnr.Bottom = nr.Bottom + 1
                    hRgn = CreateRectRgnIndirect(hnr)
                    SelectClipRgn UserControl.hDC, hRgn
                    RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                    If Not skipme Then
                        If mwraptext And Not mEditFlag1 Then
                            hnr = nr
                            s1 = listAtColumn(i, k)
                            LineAddTopOffsetPixels s1, hnr
                            PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                        Else
                            hnr = nr
                            If mEditFlag1 Then hnr.Bottom = nr.top + mytPixels + 1
                            PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), hnr, False, mcenter, mVcenter
                        End If
                    End If
                    SelectClipRgn UserControl.hDC, &H0
                    DeleteObject hRgn
                    nr.Right = onr
                    sum = LeftMarginPixels
                    nr.Left = sum
                    For k = 1 To mTabs
                        If mParts(k) > 0 Then
                            mcenter = CBool(PropAtColumnNum(i, k, "CenterText"))
                            mwraptext = CBool(PropAtColumnNum(i, k, "WrapText"))
                            mVcenter = CBool(PropAtColumnNum(i, k, "VerticalCenterText"))
                            If (NarrowSelect And mCurTab = k) Or Not NarrowSelect Then
                                nfg = fg
                                If nfg1 <> nfg Then nfg = nfg1
                                If nfg <> fg Then Me.ForeColor = nfg
                                If mEditFlag1 Then
                                    If nfg = fg Then Me.ForeColor = fg
                                ElseIf nfg = fg Then
                                    If Me.BackColor = 0 Then
                                        Me.ForeColor = &HFFFFFF
                                    Else
                                        Me.ForeColor = 0
                                    End If
                                End If
                            End If
                            nr.Right = nr.Left + mParts(k)
                            If k <> mCurTab Then
                                nr.Right = nr.Left + mParts(k) - 1
                                nr.Right = nr.Right + 1
                                nr.Bottom = nr.Bottom + 1
                                hRgn = CreateRectRgnIndirect(nr)
                                nr.Right = nr.Right - 1
                                nr.Bottom = nr.Bottom - 1
                                SelectClipRgn UserControl.hDC, hRgn
                                RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                                If Not skipme Then
                                    If mwraptext Then
                                        hnr = nr
                                        s1 = listAtColumn(i, k)
                                        LineAddTopOffsetPixels s1, hnr
                                        PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                                    Else
                                        PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), nr, mwraptext, mcenter, mVcenter
                                    End If
                                End If
                                SelectClipRgn UserControl.hDC, &H0
                                DeleteObject hRgn
                            End If
                            sum = sum + mParts(k)
                            nr.Left = sum  ' nr.Right + 1
                            ForeColor = fg
                            If nr.Left >= onr Then Exit For
                        End If
                    Next k
                Else
                    nfg = nfg1
                    If nfg <> fg Then Me.ForeColor = nfg
                    If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                        MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i, True
                    End If
                    PrintLineControlSingle UserControl.hDC, list(i), nr
                End If
                If nfg = fg Then Me.ForeColor = fg
            Else
                nfg = fg
                'RaiseEvent SpecialColor(nfg)
                If nfg1 <> nfg Then nfg = nfg1
                If ListSep(i) And list(i) = vbNullString And Not ListJoin(i) Then
                    hnr.Left = 0
                    hnr.Right = nr.Right
                    hnr.top = nr.top + mytPixels \ 2
                    hnr.Bottom = hnr.top + 1
                    FillBack UserControl.hDC, hnr, ForeColor
                Else
                    If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                        MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i
                    End If
                    If ListSep(i) Then
                        ForeColor = dcolor
                    Else
                        nfg = nfg1
                        If nfg = fg Then ForeColor = fg
                        If SELECTEDITEM - 1 = i And nfg <> fg Then
                            Me.ForeColor = nfg
                        End If
                    End If
                    If mTabs > 1 Then
                        sum = LeftMarginPixels
                        nr.Left = sum
                        currentX = scrollme
                        For k = 1 To mTabs
                            If mParts(k) > 0 Then
                            mEditFlag1 = CBool(PropAtColumnNum(i, mCurTab, "EDIT"))
                            mcenter = CBool(PropAtColumnNum(i, k, "CenterText"))
                            mwraptext = CBool(PropAtColumnNum(i, k, "WrapText"))
                            mVcenter = CBool(PropAtColumnNum(i, k, "VerticalCenterText"))
                            nr.Right = nr.Left + mParts(k) - 1
                            nr.Right = nr.Right + 1
                            nr.Bottom = nr.Bottom + 1
                            hnr = nr
                            hnr.top = ont
                            hRgn = CreateRectRgnIndirect(hnr)
                            nr.Right = nr.Right - 1
                            nr.Bottom = nr.Bottom - 1
                            SelectClipRgn UserControl.hDC, hRgn
                            RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                            If Not skipme Then
                                If mwraptext Then
                                    hnr = nr
                                    s1 = listAtColumn(i, k)
                                    LineAddTopOffsetPixels s1, hnr
                                    PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                                Else
                                    PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), nr, mwraptext, mcenter, mVcenter
                                End If
                            End If
                            SelectClipRgn UserControl.hDC, &H0
                            DeleteObject hRgn
                            sum = sum + mParts(k)
                            nr.Left = sum ' nr.Right + 1
                            If nr.Left >= onr Then Exit For
                            End If
                        Next k
                    Else
                        PrintLineControlSingle UserControl.hDC, list(i), nr
                    End If
                End If
            End If
            If mTabs = 1 Then
                If SingleLineSlide Then
                    nr.Left = LeftMarginPixels
                Else
                    nr.Left = scrollme / scrTwips + LeftMarginPixels
                End If
            Else
                nr.Left = LeftMarginPixels
            End If
        End If
        nr.top = nr.top + mytPixels * (ii - i + 1)
        nr.Bottom = nr.top + mytPixels + 1
        ForeColor = fg
    Next ii
    For ii = topitem To j + onemore
        i = ii
        currentX = scrollme
        currentY = 0
        onr = nr.Right
        ont = nr.top
        RaiseEvent ExposeRect(i, VarPtr(nr), UserControl.hDC, skipme)
        If Not skipme Then
                If Me.ListJoin(i) Then
                    Do While Me.ListJoin(i)
                        i = i - 1
                        nr.top = nr.top - mytPixels
                    Loop
                End If
                Do While Me.ListJoin(ii + 1)
                    ii = ii + 1
                    nr.Bottom = nr.Bottom + mytPixels
                Loop
            If i = SELECTEDITEM - 1 And Not NoCaretShow And Not ListSep(i) Then
                
                If mTabs = 1 Then
                        nr.Left = scrollme / scrTwips + LeftMarginPixels
                Else
                    nr.Left = LeftMarginPixels
                End If
                If mTabs > 1 Then
                    nr.Left = LeftMarginPixels
                    For k = 1 To mCurTab - 1
                        nr.Left = nr.Left + mParts(k)
                    Next k
                    k = mCurTab
                    mEditFlag1 = CBool(PropAtColumnNum(i, mCurTab, "EDIT"))
                    mcenter = CBool(PropAtColumnNum(i, mCurTab, "CenterText"))
                    mwraptext = CBool(PropAtColumnNum(i, mCurTab, "WrapText"))
                    mVcenter = CBool(PropAtColumnNum(i, mCurTab, "VerticalCenterText"))
                    nr.Right = nr.Left + mParts(k) - 1
                    hnr = nr
                    nr.Left = nr.Left + scrollme / scrTwips
                    hnr.Right = nr.Right + 1
                    hnr.Bottom = nr.Bottom + 1
                    hnr.top = ont
                    hRgn = CreateRectRgnIndirect(hnr)
                    SelectClipRgn UserControl.hDC, hRgn
                    RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                    If Not skipme Then
                        If mwraptext And Not mEditFlag1 Then
                            hnr = nr
                            s1 = listAtColumn(i, k)
                            LineAddTopOffsetPixels s1, hnr
                            PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                        Else
                            hnr = nr
                            If mEditFlag1 Then hnr.Bottom = nr.top + mytPixels + 1
                            PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), hnr, False, mcenter, mVcenter
                        End If
                    End If
                    SelectClipRgn UserControl.hDC, &H0
                    DeleteObject hRgn
                    nr.Right = onr
                    sum = LeftMarginPixels
                    nr.Left = sum
                    For k = 1 To mTabs
                        If mParts(k) > 0 Then
                            mEditFlag1 = CBool(PropAtColumnNum(i, mCurTab, "EDIT"))
                            mcenter = CBool(PropAtColumnNum(i, k, "CenterText"))
                            mwraptext = CBool(PropAtColumnNum(i, k, "WrapText"))
                            mVcenter = CBool(PropAtColumnNum(i, k, "VerticalCenterText"))
                            If (NarrowSelect And mCurTab = k) Or Not NarrowSelect Then
                                nfg = fg
                                If nfg1 <> nfg Then nfg = nfg1
                                If nfg <> fg Then Me.ForeColor = nfg
                                If mEditFlag1 Then
                                    If nfg = fg Then Me.ForeColor = fg
                                ElseIf nfg = fg Then
                                    If Me.BackColor = 0 Then
                                        Me.ForeColor = &HFFFFFF
                                    Else
                                        Me.ForeColor = 0
                                    End If
                                End If
                            End If
                            nr.Right = nr.Left + mParts(k)
                            If k <> mCurTab Then
                                nr.Right = nr.Left + mParts(k) - 1
                                nr.Right = nr.Right + 1
                                nr.Bottom = nr.Bottom + 1
                                hnr = nr
                                hnr.top = ont
                                hRgn = CreateRectRgnIndirect(hnr)
                                nr.Right = nr.Right - 1
                                nr.Bottom = nr.Bottom - 1
                                SelectClipRgn UserControl.hDC, hRgn
                                RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                                If Not skipme Then
                                    If mwraptext Then
                                        hnr = nr
                                        s1 = listAtColumn(i, k)
                                        LineAddTopOffsetPixels s1, hnr
                                        PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                                    Else
                                        PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), nr, mwraptext, mcenter, mVcenter
                                    End If
                                End If
                                SelectClipRgn UserControl.hDC, &H0
                                DeleteObject hRgn
                            End If
                            sum = sum + mParts(k)
                            nr.Left = sum  ' nr.Right + 1
                            ForeColor = fg
                            If nr.Left >= onr Then Exit For
                        End If
                    Next k
                Else
                nfg = nfg1
                If nfg <> fg Then Me.ForeColor = nfg
                
                
                If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                    MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i, True
                End If
                    PrintLineControlSingle UserControl.hDC, list(i), nr
                End If
                If nfg = fg Then Me.ForeColor = fg
            Else
                nfg = fg
                'RaiseEvent SpecialColor(nfg)
                If nfg1 <> nfg Then nfg = nfg1
                If ListSep(i) And list(i) = vbNullString And Not ListJoin(i) Then
                    hnr.Left = 0
                    hnr.Right = nr.Right
                    hnr.top = nr.top + mytPixels \ 2
                    hnr.Bottom = hnr.top + 1
                    FillBack UserControl.hDC, hnr, ForeColor
                Else
                    If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                        MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i
                    End If
                    If ListSep(i) Then
                        ForeColor = dcolor
                    Else
                        nfg = nfg1
                        If nfg = fg Then ForeColor = fg
                        If SELECTEDITEM - 1 = i And nfg <> fg Then
                            Me.ForeColor = nfg
                        End If
                    End If
                    If mTabs > 1 Then
                        sum = LeftMarginPixels
                        nr.Left = sum
                        currentX = scrollme
                        For k = 1 To mTabs
                            If mParts(k) > 0 Then
                            mcenter = CBool(PropAtColumnNum(i, k, "CenterText"))
                            mwraptext = CBool(PropAtColumnNum(i, k, "WrapText"))
                            mVcenter = CBool(PropAtColumnNum(i, k, "VerticalCenterText"))
                            nr.Right = nr.Left + mParts(k) - 1
                            nr.Right = nr.Right + 1
                            nr.Bottom = nr.Bottom + 1
                            hnr = nr
                            hnr.top = ont
                            hRgn = CreateRectRgnIndirect(hnr)
                            nr.Right = nr.Right - 1
                            nr.Bottom = nr.Bottom - 1
                            SelectClipRgn UserControl.hDC, hRgn
                            RaiseEvent ExposeRectCol(i, k, VarPtr(nr), UserControl.hDC, skipme)
                            If Not skipme Then
                                If mwraptext Then
                                    hnr = nr
                                    s1 = listAtColumn(i, k)
                                    LineAddTopOffsetPixels s1, hnr
                                    PrintLineControlSinglePrivate UserControl.hDC, s1, hnr, mwraptext, mcenter, mVcenter
                                Else
                                    PrintLineControlSinglePrivate UserControl.hDC, listAtColumn(i, k), nr, mwraptext, mcenter, mVcenter
                                End If
                            End If
                            SelectClipRgn UserControl.hDC, &H0
                            DeleteObject hRgn
                            sum = sum + mParts(k)
                            nr.Left = sum ' nr.Right + 1
                            If nr.Left >= onr Then Exit For
                            End If
                        Next k
                    Else
                        PrintLineControlSingle UserControl.hDC, list(i), nr
                    End If
                End If
            End If
            If mTabs = 1 Then
                If SingleLineSlide Then
                    nr.Left = LeftMarginPixels
                Else
                    nr.Left = scrollme / scrTwips + LeftMarginPixels
                End If
            Else
                nr.Left = LeftMarginPixels
            End If
        End If
        nr.top = nr.top + mytPixels * (ii - i + 1)
        nr.Bottom = nr.top + mytPixels + 1
        ForeColor = fg
    Next ii
    myt1 = myt - scrTwips
    DrawMode = vbInvert
    If SELECTEDITEM > 0 Then
            i = SELECTEDITEM - 1
        ii = i
        Do While ListJoin(ii + 1)
            ii = ii + 1
        Loop
        mEditFlag1 = CBool(PropAtColumnNum(i, mCurTab, "EDIT"))
        If Not NoCaretShow And Not mEditFlag1 And i <> ii Then
            If ii - topitem - 1 <= lines + onemore And (ii > topitem - 1 Or ii <= TopRows) And Not ListSep(ii - 1) Then
                If i < topitem And ii >= topitem Then
                
                i = ii - topitem
                ii = SELECTEDITEM - topitem - 1
                Else
                i = ii - i
                ii = 0
                End If
                GoTo there1
            End If
        End If
        i = ii - i
        ii = 0
        If SELECTEDITEM - topitem - 1 <= lines + onemore And (SELECTEDITEM > topitem Or SELECTEDITEM <= TopRows) And Not ListSep(SELECTEDITEM - 1) Then
        mEditFlag1 = CBool(PropAtColumnNum(SELECTEDITEM - 1, mCurTab, "EDIT"))
        mcenter = CBool(PropAtColumnNum(SELECTEDITEM - 1, mCurTab, "CenterText"))
        mwraptext = CBool(PropAtColumnNum(SELECTEDITEM - 1, mCurTab, "WrapText"))
        mVcenter = CBool(PropAtColumnNum(SELECTEDITEM - 1, mCurTab, "VerticalCenterText"))
        If Not NoCaretShow Then
             If mEditFlag1 Then
                If SelStart = 0 Then SelStart = 1
                DrawStyle = vbSolid
                RaiseEvent PureListOff
                If mcenter Then
                    If mTabs = 1 Then
                        REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + (UserControl.ScaleWidth - UserControlTextWidth(list$(SELECTEDITEM - 1))) / 2 + LeftMarginPixels * scrTwips
                        RealX2 = scrollme / 2 + REALX
                    Else
                        REALX = UserControlTextWidth(Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), 1, SelStart - 1)) + (mParts(mCurTab) * scrTwips - UserControlTextWidth(listAtColumn(SELECTEDITEM - 1, mCurTab))) / 2 + TwipsCurTab + LeftMarginPixels * scrTwips
                        RealX2 = scrollme / 2 + REALX
                    End If
                Else
                    If mTabs = 1 Then
                        skipme = False
                        RaiseEvent GetRealX1(UserControl.hDC, SelStart, list(SELECTEDITEM - 1), REALX, skipme)
                        If skipme Then
                            skipme = False
                            REALX = (REALX + LeftMarginPixels) * scrTwips
                        Else
                            REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
                        End If
                        RealX2 = scrollme + REALX
                    Else
                        skipme = False
                        RaiseEvent GetRealX1(UserControl.hDC, SelStart, listAtColumn(SELECTEDITEM - 1, mCurTab), REALX, skipme)
                        If skipme Then
                            skipme = False
                            REALX = (REALX + LeftMarginPixels) * scrTwips + TwipsCurTab
                        Else
                            REALX = UserControlTextWidth(Mid$(listAtColumn(SELECTEDITEM - 1, mCurTab), 1, SelStart - 1)) + LeftMarginPixels * scrTwips + TwipsCurTab
                        End If
                        RealX2 = scrollme + REALX
                    End If
                End If
                    RaiseEvent PureListOn
                    If Noflashingcaret Or Not havefocus Then
                        If InternalCursor Or dragfocus Then
                            Line (scrollme + REALX, (SELECTEDITEM + TopRows - topitem - 1) * myt + myt1 + mHeadlineHeightTwips)-(RealX2, (SELECTEDITEM + TopRows - topitem - 1) * myt + mHeadlineHeightTwips), ForeColor
                        End If
                    Else
                        If InternalCursor Or Not MultiLineEditBox Then
                            If mTabs > 1 Then
                                If RealX2 < TwipsCurTab Then
                                ElseIf RealX2 > TwipsCurTab + mParts(mCurTab) * scrTwips Then
                                Else
                                    ShowMyCaretInTwips RealX2, (SELECTEDITEM + TopRows - topitem - 1 + tmprows) * myt + mHeadlineHeightTwips
                                End If
                            Else
                                ShowMyCaretInTwips RealX2, (SELECTEDITEM + TopRows - topitem - 1 + tmprows) * myt + mHeadlineHeightTwips
                                End If
                        End If
                    End If
                    If mTabs = 1 Then
                        If mEditFlag1 Or Not NoScroll Then
                            If RealX2 > Width * 0.8 * dragslow Then scrollme = scrollme - Width * 0.2 * dragslow: PrepareToShow 5
                        End If
                        If RealX2 - Width * 0.2 * dragslow < 0 Then
                            If mEditFlag1 Or Not NoScroll Then
                                scrollme = scrollme + Width * 0.2 * dragslow
                                If Not mcenter Then
                                    If scrollme > 0 Then scrollme = 0 Else PrepareToShow 5
                                Else
                                   PrepareToShow 5
                                End If
                            End If
                        End If
                    Else
                
                        If mEditFlag1 Or Not NoScroll Then
                            If RealX2 > (mParts(mCurTab) * scrTwips * 0.8 + TwipsCurTab) * dragslow Then
                                scrollme = scrollme - mParts(mCurTab) * scrTwips * 0.2 * dragslow: PrepareToShow 5
                            End If
                        End If
                        If RealX2 - (mParts(mCurTab) * scrTwips * 0.2 + TwipsCurTab) * dragslow < 0 Then
                            If mEditFlag1 Or Not NoScroll Then
                                If Not mcenter Then
                                    scrollme = scrollme + mParts(mCurTab) * scrTwips * 0.2 * dragslow
                                    If scrollme > 0 Then scrollme = 0 Else PrepareToShow 5
                                Else
                                    scrollme = scrollme + mParts(mCurTab) * scrTwips * 0.2 * dragslow
                                    PrepareToShow 5
                                End If
                            End If
                        End If
                    End If
                Else
there1:
                    DrawStyle = vbInvisible
                    If mTabs = 1 Or Not NarrowSelect Then
                        If BackStyle = 1 Then
                            Line (scrTwips, (SELECTEDITEM - topitem + TopRows + tmprows + i) * myt + mHeadlineHeightTwips)-(0 + UserControl.Width, (SELECTEDITEM - topitem - 1 + TopRows + tmprows) * myt + mHeadlineHeightTwips - scrTwips / 2), 0, B
                        Else
                            Line (0, (SELECTEDITEM - topitem + TopRows + tmprows + i) * myt + mHeadlineHeightTwips)-(0 + UserControl.Width, (SELECTEDITEM - topitem + TopRows - 1 + tmprows) * myt + mHeadlineHeightTwips), 0, B
                        End If
                    Else
                        If BackStyle = 1 Then
                            Line (scrTwips + TwipsCurTab, (SELECTEDITEM - ii - topitem + TopRows + tmprows + i) * myt + mHeadlineHeightTwips)-(TwipsCurTab + mParts(mCurTab) * scrTwips - 2 * scrTwips, (SELECTEDITEM - ii - topitem - 1 + TopRows + tmprows) * myt + mHeadlineHeightTwips), 0, B
                        Else
                            Line (TwipsCurTab, (SELECTEDITEM - ii - topitem + TopRows + tmprows + i) * myt + mHeadlineHeightTwips)-(TwipsCurTab + mParts(mCurTab) * scrTwips - scrTwips, (SELECTEDITEM - ii - topitem + TopRows - 1 + tmprows) * myt + mHeadlineHeightTwips), 0, B
                        End If
                    End If
                End If
            Else
                RaiseEvent PureListOff
            End If
        Else
            HideCaret (hWnd)
        End If
    End If
    DrawStyle = vbSolid
    DrawMode = vbCopyPen
    If Grid Then
    If mHeadlineHeightTwips > 0 Then
        Line (0, mHeadlineHeightTwips)-(UserControl.Width, mHeadlineHeightTwips), GridColor
    End If
    For i = 1 To lines + 1 + TopRows
        If i <= TopRows Then ii = i Else ii = i + topitem - TopRows
        If Not ListJoin(ii) Then
            Line (0, i * myt + mHeadlineHeightTwips)-(UserControl.Width, i * myt + mHeadlineHeightTwips), GridColor
        End If
    Next
    'If Not ListJoin(ii + 1) Then
    '    Line (sum, mHeadlineHeightTwips)-(sum, UserControl.Height), GridColor
    'End If
    sum = -scrTwips
    Line (sum, mHeadlineHeightTwips)-(sum, UserControl.Height), GridColor
    If mTabs > 1 Then
    For i = 1 To mTabs
        sum = sum + mParts(i) * scrTwips
        Line (sum, mHeadlineHeightTwips)-(sum, UserControl.Height), GridColor
    Next
    End If
    End If
    currentY = 0
    currentX = 0
End If
            If BorderStyle = 1 And BackStyle = 1 Then
                hnr.Left = 0
                hnr.Right = Me.WidthPixels
                hnr.top = 0
                hnr.Bottom = Me.HeightPixels
                onemore = CreateSolidBrush(Me.BackColor)
                FrameRect UserControl.hDC, hnr, onemore
                DeleteObject onemore
            End If

RepaintScrollBar
there:
End Sub

Property Get lines() As Long
Dim l As Long
On Error GoTo ex1
 myt = UserControlTextHeight() + addpixels * scrTwips
If restrictLines > 0 Then
l = restrictLines - 1
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips - 1) / restrictLines

Else
l = Int((UserControl.ScaleHeight - mHeadlineHeightTwips) / myt) - 1
End If
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
ex1:
If l <= 0 Then
l = 0
End If
If TopRows < l Then
    lines = l - TopRows
Else
    lines = 1
End If
End Property


Private Sub LargeBar1_Change()

If Not state Then



    topitem = Value
  
RaiseEvent ScrollMove(topitem)
Timer1.enabled = True
HandleOverride = True
LastVScroll = Value

End If
End Sub
Public Function TextHeightOffset() As Variant
If restrictLines = 0 Then
TextHeightOffset = 0
Else
TextHeightOffset = (myt - UserControlTextHeight()) \ scrTwips \ 2 - 1 ' + addpixels \ 2 + 1
End If
End Function
Public Sub RepaintOld7_18()
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) \ restrictLines
Else
myt = UserControlTextHeight() + addpixels * scrTwips
End If
'HeadlineHeight = UserControlTextHeight() / SCRTWIPS
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
On Error GoTo th1

If Extender.Parent Is Nothing Then Exit Sub

If Extender.Parent.Picture.Handle <> 0 And BackStyle = 1 Then

If Me.BorderStyle = 1 Then
currentY = 0
    currentX = 0
Line (0, 0)-(ScaleWidth - scrTwips, ScaleHeight - scrTwips), Me.BackColor, B
UserControl.PaintPicture UserControl.Parent.Picture, scrTwips, scrTwips, Width - 2 * scrTwips, Height - 2 * scrTwips, Extender.Left, Extender.top, Width - 2 * scrTwips, Height - 2 * scrTwips
    currentY = 0
    currentX = 0
Else
UserControl.PaintPicture UserControl.Parent.Picture, 0, 0, , , Extender.Left, Extender.top

End If

ElseIf BackStyle = 1 Then
Dim mmo As PictureBox
RaiseEvent GetBackPicture(mmo)
If Not mmo Is Nothing Then
If mmo.Picture.Handle <> 0 Then
    UserControl.PaintPicture mmo.Picture, 0, 0, , , Extender.Left, Extender.top
    If Me.BorderStyle = 1 Then
    currentY = 0
        currentX = 0
    Line (0, 0)-(ScaleWidth - scrTwips, ScaleHeight - scrTwips), Me.BackColor, B
        currentY = 0
        currentX = 0
    End If
End If
End If
Else
th1:
UserControl.Cls
End If
End Sub
Public Sub Repaint()
Dim hnr As RECT, br As Long, pp As Object
If restrictLines > 0 Then
myt = (UserControl.ScaleHeight - mHeadlineHeightTwips) \ restrictLines
Else
myt = UserControlTextHeight() + addpixels * scrTwips
End If
'HeadlineHeight = UserControlTextHeight() / SCRTWIPS
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
On Error GoTo th1
If Not waitforparent Then Exit Sub
If Extender.Parent Is Nothing Then Exit Sub

If BackStyle = 1 Then
    
    If Not SkipForm Then
        On Error Resume Next
        If Not UserControl.Extender.Container Is UserControl.Parent Then
            Set pp = UserControl.Extender.Container
        End If
        Err.Clear
        If pp Is Nothing Then
            Set pp = UserControl.Parent
        End If
        On Error GoTo th1
        If pp.Picture.Handle <> 0 Then
            If Me.BorderStyle = 1 Then
                UserControl.PaintPicture pp.Picture, scrTwips, scrTwips, Width - 2 * scrTwips, Height - 2 * scrTwips, Extender.Left, Extender.top, Width - 2 * scrTwips, Height - 2 * scrTwips
                hnr.Left = 0
                hnr.Right = Me.WidthPixels
                hnr.top = 0
                hnr.Bottom = Me.HeightPixels
                br = CreateSolidBrush(Me.BackColor)
                FrameRect UserControl.hDC, hnr, br
                DeleteObject br
            Else
                UserControl.PaintPicture pp.Picture, 0, 0, , , Extender.Left, Extender.top
            End If
        Else
            If Me.BorderStyle = 1 Then
                UserControl.PaintPicture pp.Image, 0, 0, Width - scrTwips, Height - scrTwips, Extender.Left, Extender.top, Width - scrTwips, Height - scrTwips
                hnr.Left = 0
                hnr.Right = Me.WidthPixels
                hnr.top = 0
                hnr.Bottom = Me.HeightPixels
                br = CreateSolidBrush(Me.BackColor)
                FrameRect UserControl.hDC, hnr, br
                DeleteObject br
                
            Else
                UserControl.PaintPicture pp.Image, 0, 0, , , Extender.Left, Extender.top
            End If
        End If
    Else
        Dim mmo As Object  ' MUST BE A FORM OR A PICTURE BOX
        RaiseEvent GetBackPicture(mmo)
        If Not mmo Is Nothing Then
            If mmo.Image.Handle <> 0 Then
                UserControl.PaintPicture mmo.Image, 0, 0, , , Extender.Left - mmo.Left, Extender.top - mmo.top
                If Me.BorderStyle = 1 Then
                        UserControl.PaintPicture UserControl.Parent.Picture, scrTwips, scrTwips, Width - 2 * scrTwips, Height - 2 * scrTwips, Extender.Left, Extender.top, Width - 2 * scrTwips, Height - 2 * scrTwips
                        hnr.Left = 0
                        hnr.Right = Me.WidthPixels
                        hnr.top = 0
                        hnr.Bottom = Me.HeightPixels
                        br = CreateSolidBrush(Me.BackColor)
                        FrameRect UserControl.hDC, hnr, br
                        DeleteObject br
                End If
            End If
        End If
    End If
Else
th1:
UserControl.Cls
End If
End Sub
Private Function GetStrUntilB(Pos As Long, ByVal sStr As String, fromstr As String, Optional RemoveSstr As Boolean = True) As String
Dim i As Long
If fromstr = vbNullString Then GetStrUntilB = vbNullString: Exit Function
If Pos <= 0 Then Pos = 1
If Pos > Len(fromstr) Then
    GetStrUntilB = vbNullString
Exit Function
End If
i = InStr(Pos, fromstr, sStr)
If (i < 1 + Pos) And Not ((i > 0) And RemoveSstr) Then
    GetStrUntilB = vbNullString
    Pos = Len(fromstr) + 1
Else
    GetStrUntilB = Mid$(fromstr, Pos, i - Pos)
    If RemoveSstr Then
        Pos = i + Len(sStr)
    Else
        Pos = i
    End If
End If
End Function
Function design() As Boolean
On Error GoTo there
If GetTopUserControl(Me).Ambient.UserMode = False Then
Cls
currentX = scrTwips
currentY = scrTwips
Print UserControl.Ambient.DisplayName
currentX = 0
currentY = 0
design = True
Else
'Cls
End If
Exit Function
there:
'If listcount = 0 Then Cls

End Function
Private Sub LargeBar1_Scroll()
If Not state Then
 topitem = Value
RaiseEvent ScrollMove(topitem)
Timer1.enabled = True
LastVScroll = Value
End If
End Sub
Public Function UserControlTextWidthPixels(a$) As Long
Dim nr As RECT
If Len(a$) > 0 Then

CalcRect UserControl.hDC, a$, nr
UserControlTextWidthPixels = nr.Right
End If
End Function
Public Sub UserControlTextMetricsPixels(a$, tw As Long, th As Long)
Dim nr As RECT
If Len(a$) > 0 Then

CalcRect UserControl.hDC, a$, nr
tw = nr.Right
th = nr.Bottom
End If
End Sub
Public Function UserControlTextWidth(a$) As Long
Dim nr As RECT
CalcRect UserControl.hDC, a$, nr
UserControlTextWidth = nr.Right * scrTwips
End Function
Public Function UserControlTextWidth2(a$, ByVal n As Long) As Long
Dim nr As RECT
n = n - 1
If n > 0 Then
CalcRect UserControl.hDC, Left$(a$, n), nr
UserControlTextWidth2 = nr.Right * scrTwips
End If
End Function
Private Function UserControlTextHeight() As Long
Dim nr As RECT
If overrideTextHeight = 0 Then
CalcRect1 UserControl.hDC, "fj", nr
UserControlTextHeight = nr.Bottom * scrTwips
Exit Function
End If
UserControlTextHeight = overrideTextHeight

End Function
Private Function UserControlTextHeightPixels() As Long
Dim nr As RECT
If overrideTextHeight = 0 Then
CalcRect1 UserControl.hDC, "fj", nr
UserControlTextHeightPixels = nr.Bottom
Exit Function
End If
UserControlTextHeightPixels = overrideTextHeight / scrTwips

End Function
Private Sub PrintLineControlSinglePrivate(mHdc As Long, ByVal c As String, R As RECT, WrapText1 As Boolean, CenterText1 As Boolean, VerticalCenterText1 As Boolean)
' this is our basic print routine
Dim that As Long, cc As String, fg As Long
If CenterText1 Then that = DT_CENTER
If VerticalCenterText1 Then that = that Or DT_VCENTER
If WrapText1 Then
    'c = c + space(4)  ' 4 additional characters for DT_MODIFYSTRING
    'DrawTextEx mHdc, StrPtr(c), -1, R, DT_WORDBREAK Or DT_NOPREFIX Or DT_MODIFYSTRING Or DT_PATH_ELLIPSIS Or that, VarPtr(tParam)
    DrawTextEx mHdc, StrPtr(c), -1, R, DT_WORDBREAK Or DT_NOPREFIX Or that, VarPtr(tParam)
Else
    If LastLinePart <> "" Then
        If FadeLastLinePart > 0 Then
            cc = c + LastLinePart
            fg = Me.ForeColor
            Me.ForeColor = FadeLastLinePart
            DrawText mHdc, StrPtr(cc), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS
            Me.ForeColor = fg
            DrawTextEx mHdc, StrPtr(c), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS, VarPtr(tParam)
        Else
            cc = c + LastLinePart
            DrawTextEx mHdc, StrPtr(cc), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS, VarPtr(tParam)
        End If
    Else
        DrawTextEx mHdc, StrPtr(c), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
    End If
End If
    
End Sub
Private Sub PrintLineControlSingle(mHdc As Long, ByVal c As String, R As RECT)
' this is our basic print routine
Dim that As Long, cc As String, fg As Long
If CenterText Then that = DT_CENTER
If VerticalCenterText And Not WrapText Then that = that Or DT_VCENTER
If WrapText Then
'If Not CenterText Then
 '   c = c + space(4)  ' 4 additional characters for DT_MODIFYSTRING  '' Or DT_MODIFYSTRING Or DT_PATH_ELLIPSIS
  '  DrawTextEx mHdc, StrPtr(c), -1, r, DT_WORDBREAK Or DT_NOPREFIX Or that, VarPtr(tParam)
'Else
    DrawTextEx mHdc, StrPtr(c), -1, R, DT_WORDBREAK Or DT_NOPREFIX Or that, VarPtr(tParam)
'End If
Else
    If LastLinePart <> "" Then
        If FadeLastLinePart > 0 Then
            cc = c + LastLinePart
            fg = Me.ForeColor
            Me.ForeColor = FadeLastLinePart
            DrawText mHdc, StrPtr(cc), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS
            Me.ForeColor = fg
            DrawTextEx mHdc, StrPtr(c), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS, VarPtr(tParam)
        Else
            cc = c + LastLinePart
            DrawTextEx mHdc, StrPtr(cc), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS, VarPtr(tParam)
        End If
    Else
        DrawTextEx mHdc, StrPtr(c), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
    End If
End If
    
End Sub
Private Sub PrintLineControlSingleNoWrap(mHdc As Long, ByVal c As String, R As RECT)
' this is our basic print routine
Dim that As Long, cc As String, fg As Long
If CenterText Then that = DT_CENTER
If VerticalCenterText Then that = that Or DT_VCENTER
If LastLinePart <> "" Then
    If FadeLastLinePart > 0 Then
    cc = c + LastLinePart
    fg = Me.ForeColor
    Me.ForeColor = FadeLastLinePart
   DrawText mHdc, StrPtr(cc), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS
   Me.ForeColor = fg
   DrawTextEx mHdc, StrPtr(c), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS, VarPtr(tParam)
    
    Else
    cc = c + LastLinePart
   DrawTextEx mHdc, StrPtr(cc), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS, VarPtr(tParam)
   End If
Else

    DrawTextEx mHdc, StrPtr(c), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
    End If
    
End Sub

Private Sub PrintLineControlHeader(mHdc As Long, c As String, R As RECT, Optional that As Long = 0)
' this is our basic print routine
DrawTextEx mHdc, StrPtr(c), -1, R, DT_WORDBREAK Or DT_NOPREFIX Or that Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)

    
End Sub
  Private Sub CalcRectHeader(mHdc As Long, c As String, R As RECT, Optional that As Long = 0)
R.top = 0
R.Left = 0
If R.Right = 0 Then R.Right = UserControl.Width / scrTwips
DrawTextEx mHdc, StrPtr(c), -1, R, DT_CALCRECT Or DT_WORDBREAK Or DT_NOPREFIX Or DT_EXPANDTABS Or DT_TABSTOP Or that, VarPtr(tParam)
End Sub
Private Sub PrintLineControl(mHdc As Long, c As String, R As RECT)
    DrawTextEx mHdc, StrPtr(c), -1, R, DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS, VarPtr(tParam)

End Sub
Private Sub PrintLinePixels(dd As Object, c As String)
Dim R As RECT    ' print to a picturebox as label
R.Right = dd.ScaleWidth
R.Bottom = dd.ScaleHeight
DrawTextEx dd.hDC, StrPtr(c), -1, R, DT_NOPREFIX Or DT_WORDBREAK Or DT_EXPANDTABS, VarPtr(tParam)
End Sub
Private Sub CalcRect(mHdc As Long, c As String, R As RECT)
R.top = 0
R.Left = 0
Dim that As Long
If CenterText Then that = DT_CENTER
If VerticalCenterText Then that = that Or DT_VCENTER
If WrapText Then
    If R.Right = 0 Then R.Right = UserControl.Width / scrTwips
    DrawTextEx mHdc, StrPtr(c), -1, R, DT_CALCRECT Or DT_WORDBREAK Or DT_NOPREFIX Or that Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
Else
    DrawTextEx mHdc, StrPtr(c), -1, R, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
End If

End Sub
Private Sub CalcRect1(mHdc As Long, c As String, R As RECT)
R.top = 0
R.Left = 0

If WrapText Then
    If R.Right = 0 Then R.Right = UserControl.Width / scrTwips - LeftMarginPixels
    DrawTextEx mHdc, StrPtr(c), -1, R, DT_CALCRECT Or DT_WORDBREAK Or DT_NOPREFIX Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
Else
    DrawTextEx mHdc, StrPtr(c), -1, R, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
End If

End Sub
Public Function SpellUnicode(a$)
' use spellunicode to get numbers
' and make a ListenUnicode...with numbers for input text
Dim b$, i As Long
For i = 1 To Len(a$) - 1
b$ = b$ & CStr(AscW(Mid$(a$, i, 1))) & ","
Next i
SpellUnicode = b$ & CStr(AscW(Right$(a$, 1)))
End Function
Public Function ListenUnicode(ParamArray aa() As Variant) As String
Dim all$, i As Long
For i = 0 To UBound(aa)
    all$ = all$ & ChrW(aa(i))
Next i
ListenUnicode = all$
End Function

Public Sub RepaintFromOut(parentpic As StdPicture, Myleft As Long, mytop As Long)
On Error GoTo th1

If parentpic.Handle <> 0 Then
UserControl.PaintPicture parentpic, 0, 0, , , Myleft, mytop
Else
th1:
'UserControl.Cls
End If
End Sub
Private Sub Redraw(ParamArray status())
If EnabledBar Then
Dim fakeLargeChange As Long, NewHeight As Long, newTop As Long
Dim b As Boolean, nstatus As Boolean
Timer2bar.enabled = False
If UBound(status) >= 0 Then
nstatus = CBool(status(0))
Else
nstatus = Shape1.Visible
End If
With UserControl
If mHeadline <> "" Then
NewHeight = .Height - mHeadlineHeightTwips
newTop = mHeadlineHeightTwips
Else
NewHeight = .Height
End If

If NewHeight <= 0 Then

Else
        minimumWidth = (1 - (Max - Min) / (largechange + Max - Min)) * NewHeight * (1 - Percent * 2) + 1
        If minimumWidth < 60 Then
        
        mLargeChange = Round(-(Max - Min) / ((60 - 1) / NewHeight / (1 - Percent * 2) - 1) - Max + Min) + 1
        
        minimumWidth = (1 - (Max - Min) / (largechange + Max - Min)) * NewHeight * (1 - Percent * 2) + 1
        End If
        valuepoint = (Value - Min) / (largechange + Max - Min) * (NewHeight * (1 - 2 * Percent)) + NewHeight * Percent

       Shape Shape1, Width - BarWidth, newTop + valuepoint, BarWidth, minimumWidth
       Shape Shape2, Width - BarWidth, newTop + NewHeight * (1 - Percent), BarWidth, NewHeight * Percent ' newtop + newheight * Percent - scrTwips
        Shape Shape3, Width - BarWidth, newTop, BarWidth, NewHeight * Percent   ' left or top
End If
End With
If UBound(status) >= 0 Then
b = (CBool(status(0)) Or Spinner) And listcount > lines
If Not Shape1.Visible = b Then
Shape1.Visible = b
Shape2.Visible = b
Shape3.Visible = b

End If
End If

End If
End Sub
Private Property Get largechange() As Long
If mLargeChange < 1 Then mLargeChange = 1
largechange = mLargeChange
End Property

Private Property Let largechange(ByVal RHS As Long)
If RHS < 1 Then RHS = 1
mLargeChange = RHS
showshapes
PropertyChanged "LargeChange"
End Property
Public Property Get smallchange() As Long
smallchange = mSmallChange
End Property

Private Property Let smallchange(ByVal RHS As Long)
If RHS < 1 Then RHS = 1
mSmallChange = RHS
showshapes
PropertyChanged "SmallChange"
End Property
Private Property Get Max() As Long
Max = mmax
End Property

Private Property Let Max(ByVal RHS As Long)
If Min > RHS Then RHS = Min
If mValue > RHS Then mValue = RHS  ' change but not send event
If RHS = 0 Then RHS = 1
mmax = RHS
showshapes
PropertyChanged "Max"
End Property

Private Property Get Min() As Long
Min = mmin + TopRows
End Property
Public Sub SetSpin(Low As Long, high As Long, stepbig As Long)
If Spinner Then
mpercent = 0.33
mmax = high
mmin = Low
mLargeChange = (Max - Min) * 0.2
mSmallChange = stepbig
mjumptothemousemode = True
End If
End Sub

Private Property Let Min(ByVal RHS As Long)
If Max <= RHS Then RHS = Max
If mValue < RHS Then mValue = RHS  ' change but not send event

mmin = RHS
showshapes
PropertyChanged "LargeChange"
PropertyChanged "Min"
End Property
Public Property Get EnabledBar() As Boolean
If InfoDropBarClick Then Exit Property
EnabledBar = Not NoFire
End Property

Public Property Let EnabledBar(ByVal RHS As Boolean)
If Not myEnabled Then Exit Property
NoFire = Not EnabledBar
Shape1.Visible = Not NoFire
Shape2.Visible = Not NoFire
Shape3.Visible = Not NoFire
Shape Shape1
Shape Shape2
Shape Shape3
If Not NoFire = True Then Timer1.enabled = True
End Property
Public Property Get Value() As Long
Value = mValue
End Property
Public Property Get Visible() As Boolean
On Error Resume Next
If UserControl.Parent Is Nothing Then Exit Property
Visible = Extender.Visible
End Property

Public Property Get TopTwips() As Long
On Error Resume Next
TopTwips = CLng(Extender.top)
End Property
Public Property Let Visible(ByVal RHS As Boolean)
On Error Resume Next
Extender.Visible = RHS
End Property
Public Property Let TopTwips(ByVal RHS As Long)
On Error Resume Next
Extender.move Extender.Left, CSng(RHS)
End Property
Public Property Get HeightTwips() As Long
On Error Resume Next
HeightTwips = CLng(Extender.Height)
End Property
Public Sub GetLeftTop(Ltwips, Ttwips)
On Error Resume Next
Ltwips = CLng(Extender.Left)
Ttwips = CLng(Extender.top)
End Sub
Public Property Let HeightTwips(ByVal RHS As Long)
On Error Resume Next
Extender.move Extender.Left, Extender.top, Extender.Width, RHS
End Property
Public Sub MoveTwips(ByVal mleft As Long, ByVal mtop As Long, mWidth As Long, mHeight As Long)
On Error Resume Next
If UseHeaderOnly Then
Extender.move mleft, mtop, mWidth, mHeight
ElseIf mWidth < 100 Then
Extender.move mleft, mtop, Extender.Width, Extender.Height
ElseIf mHeight < 100 Then
Extender.move mleft, mtop, mWidth, Extender.Height
Else
Extender.move mleft, mtop, mWidth, mHeight
End If
End Sub
Public Sub ZOrder(Optional ByVal RHS As Long = 0)
On Error Resume Next
Extender.ZOrder RHS
End Sub

Public Sub SetFocus()
On Error Resume Next
If Extender.Visible Then
Extender.SetFocus
End If
End Sub
Public Property Let Value(ByVal RHS As Long)
' Dim oldvalue As Long
If RHS < Min Then RHS = Min
If RHS > Max Then RHS = Max
If state And Spinner Then
'don't fix the value
Else
mValue = RHS
End If
showshapes

If Not Spinner Then
If Not NoFire Then
LargeBar1_Change
End If
Else

RaiseEvent SpinnerValue(mmax - mValue + mmin)
Redraw

'UserControl.refresh
End If
PropertyChanged "Value"
End Property
Public Property Let ValueSilent(ByVal RHS As Long)
If Spinner Then
' no events
If RHS < Min Then RHS = Min
If RHS > Max Then RHS = Max
mValue = Max - RHS + Min
showshapes
End If
End Property
Public Property Get ValueSilent() As Long
ValueSilent = Max - mValue + Min
End Property
Private Property Get BarVisible() As Boolean
BarVisible = Shape1.Visible
End Property
Private Property Let BarVisible(ByVal RHS As Boolean)
If Not myEnabled Then
Exit Property
End If
If RHS = False And Shape1.Visible = False Then
If nopointerchange Then
UserControl.MousePointer = oldpointer
End If
Else
If listcount = 0 Then RHS = False
Shape1.Visible = RHS Or Spinner
Shape2.Visible = RHS Or Spinner
Shape3.Visible = RHS Or Spinner
Shape Shape1
Shape Shape2
Shape Shape3
If nopointerchange Then
If (RHS Or Spinner) Then
UserControl.MousePointer = 1
Else
UserControl.MousePointer = oldpointer
End If
End If
If Not NoFire = True Then Timer1.enabled = True
End If
End Property

Private Sub showshapes()
If m_showbar Or StickBar Or Spinner Or AutoHide Then
Timer2bar.enabled = True
End If
End Sub
Public Property Get Percent() As Single
Percent = mpercent
End Property

Public Property Let Percent(ByVal RHS As Single)
mpercent = RHS
PropertyChanged "Percent"
End Property
Friend Sub Goback()
    If Typename(Me.Parent) = "GuiM2000" Then
        ChooseNextLeft Me, Me.Parent, True
    End If
End Sub
Friend Sub GoON()
    If Typename(Me.Parent) = "GuiM2000" Then
        ChooseNextRight Me, Me.Parent, True
    End If
End Sub
Friend Sub TakeKey(KeyCode As Integer, shift As Integer)
    UserControl_KeyDown KeyCode, shift
If KeyCode <> 0 Then
    UserControl_KeyUp KeyCode, shift
End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, shift As Integer)
On Error GoTo fin
If shift = 0 Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        If WinKey() Then
            RaiseEvent WindowKey(KeyCode)
            KeyCode = 0
            Exit Sub
        End If
    End If
End If
PrevLocale = GetLocale()
If BypassKey Then KeyCode = 0: shift = 0: Exit Sub
lastshift = shift

If KeyCode = 27 And NoEscapeKey Then
KeyCode = 0
Exit Sub
End If
If Arrows2Tab And Not EditFlagSpecial Then
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
        If Typename(Me.Parent) = "GuiM2000" Then
        ChooseNextLeft Me, Me.Parent
        End If
        KeyCode = 0
        Exit Sub
    ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
        If Typename(Me.Parent) = "GuiM2000" Then
        ChooseNextRight Me, Me.Parent
        End If
        KeyCode = 0
        Exit Sub
    End If
End If
If KeyCode = vbKeyTab And Not EditFlagSpecial Then
    If shift = 1 Then
        choosenext
        KeyCode = 0
        Exit Sub
    End If
ElseIf KeyCode = vbKeyF4 Then
If shift = 4 Then
On Error Resume Next
If Parent.Name = "GuiM2000" Or Parent.Name = "Form2" Or Parent.Name = "Form4" Then
With UserControl.Parent
.ByeBye
End With
KeyCode = 0
Exit Sub
End If
End If
End If
If DropKey Then shift = 0: KeyCode = 0: Exit Sub
Dim i&
If shift = 4 Then
If KeyCode = 18 Then
If mynum$ = vbNullString Then mynum$ = "0"
KeyCode = 0
Exit Sub
Else
If KeyCode <> 0 Then GetKeY2 KeyCode, shift
End If
Select Case KeyCode
Case vbKeyAdd, vbKeyInsert
mynum$ = "&h"
Case vbKey0 To vbKey9
mynum$ = mynum$ + Chr$(KeyCode - vbKey0 + 48)
LastNumX = True
Case vbKeyNumpad0 To vbKeyNumpad9
LastNumX = False
mynum$ = mynum$ + Chr$(KeyCode - vbKeyNumpad0 + 48)
Case vbKeyA To vbKeyF
If Left$(mynum$, 1) = "&" Then
mynum$ = mynum$ + Chr$(KeyCode - vbKeyNumpad0 + 65)
LastNumX = True
Else
mynum$ = vbNullString
End If
Case Else
mynum$ = vbNullString
End Select
Exit Sub
End If

mynum$ = vbNullString
If shift <> 0 And KeyCode = 0 Then Exit Sub

RaiseEvent KeyDown(KeyCode, shift)
If KeyCode <> 0 Then GetKeY2 KeyCode, shift
If (KeyCode = 0) Or Not (enabled Or state) Then Exit Sub
If SELECTEDITEM < 0 Then
SELECTEDITEM = topitem + 1: ShowMe2
If Not EditFlagSpecial Then: KeyCode = 0
End If
LargeBar1KeyDown KeyCode, shift
If EnabledBar Then
Select Case KeyCode
Case vbKeyLeft, vbKeyUp
If WinKey Then

End If
If Spinner Then
If Not NoBarClick Then
    If shift Then
        Value = Value - 1
    Else
        Value = Value - mSmallChange
    End If
    End If
Else
Value = Value - mSmallChange
End If
Case vbKeyPageUp
Value = Value - largechange
Case vbKeyRight, vbKeyDown
If Spinner Then
    If Not NoBarClick Then
        If shift Then
        Value = Value + 1
        Else
        Value = Value + mSmallChange
        End If
    End If
Else
If Value + largechange + 1 <= Max Then
Value = Value + mSmallChange
End If
End If
Case vbKeyPageDown
Value = Value + largechange
End Select
End If

i = GetLastKeyPressed
 If i <> -1 And i <> 94 Then
  If i = 13 Then
 UKEY$ = vbNullString
 Else
 UKEY$ = ChrW(i)
 End If
 End If
fin:
End Sub
Public Property Get Vertical() As Boolean
Vertical = mVertical
End Property

Public Property Let Vertical(ByVal RHS As Boolean)
RHS = True ' intercept
mVertical = RHS
showshapes
PropertyChanged "Vertical"
End Property

Public Property Get jumptothemousemode() As Boolean
jumptothemousemode = mjumptothemousemode
End Property

Public Property Let jumptothemousemode(ByVal RHS As Boolean)
mjumptothemousemode = RHS
End Property
Private Function processXY(ByVal X As Single, ByVal y As Single, Optional rep As Boolean = True) As Boolean
If NoBarClick Or UseHeaderOnly Then Exit Function
Timer1bar.enabled = False
Dim checknewvalue As Long, NewHeight As Long
With UserControl
If mHeadline <> "" Then
NewHeight = .Height - mHeadlineHeightTwips
y = y - mHeadlineHeightTwips
Else
NewHeight = .Height
End If

If minimumWidth < 60 Then minimumWidth = 60  ' 4 x scrtwips
' value must have real max ...minimum MAX-60
If Vertical Then
' here minimumwidth is minimumheight
If y >= valuepoint - scrTwips And y <= minimumWidth + valuepoint - scrTwips Then
' is our scroll bar
OurDraw = Not rep

ElseIf y > NewHeight * Percent And y < NewHeight * (1 - Percent) Then
'  we are inside so take a largechange
processXY = True

        If y < valuepoint Then
         ' jump to mouse position at page (or fakepage )
                    If mjumptothemousemode Then
                     y = (y \ minimumWidth + 1) * minimumWidth - minimumWidth
                     Else
                    y = valuepoint - minimumWidth
                    End If
        Else
         ' jump to mouse position at page (or fakepage )
                If mjumptothemousemode Then
                y = (y \ minimumWidth - 1) * minimumWidth + minimumWidth
                Else
                y = valuepoint + minimumWidth
                End If
        End If
            If y < NewHeight * Percent Then y = NewHeight * Percent
            If y > Round(NewHeight * (1 - Percent)) - minimumWidth + NewHeight * Percent Then
            y = NewHeight * (1 - Percent) - minimumWidth
            End If
            checknewvalue = Round((y - NewHeight * Percent) * (Max - Min) / ((NewHeight * (1 - Percent) - minimumWidth) - NewHeight * Percent)) + Min
            If checknewvalue = Value And mjumptothemousemode Then
                 ' do nothing
                
            Else
    
                Value = checknewvalue
                If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5 ' autorepeat
                Timer1bar.enabled = True
            End If
ElseIf y >= NewHeight * (1 - Percent) And y <= NewHeight Then ' is right button
processXY = True
checknewvalue = Value + mSmallChange
If checknewvalue = Value Then
' do nothing
Else
Value = checknewvalue
If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5  '
Timer1bar.enabled = True
End If

ElseIf y < NewHeight * Percent - scrTwips Then
processXY = True
checknewvalue = Value - mSmallChange
' is  left button
If checknewvalue = Value Then
' do nothing
Else
Value = checknewvalue

If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5 ' autorepeat
Timer1bar.enabled = True
End If
End If

ElseIf Not Timer1bar.enabled Then
If X >= valuepoint - scrTwips And X <= minimumWidth + valuepoint - scrTwips Then
' is our scroll bar
OurDraw = Not rep

ElseIf X > .Width * Percent And X < .Width * (1 - Percent) Then
processXY = True
'  we are inside so take a largechange
        If X < valuepoint Then
                If mjumptothemousemode Then
                  X = (X \ minimumWidth + 1) * minimumWidth - minimumWidth
                Else
                X = valuepoint - minimumWidth
                End If
        Else
                If mjumptothemousemode Then
                X = (X \ minimumWidth - 1) * minimumWidth + minimumWidth
                Else
                X = valuepoint + minimumWidth
                End If
        End If
            If X < .Width * Percent Then X = .Width * Percent
            If X > Round(.Width * (1 - Percent)) - minimumWidth + .Width * Percent Then
            X = .Width * (1 - Percent) - minimumWidth
            End If
            checknewvalue = Round((X - .Width * Percent) * (Max - Min) / ((.Width * (1 - Percent) - minimumWidth) - .Width * Percent)) + Min
            If checknewvalue = Value And mjumptothemousemode Then
            ' do nothing
            Else
            Value = checknewvalue
            If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5  ' autorepeat
            Timer1bar.enabled = True
            End If
ElseIf X >= .Width * (1 - Percent) And X <= .Width Then
processXY = True
checknewvalue = Value + mSmallChange
If checknewvalue = Value Then
' do nothing
Else
Value = checknewvalue
If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5 ' autorepeat
Timer1bar.enabled = True
End If
' is right button
ElseIf X < .Width * Percent - scrTwips Then
processXY = True
checknewvalue = Value - mSmallChange
If checknewvalue = Value Then
' do nothing
Else
Value = checknewvalue
If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5  ' autorepeat
Timer1bar.enabled = True
' is  left button
End If
End If

End If
End With
End Function
Private Sub barMouseMove(Button As Integer, shift As Integer, X As Single, ByVal y As Single)
If Not EnabledBar Then Exit Sub
Dim ForValidValue As Long, NewHeight As Long
If OurDraw Then
If Button = 1 Then
Timer1bar.Interval = 5000
'timer2bar.enabled = False
If minimumWidth < 60 Then minimumWidth = 60  ' 4 x scrtwips

With UserControl


If Vertical Then

If mHeadline <> "" Then
y = y - mHeadlineHeightTwips
NewHeight = .Height - mHeadlineHeightTwips
Else
NewHeight = .Height
End If
        ForValidValue = y + GetOpenValue 'ForValidValue + valuepoint
        If ForValidValue < NewHeight * Percent Then
        ForValidValue = NewHeight * Percent
        Value = Min
        ElseIf ForValidValue > ((NewHeight * (1 - Percent) - minimumWidth)) Then
        ForValidValue = ((NewHeight * (1 - Percent) - minimumWidth))
        Value = Max
        Else

         Value = Round((ForValidValue - NewHeight * Percent) * (Max - Min) / ((NewHeight * (1 - Percent) - minimumWidth) - NewHeight * Percent)) + Min
         
        End If
    

Else

         ForValidValue = X + GetOpenValue
        If ForValidValue < .Width * Percent Then
        ForValidValue = .Width * Percent
        Value = Min
        ElseIf ForValidValue > ((.Width * (1 - Percent) - minimumWidth)) Then
        ForValidValue = ((.Width * (1 - Percent) - minimumWidth))
        Value = Max
        Else
        Value = Round((ForValidValue - .Width * Percent) * (Max - Min) / ((.Width * (1 - Percent) - minimumWidth) - .Width * Percent)) + Min
        
        End If
      
End If
showshapes
'Redraw


End With
If Not Spinner Then
If Not NoFire Then LargeBar1_Scroll
Else

RaiseEvent SpinnerValue(mmax - mValue + mmin)
End If
End If
End If
End Sub
Public Sub MenuItem(ByVal item As Long, checked As Boolean, radiobutton As Boolean, firstState As Boolean, Optional id$)
' Using MenuItem we want glist to act as a menu with checked and radio buttons
item = item - 1  ' from 1...to listcount as input
' now from 0 to listcount-1
If itemcount > 0 And Not BlockItemcount Then
    If item >= 0 And item < listcount Then
        If LeftMarginPixels < mytPixels Then LeftMarginPixels = mytPixels
        'mList(item).checked = checked ' means that can be checked
        With mList(item)
            .Flags = .Flags Or fchecked
            If Not checked Then .Flags = .Flags Xor fchecked
            If .content Is Nothing Then Set .content = New JsonObject
              .content.AssignPath "C.0", CVar(id$)   ' now contendID is in 0
        End With
        ListSelected(item) = firstState
        With mList(item)
            .Flags = .Flags Or fradiobutton
            If Not radiobutton Then .Flags = .Flags Xor fradiobutton
        End With
    End If
End If
End Sub
Public Function GetMenuId(id$, Pos As Long) As Boolean
' return item number with that id$
' work only in the internal list
Dim i As Long
If itemcount > 0 And Not BlockItemcount Then
For i = 0 To itemcount - 1
If Not mList(i).content Is Nothing Then
If mList(i).content.ItemPath("C.0") = CVar(id$) Then Pos = i: Exit For
End If
Next i
End If
GetMenuId = Not (i = itemcount)
End Function
Property Get id(item As Long) As String
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 And item < listcount Then
If Not mList(item).content Is Nothing Then
id = mList(item).content.ItemPath("C.0")
End If
End If
End If
End Property
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub

Private Sub MyMark(thathDC As Long, Radius As Long, X As Long, y As Long, item As Long, Optional Reverse As Boolean = False) ' circle
'
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
Dim th As RECT
th.Left = X - Radius
th.top = y - Radius
th.Right = X + Radius
th.Bottom = y + Radius
Dim old_brush As Long, old_pen As Long, my_brush As Long

    If Not ListChecked(item) Then
                If Reverse Then
                  my_brush = CreateSolidBrush(0)
                Else
                   my_brush = CreateSolidBrush(m_ForeColor)
                End If
            FillRect thathDC, th, my_brush
            DeleteObject my_brush
             Radius = Radius - 2
             If Radius = 0 Then Radius = 1
        Else
        Radius = mytPixels / 5
        If Radius < 4 Then Radius = 4
        End If
             
        th.Left = X - Radius
        th.top = y - Radius
        th.Right = X + Radius
        th.Bottom = y + Radius

        

        If ListSelected(item) Then
            If Reverse Then
                my_brush = CreateSolidBrush(0)  '
            Else
                my_brush = CreateSolidBrush(m_dcolor)  'm_CapColor
            End If
        Else
        If Reverse Then
            my_brush = CreateSolidBrush(&HFFFFFF)
        Else
            my_brush = CreateSolidBrush(m_BackColor)
        End If
             End If
     FillRect thathDC, th, my_brush
DeleteObject my_brush
 



End Sub


Public Property Get widthtwips() As Long

widthtwips = UserControl.Width
End Property
Public Property Get WidthPixels() As Long
WidthPixels = UserControl.Width / scrTwips
End Property
Public Property Get HeightPixels() As Long
HeightPixels = UserControl.Height / scrTwips
End Property
Public Sub REALCUR(where As Long, ByVal probeX As Single, realpos As Long, usedCharLength As Long)
Dim n As Long, st As Long, st1 As Long, st0 As Long, W As Integer, mark1 As Long, mark2 As Long
Dim Pad$, addlength As Long, s$, tColumn As Long

If mTabs = 1 Then
    s$ = list(where)
    RaiseEvent RealCurReplace(s$)
    n = UserControlTextWidth(s$)
    mCurTab = 1
    TwipsCurTab = 0
    If CenterText Then
    probeX = scrollme / 2 + probeX - LeftMarginPixels * scrTwips - (UserControl.ScaleWidth - LeftMarginPixels * scrTwips - n) / 2 + 2 * scrTwips
    Else
    probeX = probeX - LeftMarginPixels * scrTwips + 2 * scrTwips
    End If
Else
    st = (CLng(probeX + scrollme)) \ scrTwips
    For W = 1 To mTabs
        If mParts(W) > st Then Exit For
        st = st - mParts(W)
    Next W
    tColumn = W
    If tColumn > mTabs Then tColumn = mTabs
    s$ = listAtColumn(where, tColumn)
    mCurTab = tColumn
    RaiseEvent RealCurReplace(s$)
    
    n = UserControlTextWidth(s$)
    
    st = 0
    For W = 1 To tColumn - 1
    st = st + mParts(W)
    Next W
    TwipsCurTab = st * scrTwips
    If PropAtColumnNum(where, mCurTab, "CenterText") Then
        probeX = scrollme / 2 + probeX - (mParts(tColumn) * scrTwips - n) / 2 + 2 * scrTwips - LeftMarginPixels * scrTwips - TwipsCurTab
    Else
        probeX = probeX - LeftMarginPixels * scrTwips + 2 * scrTwips - TwipsCurTab
    End If
st = 0
End If
If probeX > n Then
If s$ = vbNullString Then
realpos = 0
usedCharLength = 1
Else
addlength = -1
mark1 = probeX \ scrTwips
st1 = Len(s$) + 1
RaiseEvent rtl(UserControl.hDC, where, st1, mark1, mark2, addlength)
If addlength >= 0 Then


realpos = addlength * scrTwips


Else

realpos = n
End If
usedCharLength = st1 - 1

End If
Else
If probeX <= 30 Then
realpos = 0
usedCharLength = 0
Exit Sub
End If
st = Len(s$)
st1 = st + 1
st0 = 1
While st > st0 + 1
st1 = (st + st0) \ 2
W = AscW(Mid$(s$, 1, st1))
If W > -10241 And W < -9216 Then
If probeX >= UserControlTextWidth2(s$, st1 + 2) Then
st0 = st1
Else
st = st1
End If
Else
If probeX >= UserControlTextWidth2(s$, st1 + 1) Then
st0 = st1
Else
st = st1
End If
End If
Wend
If probeX > UserControlTextWidth2(s$, st0) Then
    st1 = st0
End If
addlength = -1
mark1 = probeX \ scrTwips
st1 = st1 + 1
RaiseEvent rtl(UserControl.hDC, where, st1, mark1, mark2, addlength)
If addlength >= 0 Then


realpos = addlength * scrTwips

usedCharLength = st1 - 1
Exit Sub
End If
' the old one
st1 = st1 - 1

If mark1 <> 0 And mark2 <> 0 Then

If st1 + 1 >= mark2 Then
st1 = mark1 - 1

realpos = UserControlTextWidth2(s$, st1)
RaiseEvent rtl(UserControl.hDC, where, st1, mark1, mark2, addlength)

Else
W = UserControlTextWidth2(s$, mark2)
Pad$ = Mid$(s$, mark1, mark2 - mark1 + 1)
n = Len(Pad$)
st1 = 0
Do
st1 = st1 + 1
realpos = W - UserControlTextWidth2(Pad$, st1 + 1)

Loop While realpos >= probeX And st1 < n
While st1 > 1 And st1 < mark2 And UserControlTextWidth2(Pad$, st1 + 1) - UserControlTextWidth2(Pad$, st1 + 2) > 0
st1 = st1 + 1
Wend
st1 = mark1 + st1 - 1
st0 = st1 + 2
RaiseEvent rtl(UserControl.hDC, where, st0, mark1, mark2, addlength)
If st0 <> st1 + 2 Then
st1 = st0 - 2
End If
End If

Else
If probeX > UserControlTextWidth2(s$, st1 + 1) Then

st1 = st1 + 1
Else
If st1 = 2 Then
If probeX < UserControlTextWidth2(s$, 2) Then st1 = 1
End If
End If
Do
st1 = st1 - 1
If st1 > 0 Then

realpos = UserControlTextWidth2(s$, st1 + 1)
Else
realpos = 0
End If

Loop While realpos > probeX And st1 > 1
End If
If realpos > probeX Then
usedCharLength = 0
Else
st1 = st1 + 1

If mark1 <= st1 And st1 <= mark2 + 1 And mark1 <> 0 And mark2 <> 0 Then
usedCharLength = st1
Else

usedCharLength = st1 - 1
End If



End If

End If
End Sub
Public Function Pixels2Twips(pixels As Long) As Long
Pixels2Twips = pixels * scrTwips
End Function
Public Function BreakLine(Data As String, datanext As String, Optional thatTwipsPreserveRight As Long = -1, Optional aSpace$ = " ") As Boolean
Dim i As Long, k As Long, m As Long
If thatTwipsPreserveRight = -1 Then
m = widthtwips
Else
m = widthtwips - thatTwipsPreserveRight
End If
''If aSpace$ <> "" Then m = m - UserControlTextWidth(aSpace$)
REALCURb Data, m, k, i, True
datanext = Mid$(Data, 1, i)
Data = Mid$(Data, i + 1)

' lets see if we have space in data
If Len(Data) > 0 Then
    If Right$(datanext, 1) <> aSpace$ And Left$(Data, 1) <> aSpace$ Then
    ' we have a broken word
    m = InStrRev(datanext, aSpace$)
    If m > 0 Then
    ' we have a space inside datanext
    If m > 1 Then
    Data = Mid$(datanext, m + 1) + Data
    datanext = Left$(datanext, m)
    Else
    ' do nothing, we will have nothing for this line if we take the word
    End If
    Else
    ' do nothing it is a big word...
    m = InStrRev(datanext, "\")
    If m > 1 Then
    Data = Mid$(datanext, m + 1) + Data
    datanext = Left$(datanext, m)
    Else
    ' do nothing, we will have nothing for this line if we take the word
    End If
    End If
    End If
    
    i = 1
    If Data <> aSpace$ Or Data$ = vbNullString Then
    While Left$(Data, i) = aSpace$
    i = i + 1
    Wend
    End If
    datanext = datanext + Mid$(Data, 1, i - 1)
    Data = Mid$(Data, i)
    
End If
BreakLine = Data <> ""
End Function
Public Sub REALCURb(ByVal s$, ByVal probeX As Single, realpos As Long, usedCharLength As Long, Optional notextonly As Boolean = False)
' this is for breakline only
Dim n As Long, st As Long, st1 As Long, st0 As Long

If Not notextonly Then probeX = probeX - UserControlTextWidth("W") ' Else' probeX = probeX + 2 * scrTwips

n = UserControlTextWidth(s$)

probeX = probeX - 2 * LeftMarginPixels * scrTwips - 2 * scrTwips

If probeX > n Then
If s$ = vbNullString Then
realpos = 0
usedCharLength = 1
Else
realpos = n
usedCharLength = Len(s$) + 1
End If
Else
If probeX <= 30 Then
realpos = 0
usedCharLength = 1
Exit Sub
End If
st = Len(s$)
st1 = st + 1
st0 = 1
While st > st0 + 1
st1 = (st + st0) \ 2
If probeX >= UserControlTextWidth(Mid$(s$, 1, st1)) Then
st0 = st1
Else
st = st1
End If
Wend

If probeX > UserControlTextWidth(Mid$(s$, 1, st1 + 1)) Then
st1 = st1 + 1
Else
If probeX < UserControlTextWidth(Mid$(s$, 1, st1)) Then st1 = st0
If st1 = 2 Then

If probeX < UserControlTextWidth(Mid$(s$, 1, 1)) Then st1 = 1
End If
End If
s$ = Mid$(s$, 1, st1)  '
realpos = UserControlTextWidth(s$)
usedCharLength = Len(s$)
End If
End Sub


Property Let ListindexPrivateUse(item As Long)
Do While ListJoin(item) And item > 0
    item = item - 1
Loop
SkipReadEditflag = False
If item < listcount Then
SELECTEDITEM = item + 1
Else
SELECTEDITEM = 0
End If
End Property
Public Sub ListindexPrivateUseFirstFree(ByVal item As Long)
Dim X As Long
If item < listcount Then
Do While ListJoin(item) And item > 0
    item = item - 1
Loop
SkipReadEditflag = False
For X = item To listcount - 1
If (mList(X).Flags And (fline + joinpRevline)) = 0 Then SELECTEDITEM = X + 1: Exit For
Next X
If item = listcount Then SELECTEDITEM = 0
Else
SELECTEDITEM = 0
End If
End Sub

Private Property Get SELECTEDITEM() As Long
SELECTEDITEM = Mselecteditem
End Property

Private Property Let SELECTEDITEM(ByVal RHS As Long)
If RHS > listcount And RHS > 0 Then
RHS = 0

If RHS > listcount Then Exit Property
End If
Mselecteditem = RHS
End Property

Public Property Get PanPos() As Long
PanPos = scrollme

End Property
Public Property Get PanPosPixels() As Long
If scrollme <> 0 Then PanPosPixels = scrollme / scrTwips
End Property
Public Property Let PanPos(ByVal RHS As Long)
scrollme = RHS
End Property

Public Sub Refresh()
Dim a As Long
Shape Shape1
Shape Shape2
Shape Shape3
a = GdiFlush()
UserControl.Refresh
End Sub
Public Property Get PreserveNpixelsHeaderRightTwips() As Long
PreserveNpixelsHeaderRightTwips = mPreserveNpixelsHeaderRight
End Property

Public Property Let PreserveNpixelsHeaderRightTwips(ByVal RHS As Long)
mPreserveNpixelsHeaderRight = RHS
End Property
Public Property Let SelStartNoEvents(RHS As Long)
'Dim checkline As Long
'RaiseEvent PromptLine(checkline)
'If PromptLineIdent > 0 And (ListIndex = checkline) And PromptLineIdent >= RHS Then RHS = PromptLineIdent + 1
'If Not (mSelstart = RHS) Then
'mSelstart = RHS
'Else
mSelstart = RHS
If mSelstart < 1 Then mSelstart = 1
'End If


End Property

Public Property Get SelStart() As Long
If mSelstart < 1 Then
mSelstart = 1
End If
SelStart = mSelstart
End Property
Public Property Let SelStartEventAlways(ByVal RHS As Long)
Dim checkline As Long
RaiseEvent PromptLine(checkline)
If PromptLineIdent > 0 And (ListIndex = checkline) And PromptLineIdent >= RHS Then RHS = PromptLineIdent + 1
mSelstart = RHS

RaiseEvent ChangeSelStart(RHS)
mSelstart = RHS
End Property
Public Property Let SelStart(ByVal RHS As Long)
Dim checkline As Long
RaiseEvent PromptLine(checkline)
If PromptLineIdent > 0 And (ListIndex = checkline) And PromptLineIdent >= RHS Then RHS = PromptLineIdent + 1
If Not (mSelstart = RHS) Then
mSelstart = RHS
RaiseEvent ChangeSelStart(RHS)
mSelstart = RHS

Else
mSelstart = RHS
End If
End Property
Private Sub ShowMyCaretInTwips(x1 As Long, y1 As Long)
If hWnd <> 0 Then
 With UserControl
 If Not caretCreated Then

 CreateCaret hWnd, 0, .ScaleX(1, 1, 3) + 2, .ScaleY(myt, 1, 3) - 2: caretCreated = True
 End If
SetCaretPos .ScaleX(x1, 1, 3), .ScaleY(y1, 1, 3) + 1
ShowCaret (hWnd)
End With
End If
End Sub

Public Property Get EditFlag() As Boolean
EditFlag = mEditFlag
End Property
Private Property Get EditFlagSpecial() As Boolean
If mTabs > 1 And SELECTEDITEM > 0 Then
    If Not SkipReadEditflag Then
        lastEditFlag = CBool(PropAtColumnNum(SELECTEDITEM - 1, mCurTab, "EDIT"))
        SkipReadEditflag = True
    End If
End If
EditFlagSpecial = lastEditFlag
End Property

Public Property Let EditFlag(ByVal RHS As Boolean)
mEditFlag = RHS
lastEditFlag = RHS
If Not RHS Then If hWnd <> 0 Then DestroyCaret: caretCreated = False
End Property
Public Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long, Optional ByVal offsetX As Long = 0)
Dim a As RECT
CopyFromLParamToRect a, thatRect
a.Bottom = a.Bottom - 1
a.Left = a.Left + offsetX
FillBack thathDC, a, thatbgcolor
End Sub
Public Sub WriteThere(thatRect As Long, aa$, ByVal offsetX As Long, ByVal offsetY As Long, thiscolor As Long)
Dim a As RECT, fg As Long
CopyFromLParamToRect a, thatRect
If a.Left > Width Then Exit Sub
a.Right = WidthPixels
a.Left = a.Left + offsetX
a.top = a.top + offsetY
fg = ForeColor
ForeColor = thiscolor
    DrawText UserControl.hDC, StrPtr(aa$), -1, a, DT_NOPREFIX Or DT_NOCLIP
    ForeColor = fg
End Sub
Public Property Get FontBold() As Boolean
FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal RHS As Boolean)
UserControl.FontBold = RHS
CalcNewFont
 PropertyChanged "Font"
End Property

Public Property Get charset() As Integer
charset = UserControl.Font.charset
End Property

Public Property Let charset(ByVal RHS As Integer)
UserControl.Font.charset = RHS
CalcNewFont
 PropertyChanged "Font"
End Property
Public Sub ExternalCursor(ByVal ExtSelStart, that$, Curcolor As Long)
If HideCaretOnexit Then
If caretCreated Then caretCreated = False: DestroyCaret
Exit Sub
End If

 Dim REALX As Long, RealX2 As Long, myt1
 myt1 = myt - scrTwips * 2
If ExtSelStart <= 0 Then ExtSelStart = 1
                                             DrawStyle = vbNormal
             
                                   REALX = UserControlTextWidth(Mid$(that$, 1, ExtSelStart - 1)) + LeftMarginPixels * scrTwips
              
                                    RealX2 = scrollme + REALX
                                    If (Not marvel) And (havefocus And Not Noflashingcaret) Then
                                          ShowMyCaretInTwips RealX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips + scrTwips
                                    Else
                                    If caretCreated Then caretCreated = False: DestroyCaret
                                              DrawMode = vbCopyPen
                    'If Not NoCaretShow Then
                    Line (RealX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips + scrTwips)-(RealX2 + scrTwips, (SELECTEDITEM - topitem - 1) * myt + myt1 + mHeadlineHeightTwips), Curcolor, BF
                    
                                            ' DrawMode = vbCopyPen
                                 End If
                                 

                                   If Not NoScroll Then If RealX2 > Width * 0.8 * dragslow Then scrollme = scrollme - Width * 0.2 * dragslow: PrepareToShow 10
                                   If RealX2 - Width * 0.2 * dragslow < 0 Then
                              If Not NoScroll Then
                              scrollme = scrollme + Width * 0.2 * dragslow
                              If scrollme > 0 Then scrollme = 0 Else PrepareToShow 10
                                   End If
                                   End If
                     

End Sub


Public Sub ExternalCursor2(ByVal REALX As Long, Curcolor As Long)
If HideCaretOnexit Then
    If caretCreated Then caretCreated = False: DestroyCaret
    Exit Sub
End If
If marvel Then missMouseClick = False
Dim RealX2 As Long, myt1
myt1 = myt - scrTwips * 2

DrawStyle = vbNormal

REALX = REALX * scrTwips + LeftMarginPixels * scrTwips

RealX2 = scrollme + REALX
If (Not marvel) And (havefocus And Not Noflashingcaret) Then
    ShowMyCaretInTwips RealX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips + scrTwips
Else
    If caretCreated Then caretCreated = False: DestroyCaret
    DrawMode = vbCopyPen
    'If Not NoCaretShow Then
    If Not missMouseClick Then
        Line (RealX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips + scrTwips)-(RealX2 + scrTwips, (SELECTEDITEM - topitem - 1) * myt + myt1 + mHeadlineHeightTwips), Curcolor, BF
    End If
    ' DrawMode = vbCopyPen
End If


If Not NoScroll Then If RealX2 > Width * 0.8 * dragslow Then scrollme = scrollme - Width * 0.2 * dragslow: PrepareToShow 10
If RealX2 - Width * 0.2 * dragslow < 0 Then
    If Not NoScroll Then
        scrollme = scrollme + Width * 0.2 * dragslow
        If scrollme > 0 Then scrollme = 0 Else PrepareToShow 10
    End If
End If
                     

End Sub

Private Sub FindRealCursor(ByVal tothere As Long)
' from listindex to tothere
' No center text
tothere = tothere - 1
If tothere = ListIndex Then Exit Sub
Dim thatwidth As Long, c$, Dummy1 As Long
If SelStart < 1 Then
c$ = list(ListIndex)
Else
c$ = Left$(list(ListIndex), SelStart - 1)
End If

thatwidth = UserControlTextWidth(c$) + LeftMarginPixels * scrTwips
''If mSelstart > Len(c$) Then Exit Sub
Dim where As Long, oldmselstart As Long
oldmselstart = mSelstart
REALCUR tothere, thatwidth, (Dummy1), where
'If Len(list(tothere)) < Len(c$) And oldmselstart > Len(list(tothere)) Then
'mSelstart = oldmselstart
'Else
mSelstart = where + 1
'End If

End Sub

Public Sub Shutdown()
waitforparent = False
BlinkON = False
BlinkTimer.enabled = False
Timer1.enabled = False
Timer2.enabled = False
Timer3.enabled = False
BlinkTimer.Interval = 10000
Timer1.Interval = 10000
Timer2.Interval = 10000
Timer3.Interval = 10000
enabled = False
End Sub

Public Sub DragNow()
marvel = True
UserControl.OLEDrag
marvel = False
End Sub
Friend Sub MarkWord()
If ListIndex < 0 Then Exit Sub
Dim one$
Dim mline$, Pos As Long, Epos As Long, oldselstart As Long
RaiseEvent PureListOn
mline$ = list(ListIndex)
RaiseEvent PureListOff
'Enabled = False
Pos = SelStart
If Pos <> 0 Then
Dim mypos As Long, ogt As String, this$
If Pos <= 0 Then Pos = 1
Epos = Pos
Do While Pos > 0
If InStr(1, WordCharLeft, Mid$(mline$, Pos, 1)) Then Exit Do
Pos = Pos - 1
Loop
If Pos > 0 Then If InStr(1, WordCharLeftButIncluded, Mid$(mline$, Pos, 1)) Then Pos = Pos - 1
Do While Epos <= Len(mline$)
one$ = Mid$(mline$, Epos, 1)
If InStr(1, WordCharRightButIncluded, one$) Then Epos = Epos + 1: Exit Do
If InStr(1, WordCharRight, one$) Then Exit Do
Epos = Epos + 1
Loop
If (Epos - Pos - 1) > 0 Then
    If Pos = 0 Then
    Pos = MyTrimL(mline$)
    If Pos > Len(mline$) Then Pos = 0 Else Pos = Pos - 1
    End If
    this$ = Mid$(mline$, Pos + 1, Epos - Pos - 1)
    
    RaiseEvent WordMarked(this$)
    If this = vbNullString Or Not EditFlag Then Exit Sub
    oldselstart = SelStart
    MarkNext = 0
    If (oldselstart - Pos - 1) > (Epos - oldselstart) Then
        SelStart = Pos + 1
        RaiseEvent markin
        MarkNext = 1
        SelStart = Epos
        RaiseEvent MarkOut
    Else
        SelStart = Epos
        RaiseEvent markin
        SelStart = Pos + 1
        MarkNext = 1
        RaiseEvent MarkOut
        SelStart = Pos + 1
    End If
ShowMe2
ElseIf Not EditFlag Then
    PressSoft
End If
End If
'Enabled = True

End Sub
Public Sub MarkUp()
MarkNext = 0
RaiseEvent selected(ListIndex + 1)
RaiseEvent markin
MarkNext = 1
SelStart = 1
ListindexPrivateUse = 0
RaiseEvent selected(ListIndex + 1)
RaiseEvent MarkOut
ShowMe2

End Sub
Public Sub MarkDown()
MarkNext = 0
RaiseEvent selected(ListIndex + 1)
RaiseEvent markin
MarkNext = 1
ListindexPrivateUse = listcount - 1
SelStart = Len(list(ListIndex)) + 1
RaiseEvent selected(ListIndex + 1)
RaiseEvent MarkOut
ShowMe2
End Sub
Public Sub MarkALL()
MarkNext = 0
ListindexPrivateUse = 0
SelStart = 1
RaiseEvent selected(ListIndex + 1)
RaiseEvent markin
MarkNext = 1
ListindexPrivateUse = listcount - 1
SelStart = Len(list(ListIndex)) + 1
RaiseEvent selected(ListIndex + 1)
RaiseEvent MarkOut
ShowMe2
End Sub
Public Sub ShowPan()
Dim LL As Long
If listcount > 0 Then
    If ListIndex >= 0 Then
            If (ListIndex - topitem) >= 0 And (ListIndex - topitem) < lines Then
                    SkipReadEditflag = False
                    If SelStart = 0 Then
                    LL = scrollme
                    Else
                    LL = UserControlTextWidthPixels(Left$(list(ListIndex), SelStart)) + scrollme
                    End If
                    If LL < WidthPixels Then
                    ShowMe
                    Exit Sub
                    ElseIf LL >= 0 Then
                    ShowMe2
                    Exit Sub
                    End If
           
            End If
    End If
End If
chooseshow
End Sub

Public Property Get MousePointer() As Integer
MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal RHS As Integer)
UserControl.MousePointer = RHS
End Property
Public Property Set MouseIcon(RHS)
Set UserControl.MouseIcon = RHS
End Property
Function GetLocale() As Long
    Dim R&
      R = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF&
      GetLocale = val("&H" & Right(Hex(R), 4))
End Function
Function GetKeY(ascii As Integer) As String
    Dim Buffer As String, ret As Long

    Buffer = String$(514, 0)
    Dim R&
      R = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF&
      R = val("&H" & Right(Hex(R), 4))
    ret = GetLocaleInfo(R, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    If ret > 0 Then
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, CLng(val("&h" + Left$(Buffer, ret - 1))))))
    Else
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, 1033)))
    End If
End Function
Sub GetKeY2(ascii As Integer, shift As Integer)
Dim acc As Long
acc = MapVirtualKey(ascii, 2)
If ascii > 0 And acc = 0 Then acc = ascii + 500
If acc = 0 Then Exit Sub
If (shift And 1) = 1 Then acc = acc + 1000
If (shift And 2) = 2 Then acc = acc + 10000
If (shift And 4) = 4 Then acc = acc + 100000

RaiseEvent AccKey(acc)
If acc = 0 Then ascii = 0: shift = 0
End Sub
Public Function LineTopOffsetPixels()
Dim nr As RECT, a$
a$ = "fg"
CalcRect1 UserControl.hDC, a$, nr
LineTopOffsetPixels = (mytPixels - nr.Bottom) / 2
End Function
Private Sub LineAddTopOffsetPixels(c, nr As RECT)
Dim R As RECT
'' only in  WrapText
    R = nr
    R.top = 0
    R.Left = 0
    R.Right = nr.Right - nr.Left
    R.Bottom = nr.Bottom - nr.top
    DrawTextEx UserControl.hDC, StrPtr(c), -1, R, DT_CALCRECT Or DT_WORDBREAK Or DT_NOPREFIX Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
    R.top = (nr.Bottom - nr.top - R.Bottom) / 2
    nr.top = nr.top + R.top
     




End Sub

Private Sub Shape(a As Myshape, Optional Left As Long = -1, Optional top As Long = -1, Optional Width As Long = -1, Optional Height As Long = -1)
If Left <> -1 Then a.Left = Left
If top <> -1 Then a.top = top
If Width <> -1 Then a.Width = Width
If Height <> -1 Then a.Height = Height
Dim th As RECT, my_brush As Long, br2 As Long
If a.Visible Then
With th
.top = a.top / scrTwips
.Left = a.Left / scrTwips
.Bottom = .top + a.Height / scrTwips
.Right = .Left + a.Width / scrTwips
End With

 br2 = CreateSolidBrush(BarHatchColor)
   
   If a.hatchType = 1 Then

    
    If BarHatch <> -1 Then
    SetBkColor UserControl.hDC, BarColor
        my_brush = CreateHatchBrush(BarHatch, BarHatchColor)
    Else
        my_brush = CreateSolidBrush(BarColor)
    End If

  FillRect UserControl.hDC, th, my_brush
 Else
  my_brush = CreateSolidBrush(BarColor)
  FillRect UserControl.hDC, th, my_brush
End If

If BarHatch <> -1 Then FrameRect UserControl.hDC, th, br2

  DeleteObject my_brush
  DeleteObject br2
End If
End Sub
Function DoubleClickArea(ByVal X As Long, ByVal y As Long, ByVal Xorigin As Long, ByVal Yorigin As Long, setupxy As Long) As Boolean
   If Abs(X - Xorigin) < setupxy And Abs(y - Yorigin) < setupxy Then
        preservedoubleclick = doubleclick
    Else
        preservedoubleclick = 0
   End If
   DoubleClickArea = Not preservedoubleclick = 0
End Function
Function DoubleClickCheck(Button As Integer, ByVal item As Long, ByVal X As Long, ByVal y As Long, ByVal Xorigin As Long, ByVal Yorigin As Long, setupxy As Long, itemline As Long) As Boolean
' doubleclick
If item = itemline Then
   If Abs(X - Xorigin) < setupxy And Abs(y - Yorigin) < setupxy Then
        If Not nopointerchange Then MousePointer = 1
        FloatList = False
        If Button = 1 Then
            doubleclick = doubleclick + 1 + preservedoubleclick
            preservedoubleclick = 0
            If doubleclick = 1 Then
                timestamp1 = Timer
            ElseIf doubleclick > 1 Then
                If (timestamp1 + 1.5) < Timer Then
                    doubleclick = 1
                    timestamp1 = Timer
                Else
                    timestamp1 = Timer + 100
                    DoubleClickCheck = True: Exit Function
                End If
            End If
            Button = 0
        End If
    Else
        doubleclick = 0
        FloatList = True
    End If
End If
End Function
Function SingleClickCheck(Button As Integer, ByVal item As Long, ByVal X As Long, ByVal y As Long, ByVal Xorigin As Long, ByVal Yorigin As Long, setupxy As Long, itemline As Long) As Boolean
If item = itemline Then
   If Abs(X - Xorigin) < setupxy And Abs(y - Yorigin) < setupxy Then
      If Not nopointerchange Then MousePointer = 1
        FloatList = False
            If Button = 1 Then
            SingleClickCheck = True
                Exit Function
    

                       
            End If
    Else
       FloatList = True
    End If
End If
End Function
Public Property Get Parent() As Variant
On Error GoTo there
If UserControl.Parent Is Nothing Then Exit Property
Set Parent = UserControl.Parent
there:
End Property


Public Sub Curve(Optional t As Boolean = False, Optional factor As Single = 1)
Dim hRgn As Long
If Int(25 * factor) > 2 Then
hRgn = CreateRoundRectRgn(0, 0, WidthPixels, HeightPixels, 25 * factor, 25 * factor)
SetWindowRgn Me.hWnd, hRgn, t
'DeleteObject hRgn
End If
End Sub
Public Sub ShowMenu()
    'dropkey = True
    RaiseEvent DeployMenu

   
End Sub
Public Sub CascadeSelect(ByVal item As Long)
RaiseEvent CascadeSelect(item)
End Sub
Public Property Let BlinkTime(ByVal t As Long)
BlinkON = True <> 0
mBlinkTime = t
Timer1.Interval = t
Timer1.enabled = True
End Property
Public Property Get BlinkTime() As Long
BlinkTime = mBlinkTime
End Property
Sub DestCaret()
 DestroyCaret
 caretCreated = False
End Sub
Private Function MyTrimL(s$) As Long
Dim i&, l As Long
Dim P2 As Long, P1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL = 1: Exit Function
  P2 = StrPtr(s): l = l - 1
  p4 = P2 + l * 2
  For i = P2 To p4 Step 2
  GetMem2 i, P1
  Select Case P1
    Case 32, 160, 9
    Case Else
     MyTrimL = (i - P2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimL = l + 2
End Function
Private Function NLtrim$(a$)
If Len(a$) > 0 Then NLtrim$ = Mid$(a$, MyTrimL(a$))
End Function

Public Property Get FontName() As Variant
FontName = UserControl.Font
End Property
Public Sub PaintPicture1(pic As StdPicture, x1 As Long, y1 As Long, width1 As Long, height1 As Long)
    UserControl.ScaleMode = 3
    UserControl.PaintPicture pic, x1, y1, width1, height1
    UserControl.ScaleMode = 1
End Sub
Public Property Let icon(RHS)
On Error Resume Next
nopointerchange = True
If IsNumeric(RHS) Then
    UserControl.MousePointer = CInt(RHS)
Else
    Dim aPic As StdPicture, s$, Scr As Object
    s$ = RHS
    If s$ <> vbNullString Then
                    s$ = CFname(s$)
                    If s$ <> vbNullString Then
                    If LCase(Right$(s$, 4)) = ".ico" Or LCase(Right$(s$, 4)) = ".cur" Then
                        Set aPic = LoadPicture(GetDosPath(s$))
                    Else
                        
                        Set aPic = LoadMyPicture(GetDosPath(s$))
                        End If
                         Set UserControl.MouseIcon = Form1.Picture2.MouseIcon: UserControl.MousePointer = 99
                 
                        If aPic Is Nothing Then MissCdib: Exit Property
                        Set UserControl.MouseIcon = aPic
                        UserControl.MousePointer = 99
                    Else
                    MissFile
                    End If
                Else
                 Set UserControl.MouseIcon = Form1.Picture2.MouseIcon: UserControl.MousePointer = 99
                    
                    
                End If
End If
oldpointer = UserControl.MousePointer
End Property
Friend Sub SetIcon(RHS, mpointer As Integer)
On Error Resume Next
nopointerchange = True
Set UserControl.MouseIcon = RHS
oldpointer = mpointer
UserControl.MousePointer = mpointer
End Sub
Public Property Get icon()
Set icon = UserControl.MouseIcon
End Property
Private Sub chooseshow()
If MultiLineEditBox Then
ShowMe2
Else
ShowMe
End If
End Sub
Public Sub HideTheCaret()
Dim what As Long
    If hWnd <> 0 Then
   ' If caretCreated Then
    caretCreated = False: DestroyCaret: havefocus = True: missMouseClick = True
    'End If
    RaiseEvent CaretDeal(what)
    If what = 0 Then
    Timer1.enabled = False

     End If
    End If
End Sub
Public Sub ShowTheCaret()
Dim what As Long
    If hWnd <> 0 Then
        what = 1
        RaiseEvent CaretDeal(what)
        If what = 1 Then
        havefocus = True
        dragfocus = False
        SoftEnterFocus
        End If
    End If
End Sub
Private Property Get ParentPreview() As Boolean
On Error Resume Next
If UserControl.Parent.previewKey Then ParentPreview = True
End Property
Public Function GetTopUserControl(ByVal UserControl As Object) As VB.UserControl
If UserControl Is Nothing Then Exit Function
Dim TopUserControl As VB.UserControl, TempUserControl As VB.UserControl
CopyMemory TempUserControl, ObjPtr(UserControl), 4
Set TopUserControl = TempUserControl
CopyMemory TempUserControl, 0&, 4
With TopUserControl
If .ParentControls.count > 0 Then
    Dim OldParentControlsType As VBRUN.ParentControlsType
    OldParentControlsType = .ParentControls.ParentControlsType
    .ParentControls.ParentControlsType = vbExtender
    If TypeOf .ParentControls(0) Is VB.VBControlExtender Then
        .ParentControls.ParentControlsType = vbNoExtender
        CopyMemory TempUserControl, ObjPtr(.ParentControls(0)), 4
        Set TopUserControl = TempUserControl
        CopyMemory TempUserControl, 0&, 4
        Dim TempParentControlsType As VBRUN.ParentControlsType
        Do
            With TopUserControl
            If .ParentControls.count = 0 Then Exit Do
            TempParentControlsType = .ParentControls.ParentControlsType
            .ParentControls.ParentControlsType = vbExtender
            If TypeOf .ParentControls(0) Is VB.VBControlExtender Then
                .ParentControls.ParentControlsType = vbNoExtender
                CopyMemory TempUserControl, ObjPtr(.ParentControls(0)), 4
                Set TopUserControl = TempUserControl
                CopyMemory TempUserControl, 0&, 4
                .ParentControls.ParentControlsType = TempParentControlsType
            Else
                .ParentControls.ParentControlsType = TempParentControlsType
                Exit Do
            End If
            End With
        Loop
    End If
    .ParentControls.ParentControlsType = OldParentControlsType
End If
End With
Set GetTopUserControl = TopUserControl
End Function
Public Property Get Column() As Long
    Column = mCurTab
End Property
Public Property Let Column(ByVal RHS As Long)
Dim lcnt As Long
If RHS >= 1 And RHS <= mTabs Then
    mCurTab = RHS
    TwipsCurTab = 0
    For lcnt = 1 To mCurTab - 1
        TwipsCurTab = TwipsCurTab + mParts(lcnt) * scrTwips
    Next lcnt
    mSelstart = 1
End If
End Property

Public Property Get Columns() As Long
Columns = mTabs
End Property

Public Property Let Columns(ByVal RHS As Long)
If RHS >= 1 Then
mTabs = RHS
ReDim Preserve mParts(1 To mTabs)
End If
End Property
Public Property Let ColumnWidth(ByVal Col As Long, ByVal RHS)
' twips
Dim lcnt As Long
On Error Resume Next
If Col >= 1 And Col <= mTabs And IsNumeric(RHS) Then
    mParts(Col) = Abs(RHS) / scrTwips
    TwipsCurTab = 0
      For lcnt = 1 To mCurTab - 1
        TwipsCurTab = TwipsCurTab + mParts(lcnt) * scrTwips
        Next lcnt
End If
End Property
Public Property Get ColumnWidth(ByVal Col As Long)
On Error Resume Next
If Col >= 1 And Col <= mTabs Then
    ColumnWidth = CVar(mParts(Col) * scrTwips)
End If
End Property
Public Sub QuickSortExtended(ByVal Col As Long, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
On Error GoTo abc1
Dim Piv As String, tmp As String
     If UB - LB = 1 Then
     M1 = LB
      If StrComp(Mid$(listAtColumn(M1, Col), SkipChars), Mid$(listAtColumn(UB, Col), SkipChars), mSortstyle) = 1 Then SwapListItems M1, UB
      Exit Sub
     Else
       M1 = (LB + UB) \ 2 '+ 1
             If StrComp(Mid$(listAtColumn(M1, Col), SkipChars), Mid$(listAtColumn(LB, Col), SkipChars), mSortstyle) = 0 Then
                M2 = UB - 1
                M1 = LB
                Piv = Mid$(listAtColumn(LB, Col), SkipChars)
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If StrComp(Mid$(listAtColumn(UB, Col), SkipChars), Piv, mSortstyle) = -1 Then SwapListItems LB, UB
                        Exit Sub
                    End If
                Loop Until StrComp(Mid$(listAtColumn(M1, Col), SkipChars), Piv, mSortstyle) <> 0
                Piv = Mid$(listAtColumn(M1, Col), SkipChars)
                If M1 > LB Then If StrComp(Mid$(listAtColumn(LB, Col), SkipChars), Piv, mSortstyle) = 1 Then SwapListItems M1, LB: Piv = Mid$(listAtColumn(M1, Col), SkipChars)
            Else
                Piv = Mid$(listAtColumn(M1, Col), SkipChars)
                M1 = LB
                Do While StrComp(Mid$(listAtColumn(M1, Col), SkipChars), Piv, mSortstyle) = -1: M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While StrComp(Mid$(listAtColumn(M2, Col), SkipChars), Piv, mSortstyle) = 1: M2 = M2 - 1: Loop
      If M1 <= M2 Then
       If M1 <> M2 Then SwapListItems M1, M2
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While StrComp(Mid$(listAtColumn(M1, Col), SkipChars), Piv, mSortstyle) = -1: M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortExtended Col, LB, M2
    If M1 < UB Then QuickSortExtended Col, M1, UB
    Exit Sub
abc1:
    
    
End Sub

Sub SwapListItems(T1 As Long, T2 As Long)
    If T1 = T2 Then Exit Sub
    Dim emp As itemlist
    emp = mList(T1)
    mList(T1) = mList(T2)
    mList(T2) = emp
End Sub
Friend Function compact(ByVal fr As Long, ByVal ed As Long) As Long
    Dim i As Long, many As Long, pc As Long
    Dim emp() As itemlist, em As itemlist
    em.Flags = joinpRevline
    ReDim emp(fr To ed) As itemlist
    pc = fr
    For i = fr To ed
        If (mList(i).Flags And joinpRevline) = 0 Then
            emp(pc) = mList(i)
            pc = pc + 1
        Else
            many = many + 1
        End If
    Next i
    For i = fr To fr + many - 1
        mList(i) = em
    Next i
    pc = fr
    For i = i To ed
        mList(i) = emp(pc): pc = pc + 1
    Next i
    compact = many
End Function
Friend Sub Expand(ByVal fr As Long, ByVal ed As Long, many As Long)
    Dim em As itemlist, pc As Long, i As Long, h As Integer, j As Long, mm As Long
    em.Flags = joinpRevline
    
    pc = fr + many
    mm = many
    For i = fr To ed
        mList(i) = mList(pc)
        h = mList(i).morerows
        For j = i + 1 To i + h
            mList(j) = em
            mm = mm - 1
        Next j
        i = i + h
        pc = pc + 1
        If mm = 0 Then Exit For
    Next i
End Sub
Public Sub InsertColumn(where As Long, colWidthTwips As Long, many)
If itemcount = 0 Or BlockItemcount Then Exit Sub
If IsMissing(many) Then many = 1 Else many = Int(Abs(many))
If many = 0 Then many = 1
Dim i As Long, ja As JsonArray
If where > mTabs Then
    many = many + where - 1 - mTabs
    mTabs = mTabs + many
    ReDim Preserve mParts(1 To mTabs)
    For i = mTabs + 1 - many To mTabs
        mParts(i) = colWidthTwips \ scrTwips
    Next i
Else
    i = mTabs
    mTabs = mTabs + many
    ReDim Preserve mParts(1 To mTabs)
    For i = i To where Step -1
        mParts(i + many) = mParts(i)
    Next i
    For i = mTabs - many To where Step -1
        mParts(i) = colWidthTwips \ scrTwips
    Next i
End If
For i = 0 To listcount - 1
    With mList(i)
        If (.Flags And joinpRevline) = 0 Then
        If Not .content Is Nothing Then
            If .content.ExistKey("C") Then
                Set ja = .content.ValueObj
                ja.InsertNull where, CLng(many)
            End If
            If .content.ExistKey("P") Then
                Set ja = .content.ValueObj
                ja.InsertNull where, CLng(many)
            End If
        End If
        End If
    End With
Next i
where = 0
For i = 1 To mCurTab - 1
    where = where + mParts(i)
Next i
TwipsCurTab = where * scrTwips
End Sub

Friend Sub HideBarAsap()
If BarVisible Then
    Hidebar = False
    BarVisible = False
End If
End Sub
