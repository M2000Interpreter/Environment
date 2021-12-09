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
Public BypassKey As Boolean
Public BlinkON As Boolean
Private mBlinkTime
Public UseTab As Boolean
Public InternalCursor As Boolean
Public OverrideShow As Boolean
Public HideCaretOnexit As Boolean
Public overrideTextHeight As Long
Public AutoHide As Boolean, NoWheel As Boolean
Private missMouseClick As Boolean
Public bypassfirstClick As Boolean
Private Shape1 As Myshape, Shape2 As Myshape, Shape3 As Myshape
Private Type RECT
        Left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type
Private Type itemlist
    selected As Boolean  ' use this for multiselect or checked
    checked As Boolean  ' use this to use list item as menu
    radiobutton As Boolean  ' use this to checked like radio buttons ..with auto unselect between to lines...or all list if not lines foundit
    content As String
    contentID As String
    Line As Boolean
End Type
Private ehat$
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
Private Declare Function SetCaretPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hDC As Long, ByVal lpsz As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long, ByVal lpDrawTextParams As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, RetVal As Integer)

Private Const PS_NULL = 5
Private Const PS_SOLID = 0
Public restrictLines As Long
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
Dim m_backcolor As Long
Dim m_ForeColor As Long
'Dim m_Enabled As Boolean
Dim m_font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_CapColor As Long
Dim m_dcolor As Long

Dim m_showbar As Boolean
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
Event Maybelanguage()
Event MouseUp(x As Single, y As Single)
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
Event NeedDoEvents()
Event OutPopUp(x As Single, y As Single, myButton As Integer)
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
Event PreviewKeyboardUnicode(ByVal a$)
Event Find(Key As String, where As Long, skip As Boolean)
Event ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Event ExposeListcount(cListCount As Long)
Event ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
Event MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
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
Event RTL2(s$, where As Long, mark10 As Long, mark20 As Long, Offset As Long)
Event DelExpandSS()
Event SetExpandSS(val As Long)
Event ExpandSelStart(val As Long)
Event Fkey(a As Integer)
Event CaretDeal(Deal As Long)
Event AccKey(m As Long)
Private State As Boolean
Private secreset As Boolean
Private scrollme As Long
Private scrolledit As Long
Private lY As Long, dr As Boolean
Private drc As Boolean
Private scrTwips As Single
Private cY As Long
Private cX As Long
Dim myt As Long, FaceBlink As Boolean
Dim mytPixels As Long
Public NoMoveDrag As Boolean
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
Private barwidth As Long
Private NoFire As Boolean
Public addpixels As Long
Public StickBar As Boolean
Dim Hidebar As Boolean
Dim myEnabled As Boolean
Public WrapText As Boolean
Public CenterText As Boolean
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
Public HeadLineHeightMinimum As Long
Private mPreserveNpixelsHeaderRight As Long
Public AutoPanPos As Boolean   ' used if we have no EditFlag
Public FloatLimitLeft As Long
Public FloatLimitTop As Long
Public mEditFlag As Boolean
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
    x As Long
    y As Long
End Type
 Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type
Private timestamp As Double, timestamp1 As Double
Private doubleclick As Long, preservedoubleclick As Long
Private dbLx As Long, dbly As Long
Private PX As Long, PY As Long
Public SkipForm As Boolean
Public dropkey As Boolean
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
Private tm As TEXTMETRICW
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
    Exit Function
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
Public Sub PrepareToShow(Optional delay As Single = 10)
 barwidth = UserControlTextWidth("W")
 CalcAndShowBar1
Timer1.enabled = False
If delay < 1 Then delay = 1
If fast Then
fast = False
Timer1.Interval = delay
Else
Timer1.Interval = delay * 5
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
    
 If BackStyle = 0 Then UserControl.BorderStyle = m_BorderStyle Else UserControl.BorderStyle = 0
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
If UserControl.Ambient.UserMode = False Then
    Repaint
    SELECTEDITEM = 0
    CalcAndShowBar
    ShowMe
End If

PropertyChanged "Text"
End Property
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
Text = space$(5)
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
State = True

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

State = False
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
myt = (UserControl.Scaleheight - mHeadlineHeightTwips) / restrictLines
mytPixels = CLng(myt / scrTwips)
myt = CLng(mytPixels * scrTwips)
Else
mytPixels = (UserControlTextHeightPixels() + addpixels)
myt = mytPixels * scrTwips

End If


    m_showbar = RHS
    barwidth = UserControlTextWidth("W")
    
    State = True
    Value = 0
    State = False
 
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

Public Property Let backcolor(ByVal RHS As OLE_COLOR)

    m_backcolor = RHS
UserControl.backcolor = RHS
  PropertyChanged "BackColor"
    
End Property
Public Property Get backcolor() As OLE_COLOR
    backcolor = m_backcolor 'UserControl.Backcolor
End Property
Public Property Get forecolor() As OLE_COLOR
    forecolor = m_ForeColor
End Property

Public Property Let forecolor(ByVal RHS As OLE_COLOR)
    m_ForeColor = RHS
    UserControl.forecolor = Abs(RHS)
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
    Extender.TabStop = RHS
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
GetTextMetrics UserControl.hDC, tm
AveCharWith = tm.tmAveCharWidth
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
    GetTextMetrics UserControl.hDC, tm
    If restrictLines > 0 Then
        myt = (UserControl.Scaleheight - mHeadlineHeightTwips) / restrictLines
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
GetTextMetrics UserControl.hDC, tm
If restrictLines > 0 Then
myt = (UserControl.Scaleheight - mHeadlineHeightTwips) / restrictLines
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
GetTextMetrics UserControl.hDC, tm
If restrictLines > 0 Then
myt = (UserControl.Scaleheight - mHeadlineHeightTwips) / restrictLines
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
    If m_BackStyle = 0 Then UserControl.BorderStyle = m_BorderStyle Else UserControl.BorderStyle = 0
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
If Not NoWheel Then RaiseEvent RegisterGlist(Me)
End Sub

Private Sub UserControl_GotFocus()
RaiseEvent CheckGotFocus
havefocus = True
dragfocus = False
SoftEnterFocus
If Not NoWheel Then RaiseEvent RegisterGlist(Me)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

If BypassKey Then KeyAscii = 0: Exit Sub
If dropkey Then KeyAscii = 0: Exit Sub
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
End If
Else
    If Not State Then
       
        If KeyAscii = 13 And myEnabled And Not MultiLineEditBox Then
            KeyAscii = 0
            If SELECTEDITEM < 0 Then
                
            ElseIf SELECTEDITEM > 0 Then
                secreset = False
                RaiseEvent Selected2(SELECTEDITEM - 1)
            End If
            If ParentPreview Then RaiseEvent PreviewKeyboardUnicode(Chr$(13))
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
                If maxchar = 0 Or (maxchar > Len(list(SELECTEDITEM - 1)) Or MultiLineEditBox) Then
                   'If Len(UKEY$) <> 2 Then
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
                     If EditFlag Then
                    bb = enabled
                    enabled = False
                    RaiseEvent HaveMark(b1)
                    RaiseEvent PushUndoIfMarked
                    RaiseEvent MarkDelete(False)
                    If b1 Then RaiseEvent GroupUndo
                    enabled = bb
                End If
                    If EditFlag And ((KeyAscii > 32 And KeyAscii <> 127) Or (KeyAscii = 9 And UseTab)) Then
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
         list(SELECTEDITEM - 1) = Left$(list(SELECTEDITEM - 1), SelStart - 3) + kk$ + Mid$(list(SELECTEDITEM - 1), SelStart - 2)
          RaiseEvent PureListOff
         Else
         RaiseEvent SetExpandSS(mSelstart + 1)
            SelStartEventAlways = SelStart + 1
            RaiseEvent DelExpandSS
            RaiseEvent PureListOn
            
            list(SELECTEDITEM - 1) = Left$(list(SELECTEDITEM - 1), SelStart - 2) + kk$ + Mid$(list(SELECTEDITEM - 1), SelStart - 1)
            
          RaiseEvent PureListOff
          
            End If


            RaiseEvent SetExpandSS(mSelstart)
        
           '
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
KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
If BlinkON Then
    If Not BlinkTimer.enabled Then
        BlinkTimer.Interval = mBlinkTime
        BlinkTimer.enabled = True
        FaceBlink = Not FaceBlink
        RaiseEvent BlinkNow(FaceBlink)
        ShowPan
    End If
Else
    Timer1.enabled = False
    Timer1.Interval = 30
    If Not enabled Then Exit Sub

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
State = True
 On Error Resume Next
 Err.Clear

    If SELECTEDITEM >= listcount Then
 Value = listcount - 1
  State = False
  Exit Sub
        Else
    Value = topitem
    End If
    State = False
 If Timer2.enabled = False Then
If SELECTEDITEM - topitem > 0 And SELECTEDITEM - topitem - 1 <= lines And cX > 0 And cX < UserControl.Scalewidth Then
 If SELECTEDITEM > 0 Then
         If Not BlockItemcount Then
             REALCUR SELECTEDITEM - 1, cX - scrollme, dummy, mSelstart
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
RaiseEvent ScrollSelected(SELECTEDITEM, cY * myt)

End If
End Sub


Public Sub SoftEnterFocus()

If bypassfirstClick Then
missMouseClick = True
FreeMouse = True
End If
State = Not enabled
Noflashingcaret = Not enabled
If EditFlag Then
If Not Spinner Then State = Not MultiLineEditBox
End If
RaiseEvent ShowExternalCursor
If Not Timer1.enabled Then PrepareToShow 5
End Sub

Private Sub SoftExitFocus()
If Not havefocus Then Exit Sub
Noflashingcaret = True
State = True ' no keyboard input

secreset = False
Timer2.enabled = False
FreeMouse = False
If (Not BypassLeaveonChoose) And LeaveonChoose Then
If Not MultiLineEditBox Then If EditFlag And caretCreated Then caretCreated = False: DestroyCaret
SELECTEDITEM = -1: RaiseEvent Selected2(-2)
End If
If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar

RaiseEvent ShowExternalCursor
State = False
End Sub



Private Sub UserControl_Initialize()
tParam.cbSize = LenB(tParam)
tParam.iTabLength = 4
mTabStop = True
Buffer = 100
Set m_font = UserControl.Font
ReDim mList(0 To Buffer)
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
 backcolor = m_def_BackColor
   forecolor = m_def_ForeColor
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
If EditFlag Then
    RaiseEvent DelExpandSS
    If mSelstart = 1 Then
        mSelstart = Len(list(ListIndex)) - Len(NLtrim(list(ListIndex))) + 1
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
If EditFlag Then
    mSelstart = Len(list(ListIndex)) + 1
    RaiseEvent SetExpandSS(mSelstart)
Else
    SELECTEDITEM = listcount + 1
    osel = SELECTEDITEM
    secreset = False
    Do While Not (Not ListSep(ListIndex) Or ListIndex = 0)
        secreset = False
        SELECTEDITEM = osel - 1
        osel = SELECTEDITEM
    Loop
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
        If EditFlag Then
            RaiseEvent DelExpandSS
            If mSelstart = 1 Then
                mSelstart = Len(list(ListIndex)) - Len(NLtrim(list(ListIndex))) + 1
            Else
                mSelstart = 1
           End If
        End If
    ElseIf SELECTEDITEM - lines < 0 Then
        If SELECTEDITEM - 1 > 0 Then
            ShowThis SELECTEDITEM - 1
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
    osel = SELECTEDITEM
    secreset = False
    While ListSep(osel) And Not osel = 0
        secreset = False
        osel = osel - 1
    Wend
    secreset = False
    If ListSep(osel - 1) Then
        ListIndex = LastListIndex
    Else
        SELECTEDITEM = osel + 1
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
        Loop Until Not ListSep(ListIndex) Or ListIndex = 0
        secreset = False
        If Not MultiLineEditBox Then
            If ListSep(ListIndex) Then ListIndex = LastListIndex Else ShowThis osel
        Else
            ShowThis osel
        End If
    Else
        If Not MultiLineEditBox Then If ListSep(ListIndex) Then ListIndex = LastListIndex
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
        Loop Until Not ListSep(ListIndex) Or ListIndex = lcnt - 1
        secreset = False
        If Not MultiLineEditBox Then
            If ListSep(ListIndex) Then ListIndex = LastListIndex Else ShowThis osel
        Else
            ShowThis osel
        End If
    Else
        If Not MultiLineEditBox Then If ListSep(ListIndex) Then ListIndex = LastListIndex
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
        If EditFlag Then
            mSelstart = Len(list(ListIndex)) + 1
            RaiseEvent SetExpandSS(mSelstart)
        End If
    ElseIf SELECTEDITEM + (lines + 1) \ 2 >= lcnt Then
        If listcount > SELECTEDITEM Then
            ShowThis SELECTEDITEM + 1
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
    osel = SELECTEDITEM
    secreset = False
    While ListSep(osel) And Not ((osel = lcnt - 1) Or (osel < 1))
        secreset = False
        osel = osel + 1
    Wend
    secreset = False
    If ListSep(osel) Then
        ListIndex = LastListIndex
    Else
        SELECTEDITEM = osel
        ShowThis osel
    End If
    RaiseEvent ChangeSelStart(SelStart)
    If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
    If shift <> 0 Then If MarkNext > 0 Then RaiseEvent KeyDownAfter(KeyCode, shift)
    shift = 0: KeyCode = 0: Exit Sub
Case vbKeySpace
    If SELECTEDITEM > 0 Then
        If EditFlag Then
        If mSelstart = 0 Then mSelstart = 1
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
            PrepareToShow 10
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
If EditFlag Then
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
            mSelstart = Len(list(ListIndex)) + 1
            If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
            If Not NoEvents Then If SELECTEDITEM > 0 Then RaiseEvent selected(SELECTEDITEM)
        End If
    ElseIf SelStart > val Then
        mSelstart = SelStart - val
        If shift = 0 Then RaiseEvent SetExpandSS(mSelstart)
    End If
Else
    If Not NoEvents Then If SELECTEDITEM > 0 Then If Not Arrows2Tab Then RaiseEvent selected(SELECTEDITEM)
End If
Case vbKeyRight
If EditFlag Then
    val = 1
    RaiseEvent AddSelStart(val, shift)
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
    If Not NoEvents Then If SELECTEDITEM > 0 Then If Not Arrows2Tab Then RaiseEvent selected(SELECTEDITEM)
End If
Case vbKeyDelete
If EditFlag Then
    If mSelstart = 0 Then mSelstart = 1
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
End If
Case vbKeyBack
If EditFlag Then
    If SelStart > 1 Then
        val = 1
        RaiseEvent PureListOn
        RaiseEvent SubSelStart(val, shift)
        SelStart = SelStart - val  ' make it a delete because we want selstart to take place before list() take value     RaiseEvent PureListOn
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
End If
Case vbKeyReturn
If MultiLineEditBox Then
    RaiseEvent SplitLine
    RaiseEvent SetExpandSS(mSelstart)
    RaiseEvent RemoveOne(vbCrLf)
Else
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
End Sub

Private Sub UserControl_LostFocus()
doubleclick = 0
Fkey = 0
If Not NoWheel Then RaiseEvent UnregisterGlist
RaiseEvent CheckLostFocus
If myEnabled Then SoftExitFocus
havefocus = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
' cut area
On Error GoTo there
If dropkey Then Exit Sub
If missMouseClick Then Exit Sub
nowX = x
nowY = y

If (Button And 2) = 2 Then Exit Sub
If myt = 0 Then Exit Sub
FreeMouse = True
Dim YYT As Long, oldbutton As Integer
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
        RaiseEvent ExposeItemMouseMove(Button, -1, CLng(x / scrTwips), CLng(y / scrTwips))
        If (x < Width - mPreserveNpixelsHeaderRight) Or (mPreserveNpixelsHeaderRight = 0) Then RaiseEvent HeaderSelected(Button)
        If oldbutton <> Button Then
        Button = 0
        Exit Sub
        End If
    ElseIf myEnabled Then
        RaiseEvent ExposeItemMouseMove(Button, topitem + YYT - 1, CLng(x / scrTwips), CLng(y - (YYT - 1) * myt - mHeadlineHeightTwips) / scrTwips)
    End If
ElseIf myEnabled Then
    RaiseEvent ExposeItemMouseMove(Button, topitem + YYT, CLng(x / scrTwips), CLng((y - YYT * myt) / scrTwips))
End If
If oldbutton <> Button Then Exit Sub
End If
YYT = YYT + (mHeadline <> "")
lastX = x
LastY = y

If (x > Width - barwidth) And BarVisible And EnabledBar And Button = 1 Then
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
cX = x
If Not dr Then lY = x
dr = True

If cY = y \ myt Then Timer3.enabled = False: cY = y \ myt
End If

                  If MarkNext = 4 Then
                  
            RaiseEvent MarkDestroyAny
            End If
there:
End Sub

Private Sub UserControl_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
On Error GoTo there
If dropkey Then Exit Sub
Dim osel As Long
If missMouseClick Then Exit Sub
If Abs(PX - x) <= 60 And Abs(PY - y) <= 60 Then Exit Sub
PX = x
PY = y

RaiseEvent MouseMove(Button, shift, x, y)
If myt = 0 Or Not myEnabled Then Exit Sub
If (Button And 2) = 2 Then Exit Sub
Dim tListcount As Long
tListcount = listcount

If timestamp = 0 Or (timestamp - Timer) > 1 Then timestamp = Timer
If (timestamp + 0.02) > Timer And shift = 0 Then Exit Sub
timestamp = Timer
If Not FreeMouse Then Exit Sub
If Button = 0 Then If Not nopointerchange Then If mousepointer < 2 Then mousepointer = 1
If (x > Width - barwidth) And tListcount > lines + 1 And Not BarVisible Then
    Hidebar = True: BarVisible = m_showbar Or AutoHide Or MultiLineEditBox
ElseIf (x < Width - barwidth) And Button = 0 And BarVisible And (StickBar Or AutoHide) Then
    Hidebar = False
    BarVisible = False
End If
If OurDraw Then
    barMouseMove Button, shift, x, y
    Exit Sub
End If
cX = x
Timer3.enabled = False
Dim YYT As Long, oldbutton As Integer
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
If (Button And 3) > 0 And useFloatList And FloatList Then FloatListMe useFloatList, x, y: Button = 0 Else If Not nopointerchange Then If mousepointer > 1 Then mousepointer = 1
If mHeadline <> "" Then
    If YYT = 0 Then ' we move in mHeadline
    ' -1 is mHeadline
        If (Button And 3) > 0 And FloatList And Not useFloatList Then FloatListMe useFloatList, x, y: Button = 0
            RaiseEvent ExposeItemMouseMove(Button, -1, CLng(x / scrTwips), CLng(y / scrTwips))
        Else
            RaiseEvent ExposeItemMouseMove(Button, topitem + YYT - 1, CLng(x / scrTwips), CLng(y - (YYT - 1) * myt) / scrTwips)
        End If
    Else
        RaiseEvent ExposeItemMouseMove(Button, topitem + YYT, CLng(x / scrTwips), CLng(y - YYT * myt) / scrTwips)
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
        If shift = 0 And ((Not scrollme > 0) And (x > Width / 2) Or Not SingleLineSlide) And StickBar And MarkNext = 0 And tListcount > lines + 1 Then
            If Abs(LastY - y) < scrTwips * 2 Then LastY = y: Exit Sub
            Hidebar = True
            CalcAndShowBar1
            If LastY < y Then
                y = scrTwips * 2
            Else
                y = Scaleheight - scrTwips
            End If
            If Abs(lastX - x) < scrTwips * 4 Or Not MultiLineEditBox Then
                lastX = x
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
        ElseIf y > Scaleheight - myt \ 2 And (tListcount <> 1) Then
            drc = False
            Timer2.enabled = True
        ElseIf YYT >= 0 And YYT <= lines Then
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
            If topitem + YYT < tListcount Then
                If (cX > Scalewidth / 4 And cX < Scalewidth * 3 / 4) And scrollme = 0 Then x = lY
                If Not SELECTEDITEM = topitem + YYT + 1 Then
                    osel = SELECTEDITEM
                    SELECTEDITEM = topitem + YYT + 1
                    If Not BlockItemcount Then
                        REALCUR SELECTEDITEM - 1, cX - scrollme, dummy, mSelstart
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
                        If x - lY > 0 And Not NoPanRight Then
                            scrollme = (x - lY)
                        ElseIf x - lY < 0 And Not NoPanLeft Then
                            scrollme = (x - lY)
                        Else
                            If Not EditFlag Then scrollme = 0
                        End If
                        If Not EditFlag Then If scrollme > 0 Then scrollme = 0
                ElseIf cY <> YYT Then
                    cY = YYT
                    Timer3.enabled = True
                Else
                    If Not Timer1.enabled Then
                        If Not BlockItemcount Then
                            REALCUR SELECTEDITEM - 1, cX - scrollme, dummy, mSelstart
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
                    If x - lY > 0 And Not NoPanRight Then
                        scrollme = (x - lY)
                    ElseIf x - lY < 0 And Not NoPanLeft Then
                        scrollme = (x - lY)
                    Else
                        If Not EditFlag Then scrollme = 0
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

Private Sub UserControl_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
On Error GoTo there
If dropkey Then Exit Sub
If missMouseClick Then missMouseClick = False: Exit Sub
If Button = 1 Then RaiseEvent MouseUp(x / scrTwips, y / scrTwips)
If (Button And 2) = 2 Then
x = nowX
y = nowY
End If
useFloatList = False
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
 If (x < 0 Or y < 0 Or x > .Width Or y > .Height) And (LeaveonChoose And Not BypassLeaveonChoose) Then
If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar
 SELECTEDITEM = -1
 RaiseEvent Selected2(-2)
 Exit Sub
 End If
End With
cX = x
If Hidebar Then Hidebar = False: Redraw Hidebar Or m_showbar
If Timer3.enabled Then
cY = y: DOT3
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
                 If Not EditFlag Then scrollme = 0
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

If YYT >= 0 And YYT <= lines Then


If topitem + YYT < listcount Then

If (Button And 3) > 0 And myEnabled Then

    
    If secreset Then
        ' this is a double click
        secreset = False
         If Not ListSep(topitem + YYT) Then
            If MarkNext = 0 And (EditFlag Or MultiLineEditBox) Then
                If MultiLineEditBox And Not EditFlag Then
                    REALCUR SELECTEDITEM - 1, cX - scrollme, dummy, mSelstart
                    mSelstart = mSelstart + 1
                End If
                MarkWord
            Else
                
        Timer1.enabled = False
        If (((SELECTEDITEM <> (topitem + YYT + 1)) And Not secreset) Or EditFlag) And Not ListSep(topitem + YYT) Then
             SELECTEDITEM = topitem + YYT + 1 ' we have a new selected item
             ' compute selstart always
            If Not BlockItemcount Then
                REALCUR SELECTEDITEM - 1, cX - scrollme, dummy, mSelstart
                mSelstart = mSelstart + 1
                RaiseEvent SetExpandSS(mSelstart)
                RaiseEvent ChangeSelStart(mSelstart)
                
            End If
            RaiseEvent selected(SELECTEDITEM)  ' broadcast
         End If
            If SELECTEDITEM = topitem + YYT + 1 Then
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
        If (((SELECTEDITEM <> (topitem + YYT + 1)) And Not secreset) Or EditFlag) And Not ListSep(topitem + YYT) Then
             SELECTEDITEM = topitem + YYT + 1 ' we have a new selected item
             ' compute selstart always
            If Not BlockItemcount Then
                REALCUR SELECTEDITEM - 1, cX - scrollme, dummy, mSelstart
                mSelstart = mSelstart + 1
                RaiseEvent SetExpandSS(mSelstart)
                RaiseEvent ChangeSelStart(mSelstart)
                
            End If
            RaiseEvent selected(SELECTEDITEM)  ' broadcast
         End If
            If SELECTEDITEM = topitem + YYT + 1 Then
                If MultiSelect Or ListMenu(SELECTEDITEM - 1) Then
                    If (x / scrTwips > 0) And (x / scrTwips < LeftMarginPixels) Then
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
RaiseEvent OutPopUp(x, y, Button)

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

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single)
Dim something$, ok As Boolean
If dropkey Then Exit Sub
 
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

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single, State As Integer)
On Error Resume Next
If dropkey Then Exit Sub
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
 If State = vbOver Then
 
If mHeadline <> "" And y < mHeadlineHeightTwips Then
' we sent twips not pixels
' move...me..??

ElseIf (y - mHeadlineHeightTwips) < myt / 2 And (topitem + YYT > 0) Then
                drc = True
                Timer2.enabled = True
        
        ElseIf y > Scaleheight - myt \ 2 And (tListcount <> 1) Then
                drc = False
                Timer2.enabled = True
        Else
                Timer2.enabled = False
             '  If marvel Then
             
                If Not Timer1.enabled Then
                If Not havefocus Then RaiseEvent DragOverCursor(dragfocus)
                HideCaretOnexit = False: MovePos x, y
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
ElseIf State = vbLeave Then
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
        MovePos x, y
        If CBool(shift And 1) Then chooseshow
    End If
ElseIf State = vbEnter Then
ok = False
If Not marvel Then RaiseEvent DragOverNow(ok)
If Not ok Then
    If Not Timer1.enabled Then
        HideCaretOnexit = False
        MovePos x, y
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
If dropkey Then Exit Sub
If Not DragEnabled Then Exit Sub
Dim aa() As Byte, this$
RaiseEvent DragData(this$)
aa = this$ & ChrW$(0)
 Data.SetData aa(), 13
Data.SetData aa(), vbCFText
AllowedEffects = vbDropEffectCopy + vbDropEffectMove
End Sub
Public Sub MovePos(ByVal x As Single, ByVal y As Single)

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
REALCUR topitem + YYT, x - scrollme, dummy, M_CURSOR
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
backcolor = .ReadProperty("BackColor", m_def_BackColor)
forecolor = .ReadProperty("ForeColor", m_def_ForeColor)
CapColor = .ReadProperty("CapColor", m_def_CapColor)

   Text = .ReadProperty("Text", m_def_Text)

   End With
   If restrictLines > 0 Then
myt = (UserControl.Scaleheight - mHeadlineHeightTwips) / restrictLines
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
myt = (UserControl.Scaleheight - mHeadlineHeightTwips) / restrictLines
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
     .WriteProperty "Backcolor", backcolor, m_def_BackColor
       .WriteProperty "ForeColor", forecolor, m_def_ForeColor
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
Property Let ListIndex(item As Long)
On Error GoTo there
If listcount <= lines + 1 Then
BarVisible = False
Else
Redraw m_showbar And Extender.Visible

End If


If item < listcount Then SELECTEDITEM = item + 1
If SELECTEDITEM > 0 Then
RaiseEvent softSelected(SELECTEDITEM)
Else
SELECTEDITEM = 0
End If
there:
End Property
Public Sub FloatListMe(State As Boolean, x As Single, y As Single)
Static preX As Single, preY As Single
On Error GoTo there
If Extender.Parent Is Nothing Then Exit Sub
If Not State Then
preX = x
preY = y
State = True
mousepointer = 0
doubleclick = 0
Else
If Extender.Visible Then

mousepointer = 5
RaiseEvent NeedDoEvents
If MoveParent Then
If (Extender.Parent.top + (y - preY) < 0) Then preY = y + Extender.Parent.top
If (Extender.Parent.Left + (x - preX) < 0) Then preX = x + Extender.Parent.Left
If ((Extender.Parent.top + y - preY) > FloatLimitTop) And FloatLimitTop > 0 Then preY = Extender.Parent.top + y - FloatLimitTop

If ((Extender.Parent.Left + x - preX) > FloatLimitLeft) And FloatLimitLeft > 0 Then preX = Extender.Parent.Left + x - FloatLimitLeft
Extender.Parent.move Extender.Parent.Left + (x - preX), Extender.Parent.top + (y - preY)
RaiseEvent RefreshDesktop
Else
Extender.ZOrder
If (Extender.top + (y - preY) < 0) Then preY = y + Extender.top
If (Extender.Left + (x - preX) < 0) Then preX = x + Extender.Left
If ((Extender.top + y - preY) > FloatLimitTop) And FloatLimitTop > 0 Then preY = Extender.top + y - FloatLimitTop

If ((Extender.Left + x - preX) > FloatLimitLeft) And FloatLimitLeft > 0 Then preX = Extender.Left + x - FloatLimitLeft

Extender.move Extender.Left + (x - preX), Extender.top + (y - preY)
End If
End If
If Me.BackStyle = 1 Then ShowMe2
End If
there:
End Sub
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
list = mList(item).content
End If
End Property

Property Let list(item As Long, ByVal b$)
On Error GoTo nnn1
If itemcount = 0 Or BlockItemcount Then
If list(item) <> b$ Then
'RaiseEvent PureListOff

RaiseEvent ChangeListItem(item, b$)

End If

Exit Property
End If
If item >= 0 Then
With mList(item)
If Not .content = b$ Then RaiseEvent ChangeListItem(item, b$)
.content = b$
.Line = False
.selected = False
End With
End If
nnn1:
End Property
Property Let menuEnabled(item As Long, ByVal RHS As Boolean)
If item >= 0 Then
With mList(item)
.Line = Not RHS   ' The line flag used as enabled flag, in reverse logic
End With
End If
End Property
Property Let ListSep(item As Long, ByVal RHS As Boolean)
If item >= 0 Then
With mList(item)
.Line = RHS
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
If item >= 0 Then
With mList(item)
ListSep = .Line
End With
End If
End Property

Property Let ListSelected(item As Long, ByVal b As Boolean)
Dim first As Long, last As Long
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then

If mList(item).radiobutton Then
        ' erase all
        first = item
        While first > 0 And mList(first).radiobutton
        first = first - 1
        Wend
        If Not mList(first).radiobutton Then first = first + 1
        last = item
        While last < listcount - 1 And mList(last).radiobutton
        last = last + 1
        Wend
        If Not mList(last).radiobutton Then last = last - 1
        For first = first To last
        mList(first).selected = False
        Next first
End If
With mList(item)
.selected = b
End With
End If
End If
End Property
Property Let ListSelectedNoRadioCare(item As Long, ByVal b As Boolean)
Dim first As Long, last As Long
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then

With mList(item)
.selected = b
End With
End If
End If
End Property
Property Get ListSelected(item As Long) As Boolean

If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mList(item)
ListSelected = .selected
End With
End If
End If
End Property
Property Let ListRadio(item As Long, ByVal b As Boolean)
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mList(item)
.radiobutton = b
End With
End If
End If
End Property
Property Get ListRadio(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mList(item)
ListRadio = .radiobutton
End With
End If
End If
End Property
Property Get ListMenu(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mList(item)
ListMenu = .radiobutton Or .checked
End With
End If
End If
End Property
Property Let ListChecked(item As Long, ByVal b As Boolean)
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mList(item)
.checked = b
End With
End If
End If
End Property
Property Get ListChecked(item As Long) As Boolean
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 Then
With mList(item)
ListChecked = .checked
End With
End If
End If
End Property
Public Sub moveto(ByVal Key As String)
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
For i = 0 To listcount - 1
If list(i) Like Key Then Exit For
Next i
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

If listcount <= lines + 1 Then
    BarVisible = False
Else
    BarVisible = m_showbar And Extender.Visible

End If
If item > 0 And item <= listcount Then
    If MultiLineEditBox Then FindRealCursor item
    If item - topitem > 0 And item - topitem <= lines + 1 Then
        If restrictLines > 0 Then
            If listcount <= topitem + lines Then
                topitem = listcount - lines - 1
                If topitem < 0 Then topitem = 0
            End If
        End If
    
            SELECTEDITEM = item
           
            If SELECTEDITEM = listcount Then
            State = True
            Value = Max
            State = False
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
If hWnd <> 0 Then HideCaret (hWnd)
State = True
mValue = 0  ' if here you have an error you forget to apply VALUE as default property
showshapes
LastVScroll = 1

'max = 0
State = False
topitem = 0
Buffer = 100
ReDim mList(0 To Buffer)

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
.content = a$
.Line = False
.selected = False
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
.content = a$
.Line = False
.selected = False
End With

Timer1.enabled = False
Timer1.Interval = 100
Timer1.enabled = True
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
.content = a$
.Line = False
.selected = False
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
Dim REALX As Long, RealX2 As Long, myt1, oldtopitem As Long
On Error GoTo there
 If SuspDraw Then Exit Sub
If visibleme Then
 barwidth = UserControlTextWidth("W")
 CalcAndShowBar1
Timer1.enabled = True: Exit Sub
End If
If listcount = 0 And HeadLine = vbNullString Then
    Repaint
    Exit Sub
End If

Dim i As Long, j As Long, g$, nr As RECT, fg As Long, hnr As RECT, skipme As Boolean, nfg As Long
If MultiSelect And LeftMarginPixels < mytPixels Then LeftMarginPixels = mytPixels
If Not headeronly Then Repaint
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


If mHeadlineHeight <> hnr.Bottom Then
HeadlineHeight = hnr.Bottom
nr.Bottom = mHeadlineHeight
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
If SELECTEDITEM > 0 Then
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
        State = True
            On Error Resume Next
            Err.Clear
    If Not Spinner Then
            If listcount - 1 - lines < 1 Then
            Max = 1
            Else
            Max = listcount - 1 - lines
            End If
            If Err.Number > 0 Then
                Value = listcount - 1
                Max = listcount - 1
            End If
                      Value = j

        End If
        State = False
    
    End If
   
Else
    State = True
        On Error Resume Next
        Err.Clear
        If Not Spinner Then
        Max = listcount - 1
        If Err.Number > 0 Then
            Value = listcount - 1
            Max = listcount - 1
        End If
        End If
    State = False
    
End If
  
    '
j = topitem + lines
If j >= listcount Then j = listcount - 1
'Text1 = VbNullString


    If listcount = 0 Then
        ' DO NOTHING
    Else

       currentX = scrollme

  DrawStyle = vbSolid
  fg = Me.forecolor
  
  If havefocus Or dragfocus Then
  caretCreated = False
  DestroyCaret
 
  End If
  
  
  Dim onemore As Long
    If restrictLines = 0 Then
        If nr.top + (topitem - j + 1) * mytPixels < HeightPixels Then
            onemore = 1
        End If
    End If
        For i = topitem To j + onemore
        
        RaiseEvent ExposeRect(i, VarPtr(nr), UserControl.hDC, skipme)
        If Not skipme Then
            If i = SELECTEDITEM - 1 And Not NoCaretShow And Not ListSep(i) Then
                nr.Left = scrollme / scrTwips + LeftMarginPixels
                nfg = fg
                RaiseEvent SpecialColor(nfg)
                If nfg <> fg Then Me.forecolor = nfg
                If mEditFlag Then
                    If nfg = fg Then Me.forecolor = fg
                ElseIf nfg = fg Then
                    If Me.backcolor = 0 Then
                        Me.forecolor = &HFFFFFF
                    Else
                        Me.forecolor = 0
                    End If
                End If
            
                    If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                                   MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i, True
                        End If

                 PrintLineControlSingle UserControl.hDC, list(i), nr
                 Me.forecolor = fg
             Else
                 If ListSep(i) And list(i) = vbNullString Then
                   hnr.Left = 0
                   hnr.Right = nr.Right
                   hnr.top = nr.top + mytPixels \ 2
                   hnr.Bottom = hnr.top + 1
                   FillBack UserControl.hDC, hnr, forecolor
                Else
                   If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
                                 MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i
                    End If
                    If ListSep(i) Then
                        forecolor = dcolor
                    Else
         
               '    ForeColor = fg
                    End If
                    PrintLineControlSingle UserControl.hDC, list(i), nr
                    
                End If
      
             End If
                     If SingleLineSlide Then
nr.Left = LeftMarginPixels
Else
nr.Left = scrollme / scrTwips + LeftMarginPixels
End If
        End If
     nr.top = nr.top + mytPixels
nr.Bottom = nr.top + mytPixels + 1
 forecolor = fg
    Next i
  
 DrawMode = vbInvert

 myt1 = myt - scrTwips
    If SELECTEDITEM > 0 Then
        If SELECTEDITEM - topitem - 1 <= lines + onemore And Not ListSep(SELECTEDITEM - 1) Then
                If Not NoCaretShow Then
                                If EditFlag Then 'And Not BlockItemcount Then
                                If SelStart = 0 Then SelStart = 1
                                        DrawStyle = vbSolid
                                 If CenterText Then
                                        ' (UserControl.ScaleWidth- LeftMarginPixels * scrTwips-UserControlTextWidth(list$(selecteitem-1)))/2
                                          REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips + (UserControl.Scalewidth - LeftMarginPixels * scrTwips - UserControlTextWidth(list$(SELECTEDITEM - 1))) / 2
                                          RealX2 = scrollme / 2 + REALX
                                            Else
                                   REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
                                    RealX2 = scrollme + REALX
                                   End If
                                   If InternalCursor Or dragfocus Then
                                   ' not used
                                              If Noflashingcaret Or Not havefocus Then
                                                    Line (scrollme + REALX, (SELECTEDITEM - topitem - 1) * myt + myt1 + mHeadlineHeightTwips)-(RealX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips), forecolor
                
                                                Else
                                                       ShowMyCaretInTwips RealX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips
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
                                         DrawStyle = vbInvisible
                                
                                        If BackStyle = 1 Then
                            
                                            Line (scrTwips, (SELECTEDITEM - topitem) * myt + mHeadlineHeightTwips)-(scrollme + UserControl.Width - 2 * scrTwips, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips), 0, B
                                        Else
                                            Line (0, (SELECTEDITEM - topitem) * myt + mHeadlineHeightTwips)-(scrollme + UserControl.Width, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips), 0, B
                                     
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
RepaintScrollBar
RaiseEvent ScrollMove(topitem)
there:
End Sub
Public Sub ShowMe2()
On Error GoTo there
If SuspDraw Then Exit Sub
Dim YYT As Long, nr As RECT, j As Long, i As Long, skipme As Boolean, fg As Long, hnr As RECT, nfg As Long, nfg1 As Long
 Dim REALX As Long, RealX2 As Long, myt1
barwidth = UserControlTextWidth("W")
If listcount = 0 And HeadLine = vbNullString Then
Repaint
HideCaret (hWnd)
Exit Sub
End If
If MultiSelect And LeftMarginPixels < mytPixels Then LeftMarginPixels = mytPixels
Repaint

YYT = myt
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

If mHeadlineHeight <> hnr.Bottom Then
HeadlineHeight = hnr.Bottom
nr.Bottom = mHeadlineHeight
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
j = topitem + lines
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
fg = Me.forecolor
nfg = fg
nfg1 = fg
  RaiseEvent SpecialColor(nfg1)
Dim onemore As Long
If restrictLines = 0 Then
    If nr.top + (topitem - j + 1) * mytPixels < HeightPixels Then
        onemore = 1
    End If
End If
For i = topitem To j + onemore
currentX = scrollme
currentY = 0
  RaiseEvent ExposeRect(i, VarPtr(nr), UserControl.hDC, skipme)
  If Not skipme Then
  If i = SELECTEDITEM - 1 And Not NoCaretShow And Not ListSep(i) Then
    nfg = fg
  'RaiseEvent SpecialColor(nfg)
  If nfg1 <> nfg Then nfg = nfg1
  If nfg <> fg Then Me.forecolor = nfg
  nr.Left = scrollme / scrTwips + LeftMarginPixels
  If mEditFlag Then
   If nfg = fg Then Me.forecolor = fg
  ElseIf nfg = fg Then
  If Me.backcolor = 0 Then
  Me.forecolor = &HFFFFFF
  Else
  Me.forecolor = 0
  End If
  End If

   If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
 MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i, True
 End If

   PrintLineControlSingle UserControl.hDC, list(i), nr
 If nfg = fg Then Me.forecolor = fg
 Else
    nfg = fg
  'RaiseEvent SpecialColor(nfg)
    If nfg1 <> nfg Then nfg = nfg1
 If ListSep(i) And list(i) = vbNullString Then
 hnr.Left = 0
 hnr.Right = nr.Right
 hnr.top = nr.top + mytPixels \ 2
 hnr.Bottom = hnr.top + 1
 FillBack UserControl.hDC, hnr, forecolor
 Else

 If (MultiSelect Or ListMenu(i)) And itemcount > 0 Then
 MyMark UserControl.hDC, mytPixels \ 3, nr.Left - LeftMarginPixels / 2, nr.top + mytPixels / 2, i
 End If
  If ListSep(i) Then
 forecolor = dcolor
 Else
 If nfg = fg Then forecolor = fg
   If SELECTEDITEM - 1 = i And nfg <> fg Then
   Me.forecolor = nfg  'uintnew(&HFFFFFF) - uintnew(nfg)
   End If
 End If

 PrintLineControlSingle UserControl.hDC, list(i), nr
 End If

   End If
 If SingleLineSlide Then
nr.Left = LeftMarginPixels
Else
nr.Left = scrollme / scrTwips + LeftMarginPixels
End If
 
  End If
 
nr.top = nr.top + mytPixels
nr.Bottom = nr.top + mytPixels + 1
forecolor = fg
Next i

 myt1 = myt - scrTwips

' DrawStyle = vbInvisible
DrawMode = vbInvert
If SELECTEDITEM > 0 Then

    If SELECTEDITEM - topitem - 1 <= lines + onemore And SELECTEDITEM > topitem And Not ListSep(SELECTEDITEM - 1) Then
       '' cY = yyt * (i - topitem + 1) 'CurrentY
        
        If Not NoCaretShow Then
                 If EditFlag Then ' And Not BlockItemcount Then
                    If SelStart = 0 Then SelStart = 1
                                             DrawStyle = vbSolid
                                             RaiseEvent PureListOff
                                          If CenterText Then
                                        ' (UserControl.ScaleWidth- LeftMarginPixels * scrTwips-UserControlTextWidth(list$(selecteitem-1)))/2
                                          REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips + (UserControl.Scalewidth - LeftMarginPixels * scrTwips - UserControlTextWidth(list$(SELECTEDITEM - 1))) / 2
                                          RealX2 = scrollme / 2 + REALX
                                            Else
                                   REALX = UserControlTextWidth(Mid$(list(SELECTEDITEM - 1), 1, SelStart - 1)) + LeftMarginPixels * scrTwips
                                    RealX2 = scrollme + REALX
                                   End If
                                   RaiseEvent PureListOn
                           
                                  If Noflashingcaret Or Not havefocus Then
                                 
                            If InternalCursor Or dragfocus Then Line (scrollme + REALX, (SELECTEDITEM - topitem - 1) * myt + myt1 + mHeadlineHeightTwips)-(RealX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips), forecolor
                      
                                Else
                           
                                If InternalCursor Or Not MultiLineEditBox Then
                                ShowMyCaretInTwips RealX2, (SELECTEDITEM - topitem - 1) * myt + mHeadlineHeightTwips
                                End If
                                   End If
                               
                                   If Not NoScroll Then If RealX2 > Width * 0.8 * dragslow Then scrollme = scrollme - Width * 0.2 * dragslow: PrepareToShow 10
                                   If RealX2 - Width * 0.2 * dragslow < 0 Then
                                    If Not NoScroll Then
                                     scrollme = scrollme + Width * 0.2 * dragslow
                                   If scrollme > 0 Then scrollme = 0 Else PrepareToShow 10
:
                                   
                                   End If
                                   End If
                           Else
                                   DrawStyle = vbInvisible
                                   
                                   If BackStyle = 1 Then
                       
                                           Line (scrTwips, (SELECTEDITEM - topitem) * YYT + mHeadlineHeightTwips)-(0 + UserControl.Width, (SELECTEDITEM - topitem - 1) * YYT + mHeadlineHeightTwips - scrTwips / 2), 0, B
                         
                                   Else
                       
                                        Line (0, (SELECTEDITEM - topitem) * YYT + mHeadlineHeightTwips)-(0 + UserControl.Width, (SELECTEDITEM - topitem - 1) * YYT + mHeadlineHeightTwips), 0, B
                           
                                   End If
      
                End If
            Else
        RaiseEvent PureListOff
        
        End If
        Else
        
        HideCaret (hWnd)
    End If
Else


End If

 DrawStyle = vbSolid
DrawMode = vbCopyPen
currentY = 0
currentX = 0
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
myt = (UserControl.Scaleheight - mHeadlineHeightTwips - 1) / restrictLines

Else
l = Int((UserControl.Scaleheight - mHeadlineHeightTwips) / myt) - 1
End If
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
ex1:
If l <= 0 Then
l = 0
End If

lines = l
End Property


Private Sub LargeBar1_Change()

If Not State Then



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
myt = (UserControl.Scaleheight - mHeadlineHeightTwips) \ restrictLines
Else
myt = UserControlTextHeight() + addpixels * scrTwips
End If
'HeadlineHeight = UserControlTextHeight() / SCRTWIPS
mytPixels = myt / scrTwips
myt = mytPixels * scrTwips
On Error GoTo th1

If Extender.Parent Is Nothing Then Exit Sub

If Extender.Parent.Picture.handle <> 0 And BackStyle = 1 Then

If Me.BorderStyle = 1 Then
currentY = 0
    currentX = 0
Line (0, 0)-(Scalewidth - scrTwips, Scaleheight - scrTwips), Me.backcolor, B
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
If mmo.Picture.handle <> 0 Then
    UserControl.PaintPicture mmo.Picture, 0, 0, , , Extender.Left, Extender.top
    If Me.BorderStyle = 1 Then
    currentY = 0
        currentX = 0
    Line (0, 0)-(Scalewidth - scrTwips, Scaleheight - scrTwips), Me.backcolor, B
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
If restrictLines > 0 Then
myt = (UserControl.Scaleheight - mHeadlineHeightTwips) \ restrictLines
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
        If UserControl.Parent.Picture.handle <> 0 Then
            If Me.BorderStyle = 1 Then
                    currentY = 0
                    currentX = 0
                    Line (0, 0)-(Scalewidth - scrTwips, Scaleheight - scrTwips), Me.backcolor, B
                    UserControl.PaintPicture UserControl.Parent.Picture, scrTwips, scrTwips, Width - 2 * scrTwips, Height - 2 * scrTwips, Extender.Left, Extender.top, Width - 2 * scrTwips, Height - 2 * scrTwips
                    currentY = 0
                    currentX = 0
            Else
                    UserControl.PaintPicture UserControl.Parent.Picture, 0, 0, , , Extender.Left, Extender.top
            End If
            Else
            If Me.BorderStyle = 1 Then
                currentY = 0
                currentX = 0
                Line (0, 0)-(Scalewidth - scrTwips, Scaleheight - scrTwips), Me.backcolor, B
                UserControl.PaintPicture UserControl.Parent.Image, scrTwips, scrTwips, Width - 2 * scrTwips, Height - 2 * scrTwips, Extender.Left, Extender.top, Width - 2 * scrTwips, Height - 2 * scrTwips
                currentY = 0
                currentX = 0
            Else
                UserControl.PaintPicture UserControl.Parent.Image, 0, 0, , , Extender.Left, Extender.top
            End If
        End If
    Else
        Dim mmo As Object, isfrm As Boolean
        RaiseEvent GetBackPicture(mmo)
        If Not mmo Is Nothing Then
            If mmo.Image.handle <> 0 Then
                UserControl.PaintPicture mmo.Image, 0, 0, , , Extender.Left - mmo.Left, Extender.top - mmo.top
                If Me.BorderStyle = 1 Then
                currentY = 0
                    currentX = 0
                Line (0, 0)-(Scalewidth - scrTwips, Scaleheight - scrTwips), Me.backcolor, B
                    currentY = 0
                    currentX = 0
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
If UserControl.Ambient.UserMode = False Then
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
If Not State Then
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
Private Sub PrintLineControlSingle(mHdc As Long, ByVal c As String, R As RECT)
' this is our basic print routine
Dim that As Long, cc As String, fg As Long
If CenterText Then that = DT_CENTER
If VerticalCenterText Then that = that Or DT_VCENTER
If WrapText Then
c = c + space(4)  ' 4 additional characters for DT_MODIFYSTRING
DrawTextEx mHdc, StrPtr(c), -1, R, DT_WORDBREAK Or DT_NOPREFIX Or DT_MODIFYSTRING Or DT_PATH_ELLIPSIS Or that, VarPtr(tParam)
Else
If LastLinePart <> "" Then
    If FadeLastLinePart > 0 Then
    cc = c + LastLinePart
    fg = Me.forecolor
    Me.forecolor = FadeLastLinePart
   DrawText mHdc, StrPtr(cc), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or that Or DT_EXPANDTABS
   Me.forecolor = fg
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
R.Right = dd.Scalewidth
R.Bottom = dd.Scaleheight
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

If parentpic.handle <> 0 Then
UserControl.PaintPicture parentpic, 0, 0, , , Myleft, mytop
Else
th1:
'UserControl.Cls
End If
End Sub
Private Sub Redraw(ParamArray status())
If EnabledBar Then
Dim fakeLargeChange As Long, newheight As Long, newTop As Long
Dim b As Boolean, nstatus As Boolean
Timer2bar.enabled = False
If UBound(status) >= 0 Then
nstatus = CBool(status(0))
Else
nstatus = Shape1.Visible
End If
With UserControl
If mHeadline <> "" Then
newheight = .Height - mHeadlineHeightTwips
newTop = mHeadlineHeightTwips
Else
newheight = .Height
End If

If newheight <= 0 Then

Else
        minimumWidth = (1 - (Max - Min) / (largechange + Max - Min)) * newheight * (1 - Percent * 2) + 1
        If minimumWidth < 60 Then
        
        mLargeChange = Round(-(Max - Min) / ((60 - 1) / newheight / (1 - Percent * 2) - 1) - Max + Min) + 1
        
        minimumWidth = (1 - (Max - Min) / (largechange + Max - Min)) * newheight * (1 - Percent * 2) + 1
        End If
        valuepoint = (Value - Min) / (largechange + Max - Min) * (newheight * (1 - 2 * Percent)) + newheight * Percent

       Shape Shape1, Width - barwidth, newTop + valuepoint, barwidth, minimumWidth
       Shape Shape2, Width - barwidth, newTop + newheight * (1 - Percent), barwidth, newheight * Percent ' newtop + newheight * Percent - scrTwips
        Shape Shape3, Width - barwidth, newTop, barwidth, newheight * Percent   ' left or top
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
Min = mmin
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
If mWidth < 100 Then
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
If State And Spinner Then
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
UserControl.mousepointer = oldpointer
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
UserControl.mousepointer = 1
Else
UserControl.mousepointer = oldpointer
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
    ChooseNextLeft Me, Me.Parent, True
End Sub
Friend Sub GoON()
    ChooseNextRight Me, Me.Parent, True
End Sub
Friend Sub TakeKey(KeyCode As Integer, shift As Integer)
    UserControl_KeyDown KeyCode, shift
If KeyCode <> 0 Then
    UserControl_KeyUp KeyCode, shift
End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, shift As Integer)
PrevLocale = GetLocale()
If BypassKey Then KeyCode = 0: shift = 0: Exit Sub
lastshift = shift

If KeyCode = 27 And NoEscapeKey Then
KeyCode = 0
Exit Sub
End If
If Arrows2Tab And Not mEditFlag Then
    If KeyCode = vbKeyLeft Or (KeyCode = vbKeyUp And Arrows2Tab) Then
        ChooseNextLeft Me, Me.Parent
        KeyCode = 0
        Exit Sub
    ElseIf KeyCode = vbKeyRight Or (KeyCode = vbKeyDown And Arrows2Tab) Then
        ChooseNextRight Me, Me.Parent
        KeyCode = 0
        Exit Sub
    End If
End If
If KeyCode = vbKeyTab And Not mEditFlag Then
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
If dropkey Then shift = 0: KeyCode = 0: Exit Sub
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
If (KeyCode = 0) Or Not (enabled Or State) Then Exit Sub
If SELECTEDITEM < 0 Then
SELECTEDITEM = topitem + 1: ShowMe2
If Not EditFlag Then: KeyCode = 0
End If
LargeBar1KeyDown KeyCode, shift
If EnabledBar Then
Select Case KeyCode
Case vbKeyLeft, vbKeyUp
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
Private Function processXY(ByVal x As Single, ByVal y As Single, Optional rep As Boolean = True) As Boolean
If NoBarClick Then Exit Function
Timer1bar.enabled = False
Dim checknewvalue As Long, newheight As Long
With UserControl
If mHeadline <> "" Then
newheight = .Height - mHeadlineHeightTwips
y = y - mHeadlineHeightTwips
Else
newheight = .Height
End If

If minimumWidth < 60 Then minimumWidth = 60  ' 4 x scrtwips
' value must have real max ...minimum MAX-60
If Vertical Then
' here minimumwidth is minimumheight
If y >= valuepoint - scrTwips And y <= minimumWidth + valuepoint - scrTwips Then
' is our scroll bar
OurDraw = Not rep

ElseIf y > newheight * Percent And y < newheight * (1 - Percent) Then
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
            If y < newheight * Percent Then y = newheight * Percent
            If y > Round(newheight * (1 - Percent)) - minimumWidth + newheight * Percent Then
            y = newheight * (1 - Percent) - minimumWidth
            End If
            checknewvalue = Round((y - newheight * Percent) * (Max - Min) / ((newheight * (1 - Percent) - minimumWidth) - newheight * Percent)) + Min
            If checknewvalue = Value And mjumptothemousemode Then
                 ' do nothing
                
            Else
    
                Value = checknewvalue
                If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5 ' autorepeat
                Timer1bar.enabled = True
            End If
ElseIf y >= newheight * (1 - Percent) And y <= newheight Then ' is right button
processXY = True
checknewvalue = Value + mSmallChange
If checknewvalue = Value Then
' do nothing
Else
Value = checknewvalue
If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5  '
Timer1bar.enabled = True
End If

ElseIf y < newheight * Percent - scrTwips Then
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
If x >= valuepoint - scrTwips And x <= minimumWidth + valuepoint - scrTwips Then
' is our scroll bar
OurDraw = Not rep

ElseIf x > .Width * Percent And x < .Width * (1 - Percent) Then
processXY = True
'  we are inside so take a largechange
        If x < valuepoint Then
                If mjumptothemousemode Then
                  x = (x \ minimumWidth + 1) * minimumWidth - minimumWidth
                Else
                x = valuepoint - minimumWidth
                End If
        Else
                If mjumptothemousemode Then
                x = (x \ minimumWidth - 1) * minimumWidth + minimumWidth
                Else
                x = valuepoint + minimumWidth
                End If
        End If
            If x < .Width * Percent Then x = .Width * Percent
            If x > Round(.Width * (1 - Percent)) - minimumWidth + .Width * Percent Then
            x = .Width * (1 - Percent) - minimumWidth
            End If
            checknewvalue = Round((x - .Width * Percent) * (Max - Min) / ((.Width * (1 - Percent) - minimumWidth) - .Width * Percent)) + Min
            If checknewvalue = Value And mjumptothemousemode Then
            ' do nothing
            Else
            Value = checknewvalue
            If Timer1bar.Interval > 10 Then Timer1bar.Interval = Timer1bar.Interval - 5  ' autorepeat
            Timer1bar.enabled = True
            End If
ElseIf x >= .Width * (1 - Percent) And x <= .Width Then
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
ElseIf x < .Width * Percent - scrTwips Then
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
Private Sub barMouseMove(Button As Integer, shift As Integer, x As Single, ByVal y As Single)
If Not EnabledBar Then Exit Sub
Dim ForValidValue As Long, newheight As Long
If OurDraw Then
If Button = 1 Then
Timer1bar.Interval = 5000
'timer2bar.enabled = False
If minimumWidth < 60 Then minimumWidth = 60  ' 4 x scrtwips

With UserControl


If Vertical Then

If mHeadline <> "" Then
y = y - mHeadlineHeightTwips
newheight = .Height - mHeadlineHeightTwips
Else
newheight = .Height
End If
        ForValidValue = y + GetOpenValue 'ForValidValue + valuepoint
        If ForValidValue < newheight * Percent Then
        ForValidValue = newheight * Percent
        Value = Min
        ElseIf ForValidValue > ((newheight * (1 - Percent) - minimumWidth)) Then
        ForValidValue = ((newheight * (1 - Percent) - minimumWidth))
        Value = Max
        Else

         Value = Round((ForValidValue - newheight * Percent) * (Max - Min) / ((newheight * (1 - Percent) - minimumWidth) - newheight * Percent)) + Min
         
        End If
    

Else

         ForValidValue = x + GetOpenValue
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
mList(item).checked = checked ' means that can be checked
mList(item).contentID = id$
ListSelected(item) = firstState
mList(item).radiobutton = radiobutton ' one of the group can be checked
End If
End If
End Sub
Public Function GetMenuId(id$, Pos As Long) As Boolean
' return item number with that id$
' work only in the internal list
Dim i As Long
If itemcount > 0 And Not BlockItemcount Then
For i = 0 To itemcount - 1
If mList(i).contentID = id$ Then Pos = i: Exit For
Next i
End If
GetMenuId = Not (i = itemcount)
End Function
Property Get id(item As Long) As String
If itemcount > 0 And Not BlockItemcount Then
If item >= 0 And item < listcount Then
id = mList(item).contentID
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
Private Sub MyMark(thathDC As Long, Radius As Long, x As Long, y As Long, item As Long, Optional Reverse As Boolean = False) ' circle
'
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
Dim th As RECT
th.Left = x - Radius
th.top = y - Radius
th.Right = x + Radius
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
             
        th.Left = x - Radius
        th.top = y - Radius
        th.Right = x + Radius
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
            my_brush = CreateSolidBrush(m_backcolor)
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
Dim n As Long, st As Long, st1 As Long, st0 As Long, w As Integer, mark1 As Long, mark2 As Long
Dim Pad$, addlength As Long, s$
s$ = list(where)
RaiseEvent RealCurReplace(s$)

n = UserControlTextWidth(s$)

If CenterText Then
probeX = scrollme / 2 + probeX - LeftMarginPixels * scrTwips - (UserControl.Scalewidth - LeftMarginPixels * scrTwips - n) / 2 + 2 * scrTwips
Else
probeX = probeX - LeftMarginPixels * scrTwips + 2 * scrTwips
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
w = AscW(Mid$(s$, 1, st1))
If w > -10241 And w < -9984 Then
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
w = UserControlTextWidth2(s$, mark2)
Pad$ = Mid$(s$, mark1, mark2 - mark1 + 1)
n = Len(Pad$)
st1 = 0
Do
st1 = st1 + 1
realpos = w - UserControlTextWidth2(Pad$, st1 + 1)

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
If item < listcount Then
SELECTEDITEM = item + 1
Else
SELECTEDITEM = 0
End If
End Property
Public Sub ListindexPrivateUseFirstFree(item As Long)
Dim x As Long
If item < listcount Then

For x = item To listcount - 1
If Not mList(x).Line Then SELECTEDITEM = x + 1: Exit For
Next x
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

Public Property Let EditFlag(ByVal RHS As Boolean)
mEditFlag = RHS
If Not RHS Then If hWnd <> 0 Then DestroyCaret: caretCreated = False
End Property
Public Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long, Optional ByVal offsetx As Long = 0)
Dim a As RECT
CopyFromLParamToRect a, thatRect
a.Bottom = a.Bottom - 1
a.Left = a.Left + offsetx
FillBack thathDC, a, thatbgcolor
End Sub
Public Sub WriteThere(thatRect As Long, aa$, ByVal offsetx As Long, ByVal offsety As Long, thiscolor As Long)
Dim a As RECT, fg As Long
CopyFromLParamToRect a, thatRect
If a.Left > Width Then Exit Sub
a.Right = WidthPixels
a.Left = a.Left + offsetx
a.top = a.top + offsety
fg = forecolor
forecolor = thiscolor
    DrawText UserControl.hDC, StrPtr(aa$), -1, a, DT_NOPREFIX Or DT_NOCLIP
    forecolor = fg
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
Dim thatwidth As Long, c$, dummy1 As Long
If SelStart < 1 Then
c$ = list(ListIndex)
Else
c$ = Left$(list(ListIndex), SelStart - 1)
End If

thatwidth = UserControlTextWidth(c$) + LeftMarginPixels * scrTwips
''If mSelstart > Len(c$) Then Exit Sub
Dim where As Long, oldmselstart As Long
oldmselstart = mSelstart
REALCUR tothere, thatwidth, (dummy1), where
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
If Pos <= 0 Then Pos = 1  ': Stop
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

Public Property Get mousepointer() As Integer
mousepointer = UserControl.mousepointer
End Property

Public Property Let mousepointer(ByVal RHS As Integer)
UserControl.mousepointer = RHS
End Property
Public Property Set mouseicon(RHS)
Set UserControl.mouseicon = RHS
End Property
Function GetLocale() As Long
    Dim R&
      R = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      GetLocale = val("&H" & Right(Hex(R), 4))
   
    
End Function
Function GetKeY(ascii As Integer) As String
    Dim Buffer As String, ret As Long

    Buffer = String$(514, 0)
    Dim R&
      R = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
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
Function DoubleClickArea(ByVal x As Long, ByVal y As Long, ByVal Xorigin As Long, ByVal Yorigin As Long, setupxy As Long) As Boolean
   If Abs(x - Xorigin) < setupxy And Abs(y - Yorigin) < setupxy Then
        preservedoubleclick = doubleclick
    Else
        preservedoubleclick = 0
   End If
   DoubleClickArea = Not preservedoubleclick = 0
End Function
Function DoubleClickCheck(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long, ByVal Xorigin As Long, ByVal Yorigin As Long, setupxy As Long, itemline As Long) As Boolean
' doubleclick
If item = itemline Then
   If Abs(x - Xorigin) < setupxy And Abs(y - Yorigin) < setupxy Then
        If Not nopointerchange Then mousepointer = 1
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
Function SingleClickCheck(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long, ByVal Xorigin As Long, ByVal Yorigin As Long, setupxy As Long, itemline As Long) As Boolean
If item = itemline Then
   If Abs(x - Xorigin) < setupxy And Abs(y - Yorigin) < setupxy Then
      If Not nopointerchange Then mousepointer = 1
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
DeleteObject hRgn
End If
End Sub
Public Sub ShowMenu()
    'dropkey = True
    RaiseEvent DeployMenu

   
End Sub
Public Sub CascadeSelect(ByVal item As Long)
RaiseEvent CascadeSelect(item)
End Sub
Public Property Let BlinkTime(t As Variant)
BlinkON = True <> 0
mBlinkTime = t
Timer1.Interval = t
Timer1.enabled = True
End Property
Public Property Get BlinkTime() As Variant
BlinkTime = mBlinkTime
End Property
Sub DestCaret()
 DestroyCaret
 caretCreated = False
End Sub
Private Function MyTrimL(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL = 1: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 9
    Case Else
     MyTrimL = (i - p2) \ 2 + 1
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
    UserControl.mousepointer = CInt(RHS)
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
                         Set UserControl.mouseicon = Form1.Picture2.mouseicon: UserControl.mousepointer = 99
                 
                        If aPic Is Nothing Then MissCdib: Exit Property
                        Set UserControl.mouseicon = aPic
                        UserControl.mousepointer = 99
                    Else
                    MissFile
                    End If
                Else
                 Set UserControl.mouseicon = Form1.Picture2.mouseicon: UserControl.mousepointer = 99
                    
                    
                End If
End If
oldpointer = UserControl.mousepointer
End Property
Friend Sub SetIcon(RHS, mpointer As Integer)
On Error Resume Next
nopointerchange = True
Set UserControl.mouseicon = RHS
oldpointer = mpointer
UserControl.mousepointer = mpointer
End Sub
Public Property Get icon()
Set icon = UserControl.mouseicon
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
