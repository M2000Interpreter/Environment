VERSION 5.00
Begin VB.Form GuiM2000 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   4620
   ClientLeft      =   3000
   ClientTop       =   3000
   ClientWidth     =   9210
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "GuiM2000.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame ResizeMark 
      Appearance      =   0  'Flat
      BackColor       =   &H003B3B3B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   8475
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   873
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   3881787
      ForeColor       =   16777215
      CapColor        =   16777215
   End
End
Attribute VB_Name = "GuiM2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal Hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" _
        (ByVal Hdc As Long, _
        ByVal hBrush As Long, _
        ByVal lpDrawStateProc As Long, _
        ByVal lParam As Long, _
        ByVal wParam As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal fuFlags As Long) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
    Private Const GWL_WNDPROC = -4
    Private m_Caption As String
Dim setupxy As Single
Dim Lx As Single, lY As Single, dr As Boolean
Dim scrTwips As Long
Dim bordertop As Long, borderleft As Long
Dim allwidth As Long, itemWidth As Long
Private ExpandWidth As Boolean, lastfactor As Single
Private myEvent As mEvent
Private ttl As Boolean, drawminimized As Boolean
Private GuiControls As New Collection
Dim onetime As Boolean, PopupOn As Boolean
Dim alfa As New GuiButton
Public prive As Long
Private ByPassEvent As Boolean, mBarColor As Long, mIconColor As Long
Private mIndex As Long
Private mSizable As Boolean, mNoTaskBar As Boolean
Public Relax As Boolean
Private MarkSize As Long
Public MY_BACK As cDIBSection
Dim CtrlFont As New StdFont
Dim novisible As Boolean
Private mModalid As Double, mModalIdPrev As Double
Private mPopUpMenu As Boolean, mNoCaption As Boolean
Public IamPopUp As Boolean
Private mEnabled As Boolean
Public WithEvents mDoc As Document
Attribute mDoc.VB_VarHelpID = -1
Public VisibleOldState As Boolean
Private minimPos As Long
Private mQuit, mMenuWidth As Long
Private MyForm3 As Form3, mMyName$
Private IhaveLastPos As Boolean, MeTop As Long, MeLeft As Long, MeWidth As Long, MeHeight As Long
Private mShowMaximize As Boolean, EnableStandardInfo As Boolean
Private infopos As Long, NoEventInfo As Boolean
Public WithEvents Pad As Form
Attribute Pad.VB_VarHelpID = -1
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Public UseReverse As Boolean, UseInfo As Boolean
Private moveMe As Boolean, movemeX As Single, movemeY As Single, mTimes, mIcon As Boolean
Private sizeMe As Boolean
Private lastBlink As Long, LastBlinkmTimes As Boolean, lastBlinkOn As Boolean, Stored As Boolean
Private DefaultName As String, ByPassColor As Boolean
Public LastActive As String
Public RefreshList As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal BLen As Long) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC
Private Targets As Boolean, q() As target
Public NoHook As Boolean, SkipFirstClick As Boolean
Public SkipAutoPos As Boolean, previewKey
Dim lastitem As Long, safeform As LongHash
Private PRmodulename$, acclist As FastCollection, mHover As String
Friend Property Get RealHover() As String
RealHover = mHover
End Property
Friend Property Let RealHover(RHS As String)
mHover = RHS
End Property

Friend Property Set safe(RHS As Object)
Set safeform = RHS
End Property
Public Sub DisAllTargets()
DisableTargets q(), prive
End Sub
Friend Sub RenderTarget(bstack As basetask, rest$, Lang As Long, tHandle As Variant)
Dim p, w$, x
If tHandle \ 10000 <> prive Then
MyEr "target not for this form", "Ô ÛÙ¸˜ÔÚ ‰ÂÌ ÂﬂÌ·È „È· ·ıÙﬁ ÙÁ ˆ¸ÒÏ·"
Exit Sub
End If
p = tHandle Mod 10000
If p >= 0 And p < UBound(q()) Then
     
              '
While FastSymbol(rest$, ",")
x = Empty
If IsLabelSymbolNew(rest$, "÷—¡”«", "TEXT", Lang) Then
If IsStrExp(bstack, rest$, w$) Then q(p).Tag = w$
ElseIf IsLabelSymbolNew(rest$, "–≈Õ¡", "PEN", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).pen = x
ElseIf IsLabelSymbolNew(rest$, "÷œÕ‘œ", "BACK", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).back = x
ElseIf IsLabelSymbolNew(rest$, "–À¡…”…œ", "BORDER", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).fore = x
ElseIf IsLabelSymbolNew(rest$, "œƒ«√…¡", "COMMAND", Lang) Then
If IsStrExp(bstack, rest$, w$) Then q(p).Comm = w$
ElseIf IsLabelSymbolNew(rest$, "‘…Ã«", "VALUE", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).topval = Int(x * 100)
ElseIf IsLabelSymbolNew(rest$, "¬¡”«", "BASE", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).botval = Int(x * 100):
ElseIf IsLabelSymbolNew(rest$, "◊—ŸÃ¡", "COLOR", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).barC = x
ElseIf IsLabelSymbolNew(rest$, "Ã≈√≈»œ”", "SIZE", Lang) Then
If IsExp(bstack, rest$, x, , True) Then
If x > 100 Then x = 100
If x < -100 Then x = -100
q(p).imagesize = Int(x)
End If
' " ¡»≈‘«", "PORTRAIT"
ElseIf IsLabelSymbolNew(rest$, " ¡»≈‘«", "PORTRAIT", Lang) Then
    If IsExp(bstack, rest$, x, , True) Then q(p).Vertical = Int(x) <> 0
ElseIf IsLabelSymbolNew(rest$, "≈… œÕ¡", "IMAGE", Lang) Then
If IsExp(bstack, rest$, x) Then
    If bstack.lastobj Is Nothing Then
    Set q(p).drawimage = Nothing
    ElseIf TypeOf bstack.lastobj Is mHandler Then
        Dim usehandler As mHandler
        Set usehandler = bstack.lastobj
        Set bstack.lastobj = Nothing
        If usehandler.t1 = 2 Then
        Set q(p).drawimage = usehandler.objref
        Else
        GoTo err123
        End If
     
    Else
err123:
        WrongObject
        Exit Sub
    End If
    End If
End If
Wend
RTarget bstack, q(p)
End If
End Sub
Friend Function AddTarget(t As target) As Long
            If UBound(q()) < 9999 Then
                Targets = False
                ReDim Preserve q(UBound(q()) + 1)
                q(UBound(q()) - 1) = t
                AddTarget = prive * 10000 + UBound(q()) - 1
                Targets = True
            End If
End Function
Friend Sub EnableTarget(bstack As basetask, ByVal tHandle As Variant, p As Variant)
        If tHandle \ 10000 = prive Then
        tHandle = tHandle Mod 10000
        q(tHandle).Enable = Not (p = 0)
        RTarget bstack, q(tHandle)
        End If
End Sub
Public Sub ClearTargets()
    Targets = False
    ReDim q(0) As target
End Sub
Public Property Let CaptionW(ByVal NewValue As String)
If mNoCaption Then Exit Property
If LenB(NewValue) = 0 Then NewValue = "M2000"
    m_Caption = NewValue
DefWindowProcW Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue)
End Property

Friend Property Let Default(ctrlName$)
DefaultName = ctrlName$
End Property
Public Property Get CaptionW() As String
    If m_Caption = "M2000" Then
        CaptionW = vbNullString
    Else
        CaptionW = m_Caption
    End If
End Property
Public Property Let NoCaption(ByVal NewValue As Boolean) ' for task manager
mNoCaption = NewValue
If mNoCaption Then
DefWindowProcW Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue)
End If
End Property


Public Property Get Minimized() As Boolean
Minimized = Not Visible And VisibleOldState
End Property
Public Property Get TrueVisible() As Boolean
TrueVisible = Visible Or VisibleOldState
End Property
Public Property Let TrueVisible(RHS As Boolean)
    If mQuit Then
    
    Else
    VisibleOldState = RHS
    Visible = RHS
    End If
End Property
Public Sub AddGuiControl(widget As Object)
GuiControls.Add widget
End Sub
Public Sub TestModal(alfa As Double)
If mModalid = alfa Then
mModalid = mModalIdPrev
mModalIdPrev = 0
Enablecontrol = True
End If
End Sub
Public Sub NoTaskBar()
On Error Resume Next
drawminimized = False
mNoTaskBar = True
If Not MyForm3 Is Nothing Then
Unload MyForm3
Set MyForm3 = Nothing
End If
gList2.ShowMe
End Sub
Friend Property Let MyName(RHS As String)
mMyName$ = RHS
If IamPopUp Then Exit Property
drawminimized = Not IsWine
Set icon = Form1.icon


End Property
Friend Property Get MyName() As String
MyName = mMyName$
End Property
Friend Property Let TempTitle(RHS As String)
On Error Resume Next
Dim oldenable As Boolean
oldenable = gList2.enabled
gList2.enabled = True
gList2.HeadLine = vbNullString
If Trim$(RHS) = vbNullString Then RHS = " "
gList2.HeadLine = RHS
gList2.HeadlineHeight = gList2.HeightPixels
gList2.ShowMe
gList2.enabled = oldenable
End Property

Property Get Modal() As Double
    Modal = mModalid
End Property
Property Let Modal(RHS As Double)
mModalIdPrev = mModalid
mModalid = RHS
End Property
Public Property Get PopUpMenuVal() As Boolean
PopUpMenuVal = mPopUpMenu
End Property
Public Property Let PopUpMenuVal(RHS As Boolean)
mPopUpMenu = RHS
End Property
Public Property Let Enablecontrol(RHS As Boolean)
If RHS = False Then UnHook hWnd
If Len(mMyName$) = 0 Then Exit Property
If mEnabled = False And RHS = True Then Me.enabled = True
mEnabled = RHS

Dim w As Object
If Controls.Count > 0 Then
For Each w In Me.Controls
If w Is gList2 Then
gList2.enabled = RHS
gList2.mousepointer = 0
ElseIf w.Visible Then
w.enabled = RHS
If TypeOf w Is gList Then
w.TabStop = w.TabStopSoft
w.BypassKey = Not RHS
End If
ElseIf TypeOf w Is gList Then
w.BypassKey = Not RHS
End If
Next w
End If
Me.enabled = RHS
End Property
Public Property Get Enablecontrol() As Boolean
If Len(mMyName$) = 0 Then Enablecontrol = False: Exit Property
Enablecontrol = mEnabled
End Property

Property Get NeverShow() As Boolean
NeverShow = Not novisible
End Property
Friend Property Set EventObj(aEvent As Object)
Set myEvent = aEvent
Set myEvent.excludeme = New FastCollection
End Property
Friend Property Get EventObj() As Object
Set EventObj = myEvent
End Property
Friend Sub Callback(b$)
If Quit Then Exit Sub
If myEvent Is Nothing Then
Set EventObj = New mEvent
End If
If ByPassEvent Then
    If myEvent.excludeme.IamBusy Then Exit Sub
    Dim Mark$
    Mark$ = Split(b$, "(")(0)
    If myEvent.excludeme.ExistKey3(Mark$) Then Exit Sub
    If Not TaskMaster Is Nothing Then TaskMaster.tickdrop = 0
    
    If Visible Then
       myEvent.excludeme.AddKey2 Mark$
    If CallEventFromGuiOne(Me, myEvent, b$) Then
       If Not Quit Then myEvent.excludeme.Remove Mark$
    End If
    Else
        CallEventFromGuiOne Me, myEvent, b$
    End If
Else
    CallEventFromGui Me, myEvent, b$
End If
End Sub
Friend Sub CallbackNow(b$, VR())
If Quit Then Exit Sub
If myEvent Is Nothing Then
Set EventObj = New mEvent
End If

If myEvent.excludeme.IamBusy Then Exit Sub
Dim Mark$
Mark$ = Split(b$, "(")(0)
If myEvent.excludeme.ExistKey3(Mark$) Then Exit Sub
If Visible Then myEvent.excludeme.AddKey2 Mark$
If CallEventFromGuiNow(Me, myEvent, b$, VR()) Then myEvent.excludeme.Remove Mark$

End Sub


Public Sub ShowmeALL()
Dim w As Object
If IamPopUp Then
    If EnableStandardInfo Then
        glistN.menuEnabled(2) = False
    End If
    If Not MyForm3 Is Nothing Then
        drawminimized = False
        mNoTaskBar = True
        If Not MyForm3 Is Nothing Then
        Unload MyForm3
        Set MyForm3 = Nothing
        ttl = False
End If

End If
End If

If Controls.Count > 0 Then
For Each w In Controls
If w.enabled Then w.Visible = True
Next w
End If

gList2.PrepareToShow

End Sub
Public Sub RefreshALL()
Dim w As Object ', g As gList
If Controls.Count > 0 Then
For Each w In Controls
If w.Visible Then
    If TypeOf w Is gList Then w.ShowMe2
End If
Next w
End If
Refresh
End Sub

Private Sub Form_Click()
On Error Resume Next
If gList2.Visible Then gList2.SetFocus
If mIndex > -1 Then
    Callback mMyName$ + ".Click(" + CStr(index) + ")"
Else
    Callback mMyName$ + ".Click()"
End If
End Sub
Friend Sub SpreadKey(strKey As String)
Dim VR(1)
VR(0) = strKey
If mIndex > -1 Then
    CallbackNow mMyName$ + ".KeyPreview(" + CStr(index) + ")", VR()
Else
    CallbackNow mMyName$ + ".KeyPreview()", VR()
End If
strKey = VR(0)
End Sub
Private Sub Form_Activate()
On Error Resume Next
If Not Quit Then
If myEvent Is Nothing Then
Set EventObj = New mEvent
End If
If Not myEvent.excludeme.IamBusy Then
Set myEvent.excludeme = New FastCollection
End If
End If
If PopupOn Then
 mHover = ""
PopupOn = False
End If
If novisible Then Hide: Unload Me
gList2.mousepointer = 1
MarkSize = 4
ResizeMark.Width = MarkSize * dv15
ResizeMark.Height = MarkSize * dv15
ResizeMark.Left = Width - MarkSize * dv15
ResizeMark.top = Height - MarkSize * dv15
Dim XX As Long
XX = GetPixel(Me.Hdc, Width - MarkSize * dv15, Height - MarkSize * dv15)
If XX <> -1 Then
ResizeMark.backcolor = XX
End If
ResizeMark.Visible = Sizable
If Sizable Then ResizeMark.ZOrder 0
If HOOKTEST <> 0 Then UnHook HOOKTEST
If Not NoHook Then
If Typename(ActiveControl) = "gList" Then

Hook hWnd, ActiveControl

Else
Hook hWnd, Nothing
End If
End If
If DefaultName <> "" Then
Dim aa As gList

Set aa = Controls(DefaultName)
If IamPopUp Then
aa.SetFocus
Else
LastActive = DefaultName
End If
DefaultName = vbNullString
End If
If IamPopUp Then Exit Sub
If Not moveMe Then
If LastActive <> "" Then
    If Controls(LastActive).enabled Then
    If Controls(LastActive).Visible Then Controls(LastActive).SetFocus
    End If
End If
End If
If mNoTaskBar Then Exit Sub
If ttl Then
If MyForm3.Visible Then
Set MyForm3.lastform = Me: MyForm3.CaptionW = gList2.HeadLine
MyForm3.Timer1.enabled = False
Else
MyForm3.hideme = False

If MyForm3.Timer1.Interval = 10000 Then MyForm3.Timer1.Interval = 20
MyForm3.Timer1.enabled = True
MyForm3.WindowState = 0
MyForm3.Visible = True

End If
End If
End Sub

Private Sub Form_Deactivate()
UNhookMe
If PopupOn Then

Exit Sub
End If
If IamPopUp Then
If mModalid = Modalid And Modalid <> 0 Then
If Visible Then Hide
Modalid = 0
novisible = False
End If
Else
If mModalid = Modalid And Modalid <> 0 Then If Not (Visible Or Me.VisibleOldState) Then If mModalid <> 0 Then Modalid = 0
End If

End Sub


Private Sub Form_Initialize()
mEnabled = True
ClearTargets
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Me.Visible Then
If ActiveControl Is Nothing Then
Dim w As Object
    If Controls.Count > 0 Then
    For Each w In Controls
    If w.Visible Then
    If TypeOf w Is gList Then
    w.SetFocus
    Exit For
    End If
    End If
    Next w
    Set w = Nothing
    End If
    Else
    
    If Typename(ActiveControl) = "gList" Then ActiveControl.SetFocus
End If
Else
choosenext
End If
End Sub

Private Sub Form_LostFocus()
If Not IamPopUp Then mHover = ""
If mIndex > -1 Then
    Callback mMyName$ + ".LostFocus(" + CStr(index) + ")"
Else
    Callback mMyName$ + ".LostFocus()"
End If
If HOOKTEST <> 0 Then
UnHook hWnd
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
Dim bstack As basetask, oldhere$


If Not Relax Then
RealHover = ""


Relax = True

Dim sel&

If Button > 0 And Targets Then
    sel& = ScanTarget(q(), CLng(x), CLng(y), prive)
    If sel& >= 0 Then
        If Button = 1 Then

            
            Select Case q(sel&).id Mod 100
            Case Is < 10
                SwapStrings here$, oldhere$
                here$ = modulename
                Set bstack = New basetask
                Set bstack.Owner = Me
                Set bstack.Sorosref = New mStiva
                If Execute(bstack, (q(sel&).Comm), False) = 0 Then Beep
                SwapStrings here$, oldhere$
                
            Case Else
            
            If mIndex > -1 Then
                Callback mMyName$ + ".Target" + "(" + CStr(index) + "," + Str(sel& + prive * 10000) + ")"
            Else
                Callback mMyName$ + ".Target" + "(" + Str(sel& + prive * 10000) + ")"
            End If

            End Select
            
        End If
        
        Button = 0
        Relax = False
        Exit Sub
    End If
End If



If mIndex > -1 Then
    Callback mMyName$ + ".MouseDown(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
    Callback mMyName$ + ".MouseDown(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If



Relax = False
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then
Relax = True

If mIndex > -1 Then
Callback mMyName$ + ".MouseMove(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback mMyName$ + ".MouseMove(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then

Relax = True

If mIndex > -1 Then
Callback mMyName$ + ".MouseUp(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback mMyName$ + ".MouseUp(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If mModalid = Modalid And Modalid <> 0 Then
    If Visible Then
    Hide
    Quit = byPassCallback
    End If
    If mModalid <> 0 Then Modalid = 0
    Cancel = Not byPassCallback
    novisible = False
ElseIf mModalid <> 0 And Visible Then
    mModalid = mModalIdPrev
    mModalIdPrev = 0
    If mModalid > 0 Then
        Cancel = True
    Else
        Quit = True
    End If
Else
mModalIdPrev = 0
Quit = Not byPassCallback
If Quit And Visible Then Visible = False
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 0 Then WindowState = 0: Exit Sub
gList2.MoveTwips 0, 0, Me.Width, gList2.HeightTwips
ResizeMark.move Width - ResizeMark.Width, Height - ResizeMark.Height
ResizeMark.backcolor = GetPixel(Me.Hdc, Width \ dv15 - 1, Height \ dv15 - 1)
End Sub


Private Sub gList2_AccKey(m As Long)
AccProces m
End Sub
Private Sub gListN_AccKey(m As Long)
AccProces m
End Sub
Private Sub gList2_BlinkNow(Face As Boolean)
If mTimes > 0 Then
    mTimes = mTimes - 1
    If mTimes = 0 Then
    gList2.BlinkON = False: gList2.CapColor = mBarColor
    If Stored Then RestoreBlibkStatus
    Else
        State Face
    End If
Else
    State Face
End If
gList2.ShowMe
If Stored Then Exit Sub
  If mIndex >= 0 Then
   Callback mMyName$ + ".Blink(" + Str(index) + "," + Str(Face) + ")"
   Else
      Callback mMyName$ + ".Blink(" + Str(Face) + ")"
      End If
End Sub
Public Sub State(Face As Boolean)
    If Face Then
        gList2.CapColor = rgb(255, 160, 0)
    Else
        gList2.CapColor = rgb(128, 80, 0)
    End If

End Sub

Public Property Let Blink(ByVal vNewValue As Variant)
If vNewValue = 0 Then gList2.CapColor = rgb(255, 160, 0): gList2.ShowMe
gList2.BlinkTime = Abs(vNewValue)
End Property
Public Property Let BlinkTimes(ByVal vNewValue As Variant)
mTimes = vNewValue
End Property
Private Sub StoreBlinkStatus()
If Stored Then Exit Sub
LastBlinkmTimes = mTimes
lastBlink = gList2.BlinkTime
lastBlinkOn = gList2.BlinkON
Stored = True
End Sub
Private Sub RestoreBlibkStatus()
On Error Resume Next
If Not Stored Then
 gList2.ShowMe
Exit Sub
End If
mTimes = LastBlinkmTimes
gList2.BlinkTime = lastBlink
gList2.BlinkON = lastBlinkOn
Stored = False
End Sub
Private Sub gList2_CtrlPlusF1()
    If mIndex > -1 Then
        Callback mMyName$ + ".About(" + CStr(index) + ")"
    Else
        Callback mMyName$ + ".About()"
    End If
End Sub

Private Sub gList2_EnterOnly()
    If mIndex > -1 Then
        Callback mMyName$ + ".Enter(" + CStr(index) + ")"
    Else
        Callback mMyName$ + ".Enter()"
    End If
End Sub

Private Sub gList2_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = -1 Then
Dim m As Long
m = gList2.CapColor
If gList2.BlinkON Or Not gList2.NoHeaderBackground Then FillThere thisHDC, thisrect, m
FillThereMyVersionTrans thisHDC, thisrect, m, UseReverse
If mSizable And mShowMaximize Then FillThereMyVersion3Trans thisHDC, thisrect, m, 0, UseReverse Else minimPos = 0
If drawminimized Then FillThereMyVersion2 thisHDC, thisrect, minimPos, UseReverse
If UseInfo Then FillThereMyVersion4 thisHDC, thisrect, infopos, UseReverse
If mIcon Then
drawicon thisHDC
End If
skip = True
End If
End Sub
Private Sub OpenInfo()
Dim gl As Long, thisLastControl As String

gl = glistN.listcount
If gl = 0 Then Exit Sub
Pad.Width = mMenuWidth    'CLng(Width / 1.618 * dv15) \ dv15
If gl > 9 Then gl = 9

glistN.restrictLines = gl

Pad.Height = (((glistN.HeadlineHeightTwips * gl) \ dv15) + 1) * dv15
glistN.MoveTwips 0, 0, Pad.Width, Pad.Height
glistN.LeaveonChoose = True
glistN.ListindexPrivateUseFirstFree 0
glistN.FreeMouse = True
glistN.ShowBar = False
glistN.PanPos = 0
glistN.NoWheel = True
If Not UseReverse Then
PopUpPos Pad, Width - mMenuWidth, gList2.Height / 2, gList2.Height / 2
Else
PopUpPos Pad, 0, gList2.Height / 2, gList2.Height / 2
End If
End Sub

Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
On Error Resume Next ' for set focus
If Button <> 1 Then Exit Sub
ByPassColor = True
If UseReverse Then
If UseInfo Then
If gList2.SingleClickCheck(Button, item, x, y, setupxy * (1 + 2 * infopos) / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1, -1) Then
        If NoEventInfo Then
            OpenInfo
        Else
            If mIndex > -1 Then
                Callback mMyName$ + ".Info(" + CStr(index) + "," + CStr(x) + "," + CStr(y) + ")"
            Else
                Callback mMyName$ + ".Info(" + CStr(x) + "," + CStr(y) + ")"
            End If
        End If
   Exit Sub
End If
End If
If gList2.DoubleClickCheck(Button, item, x, y, gList2.WidthPixels - setupxy / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1, -1) Then
    ByeBye
    Exit Sub
End If
If mSizable And mShowMaximize Then
    If gList2.SingleClickCheck(Button, item, x, y, gList2.WidthPixels - setupxy * 3 / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1, -1) Then
        If IhaveLastPos Then
            On Error Resume Next
            With ScrInfo(FindMonitorFromMouse)
            If IsWine And .Left = 0 And .top = 0 And .Width - 1 = Width And .Height - 1 = Height And (.Left <> Left Or .top <> top) Then
                Me.move .Left, .top
                If LastActive <> "" Then
                    If Controls(LastActive).enabled Then
                    If Controls(LastActive).Visible Then Controls(LastActive).SetFocus
                    End If
                End If
                Exit Sub
            ElseIf .Width = Width And .Height = Height And (.Left <> Left Or .top <> top) Then
                Me.move .Left, .top
                If LastActive <> "" Then
                    If Controls(LastActive).enabled Then
                    If Controls(LastActive).Visible Then Controls(LastActive).SetFocus
                    End If
                End If
                Exit Sub
            Else
                IhaveLastPos = False
                If IsWine And .Left = 0 And .top = 0 Then
                    Width = MeWidth
                    Height = MeHeight
                    Me.move MeLeft, MeTop
                Else
                    Me.move MeLeft, MeTop, MeWidth, MeHeight
                End If
            End If

            End With

        Else
            IhaveLastPos = True
            MeLeft = Left
            MeTop = top
            MeWidth = Width
            MeHeight = Height
            On Error Resume Next
            With ScrInfo(FindMonitorFromMouse)
                If IsWine And .Left = 0 And .top = 0 Then
                    Width = .Width - 1
                    Height = .Height - 1
                    move .Left, .top
                Else
                move .Left, .top, .Width, .Height
                End If
            End With
        End If
        If mIndex > -1 Then
           Callback mMyName$ + ".Resize(" + CStr(index) + ")"
        Else
           Callback mMyName$ + ".Resize()"
        End If
        ResizeMark.backcolor = GetPixel(Me.Hdc, Width \ dv15 - 1, Height \ dv15 - 1)
        If IhaveLastPos Then
                
                If mIndex > -1 Then
                    Callback mMyName$ + ".Maximized(" + CStr(index) + ")"
                Else
                    Callback mMyName$ + ".Maximized()"
                End If
                MenuSet 2
        Else
                If mIndex > -1 Then
                    Callback mMyName$ + ".Restored(" + CStr(index) + ")"
                Else
                    Callback mMyName$ + ".Restored()"
                End If
               MenuSet 3
        End If
         Exit Sub
    End If

End If
If Not IsWine And drawminimized Then
    If Not MyForm3 Is Nothing Then
        If gList2.SingleClickCheck(Button, item, x, y, gList2.WidthPixels - setupxy * (3 + 2 * minimPos) / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1, -1) Then
            VisibleOldState = Visible
            Visible = False
            MinimizeON
            Exit Sub
        End If
    End If
End If
Else
If UseInfo Then
If gList2.SingleClickCheck(Button, item, x, y, gList2.WidthPixels - setupxy * (1 + 2 * infopos) / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1, -1) Then
        If NoEventInfo Then
            OpenInfo
        Else
            If mIndex > -1 Then
                Callback mMyName$ + ".Info(" + CStr(index) + "," + CStr(x) + "," + CStr(y) + ")"
            Else
                Callback mMyName$ + ".Info(" + CStr(x) + "," + CStr(y) + ")"
            End If
        End If
    Exit Sub
End If
End If
If gList2.DoubleClickCheck(Button, item, x, y, setupxy / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1, -1) Then
    ByeBye
    Exit Sub
End If
If mSizable And mShowMaximize Then
    If gList2.SingleClickCheck(Button, item, x, y, setupxy * 3 / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1, -1) Then
        If IhaveLastPos Then
            On Error Resume Next
            With ScrInfo(FindMonitorFromMouse)
            If IsWine And .Left = 0 And .top = 0 And .Width - 1 = Width And .Height - 1 = Height And (.Left <> Left Or .top <> top) Then
                Me.move .Left, .top
                Exit Sub
            ElseIf .Width = Width And .Height = Height And (.Left <> Left Or .top <> top) Then
                Me.move .Left, .top
                Exit Sub
            Else
                IhaveLastPos = False
                If IsWine And .Left = 0 And .top = 0 Then
                    Width = MeWidth
                    Height = MeHeight
                    Me.move MeLeft, MeTop
                Else
                    Me.move MeLeft, MeTop, MeWidth, MeHeight
                End If
            End If

            End With

        Else
            IhaveLastPos = True
            MeLeft = Left
            MeTop = top
            MeWidth = Width
            MeHeight = Height
            On Error Resume Next
            With ScrInfo(FindMonitorFromMouse)
                If IsWine And .Left = 0 And .top = 0 Then
                    Width = .Width - 1
                    Height = .Height - 1
                    move .Left, .top
                Else
                move .Left, .top, .Width, .Height
                End If
            End With
            
        End If
        If mIndex > -1 Then
           Callback mMyName$ + ".Resize(" + CStr(index) + ")"
        Else
           Callback mMyName$ + ".Resize()"
        End If
        ResizeMark.backcolor = GetPixel(Me.Hdc, Width \ dv15 - 1, Height \ dv15 - 1)
        If IhaveLastPos Then
                If mIndex > -1 Then
                    Callback mMyName$ + ".Maximized(" + CStr(index) + ")"
                Else
                    Callback mMyName$ + ".Maximized()"
                End If
                MenuSet 2
        Else
                If mIndex > -1 Then
                    Callback mMyName$ + ".Restored(" + CStr(index) + ")"
                Else
                    Callback mMyName$ + ".Restrored()"
                End If
                MenuSet 3
        End If
        Exit Sub
    End If
       
End If
If Not IsWine And drawminimized Then
    If Not MyForm3 Is Nothing Then
        If gList2.SingleClickCheck(Button, item, x, y, setupxy * (3 + 2 * minimPos) / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1, -1) Then
            VisibleOldState = Visible
            Visible = False
            MinimizeON
            Exit Sub
        End If
    End If
End If
End If
End Sub
Sub ByeBye()
Dim var(1) As Variant
var(0) = CLng(0)
If mIndex > -1 Then
If Not Quit Then CallEventFromGuiNow Me, myEvent, mMyName$ + ".Unload(" + CStr(mIndex) + ")", var()
Else
If Not Quit Then CallEventFromGuiNow Me, myEvent, mMyName$ + ".Unload()", var()
End If
If var(0) = 0 Then
                     If ttl And Not mNoTaskBar Then
                     If Not MyForm3 Is Nothing Then
                     MyForm3.CaptionW = vbNullString
                     If MyForm3.WindowState = 1 Then MyForm3.WindowState = 0
               
                    Unload MyForm3
                    End If
             End If
                              Unload Me
                      End If
End Sub
Friend Sub ByeBye2(ret As Long)
Dim var(1) As Variant
var(0) = CLng(0)
If mIndex > -1 Then
If Not Quit Then CallEventFromGuiNow Me, myEvent, mMyName$ + ".Unload(" + CStr(mIndex) + ")", var()
Else
If Not Quit Then CallEventFromGuiNow Me, myEvent, mMyName$ + ".Unload()", var()
End If
ret = var(0)
End Sub
Private Sub Form_Load()
If onetime Then
novisible = True
Exit Sub
End If
If Not safeform Is Nothing Then
If Not safeform.ExistKey(hWnd) Then safeform.AddKey hWnd Else safeform.ValueStr = ""
End If
SkipFirstClick = True
mShowMaximize = True
infopos = 0
onetime = True
minimPos = 1
mQuit = False
Set LastGlist = Nothing
scrTwips = Screen.TwipsPerPixelX
lastfactor = 1
setupxy = 20
gList2.FreeMouse = True
gList2.Font.Size = 14.25 * dv15 / 15
gList2.enabled = True
mIconColor = 0
mBarColor = rgb(255, 160, 0)
gList2.CapColor = mBarColor
gList2.FloatList = True
gList2.MoveParent = True
gList2.HeadLine = vbNullString
gList2.HeadLine = "Form"
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
gList2.TabStop = False
With gList2.Font
CtrlFont.Name = .Name
CtrlFont.Size = .Size
CtrlFont.bold = .bold
End With
gList2.FloatLimitTop = VirtualScreenHeight() - 600
gList2.FloatLimitLeft = VirtualScreenWidth() - 450
Dim mm As Long
mm = Forms.Count
With ScrInfo(Console)
    If (.Left + .Width / 16 + mm * dv15 * 10) > .Width * 7 / 8 Or (.top + .Height / 16 + mm * dv15 * 10) > .Height * 7 / 8 Then
    move .Left, .top
    Else
    move .Left + .Width / 16 + mm * dv15 * 10, .top + .Height / 16 + mm * dv15 * 10
    End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
UNhookMe
Quit = True
Set myEvent = Nothing
If Not glistN Is Nothing Then glistN.Shutdown
Set glistN = Nothing
If Not Pad Is Nothing Then Unload Pad
Set Pad = Nothing
If prive <> 0 Then
players(prive).used = False
players(prive).MAXXGRAPH = 0
prive = 0

End If
Dim w As Object
If GuiControls.Count > 0 Then
For Each w In GuiControls
    w.deconstruct
Next w
End If
If ttl Then If Not MyForm3 Is Nothing Then Unload MyForm3
VisibleOldState = False
If Not safeform Is Nothing Then
If safeform.ExistKey(hWnd) Then safeform.ValueStr = "skip"
End If

If TaskMaster Is Nothing Then Exit Sub
TaskMaster.CheckThreadsForThisObject Me

End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
Private Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT
CopyFromLParamToRect a, thatRect

FillBack thathDC, a, thatbgcolor
End Sub
Private Sub FillThereMyVersionTrans(thathDC As Long, thatRect As Long, thatbgcolor As Long, Reverse As Boolean)
Dim a As RECT, b As Long, aline As RECT
b = 2 * lastfactor
If b < 2 Then b = 2
If setupxy - b < 0 Then b = setupxy \ 4 + 1
CopyFromLParamToRect a, thatRect
If Reverse Then
    a.Left = a.Right - setupxy + b
    a.Right = a.Right - b
Else
    a.Left = b
    a.Right = setupxy - b
End If
a.top = b
a.Bottom = setupxy - b
If b < 1 Then b = 1 Else b = b * 2 - 1
aline = a
aline.Left = a.Right - b
FillThere thathDC, VarPtr(aline), mIconColor
aline = a
aline.Right = a.Left + b
FillThere thathDC, VarPtr(aline), mIconColor

aline = a
aline.Bottom = a.top + b
FillThere thathDC, VarPtr(aline), mIconColor
aline = a
aline.top = a.Bottom - b
FillThere thathDC, VarPtr(aline), mIconColor
End Sub
Private Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long, Reverse As Boolean)
Dim a As RECT, b As Long
b = 2 * lastfactor
If b < 2 Then b = 2
If setupxy - b < 0 Then b = setupxy \ 4 + 1
CopyFromLParamToRect a, thatRect
If Reverse Then
    a.Left = a.Right - setupxy + b
    a.Right = a.Right - b
Else
    a.Left = b
    a.Right = setupxy - b
End If
a.top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), mIconColor
CopyFromLParamToRect a, thatRect
b = 5 * lastfactor
If Reverse Then
    a.Left = a.Right - setupxy + b
    a.Right = a.Right - b
Else
    a.Left = b
    a.Right = setupxy - b
End If
a.top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), thatbgcolor
End Sub
Private Sub FillThereMyVersion2(thathDC As Long, thatRect As Long, butPos As Long, Reverse As Boolean)
Dim a As RECT, b As Long, c As Long
b = 2 * lastfactor
If b < 2 Then b = 2
If setupxy - b < 0 Then b = setupxy \ 4 + 1
c = setupxy * (butPos)
CopyFromLParamToRect a, thatRect
If Reverse Then
    a.Left = a.Right - (c + setupxy * 2) + b
    a.Right = a.Right - b - c - setupxy

Else
    a.Left = b + c + setupxy
    a.Right = setupxy - 2 * b + a.Left
End If
a.Bottom = setupxy - b
b = 5 * lastfactor

If 5 * lastfactor < 1 Then
a.top = a.Bottom - 1
Else
a.top = setupxy - b
End If
FillThere thathDC, VarPtr(a), mIconColor
End Sub
Private Sub FillThereMyVersion3Trans(thathDC As Long, thatRect As Long, thatbgcolor As Long, butPos As Long, Reverse As Boolean)
Dim a As RECT, b As Long, c As Long
b = 2 * lastfactor
If b < 2 Then b = 2
If setupxy - b < 0 Then b = setupxy \ 4 + 1
c = setupxy * butPos
CopyFromLParamToRect a, thatRect
If Reverse Then
    a.Left = a.Right - (c + setupxy * 2) + b
    a.Right = a.Right - b - c - setupxy
Else
    a.Left = b + c + setupxy
    a.Right = setupxy - 2 * b + a.Left
End If
a.top = 3 * b
a.Bottom = setupxy - b
Dim aline As RECT
If b < 2 Then b = 2 Else b = b * 2 - 1
aline = a
aline.Left = a.Right - b
FillThere thathDC, VarPtr(aline), mIconColor
aline = a
aline.Right = a.Left + b
FillThere thathDC, VarPtr(aline), mIconColor

aline = a
aline.Bottom = a.top + b
FillThere thathDC, VarPtr(aline), mIconColor
aline = a
aline.Bottom = a.Bottom - b * 2
aline.top = a.Bottom - b
FillThere thathDC, VarPtr(aline), mIconColor
End Sub
Private Sub FillThereMyVersion3(thathDC As Long, thatRect As Long, thatbgcolor As Long, butPos As Long, Reverse As Boolean)
Dim a As RECT, b As Long, c As Long
b = 2 * lastfactor
If b < 2 Then b = 2
If setupxy - b < 0 Then b = setupxy \ 4 + 1
c = setupxy * butPos
CopyFromLParamToRect a, thatRect
If Reverse Then
    a.Left = a.Right - (c + setupxy * 2) + b
    a.Right = a.Right - b - c - setupxy
Else
    a.Left = b + c + setupxy
    a.Right = setupxy - 2 * b + a.Left
End If
a.top = 3 * b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), mIconColor
CopyFromLParamToRect a, thatRect
b = 5 * lastfactor
If b < 2 Then b = 2
If Reverse Then
    a.Left = a.Right - (c + setupxy * 2) + b
    a.Right = a.Right - b - c - setupxy

Else
    a.Left = b + c + setupxy
    a.Right = setupxy - 2 * b + a.Left
End If


a.top = 4 * b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), thatbgcolor
a.Bottom = a.top - b
b = 2 * lastfactor
If b < 2 Then b = 2
a.Bottom = a.Bottom - b - b / 2

a.top = 4 * b + b / 2
FillThere thathDC, VarPtr(a), thatbgcolor


End Sub
Private Sub FillThereMyVersion4(thathDC As Long, thatRect As Long, butPos As Long, Reverse As Boolean)
Dim a As RECT, b As Long, c As Long
Dim color1 As Long
color1 = mIconColor
If Not moveMe And Not ByPassColor Then
If Screen.ActiveForm Is Me Then
If Screen.ActiveControl Is gList2 Then
If mIconColor = 16777215 Then color1 = 32768 Else color1 = 16777215
End If
End If
End If
b = 2 * lastfactor
If b < 2 Then b = 2
If setupxy - b < 0 Then b = setupxy \ 4 + 1
CopyFromLParamToRect a, thatRect
If Reverse Then
c = setupxy * (butPos - 1)
a.Left = b + c + setupxy
a.Right = setupxy - 2 * b + a.Left
Else
c = setupxy * butPos
a.Left = a.Right - (c + setupxy) + b
 a.Right = a.Right - b - c
End If
a.Bottom = setupxy - b
b = 5 * lastfactor

If 5 * lastfactor < 1 Then
a.top = a.Bottom - 1
Else
a.top = setupxy - b
End If
FillThere thathDC, VarPtr(a), color1
a.top = a.top - b
a.Bottom = a.Bottom - b
FillThere thathDC, VarPtr(a), color1
a.top = a.top - b
a.Bottom = a.Bottom - b
FillThere thathDC, VarPtr(a), color1

End Sub
Public Property Get Title() As Variant
Title = gList2.HeadLine
End Property
Public Sub ShowTaskBar()
    If IamPopUp Then Exit Sub

    If mNoTaskBar Then Exit Sub
    If Not ttl Then
        drawminimized = Not IsWine
        If Not MyForm3 Is Nothing Then
        Else
            Set MyForm3 = New Form3
        End If
        Set MyForm3.lastform = Me
        MyForm3.Timer1.enabled = False
        ttl = True
        MyForm3.WindowState = 0
    End If
    MyForm3.CaptionW = gList2.HeadLine
    MyForm3.Visible = True
    MyForm3.Refresh
    If Err Then
    Err.Clear
    Sleep 10
    MyForm3.Visible = True
    
    End If


End Sub
Public Property Let Title(ByVal vNewValue As Variant)
' A WORKAROUND TO CHANGE TITLE WHEN FORM IS DISABLED BY A MODAL FORM
On Error Resume Next
Dim oldenable As Boolean
If vNewValue <> vbNullString Then
  
    oldenable = gList2.enabled
    gList2.enabled = True
    gList2.HeadLine = vbNullString
    If Trim$(vNewValue) = vbNullString Then vNewValue = " "
    gList2.HeadLine = vNewValue
    gList2.HeadlineHeight = gList2.HeightPixels
    gList2.ShowMe
    gList2.enabled = oldenable
    If IamPopUp Then Exit Property

    If mNoTaskBar Then Exit Property
    'If Not ttl Then
     '   drawminimized = Not IsWine
     '   If Not MyForm3 Is Nothing Then
     '   Else
     '       Set MyForm3 = New Form3
     '   End If
     '   Set MyForm3.lastform = Me
     '   MyForm3.Timer1.enabled = False
     '   ttl = True
     '   MyForm3.WindowState = 0
   ' End If
   If ttl Then
    MyForm3.CaptionW = gList2.HeadLine
    MyForm3.Visible = True
    MyForm3.Refresh
    If Err Then
    Err.Clear
    Sleep 10
    MyForm3.Visible = True
    End If
    End If
Else
    If ttl Then Unload MyForm3: ttl = False
    Set MyForm3 = Nothing
    Exit Property
End If
CaptionW = gList2.HeadLine
End Property
Public Property Get index() As Long
index = mIndex
End Property

Friend Property Let index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Sub CloseNow()
Dim w As Object
    If mModalid = Modalid And Modalid <> 0 Then
        Modalid = 0
      If Visible Then Hide
    Else
    mModalid = 0
    For Each w In GuiControls
    If Left$(Typename(w), 3) = "Gui" Then
    w.deconstruct
    End If
Next w
Set w = Nothing
         If ttl And Not mNoTaskBar Then
                    Unload MyForm3
             End If

Unload Me
    End If
End Sub
Public Function Control(index) As Object
On Error Resume Next
Set Control = Controls(index)
If Err > 0 Then Set Control = Me
End Function
Public Sub Opacity(mAlpha, Optional mlColor = 0, Optional mTRMODE = 0)
SetTrans Me, CInt(Abs(mAlpha)) Mod 256, CLng(mycolor(mlColor)), CBool(mTRMODE)
End Sub
Public Sub Hold()
If Not Sizable Then
If MY_BACK Is Nothing Then Set MY_BACK = New cDIBSection
MY_BACK.ClearUp
If MY_BACK.create(Width / DXP, Height / DYP) Then
MY_BACK.LoadPictureBlt Hdc
If MY_BACK.bitsPerPixel <> 24 Then Conv24 MY_BACK
End If
End If
End Sub
Public Sub Release()
If Not Sizable Then
If MY_BACK Is Nothing Then Exit Sub
MY_BACK.PaintPicture Hdc
End If
End Sub


Public Property Get ByPass() As Variant
ByPass = ByPassEvent
End Property

Public Property Let ByPass(ByVal vNewValue As Variant)
ByPassEvent = CBool(vNewValue)
End Property
Property Get TitleHeight() As Variant
TitleHeight = gList2.Height
End Property
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.Name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set gList2.Font = aa
gList2.Height = gList2.HeadlineHeightTwips
lastfactor = gList2.HeadlineHeight / 30
setupxy = 20 * lastfactor
 gList2.Dynamic

End If
End Sub
Public Sub CtrlFontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)

If ThisFontName <> "" Then

CtrlFont.Name = ThisFontName

If ThisMode > 7 Then CtrlFont.Size = ThisMode Else CtrlFont = 7
CtrlFont.bold = ThisBold

End If
End Sub
Public Property Get CtrlFontName()
    CtrlFontName = CtrlFont.Name
End Property
Public Property Get CtrlFontSize()
    CtrlFontSize = CtrlFont.Size
End Property
Public Property Get CtrlFontBold()
    CtrlFontBold = CtrlFont.bold
End Property
Friend Sub SendFKEY(a As Integer)

 If mIndex >= 0 Then
      Callback mMyName$ + ".Fkey(" + Str(mIndex) + "," + Str(a) + ")"
   Else
      Callback mMyName$ + ".Fkey(" + Str(a) + ")"
      End If
End Sub

Private Sub gList2_Fkey(a As Integer)
If a > 1000 Then
SendFKEY a - 1000
Else
SendFKEY a
End If
End Sub

Private Sub gList2_KeyDown(keycode As Integer, shift As Integer)
If keycode = 115 And shift = 4 Then
ByeBye
Exit Sub
End If
If moveMe Then
If shift = 0 Then
Select Case keycode
Case vbKeyLeft
movemeX = movemeX - 10 * dv15
Case vbKeyRight
movemeX = movemeX + 10 * dv15
Case vbKeyUp
movemeY = movemeY - 10 * dv15
Case vbKeyDown
movemeY = movemeY + 10 * dv15
Case Else
    RestoreBlibkStatus
    sizeMe = False
     moveMe = False
     On Error Resume Next
        If LastActive <> "" Then
            If Controls(LastActive).enabled Then
            If Controls(LastActive).Visible Then Controls(LastActive).SetFocus
            End If
        End If
    Exit Sub
End Select
keycode = 0
Else
Select Case keycode
Case vbKeyLeft
movemeX = movemeX - dv15
Case vbKeyRight
movemeX = movemeX + dv15
Case vbKeyUp
movemeY = movemeY - dv15
Case vbKeyDown
movemeY = movemeY + dv15
End Select
keycode = 0
End If
If Not sizeMe Then
gList2.FloatListMe True, movemeX, movemeY
gList2.FloatListMe False, movemeX, movemeY
Else
If movemeY < 3000 Then movemeY = 3000
If movemeX < 3000 Then movemeX = 3000
If VirtualScreenHeight < movemeY Then movemeY = VirtualScreenHeight
If VirtualScreenWidth < movemeX Then movemeX = VirtualScreenWidth

                move Me.Left, Me.top, movemeX, movemeY
                If mIndex > -1 Then
                    Callback mMyName$ + ".Resize(" + CStr(index) + ")"
                Else
                    Callback mMyName$ + ".Resize()"
                End If
                Form_Resize
                

End If
keycode = 0
Exit Sub
Else

Dim VR(2)
VR(0) = keycode
VR(1) = shift
If mIndex > -1 Then
    CallbackNow mMyName$ + ".KeyDown(" + CStr(index) + ")", VR()
Else
    CallbackNow mMyName$ + ".KeyDown()", VR()
End If
shift = VR(1)
keycode = VR(0)
If keycode = 40 Then
If NoEventInfo Then
If Not Pad.Visible Then OpenInfo
End If
End If
End If
End Sub

Private Sub gList2_LostFocus()
RestoreBlibkStatus
ByPassColor = False
sizeMe = False
moveMe = False
gList2.mousepointer = 1
End Sub


Private Sub gList2_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If Button <> 0 Then RestoreBlibkStatus: moveMe = False: sizeMe = False
End Sub

Private Sub gList2_MouseUp(x As Single, y As Single)
RestoreBlibkStatus
moveMe = False
sizeMe = False
If Not Pad Is Nothing Then
If Pad.Visible Then Exit Sub
End If


            On Error Resume Next
            If Me.WindowState = 1 Then Exit Sub
            If LastActive <> vbNullString Then
            On Error GoTo 1000
            If Controls(LastActive).enabled Then
                If Controls(LastActive).Visible Then
                    If Not UseReverse Then
                        gList2.DoubleClickArea x, y, setupxy / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1
                    Else
                        gList2.DoubleClickArea x, y, gList2.WidthPixels - setupxy / 2, setupxy / 3, Abs(setupxy / 2 - 2) + 1
                    End If
                    Controls(LastActive).SetFocus
                End If
            End If
1000             If Err Then LastActive = Screen.ActiveControl.Name
            If Err Then Debug.Print "error:(" & Err.Description & ")": Exit Sub
            End If
            If MyForm3 Is Nothing Then LastActive = vbNullString: Exit Sub
            If MyForm3.WindowState <> 1 Then LastActive = vbNullString
            
        
End Sub

Private Sub gList2_RefreshDesktop()
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
End Sub
Public Sub PopUp(vv As Variant, ByVal x As Variant, ByVal y As Variant)
Dim var1() As Variant, retobject As Object, that As Object, hmonitor As Long
ReDim var1(0 To 1)
Dim var2() As String
ReDim var2(0 To 0)
hmonitor = FindMonitorFromPixel(x, y) ' FindFormSScreen(Me)
x = x + Left
y = y + top
Set that = vv
If Me Is that Then Exit Sub
If that.Visible Then
If Not that.enabled Then Exit Sub
End If
If x + that.Width > ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left Then
If y + that.Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).top Then
that.move ScrInfo(hmonitor).Width - that.Width + ScrInfo(hmonitor).Left, ScrInfo(hmonitor).Height - that.Height + ScrInfo(hmonitor).top
Else
that.move ScrInfo(hmonitor).Width - that.Width + ScrInfo(hmonitor).Left, y + ScrInfo(hmonitor).top
End If
ElseIf y + that.Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).top Then
that.move x, ScrInfo(hmonitor).Height - Height '+ ScrInfo(hmonitor).top
Else
that.move x, y
End If
var1(1) = 1
Set var1(0) = Me
that.IamPopUp = True

CallByNameFixParamArray that, "Show", VbMethod, var1(), var2(), 2
Set that = Nothing
Set var1(0) = Nothing
MyDoEvents

End Sub
Public Sub PopUpPos(vv As Variant, ByVal x As Variant, ByVal y As Variant, ByVal y1 As Variant)
Dim that As Object, hmonitor As Long

x = x + Left
y = y + top + y1
'hmonitor = FindFormSScreen(Me)
hmonitor = FindMonitorFromPixel(x, y)

Set that = vv
If Me Is that Then Exit Sub
If that.Visible Then
If Not that.enabled Then Exit Sub
End If
If x + that.Width > ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left Then
If y + that.Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).top Then
that.move ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left - that.Width, y - that.Height - y1 ' + ScrInfo(hmonitor).top
Else
that.move ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left - that.Width, y '+ ScrInfo(hmonitor).top
End If
ElseIf y + that.Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).top Then
that.move x, y - that.Height - y1
Else
that.move x, y
End If
that.ShowmeALL
PopupOn = True

that.Show , Me

End Sub
Public Sub hookme(this As gList)
Set LastGlist = this
End Sub


Private Sub gList2_PreviewKeyboardUnicode(ByVal a As String)
SpreadKey a
End Sub

Private Sub glistN_CheckGotFocus()
OneClick
End Sub

Private Sub glistN_CheckLostFocus()
On Error Resume Next
If Not moveMe Then
                If LastActive <> "" Then

                    If Controls(LastActive).enabled Then
                    If Controls(LastActive).Visible Then Controls(LastActive).SetFocus
                    End If
                End If
                Else
               If IsWine Then If glistN.Visible Then gList2.SetFocus
                End If
End Sub

Private Sub glistN_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If item = -1 Then

Else
glistN.mousepointer = 1
If lastitem = item Then Exit Sub
If glistN.ListSep(item) Then Exit Sub
glistN.ListindexPrivateUse = item
glistN.ShowMe2
lastitem = item
'glistN.ListindexPrivateUse = -1
End If
End Sub

Private Sub glistN_Fkey(a As Integer)
If a > 1000 Then
SendFKEY a - 1000
Else
SendFKEY a
End If
End Sub

Private Sub glistN_KeyDown(keycode As Integer, shift As Integer)
If keycode = vbKeyLeft Or keycode = vbKeyRight Then

keycode = 0

Pad.Visible = False
ElseIf keycode = 9 Then
keycode = 0
Pad.Visible = False
End If
End Sub

Private Sub glistN_ScrollMove(item As Long)
OneClick
End Sub

Private Sub mDoc_MayQuit(Yes As Variant)
If mQuit Or Not Visible Then Yes = True
MyDoEvents1 Me
End Sub
Private Sub ResizeMark_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If Sizable And Not dr Then
    x = x + ResizeMark.Left
    y = y + ResizeMark.top
    If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
    
    dr = Button = 1
    ResizeMark.mousepointer = vbSizeNWSE
    Lx = x
    lY = y
    If dr Then Exit Sub
    
    End If
    
End If
End Sub

Private Sub ResizeMark_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim addy As Single, addX As Single
If Not Relax Then
    x = x + ResizeMark.Left
    y = y + ResizeMark.top
    If Button = 0 Then If dr Then Me.mousepointer = 0: dr = False: Relax = False: Exit Sub
    Relax = True
    If dr Then
         If y < (Height - 150) Or y >= Height Then addy = (y - lY) Else addy = dv15 * 5
         If x < (Width - 150) Or x >= Width Then addX = (x - Lx) Else addX = dv15 * 5
         If Width + addX >= 1800 And Width + addX < VirtualScreenWidth() Then
             If Height + addy >= 1800 And Height + addy < VirtualScreenHeight() Then
                Lx = x
                lY = y
                move Left, top, Width + addX, Height + addy
                IhaveLastPos = False
                If mIndex > -1 Then
                    Callback mMyName$ + ".Resize(" + CStr(index) + ")"
                Else
                    Callback mMyName$ + ".Resize()"
                End If
                ResizeMark.backcolor = GetPixel(Me.Hdc, Width \ dv15 - 1, Height \ dv15 - 1)
            End If
        End If
        Relax = False
        Exit Sub
    Else
        If Sizable Then
            If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
                    dr = Button = 1
                    ResizeMark.mousepointer = vbSizeNWSE
                    Lx = x
                    lY = y
                    If dr Then Relax = False: Exit Sub
                Else
                    ResizeMark.mousepointer = 0
                    dr = 0
                End If
            End If
    End If
Relax = False
End If
End Sub

Public Property Get Sizable() As Variant
Sizable = mSizable
End Property

Public Property Let Sizable(ByVal vNewValue As Variant)
mSizable = vNewValue
ResizeMark.enabled = vNewValue
If ResizeMark.enabled Then
minimPos = 1
ResizeMark.Visible = Me.Visible
Else
minimPos = 0
ResizeMark.Visible = False
End If
End Property
Public Property Let SizerWidth(ByVal vNewValue As Variant)
If vNewValue \ dv15 > 1 Then
    MarkSize = vNewValue \ dv15
    With ResizeMark
    .Width = MarkSize * dv15
    .Height = MarkSize * dv15
    .move Width - .Width, Height - .Height
    End With
End If
End Property

Public Property Get Header() As Variant
Header = gList2.Visible
End Property

Public Property Let Header(ByVal vNewValue As Variant)
gList2.Visible = vNewValue
End Property


Sub GetFocus()
On Error Resume Next
If Me.Visible Then Me.SetFocus
End Sub
Public Sub UNhookMe()
Set LastGlist = Nothing
UnHook hWnd
End Sub

Public Property Get Quit() As Variant
Quit = mQuit
End Property

Public Property Let Quit(ByVal vNewValue As Variant)
mQuit = vNewValue
End Property

Public Property Get ShowMaximize() As Variant
ShowMaximize = mShowMaximize
End Property

Public Property Let ShowMaximize(ByVal vNewValue As Variant)
mShowMaximize = vNewValue
If gList2.Visible Then gList2.ShowMe
End Property
Friend Sub MinimizeOff()
           If MyForm3 Is Nothing Then Exit Sub
           If Not MyForm3.WindowState = 0 Then
            MyForm3.skiptimer = True
            MyForm3.WindowState = 0
           End If
End Sub
Friend Sub MinimizeON()
           If MyForm3 Is Nothing Then Exit Sub
           If Not MyForm3.WindowState = 1 Then
           MyForm3.skiptimer = True
           MyForm3.WindowState = 1
           End If
End Sub
Private Sub glistN_PanLeftRight(direction As Boolean)
Dim item As Long
On Error Resume Next
If direction = True Then
item = glistN.ListIndex

If glistN.ListSep(item) Then Exit Sub
If glistN.ListRadio(item) Then
    glistN.ListSelected(item) = True
    
End If

    If LastActive <> "" Then
        If EnableStandardInfo And item = 5 Then
        
            If gList2.Visible Then gList2.SetFocus
        Else
            If Controls(LastActive).enabled Then
            If Controls(LastActive).Visible Then Controls(LastActive).SetFocus
            End If
            LastActive = vbNullString
            End If
        Else
    If gList2.Visible Then gList2.SetFocus
        End If
If EnableStandardInfo Then
Select Case item
Case 1
' standard event
Case 2
    Minimize
Case 3
    If Maximize(True) Then
            If mIndex > -1 Then
           Callback mMyName$ + ".Resize(" + CStr(index) + ")"
        Else
           Callback mMyName$ + ".Resize()"
        End If
        ResizeMark.backcolor = GetPixel(Me.Hdc, Width \ dv15 - 1, Height \ dv15 - 1)
    MenuSet 2
    End If
Case 4
    If Maximize(False) Then
            If mIndex > -1 Then
           Callback mMyName$ + ".Resize(" + CStr(index) + ")"
        Else
           Callback mMyName$ + ".Resize()"
        End If
        
    MenuSet 3
    End If
Case 5
    MoveByKeyboard
    Blink = 50
    BlinkTimes = 10
    Exit Sub
Case 6
    SizeByKeyboard
    Blink = 50
    BlinkTimes = 10
    Exit Sub
    
Case glistN.listcount - 1
    ByeBye
    Exit Sub
End Select
End If
   If mIndex >= 0 Then
      Callback mMyName$ + ".InfoClick(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback mMyName$ + ".InfoClick(" + Str(item) + ")"
      End If
      
  
End If
End Sub

Private Sub glistN_Selected2(item As Long)
On Error Resume Next

If glistN.ListSep(item) Then Exit Sub
If item >= 0 Then

    If glistN.ListSep(item) Then Exit Sub
        If glistN.ListRadio(item) Then
            glistN.ListSelected(item) = True
        End If
    Pad.UNhookMe
    
    If LastActive <> "" Then
        If EnableStandardInfo And (item = 5 Or item = 6) Then
            If gList2.Visible Then gList2.SetFocus
        Else
            If Controls(LastActive).enabled Then
            If Controls(LastActive).Visible Then Controls(LastActive).SetFocus
            End If
            LastActive = vbNullString
            End If
        Else
    If gList2.Visible Then gList2.SetFocus
        End If
    If EnableStandardInfo Then
Select Case item
Case 1
' standard event
Case 2
    Minimize
Case 3
    If Maximize(True) Then
            If mIndex > -1 Then
           Callback mMyName$ + ".Resize(" + CStr(index) + ")"
        Else
           Callback mMyName$ + ".Resize()"
        End If
    MenuSet 2
    End If
Case 4
    If Maximize(False) Then
            If mIndex > -1 Then
           Callback mMyName$ + ".Resize(" + CStr(index) + ")"
        Else
           Callback mMyName$ + ".Resize()"
        End If
    MenuSet 3
    End If
Case 5
    MoveByKeyboard
    Blink = 50
    BlinkTimes = 10
    Exit Sub
Case 6
    SizeByKeyboard
    Blink = 50
    BlinkTimes = 10
    Exit Sub

Case glistN.listcount - 1
    ByeBye
    Exit Sub
End Select
End If
   If mIndex >= 0 Then
      Callback mMyName$ + ".InfoClick(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback mMyName$ + ".InfoClick(" + Str(item) + ")"
      End If
    
Else
    Pad.Visible = False
End If
End Sub


Public Sub additem(a$)
If Not NoEventInfo Then Exit Sub
glistN.additemFast a$
End Sub

Public Sub additemFast(a$)
If Not NoEventInfo Then Exit Sub
glistN.additemFast a$
End Sub
Public Property Get MenuWidth() As Long
MenuWidth = mMenuWidth
End Property

Public Property Let MenuWidth(ByVal RHS As Long)
 mMenuWidth = Abs(RHS)
 If mMenuWidth < 3000 Then mMenuWidth = 3000
End Property
Property Let menuEnabled(item As Long, ByVal RHS As Boolean)
If Not NoEventInfo Then Exit Property
glistN.menuEnabled(item) = RHS
End Property
Public Property Let Mark(item)
If Not NoEventInfo Then Exit Property
glistN.dcolor = mycolor(item)
End Property
Property Get menuEnabled(item As Long) As Boolean
If Not NoEventInfo Then Exit Property
menuEnabled = Not glistN.ListSep(item)
End Property
Public Sub Remove(item)
If Not NoEventInfo Then Exit Sub
On Error Resume Next
If item < 0 Then Exit Sub
glistN.Removeitem item
End Sub
Public Sub Insert(item, a$)
If Not NoEventInfo Then Exit Sub
On Error Resume Next
glistN.ListindexPrivateUse = item
If glistN.ListIndex > -1 Then
glistN.additemAtListIndex a$
End If
End Sub
Public Sub MenuItemAtListIndex(Optional enabledthis As Boolean = True, Optional checked As Boolean = False, Optional radiobutton As Boolean = False, Optional firstate As Boolean = False, Optional IdD)
If Not NoEventInfo Then Exit Sub
Dim item
item = glistN.ListIndex
If item < 0 Then Exit Sub
If IsMissing(IdD) Then
glistN.MenuItem item, checked, radiobutton, firstate

Else
glistN.MenuItem item, checked, radiobutton, firstate, CStr(IdD)
End If
glistN.menuEnabled(CLng(item - 1)) = enabledthis
End Sub
Public Sub MenuItem(a$, Optional enabledthis As Boolean = True, Optional checked As Boolean = False, Optional radiobutton As Boolean = False, Optional firstate As Boolean = False, Optional IdD)
If Not NoEventInfo Then Exit Sub
Dim item
If Not a$ = vbNullString Then
glistN.additemFast a$
End If
item = glistN.listcount
If a$ = vbNullString Then
glistN.AddSep
Else
If IsMissing(IdD) Then
glistN.MenuItem item, checked, radiobutton, firstate

Else
glistN.MenuItem item, checked, radiobutton, firstate, CStr(IdD)
End If
glistN.menuEnabled(CLng(item - 1)) = enabledthis
End If
End Sub
Public Sub MenuRadio(a$, Optional enabledthis As Boolean = True, Optional firstate As Boolean = False, Optional IdD)
If Not NoEventInfo Then Exit Sub
Dim item, checked As Boolean
checked = False

If Not a$ = vbNullString Then
glistN.additemFast a$
End If
item = glistN.listcount
If a$ = vbNullString Then
glistN.AddSep
Else
If IsMissing(IdD) Then
glistN.MenuItem item, True, True, False

Else
glistN.MenuItem item, True, True, False, CStr(IdD)
End If
If firstate Then glistN.ListSelectedNoRadioCare(CLng(item - 1)) = True
glistN.menuEnabled(CLng(item - 1)) = enabledthis
End If
End Sub

Property Let ListRadioPrivate(item As Long, RHS As Boolean)
If Not NoEventInfo Then Exit Property
glistN.ListSelectedNoRadioCare(item) = RHS
End Property
Property Get ListSelected(item As Long) As Boolean
If Not NoEventInfo Then Exit Property
ListSelected = glistN.ListSelected(item)
End Property
Property Let ListSelected(item As Long, RHS As Boolean)
If Not NoEventInfo Then Exit Property
glistN.ListSelected(item) = RHS
End Property
Property Get ListChecked(item As Long) As Boolean
If Not NoEventInfo Then Exit Property
ListChecked = glistN.ListChecked(item)
End Property
Property Let ListChecked(item As Long, RHS As Boolean)
If Not NoEventInfo Then Exit Property
glistN.ListChecked(item) = RHS
End Property
Property Get ListMenu(item As Long) As Boolean
If Not NoEventInfo Then Exit Property
ListMenu = glistN.ListMenu(item)
End Property

Property Get ListRadio(item As Long) As Boolean
If Not NoEventInfo Then Exit Property
ListRadio = glistN.ListRadio(item)
End Property
Property Let ListRadio(item As Long, RHS As Boolean)
If Not NoEventInfo Then Exit Property
glistN.ListRadio(item) = RHS
End Property
Property Get ListSep(item As Long) As Boolean
If Not NoEventInfo Then Exit Property
ListSep = glistN.ListSep(item)
End Property
Property Let ListSep(item As Long, RHS As Boolean)
If Not NoEventInfo Then Exit Property
glistN.ListSep(item) = RHS
End Property

Sub MakeInfo(ByVal RHS As Long)
 NoEventInfo = True
 Dim PadGui As New GuiM2000
 Set Pad = PadGui
 On Error Resume Next
 Set glistN = Pad.Controls(1)
 glistN.Arrows2Tab = False
If EnableStandardInfo Then
    glistN.Clear
   
    EnableStandardInfo = False
End If
 UseInfo = True
With glistN
    .addpixels = 4
    .FontSize = Me.CtrlFontSize
    .FontBold = True
    .backcolor = rgb(255, 255, 255)
    .forecolor = 0
    .enabled = True
    .FreeMouse = True
    .NoPanLeft = True
    .NoPanRight = False
    .SingleLineSlide = True
    .LeaveonChoose = True
    .LeftMarginPixels = 8
    .VerticalCenterText = True
    .ShowBar = False
    .StickBar = False ' True ' try with false - or hold shift to engage false
    .NoFreeMoveUpDown = True
     .CapColor = gList2.CapColor
    .dcolor = rgb(200, 200, 200)
    .BorderStyle = 1
 End With
 If Err.Number > 0 Then
 Set glistN = Pad.Controls(1)
 End If
 If RHS < 0 Then
    RHS = -RHS
    If RHS > 80 Then RHS = 80
    mMenuWidth = CLng(glistN.UserControlTextWidth("W") * RHS * 1.3)
 Else
    mMenuWidth = Abs(RHS)
    If mMenuWidth < 3000 Then mMenuWidth = 3000
 End If
 PadGui.PopUpMenuVal = True
 PadGui.NoHook = True
 With Pad
    .gList2.HeadLine = vbNullString
    .gList2.HeadLine = vbNullString
    .gList2.HeadlineHeight = .gList2.HeightPixels
End With
glistN.Dynamic
End Sub
Public Sub Minimize()
If IsWine Then Exit Sub
If Minimized Then Exit Sub
If MyForm3 Is Nothing Then Exit Sub
On Error Resume Next
If gList2.Visible Then gList2.SetFocus
MyForm3.Timer1.enabled = False
MyForm3.Timer1.Interval = 20
'MyForm3.Timer1.enabled = True
MyForm3.WindowState = 1
End Sub
Public Function Maximize(what As Boolean) As Boolean
On Error Resume Next
If Not mSizable Then Exit Function
If IhaveLastPos And Not what Then
           
            With ScrInfo(FindMonitorFromMouse)
            If IsWine And .Left = 0 And .top = 0 And .Width - 1 = Width And .Height - 1 = Height And (.Left <> Left Or .top <> top) Then
                Me.move .Left, .top
                Exit Function
            ElseIf .Width = Width And .Height = Height And (.Left <> Left Or .top <> top) Then
                Me.move .Left, .top
                Exit Function
            Else
                IhaveLastPos = False
                If IsWine And .Left = 0 And .top = 0 Then
                    Width = MeWidth
                    Height = MeHeight
                    Me.move MeLeft, MeTop
                Else
                    Me.move MeLeft, MeTop, MeWidth, MeHeight
                End If
            End If

            End With
    Maximize = True
ElseIf what And Not IhaveLastPos Then
    IhaveLastPos = True
    MeLeft = Left
    MeTop = top
    MeWidth = Width
    MeHeight = Height
    On Error Resume Next
    With ScrInfo(FindMonitorFromMouse)
                If IsWine And .Left = 0 And .top = 0 Then
                    Width = .Width - 1
                    Height = .Height - 1
                    move .Left, .top
                Else
                move .Left, .top, .Width, .Height
                End If
    End With
    Maximize = True
End If
End Function
Public Sub MoveByKeyboard()
If gList2.Visible Then
StoreBlinkStatus
movemeX = MOUSEX()
movemeY = MOUSEY
gList2.FloatListMe False, movemeX, movemeY
moveMe = True
End If
End Sub
Public Sub SizeByKeyboard()
If gList2.Visible Then
StoreBlinkStatus
movemeX = Width
movemeY = Height
moveMe = True
sizeMe = True
End If
End Sub
Public Property Get UseIcon() As Variant
    UseIcon = mIcon
End Property

Public Property Let UseIcon(ByVal vNewValue As Variant)
    If mIcon = CBool(vNewValue) Then Exit Property
    mIcon = CBool(vNewValue)
    If mIcon Then infopos = 1 Else infopos = 0
    gList2.ShowMe
End Property
Private Sub drawicon(HDC1 As Long)
Dim picthis As StdPicture, my_brush As Long, msize As Long

msize = setupxy
If Me.icon Is Nothing Then
    Set picthis = Form1.icon
Else
    Set picthis = icon
End If
my_brush = CreateSolidBrush(gList2.CapColor)
If IsWine Then
If UseReverse Then
gList2.PaintPicture1 picthis, 0, 0, msize, msize
gList2.PaintPicture1 picthis, 0, 0, msize, msize
Else
gList2.PaintPicture1 picthis, gList2.WidthPixels - msize, 0, msize, msize
gList2.PaintPicture1 picthis, gList2.WidthPixels - msize, 0, msize, msize

End If
Else
If Not UseReverse Then
If Not gList2.NoHeaderBackground Then
DrawIconEx HDC1, gList2.WidthPixels - msize, 0, picthis, msize, msize, 0, my_brush, 1
Else
DrawIconEx HDC1, gList2.WidthPixels - msize, 0, picthis, msize, msize, 0, 0, &H3
End If
DrawIconEx HDC1, gList2.WidthPixels - msize, 0, picthis, msize, msize, 0, 0, &H3

Else
If Not gList2.NoHeaderBackground Then
DrawIconEx HDC1, 0, 0, picthis, msize, msize, 0, my_brush, 1
Else
DrawIconEx HDC1, 0, 0, picthis, msize, msize, 0, 0, &H3
End If
DrawIconEx HDC1, 0, 0, picthis, msize, msize, 0, 0, &H3
End If
End If
DeleteObject my_brush
End Sub
Public Function GetPicture(ByVal s$, Optional Size, Optional ColorDepth, Optional x, Optional y) As StdPicture
Dim where$
On Error Resume Next
where$ = CFname(s$)
If LenB(where$) = 0 Then
    Set GetPicture = Form1.icon
Else
where$ = GetDosPath(where$)
If LenB(where$) = 0 Then Set GetPicture = LoadPicture(""): Exit Function
If IsMissing(Size) And IsMissing(ColorDepth) And IsMissing(x) And IsMissing(y) Then
Set GetPicture = LoadPicture(where$)
ElseIf IsMissing(ColorDepth) And IsMissing(x) And IsMissing(y) Then
Set GetPicture = LoadPicture(where$, Size)
ElseIf IsMissing(x) And IsMissing(y) Then
Set GetPicture = LoadPicture(where$, Size, ColorDepth)
ElseIf IsMissing(y) Then
Set GetPicture = LoadPicture(where$, Size, ColorDepth, x)
Else
Set GetPicture = LoadPicture(where$, Size, ColorDepth, x, y)
End If
End If
End Function
Public Sub ReloadIcon(ByVal s$, Optional Size, Optional ColorDepth, Optional x, Optional y)
Dim where$
On Error Resume Next
where$ = CFname(s$)
If LenB(where$) = 0 Then
    Set icon = Form1.icon
Else
where$ = GetDosPath(where$)
    If LenB(where$) = 0 Then Exit Sub
If IsMissing(Size) And IsMissing(ColorDepth) And IsMissing(x) And IsMissing(y) Then
Set icon = LoadPicture(where$)
ElseIf IsMissing(ColorDepth) And IsMissing(x) And IsMissing(y) Then
Set icon = LoadPicture(where$, Size)
ElseIf IsMissing(x) And IsMissing(y) Then
Set icon = LoadPicture(where$, Size, ColorDepth)
ElseIf IsMissing(y) Then
Set icon = LoadPicture(where$, Size, ColorDepth, x)
Else
Set icon = LoadPicture(where$, Size, ColorDepth, x, y)
End If
End If
If MyForm3 Is Nothing Then Exit Sub
Set MyForm3.icon = icon
End Sub
Public Sub MakeStandardInfo(RHS)
MakeInfo -7
Dim item
If RHS > 0 Then
    MenuItem "About", True
    MenuItem ""
    MenuItem "Minimize", Not IsWine And Not IamPopUp
    MenuItem "Maximize", mSizable
    MenuItem "Restore", False
    MenuItem "Move", True
    MenuItem "Size", mSizable
    MenuItem ""
    MenuItem "Close", True
Else
    MenuItem "”˜ÂÙÈÍ‹", True
    MenuItem ""
    MenuItem "¡¸ÍÒı¯Á", Not IsWine And Not IamPopUp
    MenuItem "≈›ÍÙ·ÛÁ", mSizable
    MenuItem "≈·Ì·ˆÔÒ‹", False
    MenuItem "ÃÂÙ·ÍﬂÌÁÛÁ", True
    MenuItem "Ã›„ÂËÔÚ", mSizable
    MenuItem ""
    MenuItem " ÎÂﬂÛÈÏÔ", True
End If
EnableStandardInfo = True
End Sub
Private Sub MenuSet(RHS)
If Not EnableStandardInfo Then Exit Sub
Select Case RHS
Case 2
 glistN.menuEnabled(3) = False: glistN.menuEnabled(4) = True
Case 3
 glistN.menuEnabled(3) = True: glistN.menuEnabled(4) = False
 End Select
End Sub
Sub InsertMenuItem(a$, Optional enabledthis As Boolean = True, Optional checked As Boolean = False, Optional radiobutton As Boolean = False, Optional firstate As Boolean = False, Optional IdD)
If Not NoEventInfo Then Exit Sub
Dim item
glistN.additemAtListIndex a$
item = glistN.ListIndex
If a$ = vbNullString Then
glistN.ListSep(item - 1) = True
Else
If IsMissing(IdD) Then
glistN.MenuItem item, checked, radiobutton, firstate

Else
glistN.MenuItem item, checked, radiobutton, firstate, CStr(IdD)
End If
glistN.menuEnabled(CLng(item - 1)) = enabledthis
End If
End Sub
Property Get NoFocus() As Boolean
NoFocus = AppNoFocus
End Property
Property Get LastControl() As String
If Not LastGlist Is Nothing Then
LastControl = LastActive + " True"
Else
LastControl = LastActive
End If
End Property
Property Get HookStatus() As Long
HookStatus = HOOKTEST
End Property
Property Let TitleBarColor(RHS As Long)
    mBarColor = mycolor(RHS)
    gList2.CapColor = mBarColor
End Property
Property Let TitleTextColor(RHS As Long)
    gList2.forecolor = mycolor(RHS)
End Property
Property Let TitleIconColor(RHS As Long)
    mIconColor = mycolor(RHS)
End Property

Sub TransparentTitle()
Dim x  As Long, y As Long
gList2.NoHeaderBackground = True
gList2.BackStyle = 1
gList2.GetLeftTop x, y
gList2.RepaintFromOut Me.Image, x, y
gList2.ShowMe
RefreshList = RefreshList + 1
End Sub
Sub OpaqueTtile()
On Error Resume Next
If Not glistN Is Nothing Then
gList2.NoHeaderBackground = False
gList2.BackStyle = 0
gList2.backcolor = 0
gList2.PanPos = 0
gList2.ShowMe
RefreshList = RefreshList - 1
End If
End Sub
Private Sub OneClick()
On Error Resume Next
If SkipFirstClick Then glistN.PrepareClick
End Sub

Friend Property Get modulename() As String
modulename = PRmodulename$
End Property

Friend Property Let modulename(ByVal RHS As String)
PRmodulename$ = RHS
End Property
Friend Sub RegisterAcc(m, ControlName$, Optional Opcode As Long = 0)
If acclist Is Nothing Then Set acclist = New FastCollection
If acclist.ExistKey(m) Then acclist.RemoveWithNoFind
acclist.AddKey m, ControlName$
acclist.sValue = Opcode
End Sub
Friend Sub AccProces(m As Long)
Dim todo As Long
On Error Resume Next
If Not acclist Is Nothing Then
    If acclist.ExistKey(m) Then
        If Controls(acclist.Value).enabled Then
            
            todo = acclist.sValue
            If todo = 0 Then
                Controls(acclist.Value).SetFocus
            ElseIf todo < 0 Then
                ' CALL PRESS MENU ITEM -TODO
                Controls(acclist.Value).CascadeSelect -todo
            ElseIf todo = 1 Then
                Controls(acclist.Value).SetFocus
                Controls(acclist.Value).PressSoft
            ElseIf todo = 2 Then
                Controls(acclist.Value).PressSoft
            Else
                Dim shift As Long, ctrl As Long, alt As Long
                If todo Mod 2 = 1 Then Controls(acclist.Value).SetFocus
                todo = todo \ 2
                shift = Abs(((todo \ 1000) Mod 10) <> 0)
                ctrl = Abs(((todo \ 10000) Mod 10) <> 0) * 2
                alt = Abs(((todo \ 100000) Mod 10) <> 0) * 4
                
                todo = todo Mod 1000
                Dim a As gList
                Set a = Controls(acclist.Value)
                a.TakeKey CInt(todo), shift + ctrl + alt
            End If
            m = 0
            Exit Sub
        Else
        'Debug.Print "NOT ENABLED"
        End If
        Else
        'Debug.Print "NOT EXIST"
    End If
End If
End Sub
Public Sub AccKey(a, Optional shift As Boolean, Optional ctrl As Boolean, Optional alt As Boolean, Optional Opcode As Long = 0)

If MyIsNumeric(a) Then
a = CLng(a)
If a < 0 Then Exit Sub
If a > 499 Then Exit Sub
Else
a = UCase(a)
Select Case a
Case "F1" To "F9"
a = 611 + val(Mid(a, 2))
Case Else
a = AscW(a)
If a > 126 Then a = 0
End Select
End If
If a = 0 Then Exit Sub
a = a - 1000 * shift - 10000 * ctrl - 100000 * alt
RegisterAcc a, "gList2", Opcode
End Sub
Public Sub SendScanCode(a, Optional shift As Boolean, Optional ctrl As Boolean, Optional alt As Boolean)
' send only scancodes if a<500 or a-500 for a>500 with extend option. F4 is 615 (500+115)
If MyIsNumeric(a) Then
a = CLng(a)
If a < 0 Then Exit Sub
If a > 754 Then Exit Sub
Else
a = UCase(a)
Select Case a
Case "F1" To "F9"
a = 611 + val(Mid(a, 2))
Case Else
a = AscW(a)
If a > 254 Then a = 0
End Select
End If
If a = 0 Then Exit Sub
SendAKey a, shift, ctrl, alt
End Sub
Property Get CapsLock() As Boolean
    CapsLock = CapsLockOn()
End Property
