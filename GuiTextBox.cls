VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' If you copy and paste code
' don't forget to set default property to vartext (changed from text)
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Private mprompt As String
Private mark1 As Long, mark2 As Long
Private mkind As String, mPartColor As Long
Public SkipProcAtExitFocus As Boolean, MoveOnEnter As Boolean
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_DISCARDED = &H4000
Private Const GMEM_FIXED = &H0
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GMEM_LOCKCOUNT = &HFF
Private Const GMEM_MODIFY = &H80
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetClipboardData Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Declare Function SetClipboardData Lib "user32" _
    (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function OpenClipboard Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" _
    () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Pad$
    ' for Gui purposes
Public ctrlName As String
Private mesCtrlName As String, sp As String, sp1 As String, sp2 As String
Dim mIndex As Long, mStayOpen As Boolean
Dim Callback As GuiM2000
Dim white As Long, black As Long, dimm1 As Long, dimm2 As Long
Dim tParam As DRAWTEXTPARAMS
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Function GetStringTypeExW Lib "Kernel32.dll" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByRef lpCharType As Integer) As Long
Public NumberOnly As Boolean
Public NumberIntOnly As Boolean
Friend Property Get GetCallBack() As GuiM2000
Set GetCallBack = Callback
End Property
Friend Sub Construct(there As GuiM2000, Name$)
 mIndex = -1
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set glistN = there.Controls.Add("m2000.glist", ctrlName)
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
Friend Sub ConstructArray(there As GuiM2000, Name$, i As Long)
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set glistN = there.Controls.Add("m2000.glist", ctrlName)
 mIndex = i
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
Public Property Get Text() As String
If glistN.listcount = 0 Then
glistN.AddItem mprompt
Else
Text = glistN.list(0)
End If
End Property
Public Property Get vartext() As String
If glistN.listcount = 0 Then
glistN.AddItem ""
Else
vartext = Mid$(glistN.list(0), 1 + Len(mprompt))
End If
End Property
Public Property Let vartext(ByVal RHS As String)
glistN.restrictLines = 1
glistN.HeadLine = vbNullString ' no info
glistN.HeadlineHeight = 0

If glistN.listcount = 0 Then
glistN.AddItem mprompt + RHS
Else
glistN.list(0) = mprompt + RHS
End If
glistN.SelStart = Len(mprompt + RHS) + 1
glistN.ShowMe
End Property
Public Property Let Text(ByVal RHS As String)
'Dim oldright As Long
'oldright = glistN.LeftMarginPixels
On Error Resume Next
If glistN Is Nothing Then Exit Property
glistN.restrictLines = 1
glistN.HeadLine = vbNullString ' no info
glistN.HeadlineHeight = 0
If glistN.listcount = 0 Then
glistN.AddItem RHS
Else
glistN.list(0) = RHS
End If
'glistN.LeftMarginPixels = oldright

glistN.ShowMe2
End Property
Public Property Get Info() As String
Info = glistN.HeadLine
End Property

Public Property Let Info(ByVal RHS As String)
locked = False
If RHS = vbNullString Then
glistN.InfoDropBarClick = False
Text = Text
Else
glistN.restrictLines = 0
glistN.HeadLineHeightMinimum = glistN.HeightPixels
glistN.HeadLine = RHS
glistN.InfoDropBarClick = True
glistN.ShowMe
locked = True
End If

End Property

Public Property Get top() As Variant
top = Callback.Controls(ctrlName).top
End Property
Public Property Let top(ByVal RHS As Variant)
Callback.Controls(ctrlName).top = RHS
End Property
Public Property Get Left() As Variant
Left = Callback.Controls(ctrlName).Left
End Property

Public Property Let Left(ByVal RHS As Variant)
Callback.Controls(ctrlName).Left = RHS
End Property


Public Sub SetUp()
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
Dim ff As New StdFont
With Callback
white = .Controls(1).ForeColor
ff.Name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With
MoveOnEnter = True
With glistN
.NoEscapeKey = True
.bypassfirstClick = True
.Enabled = True
Info = vbNullString
If .listcount = 0 Then Text = vbNullString
.NoCaretShow = True
.VerticalCenterText = True
If Not .CenterText Then .LeftMarginPixels = 8 Else .LeftMarginPixels = 0
.BackColor = rgb(200, 120, 60)
Set .Font = ff
.ForeColor = white
.dcolor = Callback.Controls(1).dcolor
.ShowMe
End With
locked = False
Enabled = True
Retired

End Sub


Public Property Get Prompt() As String
Prompt = mprompt
End Property

Public Property Let Prompt(ByVal RHS As String)
mprompt = RHS
glistN.PromptLineIdent = Len(RHS)
Text = RHS
End Property


Private Sub Class_Initialize()
sp1 = ",+-*/=<>^!.&[]{}()^;""" + "'!"
sp = " " + ChrW(9) + ChrW(160)
sp2 = sp1 + sp
End Sub

Private Sub Class_Terminate()
Set glistN = Nothing
End Sub

Private Sub glistN_CaretDeal(Deal As Long)

If Deal = 0 Then
    glistN.Noflashingcaret = True
    glistN.NoCaretShow = True
    glistN.ShowMe2
    Deal = 1
Else
   glistN.Noflashingcaret = False
    glistN.NoCaretShow = False
    glistN.ShowMe2
    Deal = 0
End If
End Sub


Private Sub glistN_Fkey(a As Integer)
If a > 1000 Then
Callback.SendFKEY a - 1000
Else
Callback.SendFKEY a
End If
End Sub

Private Sub glistN_GetRealX1(mHdc As Long, ByVal ExtSelStart As Long, ByVal that As String, retvalue As Long, ok As Boolean)
retvalue = RealX1(mHdc, ExtSelStart, that)
ok = True
End Sub

Private Sub glistN_OutPopUp(x As Single, y As Single, myButton As Integer)
If (myButton And &H2) = 2 Then
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".PopUp(" + Str(mIndex) + "," + Str(x + Callback.Controls(ctrlName).Left) + "," + Str(y + Callback.Controls(ctrlName).top) + "," + Str(myButton) + "," + Str(glistN.ListIndex + 1) + ")"
   Else
      Callback.Callback mesCtrlName + ".PopUp(" + Str(x + Callback.Controls(ctrlName).Left) + "," + Str(y + Callback.Controls(ctrlName).top) + "," + Str(myButton) + "," + Str(glistN.ListIndex + 1) + ")"
      End If
      End If
End Sub

Private Sub glistN_RegisterGlist(this As gList)
If Not glistN.Spinner Then this.NoWheel = True
On Error Resume Next
Set LastGlist = this
If Err.Number > 0 Then this.NoWheel = True
End Sub



Private Sub glistN_ChangeListItem(item As Long, content As String)
'If Not enabled Then Exit Sub
Dim this As String, thispos As Long, VR(2)
thispos = glistN.SelStart - Len(mprompt)
this = Mid$(content, 1 + Len(mprompt))
If NumberOnly Then
    If NLtrim$(this) = vbNullString Then
    
    ElseIf Not ValidNum(this, False, NumberIntOnly) Then
        this = Mid$(glistN.list(item), 1 + Len(mprompt))
    End If
End If

If Len(this) < thispos Then thispos = Len(this) + 1
'' RaiseEvent ValidString(this, thispos)
VR(0) = this
VR(1) = thispos
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".ValidString(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".ValidString()", VR()
      End If
      
thispos = VR(1) + Len(mprompt)
content = mprompt + VR(0)
If thispos <> glistN.SelStart And VR(0) <> "" Then glistN.SelStart = thispos: glistN.ShowMe
End Sub
Private Sub glistN_CtrlPlusF1()
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
End Sub
Public Sub ResetPan()
glistN.PanPos = 0
End Sub

Private Sub glistN_ChangeSelStart(thisselstart As Long)
If thisselstart <= Len(mprompt) Then thisselstart = Len(mprompt) + 1
'glistN.ShowPan
With glistN
'.ShowPan
.ShowMe2
'.ExternalCursor thisselstart, .list(.ListIndex), .forecolor
End With
End Sub

Public Property Get ThisKind() As String
ThisKind = mkind
End Property

Public Property Let ThisKind(ByVal RHS As String)
mkind = RHS
glistN.LastLinePart = " " & RHS
End Property


Public Property Get Enabled() As Boolean
Enabled = glistN.EditFlag
End Property

Public Property Let Enabled(ByVal RHS As Boolean)
glistN.EditFlag = RHS
glistN.Arrows2Tab = Not RHS
glistN.NoCaretShow = Not RHS
glistN.NoBarClick = Not RHS
End Property
'


Private Sub glistN_CheckGotFocus()
Callback.RealHover = ctrlName
If Not locked Then
If glistN.ListIndex < 0 And glistN.listcount > 0 Then glistN.ListindexPrivateUse = 0
' RaiseEvent GotFocus
If Enabled Then
glistN.FadeLastLinePart = mPartColor
If glistN.BackStyle <> 1 Then
glistN.ForeColor = black
glistN.BackColor = white
End If
glistN.NoCaretShow = False:
glistN.PanPos = 0
Else
glistN.FadeLastLinePart = black
If glistN.BackStyle <> 1 Then
glistN.ForeColor = white
glistN.BackColor = dimm1
End If
glistN.NoCaretShow = True
End If
Callback.LastActive = ctrlName
glistN.ShowMe
End If

End Sub
Public Sub Retired()
If mStayOpen Then
glistN.FadeLastLinePart = mPartColor
If glistN.BackStyle <> 1 Then
glistN.ForeColor = black
glistN.BackColor = white
End If
glistN.NoCaretShow = False
Else
If glistN.BackStyle <> 1 Then
glistN.FadeLastLinePart = 0
glistN.ForeColor = white
glistN.BackColor = dimm2
End If
glistN.NoCaretShow = True
End If
 glistN.PanPos = 0: glistN.ShowMe
End Sub

Private Sub glistN_CheckLostFocus()
If Not locked Then
    Callback.LastActive = ctrlName
    If mStayOpen Then
        glistN.NoCaretShow = True
        glistN.HideTheCaret
        If glistN.Visible Then glistN.PanPos = 0: glistN.PrepareToShow
        Exit Sub
    End If
    If Not SkipProcAtExitFocus Then
        If glistN.BackStyle <> 1 Then
            glistN.ForeColor = white
            glistN.BackColor = dimm2
        End If
        glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
    End If
End If
' RaiseEvent LostFocus
End Sub

Private Sub glistN_HeaderSelected(Button As Integer)
' INFO SELECTED
Text = Text
End Sub

Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)
If MoveOnEnter Then
If KeyCode = 13 Then
If shift = 1 Then
glistN.Goback
Else
glistN.GoON
End If
Exit Sub
KeyCode = 0
End If
End If
Dim VR(2)
VR(0) = KeyCode
VR(1) = shift
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".KeyDown(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".KeyDown()", VR()
      End If
KeyCode = VR(0)
shift = VR(1)
If KeyCode = vbKeyF10 And shift <> 0 Then
KeyCode = 0
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".PopUp(" + Str(mIndex) + "," + Str(MOUSEX(Callback.Left)) + "," + Str(MOUSEY(Callback.top)) + "," + Str(2) + "," + Str(glistN.ListIndex + 1) + ")"
   Else
      Callback.Callback mesCtrlName + ".PopUp(" + Str(MOUSEX(Callback.Left)) + "," + Str(MOUSEY(Callback.top)) + "," + Str(2) + "," + Str(glistN.ListIndex + 1) + ")"
      End If
Exit Sub
End If
If KeyCode = vbKeyUp Then
ElseIf KeyCode = vbKeyDown Then
ElseIf KeyCode = vbKeyV And (shift And 3) = 2 Then
If Not Enabled Then Exit Sub
Pad$ = GetTextData(13)
If Pad$ = vbNullString Then Pad$ = Clipboard.GetText(1)
If Pad$ <> "" Then
If (glistN.maxchar - Len(mprompt)) > 0 Then
    vartext = Mid$(Pad$, 1, (glistN.maxchar - Len(mprompt)))

Else
    vartext = Pad$
    End If
End If
End If

End Sub

Public Property Get SelStart() As Long
SelStart = glistN.SelStart
End Property

Public Property Let SelStart(ByVal RHS As Long)
glistN.SelStart = RHS
End Property
Property Let Visible(ByVal RHS As Boolean)
    If Not glistN Is Nothing Then
        glistN.Visible = RHS
    End If
End Property
Property Get Visible() As Boolean
    If Not glistN Is Nothing Then
        Visible = glistN.Visible
    End If
End Property
Public Sub SetFocus()
On Error Resume Next
glistN.SetFocus
End Sub

Private Sub glistN_RTL(thisHDC As Long, item As Long, where As Long, mark10 As Long, mark20 As Long, offset As Long)
offset = RealX1(thisHDC, where, Text, where, mark10)
End Sub
Function RealX1(mHdc As Long, ByVal ExtSelStart As Long, ByVal that$, Optional nValue As Long = 0, Optional found As Long = 0) As Long
If ExtSelStart <= 0 Then ExtSelStart = 1
Dim ww$(), Col() As Long, items&, i, c As String
Dim rleftacc As Long, original As Long, rx1 As Long, direction As Long
Dim OldValue As Long, orx1 As Long, j As Long, nrtl As Boolean
If Len(that$) = 0 Then
    nValue = 1: Exit Function
End If

original = ExtSelStart

ReDim ww$(0)
ww$(0) = that$


Dim M1 As Long, M2 As Long
Dim dummy As RECT, R As RECT, tabw As Long, offset As Long, hRgn As Long, estX As Long
tabw = tParam.iTabLength * glistN.AveCharWith
glistN.UserControlTextMetricsPixels that$, R.Right, R.Bottom
dummy = R
dummy.top = glistN.HeadlineHeightTwips + 1
dummy.Bottom = dummy.top + R.Bottom
R = dummy
offset = 0
estX = 0
'hRgn = CreateRectRgnIndirect(dummy)
'SelectClipRgn mHdc, hRgn
If found > 0 Then
rx1 = LowWord(GetTabbedTextExtent(mHdc, StrPtr(that$), Len(that$), 1, tabw))
If rx1 < found Then

c = Right$(that$, 1)

nValue = 1
RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
If mark1 <> 0 And mark2 <> 0 Then
nValue = Len(that$)
RTL3 that$, nValue, mark1, mark2, nrtl, M1, M2
nValue = mark1
GoTo finish
Else
nValue = Len(ww(i))
RTL3 ww(i), nValue, mark1, mark2, nrtl, M1, M2
If mark1 <> 0 And mark2 <> 0 And nValue >= mark2 Then
If nValue <= mark2 + 1 Then
nValue = mark1
GoTo finish
Else
nValue = nValue + 1
ExtSelStart = nValue
End If
Else
nValue = Len(that$) + 1
GoTo finish
End If

End If
End If
End If
For i = 0 To items&
    c = ww$(i)
    R.Left = offset
    rleftacc = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
    If Len(c) >= ExtSelStart - estX Or (found > 0 And rleftacc > found) Then
            If Not Len(c) >= ExtSelStart - estX Then
            ExtSelStart = Len(c) + estX
            End If

        nValue = ExtSelStart - estX

        RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
        If nrtl Then
            If M2 - M1 > 0 Then Mid$(c, M1, M2 - M1 + 1) = StrReverse(Mid$(c, M1, M2 - M1 + 1))
        End If
        If mark1 = 0 Or mark2 = 0 Then
        If nValue > 1 Then
            j = nValue - 1
            RTL3 c, j, mark1, mark2, nrtl, M1, M2
            If mark1 <> 0 And mark2 <> 0 And j >= mark1 And j <= mark2 Then
            If i = items& Then
            If found > 0 Then
                rx1 = offset
                If mark1 > 1 Then
                    offset = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark1 - 1, 1, tabw, 0))
                End If
                c = Mid$(ww(i), mark1, mark2 - mark1 + 1)
                If nrtl Then
                    M1 = M1 - mark1 + 1
                    M2 = M2 - mark1 + 1
                    If M2 - M1 > 0 Then Mid$(c, M1, M2 - M1 + 1) = StrReverse(Mid$(c, M1, M2 - M1 + 1))
                End If
                If found > offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c$), 1, tabw, 0)) Then
                        nValue = mark1 + estX
                        Exit For
                    Else
                        offset = rx1
                        c = ww$(i)
                    End If
                nValue = j + estX
                End If
             GoTo here1

            Else
            nValue = mark2 - 1 + estX 'j - (nValue > mark2)
            If nValue <= 0 Then nValue = 1
            If mark1 > 1 Then
            
            rx1 = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark2, 1, tabw, 0))
            If found > 0 Then
            If rx1 <= found Then
            nValue = mark1 + estX
            RealX1 = rx1
            Exit For
            End If

            End If
            RealX1 = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark1 - 1, 1, tabw, 0))
            
            End If
            Exit For
            
            End If
            End If
        End If
        ElseIf mark1 <> 0 And mark2 <> 0 And nValue = mark2 + 1 Then
        If found > 0 Then
         rx1 = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark2, 1, tabw, 0))
        If found > rx1 Then
         nValue = mark1 + estX
            RealX1 = rx1
        Exit For
        End If
        nValue = mark2
        End If
        GoTo here1
       
        End If

        If mark1 <> 0 And mark2 <> 0 And nValue >= mark1 And nValue <= mark2 Then
here1:
            If mark1 > 1 Then
                
                offset = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark1 - 1, 1, tabw, 0))
            End If
            If mark1 = nValue And found > 0 Then

            If offset > found Then
            j = mark1 - 2
            If j >= 1 Then
            Do While found < R.Left + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j, 1, tabw, 0))
            j = j - 1
            If j = 0 Then j = 1: Exit Do
            Loop
            Else
                            
                GoTo nortl
            End If
            RTL3 c, j, mark1, mark2, nrtl, M1, M2
            If mark1 <> 0 And mark2 <> 0 Then
            nValue = mark2 - (j - mark1) + estX
            Else
            nValue = j + 1 + estX
            End If

            RealX1 = offset
            Exit For
            End If
            End If
            c = Mid$(ww(i), mark1, mark2 - mark1 + 1)
            If nrtl Then
                M1 = M1 - mark1 + 1
                M2 = M2 - mark1 + 1
                If M2 - M1 > 0 Then Mid$(c, M1, M2 - M1 + 1) = StrReverse(Mid$(c, M1, M2 - M1 + 1))
            End If
           
            If nValue = mark2 + 1 Then
            If found > 0 Then
            For j = mark2 - mark1 + 1 To 1 Step -1
            If found > offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j, 1, tabw, 0)) Then
            
            If j > 1 Then
            If LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j, 1, tabw, 0)) > LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j - 1, 1, tabw, 0)) < 0 Then
            
                        nValue = mark2 - (j) + estX

            offset = found
            Exit For
            End If
            End If

            nValue = mark2 - (j - 1) + estX

            offset = found
            Exit For

            End If
            Next j
            If j = 0 Then
            nValue = mark1 - 1 + estX - 1
            If nValue <= 0 Then nValue = 1
            End If
            Else
            nValue = mark1 - 1 + estX - 1
            If nValue <= 0 Then nValue = 1
            End If
            
          
             RealX1 = offset
         
            Exit For
            ElseIf found > 0 And mark1 = nValue Then
            If found > offset Then
             nValue = mark2
             
            GoTo conthere
            Else
            nValue = mark2 + 1 + estX
            RealX1 = offset
            End If
            
            
            Else
conthere:
            rleftacc = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
             If nValue > mark1 Then
                rx1 = rleftacc - 1 - LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), nValue - mark1, 1, tabw, 0))
            Else
                rx1 = rleftacc - 1
            End If
            End If
        If found > 0 Then
         rx1 = rx1 + 1
            If found < rx1 Then
                original = nValue
                For j = nValue To mark2

                rx1 = rleftacc - LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j - mark1 + 1, 1, tabw, 0))

                    
                    original = original + 1
                    If found >= rx1 Then Exit For
                Next j
                
                j = original
                If nrtl Then
                nValue = M2 - ((j + 1 - mark1) - M1) + mark1 + estX
 
                
                
                Else
                
                
                If j > mark2 Then
      
                  nValue = mark2 + 1 + estX
                  If nValue = 0 Then nValue = 1
                      
                          RealX1 = rx1
                          Exit For
                  Else
                  
                          nValue = j
                          RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
                          If nrtl And j >= M1 And j <= M2 Then
                          nValue = nValue + 2 + estX
                          Else
                          nValue = j + estX
                          End If
            
                  
                  End If
              
                
                    RealX1 = rx1 + 1
        End If
                Else
                  original = nValue
                  For j = nValue - 1 To mark1 + 1 Step -1
                
                      rx1 = rleftacc - LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j - mark1, 1, tabw, 0))
                      If found <= rx1 Then Exit For
                      original = j
                      
                  Next j
                    j = original
                If nrtl Then
                nValue = M2 - ((j + 1 - mark1) - M1) + mark1 + estX
              Else
                    
                    If j < mark1 Then
                       j = mark1
                    End If
                 
                 
                 
                    If nrtl And j - mark1 >= M1 And j - mark1 <= M2 Then
                    nValue = nValue + estX
                    Else
                    nValue = j + estX
                    End If
     End If

                    RealX1 = rx1

            End If
            
          If found = 0 Then
                If Mid$(that$, nValue, 1) = " " Then
         c = Mid$(that$, nValue + 1, 1)
         If Len(c) > 0 Then
         RTL3 c, (1), mark1, mark2, nrtl, M1, M2
         If mark1 = 0 Or mark2 = 0 Then
         nValue = nValue + 1
         'If nvalue1 < 1 Then nvalue1 = 1
         End If
         End If
         End If
         End If
      Exit For
     End If
            
            If found = 0 And nrtl Then
            'nValue = ExtSelStart
            Else
            nValue = mark2 - (nValue - mark1) + (nValue <> mark2) + estX - 1
            End If
            
If nValue = mark1 And items& = 1 Then
rx1 = 0
           
            
            End If
            If nValue = 0 Then nValue = mark2
       
            
            
            RealX1 = rx1
            
        Else ' NO RTL
nortl:
            c = Myleft$(c, ExtSelStart - estX - 1)
            rx1 = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
            If found > 0 Then
                If found < rx1 Then
                    nValue = ExtSelStart - 1
                    If nValue - estX - 1 < 1 Then
                        RealX1 = offset
                    Else
                        c = Myleft$(c, nValue - estX - 1)
                        j = Len(c)
                        RTL3 c, j, mark1, mark2, nrtl, M1, M2      ' found=0
                        If mark1 <> 0 And mark2 <> 0 Then
                        If j <= mark2 And j >= mark1 Then
                        nValue = mark2 - (j - mark1) + estX
                        Else
                        'nValue = nValue + estX - 1
                        End If
                        Else
                       ' nValue = Len(c) + estX
                        End If
                        rx1 = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
                    End If
                    RealX1 = rx1
                    Exit For
                Else
                    RealX1 = rx1
                    
                    nValue = original
                End If
            Else
                RealX1 = rx1
                nValue = ExtSelStart
            End If
            
            Exit For
        End If
  
         RealX1 = rx1
       ' nValue = ExtSelStart
        
        Exit For
    ElseIf i = items& Then
        nValue = ExtSelStart - estX
        RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
        If mark1 <> 0 And mark2 <> 0 Then
            If mark2 = Len(c) Then
                c = Myleft$(ww(i), mark1 - 1)
                If Len(c) = 0 Then
                    rleftacc = offset
                ElseIf nrtl And M2 = mark2 Then
                rleftacc = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
                  c = ww(i)
                  
               If M2 - M1 > 0 Then c = Mid$(c, M1, M2 - M1 + 1)
                 rleftacc = rleftacc + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
                nValue = M1 - 1
            Else
                 rleftacc = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
                End If
            ElseIf found > 0 Then
            rleftacc = offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), nValue, 1, tabw, 0))
        If rleftacc < found Then
        nValue = original

        End If
                RealX1 = found
        GoTo finish
            End If
        End If
    End If
    offset = rleftacc
    estX = estX + Len(ww(i))
Next i


If i = items& + 1 Then
    RealX1 = offset
    nValue = estX
End If
finish:
'SelectClipRgn mHdc, &H0
'DeleteObject hRgn
End Function

Private Sub RTL3(what As String, where As Long, mark10 As Long, mark20 As Long, numhandle As Boolean, M1 As Long, M2 As Long)
numhandle = False
M1 = 0
M2 = 0
mark10 = 0
mark20 = 0
Dim A2() As Integer, R As Long, ii As Long
Dim wr$
Dim mcharpos As Long
If Len(what) = 0 Then Exit Sub

If where = -1 Then mcharpos = Len(what) Else mcharpos = where - 1
If mcharpos > Len(what) Then Exit Sub
ReDim A2(Len(what) + 10)


If GetStringTypeExW(&HB, 2, StrPtr(what), Len(what), A2(0)) = 0 Then Exit Sub
For ii = mcharpos To 0 Step -1
    If ii < Len(what) Then
    If A2(ii) = 9 Then
   If mark10 > 0 Then Exit For
    Exit Sub
    End If
    If A2(ii) = 2 Then
    mark10 = ii + 1
    End If
    If mark10 > 0 And A2(ii) = 1 Then
    Exit For
    End If
    End If
Next ii

If mark10 = 0 Then Exit Sub
For ii = mark10 - 1 To Len(what) - 1

    If A2(ii) = 1 Then
    Exit For
    End If
    If ii >= 0 Then
        If A2(ii) = 7 Then
            For R = ii + 1 To Len(what) - 1
                If A2(R) = 1 Then Exit For
                If A2(R) = 9 Then
                Exit For
                End If
                If A2(R) <> 7 And A2(R) < 11 Then
                    If (A2(R) And 2) = 2 Then
                        ii = R
                        R = -1
                        Exit For
                    End If
                   If mark20 > mcharpos Then: R = 0: Exit For
                End If
            Next R
            If R = 0 Then
                Exit For
            ElseIf R <> -1 Then
                ii = R - 1
            End If
        End If
        If (A2(ii) And 2) = 2 And A2(ii) < 11 And A2(ii) <> 10 Then
            mark20 = ii + 1
            If mark20 > 0 And A2(ii) = 1 Then
            Exit For
            End If
        End If
    End If
    If A2(ii) = 9 Then
    If mark20 > 0 Then mark20 = ii
    Exit For
    End If
Next ii

'If mark10 > 0 And mark20 > 0 Then
there11:
If mcharpos > 0 Then
If A2(mcharpos) = 3 And A2(mcharpos - 1) <> 3 Then
M1 = mcharpos
M2 = mcharpos
Else
M1 = mcharpos - 1
M2 = mcharpos - 1
End If
Else
Exit Sub

End If
Dim findme As Long


If A2(M1) <> 3 Then Exit Sub
' find language
For findme = mcharpos To mark1 Step -1
If A2(findme) = 2 Then Exit For
Next findme
findme = AscW(Mid$(what, findme + 1, 1))
If findme >= &H600& And findme <= &H6FF& Then
' arabic language
                    Do While M1 > 0
                        If (A2(M1 - 1) > 3 And A2(M1) < 11) And (A2(M1) > 3 And A2(M1) < 10) Then

                        ElseIf A2(M1 - 1) = 4 Then
                        If M1 > 1 Then
                        If A2(M1 - 2) >= 10 Then
                        M1 = M1 - 1
                        Exit Do
                        End If
                        End If
                        Else
                        If Not (A2(M1) = 3 Or A2(M1) = 7) Then Exit Do
                       
                    End If
                        M1 = M1 - 1
                    Loop
                    Do While M2 < mark2
                    If (A2(M2 + 1) > 3 And A2(M2) < 11) And (A2(M2) > 3 And A2(M2) < 10) Then

                        Else
                    If Not (A2(M2) = 3 Or A2(M2) = 7) Then Exit Do
                    
                    End If
                        M2 = M2 + 1
                    Loop
Else

Do While M1 > 0
    If (A2(M1 - 1) > 3 And A2(M1) < 10) Or A2(M1) = 4 Then
 If A2(M1 - 1) = 4 Then
 If M1 > 1 Then
 If A2(M1 - 2) = 11 Then
 M1 = M1 - 1: Exit Do
 End If
 End If
 ElseIf A2(M1) = 1 Then
Exit Do
End If
    Else
    If Not (A2(M1) = 3 Or A2(M1) = 7) Then Exit Do
   
End If
    M1 = M1 - 1
Loop
Do While M2 < mark2

If A2(M2) > 3 And A2(M2) < 10 Then
If A2(M2) = 4 Then
If A2(M2 + 1) <> 3 Then
Exit Do
End If
End If
    Else
If Not (A2(M2) = 3 Or A2(M2) = 7) Then Exit Do

End If
    M2 = M2 + 1
Loop
End If
M1 = M1 + 2

numhandle = True
If mcharpos = M2 Then
where = M1 - (M1 = 1)

Else
where = M2 - (mcharpos - M1)

End If


End Sub



Private Sub glistN_Selected(item As Long)
If item >= 0 Then
   '' If (Not mCallback Is Nothing) Then mCallback.Press Index
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click()"
      End If
    End If
End Sub

Private Sub glistN_Selected2(item As Long)
If item >= 0 Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Enter(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Enter()"
      End If
End If
End Sub
Private Sub glistN_SpinnerValue(ThatValue As Long)
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".SpinnerValue(" + Str(mIndex) + "," + LTrim$(Str(ThatValue)) + ")"
   Else
      Callback.Callback mesCtrlName + ".SpinnerValue(" + LTrim$(Str(ThatValue)) + ")"
      End If
End Sub
Public Property Get locked() As Boolean
locked = Not glistN.Enabled
End Property

Public Property Let locked(ByVal RHS As Boolean)
glistN.Enabled = Not RHS
End Property
Public Sub Spinner(ShowBar As Boolean, MinValue As Long, maxvalue As Long, LargeValue As Long)
If Not ShowBar Then
glistN.Spinner = False
glistN.ShowBar = False
glistN.EnabledBar = False
Else
glistN.Enabled = False
glistN.restrictLines = 1
glistN.Spinner = True
glistN.SetSpin MinValue, maxvalue, LargeValue
glistN.Enabled = True
glistN.EnabledBar = True
glistN.ShowBar = True

End If
glistN.Vertical = True
End Sub


Public Property Get Value()
On Error Resume Next
Value = 0&
If glistN.Spinner Then
Value = CVar(glistN.ValueSilent)
ElseIf NumberOnly Then
    If NLtrim$(vartext) = vbNullString Then
        If Not NumberIntOnly Then
            Value = CDbl(0)
        Else
            Value = CDec(0)
        End If
    ElseIf Not NumberIntOnly Then
        If NowDec$ <> "." Then
            Value = CDbl(val(Replace$(vartext, NowDec$, ".")))
        Else
            Value = CDbl(val(vartext))
        End If
    Else
        Value = CDec(val(vartext))
    End If
Else
    Value = vartext
End If

End Property

Public Property Let Value(ByVal RHS)
If glistN.Spinner Then
On Error Resume Next
glistN.ValueSilent = CLng(RHS)
ElseIf NumberOnly Then
    If IsNumeric(RHS) Then
        vartext = Replace$(LTrim(Str$(RHS)), NowDec$, ".")
    Else
        vartext = Replace$(RHS, NowDec$, ".")
    End If
ElseIf IsNumeric(RHS) Then
    vartext = Replace$(LTrim(Str$(RHS)), NowDec$, ".")
Else
    vartext = RHS
End If

End Property

Public Property Get MaxCharLength() As Long
MaxCharLength = glistN.maxchar
End Property

Public Property Let MaxCharLength(ByVal RHS As Long)
glistN.maxchar = RHS
End Property

Private Function Myleft$(a As String, Pos As Long)
If Pos < 1 Then Myleft$ = vbNullString Else Myleft$ = Mid$(a, 1, Pos)
End Function
Private Function GetTextData(ByVal lFormatId As Long) As String
'' Addition to myTextBox for ctrl+V
Dim bData() As Byte, sr As String, sr1 As String
sr1 = Clipboard.GetText(1)
If (OpenClipboard(0) <> 0) Then

        
    If (GetBinaryData(lFormatId, bData())) Then
        sr = bData
        
        If IsWine Then
            sr1 = Myleft$(sr, Len(sr1))
            GetTextData = Myleft$(sr1, Len(sr1))
        Else
            GetTextData = Myleft$(sr, Len(sr1))
        End If

    End If

End If
CloseClipboard
End Function
Private Function GetClipboardMemoryHandle( _
        ByVal lFormatId As Long _
    ) As Long

    

    If (IsClipboardFormatAvailable(lFormatId) <> 0) Then
         
        GetClipboardMemoryHandle = GetClipboardData(lFormatId)
        
    End If
End Function
Private Function GetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean

Dim hMem As Long, lSize As Long, lPtr As Long
    

    Erase bData
    
    hMem = GetClipboardMemoryHandle(lFormatId)

    If (hMem <> 0) Then

        lSize = GlobalSize(hMem)

        lPtr = GlobalLock(hMem)
        If (lSize > 0) Then

            ReDim bData(0 To lSize - 2) As Byte

            CopyMemory bData(0), ByVal lPtr, lSize - 1
        End If

        GlobalUnlock hMem

        GetBinaryData = (lSize > 0)

    End If
End Function

Private Sub glistN_PreviewKeyboardUnicode(a As String)
Callback.SpreadKey a
End Sub

Private Sub glistN_UnregisterGlist()
If Not glistN.Spinner Then glistN.NoWheel = True
If glistN.TabStopSoft Then Callback.LastActive = ctrlName
On Error Resume Next
Set LastGlist = Nothing
If Err.Number > 0 Then glistN.NoWheel = True

End Sub

Public Sub move(x, y, Optional W, Optional H)
If IsMissing(W) Then
glistN.MoveTwips CLng(x), CLng(y), 0, 0
ElseIf IsMissing(H) Then
glistN.MoveTwips CLng(x), CLng(y), CLng(W), 0
Else
glistN.MoveTwips CLng(x), CLng(y), CLng(W), CLng(H)
End If
End Sub
Public Sub deconstruct()

If Not glistN Is Nothing Then glistN.Shutdown
Set Callback = Nothing
Set glistN = Nothing
End Sub
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.Name = ThisFontName

If ThisMode > 7 Then aa.size = ThisMode Else aa = 7
aa.bold = ThisBold
Set glistN.Font = aa
End If
End Sub
Public Property Get index() As Long
index = mIndex
End Property

Friend Property Let index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Property Let ShowAlways(ByVal vNewValue As Variant)
mStayOpen = CBool(vNewValue)
If mStayOpen Then
glistN.ForeColor = rgb(0, 0, 0)
glistN.BackColor = rgb(255, 255, 255)
End If
End Property
Sub GetFocus()
On Error Resume Next
If Callback.Controls(ctrlName).Visible Then
Callback.Controls(ctrlName).SetFocus
End If
End Sub
Public Property Let Default(RHS)
On Error Resume Next
If CBool(RHS) Then Callback.Default = ctrlName
End Property
Sub Transparent()
Dim x  As Long, y As Long
glistN.FadeLastLinePart = mPartColor
glistN.ForeColor = black
glistN.BackStyle = 1
glistN.GetLeftTop x, y
glistN.RepaintFromOut Callback.Image, x, y
Callback.RefreshList = Callback.RefreshList + 1
End Sub
Sub TransparentOff()
On Error Resume Next
If Not glistN Is Nothing Then
glistN.BackStyle = 0
glistN.BackColor = mycolor(dimm2)
glistN.PanPos = 0
glistN.ShowMe
Callback.RefreshList = Callback.RefreshList - 1
End If
End Sub
Sub ColorsReset()
white = rgb(255, 255, 255)
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
End Sub
Sub Colors(white1, Optional black1, Optional dimm11, Optional dimm21)

white = mycolor(white1)
If Not IsMissing(black1) Then
black = mycolor(black1)
If Not IsMissing(dimm11) Then
dimm1 = mycolor(dimm11)
If Not IsMissing(dimm21) Then dimm2 = mycolor(dimm21)
End If
End If
If Enabled Then
glistN.ForeColor = white
Else
glistN.ForeColor = black
End If

'Retired
End Sub

Public Property Get FadePartColor() As Variant
FadePartColor = mPartColor
End Property

Public Property Let FadePartColor(ByVal vNewValue As Variant)
On Error Resume Next
mPartColor = CLng(vNewValue)
End Property

Public Property Let icon(RHS)
glistN.icon = RHS
End Property
Public Property Let Drag(RHS)
glistN.DragEnabled = RHS
End Property
Public Property Get Drag()
Drag = glistN.DragEnabled
End Property
Public Property Let drop(RHS)
glistN.DropEnabled = RHS
End Property
Public Property Get drop()
drop = glistN.DropEnabled
End Property
Private Sub glistN_AddSelStart(val As Long, shift As Integer)
Dim W As Long, a$, l As Long, b$
If glistN.SelStart > 0 Then
a$ = Mid$(vartext, glistN.SelStart)

    If Len(a$) > 0 Then
        W = AscW(a$)
        If (shift And 2) = 2 Then
            l = Len(a$)
            val = 1
            If l < 2 Then Exit Sub
                b$ = Mid$(a$, val, 1)
                If InStr(sp1, b$) = 0 Then
                Do While val < l And InStr(sp, b$) > 0
                    val = val + 1
                    b$ = Mid$(a$, val, 1)
                Loop
                If InStr(sp1, b$) = 0 Then
                    If val = 1 Then
                        Do While val < l And InStr(sp2, b$) = 0
                            val = val + 1
                            b$ = Mid$(a$, val, 1)
                        Loop
                        If InStr(sp2, b$) > 0 Then val = val - 1
                    Else
                        If val < l Then val = val - 1
                    End If
                Else
                    val = val - 1
                End If
            End If
        ElseIf W > -10241 And W < -9216 Then
            val = 2
        End If
    End If
End If
End Sub

Private Sub glistN_SubSelStart(val As Long, shift As Integer)
Dim W As Long, a$, l As Long, b$
If glistN.SelStart < 2 Then Exit Sub
a$ = Mid$(vartext, glistN.SelStart - 1, 1)
If Len(a$) > 0 Then
W = AscW(a$)
        If (shift And 2) = 2 Then
            a$ = Mid$(vartext, 1, glistN.SelStart - 1)
            l = Len(a$)
            val = l
            If val = 1 Then Exit Sub
            b$ = Mid$(a$, l, 1)
            If InStr(sp1, b$) = 0 Then
                Do While l > 1 And InStr(sp, b$) > 0
                    l = l - 1
                    b$ = Mid$(a$, l, 1)
                Loop
                If InStr(sp1, b$) = 0 Then
                    If l = val Then
                        Do While l > 1 And InStr(sp2, b$) = 0
                            l = l - 1
                            b$ = Mid$(a$, l, 1)
                        Loop
                        If l = 1 And Not InStr(sp2, b$) > 0 Then l = l - 1
                        
                    Else
                     If InStr(sp2, b$) > 0 Then l = l - 1
                    End If
                End If
            Else
                l = l - 1
            End If
            val = val - l
        ElseIf W > -9217 And W < -8192 Then
            val = 2
        End If

End If
End Sub
Public Sub SetBarStyle(bColor As Long, Optional bHatch As Long, Optional BHatchColor As Long)
    glistN.BarColor = mycolor(bColor)
    glistN.BarHatch = bHatch
    glistN.BarHatchColor = BHatchColor
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
If Opcode < 0 Then Opcode = 0
If a = 0 Then Exit Sub
a = a - 1000 * shift - 10000 * ctrl - 100000 * alt
Callback.RegisterAcc a, ctrlName, Opcode
End Sub
Private Sub gListN_AccKey(m As Long)
Callback.AccProces m
End Sub

Private Sub glistN_WindowKey(VbKeyThis As Integer)
Callback.ProcessWinKey VbKeyThis
End Sub
