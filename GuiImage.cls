VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents pbox As PictureBox
Attribute pbox.VB_VarHelpID = -1
Private UKEY$, preserveINK$
Public Relax As Boolean
Private once As Boolean, mShowCaret As Boolean
Dim mIndex As Long, MoveOnEnter As Boolean, LastNumX As Boolean, DropKey As Boolean
Public ctrlName As String
Private Targets As Boolean, q() As target
Dim Callback As GuiM2000
Public MY_BACK As cDIBSection
Public NoEscapeKey As Boolean, Arrows2Tab As Boolean
Private mesCtrlName As String, mprive As Long, mtext As String
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
Private mynum$, lastshift As Integer
Public BypassKey As Boolean
Friend Property Get GetCallBack() As GuiM2000
Set GetCallBack = Callback
End Property
Friend Sub Construct(there As GuiM2000, Name$)
 mIndex = -1
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set pbox = there.Controls.Add("vb.picturebox", ctrlName)
 there.AddGuiControl Me
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
Friend Sub ConstructArray(there As GuiM2000, Name$, i As Long)
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set pbox = there.Controls.Add("vb.picturebox", ctrlName)
 mIndex = i
 there.AddGuiControl Me
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
Public Sub SetUp()
Dim ff As New StdFont
With Callback
ff.Name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With
MoveOnEnter = True
With pbox
.AutoRedraw = True
.BorderStyle = 0
.BackColor = rgb(200, 120, 60)
Set .Font = ff
.BackColor = Callback.Controls(1).BackColor
.forecolor = Callback.Controls(1).forecolor
End With
End Sub
Public Property Let Text(ByVal RHS As String)
mtext = RHS
pbox.Cls
pbox.currentX = 0
pbox.currentY = 0
pbox.Print mtext
pbox.Refresh
End Property
Public Property Get Text() As String
    Text = mtext
End Property
Public Sub deconstruct()
Set Callback = Nothing
Set pbox = Nothing
End Sub
Public Sub move(x, y, Optional w, Optional h)
If IsMissing(w) Then
pbox.move CLng(x), CLng(y), 0, 0
ElseIf IsMissing(h) Then
pbox.move CLng(x), CLng(y), CLng(w), 0
Else
pbox.move CLng(x), CLng(y), CLng(w), CLng(h)
End If
End Sub

Public Property Get locked() As Boolean
locked = Not pbox.Enabled
End Property

Public Property Let locked(ByVal RHS As Boolean)
pbox.Enabled = Not RHS
End Property
Public Property Get Enabled() As Boolean
Enabled = Not BypassKey
End Property

Public Property Let Enabled(ByVal RHS As Boolean)
BypassKey = Not RHS
If BypassKey Then DestroyCaret

End Property

Public Property Let prive(ByVal RHS As Long)
    If Not pbox Is Nothing Then
        pbox.Tag = RHS
        players(RHS).controlname = myUcase(Callback.modulename + "." + mesCtrlName, True)
       
    End If
        mprive = RHS
End Property
Public Property Get prive() As Long
    If Not pbox Is Nothing Then
        pbox.Tag = mprive
    End If
    prive = mprive
End Property

Private Sub Class_Terminate()
Dim clearbasket As basket
If prive <> 0 Then
    players(prive) = clearbasket
End If
End Sub
Property Let Visible(ByVal RHS As Boolean)
    If Not pbox Is Nothing Then
        pbox.Visible = RHS
    End If
End Property
Property Get Visible() As Boolean
    If Not pbox Is Nothing Then
        Visible = pbox.Visible
    End If
End Property

Private Sub pbox_Click()
    If Not Enabled Then Exit Sub
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + ")"
    Else
      Callback.Callback mesCtrlName + ".Click()"
    End If
End Sub

Private Sub pbox_GotFocus()
INK$ = preserveINK$
If BypassKey Then Exit Sub
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".GotFocus(" + Str(mIndex) + ")"
    Else
        Callback.Callback mesCtrlName + ".GotFocus()"
    End If
ShowCaret = ShowCaret
If ShowCaret Then Module2.ShowCaret pbox.hWnd
End Sub

Private Sub pbox_KeyDown(keycode As Integer, shift As Integer)
If BypassKey Then keycode = 0: shift = 0: Exit Sub
lastshift = shift
Dim VR(2)
VR(0) = keycode
VR(1) = shift
If mIndex >= 0 Then
Callback.CallbackNow mesCtrlName + ".KeyDown(" + Str(mIndex) + ")", VR()
Else
Callback.CallbackNow mesCtrlName + ".KeyDown()", VR()
End If
keycode = VR(0)
shift = VR(1)

If keycode = 27 And NoEscapeKey Then
keycode = 0
Exit Sub
End If
If Arrows2Tab Then
    If keycode = vbKeyLeft Or (keycode = vbKeyUp And Arrows2Tab) Then
        ChooseNextLeft pbox, pbox.Parent
        keycode = 0
        Exit Sub
    ElseIf keycode = vbKeyRight Or (keycode = vbKeyDown And Arrows2Tab) Then
        ChooseNextRight pbox, pbox.Parent
        keycode = 0
        Exit Sub
    End If
End If
If keycode = vbKeyTab Then  'And Not EditFlagSpecial
    If shift = 1 Then
        choosenext
        keycode = 0
        Exit Sub
    End If
ElseIf keycode = vbKeyF4 Then
If shift = 4 Then
On Error Resume Next
If pbox.Parent.Name = "GuiM2000" Or pbox.Parent.Name = "Form2" Or pbox.Parent.Name = "Form4" Then
With pbox.Parent
.ByeBye
End With
keycode = 0
Exit Sub
End If
End If
End If
If DropKey Then keycode = 0: Exit Sub
Dim i&
If shift = 4 Then
If keycode = 18 Then
If mynum$ = vbNullString Then mynum$ = "0"
keycode = 0
Exit Sub
Else
If keycode <> 0 Then GetKeY2 keycode, shift
End If
Select Case keycode
Case vbKeyAdd, vbKeyInsert
mynum$ = "&h"
Case vbKey0 To vbKey9
mynum$ = mynum$ + Chr$(keycode - vbKey0 + 48)
LastNumX = True
Case vbKeyNumpad0 To vbKeyNumpad9
LastNumX = False
mynum$ = mynum$ + Chr$(keycode - vbKeyNumpad0 + 48)
Case vbKeyA To vbKeyF
If Left$(mynum$, 1) = "&" Then
mynum$ = mynum$ + Chr$(keycode - vbKeyNumpad0 + 65)
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
If shift <> 0 And keycode = 0 Then Exit Sub
End Sub
Private Sub EnterOnly()
If mIndex >= 0 Then
    Callback.Callback mesCtrlName + ".Enter(" + Str(mIndex) + ")"
Else
    Callback.Callback mesCtrlName + ".Enter()"
End If

End Sub
Private Property Get ParentPreview() As Boolean
On Error Resume Next
If pbox.Parent.previewKey Then ParentPreview = True
End Property
Private Sub pbox_KeyPress(KeyAscii As Integer)
On Error GoTo fin
If BypassKey Then KeyAscii = 0: Exit Sub
Dim bb As Boolean, kk$, pair$, b1 As Boolean
If KeyAscii = 13 Then EnterOnly
If mynum$ <> "" Then
    Exit Sub
End If
If UKEY$ <> "" Then
    kk$ = UKEY$
    UKEY$ = vbNullString
Else
    kk$ = GetKeY(KeyAscii)
End If
    If ParentPreview Then Callback.SpreadKey kk$
   
    If DropKey Then KeyAscii = 0: UKEY$ = vbNullString: kk$ = vbNullString: Exit Sub
    If Enabled Then
Dim last$
last$ = INK$ + kk$
        INK$ = last$
again:
        If mIndex > -1 Then
            Callback.Callback mesCtrlName + ".KeyPress(" + CStr(index) + ")"
        Else
            Callback.Callback mesCtrlName + ".KeyPress()"
        End If
        If Not INK$ = last$ Then If Len(INK$) > 0 Then last$ = INK$: Debug.Print "again", Timer:   GoTo again
    End If
fin:
    KeyAscii = 0
End Sub

Private Sub pbox_KeyUp(keycode As Integer, shift As Integer)
On Error GoTo fin
'If PrevLocale <> GetLocale Then RaiseEvent Maybelanguage
If BypassKey Then keycode = 0: shift = 0: Exit Sub
Dim i As Long, k As Integer

lastshift = shift
If keycode = 18 Then

ElseIf keycode = 112 And (shift And 2) = 2 Then
keycode = 0
shift = 0
CtrlPlusF1
Exit Sub
ElseIf keycode >= vbKeyF1 And keycode <= vbKeyF12 Then
    k = ((keycode - vbKeyF1 + 1) + 12 * (shift And 1)) + 24 * (1 + ((shift And 2) = 0)) - 1000 * ((shift And 4) = 4)
    Fkey k
    If k = 0 Then keycode = 0: shift = 0
ElseIf keycode = 16 And shift <> 0 Then
  '  RaiseEvent Maybelanguage
ElseIf keycode = vbKeyV Then
Exit Sub
Else
If keycode = 27 And NoEscapeKey Then
keycode = 0
Else
'RaiseEvent RefreshOnly
End If
End If
If DropKey Then keycode = 0: UKEY$ = vbNullString: Exit Sub
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
'If LastNumX Then pbox_KeyPress 44
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
    Dim VR(2)
    VR(0) = keycode
    VR(1) = shift
    If mIndex >= 0 Then
    Callback.CallbackNow mesCtrlName + ".KeyUp(" + Str(mIndex) + ")", VR()
    Else
    Callback.CallbackNow mesCtrlName + ".KeyUp()", VR()
    End If
    keycode = VR(0)
    shift = VR(1)

fin:
End Sub


Private Sub pbox_LostFocus()
preserveINK$ = INK$
INK$ = ""
If ShowCaret Then DestroyCaret
If pbox.TabStop And Enabled And Visible Then
    Callback.LastActive = ctrlName
End If
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".LostFocus(" + Str(mIndex) + ")"
    Else
        Callback.Callback mesCtrlName + ".LostFocus()"
    End If
End Sub

Private Sub pbox_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
Dim bstack As basetask, oldhere$
If Not Enabled Then Exit Sub
If Not Relax Then
Relax = True
Dim sel&

If Button > 0 And Targets Then
    sel& = ScanTarget(q(), CLng(x), CLng(y), prive)
    If sel& >= 0 Then
        If Button = 1 Then
            Select Case q(sel&).id Mod 100
            Case Is < 10
                SwapStrings here$, oldhere$
                here$ = Callback.modulename()
                Set bstack = New basetask
                Set bstack.Owner = pbox
                Set bstack.Sorosref = New mStiva
                If Execute(bstack, (q(sel&).Comm), False) = 0 Then Beep
                SwapStrings here$, oldhere$
                
            Case Else
            
            If mIndex > -1 Then
                Callback.Callback mesCtrlName + ".Target" + "(" + CStr(index) + "," + Str(sel& + prive * 10000) + ")"
            Else
                Callback.Callback mesCtrlName + ".Target" + "(" + Str(sel& + prive * 10000) + ")"
            End If

            End Select
            
        End If
        
        Button = 0
        Relax = False
        Exit Sub
    End If
End If



If mIndex > -1 Then
    Callback.Callback mesCtrlName + ".MouseDown(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
    Callback.Callback mesCtrlName + ".MouseDown(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If
End Sub
Private Sub pbox_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Enabled Then Exit Sub
If Not Relax Then
Relax = True
If mIndex > -1 Then
Callback.Callback mesCtrlName + ".MouseUp(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback.Callback mesCtrlName + ".MouseUp(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If
End Sub
Private Sub pbox_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Enabled Then Exit Sub
If Not Relax Then
Relax = True
If mIndex > -1 Then
Callback.Callback mesCtrlName + ".MouseMove(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback.Callback mesCtrlName + ".MouseMove(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If

End Sub
Private Sub pbox_DblClick()
    If Not Enabled Then Exit Sub
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".DblClick(" + Str(mIndex) + ")"
    Else
        Callback.Callback mesCtrlName + ".DblClick()"
    End If
End Sub

Private Sub pbox_Resize()
If mprive = 0 Then Exit Sub
Set MY_BACK = Nothing
If Not once Then
once = True
Dim VR(2)
VR(0) = Width
VR(1) = Height
    If mIndex >= 0 Then
        Callback.CallbackNow mesCtrlName + ".Resize(" + Str(mIndex) + "," + Str(mIndex) + ")", VR()
    Else
        Callback.CallbackNow mesCtrlName + ".Resize()", VR()
    End If
    If VR(0) <> Width Or VR(1) <> Height Then
        pbox.move pbox.Left, pbox.Top, VR(0), VR(1)
    End If
once = False
End If
End Sub
Public Sub CopyBack()
If pbox Is Nothing Then Exit Sub
pbox.PaintPicture pbox.Parent.Image, 0, 0, , , pbox.Left, pbox.Top
End Sub

Public Sub DisAllTargets()
DisableTargets q(), prive
End Sub
Friend Sub RenderTarget(bstack As basetask, rest$, Lang As Long, tHandle As Variant)
Dim p, w$, x
If tHandle \ 10000 <> prive Then
MyEr "target not for this image", "Ô ÛÙ¸˜ÔÚ ‰ÂÌ ÂﬂÌ·È „È· ·ıÙﬁ ÙÁÌ ÂÈÍ¸Ì·"
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

Public Property Get index() As Long
index = mIndex
End Property

Friend Property Let index(ByVal RHS As Long)
mIndex = RHS
End Property

Public Property Get Height() As Single
If Not pbox Is Nothing Then
    Height = pbox.Height
End If
End Property
Public Property Get Width() As Single
If Not pbox Is Nothing Then
    Width = pbox.Width
End If
End Property

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

Callback.AccProces acc
If acc = 0 Then ascii = 0: shift = 0
End Sub
Property Let KeyEvent(ByVal RHS As Boolean)
    If Not pbox Is Nothing Then
        DropKey = Not RHS
    End If
End Property
Property Get KeyEvent() As Boolean
    If Not pbox Is Nothing Then
        KeyEvent = Not DropKey
    End If
End Property
Property Let TabStop(ByVal RHS As Boolean)
    If Not pbox Is Nothing Then
        pbox.TabStop = RHS
        
    End If
End Property
Property Get TabStop() As Boolean
    If Not pbox Is Nothing Then
        TabStop = pbox.TabStop
    End If
End Property
Property Let ShowCaret(ByVal RHS As Boolean)
    If Not pbox Is Nothing Then
        mShowCaret = RHS
        If prive > 0 Then
        With players(prive)
        .ShowCaret = RHS
        DestroyCaret
        LCTCnew pbox, players(prive), .currow, .curpos
        If RHS Then Module2.ShowCaret pbox.hWnd
        End With
        End If
    End If
End Property
Property Get ShowCaret() As Boolean
    If Not pbox Is Nothing Then
        ShowCaret = mShowCaret
    End If
End Property

Private Sub Fkey(a As Integer)
If a > 1000 Then
Callback.SendFKEY a - 1000
Else
Callback.SendFKEY a
End If
End Sub
Private Sub CtrlPlusF1()
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
End Sub

Public Property Let Default(RHS)
On Error Resume Next
If CBool(RHS) Then Callback.Default = ctrlName
End Property
Sub GetFocus()
On Error Resume Next
DestroyCaret
Callback.Controls(ctrlName).SetFocus

End Sub
