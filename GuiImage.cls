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
Implements IControlIndex
Public WithEvents pbox As PictureBox
Attribute pbox.VB_VarHelpID = -1
Private UKEY$, preserveINK$
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
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
    X As Long
    Y As Long
End Type
Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
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
 there.AddGuiControl Me, CVar(ctrlName)
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
Friend Sub ConstructArray(there As GuiM2000, Name$, i As Long)
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set pbox = there.Controls.Add("vb.picturebox", ctrlName)
 mIndex = i
 there.AddGuiControl Me, CVar(ctrlName)
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
.ClipControls = True
.BackColor = rgb(200, 120, 60)
Set .Font = ff
.BackColor = Callback.Controls(1).BackColor
.ForeColor = Callback.Controls(1).ForeColor
End With
End Sub
Public Property Let Text(ByVal RHS As String)
mtext = RHS
pbox.Cls
pbox.currentX = 0
pbox.currentY = 0
MyPrint pbox, mtext
pbox.Refresh
End Property
Public Property Get Text() As String
    Text = mtext
End Property
Public Sub deconstruct()
Set Callback = Nothing
Set pbox = Nothing
End Sub
Public Sub move(X, Y, Optional W, Optional H)
If IsMissing(W) Then
pbox.move CLng(X), CLng(Y), 0, 0
ElseIf IsMissing(H) Then
pbox.move CLng(X), CLng(Y), CLng(W), 0
Else
pbox.move CLng(X), CLng(Y), CLng(W), CLng(H)
End If
End Sub

Public Property Get locked() As Boolean
locked = Not pbox.enabled
End Property

Public Property Let locked(ByVal RHS As Boolean)
pbox.enabled = Not RHS
End Property
Public Property Get enabled() As Boolean
enabled = Not BypassKey
End Property

Public Property Let enabled(ByVal RHS As Boolean)
BypassKey = Not RHS
If BypassKey Then DestroyCaret

End Property

Public Property Let prive(ByVal RHS As Long)
    If Not pbox Is Nothing Then
        pbox.Tag = RHS
        players(RHS).ControlName = myUcase(Callback.modulename + "." + mesCtrlName, True)
       
    End If
        mprive = RHS
End Property
Public Property Get prive() As Long
On Error Resume Next
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
    If Not enabled Then Exit Sub
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".Click(" + str(mIndex) + ")"
    Else
      Callback.Callback mesCtrlName + ".Click()"
    End If
End Sub

Private Sub pbox_GotFocus()
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".GotFocus(" + str(mIndex) + ")"
    Else
        Callback.Callback mesCtrlName + ".GotFocus()"
    End If
If BypassKey Then Exit Sub
INK$ = preserveINK$
ShowCaret = ShowCaret
If ShowCaret Then Module2.ShowCaret pbox.hWnd
End Sub

Private Sub pbox_KeyDown(KeyCode As Integer, shift As Integer)

If BypassKey Then
GetKeY2 KeyCode, shift
KeyCode = 0: shift = 0: Exit Sub
End If
lastshift = shift
Dim VR(2)
VR(0) = KeyCode
VR(1) = shift
If mIndex >= 0 Then
Callback.CallbackNow mesCtrlName + ".KeyDown(" + str(mIndex) + ")", VR()
Else
Callback.CallbackNow mesCtrlName + ".KeyDown()", VR()
End If
KeyCode = VR(0)
shift = VR(1)

If KeyCode = 27 And NoEscapeKey Then
KeyCode = 0
Exit Sub
End If
If Arrows2Tab Then
    If KeyCode = vbKeyLeft Or (KeyCode = vbKeyUp And Arrows2Tab) Then
        If Typename(pbox.Parent) = "GuiM2000" Then
            ChooseNextLeft pbox, pbox.Parent
        End If
        KeyCode = 0
        Exit Sub
    ElseIf KeyCode = vbKeyRight Or (KeyCode = vbKeyDown And Arrows2Tab) Then
        If Typename(pbox.Parent) = "GuiM2000" Then
            ChooseNextRight pbox, pbox.Parent
        End If
        KeyCode = 0
        Exit Sub
    End If
End If
If KeyCode = vbKeyTab Then  'And Not EditFlagSpecial
    If shift = 1 Then
        choosenext
        KeyCode = 0
        Exit Sub
    End If
ElseIf KeyCode = vbKeyF4 Then
If shift = 4 Then
On Error Resume Next
If pbox.Parent.Name = "GuiM2000" Or pbox.Parent.Name = "Form2" Or pbox.Parent.Name = "Form4" Then
With pbox.Parent
.ByeBye
End With
KeyCode = 0
Exit Sub
End If
End If
End If

If DropKey Then
KeyCode = 0: Exit Sub
Else
If KeyCode <> 0 Then GetKeY2 KeyCode, shift
If KeyCode = 0 Then Exit Sub
End If
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
End Sub
Private Sub EnterOnly()
If mIndex >= 0 Then
    Callback.Callback mesCtrlName + ".Enter(" + str(mIndex) + ")"
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
    If enabled Then
Dim last$
last$ = INK$ + kk$
        INK$ = last$
again:
        If mIndex > -1 Then
            Callback.Callback mesCtrlName + ".KeyPress(" + CStr(Index) + ")"
        Else
            Callback.Callback mesCtrlName + ".KeyPress()"
        End If
        If Not INK$ = last$ Then If Len(INK$) > 0 Then last$ = INK$: GoTo again
    End If
fin:
    KeyAscii = 0
End Sub

Private Sub pbox_KeyUp(KeyCode As Integer, shift As Integer)
On Error GoTo fin
'If PrevLocale <> GetLocale Then RaiseEvent Maybelanguage
If BypassKey Then KeyCode = 0: shift = 0: Exit Sub
Dim i As Long, K As Integer

lastshift = shift
If KeyCode = 18 Then

ElseIf KeyCode = 112 And (shift And 2) = 2 Then
KeyCode = 0
shift = 0
CtrlPlusF1
Exit Sub
ElseIf KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
    K = ((KeyCode - vbKeyF1 + 1) + 12 * (shift And 1)) + 24 * (1 + ((shift And 2) = 0)) - 1000 * ((shift And 4) = 4)
    Fkey K
    If K = 0 Then KeyCode = 0: shift = 0
ElseIf KeyCode = 16 And shift <> 0 Then
  '  RaiseEvent Maybelanguage
ElseIf KeyCode = vbKeyV Then
Exit Sub
Else
If KeyCode = 27 And NoEscapeKey Then
KeyCode = 0
Else
'RaiseEvent RefreshOnly
End If
End If
If DropKey Then KeyCode = 0: UKEY$ = vbNullString: Exit Sub
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
    VR(0) = KeyCode
    VR(1) = shift
    If mIndex >= 0 Then
    Callback.CallbackNow mesCtrlName + ".KeyUp(" + str(mIndex) + ")", VR()
    Else
    Callback.CallbackNow mesCtrlName + ".KeyUp()", VR()
    End If
    KeyCode = VR(0)
    shift = VR(1)

fin:
End Sub


Private Sub pbox_LostFocus()

If ShowCaret Then DestroyCaret
If pbox.TabStop And enabled And Visible Then
    Callback.LastActive = ctrlName
End If
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".LostFocus(" + str(mIndex) + ")"
    Else
        Callback.Callback mesCtrlName + ".LostFocus()"
    End If
If BypassKey Then INK$ = "": Exit Sub
preserveINK$ = INK$
INK$ = ""
End Sub

Private Sub pbox_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
Dim bstack As basetask, oldHere$
If Not enabled Then Exit Sub
If Not Relax Then
Relax = True
Dim sel&

If Button > 0 And Targets Then
    sel& = ScanTarget(q(), CLng(X), CLng(Y), prive)
    If sel& >= 0 Then
        If Button = 1 Then
            Select Case q(sel&).id Mod 100
            Case Is < 10
                SwapStrings here$, oldHere$
                here$ = Callback.modulename()
                Set bstack = New basetask
                Set bstack.Owner = pbox
                Set bstack.Sorosref = New mStiva
                If Execute(bstack, (q(sel&).Comm), False) = 0 Then Beep
                SwapStrings here$, oldHere$
                
            Case Else
            
            If mIndex > -1 Then
                Callback.Callback mesCtrlName + ".Target" + "(" + CStr(Index) + "," + str(sel& + prive * 10000) + ")"
            Else
                Callback.Callback mesCtrlName + ".Target" + "(" + str(sel& + prive * 10000) + ")"
            End If

            End Select
            
        End If
        
        Button = 0
        Relax = False
        Exit Sub
    End If
End If



If mIndex > -1 Then
    Callback.Callback mesCtrlName + ".MouseDown(" + CStr(Index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(X) + "," + CStr(Y) + ")"
Else
    Callback.Callback mesCtrlName + ".MouseDown(" + CStr(Button) + "," + CStr(shift) + "," + CStr(X) + "," + CStr(Y) + ")"
End If
Relax = False
End If
End Sub
Private Sub pbox_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
If Not enabled Then Exit Sub
If Not Relax Then
Relax = True
If mIndex > -1 Then
Callback.Callback mesCtrlName + ".MouseUp(" + CStr(Index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(X) + "," + CStr(Y) + ")"
Else
Callback.Callback mesCtrlName + ".MouseUp(" + CStr(Button) + "," + CStr(shift) + "," + CStr(X) + "," + CStr(Y) + ")"
End If
Relax = False
End If
End Sub
Private Sub pbox_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Not enabled Then Exit Sub
If Not Relax Then
Relax = True
If mIndex > -1 Then
Callback.Callback mesCtrlName + ".MouseMove(" + CStr(Index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(X) + "," + CStr(Y) + ")"
Else
Callback.Callback mesCtrlName + ".MouseMove(" + CStr(Button) + "," + CStr(shift) + "," + CStr(X) + "," + CStr(Y) + ")"
End If
Relax = False
End If

End Sub
Private Sub pbox_DblClick()
    If Not enabled Then Exit Sub
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".DblClick(" + str(mIndex) + ")"
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
        Callback.CallbackNow mesCtrlName + ".Resize(" + str(mIndex) + "," + str(mIndex) + ")", VR()
    Else
        Callback.CallbackNow mesCtrlName + ".Resize()", VR()
    End If
    If VR(0) <> Width Or VR(1) <> Height Then
        pbox.move pbox.Left, pbox.top, VR(0), VR(1)
    End If
once = False
End If
End Sub
Public Sub CopyBack()
If pbox Is Nothing Then Exit Sub
pbox.PaintPicture pbox.Parent.Image, 0, 0, , , pbox.Left, pbox.top
End Sub

Public Sub DisAllTargets()
DisableTargets q(), prive
End Sub
Friend Sub RenderTarget(bstack As basetask, rest$, Lang As Long, tHandle As Variant)
Dim p, W$, X, Y As Long
If tHandle \ 10000 <> prive Then
    MyEr "target not for this image", "ï óôü÷ïò äåí åßíáé ãéá áõôÞ ôçí åéêüíá"
    Exit Sub
End If
p = tHandle Mod 10000
If p >= 0 And p < UBound(q()) Then
    While FastSymbol(rest$, ",")
        X = Empty
        Y = 0
        If IsLabelSymbolNew(rest$, "ÖÑÁÓÇ", "TEXT", Lang, Y) Then
            If IsStrExp(bstack, rest$, W$) Then
                q(p).Tag = W$
            Else
                GoTo skipcommand
            End If
        ElseIf IsLabelSymbolNew(rest$, "ÐÅÍÁ", "PEN", Lang, Y) Then
            If IsExp(bstack, rest$, X, , True) Then
                q(p).pen = X
            Else
                GoTo skipcommand
            End If
        ElseIf IsLabelSymbolNew(rest$, "ÖÏÍÔÏ", "BACK", Lang, Y) Then
            If IsExp(bstack, rest$, X, , True) Then
                q(p).back = X
            Else
                GoTo skipcommand
            End If
        ElseIf IsLabelSymbolNew(rest$, "ÐËÁÉÓÉÏ", "BORDER", Lang, Y) Then
            If IsExp(bstack, rest$, X, , True) Then
                q(p).fore = X
            Else
                GoTo skipcommand
            End If
        ElseIf IsLabelSymbolNew(rest$, "ÏÄÇÃÉÁ", "COMMAND", Lang, Y) Then
            If IsStrExp(bstack, rest$, W$) Then
                q(p).Comm = W$
            Else
                GoTo skipcommand
            End If
        ElseIf IsLabelSymbolNew(rest$, "ÔÉÌÇ", "VALUE", Lang, Y) Then
            If IsExp(bstack, rest$, X, , True) Then
                q(p).topval = Int(X * 100)
            Else
                GoTo skipcommand
            End If
        ElseIf IsLabelSymbolNew(rest$, "ÂÁÓÇ", "BASE", Lang, Y) Then
            If IsExp(bstack, rest$, X, , True) Then
                q(p).botval = Int(X * 100)
            Else
                GoTo skipcommand
            End If
        ElseIf IsLabelSymbolNew(rest$, "×ÑÙÌÁ", "COLOR", Lang, Y) Then
            If IsExp(bstack, rest$, X, , True) Then
                q(p).barC = X
            Else
                GoTo skipcommand
            End If
        ElseIf IsLabelSymbolNew(rest$, "ÌÅÃÅÈÏÓ", "SIZE", Lang, Y) Then
            If IsExp(bstack, rest$, X, , True) Then
                If X > 100 Then X = 100
                If X < -100 Then X = -100
                q(p).imagesize = Int(X)
            Else
                GoTo skipcommand
            End If
' "ÊÁÈÅÔÇ", "PORTRAIT"
        ElseIf IsLabelSymbolNew(rest$, "ÊÁÈÅÔÇ", "PORTRAIT", Lang, Y) Then
            If IsExp(bstack, rest$, X, , True) Then
                q(p).Vertical = Int(X) <> 0
            Else
                GoTo skipcommand
            End If
        ElseIf IsLabelSymbolNew(rest$, "ÅÉÊÏÍÁ", "IMAGE", Lang, Y) Then
            If IsExp(bstack, rest$, X) Then
                If bstack.lastobj Is Nothing Then
                    Set q(p).drawimage = Nothing
                ElseIf TypeOf bstack.lastobj Is mHandler Then
                    Dim usehandler As mHandler
                    Set usehandler = bstack.lastobj
                    Set bstack.lastobj = Nothing
                    If usehandler.T1 = 2 Then
                        Set q(p).drawimage = usehandler.objref
                        Else
                        GoTo err123
                        End If
                     
                    Else
err123:
                        WrongObject
                        Exit Sub
                    End If
                Else
                    GoTo skipcommand
                End If
        Else
            GoTo skipcommand
        End If
    Wend
    RTarget bstack, q(p)
End If
Exit Sub
skipcommand:
Dim V1 As Long
V1 = 1
FastSymbol rest$, ","
Do
aheadstatusSkipParam rest$, V1
If V1 > 1 And V1 <= Len(rest$) Then
    rest$ = Mid$(rest$, V1)
    FastSymbol rest$, ","
    V1 = 1
Else
    Exit Do
End If
Loop Until V1 = 1
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

Public Property Get Index() As Long
Index = mIndex
End Property

Friend Property Let Index(ByVal RHS As Long)
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
    Dim r&
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF&
      GetLocale = val("&H" + Right(Hex(r), 4))
End Function
Function GetKeY(ascii As Integer) As String
    Dim Buffer As String, ret As Long

    Buffer = String$(514, 0)
    Dim r&
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF&
      r = val("&H" + Right(Hex(r), 4))
    ret = GetLocaleInfo(r, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
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
   Callback.Callback mesCtrlName + ".About(" + str(mIndex) + ")"
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
Friend Sub TakeKey(KeyCode As Integer, shift As Integer)
    pbox_KeyDown KeyCode, shift
If KeyCode <> 0 Then
    pbox_KeyUp KeyCode, shift
End If
End Sub
Public Sub MaskColor(ByVal trV, Optional ByVal rr As Long, Optional ref As Boolean = True)
Dim photo As New cDIBSection, myRgn As Long, tr As Long
If pbox Is Nothing Then Exit Sub
rr = Abs(rr)
tr = mycolor(trV)
photo.ClearUp
If photo.create(pbox.Width \ DXP, pbox.Height \ DYP) Then
    photo.LoadPictureBlt pbox.hDC
    If photo.bitsPerPixel <> 24 Then Conv24 photo
    myRgn = fRegionFromBitmap2(photo, tr, CInt(rr))
    If myRgn = 0 Then
        myRgn = CreateRectRgn(0, 0, photo.Width, photo.Height)
    End If
    Call SetWindowRgn(pbox.hWnd, myRgn, ref)
End If
End Sub
Public Sub DeleteMask(Optional ref As Boolean = True)
Call SetWindowRgn(pbox.hWnd, (0), ref)
End Sub
Public Sub CurveMask(Optional t As Boolean = False, Optional factor As Single = 25)
Dim hRgn As Long
If pbox Is Nothing Then Exit Sub
If Int(factor) > 2 Then
hRgn = CreateRoundRectRgn(0, 0, pbox.Width \ DXP, pbox.Height \ DYP, factor, factor)
SetWindowRgn pbox.hWnd, hRgn, t
End If
End Sub
Public Sub ZOrder(Optional ByVal t)
If pbox Is Nothing Then Exit Sub
If IsMissing(t) Then
    pbox.ZOrder
Else
    pbox.ZOrder t
End If
End Sub
Public Property Let IControlIndex_Index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Property Get IControlIndex_Index() As Long
IControlIndex_Index = mIndex
End Property

Public Sub IControlIndex_UnloadMe()
On Error Resume Next
If pbox Is Nothing Then Exit Sub
Callback.RemoveGuiControl ctrlName
mIndex = -1
Set Callback = Nothing
Set pbox = Nothing
End Sub

