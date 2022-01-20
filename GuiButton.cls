VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Dim mCaptext As String
Public ctrlName As String
Private mesCtrlName As String, mlocked As Boolean, mTimes
Dim mIndex As Long
Dim Callback As GuiM2000
Dim white As Long, black As Long, dimm1 As Long, dimm2 As Long, shadowcol As Long
Public SkipFirstClick As Boolean

Friend Property Get GetCallBack() As GuiM2000
Set GetCallBack = Callback
End Property
Private Sub OneClick()
On Error Resume Next
If SkipFirstClick Then glistN.PrepareClick
End Sub
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
Public Sub move(x, y, Optional w, Optional h)
If IsMissing(w) Then
glistN.MoveTwips CLng(x), CLng(y), 0, 0
ElseIf IsMissing(h) Then
glistN.MoveTwips CLng(x), CLng(y), CLng(w), 0
Else
glistN.MoveTwips CLng(x), CLng(y), CLng(w), CLng(h)
End If
End Sub
Public Sub deconstruct()
Shutdown
Set Callback = Nothing
  Set glistN = Nothing
End Sub
Friend Sub SetUp()
white = Callback.Controls(1).forecolor
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
Dim ff As New StdFont
With Callback
ff.Name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With
SkipFirstClick = True
With glistN
.Arrows2Tab = True
.NoArrowDown = True
.NoArrowUp = True
.NoEscapeKey = True
.bypassfirstClick = True
.NoCaretShow = True
.NoPanRight = False
.restrictLines = 1
.CenterText = True
.VerticalCenterText = True
.Text = mCaptext
.backcolor = rgb(200, 120, 60)
Set .Font = ff
.forecolor = Callback.Controls(1).forecolor
shadowcol = -.forecolor
.dcolor = Callback.Controls(1).dcolor

.enabled = True
.ListindexPrivateUse = 0
.ShowMe
End With
End Sub

Private Sub Class_Initialize()
mCaptext = "Button"
mIndex = -1
End Sub

Private Sub Class_Terminate()
Set glistN = Nothing
End Sub



Private Sub glistN_BlinkNow(Face As Boolean)
On Error GoTo ex1
    If mTimes > 0 Then mTimes = mTimes - 1: If mTimes = 0 Then glistN.BlinkON = False
    If Not glistN.BlinkON Then Face = True
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".Blink(" + Str(mIndex) + "," + Str(Face) + ")"
    Else
        Callback.Callback mesCtrlName + ".Blink(" + Str(Face) + ")"
    End If
    
    If Not glistN.BlinkON Then
      If Not glistN.backcolor = mycolor(dimm1) Then
        If Not glistN.BackStyle = 0 Then
            glistN.backcolor = mycolor(dimm1)
            glistN.ShowMe
        End If
        End If
    End If
ex1:
End Sub

Private Sub glistN_CheckGotFocus()
Callback.RealHover = ctrlName
If mlocked Then
On Error Resume Next
'Callback.Controls(ctrlName).TabStop = False
If Callback.LastActive = vbNullString Then
Callback.LastActive = "gList2"
End If
Callback.Controls(Callback.LastActive).SetFocus
Exit Sub
End If
If Me.enabled Then
If glistN.BackStyle = 1 Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".GotFocus(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".GotFocus()"
      End If
End If
glistN.backcolor = mycolor(dimm1)
glistN.ShowMe
'OneClick
End If
End Sub

Private Sub glistN_CheckLostFocus()

If Me.enabled Then
Callback.LastActive = ctrlName
If glistN.BackStyle = 1 Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".LostFocus(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".LostFocus()"
      End If
End If
If Not mlocked Then glistN.backcolor = mycolor(dimm2)
glistN.PanPos = 0
glistN.ShowMe
End If
glistN.DblClick
End Sub

Private Sub glistN_CtrlPlusF1()
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
End Sub

Private Sub glistN_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
glistN.ShowMe
End Sub


Private Sub glistN_PanLeftRight(direction As Boolean)

 
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click()"
      End If
    
End Sub

Private Sub glistN_RegisterGlist(this As gList)
this.NoWheel = True
End Sub

Private Sub glistN_selected(item As Long)
If Not SkipFirstClick Then
 If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".ClickOnce(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".ClickOnce()"
      End If
Else
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
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click()"
      End If
      OneClick
    End If
End Sub
'
'
Public Property Let Caption(myCap As String)
'
mCaptext = myCap
If Not glistN Is Nothing Then
If glistN.CenterText Then
glistN.list(0) = mCaptext
Else
glistN.list(0) = "  " + mCaptext
End If
glistN.ShowMe
End If
End Property
Public Property Get Caption() As String
Caption = mCaptext
End Property

Public Property Get index() As Long
index = mIndex
End Property

Friend Property Let index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Property Get enabled() As Long
enabled = glistN.enabled
End Property

Public Property Let enabled(ByVal RHS As Long)
If glistN Is Nothing Then Exit Property
glistN.enabled = RHS
End Property
Public Property Get WrapText() As Long
WrapText = glistN.WrapText
End Property

Public Property Let WrapText(ByVal RHS As Long)
glistN.WrapText = RHS <> 0
End Property
Sub ColorsReset()
white = rgb(255, 255, 255)
glistN.forecolor = rgb(255, 255, 255)
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
End Sub
Sub Colors(white1, Optional dimm21, Optional dimm11, Optional dimm33)
On Error Resume Next
If Not glistN Is Nothing Then
glistN.forecolor = mycolor(white1)
If Not IsMissing(dimm21) Then dimm2 = mycolor(dimm21)

If Not IsMissing(dimm11) Then dimm1 = mycolor(dimm11)

If Not IsMissing(dimm33) Then glistN.CapColor = CLng(mycolor(dimm33))
End If


'Retired
End Sub
Public Sub Shutdown()
On Error Resume Next
If Not glistN Is Nothing Then
glistN.Shutdown
glistN.enabled = False
End If
End Sub
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.Name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set glistN.Font = aa
End If
End Sub
Public Sub Curve(RHS As Single)
glistN.Curve False, RHS
End Sub
Public Property Get top() As Variant
'top = Callback.Controls(ctrlName).top
top = glistN.TopTwips
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
Sub GetFocus()
On Error Resume Next
Callback.Controls(ctrlName).SetFocus
End Sub

Public Property Get locked() As Variant
locked = mlocked
End Property

Public Property Let locked(ByVal vNewValue As Variant)
mlocked = vNewValue
glistN.dropkey = mlocked
glistN.TabStop = Not mlocked
glistN.backcolor = mycolor(dimm2)
glistN.PanPos = 0
glistN.ShowMe
End Property
Sub Show()
glistN.ShowPan
End Sub
Public Property Let BlinkTimes(ByVal vNewValue As Variant)
mTimes = vNewValue
End Property
Public Property Get BlinkTimes()
BlinkTimes = mTimes
End Property
Public Property Let Blink(ByVal vNewValue As Variant)
    If vNewValue = 0 Then
        mTimes = 1
    Else
    glistN.BlinkTime = vNewValue
    End If
End Property
Public Property Get Blink()
Blink = glistN.BlinkTime
End Property
Sub Transparent()
Dim x  As Long, y As Long
glistN.BackStyle = 1
glistN.GetLeftTop x, y
glistN.RepaintFromOut Callback.Image, x, y
glistN.ShowMe
Callback.RefreshList = Callback.RefreshList + 1
End Sub
Sub TransparentOff()
On Error Resume Next
If Not glistN Is Nothing Then
glistN.BackStyle = 0
glistN.backcolor = mycolor(dimm2)
glistN.PanPos = 0
glistN.ShowMe
Callback.RefreshList = Callback.RefreshList - 1
End If
End Sub
Sub Shadow(ByVal vNewValue As Variant)
shadowcol = -white
Colors vNewValue
Me.locked = True
Show
End Sub
Sub ShadowOff()

Colors shadowcol
Me.locked = False
Show
End Sub
Public Property Let Default(RHS)
On Error Resume Next
If CBool(RHS) Then Callback.Default = ctrlName
End Property

Private Sub glistN_softSelected(item As Long)
OneClick
End Sub

Private Sub glistN_UnregisterGlist()
On Error Resume Next
If glistN.TabStopSoft Then Callback.LastActive = ctrlName
End Sub
Public Property Let icon(RHS)
glistN.icon = RHS

End Property

Private Sub glistN_PreviewKeyboardUnicode(ByVal a As String)
Callback.SpreadKey a
End Sub
Private Sub glistN_Fkey(a As Integer)
If a > 1000 Then
Callback.SendFKEY a - 1000
Else
Callback.SendFKEY a
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
If Opcode < 0 Then Opcode = 0
If a = 0 Then Exit Sub
a = a - 1000 * shift - 10000 * ctrl - 100000 * alt
Callback.RegisterAcc a, ctrlName, Opcode
End Sub
Private Sub gListN_AccKey(m As Long)
Callback.AccProces m
End Sub
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
