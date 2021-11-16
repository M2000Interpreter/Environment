VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Private mEnabled As Boolean
Public ctrlName As String
Private mesCtrlName As String
Dim mListText As String
Dim mIndex As Long, mStayOpen As Boolean
Dim Callback As GuiM2000, Point2Me As Object
' colors
Dim white As Long, black As Long, dimm1 As Long, dimm2 As Long, trans As Boolean, noeditme As Boolean
Private mHideCaret As Boolean, mTitle As String
Private lastshift As Boolean, lastitem As Long
Public maychange As Boolean
Public SkipFirstClick As Boolean
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
Public Sub deconstruct()
Set Point2Me = Nothing
If Not glistN Is Nothing Then glistN.Shutdown
Set Callback = Nothing
Set glistN = Nothing
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

Private Sub Class_Terminate()
Set Point2Me = Nothing
End Sub

Private Sub glistN_BlockCaret(item As Long, blockme As Boolean, skipme As Boolean)
Dim var(1) As Variant
var(0) = item
If mIndex > -1 Then
Callback.CallbackNow mesCtrlName + ".Sep(" + CStr(mIndex) + ")", var()
Else
Callback.CallbackNow mesCtrlName + ".Sep()", var()
End If
If var(0) = -1 Then skipme = True: blockme = True


If (trans And noeditme) Or mHideCaret Then
blockme = True
If mHideCaret Then glistN.NoCaretShow = True
End If
End Sub


Private Sub glistN_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)



If item = -1 Then
ElseIf maychange And Not Button = 2 Then

If lastitem = item Then Exit Sub
glistN.mousepointer = 1
If glistN.ListSep(item) Then Exit Sub

If Button = 0 Then OneClick
glistN.ListindexPrivateUse = item
glistN.ShowMe2
lastitem = item
End If
End Sub

Private Sub glistN_OutPopUp(x As Single, y As Single, myButton As Integer)
If myButton And &H2 = 2 Then
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".PopUp(" + Str(mIndex) + "," + Str(x + Callback.Controls(ctrlName).Left) + "," + Str(y + Callback.Controls(ctrlName).top) + "," + Str(myButton) + "," + Str(glistN.ListIndex + 1) + ")"
   Else
      Callback.Callback mesCtrlName + ".PopUp(" + Str(x + Callback.Controls(ctrlName).Left) + "," + Str(y + Callback.Controls(ctrlName).top) + "," + Str(myButton) + "," + Str(glistN.ListIndex + 1) + ")"
      End If
      End If
End Sub

Private Sub glistN_RefreshOnly()
glistN.PrepareToShow
End Sub
Private Sub glistN_ChangeListItem(item As Long, content As String)
If Not mEnabled Then Exit Sub
Dim thispos As Long, VR(3)
thispos = glistN.SelStart
VR(0) = content
VR(1) = thispos
VR(2) = item
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".ValidList(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".ValidList()", VR()
      End If
      
thispos = VR(1)
content = VR(0)
If thispos <> glistN.SelStart And content <> "" Then glistN.SelStart = thispos: glistN.ShowMe Else lastitem = -1
End Sub
Private Sub glistN_CtrlPlusF1()
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
End Sub
Private Sub glistN_CheckGotFocus()
Callback.RealHover = ctrlName
If Not locked Then
If glistN.ListIndex < 0 And glistN.listcount > 0 Then glistN.ListindexPrivateUse = 0
If Not trans Then
    If enabled Then
    glistN.forecolor = black
    glistN.backcolor = white
    glistN.NoCaretShow = False:
    glistN.PanPos = 0
    Else
    glistN.forecolor = white
    glistN.backcolor = dimm1
    glistN.NoCaretShow = True
    End If
End If
glistN.ShowMe
OneClick
End If

End Sub
Public Sub Retired()
If Not trans Then
glistN.forecolor = white
glistN.backcolor = dimm2
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
End If
End Sub

Private Sub glistN_CheckLostFocus()
If Not locked Then
If mStayOpen Then
If glistN.Visible Then glistN.PanPos = 0: glistN.PrepareToShow
Exit Sub
End If
If Not trans Then
If glistN.BackStyle <> 1 Then
glistN.forecolor = white
glistN.backcolor = dimm2
End If
glistN.PanPos = 0: glistN.PrepareToShow
End If
End If
glistN.DblClick
End Sub
Public Sub SetFocus()
On Error Resume Next
glistN.SetFocus
End Sub
Public Sub SetUp()
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
.NoEscapeKey = True
.bypassfirstClick = True
.SingleLineSlide = True
.NoPanLeft = True
.NoPanRight = False
.Vertical = True
.VerticalCenterText = True
.Text = mListText
.backcolor = dimm2

Set .Font = ff
.forecolor = white
.dcolor = Callback.Controls(1).dcolor
.NoCaretShow = False
.enabled = True
.StickBar = True
enabled = True
ShowAlways = True
.ShowMe
End With
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
Public Property Get Index() As Long
Index = mIndex
End Property

Friend Property Let Index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Property Get locked() As Boolean
locked = Not glistN.enabled
End Property

Public Property Let locked(ByVal RHS As Boolean)
glistN.enabled = Not RHS
End Property

Private Sub glistN_ExposeListcount(cListCount As Long)
' use with readlistitem

If Not mEnabled Then Exit Sub
Dim VR(1)

VR(0) = cListCount

   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".ListCount(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".ListCount()", VR()
      End If
cListCount = VR(0)
End Sub

Private Sub glistN_GetBackPicture(pic As Object)
Set pic = Point2Me
End Sub

Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)
lastshift = (shift And 1) = 0
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
      OneClick
Exit Sub
End If
End Sub

Private Sub glistN_PanLeftRight(direction As Boolean)
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".DblClick(" + Str(mIndex) + "," + Str(glistN.ListIndex) + "," + Str(direction) + ")"
   Else
      Callback.Callback mesCtrlName + ".DblClick(" + Str(glistN.ListIndex) + "," + Str(direction) + ")"
      End If
       ' OneClick
End Sub

Private Sub glistN_ReadListItem(item As Long, content As String)
If Not mEnabled Then Exit Sub
Dim VR(1)


VR(0) = content
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".External(" + Str(item) + "," + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".External(" + Str(item) + ")", VR()
      End If
    
content = VR(0)


End Sub

Private Sub glistN_RefreshDesktop()

If Callback.Visible Then Callback.Refresh

End Sub

Private Sub glistN_RemoveOne(that As String)
'' this is the unicode key
Dim VR(1)


VR(0) = that
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".Unicode(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".Unicode()", VR()
      End If
    
that = VR(0)

End Sub

Private Sub glistN_ScrollMove(item As Long)
'

   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Move(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback mesCtrlName + ".Move(" + Str(item) + ")"
      End If
      glistN.DblClick
End Sub

Private Sub glistN_ScrollSelected(item As Long, y As Long)
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Scroll(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback mesCtrlName + ".Scroll(" + Str(item) + ")"
      End If
     ' OneClick
End Sub

Private Sub glistN_selected(item As Long)
'
If glistN.PanPos <> 0 Then glistN.PanPos = 0
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click(" + Str(item) + ")"
      End If
        'OneClick
End Sub

Private Sub glistN_Selected2(item As Long)
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".DblClick(" + Str(mIndex) + "," + Str(item) + "," + Str(lastshift) + ")"
   Else
      Callback.Callback mesCtrlName + ".DblClick(" + Str(item) + "," + Str(lastshift) + ")"
      End If
      OneClick
End Sub

Private Sub gListN_SpecialColor(RGBcolor As Long)

If Not mEnabled Then Exit Sub
Dim VR(1)

VR(0) = -RGBcolor

   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".Color(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".Color(", VR()
      End If
RGBcolor = mycolor(VR(0))
End Sub
Function Control() As gList
    Set Control = glistN
End Function

Public Property Get ListText() As Variant
ListText = mListText
End Property

Public Property Let ListText(ByVal vNewValue As Variant)
mListText = vNewValue

End Property
Public Property Get enabled() As Boolean
enabled = mEnabled
End Property

Public Property Let enabled(ByVal RHS As Boolean)
mEnabled = RHS
End Property

Public Property Get Linespace() As Variant
Linespace = dv15 * (glistN.addpixels) / 2
End Property

Public Property Let Linespace(ByVal vNewValue As Variant)
    vNewValue = (vNewValue \ 30) * 60
 
glistN.addpixels = vNewValue / dv15
If glistN.Visible Then glistN.ShowMe
End Property
Public Sub Refresh()
glistN.CalcAndShowBar
If glistN.Visible Then glistN.ShowMe2
End Sub

Public Property Get list(Index) As Variant
On Error Resume Next
list = glistN.list(CLng(Index))
End Property

Public Property Let list(Index, ByVal vNewValue As Variant)
On Error Resume Next
glistN.list(CLng(Index)) = CStr(vNewValue)
glistN.PrepareToShow
End Property
Sub Add(content$)
glistN.additem content$
End Sub
Sub AddFast(Index As Long, content$)
glistN.additemFast content$
End Sub
Sub Insert(content$)
glistN.additemAtListIndex content$
End Sub
Sub Clear()
glistN.Clear
If glistN.Visible Then glistN.ShowMe2
End Sub

Sub Delete(Index As Long)
glistN.Removeitem Index
End Sub

Public Property Get ListIndex() As Variant
ListIndex = glistN.ListIndex
End Property

Public Property Let ListIndex(ByVal vNewValue As Variant)
glistN.ListIndex = CLng(vNewValue)
'If glistN.EditFlag Then
glistN.ShowPan
'Else
'glistN.ShowThis CLng(vNewValue) + 1
'End If
End Property
Public Property Let ListIndex2(ByVal vNewValue As Variant)
glistN.ListindexPrivateUse = CLng(vNewValue)
End Property
Public Property Get listcount() As Variant
listcount = glistN.listcount
End Property
Public Sub ShowThis(Index As Long, Optional noselect)
glistN.ShowThis Index, CBool(noselect)
End Sub

Public Property Get Text() As Variant
Text = glistN.Text
End Property

Public Property Let Text(ByVal vNewValue As Variant)
On Error Resume Next
If glistN Is Nothing Then Exit Property
'If CStr(vNewValue) = vbNullString Then
'glistN.Text = " "
'Else
glistN.Text = CStr(vNewValue)
'End If
End Property
Public Sub Sort(Optional SkipChar As Long = 1)
If glistN.listcount < 2 Then Exit Sub
Dim doc As Document, m As Long
Set doc = New Document
doc.textDocFast = glistN.Text
m = glistN.ListIndex
If SkipChar < 1 Then SkipChar = 1
doc.SortDoc SkipChar, 1, glistN.listcount
glistN.Text = doc.textDoc
If m > 0 Then glistN.ListindexPrivateUse = m
glistN.ShowMe2
Set doc = Nothing
End Sub


Public Property Get External() As Variant
External = glistN.BlockItemcount
End Property

Public Property Let External(ByVal vNewValue As Variant)
glistN.BlockItemcount = CBool(vNewValue)
End Property
Public Property Let DisplayLines(ByVal vNewValue As Variant)
glistN.restrictLines = CLng(vNewValue)
End Property
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
dimm1 = dimm11
If Not IsMissing(dimm21) Then dimm2 = dimm21
End If
End If
Retired
End Sub
Property Let Transparent(ok As Boolean)
trans = ok
If ok Then
glistN.NoCaretShow = True
Set Point2Me = Callback.Picture
glistN.SkipForm = False
glistN.BackStyle = 1
Callback.RefreshList = Callback.RefreshList + 1
Else
glistN.NoCaretShow = False
Set Point2Me = Nothing
glistN.SkipForm = True
glistN.BackStyle = 0
Callback.RefreshList = Callback.RefreshList - 1
End If

End Property
Property Let Center(ok As Boolean)
    glistN.CenterText = ok
    glistN.ShowMe2
End Property
Property Let Edit(ok As Boolean)
    glistN.EditFlag = ok
    glistN.NoCaretShow = Not ok
    noeditme = Not ok
    glistN.ShowMe2
End Property
Private Sub glistN_RegisterGlist(this As gList)
On Error Resume Next
Set LastGlist = this
If Err.Number > 0 Then this.NoWheel = True
End Sub

Private Sub glistN_SyncKeyboard(item As Integer)
Dim VR(1)
VR(0) = item
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".Ascii(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".Ascii()", VR()
      End If
    
item = VR(0)
End Sub

Private Sub glistN_UnregisterGlist()
On Error Resume Next
If glistN.TabStopSoft Then Callback.LastActive = ctrlName
Set LastGlist = Nothing
If Err.Number > 0 Then glistN.NoWheel = True
End Sub
Property Let ScrollTo(there As Long)
glistN.ScrollTo there
End Property
Property Get ScrollFrom()
ScrollFrom = glistN.ScrollFrom
End Property
Property Let ShowBar(ok As Boolean)
glistN.ShowBar = ok
If ok Then
End If
End Property

Public Property Let ShowAlways(ByVal vNewValue As Variant)
mStayOpen = CBool(vNewValue)
If mStayOpen Then
glistN.forecolor = rgb(0, 0, 0)
glistN.backcolor = rgb(255, 255, 255)
End If
End Property
Public Property Get Find(ByVal FindThis$) As Variant
    Find = glistN.Find(FindThis)
End Property

Sub GetFocus()
On Error Resume Next
Callback.Controls(ctrlName).SetFocus
End Sub

Public Property Let HideCaret(ByVal vNewValue As Variant)
    mHideCaret = vNewValue
    glistN.NoCaretShow = vNewValue
End Property

Public Property Get Title() As Variant
Title = mTitle
End Property

Public Property Let Title(ByVal vNewValue As Variant)
Dim onelineheight As Long
mTitle = vNewValue
glistN.HeadLine = vbNullString
If mTitle <> "" Then
glistN.HeadLine = mTitle
onelineheight = glistN.HeadlineHeight
glistN.HeadlineHeight = onelineheight
Else
glistN.HeadlineHeight = 0
End If

End Property
Public Property Let PanDouble(ByVal vNewValue As Variant)
glistN.NoPanLeft = Not vNewValue
End Property
Public Property Let Default(RHS)
On Error Resume Next
If CBool(RHS) Then Callback.Default = ctrlName
End Property
Public Property Let icon(RHS)
glistN.icon = RHS
End Property
Public Sub SetBarStyle(bColor As Long, Optional bHatch As Long, Optional BHatchColor As Long)
    glistN.BarColor = mycolor(bColor)
    glistN.BarHatch = bHatch
    glistN.BarHatchColor = BHatchColor
End Sub
Private Sub glistN_PreviewKeyboardUnicode(ByVal a As String)
Callback.SpreadKey a
End Sub
Private Sub OneClick()
On Error Resume Next
If SkipFirstClick Then glistN.PrepareClick
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

