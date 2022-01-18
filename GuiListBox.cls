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
Dim Callback As GuiM2000
' colors
Dim white As Long, black As Long, dimm1 As Long, dimm2 As Long, trans As Boolean, noeditme As Boolean
Private mHideCaret As Boolean, mTitle As String, mfactor As Single
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
'Set Point2Me = Nothing
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

Private Sub Class_Initialize()
mfactor = 1
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

Private Sub glistN_ChangeColumnListItem(item As Long, col As Long, content As String)
If Not mEnabled Then Exit Sub
Dim thispos As Long, VR(4)
thispos = glistN.SelStart
VR(0) = content
VR(1) = thispos
VR(2) = col
VR(3) = item

   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".ValidListCol(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".ValidListCol()", VR()
      End If
      
thispos = VR(1)
content = VR(0)
If thispos <> glistN.SelStart And content <> "" Then glistN.SelStart = thispos: glistN.ShowMe Else lastitem = -1

End Sub

Private Sub glistN_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
Dim oldtab As Long, where As Long, oldscroll As Long


If item = -1 Then
ElseIf maychange And Button = 0 Then
oldtab = glistN.Column
oldscroll = glistN.PanPos
glistN.PanPos = 0
glistN.REALCUR item, CSng(x * dv15), (0), where
If lastitem = item And oldtab = glistN.Column Then
glistN.PanPos = oldscroll
Exit Sub
End If
glistN.mousepointer = 1
If glistN.ListSep(item) Then Exit Sub

If Button = 0 Then OneClick
While glistN.ListJoin(item) And item > 0
item = item - 1
Wend
glistN.ListindexPrivateUse = item
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + "," + Str(item + 1) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click(" + Str(item + 1) + ")"
      End If
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

Private Sub glistN_ReadColumnItem(item As Long, col As Long, content As String)
If Not mEnabled Then Exit Sub
Dim VR(1)
VR(0) = content
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".ExternalCol(" + Str(col) + "," + Str(item) + "," + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".ExternalCol(" + Str(col) + "," + Str(item) + ")", VR()
      End If
    
content = VR(0)
End Sub

Private Sub glistN_ReadColumnProp(item As Long, col As Long, Prop As String, content As String)
If Not mEnabled Then Exit Sub
Dim VR(2)
VR(0) = content
VR(1) = Prop
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".ItemPropertyStr(" + Str(col) + "," + Str(item) + "," + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".ItemPropertyStr(" + Str(col) + "," + Str(item) + ")", VR()
      End If
content = VR(0)

End Sub

Private Sub glistN_ReadColumnPropNum(item As Long, col As Long, Prop As String, content As Variant)
If Not mEnabled Then Exit Sub
Dim VR(2)
VR(0) = content
VR(1) = Prop
   If mIndex >= 0 Then
        Callback.CallbackNow mesCtrlName + ".ItemPropertyNum(" + Str(col) + "," + Str(item) + "," + Str(mIndex) + ")", VR()
   Else
        Callback.CallbackNow mesCtrlName + ".ItemPropertyNum(" + Str(col) + "," + Str(item) + ")", VR()
      End If
content = VR(0)
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
        glistN.NoCaretShow = False
        If Not glistN.EditFlag Then
            glistN.PanPos = 0
        End If
    Else
        glistN.forecolor = white
        glistN.backcolor = dimm1
        glistN.NoCaretShow = True
    End If
End If
If Not glistN.EditFlag Then
glistN.ShowMe
End If
OneClick
End If

End Sub
Public Sub Retired()
If Not trans Then
glistN.EditFlag = False
glistN.forecolor = white
glistN.backcolor = dimm2
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
End If
End Sub

Private Sub glistN_CheckLostFocus()
If Not locked Then
If glistN.EditFlag Then
glistN.DestCaret
End If
If mStayOpen Then
If glistN.Visible Then
If Not glistN.EditFlag Then glistN.PanPos = 0
glistN.PrepareToShow
End If
Exit Sub
End If
If Not trans Then
If glistN.BackStyle <> 1 Then
glistN.forecolor = white
glistN.backcolor = dimm2
End If
glistN.PanPos = 0
glistN.PrepareToShow
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
With glistN
'.dcolor = Callback.Controls(1).dcolor
'.CapColor = Callback.Controls(1).CapColor
End With
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
Public Property Get index() As Long
index = mIndex
End Property

Friend Property Let index(ByVal RHS As Long)
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


Private Sub glistN_KeyDown(keycode As Integer, shift As Integer)
lastshift = (shift And 1) = 0
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
If keycode = vbKeyF10 And shift <> 0 Then
keycode = 0
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
Public Property Get Columns() As Long
Columns = glistN.Columns
End Property

Public Property Let Columns(ByVal RHS As Long)
If RHS > 0 Then
    glistN.Columns = RHS
Else
    MyEr "Wrong number of columns", "Λάθος αριθμός στηλών"
End If
End Property
Public Property Let BorderStyle(RHS)

glistN.BorderStyle = CInt(RHS)
End Property
Public Property Let AdjustColumns(ByVal RHS As Boolean)
glistN.AdjustColumns = RHS
End Property
Public Property Get TopRows() As Long
    TopRows = glistN.TopRows
End Property
Public Property Let TopRows(RHS As Long)
    glistN.TopRows = RHS
End Property
Public Property Get Column() As Long
Column = glistN.Column
End Property

Public Property Let Column(ByVal RHS As Long)
If RHS > 0 And RHS <= glistN.Columns Then
    glistN.Column = RHS
Else
    MyEr "Wrong number of column", "Λάθος αριθμός στήλης"
End If
End Property
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

Private Sub glistN_softSelected(item As Long)
If glistN.PanPos <> 0 Then glistN.PanPos = 0
   If mIndex >= 0 Then
    Callback.Callback mesCtrlName + ".SoftClick(" + Str(mIndex) + "," + Str(item) + ")"
   Else
    Callback.Callback mesCtrlName + ".SoftClick(" + Str(item) + ")"
    End If
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

Property Get ListJoin(item As Long) As Boolean
ListJoin = glistN.ListJoin(item)
End Property
Property Let ListJoin(item As Long, RHS As Boolean)
If RHS Then
    glistN.JoinLine item
Else
    glistN.SplitLine item
End If
End Property
Public Property Get list(index) As Variant
On Error Resume Next
list = glistN.list(CLng(index))
End Property
Property Get ListSelected(item As Long) As Boolean
ListSelected = glistN.ListSelected(item)
End Property
Property Let ListSelected(item As Long, RHS As Boolean)
glistN.ListSelected(item) = RHS
End Property
Property Get ListChecked(item As Long) As Boolean
ListChecked = glistN.ListChecked(item)
End Property
Property Let ListChecked(item As Long, RHS As Boolean)
glistN.ListChecked(item) = RHS
End Property
Property Get ListMenu(item As Long) As Boolean
ListMenu = glistN.ListMenu(item)
End Property

Property Get ListRadio(item As Long) As Boolean
ListRadio = glistN.ListRadio(item)
End Property
Property Let ListRadio(item As Long, RHS As Boolean)
glistN.ListRadio(item) = RHS
End Property
Property Get ListSep(item As Long) As Boolean
ListSep = glistN.ListSep(item)
End Property
Property Let ListSep(item As Long, RHS As Boolean)
glistN.ListSep(item) = RHS
End Property
Public Property Let list(index, ByVal vNewValue As Variant)
On Error Resume Next
glistN.list(CLng(index)) = CStr(vNewValue)
glistN.PrepareToShow
End Property
Sub Add(content$)
glistN.additem content$
End Sub
Sub AddFast(index As Long, content$)
glistN.additemFast content$
End Sub
Sub Insert(content$)
glistN.additemAtListIndex content$
End Sub
Sub Clear()
glistN.Clear
If glistN.Visible Then glistN.ShowMe2
End Sub

Sub Delete(index As Long)
glistN.Removeitem index
End Sub

Public Property Get ListIndex() As Variant
ListIndex = glistN.ListIndex
End Property
Public Sub AddHeight(ByVal RHS As Integer)
If Not glistN Is Nothing Then
    If glistN.ListIndex >= 0 Then
        glistN.AddRowsAtListIndex RHS
    End If
End If
End Sub
Public Sub RemoveHeight(ByVal RHS As Integer)
If Not glistN Is Nothing Then
    If glistN.ListIndex >= 0 Then
        glistN.DropRowsAtListIndex RHS
    End If
End If
End Sub

Public Property Let ListIndex(ByVal vNewValue As Variant)
glistN.ListIndex = CLng(vNewValue)
glistN.ShowPan
End Property
Public Property Let ListIndex2(ByVal vNewValue As Variant)
glistN.ListindexPrivateUse = CLng(vNewValue)
End Property
Public Property Get listcount() As Variant
listcount = glistN.listcount
End Property
Public Sub ShowThis(index As Long, Optional noselect)
glistN.ShowThis index, CBool(noselect)
End Sub

Public Property Get Text() As Variant
If glistN Is Nothing Then Exit Property
Text = glistN.Text
End Property
Public Property Get TextColumn(ByVal RHS As Variant) As Variant
Dim aCol As Long
aCol = CLng(RHS)
If aCol < 1 Then
    MyEr "Columns begin from 1", "Οι στήλες ξεκινούν από το 1"
    Exit Property
End If
If glistN Is Nothing Then Exit Property
TextColumn = glistN.TextAtColumn(aCol)
End Property
Public Property Let TextColumn(ByVal RHS As Variant, ByVal vNewValue As Variant)
On Error Resume Next
Dim aCol As Long
aCol = CLng(RHS)
If aCol < 1 Then
    MyEr "Columns begin from 1", "Οι στήλες ξεκινούν από το 1"
    Exit Property
End If
If glistN Is Nothing Then Exit Property
If glistN.Columns < aCol Then
    glistN.Columns = aCol
End If
glistN.AddTextColumn aCol, CStr(vNewValue)
End Property
Public Property Let NarrowSelect(ByVal vNewValue As Variant)
If glistN Is Nothing Then Exit Property
On Error Resume Next
glistN.NarrowSelect = CBool(vNewValue)
End Property
Public Property Let Text(ByVal vNewValue As Variant)
On Error Resume Next
If glistN Is Nothing Then Exit Property
glistN.Text = CStr(vNewValue)
End Property
Public Sub Sort(Optional SkipChar As Long = 1, Optional SortType, Optional Columns, Optional LowLim, Optional HighLim)
    If Not IsMissing(SortType) Then SortType = CInt(SortType) Else SortType = vbTextCompare
    If Not IsMissing(LowLim) Then
        If LowLim > listcount - 1 Then LowLim = listcount - 1
    Else
        LowLim = 0
    End If
    If IsMissing(HighLim) Then
        HighLim = listcount - 1
    End If
    If HighLim < LowLim Then SwapVariant HighLim, LowLim
    If HighLim > listcount - 1 Then HighLim = listcount - 1

    If glistN.listcount < 2 Then Exit Sub
    If SortType = vbTextCompare Then
        glistN.mSortstyle = vbTextCompare
    Else
        glistN.mSortstyle = vbBinaryCompare
    End If
    glistN.SkipChars = SkipChar
    If IsMissing(Columns) Then Columns = 1
    Dim oEnabled As Boolean
    oEnabled = glistN.enabled
    glistN.enabled = False
    If Not glistN.JoinLinesCount > 0 Then
        glistN.QuickSortExtended CLng(Columns), CLng(LowLim), CLng(HighLim)
    Else
        While glistN.ListJoin(CLng(LowLim))
            LowLim = LowLim - 1
        Wend
        
        While glistN.ListJoin(CLng(HighLim))
            HighLim = HighLim - 1
        Wend
        HighLim = HighLim + glistN.ListMoreLines(CLng(HighLim))
        Dim many As Long
        many = glistN.compact(CLng(LowLim), CLng(HighLim))
        If many = 0 Then
            glistN.QuickSortExtended CLng(Columns), CLng(LowLim), CLng(HighLim)
        Else
            glistN.QuickSortExtended CLng(Columns), CLng(LowLim) + many, CLng(HighLim)
            glistN.Expand CLng(LowLim), CLng(HighLim), many
        End If
    End If
    glistN.enabled = oEnabled
    If oEnabled Then glistN.ListIndex = glistN.ListIndex: glistN.PrepareToShow 2
    
End Sub
Property Let ColumnWidth(ByVal col As Long, ByVal RHS)
    glistN.ColumnWidth(col) = RHS
End Property
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
Sub Colors(white1, Optional black1, Optional dimm11, Optional dimm21, Optional dimm33, Optional dimm34)
If Not glistN Is Nothing Then
white = mycolor(white1)
If Not IsMissing(black1) Then black = mycolor(black1)
If Not IsMissing(dimm11) Then dimm1 = mycolor(dimm11)
If Not IsMissing(dimm21) Then dimm2 = mycolor(dimm21)
If Not IsMissing(dimm34) Then glistN.dcolor = CLng(mycolor(dimm34))
If Not IsMissing(dimm33) Then glistN.CapColor = CLng(mycolor(dimm33))
End If
'&H333333
Retired
End Sub
Property Let Transparent(ok As Boolean)
trans = ok
If ok Then
glistN.NoCaretShow = True
glistN.SkipForm = False
glistN.BackStyle = 1
Callback.RefreshList = Callback.RefreshList + 1
Else
glistN.NoCaretShow = False
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
    glistN.NoCaretShow = False
    noeditme = Not ok
    glistN.ShowMe2
End Property
Property Get Edit() As Boolean
    Edit = glistN.EditFlag
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
glistN.forecolor = black ' rgb(0, 0, 0)
glistN.backcolor = white ' rgb(255, 255, 255)
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
Public Property Let TitleHeight(ByVal vNewValue As Variant)
mfactor = Abs(CSng(vNewValue))
If mfactor < 1 Then mfactor = 1
If mTitle <> "" Then
Dim onelineheight As Long
glistN.HeadLine = mTitle
onelineheight = glistN.HeadlineHeight
glistN.HeadlineHeight = onelineheight * mfactor
glistN.blockheight = True
Else
glistN.HeadlineHeight = 0
End If

End Property
Public Property Let Title(ByVal vNewValue As Variant)
Dim onelineheight As Long
mTitle = vNewValue
glistN.HeadLine = vbNullString
If mTitle <> "" Then
'glistN.HeadLine = mTitle
'onelineheight = glistN.HeadlineHeight
'glistN.HeadlineHeight = onelineheight
glistN.HeadLine = mTitle
onelineheight = glistN.HeadlineHeight
glistN.HeadlineHeight = onelineheight * mfactor
glistN.blockheight = True
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
Public Sub SetGrid(ByVal RHS, Optional ByVal Fcol)
If glistN Is Nothing Then Exit Sub
If Not IsMissing(Fcol) Then
    glistN.GridColor = mycolor(Fcol)
End If
glistN.Grid = CBool(RHS)
glistN.ShowMe2
End Sub
Public Property Get WrapText() As Long
WrapText = glistN.WrapText
End Property

Public Property Let WrapText(ByVal RHS As Long)
glistN.WrapText = RHS <> 0
End Property
Property Get PropAtColumnNum(ByVal item As Long, ByVal thiscolumn As Long, ByVal Prop As String) As Variant
    PropAtColumnNum = glistN.PropAtColumnNum(item, thiscolumn, Prop)
End Property
Property Get PropAtColumn(ByVal item As Long, ByVal thiscolumn As Long, ByVal Prop As String) As String
    PropAtColumn = glistN.PropAtColumn(item, thiscolumn, Prop)
End Property

Public Sub Attrib(ByVal Prop$, ByVal aval, Optional ByVal where, Optional ByVal thiscolumn)
    If IsMissing(where) And IsMissing(thiscolumn) Then
        glistN.PropertyLet Prop$, aval
    ElseIf IsMissing(thiscolumn) Then
        glistN.PropertyLet Prop$, aval, where
    Else
        glistN.PropertyLet Prop$, aval, where, thiscolumn
    End If
End Sub
Public Sub AttribColumn(ByVal Prop$, ByVal aval, ByVal thiscolumn, Optional ByVal fr, Optional ByVal ed)
    If IsMissing(ed) Then ed = listcount - 1
    If IsMissing(fr) Then fr = 0
    If fr > ed Then SwapVariant fr, ed
    If ed >= listcount Then ed = listcount - 1
    If fr < 0 Then fr = 0
    For fr = fr To ed
        glistN.PropertyLet Prop$, aval, fr, thiscolumn
    Next fr
End Sub
Public Sub InsertColumn(ByVal where As Long, Optional ByVal many)
If where > glistN.Columns Then Exit Sub
If IsMissing(many) Then
    glistN.InsertColumn where, glistN.ColumnWidth(where), CVar(1)
Else
    glistN.InsertColumn where, glistN.ColumnWidth(where), many
End If
End Sub
Public Sub AppendColumn(ByVal many As Long, Optional ByVal colWidth)
    If IsMissing(colWidth) Then colWidth = glistN.ColumnWidth(glistN.Columns)
    glistN.InsertColumn glistN.Columns + 1, CLng(colWidth), CVar(many)
End Sub

