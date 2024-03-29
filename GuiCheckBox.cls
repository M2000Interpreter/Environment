VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' From Tools make changed property as default
Implements IControlIndex
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Dim check1focus As Boolean
Dim mCaptext As String
Dim mValue As Boolean
'' FOR EVERY GuiSomething
Public ctrlName As String
Private mesCtrlName As String
Dim mIndex As Long, mStayOpen As Boolean
Dim Callback As GuiM2000
Public SkipFirstClick As Boolean
Dim white As Long, black As Long, dimm1 As Long, dimm2 As Long, shadowcol As Long

'' for every GuiSomething
Friend Property Get GetCallBack() As GuiM2000
Set GetCallBack = Callback
End Property
Private Sub OneClick()
On Error Resume Next
If SkipFirstClick Then glistN.PrepareClick
End Sub
Public Property Get Index() As Long
Index = mIndex
End Property
Friend Property Let Index(ByVal RHS As Long)
mIndex = RHS
End Property
Private Sub glistN_CtrlPlusF1()
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
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

  'Set glistN = there.Controls.Add("m2000.glist", name$ + "_" + LTrim$(Str$(i)))
'CtrlName = name$ + "_" + LTrim$(Str$(i))
 ' If Err.Number > 0 Then
 ctrlName = GetTag$
 Set glistN = there.Controls.Add("m2000.glist", ctrlName)
 
 'End If
 mIndex = i
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
Public Property Get enabled() As Long
enabled = glistN.enabled
End Property

Public Property Let enabled(ByVal RHS As Long)
If glistN Is Nothing Then Exit Property
glistN.enabled = RHS
End Property
Public Sub move(X, y, Optional W, Optional H)
If IsMissing(W) Then
glistN.MoveTwips CLng(X), CLng(y), 0, 0
ElseIf IsMissing(H) Then
glistN.MoveTwips CLng(X), CLng(y), CLng(W), 0
Else
glistN.MoveTwips CLng(X), CLng(y), CLng(W), CLng(H)
End If
End Sub
Public Sub deconstruct()
Shutdown
Set Callback = Nothing
  Set glistN = Nothing
End Sub
Public Sub SetUp()
white = Callback.Controls(1).ForeColor
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
.Clear
.Arrows2Tab = True
.NoArrowDown = True
.NoArrowUp = True
.NoEscapeKey = True
.bypassfirstClick = True
.AddItem "  " + mCaptext
.NoCaretShow = True
.NoPanRight = False
.VerticalCenterText = True
.restrictLines = 1
.MenuItem 1, True, False, mValue
.ListIndex = 0
Set .Font = ff
.ForeColor = Callback.Controls(1).ForeColor
.dcolor = Callback.Controls(1).dcolor
.LeftMarginPixels = .HeightPixels
.ShowMe
.enabled = True
End With
End Sub
Public Property Let Caption(myCap As String)
'
mCaptext = myCap
If Not glistN Is Nothing Then
glistN.list(0) = "  " + mCaptext
glistN.ShowMe
End If
End Property
Public Property Let CheckReset(RHS As Boolean)
' without event
mValue = RHS
If Not glistN Is Nothing Then
glistN.ListSelected(0) = mValue
glistN.ShowMe
End If
End Property



Private Sub Class_Initialize()
mIndex = -1
End Sub

Private Sub Class_Terminate()
Set glistN = Nothing
End Sub

Private Sub glistN_CheckGotFocus()
Callback.RealHover = ctrlName
check1focus = True
'OneClick
End Sub

Private Sub gListN_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
'
If check1focus Then
glistN.FillThere thisHDC, thisrect, rgb(100, 200, 160)
Else
If mStayOpen Then
glistN.FillThere thisHDC, thisrect, rgb(255, 255, 255)
Else
glistN.FillThere thisHDC, thisrect, rgb(200, 120, 60)
End If
End If
skip = False
End Sub


Private Sub glistN_CheckLostFocus()
check1focus = False
glistN.ShowMe
glistN.DblClick
End Sub
Private Sub glistN_MenuChecked(item As Long)
' always item is 1  so 1-1..0 (base 0)
' RaiseEvent Changed(glistN.ListSelected(0))
If item >= 0 Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Changed(" + Str(mIndex) + "," + Str(glistN.ListSelected(0)) + ")"
   Else
      Callback.Callback mesCtrlName + ".Changed(" + Str(glistN.ListSelected(0)) + ")"
      End If
    
End If

End Sub
Public Property Get checked() As Boolean
If Not glistN Is Nothing Then
checked = glistN.ListSelected(0)
Else
' error
End If
End Property

Sub Shutdown()
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
Private Sub glistN_OnResize()
glistN.LeftMarginPixels = glistN.HeightPixels
End Sub

Private Sub glistN_PanLeftRight(direction As Boolean)
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + "," + Str(checked) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click(" + Str(checked) + ")"
      End If
End Sub

Private Sub glistN_Selected2(item As Long)
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + "," + Str(checked) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click(" + Str(checked) + ")"
      End If
      OneClick
End Sub
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
Public Property Let ShowAlways(ByVal vNewValue As Variant)
mStayOpen = CBool(vNewValue)
If mStayOpen Then
glistN.ForeColor = rgb(0, 0, 0)
glistN.BackColor = rgb(255, 255, 255)
End If
End Property
Sub GetFocus()
On Error Resume Next
Callback.Controls(ctrlName).SetFocus
End Sub
Sub Show()
glistN.ShowPan
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

Private Sub glistN_PreviewKeyboardUnicode(A As String)
Callback.SpreadKey A
End Sub
Private Sub glistN_Fkey(A As Integer)
If A > 1000 Then
Callback.SendFKEY A - 1000
Else
Callback.SendFKEY A
End If
End Sub
Public Sub AccKey(A, Optional shift As Boolean, Optional ctrl As Boolean, Optional alt As Boolean, Optional Opcode As Long = 0)
If MyIsNumeric(A) Then
A = CLng(A)
If A < 0 Then Exit Sub
If A > 499 Then Exit Sub
Else
A = UCase(A)
Select Case A
Case "F1" To "F9"
A = 611 + val(Mid(A, 2))
Case Else
A = AscW(A)
If A > 126 Then A = 0
End Select
End If
If Opcode < 0 Then Opcode = 0
If A = 0 Then Exit Sub
A = A - 1000 * shift - 10000 * ctrl - 100000 * alt
Callback.RegisterAcc A, ctrlName, Opcode
End Sub
Private Sub gListN_AccKey(m As Long)
Callback.AccProces m
End Sub
Sub Colors(white1, Optional dimm21, Optional dimm11, Optional dimm33)
On Error Resume Next
If Not glistN Is Nothing Then
glistN.ForeColor = mycolor(white1)
If Not IsMissing(dimm21) Then dimm2 = mycolor(dimm21)

If Not IsMissing(dimm11) Then dimm1 = mycolor(dimm11)

If Not IsMissing(dimm33) Then glistN.CapColor = CLng(mycolor(dimm33))
End If
'Retired
End Sub
Sub ColorsReset()
white = rgb(255, 255, 255)
glistN.ForeColor = rgb(255, 255, 255)
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
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

Private Sub glistN_WindowKey(VbKeyThis As Integer)
Callback.ProcessWinKey VbKeyThis
End Sub
Public Property Let IControlIndex_Index(ByVal RHS As Long)
mIndex = RHS
End Property

Public Property Get IControlIndex_Index() As Long
IControlIndex_Index = mIndex
End Property

Public Sub IControlIndex_UnloadMe()
If glistN Is Nothing Then Exit Sub
Shutdown
Set glistN = Nothing
Callback.RemoveGuiControl ctrlName
mIndex = -1
Set Callback = Nothing
End Sub

