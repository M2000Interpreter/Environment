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
'' for every GuiSomething
Public Property Get GetCallBack() As Form
Set GetCallBack = Callback
End Property
Public Property Get index() As Long
index = mIndex
End Property

Public Property Let index(ByVal RHS As Long)
mIndex = RHS
End Property
Private Sub glistN_CtrlPlusF1()
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
End Sub
Public Sub Construct(there As Form, name$)
 mIndex = -1
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set glistN = there.Controls.Add("m2000.glist", ctrlName)
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = name$
End Sub

Public Sub ConstructArray(there As Form, name$, i As Long)
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
 mesCtrlName = name$
End Sub
Public Property Get enabled() As Long
enabled = glistN.enabled
End Property

Public Property Let enabled(ByVal RHS As Long)
If glistN Is Nothing Then Exit Property
glistN.enabled = RHS
End Property
Public Sub Move(X, Y, Optional w, Optional h)
If IsMissing(w) Then
glistN.MoveTwips CLng(X), CLng(Y), 0, 0
ElseIf IsMissing(h) Then
glistN.MoveTwips CLng(X), CLng(Y), CLng(w), 0
Else
glistN.MoveTwips CLng(X), CLng(Y), CLng(w), CLng(h)
End If
End Sub
Public Sub deconstruct()
Shutdown
Set Callback = Nothing
  Set glistN = Nothing
End Sub
Public Sub SetUp()
Dim ff As New StdFont
With Callback
ff.name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With
With glistN
glistN.clear
.additem "  " + mCaptext
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
check1focus = True
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

aa.name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set glistN.Font = aa
End If
End Sub
Private Sub glistN_OnResize()
glistN.LeftMarginPixels = glistN.HeightPixels
End Sub

Private Sub glistN_PanLeftRight(Direction As Boolean)
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
glistN.backcolor = rgb(255, 255, 255)
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
Private Sub glistN_UnregisterGlist()
On Error Resume Next
If glistN.TabStopSoft Then Callback.LastActive = ctrlName
End Sub

Public Property Let icon(RHS)
glistN.icon = RHS

End Property

