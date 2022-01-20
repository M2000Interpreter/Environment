VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Textbox As myTextBox
Attribute Textbox.VB_VarHelpID = -1
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Public WithEvents Pad As Form
Attribute Pad.VB_VarHelpID = -1
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private mEnabled As Boolean
Dim Iam As Boolean
Public UseOnlyTheList As Boolean
'' autocomplete isn't check yet for GuiDropDown...
'' next...time...
Public AutoComplete As Boolean
Public NoCaseSensitive As Boolean
Public SkipFirstClick As Boolean
' gui own
Public ctrlName As String, Hover As Boolean
Private mesCtrlName As String
Dim mIndex As Long, mStayOpen As Boolean, mMenuStyle As Boolean, mMenuWidth As Long
Dim Callback As GuiM2000
Dim mListText As String
Dim lastitem As Long, EditSure As Boolean
Dim acckeys As FastCollection, playnext As Integer, playnextshift As Integer
Friend Property Get GetCallBack() As GuiM2000
Set GetCallBack = Callback
End Property

Private Sub gListN_AccKey(m As Long)
If mMenuStyle Then
    Callback.AccProces m
End If
End Sub

Private Sub glistN_CtrlPlusF1()
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
End Sub
 Friend Sub Construct(there As GuiM2000, Name$)
 Dim PadGui As GuiM2000
 mIndex = -1
 Set Callback = there
 Set PadGui = New GuiM2000
 Set Pad = PadGui
 On Error Resume Next
 Set glistN = Pad.Controls(1)
 If Err.Number > 0 Then
 Set glistN = Pad.Controls(1)
 End If
 PadGui.PopUpMenuVal = True
 With PadGui
.NoHook = True
.gList2.HeadLine = vbNullString
.gList2.HeadLine = vbNullString
.gList2.HeadlineHeight = .gList2.HeightPixels
End With
 Set Textbox = New myTextBox
 ctrlName = GetTag$
 Set Textbox.Container = there.Controls.Add("m2000.glist", ctrlName)
 there.AddGuiControl Me
 glistN.Dynamic
 'glistN.NoWheel = True
 Textbox.glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
 Friend Sub ConstructArray(there As GuiM2000, Name$, i As Long)
 mIndex = i
 Set Callback = there
 Dim PadGui As GuiM2000
 Set PadGui = New GuiM2000
 Set Pad = PadGui

 On Error Resume Next
 Set glistN = Pad.Controls(1)
 If Err.Number > 0 Then
 Set glistN = Pad.Controls(1)
 End If
 PadGui.PopUpMenuVal = True
 With Pad
 .Controls(0).Visible = 0
 .gList2.HeadLine = vbNullString
.gList2.HeadLine = vbNullString
.gList2.HeadlineHeight = .gList2.HeightPixels
End With
 Set Textbox = New myTextBox
  ctrlName = GetTag$
 Set Textbox.Container = there.Controls.Add("m2000.glist", ctrlName)
 there.AddGuiControl Me
 glistN.Dynamic
 'glistN.NoWheel = True
 Textbox.glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub

Public Sub move(x, y, Optional w, Optional h)
With Textbox
If IsMissing(w) Then
.glistN.MoveTwips CLng(x), CLng(y), 0, 0
ElseIf IsMissing(h) Then
.glistN.MoveTwips CLng(x), CLng(y), CLng(w), 0
Else
.glistN.MoveTwips CLng(x), CLng(y), CLng(w), CLng(h)
End If
Pad.move Pad.Left, Pad.top, .glistN.widthtwips, .glistN.HeightTwips * 5
End With
glistN.restrictLines = 5
glistN.MoveTwips 0, 0, Pad.Width, Pad.Height
End Sub
Public Sub deconstruct()
If Not glistN Is Nothing Then glistN.Shutdown
Set glistN = Nothing
If Not Pad Is Nothing Then Unload Pad
If Not Textbox Is Nothing Then Textbox.Dereference
Set Textbox = Nothing
Set Callback = Nothing
Set Pad = Nothing
End Sub
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.Name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set Textbox.glistN.Font = aa
End If
End Sub
Public Property Get index() As Long
index = mIndex
End Property

Public Property Set Container(glistNN As gList)
Set glistN = glistNN
glistNN.enabled = True
glistNN.NoPanRight = True
glistNN.SingleLineSlide = True
glistNN.LeaveonChoose = True
glistNN.LeftMarginPixels = 8
glistNN.VerticalCenterText = True
glistNN.StickBar = True ' try with false - or hold shift to engage false
glistNN.NoFreeMoveUpDown = True
If Not Textbox Is Nothing Then
Textbox.Prompt = ChrW(&H25BC) + " "
End If
End Property
Public Property Get ListText()
ListText = glistN.ListValue
End Property
Public Property Let Label(a$)
Textbox.Prompt = a$ + " " + ChrW(&H25BC) & " "
Textbox.glistN.ShowMe
End Property



Private Sub glistN_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If item = -1 Then

Else
glistN.mousepointer = 1
If lastitem = item Then Exit Sub
If glistN.ListSep(item) Then Exit Sub
glistN.ListindexPrivateUse = item
glistN.ShowMe2
lastitem = item
End If
End Sub
 Private Sub PrintItem(mHdc As Long, c As String, R As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_RIGHT Or DT_VCENTER)
    DrawText mHdc, StrPtr(c), -1, R, way
    End Sub
Private Sub gListN_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim a As RECT, ofg As Long
If acckeys Is Nothing Then Exit Sub
If acckeys.ExistKey(item) Then
CopyFromLParamToRect a, thisrect
ofg = glistN.forecolor
SetTextColor thisHDC, rgb(128, 128, 64)
PrintItem thisHDC, acckeys.Value + " ", a
SetTextColor thisHDC, ofg
'glistN.forecolor = ofg
End If
End Sub

Private Sub glistN_MenuChecked(item As Long)
'
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".MenuChecked(" + Str(mIndex) + "," + Str(glistN.ListIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".MenuChecked(" + Str(glistN.ListIndex) + ")"
      End If
End Sub
Property Let menuEnabled(item As Long, ByVal RHS As Boolean)
If glistN Is Nothing Then Exit Property
glistN.menuEnabled(item) = RHS
End Property

Property Get menuEnabled(item As Long) As Boolean
If glistN Is Nothing Then Exit Property
menuEnabled = Not glistN.ListSep(item)
End Property
Public Function GetMenuId(id$, Pos As Long) As Boolean
'
GetMenuId = glistN.GetMenuId(id$, Pos)
End Function
Property Get id(item As Long) As String
'
id = glistN.id(item)
End Property


Private Sub glistN_RefreshOnly()
glistN.PrepareToShow
End Sub


Private Sub Class_Terminate()
Set Textbox = Nothing
Set glistN = Nothing
If Not Pad Is Nothing Then Unload Pad
Set Pad = Nothing
End Sub

Private Sub glistN_CheckGotFocus()
Textbox.SkipProcAtExitFocus = True
glistN.FreeMouse = True
If Not mMenuStyle Then
If Textbox.glistN.TabStopSoft Then Callback.LastActive = ctrlName
Textbox.glistN.NoCaretShow = True
End If
glistN.forecolor = rgb(0, 0, 0)
glistN.backcolor = rgb(255, 255, 255)
glistN.NoCaretShow = False:
glistN.PanPos = 0
Textbox.glistN.dropkey = False
OneClick
End Sub
Private Sub glistN_CheckLostFocus()
If Textbox.SkipProcAtExitFocus Then
    
        If Not Iam Then Textbox.Retired   'wrong...
        If Not Iam Then glistN.ListindexPrivateUse = -1
        Textbox.SkipProcAtExitFocus = False

Else
Pad.Visible = False
 Callback.RealHover = ""
End If
Iam = True
End Sub

Private Sub glistN_KeyDown(keycode As Integer, shift As Integer)
If keycode = vbKeyLeft Or keycode = vbKeyRight Then
If Textbox.glistN.EditFlag Then
Iam = True
glistN.BypassLeaveonChoose = True
If Textbox.glistN.TabStopSoft Then Callback.LastActive = ctrlName
Textbox.SetFocus
Else

Iam = False
glistN.BypassLeaveonChoose = True

Textbox.SetFocus
Pad.Visible = False
If mMenuStyle Then
playnext = keycode
playnextshift = shift
keycode = 0
shift = 0
Exit Sub
Else
keycode = 9
End If

End If
ElseIf keycode = 9 Then
keycode = 0
Iam = False
glistN.BypassLeaveonChoose = True

Textbox.SetFocus
Pad.Visible = False
End If
End Sub

Private Sub glistN_PanLeftRight(direction As Boolean)
Dim item As Long
If direction = True Then
item = glistN.ListIndex
If mMenuStyle Then
If glistN.ListSep(item) Then Exit Sub
If glistN.ListRadio(item) Then
    glistN.ListSelected(item) = True
    
End If
Else
Textbox.vartext = glistN.list(item): Textbox.SelStart = Len(Textbox.Text) + 1
End If
 Textbox.SetFocus
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".DblClick(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback mesCtrlName + ".DblClick(" + Str(item) + ")"
      End If
  
End If
End Sub

Private Sub glistN_RegisterGlist(this As gList)
this.NoWheel = True
End Sub
Private Sub glistN_UnregisterGlist()
glistN.NoWheel = True
End Sub
Private Sub glistN_ScrollMove(item As Long)
'
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Move(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback mesCtrlName + ".Move(" + Str(item) + ")"
      End If
      OneClick
End Sub

Private Sub glistN_ScrollSelected(item As Long, y As Long)
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Scroll(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback mesCtrlName + ".Scroll(" + Str(item) + ")"
      End If
End Sub


Public Property Get list(index) As Variant
On Error Resume Next
list = glistN.list(CLng(index))
End Property
Public Property Let list(index, RHS As Variant)
On Error Resume Next
glistN.list(CLng(index)) = CStr(RHS)
End Property

Private Sub glistN_selected(item As Long)
If glistN.PanPos <> 0 Then glistN.PanPos = 0
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click(" + Str(item) + ")"
      End If
      
End Sub

Private Sub Pad_Activate()
Callback.RealHover = ctrlName
glistN.GiveSoftFocus

End Sub

Private Sub Textbox_AccKey(m As Long)
Callback.AccProces m
End Sub

Private Sub Textbox_CascadeSelect(item As Long)
' if pad visible close
'If Not Pad.Visible Then ShowMenu
 If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".OpenMenu(" + Str(mIndex) + "," + Str(-1) + ")"
   Else
      Callback.Callback mesCtrlName + ".OpenMenu(" + Str(-1) + ")"
      End If
If menuEnabled(item - 1) Then

ListIndex2 = item - 1
If ListIndex = item - 1 Then

If ListMenu(item - 1) Then
    If ListRadio(item - 1) And ListSelected(item - 1) Then
    Else
        ListSelected(item - 1) = Not ListSelected(item - 1)
' send event
    End If
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".MenuChecked(" + Str(mIndex) + "," + Str(glistN.ListIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".MenuChecked(" + Str(glistN.ListIndex) + ")"
      End If
      Else
      If glistN.PanPos <> 0 Then glistN.PanPos = 0
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".DblClick(" + Str(mIndex) + "," + Str(item - 1) + ")"
   Else
      Callback.Callback mesCtrlName + ".DblClick(" + Str(item - 1) + ")"
      End If

End If

End If
End If
End Sub

Private Sub Textbox_CtrlPlusF1()
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
End Sub

Private Sub Textbox_DeployMenu()
On Error Resume Next
Textbox.SetFocus
ShowMenu
End Sub

Private Sub Textbox_Fkey(a As Integer)
Callback.SendFKEY a
End Sub

Private Sub textbox_KeyDown(keycode As Integer, shift As Integer)
'If Not glistN.enabled Then KeyCode = 0: shift = 0: Exit Sub
Dim b$
Dim gl As Long, x As Long, x1 As gList, XX As Boolean
Iam = True
If keycode = vbKeyRight And Len(MenuGroup) <> 0 Then
Set x1 = Textbox.glistN
Dim obj As gList
With Callback
.LastActive = vbNullString
For x = 1 To .Controls.Count - 1

    If TypeOf .Controls(x) Is gList Then
    If XX Then
        If .Controls(x).enabled And .Controls(x).TabStop Then
            If .Controls(x).MenuGroup = x1.MenuGroup Then
  
            'Textbox.Retired
            '.Controls(x).SetFocus
            If Pad.Visible Then Pad.Visible = False
            .Controls(x).ShowMenu
            .Controls(x).Refresh
            Textbox.Retired
            '.Controls(x).ShowMenu
            
            Exit For
            
            End If
        End If
    Else
    Set obj = .Controls(x)
        If obj Is x1 Then XX = True
        Set obj = Nothing
    End If
    End If
    
Next x

End With
Set x1 = Nothing
ElseIf keycode = vbKeyLeft And Len(MenuGroup) <> 0 Then
Set x1 = Textbox.glistN
With Callback
.LastActive = vbNullString
For x = .Controls.Count - 1 To 1 Step -1
    If TypeOf .Controls(x) Is gList Then
    If XX Then
        If .Controls(x).enabled And .Controls(x).TabStop Then
            If .Controls(x).MenuGroup = x1.MenuGroup Then
         
                    Textbox.Retired
                    .Controls(x).ShowMenu: Exit For
           End If
         
        End If
    Else
     Set obj = .Controls(x)
        If obj Is x1 Then XX = True
        Set obj = Nothing
    End If
    End If
Next x
End With
Set x1 = Nothing
ElseIf keycode = vbKeyDown Or (keycode = vbKeyReturn And UseOnlyTheList) Then
Iam = False
If glistN.listcount > 0 Then Callback.LastActive = vbNullString

If Pad.Visible Then If glistN.Visible Then glistN.enabled = True: Textbox.SkipProcAtExitFocus = True: glistN.SetFocus: Exit Sub

    keycode = 0
 If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".OpenMenu(" + Str(mIndex) + "," + Str(-1) + ")"
   Else
      Callback.Callback mesCtrlName + ".OpenMenu(" + Str(-1) + ")"
      End If
    gl = glistN.listcount
     If gl = 0 Then
     '  If Callback.Visible Then Callback.SetFocus
        Exit Sub
     End If
     If mMenuStyle Then
     If mMenuWidth > 1000 Then
     Pad.Width = mMenuWidth
     Else
     Pad.Width = Callback.Controls(ctrlName).Width
     End If
     Else
Pad.Width = CLng(Callback.Controls(ctrlName).Width / 1.618 * dv15) \ dv15
End If
If Callback.Controls(ctrlName).Height * gl < VirtualScreenHeight() * 2 / 5 Then

Pad.Controls(1).Height = Callback.Controls(ctrlName).Height * gl + 2 * dv15 * glistN.BorderStyle
glistN.restrictLines = gl
Pad.Height = Pad.Controls(1).Height ' Callback.Controls(ctrlName).Height * glistN.restrictLines + 2 * dv15
Else
Pad.Controls(1).Height = VirtualScreenHeight() * 2 / 5 + 2 * dv15 * glistN.BorderStyle
glistN.restrictLines = (VirtualScreenHeight() * 2 / 5) \ Callback.Controls(ctrlName).Height
Pad.Height = Pad.Controls(1).Height ' Callback.Controls(ctrlName).Height * glistN.restrictLines

End If

glistN.AutoHide = False
glistN.StickBar = False
glistN.ShowBar = False
glistN.MoveTwips 0, 0, Pad.Width, Pad.Height

    Textbox.SkipProcAtExitFocus = True
    Textbox.glistN.NoCaretShow = True
    Textbox.glistN.ShowPan
    glistN.LeaveonChoose = True
   ' glistN.SetFocus

     If AutoComplete Then
            glistN.moveto Text
            If glistN.ListIndex < 0 Then
                glistN.ListIndex = 0
            Else
                glistN.ShowThis glistN.ListIndex + 1
            End If
      Else
            If glistN.LastSelected >= 0 Then
                glistN.ListIndex = glistN.LastSelected - 1
            ElseIf mMenuStyle Then
                glistN.ListindexPrivateUseFirstFree 0
            Else
                glistN.moveto Text
                If glistN.ListIndex < 0 Then
                    glistN.ListIndex = 0
                Else
                    glistN.ShowPan
                End If
            End If
    End If

    If mMenuWidth <> 0 Then
    Callback.PopUpPos Pad, Left, top, Callback.Controls(ctrlName).Height
    Else
Callback.PopUpPos Pad, Left + Callback.Controls(ctrlName).Width - Pad.Width, top, Callback.Controls(ctrlName).Height
End If
If Not mMenuStyle Then If Textbox.glistN.TabStopSoft Then Callback.LastActive = ctrlName

ElseIf keycode = vbKeyTab Then
If Not mMenuStyle Then Callback.LastActive = vbNullString
Textbox.SkipProcAtExitFocus = False
Pad.Visible = False
'glistN.Visible = False
ElseIf AutoComplete And keycode = vbKeyRight And Textbox.SelStart >= Len(Textbox.Text) Then
b$ = Myleft$(Textbox.glistN.LastLinePart, 1)
If b$ <> "" Then Textbox.glistN.LastLinePart = Mid$(Textbox.glistN.LastLinePart, 2)
Textbox.Text = Textbox.Text + b$
ElseIf keycode = vbKeyReturn Then
If Edit Then
If Not mMenuStyle Then If Textbox.glistN.TabStopSoft Then Callback.LastActive = ctrlName

  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".DblClick(" + Str(mIndex) + "," + Str(-1) + ")"
   Else
      Callback.Callback mesCtrlName + ".DblClick(" + Str(-1) + ")"
      End If
   End If
End If


End Sub
Function Myleft$(a As Variant, Pos As Variant)
Myleft$ = Mid$(a, 1, Pos)
End Function

Private Sub TextBox_KeyPress(KeyAscii As Integer)
Iam = KeyAscii = 9
If KeyAscii = 9 Then Pad.Visible = False
End Sub

Private Sub textbox_LostFocus()
Textbox.glistN.NoCaretShow = True

If Textbox.glistN.BackStyle = 1 Then Textbox.glistN.CenterText = True

If Not mMenuStyle Then If Callback.LastActive <> ctrlName Then Textbox.glistN.ShowMe: Callback.LastActive = ctrlName

glistN.BypassLeaveonChoose = False
If Iam Then
If Not mMenuStyle Then Pad.Visible = False
Else

If Not glistN.enabled Then

glistN.enabled = True
End If
End If

End Sub
Private Sub textbox_GotFocus()
If playnext <> 0 Then
If Pad.Visible Then

glistN.enabled = False
glistN.BypassLeaveonChoose = False
End If
textbox_KeyDown playnext, playnextshift
playnext = 0
Exit Sub
End If

 If Not mMenuStyle Then
 Textbox.glistN.NoCaretShow = Not Textbox.enabled
   
 End If

 If Textbox.glistN.BackStyle = 1 Then Textbox.glistN.CenterText = False
 
 If Not mMenuStyle Then Textbox.glistN.ShowMe: Textbox.glistN.PrepareToShow 1
 If mMenuStyle Then Textbox.glistN.PrepareClick
If Pad.Visible Then

glistN.enabled = False
glistN.BypassLeaveonChoose = False
End If
End Sub
Private Sub glistN_Selected2(item As Long)
If glistN.ListSep(item) Then Exit Sub
If item >= 0 Then
If mMenuStyle Then
If glistN.ListSep(item) Then Exit Sub
If glistN.ListRadio(item) Then
    glistN.ListSelected(item) = True
    
End If
Else
Textbox.vartext = glistN.list(item): Textbox.SelStart = Len(Textbox.Text) + 1
End If

Textbox.SetFocus

   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".DblClick(" + Str(mIndex) + "," + Str(item) + ")"
   Else
      Callback.Callback mesCtrlName + ".DblClick(" + Str(item) + ")"
      End If
Else
Pad.Visible = False
End If
End Sub


Private Sub textbox_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim gl As Long
playnext = 0
If mMenuStyle Then
    If Screen.ActiveForm Is Nothing Then Exit Sub
    If GetForegroundWindow <> Screen.ActiveForm.hWnd Then
    Exit Sub
    End If
    If Hover Then
    
        If Callback.RealHover = ctrlName Then
            Callback.RealHover = "OPEN"
        ElseIf Callback.RealHover <> "OPEN" Then
           If Callback.RealHover = "" Then
                Callback.RealHover = "bingo"
           Else
                Textbox.SetFocus
           End If
        ElseIf Not Pad.Visible Then
                Textbox.SetFocus
                OneClick
                If glistN.listcount <> 0 Then textbox_KeyDown vbKeyDown, 0
            End If
    Else
        Textbox.OneClick
    End If
    Exit Sub
End If
If Textbox.glistN.DblClick Then
Iam = False

 Textbox.glistN.ShowMe2

 End If

 If Pad.Visible Then Iam = True: Exit Sub



If (Button And 1) = 1 And Not Pad.Visible Then

 If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".OpenMenu(" + Str(mIndex) + "," + Str(-1) + ")"
   Else
      Callback.Callback mesCtrlName + ".OpenMenu(" + Str(-1) + ")"
      End If
gl = glistN.listcount
If gl = 0 Then Exit Sub

Textbox.glistN.dropkey = True
If glistN.Visible Then glistN.enabled = True: Textbox.SkipProcAtExitFocus = True: glistN.SetFocus: Textbox.glistN.dropkey = False: Exit Sub

If mMenuStyle Then
     If mMenuWidth > 1000 Then
     Pad.Width = mMenuWidth
     Else
     Pad.Width = Callback.Controls(ctrlName).Width
     End If
     Else
Pad.Width = CLng(Callback.Controls(ctrlName).Width / 1.618 * dv15) \ dv15
End If

If Callback.Controls(ctrlName).Height * gl < ScrY() * 2 / 5 Then
Pad.Controls(1).Height = Callback.Controls(ctrlName).Height * gl + 2 * dv15 * glistN.BorderStyle
glistN.restrictLines = gl
Pad.Height = Pad.Controls(1).Height    'Pad.Height = Callback.Controls(ctrlName).Height * glistN.restrictLines
Else
Pad.Controls(1).Height = ScrY() * 2 / 5 + 2 * dv15 * glistN.BorderStyle
glistN.restrictLines = (ScrY() * 2 / 5) \ Callback.Controls(ctrlName).Height
Pad.Height = Pad.Controls(1).Height  'Pad.Height = Callback.Controls(ctrlName).Height * glistN.restrictLines
End If

glistN.MoveTwips 0, 0, Pad.Width, Pad.Height
glistN.AutoHide = False
glistN.StickBar = False
glistN.ShowBar = False
    Textbox.SkipProcAtExitFocus = True
    glistN.LeaveonChoose = True
  If AutoComplete Then
        glistN.moveto Text
        If glistN.ListIndex < 0 Then
        glistN.ListIndex = 0
        Else
        glistN.ShowThis glistN.ListIndex + 1
        
         End If
  Else
        If glistN.LastSelected >= 0 Then
             glistN.ListIndex = glistN.LastSelected
    ElseIf mMenuStyle Then
    glistN.ListindexPrivateUseFirstFree 0
    Else
            glistN.moveto Text
                If glistN.ListIndex < 0 Then
                glistN.ListIndex = 0
                Else
                glistN.ShowThis glistN.ListIndex + 1
                
                End If
    End If
    End If
'Hook Pad.hWnd, glistN
    If mMenuWidth <> 0 Then
    Callback.PopUpPos Pad, Left, top, Callback.Controls(ctrlName).Height
    Else
Callback.PopUpPos Pad, Left + Callback.Controls(ctrlName).Width - Pad.Width, top, Callback.Controls(ctrlName).Height
End If
Button = 0
Textbox.glistN.dropkey = False
Iam = Not Pad.Visible

End If

End Sub
Private Sub glistN_SyncKeyboardUnicode(a As String)
Static f As Long
f = glistN.FindItemStartWidth(a, Not NoCaseSensitive, f + 1)
If f >= 0 Then
glistN.ScrollTo f - glistN.lines / 2, f + 1
    If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".PickOther(" + Str(mIndex) + ",{" + glistN.ListValue + "})"
   Else
      Callback.Callback mesCtrlName + ".PickOther({" + glistN.ListValue + "})"
      End If
Else
f = -1
End If
End Sub

Private Sub Textbox_PreviewKeyboardUnicode(a As String)
Callback.SpreadKey a
End Sub

Private Sub TextBox_ValidString(ThatString As String, setpos As Long)
Dim l As Boolean
If Not AutoComplete Then Exit Sub
Dim f As Long
Iam = True
If ThatString = vbNullString And Textbox.glistN.LastLinePart <> "" Then
f = glistN.FindItemStartWidth(Textbox.glistN.LastLinePart, Not NoCaseSensitive, 0)

Else
f = glistN.FindItemStartWidth(ThatString, Not NoCaseSensitive, 0)
End If
If f >= 0 Then
If Not NoCaseSensitive Then ThatString = Myleft$(glistN.list(f), Len(ThatString))
Textbox.glistN.LastLinePart = Mid$(glistN.list(f), Len(ThatString) + 1)
'RaiseEvent AutoCompleteDone(glistN.List(f))
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".AutoCompleteDone(" + Str(mIndex) + ",{" + glistN.list(f) + "})"
   Else
      Callback.Callback mesCtrlName + ".AutoCompleteDone({" + glistN.list(f) + "})"
      End If

glistN.enabled = False
glistN.ScrollTo f - glistN.lines / 2, f + 1
glistN.enabled = True
glistN.ShowMe

Else
If UseOnlyTheList Then
l = Textbox.glistN.enabled
Textbox.glistN.enabled = False
Textbox.glistN.LastLinePart = Textbox.vartext + Textbox.glistN.LastLinePart
ThatString = vbNullString
setpos = 0
Textbox.glistN.enabled = l
Else
Textbox.glistN.LastLinePart = vbNullString
End If
End If
End Sub

Public Property Get Text() As String
Text = Textbox.vartext + Textbox.glistN.LastLinePart
End Property

Public Property Let Text(ByVal RHS As String)
On Error Resume Next
If glistN Is Nothing Then Exit Property
Textbox.glistN.LastLinePart = vbNullString
Textbox.vartext = RHS

End Property
Public Property Get locked() As Boolean
locked = Textbox.locked
End Property

Public Property Let locked(ByVal RHS As Boolean)
If RHS = Textbox.locked Then Exit Property
If RHS Then
If Not Textbox.locked Then
glistN.BypassLeaveonChoose = True

Textbox.SetFocus
glistN.BypassLeaveonChoose = False
End If
End If
Textbox.locked = RHS
End Property
Public Sub additem(a$)
glistN.additemFast a$
Text = a$
End Sub

Public Sub additemFast(a$)
glistN.additemFast a$
End Sub
Public Sub MenuItem(a$, Optional enabledthis As Boolean = True, Optional checked As Boolean = False, Optional radiobutton As Boolean = False, Optional firstate As Boolean = False, Optional ByVal id, Optional ByVal acc, Optional ByVal shift, Optional ByVal ctrl, Optional ByVal alt, Optional legend)
Dim item, disp$
If Not mMenuStyle Then
    MenuStyle = True
End If

If Not a$ = vbNullString Then
glistN.additemFast a$
End If
item = glistN.listcount
If a$ = vbNullString Then
glistN.AddSep
Exit Sub
Else
If IsMissing(id) Then
glistN.MenuItem item, checked, radiobutton, firstate

Else
glistN.MenuItem item, checked, radiobutton, firstate, CStr(id)
End If
glistN.menuEnabled(CLng(item - 1)) = enabledthis
End If
If Not IsMissing(acc) Then
If MyIsNumeric(acc) Then
acc = CLng(acc)
If acc < 0 Then Exit Sub
If acc > 499 Then Exit Sub
If Not IsMissing(legend) Then disp$ = CStr(legend) Else disp$ = "???"
Else
acc = UCase(acc)
If Not IsMissing(legend) Then disp$ = CStr(legend) Else disp$ = acc
Select Case acc
Case "F1" To "F9"
acc = 611 + val(Mid(acc, 2))
Case Else
acc = AscW(acc)
If acc > 126 Then acc = 0
End Select
End If
If acc = 0 Then Exit Sub
If Not IsMissing(alt) Then
acc = acc - 100000 * CBool(alt)
disp$ = "Alt+" + disp$
End If
If Not IsMissing(ctrl) Then
acc = acc - 10000 * CBool(ctrl)
disp$ = "Ctrl+" + disp$
End If
If Not IsMissing(shift) Then
acc = acc - 1000 * CBool(shift)
disp$ = "Shift+" + disp$
End If

Callback.RegisterAcc acc, ctrlName, -item
If acckeys Is Nothing Then Set acckeys = New FastCollection
If acckeys.ExistKey(item - 1) Then acckeys.Value = disp$ Else acckeys.AddKey item - 1, disp$

End If
End Sub
Public Sub MenuRadio(a$, Optional enabledthis As Boolean = True, Optional firstate As Boolean = False, Optional ByVal id, Optional ByVal acc, Optional ByVal shift, Optional ByVal ctrl, Optional ByVal alt, Optional legend)
Dim item, checked As Boolean, disp$
checked = False
If Not mMenuStyle Then
    MenuStyle = True
End If
If Not a$ = vbNullString Then
glistN.additemFast a$
End If
item = glistN.listcount
If a$ = vbNullString Then
glistN.AddSep
Exit Sub
Else
If IsMissing(id) Then
glistN.MenuItem item, True, True, False

Else
glistN.MenuItem item, True, True, False, CStr(id)
End If
If firstate Then glistN.ListSelectedNoRadioCare(CLng(item - 1)) = True
glistN.menuEnabled(CLng(item - 1)) = enabledthis
End If

If Not IsMissing(acc) Then
If MyIsNumeric(acc) Then
acc = CLng(acc)
If acc < 0 Then Exit Sub
If acc > 499 Then Exit Sub
If Not IsMissing(legend) Then disp$ = CStr(legend) Else disp$ = "???"
Else
acc = UCase(acc)
If Not IsMissing(legend) Then disp$ = CStr(legend) Else disp$ = acc
Select Case acc
Case "F1" To "F9"
acc = 611 + val(Mid(acc, 2))
Case Else
acc = AscW(acc)
If acc > 126 Then acc = 0
End Select
End If
If acc = 0 Then Exit Sub
If Not IsMissing(alt) Then
acc = acc - 100000 * CBool(alt)
disp$ = "Alt+" + disp$
End If
If Not IsMissing(ctrl) Then
acc = acc - 10000 * CBool(ctrl)
disp$ = "Ctrl+" + disp$
End If
If Not IsMissing(shift) Then
acc = acc - 1000 * CBool(shift)
disp$ = "Shift+" + disp$
End If
Callback.RegisterAcc acc, ctrlName, -item
If acckeys Is Nothing Then Set acckeys = New FastCollection
If acckeys.ExistKey(item - 1) Then acckeys.Value = disp$ Else acckeys.AddKey item - 1, disp$

End If

End Sub

Property Let ListRadioPrivate(item As Long, RHS As Boolean)
glistN.ListSelectedNoRadioCare(item) = RHS
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
Public Sub Shutdown()
'new
glistN.Shutdown
glistN.enabled = False
Textbox.locked = True

End Sub

Public Sub SetUp()
'white = rgb(255, 255, 255)
'black = 0
'dimm1 = rgb(100, 200, 160)
'dimm2 = rgb(200, 120, 60)

Dim ff As New StdFont
Dim ff1 As New StdFont
With Callback
ff.Name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
ff1.Name = .CtrlFontName
ff1.charset = Form1.DIS.Font.charset
ff1.Size = .CtrlFontSize
ff1.bold = .CtrlFontBold


End With
SkipFirstClick = True
Set Textbox.glistN.Font = ff1
Textbox.glistN.NoEscapeKey = True
Textbox.glistN.bypassfirstClick = True
Textbox.FadePartColor = &H777777
Textbox.Retired
With glistN
.SingleLineSlide = True
.NoPanLeft = True
.NoPanRight = False
.Vertical = True
.VerticalCenterText = True
.Text = mListText
.backcolor = 0 'dimm2

.MoveTwips 0, 0, Pad.Width, Pad.Height

Set .Font = ff
.forecolor = Callback.Controls(1).forecolor

.dcolor = rgb(127, 127, 127)
.NoCaretShow = False
.enabled = True
.StickBar = True
enabled = True
.ShowMe
End With

End Sub

Public Property Let ListText(ByVal vNewValue As Variant)
mListText = vNewValue
glistN.Text = mListText
End Property
Public Property Get enabled() As Boolean
enabled = Textbox.glistN.enabled
End Property

Public Property Let enabled(ByVal RHS As Boolean)
Textbox.glistN.enabled = RHS
End Property

Public Property Get Edit() As Boolean
If Textbox.glistN.EditFlag <> EditSure Then
With Textbox
.glistN.EditFlag = EditSure
.glistN.NoCaretShow = Not EditSure
End With

Else
Edit = EditSure
End If
End Property

Public Property Let Edit(ByVal RHS As Boolean)
EditSure = RHS
With Textbox
.glistN.EditFlag = RHS
.glistN.NoCaretShow = Not RHS
End With
End Property
Public Property Get Left() As Variant
Left = Callback.Controls(ctrlName).Left
End Property

Public Property Let Left(ByVal RHS As Variant)
Callback.Controls(ctrlName).Left = RHS
End Property
Public Property Get top() As Variant
top = Callback.Controls(ctrlName).top
End Property
Public Property Let top(ByVal RHS As Variant)
Callback.Controls(ctrlName).top = RHS
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
Public Property Get Find(ByVal FindThis$) As Variant
    Find = glistN.Find(FindThis)
End Property
Public Property Get ListIndex() As Variant
ListIndex = glistN.ListIndex
End Property

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
Public Property Get MaxCharLength() As Long
MaxCharLength = Textbox.glistN.maxchar
End Property

Public Property Let MaxCharLength(ByVal RHS As Long)
Textbox.glistN.maxchar = RHS
End Property
Public Property Let ShowAlways(ByVal vNewValue As Variant)
mStayOpen = CBool(vNewValue)
Textbox.ShowAlways = CBool(vNewValue)
End Property

Sub GetFocus()
On Error Resume Next
Callback.Controls(ctrlName).SetFocus
End Sub


Public Property Get MenuStyle() As Variant
MenuStyle = mMenuStyle
End Property

Public Property Let MenuStyle(ByVal vNewValue As Variant)
mMenuStyle = vNewValue
If mMenuStyle Then
glistN.BorderStyle = 1
Textbox.SkipFirstClick = True
Textbox.vartext = vbNullString
Textbox.glistN.bypassfirstClick = False
Iam = True
Else
Textbox.SkipFirstClick = False
Textbox.glistN.bypassfirstClick = True
End If
End Property

Public Property Get MenuWidth() As Variant
MenuWidth = mMenuWidth
End Property

Public Property Let MenuWidth(ByVal vNewValue As Variant)
mMenuWidth = vNewValue
End Property
Public Sub Clear()
On Error Resume Next
Iam = False
Pad.Visible = False
glistN.Clear

Textbox.SetFocus
End Sub
Public Property Let Mark(item)
glistN.dcolor = mycolor(item)
End Property
Public Sub Remove(item)
On Error Resume Next
If item < 0 Then Exit Sub
glistN.Removeitem item
End Sub
Public Sub Insert(item, a$)
On Error Resume Next
glistN.ListindexPrivateUse = item
If glistN.ListIndex > -1 Then
glistN.additemAtListIndex a$
End If
End Sub
Public Sub MenuItemAtListIndex(Optional enabledthis As Boolean = True, Optional checked As Boolean = False, Optional radiobutton As Boolean = False, Optional firstate As Boolean = False, Optional IdD)
Dim item
mMenuStyle = True
item = glistN.ListIndex
If item < 0 Then Exit Sub
If IsMissing(IdD) Then
glistN.MenuItem item, checked, radiobutton, firstate

Else
glistN.MenuItem item, checked, radiobutton, firstate, CStr(IdD)
End If
glistN.menuEnabled(CLng(item - 1)) = enabledthis
End Sub


Public Property Get MenuGroup() As Variant
MenuGroup = Textbox.glistN.MenuGroup
End Property

Public Property Let MenuGroup(ByVal RHS As Variant)
Textbox.glistN.MenuGroup = RHS
End Property

Public Property Get BorderStyle() As Variant
BorderStyle = glistN.BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewValue As Variant)
glistN.BorderStyle = vNewValue
End Property
Public Sub ShowMenu()
If mMenuStyle Then
    If glistN.listcount <> 0 Then textbox_KeyDown vbKeyDown, 0
End If
End Sub
Sub Transparent()
Dim x  As Long, y As Long
Textbox.glistN.BackStyle = 1
Textbox.glistN.GetLeftTop x, y
Textbox.glistN.RepaintFromOut Callback.Image, x, y
Callback.RefreshList = Callback.RefreshList + 1
End Sub
Public Property Get FadePartColor() As Variant
FadePartColor = Textbox.FadePartColor
End Property

Public Property Let FadePartColor(ByVal vNewValue As Variant)
On Error Resume Next
Textbox.FadePartColor = CLng(vNewValue)
End Property
Public Property Let Default(RHS)
On Error Resume Next
If CBool(RHS) Then Callback.Default = ctrlName
End Property
Public Property Let icon(RHS)
Textbox.glistN.icon = RHS
End Property
Public Sub SetBarStyle(bColor As Long, Optional bHatch As Long, Optional BHatchColor As Long)
    glistN.BarColor = mycolor(bColor)
    glistN.BarHatch = bHatch
    glistN.BarHatchColor = BHatchColor
End Sub
Public Property Get Linespace() As Variant
Linespace = dv15 * (glistN.addpixels) / 2
End Property

Public Property Let Linespace(ByVal vNewValue As Variant)
    vNewValue = (vNewValue \ 30) * 60
On Error Resume Next
glistN.addpixels = CLng(vNewValue / dv15)
If Err.Number > 0 Then Exit Property
End Property
Private Sub glistN_Fkey(a As Integer)
If a > 1000 Then
Callback.SendFKEY a - 1000
Else
Callback.SendFKEY a
End If
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

Public Property Get WrapText() As Long
WrapText = glistN.WrapText
End Property

Public Property Let WrapText(ByVal RHS As Long)
glistN.WrapText = RHS <> 0
End Property
Property Let Visible(ByVal RHS As Boolean)
    If Not Textbox Is Nothing Then
        If Not Textbox.glistN Is Nothing Then
            If RHS = False Then
                If Pad.Visible Then
                    Iam = False
                    glistN.BypassLeaveonChoose = True
                    Textbox.SetFocus
                    Pad.Visible = False
                End If
            End If
            Textbox.glistN.Visible = RHS
        End If
    End If
End Property
Property Get Visible() As Boolean
    If Not Textbox Is Nothing Then
        If Not Textbox.glistN Is Nothing Then
        Visible = Textbox.glistN.Visible
        End If
    End If
End Property
