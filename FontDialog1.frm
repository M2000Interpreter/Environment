VERSION 5.00
Begin VB.Form FontDialog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   8145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FontDialog1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin M2000.gList gList1 
      Height          =   3375
      Left            =   135
      TabIndex        =   0
      Top             =   645
      Width           =   3420
      _extentx        =   6033
      _extenty        =   5953
      max             =   1
      vertical        =   -1
      font            =   "FontDialog1.frx":000C
      dcolor          =   65535
      backcolor       =   3881787
      forecolor       =   14737632
      capcolor        =   9797738
   End
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   3420
      _extentx        =   6033
      _extenty        =   873
      max             =   1
      vertical        =   -1
      font            =   "FontDialog1.frx":0030
      backcolor       =   3881787
      forecolor       =   16777215
      capcolor        =   16777215
   End
   Begin M2000.gList glist3 
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   7635
      Width           =   3420
      _extentx        =   6033
      _extenty        =   661
      max             =   1
      vertical        =   -1
      font            =   "FontDialog1.frx":0054
      backcolor       =   8421504
      forecolor       =   14737632
      capcolor        =   49344
   End
   Begin M2000.gList gList4 
      Height          =   3060
      Left            =   75
      TabIndex        =   2
      Top             =   4350
      Width           =   3420
      _extentx        =   6033
      _extenty        =   5398
      max             =   1
      vertical        =   -1
      font            =   "FontDialog1.frx":0078
      dcolor          =   65535
      backcolor       =   3881787
      forecolor       =   14737632
      capcolor        =   16777215
   End
End
Attribute VB_Name = "FontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Public TEXT1 As myTextBox
Attribute TEXT1.VB_VarHelpID = -1
Dim firstpath As Long
Dim setupxy As Single
Dim Lx As Long, lY As Long, dr As Boolean
Dim scrTwips As Long
Dim bordertop As Long, borderleft As Long
Dim allwidth As Long, itemWidth As Long
Private LastActive As String
Public previewKey As Boolean
Private Sub Form_Activate()
    On Error Resume Next
    If LastActive = vbNullString Then LastActive = gList1.Name
    If HOOKTEST <> 0 Then UnHook HOOKTEST
        If Typename(ActiveControl) = "gList" Then
            Hook hWnd, ActiveControl
        Else
            Hook hWnd, Nothing
        End If
        If LastActive <> "" Then
            If Controls(LastActive).enabled Then
            If Controls(LastActive).Visible Then Controls(LastActive).SetFocus
        End If
    End If
End Sub
Private Sub Form_Deactivate()
UnHook hWnd
End Sub
Private Sub Form_Load()
loadfileiamloaded = True
scrTwips = Screen.TwipsPerPixelX
' clear data...
setupxy = 20
gList1.enabled = True
gList2.enabled = True
glist3.enabled = True
glist3.LeaveonChoose = True
glist3.VerticalCenterText = True
glist3.restrictLines = 1
glist3.PanPos = 0
gList2.CapColor = rgb(255, 160, 0)
gList2.HeadLine = vbNullString
gList2.FloatList = True
gList2.MoveParent = True
glist3.NoPanRight = False
gList1.NoCaretShow = True
gList1.NoPanLeft = False
gList1.StickBar = True
gList1.ShowBar = True
gList1.VerticalCenterText = True
With gList4
If DialogLang <> 0 Then
.additemFast "Font Type"
.menuEnabled(0) = False
.additemFast "  Bold"
.additemFast "  Italic"
.MenuItem 2, True, False, ReturnBold, "bold"
.MenuItem 3, True, False, ReturnItalic, "italic"
.additemFast "Font Size"
.menuEnabled(3) = False
.additemFast "  12"
.additemFast "Font Charset Number"
.menuEnabled(5) = False
.additemFast "  0"
.additemFast "Font Charset Table"
.menuEnabled(7) = False
.additemFast "  ANSI - 0"
.additemFast "  Default - 1"
.additemFast "  Greek - 161"
.additemFast "  Turkish - 162"
.additemFast "  Hebrew - 177"
.additemFast "  Arabic - 178"
.additemFast "  East Europe - 238"
.additemFast "  Russian - 204"
.additemFast "  Baltic -186"
.additemFast "Font Size Table"
Else
.additemFast "Τύπος Γραμματοσειράς"
.menuEnabled(0) = False
.additemFast "  Έντονη"
.additemFast "  Πλάγια"
.MenuItem 2, True, False, ReturnBold, "bold"
.MenuItem 3, True, False, ReturnItalic, "italic"
.additemFast "Μέγεθος Γραμματοσειράς"
.menuEnabled(3) = False
.additemFast "  12"
.additemFast "Αριθμός Κωδικοσελίδας"
.menuEnabled(5) = False
.additemFast "  0"
.additemFast "Πίνακας Κωδικοσελίδων"
.menuEnabled(7) = False
.additemFast "  ANSI - 0"
.additemFast "  Default - 1"
.additemFast "  Greek - 161"
.additemFast "  Turkish - 162"
.additemFast "  Hebrew - 177"
.additemFast "  Arabic - 178"
.additemFast "  East Europe - 238"
.additemFast "  Russian - 204"
.additemFast "  Baltic -186"
.additemFast "Πίνακας Μεγεθών"
End If
.menuEnabled(17) = False
.additemFast "  8"
.additemFast "  9"
.additemFast "  10"
.additemFast "  11"
.additemFast "  12"
.additemFast "  14"
.additemFast "  16"
.additemFast "  18"
.additemFast "  20"
.additemFast "  22"
.additemFast "  24"
.additemFast "  26"
.additemFast "  28"
.additemFast "  36"
.additemFast "  48"
.additemFast "  72"
.ShowMe
.ShowBar = True
.StickBar = True
.ShowBar = False
.NoCaretShow = True
.ListindexPrivateUse = 1
End With
  
Set TEXT1 = New myTextBox
Set TEXT1.Container = glist3
 lastfactor = ScaleDialogFix(SizeDialog)
If ExpandWidth Then
If LastWidth = 0 Then LastWidth = -1
Else
LastWidth = -1
End If
If ExpandWidth Then
If LastWidth = 0 Then LastWidth = -1
Else
LastWidth = -1
End If
ScaleDialog lastfactor, LastWidth
gList2.HeadLine = FontSelector
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
If selectorLastX = -1 And selectorLastY = -1 Then

Else
move selectorLastX, selectorLastY
End If
Dim i As Integer

For i = 0 To Screen.FontCount - 1
If Screen.Fonts(i) = ReturnFontName Then
TEXT1 = Screen.Fonts(i)
TEXT1.locked = False
glist3.Font.charset = ReturnCharset
glist3.Font.bold = ReturnBold
glist3.Font.Italic = ReturnItalic
gList1.ListindexPrivateUse = i
Exit For
End If
Next i
If ReturnSize >= 6 Then gList4.list(4) = "  " & CStr(ReturnSize)
gList4.list(6) = "  " & CStr(ReturnCharset)
gList4.enabled = True
gList2.TabStop = False
gList1.ShowMe
TEXT1.locked = False
glist3.ListIndex = 0
glist3.SoftEnterFocus
End Sub



Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)

If Button = 1 Then

If lastfactor = 0 Then lastfactor = 1

If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
dr = True

Lx = x
lY = y
End If

Else
If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then
dr = True
Lx = x
lY = y
End If

End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim addX As Long, addy As Long, factor As Single, once As Boolean
If once Then Exit Sub
If Button = 0 Then dr = False
If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then mousepointer = vbSizeNWSE Else mousepointer = 0
 Else
 If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then mousepointer = vbSizeNWSE Else mousepointer = 0
End If
If dr Then
    If y < (Height - bordertop) Or y > Height Then addy = (y - lY)
    If x < (Width - borderleft) Or x > Width Then addX = (x - Lx)
   If Not ExpandWidth Then addX = 0
        If lastfactor = 0 Then lastfactor = 1
        factor = lastfactor

        
  
        once = True
        If Height > VirtualScreenHeight() Then addy = -(Height - VirtualScreenHeight()) + addy
        If Width > VirtualScreenWidth() Then addX = -(Width - VirtualScreenWidth()) + addX
        If (addy + Height) / (8145 * DYP / 15) > 0.4 And ((Width + addX) / (3690 * DXP / 15)) > 0.4 Then
   
        If addy <> 0 Then SizeDialog = ((addy + Height) / (8145 * DYP / 15))
        lastfactor = ScaleDialogFix(SizeDialog)


        If ((Width * lastfactor / factor + addX) / Height * lastfactor / factor) < (3690 / 8145 * DXP / DYP) Then
        addX = -Width * lastfactor / factor - 1
      
           End If

        If addX = 0 Then
        If lastfactor <> factor Then ScaleDialog lastfactor, Width
        Lx = x
        
        Else
        Lx = x * lastfactor / factor
         ScaleDialog lastfactor, (Width + addX) * lastfactor / factor
         End If

        
         
        
        LastWidth = Width
        gList2.HeadlineHeight = gList2.HeightPixels
        gList2.SoftEnterFocus
      
      
        lY = lY * lastfactor / factor
    
        'End If
        End If
        Else
        Lx = x
        lY = y
   
End If
once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If dr Then Me.mousepointer = 0
dr = False

End Sub
Public Sub UNhookMe()
Set LastGlist = Nothing
UnHook hWnd
End Sub

Private Sub Form_LostFocus()
If HOOKTEST <> 0 Then
UnHook hWnd
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
UNhookMe
DestroyCaret
selectorLastX = Left
selectorLastY = top
Sleep 200
loadfileiamloaded = False
End Sub






Private Sub gList1_ExposeListcount(cListCount As Long)
cListCount = Screen.FontCount
End Sub

Private Sub gList1_GotFocus()
LastActive = gList1.Name
If gList1.ListIndex = -1 Then gList1.ListIndex = gList1.ScrollFrom
End Sub


Private Sub gList1_KeyDown(keycode As Integer, shift As Integer)
If keycode = vbKeyEscape Then
 CancelDialog = True
Unload Me
End If
End Sub



Private Sub gList1_ScrollSelected(item As Long, y As Long)
glist3.Font.Name = Screen.Fonts(item - 1)
glist3.Font.Italic = gList4.ListSelected(2)
glist3.FontSize = 11.25 * lastfactor
glist3.FontBold = gList4.ListSelected(1)
glist3.Font.charset = Trim$(gList4.list(6))
TEXT1 = Screen.Fonts(item - 1)
End Sub

Private Sub gList1_selected(item As Long)
glist3.Font.Name = Screen.Fonts(item - 1)
glist3.Font.Italic = gList4.ListSelected(2)
glist3.FontSize = 11.25 * lastfactor
glist3.FontBold = gList4.ListSelected(1)
glist3.Font.charset = Trim$(gList4.list(6))
TEXT1 = Screen.Fonts(item - 1)
End Sub

Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If gList2.DoubleClickCheck(Button, item, x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor, -1) Then
                      Unload Me
End If

End Sub


Private Sub gList1_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim a As RECT, b As RECT
Dim oldforecolor As Long
oldforecolor = gList1.forecolor
If item = -1 Then
'FillThere thisHDC, thisrect, gList1.CapColor
'FillThereMyVersion2 thisHDC, thisrect, &HF0F0F0
'skip = True
Else
skip = True
CopyFromLParamToRect a, thisrect
CopyFromLParamToRect b, thisrect
a.top = a.top + 2
If gList1.ListIndex = item Then
b.Left = 0
FillBack thisHDC, b, 0
gList1.forecolor = &HFFFFFF
Else
FillBack thisHDC, b, gList1.backcolor

End If

PrintItem thisHDC, Screen.Fonts(item), a
gList1.forecolor = forecolor
End If
End Sub
 Private Sub PrintItem(mHdc As Long, c As String, R As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_VCENTER)
    DrawText mHdc, StrPtr(c), -1, R, way
    End Sub
Private Sub gList2_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = -1 Then
FillThere thisHDC, thisrect, gList2.CapColor
FillThereMyVersion thisHDC, thisrect, &H999999
skip = True
End If
End Sub

Private Sub gList2_MouseUp(x As Single, y As Single)
            On Error Resume Next
            If LastActive <> "" Then
            If Controls(LastActive).enabled Then
            If Controls(LastActive).Visible Then
               If Not gList2.DoubleClickArea(x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor) Then
                Controls(LastActive).SetFocus
                Else
                    Unload Me
                End If
            End If
            End If
            End If
End Sub

Private Sub glist3_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If glist3.EditFlag Then Exit Sub
    If glist3.list(0) = vbNullString Then
    glist3.backcolor = &H808080
    glist3.ShowMe2
    Exit Sub
    End If
 
If Button = 1 Then
  glist3.LeftMarginPixels = glist3.WidthPixels - glist3.UserControlTextWidth(glist3.list(0)) / Screen.TwipsPerPixelX
       glist3.backcolor = rgb(0, 160, 0)
    glist3.ShowMe2
Else

    glist3.LeftMarginPixels = lastfactor * 5
  glist3.backcolor = &H808080
   glist3.ShowMe2


End If


End Sub



Private Sub glist3_LostFocus()
'If gList1.listindex > -1 Then Text1 = gList1.List(gList1.listindex)
glist3.backcolor = &H808080
glist3.ShowMe2
End Sub

Private Sub glist3_PanLeftRight(direction As Boolean)

If TEXT1 = vbNullString Then Exit Sub
If direction Then
ReturnBold = gList4.ListSelected(1)
ReturnItalic = gList4.ListSelected(2)
ReturnSize = val(Trim$(gList4.list(4)))
ReturnCharset = val(Trim$(gList4.list(6)))
If gList1.ListIndex > -1 Then ReturnFontName = Screen.Fonts(gList1.ListIndex)
Unload Me
End If
End Sub

Private Sub gList3_Selected2(item As Long)
If item = -2 Then
If glist3.PanPos <> 0 Then
glist3_PanLeftRight (True)
Exit Sub
End If

 glist3.LeftMarginPixels = lastfactor * 5
glist3.backcolor = &H808080
glist3.forecolor = &HE0E0E0
glist3.EditFlag = False
glist3.NoCaretShow = True
End If
glist3.ShowMe2
End Sub






Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Beep
End If
End Sub
Public Sub FillThereMyVersion2(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT, b As Long
b = CLng(Rnd * 3) + setupxy / 3

CopyFromLParamToRect a, thatRect
a.Left = a.Right - setupxy
a.top = b
a.Bottom = b + setupxy / 5
FillThere thathDC, VarPtr(a), thatbgcolor
a.top = b + setupxy / 5 + setupxy / 10
a.Bottom = b + setupxy \ 2
FillThere thathDC, VarPtr(a), thatbgcolor


End Sub
Public Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT, b As Long
b = 2
CopyFromLParamToRect a, thatRect
a.Left = b
a.Right = setupxy - b
a.top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), 0
b = 5
a.Left = b
a.Right = setupxy - b
a.top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), rgb(255, 160, 0)


End Sub

Private Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT
CopyFromLParamToRect a, thatRect
FillBack thathDC, a, thatbgcolor
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub

Function ScaleDialogFix(ByVal factor As Single) As Single
gList2.FontSize = 14.25 * factor * dv15 / 15
factor = gList2.FontSize / 14.25 / dv15 * 15
gList1.FontSize = 11.25 * factor * dv15 / 15
gList4.FontSize = 11.25 * factor * dv15 / 15

factor = gList1.FontSize / 11.25 / dv15 * 15

ScaleDialogFix = factor
End Function
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)
lastfactor = factor
gList1.addpixels = 10 * factor
glist3.FontSize = 11.25 * factor * dv15 / 15
setupxy = 20 * factor
gList1.LeftMarginPixels = 5 * factor
 glist3.LeftMarginPixels = factor * 5


bordertop = 10 * scrTwips * factor
borderleft = bordertop
Dim heightTop As Long, heightSelector As Long, HeightMenu As Long, HeightBottom As Long
Dim shapeHeight As Long
heightTop = 30 * factor * scrTwips
HeightBottom = 30 * factor * scrTwips
' some space here
heightSelector = 240 * factor * scrTwips
HeightMenu = 180 * factor * scrTwips

' some space here
HeightBottom = 30 * factor * scrTwips
If (NewWidth < 0) Or NewWidth <= (246 * scrTwips * factor) Then
NewWidth = 246 * scrTwips * factor
End If
itemWidth = (NewWidth - 2 * borderleft)
allwidth = NewWidth
Dim allheight As Long
gList2.FloatLimitTop = VirtualScreenHeight() - bordertop - heightTop
gList2.FloatLimitLeft = VirtualScreenWidth() - borderleft * 3

allheight = bordertop + heightTop + bordertop + heightSelector + bordertop + HeightMenu + bordertop + HeightBottom + bordertop

move Left, top, allwidth, allheight
gList2.move borderleft, bordertop, itemWidth, heightTop
gList1.move borderleft, 2 * bordertop + heightTop, itemWidth, heightSelector
gList4.move borderleft, 3 * bordertop + heightTop + heightSelector, itemWidth, HeightMenu
glist3.move borderleft, allheight - HeightBottom - bordertop, itemWidth, HeightBottom


End Sub

Private Sub gList4_ChangeListItem(item As Long, content As String)
Dim content1 As Single
If item = 4 Then
content1 = val("0" & Trim$(content))
If content1 > 144 Then
content = gList4.list(item)
Else
content = "  " & CStr(content1)
End If
ElseIf item = 6 Then
content1 = val("0" & Trim$(content))
If content1 > 255 Then
content = gList4.list(item)
Else
content = "  " & CStr(content1)
End If
End If
End Sub


Private Sub gList4_GotFocus()
If gList4.EditFlag Then gList4.NoCaretShow = False
gList4.ShowMe2
End Sub

Private Sub gList4_LostFocus()
If gList4.EditFlag Then
If val(Trim$(gList4.list(4))) < 6 Then gList4.list(4) = "  6"
End If
gList4.EditFlag = False
gList4.NoCaretShow = True
gList4.ShowMe2
End Sub

Private Sub gList4_MenuChecked(item As Long)
If item = 2 Then
If gList4.ListSelected(1) Then
glist3.Font.bold = True
Else
glist3.Font.bold = False
End If
ElseIf item = 3 Then
If gList4.ListSelected(2) Then
glist3.Font.Italic = True
Else
glist3.Font.Italic = False
End If
End If
glist3.ShowMe2
End Sub


Private Sub gList4_selected(item As Long)

If item = 5 Or item = 7 Then
If Not gList4.EditFlag Then
 gList4.EditFlag = True
 gList4.NoCaretShow = False
 gList4.ShowMe2
 If val(Trim$(gList4.list(4))) < 6 Then gList4.list(4) = "  6"
End If
 Else
 gList4.EditFlag = False
 gList4.NoCaretShow = True
 End If
glist3.Font.charset = Trim$(gList4.list(6))

End Sub
Private Sub gList4_Selected2(item As Long)
Dim t$()
If val(Trim$(gList4.list(4))) < 6 Then gList4.list(4) = "  6"
If item = 4 Or item = 6 Then
If Not gList4.EditFlag Then
 gList4.EditFlag = True
 gList4.NoCaretShow = False
 gList4.ShowMe2
 
End If
Else

 gList4.EditFlag = False
  gList4.NoCaretShow = False
 If item > 7 And item < 17 Then
 t$() = Split(gList4.list(item), " - ")
 gList4.list(6) = "  " + t$(UBound(t$()))
 gList4.ShowMe2
 ElseIf item > 17 Then
 gList4.list(4) = gList4.list(item)
 gList4.ShowMe2
 End If
End If

End Sub
Private Sub gList4_softSelected(item As Long)
gList4_selected item
End Sub

Private Sub glist3_KeyDown(keycode As Integer, shift As Integer)

If Not glist3.EditFlag Then


gList1.ShowMe2
glist3.SelStart = 1
     glist3.LeftMarginPixels = lastfactor * 5
  glist3.backcolor = &H808080
  
glist3.EditFlag = True
glist3.NoCaretShow = False
glist3.backcolor = &H0
glist3.forecolor = &HFFFFFF
glist3.ShowMe2
ElseIf keycode = vbKeyReturn Then

DestroyCaret
If TEXT1 <> "" Then
glist3.EditFlag = False
glist3.enabled = False

glist3_PanLeftRight True
keycode = 0
End If
End If

End Sub
Private Sub gList4_SpecialColor(RGBcolor As Long)
If gList4.EditFlag Then

ElseIf gList4.NoCaretShow Then
RGBcolor = rgb(255, 200, 125)
End If
End Sub
Private Sub gList4_ScrollSelected(item As Long, y As Long)
gList4_selected item
End Sub
Private Sub glist1_RegisterGlist(this As gList)
On Error Resume Next
Set LastGlist = this
If Err.Number > 0 Then this.NoWheel = True
End Sub
Public Sub hookme(this As gList)
If this Is gList1 Then Set LastGlist = this Else Set LastGlist = Nothing
End Sub
Private Sub gList2_RefreshDesktop()
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
End Sub
Private Sub gList1_CheckGotFocus()
LastActive = gList1.Name
End Sub
Private Sub gList1_UnregisterGlist()
On Error Resume Next
If gList1.TabStopSoft Then LastActive = gList1.Name
Set LastGlist = Nothing
If Err.Number > 0 Then gList1.NoWheel = True
End Sub
