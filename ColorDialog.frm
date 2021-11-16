VERSION 5.00
Begin VB.Form ColorDialog 
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
   Icon            =   "ColorDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin M2000.gList gList1 
      Height          =   6750
      Left            =   135
      TabIndex        =   0
      Top             =   720
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   11906
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dcolor          =   65535
      Backcolor       =   3881787
      ForeColor       =   14737632
      CapColor        =   9797738
   End
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   873
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   3881787
      ForeColor       =   16777215
      CapColor        =   16777215
   End
   Begin M2000.gList glist3 
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   7650
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   661
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   8421504
      ForeColor       =   14737632
      CapColor        =   49344
   End
End
Attribute VB_Name = "ColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Public TEXT1 As myTextBox
Attribute TEXT1.VB_VarHelpID = -1
Dim setupxy As Single
Dim Lx As Long, ly As Long, dr As Boolean
Dim scrTwips As Long
Dim bordertop As Long, borderleft As Long
Dim allwidth As Long, itemWidth As Long
Dim colrotate As Long
Private LastActive As Object
Public previewKey As Boolean
Private Sub Form_Activate()

     On Error Resume Next
             If LastActive Is Nothing Then Set LastActive = gList1
              If Typename(ActiveControl) = "gList" Then
                If LastActive Is ActiveControl Then
                Hook hWnd, ActiveControl
                Else
                Hook hWnd, Nothing
                End If
                Else
               
                Hook hWnd, Nothing
                End If
            If LastActive.enabled Then
            If LastActive.Visible Then If Not ActiveControl Is LastActive Then LastActive.SetFocus
         End If

End Sub
Private Sub Form_Deactivate()
UnHook hWnd
End Sub
Private Sub Form_Load()
colrotate = 0
loadfileiamloaded = True
scrTwips = Screen.TwipsPerPixelX
' clear data...
setupxy = 20
gList3.enabled = True
gList3.LeaveonChoose = True
gList3.VerticalCenterText = True
gList3.restrictLines = 1
gList3.PanPos = 0
gList2.enabled = True
gList2.CapColor = rgb(255, 160, 0)
gList2.HeadLine = vbNullString
gList2.FloatList = True
gList2.MoveParent = True
gList3.NoPanRight = False
gList1.NoPanLeft = False
gList1.enabled = True
gList1.NoFreeMoveUpDown = True
gList1.ShowBar = True
gList1.restrictLines = 8
gList1.FreeMouse = True
gList1.StickBar = False
gList1.StickBar = True
gList1.HeadLine = "B | G | R"
Set TEXT1 = New myTextBox
Set TEXT1.Container = gList3
TEXT1.locked = False
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
gList2.HeadLine = ColorSelector
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
If selectorLastX = -1 And selectorLastY = -1 Then

Else
move selectorLastX, selectorLastY
End If
TEXT1 = Right$(PACKLNG(ReturnColor), 6)
gList1.ShowThis ReturnColor + 1
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

Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
If Button = 1 Then

If lastfactor = 0 Then lastfactor = 1

If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
dr = True
mousepointer = vbSizeNWSE
Lx = x
ly = y
End If

Else
If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then
dr = True
mousepointer = vbSizeNWSE
Lx = x
ly = y
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
    If y < (Height - bordertop) Or y > Height Then addy = (y - ly)
    If x < (Width - borderleft) Or x > Width Then addX = (x - Lx)
    
   If Not ExpandWidth Then addX = 0
        If lastfactor = 0 Then lastfactor = 1
        factor = lastfactor

        
  
        once = True
        If Height > VirtualScreenHeight() Then addy = -(Height - VirtualScreenHeight()) + addy
        If Width > VirtualScreenWidth() Then addX = -(Width - VirtualScreenWidth()) + addX
        If (addy + Height) / 8145 > 0.4 And ((Width + addX) / 3690) > 0.4 Then
   
        If addy <> 0 Then SizeDialog = ((addy + Height) / (8145 * DYP / 15))
        lastfactor = ScaleDialogFix(SizeDialog)


        If ((Width * lastfactor / factor + addX) / Height * lastfactor / factor) < (3690 / 8145) Then
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
        gList2.PrepareToShow
      
      
        ly = ly * lastfactor / factor
    
        'End If
        End If
        Else
        Lx = x
        ly = y
   
End If
once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If dr Then Me.mousepointer = 0
dr = False
End Sub

Private Sub Form_Terminate()
Set LastActive = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
UNhookMe
DestroyCaret
selectorLastX = Left
selectorLastY = top
Sleep 200
loadfileiamloaded = False
End Sub




Private Sub gList1_GotFocus()
If gList1.ListIndex = -1 Then gList1.ListIndex = gList1.ScrollFrom
End Sub


Private Sub gList1_HeaderSelected(Button As Integer)
' rotate b g r
Select Case colrotate
Case 0
colrotate = 1
gList1.HeadLine = "G | R | B"
Case 1
colrotate = 2
gList1.HeadLine = "R | B | G"
Case 2
colrotate = 0
gList1.HeadLine = "B | G | R"
End Select

'
TEXT1 = Mid$(TEXT1 + TEXT1, 3, 6)
gList1.ShowThis UNPACKLNG(TEXT1) + 1
End Sub

Private Sub gList1_KeyDown(keycode As Integer, shift As Integer)
If keycode = vbKeyEscape Then

Unload Me
ElseIf keycode = vbKeyRight Then
    If colrotate > 1 Then
    colrotate = 0
    Else
    colrotate = colrotate + 1
    End If

    gList1_HeaderSelected 0
 keycode = 0
ElseIf keycode = vbKeyLeft Then
    
    gList1_HeaderSelected 0
keycode = 0
End If

End Sub



Private Sub gList1_ScrollSelected(item As Long, y As Long)
TEXT1 = Right$("00000" & Hex$(item - 1), 6)
End Sub

Private Sub gList1_selected(item As Long)
TEXT1 = Right$("00000" & Hex$(item - 1), 6)
End Sub

Private Sub gList1_selected2(item As Long)
TEXT1 = Right$("00000" & Hex$(item), 6)
Refresh
glist3_PanLeftRight True
End Sub

Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)

If gList2.DoubleClickCheck(Button, item, x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor, -1) Then
                gList1.enabled = False  '??
                    gList3.enabled = False
            Unload Me
End If
End Sub
 Private Sub PrintItem(mHdc As Long, c As String, R As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_CENTER Or DT_VCENTER)
    DrawText mHdc, StrPtr(c), -1, R, way
    End Sub

Private Sub gList1_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim a As RECT, realitem As Long, v$
If item = -1 Then
'FillThere thisHDC, thisrect, gList1.CapColor
'FillThereMyVersion2 thisHDC, thisrect, &HF0F0F0
'skip = True
Else

CopyFromLParamToRect a, thisrect
Select Case colrotate
Case 0
realitem = item
Case 1
v$ = Right$(PACKLNG(item), 6)
realitem = UNPACKLNG(Mid$(v$ & v$, 5, 6))
Case 2
v$ = Right$(PACKLNG(item), 6)
realitem = UNPACKLNG(Mid$(v$ & v$, 3, 6))
End Select
FillBack thisHDC, a, realitem
gList1.forecolor = &HFFFFFF - realitem
a.top = a.top + 2
PrintItem thisHDC, Right$("00000" & Hex$(item), 6), a
End If
End Sub

Private Sub gList1_ExposeListcount(cListCount As Long)
cListCount = &H1000000 ' all the colors are here
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
            If Not LastActive Is Nothing Then
            If LastActive.enabled Then
                If LastActive.Visible Then
                   If Not gList2.DoubleClickArea(x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor) Then
                    LastActive.SetFocus
                    Else
                        gList1.enabled = False  '??
                        gList3.enabled = False
                        Unload Me
                    End If
                End If
            End If
            End If
End Sub

Private Sub glist3_ChangeListItem(item As Long, content As String)
Dim realitem As Long
If item = 0 Then
Err.Clear
On Error Resume Next
content = Right$(PACKLNG(UNPACKLNG(content)), 6)
If Err.Number > 0 Then
content = gList3.list(0)
Else
gList1.ShowThis UNPACKLNG(content) + 1
End If
End If
End Sub

Private Sub glist3_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If gList3.EditFlag Then Exit Sub
    If gList3.list(0) = vbNullString Then
    gList3.backcolor = &H808080
    gList3.ShowMe2
    Exit Sub
    End If
 
If Button = 1 Then
  gList3.LeftMarginPixels = gList3.WidthPixels - gList3.UserControlTextWidth(gList3.list(0)) / Screen.TwipsPerPixelX
       gList3.backcolor = rgb(0, 160, 0)
    gList3.ShowMe2
Else

    gList3.LeftMarginPixels = lastfactor * 5
  gList3.backcolor = &H808080
   gList3.ShowMe2


End If


End Sub

Private Sub glist3_KeyDown(keycode As Integer, shift As Integer)

If Not gList3.EditFlag Then


gList1.ShowMe2
gList3.SelStart = 1
 gList3.LeftMarginPixels = lastfactor * 5
  gList3.backcolor = &H808080
  
gList3.EditFlag = True
gList3.NoCaretShow = False
gList3.backcolor = &H0
gList3.forecolor = &HFFFFFF
gList3.ShowMe2
ElseIf keycode = vbKeyReturn Then

DestroyCaret
If TEXT1 <> "" Then
gList3.EditFlag = False
gList3.enabled = False

glist3_PanLeftRight True
keycode = 0
End If
End If

End Sub


Private Sub glist3_LostFocus()

gList3.backcolor = &H808080
gList3.ShowMe2
End Sub

Private Sub glist3_PanLeftRight(direction As Boolean)
Dim that As New recDir, TT As Integer
If TEXT1 = vbNullString Then Exit Sub
If direction Then
    Select Case colrotate
    Case 0
    ReturnColor = UNPACKLNG(TEXT1)
    Case 1
     ReturnColor = UNPACKLNG(Mid$(TEXT1 + TEXT1, 5, 6))
    Case 2
    ReturnColor = UNPACKLNG(Mid$(TEXT1 + TEXT1, 3, 6))
    End Select

Unload Me
End If
End Sub

Private Sub gList3_Selected2(item As Long)
If item = -2 Then
    If gList3.PanPos <> 0 Then
    glist3_PanLeftRight (True)
    Exit Sub
    End If
    gList3.LeftMarginPixels = lastfactor * 5
    gList3.backcolor = &H808080
    gList3.forecolor = &HE0E0E0
    gList3.EditFlag = False
    gList3.NoCaretShow = True
ElseIf Not gList1.EditFlag Then
      gList3.LeftMarginPixels = lastfactor * 5
      gList3.backcolor = &H808080
      
    gList3.EditFlag = True
    gList3.NoCaretShow = False
    gList3.backcolor = &H0
    gList3.forecolor = &HFFFFFF
End If
    gList3.ShowMe2
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
gList1.FontSize = 26 * factor * dv15 / 15
factor = gList1.FontSize / 26 / dv15 * 15
ScaleDialogFix = factor
End Function
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)
lastfactor = factor
gList1.addpixels = 10 * factor
gList3.FontSize = 11.25 * factor * dv15 / 15
gList3.LeftMarginPixels = factor * 5
setupxy = 20 * factor

Dim hl As String
hl = gList1.HeadLine
gList1.HeadLine = vbNullString
gList1.HeadLine = hl


bordertop = 10 * scrTwips * factor
borderleft = bordertop
Dim heightTop As Long, heightSelector As Long, HeightPreview As Long, HeightBottom As Long
Dim shapeHeight As Long
heightTop = 30 * factor * scrTwips
HeightBottom = 30 * factor * scrTwips
' some space here
heightSelector = 450 * factor * scrTwips

HeightPreview = 180 * factor * scrTwips
shapeHeight = 160 * factor * scrTwips  ' and width
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

allheight = bordertop + heightTop + bordertop + heightSelector + bordertop + HeightBottom + bordertop

move Left, top, allwidth, allheight
gList2.move borderleft, bordertop, itemWidth, heightTop
gList1.move borderleft, 2 * bordertop + heightTop, itemWidth, heightSelector
gList3.move borderleft, allheight - HeightBottom - bordertop, itemWidth, HeightBottom


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
Set LastActive = gList1
End Sub
Private Sub gList1_UnregisterGlist()
On Error Resume Next
If gList1.TabStopSoft Then Set LastActive = gList1
Set LastGlist = Nothing
If Err.Number > 0 Then gList1.NoWheel = True
End Sub
