VERSION 5.00
Begin VB.Form NeoMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   Icon            =   "NeoMsgBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   ScaleHeight     =   4920
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   375
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
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
      Enabled         =   -1  'True
      Backcolor       =   3881787
      ForeColor       =   16777215
      CapColor        =   16777215
   End
   Begin M2000.gList command1 
      Height          =   525
      Left            =   4590
      TabIndex        =   1
      Top             =   4245
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   926
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      ForeColor       =   16777215
   End
   Begin M2000.gList gList1 
      Height          =   1995
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   3519
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      Backcolor       =   3881787
      ForeColor       =   16777215
   End
   Begin M2000.gList gList3 
      Height          =   315
      Left            =   3135
      TabIndex        =   3
      Top             =   3600
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
   End
   Begin M2000.gList command2 
      Height          =   525
      Left            =   480
      TabIndex        =   2
      Top             =   4200
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   926
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      ForeColor       =   16777215
   End
End
Attribute VB_Name = "NeoMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements InterPress
Public textbox1 As myTextBox
Public WithEvents ListPad As Document, previewKey As Boolean
Attribute ListPad.VB_VarHelpID = -1
Private Type myImage
    Image As StdPicture
    Height As Long
    Width As Long
    top As Long
    Left As Long
End Type
Dim Image1 As myImage
'This is my new MsgBox
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Dim iTop As Long, iLeft As Long, iwidth As Long, iheight As Long
Dim setupxy As Single
Dim lX As Long, lY As Long, dr As Boolean, drmove As Boolean
Dim prevx As Long, prevy As Long
Dim a$
Dim bordertop As Long, borderleft As Long
Dim allheight As Long, allwidth As Long, itemWidth As Long, itemwidth3 As Long, itemwidth2 As Long
Dim height1 As Long, width1 As Long
Dim myOk As myButton
Dim myCancel As myButton
Dim all As Long
Dim novisible As Boolean
Private mModalid As Variant
Dim LastActive As String, yo As Integer

'
Property Get NeverShow() As Boolean
NeverShow = Not novisible
End Property

Private Sub command1_GotFocus()
yo = 0
End Sub
Private Sub command2_GotFocus()
yo = 1
End Sub
Private Sub command1_KeyDown(KeyCode As Integer, shift As Integer)

If KeyCode = 39 Or KeyCode = 40 Then
    KeyCode = 0
ElseIf KeyCode = 37 Or KeyCode = 38 Then
    KeyCode = 0
    If command2.Visible Then command2.SetFocus: LastActive = command2.Name
End If
If KeyCode = vbKeyPause And Not BreakMe Then
ASKINUSE = False
ElseIf KeyCode = vbKeyEscape Then
                AskResponse$ = AskCancel$
                AskCancel$ = vbNullString
            ASKINUSE = False

End If
End Sub
Private Sub command2_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 39 Or KeyCode = 40 Then
KeyCode = 0

If command1.Visible Then command1.SetFocus: LastActive = command1.Name

ElseIf KeyCode = 37 Or KeyCode = 38 Then
    KeyCode = 0
End If
If KeyCode = vbKeyPause And Not BreakMe Then
ASKINUSE = False
ElseIf KeyCode = vbKeyEscape Then
            AskResponse$ = AskCancel$
                AskCancel$ = vbNullString
            ASKINUSE = False

End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
If HOOKTEST <> 0 Then UnHook HOOKTEST
If LastActive = "" Then Exit Sub
Controls(LastActive).SetFocus
LastActive = ""
End Sub

Private Sub Form_Deactivate()
  If ASKINUSE And Not Form2.Visible Then
    If Visible Then
    Me.SetFocus
    End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyPause And Not BreakMe Then
ASKINUSE = False
End If
End Sub

Private Sub Form_Load()
Dim photo As cDIBSection, aPic As StdPicture

novisible = True
''Set LastGlist = Nothing

If AskCancel$ = vbNullString Then command2.Visible = False
gList2.enabled = True

command1.enabled = True
command2.enabled = True
command1.TabIndex = 2
command2.TabIndex = 1
gList2.TabIndex = 3

height1 = 2775 * DYP / 15
width1 = 7920 * DXP / 15
lastfactor = 1
LastWidth = -1

Set textbox1 = New myTextBox
Set textbox1.Container = gList3
textbox1.MaxCharLength = 100
If AskInput Then
gList3.Visible = True
textbox1 = AskStrInput$
textbox1.locked = False
textbox1.enabled = True

Else
gList3.Visible = False  ' new from revision 17 (version 7)
End If
gList1.NoCaretShow = True
gList1.VerticalCenterText = True
gList1.LeftMarginPixels = 8
gList1.enabled = True
Set ListPad = New Document
ListPad = AskText$
If AskDIB$ = vbNullString And AskDIBicon$ <> "" Then AskDIB$ = AskDIBicon$
If AskDIB$ = vbNullString Then
    Set LoadPictureMine = Form3.icon
Else
    If Left$(AskDIB$, 4) = "cDIB" And Len(AskDIB$) > 12 Then
                Set photo = New cDIBSection
               If cDib(AskDIB$, photo) Then
                   photo.GetDpi 96, 96
                   Set LoadPictureMine = photo.Picture
               Else
                   Set LoadPictureMine = Form3.icon
               End If
               Set photo = Nothing
       Else
               If CFname(AskDIB$) <> "" Then
                  Set aPic = LoadMyPicture(GetDosPath(CFname(AskDIB$)), True, gList2.BackColor)
                    If aPic Is Nothing Then Exit Sub
               
                   Set LoadPictureMine = aPic
               Else
                   Set LoadPictureMine = Form3.icon
               End If
    End If
End If
lastfactor = ScaleDialogFix(SizeDialog)
ScaleDialog lastfactor, LastWidth
gList2.enabled = True
gList2.CapColor = rgb(255, 160, 0)
gList2.FloatList = True
gList2.MoveParent = True
gList2.HeadLine = vbNullString
gList2.HeadLine = AskTitle$
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
Set myOk = New myButton
Set myOk.Container = command1

  Set myOk.Callback = Me
  myOk.Index = 1
  If Left$(AskOk$, 1) = "*" Then
  LastActive = command1.Name
  AskOk$ = Mid$(AskOk$, 2)
  End If
  myOk.Caption = AskOk$
myOk.enabled = True
Set myCancel = New myButton
Set myCancel.Container = command2
  If Left$(AskCancel$, 1) = "*" Then
  LastActive = command2.Name
  AskCancel$ = Mid$(AskCancel$, 2)
  End If
myCancel.Caption = AskCancel$
  Set myCancel.Callback = Me
myCancel.enabled = True
ListPad.WrapAgain

all = ListPad.DocLines

gList1.ShowMe
If AskLastX = -1 And AskLastY = -1 Then

Else
move AskLastX, AskLastY
End If
If AskInput Then
gList3.TabIndex = 1
LastActive = gList3.Name
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, shift As Integer, X As Single, y As Single)

If Button = 1 Then
    
    If lastfactor = 0 Then lastfactor = 1

    If bordertop < 150 Then
    If (y > Height - 150 And y < Height) And (X > Width - 150 And X < Width) Then
    dr = True
    MousePointer = vbSizeNWSE
    lX = X
    lY = y
    End If
    
    Else
    If (y > Height - bordertop And y < Height) And (X > Width - borderleft And X < Width) Then
    dr = True
    MousePointer = vbSizeNWSE
    lX = X
    lY = y
    End If
    End If

End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
textbox1.Dereference
myOk.Shutdown
myCancel.Shutdown

gList1.Shutdown
gList2.Shutdown
gList3.Shutdown
command1.Shutdown
command2.Shutdown
novisible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set myOk = Nothing
Set myCancel = Nothing
AskDIB$ = vbNullString
AskOk$ = vbNullString
AskLastX = Left
AskLastY = top
''Sleep 200
ASKINUSE = False
End Sub

Private Sub gList1_ExposeListcount(cListCount As Long)
cListCount = all
End Sub


Private Sub gList2_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = -1 Then
FillThere thisHDC, thisrect, gList2.CapColor
FillThereMyVersion thisHDC, thisrect, &H999999
skip = True
End If
End Sub

Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal X As Long, ByVal y As Long)
If gList2.DoubleClickCheck(Button, item, X, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor, -1) Then
            AskCancel$ = vbNullString
            ASKINUSE = False
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, shift As Integer, X As Single, y As Single)
Dim addX As Long, addy As Long, factor As Single, once As Boolean
If once Then Exit Sub
If Button = 0 Then dr = False: drmove = False
If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (X > Width - 150 And X < Width) Then MousePointer = vbSizeNWSE Else If Not (dr Or drmove) Then MousePointer = 0
 Else
 If (y > Height - bordertop And y < Height) And (X > Width - borderleft And X < Width) Then MousePointer = vbSizeNWSE Else If Not (dr Or drmove) Then MousePointer = 0
End If
If dr Then



If bordertop < 150 Then

        If y < (Height - 150) Or y > Height Then addy = (y - lY)
     If X < (Width - 150) Or X > Width Then addX = (X - lX)
     
Else
    If y < (Height - bordertop) Or y > Height Then addy = (y - lY)
        If X < (Width - borderleft) Or X > Width Then addX = (X - lX)
    End If
    

    
   ''If Not ExpandWidth Then
   addX = 0
        If lastfactor = 0 Then lastfactor = 1
        factor = lastfactor

        
  
        once = True
         If Width > VirtualScreenWidth() Then addX = -(Width - VirtualScreenWidth()) + addX
        If Height > VirtualScreenHeight() Then addy = -(Height - VirtualScreenHeight()) + addy
      
        If (addy + Height) / height1 > 0.4 And ((Width + addX) / width1) > 0.4 Then
   
        If addy <> 0 Then
        If ((addy + Height) / height1) * width1 > VirtualScreenWidth() * 0.9 Then
        addy = 0: addX = 0

        Else
        SizeDialog = ((addy + Height) / height1)
        End If
        End If
        lastfactor = ScaleDialogFix(SizeDialog)


        If ((Width * lastfactor / factor + addX) / Height * lastfactor / factor) < (width1 / height1) Then
        addX = -Width * lastfactor / factor - 1
      
           End If

        If addX = 0 Then
        If lastfactor <> factor Then ScaleDialog lastfactor, Width
        lX = X
        
        Else
        lX = X * lastfactor / factor
         ScaleDialog lastfactor, (Width + addX) * lastfactor / factor
         End If

        
         
        
        LastWidth = Width
              gList2.HeadlineHeight = gList2.HeightPixels
        gList2.PrepareToShow
        gList1.PrepareToShow
          ListPad.WrapAgain
        all = ListPad.DocLines
        lY = lY * lastfactor / factor
        End If
        Else
        lX = X
        lY = y
   
End If
once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, X As Single, y As Single)

If dr Then Me.MousePointer = 0
dr = False
End Sub
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)
Dim aa As Long
On Error Resume Next
lastfactor = factor
setupxy = 16 * factor
bordertop = 8 * dv15 * factor
If AskInput Then aa = 4 Else aa = 5
gList1.StickBar = True
gList1.addpixels = 3 * factor
borderleft = bordertop
allwidth = width1 * factor
allheight = height1 * factor
itemWidth = allwidth - 2 * borderleft
itemwidth3 = (itemWidth - 2 * borderleft) / 3
itemwidth2 = (itemWidth - borderleft) / 2
move Left, top, allwidth, allheight
FontTransparent = False  ' clear background  or false to write over
gList2.move borderleft, bordertop, itemWidth, bordertop * 3
gList2.FloatLimitTop = VirtualScreenHeight() - bordertop - bordertop * 3
gList2.FloatLimitLeft = VirtualScreenWidth() - borderleft * 3

gList1.Width = itemwidth3 * 2 + borderleft
 ListPad.WrapAgain: all = ListPad.DocLines

If AskCancel$ <> "" And ListPad.DocLines < aa Then
gList1.restrictLines = ListPad.DocLines
If AskInput Then
gList1.move borderleft * 2 + itemwidth3, bordertop * (4 + (aa - ListPad.DocLines) * 2), itemwidth3 * 2 + borderleft, bordertop * ListPad.DocLines * 3
Else
gList1.move borderleft * 2 + itemwidth3, bordertop * (4 + (5 - ListPad.DocLines) * 2), itemwidth3 * 2 + borderleft, bordertop * ListPad.DocLines * 3
End If
Else
gList1.restrictLines = aa
If AskInput Then
gList1.move borderleft * 2 + itemwidth3, bordertop * 5, itemwidth3 * 2 + borderleft, bordertop * 9
Else
gList1.move borderleft * 2 + itemwidth3, bordertop * 5, itemwidth3 * 2 + borderleft, bordertop * 12
End If
End If
If AskInput Then
gList3.move borderleft * 2 + itemwidth3, bordertop * 15, itemwidth3 * 2 + borderleft, bordertop * 3

End If
If AskCancel$ <> "" Then
command2.move borderleft, bordertop * 19, itemwidth2, bordertop * 3
command1.move borderleft + itemwidth2 + borderleft, bordertop * 19, itemwidth2, bordertop * 3
Else
command1.move borderleft, bordertop * 19, itemWidth, bordertop * 3
End If
If iwidth = 0 Then iwidth = itemwidth3
If iheight = 0 Then iheight = bordertop * 12
Dim curIwidth As Long, curIheight As Long, sc As Single
If Image1.Width > 0 Then
curIwidth = Image1.Width
curIheight = Image1.Height
iLeft = borderleft
iTop = 5 * bordertop
iwidth = itemwidth3
iheight = bordertop * 12
 Line (0, 0)-(ScaleWidth - dv15, ScaleHeight - dv15), Me.BackColor, BF
If (curIwidth / iwidth) < (curIheight / iheight) Then
sc = curIheight / iheight
ImageMove Image1, iLeft + (iwidth - curIwidth / sc) / 2, iTop, curIwidth / sc, iheight
Else
sc = curIwidth / iwidth
ImageMove Image1, iLeft, iTop + (iheight - curIheight / sc) / 2, iwidth, curIheight / sc
End If
End If
End Sub
Function ScaleDialogFix(ByVal factor As Single) As Single
gList2.FontSize = 14.25 * factor * dv15 / 15
gList1.FontSize = 13.5 * factor * dv15 / 15
gList3.FontSize = 13.5 * factor * dv15 / 15


factor = gList2.FontSize / 14.25 / dv15 * 15
command1.FontSize = 11.75 * factor
factor = gList1.FontSize / 11.75 / dv15 * 15
command2.FontSize = command1.FontSize
ScaleDialogFix = factor
End Function

Public Property Set LoadApicture(aImage As StdPicture)
On Error Resume Next
Dim sc As Double
Set Image1.Image = Nothing
Image1.Width = 0
If aImage.Handle <> 0 Then
Set Image1.Image = aImage
If (aImage.Width / iwidth) < (aImage.Height / iheight) Then
sc = aImage.Height / iheight
ImageMove Image1, iLeft + (iwidth - aImage.Width / sc) / 2, iTop, aImage.Width / sc, iheight
Else
sc = aImage.Width / iwidth
ImageMove Image1, iLeft, iTop + (iheight - aImage.Height / sc) / 2, iwidth, aImage.Height / sc
End If
End If


Image1.Height = aImage.Height
Image1.Width = aImage.Width
End Property

Public Property Set LoadPictureMine(aImage As StdPicture)
On Error Resume Next
Dim sc As Double
Set Image1.Image = Nothing
Image1.Width = 0
If aImage.Handle <> 0 Then
Set Image1.Image = aImage
Image1.Height = aImage.Height
Image1.Width = aImage.Width
End If
End Property
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
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
Private Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT
CopyFromLParamToRect a, thatRect
FillBack thathDC, a, thatbgcolor
End Sub
Private Sub ImageMove(a As myImage, neoTop As Long, NeoLeft As Long, NeoWidth As Long, NeoHeight As Long)
If a.Image Is Nothing Then Exit Sub
If a.Image.Width = 0 Then Exit Sub
If a.Image.Type = vbPicTypeIcon Then
If IsWine Then
    PaintPicture a.Image, neoTop, NeoLeft, NeoWidth, NeoHeight
    PaintPicture a.Image, neoTop, NeoLeft, NeoWidth, NeoHeight
Else
    Dim aa As New cDIBSection
    aa.BackColor = BackColor
    aa.CreateFromPicture a.Image
    aa.ResetBitmapTypeToBITMAP
    PaintPicture aa.Picture, neoTop, NeoLeft, NeoWidth, NeoHeight
    End If
Else
PaintPicture a.Image, neoTop, NeoLeft, NeoWidth, NeoHeight
End If

End Sub



Private Sub gList2_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyEscape Then
                AskCancel$ = vbNullString
            ASKINUSE = False

End If
End Sub

Private Sub gList2_MouseUp(X As Single, y As Single)
If yo = 0 Then
If command1.Visible Then
    If Not gList2.DoubleClickArea(X, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor) Then
    command1.SetFocus
    Else
    AskResponse$ = AskCancel$
     AskCancel$ = vbNullString
            ASKINUSE = False
    End If
End If

Else
If command2.Visible Then
    If Not gList2.DoubleClickArea(X, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor) Then
        command2.SetFocus
    Else
        AskResponse$ = AskCancel$
        AskCancel$ = vbNullString
        ASKINUSE = False
    End If
End If
End If
End Sub

Private Sub gList3_Selected2(item As Long)
 command1.SetFocus
End Sub

Private Sub InterPress_Press(Index As Long)
If Index = 0 Then
AskResponse$ = AskCancel$
AskCancel$ = vbNullString
Else
If AskInput Then AskStrInput$ = textbox1
AskResponse$ = AskOk$
End If

AskOk$ = vbNullString
ASKINUSE = False
End Sub
Private Sub glist1_ReadListItem(item As Long, content As String)

If item >= 0 Then
content = ListPad.TextLine(item + 1)
End If
End Sub
Private Sub ListPad_BreakLine(Data As String, datanext As String)
    gList1.BreakLine Data, datanext
End Sub


Private Sub gList2_RefreshDesktop()
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
End Sub
