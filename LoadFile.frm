VERSION 5.00
Begin VB.Form LoadFile 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   8145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   BeginProperty Font 
      Name            =   "FreeSans"
      Size            =   14.25
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin m2000.gList gList1 
      Height          =   3600
      Left            =   135
      TabIndex        =   0
      Top             =   720
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   6350
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FreeSans"
         Size            =   11.25
         Charset         =   161
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
   Begin m2000.gList gList2 
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
         Name            =   "FreeSans"
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
   Begin m2000.gList glist3 
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   7635
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   661
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FreeSans"
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
   Begin VB.Image Image1 
      Height          =   2955
      Left            =   240
      Stretch         =   -1  'True
      Top             =   4350
      Width           =   3225
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   2070
      Left            =   720
      Shape           =   3  'Circle
      Top             =   4777
      Width           =   2250
   End
End
Attribute VB_Name = "LoadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long

Public WithEvents mySelector As FileSelector
Attribute mySelector.VB_VarHelpID = -1
Private Declare Function DestroyCaret Lib "user32" () As Long

Public Text1 As myTextBox
Attribute Text1.VB_VarHelpID = -1

Dim iTop As Long, iLeft As Long, iwidth As Long, iheight As Long

Dim nopreview As Boolean
Dim oldLeftMarginPixels As Long
Dim firstpath As Long
Dim setupXY As Single
Dim lx As Long, ly As Long, dr As Boolean
Dim scrTwips As Long
Dim bordertop As Long, borderleft As Long
Dim allwidth As Long, itemWidth As Long

Private Sub Form_Load()
loadfileiamloaded = True
scrTwips = Screen.TwipsPerPixelX
' clear data...
setupXY = 20
Set mySelector = New FileSelector
glist3.LeaveonChoose = True
glist3.VerticalCenterText = True
glist3.restrictLines = 1
glist3.PanPos = 0
firstpath = False
nopreview = False
oldLeftMarginPixels = 0
gList2.CapColor = RGB(255, 160, 0)
gList2.HeadLine = ""

gList2.FloatList = True
gList2.MoveParent = True
glist3.NoPanRight = False
gList1.NoPanLeft = False

Set Text1 = New myTextBox
Set Text1.Container = glist3

nopreview = True
'fHeight = gList1.Height
Set mySelector = New FileSelector
With mySelector
Set .glistN = gList1
Set .Text1 = Text1

.NostateDir = True
End With
If SetUp = SetUpGR Then
With gList1
.VerticalCenterText = True
.additemFast "Ταξινόμηση κατά"
.MenuEnabled(0) = False
.additemFast "  Χρονοσήμανση"
.additemFast "  Όνομα"
.additemFast "  Τύπο"
.MenuItem 2, False, True, False, "time"
.MenuItem 3, False, True, False, "name"
.MenuItem 4, False, True, False, "type"
.addsep
.additemFast "Παρουσίαση"
.MenuEnabled(5) = False
.additemFast "  Ένα φάκελο"
.additemFast "  Έως 3 φακέλους"
.additemFast "  Χωρίς όριο"
.MenuItem 7, True, True, False, "normal"
.MenuItem 8, True, True, False, "3levels"
.MenuItem 9, True, True, False, "recur"
.addsep
.additemFast "Συμπεριφορά"
.MenuEnabled(10) = False
.additemFast "  Σπρώξε τη λίστα"
.additemFast "  Πολλαπλή Επιλογή"
.additemFast "  Επέκταση Πλάτους"
.MenuItem 12, True, False, False, "push"
.MenuItem 13, True, False, False, "multi"
.MenuItem 14, True, False, False, "expand"
.addsep
.additemFast "Πέτα τις Αλλαγές"
.additemFast "Σταμάτα και κλείσε"
.addsep
.additemFast "Πληροφορίες"
.MenuEnabled(18) = False
.additemFast "κύλισε δεξιά το κάτω"
.MenuEnabled(19) = False
.additemFast "πλαίσιο για επιλογή ή"
.MenuEnabled(20) = False
.additemFast "με διπλό κλικ στη λίστα"
.MenuEnabled(21) = False
.addsep
.additemFast "Γιώργος Καρράς 2014"
.MenuEnabled(23) = False

oldLeftMarginPixels = .LeftMarginPixels + 10
.LeftMarginPixels = 0
PlaceSettings
ReadSettings
End With

Else
With gList1
.VerticalCenterText = True
.additemFast "Sort Type"
.MenuEnabled(0) = False
.additemFast "  By Time Stamp"
.additemFast "  By Name"
.additemFast "  By Type"
.MenuItem 2, False, True, False, "time"
.MenuItem 3, False, True, False, "name"
.MenuItem 4, False, True, False, "type"
.addsep
.additemFast "Performance"
.MenuEnabled(5) = False
.additemFast "  Normal"
.additemFast "  Recursive 3 levels"
.additemFast "  Recursive"
.MenuItem 7, True, True, False, "normal"
.MenuItem 8, True, True, False, "3levels"
.MenuItem 9, True, True, False, "recur"
.addsep
.additemFast "Behavior"
.MenuEnabled(10) = False
.additemFast "  Push to Scroll"
.additemFast "  MultiSelect"
.additemFast "  Expand Width"
.MenuItem 12, True, False, False, "push"
.MenuItem 13, True, False, False, "multi"
.MenuItem 14, True, False, False, "expand"
.addsep
.additemFast "Undo Changes"
.additemFast "Abord and Exit"
.addsep
.additemFast "Information"
.MenuEnabled(18) = False
.additemFast "slide right in the textbox"
.MenuEnabled(19) = False
.additemFast "down side to return file"
.MenuEnabled(20) = False
.additemFast "or double click the file list"
.MenuEnabled(21) = False
.addsep
.additemFast "George Karras 2014"
.MenuEnabled(23) = False

oldLeftMarginPixels = .LeftMarginPixels + 10
.LeftMarginPixels = 0
PlaceSettings
ReadSettings
End With

End If
With mySelector
.NostateDir = False
 lastfactor = ScaleDialogFix(SizeDialog)
If ExpandWidth Then
If LastWidth = 0 Then LastWidth = -1
Else
LastWidth = -1
End If
ScaleDialog lastfactor, DialogPreview, LastWidth
UserFileName = .Mydir.ExtractName(UserFileName)
.FileTypesToDisplay = FileTypesShow
.Mydir.Nofiles = FolderOnly
.Mydir.TopFolder = TopFolder
If ReturnFile <> "" Then
.selectedFile = .Mydir.ExtractName(ReturnFile)
glist3.ShowMe
.FilePath = .Mydir.ExtractPath(ReturnFile)
.Text1 = .Mydir.ExtractName(ReturnFile)
Else
.FilePath = .Mydir.ExtractPath(TopFolder)
End If
End With

If FolderOnly Then
gList2.HeadLine = SelectFolderCaption
ElseIf SaveDialog Then
gList2.HeadLine = SaveFileCaption
Else
gList2.HeadLine = LoadFileCaption
End If
gList2.HeadlineHeight = gList2.HeightPixels
gList2.PrepareToShow
If selectorLastX = -1 And selectorlasty = -1 Then
Else
Move selectorLastX, selectorlasty
End If

End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = 1 Then
If lastfactor = 0 Then lastfactor = 1
If (Y > Height - bordertop And Y < Height) And (x > Width - borderleft And x < Width) Then
dr = True
lx = x
ly = Y
End If
End If
'End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim addX As Long, addy As Long, factor As Single, once As Boolean
If once Then Exit Sub
If Button = 0 Then dr = False
If dr Then
    If Y < (Height - bordertop) Or Y > Height Then addy = (Y - ly)
    If x < (Width - borderleft) Or x > Width Then addX = (x - lx)
    
   If Not ExpandWidth Then addX = 0
        If lastfactor = 0 Then lastfactor = 1
        factor = lastfactor

        
  
        once = True
        If Height > Screen.Height Then addy = -(Height - Screen.Height) + addy
        If Width > Screen.Width Then addX = -(Width - Screen.Width) + addX
        If (addy + Height) / 8145 > 0.4 And ((Width + addX) / 3690) > 0.4 Then
   
        If addy <> 0 Then SizeDialog = ((addy + Height) / 8145)
        lastfactor = ScaleDialogFix(SizeDialog)


        If ((Width * lastfactor / factor + addX) / Height * lastfactor / factor) < (3690 / 8145) Then
        addX = -Width * lastfactor / factor - 1
      
           End If

        If addX = 0 Then
        If lastfactor <> factor Then ScaleDialog lastfactor, DialogPreview, Width
        lx = x
        
        Else
        lx = x * lastfactor / factor
         ScaleDialog lastfactor, DialogPreview, (Width + addX) * lastfactor / factor
         End If

        
         
        
        LastWidth = Width
        gList2.HeadlineHeight = gList2.HeightPixels
        gList2.PrepareToShow
        mySelector.ResetHeightSelector
        gList1.PrepareToShow
       
      
        ly = ly * lastfactor / factor
    
        'End If
        End If
        Else
        lx = x
        ly = Y
   
End If
once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
dr = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
DestroyCaret
Dim i As Long, filetosave As String
If mySelector.mselChecked Then
ReturnListOfFiles = ""
With mySelector
For i = 0 To .Mydir.listcount - 1
' prepare list for files
If .Mydir.ReadMark(i) Then
    If .recnowchecked Then
 filetosave = .Mydir.FindFolder(i) + .Mydir.List(i)
    Else
    filetosave = .Mydir.path + .Mydir.List(i)
    End If
If ReturnListOfFiles = "" Then
ReturnListOfFiles = filetosave
Else
ReturnListOfFiles = ReturnListOfFiles + "#" + filetosave
End If
End If
Next i
End With
ElseIf FolderOnly Then
If ReturnFile <> "" Then
If Not mySelector.Mydir.isdir(ReturnFile) Then MakeFolder ReturnFile
If Not mySelector.Mydir.isdir(ReturnFile) Then ReturnFile = ""
End If
End If

selectorLastX = Left
selectorlasty = Top
Sleep 200
loadfileiamloaded = False
End Sub
Private Sub MakeFolder(ByVal a$)
a$ = Left$(a$, Len(a$) - 1)
On Error Resume Next
MkDir a$
Sleep 100
End Sub

Private Sub gList1_ExposeItemMouseMove(Button As Integer, ByVal Item As Long, ByVal x As Long, ByVal Y As Long)

Static doubleclick As Long
Static dirlistindex As Long
Static dirlisttop As Long
If mySelector.IamBusy Then Exit Sub
If Item = -1 Then
    If Button = 1 And x > gList1.WidthPixels - setupXY And Y < setupXY Then
    doubleclick = doubleclick + 1
      If doubleclick > 1 Then
      doubleclick = 0
If Not gList1.HeadLine = SetUp Then
dirlisttop = gList1.ScrollFrom
dirlistindex = gList1.listindex
mySelector.NostateDir = True
gList1.LeftMarginPixels = oldLeftMarginPixels
gList1.HeadLine = "" ' reset
gList1.HeadLine = SetUp
gList1.ScrollTo 0
gList1.ListindexPrivateUse = 1
gList1.ShowMe
Else
GetSettings
If Not ReadSettings Then
mySelector.NostateDir = False
gList1.ScrollTo dirlistindex
gList1.ListindexPrivateUse = dirlistindex
gList1.LeftMarginPixels = 0
gList1.HeadLine = ""
gList1.HeadLine = " "
mySelector.ResetHeightSelector
gList1.PrepareToShow
Else
mySelector.NostateDir = False
gList1.LeftMarginPixels = 0
gList1.HeadLine = ""
gList1.HeadLine = " "
mySelector.reload
mySelector.ResetHeightSelector
gList1.PrepareToShow
End If
End If
End If


    End If
Else
doubleclick = 0

End If
End Sub

Private Sub gList1_GotFocus()
If gList1.listindex = -1 Then gList1.listindex = gList1.ScrollFrom
End Sub

Private Sub gList1_HeaderSelected(Button As Integer)
If Button = 1 And Not mySelector.NostateDir Then
gList1.CapColor = RGB(0, 160, 0)
gList1.ShowMe2
gList1.refresh
mySelector.reload
gList1.CapColor = RGB(106, 128, 149)
gList1.ShowMe2
End If
End Sub

Private Sub gList1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
mySelector.AbordAll
Unload Me
End If
End Sub

Private Sub gList1_Selected2(Item As Long)
If mySelector.NostateDir = True Then
' we ar in setup
Select Case Item
Case 15
PlaceSettings
gList1.ScrollTo 0
Case 16
With mySelector
.AbordAll
.selectedFile = ""
End With
Unload Me
End Select
End If
End Sub

Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal Item As Long, ByVal x As Long, ByVal Y As Long)
Static doubleclick As Long

If Item = -1 Then
If Button = 1 Then

End If
    If Button = 1 And x < setupXY And Y < setupXY Then
    doubleclick = doubleclick + 1
    If doubleclick > 1 Then
    mySelector.AbordAll
    Unload Me: Button = 0
    End If
    End If
Else
doubleclick = 0
End If
End Sub


Private Sub gList1_ExposeRect(ByVal Item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If Item = -1 Then
mySelector.FillThere thisHDC, thisrect, gList1.CapColor
FillThereMyVersion2 thisHDC, thisrect, &HF0F0F0
skip = True
End If
End Sub
Private Sub gList2_ExposeRect(ByVal Item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If Item = -1 Then
mySelector.FillThere thisHDC, thisrect, gList2.CapColor
FillThereMyVersion thisHDC, thisrect, &H999999
skip = True
End If
End Sub




Private Sub glist3_ExposeItemMouseMove(Button As Integer, ByVal Item As Long, ByVal x As Long, ByVal Y As Long)
If glist3.EditFlag Then Exit Sub
    If glist3.List(0) = "" Then
    glist3.Backcolor = &H808080
    glist3.ShowMe2
    Exit Sub
    End If
 
If Button = 1 Then
  glist3.LeftMarginPixels = glist3.WidthPixels - glist3.UserControlTextWidth(glist3.List(0)) / Screen.TwipsPerPixelX
       glist3.Backcolor = RGB(0, 160, 0)
    glist3.ShowMe2
Else

     glist3.LeftMarginPixels = lastfactor * 5
  glist3.Backcolor = &H808080
   glist3.ShowMe2


End If


End Sub

Private Sub glist3_GotFocus()
'
End Sub

Private Sub glist3_KeyDown(KeyCode As Integer, Shift As Integer)
If Not mySelector.Mydir.isReadOnly(mySelector.Mydir.path) Then
If Not glist3.EditFlag Then

If NewFolder Then

If Not (gList1.listindex = -1) Then
gList1.listindex = -1
gList1.ShowMe2
glist3.clear
glist3.SelStart = 1
Text1 = "NewFolder"
End If
     glist3.LeftMarginPixels = lastfactor * 5
  glist3.Backcolor = &H808080
  
glist3.EditFlag = True
glist3.NoCaretShow = False
glist3.Backcolor = &H0
glist3.ForeColor = &HFFFFFF
ElseIf Not FileExist Then
If Not (gList1.listindex = -1) Then
gList1.listindex = -1
gList1.ShowMe2
glist3.clear
glist3.SelStart = 1
If UserFileName <> "" Then
Text1 = UserFileName
Else
Text1 = "NewFile"
End If
End If
     glist3.LeftMarginPixels = lastfactor * 5
  glist3.Backcolor = &H808080
  
glist3.EditFlag = True
glist3.NoCaretShow = False
glist3.Backcolor = &H0
glist3.ForeColor = &HFFFFFF
End If
glist3.ShowMe2
KeyCode = 0

ElseIf KeyCode = vbKeyReturn Then

DestroyCaret
If Text1 <> "" Then
glist3.EditFlag = False
glist3.Enabled = False

glist3_PanLeftRight True
End If
KeyCode = 0
End If
End If

End Sub

Private Sub glist3_LostFocus()

glist3.Backcolor = &H808080
glist3.ShowMe2
End Sub

Private Sub glist3_PanLeftRight(Direction As Boolean)
Dim that As New recDir, TT As Integer
If Text1 = "" Then Exit Sub

If Direction Then
If mySelector.Mydir.path = "" Then
If gList2.HeadLine = SelectFolderCaption And Text1 <> "" And Text1 <> ".." Then
ReturnFile = Text1 + "\"
Else
ReturnFile = ""
End If
mySelector.AbordAll
Unload Me
Else
If Text1 <> "" Then
Text1 = mySelector.Mydir.CleanName(Text1.text)

    If mySelector.Mydir.Nofiles Then
        If Text1 = SelectFolderButton Then
            ReturnFile = mySelector.mDoc1.TextParagraphOrder(0)
        ElseIf Text1.glistN.EditFlag Then
            ReturnFile = mySelector.mDoc1.TextParagraphOrder(0) + Text1 + "\"
        ElseIf mySelector.glistN.listindex >= 0 Then
            ReturnFile = Mid$(mySelector.Mydir.List(mySelector.glistN.listindex), 2) + "\"
        Else
         ReturnFile = mySelector.mDoc1.TextParagraphOrder(0) + Text1 + "\"
    End If
    Else

        ReturnFile = mySelector.mDoc1.TextParagraphOrder(0) + glist3.List(0)
        
    End If

mySelector.AbordAll
Unload Me
Else
Beep
End If
End If
End If
End Sub

Private Sub glist3_Selected2(Item As Long)
If Item = -2 Then
If glist3.PanPos <> 0 Then
glist3_PanLeftRight (True)
Exit Sub
End If

 glist3.LeftMarginPixels = lastfactor * 5
glist3.Backcolor = &H808080
glist3.ForeColor = &HE0E0E0
glist3.EditFlag = False
glist3.NoCaretShow = True


ElseIf Not mySelector.Mydir.isReadOnly(mySelector.Mydir.path) Then
If NewFolder Then
If Not (gList1.listindex = -1) Then
gList1.listindex = -1
gList1.ShowMe2
Text1 = "NewFolder"
End If
     glist3.LeftMarginPixels = lastfactor * 5
  glist3.Backcolor = &H808080
  
glist3.EditFlag = True
glist3.NoCaretShow = False
glist3.Backcolor = &H0
glist3.ForeColor = &HFFFFFF
ElseIf Not FileExist Then
If Not (gList1.listindex = -1) Then
gList1.listindex = -1
gList1.ShowMe2
If UserFileName <> "" Then
Text1 = UserFileName
Else
Text1 = "NewFile"
End If
End If
    glist3.LeftMarginPixels = lastfactor * 5
  glist3.Backcolor = &H808080
  
glist3.EditFlag = True
glist3.NoCaretShow = False
glist3.Backcolor = &H0
glist3.ForeColor = &HFFFFFF
End If
End If
glist3.ShowMe2
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'??
End Sub

Private Sub mySelector_DoubleClick(file As String)
ReturnFile = file
mySelector.AbordAll
Unload Me
End Sub

Private Sub mySelector_NewHeadline(newpath As String)
If firstpath = 0 Then
Else
If Not SaveDialog Then Text1 = ""
Set Image1 = LoadPicture("")
End If
firstpath = firstpath + 1
End Sub

Private Sub mySelector_TraceFile(file As String)
If Not DialogPreview Then
Text1 = mySelector.Mydir.List(mySelector.glistN.listindex)
refresh
Else
Dim aImage As StdPicture, sc As Single
Static ihave As Boolean
If ihave Then Exit Sub
mySelector.glistN.Enabled = False
' read ratio
Set Image1 = LoadPicture("")
On Error Resume Next
Err.clear
If FileLen(file) > 1500000 Then Image1.refresh
Set aImage = LoadPicture(file)
If file = "" Or Err.Number > 0 Then Exit Sub
ihave = True
If (aImage.Width / iwidth) < (aImage.Height / iheight) Then
sc = aImage.Height / iheight
Image1.Move iLeft + (iwidth - aImage.Width / sc) / 2, iTop, aImage.Width / sc, iheight
Else
sc = aImage.Width / iwidth
Image1.Move iLeft, iTop + (iheight - aImage.Height / sc) / 2, iwidth, aImage.Height / sc
End If

Image1.Picture = aImage

Text1 = mySelector.Mydir.List(mySelector.glistN.listindex)
mySelector.glistN.Enabled = True

ihave = False
End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Beep
End If
End Sub
Public Sub FillThereMyVersion2(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT, b As Long
b = CLng(Rnd * 3) + setupXY / 3

CopyFromLParamToRect a, thatRect
a.Left = a.Right - setupXY
a.Top = b
a.Bottom = b + setupXY / 5
mySelector.FillThere thathDC, VarPtr(a), thatbgcolor
a.Top = b + setupXY / 5 + setupXY / 10
a.Bottom = b + setupXY \ 2
mySelector.FillThere thathDC, VarPtr(a), thatbgcolor
'a.Top = b + setupXY * 4 / 5
'a.Bottom = b + setupXY - setupXY / 5
'mySelector.FillThere thathDC, VarPtr(a), thatbgcolor

End Sub
Public Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT, b As Long
b = 2
CopyFromLParamToRect a, thatRect
a.Left = b
a.Right = setupXY - b
a.Top = b
a.Bottom = setupXY - b
mySelector.FillThere thathDC, VarPtr(a), 0
b = 5
a.Left = b
a.Right = setupXY - b
a.Top = b
a.Bottom = setupXY - b
mySelector.FillThere thathDC, VarPtr(a), RGB(255, 160, 0)


End Sub
Sub PlaceSettings()
' using global var settings
Dim a() As String, i As Long, j As Long
a() = Split(Settings, ",")
For i = 0 To gList1.listcount - 1
gList1.ListSelectedNoRadioCare(i) = False
Next i
For i = LBound(a()) To UBound(a())
If gList1.GetMenuId(a(i), j) Then
gList1.ListSelectedNoRadioCare(j) = True
End If
Next i
End Sub
Function ReadSettings() As Boolean
' using global var settings
' we have to read at NostateDir=true
Dim a() As String, i As Long, j As Long
a() = Split(Settings, ",")
' reset some flags
gList1.StickBar = False
mySelector.mselChecked = False
multifileselection = False
ExpandWidth = False
For i = LBound(a()) To UBound(a())
While gList1.Id(j) <> a(i)
j = j + 1
Wend
j = j + 1  ' now we are in base 1
Select Case j
Case 2, 3, 4
If Not (mySelector.Mydir.SortType = j - 2) Then
mySelector.SortType = j - 2
ReadSettings = True
End If

mySelector.SortType = j - 2
Case 7  ' normal
If (mySelector.recnowchecked Or mySelector.recnow3checked) Then
ReadSettings = True
End If
mySelector.recnowchecked = False
mySelector.recnow3checked = False
Case 8  ' recursive plus level = 3
If Not (mySelector.recnowchecked And mySelector.recnow3checked) Then
ReadSettings = True
End If
mySelector.recnowchecked = True
mySelector.recnow3checked = True
Case 9  ' recursive
If Not (mySelector.recnowchecked And Not mySelector.recnow3checked) Then
ReadSettings = True
End If
mySelector.recnowchecked = True
mySelector.recnow3checked = False
Case 12 ' stickbar
gList1.StickBar = True ' plays for the two lists
Case 13 ' multiselect
mySelector.mselChecked = True
multifileselection = True
Case 14 ' Expand Width
ExpandWidth = True
End Select


Next i
End Function

Sub GetSettings()
Dim s As String, i As Long
For i = 0 To gList1.listcount - 1
If gList1.ListSelected(i) = True Then
If s = "" Then s = gList1.Id(i) Else s = s + "," + gList1.Id(i)
End If
Next i
Settings = s
End Sub
Function ScaleDialogFix(ByVal factor As Single) As Single
gList2.FontSize = 14.25 * factor
factor = gList2.FontSize / 14.25
gList1.FontSize = 11.25 * factor
factor = gList1.FontSize / 11.25
ScaleDialogFix = factor
End Function
Sub ScaleDialog(ByVal factor As Single, PreviewFile As Boolean, Optional NewWidth As Long = -1)
lastfactor = factor
gList1.AddPixels = 10 * factor
glist3.FontSize = 11.25 * factor
 glist3.LeftMarginPixels = factor * 5
mySelector.PreserveNpixelsHeaderRight = 20 * factor
setupXY = 20 * factor
oldLeftMarginPixels = 30 * factor
If mySelector.NostateDir Then
gList1.LeftMarginPixels = oldLeftMarginPixels
End If


bordertop = 10 * scrTwips * factor
borderleft = bordertop
Dim heightTop As Long, heightSelector As Long, HeightPreview As Long, HeightBottom As Long
Dim shapeHeight As Long
heightTop = 30 * factor * scrTwips
HeightBottom = 30 * factor * scrTwips
' some space here
If Not PreviewFile Then
heightSelector = 450 * factor * scrTwips

Else
heightSelector = 240 * factor * scrTwips
End If
HeightPreview = 180 * factor * scrTwips
shapeHeight = 160 * factor * scrTwips  ' and width
' some space here
HeightBottom = 30 * factor * scrTwips
If (NewWidth < 0) Or NewWidth <= (246 * scrTwips * factor) Then
NewWidth = 246 * scrTwips * factor
End If
itemWidth = (NewWidth - 2 * borderleft)
allwidth = NewWidth 'itemWidth + 2 * borderleft
Dim allheight As Long
gList2.FloatLimitTop = Screen.Height - bordertop - heightTop
gList2.FloatLimitLeft = Screen.Width - borderleft * 3
If PreviewFile Then
allheight = bordertop + heightTop + bordertop + heightSelector + bordertop + HeightPreview + bordertop + HeightBottom + bordertop
Else
allheight = bordertop + heightTop + bordertop + heightSelector + bordertop + HeightBottom + bordertop

End If

Move Left, Top, allwidth, allheight
gList2.Move borderleft, bordertop, itemWidth, heightTop
gList1.Move borderleft, 2 * bordertop + heightTop, itemWidth, heightSelector
glist3.Move borderleft, allheight - HeightBottom - bordertop, itemWidth, HeightBottom

If iwidth = 0 Then iwidth = itemWidth
If iheight = 0 Then iheight = HeightPreview
If PreviewFile Then
Dim curIwidth As Long, curIheight As Long, sc As Single
curIwidth = Image1.Width
curIheight = Image1.Height
iLeft = borderleft
iTop = 3 * bordertop + heightTop + heightSelector
iwidth = itemWidth
iheight = HeightPreview
If (curIwidth / iwidth) < (curIheight / iheight) Then
sc = curIheight / iheight
Image1.Move iLeft + (iwidth - curIwidth / sc) / 2, iTop, curIwidth / sc, iheight
Else
sc = curIwidth / iwidth
Image1.Move iLeft, iTop + (iheight - curIheight / sc) / 2, iwidth, curIheight / sc
End If
Shape1.Move borderleft, 3 * bordertop + heightTop + 240 * factor * scrTwips + 10 * scrTwips, itemWidth, shapeHeight
Image1.visible = True
Shape1.visible = True
Else
Image1.visible = False
Shape1.visible = False
End If
End Sub
