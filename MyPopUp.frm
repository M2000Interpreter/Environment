VERSION 5.00
Begin VB.Form MyPopUp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin M2000.gList gList1 
      Height          =   5475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   9657
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
      Enabled         =   -1  'True
      dcolor          =   32896
      Backcolor       =   3881787
      ForeColor       =   14737632
      CapColor        =   9797738
   End
End
Attribute VB_Name = "MyPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gokeyboard As Boolean, lastitem As Long, part1 As String, lastgoodnum As Long
Private height1, width1
Dim lX As Long, lY As Long, dr As Boolean
Dim bordertop As Long, borderleft As Long, lastshift As Integer
Dim allheight As Long, allwidth As Long, itemWidth As Long
Private myobject As Object
Public LASTActiveForm As Form, previewKey As Boolean
Dim ttl$(1 To 2)
Public Sub Up(Optional X As Variant, Optional Y As Variant)
Dim hmonitor As Long
If IsMissing(X) Then
X = CSng(MOUSEX())
Y = CSng(MOUSEY())
Else
X = X + Form1.Left
Y = Y + Form1.top
End If
hmonitor = FindMonitorFromPixel(X, Y)
If X + Width > ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left Then
If Y + Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).top Then
move ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left - Width, Y - Height
Else
move ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left - Width, Y
End If
ElseIf Y + Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).top Then
move X, Y - Height
Else
move X, Y
End If
Show
MyDoEvents
End Sub
Public Sub UpGui(that As Object, X As Variant, Y As Variant, thistitle$)
Dim hmonitor As Long
If thistitle$ <> "" Then
gList1.HeadLine = vbNullString
gList1.HeadLine = thistitle$
gList1.HeadlineHeight = gList1.HeightPixels
Else
gList1.HeadLine = vbNullString
gList1.HeadlineHeight = 0
End If
X = X + that.Left
Y = Y + that.top
hmonitor = FindMonitorFromPixel(X, Y)
If X + Width > ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left Then
If Y + Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).top Then
move ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left - Width, Y - Height
Else
move ScrInfo(hmonitor).Width + ScrInfo(hmonitor).Left - Width, Y
End If
ElseIf Y + Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).top Then
move X, Y - Height
Else
move X, Y
End If
If thistilte$ <> "" Then

Else

End If
Show
MyDoEvents
End Sub



Public Sub feedlabels(that As Object, EditTextWord As Boolean)
Dim k As Long
Set myobject = that

With gList1
'.NoWheel = True
.restrictLines = 14
.FloatList = True
.MoveParent = True
.SingleLineSlide = True
.NoPanRight = True
.AutoHide = True
.FreeMouse = True
End With
height1 = 5475 * DYP / 15
width1 = 4155 * DXP / 15
If pagio$ = "GREEK" Then
With gList1
''
.StickBar = False
''.AddPixels = 4
.VerticalCenterText = True
If Typename(myobject) <> "GuiEditBox" Then
part1 = " " + GetStrUntil("(", (textinformCaption)) + "("
.additemFast textinformCaption
Else
part1 = " " + GetStrUntil("(", (myobject.textinform)) + "("
.additemFast myobject.textinform
End If
.AddSep
.additemFast "Αποκοπή Ctrl+X"
.menuEnabled(2) = that.Form1mn1Enabled
.additemFast "Αντιγραφή Ctrl+C"
.menuEnabled(3) = that.Form1mn2Enabled
.additemFast "Επικόλληση Ctrl+V"
.menuEnabled(4) = that.Form1mn3Enabled
If Typename(myobject) <> "GuiEditBox" Then
.AddSep
.additemFast "Έξοδος με αλλαγές (ESC)"
.AddSep
.additemFast "Έξοδος χωρίς αλλαγές shift F12"
Else
k = 4
End If
.AddSep
.additemFast "Αναζήτησε πάνω F2"
.menuEnabled(10 - k) = that.Form1supEnabled
.additemFast "Αναζήτησε κάτω F3"
.menuEnabled(11 - k) = that.Form1sdnEnabled
.additemFast "Κάνε το ίδιο παντού F4"
.menuEnabled(12 - k) = that.Form1mscatEnabled
.additemFast "Αλλαγή λέξης F5"
.menuEnabled(13 - k) = that.Form1rthisEnabled
.AddSep
.additemFast "Αναδίπλωση λέξεων F1"

.MenuItem 16 - k, True, False, Not that.nowrap, "warp"
.additemFast "Μεταφορά Κειμένου"
.MenuItem 17 - k, True, False, that.glistN.DragEnabled, "drag"
If k = 0 Then
.additemFast "Χρώμα/Σύμπτυξη Γλώσσας F11"
Else
.additemFast "Χρώμα F11"
End If
.MenuItem 18 - k, True, False, shortlang, "short"
.additemFast "Εμφάνιση Παραγράφων F10"
.MenuItem 19 - k, True, False, that.showparagraph, "para"
.additemFast "Μέτρηση λέξεων F9"
.AddSep
.additemFast "Βοήθεια ctrl+F1"
If Not EditTextWord Then
If k = 0 Then
.HeadLine = "Μ2000 Συντάκτης"
.AddSep
.additemFast "Τμήματα/Συναρτήσεις F12"
.menuEnabled(23 - k) = SubsExist()
ttl$(1) = "Εισαγωγή Αρχείου"

.additemFast ttl$(1)
.menuEnabled(24 - k) = True
ttl$(2) = "Εισαγωγή Πόρου"
.additemFast ttl$(2)
.menuEnabled(25 - k) = True

End If
End If
End With
Else
With gList1
''gList1.HeadLine = "Μ2000"
.StickBar = False
''''.AddPixels = 4
.VerticalCenterText = True
If Typename(myobject) <> "GuiEditBox" Then
part1 = " " + GetStrUntil("(", (textinformCaption)) + "("
.additemFast textinformCaption
Else
part1 = " " + GetStrUntil("(", (myobject.textinform)) + "("
.additemFast myobject.textinform
End If
.AddSep
.additemFast "Cut   Ctrl+X"
.menuEnabled(2) = that.Form1mn1Enabled
.additemFast "Copy  Ctrl+C"
.menuEnabled(3) = that.Form1mn2Enabled
.additemFast "Paste Ctrl+V"
.menuEnabled(4) = that.Form1mn3Enabled
.AddSep
If Typename(myobject) <> "GuiEditBox" Then
.additemFast "Save and Exit (ESC)"
.AddSep
.additemFast "Discard Changes shift F12"
.AddSep
Else
k = 4
End If
.additemFast "Search up F2"
.menuEnabled(10 - k) = that.Form1supEnabled
.additemFast "Search down F3"
.menuEnabled(11 - k) = that.Form1sdnEnabled
.additemFast "Make same all F4"
.menuEnabled(12 - k) = that.Form1mscatEnabled
.additemFast "Replace word F5"
.menuEnabled(13 - k) = that.Form1rthisEnabled
.AddSep
.additemFast "Word Wrap F1"
.MenuItem 16 - k, True, False, Not that.nowrap, "warp"
.additemFast "Drag Enabled"
.MenuItem 17 - k, True, False, that.glistN.DragEnabled, "drag"
If k = 0 Then
.additemFast "Color/Short Language F11"
Else
.additemFast "Color F11"
End If
.MenuItem 18 - k, True, False, shortlang, "short"
.additemFast "Paragraph Mark F10"
.MenuItem 19 - k, True, False, that.showparagraph, "para"
.additemFast "Word count F9"
.AddSep
.additemFast "Help ctrl+F1"
If Not EditTextWord Then
If k = 0 Then
.HeadLine = "Μ2000 Editor"
.AddSep
.additemFast "Modules/Functions F12"
.menuEnabled(23 - k) = SubsExist()
ttl$(1) = "Insert File"
.additemFast ttl$(1)
.menuEnabled(24 - k) = True
ttl$(2) = "Load Resource"
.additemFast ttl$(2)

.menuEnabled(25 - k) = True

End If
End If

End With
End If
If Pouplastfactor = 0 Then Pouplastfactor = 1
 Pouplastfactor = ScaleDialogFix(helpSizeDialog)
If ExpandWidth And False Then
If PopUpLastWidth = 0 Then PopUpLastWidth = -1
Else
PopUpLastWidth = -1
End If
If ExpandWidth Then
If PopUpLastWidth = 0 Then PopUpLastWidth = -1
Else
PopUpLastWidth = -1
End If
ScaleDialog Pouplastfactor, PopUpLastWidth
gList1.ListIndex = 0
gList1.ShowBar = True
gList1.ShowBar = False
gList1.NoPanLeft = False
gList1.SoftEnterFocus

End Sub
Public Sub UNhookMe()
' nothing
End Sub

Private Sub Form_Activate()
If HOOKTEST <> 0 Then UnHook HOOKTEST
End Sub

Private Sub Form_Load()
Set LASTActiveForm = Screen.ActiveForm
End Sub

Private Sub Form_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    
    If Pouplastfactor = 0 Then Pouplastfactor = 1

    If bordertop < 150 Then
    If (Y > Height - 150 And Y < Height) And (X > Width - 150 And X < Width) Then
    dr = True
    mousepointer = vbSizeNWSE
    lX = X
    lY = Y
    End If
    
    Else
    If (Y > Height - bordertop And Y < Height) And (X > Width - borderleft And X < Width) Then
    dr = True
    mousepointer = vbSizeNWSE
    lX = X
    lY = Y
    End If
    End If

End If
End Sub
Private Sub Form_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
Dim addX As Long, addy As Long, factor As Single, once As Boolean
If once Then Exit Sub
If Button = 0 Then dr = False: drmove = False
If bordertop < 150 Then
If (Y > Height - 150 And Y < Height) And (X > Width - 150 And X < Width) Then mousepointer = vbSizeNWSE Else If Not (dr Or drmove) Then mousepointer = 0
 Else
 If (Y > Height - bordertop And Y < Height) And (X > Width - borderleft And X < Width) Then mousepointer = vbSizeNWSE Else If Not (dr Or drmove) Then mousepointer = 0
End If
If dr Then



If bordertop < 150 Then

        If Y < (Height - 150) Or Y > Height Then addy = (Y - lY)
     If X < (Width - 150) Or X > Width Then addX = (X - lX)
     
Else
    If Y < (Height - bordertop) Or Y > Height Then addy = (Y - lY)
        If X < (Width - borderleft) Or X > Width Then addX = (X - lX)
    End If
    

    
  '' If Not ExpandWidth Then addX = 0
        If Pouplastfactor = 0 Then Pouplastfactor = 1
        factor = Pouplastfactor

        
  
        once = True
        If Height > VirtualScreenHeight() Then addy = -(Height - VirtualScreenHeight()) + addy
        If Width > VirtualScreenWidth() Then addX = -(Width - VirtualScreenWidth()) + addX
        If (addy + Height) / height1 > 0.4 And ((Width + addX) / width1) > 0.4 Then
   
        If addy <> 0 Then helpSizeDialog = ((addy + Height) / height1)
        Pouplastfactor = ScaleDialogFix(helpSizeDialog)


        If ((Width * Pouplastfactor / factor + addX) / Height * Pouplastfactor / factor) < (width1 / height1) Then
        addX = -Width * Pouplastfactor / factor - 1
      
           End If

        If addX = 0 Then
        
        If Pouplastfactor <> factor Then ScaleDialog Pouplastfactor, Width

        lX = X
        
        Else
        lX = X * Pouplastfactor / factor
             ScaleDialog Pouplastfactor, (Width + addX) * Pouplastfactor / factor
         
   
         End If

        
        PopUpLastWidth = Width


''gList1.PrepareToShow
        lY = lY * Pouplastfactor / factor
        End If
        Else
        lX = X
        lY = Y
   
End If
once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)

If dr Then Me.mousepointer = 0
dr = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set LASTActiveForm = Nothing
Set myobject = Nothing
End Sub

Private Sub gList1_ChangeListItem(item As Long, content As String)
Dim content1 As Long
If item = 0 Then
content1 = Int(val("0" & Trim$(Mid$(content, Len(part1) + 1))))

        If content1 > myobject.mDoc.DocLines Or content1 < 0 Then
        content = gList1.list(item)
              gList1.SelStart = Len(gList1.list(item)) - 1
        Else
        lastgoodnum = content1
        If content1 = 0 Then
        content = part1 & ")"
        gList1.SelStart = 3
        Else
        content = part1 & CStr(content1) & ")"
        End If
        
        End If
End If
End Sub



Private Sub glist1_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal X As Long, ByVal Y As Long)
''If X * dv15 > Width / 2 Then

If item = -1 Then

Else
gList1.mousepointer = 1
If gokeyboard Then Exit Sub
If gList1.ListSep(item) Then Exit Sub
gList1.EditFlag = False
''''''''''''''''''''''''''''''
If lastitem = item Then Exit Sub
If gList1.ListSep(item) Then Exit Sub
gList1.ListindexPrivateUse = item
gList1.ShowMe2

lastitem = item
End If
End Sub

Private Sub gList1_KeyDown(keycode As Integer, shift As Integer)
gokeyboard = True
If keycode = vbKeyEscape Then Unload Me: Exit Sub

If gList1.ListIndex = -1 Then gList1.ListindexPrivateUse = lastitem

If ((keycode >= vbKey0 And keycode <= vbKey9) Or (keycode >= vbKeyNumpad0 And keycode <= vbKeyNumpad9)) And gList1.EditFlag = False And gList1.ListIndex = 0 Then
                        lastitem = 0
                    gList1.PromptLineIdent = Len(part1)
                    gList1.list(0) = vbNullString
                    gList1.SelStart = 3
                    gList1.EditFlag = True

ElseIf gList1.ListIndex = 0 And gList1.EditFlag = True Then
        If keycode = vbKeyDown Or keycode = vbKeyReturn And gList1.EditFlag = True Then
        gList1.EditFlag = False
        lastitem = 0
        keycode = 0
        DoCommand 1
        gList1.ListindexPrivateUse = 0
        gList1.ShowMe2
        
        lastitem = 0
        gList1.ListindexPrivateUse = -1
        End If
End If



End Sub



Private Sub gList1_LostFocus()
Unload Me
End Sub

Private Sub gList1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)

gokeyboard = False
gList1.PromptLineIdent = 0
If lastitem = item Then Exit Sub
gList1.ListindexPrivateUse = -1
End Sub

Private Sub glist1_RegisterGlist(this As gList)
this.NoWheel = True '
End Sub

Private Sub gList1_ScrollSelected(item As Long, Y As Long)
If gokeyboard Then Exit Sub

gList1.EditFlag = False
''''''''''''''''''''''
If lastitem = item - 1 Then Exit Sub
gList1.ListindexPrivateUse = item - 1
gList1.ShowMe2

lastitem = item
gList1.ListindexPrivateUse = -1
End Sub

Private Sub gList1_selected(item As Long)
If Not gokeyboard Then
DoCommand item
Else
If Not gList1.EditFlag Then
gList1.ListindexPrivateUse = item - 1
lastitem = gList1.ListIndex
gList1.ShowMe2


End If
End If
End Sub

Private Sub gList1_selected2(item As Long)
 DoCommand item + 1

End Sub
Private Sub DoCommand(item As Long)
Dim k As Long, l As Long
Dim b As basetask, fname$, files() As String, reader As Document, neo$, s$, noinp As Double
If Typename(myobject) = "GuiEditBox" Then k = 4: l = 100
Select Case item - 1
Case -2
Exit Sub
Case 0
If lastgoodnum > 0 Then
With gList1
.menuEnabled(2) = False
.menuEnabled(3) = False
myobject.SetRowColumn lastgoodnum, 0
.PromptLineIdent = 0
lastitem = 0
.ListindexPrivateUse = -1
myobject.glistN.ShowMe


Exit Sub
End With
End If
Case 2
If k = 0 Then
    Form1.mn1sub
Else
    myobject.mn1sub
End If
Case 3
If k = 0 Then
    Form1.mn2sub
Else
    myobject.mn2sub
End If
Case 4
If k = 0 Then
    Form1.mn3sub
Else
    myobject.mn3sub
End If
Case 6 - l
    Form1.mn4sub
Case 8 - l
    Form1.mn5sub
Case 10 - k
If k = 0 Then
    Form1.supsub
Else
    myobject.supsub
End If
Case 11 - k
If k = 0 Then
    Form1.sdnSub
Else
    myobject.sdnSub
End If
Case 12 - k
If k = 0 Then
    Form1.mscatsub
Else
    myobject.mscatsub
End If
Case 13 - k
If k = 0 Then
    Form1.rthissub
Else
    myobject.rthissub
End If
Case 15 - k
gList1.ListSelectedNoRadioCare(17 - k) = Not gList1.ListChecked(17 - k)
If k = 0 Then
Form1.wordwrapsub
Else
myobject.wordwrapsub
End If
Case 16 - k
gList1.ListSelectedNoRadioCare(18 - k) = Not gList1.ListChecked(18 - k)
myobject.glistN.DragEnabled = Not myobject.glistN.DragEnabled
Case 21 - k
If k = 0 Then
    Form1.helpmeSub
Else
    myobject.helpmeSub
End If
Case 23 - l
showmodules
Case 24 - l
Set b = New basetask
Set b.Owner = LASTActiveForm
If b.Owner Is Nothing Then Set b.Owner = Form1
Form1.TEXT1.glistN.enabled = False
fname$ = GetFile(b, ttl$(1), mcd, "TXT|GSB|BCK|GM2", True)
Set reader = New Document
If fname$ <> "" Then
If ReturnListOfFiles <> "" Then
    files() = Split(ReturnListOfFiles, "#")
    For k = 0 To UBound(files())
    Form1.TEXT1.PasteText "\\$ " + files(k)
    Set reader = New Document
    reader.ReadUnicodeOrANSI files(k), True
    Form1.TEXT1.PasteText reader.textDoc
    Next k
Else
    Form1.TEXT1.PasteText "\\$ " + fname$
    reader.ReadUnicodeOrANSI fname$, True, , True
    Form1.TEXT1.PasteText reader.textDoc
End If
Form1.TEXT1.Render
End If
Form1.TEXT1.glistN.enabled = True
Form1.TEXT1.ManualInform
Unload Me
Case 26 - 1
    Set b = New basetask
    Set b.Owner = LASTActiveForm
    If b.Owner Is Nothing Then Set b.Owner = Form1
Set reader = New Document
Form1.TEXT1.glistN.enabled = False
fname$ = GetFile(b, ttl$(2), mcd, "", True)

If fname$ <> "" Then
If pagio$ = "GREEK" Then
s$ = "Πόρος"
neo$ = Trim$(InputBoxN("Όνομα Μεταβλητής (αριθμητική ή αλφαριθμητική)", "Συγγραφή Κειμένου", s$, noinp))
If noinp <> 1 Then
Form1.TEXT1.glistN.enabled = True
Exit Sub
End If
If MyTrim(neo$) = vbNullString Then neo$ = "Πόρος"
Else
s$ = "Resource"
neo$ = Trim$(InputBoxN("Variable Name (numeric or string)", "Text Editor", s$, noinp))
If noinp <> 1 Then
Form1.TEXT1.glistN.enabled = True
Exit Sub
End If
If MyTrim(neo$) = vbNullString Then neo$ = "Resource"
End If
If Right$(neo$, 1) = "$" Then
    s$ = "$"
    neo$ = Left$(neo$, Len(neo$) - 1)
ElseIf Right$(neo$, 1) = ")" Then
    s$ = ")"
    neo$ = Left$(neo$, Len(neo$) - 1)
Else
    s$ = vbNullString
End If
    If ReturnListOfFiles <> "" Then
        files() = Split(ReturnListOfFiles, "#")
        For k = 0 To UBound(files())
        Form1.TEXT1.PasteText "\\$ " + files(k)
        If pagio$ = "GREEK" Then
            Form1.TEXT1.PasteText "Δυαδικό {"
            Form1.TEXT1.PasteText FileToEncode64(files(k), 6)
            If s$ = ")" Then
            Form1.TEXT1.PasteText "} Ως " + neo$ + CStr(k) + s$
            Else
            Form1.TEXT1.PasteText "} Ως " + neo$ + CStr(k + 1) + s$
            End If
        Else
            Form1.TEXT1.PasteText "Binary {"
            Form1.TEXT1.PasteText FileToEncode64(files(k), 6)
            If s$ = ")" Then
            Form1.TEXT1.PasteText "} As " + neo$ + CStr(k) + s$
            Else
            Form1.TEXT1.PasteText "} As " + neo$ + CStr(k + 1) + s$
            End If
        End If
        
        Next k
    Else
        Form1.TEXT1.PasteText "\\$ " + fname$
        If pagio$ = "GREEK" Then
        Form1.TEXT1.PasteText "Δυαδικό {"
        Form1.TEXT1.PasteText FileToEncode64(fname$, 6)
        If s$ = ")" Then
            If Right$(neo$, 1) <> "(" Then
                Form1.TEXT1.PasteText "} Ως " + neo$ + s$
            Else
                Form1.TEXT1.PasteText "} Ως " + neo$ + "0" + s$
            End If
        Else
        Form1.TEXT1.PasteText "} Ως " + neo$ + s$
        End If
        Else
        Form1.TEXT1.PasteText "Binary {"
        Form1.TEXT1.PasteText FileToEncode64(fname$, 6)
        If s$ = ")" Then
            If Right$(neo$, 1) <> "(" Then
                Form1.TEXT1.PasteText "} As " + neo$ + s$
            Else
                Form1.TEXT1.PasteText "} As " + neo$ + "0" + s$
            End If
        Else
        Form1.TEXT1.PasteText "} As " + neo$ + s$
        End If
        End If

    End If
Form1.TEXT1.Render

End If
Form1.TEXT1.glistN.enabled = True
Form1.TEXT1.ManualInform

Unload Me
Case 17 - k
If k = 0 Then
With myobject
shortlang = Not shortlang
.ManualInform
End With
Else
With myobject
shortlang = Not shortlang
.NoColor = shortlang
End With
End If
Case 18 - k
With myobject
.showparagraph = Not .showparagraph
.Render
End With
Case 19 - k
With myobject
If .glistN.lines > 1 Then
If UserCodePage = 1253 Then
.ReplaceTitle = "Λέξεις στο κείμενο:" + CStr(.mDoc.WordCount)
Else
.ReplaceTitle = "Words in text:" + CStr(.mDoc.WordCount)
End If
End If
End With

End Select
Unload Me
End Sub

Function ScaleDialogFix(ByVal factor As Single) As Single
gList1.FontSize = 11.25 * factor * dv15 / 15
factor = gList1.FontSize / 11.25 / dv15 * 15
ScaleDialogFix = factor
End Function
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)
Dim h As Long, i As Long
Pouplastfactor = factor
gList1.LeftMarginPixels = 30 * factor
setupxy = 20 * factor
bordertop = 10 * dv15 * factor
borderleft = bordertop
If (NewWidth < 0) Or NewWidth <= width1 * factor Then
NewWidth = width1 * factor
End If
allwidth = NewWidth  ''width1 * factor
allheight = height1 * factor
itemWidth = allwidth - 2 * borderleft
''MyForm Me, Left, top, allwidth, allheight, True, factor

move Left, top, allwidth, allheight
  
gList1.addpixels = 4 * factor

gList1.move borderleft, bordertop, itemWidth, allheight - bordertop * 2

gList1.CalcAndShowBar
gList1.ShowBar = False
gList1.FloatLimitTop = VirtualScreenHeight() - bordertop - bordertop * 3
gList1.FloatLimitLeft = VirtualScreenWidth() - borderleft * 3

End Sub

Public Sub hookme(this As gList)
'' do nothing
End Sub
Private Sub gList1_SpecialColor(RGBcolor As Long)
RGBcolor = rgb(100, 132, 254)
End Sub
Private Sub gList1_RefreshDesktop()
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
End Sub

Private Sub gList1_UnregisterGlist()
gList1.NoWheel = True
End Sub
