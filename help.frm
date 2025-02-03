VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   4650
   ClientLeft      =   11925
   ClientTop       =   -6825
   ClientWidth     =   7080
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "help.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MousePointer    =   5  'Size
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin M2000.gList glist1 
      Height          =   3825
      Left            =   330
      TabIndex        =   0
      Top             =   300
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6747
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      Backcolor       =   -2147483624
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents label1 As TextViewer, previewKey As Boolean
Attribute label1.VB_VarHelpID = -1
Private L As Long
Private t As Long
Private mt As Integer
Private back$
Private jump As Boolean

 Dim setupxy As Single
 Dim scrTwips As Long

Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Dim lX As Long, lY As Long, dr As Boolean, drmove As Boolean
Dim bordertop As Long, borderleft As Long
Dim allheight As Long, allwidth As Long, itemWidth As Long
Dim UAddPixelsTop As Long, flagmarkout As Boolean

Private Sub Form_Activate()
If HOOKTEST <> 0 Then UnHook HOOKTEST
End Sub

Private Sub Form_Deactivate()
jump = False
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyF12 And ((Not mHelp) Or trace) Then
showmodules
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo there1

If Form1.Visible Then
If Not gList1.EditFlag Then
Form1.SetFocus
INK$ = StrConv(ChrW$(KeyAscii Mod 256), 64, Form1.GetLCIDFromKeyboard)
End If
End If

there1:
End Sub

Private Sub Form_Load()
Form4Loaded = True
Set LastGlist2 = Nothing
setupxy = 20 * Helplastfactor
scrTwips = Screen.TwipsPerPixelX
gList1.CapColor = rgb(255, 160, 0)
gList1.LeftMarginPixels = 4
Set label1 = New TextViewer
Set label1.Container = gList1
label1.NoCenterLineEdit = True
label1.FileName = vbNullString
label1.glistN.NoMoveDrag = True
label1.glistN.DropEnabled = False
label1.glistN.DragEnabled = Not abt
label1.NoMark = True
label1.NoColor = True
label1.EditDoc = False
label1.nowrap = False
label1.enabled = False    '' true before
label1.glistN.FloatList = True
label1.glistN.MoveParent = True
With label1.glistN
If FeedbackExec$ = vbNullString Or Not abt Then
.WordCharRightButIncluded = ChrW(160) + "("
.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", "@", Chr$(9), "#", "%", "&", "$")
.WordCharRight = ConCat(":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", Chr$(9), "#")
'
.WordCharLeftButIncluded = "#$@~"
Else
.WordCharLeft = "['"
.WordCharRight = "']"
.WordCharRightButIncluded = ChrW(160)
.WordCharLeftButIncluded = vbNullString
End If

End With

mt = DXP
If Helplastfactor = 0 Then Helplastfactor = 1
Helplastfactor = ScaleDialogFix(helpSizeDialog)
HelpLastWidth = -1
If ExpandWidth Then
If HelpLastWidth = 0 Then HelpLastWidth = -1
Else
HelpLastWidth = -1
End If

End Sub
Public Sub moveMe()
ScaleDialog Helplastfactor, HelpLastWidth
Hook2 hWnd, gList1
label1.glistN.SoftEnterFocus
If IsWine Then
If Not Screen.ActiveForm Is Nothing Then
If Not Screen.ActiveForm Is Form4 Then
Form4.Show , Screen.ActiveForm
End If
End If
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    
    If Helplastfactor = 0 Then Helplastfactor = 1

    If bordertop < 150 Then
    If (Y > Height - 150 And Y < Height) And (X > Width - 150 And X < Width) Then
    dr = True
    MousePointer = vbSizeNWSE
    lX = X
    lY = Y
    End If
    
    Else
    If (X > Width - borderleft And X < Width) Or (Y > Height - bordertop) Then  ' (y > Height - bordertop And y < Height) And
    dr = True
    MousePointer = vbSizeNWSE
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
If (Y > Height - 150 And Y < Height) And (X > Width - 150 And X < Width) Then MousePointer = vbSizeNWSE Else If Not (dr Or drmove) Then MousePointer = 0
 Else
 If (Y > Height - bordertop And Y < Height) And (X > Width - borderleft And X < Width) Then MousePointer = vbSizeNWSE Else If Not (dr Or drmove) Then MousePointer = 0
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
        If Helplastfactor = 0 Then Helplastfactor = 1
        factor = Helplastfactor

        
  
        once = True
        If Height > VirtualScreenHeight() Then addy = -(Height - VirtualScreenHeight()) + addy
        If Width > VirtualScreenWidth() Then addX = -(Width - VirtualScreenWidth()) + addX
        If (addy + Height) / vH_y > 0.4 And ((Width + addX) / vH_x) > 0.4 Then
   
        If addy <> 0 Then helpSizeDialog = ((addy + Height) / vH_y)
        Helplastfactor = ScaleDialogFix(helpSizeDialog)


        If ((Width * Helplastfactor / factor + addX) / Height * Helplastfactor / factor) < (vH_x / vH_y) Then
        addX = -Width * Helplastfactor / factor - 1
      
           End If

        If addX = 0 Then
        
        If Helplastfactor <> factor Then ScaleDialog Helplastfactor, Width

        lX = X
        
        Else
        lX = X * Helplastfactor / factor
             ScaleDialog Helplastfactor, (Width + addX) * Helplastfactor / factor
         
   
         End If

        
        HelpLastWidth = Width


''gList1.PrepareToShow
        lY = lY * Helplastfactor / factor
        End If
        Else
        lX = X
        lY = Y
   
End If
once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)

If dr Then Me.MousePointer = 0
dr = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UnHook2 hWnd
Set LastGlist2 = Nothing
End Sub

Private Sub Form_Terminate()
''''Set HelpStack.Owner = Nothing
End Sub

Private Sub ffhelp(A$)
If A$ = "������" Then A$ = "���"
If A$ = "GENERAL" Then A$ = "ALL"
If Left$(A$, 1) = "#" Then
If Mid$(A$, 2) < "�" Then
fHelp Basestack1, A$, True
Else
fHelp Basestack1, A$
End If

Else

If Left$(A$, 1) < "�" Then
fHelp Basestack1, A$, True
Else
fHelp Basestack1, A$
End If
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
label1.Dereference  ' to ensure that no reference hold objects..
Set label1 = Nothing
Helplastfactor = 1
helpSizeDialog = 1
Form4Loaded = False
End Sub

Private Sub glist1_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal X As Long, ByVal Y As Long)
If item = -1 Then
If gList1.DoubleClickCheck(Button, item, X, Y, gList1.WidthPixels - 10 * Helplastfactor, 10 * Helplastfactor, 8 * Helplastfactor, -1) Then
            HelpLastWidth = -1
            Unload Me
End If
Else
gList1.MousePointer = 1
End If
End Sub


Private Sub glist1_getpair(A As String, b As String)
If mHelp Or abt Then
gList1.EditFlag = False
    MKEY$ = MKEY$ & A
    A = vbNullString
End If
End Sub

Private Sub gList1_KeyDown(KeyCode As Integer, shift As Integer)
If shift <> 0 Then
If label1.SelectionColor = rgb(255, 64, 128) Then label1.SelectionColor = 0
label1.NoMark = False
label1.EditDoc = True
End If
Select Case KeyCode
Case vbKeyDelete, vbKeyBack, vbKeyReturn, vbKeySpace

gList1.EditFlag = False
If mHelp Or abt Then
MKEY$ = MKEY$ & ChrW$(KeyCode)
KeyCode = 0
End If
End Select
If mHelp Or abt Then shift = 0

End Sub



Private Sub glist1_MarkOut()
If flagmarkout Then
If label1.SelectionColor = rgb(255, 64, 128) Then label1.SelectionColor = 0
flagmarkout = False: Exit Sub
End If
End Sub

Private Sub gList1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
flagmarkout = True
If mHelp Then
shift = 0
End If
End Sub

Private Sub gList1_MouseUp(X As Single, Y As Single)
If gList1.DoubleClickArea(X, Y, gList1.WidthPixels - 10 * Helplastfactor, 10 * Helplastfactor, 8 * Helplastfactor) Then
            HelpLastWidth = -1
            Unload Me
            End If
End Sub

Private Sub gList1_selected2(item As Long)

label1.NoMark = False
label1.EditDoc = True
End Sub

Private Sub glist1_WordMarked(ThisWord As String)
If abt Then
feedback$ = Trim$(Replace(ThisWord, ChrW(160), " "))
feednow$ = FeedbackExec$
label1.SelLengthSilent = 0
CallGlobal feednow$
Else
If Not mHelp Then
If Form2.Visible Then
    If ThisWord = "Control" Or ThisWord = "�������" Then
    sHelp Form2.gList2.HeadLine, Form2.testpad.Text, vH_x, vH_y
    vHelp
    If TestShowCode Then
    label1.SelStartSilent = Form2.testpad.SelStart
    label1.SelLengthSilent = 0
    label1.SelectionColor = rgb(255, 64, 128)
    If Form2.testpad.SelStart > 0 And Form2.testpad.SelLength > 0 Then label1.SelLength = Form2.testpad.SelLength
    End If
    Else
    ffhelp Trim$(Replace(ThisWord, ChrW(160), " "))
    End If
    Else
    label1.SelLengthSilent = 0
    label1.SelectionColor = 0
ffhelp Trim$(Replace(ThisWord, ChrW(160), " "))
End If

End If

End If
ThisWord = vbNullString

End Sub
Public Sub FillThereMyVersion2(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT, b As Long
b = setupxy / 3

CopyFromLParamToRect A, thatRect
A.Right = A.Right - b
A.Left = A.Right - setupxy - b
A.top = b
A.Bottom = b + setupxy / 5
FillThere thathDC, VarPtr(A), thatbgcolor
A.top = b + setupxy / 5 + setupxy / 10
A.Bottom = b + setupxy \ 2
FillThere thathDC, VarPtr(A), thatbgcolor
A.top = b + 2 * (setupxy / 5 + setupxy / 10)
A.Bottom = A.top + setupxy / 5
FillThere thathDC, VarPtr(A), thatbgcolor

End Sub
Public Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT, b As Long
b = 2
CopyFromLParamToRect A, thatRect
A.Left = A.Right - b
A.Right = A.Right - setupxy + b
A.top = b
A.Bottom = setupxy - b
FillThere thathDC, VarPtr(A), gList1.dcolor
b = 5
A.Left = A.Left - 3
A.Right = A.Right + 3
A.top = b
A.Bottom = setupxy - b
FillThere thathDC, VarPtr(A), gList1.CapColor


End Sub
Public Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim A As RECT
CopyFromLParamToRect A, thatRect
FillBack thathDC, A, thatbgcolor
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub

Private Sub Label1_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = -1 Then

FillThereMyVersion thisHDC, thisrect, &HF0F0F0
''skip = True
End If
End Sub
Function ScaleDialogFix(ByVal factor As Single) As Single
gList1.FontSize = 14.25 * factor * dv15 / 15
factor = gList1.FontSize / 14.25 / dv15 * 15
ScaleDialogFix = factor
End Function
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)
Dim h As Long, I As Long
Helplastfactor = factor
setupxy = 20 * factor
bordertop = 10 * dv15 * factor
borderleft = bordertop
If (NewWidth < 0) Or NewWidth <= vH_x * factor Then
NewWidth = vH_x * factor
End If



allwidth = NewWidth  ''vH_x * factor
allheight = vH_y * factor
itemWidth = allwidth - 2 * borderleft
Dim kk As Long
If Left < MinMonitorLeft Or top < MinMonitorTop Then
kk = 0
Else
kk = 1
End If
myform Me, Left * kk, top * kk, allwidth, allheight, True, factor

  
gList1.addpixels = 4 * factor
label1.move borderleft, bordertop, itemWidth, allheight - bordertop * 2

label1.NewTitle vH_title$, (4 + UAddPixelsTop) * factor
label1.Render
gList1.FloatLimitTop = VirtualScreenHeight() - bordertop - bordertop * 3
gList1.FloatLimitLeft = VirtualScreenWidth() - borderleft * 3


End Sub
Public Sub hookme(this As gList)

''Set LastGlist2 = this

End Sub
Sub ByeBye()
Unload Me
End Sub
Private Sub gList1_RefreshDesktop()
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
End Sub
