VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   5295
   ClientLeft      =   3000
   ClientTop       =   3000
   ClientWidth     =   7860
   Icon            =   "TEST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleMode       =   0  'User
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin M2000.gList gList4 
      Height          =   1920
      Left            =   4050
      TabIndex        =   5
      Top             =   810
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   3387
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
      BorderStyle     =   1
      ShowBar         =   0   'False
      Backcolor       =   3881787
      ForeColor       =   16777215
   End
   Begin M2000.gList gList3 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   1058
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
      Backcolor       =   3881787
      ForeColor       =   16777215
   End
   Begin M2000.gList gList0 
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   4620
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   979
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
      Backcolor       =   657930
      ForeColor       =   16777215
   End
   Begin M2000.gList gList1 
      Height          =   1800
      Left            =   90
      TabIndex        =   0
      Top             =   2775
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   3175
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
      Height          =   615
      Index           =   1
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1455
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   1085
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
      Height          =   615
      Index           =   2
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2115
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   1085
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
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   135
      Width           =   7635
      _ExtentX        =   13467
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
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents testpad As TextViewer
Attribute testpad.VB_VarHelpID = -1
Public WithEvents Compute As myTextBox
Attribute Compute.VB_VarHelpID = -1
Private Label(0 To 2) As New myTextBox
Dim run_in_basestack1 As Boolean
Dim MyBaseTask As New basetask
Dim setupxy As Single
Dim Lx As Long, ly As Long, dr As Boolean, drmove As Boolean
Dim prevx As Long, prevy As Long
Dim a$
Dim bordertop As Long, borderleft As Long
Dim allheight As Long, allwidth As Long, itemWidth As Long, itemwidth3 As Long, itemwidth2 As Long
Dim height1 As Long, width1 As Long
Dim doubleclick As Long

Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Dim EXECUTED As Boolean
Dim stolemodalid As Variant
Public Property Set Process(mBtask As basetask)
If mBtask Is Nothing Then
'Set MyBaseTask = Nothing
Else
Set MyBaseTask = New basetask
mBtask.CopyStrip2 MyBaseTask
run_in_basestack1 = mBtask Is Basestack1
End If
End Property

Private Sub Command1_Click()
trace = True
STq = False
STbyST = True
End Sub
Private Sub Command2_Click()
trace = True
STq = True
STbyST = False
End Sub

Private Sub Command3_Click()
NOEXECUTION = True
trace = True
STEXIT = True
End Sub
Public Sub ComputeNow()
stackshow MyBaseTask
End Sub

Private Sub compute_KeyDown(KeyCode As Integer, shift As Integer)
Dim excode As Long
If KeyCode = 13 Then
KeyCode = 0
    If Compute.Prompt = "? " Then
        gList3(2).backcolor = &H3B3B3B
        TestShowCode = False
        stackshow MyBaseTask
    Else
            'bypasstrace = True
         If Compute <> "" Then STbyST = True
        If run_in_basestack1 Then
        If QRY And Form1.Visible Then
        INK$ = Compute + Chr$(13)
        
        End If

        Else
        tracecode = Compute
        End If
        
    End If
ElseIf KeyCode = 8 Then
If Compute = vbNullString Then
    If Compute.Prompt <> ">" Then
        Compute.Prompt = ">"
         If pagio$ = "GREEK" Then
            Compute.ThisKind = "     ������ ������� | ������� | Backspace ������� �����"
        Else
            Compute.ThisKind = "     Inline command | STOP | Backspace Select Mode"
        End If
    Else
        Compute.Prompt = "? "
        If pagio$ = "GREEK" Then
            Compute.ThisKind = "     [����,] ���� | Backspace ������� �����"
        Else
            Compute.ThisKind = "     [Expr,] Expr | Backspace Select Mode"
        End If
    End If
    Compute.vartext = vbNullString
KeyCode = 0
Exit Sub
End If
End If
End Sub
'M2000 [������� - CONTROL]

Private Sub Form_Activate()
'
trace = True
If stolemodalid = 0 Then
If Modalid <> 0 Then
stolemodalid = Modalid
Modalid = Rnd * 645677887
End If

End If
End Sub

Private Sub Form_Deactivate()
If stolemodalid <> 0 Then
Modalid = stolemodalid
stolemodalid = 0
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 27 Then
KeyCode = 0
Unload Me
ElseIf KeyCode = 13 Then
If Not EXECUTED Then
If Not STq Then
STbyST = True
End If
End If
End If
End Sub

Private Sub Form_Load()
Dim i As Long
height1 = 5280 * DYP / 15
width1 = 7860 * DXP / 15
MoveFormToOtherMonitorOnly Me
lastfactor = 1
LastWidth = -1
HelpLastWidth = -1
PopUpLastWidth = -1
setupxy = 20
lastfactor = ScaleDialogFix(SizeDialog)
ScaleDialog lastfactor, LastWidth
gList4.NoCaretShow = True
gList4.restrictLines = 3
gList4.CenterText = True
gList2.CapColor = rgb(255, 160, 0)
gList2.HeadLine = vbNullString

gList2.FloatList = True
gList2.FloatLimitTop = VirtualScreenHeight() - players(0).Yt * 2
gList2.FloatLimitLeft = VirtualScreenWidth() - players(0).Xt * 2
gList2.MoveParent = True
'gList2.enabled = True
gList1.DragEnabled = False
gList1.AutoPanPos = True
Set testpad = New TextViewer
gList1.NoWheel = True
Set testpad.Container = gList1
testpad.FileName = vbNullString
testpad.glistN.DropEnabled = False
testpad.glistN.DragEnabled = False
testpad.glistN.LeftMarginPixels = 8
testpad.NoMark = True
testpad.NoColor = False
testpad.EditDoc = False
testpad.nowrap = False
testpad.enabled = True
Set Compute = New myTextBox
Set Compute.Container = gList0
Compute.MaxCharLength = 500 ' as a limit
Compute.locked = False
Compute.enabled = True
Compute.Retired
Set Label(0).Container = gList3(0)
Set Label(1).Container = gList3(1)
Set Label(2).Container = gList3(2)
If pagio$ = "GREEK" Then
gList2.HeadLine = "�������"
Compute.Prompt = "? "
Compute.ThisKind = "     [����,] ���� + Enter | Backspace ������� �����"
Compute.FadePartColor = &H777777
Label(0).Prompt = "�����: "
Label(1).Prompt = "������: "
Label(2).Prompt = "�������: "
gList4.additemFast "������� ����"
gList4.additemFast "���� ���"
gList4.additemFast "�������"
Else
gList2.HeadLine = "Control"
Compute.Prompt = "? "
Compute.ThisKind = "     [Expr,] Expr + Enter | Backspace Select Mode"
Compute.FadePartColor = &H777777
Label(0).Prompt = "Module: "
Label(1).Prompt = "Id: "
Label(2).Prompt = "Next: "
gList4.additemFast "Next Step"
gList4.additemFast "Slow Flow"
gList4.additemFast "Stop"
End If
gList2.HeadlineHeight = gList2.HeightPixels
gList2.PrepareToShow
gList4.NoPanRight = False
gList4.SingleLineSlide = True
gList4.VerticalCenterText = True
gList4.enabled = True
gList4.ListindexPrivateUse = 0
gList4.ShowMe
End Sub
Sub FillAgainLabels()
Dim oldindex As Long
If pagio$ = "GREEK" Then
If gList2.HeadLine = "Control" Then gList2.HeadLine = "�������"
If Compute.Prompt = ">" Then
        Compute.ThisKind = "     ������ ������� + Enter | Backspace ������� �����"
Else
        Compute.ThisKind = "     [����,] ���� + Enter | Backspace ������� �����"
End If
Label(1).Prompt = "������: "
Label(2).Prompt = "�������: "
oldindex = gList4.ListIndex
gList4.Clear
gList4.additemFast "������� ����"
gList4.additemFast "���� ���"
gList4.additemFast "�������"
gList4.ListindexPrivateUse = oldindex
gList4.PrepareToShow
gList2.PrepareToShow
 Compute.glistN.ShowMe
         

        
Else
If gList2.HeadLine = "�������" Then gList2.HeadLine = "Control"
If Compute.Prompt = ">" Then
            Compute.ThisKind = "     Inline command + Enter | Backspace Select Mode"
    Else
            Compute.ThisKind = "     [Expr,] Expr + Enter | Backspace Select Mode"
    End If
Label(1).Prompt = "Id: "
Label(2).Prompt = "Next: "
oldindex = gList4.ListIndex
gList4.Clear
gList4.additemFast "Next Step"
gList4.additemFast "Slow Flow"
gList4.additemFast "Stop"
gList4.ListindexPrivateUse = oldindex
gList4.PrepareToShow
gList2.PrepareToShow
Compute.glistN.ShowMe
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
testpad.Dereference
Compute.Dereference
Set MyBaseTask = Nothing
trace = False
STq = True
End Sub




Private Sub gList1_CheckGotFocus()
gList1.backcolor = &H606060
gList1.ShowMe2
End Sub

Private Sub gList1_CheckLostFocus()

gList1.backcolor = &H3B3B3B
gList1.ShowMe2
End Sub


Private Sub gList1_MouseUp(x As Single, y As Single)
Dim monitor As Long
If Not Form4.Visible Then
monitor = FindFormSScreen(Form1)
Else
monitor = FindFormSScreen(Form4)
End If
abt = False
sHelp gList2.HeadLine, testpad.Text, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7

If TestShowCode Then
vHelp True
Form4.label1.SelStartSilent = testpad.SelStart
Form4.label1.SelLengthSilent = 0
Form4.label1.SelectionColor = rgb(255, 64, 128)
If testpad.SelStart > 0 And testpad.SelLength > 0 Then Form4.label1.SelLength = testpad.SelLength: Form4.label1.glistN.ShowMe
Else
vHelp Not Form4.Visible
End If
End Sub


Private Sub gList2_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)

If item = -1 Then
FillThere thisHDC, thisrect, gList2.CapColor
FillThereMyVersion thisHDC, thisrect, &H999999
skip = True
End If

End Sub
Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If gList2.DoubleClickCheck(Button, item, x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor, -1) Then
            Me.Visible = False
            Unload Me
End If
End Sub

Public Property Let Label1prompt(ByVal Index As Long, ByVal RHS As String)
Label(Index).Prompt = RHS
End Property

Public Property Get label1(ByVal Index As Long) As String
label1 = Label(Index)
End Property

Public Property Let label1(ByVal Index As Long, ByVal RHS As String)
Label(Index) = RHS
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

Private Sub gList2_GotFocus()
tracecounter = 100
End Sub

Private Sub gList2_LostFocus()
doubleclick = 0

End Sub

Private Sub gList2_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If Button <> 0 Then tracecounter = 100
End Sub

Private Sub gList2_MouseUp(x As Single, y As Single)
tracecounter = 0
End Sub




Private Sub glist3_CheckGotFocus(Index As Integer)
Dim s$, z$
gList4.SetFocus
If Index < 2 Then
abt = False

vH_title$ = vbNullString
s$ = Label(Index)
If Index = 1 Then
   
        Dim i As Long
        If MyBaseTask.ExistVar2(s$) Then
            IsStr1 MyBaseTask, "Type$(" + s$ + ")", z$
            
            If AscW(s$ + Mid$(" �", Abs(pagio$ = "GREEK") + 1)) < 128 Then
                sHelp "Static Variable", s$ & vbCrLf & "Type " + z$, vH_x, vH_y
            Else
                sHelp "������� ���������", s$ & vbCrLf & "����� " + z$, vH_x, vH_y
            End If
        
        ElseIf GetlocalVar(s$, i) Then
            IsStr1 MyBaseTask, "Type$(" + s$ + ")", z$
            If AscW(s$ + Mid$(" �", Abs(pagio$ = "GREEK") + 1)) < 128 Then
                sHelp "Local Identifier", s$ & vbCrLf & "Type " + z$, vH_x, vH_y
            Else
                sHelp "������ �������������", s$ & vbCrLf & "����� " + z$, vH_x, vH_y
            End If
        ElseIf GetGlobalVar(s$, i) Then
        If AscW(s$ + Mid$(" �", Abs(pagio$ = "GREEK") + 1)) < 128 Then
            sHelp "Global Identifier", s$, vH_x, vH_y
            
            Else
            sHelp "������ �������������", s$, vH_x, vH_y
            
            End If
        ElseIf GetSub(s$, i) Then
            '
            z$ = Label(2)
            If MaybeIsSymbol(z$, "=") Then
            If AscW(s$ + Mid$(" �", Abs(pagio$ = "GREEK") + 1)) < 128 Then
                sHelp "New Identifier", s$, vH_x, vH_y
            Else
                sHelp "��� �������������", s$, vH_x, vH_y
            End If
            Else
                GoTo JUMPHERE
            End If
            
        ElseIf ismine(s$) Then
            fHelp MyBaseTask, s$, AscW(s$ + Mid$(" �", Abs(pagio$ = "GREEK") + 1)) < 128
        Else
        z$ = Label(2)
        If MaybeIsSymbol(z$, "=") Then
            If AscW(s$ + Mid$(" �", Abs(pagio$ = "GREEK") + 1)) < 128 Then
                sHelp "New Identifier", s$, vH_x, vH_y
            Else
                sHelp "��� �������������", s$, vH_x, vH_y
            End If
        End If
        End If
    
    vHelp
ElseIf MyBaseTask.IamLambda Then
sHelp s$, LambdaList(MyBaseTask), vH_x, vH_y
vHelp
ElseIf MyBaseTask.IamThread Then
If MyBaseTask.Process Is Nothing Then
Else
sHelp s$, MyBaseTask.Process.CodeData, vH_x, vH_y
vHelp
End If
Else
If Index = 0 Then
i = MyBaseTask.OriginalCode
JUMPHERE:
Dim aa As Long, aaa As String
aa = i
aaa = SBcode(aa)
If Left$(aaa, 10) = "'11001EDIT" Then
SetNextLine aaa
End If
Dim monitor As Long
If Not Form4.Visible Then
monitor = FindFormSScreen(Form1)
Else
monitor = FindFormSScreen(Form4)
End If
If aa > 0 Then
sHelp s$, aaa, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
vHelp Not Form4.Visible
ElseIf subHash.Find(s$, aa) Then
sHelp s$, SBcode(aa), (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
vHelp Not Form4.Visible

ElseIf subHash.Find(subHash.LastKnown, aa) Then
sHelp s$, SBcode(aa), (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
vHelp Not Form4.Visible
Else
    fHelp MyBaseTask, s$, AscW(s$ + Mid$(" �", Abs(pagio$ = "GREEK") + 1)) < 128
    End If
Else
Select Case Left$(LTrim(Label(2)) + " ", 1)
Case "?", "!", " ", ".", ":", Is >= "A", Chr$(10), """"
    fHelp MyBaseTask, s$, AscW(s$ + Mid$(" �", Abs(pagio$ = "GREEK") + 1)) < 128
End Select
End If
End If
ElseIf Index = 2 Then
TestShowCode = Not TestShowCode
If TestShowCode Then
gList3(2).backcolor = &H606060
Label(2) = Label(2)
Else
gList3(2).backcolor = &H3B3B3B
Label(2) = Label(2)
testpad.SetRowColumn 1, 1
End If
stackshow MyBaseTask
End If
End Sub

Private Sub gList4_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim a As RECT, b As RECT
CopyFromLParamToRect a, thisrect
CopyFromLParamToRect b, thisrect
a.Left = a.Left + 1 * lastfactor
a.Right = gList4.WidthPixels
b.Right = gList4.WidthPixels
 If item = gList4.ListIndex Then
   If EXECUTED Then
   FillBack thisHDC, b, &H77FF77
   Else
   
             FillBack thisHDC, b, &H77FFFF
             End If
             'EXECUTED = False
              SetTextColor thisHDC, 0
              b.top = b.Bottom - 1 * lastfactor
       
            FillBack thisHDC, b, &H777777
           
    Else
          
          
    SetTextColor thisHDC, gList4.forecolor
    b.top = b.Bottom - 1
    FillBack thisHDC, b, 0
    End If
    If item = gList4.ListIndex Then
  a.Left = a.Left + 1 * lastfactor + gList4.PanPosPixels
  gList4.forecolor = rgb(128, 0, 128)
  End If
   
   
   PrintItem thisHDC, gList4.list(item), a
    skip = True
End Sub
 
Private Sub gList4_PanLeftRight(direction As Boolean)
EXECUTED = True
Action
End Sub

Private Sub gList4_selected(item As Long)
EXECUTED = False
gList4.ShowMe

End Sub

Private Sub gList4_Selected2(item As Long)
EXECUTED = True
Action
End Sub
Private Sub Action()
EXECUTED = True ' SO CHANGE THE BACKGROUND COLOR COLOR
Select Case gList4.ListIndex
Case 0
trace = True
STq = False
STbyST = True
Case 1
trace = True
STq = True
STbyST = False
Case 2
If Not TaskMaster Is Nothing Then
TaskMaster.Dispose
MyBaseTask.ThrowThreads
End If
Modalid = 0
NOEXECUTION = True
trace = True
STq = False
STbyST = True
End Select
gList4.PanPos = 0
gList4.ShowMe2
End Sub

Private Sub gList4_softSelected(item As Long)
EXECUTED = Not EXECUTED
gList4.ShowMe


End Sub
 Private Sub PrintItem(mHdc As Long, c As String, R As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_CENTER Or DT_VCENTER)
    DrawText mHdc, StrPtr(c), -1, R, way
    End Sub
Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)

If Button = 1 Then
    
    If lastfactor = 0 Then lastfactor = 1

    If bordertop < 150 Then
    If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
    dr = True
    tracecounter = 100
    mousepointer = vbSizeNWSE
    Lx = x
    ly = y
    End If
    
    Else
    If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then
    dr = True
    tracecounter = 100
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
If Button = 0 Then dr = False: drmove = False
If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then mousepointer = vbSizeNWSE Else If Not (dr Or drmove) Then mousepointer = 0
 Else
 If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then mousepointer = vbSizeNWSE Else If Not (dr Or drmove) Then mousepointer = 0
End If
If dr Then



If bordertop < 150 Then

        If y < (Height - 150) Or y > Height Then addy = (y - ly)
     If x < (Width - 150) Or x > Width Then addX = (x - Lx)
     
Else
    If y < (Height - bordertop) Or y > Height Then addy = (y - ly)
        If x < (Width - borderleft) Or x > Width Then addX = (x - Lx)
    End If
    

    
  If Not ExpandWidth Then addX = 0
        If lastfactor = 0 Then lastfactor = 1
        factor = lastfactor

        
  
        once = True
        If Height > VirtualScreenHeight() Then addy = -(Height - VirtualScreenHeight()) + addy
        If Width > VirtualScreenWidth() Then addX = -(Width - VirtualScreenWidth()) + addX
        If (addy + Height) / height1 > 0.4 And ((Width + addX) / width1) > 0.4 Then
   
        If addy <> 0 Then SizeDialog = ((addy + Height) / height1)
        lastfactor = ScaleDialogFix(SizeDialog)


        If ((Width * lastfactor / factor + addX) / Height * lastfactor / factor) < (width1 / height1) Then
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
        gList1.PrepareToShow
        'testpad.Render
        ly = ly * lastfactor / factor
        End If
        Else
        Lx = x
        ly = y
   
End If
once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)

If dr Then Me.mousepointer = 0
tracecounter = 0
dr = False
End Sub
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)

lastfactor = factor
setupxy = 20 * factor
bordertop = 10 * dv15 * factor

borderleft = bordertop
allwidth = width1 * factor
allheight = height1 * factor
itemWidth = allwidth - 2 * borderleft
itemwidth3 = itemWidth * 2 / 5
itemwidth2 = itemWidth * 3 / 5 - borderleft
move Left, top, allwidth, allheight
FontTransparent = False  ' clear background  or false to write over
gList2.move borderleft, bordertop, itemWidth, bordertop * 3
gList2.FloatLimitTop = VirtualScreenHeight() - bordertop - bordertop * 3
gList2.FloatLimitLeft = VirtualScreenWidth() - borderleft * 3
gList3(0).move borderleft, bordertop * 5, itemwidth2, bordertop * 4
gList3(1).move borderleft, bordertop * 9, itemwidth2, bordertop * 4
gList3(2).move borderleft, bordertop * 13, itemwidth2, bordertop * 4
gList4.move borderleft * 2 + itemwidth2, bordertop * 5, itemwidth3, bordertop * 12
gList1.move borderleft, bordertop * 18, itemWidth, bordertop * 12
gList0.move borderleft, bordertop * 31, itemWidth, bordertop * 3
End Sub
Function ScaleDialogFix(ByVal factor As Single) As Single
gList2.FontSize = 14.25 * factor * dv15 / 15
factor = gList2.FontSize / 14.25 / dv15 * 15
gList1.FontSize = 11.25 * factor * dv15 / 15
gList4.FontSize = 12 * factor * dv15 / 15
factor = gList1.FontSize / 11.25 / dv15 * 15
gList3(0).FontSize = gList1.FontSize
gList3(1).FontSize = gList1.FontSize
gList3(2).FontSize = gList1.FontSize
gList0.FontSize = gList1.FontSize
ScaleDialogFix = factor
End Function

Public Sub hookme(this As gList)
If Not this Is Nothing Then this.NoWheel = True
End Sub
Sub ByeBye()
Unload Me
End Sub

