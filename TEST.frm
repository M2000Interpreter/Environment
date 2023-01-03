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
Public previewKey As Boolean, switchview As Integer, Busy As Boolean
Private Label(0 To 2) As New myTextBox
Dim run_in_basestack1 As Boolean
Dim MyBaseTask As New basetask
Dim setupxy As Single
Dim Lx As Long, lY As Long, dr As Boolean, drmove As Boolean
Dim prevx As Long, prevy As Long
Dim bordertop As Long, borderleft As Long
Dim allheight As Long, allwidth As Long, itemWidth As Long, itemwidth3 As Long, itemwidth2 As Long
Dim height1 As Long, width1 As Long
Dim doubleclick As Long
Dim para(2) As Long, pospara(2) As Long, selpresrv(2) As Long
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Dim EXECUTED As Boolean
Dim stolemodalid As Variant, padtitle As String
Public Sub generalFkey(a As Integer)
Dim monitor As Long, titl$, once As Boolean
If once Then Exit Sub
once = True
Debug.Print "OK", a
If a = 2 Then
gList0.SetFocus
ElseIf a = 3 Then
glist3_CheckGotFocus 2
ElseIf a = 5 Then
gList4.ListIndex = 0
gList4_Selected2 0
ElseIf a = 6 Then
stoponerror = True
If pagio$ = "GREEK" Then
gList4.list(1) = "Στο Λάθος F6"
Else
gList4.list(1) = "To Error F6"
End If
' "Slow Flow" "Αργή Ροή"

gList4.ListIndex = 1
gList4_Selected2 1
ElseIf a = 7 Then
If pagio$ = "GREEK" Then
gList4.list(1) = "Αργή Ροή F7"
Else
gList4.list(1) = "Slow Flow F7"
End If
stoponerror = False
gList4.ListIndex = 1
gList4_Selected2 1
ElseIf a = 8 Then
gList4.ListIndex = 2
gList4_Selected2 2
ElseIf switchview = 2 Then
    If a = 1 Then
        Errorlog.EmptyDoc
        stackshow MyBaseTask
    ElseIf a = 4 Then
        If pagio$ = "GREEK" Then titl$ = "Καταγραφικό Λαθών" Else titl$ = "Error Log"
        GoSub findwindow
        sHelp titl$, Errorlog.textDoc, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
        If Not Form4.Visible Then vHelp True
        Form4.label1.SelStartSilent = testpad.SelStart
        vHelp
    End If
ElseIf a = 4 Then
    GoSub findwindow
    sHelp gList2.HeadLine, testpad.Text, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
    If TestShowCode Then
        If Not Form4.Visible Then vHelp True
        Form4.label1.Text = testpad.Text
        
        Form4.label1.SelStartSilent = testpad.SelStart
        Form4.label1.SelLengthSilent = testpad.SelLength
        Form4.label1.SelectionColor = rgb(255, 64, 128)
         Form4.gList1.ShowMe '  label1.glistN.Show
       ' If testpad.SelStart > 0 And testpad.SelLength > 0 Then Form4.label1.SelLength = testpad.SelLength: Form4.label1.glistN.ShowMe
    Else
       vHelp
    End If
    
End If
once = False
Exit Sub
findwindow:
    If Not Form4.Visible Then
        monitor = FindFormSScreen(Form1)
    Else
        monitor = FindFormSScreen(Form4)
    End If
    abt = False
Return
End Sub
Public Property Get Process()
    Set Process = MyBaseTask
End Property
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
        gList3(2).BackColor = &H3B3B3B
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
            Compute.ThisKind = "     Ένθεση Εντολής | ΔΙΑΚΟΠΗ | Backspace Επιλογή Τύπου"
        Else
            Compute.ThisKind = "     Inline command | STOP | Backspace Select Mode"
        End If
    Else
        Compute.Prompt = "? "
        If pagio$ = "GREEK" Then
            Compute.ThisKind = "     [Εκφρ,] Εκφρ | Backspace Επιλογή Τύπου"
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
'M2000 [ΕΛΕΓΧΟΣ - CONTROL]

Private Sub Form_Activate()
'
If HOOKTEST <> 0 Then UnHook HOOKTEST
trace = True
If stolemodalid = 0 Then
If Modalid <> 0 Then
stolemodalid = Modalid
Modalid = Rnd * 645677887
End If
Else
End If
If Not Busy Then generalFkey 2 Else Debug.Print "test is busy"
End Sub

Private Sub Form_Deactivate()
If stolemodalid <> 0 Then
Modalid = stolemodalid
stolemodalid = 0
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyTab And (shift And &H2 = 2) Then
choosenext
Sleep 100
KeyCode = 0
ElseIf KeyCode = 27 Then
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
Busy = True
height1 = 5280 * DYP / 15
width1 = 7860 * DXP / 15
MoveFormToOtherMonitorOnly Me
lastfactor = 1
LastWidth = -1
HelpLastWidth = -1
PopUpLastWidth = -1
setupxy = 20
padtitle = "F4 - Copy to help form | F1 clear error log"
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
gList1.bypassfirstClick = True
gList1.NoWheel = True
Set testpad.Container = gList1
testpad.FileName = vbNullString
testpad.glistN.CapColor = rgb(128, 128, 0)
testpad.glistN.DropEnabled = False
testpad.glistN.DragEnabled = False
testpad.glistN.LeftMarginPixels = 8
testpad.NoMark = True
testpad.NoColor = False
testpad.EditDoc = False
testpad.nowrap = False
paneltitle 0&
testpad.Enabled = True
Set Compute = New myTextBox
Set Compute.Container = gList0
Compute.MaxCharLength = 500 ' as a limit
Compute.locked = False
Compute.Enabled = True
Compute.Retired
Set Label(0).Container = gList3(0)
Set Label(1).Container = gList3(1)
Set Label(2).Container = gList3(2)
If pagio$ = "GREEK" Then
gList2.HeadLine = "Έλεγχος"
Compute.Prompt = "? "
Compute.ThisKind = "     [Εκφρ,] Εκφρ + Enter | Backspace Επιλογή Τύπου"
Compute.FadePartColor = &H777777
Label(0).Prompt = "Τμήμα: "
Label(1).Prompt = "Εντολή: "
Label(2).Prompt = "Επόμενο: "
gList4.additemFast "Επόμενο Βήμα F5"
If stoponerror Then
gList4.additemFast "Στο Λάθος F6"
Else
gList4.additemFast "Αργή Ροή F7"
End If
gList4.additemFast "Διακοπή F8"
Else
gList2.HeadLine = "Control"
Compute.Prompt = "? "
Compute.ThisKind = "     [Expr,] Expr + Enter | Backspace Select Mode"
Compute.FadePartColor = &H777777
Label(0).Prompt = "Module: "
Label(1).Prompt = "Id: "
Label(2).Prompt = "Next: "
gList4.additemFast "Next Step F5"
If stoponerror Then
gList4.additemFast "To Error F6"
Else
gList4.additemFast "Slow Flow F7"
End If
gList4.additemFast "Stop F8"
End If
gList2.HeadlineHeight = gList2.HeightPixels
gList2.PrepareToShow
gList4.NoPanRight = False
gList4.SingleLineSlide = True
gList4.VerticalCenterText = True
gList4.Enabled = True
gList4.ListindexPrivateUse = 0
gList4.ShowMe
Busy = False
End Sub
Sub FillAgainLabels()
Dim oldindex As Long
If pagio$ = "GREEK" Then
If gList2.HeadLine = "Control" Then gList2.HeadLine = "Έλεγχος"
If Compute.Prompt = ">" Then
        Compute.ThisKind = "     Ένθεση Εντολής + Enter | Backspace Επιλογή Τύπου"
Else
        Compute.ThisKind = "     [Εκφρ,] Εκφρ + Enter | Backspace Επιλογή Τύπου"
End If
Label(1).Prompt = "Εντολή: "
Label(2).Prompt = "Επόμενο: "
oldindex = gList4.ListIndex
gList4.Clear
gList4.additemFast "Επόμενο Βήμα F5"
If stoponerror Then
gList4.additemFast "Στο Λάθος F6"
Else
gList4.additemFast "Αργή Ροή F7"
End If
gList4.additemFast "Διακοπή F8"


gList4.ListindexPrivateUse = oldindex
gList4.PrepareToShow
gList2.PrepareToShow
 Compute.glistN.ShowMe
         

        
Else
If gList2.HeadLine = "Έλεγχος" Then gList2.HeadLine = "Control"
If Compute.Prompt = ">" Then
            Compute.ThisKind = "     Inline command + Enter | Backspace Select Mode"
    Else
            Compute.ThisKind = "     [Expr,] Expr + Enter | Backspace Select Mode"
    End If
Label(1).Prompt = "Id: "
Label(2).Prompt = "Next: "
oldindex = gList4.ListIndex
gList4.Clear
gList4.additemFast "Next Step F5"
If stoponerror Then
gList4.additemFast "To Error F6"
Else
gList4.additemFast "Slow Flow F7"
End If
gList4.additemFast "Stop F8"
gList4.ListindexPrivateUse = oldindex
gList4.PrepareToShow
gList2.PrepareToShow
Compute.glistN.ShowMe
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
If Busy Then Cancel = True: Exit Sub
testpad.Dereference
Compute.Dereference
Set MyBaseTask = Nothing
trace = False
STq = True
End Sub




Private Sub gList0_Fkey(a As Integer)
generalFkey a
End Sub


Private Sub gList1_CheckGotFocus()
gList1.BackColor = &H505050
'testpad.SelLengthSilent = 0
'gList1.ShowMe2
SetPanelPos switchview
End Sub

Private Sub gList1_CheckLostFocus()
gList1.BackColor = &H3B3B3B
gList1.ShowMe2
GetPanelPos
End Sub


Private Sub gList1_Fkey(a As Integer)

generalFkey a
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

Public Property Let Label1prompt(ByVal index As Long, ByVal RHS As String)
Label(index).Prompt = RHS
End Property

Public Property Get label1(ByVal index As Long) As String
label1 = Label(index)
End Property

Public Property Let label1(ByVal index As Long, ByVal RHS As String)
Label(index) = RHS
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

Private Sub gList2_Fkey(a As Integer)
generalFkey a
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




Private Sub glist3_CheckGotFocus(index As Integer)
Dim s$, z$
gList4.SetFocus
On Error GoTo there1
If index < 2 Then
abt = False

vH_title$ = vbNullString
s$ = Label(index)
If index = 1 Then
   
        Dim i As Long
        If MyBaseTask.ExistVar2(s$) Then
            IsStr1 MyBaseTask, "Type$(" + s$ + ")", z$
            
            If AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128 Then
                sHelp "Static Variable", s$ & vbCrLf & "Type " + z$, vH_x, vH_y
            Else
                sHelp "Στατική Μεταβλητή", s$ & vbCrLf & "Τύπος " + z$, vH_x, vH_y
            End If
        
        ElseIf GetlocalVar(s$, i) Then
            IsStr1 MyBaseTask, "Type$(" + s$ + ")", z$
            If AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128 Then
                sHelp "Local Identifier", s$ & vbCrLf & "Type " + z$, vH_x, vH_y
            Else
                sHelp "Τοπικό Αναγνωριστικό", s$ & vbCrLf & "Τύπος " + z$, vH_x, vH_y
            End If
        ElseIf GetGlobalVar(s$, i) Then
        If AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128 Then
            sHelp "Global Identifier", s$, vH_x, vH_y
            
            Else
            sHelp "Γενικό Αναγνωριστικό", s$, vH_x, vH_y
            
            End If
        ElseIf GetSub(s$, i) Then
            '
            z$ = Label(2)
            If MaybeIsSymbol(z$, "=") Then
            If AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128 Then
                sHelp "New Identifier", s$, vH_x, vH_y
            Else
                sHelp "Νέο Αναγνωριστικό", s$, vH_x, vH_y
            End If
            Else
                GoTo JUMPHERE
            End If
            
        ElseIf ismine(s$) Then
            fHelp MyBaseTask, s$, AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128
        Else
        z$ = Label(2)
        If MaybeIsSymbol(z$, "=") Then
            If AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128 Then
                sHelp "New Identifier", s$, vH_x, vH_y
            Else
                sHelp "Νέο Αναγνωριστικό", s$, vH_x, vH_y
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
If index = 0 Then
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
    fHelp MyBaseTask, s$, AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128
    End If
Else
Select Case Left$(LTrim(Label(2)) + " ", 1)
Case "?", "!", " ", ".", ":", Is >= "A", Chr$(10), """"
    fHelp MyBaseTask, s$, AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128
End Select
End If
End If
ElseIf index = 2 Then
    GetPanelPos
    switchview = switchview + 1: If switchview > 2 Then switchview = 0
    TestShowCode = switchview = 1
    If TestShowCode Then
        gList3(2).BackColor = &H606060
        Label(2) = Label(2)
    Else
        If switchview = 2 Then
            gList3(2).BackColor = &H3B3B6B
            Label(2) = Label(2)
        Else
            gList3(2).BackColor = &H3B3B3B
            Label(2) = Label(2)
            testpad.SetRowColumn 1, 1
            
        End If
    End If
  paneltitle switchview
  stackshow MyBaseTask
  testpad.Enabled = True
End If
there1:
End Sub

Private Sub gList3_Fkey(index As Integer, a As Integer)
generalFkey a
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
          
          
    SetTextColor thisHDC, gList4.ForeColor
    b.top = b.Bottom - 1
    FillBack thisHDC, b, 0
    End If
    If item = gList4.ListIndex Then
  a.Left = a.Left + 1 * lastfactor + gList4.PanPosPixels
  gList4.ForeColor = rgb(128, 0, 128)
  End If
   
   
   PrintItem thisHDC, gList4.list(item), a
    skip = True
End Sub
 
Private Sub gList4_Fkey(a As Integer)
generalFkey a
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
    MousePointer = vbSizeNWSE
    Lx = x
    lY = y
    End If
    
    Else
    If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then
    dr = True
    tracecounter = 100
    MousePointer = vbSizeNWSE
    Lx = x
    lY = y
    End If
    End If

End If
End Sub
Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim addX As Long, addy As Long, factor As Single, once As Boolean
If once Then Exit Sub
If Button = 0 Then dr = False: drmove = False
If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then MousePointer = vbSizeNWSE Else If Not (dr Or drmove) Then MousePointer = 0
 Else
 If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then MousePointer = vbSizeNWSE Else If Not (dr Or drmove) Then MousePointer = 0
End If
If dr Then



If bordertop < 150 Then

        If y < (Height - 150) Or y > Height Then addy = (y - lY)
     If x < (Width - 150) Or x > Width Then addX = (x - Lx)
     
Else
    If y < (Height - bordertop) Or y > Height Then addy = (y - lY)
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
        lY = lY * lastfactor / factor
        End If
        Else
        Lx = x
        lY = y
   
End If
once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)

If dr Then Me.MousePointer = 0
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
On Error Resume Next
If Not testpad Is Nothing Then
If Not testpad.glistN Is Nothing Then
testpad.NewTitle testpad.Title, 4 * factor
'testpad.ReColor
End If
End If
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
Public Sub ResetPanelPos()
Dim i As Long
For i = 0 To 2: para(i) = 0: Next i
lastpanel = 0
End Sub
Public Sub paneltitle(ByVal that As Long)
Dim oldlength As Long
    Select Case that
    Case 0
        testpad.Title = "Stack & Variables | F4 = copy to help form"
        testpad.ReplaceTitle = " "
    Case 1
        
        testpad.Title = "Code View | F4 = copy to help form"
        testpad.ReplaceTitle = " "
    Case 2
        testpad.Title = "Error Log | F1 - clear | F4 = copy to help form"
        testpad.ReplaceTitle = " "
    End Select
End Sub

Public Sub GetPanelPos(Optional ByVal that As Long = -1)
Select Case that
Case 0 To 2
    
    para(that) = testpad.mDoc.MarkParagraphID
    pospara(that) = testpad.ParaSelStart
    selpresrv(that) = testpad.SelLength
    testpad.nowrap = that > 0
    testpad.Enabled = Not STbyST
Case -1
    para(lastpanel) = testpad.mDoc.MarkParagraphID
    pospara(lastpanel) = testpad.ParaSelStart
End Select
End Sub
Public Sub SetPanelPos(ByVal that As Long)
Select Case that
Case 0 To 2
    lastpanel = that
    If para(that) = 0 Then Exit Sub
    With testpad
        .SelLengthSilent = selpresrv(that)
        .mDoc.MarkParagraphID = para(that)
        .glistN.Enabled = False
        .ParaSelStart = pospara(that)
        .glistN.Enabled = True
        .glistN.ShowMe
    End With
End Select
End Sub

