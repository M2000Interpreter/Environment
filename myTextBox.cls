VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "myTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' If you copy and paste code
' don't forget to set default property to vartext (changed from text)
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Event ValidString(ThatString As String, setpos As Long)
Event KeyDown(KeyCode As Integer, shift As Integer)
Event LostFocus()
Event GotFocus()
Event KeyPress(KeyAscii As Integer)
Event MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Event SpinnerValue(ThisValue As Long)
Event DeployMenu()
Event CtrlPlusF1()
Event SyncKeyboardUnicode(a As String)
Event Fkey(a As Integer)
Private mprompt As String
Private mkind As String, mPartColor As Long
Public SkipProcAtExitFocus As Boolean
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000&
Private Const GMEM_DISCARDABLE = &H100&
Private Const GMEM_DISCARDED = &H4000&
Private Const GMEM_FIXED = &H0&
Private Const GMEM_INVALID_HANDLE = &H8000&
Private Const GMEM_LOCKCOUNT = &HFF&
Private Const GMEM_MODIFY = &H80&
Private Const GMEM_MOVEABLE = &H2&
Private Const GMEM_NOCOMPACT = &H10&
Private Const GMEM_NODISCARD = &H20&
Private Const GMEM_NOT_BANKED = &H1000&
Private Const GMEM_NOTIFY = &H4000&
Private Const GMEM_SHARE = &H2000&
Private Const GMEM_VALID_FLAGS = &H7F72&
Private Const GMEM_ZEROINIT = &H40&
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetClipboardData Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Declare Function SetClipboardData Lib "user32" _
    (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function OpenClipboard Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" _
    () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Pad$, mStayOpen As Boolean
Public Property Get Text() As String
If glistN.listcount = 0 Then
glistN.additem mprompt
Else
Text = glistN.list(0)
End If
End Property
Public Property Get vartext() As String
Attribute vartext.VB_UserMemId = 0
If glistN.listcount = 0 Then
glistN.additem ""
Else
vartext = Mid$(glistN.list(0), 1 + Len(mprompt))
End If
End Property
Public Property Let vartext(ByVal RHS As String)
glistN.restrictLines = 1
glistN.HeadLine = vbNullString ' no info
glistN.HeadlineHeight = 0

If glistN.listcount = 0 Then
glistN.additem mprompt + RHS
Else
glistN.list(0) = mprompt + RHS
End If
glistN.SelStart = Len(mprompt + RHS) + 1
glistN.ShowMe
End Property
Public Property Let Text(ByVal RHS As String)
'Dim oldright As Long
'oldright = glistN.LeftMarginPixels
glistN.restrictLines = 1
glistN.HeadLine = vbNullString ' no info
glistN.HeadlineHeight = 0
If glistN.listcount = 0 Then
glistN.additem RHS
Else
glistN.list(0) = RHS
End If
'glistN.LeftMarginPixels = oldright

glistN.ShowMe2
End Property
Public Property Get Info() As String

Info = glistN.HeadLine

End Property

Public Property Let Info(ByVal RHS As String)
locked = False
glistN.restrictLines = 0
glistN.HeadLineHeightMinimum = glistN.HeightPixels
glistN.HeadLine = RHS

glistN.ShowMe
locked = True
End Property

Public Property Get top() As Long
top = glistN.TopTwips
End Property
Public Property Let top(RHS As Long)
glistN.TopTwips = RHS
End Property
Public Property Set Container(glistNN As gList)
Set glistN = glistNN
glistNN.enabled = True
Info = vbNullString
If glistNN.listcount = 0 Then Text = vbNullString
glistN.NoCaretShow = True
glistN.VerticalCenterText = True
If Not glistN.CenterText Then glistN.LeftMarginPixels = 8
End Property


Public Property Get Prompt() As String
Prompt = mprompt
End Property

Public Property Let Prompt(ByVal RHS As String)
mprompt = RHS
glistN.PromptLineIdent = Len(RHS)
Text = RHS
End Property



Private Sub glistN_CtrlPlusF1()
RaiseEvent CtrlPlusF1
End Sub

Private Sub glistN_DeployMenu()
'
RaiseEvent DeployMenu
End Sub


Private Sub glistN_RegisterGlist(this As gList)
If Not glistN.Spinner Then this.NoWheel = True
On Error Resume Next
Set LastGlist = this
If Err.Number > 0 Then this.NoWheel = True



End Sub



Private Sub glistN_ChangeListItem(item As Long, content As String)
If Not enabled Then Exit Sub
Dim this As String, thispos As Long
thispos = glistN.SelStart - Len(mprompt)
this = Mid$(content, 1 + Len(mprompt))
RaiseEvent ValidString(this, thispos)
thispos = thispos + Len(mprompt)
content = mprompt + this
If thispos <> glistN.SelStart And this <> "" Then glistN.SelStart = thispos: glistN.ShowMe
End Sub
Public Sub ResetPan()
glistN.PanPos = 0
End Sub

Private Sub glistN_ChangeSelStart(thisselstart As Long)
If thisselstart <= Len(mprompt) Then thisselstart = Len(mprompt) + 1
End Sub

Public Property Get ThisKind() As String
ThisKind = mkind
End Property

Public Property Let ThisKind(ByVal RHS As String)
mkind = RHS
glistN.LastLinePart = " " & RHS
End Property


Public Property Get enabled() As Boolean
enabled = glistN.EditFlag
End Property

Public Property Let enabled(ByVal RHS As Boolean)
glistN.EditFlag = RHS
glistN.NoCaretShow = Not RHS
End Property
'


Private Sub glistN_CheckGotFocus()
If Not locked Then
If glistN.ListIndex < 0 And glistN.listcount > 0 Then glistN.ListindexPrivateUse = 0
RaiseEvent GotFocus
If enabled Then
glistN.FadeLastLinePart = mPartColor
glistN.forecolor = rgb(0, 0, 0)
glistN.backcolor = rgb(255, 255, 255)
glistN.NoCaretShow = False:
glistN.PanPos = 0
Else
glistN.FadeLastLinePart = 0
glistN.forecolor = rgb(255, 255, 255)
glistN.backcolor = rgb(100, 200, 160)
glistN.NoCaretShow = True
End If
glistN.ShowMe
End If

End Sub
Public Sub Retired()
glistN.FadeLastLinePart = 0
If Not mStayOpen Then
glistN.forecolor = rgb(255, 255, 255)
glistN.backcolor = rgb(200, 120, 60)
End If
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
End Sub

Private Sub glistN_CheckLostFocus()
If Not locked Then
If mStayOpen Then
glistN.NoCaretShow = True
glistN.FadeLastLinePart = 0
If glistN.Visible Then glistN.PanPos = 0: glistN.PrepareToShow
RaiseEvent LostFocus
Exit Sub
End If
If Not SkipProcAtExitFocus Then
glistN.FadeLastLinePart = 0
glistN.forecolor = rgb(255, 255, 255)
glistN.backcolor = rgb(200, 120, 60)
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
Else
'
End If
End If
RaiseEvent LostFocus
End Sub

Private Sub glistN_HeaderSelected(Button As Integer)
' INFO SELECTED
Text = Text
End Sub

Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)
RaiseEvent KeyDown(KeyCode, shift)
If KeyCode = vbKeyUp Then
ElseIf KeyCode = vbKeyDown Then
ElseIf KeyCode = vbKeyV And (shift And 3) = 2 Then
If Not enabled Then Exit Sub
Pad$ = GetTextData(13)
If Pad$ = vbNullString Then Pad$ = Clipboard.GetText(1)
If Pad$ <> "" Then vartext = Pad$
End If

End Sub

Public Property Get SelStart() As Long
SelStart = glistN.SelStart
End Property

Public Property Let SelStart(ByVal RHS As Long)
glistN.SelStart = RHS
End Property

Public Property Get Visible() As Boolean
Visible = glistN.Visible
End Property

Public Property Let Visible(ByVal RHS As Boolean)
glistN.Visible = RHS
End Property
Public Sub SetFocus()
On Error Resume Next
glistN.SetFocus
End Sub

Private Sub glistN_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, shift, x, y)
End Sub


Private Sub glistN_SyncKeyboard(item As Integer)
RaiseEvent KeyPress(item)
End Sub
Private Sub glistN_SpinnerValue(ThatValue As Long)
RaiseEvent SpinnerValue(ThatValue)
End Sub
Public Property Get locked() As Boolean
locked = Not glistN.enabled
End Property

Public Property Let locked(ByVal RHS As Boolean)
glistN.enabled = Not RHS
End Property
Public Sub Spinner(ShowBar As Boolean, MinValue As Long, maxvalue As Long, LargeValue As Long)
If Not ShowBar Then
glistN.Spinner = False
glistN.ShowBar = False
glistN.EnabledBar = False
Else
glistN.enabled = False
glistN.restrictLines = 1
glistN.Spinner = True
glistN.SetSpin MinValue, maxvalue, LargeValue
glistN.enabled = True
glistN.EnabledBar = True
glistN.ShowBar = True

End If
End Sub


Public Property Get Value() As Long
If glistN.Spinner Then
Value = glistN.ValueSilent
End If

End Property

Public Property Let Value(ByVal RHS As Long)
If glistN.Spinner Then
glistN.ValueSilent = RHS
End If

End Property

Public Sub Dereference()
glistN.Shutdown
Set glistN = Nothing
End Sub

Public Property Get MaxCharLength() As Long
MaxCharLength = glistN.maxchar
End Property

Public Property Let MaxCharLength(ByVal RHS As Long)
glistN.maxchar = RHS
End Property


Private Function GetTextData(ByVal lFormatId As Long) As String
'' Addition to myTextBox for ctrl+V
Dim bData() As Byte, sr As String, sr1 As String
sr1 = Clipboard.GetText(1)
If (OpenClipboard(0) <> 0) Then

        
        If (GetBinaryData(lFormatId, bData())) Then
        sr = bData
        
        If IsWine Then
        sr1 = Left$(sr, Len(sr1))
        GetTextData = Left$(sr1, Len(sr1))
        Else
        GetTextData = Left$(sr, Len(sr1))
        End If

        End If

End If
CloseClipboard
End Function
Private Function GetClipboardMemoryHandle( _
        ByVal lFormatId As Long _
    ) As Long

    

    If (IsClipboardFormatAvailable(lFormatId) <> 0) Then
         
        GetClipboardMemoryHandle = GetClipboardData(lFormatId)
        
    End If
End Function
Private Function GetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean

Dim hMem As Long, lSize As Long, lPtr As Long
    

    Erase bData
    
    hMem = GetClipboardMemoryHandle(lFormatId)

    If (hMem <> 0) Then

        lSize = GlobalSize(hMem)

        lPtr = GlobalLock(hMem)
        If (lSize > 0) Then

            ReDim bData(0 To lSize - 2) As Byte

            CopyMemory bData(0), ByVal lPtr, lSize - 1
        End If

        GlobalUnlock hMem

        GetBinaryData = (lSize > 0)

    End If
End Function

Private Sub glistN_SyncKeyboardUnicode(a As String)
RaiseEvent SyncKeyboardUnicode(a)
End Sub
Private Sub glistN_Fkey(a As Integer)
If a > 1000 Then
RaiseEvent Fkey(a - 1000)
Else
RaiseEvent Fkey(a)
End If
End Sub

Private Sub glistN_UnregisterGlist()
If Not glistN.Spinner Then glistN.NoWheel = True
On Error Resume Next
Set LastGlist = Nothing
If Err.Number > 0 Then glistN.NoWheel = True

End Sub

Public Property Let ShowAlways(ByVal vNewValue As Variant)
mStayOpen = CBool(vNewValue)
If mStayOpen Then
glistN.forecolor = rgb(0, 0, 0)
glistN.FadeLastLinePart = mPartColor
glistN.backcolor = rgb(255, 255, 255)
End If
End Property

Public Property Get FadePartColor() As Variant
FadePartColor = mPartColor
End Property

Public Property Let FadePartColor(ByVal vNewValue As Variant)
On Error Resume Next
mPartColor = CLng(vNewValue)
End Property
Private Sub glistN_AddSelStart(val As Long, shift As Integer)
Dim w As Long, a$
If glistN.SelStart > 0 Then
a$ = Mid$(vartext, glistN.SelStart, 1)
If Len(a$) > 0 Then
w = AscW(a$)
If w > -10241 And w < -9984 Then
val = 2
End If
End If
End If
End Sub

Private Sub glistN_SubSelStart(val As Long, shift As Integer)
Dim w As Long, a$
If glistN.SelStart < 2 Then Exit Sub
a$ = Mid$(vartext, glistN.SelStart - 1, 1)
If Len(a$) > 0 Then
w = AscW(a$)
If w > -9215 And w < -8960 Then

val = 2

End If
End If
End Sub
