VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiEditBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' M2000 ver 1.
Option Explicit
Implements IControlIndex
' chr(34) is """" maybe somewhere in a module as public
'Const mway As Long = DT_NOPREFIX Or DT_NOCLIP
Const mway As Long = DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP
Public textinform As String
Private Col() As Long, lastdata As String, thispara As Long
Private mark1 As Long, mark2 As Long

Private TextPixelOffset As Long, LastSearchType As Long
Private Brackets As Boolean
Private nochange As Boolean, sp As String, sP1 As String, sP2 As String
Private TabControl As Long
Private mNoCenterLineEdit As Boolean, LineCommandSep As String
Public PreserveRightSpaces As Boolean
Public HighLightParagraph As Boolean, HighLightColor As Long
Public UseHtmlColors As Boolean
Dim profundo As New clsProfiler
Public SeekNextLine As Boolean
Public ShadowMarks As Boolean
Public StartSymbols As String
Public OtherSymbols As String
Private mSplitExpr As String
Public LineComment1 As String
Public LineComment2 As String
Public AssignSym As String
Public StringSep As String
Public stringsep2 As String
Public ExtraFront As String
Public OpenTag As String, CloseTag As String, DropSym As String
Public CommentSymbols As String
Public ComSymbolsWidth As Long, CommentLineLight As Boolean
Public MultiLineComment1 As String
Private mMultiLineComment2 As String
Public Form1mn1Enabled As Boolean
Public Form1mn2Enabled As Boolean
Public Form1mn3Enabled As Boolean
Public Form1sdnEnabled As Boolean
Public Form1supEnabled As Boolean
Public Form1mscatEnabled As Boolean
Public Form1rthisEnabled As Boolean
Private mColorCollection1 As String
Private mUseCase As Boolean
Public ColorCollection2 As String
Public ColorCollection3 As String
Public ColorCollection4 As String
Private M2000code As Boolean
Public AllowInsertBrackets As Boolean
Public EnablePairs As Boolean
Public EditTextWord As Boolean
Public ColorSet As Long
Public NumberOnly As Boolean
Public NumberIntOnly As Boolean
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Private blockKeyboard As Boolean
Private LastDocTitle$, para1 As Long, PosPara1 As Long, Para2 As Long, PosPara2 As Long, Para3 As Long, PosPara3 As Long
Public fState As Long
Public WithEvents mDoc As Document    ' document
Attribute mDoc.VB_VarHelpID = -1
Dim WithEvents mDoc1 As Document  ' header
Attribute mDoc1.VB_VarHelpID = -1
Dim UndoDoc As New Document  ' one level of undo no events ' no break lines
Dim undopara As Long, undopos As Long, undotext As String, undomode As Long, undotexthtml As String
Dim Redopara As Long, Redopos As Long, Redotext As String, Redomode As Long
Dim undoprev As Boolean, Redoprev As Boolean
Dim undostate As Boolean
Dim Redostate As Boolean
Dim NoHeadwrap As Boolean
Public SkipProcAtExitFocus As Boolean
Private Type UndoChain
    undopara As Long
    undopos As Long
    undomode As Long
    undotext As String
    undoprev As Boolean
End Type
Dim NeoUndo As Long, UndoMax As Long
Dim NeoRedo As Long, RedoMax As Long

Dim MultiUndo() As UndoChain
Dim MultiRedo() As UndoChain

Dim GetUndoNext As Boolean
Dim myfilename As String
Dim all As Long, onelineheight As Long
''Event inform(tLine As Long, tPos As Long)
Public LastSelStart As Long
Public NoMark As Boolean
Public NoTitle As Boolean, TitleStay As Boolean
Private mNoColor As Boolean, mNoColor1 As Boolean
Public Title As String
Public UsedAsTextBox As Boolean  '' only in m2000. Enter exit - use maxchar - No wrap
Private mEditDoc As Boolean
Public AutoNumber As Boolean
Public Charpos As Long, expandCharPos As Long
Public showparagraph As Boolean
Private showparagraphWork As Boolean
Public SelectionColor As Long
Public SelectionTextColor As Long
Private showparagraphonce As Boolean
Private mNoWrap As Boolean
Private BlockStartPara As Long, BlockStartPos As Long, BlockStartLine As Long, BlockTopline As Long
Private BlockStartCursor As Long, BlockEndCursor As Long
Private BlockEndPara As Long, BlockEndPos As Long, BlockEndline As Long, BlockBottomline As Long
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_DISCARDED = &H4000
Private Const GMEM_FIXED = &H0
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GMEM_LOCKCOUNT = &HFF
Private Const GMEM_MODIFY = &H80
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_ZEROINIT = &H40
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
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&) ' not NT?
Private Const DWL_ANYTHREAD& = 0
Const LOCALE_ILANGUAGE = 1
    ' for Gui purposes
Private s$
Public ctrlName As String
Private mesCtrlName As String
Dim mIndex As Long, mStayOpen As Boolean
Dim Callback As GuiM2000
Dim oldEditDoc As Boolean
Dim white As Long, black As Long, dimm1 As Long, dimm2 As Long, capchroma As Long
Dim Quit
Private OldPar As Long, dragCharpos As Long
' new
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, RetVal As Integer)
Private Type DRAWTEXTPARAMS
     cbSize As Long
     iTabLength As Long
     iLeftMargin As Long
     iRightMargin As Long
     uiLengthDrawn As Long
End Type
Dim tParam As DRAWTEXTPARAMS
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Integer)
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (ByRef lpRect As RECT) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hDC As Long, ByVal lpsz As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long, ByVal lpDrawTextParams As Long) As Long
Private Declare Function GetStringTypeExW Lib "kernel32.dll" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByRef lpCharType As Integer) As Long
Private tmpshow As Boolean
Private WordOnly As Boolean
Private ColArr(0 To 15) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private LastListIndex As Long



Friend Property Get GetCallBack() As GuiM2000
Set GetCallBack = Callback
End Property
Friend Sub Construct(there As GuiM2000, Name$)
 mIndex = -1
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set glistN = there.Controls.Add("m2000.glist", ctrlName)
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
Friend Sub ConstructArray(there As GuiM2000, Name$, i As Long)
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set glistN = there.Controls.Add("m2000.glist", ctrlName)
 mIndex = i
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub

Public Sub SetUp()
TabControl = 4
EditTextWord = True

black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
' for vb6 coloring &HFFFFFF  - no checking for number of digits
ColorCollection4 = "|&H|&O|"
AssignSym = "="
ExtraFront = "&"
OtherSymbols = ",;"
LineComment1 = "\"
LineComment2 = "'"
StringSep = Chr$(34)
stringsep2 = Chr$(0)
CommentSymbols = "[\][\]"
ComSymbolsWidth = 2
CommentLineLight = False
MultiLineComment1 = "/*"
MultiLineComment2 = "*/"
Dim ff As New StdFont
With Callback
white = .Controls(1).ForeColor
ff.Name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With
mStayOpen = True
With glistN
WordCharLeft = " ,[]"
WordCharRight = " ,[]"
.NoEscapeKey = True
.bypassfirstClick = True
.Vertical = True
.DropEnabled = True
.DragEnabled = True
.PreserveNpixelsHeaderRightTwips = 8
NoMark = False
mNoColor = True
mNoColor1 = True
FileName = vbNullString
glistN.DragEnabled = False
.BackColor = rgb(200, 120, 60)
Set .Font = ff
.ForeColor = white
.dcolor = Callback.Controls(1).dcolor
.LeftMarginPixels = 8
End With
'glistN.addpixels = 0
TextPixelOffset = Int(glistN.TextHeightOffset)
capchroma = glistN.CapColor
EditDoc = True
nowrap = False
enabled = True
SetME
ShowAlways = True
mDoc.ForEditBox = True
End Sub
Sub Show()
If glistN Is Nothing Then Exit Sub
TextPixelOffset = Int(glistN.TextHeightOffset)
If tmpshow Then tmpshow = False: Exit Sub
glistN.PrepareToShow 5
End Sub
Public Sub PasteDoc(thisdoc As String)

 Const CF_UNICODETEXT = 13
 ClearRedo
If thisdoc = vbNullString Then
thisdoc = GetTextData(CF_UNICODETEXT)
Else
PushUndo
End If

undotext = thisdoc
undopara = mDoc.MarkParagraphID
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + CharPosStart - glistN.SelStart
Else
undopos = CharPosStart
End If
undomode = 0
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, (thisdoc)
'ReColorBlock
Render
End Sub
Private Function Myleft$(a As String, Pos As Long)
If Pos < 1 Then Myleft$ = vbNullString Else Myleft$ = Mid$(a, 1, Pos)
End Function
Private Function GetTextData(ByVal lFormatId As Long) As String
Dim bData() As Byte, sr As String, sr1 As String
On Error Resume Next
Sleep 10
sr1 = Clipboard.GetText(1)
If Err.Number = 521 Then GoTo cont1
If Err Then Err.Clear: Beep: Exit Function
cont1:
Err.Clear
If (OpenClipboard(Form1.hWnd) <> 0) Then

        
    If (GetBinaryData(lFormatId, bData())) Then
        sr = bData
        
        If IsWine Then
            sr1 = Myleft$(sr, Len(sr1))
            GetTextData = Myleft$(sr1, Len(sr1))
        Else
            GetTextData = Myleft$(sr, Len(sr1))
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


Public Sub HorSrcoll(widthtwips As Long)
glistN.PanPos = widthtwips
End Sub
Public Property Get FileName() As String
FileName = myfilename
End Property
Sub NewTitle(RHS As String, addpixelstoo As Long, Optional factor As Single = 1)
Title = RHS
Set mDoc1 = New Document
mDoc.NoIgnoreCase = Not mUseCase
If Title = vbNullString Then
Else
NoHeadwrap = False
mDoc1.textDoc = Title

glistN.VerticalCenterText = True
If mDoc1.DocLines >= glistN.lines + 1 Then
mDoc1.EmptyDoc
NoHeadwrap = True
mDoc1.textDoc = Title
End If
glistN.HeadLine = vbNullString ' reset
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight + addpixelstoo
glistN.addpixels = 4 * factor ' 'line spacing
End If
all = mDoc.DocLines
glistN.MultiLineEditBox = True
glistN.ShowBar = True
glistN.ShowBar = False  ' now is in autohide mode
glistN.StickBar = True
mDoc.WrapAgain
mDoc1.WrapAgain
glistN.PrepareToShow 10
all = mDoc.DocLines
glistN.CalcAndShowBar
End Sub
Public Property Let FileName(ByVal RHS As String)
myfilename = RHS
onelineheight = 0
Dim thisdoc As New Document, Header As New Document
Set mDoc = thisdoc
mDoc.NoIgnoreCase = Not mUseCase
mDoc.ColorEvent = True
Set mDoc1 = Header
para1 = 0
PosPara1 = 0
Para2 = 0
PosPara2 = 0
Para3 = 0
PosPara3 = 0
BlockStartPara = 0
BlockStartPos = 1
BlockStartLine = 0
BlockTopline = 0
BlockStartCursor = 0
BlockEndCursor = 0
BlockEndPara = 0
BlockEndPos = 1
BlockEndline = 0
BlockBottomline = 0
Dim i As Long
glistN.NoCaretShow = True
ClearRedo
ClearUndo
If RHS = vbNullString Then
mDoc.textDoc = vbNullString
Else
mDoc.textDoc = ReadUnicodeOrANSI(RHS)
End If

Dim what$, ThatPara As Long, pastchars
pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
NoTitle = Title = vbNullString
nowrap = True


If Not NoTitle Then
mDoc1.textDoc = Title

glistN.HeadLine = vbNullString ' reset
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight
End If
all = mDoc.DocLines
glistN.MultiLineEditBox = True
glistN.ShowBar = True
glistN.ShowBar = False  ' now is in autohide mode
glistN.ShowMe
glistN.StickBar = True
mDoc.WrapAgain
mDoc1.WrapAgain
glistN.PrepareToShow 10
all = mDoc.DocLines
glistN.CalcAndShowBar
End Property
Public Sub ResetSelColors()
SelectionTextColor = &HFFFFFF
SelectionColor = &H666666
End Sub



Private Sub Class_Initialize()
' new
Dim i As Long
sP1 = ",+-*/=<>^!.&[]{}()^;""" + "'!"
sp = " " + ChrW(9) + ChrW(160)
sP2 = sP1 + sp
tParam.cbSize = LenB(tParam)
tParam.iTabLength = 4
ColorSet = 1
M2000code = True
For i = 0 To 15: ColArr(CLng(i)) = M2000(i): Next i
M2000code = False
Set mDoc = New Document
mDoc.NoIgnoreCase = Not mUseCase
mDoc.ColorEvent = True
Set mDoc1 = New Document
ResetSelColors
mDoc.LCID = Clid
mDoc1.LCID = Clid
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
undopara = -1
Redopara = -1
mIndex = -1
mUseCase = True
LineCommandSep = ":"
mSplitExpr = ","
HighLightColor = rgb(225, 255, 255)
End Sub
Private Sub Class_Terminate()
Quit = True
Set glistN = Nothing
End Sub
Property Let TabWidth(RHS As Long)
    If RHS = 0 Then
        tParam.iTabLength = 4
        glistN.TabWidthChar = 4
    Else
        tParam.iTabLength = Abs(RHS)
        glistN.TabWidthChar = Abs(RHS)
    End If
    TabControl = 0
End Property
Property Let SpaceIndent(RHS As Long)
    TabControl = Abs(RHS)
End Property

Private Sub glistN_addone(that As String)  '
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 2
ElseIf that = vbCr Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = vbCrLf
undopos = SelStart
undomode = 1
ElseIf undomode = 1 And undotext <> "" Then
If undopos = SelStart Then
' we add after
undotext = undotext + that
ElseIf undopos = SelStart + 1 Then
    undotext = that + undotext
    undopos = SelStart
Else
GoTo there
End If
Else
there:
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart ' always the left side of char - so 0 is inside
undomode = 1
End If


End Sub
Public Sub AddUndo(that As String)
If Not enabled Then Exit Sub
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 2
Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart ' always the left side of char - so 0 is inside
undomode = 1
End If
End Sub
Public Sub RemoveUndo(that As String)
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 3
ElseIf that = vbCr Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 3
Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + SelStart - glistN.SelStart
Else
undopos = SelStart
End If
undomode = 0
End If
End Sub
Public Sub ManualInform()
If Not glistN.enabled Then Exit Sub
If glistN.SelStart = 0 Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Inform(" + str(mIndex) + "," + str(glistN.ListIndex + 1) + ", 1)"
   Else
      Callback.Callback mesCtrlName + ".Inform(" + str(glistN.ListIndex + 1) + ", 1)"
      End If
Else
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Inform(" + str(mIndex) + "," + str(glistN.ListIndex + 1) + ", " + str(glistN.SelStart) + ")"
   Else
      Callback.Callback mesCtrlName + ".Inform(" + str(glistN.ListIndex + 1) + ", " + str(glistN.SelStart) + ")"
      End If
End If
End Sub

Private Sub glistN_GroupUndo()
If undopara <> -1 Then
undoprev = True
End If
End Sub

Private Sub glistN_SubSelStart(val As Long, shift As Integer)
expandCharPos = 0
Dim W As Long, a$, l As Long, b$
If HaveMarkedText And Not (shift And 1) = 1 Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
val = 0
Exit Sub
End If

If Charpos < 2 Or NumberOnly Or NumberIntOnly Then Exit Sub

    a$ = Mid$(mDoc.TextParagraph(mDoc.MarkParagraphID), Charpos - 1)
    If Len(a$) > 0 Then
        W = AscW(a$)
        If (shift And 2) = 2 And Charpos > 1 Then
            a$ = Mid$(mDoc.TextParagraph(mDoc.MarkParagraphID), 1, Charpos - 1)
            l = Len(a$)
            val = l
            If val = 1 Then Exit Sub
            b$ = Mid$(a$, l, 1)
            If InStr(sP1, b$) = 0 Then
                Do While l > 1 And InStr(sp, b$) > 0
                    l = l - 1
                    b$ = Mid$(a$, l, 1)
                Loop
                If InStr(sP1, b$) = 0 Then
                    If l = val Then
                        Do While l > 1 And InStr(sP2, b$) = 0
                            l = l - 1
                            b$ = Mid$(a$, l, 1)
                        Loop
                        If l = 1 And Not InStr(sP2, b$) > 0 Then l = l - 1
                        
                    Else
                     If InStr(sP2, b$) > 0 Then l = l - 1
                    End If
                End If
            Else
                l = l - 1
            End If
            val = val - l
        ElseIf W > -9217 And W < -8192 Then
            val = 2
        End If
    ElseIf Charpos > 0 Then
    W = Len(mDoc.TextParagraph(mDoc.MarkParagraphID))
    If W < Charpos Then val = Charpos - W
    End If
    
End Sub

Private Sub glistN_AddSelStart(val As Long, shift As Integer)
expandCharPos = 0
If NumberOnly Or NumberIntOnly Then Exit Sub
Dim W As Long, a$, l As Long, b$
If HaveMarkedText And Not (shift And 1) = 1 Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
val = 0
Exit Sub
End If
If Charpos > 0 Then
    a$ = Mid$(mDoc.TextParagraph(mDoc.MarkParagraphID), Charpos, 1)
    If Len(a$) > 0 Then
        W = AscW(a$)
        If (shift And 2) = 2 Then
            a$ = Mid$(mDoc.TextParagraph(mDoc.MarkParagraphID), Charpos)
            l = Len(a$)
            val = 1
            If l < 2 Then Exit Sub
                b$ = Mid$(a$, val, 1)
                If InStr(sP1, b$) = 0 Then
                Do While val < l And InStr(sp, b$) > 0
                    val = val + 1
                    b$ = Mid$(a$, val, 1)
                Loop
                If InStr(sP1, b$) = 0 Then
                    If val = 1 Then
                        Do While val < l And InStr(sP2, b$) = 0
                            val = val + 1
                            b$ = Mid$(a$, val, 1)
                        Loop
                        If InStr(sP2, b$) > 0 Then val = val - 1
                    Else
                        If val < l Then val = val - 1
                    End If
                Else
                    val = val - 1
                End If
            End If
        ElseIf W > -10241 And W < -9216 Then
            val = 2
        End If
    End If
End If



End Sub
Private Sub glistN_CheckGotFocus()
Callback.RealHover = ctrlName
If Not locked Then
glistN.HideCaretOnexit = False
If glistN.ListIndex < 0 And glistN.listcount > 0 Then glistN.ListindexPrivateUse = 0

If enabled Then
glistN.NoCaretShow = NoMark

mNoColor = mNoColor1
glistN.ForeColor = black
glistN.BackColor = white

Else
mNoColor = True
glistN.NoCaretShow = NoMark
glistN.ForeColor = white
glistN.BackColor = dimm1
End If
If Not NoCenterLineEdit Then
glistN.ShowMe
End If
Else
mNoColor = True
End If

End Sub
Private Sub glistN_CheckLostFocus()
If Not locked Then
If mStayOpen And Not SkipProcAtExitFocus Then
glistN.NoCaretShow = NoMark
glistN.HideCaretOnexit = True
If glistN.Visible Then
If NumberOnly Or UsedAsTextBox Then
SelStartSilent = 0
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End If
If Not NoCenterLineEdit Then
glistN.ShowPan
End If
End If

Exit Sub
End If
If Not SkipProcAtExitFocus Then
mNoColor = True
glistN.NoCaretShow = NoMark
glistN.HideCaretOnexit = True
glistN.ForeColor = white
If glistN.BackStyle <> 1 Then

glistN.BackColor = dimm2
End If

glistN.PanPos = 0: glistN.PrepareToShow

Else
'
End If

Else

End If

'RaiseEvent LostFocus
End Sub

Private Sub glistN_CorrectCursorAfterDrag()
    If Not mDoc.InvalidPara(OldPar) Then
        SelLengthSilent = 0
        mDoc.MarkParagraphID = OldPar
        glistN.enabled = False
        ParaSelStart = dragCharpos
        glistN.enabled = True
        ManualInform
        OldPar = 0
    End If
End Sub
Private Sub glistN_DragOverCursor(ok As Boolean)
ok = True
End Sub

Private Sub glistN_getpair(a As String, b As String)
If EnablePairs Then
Select Case a
Case StringSep
b = StringSep
Case Myleft$(stringsep2, 1)
b = Right$(stringsep2, 1)
Case "("
b = ")"
Case "{"
b = "}"
Case "["
b = "]"
Case Else
If Len(OpenTag) > 0 Then
If a = OpenTag Then b = ">"

End If
End Select
End If
End Sub

Private Sub glistN_HaveMark(Yes As Boolean)
Yes = HaveMarkedText
End Sub

Private Sub glistN_MarkCut(ThatData As String)
expandCharPos = 0
ThatData = undotext
End Sub
Private Sub glistN_DragOverDone(a As Boolean)
If Not mDoc.InvalidPara(OldPar) Then
        a = True
        SelLengthSilent = 0
        mDoc.MarkParagraphID = OldPar
        glistN.enabled = False
        ParaSelStart = dragCharpos
        glistN.enabled = True
        ManualInform
        OldPar = 0
    End If
End Sub
Private Sub glistN_DragOverNow(a As Boolean)
If OldPar = 0 Then
a = True
 SelLengthSilent = 0
 OldPar = mDoc.MarkParagraphID
 dragCharpos = Charpos
End If
End Sub
Private Sub glistN_Maybelanguage()
If Not glistN.enabled Then Exit Sub
If glistN.SelStart = 0 Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Inform(" + str(mIndex) + "," + str(glistN.ListIndex + 1) + ", 1)"
   Else
      Callback.Callback mesCtrlName + ".Inform(" + str(glistN.ListIndex + 1) + ", 1)"
      End If
Else
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Inform(" + str(mIndex) + "," + str(glistN.ListIndex + 1) + ", " + str(glistN.SelStart) + ")"
   Else
      Callback.Callback mesCtrlName + ".Inform(" + str(glistN.ListIndex + 1) + ", " + str(glistN.SelStart) + ")"
      End If
End If
Render
End Sub


Private Sub glistN_OutPopUp(X As Single, Y As Single, myButton As Integer)
'
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".PopUp(" + str(mIndex) + "," + str(X + Callback.Controls(ctrlName).Left) + "," + str(Y + Callback.Controls(ctrlName).top) + "," + str(myButton) + "," + str(glistN.ListIndex + 1) + ")"
   Else
      Callback.Callback mesCtrlName + ".PopUp(" + str(X + Callback.Controls(ctrlName).Left) + "," + str(Y + Callback.Controls(ctrlName).top) + "," + str(myButton) + "," + str(glistN.ListIndex + 1) + ")"
      End If
End Sub

Private Sub glistN_PureListOff()
showparagraphWork = True
End Sub

Private Sub glistN_PureListOn()
showparagraphWork = False
End Sub

Private Sub glistN_PushMark2Undo(that As String)
ClearRedo
undopara = BlockEndPara
undotext = that
undomode = 1
undopos = SelStart
End Sub

Private Sub glistN_PushUndoIfMarked()
If HaveMarkedText Then PushUndo
End Sub



Private Sub glistN_RealCurReplace(a As String)
If showparagraph Then
a = Replace(a, vbTab, ChrW(&H21E5))
End If
End Sub

Private Sub glistN_RemoveOne(that As String)
' push undo removeone
If Not enabled Then Exit Sub

ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 3
ElseIf undomode = 0 And undotext <> "" Then
If undopos + Len(undotext) = SelStart Then
' we add after
undotext = undotext + that

Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + SelStart - glistN.SelStart
Else
If glistN.ListIndex = all Then
undopos = SelStart - 1
Else
undopos = SelStart
End If
End If
 ' always the left side of char - so 0 is inside
undomode = 0
End If

Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + SelStart - glistN.SelStart
Else
undopos = SelStart
End If
undomode = 0
End If
End Sub


Private Sub glistN_MayRefresh(ok As Boolean)
ok = Not HaveMarkedText
End Sub

Public Property Get MaxCharLength() As Long
MaxCharLength = glistN.maxchar
End Property

Public Property Let MaxCharLength(ByVal RHS As Long)
glistN.maxchar = RHS
End Property
Private Sub glistN_ChangeListItem(item As Long, content As String)
Dim selstart2 As Long, checklistindex As Long, old$
'' only for M2000, use NumberOnly
If NumberOnly Or NumberIntOnly Then
content = MyTrim(content)
NumberOnly = True
If Len(content) > glistN.maxchar Then
Charpos = glistN.maxchar
glistN.SelStartEventAlways = glistN.SelStart - 1
Exit Sub
End If
If content = vbNullString Then
ElseIf Not ValidNum(content, False, NumberIntOnly) Then
undotext = vbNullString
content = mDoc.TextLine(item + 1)
Exit Sub
End If
End If
If UsedAsTextBox Then
If Len(content) > glistN.maxchar Then
IsRTLrun
Charpos = glistN.maxchar
glistN.SelStartEventAlways = glistN.SelStart - 1
Exit Sub
End If
End If
mDoc.TextLine(item + 1) = content
'IsRTLrun
WrapMarkedPara
If all <> mDoc.DocLines Then
    Render
Else
    checklistindex = mDoc.FirstParagraphLine(mDoc.MarkParagraphID) + mDoc.TextParagraph2Line(mDoc.MarkParagraphID, Charpos, selstart2)
    If (checklistindex - glistN.ScrollFrom) >= glistN.lines + 1 Then
     glistN.ListindexPrivateUse = checklistindex
     glistN.ScrollToTextEdit checklistindex - glistN.lines
     glistN.SelStartEventAlways = selstart2
     

    
    ElseIf checklistindex < glistN.ScrollFrom Then
     glistN.ListindexPrivateUse = checklistindex
     glistN.ScrollToTextEdit checklistindex
          glistN.SelStartEventAlways = selstart2
    Else
    glistN.ListindexPrivateUse = checklistindex
    glistN.SelStart = selstart2
        End If
End If

End Sub

Private Sub glistN_ChangeSelStart(thisselstart As Long)
If glistN.ListIndex < 0 Then Exit Sub
Dim what$, ThatPara As Long, pastchars As Long, oldpanpos As Long
pastchars = mDoc.TextParagraphLine2(glistN.ListIndex + 1, what$, ThatPara, False)
Charpos = pastchars + glistN.SelStart
If Not glistN.enabled Then Exit Sub
oldpanpos = glistN.PanPos
tmpshow = True
If glistN.SelStart = 0 Then
inform glistN.ListIndex + 1, 1
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Inform(" + str(mIndex) + "," + str(glistN.ListIndex + 1) + ", 1)"
   Else
      Callback.Callback mesCtrlName + ".Inform(" + str(glistN.ListIndex + 1) + ", 1)"
      End If
Else
 inform glistN.ListIndex + 1, (glistN.SelStart)
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Inform(" + str(mIndex) + "," + str(glistN.ListIndex + 1) + ", " + str(glistN.SelStart) + ")"
   Else
      Callback.Callback mesCtrlName + ".Inform(" + str(glistN.ListIndex + 1) + ", " + str(glistN.SelStart) + ")"
      End If
End If
On Error Resume Next

If LastListIndex <> glistN.ListIndex Then

LastListIndex = glistN.ListIndex

    If NoCenterLineEdit Then
       If Not tmpshow Then glistN.ShowMe2
    Else
    tmpshow = False
    If oldpanpos <> glistN.PanPos Then Show Else glistN.ShowMe
    End If
Else
If NoCenterLineEdit Then
 'glistN.ShowPan
Else
Show
End If
    
End If
tmpshow = False
Exit Sub
If NoCenterLineEdit Then
    If LastListIndex <> glistN.ListIndex Then
        LastListIndex = glistN.ListIndex:  glistN.ShowMe: If Not tmpshow Then glistN.ShowMe2
    Else
        If Not tmpshow Then glistN.ShowMe2
    End If

Else
    glistN.ShowMe
    If Not tmpshow Then glistN.ShowMe2
End If

End Sub

Private Sub glistN_DragData(ThatData As String)
FillBlock
ThatData = UndoDoc.textDoc
 OldPar = mDoc.MarkParagraphID
 dragCharpos = Charpos
End Sub

Private Sub glistN_DragPasteData(ThatData As String)
Dim a() As String, K As Long, a77$
a77$ = space$(EditTabWidth)
If Not glistN.UseTab Then

a() = Split(ThatData, ChrW(9))
PasteDoc Join(a(), a77$)
Else
PasteDoc ThatData
End If
RelocateMarkedText

End Sub

Private Sub glistN_DropFront(ok As Boolean)
ok = False
' we need to know if we delete mark before paste data
' when we perform dragdrop to the same document
' by defult ok=false so we drop behind so we can delete marked text and then we paste
If HaveMarkedText Then
If glistN.ListIndex < BlockStartLine Then
' we are behind
ElseIf glistN.ListIndex = BlockStartLine Then
If glistN.SelStart < BlockStartCursor Then
Else
ok = True
End If
Else
ok = True
End If
End If
End Sub

Private Sub glistN_DropOk(ok As Boolean)
If HaveMarkedText Then
If ClickInsideMarkedText Then
Else
ok = True
End If
Else
ok = True
End If
End Sub

Private Sub glistN_ExposeListcount(cListCount As Long)
cListCount = all
End Sub

Private Sub gListN_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim a As RECT, b As RECT, i As Long, tmp As Long
Dim Final$, fbypass As Boolean
Dim bb As RECT, cc As RECT
Dim OldColor As Long
Dim what$, ThatPara As Long
CopyFromLParamToRect a, thisrect
b = a
bb = a
cc = a
cc.Left = 0
If UsedAsTextBox Then
b.Bottom = a.Bottom - 2
End If
Dim skipcolor As Boolean
OldColor = glistN.ForeColor
If item < 0 Then  ' this is a headline
    FillBack thisHDC, b, glistN.CapColor
    mDoc1.TextParagraphLine2 1, what$, ThatPara, False
    showparagraphonce = False
    glistN.HeadlineHeight = mDoc1.DocLines * onelineheight
    For i = 1 To mDoc1.DocLines
        PrintItem2 thisHDC, mDoc1.TextLine2(i, False, True), a
        a.top = a.top + onelineheight
    Next i
    'RaiseEvent ExposeRect(item, thisrect, thisHDC, skip)
    skip = True
    Exit Sub
Else
    showparagraphWork = showparagraph
    a.top = a.top + glistN.addpixels \ 2
    mDoc.TextParagraphLine2 item + 1, what$, ThatPara, showparagraph
    If MarkedTextAny Then
        showparagraphWork = True
        If item > BlockTopline And item < BlockBottomline Then
            FillBack thisHDC, b, SelectionColor
            SetTextColor thisHDC, SelectionTextColor
            skipcolor = True
        ElseIf item = BlockTopline And item <> BlockBottomline Then
            If BlockEndline < BlockStartLine Then
                If BlockEndCursor > 1 Then  ' if not leave it as is
                    b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1))
                    PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                    Final$ = Mid$(glistN.list(item), BlockEndCursor)
                    fbypass = True
                End If
            Else
                If BlockStartCursor > 1 Then  ' if not leave it as is
                    b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1))
                    PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                    Final$ = Mid$(glistN.list(item), BlockStartCursor)
                    fbypass = True
                 End If
            End If
            FillBack thisHDC, b, SelectionColor
            SetTextColor thisHDC, SelectionTextColor
            skipcolor = True
        ElseIf item = BlockBottomline And item <> BlockTopline Then
            If BlockBottomline = BlockEndline Then
                If BlockEndCursor <= Len(glistN.list(item)) Then
                    If BlockEndCursor > 1 Then ' if not leave it as is
                        b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                        PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                        Final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
                        fbypass = False
                    Else
                        If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor2 RealX1(thisHDC, glistN.SelStart, mDoc.TextLine2(item + 1, showparagraph, NoColor)), OldColor
                        GoTo conthere
                    End If
                End If
            Else
                If BlockStartCursor <= Len(glistN.list(item)) Then  ' if not leave it as is
                    If BlockStartCursor > 1 Then
                        b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                        PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                        Final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
                        fbypass = False
                    Else
                        If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor2 RealX1(thisHDC, glistN.SelStart, mDoc.TextLine2(item + 1, showparagraph, NoColor)), OldColor
                        GoTo conthere
                    End If
                End If
            End If
            FillBack thisHDC, b, SelectionColor
            a.Right = b.Right + 1
            SetTextColor thisHDC, SelectionTextColor
            skipcolor = True
        ElseIf BlockBottomline = item Then ' in one line only
            skipcolor = False
            If BlockEndCursor > BlockStartCursor Then
                If BlockStartCursor > 0 Then b.Left = b.Left + RealX1(thisHDC, BlockStartCursor, mDoc.TextLine2(item + 1, showparagraph, NoColor))
                If BlockEndCursor > 0 Then b.Right = RealX1(thisHDC, BlockEndCursor, mDoc.TextLine2(item + 1, showparagraph, NoColor)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                If b.Left > b.Right Then
                    tmp = b.Right
                    b.Right = b.Left
                    b.Left = tmp
                End If
                If BlockStartCursor = 0 Then
                    Final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
                Else
                    skipcolor = False
                    Final$ = Mid$(glistN.list(item), BlockStartCursor, BlockEndCursor - BlockStartCursor)
                End If
                PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                FillBack thisHDC, b, SelectionColor
                a.Right = b.Right + 1
                SetTextColor thisHDC, SelectionTextColor
            ElseIf BlockEndCursor < BlockStartCursor Then
                If BlockEndCursor > 0 Then b.Left = b.Left + RealX1(thisHDC, BlockEndCursor, mDoc.TextLine2(item + 1, showparagraph, NoColor))
                If BlockStartCursor > 0 Then b.Right = RealX1(thisHDC, BlockStartCursor, mDoc.TextLine2(item + 1, showparagraph, NoColor)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                If b.Left > b.Right Then
                    tmp = b.Right
                    b.Right = b.Left
                    b.Left = tmp
                End If
                If BlockEndCursor > 0 Then
                    Final$ = Mid$(glistN.list(item), BlockEndCursor, BlockStartCursor - BlockEndCursor)
                Else
                    Final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
                End If
                PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                FillBack thisHDC, b, SelectionColor
                a.Right = b.Right + 1
                SetTextColor thisHDC, SelectionTextColor
            End If
            fbypass = False
        End If
    ElseIf HighLightParagraph Then
       If ThatPara = mDoc.MarkParagraphID Then
        bb = cc
        bb.top = b.top
        bb.Bottom = b.Bottom
        FillBack thisHDC, bb, HighLightColor
        End If
    End If
    If ThatPara = mDoc.MarkParagraphID Then
        If EditDoc Then
            If glistN.ListIndex = item Then
            ' why???
                If mNoWrap Or (Not skipcolor) Then glistN.ExternalCursor2 RealX1(thisHDC, glistN.SelStart, mDoc.TextLine2(item + 1, showparagraph, NoColor)), OldColor
            End If
        ElseIf Not NoMark Then
            FillBack thisHDC, b, 0
            SetTextColor thisHDC, SelectionTextColor
        End If
    ElseIf Not skipcolor Then
        SetTextColor thisHDC, glistN.ForeColor
    End If
    If Final$ <> "" Or fbypass Then
        If Not fbypass Then

            PrintItem2X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, b, mDoc.lastcolordata(), False
        Else

             PrintItem2X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, b, mDoc.lastcolordata(), True
        End If
        showparagraphWork = False
    Else
conthere:
        If skipcolor Or NoColor Then
            PrintItem3X thisHDC, mDoc.TextLine2(item + 1, showparagraph, True), a, mDoc.lastcolordata()
        Else
            PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
        End If
    End If
End If
SetTextColor thisHDC, OldColor
showparagraphonce = False
skip = True
End Sub


Private Sub gListN_ExposeRect2(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim a As RECT, b As RECT, i As Long, pastchars As Long, hRgn As Long, tmp As Long
Dim bb As RECT, cc As RECT
Dim Final$, fbypass As Boolean
Dim OldColor As Long
Dim what$, ThatPara As Long
CopyFromLParamToRect a, thisrect
b = a
bb = a
cc = a
cc.Left = 0
If UsedAsTextBox Then
b.Bottom = a.Bottom - 2
End If
Dim skipcolor As Boolean
OldColor = glistN.ForeColor
If item < 0 Then  ' this is a headline
    FillBack thisHDC, b, glistN.CapColor
    mDoc1.TextParagraphLine2 1, what$, ThatPara, False
    showparagraphonce = False
    glistN.HeadlineHeight = mDoc1.DocLines * onelineheight
    For i = 1 To mDoc1.DocLines
        PrintItem2 thisHDC, mDoc1.TextLine2(i, False, True), a
        a.top = a.top + onelineheight
    Next i
    ' RaiseEvent ExposeRect(item, thisrect, thisHDC, skip)
    skip = True
    Exit Sub
Else
    showparagraphWork = showparagraph
    a.top = a.top + glistN.addpixels \ 2
    mDoc.TextParagraphLine2 item + 1, what$, ThatPara, showparagraph
    If MarkedTextAny Then
        showparagraphWork = True
        If item > BlockTopline And item < BlockBottomline Then
            FillBack thisHDC, b, SelectionColor
            SetTextColor thisHDC, SelectionTextColor
            skipcolor = True
        ElseIf item = BlockTopline And item <> BlockBottomline Then
            If BlockEndline < BlockStartLine Then
                If BlockEndCursor > 1 Then  ' if not leave it as is
                    b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1))
                    PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                    Final$ = Mid$(glistN.list(item), BlockEndCursor)
                    fbypass = True
                 End If
             Else
                  If BlockStartCursor > 1 Then  ' if not leave it as is
                 b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1))
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
              
                 Final$ = Mid$(glistN.list(item), BlockStartCursor)
                 fbypass = True
                 End If
             End If
     
             FillBack thisHDC, b, SelectionColor
             'a.Left = b.Left + 1
             
         SetTextColor thisHDC, SelectionTextColor
         skipcolor = True
     ElseIf item = BlockBottomline And item <> BlockTopline Then
      
             If BlockBottomline = BlockEndline Then
                 If BlockEndCursor <= Len(glistN.list(item)) Then
                 If BlockEndCursor > 1 Then ' if not leave it as is
                    b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                    PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                    Final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
                    fbypass = False
                 Else
                    If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor glistN.SelStart, what$, black
                    GoTo conthere
                 End If
                 End If
                 
             Else
                  If BlockStartCursor <= Len(glistN.list(item)) Then  ' if not leave it as is
               If BlockStartCursor > 1 Then
               b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                Final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
               fbypass = False
                     Else
                       If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor glistN.SelStart, what$, black
                      GoTo conthere
                 End If
                 End If
                 
             End If
     
             FillBack thisHDC, b, SelectionColor
             a.Right = b.Right + 1
             
         SetTextColor thisHDC, SelectionTextColor
         skipcolor = True
    ElseIf BlockBottomline = item Then ' in one line only
        skipcolor = False
        If BlockEndCursor > BlockStartCursor Then
            If BlockStartCursor > 1 Then b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1))
            If BlockEndCursor > 1 Then b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
            If BlockStartCursor = 0 Then
                Final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
            Else
                skipcolor = False
                Final$ = Mid$(glistN.list(item), BlockStartCursor, BlockEndCursor - BlockStartCursor)
            End If
            PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
            FillBack thisHDC, b, SelectionColor
            a.Right = b.Right + 1
            SetTextColor thisHDC, SelectionTextColor
        ElseIf BlockEndCursor < BlockStartCursor Then
            If BlockEndCursor > 1 Then b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1))
            If BlockStartCursor > 1 Then b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
            If BlockEndCursor > 0 Then
                Final$ = Mid$(glistN.list(item), BlockEndCursor, BlockStartCursor - BlockEndCursor)
            Else
                Final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
            End If
            PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                FillBack thisHDC, b, SelectionColor
                a.Right = b.Right + 1
                SetTextColor thisHDC, SelectionTextColor
            End If
            fbypass = False
        End If
 ''  skipcolor = True
ElseIf HighLightParagraph Then
   If ThatPara = mDoc.MarkParagraphID Then
    bb = cc
    bb.top = b.top
    bb.Bottom = b.Bottom
    FillBack thisHDC, bb, HighLightColor
    End If
End If
    If ThatPara = mDoc.MarkParagraphID Then

            If EditDoc Then
                        ' old patch -- removed here
                      If mNoWrap And (glistN.ListIndex = item) Then
              '        glistN.ExternalCursor glistN.SelStart, what$, black
                       ElseIf (Not skipcolor) And (glistN.ListIndex = item) Then
                      If OldPar > 0 Then glistN.ExternalCursor glistN.SelStart, what$, black
                       End If
            ElseIf Not NoMark Then
            
                    FillBack thisHDC, b, 0
                     SetTextColor thisHDC, SelectionTextColor
                 
                     
            End If
    
    Else
         
    If Not skipcolor Then SetTextColor thisHDC, glistN.ForeColor
 

    End If
    If Final$ <> "" Or fbypass Then

   If Not fbypass Then
 
   'PrintItem thisHDC, final$, a
   hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn thisHDC, hRgn
   PrintItem thisHDC, mDoc.TextLine2(item + 1, showparagraph, False), a
   SelectClipRgn thisHDC, &H0
   DeleteObject hRgn
   Else
'   PrintItem1 thisHDC, final$, a
hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn thisHDC, hRgn
   PrintItem1 thisHDC, mDoc.TextLine2(item + 1, showparagraph, False), a
   SelectClipRgn thisHDC, &H0
   DeleteObject hRgn
   End If
   showparagraphWork = False
    Else
conthere:
If skipcolor Then
 PrintItem1 thisHDC, mDoc.TextLine2(item + 1, showparagraph, False), a
Else
  PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
End If
  End If
End If
SetTextColor thisHDC, OldColor
 showparagraphonce = False

skip = True
End Sub



Public Sub Render()
If mDoc.Busy Then Exit Sub

Dim selstart2 As Long
mDoc.Busy = True
all = mDoc.DocLines
glistN.ListIndex = mDoc.FirstParagraphLine(mDoc.MarkParagraphID) + mDoc.TextParagraph2Line(mDoc.MarkParagraphID, Charpos, selstart2, NoColor)
glistN.SelStart = selstart2
'GoTo aaa
If glistN.ListIndex < glistN.ScrollFrom Then
If glistN.ListIndex > glistN.lines Then
glistN.ScrollToSilent glistN.ListIndex - glistN.lines / 2
Else
glistN.ScrollToSilent 0
End If
'ElseIf glistN.ListIndex > glistN.lines And glistN.listcount - glistN.ListIndex <= glistN.lines Then
'    glistN.ScrollToSilent glistN.listcount - glistN.lines - 1
ElseIf glistN.ListIndex > glistN.ScrollFrom + glistN.lines Then
If glistN.listcount - glistN.ListIndex < glistN.lines Then
glistN.ScrollToSilent glistN.listcount - glistN.lines
Else
glistN.ScrollToSilent glistN.ListIndex - glistN.lines / 2
End If
ElseIf glistN.ListIndex > glistN.lines And glistN.listcount - glistN.ListIndex < glistN.lines Then
If glistN.enabled Then glistN.ScrollToSilent glistN.listcount - glistN.lines
End If
aaa:

glistN.ScrollToSilent glistN.ScrollFrom

If Not mNoCenterLineEdit Then
If glistN.enabled Then glistN.ShowPan
Else
If glistN.enabled Then glistN.ShowMe2
End If
'expandCharPos = glistN.SelStart
mDoc.Busy = False
End Sub

Public Sub WrapMarkedPara()
If glistN.ListIndex >= 0 Then
mDoc.ReWritePara mDoc.MarkParagraphID, mDoc.TextParagraph(mDoc.MarkParagraphID)
End If
End Sub

Public Sub CloseTitle()
If glistN.HeadLine <> "" Then
If NoTitle Then
glistN.HeadLine = vbNullString
onelineheight = 0
glistN.HeadlineHeight = 0
ElseIf Title <> glistN.HeadLine And Not TitleStay Then
ReplaceTitle = vbNullString
End If
End If
End Sub
Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)
If glistN.HeadLine <> "" Then
If NoTitle Then
glistN.HeadLine = vbNullString
onelineheight = 0
glistN.HeadlineHeight = 0
ElseIf Title <> glistN.HeadLine And Not TitleStay Then
ReplaceTitle = vbNullString
End If
End If
Dim VR(2)
VR(0) = KeyCode
VR(1) = shift
If mIndex > -1 Then
    Callback.CallbackNow mesCtrlName + ".KeyDown(" + CStr(Index) + ")", VR()
Else
    Callback.CallbackNow mesCtrlName + ".KeyDown()", VR()
End If
shift = VR(1)
KeyCode = VR(0)

Dim selstart2 As Long, b As Boolean, Pad$, cc As Long
Dim l1 As Long, c1 As Long
If KeyCode = 18 Then
Exit Sub
End If
If KeyCode = vbKeyF10 And (shift And 7) = 1 Then
KeyCode = 0
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".PopUp(" + str(mIndex) + "," + str(MOUSEX(Callback.Left)) + "," + str(MOUSEY(Callback.top)) + "," + str(2) + "," + str(glistN.ListIndex + 1) + ")"
   Else
      Callback.Callback mesCtrlName + ".PopUp(" + str(MOUSEX(Callback.Left)) + "," + str(MOUSEY(Callback.top)) + "," + str(2) + "," + str(glistN.ListIndex + 1) + ")"
      End If
Exit Sub
End If
If KeyCode = 16 And shift <> 0 Then
KeyCode = 0
If glistN.SelStart = 0 Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Inform(" + str(mIndex) + "," + str(glistN.ListIndex + 1) + ", 1)"
   Else
      Callback.Callback mesCtrlName + ".Inform(" + str(glistN.ListIndex + 1) + ", 1)"
      End If
Else
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Inform(" + str(mIndex) + "," + str(glistN.ListIndex + 1) + ", " + str(glistN.SelStart) + ")"
   Else
      Callback.Callback mesCtrlName + ".Inform(" + str(glistN.ListIndex + 1) + ", " + str(glistN.SelStart) + ")"
      End If
End If
Exit Sub
End If
If NoMark And Not (KeyCode <> vbKeyDown Or KeyCode <> vbKeyUp Or KeyCode <> vbKeyPageUp Or KeyCode <> vbKeyPageDown) Then KeyCode = 0: Exit Sub

If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
'
If NoMark And KeyCode = vbKeyUp Then KeyCode = vbKeyPageUp
If NoMark And KeyCode = vbKeyDown Then KeyCode = vbKeyPageDown
ElseIf KeyCode = vbKeyX And (shift And 3) = 2 Then
MarkCut
Show
KeyCode = 0
ElseIf KeyCode = vbKeyC And (shift And 3) = 2 Then
MarkCopy
KeyCode = 0
ElseIf KeyCode = vbKeyV And (shift And 3) = 2 Then
    Sleep 10
    On Error Resume Next
    Pad$ = GetTextData(13)
    If Err Or Pad$ = "" Then
    Pad$ = Clipboard.GetText(1)
    If Err Or Pad$ = "" Then Err.Clear: KeyCode = 0: Exit Sub
    End If
   
    If NumberOnly Or NumberIntOnly Then
        If Not ValidNum(Pad$, False, NumberIntOnly) Then GoTo conthere
        NumberOnly = True
        If Not mNoWrap = True Then
            mNoWrap = True
            glistN.NoScroll = False
        End If
    ElseIf UsedAsTextBox Then
        Pad$ = Replace(Pad$, Chr(13), "")
        Pad$ = Replace(Pad$, Chr(10), "")
        If Not mNoWrap = True Then
            mNoWrap = True
            glistN.NoScroll = False
        End If
    End If
  '  If Len(CurrentParagraph) + 1 < Charpos Then
  '      SelStartSilent = CharPosStart - Charpos + 1 + Len(CurrentParagraph)
  '  End If
    If ParaSelStart = 2 And glistN.list(glistN.ListIndex) = vbNullString Then
        SelStart = SelStart - 1
    End If
    AddUndo ""
    If UsedAsTextBox Then
    Pad$ = GetNextLine((Pad$))
    If Len(Text) + Len(Pad$) > glistN.maxchar Then
        Pad$ = Myleft$(Pad$, glistN.maxchar - Len(Text))
    End If
    SelText = Pad$
    RemoveUndo SelText
    ElseIf NumberOnly Or NumberIntOnly Then
    glistN.MarkNext = 0
    glistN.MarkALL
    SelText = Pad$
    glistN.MarkNext = 0
    glistN.MarkALL
    RemoveUndo SelText
    SelLength = 0
    Else
    SelText = Pad$
    RemoveUndo SelText
    End If
    KeyCode = 0
    ReColorBlock
ElseIf KeyCode = vbKeyA And (shift And 3) = 2 Then
glistN.MarkNext = 0
glistN.MarkALL
KeyCode = 0
ElseIf KeyCode = vbKeyY And (shift And 3) = 2 Then
' Redo
profundo.MARKONE
againY:
If profundo.MARKTWO > 1000 Then ProcTask2 Basestack1: profundo.MARKONE
If Redopara <> -1 Then

PushUndo

undopara = Redopara
undopos = Redopos
undotext = Redotext

undomode = Redomode
If undomode = 2 Then undopos = Redopos - 2
undoprev = Redoprev
b = enabled
enabled = False
SelLength = 0
If Redomode = 3 Then
SelStart = Redopos - 2
Else
SelStart = Redopos
End If
If Redomode = 0 Or Redomode = 5 Then ' now we do the opposite

If Redomode = 5 Then


PopRedo

If glistN.SelStart = 0 Then
SelLength = Len(undotext) + 1
Else
SelLength = Len(undotext)
End If
    glistN_MarkDelete True
    glistN.MarkNext = 0
PushUndo
InsertText = Redotext
undopara = Redopara
undopos = Redopos
undotext = Redotext
undomode = Redomode
undoprev = Redoprev
Else
InsertText = Redotext

End If
ElseIf Redomode = 1 Then
If glistN.SelStart = 0 Then
SelLength = Len(Redotext) + 1
Else
SelLength = Len(Redotext)
End If
SelText = vbNullString
ElseIf Redomode = 2 Then
   keyzLineUp
ElseIf Redomode = 3 Then
glistN_SplitLine

End If

enabled = b
If Not undoprev Then
glistN.ShowPan
ManualInform
End If
If Not PopRedo Then
Redopara = -1
'Redoprev = False
Redotext = vbNullString
Redomode = 0
End If
If undoprev Then GoTo againY
End If
KeyCode = 0
ElseIf KeyCode = vbKeyZ And (shift And 3) = 2 Then
' 2nd version of undo system
' undopos now is selstart
profundo.MARKONE
againZ:
If profundo.MARKTWO > 1000 Then ProcTask2 Basestack1: profundo.MARKONE
If undopara <> -1 Then
' seve to redo...
PushRedo
Redopara = undopara
Redopos = undopos
Redotext = undotext
Redomode = undomode
If undomode = 2 Then Redopos = undopos + 2
Redoprev = undoprev

' MOVE TO PARAGRAPH...AND CHARPOS

b = enabled
enabled = False

   SelLength = 0
   SelStart = undopos

If undomode = 0 Or undomode = 4 Then

    If glistN.SelStart = 0 Then
        SelLength = Len(undotext) + 1
    Else
        SelLength = Len(undotext)
    End If
    
    If undomode = 4 Then
    glistN_MarkDelete True
    glistN.MarkNext = 0
    PopUndo
    PushRedo
        Redopara = undopara
        Redopos = undopos
        Redotext = undotext
        Redomode = undomode
        Redoprev = undoprev
    ' now we have undomode 5
    InsertText = undotext
    Else
    SelText = vbNullString
    End If
ElseIf undomode = 3 Then
    keyzLineUp
ElseIf undomode = 2 Then
    glistN_SplitLine
Else
        InsertText = undotext
End If
enabled = b
If Not undoprev Then
glistN.ShowPan
ManualInform
End If
End If
 If Not PopUndo Then
 ClearUndo
undopara = -1
undoprev = False
undotext = vbNullString
undomode = 0
End If
If undoprev Then GoTo againZ
KeyCode = 0
ElseIf KeyCode = vbKeyA And (shift And 3) = 3 Then
glistN.MarkNext = 0
BlockEndPara = mDoc.MarkParagraphID
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.ShowPan
KeyCode = 0

ElseIf KeyCode = vbKeyDelete Then
If HaveMarkedText Then
ClearRedo
PushUndo
glistN_MarkDelete False

KeyCode = 0
End If
ElseIf KeyCode = vbKeyBack Then
If HaveMarkedText Then
ClearRedo
PushUndo
glistN_MarkDelete False

KeyCode = 0
End If
End If
conthere:
If NoMark And EditDoc = False Then glistN.ShowMe2: Exit Sub
If KeyCode = 0 And shift = 0 Then Exit Sub
If KeyCode <> 0 Or shift <> 0 Then PressKey_A KeyCode, shift


End Sub

Private Sub glistN_KeyDownAfter(KeyCode As Integer, shift As Integer)
If NoMark Then Exit Sub
If HaveMarkedText Then
If shift = 0 Then
glistN.MarkNext = 0
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".MarkNo(" + str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".MarkNo()"
      End If
glistN.RefreshNow

ElseIf (shift And 1) = 1 Then

If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
glistN_MarkOut
End If
End If
ElseIf (shift And 1) = 1 And glistN.MarkNext = 0 Then
If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
glistN.MarkNext = 1
glistN_MarkIn
End If
ElseIf (shift And 1) = 1 And glistN.MarkNext > 0 Then
If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
glistN.MarkNext = 2
glistN_MarkOut
End If
End If
End Sub

Private Sub glistN_LineDown()
Dim b$, nline As Long
mDoc.TextParagraphLine glistN.ListIndex + 1, b$, nline
mDoc.MarkParagraphID = nline

If mDoc.IsLast(mDoc.MarkParagraphID) Then Exit Sub

  
 
 If mDoc.TextParagraphLen(nline) = 0 Then
 Charpos = 1
  mDoc.RemoveEmptyParagraph nline
 mDoc.MarkParagraphID = nline
 Else
 Charpos = ParaSelStart
  mDoc.ReWritePara nline, mDoc.TextParagraph(nline) + mDoc.DeleteNextParagraph(nline)

 End If
 
  all = mDoc.DocLines
 
  ParaSelStart = Charpos
End Sub

Private Sub glistN_LineUp()
Dim olID As Long, b$ ', c$
olID = mDoc.MarkParagraphID
If ParaSelStart < 2 Then
If mDoc.ParagraphOrder(mDoc.MarkParagraphID) > 1 Then
glistN.PressKey vbKeyLeft, 0
If olID <> mDoc.MarkParagraphID Then
b$ = mDoc.DeleteNextParagraph(mDoc.MarkParagraphID)
mDoc.ReWritePara mDoc.MarkParagraphID, mDoc.TextParagraph(mDoc.MarkParagraphID) + b$
 all = mDoc.DocLines
ClearRedo
PushUndo
undotext = vbCrLf
undopara = mDoc.MarkParagraphID
undopos = SelStart
undomode = 2
'End If
Render
End If
End If
Else
glistN.PressKey vbKeyLeft, 0
End If
all = mDoc.DocLines

End Sub

Private Sub keyzLineUp()
Dim olID As Long, b$ ', c$
olID = mDoc.MarkParagraphID
If ParaSelStart < 2 Then
If mDoc.ParagraphOrder(mDoc.MarkParagraphID) > 1 Then
glistN.PressKey vbKeyLeft, 0
If olID <> mDoc.MarkParagraphID Then
b$ = mDoc.DeleteNextParagraph(mDoc.MarkParagraphID)
mDoc.ReWritePara mDoc.MarkParagraphID, mDoc.TextParagraph(mDoc.MarkParagraphID) + b$
Render
End If
End If
Else
glistN.PressKey vbKeyLeft, 0
End If
all = mDoc.DocLines
End Sub

Public Sub MarkCut()
MarkCopy
PushUndo
glistN_MarkDelete False

End Sub
Private Sub FillBlock()
Dim order1 As Long, order2 As Long, a$, b$, i As Long, nn As Long
' TAKE FIRST VBCRLF IF FIRST LINE =""
On Error GoTo fbl

 Set UndoDoc = New Document
 UndoDoc.CrTail = False
    If BlockStartPara = BlockEndPara Then
        b$ = mDoc.TextParagraph(BlockStartPara)
        
        UndoDoc.CrTail = False
        If b$ = vbNullString Then
        Else
        If BlockStartPos < BlockEndPos Then
           '' UndoDoc.textDoc = Mid$(b$, BlockStartPos, BlockEndPos - BlockStartPos)
            UndoDoc.textParagraphPlusColor Mid$(b$, BlockStartPos, BlockEndPos - BlockStartPos), mDoc.Paracolordata(BlockStartPara), BlockStartPos - 1
            If BlockEndPos > Len(b$) + 1 Then UndoDoc.CrTail = True
        Else
          '' UndoDoc.textDoc = Mid$(b$, BlockEndPos, BlockStartPos - BlockEndPos)
            UndoDoc.textParagraphPlusColor Mid$(b$, BlockEndPos, BlockStartPos - BlockEndPos), mDoc.Paracolordata(BlockStartPara), BlockEndPos - 1
            If BlockStartPos > Len(b$) + 1 Then UndoDoc.CrTail = True
        End If
        End If
    Else
        order1 = mDoc.ParagraphOrder(BlockStartPara)
        order2 = mDoc.ParagraphOrder(BlockEndPara)
        If order2 > order1 Then '  normal
            a$ = mDoc.TextParagraph(BlockStartPara)
            UndoDoc.LfLead = False
         If BlockStartPos < 1 Then
               UndoDoc.AppendParagraphPlusColor a$, mDoc.Paracolordata(BlockStartPara)
               '' UndoDoc.AppendParagraph a$
            Else
            If Len(a$) + 1 < BlockStartPos Then
                UndoDoc.LfLead = True
                UndoDoc.AppendParagraph ""  '' no color
               Else
            
                ''UndoDoc.AppendParagraph Mid$(a$, BlockStartPos)
                UndoDoc.AppendParagraphPlusColor Mid$(a$, BlockStartPos), mDoc.Paracolordata(BlockStartPara), BlockStartPos - 1
                End If
            End If
            UndoDoc.MarkParagraphID = 1
            nn = BlockStartPara
            order2 = order2 - 1
            If order2 > order1 Then
                For i = order1 + 1 To order2
               ' UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn))
                UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraphPlusColor(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn), mDoc.Paracolordata(nn))
                Next i
            End If
                b$ = mDoc.TextParagraph(BlockEndPara)
                UndoDoc.CrTail = False
                
               If BlockEndPos > 0 Then
               ''UndoDoc.AppendParagraph MyLeft$(b$, BlockEndPos - 1) 'maybe is empty
               UndoDoc.AppendParagraphPlusColor Myleft$(b$, BlockEndPos - 1), mDoc.Paracolordata(BlockEndPara)
               Else
               UndoDoc.CrTail = True
               End If
        Else
            a$ = mDoc.TextParagraph(BlockEndPara)
        '   If a$ = VbNullString Then UndoDoc.AppendParagraph ""
         UndoDoc.LfLead = False
            If BlockEndPos < 1 Then
        
              UndoDoc.AppendParagraphPlusColor a$, mDoc.Paracolordata(BlockEndPara)
               ' UndoDoc.AppendParagraph a$
                
            Else
                 If Len(a$) + 1 < BlockEndPos Then
                UndoDoc.LfLead = True
                UndoDoc.AppendParagraph ""
            Else
                ' UndoDoc.AppendParagraph Mid$(a$, BlockEndPos)
                If BlockEndPos >= 1 Then
                  UndoDoc.AppendParagraphPlusColor Mid$(a$, BlockEndPos), mDoc.Paracolordata(BlockEndPara), BlockEndPos - 1
                  Else
                  UndoDoc.AppendParagraphPlusColor Mid$(a$, BlockEndPos), mDoc.Paracolordata(BlockEndPara)
                  End If
                End If
            End If
            UndoDoc.MarkParagraphID = 1
             nn = BlockEndPara
            order1 = order1 - 1
            If order1 > order2 Then
                For i = order2 + 1 To order1
                ''UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn))
                UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraphPlusColor(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn), mDoc.Paracolordata(nn))
                Next i
            End If
                b$ = mDoc.TextParagraph(BlockStartPara)
                If BlockStartPos < 1 Then BlockStartPos = 1
                  'UndoDoc.AppendParagraph MyLeft$(b$, BlockStartPos - 1) 'maybe is empty
                  UndoDoc.AppendParagraphPlusColor Myleft$(b$, BlockStartPos - 1), mDoc.Paracolordata(BlockStartPara)
        
        End If
End If
fbl:
End Sub

Public Sub MarkCopy()
If Not NoColor Then ColorBlockOnly
If HaveMarkedText Then
FillBlock
Dim pp1 As String, pp2 As String
pp1 = UndoDoc.textDoc

If Not mNoColor Then pp2 = SelHtmlText

blockKeyboard = True
MyDoEvents
blockKeyboard = False
 On Error Resume Next
 ProcTask2 Basestack1
On Error Resume Next
Clipboard.Clear
Err.Clear
If Not mNoColor Then
    SimpleHtmlData pp2, pp1
Else
    SetTextData 13, pp1
End If
End If
End Sub
Private Function SetTextData( _
        ByVal lFormatId As Long, _
        ByVal sText As String _
    ) As Boolean
    If lFormatId = 0 Then Exit Function
    Dim hMem As Long, lPtr As Long
    Dim lSize As Long, placeSize As Long
    
        lSize = LenB(sText)
        If lSize = 0 Then Exit Function
        If lSize Mod 2 = 1 Then
        placeSize = lSize + 1
        Else
        placeSize = lSize
        End If
        hMem = GlobalAlloc(0, placeSize + 2)
        If (hMem > 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, ByVal StrPtr(String$(placeSize \ 2 + 1, Chr$(0))), placeSize + 2
        CopyMemory ByVal lPtr, ByVal StrPtr(sText), lSize
        GlobalUnlock hMem
       If (OpenClipboard(Form1.hWnd) <> 0) Then
     SetClipboardData lFormatId, hMem
      CloseClipboard
      Else
      GlobalFree hMem
       End If
    End If
End Function

Private Sub glistN_MarkDelete(preservecursor As Boolean)
Dim order1 As Long, order2 As Long, a$, b$, i As Long
Dim oldid As Long, OLDPOS As Long
oldid = mDoc.MarkParagraphID
OLDPOS = Charpos
 If HaveMarkedText Then
 Set UndoDoc = New Document
    If BlockStartPara = BlockEndPara Then
        b$ = mDoc.TextParagraph(BlockStartPara)
        If BlockStartPos < BlockEndPos Then
            UndoDoc.textDoc = Mid$(b$, BlockStartPos, BlockEndPos - BlockStartPos)
            b$ = Mid$(b$, 1, BlockStartPos - 1) + Mid$(b$, BlockEndPos)
                        If OLDPOS > BlockStartPos And oldid = BlockStartPara Then
            OLDPOS = OLDPOS - (BlockEndPos - BlockStartPos)
            End If
            Charpos = BlockStartPos
        Else
           UndoDoc.textDoc = Mid$(b$, BlockEndPos, BlockStartPos - BlockEndPos)
            b$ = Myleft$(b$, BlockEndPos - 1) + Mid$(b$, BlockStartPos)
            If OLDPOS > BlockEndPos And oldid = BlockStartPara Then
            OLDPOS = OLDPOS - (BlockStartPos - BlockEndPos)
            End If
            Charpos = BlockEndPos
        End If
        mDoc.ReWritePara BlockStartPara, b$
    Else
        order1 = mDoc.ParagraphOrder(BlockStartPara)
        order2 = mDoc.ParagraphOrder(BlockEndPara)
  
        If order2 > order1 Then '  normal
            UndoDoc.AppendParagraph "" ' start paragraph no 1
            UndoDoc.MarkParagraphID = 1
        
            b$ = mDoc.TextParagraph(BlockEndPara)
            mDoc.RemoveDocParaIndex BlockEndPara
            If BlockEndPos < 1 Then
            UndoDoc.AppendParagraph ""
            b$ = vbNullString
            Else
            UndoDoc.AppendParagraph Myleft$(b$, BlockEndPos - 1) 'maybe is empty
            b$ = Mid$(b$, BlockEndPos)
            End If
            
            
            order2 = order2 - 1
            If order2 > order1 Then
                For i = order1 + 1 To order2
                UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.DeleteNextParagraph(BlockStartPara))
                    
                Next i
            End If
            a$ = mDoc.TextParagraph(BlockStartPara)
            UndoDoc.ReWritePara 1, a$, True
            If BlockStartPos = 1 Then
                mDoc.ReWritePara BlockStartPara, b$
                UndoDoc.ReWritePara 1, a$, True
            Else
            UndoDoc.ReWritePara 1, Mid$(a$, BlockStartPos), True
                mDoc.ReWritePara BlockStartPara, Myleft$(a$, BlockStartPos - 1) + b$
            End If
            mDoc.MarkParagraphID = BlockStartPara
            Charpos = BlockStartPos
        Else
            UndoDoc.AppendParagraph "" ' start paragraph no 1
            UndoDoc.MarkParagraphID = 1
             b$ = mDoc.TextParagraph(BlockStartPara)
            mDoc.RemoveDocParaIndex BlockStartPara
             UndoDoc.AppendParagraph Myleft$(b$, BlockStartPos - 1) 'maybe is empty
            b$ = Mid$(b$, BlockStartPos)
            order1 = order1 - 1
            If order1 > order2 Then
                For i = order2 + 1 To order1
                    
                    UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.DeleteNextParagraph(BlockEndPara))
               
                Next i
            End If
            a$ = mDoc.TextParagraph(BlockEndPara)
            
            If BlockEndPos = 1 Then
                mDoc.ReWritePara BlockEndPara, b$
                UndoDoc.ReWritePara 1, a$
            Else
            UndoDoc.ReWritePara 1, Mid$(a$, BlockEndPos)
                mDoc.ReWritePara BlockEndPara, Myleft$(a$, BlockEndPos - 1) + b$
            End If
            mDoc.MarkParagraphID = BlockEndPara
            
            Charpos = BlockEndPos
        
        
        End If

End If
all = mDoc.DocLines
If Not preservecursor Then
undopara = mDoc.MarkParagraphID
undopos = SelStart
undotext = UndoDoc.textDoc
undomode = 1 ' insert text at position changed to 1
End If
Set UndoDoc = New Document


glistN.MarkNext = 0
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
If preservecursor Then
If mDoc.ParagraphOrder(oldid) > 0 Then

mDoc.MarkParagraphID = oldid
Charpos = OLDPOS
End If
End If
Render
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".MarkNo(" + str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".MarkNo()"
      End If
End If

End Sub

'
Private Sub glistN_MarkDestroy()
' destroy or open to change
If Not ClickInsideMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".MarkNo(" + str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".MarkNo()"
      End If
Else
glistN.MarkNext = 1
' we have to raise event now
End If
End Sub

Private Sub glistN_MarkDestroyAny()
' destroy always
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".MarkNo(" + str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".MarkNo()"
      End If
End Sub

Private Sub glistN_MarkIn()
expandCharPos = 0
If Not glistN.EditFlag Then glistN.MarkNext = 0: Exit Sub
'If Len(CurrentParagraph) + 1 < Charpos Then
'SelStartSilent = CharPosStart - Charpos + 1 + Len(CurrentParagraph)
'End If
BlockStartPara = mDoc.MarkParagraphID
BlockStartPos = ParaSelStart
BlockStartCursor = glistN.SelStart
BlockEndCursor = BlockStartCursor + 1

BlockStartLine = glistN.ListIndex
BlockEndline = BlockStartLine
BlockEndPara = BlockStartPara
BlockEndPos = BlockStartPos
End Sub
Function HaveMarkedText() As Boolean
If BlockEndPos < 1 Then BlockEndPos = 1
If glistN.MarkNext = 0 Then
Else
If BlockStartPara <> BlockEndPara Then
HaveMarkedText = True
ElseIf BlockStartPos <> BlockEndPos Then
HaveMarkedText = True
End If
End If
End Function
Function MarkedTextAny() As Boolean
If BlockStartPara <> BlockEndPara Then
MarkedTextAny = True
ElseIf BlockStartPos <> BlockEndPos Then
MarkedTextAny = True
End If
End Function


Function ClickInsideMarkedText() As Boolean
' we need that for drag and drop, copy or move
' it is simple
' we now the current line and selstart of glistN
If MarkedTextAny Then
If BlockTopline <> BlockBottomline Then
If glistN.ListIndex = BlockTopline Then
    If BlockEndline < BlockStartLine Then
    If glistN.SelStart >= BlockEndCursor Then ClickInsideMarkedText = True
    Else
    If glistN.SelStart >= BlockStartCursor Then ClickInsideMarkedText = True
    End If
ElseIf glistN.ListIndex = BlockBottomline Then
    If BlockEndline < BlockTopline Then
    If glistN.SelStart <= BlockStartCursor Then ClickInsideMarkedText = True
    Else
    If glistN.SelStart <= BlockEndCursor Then ClickInsideMarkedText = True
    End If
ElseIf glistN.ListIndex < BlockTopline Or glistN.ListIndex > BlockBottomline Then
' FALSE
Else
    ClickInsideMarkedText = True
End If
ElseIf BlockTopline = glistN.ListIndex Then
If BlockEndCursor < BlockStartCursor Then
If glistN.SelStart >= BlockEndCursor And glistN.SelStart <= BlockStartCursor Then ClickInsideMarkedText = True
Else
If glistN.SelStart >= BlockStartCursor And glistN.SelStart <= BlockEndCursor Then ClickInsideMarkedText = True
End If

End If
End If
End Function
Sub RelocateMarkedText()
' when we resize glist, need to calculate the position of marked text
' we use a second set of variables
'  BlockStartPara As Long, BlockStartPos As Long
' BlockEndPara As Long, BlockEndPos As Long
'
BlockStartLine = mDoc.FirstParagraphLine(BlockStartPara) + mDoc.TextParagraph2Line(BlockStartPara, BlockStartPos, BlockStartCursor)
BlockEndline = mDoc.FirstParagraphLine(BlockEndPara) + mDoc.TextParagraph2Line(BlockEndPara, BlockEndPos, BlockEndCursor)
If BlockEndline < BlockStartLine Then
    BlockTopline = BlockEndline
    BlockBottomline = BlockStartLine
Else
    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
End If
End Sub
Private Sub glistN_MarkOut()
'
expandCharPos = 0
'If Len(CurrentParagraph) + 1 < Charpos Then
'SelStartSilent = CharPosStart - Charpos + 1 + Len(CurrentParagraph)
'End If
BlockEndPara = mDoc.MarkParagraphID
BlockEndPos = ParaSelStart
BlockEndline = glistN.ListIndex
BlockEndCursor = glistN.SelStart
If BlockEndline < BlockStartLine Then
    BlockTopline = BlockEndline
    BlockBottomline = BlockStartLine
Else
    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
End If
'
If HaveMarkedText Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".MarkOut(" + str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".MarkOut()"
      End If
Else
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".MarkNo(" + str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".MarkNo()"
      End If
End If
End Sub

Private Sub glistN_ReadListItem(item As Long, content As String)
If item >= 0 Then

content = mDoc.TextLine2(item + 1, showparagraphWork And showparagraph)
If showparagraph And showparagraphWork Then
content = Replace(content, vbTab, ChrW(&H21E5))
End If
End If
End Sub



'End Sub

Private Sub glistN_ScrollSelected(item As Long, Y As Long)
If NoMark Then Exit Sub
Dim what$, ThatPara As Long
   mDoc.TextParagraphLine item, what$, ThatPara
   mDoc.MarkParagraphID = ThatPara
End Sub

Private Sub glistN_Selected(item As Long)
If undomode = 1 And undotext <> "" Then PushUndo
If mEditDoc <> glistN.EditFlag Then EditDoc = EditDoc
Dim what$, ThatPara As Long
mDoc.TextParagraphLine2 item, what$, ThatPara, showparagraphWork And showparagraph
mDoc.MarkParagraphID = ThatPara
If glistN.DragEnabled Then Exit Sub
Dim M As Long, z As Long


    If ClickInsideMarkedText Then
        If Not M2000code Then
        If BlockStartLine = BlockEndline Then
        M = SelLength
        SelStartSilent = SelStart
        SelLengthSilent = SelLength + 1
        If Right$(SelText, 1) = "." Then
            If BlockStartPos < BlockEndPos Then
            M = BlockEndPos
            Else
            M = BlockStartPos
            End If
            z = mDoc.TextParagraphLen(ThatPara)
            While InStr(glistN.WordCharLeft, Mid$(mDoc.TextParagraph(ThatPara), M, 1)) = 0 And M < z
            M = M + 1
            Wend
            If BlockStartPos < BlockEndPos Then
            SelLengthSilent = M - BlockStartPos
            Else
            SelLengthSilent = M - BlockEndPos
            End If
            FillBlock
         Else
            SelLengthSilent = SelLength - 1
        End If
        End If
    Else
        If BlockStartLine = BlockEndline Then
            If BlockStartPos < BlockEndPos Then
                If Mid$(mDoc.TextParagraph(ThatPara), BlockEndPos, 1) = "(" Then
                SelStartSilent = SelStart
                SelLengthSilent = SelLength + 1
                FillBlock
                Exit Sub
                Else
                z = InStr(2, SelText, ".")
                End If
            Else
                z = InStr(2, SelText, ".")
            End If
            If z > 1 Then
                SelStartSilent = SelStart + z
                SelLengthSilent = SelLength - z
                FillBlock
            ElseIf Right$(SelText, 1) = "(" Then
                SelStartSilent = SelStart
                SelLengthSilent = SelLength - 1
                FillBlock
            Else
                M = SelLength
                SelStartSilent = SelStart
                SelLengthSilent = SelLength + 1
                If M + 1 <> SelLength Then
                    SelLengthSilent = M
                    FillBlock
                ElseIf Right$(SelText, 1) = "(" Then
                    SelStartSilent = LastSelStart
                    FillBlock
                ElseIf Right$(SelText, 1) = "." Then
                   If BlockStartPos < BlockEndPos Then
                   M = BlockEndPos
                   Else
                   M = BlockStartPos
                   End If
                   z = mDoc.TextParagraphLen(ThatPara)
                    While InStr(glistN.WordCharLeft, Mid$(mDoc.TextParagraph(ThatPara), M, 1)) < 2 And M < z
                        M = M + 1
                    Wend
                    If BlockStartPos < BlockEndPos Then
                    If InStr(glistN.WordCharLeft, Mid$(mDoc.TextParagraph(ThatPara), M, 1)) <> 0 Then
                        SelLengthSilent = M - BlockStartPos
                    Else
                        SelLengthSilent = M - BlockStartPos + 1
                    End If
                    Else
                    SelLengthSilent = M - BlockEndPos
                    End If
                    
                    FillBlock
                Else
                    SelLengthSilent = SelLength - 1
                End If
            End If
        End If
    End If
End If

End Sub


Private Sub glistN_Selected2(item As Long)
If Not EditDoc Then
 
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click()"
      End If
Else

If undomode = 1 And undotext <> "" Then PushUndo
If ClickInsideMarkedText Then
' start drag
glistN.MarkNext = 4

glistN.DragNow
End If
End If
End Sub


Private Sub glistN_ShowExternalCursor()
glistN.PrepareToShow 5
End Sub


Private Sub glistN_SplitLine()
'
Dim b$, dummy As Long, a As Long
If NumberOnly Or NumberIntOnly Or UsedAsTextBox Then Exit Sub
If glistN.ListIndex >= 0 Then
    b$ = mDoc.TextParagraph(mDoc.MarkParagraphID)
If Charpos = 1 And b$ <> "" Then
    dummy = mDoc.InsertBeforeParagraph(mDoc.MarkParagraphID, "")
        mDoc.ColorThis (dummy)
     all = mDoc.DocLines
    mDoc.MarkParagraphID = dummy
    SetRowColumn glistN.ListIndex + 2, 0

    
  ''
   
ElseIf ParaSelStart = 1 Then
    mDoc.MarkParagraphID = mDoc.InsertAfterParagraph(mDoc.MarkParagraphID, "")
    all = mDoc.DocLines
    ParaSelStart = 1

ElseIf ParaSelStart > Len(b$) Then

    mDoc.MarkParagraphID = mDoc.InsertAfterParagraph(mDoc.MarkParagraphID, "")
    
      all = mDoc.DocLines
    ParaSelStart = Len(b$)

Else  ' ANYWHERE

dummy = ParaSelStart

    mDoc.ReWritePara mDoc.MarkParagraphID, Myleft$(b$, ParaSelStart - 1), False ' write and format
    
    mDoc.MarkParagraphID = mDoc.InsertAfterParagraph(mDoc.MarkParagraphID, Mid$(b$, dummy))
    mDoc.ColorThis (mDoc.MarkParagraphID)
    all = mDoc.DocLines
    ParaSelStart = 1

End If
End If
End Sub


Private Sub glistN_SyncKeyboard(item As Integer)
If NoMark Then item = 0
If blockKeyboard Then item = 0
If item > 32 And item <> 127 Then GetUndoNext = True
End Sub
Private Sub glistN_RegisterGlist(this As gList)
On Error Resume Next
Set LastGlist = this
If Err.Number > 0 Then this.NoWheel = True
End Sub
Private Sub glistN_UnregisterGlist()
On Error Resume Next
If glistN.TabStopSoft Then Callback.LastActive = ctrlName
Set LastGlist = Nothing
If Err.Number > 0 Then glistN.NoWheel = True
End Sub

Private Sub glistN_WindowKey(VbKeyThis As Integer)
    Callback.ProcessWinKey VbKeyThis
End Sub

Private Sub glistN_WordMarked(ThisWord As String)
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Word(" + str(mIndex) + ",{" + ThisWord + "})"
   Else
      Callback.Callback mesCtrlName + ".Word({" + ThisWord + "})"
      End If
      If mEditDoc = False Then ThisWord = vbNullString
End Sub

Private Sub mDoc_BreakLine(Data As String, datanext As String)
 ' Nowrap = True mean no Wrap
 Dim data1 As String, datanext1 As String
 If nowrap Then
    datanext = Data
    Data = vbNullString
 Else
    If showparagraph Then
         data1 = Replace(Replace(Data, " ", ChrW(&HB7)), ChrW(&H2007), ChrW(&H2423))
        
        glistN.BreakLine data1, datanext1, , ChrW(&HB7)
        If data1 <> "" Then
        
        datanext = Myleft$(Data, Len(datanext1))
        Data = Right$(Data, Len(data1))
        Else
            datanext = Data
             Data = vbNullString
        End If
      
    Else
        glistN.BreakLine Data, datanext
    End If
End If

End Sub
Private Sub makebase(basecolor As Long)
ReDim Col(1) As Long
Col(0) = basecolor
Col(1) = 1

End Sub
Private Sub addone(ByVal thispos As Long, ByVal thiscolor As Long)
Dim K As Long

K = UBound(Col())
If Col(1) = 0 Then
makebase 14
If thispos <> 1 Then
K = UBound(Col())
Else
Col(0) = thiscolor
Exit Sub
End If
End If
' LOOK THIS SECOND TIME
If K >= 1 Then
If K > 1 Then If Col(K - 1) = thiscolor Then Exit Sub
If Col(K) >= thispos Then

'If col(k) >= thispos Then
'thispos = thispos + 1
'End If
'col(k) = thispos
Exit Sub
End If
End If
ReDim Preserve Col(K + 2) As Long
Col(K + 1) = thiscolor
Col(K + 2) = thispos


End Sub

Private Sub addone2000(ByVal thispos As Long, ByVal thiscolor As Long)
Dim K As Long
If thiscolor = 0 Then thiscolor = 14
K = UBound(Col())
If Col(1) = 0 Then
makebase 14
If thispos <> 1 Then
K = UBound(Col())
Else
Col(0) = thiscolor
Exit Sub
End If
End If
' LOOK THIS SECOND TIME
If K >= 1 Then
If K > 1 Then If Col(K - 1) = thiscolor Then Exit Sub
If Col(K) = thispos Then
Col(K - 1) = thiscolor: Exit Sub
End If
End If

''''If col(k - 1) = thiscolor And k > 1 Then Exit Sub

ReDim Preserve Col(K + 2)
Col(K + 1) = thiscolor
Col(K + 2) = thispos


End Sub
Private Sub ColorizeExpressions2000(ByVal a$, K As Long, pos2 As Long, feedthis As Long)
Dim drop$, pos1 As Long, kind As Long, localfeedthis As Long, thispara1 As Long, lastpar As Long
Dim lastcol As Long, bracket As Long
Dim n$
again:
If a$ = vbNullString Then
If feedthis = 0 Then
If lastpar > 0 Then
feedthis = lastpar
End If
End If
Exit Sub
End If
drop$ = "@(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then
    If IsSymbol(a$, "@") Then
        If pos2 > 0 Then
            addone2000 pos2, 14
        Else
            addone2000 1, 14
        End If
        pos2 = K - Len(a$) + 1
    
    End If
getlabelonly:
    n$ = a$
    kind = IsLabelOnly2000(n$, drop$)
    pos2 = K - Len(a$) + 1
    If kind = 0 And n$ <> a$ Then
    feedthis = 100
    a$ = n$: Exit Sub
    End If
    a$ = n$
    If kind = 100 Then
    If pos2 > 0 Then
            addone2000 pos2, 14
        Else
            addone2000 1, 14
        End If
         pos2 = K - Len(a$) + 1
         
    kind = 0
    GoTo again
    
    End If
End If
If kind = 0 Then drop$ = "$(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "~(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then
  If FastSymbol(a$, "->", , 2) Then
    pos2 = K - Len(a$) - 1
    addone2000 pos2, 11
        a$ = NLtrim$(a$)
        pos2 = K - Len(a$) + 1
        If MaybeIsSymbolNoSpace(a$, "{") Then
        pos2 = K - Len(a$) + 1
        addone2000 pos2, 14
        a$ = Mid$(a$, 2)
         localfeedthis = 0
           pos2 = K - Len(a$) + 1
           If pos2 <= K Then
                drop$ = block(a$)
                If Len(drop$) > 0 Then
                    ColorizeInnerStatements drop$, K - Len(a$), pos2, localfeedthis
           
                    If Len(a$) > 0 Then ColorizeInnerStatements a$, K, pos2, localfeedthis
                    drop$ = vbNullString
                Else
                ColorizeInnerStatements a$, K, pos2, localfeedthis
               End If

           End If
            kind = 0
            drop$ = vbNullString
            If pos2 > K Then Exit Sub
        Else
           GoTo checklambda
            
        End If
    
    End If
End If
If kind = 0 Then drop$ = vbNullString
If kind > 0 Then
    If ismine(drop$) Then
    lastcol = 11
    If kind > 4 Then lastpar = 11
    If pos2 > 0 Then
       addone2000 pos2, 11
    Else
       addone2000 K - Len(a$) - Len(drop$) + 1, 11
       End If
        '' for the end
   Else
   If kind > 4 And kind < 9 Then lastpar = 10 - (kind = 8) * 3
   If kind = 10 Then
    If pos2 > 0 Then
       addone2000 pos2, 13
    Else
       addone2000 K - Len(a$) - Len(drop$) + 1, 13
       End If
        lastcol = 13
   kind = 0
   ElseIf kind = 20 Then
        lastcol = 12
   Else
        lastcol = 10 - (kind = 8) * 3
   End If
   
   If pos2 > 0 Then
        If kind = 8 Then
            addone2000 pos2, 10
            addone2000 pos2 + Len(drop$) - 1, lastcol
        Else
            addone2000 pos2, lastcol
            If lastcol = 13 And Myleft$(LTrim$(drop$), 1) <> "0" And Myleft$(LTrim$(drop$), 1) <> "#" Then
                If UCase(Myleft$(LTrim$(drop$), 2)) = "&H" And Len(drop$) > 2 Then
       
            Else
            addone2000 pos2 + 1, 10
            End If
       End If
        End If
    Else
        If kind = 8 Then
        addone2000 K - Len(a$) - Len(drop$), 10
        addone2000 K - Len(a$) - Len(drop$) + 1, 11
        Else
       addone2000 K - Len(a$) - Len(drop$) + 1, lastcol
       End If
    End If
  End If
  pos2 = K - Len(a$) + 1
 If kind > 3 Then
checklambda:
  Do
  pos1 = 1
  drop$ = aheadstatus2000(a$, False, pos1)
  If pos1 > 1 Then
  ColorizeExpressions2000 Mid$(a$, 1, pos1 - 1), K - Len(a$) + pos1 - 1, pos2, feedthis
  If feedthis = 14 Or feedthis = 11 Then
  If lastcol <> 11 And lastcol <> 10 And lastcol <> 13 Then lastcol = feedthis
  feedthis = 0
  End If
  If feedthis > 0 And pos1 > Len(a$) Then Exit Sub
  feedthis = 0
   a$ = Mid$(a$, pos1)
   End If
   If FastSymbol(a$, ",") Then
     addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
   Else
   Exit Do
   End If
   Loop
    If MaybeIsSymbol(a$, "(") Then
addone2000 pos2, 14
lastpar = 14
FastSymbol a$, "("
pos2 = K - Len(a$) + 1
  pos1 = 1
  drop$ = aheadstatus2000(a$, False, pos1)
    thispara1 = thispara
  If pos1 > 1 Then

  ColorizeExpressions2000 Mid$(a$, 1, pos1 - 1), K - Len(a$) + pos1 - 1, pos2, feedthis
    a$ = Mid$(a$, pos1)
    End If
    If thispara1 <> thispara Then
    Exit Sub
    End If
    If kind = 8 Then
      addone2000 pos2, 13
      FastSymbol a$, "]"
    
    Else
      addone2000 pos2, 14
      FastSymbol a$, ")"
    End If
    pos2 = K - Len(a$) + 1
ElseIf IsSymbol(a$, ",") Then
addone2000 pos2, 14
pos2 = K - Len(a$) + 1
    GoTo again
ElseIf MaybeIsSymbol(a$, "]") Then
    pos2 = K - Len(a$) + 1
    addone2000 pos2, lastcol
    If FastSymbol(a$, "][", , 2) Then
    GoTo checklambda
    Else
        pos2 = K - Len(a$) + 2
        addone2000 pos2, 14
        FastSymbol a$, "]"
        pos2 = K - Len(a$) + 1
    End If
ElseIf MaybeIsSymbol(a$, ")") Then
    pos2 = K - Len(a$) + 1
    If kind >= 10 Then
    addone2000 pos2, 14
    
    ElseIf kind > 0 Then
  addone2000 pos2, lastcol
  Else
  addone2000 pos2, 14
  End If
  pos2 = K - Len(a$) + 2
  addone2000 pos2, 14
  FastSymbol a$, ")"
   
    pos2 = K - Len(a$) + 1
  End If
  End If

GoTo again
ElseIf MaybeIsSymbol(a$, "(", , True) Then
    addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
   GoTo again
ElseIf MaybeIsSymbol(a$, "{") Then
addone2000 pos2, 15
pos1 = 1
bracket = 0
blockStringAhead2 a$, pos1, bracket
If bracket = 0 Then feedthis = 0 Else feedthis = 15
a$ = Mid$(a$, pos1 + 1)
drop$ = vbNullString
pos2 = K - Len(a$) + 1
If feedthis = 15 Then Exit Sub
    GoTo again

ElseIf MaybeIsSymbol(a$, Chr(34)) Then

If Not ISSTRINGA(a$, drop$) Then
a$ = vbNullString
addone2000 pos2, 10
Else
addone2000 pos2, 15
End If
drop$ = vbNullString
pos2 = K - Len(a$) + 1
    GoTo again
ElseIf MaybeIsSymbol(a$, Chr(34)) Then

If Not ISSTRINGA(a$, drop$) Then
a$ = vbNullString
addone2000 pos2, 10
Else
addone2000 pos2, 15
End If
drop$ = vbNullString
pos2 = K - Len(a$) + 1
    GoTo again
Else
If Col((UBound(Col()) \ 2) * 2) <> 14 Then addone2000 pos2, 14
a$ = Mid$(a$, 2)
If Len(a$) > 0 Then GoTo again
End If

If pos2 > 0 Then addone2000 pos2, 14


End Sub
Private Sub ColorizeExpressions(a$, K As Long, pos2 As Long, feedthis As Long)
Dim drop$, pos1 As Long, kind As Long
Dim lastcol As Long
again:
 
 If a$ = vbNullString Then Exit Sub
' Exit Sub
If Len(OpenTag) > 0 Then
    
a$ = LTrim$(a$)
     If a$ = vbNullString Then Exit Sub
    If Myleft$(a$, Len(MultiLineComment1)) = MultiLineComment1 Then
    feedthis = -15
    pos2 = K - Len(a$)
    Exit Sub
    ElseIf Myleft$(a$, 1) = "=" Then
   
If pos2 = 0 Then pos2 = 1
lastcol = 14
    a$ = Mid$(a$, 2)
addone pos2, lastcol
pos2 = pos2 + 1
    GoTo herestr
    ElseIf Myleft$(a$, 2) = "/>" Then
    pos2 = K - Len(a$) + 1
    
    
    a$ = Mid$(a$, 3)
    
    addone pos2, 14
    pos2 = pos2 + 2
    GoTo there001
    ElseIf Myleft$(a$, 1) = ">" Then
     pos2 = K - Len(a$) + 1
    a$ = Mid$(a$, 2)

   If feedthis = 1 Then feedthis = 0
            If pos2 > 0 Then
                addone pos2, 14
            Else
                addone 1, 14
            End If
      pos2 = pos2 + 1
there001:
     If a$ = vbNullString Then Exit Sub
    kind = IsLabelOnly(a$, drop$)
    If kind = 1 Then
    If feedthis = 0 Then
                If pos2 > 0 Then
                addone pos2, 10
            Else
                addone 1, 10
            End If
    Else
    If Len(drop$) > 0 Then Mid$(drop$, 1, 1) = Chr$(0) Else drop$ = Chr$(0)
    End If
    End If
   
    lastcol = 10
    If Not FastSymbolNoRemove(a$, MultiLineComment1, , Len(MultiLineComment1)) Then
    
     If Not MaybeIsSymbol2(a$, Len(CloseTag), CloseTag) Then
     feedthis = -2
   If MaybeIsSymbol(a$, "<") Then GoTo again
   ElseIf feedthis = -2 Then
   GoTo again
   ElseIf feedthis = 0 Then
   If MaybeIsSymbol(a$, "<") Then GoTo again
   GoTo cont123
   End If
   Else
 '  pos2 = k - Len(a$)
    feedthis = -15
    pos2 = K - Len(a$)
    Exit Sub
   End If
       feedthis = -2
cont123:
    ElseIf Myleft$(a$, Len(CloseTag)) = CloseTag Then
     pos2 = K - Len(a$) + 1
        If pos2 > 0 Then
                addone pos2, 14
            Else
                addone 1, 14
            End If
      a$ = Mid$(a$, Len(CloseTag) + 1)
             pos2 = K - Len(a$)
    If a$ = vbNullString Then Exit Sub
    kind = IsLabelOnly(a$, drop$)
   
        feedthis = 1
        
    ElseIf Myleft$(a$, Len(OpenTag)) = OpenTag Then
    pos2 = K - Len(a$) + 1
    a$ = Mid$(a$, Len(OpenTag) + 1)
    
        If pos2 > 0 Then
                addone pos2, 14
            Else
                addone 1, 14
            End If
feedthis = 1
         pos2 = pos2 + 1
         If a$ = vbNullString Then Exit Sub
         If AscW(a$) < 33 Then a$ = "": Exit Sub
         kind = IsLabelOnly(a$, drop$)
         If kind = 0 Then feedthis = 0: GoTo again
         If Len(a$) > 0 Then If MaybeIsSymbol(a$, "=") Then Exit Sub
        ' End If

    
    
    Else
    
    kind = IsLabelOnly(a$, drop$)
  
    If kind > 0 Then
        If drop$ = "" Then drop$ = Chr$(0)
        If feedthis <> -2 Then feedthis = 0 Else Mid$(drop$, 1, 1) = "-"
        pos2 = K - Len(a$) + 1 - Len(drop$)
        If pos2 = 0 Then pos2 = 1
    Else
    If Len(a$) > 0 Then kind = 1: drop$ = "-": a$ = Mid$(a$, 2)
     pos2 = K - Len(a$)
    If pos2 = 0 Then pos2 = 1
    End If
    
    
End If
Else
kind = IsLabelOnly(a$, drop$)
If kind > 0 Then feedthis = 0
End If

If kind = 100 Then

    If pos2 > 0 Then
            addone pos2 - Len(drop$), 14
        Else
            addone 1, 14
        End If
         pos2 = K - Len(a$) + 1
         
    kind = 0
    GoTo again
    
    End If
If kind = 0 Then drop$ = vbNullString
If kind > 0 Then
    If IsmineHere2(drop$) Then
    lastcol = 11
    If pos2 > 0 Then
       addone pos2, 11
    Else
       addone K - Len(a$) - Len(drop$) + 1, 11
       End If
    ElseIf IsmineHere(drop$) Then  ' FALL TO MAIN VOC
    
   If Len(OpenTag) > 0 Then
   If feedthis = 1 Then lastcol = 11 Else lastcol = 14
    Else
    lastcol = 11
    End If
             If pos2 > 0 Then
        addone pos2, lastcol
    Else
       addone K - Len(a$) - Len(drop$) + 1, lastcol
       End If
       If Len(OpenTag) > 0 Then
   If MaybeIsSymbol(a$, "[=>/]") Then
'drop$ = drop$ + " "

pos2 = K - Len(a$) + 1
lastcol = 14
drop$ = ""
GoTo cont0010
Else
drop$ = ""
End If
End If
        '' for the end
   Else
    If Len(OpenTag) > 0 And kind = 1 And feedthis <> -2 Then
 '  pos2 = k - Len(a$) + 1

        If MaybeIsSymbol2(a$, 1, "[=>/]") Then
            lastcol = 14
            addone pos2, 14
cont0010:
            If FastSymbol(a$, "=") Then

                pos2 = K - Len(a$) '+ 1

                addone pos2, lastcol
herestr:
                If FastSymbol(a$, StringSep) Then
                    pos2 = K - Len(a$)
                    drop$ = blockString2(a$)
                    addone pos2, 15
                    pos2 = K - Len(a$) + 1
                    lastcol = 14
                    drop$ = ""
                    GoTo again
                ElseIf FastSymbol(a$, stringsep2, , Len(stringsep2)) Then
                    pos2 = K - Len(a$)
                    drop$ = blockString3(a$, stringsep2)
                    addone pos2, 15
                    pos2 = K - Len(a$) + 1
                    lastcol = 14
                    drop$ = ""
                    GoTo again
                Else
                    pos2 = K - Len(a$) + 1
                    kind = IsLabelOnly(a$, drop$)
                    If Len(drop$) = 0 Then
                        addone pos2, 10
                        pos2 = K - Len(a$) + 1
                        lastcol = 14
                    ElseIf kind > 0 Then
                        lastcol = 14
                        pos2 = K - Len(a$) - Len(drop$) + 1
                        addone pos2, 10
                        pos2 = K - Len(a$) + 1
                        drop$ = ""
                    End If
                    GoTo again
                End If
       Else
        GoTo again
        lastcol = 14
       End If
       
    Else
    If lastcol = 11 Or lastcol = 15 Then lastcol = 14 Else lastcol = 10
    
    
    End If
   Else
   If kind = 10 Then lastcol = 14 Else lastcol = 10
   pos2 = K - Len(a$) '+ 1
   End If
   
   If pos2 > 0 And Len(drop$) = 0 Then
       addone pos2, lastcol
   Else
       addone K - Len(a$) - Len(drop$) + 1, lastcol
      
       End If
  End If
  
  pos2 = K - Len(a$) + 1
  If kind > 3 Then
  Do
  pos1 = 1
  drop$ = aheadstatus(a$, False, pos1)
  If pos1 > 1 Then
  ColorizeExpressions Mid$(a$, 1, pos1 - 1), K - Len(a$) + pos1 - 1, pos2, feedthis
  ' ColorizeExpressions Mid$(a$, 1, pos1 - 1), k, pos2, feedthis
  If feedthis = 1 Then Exit Sub
  ' check next paragraph for symbols as starting position, if not feedthis
 If feedthis > 0 And pos1 >= Len(a$) Then Exit Sub
  feedthis = 0
   a$ = Mid$(a$, pos1)
   End If
   If FastSymbol(a$, mSplitExpr) Then
     addone pos2, 14
    pos2 = K - Len(a$) + 1
   ElseIf MaybeIsSymbol(a$, OtherSymbols, , True) Then
     addone pos2, 14
    pos2 = K - Len(a$) + 1
   Else
   Exit Do
   End If
   Loop
    If FastSymbol(a$, ")") Then
  addone pos2, lastcol
    pos2 = K - Len(a$) + 1
  End If
  If Brackets Then
    If FastSymbol(a$, "]") Then
  addone pos2, lastcol
    pos2 = K - Len(a$) + 1
  End If
  End If
  End If

GoTo again
ElseIf IsSymbol(a$, mSplitExpr) Then
addone pos2, 14
pos2 = K - Len(a$) + 1
    GoTo again
ElseIf MaybeIsSymbol(a$, OtherSymbols, , True) Then
addone pos2, 14
pos2 = K - Len(a$) + 1
    GoTo again
 ElseIf MaybeIsSymbol(a$, LineComment2) Then
drop$ = "  "
If pos2 > 0 Then
 addone pos2, 7

Else
addone 1, 7
End If
feedthis = 1
 Exit Sub
ElseIf FastSymbolNoRemove(a$, MultiLineComment1, , Len(MultiLineComment1)) Then
pos2 = K - Len(a$)
cont0030:
If MultiLineComment1 <> "{" Then
feedthis = -15
'pos2 = k - Len(a$) + 1
'addone pos2, 7
'pos2 = k - Len(a$) - Len(MultiLineComment1) + 1
Exit Sub
ElseIf MaybeIsSymbol2(a$, ComSymbolsWidth, CommentSymbols) Then
drop$ = "  "
If pos2 > 0 Then
 addone pos2, CLng(IIf(CommentLineLight, 12, 7))

Else
addone 1, CLng(IIf(CommentLineLight, 12, 7))
End If
feedthis = 1
  Exit Sub

   
Else

    FastSymbol a$, MultiLineComment1, , Len(MultiLineComment1)
    addone pos2, 15
    
    drop$ = blockString(a$)
    If drop$ <> "" Or (drop$ = vbNullString And Myleft$(a$, 1) = mMultiLineComment2) Then
        If feedthis <> -2 Then feedthis = 0
        If Not MaybeIsSymbol2(a$, Len(mMultiLineComment2), mMultiLineComment2, , , True) Then
        feedthis = 15
        Else
        End If
    Else
        feedthis = 15
    End If
    drop$ = vbNullString
    pos2 = K - Len(a$) + 1
    If feedthis = 15 Then Exit Sub
    GoTo again
End If
ElseIf MaybeIsSymbol(a$, StringSep + Myleft$(stringsep2, 1)) Then

If Not ISSTRINGb(a$, drop$) Then
a$ = vbNullString
addone pos2, 10
Else
addone pos2, 15
End If
drop$ = vbNullString
pos2 = K - Len(a$) + 1
    GoTo again

Else
addone pos2, 14
a$ = Mid$(a$, 2)
    pos2 = K - Len(a$) + 1
    GoTo again

 End If
If pos2 > 0 Then addone pos2, 14


End Sub
Private Function blockString(s$) As String
Dim i As Long, j As Long, c As Long
Dim a1 As Boolean

If s$ = vbNullString Then Exit Function
If mMultiLineComment2 <> "}" Then
i = InStr(s$, mMultiLineComment2)
If i > 0 Then blockString = Mid$(s$, 1, i - 1): s$ = Mid$(s$, i)
Exit Function
Else

c = Len(s$)
a1 = True
i = 1
Do
Select Case Asc(Mid$(s$, i, 1))
Case 34
Do While i < c
i = i + 1
If Asc(Mid$(s$, i, 1)) = 34 Then If Asc(Mid$(s$, i - 1, 1)) <> 92 Then Exit Do
Loop
Case 123
j = j - 1
Case 125
j = j + 1: If j = 1 Then Exit Do
End Select
i = i + 1
Loop Until i > c
If j = 1 Then
blockString = Myleft$(s$, i - 1)
s$ = Mid$(s$, i)
Else
blockString = s$
s$ = vbNullString
End If
End If
End Function
Private Function blockString2(s$) As String
Dim i As Long, j As Long, c As Long

If s$ = vbNullString Then Exit Function

c = Len(s$)
i = 0
Do While i < c
i = i + 1
If Asc(Mid$(s$, i, 1)) = 34 Then
If i = 1 Then Exit Do
If Asc(Mid$(s$, i - 1, 1)) <> 92 Then Exit Do
End If
Loop
If i = 1 Then
s$ = Mid$(s$, 2)
blockString2 = Chr$(34) + Chr$(34)
ElseIf i < c Then
blockString2 = Chr$(34) + Myleft$(s$, i)
s$ = Mid$(s$, i + 1)
Else
blockString2 = Chr$(34) + s$
s$ = vbNullString
End If
End Function
Private Function blockString3(s$, ssep As String) As String
Dim i As Long, j As Long, c As Long, Add As Long
Add = Len(ssep)

If s$ = vbNullString Then Exit Function

c = Len(s$)
i = 0
Do While i < c
i = i + 1
If Mid$(s$, i, Add) = ssep Then
If i = Add Then Exit Do
If Asc(Mid$(s$, i - 1, 1)) <> 92 Then Exit Do
End If
Loop
If i = 1 Then
s$ = Mid$(s$, 1 + Add)
blockString3 = ssep + ssep
ElseIf i < c Then
blockString3 = ssep + Mid$(s$, 1, i + Add - 1)
s$ = Mid$(s$, i + Add)
Else
blockString3 = ssep + s$
s$ = vbNullString
End If
End Function
Private Function IsLabelDot(where$, a$, r$) As Long ' ok
' for left side...no &

Dim rr&, one As Boolean, c$, firstdot$, gr As Boolean
r$ = vbNullString
If a$ = vbNullString Then IsLabelDot = 0: Exit Function

a$ = NLtrim$(a$)
again:
    Do While Len(a$) > 0
    c$ = Myleft$(a$, 1)
    If Len(DropSym) > 0 Then
    If InStr(DropSym, c$) > 0 Then
        If c$ = "/" Then
            If Len(r$) > 0 Then Exit Do
            rr& = 500
            Exit Do
        Else
            c$ = "A"
        End If
    End If
    End If
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 46 '"."
            If one Then
            Exit Do
            ElseIf r$ <> "" And Len(a$) > 1 Then
            If Mid$(a$, 2, 2) = ". " Or Mid$(a$, 2, 1) = " " Then Exit Do
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1
            Else
            firstdot$ = firstdot$ + "."
            a$ = Mid$(a$, 2)
            End If
       Case 64, 91 To 94, 123 To 126
       
        Exit Do

        Case 48 To 57
           If one Then
            If firstdot$ <> "" Then a$ = firstdot$ + a$
            Exit Do
            ElseIf r$ <> "" Then
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1
            Else
            Exit Do
            End If
        Case Is < 0, Is > 64
            If one Then
            Exit Do
            Else
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1
            End If
        Case 36 ' "$"
            If one Then Exit Do
            If r$ <> "" Then
            one = True
            rr& = 1
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 37 ' "%"
            If one Then Exit Do
            If r$ <> "" Then
            one = True
            rr& = 1
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 40 ' "("
            If r$ <> "" And Not WordOnly Then
                                       Select Case rr&
                                       Case 1, 3, 4
                                       rr& = 5
                                       Case Else
                                       Exit Do
                                       End Select
                                       r$ = r$ & Myleft$(a$, 1)
                                       a$ = Mid$(a$, 2)
                                   Exit Do
                            
              '            End If
               Else
                        Exit Do
            
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
            gr = True
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        End If

    Loop
    If mUseCase Then
       r$ = firstdot$ + myUcase(r$, gr)
       Else
       r$ = firstdot$ + r$
       End If
    IsLabelDot = rr&
   'a$ = LTrim$(a$)

End Function
Function IsLabelDot2000(where$, a$, r$) As Long 'ok
' for left side...no &

Dim rr&, one As Boolean, c$, firstdot$, gr As Boolean
r$ = vbNullString
If a$ = vbNullString Then IsLabelDot2000 = 0: Exit Function

a$ = NLtrim$(a$)
    Do While Len(a$) > 0
    c$ = Myleft$(a$, 1)
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 64  '"@"
        
            If Mid$(a$, 2, 1) <> "(" And r$ <> "" Then

                r$ = r$ + "@"
                a$ = Mid$(a$, 2)
            Else
                IsLabelDot2000 = 0: a$ = firstdot$ + a$: Exit Function
            End If
        Case 63 '"?"
        If r$ = vbNullString And firstdot$ = vbNullString Then
        r$ = "?"
        a$ = Mid$(a$, 2)
        IsLabelDot2000 = 1
        Exit Function
    
        ElseIf firstdot$ = vbNullString Then
        IsLabelDot2000 = 1
        Exit Function
        Else
        IsLabelDot2000 = 0
        Exit Function
        End If
        Case 46 '"."
            If one Then
            Exit Do
            Exit Do
            ElseIf r$ <> "" And Len(a$) > 1 Then
            If Mid$(a$, 2, 2) = ". " Or Mid$(a$, 2, 1) = " " Then Exit Do
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1
            Else
            firstdot$ = firstdot$ + "."
            a$ = Mid$(a$, 2)
            End If
       Case 92, 94, 123 To 126, 160 '"\","^", "{" To "~"
        Exit Do

        Case 48 To 57, 95 '"0" To "9", "_"
           If one Then
            If firstdot$ <> "" Then a$ = firstdot$ + a$
            Exit Do
            ElseIf r$ <> "" Then
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            Else
            Exit Do
            End If
        Case Is < 0, Is > 64 ' >=A and negative
            If one Then
            Exit Do
            Else
                If Myleft$(r$, 1) <> "[" And c$ = "]" Then
                    If InStr(r$, ".[") > 0 Then
                        r$ = r$ + "]"
                        a$ = Mid$(a$, 2)
                    Else
                        Exit Do
                    End If
                ElseIf Myleft$(r$, 1) = "[" And c$ = "]" And Mid$(a$, 2, 1) = "." Then
                    r$ = r$ & "]"
                    a$ = Mid$(a$, 2)
                    Exit Do
                ElseIf Right$(r$, 2) = "]$" And c$ = "]" Then
                    Exit Do
                ElseIf Right$(r$, 1) = "]" And c$ = "]" Then
                    Exit Do
                ElseIf c$ = "[" And Len(r$) > 0 Then
                    r$ = r$ & "["
                    a$ = Mid$(a$, 2)
                    If Right$(r$, 2) <> ".[" Then
                        If Myleft$(r$, 1) <> "[" Then
                            rr& = 8
                            Exit Do
                        End If
                    End If

                Else
                    r$ = r$ & Myleft$(a$, 1)
                    a$ = Mid$(a$, 2)
                End If
                rr& = 1 'is an identifier or floating point variable
            End If
        Case 36 ' "$"
        
            If one Then
            If InStr(r$, "$.") Then
            rr& = 3          ' is string variable with string pointer
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            End If
            Exit Do
            End If
            If r$ <> "" Then
            one = True
            If Mid$(a$, 2, 1) <> "." Then
            rr& = 3
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            rr& = 1       ' is string variable
            r$ = r$ & Myleft$(a$, 2)
            a$ = Mid$(a$, 3)
            
            End If
            
            Else
            Exit Do
            End If
        Case 37 ' "%"
            If one Then Exit Do
            If r$ <> "" Then
            one = True
            rr& = 4 ' is long variable
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 40 ' "("
            If r$ <> "" Then
                            If Mid$(a$, 2, 2) = ")@" Then
                                    r$ = r$ & "()."
                                  
                                 a$ = Mid$(a$, 4)
                               Else
                                       Select Case rr&
                                       Case 1
                                       rr& = 5 ' float array or function
                                       Case 3
                                       rr& = 6 'string array or function
                                       Case 4
                                       rr& = 7 ' long array
                                       Case Else
                                       Exit Do
                                       End Select
                                       r$ = r$ & Myleft$(a$, 1)
                                       a$ = Mid$(a$, 2)
                                   Exit Do
                            
                          End If
               Else
                        Exit Do
            
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
                If Myleft$(r$, 1) <> "[" And c$ = "]" Then
                    If InStr(r$, ".[") > 0 Then
                        r$ = r$ + "]"
                        a$ = Mid$(a$, 2)
                    Else
                        Exit Do
                    End If
                ElseIf Myleft$(r$, 1) = "[" And c$ = "]" And Mid$(a$, 2, 1) = "." Then
                    r$ = r$ & "]"
                    a$ = Mid$(a$, 2)
                    Exit Do
                ElseIf Right$(r$, 2) = "]$" And c$ = "]" Then
                    Exit Do
                ElseIf Right$(r$, 1) = "]" And c$ = "]" Then
                    Exit Do
                ElseIf c$ = "[" And Len(r$) > 0 Then
                    r$ = r$ & "["
                    a$ = Mid$(a$, 2)
                    If Right$(r$, 2) <> ".[" Then
                        If Myleft$(r$, 1) <> "[" Then
                            rr& = 8
                            Exit Do
                        End If
                    End If

                Else
                    r$ = r$ & Myleft$(a$, 1)
                    a$ = Mid$(a$, 2)
                End If
                rr& = 1 'is an identifier or floating point variable
            End If
        End If

    Loop
       r$ = firstdot$ + myUcase(r$, gr)
    IsLabelDot2000 = rr&
   'a$ = LTrim$(a$)

End Function


Private Function IsLabelOnly(a$, r$) As Long 'ok

Dim rr&, one As Boolean, c$, dot&, Extra&, ExtraYes As Boolean, use14 As Boolean
r$ = vbNullString
If a$ = vbNullString Then IsLabelOnly = 0: Exit Function
a$ = NLtrim$(a$)
again1:
    Do While Len(a$) > 0
     c$ = Myleft$(a$, 1) 'ANYCHAR HERE
        If Len(DropSym) > 0 Then
            If InStr(DropSym, c$) > 0 Then
                c$ = "A"
            End If
        End If
    
    If OtherSymbols <> "" Then
    If InStr(OtherSymbols, c$) > 0 Then a$ = Mid$(a$, 2): rr& = 10: Extra& = 0: Exit Do
    End If
     If stringsep2 <> Chr(0) Then
    If InStr(stringsep2, c$) > 0 Then Exit Do
    End If
    If ExtraFront <> "" And Not use14 Then
    If ExtraYes Then
    If IsmineHere3(r$) Then
    use14 = True
    End If
    ElseIf LenB(r$) = 0 Then
    If InStr(ExtraFront, c$) > 0 Then
        r$ = c$
        a$ = Mid$(a$, 2)
        one = False
        rr& = 1
        ExtraYes = True
        GoTo again1
    End If
    End If
    End If
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        
        Case 46 '"."
            If one Then
                Exit Do
            ElseIf r$ <> "" Then
                r$ = r$ & Myleft$(a$, 1)
                a$ = Mid$(a$, 2)
           ElseIf Not Mid$(a$, 2, 1) Like "[0-9]" Then
                       If r$ <> "" Then
                           r$ = r$ & Myleft$(a$, 1)
                          rr& = 1
                                      
                        Else
                            
                        dot& = dot& + 1
                    
                        End If
                        If Mid$(a$, 2, 1) Like "[eEεΕ]" Then
                            r$ = r$ & Mid$(a$, 2, 1)
                             a$ = Mid$(a$, 3)
                        Else
                             a$ = Mid$(a$, 2)
                        End If
            Else
                If r$ = vbNullString And dot& > 0 Then
                    r$ = String$(dot& + 1, ".")
                    a$ = Mid$(a$, 2)
                   IsLabelOnly = 1
                Else
                    If MaybeIsSymbolNoSpace(Mid$(a$, 2, 1), "1234567890") Then
                        FastSymbol a$, "."
                        While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                        Wend
                        If Len(OpenTag) = 0 Then
                        If MaybeIsSymbolNoSpace(Mid$(a$, 3, 1), "EeεΕ") Then
                            a$ = Mid$(a$, 4)
                            MaybeIsSymbolNoSpace a$, "+-", True
                            MaybeIsSymbolNoSpace a$, "123456789", True
                            While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                            Wend
                            If MaybeIsSymbolNoSpace(a$, ".") Then
                                a$ = vbNullString
                            End If
                        End If
                        End If
                        If MaybeIsSymbolNoSpace(a$, ".") Then
                    'a$ = vbnullstring
                            IsLabelOnly = 1
                        Else
                            IsLabelOnly = 100
                        End If
                        Exit Function
                    End If
                    IsLabelOnly = 0
                End If
                Exit Function
            End If
     Case 35
        If r$ <> "" Or Not UseHtmlColors Then
            Exit Do
        Else
        
            r$ = "#"
            a$ = Mid$(a$, 2)
            rr& = 10
        End If
     Case 64, 91 To 94, 123 To 126, 160  '"\","^", "{" To "~"
     
Exit Do
        
        Case 48 To 57, 95 '"0" To "9", "_"
       
           If one Then
           
            Exit Do
            ElseIf Extra& > 0 Then
                
                r$ = r$ & Myleft$(a$, 1)
                a$ = Mid$(a$, 2)
                
                If Len(r$) > 18 Then rr& = 0: Exit Do
                rr& = 1
            If MaybeIsSymbolNoSpace(a$, "&", True) Then
                If Len(r$) > 10 Then rr& = 0: Exit Do
                r$ = r$ + "&": Exit Do
            ElseIf MaybeIsSymbolNoSpace(a$, "%", True) Then
                If Len(r$) > 6 Then rr& = 0: Exit Do
                r$ = r$ + "%": Exit Do
            End If
                
            ElseIf r$ <> "" Then
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
               rr& = 1 'is an identifier or floating point variable
            ElseIf Mid$(a$ + "__", 2, 2) Like "[ΧχXx][0-9A-Fa-f]" Then
evryka:
                r$ = r$ & Myleft$(a$, 3)
                a$ = Mid$(a$, 4)
                Extra& = 9
                If Myleft$(a$, 1) = "%" Then
                    a$ = Mid$(a$, 2)
                    rr& = 1
                    Exit Do
                ElseIf Myleft$(a$, 1) = "&" Then
                    rr& = 1
                    a$ = Mid$(a$, 2)
                    Exit Do
                Else
                    rr& = 1
                End If
            ElseIf Mid$(a$, 2, 1) = "%" Then
                a$ = Mid$(a$, 3)
                rr& = 100: Exit Do
            ElseIf Mid$(a$, 2, 1) = "&" Then
                a$ = Mid$(a$, 3)
                rr& = 100: Exit Do
            Else
                If dot& > 0 Then
                    a$ = "." + a$: dot& = 0
                ElseIf Mid$(a$, 2, 1) = "." Then
                    a$ = Mid$(a$, 3)
checkexp:
              While MaybeIsSymbolNoSpace(a$, "1234567890", True)
              Wend
If Len(OpenTag) = 0 Then
              If MaybeIsSymbolNoSpace(a$, "EeΕε", True) Then
                  MaybeIsSymbolNoSpace a$, "+-", True
                  If MaybeIsSymbolNoSpace(a$, "+-", True) Then
                  IsLabelOnly = 1
                  Exit Function
                  End If
                  MaybeIsSymbolNoSpace a$, "123456789", True
                  While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                  Wend
                  If MaybeIsSymbolNoSpace(a$, ".") Then
                  IsLabelOnly = 1
                  a$ = vbNullString
                  Exit Function
                  End If
              Else
              While MaybeIsSymbolNoSpace(a$, "1234567890", True)
              Wend
              End If
              End If
            If MaybeIsSymbolNoSpace(a$, ".") Then
             
              IsLabelOnly = 1
              Else
              If Len(OpenTag) > 0 Then IsLabelOnly = 1 Else IsLabelOnly = 100
              End If
              
              Exit Function
            ElseIf MaybeIsSymbolNoSpace(a$, "1234567890", True) Then
              While MaybeIsSymbolNoSpace(a$, "1234567890", True)
              Wend
                  If MaybeIsSymbolNoSpace(a$, ".", True) Then
                                
                    GoTo checkexp
                  End If
                  
                  If Len(OpenTag) > 0 Then
                  IsLabelOnly = 1
                  Else
                  MaybeIsSymbolNoSpace a$, "&", True
                  IsLabelOnly = 100
                  End If
                Exit Function
            End If
            Exit Do
            End If
        Case Is < 0, Is > 64 ' >=A and negative

            If one Then
            Exit Do
            Else
                If Extra& > 0 Then
                If Not c$ Like "[0-9a-fA-F]" Then Exit Do
                If Len(r$) > 18 Then rr& = 0: a$ = Mid$(a$, 2): Exit Do
                End If
            
              r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            If MaybeIsSymbolNoSpace(a$, "&", True) Then
                If Len(r$) > 10 Then rr& = 0: Exit Do
                r$ = r$ + "&": Exit Do
            ElseIf MaybeIsSymbolNoSpace(a$, "%", True) Then
                If Len(r$) > 6 Then rr& = 0: Exit Do
                r$ = r$ + "%": Exit Do
            End If
            rr& = 1
  
            End If
       Case 36 ' "$"
       If one Then Exit Do
            If r$ <> "" Then
            
            rr& = 3 ' is string variable
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 37 ' "%"
            If one Then Exit Do
            If r$ <> "" Then
            one = True
            rr& = 1 ' is long variable but here we want to be as numeric variable
            r$ = r$ & Myleft$(a$, 1)
        
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 38
            If one Then Exit Do
            If LCase$(Mid$(a$, 2, 1)) = "h" Then
                If Mid$(a$ + "__", 3, 1) Like "[0-9A-Fa-f]" Then
                    GoTo evryka
                End If
            End If
            a$ = Mid$(a$, 2)
            rr& = 100
            Exit Do
        
        Case 40 ' "("
            If r$ <> "" And Not ((Len(r$) = 1 And ExtraYes) Or WordOnly) Then
                Select Case rr&
                Case 1, 3, 4
                rr& = 5
                Case Else
                Exit Do
                End Select
                r$ = r$ & Myleft$(a$, 1)
                a$ = Mid$(a$, 2)
               Exit Do
            
            'End If
           Else
           Exit Do
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
             r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        End If
    Loop
    ' SPECIFIC M2000 - FOR HTML COLORS
    If rr& > 0 Then
        If Myleft$(r$, 1) = "#" Then
          If Len(r$) > 1 Then
            If Mid$(r$, 2) Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                rr& = 10
            Else
                a$ = r$ + a$
                rr& = 0
            End If
         Else
            a$ = r$ + a$
             rr& = 0
        End If
        End If
    End If
If ExtraYes And Len(r$) = 1 Then
   If AscW(r$) < 65 Then
        a$ = r$ + a$
        r$ = vbNullString
        IsLabelOnly = 0
    Else
         IsLabelOnly = rr&
    End If
Else
If use14 Then
    IsLabelOnly = 10
Else
    IsLabelOnly = rr& + Extra&
    End If
    End If
 'If IsLabelOnly = 1 And r$ = "" Then Stop
End Function
Function FastSymbol(a$, c$, Optional mis As Boolean = False, Optional cl As Long = 1) As Boolean
Dim i As Long, j As Long
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function  ' this is not good
If j - i < cl - 1 Then
If mis Then MyEr "missing " & c$, "λείπει " & c$
Exit Function
End If
If InStr(c$, Mid$(a$, i, cl)) > 0 Then
a$ = Mid$(a$, cl + i)
FastSymbol = True
ElseIf mis Then
MyEr "missing " & c$, "λείπει " & c$
End If
End Function
Function FastSymbolNoRemove(a$, c$, Optional mis As Boolean = False, Optional cl As Long = 1) As Boolean
Dim i As Long, j As Long
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function  ' this is not good
If j - i < cl - 1 Then Exit Function
If InStr(c$, Mid$(a$, i, cl)) > 0 Then FastSymbolNoRemove = True
End Function
Function IsLabelOnly2000(a$, r$) As Long 'ok
Dim rr&, one As Boolean, c$, dot&, Extra&, num As Boolean
r$ = vbNullString
If a$ = vbNullString Then IsLabelOnly2000 = 0: Exit Function
a$ = NLtrim$(a$)
Do While Len(a$) > 0
    c$ = Myleft$(a$, 1) 'ANYCHAR HERE
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 35 ' "#"
            If r$ <> "" Then
                Exit Do
            Else
                r$ = "#"
                a$ = Mid$(a$, 2)
                rr& = 10
           End If
        Case 64  '"@"
           If r$ = vbNullString Then
              a$ = Mid$(a$, 2)
              ElseIf Mid$(a$, 2, 1) <> "(" And r$ <> "" Then
            r$ = r$ & "."
              a$ = Mid$(a$, 2)
              Else
                 IsLabelOnly2000 = 0: Exit Function
            End If
        Case 46 '"."
            If one Then
                Exit Do
            ElseIf r$ <> "" Then
                r$ = r$ & Myleft$(a$, 1)
                a$ = Mid$(a$, 2)
            ElseIf Not Mid$(a$, 2, 1) Like "[0-9]" Then
                       If r$ <> "" Then
                            r$ = r$ & Myleft$(a$, 1)
                            rr& = 1
                        Else
                            dot& = dot& + 1
                        End If
                        If Mid$(a$, 2, 1) Like "[eEεΕ]" Then
                            r$ = r$ & Mid$(a$, 2, 1)
                            a$ = Mid$(a$, 3)
                        Else
                             a$ = Mid$(a$, 2)
                        End If
            Else
                If r$ = vbNullString And dot& > 0 Then
                    r$ = String$(dot& + 1, ".")
                    a$ = Mid$(a$, 2)
                    IsLabelOnly2000 = 1
                Else
                If MaybeIsSymbolNoSpace(Mid$(a$, 2, 1), "1234567890") Then
                    FastSymbol a$, "."
                    While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                    Wend
                    If MaybeIsSymbolNoSpace(Mid$(a$, 3, 1), "EeΕε") Then
                        a$ = Mid$(a$, 4)
                        MaybeIsSymbolNoSpace a$, "+-", True
                        MaybeIsSymbolNoSpace a$, "123456789", True
                        While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                        Wend
                        If MaybeIsSymbolNoSpace(a$, ".") Then
                            a$ = vbNullString
                        End If
                    End If
                    If MaybeIsSymbolNoSpace(a$, ".") Then
                        IsLabelOnly2000 = 1
                    Else
                        IsLabelOnly2000 = 100
                    End If
                    Exit Function
                End If
                IsLabelOnly2000 = 0
            End If
            Exit Function
        End If
    Case 92, 94, 123 To 126, 160 '"\","^", "{" To "~"
        Exit Do
    Case 48 To 57, 95 '"0" To "9", "_"
        If one Then
            Exit Do
        ElseIf r$ <> "" Then
            If c$ = "_" Then
            r$ = " " & r$
            Else
            r$ = r$ & Myleft$(a$, 1)
            End If
            a$ = Mid$(a$, 2)
            If Extra& > 0 Then
                If Myleft$(a$, 1) = "%" Then
                    a$ = Mid$(a$, 2)
                    If Len(LTrim$(r$)) > 6 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                    Exit Do
                ElseIf Myleft$(a$, 2) = "&&" Then
                    a$ = Mid$(a$, 3)
                    If Len(LTrim$(r$)) > 18 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                    Exit Do
                ElseIf Myleft$(a$, 1) = "&" Then
                    a$ = Mid$(a$, 2)
                    If Len(LTrim$(r$)) > 10 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                    Exit Do
                ElseIf Myleft$(a$, 1) = "@" Then
                    a$ = Mid$(a$, 2)
                    If Len(LTrim$(r$)) > 18 Then rr& = 0 Else rr& = 1
                    Exit Do
                ElseIf LCase$(Myleft$(a$, 2)) = "ub" Then
                    a$ = Mid$(a$, 3)
                    If Len(LTrim$(r$)) > 4 Then rr& = 0 Else rr& = 1
                    Exit Do
                End If
            End If
            If Len(LTrim$(r$)) > 18 And Extra& Then rr& = 0 Else rr& = 1
        ElseIf Mid$(a$ + "__", 2, 2) Like "[ΧχXx][0-9A-Fa-f]" Then
evryka:
            r$ = Myleft$(a$, 3)
            a$ = Mid$(a$, 4)
            Extra& = 9
            If Myleft$(a$, 1) = "%" Then
                a$ = Mid$(a$, 2)
                rr& = 1: Extra& = Extra& + 10
                Exit Do
            ElseIf Myleft$(a$, 2) = "&&" Then
                Extra& = Extra& + 10
                a$ = Mid$(a$, 3)
                Exit Do
            ElseIf Myleft$(a$, 1) = "&" Then
                rr& = 1: Extra& = Extra& + 10
                a$ = Mid$(a$, 2)
                Exit Do
            ElseIf Myleft$(a$, 1) = "@" Then
                a$ = Mid$(a$, 2)
                rr& = 1
                Exit Do
            ElseIf LCase$(Myleft$(a$, 2)) = "ub" Then
                a$ = Mid$(a$, 3)
                rr& = 1
                Exit Do
            Else
            rr& = 1
            End If
        ElseIf Mid$(a$, 2, 1) = "%" Then
            a$ = Mid$(a$, 3)
            
            rr& = 100: Exit Do
        ElseIf Mid$(a$, 2, 1) = "&" Then
           a$ = Mid$(a$, 3)
            
            MaybeIsSymbolNoSpace a$, "&", True
            rr& = 100: Exit Do
            
            Exit Do
        Else
            If dot& > 0 Then
                a$ = "." + a$: dot& = 0
            ElseIf Mid$(a$, 2, 1) = "." Then
                a$ = Mid$(a$, 3)
checkexp:
                While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                Wend
                If MaybeIsSymbolNoSpace(a$, "EeΕε", True) Then
                    MaybeIsSymbolNoSpace a$, "+-", True
                    If MaybeIsSymbolNoSpace(a$, "+-", True) Then
                        IsLabelOnly2000 = 1
                        Exit Function
                    End If
                    MaybeIsSymbolNoSpace a$, "123456789", True
                    While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                    Wend
                    If MaybeIsSymbolNoSpace(a$, ".") Then
                      IsLabelOnly2000 = 1
                      a$ = vbNullString
                      Exit Function
                    End If
                Else
                    While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                    Wend
                    If MaybeIsSymbolNoSpace(a$, "uU", True) Then
                    If MaybeIsSymbolNoSpace(a$, "dD", True) Then
                    End If
                    End If
                End If
                If MaybeIsSymbolNoSpace(a$, ".") Then
                    IsLabelOnly2000 = 1
                Else
                    IsLabelOnly2000 = 100
                End If
                Exit Function
            ElseIf MaybeIsSymbolNoSpace(a$, "1234567890", True) Then
                While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                Wend
                If MaybeIsSymbolNoSpace(a$, ".", True) Then
                    GoTo checkexp
                ElseIf MaybeIsSymbolNoSpace(a$, "EeΕε") Then
                    GoTo checkexp
                End If
                If MaybeIsSymbolNoSpace(a$, "&", True) Then
                    MaybeIsSymbolNoSpace a$, "&", True
                
                ElseIf MaybeIsSymbolNoSpace(a$, "%", True) Then
                ElseIf MaybeIsSymbolNoSpace(a$, "uU", True) Then
                    If MaybeIsSymbolNoSpace(a$, "bB", True) Then
                    ElseIf MaybeIsSymbolNoSpace(a$, "dD", True) Then
                    End If
                End If
                    
                IsLabelOnly2000 = 100
                Exit Function
            End If
            Exit Do
        End If
    Case Is < 0, Is > 64 ' >=A and negative
        If one Then
            Exit Do
        Else
            If Extra& > 0 Then
                If Not c$ Like "[a-fA-F]" Then Exit Do
                r$ = r$ & Myleft$(a$, 1)
                a$ = Mid$(a$, 2)
                If Myleft$(a$, 1) = "%" Then
                    If Len(LTrim$(r$)) > 6 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                    a$ = Mid$(a$, 2)
                    Exit Do
                ElseIf Myleft$(a$, 2) = "&&" Then
                   a$ = Mid$(a$, 3)
                   If Len(LTrim$(r$)) > 18 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                   Exit Do
                ElseIf Myleft$(a$, 1) = "&" Then
                   a$ = Mid$(a$, 2)
                   If Len(LTrim$(r$)) > 10 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                   Exit Do
                ElseIf Myleft$(a$, 1) = "@" Then
                    a$ = Mid$(a$, 2)
                    If Len(LTrim$(r$)) > 18 Then rr& = 0 Else rr& = 1
                    Exit Do
                End If
                If Len(LTrim$(r$)) > 18 Then rr& = 0
            Else
                If Myleft$(r$, 1) <> "[" And c$ = "]" Then
                    If InStr(r$, ".[") > 0 Then
                        r$ = r$ + "]"
                        a$ = Mid$(a$, 2)
                    Else
                        Exit Do
                    End If
                ElseIf Myleft$(r$, 1) = "[" And c$ = "]" And Mid$(a$, 2, 1) = "." Then
                    r$ = r$ & "]"
                    a$ = Mid$(a$, 2)
                    Exit Do
                ElseIf Right$(r$, 2) = "]$" And c$ = "]" Then
                    Exit Do
                ElseIf Right$(r$, 1) = "]" And c$ = "]" Then
                    Exit Do
                ElseIf c$ = "[" And Len(r$) > 0 Then
                    r$ = r$ & "["
                    a$ = Mid$(a$, 2)
                    If Right$(r$, 2) <> ".[" Then
                        If Myleft$(r$, 1) <> "[" Then
                            If dot& > 0 Then
                                r$ = String$(dot&, ".") + r$
                            End If
                            rr& = 8
                            Exit Do
                        End If
                    End If

                Else
                    r$ = r$ & Myleft$(a$, 1)
                    a$ = Mid$(a$, 2)
                End If
                rr& = 1 'is an identifier or floating point variable
            End If
        End If
    Case 36 ' "$"
       If one Then Exit Do
            If r$ <> "" Then
            one = True
            rr& = 3 ' is string variable
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
     Case 37 ' "%"
      If one Then
            If InStr(r$, "$.") Then
            rr& = 3          ' is string variable with string pointer
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            End If
            Exit Do
            End If
            If r$ <> "" Then
            one = True
            If Mid$(a$, 2, 1) <> "." Then
            rr& = 3
            r$ = r$ & Myleft$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            rr& = 1       ' is string variable
            r$ = r$ & Myleft$(a$, 2)
            a$ = Mid$(a$, 3)
            
            End If
            
            Else
                Exit Do
            End If
        Case 38 ' "&"
        If one Then Exit Do
        If LCase$(Mid$(a$, 2, 1)) = "h" Then
        If Mid$(a$ + "__", 3, 1) Like "[0-9A-Fa-f]" Then
            GoTo evryka
        End If
        End If
        If r$ = vbNullString Then rr& = 10: a$ = Mid$(a$, 2): r$ = "&"
        Exit Do
        Case 40 ' "("
            If r$ <> "" Then
                If Mid$(a$, 2, 2) = ")@" Then
                 r$ = r$ & "()."
              a$ = Mid$(a$, 4)
                Else
                Select Case rr&
                Case 1
                rr& = 5 ' float array or function
                Case 3
                rr& = 6 'string array or function
                Case 4
                rr& = 7 ' long array
                Case Else
                Exit Do
                End Select
                r$ = r$ & Myleft$(a$, 1)
                a$ = Mid$(a$, 2)
               Exit Do
            
            End If
           Else
           Exit Do
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
                If Myleft$(r$, 1) <> "[" And c$ = "]" Then
                    If InStr(r$, ".[") > 0 Then
                        r$ = r$ + "]"
                        a$ = Mid$(a$, 2)
                    Else
                        Exit Do
                    End If
                ElseIf Myleft$(r$, 1) = "[" And c$ = "]" And Mid$(a$, 2, 1) = "." Then
                    r$ = r$ & "]"
                    a$ = Mid$(a$, 2)
                    Exit Do
                ElseIf Right$(r$, 2) = "]$" And c$ = "]" Then
                    Exit Do
                ElseIf Right$(r$, 1) = "]" And c$ = "]" Then
                    Exit Do
                ElseIf c$ = "[" And Len(r$) > 0 Then
                    r$ = r$ & "["
                    a$ = Mid$(a$, 2)
                    If Right$(r$, 2) <> ".[" Then
                        If Myleft$(r$, 1) <> "[" Then
                            If dot& > 0 Then
                                r$ = String$(dot&, ".") + r$
                            End If
                            rr& = 8
                            Exit Do
                        End If
                    End If

                Else
                    r$ = r$ & Myleft$(a$, 1)
                    a$ = Mid$(a$, 2)
                End If
                rr& = 1 'is an identifier or floating point variable
            End If
        End If
    Loop
    If rr& > 0 Then
        If Myleft$(r$, 1) = "#" Then
          If Len(r$) > 1 Then
            If Mid$(r$, 2) Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                rr& = 10
            Else
                a$ = r$ + a$
                rr& = 0
            End If
         Else
        a$ = r$ + a$
        rr& = 0
        End If
        End If
    End If
    
    If r$ = "&" Then
        IsLabelOnly2000 = 10
    ElseIf rr& > 3 And rr& < 9 Then
        IsLabelOnly2000 = rr& + Extra&
    ElseIf CheckFree(Myleft$(a$, 1)) Then
        IsLabelOnly2000 = rr& + Extra&
    ElseIf rr& > 0 Then
        If Myleft$(r$, 1) = "&" Then
            a$ = Mid$(r$, 2) + a$: r$ = "&"
            IsLabelOnly2000 = 10
        Else
            IsLabelOnly2000 = 1
        End If
    End If
End Function



Private Function MaybeIsSymbol(a$, c$, Optional tail$ = vbNullString, Optional Remove As Boolean = False) As Boolean
Dim i As Long, ok As Boolean
If a$ = vbNullString Or c$ = vbNullString Then Exit Function
i = 1
While Mid$(a$, i, 1) = " "
    i = i + 1
Wend
If i > Len(a$) Then Exit Function
ok = InStr(c$, Mid$(a$, i, 1)) > 0
If tail$ <> "" And ok Then
ok = Mid$(a$, i + 1, Len(tail$)) Like tail$
End If
If Remove Then If ok Then a$ = Mid$(a$, i + 1 + Len(tail$))
MaybeIsSymbol = ok
End Function
Private Function MaybeIsSymbol2(a$, num As Long, c$, Optional num2 As Long, Optional tail$ = vbNullString, Optional Remove As Boolean = False) As Boolean
If num <= 0 Then Exit Function
Dim i As Long, ok As Boolean
If a$ = vbNullString Or c$ = vbNullString Then Exit Function
i = 1
While Mid$(a$, i, 1) = " "
    i = i + 1
Wend
If i + num - 1 > Len(a$) Then Exit Function
''If i > num Then Exit Function
ok = Mid$(a$, i, num) Like c$
If num2 > 0 And ok Then
ok = Mid$(a$, i + num, num2) Like tail$  '' tail$="[abc]" for one num2
End If
If Remove Then If ok Then a$ = Mid$(a$, i + num + num2)
MaybeIsSymbol2 = ok
End Function

Private Function MaybeIsSymbolNoSpace(a$, c$, Optional Remove As Boolean = False) As Boolean
Dim ok As Boolean

If a$ = vbNullString Then Exit Function
ok = InStr(c$, Myleft$(a$, 1))
If Remove And ok Then a$ = Mid$(a$, 2)
MaybeIsSymbolNoSpace = ok
End Function
Public Property Let SetM2000(RHS)
    M2000code = CBool(RHS)
    mDoc.ForEditBox = Not CBool(RHS)
    If M2000code Then
    EnablePairs = True
    AutoNumber = True
    AllowInsertBrackets = True
    glistN.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", "@", Chr$(9), "#", "%", "&")
    glistN.WordCharRight = ConCat(".", ":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", Chr$(9), "#")
    glistN.WordCharRightButIncluded = "(" ' so aaa(sdd) give aaa( as word
    glistN.WordCharLeftButIncluded = "#"
    
   End If
   EditTextWord = Not M2000code
End Property
Public Property Let SetHTML(RHS)
M2000code = False
mDoc.ForEditBox = Not CBool(RHS)
If RHS Then
EnablePairs = True
AllowInsertBrackets = True
NoColor = False
UseCase = True
ExtraFront = "_"
LineComment1 = ""
LineComment2 = ""
CommentSymbols = "[<][!]"
ComSymbolsWidth = 2
CommentLineLight = True
MultiLineComment1 = "<!--"
MultiLineComment2 = "-->"
stringsep2 = "'"
WordCharLeft = " .,+=*^()/[]:;!&|><'?{}@" + Chr$(34)
WordCharRight = " ,+=*^()[];<>!&|#?'{}" + Chr$(34)
WordCharLeftButIncluded = "&!"
OtherSymbols = Chr$(0)
OpenTag = "<"
CloseTag = "</"
UseHtmlColors = True
DropSym = "!#&@1234567890:.&;-/(){}" + Chr$(34)
AssignSym = "_"
StartSymbols = "<>"
TabWidth = 2
EditTextWord = False
SpaceIndent = 2

End If
End Property
Private Sub mDoc_ColorDataEvent(thispara1 As Long, Data As String, feedback As Long)
If M2000code Then
    mDoc_ColorDataEvent2 thispara1, Data, feedback
Else
    mDoc_ColorDataEvent1 thispara1, Data, feedback
End If
End Sub
Private Sub mDoc_ColorDataEvent2(thispara1 As Long, Data As String, feedback As Long)

thispara = thispara1
Col() = mDoc.Paracolordata(thispara)
If NoColor Then
If UBound(Col()) > 1 Then
ReDim Col(1)
mDoc.Paracolordata(thispara) = Col()
End If
Exit Sub
End If
If feedback = 0 Then
ReDim Col(1)
End If
lastdata = Data
ColorizeInnerStatements (Data), Len(Data), (0), feedback
If feedback = -1 Then
If thispara <> 0 Then
mDoc.Paracolordata(thispara) = Col()
thispara1 = thispara
End If
Else
mDoc.Paracolordata(thispara) = Col()
thispara1 = thispara
End If
End Sub
Private Sub ColorizeInnerStatements(a$, K As Long, pos2 As Long, feedback As Long, Optional jump As Boolean = False)
Dim drop$, pos1 As Long, onetime As Boolean, kind As Long, lastkind As Long, feedthis As Long, counter As Long
Dim lastcol As Long
a$ = Replace(a$, vbTab, " ")
If jump Then
kind = IsLabelOnly2000((a$), drop$)
If kind = 0 Then
addone2000 pos2, 14
lastcol = 14
If FastSymbol(a$, "(") Then
 pos2 = K - Len(a$) + 1
End If
GoTo again2
End If
GoTo fromjump

End If
onetime = True

If K < 1 Then
makebase 14
mDoc.Paracolordata(thispara) = Col()
If mDoc.startfrom(thispara) = thispara Then
feedback = thispara
ReDim Col(1)
Do While mDoc.Advance(thispara)
If mDoc.startfrom(thispara) = feedback Then
mDoc.startfrom(thispara) = 0
mDoc.Paracolordata(thispara) = Col()
End If
Loop
feedback = feedback <> thispara
Exit Sub
End If

End If
If feedback > 0 Then
'lastcol = 15
a$ = Mid$(a$, feedback + 1)
'pos2 = feedback + 1
 addone2000 feedback + 1, lastcol
feedback = 0
GoTo checkthistoo
Else
lastcol = 14
End If
again:
If RTrim$(a$) = vbNullString Then
If pos2 > 0 Then
If Col((UBound(Col()) \ 2) * 2) <> lastcol Then
If pos2 < K Then
addone2000 pos2 + 1, lastcol
End If
ElseIf UBound(Col()) < 2 Then
addone2000 pos2, lastcol
End If
  Else
  makebase 14
  End If
mDoc.Paracolordata(thispara) = Col()
Exit Sub
' FOR TESTS
  pos1 = 0
  For K = 1 To UBound(Col()) Step 2
    If VarType(Col(K)) = 0 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
    If Col(K) > pos1 Then pos1 = Col(K)
    If Col(K) < pos1 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
  Next K
    For K = 0 To UBound(Col()) Step 2
   If Col(K) > 15 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
  Next K
  
End If
If FastSymbol(a$, "&") Then
            addone2000 K - Len(a$), 13
            pos2 = K - Len(a$) + 1
ElseIf FastSymbol(a$, "...", , 3) Then
            addone2000 K - Len(a$) - 2, 11
            pos2 = K - Len(a$) + 1
End If
If MaybeIsSymbol(a$, ":", "\") Then
a$ = Mid$(a$, 2)
If AscW(Myleft$(a$, 1) + " ") = 92 Then
If pos2 > 1 Then
    If Col(UBound(Col())) = 1 Or InStr(" #@.+-~^%$,!({'?|/&)}" + StringSep + stringsep2, Mid$(lastdata, pos2 - 1, 1)) > 0 Then

    If pos2 > 0 Then
        addone2000 pos2, 14
    Else
        addone2000 1, 14
    End If
 pos2 = K - Len(a$) + 1
GoTo novar
End If
End If
a$ = Mid$(a$, 2)
End If

If Len(a$) > 1 Then
If AscW(a$) <> 32 Then
lastkind = 1
GoTo again3

End If
Else
'lastkind = 1
'GoTo again3
End If
Else
novar:
While FastSymbol(a$, ":")
If pos2 > 0 Then
addone2000 pos2, 14
Else
addone2000 1, 14
End If
 pos2 = K - Len(a$) + 1
Wend
End If
If onetime And MaybeIsSymbol(a$, "(") Then
GoTo CheckThis
End If
drop$ = "@(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then
If IsSymbol(a$, "@") Then
        If pos2 > 0 Then
        addone2000 pos2, 14
        Else
        addone2000 1, 14
        End If
     pos2 = K - Len(a$) + 1
End If
End If
again500:
kind = IsLabelDot2000("", a$, drop$)

If kind = 0 Then drop$ = "$(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "~(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "->": kind = Abs(FastSymbol(a$, drop$, , 2))
fromjump:
If kind = 0 Then drop$ = vbNullString
If kind > 0 Then
    If ismine(drop$) Then
                lastcol = 11
                If pos2 > 0 Then
                   addone2000 pos2, 11
                Else
                   addone2000 K - Len(a$) - Len(drop$) + 1, 11
                   End If
        '' for the end
                If ismine2(drop$) Then
              
                kind = -2
                If Trim$(a$) <> "" Then
                                        pos2 = K - Len(a$) + 1
                                        If ismine22(drop$) Then
                                        
                                        ElseIf MaybeIsSymbol(a$, "{") Then
                                        
                                            addone2000 pos2, 14
                                            pos2 = K - Len(a$) + 1
                                            onetime = True
                                            GoTo again
                                        ElseIf ismine5(drop$) Then
                                            GoTo again500
                                        ElseIf ismine3(drop$) Then
                                            onetime = True
                                            GoTo again
                                        End If
                                End If
             
        
                ElseIf ismine1(drop$) Then
                  
                kind = -1
                                If Trim$(a$) <> "" Then
                                                  pos2 = K - Len(a$) + 1
                                        If FastSymbol(a$, "{") Then
                                    
                                                  addone2000 pos2, 14
                                              
                                                  pos2 = K - Len(a$) + 1
                                                  
                                                  onetime = True
                                                  GoTo again
                                          Else
                                           mDoc.Paracolordata(thispara) = Col()
                                                  Exit Sub
                                        End If
                                End If
                ElseIf ismine33(drop$) Then
                ' fix for case >100, =50...
                pos2 = K - Len(a$) + 1
                Do
                    lastcol = 14
                    addone2000 pos2, 14
                    If FastSymbol(a$, "<") Then
                        If Not FastSymbol(a$, "=") Then
                            FastSymbol a$, ">"
                        End If
                    ElseIf FastSymbol(a$, ">") Then
                        FastSymbol a$, "="
                    ElseIf FastSymbol(a$, "=") Then
                        FastSymbol a$, "="
                    ElseIf FastSymbol(a$, "~") Then
                    End If
                    pos1 = 1
                    drop$ = aheadstatus2000(a$, False, pos1)
                    If pos1 > 0 Then
                        drop$ = Mid$(a$, 1, pos1 - 1)
                        ColorizeExpressions2000 drop$, K - Len(a$) + Len(drop$), pos2, feedthis
                    If feedthis <> 15 Then
                    a$ = Mid$(a$, pos1)
                    pos2 = K - Len(a$) + 1
                    addone2000 pos2, 14
                Else
                    counter = 1
nextline1:
                    feedback = 0
                    drop$ = mDoc.NextTextParagraph(thispara, feedback)
                    If feedback = 0 Then
                        a$ = Mid$(a$, pos1)
                        mDoc.Paracolordata(thispara) = Col()
                        Exit Sub
                    End If
                    pos1 = 1
                    countbrakcet drop$, pos1, counter
                    mDoc.Paracolordata(thispara) = Col()
                    a$ = drop$
                    
                    K = Len(a$)
                    
                    If mDoc.startfrom(thispara) = 0 Then
                        mDoc.startfrom(thispara) = thispara
                    End If
                    mDoc.startfrom(feedback) = mDoc.startfrom(thispara)
                    thispara = feedback
                    lastdata = a$
                    pos2 = 1
                    If Len(a$) > 0 Then
                        ReDim Col(3)
                        Col(0) = 15
                        Col(1) = 1
                        Col(2) = 15
                        Col(3) = K + 1
                    Else
                        ReDim Col(1)
                        Col(0) = 15
                        Col(1) = 1
                    End If
                    If counter = 0 Then
                        If UBound(Col()) > 1 Then
                            Col(3) = pos1
                            If Col(3) = Col(1) Then
                                pos1 = pos1 + 1
                                Col(3) = pos1
                                Col(2) = 14
                            End If
                        Else
                            Col(1) = pos1
                        End If
                        If Myleft(a$, 1) <> "}" Then
                            a$ = Mid$(a$, pos1 + 1)
                            pos2 = K - Len(a$) + 1
                        Else
                            a$ = Mid$(a$, 2)
                            pos2 = K - Len(a$) + 2
                            addone2000 pos2, 14
                        End If
                        feedback = 0 'lastcol
                        feedthis = 0
                        If MaybeIsSymbol(a$, "<>=~") Then
                            onetime = False
                            pos2 = K - Len(a$) + 1
                            addone2000 pos2, 14
                            If FastSymbol(a$, "<") Then
                                If FastSymbol(a$, ">") Then
                                ElseIf FastSymbol(a$, "=") Then
                                End If
                            ElseIf FastSymbol(a$, ">") Then
                                If FastSymbol(a$, "=") Then
                                End If
                            ElseIf FastSymbol(a$, "=") Then
                                ElseIf FastSymbol(a$, "~") Then
                                End If
                            End If
                        Else
                            GoTo nextline1
                        End If
                    End If
                
                End If
                Loop Until Not FastSymbol(a$, ",")
                onetime = True
                
                GoTo here1
                End If
                 Else
            If kind = 8 Then
                lastcol = 13
            Else
                lastcol = 10
            End If
            If Len(a$) > 0 Then
                If AscW(a$) = 92 Then
                    lastkind = 1
                    End If
            End If
            If pos2 > 0 Then
                If kind = 8 Then
                   addone2000 K - Len(a$) - Len(drop$) + 1, 10
                   addone2000 K - Len(a$), 13
                Else
                    addone2000 pos2, 10
                End If
            Else
                lastkind = 0
                If kind = 8 Then
                    addone2000 K - Len(a$) - Len(drop$) + 1, 10
                    addone2000 K - Len(a$), 13
                Else
                    addone2000 K - Len(a$) - Len(drop$) + 1, 10
                End If
            End If
        End If
        pos2 = K - Len(a$) + 1
        If kind > 3 Then
            If kind = 8 Then
                lastcol = 13 '10
                    If Myleft$(a$, 1) = "]" Then
                    GoTo checkthistoo
                End If
            End If
CheckThis:
            pos1 = 1
            drop$ = aheadstatus2000(a$, False, pos1)
marvel:
            If pos1 > 1 Then
marvel2:
                drop$ = Mid$(a$, 1, pos1 - 1)
                ColorizeExpressions2000 drop$, K - Len(a$) + Len(drop$), pos2, feedthis
                If feedthis = 11 Or feedthis = 14 Or feedthis = 10 Then
                    feedthis = 0
                End If
newjump:
                feedback = feedthis
                If Col((UBound(Col()) \ 2) * 2) = 7 Then
                    a$ = vbNullString
                Else
                    a$ = Mid$(a$, pos1)
                End If
checkthistoo:
                If RTrim$(a$) = vbNullString Then
                    If pos2 > 0 Then
                        If feedthis <> 15 Then
                            If Col((UBound(Col()) \ 2) * 2) <> lastcol And pos2 < pos1 Then addone2000 pos2, lastcol
                        End If
                    Else
                        makebase 14
                    End If
                    mDoc.Paracolordata(thispara) = Col()
                    If feedthis = 15 Then GoTo use15
                    Exit Sub
                End If
                If kind = 8 And Myleft$(a$, 1) = "]" Then '****************************fix this
                    pos2 = K - Len(a$) + 1
                    addone2000 pos2, 13
                    a$ = Mid$(a$, 2)
                    pos2 = pos2 + 1 'k - Len(a$) + 1
                    If Myleft$(a$, 1) = "[" Then
                        a$ = Mid$(a$, 2)
                        pos2 = K - Len(a$) + 1
                        GoTo CheckThis
                    End If
                    onetime = False
                ElseIf kind > 4 And onetime Then
                    If FastSymbol(a$, ",") Then
  addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
  pos1 = 1
  GoTo CheckThis
  End If
End If

If FastSymbol(a$, "//", , 2) Then
addone pos2, 7
 pos2 = K - Len(a$) + 1
 mDoc.Paracolordata(thispara) = Col()
 Exit Sub
ElseIf FastSymbol(a$, "->", , 2) Then
addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
If FastSymbol(a$, "{") Then
    pos2 = K - Len(a$) + 1
    onetime = True
    GoTo again
End If
  ElseIf FastSymbol(a$, ",") Then
  addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
  pos1 = 1

  End If
ElseIf kind > 4 And onetime Then
If FastSymbol(a$, ",") Then
  addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
  pos1 = 1
  GoTo CheckThis
  End If

  End If

    If FastSymbol(a$, ")") Then
    pos2 = K - Len(a$)
    If lastcol = 0 Then
    addone2000 pos2, 14
    Else
  addone2000 pos2, lastcol
  End If
    pos2 = K - Len(a$) + 1
  End If
  End If


here3:
If kind > 0 Then
If MaybeIsSymbol(a$, ".") Then
If FastSymbol(a$, ".(", , 2) Then
addone2000 pos2, 14
a$ = vbNullString
    pos2 = K + 1
    
ElseIf FastSymbol(a$, "..", , 2) Then
    If Myleft$(a$, 1) = "." Then
        a$ = Mid$(a$, 2)
        lastcol = 11
        pos2 = K - Len(a$) + 1
        lastkind = 1
        addone2000 pos2 - 3, lastcol
    Else
        lastcol = 10
        lastkind = 1
        addone2000 pos2, lastcol
        pos2 = K - Len(a$) + 1
    End If
    GoTo here3
Else
    pos1 = pos1 + 1
    lastcol = 14
    GoTo CheckThis
  End If
End If
End If
If FastSymbol(a$, "<=", , 2) Then
  addone2000 pos2, 15
    pos2 = K - Len(a$) + 1
    Else
  
again4:
  If MaybeIsSymbolNoSpace(a$, "~", True) Then
    addone2000 pos2, 15
    pos2 = K - Len(a$) + 1
  ElseIf FastSymbol(a$, "=>", , 2) Then
    addone2000 pos2, 14
  
    pos2 = K - Len(a$) + 1
  ElseIf FastSymbol(a$, "=") Then
  If onetime And lastcol = 11 Then
    
  Else
    addone2000 pos2, 15
  End If
    pos2 = K - Len(a$) + 1
ElseIf FastSymbol(a$, "//", , 2) Then
addone pos2, 7
 pos2 = K - Len(a$) + 1
 mDoc.Paracolordata(thispara) = Col()
 Exit Sub
  ElseIf MaybeIsSymbol(a$, "+-*/", "=", True) Then
  addone2000 pos2, 15
    pos2 = K - Len(a$) + 1
  ElseIf FastSymbol(a$, "/*-+=~^|<>") Then
  addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
  If FastSymbol(a$, "/*-+=~^&|<>") Then
  addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
  End If
End If
  End If
again2:
Do
again3:
  pos1 = 1
  
  drop$ = aheadstatus2000(a$, False, pos1)
  
  If pos1 > 1 Or drop$ = "S" Then
  If pos1 = 1 Then
  drop$ = a$
  pos1 = Len(a$) + 1
  counter = 1
  Else
  drop$ = Mid$(a$, 1, pos1 - 1)
  If drop$ = vbNullString Then drop$ = " "
  End If
feedback = 0
 ColorizeExpressions2000 drop$, K - Len(a$) + Len(drop$), pos2, feedback
If feedback <> 15 Then
If feedback = 0 And lastcol = 0 Then feedback = 14
If drop$ = "(" Then
If feedback = 100 Then
    lastcol = 14
Else
lastcol = feedback
End If
Else 'If lastcol <> 11 And lastcol <> 10 Then
If feedback = 100 Then
    lastcol = 14
Else
If feedback > 0 And kind > 0 Then
lastcol = feedback
feedback = 0

a$ = Mid$(a$, pos1)

ColorizeInnerStatements a$, K, pos2, feedback, True

If MaybeIsSymbol(a$, ")") Then
addone2000 pos2, lastcol
FastSymbol a$, ")"
Exit Sub
End If
End If
End If
End If
feedback = 0
End If
 If feedback = 100 Then
  a$ = Mid$(a$, pos1)
 GoTo again112
 ElseIf feedback = 15 Then
use15:
 counter = 1
nextline:
 feedback = 0
 
 drop$ = mDoc.NextTextParagraph(thispara, feedback)
  If feedback = 0 Then
  a$ = Mid$(a$, pos1)
 mDoc.Paracolordata(thispara) = Col()
  Exit Sub
 End If
 pos1 = 1
 countbrakcet drop$, pos1, counter
 mDoc.Paracolordata(thispara) = Col()
  a$ = drop$

 K = Len(a$)

 If mDoc.startfrom(thispara) = 0 Then
 mDoc.startfrom(thispara) = thispara
  End If
  mDoc.startfrom(feedback) = mDoc.startfrom(thispara)
 thispara = feedback
 
    
    lastdata = a$
    pos2 = 1
    If Len(a$) > 0 Then
    ReDim Col(3)
             Col(0) = 15
            Col(1) = 1
             Col(2) = 15
             Col(3) = K + 1
             Else
    ReDim Col(1)
             Col(0) = 15
            Col(1) = 1
             End If
If counter = 0 Then
If UBound(Col()) > 1 Then
            
            Col(3) = pos1
            If Col(3) = Col(1) Then
            pos1 = pos1 + 1
            Col(3) = pos1
            Col(2) = 14
            End If
            Else
            Col(1) = pos1
            End If
If Myleft(a$, 1) = "}" Then
Else
a$ = Mid$(a$, pos1 + 1)
pos2 = K - Len(a$) + 1
End If
feedback = 0 'lastcol
feedthis = 0
If MaybeIsSymbol(a$, "<>=~") Then
    onetime = False
            pos2 = K - Len(a$) + 1
            addone2000 pos2, 14

    If FastSymbol(a$, "<") Then
        If FastSymbol(a$, ">") Then
        ElseIf FastSymbol(a$, "=") Then
        End If
    ElseIf FastSymbol(a$, ">") Then
        
        If FastSymbol(a$, "=") Then
        End If
        
        
    ElseIf FastSymbol(a$, "=") Then
    ElseIf FastSymbol(a$, "~") Then
    End If
    
End If
Else
    GoTo nextline

End If
    'kind = 0
   ' feedback = 0

            pos1 = 1

GoTo CheckThis
ElseIf feedback = 0 And pos1 > 1 And kind > 0 Then

GoTo newjump

 End If
 If pos1 > 0 Then
    If Mid$(a$, pos1 - 1, 1) = " " Then lastkind = 0
 
  a$ = Mid$(a$, pos1)
End If
If MaybeIsSymbol(a$, ":{") Then
   
    onetime = True
    GoTo again
End If
If InStr(drop$, " ") = 0 Or kind >= 0 Then
GoTo CheckThis

End If
  End If
again112:
    If MaybeIsSymbol(a$, ",;", , True) Then
     addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
    lastkind = 0
    onetime = False
    ElseIf FastSymbol(a$, "}") Then
    lastkind = 0
         addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
    onetime = True
   If MaybeIsSymbol(a$, ".;+(<>=~") Then
    onetime = False
    If FastSymbol(a$, "<") Then
        
        If FastSymbol(a$, ">") Then
        ElseIf FastSymbol(a$, "=") Then
        End If
    ElseIf FastSymbol(a$, ">") Then
        
        If FastSymbol(a$, "=") Then
        End If

    ElseIf FastSymbol(a$, "=") Then
    ElseIf FastSymbol(a$, "~") Then
    End If

GoTo CheckThis
End If
    GoTo again112
    ElseIf MaybeIsSymbol(a$, "'") Then
        lastkind = 0
         pos2 = K - Len(a$) + 1
        FastSymbol a$, "'"
        
        If pos2 > 0 Then
         addone2000 pos2, 7
         
         Else
        addone2000 1, 7
        End If
        a$ = vbNullString
        pos2 = K
        If pos2 > 0 Then
         addone2000 pos2 + 1, 14
        End If
   mDoc.Paracolordata(thispara) = Col()
   Exit Sub
    Else
   onetime = True
    mDoc.Paracolordata(thispara) = Col()
   Exit Do
   End If
   Loop
'If a$ = vbNullString Then
'
 'mDoc.Paracolordata(thispara) = col()
 'Exit Sub

'End If
GoTo again
ElseIf FastSymbol(a$, "//", , 2) Then
GoTo here1
ElseIf MaybeIsSymbol(a$, "\", , True) Then

If lastkind = 1 Then
 pos2 = K - Len(a$) + 1
 If AscW(a$ + ".") <> 32 Then
 If MaybeIsSymbol(a$, ".", ".", True) Then
 
 End If
 If MaybeIsSymbolNoSpace(a$, "\") Then GoTo again
    pos2 = K - Len(a$) + 1
    
    lastkind = 0
  If MaybeIsSymbol(a$, "\", , True) Then GoTo here1
  GoTo again
 
  End If
 Else
here1:
 
 If pos2 > 0 Then
     addone2000 pos2, 7
    
    Else
    addone2000 1, 7
    End If
 pos2 = K - Len(a$) + 1
 mDoc.Paracolordata(thispara) = Col()
 Exit Sub
 End If

ElseIf MaybeIsSymbol(a$, "'") Then

    If pos2 > 1 Then
         addone2000 pos2, 7
         Else
    
    addone2000 1, 7
    End If
 pos2 = K - Len(a$) + 1
 mDoc.Paracolordata(thispara) = Col()
 Exit Sub

ElseIf Not onetime Then
If FastSymbol(a$, ",") Then
mDoc.Paracolordata(thispara) = Col()
Exit Sub  ' stop
ElseIf FastSymbol(a$, ":") Then
onetime = True
ElseIf FastSymbol(a$, ")") Then
addone2000 pos2, lastcol
pos2 = K - Len(a$) + 1
    GoTo again4
ElseIf FastSymbol(a$, ";") Then
addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
    GoTo again2
ElseIf FastSymbol(a$, "}") Then
addone2000 pos2, 14
addone2000 K - Len(a$), 11
pos2 = K - Len(a$) + 1
    GoTo again3

ElseIf FastSymbol(a$, "{") Then
onetime = True
ElseIf FastSymbol(a$, "?") Then
addone2000 pos2, 14
addone2000 K - Len(a$), 11
pos2 = K - Len(a$) + 1
GoTo again2
Else
addone2000 pos2, lastcol
mDoc.Paracolordata(thispara) = Col()
Exit Sub
End If
addone2000 pos2, 14
    pos2 = K - Len(a$) + 1
    GoTo again
ElseIf IsSymbol(a$, "?") Then

addone2000 1, 11
        pos2 = K - Len(a$) + 1
        GoTo again2
ElseIf IsSymbol(a$, "=") Then

If pos2 > 0 Then
addone2000 pos2, 11
Else
addone2000 1, 11
End If
        pos2 = K - Len(a$) + 1
        onetime = False
        GoTo again2
ElseIf FastSymbol(a$, "{") Then
onetime = True
If pos2 > 0 Then
If Col((UBound(Col()) \ 2) * 2) <> 14 Then addone2000 pos2, 14
       pos2 = K - Len(a$) + 1
        If FastSymbol(a$, "&") Then
            addone2000 K - Len(a$), 13
            pos2 = K - Len(a$) + 1
        End If
Else
addone2000 1, 14
    pos2 = K - Len(a$) + 1
End If
    
        GoTo again

ElseIf FastSymbol(a$, "}") Then
If pos2 > 0 Then
addone2000 pos2, 14
Else
addone2000 1, 14
If MaybeIsSymbol(a$, "(") Then
GoTo CheckThis
ElseIf MaybeIsSymbol(a$, ",;") Then
a$ = Mid$(a$, 2)

 GoTo CheckThis
End If
End If
        pos2 = K - Len(a$) + 1
        GoTo again

Else
If IsNumberLabel(a$, drop$) Then
    If pos2 > 0 Then
        addone K - Len(a$) - Len(drop$) + 1, 14
    Else
        addone K - Len(a$) - Len(drop$) + 1, 15
        End If
        pos2 = K - Len(a$) + 1
        GoTo again
 End If
 End If
If pos2 > 0 Then addone2000 pos2, 14
If Col(1) = 0 Then
makebase 0
End If
End Sub



Private Sub mDoc_ColorDataEvent1(thispara As Long, data1 As String, feedback As Long)
Col() = mDoc.Paracolordata(thispara)
Dim drop$, Data As String
Data = Replace(data1, vbTab, " ")
Dim feedthis As Long, oldpara As Long
oldpara = thispara
If mDoc.startfrom(thispara) <> 0 Then
If mDoc.startfrom(thispara) <> thispara Then
drop$ = Data
GoTo entry100
End If
End If
If OtherSymbols = vbNullString Then OtherSymbols = ";"

If mNoColor Then
If UBound(Col()) > 1 Then
ReDim Col(1)
mDoc.Paracolordata(thispara) = Col()
End If
Exit Sub
End If
ReDim Col(1)
Dim a$, K As Long, pos1 As Long, pos2 As Long, onetime As Boolean, kind As Long
Dim lastcol As Long
a$ = Data
onetime = True
 K = Len(a$)
If K < 1 Then
makebase 14
mDoc.Paracolordata(thispara) = Col()
Exit Sub
End If
If feedback > 0 Then
lastcol = 15
pos2 = K - Len(a$) + 1

GoTo herecol
ElseIf MultiLineComment1 <> "{" Then
If FastSymbol(a$, mMultiLineComment2, , Len(mMultiLineComment2)) Then

GoTo entry3
End If
entry2:

If FastSymbol(a$, MultiLineComment1, , Len(MultiLineComment1)) Then

pos2 = K - Len(a$) + 1 - Len(MultiLineComment1)
feedback = 15
addone pos2, 15
GoTo entry2_6
entry2_5:
drop$ = ""

If pos2 > 0 Then
pos2 = K - Len(a$) - 1
  addone pos2, 15
  
  Else
  makebase 15
  'addone 1, 15
  End If
feedback = 15

entry2_6:
If InStr(a$, mMultiLineComment2) > 0 Then
a$ = Mid$(a$, InStr(a$, mMultiLineComment2) + Len(mMultiLineComment2))
entry3:
pos2 = K - Len(a$) + 1
If mDoc.startfrom(thispara) <> 0 Then
If mDoc.startfrom(thispara) = thispara Then mDoc.startfrom(thispara) = 0
End If
feedback = 0
pos1 = 1
GoTo marvel
Else

   If mDoc.startfrom(thispara) = 0 And feedback = 15 Then
   mDoc.startfrom(thispara) = thispara
   Else
   feedback = 15
   drop$ = a$
   End If
entry100:
    Do While thispara <> 0
  If oldpara <> thispara Then mDoc.startfrom(thispara) = mDoc.startfrom(oldpara)
  K = InStr(drop$, mMultiLineComment2)
  

    kind = 0
   oldpara = thispara
  If K > 0 Then
  a$ = Mid$(drop$, K + Len(mMultiLineComment2))
     K = Len(drop$)
    feedback = 0
    pos2 = K - Len(a$) + 1
    drop$ = vbNullString
    Exit Do
  

  End If
  mDoc.Paracolordata(oldpara) = Col()
  thispara = 0
  drop$ = mDoc.NextTextParagraph(oldpara, thispara)
  If thispara <> 0 Then
    pos2 = 1
    makebase 15
  End If
  Loop
  mDoc.Paracolordata(oldpara) = Col()
  If thispara = 0 Then thispara = oldpara
  If a$ = vbNullString Then Exit Sub
  pos1 = 1
  drop$ = ""
If K = 0 Then
GoTo thisexit
Else
GoTo marvel
End If
End If
End If
lastcol = 14
Else
lastcol = 14
End If
again:
If RTrim$(a$) = vbNullString Then
thisexit:
If pos2 > 0 Then
  If pos2 <= K Then addone pos2, lastcol
  Else
  makebase 14
  End If
mDoc.Paracolordata(thispara) = Col()
Exit Sub
' FOR TESTS
  pos1 = 0
  For K = 1 To UBound(Col()) Step 2
    If VarType(Col(K)) = 0 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
    If Col(K) > pos1 Then pos1 = Col(K)
    If Col(K) < pos1 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
  Next K
    For K = 0 To UBound(Col()) Step 2
   If Col(K) > 15 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
  Next K
  
End If
'pos2 = k - Len(a$) + 1
If Len(OpenTag) > 0 Then
If MaybeIsSymbol(a$, StartSymbols) Then
lastcol = 14
pos2 = pos2 + 1
GoTo herecol
End If
Else
While FastSymbol(a$, LineCommandSep)

If pos2 > 0 Then
addone pos2, 14
Else
addone 1, 14
End If
 pos2 = K - Len(a$) + 1
 
Wend
End If
If kind = 0 Then
'If IsSymbol(a$, "@") Then
      '  If pos2 > 0 Then
     '   addone pos2, 14
    '    Else
   '     addone 1, 14
  '      End If
 '    pos2 = k - Len(a$) + 1
'End If
End If

kind = IsLabelDot("", a$, drop$)

If kind = 0 Then
drop$ = vbNullString
End If
If kind > 0 Then
    If kind = 500 Then kind = 0: GoTo marvel
    
    If IsmineHere(drop$) Then
                If Len(OpenTag) > 0 Then
                 If MaybeIsSymbol2(a$, 1, "[=>/]") Then
        lastcol = 14
        Else
        lastcol = 10
        End If
                If pos2 > 0 Then
                   addone pos2, lastcol
                Else
                   addone K - Len(a$) - Len(drop$) + 1, lastcol
                   End If
                   pos2 = K - Len(a$) + 1

If MaybeIsSymbol2(a$, 1, "[=>/]") Then
    lastcol = 14

    If FastSymbol(a$, "=") Then
    pos2 = K - Len(a$)
        addone pos2, lastcol
        
        If FastSymbol(a$, Chr$(34)) Then
        
        pos2 = pos2 + 1
            drop$ = blockString2(a$)
            
            
            addone pos2, 15
        ElseIf FastSymbol(a$, stringsep2, , Len(stringsep2)) Then
        
        pos2 = pos2 + 1
            drop$ = blockString3(a$, stringsep2)
            
            
            addone pos2, 15
       
        Else
        
        pos2 = K - Len(a$) + 1
        GoTo again2
        End If
    End If
Else
       pos1 = InStr(a$, OpenTag)
If pos1 = 0 Then
mDoc.Paracolordata(thispara) = Col()
Exit Sub
End If
If pos1 > 1 Or pos2 = 0 Then
a$ = Mid$(a$, pos1)
If pos2 = 0 Then pos2 = 1
addone pos2, 10
pos2 = K - Len(a$) + 1
End If
   GoTo herecol

End If

               ' GoTo cont0020
                Else
                lastcol = 11
                If pos2 > 0 Then
                   addone pos2, lastcol
                Else
                   addone K - Len(a$) - Len(drop$) + 1, lastcol
                   End If
                End If

        '' for the end
                If IsmineHere1(drop$) Then
              
                kind = -2
                If Trim$(a$) <> "" Then
                                        pos2 = K - Len(a$) + 1
                                        If FastSymbol(a$, "{") Then
                                            addone pos2, 14
                                            pos2 = K - Len(a$) + 1
                                            onetime = True
                                            GoTo again
                                        ElseIf IsmineHere2(drop$) Then
                                        onetime = True
                                        kind = -2
                                        GoTo again2
                                        End If
                                End If
                ElseIf IsmineHere1(drop$) Then
                  
                kind = -1
                                If Trim$(a$) <> "" Then
                                                  pos2 = K - Len(a$) + 1
                                        If FastSymbol(a$, "{") Then
                                    
                                                  addone pos2, 14
                                              
                                                  pos2 = K - Len(a$) + 1
                                                  
                                                  onetime = True
                                                  GoTo again
                                          Else
                                           mDoc.Paracolordata(thispara) = Col()
                                                  Exit Sub
                                        End If
                                End If
                            
                End If
    
   Else
cont0020:
   If Len(OpenTag) > 0 Then
   If Len(drop$) > 0 Then
   
   If IsLabelOnly2000((drop$), drop$) > 1 Then
   If pos2 = 0 Then pos2 = 1
        addone pos2, 10
        pos2 = K - Len(a$) + 1
        GoTo herecol
   End If
   ElseIf Len(drop$) > 0 Then
   If pos2 = 0 Then pos2 = 1
        addone pos2, 10
        pos2 = K - Len(a$) + 1
        GoTo herecol

   End If
       pos1 = InStr(a$, OpenTag)
If pos1 = 0 Then
If FastSymbolNoRemove(a$, "=") Then
pos2 = K - Len(a$)
If pos2 = 0 Then pos2 = 1
addone pos2, 14
pos2 = K - Len(a$) + 1
GoTo herecol
Else
makebase 10
mDoc.Paracolordata(thispara) = Col()
Exit Sub
End If
ElseIf FastSymbolNoRemove(a$, "=") Then
pos2 = K - Len(a$)
If pos2 = 0 Then pos2 = 1
addone pos2, 14
pos2 = K - Len(a$) + 1
GoTo herecol

End If
a$ = Mid$(a$, pos1)
If pos2 = 0 Then pos2 = 1
addone pos2, 10
pos2 = K - Len(a$) + 1
   GoTo herecol
   
   Else
                lastcol = 10
                 If pos2 > 0 Then
                 addone pos2, lastcol
                Else
                 addone K - Len(a$) - Len(drop$) + 1, lastcol
                
                 End If
                End If
                

                 
  End If
  pos2 = K - Len(a$) + 1
  If kind > 3 Then
herecol:
If FastSymbolNoRemove(a$, MultiLineComment1, , Len(MultiLineComment1)) Then

'pos2 = k - Len(a$) + 1 - Len(MultiLineComment1)
   onetime = False
   lastcol = 15
   feedthis = 15
   drop$ = ""
   drop$ = ""

If pos2 > 0 Then
  addone pos2, 15
  
  Else
  makebase 15
  'addone 1, 15
  End If
feedback = 0
   
GoTo entry2_6
End If
  pos1 = 1
 drop$ = aheadstatus(a$, False, pos1)
''  pos1 = Len(A$) + 1

marvel:

  If pos1 > 1 Then
marvel2:

 ColorizeExpressions Mid$(a$, 1, pos1 - 1), K - Len(a$) + pos1 - 1, pos2, feedthis
'ColorizeExpressions Mid$(a$, 1, pos1 - 1), k, pos2, feedthis

 If feedthis = 1 Then
 mDoc.Paracolordata(thispara) = Col()
 Exit Sub
 
 End If
 feedback = feedthis
  
If feedthis = -15 Then
   a$ = Right$(a$, K - pos2)
   onetime = False
   lastcol = 15
   feedthis = 15
   feedback = 15
   drop$ = ""
   
   GoTo entry2
   Else
   a$ = Mid$(a$, pos1)
 End If
  If SeekNextLine And a$ = vbNullString Then
  thispara = 0
  drop$ = mDoc.NextTextParagraph(oldpara, thispara)
  If thispara <> 0 Then
  ' for java
  If MaybeIsSymbol(drop$, ";,+-*/&|!=)]" + Chr(34)) Then
  mDoc.Paracolordata(oldpara) = Col()
    
    a$ = drop$
    K = Len(a$)
    pos2 = 1
    ReDim Col(1)
    kind = 0
   oldpara = thispara
  End If
  Else
  thispara = oldpara
  End If
  End If

  If FastSymbol(a$, mSplitExpr) Then
  If Len(OpenTag) > 0 Then
  addone pos2, 10
  Else
       addone pos2, 14
       End If
    pos2 = K - Len(a$) + 1
  pos1 = 1
  drop$ = aheadstatus(a$, False, pos1)
GoTo marvel2
ElseIf MaybeIsSymbol(a$, mSplitExpr, , True) Then
addone pos2, 14
    pos2 = K - Len(a$) + 1
  pos1 = 1
  drop$ = aheadstatus(a$, False, pos1)
GoTo marvel2
ElseIf MaybeIsSymbol(a$, OtherSymbols, , True) Then

addone pos2, 14
    pos2 = K - Len(a$) + 1
  pos1 = 1
  drop$ = aheadstatus(a$, False, pos1)
GoTo marvel2
' elseif  ****
End If
  
  End If
  If SeekNextLine And a$ = vbNullString Then
  thispara = 0
  drop$ = mDoc.NextTextParagraph(oldpara, thispara)
  If thispara <> 0 Then
  ' for java
  If MaybeIsSymbol(drop$, ";,+-*/&|!=)" + Chr(34)) Or kind > 4 Then
  mDoc.Paracolordata(oldpara) = Col()
  If mDoc.startfrom(oldpara) = 0 Then
  mDoc.startfrom(oldpara) = oldpara
  End If
  mDoc.startfrom(thispara) = mDoc.startfrom(oldpara)
  kind = 0
    a$ = drop$
    K = Len(a$)
    pos2 = 1
    ReDim Col(1)
    oldpara = thispara
   GoTo herecol
   Else
   thispara = oldpara
  End If
  Else
  thispara = oldpara
  End If
  End If
    If FastSymbol(a$, ")") Then
    If kind <> 5 Then
   addone pos2, 14
   Else
   addone pos2, lastcol
   End If
   'lastcol ' '
    pos2 = K - Len(a$) + 1
  End If
  If Brackets Then
    If FastSymbol(a$, "]") Then
  If kind <> 5 Then
   addone pos2, 14
   Else
   addone pos2, lastcol
   End If
    pos2 = K - Len(a$) + 1
  End If
  End If
  
  
  
  End If
If FastSymbol(a$, mSplitExpr) Then
  addone pos2, 14 ' WAS LASTCOL
    pos2 = K - Len(a$) + 1
End If

  
again4:
'
  If FastSymbol(a$, AssignSym, , Len(AssignSym)) Then
  addone pos2, 15
    pos2 = K - Len(a$) + 1
    
    onetime = False
    
  ElseIf MaybeIsSymbol(a$, "+-*/", "=", True) Then
  addone pos2, 14
    pos2 = K - Len(a$) + 1
  ElseIf MaybeIsSymbolNoSpace(a$, "~", True) Then
  addone pos2, 14
    pos2 = K - Len(a$) + 1
End If
again2:
Do
again3:
  pos1 = 1
 '' kind = -2 maybe
a$ = NLtrim$(a$)
pos2 = K - Len(a$) + 1
  drop$ = aheadstatus(a$, False, pos1)
  
  If pos1 > 1 Or drop$ = "S" Then
  If pos1 = 1 Then
  drop$ = a$
  pos1 = Len(a$) + 1
  Else
  drop$ = Mid$(a$, 1, pos1 - 1)
  If Brackets Then
  If onetime And kind = 1 Then
  If Myleft$(a$, 1) = "(" Or Myleft$(a$, 1) = "[" Then
  pos1 = 1
  drop$ = Mid$(drop$, 2)
  a$ = Mid$(a$, 2)
  addone pos2, 14
  pos2 = pos2 + 1
  aheadstatus drop$, False, pos1
  drop$ = Mid$(a$, 1, pos1 - 1)
 
  
  
  'pos2 = k - Len(a$)
  kind = 40
  End If
  End If
  End If
  If drop$ = vbNullString Then drop$ = " "
  End If
feedback = 0
 ColorizeExpressions drop$, K - Len(a$) + pos1 - 1, pos2, feedback
 If kind = 40 Then
    kind = 1
    a$ = NLtrim$(Mid$(a$, pos1))
    
    pos2 = K - Len(a$) + 1
    
    If a$ = vbNullString Then
    ElseIf feedback = -15 Then
            a$ = Mid$(a$, pos1)
        pos2 = K - Len(a$) ' + 1 - Len(MultiLineComment1)
           onetime = False
   lastcol = 15
   feedthis = 15
   drop$ = ""
        GoTo entry2_5

    ElseIf Myleft(a$, 1) = LineCommandSep Then
    a$ = Mid$(a$, 2)
    pos2 = K - Len(a$) + 1
    If Not a$ = vbNullString Then
    
    GoTo herecol
    End If
    ElseIf Myleft(a$, 1) = mSplitExpr Then
        kind = 40
        If Len(OpenTag) > 0 Then
        addone pos2, 10
        Else
        addone pos2, 14
        End If
        a$ = Mid$(a$, 2)
        pos2 = pos2 + 1
         GoTo again3
    ElseIf MaybeIsSymbol(a$, ")]", , True) Then
    kind = 0
        addone pos2, 14
        pos2 = K - Len(a$) + 1
        GoTo again4
    End If
 ElseIf feedback = -15 Then

            a$ = Mid$(a$, pos1 - Len(drop$) + Len(MultiLineComment1))
        pos2 = K - Len(a$)
        
          ' onetime = False
   lastcol = 15
   feedthis = 15
   drop$ = ""
        GoTo entry2_5

Else
a$ = Mid$(a$, pos1)
pos2 = K - Len(a$) + 1
If drop$ = "(" Then
    If feedback = 100 Then
        lastcol = 14
    End If
ElseIf drop$ = "[" Then
     
    If feedback = 100 Then
        lastcol = 14
    End If

End If

End If
 If feedback = 1 Then
 feedthis = 0
 mDoc.Paracolordata(thispara) = Col()
 Exit Sub
 
  End If
   '''was pos1
    If SeekNextLine And a$ = vbNullString Then
  thispara = 0
  drop$ = mDoc.NextTextParagraph(oldpara, thispara)
  If thispara <> 0 Then
  

  If MaybeIsSymbol(drop$, ";,+-*/&|!=)" + Chr(34)) Then
  mDoc.Paracolordata(oldpara) = Col()
    kind = 0
    a$ = drop$
    K = Len(a$)
    pos2 = 1
    ReDim Col(1)
    oldpara = thispara
   
  End If
  Else
  thispara = oldpara
  End If
  End If
''  If InStr(DROP$, " ") = 0 Then
If InStr(drop$, " ") = 0 Or kind >= 0 Then
If MaybeIsSymbol(a$, mSplitExpr, , True) Then
     addone pos2, 14
    pos2 = K - Len(a$) + 1
    onetime = False
    End If
    While FastSymbol(a$, LineCommandSep)

If pos2 > 0 Then
addone pos2, 14
Else
addone 1, 14
End If
 pos2 = K - Len(a$) + 1
 
Wend
    
GoTo herecol
End If
  End If
    If MaybeIsSymbol(a$, OtherSymbols, , True) Then
     addone pos2, 14
    pos2 = K - Len(a$) + 1
    onetime = False
    ElseIf FastSymbol(a$, "}") Then
         addone pos2, 14
    pos2 = K - Len(a$) + 1
    onetime = True
    GoTo again
   Else
   onetime = True
   Exit Do
   End If
   Loop

GoTo again
ElseIf MaybeIsSymbol(a$, LineComment2) Then
If pos2 > 0 Then
 addone pos2, 7

Else
addone 1, 7
End If
    If FastSymbol(a$, MultiLineComment1, , Len(MultiLineComment1)) Then

        pos2 = K - Len(a$) + 1 - Len(MultiLineComment1)
           onetime = False
   lastcol = 15
   feedthis = 15
   drop$ = ""
        GoTo entry2_5
        End If
 pos2 = K - Len(a$) + 1
 mDoc.Paracolordata(thispara) = Col()
 Exit Sub
ElseIf MaybeIsSymbol2(a$, ComSymbolsWidth, CommentSymbols) Then
If pos2 > 0 Then
 addone pos2, CLng(IIf(CommentLineLight, 12, 7))

Else
addone 1, CLng(IIf(CommentLineLight, 12, 7))
End If
 pos2 = K - Len(a$) + 1
 mDoc.Paracolordata(thispara) = Col()
 Exit Sub
ElseIf Not onetime Then
If FastSymbol(a$, mSplitExpr) Then
If Len(OpenTag) > 0 Then
addone pos2, 14
    pos2 = K - Len(a$) + 1
    End If
ElseIf FastSymbol(a$, LineCommandSep) Then
onetime = True
ElseIf FastSymbol(a$, ")") Then
addone pos2, lastcol
pos2 = K - Len(a$) + 1
    GoTo again4
ElseIf MaybeIsSymbol(a$, OtherSymbols, , True) Then
addone pos2, 14
    pos2 = K - Len(a$) + 1
    GoTo again2
ElseIf FastSymbol(a$, "}") Then
addone pos2, 14
addone K - Len(a$), 11
pos2 = K - Len(a$) + 1
    GoTo again3

ElseIf FastSymbol(a$, "{") Then
onetime = True
ElseIf FastSymbol(a$, "?") Then
addone pos2, 14
addone K - Len(a$), 11
pos2 = K - Len(a$) + 1
GoTo again2
Else
If Brackets Then
    If FastSymbol(a$, "]") Then
  addone pos2, lastcol
    pos2 = K - Len(a$) + 1
  End If
  End If
addone pos2, lastcol
mDoc.Paracolordata(thispara) = Col()
Exit Sub
End If
addone pos2, 14
    pos2 = K - Len(a$) + 1
    GoTo again
ElseIf IsSymbol(a$, "?") Then

If pos2 > 0 Then
addone pos2, 14
Else
addone 1, 14
End If
        pos2 = K - Len(a$) + 1
        GoTo again2
ElseIf IsSymbol(a$, "=") Then
If Len(OpenTag) > 0 Then
If pos2 > 0 Then
addone pos2, 14
Else
addone 1, 14
End If
Else
If pos2 > 0 Then
addone pos2, 11
Else
addone 1, 11
End If
End If
        pos2 = K - Len(a$) + 1
        GoTo again2
ElseIf FastSymbol(a$, "{") Then
onetime = True
If pos2 > 0 Then
addone pos2, 14
 pos2 = K - Len(a$) + 1
    If FastSymbol(a$, "&") Then
          addone pos2, 10
      pos2 = K - Len(a$) + 1
End If
Else
addone 1, 14
    pos2 = K - Len(a$) + 1
End If
    
        GoTo again

ElseIf FastSymbol(a$, "}") Then
If pos2 > 0 Then
addone pos2, 14
Else
addone 1, 14
End If
        pos2 = K - Len(a$) + 1
        GoTo again

ElseIf pos2 = 0 Then
If Myleft$(a$, 1) <> "" Then
If InStr(StartSymbols + Chr$(0), Myleft$(a$, 1)) > 0 Then
pos1 = 1
 drop$ = aheadstatus(a$, False, pos1)
 pos2 = K - Len(a$) + 1
If drop$ <> "" Then GoTo marvel

End If
End If
If AutoNumber Then
If IsNumberLabel(a$, drop$) Then
        addone K - Len(a$) - Len(drop$) + 1, 15
        pos2 = K - Len(a$) + 1
        GoTo again
 End If
 
 End If
          
       ' pos2 = k - Len(a$) + 1
 End If
If pos2 > 0 Then
If Len(OpenTag) > 0 Then
addone pos2, 10
Else
addone pos2, 14
End If
End If
If Col(1) = 0 Then

makebase 0
End If
mDoc.Paracolordata(thispara) = Col()
End Sub


Private Sub mDoc_Lastline()
If showparagraph Then showparagraphonce = True
End Sub


Private Sub mDoc1_BreakLine(Data As String, datanext As String)
If NoHeadwrap Or (NoTitle) Or UsedAsTextBox Then
    datanext = Data
    Data = vbNullString
Else
 glistN.BreakLine Data, datanext
End If
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
 Private Sub PrintItem1(mHdc As Long, c As String, r As RECT, Optional way As Long = DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP)
 Dim cc As String
r.top = r.top + TextPixelOffset
 If showparagraphonce Then
         cc = c & "Ά"
  DrawTextEx mHdc, StrPtr(cc), -1, r, way, VarPtr(tParam)
 Else
  DrawTextEx mHdc, StrPtr(c), -1, r, way, VarPtr(tParam)
   
    End If
    r.top = r.top - TextPixelOffset
    End Sub
 Private Sub PrintItem(mHdc As Long, c As String, r As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP)
  r.top = r.top + TextPixelOffset
  DrawTextEx mHdc, StrPtr(c), -1, r, way, VarPtr(tParam)
r.top = r.top - TextPixelOffset
    End Sub
     Private Sub PrintItem2(mHdc As Long, c As String, r As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP)
 Dim that As Long
 that = way
 r.Left = r.Left + 4
If glistN.VerticalCenterText Then that = that Or DT_VCENTER Or DT_EXPANDTABS Or DT_TABSTOP

    DrawTextEx mHdc, StrPtr(c), -1, r, that, VarPtr(tParam)

    End Sub
Public Property Get EditDoc() As Boolean
EditDoc = mEditDoc
End Property

Public Property Let EditDoc(ByVal RHS As Boolean)
mEditDoc = RHS
If Not RHS Then glistN.MarkNext = 0
glistN.EditFlag = RHS
If glistN.enabled Then
glistN.ShowMe2
End If
End Property

Public Property Get ParaSelStart() As Long
If glistN.ListIndex < 0 Then ParaSelStart = 1: Exit Property
Dim what$, ThatPara As Long, pastchars
    pastchars = mDoc.TextParagraphLine(glistN.ListIndex + 1, what$, ThatPara)
Charpos = pastchars + glistN.SelStart
If Charpos = 0 Then
ParaSelStart = 1
Else
ParaSelStart = Charpos
End If
End Property
Public Sub WrapAll()
If mDoc Is Nothing Then Exit Sub
mDoc.WrapAgain
End Sub
Public Property Let ParaSelStart(ByVal RHS As Long)
' we have a new ParaSelStart
' we want a new listindex...and a new selstart

Charpos = RHS
Render
End Property

Public Property Get nowrap() As Boolean
nowrap = mNoWrap
End Property

Public Property Let nowrap(ByVal RHS As Boolean)
If UsedAsTextBox Then Exit Property
If Not (mNoWrap = RHS) Then
mNoWrap = RHS
glistN.NoScroll = Not mNoWrap
    If Not mDoc.IsEmpty Then
    WrapAll
   If HaveMarkedText Then RelocateMarkedText
    
    HorSrcoll 0
    Render
    End If
End If
End Property

Public Property Get enabled() As Boolean
enabled = glistN.enabled
End Property

Public Property Let enabled(ByVal RHS As Boolean)
'glistN.HideCaretOnexit = Not rhs
If Not RHS Then
glistN.Noflashingcaret = False
Show
End If
glistN.enabled = RHS
End Property
Sub SetFocus()
If glistN.Visible Then glistN.SetFocus
End Sub

Public Property Get ScrollKeyUpDown() As Boolean
ScrollKeyUpDown = glistN.NoFreeMoveUpDown
End Property

Public Property Let ScrollKeyUpDown(ByVal RHS As Boolean)
glistN.NoFreeMoveUpDown = RHS
End Property
Public Sub SetRowColumn(Row As Long, ByVal Column As Long, Optional noShow As Boolean = False)
' base 0 for column as input
Column = Column + 1
' base 1
If Column < 0 Then Column = 0
If Row = -1 Then
glistN.ListindexPrivateUse = glistN.listcount - 1
Else
If Row > glistN.listcount + 1 Then
Row = glistN.listcount + 1
End If
glistN.ListindexPrivateUse = Row - 1  ' no events
End If
Dim a As String, b As Long
mDoc.TextParagraphLine glistN.ListIndex + 1, a, b
mDoc.MarkParagraphID = b
If Len(glistN.list(glistN.ListIndex)) < Column Then
If Row = glistN.listcount + 1 Then
glistN.SelStartEventAlways = Len(glistN.list(glistN.ListIndex))
Else
glistN.SelStartEventAlways = Column  'Len(glistN.List(glistN.listindex)) + 1
End If
Else
glistN.SelStartEventAlways = Column
End If
If glistN.enabled Then
If Not noShow Then glistN.ShowMe2
End If
End Sub
Public Function TextFormatBreak(lineBreaker As String) As String
Text = mDoc.textFormat(lineBreaker)
End Function

Public Property Get Text() As String
If mDoc Is Nothing Then Exit Property
Text = mDoc.textDoc
End Property
Public Property Let Text(RHS As String)
On Error Resume Next
If glistN Is Nothing Then Exit Property
LastListIndex = 0
If UsedAsTextBox Or NumberOnly Or NumberIntOnly Then
If Not mNoWrap = True Then
    mNoWrap = True
    glistN.NoScroll = False
End If
End If
Dim oldenabled As Boolean
If Me.HaveMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End If
oldenabled = glistN.enabled
glistN.enabled = False
glistN.PanPos = 0
glistN.ScrollToSilent 0
Set mDoc = New Document
mDoc.ColorEvent = True
mDoc.Final = mMultiLineComment2
mDoc.textDoc = RHS
mDoc.NoIgnoreCase = Not mUseCase
glistN.ListindexPrivateUse = 0
para1 = 0
PosPara1 = 0
Para2 = 0
PosPara2 = 0
Para3 = 0
PosPara3 = 0
expandCharPos = 0
Charpos = 0
BlockStartPara = 0
BlockStartPos = 1
BlockStartLine = 0
BlockTopline = 0
BlockStartCursor = 0
BlockEndCursor = 0
BlockEndPara = 0
BlockEndPos = 1
BlockEndline = 0
BlockBottomline = 0
Charpos = 0
all = mDoc.DocLines
 glistN.SelStart = 0
ClearRedo
ClearUndo
glistN.enabled = oldenabled
Dim what$, ThatPara As Long, pastchars
   pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
End Property
Public Property Get TextParagraphs() As Long
TextParagraphs = mDoc.DocParagraphs
End Property
Public Property Get TextLines() As Long
TextLines = mDoc.DocLines
End Property
Public Property Let LineSpacingSilent(ByVal RHS As Long)
RHS = Abs(RHS) + 2
glistN.addpixels = RHS
''Render
End Property
Public Property Let LineSpacing(ByVal RHS As Long)
RHS = Abs(RHS) + 2
glistN.addpixels = RHS
Render
End Property

Public Property Get SelLength() As Long
If HaveMarkedText Then
If BlockStartLine = BlockEndline Then
SelLength = Abs(BlockStartCursor - BlockEndCursor)
Else
Dim a$(), oldmydoc As Document, i As Long, sum As Long
' maybe we didn't need to preserve old UndoDoc...I need to clarify this..
Set oldmydoc = UndoDoc
Set UndoDoc = New Document
FillBlock
sum = UndoDoc.SizeCRLF + 2
If UndoDoc.LfLead Then sum = sum - 1
If UndoDoc.CrTail Then sum = sum + 1
SelLength = sum - 2 ' no the last vbcrlf
End If
Set UndoDoc = oldmydoc
End If
End Property
Public Property Let SelLengthSilent(ByVal RHS As Long)
If RHS <= 0 Then
    ' a copy of glistN_MarkDestroyAny
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
'If glistN.Enabled Then glistN.ShowPan
Else
glistN.MarkNext = 1
   BlockStartCursor = glistN.SelStart
    BlockStartPara = mDoc.MarkParagraphID
    BlockStartPos = ParaSelStart  ' selstart isn't from start of the doc...
 
    BlockStartLine = glistN.ListIndex
        BlockEndPara = BlockStartPara
        BlockEndPos = BlockStartPos
        BlockEndline = BlockStartLine
    BlockEndCursor = BlockStartCursor + 1
    
mDoc.FindPos BlockStartLine + 1, BlockStartCursor, RHS, BlockEndline, BlockEndCursor, BlockEndPara, BlockEndPos

BlockEndline = BlockEndline - 1 ' base 0


    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
If RHS <> 0 Then
glistN.MarkNext = 2
'If glistN.Enabled Then glistN.ShowPan
End If
End If
End Property
Public Property Let SelLength(ByVal RHS As Long)
If RHS <= 0 Then
    ' a copy of glistN_MarkDestroyAny
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
If glistN.enabled Then glistN.ShowPan
Else
glistN.MarkNext = 1
   BlockStartCursor = glistN.SelStart
    BlockStartPara = mDoc.MarkParagraphID
    BlockStartPos = ParaSelStart  ' selstart isn't from start of the doc...
 
    BlockStartLine = glistN.ListIndex
        BlockEndPara = BlockStartPara
        BlockEndPos = BlockStartPos
        BlockEndline = BlockStartLine
    BlockEndCursor = BlockStartCursor + 1
    
mDoc.FindPos BlockStartLine + 1, BlockStartCursor, RHS, BlockEndline, BlockEndCursor, BlockEndPara, BlockEndPos

BlockEndline = BlockEndline - 1 ' base 0

    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
If RHS <> 0 Then
glistN.MarkNext = 2
If glistN.enabled Then glistN.ShowPan
End If
End If
End Property

Property Let Visible(ByVal RHS As Boolean)
    If Not glistN Is Nothing Then
        glistN.Visible = RHS
    End If
End Property
Property Get Visible() As Boolean
    If Not glistN Is Nothing Then
        Visible = glistN.Visible
    End If
End Property
Public Property Get FontSize() As Single

  FontSize = glistN.FontSize
 
End Property

Public Property Let FontSize(New_FontSize As Single)
glistN.FontSize = New_FontSize
End Property
Public Property Get Font() As Font
 Set Font = glistN.Font
End Property

Public Property Set Font(New_Font As Font)
    Set glistN.Font = New_Font
End Property
Public Property Let BackColor(ByVal RHS As OLE_COLOR)

    
glistN.BackColor = RHS
  
    
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = glistN.BackColor
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = glistN.ForeColor
End Property

Public Property Let ForeColor(ByVal RHS As OLE_COLOR)

    glistN.ForeColor = RHS

End Property

Public Property Get top() As Long
top = Callback.Controls(ctrlName).top
End Property

Public Property Let top(ByVal RHS As Long)
Callback.Controls(ctrlName).top = RHS
End Property
Public Property Get Left() As Variant
Left = Callback.Controls(ctrlName).Left
End Property

Public Property Let Left(ByVal RHS As Variant)
Callback.Controls(ctrlName).Left = RHS
End Property
Public Property Get Height() As Long
Height = glistN.HeightTwips
End Property
Public Property Let Height(ByVal RHS As Long)
glistN.HeightTwips = RHS
End Property
Public Sub Dereference()
Set glistN = Nothing
Set mDoc = Nothing
Set mDoc1 = Nothing

End Sub

Public Property Get SelText() As String
If HaveMarkedText Then
FillBlock

If UndoDoc.LfLead Then
SelText = Mid$(UndoDoc.textDoc, 2)
ElseIf UndoDoc.CrTail Then
SelText = UndoDoc.textDoc + vbCr
Else  ''If UndoDoc.LfLead = UndoDoc.CrTail Then
SelText = UndoDoc.textDoc
End If
End If
End Property
Public Property Get SelHtmlText() As String
If HaveMarkedText Then
''FillBlock
Dim curcolors(0 To 15) As String, i As Long
For i = 0 To 15
curcolors(i) = Right$("000000" + Hex$(htmlM2000(i)), 6)

Next i

If UndoDoc.LfLead Then
SelHtmlText = Replace(Mid$(UndoDoc.textHtmlDoc(curcolors()), 2), vbCrLf, "<br>")
Else
SelHtmlText = Replace(UndoDoc.textHtmlDoc(curcolors()), vbCrLf, "<br>")
End If


End If
End Property
Public Property Let SelText(ByVal RHS As String)
Dim kk As Long, Pad$, prev As Boolean
kk = SelStart
SelStart = kk
prev = HaveMarkedText
glistN_MarkDelete False
glistN.MarkNext = 1
BlockStartPara = mDoc.MarkParagraphID
BlockStartPos = ParaSelStart
BlockStartCursor = glistN.SelStart '- 1
BlockEndCursor = BlockStartCursor + 1
BlockStartLine = glistN.ListIndex
BlockEndline = BlockStartLine
BlockEndPara = BlockStartPara
BlockEndPos = BlockStartPos
Dim K As Long, l As Long
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
If RHS <> "" Then
If prev Then undoprev = prev
If NumberOnly Then
glistN.list(0) = RHS
Else
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS, , , Not NoColor
End If
Render
End If
BlockEndPara = mDoc.MarkParagraphID
BlockEndPos = ParaSelStart
BlockEndline = glistN.ListIndex
BlockEndCursor = glistN.SelStart
If BlockEndline < BlockStartLine Then
    BlockTopline = BlockEndline
    BlockBottomline = BlockStartLine
Else
    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
End If


K = kk
glistN.MarkNext = 2
SelStart = K

End Property
Public Property Let InsertTextNoRender(ByVal RHS As String)
Dim b As Boolean
ClearRedo
PushUndo
glistN_MarkDelete False
glistN.MarkNext = 0
undomode = 5 ' like 1
RemoveUndo (RHS)
undomode = 4 ' like 0
b = enabled
If b Then enabled = False
Dim K As Long, l As Long
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
If RHS <> "" Then


mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
all = mDoc.DocLines

End If
If b <> enabled Then enabled = b
End Property
Private Property Let InsertText2(ByVal RHS As String)
glistN_MarkDelete True
glistN.MarkNext = 0
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
mDoc.ResetColor mDoc.MarkParagraphID
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
End Property
Public Property Let InsertText(ByVal RHS As String)
glistN_MarkDelete False
glistN.MarkNext = 0
Dim K As Long, l As Long
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
If RHS <> "" Then
mDoc.ResetColor mDoc.MarkParagraphID
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
Render

End If
End Property
Public Property Get CharPosStart() As Long
Dim findit As Long
Dim a As Long, b As Long, c As Long, D As Long
findit = -1
mDoc.FindPos glistN.ListIndex + 1, glistN.SelStart, findit, a, b, c, D
CharPosStart = findit
End Property
Public Property Get SelStart() As Long
'
Dim findit As Long
Dim a As Long, b As Long, c As Long, D As Long
findit = -1
If glistN.MarkNext > 0 Then
If BlockTopline = BlockBottomline Then
If BlockStartCursor < BlockEndCursor Then
mDoc.FindPos BlockTopline + 1, BlockStartCursor, findit, a, b, c, D
Else
mDoc.FindPos BlockTopline + 1, BlockEndCursor, findit, a, b, c, D
End If
ElseIf BlockTopline = BlockStartLine Then
mDoc.FindPos BlockTopline + 1, BlockStartCursor, findit, a, b, c, D
Else
mDoc.FindPos BlockTopline + 1, BlockEndCursor, findit, a, b, c, D
End If
Else
mDoc.FindPos glistN.ListIndex + 1, glistN.SelStart, findit, a, b, c, D
End If
If findit <> -1 Then LastSelStart = findit
SelStart = findit
End Property

Public Property Let SelStart(ByVal RHS As Long)
Dim findit As Long, sel1 As Long
sel1 = SelLength
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
Dim a As Long, b As Long, c As Long, D As Long
findit = RHS
mDoc.FindPos 1, 0, findit, a, b, c, D

SetRowColumn a, b, sel1 <> 0
If sel1 <> 0 Then SelLength = sel1: LastSelStart = sel1
End Property
Public Property Let SelStartSilent(ByVal RHS As Long)
Dim findit As Long
Dim a As Long, b As Long, c As Long, D As Long
LastSelStart = RHS
findit = RHS
mDoc.FindPos 1, 0, findit, a, b, c, D
SetRowColumn a, b, True
End Property

Sub ClearUndo()
undopara = -1
NeoUndo = 0
UndoMax = 20
ReDim MultiUndo(1 To UndoMax)
End Sub
Function FindNewUndo() As Long
If NeoUndo = 0 Then
UndoMax = 20
ReDim MultiUndo(1 To UndoMax)
End If
NeoUndo = NeoUndo + 1
If NeoUndo > UndoMax Then
UndoMax = UndoMax * 2
ReDim Preserve MultiUndo(1 To UndoMax)
ElseIf UndoMax > 20 And NeoUndo < UndoMax / 2 Then
UndoMax = UndoMax / 2
ReDim Preserve MultiUndo(1 To UndoMax)
End If
FindNewUndo = NeoUndo
End Function
Private Sub PushUndo()
If undopara = -1 Then Exit Sub
Dim b As Long
b = FindNewUndo
With MultiUndo(b)
.undotext = undotext  ' new  seltext
.undopara = undopara  ' not need for this 2nd version
.undomode = undomode  ' needed
.undopos = undopos  ' new not charpos but selstart
.undoprev = undoprev
undopos = ParaSelStart
undotext = vbNullString
undopara = mDoc.MarkParagraphID  ' nod needed maybe for trace errors later
undomode = -1
undopara = -1
undoprev = False
End With
End Sub
Private Function PopUndo() As Boolean
' false ...no undo saved
If NeoUndo < 1 Then Exit Function

With MultiUndo(NeoUndo)
undotext = .undotext
.undotext = vbNullString  ' flush
undopara = .undopara
undomode = .undomode
undopos = .undopos
undoprev = .undoprev
End With
NeoUndo = NeoUndo - 1
PopUndo = True
If undopara < 0 Then Exit Function

End Function
Sub ClearRedo()
Redopara = -1
NeoRedo = 0
RedoMax = 20
ReDim MultiRedo(1 To RedoMax)
End Sub
Function FindNewRedo() As Long
If NeoRedo = 0 Then
RedoMax = 20
ReDim MultiRedo(1 To RedoMax)
End If
NeoRedo = NeoRedo + 1
If NeoRedo > RedoMax Then
RedoMax = RedoMax * 2
ReDim Preserve MultiRedo(1 To RedoMax)
ElseIf RedoMax > 20 And NeoRedo < RedoMax / 2 Then
RedoMax = RedoMax / 2
ReDim Preserve MultiRedo(1 To RedoMax)
End If
FindNewRedo = NeoRedo
End Function
Private Sub PushRedo()

If Redopara = -1 Then Exit Sub
Dim b As Long
b = FindNewRedo
With MultiRedo(b)
.undotext = Redotext
.undopara = Redopara
.undomode = Redomode
.undopos = Redopos
.undoprev = Redoprev
Redopos = ParaSelStart
Redotext = vbNullString
Redopara = mDoc.MarkParagraphID
Redomode = -1
'Redoprev = False
End With
End Sub
Private Function PopRedo() As Boolean
' false ...no Redo saved
If NeoRedo < 1 Then Exit Function
With MultiRedo(NeoRedo)
Redotext = .undotext
.undotext = vbNullString  ' flush
Redopara = .undopara
Redomode = .undomode
Redopos = .undopos
Redoprev = .undoprev
End With
NeoRedo = NeoRedo - 1
PopRedo = True
End Function

Public Sub ResetUndoRedo()
ClearUndo
ClearRedo
Redopara = -1
undopara = -1
End Sub
Public Function CurrentParagraph() As String
CurrentParagraph = mDoc.TextParagraph(mDoc.MarkParagraphID)
End Function
Public Function NextParagraph() As String
Dim dummy As Long
If showparagraphWork And showparagraph Then
NextParagraph = Replace(Replace(mDoc.NextTextParagraph(mDoc.MarkParagraphID, dummy), " ", ChrW(&HB7)), ChrW(&H2007), ChrW(&H2423))
Else
NextParagraph = mDoc.NextTextParagraph(mDoc.MarkParagraphID, dummy)
End If
End Function
Public Function AutoIntNewLine() As Boolean
' if autonumber
Dim a$, newid As Long 'mDoc.MarkParagraphID
Dim b$, c$, D$, e$, skip As Boolean, tr$, maxw As Long, Ns As String, och As Long
och = -1
showparagraphWork = False
b$ = CurrentParagraph
If Charpos > Len(b$) + 1 Then SelStartSilent = SelStart - (Charpos - Len(b$) - 1)
FixCharPos
If AutoNumber Then
    If IsNumberLabel(b$, c$) Then
        D$ = NextParagraph: If D$ <> "" Then If IsNumberLabel(D$, e$) Then skip = True
        If Not skip Then c$ = Trim$(c$)
        If e$ <> "" Then If val(e$) > val(c$) + 10 Then skip = False
        If (val(c$) Mod 10 = 0) And Not skip Then
            tr$ = String$(Len(NLTrim2(c$)), "0")
            a$ = Format((val(c$) + 10), tr$) + Mid$(CurrentParagraph, 1, MyTrimL(CurrentParagraph) - 1) + Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph), MyTrimL(Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph))) - 1)
        Else
            If e$ <> "" And Trim$(str((val(c$) + 1))) = Trim$(e$) Then
                maxw = glistN.UserControlTextWidth(c$)
                a$ = space$(Len(c$))
                While glistN.UserControlTextWidth(a$) < maxw
                a$ = a$ + " "
                Wend
                c$ = vbNullString
            Else
                tr$ = String$(Len(NLTrim2(c$)), "0")
                a$ = Format((val(c$) + 1), tr$) + Mid$(CurrentParagraph, 1, MyTrimL(CurrentParagraph) - 1) + Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph), MyTrimL(Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph))) - 1)
            End If
        End If
    Else
        GoTo there
        
    End If
Else
there:
        If HaveMarkedText Then
        ClearRedo
        PushUndo
        glistN_MarkDelete False
        undoprev = True
        End If
        Ns = NLtrim(CurrentParagraph)
        a$ = Myleft$(CurrentParagraph, Len(CurrentParagraph) - Len(Ns))
 If Len(a$) > Charpos - 1 Then
SelStartSilent = SelStart - Charpos + Len(a$) + 1
Else
If Charpos > 1 And Not PreserveRightSpaces Then
maxw = MyTrimRNoCr(Myleft$(CurrentParagraph, Charpos - 1))
If Charpos > maxw Then
och = Charpos - maxw - 1
If och > 0 Then
skip = enabled
enabled = False

SelStartSilent2 = SelStart - och
SelLengthSilent = och
InsertTextNoRender = ""
undoprev = True
enabled = skip
End If
End If
End If
 End If

End If

If a$ <> "" Then
RemoveUndo vbCrLf + a$
InsertText = vbCrLf + a$
AutoIntNewLine = True
End If
showparagraphWork = True

End Function
Public Property Let SelStartSilent2(ByVal RHS As Long)
Dim findit As Long, sel1 As Long
sel1 = SelLength
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
Dim a As Long, b As Long, c As Long, D As Long
findit = RHS
mDoc.FindPos 1, 0, findit, a, b, c, D
SetRowColumn a, b, True
If sel1 <> 0 Then SelLength = sel1: LastSelStart = sel1
End Property
Public Property Get Length() As Long
Length = mDoc.SizeCRLF
End Property
Public Sub dothis()
glistN.MarkWord
End Sub
Public Sub insertbrackets()
Dim i As Long, j As Long, ss$, idnt$, tc As Long
Dim ww As Long
If SelLength > 0 Then
ClearRedo
PushUndo
glistN_MarkDelete False
ss$ = undotext
undoprev = True
End If
Dim a$
tc = TabControl

glistN.enabled = False
    a$ = Myleft(CurrentParagraph, Charpos)

    If MyTrimL2NoTab(a$) <> 1 And tc = 0 And MyTrimL(a$) > 1 Then
        tc = Int(glistN.UserControlTextWidthPixels(Chr(9)) / glistN.UserControlTextWidthPixels("m") * 2)
    End If
    If tc = 0 Then

    j = SelStart
    i = Len(a$) - Len(NLtrim$(a$))
    idnt$ = Myleft$(a$, i)
   
    If ss$ <> "" Then
        ss$ = Replace(ss$, vbCrLf, vbCrLf + vbTab)
        RemoveUndo "{" + vbCrLf + idnt$ + vbTab + ss$ + vbCrLf + idnt$ + "}"
        InsertText = "{" + vbCrLf + idnt$ + vbTab + ss$ + vbCrLf + idnt$ + "}"
    Else
        RemoveUndo "{" + vbCrLf + idnt$ + vbTab + vbCrLf + idnt$ + "}"
        InsertText = "{" + vbCrLf + idnt$ + vbTab + vbCrLf + idnt$ + "}"
    End If
    SelStartSilent = j + Len(idnt$) + 4
    
    
    Else
    j = SelStart
    
    ww = tc
    i = Len(a$) - Len(NLtrim$(a$))
    idnt$ = Myleft$(a$, i)
   
    If ss$ <> "" Then
        ss$ = Replace(ss$, vbCrLf, vbCrLf + space$(ww))
        RemoveUndo "{" + vbCrLf + idnt$ + space$(ww) + ss$ + vbCrLf + idnt$ + "}"
        InsertText = "{" + vbCrLf + idnt$ + space$(ww) + ss$ + vbCrLf + idnt$ + "}"
    Else
        RemoveUndo "{" + vbCrLf + idnt$ + space$(ww) + vbCrLf + idnt$ + "}"
        InsertText = "{" + vbCrLf + idnt$ + space$(ww) + vbCrLf + idnt$ + "}"
    End If
    SelStartSilent = j + Len(idnt$) + 3 + ww
    End If
glistN.enabled = True
Render
End Sub
Function GetLCIDFromKeyboard() As Long
    Dim Buffer As String, ret&, r&
    Buffer = String$(514, 0)
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF&
      r = val("&H" & Right(Hex(r), 4))
        ret = GetLocaleInfo(r, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    GetLCIDFromKeyboard = CLng(val("&h" + Myleft$(Buffer, ret - 1)))
End Function
Public Property Let ReplaceTitle(a$)  ' at the end
Dim curlines As Long
NoTitle = Title = vbNullString
NoHeadwrap = False
curlines = mDoc1.DocLines
mDoc1.EmptyDoc
mDoc1.textDoc = RTrim$(Title$ + " " + a$)
mDoc1.WrapAgain
If mDoc1.DocLines >= curlines + glistN.lines + 1 Then
mDoc1.EmptyDoc
NoHeadwrap = True
mDoc1.textDoc = RTrim$(Title$ + " " + a$)
End If
glistN.HeadLine = vbNullString ' reset
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight
glistN.HeadlineHeight = mDoc1.DocLines * onelineheight
glistN.ShowMe

End Property
Public Property Let ReplaceTitleHold(a$)
ReplaceTitle = a$
TitleStay = a$ <> ""
End Property
Private Sub PrintItem2X(mHdc As Long, ByVal c As String, r As RECT, area As RECT, Col() As Long, ver1 As Boolean)
On Error Resume Next
Dim cc As String, all As Long, offset As Long, addthis As Long, hRgn As Long
Dim b As RECT, Origin As Long, tabw As Long, dummy As RECT
tabw = tParam.iTabLength * glistN.AveCharWith
Origin = r.Left
dummy = r
dummy.top = r.Bottom + 10
dummy.Bottom = dummy.top + r.Bottom - r.top
Dim ww$(), items&, i, rleft
If UBound(Col()) < 1 Then
    b = area
    hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn mHdc, hRgn
    If ver1 Then
        PrintItem1 mHdc, c, r
    Else
        PrintItem mHdc, c, r
    End If
    SelectClipRgn mHdc, &H0
    DeleteObject hRgn
    Exit Sub
End If
If UBound(Col()) = 2 Then
If Col(0) = 0 And Col(1) = 0 And Col(2) = 0 Then GoTo there
End If
If UBound(Col()) = 1 And (Col(0) = 14 Or Col(0) = 0 Or Col(0) = 7) Then
there:
    b = area
    hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn mHdc, hRgn
    If ver1 Then
        PrintItem1 mHdc, c, r
    Else
        PrintItem mHdc, c, r
    End If
    SelectClipRgn mHdc, &H0
    DeleteObject hRgn
    Exit Sub
End If
rleft = r.Left
all = rleft + glistN.UserControlTextWidthPixels(c)
r.Left = all
ww$() = SplitStringUsingArrayNoTab(c, Col(), items&)

        If showparagraphonce And ver1 Then
            cc = "Ά"
            b = area
        hRgn = CreateRectRgnIndirect(b)
        SelectClipRgn mHdc, hRgn
        DrawTextEx mHdc, StrPtr(cc), -1, r, mway, VarPtr(tParam)
        SelectClipRgn mHdc, &H0
        DeleteObject hRgn
            
        End If
    cc = vbNullString
    offset = rleft
    For i = 0 To items&
    c = ww$(i)
    r.Left = offset
    r.top = r.top + TextPixelOffset
        b = r

        If b.Right <= area.Left Or b.Left >= area.Right Then
        hRgn = CreateRectRgnIndirect(dummy)
        SelectClipRgn mHdc, hRgn
        offset = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, Origin))

        SelectClipRgn mHdc, &H0
        DeleteObject hRgn
         
        Else
        hRgn = CreateRectRgnIndirect(area)
        SelectClipRgn mHdc, hRgn
         offset = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, Origin))

        SelectClipRgn mHdc, &H0
        DeleteObject hRgn
        End If
        If r.Left > area.Right Then Exit For
        r.top = r.top - TextPixelOffset
    Next i
alfa1:
r.Left = rleft
End Sub

Private Sub PrintItem3X(mHdc As Long, ByVal c As String, r As RECT, Col() As Long)
On Error Resume Next
Dim cc As String, all As Long, offset As Long, addthis As Long
Dim b As RECT, Origin As Long, tabw As Long
tabw = tParam.iTabLength * glistN.AveCharWith
Origin = r.Left
Dim ww$(), items&, i, rleft
If UBound(Col()) < 1 Then
    PrintItem1 mHdc, c, r
    Exit Sub
End If
If UBound(Col()) = 2 Then
If Col(0) = 0 And Col(1) = 0 And Col(2) = 0 Then GoTo there
End If
If UBound(Col()) = 1 And (Col(0) = 14 Or Col(0) = 0 Or Col(0) = 7) Then
there:
    PrintItem1 mHdc, c, r
    Exit Sub
End If
rleft = r.Left
'all = rleft + glistN.UserControlTextWidthPixels(c)
'r.Left = all

ww$() = SplitStringUsingArrayNoTab(c, Col(), items&)

cc = vbNullString
offset = rleft
For i = 0 To items&
    c = ww$(i)
       r.Left = offset '+ rleft + addthis
       If r.Left > r.Right Then Exit For
    r.top = r.top + TextPixelOffset
    offset = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, Origin))
    r.top = r.top - TextPixelOffset
    
Next i
r.Left = rleft
End Sub


Private Sub PrintItem1X(mHdc As Long, ByVal c As String, r As RECT, Col() As Long)
On Error Resume Next
Dim cc As String, all As Long, offset As Long
Dim b As RECT, Origin As Long, tabw As Long
tabw = tParam.iTabLength * glistN.AveCharWith
Origin = r.Left
Dim ww$(), items&, i, rleft
If UBound(Col()) < 1 Or NoColor Then
    PrintItem1 mHdc, c, r
    Exit Sub
End If
If UBound(Col()) = 2 Then
If Col(0) = 0 And Col(1) = 0 And Col(2) = 0 Then GoTo there
End If
If UBound(Col()) = 1 And (Col(0) = 14 Or Col(0) = 0 Or Col(0) = 7) Then
there:
    PrintItem1 mHdc, c, r
    Exit Sub
End If
rleft = r.Left

ww$() = SplitStringUsingArrayNoTab(c, Col(), items&)


offset = rleft
For i = 0 To items&
    c = ww$(i)
       r.Left = offset
       If r.Left > r.Right Then Exit For
    r.top = r.top + TextPixelOffset
    
    If Not NoColor Then SetTextColor mHdc, M2000(Col(i * 2))
    offset = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, Origin))
    r.top = r.top - TextPixelOffset
    
Next i
 If r.Left < r.Right Then
If showparagraphonce Then
    cc = "Ά"
    SetTextColor mHdc, M2000(14)
    r.Left = offset
    DrawTextEx mHdc, StrPtr(cc), -1, r, mway, VarPtr(tParam)
    cc = vbNullString
End If
End If
r.Left = rleft
End Sub


Private Sub PrintItem1Xold(mHdc As Long, ByVal c As String, r As RECT, Col() As Long)
On Error Resume Next
    Dim cc As String, all As Long, CM As String, offset As Long, addthis As Long, hRgn As Long
    Dim b As RECT
    Dim ww$(), items&, i, rleft, cuts() As Long
    If UBound(Col()) <= 1 Or mNoColor Then
    If UBound(Col()) = 1 And Not mNoColor Then
    SetTextColor mHdc, M2000(Col(0))
    PrintItem1 mHdc, c, r
    Else
        PrintItem1 mHdc, c, r
        End If
        Exit Sub
    End If
    rleft = r.Left
    all = rleft + glistN.UserControlTextWidthPixels(c)
    r.Left = all
    If InStr(c, vbTab) > 0 Then
    ReDim cuts(UBound(Col()))
    ww$() = SplitStringUsingArray(c, Col(), items&, cuts())
    If showparagraphonce Then
        cc = "Ά"
        DrawTextEx mHdc, StrPtr(cc), -1, r, mway, VarPtr(tParam)
    End If
    cc = vbNullString
    
    
    For i = items& To 0 Step -1
        CM = ww$(i)
        r.Left = offset + rleft + addthis
        r.top = r.top + TextPixelOffset
        b = r
        If i > 0 Then
        
            b.Left = rleft + glistN.UserControlTextWidthPixels(Myleft$(c, cuts(i - 1))) + 1
          If b.Left < 0 Then b.Left = 0
          
        
        End If

        If Not mNoColor1 Then SetTextColor mHdc, M2000(Col(i * 2))
         hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn mHdc, hRgn
        DrawTextEx mHdc, StrPtr(CM), -1, r, mway, VarPtr(tParam)
          SelectClipRgn mHdc, &H0
   DeleteObject hRgn
           r.top = r.top - TextPixelOffset
   
    Next i
Else
    ww$() = SplitStringUsingArrayNoTab(c, Col(), items&)
    If showparagraphonce Then
        cc = "Ά"
        DrawTextEx mHdc, StrPtr(cc), -1, r, mway, VarPtr(tParam)
    End If
    cc = vbNullString
    CM = c
    For i = items& To 0 Step -1
        c = ww$(i)
        CM = Myleft$(CM, Len(CM) - Len(ww$(i)))
        offset = glistN.UserControlTextWidthPixels(CM)
        ' for Kenring
        addthis = glistN.UserControlTextWidthPixels(CM + c) - offset - glistN.UserControlTextWidthPixels(c)
        r.Left = offset + rleft + addthis
        r.top = r.top + TextPixelOffset
        
        If Not mNoColor1 Then SetTextColor mHdc, M2000(Col(i * 2))
    
        DrawTextEx mHdc, StrPtr(c), -1, r, mway, VarPtr(tParam)
        r.top = r.top - TextPixelOffset
        If r.Left < 0 Then Exit For
    Next i

End If
    r.Left = rleft
    End Sub


Function SplitStringUsingArray(ByVal c$, bb() As Long, wc&, cuts() As Long) As String()
Dim cc$()
ReDim cc$((UBound(bb()) + 1) / 2)
Dim st&, ste&, W&
wc& = 0
ste& = UBound(bb())
For W& = 1 To ste& - 2 Step 2
On Error Resume Next
If bb(W& + 2) - bb(W&) < 1 Then
cc$(wc&) = Mid$(c$, 1, bb(W&))
cuts(wc&) = bb(W& + 2)
Else
If wc& > 0 Then
cc$(wc&) = cc$(wc& - 1) + Mid$(c$, bb(W&), bb(W& + 2) - bb(W&))
Else
cc$(wc&) = Mid$(c$, bb(W&), bb(W& + 2) - bb(W&))
End If
cuts(wc&) = Len(cc$(wc&))
End If
If bb(W& + 2) - bb(W&) = 0 Then
wc& = wc& - 1
SplitStringUsingArray = cc$()
Exit Function
End If
wc& = wc& + 1
Next W&
If bb(W&) > 0 Then
If Trim$(Mid$(c$, bb(W&))) <> vbNullString Then

cc$(wc&) = c$ 'Mid$(c$,  bb(w&))
cuts(wc&) = Len(c$)
Else
 wc& = wc& - 1
End If
End If
SplitStringUsingArray = cc$()
End Function
Function SplitStringUsingArrayNoTab(c$, bb() As Long, wc&) As String()
Dim cc$()
ReDim cc$((UBound(bb()) + 1) / 2)
Dim st&, ste&, W&
wc& = 0
ste& = UBound(bb())
For W& = 1 To ste& - 2 Step 2
On Error Resume Next
If bb(W& + 2) - bb(W&) < 1 Then
cc$(wc&) = Mid$(c$, bb(W&))
Else
cc$(wc&) = Mid$(c$, bb(W&), bb(W& + 2) - bb(W&))
End If
If cc$(wc&) = vbNullString Then
wc& = wc& - 1
SplitStringUsingArrayNoTab = cc$()
Exit Function
End If
wc& = wc& + 1
Next W&
If bb(W&) > 0 Then
cc$(wc&) = Mid$(c$, bb(W&))
End If
If cc$(wc&) = vbNullString Then wc& = wc& - 1
SplitStringUsingArrayNoTab = cc$()
End Function

Function M2000(ch As Long) As Long
Dim b As Integer
If M2000code Then
Select Case ColorSet
Case 0
    If ch = 13 Then
    'M2000 = rgb(&HEE, &HEE, &H33)
    M2000 = rgb(&HC5, &HF7, &HF7)
    ElseIf ch = 12 Then
    M2000 = rgb(&HDD, &HAA, 0)
    Else
    If ch = 0 Then
    M2000 = rgb(32, 32, 255)
    Else
    M2000 = QBColor(ch Mod 16)
    End If
    End If
Case 1
    Select Case ch
    Case 7
    M2000 = rgb(&H68, &H68, &H13)
    Case 10
    M2000 = 0   '' black  '' basic
    Case 11
    M2000 = rgb(32, 32, 255)
    Case 12
    M2000 = rgb(&H80, &H40, &H10)
    Case 13
    M2000 = rgb(&HAA, &H90, &H20)
    Case 14
    M2000 = rgb(&HAA, &H60, &H30)
    Case 15
    M2000 = rgb(&H63, &H30, &H63) '' string
    End Select
Case Is < 0
M2000 = ColArr(CLng(ch) Mod 16)
End Select
Else
Select Case ColorSet
    Case 0
    
    If ch = 0 Then
    M2000 = rgb(225, 225, 255)
    Else
    M2000 = QBColor(ch Mod 16)
    End If
    
    Case 1
        Select Case ch
        Case 7
        M2000 = rgb(&H68, &H68, &H13)
        Case 10
        M2000 = 0   '' black  '' basic
        Case 11
        
        M2000 = rgb(32, 32, 255)
        Case 12
        M2000 = rgb(&H80, &H40, &H10)
        Case 13
        M2000 = rgb(&HAA, &H90, &H20)
        Case 14

        M2000 = rgb(&HAA, &H60, &H30)
        Case 15
        M2000 = rgb(&H63, &H30, &H63) '' string
        
        End Select
    Case Is < 0

    M2000 = ColArr(CLng(ch) Mod 16)

End Select
End If
End Function
Public Sub UserColorSet(ResetTable, Optional col7, Optional col15, Optional col14, Optional col10, Optional col11, Optional col12, Optional col0, Optional col13)
Dim i As Integer
If CBool(ResetTable) Or Not ColorSet = -1 Then
If CBool(ResetTable) Then
    For i = 0 To 15: ColArr(CLng(i)) = QBColor(i): Next i
Else
    For i = 0 To 15: ColArr(CLng(i)) = M2000(CLng(i)): Next i
End If
End If
If Not IsMissing(col7) Then ColArr(7) = col7
If Not IsMissing(col15) Then ColArr(15) = col15
If Not IsMissing(col10) Then ColArr(10) = col10
If Not IsMissing(col11) Then ColArr(11) = col11
If Not IsMissing(col12) Then ColArr(12) = col12
If Not IsMissing(col0) Then ColArr(0) = col0  ' not processed line
If Not IsMissing(col13) Then ColArr(13) = col13
If Not IsMissing(col14) Then ColArr(14) = col14
' need a colorset to -1 value
End Sub
Function htmlM2000(ch) As Long
Dim b As Integer

Select Case Abs(pagiohtml$ = "DARK")
Case 0

htmlM2000 = (QBColor(ch Mod 16) And &HFF00&) + (QBColor(ch Mod 16) \ &H10000) + ((QBColor(ch Mod 16) And &HFF) * &H10000)
Case 1
Select Case ch
Case 7
htmlM2000 = rgb(&H13, &H68, &H68)
Case 10
htmlM2000 = 0   '' black  '' basic
Case 11
htmlM2000 = rgb(255, 32, 32)
Case 14
htmlM2000 = rgb(&H30, &H60, &HAA)
Case 15
htmlM2000 = rgb(&H63, &H30, &H63) '' string

End Select
End Select

End Function
Private Function aheadstatus2000(a$, Optional srink As Boolean = True, Optional Pos As Long = 1) As String 'ok
Dim b$, part$, W$, pos2 As Long, word$

If a$ = vbNullString Then Exit Function
If Pos = 0 Then Pos = 1
Do While Pos <= Len(a$)

    W$ = Mid$(a$, Pos, 1)
    If part$ = vbNullString And W$ = "0" Then
        If Pos + 2 <= Len(a$) Then
            If LCase(Mid$(a$, Pos, 2)) Like "0[xχ]" Then
            'hexadecimal literal number....
                Pos = Pos + 2
                Do While Pos <= Len(a$)
                If Not Mid$(a$, Pos, 1) Like "[0-9a-fA-F]" Then Exit Do
                Pos = Pos + 1
                Loop
                If Mid$(a$, Pos, 1) = "&" Then
                If Mid$(a$, Pos + 1, 1) = "&" Then Pos = Pos + 2 Else Pos = Pos + 1
                End If
                b$ = b$ & "N"
                If Pos <= Len(a$) Then
                    W$ = Mid$(a$, Pos, 1)
                Else
                    Exit Do
                End If
            End If
        End If
    End If

    If W$ = Chr(34) Then
    
        If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        Pos = Pos + 1
        Do While Pos <= Len(a$)
       ' If Mid$(a$, Pos, 1) = Chr(34) And Mid$(a$, Pos - 1, 1) <> "`" Then Exit Do
       If Mid$(a$, Pos, 1) = Chr(34) Then Exit Do
    If Asc(Mid$(a$, Pos, 1)) < 32 Then Exit Do
   
        Pos = Pos + 1
        Loop
    ElseIf W$ = "(" Or W$ = "[" Then
        If part$ <> "" Then
            ' after
            ElseIf Right$(b$, 1) = "a" Then
            b$ = Myleft(b$, Len(b$) - 1)
            part$ = vbNullString
            Else
            part$ = "N"
              
        End If
      Pos = Pos + 1
        If W$ = "(" Then
            If Not BlockParam2(a$, Pos, 1, 0) Then Exit Do
            If Mid$(a$, Pos + 1, 1) <> "." Then b$ = b$ + part$
            part$ = vbNullString
        Else
            If Not BlockParam2(a$, Pos, 0, 1) Then Exit Do
            If Len(word$) = 0 Then
            
            word$ = Mid$(a$, 1, Pos)
            b$ = b$ + part$
            part$ = "N"
            End If
        End If
       
        
        
    ElseIf W$ = "{" Then
If word$ <> "" Then
  Select Case myUcase(word$, True)
  Case "THEN", "ELSE", "ΤΟΤΕ"
  Pos = Pos - 4
  Exit Do
  Case "ELSE.IF"
  Pos = Pos - 7
  Exit Do
  Case "ΑΛΛΙΩΣ"
  Pos = Pos - 6
  Exit Do
  Case "ΑΛΛΙΩΣ.ΑΝ"
  Pos = Pos - 9
  Exit Do
  End Select
  End If
    

    If part$ <> "" Then b$ = b$ & part$
        If LenB(b$) = 0 Then
        ElseIf Right(b$, 1) <> "l" Then
        Exit Do
        End If
 
        blockStringAhead a$, Pos
        Pos = Pos + 1
        part$ = "S"
        

            If Pos <= Len(a$) Then
        If Not blockStringAhead(a$, Pos) Then Exit Do
        Pos = Pos - 1
        End If
''Exit Do

    Else
        Select Case W$
        Case ","  ' bye bye
        Exit Do
        Case "%"
            If part$ = vbNullString Then
            ' lathos
            End If
        Case "$"
            If part$ = vbNullString Then
                If b$ = vbNullString Then
                    part$ = "N"
                ElseIf Right$(b$, 1) = "o" Then
                    part$ = "N"
                Else
        
                    aheadstatus2000 = b$
                    Exit Function
                End If
            ElseIf part$ = "N" Then
                    b$ = b$ & "Sa"
                    part$ = vbNullString
            End If

        Case "-", "+", "|"
                    b$ = b$ & part$
                    If b$ = vbNullString Then
                        part$ = "o"
                    Else
                    
                part$ = "o"
                End If
        Case "/"
            If Mid$(a$, Pos + 1, 1) = "/" Then Exit Do
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case "*", "^"
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case " ", ChrW(160)
          If word$ <> "" Then
  Select Case myUcase(word$, True)
  Case "THEN", "ELSE", "ΤΟΤΕ"
  Pos = Pos - 4
  Exit Do
  Case "ELSE.IF"
  Pos = Pos - 7
  Exit Do
  Case "ΑΛΛΙΩΣ"
  Pos = Pos - 6
  Exit Do
  Case "ΑΛΛΙΩΣ.ΑΝ"
  Pos = Pos - 9
  Exit Do
  End Select
  End If
JUMPHERE:
        word$ = vbNullString
            If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            Else
            'skip
            End If
        Case "0" To "9", "."
            If part$ = "N" Then
            If Len(a$) < Pos Then
                If Mid$(a$, Pos + 1, 1) Like "[@#%~]" Then
                    Pos = Pos + 1
                ElseIf Mid$(a$, Pos + 1, 1) = "&" Then
                    If Mid$(a$, Pos + 2, 1) = "&" Then Pos = Pos + 2 Else Pos = Pos + 1
                End If
            End If
            ElseIf part$ = "S" Then
            
            Else
            b$ = b$ & part$
            part$ = "N"
            End If
        Case "&"
        If part$ = vbNullString Then
        part$ = "S"
        Else
        b$ = part$
        part$ = "S"
        End If
        Case "e", "E", "ε", "Ε"
        If LenB(word$) = 0 Then part$ = "N"
        
        word$ = word$ + W$
           If part$ = "S" Then
            
            
            ElseIf part$ = "N" Then
            
            b$ = b$ & part$
            'part$ = "N"
            End If
            
          Case "~"
        GoTo there1
        Case ">", "<"
            If Len(a$) > Pos Then
                If Mid$(a$, Pos + 1, 1) = "=" Then
                     b$ = b$ & part$
                     If Len(b$) > 1 Then
                        part$ = "o"
                        If Mid$(a$, Pos, 3) = "<=>" Then
                            Pos = Pos + 2
                        Else
                            Pos = Pos + 1
                        End If
                    End If
                ElseIf Mid$(a$, Pos, 2) = "<>" Then
                    b$ = b$ & part$
                    If Len(b$) > 1 Then
                        part$ = "o"
                        Pos = Pos + 1
                    End If
                ElseIf W$ = ">" And Pos > 1 Then
                    If Mid$(a$, Pos - 1, 2) = "->" Then ' "->"
                        If Right$(b$, 1) = "S" Then
                            b$ = b$ + part$
                            part$ = "N"
                        End If
                    End If
                End If
            End If
            GoTo there1
        Case "="
            If Len(a$) > Pos Then
                If InStr(">=", Mid$(a$, Pos + 1, 1)) > 0 Then
                    Pos = Pos + 2
                    GoTo conthere
                End If
            End If
there1:
                If b$ & part$ <> "" Then
                W$ = Replace(b$ & part$, "a", "")
            part$ = vbNullString
               If srink Then
                  Do
                b$ = W$
                W$ = Replace(b$, "NN", "N")
                Loop While W$ <> b$
                         Do
                        b$ = W$
                          W$ = Replace(b$, "SlS", "N")
                          Loop While W$ <> b$
                            Do
                          b$ = W$
                          W$ = Replace(b$, "NlN", "N")
                          Loop While W$ <> b$
    
                Do
                b$ = W$
                W$ = Replace(b$, "NoN", "N")
                Loop While W$ <> b$
                
                Do
                b$ = W$
                W$ = Replace(b$, "SoS", "S")
                Loop While W$ <> b$
                Else
              b$ = W$
               End If
               
                If Myleft$(b$, Len(b$) - 1) <> "l" Then part$ = "l"
                Else
                Exit Do
                End If
        
        Case ")", "}", Is < " ", ":", ";", "'", "\", "/", "]"
            Exit Do
        Case Else
            If part$ = "N" Then
                word$ = word$ + W$
            ElseIf part$ = "S" Then
            Else
                b$ = b$ & part$
                part$ = "N"
                word$ = word$ + W$
            End If
           
        End Select
        End If
        
    Pos = Pos + 1
conthere:
Loop

    W$ = Replace(b$ & part$, "a", "")
    
    b$ = W$
If srink Then
         Do
  b$ = W$

    W$ = Replace(b$, "SlS", "N")
    Loop While W$ <> b$
      Do
    b$ = W$
    W$ = Replace(b$, "NlN", "N")
    Loop While W$ <> b$
    
    Do
    b$ = W$
    W$ = Replace(b$, "NoN", "N")
    Loop While W$ <> b$
    
    Do
    b$ = W$
    W$ = Replace(b$, "SoS", "S")
    Loop While W$ <> b$
End If
   
   
 
    aheadstatus2000 = b$




End Function


Private Function aheadstatus(a$, Optional srink As Boolean = True, Optional Pos As Long = 1) As String 'ok
Dim b$, part$, W$, pos2 As Long, word$

If a$ = vbNullString Then Exit Function
If Pos = 0 Then Pos = 1
Do While Pos <= Len(a$)
    W$ = Mid$(a$, Pos, 1)
    If part$ = vbNullString And W$ = "0" Then
        If Pos + 2 <= Len(a$) Then
            If LCase(Mid$(a$, Pos, 2)) Like "0[xχ]" Then
            'hexadecimal literal number....
                Pos = Pos + 2
                Do While Pos <= Len(a$)
                If Not Mid$(a$, Pos, 1) Like "[0-9a-fA-F]" Then Exit Do
                Pos = Pos + 1
                Loop
                b$ = b$ & "N"
                If Pos <= Len(a$) Then
                    W$ = Mid$(a$, Pos, 1)
                Else
                    Exit Do
                End If
            End If
        End If
    End If
        
        If W$ = StringSep Then
            If part$ <> "" Then
                b$ = b$ & part$
                part$ = vbNullString
            ElseIf W$ = Myleft$(stringsep2, 1) Then
                If part$ <> "" Then
                b$ = b$ & part$
                End If
                If Not (part$ = vbNullString Or part$ = "o") Then Exit Do
                part$ = "S"
                Pos = Pos + 1
                Do While Pos <= Len(a$)
                    If Pos > 1 Then
                    If Mid$(a$, Pos - 1, 2) = "\" + Right$(stringsep2, 1) Then GoTo cont001
                    
                    End If
                    If Mid$(a$, Pos, 1) = Right$(stringsep2, 1) Then Exit Do
                    If Asc(Mid$(a$, Pos, 1)) < 32 Then Exit Do
cont001:
                    Pos = Pos + 1
                    
                Loop
            If Pos > Len(a$) Then Pos = Len(a$)
            End If
            If Len(OpenTag) = 0 Then If Not (part$ = vbNullString Or part$ = "o") Then Exit Do
            part$ = "S"
            Pos = Pos + 1
            Do While Pos <= Len(a$)
                     If Pos > 1 Then
                    If Mid$(a$, Pos - 1, 2) = "\" + Right$(StringSep, 1) Then GoTo cont002
                    
                    End If
                

                If Mid$(a$, Pos, 1) = StringSep Then Exit Do
                If Asc(Mid$(a$, Pos, 1)) < 32 Then Exit Do
cont002:
                Pos = Pos + 1
            Loop
            If Pos > Len(a$) Then Pos = Len(a$)
        ElseIf W$ = Myleft$(stringsep2, 1) Then
            If part$ <> "" Then
                b$ = b$ & part$
            End If
            If Len(OpenTag) = 0 Then If Not (part$ = vbNullString Or part$ = "o") Then Exit Do
            part$ = "S"
            Pos = Pos + 1
            Do While Pos <= Len(a$)
                If Mid$(a$, Pos, 1) = Right$(stringsep2, 1) Then Exit Do
                If Asc(Mid$(a$, Pos, 1)) < 32 Then Exit Do
                Pos = Pos + 1
            Loop
            If Pos > Len(a$) Then Pos = Len(a$)
           
        
        ElseIf W$ = "(" Or W$ = "[" Then
        
        If part$ <> "" Then
            ' after
            ElseIf Right$(b$, 1) = "a" Then
            b$ = Myleft$(b$, Len(b$) - 1)
            part$ = vbNullString
            Else
            part$ = "N"
              
        End If
      Pos = Pos + 1
        If Not BlockParam4(a$, Pos, Asc(W$)) Then Pos = Len(a$) + 1: Exit Do
       If Mid$(a$, Pos + 1, 1) <> "." Then
       b$ = b$ & part$
        Else
        Pos = Pos - 1
        End If
        part$ = vbNullString
   Else
    If Len(OpenTag) = 0 Then
        Select Case W$
        Case mSplitExpr  ' bye bye
        Exit Do
        Case "%"
            If part$ = vbNullString Then
            ' lathos
            End If
        Case "$"
            If part$ = vbNullString Then
                If b$ = vbNullString Then
                    part$ = "N"
                ElseIf Right$(b$, 1) = "o" Then
                    part$ = "N"
                Else
        
                    aheadstatus = b$
                    Exit Function
                End If
            ElseIf part$ = "N" Then
                    b$ = b$ & "Sa"
                    part$ = vbNullString
            End If
        Case "+", "-"
                    b$ = b$ & part$
                    If b$ = vbNullString Then
                        part$ = "o"
                    Else
                    
                part$ = "o"
                End If
        Case "/", "^"
            If Mid$(a$, Pos + 1, 1) = "/" Then Exit Do
            If Mid$(a$, Pos + 1, 1) = "/*" Then Exit Do
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case "^"
            If Mid$(a$, Pos + 1, 1) = "^" Then Exit Do
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
            
        Case "*"
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case " ", ChrW(160)
        word$ = vbNullString
            If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            Else
            'skip
            End If
        Case "0" To "9", "."
            If part$ = "N" Then
            ElseIf part$ = "S" Then
            
            Else
            b$ = b$ & part$
            part$ = "N"
            End If
        Case "&"
        If part$ = vbNullString Then
        part$ = "S"
        Else
        b$ = part$
        part$ = "S"
        End If
   
        Case "e", "E", "ε", "Ε"
        word$ = word$ + W$
           If part$ = "S" Then
            
            
            ElseIf part$ = "N" Then
            
            b$ = b$ & part$
            'part$ = "N"
            End If
        Case "=", ">", "<", "~"
         If b$ & part$ = "" And InStr(StartSymbols + Chr$(0), W$) > 0 Then
         Pos = Pos + 1
         b$ = "N"
         
         Exit Do
         End If
                If b$ & part$ <> "" Then
                W$ = Replace(b$ & part$, "a", "")
            part$ = vbNullString
               If srink Then
                  Do
                b$ = W$
                W$ = Replace(b$, "NN", "N")
                Loop While W$ <> b$
                         Do
                        b$ = W$
                          W$ = Replace(b$, "SlS", "N")
                          Loop While W$ <> b$
                            Do
                          b$ = W$
                          W$ = Replace(b$, "NlN", "N")
                          Loop While W$ <> b$
    
                Do
                b$ = W$
                W$ = Replace(b$, "NoN", "N")
                Loop While W$ <> b$
                
                Do
                b$ = W$
                W$ = Replace(b$, "SoS", "S")
                Loop While W$ <> b$
                Else
              b$ = W$
               End If
               If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            End If
               If Myleft$(b$, Len(b$) - 1) <> "l" Then part$ = "l"
               
                Else
                Exit Do
                End If
        Case ")", "]"
        
        If Mid$(a$, Pos + 1, 1) = "." Then
        Pos = Pos + 1
                   If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            End If
        Else
        Exit Do
        End If
        Case Is < " ", LineCommandSep, LineComment1, LineComment2
        Exit Do
        Case Else
    
         If part$ = "N" Then
        word$ = word$ + W$
        ElseIf part$ = "S" Then
       
        Else
        
     b$ = b$ & part$
     part$ = "N"
     word$ = word$ + W$
            End If
        End Select
        Else
         'If Len(DropSym) > 0 Then If InStr(DropSym, w$) > 0 Then Exit Do
        End If
     End If
    Pos = Pos + 1
 Loop

    W$ = Replace(b$ & part$, "a", "")
    
    b$ = W$
If srink Then
         Do
  b$ = W$

    W$ = Replace(b$, "SlS", "N")
    Loop While W$ <> b$
      Do
    b$ = W$
    W$ = Replace(b$, "NlN", "N")
    Loop While W$ <> b$
    
    Do
    b$ = W$
    W$ = Replace(b$, "NoN", "N")
    Loop While W$ <> b$
    
    Do
    b$ = W$
    W$ = Replace(b$, "SoS", "S")
    Loop While W$ <> b$
End If
aheadstatus = b$
End Function

Public Sub move(X, Y, Optional W, Optional H)
glistN.HideBarAsap
If IsMissing(W) Then
glistN.MoveTwips CLng(X), CLng(Y), 0, 0
ElseIf IsMissing(H) Then

glistN.MoveTwips CLng(X), CLng(Y), CLng(W), 0
RelocateMarkedText
Else
glistN.MoveTwips CLng(X), CLng(Y), CLng(W), CLng(H)
RelocateMarkedText
End If


End Sub
Public Sub deconstruct()
If Not glistN Is Nothing Then glistN.Shutdown
Set Callback = Nothing
Set glistN = Nothing
Set mDoc = Nothing
Set mDoc1 = Nothing
End Sub
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.Name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set glistN.Font = aa
TextPixelOffset = Int(glistN.TextHeightOffset)
End If
End Sub

Public Property Get locked() As Boolean
locked = Not glistN.enabled
End Property

Public Property Let locked(ByVal RHS As Boolean)
glistN.enabled = Not RHS
End Property
Public Property Get Index() As Long
Index = mIndex
End Property

Friend Property Let Index(ByVal RHS As Long)
mIndex = RHS
End Property

Public Property Get LineSpace() As Variant
LineSpace = dv15 * (glistN.addpixels) / 2
End Property

Public Property Let LineSpace(ByVal vNewValue As Variant)
    vNewValue = (vNewValue \ 30) * 60
On Error Resume Next
glistN.addpixels = CLng(vNewValue / dv15)
If Err.Number > 0 Then Exit Property
Render
End Property

Public Property Get LineHeight() As Variant
''
LineHeight = glistN.HeightTwips / glistN.lines
End Property
Public Sub PressKey(KeyCode As Integer, shift As Integer)
Dim b As Boolean
  b = enabled
    PressKey_A KeyCode, shift
enabled = b
  
End Sub
Private Sub PressKey_A(KeyCode As Integer, shift As Integer)
Static ctrl As Boolean, alt As Boolean, noentrance As Boolean, where As Long, noinp As Double
Dim aa$, a$, JJ As Long, ii As Long
If alt = True Then shift = shift Or 4: alt = False
If noentrance Then
KeyCode = 0
Exit Sub
End If
noentrance = True
Select Case KeyCode
Case vbKeyReturn
nochange = True
If UsedAsTextBox Or NumberOnly Or NumberIntOnly Then
KeyCode = 0
If Not Callback Is Nothing Then
With Callback.Control(1)
.PressKey 13, 0
End With
End If
ElseIf (shift And 2) = 2 And AllowInsertBrackets Then
KeyCode = 0
shift = 0
insertbrackets
ElseIf AutoIntNewLine Then
KeyCode = 0
Else

End If

nochange = False

Case vbKeyControl
ctrl = True
alt = (shift And 4) = 4
KeyCode = 0
Case vbKeyF1
If shift > 0 Then If (shift And 2) <> 2 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub

If (shift And 2) = 2 Then
shift = 0
If SelText <> "" Then
helpmeSub
Else

  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
End If
Else
nowrap = Not nowrap
Render
ManualInform
End If

KeyCode = 0
Case vbKeyF2
If (shift And 4) = 4 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub

If shift <> 0 Then
LastSearchType = 2 - Abs(shift Mod 2 = 1)
shift = 0
If SelText <> "" Then s$ = SelText
If pagio$ = "GREEK" Then
s$ = InputBoxN("Αναζήτησε προς την αρχή " & IIf(LastSearchType = 1, "σειρά χαρακτήρων" & vbCrLf & " Ctrl+F2 αναζήτηση λέξης", "λέξη" & vbCrLf & " Shift+F2 αναζήτηση χαρακτήρων"), "Συγγραφή Κειμένου", s$, noinp)

Else
s$ = InputBoxN("Search to the top " & IIf(LastSearchType = 1, " characters" & vbCrLf & " Ctrl+F2 search word", "word" & vbCrLf & " Shift+F2 search characters/symbols"), "Text Editor", s$, noinp)
End If
If Not noinp = 1 Then
noentrance = False
KeyCode = 0
Exit Sub
End If
End If
If noinp = 1 Then
    If MyTrim(s$) <> "" Then Searchup s$, LastSearchType = 1 Else LastSearchType = 0
    shift = 0
ElseIf SelText <> "" Or s$ <> "" Then
If LastSearchType > 0 Then
    If SelText <> "" Then s$ = SelText
        Searchup s$, LastSearchType = 1
    Else
        supsub
    End If
End If
KeyCode = 0
Case vbKeyF3
If (shift And 4) = 4 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub
    If shift <> 0 Then
    shift = 0
        If s$ = vbNullString And SelText <> "" Then s$ = SelText
        If pagio$ = "GREEK" Then
            s$ = InputBoxN("Αναζήτησε προς το τέλος " & IIf(LastSearchType = 1, "σειρά χαρακτήρων" & vbCrLf & " Ctrl+F3 αναζήτηση λέξης", "λέξη" & vbCrLf & " Shift+F3 αναζήτηση χαρακτήρων"), "Συγγραφή Κειμένου", s$, noinp)
        Else
            s$ = InputBoxN("Search to the end " & IIf(LastSearchType = 1, " characters" & vbCrLf & " Ctrl+F3 search word", "word" & vbCrLf & " Shift+F3 search characters/symbols"), "Text Editor", s$, noinp)
        End If
        If Not noinp = 1 Then
            noentrance = False
            KeyCode = 0
            Exit Sub
        End If
    End If
    If noinp = 1 Then
    If MyTrim(s$) <> "" Then SearchDown s$, LastSearchType = 1 Else LastSearchType = 0
    shift = 0
    ElseIf SelText <> "" Or s$ <> "" Then
        If LastSearchType > 0 Then
            If SelText <> "" Then s$ = SelText
            SearchDown s$, LastSearchType = 1
        Else
            sdnSub
        End If
    End If
Case vbKeyF4
If (shift And 7) > 0 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub
If SelText <> "" Then mscatsub Else dothis
shift = 0
KeyCode = 0
Case vbKeyF5
If (shift And 7) > 0 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub
If SelText <> "" Then rthissub shift
KeyCode = 0
shift = 0
Case vbKeyF6  ' Set/Show/Reset Para1
If (shift And 7) > 0 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub
If Not (UsedAsTextBox Or NumberOnly Or NumberIntOnly) Then MarkSoftButton para1, PosPara1
KeyCode = 0
shift = 0
Case vbKeyF7  'Set/Show/Reset Para2
If (shift And 7) > 0 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub
If Not (UsedAsTextBox Or NumberOnly Or NumberIntOnly) Then MarkSoftButton Para2, PosPara2
KeyCode = 0
shift = 0
Case vbKeyF8  'Set/Show/Reset Para2
If (shift And 7) > 0 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub
If Not (UsedAsTextBox Or NumberOnly Or NumberIntOnly) Then MarkSoftButton Para3, PosPara3
KeyCode = 0
shift = 0
Case vbKeyF9  ' Count Words
'If (shift And 7) > 1 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub
If (shift And 7) = 1 Then
shift = 0
NoCenterLineEdit = Not NoCenterLineEdit
If UserCodePage = 1253 Then
If NoCenterLineEdit Then
ReplaceTitle = "Ελεύθερη διόρθωση στο κείμενο"
Else
ReplaceTitle = "Διόρθωση στη κεντρική γραμμή"
End If
Else
If NoCenterLineEdit Then
ReplaceTitle = "Free line edit mode"
Else
ReplaceTitle = "Center line edit mode"
End If
End If

Else
If Not (UsedAsTextBox Or NumberOnly Or NumberIntOnly) Then
If glistN.lines > 1 Then
If UserCodePage = 1253 Then
ReplaceTitle = "Λέξεις στο κείμενο:" + str(mDoc.WordCount)
Else
ReplaceTitle = "Words in text:" + str(mDoc.WordCount)
End If
End If
End If
End If
KeyCode = 0
Case vbKeyF10
If (shift And 7) > 1 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub

If Not (UsedAsTextBox Or NumberOnly Or NumberIntOnly) Then
showparagraph = Not showparagraph
mDoc.WrapAgain
Render
End If
ManualInform
KeyCode = 0
Case vbKeyF11
If (shift And 7) = 0 Then KeyCode = 0: shift = 0: noentrance = False: Exit Sub
fState = fState + 1
SetME
WrapAll
ManualInform
glistN.ShowMe
KeyCode = 0
Case vbKeyPageUp
Case vbKeyPageDown
Case vbKeyTab

If UseTab And Not BlockStartPos <> BlockEndPos And (TabControl = 0 Or MyTrimLNoTab(Mid$(CurrentParagraph, 1, Charpos)) < Charpos) Then
    shift = shift And &H1
    ctrl = False
    alt = False
    JJ = SelStart
    If shift = 0 Then
    If showparagraph Then
    showparagraph = False
    'SelStartSilent = JJ
    RemoveUndo ChrW(9)
    InsertText2 = ChrW(9)
     showparagraph = True
     SelStart = JJ + 1
    Else
    SelStartSilent = JJ
        RemoveUndo ChrW(9)
        InsertText = ChrW(9)
         
          SelStart = JJ + 1
        End If
       
         KeyCode = 0
    ElseIf NLtrim(Mid$(CurrentParagraph, Charpos, 1)) = vbNullString Then
        SelStart = SelStart + 1
        KeyCode = 8
    ElseIf NLtrim(Myleft$(CurrentParagraph, Charpos - 1)) = vbNullString Then
        KeyCode = 8
    End If
    'End If
   
    shift = 0
 
ElseIf TabControl > 0 Then
conttab:
    If Not (UsedAsTextBox Or NumberOnly Or NumberIntOnly) Then
    nochange = True
    shift = shift And &H1
    glistN.enabled = False
    
    If HaveMarkedText Then SelStartSilent = SelStart
    JJ = SelStart
    where = JJ
    ii = 1 + SelStart - ParaSelStart
    
    If SelLength > 0 Then
    
    JJ = SelLength + JJ - ii
    SelStart = ii
    SelLength = JJ
    JJ = where
    Else
    SelStart = ii
    End If
    
    
    If SelText <> "" Then
    
        a$ = vbCrLf + SelText & "*"
        If shift <> 0 Then
        shift = 0
        a$ = Replace(a$, vbCrLf + String(TabControl + (Len(CurrentParagraph) - Len(LTrim(CurrentParagraph))) Mod TabControl, ChrW(160)), vbCrLf)
            a$ = Replace(a$, vbCrLf + space$(TabControl + (Len(CurrentParagraph) - Len(LTrim(CurrentParagraph))) Mod TabControl), vbCrLf)
            InsertTextNoRender = Mid$(a$, 3, Len(a$) - 3)
             SelStartSilent = ii
             SelLengthSilent = Len(a$) - 3
           
        Else
            a$ = Replace(a$, vbCrLf, vbCrLf + space$(TabControl))
            InsertTextNoRender = Mid$(a$, 3, Len(a$) - 3)
            SelStartSilent = where + TabControl
            SelLengthSilent = Len(a$) - 3 - (where + TabControl - ii)
      
        End If
        
    Else
    If shift <> 0 Then
    
      shift = 0
        If MyTrimL2(Mid$(CurrentParagraph, 1, TabControl)) = 0 Then
    
                SelStartSilent = ii
                SelLengthSilent = TabControl
                InsertTextNoRender = vbNullString
                SelStartSilent = ii
        Else
                SelStartSilent = ii
                SelLengthSilent = Len(CurrentParagraph) - Len(LTrim(CurrentParagraph))
                InsertTextNoRender = vbNullString
                SelStartSilent = ii
        
        End If
        Else
            SelStartSilent = JJ
            RemoveUndo space(TabControl)
            InsertText = space(TabControl)
            
            SelStartSilent = where + TabControl
        End If
    End If
    glistN.enabled = True
        ReColorBlock
    Render
    nochange = False
    End If
    KeyCode = 0  ' new
    shift = 0 ' new
ElseIf BlockStartPos <> BlockEndPos And HaveMarkedText Then
conttab1:
If Not (UsedAsTextBox Or NumberOnly Or NumberIntOnly) Then
nochange = True
shift = shift And &H1
glistN.enabled = False

If HaveMarkedText Then SelStartSilent = SelStart
JJ = SelStart
where = JJ
ii = 1 + SelStart - ParaSelStart

If SelLength > 0 Then

JJ = SelLength + JJ - ii
SelStart = ii
SelLength = JJ
JJ = where
Else
SelStart = ii
End If
' code for adding/removing tabs from left, usefull for code

If SelText <> "" Then
    a$ = vbCrLf + SelText & "*"
    If shift <> 0 Then
    shift = 0
    a$ = Replace(a$, vbCrLf + ChrW(9), vbCrLf)
       ' a$ = Replace(a$, vbCrLf + String(1 + (Len(CurrentParagraph) - Len(LTrim(CurrentParagraph))), ChrW(9)), vbCrLf)
        InsertTextNoRender = Mid$(a$, 3, Len(a$) - 3)
         SelStartSilent = ii
         SelLengthSilent = Len(a$) - 3
       
    Else
        a$ = Replace(a$, vbCrLf, vbCrLf + ChrW(9))
        InsertTextNoRender = Mid$(a$, 3, Len(a$) - 3)
        SelStartSilent = where + 1
        SelLengthSilent = Len(a$) - 3 - (where + 1 - ii)
  
    End If
    

Else
If shift <> 0 Then
shift = 0
    If Mid$(CurrentParagraph, 1, 1) = ChrW(9) Or Mid$(CurrentParagraph, 1, 1) = ChrW(9) Then

            SelStartSilent = ii
            SelLengthSilent = 1
            InsertTextNoRender = vbNullString
            SelStartSilent = ii
    Else
            SelStartSilent = ii
            SelLengthSilent = Len(CurrentParagraph) - Len(LTrim(CurrentParagraph))
            InsertTextNoRender = vbNullString
            SelStartSilent = ii
    
    End If
    Else
        SelStartSilent = JJ
        RemoveUndo space(1)
        InsertText = space(1)
        
        SelStartSilent = where + 1
    End If
End If
glistN.enabled = True
    ReColorBlock
Render
nochange = False
End If
KeyCode = 0
shift = 0

End If

Case Else
alt = False
ctrl = False
End Select
noentrance = False
End Sub
Sub SearchDown(s$, Optional anystr As Boolean = False)
Dim l As Long, W As Long, TempLcid As Long, OldLcid As Long
W = mDoc.MarkParagraphID   ' this is the not the order
SelStartSilent = SelStart
l = Charpos

OldLcid = mDoc.LCID
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then mDoc.LCID = TempLcid
If EditTextWord Or anystr Then
    If anystr Then
  If Not mDoc.FindStrDown(s$, W, l) Then GoTo sdnOut
  Else
    If Not mDoc.FindWord(s$, True, W, l) Then GoTo sdnOut
    End If
Else
    If WordOnly Then
    If Not mDoc.FindWord(s$, True, W, l) Then GoTo sdnOut
    Else
    If Not mDoc.FindIdentifier(s$, True, W, l) Then GoTo sdnOut
    End If
End If
SelLengthSilent = 0
mDoc.MarkParagraphID = W
glistN.enabled = False
ParaSelStart = l
glistN.enabled = True
SelLength = Len(s$)
sdnOut:
mDoc.LCID = OldLcid
End Sub
Sub Searchup(s$, Optional anystr As Boolean = False)
Dim l As Long, W As Long, TempLcid As Long, OldLcid As Long
W = mDoc.MarkParagraphID
SelStartSilent = SelStart - (SelLength > 1)
l = Charpos + Len(s$)
OldLcid = mDoc.LCID
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then mDoc.LCID = TempLcid
If EditTextWord Or anystr Then
   If anystr Then
   If Not mDoc.FindStrUp(s$, W, l) Then GoTo sdupOut
   Else
       If Not mDoc.FindWord(s$, False, W, l) Then GoTo sdupOut
    End If
Else
    If WordOnly Then
        If Not mDoc.FindWord(s$, False, W, l) Then GoTo sdupOut
    Else
        If Not mDoc.FindIdentifier(s$, False, W, l) Then GoTo sdupOut
    End If
End If
SelLengthSilent = 0
mDoc.MarkParagraphID = W
glistN.enabled = False
ParaSelStart = l
glistN.enabled = True
SelLength = Len(s$)
sdupOut:
mDoc.LCID = OldLcid
End Sub
Public Sub helpmeSub()
      If mIndex >= 0 Then
       Callback.Callback mesCtrlName + ".Help(" + str(mIndex) + ",{" + Trim$(SelText) + "})"
       Else
          Callback.Callback mesCtrlName + ".Help({" + Trim$(SelText) + "})"
          End If
End Sub
Public Sub supsub()
Dim b$
b$ = s$
s$ = SelText
If s$ = vbNullString Or InStr(s$, Chr$(13)) > 0 Or InStr(s$, Chr$(10)) > 0 Then s$ = b$
Searchup s$
End Sub
Public Sub sdnSub()
Dim b$
b$ = s$
s$ = SelText
If s$ = vbNullString Or InStr(s$, Chr$(13)) > 0 Or InStr(s$, Chr$(10)) > 0 Then s$ = b$
SearchDown s$
End Sub
Public Sub mscatsub()
''
Dim l As Long, W As Long, s$, TempLcid As Long, OldLcid As Long
Dim el As Long, eW As Long, safety As Long, TT$

W = mDoc.MarkParagraphID
eW = W
SelStartSilent = SelStart  'MOVE CHARPOS TO SELSTART

el = Charpos  ' charpos maybe is in the start or the end of block
s$ = SelText
OldLcid = mDoc.LCID
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then mDoc.LCID = TempLcid

l = el + 1
If EditTextWord Then
Do
If mDoc.FindWord(s$, True, W, l) Then
TT$ = mDoc.TextParagraph(W)
Mid$(TT$, l, Len(s$)) = s$
mDoc.ReWritePara W, TT$
mDoc.WrapAgainBlock W, W
mDoc.ColorThis (W)
Else
W = 1
l = 0
safety = safety + 1
End If
Loop Until (W = eW And l = el) Or safety = 2

Else
Do
If mDoc.FindIdentifier(s$, True, W, l) Then
TT$ = mDoc.TextParagraph(W)
Mid$(TT$, l, Len(s$)) = s$
mDoc.TextParagraph(W) = TT$
mDoc.WrapAgainBlock W, W
mDoc.ColorThis (W)

Else
W = 1
l = 0
safety = safety + 1
End If
Loop Until (W = eW And l = el) Or safety = 2

End If
mDoc.LCID = OldLcid
Render
If glistN.Visible Then glistN.SetFocus
End Sub

Public Sub rthissub(Optional shift As Integer)
Dim l As Long, W As Long, s$, TempLcid As Long, OldLcid As Long
Dim el As Long, eW As Long, safety As Long, TT$, w1 As Long, i1 As Long, noinp As Double
Dim neo$, mDoc10 As Document, addthat As Long, w2 As Long, ok1 As Boolean
Dim anystr As Boolean
Dim prof1 As New clsProfiler
W = mDoc.MarkParagraphID
eW = W
w2 = -1
SelStartSilent = SelStart  'MOVE CHARPOS TO SELSTART
el = Charpos  ' charpos maybe is in the start or the end of block
s$ = Trim$(SelText)
SelStartSilent = SelStart
el = Charpos  ' charpos maybe is in the start or the end of block
If shift Mod 2 = 1 Then
shift = 0
anystr = True
If pagio$ = "GREEK" Then
neo$ = InputBoxN("Αλλαγή Χαρακτήρων", "Διορθωτής Κειμένου", s$, noinp)
Else
neo$ = InputBoxN("Replace Characters", "Text Editor", s$, noinp)
End If
ElseIf WordOnly Or Not EditTextWord Then
If pagio$ = "GREEK" Then
neo$ = InputBoxN("Αλλαγή Λέξης", "Διορθωτής Κειμένου", s$, noinp)
Else
neo$ = InputBoxN("Replace Word", "Text Editor", s$, noinp)
End If
Else
If EditTextWord Then
    If pagio$ = "GREEK" Then
        neo$ = InputBoxN("Αλλαγή " & IIf(anystr, "μέρους ", "") & "Λέξης", "Συγγραφή " & IIf(EditTextWord, "Κειμένου", "Κώδικα"), s$, noinp)
    Else
        neo$ = InputBoxN("Replace " & IIf(anystr, "part of ", "") & "Word", IIf(EditTextWord, "Text", "Code") & " Editor", s$, noinp)
    End If
End If

End If

If noinp <> 1 Then Exit Sub
OldLcid = mDoc.LCID
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then mDoc.LCID = TempLcid
If Len(neo$) >= Len(s$) Then
    Set mDoc10 = New Document
    
    mDoc10 = neo$
    mDoc10.NoIgnoreCase = Not mUseCase
    w1 = 0
    i1 = 0
    
    If EditTextWord Or WordOnly Then
        If mDoc10.FindWord(s$, True, w1, i1) Then addthat = i1 - 1: If Len(neo$) = Len(s$) And addthat = 0 Then Exit Sub
    ElseIf anystr Then
        If mDoc10.FindStrDown(s$, w1, i1) Then addthat = i1 - 1: If Len(neo$) = Len(s$) And addthat = 0 Then Exit Sub
    Else
        If mDoc10.FindIdentifier(s$, True, w1, i1) Then addthat = i1 - 1: If Len(neo$) = Len(s$) And addthat = 0 Then Exit Sub
    End If
    
End If
prof1.MARKONE
glistN.SuspDraw = True
i1 = el
l = i1 + addthat
w1 = W
If EditTextWord Or WordOnly Then
glistN.DropKey = True
Do

If anystr Then
ok1 = mDoc.FindStrDown(s$, W, l)
Else
ok1 = mDoc.FindWord(s$, True, W, l)
End If
If ok1 Then
If w2 > 0 Then If w2 <> W Then mDoc.WrapAgainBlock w2, w2:  mDoc.ColorThis w2
w2 = W
If safety And W = w1 Then

If l = i1 Then
 SelLengthSilent = 0
mDoc.MarkParagraphID = W
glistN.enabled = False
ParaSelStart = l
glistN.enabled = True
SelLength = Len(s$)
AddUndo ""
SelText = neo$
RemoveUndo neo$
Exit Do
ElseIf l - addthat < i1 Then
i1 = i1 + Len(neo$) - Len(s$)
Else

End If
End If
SelLengthSilent = 0
mDoc.MarkParagraphID = W
glistN.enabled = False
ParaSelStart = l
glistN.enabled = True
SelLength = Len(s$)
AddUndo ""
SelText = neo$

RemoveUndo neo$
undoprev = True
'l = l + Len(neo$)

Else
W = 1
l = 0
safety = safety + 1
End If
If prof1.MARKTWO > 1000 Then ProcTask2 Basestack1: prof1.MARKTWO
Loop Until safety = 2 Or KeyPressed(16)
glistN.DropKey = False

Else

glistN.DropKey = True
Do
If mDoc.FindIdentifier(s$, True, W, l) Then
If w2 > 0 Then If w2 <> W Then mDoc.WrapAgainBlock w2, w2:  mDoc.ColorThis w2
w2 = W
If safety And W = w1 Then
If l = i1 Then
 SelLengthSilent = 0
mDoc.MarkParagraphID = W
glistN.enabled = False
ParaSelStart = l
glistN.enabled = True
SelLength = Len(s$)
AddUndo ""
SelText = neo$
RemoveUndo neo$
Exit Do
ElseIf l - addthat < i1 Then
i1 = i1 + Len(neo$) - Len(s$)
Else

End If
End If
SelLengthSilent = 0
mDoc.MarkParagraphID = W
glistN.enabled = False
ParaSelStart = l
glistN.enabled = True
SelLength = Len(s$)
AddUndo ""
SelText = neo$
RemoveUndo neo$
undoprev = True
l = l + Len(neo$)

Else
W = 1
l = 0
safety = safety + 1
End If
If prof1.MARKTWO > 1000 Then ProcTask2 Basestack1: prof1.MARKTWO
Loop Until safety = 2 Or KeyPressed(16)
glistN.DropKey = False
End If
mDoc.LCID = OldLcid
If w2 > 0 Then mDoc.WrapAgainBlock w2, w2:   mDoc.ColorThis w2
glistN.SuspDraw = False
Render
If glistN.Visible Then glistN.SetFocus
End Sub
Private Sub MarkSoftButton(para As Long, pospara As Long)
If glistN.lines = 1 Then Exit Sub
If ShadowMarks Then Exit Sub
If para = 0 Then 'set
    para = mDoc.MarkParagraphID
    pospara = ParaSelStart
    
    If UserCodePage = 1253 Then
        ReplaceTitle = "Ο δείκτης τώρα θα δείχνει αυτή την παράγραφο"
    Else
        ReplaceTitle = "Mark now move to this Paragraph and Position"
    End If
ElseIf para = mDoc.MarkParagraphID And pospara = Charpos Then 'Reset
    para = 0
    
    If UserCodePage = 1253 Then
    ReplaceTitle = "Διαγραφή Δείκτη"
    Else
    ReplaceTitle = "Mark Deleted"
    End If
Else ' goto that paragraph
    If Not mDoc.InvalidPara(para) Then
        SelLengthSilent = 0
        mDoc.MarkParagraphID = para
        glistN.enabled = False
        ParaSelStart = pospara
        glistN.enabled = True
        ManualInform
    Else
        para = 0
        If UserCodePage = 1253 Then
            ReplaceTitle = "Δεν βρέθηκε παράγραφος - διαγράφτηκε ο δείκτης"
        Else
            ReplaceTitle = "Paragraph not found - mark deleted"
        End If
    End If
End If
End Sub
Private Sub SetME()
'If (600 - hueconv(backcolor)) Mod 360 > 30 And lightconv(backcolor) >= 128 Then ColorSet = 1 Else ColorSet = 0
Select Case fState
Case 0
'shortlang = False
If M2000code Then
    NoColor = False
Else
    NoColor = EditTextWord
End If
Case 1
'shortlang = False
NoColor = True
Case 2
'shortlang = True
If M2000code Then
    NoColor = False
Else
    NoColor = EditTextWord
End If
Case 3
'shortlang = True
NoColor = True
fState = -1
End Select
mDoc.ColorEvent = Not NoColor
End Sub
Private Function IsmineHere(a$) As Boolean
If mUseCase Then
IsmineHere = InStr(mColorCollection1, "|" + UCase(a$) + "|") > 0
Else
IsmineHere = InStr(mColorCollection1, "|" + a$ + "|") > 0
End If
End Function
Private Function IsmineHere1(a$) As Boolean
If mUseCase Then
IsmineHere1 = InStr(ColorCollection2, "|" + UCase(a$) + "|") > 0
Else
IsmineHere1 = InStr(ColorCollection2, "|" + a$ + "|") > 0
End If
End Function
Private Function IsmineHere2(a$) As Boolean
If mUseCase Then
IsmineHere2 = InStr(ColorCollection3, "|" + UCase(a$) + "|") > 0
Else
IsmineHere2 = InStr(ColorCollection3, "|" + a$ + "|") > 0
End If
End Function
Private Function IsmineHere3(a$) As Boolean
If mUseCase Then
IsmineHere3 = InStr(ColorCollection4, "|" + UCase(a$) + "|") > 0
Else
IsmineHere3 = InStr(ColorCollection4, "|" + a$ + "|") > 0
End If
End Function


Public Property Let MultiLineComment2(ByVal vNewValue As Variant)
mMultiLineComment2 = vNewValue
mDoc.Final = vNewValue
End Property
Public Sub PopUpMenu(poptitle$, X As Variant, Y As Variant)
Dim oldModalId As Variant, i As Long
''If UsedAsTextBox Then Exit Sub
i = SelLength
Form1mn1Enabled = i > 0
Form1mn2Enabled = i > 0
Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
Form1sdnEnabled = i > 0 And (Length - SelStart) > SelLength
Form1supEnabled = i > 0 And SelStart > SelLength
Form1mscatEnabled = Form1sdnEnabled Or Form1supEnabled
Form1rthisEnabled = Form1mscatEnabled

oldModalId = Modalid
Modalid = Rnd * 113314324
'UnHook Callback.hWnd
UNhookMe
MyPopUp.feedlabels Me, False
MyPopUp.UpGui Callback, X, Y, poptitle$
MyPopUp.Visible = True
MyPopUp.ZOrder 0

        Do While MyPopUp.Visible
   ProcTask2 Basestack1
   If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow <> Screen.ActiveForm.hWnd Then MyPopUp.Visible = False
End If
                '     mywait basestack1, 1
                If ExTarget Then Exit Do
                Loop
                Modalid = oldModalId
         '  Hook Callback.hWnd, Nothing
End Sub


Public Sub mn1sub()
MarkCut
End Sub

Public Sub mn2sub()
MarkCopy
End Sub

Public Sub mn3sub()
On Error Resume Next
Dim Pad$
Sleep 10
    On Error Resume Next
    Pad$ = GetTextData(13)
    If Err Or Pad$ = "" Then
    Pad$ = Clipboard.GetText(1)
    If Err Or Pad$ = "" Then Err.Clear: Exit Sub
    End If
If NumberOnly Or NumberIntOnly Then
    If Not ValidNum(Pad$, False, NumberIntOnly) Then Exit Sub
    NumberOnly = True
    If Not mNoWrap = True Then
        mNoWrap = True
        glistN.NoScroll = False
    End If
ElseIf UsedAsTextBox Then
    Pad$ = Replace(Pad$, Chr(13), "")
    Pad$ = Replace(Pad$, Chr(10), "")
    If Not mNoWrap = True Then
        mNoWrap = True
        glistN.NoScroll = False
    End If
End If
If ParaSelStart = 2 And glistN.list(glistN.ListIndex) = vbNullString Then
SelStart = SelStart - 1
End If
AddUndo ""
SelText = Pad$
RemoveUndo Pad$
ReColorBlock
End Sub


Public Sub wordwrapsub()
nowrap = Not nowrap
Render
ManualInform
End Sub
Public Sub Resize()
If Not nowrap Then WrapAll
Render
ManualInform
End Sub
Private Sub inform(tLine As Long, tPos As Long)
If UsedAsTextBox Then

Else
If UserCodePage = 1253 Then
textinform = "Γραμμή(" + CStr(tLine) + ")-Θέση(" + CStr(Charpos) + ")"

Else
textinform = "Line(" + CStr(tLine) + ")-Pos(" + CStr(Charpos) + ")"
End If

End If
End Sub

Public Property Get ColorCollection1() As Variant
ColorCollection1 = mColorCollection1
End Property

Public Property Let ColorCollection1(ByVal vNewValue As Variant)
mColorCollection1 = vNewValue
EditTextWord = vNewValue = vbNullString
SetME
End Property

Public Property Get NoColor() As Variant
NoColor = mNoColor
End Property

Public Property Let NoColor(ByVal vNewValue As Variant)
mNoColor = vNewValue
mNoColor1 = vNewValue
End Property
Public Property Let ShowAlways(ByVal vNewValue As Variant)
mStayOpen = CBool(vNewValue)
If mStayOpen Then
glistN.ForeColor = black 'rgb(0, 0, 0)
glistN.BackColor = white ' rgb(255, 255, 255)
End If
End Property

Sub GetFocus()
On Error Resume Next
Callback.Controls(ctrlName).SetFocus
End Sub
Sub ReColor()
On Error GoTo end1
Set Callback.mDoc = mDoc
WrapAll
If mDoc Is Nothing Then Set Callback.mDoc = Nothing: Exit Sub
Set Callback.mDoc = Nothing
ManualInform
end1:
End Sub
Private Sub ReColorBlock()
Dim order1, order2
order1 = 0: order2 = 0
If BlockEndPara <> BlockStartPara Then
order1 = mDoc.ParagraphOrder(BlockStartPara)
order2 = mDoc.ParagraphOrder(BlockEndPara)
End If
If order1 <= order2 Then
mDoc.WrapAgainBlock BlockStartPara, BlockEndPara
mDoc.WrapAgainColorBlock2 BlockStartPara, BlockEndPara
Else
mDoc.WrapAgainBlock BlockEndPara, BlockStartPara
mDoc.WrapAgainColorBlock2 BlockEndPara, BlockStartPara

End If
ManualInform
End Sub
Private Sub ColorBlockOnly()
Dim order1, order2
order1 = 0: order2 = 0
If BlockEndPara <> BlockStartPara Then
order1 = mDoc.ParagraphOrder(BlockStartPara)
order2 = mDoc.ParagraphOrder(BlockEndPara)
End If
If order1 <= order2 Then
mDoc.WrapAgainColorBlock2 BlockStartPara, BlockEndPara
Else
mDoc.WrapAgainColorBlock2 BlockEndPara, BlockStartPara
End If
''ManualInform
End Sub
Sub TextViewOnly()
mStayOpen = False
NoMark = True
mNoColor = True
mNoColor1 = True
EditDoc = False
glistN.PanPos = 0
nowrap = False
mDoc1.WrapAgain
glistN.ScrollToTextEdit 0
Me.SetRowColumn 1, 1, True
End Sub
Sub EditTextView(ByVal OnOff As Boolean)
mStayOpen = True
glistN.ForeColor = black 'rgb(0, 0, 0)
glistN.BackColor = white ' rgb(255, 255, 255)
NoMark = Not OnOff
mEditDoc = OnOff
If Not OnOff Then glistN.MarkNext = 0 Else glistN.MarkNext = 2
glistN.EditFlag = OnOff
glistN.NoCaretShow = NoMark
End Sub
Sub Transparent()
Dim X  As Long, Y As Long
glistN.BackStyle = 1
glistN.GetLeftTop X, Y
glistN.RepaintFromOut Callback.Image, X, Y
glistN.ShowMe
Callback.RefreshList = Callback.RefreshList + 1
End Sub
Sub TransparentOff()
On Error Resume Next
If Not glistN Is Nothing Then
glistN.BackStyle = 0
glistN.BackColor = mycolor(dimm2)
glistN.PanPos = 0
glistN.ShowMe
Callback.RefreshList = Callback.RefreshList - 1
End If
End Sub
Sub ColorsReset()
white = rgb(255, 255, 255)
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)

With glistN
.dcolor = Callback.Controls(1).dcolor
.CapColor = capchroma
End With
Retired

End Sub
Sub Colors(white1, Optional black1, Optional dimm11, Optional dimm21, Optional dimm33, Optional dimm34)
If Not glistN Is Nothing Then
white = mycolor(white1)
If Not IsMissing(black1) Then black = mycolor(black1)
If Not IsMissing(dimm11) Then dimm1 = mycolor(dimm11)
If Not IsMissing(dimm21) Then dimm2 = mycolor(dimm21)
If Not IsMissing(dimm34) Then glistN.dcolor = CLng(mycolor(dimm34))
If Not IsMissing(dimm33) Then glistN.CapColor = CLng(mycolor(dimm33))
Retired
End If
End Sub
Public Sub Retired()
If mStayOpen Then
glistN.ForeColor = black
glistN.BackColor = white
glistN.NoCaretShow = False
Else
glistN.ForeColor = white
glistN.BackColor = dimm2
glistN.NoCaretShow = True
End If
 glistN.PanPos = 0: glistN.ShowMe
End Sub
Public Property Let DisplayLines(ByVal vNewValue As Variant)
glistN.addpixels = 0
glistN.restrictLines = CLng(vNewValue)
glistN.CalcNewFont
TextPixelOffset = Int(glistN.TextHeightOffset)
End Property

Private Sub mDoc_MayQuit(Yes As Variant)
If Quit Then Yes = True
'MyDoEvents1 Form3
End Sub


Public Property Get Value() As Variant
If Not (NumberOnly Or NumberIntOnly) Then Exit Property
NumberOnly = True
If mDoc.SizeCRLF = 0 Then
Value = 0
Else
If Not NumberIntOnly Then
If NowDec$ <> "." Then
Value = val(Replace$(Text, NowDec$, "."))
Else
Value = val(Replace$(Text, NowDec$, "."))
End If
Else
Value = val(Text)
End If
End If

End Property

Public Property Let Value(ByVal vNewValue As Variant)
NumberOnly = True
Dim content As String

If MyIsNumeric(vNewValue) Then
content = NLtrim$(str(vNewValue))
If Not ValidNum(content, True, NumberIntOnly) Then
    Text = vbNullString
ElseIf vNewValue = 0 Then
    Text = vbNullString
Else
    Text = content
    End If
    Else
    Text = vbNullString
    End If
End Property
Public Property Let Default(RHS)
On Error Resume Next
If CBool(RHS) Then Callback.Default = ctrlName
End Property


Public Property Let WordCharLeft(ByVal vNewValue As String)
If Not glistN Is Nothing Then
glistN.WordCharLeft = vNewValue
End If
End Property

Public Property Let WordCharRight(ByVal vNewValue As String)
If Not glistN Is Nothing Then
glistN.WordCharRight = vNewValue
End If
End Property
Public Property Let WordCharLeftButIncluded(ByVal vNewValue As String)
If Not glistN Is Nothing Then
glistN.WordCharLeftButIncluded = vNewValue
End If

End Property
Public Property Let WordCharRightButIncluded(ByVal vNewValue As String)
If Not glistN Is Nothing Then
glistN.WordCharRightButIncluded = vNewValue
End If

End Property
Public Property Let icon(RHS)
glistN.icon = RHS

End Property
Public Property Let Drag(RHS)
glistN.DragEnabled = RHS
End Property
Public Property Get Drag()
Drag = glistN.DragEnabled
End Property
Public Property Let drop(RHS)
glistN.DropEnabled = RHS
End Property
Public Property Get drop()
drop = glistN.DropEnabled
End Property
Public Property Get UseTab() As Boolean
UseTab = glistN.UseTab
End Property

Public Property Let UseTab(ByVal vNewValue As Boolean)
glistN.UseTab = vNewValue
End Property
Public Property Let TabWidthChar(ByVal vNewValue As Long)
glistN.TabWidthChar = vNewValue
tParam.iTabLength = Abs(vNewValue)
End Property
Private Function MyTrimL(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL = 1: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 9
    Case Else
     MyTrimL = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimL = l + 2
End Function
Private Function MyTrimLNoTab(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimLNoTab = 1: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     MyTrimLNoTab = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimLNoTab = l + 2
End Function
Private Function MyTrimL2(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL2 = 0: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     MyTrimL2 = (i - p2) \ 2 '+ 1
   Exit Function
  End Select
  Next i
 MyTrimL2 = 0
End Function
Private Function NLTrim2(s$) As String
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then NLTrim2 = 1: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     NLTrim2 = Mid$(s$, (i - p2) \ 2 + 1)
   Exit Function
  End Select
  Next i
 NLTrim2 = vbNullString
End Function
Private Function NLtrim$(a$)
If Len(a$) > 0 Then NLtrim$ = Mid$(a$, MyTrimL(a$))
End Function
Private Sub countbrakcet(a$, K As Long, counter As Long)
Dim i As Long, j As Integer, blind As Boolean
For i = Abs(K) To Len(a$)
j = AscW(Mid$(a$, i, 1))
Select Case j
Case 34
blind = Not blind
Case 123
If Not blind Then counter = counter + 1
Case 125
If Not blind Then counter = counter - 1: If counter = 0 Then K = i: Exit For
End Select
Next i

End Sub
Function blockStringAhead2(s$, pos1 As Long, j As Long) As Boolean
Dim i As Long, c As Long
c = Len(s$)
i = pos1
If i > c Then blockStringAhead2 = False: Exit Function
Do

Select Case AscW(Mid$(s$, i, 1))
Case 34
Do While i < c
i = i + 1
If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
Loop
Case 123
j = j - 1
Case 125
j = j + 1
End Select
i = i + 1

Loop Until i > c Or j = 0
pos1 = i - 1
blockStringAhead2 = True


End Function

Function blockStringAhead(s$, pos1 As Long) As Long
Dim i As Long, j As Long, c As Long
c = Len(s$)
i = pos1
If i > c Then blockStringAhead = False: Exit Function
Do

Select Case AscW(Mid$(s$, i, 1))
Case 34
Do While i < c
i = i + 1
If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
Loop
Case 123
j = j - 1
Case 125
j = j + 1
End Select
i = i + 1

Loop Until i > c Or j = 0
pos1 = i - 1
blockStringAhead = True


End Function
Public Sub UNhookMe()
On Error Resume Next
Set LastGlist = Nothing
UnHook Callback.hWnd
End Sub
Function ISSTRINGb(bb$, r$) As Boolean
'
Dim q$, W As Long, a$
a$ = NLtrim$(bb$)
r$ = vbNullString
If a$ = vbNullString Then Exit Function
Select Case Myleft$(a$, 1)
Case StringSep
q$ = StringSep
Case Myleft$(stringsep2, 1)
q$ = Right$(stringsep2, 1)
End Select
If q$ = vbNullString Or Len(a$) < 2 Then ISSTRINGb = False: Exit Function
If q$ = Chr(34) Then
    W = 1
    Do
    W = W + 1
    W = InStr(W, a$, q$)
    If W > 0 Then
    If Mid$(a$, W - 1, 1) <> "\" Then Exit Do
    ElseIf W = 0 Then
    Exit Do
    End If
    Loop

Else
W = InStr(2, a$, q$, vbBinaryCompare)
End If
If W = 0 Then ISSTRINGb = False: Exit Function
r$ = Mid$(a$, 2, W - 2)
bb$ = NLtrim$(Mid$(a$, W + 1))

ISSTRINGb = True

End Function


Public Property Let UseCase(ByVal RHS As Boolean)
mUseCase = Not RHS
mDoc.NoIgnoreCase = RHS
End Property


Public Property Let SearchWords(ByVal RHS As Boolean)
WordOnly = RHS
Brackets = RHS
End Property
Public Property Let BracketsForArrays(ByVal RHS As Boolean)
Brackets = RHS
End Property
Function BlockParam2(s$, Pos As Long, j As Long, j1 As Long) As Boolean
' need to be open
Dim i As Long, ii As Long, exit1 As Boolean
exit1 = j <> 0
For i = Pos To Len(s$)
Select Case AscW(Mid$(s$, i, 1))
Case 0
Exit For
Case 34
again:
ii = InStr(i + 1, s$, """")
If ii = 0 Then Exit Function
 i = ii
'If Mid$(s$, ii - 1, 1) = "`" Then GoTo again
Case 91
j1 = j1 + 1
Case 93
j1 = j1 - 1
If Not exit1 Then If j1 = 0 Then Exit For
If j1 < 0 Then Exit For
Case 40
j = j + 1
Case 41
j = j - 1
If exit1 Then If j = 0 Then Exit For
If j < 0 Then Exit For
Case 123
i = i + 1
If blockStringPOS(s$, i) Then
Else
i = 0
End If
If i = 0 Then Exit Function
End Select
Next i
If j = 0 And j1 = 0 Then Pos = i: BlockParam2 = True
End Function

Function BlockParam4(s$, Pos As Long, what As Long) As Boolean
' need to be open
Dim i As Long, j As Long, ii As Long
If what = 91 Then
j = 1000
Else
j = 1
End If
For i = Pos To Len(s$)
Select Case AscW(Mid$(s$, i, 1))
Case 0
Exit For
Case 34
again:
ii = InStr(i + 1, s$, """")

If ii = 0 Then Exit Function

 i = ii
If Mid$(s$, ii - 1, 1) = Chr$(34) Then GoTo again
If Mid$(s$, ii - 1, 1) = "\" Then GoTo again
Case 91
If WordOnly Then
j = j + 1000
End If
Case 93
If WordOnly Then
j = j - 1000
If j = 0 Then Exit For
End If
Case 40
j = j + 1
Case 41
j = j - 1
If j = 0 Then Exit For
Case 123
i = i + 1
If blockStringPOS(s$, i) Then
Else
i = 0
End If
If i = 0 Then Exit Function
End Select
Next i
If j = 0 Then Pos = i: BlockParam4 = True
End Function
Public Property Let SplitExpr(ByVal RHS As String)
If RHS = "" Then
mSplitExpr = ","
Else
mSplitExpr = Myleft$(RHS, 1)
End If
End Property

Public Property Let UseColon(ByVal RHS As String)
If RHS = "" Then
LineCommandSep = Chr$(0)
Else
LineCommandSep = Myleft$(RHS, 1)
End If
End Property
Private Function MyTrimRNoCr(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimRNoCr = 1: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p4 To p2 Step -2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 9
    Case Else
     MyTrimRNoCr = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimRNoCr = l + 2
End Function
Private Function MyTrimL2NoTab(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL2NoTab = 0: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     MyTrimL2NoTab = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimL2NoTab = 0
End Function
Sub Keyboard(RHS As String)
    RemoveUndo RHS
    InsertTextNoRender = RHS
    Render
End Sub
Private Function IsRTLrun() As Boolean
Dim what$
mark1 = 0
mark2 = 0
what$ = mDoc.TextParagraph(mDoc.MarkParagraphID)
Dim A2() As Integer, WHAT1$, r As Long, ii As Long
Dim wr$
Dim mcharpos As Long
mcharpos = Charpos
If Len(what$) = 0 Then GoTo ex1
If mcharpos > Len(what$) Then mcharpos = Len(what$)
ReDim A2(Len(what$) + 10)
If GetStringTypeExW(&HB, 2, StrPtr(what$), Len(what$), A2(0)) = 0 Then Exit Function

'If A2(mcharpos) = 1 Then goto ex1
For ii = mcharpos To 0 Step -1
If ii < Len(what$) Then
If A2(ii) = 2 Then
mark1 = ii + 1
End If

If mark1 > 0 And A2(ii) = 1 Then Exit For
End If
Next ii
If mark1 = 0 Then GoTo ex1
For ii = mark1 - 1 To Len(what$) - 1
If A2(ii) = 1 Then Exit For
If ii >= 0 Then
If A2(ii) = 7 Then
For r = ii + 1 To Len(what$) - 1
If A2(r) = 1 Then Exit For
If A2(r) <> 7 And A2(r) < 10 Then
If (A2(r) And 2) = 2 Then
ii = r
r = -1
Exit For
End If

If mark2 > mcharpos Then: r = 0: Exit For
End If
Next r
If r = 0 Then
Exit For
ElseIf r <> -1 Then
ii = r - 1
End If
End If
If (A2(ii) And 2) = 2 And A2(ii) <> 10 Then
mark2 = ii + 1
If mark2 > 0 And A2(ii) = 1 Then Exit For
End If
End If
Next ii
If mark2 < mcharpos Then
GoTo ex1
End If
If mark1 <> 0 And mark2 <> 0 Then
IsRTLrun = True

End If
ex1:


    
End Function
Private Sub glistN_RTL(thisHDC As Long, item As Long, where As Long, mark10 As Long, mark20 As Long, offset As Long)
offset = RealX1(thisHDC, where, mDoc.TextLine2(item + 1, showparagraph, NoColor), where, mark10)

End Sub
Function RealX1(mHdc As Long, ByVal ExtSelStart As Long, ByVal that$, Optional nValue As Long = 0, Optional found As Long = 0) As Long
If ExtSelStart <= 0 Then ExtSelStart = 1
Dim ww$(), Col() As Long, items&, i, c As String
Dim rleftacc As Long, original As Long, rx1 As Long, direction As Long
Dim OldValue As Long, orx1 As Long, j As Long, nrtl As Boolean
If Len(that$) = 0 Then
    nValue = 1: Exit Function
End If
If showparagraph Then
        that$ = Replace(that$, vbTab, ChrW(&H21E5))
End If
Col() = mDoc.lastcolordata()
original = ExtSelStart
If UBound(Col()) < 1 Then
GoTo there
ElseIf UBound(Col()) = 2 Then
    If Col(0) = 0 And Col(1) = 0 And Col(2) = 0 Then GoTo there
    ww$() = SplitStringUsingArrayNoTab(that$, Col(), items&)
ElseIf UBound(Col()) = 1 And (Col(0) = 14 Or Col(0) = 0 Or Col(0) = 7) Then
there:
ReDim ww$(0)
ww$(0) = that$
Else
ww$() = SplitStringUsingArrayNoTab(that$, Col(), items&)
End If


Dim M1 As Long, M2 As Long
Dim dummy As RECT, r As RECT, tabw As Long, offset As Long, hRgn As Long, estX As Long
tabw = tParam.iTabLength * glistN.AveCharWith
glistN.UserControlTextMetricsPixels that$, r.Right, r.Bottom
dummy = r
dummy.top = r.Bottom + 1
dummy.Bottom = dummy.top + r.Bottom
offset = 0
estX = 0
hRgn = CreateRectRgnIndirect(dummy)
SelectClipRgn mHdc, hRgn
If found > 0 Then
rx1 = LowWord(GetTabbedTextExtent(mHdc, StrPtr(that$), Len(that$), 1, tabw))
If rx1 < found Then

c = Right$(that$, 1)

nValue = 1
RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
If mark1 <> 0 And mark2 <> 0 Then
nValue = Len(that$)
RTL3 that$, nValue, mark1, mark2, nrtl, M1, M2
nValue = mark1
GoTo finish
Else
nValue = Len(ww(i))
RTL3 ww(i), nValue, mark1, mark2, nrtl, M1, M2
If mark1 <> 0 And mark2 <> 0 And nValue >= mark2 Then
If nValue <= mark2 + 1 Then
nValue = mark1
GoTo finish
Else
nValue = nValue + 1
ExtSelStart = nValue
End If
Else
nValue = Len(that$) + 1
GoTo finish
End If

End If
End If
End If
For i = 0 To items&
    c = ww$(i)
    r.Left = offset
    rleftacc = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, 0))
    If Len(c) >= ExtSelStart - estX Or (found > 0 And rleftacc > found) Then
            If Not Len(c) >= ExtSelStart - estX Then
            ExtSelStart = Len(c) + estX
            End If

        nValue = ExtSelStart - estX

        RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
        If nrtl Then
            If M2 - M1 > 0 Then Mid$(c, M1, M2 - M1 + 1) = StrReverse(Mid$(c, M1, M2 - M1 + 1))
        End If
        If mark1 = 0 Or mark2 = 0 Then
        If nValue > 1 Then
            j = nValue - 1
            RTL3 c, j, mark1, mark2, nrtl, M1, M2
            If mark1 <> 0 And mark2 <> 0 And j >= mark1 And j <= mark2 Then
            If i = items& Then
            If found > 0 Then
                rx1 = offset
                If mark1 > 1 Then
                    offset = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), mark1 - 1, 1, tabw, 0))
                End If
                c = Mid$(ww(i), mark1, mark2 - mark1 + 1)
                If nrtl Then
                    M1 = M1 - mark1 + 1
                    M2 = M2 - mark1 + 1
                    If M2 - M1 > 0 Then Mid$(c, M1, M2 - M1 + 1) = StrReverse(Mid$(c, M1, M2 - M1 + 1))
                End If
                If found > offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c$), 1, tabw, 0)) Then
                        nValue = mark1 + estX
                        Exit For
                    Else
                        offset = rx1
                        c = ww$(i)
                    End If
                nValue = j + estX
                End If
             GoTo here1

            Else
            nValue = mark2 - 1 + estX 'j - (nValue > mark2)
            If nValue <= 0 Then nValue = 1
            If mark1 > 1 Then
            
            rx1 = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), mark2, 1, tabw, 0))
            If found > 0 Then
            If rx1 <= found Then
            nValue = mark1 + estX
            RealX1 = rx1
            Exit For
            End If

            End If
            RealX1 = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), mark1 - 1, 1, tabw, 0))
            
            End If
            Exit For
            
            End If
            End If
        End If
        ElseIf mark1 <> 0 And mark2 <> 0 And nValue = mark2 + 1 Then
        If found > 0 Then
         rx1 = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), mark2, 1, tabw, 0))
        If found > rx1 Then
         nValue = mark1 + estX
            RealX1 = rx1
        Exit For
        End If
        nValue = mark2
        End If
        GoTo here1
       
        End If

        If mark1 <> 0 And mark2 <> 0 And nValue >= mark1 And nValue <= mark2 Then
here1:
            If mark1 > 1 Then
                
                offset = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), mark1 - 1, 1, tabw, 0))
            End If
            If mark1 = nValue And found > 0 Then

            If offset > found Then
            j = mark1 - 2
            If j >= 1 Then
            Do While found < r.Left + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), j, 1, tabw, 0))
            j = j - 1
            If j = 0 Then j = 1: Exit Do
            Loop
            Else
                            
                GoTo nortl
            End If
            RTL3 c, j, mark1, mark2, nrtl, M1, M2
            If mark1 <> 0 And mark2 <> 0 Then
            nValue = mark2 - (j - mark1) + estX
            Else
            nValue = j + 1 + estX
            End If

            RealX1 = offset
            Exit For
            End If
            End If
            c = Mid$(ww(i), mark1, mark2 - mark1 + 1)
            If nrtl Then
                M1 = M1 - mark1 + 1
                M2 = M2 - mark1 + 1
                If M2 - M1 > 0 Then Mid$(c, M1, M2 - M1 + 1) = StrReverse(Mid$(c, M1, M2 - M1 + 1))
            End If
           
            If nValue = mark2 + 1 Then
            If found > 0 Then
            For j = mark2 - mark1 + 1 To 1 Step -1
            If found > offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), j, 1, tabw, 0)) Then
            
            If j > 1 Then
            If LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), j, 1, tabw, 0)) > LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), j - 1, 1, tabw, 0)) < 0 Then
            
                        nValue = mark2 - (j) + estX

            offset = found
            Exit For
            End If
            End If

            nValue = mark2 - (j - 1) + estX

            offset = found
            Exit For

            End If
            Next j
            If j = 0 Then
            nValue = mark1 - 1 + estX - 1
            If nValue <= 0 Then nValue = 1
            End If
            Else
            nValue = mark1 - 1 + estX - 1
            If nValue <= 0 Then nValue = 1
            End If
            
          
             RealX1 = offset
         
            Exit For
            ElseIf found > 0 And mark1 = nValue Then
            If found > offset Then
             nValue = mark2
             
            GoTo conthere
            Else
            nValue = mark2 + 1 + estX
            RealX1 = offset
            End If
            
            
            Else
conthere:
            rleftacc = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, 0))
             If nValue > mark1 Then
                rx1 = rleftacc - 1 - LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), nValue - mark1, 1, tabw, 0))
            Else
                rx1 = rleftacc - 1
            End If
            End If
        If found > 0 Then
         rx1 = rx1 + 1
            If found < rx1 Then
                original = nValue
                For j = nValue To mark2

                rx1 = rleftacc - LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), j - mark1 + 1, 1, tabw, 0))

                    
                    original = original + 1
                    If found >= rx1 Then Exit For
                Next j
                
                j = original
                If nrtl Then
                nValue = M2 - ((j + 1 - mark1) - M1) + mark1 + estX
 
                
                
                Else
                
                
                If j > mark2 Then
      
                  nValue = mark2 + 1 + estX
                  If nValue = 0 Then nValue = 1
                      
                          RealX1 = rx1
                          Exit For
                  Else
                  
                          nValue = j
                          RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
                          If nrtl And j >= M1 And j <= M2 Then
                          nValue = nValue + 2 + estX
                          Else
                          nValue = j + estX
                          End If
            
                  
                  End If
              
                
                    RealX1 = rx1 + 1
        End If
                Else
                  original = nValue
                  For j = nValue - 1 To mark1 + 1 Step -1
                
                      rx1 = rleftacc - LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), j - mark1, 1, tabw, 0))
                      If found <= rx1 Then Exit For
                      original = j
                      
                  Next j
                    j = original
                If nrtl Then
                nValue = M2 - ((j + 1 - mark1) - M1) + mark1 + estX
              Else
                    
                    If j < mark1 Then
                       j = mark1
                    End If
                 
                 
                 
                    If nrtl And j - mark1 >= M1 And j - mark1 <= M2 Then
                    nValue = nValue + estX
                    Else
                    nValue = j + estX
                    End If
     End If

                    RealX1 = rx1

            End If
            
          If found = 0 Then
                If Mid$(that$, nValue, 1) = " " Then
         c = Mid$(that$, nValue + 1, 1)
         If Len(c) > 0 Then
         RTL3 c, (1), mark1, mark2, nrtl, M1, M2
         If mark1 = 0 Or mark2 = 0 Then
         nValue = nValue + 1
         'If nvalue1 < 1 Then nvalue1 = 1
         End If
         End If
         End If
         End If
      Exit For
     End If
            
            If found = 0 And nrtl Then
            'nValue = ExtSelStart
            Else
            nValue = mark2 - (nValue - mark1) + (nValue <> mark2) + estX - 1
            End If
            
If nValue = mark1 And items& = 1 Then
rx1 = 0
           
            
            End If
            If nValue = 0 Then nValue = mark2
       
            
            
            RealX1 = rx1
            
        Else ' NO RTL
nortl:
            c = Myleft$(c, ExtSelStart - estX - 1)
            rx1 = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, 0))
            If found > 0 Then
                If found < rx1 Then
                    nValue = ExtSelStart - 1
                    If nValue - estX - 1 < 1 Then
                        RealX1 = offset
                    Else
                        c = Myleft$(c, nValue - estX - 1)
                        j = Len(c)
                        RTL3 c, j, mark1, mark2, nrtl, M1, M2      ' found=0
                        If mark1 <> 0 And mark2 <> 0 Then
                        If j <= mark2 And j >= mark1 Then
                        nValue = mark2 - (j - mark1) + estX
                        Else
                        'nValue = nValue + estX - 1
                        End If
                        Else
                       ' nValue = Len(c) + estX
                        End If
                        rx1 = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, 0))
                    End If
                    RealX1 = rx1
                    Exit For
                Else
                    RealX1 = rx1
                    
                    nValue = original
                End If
            Else
                RealX1 = rx1
                nValue = ExtSelStart
            End If
            
            Exit For
        End If
  
         RealX1 = rx1
       ' nValue = ExtSelStart
        
        Exit For
    ElseIf i = items& Then
        nValue = ExtSelStart - estX
        RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
        If mark1 <> 0 And mark2 <> 0 Then
            If mark2 = Len(c) Then
                c = Myleft$(ww(i), mark1 - 1)
                If Len(c) = 0 Then
                    rleftacc = offset
                ElseIf nrtl And M2 = mark2 Then
                rleftacc = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, 0))
                  c = ww(i)

               If M2 - M1 > 0 Then c = Mid$(c, M1, M2 - M1 + 1)
                   rleftacc = rleftacc + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, 0))
                nValue = M1 - 1
            Else
                 rleftacc = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, 0))
                End If
            ElseIf found > 0 Then
            rleftacc = offset + LowWord(TabbedTextOut(mHdc, r.Left, r.top, StrPtr(c), nValue, 1, tabw, 0))
        If rleftacc < found Then
        nValue = original

        End If
                RealX1 = found
        GoTo finish
            End If
        End If
    End If
    offset = rleftacc
    estX = estX + Len(ww(i))
Next i


If i = items& + 1 Then
    RealX1 = offset
    nValue = estX
End If
finish:
SelectClipRgn mHdc, &H0
DeleteObject hRgn
End Function

Private Sub RTL3(what As String, where As Long, mark10 As Long, mark20 As Long, numhandle As Boolean, M1 As Long, M2 As Long)
numhandle = False
M1 = 0
M2 = 0
mark10 = 0
mark20 = 0
Dim A2() As Integer, r As Long, ii As Long
Dim wr$
Dim mcharpos As Long
If Len(what) = 0 Then Exit Sub

If where = -1 Then mcharpos = Len(what) Else mcharpos = where - 1
If mcharpos > Len(what) Then Exit Sub
ReDim A2(Len(what) + 10)


If GetStringTypeExW(&HB, 2, StrPtr(what), Len(what), A2(0)) = 0 Then Exit Sub
For ii = mcharpos To 0 Step -1
    If ii < Len(what) Then
    If A2(ii) = 9 Then
   If mark10 > 0 Then Exit For
    Exit Sub
    End If
    If A2(ii) = 2 Then
    mark10 = ii + 1
    End If
    If mark10 > 0 And A2(ii) = 1 Then
    Exit For
    End If
    End If
Next ii

If mark10 = 0 Then Exit Sub
For ii = mark10 - 1 To Len(what) - 1

    If A2(ii) = 1 Then
    Exit For
    End If
    If ii >= 0 Then
        If A2(ii) = 7 Then
            For r = ii + 1 To Len(what) - 1
                If A2(r) = 1 Then Exit For
                If A2(r) = 9 Then
                Exit For
                End If
                If A2(r) <> 7 And A2(r) < 11 Then
                    If (A2(r) And 2) = 2 Then
                        ii = r
                        r = -1
                        Exit For
                    End If
                   If mark20 > mcharpos Then: r = 0: Exit For
                End If
            Next r
            If r = 0 Then
                Exit For
            ElseIf r <> -1 Then
                ii = r - 1
            End If
        End If
        If (A2(ii) And 2) = 2 And A2(ii) < 11 And A2(ii) <> 10 Then
            mark20 = ii + 1
            If mark20 > 0 And A2(ii) = 1 Then
            Exit For
            End If
        End If
    End If
    If A2(ii) = 9 Then
    If mark20 > 0 Then mark20 = ii
    Exit For
    End If
Next ii

'If mark10 > 0 And mark20 > 0 Then
there11:
If mcharpos > 0 Then
If A2(mcharpos) = 3 And A2(mcharpos - 1) <> 3 Then
M1 = mcharpos
M2 = mcharpos
Else
M1 = mcharpos - 1
M2 = mcharpos - 1
End If
Else
Exit Sub

End If
Dim findme As Long


If A2(M1) <> 3 Then Exit Sub
' find language
For findme = mcharpos To mark1 Step -1
If A2(findme) = 2 Then Exit For
Next findme
findme = AscW(Mid$(what, findme + 1, 1))
If findme >= &H600& And findme <= &H6FF& Then
' arabic language
                    Do While M1 > 0
                        If (A2(M1 - 1) > 3 And A2(M1) < 11) And (A2(M1) > 3 And A2(M1) < 10) Then

                        ElseIf A2(M1 - 1) = 4 Then
                        If M1 > 1 Then
                        If A2(M1 - 2) >= 10 Then
                        M1 = M1 - 1
                        Exit Do
                        End If
                        End If
                        Else
                        If Not (A2(M1) = 3 Or A2(M1) = 7) Then Exit Do
                       
                    End If
                        M1 = M1 - 1
                    Loop
                    Do While M2 < mark2
                    If (A2(M2 + 1) > 3 And A2(M2) < 11) And (A2(M2) > 3 And A2(M2) < 10) Then

                        Else
                    If Not (A2(M2) = 3 Or A2(M2) = 7) Then Exit Do
                    
                    End If
                        M2 = M2 + 1
                    Loop
Else

Do While M1 > 0
    If (A2(M1 - 1) > 3 And A2(M1) < 10) Or A2(M1) = 4 Then
 If A2(M1 - 1) = 4 Then
 If M1 > 1 Then
 If A2(M1 - 2) = 11 Then
 M1 = M1 - 1: Exit Do
 End If
 End If
 ElseIf A2(M1) = 1 Then
Exit Do
End If
    Else
    If Not (A2(M1) = 3 Or A2(M1) = 7) Then Exit Do
   
End If
    M1 = M1 - 1
Loop
Do While M2 < mark2

If A2(M2) > 3 And A2(M2) < 10 Then
If A2(M2) = 4 Then
If A2(M2 + 1) <> 3 Then
Exit Do
End If
End If
    Else
If Not (A2(M2) = 3 Or A2(M2) = 7) Then Exit Do

End If
    M2 = M2 + 1
Loop
End If
M1 = M1 + 2

numhandle = True
If mcharpos = M2 Then
where = M1 - (M1 = 1)

Else
where = M2 - (mcharpos - M1)

End If


End Sub

Private Sub glistN_ExpandSelStart(val As Long)
If val < 1 Then val = 1
glistN.SelStartNoEvents = expandCharPos
val = expandCharPos
FixCharPos
expandCharPos = val
End Sub
Private Sub glistN_SetExpandSS(val As Long)
expandCharPos = val
End Sub
Private Sub glistN_DelExpandSS()
expandCharPos = 0

End Sub
Private Sub FixCharPos()
Dim what$, ThatPara As Long, pastchars As Long
pastchars = mDoc.TextParagraphLine2(glistN.ListIndex + 1, what$, ThatPara, False)
If Len(what$) + 1 < glistN.SelStart Then
    Charpos = pastchars + Len(what$) + 1
    glistN.SelStartNoEvents = Len(what$) + 1
Else
    Charpos = pastchars + glistN.SelStart
End If
End Sub

Public Property Get NoCenterLineEdit() As Boolean
NoCenterLineEdit = mNoCenterLineEdit
End Property

Public Property Let NoCenterLineEdit(ByVal RHS As Boolean)
mNoCenterLineEdit = RHS
glistN.OverrideShow = Not RHS
End Property
Public Sub SetBarStyle(bColor As Long, Optional bHatch As Long, Optional BHatchColor As Long)
    glistN.BarColor = mycolor(bColor)
    glistN.BarHatch = bHatch
    glistN.BarHatchColor = mycolor(BHatchColor)
End Sub
Private Sub glistN_PreviewKeyboardUnicode(a As String)
Callback.SpreadKey a
End Sub
Private Sub glistN_Fkey(a As Integer)
If a > 1000 Then
Callback.SendFKEY a - 1000
Else
If UsedAsTextBox Or NumberOnly Or NumberIntOnly Then Callback.SendFKEY a
End If
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
Private Sub gListN_AccKey(M As Long)
Callback.AccProces M
End Sub
Public Property Let IControlIndex_Index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Property Get IControlIndex_Index() As Long
IControlIndex_Index = mIndex
End Property

Public Sub IControlIndex_UnloadMe()
If glistN Is Nothing Then Exit Sub
Set glistN = Nothing
Callback.RemoveGuiControl ctrlName
mIndex = -1
Set Callback = Nothing
Set mDoc = Nothing
Set mDoc1 = Nothing
End Sub

